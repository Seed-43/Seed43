# -*- coding: utf-8 -*-
# view_organiser.py
# --- Imports ---
import clr
import os
import json
import time
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB import (
    FilteredElementCollector, Transaction, TransactionGroup,
    ElementId, View, ViewSheet, ViewType, StorageType
)
from System.Collections.Generic import List
from pyrevit import revit, forms, script

doc    = revit.doc
output = script.get_output()

# ElementId.IntegerValue was removed in Revit 2024 — use .Value instead.
# This helper works across all versions.
def eid_int(element_id):
    try:
        return element_id.Value          # Revit 2024+
    except AttributeError:
        return element_id.IntegerValue   # Revit 2023 and earlier

# ── Config (folder parameter mapping) ───────────────────────────────────────
# Saved next to this script so it travels with the extension.
CONFIG_PATH = os.path.join(os.path.dirname(__file__), "view_organiser_config.json")
NONE_LABEL  = "— None —"

def get_text_param_names(elements):
    """
    Return sorted list of writable text parameter names across a list of elements.
    Sampling multiple elements catches project parameters that only appear on
    certain view types.
    """
    names = set()
    for element in elements:
        try:
            for p in element.Parameters:
                try:
                    if (p.StorageType == StorageType.String
                            and not p.IsReadOnly
                            and p.Definition
                            and p.Definition.Name):
                        names.add(p.Definition.Name)
                except Exception:
                    pass
        except Exception:
            pass
    return sorted(names)

def pick_folder_params(all_sheets, all_views):
    """
    Two dropdowns:
      1. Which text parameter on SHEETS is the folder source.
      2. Which text parameter on VIEWS is the folder destination.
    Either can be set to None to skip syncing.
    Returns (sheet_param_name, view_param_name).
    """
    sample_views = [
        v for v in all_views
        if not v.IsTemplate
        and v.ViewType not in [ViewType.DrawingSheet]
    ][:20]

    sheet_params = [NONE_LABEL] + get_text_param_names(all_sheets[:10])
    view_params  = [NONE_LABEL] + get_text_param_names(sample_views)

    if len(sheet_params) == 1:
        print_warning("No writable text parameters found on sheets.")
    if len(view_params) == 1:
        print_warning("No writable text parameters found on views.")

    # Sheet picker (source)
    sheet_choice = forms.ask_for_one_item(
        sheet_params,
        default=NONE_LABEL,
        prompt="Which parameter on SHEETS holds the folder grouping?\n"
               "(This value will be read from the sheet and copied to the view.)\n\n"
               "Select 'None' to skip folder syncing.",
        title="View Organiser — Sheet Folder Parameter (source)"
    )
    if sheet_choice is None:
        script.exit()

    # View picker (destination)
    view_choice = forms.ask_for_one_item(
        view_params,
        default=NONE_LABEL,
        prompt="Which parameter on VIEWS should receive the folder value?\n\n"
               "Select 'None' to skip folder syncing.",
        title="View Organiser — View Folder Parameter (destination)"
    )
    if view_choice is None:
        script.exit()

    return (
        None if sheet_choice == NONE_LABEL else sheet_choice,
        None if view_choice  == NONE_LABEL else view_choice
    )

def load_config():
    """Return config dict from JSON, or None if file doesn't exist."""
    if os.path.isfile(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r") as f:
                return json.load(f)
        except Exception:
            return None
    return None

def save_config(sheet_param, view_param):
    """Persist the chosen parameter names to JSON."""
    data = {
        "sheet_folder_param": sheet_param,
        "view_folder_param":  view_param
    }
    try:
        with open(CONFIG_PATH, "w") as f:
            json.dump(data, f, indent=2)
    except Exception as e:
        print_warning("Could not save config: {}".format(str(e)))

# --- Terminal output styles ---
output.add_style("""
body {
    background-color: #232933;
    color: #F4FAFF;
    font-family: Consolas, Courier New, monospace;
    padding: 20px;
}
.header {
    color: #2B933F;
    font-weight: bold;
    font-size: 1.2em;
}
.sheet {
    color: #F4FAFF;
    padding-left: 15px;
}
.rev {
    color: #8B9199;
    padding-left: 30px;
}
.warn {
    color: #E0A040;
    padding-left: 30px;
}
.line {
    color: #3B4553;
}
""")

def print_header(text):
    output.print_html("<div class='header'>{}</div>".format(text))

def print_separator():
    output.print_html("<div class='line'>────────────────────────────────────</div>")

def print_success(text):
    output.print_html("<div class='sheet'>{}</div>".format(text))

def print_warning(text):
    output.print_html("<div class='warn'>WARNING: {}</div>".format(text))

def print_error(text):
    output.print_html("<div class='warn'>ERROR: {}</div>".format(text))

def print_info(text):
    output.print_html("<div class='sheet'>{}</div>".format(text))

def print_dim(text):
    output.print_html("<div class='rev'>→ {}</div>".format(text))

# --- Helpers ---

def clean_title(title):
    if title:
        return title.replace("\r\n", " ").replace("\n", " ").strip()
    return ""

def get_parameter_by_name(element, param_name):
    try:
        return element.LookupParameter(param_name)
    except Exception:
        return None

def should_process_view(view):
    """Return True for normal, non-template views that can be renamed."""
    try:
        return (
            view.ViewType not in [ViewType.Legend, ViewType.Schedule, ViewType.DrawingSheet]
            and not view.IsTemplate
        )
    except Exception:
        return False

def unique_name(base_name, used_names):
    """Return base_name, or base_name (N) if already taken. Registers the result."""
    name = base_name
    if name in used_names:
        counter = 2
        while True:
            name = "{} ({})".format(base_name, counter)
            if name not in used_names:
                break
            counter += 1
    used_names[name] = True
    return name

# ══════════════════════════════════════════════════════════════════════════════
start_time = time.time()
print_header("VIEW ORGANISER ENGINE INITIALISING")
print_separator()

# ── Collect once, correctly ───────────────────────────────────────────────────
# WhereElementIsNotElementType() ensures we get instances only, not type objects.
all_views = list(
    FilteredElementCollector(doc)
    .OfClass(View)
    .WhereElementIsNotElementType()
    .ToElements()
)

all_sheets = list(
    FilteredElementCollector(doc)
    .OfClass(ViewSheet)
    .ToElements()
)

# ── Folder parameter config ──────────────────────────────────────────────────
config = load_config()
if config is None:
    print_info("First run — picking folder parameters...")
    sheet_folder_param_name, view_folder_param_name = pick_folder_params(all_sheets, all_views)
    save_config(sheet_folder_param_name, view_folder_param_name)
    if sheet_folder_param_name or view_folder_param_name:
        print_success("Config saved → {}".format(CONFIG_PATH))
        print_dim("  Sheet param: {}".format(sheet_folder_param_name or "None"))
        print_dim("  View param:  {}".format(view_folder_param_name  or "None"))
    else:
        print_warning("Both set to None — folder syncing will be skipped")
else:
    sheet_folder_param_name = config.get("sheet_folder_param")
    view_folder_param_name  = config.get("view_folder_param")
    print_success("Config loaded:")
    print_dim("  Sheet param: {}".format(sheet_folder_param_name or "None"))
    print_dim("  View param:  {}".format(view_folder_param_name  or "None"))

SYNC_FOLDERS = bool(sheet_folder_param_name and view_folder_param_name)
print_separator()

# ══════════════════════════════════════════════════════════════════════════════
# PART 1 — Uncheck "Folder" in all view templates
# Committed as its own transaction so a Part 2 failure can't roll this back.
# ══════════════════════════════════════════════════════════════════════════════
print_header("PART 1: TEMPLATE PARAMETER MANAGEMENT")

view_templates = [v for v in all_views if v.IsTemplate]
print_dim("Found {} view templates".format(len(view_templates)))

view_folder_param_id = None
if view_templates:
    print_info("Searching for 'Folder' parameter...")
    for param_id in view_templates[0].GetTemplateParameterIds():
        param = doc.GetElement(param_id)
        if param and hasattr(param, "Name") and param.Name == "Folder":
            view_folder_param_id = param_id
            print_success("'Folder' parameter located")
            break

templates_updated = 0
if view_folder_param_id is not None:
    t1 = Transaction(doc, "Uncheck 'Folder' in all View Templates")
    t1.Start()
    try:
        for vt in view_templates:
            non_controlled = list(vt.GetNonControlledTemplateParameterIds())
            if view_folder_param_id not in non_controlled:
                non_controlled.append(view_folder_param_id)
                vt.SetNonControlledTemplateParameterIds(List[ElementId](non_controlled))
                templates_updated += 1
        t1.Commit()
        print_success("Unchecked 'Folder' in {} view templates".format(templates_updated))
    except Exception as e:
        t1.RollBack()
        print_error("Part 1 failed and rolled back: {}".format(str(e)))
else:
    print_warning("'Folder' parameter not found in view templates — skipping Part 1")

print_separator()

# ══════════════════════════════════════════════════════════════════════════════
# PART 2 — View renaming and folder sync
# Split into two committed transactions:
#   T2a — temp renames  (clears namespace)
#   T2b — final renames + parameter writes
# If T2b fails, T2a is already committed but temp names are easy to undo;
# a failed single transaction would have left temp names with no way out.
# ══════════════════════════════════════════════════════════════════════════════
print_header("PART 2: VIEW UPDATE OPERATIONS")

# Build view → sheet map
print_info("Building view-to-sheet mapping...")
view_to_sheet_map = {}
for sheet in all_sheets:
    try:
        for viewport_id in sheet.GetAllViewports():
            viewport = doc.GetElement(viewport_id)
            if viewport and hasattr(viewport, "ViewId"):
                view_to_sheet_map[viewport.ViewId] = sheet
    except Exception as e:
        print_warning("Error reading sheet {}: {}".format(
            getattr(sheet, "SheetNumber", "Unknown"), str(e)))

print_dim("Mapped {} views to sheets".format(len(view_to_sheet_map)))

# Filter to processable views
processable_views = [v for v in all_views if should_process_view(v)]
views_on_sheets   = [v for v in processable_views if v.Id in view_to_sheet_map]
print_success("Found {} views placed on sheets".format(len(views_on_sheets)))

# Identify parent views whose dependents are on sheets
print_info("Identifying parent views with placed dependents...")
parent_views_to_process = {}   # parent ElementId → one dependent ElementId
for view in views_on_sheets:
    try:
        parent_id = view.GetPrimaryViewId()
        if parent_id != ElementId.InvalidElementId and parent_id != view.Id:
            parent_view = doc.GetElement(parent_id)
            if parent_view and should_process_view(parent_view):
                if parent_id not in parent_views_to_process:
                    parent_views_to_process[parent_id] = view.Id
    except Exception as e:
        print_warning("Could not determine parent for view '{}': {}".format(
            getattr(view, "Name", "Unknown"), str(e)))

print_dim("Found {} parent views".format(len(parent_views_to_process)))
print_separator()

if not views_on_sheets and not parent_views_to_process:
    print_warning("No views found to process")
    print_header("OPERATION COMPLETE")
else:
    # Snapshot original names BEFORE T2a wipes them — needed for Title on Sheet
    # preservation in T2b where view.Name is already TEMP_RENAME_...
    original_names = {}   # eid_int(view.Id) → original name string
    for view in views_on_sheets:
        original_names[eid_int(view.Id)] = clean_title(view.Name)

    # ── T2a: Temp renames ────────────────────────────────────────────────────
    # Rename everything we're going to touch to TEMP_ names so Phase 2 starts
    # with a clean namespace and can't collide with existing names.
    print_info("T2a: Clearing namespace with temporary names...")
    t2a = Transaction(doc, "View Organizer — temp renames")
    t2a.Start()
    temp_count = 0
    try:
        for view in views_on_sheets:
            try:
                view.Name = "TEMP_RENAME_{}_{}".format(eid_int(view.Id), temp_count)
                temp_count += 1
            except Exception as e:
                print_warning("Could not temp-rename '{}': {}".format(view.Name, str(e)))

        for parent_id in parent_views_to_process:
            try:
                parent_view = doc.GetElement(parent_id)
                parent_view.Name = "TEMP_PARENT_{}_{}".format(eid_int(parent_id), temp_count)
                temp_count += 1
            except Exception as e:
                print_warning("Could not temp-rename parent view {}: {}".format(
                    eid_int(parent_id), str(e)))

        t2a.Commit()
        print_success("Temporarily renamed {} views".format(temp_count))
    except Exception as e:
        t2a.RollBack()
        print_error("T2a failed and rolled back — aborting: {}".format(str(e)))
        raise   # stop here; nothing has been permanently changed

    print_separator()

    # ── T2b: Final renames + parameter writes ────────────────────────────────
    print_info("T2b: Applying final names and updating parameters...")
    t2b = Transaction(doc, "View Organizer — final renames and folder sync")
    t2b.Start()
    try:
        # ── Phase 0: Preserve original name in "Title on Sheet" if blank ─────
        # view.Name is now TEMP_RENAME_... so we use the original_names snapshot
        # captured before T2a ran.
        print_info("Phase 0: Preserving original view names in Title on Sheet...")
        title_preserved_count = 0
        for view in views_on_sheets:
            try:
                title_param = get_parameter_by_name(view, "Title on Sheet")
                if title_param and not title_param.IsReadOnly:
                    if not (title_param.AsString() or ""):
                        original = original_names.get(eid_int(view.Id), "")
                        if original:
                            title_param.Set(original.upper())
                            title_preserved_count += 1
                            print_dim("Preserved: '{}' → Title on Sheet".format(original))
            except Exception as e:
                print_warning("Could not preserve title for id {}: {}".format(
                    eid_int(view.Id), str(e)))

        print_success("Preserved {} original names in Title on Sheet".format(title_preserved_count))
        print_separator()

        # ── Phase 1: Build used_names from views NOT being touched ────────────
        # Re-collect live names now that T2a has committed, so the snapshot
        # accurately reflects what's actually in the model right now.
        live_views = list(
            FilteredElementCollector(doc)
            .OfClass(View)
            .WhereElementIsNotElementType()
            .ToElements()
        )
        ids_being_processed = set(
            list(view_to_sheet_map.keys()) + list(parent_views_to_process.keys())
        )
        used_names = {}
        for v in live_views:
            if v.Id not in ids_being_processed and not v.Name.startswith("TEMP_"):
                used_names[v.Name] = True

        # ── Phase 2: Rename views on sheets ──────────────────────────────────
        processed_count          = 0
        error_count              = 0
        renamed_count            = 0
        folder_updated_count     = 0
        duplicate_suffix_count   = 0
        parent_views_renamed     = 0
        parent_final_names = {}   # parent_id → base name for its PARENT VIEW suffix

        for view in views_on_sheets:
            try:
                sheet        = view_to_sheet_map[view.Id]
                sheet_number = sheet.SheetNumber or "XX"

                # Sheet folder value — read from the SHEET
                sheet_folder_value = ""
                if SYNC_FOLDERS:
                    sh_folder_param = get_parameter_by_name(sheet, sheet_folder_param_name)
                    if sh_folder_param:
                        sheet_folder_value = sh_folder_param.AsString() or ""

                # Title on Sheet (now populated from Phase 0)
                title_param = get_parameter_by_name(view, "Title on Sheet")
                name_title  = ""
                if title_param:
                    name_title = clean_title(title_param.AsString() or "").upper()
                if not name_title:
                    name_title = "UNTITLED"

                # Detail number
                detail_param  = get_parameter_by_name(view, "Detail Number")
                detail_number = detail_param.AsString() or "XX" if detail_param else "XX"

                # Compose and de-duplicate
                base_name = "{} - {} - {}".format(sheet_number, detail_number, name_title)
                new_name  = unique_name(base_name, used_names)
                if new_name != base_name:
                    duplicate_suffix_count += 1
                    print_dim("Duplicate resolved: '{}' → '{}'".format(base_name, new_name))

                # Track for parent naming
                parent_id = view.GetPrimaryViewId()
                if parent_id != ElementId.InvalidElementId and parent_id != view.Id:
                    if parent_id not in parent_final_names:
                        parent_final_names[parent_id] = base_name

                try:
                    view.Name = new_name
                    renamed_count += 1
                except Exception as e:
                    print_error("Could not rename view to '{}': {}".format(new_name, str(e)))

                # Sync folder
                if SYNC_FOLDERS and sheet_folder_value:
                    view_folder_param = get_parameter_by_name(view, view_folder_param_name)
                    if view_folder_param and not view_folder_param.IsReadOnly:
                        try:
                            if (view_folder_param.AsString() or "") != sheet_folder_value:
                                view_folder_param.Set(sheet_folder_value)
                                folder_updated_count += 1
                        except Exception as e:
                            print_warning("Could not set Folder for '{}': {}".format(
                                new_name, str(e)))

                processed_count += 1

            except Exception as e:
                error_count += 1
                print_error("Error processing view id {}: {}".format(
                    eid_int(view.Id), str(e)))

        # ── Phase 3: Rename parent views ──────────────────────────────────────
        print_info("Processing parent views...")
        for parent_id in parent_views_to_process:
            try:
                parent_view = doc.GetElement(parent_id)
                base_name   = (parent_final_names.get(parent_id) or parent_view.Name) + " PARENT VIEW"
                new_name    = unique_name(base_name, used_names)
                try:
                    parent_view.Name = new_name
                    parent_views_renamed += 1
                    print_dim("Parent renamed: '{}'".format(new_name))
                except Exception as e:
                    print_error("Could not rename parent view to '{}': {}".format(new_name, str(e)))
            except Exception as e:
                print_error("Error processing parent view {}: {}".format(
                    eid_int(parent_id), str(e)))

        t2b.Commit()

        elapsed = time.time() - start_time
        print_separator()
        print_header("OPERATION COMPLETE")
        print_success("SUMMARY:")
        print_dim("Templates updated:            {}".format(templates_updated))
        print_dim("Original names preserved:     {}".format(title_preserved_count))
        print_dim("Views on sheets processed:    {}".format(processed_count))
        print_dim("Views renamed:                {}".format(renamed_count))
        print_dim("Parent views renamed:         {}".format(parent_views_renamed))
        print_dim("Duplicate suffixes added:     {}".format(duplicate_suffix_count))
        print_dim("Folders synced:               {}".format(folder_updated_count))
        print_separator()
        print_success("Total elements updated:  {}".format(
            templates_updated + title_preserved_count + renamed_count
            + parent_views_renamed + folder_updated_count))
        print_success("Processing time:  {:.2f} seconds".format(elapsed))
        print_separator()
        if error_count > 0:
            print_warning("Errors encountered on {} views — check log above".format(error_count))
        else:
            print_success("No errors — all done")

    except Exception as e:
        t2b.RollBack()
        print_error("T2b failed and rolled back: {}".format(str(e)))
        print_warning("Views may still have TEMP_ names — use Ctrl+Z to undo T2a as well.")

print_separator()
print_header("SCRIPT FINISHED")
