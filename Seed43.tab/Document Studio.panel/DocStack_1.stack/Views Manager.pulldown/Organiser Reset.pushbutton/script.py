# -*- coding: utf-8 -*-
# view_organiser.py

import os
import json
import time

import clr
clr.AddReference("System")
from System.Collections.Generic import List
from pyrevit import revit, DB, forms, script
from Snippets import _userdata

doc    = revit.doc
output = script.get_output()

# ElementId.IntegerValue was removed in Revit 2024 -- use .Value instead.
# This helper works across all versions.
def eid_int(element_id):
    try:
        return element_id.Value          # Revit 2024+
    except AttributeError:
        return element_id.IntegerValue   # Revit 2023 and earlier

# ── Config (folder parameter mapping) ───────────────────────────────────────
# Stored in .user so an update can never overwrite it. Organiser Reset reads
# the same file and resolves it the same way - keep the two in step.
# A config saved by an older version, back when it sat beside this script, is
# moved across on first run.
CONFIG_PATH = _userdata.migrate(
    os.path.join(os.path.dirname(__file__), "view_organiser_config.json"),
    _userdata.user_path("ViewOrganiser", "view_organiser_config.json"))
NONE_LABEL  = "- None -"

def get_text_param_names(elements):
    """
    Return sorted list of writable text parameter names across a list of
    elements. Sampling multiple elements catches project parameters that
    only appear on certain view types.
    """
    names = set()
    for element in elements:
        try:
            for p in element.Parameters:
                try:
                    if (p.StorageType == DB.StorageType.String
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
        and v.ViewType not in [DB.ViewType.DrawingSheet]
    ][:20]

    sheet_params = [NONE_LABEL] + get_text_param_names(all_sheets[:10])
    view_params  = [NONE_LABEL] + get_text_param_names(sample_views)

    if len(sheet_params) == 1:
        print_warning("No writable text parameters found on sheets.")
    if len(view_params) == 1:
        print_warning("No writable text parameters found on views.")

    sheet_choice = forms.ask_for_one_item(
        sheet_params,
        default=NONE_LABEL,
        prompt="Which parameter on SHEETS holds the folder grouping?\n"
               "(This value will be read from the sheet and copied to the view.)\n\n"
               "Select 'None' to skip folder syncing.",
        title="View Organiser - Sheet Folder Parameter (source)"
    )
    if sheet_choice is None:
        script.exit()

    view_choice = forms.ask_for_one_item(
        view_params,
        default=NONE_LABEL,
        prompt="Which parameter on VIEWS should receive the folder value?\n\n"
               "Select 'None' to skip folder syncing.",
        title="View Organiser - View Folder Parameter (destination)"
    )
    if view_choice is None:
        script.exit()

    CASE_OPTIONS = [
        "ALL CAPS  (e.g. GROUND FLOOR PLAN)",
        "Title Case  (e.g. Ground Floor Plan)",
        "Ignore  (leave as-is)",
    ]
    CASE_MAP = {
        CASE_OPTIONS[0]: "upper",
        CASE_OPTIONS[1]: "title",
        CASE_OPTIONS[2]: "ignore",
    }
    title_case_choice = forms.ask_for_one_item(
        CASE_OPTIONS,
        default=CASE_OPTIONS[0],
        prompt="How should 'Title on Sheet' values be cased?\n\n"
               "ALL CAPS   - forces every title to uppercase\n"
               "Title Case - capitalises the first letter of each word\n"
               "Ignore     - no case changes applied",
        title="View Organiser - Title on Sheet Casing"
    )
    if title_case_choice is None:
        script.exit()

    sheet_name_case_choice = forms.ask_for_one_item(
        CASE_OPTIONS,
        default=CASE_OPTIONS[0],
        prompt="How should Sheet Name values be cased?\n\n"
               "ALL CAPS   - forces every sheet name to uppercase\n"
               "Title Case - capitalises the first letter of each word\n"
               "Ignore     - no case changes applied",
        title="View Organiser - Sheet Name Casing"
    )
    if sheet_name_case_choice is None:
        script.exit()

    return (
        None if sheet_choice == NONE_LABEL else sheet_choice,
        None if view_choice  == NONE_LABEL else view_choice,
        CASE_MAP[title_case_choice],
        CASE_MAP[sheet_name_case_choice],
    )

def load_config():
    """Return config dict from JSON, or None if file does not exist."""
    if os.path.isfile(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r") as f:
                return json.load(f)
        except Exception:
            return None
    return None

def save_config(sheet_param, view_param, title_on_sheet_case="upper", sheet_name_case="upper"):
    """Persist the chosen parameter names and casing modes to JSON."""
    data = {
        "sheet_folder_param":   sheet_param,
        "view_folder_param":    view_param,
        "title_on_sheet_case":  title_on_sheet_case,
        "sheet_name_case":      sheet_name_case,
    }
    try:
        with open(CONFIG_PATH, "w") as f:
            json.dump(data, f, indent=2)
    except Exception as e:
        print_warning("Could not save config: {}".format(str(e)))

# ── Output styles ─────────────────────────────────────────────────────────────
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
    output.print_html(
        "<div class='line'>------------------------------------</div>")

def print_success(text):
    output.print_html("<div class='sheet'>{}</div>".format(text))

def print_warning(text):
    output.print_html("<div class='warn'>WARNING: {}</div>".format(text))

def print_error(text):
    output.print_html("<div class='warn'>ERROR: {}</div>".format(text))

def print_info(text):
    output.print_html("<div class='sheet'>{}</div>".format(text))

def print_dim(text):
    output.print_html("<div class='rev'>-> {}</div>".format(text))

# ── Helpers ───────────────────────────────────────────────────────────────────

def apply_case(text, mode):
    if mode == "upper":
        return text.upper()
    if mode == "title":
        return text.title()
    return text

def clean_title(title):
    if title:
        return title.replace("\r\n", " ").replace("\n", " ").strip()
    return ""

def sanitize_name(name):
    """Strip or replace characters prohibited in Revit view/sheet names.

    Revit's actual restriction (per the in-app "Name cannot contain any of
    the following characters" error) is: \\ : { } [ ] | ; < > ? ` ~
    Forward slash ("/") is NOT on that list and is a legal, commonly used
    character in sheet and view names - it must not be touched here.

    Rules:
      \\ : | -> replaced with -
      { } [ ] < > ? " ; ` ~ -> removed entirely
      Multiple consecutive - or spaces collapsed
      Leading/trailing whitespace stripped
    """
    import re
    if not name:
        return name
    for ch in "\\:|":
        name = name.replace(ch, "-")
    for ch in "{}[]<>?\";`~\r\n":
        name = name.replace(ch, "")
    name = re.sub(r"\s*-+\s*", " - ", name)
    name = re.sub(r" {2,}", " ", name)
    name = name.strip(" -")
    return name

def get_parameter_by_name(element, param_name):
    try:
        return element.LookupParameter(param_name)
    except Exception:
        return None

def should_process_view(view):
    """Return True for normal, non-template views that can be renamed."""
    try:
        return (
            view.ViewType not in [
                DB.ViewType.Legend,
                DB.ViewType.Schedule,
                DB.ViewType.DrawingSheet
            ]
            and not view.IsTemplate
        )
    except Exception:
        return False

def unique_name(base_name, used_names):
    """Return base_name, or base_name (N) if already taken."""
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
all_views = list(
    DB.FilteredElementCollector(doc)
    .OfClass(DB.View)
    .WhereElementIsNotElementType()
    .ToElements()
)

all_sheets = list(
    DB.FilteredElementCollector(doc)
    .OfClass(DB.ViewSheet)
    .ToElements()
)

# ── Folder parameter config ───────────────────────────────────────────────────
config = load_config()
if config is None:
    print_info("First run - picking folder parameters...")
    sheet_folder_param_name, view_folder_param_name, title_on_sheet_case, sheet_name_case = pick_folder_params(
        all_sheets, all_views)
    save_config(sheet_folder_param_name, view_folder_param_name, title_on_sheet_case, sheet_name_case)
    if sheet_folder_param_name or view_folder_param_name:
        print_success("Config saved -> {}".format(CONFIG_PATH))
        print_dim("  Sheet param:       {}".format(sheet_folder_param_name or "None"))
        print_dim("  View param:        {}".format(view_folder_param_name  or "None"))
        print_dim("  Title on Sheet:    {}".format(title_on_sheet_case))
        print_dim("  Sheet name case:   {}".format(sheet_name_case))
    else:
        print_warning("Both folder params set to None - folder syncing will be skipped")
else:
    sheet_folder_param_name = config.get("sheet_folder_param")
    view_folder_param_name  = config.get("view_folder_param")
    title_on_sheet_case     = config.get("title_on_sheet_case", "upper")
    sheet_name_case         = config.get("sheet_name_case", "upper")
    print_success("Config loaded:")
    print_dim("  Sheet param:       {}".format(sheet_folder_param_name or "None"))
    print_dim("  View param:        {}".format(view_folder_param_name  or "None"))
    print_dim("  Title on Sheet:    {}".format(title_on_sheet_case))
    print_dim("  Sheet name case:   {}".format(sheet_name_case))

SYNC_FOLDERS = bool(sheet_folder_param_name and view_folder_param_name)
print_separator()

# ══════════════════════════════════════════════════════════════════════════════
# PART 1 - Uncheck "Folder" in all view templates
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
    with revit.Transaction("Uncheck 'Folder' in all View Templates"):
        try:
            for vt in view_templates:
                non_controlled = list(vt.GetNonControlledTemplateParameterIds())
                if view_folder_param_id not in non_controlled:
                    non_controlled.append(view_folder_param_id)
                    vt.SetNonControlledTemplateParameterIds(
                        List[DB.ElementId](non_controlled))
                    templates_updated += 1
            print_success(
                "Unchecked 'Folder' in {} view templates".format(
                    templates_updated))
        except Exception as e:
            print_error("Part 1 failed: {}".format(str(e)))
else:
    print_warning(
        "'Folder' parameter not found in view templates - skipping Part 1")

print_separator()

# ══════════════════════════════════════════════════════════════════════════════
# PART 2 - View renaming and folder sync
# T2a clears namespace; T2b applies final names and parameter writes.
# ══════════════════════════════════════════════════════════════════════════════
print_header("PART 2: VIEW UPDATE OPERATIONS")

# Build view to sheet map
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

processable_views = [v for v in all_views if should_process_view(v)]
views_on_sheets   = [v for v in processable_views if v.Id in view_to_sheet_map]
print_success("Found {} views placed on sheets".format(len(views_on_sheets)))

# Identify parent views whose dependents are on sheets
print_info("Identifying parent views with placed dependents...")
parent_views_to_process = {}   # parent ElementId -> one dependent ElementId
for view in views_on_sheets:
    try:
        parent_id = view.GetPrimaryViewId()
        if (parent_id != DB.ElementId.InvalidElementId
                and parent_id != view.Id):
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
    original_names = {}   # eid_int(view.Id) -> original name string
    for view in views_on_sheets:
        original_names[eid_int(view.Id)] = clean_title(view.Name)

    # ── T2a: Temp renames ────────────────────────────────────────────────────
    print_info("T2a: Clearing namespace with temporary names...")
    temp_count = 0

    with revit.Transaction("View Organizer - temp renames"):
        for view in views_on_sheets:
            try:
                view.Name = "TEMP_RENAME_{}_{}".format(
                    eid_int(view.Id), temp_count)
                temp_count += 1
            except Exception as e:
                print_warning("Could not temp-rename '{}': {}".format(
                    view.Name, str(e)))

        for parent_id in parent_views_to_process:
            try:
                parent_view = doc.GetElement(parent_id)
                parent_view.Name = "TEMP_PARENT_{}_{}".format(
                    eid_int(parent_id), temp_count)
                temp_count += 1
            except Exception as e:
                print_warning(
                    "Could not temp-rename parent view {}: {}".format(
                        eid_int(parent_id), str(e)))

    print_success("Temporarily renamed {} views".format(temp_count))
    print_separator()

    # ── T2b: Final renames + parameter writes ─────────────────────────────────
    print_info("T2b: Applying final names and updating parameters...")
    try:
        with revit.Transaction(
                "View Organizer - final renames and folder sync"):

            # Phase 0: Preserve original name in "Title on Sheet" if blank
            print_info(
                "Phase 0: Preserving original view names in Title on Sheet...")
            title_preserved_count = 0
            for view in views_on_sheets:
                try:
                    title_param = get_parameter_by_name(view, "Title on Sheet")
                    if title_param and not title_param.IsReadOnly:
                        if not (title_param.AsString() or ""):
                            original = original_names.get(eid_int(view.Id), "")
                            if original:
                                title_param.Set(apply_case(original, title_on_sheet_case))
                                title_preserved_count += 1
                                print_dim("Preserved: '{}' -> Title on Sheet".format(
                                    original))
                except Exception as e:
                    print_warning(
                        "Could not preserve title for id {}: {}".format(
                            eid_int(view.Id), str(e)))

            print_success("Preserved {} original names in Title on Sheet".format(
                title_preserved_count))
            print_separator()

            # Phase 1: Build used_names from views not being touched
            live_views = list(
                DB.FilteredElementCollector(doc)
                .OfClass(DB.View)
                .WhereElementIsNotElementType()
                .ToElements()
            )
            ids_being_processed = set(
                list(view_to_sheet_map.keys())
                + list(parent_views_to_process.keys())
            )
            used_names = {}
            for v in live_views:
                if (v.Id not in ids_being_processed
                        and not v.Name.startswith("TEMP_")):
                    used_names[v.Name] = True

            # Phase 2: Rename views on sheets
            processed_count        = 0
            error_count            = 0
            renamed_count          = 0
            folder_updated_count   = 0
            duplicate_suffix_count = 0
            parent_views_renamed   = 0
            parent_final_names     = {}   # parent_id -> base name

            for view in views_on_sheets:
                try:
                    sheet        = view_to_sheet_map[view.Id]
                    sheet_number = sheet.SheetNumber or "XX"

                    sheet_folder_value = ""
                    if SYNC_FOLDERS:
                        sh_folder_param = get_parameter_by_name(
                            sheet, sheet_folder_param_name)
                        if sh_folder_param:
                            sheet_folder_value = (
                                sh_folder_param.AsString() or "")

                    title_param = get_parameter_by_name(view, "Title on Sheet")
                    name_title  = ""
                    if title_param:
                        name_title = apply_case(
                            clean_title(title_param.AsString() or ""), title_on_sheet_case)
                        if name_title and not title_param.IsReadOnly:
                            try:
                                if title_param.AsString() != name_title:
                                    title_param.Set(name_title)
                            except Exception as e:
                                print_warning("Could not set Title on Sheet for view {}: {}".format(
                                    eid_int(view.Id), str(e)))
                    if not name_title:
                        name_title = "UNTITLED"

                    detail_param  = get_parameter_by_name(view, "Detail Number")
                    detail_number = (
                        detail_param.AsString() or "XX"
                        if detail_param else "XX"
                    )

                    raw_name  = "{} - {} - {}".format(
                        sheet_number, detail_number, name_title)
                    base_name = sanitize_name(raw_name)
                    if base_name != raw_name:
                        print_dim("Sanitized: '{}' -> '{}'".format(
                            raw_name, base_name))
                    new_name  = unique_name(base_name, used_names)
                    if new_name != base_name:
                        duplicate_suffix_count += 1
                        print_dim("Duplicate resolved: '{}' -> '{}'".format(
                            base_name, new_name))

                    parent_id = view.GetPrimaryViewId()
                    if (parent_id != DB.ElementId.InvalidElementId
                            and parent_id != view.Id):
                        if parent_id not in parent_final_names:
                            parent_final_names[parent_id] = base_name

                    try:
                        view.Name = new_name
                        renamed_count += 1
                    except Exception as e:
                        print_error(
                            "Could not rename view to '{}': {}".format(
                                new_name, str(e)))

                    if SYNC_FOLDERS and sheet_folder_value:
                        view_folder_param = get_parameter_by_name(
                            view, view_folder_param_name)
                        if view_folder_param and not view_folder_param.IsReadOnly:
                            try:
                                if (view_folder_param.AsString() or "") \
                                        != sheet_folder_value:
                                    view_folder_param.Set(sheet_folder_value)
                                    folder_updated_count += 1
                            except Exception as e:
                                print_warning(
                                    "Could not set Folder for '{}': {}".format(
                                        new_name, str(e)))

                    processed_count += 1

                except Exception as e:
                    error_count += 1
                    print_error("Error processing view id {}: {}".format(
                        eid_int(view.Id), str(e)))

            # Phase 3: Rename parent views
            print_info("Processing parent views...")
            for parent_id in parent_views_to_process:
                try:
                    parent_view = doc.GetElement(parent_id)
                    raw_name    = (
                        parent_final_names.get(parent_id) or parent_view.Name
                    ) + " PARENT VIEW"
                    base_name   = sanitize_name(raw_name)
                    if base_name != raw_name:
                        print_dim("Sanitized: '{}' -> '{}'".format(
                            raw_name, base_name))
                    new_name    = unique_name(base_name, used_names)
                    try:
                        parent_view.Name = new_name
                        parent_views_renamed += 1
                        print_dim("Parent renamed: '{}'".format(new_name))
                    except Exception as e:
                        print_error(
                            "Could not rename parent view to '{}': {}".format(
                                new_name, str(e)))
                except Exception as e:
                    print_error(
                        "Error processing parent view {}: {}".format(
                            eid_int(parent_id), str(e)))

        elapsed = time.time() - start_time
        print_separator()
        print_header("OPERATION COMPLETE")
        print_success("SUMMARY:")
        print_dim("Templates updated:            {}".format(templates_updated))
        print_dim("Original names preserved:     {}".format(
            title_preserved_count))
        print_dim("Views on sheets processed:    {}".format(processed_count))
        print_dim("Views renamed:                {}".format(renamed_count))
        print_dim("Parent views renamed:         {}".format(
            parent_views_renamed))
        print_dim("Duplicate suffixes added:     {}".format(
            duplicate_suffix_count))
        print_dim("Folders synced:               {}".format(
            folder_updated_count))
        print_separator()
        print_success("Total elements updated:  {}".format(
            templates_updated + title_preserved_count + renamed_count
            + parent_views_renamed + folder_updated_count))
        print_success("Processing time:  {:.2f} seconds".format(elapsed))
        print_separator()
        if error_count > 0:
            print_warning(
                "Errors encountered on {} views - check log above".format(
                    error_count))
        else:
            print_success("No errors - all done")

    except Exception as e:
        print_error("T2b failed and rolled back: {}".format(str(e)))
        print_warning(
            "Views may still have TEMP_ names - use Ctrl+Z to undo T2a as well.")

# ══════════════════════════════════════════════════════════════════════════════
# PART 3 - Apply casing to sheet Name property across all ViewSheet elements
# ══════════════════════════════════════════════════════════════════════════════
print_header("PART 3: SHEET NAME CASING")

CASE_LABELS = {
    "upper":  "ALL CAPS",
    "title":  "Title Case",
    "ignore": "Ignore (skipping)",
}
print_dim("Mode: {}".format(CASE_LABELS.get(sheet_name_case, sheet_name_case)))

if sheet_name_case == "ignore":
    print_info("Casing mode is Ignore - Part 3 skipped.")
else:
    all_sheets_p3 = list(
        DB.FilteredElementCollector(doc)
        .OfClass(DB.ViewSheet)
        .ToElements()
    )
    print_dim("Found {} sheets to check".format(len(all_sheets_p3)))

    sheet_updated_count = 0
    sheet_skipped_count = 0
    sheet_error_count   = 0

    try:
        with revit.Transaction("View Organiser - casing sheet names"):
            for sheet in all_sheets_p3:
                try:
                    current_val = clean_title(sheet.Name or "")
                    if not current_val:
                        sheet_skipped_count += 1
                        continue

                    cased     = sanitize_name(apply_case(current_val, sheet_name_case))
                    pre_sanit = apply_case(current_val, sheet_name_case)
                    if pre_sanit != cased:
                        print_dim("Sanitized sheet: '{}' -> '{}'".format(
                            pre_sanit, cased))
                    if cased == current_val:
                        sheet_skipped_count += 1
                        continue

                    sheet.Name = cased
                    sheet_updated_count += 1
                    print_dim("Cased: '{}' -> '{}'".format(current_val, cased))

                except Exception as e:
                    sheet_error_count += 1
                    print_error("Error on sheet '{}': {}".format(
                        getattr(sheet, "SheetNumber", str(sheet.Id)), str(e)))

        print_success("Sheet names updated:  {}".format(sheet_updated_count))
        print_success("Already correct / blank / skipped:  {}".format(
            sheet_skipped_count))
        if sheet_error_count:
            print_warning("{} sheet(s) failed - check log above".format(
                sheet_error_count))

    except Exception as e:
        print_error("Part 3 transaction failed: {}".format(str(e)))

print_separator()
print_header("SCRIPT FINISHED")
