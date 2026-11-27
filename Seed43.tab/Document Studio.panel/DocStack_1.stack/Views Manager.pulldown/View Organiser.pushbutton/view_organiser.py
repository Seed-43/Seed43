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

# ── Config (sheet -> view / schedule / legend parameter links) ───────────────
# The setup window, the saved file and its migration all live in
# Snippets/_organiser_setup.py, shared with Organiser Reset so both always
# read and write the same file (in .user, so an update never overwrites it).
from Snippets import _organiser_setup as setup

CONFIG_PATH = setup.CONFIG_PATH

def value_text(param):
    """A parameter's value as text: a string as it is, an ElementId one (e.g.
    Sheet Collection) as the name of what it points at."""
    try:
        if param.StorageType == DB.StorageType.String:
            return param.AsString() or ""
        text = param.AsValueString() or ""
    except Exception:
        return ""
    return "" if text.strip() in ("", "<None>") else text

def sheet_value(sheet, name):
    """Text value of a sheet parameter. A sheet can carry more than one
    parameter of the same name (Sheet Collection has a read-only twin), so
    the first that holds a value wins."""
    try:
        for p in sheet.GetParameters(name):
            text = value_text(p)
            if text:
                return text
    except Exception:
        pass
    return ""

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
    name = re.sub(r"-{2,}", "-", name)      # collapse repeated dashes, no spacing added
    name = re.sub(r" {2,}", " ", name)      # collapse repeated spaces
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

all_schedules = [
    v for v in DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule).ToElements()
    if not v.IsTemplate
]

# ── Parameter links config ────────────────────────────────────────────────────
# Asked once in the setup window. It only comes back when the saved file is
# missing, or lacks a section (added by an update). Whatever is already saved
# is pre-filled in the window.
raw_config = setup.load_config()
config = setup.migrate_config(raw_config)
todo = setup.missing_parts(config)
broken = [] if todo else setup.unavailable_links(doc, config)
if todo or broken:
    if raw_config is None:
        notice = ("First run: map the sheet parameters to the views, schedules "
                  "and legends placed on them. This is asked once and saved.")
    elif todo:
        notice = ("There has been an update to View Organiser. Its saved setup "
                  "is missing: {}. Please check the links below (your earlier "
                  "choices are filled in) and save to continue.".format(
                      ", ".join(todo)))
    else:
        notice = ("Some saved parameters are no longer available, so the links "
                  "need remapping: {}. Choose new ones (or None) and save to "
                  "continue.".format("; ".join(broken)))
    print_info("Opening the setup window...")
    config = setup.show_setup(doc, config, notice)
    if config is None:
        print_warning("Setup cancelled - nothing was changed.")
        script.exit()
    print_success("Config saved -> {}".format(CONFIG_PATH))

SECTIONS      = setup.SECTIONS
SECTION_LABEL = {"views": "VIEWS", "schedules": "SCHEDULES", "legends": "LEGENDS"}
section_targets = {
    "views":     [v for v in all_views if should_process_view(v)],
    "schedules": all_schedules,
    "legends":   [v for v in all_views
                  if not v.IsTemplate and v.ViewType == DB.ViewType.Legend],
}
title_on_sheet_case = config.get("title_on_sheet_case", "upper")
sheet_name_case     = config.get("sheet_name_case", "upper")
MAPS = config["maps"]

print_success("Config:")
for section in SECTIONS:
    links = MAPS.get(section) or []
    if not links:
        print_dim("  {}: no links".format(SECTION_LABEL[section]))
    for l in links:
        print_dim("  {}: sheet '{}' -> '{}'".format(
            SECTION_LABEL[section], l["sheet"], l["target"]))
print_dim("  Title on Sheet:    {}".format(title_on_sheet_case))
print_dim("  Sheet name case:   {}".format(sheet_name_case))
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

                    for link in MAPS.get("views", []):
                        link_value = sheet_value(sheet, link["sheet"])
                        if not link_value:
                            continue
                        dst = get_parameter_by_name(view, link["target"])
                        if (dst is None or dst.IsReadOnly
                                or dst.StorageType != DB.StorageType.String):
                            continue
                        try:
                            if (dst.AsString() or "") != link_value:
                                dst.Set(link_value)
                                folder_updated_count += 1
                        except Exception as e:
                            print_warning(
                                "Could not set '{}' for '{}': {}".format(
                                    link["target"], new_name, str(e)))

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
        print_dim("View parameters synced:      {}".format(
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

# ══════════════════════════════════════════════════════════════════════════════
# PART 4 / 5 - Copy sheet parameters to the schedules and legends on them
# ══════════════════════════════════════════════════════════════════════════════
def sync_links(section, target_to_sheet):
    """Write each sheet's linked parameter values to the schedules / legends
    placed on it. target_to_sheet: {ElementId: ViewSheet}."""
    links = MAPS.get(section) or []
    label = SECTION_LABEL[section]
    if not links:
        print_info("No {} links set - skipped. (Organiser Reset can ask "
                   "again.)".format(label.lower()))
        return
    print_dim("Found {} {} placed on sheets".format(
        len(target_to_sheet), label.lower()))
    updated = same = blank = unwritable = 0
    try:
        with revit.Transaction("View Organiser - {} links".format(label.lower())):
            for target_id, sheet in target_to_sheet.items():
                target = doc.GetElement(target_id)
                for link in links:
                    value = sheet_value(sheet, link["sheet"])
                    if not value:
                        blank += 1
                        continue
                    dst = get_parameter_by_name(target, link["target"])
                    if (dst is None or dst.IsReadOnly
                            or dst.StorageType != DB.StorageType.String):
                        unwritable += 1
                        print_warning("Cannot write '{}' on '{}'".format(
                            link["target"], target.Name))
                        continue
                    if (dst.AsString() or "") == value:
                        same += 1
                        continue
                    dst.Set(value)
                    updated += 1
        print_success("{} parameters set:  {}".format(label.title(), updated))
        print_dim("Already correct:  {}".format(same))
        print_dim("Sheet value blank:  {}".format(blank))
        if unwritable:
            print_warning("{} could not be written".format(unwritable))
    except Exception as e:
        print_error("{} links failed: {}".format(label.title(), str(e)))

print_header("PART 4: SCHEDULE LINKS")
sched_to_sheet = {}
for inst in DB.FilteredElementCollector(doc).OfClass(DB.ScheduleSheetInstance):
    try:
        if inst.IsTitleblockRevisionSchedule:
            continue
        sh = doc.GetElement(inst.OwnerViewId)
        if isinstance(sh, DB.ViewSheet) and inst.ScheduleId not in sched_to_sheet:
            sched_to_sheet[inst.ScheduleId] = sh
    except Exception:
        pass
sync_links("schedules", sched_to_sheet)
print_separator()

print_header("PART 5: LEGEND LINKS")
legend_to_sheet = {}
for lg in section_targets["legends"]:
    if lg.Id in view_to_sheet_map:
        legend_to_sheet[lg.Id] = view_to_sheet_map[lg.Id]
sync_links("legends", legend_to_sheet)
print_separator()
print_header("SCRIPT FINISHED")
