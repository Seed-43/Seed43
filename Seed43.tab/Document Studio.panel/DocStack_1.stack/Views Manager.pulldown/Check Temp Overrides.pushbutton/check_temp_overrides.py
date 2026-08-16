# -*- coding: utf-8 -*-
# check_temp_overrides.py
# Finds all views placed on sheets that have temporary graphic overrides active.
# Outputs a clickable report with links to navigate to the parent sheet.

import time
from pyrevit import revit, DB, script

doc    = revit.doc
output = script.get_output()

def eid_int(element_id):
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue

# ── Output styles ──────────────────────────────────────────────────────────────
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
    output.print_html("<div class='line'>------------------------------------</div>")

def print_success(text):
    output.print_html("<div class='sheet'>{}</div>".format(text))

def print_warning(text):
    output.print_html("<div class='warn'>&#9888; {}</div>".format(text))

def print_dim(text):
    output.print_html("<div class='rev'>-> {}</div>".format(text))

def print_info(text):
    output.print_html("<div class='sheet'>{}</div>".format(text))

# ── Helpers ────────────────────────────────────────────────────────────────────

def has_temp_overrides(view):
    """
    Return True if the view has any temporary graphic overrides active.
    Revit stores these on View.TemporaryViewModes -- if any mode other than
    None/Normal is active the view is in a temporary state.
    Also checks AreGraphicsOverridesTemporary() where available.
    """
    try:
        tvm = view.TemporaryViewModes
        # TemporaryViewPropertyType values that indicate an active temp state
        TEMP_PROPS = [
            DB.TemporaryViewPropertyType.TemporaryHideIsolate,
            DB.TemporaryViewPropertyType.RevealHiddenElements,
            DB.TemporaryViewPropertyType.WorksharingDisplay,
        ]
        # getattr, because 'None' is a Python keyword and cannot be written as
        # an attribute - same workaround as WindowStyle.None in _dialogs.py.
        _NONE = getattr(DB.TemporaryViewPropertyType, 'None')
        for prop in TEMP_PROPS:
            try:
                if tvm.GetTemporaryViewProperty(prop) != _NONE:
                    return True
            except Exception:
                pass
    except Exception:
        pass

    # Fallback: check HideIsolateActive flag directly
    try:
        if view.IsInTemporaryViewMode(DB.TemporaryViewMode.TemporaryHideIsolate):
            return True
    except Exception:
        pass

    try:
        if view.IsInTemporaryViewMode(DB.TemporaryViewMode.RevealHiddenElements):
            return True
    except Exception:
        pass

    return False

def get_temp_override_types(view):
    """Return a list of human-readable override type strings active on the view."""
    active = []
    try:
        if view.IsInTemporaryViewMode(DB.TemporaryViewMode.TemporaryHideIsolate):
            active.append("Hide/Isolate")
    except Exception:
        pass
    try:
        if view.IsInTemporaryViewMode(DB.TemporaryViewMode.RevealHiddenElements):
            active.append("Reveal Hidden Elements")
    except Exception:
        pass
    try:
        if view.IsInTemporaryViewMode(DB.TemporaryViewMode.WorksharingDisplay):
            active.append("Worksharing Display")
    except Exception:
        pass
    return active if active else ["Unknown override"]

# ── Main ───────────────────────────────────────────────────────────────────────
start_time = time.time()
print_header("TEMPORARY OVERRIDE CHECKER")
print_separator()

# Collect all sheets and build view -> sheet map
all_sheets = list(
    DB.FilteredElementCollector(doc)
    .OfClass(DB.ViewSheet)
    .ToElements()
)

view_to_sheet = {}   # view Id -> (sheet, viewport Id)
for sheet in all_sheets:
    try:
        for vp_id in sheet.GetAllViewports():
            vp = doc.GetElement(vp_id)
            if vp and hasattr(vp, "ViewId"):
                view_to_sheet[vp.ViewId] = (sheet, vp_id)
    except Exception:
        pass

print_dim("Checked {} sheets, {} views on sheets".format(
    len(all_sheets), len(view_to_sheet)))
print_separator()

# Check each view on a sheet for temp overrides
hits = []   # list of (sheet, view, override_types)

all_views = list(
    DB.FilteredElementCollector(doc)
    .OfClass(DB.View)
    .WhereElementIsNotElementType()
    .ToElements()
)

checked = 0
for view in all_views:
    try:
        if view.IsTemplate:
            continue
        if view.ViewType == DB.ViewType.DrawingSheet:
            continue
        if view.Id not in view_to_sheet:
            continue
        checked += 1
        if has_temp_overrides(view):
            sheet, vp_id = view_to_sheet[view.Id]
            override_types = get_temp_override_types(view)
            hits.append((sheet, vp_id, view, override_types))
    except Exception:
        pass

print_dim("Scanned {} views on sheets".format(checked))
print_separator()

# ── Report ─────────────────────────────────────────────────────────────────────
if not hits:
    print_header("ALL CLEAR")
    print_success("No views with temporary overrides found.")
else:
    print_header("FOUND {} VIEW(S) WITH TEMPORARY OVERRIDES".format(len(hits)))
    print_separator()

    # Group by sheet for cleaner output
    sheets_seen = {}
    for sheet, vp_id, view, override_types in hits:
        key = eid_int(sheet.Id)
        if key not in sheets_seen:
            sheets_seen[key] = {"sheet": sheet, "views": []}
        sheets_seen[key]["views"].append((vp_id, view, override_types))

    for key in sorted(sheets_seen.keys()):
        entry  = sheets_seen[key]
        sheet  = entry["sheet"]
        s_num  = sheet.SheetNumber or "??"
        s_name = sheet.Name or "Unnamed Sheet"

        output.print_html(
            "<div class='sheet'>{} - {}</div>".format(s_num, s_name))

        for vp_id, view, override_types in entry["views"]:
            v_name  = view.Name or "Unnamed View"
            vp_link = output.linkify(vp_id)
            output.print_html(
                "<div class='rev'>-> {} &nbsp; "
                "<span style='color:#E0A040'>[{}]</span></div>".format(
                    vp_link, ", ".join(override_types)))

        print_separator()

elapsed = time.time() - start_time
print_success("Scan complete in {:.2f} seconds".format(elapsed))
print_separator()
print_header("SCRIPT FINISHED")