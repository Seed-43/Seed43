# -*- coding: utf-8 -*-
"""Find which sheet(s) the selected group(s) are placed on."""

from pyrevit import revit, DB, script, forms

doc = revit.doc
output = script.get_output()

# ---------------------------------------------------------------------------
# 1. Get selected groups (model or detail)
# ---------------------------------------------------------------------------
selection = revit.get_selection()
groups = [el for el in selection if isinstance(el, DB.Group)]

if not groups:
    forms.alert("Select one or more Groups first.", exitscript=True)

# ---------------------------------------------------------------------------
# 2. Collect every (sheet, view) pair from actual viewports on sheets
# ---------------------------------------------------------------------------
all_sheets = DB.FilteredElementCollector(doc)\
               .OfClass(DB.ViewSheet)\
               .WhereElementIsNotElementType()\
               .ToElements()

placed_views = []  # list of (ViewSheet, View)
for sheet in all_sheets:
    viewports = DB.FilteredElementCollector(doc, sheet.Id)\
                   .OfClass(DB.Viewport)\
                   .ToElements()
    for vp in viewports:
        view = doc.GetElement(vp.ViewId)
        if view:
            placed_views.append((sheet, view))

# ---------------------------------------------------------------------------
# 3. Cache visible element ids per view, so each view is only collected once
#    even when checking multiple selected groups.
# ---------------------------------------------------------------------------
view_visible_ids = {}

def get_visible_ids(view):
    if view.Id not in view_visible_ids:
        try:
            ids = set(DB.FilteredElementCollector(doc, view.Id).ToElementIds())
        except Exception:
            ids = set()
        view_visible_ids[view.Id] = ids
    return view_visible_ids[view.Id]

# ---------------------------------------------------------------------------
# 4. Match each selected group against every placed view
# ---------------------------------------------------------------------------
results = {}
for group in groups:
    hits = []
    for sheet, view in placed_views:
        if group.Id in get_visible_ids(view):
            hits.append((sheet.SheetNumber, sheet.Name, view.Name))
    results[group.Id] = hits

# ---------------------------------------------------------------------------
# 5. Report
# ---------------------------------------------------------------------------
output.print_md("## Group Sheet Lookup")

for group in groups:
    group_name = group.Name if hasattr(group, "Name") else "Group"
    output.print_md("**{}** (id {})".format(group_name, output.linkify(group.Id)))

    hits = results[group.Id]
    if not hits:
        output.print_md("- Not placed on any sheet")
    else:
        for sheet_number, sheet_name, view_name in hits:
            output.print_md("- Sheet **{}** - {} (via view {})".format(
                sheet_number, sheet_name, view_name))