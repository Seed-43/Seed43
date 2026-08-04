# -*- coding: utf-8 -*-
"""
Find which sheet(s) the selected element(s) are placed on.

Works with any element selected in the Project Browser, a view,
the model, or a sheet — including groups, categories, family types,
individual elements, etc.
"""

from pyrevit import revit, DB, script, forms

doc = revit.doc
output = script.get_output()

# ---------------------------------------------------------------------------
# Helper: Safely get 'Name' from Revit elements
# ---------------------------------------------------------------------------
def safe_get_name(obj):
    """Try multiple ways to get a readable name from a Revit object."""
    if obj is None:
        return "Unknown"
    
    # Try getting name attribute
    try:
        name = getattr(obj, 'Name', None)
        if name and isinstance(name, str) and name.strip():
            return name
    except Exception:
        pass
    
    # Try built-in string representation
    try:
        str_repr = str(obj)
        if str_repr and str_repr.strip():
            return str_repr
    except Exception:
        pass
    
    # Fall back to ID
    try:
        if hasattr(obj, 'Id'):
            return "ID: {}".format(obj.Id)
    except Exception:
        pass
    
    return "Unknown Element"

# ---------------------------------------------------------------------------
# 1. Get all currently selected elements (anything the user picked)
# ---------------------------------------------------------------------------
selection = revit.get_selection()
selected_elements = []
try:
    if selection.elements:
        selected_elements = [el for el in selection.elements]
except Exception:
    pass

if not selected_elements:
    forms.alert("Select one or more elements first, then run the script.",
                exitscript=True)

# ---------------------------------------------------------------------------
# 2. Resolve "search targets" from the selection.
# ---------------------------------------------------------------------------
search_ids = set()          # element IDs we actually look for on sheets
display_map = {}            # ElementId -> human-readable label for the report

for el in selected_elements:

    # --- Category selected (project browser "Categories" node) ------------
    try:
        if isinstance(el, DB.Category):
            cat_name = safe_get_name(el)
            cat_id = el.Id
            collector = DB.FilteredElementCollector(doc)\
                           .OfCategoryId(cat_id)\
                           .WhereElementIsNotElementType()\
                           .ToElements()
            for item in collector:
                search_ids.add(item.Id)
                item_name = safe_get_name(item)
                display_map[item.Id] = "{} – {}".format(cat_name, item_name)
            if not collector:
                display_map[cat_id] = "Category: {} (no instances)".format(cat_name)
            continue
    except Exception:
        pass

    # --- Family (project browser Families node) ---------------------------
    try:
        if isinstance(el, DB.Family):
            fam_name = safe_get_name(el)
            symbols = el.GetFamilySymbolIds()
            for sym_id in symbols:
                instances = DB.FilteredElementCollector(doc)\
                              .OfClass(DB.FamilyInstance)\
                              .WhereElementIsNotElementType()\
                              .ToElements()
                for inst in instances:
                    if inst.Symbol.Id == sym_id:
                        search_ids.add(inst.Id)
                        inst_name = safe_get_name(inst)
                        display_map[inst.Id] = "Family {} – {}".format(fam_name, inst_name)
            continue
    except Exception:
        pass

    # --- FamilySymbol (family type selected) ------------------------------
    try:
        if isinstance(el, DB.FamilySymbol):
            sym_name = safe_get_name(el)
            instances = DB.FilteredElementCollector(doc)\
                          .OfClass(DB.FamilyInstance)\
                          .WhereElementIsNotElementType()\
                          .ToElements()
            for inst in instances:
                if inst.Symbol.Id == el.Id:
                    search_ids.add(inst.Id)
                    inst_name = safe_get_name(inst)
                    display_map[inst.Id] = "Type {} – {}".format(sym_name, inst_name)
            continue
    except Exception:
        pass

    # --- Group ------------------------------------------------------------
    try:
        if isinstance(el, DB.Group):
            grp_name = safe_get_name(el)
            search_ids.add(el.Id)
            display_map[el.Id] = "Group: {}".format(grp_name)
            continue
    except Exception:
        pass

    # --- Any other element (walls, doors, views, sheets, etc.) -------------
    try:
        cat_name = safe_get_name(el.Category) if hasattr(el, 'Category') else "Unknown"
        elem_name = safe_get_name(el)
        search_ids.add(el.Id)
        display_map[el.Id] = "{}: {}".format(cat_name, elem_name)
    except Exception:
        display_map[el.Id] = "Element: {}".format(el.Id)

if not search_ids:
    forms.alert("No searchable elements were resolved from the selection.",
                exitscript=True)

# ---------------------------------------------------------------------------
# 3. Collect every (sheet, view) pair from actual viewports on sheets
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
        try:
            view = doc.GetElement(vp.ViewId)
            if view:
                placed_views.append((sheet, view))
        except Exception:
            pass

# ---------------------------------------------------------------------------
# 4. Cache visible element IDs per view (collected once per view)
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
# 5. Match each search ID against every placed view
# ---------------------------------------------------------------------------
results = {}  # ElementId -> list of (sheet_number, sheet_name, view_name)
for eid in search_ids:
    hits = []
    for sheet, view in placed_views:
        if eid in get_visible_ids(view):
            try:
                hits.append((sheet.SheetNumber, sheet.Name, view.Name))
            except Exception:
                hits.append(("?", "?", "?"))
    results[eid] = hits

# ---------------------------------------------------------------------------
# 6. Report
# ---------------------------------------------------------------------------
output.print_md("## Element → Sheet Lookup")

total_found = 0
total_not_found = 0

for eid in sorted(results, key=lambda x: display_map.get(x, "")):
    label = display_map.get(eid, "Element {}".format(eid))
    hits = results[eid]

    output.print_md("**{}** (id {})".format(label, output.linkify(eid)))

    if not hits:
        output.print_md("- Not placed on any sheet")
        total_not_found += 1
    else:
        for sheet_number, sheet_name, view_name in hits:
            output.print_md("- Sheet **{}** – {} (via view {})".format(
                sheet_number, sheet_name, view_name))
        total_found += 1

# Summary
output.print_md("---")
output.print_md("**Summary:** {} element(s) found on sheets · "
                "{} element(s) not on any sheet".format(total_found,
                                                         total_not_found))