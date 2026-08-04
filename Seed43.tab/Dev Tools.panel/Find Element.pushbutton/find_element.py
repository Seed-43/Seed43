# -*- coding: utf-8 -*-
"""
Find which sheet(s) every element of the selected element's type is placed on.

A selection is expanded to all placed instances before searching, so picking a
Filled Region type in the Project Browser reports every placed filled region of
that type, not just one. Works with categories, families, family types, system
types (FilledRegionType, WallType, ...), groups, and plain instances.

Each result is linkified, so clicking an id selects and zooms that element.
"""

from pyrevit import revit, DB, script, forms

doc = revit.doc
output = script.get_output()

# ---------------------------------------------------------------------------
# Output styles (dark theme, matching view_organiser.py)
# ---------------------------------------------------------------------------
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
.element {
    color: #F4FAFF;
    font-weight: bold;
    padding-left: 5px;
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
    padding-left: 15px;
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

def print_element(text):
    output.print_html("<div class='element'>{}</div>".format(text))

def print_sheet(text):
    output.print_html("<div class='sheet'>{}</div>".format(text))

def print_warn(text):
    output.print_html("<div class='warn'>{}</div>".format(text))

# ---------------------------------------------------------------------------
# Helper: Safely get 'Name' from Revit elements
# ---------------------------------------------------------------------------
def safe_get_name(obj):
    """Try multiple ways to get a readable name from a Revit object.

    Everything returns unicode: type names routinely carry non-ASCII (a Filled
    Region called "Solid – 50%"), and a bare str() on a .NET string containing
    those raises UnicodeEncodeError under IronPython 2. Same reasoning as
    _revisions.safe_str().
    """
    if obj is None:
        return u"Unknown"

    # basestring, not str - a .NET name can arrive as unicode
    try:
        name = getattr(obj, 'Name', None)
        if name and isinstance(name, basestring) and name.strip():
            return unicode(name)
    except Exception:
        pass

    try:
        str_repr = unicode(obj)
        if str_repr and str_repr.strip():
            return str_repr
    except Exception:
        pass

    try:
        if hasattr(obj, 'Id'):
            return u"ID: {}".format(obj.Id)
    except Exception:
        pass

    return u"Unknown Element"

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
    try:
        selected_elements = revit.pick_elements(
            "Select one or more elements to search for")
    except Exception:
        selected_elements = []

if not selected_elements:
    forms.alert("No elements were selected.", exitscript=True)

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
                display_map[item.Id] = u"{} – {}".format(cat_name, item_name)
            if not collector:
                display_map[cat_id] = u"Category: {} (no instances)".format(cat_name)
            continue
    except Exception:
        pass

    # --- Family (project browser Families node) ---------------------------
    try:
        if isinstance(el, DB.Family):
            fam_name = safe_get_name(el)
            # Collect once and match against the symbol set - re-running the
            # collector per symbol is a full model sweep each time.
            symbol_ids = set(el.GetFamilySymbolIds())
            instances = DB.FilteredElementCollector(doc)\
                          .OfClass(DB.FamilyInstance)\
                          .WhereElementIsNotElementType()\
                          .ToElements()
            for inst in instances:
                if inst.Symbol.Id in symbol_ids:
                    search_ids.add(inst.Id)
                    inst_name = safe_get_name(inst)
                    display_map[inst.Id] = u"Family {} – {}".format(fam_name, inst_name)
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
                    display_map[inst.Id] = u"Type {} – {}".format(sym_name, inst_name)
            continue
    except Exception:
        pass

    # --- Any other ElementType selected directly (e.g. Project Browser ----
    # type node for a system family: FilledRegionType, WallType, FloorType,
    # etc). These are not FamilySymbols but are still just a type, not an
    # instance, so expand to every instance of that type.
    try:
        if isinstance(el, DB.ElementType) and not isinstance(el, DB.FamilySymbol):
            type_name = safe_get_name(el)
            cat_name = safe_get_name(el.Category) if hasattr(el, 'Category') and el.Category else "Unknown"

            same_type_collector = DB.FilteredElementCollector(doc)\
                                     .WhereElementIsNotElementType()
            if hasattr(el, 'Category') and el.Category:
                same_type_collector = same_type_collector.OfCategoryId(el.Category.Id)

            same_elements = [e for e in same_type_collector.ToElements()
                              if e.GetTypeId() == el.Id]

            if same_elements:
                for same_el in same_elements:
                    search_ids.add(same_el.Id)
                    same_name = safe_get_name(same_el)
                    display_map[same_el.Id] = u"{} ({}) – {}".format(
                        cat_name, type_name, same_name)
            else:
                search_ids.add(el.Id)
                display_map[el.Id] = u"{}: {} (no instances in model)".format(
                    cat_name, type_name)
            continue
    except Exception:
        pass

    # --- Group ------------------------------------------------------------
    try:
        if isinstance(el, DB.Group):
            grp_name = safe_get_name(el)
            search_ids.add(el.Id)
            display_map[el.Id] = u"Group: {}".format(grp_name)
            continue
    except Exception:
        pass

    # --- Any other element (walls, doors, views, sheets, etc.) -------------
    # Expand to every other element sharing the same type, not just the
    # one that was picked.
    try:
        cat_name = safe_get_name(el.Category) if hasattr(el, 'Category') else "Unknown"

        type_id = DB.ElementId.InvalidElementId
        try:
            type_id = el.GetTypeId()
        except Exception:
            pass

        if type_id and type_id != DB.ElementId.InvalidElementId:
            type_elem = doc.GetElement(type_id)
            type_name = safe_get_name(type_elem)

            same_type_collector = DB.FilteredElementCollector(doc)\
                                     .WhereElementIsNotElementType()
            if hasattr(el, 'Category') and el.Category:
                same_type_collector = same_type_collector.OfCategoryId(el.Category.Id)

            same_elements = [e for e in same_type_collector.ToElements()
                              if e.GetTypeId() == type_id]

            if not same_elements:
                same_elements = [el]

            for same_el in same_elements:
                search_ids.add(same_el.Id)
                same_name = safe_get_name(same_el)
                display_map[same_el.Id] = u"{} ({}) – {}".format(
                    cat_name, type_name, same_name)
        else:
            # No type to match against, fall back to just this element
            elem_name = safe_get_name(el)
            search_ids.add(el.Id)
            display_map[el.Id] = u"{}: {}".format(cat_name, elem_name)
    except Exception:
        search_ids.add(el.Id)
        display_map[el.Id] = u"Element: {}".format(el.Id)

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
    results[eid] = []

# Walk views once and intersect, rather than re-scanning every view per id -
# expanding a type to all its instances can put thousands of ids in search_ids.
for sheet, view in placed_views:
    matched = search_ids.intersection(get_visible_ids(view))
    if not matched:
        continue
    try:
        entry = (sheet.SheetNumber, sheet.Name, view.Name)
    except Exception:
        entry = ("?", "?", "?")
    for eid in matched:
        results[eid].append(entry)

# ---------------------------------------------------------------------------
# 6. Report
# ---------------------------------------------------------------------------
print_header("Element &rarr; Sheet Lookup")

total_found = 0
total_not_found = 0

for eid in sorted(results, key=lambda x: display_map.get(x, "")):
    label = display_map.get(eid, "Element {}".format(eid))
    hits = results[eid]

    print_element("{} (id {})".format(label, output.linkify(eid)))

    if not hits:
        print_warn("Not placed on any sheet")
        total_not_found += 1
    else:
        for sheet_number, sheet_name, view_name in hits:
            print_sheet("Sheet {} - {} (via view {})".format(
                sheet_number, sheet_name, view_name))
        total_found += 1

# Summary
print_separator()
print_element("Summary: {} element(s) found on sheets, "
               "{} element(s) not on any sheet".format(total_found,
                                                         total_not_found))