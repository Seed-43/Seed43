# -*- coding: utf-8 -*-
# isolate_levels.py
from pyrevit import revit, DB, forms, script
from System.Collections.Generic import List

# ── [LIB] Snippets/_selection.py ─────────────────────────────────────────────
from Snippets._selection import get_levels

doc = revit.doc

# ── FUNCTIONS ─────────────────────────────────────────────────────────────────

def select_level(levels):
    """Present a dialog to select a level."""
    level_options = {level.Name: level for level in levels}
    selected_name = forms.SelectFromList.show(
        sorted(level_options.keys()),
        title="Select a Level",
        multiselect=False
    )
    return level_options.get(selected_name)

def get_elements_at_level(level):
    """Get all elements constrained to the specified level."""
    level_filter = DB.ElementLevelFilter(level.Id)
    return DB.FilteredElementCollector(doc)\
        .WherePasses(level_filter)\
        .WhereElementIsNotElementType()\
        .ToElements()

def isolate_elements(elements):
    """Isolate the specified elements in the active view."""
    active_view = doc.ActiveView
    element_ids = List[DB.ElementId]([elem.Id for elem in elements])
    with revit.Transaction("Isolate Elements"):
        active_view.IsolateElementsTemporary(element_ids)

# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    levels         = get_levels(doc)
    selected_level = select_level(levels)

    if not selected_level:
        forms.alert("No level selected. Operation cancelled.")
        script.exit()

    elements_at_level = get_elements_at_level(selected_level)
    isolate_elements(elements_at_level)
    forms.alert("Isolated {} elements on level: {}".format(
        len(elements_at_level), selected_level.Name))

if __name__ == "__main__":
    main()
