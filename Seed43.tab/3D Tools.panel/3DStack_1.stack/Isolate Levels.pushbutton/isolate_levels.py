# -*- coding: utf-8 -*-
# isolate_levels.py
# "Isolate Levels"
# "Seed43"
# """
# Temporarily isolate everything tied to a chosen level in the active view,
# to see what deleting that level would take with it. Revit warns that
# elements will be deleted but never says which ones.
# """
from pyrevit import revit, DB, forms, script
from System.Collections.Generic import List

# ── [LIB] Snippets/_selection.py ─────────────────────────────────────────────
from Snippets._selection import get_levels

doc = revit.doc

# ── CONSTANTS ───────────────────────────────────────────────────────────────

# Every way an element can be tied to a level. Element.LevelId covers the
# common case, but plenty of categories record their level in a parameter
# instead and report InvalidElementId - stairs, roofs, and anything with a
# separate base and top.
#
# Deliberately association, not elevation: this tool exists to answer "what
# do I lose if I delete this level", and Revit deletes what is ASSOCIATED
# with a level, not what happens to sit at that height. A beam tied to Level
# 1 but drawn at Level 2 goes with Level 1; a generic model floating at the
# same height with no level at all is untouched.
#
# Resolved by name at runtime: the set of BuiltInParameters differs between
# Revit versions, so a name missing on this one is skipped rather than
# breaking the whole lookup.
_LEVEL_PARAM_NAMES = [
    'LEVEL_PARAM',
    'FAMILY_LEVEL_PARAM',
    'SCHEDULE_LEVEL_PARAM',
    'FAMILY_BASE_LEVEL_PARAM',
    'FAMILY_TOP_LEVEL_PARAM',
    'WALL_BASE_CONSTRAINT',
    'WALL_HEIGHT_TYPE',
    'ROOF_BASE_LEVEL_PARAM',
    'ROOF_UPTO_LEVEL_PARAM',
    'STAIRS_BASE_LEVEL_PARAM',
    'STAIRS_TOP_LEVEL_PARAM',
    'INSTANCE_REFERENCE_LEVEL_PARAM',
    'INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM',
    'GROUP_LEVEL',
]


def _level_params():
    """The BuiltInParameters above that exist on this Revit version."""
    found = []
    for name in _LEVEL_PARAM_NAMES:
        try:
            found.append(getattr(DB.BuiltInParameter, name))
        except AttributeError:
            continue  # not in this version, nothing to do about it
    return found


LEVEL_PARAMS = _level_params()


# ── LEVEL SELECTION ─────────────────────────────────────────────────────────

def select_level(levels):
    """Present a dialog to select a level, or None if cancelled."""
    level_options = {level.Name: level for level in levels}
    selected_name = forms.SelectFromList.show(
        sorted(level_options.keys()),
        title="Select a Level",
        multiselect=False
    )
    return level_options.get(selected_name)


# ── ELEMENT COLLECTION ──────────────────────────────────────────────────────

def _is_tied_to(element, level_id):
    """True if the element is associated with this level by any route.

    Checks Element.LevelId first, then every level parameter, because the two
    disagree often enough to matter: stairs, roofs and multi-level families
    frequently report InvalidElementId while still holding the level in a
    parameter.
    """
    try:
        if element.LevelId == level_id:
            return True
    except Exception:
        pass

    for bip in LEVEL_PARAMS:
        try:
            param = element.get_Parameter(bip)
            if param and param.StorageType == DB.StorageType.ElementId:
                if param.AsElementId() == level_id:
                    return True
        except Exception:
            continue
    return False


def get_elements_on_level(level, view):
    """Every element in the view tied to the given level.

    Collected from the view rather than the document, so what is isolated
    matches what is on screen - isolating something the view never showed
    would be invisible anyway.
    """
    matched = []
    for element in (DB.FilteredElementCollector(doc, view.Id)
                    .WhereElementIsNotElementType()):
        try:
            if not element.CanBeHidden(view):
                continue  # pointless to isolate what the view cannot hide
        except Exception:
            continue
        if _is_tied_to(element, level.Id):
            matched.append(element)
    return matched


# ── ISOLATION ───────────────────────────────────────────────────────────────

def isolate_elements(elements, view):
    """Temporarily isolate the elements in the given view."""
    element_ids = List[DB.ElementId]([element.Id for element in elements])
    with revit.Transaction("Isolate Level"):
        view.IsolateElementsTemporary(element_ids)


# ── MAIN ────────────────────────────────────────────────────────────────────

def main():
    view = doc.ActiveView

    # Schedules, sheets and legends have no temporary view modes, and asking
    # anyway throws rather than failing quietly.
    if not view.CanUseTemporaryVisibilityModes():
        forms.alert("This view does not support temporary isolate.\n\n"
                    "Open a plan, section or 3D view and try again.",
                    title="Isolate Levels")
        script.exit()

    levels = get_levels(doc)
    if not levels:
        forms.alert("No levels found in this model.", title="Isolate Levels")
        script.exit()

    selected_level = select_level(levels)
    if not selected_level:
        script.exit()  # cancelled, no need to tell them what they just did

    elements = get_elements_on_level(selected_level, view)

    # Isolating nothing hides the entire view, which looks like a crash.
    # It is also the answer worth knowing: nothing is tied to this level,
    # so it is safe to delete.
    if not elements:
        forms.alert("Nothing is tied to level: {}\n\n"
                    "Safe to delete, as far as this view can see."
                    .format(selected_level.Name), title="Isolate Levels")
        script.exit()

    isolate_elements(elements, view)
    forms.alert("Isolated {} elements tied to level: {}\n\n"
                "These are what deleting the level would affect."
                .format(len(elements), selected_level.Name),
                title="Isolate Levels")


if __name__ == "__main__":
    main()
