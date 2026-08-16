# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import ElementType, ElementId, ImportInstance
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType


# ── ELEMENT TYPE ──────────────────────────────────────────────────────────────

def get_element_type(element, doc):
    """
    Return the ElementType for a given element.

    If the element is already an ElementType it is returned directly.
    Returns None if the type cannot be resolved.
    """
    if isinstance(element, ElementType):
        return element
    type_id = element.GetTypeId()
    if type_id != ElementId.InvalidElementId:
        element_type = doc.GetElement(type_id)
        if isinstance(element_type, ElementType):
            return element_type
    return None


def get_element_type_linked(sourcedoc, element):
    """
    Return the ElementType for an element that lives in a linked document.

    Pass the linked document as sourcedoc. Returns None if unresolvable.
    """
    if isinstance(element, ElementType):
        return element
    type_id = element.GetTypeId()
    if type_id != ElementId.InvalidElementId:
        element_type = sourcedoc.GetElement(type_id)
        if isinstance(element_type, ElementType):
            return element_type
    return None


# ── ELEMENT CATEGORY ──────────────────────────────────────────────────────────

def get_element_category(element, doc):
    """
    Return the Category for an element.

    Checks the element directly first, then falls back to its type.
    Returns None if no category can be found.
    """
    cat = getattr(element, "Category", None)
    if cat:
        return cat
    et = get_element_type(element, doc)
    if et:
        return getattr(et, "Category", None)
    return None


# ── CAD SELECTION ─────────────────────────────────────────────────────────────

class _CADImportFilter(ISelectionFilter):
    """Selection filter that only allows CAD Import or Link instances."""
    def AllowElement(self, element):
        return isinstance(element, ImportInstance)
    def AllowReference(self, reference, xyz):
        return False


def resolve_cad_instance(uidoc, doc, revit, forms, script):
    """
    Return the CAD ImportInstance the user wants to work with.

    Handles three cases:
    - Exactly one CAD instance is already selected, use it.
    - Multiple CAD instances are selected, ask which one.
    - Nothing is selected, prompt the user to click on one.

    Exits the script via script.exit() if the user cancels.
    """
    selected_cads = [
        el for el in revit.get_selection()
        if isinstance(el, ImportInstance)
    ]

    if len(selected_cads) == 1:
        return selected_cads[0]

    if len(selected_cads) > 1:
        name_map = {
            "{} (id {})".format(
                el.Category.Name if el.Category else "CAD", el.Id): el
            for el in selected_cads
        }
        chosen_key = forms.ask_for_one_item(
            sorted(name_map.keys()),
            default=sorted(name_map.keys())[0],
            prompt="Multiple CAD files are selected. Which one do you want to manage?",
            title="Select CAD File"
        )
        if not chosen_key:
            script.exit()
        return name_map[chosen_key]

    # Nothing selected, ask the user to pick
    from pyrevit import forms as _forms
    with _forms.WarningBar(title="Pick the CAD Import or Link"):
        try:
            ref = uidoc.Selection.PickObject(
                ObjectType.Element,
                _CADImportFilter(),
                "Click on the CAD file (Import or Link)"
            )
            return doc.GetElement(ref.ElementId)
        except Exception:
            script.exit()


# ── LEVELS ────────────────────────────────────────────────────────────────────

def get_levels(doc):
    """Return all Level elements in the project as a list."""
    from Autodesk.Revit.DB import FilteredElementCollector, Level
    return FilteredElementCollector(doc).OfClass(Level).ToElements()


# ── DATUM PICKING (levels, grids) ─────────────────────────────────────────────

# Shown as the first entry in the list, so picking in the model is a choice
# alongside the names rather than a separate mode to get into.
PICK_IN_MODEL = u'⊕  Select in model…'


class _ClassFilter(ISelectionFilter):
    """Selection filter allowing a single element class."""
    def __init__(self, element_class):
        self._class = element_class

    def AllowElement(self, element):
        return isinstance(element, self._class)

    def AllowReference(self, reference, xyz):
        return False


def _natural_key(name):
    """Sort Level 2 before Level 10, and grid A2 before A10."""
    import re
    return [int(p) if p.isdigit() else p.lower()
            for p in re.split(r'(\d+)', unicode(name))]


def choose_datum(uidoc, doc, revit, forms, element_class, elements, title,
                 pick_prompt):
    """Return one chosen element of element_class, or None if cancelled.

    Three routes, in order of how little work they ask of the user:

      1. Already selected before running - use it, no dialog at all.
      2. Pick from the list of names.
      3. The list's first entry switches to clicking it in the model.

    Route 3 matters most for grids, where the name means little and the
    position means everything. It can fail on its own terms though: levels
    and grids are only clickable in a view that shows them, and a 3D view
    hides both unless they are turned on in Visibility/Graphics. A failed or
    cancelled pick returns None rather than looping, so the user is never
    trapped in a mode they cannot escape.
    """
    selected = [el for el in revit.get_selection()
                if isinstance(el, element_class)]
    if len(selected) == 1:
        return selected[0]

    by_name = {}
    for element in elements:
        by_name[element.Name] = element

    options = [PICK_IN_MODEL] + sorted(by_name.keys(), key=_natural_key)
    chosen = forms.SelectFromList.show(options, title=title, multiselect=False)
    if not chosen:
        return None
    if chosen != PICK_IN_MODEL:
        return by_name.get(chosen)

    try:
        with forms.WarningBar(title=pick_prompt):
            reference = uidoc.Selection.PickObject(
                ObjectType.Element, _ClassFilter(element_class), pick_prompt)
        return doc.GetElement(reference.ElementId)
    except Exception:
        return None  # cancelled, or nothing pickable in this view
