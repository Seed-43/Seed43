# -*- coding: utf-8 -*-
__title__  = "_selection"
__author__  = "Seed43"
__doc__     = """
VERSION 260507
_____________________________________________________________________
Description:
Shared helper functions for picking and resolving elements in Revit.

Used across Filter Manager, CAD Layer Manager, and Isolate Levels.
Import the functions you need.

Example:
    from Snippets._selection import get_element_type, get_element_category
    from Snippets._selection import resolve_cad_instance
    from Snippets._selection import get_levels
_____________________________________________________________________
Functions:
get_element_type(element, doc)
    Return the ElementType for a given element.

get_element_type_linked(sourcedoc, element)
    Same as get_element_type but reads from a linked document.

get_element_category(element, doc)
    Return the Category for an element, falling back to its type.

resolve_cad_instance(uidoc, doc, revit, forms, script)
    Return the CAD ImportInstance to work with, from selection or pick.

get_levels(doc)
    Return all levels in the project.
_____________________________________________________________________
Last update:
- Initial release, extracted from Filter Manager scripts, CAD Layer
  Manager, and Isolate Levels
_____________________________________________________________________
"""

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
