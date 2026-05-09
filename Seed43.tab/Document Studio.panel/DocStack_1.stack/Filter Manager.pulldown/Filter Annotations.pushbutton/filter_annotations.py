# -*- coding: utf-8 -*-
__title__  = "Filter Annotations"
__author__  = "Seed43"
__doc__     = """
𝐕𝐄𝐑𝐒𝐈𝐎𝐍 𝟐𝟔𝟎𝟓𝟎𝟏
_____________________________________________________________________
Description:
Creates smart View Filters for annotation and datum elements directly
from user selection.

- Select any annotation or datum element in the model
- Automatically detects the best parameter to build a filter
- Creates a filter matching:
    - Category
    - Type, Name, View, or other valid parameters
- Applies the filter to the active view or its template
- Automatically hides matching elements

Smart parameter detection:
- Tests multiple built-in parameters in priority order
- Verifies parameter validity before creating the filter
- Falls back safely if the parameter is not valid for the category

Fallback system:
- If no valid parameter is found:
    - Creates a category-only filter (affects ALL elements in category)
    - Prompts user before applying

Designed for fast isolation and cleanup of annotation graphics.
_____________________________________________________________________
How-to:
-> Run the tool
-> Select an annotation or datum element
   (e.g. section line, level, grid, tag, etc.)

-> Tool will:
    - Detect category and type
    - Find a valid filter parameter
    - Create a filter automatically

-> If filter already exists:
    - Use Existing, applies it
    - Create New, duplicates with unique name
    - Skip, does nothing

-> Result:
    - Filter is applied to the current view (or view template)
    - Matching elements are hidden

-> Press ESC to exit the tool
_____________________________________________________________________
Notes:
- Works on annotation and datum elements, NOT host model elements
- If you select a model element:
    - The tool will warn and suggest the correct workflow

- View Template support:
    - If a template controls filters, the filter is applied to the
      template instead

- Category-only filters:
    - Will hide ALL elements of that category in the view
    - Use with caution

- Parameter detection is dynamic:
    - Not all categories support all parameters
    - The tool tests validity before creating the filter
_____________________________________________________________________
Last update:
- Initial release
_____________________________________________________________________
"""

# pylint: disable=import-error,invalid-name,broad-except

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from System.Collections.Generic import List
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *
from pyrevit import revit, forms, script

# ── [LIB] Snippets/_filters.py ───────────────────────────────────────────────
from Snippets._filters import (
    find_existing_filter,
    get_unique_filter_name,
    apply_filter_to_target,
    create_category_filter,
)

# ── [LIB] Snippets/_selection.py ─────────────────────────────────────────────
from Snippets._selection import get_element_type, get_element_category

doc    = revit.doc
uidoc  = revit.uidoc
logger = script.get_logger()

FILTER_NAME_FORMAT     = "Annotation - {0} ({1})"
FILTER_NAME_FORMAT_ALL = "Annotation - {0} (All)"


# ── VARIABLES ─────────────────────────────────────────────────────────────────
# Hardcoded param IDs confirmed working from debug output.
# Ordered list of (param_integer_id, use_AsValueString) to try in order.

PARAM_CANDIDATES = [
    (-1002050, False),   # Type, e.g. "Main Section", "5mm Bubble", "Story"
    (-1002001, False),   # Type Name
    (-1002052, False),   # Family and Type, e.g. "Section: Main Section"
    (-1005112, False),   # View Name (for views)
    (int(BuiltInParameter.ALL_MODEL_TYPE_NAME), False),
    (int(BuiltInParameter.DATUM_TEXT),          False),
    (int(BuiltInParameter.ROOM_NAME),           False),
    (int(BuiltInParameter.ALL_MODEL_MARK),      False),
]


# ── HELPERS ───────────────────────────────────────────────────────────────────

def _safe_bic(name):
    try:
        return getattr(BuiltInCategory, name)
    except AttributeError:
        return None


_HOST_BICS = set(filter(None, [
    _safe_bic("OST_Walls"),             _safe_bic("OST_Floors"),
    _safe_bic("OST_Roofs"),             _safe_bic("OST_Columns"),
    _safe_bic("OST_StructuralColumns"), _safe_bic("OST_Doors"),
    _safe_bic("OST_Windows"),           _safe_bic("OST_Furniture"),
    _safe_bic("OST_GenericModel"),      _safe_bic("OST_MechanicalEquipment"),
    _safe_bic("OST_ElectricalEquipment"), _safe_bic("OST_PlumbingFixtures"),
    _safe_bic("OST_StructuralFraming"), _safe_bic("OST_Stairs"),
    _safe_bic("OST_Ramps"),             _safe_bic("OST_Railings"),
    _safe_bic("OST_Ceilings"),          _safe_bic("OST_CurtainWallPanels"),
    _safe_bic("OST_CurtainWallMullions"),
]))


def is_host_model(category):
    try:
        return category.BuiltInCategory in _HOST_BICS
    except Exception:
        return False


# ── SELECTION FILTER ──────────────────────────────────────────────────────────

class AnnotationSelFilter(ISelectionFilter):
    def AllowElement(self, element):
        return not isinstance(element, RevitLinkInstance)
    def AllowReference(self, reference, xyz):
        return False


# ── FIND BEST PARAMETER ───────────────────────────────────────────────────────

def find_working_param(element, element_type, category):
    """
    Try each param candidate in order.
    Attempts to create a test filter to confirm the param is valid,
    then deletes it and returns the working param and its value.
    Returns (ElementId param_id, str value) or (None, None).
    """
    cat_ids = List[ElementId]([category.Id])

    for (pid_int, use_vs) in PARAM_CANDIDATES:
        try:
            param_id = ElementId(pid_int)
        except Exception:
            continue

        value = None
        for host in [element_type, element]:
            if host is None:
                continue
            try:
                p = host.get_Parameter(param_id)
                if p is None:
                    try:
                        bip = BuiltInParameter(pid_int)
                        p   = host.get_Parameter(bip)
                    except Exception:
                        pass
                if p and p.HasValue:
                    val = p.AsValueString() if use_vs else p.AsString()
                    if val and val.strip():
                        value = val.strip()
                        break
            except Exception:
                pass

        if not value:
            continue

        try:
            provider  = ParameterValueProvider(param_id)
            evaluator = FilterStringEquals()
            rule      = FilterStringRule(provider, evaluator, value)
            ef        = ElementParameterFilter(rule)

            test_name = "__annot_filter_test__"
            for f in FilteredElementCollector(doc).OfClass(ParameterFilterElement):
                if f.Name == test_name:
                    with revit.Transaction("Cleanup test filter"):
                        doc.Delete(f.Id)
                    break

            with revit.Transaction("Test filter param"):
                pf = ParameterFilterElement.Create(doc, test_name, cat_ids, ef)
                doc.Delete(pf.Id)

            return param_id, value

        except Exception as ex:
            logger.debug("Param {} failed for {}: {}", pid_int, category.Name, str(ex))
            continue

    return None, None


# ── CREATE NAME FILTER ────────────────────────────────────────────────────────

def create_name_filter(filter_name, category, param_id, value):
    """Create a ParameterFilterElement using a specific param ID and value."""
    cat_ids   = List[ElementId]([category.Id])
    provider  = ParameterValueProvider(param_id)
    evaluator = FilterStringEquals()
    rule      = FilterStringRule(provider, evaluator, value)
    ef        = ElementParameterFilter(rule)
    return ParameterFilterElement.Create(doc, filter_name, cat_ids, ef)


# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    try:
        with forms.WarningBar(title="Filter Annotation: Select element. ESC to cancel."):
            while True:
                try:
                    picked_ref = uidoc.Selection.PickObject(
                        ObjectType.Element,
                        AnnotationSelFilter(),
                        "Select annotation or datum element"
                    )
                    element = doc.GetElement(picked_ref.ElementId)

                    if isinstance(element, RevitLinkInstance):
                        forms.alert("Please select a host model element.",
                                    title="Invalid Selection")
                        continue

                    category = get_element_category(element, doc)
                    if not category:
                        forms.alert("Element has no valid category.", title="Error")
                        continue

                    if is_host_model(category):
                        opt = forms.alert(
                            "This looks like a host model element ({}).\n\n"
                            "Use 'Filter Element (Host)' instead.\n\n"
                            "Continue anyway?".format(category.Name),
                            options=["Continue", "Cancel"],
                            title="Wrong Script?"
                        )
                        if opt != "Continue":
                            continue

                    element_type = get_element_type(element, doc)

                    param_id, param_value = find_working_param(
                        element, element_type, category)

                    if param_id is not None:
                        filter_name       = FILTER_NAME_FORMAT.format(
                            category.Name, param_value)
                        use_category_only = False
                    else:
                        filter_name       = FILTER_NAME_FORMAT_ALL.format(category.Name)
                        use_category_only = True

                    if use_category_only:
                        opt = forms.alert(
                            "Could not find a working filter parameter for '{}'.\n\n"
                            "The only option is to hide ALL '{}' elements "
                            "in this view.\n\nProceed?".format(
                                category.Name, category.Name),
                            options=["Yes, hide all", "Cancel"],
                            title="Category-Only Filter"
                        )
                        if opt != "Yes, hide all":
                            continue

                    existing = find_existing_filter(filter_name, doc)
                    if existing:
                        opt = forms.alert(
                            "Filter '{}' already exists.\n\nWhat do you want to do?".format(
                                filter_name),
                            options=["Use Existing", "Create New", "Skip"],
                            title="Filter Already Exists"
                        )
                        if opt == "Skip":
                            continue
                        elif opt == "Use Existing":
                            apply_filter_to_target(existing, doc, revit)
                            continue
                        elif opt == "Create New":
                            filter_name = get_unique_filter_name(filter_name, doc)

                    with revit.Transaction("Create Annotation Filter"):
                        if use_category_only:
                            param_filter = create_category_filter(
                                filter_name, category, doc)
                        else:
                            param_filter = create_name_filter(
                                filter_name, category, param_id, param_value)

                        if param_filter:
                            apply_filter_to_target(param_filter, doc, revit)
                        else:
                            forms.alert("Failed to create filter.", title="Error")

                except Exception as ex:
                    ex_str = str(ex).lower()
                    if any(w in ex_str for w in
                           ("cancelled", "aborted", "operation was cancelled")):
                        break
                    logger.error("Error: {}", str(ex))
                    forms.alert(
                        "An error occurred:\n\n{}\n\nCheck the pyRevit console.".format(
                            str(ex)),
                        title="Error"
                    )
                    break

    except Exception as ex:
        ex_str = str(ex).lower()
        if any(w in ex_str for w in ("cancelled", "aborted")):
            return
        logger.error("Error: {}", str(ex))
        forms.alert("Error occurred. Check pyRevit console.", title="Error")


if __name__ == "__main__":
    main()
