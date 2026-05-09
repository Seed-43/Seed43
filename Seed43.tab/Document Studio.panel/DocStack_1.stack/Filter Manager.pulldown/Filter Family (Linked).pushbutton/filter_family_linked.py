# -*- coding: utf-8 -*-
__title__  = "Filter Family (Linked)"
__author__  = "Seed43"
__doc__     = """
𝐕𝐄𝐑𝐒𝐈𝐎𝐍 𝟐𝟔𝟎𝟓𝟎𝟏
_____________________________________________________________________
Description:
Creates View Filters for elements selected from linked models, based
on family name.

- Select a linked model, then pick elements inside it
- Automatically detects:
    - Category
    - Family name
- Creates a filter matching:
    - Category and Family Name
- Applies the filter to the active view or its template
- Automatically hides matching elements

Persistent workflow:
- Remembers the selected linked model
- Allows repeated element selection without re-picking the link

Designed for fast isolation and control of linked model content.
_____________________________________________________________________
How-to:
-> Run the tool
-> Select a linked model
-> Select an element inside the linked model (use TAB if needed)

-> Tool will:
    - Detect category and family
    - Create a filter based on family name
    - Apply filter to view (or template)

-> If filter already exists:
    - Use Existing, applies it
    - Create New, duplicates with unique name
    - Skip, does nothing

-> Continue selecting elements from the same link
-> Press ESC to exit the tool
_____________________________________________________________________
Notes:
- Works on linked model elements only, NOT host elements
- The linked model must be loaded
- Ensures the selected element belongs to the chosen link

- View Template support:
    - If a template controls filters, the filter is applied to the
      template instead

- Filter logic:
    - Uses the Family Name parameter
    - Matches all instances of that family within the category

- Persistent link selection:
    - Improves speed when working with large linked models
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

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *
from pyrevit import revit, forms, script

# ── [LIB] Snippets/_filters.py ───────────────────────────────────────────────
from Snippets._filters import (
    find_existing_filter,
    get_unique_filter_name,
    apply_filter_to_target,
    create_parameter_filter,
)

# ── [LIB] Snippets/_selection.py ─────────────────────────────────────────────
from Snippets._selection import get_element_type_linked

doc    = revit.doc
uidoc  = revit.uidoc
logger = script.get_logger()

BUILTIN_PARAM      = BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM
FILTER_NAME_FORMAT = "Linked - {0} - {1}"


# ── SELECTION FILTER ──────────────────────────────────────────────────────────

class LinkInstanceFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, RevitLinkInstance)
    def AllowReference(self, reference, xyz):
        return False


# ── FUNCTIONS ─────────────────────────────────────────────────────────────────

def get_family_name(element_type, element):
    """Return the family name, checking multiple sources in order."""
    if element_type:
        param = element_type.get_Parameter(BUILTIN_PARAM)
        if param and param.HasValue:
            val = param.AsString()
            if val:
                return val
        try:
            if hasattr(element_type, "FamilyName") and element_type.FamilyName:
                return element_type.FamilyName
        except Exception:
            pass
    try:
        if hasattr(element, "Symbol") and element.Symbol:
            fam = element.Symbol.Family
            if fam and fam.Name:
                return fam.Name
    except Exception:
        pass
    return "Unknown"


# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    try:
        last_link_instance = None
        last_source_doc    = None

        with forms.WarningBar(
            title="Filter Family (Linked): Select LINK first, then element. ESC to cancel."
        ):
            while True:
                try:
                    if not last_link_instance:
                        picked_link   = uidoc.Selection.PickObject(
                            ObjectType.Element, LinkInstanceFilter(), "Select linked model")
                        link_instance = doc.GetElement(picked_link.ElementId)

                        if not isinstance(link_instance, RevitLinkInstance):
                            forms.alert("Please select a linked model.",
                                        title="Invalid Selection")
                            continue

                        source_doc = link_instance.GetLinkDocument()
                        if not source_doc:
                            forms.alert("Linked model not loaded.", title="Error")
                            continue

                        last_link_instance = link_instance
                        last_source_doc    = source_doc
                    else:
                        link_instance = last_link_instance
                        source_doc    = last_source_doc

                    picked_ref = uidoc.Selection.PickObject(
                        ObjectType.LinkedElement,
                        "Select element in linked model (TAB to highlight)"
                    )

                    if picked_ref.LinkedElementId == ElementId.InvalidElementId:
                        forms.alert("Invalid linked element.", title="Error")
                        continue
                    if picked_ref.ElementId != link_instance.Id:
                        forms.alert("Element not from selected link.", title="Error")
                        continue

                    element      = source_doc.GetElement(picked_ref.LinkedElementId)
                    element_type = get_element_type_linked(source_doc, element)
                    if not element_type:
                        forms.alert("Could not get valid element type.", title="Error")
                        continue

                    category = element.Category or element_type.Category
                    if not category:
                        forms.alert("Element has no valid category.", title="Error")
                        continue

                    family_name = get_family_name(element_type, element)
                    filter_name = FILTER_NAME_FORMAT.format(category.Name, family_name)
                    existing    = find_existing_filter(filter_name, doc)

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

                    with revit.Transaction("Create Parameter Filter"):
                        param_filter = create_parameter_filter(
                            filter_name, category,
                            ElementId(BUILTIN_PARAM), family_name, doc)
                        if param_filter:
                            apply_filter_to_target(param_filter, doc, revit)
                        else:
                            forms.alert("Failed to create filter.", title="Error")

                except Exception as ex:
                    if "cancelled" in str(ex).lower() or "aborted" in str(ex).lower():
                        break
                    logger.error("Error: {}".format(str(ex)))
                    forms.alert("Error occurred. Check pyRevit console.", title="Error")
                    break

    except Exception as ex:
        if "cancelled" in str(ex).lower() or "aborted" in str(ex).lower():
            return
        logger.error("Error: {}".format(str(ex)))
        forms.alert("Error occurred. Check pyRevit console.", title="Error")


if __name__ == "__main__":
    main()
