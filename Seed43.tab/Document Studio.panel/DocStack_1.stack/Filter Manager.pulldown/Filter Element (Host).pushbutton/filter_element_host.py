# -*- coding: utf-8 -*-
__title__  = "Filter Element (Host)"
__author__  = "Seed43"
__doc__     = """
𝐕𝐄𝐑𝐒𝐈𝐎𝐍 𝟐𝟔𝟎𝟓𝟎𝟏
_____________________________________________________________________
Description:
Creates View Filters for host model elements based on the selected
element type.

- Select any host model element (walls, doors, furniture, etc.)
- Automatically detects:
    - Category
    - Type Name
- Creates a filter matching:
    - Category and Type Name
- Applies the filter to the active view or its template
- Automatically hides matching elements

Designed for fast isolation and control of model elements by type.
_____________________________________________________________________
How-to:
-> Run the tool
-> Select a host model element

-> Tool will:
    - Detect element type
    - Create a filter based on type name
    - Apply filter to view (or template)

-> If filter already exists:
    - Use Existing, applies it
    - Create New, duplicates with unique name
    - Skip, does nothing

-> Result:
    - All matching elements are hidden in the view

-> Press ESC to exit the tool
_____________________________________________________________________
Notes:
- Works on host model elements only, NOT annotation elements
- Linked elements are ignored
- Requires a valid element type and category

- View Template support:
    - If a template is assigned, the filter is applied to the template

- Filter logic:
    - Uses the Type Name parameter
    - Matches all instances of that type within the category
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
from Snippets._selection import get_element_type

doc    = revit.doc
uidoc  = revit.uidoc
logger = script.get_logger()

BUILTIN_PARAM      = BuiltInParameter.SYMBOL_NAME_PARAM
FILTER_NAME_FORMAT = "Host - {0} - {1}"


# ── SELECTION FILTER ──────────────────────────────────────────────────────────

class HostElementFilter(ISelectionFilter):
    def AllowElement(self, element):
        return not isinstance(element, RevitLinkInstance)
    def AllowReference(self, reference, xyz):
        return False


# ── FUNCTIONS ─────────────────────────────────────────────────────────────────

def get_element_name(element_type, element):
    """Return the type name, falling back to element name or param."""
    if hasattr(element_type, "Name") and element_type.Name:
        return element_type.Name
    element_name = getattr(element, "Name", None)
    if element_name:
        return element_name
    param = element_type.get_Parameter(BUILTIN_PARAM)
    if param and param.HasValue:
        return param.AsString()
    return "Unknown"


# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    try:
        with forms.WarningBar(title="Filter Element (Host): Select element. ESC to cancel."):
            while True:
                try:
                    picked_ref = uidoc.Selection.PickObject(
                        ObjectType.Element,
                        HostElementFilter(),
                        "Select host element"
                    )
                    element = doc.GetElement(picked_ref.ElementId)

                    if isinstance(element, RevitLinkInstance):
                        forms.alert("Please select a host model element.",
                                    title="Invalid Selection")
                        continue

                    element_type = get_element_type(element, doc)
                    if not element_type:
                        forms.alert("Could not get valid element type.", title="Error")
                        continue

                    category = element.Category or element_type.Category
                    if not category:
                        forms.alert("Element has no valid category.", title="Error")
                        continue

                    element_name = get_element_name(element_type, element)
                    filter_name  = FILTER_NAME_FORMAT.format(category.Name, element_name)
                    existing     = find_existing_filter(filter_name, doc)

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
                            ElementId(BUILTIN_PARAM), element_name, doc)
                        if param_filter:
                            apply_filter_to_target(param_filter, doc, revit)
                        else:
                            forms.alert("Failed to create filter.", title="Error")

                except Exception as ex:
                    if "cancelled" in str(ex).lower() or "aborted" in str(ex).lower():
                        break
                    logger.error("Error: {}", str(ex))
                    forms.alert("Error occurred. Check pyRevit console.", title="Error")
                    break

    except Exception as ex:
        if "cancelled" in str(ex).lower() or "aborted" in str(ex).lower():
            return
        logger.error("Error: {}", str(ex))
        forms.alert("Error occurred. Check pyRevit console.", title="Error")


if __name__ == "__main__":
    main()
