# -*- coding: utf-8 -*-
__title__  = "Update Tag Value To All Elements"
__author__  = "Seed43"
__doc__     = """
𝐕𝐄𝐑𝐒𝐈𝐎𝐍 𝟐𝟔𝟎𝟓𝟎𝟏
_____________________________________________________________________
Description:
Picks a tag, reads a text parameter from the element it is tagging,
then finds every element in the model that has the same value for
that parameter and updates them all to a new value you type in.

Useful for fixing a parameter value that was applied inconsistently
across many elements.
_____________________________________________________________________
How-to:
-> Run the tool
-> Click on a tag in the model
-> Choose which text parameter to edit from the list
-> Type the new value you want to apply
-> All elements with the old value are updated automatically
_____________________________________________________________________
Notes:
- Only text-type parameters are shown in the list
- Read-only parameters cannot be updated and are excluded
- The tag must reference a host element (not a linked element)
_____________________________________________________________________
Last update:
- Initial release
_____________________________________________________________________
"""

from pyrevit import revit, DB, forms
import sys

doc = revit.doc


# ── FUNCTIONS ─────────────────────────────────────────────────────────────────

def get_param_value(param):
    """Return the string value of a parameter, or None if not a string."""
    try:
        if param.StorageType == DB.StorageType.String:
            return param.AsString()
    except Exception:
        pass
    return None


# ── PICK TAG ─────────────────────────────────────────────────────────────────

tag_elem = revit.pick_element(message="Select a Tag element")
if not tag_elem:
    forms.alert("No tag element selected. Exiting.")
    sys.exit()

try:
    host_ids = list(tag_elem.GetTaggedLocalElementIds())
except Exception:
    forms.alert("Selected element is not a tag or is an unsupported tag type.")
    sys.exit()

if not host_ids:
    forms.alert("Tag does not reference any host elements.")
    sys.exit()

host_elem = doc.GetElement(host_ids[0])
if not host_elem:
    forms.alert("Could not find the host element for the selected tag.")
    sys.exit()


# ── SELECT PARAMETER ─────────────────────────────────────────────────────────

text_params = [
    p for p in host_elem.Parameters
    if p.StorageType == DB.StorageType.String and not p.IsReadOnly
]

if not text_params:
    forms.alert("No editable text parameters found on the host element.")
    sys.exit()

param_names        = [p.Definition.Name for p in text_params]
selected_param_name = forms.SelectFromList.show(
    param_names, title="Select Text Parameter to Edit")

if not selected_param_name:
    forms.alert("No parameter selected. Exiting.")
    sys.exit()

selected_param = next(
    (p for p in text_params if p.Definition.Name == selected_param_name),
    None
)

old_value = get_param_value(selected_param)


# ── GET NEW VALUE ─────────────────────────────────────────────────────────────

new_value = forms.ask_for_string(
    default=old_value if old_value else "",
    prompt="Old value: '{}'. Enter new value to set:".format(old_value)
)
if new_value is None:
    forms.alert("No new value entered. Exiting.")
    sys.exit()


# ── FIND MATCHING ELEMENTS ────────────────────────────────────────────────────

elements_to_update = []
for elem in DB.FilteredElementCollector(doc).WhereElementIsNotElementType():
    param = elem.LookupParameter(selected_param_name)
    if param and param.StorageType == DB.StorageType.String:
        if param.AsString() == old_value:
            elements_to_update.append(elem)


# ── UPDATE ────────────────────────────────────────────────────────────────────

t = DB.Transaction(doc, "Update Text Parameter Values")
t.Start()
for elem in elements_to_update:
    param = elem.LookupParameter(selected_param_name)
    if param and not param.IsReadOnly:
        try:
            param.Set(new_value)
        except Exception:
            pass
t.Commit()

forms.alert(
    "Updated {} elements with parameter '{}' from '{}' to '{}'.".format(
        len(elements_to_update), selected_param_name, old_value, new_value)
)
