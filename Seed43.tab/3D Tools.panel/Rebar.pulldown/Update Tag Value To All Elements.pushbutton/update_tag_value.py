# -*- coding: utf-8 -*-
# update_tag_value.py
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
