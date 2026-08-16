# -*- coding: utf-8 -*-
# update_tag_value.py
#
# Picks a tag, reads a text parameter from the element it tags, then finds
# every element in the model sharing that value and updates them all at once.

from pyrevit import revit, DB, forms, script

doc = revit.doc

try:
    from Snippets import _dialogs as sdlg
except Exception:
    sdlg = None

TITLE = "Update Tag Value"

# ── DIALOGS ───────────────────────────────────────────────────────────────────

def _alert(message, title=TITLE, exitscript=False):
    """Themed popup via the shared Snippets dialog lib, falls back to
    pyRevit's default forms.alert if the shared lib isn't available."""
    if sdlg:
        sdlg.message(message, title=title)
    else:
        forms.alert(message, title=title)
    if exitscript:
        script.exit()


def _confirm(message, title=TITLE, no="Cancel"):
    """Themed yes/no popup, returns True on yes."""
    if sdlg:
        return sdlg.confirm(message, title=title, no=no)
    return bool(forms.alert(message, title=title, ok=False, yes=True, no=True))

# ── HELPERS ───────────────────────────────────────────────────────────────────

def eid_int(element_id):
    try:
        return element_id.Value          # Revit 2024+
    except AttributeError:
        return element_id.IntegerValue   # Revit 2023 and earlier


def get_param_value(param):
    """Return the string value of a parameter, or None if not a string."""
    try:
        if param.StorageType == DB.StorageType.String:
            return param.AsString()
    except Exception:
        pass
    return None


def param_by_id(element, param_id_value):
    """Return the parameter on this element carrying param_id_value, or None.

    Matched on id rather than name so the update stays in step with the
    id-based filter that found the elements in the first place.
    """
    for p in element.Parameters:
        try:
            if eid_int(p.Id) == param_id_value:
                return p
        except Exception:
            pass
    return None

# ── PICK TAG ─────────────────────────────────────────────────────────────────

tag_elem = revit.pick_element(message="Select a Tag element")
if not tag_elem:
    _alert("No tag element selected. Exiting.", exitscript=True)

try:
    host_ids = list(tag_elem.GetTaggedLocalElementIds())
except Exception:
    _alert("Selected element is not a tag or is an unsupported tag type.",
           exitscript=True)

if not host_ids:
    _alert("Tag does not reference any host elements.", exitscript=True)

host_elem = doc.GetElement(host_ids[0])
if not host_elem:
    _alert("Could not find the host element for the selected tag.",
           exitscript=True)

# ── SELECT PARAMETER ─────────────────────────────────────────────────────────

text_params = [
    p for p in host_elem.Parameters
    if p.StorageType == DB.StorageType.String and not p.IsReadOnly
]

if not text_params:
    _alert("No editable text parameters found on the host element.",
           exitscript=True)

param_names         = [p.Definition.Name for p in text_params]
selected_param_name = forms.SelectFromList.show(
    param_names, title="Select Text Parameter to Edit")

if not selected_param_name:
    _alert("No parameter selected. Exiting.", exitscript=True)

selected_param = next(
    (p for p in text_params if p.Definition.Name == selected_param_name),
    None
)

old_value = get_param_value(selected_param)

# An unset text parameter reads back as None, and matching on that would sweep
# up every element in the model whose parameter is simply blank -- on something
# common like Comments that is most of the model, overwritten in one click with
# no warning. There is nothing safe to search for, so stop here.
if old_value is None or not old_value.strip():
    _alert(
        "'{}' is empty on the tagged element.\n\n"
        "Searching for a blank value would match every element in the model "
        "with that parameter unset, so there is nothing to update from."
        .format(selected_param_name),
        exitscript=True)

# ── GET NEW VALUE ─────────────────────────────────────────────────────────────

new_value = forms.ask_for_string(
    default=old_value,
    prompt="Old value: '{}'. Enter new value to set:".format(old_value)
)
if new_value is None:
    _alert("No new value entered. Exiting.", exitscript=True)

if new_value == old_value:
    _alert("The new value matches the old value. Nothing to do.",
           exitscript=True)

# ── FIND MATCHING ELEMENTS ────────────────────────────────────────────────────

# An ElementParameterFilter pushes the match down into Revit's own filtering.
# The previous version walked every element in the document and called
# LookupParameter on each one, which is a linear parameter scan per element.
#
# NOTE: this matches on the parameter's id rather than its name. Built-in,
# shared and project parameters carry the same id on every element, so the
# usual case behaves exactly as before; a family parameter that merely shares
# a name with the tagged element's one is no longer swept in.
provider     = DB.ParameterValueProvider(selected_param.Id)
rule         = DB.FilterStringRule(provider, DB.FilterStringEquals(), old_value)
param_filter = DB.ElementParameterFilter(rule)

elements_to_update = list(
    DB.FilteredElementCollector(doc)
      .WhereElementIsNotElementType()
      .WherePasses(param_filter)
)

if not elements_to_update:
    _alert("No elements found with '{}' set to '{}'.".format(
        selected_param_name, old_value), exitscript=True)

if not _confirm(
        "{} element(s) have '{}' set to '{}'.\n\n"
        "Change them all to '{}'?".format(
            len(elements_to_update), selected_param_name,
            old_value, new_value),
        no="Cancel"):
    script.exit()

# ── UPDATE ────────────────────────────────────────────────────────────────────

param_id_value = eid_int(selected_param.Id)

updated = 0
skipped = 0

t = DB.Transaction(doc, "Update Text Parameter Values")
t.Start()

try:
    for elem in elements_to_update:
        param = param_by_id(elem, param_id_value)
        if param is None or param.IsReadOnly:
            skipped += 1
            continue
        try:
            if param.Set(new_value):
                updated += 1
            else:
                skipped += 1
        except Exception:
            skipped += 1
    t.Commit()
except Exception as e:
    if t.HasStarted():
        t.RollBack()
    _alert("Error: {}".format(str(e)), exitscript=True)

# Report what actually changed -- the old summary counted candidates found,
# not writes that succeeded, so read-only and other-user-owned elements were
# reported as updated.
summary = ["Updated {} element(s): '{}' from '{}' to '{}'.".format(
    updated, selected_param_name, old_value, new_value)]
if skipped:
    summary.append(
        "{} skipped - read-only, or owned by another user.".format(skipped))

_alert("\n\n".join(summary))
