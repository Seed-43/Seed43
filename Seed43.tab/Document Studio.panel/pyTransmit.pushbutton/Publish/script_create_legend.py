# -*- coding: utf-8 -*-
# script_create_legend.py

# ── IMPORTS ───────────────────────────────────────────────────────────────────

_p = globals().get('PYTRANSMIT_PAYLOAD', {})

from pyrevit.framework import List
from pyrevit import revit, script, DB, forms
from Autodesk.Revit.DB import (
    FilteredElementCollector, CurveElement, TextNote,
    ImageInstance, FilledRegion, ViewFamilyType, ViewFamily,
    ElementTransformUtils, CopyPasteOptions,
    ViewType, ViewDuplicateOption, ViewDrafting,
)
import os

output = script.get_output()
output.hide()


# ── Log capture via payload ───────────────────────────────────────────
_log_lines = PYTRANSMIT_PAYLOAD.get('_log_lines', [])
def _log(msg):
    try: _log_lines.append(str(msg))
    except Exception: pass

class _SilentOutput(object):
    def print_md(self, msg, *a, **kw): _log(msg)
    def hide(self): pass
output = _SilentOutput()
doc    = revit.doc

# ── VARIABLES ─────────────────────────────────────────────────────────────────

TEMP_VIEW_NAME   = "pyTransmit TEMP"
LEGEND_VIEW_NAME = "pyTransmit Document"

# ── STEP 1, RUN DRAFTING VIEW SCRIPT ─────────────────────────────────────────

_script_dir    = os.path.dirname(os.path.abspath(__file__))
_drafting_path = os.path.join(_script_dir, 'script_create_drafting_view.py')

import sys as _sys
_settings_dir = os.path.join(os.path.dirname(_script_dir), 'Settings')
if _settings_dir not in _sys.path:
    _sys.path.insert(0, _settings_dir)
from Dialogs import Dialogs

if not os.path.exists(_drafting_path):
    forms.alert(
        "script_create_drafting_view.py not found at:\n{}".format(_drafting_path),
        exitscript=True
    )

_payload_for_drafting = dict(_p)
_payload_for_drafting['_legend_temp_view_name'] = TEMP_VIEW_NAME

_ns = {
    '__name__':           'drafting_view_for_legend',
    '__file__':           _drafting_path,
    '__builtins__':       __builtins__,
    'PYTRANSMIT_PAYLOAD': _payload_for_drafting,
}
with open(_drafting_path, 'r') as _f:
    _src = _f.read()
try:
    exec(_src, _ns)
except Exception as _e:
    import traceback as _tb
    forms.alert(
        "Error running drafting view script:\n{}".format(_tb.format_exc() or str(_e)),
        exitscript=True
    )

# ── STEP 2, FIND TEMP DRAFTING VIEW ──────────────────────────────────────────

temp_view = None
for v in revit.query.get_elements_by_class(ViewDrafting, doc=doc):
    try:
        if v.Name == TEMP_VIEW_NAME:
            temp_view = v
            break
    except Exception:
        pass

if not temp_view:
    forms.alert(
        "Temp drafting view '{}' not found after generation.".format(TEMP_VIEW_NAME),
        exitscript=True
    )

# ── STEP 3, FIND OR CREATE LEGEND VIEW ───────────────────────────────────────

existing_legend = None
base_legend     = None
for v in revit.query.get_elements_by_class(DB.View, doc=doc):
    try:
        if v.ViewType == ViewType.Legend and not v.IsTemplate:
            if v.Name in (LEGEND_VIEW_NAME, LEGEND_VIEW_NAME + " (Transmittal)"):
                existing_legend = v
            if base_legend is None:
                base_legend = v
    except Exception:
        pass

if not base_legend:
    Dialogs.alert(
        "No Legend View Found",
        "This model does not have a Legend view. Create one first: View tab > New > Legend, then re-run pyTransmit."
    )
    import sys; sys.exit(0)

# ── STEP 4, COLLECT ELEMENTS FROM TEMP VIEW ──────────────────────────────────

elements_to_copy = []
for el in FilteredElementCollector(doc, temp_view.Id).ToElements():
    try:
        if el.Category:
            elements_to_copy.append(el.Id)
    except Exception:
        pass

if not elements_to_copy:
    forms.alert("Temp drafting view is empty, nothing to copy.", exitscript=True)

# ── STEP 5, COPY TO LEGEND AND DELETE TEMP ───────────────────────────────────

class _CopyUseDestination(DB.IDuplicateTypeNamesHandler):
    def OnDuplicateTypeNamesFound(self, args):
        return DB.DuplicateTypeAction.UseDestinationTypes

with revit.Transaction("pyTransmit - Create Legend") as _t:

    if existing_legend:
        dest_legend = existing_legend
        legend_name = existing_legend.Name
        for cls in (CurveElement, TextNote, ImageInstance, FilledRegion):
            for el in list(FilteredElementCollector(doc, dest_legend.Id).OfClass(cls).ToElements()):
                try: doc.Delete(el.Id)
                except Exception: pass
        output.print_md("Cleared existing legend '{}'".format(legend_name))
    else:
        dest_legend = doc.GetElement(
            base_legend.Duplicate(ViewDuplicateOption.Duplicate)
        )
        legend_name = LEGEND_VIEW_NAME
        try:
            dest_legend.Name = legend_name
        except Exception:
            legend_name = LEGEND_VIEW_NAME + " (Transmittal)"
            try: dest_legend.Name = legend_name
            except Exception: pass

    try: dest_legend.Scale = 1
    except Exception: pass
    try:
        _sp = dest_legend.get_Parameter(DB.BuiltInParameter.VIEW_SCALE_PULLDOWN_METRIC)
        if _sp and not _sp.IsReadOnly:
            _sp.Set(1)
    except Exception: pass

    options = CopyPasteOptions()
    options.SetDuplicateTypeNamesHandler(_CopyUseDestination())
    # Copy elements one at a time, CopyElements does not guarantee the
    # returned ids are in the same order as the input ids, so copying
    # individually keeps each source -> destination override mapping correct.
    for src_id in elements_to_copy:
        try:
            copied = ElementTransformUtils.CopyElements(
                temp_view,
                List[DB.ElementId]([src_id]),
                dest_legend,
                None,
                options
            )
            for dest_id in copied:
                try:
                    dest_legend.SetElementOverrides(dest_id, temp_view.GetElementOverrides(src_id))
                except Exception:
                    pass
        except Exception:
            pass


# Delete temp view in a separate transaction after the copy is fully committed
with revit.Transaction("pyTransmit - Delete Temp View") as _td:
    try:
        doc.Delete(temp_view.Id)
    except Exception as _del_err:
        _td.RollBack()
        output.print_md("Could not delete temp view: {}".format(_del_err))

output.print_md("## Done!")
output.print_md("Legend **'{}'** updated. Find it in the Project Browser under Legends.".format(legend_name))