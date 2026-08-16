# -*- coding: utf-8 -*-
# script_create_legend_studio.py
#
# Legend writer for pyTransmit STUDIO layouts.
#
# Same two-step as script_create_legend.py, because a Legend view cannot be
# drawn into directly the way a Drafting View can: draw the transmittal into a
# temporary drafting view, copy every element across to the Legend, then throw
# the temporary view away. The only difference from the Layout Builder version
# is which drafting-view writer does the drawing.
#
# Payload keys used: whatever script_create_drafting_view_studio.py needs; this
# script only adds _legend_temp_view_name.

_p = globals().get('PYTRANSMIT_PAYLOAD', {})

import os

from pyrevit.framework import List
from pyrevit import revit, script, DB, forms

from Autodesk.Revit.DB import (
    FilteredElementCollector, CurveElement, TextNote,
    ImageInstance, FilledRegion,
    ElementTransformUtils, CopyPasteOptions,
    ViewType, ViewDuplicateOption, ViewDrafting,
)

try:
    from Snippets import _dialogs as sdlg
except Exception:
    sdlg = None


def _alert(message, title='', exitscript=False):
    """Themed popup via the shared Snippets dialog lib, falls back to
    pyRevit's default forms.alert if the shared lib isn't available."""
    if sdlg:
        sdlg.message(message, title=title)
    else:
        forms.alert(message, title=title)
    if exitscript:
        script.exit()


_log_lines = _p.get('_log_lines', [])


def _log(msg):
    try:
        _log_lines.append(str(msg))
    except Exception:
        pass


doc = revit.doc
try:
    script.get_output().hide()
except Exception:
    pass

TITLE = 'pyTransmit Studio - Legend'
TEMP_VIEW_NAME = 'pyTransmit TEMP'
LEGEND_VIEW_NAME = 'pyTransmit Document'

# ── Step 1: draw into a temporary drafting view ──────────────────────────────
_script_dir = os.path.dirname(os.path.abspath(__file__))
_drafting_path = os.path.join(_script_dir, 'script_create_drafting_view_studio.py')

if not os.path.isfile(_drafting_path):
    _alert('script_create_drafting_view_studio.py not found at:\n{}'.format(
        _drafting_path), title=TITLE, exitscript=True)

_payload_for_drafting = dict(_p)
_payload_for_drafting['_legend_temp_view_name'] = TEMP_VIEW_NAME

_ns = {
    '__name__': 'drafting_view_studio_for_legend',
    '__file__': _drafting_path,
    '__builtins__': __builtins__,
    'PYTRANSMIT_PAYLOAD': _payload_for_drafting,
}
with open(_drafting_path, 'r') as _f:
    _src = _f.read()
try:
    exec(_src, _ns)
except SystemExit:
    # The drafting writer bailed out and has already said why.
    raise
except Exception as _e:
    import traceback as _tb
    _alert('Error running the Studio drafting view script:\n{}'.format(
        _tb.format_exc() or str(_e)), title=TITLE, exitscript=True)

# ── Step 2: find the temporary view ──────────────────────────────────────────
temp_view = None
for _v in revit.query.get_elements_by_class(ViewDrafting, doc=doc):
    try:
        if _v.IsValidObject and _v.Name == TEMP_VIEW_NAME:
            temp_view = _v
            break
    except Exception:
        pass

if not temp_view:
    _alert('Temp drafting view "{}" not found after generation.'.format(
        TEMP_VIEW_NAME), title=TITLE, exitscript=True)

# ── Step 3: find or create the legend ────────────────────────────────────────
existing_legend = None
base_legend = None
for _v in revit.query.get_elements_by_class(DB.View, doc=doc):
    try:
        if _v.ViewType == ViewType.Legend and not _v.IsTemplate:
            if _v.Name in (LEGEND_VIEW_NAME, LEGEND_VIEW_NAME + ' (Transmittal)'):
                existing_legend = _v
            if base_legend is None:
                base_legend = _v
    except Exception:
        pass

if not base_legend:
    # A Legend can only be made by duplicating one, and the API cannot create
    # the first one - so this is a genuine dead end, not a bug.
    _alert('This model does not have a Legend view. Create one first: '
           'View tab > New > Legend, then re-run pyTransmit.',
           title='No Legend View Found')
    _log('Legend export skipped - the model has no Legend view to copy.')
    import sys
    sys.exit(0)

# ── Step 4: collect what was drawn ───────────────────────────────────────────
elements_to_copy = []
for _el in FilteredElementCollector(doc, temp_view.Id).ToElements():
    try:
        if _el.Category:
            elements_to_copy.append(_el.Id)
    except Exception:
        pass

if not elements_to_copy:
    _alert('The temp drafting view is empty, nothing to copy.',
           title=TITLE, exitscript=True)


# ── Step 5: copy into the legend, then drop the temp view ────────────────────
class _CopyUseDestination(DB.IDuplicateTypeNamesHandler):
    def OnDuplicateTypeNamesFound(self, args):
        return DB.DuplicateTypeAction.UseDestinationTypes


with revit.Transaction('pyTransmit Studio - Create Legend') as _t:

    if existing_legend:
        dest_legend = existing_legend
        legend_name = existing_legend.Name
        for _cls in (CurveElement, TextNote, ImageInstance, FilledRegion):
            for _el in list(FilteredElementCollector(
                    doc, dest_legend.Id).OfClass(_cls).ToElements()):
                try:
                    doc.Delete(_el.Id)
                except Exception:
                    pass
        _log('Cleared existing legend "{}"'.format(legend_name))
    else:
        dest_legend = doc.GetElement(
            base_legend.Duplicate(ViewDuplicateOption.Duplicate))
        legend_name = LEGEND_VIEW_NAME
        try:
            dest_legend.Name = legend_name
        except Exception:
            legend_name = LEGEND_VIEW_NAME + ' (Transmittal)'
            try:
                dest_legend.Name = legend_name
            except Exception:
                pass

    try:
        dest_legend.Scale = 1
    except Exception:
        pass
    try:
        _sp = dest_legend.get_Parameter(DB.BuiltInParameter.VIEW_SCALE_PULLDOWN_METRIC)
        if _sp and not _sp.IsReadOnly:
            _sp.Set(1)
    except Exception:
        pass

    _options = CopyPasteOptions()
    _options.SetDuplicateTypeNamesHandler(_CopyUseDestination())
    # Copied one at a time: CopyElements does not guarantee the returned ids
    # are in the same order as the input ids, so copying individually keeps
    # each source -> destination override mapping correct.
    for _src_id in elements_to_copy:
        try:
            _copied = ElementTransformUtils.CopyElements(
                temp_view, List[DB.ElementId]([_src_id]), dest_legend, None, _options)
            for _dest_id in _copied:
                try:
                    dest_legend.SetElementOverrides(
                        _dest_id, temp_view.GetElementOverrides(_src_id))
                except Exception:
                    pass
        except Exception:
            pass

# Deleted in its own transaction, after the copy is fully committed.
with revit.Transaction('pyTransmit Studio - Delete Temp View') as _td:
    try:
        doc.Delete(temp_view.Id)
    except Exception as _del_err:
        _td.RollBack()
        _log('Could not delete the temp view: {}'.format(_del_err))

_log('Legend "{}" updated: {} element(s) copied.'.format(
    legend_name, len(elements_to_copy)))
