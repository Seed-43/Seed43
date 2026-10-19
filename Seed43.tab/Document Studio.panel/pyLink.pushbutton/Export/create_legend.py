# -*- coding: utf-8 -*-
"""
pyLink - Export/create_legend.py

Creates a Legend View from arbitrary tabular data.
Mirrors pyTransmit's script_create_legend.py exactly:
  1. Run create_drafting.py to build a temp Drafting View
  2. Copy all elements from the temp view into a Legend View
  3. Delete the temp Drafting View

Called by script.py via exec() with PYLINK_PAYLOAD injected.
"""

_p = globals().get('PYLINK_PAYLOAD', {})

from pyrevit.framework import List
from pyrevit import revit, DB, script, forms
from Autodesk.Revit.DB import (
    FilteredElementCollector, CurveElement, TextNote,
    ImageInstance, FilledRegion, ViewDrafting,
    ElementTransformUtils, CopyPasteOptions,
    ViewType, ViewDuplicateOption, ElementId,
)

import os

logger = script.get_logger()
doc = revit.doc

# ---------------------------------------------------------------------------
# Payload
# ---------------------------------------------------------------------------

view_name  = _p.get('view_name', 'pyLink Legend')
TEMP_NAME  = '__pyLink_TEMP__'
# Position in a run of rows sharing this legend. 0 clears and lands at
# the top; higher numbers are copied in below what is already there.
stack_index = int(_p.get('stack_index', 0) or 0)
STACK_GAP_FT = 8.0 / 304.8

# ---------------------------------------------------------------------------
# Step 1 — Run create_drafting.py to build the temp drafting view
# ---------------------------------------------------------------------------

_script_dir   = os.path.dirname(os.path.abspath(__file__))
_drafting_path = os.path.join(_script_dir, 'create_drafting.py')

if not os.path.exists(_drafting_path):
    raise Exception(
        'create_drafting.py not found at: {}'.format(_drafting_path)
    )

_payload_for_drafting = dict(_p)
_payload_for_drafting['_legend_temp_view_name'] = TEMP_NAME
# The temp view is created fresh every time, so it always draws at the
# origin. Stacking happens when the elements are copied into the legend
# below, where there is something to stack under.
_payload_for_drafting['stack_index'] = 0

_ns = {
    '__name__':         'drafting_for_legend',
    '__file__':         _drafting_path,
    '__builtins__':     __builtins__,
    'PYLINK_PAYLOAD':  _payload_for_drafting,
}
with open(_drafting_path, 'r') as _f:
    _src = _f.read()

# create_drafting.py modifies the document so needs a transaction
with revit.Transaction('pyLink - Temp drafting for legend'):
    exec(_src, _ns)

# ---------------------------------------------------------------------------
# Step 2 — Find the temp drafting view
# ---------------------------------------------------------------------------

temp_view = None
for v in FilteredElementCollector(doc)\
        .OfClass(ViewDrafting)\
        .WhereElementIsNotElementType():
    if v.Name == TEMP_NAME:
        temp_view = v
        break

if not temp_view:
    raise Exception(
        'Temp drafting view "{}" not found after generation'.format(TEMP_NAME)
    )

# ---------------------------------------------------------------------------
# Step 3 — Find or create the legend view
# ---------------------------------------------------------------------------

existing_legend = None
base_legend     = None

for v in FilteredElementCollector(doc)\
        .OfClass(DB.View)\
        .WhereElementIsNotElementType():
    try:
        if v.ViewType == ViewType.Legend and not v.IsTemplate:
            if v.Name == view_name:
                existing_legend = v
            if base_legend is None:
                base_legend = v
    except Exception:
        pass

if not base_legend:
    forms.alert(
        'No Legend view found in this project. '
        'Create one first via View tab > New > Legend, '
        'then run pyLink again.',
        exitscript=True
    )

# ---------------------------------------------------------------------------
# Step 4 — Collect elements from temp view
# ---------------------------------------------------------------------------

elements_to_copy = []
for el in FilteredElementCollector(doc, temp_view.Id).ToElements():
    try:
        if el.Category:
            elements_to_copy.append(el.Id)
    except Exception:
        pass

if not elements_to_copy:
    raise Exception('Temp drafting view is empty — nothing to copy')

# Record each element's view-specific override on the temp view BEFORE
# copying — CopyElements does not carry these over on its own (an
# override is stored against the (view, elementId) pair, not the
# element), so without this every colour create_drafting.py just
# applied would be silently lost the moment these land in the legend.
overrides_by_old_id = {}
for el_id in elements_to_copy:
    try:
        ogs = temp_view.GetElementOverrides(el_id)
        if ogs is not None:
            overrides_by_old_id[el_id] = ogs
    except Exception as ex:
        logger.debug('Read override {}: {}'.format(el_id, ex))

# ---------------------------------------------------------------------------
# Step 5 — Copy into legend, delete temp
# ---------------------------------------------------------------------------

class _UseDestination(DB.IDuplicateTypeNamesHandler):
    def OnDuplicateTypeNamesFound(self, args):
        return DB.DuplicateTypeAction.UseDestinationTypes

with revit.Transaction('pyLink - Create Legend: {}'.format(view_name)):

    _copy_dy_ft = 0.0
    if existing_legend:
        dest = existing_legend
        if stack_index > 0:
            # Sharing this legend with a table already on it: keep that
            # and drop this one in underneath. Only the first row of a
            # run clears, so re-applying the run replaces it instead of
            # stacking another copy.
            _lowest = None
            for el in FilteredElementCollector(doc, dest.Id).ToElements():
                try:
                    bb = el.get_BoundingBox(dest)
                except Exception:
                    bb = None
                if bb is None:
                    continue
                if _lowest is None or bb.Min.Y < _lowest:
                    _lowest = bb.Min.Y
            if _lowest is not None:
                _copy_dy_ft = _lowest - STACK_GAP_FT
            logger.debug('stacking legend table {} at y={:.1f}mm'.format(
                stack_index, _copy_dy_ft * 304.8))
        else:
            # Clear existing content
            for cls in (CurveElement, TextNote, ImageInstance, FilledRegion):
                for el in list(
                    FilteredElementCollector(doc, dest.Id)
                    .OfClass(cls).ToElements()
                ):
                    try:
                        doc.Delete(el.Id)
                    except Exception:
                        pass
    else:
        dest = doc.GetElement(
            base_legend.Duplicate(ViewDuplicateOption.Duplicate)
        )
        try:
            dest.Name = view_name
        except Exception:
            dest.Name = view_name + ' (pyLink)'

    try:
        dest.Scale = int(_p.get('view_scale', 1))
    except Exception:
        pass

    opts = CopyPasteOptions()
    opts.SetDuplicateTypeNamesHandler(_UseDestination())

    # The temp view drew at the origin, so a stacked table is shifted
    # down on the way in rather than being redrawn somewhere else.
    _xform = None
    if _copy_dy_ft:
        _xform = DB.Transform.CreateTranslation(DB.XYZ(0, _copy_dy_ft, 0))

    new_ids = ElementTransformUtils.CopyElements(
        temp_view,
        List[DB.ElementId](elements_to_copy),
        dest,
        _xform,
        opts
    )

    # Re-apply each element's recorded override onto its new copy in
    # the legend view. CopyElements returns the new ids in the same
    # order as the input collection, so old/new pair up by position.
    new_ids_list = list(new_ids)
    for old_id, new_id in zip(elements_to_copy, new_ids_list):
        ogs = overrides_by_old_id.get(old_id)
        if ogs is not None:
            try:
                dest.SetElementOverrides(new_id, ogs)
            except Exception as ex:
                logger.debug(
                    'Reapply override {} -> {}: {}'.format(
                        old_id, new_id, ex
                    )
                )

# Delete temp in its own transaction (same as pyTransmit)
with revit.Transaction('pyLink - Delete temp view'):
    try:
        doc.Delete(temp_view.Id)
    except Exception as ex:
        logger.debug('Could not delete temp view: {}'.format(ex))

logger.debug(
    'create_legend: "{}" complete'.format(view_name)
)

try:
    _view_id_int = dest.Id.IntegerValue
except AttributeError:
    _view_id_int = int(dest.Id.Value)
PYLINK_RESULT = {'view_id': _view_id_int}
