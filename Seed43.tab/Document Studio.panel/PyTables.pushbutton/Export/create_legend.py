# -*- coding: utf-8 -*-
"""
pyTable - Export/create_legend.py

Creates a Legend View from arbitrary tabular data.
Mirrors pyTransmit's script_create_legend.py exactly:
  1. Run create_drafting.py to build a temp Drafting View
  2. Copy all elements from the temp view into a Legend View
  3. Delete the temp Drafting View

Called by script.py via exec() with PYTABLE_PAYLOAD injected.
"""

_p = globals().get('PYTABLE_PAYLOAD', {})

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

view_name  = _p.get('view_name', 'pyTable Legend')
TEMP_NAME  = '__pyTable_TEMP__'

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

_ns = {
    '__name__':         'drafting_for_legend',
    '__file__':         _drafting_path,
    '__builtins__':     __builtins__,
    'PYTABLE_PAYLOAD':  _payload_for_drafting,
}
with open(_drafting_path, 'r') as _f:
    _src = _f.read()

# create_drafting.py modifies the document so needs a transaction
with revit.Transaction('pyTable - Temp drafting for legend'):
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
        'then run pyTable again.',
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

# ---------------------------------------------------------------------------
# Step 5 — Copy into legend, delete temp
# ---------------------------------------------------------------------------

class _UseDestination(DB.IDuplicateTypeNamesHandler):
    def OnDuplicateTypeNamesFound(self, args):
        return DB.DuplicateTypeAction.UseDestinationTypes

with revit.Transaction('pyTable - Create Legend: {}'.format(view_name)):

    if existing_legend:
        dest = existing_legend
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
            dest.Name = view_name + ' (pyTable)'

    try:
        dest.Scale = int(_p.get('view_scale', 1))
    except Exception:
        pass

    opts = CopyPasteOptions()
    opts.SetDuplicateTypeNamesHandler(_UseDestination())

    ElementTransformUtils.CopyElements(
        temp_view,
        List[DB.ElementId](elements_to_copy),
        dest,
        None,
        opts
    )

# Delete temp in its own transaction (same as pyTransmit)
with revit.Transaction('pyTable - Delete temp view'):
    try:
        doc.Delete(temp_view.Id)
    except Exception as ex:
        logger.debug('Could not delete temp view: {}'.format(ex))

logger.debug(
    'create_legend: "{}" complete'.format(view_name)
)
