# -*- coding: utf-8 -*-
"""
pyTable - Export/create_drafting.py

Creates a Drafting View from arbitrary tabular data.
Called by script.py via exec() with PYTABLE_PAYLOAD injected.

Payload keys:
    view_name    : str   - drafting view name in Revit
    fields       : list  - column header strings
    records      : list  - list of lists of cell value strings
    font         : str   - font name
    size_hdr_mm  : float - header font size in mm
    size_dat_mm  : float - data font size in mm
    hdr_tt_id    : DB.ElementId - pre-created header TextNoteType id
    dat_tt_id    : DB.ElementId - pre-created data TextNoteType id
    view_scale   : int   - view scale (default 1)

Pattern:
    Draws table using TextNote elements (one per cell).
    FilledRegion is used for header row background.
    Mirrors pyTransmit's script_create_drafting_view.py approach.
"""

_p = globals().get('PYTABLE_PAYLOAD', {})

from pyrevit import revit, DB, script
from Autodesk.Revit.DB import (
    FilteredElementCollector, ViewFamilyType, ViewFamily,
    ViewDrafting, TextNote, TextNoteOptions,
    HorizontalTextAlignment, FilledRegion, FilledRegionType,
    CurveLoop, Line, XYZ, ElementId,
    CurveElement, ImageInstance,
)

logger = script.get_logger()
doc = revit.doc

MM = 1.0 / 304.8

# ---------------------------------------------------------------------------
# Payload
# ---------------------------------------------------------------------------

view_name   = _p.get('view_name',   'pyTable Drafting View')
fields      = _p.get('fields',      [])
records     = _p.get('records',     [])
font        = _p.get('font',        'Arial')
size_hdr_mm = float(_p.get('size_hdr_mm', 2.5))
size_dat_mm = float(_p.get('size_dat_mm', 2.3))
hdr_tt_id   = _p.get('hdr_tt_id',  ElementId.InvalidElementId)
dat_tt_id   = _p.get('dat_tt_id',  ElementId.InvalidElementId)
view_scale  = int(_p.get('view_scale', 1))
# Legend mode: if set, use this name for the temp view
_temp_name  = _p.get('_legend_temp_view_name', None)
_final_name = _temp_name if _temp_name else view_name

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _get_text_type(tt_id):
    """Return TextNoteType from ID, or first available."""
    if tt_id and tt_id != ElementId.InvalidElementId:
        tt = doc.GetElement(tt_id)
        if tt:
            return tt.Id
    return FilteredElementCollector(doc)\
        .OfClass(DB.TextNoteType)\
        .FirstElementId()


def _draw_text(view, x, y, text, tt_id):
    """Place a TextNote at (x, y)."""
    if not str(text).strip():
        return
    try:
        opts = TextNoteOptions(tt_id)
        opts.HorizontalAlignment = HorizontalTextAlignment.Left
        pt = XYZ(x + 1.5 * MM, y - 1.5 * MM, 0)
        TextNote.Create(doc, view.Id, pt, str(text), opts)
    except Exception as ex:
        logger.debug('TextNote ({},{}) {}'.format(x, y, ex))


def _draw_filled_rect(view, x, y, w, h, frt_id):
    """Draw a solid FilledRegion rectangle."""
    try:
        loop = CurveLoop()
        pts = [
            XYZ(x,     y,     0),
            XYZ(x + w, y,     0),
            XYZ(x + w, y - h, 0),
            XYZ(x,     y - h, 0),
        ]
        for i in range(4):
            seg = Line.CreateBound(pts[i], pts[(i + 1) % 4])
            loop.Append(seg)
        FilledRegion.Create(doc, frt_id, view.Id, [loop])
    except Exception as ex:
        logger.debug('FilledRegion ({},{}) {}'.format(x, y, ex))


def _get_filled_region_type_id():
    frt = FilteredElementCollector(doc)\
        .OfClass(FilledRegionType)\
        .FirstElement()
    return frt.Id if frt else ElementId.InvalidElementId


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

n_cols = len(fields)
n_rows = len(records)

if n_cols == 0:
    logger.error('create_drafting: no columns in data')
    raise Exception('No column data provided to create_drafting.py')

col_w_mm   = max(25.0, 180.0 / n_cols)
col_w      = col_w_mm * MM
row_h_hdr  = max(6.0, size_hdr_mm * 3.0) * MM
row_h_data = max(5.0, size_dat_mm * 2.5) * MM

# ── Get or create drafting view ───────────────────────────────────────────────

view = None
for v in FilteredElementCollector(doc)\
        .OfClass(ViewDrafting)\
        .WhereElementIsNotElementType():
    if v.Name == _final_name:
        view = v
        # Clear existing elements
        for cls in (CurveElement, TextNote, FilledRegion, ImageInstance):
            for el in list(FilteredElementCollector(doc, view.Id)
                           .OfClass(cls).ToElements()):
                try:
                    doc.Delete(el.Id)
                except Exception:
                    pass
        break

if view is None:
    vft = None
    # Try ViewFamily.Drafting first
    for v in FilteredElementCollector(doc)\
            .OfClass(ViewFamilyType)\
            .WhereElementIsNotElementType():
        if v.ViewFamily == ViewFamily.Drafting:
            vft = v
            break
    # Broader fallback — any ViewFamilyType that supports drafting views
    if vft is None:
        for v in FilteredElementCollector(doc)\
                .OfClass(ViewFamilyType)\
                .WhereElementIsNotElementType():
            try:
                if 'draft' in str(v.ViewFamily).lower():
                    vft = v
                    break
            except Exception:
                pass
    if vft is None:
        raise Exception(
            'No Drafting View family type found. '
            'Create any Drafting View manually first, then re-run pyTable.'
        )
    view = ViewDrafting.Create(doc, vft.Id)
    view.Name = _final_name

try:
    view.Scale = view_scale
except Exception:
    pass

# ── Draw table ────────────────────────────────────────────────────────────────

hdr_txt_id = _get_text_type(hdr_tt_id)
dat_txt_id = _get_text_type(dat_tt_id)
frt_id     = _get_filled_region_type_id()

# Header row background
if frt_id != ElementId.InvalidElementId:
    for ci in range(n_cols):
        x = ci * col_w
        _draw_filled_rect(view, x, 0.0, col_w, row_h_hdr, frt_id)

# Header text
for ci, field in enumerate(fields):
    x = ci * col_w
    _draw_text(view, x, 0.0, field, hdr_txt_id)

# Data rows
for ri, record in enumerate(records):
    y = -(row_h_hdr + ri * row_h_data)
    for ci, cell in enumerate(record):
        if ci >= n_cols:
            break
        x = ci * col_w
        _draw_text(view, x, y, cell, dat_txt_id)

logger.debug(
    'create_drafting: "{}" {}r x {}c'.format(
        _final_name, 1 + n_rows, n_cols
    )
)
