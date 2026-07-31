# -*- coding: utf-8 -*-
"""
pyLink - Export/create_drafting.py

Creates a Drafting View from arbitrary tabular data.
Called by script.py via exec() with PYLINK_PAYLOAD injected.

Payload keys:
    view_name    : str   - drafting view name in Revit
    fields       : list  - column header strings
    records      : list  - list of lists of cell value strings
    font         : str   - font name
    size_hdr_mm  : float - header font size in mm
    size_dat_mm  : float - data font size in mm
    hdr_tt_id    : DB.ElementId - header TextNoteType id (plain, black)
    dat_tt_id    : DB.ElementId - data TextNoteType id (plain, black)
    view_scale   : int   - view scale (default 1)
    cell_styles  : dict  - {(row,col): style_dict} per-cell Excel formatting
                           (row 0 = header row, relative to the named range)
    merges       : list  - [(r1,c1,r2,c2), ...] merged cell ranges, inclusive
    row_heights  : dict  - {row_idx: pts} Excel row heights, in points
    col_widths   : dict  - {col_idx: mm} Excel column widths, in mm
    default_row_height : float - fallback row height in points
    fill_type_id : DB.ElementId - single reusable solid-fill FilledRegionType

Pattern:
    Draws a flat 2D copy of the Excel table using one reusable
    FilledRegionType and the default Revit <Line> line style for
    everything — per-cell colour (fill, text, border) is applied
    afterwards as a view-specific graphic override on that individual
    element instance (the same as manually right-clicking an element >
    Override Graphics in View > By Element), rather than creating a
    new Type or line-style subcategory per colour. Only sides/fills
    Excel actually set get drawn; nothing is invented. Cell geometry
    uses Excel's own row heights/column widths and respects merged
    cells. Mirrors pyTransmit's script_create_drafting_view.py
    approach, extended with real per-cell formatting.
"""

_p = globals().get('PYLINK_PAYLOAD', {})

from pyrevit import revit, DB, script
from Autodesk.Revit.DB import (
    FilteredElementCollector, ViewFamilyType, ViewFamily,
    ViewDrafting, TextNote, TextNoteOptions,
    HorizontalTextAlignment, FilledRegion, FilledRegionType,
    CurveLoop, Line, XYZ, ElementId, Color,
    CurveElement, ImageInstance, OverrideGraphicSettings,
    BuiltInCategory, GraphicsStyleType, ElementTransformUtils,
    BuiltInParameter,
)
import math

logger = script.get_logger()
doc = revit.doc

MM    = 1.0 / 304.8   # feet per mm
PT_MM = 0.352778      # mm per point (1pt = 1/72 in = 25.4/72 mm)

# ---------------------------------------------------------------------------
# Payload
# ---------------------------------------------------------------------------

view_name     = _p.get('view_name',   'pyLink Drafting View')
fields        = _p.get('fields',      [])
records       = _p.get('records',     [])
font          = _p.get('font',        'Arial')
size_hdr_mm   = float(_p.get('size_hdr_mm', 2.5))
size_dat_mm   = float(_p.get('size_dat_mm', 2.3))
hdr_tt_id     = _p.get('hdr_tt_id',  ElementId.InvalidElementId)
dat_tt_id     = _p.get('dat_tt_id',  ElementId.InvalidElementId)
view_scale    = int(_p.get('view_scale', 1))
cell_styles   = _p.get('cell_styles',   {})
merges        = _p.get('merges',        [])
row_heights   = _p.get('row_heights',   {})
col_widths    = _p.get('col_widths',    {})
default_row_height = float(_p.get('default_row_height', 14.0))
fill_type_id  = _p.get('fill_type_id', ElementId.InvalidElementId)
# Legend mode: if set, use this name for the temp view
_temp_name  = _p.get('_legend_temp_view_name', None)
_final_name = _temp_name if _temp_name else view_name

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _get_text_type(tt_id):
    """Return TextNoteType Id from ID, or first available."""
    if tt_id and tt_id != ElementId.InvalidElementId:
        tt = doc.GetElement(tt_id)
        if tt:
            return tt.Id
    return FilteredElementCollector(doc)\
        .OfClass(DB.TextNoteType)\
        .FirstElementId()


def _safe_name(el):
    """Read an element's name safely across element types. Raw .Name
    works fine for GraphicsStyle (it was never affected by the Revit
    2023+ break that hit ElementType.Name specifically - that only
    applies to Type-derived classes like FilledRegionType/
    TextNoteType), so try it first; fall back to the
    BuiltInParameter/Element.Name.GetValue routes only for classes
    where raw access actually throws."""
    try:
        return el.Name
    except Exception:
        pass
    try:
        p = el.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if p:
            return p.AsString()
    except Exception:
        pass
    try:
        return DB.Element.Name.GetValue(el)
    except Exception:
        return None


def _default_line_style_id():
    """The project's own default '<Lines>' line style — used for every
    border instead of creating a custom Lines subcategory per colour.
    Colour still comes from a per-instance view override layered on
    top of this (see _override_color), specifically so a project's own
    custom appearance for <Lines> — weight, pattern, whatever a user
    has set it to — is left alone and only the colour gets overridden,
    never swapped for a different style."""
    lines_cat = None
    try:
        lines_cat = doc.Settings.Categories.get_Item(
            BuiltInCategory.OST_Lines
        )
    except Exception as ex:
        logger.debug('OST_Lines category lookup: {}'.format(ex))

    # Primary: the project's own Lines category subcategories - this
    # is the authoritative, unambiguous source. A document-wide flat
    # scan can also turn up same-named '<Lines>' entries belonging to
    # linked models, which would be the wrong element to use here.
    if lines_cat is not None:
        try:
            for sub in lines_cat.SubCategories:
                try:
                    if _safe_name(sub) == '<Lines>':
                        gs = sub.GetGraphicsStyle(GraphicsStyleType.Projection)
                        if gs:
                            return gs.Id
                except Exception:
                    continue
        except Exception as ex:
            logger.debug('Default <Lines> subcategory search: {}'.format(ex))

    # Fallback 1: flat scan of every GraphicsStyle in the document, in
    # case the subcategory walk missed it for some reason.
    try:
        for gs in FilteredElementCollector(doc).OfClass(DB.GraphicsStyle):
            try:
                if _safe_name(gs) == '<Lines>':
                    return gs.Id
            except Exception:
                continue
    except Exception as ex:
        logger.debug('Default <Lines> flat search: {}'.format(ex))

    # Fallback 2: the Lines category's own top-level default style —
    # last resort only. This is what was producing plain "Lines"
    # instead of "<Lines>" before, so it's kept purely so borders at
    # least land on *some* real style rather than none.
    if lines_cat is not None:
        try:
            gs = lines_cat.GetGraphicsStyle(GraphicsStyleType.Projection)
            if gs:
                logger.warning(
                    'Could not find the exact "<Lines>" style - falling '
                    'back to the Lines category default instead.'
                )
                return gs.Id
        except Exception as ex:
            logger.debug('Default <Lines> style via category: {}'.format(ex))

    logger.error(
        'No usable line style found at all for borders - they will '
        'be drawn on whatever Revit assigns by default.'
    )
    return ElementId.InvalidElementId


# Excel's border style keywords -> Revit's 1-16 line weight scale.
# There's no exact correspondence between the two systems (Excel has
# no numeric weight, just named styles), so this is a best-effort
# mapping, not a precise conversion - picked to roughly preserve the
# thin/medium/thick relationships Excel itself distinguishes.
_EXCEL_WEIGHT_MAP = {
    'hair':              1,
    'thin':              1,
    'dotted':            1,
    'dashed':            1,
    'dashDot':           1,
    'dashDotDot':        1,
    'slantDashDot':      2,
    'medium':            3,
    'mediumDashed':      3,
    'mediumDashDot':     3,
    'mediumDashDotDot':  3,
    'double':            3,
    'thick':             5,
}


def _excel_border_weight(style_val):
    """Map an Excel border style keyword to a Revit line weight (1-16).
    Falls back to 1 (thin) for anything unrecognised."""
    return _EXCEL_WEIGHT_MAP.get(style_val, 1)


def _override_color(element_id, rgb, weight=None):
    """Apply a view-specific 'Override Graphics in View > By Element'
    projection line colour (and optionally weight) AND surface fill
    colour to one element - covers TextNote, DetailLine and
    FilledRegion alike."""
    if not rgb:
        return
    try:
        ogs = OverrideGraphicSettings()
        color = Color(rgb[0], rgb[1], rgb[2])
        ogs.SetProjectionLineColor(color)
        if weight:
            try:
                ogs.SetProjectionLineWeight(weight)
            except Exception as ex:
                logger.debug('Set line weight {}: {}'.format(weight, ex))
        try:
            ogs.SetSurfaceForegroundPatternColor(color)
            ogs.SetSurfaceForegroundPatternVisible(True)
        except Exception:
            pass  # not every element type supports surface pattern overrides
        view.SetElementOverrides(element_id, ogs)
    except Exception as ex:
        logger.debug('Override colour on {}: {}'.format(element_id, ex))


def _excel_rotation_to_radians(excel_rot):
    """
    Map Excel's alignment.textRotation convention to a Revit rotation
    angle in radians, for use with ElementTransformUtils.RotateElement.

    Excel: 0 = horizontal, 1-90 = counter-clockwise from horizontal
    (reading bottom-to-top), 91-180 = clockwise (stored as 90+angle,
    reading top-to-bottom), 255 = stacked/vertical (not supported here).

    Revit's ElementTransformUtils.RotateElement angle is positive =
    counter-clockwise about the given axis (standard math convention
    when viewed from the +Z side, i.e. looking straight down at the
    view in plan). Excel's 90 (bottom-to-top) is a +90 degree turn in
    that same sense, so it maps straight across; the 91-180 range
    (top-to-bottom) maps to the mirrored angle in the other direction.
    """
    if not excel_rot:
        return 0.0
    if 1 <= excel_rot <= 90:
        return math.radians(excel_rot)
    if 91 <= excel_rot <= 180:
        return -math.radians(excel_rot - 90)
    return 0.0


_TEXT_MARGIN = 1.5 * MM


def _force_top_anchor(tn):
    """
    Force this TextNote's vertical anchor to Top - i.e. the insertion
    point represents the box's TOP edge, staying put as text height
    changes. Every position calculation in _draw_text assumes this
    (origin = box's top-left corner); previously this was only ever
    assumed from an unverified default, never actually set. Try the
    strongly-typed property first, fall back to the raw parameter -
    either can be missing depending on Revit API version.
    """
    try:
        from Autodesk.Revit.DB import VerticalAlignment
        tn.VerticalAlignment = VerticalAlignment.Top
        return
    except Exception as ex:
        logger.debug('VerticalAlignment property set failed: {}'.format(ex))
    try:
        p = tn.get_Parameter(BuiltInParameter.TEXT_ALIGN_VERT)
        if p and not p.IsReadOnly:
            p.Set(0)  # 0 = Top in the TEXT_ALIGN_VERT built-in enum
    except Exception as ex:
        logger.debug('TEXT_ALIGN_VERT parameter set failed: {}'.format(ex))


def _draw_text(view, x, y, w, h, text, tt_id, color_rgb,
                rotation_rad=0.0, halign='Left', valign='Bottom',
                wrap=False):
    """
    Place a TextNote so its (possibly rotated) box lands inside the
    cell rectangle (x, y, w, h) — (x, y) is the cell's top-left corner
    — positioned per Excel's own horizontal/vertical alignment, with
    optional wrapping.

    General approach:
      1. Create the note UNROTATED at a scratch point, with Width set
         to the wrap constraint (only if Excel's wrap_text was on) so
         Revit computes real line breaks off actual font metrics.
      2. Read back its real (unrotated) box size from its bounding box
         — wrapping can produce more/fewer lines than guessed.
      3. Excel keeps halign/valign meaning literal screen X/Y even when
         rotated (halign -> horizontal position in the cell, valign ->
         vertical), so work out where the UNROTATED box's origin needs
         to sit such that, after rotating 90 around that origin, the
         rotated footprint lands where halign/valign says it should.
      4. Discard that probe note, recreate the real one at the correct
         origin, then rotate around that same point.

    rotation_rad: 0, or the drafting rotation for Excel textRotation
    (see _excel_rotation_to_radians) - positive (CCW, Excel 1-90,
    "bottom-to-top") or negative (CW, Excel 91-180, "top-to-bottom").
    """
    if not str(text).strip():
        return
    try:
        # Shrink the working rect by the margin on every side so text
        # never touches the border lines.
        mx, my  = x + _TEXT_MARGIN, y - _TEXT_MARGIN
        mw, mh  = max(w - 2 * _TEXT_MARGIN, 0), max(h - 2 * _TEXT_MARGIN, 0)

        opts = TextNoteOptions(tt_id)
        # Always Left - our own ox/oy math below is solely responsible
        # for halign/valign positioning. If Revit's own justification
        # is also set to Center/Right, it shifts what the insertion
        # point means relative to the box (left edge vs. center vs.
        # right edge), which double-applies alignment on top of ours
        # and throws the box off by about half its size.
        opts.HorizontalAlignment = HorizontalTextAlignment.Left

        # ── Pass 1: create unrotated at a scratch point purely to
        # measure the real wrapped box size, then discard it. Safer
        # than creating once and relocating via tn.Coord, whose setter
        # reliability isn't confirmed - recreating at the right spot
        # from the start avoids depending on it.
        scratch = XYZ(0, 0, 0)
        wrap_width = mh if rotation_rad else mw
        if wrap and wrap_width > 0:
            probe = TextNote.Create(doc, view.Id, scratch, wrap_width, str(text), opts)
        else:
            probe = TextNote.Create(doc, view.Id, scratch, str(text), opts)
        _force_top_anchor(probe)
        # A TextNote's bounding box is not reliable until the document
        # regenerates after creation - without this, get_BoundingBox()
        # can return None (silently dropping the text below, since
        # that's caught by this function's own try/except).
        doc.Regenerate()
        bbox = probe.get_BoundingBox(view)
        if bbox is None:
            # Extremely defensive fallback - should not happen after
            # Regenerate(), but better a plain top-left-anchored note
            # than no text at all.
            logger.warning(
                'TextNote bbox still None after Regenerate for "{}" - '
                'falling back to plain top-left placement'.format(text)
            )
            doc.Delete(probe.Id)
            fallback_pt = XYZ(mx, my, 0)
            if wrap and wrap_width > 0:
                tn = TextNote.Create(doc, view.Id, fallback_pt, wrap_width, str(text), opts)
            else:
                tn = TextNote.Create(doc, view.Id, fallback_pt, str(text), opts)
            _force_top_anchor(tn)
            if color_rgb:
                _override_color(tn.Id, color_rgb)
            if rotation_rad:
                try:
                    axis = Line.CreateBound(fallback_pt, fallback_pt + XYZ.BasisZ)
                    ElementTransformUtils.RotateElement(doc, tn.Id, axis, rotation_rad)
                except Exception as rex:
                    logger.debug('TextNote rotate ({},{}) {}'.format(x, y, rex))
            return
        box_w = bbox.Max.X - bbox.Min.X
        box_h = bbox.Max.Y - bbox.Min.Y
        doc.Delete(probe.Id)

        # ── Work out the pre-rotation origin (top-left, as authored) ─
        # so that after rotating, the box lands per halign/valign.
        # See module docstring math: unrotated footprint is
        # X:[ox, ox+box_w], Y:[oy-box_h, oy] before rotation.
        if rotation_rad > 0:
            # +90 CCW: post-rotation footprint becomes
            # X:[ox, ox+box_h] (halign axis), Y:[oy, oy+box_w] (valign axis)
            ox = {
                'Left':   mx,
                'Center': mx + (mw - box_h) / 2.0,
                'Right':  mx + mw - box_h,
            }.get(halign, mx)
            oy = {
                'Top':    my - box_w,
                'Center': my - mh / 2.0 - box_w / 2.0,
                'Bottom': my - mh,
            }.get(valign, my - mh)
        elif rotation_rad < 0:
            # -90 CW: post-rotation footprint becomes
            # X:[ox-box_h, ox] (halign axis), Y:[oy-box_w, oy] (valign axis)
            ox = {
                'Left':   mx + box_h,
                'Center': mx + (mw + box_h) / 2.0,
                'Right':  mx + mw,
            }.get(halign, mx + box_h)
            oy = {
                'Top':    my,
                'Center': my - mh / 2.0 + box_w / 2.0,
                'Bottom': my - mh + box_w,
            }.get(valign, my)
        else:
            # No rotation: standard top-left-anchored box.
            ox = {
                'Left':   mx,
                'Center': mx + (mw - box_w) / 2.0,
                'Right':  mx + mw - box_w,
            }.get(halign, mx)
            oy = {
                'Top':    my,
                'Center': my - (mh - box_h) / 2.0,
                'Bottom': my - mh + box_h,
            }.get(valign, my)

        pt = XYZ(ox, oy, 0)
        if wrap and wrap_width > 0:
            tn = TextNote.Create(doc, view.Id, pt, wrap_width, str(text), opts)
        else:
            tn = TextNote.Create(doc, view.Id, pt, str(text), opts)
        _force_top_anchor(tn)

        if color_rgb:
            _override_color(tn.Id, color_rgb)

        if rotation_rad:
            try:
                axis = Line.CreateBound(pt, pt + XYZ.BasisZ)
                ElementTransformUtils.RotateElement(doc, tn.Id, axis, rotation_rad)
            except Exception as rex:
                logger.debug('TextNote rotate ({},{}) {}'.format(x, y, rex))
    except Exception as ex:
        logger.debug('TextNote ({},{}) {}'.format(x, y, ex))


def _draw_filled_rect(view, x, y, w, h, frt_id, color_rgb):
    """Draw a solid FilledRegion rectangle, tinted via view override.
    (x, y) is the top-left corner."""
    if not frt_id or frt_id == ElementId.InvalidElementId or not color_rgb:
        return
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
        fr = FilledRegion.Create(doc, frt_id, view.Id, [loop])
        _override_color(fr.Id, color_rgb)
    except Exception as ex:
        logger.debug('FilledRegion ({},{}) {}'.format(x, y, ex))


def _draw_border_line(view, x1, y1, x2, y2, line_style_id, color_rgb, weight=None):
    """Draw a single DetailLine on the default <Lines> style, tinted via
    a per-instance view override to the border's actual colour and
    (best-effort) weight."""
    try:
        seg = Line.CreateBound(XYZ(x1, y1, 0), XYZ(x2, y2, 0))
        dc = doc.Create.NewDetailCurve(view, seg)
        if line_style_id and line_style_id != ElementId.InvalidElementId:
            try:
                dc.LineStyle = doc.GetElement(line_style_id)
            except Exception as ex:
                logger.warning('Set <Lines> style: {}'.format(ex))
        _override_color(dc.Id, color_rgb or (0, 0, 0), weight)
    except Exception as ex:
        logger.debug(
            'Border ({},{})-({},{}) {}'.format(x1, y1, x2, y2, ex)
        )


def _cell_border_rgb(style, side):
    """Return the (r,g,b) for one border side, or None if Excel didn't
    set a border on that side of this cell."""
    val = style.get('border_' + side)
    if not val or val in ('', 'none'):
        return None
    rgb = style.get('border_' + side + '_color')
    return tuple(rgb) if rgb else (0, 0, 0)


def _cell_border_weight(style, side):
    """Revit line weight (1-16) for one border side, mapped from
    Excel's own border style keyword for that side."""
    val = style.get('border_' + side)
    if not val or val in ('', 'none'):
        return None
    return _excel_border_weight(val)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

n_cols = len(fields)
n_rows = len(records)

if n_cols == 0:
    logger.error('create_drafting: no columns in data')
    raise Exception('No column data provided to create_drafting.py')

n_total_rows = 1 + n_rows  # header row + data rows

# ── Column x positions/widths, from Excel's own column widths ──────────
default_col_w_mm = max(25.0, 180.0 / n_cols)
col_w_mm = [col_widths.get(c, default_col_w_mm) for c in range(n_cols)]
col_x_mm = [0.0] * (n_cols + 1)
for c in range(n_cols):
    col_x_mm[c + 1] = col_x_mm[c] + col_w_mm[c]

# ── Row y positions/heights, from Excel's own row heights ──────────────
# Still floored to a minimum so the chosen font size never gets clipped,
# even if Excel's row was set shorter than what pyLink's font needs.
min_row_h_mm = [
    (max(6.0, size_hdr_mm * 3.0) if r == 0 else max(5.0, size_dat_mm * 2.5))
    for r in range(n_total_rows)
]
row_h_mm = []
for r in range(n_total_rows):
    pts = row_heights.get(r, default_row_height)
    row_h_mm.append(max(pts * PT_MM, min_row_h_mm[r]))
row_y_mm = [0.0] * (n_total_rows + 1)
for r in range(n_total_rows):
    row_y_mm[r + 1] = row_y_mm[r] + row_h_mm[r]

# ── Merge lookup ─────────────────────────────────────────────────────────
# merge_anchor[(r,c)] = (r_end, c_end) for the top-left cell of a merge.
# merge_covered holds every other cell inside that merge, so they're
# skipped entirely (no duplicate background/border/text on top of the
# anchor's larger rectangle).
merge_anchor  = {}
merge_covered = set()
for (r1, c1, r2, c2) in merges:
    merge_anchor[(r1, c1)] = (r2, c2)
    for rr in range(r1, r2 + 1):
        for cc in range(c1, c2 + 1):
            if (rr, cc) != (r1, c1):
                merge_covered.add((rr, cc))

# ── Get or create drafting view ─────────────────────────────────────────

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
            .OfClass(ViewFamilyType):
        if v.ViewFamily == ViewFamily.Drafting:
            vft = v
            break
    # Broader fallback — any ViewFamilyType that supports drafting views
    if vft is None:
        for v in FilteredElementCollector(doc)\
                .OfClass(ViewFamilyType):
            try:
                if 'draft' in str(v.ViewFamily).lower():
                    vft = v
                    break
            except Exception:
                pass
    if vft is None:
        raise Exception(
            'No Drafting View family type found. '
            'Create any Drafting View manually first, then re-run pyLink.'
        )
    view = ViewDrafting.Create(doc, vft.Id)
    view.Name = _final_name

try:
    view.Scale = view_scale
except Exception:
    pass

# ── Draw table ────────────────────────────────────────────────────────────

hdr_txt_id = _get_text_type(hdr_tt_id)
dat_txt_id = _get_text_type(dat_tt_id)
default_line_id = _default_line_style_id()

for r in range(n_total_rows):
    is_header = (r == 0)
    for c in range(n_cols):
        if (r, c) in merge_covered:
            continue

        r_end, c_end = merge_anchor.get((r, c), (r, c))
        x = col_x_mm[c] * MM
        y = -row_y_mm[r] * MM
        w = (col_x_mm[c_end + 1] - col_x_mm[c]) * MM
        h = (row_y_mm[r_end + 1] - row_y_mm[r]) * MM

        style = cell_styles.get((r, c), {})

        # Background — only drawn if Excel actually set a fill on this
        # cell; nothing invented for cells Excel left plain.
        fill_rgb = style.get('fill_rgb')
        if fill_rgb:
            _draw_filled_rect(view, x, y, w, h, fill_type_id, tuple(fill_rgb))

        # Borders — only the sides Excel actually set, in that side's
        # own colour, all on the default <Line> style.
        top_rgb    = _cell_border_rgb(style, 'top')
        bottom_rgb = _cell_border_rgb(style, 'bottom')
        left_rgb   = _cell_border_rgb(style, 'left')
        right_rgb  = _cell_border_rgb(style, 'right')
        if top_rgb:
            _draw_border_line(view, x,     y,     x + w, y,
                               default_line_id, top_rgb,
                               _cell_border_weight(style, 'top'))
        if bottom_rgb:
            _draw_border_line(view, x,     y - h, x + w, y - h,
                               default_line_id, bottom_rgb,
                               _cell_border_weight(style, 'bottom'))
        if left_rgb:
            _draw_border_line(view, x,     y,     x,     y - h,
                               default_line_id, left_rgb,
                               _cell_border_weight(style, 'left'))
        if right_rgb:
            _draw_border_line(view, x + w, y,     x + w, y - h,
                               default_line_id, right_rgb,
                               _cell_border_weight(style, 'right'))

        # Text — plain black type, tinted via view override to this
        # cell's own colour when Excel set one.
        if is_header:
            text = fields[c] if c < len(fields) else ''
        else:
            record = records[r - 1] if (r - 1) < len(records) else []
            text = record[c] if c < len(record) else ''
        tt_id = hdr_txt_id if is_header else dat_txt_id
        color_rgb = style.get('color_rgb')
        rotation_rad = _excel_rotation_to_radians(style.get('rotation', 0)) if is_header else 0.0
        if is_header:
            halign = style.get('halign', 'Center')
            valign = style.get('valign', 'Bottom')
            wrap   = style.get('wrap', False)
        else:
            # Data rows: fixed Left/Top, ignoring Excel's actual
            # halign/valign/wrap (often Center/Center) - matches the
            # original pre-containment-fix look, which read better for
            # a plain data table than true centering does.
            halign = 'Left'
            valign = 'Top'
            wrap   = False
        _draw_text(view, x, y, w, h, text, tt_id,
                   tuple(color_rgb) if color_rgb else None,
                   rotation_rad, halign, valign, wrap)

logger.debug(
    'create_drafting: "{}" {}r x {}c'.format(
        _final_name, n_total_rows, n_cols
    )
)

try:
    _view_id_int = view.Id.IntegerValue
except AttributeError:
    _view_id_int = int(view.Id.Value)
PYLINK_RESULT = {'view_id': _view_id_int}
