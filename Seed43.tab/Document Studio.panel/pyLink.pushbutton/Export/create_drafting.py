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
    BuiltInParameter, FormattedText,
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
    # last resort only, and not the exact "<Lines>" subcategory (plain
    # "Lines" instead) - kept so borders at least land on *some* real
    # style rather than none.
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
_TAB_MM = 1.0 * MM  # matches TEXT_TAB_SIZE forced on every pyLink text
                     # type (see get_or_create_text_type in
                     # pylink_excel.py) - used here as the Top/Bottom
                     # edge padding in the Middle-Align-always vertical
                     # positioning formula below.

# Cache of measured probe offsets, keyed by every setting that can
# change where Revit anchors a text box: (tt_id, width, halign, bold,
# text-component). Bold is included because bold glyphs measure wider/
# taller than regular ones at the same type/size, which is exactly why
# bold is in the cache key even though it's no longer part of the TYPE
# itself (see get_or_create_text_type in pylink_excel.py - bold/
# italic/underline are applied per-instance via FormattedText below,
# not baked into a dedicated "Bold" type per size).
#
# The text-component (see _cache_text_component) is the cell's real
# text for HEADER cells only, else a shared placeholder. Two different
# strings sharing the same width/halign/bold can measure a different
# bounding box if one wraps onto more lines than the other - narrow
# header columns are exactly that case (several often share the same
# Excel width, so a short one-line header and a long one that wraps to
# 3 lines would otherwise land on the same key). Data cells share one
# placeholder key per (type, width, halign, bold) regardless of their
# own text - single-line values at a shared width/halign/bold measure
# the same box height, and this keeps the cache doing its actual job:
# a whole column of data (e.g. 40 rows of "Grade 8.8") shares ONE
# probe instead of one per cell.
#
# Every key this table's text needs is pre-measured in a single batch
# (_batch_measure_probes, called once per view) so the doc.Regenerate()
# this all depends on - expensive, it walks the whole document, not
# just the view - runs once per view instead of once per cache miss.
_probe_cache = {}


def _cache_text_component(text, is_header):
    """The text portion of a probe cache key - see _probe_cache's
    docstring for why header cells use their real text and data cells
    share one placeholder."""
    return str(text) if is_header else ''


def _apply_instance_formatting(tn, bold=False, italic=False, underline=False):
    """Apply bold/italic/underline to the FULL text of one TextNote
    instance via FormattedText - Revit's per-instance formatting API,
    independent of the note's Type. This is what lets a single
    'pyLink 7pt' type serve both a bold header cell and a regular
    data cell at the same point size, instead of needing a dedicated
    Bold type purely to carry that one flag (each per-cell Excel
    bold/italic/underline flag is honoured exactly as set, nothing
    forced or assumed)."""
    if not (bold or italic or underline):
        return
    try:
        ft = tn.GetFormattedText()
        rng = ft.AsTextRange()
        if bold:
            ft.SetBoldStatus(rng, True)
        if italic:
            ft.SetItalicStatus(rng, True)
        if underline:
            ft.SetUnderlineStatus(rng, True)
        tn.SetFormattedText(ft)
    except Exception as ex:
        logger.debug('Instance formatting failed: {}'.format(ex))


def _probe_key(tt_id, width, halign, bold, text, is_header):
    return (
        tt_id.IntegerValue if hasattr(tt_id, 'IntegerValue') else int(tt_id.Value),
        round(width, 6), halign, bold, _cache_text_component(text, is_header)
    )


def _measure_probe(tt_id, width, halign, sample_text, bold=False, is_header=False):
    """Cache lookup for one text box's measured offset. Normally
    every key is already populated by _batch_measure_probes before
    the draw pass runs, so this just returns the cached value; the
    create/regenerate/measure/delete path below only runs as a
    fallback for a key that pass somehow missed, so correctness never
    depends on the batch pass being exhaustive - just performance."""
    key = _probe_key(tt_id, width, halign, bold, sample_text, is_header)
    if key in _probe_cache:
        return _probe_cache[key]
    opts = TextNoteOptions(tt_id)
    opts.HorizontalAlignment = getattr(
        HorizontalTextAlignment, halign, HorizontalTextAlignment.Left
    )
    scratch = XYZ(0, 0, 0)
    probe = TextNote.Create(doc, view.Id, scratch, width, str(sample_text), opts)
    if bold:
        _apply_instance_formatting(probe, bold=True)
    doc.Regenerate()  # bbox is unreliable until regeneration
    bbox = probe.get_BoundingBox(view)
    doc.Delete(probe.Id)
    if bbox is None:
        return None
    result = (bbox.Min.X, bbox.Max.X, bbox.Min.Y, bbox.Max.Y)
    _probe_cache[key] = result
    return result


def _batch_measure_probes(cell_specs):
    """Pre-measure every distinct probe key this table's text needs
    in ONE doc.Regenerate() call instead of one per key: create every
    still-uncached probe TextNote first (no regenerate yet), regenerate
    once, then measure and delete them all. Populates _probe_cache
    directly; _measure_probe's own single-key path still runs during
    the draw pass as a fallback for anything this missed."""
    pending = []  # [(key, probe_element), ...]
    seen_keys = set()

    for spec in cell_specs:
        text = spec['text']
        if not str(text).strip():
            continue
        width = spec['h'] if spec['rotation_rad'] else spec['w']
        if width <= 0:
            continue

        tt_id, halign, bold, is_header = (
            spec['tt_id'], spec['halign'], spec['bold'], spec['is_header']
        )
        key = _probe_key(tt_id, width, halign, bold, text, is_header)
        if key in _probe_cache or key in seen_keys:
            continue
        seen_keys.add(key)

        try:
            opts = TextNoteOptions(tt_id)
            opts.HorizontalAlignment = getattr(
                HorizontalTextAlignment, halign, HorizontalTextAlignment.Left
            )
            probe = TextNote.Create(
                doc, view.Id, XYZ(0, 0, 0), width, str(text), opts
            )
            if bold:
                _apply_instance_formatting(probe, bold=True)
            pending.append((key, probe))
        except Exception as ex:
            logger.debug('Batch probe create failed for {}: {}'.format(key, ex))

    if not pending:
        return

    doc.Regenerate()  # single whole-document pass for every probe above

    for key, probe in pending:
        try:
            bbox = probe.get_BoundingBox(view)
            if bbox is not None:
                _probe_cache[key] = (bbox.Min.X, bbox.Max.X, bbox.Min.Y, bbox.Max.Y)
        except Exception as ex:
            logger.debug('Batch probe measure failed for {}: {}'.format(key, ex))

    for _key, probe in pending:
        try:
            doc.Delete(probe.Id)
        except Exception:
            pass


def _draw_text(view, x, y, w, h, text, tt_id, color_rgb,
                rotation_rad=0.0, halign='Left', valign='Bottom',
                bold=False, italic=False, underline=False, is_header=False):
    """
    Place a TextNote inside the cell (x, y, w, h) — (x, y) top-left —
    per Excel's halign/valign, by MEASURING where Revit actually puts
    a probe note rather than assuming what any alignment property
    does on its own.

    Horizontal: opts.HorizontalAlignment is set to match Excel's
    halign at creation time, and the probe measurement corrects the
    ORIGIN so the resulting box's left/center/right lands exactly on
    the cell's own left/center/right. Do not also set
    TextElement.HorizontalAlignment after creation - that re-justifies
    the box against Revit's own internal reference, fighting the
    already-correct position instead of complementing it.

    Vertical: Excel's Top/Center/Bottom is encoded purely as an OFFSET
    for where the text's center line should sit: Top = tab size + half
    the text's own measured height down from the cell's top line;
    Bottom = the mirror of that, up from the bottom line; Center = the
    cell's exact vertical middle. The manual probe-correction math
    below (same mechanism as horizontal) lands the note exactly there
    on its own - do NOT also set TextElement.VerticalAlignment after
    creation, same reason as HorizontalAlignment above.

      1. Create a probe with the EXACT same options/Width/text at a
         scratch origin (0,0,0).
      2. Regenerate, then read its real bounding box - this captures
         wherever Revit actually placed the box relative to that
         origin, whatever mechanism caused it.
      3. Work out, via the rotation math below, what origin WOULD
         have produced a box landing exactly on the target position -
         using the probe's measured offsets, not an assumed corner.
      4. Discard the probe, recreate at that corrected origin, rotate,
         then set VerticalAlignment=Middle.

    Width is always set to the cell's full trivial-axis dimension
    (matching the filled region exactly, not a separately-inset one).

    rotation_rad: 0, or the drafting rotation for Excel textRotation -
    positive (CCW, Excel 1-90) or negative (CW, Excel 91-180).
    """
    if not str(text).strip():
        return
    try:
        opts = TextNoteOptions(tt_id)
        opts.HorizontalAlignment = getattr(
            HorizontalTextAlignment, halign, HorizontalTextAlignment.Left
        )
        width = h if rotation_rad else w
        if width <= 0:
            return

        # ── Probe: measure exactly where Revit puts this box for
        # these exact settings, relative to a scratch origin - cached
        # per (type, width, halign, bold) so repeat cells (most of a
        # table) skip the expensive doc.Regenerate() entirely.
        probe_result = _measure_probe(tt_id, width, halign, text, bold=bold, is_header=is_header)
        if probe_result is None:
            logger.debug('TextNote bbox None for "{}" - skipped'.format(text))
            return
        bminx, bmaxx, bminy, bmaxy = probe_result

        # Margin only applies to the axis perpendicular to Width, to
        # keep text off the border line on that side.
        my, mh = y - _TEXT_MARGIN, max(h - 2 * _TEXT_MARGIN, 0)
        mx, mw = x + _TEXT_MARGIN, max(w - 2 * _TEXT_MARGIN, 0)

        if rotation_rad == 0:
            box_h = bmaxy - bminy
            box_w = bmaxx - bminx
            # Full cell width (x, w), not the margin-inset mx/mw - the
            # frame itself is declared at width=w (see "width" above),
            # so the target has to match that exactly or the two
            # disagree by the margin amount. Margin here is reserved
            # for the perpendicular (vertical) axis only, same as the
            # comment above already establishes.
            target_left = {
                'Left':   x,
                'Center': x + (w - box_w) / 2.0,
                'Right':  x + w - box_w,
            }.get(halign, x)
            # Vertical: always target the text's CENTER line (Revit's
            # own Middle Align, set below, is what actually holds it
            # there) - Excel's Top/Center/Bottom become an offset FOR
            # that center point, not an edge to match directly:
            #   Top    -> center sits tab + half the text's own height
            #             down from the cell's top line
            #   Bottom -> mirror of that, up from the bottom line
            #   Center -> the cell's own vertical middle, exactly
            half_text = box_h / 2.0
            target_center_y = {
                'Top':    y - _TAB_MM - half_text,
                'Center': y - h / 2.0,
                'Bottom': y - h + _TAB_MM + half_text,
            }.get(valign, y - h / 2.0)
            target_top = target_center_y + half_text
            ox = target_left - bminx
            oy = target_top - bmaxy
        elif rotation_rad > 0:
            # +90 CCW: local (dx,dy) -> screen (-dy, dx) relative to
            # origin, so post-rotation footprint (relative to origin)
            # is X:[-bmaxy,-bminy], Y:[bminx,bmaxx].
            box_h = bmaxy - bminy  # -> screen X extent (halign axis)
            box_w = bmaxx - bminx  # -> screen Y extent (valign axis)
            target_x_min = {
                'Left':   mx,
                'Center': mx + (mw - box_h) / 2.0,
                'Right':  mx + mw - box_h,
            }.get(halign, mx)
            # Same Middle-Align-always technique as the unrotated case,
            # just along the row-height axis (box_w here, since that's
            # what maps to "reading-direction extent" once rotated).
            half_text = box_w / 2.0
            target_center_y = {
                'Top':    y - _TAB_MM - half_text,
                'Center': y - h / 2.0,
                'Bottom': y - h + _TAB_MM + half_text,
            }.get(valign, y - h / 2.0)
            target_y_min = target_center_y - half_text
            ox = target_x_min + bmaxy
            oy = target_y_min - bminx
        else:
            # -90 CW: local (dx,dy) -> screen (dy, -dx) relative to
            # origin, so post-rotation footprint (relative to origin)
            # is X:[bminy,bmaxy], Y:[-bmaxx,-bminx].
            box_h = bmaxy - bminy  # -> screen X extent (halign axis)
            box_w = bmaxx - bminx  # -> screen Y extent (valign axis)
            target_x_min = {
                'Left':   mx,
                'Center': mx + (mw - box_h) / 2.0,
                'Right':  mx + mw - box_h,
            }.get(halign, mx)
            half_text = box_w / 2.0
            target_center_y = {
                'Top':    y - _TAB_MM - half_text,
                'Center': y - h / 2.0,
                'Bottom': y - h + _TAB_MM + half_text,
            }.get(valign, y - h / 2.0)
            target_y_min = target_center_y - half_text
            ox = target_x_min - bminy
            oy = target_y_min + bmaxx

        pt = XYZ(ox, oy, 0)
        tn = TextNote.Create(doc, view.Id, pt, width, str(text), opts)

        # NEITHER TextElement.HorizontalAlignment NOR .VerticalAlignment
        # get set here post-creation, despite both being real, settable
        # per-INSTANCE properties (since Revit 2016, replacing the old
        # TextNote.Align). The manual math above already computes and
        # lands on the exact correct position for both axes; setting
        # either alignment property again AFTERWARD re-justifies the
        # already-correctly-placed note against Revit's own internal
        # reference instead of complementing the position that's
        # already there. If Revit's own Paragraph-panel Vertical/
        # Horizontal Align ever needs to show something other than its
        # unset default for cosmetic/editability reasons, that's a
        # SEPARATE concern from positioning, not assumed harmless.

        if color_rgb:
            _override_color(tn.Id, color_rgb)
        _apply_instance_formatting(tn, bold=bold, italic=italic, underline=underline)

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


def _merge_collinear_segments(segments):
    """Merge touching/overlapping collinear border segments that share
    the same style into as few pieces as possible - a table row where
    every cell has the same bottom-border colour/weight draws as ONE
    line spanning the row instead of one abutting DetailLine per cell.
    `segments` is [(pos, start, end, rgb, weight), ...] where `pos` is
    the fixed coordinate (y for a horizontal run, x for a vertical
    one) and start/end is the extent along the other axis. Two cells'
    borders at the exact same position (e.g. one cell's bottom and the
    cell below's top, both set to the same colour) collapse into the
    same merged piece rather than drawing twice."""
    TOL = 1e-6
    by_key = {}
    for pos, s, e, rgb, weight in segments:
        key = (round(pos, 6), rgb, weight)
        by_key.setdefault(key, []).append((min(s, e), max(s, e)))

    merged = []
    for (pos, rgb, weight), spans in by_key.items():
        spans.sort()
        cur_s, cur_e = spans[0]
        for s, e in spans[1:]:
            if s <= cur_e + TOL:
                cur_e = max(cur_e, e)
            else:
                merged.append((pos, cur_s, cur_e, rgb, weight))
                cur_s, cur_e = s, e
        merged.append((pos, cur_s, cur_e, rgb, weight))
    return merged


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

# Pass 1: backgrounds (drawn immediately - no batching benefit, each
# is its own FilledRegion) plus collecting border segments and text
# specs for passes 2-4 below, instead of drawing borders/text inline
# per cell.
h_border_segments = []  # [(y, x1, x2, rgb, weight), ...]
v_border_segments = []  # [(x, y1, y2, rgb, weight), ...]
cell_specs = []          # text-drawing specs, one per non-covered cell

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

        # Borders — collected here, merged and drawn in pass 2 below:
        # a whole row/column of cells sharing the same border colour/
        # weight ends up as ONE DetailLine instead of one per cell.
        top_rgb    = _cell_border_rgb(style, 'top')
        bottom_rgb = _cell_border_rgb(style, 'bottom')
        left_rgb   = _cell_border_rgb(style, 'left')
        right_rgb  = _cell_border_rgb(style, 'right')
        if top_rgb:
            h_border_segments.append(
                (y, x, x + w, top_rgb, _cell_border_weight(style, 'top')))
        if bottom_rgb:
            h_border_segments.append(
                (y - h, x, x + w, bottom_rgb, _cell_border_weight(style, 'bottom')))
        if left_rgb:
            v_border_segments.append(
                (x, y, y - h, left_rgb, _cell_border_weight(style, 'left')))
        if right_rgb:
            v_border_segments.append(
                (x + w, y, y - h, right_rgb, _cell_border_weight(style, 'right')))

        # Text spec — measured in pass 3 (one batch doc.Regenerate()
        # for every distinct probe this table needs) and drawn in
        # pass 4. Bold/italic/underline come from Excel's OWN per-cell
        # flags, not assumed from the row being a header - the type
        # itself is never bold (see get_or_create_text_type), so a
        # header cell Excel didn't actually bold stays regular, and a
        # data cell Excel DID bold comes through bold too.
        if is_header:
            text = fields[c] if c < len(fields) else ''
        else:
            record = records[r - 1] if (r - 1) < len(records) else []
            text = record[c] if c < len(record) else ''
        color_rgb = style.get('color_rgb')
        cell_specs.append({
            'x': x, 'y': y, 'w': w, 'h': h,
            'text':  text,
            'tt_id': hdr_txt_id if is_header else dat_txt_id,
            'color_rgb': tuple(color_rgb) if color_rgb else None,
            'rotation_rad': _excel_rotation_to_radians(style.get('rotation', 0)) if is_header else 0.0,
            'halign': style.get('halign', 'Center'),
            'valign': style.get('valign', 'Center' if not is_header else 'Bottom'),
            'bold':      bool(style.get('bold', False)),
            'italic':    bool(style.get('italic', False)),
            'underline': bool(style.get('underline', False)),
            'is_header': is_header,
        })

# Pass 2: merge collinear border segments and draw the result.
for y, x1, x2, rgb, weight in _merge_collinear_segments(h_border_segments):
    _draw_border_line(view, x1, y, x2, y, default_line_id, rgb, weight)
for x, y1, y2, rgb, weight in _merge_collinear_segments(v_border_segments):
    _draw_border_line(view, x, y1, x, y2, default_line_id, rgb, weight)

# Pass 3: pre-measure every distinct text-box probe this table needs
# in one doc.Regenerate() call - see _batch_measure_probes.
_batch_measure_probes(cell_specs)

# Pass 4: place every cell's text. _measure_probe now hits the cache
# pass 3 already populated, so no further doc.Regenerate() calls
# happen here - the probe cache (keyed on type/width/halign/bold, plus
# text for header cells) keeps most cells sharing one measurement.
for spec in cell_specs:
    _draw_text(view, spec['x'], spec['y'], spec['w'], spec['h'],
               spec['text'], spec['tt_id'], spec['color_rgb'],
               spec['rotation_rad'], spec['halign'], spec['valign'],
               bold=spec['bold'], italic=spec['italic'],
               underline=spec['underline'], is_header=spec['is_header'])

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
