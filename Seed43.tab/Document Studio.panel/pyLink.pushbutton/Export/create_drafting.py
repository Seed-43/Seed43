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
    CurveArray,
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
# Which table this is in a run sharing one view: 0 draws at the top and
# clears whatever was there, higher numbers stack below.
stack_index   = int(_p.get('stack_index', 0) or 0)
default_row_height = float(_p.get('default_row_height', 14.0))
# The card's scale. Columns and rows are Excel's own sizes times this,
# the text was set to Excel's point size times it by the caller - the
# same rule schedules follow, so a card comes out the same size whichever
# view type it builds.
native_scale  = float(_p.get('native_scale') or 0.7)
# Excel's real column widths, not the Revit-calibrated ones; absent for
# .ods, which falls back to col_widths.
col_widths_excel = _p.get('col_widths_excel') or col_widths
fill_type_id  = _p.get('fill_type_id', ElementId.InvalidElementId)
# Excel's manual page breaks (row index each new part starts at) and the
# Print Titles rows repeated at the top of every part; None = not set.
row_breaks    = sorted(set(int(b) for b in (_p.get('row_breaks') or [])))
repeat_rows   = _p.get('repeat_rows')
# The row's split choice: 'Excel' uses row_breaks above; 'Rows' puts
# split_value data rows in each part; 'Height' keeps each part no taller
# than split_value mm.
split_mode    = _p.get('split_mode') or 'Excel'
try:
    split_value = float(_p.get('split_value') or 0)
except Exception:
    split_value = 0.0
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
# Gap between flat text and the cell's left and right edges, like the one
# a schedule leaves inside its cells. Without it left-aligned text sat on
# the border line.
_CELL_PAD = 0.8 * MM
_TAB_MM = 1.0 * MM  # matches TEXT_TAB_SIZE forced on every pyLink text


def _note_layout(w, h, rotation_rad, halign, valign):
    """(width, alignment) a cell's TextNote is created with.

    Flat text: the cell width less its side padding, aligned as Excel's
    horizontal alignment says.

    Rotated text runs along the cell's HEIGHT, so Excel's VERTICAL
    alignment decides where along the note it sits: a bottom-aligned
    header starts at the foot of the cell, as Excel draws it. The note
    spans the cell height less its margins and is justified within that.
    Reading upwards (+90) the note's left end is at the bottom; reading
    downwards (-90) it is at the top.
    """
    if not rotation_rad:
        return w - 2 * _CELL_PAD, halign
    if rotation_rad > 0:
        align = {'Bottom': 'Left', 'Top': 'Right'}.get(valign, 'Center')
    else:
        align = {'Bottom': 'Right', 'Top': 'Left'}.get(valign, 'Center')
    return h - 2 * _TEXT_MARGIN, align
                     # type (see get_or_create_text_type in
                     # pylink_excel.py) - used here as the Top/Bottom
                     # edge padding in the Middle-Align-always vertical
                     # positioning formula below.

# Cache of measured probe offsets, keyed by everything that moves where Revit
# anchors a text box: (tt_id, width, halign, bold, text-component). Bold is in
# the key because bold glyphs measure larger at the same type/size, even though
# it is no longer part of the TYPE (see get_or_create_text_type in
# pylink_excel.py - FormattedText applies it per instance).
#
# The text-component (see _cache_text_component) is the real text for HEADER
# cells only, otherwise a shared placeholder. Two strings at the same
# width/halign/bold can measure different boxes if one wraps to more lines,
# which is exactly the narrow-header case. Data cells share one placeholder key
# per (type, width, halign, bold), since single-line values at those settings
# measure the same height - that is what keeps a 40-row column of "Grade 8.8"
# on ONE probe instead of 40.
#
# _batch_measure_probes pre-measures every key once per view, so the
# doc.Regenerate() it depends on - expensive, it walks the whole document -
# runs once per view rather than once per cache miss.
_probe_cache = {}


def _cache_text_component(text, is_header):
    """The text portion of a probe cache key - see _probe_cache's
    docstring for why header cells use their real text and data cells
    share one placeholder."""
    # Text that runs over more than one line always measures on its own:
    # a two-line heading below the first row shared a one-line cell's
    # measurement at the same width and sat half a line too low.
    t = unicode(text)
    if is_header or '\r' in t.strip() or '\n' in t.strip():
        return t
    return ''


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
        width, halign = _note_layout(spec['w'], spec['h'],
                                     spec['rotation_rad'], spec['halign'],
                                     spec['valign'])
        if width <= 0:
            continue

        tt_id, bold, is_header = (
            spec['tt_id'], spec['bold'], spec['is_header']
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
        width, note_align = _note_layout(w, h, rotation_rad, halign, valign)
        if width <= 0:
            return
        opts = TextNoteOptions(tt_id)
        opts.HorizontalAlignment = getattr(
            HorizontalTextAlignment, note_align, HorizontalTextAlignment.Left
        )
        # Flat text is laid out inside the cell less its side padding, so
        # it both starts off the border and wraps short of the far one.
        if not rotation_rad:
            x, w = x + _CELL_PAD, w - 2 * _CELL_PAD

        # ── Probe: measure exactly where Revit puts this box for
        # these exact settings, relative to a scratch origin - cached
        # per (type, width, halign, bold) so repeat cells (most of a
        # table) skip the expensive doc.Regenerate() entirely.
        probe_result = _measure_probe(tt_id, width, note_align, text, bold=bold, is_header=is_header)
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
            # The padded width (x, w), not the margin-inset mx/mw: the frame
            # is declared at width=w above, so the target must match or the
            # two disagree by the margin.
            target_left = {
                'Left':   x,
                'Center': x + (w - box_w) / 2.0,
                'Right':  x + w - box_w,
            }.get(halign, x)
            # Vertical: always target the text's CENTER line, held there by
            # Revit's Middle Align set below. Excel's Top/Center/Bottom become
            # an offset for that center point, not an edge to match:
            #   Top    -> tab + half the text height down from the top line
            #   Bottom -> mirrored, up from the bottom line
            #   Center -> the cell's exact vertical middle
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
            # The note spans the cell height less its margins and is
            # justified inside itself (see _note_layout), so the box is
            # simply centred on the cell.
            target_y_min = y - h / 2.0 - box_w / 2.0
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
            # The note spans the cell height less its margins and is
            # justified inside itself (see _note_layout), so the box is
            # simply centred on the cell.
            target_y_min = y - h / 2.0 - box_w / 2.0
            ox = target_x_min - bminy
            oy = target_y_min + bmaxx

        pt = XYZ(ox, oy, 0)
        tn = TextNote.Create(doc, view.Id, pt, width, str(text), opts)

        # Deliberately sets NEITHER TextElement.HorizontalAlignment NOR
        # .VerticalAlignment post-creation, though both are settable per
        # instance since Revit 2016. The math above already lands the note
        # exactly on both axes; setting either afterward re-justifies the
        # already-placed note against Revit's internal reference rather than
        # complementing it. Making the Paragraph panel show a non-default
        # align for cosmetic reasons is a separate concern from positioning.

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


def _curve_loop(pts):
    """Closed CurveLoop through pts (XY tuples, feet), in order."""
    loop = CurveLoop()
    n = len(pts)
    for i in range(n):
        a, b = pts[i], pts[(i + 1) % n]
        loop.Append(Line.CreateBound(XYZ(a[0], a[1], 0), XYZ(b[0], b[1], 0)))
    return loop


def _draw_filled_rect(view, x, y, w, h, frt_id, color_rgb):
    """Draw a solid FilledRegion rectangle, tinted via view override.
    (x, y) is the top-left corner."""
    if not frt_id or frt_id == ElementId.InvalidElementId or not color_rgb:
        return
    try:
        loop = _curve_loop([(x, y), (x + w, y), (x + w, y - h), (x, y - h)])
        fr = FilledRegion.Create(doc, frt_id, view.Id, [loop])
        _override_color(fr.Id, color_rgb)
    except Exception as ex:
        logger.debug('FilledRegion ({},{}) {}'.format(x, y, ex))


def _fill_outlines(rects):
    """Outline loops of each joined group of rectangles.

    rects: [(x0, x1, y_bottom, y_top), ...] in feet, all one colour. Cells
    that share an edge belong to one group; each group comes back as a
    list of loops (its outline, plus any holes), corners only. Works on
    the grid every rectangle edge sits on, so it handles merged cells and
    any shape a table makes, not just straight strips.
    """
    def k(v):
        return round(v, 7)
    xs = sorted(set(k(v) for r in rects for v in (r[0], r[1])))
    ys = sorted(set(k(v) for r in rects for v in (r[2], r[3])))
    xi = dict((v, i) for i, v in enumerate(xs))
    yi = dict((v, i) for i, v in enumerate(ys))
    cells = set()
    for x0, x1, y0, y1 in rects:
        for i in range(xi[k(x0)], xi[k(x1)]):
            for j in range(yi[k(y0)], yi[k(y1)]):
                cells.add((i, j))

    groups = []
    seen = set()
    for start in sorted(cells):
        if start in seen:
            continue
        comp, stack = [], [start]
        seen.add(start)
        while stack:
            i, j = stack.pop()
            comp.append((i, j))
            for nb in ((i + 1, j), (i - 1, j), (i, j + 1), (i, j - 1)):
                if nb in cells and nb not in seen:
                    seen.add(nb)
                    stack.append(nb)
        groups.append(comp)

    out = []
    for comp in groups:
        cs = set(comp)
        # Counter-clockwise edges of every cell; an edge two cells share
        # appears once each way and cancels, leaving the outline.
        nxt = {}
        for i, j in comp:
            if (i, j - 1) not in cs:
                nxt.setdefault((i, j), []).append((i + 1, j))
            if (i + 1, j) not in cs:
                nxt.setdefault((i + 1, j), []).append((i + 1, j + 1))
            if (i, j + 1) not in cs:
                nxt.setdefault((i + 1, j + 1), []).append((i, j + 1))
            if (i - 1, j) not in cs:
                nxt.setdefault((i, j + 1), []).append((i, j))
        loops = []
        while nxt:
            p0 = min(nxt)
            pts, p = [p0], p0
            while True:
                q = nxt[p].pop()
                if not nxt[p]:
                    del nxt[p]
                if q == p0:
                    break
                pts.append(q)
                p = q
            # Corners only: drop points in the middle of a straight run.
            n = len(pts)
            corners = [pts[m] for m in range(n)
                       if not ((pts[m - 1][0] == pts[m][0] == pts[(m + 1) % n][0])
                               or (pts[m - 1][1] == pts[m][1] == pts[(m + 1) % n][1]))]
            loops.append([(xs[a], ys[b]) for a, b in corners])
        out.append((loops, comp))
    return out, xs, ys


def _draw_merged_fills(view, fill_cells, frt_id):
    """One FilledRegion per joined group of same-colour cells, instead of
    one per cell: a banded row is one strip, a block of shaded cells one
    shape. A group Revit refuses as one shape (two cells touching only at
    a corner) falls back to a region per cell, as before."""
    if not frt_id or frt_id == ElementId.InvalidElementId:
        return
    by_rgb = {}
    for x, y, w, h, rgb in fill_cells:
        by_rgb.setdefault(rgb, []).append((x, x + w, y - h, y))
    for rgb, rects in by_rgb.items():
        groups, xs, ys = _fill_outlines(rects)
        for loops, comp in groups:
            try:
                fr = FilledRegion.Create(doc, frt_id, view.Id,
                                         [_curve_loop(lp) for lp in loops])
                _override_color(fr.Id, rgb)
            except Exception as ex:
                logger.debug('Merged fill fell back to cells: {}'.format(ex))
                for i, j in comp:
                    _draw_filled_rect(view, xs[i], ys[j + 1],
                                      xs[i + 1] - xs[i], ys[j + 1] - ys[j],
                                      frt_id, rgb)


_LINE_PROF = [0.0, 0.0, 0.0]   # create / set style / override seconds (profiling)


def _draw_border_line(view, x1, y1, x2, y2, line_style_id, color_rgb, weight=None):
    """Draw a single DetailLine on the default <Lines> style, tinted via
    a per-instance view override to the border's actual colour and
    (best-effort) weight."""
    try:
        import time as _t
        _t0 = _t.time()
        seg = Line.CreateBound(XYZ(x1, y1, 0), XYZ(x2, y2, 0))
        dc = doc.Create.NewDetailCurve(view, seg)
        _t1 = _t.time()
        if line_style_id and line_style_id != ElementId.InvalidElementId:
            try:
                dc.LineStyle = doc.GetElement(line_style_id)
            except Exception as ex:
                logger.warning('Set <Lines> style: {}'.format(ex))
        _t2 = _t.time()
        _override_color(dc.Id, color_rgb or (0, 0, 0), weight)
        _t3 = _t.time()
        _LINE_PROF[0] += _t1 - _t0
        _LINE_PROF[1] += _t2 - _t1
        _LINE_PROF[2] += _t3 - _t2
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

# ── Column x positions/widths ──────────────────────────────────────────
# Excel's own widths are the starting point, but they cannot be the last
# word: a column narrower than a word it shows has Revit break that word
# wherever it runs out, mid-word, into the line below. Hence the floor:
# every column ends up at least as wide as the widest unbreakable word it
# has to show flat, at the text size this view is actually using. A
# rotated header only needs one line of text across the column.
#
# Same rule the schedule builder follows, and for the same reason:
# wrapping can only break at a space, so a column narrower than the
# longest word is a column that clips or splits it.
default_col_w_mm = max(25.0, 180.0 / n_cols) * native_scale
col_w_mm = [col_widths_excel[c] * native_scale if c in col_widths_excel
            else default_col_w_mm for c in range(n_cols)]

_all_rows = [fields] + list(records)
_merge_span = {}
for _r1, _c1, _r2, _c2 in merges:
    _merge_span[(_r1, _c1)] = (_r2, _c2)

try:
    from pylink_shared import (measure_text_mm, longest_word_mm,
                               wrap_text_to_lines, CELL_LINE_BREAK)

    _word_need = {}          # {col: mm} single-column cells
    _span_need = []          # [(mm, [cols])] merged cells
    for _ri, _row in enumerate(_all_rows):
        _size = size_hdr_mm if _ri == 0 else size_dat_mm
        for _ci in range(min(len(_row), n_cols)):
            _cell = _row[_ci]
            _text = u'' if _cell is None else unicode(_cell).strip()
            if not _text:
                continue
            _st = cell_styles.get((_ri, _ci), {})
            if _ri == 0 and _st.get('rotation'):
                # A rotated header runs up the cell, so across the column
                # it needs one line of text, not its longest word. Measured
                # flat it widened every rotated column to a different size.
                _need = _size * 1.2
            else:
                _need = longest_word_mm(_text, _st.get('font_name', font),
                                        _size, _st.get('bold'),
                                        _st.get('italic'))
            if (_ri, _ci) in _merge_span:
                _last = min(_merge_span[(_ri, _ci)][1], n_cols - 1)
                _span_need.append((_need, list(range(_ci, _last + 1))))
            else:
                _word_need[_ci] = max(_word_need.get(_ci, 0.0), _need)

    _pad = 1.5 + 2 * _CELL_PAD / MM   # text slack plus the side padding
    for _c in range(n_cols):
        if _c in _word_need:
            col_w_mm[_c] = max(col_w_mm[_c], _word_need[_c] + _pad)

    # A merged heading cannot stretch one column, so any shortfall is
    # shared across the ones it covers.
    for _need, _cols in _span_need:
        _short = _need + _pad - sum(col_w_mm[c] for c in _cols)
        if _short > 0:
            for _c in _cols:
                col_w_mm[_c] += _short / len(_cols)

    logger.debug('drafting col widths mm after word floor: {}'.format(
        ['%.1f' % w for w in col_w_mm]))
except Exception as _ex:
    logger.warning('drafting column floor failed, using Excel widths: '
                   '{}'.format(_ex))

col_x_mm = [0.0] * (n_cols + 1)
for c in range(n_cols):
    col_x_mm[c + 1] = col_x_mm[c] + col_w_mm[c]


def _dv_prewrapped(ri, ci, text, available_mm):
    """Cell text broken at spaces to fit the space it has.

    Revit wraps a TextNote by itself, but it will break inside a word to
    do it - which is how "Thickness (ti)" arrived as "Thickn / ess (ti)".
    Breaking here first puts every break on a space, and the column floor
    computed above guarantees each resulting line actually fits.

    available_mm is the cell's own extent along the direction the text
    runs, IN MILLIMETRES: its width normally, its height for a rotated
    header. Callers hold those in feet, so they divide by MM first.
    """
    text = u'' if text is None else unicode(text)
    if not text.strip() or available_mm <= 0:
        return text
    st = cell_styles.get((ri, ci), {})
    try:
        lines = wrap_text_to_lines(
            text, st.get('font_name', font),
            size_hdr_mm if ri == 0 else size_dat_mm,
            available_mm - 1.0, st.get('bold'), st.get('italic'))
    except Exception:
        return text
    return CELL_LINE_BREAK.join(lines) if len(lines) > 1 else text

# ── Row y positions/heights, from Excel's own row heights ──────────────
# Excel's height times the card's scale, like a schedule's rows. Floored
# at a line and a half of text so a row Excel set too short for its font
# still can't clip it.
min_row_h_mm = [(size_hdr_mm if r == 0 else size_dat_mm) * 1.5
                for r in range(n_total_rows)]
row_h_mm = []
for r in range(n_total_rows):
    pts = row_heights.get(r, default_row_height)
    row_h_mm.append(max(pts * PT_MM * native_scale, min_row_h_mm[r]))
# ── Parts: Excel's manual page breaks ───────────────────────────────────
# A long table split with Page Layout > Breaks is drawn as parts side by
# side, each starting with the rows Excel's Print Titles repeat. With no
# Print Titles, the rows above the first one holding a number repeat -
# for a typical data table, its title and column headings.
PART_GAP_MM = 10.0


def _is_number(v):
    try:
        float(unicode(v).strip())
        return True
    except Exception:
        return False


if repeat_rows is not None:
    _repeat = sorted(set(int(r) for r in repeat_rows
                         if 0 <= int(r) < n_total_rows))
else:
    _first_num = n_total_rows
    for _ri, _row in enumerate(_all_rows):
        if any(_is_number(_v) for _v in _row if _v not in (None, '')):
            _first_num = _ri
            break
    _repeat = list(range(0, max(1, _first_num)))
# Data starts below the heading rows; a split never cuts into them.
_data0 = (max(_repeat) + 1) if _repeat else 1
_head_mm = sum(row_h_mm[r] for r in _repeat)

if split_mode == 'Rows' and split_value >= 1:
    _n = int(round(split_value))
    _breaks = list(range(_data0 + _n, n_total_rows, _n))
elif split_mode == 'Height' and split_value > 0:
    _breaks = []
    _acc = sum(row_h_mm[r] for r in range(_data0))   # part 1's headings
    for _r in range(_data0, n_total_rows):
        if _acc + row_h_mm[_r] > split_value and _r > (_breaks[-1] if _breaks
                                                         else _data0):
            _breaks.append(_r)
            _acc = _head_mm
        _acc += row_h_mm[_r]
elif split_mode == 'Excel':
    _breaks = [b for b in row_breaks if 0 < b < n_total_rows]
else:
    _breaks = []
if split_mode != 'Excel':
    logger.debug('split {} {} -> parts start at {}'.format(
        split_mode, split_value, _breaks))

parts = []   # [[source row, ...], ...]
_edges = [0] + _breaks + [n_total_rows]
for _p_i in range(len(_edges) - 1):
    _body = list(range(_edges[_p_i], _edges[_p_i + 1]))
    if _p_i == 0:
        parts.append(_body)
    else:
        parts.append([r for r in _repeat if r < _edges[_p_i]] + _body)

_table_w_mm = col_x_mm[n_cols]


def _part_layout(rows):
    """Top y (mm down from the part's top) of each row, plus this part's
    own merges: a merge is kept for the rows of it the part shows, so a
    titles row merged across columns stays merged in every part."""
    y = {}
    acc = 0.0
    for r in rows:
        y[r] = acc
        acc += row_h_mm[r]
    pos = dict((r, i) for i, r in enumerate(rows))
    anchor, covered = {}, set()
    for (r1, c1, r2, c2) in merges:
        shown = [r for r in rows if r1 <= r <= r2]
        if not shown:
            continue
        # Rows of a merge are only joined where the part shows them one
        # after another; anything else is drawn as separate cells.
        if pos[shown[-1]] - pos[shown[0]] != len(shown) - 1:
            continue
        top, bot = shown[0], shown[-1]
        anchor[(top, c1)] = (bot, c2)
        for rr in shown:
            for cc in range(c1, c2 + 1):
                if (rr, cc) != (top, c1):
                    covered.add((rr, cc))
    return y, acc, anchor, covered

# ── Get or create drafting view ─────────────────────────────────────────

# Where this table starts, in feet. Zero unless it is stacking below
# another table already drawn on this view - see stack_index.
_stack_y0_ft = 0.0

# Gap between stacked tables, in mm. Roughly two data rows, which reads
# as a deliberate break rather than a mis-set row height.
STACK_GAP_MM = 8.0

import time as _tm
_T = [('begin', _tm.time())]
view = None
for v in FilteredElementCollector(doc)\
        .OfClass(ViewDrafting)\
        .WhereElementIsNotElementType():
    if v.Name == _final_name:
        view = v
        if stack_index > 0:
            # Sharing this view with a table already on it: keep what is
            # there and start below its lowest point. Only the FIRST row
            # of a run clears, which is what makes re-applying the run
            # replace it rather than pile another copy on top.
            _lowest = None
            for el in FilteredElementCollector(doc, view.Id).ToElements():
                try:
                    bb = el.get_BoundingBox(view)
                except Exception:
                    bb = None
                if bb is None:
                    continue
                if _lowest is None or bb.Min.Y < _lowest:
                    _lowest = bb.Min.Y
            if _lowest is not None:
                _stack_y0_ft = _lowest - STACK_GAP_MM * MM
            logger.debug('stacking table {} at y={:.1f}mm'.format(
                stack_index, _stack_y0_ft / MM))
        else:
            # Clear existing elements. One batched delete: Revit works out
            # dependencies once instead of once per element, which is most
            # of the time a big table spent being cleared. Falls back to
            # one at a time if the batch is refused.
            import clr, System
            from System.Collections.Generic import List as _List
            from Autodesk.Revit.DB import ElementMulticlassFilter
            _classes = _List[System.Type]()
            for cls in (CurveElement, TextNote, FilledRegion, ImageInstance):
                _classes.Add(clr.GetClrType(cls))
            _ids = _List[ElementId](
                FilteredElementCollector(doc, view.Id)
                .WherePasses(ElementMulticlassFilter(_classes))
                .ToElementIds())
            if _ids.Count:
                try:
                    doc.Delete(_ids)
                except Exception:
                    for _i in _ids:
                        try:
                            doc.Delete(_i)
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

_T.append(('find_clear_view', _tm.time()))
# Pass 1: backgrounds (drawn immediately - no batching benefit, each
# is its own FilledRegion) plus collecting border segments and text
# specs for passes 2-4 below, instead of drawing borders/text inline
# per cell.
h_border_segments = []  # [(y, x1, x2, rgb, weight), ...]
v_border_segments = []  # [(x, y1, y2, rgb, weight), ...]
cell_specs = []          # text-drawing specs, one per non-covered cell
fill_cells = []          # [(x, y, w, h, rgb), ...] joined up after the loop

_cells = []
for _p_i, _rows in enumerate(parts):
    _py, _ph, _anchor, _covered = _part_layout(_rows)
    _x0_mm = _p_i * (_table_w_mm + PART_GAP_MM)
    for r in _rows:
        for c in range(n_cols):
            if (r, c) in _covered:
                continue
            _cells.append((r, c, _x0_mm, _py, _anchor))

for r, c, _x0_mm, _py, _anchor in _cells:
    is_header = (r == 0)
    r_end, c_end = _anchor.get((r, c), (r, c))
    x = (_x0_mm + col_x_mm[c]) * MM
    y = _stack_y0_ft - _py[r] * MM
    w = (col_x_mm[c_end + 1] - col_x_mm[c]) * MM
    h = (_py[r_end] + row_h_mm[r_end] - _py[r]) * MM

    style = cell_styles.get((r, c), {})

    # Background — only where Excel actually set a fill; nothing
    # invented for cells Excel left plain. Collected and drawn after
    # the loop so touching cells of one colour become one region.
    fill_rgb = style.get('fill_rgb')
    if fill_rgb:
        fill_cells.append((x, y, w, h, tuple(fill_rgb)))

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

    # Text spec - measured in pass 3 (one batch doc.Regenerate() covering
    # every distinct probe) and drawn in pass 4. Bold/italic/underline come
    # from Excel's own per-cell flags rather than being inferred from
    # header-ness, and the type is never bold (see get_or_create_text_type),
    # so an unbolded header stays regular and a bolded data cell comes
    # through bold.
    if is_header:
        text = fields[c] if c < len(fields) else ''
    else:
        record = records[r - 1] if (r - 1) < len(records) else []
        text = record[c] if c < len(record) else ''
    color_rgb = style.get('color_rgb')
    _rot_rad = (_excel_rotation_to_radians(style.get('rotation', 0))
                if is_header else 0.0)
    # Rotated text runs down the cell, so the height is what it has
    # to fit inside; everything else runs across the width. x/y/w/h
    # are Revit internal units (feet) by this point, and the wrap
    # measures in millimetres - without the conversion every word
    # was wider than the cell and each one took its own line.
    text = _dv_prewrapped(r, c, text,
                          (h - 2 * _TEXT_MARGIN if _rot_rad
                           else w - 2 * _CELL_PAD) / MM)
    cell_specs.append({
        'x': x, 'y': y, 'w': w, 'h': h,
        'text':  text,
        'tt_id': hdr_txt_id if is_header else dat_txt_id,
        'color_rgb': tuple(color_rgb) if color_rgb else None,
        'rotation_rad': _rot_rad,
        'halign': style.get('halign', 'Center'),
        'valign': style.get('valign', 'Center' if not is_header else 'Bottom'),
        'bold':      bool(style.get('bold', False)),
        'italic':    bool(style.get('italic', False)),
        'underline': bool(style.get('underline', False)),
        'is_header': is_header,
    })

_T.append(('layout_loop', _tm.time()))
# Pass 3 runs FIRST: the probe measurement needs one doc.Regenerate(), and
# that regenerate also has to process every element drawn before it. Done
# here, before any fill, border or text exists, it only regenerates the
# probes themselves. Nothing below needs a regenerate until the commit.
_batch_measure_probes(cell_specs)
_T.append(('measure', _tm.time()))

# Pass 1b: backgrounds, touching same-colour cells joined into one region.
_draw_merged_fills(view, fill_cells, fill_type_id)

_T.append(('fills', _tm.time()))
# Pass 2: merge collinear border segments and draw the result.
_hm = _merge_collinear_segments(h_border_segments)
_vm = _merge_collinear_segments(v_border_segments)
_nlines = len(_hm) + len(_vm)
# NewDetailCurve costs more the more is already in the view (about 40 ms
# a line on a big table), so every border goes in with ONE
# NewDetailCurveArray call and is styled afterwards. Falls back to a line
# at a time if the array call is refused.
_all_lines = ([(x1, y, x2, y, rgb, w) for y, x1, x2, rgb, w in _hm] +
              [(x, y1, x, y2, rgb, w) for x, y1, y2, rgb, w in _vm])
_made = False
if _all_lines:
    try:
        import time as _t2
        _a0 = _t2.time()
        _ca = CurveArray()
        for _x1, _y1, _x2, _y2, _rgb, _w in _all_lines:
            _ca.Append(Line.CreateBound(XYZ(_x1, _y1, 0), XYZ(_x2, _y2, 0)))
        _dcs = doc.Create.NewDetailCurveArray(view, _ca)
        _made = True   # lines exist now; never redraw them singly
        _LINE_PROF[0] += _t2.time() - _a0
        _style_el = (doc.GetElement(default_line_id)
                     if default_line_id and default_line_id != ElementId.InvalidElementId
                     else None)
        _i = 0
        for _dc in _dcs:
            if _style_el is not None:
                try:
                    _dc.LineStyle = _style_el
                except Exception as _ex:
                    logger.warning('Set <Lines> style: {}'.format(_ex))
            _rgb, _w = _all_lines[_i][4], _all_lines[_i][5]
            _override_color(_dc.Id, _rgb or (0, 0, 0), _w)
            _i += 1
    except Exception as _aex:
        logger.debug('Border array failed: {}'.format(_aex))
if _all_lines and not _made:
    # partial/failed array: clear nothing here - a failed call made no lines
    for x1, y1, x2, y2, rgb, weight in _all_lines:
        _draw_border_line(view, x1, y1, x2, y2, default_line_id, rgb, weight)

_T.append(('borders', _tm.time()))
# (Pass 3, the probe measurement, now runs before the fills - see above.)

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

_T.append(('text', _tm.time()))
try:
    with open('C:/Users/Fred/Proton Drive/juvenciodas/My files/.Claude/Work Files/pylink_timing.log', 'a') as _fh:
        _fh.write('   DRAFT cells=%d borders=%d+%d->%d lines [create=%.2f style=%.2f override=%.2f] fills=%d | %s\n' % (
            len(cell_specs), len(h_border_segments), len(v_border_segments), _nlines, _LINE_PROF[0], _LINE_PROF[1], _LINE_PROF[2],
            len(fill_cells), ' '.join('%s=%.2f' % (_T[i+1][0], _T[i+1][1]-_T[i][1]) for i in range(len(_T)-1))))
except Exception:
    pass
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
