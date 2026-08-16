# -*- coding: utf-8 -*-
# v3-no-fontname
"""
pyLink - Export/create_schedule.py

Creates a dumb ViewSchedule from arbitrary tabular data.
Called by script.py via exec() with PYLINK_PAYLOAD injected.

Payload keys:
    view_name          : str   - schedule name in Revit
    fields             : list  - column header strings (row 0 of named range)
    records            : list  - list of lists of cell value strings
    font               : str   - font name, e.g. 'Arial'
    size_hdr_mm        : float - header row font size in mm (used as fallback only)
    size_dat_mm        : float - data row font size in mm (used as fallback only)
    hdr_tt_id          : DB.ElementId - pre-created pyLink header TextNoteType id
    dat_tt_id          : DB.ElementId - pre-created pyLink data TextNoteType id
    cell_styles        : dict  - {(r,c): style_dict} per-cell Excel formatting
    merges             : list  - [(r1,c1,r2,c2), ...] merged cell ranges
    row_heights        : dict  - {row_idx: pts} Excel row heights in points
    col_widths         : dict  - {col_idx: mm} Excel column widths in mm
    default_row_height : float - fallback row height in points

Pattern - identical to pyTransmit script_create_schedule.py:
  - ViewSchedule.CreateSchedule(doc, ElementId.InvalidElementId)
  - Assembly Code field + two impossible filters = body always empty
  - body.SetColumnWidth collapses the single empty body column
  - All data written into Header section via InsertRow/InsertColumn/SetCellText

Border colours - pyTransmit pattern:
  - Named line style subcategories under the Lines category are created
    on demand, one per unique RGB colour plus pyT Off (white = invisible).
  - Each border side receives the ElementId of its colour's subcategory.
  - This is the only correct way to set per-border colours in a schedule;
    there is no BorderXxxColor property on TableCellStyle.

Font colour override flag:
  - The property on TableCellStyle is TextColor.
  - The flag on CellStyleOverrideOptions is FontColor (NOT TextColor).
  - Both must be used together or the colour override is silently ignored.

Merged cells:
  - MergeCells() writes to the anchor (top-left) only.
  - Non-anchor cells in the merged range need force_bg / force_fg passes
    to pick up background and text colour from the anchor style.
"""

_p = globals().get('PYLINK_PAYLOAD', {})

from pyrevit import revit, DB, script
from Autodesk.Revit.DB import (
    ViewSchedule, ScheduleFilter, ScheduleFilterType,
    SectionType, ElementId, TableCellStyle, TableMergedCell,
    HorizontalAlignmentStyle, VerticalAlignmentStyle, Color,
    FilteredElementCollector, LinePatternElement,
)

logger = script.get_logger()
doc = revit.doc

MM = 1.0 / 304.8   # millimetres -> Revit internal feet
PT_TO_FT = 1.0 / 864.0  # points -> feet  (72pt/in, 12in/ft = 864)
# Row height visual correction factor:
# Excel 20pt row looks better at 8.0mm in Revit (standard = 7.056mm)
# Factor = 8.0 / 7.056 = 1.1339
# This compensates for Revit rendering text at ~88.6% of typographic em height,
# requiring taller rows to maintain the same visual padding as Excel.
ROW_H_FACTOR = 1.1339

# ---------------------------------------------------------------------------
# Payload
# ---------------------------------------------------------------------------

view_name   = _p.get('view_name',   'pyLink Schedule')
fields      = _p.get('fields',      [])
records     = _p.get('records',     [])
font        = _p.get('font',        'Arial')
size_hdr_mm = float(_p.get('size_hdr_mm', 2.5))
size_dat_mm = float(_p.get('size_dat_mm', 2.3))
hdr_tt_id   = _p.get('hdr_tt_id',  ElementId.InvalidElementId)
dat_tt_id   = _p.get('dat_tt_id',  ElementId.InvalidElementId)
cell_styles = _p.get('cell_styles', {})   # {(r,c): style_dict}
merges      = _p.get('merges',      [])   # [(r1,c1,r2,c2), ...]
row_heights = _p.get('row_heights', {})   # {row_idx: pts}
col_widths  = _p.get('col_widths',  {})   # {col_idx: mm}
DEFAULT_ROW_H_PT = _p.get('default_row_height', 14.0)

# ---------------------------------------------------------------------------
# Line style registry (pyTransmit border pattern)
#
# Line style subcategories CANNOT be created inside an open transaction.
# script.py pre-creates all needed subcategories before the main transaction
# and passes their ElementIds in the payload under 'line_ids'.
#
# _border_line_id and _border_off_id look up from that pre-built dict.
# ---------------------------------------------------------------------------

_line_ids = _p.get('line_ids', {})   # rgb_tuple -> ElementId


def _border_line_id(rgb):
    """Return the pre-created line style ElementId for this colour, or
    None if it genuinely isn't available (never InvalidElementId - the
    caller must treat None as 'leave this override alone', matching
    pyTransmit's guarded pattern)."""
    key = tuple(rgb) if not isinstance(rgb, tuple) else rgb
    eid = _line_ids.get(key)
    if eid is None:
        # Fallback: try black
        eid = _line_ids.get((0, 0, 0))
    return eid


def _border_off_id():
    """Return the invisible (white) border line style ElementId, or None
    if it genuinely isn't available."""
    return _line_ids.get((255, 255, 255))


# ---------------------------------------------------------------------------
# Schedule field helpers (pyTransmit pattern)
# ---------------------------------------------------------------------------

def get_sf_by_id(sched_def, param_id_int):
    """Find a schedulable field by its parameter integer ID."""
    for sf in sched_def.GetSchedulableFields():
        try:
            pid = None
            try:
                pid = sf.ParameterId.IntegerValue
            except Exception:
                try:
                    pid = int(sf.ParameterId.Value)
                except Exception:
                    pass
            if pid == param_id_int:
                return sf
        except Exception:
            pass
    return None


# ---------------------------------------------------------------------------
# Cell text helper
# ---------------------------------------------------------------------------

def _safe_text(sec, r, c, text):
    try:
        sec.SetCellText(r, c, str(text))
        logger.debug('SetCellText({},{}) = "{}"'.format(r, c, text))
    except Exception as ex:
        logger.error('SetCellText({},{}) FAILED: {}'.format(r, c, ex))


# ---------------------------------------------------------------------------
# Core style application
#
# Follows the pyTransmit apply_style pattern exactly:
#   1. Create TableCellStyle
#   2. Set all value properties (font, colours, alignment)
#   3. Get CellStyleOverrideOptions
#   4. Enable each override flag
#   5. Set the options back onto the style
#   6. Apply to cell via SetCellStyle
#
# NOTE: the property and its override flag are named differently and must
# both be set:
#   style.TextColor  sets the font colour
#   opts.FontColor   enables the override
#
# Border colours come from named Lines subcategories (see the line style
# registry above) - there is no BorderXxxColor property.
# ---------------------------------------------------------------------------

def _excel_rotation_to_revit(excel_rot):
    """
    Map Excel's alignment.textRotation convention to Revit's TableCellStyle
    TextOrientation convention (as used by pyTransmit: 90 or 270).

    Excel: 0 = horizontal, 1-90 = counter-clockwise from horizontal (reading
    bottom-to-top), 91-180 = clockwise (stored as 90+angle, reading
    top-to-bottom), 255 = stacked/vertical (not supported here).
    Revit:  90 = one direction, 270 = the other (pyTransmit's proven values).

    Excel's common 90 (bottom-to-top) maps to Revit's 270; anything in the
    91-180 clockwise range maps to Revit's 90. Anything else (0, 255,
    unrecognised) returns 0, meaning "no rotation override".
    """
    if not excel_rot:
        return 0
    if 1 <= excel_rot <= 90:
        return 270
    if 91 <= excel_rot <= 180:
        return 90
    return 0


def _apply_style(sec, r, c,
                 bold=False, italic=False, underline=False, size_pt=8.0,
                 bg_rgb=None, fg_rgb=None, halign='Left', valign='Bottom',
                 rotation=0,
                 font_name='Arial', tt_id=None,
                 border_top='',    border_top_color=None,
                 border_bottom='', border_bottom_color=None,
                 border_left='',   border_left_color=None,
                 border_right='',  border_right_color=None):
    """
    Apply a TableCellStyle to a header section cell.

    bg_rgb / fg_rgb: (r, g, b) tuples or None.
    border_*: non-empty string means show border, '' means hide.
    border_*_color: (r, g, b) for that side's line colour, or None
                    (None defaults to black when showing).
    """
    try:
        style = TableCellStyle()

        # Font name — resolve from TextNoteType when available so the
        # schedule inherits the pre-created pyLink type's font exactly.
        resolved_font = font_name
        if tt_id and tt_id != ElementId.InvalidElementId:
            try:
                tt = doc.GetElement(tt_id)
                if tt:
                    fp = tt.get_Parameter(DB.BuiltInParameter.TEXT_FONT)
                    if fp:
                        resolved_font = fp.AsString()
            except Exception:
                pass

        style.IsFontBold      = bold
        style.IsFontItalic    = italic
        style.IsFontUnderline = underline
        style.FontName     = resolved_font
        # TableCellStyle.TextSize calibration:
        # Excel 16pt visually matches Revit TextSize that getter returns 5.0mm.
        # 16pt typographic = 5.644mm, but Revit renders cap height ~5.0mm.
        # Empirical factor: pt * 1.1812 -> getter returns pt * 25.4/72 * 0.886
        # This gives: 16pt->5.0mm, 12pt->3.75mm, 11pt->3.44mm
        style.TextSize     = float(size_pt) * 1.1812

        if bg_rgb:
            style.BackgroundColor = Color(bg_rgb[0], bg_rgb[1], bg_rgb[2])

        if fg_rgb:
            # Property name: TextColor
            # Override flag name: FontColor  ← different, see module docstring
            style.TextColor = Color(fg_rgb[0], fg_rgb[1], fg_rgb[2])

        style.FontHorizontalAlignment = getattr(
            HorizontalAlignmentStyle, halign,
            HorizontalAlignmentStyle.Left
        )
        # Excel 'Center' -> Revit 'Middle'; everything else (Top/Bottom)
        # maps straight across by name.
        _valign_name = 'Middle' if valign == 'Center' else valign
        style.FontVerticalAlignment = getattr(
            VerticalAlignmentStyle, _valign_name,
            VerticalAlignmentStyle.Bottom
        )

        # ── Override options ──────────────────────────────────────────────
        opts = style.GetCellStyleOverrideOptions()
        opts.Bold                = True
        opts.Italics             = True
        opts.Underline           = True
        opts.FontSize            = True
        # opts.FontName does not exist on TableCellStyleOverrideOptions in Revit 2026
        # Font name is applied via style.FontName directly without a flag
        opts.BackgroundColor     = bg_rgb is not None
        opts.FontColor           = fg_rgb is not None  # flag is FontColor
        opts.HorizontalAlignment = True
        opts.VerticalAlignment   = True

        # ── Borders via line style ElementIds ────────────────────────────
        # pyTransmit pattern: a coloured subcategory Id shows a border, the
        # white 'pyT Off' Id hides it - colour lives in the subcategory, there
        # is no opts.BorderXxxColor. The master flag below is required on top
        # of the four per-side ones or Revit ignores the per-side style ids.
        opts.BorderLineStyle = True
        sides = [
            (border_top,    border_top_color,    'BorderTopLineStyle'),
            (border_bottom, border_bottom_color, 'BorderBottomLineStyle'),
            (border_left,   border_left_color,   'BorderLeftLineStyle'),
            (border_right,  border_right_color,  'BorderRightLineStyle'),
        ]
        for bstyle, bcolor, prop in sides:
            try:
                if bstyle and bstyle not in ('none', ''):
                    colour = bcolor if bcolor else (0, 0, 0)
                    line_id = _border_line_id(colour)
                else:
                    line_id = _border_off_id()
                if line_id is not None:
                    # Only touch this side if we actually have a real,
                    # valid line style for it (pyTransmit pattern) -
                    # never assign InvalidElementId, which produces
                    # undefined rendering instead of "no border".
                    setattr(opts,  prop, True)
                    setattr(style, prop, line_id)
                else:
                    setattr(opts, prop, False)
                    logger.warning(
                        'border {} has no valid line style - left unset'
                        .format(prop)
                    )
            except Exception as bex:
                logger.warning('border {} FAILED: {}'.format(prop, bex))

        style.SetCellStyleOverrideOptions(opts)

        # ── Text rotation (pyTransmit pattern) ───────────────────────────
        # rotation here is already mapped to Revit's convention (90/270)
        # by the caller — see _excel_rotation_to_revit().
        if rotation in (90, 270):
            try:
                opts2 = style.GetCellStyleOverrideOptions()
                style.TextOrientation = rotation
                opts2.TextOrientation = True
                style.SetCellStyleOverrideOptions(opts2)
            except Exception as rex:
                logger.warning('rotation FAILED: {}'.format(rex))

        sec.SetCellStyle(r, c, style)

    except Exception as ex:
        import traceback as _tb
        logger.warning('_apply_style({},{}) EXCEPTION: {} | {}'.format(
            r, c, ex, _tb.format_exc().splitlines()[-1]))


# ---------------------------------------------------------------------------
# Read-then-patch helpers for merged cells (pyTransmit force_bg / force_fg)
#
# MergeCells() applies a style only to the anchor (top-left) cell.
# The remaining cells in the merged range retain their own individual styles.
# To ensure a uniform background and text colour across the entire merged
# region, we read each non-anchor cell's existing style and patch just the
# colour properties — without disturbing borders or font settings.
# ---------------------------------------------------------------------------

def _force_bg(sec, r, c, bg_rgb):
    """Patch background colour on an already-styled cell (read-then-write)."""
    try:
        style = sec.GetTableCellStyle(r, c)
        opts  = style.GetCellStyleOverrideOptions()
        opts.BackgroundColor  = True
        style.BackgroundColor = Color(bg_rgb[0], bg_rgb[1], bg_rgb[2])
        style.SetCellStyleOverrideOptions(opts)
        sec.SetCellStyle(r, c, style)
    except Exception as ex:
        logger.debug('_force_bg({},{}) {}'.format(r, c, ex))


def _force_fg(sec, r, c, fg_rgb):
    """Patch font colour on an already-styled cell (read-then-write)."""
    try:
        style = sec.GetTableCellStyle(r, c)
        opts  = style.GetCellStyleOverrideOptions()
        opts.FontColor  = True       # flag is FontColor, not TextColor
        style.TextColor = Color(fg_rgb[0], fg_rgb[1], fg_rgb[2])
        style.SetCellStyleOverrideOptions(opts)
        sec.SetCellStyle(r, c, style)
    except Exception as ex:
        logger.debug('_force_fg({},{}) {}'.format(r, c, ex))


# ---------------------------------------------------------------------------
# Merge helper
# ---------------------------------------------------------------------------

def _safe_merge(sec, r1, c1, r2, c2):
    try:
        mc        = TableMergedCell()
        mc.Top    = r1;  mc.Bottom = r2
        mc.Left   = c1;  mc.Right  = c2
        sec.MergeCells(mc)
    except Exception as ex:
        logger.debug('MergeCells({},{},{},{}) {}'.format(r1, c1, r2, c2, ex))


# ---------------------------------------------------------------------------
# Main schedule build
# ---------------------------------------------------------------------------

n_cols     = len(fields)
n_rows     = len(records)
total_rows = 1 + n_rows   # header row + data rows
total_cols = n_cols

if n_cols == 0:
    logger.error('create_schedule: no columns in data')
    raise Exception('No column data provided to create_schedule.py')

# ── Column widths: Excel mm -> Revit feet ─────────────────────────────────────
default_col_w_ft = max(20.0, 190.0 / n_cols) / 304.8
col_w_ft = {}
for ci in range(n_cols):
    w_mm = col_widths.get(ci)
    col_w_ft[ci] = (w_mm / 304.8) if w_mm else default_col_w_ft

# ── Delete existing schedule with the same name ───────────────────────────────
# Always recreate to avoid stale SectionData handle errors on update.
for v in revit.query.get_elements_by_class(ViewSchedule, doc=doc):
    if v.Name == view_name:
        try:
            doc.Delete(v.Id)
        except Exception as _de:
            logger.error('Could not delete existing schedule: {}'.format(_de))
        break

# ── Create schedule (pyTransmit pattern) ──────────────────────────────────────
sched     = ViewSchedule.CreateSchedule(doc, ElementId.InvalidElementId)
sched.Name = view_name
sched_def  = sched.Definition

# ── Assembly Code field + two impossible filters = body always empty ──────────
FIELD_ID_ASM_CODE = -1002500
sf_asm = get_sf_by_id(sched_def, FIELD_ID_ASM_CODE)
if sf_asm is None:
    for sf in sched_def.GetSchedulableFields():
        sf_asm = sf
        break
if sf_asm is None:
    raise Exception('No schedulable fields available for the empty-body trick')

field_asm = sched_def.AddField(sf_asm)
field_asm.ColumnHeading = ''
# Do NOT set IsHidden — it collapses body columns and breaks header rendering.

sched_def.AddFilter(ScheduleFilter(
    field_asm.FieldId, ScheduleFilterType.Equal, 'NO VALUES FOUND'
))
sched_def.AddFilter(ScheduleFilter(
    field_asm.FieldId, ScheduleFilterType.Equal, 'ALL VALUES FOUND'
))

# ShowGridLines is Revit's blanket gridline layer, drawn underneath and
# independent of per-cell border overrides, so any cell without a successful
# override still fell back to it - which is why data cells with no Excel
# border rendered bordered anyway. Off, so borders come only from real
# per-cell TableCellStyle overrides.
#
# NOTE: do NOT set ShowHeaders=False or ShowTitle=False on ScheduleDefinition.
# They collapse whole table sections and break Header rendering, and
# SectionType.Header is where all custom content lives.
try:
    sched_def.ShowGridLines = False
except Exception as ex:
    logger.debug('ShowGridLines: {}'.format(ex))

# ── Fetch section handles ──────────────────────────────────────────────────────
table_data = sched.GetTableData()
hdr  = table_data.GetSectionData(SectionType.Header)
body = table_data.GetSectionData(SectionType.Body)

logger.debug('Sections: hdr={}r x {}c  body={}r x {}c'.format(
    hdr.NumberOfRows, hdr.NumberOfColumns,
    body.NumberOfRows, body.NumberOfColumns
))

# ── Collapse the empty body row to full schedule width ────────────────────────
total_w = sum(col_w_ft.values())
try:
    body.SetColumnWidth(0, total_w)
except Exception as ex:
    logger.debug('body.SetColumnWidth: {}'.format(ex))

# Hide all body borders so the collapsed row is truly invisible.
try:
    off_id = _border_off_id()
    if off_id is not None:
        _bs = TableCellStyle()
        _bo = _bs.GetCellStyleOverrideOptions()
        _bo.BorderTopLineStyle    = True;  _bs.BorderTopLineStyle    = off_id
        _bo.BorderBottomLineStyle = True;  _bs.BorderBottomLineStyle = off_id
        _bo.BorderLeftLineStyle   = True;  _bs.BorderLeftLineStyle   = off_id
        _bo.BorderRightLineStyle  = True;  _bs.BorderRightLineStyle  = off_id
        _bs.SetCellStyleOverrideOptions(_bo)
        body.SetCellStyle(_bs)
    else:
        logger.warning('body border suppression: no off_id available, skipped')
except Exception as ex:
    logger.debug('body border suppression: {}'.format(ex))

# ── Grow header grid to required size ─────────────────────────────────────────
# pyTransmit uses while-loops, not for-range, for IronPython 2 safety.
while hdr.NumberOfColumns < total_cols:
    hdr.InsertColumn(hdr.NumberOfColumns)
while hdr.NumberOfRows < total_rows:
    hdr.InsertRow(hdr.NumberOfRows)

logger.debug('Header grid: {}r x {}c'.format(
    hdr.NumberOfRows, hdr.NumberOfColumns
))

# ── Set column widths ─────────────────────────────────────────────────────────
for ci in range(total_cols):
    try:
        hdr.SetColumnWidth(ci, col_w_ft.get(ci, default_col_w_ft))
    except Exception as ex:
        logger.debug('SetColumnWidth({}): {}'.format(ci, ex))

# ── Set row heights from Excel (points -> feet) ───────────────────────────────
for ri in range(total_rows):
    excel_ht = row_heights.get(ri)
    ht_ft = (excel_ht if excel_ht else DEFAULT_ROW_H_PT) * PT_TO_FT * ROW_H_FACTOR
    try:
        hdr.SetRowHeight(ri, ht_ft)
    except Exception as ex:
        logger.debug('SetRowHeight({}): {}'.format(ri, ex))

# ── Apply merged cells ────────────────────────────────────────────────────────
# Merges must be applied before SetCellText and SetCellStyle.
# After merging, text and styles target the anchor cell (top-left).
for r1, c1, r2, c2 in merges:
    if r1 < total_rows and c1 < n_cols:
        _safe_merge(
            hdr, r1, c1,
            min(r2, total_rows - 1),
            min(c2, n_cols - 1)
        )

# ── Input + Revit state report ───────────────────────────────────────────────
from pyrevit import output as _out_mod
_out = _out_mod.get_output()

_out.print_md('---')
_out.print_md('## pyLink Report: `{}`'.format(view_name))

# -- Schedule state after setup --
_out.print_md('### Schedule grid')
_out.print_md('- hdr rows={} cols={}'.format(hdr.NumberOfRows, hdr.NumberOfColumns))
_out.print_md('- body rows={} cols={}'.format(body.NumberOfRows, body.NumberOfColumns))
_out.print_md('- total_rows={} total_cols={}'.format(total_rows, total_cols))

# -- Column widths --
_out.print_md('### Column widths (ft in Revit)')
for _ci in range(total_cols):
    _out.print_md('- col {}: {:.4f} ft ({:.1f} mm)'.format(
        _ci, col_w_ft.get(_ci, 0), col_w_ft.get(_ci, 0) * 304.8))

# -- line_ids --
_out.print_md('### Line style IDs in payload')
for _rgb, _eid in _line_ids.items():
    _out.print_md('- {} -> ElementId {}'.format(_rgb, _eid))

# -- Per-cell input styles --
_out.print_md('### Per-cell Excel styles (input)')
_all_input = [fields] + list(records)
for _ri, _row in enumerate(_all_input):
    for _ci in range(len(_row)):
        _cs = cell_styles.get((_ri, _ci), {})
        if not _cs:
            continue
        _out.print_md(
            '- **cell({},{})** `{}` '
            'bold=**{}** italic={} size={}pt font=`{}` halign={} '
            'fill=`{}` color=`{}` '
            'borders T=`{}` B=`{}` L=`{}` R=`{}`'.format(
                _ri, _ci, repr(str(_all_input[_ri][_ci])[:15]),
                _cs.get('bold'), _cs.get('italic'),
                round(_cs.get('font_size', 0), 1),
                _cs.get('font_name', ''),
                _cs.get('halign', ''),
                _cs.get('fill_rgb'),
                _cs.get('color_rgb'),
                _cs.get('border_top', ''),
                _cs.get('border_bottom', ''),
                _cs.get('border_left', ''),
                _cs.get('border_right', ''),
            )
        )

# ── Fill cells: text + per-cell Excel styling ─────────────────────────────────
all_rows = [fields] + list(records)

for ri, row_data in enumerate(all_rows):
    for ci, cell in enumerate(row_data):
        if ci >= n_cols:
            break

        cs = cell_styles.get((ri, ci), {})

        fn       = cs.get('font_name',  font)
        # cell_styles stores font_size in points; convert to mm for _apply_style
        # font_size in cell_styles is already in points (from Excel)
        # Fallback uses size_hdr_mm/size_dat_mm converted to pt
        fs_pt    = cs.get('font_size',
                          size_hdr_mm * (72.0 / 25.4) if ri == 0
                          else size_dat_mm * (72.0 / 25.4))
        bold     = cs.get('bold',   ri == 0)
        italic   = cs.get('italic', False)
        underline = cs.get('underline', False)
        halign   = cs.get('halign', 'Center' if ri == 0 else 'Left')
        valign   = cs.get('valign', 'Bottom')
        fill_rgb = cs.get('fill_rgb', (220, 220, 220) if ri == 0 else None)
        fg_rgb   = cs.get('color_rgb', None)
        # Rotation only matters on the header row - body data isn't rotated
        rotation = _excel_rotation_to_revit(cs.get('rotation', 0)) if ri == 0 else 0

        _safe_text(hdr, ri, ci, cell)
        # Data rows never get borders, regardless of what Excel's own
        # per-cell formatting says - all visible content lives in this
        # Header section (the Body is deliberately kept empty), and
        # borders are only wanted on the header row itself. Only the
        # header row (ri == 0) reads Excel's actual border formatting.
        if ri == 0:
            bt, bt_c = cs.get('border_top', ''),    cs.get('border_top_color', None)
            bb, bb_c = cs.get('border_bottom', ''), cs.get('border_bottom_color', None)
            bl, bl_c = cs.get('border_left', ''),   cs.get('border_left_color', None)
            br, br_c = cs.get('border_right', ''),  cs.get('border_right_color', None)
        else:
            bt, bt_c = '', None
            bb, bb_c = '', None
            bl, bl_c = '', None
            br, br_c = '', None
        _apply_style(
            hdr, ri, ci,
            bold=bold,
            italic=italic,
            underline=underline,
            size_pt=fs_pt,
            bg_rgb=fill_rgb,
            fg_rgb=fg_rgb,
            halign=halign,
            valign=valign,
            rotation=rotation,
            font_name=fn,
            tt_id=hdr_tt_id if ri == 0 else dat_tt_id,
            border_top=bt,       border_top_color=bt_c,
            border_bottom=bb,    border_bottom_color=bb_c,
            border_left=bl,      border_left_color=bl_c,
            border_right=br,     border_right_color=br_c,
        )

# ── Force background and font colour on non-anchor merged cells ───────────────
# SetCellStyle on the anchor cell does not propagate colour to the other cells
# in the merged region.  We do a read-then-patch pass here to close that gap.
for r1, c1, r2, c2 in merges:
    r2 = min(r2, total_rows - 1)
    c2 = min(c2, n_cols - 1)
    anchor_cs = cell_styles.get((r1, c1), {})
    anchor_bg = anchor_cs.get('fill_rgb', None)
    anchor_fg = anchor_cs.get('color_rgb', None)

    for mr in range(r1, r2 + 1):
        for mc in range(c1, c2 + 1):
            if mr == r1 and mc == c1:
                continue   # skip anchor, already styled
            if anchor_bg:
                _force_bg(hdr, mr, mc, anchor_bg)
            if anchor_fg:
                _force_fg(hdr, mr, mc, anchor_fg)

logger.debug('create_schedule complete: "{}" {}r x {}c'.format(
    view_name, hdr.NumberOfRows, hdr.NumberOfColumns
))

# ── Read back what Revit actually stored ─────────────────────────────────────
# ── Column and row dimension readback ────────────────────────────────────────
_out.print_md('### Revit dimensions (read-back)')
_out.print_md('**Columns:**')
for _ci in range(hdr.NumberOfColumns):
    try:
        _w_ft = hdr.GetColumnWidth(_ci)
        _w_mm = _w_ft * 304.8
        _excel_mm = col_widths.get(_ci, 0)
        _out.print_md('- col {}: {:.4f}ft = **{:.3f}mm** (Excel: {:.3f}mm, diff: {:.3f}mm)'.format(
            _ci, _w_ft, _w_mm, _excel_mm, _w_mm - _excel_mm))
    except Exception as _e:
        _out.print_md('- col {} readback failed: {}'.format(_ci, _e))

_out.print_md('**Rows:**')
for _ri in range(hdr.NumberOfRows):
    try:
        _h_ft = hdr.GetRowHeight(_ri)
        _h_mm = _h_ft * 304.8
        _excel_pt = row_heights.get(_ri, DEFAULT_ROW_H_PT)
        _excel_mm = _excel_pt * 25.4 / 72.0
        _out.print_md('- row {}: {:.5f}ft = **{:.3f}mm** (Excel: {:.1f}pt = {:.3f}mm, diff: {:.3f}mm)'.format(
            _ri, _h_ft, _h_mm, _excel_pt, _excel_mm, _h_mm - _excel_mm))
    except Exception as _e:
        _out.print_md('- row {} readback failed: {}'.format(_ri, _e))

_out.print_md('### Revit cell styles (read-back after write)')
for _ri in range(hdr.NumberOfRows):
    for _ci in range(hdr.NumberOfColumns):
        try:
            _rs = hdr.GetTableCellStyle(_ri, _ci)
            _ro = _rs.GetCellStyleOverrideOptions()
            _out.print_md(
                '- **cell({},{})** text=`{}` '
                'bold={} size={:.1f}pt font=`{}` '
                'bg=({},{},{}) fg=({},{},{}) '
                'overrides: bold={} size={} bg={} fg={}'.format(
                    _ri, _ci,
                    repr(hdr.GetCellText(_ri, _ci)[:15]),
                    _rs.IsFontBold,
                    _rs.TextSize,
                    _rs.FontName,
                    _rs.BackgroundColor.Red,
                    _rs.BackgroundColor.Green,
                    _rs.BackgroundColor.Blue,
                    _rs.TextColor.Red,
                    _rs.TextColor.Green,
                    _rs.TextColor.Blue,
                    _ro.Bold, _ro.FontSize,
                    _ro.BackgroundColor, _ro.FontColor,
                )
            )
        except Exception as _re:
            _out.print_md('- cell({},{}) read-back failed: {}'.format(_ri, _ci, _re))
