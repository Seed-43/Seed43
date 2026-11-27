# -*- coding: utf-8 -*-
from pyrevit import revit, DB
from pyrevit import forms
from pyrevit import script

import os
import json as _json
import zipfile as _zipfile
import re
import time as _time
import threading as _threading
import wpf
from System import Action as _Action
from System import Int64 as _Int64
from System.Windows import (
    Visibility, Thickness,
    VerticalAlignment, HorizontalAlignment,
    FontWeights, CornerRadius, TextTrimming,
    GridLength, GridUnitType
)
from System import DateTime
from System.Windows.Controls import (
    StackPanel, Border, CheckBox, TextBlock, TextBox,
    ComboBox, Button, Orientation, ScrollViewer,
    Grid, ColumnDefinition
)
from System.Windows.Controls.Primitives import Popup, ToggleButton
from System.Windows.Input import Cursors
from System.Windows.Media import Brushes
from System.Windows.Shapes import Ellipse

logger = script.get_logger()
doc = revit.doc
uidoc = revit.uidoc

try:
    from Snippets import _dialogs as sdlg
except Exception:
    sdlg = None

try:
    from Snippets._icons import make_icon as _mi
except Exception:
    _mi = None

"""
pylink_excel.py -- everything specific to the Excel side of pyLink:
dispatching xlsx/ods reads to tools/format/ (the actual zip/XML
parsing lives there - see pylink_xlsx.py and pylink_ods.py), building
native Schedule/Legend/Drafting views from that data, and the
ExcelCardMixin providing the Excel-specific parts of the row/card UI
(mixed into PyLinkWindow alongside WordCardMixin in pyLink.py).
"""

import sys as _sys
_sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pylink_shared import (
    hb, Row, VIEW_TYPES, SRC_COLOURS, STATUS_COLOURS,
    _run_export_script, _confirm, _alert, set_view_pylink_data, get_view_pylink_data,
    view_matches_type, capture_view_for_swap, restore_view_after_swap,
    load_excel_font_settings, _is_font_installed,
    NATIVE_SCALE_OPTIONS, DEFAULT_NATIVE_SCALE, BODY_PT_OPTIONS,
    detect_body_pt, detect_body_pt_multi, auto_native_scale, card_body_pt,
    SCHED_ONE, SCHED_PER_PART, SCHED_MODES, card_sched_mode,
    VIEW_NAME_RANGE, VIEW_NAME_SHEET, VIEW_NAME_BOTH, VIEW_NAME_MODES,
    card_native_scale, native_text_type_mm,
    SPLIT_EXCEL, SPLIT_ROWS, SPLIT_HEIGHT, SPLIT_MODES,
)

_sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'format'))
from pylink_xlsx import (
    read_xlsx_named_ranges, read_xlsx_range_data, read_xlsx_range_formatting,
)
from pylink_ods import (
    read_ods_named_ranges, read_ods_range_data, read_ods_range_formatting,
)


VIEW_TYPE_SCHEDULE = 'Schedule View'
VIEW_TYPE_DRAFTING = 'Drafting View'
VIEW_TYPE_LEGEND   = 'Legend View'

MM     = 1.0 / 304.8   # millimetres to Revit internal feet
PT_MM  = 0.352778      # millimetres per point (1pt = 1/72in = 25.4/72mm) -
                        # this is the answer to "what is 7pt in mm": 7 * PT_MM ~ 2.47mm


def _eid_int(eid):
    """ElementId as a plain int, across Revit API versions - Revit
    2024+ removed .IntegerValue in favour of .Value (Int64)."""
    try:
        return eid.IntegerValue
    except AttributeError:
        return int(eid.Value)



class TableRow(object):
    """Configuration for one import row from the UI."""

    def __init__(self):
        self.view_name   = ''
        self.named_range = ''
        self.sheet_name  = ''
        self.view_type   = VIEW_TYPE_SCHEDULE
        self.view_scale  = 1
        self.native_scale = None   # card's scale override (None = automatic)
        self.body_pt = None        # card's body text size override (pt)
        self.sched_mode  = SCHED_ONE               # card's schedule split mode
        self.file_path   = ''
        self.auto_sync   = False
        self.font        = 'Arial'    # used for pyLink text type lookup
        self.size_hdr_mm = 2.5        # header row font size in mm
        self.size_dat_mm = 2.3        # data row font size in mm
        # The Revit view name this row's UI counterpart previously
        # created (if any) — apply_row() uses this to refuse
        # overwriting a same-named view it doesn't actually own.
        self.applied_view_name = None
        # Same idea but by ElementId — the authoritative ownership
        # proof, since it survives the view being renamed outside
        # pyLink (unlike applied_view_name, which goes stale on rename).
        self.applied_view_id = None
        # Position in a run of rows that share one legend/drafting view.
        # 0 clears the view and draws at the top; anything higher draws
        # below what is already there. Schedules ignore it - a schedule
        # is one table by construction.
        self.stack_index = 0
        # Where a Legend/Drafting table splits into side-by-side parts -
        # see SPLIT_MODES in pylink_shared.
        self.split_mode  = SPLIT_EXCEL
        self.split_value = ''


# ── Read spreadsheet metadata + cell data (format dispatch) ──
# The zip/XML parsing lives in tools/format/ (pylink_xlsx.py, pylink_ods.py);
# these three functions only sniff which format a file's zip really is and
# hand off, keeping create_schedule/create_legend/create_drafting and
# ExcelCardMixin format-agnostic.

def _sniff_format(file_path):
    """'xlsx' or 'ods', from the zip's own namelist - or None if
    neither is recognised (legacy binary .xls, a corrupt/locked file,
    or something that isn't a zip at all)."""
    try:
        with _zipfile.ZipFile(file_path, 'r') as z:
            names = z.namelist()
    except Exception as e:
        logger.warning(
            'Could not read "{}" as a spreadsheet archive: {}'.format(
                os.path.basename(file_path), e
            )
        )
        return None
    if 'xl/workbook.xml' in names:
        return 'xlsx'
    if 'content.xml' in names:
        return 'ods'
    return None


def get_named_ranges_from_workbook(file_path):
    """
    Read named ranges and sheet names from an xlsx OR ods spreadsheet.
    Returns {'named_ranges': [...], 'sheets': [...], 'sheet_ranges': {...}}.
    """
    fmt = _sniff_format(file_path)
    if fmt == 'xlsx':
        return read_xlsx_named_ranges(file_path)
    if fmt == 'ods':
        return read_ods_named_ranges(file_path)
    logger.debug(
        'No named ranges found in "{}" (not an xlsx or ods archive - '
        'legacy .xls is not yet supported)'.format(
            os.path.basename(file_path)
        )
    )
    return {'named_ranges': [], 'sheets': [], 'sheet_ranges': {}}


# --- Parse cache ---
# One Apply used to open and fully parse the same workbook up to six times
# per row (data + formatting to draw, again for the view's tag hash, again
# for the card's hash). The parse is keyed on the file's own modified time
# and size, so saving the workbook in Excel always forces a fresh read.
# Results are copied out so no caller can change what the next one gets.
_PARSE_CACHE = {}
_PARSE_CACHE_MAX = 2000   # was 64: 43 rows x (data+format) overflowed it and
                          # wiped the cache on every status refresh
_HASH_CACHE = {}          # (file, range, sheet) -> (file stamp, md5)


def _file_stamp(file_path):
    try:
        st = os.stat(file_path)
        return (st.st_mtime, st.st_size)
    except Exception:
        return None


def _cached_parse(kind, file_path, named_range, sheet_name, parser):
    import copy as _copy
    stamp = _file_stamp(file_path)
    if stamp is None:
        return parser()
    key = (kind, file_path, named_range, sheet_name)
    hit = _PARSE_CACHE.get(key)
    if hit is not None and hit[0] == stamp:
        return _copy.deepcopy(hit[1])
    value = parser()
    if len(_PARSE_CACHE) >= _PARSE_CACHE_MAX:
        _PARSE_CACHE.clear()
    _PARSE_CACHE[key] = (stamp, _copy.deepcopy(value))
    return value


def read_named_range_data(file_path, named_range, sheet_name):
    """
    Read rows from an xlsx OR ods named range. Returns list of lists:
    [[header1, header2, ...], [val, val, ...], ...] - first row is
    column headers.
    """
    def _parse():
        fmt = _sniff_format(file_path)
        if fmt == 'xlsx':
            return read_xlsx_range_data(file_path, named_range, sheet_name)
        if fmt == 'ods':
            return read_ods_range_data(file_path, named_range, sheet_name)
        return []
    return _cached_parse('data', file_path, named_range, sheet_name, _parse)


def detect_card_body_pt(rows):
    """Body text size (pt) across all the ranges of a card's rows, so a card
    is understood as a whole document rather than range by range."""
    styles = []
    for r in rows:
        try:
            fmt = read_range_formatting(r.FilePath, r.NamedRange, r.Sheet)
            styles.append(fmt.get('cell_styles', {}))
        except Exception:
            continue
    return detect_body_pt_multi(styles)


def read_range_formatting(file_path, named_range, sheet_name):
    """
    Read cell formatting (fonts, fills, borders, alignment, merges,
    row/column sizing) from an xlsx OR ods named range - see
    pylink_xlsx.read_xlsx_range_formatting's docstring for the full
    result shape (cell_styles/merges/row_heights/col_widths).
    """
    def _parse():
        fmt = _sniff_format(file_path)
        if fmt == 'xlsx':
            return read_xlsx_range_formatting(file_path, named_range,
                                              sheet_name)
        if fmt == 'ods':
            return read_ods_range_formatting(file_path, named_range,
                                             sheet_name)
        return {'cell_styles': {}, 'merges': [], 'row_heights': {},
                'col_widths': {}}
    return _cached_parse('fmt', file_path, named_range, sheet_name, _parse)

# ── Schedule creation - pyTransmit pattern ──
# Uses ViewSchedule.CreateSchedule with ElementId.InvalidElementId (no
# category) and builds the entire table in the Header section.
# No Key Schedule, no project parameters required.

# ── pyLink text type manager ──
# Text types are named "pyLink <N>pt <Font>" and matched on font/size ONLY;
# bold/italic/underline are deliberately not part of a type's identity.
# FormattedText applies those per text RANGE inside a TextNote, independent of
# Type, so one "pyLink 7pt" serves every 7pt cell whatever its weight. A
# per-attribute naming scheme would instead bloat 100s of cells into 100s of
# near-duplicate types. Matching types are reused, otherwise one is created.

PREFIX = 'pyLink '

def _text_type_fingerprint(font, size_mm):
    """
    Canonical string key for a text style.
    Used to match existing types without creating duplicates.
    """
    return '{}__{:.4f}'.format(font.lower().strip(), size_mm)

def _read_fingerprint(tt):
    """
    Read font and size from an existing TextNoteType and return its
    fingerprint string. Returns None if the type cannot be read.
    """
    try:
        font_p = tt.get_Parameter(DB.BuiltInParameter.TEXT_FONT)
        size_p = tt.get_Parameter(DB.BuiltInParameter.TEXT_SIZE)

        if not font_p or not size_p:
            return None

        font    = font_p.AsString() or 'Arial'
        size_ft = size_p.AsDouble()
        size_mm = size_ft / MM  # convert feet back to mm

        return _text_type_fingerprint(font, size_mm)
    except Exception:
        return None

def _pt_name(size_mm, font):
    """'pyLink 7pt Arial' / 'pyLink 9pt Artifakt Element' - size
    derived from mm back to points (the unit Excel and the Revit UI's
    Text Size dialog both think in), so the name reads the same as
    what a person set in Excel. No Bold/Italic/Underline suffix -
    those are instance-level now, not part of the type."""
    pt = size_mm / PT_MM
    # Round to the nearest sane fraction so float noise (e.g. 7.0001)
    # doesn't spawn a near-duplicate type; 0.5pt is finer than anyone
    # actually picks in Excel's font-size box.
    pt = round(pt * 2) / 2.0
    pt_str = ('{:g}'.format(pt))
    return '{}{}pt {}'.format(PREFIX, pt_str, font)


def get_or_create_text_type(font='Arial', size_mm=2.3):
    """
    Return a TextNoteType matching font/size. Always plain black,
    never bold/italic/underlined at the TYPE level - per-cell text
    colour AND per-cell bold/italic/underline are both applied
    separately as instance-level formatting (view-specific graphic
    override for colour, FormattedText/TextRange for bold/italic/
    underline - see _draw_text in Export/create_drafting.py), not by
    creating a new Type per combination.
    Searches existing 'pyLink <N>pt <Font>' types first.
    Creates a new one if none match.

    Every pyLink type is also forced to Leader/Border Offset = 0,
    Tab Size = 1mm, and Width Factor = 0.75 on creation (and re-checked
    on reuse, same as the transparency fixup below) - Revit pads a
    TextNote's rendered width by the Leader/Border Offset on both
    sides, so leaving it non-zero is why a box sized exactly to a
    cell's width visually overflows or undershoots it; forcing it to 0
    makes the Width passed to TextNote.Create the actual visible text
    width, matching the cell. Width Factor (TEXT_WIDTH_SCALE) narrows
    the glyphs themselves so cells whose text was overflowing at 1.0
    fit properly at 0.75.
    """
    target_fp = _text_type_fingerprint(font, size_mm)
    target_name = _pt_name(size_mm, font)

    all_tt = list(
        DB.FilteredElementCollector(doc)
        .OfClass(DB.TextNoteType)
        .ToElements()
    )

    # Collect all existing pyLink types and check for a match
    pylink_types = []
    for tt in all_tt:
        try:
            name = tt.get_Parameter(
                DB.BuiltInParameter.SYMBOL_NAME_PARAM
            ).AsString()
        except Exception:
            continue

        if name and name.startswith(PREFIX):
            pylink_types.append((name, tt))
            fp = _read_fingerprint(tt)
            if fp == target_fp:
                logger.debug(
                    'Reusing text type "{}": {}'.format(name, target_fp)
                )
                # None of these are part of the fingerprint, so a type
                # created before these fixes existed (or before Bold
                # got folded into instance-level formatting) would
                # otherwise keep its old (wrong) values forever on
                # reuse. Check and correct each one.
                try:
                    fixes = []
                    bg_p = tt.get_Parameter(DB.BuiltInParameter.TEXT_BACKGROUND)
                    # 1 = Transparent, 0 = Opaque - NOT the other way
                    # around; an opaque background hides a cell's own
                    # fill colour drawn behind the text.
                    if bg_p and not bg_p.IsReadOnly and bg_p.AsInteger() != 1:
                        fixes.append((bg_p, 1))
                    lo_p = tt.get_Parameter(DB.BuiltInParameter.LEADER_OFFSET_SHEET)
                    if lo_p and not lo_p.IsReadOnly and lo_p.AsDouble() != 0.0:
                        fixes.append((lo_p, 0.0))
                    ts_p = tt.get_Parameter(DB.BuiltInParameter.TEXT_TAB_SIZE)
                    tab_ft = 1.0 * MM
                    if ts_p and not ts_p.IsReadOnly and abs(ts_p.AsDouble() - tab_ft) > 1e-9:
                        fixes.append((ts_p, tab_ft))
                    ws_p = tt.get_Parameter(DB.BuiltInParameter.TEXT_WIDTH_SCALE)
                    if ws_p and not ws_p.IsReadOnly and abs(ws_p.AsDouble() - 0.75) > 1e-9:
                        fixes.append((ws_p, 0.75))
                    # Bold/Italic aren't type-level, but a type from an older
                    # pyLink version may still carry them (a leftover
                    # "pyLink 7pt Bold" some cells still reference). Force
                    # both off so formatting only ever comes from the
                    # per-instance FormattedText path in create_drafting.py.
                    b_p = tt.get_Parameter(DB.BuiltInParameter.TEXT_STYLE_BOLD)
                    if b_p and not b_p.IsReadOnly and b_p.AsInteger() != 0:
                        fixes.append((b_p, 0))
                    i_p = tt.get_Parameter(DB.BuiltInParameter.TEXT_STYLE_ITALIC)
                    if i_p and not i_p.IsReadOnly and i_p.AsInteger() != 0:
                        fixes.append((i_p, 0))
                    # Apply naming-scheme changes to existing types too, not
                    # just new ones ("pyLink 7pt" -> "pyLink 7pt font type").
                    # If another type already holds the target name, Set()
                    # raises and this one keeps its current name - not worth
                    # failing the whole reuse over.
                    if name != target_name:
                        name_p = tt.get_Parameter(
                            DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                        if name_p and not name_p.IsReadOnly:
                            fixes.append((name_p, target_name))
                    if fixes:
                        with revit.Transaction(
                            'pyLink - Fix text type settings: {}'.format(name)
                        ):
                            for p, v in fixes:
                                p.Set(v)
                except Exception as ex:
                    logger.debug(
                        'Settings fixup on "{}": {}'.format(name, ex)
                    )
                return tt

    # No match — create a new one, named after its point size
    new_name = target_name
    if any(n == new_name for n, _ in pylink_types):
        # Extremely unlikely (would mean the fingerprint changed but
        # the rounded name didn't) - fall back to a suffix rather than
        # silently overwriting/duplicating a real type.
        new_name = '{} ({})'.format(new_name, len(pylink_types) + 1)

    # Duplicate from any existing TextNoteType as a base
    base = all_tt[0] if all_tt else None
    if not base:
        logger.error('No TextNoteType found to duplicate')
        return None

    size_ft = size_mm * MM

    with revit.Transaction(
        'pyLink - Create text type: {}'.format(new_name)
    ):
        new_tt = base.Duplicate(new_name)
        for bip, val in [
            (DB.BuiltInParameter.TEXT_FONT,         font),
            (DB.BuiltInParameter.TEXT_SIZE,          size_ft),
            # Always plain at type level: create_drafting.py applies
            # bold/italic/underline per instance via FormattedText from each
            # cell's own Excel flags, so ONE type covers both a bold header
            # and a plain data cell at the same point size.
            (DB.BuiltInParameter.TEXT_STYLE_BOLD,    0),
            (DB.BuiltInParameter.TEXT_STYLE_ITALIC,  0),
            # Black, and transparent rather than opaque — an opaque
            # text background would sit on top of and hide a cell's
            # fill colour, drawn separately as a FilledRegion behind
            # the text. TEXT_BACKGROUND is 1 = Transparent, 0 = Opaque.
            (DB.BuiltInParameter.LINE_COLOR,         0),
            (DB.BuiltInParameter.TEXT_BACKGROUND,    1),
            # Leader/Border Offset pads the rendered text box on both
            # sides - zero it so a TextNote's Width parameter equals
            # its actual visible width, matching the cell it's drawn
            # into exactly rather than overflowing/undershooting it.
            (DB.BuiltInParameter.LEADER_OFFSET_SHEET, 0.0),
            (DB.BuiltInParameter.TEXT_TAB_SIZE,       1.0 * MM),
            # Width Factor - narrows glyphs so text that was
            # overflowing a cell at the normal 1.0 factor fits at 0.75.
            (DB.BuiltInParameter.TEXT_WIDTH_SCALE,    0.75),
        ]:
            try:
                p = new_tt.get_Parameter(bip)
                if p and not p.IsReadOnly:
                    p.Set(val)
            except Exception as ex:
                logger.debug(
                    'Set param {} on {}: {}'.format(bip, new_name, ex)
                )

    logger.debug(
        'Created text type "{}": {}'.format(new_name, target_fp)
    )
    return new_tt


def purge_unused_pylink_text_types():
    """
    Delete every 'pyLink <N>pt <Font>' TextNoteType with zero TextNote
    instances currently using it anywhere in the document. Scoped to
    types this tool created (by name prefix) - never touches anything
    else in the project, unlike Revit's own Purge Unused command,
    which is document-wide and would strip a user's own unrelated
    types right alongside pyLink's accumulated ones.

    Bold/Italic no longer being part of a type's identity means far
    fewer types accumulate in the first place (one per point size,
    not one per size+bold combination) - this is for the types that
    piled up before that change, or from a range that's since been
    resized/reformatted so its old size is no longer used anywhere.

    A kept type is one that TextNotes still point at. "Still in use" on
    its own was no help when the project looked empty of text, so which
    views hold those notes comes back too - the answer is nearly always
    a legend or drafting view that is still there.

    Schedules are NOT a user: a schedule cell copies a text type's font
    NAME into its TableCellStyle and keeps no reference to the type, so
    a type held only by a schedule is genuinely unused and safe to go.

    Returns (purged_count, kept_count, failed_names, usage), where usage
    is [(type_name, n_notes, [view name, ...])].
    """
    all_tt = list(
        DB.FilteredElementCollector(doc)
        .OfClass(DB.TextNoteType)
        .ToElements()
    )
    pylink_tt = {}
    for tt in all_tt:
        try:
            name = tt.get_Parameter(
                DB.BuiltInParameter.SYMBOL_NAME_PARAM
            ).AsString()
        except Exception:
            continue
        if name and name.startswith(PREFIX):
            pylink_tt[_eid_int(tt.Id)] = (name, tt)

    if not pylink_tt:
        return (0, 0, [], [])

    # What is using each type, and which view it sits in. The view name
    # is the whole point: "2 still in use" against a project that looks
    # empty tells you nothing, "used by 14 notes in Legend AC1-1" tells
    # you where to go and delete them.
    note_use = {}                       # {type_id: [view name, ...]}
    try:
        for tn in DB.FilteredElementCollector(doc).OfClass(DB.TextNote):
            try:
                tid = _eid_int(tn.GetTypeId())
            except Exception:
                continue
            where = ''
            try:
                owner = doc.GetElement(tn.OwnerViewId)
                where = owner.Name if owner is not None else ''
            except Exception:
                pass
            note_use.setdefault(tid, []).append(where)
    except Exception as ex:
        logger.warning('Purge: TextNote scan failed: {}'.format(ex))
        return (0, len(pylink_tt), [], [])

    used_ids = set(note_use)

    unused = [
        (name, tt) for tid, (name, tt) in pylink_tt.items()
        if tid not in used_ids
    ]

    purged = 0
    failed = []
    if unused:
        with revit.Transaction('pyLink - Purge unused text types'):
            for name, tt in unused:
                try:
                    doc.Delete(tt.Id)
                    purged += 1
                except Exception as ex:
                    logger.debug(
                        'Purge failed for "{}": {}'.format(name, ex)
                    )
                    failed.append(name)

    kept = len(pylink_tt) - purged
    purged_ids = set()
    for name, tt in unused:
        if name not in failed:
            try:
                purged_ids.add(_eid_int(tt.Id))
            except Exception:
                pass

    usage = []
    for tid, (name, _tt) in pylink_tt.items():
        if tid in purged_ids:
            continue
        notes = note_use.get(tid, [])
        where = []
        for w in notes:
            if w and w not in where:
                where.append(w)
        usage.append((name, len(notes), where))
    usage.sort()
    return (purged, kept, failed, usage)

def _clear_header(hdr):
    """Strip header back to a single 1x1 cell."""
    while hdr.NumberOfRows > 1:
        try:
            hdr.RemoveRow(hdr.NumberOfRows - 1)
        except Exception:
            break
    while hdr.NumberOfColumns > 1:
        try:
            hdr.RemoveColumn(hdr.NumberOfColumns - 1)
        except Exception:
            break

def _safe_text(hdr, r, c, text):
    try:
        hdr.SetCellText(r, c, str(text))
    except Exception as ex:
        logger.debug('SetCellText({},{}) {}'.format(r, c, ex))

def _apply_style(hdr, r, c, bold=False, size_mm=2.5,
                 bg_rgb=None, halign='Left',
                 font='Arial', tt_id=None):
    """
    Apply font and background style to a header cell.
    Uses a pre-resolved pyLink TextNoteType ID (tt_id) when provided
    so no transaction is opened inside the schedule transaction.
    Falls back to raw font/size if no type ID is given.
    """
    try:
        from Autodesk.Revit.DB import (
            Color, HorizontalAlignmentStyle,
            VerticalAlignmentStyle, TableCellStyle
        )

        style = TableCellStyle()
        style.IsFontBold   = bold
        style.IsFontItalic = False

        # Resolve font name from the pre-created pyLink type
        resolved_font = font
        if tt_id and tt_id != DB.ElementId.InvalidElementId:
            try:
                tt = doc.GetElement(tt_id)
                if tt:
                    fp = tt.get_Parameter(DB.BuiltInParameter.TEXT_FONT)
                    if fp:
                        resolved_font = fp.AsString()
            except Exception:
                pass

        style.FontName = resolved_font
        style.TextSize = (size_mm / 0.75) * (72.0 / 25.4)

        if bg_rgb:
            style.BackgroundColor = Color(
                bg_rgb[0], bg_rgb[1], bg_rgb[2]
            )

        style.FontHorizontalAlignment = getattr(
            HorizontalAlignmentStyle, halign,
            HorizontalAlignmentStyle.Left
        )
        style.FontVerticalAlignment = VerticalAlignmentStyle.Middle

        opts = style.GetCellStyleOverrideOptions()
        opts.Bold                = True
        opts.Italics             = True
        opts.FontSize            = True
        # opts.FontName does not exist in Revit 2026 API
        opts.BackgroundColor     = bg_rgb is not None
        opts.HorizontalAlignment = True
        opts.VerticalAlignment   = True
        style.SetCellStyleOverrideOptions(opts)

        hdr.SetCellStyle(r, c, style)

    except Exception as ex:
        logger.debug('apply_style({},{}) {}'.format(r, c, ex))

def _safe_merge(hdr, r1, c1, r2, c2):
    try:
        from Autodesk.Revit.DB import TableMergedCell
        mc = TableMergedCell()
        mc.Top = r1;    mc.Bottom = r2
        mc.Left = c1;   mc.Right = c2
        hdr.MergeCells(mc)
    except Exception as ex:
        logger.debug(
            'MergeCells({},{},{},{}) {}'.format(r1, c1, r2, c2, ex)
        )

def create_schedule_from_data(view_name, fields, records,
                              font='Arial',
                              size_hdr_mm=2.5,
                              size_dat_mm=2.3,
                              hdr_tt_id=None,
                              dat_tt_id=None):
    """
    Create a dumb ViewSchedule with Excel data in the Header section.
    Exactly matches pyTransmit's approach:
    - CreateSchedule with InvalidElementId (multiple categories)
    - Assembly Code field with two impossible filters = body always empty
    - body.SetColumnWidth collapses the empty body row
    - All data goes into the Header section
    - Uses while loops to insert rows/cols matching pyTransmit's pattern

    Returns (ViewSchedule, status_string) or (None, error_string).
    """
    from Autodesk.Revit.DB import (
        ScheduleFilter, ScheduleFilterType, TableCellStyle,
        HorizontalAlignmentStyle, VerticalAlignmentStyle, Color
    )

    # Check for existing
    for v in revit.query.get_elements_by_class(DB.ViewSchedule, doc=doc):
        if v.Name == view_name:
            action = _confirm(
                'A schedule named "{}" already exists.\n'
                'Overwrite or skip?'.format(view_name),
                title='Schedule exists', yes='Overwrite', no='Skip'
            )
            if not action:
                return None, 'skipped'
            doc.Delete(v.Id)
            break

    n_cols = len(fields)
    n_rows = len(records)

    if n_cols == 0:
        return None, 'No columns in data'

    col_w_mm   = max(20.0, 190.0 / n_cols)
    col_w      = col_w_mm * MM
    row_h_hdr  = max(6.0, size_hdr_mm * 3.0) * MM
    row_h_data = max(5.0, size_dat_mm * 2.5) * MM
    total_cols = n_cols
    total_rows = 1 + n_rows  # 1 header + data rows

    # Create schedule - no category
    vs = DB.ViewSchedule.CreateSchedule(
        doc, DB.ElementId.InvalidElementId
    )
    vs.Name = view_name
    sched_def = vs.Definition

    # Add Assembly Code as hidden field with impossible filters
    # so body section always has zero rows (pyTransmit pattern)
    FIELD_ID_ASM_CODE = -1002500
    asm_sf = None
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
            if pid == FIELD_ID_ASM_CODE:
                asm_sf = sf
                break
        except Exception:
            pass

    # Fallback to first available field
    if asm_sf is None:
        for sf in sched_def.GetSchedulableFields():
            try:
                asm_sf = sf
                break
            except Exception:
                pass

    if asm_sf:
        try:
            f = sched_def.AddField(asm_sf)
            f.ColumnHeading = ''
            f.IsHidden = True
            sched_def.AddFilter(ScheduleFilter(
                f.FieldId, ScheduleFilterType.Equal, 'NO VALUES FOUND'
            ))
            sched_def.AddFilter(ScheduleFilter(
                f.FieldId, ScheduleFilterType.Equal, 'ALL VALUES FOUND'
            ))
            logger.debug('Hidden field: {}'.format(asm_sf.GetName(doc)))
        except Exception as ex:
            logger.debug('Hidden field failed: {}'.format(ex))

    try:
        sched_def.ShowGridLines = True
    except Exception:
        pass

    table_data = vs.GetTableData()
    hdr  = table_data.GetSectionData(DB.SectionType.Header)
    body = table_data.GetSectionData(DB.SectionType.Body)

    # Collapse body - pyTransmit sets body col to full width
    total_w = col_w * n_cols
    try:
        body.SetColumnWidth(0, total_w)
    except Exception:
        pass

    # Hide body borders
    try:
        _bs = TableCellStyle()
        _bo = _bs.GetCellStyleOverrideOptions()
        _bo.BorderTopLineStyle    = False
        _bo.BorderBottomLineStyle = False
        _bo.BorderLeftLineStyle   = False
        _bo.BorderRightLineStyle  = False
        _bs.SetCellStyleOverrideOptions(_bo)
        body.SetCellStyle(_bs)
    except Exception:
        pass

    # Insert columns - header starts at 1 col, use while loop like pyTransmit
    while hdr.NumberOfColumns < total_cols:
        hdr.InsertColumn(hdr.NumberOfColumns)
    for ci in range(total_cols):
        try:
            hdr.SetColumnWidth(ci, col_w)
        except Exception:
            pass

    # Insert rows - header starts at 1 row, use while loop like pyTransmit
    while hdr.NumberOfRows < total_rows:
        hdr.InsertRow(hdr.NumberOfRows)

    # Set row heights
    try:
        hdr.SetRowHeight(0, row_h_hdr)
    except Exception:
        pass
    for ri in range(1, total_rows):
        try:
            hdr.SetRowHeight(ri, row_h_data)
        except Exception:
            pass

    # Fill header row
    for ci, field in enumerate(fields):
        _safe_text(hdr, 0, ci, field)
        _apply_style(
            hdr, 0, ci,
            bold=True,
            size_mm=size_hdr_mm,
            bg_rgb=(220, 220, 220),
            halign='Center',
            font=font,
            tt_id=hdr_tt_id
        )

    # Fill data rows
    for ri, record in enumerate(records):
        for ci, cell in enumerate(record):
            if ci >= n_cols:
                break
            _safe_text(hdr, ri + 1, ci, cell)
            _apply_style(
                hdr, ri + 1, ci,
                bold=False,
                size_mm=size_dat_mm,
                halign='Left',
                font=font,
                tt_id=dat_tt_id
            )

    logger.debug(
        'Schedule "{}" — {}r x {}c'.format(view_name, total_rows, total_cols)
    )
    return vs, 'success'


# ── Legend / Drafting view creation ──

def create_drafting_view(view_name, scale):
    """Create a blank Drafting view."""
    vft = None
    for v in DB.FilteredElementCollector(doc)\
            .OfClass(DB.ViewFamilyType)\
            .WhereElementIsNotElementType():
        if v.ViewFamily == DB.ViewFamily.Drafting:
            vft = v
            break
    if not vft:
        _alert(
            'No Drafting View family type found.',
            exitscript=True
        )
    view = DB.ViewDrafting.Create(doc, vft.Id)
    view.Name = view_name
    view.Scale = scale
    return view

def create_legend_view(view_name, scale):
    """Create a Legend view by duplicating an existing one."""
    template = None
    for v in DB.FilteredElementCollector(doc)\
            .OfClass(DB.View)\
            .WhereElementIsNotElementType():
        if v.ViewType == DB.ViewType.Legend:
            template = v
            break
    if not template:
        _alert(
            'No Legend view found. Create one first.',
            exitscript=True
        )
    new_id = template.Duplicate(DB.ViewDuplicateOption.WithDetailing)
    view = doc.GetElement(new_id)
    view.Name = view_name
    view.Scale = scale
    return view

def place_table_in_view(view, fields, records):
    """
    Draw table data as native TextNote elements in a view.
    Used for Legend and Drafting views.
    Each cell = one TextNote positioned in a grid.
    """
    txt_type_id = DB.FilteredElementCollector(doc)\
        .OfClass(DB.TextNoteType)\
        .FirstElementId()

    if txt_type_id == DB.ElementId.InvalidElementId:
        logger.error('No TextNoteType found')
        return

    col_width  = 0.5    # feet
    row_height = 0.15   # feet

    all_rows = [fields] + [r for r in records]

    for r_idx, row in enumerate(all_rows):
        y = -(r_idx * row_height)
        for c_idx, cell in enumerate(row):
            if not str(cell).strip():
                continue
            x = c_idx * col_width
            pt = DB.XYZ(x, y, 0)
            opts = DB.TextNoteOptions(txt_type_id)
            opts.HorizontalAlignment = DB.HorizontalTextAlignment.Left
            DB.TextNote.Create(doc, view.Id, pt, str(cell), opts)


# ── Core apply logic ──

def _get_or_create_line_style(name, rgb):
    """
    Return the ElementId of a Lines subcategory with the given name/colour,
    or None if creation genuinely failed.

    Mirrors pyTransmit's _get_or_create_line_style exactly (proven pattern):
    explicitly sets a line weight (a fresh subcategory's default weight is
    not guaranteed to render visibly at schedule scale), and returns None
    on any failure instead of letting an exception or a half-built
    GraphicsStyle leak out as a usable-looking ElementId.

    Must be called INSIDE the transaction the caller wraps around it.
    """
    try:
        lines_cat = doc.Settings.Categories.get_Item('Lines')
        if lines_cat is None:
            return None
        existing = None
        for sub in lines_cat.SubCategories:
            if sub.Name == name:
                existing = sub
                break
        if existing is None:
            existing = doc.Settings.Categories.NewSubcategory(lines_cat, name)
        existing.LineColor = DB.Color(rgb[0], rgb[1], rgb[2])
        try:
            existing.SetLineWeight(1, DB.GraphicsStyleType.Projection)
        except Exception:
            pass
        gs = existing.GetGraphicsStyle(DB.GraphicsStyleType.Projection)
        return gs.Id if gs else None
    except Exception as ex:
        logger.error('Line style {} failed: {}'.format(name, ex))
        return None

def _pre_create_line_styles(cell_styles):
    """
    Pre-create all line style subcategories needed for borders in
    cell_styles, plus the always-needed 'pyT On' (black) and 'pyT Off'
    (invisible) styles.
    Returns a dict mapping rgb tuple -> ElementId. A colour that failed to
    create is simply absent from the dict — callers must treat a missing
    key as "no override available", never fall back to an invalid id.

    Must be called before the main schedule transaction, in its own
    transaction (mirrors pyTransmit: one transaction for all line styles,
    not one per style, and (0,0,0)/(255,255,255) are ALWAYS created
    regardless of what's actually in cell_styles - they're not optional
    extras, the border system depends on both existing).
    """
    needed = {(0, 0, 0), (255, 255, 255)}   # always needed, unconditionally
    for cs in cell_styles.values():
        for side in ('border_top_color', 'border_bottom_color',
                     'border_left_color', 'border_right_color'):
            c = cs.get(side)
            if c:
                needed.add(tuple(c))

    # Every row used to open a transaction and rewrite colour and weight of
    # each style even when nothing had changed. Styles already made this
    # session are reused as long as they still exist in the document.
    cache = _LINE_STYLE_CACHE.setdefault(_doc_key(), {})
    line_ids = {}
    todo = []
    for rgb in needed:
        eid = cache.get(rgb)
        if eid is not None and _id_alive(eid):
            line_ids[rgb] = eid
        else:
            todo.append(rgb)
    if not todo:
        return _with_invisible_off(line_ids)

    with revit.Transaction('pyLink - Line Styles'):
        for rgb in todo:
            if rgb == (255, 255, 255):
                name = 'pyT Off'
            elif rgb == (0, 0, 0):
                name = 'pyT On'
            else:
                name = 'pyT {:02X}{:02X}{:02X}'.format(rgb[0], rgb[1], rgb[2])
            eid = _get_or_create_line_style(name, rgb)
            if eid is not None:
                line_ids[rgb] = eid
                cache[rgb] = eid
    return _with_invisible_off(line_ids)


_INVISIBLE_CACHE = {}


def _invisible_lines_id():
    """ElementId of Revit's built-in <Invisible lines> style, or None."""
    key = _doc_key()
    eid = _INVISIBLE_CACHE.get(key)
    if eid is not None and _id_alive(eid):
        return eid
    try:
        for gs in DB.FilteredElementCollector(doc).OfClass(DB.GraphicsStyle):
            try:
                if (gs.Name == '<Invisible lines>'
                        and gs.GraphicsStyleType == DB.GraphicsStyleType.Projection):
                    _INVISIBLE_CACHE[key] = gs.Id
                    return gs.Id
            except Exception:
                continue
    except Exception as ex:
        logger.debug('Invisible lines lookup: {}'.format(ex))
    return None


def _with_invisible_off(line_ids):
    """A cell side with no Excel border is drawn with the 'off' style. A
    white line is invisible on screen but prints (and exports) black when
    the print setup makes every line black, so it showed as a grid. The
    <Invisible lines> style never prints. Falls back to the white style."""
    # Disabled: the white 'pyT Off' style is what pyTransmit uses, and
    # <Invisible lines> left the schedule's own grid showing.
    return line_ids


# Per-document caches for ids pyLink makes once and reuses every row.
_LINE_STYLE_CACHE = {}
_FILL_TYPE_CACHE = {}


def _doc_key():
    try:
        return doc.PathName or doc.Title
    except Exception:
        return ''


def _id_alive(eid):
    try:
        return doc.GetElement(eid) is not None
    except Exception:
        return False

def _type_name(el):
    """Safely read an ElementType's name. Reading .Name directly on
    Type-derived classes (FilledRegionType, TextNoteType, etc.) can
    throw in IronPython - this is why get_or_create_text_type above
    reads TextNoteType's name via SYMBOL_NAME_PARAM instead of .Name
    directly. Same fix, generalised for reuse here."""
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

def _get_or_create_pylink_fill_type():
    """
    Return the ElementId of a single reusable solid-fill FilledRegionType
    ('pyLink Fill') used for every cell background in drafting/legend
    views. Each cell's actual colour is applied afterwards as a
    view-specific graphic override on that FilledRegion instance
    (Override Graphics in View > By Element) rather than by creating a
    separate Type per colour — one Type total, not one per colour.
    Must be called OUTSIDE any open transaction — creating a type
    requires its own transaction context in Revit.

    Never raises — on any failure this returns InvalidElementId and
    logs the reason, rather than letting a type-creation problem take
    down the whole Apply run. The drawing code already treats a
    missing fill type as "skip the background for this run", same as
    a cell with no fill colour at all.
    """
    name = 'pyLink Fill'
    _cached = _FILL_TYPE_CACHE.get(_doc_key())
    if _cached is not None and _id_alive(_cached):
        return _cached
    try:
        for frt in DB.FilteredElementCollector(doc).OfClass(DB.FilledRegionType):
            if _type_name(frt) == name:
                _FILL_TYPE_CACHE[_doc_key()] = frt.Id
                return frt.Id
    except Exception as ex:
        logger.error('Fill type lookup failed: {}'.format(ex))
        return DB.ElementId.InvalidElementId

    # Find an actual "solid fill" drafting pattern to base the type on
    solid_pattern_id = DB.ElementId.InvalidElementId
    for fp in DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement):
        try:
            if fp.GetFillPattern().IsSolidFill:
                solid_pattern_id = fp.Id
                break
        except Exception:
            continue

    candidates = list(
        DB.FilteredElementCollector(doc).OfClass(DB.FilledRegionType)
    )
    if not candidates:
        logger.error('No FilledRegionType found to duplicate')
        return DB.ElementId.InvalidElementId

    # Try every candidate in turn — the first one Revit hands back from
    # the collector isn't guaranteed to be duplicable (e.g. a built-in
    # Masking Region type can refuse), so don't give up after one try.
    for base in candidates:
        try:
            with revit.Transaction('pyLink - Fill type: {}'.format(name)):
                new_frt = base.Duplicate(name)
                try:
                    if solid_pattern_id != DB.ElementId.InvalidElementId:
                        new_frt.ForegroundPatternId = solid_pattern_id
                    new_frt.IsMasking = False
                except Exception as ex:
                    logger.debug('Fill type props {} {}'.format(name, ex))
            return new_frt.Id
        except Exception as ex:
            logger.debug(
                'Fill type duplicate from "{}" failed: {}'.format(
                    _type_name(base), ex
                )
            )
            continue

    logger.error(
        'Fill type "{}" could not be created from any candidate — '
        'cell backgrounds will be skipped this run.'.format(name)
    )
    return DB.ElementId.InvalidElementId

_TIMING_LOG = 'C:/Users/Fred/Proton Drive/juvenciodas/My files/.Claude/Work Files/pylink_timing.log'


def _timing(row, marks):
    """Append one line of per-phase seconds for a row (profiling aid)."""
    try:
        out = ['%s=%.2f' % (marks[i + 1][0], marks[i + 1][1] - marks[i][1])
               for i in range(len(marks) - 1)]
        with open(_TIMING_LOG, 'a') as fh:
            fh.write('%s | total=%.2f | %s | %s\n' % (
                row.view_name[:40], marks[-1][1] - marks[0][1],
                ' '.join(out), row.view_type))
    except Exception:
        pass



# ── Schedule parts ──────────────────────────────────────────────────────────
# A long table splits into parts the way a drafting view does (Excel page
# breaks, N rows, or a height), the heading rows repeating on each part.
# A schedule built from the header section cannot be split by Revit, so
# the parts are built here instead: side by side in ONE wide schedule, or
# as a schedule each.

SCHED_GAP_MM = 6.0     # empty column between parts in a single wide schedule
PART_PLACE_GAP_MM = 10.0   # distance between separate schedules on a sheet


def _num_like(v):
    try:
        float(u'{}'.format(v).strip())
        return True
    except Exception:
        return False


def _schedule_part_rows(payload, scale):
    """[[source row index, ...], ...] - one list per part. Part 1 holds the
    rows from the top; every later part starts with the heading rows again.
    A single list means the table does not split."""
    all_rows = [payload['fields']] + list(payload['records'])
    n = len(all_rows)
    heights = payload.get('row_heights', {}) or {}
    dflt = payload.get('default_row_height', 14.0)
    rh = [heights.get(r, dflt) * PT_MM * scale for r in range(n)]
    rep = payload.get('repeat_rows')
    if rep is not None:
        repeat = sorted(set(int(r) for r in rep if 0 <= int(r) < n))
    else:
        first_num = n
        for ri, row in enumerate(all_rows):
            if any(_num_like(v) for v in row if v not in (None, '')):
                first_num = ri
                break
        repeat = list(range(0, max(1, first_num)))
    data0 = (max(repeat) + 1) if repeat else 1
    head = sum(rh[r] for r in repeat)
    mode = payload.get('split_mode')
    try:
        val = float(payload.get('split_value') or 0)
    except Exception:
        val = 0.0
    if mode == SPLIT_ROWS and val >= 1:
        k = int(round(val))
        breaks = list(range(data0 + k, n, k))
    elif mode == SPLIT_HEIGHT and val > 0:
        breaks = []
        acc = sum(rh[r] for r in range(data0))
        for r in range(data0, n):
            if acc + rh[r] > val and r > (breaks[-1] if breaks else data0):
                breaks.append(r)
                acc = head
            acc += rh[r]
    elif mode == SPLIT_EXCEL:
        breaks = [b for b in (payload.get('row_breaks') or []) if 0 < b < n]
    else:
        breaks = []
    edges = [0] + breaks + [n]
    parts = []
    for i in range(len(edges) - 1):
        body = list(range(edges[i], edges[i + 1]))
        if i == 0:
            parts.append(body)
        else:
            parts.append([r for r in repeat if r < edges[i]] + body)
    return parts


def _parts_payload(payload, parts, side_by_side):
    """A copy of payload rebuilt for the given parts. side_by_side puts every
    part into one wide table (a gap column between them, row i of each part
    on row i); otherwise parts[0] alone becomes the whole table."""
    all_rows = [payload['fields']] + list(payload['records'])
    n_cols = len(payload['fields'])
    cs = payload.get('cell_styles', {}) or {}
    merges = payload.get('merges', []) or []
    heights = payload.get('row_heights', {}) or {}
    cw = payload.get('col_widths', {}) or {}
    cwe = payload.get('col_widths_excel', {}) or {}
    cust_r = payload.get('custom_rows', []) or []
    cust_c = payload.get('custom_cols', []) or []
    if side_by_side:
        n_parts = len(parts)
        total_cols = n_parts * (n_cols + 1) - 1
        placed = [(rows, k * (n_cols + 1)) for k, rows in enumerate(parts)]
    else:
        total_cols = n_cols
        placed = [(parts[0], 0)]
    n_rows = max(len(rows) for rows, _ in placed)
    matrix = [[''] * total_cols for _ in range(n_rows)]
    nst, nmg, nh, ncr = {}, [], {}, set()
    for rows, off in placed:
        pos = dict((r, i) for i, r in enumerate(rows))
        for i, src in enumerate(rows):
            for c in range(n_cols):
                row = all_rows[src]
                matrix[i][off + c] = row[c] if c < len(row) else ''
                st = cs.get((src, c))
                if st is not None:
                    nst[(i, off + c)] = st
            if src in heights:
                nh[i] = max(nh.get(i, 0), heights[src])
            if src in cust_r:
                ncr.add(i)
        for (r1, c1, r2, c2) in merges:
            shown = [r for r in rows if r1 <= r <= r2]
            if not shown:
                continue
            if pos[shown[-1]] - pos[shown[0]] != len(shown) - 1:
                continue
            a, b = pos[shown[0]], pos[shown[-1]]
            if a == b and c1 == c2:
                continue
            nmg.append((a, c1 + off, b, c2 + off))
    ncw, ncwe, ncc = {}, {}, []
    for k, (rows, off) in enumerate(placed):
        for c in range(n_cols):
            if c in cw:
                ncw[off + c] = cw[c]
            if c in cwe:
                ncwe[off + c] = cwe[c]
            if c in cust_c:
                ncc.append(off + c)
    if side_by_side:
        for k in range(len(placed) - 1):
            g = k * (n_cols + 1) + n_cols
            ncw[g] = SCHED_GAP_MM
            ncwe[g] = SCHED_GAP_MM
            ncc.append(g)
            # a schedule fills its title row grey by default; the gap between
            # parts is no part of either title
            nst[(0, g)] = {'fill_rgb': (255, 255, 255), 'halign': 'Center'}
    out = dict(payload)
    out['fields'] = matrix[0]
    out['records'] = matrix[1:]
    out['cell_styles'] = nst
    out['merges'] = nmg
    out['row_heights'] = nh
    out['col_widths'] = ncw
    out['col_widths_excel'] = ncwe
    out['custom_rows'] = sorted(ncr)
    out['custom_cols'] = ncc
    out['row_breaks'] = []
    out['repeat_rows'] = None
    out['split_mode'] = SPLIT_EXCEL
    out['split_value'] = ''
    return out


# ── Group dividers as thin black columns ────────────────────────────────────
# A schedule only draws cell borders while its grid lines are on, and the
# grid then boxes every cell. So the schedule is built with the grid off, and
# each vertical rule the spreadsheet draws between column groups is made a
# very thin column filled black instead.

DIVIDER_MM = 0.25       # Excel-side width; the schedule draws it at the card scale
DIVIDER_PAD_MM = 0.8    # extra width (Excel-side) for each column beside a divider, so text clears the rule
DIVIDER_FRACTION = 0.6  # a boundary is a divider when this share of data rows rule it


def _has_border(st, side):
    v = (st or {}).get('border_' + side)
    return bool(v) and v not in ('', 'none')


def _divider_boundaries(payload):
    """Column indexes b (1..n-1) where a vertical rule sits between column
    b-1 and column b on most data rows."""
    n_cols = len(payload['fields'])
    cs = payload.get('cell_styles', {}) or {}
    n_rows = 1 + len(payload['records'])
    # heading rows are the ones above the first row holding a number
    first_num = n_rows
    all_rows = [payload['fields']] + list(payload['records'])
    for ri, row in enumerate(all_rows):
        if any(_num_like(v) for v in row if v not in (None, '')):
            first_num = ri
            break
    data_rows = [r for r in range(first_num, n_rows)]
    if not data_rows:
        return []
    out = []
    for b in range(1, n_cols):
        hits = 0
        for r in data_rows:
            if (_has_border(cs.get((r, b)), 'left')
                    or _has_border(cs.get((r, b - 1)), 'right')):
                hits += 1
        if hits >= DIVIDER_FRACTION * len(data_rows):
            out.append(b)
    return out


def _with_dividers(payload):
    """payload with a thin black column inserted at every group boundary.
    Returned unchanged when the table has none."""
    bounds = _divider_boundaries(payload)
    if not bounds:
        return payload
    n_cols = len(payload['fields'])
    # old column -> new column; divider k sits just before old column bounds[k]
    shift = {}
    for c in range(n_cols):
        shift[c] = c + len([b for b in bounds if b <= c])
    dcol = [bounds[k] + k for k in range(len(bounds))]
    new_n = n_cols + len(bounds)
    cs = payload.get('cell_styles', {}) or {}
    n_rows = 1 + len(payload['records'])

    def widen(row):
        out = [u''] * new_n
        for c, v in enumerate(list(row)[:n_cols]):
            out[shift[c]] = v
        return out

    new = dict(payload)
    new['fields'] = widen(payload['fields'])
    new['records'] = [widen(r) for r in payload['records']]
    ncs = {}
    for (r, c), st in cs.items():
        if c >= n_cols:
            continue
        st2 = st
        if c in bounds or (c + 1) in bounds:
            st2 = dict(st)
            if c in bounds:
                st2['border_left'] = ''
            if (c + 1) in bounds:
                st2['border_right'] = ''
        ncs[(r, shift[c])] = st2
    # the rule itself, on each row where the spreadsheet rules that boundary
    # (the group heading row and the data rows - not the sub-heading row,
    # which has no borders in the spreadsheet)
    for k, d in enumerate(dcol):
        bnd = bounds[k]
        for r in range(1, n_rows):
            ruled = (_has_border(cs.get((r, bnd)), 'left')
                     or _has_border(cs.get((r, bnd - 1)), 'right'))
            if ruled:
                ncs[(r, d)] = {'fill_rgb': (0, 0, 0), 'halign': 'Center'}
            else:
                # no rule on this row: the gap takes the colour of the cell it
                # sits against, so a banded or grey row has no white notch
                fill = (cs.get((r, bnd - 1)) or {}).get('fill_rgb')
                if fill:
                    ncs[(r, d)] = {'fill_rgb': tuple(fill), 'halign': 'Center'}
    new['cell_styles'] = ncs
    nm = []
    for (r1, c1, r2, c2) in payload.get('merges', []) or []:
        nm.append((r1, shift.get(c1, c1), r2, shift.get(c2, c2)))
    new['merges'] = nm
    for key in ('col_widths', 'col_widths_excel'):
        old = payload.get(key, {}) or {}
        nw = {}
        for c, w in old.items():
            if c in shift:
                nw[shift[c]] = w
        for d in dcol:
            nw[d] = DIVIDER_MM
        # room for the text beside each rule: it is centred in its column and
        # a schedule draws it wider than it measures, so it touched the line
        for b in bounds:
            for oc in (b - 1, b):
                nc = shift[oc]
                if nc in nw:
                    nw[nc] = nw[nc] + DIVIDER_PAD_MM
        new[key] = nw
    nc_set = set(shift[c] for c in (payload.get('custom_cols', []) or []) if c in shift)
    nc_set.update(dcol)
    for b in bounds:
        nc_set.add(shift[b - 1])
        nc_set.add(shift[b])
    new['custom_cols'] = sorted(nc_set)
    return new


def _find_schedule(name):
    for v in DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule):
        try:
            if v.Name == name and not v.IsTitleblockRevisionSchedule:
                return v
        except Exception:
            continue
    return None


def _schedule_instances(sched):
    out = []
    for inst in DB.FilteredElementCollector(doc).OfClass(
            DB.ScheduleSheetInstance).ToElements():
        try:
            if inst.ScheduleId == sched.Id:
                out.append(inst)
        except Exception:
            pass
    return out


def _place_beside(prev, new):
    """Put `new` on every sheet `prev` is on, just to its right and level
    with its top, unless `new` is already on that sheet."""
    gap = PART_PLACE_GAP_MM / 304.8
    on_sheets = set(_eid_int(i.OwnerViewId) for i in _schedule_instances(new))
    with revit.Transaction('pyLink - place schedule part'):
        for inst in _schedule_instances(prev):
            try:
                if _eid_int(inst.OwnerViewId) in on_sheets:
                    continue
                sheet = doc.GetElement(inst.OwnerViewId)
                bb = inst.get_BoundingBox(sheet)
                tl = (bb.Max.X + gap, bb.Max.Y)
                el = DB.ScheduleSheetInstance.Create(
                    doc, inst.OwnerViewId, new.Id, DB.XYZ(tl[0], tl[1], 0))
                doc.Regenerate()
                nb = el.get_BoundingBox(sheet)
                dx, dy = tl[0] - nb.Min.X, tl[1] - nb.Max.Y
                if abs(dx) > 1e-9 or abs(dy) > 1e-9:
                    DB.ElementTransformUtils.MoveElement(
                        doc, el.Id, DB.XYZ(dx, dy, 0))
            except Exception as ex:
                logger.warning('Could not place schedule part: {}'.format(ex))


def _remove_extra_schedules(base_name, keep):
    """Delete "<base> (k)" schedules for k > keep that pyLink made (they
    carry a pyLink tag), when a table now has fewer parts."""
    k = max(keep, 1) + 1
    misses = 0
    with revit.Transaction('pyLink - remove unused schedule parts'):
        while misses < 3 and k < 60:
            v = _find_schedule(u'{} ({})'.format(base_name, k))
            if v is None:
                misses += 1
            else:
                misses = 0
                try:
                    if get_view_pylink_data(v):
                        doc.Delete(v.Id)
                except Exception as ex:
                    logger.debug('remove extra schedule: {}'.format(ex))
            k += 1


def _reapply_schedule_borders(view_id, borders):
    """Write a schedule's cell borders again in a transaction of their own,
    after the build has committed. Borders set inside the transaction that
    built the table show for a few seconds and are then lost as Revit
    settles the table; written afterwards, from fresh section data, they
    stay (the order pyTransmit uses)."""
    if not borders or view_id is None:
        return
    try:
        sched = doc.GetElement(DB.ElementId(_Int64(view_id)))
        with revit.Transaction('pyLink - schedule borders'):
            hdr = sched.GetTableData().GetSectionData(DB.SectionType.Header)
            for (r, c), sides in borders.items():
                try:
                    st = hdr.GetTableCellStyle(r, c)
                    opts = st.GetCellStyleOverrideOptions()
                    opts.BorderLineStyle = True
                    for prop, eid in sides.items():
                        setattr(opts, prop, True)
                        setattr(st, prop, eid)
                    st.SetCellStyleOverrideOptions(opts)
                    hdr.SetCellStyle(r, c, st)
                except Exception as ex:
                    logger.debug('border ({},{}): {}'.format(r, c, ex))
    except Exception as ex:
        logger.warning('Schedule borders pass failed: {}'.format(ex))


def _whiten(payload):
    """Colour only where the spreadsheet has colour: a cell with no fill is
    WHITE, never left to a schedule's own default (it fills the title row,
    among others, grey)."""
    cs = dict(payload.get('cell_styles', {}) or {})
    n_rows = 1 + len(payload['records'])
    n_cols = len(payload['fields'])
    for r in range(n_rows):
        for c in range(n_cols):
            st = cs.get((r, c))
            if st is None:
                cs[(r, c)] = {'fill_rgb': (255, 255, 255)}
            elif not st.get('fill_rgb'):
                st = dict(st)
                st['fill_rgb'] = (255, 255, 255)
                cs[(r, c)] = st
    out = dict(payload)
    out['cell_styles'] = cs
    return out


def _export_schedule(export_script, payload):
    """_run_export_script for a schedule, then its borders in their own
    transaction. Returns the export result."""
    payload = _whiten(payload)
    res = _run_export_script(export_script, payload)
    if res and res.get('borders'):
        _reapply_schedule_borders(res.get('view_id'), res.get('borders'))
    return res


def _run_schedule_build(export_script, payload, row):
    """Run the schedule export for a row, split into parts when the row
    says so. Returns the export result of the first (or only) schedule."""
    # A schedule fills any title-row cell the spreadsheet left without a
    # fill grey. Make them white, so a column the title does not cover
    # does not show as a grey block.
    _cs = dict(payload.get('cell_styles', {}) or {})
    for _c in range(len(payload['fields'])):
        _st = _cs.get((0, _c))
        if _st is None:
            _cs[(0, _c)] = {'fill_rgb': (255, 255, 255)}
        elif not _st.get('fill_rgb'):
            _st = dict(_st)
            _st['fill_rgb'] = (255, 255, 255)
            _cs[(0, _c)] = _st
    payload = dict(payload)
    payload['cell_styles'] = _cs
    payload = _with_dividers(payload)
    scale = getattr(row, 'native_scale', DEFAULT_NATIVE_SCALE)
    parts = _schedule_part_rows(payload, scale)
    name = payload['view_name']
    if len(parts) <= 1:
        res = _export_schedule(export_script, payload)
        _remove_extra_schedules(name, 1)
        return res
    mode = getattr(row, 'sched_mode', SCHED_ONE)
    if mode != SCHED_PER_PART:
        res = _export_schedule(export_script,
                                 _parts_payload(payload, parts, True))
        _remove_extra_schedules(name, 1)
        return res
    first, prev = None, None
    for k, rows in enumerate(parts):
        p = _parts_payload(payload, [rows], False)
        p['view_name'] = name if k == 0 else u'{} ({})'.format(name, k + 1)
        res = _export_schedule(export_script, p)
        sched = _find_schedule(p['view_name'])
        if k == 0:
            first = res
        elif sched is not None:
            try:
                set_view_pylink_data(
                    sched, FP=row.file_path, SH=row.sheet_name,
                    RG=row.named_range, VT=row.view_type,
                    H=_hash_range(row.file_path, row.named_range,
                                  row.sheet_name))
            except Exception as ex:
                logger.debug('part tag failed: {}'.format(ex))
        if prev is not None and sched is not None:
            _place_beside(prev, sched)
        prev = sched
    _remove_extra_schedules(name, len(parts))
    return first


def apply_row(row):
    """
    Process one UI row:
    1. Read data from the Excel named range
    2. Pre-create pyLink text types (before any transaction)
    3. Call the appropriate Export/ script via exec()

    Returns {'view_name', 'status', 'message'}
    """
    result = {
        'view_name': row.view_name,
        'status':    'error',
        'message':   ''
    }

    _marks = [('start', _time.time())]
    logger.debug('pyLink: {} | {} | {}/{}'.format(
        row.view_name, row.file_path, row.sheet_name, row.named_range))

    # Read data from xlsx
    rows = read_named_range_data(
        row.file_path,
        row.named_range,
        row.sheet_name
    )

    if not rows:
        result['message'] = (
            'No data found in named range "{}". '
            'Check the range name and sheet.'.format(row.named_range)
        )
        return result

    _marks.append(('read_data', _time.time()))
    fields  = [str(h) for h in rows[0]]
    records = [[str(c) for c in r] for r in rows[1:]]

    if not fields:
        result['message'] = 'Named range has no header row.'
        return result

    # Ownership guard — never silently overwrite a view pyLink itself
    # didn't create. ElementId is the authoritative proof (survives the
    # view being renamed outside pyLink); the old name-match proof is
    # only a fallback for a row that hasn't recorded an id yet (first
    # apply under this version, or state saved before this existed).
    # A row stacking onto a shared legend/drafting view is drawing on a
    # view the FIRST row of its run owns, by design, and it appends
    # rather than clearing - so there is nothing here to protect. The
    # guard still runs for that first row, which is the one that clears.
    _stacking = getattr(row, 'stack_index', 0) > 0
    try:
        existing_view = None
        # Fast path: the view this row made last time, found by id. Only a
        # name hit counts, so a renamed or deleted view falls through to
        # the full scan below, which looks at every view by name.
        _aid = getattr(row, 'applied_view_id', None)
        if _aid is not None and not _stacking:
            try:
                _cand = doc.GetElement(DB.ElementId(_Int64(_aid)))
                if _cand is not None and isinstance(_cand, DB.View):
                    # Smart link: the id is the link, not the name. If the
                    # view was renamed in Revit, follow the rename and keep
                    # building into that view instead of making a new one.
                    existing_view = _cand
                    if _cand.Name != row.view_name:
                        if row.view_name == getattr(row, 'applied_view_name', None):
                            # pyLink's name is unchanged since the last
                            # apply, so the rename happened in Revit:
                            # follow it.
                            row.view_name = _cand.Name
                            result['view_name'] = _cand.Name
                        else:
                            # The name was changed in pyLink: push it to
                            # the view (the build then reuses that name).
                            _taken = any(
                                _o.Name == row.view_name
                                and _o.Id != _cand.Id
                                for _o in DB.FilteredElementCollector(doc)
                                .OfClass(DB.View))
                            if not _taken:
                                with revit.Transaction(
                                        'pyLink - rename view'):
                                    _cand.Name = row.view_name
            except Exception:
                existing_view = None
        for v in ([] if (_stacking or existing_view is not None) else
                  DB.FilteredElementCollector(revit.doc).OfClass(DB.View)):
            try:
                if v.IsValidObject and v.Name == row.view_name:
                    existing_view = v
                    break
            except Exception:
                continue
        if existing_view is not None:
            owns_by_id = (
                getattr(row, 'applied_view_id', None) is not None
                and _eid_int(existing_view.Id) == row.applied_view_id
            )
            owns_by_name_fallback = (
                getattr(row, 'applied_view_id', None) is None
                and row.view_name == row.applied_view_name
            )
            if not (owns_by_id or owns_by_name_fallback):
                result['message'] = (
                    'A view named "{}" already exists and was not '
                    'created by this pyLink row - refusing to '
                    'overwrite it. Rename either the existing '
                    'view or this row.'.format(row.view_name)
                )
                return result
    except Exception as ex:
        logger.warning(
            'Ownership check failed, proceeding cautiously: {}'.format(ex)
        )

    # Read cell formatting from xlsx first - text type creation below
    # needs it to know each cell's actual color.
    fmt = read_range_formatting(
        row.file_path,
        row.named_range,
        row.sheet_name
    )
    cell_styles = fmt.get('cell_styles', {})
    _marks.append(('ownership+fmt', _time.time()))

    # Take text size from Excel's own font_size in cell_styles, not the fixed
    # 2.5mm/2.3mm TableRow defaults, which are unrelated to the file - a 9pt
    # Excel header used to land at ~7pt because row.size_hdr_mm never reflected
    # the source. Row (0,*) is the header, (1,*) the first data row; use the
    # first cell in each that states a size, falling back to the row default
    # only if Excel genuinely has none (rare - every real xlsx cell has one).
    def _first_font_size(row_idx, n_cols):
        for c in range(n_cols):
            fs = cell_styles.get((row_idx, c), {}).get('font_size')
            if fs:
                return float(fs)
        return None

    hdr_pt = _first_font_size(0, len(fields))
    dat_pt = _first_font_size(1, len(fields))
    size_hdr_mm = (hdr_pt * PT_MM) if hdr_pt else row.size_hdr_mm
    size_dat_mm = (dat_pt * PT_MM) if dat_pt else row.size_dat_mm

    # Card scale: an override from the card wins. Otherwise the scale is
    # worked out so the table's BODY text (the card's pt override, else the
    # size most rows use, else the first data row) draws at the Text Map
    # target, so every card ends up with the same text size on the sheet.
    if getattr(row, 'native_scale', None) is None:
        _body_pt = (getattr(row, 'body_pt', None)
                    or detect_body_pt(cell_styles) or dat_pt or 10.0)
        try:
            _target = float(load_excel_font_settings().get('target_mm', 2.0))
        except Exception:
            _target = 2.0
        row.native_scale = auto_native_scale(_body_pt, _target)
        logger.debug('pyLink auto scale: body {}pt -> {:.3f}'.format(
            _body_pt, row.native_scale))

    # Every view type is built at Excel's own sizes times the card's scale
    # (see NATIVE_SCALE_OPTIONS), so the same card comes out the same size
    # as a schedule, a drafting view or a legend. Schedules work their
    # per-row heights out in create_schedule.py; a drafting view and a
    # legend carry two text types, header and data, set here at the size
    # a schedule cell really draws (native_text_type_mm). The old
    # per-project "Schedule Text Size" answers are no longer asked for.
    row_size_mm = {}
    _scale = getattr(row, 'native_scale', DEFAULT_NATIVE_SCALE)
    if row.view_type != VIEW_TYPE_SCHEDULE:
        size_hdr_mm = native_text_type_mm(hdr_pt or 10.0, _scale)
        size_dat_mm = native_text_type_mm(dat_pt or 10.0, _scale)

    # Same idea for font NAME - use Excel's own font (e.g. 'Aptos
    # Narrow') if it's actually installed on this machine, otherwise
    # fall back to the user-configured default (hamburger menu ->
    # Default Font) rather than handing Revit a font name it doesn't
    # have and letting it silently substitute something unpredictable.
    def _first_font_name(row_idx, n_cols):
        for c in range(n_cols):
            fn = cell_styles.get((row_idx, c), {}).get('font_name')
            if fn:
                return fn
        return None

    font_settings = load_excel_font_settings()
    if font_settings.get('force_fallback'):
        # Toggle is on - always the configured default, don't even
        # look at what Excel says.
        font_name = font_settings.get('fallback_font', 'Arial')
    else:
        excel_font = _first_font_name(0, len(fields)) or _first_font_name(1, len(fields))
        if excel_font and _is_font_installed(excel_font):
            font_name = excel_font
        else:
            font_name = font_settings.get('fallback_font', 'Arial')

    # Pre-create text types outside any transaction. Size only, always
    # plain/black — per-cell text colour and bold/italic/underline are
    # applied afterwards as instance-level formatting (view-specific
    # graphic override for colour, FormattedText for bold/italic/
    # underline) rather than by creating a new Type per combination.
    #
    # A schedule draws every cell at its own text type's size until Revit
    # is made to redraw it, whatever size the cell itself was given - so
    # its types are made at the size the scaled cells really draw at.
    tt_hdr_mm, tt_dat_mm = size_hdr_mm, size_dat_mm
    if row.view_type == VIEW_TYPE_SCHEDULE:
        tt_hdr_mm = native_text_type_mm(hdr_pt or 10.0, _scale)
        tt_dat_mm = native_text_type_mm(dat_pt or 10.0, _scale)
    hdr_tt = get_or_create_text_type(
        font=font_name, size_mm=tt_hdr_mm
    )
    dat_tt = get_or_create_text_type(
        font=font_name, size_mm=tt_dat_mm
    )
    hdr_tt_id = hdr_tt.Id if hdr_tt else DB.ElementId.InvalidElementId
    dat_tt_id = dat_tt.Id if dat_tt else DB.ElementId.InvalidElementId

    # Pre-create types outside any transaction - Revit needs new
    # types/subcategories in their own transaction, separate from the
    # drafting/legend/schedule one. line_ids feeds create_schedule.py's
    # separate cell-style API. Drafting/legend fill uses one reusable type
    # coloured per instance via view overrides, not a Type per colour.
    line_ids     = _pre_create_line_styles(cell_styles)
    fill_type_id = _get_or_create_pylink_fill_type()

    _marks.append(('types+styles', _time.time()))
    # Build payload for export script
    payload = {
        'view_name':          row.view_name,
        'fields':             fields,
        'records':            records,
        'font':               font_name,
        'size_hdr_mm':        size_hdr_mm,
        'size_dat_mm':        size_dat_mm,
        # Already resolved, so create_schedule.py does not ask again.
        'row_size_mm':        row_size_mm,
        'hdr_tt_id':          hdr_tt_id,
        'dat_tt_id':          dat_tt_id,
        'view_scale':         row.view_scale,
        'native_scale':       getattr(row, 'native_scale', DEFAULT_NATIVE_SCALE),
        'cell_styles':        cell_styles,
        'merges':             fmt.get('merges', []),
        'row_heights':        fmt.get('row_heights', {}),
        'col_widths':         fmt.get('col_widths', {}),
        # Excel's own column geometry, used only to work out where the
        # spreadsheet wrapped its text. Absent for .ods, which falls
        # back to col_widths.
        'col_widths_excel':   fmt.get('col_widths_excel', {}),
        'stack_index':        getattr(row, 'stack_index', 0),
        # Excel's manual page breaks and Print Titles: a drafting view or
        # legend draws each part beside the last, titles repeated.
        'row_breaks':         fmt.get('row_breaks', []),
        'repeat_rows':        fmt.get('repeat_rows'),
        # The row's own choice: Excel (the breaks above), Rows or Height.
        'split_mode':         getattr(row, 'split_mode', SPLIT_EXCEL),
        'split_value':        getattr(row, 'split_value', ''),
        # Which of those sizes the user actually chose, as opposed to
        # Excel's defaults - the schedule builder autofits the rest.
        'custom_rows':        fmt.get('custom_rows', []),
        'custom_cols':        fmt.get('custom_cols', []),
        'default_row_height': fmt.get('default_row_height', 14.0),
        'line_ids':           line_ids,
        'fill_type_id':       fill_type_id,
    }

    # Map view type to export script
    script_map = {
        VIEW_TYPE_SCHEDULE: 'create_schedule.py',
        VIEW_TYPE_DRAFTING: 'create_drafting.py',
        VIEW_TYPE_LEGEND:   'create_legend.py',
    }

    export_script = script_map.get(row.view_type)
    if not export_script:
        result['message'] = 'Unknown view type: {}'.format(row.view_type)
        return result

    # View type changed (Drafting <-> Legend <-> Schedule): the old view is
    # captured (sheets, top-left corners, Title on Sheet) and deleted so the
    # new one can take its name, then placed where the old one was.
    _swap = None
    if (not _stacking and existing_view is not None
            and not view_matches_type(existing_view, row.view_type)):
        try:
            _swap = capture_view_for_swap(existing_view)
            with revit.Transaction('pyLink - replace view: {}'.format(
                    row.view_name)):
                doc.Delete(existing_view.Id)
            logger.debug('View type changed - replaced "{}" ({} sheet(s))'
                         .format(row.view_name, len(_swap.get('sheets', []))))
        except Exception as ex:
            result['message'] = 'Could not replace the old view: {}'.format(ex)
            return result

    try:
        if row.view_type == VIEW_TYPE_SCHEDULE and not _stacking:
            export_result = _run_schedule_build(export_script, payload, row)
        else:
            export_result = _run_export_script(export_script, payload)
        _marks.append(('export_script', _time.time()))
        result['status']  = 'success'
        result['message'] = 'Created'
        view_id = None
        if export_result and export_result.get('view_id') is not None:
            view_id = export_result['view_id']
            result['view_id'] = view_id
            try:
                # Same explicit-Int64 disambiguation as the reload path
                # in pyLink.py - a plain int is ambiguous against
                # ElementId's BuiltInParameter/BuiltInCategory/Int64
                # overloads under IronPython (this exact "Multiple
                # targets could match" error was seen in Fred's log).
                view = doc.GetElement(DB.ElementId(_Int64(view_id)))
                if view and _swap is not None:
                    try:
                        restore_view_after_swap(view, _swap)
                    except Exception as ex:
                        logger.warning('Placing replaced view failed: {}'
                                       .format(ex))
                if view:
                    set_view_pylink_data(
                        view,
                        FP=row.file_path,
                        SH=row.sheet_name,
                        RG=row.named_range,
                        VT=row.view_type,
                        H=_hash_range(row.file_path, row.named_range,
                                      row.sheet_name),
                    )
            except Exception as ex:
                logger.debug('View tag failed: {}'.format(ex))
    except Exception as e:
        import traceback
        result['message'] = str(e)
        logger.error(traceback.format_exc())

    _marks.append(('tag+hash', _time.time()))
    _timing(row, _marks)
    return result


# ── Entry point ──
# ── Word document reading ──

def _hash_range(file_path, named_range, sheet_name):
    """
    Compute a quick hash of the named range content + formatting.
    Used to detect changes between applies.

    Both the outer cell_styles dict AND every inner per-cell style dict
    must be sorted before repr(), plain repr(dict) is not safe here:
    dict key iteration order can differ between separate Revit sessions
    (hash randomisation), which would otherwise make every row show as
    "changed" even when the source file was never touched.
    """
    import hashlib
    # Same file (same modified time and size) gives the same hash, so the
    # range is read and hashed once, not on every status refresh.
    _hk = (file_path, named_range, sheet_name)
    _stamp = _file_stamp(file_path)
    _hit = _HASH_CACHE.get(_hk)
    if _stamp is not None and _hit is not None and _hit[0] == _stamp:
        return _hit[1]
    try:
        rows = read_named_range_data(file_path, named_range, sheet_name)
        fmt  = read_range_formatting(file_path, named_range, sheet_name)
        # Hash cell values
        content = repr(rows)
        # Hash cell styles (fills, colours, borders) — sort both the
        # outer dict (by cell coordinate) and every inner style dict
        # (by property name) so the string is identical across sessions
        # for identical data.
        cell_styles = fmt.get('cell_styles', {})
        styles_sorted = sorted(
            (coord, sorted(style.items()))
            for coord, style in cell_styles.items()
        )
        styles = repr(styles_sorted)
        combined = content + styles
        # Page breaks change the drawing too. Only added when there are
        # any, so tables without breaks keep the hash they already have.
        if fmt.get('row_breaks'):
            combined += repr((fmt.get('row_breaks'),
                              fmt.get('repeat_rows')))
        _h = hashlib.md5(combined.encode('utf-8', errors='replace')).hexdigest()
        if _stamp is not None:
            _HASH_CACHE[_hk] = (_stamp, _h)
        return _h
    except Exception:
        return None


class ExcelCardMixin(object):
    """Excel-specific card/row UI methods, mixed into
    PyLinkWindow. Everything here assumes fd.get('source_type') == 'xl'."""

    # ── Column-header dropdowns (batch set on the card's ticked rows) ──

    def _header_button(self, text, width, path, kind):
        """A column heading that opens a dropdown setting that column on
        every ticked row of the card."""
        btn = Button()
        btn.Content = u'{} \u25BE'.format(text)
        btn.Width = width
        btn.FontSize = 10
        btn.Foreground = hb('#F4FAFF')
        btn.Opacity = 0.75
        btn.Background = Brushes.Transparent
        btn.BorderThickness = Thickness(0)
        btn.Padding = Thickness(4, 0, 0, 0)
        # Plain template: the default button chrome is what drew the black
        # bar behind these headings. This one is just the text, on the card.
        try:
            from System.Windows.Markup import XamlReader
            btn.Template = XamlReader.Parse(
                '<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/'
                '2006/xaml/presentation" TargetType="Button">'
                '<Border Background="Transparent" Padding="{TemplateBinding Padding}">'
                '<ContentPresenter HorizontalAlignment="{TemplateBinding '
                'HorizontalContentAlignment}" VerticalAlignment="Center"/>'
                '</Border></ControlTemplate>')
        except Exception:
            pass
        btn.HorizontalContentAlignment = HorizontalAlignment.Left
        btn.VerticalAlignment = VerticalAlignment.Center
        btn.Cursor = Cursors.Hand
        btn.FocusVisualStyle = None
        btn.ToolTip = 'Set {} for every ticked row in this card'.format(text)
        btn.Tag = (path, kind)
        btn.Click += self._header_menu
        return btn

    def _header_menu(self, sender, e):
        path, kind = sender.Tag
        fd = self._file_data.get(path)
        if fd is None:
            return
        pops = fd.setdefault('hdr_popups', {})
        if kind not in pops:
            pops[kind] = self._build_dropdown_popup(sender)
        popup, panel = pops[kind]
        if popup.IsOpen:
            popup.IsOpen = False
            return
        panel.Children.Clear()

        def item(label, fn):
            return self._make_menu_item(label, fn, popup)

        def ticked():
            return self._ticked(lambda: list(fd.get('rows', [])))

        if kind == 'vn':
            for lab in VIEW_NAME_MODES:
                panel.Children.Add(item(
                    lab, lambda s, ev, v=lab: self._batch_view_name(ticked(), v)))
        elif kind == 'vt':
            for vt in VIEW_TYPES:
                panel.Children.Add(item(
                    vt, lambda s, ev, v=vt: self._batch_view_type(ticked(), v)))
        elif kind == 'sm':
            for m in SCHED_MODES:
                panel.Children.Add(item(
                    m, lambda s, ev, v=m: self._batch_sched_mode(ticked(), v)))
            panel.Children.Add(self._make_menu_separator())
            panel.Children.Add(item(
                u'Scale\u2026', lambda s, ev: self._batch_scale(ticked())))
        elif kind == 'sp':
            for m in SPLIT_MODES:
                panel.Children.Add(item(
                    m, lambda s, ev, v=m: self._batch_split_mode(ticked(), v)))
        elif kind == 'rv':
            panel.Children.Add(item(
                u'Set Rows / mm\u2026',
                lambda s, ev: self._batch_split_value(ticked())))
        popup.IsOpen = True

    def _mark_row_changed(self, row):
        """A setting that changes what a row builds was edited: show it as
        needing an update (blue), tick it, so Apply picks it up."""
        if row.Status == 'success':
            row.Status = 'sync'
            try:
                if row._dot is not None:
                    row._dot.Fill = hb(STATUS_COLOURS.get('sync', '#3B82F6'))
                self._set_sync_btn_state(row._refresh_btn, True)
            except Exception:
                pass
        self._auto_check_row(row)

    def _sync_scale_mode_widgets(self, row):
        """Schedule rows show the mode dropdown in the Scale column;
        Drafting and Legend rows show the scale box."""
        is_sched = (row.ViewType == VIEW_TYPE_SCHEDULE)
        tb = getattr(row, '_scale_tb', None)
        cb = getattr(row, '_sched_cb', None)
        if tb is not None:
            tb.Visibility = Visibility.Collapsed if is_sched else Visibility.Visible
        if cb is not None:
            cb.Visibility = Visibility.Visible if is_sched else Visibility.Collapsed

    def _sched_mode_row_changed(self, sender, e):
        """Row dropdown -> how this schedule's parts are built. The row has
        to be rebuilt, so it shows as changed."""
        row = sender.Tag
        if not row or sender.SelectedItem is None:
            return
        new = str(sender.SelectedItem)
        if new == getattr(row, 'SchedMode', SCHED_ONE):
            return
        row.SchedMode = new
        if row.Status == 'success':
            row.Status = 'sync'
            try:
                if row._dot is not None:
                    row._dot.Fill = hb(STATUS_COLOURS.get('sync', '#3B82F6'))
                self._set_sync_btn_state(row._refresh_btn, True)
                self._refresh_status_card()
            except Exception:
                pass
        self._auto_check_row(row)
        self._save_persisted_state()

    def _sched_mode_changed(self, sender, e):
        """Card dropdown -> how split schedules are built. Schedule rows
        that were up to date now need rebuilding, so they show as changed."""
        fd = self._file_data.get(sender.Tag)
        if fd is None or sender.SelectedItem is None:
            return
        new = str(sender.SelectedItem)
        if new == card_sched_mode(fd):
            return
        fd['sched_mode'] = new
        for r in fd.get('rows', []):
            if r.ViewType == VIEW_TYPE_SCHEDULE and r.Status == 'success':
                r.Status = 'sync'
                try:
                    if r._dot is not None:
                        r._dot.Fill = hb(STATUS_COLOURS.get('sync', '#3B82F6'))
                    self._set_sync_btn_state(r._refresh_btn, True)
                except Exception:
                    pass
        try:
            self._refresh_status_card()
        except Exception:
            pass
        self._save_persisted_state()

    def _native_scale_changed(self, sender, e):
        """Card scale dropdown -> remember the new scale for this card
        ('None' = automatic from the Text Map target)."""
        fd = self._file_data.get(sender.Tag)
        if fd is None or sender.SelectedItem is None:
            return
        text = str(sender.SelectedItem)
        try:
            fd['native_scale'] = None if text == 'None' else float(text)
        except Exception:
            return
        self._save_persisted_state()

    def _body_pt_changed(self, sender, e):
        """Card body-text dropdown -> the Excel text size this card's
        scale is matched to ('None' = detect it)."""
        fd = self._file_data.get(sender.Tag)
        if fd is None or sender.SelectedItem is None:
            return
        text = str(sender.SelectedItem).replace('pt', '').strip()
        try:
            fd['body_pt'] = None if text == 'None' else float(text)
        except Exception:
            return
        self._save_persisted_state()

    def _make_excel_card(self, path):
        """Build the flat, single-card layout Excel cards use --
        one fd per card, no grouping/multi-view nesting like Word.
        Called from the shared _make_card dispatcher."""
        fd = self._file_data[path]
        outer = Border()
        try:
            outer.Style = self.FindResource('CardStyle')
        except Exception as e:
            logger.warning('Failed to apply CardStyle: {}'.format(e))

        self._slim_card(outer)
        inner = StackPanel()
        inner.Orientation = Orientation.Vertical

        # Card header row — Grid so the right-hand group (reload,
        # close) stays flush right regardless of path length. Left
        # group: collapse toggle, source badge, path, last-modified.
        header_row = Grid()
        header_row.Margin = Thickness(0, 0, 0, 6)
        _hcol_l = ColumnDefinition()
        _hcol_l.Width = GridLength(1, GridUnitType.Star)
        _hcol_r = ColumnDefinition()
        _hcol_r.Width = GridLength.Auto
        header_row.ColumnDefinitions.Add(_hcol_l)
        header_row.ColumnDefinitions.Add(_hcol_r)

        header_left = StackPanel()
        header_left.Orientation = Orientation.Horizontal
        Grid.SetColumn(header_left, 0)
        header_row.Children.Add(header_left)

        header_right = StackPanel()
        header_right.Orientation = Orientation.Horizontal
        header_right.HorizontalAlignment = HorizontalAlignment.Right
        Grid.SetColumn(header_right, 1)
        header_row.Children.Add(header_right)

        # Collapse/expand toggle — chevron-circle vector icon from the
        # shared Seed43 icon lib. The icon's own circle outline is the
        # chrome, so the button itself stays transparent; only the
        # icon's fill colour changes (grey while expanded, green once
        # collapsed) to match the rest of the card-state language.
        collapse_btn = ToggleButton()
        collapse_btn.IsChecked = True
        try:
            collapse_btn.Style = self.FindResource('PrimarySecondaryToggleButtonStyle')
        except Exception as e:
            logger.warning('Failed to apply PrimarySecondaryToggleButtonStyle: {}'.format(e))
        collapse_btn.FocusVisualStyle = None
        self._flatten_arrow_button(collapse_btn)
        collapse_btn.HorizontalContentAlignment = HorizontalAlignment.Center
        collapse_btn.VerticalContentAlignment   = VerticalAlignment.Center
        collapse_btn.VerticalAlignment          = VerticalAlignment.Center
        collapse_btn.Margin          = Thickness(0, 0, 8, 0)
        collapse_btn.Tag     = path
        collapse_btn.ToolTip = 'Collapse'
        collapse_btn.Click  += self._toggle_card_collapse
        # Kept so the card can be collapsed from code as well as by
        # clicking - restoring a saved collapsed state needs the button
        # driven too, or its chevron points the wrong way.
        self._file_data.setdefault(path, {})['collapse_btn'] = collapse_btn
        self._set_collapse_icon(collapse_btn, False)
        header_left.Children.Add(collapse_btn)

        # Source-type badge (W / XL / ODS / ODT) - same look as the per-row
        # badge, slightly bigger for card-header scale. ODS shares the xlsx
        # code path (source_type='xl'), so the real extension, not
        # source_type, picks the label/colour; otherwise every LibreOffice
        # Calc file would show as a plain Excel badge.
        real_ext = os.path.splitext(fd.get('real_path', path) or '')[1].lower()
        if real_ext == '.ods':
            badge_key, badge_text = 'ods', 'ODS'
        elif fd.get('source_type') == 'word':
            badge_key, badge_text = 'word', 'W'
        else:
            badge_key, badge_text = 'xl', 'XL'
        src_badge = Border()
        src_badge.Width             = 24
        src_badge.Height            = 24
        src_badge.CornerRadius      = CornerRadius(4)
        src_badge.Background        = hb(SRC_COLOURS.get(badge_key, '#555'))
        src_badge.Margin            = Thickness(0, 0, 8, 0)
        src_badge.VerticalAlignment = VerticalAlignment.Center
        src_lbl = TextBlock()
        src_lbl.Text                = badge_text
        src_lbl.FontSize            = 8 if len(badge_text) > 2 else 9
        src_lbl.FontWeight          = FontWeights.Bold
        src_lbl.Foreground          = hb('#FFFFFF')
        src_lbl.HorizontalAlignment = HorizontalAlignment.Center
        src_lbl.VerticalAlignment   = VerticalAlignment.Center
        src_badge.Child = src_lbl
        header_left.Children.Add(src_badge)

        # Two lines: the file name in white with its modified time beside
        # it, and the path underneath in grey.
        title_stack = StackPanel()
        title_stack.Orientation = Orientation.Vertical
        title_stack.VerticalAlignment = VerticalAlignment.Center
        header_left.Children.Add(title_stack)
        title_line = StackPanel()
        title_line.Orientation = Orientation.Horizontal
        title_stack.Children.Add(title_line)

        name_lbl = TextBlock()
        name_lbl.Text         = os.path.basename(fd.get('real_path', path) or path)
        name_lbl.FontSize     = 14
        name_lbl.FontWeight   = FontWeights.SemiBold
        name_lbl.Foreground   = hb('#FFFFFF')
        name_lbl.VerticalAlignment = VerticalAlignment.Bottom
        title_line.Children.Add(name_lbl)
        fd['name_label'] = name_lbl

        heading = TextBlock()
        heading.Text         = path
        heading.TextTrimming = TextTrimming.CharacterEllipsis
        heading.ToolTip      = path
        heading.Foreground   = hb('#6B7280')
        heading.FontWeight   = FontWeights.SemiBold
        heading.FontSize     = 11
        title_stack.Children.Add(heading)
        fd['heading_label']  = heading

        lm_text = TextBlock()
        try:
            dt = DateTime.FromFileTime(
                int(os.path.getmtime(path) * 10000000) + 116444736000000000)
            lm_text.Text = dt.ToString('dd/MM/yyyy HH:mm')
        except Exception:
            lm_text.Text = ''
        lm_text.FontSize         = 10
        lm_text.Foreground       = hb('#F4FAFF')
        lm_text.Opacity          = 0.55
        lm_text.VerticalAlignment = VerticalAlignment.Bottom
        lm_text.Margin           = Thickness(8, 0, 0, 2)
        title_line.Children.Add(lm_text)

        # Card scale - every schedule, drafting view and legend this card
        # builds is drawn at Excel's own text, column and row sizes times
        # this, so it keeps the sheet's proportions. Takes effect on the
        # next Apply.
        ov_lbl = TextBlock()
        ov_lbl.Text = u'Override'
        ov_lbl.FontSize = 10
        ov_lbl.Foreground = hb('#F4FAFF')
        ov_lbl.Opacity = 0.55
        ov_lbl.VerticalAlignment = VerticalAlignment.Center
        ov_lbl.Margin = Thickness(8, 0, 0, 0)
        ov_lbl.ToolTip = (u'None = automatic: the scale is worked out so the '
                          u'body text draws at the Text Map size. Pick a '
                          u'value to override it.')
        header_right.Children.Add(ov_lbl)

        scale_combo = ComboBox()
        self._combo_style(scale_combo, 72)
        scale_combo.Margin = Thickness(6, 0, 0, 0)
        scale_combo.ToolTip = ('Scale override: Excel text, column and row '
                               'sizes multiplied by this. None = automatic.')
        scale_combo.Items.Add(u'None')
        for opt in NATIVE_SCALE_OPTIONS:
            scale_combo.Items.Add(u'{:g}'.format(opt))
        _sc = card_native_scale(fd)
        scale_combo.SelectedItem = u'None' if _sc is None else u'{:g}'.format(_sc)
        fd['native_scale'] = _sc
        scale_combo.Tag = path
        scale_combo.SelectionChanged += self._native_scale_changed
        header_right.Children.Add(scale_combo)

        pt_combo = ComboBox()
        self._combo_style(pt_combo, 78)
        pt_combo.Margin = Thickness(4, 0, 0, 0)
        pt_combo.ToolTip = ('Body text override: the Excel font size (pt) '
                            'the automatic scale matches. None = detect it '
                            'from the table. Set it when a table has mixed '
                            'sizes and the wrong one is picked.')
        pt_combo.Items.Add(u'None')
        for opt in BODY_PT_OPTIONS:
            pt_combo.Items.Add(u'{:g} pt'.format(opt))
        _bp = card_body_pt(fd)
        pt_combo.SelectedItem = u'None' if _bp is None else u'{:g} pt'.format(_bp)
        fd['body_pt'] = _bp
        pt_combo.Tag = path
        pt_combo.SelectionChanged += self._body_pt_changed
        header_right.Children.Add(pt_combo)

        # Quick Add — same action as the control-row Add button below,
        # placed here too for one-click access without scrolling past
        # a long row list to reach it.
        header_add_btn = self._green_btn(u'+ Add Row')
        header_add_btn.Tag    = path
        header_add_btn.Click += self._add_row_for_card
        header_add_btn.VerticalAlignment = VerticalAlignment.Center
        header_add_btn.Margin = Thickness(6, 0, 0, 0)
        header_right.Children.Add(header_add_btn)

        # Per-card Batch menu — Delete selected / Duplicate / path mode /
        # Open File / Open Folder / Unlink View / Remove view. Separate
        # from the global toolbar Batch menu, which operates across all
        # cards (Enabled toggling, bulk view-type set).
        batch_btn = Button()
        batch_btn.Content    = u'Batch'
        try:
            batch_btn.Style = self.FindResource('DropdownButtonStyle')
        except Exception as e:
            logger.warning('Failed to apply DropdownButtonStyle: {}'.format(e))
        batch_btn.FocusVisualStyle = None
        batch_btn.Width       = 90
        batch_btn.VerticalAlignment = VerticalAlignment.Center
        batch_btn.Margin      = Thickness(6, 0, 0, 0)
        batch_btn.Tag         = path
        batch_btn.Click      += self._card_batch_menu
        header_right.Children.Add(batch_btn)

        # Per-card reload — grey when everything's in sync, blue when any
        # row in this card needs reapplying, matching each row's own
        # sync-dot state aggregated up to the card.
        reload_btn = Button()
        if _mi is not None:
            try:
                reload_btn.Content = _mi('reload', size=14, color='#FFFFFF')
            except Exception:
                reload_btn.Content = u'\u21bb'
        else:
            reload_btn.Content = u'\u21bb'
        try:
            reload_btn.Style = self.FindResource('RoundPrimaryButtonStyle')
        except Exception as e:
            logger.warning('Failed to apply RoundPrimaryButtonStyle: {}'.format(e))
        reload_btn.FocusVisualStyle = None
        self._style_card_reload_btn(reload_btn)
        reload_btn.HorizontalContentAlignment = HorizontalAlignment.Center
        reload_btn.VerticalContentAlignment   = VerticalAlignment.Center
        reload_btn.VerticalAlignment          = VerticalAlignment.Center
        reload_btn.Margin  = Thickness(6, 0, 0, 0)
        reload_btn.Tag     = path
        reload_btn.ToolTip = 'Reload rows that need updating'
        reload_btn.Click  += self._card_reload_click
        header_right.Children.Add(reload_btn)
        fd['reload_btn'] = reload_btn

        # Close card — always-red filled circle (more prominent than
        # the subtle row-level x buttons, since closing a whole card
        # is a bigger action).
        del_card_btn = Button()
        del_card_btn.Content         = u'\u2715'
        del_card_btn.FontSize        = 14
        try:
            del_card_btn.Style = self.FindResource('CloseButtonStyle')
        except Exception as e:
            logger.warning('Failed to apply DeleteButtonStyle: {}'.format(e))
        del_card_btn.FocusVisualStyle = None
        del_card_btn.HorizontalContentAlignment = HorizontalAlignment.Center
        del_card_btn.VerticalContentAlignment   = VerticalAlignment.Center
        del_card_btn.VerticalAlignment   = VerticalAlignment.Center
        del_card_btn.Margin          = Thickness(6, 0, 0, 0)
        del_card_btn.ToolTip         = 'Close (remove this card)'
        del_card_btn.Tag             = path
        del_card_btn.Click          += self._del_card_click
        header_right.Children.Add(del_card_btn)

        # Row container
        row_panel = StackPanel()
        row_panel.Orientation = Orientation.Vertical

        inner.Children.Add(header_row)

        # Column headers inside card
        col_hdr = StackPanel()
        col_hdr.Orientation = Orientation.Horizontal
        col_hdr.Margin      = Thickness(0, 0, 0, 8)

        def _ch(text, width, pad_left=4):
            tb = TextBlock()
            tb.Text             = text
            tb.Width            = width
            tb.FontSize         = 10
            tb.Foreground       = hb('#F4FAFF')
            tb.Opacity          = 0.45
            tb.VerticalAlignment = VerticalAlignment.Center
            tb.Padding          = Thickness(pad_left, 0, 0, 0)
            return tb

        # Tri-state select-all — indeterminate when some but not all
        # rows in this card are checked, so the header always reflects
        # the actual selection state at a glance. No fixed Width (auto-
        # sizes like the row checkboxes do) so it lines up with them
        # instead of drifting off to a different horizontal position.
        tri_cb = CheckBox()
        tri_cb.IsThreeState      = True
        tri_cb.Margin            = Thickness(0, 0, 6, 0)
        tri_cb.VerticalAlignment = VerticalAlignment.Center
        tri_cb.Tag               = path
        tri_cb.ToolTip           = 'Select all / none rows in this card'
        tri_cb.Click            += self._card_select_all_click
        fd['select_all_cb'] = tri_cb

        col_hdr.Children.Add(_ch('',           14, 0))
        col_hdr.Children.Add(tri_cb)                   # select-all
        col_hdr.Children.Add(self._header_button('View Name', 124, path, 'vn'))
        col_hdr.Children.Add(_ch('Sheet',     124))
        col_hdr.Children.Add(_ch('Range',     134))
        col_hdr.Children.Add(_ch('Modified',   96))
        # These four are buttons: each opens a dropdown that sets that one
        # thing on every TICKED row of this card.
        col_hdr.Children.Add(self._header_button('View Type', 134, path, 'vt'))
        col_hdr.Children.Add(self._header_button('Scale / Mode', 124, path, 'sm'))
        col_hdr.Children.Add(self._header_button('Split', 88, path, 'sp'))
        col_hdr.Children.Add(self._header_button('Rows / mm', 88, path, 'rv'))

        inner.Children.Add(col_hdr)
        inner.Children.Add(row_panel)
        outer.Child = inner

        fd['card_panel']    = row_panel
        fd['sections_hdr']  = col_hdr
        fd['card_border']   = outer
        fd['card_inner']    = inner

        outer.Tag = fd.get('real_path', path)        # what the list is sorted by
        self.CardsPanel.Children.Add(outer)
        self._sort_cards()
        self._update_card_reload_indicator(path)
        self._update_card_link_badge(path)
        self._update_tri_select_state(path)

    def _excel_view_name_live_check(self, sender, e):
        """Live (as-you-type) conflict check for an Excel row's View
        Name box."""
        row = sender.Tag
        if row is None:
            return
        name = sender.Text.strip()
        taken = self._view_name_taken(name, exclude_row=row)
        self._style_view_name_conflict(sender, taken)

    def _vn_lost(self, sender, e):
        row = sender.Tag
        if not row:
            return
        _old_name = row.ViewName
        row.ViewName = sender.Text.strip()
        # Typed by hand, so it stops following the named range. Clearing
        # the box hands it back: an empty name is not a chosen one.
        row._name_auto = not row.ViewName
        sender.Text  = row.ViewName
        if not row.ViewName:
            # Blank = neutral, just reset border — dot stays as-is
            self._style_view_name_conflict(sender, False)
            if row._dot and row.Status not in ('success', 'error'):
                row._dot.Fill = hb('#6B7280')
                row.Status = 'pending'
            return
        self._auto_check_row(row)
        taken = self._view_name_taken(row.ViewName, exclude_row=row)
        self._style_view_name_conflict(sender, taken)
        if taken:
            if row._dot:
                row._dot.Fill = hb('#DC2626')
                row.Status = 'error'
        else:
            if row._dot and row.Status not in ('success', 'skipped'):
                row._dot.Fill = hb('#6B7280')
                row.Status = 'pending'
        self._revalidate_all_view_name_boxes()
        if not taken and row.ViewName != _old_name:
            self._mark_row_changed(row)

    def _sc_changed(self, sender, e):
        if sender.SelectedItem is None:
            return
        tag = sender.Tag
        if not isinstance(tag, tuple):
            return
        row, rc = tag
        new_sheet = sender.SelectedItem
        _sheet_changed = (row.Sheet != new_sheet)
        row.Sheet = new_sheet
        rc.Items.Clear()
        for r in row.ranges_for(new_sheet):
            rc.Items.Add(r)
        if rc.Items.Count > 0:
            rc.SelectedIndex = 0
            row.NamedRange = rc.Items[0]
            if self._should_follow_range(row):
                self._apply_auto_view_name(row)
        self._auto_check_row(row)
        if _sheet_changed:
            self._mark_row_changed(row)

    def _should_follow_range(self, row):
        """Whether this row's view name should track its named range.

        Three cases, and only the middle one is new:

          * blank - nothing to lose;
          * never typed by hand - the name was filled in from the last
            range, so it belongs to the range and should move with it.
            Before, it stuck to whatever the FIRST range happened to be
            and the row then had to be cleared by hand before it would
            take a new one;
          * in conflict - a duplicate name blocks the row, and pointing
            it at a different range is exactly how you would expect to
            resolve that. Holding the clashing name would leave the row
            stuck on an error the user has just acted to fix.
        """
        if not row.ViewName:
            return True
        if getattr(row, '_name_auto', True):
            return True
        try:
            return self._view_name_taken(row.ViewName, exclude_row=row)
        except Exception:
            return False

    def _apply_auto_view_name(self, row):
        """Put the named range in the view name box and mark it ours."""
        row.ViewName = row.NamedRange
        row._name_auto = True
        if row._vn_textbox is not None:
            row._vn_textbox.Text = row.ViewName
        try:
            self._revalidate_all_view_name_boxes()
        except Exception:
            pass

    def _rc_changed(self, sender, e):
        if sender.SelectedItem is None:
            return
        row = sender.Tag
        if row:
            _range_changed = (row.NamedRange != sender.SelectedItem)
            row.NamedRange = sender.SelectedItem
            if self._should_follow_range(row):
                self._apply_auto_view_name(row)
            self._auto_check_row(row)
            if _range_changed:
                self._mark_row_changed(row)

    def _vt_changed(self, sender, e):
        if sender.SelectedItem is None:
            return
        row = sender.Tag
        if row:
            changed = (row.ViewType != sender.SelectedItem)
            row.ViewType = sender.SelectedItem
            self._sync_scale_mode_widgets(row)
            # A different view type means the view has to be rebuilt, so the
            # row must show as needing an update (and Apply must not skip it).
            if changed and row.Status == 'success':
                row.Status = 'sync'
                try:
                    if row._dot is not None:
                        row._dot.Fill = hb(STATUS_COLOURS.get('sync', '#3B82F6'))
                    self._set_sync_btn_state(row._refresh_btn, True)
                    self._refresh_status_card()
                except Exception:
                    pass
            if changed:
                self._auto_check_row(row)
                self._save_persisted_state()

    def _view_scale_changed(self, sender, e):
        row = sender.Tag
        if not row:
            return
        _old_scale = row.ViewScale
        try:
            val = int(sender.Text.strip())
            row.ViewScale = val if val > 0 else 1
        except Exception:
            row.ViewScale = 1
        sender.Text = str(row.ViewScale)
        if row.ViewScale != _old_scale:
            self._mark_row_changed(row)
        self._save_persisted_state()

    def _split_value_enable(self, row, mode):
        """The value box only means something for Rows and Height."""
        tb = getattr(row, '_split_value_tb', None)
        if tb is not None:
            tb.IsEnabled = mode in (SPLIT_ROWS, SPLIT_HEIGHT)
            tb.Opacity = 1.0 if tb.IsEnabled else 0.4

    def _split_mode_changed(self, sender, e):
        row = sender.Tag
        if not row or sender.SelectedItem is None:
            return
        changed = (row.SplitMode != str(sender.SelectedItem))
        row.SplitMode = str(sender.SelectedItem)
        self._split_value_enable(row, row.SplitMode)
        if changed:
            self._mark_row_changed(row)
            try:
                self._refresh_status_card()
            except Exception:
                pass
        self._save_persisted_state()

    def _split_value_changed(self, sender, e):
        """Whole rows for Rows, mm for Height; anything else clears it."""
        row = sender.Tag
        if not row:
            return
        _old_val = getattr(row, 'SplitValue', '')
        text = sender.Text.strip().lower().replace('mm', '').strip()
        try:
            val = float(text)
            if val <= 0:
                raise ValueError(text)
            row.SplitValue = (str(int(round(val)))
                              if row.SplitMode != SPLIT_HEIGHT
                              else u'{:g}'.format(val))
        except Exception:
            row.SplitValue = ''
        sender.Text = row.SplitValue
        if row.SplitValue != _old_val:
            self._mark_row_changed(row)
        self._save_persisted_state()

    def _parse_excel(self, path):
        real_path = path
        if path in self._file_data:
            # Already loaded — picking it again via + Add Tables reads
            # as "I want another card for this file", not "do nothing
            # and quietly jump back to the existing one".
            path = self._next_duplicate_key(real_path)
        try:
            wb = get_named_ranges_from_workbook(real_path)
        except Exception as ex:
            logger.error('Read failed: {}'.format(ex))
            wb = {}
        sheets = wb.get('sheets', [])
        srmap  = wb.get('sheet_ranges', {})
        if not srmap:
            all_r = wb.get('named_ranges', [])
            srmap = {s: all_r for s in sheets}
        self._file_data[path] = {
            'sheets':          sheets,
            'sheet_range_map': srmap,
            'source_type':     'xl',
            'rows':            [],
            'card_panel':      None,
            'card_border':     None,
            'real_path':       real_path,
        }
        self._active_file = path
        self._make_card(path)

