# -*- coding: utf-8 -*-
# script_create_schedule_studio.py
#
# Revit Schedule writer for pyTransmit STUDIO layouts.
#
# A separate script from script_create_schedule.py on purpose, not a patch of
# it. The two read different layout schemas:
#
#     script_create_schedule.py   Layout Builder: rows of exactly 4 slots,
#                                 slot 3 a "spine" fanned out into rev_count
#                                 schedule columns, with a style algorithm
#                                 reconstructing which borders belong where.
#     this file                   Studio: a real grid, and a Revit schedule
#                                 header is also a real grid - so one Studio
#                                 cell becomes one schedule cell, a merge
#                                 becomes MergeCells, and a column width in mm
#                                 becomes a column width in feet.
#
# The whole document is drawn in the schedule's HEADER section. The body is
# filtered down to nothing (two contradictory filters on Assembly Code), which
# is the same trick script_create_schedule.py uses: a schedule body can only
# show model elements, and a transmittal is not a list of model elements.
#
# The grid, the row plan and every cell's text come from
# Studio/studio_publish.py - the same reading the Studio drafting-view writer
# uses, built on the same studio_rows module the Studio canvas draws from.
#
# Payload keys used:
#   layout_json_path   the Studio template to write
#   group_params       sheet parameters to group the documentation table by
#   group_label        False = group header rows present but blank
#   page_height_mode   'none' = one schedule, never split
#   page_height_mm     printable height override from the Setup panel

_p = globals().get('PYTRANSMIT_PAYLOAD', {})

import os
import re

from pyrevit import revit, script, DB, forms

from Autodesk.Revit.DB import (
    FilteredElementCollector, ViewSchedule, ScheduleFilter, ScheduleFilterType,
    SectionType, ElementId, Color, TableCellStyle, TableMergedCell,
    HorizontalAlignmentStyle, VerticalAlignmentStyle, GraphicsStyleType,
)

# ── Studio modules (pure Python: no WPF pulled in here) ──────────────────────
_PUBLISH_DIR = os.path.dirname(os.path.abspath(__file__))
_PT_ROOT = os.path.dirname(_PUBLISH_DIR)
_STUDIO_DIR = os.path.join(_PT_ROOT, 'Studio')
import sys as _sys
for _d in (_PT_ROOT, _STUDIO_DIR):
    if _d not in _sys.path:
        _sys.path.insert(0, _d)

import studio_publish
import studio_live_data
from pytransmit_paths import SETTINGS_DIR

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

# ── Constants ────────────────────────────────────────────────────────────────
MM = 1.0 / 304.8
FIELD_ID_ASM_CODE = -1002500     # Assembly Code, the field the body is filtered on
TITLE = 'pyTransmit Studio - Schedule'
SCHEDULE_NAME_FMT = 'pyTransmit Schedule {:02d}-{:02d}'

# Studio stores font sizes in mm, TableCellStyle.TextSize wants points.
#
# NOTE: script_create_schedule.py uses (size_mm / 0.75) * (72 / 25.4), which is
# this conversion times 4/3 - a stray px-to-pt factor that makes Layout Builder
# schedules print a third larger than the millimetres they name. Studio's
# canvas draws true millimetres and its templates were sized against it, so
# the true conversion is the right one here. Change this one constant if
# Studio schedules ever need to match the older ones rather than the layout.
PT_PER_MM = 72.0 / 25.4

_DEFAULT_SIZE_MM = 9.0 * studio_publish.MM_PER_PT

# ── Load the Studio layout ───────────────────────────────────────────────────
_layout_path = _p.get('layout_json_path')
if not _layout_path or not os.path.isfile(_layout_path):
    _alert('No Studio layout was assigned for the Schedule output.',
           title=TITLE, exitscript=True)

try:
    LAYOUT = studio_publish.load_layout(_layout_path)
except ValueError:
    # Guarded in pyTransmit too, but a script that can be run directly should
    # not depend on its caller having checked.
    _alert('"{}" is not a pyTransmit Studio layout.\n\nUse '
           'script_create_schedule.py for Layout Builder templates.'.format(
               os.path.basename(_layout_path)),
           title=TITLE, exitscript=True)
except Exception as _e:
    _alert('Could not read the Studio layout:\n{}\n\n{}'.format(_layout_path, _e),
           title=TITLE, exitscript=True)

_log('Studio layout: {} ({}x{})'.format(
    os.path.basename(_layout_path),
    LAYOUT.get('n_rows'), LAYOUT.get('n_cols')))

# The same reader the Studio canvas uses, so the preview and the schedule are
# built from one description of the model rather than two.
DATA = studio_live_data.get_live_data(SETTINGS_DIR)
SL = studio_publish.StudioLayout(LAYOUT, DATA, _p, log=_log)

N_COLS = SL.n_cols
if N_COLS < 1:
    _alert('The Studio layout has no columns.', title=TITLE, exitscript=True)

PLACEMENTS = SL.placements()

if any((pl['block'] or {}).get('type') == 'logo' for pl in PLACEMENTS):
    # Worth saying rather than silently dropping it: a Revit schedule holds
    # text and nothing else, which is why the Layout Builder schedule has no
    # logo either.
    _log('The layout has a Logo block. A Revit schedule cannot hold an image, '
         'so it is left out - use the Drafting View or Legend output for a '
         'transmittal with a logo.')

# ── Page splitting ───────────────────────────────────────────────────────────
# A schedule header cannot break across pages, so overflow becomes a SECOND
# schedule view - "pyTransmit Schedule 02-03" - which is what
# script_create_schedule.py does and what the sheet-placing workflow expects.
# Which rows go on which page is studio_publish's decision, so the Studio
# drafting-view writer breaks in the same places.
_SPLIT = (_p.get('page_height_mode') or 'a4') != 'none'
# The Setup panel's page height wins over the layout's own when it is set -
# the layout describes the paper it was drawn for, Setup describes the paper
# this transmittal is going onto.
_PRINTABLE_H_MM = float(_p.get('page_height_mm') or SL.printable_h())

PAGES = SL.page_rows(printable_h_mm=_PRINTABLE_H_MM, split=_SPLIT)
TOTAL_PAGES = len(PAGES)

# ── Line styles ──────────────────────────────────────────────────────────────
# Revit draws a schedule's grid lines unless every cell says otherwise, and
# "no border" is expressed by pointing the border at a WHITE line style rather
# than by switching it off - the same pair script_create_schedule.py creates.
_ON_ID = None
_OFF_ID = None


def _get_or_create_line_style(name, rgb):
    """Return the ElementId of a Lines subcategory in the given colour."""
    try:
        lines_cat = doc.Settings.Categories.get_Item('Lines')
        if lines_cat is None:
            return None
        sub = None
        for existing in lines_cat.SubCategories:
            if existing.Name == name:
                sub = existing
                break
        if sub is None:
            sub = doc.Settings.Categories.NewSubcategory(lines_cat, name)
        sub.LineColor = Color(rgb[0], rgb[1], rgb[2])
        try:
            sub.SetLineWeight(1, GraphicsStyleType.Projection)
        except Exception:
            pass
        return sub.Id if sub.GetGraphicsStyle(GraphicsStyleType.Projection) else None
    except Exception as e:
        _log('Line style "{}" could not be created: {}'.format(name, e))
        return None


with revit.Transaction('pyTransmit Studio - Line styles') as _t_ls:
    _ON_ID = _get_or_create_line_style('pyT On', (0, 0, 0))
    _OFF_ID = _get_or_create_line_style('pyT Off', (255, 255, 255))

_NO_LINE = _OFF_ID if _OFF_ID else ElementId.InvalidElementId
_LINE = _ON_ID if _ON_ID else ElementId.InvalidElementId

# ── Cell helpers ─────────────────────────────────────────────────────────────
_H_ALIGN = {'left': HorizontalAlignmentStyle.Left,
            'center': HorizontalAlignmentStyle.Center,
            'right': HorizontalAlignmentStyle.Right}
_V_ALIGN = {'top': VerticalAlignmentStyle.Top,
            'middle': VerticalAlignmentStyle.Middle,
            'bottom': VerticalAlignmentStyle.Bottom}


def _hex_to_rgb(value):
    """'#RRGGBB' / 'RGB' -> (r, g, b), or None when it isn't a colour."""
    try:
        h = str(value or '').strip().lstrip('#')
        if len(h) == 3:
            h = h[0] * 2 + h[1] * 2 + h[2] * 2
        if len(h) == 6:
            return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except Exception:
        pass
    return None


def _fill_rgb(placement):
    """The fill a placement wants, or None for no fill."""
    block = placement['block'] or {}
    if placement['kind'] == 'space':
        # A deliberate gap: no fill, no rules, nothing.
        return None
    if placement['kind'] == 'group':
        return _hex_to_rgb(block.get('group_color') or block.get('bg_color')
                           or '#E8E8E8')
    if placement['alt']:
        return _hex_to_rgb(block.get('alt_color') or '#F5F7FA')
    return _hex_to_rgb(block.get('bg_color'))


def _cell_text(placement):
    """Text for one schedule cell.

    Revit's schedule cells break lines on \\r\\n, so a stacked Reason/Method
    legend has to have its newlines translated or it prints as one run-on
    line.
    """
    text = u'{}'.format(placement['text'] or '')
    return text.replace('\r\n', '\n').replace('\n', '\r\n')


def _set_text(sec, r, c, text):
    try:
        sec.SetCellText(r, c, text)
    except Exception as e:
        _log('Cell text ({}, {}) failed: {}'.format(r, c, e))


def _merge(sec, r1, c1, r2, c2):
    if r1 == r2 and c1 == c2:
        return
    try:
        mc = TableMergedCell()
        mc.Top = r1
        mc.Bottom = r2
        mc.Left = c1
        mc.Right = c2
        sec.MergeCells(mc)
    except Exception as e:
        _log('Merge ({}, {}) - ({}, {}) failed: {}'.format(r1, c1, r2, c2, e))


def _blank_style():
    """A cell with no content and no rules - what an untouched header cell
    should look like once ShowGridLines has drawn its default box."""
    style = TableCellStyle()
    opts = style.GetCellStyleOverrideOptions()
    for attr in ('BorderTopLineStyle', 'BorderBottomLineStyle',
                 'BorderLeftLineStyle', 'BorderRightLineStyle'):
        setattr(opts, attr, True)
        setattr(style, attr, _NO_LINE)
    style.SetCellStyleOverrideOptions(opts)
    return style


def _cell_style(placement, borders):
    """Build the TableCellStyle for one placement.

    borders is (top, bottom, left, right) - already worked out for the edge of
    the placement this cell sits on, so a merged block only rules its outside.
    """
    block = placement['block'] or {}
    style = TableCellStyle()

    bold = bool(block.get('bold'))
    if placement['kind'] == 'group':
        # Group header band, matching how the Excel writers emphasise it and
        # how the Studio canvas draws it.
        bold = True

    style.IsFontBold = bold
    style.IsFontItalic = bool(block.get('italic'))
    style.TextSize = float(block.get('size_mm') or _DEFAULT_SIZE_MM) * PT_PER_MM
    try:
        style.FontName = block.get('font') or 'Arial'
    except Exception:
        pass

    fg = _hex_to_rgb(block.get('color')) if placement['kind'] != 'space' else None
    if fg:
        style.TextColor = Color(fg[0], fg[1], fg[2])
    bg = _fill_rgb(placement)
    if bg:
        style.BackgroundColor = Color(bg[0], bg[1], bg[2])

    style.FontHorizontalAlignment = _H_ALIGN.get(
        block.get('just'), HorizontalAlignmentStyle.Left)
    style.FontVerticalAlignment = _V_ALIGN.get(
        block.get('v_just'), VerticalAlignmentStyle.Middle)

    rotation = int(block.get('rotation') or 0)
    if rotation in (90, 270):
        try:
            style.TextOrientation = rotation
        except Exception:
            rotation = 0

    opts = style.GetCellStyleOverrideOptions()
    opts.Bold = True
    opts.Italics = True
    opts.FontSize = True
    opts.HorizontalAlignment = True
    opts.VerticalAlignment = True
    opts.FontColor = (fg is not None)
    opts.BackgroundColor = (bg is not None)
    if rotation in (90, 270):
        try:
            opts.TextOrientation = True
        except Exception:
            pass
    for attr, show in (('BorderTopLineStyle', borders[0]),
                       ('BorderBottomLineStyle', borders[1]),
                       ('BorderLeftLineStyle', borders[2]),
                       ('BorderRightLineStyle', borders[3])):
        setattr(opts, attr, True)
        setattr(style, attr, _LINE if show else _NO_LINE)
    style.SetCellStyleOverrideOptions(opts)
    return style


def _force_colours(sec, r, c, bg, fg):
    """Re-apply a cell's colours on top of whatever style it ended up with.

    Some schedule cells refuse a wholesale SetCellStyle - AllowOverrideCellStyle
    is False for them - and the fill and text colour are what go missing.
    Reading the style back and patching just those two gets through, which is
    what script_create_schedule.py's force_bg()/force_fg() do.
    """
    if not bg and not fg:
        return
    try:
        style = sec.GetTableCellStyle(r, c)
        opts = style.GetCellStyleOverrideOptions()
        if bg:
            opts.BackgroundColor = True
            style.BackgroundColor = Color(bg[0], bg[1], bg[2])
        if fg:
            opts.FontColor = True
            style.TextColor = Color(fg[0], fg[1], fg[2])
        style.SetCellStyleOverrideOptions(opts)
        sec.SetCellStyle(r, c, style)
    except Exception:
        pass


def _clear_header(sec):
    """Strip a header section back to one cell, ready to rebuild."""
    while sec.NumberOfRows > 1:
        try:
            sec.RemoveRow(sec.NumberOfRows - 1)
        except Exception:
            break
    while sec.NumberOfColumns > 1:
        try:
            sec.RemoveColumn(sec.NumberOfColumns - 1)
        except Exception:
            break


def _get_schedulable_field(sched_def, pid):
    for sf in sched_def.GetSchedulableFields():
        try:
            if sf.ParameterId.Value == pid:
                return sf
        except Exception:
            pass
    return None


def _get_or_create_schedule(name):
    """Return (schedule, existed) for a schedule of this name."""
    for vs in revit.query.get_elements_by_class(ViewSchedule, doc=doc):
        try:
            if vs.IsValidObject and vs.Name == name:
                return vs, True
        except Exception:
            pass
    vs = ViewSchedule.CreateSchedule(doc, ElementId.InvalidElementId)
    vs.Name = name
    return vs, False


# ── Stale schedules from a previous, longer run ──────────────────────────────
# A transmittal that used to need four schedules and now needs two would
# otherwise leave 03 and 04 behind, still showing the old issue.
_stale = []
for _vs in revit.query.get_elements_by_class(ViewSchedule, doc=doc):
    try:
        m = re.match(r'^pyTransmit Schedule (\d+)-\d+$', _vs.Name)
        if m and int(m.group(1)) > TOTAL_PAGES:
            _stale.append(_vs.Id)
    except Exception:
        pass
if _stale:
    with revit.Transaction('pyTransmit Studio - Remove stale schedules') as _t_stale:
        for _sid in _stale:
            try:
                doc.Delete(_sid)
            except Exception:
                pass
    _log('Removed {} stale schedule(s).'.format(len(_stale)))


# ── Render one page ──────────────────────────────────────────────────────────
def _render_page(hdr, page_vs):
    """Write one page's output rows into a schedule header section."""
    ri_of = dict((v, i) for i, v in enumerate(page_vs))
    n_rows = len(page_vs)

    while hdr.NumberOfColumns < N_COLS:
        hdr.InsertColumn(hdr.NumberOfColumns)
    for ci in range(N_COLS):
        try:
            hdr.SetColumnWidth(ci, SL.col_w(ci) * MM)
        except Exception:
            pass

    while hdr.NumberOfRows < n_rows:
        hdr.InsertRow(hdr.NumberOfRows)
    for v, ri in ri_of.items():
        try:
            hdr.SetRowHeight(ri, SL.row_h(v) * MM)
        except Exception:
            pass

    # Every cell starts ruleless; the placements below rule their own edges.
    # Without this the schedule's own grid lines show through wherever the
    # layout deliberately left a cell open.
    blank = _blank_style()
    for ri in range(n_rows):
        for ci in range(N_COLS):
            try:
                hdr.SetCellStyle(ri, ci, blank)
            except Exception:
                pass

    # Merges first: Revit will not accept a merge whose cells have already
    # been written, and a merged block takes the text and style of its origin.
    drawn = []
    for pl in PLACEMENTS:
        vs = [v for v in range(pl['v'], pl['v'] + pl['n_v']) if v in ri_of]
        if not vs:
            continue
        r1 = ri_of[vs[0]]
        r2 = ri_of[vs[-1]]
        c1 = pl['c']
        c2 = min(N_COLS - 1, c1 + max(1, pl['col_span']) - 1)
        _merge(hdr, r1, c1, r2, c2)
        drawn.append((pl, r1, r2, c1, c2))

    for pl, r1, r2, c1, c2 in drawn:
        _set_text(hdr, r1, c1, _cell_text(pl))
        # A deliberate gap carries no rules at all - drawn with the block's
        # borders it would read as an empty data row rather than a separator.
        b = {} if pl['kind'] == 'space' else ((pl['block'] or {}).get('borders') or {})
        bg = _fill_rgb(pl)
        fg = _hex_to_rgb((pl['block'] or {}).get('color')) if pl['kind'] != 'space' else None
        for ri in range(r1, r2 + 1):
            for ci in range(c1, c2 + 1):
                # Only the outside of a merged block is ruled; its inner
                # edges would print as a grid across a single cell.
                borders = (bool(b.get('t')) and ri == r1,
                           bool(b.get('b')) and ri == r2,
                           bool(b.get('l')) and ci == c1,
                           bool(b.get('r')) and ci == c2)
                try:
                    hdr.SetCellStyle(ri, ci, _cell_style(pl, borders))
                except Exception:
                    pass
                _force_colours(hdr, ri, ci, bg, fg)


# ── Write the schedules ──────────────────────────────────────────────────────
_written = []
for _page_idx, _page_vs in enumerate(PAGES, start=1):
    _name = SCHEDULE_NAME_FMT.format(_page_idx, TOTAL_PAGES)

    with revit.Transaction('pyTransmit Studio - Schedule {}'.format(_page_idx)) as _t:
        sched, existed = _get_or_create_schedule(_name)
        sched_def = sched.Definition

        if existed:
            # Filters go by index, not by object, and removing one shifts
            # every index after it - so they come off back to front. Fields
            # come off after, since removing a field takes its filters with it.
            for _fi in reversed(range(sched_def.GetFilterCount())):
                try:
                    sched_def.RemoveFilter(_fi)
                except Exception:
                    pass
            for fid in list(sched_def.GetFieldOrder()):
                try:
                    sched_def.RemoveField(fid)
                except Exception:
                    pass

        # One field, filtered down to nothing by two conditions that cannot
        # both hold: the body has to exist for the schedule to be valid, but
        # must never list anything.
        sf = _get_schedulable_field(sched_def, FIELD_ID_ASM_CODE)
        if sf is not None:
            field = sched_def.AddField(sf)
            field.ColumnHeading = ''
            sched_def.AddFilter(ScheduleFilter(
                field.FieldId, ScheduleFilterType.Equal, 'NO VALUES FOUND'))
            sched_def.AddFilter(ScheduleFilter(
                field.FieldId, ScheduleFilterType.Equal, 'ALL VALUES FOUND'))

        table = sched.GetTableData()
        hdr = table.GetSectionData(SectionType.Header)
        body = table.GetSectionData(SectionType.Body)

        if existed:
            _clear_header(hdr)

        try:
            body.SetColumnWidth(0, SL.table_w() * MM)
        except Exception:
            pass
        try:
            body.SetCellStyle(_blank_style())
        except Exception:
            pass
        try:
            sched_def.ShowGridLines = True
        except Exception:
            pass

        _render_page(hdr, _page_vs)

    _written.append(_name)

_log('Schedule(s) written: {} ({} output rows).'.format(
    ', '.join(_written), len(SL.vrows)))
