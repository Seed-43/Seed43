# -*- coding: utf-8 -*-
# script_create_excel_studio.py
#
# Excel writer for pyTransmit STUDIO layouts.
#
# A separate script from script_create_excel.py on purpose, not a patch of it.
# The two read different layout schemas and share no logic worth sharing:
#
#     script_create_excel.py   Layout Builder: rows of exactly 4 slots, slot 3
#                              a "spine" that fans out into rev_count columns.
#     this file                Studio: a real grid - n_rows x n_cols of cells
#                              with free rectangular merges, one revision per
#                              column, mm row heights and column widths.
#
# Because Studio's grid already IS a spreadsheet, this writer is mostly a
# direct transcription: one grid cell becomes one Excel cell, a merge becomes
# merge_range, a column width in mm becomes a column width in characters. The
# awkward parts of the Layout Builder writer - working out which slot spans
# what, expanding the spine - have no equivalent here.
#
# The row plan (grouping, one row per sheet) comes from Studio/studio_rows.py,
# the SAME module the canvas draws from, so what Studio shows and what Studio
# exports cannot drift apart.
#
# Payload keys used:
#   layout_json_path      the Studio template to write
#   _excel_save_path      where to save; None means "skip"
#   _pdf_temp_xlsx_path   set by script_create_pdf.py - write here, silently
#   group_params          sheet parameters to group the documentation table by
#   group_label           False = group header rows present but blank

_p = globals().get('PYTRANSMIT_PAYLOAD', {})

import os
import re
import json

from pyrevit import revit, script, forms

# ── Studio modules (pure Python: no WPF pulled in here) ──────────────────────
_PUBLISH_DIR = os.path.dirname(os.path.abspath(__file__))
_PT_ROOT = os.path.dirname(_PUBLISH_DIR)
_STUDIO_DIR = os.path.join(_PT_ROOT, 'Studio')
import sys as _sys
for _d in (_PT_ROOT, _STUDIO_DIR):
    if _d not in _sys.path:
        _sys.path.insert(0, _d)

import studio_rows
import studio_live_data
from pytransmit_paths import SETTINGS_DIR

try:
    from Snippets import _dialogs as sdlg
except Exception:
    sdlg = None


def _alert(message, title='', exitscript=False):
    if sdlg:
        sdlg.message(message, title=title)
    else:
        forms.alert(message, title=title)
    if exitscript:
        script.exit()


def _confirm(message, title='', no='No'):
    if sdlg:
        return sdlg.confirm(message, title=title, no=no)
    return bool(forms.alert(message, title=title, ok=False, yes=True, no=True))


_log_lines = _p.get('_log_lines', [])


def _log(msg):
    try:
        _log_lines.append(str(msg))
    except Exception:
        pass


doc = revit.doc

# ── Unit conversion ──────────────────────────────────────────────────────────
# Same constants the Layout Builder writer uses, so a column of a given width
# comes out the same size whichever tool drew it.
MM_TO_PT = 2.835          # 1 mm = 2.835 points (72 pt per inch) - row heights
MM_PER_PT = 0.352778


def mm_to_char_width(mm):
    """Millimetres -> Excel column width, in character units.

    Excel measures columns in characters, not length, and the relationship is
    fixed by its own definition: the default column is 8.43 characters = 64
    pixels = 16.93 mm at 96 dpi, with 5 pixels of cell padding. So

        pixels     = mm / 25.4 * 96
        characters = (pixels - 5) / 7

    where 7 px is the max digit width of the standard font.

    This replaces an inherited `mm * 0.175 * 5.5`, which works out at 0.9625
    characters per mm against a true 0.498 - every column came out 1.93 times
    too wide, which is why an A4 layout spilled across three pages. The same
    constant is still used by script_create_excel.py; its output is too wide
    for the same reason.
    """
    px = float(mm) / 25.4 * 96.0
    return max(0.0, (px - 5.0) / 7.0)


def mm_to_pt(mm):
    try:
        return float(mm) / MM_PER_PT
    except Exception:
        return 9.0


# ── Load the Studio layout ───────────────────────────────────────────────────
_layout_path = _p.get('layout_json_path')
if not _layout_path or not os.path.isfile(_layout_path):
    _alert('No Studio layout was assigned for the Excel output.',
           title='pyTransmit Studio - Excel', exitscript=True)

try:
    with open(_layout_path, 'r') as _f:
        LAYOUT = json.load(_f)
except Exception as _e:
    _alert('Could not read the Studio layout:\n{}\n\n{}'.format(_layout_path, _e),
           title='pyTransmit Studio - Excel', exitscript=True)

if 'cells' not in LAYOUT:
    # Guarded in pyTransmit too, but a script that can be run directly should
    # not depend on its caller having checked.
    _alert('"{}" is not a pyTransmit Studio layout.\n\nUse '
           'script_create_excel.py for Layout Builder templates.'.format(
               os.path.basename(_layout_path)),
           title='pyTransmit Studio - Excel', exitscript=True)

N_ROWS = int(LAYOUT.get('n_rows', 0))
N_COLS = int(LAYOUT.get('n_cols', 0))
ROW_H_MM = list(LAYOUT.get('row_heights') or [])
COL_W_MM = list(LAYOUT.get('col_widths') or [])
# The layout's own logo (chosen in Studio's Logo picker) wins; otherwise fall
# back to Branding's, which is what the Layout Builder writer uses. A layout
# saved before the logo picker existed has no logo_path at all, and silently
# printing nothing in that case is the wrong answer.
LOGO_PATH = (LAYOUT.get('logo_path', '') or '').strip()
if not LOGO_PATH:
    LOGO_PATH = (_p.get('logo_path', '') or '').strip()
    if LOGO_PATH:
        _log('Layout has no logo of its own; using the Branding logo.')
PAGE_W_MM = float(LAYOUT.get('page_w_mm', 210))
PAGE_H_MM = float(LAYOUT.get('page_h_mm', 297))
MARGIN_MM = float(LAYOUT.get('margin_mm', 10))
ORIENTATION = LAYOUT.get('orientation', 'portrait')
ROW_SECTIONS = list(LAYOUT.get('row_sections') or [])

# cells is a flat list of merge origins; covered cells are implied by the spans
CELLS = {}
for _entry in LAYOUT.get('cells', []) or []:
    CELLS[(_entry.get('r'), _entry.get('c'))] = _entry

_log('Studio layout: {} ({}x{})'.format(
    os.path.basename(_layout_path), N_ROWS, N_COLS))

# Same correction the Studio window applies when it opens a layout. Done here
# too because the window's version only reaches the file if the user saves,
# and a template carried over from before rows were expanded still holds the
# old "tall enough for the whole list" height - which would export as ~96pt
# rows with the text lost inside them.
_row_blocks = {}
for (_r, _c), _entry in CELLS.items():
    _row_blocks.setdefault(_r, []).append(_entry.get('block'))
ROW_H_MM, _rescaled = studio_rows.normalise_row_heights(ROW_H_MM, _row_blocks)
if _rescaled:
    _log('{} list row(s) resized to one row per item'.format(_rescaled))

# ── Live data ────────────────────────────────────────────────────────────────
# The same reader the Studio canvas uses, so the preview and the workbook are
# built from one description of the model rather than two.
DATA = studio_live_data.get_live_data(
    SETTINGS_DIR, group_params=_p.get('group_params') or [])
REVISIONS = DATA.get('revisions', []) or []
DOCS = DATA.get('docs', []) or []
DISTRIBUTION = DATA.get('distribution', []) or []

# Issue metadata chosen in the pyTransmit window right now overrides what the
# revision carries. studio_live_data reads Reason / Method / Document Type /
# Page Size out of the revision's IssuedTo tags ([R:] [M:] [F:] [P:]), and
# those tags are only written once a transmittal has been published - so on a
# first issue they are empty and those rows printed blank, while the Layout
# Builder writer showed them. It applies the same override
# (script_create_excel.py's _meta_vals), and so does this.
_META_KEYS = {'issued by': 'initials', 'initials': 'initials',
              'reason for issue': 'reason', 'method of issue': 'method',
              'document format': 'doc_format', 'paper size': 'paper_size'}
_meta_vals = {}
for _lbl, _val in (_p.get('meta_rows') or []):
    _k = _META_KEYS.get(str(_lbl).lower().strip())
    if _k and str(_val or '').strip():
        _meta_vals[_k] = _val
if _meta_vals:
    # Applied to every revision column, matching the Layout Builder writer.
    # Worth knowing: these are the CURRENT issue's settings, so on a
    # multi-revision transmittal they also overwrite older columns.
    for _rev in REVISIONS:
        _rev.update(_meta_vals)
    _log('Issue metadata from the pyTransmit window: {}'.format(
        ', '.join(sorted(_meta_vals.keys()))))

GROUP_PARAMS = _p.get('group_params') or []
# Border sides are tri-state (studio_rows.border_state). Settling them
# against the neighbouring cells is done once by studio_publish, the same
# call the canvas and the Revit writers make, so all four agree on which
# lines exist. EDGES is keyed (output row, column).
import studio_publish
EDGES = studio_publish.StudioLayout(LAYOUT, DATA, _p, log=_log).resolved_edges()
# pyTransmit's Text On/Off toggle decides this, always. The layout is for
# LAYOUT - the gaps, the colours, where things sit - not for which grouping to
# use or whether to name it. A 'group_label' key written by an earlier build is
# deliberately ignored, so a template cannot quietly override the toggle.
GROUP_LABEL = bool(_p.get('group_label', True))

# condense=False always: Condense Rows shortens the PREVIEW so a 2000-sheet
# layout stays workable on screen. The transmittal must list every sheet.
# The layout stores a gap pair for each group-text state; the state in force
# chooses which pair applies, so one template covers both.
_legacy_first = bool(LAYOUT.get('space_first_group', False))
_legacy_between = bool(LAYOUT.get('space_between_groups', False))
_suffix = 'on' if GROUP_LABEL else 'off'
_gap_first = bool(LAYOUT.get('space_first_' + _suffix, _legacy_first))
_gap_between = bool(LAYOUT.get('space_between_' + _suffix, _legacy_between))

ROW_PLAN = studio_rows.sheet_row_plan(
    DATA, GROUP_PARAMS, GROUP_LABEL, condense=False,
    space_first_group=_gap_first, space_between_groups=_gap_between)
VROWS, VSPANS = studio_rows.expand_rows(LAYOUT, DATA, ROW_PLAN)
# Stripe index per row: restarts at every group so each group opens white.
BAND_INDEX = studio_rows.band_indices(ROW_PLAN)

_log('{} revisions, {} sheets, {} recipients -> {} output rows'.format(
    len(REVISIONS), len(DOCS), len(DISTRIBUTION), len(VROWS)))

# ── Save path ────────────────────────────────────────────────────────────────
proj_number = DATA.get('proj_number', '') or ''
proj_name = DATA.get('proj_name', '') or (doc.Title if doc else '')
safe_base = re.sub(r'[\\/*?:"<>|]', '_',
                   'Document_Transmittal_{}_{}'.format(proj_number, proj_name))[:60]

save_path = None
_pdf_temp = _p.get('_pdf_temp_xlsx_path')
if _pdf_temp:
    save_path = _pdf_temp
elif '_excel_save_path' in _p:
    save_path = _p['_excel_save_path']
    if save_path is None:
        _log('Excel export skipped.')
        _sys.exit(0)
else:
    try:
        from System.Windows.Forms import SaveFileDialog, DialogResult
        dlg = SaveFileDialog()
        dlg.Title = 'Save Transmittal - Excel'
        dlg.Filter = 'Excel Workbook (*.xlsx)|*.xlsx'
        dlg.FileName = '{}.xlsx'.format(safe_base)
        dlg.InitialDirectory = os.path.expanduser('~\\Desktop')
        if dlg.ShowDialog() == DialogResult.OK:
            save_path = dlg.FileName
        else:
            _log('Save cancelled.')
            script.exit()
    except Exception:
        save_path = os.path.join(os.path.expanduser('~\\Desktop'),
                                 '{}.xlsx'.format(safe_base))

if not save_path:
    script.exit()

# ── Workbook ─────────────────────────────────────────────────────────────────
import xlsxwriter

workbook = xlsxwriter.Workbook(save_path)
ws = workbook.add_worksheet('Transmittal')

for _ci in range(N_COLS):
    _w_mm = COL_W_MM[_ci] if _ci < len(COL_W_MM) else 30.0
    ws.set_column(_ci, _ci, mm_to_char_width(_w_mm))

_content_mm = sum(COL_W_MM[:N_COLS])
_printable_mm = max(10.0, PAGE_W_MM - 2 * MARGIN_MM)
# Fit Width sizes the columns to EXACTLY the printable width, leaving no
# slack at all - and Excel renders each column in whole pixels, so a few
# columns rounding up is enough to push the last one onto a second page.
# Where the layout is meant to fit, tell Excel one page across: it then
# absorbs that rounding with a scale factor of about 99%, invisible on the
# page. This is not the blanket fit_to_pages that made an over-wide layout
# microscopic - a layout genuinely too wide for its paper is left at 100% and
# allowed to break, which is what Studio's page-break lines already showed.
_WIDTH_ROUNDING_TOLERANCE = 1.02
if _content_mm <= _printable_mm * _WIDTH_ROUNDING_TOLERANCE:
    # An explicit print scale, NOT fit_to_pages(1, 0).
    #
    # fit_to_pages writes fitToWidth="1" fitToHeight="0", where 0 means "no
    # limit" in OOXML - but LibreOffice, which converts the PDF, reads that 0
    # as ONE page tall and shrinks the whole sheet to fit the height. A long
    # sheet list then came out scaled to ~85%, which is where the extra
    # white band down the right-hand side came from.
    #
    # The scale is knowable here: content width against printable width. Doing
    # the arithmetic ourselves is deterministic, needs no cooperation from the
    # renderer, and only ever shrinks by the rounding Fit Width leaves behind.
    # One page ACROSS, an explicit "many" pages down.
    #
    # A computed print scale is not enough on its own: Excel lays each column
    # out in whole pixels, so thirteen columns each rounding up a fraction can
    # still spill onto a second page even when the millimetres add up to
    # exactly the printable width, which is what put this back to three pages.
    # Constraining the width lets Excel absorb that rounding itself.
    #
    # fitToHeight is set to Excel's maximum rather than 0. Zero means "no
    # limit" in OOXML, but LibreOffice - which converts the PDF - reads it as
    # ONE page tall and shrinks the whole sheet to fit, which is where the
    # white band down the right-hand side came from. A large explicit number
    # means the same thing to both and is misread by neither.
    _FIT_HEIGHT_UNLIMITED = 32767
    ws.fit_to_pages(1, _FIT_HEIGHT_UNLIMITED)
    _log('Fit to 1 page wide - columns are {:.2f}mm against {:.2f}mm '
         'printable.'.format(float(_content_mm), float(_printable_mm)))
else:
    _log('Layout is {:.0f}mm wide but {} {} printable width is {:.0f}mm - it '
         'will print {:.1f} pages across. Use Fit Width in Studio, or a wider '
         'paper size.'.format(float(_content_mm),
                              LAYOUT.get('page_size_name', 'the'),
                              ORIENTATION, float(_printable_mm),
                              float(_content_mm) / float(_printable_mm)))

# ── Page setup ───────────────────────────────────────────────────────────────
# Studio's page size, orientation and margins are what its page-break preview
# was drawn against, so carrying them over is what makes the printed workbook
# break in the same places the canvas showed.
_PAPER_CODES = {   # xlsxwriter set_paper() codes
    (210, 297): 9,    # A4
    (297, 420): 8,    # A3
    (420, 594): 66,   # A2
    (216, 279): 1,    # Letter
    (279, 432): 3,    # Tabloid
}
_portrait_dims = (int(round(min(PAGE_W_MM, PAGE_H_MM))),
                  int(round(max(PAGE_W_MM, PAGE_H_MM))))
_paper = _PAPER_CODES.get(_portrait_dims)
if _paper:
    ws.set_paper(_paper)
if ORIENTATION == 'landscape':
    ws.set_landscape()
else:
    ws.set_portrait()
_margin_in = MARGIN_MM / 25.4
ws.set_margins(left=_margin_in, right=_margin_in,
               top=_margin_in, bottom=_margin_in)
# Print scaling is decided with the column widths, further down - see the
# fit-to-width block there. Nothing is set here on purpose: set_print_scale()
# and fit_to_pages() are mutually exclusive in xlsxwriter (setting a scale
# turns fit-to-page back OFF), so calling set_print_scale(100) here silently
# undid the fit_to_pages() decided earlier and the sheet printed two pages
# wide again. Whichever runs last wins, so only one of them is ever called.

# Repeat-at-top rows become Excel's own print titles, which is what that
# setting means: the band reappears on every printed page.
_repeat_top = [i for i, (mr, _it) in enumerate(VROWS)
               if (ROW_SECTIONS[mr] if mr < len(ROW_SECTIONS) else '') == 'repeat_top']
if _repeat_top:
    ws.repeat_rows(min(_repeat_top), max(_repeat_top))


# ── Formats ──────────────────────────────────────────────────────────────────
_fmt_cache = {}
_NO_EDGES = (False, False, False, False)


def _hex(value, default='#000000'):
    v = (value or default)
    if not str(v).startswith('#'):
        v = '#' + str(v)
    return v


_JUST = {'left': 'left', 'center': 'center', 'right': 'right'}
_VJUST = {'top': 'top', 'middle': 'vcenter', 'bottom': 'bottom'}


def cell_format(block, kind='doc', wrap=False, alt=False,
                edges=(False, False, False, False)):
    """xlsxwriter format for a block, cached so the workbook doesn't end up
    with one format object per cell (Excel has a hard limit on those, and a
    2000-sheet transmittal would sail past it)."""
    b = block or {}
    bg = b.get('bg_color')
    # Border sides are tri-state and xlsx has no way to record "suppressed",
    # so what goes in the workbook is the RESOLVED edge - worked out against
    # the neighbouring cells once, up front, and passed in here.
    borders = {'t': edges[0], 'b': edges[1], 'l': edges[2], 'r': edges[3]}
    if kind == 'space':
        # A deliberate gap: no fill, nothing.
        bg = None
        bold = False
    elif kind == 'group':
        # Group header band, matching how the Layout Builder writer emphasises
        # it and how the Studio canvas draws it.
        bold = True
        bg = b.get('group_color') or bg or '#E8E8E8'
    else:
        bold = bool(b.get('bold'))
        if alt:
            # Banded rows: every other repeated row takes the band colour, so
            # a long sheet list reads white / grey / white down the page. The
            # canvas does the same, from the same two fields.
            bg = b.get('alt_color') or '#F5F7FA'
    rotation = int(b.get('rotation') or 0)

    key = (b.get('font'), b.get('size_mm'), bold, bool(b.get('italic')),
           bool(b.get('underline')), bool(b.get('strike')), b.get('color'), bg, alt,
           b.get('just'), b.get('v_just'),
           bool(borders.get('t')), bool(borders.get('b')),
           bool(borders.get('l')), bool(borders.get('r')),
           rotation, wrap, kind)
    if key in _fmt_cache:
        return _fmt_cache[key]

    d = {
        'font_name': b.get('font') or 'Arial',
        'font_size': round(mm_to_pt(b.get('size_mm') or (9 * MM_PER_PT)), 1),
        'bold': bold,
        'italic': bool(b.get('italic')),
        'underline': 1 if b.get('underline') else 0,
        'font_strikeout': bool(b.get('strike')),
        'font_color': _hex(b.get('color'), '#000000'),
        'align': _JUST.get(b.get('just'), 'left'),
        'valign': _VJUST.get(b.get('v_just'), 'vcenter'),
        'text_wrap': bool(wrap),
    }
    if bg:
        d['bg_color'] = _hex(bg, '#FFFFFF')
        d['pattern'] = 1
    if borders.get('t'):
        d['top'] = 1
    if borders.get('b'):
        d['bottom'] = 1
    if borders.get('l'):
        d['left'] = 1
    if borders.get('r'):
        d['right'] = 1
    # Studio stores 270 for the rotated Date of Issue headers; Excel counts
    # anticlockwise, so 270 there is 90 here.
    if rotation == 270:
        d['rotation'] = 90
    elif rotation:
        d['rotation'] = rotation

    f = workbook.add_format(d)
    _fmt_cache[key] = f
    return f


# ── Static (non-repeating) cell values ───────────────────────────────────────
_KV = {
    'proj_org': 'proj_org', 'proj_client': 'proj_client',
    'proj_number': 'proj_number', 'proj_name': 'proj_name',
    'doc_type': 'doc_type', 'print_size': 'print_size',
}
_KV_LABELS = {
    'proj_org': 'Organisation', 'proj_client': 'Client',
    'proj_number': 'Project No', 'proj_name': 'Project',
    'doc_type': 'Document Type', 'print_size': 'Print Size',
}
_SPINE_FIELD = {
    'spine_dates': 'date', 'spine_initials': 'initials',
    'spine_reason': 'reason', 'spine_method': 'method',
    'spine_doc_type': 'doc_format', 'spine_print_size': 'paper_size',
}


def rev_column_map():
    """Column index -> revision index.

    Studio's rule, copied exactly: every column holding at least one revision
    block is a revision column, numbered left to right, so a block dropped one
    column over reads the next revision. A block can pin itself with an
    integer rev_index.
    """
    rev_cols = set()
    for (r, c), entry in CELLS.items():
        block = entry.get('block') or {}
        if str(block.get('type', '')).startswith('spine_'):
            rev_cols.add(c)
    return dict((c, i) for i, c in enumerate(sorted(rev_cols)))


REV_COL_MAP = rev_column_map()


def rev_index_for(block, col):
    pinned = (block or {}).get('rev_index', 'auto')
    if isinstance(pinned, int):
        return max(0, pinned)
    return REV_COL_MAP.get(col, 0)


def static_value(block, col):
    """Text for a non-repeating block. Mirrors studio_blocks.render_block()'s
    value choices; only the drawing differs."""
    b = block or {}
    t = b.get('type', '')

    if t == 'text':
        return b.get('content', '') or ''
    if t == 'blank':
        return ''
    if t in _KV:
        value = str(DATA.get(_KV[t], '') or '')
        if not value:
            return ''
        label = b.get('label') or _KV_LABELS.get(t, '')
        return u'{}: {}'.format(label, value) if label else value
    if t == 'page_count':
        total = len(DOCS)
        fmt = b.get('page_format', 'Page X of Y')
        value = {'Page X': 'Page 1',
                 'Page X of Y': 'Page 1 of {}'.format(total),
                 'X of Y': '1 of {}'.format(total),
                 'X / Y': '1 / {}'.format(total)}.get(fmt, str(total))
        return u' '.join([p for p in (b.get('prefix', ''), value,
                                      b.get('suffix', '')) if p])
    if t == 'issue_date':
        value = REVISIONS[-1].get('date', '') if REVISIONS else ''
        return u' '.join([p for p in (b.get('prefix', ''), value,
                                      b.get('suffix', '')) if p])
    if t in ('reason_list', 'method_list'):
        items = DATA.get('reasons' if t == 'reason_list' else 'methods', []) or []
        pairs = [u'{}  {}'.format(i.get('code', ''), i.get('label', ''))
                 for i in items]
        # 'row' spreads across one row, 'list' stacks - in a single cell the
        # difference is the separator.
        return (u'    '.join(pairs) if b.get('list_style') == 'row'
                else u'\n'.join(pairs))
    if t.startswith('spine_'):
        idx = rev_index_for(b, col)
        if idx >= len(REVISIONS):
            return ''
        return str(REVISIONS[idx].get(_SPINE_FIELD.get(t, 'rev'), '') or '')
    if t == 'logo':
        return ''   # inserted as a picture, not text
    return ''


# ── Write the grid ───────────────────────────────────────────────────────────
def vrow_range(r0, r1):
    """Model row range -> (first output row, count)."""
    first = VSPANS.get(r0, (r0, 1))[0]
    last_start, last_n = VSPANS.get(r1, (r1, 1))
    return first, (last_start + last_n) - first


_sheet_driven_rows = set()
for _r in range(N_ROWS):
    for _c in range(N_COLS):
        _e = CELLS.get((_r, _c))
        if _e is not None and studio_rows.repeat_domain(_e.get('block')) == 'sheet':
            _sheet_driven_rows.add(_r)
            break


def _row_is_sheet_driven(r):
    """Does this model row repeat over the sheet list (rather than over the
    recipient list)? Only sheet-driven rows have group and condensed-marker
    rows in them, so only they need those rows treated specially."""
    return r in _sheet_driven_rows


# Row heights are applied AFTER the cells, because a rotated cell only turns
# out to need more height once its text is known - see _need_height().
_row_pt = []
for _v, (_mr, _it) in enumerate(VROWS):
    _h_mm = ROW_H_MM[_mr] if _mr < len(ROW_H_MM) else 8.0
    _row_pt.append(_h_mm * MM_TO_PT)

_logo_cells = []

# ── Page footer ──────────────────────────────────────────────────────────────
# "Anchor to Bottom of Every Page" is Excel's page FOOTER, not a row near the
# end of the sheet. Written as sheet rows it simply followed the last sheet in
# the table and floated in the middle of the page; as a footer it prints at
# the bottom margin of every page, which is what the setting says.
#
# It also lets Page Count mean what it says: in a footer Excel supplies the
# real page number (&P) and page total (&N). In a cell it could only ever
# print "Page 1 of <number of sheets>", which is why the transmittal claimed
# "Page 1 of 12" on a two page document.
_FOOTER_SECTION = 'repeat_bottom'
_footer_rows = set(r for r in range(N_ROWS)
                   if (ROW_SECTIONS[r] if r < len(ROW_SECTIONS) else '') == _FOOTER_SECTION)

_PAGE_CODES = {
    'Page X of Y': 'Page &P of &N',
    'Page X': 'Page &P',
    'X of Y': '&P of &N',
    'X / Y': '&P / &N',
}


def _esc(text):
    """Excel reads a lone & in a footer as the start of a field code, so text
    the user typed has to double its own ampersands. Applied ONLY to literal
    text - doubling the & in the &P / &N codes below would print them."""
    return u'{}'.format(text or '').replace('&', '&&')


def _footer_text(block, col):
    """One footer cell's text, already escaped, with Excel's own page fields
    where the block is a Page Count."""
    b = block or {}
    if b.get('type') == 'page_count':
        code = _PAGE_CODES.get(b.get('page_format', 'Page X of Y'), 'Page &P of &N')
        return u' '.join([p for p in (_esc(b.get('prefix', '')), code,
                                      _esc(b.get('suffix', ''))) if p])
    return _esc(static_value(b, col))


def _build_footer():
    """Assemble &L / &C / &R sections from the anchored rows' cells, placing
    each by its own horizontal alignment. Ampersands in user text are doubled,
    since Excel reads a single & as the start of a field code."""
    parts = {'left': [], 'center': [], 'right': []}
    font = None
    for r in sorted(_footer_rows):
        for c in range(N_COLS):
            entry = CELLS.get((r, c))
            if entry is None:
                continue
            block = entry.get('block') or {}
            text = _footer_text(block, c)
            if not str(text or '').strip():
                continue
            if font is None:
                font = (block.get('font') or 'Arial',
                        int(round(mm_to_pt(block.get('size_mm') or (9 * MM_PER_PT)))))
            # Already escaped by _footer_text(); escaping again here would
            # turn the page-number codes into literal text.
            parts[block.get('just') if block.get('just') in parts else 'left'].append(text)
    if not any(parts.values()):
        return ''
    # The trailing space after the size is NOT cosmetic. Excel's &<size> code
    # has no terminator - it swallows every digit that follows - so
    # "&7" + "2026/06/06" is read as font size 72026 with the text "/06/06",
    # which is exactly how the issue date lost its year. A non-digit after the
    # size ends it; a single space is the least intrusive one.
    prefix = u'&"{}"&{} '.format(font[0], font[1]) if font else u''
    out = u''
    for code, key in (('&L', 'left'), ('&C', 'center'), ('&R', 'right')):
        if parts[key]:
            out += code + prefix + u' '.join(parts[key])
    return out


_footer = _build_footer()
if _footer:
    ws.set_footer(_footer, {'margin': max(0.2, MARGIN_MM / 25.4 * 0.6)})
    _log('Anchored row(s) {} moved into the page footer.'.format(
        ', '.join(str(r + 1) for r in sorted(_footer_rows))))


def _write(row, col, text, fmt):
    """Always write_string, never write().

    write() dispatches on type and will turn a string that looks like a
    number, a date or a URL into one - so a revision date typed into Revit as
    "2026/06/06" could come out of Excel reformatted as something else
    entirely. Everything Studio writes is text that the user typed or that
    Revit reported, and it goes into the workbook exactly as it arrived.
    """
    ws.write_string(row, col, u'' if text is None else u'{}'.format(text), fmt)


def _merge(r1, c1, r2, c2, text, fmt):
    ws.merge_range(r1, c1, r2, c2, u'' if text is None else u'{}'.format(text), fmt)


def _need_height(block, text, v_first, v_count):
    """Grow rows so rotated text is not cut off.

    Sideways text needs HEIGHT in proportion to how long it is, and the row
    height that suits horizontal text is nowhere near enough - which is why
    the rotated Date of Issue fitted on the Studio canvas (which lays the text
    out for real) but came out clipped in the workbook. Excel will not
    auto-fit a row containing a merged or rotated cell, so the space has to be
    reserved here.

    The estimate is deliberately generous: too tall merely leaves white space,
    too short loses characters.
    """
    b = block or {}
    if int(b.get('rotation') or 0) not in (90, 270):
        return
    text = u'{}'.format(text or '')
    if not text:
        return
    pt = mm_to_pt(b.get('size_mm') or (9 * MM_PER_PT))
    needed = len(text) * pt * 0.62 + 8.0
    have = sum(_row_pt[v] for v in range(v_first, min(v_first + v_count, len(_row_pt))))
    if needed <= have or v_count <= 0:
        return
    # Spread the shortfall over the rows the cell spans, so a merged rotated
    # header grows as a block rather than leaving one giant row.
    extra = (needed - have) / float(v_count)
    for v in range(v_first, min(v_first + v_count, len(_row_pt))):
        _row_pt[v] += extra

for r in range(N_ROWS):
    if r in _footer_rows:
        continue   # drawn by Excel at the foot of every page, not here
    for c in range(N_COLS):
        entry = CELLS.get((r, c))
        if entry is None:
            continue
        block = entry.get('block')
        row_span = int(entry.get('row_span', 1) or 1)
        col_span = int(entry.get('col_span', 1) or 1)
        first_v, n_v = VSPANS.get(r, (r, 1))
        domain = studio_rows.repeat_domain(block)
        # If the ROW repeats, every column in it repeats - not just the
        # columns holding a list block. A static cell beside the sheet list
        # (a label, or an empty column) is a cell on each of those rows, not
        # one tall box spanning them: that is what the printed table looks
        # like, and a tall merge here would also collide with the full-width
        # group header rows below.
        repeats = (row_span == 1 and VROWS[first_v][1] is not None)

        if repeats:
            rev_idx = rev_index_for(block, c)
            static_text = None if domain else static_value(block, c)
            for i in range(n_v):
                v = first_v + i
                item = VROWS[v][1]
                if domain:
                    text, kind = studio_rows.repeat_cell_text(
                        block, DATA, ROW_PLAN, item, rev_idx)
                else:
                    text = static_text
                    kind = ROW_PLAN[item][0] if (
                        domain is None and 0 <= item < len(ROW_PLAN)
                        and _row_is_sheet_driven(r)) else 'doc'

                if kind in ('group', 'more'):
                    # The group label spans the whole table and is written by
                    # the Sheet Number column alone - every other column stays
                    # out of the way, exactly as script_create_excel.py does
                    # (it writes the merged group row from sheet_number and
                    # explicitly skips it for sheet_desc).
                    if (block or {}).get('type') != 'sheet_number':
                        continue
                    fmt = cell_format(block, kind,
                                      edges=EDGES.get((v, c), _NO_EDGES))
                    if N_COLS - 1 > c:
                        _merge(v, c, v, N_COLS - 1, text, fmt)
                    else:
                        _write(v, c, text, fmt)
                    continue

                # Band every other repeated row. The index is the position
                # in the row plan, so all the columns of one sheet band
                # together and group headers do not shift the pattern.
                bi = BAND_INDEX[item] if (domain == 'sheet' and item is not None
                                          and item < len(BAND_INDEX)) else item
                alt = (bool((block or {}).get('alt_rows'))
                       and bi is not None and bi % 2 == 1)
                fmt = cell_format(block, kind, alt=alt,
                                  edges=EDGES.get((v, c), _NO_EDGES))
                _need_height(block, text, v, 1)
                if col_span > 1:
                    _merge(v, c, v, c + col_span - 1, text, fmt)
                else:
                    _write(v, c, text, fmt)
            continue

        # Static cell: spans however many output rows its model rows expand to.
        v_first, v_count = vrow_range(r, r + row_span - 1)
        value = static_value(block, c)
        wrap = ('\n' in value) or ((block or {}).get('type') == 'text')
        fmt = cell_format(block, 'doc', wrap=wrap,
                          edges=EDGES.get((v_first, c), _NO_EDGES))
        _need_height(block, value, v_first, v_count)
        if v_count > 1 or col_span > 1:
            _merge(v_first, c, v_first + v_count - 1,
                   c + col_span - 1, value, fmt)
        else:
            _write(v_first, c, value, fmt)

        if (block or {}).get('type') == 'logo':
            if LOGO_PATH:
                _logo_cells.append((v_first, c, v_count, col_span, block))
            else:
                _log('Logo block at row {} col {} has no logo assigned - pick '
                     'one in Studio (Layout tab > Logo).'.format(r + 1, c + 1))

# Row heights last, now that rotated cells have had their say. Rows that went
# into the page footer are hidden rather than removed, so every row index
# computed above still refers to the same place.
for _v, _pt in enumerate(_row_pt):
    if VROWS[_v][0] in _footer_rows:
        ws.set_row(_v, None, None, {'hidden': True})
    else:
        ws.set_row(_v, _pt)

# ── Logo images ──────────────────────────────────────────────────────────────
# Inserted after the cells so they sit on top, and scaled to the cell they
# were placed in rather than at native size - a 2000px PNG would otherwise
# cover half the sheet.
if LOGO_PATH and os.path.isfile(LOGO_PATH) and _logo_cells:
    _PAD_PX = 2
    for (v, c, v_count, col_span, _lblock) in _logo_cells:
        try:
            cell_w_mm = sum(COL_W_MM[c:c + col_span]) if COL_W_MM else 30.0
            cell_h_mm = sum(ROW_H_MM[VROWS[vv][0]]
                            for vv in range(v, min(v + v_count, len(VROWS))))
            # Excel positions pictures in pixels at 96 dpi.
            box_w = max(8.0, cell_w_mm / 25.4 * 96 - 2 * _PAD_PX)
            box_h = max(8.0, cell_h_mm / 25.4 * 96 - 2 * _PAD_PX)

            # Excel sizes a picture from the file's DPI metadata, not its raw
            # pixel count: a 300 dpi logo is placed at roughly a third of its
            # pixel width. Scaling against raw pixels therefore produced a
            # picture a different size from the one being positioned, so the
            # alignment offsets - computed from that wrong width - dropped it
            # somewhere between centre and right instead of flush right.
            img_w = img_h = None
            try:
                from System.Drawing import Image as _Img
                _im = _Img.FromFile(LOGO_PATH)
                _dpi_x = float(_im.HorizontalResolution or 96.0) or 96.0
                _dpi_y = float(_im.VerticalResolution or 96.0) or 96.0
                img_w = float(_im.Width) * 96.0 / _dpi_x
                img_h = float(_im.Height) * 96.0 / _dpi_y
                _im.Dispose()
            except Exception:
                pass

            opts = {'object_position': 1}
            if img_w and img_h:
                # Fill the row height, and only shrink further if that would
                # make it wider than the cell. Aspect ratio is preserved, so a
                # logo always looks like itself.
                scale = box_h / img_h
                if img_w * scale > box_w:
                    scale = box_w / img_w
                draw_w, draw_h = img_w * scale, img_h * scale
                opts['x_scale'] = scale
                opts['y_scale'] = scale
            else:
                # Size unknown (no System.Drawing): centre it in the cell at
                # native size rather than guessing a scale.
                draw_w, draw_h = box_w, box_h

            # Placed by the block's own alignment, like any other cell content.
            just = (_lblock or {}).get('just', 'left')
            v_just = (_lblock or {}).get('v_just', 'middle')
            x_free, y_free = max(0.0, box_w - draw_w), max(0.0, box_h - draw_h)
            opts['x_offset'] = int(_PAD_PX + {'center': x_free / 2.0,
                                              'right': x_free}.get(just, 0.0))
            opts['y_offset'] = int(_PAD_PX + {'middle': y_free / 2.0,
                                              'bottom': y_free}.get(v_just, 0.0))
            ws.insert_image(v, c, LOGO_PATH, opts)
        except Exception as _img_e:
            _log('Logo could not be placed: {}'.format(_img_e))
elif LOGO_PATH and not os.path.isfile(LOGO_PATH):
    _log('Logo file is missing, skipped: {}'.format(LOGO_PATH))
elif LOGO_PATH and not _logo_cells:
    _log('A logo is assigned ({}) but the layout has no Logo block to put it '
         'in - drag one onto a cell in Studio.'.format(os.path.basename(LOGO_PATH)))

# ── Save ─────────────────────────────────────────────────────────────────────
# Retry rather than throw the work away. The usual cause is the previous
# transmittal still being open in Excel, which the user can fix in two seconds
# - but the workbook has already been built by this point, so cancelling means
# rebuilding it from scratch. xlsxwriter only marks the file closed on success,
# so calling close() again genuinely retries the write.
while True:
    try:
        workbook.close()
        break
    except Exception as _close_e:
        if not _confirm(
                'Could not save the transmittal:\n\n{}\n\n'
                'It is probably still open in Excel. Close it, then choose '
                'Retry.\n\nFile:\n{}'.format(_close_e, save_path),
                title='pyTransmit Studio - Excel', no='Cancel'):
            _log('Excel export cancelled - could not write {}'.format(save_path))
            _sys.exit(0)

_log('Workbook written: {}'.format(save_path))

# Silent when writing the PDF's temp file - script_create_pdf.py reports.
if not _pdf_temp and os.path.isfile(save_path):
    _open_dlg = _p.get('_open_file_dialog')
    if _open_dlg:
        if _open_dlg(u'Excel Saved', u'Open the file?'):
            os.startfile(save_path)
    elif '_excel_save_path' not in _p:
        if _confirm('Transmittal saved!\n\nOpen the file?'):
            os.startfile(save_path)
