# -*- coding: utf-8 -*-
"""
LayoutSettings.py  —  pyTransmit Layout Builder v5
====================================================
Full WPF port of the HTML v5 prototype.

Features:
  - Block palette grouped by function (Layout / Project / Distribution /
    Documentation / Revision / Metadata)
  - Canvas rows with 4 column slots; drag-and-drop from palette and between slots
  - Column span stepper per block + row span for merged groups
  - Excel-style column drag-resize + editable percentages + settings dialog
  - Page size selector (A4/A3 portrait/landscape + custom mm)
  - Row merging: link adjacent rows into groups, move/delete as unit
  - Row section tags: Body / Repeat Header / Footer (for Excel print layout)
  - Per-block settings panel:
      · Horizontal + vertical justification
      · Background colour
      · Borders (Top / Bottom / Left / Right)
      · Data grid lines (Horizontal / Vertical) for data blocks
      · Spine options (Vertical lines / Rotation) for revision blocks
      · Text style reference
      · Alternate row colours (data blocks only)
      · List vs Row display (Reason / Method blocks)
      · Page format / date format / prefix / suffix (Page Count / Issue Date)
  - Global text styles (Title / Header / Data + user-defined)
      · Font family (system fonts), size (mm), Bold / Italic / Underline, colour
  - Logo path setting with browse
  - Revision count control (1-20)
  - Named templates with add / delete / save as
  - Per-template JSON files in Settings/layouts/
  - Live paper preview (scales to panel width, paper padding)

Usage:
    from LayoutSettings import LayoutSettingsWindow
    win = LayoutSettingsWindow(script_dir)
    win.ShowDialog()

    layout = win.get_active_layout()   # dict with rows, rev_count, col_pct ...
"""

import os
import json
import copy

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

import System
import System.Windows as _SW
import System.Windows.Controls as _SWC
import System.Windows.Documents as _SWD
import System.Windows.Input as _SWI
import System.Windows.Media as _SWM
import System.Windows.Media.Imaging as _SWMI
import System.Windows.Shapes as _SWS

from pyrevit.forms import WPFWindow


# ── Paths ─────────────────────────────────────────────────────────
LAYOUT_CONFIG_FILE = 'layout_config.json'    # UI state only
LAYOUTS_SUBDIR     = 'layouts'                # per-template JSON files

# ── Palette definition ────────────────────────────────────────────
# (type, label, icon, group, subtext)
# group: layout | project | dist | docs | revision | meta
PALETTE = [
    ('__grp__', 'Layout',         '', 'layout',   ''),
    ('logo',          'Logo',              '🖼', 'layout',   ''),
    ('text',          'Text',              '¶',  'layout',   'body / {{token}}'),
    ('blank',         'Blank Row',         '▭',  'layout',   '% height spacer'),
    ('page_count',    'Page Count',        '🔢', 'layout',   'page X of Y'),
    ('issue_date',    'Current Issue Date','📅', 'layout',   'latest revision date'),

    ('__grp__', 'Project Info',   '', 'project',  ''),
    ('proj_org',      'Organisation',      '🏢', 'project',  ''),
    ('proj_client',   'Client',            '👤', 'project',  ''),
    ('proj_number',   'Project Number',    '#',  'project',  ''),
    ('proj_name',     'Project Name',      '📋', 'project',  ''),

    ('__grp__', 'Distribution',   '', 'dist',     ''),
    ('sent_to',       'Sent To',           '📬', 'dist',     'company names'),
    ('attn_to',       'Attention To',      '✉',  'dist',     'contact names'),

    ('__grp__', 'Documentation',  '', 'docs',     ''),
    ('sheet_number',  'Sheet Number',      '#',  'docs',     'drawing numbers'),
    ('sheet_desc',    'Sheet Description', '📝', 'docs',     'drawing titles'),
    ('drawing_group', 'Drawing Group',     '📁', 'docs',     'group headers'),

    ('__grp__', 'Revision',       '', 'revision', ''),
    ('spine_dates',   'Date of Issue',     '📅', 'revision', 'rotated headers'),
    ('spine_copies',  'Number of Copies',  '📦', 'revision', 'per revision'),
    ('spine_rev',     'Revision Marks',    '🔄', 'revision', 'A B C per drawing'),
    ('spine_initials','Initials',          '✍',  'revision', 'issued by'),
    ('spine_reason',  'Reason for Issue Code', '🔖', 'revision', 'per revision'),
    ('spine_method',  'Method of Issue Code',  '📨', 'revision', 'per revision'),
    ('spine_doc_type','Document Type Code',     '🗂', 'revision', 'per revision'),
    ('spine_print_size','Print Size Code',      '📐', 'revision', 'per revision'),

    ('__grp__', 'Metadata',       '', 'meta',     ''),
    ('reason_list',   'Reason for Issue',  '❓', 'meta',     'list or row'),
    ('method_list',   'Method of Issue',   '📮', 'meta',     'list or row'),
    ('doc_type',      'Document Type',     '🗂', 'meta',     ''),
    ('print_size',    'Print Size',        '📐', 'meta',     ''),
]

TYPE_NAMES  = {p[0]: p[1] for p in PALETTE if p[0] != '__grp__'}
TYPE_ICONS  = {p[0]: p[2] for p in PALETTE if p[0] != '__grp__'}
TYPE_GROUP  = {p[0]: p[3] for p in PALETTE if p[0] != '__grp__'}
# Spine blocks aren't in the palette — group them as 'revision'
for _t in ('spine_dates','spine_initials','spine_reason','spine_method','spine_doc_type','spine_print_size','spine_copies','spine_rev'):
    TYPE_GROUP.setdefault(_t, 'revision')

GROUP_COLORS = {
    'layout':   _SWM.Color.FromRgb(0x8A, 0x96, 0xA8),
    'project':  _SWM.Color.FromRgb(0x29, 0x80, 0xB9),
    'dist':     _SWM.Color.FromRgb(0x27, 0xAE, 0x60),
    'docs':     _SWM.Color.FromRgb(0x8E, 0x44, 0xAD),
    'revision': _SWM.Color.FromRgb(0xE6, 0x7E, 0x22),
    'meta':     _SWM.Color.FromRgb(0x16, 0xA0, 0x85),
}

GROUP_LABELS = {
    'layout':   'LAYOUT',
    'project':  'PROJECT INFO',
    'dist':     'DISTRIBUTION',
    'docs':     'DOCUMENTATION',
    'revision': 'REVISION',
    'meta':     'METADATA',
}

SLOT_COLORS = [
    _SWM.Color.FromRgb(0x29, 0x80, 0xB9),
    _SWM.Color.FromRgb(0x27, 0xAE, 0x60),
    _SWM.Color.FromRgb(0x8E, 0x44, 0xAD),
    _SWM.Color.FromRgb(0xE6, 0x7E, 0x22),
]

# Data blocks that support alternate row colouring
DATA_BLOCK_TYPES = {'sent_to','attn_to','sheet_number','sheet_desc',
                    'reason_list','method_list','drawing_group',
                    'spine_copies','spine_rev'}

SPINE_HEADER_TYPES = {'spine_dates','spine_initials',
                      'spine_reason','spine_method','spine_doc_type','spine_print_size'}

# ── Default text styles ───────────────────────────────────────────
DEFAULT_TEXT_STYLES = {
    'Title':   {'font':'Arial','size_mm':4.5, 'bold':True, 'italic':False,'underline':False,'color':'#000000'},
    'Header':  {'font':'Arial','size_mm':2.5, 'bold':True, 'italic':False,'underline':False,'color':'#000000'},
    'Data':    {'font':'Arial','size_mm':2.2, 'bold':False,'italic':False,'underline':False,'color':'#000000'},
}

# Page sizes in mm (portrait)
PAGE_SIZES = {'a4': (210, 297), 'a4l': (297, 210), 'a3': (297, 420), 'a3l': (420, 297)}

# ── Dummy data ────────────────────────────────────────────────────
DUMMY = {
    'proj_org':    'Organization Name',
    'proj_client': 'Client Name',
    'proj_number': 'Project Number',
    'proj_name':   'Project Name',
    'doc_type':    'Structural Drawings',
    'print_size':  'A1',
    'distribution': [
        {'to': 'Architect/Designer', 'attn': 'Architect', 'copies': 1},
        {'to': 'Owner/Developer',    'attn': '',           'copies': 0},
        {'to': 'Contractor',         'attn': '',           'copies': 0},
        {'to': 'Local Authority',    'attn': '',           'copies': 0},
    ],
    'revisions': [
        {'rev':'A','date':'25/09/25','initials':'JD','reason':'C','method':'E','doc_type':'C','print_size':'A1'},
        {'rev':'B','date':'14/11/25','initials':'JD','reason':'IFC','method':'E','doc_type':'IFC','print_size':'A1'},
        {'rev':'C','date':'10/01/26','initials':'AB','reason':'BC','method':'P','doc_type':'BC','print_size':'A1'},
        {'rev':'D','date':'03/03/26','initials':'AB','reason':'A','method':'E','doc_type':'A','print_size':'A3'},
        {'rev':'E','date':'15/04/26','initials':'JD','reason':'C','method':'H','doc_type':'C','print_size':'A1'},
        {'rev':'F','date':'01/06/26','initials':'JD','reason':'IFC','method':'E','doc_type':'IFC','print_size':'A1'},
        {'rev':'G','date':'22/07/26','initials':'AB','reason':'A','method':'P','doc_type':'A','print_size':'A3'},
        {'rev':'H','date':'10/09/26','initials':'AB','reason':'BC','method':'E','doc_type':'BC','print_size':'A1'},
    ],
    'reasons': [
        {'code':'A','label':'Approval'},
        {'code':'BC','label':'Building Consent'},
        {'code':'C','label':'Construction'},
        {'code':'O','label':'Other'},
    ],
    'methods': [
        {'code':'E','label':'Email'},
        {'code':'P','label':'Post'},
        {'code':'H','label':'Hand Delivery'},
    ],
    'docs': [
        {'sheet':'G100', 'desc':'Cover',                         'revs':['A','B','','','','','','']},
        {'sheet':'C101', 'desc':'Site Plan',                     'revs':['A','B','C','','','','','']},
        {'sheet':'A101', 'desc':'Architectural Floor Plans',     'revs':['A','B','C','D','','','','']},
        {'sheet':'A201', 'desc':'Architectural Building Elevations','revs':['A','B','C','D','E','','','']},
        {'sheet':'A301', 'desc':'Architectural Building Sections','revs':['A','B','C','','','','','']},
        {'sheet':'A401', 'desc':'Architectural Enlarged Views',  'revs':['A','B','','','','','','']},
        {'sheet':'A501', 'desc':'Architectural Details',         'revs':['A','B','C','D','','','','']},
        {'sheet':'A601', 'desc':'Architectural Schedules',       'revs':['A','B','','','','','','']},
        {'sheet':'E101', 'desc':'Power Plans',                   'revs':['A','B','C','','','','','']},
        {'sheet':'S101', 'desc':'Structural Plans',              'revs':['A','B','C','D','','','','']},
        {'sheet':'M101', 'desc':'Mechanical Plans',              'revs':['A','B','','','','','','']},
        {'sheet':'P101', 'desc':'Plumbing Plans',                'revs':['A','B','C','','','','','']},
    ],
}

# ── Default template layout ───────────────────────────────────────
DATE_FORMATS = [
    'dd/MM/yyyy',
    'dd/MM/yy',
    'dd.MM.yyyy',
    'dd.MM.yy',
    'dd-MM-yyyy',
    'dddd, dd MMMM yyyy',
    'dd MMMM yyyy',
]

PAGE_COUNT_FORMATS = [
    'Page X',
    'Page X of Y',
    'X of Y',
    'X / Y',
    'Count only',
]

def _mk(t, **kw):
    d = {'type':t,'label':'','enabled':True,'span':1,'row_span':1,'just':'left','v_just':'middle',
         'borders':{'t':True,'b':True,'l':False,'r':False},
         'data_borders':{'h':True,'v':True},
         'text_style':'Data','list_style':'list','bg_color':None,
         'alt_rows':False,'alt_color':'#F5F7FA','height_pct':None,
         'content':'','rotation':0,
         'prefix':'','suffix':'',
         'page_format':'Page X of Y',
         'date_format':'dd/MM/yyyy'}
    d.update(kw)
    return d

def _txt(content, btype='text', just='left'):
    ts = 'Header' if btype == 'heading' else 'Data'
    # All text blocks are now type='text' (heading/title types removed from palette)
    return _mk('text', content=content, just=just, text_style=ts)

DEFAULT_ROWS = [
    {'blocks': [_mk('proj_org',span=2), None, _mk('proj_client'), _mk('proj_number')], 'merge_down': False},
    {'blocks': [_mk('proj_name',span=3), None, None, None], 'merge_down': False},
    {'blocks': [_txt('Distribution List','heading'), None, None, None], 'merge_down': False},
    {'blocks': [_txt('Sent To','text'), _txt('Attention To','text'), None, _txt('Number of Copies','text')], 'merge_down': False},
    {'blocks': [_mk('sent_to',alt_rows=True), _mk('attn_to',alt_rows=True), None, _mk('spine_copies')], 'merge_down': False},
    {'blocks': [_txt('Revision Legend','heading'), None, None, None], 'merge_down': False},
    {'blocks': [_mk('reason_list',list_style='list'), None, _txt('Date of Issue','text'), _mk('spine_dates')], 'merge_down': False},
    {'blocks': [None, None, _txt('Reason for Issue','text'), _mk('spine_reason')], 'merge_down': False},
    {'blocks': [_txt('Documentation List','heading'), None, None, None], 'merge_down': False},
    {'blocks': [_txt('Sheet','text'), _txt('Description','text'), None, _txt('Revision','text')], 'merge_down': False},
    {'blocks': [_mk('sheet_number',alt_rows=True), _mk('sheet_desc',span=2,alt_rows=True), None, _mk('spine_rev')], 'merge_down': False},
]

# ── Colour helpers ────────────────────────────────────────────────
def _brush(r, g, b, a=255):
    return _SWM.SolidColorBrush(_SWM.Color.FromArgb(a, r, g, b))

def _hbrush(h):
    h = (h or '#000000').lstrip('#')
    if len(h) == 3: h = h[0]*2 + h[1]*2 + h[2]*2
    return _brush(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

def _color_from_hex(h):
    h = (h or '#000000').lstrip('#')
    if len(h) == 3: h = h[0]*2 + h[1]*2 + h[2]*2
    return _SWM.Color.FromRgb(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

BK = {
    'bg':     _brush(0x3B,0x45,0x53),
    'deep':   _brush(0x23,0x29,0x33),
    'card':   _brush(0x2B,0x33,0x40),
    'row':    _brush(0x32,0x3D,0x4D),
    'hover':  _brush(0x3A,0x47,0x60),
    'accent': _brush(0x20,0x8A,0x3C),
    'aclt':   _brush(0x27,0xA8,0x49),
    'text':   _brush(0xF4,0xFA,0xFF),
    'muted':  _brush(0xE0,0xE6,0xEE),
    'bdr':    _brush(0x40,0x4E,0x62),
    'danger': _brush(0xC0,0x39,0x2B),
    'white':  _SWM.Brushes.White,
    'black':  _SWM.Brushes.Black,
    'transp': _SWM.Brushes.Transparent,
}
SLOT_BRUSHES = [_SWM.SolidColorBrush(c) for c in SLOT_COLORS]


# ══════════════════════════════════════════════════════════════════
# PREVIEW BUILDER
# ══════════════════════════════════════════════════════════════════

class PreviewBuilder(object):

    # Paper widths for preview (pixels) based on available width
    # All measurements derived from mm → px at preview scale

    def build(self, rows, rev_count, col_pct, page_w_mm, avail_px,
              text_styles, logo_path='', hlines=None, vlines=None):
        """Build and return (StackPanel content, paper_width_px)."""
        self._hlines = hlines or {}
        self._vlines = vlines or {}
        # Map from row dict id -> row index in _rows (for hlines/vlines lookup)
        self._row_ri = {id(r): i for i, r in enumerate(rows)}
        scale = avail_px / page_w_mm          # px per mm
        paper_px = int(avail_px)
        pad_px = int(10 * scale)              # 10mm margin
        inner_px = paper_px - pad_px * 2

        col_px = self._col_px(inner_px, col_pct, rev_count, scale)
        revs = DUMMY['revisions'][:rev_count]
        cw = max(int(col_px[3] / max(len(revs),1)), 6)

        root = _SWC.StackPanel()

        # Outer padding border
        pad_border = _SWC.Border()
        pad_border.Padding = _SW.Thickness(pad_px)
        inner_sp = _SWC.StackPanel()
        pad_border.Child = inner_sp
        root.Children.Add(pad_border)

        active = [r for r in rows if any(b and b.get('enabled') for b in r['blocks'])]
        if not active:
            tb = _SWC.TextBlock()
            tb.Text = 'No blocks enabled'
            tb.Foreground = BK['muted']; tb.FontSize = 8
            tb.HorizontalAlignment = _SW.HorizontalAlignment.Center
            tb.Margin = _SW.Thickness(0, 16, 0, 16)
            inner_sp.Children.Add(tb)
            return root, paper_px

        # Build row groups for merged-row support
        ri = 0
        while ri < len(active):
            # Find group extent in the active list
            grp_start = ri
            grp_end = ri
            while grp_end < len(active) - 1 and active[grp_end].get('merge_down', False):
                grp_end += 1

            if grp_end > grp_start:
                # Merged group — build a single combined row
                el = self._build_merged_rows(active[grp_start:grp_end + 1],
                                             col_px, revs, cw, scale, text_styles, logo_path)
                if el:
                    inner_sp.Children.Add(el)
            else:
                # Single row
                el = self._build_row(active[ri], col_px, revs, cw, scale, text_styles, logo_path)
                if el:
                    inner_sp.Children.Add(el)
            ri = grp_end + 1

        return root, paper_px

    def _build_merged_rows(self, rows, col_px, revs, cw, scale, text_styles, logo_path):
        """Build a single preview element for a group of merged rows.
        Uses a Grid with explicit rows and columns so column spans work correctly."""
        n_rows = len(rows)

        # Build a Grid with 4 columns and n_rows rows
        grid = _SWC.Grid()
        for px in col_px:
            cd = _SWC.ColumnDefinition(); cd.Width = _SW.GridLength(px)
            grid.ColumnDefinitions.Add(cd)
        for _ in range(n_rows):
            rd = _SWC.RowDefinition(); rd.Height = _SW.GridLength(1, _SW.GridUnitType.Auto)
            grid.RowDefinitions.Add(rd)

        # Track which cells are occupied (by column span or row span)
        occupied = set()  # set of (row_idx, col_idx) that are consumed

        for ri_idx in range(n_rows):
            row = rows[ri_idx]
            for ci in range(4):
                if (ri_idx, ci) in occupied:
                    continue

                b = row['blocks'][ci]
                if not b or not b.get('enabled', True):
                    continue

                col_span = min(b.get('span', 1), 4 - ci)
                row_span = min(b.get('row_span', 1), n_rows - ri_idx)

                # Mark occupied cells
                for rs in range(row_span):
                    for cs in range(col_span):
                        if rs == 0 and cs == 0:
                            continue  # the block's own cell
                        occupied.add((ri_idx + rs, ci + cs))

                # Also mark column-spanned cells in current row
                for cs in range(1, col_span):
                    occupied.add((ri_idx, ci + cs))

                px = sum(col_px[ci:ci + col_span])

                cell = _SWC.Border()
                cell.BorderBrush = BK['black']
                _gbrd = self._cell_border_from_grid(row, ci)
                cell.BorderThickness = _gbrd if _gbrd is not None else self._border_thickness(b.get('borders'))
                cell.VerticalAlignment = _SW.VerticalAlignment.Stretch
                cell.HorizontalAlignment = _SW.HorizontalAlignment.Stretch
                if b.get('bg_color'):
                    try: cell.Background = _hbrush(b['bg_color'])
                    except Exception: pass

                el = self._block_el(b, ci, px, revs, cw, scale, text_styles, logo_path)
                if el:
                    # Ensure the block element also stretches
                    if hasattr(el, 'VerticalAlignment'):
                        el.VerticalAlignment = _SW.VerticalAlignment.Stretch
                    if hasattr(el, 'HorizontalAlignment'):
                        el.HorizontalAlignment = _SW.HorizontalAlignment.Stretch
                    cell.Child = el

                _SWC.Grid.SetRow(cell, ri_idx)
                _SWC.Grid.SetColumn(cell, ci)
                if col_span > 1:
                    _SWC.Grid.SetColumnSpan(cell, col_span)
                if row_span > 1:
                    _SWC.Grid.SetRowSpan(cell, row_span)

                grid.Children.Add(cell)

        return grid

    def _col_px(self, inner_px, col_pct, rev_count, scale):
        a = int(inner_px * col_pct[0] / 100)
        b = int(inner_px * col_pct[1] / 100)
        c = int(inner_px * col_pct[2] / 100)
        d = inner_px - a - b - c
        return [a, b, c, d]

    def _occupied(self, row):
        occ = set()
        for i, b in enumerate(row['blocks']):
            if b and b.get('span', 1) > 1:
                for s in range(1, b['span']):
                    if i + s < 4: occ.add(i + s)
        return occ

    def _build_row(self, row, col_px, revs, cw, scale, text_styles, logo_path):
        occ = self._occupied(row)
        visible = []
        for ci in range(4):
            if ci in occ: continue
            b = row['blocks'][ci]
            if b and not b.get('enabled', True): b = None
            span = b.get('span', 1) if b else 1
            px = sum(col_px[min(ci+s, 3)] for s in range(span))
            visible.append({'ci': ci, 'block': b, 'px': px})

        has_content = any(v['block'] for v in visible)
        if not has_content: return None

        if len(visible) == 1 and visible[0]['block']:
            wrap = _SWC.Border()
            b = visible[0]['block']
            wrap.BorderBrush = BK['black']
            _gbrd0 = self._cell_border_from_grid(row, visible[0]['ci'])
            wrap.BorderThickness = _gbrd0 if _gbrd0 is not None else self._border_thickness(b.get('borders'))
            if b.get('bg_color'):
                try: wrap.Background = _hbrush(b['bg_color'])
                except Exception: pass
            el = self._block_el(visible[0]['block'], visible[0]['ci'],
                                visible[0]['px'], revs, cw, scale, text_styles, logo_path)
            if el: wrap.Child = el
            return wrap

        grid = _SWC.Grid()
        for v in visible:
            cd = _SWC.ColumnDefinition(); cd.Width = _SW.GridLength(v['px'])
            grid.ColumnDefinitions.Add(cd)

        outer = _SWC.Border()
        # Row-level border left off — each cell gets its own border from block settings

        for idx, v in enumerate(visible):
            cell = _SWC.Border()
            cell.BorderBrush = BK['black']
            if v['block']:
                _gbrdN = self._cell_border_from_grid(row, v['ci'])
                cell.BorderThickness = _gbrdN if _gbrdN is not None else self._border_thickness(v['block'].get('borders'))
                if v['block'].get('bg_color'):
                    try: cell.Background = _hbrush(v['block']['bg_color'])
                    except Exception: pass
                el = self._block_el(v['block'], v['ci'], v['px'],
                                    revs, cw, scale, text_styles, logo_path)
                if el: cell.Child = el
            else:
                cell.BorderThickness = _SW.Thickness(0)  # no border on empty cells
            _SWC.Grid.SetColumn(cell, idx)
            grid.Children.Add(cell)

        outer.Child = grid
        return outer

    def _apply_text_style(self, el, style_name, just, text_styles, scale):
        st = text_styles.get(style_name) or text_styles.get('Data') or DEFAULT_TEXT_STYLES['Data']
        try: el.FontFamily = _SWM.FontFamily(st.get('font','Arial'))
        except Exception: pass
        el.FontSize = max(4, st.get('size_mm', 2.5) * scale)
        el.FontWeight = _SW.FontWeights.Bold if st.get('bold') else _SW.FontWeights.Normal
        el.FontStyle  = _SW.FontStyles.Italic if st.get('italic') else _SW.FontStyles.Normal
        el.TextDecorations = _SW.TextDecorations.Underline if st.get('underline') else None
        try: el.Foreground = _hbrush(st.get('color','#000000'))
        except Exception: el.Foreground = BK['black']
        ta_map = {'left': _SW.TextAlignment.Left,
                  'center': _SW.TextAlignment.Center,
                  'right': _SW.TextAlignment.Right}
        el.TextAlignment = ta_map.get(just or 'left', _SW.TextAlignment.Left)

    def _v_align(self, v_just):
        m = {'top': _SW.VerticalAlignment.Top,
             'middle': _SW.VerticalAlignment.Center,
             'bottom': _SW.VerticalAlignment.Bottom}
        return m.get(v_just or 'middle', _SW.VerticalAlignment.Center)

    def _h_align(self, just):
        m = {'left':  _SW.HorizontalAlignment.Left,
             'center': _SW.HorizontalAlignment.Center,
             'right': _SW.HorizontalAlignment.Right}
        return m.get(just or 'left', _SW.HorizontalAlignment.Left)

    def _cell_border_from_grid(self, row, ci):
        # Get border thickness from hlines/vlines grid for preview
        ri = self._row_ri.get(id(row), -1)
        if ri < 0 or (not self._hlines and not self._vlines):
            return None  # fall back to block.borders

        def _hl(r, c):
            row_hl = self._hlines.get(r, [False]*4)
            return bool(row_hl[c]) if c < len(row_hl) else False

        def _vl(r, pos):
            row_vl = self._vlines.get(r, [False]*5)
            return bool(row_vl[pos]) if pos < len(row_vl) else False

        b = row['blocks'][ci]
        span = b.get('span', 1) if b else 1
        t = _hl(ri, ci)
        bm = _hl(ri + 1, ci)
        l = _vl(ri, ci)
        r2 = _vl(ri, ci + span)
        scale = 0.5  # preview line weight
        return _SW.Thickness(
            scale if l  else 0,
            scale if t  else 0,
            scale if r2 else 0,
            scale if bm else 0,
        )

    def _border_thickness(self, borders):
        b = borders or {}
        return _SW.Thickness(
            1 if b.get('l') else 0,
            1 if b.get('t') else 0,
            1 if b.get('r') else 0,
            1 if b.get('b') else 0,
        )

    def _block_el(self, block, ci, px, revs, cw, scale, text_styles, logo_path):
        t = block.get('type', '')
        sp = _SWC.StackPanel()
        sp.HorizontalAlignment = _SW.HorizontalAlignment.Stretch

        # Background colour
        if block.get('bg_color'):
            try: sp.Background = _hbrush(block['bg_color'])
            except Exception: pass

        pad = max(1, int(2 * scale))
        thick = _SW.Thickness(pad)

        # Data row height calculation — for consistent alignment across data blocks
        _st = text_styles.get(block.get('text_style', 'Data')) or DEFAULT_TEXT_STYLES.get('Data') or {'size_mm': 2.2}
        st_mm = _st.get('size_mm', 2.2)

        # Fixed data row height — ensures all data blocks align row-by-row in merged groups
        data_row_h = max(10, int(st_mm * scale * 1.8 + 2 * pad))

        def tb_row(text, style_name, just, alt=False, alt_color='#F5F7FA', show_h=True):
            row_b = _SWC.Border()
            row_b.BorderBrush = _hbrush('#E8E8E8')
            row_b.BorderThickness = _SW.Thickness(0, 0, 0, 1 if show_h else 0)
            row_b.Height = data_row_h
            if alt: row_b.Background = _hbrush(alt_color)
            tb = _SWC.TextBlock()
            tb.Text = str(text)
            tb.Padding = thick
            tb.TextTrimming = _SW.TextTrimming.CharacterEllipsis
            tb.HorizontalAlignment = _SW.HorizontalAlignment.Stretch
            tb.VerticalAlignment = _SW.VerticalAlignment.Center
            self._apply_text_style(tb, style_name, just, text_styles, scale)
            row_b.Child = tb
            return row_b

        # ── Blank row ────────────────────────────────────────────
        if t == 'blank':
            base = max(4, int(6 * scale))
            h_pct = block.get('height_pct') or 100
            el = _SWC.Border()
            el.Height = base * h_pct / 100
            return el

        # ── Text / Heading / Title ────────────────────────────────
        if t in ('title', 'heading', 'text'):
            # Use Grid so VerticalAlignment on inner wrapper is respected
            g = _SWC.Grid()
            # MinHeight >> text natural height so vertical alignment has room to work
            st_name = block.get('text_style', 'Header')
            st = text_styles.get(st_name) or DEFAULT_TEXT_STYLES.get(st_name) or DEFAULT_TEXT_STYLES['Data']
            line_h = max(8, st.get('size_mm', 2.5) * scale * 1.3)
            if block.get('bg_color'):
                try: g.Background = _hbrush(block['bg_color'])
                except Exception: pass
            wrapper = _SWC.Border()
            wrapper.Padding = _SW.Thickness(pad * 2, pad, pad * 2, pad)
            wrapper.VerticalAlignment = self._v_align(block.get('v_just'))
            wrapper.HorizontalAlignment = _SW.HorizontalAlignment.Stretch
            tb = _SWC.TextBlock()
            tb.Text = block.get('content', '')
            tb.TextWrapping = _SW.TextWrapping.Wrap
            self._apply_text_style(tb, block.get('text_style','Header'),
                                   block.get('just','left'), text_styles, scale)
            wrapper.Child = tb
            g.Children.Add(wrapper)
            return g

        # ── Logo ─────────────────────────────────────────────────
        if t == 'logo':
            # Wrap in Grid so vertical alignment has room to work
            g = _SWC.Grid()
            g.MinHeight = max(20 * scale / 3, 30)
            wrapper = _SWC.Border()
            wrapper.Padding = _SW.Thickness(pad * 2)
            wrapper.HorizontalAlignment = self._h_align(block.get('just'))
            wrapper.VerticalAlignment = self._v_align(block.get('v_just'))
            inner = _SWC.Border()
            inner.Background = _hbrush('#E8E8E8')
            inner.BorderBrush = _hbrush('#AAAAAA')
            inner.BorderThickness = _SW.Thickness(1)
            tb = _SWC.TextBlock()
            fname = os.path.basename(logo_path) if logo_path else ''
            tb.Text = '[ {} ]'.format(fname) if fname else '[ LOGO ]'
            tb.Padding = _SW.Thickness(pad * 2, pad, pad * 2, pad)
            self._apply_text_style(tb, 'Data', 'center', text_styles, scale)
            inner.Child = tb
            wrapper.Child = inner
            g.Children.Add(wrapper)
            sp.Children.Add(g)
            return sp

        # ── KV fields ─────────────────────────────────────────────
        KV_MAP = {
            'proj_org':    'Organisation',  'proj_client': 'Client',
            'proj_number': 'Project No',    'proj_name':   'Project',
            'doc_type':    'Document Type', 'print_size':  'Print Size',
        }
        KV_VAL = {
            'proj_org': DUMMY['proj_org'], 'proj_client': DUMMY['proj_client'],
            'proj_number': DUMMY['proj_number'], 'proj_name': DUMMY['proj_name'],
            'doc_type': DUMMY['doc_type'], 'print_size': DUMMY['print_size'],
        }
        if t in KV_MAP:
            g = _SWC.Grid()
            st_name = block.get('text_style', 'Data')
            st = text_styles.get(st_name) or DEFAULT_TEXT_STYLES.get(st_name) or DEFAULT_TEXT_STYLES['Data']
            line_h = max(8, st.get('size_mm', 2.5) * scale * 1.3)
            if block.get('bg_color'):
                try: g.Background = _hbrush(block['bg_color'])
                except Exception: pass
            wrapper = _SWC.Border()
            wrapper.Padding = _SW.Thickness(pad * 2, pad, pad * 2, pad)
            wrapper.VerticalAlignment = self._v_align(block.get('v_just'))
            wrapper.HorizontalAlignment = _SW.HorizontalAlignment.Stretch
            tb = _SWC.TextBlock()
            tb.TextTrimming = _SW.TextTrimming.CharacterEllipsis
            lbl_run = _SWD.Run()
            lbl_run.Text = (block.get('label') or KV_MAP[t]) + ': '
            lbl_run.Foreground = _hbrush('#555555')
            val_run = _SWD.Run()
            val_run.Text = KV_VAL.get(t, '')
            val_run.FontWeight = _SW.FontWeights.Bold
            tb.Inlines.Add(lbl_run); tb.Inlines.Add(val_run)
            self._apply_text_style(tb, block.get('text_style','Data'),
                                   block.get('just','left'), text_styles, scale)
            wrapper.Child = tb
            g.Children.Add(wrapper)
            return g

        # ── Page Count ────────────────────────────────────────────
        if t == 'page_count':
            g = _SWC.Grid()
            st = text_styles.get(block.get('text_style','Data')) or DEFAULT_TEXT_STYLES['Data']
            line_h = max(8, st.get('size_mm', 2.5) * scale * 1.3)
            wrapper = _SWC.Border()
            wrapper.Padding = thick
            wrapper.HorizontalAlignment = self._h_align(block.get('just'))
            wrapper.VerticalAlignment = self._v_align(block.get('v_just'))
            tb = _SWC.TextBlock()
            tb.TextTrimming = _SW.TextTrimming.CharacterEllipsis
            prefix = block.get('prefix', '')
            suffix = block.get('suffix', '')
            fmt = block.get('page_format', 'Page X of Y')
            total = len(DUMMY['docs'])
            if fmt == 'Page X':
                val = 'Page 1'
            elif fmt == 'Page X of Y':
                val = 'Page 1 of {}'.format(total)
            elif fmt == 'X of Y':
                val = '1 of {}'.format(total)
            elif fmt == 'X / Y':
                val = '1 / {}'.format(total)
            else:
                val = str(total)
            parts = [prefix, val, suffix]
            tb.Text = ' '.join(p for p in parts if p)
            self._apply_text_style(tb, block.get('text_style','Data'),
                                   block.get('just','left'), text_styles, scale)
            wrapper.Child = tb
            g.Children.Add(wrapper)
            sp.Children.Add(g)
            return sp

        # ── Current Issue Date ───────────────────────────────────
        if t == 'issue_date':
            g = _SWC.Grid()
            st = text_styles.get(block.get('text_style','Data')) or DEFAULT_TEXT_STYLES['Data']
            line_h = max(8, st.get('size_mm', 2.5) * scale * 1.3)
            wrapper = _SWC.Border()
            wrapper.Padding = thick
            wrapper.HorizontalAlignment = self._h_align(block.get('just'))
            wrapper.VerticalAlignment = self._v_align(block.get('v_just'))
            tb = _SWC.TextBlock()
            prefix = block.get('prefix', '')
            suffix = block.get('suffix', '')
            # Use latest revision date from dummy data
            try:
                from datetime import datetime as _dt
                today = _dt.now()
                dfmt = block.get('date_format', 'dd/MM/yyyy')
                # Convert .NET-style format to Python strftime
                py_fmt = dfmt.replace('dddd', '%A').replace('dd', '%d').replace('MMMM', '%B').replace('MM', '%m').replace('yyyy', '%Y').replace('yy', '%y')
                date_str = today.strftime(py_fmt)
            except Exception:
                date_str = '14/03/2025'
            parts = [prefix, date_str, suffix]
            tb.Text = ' '.join(p for p in parts if p)
            self._apply_text_style(tb, block.get('text_style','Data'),
                                   block.get('just','left'), text_styles, scale)
            wrapper.Child = tb
            g.Children.Add(wrapper)
            sp.Children.Add(g)
            return sp

        # ── Row-per-item data blocks ──────────────────────────────
        alt = block.get('alt_rows', False)
        alt_c = block.get('alt_color', '#F5F7FA')
        st_name = block.get('text_style', 'Data')
        just_val = block.get('just', 'left')

        ROW_DATA = {
            'sent_to':     [d['to']   for d in DUMMY['distribution']],
            'attn_to':     [d['attn'] for d in DUMMY['distribution']],
            'sheet_number':[d['sheet'] for d in DUMMY['docs']],
            'sheet_desc':  [d['desc']  for d in DUMMY['docs']],
        }
        if t in ROW_DATA:
            db = block.get('data_borders', {'h':True,'v':True})
            show_h = bool(db.get('h', True))
            for i, txt in enumerate(ROW_DATA[t]):
                row_b = tb_row(txt, st_name, just_val, alt and i%2==1, alt_c, show_h)
                row_b.HorizontalAlignment = _SW.HorizontalAlignment.Stretch
                sp.Children.Add(row_b)
            # Wrap in Grid so v_just on the whole row block works
            g = _SWC.Grid()
            sp.VerticalAlignment = self._v_align(block.get('v_just'))
            sp.HorizontalAlignment = _SW.HorizontalAlignment.Stretch
            g.Children.Add(sp)
            return g

        # ── Drawing groups ────────────────────────────────────────
        if t == 'drawing_group':
            db = block.get('data_borders', {'h':True,'v':True})
            show_h = bool(db.get('h', True))
            last_g = ''
            for i, doc in enumerate(DUMMY['docs']):
                g = doc['sheet'].rstrip('0123456789.').rstrip('0123456789')
                if g != last_g:
                    last_g = g
                    grp_b = _SWC.Border()
                    grp_b.Background = _hbrush('#E8E8E8')
                    grp_b.BorderBrush = _hbrush('#CCCCCC')
                    grp_b.BorderThickness = _SW.Thickness(0, 0, 0, 1 if show_h else 0)
                    grp_tb = _SWC.TextBlock()
                    grp_tb.Text = g
                    grp_tb.Padding = thick
                    grp_tb.FontWeight = _SW.FontWeights.Bold
                    self._apply_text_style(grp_tb, 'Header', block.get('just','left'), text_styles, scale)
                    grp_b.Child = grp_tb
                    sp.Children.Add(grp_b)
                empty_b = _SWC.Border()
                empty_b.Background = _hbrush(alt_c) if (alt and i%2==1) else BK['white']
                empty_b.BorderBrush = _hbrush('#E8E8E8')
                empty_b.BorderThickness = _SW.Thickness(0, 0, 0, 1 if show_h else 0)
                empty_b.Height = data_row_h
                sp.Children.Add(empty_b)
            g = _SWC.Grid()
            sp.VerticalAlignment = self._v_align(block.get('v_just'))
            sp.HorizontalAlignment = _SW.HorizontalAlignment.Stretch
            g.Children.Add(sp)
            return g

        # ── Reason / Method ───────────────────────────────────────
        items_map = {
            'reason_list': DUMMY['reasons'],
            'method_list': DUMMY['methods'],
        }
        if t in items_map:
            items = items_map[t]
            if block.get('list_style') == 'row':
                # Horizontal text row — evenly spaced, no boxes
                row_g = _SWC.Grid()
                row_g.Margin = _SW.Thickness(pad)
                for _ in items:
                    cd = _SWC.ColumnDefinition()
                    cd.Width = _SW.GridLength(1, _SW.GridUnitType.Star)
                    row_g.ColumnDefinitions.Add(cd)
                for idx, item in enumerate(items):
                    item_tb = _SWC.TextBlock()
                    item_tb.Text = '{}  {}'.format(item['code'], item['label'])
                    item_tb.Padding = _SW.Thickness(pad, 1, pad, 1)
                    item_tb.HorizontalAlignment = _SW.HorizontalAlignment.Stretch
                    self._apply_text_style(item_tb, st_name, just_val, text_styles, scale)
                    _SWC.Grid.SetColumn(item_tb, idx)
                    row_g.Children.Add(item_tb)
                sp.Children.Add(row_g)
            else:
                db = block.get('data_borders', {'h':True,'v':True})
                show_h = bool(db.get('h', True))
                show_v = bool(db.get('v', True))
                for i, item in enumerate(items):
                    row_g = _SWC.Grid()
                    row_g.Background = _hbrush(alt_c) if (alt and i%2==1) else BK['white']
                    cd1 = _SWC.ColumnDefinition(); cd1.Width = _SW.GridLength(18, _SW.GridUnitType.Star)
                    cd2 = _SWC.ColumnDefinition(); cd2.Width = _SW.GridLength(82, _SW.GridUnitType.Star)
                    row_g.ColumnDefinitions.Add(cd1); row_g.ColumnDefinitions.Add(cd2)
                    row_b = _SWC.Border()
                    row_b.Height = data_row_h
                    row_b.BorderBrush = _hbrush('#E8E8E8')
                    row_b.BorderThickness = _SW.Thickness(0, 0, 0, 1 if show_h else 0)
                    c1 = _SWC.Border()
                    c1.BorderBrush = _hbrush('#CCCCCC')
                    c1.BorderThickness = _SW.Thickness(0, 0, 1 if show_v else 0, 0)
                    t1 = _SWC.TextBlock(); t1.Text = item['code']; t1.Padding = thick
                    self._apply_text_style(t1, st_name, just_val, text_styles, scale); c1.Child = t1
                    c2 = _SWC.Border()
                    t2 = _SWC.TextBlock(); t2.Text = item['label']; t2.Padding = thick
                    self._apply_text_style(t2, st_name, just_val, text_styles, scale); c2.Child = t2
                    _SWC.Grid.SetColumn(c1, 0); _SWC.Grid.SetColumn(c2, 1)
                    row_g.Children.Add(c1); row_g.Children.Add(c2)
                    row_b.Child = row_g
                    sp.Children.Add(row_b)
            g_outer = _SWC.Grid()
            sp.VerticalAlignment = self._v_align(block.get('v_just'))
            sp.HorizontalAlignment = _SW.HorizontalAlignment.Stretch
            g_outer.Children.Add(sp)
            return g_outer

        # ── Spine blocks ──────────────────────────────────────────
        n = len(revs)
        if n == 0: return sp

        def spine_grid():
            g = _SWC.Grid()
            for _ in range(n):
                cd = _SWC.ColumnDefinition(); cd.Width = _SW.GridLength(cw)
                g.ColumnDefinitions.Add(cd)
            return g

        # Spine blocks honour data_borders too
        sp_db = block.get('data_borders', {'h':True,'v':True})
        sp_show_h = bool(sp_db.get('h', True))
        sp_show_v = bool(sp_db.get('v', True))

        # Read rotation from block (default 270 for spine headers)
        sp_rotation = block.get('rotation', 270)
        sp_just = block.get('just', 'center')

        def rot_cell(txt, ci2):
            b = _SWC.Border()
            b.BorderBrush = BK['black']
            b.BorderThickness = _SW.Thickness(0, 0, 1 if (ci2 < n-1 and sp_show_v) else 0, 0)
            tb = _SWC.TextBlock()
            tb.Text = str(txt or '')
            if sp_rotation == 270:
                tb.LayoutTransform = _SWM.RotateTransform(270)
                tb.Margin = _SW.Thickness(1, int(4*scale), 1, int(4*scale))
                self._apply_text_style(tb, st_name, 'center', text_styles, scale)
                g = _SWC.Grid()
                g.VerticalAlignment = _SW.VerticalAlignment.Stretch
                g.HorizontalAlignment = _SW.HorizontalAlignment.Stretch
                tb.HorizontalAlignment = self._h_align(sp_just)
                tb.VerticalAlignment = _SW.VerticalAlignment.Center
                g.Children.Add(tb)
                b.Child = g
            else:
                tb.Margin = _SW.Thickness(int(2*scale), 1, int(2*scale), 1)
                tb.HorizontalAlignment = _SW.HorizontalAlignment.Stretch
                tb.VerticalAlignment = self._v_align(block.get('v_just', 'middle'))
                self._apply_text_style(tb, st_name, sp_just, text_styles, scale)
                g = _SWC.Grid()
                g.VerticalAlignment = _SW.VerticalAlignment.Stretch
                g.HorizontalAlignment = _SW.HorizontalAlignment.Stretch
                g.Children.Add(tb)
                b.Child = g
            b.VerticalAlignment = _SW.VerticalAlignment.Stretch
            _SWC.Grid.SetColumn(b, ci2)
            return b

        def data_cell(txt, ci2, is_alt=False):
            b = _SWC.Border()
            if is_alt: b.Background = _hbrush(alt_c)
            b.BorderBrush = _hbrush('#D0D0D0')
            b.BorderThickness = _SW.Thickness(0, 0, 1 if (ci2 < n-1 and sp_show_v) else 0, 0)
            tb = _SWC.TextBlock()
            tb.Text = str(txt or '')
            tb.HorizontalAlignment = _SW.HorizontalAlignment.Stretch
            tb.VerticalAlignment = _SW.VerticalAlignment.Center
            tb.Margin = _SW.Thickness(1)
            self._apply_text_style(tb, st_name, sp_just, text_styles, scale)
            g = _SWC.Grid()
            g.VerticalAlignment = _SW.VerticalAlignment.Stretch
            g.HorizontalAlignment = _SW.HorizontalAlignment.Stretch
            g.Children.Add(tb)
            b.Child = g
            b.VerticalAlignment = _SW.VerticalAlignment.Stretch
            _SWC.Grid.SetColumn(b, ci2)
            return b

        if t == 'spine_dates':
            g = spine_grid()
            for i, r in enumerate(revs): g.Children.Add(rot_cell(r['date'], i))
            sp.Children.Add(g); return sp

        if t == 'spine_initials':
            g = spine_grid()
            for i, r in enumerate(revs): g.Children.Add(rot_cell(r['initials'], i))
            sp.Children.Add(g); return sp

        if t == 'spine_reason':
            g = spine_grid()
            for i, r in enumerate(revs): g.Children.Add(rot_cell(r['reason'], i))
            sp.Children.Add(g); return sp

        if t == 'spine_method':
            g = spine_grid()
            for i, r in enumerate(revs): g.Children.Add(rot_cell(r.get('method',''), i))
            sp.Children.Add(g); return sp

        if t == 'spine_doc_type':
            g = spine_grid()
            for i, r in enumerate(revs): g.Children.Add(rot_cell(r.get('doc_type',''), i))
            sp.Children.Add(g); return sp

        if t == 'spine_print_size':
            g = spine_grid()
            for i, r in enumerate(revs): g.Children.Add(rot_cell(r.get('print_size',''), i))
            sp.Children.Add(g); return sp

        if t == 'spine_copies':
            for ri, dist in enumerate(DUMMY['distribution']):
                g = spine_grid()
                row_outer = _SWC.Border()
                row_outer.Height = data_row_h
                row_outer.BorderBrush = _hbrush('#E0E0E0')
                row_outer.BorderThickness = _SW.Thickness(0, 0, 0, 1 if sp_show_h else 0)
                for ci2 in range(n):
                    val = dist['copies'] if ci2 == 0 and dist['copies'] else ''
                    g.Children.Add(data_cell(val, ci2, alt and ri%2==1))
                row_outer.Child = g
                sp.Children.Add(row_outer)
            return sp

        if t == 'spine_rev':
            for ri, doc in enumerate(DUMMY['docs']):
                g = spine_grid()
                row_outer = _SWC.Border()
                row_outer.Height = data_row_h
                row_outer.BorderBrush = _hbrush('#E0E0E0')
                row_outer.BorderThickness = _SW.Thickness(0, 0, 0, 1 if sp_show_h else 0)
                for ci2 in range(n):
                    val = doc['revs'][ci2] if ci2 < len(doc['revs']) else ''
                    g.Children.Add(data_cell(val, ci2, alt and ri%2==1))
                row_outer.Child = g
                sp.Children.Add(row_outer)
            return sp

        # Fallback
        fb = _SWC.Border(); fb.Padding = thick
        ftb = _SWC.TextBlock(); ftb.Text = TYPE_NAMES.get(t, t)
        ftb.Foreground = BK['muted']
        self._apply_text_style(ftb, 'Data', 'left', text_styles, scale)
        fb.Child = ftb; sp.Children.Add(fb)
        return sp


# ══════════════════════════════════════════════════════════════════
# WINDOW
# ══════════════════════════════════════════════════════════════════

class LayoutSettingsWindow(WPFWindow):

    def __init__(self, script_dir=None):
        if script_dir is None:
            script_dir = os.path.dirname(os.path.abspath(__file__))
        self._script_dir   = script_dir
        # Config and layouts live alongside this script (in the Layout/ folder)
        self._config_path  = os.path.join(script_dir, LAYOUT_CONFIG_FILE)
        self._layouts_dir  = os.path.join(script_dir, LAYOUTS_SUBDIR)

        if not os.path.isdir(self._layouts_dir):
            try: os.makedirs(self._layouts_dir)
            except: pass

        # Resolve XAML — same folder as this script
        xaml = os.path.join(script_dir, 'LayoutSettings.xaml')
        WPFWindow.__init__(self, xaml)

        # State
        self._rows          = []
        self._templates     = {}
        self._active_tmpl   = 'Excel'
        self._rev_count     = 4
        self._col_pct       = [18, 35, 17]   # A B C; D = remainder
        self._page_w_mm     = 210
        self._page_h_mm     = 297
        self._orientation   = 'portrait'
        self._text_styles   = copy.deepcopy(DEFAULT_TEXT_STYLES)
        self._logo_path     = ''
        self._uid           = 0
        self._drag_type      = None
        self._drag_from      = None
        self._pal_drag_start = None
        self._switching      = False
        self._updating_rev   = False
        self._preview_builder = PreviewBuilder()

        self._load_config()
        self._sync_ui_from_state()
        self._build_palette()
        self._render_canvas()
        self._render_preview()

        # Wire preview resize in Python (can't do in XAML with IronPython)
        self.preview_scroll.SizeChanged += self._on_prev_resize

        # Settings tab is open by default (XAML Visibility="Visible")
        # Set tab button highlight to match
        self.pal_settings_btn.Background = BK['accent']
        self.pal_blocks_btn.Background   = BK['row']
        self._render_style_cards()

        # Wire column-splitter drag handlers (Excel-style)
        try:
            self._wire_col_splitters()
        except Exception: pass

        # Window-level: accept all drags so they are never cancelled before
        # reaching cell drop targets
        self.AllowDrop = True
        self.DragOver += self._window_drag_over

    def _window_drag_over(self, sender, args):
        """Accept drag at window level — prevents WPF returning None effect."""
        try:
            args.Effects = _SW.DragDropEffects.Move
            args.Handled = True
        except Exception:
            pass

    # ── nid ─────────────────────────────────────────────────────
    def _nid(self):
        self._uid += 1
        return self._uid

    # ── Config I/O ───────────────────────────────────────────────
    def _template_file(self, name):
        """Sanitise template name for filesystem use."""
        safe = name.replace('/', '_').replace('\\', '_').replace(':', '_').strip()
        return os.path.join(self._layouts_dir, safe + '.json')

    def _load_config(self):
        # 1) Load UI state
        ui = {}
        try:
            with open(self._config_path, 'r') as f:
                ui = json.load(f)
        except Exception:
            ui = {}

        self._active_tmpl = ui.get('active_template', 'Excel')
        self._text_styles = ui.get('text_styles', copy.deepcopy(DEFAULT_TEXT_STYLES))
        self._logo_path   = ui.get('logo_path', '')
        self._col_pct     = ui.get('col_pct', [18, 35, 17])

        # 2) Load every template JSON found in the layouts dir
        self._templates = {}
        try:
            if os.path.isdir(self._layouts_dir):
                for fn in os.listdir(self._layouts_dir):
                    if not fn.lower().endswith('.json'): continue
                    path = os.path.join(self._layouts_dir, fn)
                    try:
                        with open(path, 'r') as f:
                            td = json.load(f)
                        # Template name = filename without .json (or use the template field)
                        tname = td.get('template') or os.path.splitext(fn)[0]
                        self._templates[tname] = {
                            'rev_count': int(td.get('rev_count', 4)),
                            'rows':      td.get('rows', []),
                            'col_pct':   td.get('col_pct', self._col_pct),
                            'page_w_mm':   td.get('page_w_mm', 210),
                            'page_h_mm':   td.get('page_h_mm', 297),
                            'orientation': td.get('orientation', 'portrait'),
                            'hlines':    td.get('hlines', {}),
                            'vlines':    td.get('vlines', {}),
                        }
                    except Exception:
                        continue
        except Exception: pass

        # 3) Ensure default templates exist (seed if missing)
        if 'Excel' not in self._templates:
            self._templates['Excel'] = {
                'rev_count': 4,
                'rows': copy.deepcopy(DEFAULT_ROWS),
                'col_pct': [18, 35, 17],
                'page_w_mm': 210,
            }
        for name in ('PDF', 'Revit Schedule', 'Revit Drafting View', 'Revit Legend'):
            if name not in self._templates:
                self._templates[name] = {
                    'rev_count': 4, 'rows': [],
                    'col_pct': [18, 35, 17], 'page_w_mm': 210,
                }

        if self._active_tmpl not in self._templates:
            self._active_tmpl = list(self._templates.keys())[0]

        # Migrate legacy block types (heading/title removed — text covers both)
        for tname, td in self._templates.items():
            for row in td.get('rows', []):
                for b in row.get('blocks', []):
                    if not b: continue
                    if b.get('type') == 'heading':
                        b['type'] = 'text'
                        if not b.get('text_style'): b['text_style'] = 'Header'
                    elif b.get('type') == 'title':
                        b['type'] = 'text'
                        if not b.get('text_style'): b['text_style'] = 'Title'

        td = self._templates[self._active_tmpl]
        self._rows        = copy.deepcopy(td.get('rows', []))
        self._rev_count   = int(td.get('rev_count', 4))
        self._page_w_mm   = td.get('page_w_mm', 210)
        self._page_h_mm   = td.get('page_h_mm', 297)
        self._orientation = td.get('orientation', 'portrait')
        self._hlines = {int(k): list(v) for k,v in td.get('hlines', {}).items()}
        self._vlines = {int(k): list(v) for k,v in td.get('vlines', {}).items()}

    def _save_config(self):
        self._flush_template()

        # 1) Write UI state (for editor next launch)
        ui = {
            'active_template': self._active_tmpl,
            'text_styles':     self._text_styles,
            'logo_path':       self._logo_path,
            'col_pct':         self._col_pct,
        }
        try:
            with open(self._config_path, 'w') as f:
                json.dump(ui, f, indent=2)
        except Exception as e:
            self._status('Config save failed: {}'.format(e)); return

        # 2) Write each template as its own self-contained JSON
        failed = []
        for name, td in self._templates.items():
            payload = {
                'template':    name,
                'rev_count':   int(td.get('rev_count', 4)),
                'col_pct':     td.get('col_pct', self._col_pct),
                'page_w_mm':   td.get('page_w_mm', getattr(self, '_page_w_mm', 210)),
                'page_h_mm':   td.get('page_h_mm', 297),
                'orientation': td.get('orientation', 'portrait'),
                'text_styles': copy.deepcopy(self._text_styles),
                'logo_path':   self._logo_path,
                'rows':        td.get('rows', []),
                'hlines':      td.get('hlines', {}),
                'vlines':      td.get('vlines', {}),
            }
            try:
                with open(self._template_file(name), 'w') as f:
                    json.dump(payload, f, indent=2)
            except Exception:
                failed.append(name)

        if failed:
            self._status('Save failed for: {}'.format(', '.join(failed)))
        else:
            self._status('Saved {} template(s) \u2713'.format(len(self._templates)))

    def _flush_template(self):
        # Re-read custom dimension TextBoxes directly so a partial edit is captured
        try:
            _tag = self.page_cb.SelectedItem.Tag if self.page_cb.SelectedItem else None
            if str(_tag) == 'custom':
                _w = getattr(self, 'custom_w_tb', None)
                _h = getattr(self, 'custom_h_tb', None)
                if _w:
                    try: self._page_w_mm = max(50, min(600,  int(float(_w.Text or '210'))))
                    except Exception: pass
                if _h:
                    try: self._page_h_mm = max(50, min(1200, int(float(_h.Text or '297'))))
                    except Exception: pass
        except Exception: pass
        self._templates[self._active_tmpl] = {
            'rev_count': self._rev_count,
            'rows':      copy.deepcopy(self._rows),
            'col_pct':   list(self._col_pct),
            'page_w_mm':   getattr(self, '_page_w_mm', 210),
            'page_h_mm':   getattr(self, '_page_h_mm', 297),
            'orientation': getattr(self, '_orientation', 'portrait'),
            'hlines': {str(k): list(v) for k,v in getattr(self, '_hlines', {}).items()},
            'vlines': {str(k): list(v) for k,v in getattr(self, '_vlines', {}).items()},
        }

    # ── Border grid helpers ─────────────────────────────────────────────
    def _get_hline(self, ri, ci):
        hl = getattr(self, '_hlines', {})
        row_hl = hl.get(ri, [False]*4)
        return bool(row_hl[ci]) if ci < len(row_hl) else False

    def _get_vline(self, ri, pos):
        vl = getattr(self, '_vlines', {})
        row_vl = vl.get(ri, [False]*5)
        return bool(row_vl[pos]) if pos < len(row_vl) else False

    def _toggle_hline(self, ri, ci_start, span):
        if not hasattr(self, '_hlines'): self._hlines = {}
        if ri not in self._hlines: self._hlines[ri] = [False]*4
        new_val = not self._hlines[ri][ci_start]
        for ci in range(ci_start, min(ci_start+span, 4)):
            self._hlines[ri][ci] = new_val
        self._render_canvas(); self._render_preview()

    def _toggle_vline(self, ri, pos, locked=False):
        if locked: return
        if not hasattr(self, '_vlines'): self._vlines = {}
        if ri not in self._vlines: self._vlines[ri] = [False]*5
        self._vlines[ri][pos] = not self._vlines[ri][pos]
        self._render_canvas(); self._render_preview()

    # ── Sync UI controls ─────────────────────────────────────────
    # Canonical template order
    _TEMPLATE_ORDER = ['Excel', 'PDF', 'Revit Schedule', 'Revit Drafting View', 'Revit Legend']

    def _ordered_templates(self):
        """Return template names in canonical order, with any custom (non-built-in) names appended A-Z."""
        names = list(self._templates.keys())
        ordered = [n for n in self._TEMPLATE_ORDER if n in names]
        custom = sorted([n for n in names if n not in self._TEMPLATE_ORDER], key=lambda x: x.lower())
        return ordered + custom

    def _sync_ui_from_state(self):
        self._switching = True
        try:
            # Template combo
            self.template_cb.Items.Clear()
            for name in self._ordered_templates():
                self.template_cb.Items.Add(name)
            self.template_cb.SelectedItem = self._active_tmpl
        finally:
            self._switching = False

        self._updating_rev = True
        try: self.rev_tb.Text = str(self._rev_count)
        except Exception: pass
        finally: self._updating_rev = False

        try: self.logo_path_tb.Text = self._logo_path or ''
        except Exception: pass

        try:
            self._update_col_labels()
        except Exception: pass

        self._update_rev_label()
        self._update_prev_info()

    def _update_rev_label(self):
        # No-op now — lbl_d shows the column percentage, not rev count
        # (rev count is shown in the header rev stepper and prev_info_tb)
        pass

    def _update_prev_info(self):
        try:
            self.prev_info_tb.Text = '{}x{}mm {} · {} revs · dummy data'.format(
                int(self._page_w_mm), int(self._page_h_mm), self._orientation, self._rev_count)
        except Exception: pass

    def _update_col_labels(self):
        try:
            a, b, c = self._col_pct[0], self._col_pct[1], self._col_pct[2]
            d = max(5, 100 - a - b - c)
            # lbl_a/b/c are editable TextBoxes (no '%')
            self.lbl_a.Text = '{}'.format(a)
            self.lbl_b.Text = '{}'.format(b)
            self.lbl_c.Text = '{}'.format(c)
            # lbl_d is a TextBlock (residual, no '%' — % is adjacent)
            self.lbl_d.Text = '{}'.format(d)
            # Update the splitter bar's column ratios
            try:
                self.csb_a.Width = _SW.GridLength(a, _SW.GridUnitType.Star)
                self.csb_b.Width = _SW.GridLength(b, _SW.GridUnitType.Star)
                self.csb_c.Width = _SW.GridLength(c, _SW.GridUnitType.Star)
                self.csb_d.Width = _SW.GridLength(d, _SW.GridUnitType.Star)
            except Exception: pass
        except Exception: pass

    def col_pct_keydown(self, s, e):
        """Commit on Enter; revert on Escape."""
        try:
            if e.Key == _SWI.Key.Enter:
                self._commit_pct_edit(s)
                e.Handled = True
            elif e.Key == _SWI.Key.Escape:
                self._update_col_labels()  # revert from current state
                e.Handled = True
        except Exception: pass

    def col_pct_lostfocus(self, s, e):
        """Commit when textbox loses focus."""
        self._commit_pct_edit(s)

    def _commit_pct_edit(self, tb):
        """Apply the user-entered % from a TextBox (Tag holds index 0/1/2 for A/B/C)."""
        try:
            idx = int(str(tb.Tag))
            new_val = int(round(float(tb.Text.replace('%', '').strip())))
            new_val = max(5, min(80, new_val))
            # Get other two values
            others = list(self._col_pct)
            others[idx] = new_val
            a, b, c = others
            # Ensure D >= 5%
            if a + b + c > 95:
                # squeeze others proportionally to make room for D=5
                excess = (a + b + c) - 95
                if idx != 0: a = max(5, a - excess // 2); excess = (a + b + c) - 95
                if excess > 0 and idx != 1: b = max(5, b - excess // 2); excess = (a + b + c) - 95
                if excess > 0 and idx != 2: c = max(5, c - excess);
            self._col_pct = [a, b, c]
            self._on_columns_changed()
        except Exception:
            self._update_col_labels()

    def _wire_col_splitters(self):
        """Wire mouse-drag handlers on the column splitter borders.
        Each splitter resizes the columns on either side of it."""
        # Splitter index maps to which boundary it controls:
        #  splitter_ab: between A and B  -> moves pct between [0] and [1]
        #  splitter_bc: between B and C  -> moves pct between [1] and [2]
        #  splitter_cd: between C and D  -> moves pct between [2] and (D=residual)
        boundaries = [
            (self.splitter_ab, 0),
            (self.splitter_bc, 1),
            (self.splitter_cd, 2),
        ]
        for splitter, idx in boundaries:
            self._wire_one_splitter(splitter, idx)

    def _wire_one_splitter(self, splitter, idx):
        state = {'dragging': False, 'start_x': 0.0, 'start_pct': None, 'bar_w': 1.0}

        def on_down(s, e, _idx=idx, _state=state):
            try:
                _state['dragging'] = True
                _state['start_x'] = e.GetPosition(self.col_splitter_bar).X
                _state['start_pct'] = list(self._col_pct)
                _state['bar_w'] = max(1.0, self.col_splitter_bar.ActualWidth)
                s.CaptureMouse()
                s.Background = BK['accent']
            except Exception: pass

        def on_move(s, e, _idx=idx, _state=state):
            if not _state['dragging']: return
            try:
                cur_x = e.GetPosition(self.col_splitter_bar).X
                delta_x = cur_x - _state['start_x']
                # Convert pixel delta to percentage delta
                delta_pct = int(round(delta_x / _state['bar_w'] * 100.0))
                a0, b0, c0 = _state['start_pct']
                d0 = max(5, 100 - a0 - b0 - c0)

                if _idx == 0:
                    # AB boundary: move from B to A (or vice versa)
                    new_a = max(5, min(60, a0 + delta_pct))
                    actual = new_a - a0
                    new_b = max(5, b0 - actual)
                    self._col_pct = [new_a, new_b, c0]
                elif _idx == 1:
                    # BC boundary
                    new_b = max(5, min(70, b0 + delta_pct))
                    actual = new_b - b0
                    new_c = max(5, c0 - actual)
                    self._col_pct = [a0, new_b, new_c]
                else:
                    # CD boundary — moving means changing C; D = residual
                    new_c = max(5, min(60, c0 + delta_pct))
                    # Ensure D stays >= 5%
                    max_c = 100 - a0 - b0 - 5
                    new_c = min(new_c, max_c)
                    self._col_pct = [a0, b0, new_c]

                self._on_columns_changed()
            except Exception: pass

        def on_up(s, e, _state=state):
            try:
                if _state['dragging']:
                    _state['dragging'] = False
                    s.ReleaseMouseCapture()
                    s.Background = BK['bdr']
            except Exception: pass

        splitter.MouseLeftButtonDown += on_down
        splitter.MouseMove           += on_move
        splitter.MouseLeftButtonUp   += on_up

    # ── Header events ─────────────────────────────────────────────
    def rev_dec_click(self, s, e):
        if self._rev_count > 1:
            self._rev_count -= 1
            self._updating_rev = True
            try: self.rev_tb.Text = str(self._rev_count)
            except Exception: pass
            finally: self._updating_rev = False
            self._update_rev_label(); self._update_prev_info(); self._render_preview()

    def rev_inc_click(self, s, e):
        if self._rev_count < 20:
            self._rev_count += 1
            self._updating_rev = True
            try: self.rev_tb.Text = str(self._rev_count)
            except Exception: pass
            finally: self._updating_rev = False
            self._update_rev_label(); self._update_prev_info(); self._render_preview()

    def rev_changed(self, s, e):
        if self._updating_rev: return
        try:
            v = max(1, min(20, int(self.rev_tb.Text)))
            if v != self._rev_count:
                self._rev_count = v
                self._update_rev_label(); self._update_prev_info(); self._render_preview()
        except Exception: pass

    def page_changed(self, s, e):
        try:
            tag = self.page_cb.SelectedItem.Tag if self.page_cb.SelectedItem else 'a4'
            _is_custom = (tag == 'custom')
            _cvis = _SW.Visibility.Visible if _is_custom else _SW.Visibility.Collapsed
            _dvis = _SW.Visibility.Collapsed if _is_custom else _SW.Visibility.Visible
            self.custom_w_tb.Visibility  = _cvis
            self.custom_x_tb.Visibility  = _cvis
            self.custom_h_tb.Visibility  = _cvis
            self.custom_mm_tb.Visibility = _cvis
            self.page_dims_tb.Visibility = _dvis
            if not _is_custom:
                dims = PAGE_SIZES.get(str(tag), (210, 297))
                self._page_w_mm   = dims[0]
                self._page_h_mm   = dims[1]
                self._orientation = 'landscape' if dims[0] > dims[1] else 'portrait'
                self.page_dims_tb.Text = '{}x{}mm'.format(int(self._page_w_mm), int(self._page_h_mm))
            self._update_prev_info(); self._render_preview()
        except Exception: pass

    def _sync_page_cb(self):
        """Update page_cb combobox to match current _page_w_mm/_page_h_mm."""
        _tag_map = {(210,297):'a4', (297,210):'a4l', (297,420):'a3', (420,297):'a3l'}
        _tag = _tag_map.get((int(self._page_w_mm), int(self._page_h_mm)), 'custom')
        self._switching = True
        try:
            for item in self.page_cb.Items:
                if str(item.Tag) == _tag:
                    self.page_cb.SelectedItem = item
                    break
        except Exception: pass
        finally: self._switching = False
        _is_custom = (_tag == 'custom')
        _cvis = _SW.Visibility.Visible if _is_custom else _SW.Visibility.Collapsed
        _dvis = _SW.Visibility.Collapsed if _is_custom else _SW.Visibility.Visible
        try:
            self.custom_w_tb.Visibility  = _cvis
            self.custom_x_tb.Visibility  = _cvis
            self.custom_h_tb.Visibility  = _cvis
            self.custom_mm_tb.Visibility = _cvis
            self.page_dims_tb.Visibility = _dvis
            if _is_custom:
                self.custom_w_tb.Text = str(int(self._page_w_mm))
                self.custom_h_tb.Text = str(int(self._page_h_mm))
            else:
                self.page_dims_tb.Text = '{}x{}mm'.format(int(self._page_w_mm), int(self._page_h_mm))
        except Exception: pass

    def custom_w_changed(self, s, e):
        try:
            v = max(50, min(600, int(self.custom_w_tb.Text)))
            self._page_w_mm = v
            self._orientation = 'landscape' if self._page_w_mm > self._page_h_mm else 'portrait'
            self._update_prev_info(); self._render_preview()
        except Exception: pass

    def custom_h_changed(self, s, e):
        try:
            v = max(50, min(1200, int(self.custom_h_tb.Text)))
            self._page_h_mm = v
            self._orientation = 'landscape' if self._page_w_mm > self._page_h_mm else 'portrait'
            self._update_prev_info(); self._render_preview()
        except Exception: pass

    def _on_columns_changed(self):
        """Called when column percentages change (via drag splitter or mm dialog).
        Updates labels, preview immediately, debounces canvas rebuild."""
        self._update_col_labels()
        self._render_preview()
        self._schedule_canvas_rebuild()

    def col_settings_click(self, s, e):
        """Open dialog to set column widths in %."""
        a_pct, b_pct, c_pct = self._col_pct[0], self._col_pct[1], self._col_pct[2]
        d_pct = max(5, 100 - a_pct - b_pct - c_pct)
        prompt_str = ('Enter column widths as percentages.\n'
                      'Format: A,B,C,D  e.g. "18,35,17,30" (must sum to 100)')
        default = '{},{},{},{}'.format(a_pct, b_pct, c_pct, d_pct)
        result = self._prompt(prompt_str, default=default)
        if not result: return
        try:
            parts = [float(p.strip()) for p in result.split(',')]
            if len(parts) != 4:
                self._status('Need 4 values: A,B,C,D'); return
            total = sum(parts)
            if total <= 0:
                self._status('Invalid widths'); return
            # Normalise to 100% and clamp each to 5+
            a_p = max(5, int(round(parts[0] / total * 100)))
            b_p = max(5, int(round(parts[1] / total * 100)))
            c_p = max(5, int(round(parts[2] / total * 100)))
            # Ensure D >= 5
            if a_p + b_p + c_p > 95:
                excess = (a_p + b_p + c_p) - 95
                a_p -= excess // 3; b_p -= excess // 3; c_p -= excess - 2 * (excess // 3)
            self._col_pct = [a_p, b_p, c_p]
            self._on_columns_changed()
            self._status('Column widths set')
        except Exception as ex:
            self._status('Invalid input: {}'.format(ex))
            self._status('Invalid input: {}'.format(ex))

    def _schedule_canvas_rebuild(self):
        """Coalesce rapid slider changes into a single canvas rebuild after ~80ms of inactivity."""
        import System.Windows.Threading as _SWT
        if getattr(self, '_canvas_debounce_timer', None) is None:
            self._canvas_debounce_timer = _SWT.DispatcherTimer()
            self._canvas_debounce_timer.Interval = System.TimeSpan.FromMilliseconds(80)
            def _tick(s, e):
                try: self._canvas_debounce_timer.Stop()
                except Exception: pass
                self._render_canvas()
            self._canvas_debounce_timer.Tick += _tick
        # Reset the timer — keep extending it until slider stops moving
        try: self._canvas_debounce_timer.Stop()
        except Exception: pass
        self._canvas_debounce_timer.Start()

    def logo_path_changed(self, s, e):
        try:
            self._logo_path = self.logo_path_tb.Text.strip()
            self._render_preview()
        except Exception: pass

    def logo_browse_click(self, s, e):
        try:
            clr.AddReference('Microsoft.Win32')
        except Exception: pass
        try:
            from Microsoft.Win32 import OpenFileDialog
            dlg = OpenFileDialog()
            dlg.Title = 'Select logo file'
            dlg.Filter = 'Image files (*.png;*.jpg;*.jpeg;*.bmp;*.gif)|*.png;*.jpg;*.jpeg;*.bmp;*.gif|All files (*.*)|*.*'
            if self._logo_path:
                try:
                    import os
                    d = os.path.dirname(self._logo_path)
                    if d and os.path.isdir(d): dlg.InitialDirectory = d
                except Exception: pass
            result = dlg.ShowDialog()
            if result:
                self.logo_path_tb.Text = dlg.FileName
        except Exception as ex:
            self._status('Browse failed: {}'.format(ex))

    def template_changed(self, s, e):
        if self._switching: return
        sel = self.template_cb.SelectedItem
        if not sel or str(sel) == self._active_tmpl: return
        self._flush_template()
        self._active_tmpl = str(sel)
        td = self._templates[self._active_tmpl]
        self._rows        = copy.deepcopy(td.get('rows', []))
        self._rev_count   = int(td.get('rev_count', 4))
        self._page_w_mm   = td.get('page_w_mm', 210)
        self._page_h_mm   = td.get('page_h_mm', 297)
        self._orientation = td.get('orientation', 'portrait')
        self._hlines = {int(k): list(v) for k,v in td.get('hlines', {}).items()}
        self._vlines = {int(k): list(v) for k,v in td.get('vlines', {}).items()}
        self._updating_rev = True
        try: self.rev_tb.Text = str(self._rev_count)
        except Exception: pass
        finally: self._updating_rev = False
        self._sync_page_cb()
        self._update_rev_label(); self._update_prev_info(); self._render_canvas(); self._render_preview()

    def add_template_click(self, s, e):
        name = self._prompt('New template name:')
        if not name: return
        if name in self._templates: self._status('Already exists'); return
        self._flush_template()
        self._templates[name] = copy.deepcopy(self._templates[self._active_tmpl])
        self._active_tmpl = name
        self._sync_ui_from_state()
        self._status('Added: ' + name)

    def del_template_click(self, s, e):
        if len(self._templates) <= 1:
            self._status('Cannot delete last template'); return
        name = self._active_tmpl
        confirm = self._prompt('To delete, type the template name exactly:\n"{}"'.format(name), default='')
        if not confirm or confirm.strip() != name:
            self._status('Delete cancelled — name did not match'); return
        del self._templates[name]
        # Remove the template JSON file from disk
        try:
            path = self._template_file(name)
            if os.path.isfile(path): os.remove(path)
        except Exception: pass
        self._active_tmpl = list(self._templates.keys())[0]
        td = self._templates[self._active_tmpl]
        self._rows = copy.deepcopy(td.get('rows', []))
        self._rev_count = int(td.get('rev_count', 4))
        self._sync_ui_from_state()
        self._render_canvas(); self._render_preview()
        self._status('Deleted: ' + name)

    def clear_click(self, s, e):
        self._rows = []
        self._render_canvas(); self._render_preview()

    def save_click(self, s, e):
        self._flush_template()
        self._save_config()
        self._status('Saved: {} template(s)'.format(len(self._templates)))

    def save_as_click(self, s, e):
        """Prompt for a template name. If it exists, confirm overwrite."""
        name = self._prompt('Save current layout as:', default=self._active_tmpl)
        if not name: return
        name = name.strip()
        if not name:
            self._status('Name cannot be blank'); return
        if name in self._templates and name != self._active_tmpl:
            confirm = self._prompt('Template "{}" already exists.\nType YES to overwrite:'.format(name), default='')
            if not confirm or confirm.strip().upper() != 'YES':
                self._status('Save As cancelled'); return
        # Snapshot current state into the target template
        self._flush_template()
        self._templates[name] = copy.deepcopy(self._templates[self._active_tmpl])
        self._active_tmpl = name
        self._sync_ui_from_state()
        self._save_config()
        self._status('Saved as: ' + name)

    def close_click(self, s, e):
        self.Close()

    # ── Palette tabs ──────────────────────────────────────────────
    def pal_tab_blocks(self, s, e):
        self.pal_blocks_scroll.Visibility    = _SW.Visibility.Visible
        self.pal_settings_scroll.Visibility  = _SW.Visibility.Collapsed
        self.pal_inspector_scroll.Visibility = _SW.Visibility.Collapsed
        self.pal_blocks_btn.Background    = BK['accent']
        self.pal_settings_btn.Background  = BK['row']
        self.pal_inspector_btn.Background = BK['row']
        if not getattr(self, '_palette_built', False):
            self._build_palette()
            self._palette_built = True

    def pal_tab_settings(self, s, e):
        self.pal_blocks_scroll.Visibility    = _SW.Visibility.Collapsed
        self.pal_settings_scroll.Visibility  = _SW.Visibility.Visible
        self.pal_inspector_scroll.Visibility = _SW.Visibility.Collapsed
        self.pal_blocks_btn.Background    = BK['row']
        self.pal_settings_btn.Background  = BK['accent']
        self.pal_inspector_btn.Background = BK['row']
        self._render_style_cards()

    def pal_tab_inspector(self, s, e):
        self.pal_blocks_scroll.Visibility    = _SW.Visibility.Collapsed
        self.pal_settings_scroll.Visibility  = _SW.Visibility.Collapsed
        self.pal_inspector_scroll.Visibility = _SW.Visibility.Visible
        self.pal_blocks_btn.Background    = BK['row']
        self.pal_settings_btn.Background  = BK['row']
        self.pal_inspector_btn.Background = BK['accent']

    def _show_inspector(self, ri, ci, block):
        # Populate and switch to the Inspector tab when a block is selected
        self.pal_tab_inspector(None, None)
        sp = self.inspector_panel
        sp.Children.Clear()
        if block is None:
            lbl = _SWC.TextBlock()
            lbl.Text = 'Select a block to inspect'
            lbl.Foreground = BK['muted']; lbl.FontSize = 10
            lbl.Margin = _SW.Thickness(0, 12, 0, 0)
            lbl.HorizontalAlignment = _SW.HorizontalAlignment.Center
            sp.Children.Add(lbl)
            return

        # Block title
        hdr = _SWC.TextBlock()
        hdr.Text = block.get('type','').replace('_',' ').title()
        hdr.Foreground = BK['accent']; hdr.FontSize = 11; hdr.FontWeight = _SW.FontWeights.Bold
        hdr.Margin = _SW.Thickness(0,0,0,8)
        sp.Children.Add(hdr)

        # Edit text button for text blocks
        if block.get('type') in ('text','title','heading'):
            edit_btn = _SWC.Button()
            edit_btn.Content = 'Edit Text'
            edit_btn.Height = 24; edit_btn.FontSize = 10
            edit_btn.Background = BK['row']; edit_btn.Foreground = BK['text']
            edit_btn.BorderBrush = BK['bdr']; edit_btn.BorderThickness = _SW.Thickness(1)
            edit_btn.Margin = _SW.Thickness(0,0,0,8); edit_btn.Cursor = _SWI.Cursors.Hand
            edit_btn.Click += lambda s,e,r=ri,c=ci: self._edit_text_block(r,c)
            sp.Children.Add(edit_btn)

        # ── Hint about border editing ─────────────────────────────
        hint = _SWC.TextBlock()
        hint.Text = 'Click the white lines on the canvas to set borders'
        hint.Foreground = BK['muted']; hint.FontSize = 9
        hint.Margin = _SW.Thickness(0, 0, 0, 8)
        sp.Children.Add(hint)

        # ── Block settings ────────────────────────────────────────
        settings_panel = self._make_settings_panel(ri, ci, block)
        sp.Children.Add(settings_panel)

    # ── Palette builder ───────────────────────────────────────────
    def _build_palette(self):
        """Build blocks into palette_stack using the wide item layout that
        properly stretches to scrollviewer width (proven via Settings 3)."""
        stack = self.palette_stack
        stack.Children.Clear()
        for (tid, label, icon, grp, sub) in PALETTE:
            if tid == '__grp__':
                hdr = _SWC.TextBlock()
                hdr.Text = label.upper()
                hdr.Foreground = _SWM.SolidColorBrush(GROUP_COLORS.get(grp, GROUP_COLORS['layout']))
                hdr.FontSize = 9
                hdr.FontWeight = _SW.FontWeights.Bold
                hdr.Margin = _SW.Thickness(0, 8, 0, 3)
                stack.Children.Add(hdr)
                continue
            item = self._make_palette2_item(tid, label, icon, grp, sub)
            stack.Children.Add(item)

    def _make_palette2_item(self, tid, label, icon, grp, sub):
        """Edge-to-edge palette card — mirrors _make_style_card structure:
        Border with BgCard background, CornerRadius=6, Bdr border, Padding=10, Margin bottom=8.
        Inner StackPanel with a header line and optional muted subtitle.
        """
        grp_color = _SWM.SolidColorBrush(GROUP_COLORS.get(grp, GROUP_COLORS['layout']))

        outer = _SWC.Border()
        outer.Tag = tid
        outer.Background = BK['card']
        outer.CornerRadius = _SW.CornerRadius(6)
        outer.BorderBrush = BK['bdr']
        outer.BorderThickness = _SW.Thickness(1)
        outer.Padding = _SW.Thickness(10)
        outer.Margin = _SW.Thickness(0, 0, 0, 8)
        outer.Cursor = _SWI.Cursors.Hand
        outer.AllowDrop = False
        outer.HorizontalAlignment = _SW.HorizontalAlignment.Stretch

        sp = _SWC.StackPanel()
        outer.Child = sp

        # Header row (icon + label) — matches style-card header (Foreground=Accent, FontSize=11, SemiBold)
        hdr_tb = _SWC.TextBlock()
        hdr_tb.Text = (icon or '') + '  ' + label
        hdr_tb.Foreground = grp_color
        hdr_tb.FontSize = 11
        hdr_tb.FontWeight = _SW.FontWeights.SemiBold
        sp.Children.Add(hdr_tb)

        if sub:
            sub_tb = _SWC.TextBlock()
            sub_tb.Text = sub
            sub_tb.Foreground = BK['muted']
            sub_tb.FontSize = 9
            sub_tb.Margin = _SW.Thickness(0, 4, 0, 0)
            sp.Children.Add(sub_tb)

        # Hover: highlight border with the group colour
        def on_enter(s, e, gc=grp_color):
            s.BorderBrush = gc
        def on_leave(s, e):
            s.BorderBrush = BK['bdr']
        outer.MouseEnter += on_enter
        outer.MouseLeave += on_leave

        # Drag-and-drop wiring
        outer.PreviewMouseLeftButtonDown += self._pal_mouse_down
        outer.MouseMove += self._pal_mouse_move

        return outer

    def _pal_mouse_down(self, sender, args):
        try:
            self._pal_drag_start = args.GetPosition(sender)
            self._pal_drag_source = sender
        except Exception:
            self._pal_drag_start = None

    def _pal_mouse_move(self, sender, args):
        try:
            import System.Windows.Input as _WI
            if args.LeftButton != _WI.MouseButtonState.Pressed: return
            if not hasattr(self, '_pal_drag_start') or self._pal_drag_start is None: return
            pos = args.GetPosition(sender)
            dx = abs(pos.X - self._pal_drag_start.X)
            dy = abs(pos.Y - self._pal_drag_start.Y)
            if dx < 4 and dy < 4: return
            self._pal_drag_start = None
            tid = str(sender.Tag)
            self._drag_type = tid  # fallback if DataObject fails
            data = _SW.DataObject()
            data.SetData('pal_type', tid)
            # Use Copy|Move so any target can accept
            result = _SW.DragDrop.DoDragDrop(
                sender, data,
                _SW.DragDropEffects.Copy | _SW.DragDropEffects.Move)
            self._drag_type = None
        except Exception:
            self._drag_type = None

    # ── Canvas rows ───────────────────────────────────────────────
    def add_row_click(self, s, e):
        self._rows.append({'blocks': [None, None, None, None], 'merge_down': False, 'section': 'body'})
        self._render_canvas(); self._render_preview()

    def _get_group(self, ri):
        """Return (start, end) indices of the merged group containing row ri.
        end is inclusive. A standalone row returns (ri, ri)."""
        # Walk up to find group start
        start = ri
        while start > 0 and self._rows[start - 1].get('merge_down', False):
            start -= 1
        # Walk down to find group end
        end = ri
        while end < len(self._rows) - 1 and self._rows[end].get('merge_down', False):
            end += 1
        return (start, end)

    def _toggle_merge_down(self, ri):
        """Toggle merge_down on this row (links it with the row below)."""
        if ri >= len(self._rows) - 1:
            self._status('Cannot merge last row down'); return
        self._rows[ri]['merge_down'] = not self._rows[ri].get('merge_down', False)
        self._render_canvas(); self._render_preview()

    def _cycle_section(self, ri):
        """Cycle the row's section tag: body → repeat_header → footer → body.
        For merged groups, applies to ALL rows in the group."""
        order = ['body', 'repeat_header', 'footer']
        start, end = self._get_group(ri)
        current = self._rows[start].get('section', 'body')
        idx = order.index(current) if current in order else 0
        new_section = order[(idx + 1) % len(order)]
        # Apply to all rows in the group
        for r in range(start, end + 1):
            self._rows[r]['section'] = new_section
        self._render_canvas()
        labels = {'body': 'Body', 'repeat_header': 'Repeat Header', 'footer': 'Footer'}
        self._status('Section: {}'.format(labels.get(new_section, new_section)))

    def _remove_row(self, ri):
        start, end = self._get_group(ri)
        # Remove all rows in the group (bottom-up to preserve indices)
        for i in range(end, start - 1, -1):
            del self._rows[i]
        self._render_canvas(); self._render_preview()

    def _move_row_up(self, ri):
        start, end = self._get_group(ri)
        if start <= 0: return
        # Find the group above and jump over it entirely
        above_start, above_end = self._get_group(start - 1)
        # Extract our group, insert before the above group
        group = self._rows[start:end + 1]
        del self._rows[start:end + 1]
        for i, row in enumerate(group):
            self._rows.insert(above_start + i, row)
        self._render_canvas(); self._render_preview()

    def _move_row_down(self, ri):
        start, end = self._get_group(ri)
        if end >= len(self._rows) - 1: return
        # Find the group below and jump over it entirely
        below_start, below_end = self._get_group(end + 1)
        # Extract our group, insert after the below group
        group = self._rows[start:end + 1]
        del self._rows[start:end + 1]
        # After deletion, below_end shifted by the number of rows removed
        insert_at = below_end - (end - start)
        for i, row in enumerate(group):
            self._rows.insert(insert_at + i, row)
        self._render_canvas(); self._render_preview()

    # ── Block operations ──────────────────────────────────────────
    def _occupied(self, ri):
        occ = set()
        for i, b in enumerate(self._rows[ri]['blocks']):
            if b and b.get('span', 1) > 1:
                for s in range(1, b['span']):
                    if i + s < 4: occ.add(i + s)
        return occ

    def _v_occupied(self, ri, ci):
        """Check if column ci at row ri is vertically occupied by a block
        in a row above that has row_span > 1. Returns the owning (row_idx, block) or None."""
        start, end = self._get_group(ri)
        for check_ri in range(start, ri):
            b = self._rows[check_ri]['blocks'][ci]
            if b and b.get('row_span', 1) > 1:
                # This block spans from check_ri to check_ri + row_span - 1
                if ri <= check_ri + b.get('row_span', 1) - 1:
                    return (check_ri, b)
        return None

    def _max_row_span(self, ri, ci):
        """Max row_span for a block at (ri, ci) within its merged group."""
        start, end = self._get_group(ri)
        mx = 1
        for s in range(1, end - ri + 1):
            target_ri = ri + s
            # Check if this row's column is free (no block and not v-occupied by another)
            tb = self._rows[target_ri]['blocks'][ci]
            if tb is not None:
                break  # occupied by another block
            v_owner = self._v_occupied(target_ri, ci)
            if v_owner and v_owner[0] != ri:
                break  # occupied by a different block's row span
            mx = s + 1
        return mx

    def _increase_row_span(self, ri, ci):
        b = self._rows[ri]['blocks'][ci]
        if not b: return
        mx = self._max_row_span(ri, ci)
        if b.get('row_span', 1) >= mx: return
        b['row_span'] = b.get('row_span', 1) + 1
        self._render_canvas(); self._render_preview()

    def _decrease_row_span(self, ri, ci):
        b = self._rows[ri]['blocks'][ci]
        if not b: return
        if b.get('row_span', 1) <= 1: return
        b['row_span'] = b.get('row_span', 1) - 1
        self._render_canvas(); self._render_preview()

    def _max_span(self, ri, ci):
        occ = self._occupied(ri)
        mx = 1
        for s in range(1, 4 - ci):
            nb = self._rows[ri]['blocks'][ci + s]
            if nb is not None and (ci + s) not in occ: break
            mx = s + 1
        return min(mx, 4 - ci)

    def _increase_span(self, ri, ci):
        b = self._rows[ri]['blocks'][ci]
        if not b: return
        mx = self._max_span(ri, ci)
        if b.get('span', 1) >= mx: return
        ns = ci + b.get('span', 1)
        if ns < 4: self._rows[ri]['blocks'][ns] = None
        b['span'] = b.get('span', 1) + 1
        self._render_canvas(); self._render_preview()

    def _decrease_span(self, ri, ci):
        b = self._rows[ri]['blocks'][ci]
        if not b or b.get('span', 1) <= 1: return
        b['span'] -= 1
        self._render_canvas(); self._render_preview()

    def _place_block(self, ri, ci, btype):
        occ = self._occupied(ri)
        if ci in occ: self._status('Slot occupied by spanning block'); return
        if btype in ('title', 'heading', 'text'):
            content = self._prompt('Content for {} block:'.format(TYPE_NAMES.get(btype, btype)))
            if content is None: return
            ts = 'Title' if btype == 'title' else ('Header' if btype == 'heading' else 'Data')
            block = _mk(btype, content=content, text_style=ts,
                        just='center' if btype=='title' else 'left',
                        borders={'t':False,'b':True,'l':False,'r':False})
        else:
            block = _mk(btype)
        block['_id'] = self._nid()
        self._rows[ri]['blocks'][ci] = block
        self._render_canvas(); self._render_preview()
        self._status('Added: ' + TYPE_NAMES.get(btype, btype))

    def _remove_block(self, ri, ci):
        self._rows[ri]['blocks'][ci] = None
        self._render_canvas(); self._render_preview()

    def _toggle_block(self, ri, ci):
        b = self._rows[ri]['blocks'][ci]
        if b: b['enabled'] = not b.get('enabled', True); self._render_canvas(); self._render_preview()

    def _move_block(self, fr, fc, tr, tc):
        mv = self._rows[fr]['blocks'][fc]
        dis = self._rows[tr]['blocks'][tc]
        if mv: mv['span'] = 1
        if dis: dis['span'] = 1
        self._rows[tr]['blocks'][tc] = mv
        self._rows[fr]['blocks'][fc] = dis
        self._render_canvas(); self._render_preview()

    def _update_block_label(self, ri, ci, val):
        b = self._rows[ri]['blocks'][ci]
        if b: b['label'] = val; self._render_preview()

    def _edit_text_block(self, ri, ci):
        b = self._rows[ri]['blocks'][ci]
        if not b: return
        new_content = self._prompt('Edit content:', default=b.get('content', ''))
        if new_content is not None:
            b['content'] = new_content
            self._render_canvas()
            self._render_preview()

    # ── Canvas render ─────────────────────────────────────────────
    def _render_canvas(self):
        stack = self.canvas_stack
        stack.Children.Clear()

        if not self._rows:
            tb = _SWC.TextBlock()
            tb.Text = 'Add a row, then drag blocks from the palette into slots'
            tb.Foreground = BK['muted']; tb.FontSize = 11
            tb.HorizontalAlignment = _SW.HorizontalAlignment.Center
            tb.Margin = _SW.Thickness(0, 20, 0, 0)
            stack.Children.Add(tb)
            try: self.count_tb.Text = '0 rows'
            except Exception: pass
            return

        d = max(5, 100 - self._col_pct[0] - self._col_pct[1] - self._col_pct[2])
        fracs = self._col_pct + [d]

        ri = 0
        while ri < len(self._rows):
            start, end = self._get_group(ri)
            is_group = (end > start)

            if is_group:
                # Determine group section and colour
                section = self._rows[start].get('section', 'body')
                sec_border_colors = {
                    'body':          BK['accent'],
                    'repeat_header': _SWM.SolidColorBrush(_SWM.Color.FromRgb(0x29, 0x80, 0xB9)),
                    'footer':        _SWM.SolidColorBrush(_SWM.Color.FromRgb(0x8E, 0x44, 0xAD)),
                }
                sec_label_text = {'body': '', 'repeat_header': 'REPEAT HEADER', 'footer': 'FOOTER'}

                group_border = _SWC.Border()
                group_border.BorderBrush = sec_border_colors.get(section, BK['accent'])
                group_border.BorderThickness = _SW.Thickness(2)
                group_border.CornerRadius = _SW.CornerRadius(4)
                group_border.Margin = _SW.Thickness(0, 0, 0, 4)
                group_sp = _SWC.StackPanel()
                group_border.Child = group_sp

                # Section label at top of group (if not body)
                if sec_label_text.get(section):
                    sec_lbl = _SWC.TextBlock()
                    sec_lbl.Text = sec_label_text[section]
                    sec_lbl.Foreground = sec_border_colors.get(section, BK['accent'])
                    sec_lbl.FontSize = 8
                    sec_lbl.FontWeight = _SW.FontWeights.Bold
                    sec_lbl.Margin = _SW.Thickness(6, 2, 0, 0)
                    group_sp.Children.Add(sec_lbl)

                for gri in range(start, end + 1):
                    is_top = (gri == start)
                    if is_top and start == 0:
                        group_sp.Children.Add(self._make_hline_row(gri, fracs, above=True))
                    row_el = self._make_canvas_row(gri, self._rows[gri], fracs,
                                                   show_controls=is_top, in_group=True)
                    group_sp.Children.Add(row_el)
                    group_sp.Children.Add(self._make_hline_row(gri, fracs, above=False))

                stack.Children.Add(group_border)
            else:
                row_el = self._make_canvas_row(ri, self._rows[ri], fracs,
                                               show_controls=True, in_group=False)
                # Hline above first row only
                if ri == 0:
                    stack.Children.Add(self._make_hline_row(ri, fracs, above=True))
                section = self._rows[ri].get('section', 'body')
                if section != 'body':
                    # Wrap standalone row in a section-coloured border
                    sec_border_colors = {
                        'repeat_header': _SWM.SolidColorBrush(_SWM.Color.FromRgb(0x29, 0x80, 0xB9)),
                        'footer':        _SWM.SolidColorBrush(_SWM.Color.FromRgb(0x8E, 0x44, 0xAD)),
                    }
                    sec_label_text = {'repeat_header': 'REPEAT HEADER', 'footer': 'FOOTER'}
                    wrap = _SWC.Border()
                    wrap.BorderBrush = sec_border_colors.get(section, BK['accent'])
                    wrap.BorderThickness = _SW.Thickness(2)
                    wrap.CornerRadius = _SW.CornerRadius(4)
                    wrap.Margin = _SW.Thickness(0, 0, 0, 4)
                    wrap_sp = _SWC.StackPanel()
                    lbl = _SWC.TextBlock()
                    lbl.Text = sec_label_text.get(section, '')
                    lbl.Foreground = sec_border_colors.get(section, BK['accent'])
                    lbl.FontSize = 8; lbl.FontWeight = _SW.FontWeights.Bold
                    lbl.Margin = _SW.Thickness(6, 2, 0, 0)
                    wrap_sp.Children.Add(lbl)
                    wrap_sp.Children.Add(row_el)
                    wrap.Child = wrap_sp
                    stack.Children.Add(wrap)
                else:
                    stack.Children.Add(row_el)
                # Hline below this row (bottom edge)
                stack.Children.Add(self._make_hline_row(ri, fracs, above=False))

            ri = end + 1

        try:
            n = len(self._rows)
            self.count_tb.Text = '{} row{}'.format(n, 's' if n!=1 else '')
        except Exception: pass

    def _make_hline_row(self, ri, fracs, above=False):
        # Horizontal separator line above (above=True) or below (above=False) row ri.
        # When above=True: this is the TOP edge of row ri = BOTTOM edge of row ri-1
        # The hline data key: bottom of row (ri-1) when above, or bottom of row ri when below.
        hline_ri = ri if above else ri + 1
        # Segments match the block spans in the reference row
        ref_ri = ri  # use current row's spans for segment layout
        HLINE_H = 10  # 3px pad + 4px line + 3px pad
        CTRL_W  = 22  # approximate right-side controls width

        occ = self._occupied(ri)
        row = self._rows[ri]

        outer = _SWC.Grid()
        outer.Height = HLINE_H

        # Right spacer to match controls area
        cd_main = _SWC.ColumnDefinition(); cd_main.Width = _SW.GridLength(1, _SW.GridUnitType.Star)
        cd_ctrl = _SWC.ColumnDefinition(); cd_ctrl.Width = _SW.GridLength(CTRL_W)
        outer.ColumnDefinitions.Add(cd_main); outer.ColumnDefinitions.Add(cd_ctrl)

        # Inner segment grid
        segs_g = _SWC.Grid()
        segs_g.VerticalAlignment = _SW.VerticalAlignment.Stretch
        _SWC.Grid.SetColumn(segs_g, 0)

        ci = 0
        seg_idx = 0
        while ci < 4:
            if ci in occ: ci+=1; continue
            b = row['blocks'][ci]
            span = b.get('span', 1) if b else 1

            total_frac = sum(fracs[min(ci+ss, 3)] for ss in range(span))
            cd = _SWC.ColumnDefinition()
            cd.Width = _SW.GridLength(total_frac, _SW.GridUnitType.Star)
            segs_g.ColumnDefinitions.Add(cd)

            is_on = self._get_hline(hline_ri, ci)

            # Check if a row_span block covers this hline segment
            _vspan_locked = False
            for _check_ri in range(max(0, ri - 6), ri + 2):
                if _check_ri >= len(self._rows): continue
                _rblocks = self._rows[_check_ri].get('blocks', [])
                _cb = _rblocks[ci] if ci < len(_rblocks) else None
                if _cb and _cb.get('row_span', 1) > 1:
                    _span_end_ri = _check_ri + _cb['row_span'] - 1
                    if _check_ri < hline_ri <= _span_end_ri:
                        _vspan_locked = True
                        break

            seg = _SWC.Border()
            seg.VerticalAlignment = _SW.VerticalAlignment.Stretch
            seg.Background = _SWM.Brushes.Transparent
            seg.Cursor = _SWI.Cursors.Arrow if _vspan_locked else _SWI.Cursors.Hand

            inner = _SWC.Border()
            inner.Height = 4; inner.CornerRadius = _SW.CornerRadius(2)
            inner.VerticalAlignment = _SW.VerticalAlignment.Center
            inner.HorizontalAlignment = _SW.HorizontalAlignment.Stretch
            inner.Margin = _SW.Thickness(3, 3, 3, 3)
            inner.IsHitTestVisible = False
            inner.Background = _SWM.SolidColorBrush(
                _SWM.Color.FromArgb(0 if _vspan_locked else (230 if is_on else 18), 255, 255, 255))
            seg.Child = inner

            if not _vspan_locked:
                def on_henter(ss, ee, il=inner, r=ri, c=ci):
                    il.Background = _SWM.SolidColorBrush(_SWM.Color.FromArgb(140, 255, 255, 255))
                def on_hleave(ss, ee, il=inner, r=hline_ri, c=ci):
                    cur = self._get_hline(r, c)
                    il.Background = _SWM.SolidColorBrush(
                        _SWM.Color.FromArgb(230 if cur else 18, 255, 255, 255))
                def on_hclick(ss, ee, r=hline_ri, c=ci, sp=span):
                    self._toggle_hline(r, c, sp)
                seg.MouseEnter += on_henter
                seg.MouseLeave += on_hleave
                seg.MouseLeftButtonUp += on_hclick

            _SWC.Grid.SetColumn(seg, seg_idx)
            segs_g.Children.Add(seg)
            seg_idx += 1
            ci += span

        outer.Children.Add(segs_g)
        return outer

    def _make_canvas_row(self, ri, row, fracs, show_controls=True, in_group=False):
        occ = self._occupied(ri)

        # Single-column outer — controls float right via negative margin
        outer = _SWC.Border()
        outer.Margin = _SW.Thickness(0, 0, 0, 2 if in_group else 4)
        outer.MinHeight = 80  # matches cell Height for uniform row height

        # Use DockPanel: controls dock right, slots fill remaining width
        dock = _SWC.DockPanel()
        dock.LastChildFill = True
        dock.VerticalAlignment = _SW.VerticalAlignment.Stretch
        outer.Child = dock

        # Row controls docked to the right — fixed width, no overlap
        ctrl = _SWC.StackPanel()
        ctrl.VerticalAlignment = _SW.VerticalAlignment.Center
        ctrl.Margin = _SW.Thickness(2, 0, 0, 0)
        _SWC.DockPanel.SetDock(ctrl, _SWC.Dock.Right)

        if show_controls:
            for lbl, fn in [('↑', lambda s,e,r=ri: self._move_row_up(r)),
                            ('↓', lambda s,e,r=ri: self._move_row_down(r)),
                            ('✕', lambda s,e,r=ri: self._remove_row(r))]:
                btn = _SWC.Border()
                btn.Width = 18; btn.Height = 18
                btn.Background = BK['card']
                btn.BorderBrush = BK['bdr']
                btn.BorderThickness = _SW.Thickness(1)
                btn.CornerRadius = _SW.CornerRadius(2)
                btn.Margin = _SW.Thickness(0, 1, 0, 1)
                btn.Cursor = _SWI.Cursors.Hand

                tb_lbl = _SWC.TextBlock()
                tb_lbl.Text = lbl
                tb_lbl.Foreground = BK['white']
                tb_lbl.FontSize = 9
                tb_lbl.HorizontalAlignment = _SW.HorizontalAlignment.Center
                tb_lbl.VerticalAlignment = _SW.VerticalAlignment.Center
                btn.Child = tb_lbl

                def on_enter(s, e):
                    s.BorderBrush = BK['accent']
                def on_leave(s, e):
                    s.BorderBrush = BK['bdr']
                btn.MouseEnter += on_enter
                btn.MouseLeave += on_leave
                btn.MouseLeftButtonUp += fn
                ctrl.Children.Add(btn)

            # Section tag button — only on top row of group (or standalone)
            section = row.get('section', 'body')
            sec_labels = {'body': 'B', 'repeat_header': 'R', 'footer': 'F'}
            sec_tips = {'body': 'Body (normal content)',
                        'repeat_header': 'Repeat Header (prints on every page)',
                        'footer': 'Footer (page bottom)'}
            sec_colors = {'body': BK['card'],
                          'repeat_header': _SWM.SolidColorBrush(_SWM.Color.FromRgb(0x29, 0x80, 0xB9)),
                          'footer': _SWM.SolidColorBrush(_SWM.Color.FromRgb(0x8E, 0x44, 0xAD))}
            sec_btn = _SWC.Border()
            sec_btn.Width = 18; sec_btn.Height = 18
            sec_btn.Background = sec_colors.get(section, BK['card'])
            sec_btn.BorderBrush = BK['bdr']
            sec_btn.BorderThickness = _SW.Thickness(1)
            sec_btn.CornerRadius = _SW.CornerRadius(2)
            sec_btn.Margin = _SW.Thickness(0, 1, 0, 1)
            sec_btn.Cursor = _SWI.Cursors.Hand
            sec_btn.ToolTip = sec_tips.get(section, 'Body')

            sec_tb = _SWC.TextBlock()
            sec_tb.Text = sec_labels.get(section, 'B')
            sec_tb.Foreground = BK['white']
            sec_tb.FontSize = 9; sec_tb.FontWeight = _SW.FontWeights.Bold
            sec_tb.HorizontalAlignment = _SW.HorizontalAlignment.Center
            sec_tb.VerticalAlignment = _SW.VerticalAlignment.Center
            sec_btn.Child = sec_tb
            sec_btn.MouseLeftButtonUp += lambda s,e,r=ri: self._cycle_section(r)
            ctrl.Children.Add(sec_btn)

        # Merge button — always shown (merge this row with the one below)
        is_merged = row.get('merge_down', False)
        merge_btn = _SWC.Border()
        merge_btn.Width = 18; merge_btn.Height = 18
        merge_btn.Background = BK['accent'] if is_merged else BK['card']
        merge_btn.BorderBrush = BK['accent'] if is_merged else BK['bdr']
        merge_btn.BorderThickness = _SW.Thickness(1)
        merge_btn.CornerRadius = _SW.CornerRadius(2)
        merge_btn.Margin = _SW.Thickness(0, 1, 0, 1)
        merge_btn.Cursor = _SWI.Cursors.Hand
        merge_btn.ToolTip = 'Unmerge' if is_merged else 'Merge with row below'

        merge_tb = _SWC.TextBlock()
        merge_tb.Text = '⛓' if is_merged else '🔗'
        merge_tb.Foreground = BK['white']
        merge_tb.FontSize = 9
        merge_tb.HorizontalAlignment = _SW.HorizontalAlignment.Center
        merge_tb.VerticalAlignment = _SW.VerticalAlignment.Center
        merge_btn.Child = merge_tb
        merge_btn.MouseLeftButtonUp += lambda s,e,r=ri: self._toggle_merge_down(r)
        ctrl.Children.Add(merge_btn)

        dock.Children.Add(ctrl)

        # Slot grid with interleaved vertical separator lines
        # Layout: vline(0) | slot | vline(1) | slot | vline(2) | slot | vline(3) | slot | vline(4)
        VLINE_W = 10  # 3px pad + 4px line + 3px pad
        slots_g = _SWC.Grid()
        slots_g.AllowDrop = True
        slots_g.VerticalAlignment = _SW.VerticalAlignment.Stretch

        # Build locked set — vline positions interior to a span are locked
        locked_vpos = set()
        for ci2 in range(4):
            if ci2 in occ: continue
            b2 = row['blocks'][ci2]
            if not b2: continue
            span2 = b2.get('span', 1)
            for interior in range(ci2+1, ci2+span2):
                locked_vpos.add(interior)  # vline position = col index to its left

        # Column definitions: alternating vline cols and slot cols
        # vline positions: 0=outer-left, 1=A-B, 2=B-C, 3=C-D, 4=outer-right
        # col def order: vline0, slot0, vline1, slot1, vline2, slot2, vline3, slot3, vline4
        grid_col_map = {}  # ci -> grid column index
        gc = 0
        # outer left vline
        cd = _SWC.ColumnDefinition(); cd.Width = _SW.GridLength(VLINE_W); slots_g.ColumnDefinitions.Add(cd)
        gc += 1
        for ci in range(4):
            if ci in occ: continue
            b = row['blocks'][ci]
            span = b.get('span', 1) if b else 1
            total_frac = sum(fracs[min(ci+ss, 3)] for ss in range(span))
            cd = _SWC.ColumnDefinition()
            cd.Width = _SW.GridLength(total_frac, _SW.GridUnitType.Star)
            slots_g.ColumnDefinitions.Add(cd)
            grid_col_map[ci] = gc; gc += 1
            # vline after this slot (positions 1-4)
            vpos = ci + span  # vline position to the right of this span
            cd2 = _SWC.ColumnDefinition(); cd2.Width = _SW.GridLength(VLINE_W); slots_g.ColumnDefinitions.Add(cd2)
            gc += 1

        # Add vline borders
        gc = 0
        vpos_grid_cols = {}  # vpos -> grid column index
        vpos_grid_cols[0] = 0  # outer left
        vci = 0
        for ci in range(4):
            if ci in occ: continue
            b = row['blocks'][ci]
            span = b.get('span', 1) if b else 1
            slot_gc = grid_col_map[ci]
            vpos_after = ci + span
            vpos_grid_cols[vpos_after] = slot_gc + 1  # vline grid col after slot

        for vpos, vgc in vpos_grid_cols.items():
            is_locked = vpos in locked_vpos
            is_on = self._get_vline(ri, vpos)
            vline_el = _SWC.Border()
            vline_el.Width = VLINE_W
            vline_el.VerticalAlignment = _SW.VerticalAlignment.Stretch
            vline_el.Background = _SWM.Brushes.Transparent
            vline_el.Cursor = _SWI.Cursors.Arrow if is_locked else _SWI.Cursors.Hand
            # Inner line
            inner = _SWC.Border()
            inner.Width = 4; inner.CornerRadius = _SW.CornerRadius(2)
            inner.VerticalAlignment = _SW.VerticalAlignment.Stretch
            inner.HorizontalAlignment = _SW.HorizontalAlignment.Center
            inner.Margin = _SW.Thickness(3, 0, 3, 0)
            inner.IsHitTestVisible = False
            if is_locked:
                inner.Background = _SWM.SolidColorBrush(_SWM.Color.FromArgb(5, 255, 255, 255))
            elif is_on:
                inner.Background = _SWM.SolidColorBrush(_SWM.Color.FromArgb(230, 255, 255, 255))
            else:
                inner.Background = _SWM.SolidColorBrush(_SWM.Color.FromArgb(18, 255, 255, 255))
            vline_el.Child = inner
            if not is_locked:
                def on_venter(ss, ee, il=inner, on=is_on):
                    il.Background = _SWM.SolidColorBrush(_SWM.Color.FromArgb(140, 255, 255, 255))
                def on_vleave(ss, ee, il=inner, r=ri, vp=vpos):
                    cur = self._get_vline(r, vp)
                    il.Background = _SWM.SolidColorBrush(
                        _SWM.Color.FromArgb(230 if cur else 18, 255, 255, 255))
                def on_vclick(ss, ee, r=ri, vp=vpos):
                    self._toggle_vline(r, vp)
                vline_el.MouseEnter += on_venter
                vline_el.MouseLeave += on_vleave
                vline_el.MouseLeftButtonUp += on_vclick
            _SWC.Grid.SetColumn(vline_el, vgc)
            slots_g.Children.Add(vline_el)

        # Add cells
        for ci in range(4):
            if ci in occ: continue
            b = row['blocks'][ci]
            mx = self._max_span(ri, ci)
            cell = self._make_cell(ri, ci, b, mx, in_group=in_group)
            _SWC.Grid.SetColumn(cell, grid_col_map[ci])
            slots_g.Children.Add(cell)

        dock.Children.Add(slots_g)
        return outer

    def _make_cell(self, ri, ci, block, mx, in_group=False):
        # Check if this cell is vertically occupied by a block above
        v_owner = self._v_occupied(ri, ci) if in_group else None
        if v_owner:
            # Cell is occupied by a row-spanning block above — show a greyed-out indicator
            wrapper = _SWC.Border()
            wrapper.CornerRadius = _SW.CornerRadius(5)
            wrapper.Margin = _SW.Thickness(0, 0, 3, 0)
            wrapper.ClipToBounds = True
            wrapper.AllowDrop = False
            wrapper.VerticalAlignment = _SW.VerticalAlignment.Stretch
            wrapper.HorizontalAlignment = _SW.HorizontalAlignment.Stretch

            cell = _SWC.Border()
            cell.Height = 80
            cell.Background = BK['deep']
            cell.BorderBrush = BK['accent']
            cell.BorderThickness = _SW.Thickness(1)
            cell.Opacity = 0.4
            cell.CornerRadius = _SW.CornerRadius(3)
            lbl = _SWC.TextBlock()
            lbl.Text = '↕'
            lbl.Foreground = BK['accent']; lbl.FontSize = 12
            lbl.HorizontalAlignment = _SW.HorizontalAlignment.Center
            lbl.VerticalAlignment   = _SW.VerticalAlignment.Center
            cell.Child = lbl
            wrapper.Child = cell
            return wrapper

        # Outer wrapper (cell + settings panel stacked vertically)
        wrapper = _SWC.Border()
        wrapper.CornerRadius = _SW.CornerRadius(5)
        wrapper.Margin = _SW.Thickness(0, 0, 3, 0)
        wrapper.ClipToBounds = True
        # Wire drop events on the wrapper so palette drags are caught
        # even when the drag enters the cell's child elements
        wrapper.AllowDrop = True
        # PreviewDragOver tunnels down — set effects here to allow drop
        wrapper.PreviewDragOver += lambda s, e, r=ri, c=ci: self._cell_drag_over(s, e, r, c)
        wrapper.PreviewDragLeave += lambda s, e, r=ri, c=ci: self._cell_drag_leave(s, e, r, c)
        # Use bubbling Drop (not PreviewDrop) — fires after tunneling completes
        wrapper.Drop += lambda s, e, r=ri, c=ci: self._cell_drop(s, e, r, c)

        wrapper.VerticalAlignment = _SW.VerticalAlignment.Stretch
        wrapper.HorizontalAlignment = _SW.HorizontalAlignment.Stretch

        cell_sp = _SWC.StackPanel()
        cell_sp.VerticalAlignment = _SW.VerticalAlignment.Stretch
        wrapper.Child = cell_sp

        # ── Top part: the block cell itself ──────────────────────
        cell = _SWC.Border()
        cell.Height = 80  # fixed — all canvas blocks (filled and empty) identical height
        cell.VerticalAlignment = _SW.VerticalAlignment.Stretch
        cell.HorizontalAlignment = _SW.HorizontalAlignment.Stretch
        cell.Cursor = _SWI.Cursors.Arrow

        if not block:
            cell.Background = BK['card']
            cell.BorderBrush = BK['bdr']
            cell.BorderThickness = _SW.Thickness(1)
            cell.Opacity = 0.5
            lbl = _SWC.TextBlock()
            lbl.Text = ['A','B','C','D'][ci]
            lbl.Foreground = BK['bdr']; lbl.FontSize = 9
            lbl.HorizontalAlignment = _SW.HorizontalAlignment.Center
            lbl.VerticalAlignment   = _SW.VerticalAlignment.Center
            cell.Child = lbl
            cell_sp.Children.Add(cell)
            return wrapper

        # Filled block
        col_br = SLOT_BRUSHES[ci]
        cell.Background = BK['card']
        cell.BorderBrush = col_br
        cell.BorderThickness = _SW.Thickness(1, 2, 1, 1)
        if not block.get('enabled', True): cell.Opacity = 0.4

        # Inner grid: handle | content | actions
        ig = _SWC.Grid()
        for w in (16, None, 22):
            cd = _SWC.ColumnDefinition()
            cd.Width = _SW.GridLength(w) if w else _SW.GridLength(1, _SW.GridUnitType.Star)
            ig.ColumnDefinitions.Add(cd)

        # Handle
        handle = _SWC.Border()
        handle.Background = BK['row']; handle.CornerRadius = _SW.CornerRadius(4,0,0,4)
        handle.BorderBrush = BK['bdr']; handle.BorderThickness = _SW.Thickness(0,0,1,0)
        handle.Cursor = _SWI.Cursors.SizeAll
        handle_tb = _SWC.TextBlock()
        handle_tb.Text = '≡'; handle_tb.Foreground = BK['muted']; handle_tb.FontSize = 12
        handle_tb.HorizontalAlignment = _SW.HorizontalAlignment.Center
        handle_tb.VerticalAlignment   = _SW.VerticalAlignment.Center
        handle.Child = handle_tb
        handle.MouseMove += lambda s,e,r=ri,c=ci: self._block_drag_start(s,e,r,c)
        _SWC.Grid.SetColumn(handle, 0); ig.Children.Add(handle)

        # Content
        content_sp = _SWC.StackPanel()
        content_sp.Margin = _SW.Thickness(6, 3, 4, 3)
        content_sp.VerticalAlignment = _SW.VerticalAlignment.Center

        # Group label (PROJECT INFO, DISTRIBUTION, etc.) coloured by group
        grp = TYPE_GROUP.get(block['type'], 'layout')
        grp_tb = _SWC.TextBlock()
        grp_tb.Text = GROUP_LABELS.get(grp, grp.upper())
        grp_tb.Foreground = _SWM.SolidColorBrush(GROUP_COLORS.get(grp, GROUP_COLORS['layout']))
        grp_tb.FontSize = 7
        grp_tb.FontWeight = _SW.FontWeights.Bold
        grp_tb.Margin = _SW.Thickness(0, 0, 0, 1)
        content_sp.Children.Add(grp_tb)

        type_tb = _SWC.TextBlock()
        icon = TYPE_ICONS.get(block['type'], '')
        type_tb.Text = '{} {}'.format(icon, TYPE_NAMES.get(block['type'], block['type']))
        type_tb.Foreground = _SWM.SolidColorBrush(SLOT_COLORS[ci])
        type_tb.FontSize = 8; type_tb.FontWeight = _SW.FontWeights.Bold
        content_sp.Children.Add(type_tb)

        is_text  = block['type'] in ('title', 'heading', 'text')
        is_spine = block['type'].startswith('spine_')
        is_blank = block['type'] == 'blank'

        if is_text:
            meta = _SWC.TextBlock()
            meta.Text = block.get('content','') or '(empty)'
            meta.Foreground = BK['muted']; meta.FontSize = 9
            meta.FontStyle = _SW.FontStyles.Italic
            meta.TextTrimming = _SW.TextTrimming.CharacterEllipsis
            content_sp.Children.Add(meta)
        elif is_blank:
            h_sp = _SWC.StackPanel()
            h_sp.Orientation = _SWC.Orientation.Horizontal; h_sp.Margin = _SW.Thickness(0,2,0,0)
            h_lbl = _SWC.TextBlock(); h_lbl.Text = 'Height: '
            h_lbl.Foreground = BK['muted']; h_lbl.FontSize = 9; h_lbl.VerticalAlignment = _SW.VerticalAlignment.Center
            h_tb = _SWC.TextBox()
            h_tb.Text = str(block.get('height_pct') or 100)
            h_tb.Width = 40; h_tb.Height = 18; h_tb.FontSize = 9
            h_tb.Background = BK['row']; h_tb.Foreground = BK['text']
            h_tb.BorderBrush = BK['bdr']; h_tb.BorderThickness = _SW.Thickness(1)
            h_tb.Padding = _SW.Thickness(2, 1, 2, 1)
            h_tb.TextChanged += lambda s,e,r=ri,c=ci: self._set_height(r,c,s.Text)
            h_pct = _SWC.TextBlock(); h_pct.Text = '%'
            h_pct.Foreground = BK['muted']; h_pct.FontSize = 9; h_pct.Margin = _SW.Thickness(2,0,0,0); h_pct.VerticalAlignment = _SW.VerticalAlignment.Center
            h_sp.Children.Add(h_lbl); h_sp.Children.Add(h_tb); h_sp.Children.Add(h_pct)
            content_sp.Children.Add(h_sp)
        elif is_spine:
            rev_tb = _SWC.TextBlock()
            rev_tb.Text = '{} col{}'.format(self._rev_count, 's' if self._rev_count!=1 else '')
            rev_tb.Foreground = BK['muted']; rev_tb.FontSize = 9
            content_sp.Children.Add(rev_tb)
        else:
            lbl_tb = _SWC.TextBox()
            lbl_tb.Text = block.get('label','')
            lbl_tb.Background = _SWM.Brushes.Transparent
            lbl_tb.BorderBrush = BK['bdr']; lbl_tb.BorderThickness = _SW.Thickness(0,0,0,1)
            lbl_tb.Foreground = BK['text']; lbl_tb.FontSize = 10
            lbl_tb.Padding = _SW.Thickness(2,1,2,1)
            lbl_tb.ToolTip = 'Optional label'
            lbl_tb.TextChanged += lambda s,e,r=ri,c=ci: self._update_block_label(r,c,s.Text)
            content_sp.Children.Add(lbl_tb)

        # Span stepper
        span = block.get('span', 1)
        span_sp = _SWC.StackPanel()
        span_sp.Orientation = _SWC.Orientation.Horizontal
        span_sp.Margin = _SW.Thickness(0, 3, 0, 0)

        # Slot colour for this cell (used as hover border colour for span arrows)
        slot_color_brush = SLOT_BRUSHES[ci]

        def make_sb(txt, enabled, fn):
            # Border-based button (Button template can override foreground colour)
            b = _SWC.Border()
            b.Background = BK['row']
            b.BorderBrush = BK['bdr']
            b.BorderThickness = _SW.Thickness(1)
            b.CornerRadius = _SW.CornerRadius(2)
            b.Width = 15; b.Height = 13
            b.Cursor = _SWI.Cursors.Hand if enabled else _SWI.Cursors.Arrow

            t_tb = _SWC.TextBlock()
            t_tb.Text = txt
            t_tb.Foreground = BK['text']  # always white
            t_tb.FontSize = 10
            t_tb.HorizontalAlignment = _SW.HorizontalAlignment.Center
            t_tb.VerticalAlignment = _SW.VerticalAlignment.Center
            if not enabled:
                t_tb.Opacity = 0.35
            b.Child = t_tb

            # Hover: switch border to slot colour, keep text white
            def on_enter(s, e, sc=slot_color_brush):
                if enabled:
                    s.BorderBrush = sc
            def on_leave(s, e):
                s.BorderBrush = BK['bdr']
            b.MouseEnter += on_enter
            b.MouseLeave += on_leave
            if enabled:
                b.MouseLeftButtonUp += fn
            return b

        dec_btn = make_sb('←', span > 1, lambda s,e,r=ri,c=ci: self._decrease_span(r,c))
        span_sp.Children.Add(dec_btn)

        for s in range(span):
            dot = _SWC.Border()
            dot.Width = 6; dot.Height = 6; dot.CornerRadius = _SW.CornerRadius(1)
            dot.Background = SLOT_BRUSHES[min(ci+s, 3)]
            dot.Margin = _SW.Thickness(2, 0, 1, 0)
            dot.VerticalAlignment = _SW.VerticalAlignment.Center
            span_sp.Children.Add(dot)

        span_lbl = _SWC.TextBlock()
        span_lbl.Text = '{}col{}'.format(span, 's' if span>1 else '')
        span_lbl.FontSize = 9; span_lbl.Foreground = BK['muted']
        span_lbl.VerticalAlignment = _SW.VerticalAlignment.Center; span_lbl.Margin = _SW.Thickness(2,0,2,0)
        span_sp.Children.Add(span_lbl)

        inc_btn = make_sb('→', span < mx, lambda s,e,r=ri,c=ci: self._increase_span(r,c))
        span_sp.Children.Add(inc_btn)
        content_sp.Children.Add(span_sp)

        # Row-span arrows (only shown when row is in a merged group)
        if in_group:
            rs = block.get('row_span', 1)
            max_rs = self._max_row_span(ri, ci)
            rspan_sp = _SWC.StackPanel()
            rspan_sp.Orientation = _SWC.Orientation.Horizontal
            rspan_sp.Margin = _SW.Thickness(0, 2, 0, 0)

            def make_rsb(txt, enabled, fn):
                b = _SWC.Border()
                b.Background = BK['row']
                b.BorderBrush = BK['bdr']
                b.BorderThickness = _SW.Thickness(1)
                b.CornerRadius = _SW.CornerRadius(2)
                b.Width = 15; b.Height = 13
                b.Cursor = _SWI.Cursors.Hand if enabled else _SWI.Cursors.Arrow
                t_tb = _SWC.TextBlock()
                t_tb.Text = txt
                t_tb.Foreground = BK['text']
                t_tb.FontSize = 10
                t_tb.HorizontalAlignment = _SW.HorizontalAlignment.Center
                t_tb.VerticalAlignment = _SW.VerticalAlignment.Center
                if not enabled: t_tb.Opacity = 0.35
                b.Child = t_tb
                def on_enter(s, e, sc=BK['accent']):
                    if enabled: s.BorderBrush = sc
                def on_leave(s, e):
                    s.BorderBrush = BK['bdr']
                b.MouseEnter += on_enter
                b.MouseLeave += on_leave
                if enabled: b.MouseLeftButtonUp += fn
                return b

            rs_dec = make_rsb('↑', rs > 1, lambda s,e,r=ri,c=ci: self._decrease_row_span(r,c))
            rspan_sp.Children.Add(rs_dec)

            rs_lbl = _SWC.TextBlock()
            rs_lbl.Text = '{}row{}'.format(rs, 's' if rs > 1 else '')
            rs_lbl.Foreground = BK['accent']; rs_lbl.FontSize = 8
            rs_lbl.VerticalAlignment = _SW.VerticalAlignment.Center
            rs_lbl.Margin = _SW.Thickness(3, 0, 3, 0)
            rspan_sp.Children.Add(rs_lbl)

            rs_inc = make_rsb('↓', rs < max_rs, lambda s,e,r=ri,c=ci: self._increase_row_span(r,c))
            rspan_sp.Children.Add(rs_inc)
            content_sp.Children.Add(rspan_sp)

        _SWC.Grid.SetColumn(content_sp, 1); ig.Children.Add(content_sp)

        # Actions
        act_sp = _SWC.StackPanel()
        act_sp.VerticalAlignment = _SW.VerticalAlignment.Center
        act_sp.Margin = _SW.Thickness(2, 0, 3, 0)

        def act_btn(content, fn, tip=''):
            # Border-based button: white icon, slot-accent hover border
            b = _SWC.Border()
            b.Background = _SWM.Brushes.Transparent
            b.BorderBrush = _SWM.Brushes.Transparent
            b.BorderThickness = _SW.Thickness(1)
            b.CornerRadius = _SW.CornerRadius(2)
            b.Padding = _SW.Thickness(2,1,2,1)
            b.Cursor = _SWI.Cursors.Hand
            b.ToolTip = tip

            tb = _SWC.TextBlock()
            tb.Text = content
            tb.Foreground = BK['white']
            tb.FontSize = 10
            tb.HorizontalAlignment = _SW.HorizontalAlignment.Center
            tb.VerticalAlignment = _SW.VerticalAlignment.Center
            b.Child = tb

            def on_enter(s, e, sc=slot_color_brush):
                s.BorderBrush = sc
            def on_leave(s, e):
                s.BorderBrush = _SWM.Brushes.Transparent
            b.MouseEnter += on_enter
            b.MouseLeave += on_leave
            b.MouseLeftButtonUp += fn
            return b

        tog = act_btn('●' if block.get('enabled',True) else '○',
                      lambda s,e,r=ri,c=ci: self._toggle_block(r,c), 'Show/hide')
        # Tog (enable/disable) uses accent green when enabled to indicate state
        try:
            tog.Child.Foreground = BK['accent'] if block.get('enabled',True) else BK['white']
        except Exception: pass
        act_sp.Children.Add(tog)
        act_sp.Children.Add(act_btn('✕', lambda s,e,r=ri,c=ci: self._remove_block(r,c), 'Remove'))

        _SWC.Grid.SetColumn(act_sp, 2); ig.Children.Add(act_sp)
        cell.Child = ig
        cell.Cursor = _SWI.Cursors.Hand
        cell.MouseLeftButtonUp += lambda s,e,r=ri,c=ci,b=block: self._show_inspector(r,c,b)
        cell_sp.Children.Add(cell)

        return wrapper

    def _open_inspector(self, ri, ci, cell_sp):
        block = self._rows[ri]['blocks'][ci]
        self._show_inspector(ri, ci, block)
        # Also toggle inline panel
        self._toggle_block_settings(ri, ci, cell_sp)

    def _toggle_block_settings(self, ri, ci, cell_sp):
        for child in cell_sp.Children:
            if hasattr(child, 'Tag') and str(child.Tag or '') == 'settings_panel':
                V = _SW.Visibility
                child.Visibility = V.Collapsed if child.Visibility == V.Visible else V.Visible
                break

    def _make_settings_panel(self, ri, ci, block):
        panel = _SWC.Border()
        panel.Background = BK['deep']
        panel.BorderBrush = BK['bdr']; panel.BorderThickness = _SW.Thickness(0,1,0,0)
        panel.Padding = _SW.Thickness(8, 6, 8, 8)
        panel.Tag = 'settings_panel_border'

        sp = _SWC.StackPanel()
        panel.Child = sp

        def sec(lbl):
            tb = _SWC.TextBlock()
            tb.Text = lbl; tb.Foreground = BK['muted']; tb.FontSize = 8
            tb.FontWeight = _SW.FontWeights.Bold; tb.Margin = _SW.Thickness(0,4,0,2)
            sp.Children.Add(tb)

        def row_sp():
            s = _SWC.StackPanel(); s.Orientation = _SWC.Orientation.Horizontal
            s.Margin = _SW.Thickness(0,0,0,4); sp.Children.Add(s); return s

        # ── Horizontal Justification ──────────────────────────────
        sec('JUSTIFY (Horizontal)')
        just_row = row_sp()

        def just_svg(jtype, active):
            c = _SWC.Canvas(); c.Width = 18; c.Height = 18
            configs = {
                'left':   [(2,4,14,4),(2,7,10,7),(2,10,14,10),(2,13,8,13)],
                'center': [(2,4,14,4),(4,7,12,7),(2,10,14,10),(5,13,11,13)],
                'right':  [(2,4,14,4),(6,7,14,7),(2,10,14,10),(8,13,14,13)],
            }
            for x1,y1,x2,y2 in configs[jtype]:
                line = _SWS.Line()
                line.X1=x1;line.Y1=y1;line.X2=x2;line.Y2=y2
                line.StrokeThickness = 1.5
                line.StrokeStartLineCap = _SWM.PenLineCap.Round
                line.StrokeEndLineCap   = _SWM.PenLineCap.Round
                line.Stroke = BK['white']
                c.Children.Add(line)
            return c

        just_btns = {}
        for jval in ('left', 'center', 'right'):
            btn = _SWC.Button()
            btn.Width = 30; btn.Height = 24
            active = block.get('just','left') == jval
            btn.Content = just_svg(jval, active)
            btn.Background = BK['accent'] if active else BK['row']
            btn.BorderBrush = BK['bdr']; btn.BorderThickness = _SW.Thickness(1)
            btn.Cursor = _SWI.Cursors.Hand; btn.Margin = _SW.Thickness(0,0,3,0)
            btn.Tag = jval
            just_btns[jval] = btn
            btn.Click += lambda s,e,r=ri,c=ci,btns=just_btns: self._set_just_h(r,c,str(s.Tag), btns, just_svg)
            just_row.Children.Add(btn)

        # ── Vertical Justification ────────────────────────────────
        sec('JUSTIFY (Vertical)')
        vj_row = row_sp()

        def vj_svg(vtype, active):
            """Draw vertical justify icon: frame + 3 horizontal bars stacked top/middle/bottom."""
            c = _SWC.Canvas(); c.Width = 18; c.Height = 18
            stroke = BK['white']
            # Outer frame
            frame = _SWS.Rectangle()
            frame.Width = 14; frame.Height = 14
            frame.Stroke = stroke; frame.StrokeThickness = 1.2
            frame.Fill = BK['transp']
            _SWC.Canvas.SetLeft(frame, 2); _SWC.Canvas.SetTop(frame, 2)
            c.Children.Add(frame)
            # Three bars positioned by vtype
            y_map = {'top': [4, 6.5, 9], 'middle': [6, 8.5, 11], 'bottom': [8, 10.5, 13]}
            widths = [8, 6, 8]  # slight variance like the icon
            ys = y_map[vtype]
            for i, y in enumerate(ys):
                bar = _SWS.Line()
                bar.X1 = 5; bar.Y1 = y; bar.X2 = 5 + widths[i]; bar.Y2 = y
                bar.Stroke = stroke; bar.StrokeThickness = 1.2
                bar.StrokeStartLineCap = _SWM.PenLineCap.Round
                bar.StrokeEndLineCap   = _SWM.PenLineCap.Round
                c.Children.Add(bar)
            return c

        vj_btns = {}
        for vval in ('top', 'middle', 'bottom'):
            btn = _SWC.Button()
            btn.Width = 30; btn.Height = 24
            active = block.get('v_just','middle') == vval
            btn.Content = vj_svg(vval, active)
            btn.Background = BK['accent'] if active else BK['row']
            btn.BorderBrush = BK['bdr']; btn.BorderThickness = _SW.Thickness(1)
            btn.Cursor = _SWI.Cursors.Hand; btn.Margin = _SW.Thickness(0,0,3,0)
            btn.Tag = vval
            vj_btns[vval] = btn
            btn.Click += lambda s,e,r=ri,c=ci,btns=vj_btns: self._set_just_v(r,c,str(s.Tag), btns, vj_svg)
            vj_row.Children.Add(btn)

        # ── Background colour ─────────────────────────────────────
        sec('BACKGROUND')
        bg_row = row_sp()

        swatch = _SWC.Border()
        swatch.Width = 28; swatch.Height = 24
        swatch.CornerRadius = _SW.CornerRadius(3)
        swatch.BorderBrush = BK['bdr']; swatch.BorderThickness = _SW.Thickness(1)
        swatch.Cursor = _SWI.Cursors.Hand; swatch.Margin = _SW.Thickness(0,0,4,0)
        swatch.Tag = 'bg_swatch'
        self._paint_bg_swatch(swatch, block.get('bg_color'))
        swatch.MouseLeftButtonUp += lambda s,e,r=ri,c=ci,sw=swatch,lb=None: self._pick_bg_color(r,c,sw)
        bg_row.Children.Add(swatch)

        clear_btn = _SWC.Button()
        clear_btn.Content = '✕'; clear_btn.Width = 24; clear_btn.Height = 24
        clear_btn.FontSize = 10; clear_btn.Cursor = _SWI.Cursors.Hand
        clear_btn.Background = BK['row']; clear_btn.Foreground = BK['muted']
        clear_btn.BorderBrush = BK['bdr']; clear_btn.BorderThickness = _SW.Thickness(1)
        clear_btn.Click += lambda s,e,r=ri,c=ci,sw=swatch: self._clear_bg(r,c,sw)
        bg_row.Children.Add(clear_btn)

        bg_lbl = _SWC.TextBlock()
        bg_lbl.Text = block.get('bg_color') or 'none'
        bg_lbl.Foreground = BK['muted']; bg_lbl.FontSize = 9
        bg_lbl.VerticalAlignment = _SW.VerticalAlignment.Center
        bg_lbl.Margin = _SW.Thickness(6,0,0,0); bg_lbl.Tag = 'bg_lbl'
        swatch.Tag = bg_lbl  # so pick_bg_color can update label
        bg_row.Children.Add(bg_lbl)

        # ── Text style ────────────────────────────────────────────
        sec('TEXT STYLE')
        ts_cb = _SWC.ComboBox()
        try:
            ts_cb.Style = self.FindResource('MCB')
        except Exception:
            # Manual fallback styling
            ts_cb.Background = BK['row']
            ts_cb.Foreground = BK['text']
            ts_cb.BorderBrush = BK['bdr']
            ts_cb.BorderThickness = _SW.Thickness(1)
        ts_cb.Height = 24; ts_cb.FontSize = 10
        ts_cb.Margin = _SW.Thickness(0, 0, 0, 4)
        # Build items with proper item container style for dark popup
        item_style = _SW.Style()
        item_style.TargetType = _SWC.ComboBoxItem
        item_style.Setters.Add(_SW.Setter(_SWC.Control.BackgroundProperty, BK['card']))
        item_style.Setters.Add(_SW.Setter(_SWC.Control.ForegroundProperty, BK['text']))
        item_style.Setters.Add(_SW.Setter(_SWC.Control.PaddingProperty, _SW.Thickness(8, 4, 8, 4)))
        item_style.Setters.Add(_SW.Setter(_SWC.Control.FontSizeProperty, 11.0))
        ts_cb.ItemContainerStyle = item_style

        current_style = block.get('text_style', 'Data')
        style_names = self._sorted_style_names()
        for name in style_names:
            ts_cb.Items.Add(name)
        try:
            if current_style in style_names:
                ts_cb.SelectedIndex = style_names.index(current_style)
        except Exception:
            ts_cb.SelectedIndex = 0
        ts_cb.SelectionChanged += lambda s, e, r=ri, c=ci: self._set_text_style(
            r, c, str(s.SelectedItem) if s.SelectedItem else 'Data')
        sp.Children.Add(ts_cb)



        # ── Spine block settings (vertical lines + rotation) ─────
        # ── Data grid lines — H/V icons (data blocks + spine blocks) ──
        _has_h = block.get('type') in DATA_BLOCK_TYPES
        _has_v = block.get('type') in DATA_BLOCK_TYPES or block.get('type') in SPINE_HEADER_TYPES
        if _has_h or _has_v:
            sec('DATA GRID LINES')
            db_row = row_sp()
            db = block.get('data_borders', {'h':True,'v':True})
            # H icon: horizontal lines ≡
            if _has_h:
                h_btn = _SWC.Button()
                h_btn.Padding = _SW.Thickness(4,2,4,2)
                h_btn.FontSize = 13; h_btn.Height = 26; h_btn.Width = 34
                h_btn.Cursor = _SWI.Cursors.Hand; h_btn.Margin = _SW.Thickness(0,0,4,0)
                h_btn.Background = BK['accent'] if db.get('h') else BK['row']
                h_btn.Foreground = BK['text']
                h_btn.BorderBrush = BK['bdr']; h_btn.BorderThickness = _SW.Thickness(1)
                h_btn.ToolTip = 'Horizontal grid lines'
                h_btn.Tag = 'h'
                # ≡ icon drawn as three stacked lines via TextBlock
                h_tb = _SWC.TextBlock()
                h_tb.Text = u'☰'; h_tb.FontSize = 13
                h_tb.HorizontalAlignment = _SW.HorizontalAlignment.Center
                h_tb.VerticalAlignment = _SW.VerticalAlignment.Center
                h_btn.Content = h_tb
                h_btn.Click += lambda s,e,r=ri,c=ci: self._toggle_data_border(r,c,'h',s)
                db_row.Children.Add(h_btn)
            # V icon: vertical lines |||
            if _has_v:
                v_btn2 = _SWC.Button()
                v_btn2.Padding = _SW.Thickness(4,2,4,2)
                v_btn2.FontSize = 13; v_btn2.Height = 26; v_btn2.Width = 34
                v_btn2.Cursor = _SWI.Cursors.Hand; v_btn2.Margin = _SW.Thickness(0,0,0,0)
                v_btn2.Background = BK['accent'] if db.get('v') else BK['row']
                v_btn2.Foreground = BK['text']
                v_btn2.BorderBrush = BK['bdr']; v_btn2.BorderThickness = _SW.Thickness(1)
                v_btn2.ToolTip = 'Vertical grid lines'
                v_btn2.Tag = 'v'
                v_tb = _SWC.TextBlock()
                v_tb.Text = u'‖'; v_tb.FontSize = 13
                v_tb.HorizontalAlignment = _SW.HorizontalAlignment.Center
                v_tb.VerticalAlignment = _SW.VerticalAlignment.Center
                v_btn2.Content = v_tb
                v_btn2.Click += lambda s,e,r=ri,c=ci: self._toggle_data_border(r,c,'v',s)
                db_row.Children.Add(v_btn2)

        # ── Text orientation (spine blocks only) ──────────────────
        if block.get('type') in SPINE_HEADER_TYPES:
            sec('TEXT ORIENTATION')
            orient_row = row_sp()
            is_rotated = block.get('rotation', 270) == 270
            for lbl_o, rotated_val in [('Vertical', True), ('Horizontal', False)]:
                o_btn = _SWC.Button()
                o_btn.Content = lbl_o; o_btn.Padding = _SW.Thickness(6,2,6,2)
                o_btn.FontSize = 9; o_btn.Height = 22; o_btn.Cursor = _SWI.Cursors.Hand
                o_btn.Margin = _SW.Thickness(0,0,4,0)
                o_btn.Background = BK['accent'] if is_rotated == rotated_val else BK['row']
                o_btn.Foreground = BK['text']
                o_btn.BorderBrush = BK['bdr']; o_btn.BorderThickness = _SW.Thickness(1)
                o_btn.Click += lambda s,e,r=ri,c=ci,rv=rotated_val: self._set_rotation(r,c,rv,s)
                orient_row.Children.Add(o_btn)

        # ── Alternate rows (data blocks) ──────────────────────────
        if block.get('type') in DATA_BLOCK_TYPES:
            sec('ALTERNATE ROWS')
            alt_row_sp = row_sp()

            alt_toggle = _SWC.Button()
            alt_toggle.Content = 'On' if block.get('alt_rows') else 'Off'
            alt_toggle.Width = 36; alt_toggle.Height = 22; alt_toggle.FontSize = 9
            alt_toggle.Background = BK['accent'] if block.get('alt_rows') else BK['row']
            alt_toggle.Foreground = BK['text']
            alt_toggle.BorderBrush = BK['bdr']; alt_toggle.BorderThickness = _SW.Thickness(1)
            alt_toggle.Cursor = _SWI.Cursors.Hand; alt_toggle.Margin = _SW.Thickness(0,0,6,0)
            alt_toggle.Click += lambda s,e,r=ri,c=ci,btn=alt_toggle: self._toggle_alt_rows(r,c,btn)
            alt_row_sp.Children.Add(alt_toggle)

            alt_color_swatch = _SWC.Border()
            alt_color_swatch.Width = 22; alt_color_swatch.Height = 22
            alt_color_swatch.CornerRadius = _SW.CornerRadius(3)
            alt_color_swatch.BorderBrush = BK['bdr']; alt_color_swatch.BorderThickness = _SW.Thickness(1)
            alt_color_swatch.Cursor = _SWI.Cursors.Hand
            try: alt_color_swatch.Background = _hbrush(block.get('alt_color','#F5F7FA'))
            except Exception: alt_color_swatch.Background = BK['row']
            alt_color_swatch.MouseLeftButtonUp += lambda s,e,r=ri,c=ci,sw=alt_color_swatch: self._pick_alt_color(r,c,sw)
            alt_row_sp.Children.Add(alt_color_swatch)

            alt_lbl = _SWC.TextBlock()
            alt_lbl.Text = block.get('alt_color','#F5F7FA')
            alt_lbl.Foreground = BK['muted']; alt_lbl.FontSize = 9
            alt_lbl.VerticalAlignment = _SW.VerticalAlignment.Center
            alt_lbl.Margin = _SW.Thickness(5,0,0,0)
            alt_row_sp.Children.Add(alt_lbl)

        # ── List/Row for reason/method ────────────────────────────
        if block.get('type') in ('reason_list', 'method_list'):
            sec('DISPLAY')
            ls_row = row_sp()
            for val, lbl in [('list','List'),('row','Row')]:
                btn = _SWC.Button()
                btn.Content = lbl; btn.Padding = _SW.Thickness(8,2,8,2)
                btn.FontSize = 9; btn.Height = 22; btn.Cursor = _SWI.Cursors.Hand
                btn.Margin = _SW.Thickness(0,0,4,0)
                btn.Background = BK['accent'] if block.get('list_style','list')==val else BK['row']
                btn.Foreground = BK['text']
                btn.BorderBrush = BK['bdr']; btn.BorderThickness = _SW.Thickness(1)
                btn.Tag = val
                btn.Click += lambda s,e,r=ri,c=ci: self._set_list_style(r,c,str(s.Tag),ls_row)
                ls_row.Children.Add(btn)

        # ── Page Count settings ───────────────────────────────────
        if block.get('type') == 'page_count':
            sec('FORMAT')
            fmt_cb = _SWC.ComboBox()
            try: fmt_cb.Style = self.FindResource('MCB')
            except Exception:
                fmt_cb.Background = BK['row']; fmt_cb.Foreground = BK['text']
                fmt_cb.BorderBrush = BK['bdr']; fmt_cb.BorderThickness = _SW.Thickness(1)
            fmt_cb.Height = 24; fmt_cb.FontSize = 10
            fmt_cb.Margin = _SW.Thickness(0, 0, 0, 4)
            item_style = _SW.Style()
            item_style.TargetType = _SWC.ComboBoxItem
            item_style.Setters.Add(_SW.Setter(_SWC.Control.BackgroundProperty, BK['card']))
            item_style.Setters.Add(_SW.Setter(_SWC.Control.ForegroundProperty, BK['text']))
            item_style.Setters.Add(_SW.Setter(_SWC.Control.PaddingProperty, _SW.Thickness(8, 4, 8, 4)))
            fmt_cb.ItemContainerStyle = item_style
            cur_fmt = block.get('page_format', 'Page X of Y')
            for f in PAGE_COUNT_FORMATS:
                fmt_cb.Items.Add(f)
            try:
                if cur_fmt in PAGE_COUNT_FORMATS:
                    fmt_cb.SelectedIndex = PAGE_COUNT_FORMATS.index(cur_fmt)
            except Exception: fmt_cb.SelectedIndex = 1
            fmt_cb.SelectionChanged += lambda s,e,r=ri,c=ci: self._set_block_field(r,c,'page_format',str(s.SelectedItem) if s.SelectedItem else 'Page X of Y')
            sp.Children.Add(fmt_cb)

            sec('PREFIX')
            pfx_tb = _SWC.TextBox()
            pfx_tb.Text = block.get('prefix', ''); pfx_tb.FontSize = 10; pfx_tb.Height = 24
            pfx_tb.Background = BK['row']; pfx_tb.Foreground = BK['text']
            pfx_tb.BorderBrush = BK['bdr']; pfx_tb.BorderThickness = _SW.Thickness(1)
            pfx_tb.Margin = _SW.Thickness(0, 0, 0, 4)
            pfx_tb.LostFocus += lambda s,e,r=ri,c=ci: self._set_block_field(r,c,'prefix',s.Text)
            sp.Children.Add(pfx_tb)

            sec('SUFFIX')
            sfx_tb = _SWC.TextBox()
            sfx_tb.Text = block.get('suffix', ''); sfx_tb.FontSize = 10; sfx_tb.Height = 24
            sfx_tb.Background = BK['row']; sfx_tb.Foreground = BK['text']
            sfx_tb.BorderBrush = BK['bdr']; sfx_tb.BorderThickness = _SW.Thickness(1)
            sfx_tb.Margin = _SW.Thickness(0, 0, 0, 4)
            sfx_tb.LostFocus += lambda s,e,r=ri,c=ci: self._set_block_field(r,c,'suffix',s.Text)
            sp.Children.Add(sfx_tb)

        # ── Current Issue Date settings ───────────────────────────
        if block.get('type') == 'issue_date':
            sec('DATE FORMAT')
            dfmt_cb = _SWC.ComboBox()
            try: dfmt_cb.Style = self.FindResource('MCB')
            except Exception:
                dfmt_cb.Background = BK['row']; dfmt_cb.Foreground = BK['text']
                dfmt_cb.BorderBrush = BK['bdr']; dfmt_cb.BorderThickness = _SW.Thickness(1)
            dfmt_cb.Height = 24; dfmt_cb.FontSize = 10
            dfmt_cb.Margin = _SW.Thickness(0, 0, 0, 4)
            item_style2 = _SW.Style()
            item_style2.TargetType = _SWC.ComboBoxItem
            item_style2.Setters.Add(_SW.Setter(_SWC.Control.BackgroundProperty, BK['card']))
            item_style2.Setters.Add(_SW.Setter(_SWC.Control.ForegroundProperty, BK['text']))
            item_style2.Setters.Add(_SW.Setter(_SWC.Control.PaddingProperty, _SW.Thickness(8, 4, 8, 4)))
            dfmt_cb.ItemContainerStyle = item_style2
            cur_dfmt = block.get('date_format', 'dd/MM/yyyy')
            for f in DATE_FORMATS:
                dfmt_cb.Items.Add(f)
            try:
                if cur_dfmt in DATE_FORMATS:
                    dfmt_cb.SelectedIndex = DATE_FORMATS.index(cur_dfmt)
            except Exception: dfmt_cb.SelectedIndex = 0
            dfmt_cb.SelectionChanged += lambda s,e,r=ri,c=ci: self._set_block_field(r,c,'date_format',str(s.SelectedItem) if s.SelectedItem else 'dd/MM/yyyy')
            sp.Children.Add(dfmt_cb)

            sec('PREFIX')
            pfx_tb2 = _SWC.TextBox()
            pfx_tb2.Text = block.get('prefix', ''); pfx_tb2.FontSize = 10; pfx_tb2.Height = 24
            pfx_tb2.Background = BK['row']; pfx_tb2.Foreground = BK['text']
            pfx_tb2.BorderBrush = BK['bdr']; pfx_tb2.BorderThickness = _SW.Thickness(1)
            pfx_tb2.Margin = _SW.Thickness(0, 0, 0, 4)
            pfx_tb2.LostFocus += lambda s,e,r=ri,c=ci: self._set_block_field(r,c,'prefix',s.Text)
            sp.Children.Add(pfx_tb2)

            sec('SUFFIX')
            sfx_tb2 = _SWC.TextBox()
            sfx_tb2.Text = block.get('suffix', ''); sfx_tb2.FontSize = 10; sfx_tb2.Height = 24
            sfx_tb2.Background = BK['row']; sfx_tb2.Foreground = BK['text']
            sfx_tb2.BorderBrush = BK['bdr']; sfx_tb2.BorderThickness = _SW.Thickness(1)
            sfx_tb2.Margin = _SW.Thickness(0, 0, 0, 4)
            sfx_tb2.LostFocus += lambda s,e,r=ri,c=ci: self._set_block_field(r,c,'suffix',s.Text)
            sp.Children.Add(sfx_tb2)

        return panel

    def _paint_bg_swatch(self, swatch, hex_color):
        """Paint a swatch showing either a colour or a '+' placeholder."""
        if hex_color:
            try: swatch.Background = _hbrush(hex_color)
            except Exception: swatch.Background = BK['row']
            swatch.Child = None
        else:
            swatch.Background = BK['row']
            plus = _SWC.TextBlock()
            plus.Text = '+'; plus.FontSize = 12
            plus.Foreground = BK['muted']
            plus.HorizontalAlignment = _SW.HorizontalAlignment.Center
            plus.VerticalAlignment   = _SW.VerticalAlignment.Center
            swatch.Child = plus

    # ── Block setting setters ─────────────────────────────────────
    def _set_just_h(self, ri, ci, val, btns, svg_fn):
        b = self._rows[ri]['blocks'][ci]
        if not b: return
        b['just'] = val
        # Update button highlights AND icon strokes in-place
        for jval, btn in btns.items():
            active = (jval == val)
            btn.Background = BK['accent'] if active else BK['row']
            btn.Content = svg_fn(jval, active)
        self._render_preview()

    def _set_just_v(self, ri, ci, val, btns, svg_fn):
        b = self._rows[ri]['blocks'][ci]
        if not b: return
        b['v_just'] = val
        for vval, btn in btns.items():
            active = (vval == val)
            btn.Background = BK['accent'] if active else BK['row']
            btn.Content = svg_fn(vval, active)
        self._render_preview()

    def _pick_bg_color(self, ri, ci, swatch):
        b = self._rows[ri]['blocks'][ci]
        if not b: return
        color = self._color_dialog(b.get('bg_color'))
        if color is not None:
            b['bg_color'] = color
            self._paint_bg_swatch(swatch, color)
            # Update hex label if it exists (stored in swatch.Tag)
            try:
                lbl = swatch.Tag
                if lbl: lbl.Text = color
            except Exception: pass
            self._render_preview()

    def _clear_bg(self, ri, ci, swatch):
        b = self._rows[ri]['blocks'][ci]
        if not b: return
        b['bg_color'] = None
        self._paint_bg_swatch(swatch, None)
        try:
            lbl = swatch.Tag
            if lbl: lbl.Text = 'none'
        except Exception: pass
        self._render_preview()

    def _toggle_border(self, ri, ci, side, btn):
        b = self._rows[ri]['blocks'][ci]
        if not b: return
        b.setdefault('borders', {'t':True,'b':True,'l':False,'r':False})
        b['borders'][side] = not b['borders'].get(side, True)
        btn.Background = BK['accent'] if b['borders'][side] else BK['row']
        self._render_preview()

    def _toggle_data_border(self, ri, ci, axis, btn):
        b = self._rows[ri]['blocks'][ci]
        if not b: return
        b.setdefault('data_borders', {'h':True,'v':True})
        b['data_borders'][axis] = not b['data_borders'].get(axis, True)
        btn.Background = BK['accent'] if b['data_borders'][axis] else BK['row']
        self._render_preview()

    def _set_rotation(self, ri, ci, rotated, btn_clicked):
        b = self._rows[ri]['blocks'][ci]
        if not b: return
        b['rotation'] = 270 if rotated else 0
        self._render_canvas(); self._render_preview()

    def _toggle_rotation(self, ri, ci, btn):
        b = self._rows[ri]['blocks'][ci]
        if not b: return
        cur = b.get('rotation', 270)
        b['rotation'] = 0 if cur == 270 else 270
        btn.Content = 'Rotated' if b['rotation'] == 270 else 'Normal'
        btn.Background = BK['accent'] if b['rotation'] == 270 else BK['row']
        self._render_preview()

    def _set_text_style(self, ri, ci, val):
        b = self._rows[ri]['blocks'][ci]
        if b: b['text_style'] = val; self._render_preview()

    def _toggle_alt_rows(self, ri, ci, btn):
        b = self._rows[ri]['blocks'][ci]
        if not b: return
        b['alt_rows'] = not b.get('alt_rows', False)
        btn.Content = 'On' if b['alt_rows'] else 'Off'
        btn.Background = BK['accent'] if b['alt_rows'] else BK['row']
        self._render_preview()

    def _pick_alt_color(self, ri, ci, swatch):
        b = self._rows[ri]['blocks'][ci]
        if not b: return
        color = self._color_dialog(b.get('alt_color','#F5F7FA'))
        if color is not None:
            b['alt_color'] = color
            try: swatch.Background = _hbrush(color)
            except Exception: pass
            self._render_preview()

    def _set_list_style(self, ri, ci, val, ls_row):
        b = self._rows[ri]['blocks'][ci]
        if not b: return
        b['list_style'] = val
        for child in ls_row.Children:
            if hasattr(child,'Tag'):
                child.Background = BK['accent'] if str(child.Tag)==val else BK['row']
        self._render_preview()

    def _set_height(self, ri, ci, val):
        b = self._rows[ri]['blocks'][ci]
        if not b: return
        try: b['height_pct'] = max(10, min(200, int(val)))
        except Exception: pass
        self._render_preview()

    def _set_block_field(self, ri, ci, field, val):
        """Generic setter for block fields (prefix, suffix, page_format, date_format, etc)."""
        b = self._rows[ri]['blocks'][ci]
        if not b: return
        b[field] = val
        self._render_preview()

    # ── Drag events ───────────────────────────────────────────────
    def _cell_drag_over(self, sender, args, ri, ci):
        try:
            # Always show Move cursor — we verify data in Drop handler
            args.Effects = _SW.DragDropEffects.Move
            args.Handled = True
            sender.BorderBrush = BK['accent']
            sender.BorderThickness = _SW.Thickness(2)
        except Exception:
            pass

    def _cell_drag_leave(self, sender, args, ri, ci):
        try:
            sender.BorderBrush = BK['bdr']
            sender.BorderThickness = _SW.Thickness(1)
        except Exception:
            pass

    def _cell_drop(self, sender, args, ri, ci):
        try:
            sender.BorderBrush = BK['bdr']
            sender.BorderThickness = _SW.Thickness(1)
            args.Handled = True
            # Try palette type first
            try:
                raw = args.Data.GetData('pal_type')
                if raw is not None:
                    self._place_block(ri, ci, str(raw))
                    return
            except Exception:
                pass
            # Try block move
            try:
                fr_raw = args.Data.GetData('blk_row')
                fc_raw = args.Data.GetData('blk_col')
                if fr_raw is not None and fc_raw is not None:
                    self._move_block(int(str(fr_raw)), int(str(fc_raw)), ri, ci)
                    return
            except Exception:
                pass
            # Fallback: check _drag_type set during palette drag
            if self._drag_type:
                self._place_block(ri, ci, self._drag_type)
                self._drag_type = None
        except Exception:
            pass

    def _block_drag_start(self, sender, args, ri, ci):
        try:
            import System.Windows.Input as _WI
            if args.LeftButton != _WI.MouseButtonState.Pressed: return
            data = _SW.DataObject()
            data.SetData('blk_row', str(ri)); data.SetData('blk_col', str(ci))
            _SW.DragDrop.DoDragDrop(sender, data, _SW.DragDropEffects.Move)
        except Exception: pass

    # ── Text styles panel ─────────────────────────────────────────
    def add_style_click(self, s, e):
        name = self._prompt('Style name:')
        if not name or name in self._text_styles:
            self._status('Invalid or duplicate name'); return
        self._text_styles[name] = {'font':'Arial','size_mm':2.5,'bold':False,'italic':False,'underline':False,'color':'#000000'}
        self._render_style_cards()

    def _get_system_fonts(self):
        """Return a cached, sorted list of system font family names (as Revit uses)."""
        if getattr(self, '_cached_fonts', None) is None:
            try:
                families = _SWM.Fonts.SystemFontFamilies
                names = set()
                for ff in families:
                    try:
                        names.add(str(ff.Source))
                    except Exception: pass
                self._cached_fonts = sorted(names, key=lambda x: x.lower())
            except Exception:
                # Fallback if font enumeration fails
                self._cached_fonts = ['Arial','Times New Roman','Courier New','Segoe UI','Calibri']
        return self._cached_fonts

    def _sorted_style_names(self):
        """Built-ins first (Title, Header, Data), then custom styles A-Z, 0-9."""
        priority = ['Title', 'Header', 'Data']
        names = list(self._text_styles.keys())
        built_ins = [n for n in priority if n in names]
        custom = sorted([n for n in names if n not in priority], key=lambda x: x.lower())
        return built_ins + custom

    def _render_style_cards(self):
        stack = self.style_cards_stack
        stack.Children.Clear()
        for name in self._sorted_style_names():
            st = self._text_styles.get(name)
            if st is None: continue
            card = self._make_style_card(name, st)
            stack.Children.Add(card)

    def _make_style_card(self, name, st):
        outer = _SWC.Border()
        outer.Background = BK['card']; outer.CornerRadius = _SW.CornerRadius(6)
        outer.BorderBrush = BK['bdr']; outer.BorderThickness = _SW.Thickness(1)
        outer.Padding = _SW.Thickness(10); outer.Margin = _SW.Thickness(0,0,0,8)

        sp = _SWC.StackPanel()
        outer.Child = sp

        # Header
        hdr = _SWC.Grid()
        h_tb = _SWC.TextBlock(); h_tb.Text = name; h_tb.Foreground = BK['accent']
        h_tb.FontSize = 11; h_tb.FontWeight = _SW.FontWeights.SemiBold
        del_btn = _SWC.Button(); del_btn.Content = '✕'; del_btn.FontSize = 10
        del_btn.Background = _SWM.Brushes.Transparent; del_btn.BorderThickness = _SW.Thickness(0)
        del_btn.Foreground = BK['muted']; del_btn.Cursor = _SWI.Cursors.Hand
        del_btn.HorizontalAlignment = _SW.HorizontalAlignment.Right
        built_in = name in ('Title','Header','Data')
        # Built-in styles cannot be deleted — hide the X entirely
        del_btn.Visibility = _SW.Visibility.Collapsed if built_in else _SW.Visibility.Visible
        del_btn.Click += lambda s,e,n=name: self._delete_style(n)
        hdr.Children.Add(h_tb); hdr.Children.Add(del_btn)
        sp.Children.Add(hdr)

        # Fields grid
        fields = _SWC.Grid(); fields.Margin = _SW.Thickness(0,8,0,0)
        fields.ColumnDefinitions.Add(_SWC.ColumnDefinition()); fields.ColumnDefinitions.Add(_SWC.ColumnDefinition())
        fields.RowDefinitions.Add(_SWC.RowDefinition()); fields.RowDefinitions.Add(_SWC.RowDefinition())

        def add_field(grid, col, row, label, control):
            fp = _SWC.StackPanel(); fp.Margin = _SW.Thickness(0,0,6,6)
            lbl = _SWC.TextBlock(); lbl.Text = label; lbl.Foreground = BK['muted']
            lbl.FontSize = 9; lbl.Margin = _SW.Thickness(0,0,0,2)
            fp.Children.Add(lbl); fp.Children.Add(control)
            _SWC.Grid.SetColumn(fp, col); _SWC.Grid.SetRow(fp, row); grid.Children.Add(fp)

        # Font — enumerate system fonts (same set Revit uses for text styles)
        font_cb = _SWC.ComboBox()
        try: font_cb.Style = self.FindResource('MCB')
        except Exception:
            font_cb.Background = BK['row']
            font_cb.Foreground = BK['text']
            font_cb.BorderBrush = BK['bdr']
            font_cb.BorderThickness = _SW.Thickness(1)
        font_cb.FontSize = 10; font_cb.Height = 24
        # ItemContainerStyle for dark popup
        font_item_style = _SW.Style()
        font_item_style.TargetType = _SWC.ComboBoxItem
        font_item_style.Setters.Add(_SW.Setter(_SWC.Control.BackgroundProperty, BK['card']))
        font_item_style.Setters.Add(_SW.Setter(_SWC.Control.ForegroundProperty, BK['text']))
        font_item_style.Setters.Add(_SW.Setter(_SWC.Control.PaddingProperty, _SW.Thickness(8, 4, 8, 4)))
        font_item_style.Setters.Add(_SW.Setter(_SWC.Control.FontSizeProperty, 11.0))
        font_cb.ItemContainerStyle = font_item_style
        for f in self._get_system_fonts():
            font_cb.Items.Add(f)
        try: font_cb.SelectedItem = st.get('font','Arial')
        except Exception: font_cb.SelectedIndex=0
        font_cb.SelectionChanged += lambda s,e,n=name: self._update_style(n,'font',str(s.SelectedItem) if s.SelectedItem else 'Arial')
        add_field(fields, 0, 0, 'Font', font_cb)

        # Size mm
        size_tb = _SWC.TextBox(); size_tb.Text=str(st.get('size_mm',2.5)); size_tb.Height=24; size_tb.FontSize=10
        size_tb.Background=BK['row']; size_tb.Foreground=BK['text']; size_tb.BorderBrush=BK['bdr']; size_tb.BorderThickness=_SW.Thickness(1)
        size_tb.Padding=_SW.Thickness(4,2,4,2); size_tb.VerticalContentAlignment=_SW.VerticalAlignment.Center
        size_tb.TextChanged += lambda s,e,n=name: self._update_style_size(n,s.Text)
        add_field(fields, 1, 0, 'Size (mm)', size_tb)

        # Colour (swatch button)
        color_btn = _SWC.Button(); color_btn.Height=24; color_btn.FontSize=9; color_btn.Cursor=_SWI.Cursors.Hand
        color_btn.BorderBrush=BK['bdr']; color_btn.BorderThickness=_SW.Thickness(1)
        try: color_btn.Background=_hbrush(st.get('color','#000000'))
        except Exception: color_btn.Background=BK['card']
        color_btn.Content=st.get('color','#000000'); color_btn.Foreground=BK['text']
        color_btn.Click += lambda s,e,n=name,btn=color_btn: self._pick_style_color(n,btn)
        add_field(fields, 0, 1, 'Colour', color_btn)

        # Style toggles
        tog_sp = _SWC.StackPanel(); tog_sp.Orientation=_SWC.Orientation.Horizontal
        for prop, lbl in [('bold','B'),('italic','I'),('underline','U')]:
            tb = _SWC.Button()
            # Use TextBlock as content so we can apply TextDecorations (buttons don't support it)
            lbl_tb = _SWC.TextBlock(); lbl_tb.Text = lbl
            lbl_tb.HorizontalAlignment = _SW.HorizontalAlignment.Center
            lbl_tb.VerticalAlignment = _SW.VerticalAlignment.Center
            if prop == 'bold': lbl_tb.FontWeight = _SW.FontWeights.Bold
            if prop == 'italic': lbl_tb.FontStyle = _SW.FontStyles.Italic
            if prop == 'underline': lbl_tb.TextDecorations = _SW.TextDecorations.Underline
            tb.Content = lbl_tb
            tb.Width = 26; tb.Height = 24; tb.FontSize = 10
            tb.Cursor = _SWI.Cursors.Hand; tb.Margin = _SW.Thickness(0,0,3,0)
            tb.Background = BK['accent'] if st.get(prop) else BK['row']
            tb.Foreground = BK['text']; tb.BorderBrush = BK['bdr']
            tb.BorderThickness = _SW.Thickness(1)
            tb.Tag = prop
            tb.Click += lambda s,e,n=name: self._toggle_style_prop(n, str(s.Tag), s)
            tog_sp.Children.Add(tb)
        add_field(fields, 1, 1, 'Style', tog_sp)

        sp.Children.Add(fields)

        # Preview swatch
        prev_b = _SWC.Border(); prev_b.Background=BK['row']; prev_b.CornerRadius=_SW.CornerRadius(3)
        prev_b.Padding=_SW.Thickness(6,3,6,3); prev_b.Margin=_SW.Thickness(0,6,0,0)
        prev_tb = _SWC.TextBlock(); prev_tb.Text=name+' — The quick brown fox'
        try:
            prev_tb.FontFamily=_SWM.FontFamily(st.get('font','Arial'))
            prev_tb.FontSize=max(4, st.get('size_mm',2.5)*3.5)
            prev_tb.FontWeight=_SW.FontWeights.Bold if st.get('bold') else _SW.FontWeights.Normal
            prev_tb.FontStyle=_SW.FontStyles.Italic if st.get('italic') else _SW.FontStyles.Normal
            if st.get('underline'): prev_tb.TextDecorations=_SW.TextDecorations.Underline
            prev_tb.Foreground=_hbrush(st.get('color','#000000'))
        except Exception: prev_tb.Foreground=BK['black']
        prev_b.Child=prev_tb; sp.Children.Add(prev_b)

        return outer

    def _update_style(self, name, prop, val):
        if name not in self._text_styles: return
        self._text_styles[name][prop] = val
        self._render_style_cards(); self._render_preview()

    def _update_style_size(self, name, val):
        try:
            v = max(0.5, min(20.0, float(val)))
            self._text_styles[name]['size_mm'] = v
            self._render_preview()
        except Exception: pass

    def _pick_style_color(self, name, btn):
        color = self._color_dialog(self._text_styles[name].get('color','#000000'))
        if color is not None:
            self._text_styles[name]['color'] = color
            try: btn.Background=_hbrush(color); btn.Content=color
            except Exception: pass
            self._render_style_cards(); self._render_preview()

    def _toggle_style_prop(self, name, prop, btn):
        if name not in self._text_styles: return
        self._text_styles[name][prop] = not self._text_styles[name].get(prop, False)
        btn.Background = BK['accent'] if self._text_styles[name][prop] else BK['row']
        self._render_style_cards(); self._render_preview()

    def _delete_style(self, name):
        if name in ('Title','Header','Data'):
            self._status('Cannot delete built-in style'); return
        del self._text_styles[name]
        self._render_style_cards()

    # ── Preview ───────────────────────────────────────────────────
    def _on_prev_resize(self, sender, args):
        self._render_preview()

    def _render_preview(self):
        try:
            avail = max(100, self.preview_scroll.ActualWidth - 24)
            content, pw = self._preview_builder.build(
                self._rows, self._rev_count, self._col_pct,
                self._page_w_mm, avail, self._text_styles, self._logo_path,
                hlines=getattr(self,'_hlines',{}),
                vlines=getattr(self,'_vlines',{}))
            self.preview_stack.Children.Clear()
            self.preview_stack.Children.Add(content)
            self.preview_paper.Width = pw
        except Exception as ex:
            try:
                self.preview_stack.Children.Clear()
                tb = _SWC.TextBlock()
                tb.Text = 'Preview error: {}'.format(ex)
                tb.Foreground = BK['danger']; tb.FontSize = 9; tb.Margin = _SW.Thickness(8)
                self.preview_stack.Children.Add(tb)
            except Exception: pass

    # ── Colour picker dialog ──────────────────────────────────────
    def _color_dialog(self, current_hex=None):
        """Show a simple hex-input dialog for colour picking."""
        import System.Windows.Markup as _Markup
        xaml_str = u'''<Window
            xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
            Title="" Width="300" SizeToContent="Height"
            WindowStyle="None" ResizeMode="NoResize"
            WindowStartupLocation="CenterScreen"
            Background="Transparent" AllowsTransparency="True" FontFamily="Segoe UI">
          <Border Background="#2B3340" CornerRadius="8" Margin="10" Padding="18,14">
            <Border.Effect><DropShadowEffect Color="Black" Opacity="0.5" ShadowDepth="3" BlurRadius="10"/></Border.Effect>
            <StackPanel>
              <Border Background="#208A3C" Height="2" CornerRadius="1" Margin="0,0,0,10"/>
              <TextBlock Text="Background Colour" Foreground="#F4FAFF" FontSize="12" FontWeight="SemiBold" Margin="0,0,0,8"/>
              <TextBlock Text="Enter hex colour (#RRGGBB)" Foreground="#8A96A8" FontSize="10" Margin="0,0,0,6"/>
              <TextBox x:Name="hex_tb" Background="#323D4D" Foreground="#F4FAFF"
                       BorderBrush="#404E62" BorderThickness="1" Padding="6,4"
                       FontSize="12" Margin="0,0,0,12"/>
              <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="ok_btn" Content="OK" Width="64" Height="26"
                        Background="#208A3C" Foreground="#F4FAFF" BorderThickness="0"
                        FontSize="11" Margin="0,0,6,0"/>
                <Button x:Name="cancel_btn" Content="Cancel" Width="64" Height="26"
                        Background="#404E62" Foreground="#F4FAFF" BorderThickness="0" FontSize="11"/>
              </StackPanel>
            </StackPanel>
          </Border>
        </Window>'''
        try:
            dlg = _Markup.XamlReader.Parse(xaml_str)
            dlg.FindName('hex_tb').Text = current_hex or '#FFFFFF'
            result = [None]
            def ok(s, e):
                val = dlg.FindName('hex_tb').Text.strip()
                if not val.startswith('#'): val = '#' + val
                result[0] = val
                dlg.Close()
            def cancel(s, e): dlg.Close()
            dlg.FindName('ok_btn').Click    += ok
            dlg.FindName('cancel_btn').Click += cancel
            dlg.ShowDialog()
            return result[0]
        except Exception:
            return None

    # ── Helpers ───────────────────────────────────────────────────
    def _prompt(self, message, default=''):
        import System.Windows.Markup as _Markup
        xaml_str = u'''<Window
            xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
            Title="" Width="360" SizeToContent="Height"
            WindowStyle="None" ResizeMode="NoResize"
            WindowStartupLocation="CenterScreen"
            Background="Transparent" AllowsTransparency="True" FontFamily="Segoe UI">
          <Border Background="#2B3340" CornerRadius="8" Margin="10" Padding="18,14">
            <Border.Effect><DropShadowEffect Color="Black" Opacity="0.5" ShadowDepth="3" BlurRadius="10"/></Border.Effect>
            <StackPanel>
              <Border Background="#208A3C" Height="2" CornerRadius="1" Margin="0,0,0,10"/>
              <TextBlock x:Name="msg_tb" Foreground="#F4FAFF" FontSize="12" TextWrapping="Wrap" Margin="0,0,0,8"/>
              <TextBox x:Name="inp_tb" Background="#323D4D" Foreground="#F4FAFF"
                       BorderBrush="#404E62" BorderThickness="1" Padding="6,4"
                       FontSize="12" Margin="0,0,0,12"/>
              <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="ok_btn" Content="OK" Width="64" Height="26"
                        Background="#208A3C" Foreground="#F4FAFF" BorderThickness="0"
                        FontSize="11" Margin="0,0,6,0"/>
                <Button x:Name="cancel_btn" Content="Cancel" Width="64" Height="26"
                        Background="#404E62" Foreground="#F4FAFF" BorderThickness="0" FontSize="11"/>
              </StackPanel>
            </StackPanel>
          </Border>
        </Window>'''
        try:
            dlg = _Markup.XamlReader.Parse(xaml_str)
            dlg.FindName('msg_tb').Text = message
            inp = dlg.FindName('inp_tb')
            inp.Text = default or ''
            inp.Focus()
            inp.SelectAll()
            try: dlg.Owner = self
            except Exception: pass
            result = [None]
            def ok(s, e): result[0] = dlg.FindName('inp_tb').Text; dlg.Close()
            def cancel(s, e): dlg.Close()
            dlg.FindName('ok_btn').Click    += ok
            dlg.FindName('cancel_btn').Click += cancel
            dlg.ShowDialog()
            return result[0]
        except Exception as ex:
            self._status('Prompt error: {}'.format(ex))
            return default

    def _status(self, msg):
        try:
            self.status_tb.Text = msg
            import System.Windows.Threading as _SWT
            t = _SWT.DispatcherTimer()
            t.Interval = System.TimeSpan(0, 0, 3)
            def tick(s, e): 
                try: self.status_tb.Text = ''
                except Exception: pass
                t.Stop()
            t.Tick += tick; t.Start()
        except Exception: pass

    # ── Public API ────────────────────────────────────────────────
    def get_active_layout(self):
        self._flush_template()
        return {
            'template':    self._active_tmpl,
            'rev_count':   self._rev_count,
            'col_pct':     self._col_pct,
            'page_w_mm':   self._page_w_mm,
            'rows':        copy.deepcopy(self._rows),
            'text_styles': copy.deepcopy(self._text_styles),
            'logo_path':   self._logo_path,
        }


# ── Entry point ───────────────────────────────────────────────────
if __name__ == '__main__':
    script_dir = os.path.dirname(os.path.abspath(__file__))
    win = LayoutSettingsWindow(script_dir)
    win.ShowDialog()
