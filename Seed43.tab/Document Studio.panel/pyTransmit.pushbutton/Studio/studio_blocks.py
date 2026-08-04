# -*- coding: utf-8 -*-
# studio_blocks.py
#
# Block catalog + per-cell renderer for pyTransmit Studio.
#
# A deliberate fork of ../Layout/LayoutSettings.py's PALETTE and
# PreviewBuilder._block_el, NOT an import of them, because:
#   1. LayoutSettings.py must stay untouched (existing tool).
#   2. PreviewBuilder reads a module-level DUMMY dict rather than taking data
#      as a parameter, so pointing DUMMY at live Revit data would mutate
#      shared global state and corrupt the Layout Builder's preview if both
#      tools are open. Studio's renderer takes `data` explicitly instead.

import os

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

import System.Windows as _SW
import System.Windows.Controls as _SWC
import System.Windows.Documents as _SWD
import System.Windows.Media as _SWM


# -- Palette definition -------------------------------------------------------
# (type, label, icon, group, subtext) - same catalog as the Layout Builder so
# both tools present a familiar set of content blocks.
PALETTE = [
    ('__grp__',        'Layout',                '', 'layout',   ''),
    ('logo',           'Logo',                  'logo_image',         'layout',   ''),
    ('text',           'Text',                  'text',               'layout',   'body / free text'),
    ('page_count',     'Page Count',            'revisions',          'layout',   'page X of Y'),
    ('issue_date',     'Current Issue Date',    'select_revision',    'layout',   'latest revision date'),
    ('blank',          'Blank Spacer',          'show_all',           'layout',   'empty filler row'),

    ('__grp__',        'Project Info',          '', 'project',  ''),
    ('proj_org',       'Organisation',          'client_list',        'project',  ''),
    ('proj_client',    'Client',                'recipients',         'project',  ''),
    ('proj_number',    'Project Number',        'parameters_data',    'project',  ''),
    ('proj_name',      'Project Name',          'sheet_document',     'project',  ''),

    ('__grp__',        'Distribution',          '', 'dist',     ''),
    ('sent_to',        'Sent To',               'distribution_list',  'dist',     'company names'),
    ('attn_to',        'Attention To',          'issued_by',          'dist',     'contact names'),

    ('__grp__',        'Documentation',         '', 'docs',     ''),
    ('sheet_number',   'Sheet Number',          'parameters_data',    'docs',     'drawing numbers'),
    ('sheet_desc',     'Sheet Description',     'text',               'docs',     'drawing titles'),
    ('drawing_group',  'Drawing Group',         'families_components','docs',     'group headers'),

    ('__grp__',        'Revision',              '', 'revision', ''),
    ('spine_dates',    'Date of Issue',         'select_revision',    'revision', 'rotated headers'),
    ('spine_rev',      'Revision Marks',        'revisions',          'revision', 'A B C per drawing'),
    ('spine_initials', 'Initials',              'issued_by',          'revision', 'issued by'),
    ('spine_reason',   'Reason for Issue Code', 'reason_for_issue',   'revision', 'per revision'),
    ('spine_method',   'Method of Issue Code',  'method_of_issue',    'revision', 'per revision'),
    ('spine_copies',     'Number of Copies',      'copy_match',       'revision', 'per recipient'),
    ('spine_doc_type',   'Document Type Code',    'sheet_document',   'revision', 'per revision'),
    ('spine_print_size', 'Print Size Code',       'print_size',       'revision', 'per revision'),

    ('__grp__',            'Metadata',                '', 'meta',                    ''),
    ('reason_list',        'Reason (Vertical)',       'reason_for_issue',   'meta', 'code + label rows'),
    ('reason_list_row',    'Reason (Horizontal)',     'reason_for_issue',   'meta', 'spread across one row'),
    ('method_list',        'Method (Vertical)',       'method_of_issue',    'meta', 'code + label rows'),
    ('method_list_row',    'Method (Horizontal)',     'method_of_issue',    'meta', 'spread across one row'),
    ('doc_type',           'Document Type',           'sheet_document',     'meta', ''),
    ('print_size',         'Print Size',              'print_size',         'meta', ''),
]

# Palette id -> (actual block type, extra defaults applied on insert), so one
# block type can appear in the palette more than once pre-configured
# differently: the Reason/Method legends render as stacked code+label rows
# ('list') or across a single row ('row'), the same two placements the Layout
# Builder exposes via list_style. Unlisted ids map to themselves.
PALETTE_SPEC = {
    'reason_list':     ('reason_list', {'list_style': 'list'}),
    'reason_list_row': ('reason_list', {'list_style': 'row'}),
    'method_list':     ('method_list', {'list_style': 'list'}),
    'method_list_row': ('method_list', {'list_style': 'row'}),
    # A brand-new text block with content='' renders as a blank cell, which
    # reads as "the block didn't get placed" - seed it with visible text.
    'text':            ('text', {'content': 'Text'}),
}

TYPE_NAMES = {p[0]: p[1] for p in PALETTE if p[0] != '__grp__'}
TYPE_ICONS = {p[0]: p[2] for p in PALETTE if p[0] != '__grp__'}
TYPE_GROUP = {p[0]: p[3] for p in PALETTE if p[0] != '__grp__'}

GROUP_COLORS = {
    'layout':   '#8A96A8',
    'project':  '#2980B9',
    'dist':     '#27AE60',
    'docs':     '#8E44AD',
    'revision': '#E67E22',
    'meta':     '#16A085',
}

GROUP_LABELS = {
    'layout':   'LAYOUT',
    'project':  'PROJECT INFO',
    'dist':     'DISTRIBUTION',
    'docs':     'DOCUMENTATION',
    'revision': 'REVISION',
    'meta':     'METADATA',
}

# Font sizes are stored in mm (the renderer works in mm throughout, since
# this lays out a printed page), but are shown to the user in points -
# that's the unit Excel uses, and typing "9" should mean 9pt, not 9mm.
MM_PER_PT = 0.352778


def pt_to_mm(pt):
    return float(pt) * MM_PER_PT


def mm_to_pt(mm):
    return float(mm) / MM_PER_PT


# Excel's own font-size dropdown list.
FONT_SIZES_PT = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72]

DEFAULT_FONT = {'font': 'Arial', 'size_mm': pt_to_mm(9), 'bold': False,
                'italic': False, 'underline': False, 'color': '#000000'}

KV_MAP = {
    'proj_org':    'Organisation',  'proj_client': 'Client',
    'proj_number': 'Project No',    'proj_name':   'Project',
    'doc_type':    'Document Type', 'print_size':  'Print Size',
}


def new_block(t, **kw):
    """Same shape as LayoutSettings._mk() - the schema is shared even though
    the code isn't, so a block dict looks familiar to anyone who knows the
    original tool."""
    d = {
        'type': t, 'label': '', 'enabled': True,
        'just': 'left', 'v_just': 'middle',
        'borders': {'t': True, 'b': True, 'l': False, 'r': False},
        'data_borders': {'h': True, 'v': True},
        'list_style': 'list', 'bg_color': None,
        'alt_rows': False, 'alt_color': '#F5F7FA',
        'content': '', 'rotation': 0,
        'prefix': '', 'suffix': '',
        'page_format': 'Page X of Y',
        'date_format': 'dd/MM/yyyy',
        'height_pct': None,   # 'blank' spacer only
        # Revision blocks only. 'auto' = work it out from which revision
        # column this cell sits in, so placing the same block one column
        # right reads the next revision. An int pins it to that revision.
        'rev_index': 'auto',
    }
    d.update(DEFAULT_FONT)
    d.update(kw)
    return d


def new_block_from_palette(palette_id):
    """Build the block a given palette button should insert, applying any
    preset defaults from PALETTE_SPEC (see there)."""
    block_type, extra = PALETTE_SPEC.get(palette_id, (palette_id, {}))
    return new_block(block_type, **extra)


# -- Colour helpers -----------------------------------------------------------
def _brush(r, g, b, a=255):
    return _SWM.SolidColorBrush(_SWM.Color.FromArgb(a, r, g, b))


def _hbrush(h):
    h = (h or '#000000').lstrip('#')
    if len(h) == 3:
        h = h[0] * 2 + h[1] * 2 + h[2] * 2
    return _brush(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


WHITE = _SWM.Brushes.White
BLACK = _SWM.Brushes.Black
MUTED = _hbrush('#8A96A8')


def _v_align(v_just):
    m = {'top': _SW.VerticalAlignment.Top,
         'middle': _SW.VerticalAlignment.Center,
         'bottom': _SW.VerticalAlignment.Bottom}
    return m.get(v_just or 'middle', _SW.VerticalAlignment.Center)


def _h_align(just):
    m = {'left': _SW.HorizontalAlignment.Left,
         'center': _SW.HorizontalAlignment.Center,
         'right': _SW.HorizontalAlignment.Right}
    return m.get(just or 'left', _SW.HorizontalAlignment.Left)


def _apply_text_style(el, block_or_font, just, scale):
    """`block_or_font` is either a block dict (its own font/size_mm/bold/
    italic/underline/color fields are read directly - no named-preset
    indirection, each cell carries its own formatting like a real
    spreadsheet cell) or a plain font dict for auxiliary labels that aren't
    tied to a single block (falls back to DEFAULT_FONT for anything missing)."""
    st = block_or_font or DEFAULT_FONT
    try:
        el.FontFamily = _SWM.FontFamily(st.get('font', DEFAULT_FONT['font']))
    except Exception:
        pass
    el.FontSize = max(4, st.get('size_mm', DEFAULT_FONT['size_mm']) * scale)
    el.FontWeight = _SW.FontWeights.Bold if st.get('bold') else _SW.FontWeights.Normal
    el.FontStyle = _SW.FontStyles.Italic if st.get('italic') else _SW.FontStyles.Normal
    el.TextDecorations = _SW.TextDecorations.Underline if st.get('underline') else None
    try:
        el.Foreground = _hbrush(st.get('color', DEFAULT_FONT['color']))
    except Exception:
        el.Foreground = BLACK
    ta_map = {'left': _SW.TextAlignment.Left, 'center': _SW.TextAlignment.Center, 'right': _SW.TextAlignment.Right}
    el.TextAlignment = ta_map.get(just or 'left', _SW.TextAlignment.Left)


def block_display_summary(block):
    """Short human string for the formula bar / hover tooltips."""
    if not block:
        return ''
    t = block.get('type')
    if t == 'text':
        return block.get('content', '') or '(empty text)'
    name = TYPE_NAMES.get(t, t)
    if t in ('reason_list', 'method_list'):
        # These appear twice in the palette (vertical/horizontal), so report
        # the placement this particular block is actually using.
        base = name.split(' (')[0]
        placement = 'Horizontal' if block.get('list_style') == 'row' else 'Vertical'
        return '{} ({})'.format(base, placement)
    return name


def placeholder(block, scale, note=''):
    """Ghost label for a block that currently has nothing to draw - no live
    Revit data yet, an empty field, a revision that doesn't exist.

    Without this the cell renders completely empty, which is
    indistinguishable from "no block was ever placed here". This is a
    design tool, so what's *assigned* to a cell matters even when the model
    can't fill it in yet.
    """
    t = block.get('type', '')
    name = TYPE_NAMES.get(t, t)
    if t in ('reason_list', 'method_list'):
        name = name.split(' (')[0]
        name += ' (Horizontal)' if block.get('list_style') == 'row' else ' (Vertical)'

    outer = _SWC.Border()
    outer.BorderBrush = _hbrush('#C3CCD8')
    outer.BorderThickness = _SW.Thickness(1)
    outer.Background = _hbrush('#F1F4F8')
    outer.Margin = _SW.Thickness(1)
    outer.Padding = _SW.Thickness(3, 1, 3, 1)

    tb = _SWC.TextBlock()
    tb.Text = '{} {}'.format(name, note).strip()
    tb.TextTrimming = _SW.TextTrimming.CharacterEllipsis
    tb.TextWrapping = _SW.TextWrapping.Wrap
    tb.HorizontalAlignment = _SW.HorizontalAlignment.Center
    tb.VerticalAlignment = _SW.VerticalAlignment.Center
    tb.TextAlignment = _SW.TextAlignment.Center
    tb.FontStyle = _SW.FontStyles.Italic
    tb.Foreground = _hbrush('#7C8899')
    tb.FontSize = max(7, min(11, 2.3 * scale))
    outer.Child = tb

    g = _SWC.Grid()
    g.Children.Add(outer)
    return g


def render_block(block, data, scale, logo_path='', rev_index=0):
    """Render one block's content, given an explicit `data` dict shaped like
    Layout/LayoutSettings.py's DUMMY (proj_org, proj_client, distribution,
    revisions, reasons, methods, docs, ...). Returns a UIElement suitable as
    a cell's content, or None for an empty/disabled block."""
    if not block or not block.get('enabled', True):
        return None

    t = block.get('type', '')
    pad = max(1, int(2 * scale))
    thick = _SW.Thickness(pad)

    root = _SWC.Grid()
    if block.get('bg_color'):
        try:
            root.Background = _hbrush(block['bg_color'])
        except Exception:
            pass

    # -- Blank spacer ----------------------------------------------------
    # Deliberately renders nothing visible: it exists to reserve vertical
    # space, so a ghost label would defeat its purpose.
    if t == 'blank':
        spacer = _SWC.Border()
        pct = block.get('height_pct')
        if pct:
            spacer.MinHeight = max(2, int(6 * scale * float(pct) / 100.0))
        if block.get('bg_color'):
            try:
                spacer.Background = _hbrush(block['bg_color'])
            except Exception:
                pass
        root.Children.Add(spacer)
        return root

    # -- Text ------------------------------------------------------------
    if t == 'text':
        if not (block.get('content') or '').strip():
            return placeholder(block, scale, '(empty)')
        wrapper = _SWC.Border()
        wrapper.Padding = _SW.Thickness(pad * 2, pad, pad * 2, pad)
        wrapper.VerticalAlignment = _v_align(block.get('v_just'))
        wrapper.HorizontalAlignment = _SW.HorizontalAlignment.Stretch
        tb = _SWC.TextBlock()
        tb.Text = block.get('content', '')
        tb.TextWrapping = _SW.TextWrapping.Wrap
        _apply_text_style(tb, block, block.get('just', 'left'), scale)
        wrapper.Child = tb
        root.Children.Add(wrapper)
        return root

    # -- Logo --------------------------------------------------------------
    if t == 'logo':
        wrapper = _SWC.Border()
        wrapper.Padding = _SW.Thickness(pad * 2)
        wrapper.HorizontalAlignment = _h_align(block.get('just'))
        wrapper.VerticalAlignment = _v_align(block.get('v_just'))
        inner = _SWC.Border()
        inner.Background = _hbrush('#E8E8E8')
        inner.BorderBrush = _hbrush('#AAAAAA')
        inner.BorderThickness = _SW.Thickness(1)
        tb = _SWC.TextBlock()
        fname = os.path.basename(logo_path) if logo_path else ''
        tb.Text = '[ {} ]'.format(fname) if fname else '[ LOGO ]'
        tb.Padding = _SW.Thickness(pad * 2, pad, pad * 2, pad)
        _apply_text_style(tb, DEFAULT_FONT, 'center', scale)
        inner.Child = tb
        wrapper.Child = inner
        root.Children.Add(wrapper)
        return root

    # -- KV project fields --------------------------------------------------
    if t in KV_MAP:
        if not str(data.get(t, '') or '').strip():
            return placeholder(block, scale, '(no data)')
        wrapper = _SWC.Border()
        wrapper.Padding = _SW.Thickness(pad * 2, pad, pad * 2, pad)
        wrapper.VerticalAlignment = _v_align(block.get('v_just'))
        wrapper.HorizontalAlignment = _SW.HorizontalAlignment.Stretch
        tb = _SWC.TextBlock()
        tb.TextTrimming = _SW.TextTrimming.CharacterEllipsis
        lbl_run = _SWD.Run()
        lbl_run.Text = (block.get('label') or KV_MAP[t]) + ': '
        lbl_run.Foreground = _hbrush('#555555')
        val_run = _SWD.Run()
        val_run.Text = str(data.get(t, '') or '')
        val_run.FontWeight = _SW.FontWeights.Bold
        tb.Inlines.Add(lbl_run)
        tb.Inlines.Add(val_run)
        _apply_text_style(tb, block, block.get('just', 'left'), scale)
        wrapper.Child = tb
        root.Children.Add(wrapper)
        return root

    # -- Page count -----------------------------------------------------------
    if t == 'page_count':
        wrapper = _SWC.Border()
        wrapper.Padding = thick
        wrapper.HorizontalAlignment = _h_align(block.get('just'))
        wrapper.VerticalAlignment = _v_align(block.get('v_just'))
        tb = _SWC.TextBlock()
        total = len(data.get('docs', []))
        fmt = block.get('page_format', 'Page X of Y')
        val = {'Page X': 'Page 1', 'Page X of Y': 'Page 1 of {}'.format(total),
               'X of Y': '1 of {}'.format(total), 'X / Y': '1 / {}'.format(total)}.get(fmt, str(total))
        parts = [block.get('prefix', ''), val, block.get('suffix', '')]
        tb.Text = ' '.join(p for p in parts if p)
        _apply_text_style(tb, block, block.get('just', 'left'), scale)
        wrapper.Child = tb
        root.Children.Add(wrapper)
        return root

    # -- Current issue date -----------------------------------------------------
    if t == 'issue_date':
        wrapper = _SWC.Border()
        wrapper.Padding = thick
        wrapper.HorizontalAlignment = _h_align(block.get('just'))
        wrapper.VerticalAlignment = _v_align(block.get('v_just'))
        tb = _SWC.TextBlock()
        revs = data.get('revisions', [])
        if not revs:
            return placeholder(block, scale, '(no revisions)')
        date_str = revs[-1].get('date', '') if revs else ''
        parts = [block.get('prefix', ''), date_str, block.get('suffix', '')]
        tb.Text = ' '.join(p for p in parts if p) or '(no issued revisions yet)'
        _apply_text_style(tb, block, block.get('just', 'left'), scale)
        wrapper.Child = tb
        root.Children.Add(wrapper)
        return root

    # (doc_type / print_size are handled by the KV_MAP branch above - they're
    # listed in KV_MAP, so a separate branch for them here would be dead code.)

    # -- Row-per-item data blocks --------------------------------------------
    sp = _SWC.StackPanel()
    sp.HorizontalAlignment = _SW.HorizontalAlignment.Stretch
    alt = block.get('alt_rows', False)
    alt_c = block.get('alt_color', '#F5F7FA')
    just_val = block.get('just', 'left')
    data_row_h = max(10, int(block.get('size_mm', DEFAULT_FONT['size_mm']) * scale * 1.8 + 2 * pad))

    def tb_row(txt, alt_row=False, show_h=True):
        row_b = _SWC.Border()
        row_b.BorderBrush = _hbrush('#E8E8E8')
        row_b.BorderThickness = _SW.Thickness(0, 0, 0, 1 if show_h else 0)
        row_b.Height = data_row_h
        if alt_row:
            row_b.Background = _hbrush(alt_c)
        tb = _SWC.TextBlock()
        tb.Text = str(txt or '')
        tb.Padding = thick
        tb.TextTrimming = _SW.TextTrimming.CharacterEllipsis
        tb.VerticalAlignment = _SW.VerticalAlignment.Center
        _apply_text_style(tb, block, just_val, scale)
        row_b.Child = tb
        return row_b

    ROW_DATA = {
        'sent_to':      [d.get('to', '') for d in data.get('distribution', [])],
        'attn_to':      [d.get('attn', '') for d in data.get('distribution', [])],
        'sheet_number': [d.get('sheet', '') for d in data.get('docs', [])],
        'sheet_desc':   [d.get('desc', '') for d in data.get('docs', [])],
    }
    if t in ROW_DATA:
        db = block.get('data_borders', {'h': True, 'v': True})
        show_h = bool(db.get('h', True))
        items = ROW_DATA[t]
        if not [v for v in items if str(v or '').strip()]:
            # Nothing at all, or every row blank - either way an empty
            # stack of rows is indistinguishable from an empty cell.
            return placeholder(block, scale, '(no data)')
        for i, txt in enumerate(items):
            sp.Children.Add(tb_row(txt, alt and i % 2 == 1, show_h))
        sp.VerticalAlignment = _v_align(block.get('v_just'))
        root.Children.Add(sp)
        return root

    if t == 'drawing_group':
        db = block.get('data_borders', {'h': True, 'v': True})
        show_h = bool(db.get('h', True))
        last_g = ''
        docs = data.get('docs', [])
        if not docs:
            return placeholder(block, scale, '(no sheets)')
        for i, doc in enumerate(docs):
            sheet = doc.get('sheet', '')
            g = sheet.rstrip('0123456789.').rstrip('0123456789')
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
                _apply_text_style(grp_tb, dict(DEFAULT_FONT, bold=True, size_mm=2.6), just_val, scale)
                grp_b.Child = grp_tb
                sp.Children.Add(grp_b)
            empty_b = _SWC.Border()
            empty_b.Background = _hbrush(alt_c) if (alt and i % 2 == 1) else WHITE
            empty_b.BorderBrush = _hbrush('#E8E8E8')
            empty_b.BorderThickness = _SW.Thickness(0, 0, 0, 1 if show_h else 0)
            empty_b.Height = data_row_h
            sp.Children.Add(empty_b)
        sp.VerticalAlignment = _v_align(block.get('v_just'))
        root.Children.Add(sp)
        return root

    items_map = {'reason_list': data.get('reasons', []), 'method_list': data.get('methods', [])}
    if t in items_map:
        items = items_map[t]
        if not items:
            return placeholder(block, scale, '(none configured)')
        if block.get('list_style') == 'row':
            row_g = _SWC.Grid()
            row_g.Margin = _SW.Thickness(pad)
            for _ in (items or [None]):
                cd = _SWC.ColumnDefinition()
                cd.Width = _SW.GridLength(1, _SW.GridUnitType.Star)
                row_g.ColumnDefinitions.Add(cd)
            for idx, item in enumerate(items):
                item_tb = _SWC.TextBlock()
                item_tb.Text = '{}  {}'.format(item.get('code', ''), item.get('label', ''))
                item_tb.Padding = _SW.Thickness(pad, 1, pad, 1)
                _apply_text_style(item_tb, block, just_val, scale)
                _SWC.Grid.SetColumn(item_tb, idx)
                row_g.Children.Add(item_tb)
            sp.Children.Add(row_g)
        else:
            db = block.get('data_borders', {'h': True, 'v': True})
            show_h = bool(db.get('h', True))
            show_v = bool(db.get('v', True))
            for i, item in enumerate(items):
                row_g = _SWC.Grid()
                row_g.Background = _hbrush(alt_c) if (alt and i % 2 == 1) else WHITE
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
                t1 = _SWC.TextBlock(); t1.Text = item.get('code', ''); t1.Padding = thick
                _apply_text_style(t1, block, just_val, scale); c1.Child = t1
                c2 = _SWC.Border()
                t2 = _SWC.TextBlock(); t2.Text = item.get('label', ''); t2.Padding = thick
                _apply_text_style(t2, block, just_val, scale); c2.Child = t2
                _SWC.Grid.SetColumn(c1, 0); _SWC.Grid.SetColumn(c2, 1)
                row_g.Children.Add(c1); row_g.Children.Add(c2)
                row_b.Child = row_g
                sp.Children.Add(row_b)
        sp.VerticalAlignment = _v_align(block.get('v_just'))
        root.Children.Add(sp)
        return root

    # -- Spine blocks: ONE revision per cell -----------------------------------
    # Excel model - a revision owns its own column, so each cell renders one
    # revision instead of packing them all into a single cell like the Layout
    # Builder. `rev_index` binds the cell to a revision; the caller derives it
    # from column position (StudioSettings._rev_index_for_column), so the same
    # block dropped one column over reads the next revision.
    revs = data.get('revisions', [])
    if t.startswith('spine_'):
        rev_label = '(rev {})'.format(rev_index + 1)
        if rev_index >= len(revs):
            return placeholder(block, scale, rev_label)

        rev = revs[rev_index]
        sp_just = block.get('just', 'center')
        rotation = block.get('rotation', 270 if t == 'spine_dates' else 0)

        if t == 'spine_copies':
            # Copies are per recipient, so this is a column of counts for
            # THIS revision - one row per distribution entry.
            dist = data.get('distribution', [])
            vals = []
            for d in dist:
                per_rev = d.get('copies_by_rev') or []
                vals.append(str(per_rev[rev_index] if rev_index < len(per_rev)
                                else (d.get('copies', '') or '')) or '')
            # Recipients can exist while every copies figure is blank - that
            # renders as a stack of empty rows, indistinguishable from an
            # empty cell, so fall back to the ghost label.
            if not dist or not [v for v in vals if v.strip()]:
                return placeholder(block, scale, rev_label)
            sp_db = block.get('data_borders', {'h': True, 'v': True})
            show_h = bool(sp_db.get('h', True))
            row_h = max(10, int(block.get('size_mm', DEFAULT_FONT['size_mm']) * scale * 1.8 + 2 * pad))
            stack = _SWC.StackPanel()
            for i, d in enumerate(dist):
                rb = _SWC.Border()
                rb.Height = row_h
                rb.BorderBrush = _hbrush('#E8E8E8')
                rb.BorderThickness = _SW.Thickness(0, 0, 0, 1 if show_h else 0)
                ctb = _SWC.TextBlock()
                ctb.Text = vals[i] if i < len(vals) else ''
                ctb.VerticalAlignment = _SW.VerticalAlignment.Center
                _apply_text_style(ctb, block, sp_just, scale)
                rb.Child = ctb
                stack.Children.Add(rb)
            stack.VerticalAlignment = _v_align(block.get('v_just'))
            root.Children.Add(stack)
            return root

        if t == 'spine_rev':
            # Revision marks are per-document, so this renders a column of
            # marks (one row per sheet) for THIS revision only.
            docs = data.get('docs', [])
            if not docs:
                return placeholder(block, scale, rev_label)
            sp_db = block.get('data_borders', {'h': True, 'v': True})
            show_h = bool(sp_db.get('h', True))
            alt = block.get('alt_rows', False)
            alt_c = block.get('alt_color', '#F5F7FA')
            row_h = max(10, int(block.get('size_mm', DEFAULT_FONT['size_mm']) * scale * 1.8 + 2 * pad))
            stack = _SWC.StackPanel()
            for i, doc in enumerate(docs):
                marks = doc.get('revs', [])
                val = marks[rev_index] if rev_index < len(marks) else ''
                rb = _SWC.Border()
                rb.Height = row_h
                rb.BorderBrush = _hbrush('#E8E8E8')
                rb.BorderThickness = _SW.Thickness(0, 0, 0, 1 if show_h else 0)
                if alt and i % 2 == 1:
                    rb.Background = _hbrush(alt_c)
                mtb = _SWC.TextBlock()
                mtb.Text = str(val or '')
                mtb.VerticalAlignment = _SW.VerticalAlignment.Center
                _apply_text_style(mtb, block, sp_just, scale)
                rb.Child = mtb
                stack.Children.Add(rb)
            stack.VerticalAlignment = _v_align(block.get('v_just'))
            root.Children.Add(stack)
            return root

        field_map = {
            'spine_dates': 'date', 'spine_initials': 'initials',
            'spine_reason': 'reason', 'spine_method': 'method',
            'spine_doc_type': 'doc_format', 'spine_print_size': 'paper_size',
        }
        val = str(rev.get(field_map.get(t, 'rev'), '') or '')
        if not val:
            return placeholder(block, scale, rev_label)

        wrapper = _SWC.Border()
        wrapper.Padding = thick
        wrapper.HorizontalAlignment = _h_align('center' if rotation == 270 else block.get('just'))
        wrapper.VerticalAlignment = _v_align(block.get('v_just'))
        tb = _SWC.TextBlock()
        tb.Text = val
        if rotation == 270:
            tb.LayoutTransform = _SWM.RotateTransform(270)
        _apply_text_style(tb, block, sp_just, scale)
        wrapper.Child = tb
        root.Children.Add(wrapper)
        return root

    # Fallback
    fb_tb = _SWC.TextBlock()
    fb_tb.Text = TYPE_NAMES.get(t, t)
    fb_tb.Foreground = MUTED
    fb_tb.Padding = thick
    _apply_text_style(fb_tb, DEFAULT_FONT, 'left', scale)
    root.Children.Add(fb_tb)
    return root
