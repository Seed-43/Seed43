# -*- coding: utf-8 -*-
"""
layout_Style_Algoruthum.py
===========================
General border and cell colour resolution algorithm for pyTransmit exporters.

Used by: script_create_excel.py, script_create_schedule.py, and future exporters.

Resolves:
  - On-beats-off shared edge logic for borders (t/b/l/r)
  - Cell background colour from block bg_color
  - Cell foreground colour from text_style
  - Interior border suppression for merged cells (l/r hidden on interior cols)

The algorithm produces a flat dict keyed by (layout_row_index, canvas_col_index)
with resolved border and colour values. Exporters map these to their own cell API.

Usage
-----
    from layout_Style_Algoruthum import compute_cell_styles

    styles = compute_cell_styles(
        rows         = ROWS,           # list of layout row dicts
        groups       = GROUPS,         # list of (start_ri, end_ri) tuples
        text_styles  = TEXT_STYLES,    # dict of text style name -> {color, bg, ...}
        max_rev_col  = 3,              # canvas col index of rev columns (D=3)
        n_revs       = MAX_REVS,       # number of revision columns
    )

    # styles[(ri, ci)] = {
    #     't': bool,   # top border on
    #     'b': bool,   # bottom border on
    #     'l': bool,   # left border on
    #     'r': bool,   # right border on
    #     'bg': str or None,   # background hex colour e.g. '#2B3340'
    #     'fg': str or None,   # foreground/text hex colour
    #     'interior_l': bool,  # suppress left (interior of merged span)
    #     'interior_r': bool,  # suppress right (interior of merged span)
    # }
"""


def compute_cell_styles_from_grid(rows, groups, hlines, vlines,
                                   text_styles=None, max_rev_col=3, n_revs=10):
    """
    Compute cell styles using the new hlines/vlines border grid.

    hlines: {ri: [bool*4]} — bottom border of each col in row ri
    vlines: {ri: [bool*5]} — right edge of vline positions 0-4 in row ri
            pos 0=outer-left, 1=A-B, 2=B-C, 3=C-D, 4=outer-right

    Returns same format as compute_cell_styles.
    """
    text_styles = text_styles or {}
    styles = {}

    def _b(ri, ci):
        # hlines[ri] = TOP of row ri, hlines[ri+1] = BOTTOM of row ri
        row_hl = hlines.get(ri + 1, [False]*4)
        return bool(row_hl[ci]) if ci < len(row_hl) else False

    def _t(ri, ci):
        # top of row ri
        row_hl = hlines.get(ri, [False]*4)
        return bool(row_hl[ci]) if ci < len(row_hl) else False

    def _l(ri, pos):
        row_vl = vlines.get(ri, [False]*5)
        return bool(row_vl[pos]) if pos < len(row_vl) else False

    for ri, row in enumerate(rows):
        if row.get('section') == 'footer':
            continue
        blocks = row.get('blocks', [])
        ci_canvas = 0

        while ci_canvas < 4:
            block = blocks[ci_canvas] if ci_canvas < len(blocks) else None
            if block is None:
                ci_canvas += 1
                continue

            span = int(block.get('span', 1))
            bg   = _resolve_bg(block, text_styles)
            fg   = _resolve_fg(block, text_styles)

            # vline position to the left of this block = ci_canvas
            # vline position to the right = ci_canvas + span
            left_pos  = ci_canvas        # outer or between-col vline
            right_pos = ci_canvas + span # next vline

            show_l = _l(ri, left_pos)
            show_r = _l(ri, right_pos)

            if ci_canvas >= max_rev_col:
                _DATA_TYPES_REV = {'sent_to','attn_to','sheet_number','sheet_desc',
                                   'spine_rev','spine_dates','spine_copies',
                                   'reason_list','method_list','spine_initials',
                                   'spine_reason','spine_method','spine_doc_type','spine_print_size'}
                _is_data_rev = block.get('type','') in _DATA_TYPES_REV
                v_grid = block.get('data_borders', {}).get('v', False) if _is_data_rev else False
                for rev_i in range(n_revs):
                    col_key = max_rev_col + rev_i
                    is_first = (rev_i == 0)
                    is_last  = (rev_i == n_revs - 1)
                    styles[(ri, col_key)] = {
                        't': _t(ri, max_rev_col),
                        'b': _b(ri, max_rev_col),
                        'l': show_l if is_first else v_grid,
                        'r': show_r if is_last  else v_grid,
                        'bg': bg, 'fg': fg,
                        'interior_l': not is_first,
                        'interior_r': not is_last,
                        'block': block,
                    }
            else:
                ci_end = min(ci_canvas + span - 1, max_rev_col - 1)
                _DATA_TYPES = {'sent_to','attn_to','sheet_number','sheet_desc',
                               'spine_rev','spine_dates','spine_copies',
                               'reason_list','method_list','spine_initials',
                               'spine_reason','spine_method','spine_doc_type','spine_print_size'}
                _is_data = block.get('type','') in _DATA_TYPES
                v_grid_fixed = block.get('data_borders', {}).get('v', False) if _is_data else False
                _other_blocks = [b for i2, b in enumerate(blocks) if b and i2 != ci_canvas]
                _has_neighbours = len(_other_blocks) > 0 and _is_data

                for ci in range(ci_canvas, min(ci_canvas + span, max_rev_col)):
                    is_first_col = (ci == ci_canvas)
                    is_last_col  = (ci == ci_end or ci == max_rev_col - 1)
                    _l_val = show_l if is_first_col else (v_grid_fixed if _has_neighbours else False)
                    _r_val = show_r if is_last_col  else (v_grid_fixed if _has_neighbours else False)
                    styles[(ri, ci)] = {
                        't': _t(ri, ci),
                        'b': _b(ri, ci),
                        'l': _l_val,
                        'r': _r_val,
                        'bg': bg, 'fg': fg,
                        'interior_l': not is_first_col,
                        'interior_r': not is_last_col,
                        'block': block,
                    }

            ci_canvas += span

    return styles


def compute_cell_styles(rows, groups, text_styles=None, max_rev_col=3, n_revs=10):
    """
    Compute per-cell border and colour styles from layout JSON rows.

    Parameters
    ----------
    rows : list of dict
        Layout row dicts with 'blocks', 'merge_down', 'section'.
    groups : list of (int, int)
        (start_ri, end_ri) group spans from get_groups().
    text_styles : dict, optional
        {style_name: {color, bg_color, bold, ...}} from LAYOUT.text_styles.
    max_rev_col : int
        Canvas column index where revision columns start (default 3 = col D).
    n_revs : int
        Number of revision columns.

    Returns
    -------
    dict : {(ri, ci): cell_style_dict}
        ri = layout row index, ci = canvas column index (0-3 for A-D, or
        expanded rev cols as 3, 3+1, 3+2, ... 3+n_revs-1)
    """
    text_styles = text_styles or {}
    styles = {}

    # ── Pass 1: resolve raw border and colour per block ──────────────────────
    # Build a map of (ri, ci) -> raw block info before shared-edge resolution
    raw = {}  # (ri, ci) -> {'brd': {t,b,l,r}, 'bg': str|None, 'fg': str|None, 'block': dict}

    for ri, row in enumerate(rows):
        if row.get('section') == 'footer':
            continue
        blocks = row.get('blocks', [])
        ci_canvas = 0

        while ci_canvas < 4:
            block = blocks[ci_canvas] if ci_canvas < len(blocks) else None
            if block is None:
                ci_canvas += 1
                continue

            span = int(block.get('span', 1))
            brd  = block.get('borders', {})
            bg   = _resolve_bg(block, text_styles)
            fg   = _resolve_fg(block, text_styles)

            # For rev columns (ci_canvas == max_rev_col), expand to n_revs cols
            if ci_canvas >= max_rev_col:
                # Only data/spine block types use v_grid for interior verticals
                _DATA_TYPES = (
                    'spine_dates', 'spine_rev', 'spine_initials', 'spine_reason',
                    'spine_method', 'spine_doc_type', 'spine_print_size',
                    'spine_copies', 'sheet_number', 'sheet_desc', 'sent_to',
                    'attn_to', 'reason_list', 'method_list',
                )
                _block_type = block.get('type', '')
                v_grid = (block.get('data_borders', {}).get('v', False)
                          if _block_type in _DATA_TYPES else False)
                for rev_i in range(n_revs):
                    col_key = max_rev_col + rev_i
                    is_first = (rev_i == 0)
                    is_last  = (rev_i == n_revs - 1)
                    raw[(ri, col_key)] = {
                        'brd': {
                            't': brd.get('t', False),
                            'b': brd.get('b', False),
                            'l': brd.get('l', False) if is_first else v_grid,
                            'r': brd.get('r', False) if is_last  else v_grid,
                        },
                        'bg': bg, 'fg': fg, 'block': block,
                        'span': 1, 'ci_start': col_key, 'ci_end': col_key,
                    }
            else:
                # Fixed columns A/B/C — span may cover multiple canvas cols
                ci_end = min(ci_canvas + span - 1, max_rev_col - 1)
                if ci_canvas + span - 1 >= max_rev_col:
                    ci_end = max_rev_col + n_revs - 1  # spans into rev cols

                v_grid_fixed = block.get('data_borders', {}).get('v', False)
                # Only apply v_grid between cols when there are other blocks in the row
                _other_blocks = [b for i, b in enumerate(row.get('blocks', []))
                                 if b and i != ci_canvas]
                _has_neighbours = len(_other_blocks) > 0

                for ci in range(ci_canvas, min(ci_canvas + span, max_rev_col)):
                    is_first_col = (ci == ci_canvas)
                    is_last_col  = (ci == ci_end or ci == max_rev_col - 1)
                    # Left: from brd on first col; v_grid between adjacent fixed blocks
                    _l = brd.get('l', False) if is_first_col else (v_grid_fixed if _has_neighbours else False)
                    # Right: from brd on last col; v_grid between adjacent fixed blocks
                    _r = brd.get('r', False) if is_last_col else (v_grid_fixed if _has_neighbours else False)
                    raw[(ri, ci)] = {
                        'brd': {
                            't': brd.get('t', False),
                            'b': brd.get('b', False),
                            'l': _l,
                            'r': _r,
                        },
                        'bg': bg, 'fg': fg, 'block': block,
                        'span': span, 'ci_start': ci_canvas, 'ci_end': ci_end,
                        'interior_l': not is_first_col,
                        'interior_r': not is_last_col,
                    }

            ci_canvas += span

    # ── Pass 2: on-beats-off shared edge resolution ──────────────────────────
    # A shared horizontal edge (bottom of row N = top of row N+1) is ON
    # if EITHER side wants it on — but only if the block on that side
    # participates in bordering (has at least one border set).

    def _has_any_border(ri, ci):
        cell = raw.get((ri, ci))
        if not cell:
            return False
        b = cell['brd']
        return any(b.get(k, False) for k in ('t', 'b', 'l', 'r'))

    # ── Pass 2a: same-row t/b propagation ───────────────────────────────────
    # If any block in a row has t or b, propagate to all blocks in that row.
    # This ensures "Revision" text gets same bottom as spine_rev in same row.
    # Exception: reason_list/method_list blocks are content within a vertical
    # span — they should not receive same-row propagation from other columns.
    _CONTENT_TYPES = ('reason_list', 'method_list')
    _row_top = {}  # ri -> bool
    _row_bot = {}  # ri -> bool
    for (ri, ci), cell in raw.items():
        _bt = cell.get('block', {}).get('type', '') if cell.get('block') else ''
        if _bt in _CONTENT_TYPES: continue  # don't let content blocks drive row-level propagation
        if cell['brd'].get('t', False): _row_top[ri] = True
        if cell['brd'].get('b', False): _row_bot[ri] = True

    # ── Pass 2b: on-beats-off shared edge resolution ─────────────────────────
    # For non-content blocks: use row-level maps (_row_bot/_row_top) so that
    # a border on any col in a row propagates to all cols (e.g. spine_rev b:True
    # gives the "Revision" text cell its bottom line).
    # For content blocks (reason_list/method_list): only check same-ci neighbours
    # to avoid bleeding from spine_dates b:True into the legend content above/below.
    # blank blocks are fully immune.
    for (ri, ci), cell in raw.items():
        resolved_brd = dict(cell['brd'])
        _btype = cell.get('block', {}).get('type', '') if cell.get('block') else ''

        if _btype == 'blank':
            pass  # immune

        elif _btype in _CONTENT_TYPES:
            # Same-ci only — no row-level cross-column bleeding
            above_ci = raw.get((ri - 1, ci))
            if above_ci and above_ci['brd'].get('b', False):
                resolved_brd['t'] = True
            below_ci = raw.get((ri + 1, ci))
            if below_ci and below_ci['brd'].get('t', False):
                resolved_brd['b'] = True

        else:
            # Top: on if this cell wants t OR any block in row above has b
            if _row_top.get(ri) or _row_bot.get(ri - 1):
                resolved_brd['t'] = True
            # Bottom: on if THIS CELL wants b OR any block in row below has t
            # Note: _row_bot[ri] is NOT used — it would bleed b:True from spine
            # cols (ci=2/3) into label cols (ci=0/1) in the same row.
            if cell['brd'].get('b', False) or _row_top.get(ri + 1):
                resolved_brd['b'] = True

        styles[(ri, ci)] = {
                't':          resolved_brd.get('t', False),
                'b':          resolved_brd.get('b', False),
                'l':          resolved_brd.get('l', False),
                'r':          resolved_brd.get('r', False),
                'bg':         cell.get('bg'),
                'fg':         cell.get('fg'),
                'interior_l': cell.get('interior_l', False),
                'interior_r': cell.get('interior_r', False),
                'block':      cell.get('block'),
            }

    return styles


def _resolve_bg(block, text_styles):
    """Resolve background colour from block or its text style."""
    # Block-level bg_color takes priority
    bg = block.get('bg_color')
    if bg:
        return _normalise_hex(bg)
    # Fall back to text style bg
    style_name = block.get('text_style', 'Data')
    ts = text_styles.get(style_name, {})
    bg = ts.get('bg_color') or ts.get('bg')
    return _normalise_hex(bg) if bg else None


def _resolve_fg(block, text_styles):
    """Resolve foreground/text colour from block's text style."""
    style_name = block.get('text_style', 'Data')
    ts = text_styles.get(style_name, {})
    fg = ts.get('color') or ts.get('fg') or ts.get('text_color')
    if fg:
        return _normalise_hex(fg)
    return '#000000'


def _normalise_hex(h):
    """Normalise a hex colour string to '#RRGGBB' format."""
    if not h:
        return None
    try:
        h = str(h).strip().lstrip('#')
        if len(h) == 3:
            h = h[0]*2 + h[1]*2 + h[2]*2
        if len(h) == 6:
            return '#{}'.format(h.upper())
    except Exception:
        pass
    return None


def get_groups(rows):
    """
    Build group spans from merge_down flags.
    Returns list of (start_ri, end_ri) tuples.
    """
    groups = []
    i = 0
    while i < len(rows):
        start = i
        while i < len(rows) - 1 and rows[i].get('merge_down', False):
            i += 1
        groups.append((start, i))
        i += 1
    return groups


def style_for_excel(cell_style):
    """
    Convert a cell_style dict to xlsxwriter format kwargs.

    Returns dict suitable for passing to workbook.add_format() or
    merging into an existing format override dict.
    """
    fmt = {}
    if cell_style.get('t'): fmt['top']    = 1
    if cell_style.get('b'): fmt['bottom'] = 1
    if cell_style.get('l'): fmt['left']   = 1
    if cell_style.get('r'): fmt['right']  = 1
    if cell_style.get('bg'):
        fmt['bg_color'] = cell_style['bg']
    if cell_style.get('fg'):
        fmt['font_color'] = cell_style['fg']
    return fmt


def style_for_revit(cell_style, on_id, off_id):
    """
    Convert a cell_style dict to Revit schedule border ElementIds.

    Parameters
    ----------
    on_id  : ElementId — line style for visible borders (e.g. 'pyT On')
    off_id : ElementId — line style for hidden borders (e.g. 'pyT Off')

    Returns dict with keys matching TableCellStyle border properties:
        BorderTopLineStyle, BorderBottomLineStyle,
        BorderLeftLineStyle, BorderRightLineStyle
    """
    def _pick(show):
        return on_id if show else off_id

    return {
        'BorderTopLineStyle':    _pick(cell_style.get('t', False)),
        'BorderBottomLineStyle': _pick(cell_style.get('b', False)),
        'BorderLeftLineStyle':   _pick(cell_style.get('l', False)),
        'BorderRightLineStyle':  _pick(cell_style.get('r', False)),
    }
