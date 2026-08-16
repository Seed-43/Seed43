# -*- coding: utf-8 -*-
"""convert_layout_templates.py

One-off/rerunnable converter: reads the Layout Builder's templates
(../Layout/Layouts/*.json) and writes the pyTransmit Studio equivalents
into ./studio_layouts/.

Run with ordinary CPython from the Studio folder:

    python convert_layout_templates.py

It imports nothing from the WPF modules, so it does NOT need Revit or
IronPython.

The two formats differ in one structural way that drives most of this file:

    Layout Builder  every row has exactly 4 slots; slot 3 is a "spine" that
                    internally splits into `rev_count` revision columns.
    Studio          a real grid; one revision per column.

So slots 0-2 map straight across, and slot 3 fans out into `rev_count`
real columns. A spine block in slot 3 becomes one block per revision
column (each 'auto', so it binds to its own revision); a non-spine block
in slot 3 becomes a single cell spanning all of them.
"""

import os

import sys as _sys
_PT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PT_ROOT not in _sys.path:
    _sys.path.insert(0, _PT_ROOT)
from pytransmit_paths import LAYOUTS_DIR, STUDIO_LAYOUTS_DIR
import json
import copy

HERE = os.path.dirname(os.path.abspath(__file__))
# Both live in .user now, not beside the tool.
SRC_DIR = LAYOUTS_DIR
OUT_DIR = STUDIO_LAYOUTS_DIR

MM_PER_PT = 0.352778
PAGE_SIZES = {'A4': (210, 297), 'A3': (297, 420), 'A2': (420, 594),
              'Letter': (216, 279), 'Tabloid': (279, 432)}

SPINE_TYPES = set(['spine_dates', 'spine_rev', 'spine_initials', 'spine_reason',
                   'spine_method', 'spine_copies', 'spine_doc_type', 'spine_print_size'])

# Layout Builder row height isn't stored - it auto-sizes to content. Studio
# needs explicit mm, so heights are inferred from what the row contains.
H_DEFAULT = 7.0
H_TITLE = 12.0
H_DATA = 30.0      # blocks that render a list of rows (sheets/recipients)
DATA_BLOCKS = set(['sent_to', 'attn_to', 'sheet_number', 'sheet_desc',
                   'drawing_group', 'spine_rev', 'spine_copies', 'reason_list',
                   'method_list'])


def studio_block(lb_block, text_styles):
    """Convert one Layout Builder block into a Studio block.

    The big change: Layout Builder blocks reference a *named* text style
    ('Title'/'Header'/'Data'); Studio cells carry their own font, so the
    named style is resolved and baked in here.
    """
    style_name = lb_block.get('text_style', 'Data')
    st = text_styles.get(style_name) or {}
    b = {
        'type': lb_block.get('type'),
        'label': lb_block.get('label', ''),
        'enabled': lb_block.get('enabled', True),
        'just': lb_block.get('just', 'left'),
        'v_just': lb_block.get('v_just', 'middle'),
        'borders': dict(lb_block.get('borders') or {'t': True, 'b': True, 'l': False, 'r': False}),
        'data_borders': dict(lb_block.get('data_borders') or {'h': True, 'v': True}),
        'list_style': lb_block.get('list_style', 'list'),
        'bg_color': lb_block.get('bg_color'),
        'alt_rows': lb_block.get('alt_rows', False),
        'alt_color': lb_block.get('alt_color', '#F5F7FA'),
        'content': lb_block.get('content', ''),
        'rotation': lb_block.get('rotation', 0),
        'prefix': lb_block.get('prefix', ''),
        'suffix': lb_block.get('suffix', ''),
        'page_format': lb_block.get('page_format', 'Page X of Y'),
        'date_format': lb_block.get('date_format', 'dd/MM/yyyy'),
        'height_pct': lb_block.get('height_pct'),
        'rev_index': 'auto',
        # resolved font
        'font': st.get('font', 'Arial'),
        'size_mm': st.get('size_mm', 2.3),
        'bold': st.get('bold', False),
        'italic': st.get('italic', False),
        'underline': st.get('underline', False),
        'color': st.get('color', '#000000'),
    }
    return b


def row_height_for(lb_row):
    h = H_DEFAULT
    for b in lb_row.get('blocks', []):
        if not b:
            continue
        t = b.get('type')
        if t in DATA_BLOCKS:
            h = max(h, H_DATA)
        elif b.get('text_style') == 'Title':
            h = max(h, H_TITLE)
        elif t == 'logo':
            h = max(h, 14.0)
        elif t == 'blank':
            pct = b.get('height_pct')
            h = max(h, H_DEFAULT * (float(pct) / 100.0) if pct else H_DEFAULT)
    return h


def section_for(lb_row):
    """Layout Builder's (section, repeat_every_page) pair -> Studio section.

    Only the 'every page' variants become repeating bands. The
    'first page only' variants are ordinary rows in a real grid: a row at
    the top of the sheet already appears on page 1 and nowhere else, so
    there is nothing to mark.
    """
    sec = lb_row.get('section', 'body')
    rep = bool(lb_row.get('repeat_every_page'))
    if sec == 'repeat_header' and rep:
        return 'repeat_top'
    if sec == 'footer' and rep:
        return 'repeat_bottom'
    return 'body'


def convert(path):
    lb = json.load(open(path, 'r'))
    name = lb.get('template') or os.path.splitext(os.path.basename(path))[0]
    rev_count = int(lb.get('rev_count', 10))
    col_pct = lb.get('col_pct', [22, 28, 20])
    page_w = lb.get('page_w_mm', 210)
    page_h = lb.get('page_h_mm', 297)
    orientation = lb.get('orientation', 'portrait')
    text_styles = lb.get('text_styles') or {}
    lb_rows = lb.get('rows', [])

    # ---- columns: A,B,C straight across; slot D fans into rev_count cols ----
    margin = 10.0                    # Layout Builder's implicit side margin
    inner_w = page_w - 2 * margin
    a = inner_w * col_pct[0] / 100.0
    b = inner_w * col_pct[1] / 100.0
    c = inner_w * col_pct[2] / 100.0
    spine_w = max(inner_w - a - b - c, rev_count * 4.0)
    rev_w = spine_w / float(rev_count)
    col_widths = [round(a, 2), round(b, 2), round(c, 2)] + [round(rev_w, 2)] * rev_count
    n_cols = 3 + rev_count

    n_rows = len(lb_rows)
    row_heights = [row_height_for(r) for r in lb_rows]
    row_sections = [section_for(r) for r in lb_rows]

    cells = []
    occupied = set()
    stats = {'blocks': 0, 'spine_expanded': 0, 'row_spans': 0}

    for ri, lb_row in enumerate(lb_rows):
        for slot in range(4):
            blk = lb_row['blocks'][slot] if slot < len(lb_row.get('blocks', [])) else None
            if not blk:
                continue
            span = max(1, int(blk.get('span', 1)))
            row_span = max(1, int(blk.get('row_span', 1)))
            if row_span > 1:
                stats['row_spans'] += 1
            sb = studio_block(blk, text_styles)
            btype = sb.get('type')

            if slot == 3 or (slot + span) > 3:
                # Touches the spine column.
                if slot == 3 and btype in SPINE_TYPES:
                    # One block per revision column - the Studio model.
                    for k in range(rev_count):
                        col = 3 + k
                        if (ri, col) in occupied:
                            continue
                        cells.append({'r': ri, 'c': col, 'row_span': row_span,
                                      'col_span': 1, 'block': copy.deepcopy(sb)})
                        occupied.add((ri, col))
                    stats['spine_expanded'] += 1
                    stats['blocks'] += 1
                    continue
                # Non-spine content reaching the spine: span the whole band.
                start = slot
                col_span = (3 - slot) + rev_count
            else:
                start = slot
                col_span = span

            if (ri, start) in occupied:
                continue
            cells.append({'r': ri, 'c': start, 'row_span': row_span,
                          'col_span': col_span, 'block': sb})
            for rr in range(ri, ri + row_span):
                for cc in range(start, start + col_span):
                    occupied.add((rr, cc))
            stats['blocks'] += 1

    size_name = 'Custom'
    for nm, (w, h) in PAGE_SIZES.items():
        if (abs(w - page_w) < 1 and abs(h - page_h) < 1) or \
           (abs(h - page_w) < 1 and abs(w - page_h) < 1):
            size_name = nm
            break

    return {
        'name': name,
        'n_rows': n_rows, 'n_cols': n_cols,
        'row_heights': row_heights, 'col_widths': col_widths,
        'row_sections': row_sections,
        'page_size_name': size_name, 'orientation': orientation,
        'margin_mm': margin,
        'page_w_mm': page_w, 'page_h_mm': page_h,
        'logo_path': lb.get('logo_path', ''),
        'cells': cells,
    }, stats


def main():
    if not os.path.isdir(OUT_DIR):
        os.makedirs(OUT_DIR)
    names = sorted(f for f in os.listdir(SRC_DIR) if f.lower().endswith('.json'))
    for fn in names:
        src = os.path.join(SRC_DIR, fn)
        out, stats = convert(src)
        dest = os.path.join(OUT_DIR, fn)
        with open(dest, 'w') as f:
            json.dump(out, f, indent=2)
        print('%-26s -> %2d rows x %2d cols, %3d cells  '
              '(spine blocks expanded: %d, row-spans: %d)'
              % (fn, out['n_rows'], out['n_cols'], len(out['cells']),
                 stats['spine_expanded'], stats['row_spans']))
        reps = [i + 1 for i, s in enumerate(out['row_sections']) if s != 'body']
        if reps:
            print('%-26s    repeating rows: %s' % ('', reps))


if __name__ == '__main__':
    main()
