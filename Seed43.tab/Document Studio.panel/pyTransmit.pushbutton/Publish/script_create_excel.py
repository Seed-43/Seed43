# -*- coding: utf-8 -*-
"""
pyTransmit — Excel Creator  (Layout JSON-driven)
=================================================
Driven by PYTRANSMIT_PAYLOAD injected from script.py via exec().
Layout structure is read from Layout/Layouts/Excel.json.

The layout JSON defines:
  - Row structure (block types, spans, row_span, section, merge_down)
  - Text styles (font, size_mm, bold, italic, colour, bg_color)
  - Column percentages (col_pct)
  - Logo path
  - Revision count

The payload provides runtime data:
  - Recipients (labels, attn, copies)
  - Reason / Method legend text
  - Meta row values per revision (from Revit)
  - Branding (bg colours etc.)
"""

_p = globals().get('PYTRANSMIT_PAYLOAD', {})

from pyrevit import revit, script, DB, forms
from Autodesk.Revit.DB import FilteredElementCollector
import re, os, json, copy

output = script.get_output()
doc    = revit.doc

# ── Constants ────────────────────────────────────────────────────────────────

MONTHS = ["January","February","March","April","May","June",
          "July","August","September","October","November","December"]

MM_TO_PT = 2.835          # 1 mm = 2.835 Excel points (row height)
MM_TO_CH = 0.175          # 1 mm ≈ 0.175 Excel character-width units (approx, Arial)

# ── Load layout JSON ─────────────────────────────────────────────────────────

def _load_layout():
    # 1. Explicit path injected by script.py from the Export Settings assignment
    _explicit = _p.get('layout_json_path')
    if _explicit and os.path.isfile(_explicit):
        try:
            with open(_explicit, 'r') as f:
                data = json.load(f)
            output.print_md("Loaded layout: **{}** ({} rows)".format(_explicit, len(data.get('rows', []))))
            return data
        except Exception as e:
            output.print_md("Warning: could not load assigned layout: {}".format(e))

    # 2. Fallback: search by convention
    script_dir = (_p.get('script_dir') or
                  os.path.dirname(os.path.abspath(__file__)))
    candidates = [
        os.path.join(script_dir, 'Layout', 'Layouts', 'Excel.json'),
        os.path.join(script_dir, 'Layouts', 'Excel.json'),
        os.path.join(script_dir, 'Excel.json'),
    ]
    for path in candidates:
        output.print_md("Checking layout path: `{}`  exists={}".format(
            path, os.path.isfile(path)))
        if os.path.isfile(path):
            try:
                with open(path, 'r') as f:
                    data = json.load(f)
                output.print_md("Loaded layout: **{}** ({} rows)".format(
                    path, len(data.get('rows', []))))
                return data
            except Exception as e:
                output.print_md("Warning: could not load layout JSON: {}".format(e))
    output.print_md("**Warning: Excel.json not found — workbook will be empty.**")
    return None

LAYOUT = _load_layout()

if LAYOUT:
    ROWS        = LAYOUT.get('rows', [])
    COL_PCT     = LAYOUT.get('col_pct', [20, 20, 20])
    REV_COUNT   = LAYOUT.get('rev_count', 10)
    LOGO_PATH   = LAYOUT.get('logo_path', '')
    TEXT_STYLES = LAYOUT.get('text_styles', {})
    PAGE_W_MM   = LAYOUT.get('page_w_mm', 210)
    PAGE_H_MM   = LAYOUT.get('page_h_mm', 297)
    ORIENTATION = LAYOUT.get('orientation', 'portrait')
else:
    ROWS = []; COL_PCT = [20,20,20]; REV_COUNT = 10; LOGO_PATH = ''; TEXT_STYLES = {}
    PAGE_W_MM = 210; PAGE_H_MM = 297; ORIENTATION = 'portrait'

MAX_REVS = REV_COUNT

# ── Runtime data from payload ─────────────────────────────────────────────────

def _hex(h, default='#000000'):
    try:
        h = (h or '').strip().lstrip('#')
        if len(h) == 3: h = h[0]*2 + h[1]*2 + h[2]*2
        return '#{}'.format(h) if len(h) == 6 else default
    except: return default

_RECIP_DATA = (_p.get('recipients') or [
    {'label': 'Architect/Designer', 'attn': '', 'copies': ''},
    {'label': 'Owner/Developer',    'attn': '', 'copies': ''},
    {'label': 'Contractor',         'attn': '', 'copies': ''},
    {'label': 'Local Authority',    'attn': '', 'copies': ''},
])

REASON_LEGEND = _p.get('reason_legend') or (
    "P  Prelim\nT  Tender\nC  Construction\nI  Information\n"
    "Ab As Built\nCT Concept\nCA Consent Application"
)
METHOD_LEGEND = _p.get('method_legend') or (
    "M  Mail\nH  Hand\nC  Courier\nP  Pickup\n"
    "E  Email\nSF Sharefile\nCD CD/Flash Drive"
)

def _to_1line(legend):
    return '  |  '.join(line.strip() for line in legend.splitlines() if line.strip())

REASON_LEGEND_1LINE = _p.get('reason_legend_1line') or _to_1line(REASON_LEGEND)
METHOD_LEGEND_1LINE = _p.get('method_legend_1line') or _to_1line(METHOD_LEGEND)

# ── Helpers ───────────────────────────────────────────────────────────────────

def natural_sort_key(s):
    parts = re.split(r'(\d+)', str(s))
    return [int(p) if p.isdigit() else p.lower() for p in parts]

def rev_letter(seq):
    n = seq - 1
    if n < 0: return '?'
    result = ''
    while True:
        result = chr(65 + (n % 26)) + result
        n = n // 26 - 1
        if n < 0: break
    return result

def parse_date(raw, fmt='dd/MM/yyyy'):
    if not raw: return ""
    m = re.search(r'(\d{1,2})\D(\d{1,2})\D(\d{2,4})', str(raw).strip())
    if not m: return str(raw)
    d, mo, yr = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if yr < 100: yr += 2000
    if not (1 <= mo <= 12): return str(raw)
    sep = '/' if '/' in (fmt or '') else ('.' if '.' in (fmt or '') else '-')
    if fmt and 'MMMM' in fmt:
        if 'dddd' in fmt:
            import datetime
            try:
                dt = datetime.date(yr, mo, d)
                return dt.strftime('%A, %d %B %Y')
            except: pass
        return "{} {} {}".format(d, MONTHS[mo-1], yr)
    short_yr = fmt and 'yy' in fmt and 'yyyy' not in fmt
    yr_str = str(yr)[2:] if short_yr else str(yr)
    return "{:02d}{}{:02d}{}{}".format(d, sep, mo, sep, yr_str)

def parse_date_long(raw):
    if not raw: return ""
    m = re.search(r'(\d{1,2})\D(\d{1,2})\D(\d{2,4})', str(raw).strip())
    if not m: return str(raw)
    d, mo, yr = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if yr < 100: yr += 2000
    return "{} {} {}".format(d, MONTHS[mo-1], yr) if 1 <= mo <= 12 else str(raw)

def get_param(el, name):
    try:
        p = el.LookupParameter(name)
        if p and p.HasValue:
            return (p.AsString() or p.AsValueString() or "").strip()
    except: pass
    return ""

def _parse_copies(issued_to_str, recipient_label, recipient_index=0):
    import re as _re
    _block = ''
    for _part in issued_to_str.split(' | '):
        _part = _part.strip()
        if _part.startswith('DL:') or _part.startswith('CL:'):
            _block = _part[3:].strip(); break
    if not _block and ' | ' in issued_to_str:
        _block = issued_to_str.split(' | ', 1)[1].strip()
    m = _re.search(r'{}[A-Za-z]\.\[([^\]]*)\](\d*)'.format(recipient_index + 1), _block)
    if m: return m.group(2)
    first = recipient_label[0].upper() if recipient_label else ''
    if first:
        m2 = _re.search(r'(?:^| )' + first + r'\.\[([^\]]*)\](\d*)', _block)
        if m2: return m2.group(2)
    if recipient_label:
        m3 = _re.search(r'\[' + _re.escape(recipient_label[:6]) + r'[^\]]*\](\d+)', issued_to_str)
        if m3: return m3.group(1)
    return ''

# ── Collect Revit data ────────────────────────────────────────────────────────

proj_info   = doc.ProjectInformation
org_name    = get_param(proj_info, "Organization Name") or ""
client_name = get_param(proj_info, "Client Name")       or ""
proj_number = get_param(proj_info, "Project Number")    or ""
proj_name   = get_param(proj_info, "Project Name")      or doc.Title or ""

safe_base = re.sub(r'[\\/*?:"<>|]', '_',
                   "Document_Transmittal_{}_{}".format(proj_number, proj_name))[:60]

all_revs    = list(FilteredElementCollector(doc).OfClass(DB.Revision).ToElements())
issued_revs = sorted([r for r in all_revs if r.Issued], key=lambda r: r.SequenceNumber)
n_revs      = min(len(issued_revs), MAX_REVS)

_meta_vals = {}
if _p.get('meta_rows'):
    _LABEL_TO_KEY = {'issued by':'initials','initials':'initials',
                     'reason for issue':'reason','method of issue':'method',
                     'document format':'doc_format','paper size':'paper_size'}
    for _lbl, _val in _p['meta_rows']:
        _k = _LABEL_TO_KEY.get(_lbl.lower().strip())
        if _k: _meta_vals[_k] = _val

rev_meta = []
for i in range(MAX_REVS):
    if i < n_revs:
        rev = issued_revs[i]
        raw_date = get_param(rev, "Revision Date") or str(rev.RevisionDate)
        ito = rev.IssuedTo or ""
        def _f(key, raw=ito):
            m2 = re.search(r'\[' + key + r':([^\]]*)\]', raw)
            return m2.group(1).strip() if m2 else ""
        rev_meta.append({
            'date':       raw_date,
            'initials':   _meta_vals.get('initials') or (rev.IssuedBy or '').strip() or _f('I'),
            'reason':     _meta_vals.get('reason')   or _f('R'),
            'method':     _meta_vals.get('method')   or _f('M'),
            'doc_format': _meta_vals.get('doc_format') or _f('F'),
            'paper_size': _meta_vals.get('paper_size') or _f('P'),
            'letter':     rev_letter(rev.SequenceNumber),
        })
    else:
        rev_meta.append({k: "" for k in
            ['date','initials','reason','method','doc_format','paper_size','letter']})

issued_date_long = parse_date_long(rev_meta[n_revs-1]['date']) if n_revs > 0 else ""

tx_sheets = sorted(
    [s for s in FilteredElementCollector(doc).OfClass(DB.ViewSheet).ToElements()
     if any(r.Id in set(s.GetAllRevisionIds()) for r in issued_revs)],
    key=lambda s: natural_sort_key(s.SheetNumber)
)

GROUP_PARAMS = _p.get('group_params') or []
GROUP_LABEL  = _p.get('group_label', True)

# Build grouped sheet render list: [('group', label), ('sheet', sheet_obj), ...]
# Also pre-compute _GROUP_BOUNDARY_SI for border closing.
def _gpv2(s, pn):
    try:
        p = s.LookupParameter(pn)
        return (p.AsString() or p.AsValueString() or '').strip() if p and p.HasValue else ''
    except: return ''

def _get_sheet_gl(sheet):
    return u'|'.join([_gpv2(sheet, pn) for pn in GROUP_PARAMS])

_GROUP_BOUNDARY_SI = set()
if GROUP_PARAMS:
    from collections import OrderedDict as _OD_grp
    _grp_map = _OD_grp()
    for _gs in tx_sheets:
        _grp_map.setdefault(_get_sheet_gl(_gs), []).append(_gs)
    # Build flat render list with group headers
    _sheet_render = []
    _first_grp_xl = True
    for _gl, _gs_list in _grp_map.items():
        if _gl and (GROUP_LABEL or not _first_grp_xl):
            _sheet_render.append(('group', _gl))
        for _gs in _gs_list:
            _sheet_render.append(('sheet', _gs))
            _first_grp_xl = False
    # Build boundary set: last sheet index before a group change
    _prev_gl = None
    for _gsi, _gs in enumerate(tx_sheets):
        _gl2 = _get_sheet_gl(_gs)
        if _prev_gl is not None and _gl2 != _prev_gl:
            _GROUP_BOUNDARY_SI.add(_gsi - 1)
        _prev_gl = _gl2
else:
    _sheet_render = [('sheet', s) for s in tx_sheets]

output.print_md("# pyTransmit — Excel Generator (JSON-driven)")
output.print_md("Layout: **{}** | Revisions: **{}** | Sheets: **{}**".format(
    LAYOUT.get('template','?') if LAYOUT else 'default', n_revs, len(tx_sheets)))

# ── Save path ─────────────────────────────────────────────────────────────────

save_path = None
default_fn = "{}.xlsx".format(safe_base)
_pdf_temp_path = _p.get('_pdf_temp_xlsx_path')
if _pdf_temp_path:
    save_path = _pdf_temp_path
else:
    try:
        from System.Windows.Forms import SaveFileDialog, DialogResult
        dlg = SaveFileDialog()
        dlg.Title = "Save Transmittal — Excel"
        dlg.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
        dlg.FileName = default_fn
        dlg.InitialDirectory = os.path.expanduser("~\\Desktop")
        if dlg.ShowDialog() == DialogResult.OK:
            save_path = dlg.FileName
        else:
            output.print_md("Save cancelled."); script.exit()
    except Exception:
        save_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), default_fn)

if not save_path:
    script.exit()

# ── Workbook ──────────────────────────────────────────────────────────────────

import xlsxwriter
workbook = xlsxwriter.Workbook(save_path)
ws       = workbook.add_worksheet("Transmittal")

# ── Column layout ─────────────────────────────────────────────────────────────
# Layout has 4 logical columns (A B C D).
# Col D expands into MAX_REVS revision columns in Excel.
# col_pct = [A%, B%, C%], D% = 100 - sum
# Convert mm to Excel character widths (1 ch ≈ 7px for Arial 10pt)

_d_pct = max(5, 100 - sum(COL_PCT[:3]))
_pcts  = list(COL_PCT[:3]) + [_d_pct]

# Each layout col gets an Excel column index
# A → 0, B → 1, C → 2, D → 3..3+MAX_REVS-1
COL_A     = 0
COL_B     = 1
COL_C     = 2
REV_START = 3
LAST_COL  = REV_START + MAX_REVS - 1

# Set column widths from col_pct (PAGE_W_MM minus ~20mm margins = usable width)
_usable_mm = PAGE_W_MM - 20.0
for _ci, _pct in enumerate(_pcts[:3]):
    _w_mm = _usable_mm * _pct / 100.0
    ws.set_column(_ci, _ci, _w_mm * MM_TO_CH * 5.5)  # tuned multiplier for Arial

# Rev columns: D percentage split across MAX_REVS columns
# Use same MM_TO_CH * 5.5 multiplier for consistency
_d_w_mm  = _usable_mm * _d_pct / 100.0
_rev_w   = max(2.5, (_d_w_mm / MAX_REVS) * MM_TO_CH * 5.5)
ws.set_column(REV_START, LAST_COL, _rev_w)

# ── Layout analysis pass ──────────────────────────────────────────────────────
# Walk the layout rows to:
#   1. Determine which rows are merged groups
#   2. Count expandable rows (recipients, sheets) to assign Excel row numbers
#   3. Identify repeat_header and footer rows

def _get_groups(rows):
    """Return list of (start, end) group spans. Standalone rows are (i, i)."""
    groups = []
    i = 0
    while i < len(rows):
        start = i
        while i < len(rows) - 1 and rows[i].get('merge_down', False):
            i += 1
        groups.append((start, i))
        i += 1
    return groups

GROUPS = _get_groups(ROWS)

# Data-expanding block types: each needs one Excel row per data item
DATA_EXPAND = {
    'sent_to':    len(_RECIP_DATA),
    'attn_to':    len(_RECIP_DATA),
    'spine_copies': len(_RECIP_DATA),
    'sheet_number': len(_sheet_render),
    'sheet_desc':   len(_sheet_render),
    'spine_rev':    len(_sheet_render),
}
DRAWING_GROUP_EXPAND = max(3, len(_sheet_render))

# Build a map: layout_row_index → excel_row_start
# and layout_row_index → n_excel_rows
_row_excel_start = {}
_row_excel_count = {}
_cur_excel_row = 0

for grp_start, grp_end in GROUPS:
    for ri in range(grp_start, grp_end + 1):
        row = ROWS[ri]
        sec = row.get('section', 'body')
        if sec == 'footer':
            _row_excel_start[ri] = _cur_excel_row  # assign but don't advance
            _row_excel_count[ri] = 0
            continue
        n_excel = 1
        for b in row.get('blocks', []):
            if not b: continue
            t = b.get('type', '')
            if t in DATA_EXPAND:
                n_excel = max(n_excel, DATA_EXPAND[t])
            elif t == 'drawing_group':
                n_excel = max(n_excel, DRAWING_GROUP_EXPAND)
        _row_excel_start[ri] = _cur_excel_row
        _row_excel_count[ri] = n_excel
        _cur_excel_row += n_excel

_total_excel_rows = _cur_excel_row

# Track repeat_header and footer rows
_repeat_header_rows = []
_footer_blocks = []

for ri, row in enumerate(ROWS):
    sec = row.get('section', 'body')
    if sec == 'repeat_header':
        er_start = _row_excel_start[ri]
        er_count = _row_excel_count[ri]
        _repeat_header_rows.append((er_start, er_start + er_count - 1))
    elif sec == 'footer':
        for b in row.get('blocks', []):
            if b and b.get('type') in ('page_count', 'issue_date'):
                _footer_blocks.append(b)

# ── Format factory ────────────────────────────────────────────────────────────

BORDER_COLOR = '#808080'

def _style_to_pt(style_name):
    """Convert size_mm from text style to Excel font points."""
    ts = TEXT_STYLES.get(style_name, {})
    size_mm = ts.get('size_mm', 2.3)
    return max(6, int(round(size_mm * MM_TO_PT * 2.0)))  # × 2 to match visual size

def _style_font(style_name):
    ts = TEXT_STYLES.get(style_name, {})
    return ts.get('font', 'Arial')

def _style_bold(style_name):
    ts = TEXT_STYLES.get(style_name, {})
    return bool(ts.get('bold', False))

def _style_color(style_name):
    ts = TEXT_STYLES.get(style_name, {})
    return _hex(ts.get('color', '#000000'))

def _just_to_xl(just):
    m = {'left': 'left', 'center': 'center', 'right': 'right'}
    return m.get(just or 'left', 'left')

def _valign_to_xl(v_just):
    m = {'top': 'top', 'middle': 'vcenter', 'bottom': 'bottom'}
    return m.get(v_just or 'middle', 'vcenter')

_fmt_cache = {}

def fmt(bold=False, pt=8, font='Arial', fg='#000000', bg=None,
        align='left', valign='vcenter', wrap=False, rotation=0,
        top=0, bottom=0, left=0, right=0):
    key = (bold, pt, font, fg, bg, align, valign, wrap, rotation,
           top, bottom, left, right)
    if key in _fmt_cache:
        return _fmt_cache[key]
    d = {
        'font_name': font, 'font_size': pt, 'bold': bold,
        'font_color': fg, 'align': align, 'valign': valign,
        'text_wrap': wrap, 'rotation': rotation,
        'top': top, 'bottom': bottom, 'left': left, 'right': right,
    }
    if top    > 0: d['top_color']    = BORDER_COLOR
    if bottom > 0: d['bottom_color'] = BORDER_COLOR
    if left   > 0: d['left_color']   = BORDER_COLOR
    if right  > 0: d['right_color']  = BORDER_COLOR
    if bg:
        try: d['bg_color'] = bg
        except: pass
    f = workbook.add_format(d)
    _fmt_cache[key] = f
    return f

def block_fmt(block, override=None, rotation=0, data_row=False):
    """Build an xlsxwriter format from a block dict + optional overrides."""
    if not block: block = {}
    style_name = block.get('text_style', 'Data')
    b = override or {}
    _bold    = b.get('bold',     _style_bold(style_name))
    _pt      = b.get('pt',       _style_to_pt(style_name))
    _font    = b.get('font',     _style_font(style_name))
    _fg      = b.get('fg',       _style_color(style_name))
    _bg      = b.get('bg') or (_hex(block.get('bg_color'), None) if block.get('bg_color') else None)
    _align   = b.get('align',    _just_to_xl(block.get('just', 'left')))
    _valign  = b.get('valign',   _valign_to_xl(block.get('v_just', 'middle')))
    _wrap    = b.get('wrap',     False)
    _rot     = b.get('rotation', rotation)
    bords    = block.get('borders', {})
    _top     = b.get('top',     1 if bords.get('t', False) else 0)
    _bot     = b.get('bottom',  1 if bords.get('b', False) else 0)
    _lft     = b.get('left',    1 if bords.get('l', False) else 0)
    _rgt     = b.get('right',   1 if bords.get('r', False) else 0)
    return fmt(bold=_bold, pt=_pt, font=_font, fg=_fg, bg=_bg,
               align=_align, valign=_valign, wrap=_wrap, rotation=_rot,
               top=_top, bottom=_bot, left=_lft, right=_rgt)


# ── Data row format helper ────────────────────────────────────────────────────

def data_row_fmt(block, row_idx, n_rows, bot_b, extra=None):
    """Format for a data row: applies alt_rows bg and h-border between rows."""
    _is_last = (row_idx == n_rows - 1)
    _hlines  = block.get('data_borders', {}).get('h', True)
    _bot_i   = bot_b if _is_last else (1 if _hlines else 0)
    _alt     = block.get('alt_rows', False) and (row_idx % 2 == 1)
    _alt_col = _hex(block.get('alt_color', '#F5F7FA'), None) if _alt else None
    ov = {'top': 0, 'bottom': _bot_i}
    if _alt_col:
        ov['bg'] = _alt_col
    if extra:
        ov.update(extra)
    return block_fmt(block, override=ov)

# Rich-string sub-formats

# ── Border resolution helpers ─────────────────────────────────────────────────

def _resolve_borders(block, above_block, below_block):
    """Resolve shared borders: on beats off.
    A shared edge is ON if either side wants it — but only if the current
    block participates in bordering (has at least one border set).
    Returns (top, bottom) for THIS block after resolving with neighbours."""
    bords  = block.get('borders', {}) if block else {}
    top    = 1 if bords.get('t', False) else 0
    bottom = 1 if bords.get('b', False) else 0
    _any_border = any(bords.get(k, False) for k in ('t','b','l','r'))

    if _any_border:
        # On beats off: if row above wants bottom, this row's top is on
        if above_block:
            ab = above_block.get('borders', {})
            if ab.get('b', False):
                top = 1
        # On beats off: if row below wants top, this row's bottom is on
        if below_block:
            bb = below_block.get('borders', {})
            if bb.get('t', False):
                bottom = 1

    return top, bottom

# ── Write helpers ─────────────────────────────────────────────────────────────

def w(r, c, val, f):
    ws.write(r, c, val, f)

def mrg(r1, c1, r2, c2, val, f):
    if r1 == r2 and c1 == c2:
        ws.write(r1, c1, val, f)
    else:
        ws.merge_range(r1, c1, r2, c2, val, f)

def write_legend(r1, c, r2, title, legend_text, block):
    """Write a merged legend cell (reason/method list)."""
    _pt   = _style_to_pt(block.get('text_style', 'Data'))
    _font = _style_font(block.get('text_style', 'Data'))
    _fg   = _style_color(block.get('text_style', 'Data'))
    _bg   = _hex(block.get('bg_color'), None) if block.get('bg_color') else None

    _bold_fmt  = workbook.add_format({'font_name': _font, 'font_size': _pt,
                                      'bold': True,  'font_color': _fg})
    _plain_fmt = workbook.add_format({'font_name': _font, 'font_size': _pt,
                                      'bold': False, 'font_color': _fg})
    _cell_fmt  = block_fmt(block, override={'bold': False, 'valign': 'top', 'wrap': True,
                                             'top': 0, 'bottom': 1, 'left': 0, 'right': 0})
    if _bg: _cell_fmt.set_bg_color(_bg)
    if r1 == r2:
        ws.write(r1, c, "{}\n{}".format(title, legend_text), _cell_fmt)
    else:
        ws.merge_range(r1, c, r2, c, "", _cell_fmt)
        ws.write_rich_string(r1, c,
            _bold_fmt,  title + "\n",
            _plain_fmt, legend_text,
            _cell_fmt)

# ── Row height calculator ─────────────────────────────────────────────────────

def _row_height_pt(block, n_data=1):
    """Calculate Excel row height in points based on block's text style."""
    if not block:
        return 15.0
    t = block.get('type', '')
    style = block.get('text_style', 'Data')
    pt = _style_to_pt(style)  # already doubled for visual match
    line_h = pt * 1.5
    if t == 'spine_dates':
        return max(65.0, pt * 5)
    if t in ('reason_list', 'method_list'):
        # list_style='row' → single inline line; 'list' → multiline
        if block.get('list_style', 'list') == 'row':
            return max(15.0, line_h)
        _legend = REASON_LEGEND if t == 'reason_list' else METHOD_LEGEND
        _lines = len([l for l in _legend.splitlines() if l.strip()])
        return max(15.0, pt * 1.4 * _lines + 4)
    if t == 'logo':
        return max(40.0, pt * 3)
    if t == 'blank':
        pct = block.get('height_pct', 50) or 50
        return max(6.0, line_h * pct / 100.0 * 2)
    if t in ('text',) and block.get('content', ''):
        return max(line_h, pt * 1.8)
    return max(15.0, line_h)

# ── Logical → Excel column mapper ─────────────────────────────────────────────

def layout_col_to_excel(ci, span=1):
    """Convert a layout column index (0-3) and span to (excel_col_start, excel_col_end).
    Column 3 (D) maps to REV_START..LAST_COL."""
    if ci == 3:
        return REV_START, LAST_COL
    if ci + span - 1 >= 3:
        # spans into col D
        return ci, LAST_COL
    return ci, ci + span - 1

# ── Cell styles via layout_Style_Algoruthum ─────────────────────────────────
_style_algo_path = None
for _style_dir in [
    _p.get('script_dir'),
    os.path.dirname(os.path.abspath(__file__)) if '__file__' in dir() else None,
    os.path.dirname(os.path.abspath(script.get_script_path())) if hasattr(script, 'get_script_path') else None,
]:
    if _style_dir:
        _sc = os.path.join(_style_dir, 'layout_Style_Algoruthum.py')
        if os.path.isfile(_sc):
            _style_algo_path = _sc
            break

output.print_md("layout_Style_Algoruthum path: `{}`".format(_style_algo_path or 'NOT FOUND'))

if _style_algo_path:
    import imp as _imp2
    _style_algo = _imp2.load_source('layout_Style_Algoruthum', _style_algo_path)
    # Use grid-based borders if hlines/vlines present, else fall back to block.borders
    _hlines = {int(k): v for k,v in LAYOUT.get('hlines', {}).items()}
    _vlines = {int(k): v for k,v in LAYOUT.get('vlines', {}).items()}
    if _hlines or _vlines:
        _cell_styles = _style_algo.compute_cell_styles_from_grid(
            rows        = ROWS,
            groups      = GROUPS,
            hlines      = _hlines,
            vlines      = _vlines,
            text_styles = TEXT_STYLES,
            max_rev_col = REV_START,
            n_revs      = MAX_REVS,
        )
    else:
        _cell_styles = _style_algo.compute_cell_styles(
            rows        = ROWS,
            groups      = GROUPS,
            text_styles = TEXT_STYLES,
            max_rev_col = REV_START,
            n_revs      = MAX_REVS,
        )
else:
    output.print_md("**Warning: layout_Style_Algoruthum.py not found — borders/colours computed inline.**")
    _cell_styles = {}

def _get_cell_style(ri, ci):
    """Get pre-computed cell style, falling back to empty dict."""
    return _cell_styles.get((ri, ci), {})

# ── Row heights via layout_Algoruthum ───────────────────────────────────────
# Try multiple path strategies for IronPython/pyRevit compatibility
_algo_path = None
for _algo_dir in [
    _p.get('script_dir'),
    os.path.dirname(os.path.abspath(__file__)) if '__file__' in dir() else None,
    os.path.dirname(os.path.abspath(script.get_script_path())) if hasattr(script, 'get_script_path') else None,
]:
    if _algo_dir:
        _candidate = os.path.join(_algo_dir, 'layout_Algoruthum.py')
        if os.path.isfile(_candidate):
            _algo_path = _candidate
            break

output.print_md("layout_Algoruthum path: `{}`".format(_algo_path or 'NOT FOUND'))

if _algo_path:
    import imp as _imp
    _algo = _imp.load_source('layout_Algoruthum', _algo_path)
    _row_heights = _algo.compute_row_heights(
        rows            = ROWS,
        groups          = GROUPS,
        excel_starts    = _row_excel_start,
        excel_counts    = _row_excel_count,
        block_height_fn = _row_height_pt,
        single_line_h   = 18.0,
        min_h           = 15.0,
    )
else:
    output.print_md("**Warning: layout_Algoruthum.py not found — using simple row heights.**")
    _row_heights = {}
    for _ri2, _row2 in enumerate(ROWS):
        _er_s2 = _row_excel_start.get(_ri2, 0)
        for _b2 in _row2.get('blocks', []):
            if _b2:
                _h2 = _row_height_pt(_b2)
                _row_heights[_er_s2] = max(_row_heights.get(_er_s2, 0), _h2)

for _er, _h in _row_heights.items():
    ws.set_row(_er, _h)

# ── Main writing loop ─────────────────────────────────────────────────────────

# _prev_row_blocks no longer needed — borders resolved by layout_Style_Algoruthum
_prev_row_blocks = [None, None, None, None]  # kept for write_legend compatibility

output.print_md("Writing {} layout rows → {} Excel rows...".format(len(ROWS), _total_excel_rows))

# Pre-compute occupied cells (col-spanned) per layout row
def _occupied(row):
    occ = set()
    for i, b in enumerate(row.get('blocks', [])):
        if b and b.get('span', 1) > 1:
            for s in range(1, b['span']):
                if i + s < 4: occ.add(i + s)
    return occ

# ── Write each layout row ────────────────────────────────────────────────────

for ri, row in enumerate(ROWS):
    er_start = _row_excel_start[ri]
    er_count = _row_excel_count[ri]
    sec = row.get('section', 'body')
    occ = _occupied(row)
    blocks = row.get('blocks', [])

    # Find this row's group info
    grp_start, grp_end = ri, ri
    for gs, ge in GROUPS:
        if gs <= ri <= ge:
            grp_start, grp_end = gs, ge
            break

    # ── Write each block in this layout row ───────────────────────
    if sec == 'footer':
        continue  # footer blocks handled via ws.set_footer()
    ci = 0
    while ci < 4:
        if ci in occ:
            ci += 1
            continue

        b = blocks[ci] if ci < len(blocks) else None

        if not b or not b.get('enabled', True):
            ci += 1
            continue

        t = b.get('type', '')
        span = min(b.get('span', 1), 4 - ci)
        row_span = b.get('row_span', 1)

        ec_start, ec_end = layout_col_to_excel(ci, span)

        # Get pre-computed cell style from Style Algorithm
        _cs = _get_cell_style(ri, ci)
        top_b  = 1 if _cs.get('t', False) else 0
        bot_b  = 1 if _cs.get('b', False) else 0
        left_b = 1 if _cs.get('l', False) else 0
        rgt_b  = 1 if _cs.get('r', False) else 0
        _cell_bg = _cs.get('bg') or (_hex(b.get('bg_color'), None) if b.get('bg_color') else None)
        # ── Determine vertical Excel row span for this block ──────
        # merge_down: if this col is present here but None in next layout rows
        # of the same group, span those Excel rows too (like rowspan in HTML)
        er_end = er_start + er_count - 1
        if row_span > 1 and grp_end >= ri + row_span - 1:
            er_end = _row_excel_start[ri + row_span - 1] + _row_excel_count[ri + row_span - 1] - 1
        elif row.get('merge_down', False):
            # Walk forward through group: span rows where this ci is None/absent
            _span_end_ri = ri
            _nri = ri + 1
            while _nri <= grp_end:
                _nrow_blocks = ROWS[_nri].get('blocks', [])
                _nblock = _nrow_blocks[ci] if ci < len(_nrow_blocks) else None
                if _nblock is None:
                    _span_end_ri = _nri
                    _nri += 1
                else:
                    break
            if _span_end_ri > ri:
                er_end = _row_excel_start[_span_end_ri] + _row_excel_count[_span_end_ri] - 1

        # ══ TEXT block ════════════════════════════════════════════
        if t == 'text':
            content = b.get('content', '')
            _ov = {'top': top_b, 'bottom': bot_b, 'left': left_b, 'right': rgt_b}
            if _cell_bg: _ov['bg'] = _cell_bg
            _fmt = block_fmt(b, override=_ov)
            mrg(er_start, ec_start, er_end, ec_end, content, _fmt)

        # ══ LOGO block ════════════════════════════════════════════
        elif t == 'logo':
            _fmt = block_fmt(b, override={'top': 0, 'bottom': 0, 'left': 0, 'right': 0})
            mrg(er_start, ec_start, er_end, ec_end, "", _fmt)
            # Resolve logo path
            _sd = _p.get('script_dir') or os.path.dirname(os.path.abspath(__file__))
            _lp = LOGO_PATH if LOGO_PATH and os.path.isfile(LOGO_PATH) else None
            if not _lp:
                for _ext in ('png','jpg','jpeg','PNG','JPG','JPEG'):
                    _c = os.path.join(_sd, 'Layout', 'logo.{}'.format(_ext))
                    if os.path.isfile(_c): _lp = _c; break
            if not _lp:
                for _ext in ('png','jpg','jpeg','PNG','JPG','JPEG'):
                    _c = os.path.join(_sd, 'logo.{}'.format(_ext))
                    if os.path.isfile(_c): _lp = _c; break
            if _lp:
                try:
                    import struct as _struct
                    # Read image pixel dimensions
                    _w_px, _h_px, _dpi_x, _dpi_y = None, None, 96.0, 96.0
                    try:
                        with open(_lp, 'rb') as _fh: _hdr = _fh.read(4)
                        if _hdr[:4] == b'\x89PNG':
                            with open(_lp, 'rb') as _fh: _d = _fh.read()
                            _w_px = _struct.unpack('>I', _d[16:20])[0]
                            _h_px = _struct.unpack('>I', _d[20:24])[0]
                            _idx = _d.find(b'pHYs')
                            if _idx != -1:
                                _ppux = _struct.unpack('>I', _d[_idx+4:_idx+8])[0]
                                _ppuy = _struct.unpack('>I', _d[_idx+8:_idx+12])[0]
                                _unit = _struct.unpack('B', _d[_idx+12:_idx+13])[0]
                                if _unit == 1 and _ppux > 0: _dpi_x = _ppux * 0.0254; _dpi_y = _ppuy * 0.0254
                        elif _hdr[:2] == b'\xff\xd8':
                            with open(_lp, 'rb') as _fh: _d = _fh.read()
                            _i = 2
                            while _i < len(_d) - 4:
                                if _d[_i] != 0xFF: break
                                _mk = _d[_i+1]; _sl = _struct.unpack('>H', _d[_i+2:_i+4])[0]
                                if _mk == 0xE0 and _d[_i+4:_i+9] == b'JFIF\x00':
                                    _u = _d[_i+9]; _dx = _struct.unpack('>H', _d[_i+10:_i+12])[0]
                                    if _u == 1 and _dx > 0: _dpi_x = float(_dx); _dpi_y = _dpi_x
                                if _mk in (0xC0, 0xC1, 0xC2):
                                    _h_px = _struct.unpack('>H', _d[_i+5:_i+7])[0]
                                    _w_px = _struct.unpack('>H', _d[_i+7:_i+9])[0]; break
                                _i += 2 + _sl
                    except Exception: pass

                    _just = b.get('just', 'left')

                    if _w_px and _h_px:
                        # Scale logo to fit row height
                        _row_h_px = _row_height_pt(b) / 72.0 * 96.0  # pt → screen px
                        _phys_h_cm = (_h_px / _dpi_y) * 2.54
                        _phys_w_cm = (_w_px / _dpi_x) * 2.54
                        _y_scale = (_row_h_px / 96.0 * 2.54) / _phys_h_cm
                        _x_scale = _y_scale
                        # Logo screen width in px (96dpi screen)
                        _logo_screen_w = int(_phys_w_cm * _x_scale * 96.0 / 2.54)

                        if _just == 'right':
                            # xlsxwriter col width in chars → pixels: 1 char ≈ 7px for Arial 10pt
                            # _rev_w is the width we set for rev cols in character units
                            _col_char_w = _rev_w  # set earlier as ws.set_column width
                            _col_px = int(_col_char_w * 7.0)  # approx chars → px
                            # Walk back from LAST_COL to find enough columns to cover logo width
                            _covered = 0; _anchor_col = LAST_COL
                            for _ci_logo in range(LAST_COL, ec_start - 1, -1):
                                _covered += _col_px
                                _anchor_col = _ci_logo
                                if _covered >= _logo_screen_w: break
                            # x_offset = remaining space to push logo right
                            _x_off = max(0, _covered - _logo_screen_w)
                            ws.insert_image(er_start, _anchor_col, _lp, {
                                'x_scale': _x_scale, 'y_scale': _y_scale,
                                'x_offset': _x_off, 'y_offset': 2,
                                'object_position': 1})
                        elif _just == 'center':
                            _anchor_col = (ec_start + ec_end) // 2
                            ws.insert_image(er_start, _anchor_col, _lp, {
                                'x_scale': _x_scale, 'y_scale': _y_scale,
                                'x_offset': 0, 'y_offset': 2,
                                'object_position': 1})
                        else:
                            ws.insert_image(er_start, ec_start, _lp, {
                                'x_scale': _x_scale, 'y_scale': _y_scale,
                                'x_offset': 0, 'y_offset': 2,
                                'object_position': 1})
                    else:
                        # No dimension info — use fixed scale
                        _anchor = LAST_COL if _just == 'right' else ec_start
                        ws.insert_image(er_start, _anchor, _lp,
                            {'x_scale': 0.5, 'y_scale': 0.5, 'object_position': 1})
                except Exception as _e:
                    output.print_md("Logo insert error: {}".format(_e))

        # ══ PROJECT INFO blocks ═══════════════════════════════════
        elif t == 'proj_org':
            _lbl_fmt = block_fmt(b, override={'bold': True, 'top': top_b, 'bottom': bot_b, 'left': left_b, 'right': rgt_b})
            _val_fmt = block_fmt(b, override={'bold': False, 'top': top_b, 'bottom': bot_b, 'left': left_b, 'right': rgt_b})
            if ec_start == ec_end:
                w(er_start, ec_start, "Organisation:  " + org_name, _lbl_fmt)
            else:
                w(er_start, ec_start, "Organisation:", _lbl_fmt)
                mrg(er_start, ec_start + 1, er_end, ec_end, org_name, _val_fmt)

        elif t == 'proj_client':
            _lbl_fmt = block_fmt(b, override={'bold': True, 'top': top_b, 'bottom': bot_b, 'left': left_b, 'right': rgt_b})
            _val_fmt = block_fmt(b, override={'bold': False, 'top': top_b, 'bottom': bot_b, 'left': left_b, 'right': rgt_b})
            if ec_start == ec_end:
                w(er_start, ec_start, "Client:  " + client_name, _lbl_fmt)
            else:
                w(er_start, ec_start, "Client:", _lbl_fmt)
                mrg(er_start, ec_start + 1, er_end, ec_end, client_name, _val_fmt)

        elif t == 'proj_number':
            _lbl_fmt = block_fmt(b, override={'bold': True, 'top': top_b, 'bottom': bot_b, 'left': left_b, 'right': rgt_b})
            _val_fmt = block_fmt(b, override={'bold': False, 'top': top_b, 'bottom': bot_b, 'left': left_b, 'right': rgt_b})
            if ec_start == ec_end:
                w(er_start, ec_start, "Project No:  " + proj_number, _lbl_fmt)
            else:
                w(er_start, ec_start, "Project No:", _lbl_fmt)
                mrg(er_start, ec_start + 1, er_end, ec_end, proj_number, _val_fmt)

        elif t == 'proj_name':
            _lbl_fmt = block_fmt(b, override={'bold': True, 'top': top_b, 'bottom': bot_b, 'left': left_b, 'right': rgt_b})
            _val_fmt = block_fmt(b, override={'bold': False, 'top': top_b, 'bottom': bot_b, 'left': left_b, 'right': rgt_b})
            if ec_start == ec_end:
                w(er_start, ec_start, "Project:  " + proj_name, _lbl_fmt)
            else:
                w(er_start, ec_start, "Project:", _lbl_fmt)
                mrg(er_start, ec_start + 1, er_end, ec_end, proj_name, _val_fmt)

        # ══ DISTRIBUTION blocks ═══════════════════════════════════
        elif t == 'sent_to':
            for _i, _rec in enumerate(_RECIP_DATA):
                _er = er_start + _i
                _rfmt = data_row_fmt(b, _i, len(_RECIP_DATA), bot_b, extra={"top": top_b if _i == 0 else 0, "left": left_b, "right": rgt_b})
                w(_er, ec_start, _rec.get('label', ''), _rfmt)

        elif t == 'attn_to':
            for _i, _rec in enumerate(_RECIP_DATA):
                _er = er_start + _i
                _rfmt = data_row_fmt(b, _i, len(_RECIP_DATA), bot_b, extra={"top": top_b if _i == 0 else 0, "left": left_b, "right": rgt_b})
                mrg(_er, ec_start, _er, ec_end, _rec.get('attn', ''), _rfmt)

        elif t == 'spine_copies':
            _vlines_sc = b.get('data_borders', {}).get('v', True)
            for _i, _rec in enumerate(_RECIP_DATA):
                _er = er_start + _i
                for _rci in range(MAX_REVS):
                    _copies = ''
                    if _rci < n_revs:
                        _copies = _parse_copies(
                            issued_revs[_rci].IssuedTo or '',
                            _rec.get('label', ''), _i)
                    _right_b = (1 if _vlines_sc else 0) if _rci < MAX_REVS - 1 else (1 if b.get('borders',{}).get('r', False) else 0)
                    _left_b  = (1 if b.get('borders',{}).get('l', False) else 0) if _rci == 0 else 0
                    _rfmt = data_row_fmt(b, _i, len(_RECIP_DATA), bot_b,
                                        extra={'align': 'center', 'right': _right_b, 'left': _left_b})
                    w(_er, REV_START + _rci, _copies, _rfmt)

        # ══ REVISION LEGEND blocks ════════════════════════════════
        elif t == 'reason_list':
            _brd = b.get('borders', {})
            _is_row = b.get('list_style', 'list') == 'row' or ec_end >= REV_START
            _rl_text = REASON_LEGEND_1LINE if _is_row else REASON_LEGEND
            _wrap    = False if _is_row else True
            _ec_e    = LAST_COL if _is_row else ec_end
            _rl_fmt  = block_fmt(b, override={'top': top_b, 'bottom': bot_b,
                                              'left': 1 if _brd.get('l', False) else 0,
                                              'right': 1 if _brd.get('r', False) else 0,
                                              'wrap': _wrap, 'valign': 'top'})
            mrg(er_start, ec_start, er_end, _ec_e, _rl_text, _rl_fmt)

        elif t == 'method_list':
            _brd = b.get('borders', {})
            _is_row = b.get('list_style', 'list') == 'row' or ec_end >= REV_START
            _ml_text = METHOD_LEGEND_1LINE if _is_row else METHOD_LEGEND
            _wrap    = False if _is_row else True
            _ec_e    = LAST_COL if _is_row else ec_end
            _ml_fmt  = block_fmt(b, override={'top': top_b, 'bottom': bot_b,
                                              'left': 1 if _brd.get('l', False) else 0,
                                              'right': 1 if _brd.get('r', False) else 0,
                                              'wrap': _wrap, 'valign': 'top'})
            mrg(er_start, ec_start, er_end, _ec_e, _ml_text, _ml_fmt)

        # ══ SPINE HEADER blocks (rotated per revision) ════════════
        elif t == 'spine_dates':
            _vlines = b.get('data_borders', {}).get('v', True)
            _hlines = b.get('data_borders', {}).get('h', True)
            _rot = 90 if b.get('rotation', 0) == 270 else b.get('rotation', 0)
            for _rci in range(MAX_REVS):
                _date_raw = rev_meta[_rci]['date'] if _rci < n_revs else ''
                _dfmt_str = b.get('date_format', 'dd/MM/yyyy')
                _date_val = parse_date(_date_raw, _dfmt_str)
                _right_b  = (1 if _vlines else 0) if _rci < MAX_REVS - 1 else (1 if b.get('borders',{}).get('r', False) else 0)
                _left_b   = (1 if b.get('borders',{}).get('l', False) else 0) if _rci == 0 else 0
                _rng_fmt  = block_fmt(b, override={
                    'top': top_b, 'bottom': bot_b, 'left': _left_b, 'right': _right_b,
                    'align': 'center', 'rotation': _rot})
                mrg(er_start, REV_START + _rci, er_end, REV_START + _rci, _date_val, _rng_fmt)

        elif t in ('spine_initials','spine_reason','spine_method',
                   'spine_doc_type','spine_print_size'):
            _key_map = {
                'spine_initials':   'initials',
                'spine_reason':     'reason',
                'spine_method':     'method',
                'spine_doc_type':   'doc_format',
                'spine_print_size': 'paper_size',
            }
            _key  = _key_map.get(t, 'initials')
            _rot  = 90 if b.get('rotation', 0) == 270 else b.get('rotation', 0)
            _vlines = b.get('data_borders', {}).get('v', True)
            for _rci in range(MAX_REVS):
                _val = rev_meta[_rci][_key] if _rci < n_revs else ''
                _right_b = (1 if _vlines else 0) if _rci < MAX_REVS - 1 else (1 if b.get('borders',{}).get('r', False) else 0)
                _left_b  = (1 if b.get('borders',{}).get('l', False) else 0) if _rci == 0 else 0
                _rng_fmt = block_fmt(b, override={
                    'top': top_b, 'bottom': bot_b,
                    'left': _left_b, 'right': _right_b,
                    'align': 'center', 'rotation': _rot})
                w(er_start, REV_START + _rci, _val, _rng_fmt)

        # ══ DOCUMENTATION blocks ══════════════════════════════════
        elif t == 'sheet_number':
            for _si, (_kind, _item) in enumerate(_sheet_render):
                _er = er_start + _si
                if _kind == 'group':
                    _gfmt = block_fmt(b, override={'bold': True, 'top': 0 if _si > 0 else top_b,
                                                    'bottom': 0, 'left': 0, 'right': 0})
                    mrg(_er, ec_start, _er, LAST_COL,
                        _item if GROUP_LABEL else '', _gfmt)
                else:
                    _n_eff = len(_sheet_render) if _si not in _GROUP_BOUNDARY_SI else _si + 1
                    _sfmt = data_row_fmt(b, _si, _n_eff, bot_b,
                                        extra={"top": top_b if _si == 0 else 0,
                                               "left": left_b, "right": rgt_b})
                    w(_er, ec_start, str(_item.SheetNumber), _sfmt)

        elif t == 'sheet_desc':
            for _si, (_kind, _item) in enumerate(_sheet_render):
                _er = er_start + _si
                if _kind == 'group':
                    pass  # group label written by sheet_number block
                else:
                    _n_eff = len(_sheet_render) if _si not in _GROUP_BOUNDARY_SI else _si + 1
                    _sfmt = data_row_fmt(b, _si, _n_eff, bot_b,
                                        extra={"top": top_b if _si == 0 else 0,
                                               "left": left_b, "right": rgt_b})
                    mrg(_er, ec_start, _er, ec_end, str(_item.Name), _sfmt)

        elif t == 'spine_rev':
            _vlines = b.get('data_borders', {}).get('v', True)
            for _si, (_kind, _item) in enumerate(_sheet_render):
                _er = er_start + _si
                if _kind == 'group':
                    pass  # group row handled by sheet_number block
                else:
                    _n_eff = len(_sheet_render) if _si not in _GROUP_BOUNDARY_SI else _si + 1
                    _sheet_rev_ids = set(_item.GetAllRevisionIds())
                    for _rci in range(MAX_REVS):
                        _val = (rev_meta[_rci]['letter']
                                if (_rci < n_revs and issued_revs[_rci].Id in _sheet_rev_ids)
                                else "")
                        _right_b = (1 if _vlines else 0) if _rci < MAX_REVS - 1 else (1 if b.get('borders',{}).get('r', False) else 0)
                        _left_b  = (1 if b.get('borders',{}).get('l', False) else 0) if _rci == 0 else 0
                        _rfmt = data_row_fmt(b, _si, _n_eff, bot_b,
                                            extra={'top': top_b if _si == 0 else 0,
                                                   'left': _left_b, 'right': _right_b, 'align': 'center'})
                        w(_er, REV_START + _rci, _val, _rfmt)

        # ══ BLANK row ═════════════════════════════════════════════
        elif t == 'blank':
            _rfmt = block_fmt(b, override={'top': 0, 'bottom': 0, 'left': 0, 'right': 0})
            mrg(er_start, ec_start, er_end, ec_end, "", _rfmt)

        # ══ PAGE COUNT / ISSUE DATE → footer only (handled below) ═
        elif t in ('page_count', 'issue_date'):
            pass  # handled in footer setup

        # ══ DOC TYPE / PRINT SIZE (metadata display blocks) ═══════
        elif t == 'doc_type':
            _val = _p.get('doc_type_val', _meta_vals.get('doc_format', ''))
            _rfmt = block_fmt(b, override={'top': top_b, 'bottom': bot_b, 'left': left_b, 'right': rgt_b})
            mrg(er_start, ec_start, er_end, ec_end, _val, _rfmt)

        elif t == 'print_size':
            _val = _p.get('print_size_val', _meta_vals.get('paper_size', ''))
            _rfmt = block_fmt(b, override={'top': top_b, 'bottom': bot_b, 'left': left_b, 'right': rgt_b})
            mrg(er_start, ec_start, er_end, ec_end, _val, _rfmt)

        # ══ DRAWING GROUP ═════════════════════════════════════════
        elif t == 'drawing_group':
            if GROUP_PARAMS:
                from collections import OrderedDict as _OD
                _grps = _OD()
                for _s in tx_sheets:
                    def _gpv(s, pn):
                        try:
                            p = s.LookupParameter(pn)
                            return (p.AsString() or p.AsValueString() or '').strip() if p and p.HasValue else ''
                        except: return ''
                    _gl = u' \u2014 '.join([_gpv(_s, pn) for pn in GROUP_PARAMS if _gpv(_s, pn)])
                    _grps.setdefault(_gl, []).append(_s)
                _g_row = er_start
                _g_hdr_fmt = block_fmt(b, override={'bold': True, 'top': 0, 'bottom': 0})
                for _gl, _gs in _grps.items():
                    if _gl:
                        if GROUP_LABEL:
                            mrg(_g_row, ec_start, _g_row, ec_end, _gl, _g_hdr_fmt)
                        _g_row += 1
                    _g_row += len(_gs)

        # Update previous block tracker
        _prev_row_blocks[ci] = b

        ci += 1 + (span - 1)  # skip spanned columns

    # Update prev blocks for non-occupied columns
    for _ci2 in range(4):
        if _ci2 not in occ and (_ci2 >= len(blocks) or not blocks[_ci2]):
            _prev_row_blocks[_ci2] = None

# ── Logo (robust re-attempt using read_image_info) ────────────────────────────
# Already inserted inline above in logo block; this pass is a no-op

# ── Print setup ──────────────────────────────────────────────────────────────

# Paper size: xlsxwriter codes — 9=A4, 8=A3, 1=Letter
_paper_map = {
    (210, 297): 9,   # A4 portrait
    (297, 210): 9,   # A4 landscape
    (297, 420): 8,   # A3 portrait
    (420, 297): 8,   # A3 landscape
}
_paper_code = _paper_map.get((int(PAGE_W_MM), int(PAGE_H_MM)), 9)
ws.set_paper(_paper_code)
if ORIENTATION == 'landscape':
    ws.set_landscape()
else:
    ws.set_portrait()
ws.set_margins(left=0.39, right=0.39, top=0.39, bottom=0.75)
ws.center_horizontally()
ws.fit_to_pages(1, 0)   # fit to 1 page wide, natural height
ws.hide_gridlines(1)

# Set print area to exactly the used columns
import xlsxwriter.utility as _xlu
_last_col_letter = _xlu.xl_col_to_name(LAST_COL)
ws.print_area('A1:{}{}'.format(_last_col_letter, _total_excel_rows))

# ── Repeat rows (from repeat_header sections) ────────────────────────────────

if _repeat_header_rows:
    first_rr = min(r[0] for r in _repeat_header_rows)
    last_rr  = max(r[1] for r in _repeat_header_rows)
    ws.repeat_rows(first_rr, last_rr)

# ── Footer (from footer section blocks) ─────────────────────────────────────

_footer_left  = ''
_footer_right = ''

for _fb in _footer_blocks:
    _t = _fb.get('type', '')
    _just = _fb.get('just', 'left')
    _prefix = _fb.get('prefix', '')
    _suffix = _fb.get('suffix', '')

    if _t == 'issue_date':
        _dfmt = _fb.get('date_format', 'dd/MM/yyyy')
        _date_str = parse_date(rev_meta[n_revs-1]['date'] if n_revs > 0 else '', _dfmt)
        _parts = [p for p in [_prefix, _date_str, _suffix] if p]
        _val = ' '.join(_parts)
        if _just == 'right':
            _footer_right = _val
        else:
            _footer_left = _val

    elif _t == 'page_count':
        _pfmt = _fb.get('page_format', 'Page X of Y')
        if _pfmt == 'Page X':
            _page_code = 'Page &P'
        elif _pfmt == 'Page X of Y':
            _page_code = 'Page &P of &N'
        elif _pfmt == 'X of Y':
            _page_code = '&P of &N'
        elif _pfmt == 'X / Y':
            _page_code = '&P / &N'
        else:
            _page_code = '&N'
        _parts = [p for p in [_prefix, _page_code, _suffix] if p]
        _val = ' '.join(_parts)
        if _just == 'right':
            _footer_right = _val
        else:
            _footer_left = _val

_footer_str = ''
if _footer_left:
    _footer_str += '&L{}'.format(_footer_left)
if _footer_right:
    _footer_str += '&R{}'.format(_footer_right)

if _footer_str:
    ws.set_footer(_footer_str, {'font_size': 8})

ws.set_header('')

# ── Save ──────────────────────────────────────────────────────────────────────

import time

MAX_RETRIES = 5
saved = False

for attempt in range(MAX_RETRIES):
    try:
        workbook.close()
        saved = True
        break
    except Exception as e:
        err_msg = str(e)
        if attempt < MAX_RETRIES - 1:
            retry = forms.alert(
                "Could not save — the file may already be open in Excel.\n\n"
                "Please close the file then click Retry, or Cancel to abort.\n\n"
                "File: {}\n\nError: {}".format(save_path, err_msg),
                title="File Locked — Please Close Excel",
                yes=True, no=True)
            if not retry:
                output.print_md("Save aborted.")
                script.exit()
            time.sleep(1)
        else:
            output.print_md("## Save failed after {} attempts".format(MAX_RETRIES))
            forms.alert("Could not save after {} attempts:\n{}".format(MAX_RETRIES, err_msg))
            script.exit()

if saved:
    output.print_md("\n## Saved  `{}`".format(save_path))
    output.print_md("Sheets: {} | Revisions: {} | Layout rows: {}".format(
        len(tx_sheets), n_revs, len(ROWS)))
    if not _pdf_temp_path:
        if forms.alert("Excel saved!\n\nOpen the file?", yes=True, no=True):
            os.startfile(save_path)
