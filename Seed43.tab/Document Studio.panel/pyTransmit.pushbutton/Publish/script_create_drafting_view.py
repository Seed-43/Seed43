# -*- coding: utf-8 -*-
__title__     = "pyTransmit - Drafting View"
__author__    = "Nagel Consultants"
__doc__       = """
VERSION 250507
_____________________________________________________________________
Description:
Creates or updates a Revit Drafting View containing the full
transmittal document layout. The layout is driven by the
Revit Drafting View.json file, so column widths, row heights,
text sizes, and which sections appear are all controlled from
the Layout Builder without editing this script.

_____________________________________________________________________
How-to:
This script is run automatically by pyTransmit when you click
Publish. You do not need to run it directly. Once complete,
find the view in the Project Browser under Drafting Views.

_____________________________________________________________________
Notes:
The layout JSON controls which blocks appear and in what order.
Supported block types: logo, text, sent_to, attn_to,
spine_copies, spine_dates, spine_initials, spine_reason,
spine_method, spine_doc_type, spine_print_size, reason_list,
method_list, sheet_number, sheet_desc, spine_rev, blank.

If the view already exists it will be cleared and redrawn.
Overflow sheet rows are placed in side-by-side columns on the
same view rather than creating extra pages.

_____________________________________________________________________
Last update:
250507 - Rewrote to be fully JSON-driven. Column widths, row
heights, and text styles now come from the layout JSON instead
of hardcoded constants. Matches the pattern used by the Excel
and Schedule exporters.
_____________________________________________________________________
"""

# ── IMPORTS ──────────────────────────────────────────────────────────────────

_p = globals().get('PYTRANSMIT_PAYLOAD', {})

from pyrevit import revit, script, DB, forms
from Autodesk.Revit.DB import (
    FilteredElementCollector, XYZ, Line, TextNote, TextNoteType,
    Transaction, CurveElement, ViewFamilyType, ViewFamily, ViewDrafting,
    TextNoteOptions, HorizontalTextAlignment,
    ImageType, ImageTypeOptions, ImageInstance, ImageTypeSource,
    ImagePlacementOptions, BoxPlacement,
    FilledRegion, FilledRegionType, CurveLoop,
    GraphicsStyleType, FillPatternElement,
    Color,
)
import re, os, json, math

output = script.get_output()
doc    = revit.doc

# ── CONSTANTS ─────────────────────────────────────────────────────────────────

MM                = 1.0 / 304.8
SHORT_CURVE_TOL   = 0.002083333
INDENT            = 0.8 * MM
MONTHS            = ["January","February","March","April","May","June",
                     "July","August","September","October","November","December"]

# ── LOAD LAYOUT JSON ─────────────────────────────────────────────────────────

def _load_layout():
    _explicit = _p.get('layout_json_path')
    if _explicit and os.path.isfile(_explicit):
        try:
            with open(_explicit, 'r') as f:
                return json.load(f)
        except Exception as e:
            output.print_md("Warning: could not load assigned layout: {}".format(e))
    _script_dir = _p.get('script_dir') or os.path.dirname(os.path.abspath(__file__))
    _candidates = [
        os.path.join(_script_dir, 'Layout', 'Layouts', 'Revit Drafting View.json'),
        os.path.join(_script_dir, 'Layouts', 'Revit Drafting View.json'),
        os.path.join(_script_dir, 'Revit Drafting View.json'),
    ]
    for path in _candidates:
        if os.path.isfile(path):
            try:
                with open(path, 'r') as f:
                    return json.load(f)
            except Exception as e:
                output.print_md("Warning: could not load layout JSON: {}".format(e))
    output.print_md("**Warning: Revit Drafting View.json not found, using defaults.**")
    return None

LAYOUT = _load_layout()

if LAYOUT:
    ROWS        = LAYOUT.get('rows', [])
    COL_PCT     = LAYOUT.get('col_pct', [22, 28, 20])
    MAX_REVS    = int(LAYOUT.get('rev_count', 10))
    TEXT_STYLES = LAYOUT.get('text_styles', {})
    PAGE_W_MM   = float(LAYOUT.get('page_w_mm', 210))
    PAGE_H_MM   = float(LAYOUT.get('page_h_mm', 297))
    LOGO_PATH   = LAYOUT.get('logo_path', '')
else:
    ROWS = []; COL_PCT = [22, 28, 20]; MAX_REVS = 10; TEXT_STYLES = {}
    PAGE_W_MM = 210.0; PAGE_H_MM = 297.0; LOGO_PATH = ''

# ── COLUMN GEOMETRY FROM JSON ─────────────────────────────────────────────────
# Columns A, B, C share the first three col_pct values.
# Column D (revision columns) gets the remainder, split across MAX_REVS.

_MARGIN_MM  = 5.0
_USABLE_MM  = PAGE_W_MM - 2 * _MARGIN_MM
_D_PCT      = max(5, 100 - sum(COL_PCT[:3]))
_PCTS       = list(COL_PCT[:3]) + [_D_PCT]

C_A   = _USABLE_MM * _PCTS[0] / 100.0 * MM
C_B   = _USABLE_MM * _PCTS[1] / 100.0 * MM
C_C   = _USABLE_MM * _PCTS[2] / 100.0 * MM
_D_W  = _USABLE_MM * _D_PCT   / 100.0 * MM
C_REV = max(5.0 * MM, _D_W / MAX_REVS)

COL_A_X = 0.0
COL_B_X = COL_A_X + C_A
COL_C_X = COL_B_X + C_B
REV_X   = [COL_C_X + C_C + i * C_REV for i in range(MAX_REVS)]
TABLE_W = REV_X[-1] + C_REV

# ── ROW HEIGHTS FROM TEXT STYLES ──────────────────────────────────────────────

def _style_mm(name, default=2.3):
    return TEXT_STYLES.get(name, {}).get('size_mm', default)

def _style_bold(name):
    return bool(TEXT_STYLES.get(name, {}).get('bold', False))

def _style_font(name):
    return TEXT_STYLES.get(name, {}).get('font', 'Arial')

_DATA_MM   = _style_mm('Data',   2.3)
_HDR_MM    = _style_mm('Header', 2.5)
_TITLE_MM  = _style_mm('Title',  4.5)

H_TITLE   = max(12.0, _TITLE_MM  * 4.0) * MM
H_HDR     = max(6.0,  _HDR_MM    * 2.5) * MM
H_COL_HDR = max(6.0,  _HDR_MM    * 2.5) * MM
H_DATA    = max(5.0,  _DATA_MM   * 2.2) * MM
H_RECIP   = max(5.0,  _DATA_MM   * 2.2) * MM
H_META    = max(5.0,  _HDR_MM    * 2.2) * MM
H_DATE    = max(18.0, _DATA_MM   * 8.0) * MM
H_SPACER  = 2.0 * MM
H_INFO    = max(5.0,  _HDR_MM    * 2.2) * MM

# ── RUNTIME DATA FROM PAYLOAD ─────────────────────────────────────────────────

def _hex_to_rgb(h, default):
    try:
        h = (h or '').strip().lstrip('#')
        if len(h) == 3: h = h[0]*2 + h[1]*2 + h[2]*2
        if len(h) == 6:
            return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except Exception:
        pass
    return default

_title_bg  = _hex_to_rgb(_p.get('title_bg_color'),  (255, 255, 255))
_title_fg  = _hex_to_rgb(_p.get('title_fg_color'),  (  0,   0,   0))
_header_bg = _hex_to_rgb(_p.get('header_bg_color'), (255, 255, 255))
_header_fg = _hex_to_rgb(_p.get('header_fg_color'), (  0,   0,   0))

if LOGO_PATH and not os.path.isfile(LOGO_PATH):
    LOGO_PATH = _p.get('logo_path', '')

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

_LABEL_TO_KEY = {
    'issued by': 'initials', 'initials': 'initials',
    'reason for issue': 'reason', 'method of issue': 'method',
    'document format': 'format', 'paper size': 'paper',
}
_ALL_META = [
    ("Issued By",        'initials'),
    ("Reason for Issue", 'reason'),
    ("Method of Issue",  'method'),
    ("Document Format",  'format'),
    ("Paper Size",       'paper'),
]
_payload_meta = _p.get('meta_rows')
if _payload_meta is None:
    _filtered_meta = list(_ALL_META)
else:
    _enabled_keys = set()
    for _lbl, _val in _payload_meta:
        _k = _LABEL_TO_KEY.get(_lbl.lower().strip())
        if _k:
            _enabled_keys.add(_k)
    _filtered_meta = [(lbl, key) for lbl, key in _ALL_META if key in _enabled_keys]

# ── HELPERS ───────────────────────────────────────────────────────────────────

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

def parse_date_slash(raw):
    if not raw: return ""
    m = re.search(r'(\d{1,2})\D(\d{1,2})\D(\d{2,4})', str(raw).strip())
    if not m: return str(raw)
    d, mo, yr = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if yr < 100: yr += 2000
    return "{:02d}/{:02d}/{}".format(d, mo, yr)

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
    except Exception:
        pass
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

# ── COLLECT REVIT DATA ────────────────────────────────────────────────────────

proj_info   = doc.ProjectInformation
org_name    = get_param(proj_info, "Organization Name") or ""
client_name = get_param(proj_info, "Client Name")       or ""
proj_number = get_param(proj_info, "Project Number")    or ""
proj_name   = get_param(proj_info, "Project Name")      or doc.Title or ""

_SHOW_FROM     = _p.get('show_from',     True)
_SHOW_CLIENT   = _p.get('show_client',   True)
_SHOW_PROJNO   = _p.get('show_projno',   True)
_SHOW_PROJNAME = _p.get('show_projname', True)

all_revs    = list(FilteredElementCollector(doc).OfClass(DB.Revision).ToElements())
issued_revs = sorted([r for r in all_revs if r.Issued], key=lambda r: r.SequenceNumber)
n_revs      = min(len(issued_revs), MAX_REVS)

_meta_vals = {}
if _p.get('meta_rows'):
    for _lbl, _val in _p['meta_rows']:
        _k = _LABEL_TO_KEY.get(_lbl.lower().strip())
        if _k:
            _meta_vals[_k] = _val

rev_meta = []
for i in range(MAX_REVS):
    if i < n_revs:
        rev      = issued_revs[i]
        raw_date = get_param(rev, "Revision Date") or str(rev.RevisionDate)
        ito      = rev.IssuedTo or ""
        def _f(key, raw=ito):
            m2 = re.search(r'\[' + key + r':([^\]]*)\]', raw)
            return m2.group(1).strip() if m2 else ""
        rev_meta.append({
            'date':     parse_date_slash(raw_date),
            'initials': _meta_vals.get('initials') or (rev.IssuedBy or '').strip() or _f('I'),
            'reason':   _meta_vals.get('reason')   or _f('R'),
            'method':   _meta_vals.get('method')   or _f('M'),
            'format':   _meta_vals.get('format')   or _f('F'),
            'paper':    _meta_vals.get('paper')    or _f('P'),
            'letter':   rev_letter(rev.SequenceNumber),
        })
    else:
        rev_meta.append({k: "" for k in
            ['date', 'initials', 'reason', 'method', 'format', 'paper', 'letter']})

tx_sheets = sorted(
    [s for s in FilteredElementCollector(doc).OfClass(DB.ViewSheet).ToElements()
     if any(r.Id in set(s.GetAllRevisionIds()) for r in issued_revs)],
    key=lambda s: natural_sort_key(s.SheetNumber)
)

output.print_md("# pyTransmit - Drafting View Generator")
output.print_md("Layout: **{}** | Revisions: **{}** | Sheets: **{}**".format(
    LAYOUT.get('template', '?') if LAYOUT else 'default', n_revs, len(tx_sheets)))

# ── PAGE SPLITTING ────────────────────────────────────────────────────────────
# Overflow sheet rows go in side-by-side columns on the same view.

# Use page height from the layout JSON, not the Setup panel.
# The Setup panel page_height_mode only controls whether to split at all.
_PAGE_H_MODE = _p.get('page_height_mode') or 'a4'
_SPLIT       = _PAGE_H_MODE != 'none'
_PAGE_H_MM   = PAGE_H_MM - 2 * _MARGIN_MM   # usable height from JSON
_PAGE_H_FT   = _PAGE_H_MM * MM

# Measure the fixed section heights from the JSON row order
# (everything that is not sheet data rows)
_fixed_h  = 0.0
_repeat_h = 0.0
for row in ROWS:
    blocks = row.get('blocks', [])
    first_block = next((b for b in blocks if b), None)
    if not first_block:
        continue
    t = first_block.get('type', '')
    # Sheet data rows are the expanding ones
    if t in ('sheet_number', 'sheet_desc', 'spine_rev'):
        continue
    # Estimate row height from text style
    style = first_block.get('text_style', 'Data')
    _smm  = _style_mm(style)
    _rh   = max(5.0, _smm * 2.5) * MM
    if t == 'spine_dates':
        _rh = H_DATE
    elif t in ('blank',):
        _rh = H_SPACER
    _fixed_h += _rh
    # Rows that repeat at the top of overflow columns
    if first_block.get('content', '') in ('Documentation List', 'Sheet', 'Description', 'Revision'):
        _repeat_h += _rh

_rows_page1  = max(1, int((_PAGE_H_FT - _fixed_h)  / H_DATA)) if _SPLIT else 999999
_rows_page_n = max(1, int((_PAGE_H_FT - _repeat_h) / H_DATA)) if _SPLIT else 999999

output.print_md("Page 1 rows: {}  |  Overflow rows: {}".format(_rows_page1, _rows_page_n))

# ── REVIT DRAW HELPERS ────────────────────────────────────────────────────────

def get_line_style(name):
    try:
        cats = doc.Settings.Categories
        lines_cat = cats.get_Item(DB.BuiltInCategory.OST_Lines)
        for sc in lines_cat.SubCategories:
            if sc.Name == name:
                return sc.GetGraphicsStyle(GraphicsStyleType.Projection)
    except Exception:
        pass
    return None

def create_line(view, x1, y1, x2, y2):
    try:
        start = XYZ(float(x1), float(y1), 0.0)
        end   = XYZ(float(x2), float(y2), 0.0)
        if (end - start).GetLength() < SHORT_CURVE_TOL:
            return None
        return doc.Create.NewDetailCurve(view, Line.CreateBound(start, end))
    except Exception:
        return None

def get_or_create_filled_region_type(name, rgb=(255, 255, 255)):
    _col = Color(rgb[0], rgb[1], rgb[2])
    solid_pat = None
    for fp in FilteredElementCollector(doc).OfClass(FillPatternElement).ToElements():
        try:
            if fp.GetFillPattern().IsSolidFill:
                solid_pat = fp; break
        except Exception:
            pass
    if not solid_pat:
        return None

    def _apply(frt):
        try:
            frt.ForegroundPatternId    = solid_pat.Id
            frt.ForegroundPatternColor = _col
            frt.BackgroundPatternColor = _col
        except Exception:
            pass
        for _ls_name in ("<Invisible lines>", "<Invisible Lines>"):
            inv_gs = get_line_style(_ls_name)
            if inv_gs:
                try: frt.LineStyleId = inv_gs.Id; return
                except Exception: pass

    for frt in FilteredElementCollector(doc).OfClass(FilledRegionType).ToElements():
        try:
            if frt.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == name:
                _apply(frt)
                return frt
        except Exception:
            pass
    all_frt = list(FilteredElementCollector(doc).OfClass(FilledRegionType).ToElements())
    if not all_frt:
        return None
    new_frt = all_frt[0].Duplicate(name)
    _apply(new_frt)
    return new_frt

def create_filled_region(view, frt, x1, y1, x2, y2):
    try:
        loop = CurveLoop()
        pts  = [
            XYZ(float(x1), float(y1), 0.0),
            XYZ(float(x2), float(y1), 0.0),
            XYZ(float(x2), float(y2), 0.0),
            XYZ(float(x1), float(y2), 0.0),
        ]
        for i in range(4):
            loop.Append(Line.CreateBound(pts[i], pts[(i + 1) % 4]))
        region = FilledRegion.Create(doc, frt.Id, view.Id, [loop])
        if region:
            valid_ids = FilledRegion.GetValidLineStyleIdsForFilledRegion(doc)
            for ls_id in valid_ids:
                el = doc.GetElement(ls_id)
                if el:
                    n = el.Name if hasattr(el, 'Name') else ""
                    if not n:
                        try: n = el.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                        except: pass
                    if n and "invisible" in n.lower():
                        region.SetLineStyleId(ls_id); break
        return region
    except Exception as e:
        output.print_md("FilledRegion error: {}".format(e))
        return None

def create_text(view, x, y, text, ttype, align=HorizontalTextAlignment.Left, rotation=0.0):
    if not text or not str(text).strip():
        return None
    try:
        opts = TextNoteOptions(ttype.Id)
        opts.HorizontalAlignment = align
        if rotation:
            opts.Rotation = rotation
        return TextNote.Create(doc, view.Id, XYZ(float(x), float(y), 0.0), str(text), opts)
    except Exception as e:
        output.print_md("Text error '{}': {}".format(str(text)[:30], e))
        return None

def create_text_w(view, x, y, text, width, ttype, align=HorizontalTextAlignment.Left):
    if not text or not str(text).strip():
        return None
    try:
        opts = TextNoteOptions(ttype.Id)
        opts.HorizontalAlignment = align
        return TextNote.Create(doc, view.Id, XYZ(float(x), float(y), 0.0),
                               float(width), str(text), opts)
    except Exception:
        return None

# ── TEXT STYLE HELPERS ────────────────────────────────────────────────────────

def get_or_create_text_style(name, font, size_mm, bold=False, italic=False):
    size_ft   = size_mm * MM
    all_types = list(FilteredElementCollector(doc).OfClass(TextNoteType).ToElements())
    existing  = None
    for tt in all_types:
        try:
            if tt.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == name:
                existing = tt; break
        except Exception:
            pass
    with Transaction(doc, "Create/Update text style: {}".format(name)) as t:
        t.Start()
        new_tt = existing if existing else all_types[0].Duplicate(name)
        one_mm = MM
        for bip, val in [
            (DB.BuiltInParameter.TEXT_SIZE,          size_ft),
            (DB.BuiltInParameter.TEXT_FONT,           font),
            (DB.BuiltInParameter.TEXT_STYLE_BOLD,     1 if bold   else 0),
            (DB.BuiltInParameter.TEXT_STYLE_ITALIC,   1 if italic else 0),
            (DB.BuiltInParameter.TEXT_BACKGROUND,     1),
            (DB.BuiltInParameter.LEADER_OFFSET_SHEET, one_mm),
            (DB.BuiltInParameter.TEXT_TAB_SIZE,       one_mm),
        ]:
            try:
                p = new_tt.get_Parameter(bip)
                if p and not p.IsReadOnly:
                    p.Set(val)
            except Exception:
                pass
        try:
            ap = new_tt.get_Parameter(DB.BuiltInParameter.LEADER_ARROWHEAD)
            if ap and not ap.IsReadOnly:
                ap.Set(DB.ElementId.InvalidElementId)
        except Exception:
            pass
        t.Commit()
    return new_tt

def _make_text_types():
    """Build one TextNoteType per unique text style defined in the JSON."""
    _types = {}
    for style_name, style_def in TEXT_STYLES.items():
        _types[style_name] = get_or_create_text_style(
            "Transmittal {}".format(style_name),
            style_def.get('font', 'Arial'),
            style_def.get('size_mm', 2.3),
            bold   = style_def.get('bold',   False),
            italic = style_def.get('italic', False),
        )
    # Fallbacks if styles were not in JSON
    for name, font, size, bold in [
        ('Data',   'Arial', _DATA_MM,  False),
        ('Header', 'Arial', _HDR_MM,   True),
        ('Title',  'Arial', _TITLE_MM, True),
    ]:
        if name not in _types:
            _types[name] = get_or_create_text_style(
                "Transmittal {}".format(name), font, size, bold=bold)
    return _types

# ── LOGO HELPER ───────────────────────────────────────────────────────────────

def _get_or_create_logo():
    _logo_file = None
    if LOGO_PATH and os.path.isfile(LOGO_PATH):
        _logo_file = LOGO_PATH
    else:
        _sdir = _p.get('script_dir') or os.path.dirname(os.path.abspath(__file__))
        for _sd in (os.path.join(_sdir, 'Settings'), _sdir):
            for _ext in ('png', 'jpg', 'jpeg', 'PNG', 'JPG', 'JPEG'):
                _cand = os.path.join(_sd, 'logo.{}'.format(_ext))
                if os.path.isfile(_cand):
                    _logo_file = _cand; break
            if _logo_file:
                break
    if not _logo_file:
        return None

    def _src_path(img):
        try:
            p = img.SourcePath
            if p: return p.strip()
        except Exception:
            pass
        return ""

    _norm = os.path.normcase(_logo_file)
    for _img in FilteredElementCollector(doc).OfClass(ImageType).ToElements():
        try:
            if os.path.normcase(_src_path(_img)) == _norm:
                return _img
        except Exception:
            pass
    try:
        _opts = ImageTypeOptions(_logo_file, False, ImageTypeSource.Link)
        _opts.Resolution = 300
        return ImageType.Create(doc, _opts)
    except Exception as _ce:
        output.print_md("Logo create failed: {}".format(_ce))
        return None

# ── DRAFTING VIEW SETUP ───────────────────────────────────────────────────────

VIEW_NAME = _p.get('_legend_temp_view_name') or "pyTransmit Document View"

drafting_view = None
for v in FilteredElementCollector(doc).OfClass(ViewDrafting).WhereElementIsNotElementType():
    try:
        if v.IsValidObject and v.Name == VIEW_NAME:
            drafting_view = v; break
    except Exception:
        pass

if not drafting_view:
    _drafting_vft = None
    for vft in FilteredElementCollector(doc).OfClass(ViewFamilyType):
        if vft.ViewFamily == ViewFamily.Drafting:
            _drafting_vft = vft; break
    if not _drafting_vft:
        forms.alert("No Drafting ViewFamilyType found.", exitscript=True)
    with Transaction(doc, "pyTransmit DV - Create view") as t:
        t.Start()
        drafting_view = ViewDrafting.Create(doc, _drafting_vft.Id)
        drafting_view.Name = VIEW_NAME
        try: drafting_view.Scale = 1
        except Exception: pass
        t.Commit()
else:
    with Transaction(doc, "pyTransmit DV - Set scale") as t:
        t.Start()
        try: drafting_view.Scale = 1
        except Exception: pass
        t.Commit()

with Transaction(doc, "pyTransmit DV - Clear view") as t:
    t.Start()
    for _cls in (CurveElement, TextNote, ImageInstance, FilledRegion):
        for _el in list(FilteredElementCollector(doc, drafting_view.Id).OfClass(_cls).ToElements()):
            try: doc.Delete(_el.Id)
            except Exception: pass
    t.Commit()

# ── BUILD TEXT TYPES AND FILLED REGION TYPES ─────────────────────────────────

TEXT_TYPES = _make_text_types()

def _tt(style_name):
    """Return TextNoteType for the given style name."""
    return TEXT_TYPES.get(style_name) or TEXT_TYPES.get('Data')

with Transaction(doc, "pyTransmit DV - Create fill types") as _t:
    _t.Start()
    FRT_TITLE = get_or_create_filled_region_type("Transmittal Title",  _title_bg)
    FRT_HDR   = get_or_create_filled_region_type("Transmittal Header", _header_bg)
    _t.Commit()

# ── JSON BORDER MAPS ──────────────────────────────────────────────────────────
# hlines[row_index] = [col0, col1, col2, col3]  , draw bottom line per column
# vlines[row_index] = [col0_left, col1_left, col2_left, rev_start, table_right]

_HLINES = {int(k): v for k, v in LAYOUT.get('hlines', {}).items()} if LAYOUT else {}
_VLINES = {int(k): v for k, v in LAYOUT.get('vlines', {}).items()} if LAYOUT else {}

# X positions that map to vlines indices 0-4
def _vline_xs():
    return [COL_A_X, COL_B_X, COL_C_X, REV_X[0], TABLE_W]

def draw_row_borders(vw, ri, y_top, y_bot, draw_inner_rev_vlines=False, blocks=None):
    """Draw borders for layout row ri using hlines and vlines from the JSON.

    Falls back to per-block border flags when hlines or vlines have no entry
    for this row. If draw_inner_rev_vlines is True, also draws dividers between
    each revision column.
    """
    _xs = _vline_xs()

    # Top horizontal line, hlines[ri] = top edge of row ri
    _htop = _HLINES.get(ri)
    if _htop is not None:
        if any(_htop):
            create_line(vw, COL_A_X, y_top, TABLE_W, y_top)
    elif blocks:
        if any(b.get('borders', {}).get('t') for b in blocks if b):
            create_line(vw, COL_A_X, y_top, TABLE_W, y_top)

    # Bottom horizontal line, hlines[ri+1] = bottom edge of row ri
    _hbot = _HLINES.get(ri + 1)
    if _hbot is not None:
        if any(_hbot):
            create_line(vw, COL_A_X, y_bot, TABLE_W, y_bot)
    elif blocks:
        if any(b.get('borders', {}).get('b') for b in blocks if b):
            create_line(vw, COL_A_X, y_bot, TABLE_W, y_bot)

    # Vertical lines
    _vrow = _VLINES.get(ri)
    if _vrow is not None:
        for _vi, _draw in enumerate(_vrow):
            if _draw and _vi < len(_xs):
                create_line(vw, _xs[_vi], y_top, _xs[_vi], y_bot)
    elif blocks:
        # Fallback: draw left/right outer borders from block flags
        _ci = 0
        for _b in blocks:
            if not _b:
                _ci += 1; continue
            _bords = _b.get('borders', {})
            _bx    = _xs[_ci] if _ci < len(_xs) else _xs[-1]
            if _bords.get('l'):
                create_line(vw, _bx, y_top, _bx, y_bot)
            if _bords.get('r'):
                _rx2 = _xs[min(_ci + _b.get('span', 1), len(_xs) - 1)]
                create_line(vw, _rx2, y_top, _rx2, y_bot)
            _ci += _b.get('span', 1)

    # Inner revision column dividers
    if draw_inner_rev_vlines:
        for _rx in REV_X[1:]:
            create_line(vw, _rx, y_top, _rx, y_bot)



def _col_x(ci):
    """Return the X position for layout column index 0-3."""
    return [COL_A_X, COL_B_X, COL_C_X, REV_X[0]][ci]

def _col_w(ci, span=1):
    """Return the width for a block starting at column ci with the given span."""
    widths = [C_A, C_B, C_C, C_REV * MAX_REVS]
    total  = 0.0
    for s in range(span):
        if ci + s < 3:
            total += widths[ci + s]
        else:
            total += C_REV * MAX_REVS
            break
    return total

def _row_height_from_block(block):
    """Estimate drawing row height from block type and text style."""
    if not block:
        return H_DATA
    t      = block.get('type', '')
    style  = block.get('text_style', 'Data')
    _smm   = _style_mm(style)
    _line_h = max(5.0, _smm * 2.2) * MM
    if t == 'spine_dates':
        return H_DATE
    if t == 'blank':
        pct = block.get('height_pct') or 50
        return max(2.0 * MM, _line_h * pct / 100.0 * 2)
    if t == 'logo':
        return H_TITLE
    if t in ('sheet_number', 'sheet_desc', 'spine_rev',
             'sent_to', 'attn_to', 'spine_copies'):
        return _line_h  # per data row, not the total block height
    if t in ('reason_list', 'method_list'):
        _legend = REASON_LEGEND if t == 'reason_list' else METHOD_LEGEND
        _lines  = len([l for l in _legend.splitlines() if l.strip()])
        return max(_line_h, _smm * 1.4 * _lines * MM + 2.0 * MM)
    return _line_h

def _ty(y_mid, size_mm):
    """Vertically centre text of size_mm at y_mid (Revit draws from top of text)."""
    return y_mid + (size_mm * MM) / 2.0

# ── ROW TRACKER ───────────────────────────────────────────────────────────────

_y = 0.0  # current Y cursor, decrements downward

def row(height):
    global _y
    y_top = _y
    y_bot = _y - height
    y_mid = (_y + y_bot) / 2.0
    _y    = y_bot
    return y_top, y_mid, y_bot

# ── BORDER DRAWING ────────────────────────────────────────────────────────────

def _draw_table_borders(vw, x_off, y_top, y_bot, row_ys,
                        left_xs, right_x, draw_h=True, draw_v=True):
    """Draw outer rect, row dividers, and column vlines for a table block.

    left_xs - X positions for vertical lines (index 0 = left outer edge).
    right_x - right outer edge.
    draw_h  - draw horizontal dividers between rows.
    draw_v  - draw vertical dividers between inner columns.
    """
    if not row_ys:
        return
    # Outer box always drawn
    create_line(vw, x_off + left_xs[0], y_top, x_off + right_x, y_top)
    create_line(vw, x_off + left_xs[0], y_bot, x_off + right_x, y_bot)
    create_line(vw, x_off + left_xs[0], y_top, x_off + left_xs[0], y_bot)
    create_line(vw, x_off + right_x,    y_top, x_off + right_x,    y_bot)
    # Row dividers
    if draw_h:
        for _rt, _rm, _rb in row_ys[:-1]:
            create_line(vw, x_off + left_xs[0], _rb, x_off + right_x, _rb)
    # Column dividers
    if draw_v:
        for _vx in left_xs[1:]:
            create_line(vw, x_off + _vx, y_top, x_off + _vx, y_bot)

# ── GROUP PARAMS ──────────────────────────────────────────────────────────────

GROUP_PARAMS = _p.get('group_params') or []
GROUP_LABEL  = _p.get('group_label', True)

def _get_group_label(sheet):
    def _gpv(s, pn):
        try:
            p = s.LookupParameter(pn)
            return (p.AsString() or p.AsValueString() or '').strip() if p and p.HasValue else ''
        except Exception:
            return ''
    parts = [_gpv(sheet, pn) for pn in GROUP_PARAMS]
    parts = [p for p in parts if p]
    return u' \u2014 '.join(parts)

# ── MAIN DRAW TRANSACTION ─────────────────────────────────────────────────────

vw = drafting_view
_doc_list_y_start = None

with Transaction(doc, "pyTransmit DV - Draw transmittal") as t:
    t.Start()

    _logo_type = _get_or_create_logo()

    # Walk each layout row and dispatch to the correct draw routine
    _doc_list_y_start = None  # Y where documentation list header begins
    _doc_list_drawn   = False
    _in_doc_list      = False  # once we hit sheet rows, track for side-by-side logic

    # ── Pre-pass: record doc-list header row indices ──────────────────────────
    # Documentation list header rows are plain text rows whose content labels
    # the sheet table (Documentation List, Sheet, Description, Revision).
    # We identify them by content so the main loop can handle them as a group.
    _DOC_HDR_CONTENTS = {'Documentation List', 'Sheet', 'Description', 'Revision'}
    _doc_list_row_indices = set()
    for _ri, _row in enumerate(ROWS):
        _first = next((b for b in _row.get('blocks', []) if b), None)
        if _first and _first.get('content', '') in _DOC_HDR_CONTENTS:
            _doc_list_row_indices.add(_ri)

    # ── Row type resolver ─────────────────────────────────────────────────────
    # Rows often have a plain text label in col A and the data block in col D.
    # Dispatch on the first non-text block type, falling back to 'text'.
    _DATA_TYPES = {
        'spine_dates', 'spine_initials', 'spine_reason', 'spine_method',
        'spine_doc_type', 'spine_print_size', 'spine_copies', 'spine_rev',
        'sent_to', 'attn_to', 'sheet_number', 'sheet_desc',
        'reason_list', 'method_list', 'logo', 'blank',
    }

    def _row_dispatch_type(blocks):
        for b in blocks:
            if b and b.get('type', 'text') in _DATA_TYPES:
                return b.get('type')
        first = next((b for b in blocks if b), None)
        return first.get('type', 'text') if first else 'text'

    # ── Draw rows ─────────────────────────────────────────────────────────────
    _x_off           = 0.0
    _sheet_ys        = []
    _sheet_hdr_y_bot = None
    _rows_in_block   = 0
    _is_first_block  = True

    def _draw_doc_list_headers(x_off):
        """Redraw doc-list heading + col header at x_off, updating _sheet_hdr_y_bot."""
        global _y, _sheet_hdr_y_bot
        for _hri in sorted(_doc_list_row_indices):
            _hrow   = ROWS[_hri]
            _hblocks = _hrow.get('blocks', [])
            _hfirst  = next((b for b in _hblocks if b), None)
            if not _hfirst:
                continue
            _hsmm = _style_mm(_hfirst.get('text_style', 'Header'))
            _hrh  = max(H_HDR, _hsmm * 2.5 * MM)
            _hy_top, _hy_mid, _hy_bot = row(_hrh)
            _hcontent = _hfirst.get('content', '')
            if _hcontent == 'Documentation List':
                create_text(vw, x_off + COL_A_X + INDENT,
                            _ty(_hy_mid, _hsmm), _hcontent,
                            _tt(_hfirst.get('text_style', 'Header')))
            else:
                if FRT_HDR:
                    create_filled_region(vw, FRT_HDR,
                        x_off + COL_A_X, _hy_top, x_off + TABLE_W, _hy_bot)
                _sheet_hdr_y_bot = _hy_bot
                _ci2 = 0
                for _b2 in _hblocks:
                    if not _b2:
                        _ci2 += 1; continue
                    _bx2   = _col_x(_ci2) if _ci2 < 3 else REV_X[0]
                    _bsmm2 = _style_mm(_b2.get('text_style', 'Header'))
                    _bjust2 = _b2.get('just', 'left')
                    _ba2 = (HorizontalTextAlignment.Right if _bjust2 == 'right' else
                            HorizontalTextAlignment.Center if _bjust2 == 'center' else
                            HorizontalTextAlignment.Left)
                    _bxd = (x_off + _bx2 + INDENT if _bjust2 != 'right'
                            else x_off + _bx2 + _col_w(_ci2, _b2.get('span', 1)) - INDENT)
                    create_text(vw, _bxd, _ty(_hy_mid, _bsmm2),
                                _b2.get('content', ''),
                                _tt(_b2.get('text_style', 'Header')), _ba2)
                    _ci2 += _b2.get('span', 1)

    for _ri, _row in enumerate(ROWS):
        _blocks  = _row.get('blocks', [])
        _first   = next((b for b in _blocks if b), None)
        if not _first or not _first.get('enabled', True):
            continue

        _dispatch = _row_dispatch_type(_blocks)
        _style    = _first.get('text_style', 'Data')
        _smm      = _style_mm(_style)
        _rh       = _row_height_from_block(_first)
        _just     = _first.get('just', 'left')
        _content  = _first.get('content', '')

        # ── Documentation list header rows ────────────────────────────────────
        if _ri in _doc_list_row_indices:
            if _doc_list_y_start is None:
                _doc_list_y_start = _y
            y_top, y_mid, y_bot = row(max(H_HDR, _smm * 2.5 * MM))
            if _content == 'Documentation List':
                _sheet_hdr_y_bot = None
                create_text(vw, _x_off + COL_A_X + INDENT, _ty(y_mid, _smm),
                            _content, _tt(_style))
            else:
                if FRT_HDR:
                    create_filled_region(vw, FRT_HDR,
                        _x_off + COL_A_X, y_top, _x_off + TABLE_W, y_bot)
                _sheet_hdr_y_bot = y_bot
                _ci = 0
                for _b in _blocks:
                    if not _b:
                        _ci += 1; continue
                    _bx   = _col_x(_ci) if _ci < 3 else REV_X[0]
                    _bsmm = _style_mm(_b.get('text_style', _style))
                    _bjust = _b.get('just', 'left')
                    _ba = (HorizontalTextAlignment.Right if _bjust == 'right' else
                           HorizontalTextAlignment.Center if _bjust == 'center' else
                           HorizontalTextAlignment.Left)
                    _bxd = (_x_off + _bx + INDENT if _bjust != 'right'
                            else _x_off + _bx + _col_w(_ci, _b.get('span', 1)) - INDENT)
                    create_text(vw, _bxd, _ty(y_mid, _bsmm),
                                _b.get('content', ''), _tt(_b.get('text_style', _style)), _ba)
                    _ci += _b.get('span', 1)
            continue

        # ── Sheet data rows ───────────────────────────────────────────────────
        if _dispatch in ('sheet_number', 'sheet_desc', 'spine_rev'):
            if GROUP_PARAMS:
                from collections import OrderedDict as _OD
                _grps = _OD()
                for s in tx_sheets:
                    _grps.setdefault(_get_group_label(s), []).append(s)
                _render_items = []
                _first_grp = True
                for gl, gs in _grps.items():
                    if gl and (GROUP_LABEL or not _first_grp):
                        _render_items.append(('group', gl))
                    for s in gs:
                        _render_items.append(('sheet', s))
                        _first_grp = False
            else:
                _render_items = [('sheet', s) for s in tx_sheets]

            _doc_list_y_top = _doc_list_y_start or _y

            for _kind, _item in _render_items:
                _limit = _rows_page1 if _is_first_block else _rows_page_n
                if _rows_in_block >= _limit and _sheet_ys:
                    # Close current column and open a new one to the right
                    if _sheet_hdr_y_bot:
                        _rev_b2 = next((b for b in _blocks if b and b.get('type') == 'spine_rev'), None)
                        _db2    = (_rev_b2 or {}).get('data_borders', {})
                        _draw_table_borders(vw, _x_off, _sheet_hdr_y_bot,
                                            _sheet_ys[-1][2], _sheet_ys,
                                            [COL_A_X, COL_B_X, REV_X[0]], TABLE_W,
                                            draw_h=_db2.get('h', True), draw_v=_db2.get('v', True))
                        if _db2.get('v', True):
                            for _rx2 in REV_X[1:]:
                                create_line(vw, _x_off + _rx2, _sheet_hdr_y_bot,
                                            _x_off + _rx2, _sheet_ys[-1][2])
                    _sheet_ys       = []
                    _rows_in_block  = 0
                    _is_first_block = False
                    _x_off += TABLE_W + 5.0 * MM
                    _saved_y = _y
                    _y = _doc_list_y_top
                    _draw_doc_list_headers(_x_off)
                    _y = _saved_y

                if _kind == 'group':
                    if _sheet_ys:
                        _rev_b2 = next((b for b in _blocks if b and b.get('type') == 'spine_rev'), None)
                        _db2    = (_rev_b2 or {}).get('data_borders', {})
                        _draw_table_borders(vw, _x_off, _sheet_hdr_y_bot,
                                            _sheet_ys[-1][2], _sheet_ys,
                                            [COL_A_X, COL_B_X, REV_X[0]], TABLE_W,
                                            draw_h=_db2.get('h', True), draw_v=_db2.get('v', True))
                        if _db2.get('v', True):
                            for _rx2 in REV_X[1:]:
                                create_line(vw, _x_off + _rx2, _sheet_hdr_y_bot,
                                            _x_off + _rx2, _sheet_ys[-1][2])
                        _sheet_ys = []
                    _gy_top, _gy_mid, _gy_bot = row(H_HDR)
                    if GROUP_LABEL:
                        create_text(vw, _x_off + COL_A_X + INDENT,
                                    _ty(_gy_mid, _HDR_MM), _item, _tt('Header'))
                    _sheet_hdr_y_bot = _gy_bot
                else:
                    sheet = _item
                    _sy_top, _sy_mid, _sy_bot = row(H_DATA)
                    _sheet_ys.append((_sy_top, _sy_mid, _sy_bot))
                    _rows_in_block += 1
                    _ty2 = _ty(_sy_mid, _DATA_MM)
                    create_text(vw, _x_off + COL_A_X + INDENT, _ty2,
                                str(sheet.SheetNumber), _tt('Data'))
                    create_text_w(vw, _x_off + COL_B_X + INDENT, _ty2,
                                  str(sheet.Name), C_B + C_C - INDENT * 2, _tt('Data'))
                    _sheet_rev_ids = set(sheet.GetAllRevisionIds())
                    for _rci in range(MAX_REVS):
                        if _rci < n_revs and issued_revs[_rci].Id in _sheet_rev_ids:
                            create_text(vw, _x_off + REV_X[_rci] + C_REV / 2.0, _ty2,
                                        rev_meta[_rci]['letter'], _tt('Data'),
                                        HorizontalTextAlignment.Center)

            if _sheet_ys and _sheet_hdr_y_bot:
                _rev_b = next((b for b in _blocks if b and b.get('type') == 'spine_rev'), None)
                _db    = (_rev_b or {}).get('data_borders', {})
                _draw_table_borders(vw, _x_off, _sheet_hdr_y_bot,
                                    _sheet_ys[-1][2], _sheet_ys,
                                    [COL_A_X, COL_B_X, REV_X[0]], TABLE_W,
                                    draw_h=_db.get('h', True), draw_v=_db.get('v', True))
                if _db.get('v', True):
                    for _rx in REV_X[1:]:
                        create_line(vw, _x_off + _rx, _sheet_hdr_y_bot,
                                    _x_off + _rx, _sheet_ys[-1][2])
            continue

        # ── spine_dates row (label col A-C, rotated dates col D) ─────────────
        if _dispatch == 'spine_dates':
            y_top, y_mid, y_bot = row(H_DATE)
            _label_b = next((b for b in _blocks if b and b.get('type') == 'text'), None)
            _date_b  = next((b for b in _blocks if b and b.get('type') == 'spine_dates'), None)
            if _label_b:
                _lsmm = _style_mm(_label_b.get('text_style', 'Header'))
                _la   = (HorizontalTextAlignment.Right
                         if _label_b.get('just', 'right') == 'right'
                         else HorizontalTextAlignment.Left)
                _lx   = (COL_C_X + C_C - INDENT if _la == HorizontalTextAlignment.Right
                         else COL_A_X + INDENT)
                create_text(vw, _lx, _ty(y_mid, _lsmm),
                            _label_b.get('content', 'Date of Issue'),
                            _tt(_label_b.get('text_style', 'Header')), _la)
            if _date_b:
                _dsmm = _style_mm(_date_b.get('text_style', 'Data'))
                for _rci in range(MAX_REVS):
                    _dv = rev_meta[_rci]['date'] if _rci < n_revs else ''
                    if _dv:
                        _x_ctr = REV_X[_rci] + C_REV / 2.0 - 2.0 * MM
                        _y_ctr = y_mid - len(_dv) * _dsmm * 0.65 * MM / 2.0
                        create_text(vw, _x_ctr, _y_ctr, _dv,
                                    _tt(_date_b.get('text_style', 'Data')),
                                    HorizontalTextAlignment.Left, math.pi / 2.0)
            _dates_db = (_date_b or {}).get('data_borders', {})
            draw_row_borders(vw, _ri, y_top, y_bot,
                             draw_inner_rev_vlines=_dates_db.get('v', True), blocks=_blocks)
            continue

        # ── spine meta rows (initials, reason, method, format, paper) ─────────
        if _dispatch in ('spine_initials', 'spine_reason', 'spine_method',
                         'spine_doc_type', 'spine_print_size'):
            _key_map = {
                'spine_initials':   'initials',
                'spine_reason':     'reason',
                'spine_method':     'method',
                'spine_doc_type':   'format',
                'spine_print_size': 'paper',
            }
            _mk     = _key_map[_dispatch]
            _data_b = next((b for b in _blocks if b and b.get('type') == _dispatch), None)
            _dsmm2  = _style_mm(_data_b.get('text_style', 'Data')) if _data_b else _DATA_MM
            y_top, y_mid, y_bot = row(H_META)
            _label_b3 = next((b for b in _blocks if b and b.get('type') == 'text'), None)
            if _label_b3:
                _lsmm3 = _style_mm(_label_b3.get('text_style', 'Header'))
                _la3   = (HorizontalTextAlignment.Right
                          if _label_b3.get('just', 'right') == 'right'
                          else HorizontalTextAlignment.Left)
                _lx3   = (COL_C_X + C_C - INDENT if _la3 == HorizontalTextAlignment.Right
                          else COL_A_X + INDENT)
                create_text(vw, _lx3, _ty(y_mid, _lsmm3),
                            _label_b3.get('content', ''),
                            _tt(_label_b3.get('text_style', 'Header')), _la3)
            for _rci in range(MAX_REVS):
                _val = rev_meta[_rci].get(_mk, '') if _rci < n_revs else ''
                if _val:
                    create_text(vw, REV_X[_rci] + C_REV / 2.0, _ty(y_mid, _dsmm2),
                                _val, _tt('Data'), HorizontalTextAlignment.Center)
            _meta_db = (_data_b or {}).get('data_borders', {})
            draw_row_borders(vw, _ri, y_top, y_bot,
                             draw_inner_rev_vlines=_meta_db.get('v', True), blocks=_blocks)
            continue

        # ── sent_to row (recipients: label, attn, copies per revision) ────────
        if _dispatch == 'sent_to':
            _dsmm    = _style_mm(_first.get('text_style', 'Data'))
            _recip_ys = []
            _dist_y_top = _y
            for _rec_i, _rec in enumerate(_RECIP_DATA):
                _ry_top, _ry_mid, _ry_bot = row(H_RECIP)
                _recip_ys.append((_ry_top, _ry_mid, _ry_bot))
                _label   = _rec.get('label', '')
                _attn    = _rec.get('attn', '')
                _rty     = _ty(_ry_mid, _dsmm)
                create_text(vw, COL_A_X + INDENT, _rty, _label, _tt('Data'))
                if _attn:
                    create_text(vw, COL_B_X + INDENT, _rty, _attn, _tt('Data'))
                for _rci in range(MAX_REVS):
                    if _rci >= n_revs: break
                    _ito = issued_revs[_rci].IssuedTo or ''
                    _copies_val = _parse_copies(_ito, _label, _rec_i)
                    if _copies_val:
                        create_text(vw, REV_X[_rci] + C_REV / 2.0, _rty,
                                    _copies_val, _tt('Data'), HorizontalTextAlignment.Center)
            _dist_y_bot  = _y
            _copy_b      = next((b for b in _blocks if b and b.get('type') == 'spine_copies'), None)
            _sent_db     = (_copy_b or _first or {}).get('data_borders', {})
            _vrow6 = _VLINES.get(_ri, [True, True, False, True, True])
            _xs6   = _vline_xs()
            for _vi6, _draw6 in enumerate(_vrow6):
                if _draw6 and _vi6 < len(_xs6):
                    create_line(vw, _xs6[_vi6], _dist_y_top, _xs6[_vi6], _dist_y_bot)
            if any(_HLINES.get(_ri, [])):
                create_line(vw, COL_A_X, _dist_y_top, TABLE_W, _dist_y_top)
            if any(_HLINES.get(_ri + 1, [])):
                create_line(vw, COL_A_X, _dist_y_bot, TABLE_W, _dist_y_bot)
            if _sent_db.get('v', True):
                for _rx in REV_X[1:]:
                    create_line(vw, _rx, _dist_y_top, _rx, _dist_y_bot)
            if _sent_db.get('h', True):
                for _ry_t, _ry_m, _ry_b in _recip_ys[:-1]:
                    create_line(vw, COL_A_X, _ry_b, TABLE_W, _ry_b)
            continue

        # ── reason_list / method_list (inline legend row) ─────────────────────
        if _dispatch in ('reason_list', 'method_list'):
            _legend   = REASON_LEGEND if _dispatch == 'reason_list' else METHOD_LEGEND
            _leg_b    = next((b for b in _blocks if b and b.get('type') == _dispatch), None)
            _leg_smm  = _style_mm(_leg_b.get('text_style', 'Data')) if _leg_b else _DATA_MM
            _lines    = [l for l in _legend.splitlines() if l.strip()]
            _leg_rh   = max(H_DATA, _leg_smm * 1.4 * len(_lines) * MM + 2.0 * MM)
            y_top, y_mid, y_bot = row(_leg_rh)
            # Label in col A (plain text block)
            _label_b4 = next((b for b in _blocks if b and b.get('type') == 'text'), None)
            if _label_b4:
                _lsmm4 = _style_mm(_label_b4.get('text_style', 'Data'))
                create_text(vw, COL_A_X + INDENT,
                            y_top - INDENT - (_lsmm4 * MM),
                            _label_b4.get('content', ''), _tt(_label_b4.get('text_style', 'Data')))
            # Legend text from col B onward
            _leg_span_start = 1  # start after col A label
            _lx_start = _col_x(_leg_span_start)
            _lwidth   = TABLE_W - _lx_start - INDENT
            create_text_w(vw, _lx_start + INDENT,
                          y_top - INDENT - (_leg_smm * MM),
                          "  |  ".join(_lines), _lwidth, _tt('Data'))
            continue

        # ── logo / title bar ──────────────────────────────────────────────────
        if _dispatch == 'logo':
            y_top, y_mid, y_bot = row(H_TITLE)
            if FRT_TITLE:
                create_filled_region(vw, FRT_TITLE, COL_A_X, y_top, TABLE_W, y_bot)
            create_text(vw, COL_A_X + INDENT, _ty(y_mid, _TITLE_MM),
                        "Transmittal Document", _tt('Title'))
            if _logo_type:
                try:
                    _iw = _logo_type.Width; _ih = _logo_type.Height
                    _sc = H_TITLE / _ih if _ih > 0 else 1.0
                    _po = ImagePlacementOptions()
                    _po.Location = XYZ(float(TABLE_W - INDENT), float(y_top), 0.0)
                    _ii = ImageInstance.Create(doc, vw, _logo_type.Id, _po)
                    if _ii:
                        _wp = _ii.LookupParameter("Width")
                        _hp = _ii.LookupParameter("Height")
                        if _wp and not _wp.IsReadOnly: _wp.Set(_iw * _sc)
                        if _hp and not _hp.IsReadOnly: _hp.Set(_ih * _sc)
                        _ii.SetLocation(XYZ(float(TABLE_W - INDENT), float(y_top), 0.0),
                                        BoxPlacement.TopRight)
                except Exception as _le:
                    output.print_md("Logo failed: {}".format(_le))
            continue

        # ── plain text / section header ───────────────────────────────────────
        if _dispatch == 'text':
            y_top, y_mid, y_bot = row(max(H_HDR, _smm * 2.5 * MM))
            _align = (HorizontalTextAlignment.Right  if _just == 'right'  else
                      HorizontalTextAlignment.Center if _just == 'center' else
                      HorizontalTextAlignment.Left)
            _bx_draw = (COL_A_X + INDENT if _just != 'right'
                        else COL_C_X + C_C - INDENT)
            create_text(vw, _bx_draw, _ty(y_mid, _smm), _content, _tt(_style), _align)
            _bords = _first.get('borders', {})
            if _bords.get('t'): create_line(vw, COL_A_X, y_top, TABLE_W, y_top)
            if _bords.get('b'): create_line(vw, COL_A_X, y_bot, TABLE_W, y_bot)
            if _bords.get('l'): create_line(vw, COL_A_X, y_top, COL_A_X, y_bot)
            if _bords.get('r'): create_line(vw, TABLE_W, y_top, TABLE_W, y_bot)
            continue

        # ── blank spacer ──────────────────────────────────────────────────────
        if _dispatch == 'blank':
            row(_rh)
            continue

    t.Commit()

output.print_md("## Done!")
output.print_md("Drafting view **'{}'** updated.".format(VIEW_NAME))
output.print_md("Open it in the Project Browser under Drafting Views.")
