# -*- coding: utf-8 -*-
"""
pyTransmit — Schedule Creator (Full Layout v3)
"""

# Must be first — reads the payload injected by pyTransmit via exec()
_p = globals().get('PYTRANSMIT_PAYLOAD', {})

from pyrevit import revit, script, DB, forms
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory,
    ViewSchedule, ScheduleFilter, ScheduleFilterType,
    SectionType, ElementId, Transaction,
)
import re
from datetime import datetime as _dt
import clr
clr.AddReference('RevitAPI')

output = script.get_output()
doc    = revit.doc

#  Config 

FIELD_ID_ASM_CODE = -1002500
MM = 1.0 / 304.8

MONTHS = ["January","February","March","April","May","June",
          "July","August","September","October","November","December"]

# ── Load layout JSON ──────────────────────────────────────────────────────────

import json as _json, os as _os

def _load_schedule_layout():
    _explicit = _p.get('layout_json_path')
    if _explicit and _os.path.isfile(_explicit):
        try:
            with open(_explicit, 'r') as _f: return _json.load(_f)
        except Exception: pass
    _script_dir = (_p.get('script_dir') or _os.path.dirname(_os.path.abspath(__file__)))
    for _candidate in [
        _os.path.join(_script_dir, 'Layout', 'Layouts', 'Revit_Schedule.json'),
        _os.path.join(_os.path.dirname(_script_dir), 'Layout', 'Layouts', 'Revit_Schedule.json'),
        _os.path.join(_script_dir, 'Revit_Schedule.json'),
    ]:
        if _os.path.isfile(_candidate):
            try:
                with open(_candidate, 'r') as _f: return _json.load(_f)
            except Exception: pass
    return None

LAYOUT = _load_schedule_layout()

if LAYOUT:
    _ROWS       = LAYOUT.get('rows', [])
    _COL_PCT    = LAYOUT.get('col_pct', [20, 35, 17])
    MAX_REVS    = int(LAYOUT.get('rev_count', 10))
    _TEXT_STYLES= LAYOUT.get('text_styles', {})
    PAGE_W_MM   = LAYOUT.get('page_w_mm', 210)
    PAGE_H_MM   = LAYOUT.get('page_h_mm', 297)
else:
    _ROWS = []; _COL_PCT = [20, 35, 17]; MAX_REVS = 10; _TEXT_STYLES = {}
    PAGE_W_MM = 210; PAGE_H_MM = 297

# ── Derive sizes from JSON ────────────────────────────────────────────────────

def _style_mm(name, default=2.3):
    return _TEXT_STYLES.get(name, {}).get('size_mm', default)

def _style_bold(name):
    return bool(_TEXT_STYLES.get(name, {}).get('bold', False))

# Page / usable area
_MARGIN_MM   = 5.0
_PAGE_MODE   = _p.get('page_height_mode') or 'a4'
_SPLIT       = _PAGE_MODE != 'none'
A4_USABLE    = float(_p.get('page_height_mm') or PAGE_H_MM - 2 * _MARGIN_MM)

# Usable width in Revit internal units (ft)
_usable_w_mm = PAGE_W_MM - 2 * _MARGIN_MM   # e.g. 200mm for A4
_d_pct       = max(5, 100 - sum(_COL_PCT[:3]))
_pcts        = list(_COL_PCT[:3]) + [_d_pct]

# Column widths from col_pct
C_A   = _usable_w_mm * _pcts[0] / 100.0 * MM
C_B   = _usable_w_mm * _pcts[1] / 100.0 * MM
C_C   = _usable_w_mm * _pcts[2] / 100.0 * MM
_d_w  = _usable_w_mm * _d_pct  / 100.0 * MM
C_REV = max(5.0 * MM, _d_w / MAX_REVS)

# Backwards-compatible aliases used throughout script
C_REASON = C_A
C_METHOD = C_B
C_LBL    = C_C

# Row heights from text styles
_DATA_MM   = _style_mm('Data',   2.3)
_HDR_MM    = _style_mm('Header', 2.5)
_TITLE_MM  = _style_mm('Title',  4.5)

ROW_TITLE  = max(10.0, _TITLE_MM  * 4.0) * MM
ROW_NORMAL = max(4.0,  _DATA_MM   * 2.5) * MM
ROW_LEGEND = 34 * MM
ROW_DATE_H = 18 * MM
ROW_SPACER =  2 * MM

SCHEDULE_NAME = "pyTransmit Schedule 01-01"  # updated below once total_pages known

def parse_date_long(raw):
    """Parse any DD?MM?YYYY or DD?MM?YY date string -> '19 December 2025'."""
    if not raw:
        return ""
    m = re.search(r'(\d{1,2})\D(\d{1,2})\D(\d{2,4})', str(raw).strip())
    if not m:
        return str(raw)
    day, month, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if year < 100:
        year += 2000
    if 1 <= month <= 12:
        return "{} {} {}".format(day, MONTHS[month - 1], year)
    return str(raw)

def parse_date_short(raw):
    """Parse any date string -> 'DD/MM/YYYY' for the Issued footer.
    Uses \\D+ (one-or-more non-digits) so it correctly handles the \\r\\n
    separators produced by format_revit_date as well as plain / . - separators.
    """
    if not raw:
        return ""
    m = re.search(r'(\d{1,2})\D+(\d{1,2})\D+(\d{2,4})', str(raw).strip())
    if not m:
        return str(raw).replace('\r', '').replace('\n', '/')
    day, month, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if year < 100:
        year += 2000
    return "{:02d}/{:02d}/{}".format(day, month, year)

# Column/row sizes now derived from JSON above

#  Payload from pyTransmit (falls back to defaults if run standalone) 

RECIPIENTS = (
    [r['label'] for r in _p['recipients']]
    if _p.get('recipients')
    else ["Architect/Designer", "Owner/Developer", "Contractor", "Local Authority"]
)
N_REC = max(1, len(RECIPIENTS))   # actual recipient row count (at least 1)

# Whether to show reason/method legend section — from meta_rows payload
_SHOW_REASON = any(lbl.lower() in ('reason for issue', 'reason')
                   for lbl, _ in (_p.get('meta_rows') or []))
_SHOW_METHOD = any(lbl.lower() in ('method of issue', 'method')
                   for lbl, _ in (_p.get('meta_rows') or []))
_SHOW_LEGEND = _SHOW_REASON or _SHOW_METHOD  # show legend block only if needed

#  Determine which meta rows to show — must be done here so RI_ indices are correct 
_LABEL_TO_KEY = {
    'issued by':        'initials',
    'initials':         'initials',
    'reason for issue': 'reason',
    'reason':           'reason',
    'method of issue':  'method',
    'method':           'method',
    'document format':  'doc_format',
    'format':           'doc_format',
    'paper size':       'paper_size',
    'print size':       'paper_size',
    'document type':    'doc_format',   # alias
}
_ALL_META = [
    ("Issued By",       'initials'),
    ("Reason for Issue", 'reason'),
    ("Method of Issue",  'method'),
    ("Document Format",  'doc_format'),
    ("Paper Size",       'paper_size'),
]
_payload_meta = _p.get('meta_rows')
if _payload_meta is None:
    _filtered = list(_ALL_META)  # standalone: show all
else:
    # meta_rows was explicitly provided (even if empty = all fields off)
    _enabled_keys = set()
    for _lbl, _val in _payload_meta:
        _k = _LABEL_TO_KEY.get(_lbl.lower().strip())
        if _k:
            _enabled_keys.add(_k)
    _filtered = [(lbl, key) for lbl, key in _ALL_META if key in _enabled_keys]

# Reason/Method legend text — from payload (built from live OptionsSettings data)
# If payload absent (standalone run), read directly from the Options JSON files.
def _read_coded_json(path, one_line=False):
    """Read [{'code','separator','description'}] → formatted legend text."""
    try:
        import json as _jj, os as _oo
        with open(path, 'r') as _f:
            rows = _jj.load(_f)
        lines = []
        for r in rows:
            c = str(r.get('code', '') or '').strip()
            s = str(r.get('separator', '-') or '-').strip()
            d = str(r.get('description', '') or '').strip()
            if c or d:
                parts = [p for p in [c, s, d] if p]
                lines.append(' '.join(parts))
        sep = '  |  ' if one_line else '\n'
        return sep.join(lines)
    except Exception:
        return ''

_script_dir   = _p.get('_script_dir') or ''
_settings_dir = _p.get('_settings_dir') or ''
_reason_json  = (_settings_dir and __import__('os').path.join(_settings_dir, 'reason.json')) or ''
_method_json  = (_settings_dir and __import__('os').path.join(_settings_dir, 'method.json')) or ''

REASON_LEGEND = (
    _p.get('reason_legend')
    or (_reason_json and _read_coded_json(_reason_json))
    or "A - Issued for Approval\nB - Issued for Construction\nC - Issued for Coordination\nD - Issued for Information\nE - Issued for Review"
)
METHOD_LEGEND = (
    _p.get('method_legend')
    or (_method_json and _read_coded_json(_method_json))
    or "E - Email\nP - Post\nH - By Hand\nU - Uploaded to Portal"
)
# One-line versions for schedule cells (Revit can't word-wrap multiline in narrow cells)
REASON_LEGEND_1LINE = (
    (_reason_json and _read_coded_json(_reason_json, one_line=True))
    or '  |  '.join(l for l in REASON_LEGEND.split('\n') if l.strip())
)
METHOD_LEGEND_1LINE = (
    (_method_json and _read_coded_json(_method_json, one_line=True))
    or '  |  '.join(l for l in METHOD_LEGEND.split('\n') if l.strip())
)

#  Styling from payload (with sensible defaults) 
def _hex_to_rgb(h):
    """Convert '#RRGGBB', 'RRGGBB', '#RGB', or 'RGB' to (r,g,b) tuple, or return None on failure."""
    try:
        h = h.strip().lstrip('#')
        if len(h) == 3:
            # Expand shorthand: #F07 -> #FF0077
            h = h[0]*2 + h[1]*2 + h[2]*2
        return (int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))
    except: return None

_title_bg  = _hex_to_rgb(_p.get('title_bg_color')  or '') or (255, 255, 255)  # white
_title_fg  = _hex_to_rgb(_p.get('title_fg_color')  or '') or (0,   0,   0)    # black
_header_bg = _hex_to_rgb(_p.get('header_bg_color') or '') or (255, 255, 255)  # white
_header_fg = _hex_to_rgb(_p.get('header_fg_color') or '') or (0,   0,   0)    # black
LOGO_PATH  = _p.get('logo_path', '')

#  Helpers 

def natural_sort_key(s):
    parts = re.split(r'(\d+)', str(s))
    return [int(p) if p.isdigit() else p.lower() for p in parts]

def format_revit_date(rev):
    """Return date as Day\r\nMonth\r\nYear (three lines — Revit schedule line break)."""
    try:
        d = rev.RevisionDate
        if hasattr(d, 'Day'):
            return '{:02d}\r\n{:02d}\r\n{}'.format(d.Day, d.Month, d.Year)
        s = str(d).strip()
        for fmt in ('%d %B %Y','%B %d, %Y','%Y-%m-%d','%d/%m/%Y','%m/%d/%Y','%d.%m.%Y','%d/%m/%y','%d.%m.%y'):
            try:
                dt = _dt.strptime(s, fmt)
                return '{:02d}\r\n{:02d}\r\n{}'.format(dt.day, dt.month, dt.year)
            except Exception: pass
        return s
    except Exception: return ''

def rev_letter(seq):
    n = seq - 1
    if n < 0: return '?'
    result = ''
    while True:
        result = chr(65 + (n % 26)) + result
        n = n // 26 - 1
        if n < 0: break
    return result

def get_param(element, name):
    try:
        p = element.LookupParameter(name)
        if p and p.HasValue:
            return (p.AsString() or p.AsValueString() or "").strip()
    except Exception: pass
    return ""

def get_sf_by_id(sched_def, pid):
    for sf in sched_def.GetSchedulableFields():
        try:
            if sf.ParameterId.Value == pid:
                return sf
        except Exception: pass
    return None

def safe_merge(sec, r1, c1, r2, c2):
    try:
        from Autodesk.Revit.DB import TableMergedCell
        mc = TableMergedCell()
        mc.Top    = r1;  mc.Bottom = r2
        mc.Left   = c1;  mc.Right  = c2
        sec.MergeCells(mc)
    except Exception as e:
        output.print_md("   merge({},{},{},{}) {}".format(r1,c1,r2,c2,e))

def safe_text(sec, r, c, text):
    try: sec.SetCellText(r, c, str(text))
    except Exception as e:
        output.print_md("   text({},{}) {}".format(r,c,e))

def apply_style(sec, r, c, bold=False, size_mm=_DATA_MM, bg_rgb=None, italic=False,
                halign=None, valign="Middle", fg_rgb=None, font=None):
    """
    Apply cell style.
    opts.FontColor = flag to enable text colour override (TableCellStyleOverrideOptions)
    style.TextColor = the actual colour property (TableCellStyle)
    SetCellStyleOverrideOptions called AFTER all properties are set.
    """
    try:
        from Autodesk.Revit.DB import Color, HorizontalAlignmentStyle, VerticalAlignmentStyle, TableCellStyle

        style = TableCellStyle()

        # Set all property values first
        style.IsFontBold   = bold
        style.IsFontItalic = italic
        style.TextSize     = (size_mm / 0.75) * (72.0 / 25.4)
        if font:
            try: style.FontName = font
            except Exception: pass

        if bg_rgb:
            style.BackgroundColor = Color(bg_rgb[0], bg_rgb[1], bg_rgb[2])

        if fg_rgb:
            style.TextColor = Color(fg_rgb[0], fg_rgb[1], fg_rgb[2])

        if halign:
            style.FontHorizontalAlignment = getattr(
                HorizontalAlignmentStyle, halign, HorizontalAlignmentStyle.Left)

        style.FontVerticalAlignment = getattr(
            VerticalAlignmentStyle, valign, VerticalAlignmentStyle.Middle)

        # Enable overrides AFTER setting values
        opts = style.GetCellStyleOverrideOptions()
        opts.Bold                = True
        opts.Italics             = True
        opts.FontSize            = True
        opts.BackgroundColor     = (bg_rgb is not None)
        opts.FontColor           = (fg_rgb is not None)
        opts.HorizontalAlignment = (halign is not None)
        opts.VerticalAlignment   = True
        style.SetCellStyleOverrideOptions(opts)

        sec.SetCellStyle(r, c, style)
    except Exception as e:
        output.print_md("   style({},{}) {}".format(r, c, e))

def force_bg(sec, r, c, bg_rgb):
    """
    Force background colour on a cell regardless of AllowOverrideCellStyle.
    Used for non-anchor cells in merged rows where AllowOverrideCellStyle returns False.
    """
    try:
        from Autodesk.Revit.DB import Color
        style = sec.GetTableCellStyle(r, c)
        opts  = style.GetCellStyleOverrideOptions()
        opts.BackgroundColor  = True
        style.BackgroundColor = Color(bg_rgb[0], bg_rgb[1], bg_rgb[2])
        style.SetCellStyleOverrideOptions(opts)
        sec.SetCellStyle(r, c, style)
    except Exception:
        pass

def force_fg(sec, r, c, fg_rgb):
    """
    Force text colour on a cell by reading the existing style and patching TextColor.
    Use on merge anchor cells where apply_style text colour is not sticking.
    """
    try:
        from Autodesk.Revit.DB import Color
        style = sec.GetTableCellStyle(r, c)
        opts  = style.GetCellStyleOverrideOptions()
        opts.FontColor  = True
        style.TextColor = Color(fg_rgb[0], fg_rgb[1], fg_rgb[2])
        style.SetCellStyleOverrideOptions(opts)
        sec.SetCellStyle(r, c, style)
    except Exception as e:
        output.print_md("   force_fg({},{}) {}".format(r, c, e))

def remove_cell_borders(sec, r, c):
    """Remove all four borders (top/bottom/left/right) from a single cell."""
    try:
        style = sec.GetTableCellStyle(r, c)
        opts  = style.GetCellStyleOverrideOptions()
        opts.BorderTopLineStyle    = True
        opts.BorderBottomLineStyle = True
        opts.BorderLeftLineStyle   = True
        opts.BorderRightLineStyle  = True
        style.BorderTopLineStyle    = ElementId.InvalidElementId
        style.BorderBottomLineStyle = ElementId.InvalidElementId
        style.BorderLeftLineStyle   = ElementId.InvalidElementId
        style.BorderRightLineStyle  = ElementId.InvalidElementId
        style.SetCellStyleOverrideOptions(opts)
        sec.SetCellStyle(r, c, style)
    except Exception as e:
        output.print_md("   remove_cell_borders({},{}) {}".format(r, c, e))

def remove_lr_borders(sec, r, c):
    """Set left and right borders to white (invisible) on every cell in the row."""
    n = sec.NumberOfColumns
    for ci in range(n):
        try:
            style = sec.GetTableCellStyle(r, ci)
            opts  = style.GetCellStyleOverrideOptions()
            if _OFF_ID:
                opts.BorderLeftLineStyle  = True;  style.BorderLeftLineStyle  = _OFF_ID
                opts.BorderRightLineStyle = True;  style.BorderRightLineStyle = _OFF_ID
            else:
                opts.BorderLeftLineStyle  = False
                opts.BorderRightLineStyle = False
            style.SetCellStyleOverrideOptions(opts)
            sec.SetCellStyle(r, ci, style)
        except Exception:
            pass

def remove_row_borders(sec, r, n_cols):
    """Set all borders to white (invisible) on every cell in a row."""
    for ci in range(n_cols):
        try:
            style = sec.GetTableCellStyle(r, ci)
            opts  = style.GetCellStyleOverrideOptions()
            if _OFF_ID:
                opts.BorderTopLineStyle    = True;  style.BorderTopLineStyle    = _OFF_ID
                opts.BorderBottomLineStyle = True;  style.BorderBottomLineStyle = _OFF_ID
                opts.BorderLeftLineStyle   = True;  style.BorderLeftLineStyle   = _OFF_ID
                opts.BorderRightLineStyle  = True;  style.BorderRightLineStyle  = _OFF_ID
            else:
                opts.BorderTopLineStyle    = False
                opts.BorderBottomLineStyle = False
                opts.BorderLeftLineStyle   = False
                opts.BorderRightLineStyle  = False
            style.SetCellStyleOverrideOptions(opts)
            sec.SetCellStyle(r, ci, style)
        except Exception:
            pass

def clear_schedule_header(h):
    """Remove all extra rows and columns from a schedule header, leaving 1x1."""
    while h.NumberOfRows > 1:
        try: h.RemoveRow(h.NumberOfRows - 1)
        except Exception: break
    while h.NumberOfColumns > 1:
        try: h.RemoveColumn(h.NumberOfColumns - 1)
        except Exception: break

def get_or_create_schedule(doc, name):
    """Return existing schedule by name, or create a new one."""
    for v in FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements():
        if v.Name == name:
            return v, True   # (schedule, existed)
    vs = ViewSchedule.CreateSchedule(doc, ElementId.InvalidElementId)
    vs.Name = name
    return vs, False  # (schedule, existed)

#  Collect project info 

proj_info   = doc.ProjectInformation
proj_number = get_param(proj_info, "Project Number") or ""
proj_name   = get_param(proj_info, "Project Name")   or doc.Title or ""
org_name    = get_param(proj_info, "Organization Name") or ""
client_name = get_param(proj_info, "Client Name")    or ""

# Project info visibility flags from payload
_SHOW_FROM     = _p.get('show_from',     True)
_SHOW_CLIENT   = _p.get('show_client',   True)
_SHOW_PROJNO   = _p.get('show_projno',   True)
_SHOW_PROJNAME = _p.get('show_projname', True)

# Build list of visible project info rows (label, value)
_proj_info_rows = []
if _SHOW_FROM:     _proj_info_rows.append(("Organisation:", org_name))
if _SHOW_CLIENT:   _proj_info_rows.append(("Client:",       client_name))
if _SHOW_PROJNO:   _proj_info_rows.append(("Project No:",   proj_number))
if _SHOW_PROJNAME: _proj_info_rows.append(("Project:",      proj_name))
N_INFO = len(_proj_info_rows)

title_text = "Transmittal Document"

#  Collect issued revisions 

output.print_md("# pyTransmit — Schedule Creator")

all_revisions = list(FilteredElementCollector(doc).OfClass(DB.Revision).ToElements())
issued_revisions = sorted(
    [r for r in all_revisions if r.Issued], key=lambda r: r.SequenceNumber)
if not issued_revisions:
    forms.alert("No issued revisions found.", exitscript=True)
if len(issued_revisions) > MAX_REVS:
    issued_revisions = issued_revisions[-MAX_REVS:]

n_revs     = len(issued_revisions)
rev_start  = 3
TOTAL_COLS = 3 + MAX_REVS   # 13 cols — 3 fixed + 10 rev

def _tag(s, tag):
    """Parse TAG:value format from IssuedTo string, e.g. R:C M:E F:PDF S:A3"""
    import re as _re
    m = _re.search(r'(?:^| )' + _re.escape(tag) + r':([^ |]+)', s or '')
    return m.group(1) if m else ""

def parse_reason(s):
    """Parse reason code from IssuedTo string."""
    return _tag(s, "R")

rev_meta = []
for rev in issued_revisions:
    rev_meta.append({
        'letter':     rev_letter(rev.SequenceNumber),
        'date':       format_revit_date(rev),
        'initials':   (rev.IssuedBy or "").strip() or "XX",
        'reason':     parse_reason(rev.IssuedTo or ""),
        'method':     _tag(rev.IssuedTo, 'M'),
        'doc_format': _tag(rev.IssuedTo, 'F'),
        'paper_size': _tag(rev.IssuedTo, 'S'),
    })

output.print_md("**Revisions:** {}  |  **Title:** {}".format(n_revs, title_text))

#  Collect transmittal sheets 

issued_id_set = set(r.Id for r in issued_revisions)
all_sheets = list(FilteredElementCollector(doc)
    .OfCategory(BuiltInCategory.OST_Sheets)
    .WhereElementIsNotElementType().ToElements())
tx_sheets = sorted(
    [s for s in all_sheets
     if any(rid in issued_id_set for rid in s.GetAllRevisionIds())],
    key=lambda s: natural_sort_key(s.SheetNumber))
output.print_md("**Sheets:** {}".format(len(tx_sheets)))

#  Group sheets by selected parameters 

GROUP_PARAMS = _p.get('group_params') or []
GROUP_LABEL  = _p.get('group_label', True)

def get_sheet_param(sheet, param_name):
    """Return stripped string value of a sheet parameter, or '' if missing/blank."""
    try:
        p = sheet.LookupParameter(param_name)
        if p and p.HasValue:
            v = (p.AsString() or p.AsValueString() or '').strip()
            return v
    except:
        pass
    return ''

def get_group_label(sheet, params):
    """Build 'Folder \u2014 Stage' label from whichever params have values on this sheet."""
    parts = [get_sheet_param(sheet, pn) for pn in params]
    parts = [p for p in parts if p]   # drop blanks
    return u' \u2014 '.join(parts) if parts else ''

if GROUP_PARAMS:
    # Stable-sort preserving SheetNumber order within each group
    from collections import OrderedDict as _OD
    _groups = _OD()
    for s in tx_sheets:
        key = get_group_label(s, GROUP_PARAMS)
        _groups.setdefault(key, []).append(s)
    # grouped_sheets: list of (label, [sheets]) — '' label means ungrouped
    grouped_sheets = list(_groups.items())
else:
    grouped_sheets = [('', tx_sheets)]  # single group, no header row



# Line styles: pyT On (visible) and pyT Off (invisible/white)

def _get_or_create_line_style(name, rgb):
    try:
        from Autodesk.Revit.DB import GraphicsStyleType, Color
        _lines_cat = doc.Settings.Categories.get_Item('Lines')
        if _lines_cat is None:
            return None
        _existing = None
        for _sub in _lines_cat.SubCategories:
            if _sub.Name == name:
                _existing = _sub
                break
        if _existing is None:
            _existing = doc.Settings.Categories.NewSubcategory(_lines_cat, name)
        _existing.LineColor = Color(rgb[0], rgb[1], rgb[2])
        try:
            from Autodesk.Revit.DB import GraphicsStyleType as _GST
            _existing.SetLineWeight(1, _GST.Projection)
        except Exception:
            pass
        _gs = _existing.GetGraphicsStyle(GraphicsStyleType.Projection)
        return _existing.Id if _gs else None
    except Exception as _e:
        output.print_md("  line style {}: {}".format(name, _e))
        return None

_ON_ID  = None   # pyT On  — black visible border
_OFF_ID = None   # pyT Off — white invisible border
_BG_ID_CACHE = {}  # hex -> ElementId, for bg-colour-matched border styles

def _bg_line_id(hex_colour):
    """Return a line style ID whose colour matches a bg hex colour (cached)."""
    if not hex_colour: return _OFF_ID
    if hex_colour in _BG_ID_CACHE: return _BG_ID_CACHE[hex_colour]
    rgb = _hex_to_rgb(hex_colour) or (255, 255, 255)
    lid = _get_or_create_line_style('pyT ' + hex_colour.upper(), rgb)
    _BG_ID_CACHE[hex_colour] = lid
    return lid

# ── Text style helper (same as drafting view script) ─────────────────────────

def get_or_create_text_style(name, font, size_mm, bold=False, italic=False):
    """Return existing or newly-created TextNoteType, sized from JSON."""
    from Autodesk.Revit.DB import TextNoteType
    size_ft   = size_mm * MM
    all_types = list(FilteredElementCollector(doc).OfClass(TextNoteType).ToElements())
    existing  = None
    for tt in all_types:
        try:
            if tt.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == name:
                existing = tt; break
        except Exception: pass
    with Transaction(doc, "pyTransmit text style: {}".format(name)) as _tts:
        _tts.Start()
        new_tt = existing if existing else all_types[0].Duplicate(name)
        for bip, val in [
            (DB.BuiltInParameter.TEXT_SIZE,        size_ft),
            (DB.BuiltInParameter.TEXT_FONT,         font),
            (DB.BuiltInParameter.TEXT_STYLE_BOLD,   1 if bold else 0),
            (DB.BuiltInParameter.TEXT_STYLE_ITALIC, 0),
            (DB.BuiltInParameter.TEXT_BACKGROUND,   1),
            (DB.BuiltInParameter.TEXT_TAB_SIZE,     MM),
        ]:
            try:
                p = new_tt.get_Parameter(bip)
                if p and not p.IsReadOnly: p.Set(val)
            except Exception: pass
        _tts.Commit()
    return new_tt

# Create/update text types from JSON text_styles
_ts_data = {name: {
    'font':   s.get('font', 'Arial'),
    'size':   s.get('size_mm', 2.3),
    'bold':   bool(s.get('bold', False)),
} for name, s in _TEXT_STYLES.items()}

_TNTYPE = {}  # name -> TextNoteType (not used by schedule cells, but ensures types exist)
for _tsn, _tsd in _ts_data.items():
    try:
        _TNTYPE[_tsn] = get_or_create_text_style(
            'pyT Sched Body' if _tsn == 'Data' else 'pyT Sched {}'.format(_tsn),
            _tsd['font'], _tsd['size'], _tsd['bold'])
    except Exception as _e:
        output.print_md("  text style {}: {}".format(_tsn, _e))

with Transaction(doc, "pyTransmit — Line Styles") as _tls:
    _tls.Start()
    _ON_ID  = _get_or_create_line_style('pyT On',  (0,   0,   0  ))
    _OFF_ID = _get_or_create_line_style('pyT Off', (255, 255, 255))
    _tls.Commit()

#  Pre-calculate page count to clean up stale overflow schedules 
# This runs BEFORE the main transaction so we can safely delete in its own transaction.
import re as _re_pre
def _pre_calc_pages():
    try:
        from collections import OrderedDict as _OD2
        _all_s = list(FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_Sheets)
            .WhereElementIsNotElementType().ToElements())
        _issued_ids = set(r.Id for r in issued_revisions)
        _sheets = sorted(
            [s for s in _all_s if any(rid in _issued_ids for rid in s.GetAllRevisionIds())],
            key=lambda s: natural_sort_key(s.SheetNumber))
        if GROUP_PARAMS:
            _grps = _OD2()
            for s in _sheets:
                _grps.setdefault(get_group_label(s, GROUP_PARAMS), []).append(s)
            _rr = []
            _first_grp_done = False
            for gl, gs in _grps.items():
                # Always skip gap before the very first group (nothing above it to separate).
                # After the first group, always add the gap row (with or without label text).
                if gl and _first_grp_done:
                    _rr.append(('group', gl))
                for s in gs:
                    _rr.append(('sheet', s))
                _first_grp_done = True
        else:
            _rr = [('sheet', s) for s in _sheets]
        _fh = (ROW_TITLE + ROW_NORMAL * N_INFO + ROW_NORMAL + ROW_NORMAL + ROW_NORMAL * N_REC + ROW_NORMAL +
               (ROW_NORMAL + ROW_NORMAL * (1 + len(_filtered))) + ROW_NORMAL + ROW_NORMAL) / MM
        _rh = ROW_NORMAL / MM
        _p1 = max(1, int((A4_USABLE - _fh) / _rh) - 1)
        _pn = max(1, int((A4_USABLE - (ROW_NORMAL + ROW_NORMAL) / MM) / _rh) - 1)
        _pages = [_rr[:_p1]]
        _rem = _rr[_p1:]
        while _rem:
            _pages.append(_rem[:_pn]); _rem = _rem[_pn:]
        return len(_pages)
    except Exception:
        return 1

_pre_total_pages = _pre_calc_pages()

# Delete any pyTransmit Schedule XX-YY schedules where XX > _pre_total_pages
_stale_ids = []
for _vs in FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements():
    _m = _re_pre.match(r'^pyTransmit Schedule (\d+)-\d+$', _vs.Name)
    if _m and int(_m.group(1)) > _pre_total_pages:
        _stale_ids.append(_vs.Id)
if _stale_ids:
    with Transaction(doc, "pyTransmit — Remove stale schedules") as _tds:
        _tds.Start()
        for _sid in _stale_ids:
            try: doc.Delete(_sid)
            except Exception: pass
        _tds.Commit()
    output.print_md("  Removed {} stale schedule(s)".format(len(_stale_ids)))

#  Get or create main schedule 

def _parse_copies_for_recipient(issued_to_str, recipient_label, recipient_index=0):
    import re as _re
    if not issued_to_str: return ''
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
        m2 = _re.search(r'(?:^| )' + _re.escape(first) + r'\.\[([^\]]*)\](\d+)', _block)
        if m2: return m2.group(2)
    if recipient_label:
        m3 = _re.search(r'\[' + _re.escape(recipient_label[:6]) + r'[^\]]*\](\d+)', issued_to_str)
        if m3: return m3.group(1)
    return ''

# ── JSON-driven row plan ──────────────────────────────────────────────────────
# Walk _ROWS from the layout JSON and build a flat list of "render items":
#   ('fixed',  ri, blocks)         — one schedule row
#   ('data',   ri, blocks, items)  — expands to len(items) rows
#   ('sheets', ri, blocks)         — expands to sheet rows (page-split here)

def _block_just(b):
    j = b.get('just', 'left')
    return {'left': 'Left', 'center': 'Center', 'right': 'Right'}.get(j, 'Left')

def _block_size(b):
    ts = _TEXT_STYLES.get(b.get('text_style', 'Data'), {})
    return ts.get('size_mm', _DATA_MM)

def _block_bold(b):
    ts = _TEXT_STYLES.get(b.get('text_style', 'Data'), {})
    return bool(ts.get('bold', False))

def _block_font(b):
    ts = _TEXT_STYLES.get(b.get('text_style', 'Data'), {})
    return ts.get('font', 'Arial')

def _apply_block_cell_borders(hdr, ri, ec_s, ec_e, brd):
    """Apply border show/hide from JSON block borders dict to a merged cell range.
    For merged cells: top/bottom apply to all cells, left only to ec_s, right only to ec_e.
    Interior vertical borders (between merged cells) are always hidden.
    """
    def _set_border(sty, opts, attr, show):
        setattr(opts, attr, True)
        if show:
            setattr(sty, attr, _ON_ID if _ON_ID else ElementId.InvalidElementId)
        elif _OFF_ID:
            setattr(sty, attr, _OFF_ID)
        else:
            setattr(opts, attr, False)

    for _ci in range(ec_s, ec_e + 1):
        try:
            _sty  = hdr.GetTableCellStyle(ri, _ci)
            _opts = _sty.GetCellStyleOverrideOptions()
            # Top and bottom apply to every cell in the range
            _set_border(_sty, _opts, 'BorderTopLineStyle',    brd.get('t', False))
            _set_border(_sty, _opts, 'BorderBottomLineStyle', brd.get('b', False))
            # Left only on the first cell; interior cells left to ShowGridLines
            if _ci == ec_s:
                _set_border(_sty, _opts, 'BorderLeftLineStyle', brd.get('l', False))
            # Right only on the last cell; interior cells left to ShowGridLines
            if _ci == ec_e:
                _set_border(_sty, _opts, 'BorderRightLineStyle', brd.get('r', False))
            _sty.SetCellStyleOverrideOptions(_opts)
            hdr.SetCellStyle(ri, _ci, _sty)
        except Exception: pass

def _apply_data_row_borders(hdr, ri, ec_s, ec_e, is_last, outer_b, data_b, rev_cols=False):
    """Apply borders for a data row (sheet/recip).
    - Top: always from outer_b (first row) or hidden (subsequent rows handled by prev bottom)
    - Bottom: outer_b on last row, data_b.h grid line otherwise
    - Left: outer_b on leftmost cell only
    - Right: outer_b on rightmost cell only; data_b.v between rev columns
    """
    _h = data_b.get('h', True)
    _v = data_b.get('v', True)
    _last_col = TOTAL_COLS - 1 if rev_cols else ec_e

    def _sb(sty, opts, attr, show):
        setattr(opts, attr, True)
        if show:
            setattr(sty, attr, _ON_ID if _ON_ID else ElementId.InvalidElementId)
        elif _OFF_ID:
            setattr(sty, attr, _OFF_ID)
        else:
            setattr(opts, attr, False)

    for _ci in range(ec_s, TOTAL_COLS if rev_cols else ec_e + 1):
        try:
            _sty  = hdr.GetTableCellStyle(ri, _ci)
            _opts = _sty.GetCellStyleOverrideOptions()
            _in_rev = _ci >= REV_START
            # Top: show outer top only (Revit handles row separation via bottom of prev row)
            _sb(_sty, _opts, 'BorderTopLineStyle', outer_b.get('t', True))
            # Bottom: outer on last row, h-grid otherwise
            _sb(_sty, _opts, 'BorderBottomLineStyle',
                outer_b.get('b', True) if is_last else _h)
            # Left: outer on leftmost col; v-grid on rev cols; leave others to ShowGridLines
            if _ci == ec_s:
                _sb(_sty, _opts, 'BorderLeftLineStyle', outer_b.get('l', True))
            elif _in_rev:
                _sb(_sty, _opts, 'BorderLeftLineStyle', _v)
            # Right: outer on rightmost col; v-grid on rev cols; leave others to ShowGridLines
            if _ci == _last_col:
                _sb(_sty, _opts, 'BorderRightLineStyle', outer_b.get('r', True))
            elif _in_rev:
                _sb(_sty, _opts, 'BorderRightLineStyle', _v)
            _sty.SetCellStyleOverrideOptions(_opts)
            hdr.SetCellStyle(ri, _ci, _sty)
        except Exception: pass

def _block_borders(b):
    brd = b.get('borders', {})
    return {
        't': brd.get('t', True),
        'b': brd.get('b', True),
        'l': brd.get('l', False),
        'r': brd.get('r', False),
    }

# Resolve column span → Excel-style column start/end
# Layout cols: A=0, B=1, C=2, D=rev columns (3..3+MAX_REVS-1)
REV_START  = 3
LAST_COL   = REV_START + MAX_REVS - 1
TOTAL_COLS = REV_START + MAX_REVS

def _col_range(ci, span, blocks):
    """Return (ec_start, ec_end) for block at canvas col ci with given span."""
    # canvas col 0→excel 0, 1→1, 2→2, 3(D)→REV_START..LAST_COL
    if ci == 3 or ci >= 3:
        return REV_START, LAST_COL
    ec_start = ci
    # span across layout cols
    ec_end = min(ci + span - 1, 2)  # caps at col C (=2) unless into D
    if ci + span - 1 >= 3:          # spans into D column
        ec_end = LAST_COL
    return ec_start, ec_end

# Build render plan: list of dicts
# { 'type': 'fixed'|'data'|'sheets', 'row': JSON row dict,
#   'height': float(ft), 'items': list (for data/sheets) }

render_plan = []
_sheet_render_rows = []  # ('group', label) or ('sheet', sheet_obj)
_first_grp_done = False
for grp_label, grp_sheets in grouped_sheets:
    # Skip group row before the first group, only add gaps between groups.
    if grp_label and _first_grp_done:
        _sheet_render_rows.append(('group', grp_label))
    for s in grp_sheets:
        _sheet_render_rows.append(('sheet', s))
    _first_grp_done = True

for _lri, row in enumerate(_ROWS):
    sec = row.get('section', 'body')
    if sec == 'footer':
        continue
    blocks = row.get('blocks', [])
    # Determine if this row expands
    _types = [b.get('type','') for b in blocks if b]
    if any(t in ('sent_to','attn_to','spine_copies') for t in _types):
        render_plan.append({'kind': 'recip', 'row': row, 'height': ROW_NORMAL,
                            'items': list(range(len(RECIPIENTS))), 'layout_ri': _lri})
    elif any(t in ('sheet_number','sheet_desc','spine_rev') for t in _types):
        render_plan.append({'kind': 'sheets', 'row': row, 'height': ROW_NORMAL,
                            'items': _sheet_render_rows, 'layout_ri': _lri})
    else:
        render_plan.append({'kind': 'fixed', 'row': row, 'height': ROW_NORMAL, 'layout_ri': _lri})

# ── Page splitting ────────────────────────────────────────────────────────────
# Count fixed header height (everything except the sheet-expansion rows)
_fixed_h = sum(
    p['height'] for p in render_plan if p['kind'] != 'sheets'
) / MM

_row_h_mm = ROW_NORMAL / MM
_slim_h_mm = (ROW_NORMAL * 2) / MM  # overflow page: doc list + col header

rows_page1 = max(1, int((A4_USABLE - _fixed_h) / _row_h_mm) - 1) if _SPLIT else 999999
rows_page_n = max(1, int((A4_USABLE - _slim_h_mm) / _row_h_mm) - 1) if _SPLIT else 999999

# Split sheet rows into pages
_all_sheet_items = _sheet_render_rows
pages = [_all_sheet_items[:rows_page1]]
_rem  = _all_sheet_items[rows_page1:]
while _rem:
    pages.append(_rem[:rows_page_n]); _rem = _rem[rows_page_n:]

total_pages  = max(1, len(pages))
SCHEDULE_NAME = "pyTransmit Schedule 01-{:02d}".format(total_pages)

issued_date_long = parse_date_short(rev_meta[n_revs-1]['date']) if n_revs > 0 else ''

output.print_md("**Pages:** {}  |  Page1 rows: {}  |  Page2+ rows: {}".format(
    total_pages, rows_page1, rows_page_n))

# ── Helper: render one page of the schedule into hdr section ─────────────────

def _render_page(hdr, page_sheet_items, is_first_page):
    """
    Build all rows/merges/text/styles for one schedule page.
    Returns list of group row indices (for border cleanup).
    """
    last_col = TOTAL_COLS - 1

    # ── Insert columns ────────────────────────────────────────────────────────
    while hdr.NumberOfColumns < TOTAL_COLS:
        hdr.InsertColumn(hdr.NumberOfColumns)
    col_widths = [C_A, C_B, C_C] + [C_REV] * MAX_REVS
    for ci, w in enumerate(col_widths):
        try: hdr.SetColumnWidth(ci, w)
        except Exception: pass

    # ── Build row list for this page ─────────────────────────────────────────
    # Each item: {'ri': int, 'kind': str, 'row': dict, 'item': any, 'is_last': bool}
    page_rows = []
    ri = 0

    if is_first_page:
        _plan = render_plan
    else:
        # Overflow page: only show doc-list header, col-header, sheets, footer
        _sheet_row = next((p for p in render_plan if p['kind'] == 'sheets'), None)
        _sheet_fixed = [p for p in render_plan
                        if any(b and b.get('type','') in ('text',) for b in p['row'].get('blocks',[]))
                        and p['kind'] == 'fixed'
                        and any(b and b.get('content','') in ('Documentation List','Sheet','Description','Revision')
                                or b and b.get('label','') in ('Documentation List','Sheet','Description','Revision')
                                for b in p['row'].get('blocks',[]) if b)]
        # Simpler: just emit doc-list row + col-header row + sheets
        _plan_overflow = []
        _in_doc = False
        for p in render_plan:
            _blk_types = [b.get('type','') for b in p['row'].get('blocks',[]) if b]
            _blk_labels = [b.get('label','') or b.get('content','') for b in p['row'].get('blocks',[]) if b]
            if 'Documentation List' in _blk_labels or any(t == 'text' and 'Documentation' in str(p['row']) for t in _blk_types):
                _plan_overflow.append(p); _in_doc = True
            elif _in_doc and any(t == 'text' for t in _blk_types) and any(
                    l in ('Sheet','Description','Revision') for l in _blk_labels):
                _plan_overflow.append(p)
            elif p['kind'] == 'sheets':
                _plan_overflow.append(p)
        _plan = _plan_overflow if _plan_overflow else [p for p in render_plan if p['kind'] == 'sheets']

    for plan_item in _plan:
        kind = plan_item['kind']
        row  = plan_item['row']
        _plan_lri = plan_item.get('layout_ri', 0)
        if kind == 'fixed':
            page_rows.append({'ri': ri, 'kind': 'fixed', 'row': row, 'item': None, 'layout_ri': _plan_lri})
            ri += 1
        elif kind == 'recip':
            if not is_first_page: continue
            for idx, _ in enumerate(RECIPIENTS):
                page_rows.append({'ri': ri, 'kind': 'recip', 'row': row,
                                  'item': idx, 'is_last': idx == len(RECIPIENTS)-1, 'layout_ri': _plan_lri})
                ri += 1
        elif kind == 'sheets':
            items = page_sheet_items
            for idx, sr in enumerate(items):
                # Treat as last if final item, next item is a group header row,
                # or the current sheet and next sheet belong to different groups.
                _next = items[idx + 1] if idx + 1 < len(items) else None
                _is_last_s = (idx == len(items) - 1)
                if not _is_last_s and _next is not None:
                    if _next[0] == 'group':
                        _is_last_s = True
                    elif _next[0] == 'sheet' and sr[0] == 'sheet' and GROUP_PARAMS:
                        # Detect group boundary when group rows were skipped
                        _cur_gl  = get_group_label(sr[1],   GROUP_PARAMS)
                        _next_gl = get_group_label(_next[1], GROUP_PARAMS)
                        if _cur_gl != _next_gl:
                            _is_last_s = True
                page_rows.append({'ri': ri, 'kind': 'sheet_row', 'row': row,
                                  'item': sr, 'is_last': _is_last_s, 'layout_ri': _plan_lri})
                ri += 1

    total_rows = ri
    footer_ri  = -1  # no hardcoded footer

    # ── Insert rows ───────────────────────────────────────────────────────────
    while hdr.NumberOfRows < total_rows:
        hdr.InsertRow(hdr.NumberOfRows)
    for pr in page_rows:
        _h = ROW_NORMAL
        _blocks = (pr['row'].get('blocks',[]) if pr['row'] else [])
        _btypes = [b.get('type','') for b in _blocks if b]
        if any(t == 'spine_dates' for t in _btypes):
            _h = ROW_DATE_H
        try: hdr.SetRowHeight(pr['ri'], _h)
        except Exception: pass

    # ── Merges, text and styles ───────────────────────────────────────────────
    group_row_indices = []
    blank_row_indices  = []
    border_tasks      = []  # (kind, sched_ri, ec_s, ec_e, brd, extra, layout_ri, layout_ci)

    for pr in page_rows:
        ri2   = pr['ri']
        kind2 = pr['kind']

        if kind2 == 'footer':
            continue  # footer driven by JSON only — no hardcoded row

        row2   = pr['row']
        blocks = row2.get('blocks', [])
        ci     = 0  # canvas column index

        for b in blocks:
            if b is None:
                ci += 1
                continue
            t    = b.get('type', '')
            span = int(b.get('span', 1))
            ec_s, ec_e = _col_range(ci, span, blocks)
            bold = _block_bold(b)
            sz   = _block_size(b)
            just = _block_just(b)
            brd  = _block_borders(b)

            # ── fixed text block ──────────────────────────────────────────────
            if t == 'text':
                label = b.get('label') or b.get('content', '')
                if ec_s != ec_e:
                    safe_merge(hdr, ri2, ec_s, ri2, ec_e)
                safe_text(hdr, ri2, ec_s, label)
                apply_style(hdr, ri2, ec_s, bold=bold, size_mm=sz, halign=just,
                            bg_rgb=_header_bg if bold else None,
                            fg_rgb=_header_fg if bold else None, font=_block_font(b))
                if bold:
                    for _c in range(ec_s, ec_e + 1):
                        try:
                            _sty = hdr.GetTableCellStyle(ri2, _c)
                            _opt = _sty.GetCellStyleOverrideOptions()
                            # Apply _ON_ID only on sides JSON requests; hide others
                            for _attr, _side in [
                                ('BorderTopLineStyle',    brd.get('t', False)),
                                ('BorderBottomLineStyle', brd.get('b', False)),
                                ('BorderLeftLineStyle',   brd.get('l', False) if _c == ec_s else False),
                                ('BorderRightLineStyle',  brd.get('r', False) if _c == ec_e else False),
                            ]:
                                setattr(_opt, _attr, True)
                                setattr(_sty, _attr, _ON_ID if _side else _OFF_ID)
                            _sty.SetCellStyleOverrideOptions(_opt)
                            hdr.SetCellStyle(ri2, _c, _sty)
                        except Exception: pass
                    force_bg(hdr, ri2, ec_s, _header_bg)
                    force_fg(hdr, ri2, ec_s, _header_fg)
                else:
                    # Non-bold: queue for border cleanup transaction
                    border_tasks.append(("block", ri2, ec_s, ec_e, brd, None, pr.get('layout_ri',0), ci))

            # ── blank ─────────────────────────────────────────────────────────
            elif t == 'blank':
                safe_merge(hdr, ri2, ec_s, ri2, ec_e)
                remove_row_borders(hdr, ri2, TOTAL_COLS)
                if ri2 not in blank_row_indices:
                    blank_row_indices.append(ri2)

            # ── spine_dates ───────────────────────────────────────────────────
            elif t == 'spine_dates':
                _rot    = b.get('rotation', 270)
                _dfmt   = b.get('date_format', 'dd/MM/yyyy')
                for j in range(MAX_REVS):
                    _raw = rev_meta[j]['date'] if j < n_revs else ''
                    val  = parse_date_short(_raw.replace('\r\n', '/').replace('\r', '/').replace('\n', '/')) if _raw else ''
                    safe_text(hdr, ri2, REV_START + j, val)
                    apply_style(hdr, ri2, REV_START + j, bold=False, size_mm=_DATA_MM,
                                halign='Center', valign='Middle')
                    if _rot in (90, 270):
                        try:
                            from Autodesk.Revit.DB import TableCellStyle as _TCSR
                            _rs = hdr.GetTableCellStyle(ri2, REV_START + j)
                            _ro = _rs.GetCellStyleOverrideOptions()
                            _rs.TextOrientation = _rot
                            _ro.TextOrientation = True
                            _rs.SetCellStyleOverrideOptions(_ro)
                            hdr.SetCellStyle(ri2, REV_START + j, _rs)
                        except Exception: pass
                border_tasks.append(("block", ri2, REV_START, LAST_COL, brd, None, pr.get('layout_ri',0), ci))

            # ── reason_list / method_list ─────────────────────────────────────
            elif t in ('reason_list', 'method_list'):
                _leg = REASON_LEGEND_1LINE if t == 'reason_list' else METHOD_LEGEND_1LINE
                if ec_s != ec_e:
                    safe_merge(hdr, ri2, ec_s, ri2, ec_e)
                safe_text(hdr, ri2, ec_s, _leg)
                apply_style(hdr, ri2, ec_s, bold=False, size_mm=sz, halign='Left', valign='Top')
                _apply_block_cell_borders(hdr, ri2, ec_s, ec_e, brd)

            # ── spine_reason / spine_method / spine_initials / etc ────────────
            elif t in ('spine_reason','spine_method','spine_initials',
                       'spine_doc_type','spine_print_size'):
                _key_map = {'spine_reason':'reason','spine_method':'method',
                            'spine_initials':'initials','spine_doc_type':'doc_format',
                            'spine_print_size':'paper_size'}
                _key = _key_map.get(t, 'initials')
                for j in range(MAX_REVS):
                    val = rev_meta[j][_key] if j < n_revs else ''
                    safe_text(hdr, ri2, REV_START + j, val)
                    apply_style(hdr, ri2, REV_START + j, bold=False, size_mm=sz, halign='Center')
                _apply_block_cell_borders(hdr, ri2, REV_START, LAST_COL, brd)

            # ── recipient blocks (sent_to / attn_to / spine_copies) ───────────
            elif kind2 == 'recip':
                idx = pr['item']
                rec = RECIPIENTS[idx]
                _attn = (_p['recipients'][idx].get('attn','')
                         if _p.get('recipients') and idx < len(_p['recipients']) else '')
                if t == 'sent_to':
                    safe_text(hdr, ri2, ec_s, rec)
                    apply_style(hdr, ri2, ec_s, bold=False, size_mm=sz, halign=just, font=_block_font(b))
                elif t == 'attn_to':
                    if ec_s != ec_e: safe_merge(hdr, ri2, ec_s, ri2, ec_e)
                    safe_text(hdr, ri2, ec_s, _attn)
                    apply_style(hdr, ri2, ec_s, bold=False, size_mm=sz, halign=just, font=_block_font(b))
                elif t == 'spine_copies':
                    for j in range(MAX_REVS):
                        _copies = ''
                        if j < n_revs:
                            _copies = _parse_copies_for_recipient(
                                issued_revisions[j].IssuedTo or '', rec, idx)
                        safe_text(hdr, ri2, REV_START + j, _copies)
                        apply_style(hdr, ri2, REV_START + j, bold=False, size_mm=sz, halign='Center')
            # Queue data-row borders for recipient rows
            if kind2 == 'recip':
                _is_last_r = pr.get('is_last', False)
                border_tasks.append(("data", ri2, 0, 2, brd,
                                     (b.get('data_borders', {'h':True,'v':True}), _is_last_r),
                                     pr.get('layout_ri',0), ci))

            # ── sheet rows ────────────────────────────────────────────────────
            elif kind2 == 'sheet_row':
                sr_kind, sr_item = pr['item']
                if sr_kind == 'group':
                    safe_merge(hdr, ri2, 0, ri2, last_col)
                    if GROUP_LABEL:
                        safe_text(hdr, ri2, 0, sr_item)
                        apply_style(hdr, ri2, 0, bold=True, size_mm=_DATA_MM, halign='Left')
                    group_row_indices.append(ri2)
                else:
                    sheet = sr_item
                    if t == 'sheet_number':
                        safe_text(hdr, ri2, ec_s, str(sheet.SheetNumber))
                        apply_style(hdr, ri2, ec_s, bold=False, size_mm=sz, halign=just, font=_block_font(b))
                    elif t == 'sheet_desc':
                        if ec_s != ec_e: safe_merge(hdr, ri2, ec_s, ri2, ec_e)
                        safe_text(hdr, ri2, ec_s, str(sheet.Name))
                        apply_style(hdr, ri2, ec_s, bold=False, size_mm=sz, halign=just, font=_block_font(b))
                    elif t == 'spine_rev':
                        _srev_ids = set(sheet.GetAllRevisionIds())
                        for j in range(MAX_REVS):
                            _vlines = b.get('data_borders', {}).get('v', False)
                            val = ''
                            if j < n_revs:
                                _rv = issued_revisions[j]
                                val = rev_letter(_rv.SequenceNumber) if _rv.Id in _srev_ids else ''
                            safe_text(hdr, ri2, REV_START + j, val)
                            apply_style(hdr, ri2, REV_START + j, bold=False,
                                        size_mm=sz, halign='Center')
                # Queue data-row borders for sheet rows
                if sr_kind == 'sheet':
                    _is_last_s = pr.get('is_last', False)
                    border_tasks.append(("data", ri2, 0, 2, brd,
                                         (b.get('data_borders', {'h':True,'v':True}), _is_last_s),
                                         pr.get('layout_ri',0), ci))

            ci += span

    return group_row_indices, footer_ri, blank_row_indices, border_tasks


# ── Main schedule transaction ─────────────────────────────────────────────────

with Transaction(doc, "pyTransmit \u2014 Update Schedule") as t:
    t.Start()

    sched, existed = get_or_create_schedule(doc, SCHEDULE_NAME)
    sched_def = sched.Definition

    if existed:
        output.print_md(" Updating existing schedule")
        for fid in list(sched_def.GetFieldOrder()):
            try: sched_def.RemoveField(fid)
            except Exception: pass
        for fi in list(sched_def.GetFilters()):
            try: sched_def.RemoveFilter(fi)
            except Exception: pass
    else:
        output.print_md(" Creating new schedule")

    sf_asm    = get_sf_by_id(sched_def, FIELD_ID_ASM_CODE)
    field_asm = sched_def.AddField(sf_asm)
    field_asm.ColumnHeading = ""

    filt1 = ScheduleFilter(field_asm.FieldId, ScheduleFilterType.Equal, "NO VALUES FOUND")
    sched_def.AddFilter(filt1)
    filt2 = ScheduleFilter(field_asm.FieldId, ScheduleFilterType.Equal, "All VALUES FOUND")
    sched_def.AddFilter(filt2)

    table_data = sched.GetTableData()
    hdr  = table_data.GetSectionData(SectionType.Header)
    body = table_data.GetSectionData(SectionType.Body)

    if existed:
        clear_schedule_header(hdr)

    total_w = C_A + C_B + C_C + C_REV * MAX_REVS
    body.SetColumnWidth(0, total_w)

    try:
        from Autodesk.Revit.DB import TableCellStyle as _TCS
        _bs = _TCS()
        _bo = _bs.GetCellStyleOverrideOptions()
        if _OFF_ID:
            _bo.BorderTopLineStyle    = True;  _bs.BorderTopLineStyle    = _OFF_ID
            _bo.BorderBottomLineStyle = True;  _bs.BorderBottomLineStyle = _OFF_ID
            _bo.BorderLeftLineStyle   = True;  _bs.BorderLeftLineStyle   = _OFF_ID
            _bo.BorderRightLineStyle  = True;  _bs.BorderRightLineStyle  = _OFF_ID
        else:
            _bo.BorderTopLineStyle = False; _bo.BorderBottomLineStyle = False
            _bo.BorderLeftLineStyle = False; _bo.BorderRightLineStyle = False
        _bs.SetCellStyleOverrideOptions(_bo)
        body.SetCellStyle(_bs)
    except Exception as e:
        output.print_md("  body border clear: {}".format(e))

    try: sched_def.ShowGridLines = True
    except Exception: pass

    # Render page 1
    _p1_group_ris, _p1_footer_ri, _p1_blank_ris, _p1_border_tasks = _render_page(hdr, pages[0], is_first_page=True)

    t.Commit()

# ── Load Style Algorithm for border/colour resolution ────────────────────────

_style_algo     = None
_sched_cell_styles = {}

import os as _os_sa
_sa_dir = _os_sa.path.dirname(_os_sa.path.abspath(__file__))
for _sa_cand_dir in [_sa_dir, _os_sa.path.join(_sa_dir, 'Publish')]:
    _sa_path = _os_sa.path.join(_sa_cand_dir, 'layout_Style_Algoruthum.py')
    if _os_sa.path.isfile(_sa_path):
        try:
            import imp as _imp_sa
            _style_algo = _imp_sa.load_source('layout_Style_Algoruthum', _sa_path)
            _hl = {int(k): v for k,v in LAYOUT.get('hlines', {}).items()}
            _vl = {int(k): v for k,v in LAYOUT.get('vlines', {}).items()}
            if _hl or _vl:
                _sched_cell_styles = _style_algo.compute_cell_styles_from_grid(
                    rows        = LAYOUT.get('rows', []),
                    groups      = _style_algo.get_groups(LAYOUT.get('rows', [])),
                    hlines      = _hl,
                    vlines      = _vl,
                    text_styles = LAYOUT.get('text_styles', {}),
                    max_rev_col = REV_START,
                    n_revs      = MAX_REVS,
                )
            else:
                _sched_cell_styles = _style_algo.compute_cell_styles(
                    rows        = LAYOUT.get('rows', []),
                    groups      = _style_algo.get_groups(LAYOUT.get('rows', [])),
                    text_styles = LAYOUT.get('text_styles', {}),
                    max_rev_col = REV_START,
                    n_revs      = MAX_REVS,
                )
            output.print_md("   Style algo loaded: `{}`".format(_sa_path))
        except Exception as _sa_err:
            output.print_md("   Style algo load failed: {}".format(_sa_err))
        break

def _get_sched_style(layout_ri, layout_ci):
    """Get pre-computed cell style for a layout row/col."""
    return _sched_cell_styles.get((layout_ri, layout_ci), {})

def _apply_algo_border(hdr_sec, sched_ri, layout_ri, layout_ci):
    """Apply style-algo border to a Revit schedule cell."""
    cs = _get_sched_style(layout_ri, layout_ci)
    if not cs and not _style_algo:
        return
    try:
        _sty  = hdr_sec.GetTableCellStyle(sched_ri, layout_ci if layout_ci < REV_START else REV_START)
        _opts = _sty.GetCellStyleOverrideOptions()
        def _sb(attr, show):
            setattr(_opts, attr, True)
            setattr(_sty, attr, (_ON_ID if _ON_ID else ElementId.InvalidElementId) if show else _OFF_ID)
        _sb('BorderTopLineStyle',    cs.get('t', False))
        _sb('BorderBottomLineStyle', cs.get('b', False))
        _sb('BorderLeftLineStyle',   cs.get('l', False))
        _sb('BorderRightLineStyle',  cs.get('r', False))
        _sty.SetCellStyleOverrideOptions(_opts)
        hdr_sec.SetCellStyle(sched_ri, layout_ci if layout_ci < REV_START else REV_START, _sty)
    except Exception: pass

# ── Border cleanup transaction — algo-driven ──────────────────────────────────

with Transaction(doc, "pyTransmit footer") as tf:
    tf.Start()

    hdr_f = sched.GetTableData().GetSectionData(SectionType.Header)

    def _set_cell_border(sec, r, c, show_t, show_b, show_l, show_r):
        try:
            _sty  = sec.GetTableCellStyle(r, c)
            _opts = _sty.GetCellStyleOverrideOptions()
            def _sb(attr, show):
                setattr(_opts, attr, True)
                setattr(_sty, attr,
                    (_ON_ID if _ON_ID else ElementId.InvalidElementId) if show else _OFF_ID)
            _sb('BorderTopLineStyle',    show_t)
            _sb('BorderBottomLineStyle', show_b)
            _sb('BorderLeftLineStyle',   show_l)
            _sb('BorderRightLineStyle',  show_r)
            _sty.SetCellStyleOverrideOptions(_opts)
            sec.SetCellStyle(r, c, _sty)
        except Exception: pass

    if _style_algo and _sched_cell_styles:
        # Build map: sched_ri -> layout_ri, is_last, data_borders
        _sched_to_layout = {}
        _sched_is_last   = {}
        _sched_data_b    = {}
        for _btask in _p1_border_tasks:
            _bkind = _btask[0]; _bri = _btask[1]; _bextra = _btask[5]
            _blri  = _btask[6] if len(_btask) >= 8 else 0
            _sched_to_layout[_bri] = _blri
            if _bkind == "data" and _bextra:
                _sched_is_last[_bri] = _bextra[1]
                _sched_data_b[_bri]  = _bextra[0] if isinstance(_bextra[0], dict) else {'h':True,'v':True}

        for _sched_ri, _layout_ri in _sched_to_layout.items():
            _is_last = _sched_is_last.get(_sched_ri, True)
            _data_b  = _sched_data_b.get(_sched_ri, {})
            _h       = _data_b.get('h', True)
            _v       = _data_b.get('v', True)
            _is_data = _sched_ri in _sched_is_last

            for _lci in range(4):
                _cs = _get_sched_style(_layout_ri, _lci)
                if not _cs:
                    continue
                _show_t = _cs.get('t', False)
                # Group boundary rows need h-line to close the group block.
                # Truly last row uses the cell style's outer bottom border.
                # Intermediate data rows use h for row dividers.
                _is_group_boundary = _is_last and _sched_is_last.get(_sched_ri) is True and _sched_ri != max(_sched_is_last.keys()) if _sched_is_last else False
                if not _is_data:
                    _show_b = _cs.get('b', False)
                elif _is_last and _sched_ri == max(_sched_is_last.keys()):
                    _show_b = _cs.get('b', False)   # truly last sheet row, use outer border
                elif _is_last:
                    _show_b = _h                     # group boundary, close with h-line
                else:
                    _show_b = _h                     # intermediate row divider
                _show_l = _cs.get('l', False)
                _show_r = _cs.get('r', False)

                if _lci < REV_START:
                    _set_cell_border(hdr_f, _sched_ri, _lci, _show_t, _show_b, _show_l, _show_r)
                else:
                    for _rci in range(MAX_REVS):
                        _sci = REV_START + _rci
                        _l = _show_l if _rci == 0 else _v
                        _r = _show_r if _rci == MAX_REVS - 1 else _v
                        _set_cell_border(hdr_f, _sched_ri, _sci, _show_t, _show_b, _l, _r)

        # Blank rows — mirror row above bottom for top, hide bottom/left/right
        for _blank_ri in _p1_blank_ris:
            _above_layout_ri = _sched_to_layout.get(_blank_ri - 1)
            for _ci in range(TOTAL_COLS):
                # Top mirrors what row above set as bottom
                _lci_above = _ci if _ci < REV_START else REV_START
                _cs_above = _get_sched_style(_above_layout_ri, _lci_above) if _above_layout_ri is not None else {}
                _above_b = _cs_above.get('b', False)
                _set_cell_border(hdr_f, _blank_ri, _ci, _above_b, False, False, False)

        # Group header rows, draw top/bottom borders from data_borders.h
        _data_h = True  # data_borders.h default
        for _bt in _p1_border_tasks:
            if _bt[0] == "data" and _bt[5]:
                _data_h = _bt[5][0].get('h', True) if isinstance(_bt[5][0], dict) else True
                break
        _first_gri = min(_p1_group_ris) if _p1_group_ris else None
        for gri in _p1_group_ris:
            _is_first_group_row = (gri == _first_gri)
            for _ci in range(TOTAL_COLS):
                _lci_g = _ci if _ci < REV_START else REV_START
                _cs_g  = _get_sched_style(_sched_to_layout.get(gri, 0), _lci_g)
                _show_l_g = _cs_g.get('l', False) if _ci == 0 else False
                _show_r_g = _cs_g.get('r', False) if _ci == TOTAL_COLS - 1 else False
                # No top border on first group row, it sits directly under the column headers
                _top_g = False if _is_first_group_row else _data_h
                _set_cell_border(hdr_f, gri, _ci, _top_g, _data_h, _show_l_g, _show_r_g)

    tf.Commit()
# ── Overflow pages ────────────────────────────────────────────────────────────

for page_idx, page_items in enumerate(pages[1:], start=2):
    sched_name = "pyTransmit Schedule {:02d}-{:02d}".format(page_idx, total_pages)

    with Transaction(doc, "pyTransmit page {}".format(page_idx)) as t2:
        t2.Start()

        vs2, existed2 = get_or_create_schedule(doc, sched_name)
        sd2 = vs2.Definition
        sd2.IsItemized = True

        if existed2:
            for fid in list(sd2.GetFieldOrder()):
                try: sd2.RemoveField(fid)
                except Exception: pass
            for fi in list(sd2.GetFilters()):
                try: sd2.RemoveFilter(fi)
                except Exception: pass

        sf2 = get_sf_by_id(sd2, FIELD_ID_ASM_CODE)
        if sf2:
            f2 = sd2.AddField(sf2)
            f2.ColumnHeading = ""
            sd2.AddFilter(ScheduleFilter(f2.FieldId, ScheduleFilterType.Equal, "NO VALUES FOUND"))
            sd2.AddFilter(ScheduleFilter(f2.FieldId, ScheduleFilterType.Equal, "ALL VALUES FOUND"))

        hdr2  = vs2.GetTableData().GetSectionData(SectionType.Header)
        body2 = vs2.GetTableData().GetSectionData(SectionType.Body)

        if existed2:
            clear_schedule_header(hdr2)

        try: sd2.ShowGridLines = True
        except Exception: pass

        try:
            from Autodesk.Revit.DB import TableCellStyle as _TCS2
            _bs2 = _TCS2()
            _bo2 = _bs2.GetCellStyleOverrideOptions()
            if _OFF_ID:
                _bo2.BorderTopLineStyle    = True;  _bs2.BorderTopLineStyle    = _OFF_ID
                _bo2.BorderBottomLineStyle = True;  _bs2.BorderBottomLineStyle = _OFF_ID
                _bo2.BorderLeftLineStyle   = True;  _bs2.BorderLeftLineStyle   = _OFF_ID
                _bo2.BorderRightLineStyle  = True;  _bs2.BorderRightLineStyle  = _OFF_ID
            else:
                _bo2.BorderTopLineStyle = False; _bo2.BorderBottomLineStyle = False
                _bo2.BorderLeftLineStyle = False; _bo2.BorderRightLineStyle = False
            _bs2.SetCellStyleOverrideOptions(_bo2)
            body2.SetCellStyle(_bs2)
        except Exception: pass

        _p2_group_ris, _p2_footer_ri, _p2_blank_ris, _p2_border_tasks = _render_page(hdr2, page_items, is_first_page=False)

        # Update footer page number
        hdr2_ref = vs2.GetTableData().GetSectionData(SectionType.Header)
        try:
            safe_text(hdr2_ref, _p2_footer_ri, TOTAL_COLS - 3,
                      'Page {} of {}'.format(page_idx, total_pages))
        except Exception: pass

        t2.Commit()

        # Border cleanup for overflow page
        with Transaction(doc, "pyTransmit p{} borders".format(page_idx)) as tb2:
            tb2.Start()
            hdr2_f = vs2.GetTableData().GetSectionData(SectionType.Header)
            if _style_algo and _sched_cell_styles:
                _s2_to_layout = {}; _s2_is_last = {}; _s2_data_b = {}
                for _bt2 in _p2_border_tasks:
                    _bk2 = _bt2[0]; _br2 = _bt2[1]; _be2 = _bt2[5]
                    _bl2 = _bt2[6] if len(_bt2) >= 8 else 0
                    _s2_to_layout[_br2] = _bl2
                    if _bk2 == "data" and _be2:
                        _s2_is_last[_br2] = _be2[1]
                        _s2_data_b[_br2]  = _be2[0] if isinstance(_be2[0], dict) else {'h':True,'v':True}
                for _sr2, _lr2 in _s2_to_layout.items():
                    _il2 = _s2_is_last.get(_sr2, True)
                    _db2 = _s2_data_b.get(_sr2, {})
                    _h2  = _db2.get('h', True); _v2 = _db2.get('v', True)
                    _id2 = _sr2 in _s2_is_last
                    for _lc2 in range(4):
                        _cs2 = _get_sched_style(_lr2, _lc2)
                        if not _cs2: continue
                        _t2 = _cs2.get('t',False); _b2 = _cs2.get('b',False) if (_il2 or not _id2) else _h2
                        _l2 = _cs2.get('l',False); _r2 = _cs2.get('r',False)
                        if _lc2 < REV_START:
                            _set_cell_border(hdr2_f, _sr2, _lc2, _t2, _b2, _l2, _r2)
                        else:
                            for _rc2 in range(MAX_REVS):
                                _sc2 = REV_START + _rc2
                                _set_cell_border(hdr2_f, _sr2, _sc2, _t2, _b2,
                                    _l2 if _rc2==0 else _v2, _r2 if _rc2==MAX_REVS-1 else _v2)
                for _br in _p2_blank_ris:
                    _above_lr2 = _s2_to_layout.get(_br - 1)
                    for _ci in range(TOTAL_COLS):
                        _lci_a2 = _ci if _ci < REV_START else REV_START
                        _cs_a2 = _get_sched_style(_above_lr2, _lci_a2) if _above_lr2 is not None else {}
                        _set_cell_border(hdr2_f, _br, _ci, _cs_a2.get('b', False), False, False, False)
                for _gr in _p2_group_ris:
                    for _ci in range(TOTAL_COLS):
                        _cs_g2  = _get_sched_style(_s2_to_layout.get(_gr, 0),
                                                             _ci if _ci < REV_START else REV_START)
                        _show_l_g2 = _cs_g2.get('l', False) if _ci == 0 else False
                        _show_r_g2 = _cs_g2.get('r', False) if _ci == TOTAL_COLS - 1 else False
                        _first_gri2 = min(_p2_group_ris) if _p2_group_ris else None
                        _top_g2 = False if _gr == _first_gri2 else _data_h
                        _set_cell_border(hdr2_f, _gr, _ci, _top_g2, _data_h, _show_l_g2, _show_r_g2)
            tb2.Commit()

        output.print_md("   Created `{}`  ({} rows)".format(sched_name, len(page_items)))


#  Logo insertion 
import os as _os
logo_path = LOGO_PATH
if not logo_path:
    _script_dir_logo = _os.path.dirname(_os.path.abspath(__file__))
    for _sd in (_os.path.join(_script_dir_logo, 'Settings'), _script_dir_logo):
        for _fn in ('logo.png','logo.PNG','logo.jpg','logo.JPG','logo.jpeg','logo.JPEG'):
            _cand = _os.path.join(_sd, _fn)
            if _os.path.exists(_cand):
                logo_path = _cand; break
        if logo_path: break

if logo_path and _os.path.exists(logo_path):
    try:
        with Transaction(doc, "pyTransmit logo") as tl:
            tl.Start()
            from Autodesk.Revit.DB import ImageType, ImageTypeOptions, ImageTypeSource
            existing_img = None
            for el in FilteredElementCollector(doc).OfClass(ImageType).ToElements():
                try:
                    if (getattr(el, 'Path', '') or '') == logo_path:
                        existing_img = el; break
                except: pass
            if existing_img is None:
                try:
                    opts = ImageTypeOptions(logo_path, False, ImageTypeSource.Import)
                    img_type = ImageType.Create(doc, opts)
                except Exception:
                    img_type = ImageType.Create(doc, logo_path)
            else:
                img_type = existing_img
            hdr_logo = sched.GetTableData().GetSectionData(SectionType.Header)
            logo_col = TOTAL_COLS - 3
            try:
                hdr_logo.InsertImage(0, logo_col, img_type.Id)
                output.print_md("   Logo inserted from: `{}`".format(logo_path))
            except Exception as img_err:
                output.print_md("   InsertImage error: {}".format(img_err))
            tl.Commit()
    except Exception as logo_err:
        output.print_md("   Logo insertion failed: {}".format(logo_err))


output.print_md("\n##  Complete — `{}`".format(SCHEDULE_NAME))
output.print_md("Sheets written: {}  |  Revisions: {}  |  Pages: {}".format(
    len(tx_sheets), n_revs, total_pages))
