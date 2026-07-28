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
pytable_excel.py -- everything specific to the Excel side of pyTable:
reading .xlsx named ranges/formatting directly from the zip (no COM),
building native Schedule/Legend/Drafting views from that data, and
the ExcelCardMixin providing the Excel-specific parts of the row/card
UI (mixed into PyTableWindow alongside WordCardMixin in PyTable.py).
"""

import sys as _sys
_sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pytable_shared import (
    hb, Row, VIEW_TYPES, SRC_COLOURS, STATUS_COLOURS,
    _run_export_script, _confirm, _alert,
)


VIEW_TYPE_SCHEDULE = 'Schedule View'
VIEW_TYPE_DRAFTING = 'Drafting View'
VIEW_TYPE_LEGEND   = 'Legend View'

MM     = 1.0 / 304.8   # millimetres to Revit internal feet
PREFIX = 'pyTable Table '


class TableRow(object):
    """Configuration for one import row from the UI."""

    def __init__(self):
        self.view_name   = ''
        self.named_range = ''
        self.sheet_name  = ''
        self.view_type   = VIEW_TYPE_SCHEDULE
        self.view_scale  = 1
        self.file_path   = ''
        self.auto_sync   = False
        self.font        = 'Arial'    # used for pyTable text type lookup
        self.size_hdr_mm = 2.5        # header row font size in mm
        self.size_dat_mm = 2.3        # data row font size in mm


# ── Read xlsx metadata (named ranges + sheet names) ──

def get_named_ranges_from_workbook(file_path):
    """
    Read named ranges and sheet names from an xlsx file.

    Returns:
        {
          'named_ranges': [str, ...],           # all range names
          'sheets':       [str, ...],           # all sheet names
          'sheet_ranges': {sheet: [name, ...]}, # ranges per sheet
        }

    Sheet assignment logic:
      1. If definedName has localSheetId attribute, it belongs to that
         sheet (0-based index into the sheets list).
      2. Otherwise parse the sheet name from the range reference formula,
         e.g. "Sheet1!$A$1" -> Sheet1.
      3. If neither applies (bare reference like "A1"), the range is
         treated as global and added to ALL sheets.
    """
    result = {'named_ranges': [], 'sheets': [], 'sheet_ranges': {}}

    try:
        with _zipfile.ZipFile(file_path, 'r') as z:
            xml_bytes = z.read('xl/workbook.xml')
    except Exception as e:
        # File couldn't even be opened/read as a zip - a genuinely
        # unexpected problem (locked file, wrong extension, corrupt
        # download), worth a real warning.
        logger.warning(
            'Could not read "{}" as an xlsx workbook: {}'.format(
                os.path.basename(file_path), e
            )
        )
        return result

    try:
        import clr
        clr.AddReference('System.Xml')
        from System.Xml import XmlDocument

        xml_doc = XmlDocument()
        xml_doc.LoadXml(xml_bytes.decode('utf-8'))
        ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

        # ── Sheets ──
        sheets = []
        sheet_nodes = xml_doc.GetElementsByTagName('sheet', ns)
        if sheet_nodes.Count == 0:
            sheet_nodes = xml_doc.GetElementsByTagName('sheet')
        for i in range(sheet_nodes.Count):
            a = sheet_nodes[i].Attributes.GetNamedItem('name')
            if a:
                sheets.append(a.Value)
        result['sheets'] = sheets

        # ── Named ranges ──
        sheet_ranges = {s: [] for s in sheets}
        named_ranges = []

        dn_nodes = xml_doc.GetElementsByTagName('definedName', ns)
        if dn_nodes.Count == 0:
            dn_nodes = xml_doc.GetElementsByTagName('definedName')

        for i in range(dn_nodes.Count):
            node = dn_nodes[i]
            name_attr = node.Attributes.GetNamedItem('name')
            if not name_attr or name_attr.Value.startswith('_xlnm'):
                continue
            name = name_attr.Value
            named_ranges.append(name)

            # Determine which sheet this range belongs to
            target_sheet = None

            # 1. localSheetId (sheet-scoped named range)
            local_attr = node.Attributes.GetNamedItem('localSheetId')
            if local_attr is not None:
                try:
                    idx = int(local_attr.Value)
                    if 0 <= idx < len(sheets):
                        target_sheet = sheets[idx]
                except Exception:
                    pass

            # 2. Parse sheet name from reference formula content
            #    e.g. "Sheet1!$A$1:$B$2" or "'My Sheet'!$A$1"
            if target_sheet is None:
                ref = (node.InnerText or '').strip()
                if '!' in ref:
                    sheet_part = ref.split('!')[0].strip("'").strip('"')
                    if sheet_part in sheets:
                        target_sheet = sheet_part

            # 3. Global range -> add to all sheets
            if target_sheet is None:
                for s in sheets:
                    sheet_ranges[s].append(name)
            else:
                sheet_ranges[target_sheet].append(name)

        result['named_ranges'] = named_ranges
        result['sheet_ranges'] = sheet_ranges

    except Exception:
        # workbook.xml existed but couldn't be parsed as XML (or had no
        # usable defined-name data) - this is the ordinary "file has no
        # named ranges" case, not a crash. Report it as such rather than
        # surfacing the raw XML parser exception as an ERROR.
        logger.debug(
            'No named ranges found in "{}" (workbook.xml unreadable or empty)'.format(
                os.path.basename(file_path)
            )
        )

    return result




# ── Read cell data from xlsx named range ──

def read_named_range_data(file_path, named_range, sheet_name):
    """
    Read rows from an Excel named range directly from the xlsx zip.
    Returns list of lists: [[header1, header2, ...], [val, val, ...], ...]
    First row is column headers.
    """
    try:
        import clr
        clr.AddReference('System.Xml')
        from System.Xml import XmlDocument

        with _zipfile.ZipFile(file_path, 'r') as z:
            names_in_zip = z.namelist()

            # Shared strings table
            shared_strings = _read_shared_strings(z, XmlDocument)

            # Workbook xml - resolve named range address
            wb_xml = XmlDocument()
            wb_xml.LoadXml(z.read('xl/workbook.xml').decode('utf-8'))

            range_ref = _resolve_named_range(wb_xml, named_range)

            # Determine target sheet
            target_sheet = sheet_name
            if range_ref and '!' in range_ref:
                target_sheet = range_ref.split('!')[0].strip("'")

            # Find worksheet file
            sheet_file = _find_sheet_file(z, wb_xml, target_sheet, XmlDocument)
            if not sheet_file or sheet_file not in names_in_zip:
                logger.error(
                    'Sheet file not found for: {}'.format(target_sheet)
                )
                return []

            ws_xml = XmlDocument()
            ws_xml.LoadXml(z.read(sheet_file).decode('utf-8'))

        # Parse cell bounds from range reference
        min_col, min_row, max_col, max_row = _parse_range_bounds(range_ref)

        # Extract and return rows
        return _extract_rows(
            ws_xml, shared_strings,
            min_col, min_row, max_col, max_row
        )

    except Exception as e:
        logger.error(
            'read_named_range_data failed for "{}": {}'.format(
                named_range, e
            )
        )
        return []

def _read_shared_strings(z, XmlDocument):
    shared_strings = []
    if 'xl/sharedStrings.xml' not in z.namelist():
        return shared_strings
    ss_xml = XmlDocument()
    ss_xml.LoadXml(z.read('xl/sharedStrings.xml').decode('utf-8'))
    si_nodes = ss_xml.GetElementsByTagName('si')
    if si_nodes.Count == 0:
        si_nodes = ss_xml.GetElementsByTagName(
            'si',
            'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        )
    for i in range(si_nodes.Count):
        t_nodes = si_nodes[i].GetElementsByTagName('t')
        if t_nodes.Count == 0:
            t_nodes = si_nodes[i].GetElementsByTagName(
                't',
                'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
            )
        text = ''.join(
            t_nodes[j].InnerText for j in range(t_nodes.Count)
        )
        shared_strings.append(text)
    return shared_strings

def _resolve_named_range(wb_xml, named_range):
    dn_nodes = wb_xml.GetElementsByTagName('definedName')
    if dn_nodes.Count == 0:
        dn_nodes = wb_xml.GetElementsByTagName(
            'definedName',
            'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        )
    for i in range(dn_nodes.Count):
        attr = dn_nodes[i].Attributes.GetNamedItem('name')
        if attr and attr.Value == named_range:
            return dn_nodes[i].InnerText
    return None

def _find_sheet_file(z, wb_xml, sheet_name, XmlDocument):
    sheet_nodes = wb_xml.GetElementsByTagName('sheet')
    if sheet_nodes.Count == 0:
        sheet_nodes = wb_xml.GetElementsByTagName(
            'sheet',
            'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        )

    r_id = None
    for i in range(sheet_nodes.Count):
        name_attr = sheet_nodes[i].Attributes.GetNamedItem('name')
        if name_attr and name_attr.Value == sheet_name:
            id_attr = (
                sheet_nodes[i].Attributes.GetNamedItem('r:id') or
                sheet_nodes[i].Attributes.GetNamedItem('id')
            )
            if id_attr:
                r_id = id_attr.Value
            break

    if not r_id:
        return None

    rels_path = 'xl/_rels/workbook.xml.rels'
    if rels_path not in z.namelist():
        return None

    rels_xml = XmlDocument()
    rels_xml.LoadXml(z.read(rels_path).decode('utf-8'))
    rel_nodes = rels_xml.GetElementsByTagName('Relationship')
    for i in range(rel_nodes.Count):
        id_attr = rel_nodes[i].Attributes.GetNamedItem('Id')
        if id_attr and id_attr.Value == r_id:
            target = rel_nodes[i].Attributes.GetNamedItem('Target').Value
            if not target.startswith('xl/'):
                target = 'xl/' + target
            return target
    return None

def _col_letter_to_index(col_str):
    result = 0
    for ch in col_str.upper():
        result = result * 26 + (ord(ch) - ord('A') + 1)
    return result - 1

def _parse_range_bounds(range_ref):
    try:
        if not range_ref:
            return 0, 0, None, None
        ref = range_ref.split('!')[-1].replace('$', '')
        start, end = (ref.split(':') + [ref])[:2]
        m1 = re.match(r'([A-Za-z]+)(\d+)', start)
        m2 = re.match(r'([A-Za-z]+)(\d+)', end)
        return (
            _col_letter_to_index(m1.group(1)), int(m1.group(2)) - 1,
            _col_letter_to_index(m2.group(1)), int(m2.group(2)) - 1
        )
    except Exception:
        return 0, 0, None, None

def read_range_formatting(file_path, named_range, sheet_name):
    """
    Read cell formatting from an xlsx named range.
    Returns a dict:
    {
        'cell_styles': {(row_idx, col_idx): {
            'font_name': str,
            'font_size': float,  # in points
            'bold': bool,
            'italic': bool,
            'halign': str,       # 'Left', 'Center', 'Right'
            'fill_rgb': tuple or None,  # (r, g, b)
            'border_top': str,
            'border_bottom': str,
            'border_left': str,
            'border_right': str,
        }},
        'merges': [(row_start, col_start, row_end, col_end), ...],
        'row_heights': {row_idx: float},  # in points
        'col_widths': {col_idx: float},   # in mm
    }
    Row/col indices are relative to the named range (0-based).
    """
    result = {
        'cell_styles': {},
        'merges':      [],
        'row_heights': {},
        'col_widths':  {},
    }

    try:
        import clr
        clr.AddReference('System.Xml')
        from System.Xml import XmlDocument

        with _zipfile.ZipFile(file_path, 'r') as z:
            names = z.namelist()

            # --- theme colours ---
            # Read xl/theme/theme1.xml and build cache before parsing fonts/fills
            theme_colours = {}
            theme_files = [n for n in names if 'theme' in n and n.endswith('.xml')]
            if theme_files:
                try:
                    theme_xml = XmlDocument()
                    theme_xml.LoadXml(z.read(theme_files[0]).decode('utf-8'))
                    # Theme XML uses a: namespace prefix throughout:
                    # <a:dk1>, <a:srgbClr>, <a:sysClr> etc.
                    # Must use the drawingml namespace URI for lookups.
                    THEME_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

                    def _theme_tag(tag):
                        # Try with namespace first, fallback without
                        n = theme_xml.GetElementsByTagName(tag, THEME_NS)
                        if n.Count == 0:
                            n = theme_xml.GetElementsByTagName(tag)
                        return n

                    def _child_tag(node, tag):
                        n = node.GetElementsByTagName(tag, THEME_NS)
                        if n.Count == 0:
                            n = node.GetElementsByTagName(tag)
                        return n

                    # Build ordered list matching Excel theme index.
                    # NOTE: ECMA-376 quirk — styles.xml theme indices swap
                    # the first two pairs relative to clrScheme element order:
                    # 0=lt1, 1=dk1, 2=lt2, 3=dk2, 4=accent1 ... 9=accent6
                    all_colour_els = []
                    for tag in ['lt1','dk1','lt2','dk2',
                                'accent1','accent2','accent3',
                                'accent4','accent5','accent6',
                                'hlink','folHlink']:
                        tag_nodes = _theme_tag(tag)
                        if tag_nodes.Count > 0:
                            child = tag_nodes[0]
                            srgb = _child_tag(child, 'srgbClr')
                            sys  = _child_tag(child, 'sysClr')
                            if srgb.Count > 0:
                                attr = srgb[0].Attributes.GetNamedItem('val')
                                if attr and len(attr.Value) == 6:
                                    h = attr.Value
                                    all_colour_els.append((
                                        int(h[0:2],16),
                                        int(h[2:4],16),
                                        int(h[4:6],16)
                                    ))
                                else:
                                    all_colour_els.append(None)
                            elif sys.Count > 0:
                                attr = sys[0].Attributes.GetNamedItem('lastClr')
                                if attr and len(attr.Value) == 6:
                                    h = attr.Value
                                    all_colour_els.append((
                                        int(h[0:2],16),
                                        int(h[2:4],16),
                                        int(h[4:6],16)
                                    ))
                                else:
                                    all_colour_els.append(None)
                            else:
                                all_colour_els.append(None)

                    theme_colours = {
                        i: c for i, c in enumerate(all_colour_els)
                        if c is not None
                    }
                    _set_theme_cache(theme_colours)
                except Exception as te:
                    logger.debug('Theme parse failed: {}'.format(te))

            # --- styles ---
            fonts, fills, borders, xfs = [], [], [], []
            if 'xl/styles.xml' in names:
                styles_xml = XmlDocument()
                styles_xml.LoadXml(z.read('xl/styles.xml').decode('utf-8'))
                fonts  = _parse_fonts(styles_xml)
                fills  = _parse_fills(styles_xml)
                borders= _parse_borders(styles_xml)
                xfs    = _parse_xfs(styles_xml)
            # --- workbook for range ref ---
            wb_xml = XmlDocument()
            wb_xml.LoadXml(z.read('xl/workbook.xml').decode('utf-8'))
            range_ref = _resolve_named_range(wb_xml, named_range)

            target_sheet = sheet_name
            if range_ref and '!' in range_ref:
                target_sheet = range_ref.split('!')[0].strip("'")

            sheet_file = _find_sheet_file(z, wb_xml, target_sheet, XmlDocument)
            if not sheet_file or sheet_file not in names:
                return result

            ws_xml = XmlDocument()
            ws_xml.LoadXml(z.read(sheet_file).decode('utf-8'))

        min_col, min_row, max_col, max_row = _parse_range_bounds(range_ref)

        # --- parse merges ---
        merge_nodes = ws_xml.GetElementsByTagName('mergeCell')
        for i in range(merge_nodes.Count):
            ref = merge_nodes[i].Attributes.GetNamedItem('ref')
            if not ref:
                continue
            parts = ref.Value.replace('$','').split(':')
            if len(parts) != 2:
                continue
            m1 = re.match(r'([A-Za-z]+)(\d+)', parts[0])
            m2 = re.match(r'([A-Za-z]+)(\d+)', parts[1])
            if not m1 or not m2:
                continue
            r1 = int(m1.group(2)) - 1 - min_row
            c1 = _col_letter_to_index(m1.group(1)) - min_col
            r2 = int(m2.group(2)) - 1 - min_row
            c2 = _col_letter_to_index(m2.group(1)) - min_col
            if r1 >= 0 and c1 >= 0:
                result['merges'].append((r1, c1, r2, c2))

        # --- parse row heights ---
        row_nodes = ws_xml.GetElementsByTagName('row')
        for i in range(row_nodes.Count):
            r_attr = row_nodes[i].Attributes.GetNamedItem('r')
            ht_attr = row_nodes[i].Attributes.GetNamedItem('ht')
            if not r_attr:
                continue
            row_idx = int(r_attr.Value) - 1
            if row_idx < min_row:
                continue
            if max_row is not None and row_idx > max_row:
                break
            rel_row = row_idx - min_row
            if ht_attr and ht_attr.Value:
                try:
                    result['row_heights'][rel_row] = float(ht_attr.Value)
                except Exception:
                    pass

        # --- parse cell styles ---
        row_nodes2 = ws_xml.GetElementsByTagName('row')
        for i in range(row_nodes2.Count):
            r_attr = row_nodes2[i].Attributes.GetNamedItem('r')
            if not r_attr:
                continue
            row_idx = int(r_attr.Value) - 1
            if row_idx < min_row:
                continue
            if max_row is not None and row_idx > max_row:
                break
            rel_row = row_idx - min_row

            c_nodes = row_nodes2[i].GetElementsByTagName('c')
            for j in range(c_nodes.Count):
                ref = c_nodes[j].Attributes.GetNamedItem('r')
                s_attr = c_nodes[j].Attributes.GetNamedItem('s')
                if not ref:
                    continue
                m = re.match(r'([A-Za-z]+)(\d+)', ref.Value)
                if not m:
                    continue
                col_idx = _col_letter_to_index(m.group(1))
                if col_idx < min_col:
                    continue
                if max_col is not None and col_idx > max_col:
                    continue
                rel_col = col_idx - min_col

                xf_idx = int(s_attr.Value) if s_attr else 0
                xf = xfs[xf_idx] if xf_idx < len(xfs) else {}

                font_idx   = xf.get('fontId',   0)
                fill_idx   = xf.get('fillId',   0)
                border_idx = xf.get('borderId', 0)
                halign     = xf.get('align', '') or 'Left'
                halign     = halign.capitalize() if halign else 'Left'

                font   = fonts[font_idx]     if font_idx   < len(fonts)   else {}
                fill   = fills[fill_idx]     if fill_idx   < len(fills)   else None
                border = borders[border_idx] if border_idx < len(borders) else {}

                result['cell_styles'][(rel_row, rel_col)] = {
                    'font_name':          font.get('name', 'Arial'),
                    'font_size':          float(font.get('size', 11)),
                    'bold':               font.get('bold', False),
                    'italic':             font.get('italic', False),
                    'underline':          font.get('underline', False),
                    'color_rgb':          font.get('color_rgb', None),
                    'halign':             halign,
                    'fill_rgb':           fill if fill else None,
                    'border_top':         border.get('top', ''),
                    'border_top_color':   border.get('top_color', None),
                    'border_bottom':      border.get('bottom', ''),
                    'border_bottom_color':border.get('bottom_color', None),
                    'border_left':        border.get('left', ''),
                    'border_left_color':  border.get('left_color', None),
                    'border_right':       border.get('right', ''),
                    'border_right_color': border.get('right_color', None),
                }

        # --- parse column widths ---
        # Read default column width from sheetFormatPr
        default_col_w_ch = 8.0  # Excel character units
        fmt_nodes = ws_xml.GetElementsByTagName('sheetFormatPr')
        if fmt_nodes.Count > 0:
            dcw = fmt_nodes[0].Attributes.GetNamedItem('defaultColWidth')
            drh = fmt_nodes[0].Attributes.GetNamedItem('defaultRowHeight')
            if dcw:
                try:
                    default_col_w_ch = float(dcw.Value)
                except Exception:
                    pass
            if drh:
                try:
                    result['default_row_height'] = float(drh.Value)
                except Exception:
                    pass

        # Max Digit Width depends on the workbook default font:
        # Aptos Narrow (Excel 2024+): 7.41px
        # Calibri Light: 6.8px
        # Calibri / Arial / most others: 7.0px
        _default_font_name = (fonts[0].get('name', '') if fonts else '').lower()
        # MDW calibrated from physical measurement:
        # Excel 10 char units -> 22.93mm measured in Revit -> MDW=8.17
        # This accounts for Aptos Narrow metrics + Revit's internal padding
        if 'aptos' in _default_font_name:
            _col_mdw = 8.17
        elif 'calibri light' in _default_font_name:
            _col_mdw = 7.4
        else:
            _col_mdw = 7.6   # Calibri, Arial and others (approx)

        # Read explicit column widths for cols in range
        col_nodes = ws_xml.GetElementsByTagName('col')
        for i in range(col_nodes.Count):
            mn_attr = col_nodes[i].Attributes.GetNamedItem('min')
            mx_attr = col_nodes[i].Attributes.GetNamedItem('max')
            w_attr  = col_nodes[i].Attributes.GetNamedItem('width')
            if not mn_attr or not w_attr:
                continue
            try:
                mn = int(mn_attr.Value) - 1  # 0-based
                mx = int(mx_attr.Value) - 1 if mx_attr else mn
                w_ch = float(w_attr.Value)
                for ci in range(mn, mx + 1):
                    rel_ci = ci - min_col
                    if 0 <= rel_ci <= (max_col - min_col if max_col is not None else 999):
                        # OOXML formula: px = int((chars*MDW+5)/MDW*256)/256*MDW
                        # MDW=7.41 for Aptos Narrow, 7.0 for Calibri/Arial
                        _px = int((w_ch * _col_mdw + 5) / _col_mdw * 256) / 256.0 * _col_mdw
                        result['col_widths'][rel_ci] = _px * 25.4 / 96.0
            except Exception:
                pass

        # Fill any missing col widths with default
        n_range_cols = (max_col - min_col + 1) if max_col is not None else 0
        for ci in range(n_range_cols):
            if ci not in result['col_widths']:
                _px = int((default_col_w_ch * _col_mdw + 5) / _col_mdw * 256) / 256.0 * _col_mdw
                result['col_widths'][ci] = _px * 25.4 / 96.0

    except Exception as e:
        logger.error('read_range_formatting failed: {}'.format(e))

    return result

def _resolve_colour(xml_doc, color_node, skip_white=True):
    """
    Resolve an Excel colour node to an (r, g, b) tuple or None.

    Priority:
    1. Explicit rgb attribute (ARGB hex, e.g. FFFF0000)
    2. theme attribute → look up in theme cache
    3. indexed attribute → standard Excel 56-colour palette
    4. auto / missing → None (use Revit default)

    skip_white: if True, returns None for white/near-white colours.
    White text is only valid on dark backgrounds — without a matching
    dark fill it would be invisible in Revit.
    """
    if color_node is None:
        return None

    auto = color_node.Attributes.GetNamedItem('auto')
    if auto and auto.Value == '1':
        return None

    rgb = None

    # --- explicit RGB ---
    rgb_attr = color_node.Attributes.GetNamedItem('rgb')
    if rgb_attr and rgb_attr.Value and len(rgb_attr.Value) == 8:
        hex_c = rgb_attr.Value[2:]
        try:
            rgb = (
                int(hex_c[0:2], 16),
                int(hex_c[2:4], 16),
                int(hex_c[4:6], 16)
            )
        except Exception:
            pass

    # --- theme colour ---
    if rgb is None:
        theme_attr = color_node.Attributes.GetNamedItem('theme')
        if theme_attr:
            try:
                base = _get_theme_colour(xml_doc, int(theme_attr.Value))
                if base:
                    rgb = base
            except Exception:
                pass

    # --- indexed colour ---
    if rgb is None:
        indexed_attr = color_node.Attributes.GetNamedItem('indexed')
        if indexed_attr:
            try:
                base = _EXCEL_INDEXED_COLOURS.get(int(indexed_attr.Value))
                if base:
                    rgb = base
            except Exception:
                pass

    if rgb is None:
        return None

    # Apply tint
    rgb = _apply_tint(rgb, color_node)

    # Skip white/near-white — invisible without a dark background fill
    if skip_white and rgb[0] > 230 and rgb[1] > 230 and rgb[2] > 230:
        return None

    return rgb

def _apply_tint(rgb, color_node):
    """Apply Excel tint (-1 to 1) to an RGB tuple. Returns modified RGB."""
    if color_node is None:
        return rgb
    tint_attr = color_node.Attributes.GetNamedItem('tint')
    if not tint_attr:
        return rgb
    try:
        tint = float(tint_attr.Value)
        if tint == 0.0:
            return rgb
        r, g, b = rgb
        if tint > 0:
            # lighten toward white
            r = int(r + (255 - r) * tint)
            g = int(g + (255 - g) * tint)
            b = int(b + (255 - b) * tint)
        else:
            # darken toward black
            r = int(r * (1 + tint))
            g = int(g * (1 + tint))
            b = int(b * (1 + tint))
        return (
            max(0, min(255, r)),
            max(0, min(255, g)),
            max(0, min(255, b))
        )
    except Exception:
        return rgb


# Cache for theme colours so we only parse once per workbook read
_theme_colour_cache = {}

def _get_theme_colour(xml_doc, idx):
    """
    Return (r, g, b) for a theme colour index.
    Uses the cache populated by _set_theme_cache during read_range_formatting.
    """
    global _theme_colour_cache
    return _theme_colour_cache.get(idx)

def _set_theme_cache(colours):
    """Called from read_range_formatting with resolved theme colours."""
    global _theme_colour_cache
    _theme_colour_cache = colours


# Standard Excel 56-colour indexed palette (indices 0-55)
_EXCEL_INDEXED_COLOURS = {
    0:  (0,   0,   0),    # black
    1:  (255, 255, 255),  # white
    2:  (255, 0,   0),    # red
    3:  (0,   255, 0),    # green
    4:  (0,   0,   255),  # blue
    5:  (255, 255, 0),    # yellow
    6:  (255, 0,   255),  # magenta
    7:  (0,   255, 255),  # cyan
    8:  (0,   0,   0),
    9:  (255, 255, 255),
    10: (255, 0,   0),
    11: (0,   255, 0),
    12: (0,   0,   255),
    13: (255, 255, 0),
    14: (255, 0,   255),
    15: (0,   255, 255),
    16: (128, 0,   0),
    17: (0,   128, 0),
    18: (0,   0,   128),
    19: (128, 128, 0),
    20: (128, 0,   128),
    21: (0,   128, 128),
    22: (192, 192, 192),
    23: (128, 128, 128),
    24: (153, 153, 255),
    25: (153, 51,  102),
    26: (255, 255, 204),
    27: (204, 255, 255),
    28: (102, 0,   102),
    29: (255, 128, 128),
    30: (0,   102, 204),
    31: (204, 204, 255),
    32: (0,   0,   128),
    33: (255, 0,   255),
    34: (255, 255, 0),
    35: (0,   255, 255),
    36: (128, 0,   128),
    37: (128, 0,   0),
    38: (0,   128, 128),
    39: (0,   0,   255),
    40: (0,   204, 255),
    41: (204, 255, 255),
    42: (204, 255, 204),
    43: (255, 255, 153),
    44: (153, 204, 255),
    45: (255, 153, 204),
    46: (204, 153, 255),
    47: (255, 204, 153),
    48: (51,  102, 255),
    49: (51,  204, 204),
    50: (153, 204, 0),
    51: (255, 204, 0),
    52: (255, 153, 0),
    53: (255, 102, 0),
    54: (102, 102, 153),
    55: (150, 150, 150),
    64: (0,   0,   0),    # system foreground
    65: (255, 255, 255),  # system background
}

def _parse_fonts(xml_doc):
    """
    Parse font definitions from styles XmlDocument.
    Resolves theme and indexed colours using _resolve_colour.

    Namespace handling: styles.xml always carries the spreadsheetml default
    namespace.  System.Xml GetElementsByTagName without a namespace URI only
    matches elements that have NO namespace, so it returns nothing on a
    real xlsx file.  We try without namespace first (covers edge cases) then
    fall back to the full URI so at least one call succeeds.

    Bold detection: <b/> means bold=True.  <b val="1"/> also means bold=True.
    <b val="0"/> explicitly means bold=False (override inheritance).
    Checking Count > 0 alone would misread <b val="0"/> as bold.
    """
    NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    fonts = []
    nodes = xml_doc.GetElementsByTagName('font')
    if nodes.Count == 0:
        nodes = xml_doc.GetElementsByTagName('font', NS)
    for i in range(nodes.Count):
        node = nodes[i]
        f = {}

        def _child(tag):
            n = node.GetElementsByTagName(tag)
            if n.Count == 0:
                n = node.GetElementsByTagName(tag, NS)
            return n

        b     = _child('b')
        it    = _child('i')
        sz    = _child('sz')
        name  = _child('name')
        color = _child('color')

        # <b/> or <b val="1"/> = bold.  <b val="0"/> = explicitly NOT bold.
        if b.Count > 0:
            val_attr = b[0].Attributes.GetNamedItem('val')
            f['bold'] = (val_attr is None or val_attr.Value not in ('0', 'false'))
        else:
            f['bold'] = False

        # Same pattern for italic
        if it.Count > 0:
            val_attr = it[0].Attributes.GetNamedItem('val')
            f['italic'] = (val_attr is None or val_attr.Value not in ('0', 'false'))
        else:
            f['italic'] = False

        f['size']   = (float(sz[0].Attributes.GetNamedItem('val').Value)
                       if sz.Count > 0 else 11.0)
        f['name']   = (name[0].Attributes.GetNamedItem('val').Value
                       if name.Count > 0 else 'Arial')
        # skip_white=False: white text is valid on dark/coloured fills
        # (e.g. white text on grey cell, white text on Accent1 blue)
        f['color_rgb'] = _resolve_colour(
            xml_doc,
            color[0] if color.Count > 0 else None,
            skip_white=False
        )
        # Underline: <u/> or <u val="single"/> = underline on
        #            <u val="none"/> = underline off (explicit)
        u = _child('u')
        if u.Count > 0:
            val_attr = u[0].Attributes.GetNamedItem('val')
            f['underline'] = (val_attr is None or
                              val_attr.Value not in ('none', 'false'))
        else:
            f['underline'] = False
        fonts.append(f)
    return fonts

def _parse_fills(xml_doc):
    """
    Parse fill definitions from styles XmlDocument.
    Does not skip white fills — white can be an explicit background.

    Namespace handling: same fallback pattern as _parse_fonts.
    The first two fill entries in OOXML are always the system 'none' and
    'gray125' fills — we preserve their index positions as None so that
    fillId references from xf records remain correctly indexed.
    """
    NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    fills = []
    nodes = xml_doc.GetElementsByTagName('fill')
    if nodes.Count == 0:
        nodes = xml_doc.GetElementsByTagName('fill', NS)
    for i in range(nodes.Count):
        node = nodes[i]

        def _child_f(tag):
            n = node.GetElementsByTagName(tag)
            if n.Count == 0:
                n = node.GetElementsByTagName(tag, NS)
            return n

        pf_nodes = _child_f('patternFill')
        pattern = ''
        if pf_nodes.Count > 0:
            pt = pf_nodes[0].Attributes.GetNamedItem('patternType')
            pattern = pt.Value if pt else ''

        # 'none' and 'gray125' are the two reserved system fills
        if pattern in ('none', 'gray125', ''):
            fills.append(None)
            continue

        fg = _child_f('fgColor')
        rgb = _resolve_colour(
            xml_doc,
            fg[0] if fg.Count > 0 else None,
            skip_white=False  # fills can legitimately be white
        )
        fills.append(rgb)
    return fills

def _parse_borders(xml_doc):
    """
    Parse border definitions from styles XmlDocument.
    Returns list of dicts with style and colour per side.
    Border colours use skip_white=False since dark borders on
    white backgrounds are common.

    Namespace handling: same fallback pattern as _parse_fonts.
    """
    NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    borders = []
    nodes = xml_doc.GetElementsByTagName('border')
    if nodes.Count == 0:
        nodes = xml_doc.GetElementsByTagName('border', NS)
    for i in range(nodes.Count):
        node = nodes[i]
        b = {}
        for side in ('left', 'right', 'top', 'bottom'):
            side_nodes = node.GetElementsByTagName(side)
            if side_nodes.Count == 0:
                side_nodes = node.GetElementsByTagName(side, NS)
            if side_nodes.Count > 0:
                style_attr = side_nodes[0].Attributes.GetNamedItem('style')
                b[side] = style_attr.Value if style_attr else ''
                color_node = side_nodes[0].GetElementsByTagName('color')
                if color_node.Count == 0:
                    color_node = side_nodes[0].GetElementsByTagName('color', NS)
                b[side + '_color'] = _resolve_colour(
                    xml_doc,
                    color_node[0] if color_node.Count > 0 else None,
                    skip_white=False
                )
            else:
                b[side] = ''
                b[side + '_color'] = None
        borders.append(b)
    return borders

def _parse_xfs(xml_doc):
    """
    Parse cellXfs (per-cell format table) from styles XmlDocument.

    OOXML has two xf tables:
      <cellStyleXfs>  named styles (Normal, Heading 1 ...) — fills live here
      <cellXfs>       per-cell overrides — each entry has an xfId pointing
                      back to its parent style in cellStyleXfs

    Excel stores fill, font and border on the PARENT style in cellStyleXfs
    and only writes the override in cellXfs when the cell explicitly changes
    that property.  So to get the real fillId for a cell we must:
      1. Read the xfId from the cellXfs entry
      2. Look up that index in cellStyleXfs
      3. Use the parent fillId when the cellXfs fillId is 0 (unset)

    The applyFill / applyFont flags tell us which properties the cellXfs
    entry is actually overriding vs inheriting from the parent style.
    When applyFill is absent or 0, the fill comes from the parent style.
    """
    NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

    def _get_xf_nodes(tag):
        nodes = xml_doc.GetElementsByTagName(tag)
        if nodes.Count == 0:
            nodes = xml_doc.GetElementsByTagName(tag, NS)
        return nodes

    def _int_attr_node(node, name, default=0):
        a = node.Attributes.GetNamedItem(name)
        try:
            return int(a.Value) if a else default
        except Exception:
            return default

    def _bool_attr(node, name):
        """Return True when attribute is '1' or absent (absent = inherit)."""
        a = node.Attributes.GetNamedItem(name)
        if a is None:
            return None   # absent = not set by this record, inherit from parent
        return a.Value not in ('0', 'false')

    # --- parse cellStyleXfs (parent / named styles) ---
    style_xfs = []
    sxfs_nodes = _get_xf_nodes('cellStyleXfs')
    if sxfs_nodes.Count > 0:
        for i in range(sxfs_nodes[0].ChildNodes.Count):
            node = sxfs_nodes[0].ChildNodes[i]
            try:
                if node.Attributes is None:
                    continue
            except Exception:
                continue
            style_xfs.append({
                'fontId':   _int_attr_node(node, 'fontId'),
                'fillId':   _int_attr_node(node, 'fillId'),
                'borderId': _int_attr_node(node, 'borderId'),
            })

    # --- parse cellXfs (per-cell formats) ---
    # IMPORTANT: use ChildNodes not GetElementsByTagName.
    # GetElementsByTagName('xf') on the cellXfs node also matches <xf>
    # elements nested inside cellStyleXfs when namespace handling is
    # inconsistent, returning fewer entries than actually exist.
    # ChildNodes iterates only the direct children of cellXfs, which
    # are exactly the 26 (or N) <xf> elements we want.
    xfs = []
    xfs_nodes = _get_xf_nodes('cellXfs')
    if xfs_nodes.Count == 0:
        return xfs
    xf_nodes = xfs_nodes[0].ChildNodes

    for i in range(xf_nodes.Count):
        node = xf_nodes[i]
        # ChildNodes includes text/whitespace nodes — skip non-element nodes.
        # In IronPython, NodeType is an enum so comparing to int 1 always fails.
        # Use Attributes as a proxy: only element nodes have an Attributes collection.
        try:
            if node.Attributes is None:
                continue
        except Exception:
            continue

        # xfId points to the parent style in cellStyleXfs
        xf_id = _int_attr_node(node, 'xfId', 0)
        parent = style_xfs[xf_id] if xf_id < len(style_xfs) else {}

        # applyFill/applyFont/applyBorder:
        #   None (absent) = inherit from parent style
        #   True ('1')    = this cellXfs entry overrides
        #   False ('0')   = explicitly NOT overriding (use parent)
        apply_fill   = _bool_attr(node, 'applyFill')
        apply_font   = _bool_attr(node, 'applyFont')
        apply_border = _bool_attr(node, 'applyBorder')

        cell_fill_id   = _int_attr_node(node, 'fillId')
        cell_font_id   = _int_attr_node(node, 'fontId')
        cell_border_id = _int_attr_node(node, 'borderId')

        # Use parent value when cell is not overriding that property
        if apply_fill is False or (apply_fill is None and cell_fill_id == 0):
            fill_id = parent.get('fillId', 0)
        else:
            fill_id = cell_fill_id

        if apply_font is False or (apply_font is None and cell_font_id == 0):
            font_id = parent.get('fontId', 0)
        else:
            font_id = cell_font_id

        if apply_border is False or (apply_border is None and cell_border_id == 0):
            border_id = parent.get('borderId', 0)
        else:
            border_id = cell_border_id

        xf = {
            'fontId':   font_id,
            'fillId':   fill_id,
            'borderId': border_id,
            'align':    '',
        }

        align_nodes = node.GetElementsByTagName('alignment')
        if align_nodes.Count == 0:
            align_nodes = node.GetElementsByTagName('alignment', NS)
        if align_nodes.Count > 0:
            h = align_nodes[0].Attributes.GetNamedItem('horizontal')
            xf['align'] = h.Value if h else ''

        xfs.append(xf)
    return xfs

def _extract_rows(ws_xml, shared_strings,
                  min_col, min_row, max_col, max_row):
    row_nodes = ws_xml.GetElementsByTagName('row')
    if row_nodes.Count == 0:
        row_nodes = ws_xml.GetElementsByTagName(
            'row',
            'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        )
    result = []
    for ri in range(row_nodes.Count):
        row_node = row_nodes[ri]
        r_attr = row_node.Attributes.GetNamedItem('r')
        if not r_attr:
            continue
        row_idx = int(r_attr.Value) - 1
        if row_idx < min_row:
            continue
        if max_row is not None and row_idx > max_row:
            break

        cell_nodes = row_node.GetElementsByTagName('c')
        if cell_nodes.Count == 0:
            cell_nodes = row_node.GetElementsByTagName(
                'c',
                'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
            )

        col_vals = {}
        for ci in range(cell_nodes.Count):
            cell = cell_nodes[ci]
            ref = cell.Attributes.GetNamedItem('r')
            if not ref:
                continue
            m = re.match(r'([A-Za-z]+)(\d+)', ref.Value)
            if not m:
                continue
            col_idx = _col_letter_to_index(m.group(1))
            if col_idx < min_col:
                continue
            if max_col is not None and col_idx > max_col:
                continue

            t_attr = cell.Attributes.GetNamedItem('t')
            v_nodes = cell.GetElementsByTagName('v')
            if v_nodes.Count == 0:
                v_nodes = cell.GetElementsByTagName(
                    'v',
                    'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
                )
            if v_nodes.Count == 0:
                col_vals[col_idx] = ''
                continue

            raw = v_nodes[0].InnerText
            if t_attr and t_attr.Value == 's':
                idx = int(raw)
                col_vals[col_idx] = (
                    shared_strings[idx]
                    if idx < len(shared_strings) else ''
                )
            else:
                col_vals[col_idx] = raw

        if not col_vals:
            continue

        end_col = (
            max_col if max_col is not None
            else max(col_vals.keys())
        )
        result.append([
            col_vals.get(c, '')
            for c in range(min_col, end_col + 1)
        ])

    return result


# ── Schedule creation - pyTransmit pattern ──
# Uses ViewSchedule.CreateSchedule with ElementId.InvalidElementId (no
# category) and builds the entire table in the Header section.
# No Key Schedule, no project parameters required.

MM = 1.0 / 304.8   # millimetres to Revit internal feet

# ── pyTable text type manager ──
# Text types are named "pyTable Table XX" and matched by font/size/bold.
# If an existing type matches the fingerprint it is reused.
# If not, a new one is created with the next available number.

PREFIX = 'pyTable Table '

def _text_type_fingerprint(font, size_mm, bold):
    """
    Canonical string key for a text style.
    Used to match existing types without creating duplicates.
    """
    return '{}__{:.4f}__{}'.format(
        font.lower().strip(), size_mm, 'bold' if bold else 'regular'
    )

def _read_fingerprint(tt):
    """
    Read font, size and bold from an existing TextNoteType and return
    its fingerprint string. Returns None if the type cannot be read.
    """
    try:
        font_p = tt.get_Parameter(DB.BuiltInParameter.TEXT_FONT)
        size_p = tt.get_Parameter(DB.BuiltInParameter.TEXT_SIZE)
        bold_p = tt.get_Parameter(DB.BuiltInParameter.TEXT_STYLE_BOLD)

        if not font_p or not size_p:
            return None

        font    = font_p.AsString() or 'Arial'
        size_ft = size_p.AsDouble()
        size_mm = size_ft / MM  # convert feet back to mm
        bold    = bool(bold_p.AsInteger()) if bold_p else False

        return _text_type_fingerprint(font, size_mm, bold)
    except Exception:
        return None

def get_or_create_text_type(font='Arial', size_mm=2.3, bold=False):
    """
    Return a TextNoteType matching font/size/bold.
    Searches existing 'pyTable Table XX' types first.
    Creates a new one if none match.

    All pyTable types are named 'pyTable Table 01', '02', etc.
    so they stay grouped in the Type Selector and are easy to manage.
    """
    target_fp = _text_type_fingerprint(font, size_mm, bold)

    all_tt = list(
        DB.FilteredElementCollector(doc)
        .OfClass(DB.TextNoteType)
        .ToElements()
    )

    # Collect all existing pyTable types and check for a match
    pytable_types = []
    for tt in all_tt:
        try:
            name = tt.get_Parameter(
                DB.BuiltInParameter.SYMBOL_NAME_PARAM
            ).AsString()
        except Exception:
            continue

        if name and name.startswith(PREFIX):
            pytable_types.append((name, tt))
            fp = _read_fingerprint(tt)
            if fp == target_fp:
                logger.debug(
                    'Reusing text type "{}": {}'.format(name, target_fp)
                )
                return tt

    # No match — create a new one
    next_num = len(pytable_types) + 1
    new_name = '{}{:02d}'.format(PREFIX, next_num)

    # Duplicate from any existing TextNoteType as a base
    base = all_tt[0] if all_tt else None
    if not base:
        logger.error('No TextNoteType found to duplicate')
        return None

    size_ft = size_mm * MM

    with revit.Transaction(
        'pyTable - Create text type: {}'.format(new_name)
    ):
        new_tt = base.Duplicate(new_name)
        for bip, val in [
            (DB.BuiltInParameter.TEXT_FONT,         font),
            (DB.BuiltInParameter.TEXT_SIZE,          size_ft),
            (DB.BuiltInParameter.TEXT_STYLE_BOLD,    1 if bold else 0),
            (DB.BuiltInParameter.TEXT_STYLE_ITALIC,  0),
            (DB.BuiltInParameter.TEXT_BACKGROUND,    1),  # opaque
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
    Uses a pre-resolved pyTable TextNoteType ID (tt_id) when provided
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

        # Resolve font name from the pre-created pyTable type
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
    Return the ElementId of a Lines subcategory with the given name/colour.
    Must be called OUTSIDE any open transaction — creates a subcategory
    which requires its own transaction context in Revit.
    """
    lines_cat = doc.Settings.Categories.get_Item('Lines')
    existing = None
    for sub in lines_cat.SubCategories:
        if sub.Name == name:
            existing = sub
            break
    if existing is None:
        with revit.Transaction('pyTable - Line style: {}'.format(name)):
            existing = doc.Settings.Categories.NewSubcategory(lines_cat, name)
            existing.LineColor = DB.Color(rgb[0], rgb[1], rgb[2])
    else:
        # Update colour in case it changed
        with revit.Transaction('pyTable - Line style colour: {}'.format(name)):
            existing.LineColor = DB.Color(rgb[0], rgb[1], rgb[2])
    gs = existing.GetGraphicsStyle(
        DB.GraphicsStyleType.Projection
    )
    return gs.Id

def _pre_create_line_styles(cell_styles):
    """
    Pre-create all line style subcategories needed for borders in
    cell_styles, plus the invisible 'pyT Off' style.
    Returns a dict mapping rgb tuple -> ElementId.

    Must be called before the main schedule transaction.
    """
    needed = {(255, 255, 255)}   # always need the off/invisible style
    for cs in cell_styles.values():
        for side in ('border_top_color', 'border_bottom_color',
                     'border_left_color', 'border_right_color'):
            c = cs.get(side)
            if c:
                needed.add(tuple(c))
        # Default black border colour when style is set but no colour given
        for side in ('border_top', 'border_bottom',
                     'border_left', 'border_right'):
            if cs.get(side) and cs.get(side) not in ('', 'none'):
                needed.add((0, 0, 0))
                break

    line_ids = {}
    for rgb in needed:
        if rgb == (255, 255, 255):
            name = 'pyT Off'
        else:
            name = 'pyT {:02X}{:02X}{:02X}'.format(rgb[0], rgb[1], rgb[2])
        try:
            line_ids[rgb] = _get_or_create_line_style(name, rgb)
        except Exception as ex:
            logger.error('Line style {} failed: {}'.format(name, ex))
    return line_ids

def apply_row(row):
    """
    Process one UI row:
    1. Read data from the Excel named range
    2. Pre-create pyTable text types (before any transaction)
    3. Call the appropriate Export/ script via exec()

    Returns {'view_name', 'status', 'message'}
    """
    result = {
        'view_name': row.view_name,
        'status':    'error',
        'message':   ''
    }

    logger.debug('pyTable: {} | {} | {}/{}'.format(
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

    fields  = [str(h) for h in rows[0]]
    records = [[str(c) for c in r] for r in rows[1:]]

    if not fields:
        result['message'] = 'Named range has no header row.'
        return result

    # Pre-create text types outside any transaction
    hdr_tt = get_or_create_text_type(
        font=row.font, size_mm=row.size_hdr_mm, bold=True
    )
    dat_tt = get_or_create_text_type(
        font=row.font, size_mm=row.size_dat_mm, bold=False
    )
    hdr_tt_id = hdr_tt.Id if hdr_tt else DB.ElementId.InvalidElementId
    dat_tt_id = dat_tt.Id if dat_tt else DB.ElementId.InvalidElementId

    # Read cell formatting from xlsx
    fmt = read_range_formatting(
        row.file_path,
        row.named_range,
        row.sheet_name
    )

    # Pre-create line styles outside any transaction (Revit requirement)
    line_ids = _pre_create_line_styles(fmt.get('cell_styles', {}))

    # Build payload for export script
    payload = {
        'view_name':          row.view_name,
        'fields':             fields,
        'records':            records,
        'font':               row.font,
        'size_hdr_mm':        row.size_hdr_mm,
        'size_dat_mm':        row.size_dat_mm,
        'hdr_tt_id':          hdr_tt_id,
        'dat_tt_id':          dat_tt_id,
        'view_scale':         row.view_scale,
        'cell_styles':        fmt.get('cell_styles', {}),
        'merges':             fmt.get('merges', []),
        'row_heights':        fmt.get('row_heights', {}),
        'col_widths':         fmt.get('col_widths', {}),
        'default_row_height': fmt.get('default_row_height', 14.0),
        'line_ids':           line_ids,
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

    try:
        _run_export_script(export_script, payload)
        result['status']  = 'success'
        result['message'] = 'Created'
    except Exception as e:
        import traceback
        result['message'] = str(e)
        logger.error(traceback.format_exc())

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
    (hash randomisation), which was previously causing every row to show
    as "changed" even when the source file was never touched.
    """
    import hashlib
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
        return hashlib.md5(combined.encode('utf-8', errors='replace')).hexdigest()
    except Exception:
        return None


class ExcelCardMixin(object):
    """Excel-specific card/row UI methods, mixed into
    PyTableWindow. Everything here assumes fd.get('source_type') == 'xl'."""

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

        inner = StackPanel()
        inner.Orientation = Orientation.Vertical

        # Card header row — Grid so the right-hand group (reload,
        # close) stays flush right regardless of path length. Left
        # group: collapse toggle, source badge, path, last-modified.
        header_row = Grid()
        header_row.Margin = Thickness(0, 0, 0, 8)
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
        collapse_btn.HorizontalContentAlignment = HorizontalAlignment.Center
        collapse_btn.VerticalContentAlignment   = VerticalAlignment.Center
        collapse_btn.VerticalAlignment          = VerticalAlignment.Center
        collapse_btn.Margin          = Thickness(0, 0, 8, 0)
        collapse_btn.Tag     = path
        collapse_btn.ToolTip = 'Collapse'
        collapse_btn.Click  += self._toggle_card_collapse
        self._set_collapse_icon(collapse_btn, False)
        header_left.Children.Add(collapse_btn)

        # Source-type badge (W / XL) — same look as the per-row badge,
        # a touch bigger to hold its own at card-header scale.
        src_badge = Border()
        src_badge.Width             = 24
        src_badge.Height            = 24
        src_badge.CornerRadius      = CornerRadius(4)
        src_badge.Background        = hb(SRC_COLOURS.get(
            fd.get('source_type', 'xl'), '#555'))
        src_badge.Margin            = Thickness(0, 0, 8, 0)
        src_badge.VerticalAlignment = VerticalAlignment.Center
        src_lbl = TextBlock()
        src_lbl.Text                = ('W' if fd.get('source_type') == 'word'
                                        else 'XL')
        src_lbl.FontSize            = 9
        src_lbl.FontWeight          = FontWeights.Bold
        src_lbl.Foreground          = hb('#FFFFFF')
        src_lbl.HorizontalAlignment = HorizontalAlignment.Center
        src_lbl.VerticalAlignment   = VerticalAlignment.Center
        src_badge.Child = src_lbl
        header_left.Children.Add(src_badge)

        heading = TextBlock()
        heading.Text         = path
        heading.TextTrimming = TextTrimming.CharacterEllipsis
        heading.ToolTip      = path
        heading.VerticalAlignment = VerticalAlignment.Center
        heading.Foreground   = hb('#6B7280')
        heading.FontWeight   = FontWeights.SemiBold
        heading.FontSize     = 13
        header_left.Children.Add(heading)
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
        lm_text.VerticalAlignment = VerticalAlignment.Center
        lm_text.Margin           = Thickness(10, 0, 0, 0)
        header_left.Children.Add(lm_text)

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
        batch_btn.Content    = u'Batch \u25be'
        try:
            batch_btn.Style = self.FindResource('SecondaryButtonStyle')
        except Exception as e:
            logger.warning('Failed to apply SecondaryButtonStyle: {}'.format(e))
        batch_btn.FocusVisualStyle = None
        batch_btn.Height      = 24
        batch_btn.Padding     = Thickness(12, 0, 12, 0)
        batch_btn.FontSize    = 11
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
        reload_btn.Width       = 28
        reload_btn.Height      = 28
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
        del_card_btn.FontSize        = 11
        try:
            del_card_btn.Style = self.FindResource('DeleteButtonStyle')
        except Exception as e:
            logger.warning('Failed to apply DeleteButtonStyle: {}'.format(e))
        del_card_btn.FocusVisualStyle = None
        del_card_btn.Width            = 28
        del_card_btn.Height           = 28
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
        col_hdr.Children.Add(_ch('View Name', 124))
        col_hdr.Children.Add(_ch('Sheet',     124))
        col_hdr.Children.Add(_ch('Range',     134))
        col_hdr.Children.Add(_ch('Modified',   96))
        col_hdr.Children.Add(_ch('View Type', 134))

        inner.Children.Add(col_hdr)
        inner.Children.Add(row_panel)
        outer.Child = inner

        fd['card_panel']    = row_panel
        fd['sections_hdr']  = col_hdr
        fd['card_border']   = outer
        fd['card_inner']    = inner

        self.CardsPanel.Children.Add(outer)
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
        row.ViewName = sender.Text.strip()
        sender.Text  = row.ViewName
        if not row.ViewName:
            # Blank = neutral, just reset border — dot stays as-is
            self._style_view_name_conflict(sender, False)
            if row._dot and row.Status not in ('success', 'error'):
                row._dot.Fill = hb('#6B7280')
                row.Status = 'pending'
            return
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

    def _sc_changed(self, sender, e):
        if sender.SelectedItem is None:
            return
        tag = sender.Tag
        if not isinstance(tag, tuple):
            return
        row, rc = tag
        new_sheet = sender.SelectedItem
        row.Sheet = new_sheet
        rc.Items.Clear()
        for r in row.ranges_for(new_sheet):
            rc.Items.Add(r)
        if rc.Items.Count > 0:
            rc.SelectedIndex = 0
            row.NamedRange = rc.Items[0]
            if not row.ViewName:
                row.ViewName = row.NamedRange
                if row._vn_textbox is not None:
                    row._vn_textbox.Text = row.ViewName

    def _rc_changed(self, sender, e):
        if sender.SelectedItem is None:
            return
        row = sender.Tag
        if row:
            row.NamedRange = sender.SelectedItem
            if not row.ViewName:
                row.ViewName = row.NamedRange
                if row._vn_textbox is not None:
                    row._vn_textbox.Text = row.ViewName

    def _vt_changed(self, sender, e):
        if sender.SelectedItem is None:
            return
        row = sender.Tag
        if row:
            row.ViewType = sender.SelectedItem

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

