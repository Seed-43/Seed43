# -*- coding: utf-8 -*-
from pyrevit import script

import os
import re
import sys as _sys
import zipfile as _zipfile

logger = script.get_logger()

_sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from pylink_shared import _col_letter_to_index

"""
pylink_xlsx.py -- xlsx-specific reading for pyLink: named ranges and
sheet names, named-range cell values, and per-cell formatting (fonts,
fills, borders, alignment, merges, row/column sizing) - all parsed
directly from the xlsx zip, no COM, no third-party libraries. Returns
the same dict/list shapes pylink_ods.py returns, so pylink_excel.py's
dispatchers and the Export/create_*.py view builders don't need to
know which spreadsheet format a table actually came from.
"""


# ── Named ranges + sheet names ──

def read_xlsx_named_ranges(file_path):
    """
    Read named ranges and sheet names from an xlsx workbook.

    Returns:
        {
          'named_ranges': [str, ...],           # all range names
          'sheets':       [str, ...],           # all sheet names
          'sheet_ranges': {sheet: [name, ...]}, # ranges per sheet
        }

    Sheet assignment:
      1. If definedName has a localSheetId attribute, it belongs to
         that sheet (0-based index into the sheets list).
      2. Otherwise parse the sheet name from the range reference
         formula, e.g. "Sheet1!$A$1" -> Sheet1.
      3. If neither applies (bare reference like "A1"), the range is
         global and added to ALL sheets.
    """
    result = {'named_ranges': [], 'sheets': [], 'sheet_ranges': {}}

    try:
        with _zipfile.ZipFile(file_path, 'r') as z:
            xml_bytes = z.read('xl/workbook.xml')

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
        logger.debug(
            'No named ranges found in xlsx workbook.xml (unreadable or empty)'
        )

    return result


# ── Cell data ──

def read_xlsx_range_data(file_path, named_range, sheet_name):
    """
    Read rows from an xlsx named range directly from its zip.
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
            'read_xlsx_range_data failed for "{}": {}'.format(
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


# ── Cell formatting ──

def read_xlsx_range_formatting(file_path, named_range, sheet_name):
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
            'valign': str,       # 'Top', 'Center', 'Bottom' (Excel default when unset is Bottom)
            'wrap': bool,        # Excel alignment/@wrapText
            'rotation': int,     # Excel textRotation: 0=horizontal, 1-90=CCW, 91-180=CW (90+angle), 255=stacked
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
                rotation   = xf.get('rotation', 0)
                raw_align  = xf.get('align', '')
                if raw_align:
                    halign = raw_align.capitalize()
                elif rotation:
                    # Excel's own default rendering for ROTATED text
                    # with no explicit horizontal alignment set is
                    # visually centered in the column - "General"/
                    # unset does not mean Left once rotation is
                    # involved, unlike unrotated text where Left IS
                    # the right default for General.
                    halign = 'Center'
                else:
                    halign = 'Left'
                valign     = xf.get('valign', '') or 'bottom'
                valign     = valign.capitalize()
                wrap       = xf.get('wrap', False)

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
                    'valign':             valign,
                    'wrap':               wrap,
                    'rotation':           rotation,
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
        logger.error('read_xlsx_range_formatting failed: {}'.format(e))

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
    Uses the cache populated by _set_theme_cache during read_xlsx_range_formatting.
    """
    global _theme_colour_cache
    return _theme_colour_cache.get(idx)

def _set_theme_cache(colours):
    """Called from read_xlsx_range_formatting with resolved theme colours."""
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
            'valign':   '',
            'wrap':     False,
            'rotation': 0,
        }

        align_nodes = node.GetElementsByTagName('alignment')
        if align_nodes.Count == 0:
            align_nodes = node.GetElementsByTagName('alignment', NS)
        if align_nodes.Count > 0:
            h = align_nodes[0].Attributes.GetNamedItem('horizontal')
            xf['align'] = h.Value if h else ''
            # Excel's own default vertical alignment (when the attribute
            # is absent) is 'bottom' - matches Demo.xlsx's D/G/J columns
            # (v=None in openpyxl) which visually sit at the cell bottom.
            v = align_nodes[0].Attributes.GetNamedItem('vertical')
            xf['valign'] = v.Value if v else 'bottom'
            w = align_nodes[0].Attributes.GetNamedItem('wrapText')
            if w:
                xf['wrap'] = w.Value in ('1', 'true', 'True')
            # textRotation: 0-90 = counter-clockwise from horizontal,
            # 91-180 = clockwise (stored as 90 + actual angle), 255 = stacked/vertical.
            r = align_nodes[0].Attributes.GetNamedItem('textRotation')
            if r:
                try:
                    xf['rotation'] = int(r.Value)
                except Exception:
                    xf['rotation'] = 0

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
