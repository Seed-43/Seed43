# -*- coding: utf-8 -*-
from pyrevit import script

import os
import re
import sys as _sys
import zipfile as _zipfile

logger = script.get_logger()

_sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from pylink_shared import _col_letter_to_index, PT_MM

"""
pylink_ods.py -- ODS (OpenDocument Spreadsheet)-specific reading for
pyLink: named ranges and sheet names, named-range cell values, and
per-cell formatting, all parsed directly from content.xml inside the
ods zip - no COM, no third-party libraries. Returns the same dict/list
shapes pylink_xlsx.py returns, so pylink_excel.py's dispatchers and
the Export/create_*.py view builders don't need to know which
spreadsheet format a table actually came from.
"""


# ── Named ranges + sheet names ──

def _ods_sheet_from_ref(ref):
    """Pull the sheet name out of an ODS table:cell-range-address /
    table:base-cell-address value, e.g. "$Sheet1.$E$5:.$F$8" -> "Sheet1",
    or "$'My Sheet'.$A$1" -> "My Sheet" (single-quoted when the sheet
    name itself contains a space or other special character)."""
    ref = (ref or '').strip()
    if ref.startswith('$'):
        ref = ref[1:]
    if ref.startswith("'"):
        end = ref.find("'", 1)
        if end != -1:
            return ref[1:end]
    if '.' in ref:
        return ref.split('.')[0]
    return None


def read_ods_named_ranges(file_path):
    """
    Read named ranges and sheet names from an ODS spreadsheet.

    Returns:
        {
          'named_ranges': [str, ...],           # all range names
          'sheets':       [str, ...],           # all sheet names
          'sheet_ranges': {sheet: [name, ...]}, # ranges per sheet
        }

    Sheet assignment: table:named-range's own table:cell-range-address
    attribute always includes the sheet name (e.g. "$Sheet1.$E$5:.$F$8"
    or "$'My Sheet'.$A$1"), so there's no separate localSheetId/global
    case to handle like xlsx has - every ODS named range is inherently
    sheet-scoped.
    """
    result = {'named_ranges': [], 'sheets': [], 'sheet_ranges': {}}

    try:
        with _zipfile.ZipFile(file_path, 'r') as z:
            xml_bytes = z.read('content.xml')

        import clr
        clr.AddReference('System.Xml')
        from System.Xml import XmlDocument

        xml_doc = XmlDocument()
        xml_doc.LoadXml(xml_bytes.decode('utf-8'))

        # ── Sheets ── (table:table elements, qualified-name match works
        # fine here since content.xml consistently uses the table: prefix)
        sheets = []
        table_nodes = xml_doc.GetElementsByTagName('table:table')
        for i in range(table_nodes.Count):
            name = table_nodes[i].GetAttribute('table:name')
            if name:
                sheets.append(name)
        result['sheets'] = sheets

        # ── Named ranges ──
        sheet_ranges = {s: [] for s in sheets}
        named_ranges = []

        nr_nodes = xml_doc.GetElementsByTagName('table:named-range')
        for i in range(nr_nodes.Count):
            node = nr_nodes[i]
            name = node.GetAttribute('table:name')
            if not name:
                continue
            named_ranges.append(name)

            ref = (node.GetAttribute('table:cell-range-address') or
                   node.GetAttribute('table:base-cell-address'))
            target_sheet = _ods_sheet_from_ref(ref)

            if target_sheet and target_sheet in sheet_ranges:
                sheet_ranges[target_sheet].append(name)
            else:
                # Sheet name didn't match any known sheet (shouldn't
                # normally happen) - fall back to treating it as global,
                # same as xlsx's own fallback for an unresolvable ref.
                for s in sheets:
                    sheet_ranges[s].append(name)

        result['named_ranges'] = named_ranges
        result['sheet_ranges'] = sheet_ranges

    except Exception:
        logger.debug(
            'No named ranges found in ods content.xml (unreadable or empty)'
        )

    return result


# ── Cell data ──
# _ods_read_table_grid() does the single grid walk that both
# read_ods_range_data and read_ods_range_formatting need (values,
# table:style-name per cell, merges, row style names). The style:style
# resolution below turns those names into the same
# cell_styles/merges/row_heights/col_widths shape pylink_xlsx.py returns.

_ODS_CELL_RE = re.compile(r'\$?([A-Za-z]+)\$?(\d+)')


def _ods_parse_range_bounds(ref):
    """Parse an ODS cell-range-address, e.g. "$Sheet1.$E$5:.$F$8" or a
    single-cell "$Sheet1.$A$1", into 0-based inclusive (min_col,
    min_row, max_col, max_row). The end half of a range only repeats
    the sheet name when it differs from the start (rare) - normally
    it's just ".$F$8".

    Each half's sheet-name prefix is stripped (everything up to and
    including the LAST '.' - a well-formed ODF cell address has no
    further literal dots after the sheet/cell separator) before the
    cell regex runs, same as the xlsx path stripping everything before
    '!'. Skipping this let a sheet name like "Sheet2 (2)" get matched
    AS the cell reference instead - "Sheet2" alone (letters directly
    followed by digits) satisfies the same regex meant for "E5", and
    since it appears first in the string an unanchored search on the
    raw ref found it before the real cell reference."""
    if not ref:
        return 0, 0, None, None
    halves = ref.split(':')

    def _cell(part):
        cell_part = part.rsplit('.', 1)[-1]
        m = _ODS_CELL_RE.search(cell_part)
        if not m:
            return None
        return _col_letter_to_index(m.group(1)), int(m.group(2)) - 1

    start = _cell(halves[0])
    if start is None:
        return 0, 0, None, None
    end = _cell(halves[1]) if len(halves) > 1 else start
    end = end or start
    return start[0], start[1], end[0], end[1]


def _ods_resolve_range(xml_doc, named_range, sheet_name):
    """Find `named_range`'s table:named-range node and return (sheet,
    min_col, min_row, max_col, max_row), 0-based/inclusive. Falls back
    to the caller-supplied sheet_name if the range itself can't be
    resolved, same fallback the xlsx path uses."""
    nr_nodes = xml_doc.GetElementsByTagName('table:named-range')
    for i in range(nr_nodes.Count):
        node = nr_nodes[i]
        if node.GetAttribute('table:name') == named_range:
            ref = (node.GetAttribute('table:cell-range-address') or
                   node.GetAttribute('table:base-cell-address'))
            sheet = _ods_sheet_from_ref(ref) or sheet_name
            min_col, min_row, max_col, max_row = _ods_parse_range_bounds(ref)
            return sheet, min_col, min_row, max_col, max_row
    return None, 0, 0, None, None


def _ods_find_table(xml_doc, sheet_name):
    tables = xml_doc.GetElementsByTagName('table:table')
    for i in range(tables.Count):
        if tables[i].GetAttribute('table:name') == sheet_name:
            return tables[i]
    return None


def _ods_cell_text(cell):
    """A table:table-cell's display text: joined text:p children when
    present (the normal case), else the raw office:*-value attribute
    for the cell's value type - covers numeric/date/boolean cells
    that carry their value as an attribute with no text:p child."""
    p_nodes = cell.GetElementsByTagName('text:p')
    if p_nodes.Count > 0:
        return '\n'.join(
            p_nodes[i].InnerText for i in range(p_nodes.Count)
        )
    for attr in ('office:string-value', 'office:value',
                 'office:date-value', 'office:time-value',
                 'office:boolean-value'):
        val = cell.GetAttribute(attr)
        if val:
            return val
    return ''


def _ods_read_table_grid(table, min_col, min_row, max_col, max_row):
    """Walk an ODS table:table's rows/cells within [min_row..max_row,
    min_col..max_col] (0-based inclusive; max_col/max_row of None
    means unbounded), expanding table:number-*-repeated compression
    and resolving table:number-*-spanned merges. Returns:
        {
          'rows':             [[val, ...], ...],  # same shape _extract_rows returns for xlsx
          'cell_style_names': {(rel_row, rel_col): table:style-name},
          'merges':           [(r1, c1, r2, c2), ...],
          'row_style_names':  {rel_row: table:style-name},
        }
    Blank rows inside the range still get an (empty-valued) entry in
    'rows', so its indices stay aligned with row_style_names/merges.
    """
    rows_out          = []
    cell_style_names  = {}
    merges            = []
    row_style_names   = {}

    row_idx = 0
    row_nodes = table.GetElementsByTagName('table:table-row')
    for ri in range(row_nodes.Count):
        row_node = row_nodes[ri]
        row_repeat = int(row_node.GetAttribute('table:number-rows-repeated') or '1')

        if row_idx + row_repeat - 1 < min_row:
            row_idx += row_repeat
            continue
        if max_row is not None and row_idx > max_row:
            break

        cell_nodes = row_node.ChildNodes
        for _rep in range(row_repeat):
            if max_row is not None and row_idx > max_row:
                break
            if row_idx < min_row:
                row_idx += 1
                continue

            rel_row = row_idx - min_row
            row_style_names[rel_row] = row_node.GetAttribute('table:style-name')

            col_idx = 0
            col_vals = {}
            for ci in range(cell_nodes.Count):
                cell = cell_nodes[ci]
                try:
                    if cell.Attributes is None:
                        continue
                except Exception:
                    continue
                tag = cell.Name
                if tag not in ('table:table-cell', 'table:covered-table-cell'):
                    continue

                repeat = int(cell.GetAttribute('table:number-columns-repeated') or '1')
                if tag == 'table:table-cell' and col_idx >= min_col and \
                        (max_col is None or col_idx <= max_col):
                    rel_col = col_idx - min_col
                    col_vals[rel_col] = _ods_cell_text(cell)
                    style_name = cell.GetAttribute('table:style-name')
                    if style_name:
                        cell_style_names[(rel_row, rel_col)] = style_name
                    span_c = int(cell.GetAttribute('table:number-columns-spanned') or '1')
                    span_r = int(cell.GetAttribute('table:number-rows-spanned') or '1')
                    if span_c > 1 or span_r > 1:
                        merges.append((
                            rel_row, rel_col,
                            rel_row + span_r - 1, rel_col + span_c - 1
                        ))

                col_idx += repeat
                if max_col is not None and col_idx > max_col:
                    break

            if col_vals:
                end_col = max_col if max_col is not None else max(col_vals.keys())
                width = end_col - min_col + 1
            else:
                width = (max_col - min_col + 1) if max_col is not None else 1
            rows_out.append([col_vals.get(c, '') for c in range(width)])

            row_idx += 1

    return {
        'rows':             rows_out,
        'cell_style_names': cell_style_names,
        'merges':           merges,
        'row_style_names':  row_style_names,
    }


def read_ods_range_data(file_path, named_range, sheet_name):
    """
    Read rows from an ODS named range directly from its zip.
    Returns list of lists: [[header1, header2, ...], [val, val, ...], ...]
    First row is column headers.
    """
    try:
        import clr
        clr.AddReference('System.Xml')
        from System.Xml import XmlDocument

        with _zipfile.ZipFile(file_path, 'r') as z:
            xml_doc = XmlDocument()
            xml_doc.LoadXml(z.read('content.xml').decode('utf-8'))

        target_sheet, min_col, min_row, max_col, max_row = _ods_resolve_range(
            xml_doc, named_range, sheet_name
        )
        if target_sheet is None:
            logger.error('ODS named range not found: {}'.format(named_range))
            return []

        table = _ods_find_table(xml_doc, target_sheet)
        if table is None:
            logger.error('Sheet not found in ODS: {}'.format(target_sheet))
            return []

        grid = _ods_read_table_grid(table, min_col, min_row, max_col, max_row)
        return grid['rows']

    except Exception as e:
        logger.error(
            'read_ods_range_data failed for "{}": {}'.format(named_range, e)
        )
        return []


# ── Cell formatting ──

def _ods_collect_styles(xml_doc):
    """Every style:style element under office:automatic-styles and
    office:styles, keyed by style:name - ODF splits per-document
    ('automatic') and named/inherited ('office:styles') definitions
    across both containers, and a cell's own style commonly inherits
    from one in the other via style:parent-style-name."""
    styles = {}
    for container_tag in ('office:automatic-styles', 'office:styles'):
        containers = xml_doc.GetElementsByTagName(container_tag)
        for ci in range(containers.Count):
            nodes = containers[ci].GetElementsByTagName('style:style')
            for i in range(nodes.Count):
                name = nodes[i].GetAttribute('style:name')
                if name and name not in styles:
                    styles[name] = nodes[i]
    return styles


def _ods_font_face_map(xml_doc):
    """Map an ODF style:font-name reference (e.g. "Arial1") to its
    real font family (e.g. "Arial") via office:font-face-decls -
    LibreOffice/OpenOffice usually declare the face separately rather
    than writing the family name directly on style:font-name."""
    face_map = {}
    decls = xml_doc.GetElementsByTagName('office:font-face-decls')
    for di in range(decls.Count):
        faces = decls[di].GetElementsByTagName('style:font-face')
        for fi in range(faces.Count):
            name   = faces[fi].GetAttribute('style:name')
            family = faces[fi].GetAttribute('svg:font-family')
            if name and family:
                face_map[name] = family.strip('\'"')
    return face_map


def _ods_style_prop(styles, style_name, prop_tag, attr, depth=0):
    """One text/table-cell/paragraph-properties attribute for a style,
    walking style:parent-style-name up to 4 levels deep - ODF styles
    commonly inherit most properties from a base style rather than
    repeating every one of them on each cell's own style. Stops at the
    first value found, or when the chain runs out."""
    if not style_name or style_name not in styles or depth > 4:
        return None
    node = styles[style_name]
    prop_nodes = node.GetElementsByTagName(prop_tag)
    if prop_nodes.Count > 0:
        val = prop_nodes[0].GetAttribute(attr)
        if val:
            return val
    parent = node.GetAttribute('style:parent-style-name')
    if parent and parent != style_name:
        return _ods_style_prop(styles, parent, prop_tag, attr, depth + 1)
    return None


def _ods_pt_value(s, default=11.0):
    """"11pt" -> 11.0. Any leading number, not just whole pt values."""
    if not s:
        return default
    m = re.match(r'([\d.]+)', s.strip())
    return float(m.group(1)) if m else default


_ODS_COLOR_RE = re.compile(r'^#([0-9A-Fa-f]{6})$')


def _ods_hex_to_rgb(value):
    if not value:
        return None
    m = _ODS_COLOR_RE.match(value.strip())
    if not m:
        return None
    h = m.group(1)
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _ods_rotation_to_excel(raw):
    """Convert an ODF style:rotation-angle (plain CCW degrees, 0-359)
    into the 0 / 1-90 / 91-180 encoding Excel's own textRotation uses
    (see _excel_rotation_to_radians in Export/create_drafting.py), so
    an ODS-rotated header renders the same way an Excel-rotated one
    does. ODS's very common "read top to bottom" vertical header is
    rotation-angle=270 (90 clockwise), which lands on Excel's 180 -
    exactly the value Excel itself uses for that same reading
    direction."""
    if not raw:
        return 0
    try:
        angle = int(float(raw)) % 360
    except Exception:
        return 0
    if angle == 0:
        return 0
    if 1 <= angle <= 90:
        return angle
    cw = (360 - angle) % 360
    return (90 + cw) if 1 <= cw <= 90 else 0


def _ods_parse_border(raw):
    """Parse an ODF border shorthand like "0.75pt solid #1F4E78" into
    (excel_style_keyword, (r, g, b)) - bucketed by width into the same
    thin/medium/thick vocabulary _excel_border_weight (in
    Export/create_drafting.py) already maps, so an ODS border renders
    at a comparable weight to an equivalent Excel one instead of a
    second, parallel weight system."""
    if not raw or raw.strip().lower() == 'none':
        return '', None
    width_pt = 1.0
    color = None
    for part in raw.strip().split():
        if part.lower().endswith('pt'):
            try:
                width_pt = float(part[:-2])
            except Exception:
                pass
        elif part.startswith('#'):
            color = _ods_hex_to_rgb(part)
    if width_pt >= 2.0:
        style = 'thick'
    elif width_pt >= 1.0:
        style = 'medium'
    else:
        style = 'thin'
    return style, (color or (0, 0, 0))


_ODS_HALIGN = {'start': 'Left', 'left': 'Left', 'center': 'Center',
               'end': 'Right', 'right': 'Right', 'justify': 'Left'}
_ODS_VALIGN = {'top': 'Top', 'middle': 'Center', 'bottom': 'Bottom'}

_ODS_UNIT_TO_MM = {'cm': 10.0, 'mm': 1.0, 'in': 25.4, 'pt': PT_MM,
                    'px': 25.4 / 96.0}


def _ods_length_to_mm(value):
    """Parse an ODF length like "2.267cm" / "56.7pt" / "0.5in" into mm.
    A bare number with no unit suffix is treated as cm, matching what
    LibreOffice/OpenOffice actually write for column widths."""
    if not value:
        return None
    value = value.strip()
    for unit, factor in _ODS_UNIT_TO_MM.items():
        if value.endswith(unit):
            try:
                return float(value[:-len(unit)]) * factor
            except Exception:
                return None
    try:
        return float(value) * _ODS_UNIT_TO_MM['cm']
    except Exception:
        return None


def _ods_cell_style(styles, face_map, style_name):
    """Resolve one ODS table-cell style:style into the same shape
    pylink_xlsx.py's read_xlsx_range_formatting builds in
    cell_styles, so create_drafting.py/create_schedule.py/
    create_legend.py don't need to know which spreadsheet format a
    table actually came from."""
    def prop(tag, attr):
        return _ods_style_prop(styles, style_name, tag, attr)

    font_raw = prop('style:text-properties', 'style:font-name')

    result = {
        'font_name': face_map.get(font_raw, font_raw) if font_raw else 'Arial',
        'font_size': _ods_pt_value(prop('style:text-properties', 'fo:font-size')),
        'bold':      prop('style:text-properties', 'fo:font-weight') == 'bold',
        'italic':    prop('style:text-properties', 'fo:font-style') == 'italic',
        'underline': (prop('style:text-properties', 'style:text-underline-style')
                      or 'none') != 'none',
        'color_rgb': _ods_hex_to_rgb(prop('style:text-properties', 'fo:color')),
        'halign':    _ODS_HALIGN.get(
            (prop('style:paragraph-properties', 'fo:text-align') or '').lower(), 'Left'),
        'valign':    _ODS_VALIGN.get(
            (prop('style:table-cell-properties', 'style:vertical-align') or '').lower(),
            'Bottom'),
        'wrap':      prop('style:table-cell-properties', 'fo:wrap-option') == 'wrap',
        'rotation':  _ods_rotation_to_excel(
            prop('style:table-cell-properties', 'style:rotation-angle')),
        'fill_rgb':  _ods_hex_to_rgb(
            prop('style:table-cell-properties', 'fo:background-color')),
    }
    for side in ('top', 'bottom', 'left', 'right'):
        b_style, b_color = _ods_parse_border(
            prop('style:table-cell-properties', 'fo:border-' + side))
        result['border_' + side] = b_style
        result['border_' + side + '_color'] = b_color
    return result


def _ods_col_style_names(table, min_col, max_col):
    """Map each 0-based column index in [min_col..max_col] to its
    table:table-column table:style-name, expanding
    table:number-columns-repeated the same way row/cell repeats are
    expanded in _ods_read_table_grid."""
    result = {}
    col_idx = 0
    col_nodes = table.GetElementsByTagName('table:table-column')
    for i in range(col_nodes.Count):
        node = col_nodes[i]
        repeat = int(node.GetAttribute('table:number-columns-repeated') or '1')
        style_name = node.GetAttribute('table:style-name')
        for _ in range(repeat):
            if max_col is not None and col_idx > max_col:
                return result
            if col_idx >= min_col:
                result[col_idx - min_col] = style_name
            col_idx += 1
    return result


def read_ods_range_formatting(file_path, named_range, sheet_name):
    """
    Read cell formatting from an ODS named range. Returns the exact
    same result shape pylink_xlsx.py's read_xlsx_range_formatting
    does (cell_styles/merges/row_heights/col_widths) - see that
    function's docstring for the full shape.
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
            xml_doc = XmlDocument()
            xml_doc.LoadXml(z.read('content.xml').decode('utf-8'))

        target_sheet, min_col, min_row, max_col, max_row = _ods_resolve_range(
            xml_doc, named_range, sheet_name
        )
        if target_sheet is None:
            return result

        table = _ods_find_table(xml_doc, target_sheet)
        if table is None:
            return result

        grid      = _ods_read_table_grid(table, min_col, min_row, max_col, max_row)
        styles    = _ods_collect_styles(xml_doc)
        face_map  = _ods_font_face_map(xml_doc)

        result['merges'] = grid['merges']

        for (r, c), style_name in grid['cell_style_names'].items():
            result['cell_styles'][(r, c)] = _ods_cell_style(styles, face_map, style_name)

        for row_idx, style_name in grid['row_style_names'].items():
            mm = _ods_length_to_mm(_ods_style_prop(
                styles, style_name, 'style:table-row-properties', 'style:row-height'))
            if mm:
                result['row_heights'][row_idx] = mm / PT_MM

        for col_idx, style_name in _ods_col_style_names(table, min_col, max_col).items():
            mm = _ods_length_to_mm(_ods_style_prop(
                styles, style_name, 'style:table-column-properties', 'style:column-width'))
            if mm:
                result['col_widths'][col_idx] = mm

    except Exception as e:
        logger.error('read_ods_range_formatting failed: {}'.format(e))

    return result
