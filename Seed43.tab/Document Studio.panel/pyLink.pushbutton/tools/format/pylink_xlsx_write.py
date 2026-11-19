# -*- coding: utf-8 -*-
"""
pylink_xlsx_write.py -- writes a pyLink Table Builder table as a plain
.xlsx workbook: one sheet, one named range over the whole table, every
cell's text, font, fill, borders, alignment, merges, column widths and row
heights.

The file is the bridge between the Table Builder and the rest of pyLink:
once it is written, pyLink loads it like any Excel file and its own reader
(pylink_xlsx.py) and schedule writer (Export/create_schedule.py) do the
rest, so a built table and a linked spreadsheet cannot behave differently.
It also opens in Excel as an ordinary workbook.

Pure Python (zipfile + strings), IronPython 2 and CPython 3.

The table model (see pylink_table_builder.TableModel.to_dict):
    {'name': str, 'font': str, 'n_rows': int, 'n_cols': int,
     'col_w_mm': [float], 'row_h_mm': [float],
     'merges': [[r0, c0, r1, c1], ...],
     'cells': {"r,c": {'text', 'bold', 'italic', 'underline', 'size',
                        'halign', 'valign', 'wrap', 'fill', 'color',
                        'b_top', 'b_bottom', 'b_left', 'b_right'}}}
"""
import re
import zipfile

SHEET = 'Table'
DEFAULT_FONT = 'Arial'
DEFAULT_SIZE = 10.0
MDW = 7.6          # Arial max digit width, px - the same figure pylink_xlsx.py reads back with
PT_PER_MM = 1.0 / 0.352778


def _esc(s):
    return (s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
             .replace('"', '&quot;'))


def _col_letter(i):
    s = ''
    i += 1
    while i:
        i, rem = divmod(i - 1, 26)
        s = chr(65 + rem) + s
    return s


def range_name(name):
    """A legal Excel defined name from the table's name: letters, digits,
    underscores and dots, starting with a letter or underscore, and not
    something Excel would read as a cell reference (like A1 or R1C1)."""
    n = re.sub(r'[^A-Za-z0-9_.]', '_', (name or '').strip()) or 'Table'
    if not re.match(r'[A-Za-z_]', n[0]):
        n = '_' + n
    if re.match(r'^[A-Za-z]{1,3}\d+$', n) or re.match(r'^[Rr]\d*[Cc]\d*$', n):
        n = '_' + n
    return n[:250]


def _width_chars(mm):
    """Excel column width (characters) whose drawn width is `mm`, inverting
    the formula pylink_xlsx.py uses to read it back."""
    px = mm * 96.0 / 25.4
    return max(round((px - 5.0) / MDW, 2), 0.5)


def _hex(rgb):
    return 'FF%02X%02X%02X' % tuple(int(v) for v in rgb)


def _parse_rgb(v):
    if not v:
        return None
    if isinstance(v, (list, tuple)):
        return tuple(int(x) for x in v[:3])
    s = str(v).lstrip('#')
    if len(s) == 6:
        return (int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))
    return None


class _Styles(object):
    """De-duplicating font / fill / border / cellXfs tables."""

    def __init__(self, font_name):
        self.font_name = font_name
        self.fonts = [(font_name, DEFAULT_SIZE, False, False, False, None)]
        self.fills = ['none', 'gray125']            # the two fills Excel reserves
        self.borders = [('', '', '', '')]
        self.xfs = [(0, 0, 0, '', '', False)]

    def _idx(self, table, item):
        try:
            return table.index(item)
        except ValueError:
            table.append(item)
            return len(table) - 1

    def xf_for(self, cell):
        font = (self.font_name, float(cell.get('size') or DEFAULT_SIZE),
                bool(cell.get('bold')), bool(cell.get('italic')), bool(cell.get('underline')),
                _parse_rgb(cell.get('color')))
        fill = _parse_rgb(cell.get('fill'))
        border = tuple(cell.get(k) or '' for k in ('b_left', 'b_right', 'b_top', 'b_bottom'))
        fi = self._idx(self.fonts, font)
        fl = self._idx(self.fills, fill) if fill else 0
        bi = self._idx(self.borders, border)
        xf = (fi, fl, bi, (cell.get('halign') or '').lower(), (cell.get('valign') or '').lower(),
              bool(cell.get('wrap')))
        return self._idx(self.xfs, xf)

    def xml(self):
        out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
               '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">']
        out.append('<fonts count="%d">' % len(self.fonts))
        for name, size, b, i, u, color in self.fonts:
            parts = []
            if b:
                parts.append('<b/>')
            if i:
                parts.append('<i/>')
            if u:
                parts.append('<u/>')
            parts.append('<sz val="%g"/>' % size)
            parts.append('<color rgb="%s"/>' % _hex(color) if color else '<color rgb="FF000000"/>')
            parts.append('<name val="%s"/>' % _esc(name))
            out.append('<font>%s</font>' % ''.join(parts))
        out.append('</fonts>')
        out.append('<fills count="%d">' % len(self.fills))
        for f in self.fills:
            if f in ('none', 'gray125'):
                out.append('<fill><patternFill patternType="%s"/></fill>' % f)
            else:
                out.append('<fill><patternFill patternType="solid"><fgColor rgb="%s"/>'
                           '<bgColor indexed="64"/></patternFill></fill>' % _hex(f))
        out.append('</fills>')
        out.append('<borders count="%d">' % len(self.borders))
        for b in self.borders:
            sides = []
            for tag, style in zip(('left', 'right', 'top', 'bottom'), b):
                if style:
                    sides.append('<%s style="%s"><color rgb="FF000000"/></%s>' % (tag, style, tag))
                else:
                    sides.append('<%s/>' % tag)
            out.append('<border>%s<diagonal/></border>' % ''.join(sides))
        out.append('</borders>')
        out.append('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>')
        out.append('<cellXfs count="%d">' % len(self.xfs))
        for fi, fl, bi, h, v, wrap in self.xfs:
            attrs = 'numFmtId="0" fontId="%d" fillId="%d" borderId="%d" xfId="0"' % (fi, fl, bi)
            if fi:
                attrs += ' applyFont="1"'
            if fl:
                attrs += ' applyFill="1"'
            if bi:
                attrs += ' applyBorder="1"'
            al = []
            if h:
                al.append('horizontal="%s"' % h)
            if v:
                al.append('vertical="%s"' % v)
            if wrap:
                al.append('wrapText="1"')
            if al:
                out.append('<xf %s applyAlignment="1"><alignment %s/></xf>' % (attrs, ' '.join(al)))
            else:
                out.append('<xf %s/>' % attrs)
        out.append('</cellXfs>')
        out.append('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>')
        out.append('</styleSheet>')
        return ''.join(out)


def write_table_xlsx(path, model):
    """Write `model` (a dict, see the module docstring) to `path`. Returns
    the named range the table can be found under."""
    n_rows, n_cols = int(model['n_rows']), int(model['n_cols'])
    cells = model.get('cells', {})
    styles = _Styles(model.get('font') or DEFAULT_FONT)
    rname = range_name(model.get('name'))

    covered = set()
    for r0, c0, r1, c1 in model.get('merges', []):
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                if (r, c) != (r0, c0):
                    covered.add((r, c))

    rows_xml = []
    for r in range(n_rows):
        h_mm = model['row_h_mm'][r]
        cs = []
        for c in range(n_cols):
            cell = cells.get('%d,%d' % (r, c)) or {}
            # a covered cell keeps the merge origin's borders/fill, as Excel does,
            # so the merged block's outline survives
            if (r, c) in covered:
                for r0, c0, r1, c1 in model.get('merges', []):
                    if r0 <= r <= r1 and c0 <= c <= c1:
                        origin = dict(cells.get('%d,%d' % (r0, c0)) or {})
                        origin['text'] = ''
                        cell = origin
                        break
            s = styles.xf_for(cell)
            ref = '%s%d' % (_col_letter(c), r + 1)
            text = cell.get('text') or ''
            if text and (r, c) not in covered:
                cs.append('<c r="%s" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                          % (ref, s, _esc(text)))
            else:
                cs.append('<c r="%s" s="%d"/>' % (ref, s))
        rows_xml.append('<row r="%d" ht="%.2f" customHeight="1">%s</row>'
                        % (r + 1, h_mm * PT_PER_MM, ''.join(cs)))

    cols_xml = ''.join('<col min="%d" max="%d" width="%.2f" customWidth="1"/>'
                       % (c + 1, c + 1, _width_chars(model['col_w_mm'][c])) for c in range(n_cols))
    merges = model.get('merges', [])
    merge_xml = ''
    if merges:
        merge_xml = '<mergeCells count="%d">%s</mergeCells>' % (len(merges), ''.join(
            '<mergeCell ref="%s%d:%s%d"/>' % (_col_letter(c0), r0 + 1, _col_letter(c1), r1 + 1)
            for r0, c0, r1, c1 in merges))
    last = '%s%d' % (_col_letter(n_cols - 1), n_rows)
    sheet = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
             'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
             '<dimension ref="A1:%s"/>'
             '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
             '<sheetFormatPr defaultRowHeight="15"/>'
             '<cols>%s</cols><sheetData>%s</sheetData>%s'
             '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
             '</worksheet>') % (last, cols_xml, ''.join(rows_xml), merge_xml)

    workbook = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<sheets><sheet name="%s" sheetId="1" r:id="rId1"/></sheets>'
                '<definedNames><definedName name="%s">%s!$A$1:$%s$%d</definedName></definedNames>'
                '</workbook>') % (SHEET, rname, SHEET, _col_letter(n_cols - 1), n_rows)

    content_types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                     '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                     '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                     '<Default Extension="xml" ContentType="application/xml"/>'
                     '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                     '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                     '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                     '</Types>')
    root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                 '</Relationships>')
    wb_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
               '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
               '</Relationships>')

    def _b(s):
        return s.encode('utf-8') if not isinstance(s, bytes) else s

    tmp = path + '.tmp'
    z = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
    try:
        z.writestr('[Content_Types].xml', _b(content_types))
        z.writestr('_rels/.rels', _b(root_rels))
        z.writestr('xl/workbook.xml', _b(workbook))
        z.writestr('xl/_rels/workbook.xml.rels', _b(wb_rels))
        z.writestr('xl/styles.xml', _b(styles.xml()))
        z.writestr('xl/worksheets/sheet1.xml', _b(sheet))
    finally:
        z.close()
    import os
    if os.path.exists(path):
        os.remove(path)
    os.rename(tmp, path)
    return rname
