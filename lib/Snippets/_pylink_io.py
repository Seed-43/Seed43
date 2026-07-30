# -*- coding: utf-8 -*-
"""
Format-agnostic export/import I/O for PyLink.

Deliberately has ZERO third-party dependencies. xlsx and ods are both just
zip archives of XML - this reads/writes that XML directly via zipfile and
xml.etree.ElementTree, both in the Python 2.7 stdlib IronPython ships. No
openpyxl, no odfpy, no xlsxwriter - none of those are guaranteed present in
a given pyRevit environment (there's no pip under IronPython 2, so any
third-party package has to be manually vendored into the extension), and
guessing wrong here has already cost several rounds of "No module named X".
This trades some polish (no cell shading for read-only columns, no hidden
columns, no merged-cell handling on read) for something that just works
without any external dependency at all.

Shared data shape used by both writers/readers so the picker UI and the
diff/import logic never need to know which file format was used:

    columns : ordered list of dicts, each:
        {
            "key":      "Comments",      # parameter name, also the header text
            "kind":     "instance" | "type",
            "readonly": True | False,
        }

    rows : list of dicts, each:
        {
            "_eid":      123456,         # Element.Id.Value - hidden key column
            "_category": "Walls",        # hidden, for grouping/display only
            "Comments":  "some value",
            ...                          # one key per column["key"]
        }

Hidden columns (prefixed with "_") are always written first and are the
match key on import - never shown for editing, never diffed as a parameter.

snippets.yaml entry:
  _pylink_io.py:
    description: >
      Format-agnostic export/import I/O for PyLink (xlsx and ods, both
      read directly via zipfile + xml.etree.ElementTree - no third-party
      packages). Used by PyLink.pushbutton only.
    functions:
      write_xlsx:  Write columns/rows to an .xlsx file.
      write_ods:   Write columns/rows to an .ods file.
      read_xlsx:   Read an .xlsx file back into columns/rows.
      read_ods:    Read an .ods file back into columns/rows.
      write_workbook: Route to write_xlsx/write_ods by file extension.
      read_workbook:  Route to read_xlsx/read_ods by file extension.
"""

import os
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape as _xml_escape

HIDDEN_COLUMNS = ("_eid", "_category")

XML_DECL = u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'


# ── shared helpers ─────────────────────────────────────────────────────────

def _header_row(columns):
    """Return the full ordered header list: hidden key columns + visible params."""
    return list(HIDDEN_COLUMNS) + [c["key"] for c in columns]


def _row_values(row, headers):
    return [row.get(h, u"") for h in headers]


def _cell_text(value):
    if value is None:
        return u""
    if not isinstance(value, unicode):
        value = unicode(value)
    return value


def _col_letter(idx):
    """0-based column index -> Excel-style column letter(s) (0 -> A, 26 -> AA)."""
    letters = u""
    idx += 1
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters = unichr(65 + rem) + letters
    return letters


def _col_index(letter):
    """Excel-style column letter(s) -> 0-based column index."""
    idx = 0
    for ch in letter:
        idx = idx * 26 + (ord(ch) - 64)
    return idx - 1


# ══ XLSX ═════════════════════════════════════════════════════════════════

_NS_MAIN = u"http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_RELS = u"http://schemas.openxmlformats.org/package/2006/relationships"
_NS_DOC_RELS = u"http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def write_xlsx(path, columns, rows):
    headers = _header_row(columns)

    content_types = XML_DECL + (
        u'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        u'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        u'<Default Extension="xml" ContentType="application/xml"/>'
        u'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        u'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        u'</Types>'
    )

    root_rels = XML_DECL + (
        u'<Relationships xmlns="{ns}">'
        u'<Relationship Id="rId1" Type="{doc_ns}/officeDocument" Target="xl/workbook.xml"/>'
        u'</Relationships>'
    ).format(ns=_NS_RELS, doc_ns=_NS_DOC_RELS)

    workbook_xml = XML_DECL + (
        u'<workbook xmlns="{main}" xmlns:r="{doc_ns}">'
        u'<sheets><sheet name="PyLink Export" sheetId="1" r:id="rId1"/></sheets>'
        u'</workbook>'
    ).format(main=_NS_MAIN, doc_ns=_NS_DOC_RELS)

    workbook_rels = XML_DECL + (
        u'<Relationships xmlns="{ns}">'
        u'<Relationship Id="rId1" Type="{doc_ns}/worksheet" Target="worksheets/sheet1.xml"/>'
        u'</Relationships>'
    ).format(ns=_NS_RELS, doc_ns=_NS_DOC_RELS)

    # Every cell is written as an inline string - our own data is always
    # display text (from get_param_display_value), so there's no need to
    # distinguish numeric cells on write. Readers (including Excel/LibreOffice
    # themselves) handle inlineStr cells natively.
    row_xml_parts = []
    for r_idx, row in enumerate(rows, start=2):  # row 1 is the header
        cells = []
        for c_idx, header in enumerate(headers):
            value = _cell_text(row.get(header, u""))
            ref = u"{0}{1}".format(_col_letter(c_idx), r_idx)
            cells.append(
                u'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{val}</t></is></c>'.format(
                    ref=ref, val=_xml_escape(value)))
        row_xml_parts.append(u'<row r="{0}">{1}</row>'.format(r_idx, u"".join(cells)))

    header_cells = []
    for c_idx, header in enumerate(headers):
        ref = u"{0}1".format(_col_letter(c_idx))
        header_cells.append(
            u'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{val}</t></is></c>'.format(
                ref=ref, val=_xml_escape(header)))
    header_row_xml = u'<row r="1">{0}</row>'.format(u"".join(header_cells))

    sheet_xml = XML_DECL + (
        u'<worksheet xmlns="{main}"><sheetData>{header}{rows}</sheetData></worksheet>'
    ).format(main=_NS_MAIN, header=header_row_xml, rows=u"".join(row_xml_parts))

    zf = zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED)
    try:
        zf.writestr("[Content_Types].xml", content_types.encode("utf-8"))
        zf.writestr("_rels/.rels", root_rels.encode("utf-8"))
        zf.writestr("xl/workbook.xml", workbook_xml.encode("utf-8"))
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels.encode("utf-8"))
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml.encode("utf-8"))
    finally:
        zf.close()


def _xlsx_find_first_sheet_path(zf):
    """Resolve the actual worksheet part path via workbook.xml + its rels,
    since a real Excel/LibreOffice save can rename sheet1.xml to anything."""
    try:
        wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
        sheet_el = wb_root.find(
            "{{{0}}}sheets/{{{0}}}sheet".format(_NS_MAIN))
        r_id = sheet_el.get("{{{0}}}id".format(_NS_DOC_RELS))

        rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        for rel in rels_root:
            if rel.get("Id") == r_id:
                target = rel.get("Target")
                return "xl/" + target if not target.startswith("/") else target.lstrip("/")
    except Exception:
        pass
    return "xl/worksheets/sheet1.xml"


def read_xlsx(path):
    zf = zipfile.ZipFile(path, "r")
    try:
        shared_strings = []
        if "xl/sharedStrings.xml" in zf.namelist():
            sst_root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in sst_root:
                # <si> can contain a single <t>, or multiple <r><t> runs.
                texts = [t.text or u"" for t in si.iter("{{{0}}}t".format(_NS_MAIN))]
                shared_strings.append(u"".join(texts))

        sheet_path = _xlsx_find_first_sheet_path(zf)
        sheet_root = ET.fromstring(zf.read(sheet_path))
        sheet_data = sheet_root.find("{{{0}}}sheetData".format(_NS_MAIN))
        if sheet_data is None:
            return [], []

        def _cell_value(c_el):
            t = c_el.get("t")
            if t == "inlineStr":
                is_el = c_el.find("{{{0}}}is".format(_NS_MAIN))
                if is_el is not None:
                    texts = [t_el.text or u"" for t_el in is_el.iter("{{{0}}}t".format(_NS_MAIN))]
                    return u"".join(texts)
                return u""
            v_el = c_el.find("{{{0}}}v".format(_NS_MAIN))
            if v_el is None or v_el.text is None:
                return u""
            if t == "s":
                try:
                    return shared_strings[int(v_el.text)]
                except (ValueError, IndexError):
                    return u""
            # Plain numeric cell (no t attribute) or t="str" formula result.
            return v_el.text

        rows_out = []
        headers = None
        for row_el in sheet_data.findall("{{{0}}}row".format(_NS_MAIN)):
            row_values = {}
            for c_el in row_el.findall("{{{0}}}c".format(_NS_MAIN)):
                ref = c_el.get("r", u"")
                letters = u"".join(ch for ch in ref if ch.isalpha())
                if not letters:
                    continue
                col_idx = _col_index(letters)
                row_values[col_idx] = _cell_value(c_el)

            if not row_values:
                continue
            max_col = max(row_values.keys())
            ordered = [row_values.get(i, u"") for i in range(max_col + 1)]

            if headers is None:
                headers = ordered
                continue

            if not any(v for v in ordered):
                continue
            rows_out.append(dict(zip(headers, ordered)))

        return (headers or []), rows_out
    finally:
        zf.close()


# ══ ODS ══════════════════════════════════════════════════════════════════

_NS_OFFICE = u"urn:oasis:names:tc:opendocument:xmlns:office:1.0"
_NS_TABLE = u"urn:oasis:names:tc:opendocument:xmlns:table:1.0"
_NS_TEXT = u"urn:oasis:names:tc:opendocument:xmlns:text:1.0"


def write_ods(path, columns, rows):
    headers = _header_row(columns)

    def _cell_xml(value):
        return (
            u'<table:table-cell office:value-type="string">'
            u'<text:p>{0}</text:p></table:table-cell>'
        ).format(_xml_escape(_cell_text(value)))

    row_parts = []
    header_cells = u"".join(_cell_xml(h) for h in headers)
    row_parts.append(u"<table:table-row>{0}</table:table-row>".format(header_cells))
    for row in rows:
        cells = u"".join(_cell_xml(row.get(h, u"")) for h in headers)
        row_parts.append(u"<table:table-row>{0}</table:table-row>".format(cells))

    content_xml = XML_DECL + (
        u'<office:document-content '
        u'xmlns:office="{office}" xmlns:table="{table}" xmlns:text="{text}" '
        u'office:version="1.2">'
        u'<office:body><office:spreadsheet>'
        u'<table:table table:name="PyLink Export">{rows}</table:table>'
        u'</office:spreadsheet></office:body>'
        u'</office:document-content>'
    ).format(office=_NS_OFFICE, table=_NS_TABLE, text=_NS_TEXT, rows=u"".join(row_parts))

    manifest_xml = XML_DECL + (
        u'<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" '
        u'manifest:version="1.2">'
        u'<manifest:file-entry manifest:full-path="/" manifest:version="1.2" '
        u'manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>'
        u'<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>'
        u'</manifest:manifest>'
    )

    # mimetype must be the first entry, stored (uncompressed), no extra
    # attributes - LibreOffice/Excel both check this on open.
    zf = zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED)
    try:
        zf.writestr(zipfile.ZipInfo("mimetype"), b"application/vnd.oasis.opendocument.spreadsheet",
                    zipfile.ZIP_STORED)
        zf.writestr("META-INF/manifest.xml", manifest_xml.encode("utf-8"))
        zf.writestr("content.xml", content_xml.encode("utf-8"))
    finally:
        zf.close()


def read_ods(path):
    zf = zipfile.ZipFile(path, "r")
    try:
        root = ET.fromstring(zf.read("content.xml"))
        table = root.find(
            ".//{{{0}}}table".format(_NS_TABLE))
        if table is None:
            return [], []

        def _row_cells(row_el):
            values = []
            for cell_el in row_el.findall("{{{0}}}table-cell".format(_NS_TABLE)):
                p_el = cell_el.find("{{{0}}}p".format(_NS_TEXT))
                text = (p_el.text or u"") if p_el is not None else u""

                repeat = cell_el.get(
                    "{{{0}}}number-columns-repeated".format(_NS_TABLE))
                repeat_n = int(repeat) if repeat else 1

                if not text and repeat_n > 50:
                    # LibreOffice pads trailing empty cells with a huge
                    # repeat count (often 16384 = max columns) - that's
                    # end-of-row padding, not real data. Stop here.
                    break

                values.extend([text] * repeat_n)
            return values

        trs = table.findall("{{{0}}}table-row".format(_NS_TABLE))
        if not trs:
            return [], []

        headers = _row_cells(trs[0])
        rows_out = []
        for tr in trs[1:]:
            values = _row_cells(tr)
            if not any(values):
                continue
            rows_out.append(dict(zip(headers, values)))

        return headers, rows_out
    finally:
        zf.close()


# ── format router ────────────────────────────────────────────────────────

def write_workbook(path, columns, rows):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xlsx":
        write_xlsx(path, columns, rows)
    elif ext == ".ods":
        write_ods(path, columns, rows)
    else:
        raise ValueError("Unsupported export format: {0}".format(ext))


def read_workbook(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xlsx":
        return read_xlsx(path)
    elif ext == ".ods":
        return read_ods(path)
    else:
        raise ValueError("Unsupported import format: {0}".format(ext))
