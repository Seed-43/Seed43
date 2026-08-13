# -*- coding: utf-8 -*-
"""
Shared xlsx/ods reader and writer.

Started life as pyTable's tools/pytable_io.py and moved here when pySheets
needed the same writer for its schedule export. Two consumers now:

  * pyTable   - exports model/schedule parameters, reads the edited file back.
  * pySheets  - writes one workbook per export, a tab per selected schedule.

Rows carry their tab name in a "_category" key, whatever that grouping
actually means to the caller (an element category for pyTable, a schedule
name for pySheets).

Deliberately has ZERO third-party dependencies. xlsx and ods are both just
zip archives of XML - this reads/writes that XML directly via zipfile and
xml.etree.ElementTree, both in the Python 2.7 stdlib IronPython ships. No
openpyxl, no odfpy, no xlsxwriter - pyTable's own xlsx code (pytable_excel.py)
independently arrived at the same zipfile-based approach, confirming it's
the right pattern for this environment (no pip under IronPython 2, so any
third-party package would have to be manually vendored).

Layout matches DiRoots SheetLink's structure - one sheet/table PER CATEGORY,
plus a Legend tab:
  1. "Legend"    - first tab, active on open, gridlines off (matches
                   SheetLink's own Instructions tab). Explains the column
                   colour coding, then lists a colour-coded jump link to
                   every category tab - link colour matches that tab's
                   colour.
  2. one sheet per category (e.g. "Walls", "Doors", ...) - normal
     gridlines, header row + Type/Read-only column tinting, tab colour
     matches its Legend link colour.

Column colours are deliberately different from DiRoots SheetLink's
(yellow/pink) while keeping the same idea - a distinct tint per category:
  Type columns:      soft teal      #CDEEEA
  Read-only columns: soft lavender  #DCE0F5
A column that's both Type and Read-only gets the Read-only tint (read-only
is the more important warning - don't edit this at all, regardless of why).

Category TAB colours are a separate, rotating palette (TAB_COLOR_PALETTE
below) - one distinct colour per category, cycling if there are more
categories than palette entries.

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
            "_category": "Walls",        # hidden - which tab this row goes on
            "Comments":  "some value",
            ...                          # one key per column["key"]
        }

Hidden columns (prefixed with "_") are always written first and are the
match key on import - never shown for editing, never diffed as a parameter.

On read, every sheet/table EXCEPT "Legend" is treated as data and merged
back into one row list (each row still carries "_category" from its
original tab) - there's no single fixed "data sheet name" any more now
that categories get their own tabs.

Lives in pyTable.pushbutton/tools/ alongside the rest of this tool's own
code - not in the shared lib/Snippets, since this I/O format is specific
to pyTable (same pattern as your existing pyTable's tools/pytable_excel.py).

ODS limitation: tab colour and Legend gridline removal are XLSX-only.
ODF has no standard per-sheet gridline attribute (LibreOffice treats grid
visibility as a whole-document view setting, not per-tab), and per-tab
colour is a LibreOffice extension (loext:tab-color) rather than a stable
ODF 1.2 feature, so it's skipped rather than risk depending on something
unconfirmed. ODS still gets per-category tables and colour-matched
hyperlink text on the Legend table.
"""

import os
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape as _xml_escape
from collections import OrderedDict

HIDDEN_COLUMNS = ("_eid", "_category")
LEGEND_SHEET_NAME = u"Legend"
LISTS_SHEET_NAME = u"Lists"
UNCATEGORISED_NAME = u"Uncategorised"

XML_DECL = u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

TYPE_FILL_HEX = u"CDEEEA"       # soft teal
READONLY_FILL_HEX = u"DCE0F5"   # soft lavender
NA_FILL_HEX = u"D9D9D9"         # soft grey - parameter not applicable to this element
CONSTRAINED_FILL_HEX = u"F5D9B8"  # soft peach - fixed set of values, but still user-editable via dropdown
HEADER_FILL_HEX = u"E8E8E8"     # light grey, header row

# One colour per category tab, cycling if there are more categories than
# entries. Kept distinct from the teal/lavender column tints above so
# there's no confusion between "column colour" and "tab colour".
TAB_COLOR_PALETTE = [
    u"C0392B",  # red
    u"2980B9",  # blue
    u"27AE60",  # green
    u"8E44AD",  # purple
    u"D35400",  # burnt orange
    u"16A085",  # teal (dark, distinct from the pale column teal)
    u"2C3E50",  # navy
    u"C2185B",  # magenta
    u"7F8C8D",  # grey
    u"F39C12",  # amber
    u"1ABC9C",  # turquoise
    u"E67E22",  # orange
]


# ── shared helpers ─────────────────────────────────────────────────────────

def _header_row(columns):
    """Return the full ordered header list: hidden key columns + visible params."""
    return list(HIDDEN_COLUMNS) + [c["key"] for c in columns]


def _cell_text(value):
    if value is None:
        return u""
    if not isinstance(value, unicode):
        value = unicode(value)
    # XML 1.0 only allows #x9, #xA, #xD, and [#x20-#xD7FF]/[#xE000-#xFFFD]/
    # [#x10000-#x10FFFF]. Real Revit data (notes, comments, family names)
    # can contain paste artifacts with other control characters that our
    # &/</> escaping doesn't cover - strip them rather than let them
    # silently corrupt the file.
    return u"".join(
        ch for ch in value
        if ch in u"\t\n\r" or (0x20 <= ord(ch) <= 0xD7FF) or (0xE000 <= ord(ch) <= 0xFFFD)
    )


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


def _column_style_kind(col):
    """Which fill a data column gets: 'constrained' (has a fixed dropdown
    of valid values) beats 'readonly' beats 'type' beats None.

    Constrained wins over Read-only deliberately: Revit's IsReadOnly flag
    is misleading for some enum-backed UI dropdown parameters (e.g. Detail
    Level) - they report read-only via the raw Parameter API even though
    they're genuinely user-editable through the dropdown Revit's own UI
    shows. If we can offer a list of valid values, treat it as editable
    within that set rather than locking it.
    """
    if col.get("options"):
        return "constrained"
    if col.get("readonly"):
        return "readonly"
    if col.get("kind") == "type":
        return "type"
    return None


def _cell_style_kind(col, value):
    """Per-cell fill: a missing/not-applicable value (None) always wins,
    since that's a stronger, more specific signal about THIS cell than the
    column's general classification."""
    if value is None:
        return "na"
    return _column_style_kind(col)


def _group_rows_by_category(rows):
    """Preserve first-seen order internally, but callers sort the category
    names themselves for tab order (alphabetical)."""
    groups = OrderedDict()
    for row in rows:
        cat = row.get("_category") or UNCATEGORISED_NAME
        groups.setdefault(cat, []).append(row)
    return groups


LEGEND_INTRO_ROWS = [
    (None, u"pyTable Export \u2014 Column Colour Legend"),
    ("type", u"Type Parameter \u2014 editing this value updates every element sharing this Type"),
    ("readonly", u"Read-only \u2014 shown for reference only, will not be written back on import"),
    ("na", u"Not Applicable \u2014 this parameter doesn't exist on this particular element"),
    ("constrained", u"Constrained \u2014 must be one of a fixed set of values (see the dropdown), "
                     u"but is still user-editable"),
    (None, u""),
    (None, u"Tip: check the Old/New values on the Preview/Edit diff before confirming a Type "
           u"parameter change \u2014 it applies to the whole Type, not just this row."),
    (None, u""),
    (None, u"To edit a locked (Read-only/Not Applicable) cell: in Excel, Review > Unprotect Sheet; "
           u"in LibreOffice Calc, Tools > Protect Sheet (untick). No password is set."),
    (None, u""),
    (None, u"Jump to category:"),
]


# ══ XLSX ═════════════════════════════════════════════════════════════════

_NS_MAIN = u"http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_RELS = u"http://schemas.openxmlformats.org/package/2006/relationships"
_NS_DOC_RELS = u"http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Fixed cellXfs indices (referenced via s="N" on <c> elements). Per-category
# hyperlink styles are appended dynamically after these, starting at
# _XF_FIRST_DYNAMIC. Combined with <sheetProtection> on each data sheet:
# Read-only/N-A/hidden-ID cells stay locked (the default), everything else
# is explicitly unlocked so it's still editable once the sheet is protected.
_XF_DEFAULT = 0
_XF_HEADER = 1
_XF_TYPE = 2
_XF_READONLY = 3
_XF_NA = 4
_XF_CONSTRAINED = 5
_XF_HIDDEN_LOCKED = 6
_XF_LEGEND_TITLE = 7
_XF_LEGEND_SWATCH_TYPE = 8
_XF_LEGEND_SWATCH_READONLY = 9
_XF_LEGEND_SWATCH_NA = 10
_XF_LEGEND_SWATCH_CONSTRAINED = 11
_XF_FIRST_DYNAMIC = 12

_STYLE_KIND_TO_XF = {
    "type": _XF_TYPE, "readonly": _XF_READONLY, "na": _XF_NA,
    "constrained": _XF_CONSTRAINED, None: _XF_DEFAULT,
}


def _sanitize_sheet_name(name, used_lower):
    """Excel sheet names: max 31 chars, no : \\ / ? * [ ] , can't be blank,
    can't collide (case-insensitive) with an existing name."""
    invalid = set(u':\\/?*[]')
    cleaned = u"".join(ch for ch in (name or u"") if ch not in invalid).strip()
    if not cleaned:
        cleaned = u"Sheet"
    cleaned = cleaned[:31]

    candidate = cleaned
    n = 2
    while candidate.lower() in used_lower:
        suffix = u" ({0})".format(n)
        candidate = cleaned[:31 - len(suffix)] + suffix
        n += 1
    used_lower.add(candidate.lower())
    return candidate


def _xlsx_escape_sheet_ref(sheet_name):
    """Sheet names with a single quote need it doubled inside a quoted
    reference, e.g. location="'It''s Walls'!A1"."""
    return sheet_name.replace(u"'", u"''")


def _xlsx_styles_xml(tab_colors):
    """tab_colors: ordered list of hex strings (no #), one per category -
    used to generate one extra coloured+underlined font/xf pair per
    category for the Legend's jump links."""
    dynamic_fonts = u"".join(
        u'<font><u/><sz val="11"/><color rgb="FF{0}"/><name val="Calibri"/></font>'.format(c)
        for c in tab_colors)
    dynamic_xfs = u"".join(
        u'<xf numFmtId="0" fontId="{0}" fillId="0" borderId="0" xfId="0" applyFont="1"/>'.format(2 + i)
        for i in range(len(tab_colors)))

    total_fonts = 2 + len(tab_colors)
    total_xfs = _XF_FIRST_DYNAMIC + len(tab_colors)

    return XML_DECL + (
        u'<styleSheet xmlns="{main}">'
        u'<fonts count="{total_fonts}">'
        u'<font><sz val="11"/><name val="Calibri"/></font>'
        u'<font><b/><sz val="11"/><name val="Calibri"/></font>'
        u'{dynamic_fonts}'
        u'</fonts>'
        u'<fills count="7">'
        u'<fill><patternFill patternType="none"/></fill>'
        u'<fill><patternFill patternType="gray125"/></fill>'
        u'<fill><patternFill patternType="solid"><fgColor rgb="FF{header}"/><bgColor indexed="64"/></patternFill></fill>'
        u'<fill><patternFill patternType="solid"><fgColor rgb="FF{type}"/><bgColor indexed="64"/></patternFill></fill>'
        u'<fill><patternFill patternType="solid"><fgColor rgb="FF{readonly}"/><bgColor indexed="64"/></patternFill></fill>'
        u'<fill><patternFill patternType="solid"><fgColor rgb="FF{na}"/><bgColor indexed="64"/></patternFill></fill>'
        u'<fill><patternFill patternType="solid"><fgColor rgb="FF{constrained}"/><bgColor indexed="64"/></patternFill></fill>'
        u'</fills>'
        u'<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        u'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        u'<cellXfs count="{total_xfs}">'
        u'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1"><protection locked="0"/></xf>'                      # 0 default (unlocked - editable)
        u'<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'                                          # 1 header (locked)
        u'<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1" applyProtection="1"><protection locked="0"/></xf>'        # 2 type col (unlocked - editable)
        u'<xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1" applyProtection="1"><protection locked="1"/></xf>'        # 3 readonly col (locked)
        u'<xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1" applyProtection="1"><protection locked="1"/></xf>'        # 4 na cell (locked)
        u'<xf numFmtId="0" fontId="0" fillId="6" borderId="0" xfId="0" applyFill="1" applyProtection="1"><protection locked="0"/></xf>'        # 5 constrained col (unlocked - editable within the dropdown)
        u'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1"><protection locked="1"/></xf>'                      # 6 hidden id column (locked)
        u'<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'                      # 7 legend title
        u'<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>'                      # 8 legend swatch: type
        u'<xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/>'                      # 9 legend swatch: readonly
        u'<xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1"/>'                      # 10 legend swatch: na
        u'<xf numFmtId="0" fontId="0" fillId="6" borderId="0" xfId="0" applyFill="1"/>'                      # 11 legend swatch: constrained
        u'{dynamic_xfs}'
        u'</cellXfs>'
        u'<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        u'</styleSheet>'
    ).format(main=_NS_MAIN, header=HEADER_FILL_HEX, type=TYPE_FILL_HEX, readonly=READONLY_FILL_HEX, na=NA_FILL_HEX,
             constrained=CONSTRAINED_FILL_HEX, total_fonts=total_fonts, total_xfs=total_xfs,
             dynamic_fonts=dynamic_fonts, dynamic_xfs=dynamic_xfs)


def _xlsx_cell(ref, value, style_idx=None):
    s_attr = u' s="{0}"'.format(style_idx) if style_idx else u""
    return u'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">{val}</t></is></c>'.format(
        ref=ref, s=s_attr, val=_xml_escape(_cell_text(value)))


def _xlsx_legend_sheet_xml(cat_names, cat_sheet_names, cat_row_counts, cat_xf_indices):
    rows_xml = []
    hyperlinks_xml = []
    r_idx = 0

    _SWATCH_XF = {
        "type": _XF_LEGEND_SWATCH_TYPE, "readonly": _XF_LEGEND_SWATCH_READONLY,
        "na": _XF_LEGEND_SWATCH_NA, "constrained": _XF_LEGEND_SWATCH_CONSTRAINED,
    }
    for r_idx, (kind, text) in enumerate(LEGEND_INTRO_ROWS, start=1):
        cells = []
        if kind is not None:
            swatch_style = _SWATCH_XF[kind]
            cells.append(_xlsx_cell(u"A{0}".format(r_idx), u"", swatch_style))
            cells.append(_xlsx_cell(u"B{0}".format(r_idx), text))
        else:
            style = _XF_LEGEND_TITLE if (r_idx == 1 or text == u"Jump to category:") else None
            cells.append(_xlsx_cell(u"A{0}".format(r_idx), text, style))
        rows_xml.append(u'<row r="{0}">{1}</row>'.format(r_idx, u"".join(cells)))

    link_start_row = r_idx + 1
    for i, cat in enumerate(cat_names):
        row_n = link_start_row + i
        sheet_name = cat_sheet_names[cat]
        xf_idx = cat_xf_indices[cat]
        ref = u"A{0}".format(row_n)
        cells = [
            _xlsx_cell(ref, cat, xf_idx),
            _xlsx_cell(u"B{0}".format(row_n), u"{0} row(s)".format(cat_row_counts[cat])),
        ]
        rows_xml.append(u'<row r="{0}">{1}</row>'.format(row_n, u"".join(cells)))
        hyperlinks_xml.append(
            u'<hyperlink ref="{ref}" location="\'{sheet}\'!A1" display="{disp}"/>'.format(
                ref=ref, sheet=_xlsx_escape_sheet_ref(sheet_name), disp=_xml_escape(cat)))

    cols_xml = u'<cols><col min="1" max="1" width="26" customWidth="1"/><col min="2" max="2" width="90" customWidth="1"/></cols>'
    hyperlinks_block = u"<hyperlinks>{0}</hyperlinks>".format(u"".join(hyperlinks_xml)) if hyperlinks_xml else u""

    return XML_DECL + (
        u'<worksheet xmlns="{main}" xmlns:r="{doc_ns}">'
        u'<sheetViews><sheetView tabSelected="1" showGridLines="0" workbookViewId="0"/></sheetViews>'
        u'{cols}<sheetData>{rows}</sheetData>{hyperlinks}</worksheet>'
    ).format(main=_NS_MAIN, doc_ns=_NS_DOC_RELS, cols=cols_xml, rows=u"".join(rows_xml), hyperlinks=hyperlinks_block)


def _xlsx_col_widths(headers, rows):
    """Rough auto-fit: width in characters = longest value seen in that
    column (header included), padded a little, capped to a sane range so
    one garbage-long value can't blow out the whole sheet."""
    widths = [len(h) for h in headers]
    for row in rows:
        for c_idx, h in enumerate(headers):
            val_len = len(_cell_text(row.get(h, u"")))
            if val_len > widths[c_idx]:
                widths[c_idx] = val_len
    return [min(max(w + 2, 8), 60) for w in widths]


def _xlsx_data_sheet_xml(columns, rows, tab_color_hex, dropdown_ranges):
    headers = _header_row(columns)
    col_meta = [{"kind": None, "readonly": False}] * len(HIDDEN_COLUMNS) + list(columns)
    col_widths = _xlsx_col_widths(headers, rows)

    cols_parts = []
    for c_idx, width in enumerate(col_widths):
        col_num = c_idx + 1
        if c_idx < len(HIDDEN_COLUMNS):
            cols_parts.append(u'<col min="{0}" max="{0}" width="{1}" hidden="1" customWidth="1"/>'.format(
                col_num, width))
        else:
            cols_parts.append(u'<col min="{0}" max="{0}" width="{1}" customWidth="1"/>'.format(
                col_num, width))
    cols_xml = u'<cols>{0}</cols>'.format(u"".join(cols_parts))

    header_cells = [
        _xlsx_cell(u"{0}1".format(_col_letter(c_idx)), header, _XF_HEADER)
        for c_idx, header in enumerate(headers)
    ]
    row_xml_parts = [u'<row r="1">{0}</row>'.format(u"".join(header_cells))]

    for r_idx, row in enumerate(rows, start=2):
        cells = []
        for c_idx, header in enumerate(headers):
            ref = u"{0}{1}".format(_col_letter(c_idx), r_idx)
            value = row.get(header, u"" if c_idx < len(HIDDEN_COLUMNS) else None)
            if c_idx < len(HIDDEN_COLUMNS):
                style_idx = _XF_HIDDEN_LOCKED
            else:
                style_idx = _STYLE_KIND_TO_XF[_cell_style_kind(col_meta[c_idx], value)]
            cells.append(_xlsx_cell(ref, value, style_idx))
        row_xml_parts.append(u'<row r="{0}">{1}</row>'.format(r_idx, u"".join(cells)))

    last_row = len(rows) + 1
    last_col_letter = _col_letter(len(headers) - 1)
    dimension_xml = u'<dimension ref="A1:{0}{1}"/>'.format(last_col_letter, last_row)
    # Freeze the header row so it stays visible while scrolling.
    sheet_views_xml = (
        u'<sheetViews><sheetView workbookViewId="0">'
        u'<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        u'<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>'
        u'</sheetView></sheetViews>'
    )
    autofilter_xml = u'<autoFilter ref="A1:{0}{1}"/>'.format(last_col_letter, last_row)
    validations_xml = _xlsx_data_validations_xml(headers, dropdown_ranges, last_row)
    # No password - a soft guard against accidental edits to Read-only/N-A/
    # hidden-ID cells, not real security. Users can still Unprotect Sheet
    # if they genuinely need to.
    protection_xml = u'<sheetProtection sheet="1" autoFilter="0" sort="0"/>'

    return XML_DECL + (
        u'<worksheet xmlns="{main}"><sheetPr><tabColor rgb="FF{tab_color}"/></sheetPr>'
        u'{dimension}{views}{cols}<sheetData>{rows}</sheetData>{protection}{autofilter}{validations}</worksheet>'
    ).format(main=_NS_MAIN, dimension=dimension_xml, views=sheet_views_xml, tab_color=tab_color_hex,
             cols=cols_xml, rows=u"".join(row_xml_parts), protection=protection_xml,
             autofilter=autofilter_xml, validations=validations_xml)


def _collect_dropdown_lists(columns):
    """Return an ordered {display_text: [options...]} for every column that
    has a non-empty options list, in a deterministic (sorted key) order."""
    result = OrderedDict()
    for key in sorted(c["key"] for c in columns if c.get("options")):
        col = next(c for c in columns if c["key"] == key)
        result[key] = col["options"]
    return result


def _xlsx_lists_sheet_xml(dropdown_lists):
    """One column per constrained parameter, values listed down the rows -
    referenced by each category sheet's <dataValidation> as
    Lists!$A$1:$A$N. Returns (sheet_xml, {display_text: (col_letter, count)})."""
    order = list(dropdown_lists.keys())
    ranges = {key: (_col_letter(i), len(dropdown_lists[key])) for i, key in enumerate(order)}
    lengths = [len(v) for v in dropdown_lists.values()]
    max_len = max(lengths) if lengths else 0

    row_parts = []
    for r_idx in range(1, max_len + 1):
        cells = []
        for c_idx, key in enumerate(order):
            opts = dropdown_lists[key]
            if r_idx <= len(opts):
                ref = u"{0}{1}".format(_col_letter(c_idx), r_idx)
                cells.append(_xlsx_cell(ref, opts[r_idx - 1]))
        if cells:
            row_parts.append(u'<row r="{0}">{1}</row>'.format(r_idx, u"".join(cells)))

    sheet_xml = XML_DECL + (
        u'<worksheet xmlns="{main}"><sheetData>{rows}</sheetData></worksheet>'
    ).format(main=_NS_MAIN, rows=u"".join(row_parts))
    return sheet_xml, ranges


def _xlsx_data_validations_xml(headers, dropdown_ranges, last_row):
    if not dropdown_ranges or last_row < 2:
        return u""
    parts = []
    for c_idx, header in enumerate(headers):
        if header not in dropdown_ranges:
            continue
        col_letter, opt_count = dropdown_ranges[header]
        sqref = u"{0}2:{0}{1}".format(_col_letter(c_idx), last_row)
        formula = u"{2}!${0}$1:${0}${1}".format(col_letter, opt_count, LISTS_SHEET_NAME)
        parts.append(
            u'<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="{sqref}">'
            u'<formula1>{formula}</formula1></dataValidation>'.format(sqref=sqref, formula=formula))
    if not parts:
        return u""
    return u'<dataValidations count="{0}">{1}</dataValidations>'.format(len(parts), u"".join(parts))


def write_xlsx(path, columns, rows):
    groups = _group_rows_by_category(rows)
    cat_names = sorted(groups.keys())

    used_lower = set([LEGEND_SHEET_NAME.lower()])
    cat_sheet_names = {}
    cat_tab_colors = {}
    cat_xf_indices = {}
    cat_row_counts = {}
    for i, cat in enumerate(cat_names):
        cat_sheet_names[cat] = _sanitize_sheet_name(cat, used_lower)
        cat_tab_colors[cat] = TAB_COLOR_PALETTE[i % len(TAB_COLOR_PALETTE)]
        cat_xf_indices[cat] = _XF_FIRST_DYNAMIC + i
        cat_row_counts[cat] = len(groups[cat])

    tab_colors_ordered = [cat_tab_colors[c] for c in cat_names]
    dropdown_lists = _collect_dropdown_lists(columns)

    content_types_parts = [
        u'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        u'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    ]
    root_rels = XML_DECL + (
        u'<Relationships xmlns="{ns}">'
        u'<Relationship Id="rId1" Type="{doc_ns}/officeDocument" Target="xl/workbook.xml"/>'
        u'</Relationships>'
    ).format(ns=_NS_RELS, doc_ns=_NS_DOC_RELS)

    sheets_xml = [u'<sheet name="{0}" sheetId="1" r:id="rId1"/>'.format(_xml_escape(LEGEND_SHEET_NAME))]
    workbook_rels_parts = [
        u'<Relationship Id="rId1" Type="{0}/worksheet" Target="worksheets/sheet1.xml"/>'.format(_NS_DOC_RELS)
    ]

    sheet_files = {
        "xl/worksheets/sheet1.xml": _xlsx_legend_sheet_xml(
            cat_names, cat_sheet_names, cat_row_counts, cat_xf_indices).encode("utf-8")
    }
    content_types_parts.append(
        u'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')

    dropdown_ranges = {}
    for i, cat in enumerate(cat_names):
        sheet_num = i + 2  # sheet1 = Legend
        r_id = u"rId{0}".format(sheet_num)
        sheets_xml.append(u'<sheet name="{0}" sheetId="{1}" r:id="{2}"/>'.format(
            _xml_escape(cat_sheet_names[cat]), sheet_num, r_id))
        workbook_rels_parts.append(
            u'<Relationship Id="{0}" Type="{1}/worksheet" Target="worksheets/sheet{2}.xml"/>'.format(
                r_id, _NS_DOC_RELS, sheet_num))
        sheet_path = u"xl/worksheets/sheet{0}.xml".format(sheet_num)
        sheet_files[sheet_path] = _xlsx_data_sheet_xml(
            columns, groups[cat], cat_tab_colors[cat], dropdown_ranges).encode("utf-8")
        content_types_parts.append(
            u'<Override PartName="/{0}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'.format(sheet_path))

    next_sheet_num = len(cat_names) + 2
    if dropdown_lists:
        lists_xml, dropdown_ranges_computed = _xlsx_lists_sheet_xml(dropdown_lists)
        dropdown_ranges.update(dropdown_ranges_computed)
        # Re-render category sheets now that dropdown_ranges is populated -
        # the first pass above ran before ranges existed, so redo it with
        # the real ranges available.
        for i, cat in enumerate(cat_names):
            sheet_num = i + 2
            sheet_path = u"xl/worksheets/sheet{0}.xml".format(sheet_num)
            sheet_files[sheet_path] = _xlsx_data_sheet_xml(
                columns, groups[cat], cat_tab_colors[cat], dropdown_ranges).encode("utf-8")

        lists_sheet_num = next_sheet_num
        r_id = u"rId{0}".format(lists_sheet_num)
        sheets_xml.append(u'<sheet name="{0}" sheetId="{1}" state="hidden" r:id="{2}"/>'.format(
            _xml_escape(LISTS_SHEET_NAME), lists_sheet_num, r_id))
        workbook_rels_parts.append(
            u'<Relationship Id="{0}" Type="{1}/worksheet" Target="worksheets/sheet{2}.xml"/>'.format(
                r_id, _NS_DOC_RELS, lists_sheet_num))
        lists_path = u"xl/worksheets/sheet{0}.xml".format(lists_sheet_num)
        sheet_files[lists_path] = lists_xml.encode("utf-8")
        content_types_parts.append(
            u'<Override PartName="/{0}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'.format(lists_path))
        next_sheet_num += 1

    styles_r_id = u"rId{0}".format(next_sheet_num)
    workbook_rels_parts.append(
        u'<Relationship Id="{0}" Type="{1}/styles" Target="styles.xml"/>'.format(styles_r_id, _NS_DOC_RELS))

    content_types = XML_DECL + (
        u'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        u'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        u'<Default Extension="xml" ContentType="application/xml"/>'
        u'{0}'
        u'</Types>'
    ).format(u"".join(content_types_parts))

    workbook_xml = XML_DECL + (
        u'<workbook xmlns="{main}" xmlns:r="{doc_ns}">'
        u'<bookViews><workbookView activeTab="0"/></bookViews>'
        u'<sheets>{sheets}</sheets>'
        u'</workbook>'
    ).format(main=_NS_MAIN, doc_ns=_NS_DOC_RELS, sheets=u"".join(sheets_xml))

    workbook_rels = XML_DECL + (
        u'<Relationships xmlns="{ns}">{rels}</Relationships>'
    ).format(ns=_NS_RELS, rels=u"".join(workbook_rels_parts))

    zf = zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED)
    try:
        zf.writestr("[Content_Types].xml", content_types.encode("utf-8"))
        zf.writestr("_rels/.rels", root_rels.encode("utf-8"))
        zf.writestr("xl/workbook.xml", workbook_xml.encode("utf-8"))
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels.encode("utf-8"))
        zf.writestr("xl/styles.xml", _xlsx_styles_xml(tab_colors_ordered).encode("utf-8"))
        for sheet_path, content in sheet_files.items():
            zf.writestr(sheet_path, content)
    finally:
        zf.close()


def _xlsx_list_data_sheet_paths(zf):
    """Return [(sheet_name, part_path), ...] for every sheet EXCEPT Legend
    and the hidden Lists helper sheet (dropdown option values, not data)."""
    wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
    sheets = wb_root.findall("{{{0}}}sheets/{{{0}}}sheet".format(_NS_MAIN))
    rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_targets = {rel.get("Id"): rel.get("Target") for rel in rels_root}

    result = []
    for sheet_el in sheets:
        name = sheet_el.get("name")
        if name in (LEGEND_SHEET_NAME, LISTS_SHEET_NAME):
            continue
        r_id = sheet_el.get("{{{0}}}id".format(_NS_DOC_RELS))
        target = rel_targets.get(r_id)
        if not target:
            continue
        path = "xl/" + target if not target.startswith("/") else target.lstrip("/")
        result.append((name, path))
    return result


def read_xlsx(path):
    zf = zipfile.ZipFile(path, "r")
    try:
        shared_strings = []
        if "xl/sharedStrings.xml" in zf.namelist():
            sst_root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in sst_root:
                texts = [t.text or u"" for t in si.iter("{{{0}}}t".format(_NS_MAIN))]
                shared_strings.append(u"".join(texts))

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
            return v_el.text

        def _read_sheet(sheet_path):
            sheet_root = ET.fromstring(zf.read(sheet_path))
            sheet_data = sheet_root.find("{{{0}}}sheetData".format(_NS_MAIN))
            if sheet_data is None:
                return [], []

            sheet_headers = None
            sheet_rows = []
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

                if sheet_headers is None:
                    sheet_headers = ordered
                    continue

                if not any(v for v in ordered):
                    continue
                sheet_rows.append(dict(zip(sheet_headers, ordered)))

            return (sheet_headers or []), sheet_rows

        all_headers = []
        all_rows = []
        for _sheet_name, sheet_path in _xlsx_list_data_sheet_paths(zf):
            headers, rows = _read_sheet(sheet_path)
            if headers and len(headers) > len(all_headers):
                all_headers = headers
            all_rows.extend(rows)

        return all_headers, all_rows
    finally:
        zf.close()


# ══ ODS ══════════════════════════════════════════════════════════════════

_NS_OFFICE = u"urn:oasis:names:tc:opendocument:xmlns:office:1.0"
_NS_TABLE = u"urn:oasis:names:tc:opendocument:xmlns:table:1.0"
_NS_TEXT = u"urn:oasis:names:tc:opendocument:xmlns:text:1.0"
_NS_STYLE = u"urn:oasis:names:tc:opendocument:xmlns:style:1.0"
_NS_FO = u"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
_NS_XLINK = u"http://www.w3.org/1999/xlink"


def _ods_cell_xml(value, style_name=None, content_validation_name=None):
    s_attr = u' table:style-name="{0}"'.format(style_name) if style_name else u""
    v_attr = u' table:content-validation-name="{0}"'.format(content_validation_name) if content_validation_name else u""
    return (
        u'<table:table-cell office:value-type="string"{s}{v}>'
        u'<text:p>{val}</text:p></table:table-cell>'
    ).format(s=s_attr, v=v_attr, val=_xml_escape(_cell_text(value)))


def _ods_hyperlink_cell_xml(text, target_table_name, style_name, cell_style_name=None):
    """Internal jump link to A1 of another table, ODF form:
    <text:a xlink:href="#TableName.A1">text</text:a>, wrapped in a styled p.
    cell_style_name (table-cell family) is separate from style_name (the
    text/link colour) - one controls the cell background, the other the
    link's font colour."""
    href = u"#{0}.A1".format(target_table_name.replace(u"'", u"&apos;"))
    cell_attr = u' table:style-name="{0}"'.format(cell_style_name) if cell_style_name else u""
    return (
        u'<table:table-cell office:value-type="string"{cell_attr}>'
        u'<text:p><text:a xlink:type="simple" xlink:href="{href}" text:style-name="{style}">'
        u'{val}</text:a></text:p></table:table-cell>'
    ).format(cell_attr=cell_attr, href=_xml_escape(href), style=style_name, val=_xml_escape(_cell_text(text)))


def _sanitize_table_name(name, used_lower):
    """ODS table names can't contain ' [ ] / \\ * ? : or start/end with a
    space, and can't collide (case-insensitive)."""
    invalid = set(u"'[]/\\*?:")
    cleaned = u"".join(ch for ch in (name or u"") if ch not in invalid).strip()
    if not cleaned:
        cleaned = u"Table"

    candidate = cleaned
    n = 2
    while candidate.lower() in used_lower:
        candidate = u"{0} ({1})".format(cleaned, n)
        n += 1
    used_lower.add(candidate.lower())
    return candidate


def write_ods(path, columns, rows):
    headers = _header_row(columns)
    groups = _group_rows_by_category(rows)
    cat_names = sorted(groups.keys())

    used_lower = set([LEGEND_SHEET_NAME.lower()])
    cat_table_names = {}
    cat_colors = {}
    cat_text_style_names = {}
    for i, cat in enumerate(cat_names):
        cat_table_names[cat] = _sanitize_table_name(cat, used_lower)
        cat_colors[cat] = TAB_COLOR_PALETTE[i % len(TAB_COLOR_PALETTE)]
        cat_text_style_names[cat] = u"Link{0}".format(i)

    col_meta = [{"kind": None, "readonly": False}] * len(HIDDEN_COLUMNS) + list(columns)
    _ODS_STYLE_NAME = {
        "type": u"ColType", "readonly": u"ColReadonly", "na": u"ColNA",
        "constrained": u"ColConstrained", None: u"ColEditable",
    }

    validation_names = {}
    for i, col in enumerate(columns):
        if col.get("options"):
            validation_names[col["key"]] = u"valList{0}".format(i)

    dynamic_link_styles = u"".join(
        u'<style:style style:name="{0}" style:family="text">'
        u'<style:text-properties fo:color="#{1}" style:text-underline-style="solid"/>'
        u'</style:style>'.format(cat_text_style_names[c], cat_colors[c])
        for c in cat_names
    )

    automatic_styles = (
        u'<style:style style:name="ColHeader" style:family="table-cell">'
        u'<style:table-cell-properties fo:background-color="#{0}" style:cell-protect="protected"/>'
        u'<style:text-properties fo:font-weight="bold"/></style:style>'
        u'<style:style style:name="ColType" style:family="table-cell">'
        u'<style:table-cell-properties fo:background-color="#{1}" style:cell-protect="none"/></style:style>'
        u'<style:style style:name="ColReadonly" style:family="table-cell">'
        u'<style:table-cell-properties fo:background-color="#{2}" style:cell-protect="protected"/></style:style>'
        u'<style:style style:name="ColNA" style:family="table-cell">'
        u'<style:table-cell-properties fo:background-color="#{3}" style:cell-protect="protected"/></style:style>'
        u'<style:style style:name="ColConstrained" style:family="table-cell">'
        u'<style:table-cell-properties fo:background-color="#{4}" style:cell-protect="none"/></style:style>'
        u'<style:style style:name="ColEditable" style:family="table-cell">'
        u'<style:table-cell-properties style:cell-protect="none"/></style:style>'
        u'<style:style style:name="ColHiddenLocked" style:family="table-cell">'
        u'<style:table-cell-properties style:cell-protect="protected"/></style:style>'
        u'<style:style style:name="LegendDefault" style:family="table-cell">'
        u'<style:table-cell-properties fo:background-color="#FFFFFF"/></style:style>'
        u'<style:style style:name="LegendTitle" style:family="table-cell">'
        u'<style:table-cell-properties fo:background-color="#FFFFFF"/>'
        u'<style:text-properties fo:font-weight="bold"/></style:style>'
        u'{5}'
    ).format(HEADER_FILL_HEX, TYPE_FILL_HEX, READONLY_FILL_HEX, NA_FILL_HEX, CONSTRAINED_FILL_HEX, dynamic_link_styles)

    # ── Legend table ──
    _SWATCH_STYLE = {"type": u"ColType", "readonly": u"ColReadonly", "na": u"ColNA", "constrained": u"ColConstrained"}
    legend_rows = []
    for kind, text in LEGEND_INTRO_ROWS:
        if kind is not None:
            cells = _ods_cell_xml(u"", _SWATCH_STYLE[kind]) + _ods_cell_xml(text, u"LegendDefault")
        else:
            is_title = (text == LEGEND_INTRO_ROWS[0][1] or text == u"Jump to category:")
            cells = _ods_cell_xml(text, u"LegendTitle" if is_title else u"LegendDefault")
        legend_rows.append(u"<table:table-row>{0}</table:table-row>".format(cells))

    for cat in cat_names:
        link_cell = _ods_hyperlink_cell_xml(
            cat, cat_table_names[cat], cat_text_style_names[cat], cell_style_name=u"LegendDefault")
        count_cell = _ods_cell_xml(u"{0} row(s)".format(len(groups[cat])), u"LegendDefault")
        legend_rows.append(u"<table:table-row>{0}{1}</table:table-row>".format(link_cell, count_cell))

    # Column defaults only paint cells within DECLARED rows - extend with a
    # large block of empty repeated rows so the white background actually
    # covers the visible page area, not just the ~10 rows of real content.
    legend_rows.append(u'<table:table-row table:number-rows-repeated="500"/>')

    legend_columns = u'<table:table-column table:number-columns-repeated="30" table:default-cell-style-name="LegendDefault"/>'
    legend_table = u'<table:table table:name="{0}">{1}{2}</table:table>'.format(
        _xml_escape(LEGEND_SHEET_NAME), legend_columns, u"".join(legend_rows))

    # ── Per-category tables ──
    category_tables = []
    database_ranges = []
    for cat in cat_names:
        header_cells = u"".join(_ods_cell_xml(h, u"ColHeader") for h in headers)
        data_rows = [u"<table:table-row>{0}</table:table-row>".format(header_cells)]
        for row in groups[cat]:
            cells = []
            for c_idx, h in enumerate(headers):
                value = row.get(h, u"" if c_idx < len(HIDDEN_COLUMNS) else None)
                if c_idx < len(HIDDEN_COLUMNS):
                    style_name = u"ColHiddenLocked"
                    validation_name = None
                else:
                    style_name = _ODS_STYLE_NAME[_cell_style_kind(col_meta[c_idx], value)]
                    validation_name = validation_names.get(h)
                cells.append(_ods_cell_xml(value, style_name, validation_name))
            data_rows.append(u"<table:table-row>{0}</table:table-row>".format(u"".join(cells)))
        table_name = cat_table_names[cat]
        category_tables.append(u'<table:table table:name="{0}" table:protected="true">{1}</table:table>'.format(
            _xml_escape(table_name), u"".join(data_rows)))

        last_col_letter = _col_letter(len(headers) - 1)
        last_row = len(groups[cat]) + 1
        escaped_name = table_name.replace(u"'", u"''")
        database_ranges.append(
            u'<table:database-range table:name="{safe_name}" '
            u'table:target-range-address="\'{name}\'.A1:\'{name}\'.{last_col}{last_row}" '
            u'table:display-filter-buttons="true"/>'.format(
                safe_name=_xml_escape(u"DB_" + u"".join(ch if ch.isalnum() else u"_" for ch in table_name)),
                name=escaped_name, last_col=last_col_letter, last_row=last_row))

    database_ranges_xml = u"<table:database-ranges>{0}</table:database-ranges>".format(
        u"".join(database_ranges)) if database_ranges else u""

    content_validations_parts = []
    for col in columns:
        opts = col.get("options")
        if not opts:
            continue
        val_name = validation_names[col["key"]]
        # ODF list-from-values validation: of:cell-content-is-in-list("A";"B";"C").
        # table:base-cell-address is REQUIRED - LibreOffice silently strips the
        # whole condition without it. table:display-list="unsorted" is what
        # renders the dropdown arrow; allow-empty-cell only restricts typed
        # input and adds no UI affordance.
        quoted_opts = u";".join(
            u"&quot;{0}&quot;".format(_xml_escape(o).replace(u'"', u"&quot;")) for o in opts)
        condition = u"of:cell-content-is-in-list({0})".format(quoted_opts)
        content_validations_parts.append(
            u'<table:content-validation table:name="{name}" table:condition="{cond}" '
            u'table:base-cell-address="{base}" table:allow-empty-cell="true" '
            u'table:display-list="unsorted">'
            u'<table:error-message table:display="true" table:message-type="stop"/>'
            u'</table:content-validation>'.format(
                name=val_name, cond=condition, base=_xml_escape(LEGEND_SHEET_NAME) + u".A1"))
    content_validations_xml = (
        u"<table:content-validations>{0}</table:content-validations>".format(
            u"".join(content_validations_parts)) if content_validations_parts else u"")

    content_xml = XML_DECL + (
        u'<office:document-content '
        u'xmlns:office="{office}" xmlns:table="{table}" xmlns:text="{text}" '
        u'xmlns:style="{style}" xmlns:fo="{fo}" xmlns:xlink="{xlink}" '
        u'office:version="1.2">'
        u'<office:automatic-styles>{styles}</office:automatic-styles>'
        u'<office:body><office:spreadsheet>{validations}{legend}{categories}{dbranges}</office:spreadsheet></office:body>'
        u'</office:document-content>'
    ).format(office=_NS_OFFICE, table=_NS_TABLE, text=_NS_TEXT, style=_NS_STYLE, fo=_NS_FO, xlink=_NS_XLINK,
             styles=automatic_styles, validations=content_validations_xml,
             legend=legend_table, categories=u"".join(category_tables),
             dbranges=database_ranges_xml)

    manifest_xml = XML_DECL + (
        u'<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" '
        u'manifest:version="1.2">'
        u'<manifest:file-entry manifest:full-path="/" manifest:version="1.2" '
        u'manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>'
        u'<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>'
        u'<manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/>'
        u'</manifest:manifest>'
    )

    # Freeze the header row per category table so it survives scrolling. ODF
    # keeps this as a per-table view setting in settings.xml, not content.xml.
    # TODO: VerticalSplitMode=2 + ActiveSplitRange=2 is the assumed encoding of
    # "frozen at row 1" - unverified, confirm it renders frozen on open.
    table_view_entries = u"".join(
        u'<config:config-item-map-entry config:name="{name}">'
        u'<config:config-item config:name="VerticalSplitMode" config:type="short">2</config:config-item>'
        u'<config:config-item config:name="VerticalSplitPosition" config:type="int">1</config:config-item>'
        u'<config:config-item config:name="ActiveSplitRange" config:type="short">2</config:config-item>'
        u'<config:config-item config:name="PositionTop" config:type="int">0</config:config-item>'
        u'<config:config-item config:name="PositionBottom" config:type="int">1</config:config-item>'
        u'</config:config-item-map-entry>'.format(name=_xml_escape(cat_table_names[cat]))
        for cat in cat_names
    )
    settings_xml = XML_DECL + (
        u'<office:document-settings xmlns:office="{office}" xmlns:config="{config}" '
        u'office:version="1.2">'
        u'<office:settings>'
        u'<config:config-item-set config:name="ooo:view-settings">'
        u'<config:config-item-map-indexed config:name="Views">'
        u'<config:config-item-map-entry>'
        u'<config:config-item-map-named config:name="Tables">{tables}</config:config-item-map-named>'
        u'</config:config-item-map-entry>'
        u'</config:config-item-map-indexed>'
        u'</config:config-item-set>'
        u'</office:settings>'
        u'</office:document-settings>'
    ).format(office=_NS_OFFICE, config=u"urn:oasis:names:tc:opendocument:xmlns:config:1.0",
             tables=table_view_entries)

    zf = zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED)
    try:
        zf.writestr(zipfile.ZipInfo("mimetype"), b"application/vnd.oasis.opendocument.spreadsheet",
                    zipfile.ZIP_STORED)
        zf.writestr("META-INF/manifest.xml", manifest_xml.encode("utf-8"))
        zf.writestr("content.xml", content_xml.encode("utf-8"))
        zf.writestr("settings.xml", settings_xml.encode("utf-8"))
    finally:
        zf.close()


def read_ods(path):
    zf = zipfile.ZipFile(path, "r")
    try:
        root = ET.fromstring(zf.read("content.xml"))
        tables = root.findall(".//{{{0}}}table".format(_NS_TABLE))

        def _row_cells(row_el):
            values = []
            for cell_el in row_el.findall("{{{0}}}table-cell".format(_NS_TABLE)):
                p_el = cell_el.find("{{{0}}}p".format(_NS_TEXT))
                text = u""
                if p_el is not None:
                    a_el = p_el.find("{{{0}}}a".format(_NS_TEXT))
                    if a_el is not None:
                        text = u"".join(a_el.itertext())
                    else:
                        text = p_el.text or u""

                repeat = cell_el.get("{{{0}}}number-columns-repeated".format(_NS_TABLE))
                repeat_n = int(repeat) if repeat else 1

                if not text and repeat_n > 50:
                    break

                values.extend([text] * repeat_n)
            return values

        def _read_table(t_el):
            trs = t_el.findall("{{{0}}}table-row".format(_NS_TABLE))
            if not trs:
                return [], []
            t_headers = _row_cells(trs[0])
            t_rows = []
            for tr in trs[1:]:
                values = _row_cells(tr)
                if not any(values):
                    continue
                t_rows.append(dict(zip(t_headers, values)))
            return t_headers, t_rows

        all_headers = []
        all_rows = []
        for t_el in tables:
            if t_el.get("{{{0}}}name".format(_NS_TABLE)) == LEGEND_SHEET_NAME:
                continue
            headers, rows = _read_table(t_el)
            if headers and len(headers) > len(all_headers):
                all_headers = headers
            all_rows.extend(rows)

        return all_headers, all_rows
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
