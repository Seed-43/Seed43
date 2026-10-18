# -*- coding: utf-8 -*-
# """
# CSV and saved category groups for pyRameters' Parameters tab.
#
# No Revit and no WPF in here, so all of it is tested under desktop CPython.
# Written to run unchanged on IronPython 2 and CPython 3.
#
# WHY NOT THE csv MODULE
#     IronPython 2's csv module works in bytes and breaks on non-ASCII text,
#     and parameter names are full of it (degree signs, accented names). The
#     reader and writer below are small and handle what Excel produces:
#     quoted fields, doubled quotes, line breaks inside a quoted cell.
#
# WHAT EXCEL WRITES
#     "CSV UTF-8" is UTF-8 with a BOM; plain "CSV" is the ANSI code page, and
#     in some locales uses ';' instead of ','. The export writes UTF-8 with a
#     BOM so Excel opens accents correctly; the import accepts any of these.
#
# Target: IronPython 2 (pyRevit) and CPython 3.
# """

import io
import json
import os

try:
    unicode
except NameError:       # CPython 3
    unicode = str

__all__ = ["CSV_COLUMNS", "write_csv", "read_csv", "CsvError", "CategoryGroups"]

CSV_COLUMNS = [u"Name", u"Kind", u"Discipline", u"Type", u"Group under",
               u"Instance/Type", u"Value", u"Formula", u"Categories", u"GUID"]

CATEGORY_SEPARATOR = u"; "


class CsvError(Exception):
    pass


# ── WRITING ─────────────────────────────────────────────────────────────────

def _quote(text):
    text = u"" if text is None else unicode(text)
    needs = (u"," in text or u'"' in text or u"\n" in text or u"\r" in text
             or u";" in text or text != text.strip())
    if needs:
        return u'"' + text.replace(u'"', u'""') + u'"'
    return text


def write_csv(path, rows):
    """Write rows (dicts keyed by CSV_COLUMNS) as UTF-8 with a BOM."""
    lines = [u",".join(_quote(c) for c in CSV_COLUMNS)]
    for row in rows:
        lines.append(u",".join(_quote(row.get(c, u"")) for c in CSV_COLUMNS))
    with io.open(path, "w", encoding="utf-8-sig", newline="") as handle:
        handle.write(u"\r\n".join(lines) + u"\r\n")


# ── READING ─────────────────────────────────────────────────────────────────

def _decode(raw):
    for encoding in ("utf-8-sig", "cp1252"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1")


def _delimiter(first_line):
    counts = [(first_line.count(d), d) for d in (u",", u";", u"\t")]
    return max(counts)[1] if max(counts)[0] else u","


def _parse(text, delim):
    """Split CSV text into a list of rows of fields."""
    rows, row, field = [], [], []
    quoted = False
    i, n = 0, len(text)
    while i < n:
        c = text[i]
        if quoted:
            if c == u'"':
                if i + 1 < n and text[i + 1] == u'"':
                    field.append(u'"')
                    i += 1
                else:
                    quoted = False
            else:
                field.append(c)
        elif c == u'"' and not field:
            quoted = True
        elif c == delim:
            row.append(u"".join(field))
            field = []
        elif c in (u"\r", u"\n"):
            row.append(u"".join(field))
            field = []
            if any(cell.strip() for cell in row):
                rows.append(row)
            row = []
            if c == u"\r" and i + 1 < n and text[i + 1] == u"\n":
                i += 1
        else:
            field.append(c)
        i += 1
    row.append(u"".join(field))
    if any(cell.strip() for cell in row):
        rows.append(row)
    return rows


def read_csv(path):
    """Rows of a CSV as dicts keyed by CSV_COLUMNS (missing columns absent).

    Headers are matched case-insensitively, so a file edited in Excel and
    saved back reads the same. A Name column is the one requirement.
    """
    with io.open(path, "rb") as handle:
        text = _decode(handle.read())
    first = text.split(u"\n", 1)[0]
    rows = _parse(text, _delimiter(first))
    if not rows:
        raise CsvError(u"The file is empty.")
    known = dict((c.lower(), c) for c in CSV_COLUMNS)
    header = [known.get(h.strip().lower()) for h in rows[0]]
    if u"Name" not in header:
        raise CsvError(u"No Name column. The first row should be the column "
                       u"headings an export writes: " + u", ".join(CSV_COLUMNS))
    out = []
    for cells in rows[1:]:
        record = {}
        for index, column in enumerate(header):
            if column and index < len(cells):
                record[column] = cells[index].strip()
        if record.get(u"Name"):
            out.append(record)
    return out


# ── CATEGORY GROUPS ─────────────────────────────────────────────────────────

class CategoryGroups(object):
    """Named sets of categories, saved per user in a JSON file.

    Stored by category id with the names beside them for readability. The
    ids are what count: every bindable category is built in, and a built-in
    category's id is the same in every project.
    """

    def __init__(self, path):
        self.path = path
        self.groups = {}
        try:
            with io.open(path, encoding="utf-8") as handle:
                data = json.load(handle)
            for name, entry in (data.get("groups") or {}).items():
                ids = [int(i) for i in entry.get("ids", [])]
                if ids:
                    self.groups[name] = {"ids": ids,
                                         "names": entry.get("names", [])}
        except (IOError, OSError, ValueError, AttributeError, TypeError):
            self.groups = {}        # missing or unreadable: start empty

    def names(self):
        return sorted(self.groups, key=lambda n: n.lower())

    def ids(self, name):
        return list(self.groups.get(name, {}).get("ids", []))

    def save(self, name, ids, names):
        self.groups[name] = {"ids": sorted(int(i) for i in ids),
                             "names": sorted(names)}
        self._write()

    def delete(self, name):
        if self.groups.pop(name, None) is not None:
            self._write()

    def _write(self):
        """Write to one side and move into place, so a crash cannot leave
        half a file where the user's groups were."""
        folder = os.path.dirname(self.path)
        if folder and not os.path.isdir(folder):
            os.makedirs(folder)
        tmp = self.path + u".tmp"
        with io.open(tmp, "w", encoding="utf-8") as handle:
            handle.write(unicode(json.dumps({"groups": self.groups}, indent=2,
                                            sort_keys=True, ensure_ascii=False)))
        if os.path.exists(self.path):
            os.remove(self.path)
        os.rename(tmp, self.path)
