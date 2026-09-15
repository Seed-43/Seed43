# -*- coding: utf-8 -*-
from pyrevit import revit, DB
from pyrevit import forms
from pyrevit import script
from Snippets import _userdata

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
from System.Windows.Shapes import Ellipse
from System.Windows.Media import SolidColorBrush, Color

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
pylink_shared.py -- genuinely cross-cutting pieces used by BOTH the
Excel and Word sides of pyLink: the Row class, colour/status
constants, the hb() colour helper, and Revit shared-parameter state
persistence (save/load). Deliberately has NO dependency on
pylink_excel.py or pylink_word.py, so both of those (and the main
pyLink.py) can import from here with zero circularity risk.
"""


MM     = 1.0 / 304.8   # millimetres to Revit internal feet
PT_MM  = 0.352778      # millimetres per point (1pt = 1/72in = 25.4/72mm) -
                        # this is the answer to "what is 7pt in mm": 7 * PT_MM ~ 2.47mm


def _col_letter_to_index(col_str):
    """Spreadsheet column letters ("A", "AB", ...) to a 0-based index -
    shared by the xlsx and ods format readers (tools/format/), both of
    which resolve named-range cell references the same way."""
    result = 0
    for ch in col_str.upper():
        result = result * 26 + (ord(ch) - ord('A') + 1)
    return result - 1


VIEW_TYPES      = ['Schedule View', 'Legend View', 'Drafting View']
WORD_VIEW_TYPES = ['Legend View', 'Drafting View']
SHEET_SIZES     = [
    'A4 Landscape', 'A4 Portrait',
    'A3 Landscape', 'A3 Portrait',
    'A2 Landscape', 'A2 Portrait',
    'A1 Landscape', 'A1 Portrait',
    'A0 Landscape', 'A0 Portrait',
]

SRC_COLOURS    = {'xl': '#217346', 'word': '#2B579A', 'ods': '#0E8C7B', 'odt': '#6B3FA0'}
STATUS_COLOURS = {
    'pending': '#6B7280', 'success': '#16A34A',
    'error':   '#DC2626', 'skipped': '#CA8A04',
    'sync':    '#3B82F6',
}

PYLINK_PARAM_GUID = 'f0a46d4c-c148-4ff4-95c8-9750eec5d480'
PYLINK_PARAM_NAME = 'pyLink'
PYLINK_PARAM_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), '..', 'pyLink.txt')

# ── Default text font (shared infra, not Excel-reading logic - the
# Excel side just asks "is this font installed / what's the fallback",
# it doesn't own the setting) ──
# Stored in .user so an update cannot overwrite it. The old userdata/ folder
# beside the tool is migrated across on first run; the extra folder level is
# dropped, since .user/pyLink/ is already the userdata folder.
EXCEL_FONT_SETTINGS_PATH = _userdata.migrate(
    os.path.join(os.path.dirname(os.path.abspath(__file__)), '..',
                 'userdata', 'excel_font_settings.json'),
    _userdata.user_path('pyLink', 'excel_font_settings.json'))

DEFAULT_EXCEL_FONT_SETTINGS = {'fallback_font': 'Arial', 'force_fallback': False}


def load_excel_font_settings():
    """The font pyLink falls back to when a table's own Excel font
    (e.g. 'Aptos Narrow') isn't actually installed on this machine -
    user-configurable via the hamburger menu's 'Default Font' option."""
    try:
        if os.path.exists(EXCEL_FONT_SETTINGS_PATH):
            with open(EXCEL_FONT_SETTINGS_PATH, 'r') as f:
                data = _json.load(f)
            if isinstance(data, dict) and 'fallback_font' in data:
                return data
    except Exception as ex:
        logger.warning('excel_font_settings.json load failed: {}'.format(ex))
    return dict(DEFAULT_EXCEL_FONT_SETTINGS)


def save_excel_font_settings(data):
    try:
        folder = os.path.dirname(EXCEL_FONT_SETTINGS_PATH)
        if not os.path.exists(folder):
            os.makedirs(folder)
        with open(EXCEL_FONT_SETTINGS_PATH, 'w') as f:
            _json.dump(data, f, indent=2)
    except Exception as ex:
        logger.warning('excel_font_settings.json save failed: {}'.format(ex))


# ── Per-project schedule text sizes ──────────────────────────────────────────
# A schedule's text height can't come from the spreadsheet: Excel points
# describe text on a screen, not text on a sheet at a plot scale, so
# importing 10pt literally gives an unreadable schedule. pyLink asks
# instead, once per project, and remembers the answer.
#
# The questions are per KIND of row, not per row. Rows are keyed by how
# Excel formats them, so a 14pt bold title, the 10pt bold header rows and
# the plain 10pt data rows are three questions - and the two header rows
# of a grouped table share one answer because they're formatted alike.
#
# Stored in .user so a tool update can't overwrite it, matching
# EXCEL_FONT_SETTINGS_PATH above.
SCHEDULE_TEXT_SIZES_PATH = _userdata.user_path(
    'pyLink', 'schedule_text_sizes.json')

DEFAULT_SCHEDULE_TEXT_SIZE_MM = 2.5
MIN_SCHEDULE_TEXT_SIZE_MM = 0.5
MAX_SCHEDULE_TEXT_SIZE_MM = 20.0


def load_schedule_text_sizes():
    """The whole store: {'projects': {<key>: {'title':.., 'sizes':{..}}}}."""
    try:
        if os.path.exists(SCHEDULE_TEXT_SIZES_PATH):
            with open(SCHEDULE_TEXT_SIZES_PATH, 'r') as f:
                data = _json.load(f)
            if isinstance(data, dict) and isinstance(data.get('projects'), dict):
                return data
    except Exception as ex:
        logger.warning('schedule_text_sizes.json load failed: {}'.format(ex))
    return {'projects': {}}


def save_schedule_text_sizes(data):
    try:
        folder = os.path.dirname(SCHEDULE_TEXT_SIZES_PATH)
        if not os.path.exists(folder):
            os.makedirs(folder)
        with open(SCHEDULE_TEXT_SIZES_PATH, 'w') as f:
            _json.dump(data, f, indent=2)
    except Exception as ex:
        logger.warning('schedule_text_sizes.json save failed: {}'.format(ex))


def _safe_doc_title(doc):
    try:
        return unicode(doc.Title)
    except Exception:
        return u'unknown'


def project_key(doc):
    """A stable id for the open project.

    ProjectInformation's UniqueId survives Save As and is the same in a
    workshared local copy as in the central, so everyone on a job shares
    one set of sizes. Title is only stored alongside to keep the JSON
    readable by eye.
    """
    try:
        uid = doc.ProjectInformation.UniqueId
        if uid:
            return unicode(uid)
    except Exception:
        pass
    return _safe_doc_title(doc)


def row_style_signature(cell_style):
    """Key that groups rows Excel formats identically.

    Two rows only need one question between them if they would look the
    same, so bold, italic and the Excel point size make up the key. This
    is why row 3 of a grouped table isn't asked about separately - it is
    bold at the same size as row 2, so it reuses that answer.
    """
    if not cell_style:
        return 'plain'
    try:
        size = round(float(cell_style.get('font_size', 11.0)), 1)
    except (TypeError, ValueError):
        size = 11.0
    return 'b{0}|i{1}|s{2}'.format(
        1 if cell_style.get('bold') else 0,
        1 if cell_style.get('italic') else 0,
        size)


def get_project_text_sizes(doc):
    """The sizes already answered for this project, {signature: mm}."""
    store = load_schedule_text_sizes()
    entry = store['projects'].get(project_key(doc), {})
    sizes = entry.get('sizes', {})
    return sizes if isinstance(sizes, dict) else {}


def store_project_text_size(doc, signature, size_mm):
    store = load_schedule_text_sizes()
    key = project_key(doc)
    entry = store['projects'].get(key)
    if not isinstance(entry, dict):
        entry = {}
        store['projects'][key] = entry
    entry['title'] = _safe_doc_title(doc)
    if not isinstance(entry.get('sizes'), dict):
        entry['sizes'] = {}
    entry['sizes'][signature] = size_mm
    save_schedule_text_sizes(store)


def store_schedule_row_sizes(doc, view_name, row_size_mm, row_overrides_mm=None,
                             wrap_cells=None):
    """Record the text height actually used for each row of a schedule.

    Refit needs to know how tall the text is to work out how much room
    it takes. Reading TableCellStyle.TextSize back and converting it is
    guesswork - the value isn't in millimetres and the conversion has
    changed between Revit versions - so the build writes down what it
    used and refit reads that instead.

    Which cells WRAP is written down for the same reason: Revit's
    TableCellStyle doesn't carry the flag, so without this a later
    "Reset Cell Sizes" would size every heading for one line and undo
    the layout the import produced.
    """
    store = load_schedule_text_sizes()
    key = project_key(doc)
    entry = store['projects'].get(key)
    if not isinstance(entry, dict):
        entry = {}
        store['projects'][key] = entry
    entry['title'] = _safe_doc_title(doc)
    if not isinstance(entry.get('schedules'), dict):
        entry['schedules'] = {}
    entry['schedules'][unicode(view_name)] = {
        'text': dict((str(ri), mm) for ri, mm in row_size_mm.items()),
        # Row heights the spreadsheet asked for. Recorded so that
        # "Reset Cell Sizes" puts the schedule back to how the import
        # left it, rather than to something the import never produced.
        'rows': dict((str(ri), mm)
                     for ri, mm in (row_overrides_mm or {}).items()),
        # Cells the spreadsheet wrapped, as "row,col".
        'wrap': ['%d,%d' % (r, c) for r, c in sorted(wrap_cells or ())],
    }
    save_schedule_text_sizes(store)


def _schedule_record(doc, view_name, part):
    store = load_schedule_text_sizes()
    entry = store['projects'].get(project_key(doc), {})
    saved = (entry.get('schedules') or {}).get(unicode(view_name))
    if not isinstance(saved, dict):
        return {}
    # Records written before this held the text sizes directly.
    block = saved.get(part) if 'text' in saved or 'rows' in saved else (
        saved if part == 'text' else {})
    if not isinstance(block, dict):
        return {}
    out = {}
    for ri, mm in block.items():
        try:
            out[int(ri)] = float(mm)
        except (TypeError, ValueError):
            continue
    return out


def get_schedule_row_sizes(doc, view_name):
    """The recorded per-row text heights for one schedule, {row: mm}."""
    return _schedule_record(doc, view_name, 'text')


def get_schedule_row_overrides(doc, view_name):
    """The row heights the spreadsheet asked for, {row: mm}."""
    return _schedule_record(doc, view_name, 'rows')


def get_schedule_wrap_cells(doc, view_name):
    """The cells the spreadsheet wrapped, as a set of (row, col).

    Empty for a schedule imported before this was recorded, which just
    means refit falls back to its old one-line-per-heading sizing.
    """
    store = load_schedule_text_sizes()
    entry = store['projects'].get(project_key(doc), {})
    saved = (entry.get('schedules') or {}).get(unicode(view_name))
    if not isinstance(saved, dict):
        return set()
    out = set()
    for item in saved.get('wrap') or ():
        try:
            r, c = str(item).split(',')
            out.add((int(r), int(c)))
        except (TypeError, ValueError):
            continue
    return out


def reset_project_text_sizes(doc):
    """Forget this project's answers so the next run asks again."""
    store = load_schedule_text_sizes()
    if store['projects'].pop(project_key(doc), None) is None:
        return False
    save_schedule_text_sizes(store)
    return True


def _row_sample(row, limit=64):
    """The row's own text, for showing the user which row is being asked."""
    parts = []
    for cell in row:
        if cell is None:
            continue
        try:
            text = unicode(cell).strip()
        except Exception:
            continue
        if text:
            parts.append(text)
    text = u' | '.join(parts)
    if len(text) > limit:
        text = text[:limit - 1] + u'…'
    return text or u'(empty row)'


def _ask_text_size(row_index, sample, default_mm):
    """Ask for one row-kind's text height, during a build.

    None if the user cancels.
    """
    prompt = (u'Row {0}:\n{1}\n\n'
              u'Text height in mm for every row formatted like this one.'
              .format(row_index + 1, sample))
    return ask_size_mm(prompt, default_mm)


def ask_size_mm(prompt, default_mm):
    """Ask for a text height in mm, validated. None if cancelled.

    Uses the themed Seed43 input dialog when it's available, the same
    way _alert and _confirm do, so this matches the rest of the tool.
    Split out of _ask_text_size so the settings editor can reuse the
    same prompt and the same bounds instead of growing its own.
    """
    default = '{0:g}'.format(default_mm)
    title = 'pyLink - Excel Text Size'
    error = ''
    while True:
        if sdlg:
            answer = sdlg.ask_string(prompt, title=title,
                                     default=default, error=error)
        else:
            answer = forms.ask_for_string(default=default, prompt=prompt,
                                          title=title)
        text = unicode(answer).strip() if answer is not None else u''
        if not text:
            return None                      # cancelled or left blank
        try:
            value = float(text)
        except (TypeError, ValueError):
            error = 'Enter a number, e.g. 2.5'
            continue
        if MIN_SCHEDULE_TEXT_SIZE_MM <= value <= MAX_SCHEDULE_TEXT_SIZE_MM:
            return value
        error = 'Enter a size between {0:g} and {1:g} mm.'.format(
            MIN_SCHEDULE_TEXT_SIZE_MM, MAX_SCHEDULE_TEXT_SIZE_MM)


def describe_signature(signature):
    """'b1|i0|s14.0' -> 'Bold 14pt rows', for the settings dialog."""
    if signature == 'plain':
        return u'Unformatted rows'
    try:
        parts = {}
        for chunk in signature.split('|'):
            if chunk:
                parts[chunk[0]] = chunk[1:]
        bits = []
        if parts.get('b') == '1':
            bits.append(u'Bold')
        if parts.get('i') == '1':
            bits.append(u'Italic')
        size = parts.get('s', '?')
        try:
            size = u'{0:g}'.format(float(size))   # 14.0 -> 14
        except (TypeError, ValueError):
            pass
        bits.append(u'{0}pt'.format(size))
        return u' '.join(bits) + u' rows'
    except Exception:
        return signature


def resolve_row_text_sizes(doc, all_rows, cell_styles, ask=True):
    """Return {row_index: text height in mm} for every row in the table.

    Asks the user once per distinct row style, in the order the styles
    first appear, and remembers each answer against this project so the
    next schedule in the same job runs without prompting. Cancelling a
    prompt accepts the default for that style rather than aborting the
    whole export.
    """
    sizes = dict(get_project_text_sizes(doc))
    row_signature = {}
    first_seen = []
    seen = set()

    for ri, row in enumerate(all_rows):
        style = None
        for ci in range(len(row)):
            cell = row[ci]
            if cell is None or not unicode(cell).strip():
                continue
            style = cell_styles.get((ri, ci))
            if style:
                break
        signature = row_style_signature(style)
        row_signature[ri] = signature
        if signature not in seen:
            seen.add(signature)
            first_seen.append((signature, ri))

    for signature, ri in first_seen:
        if signature in sizes:
            continue
        if not ask:
            sizes[signature] = DEFAULT_SCHEDULE_TEXT_SIZE_MM
            continue
        answer = _ask_text_size(ri, _row_sample(all_rows[ri]),
                                DEFAULT_SCHEDULE_TEXT_SIZE_MM)
        if answer is None:
            answer = DEFAULT_SCHEDULE_TEXT_SIZE_MM
        sizes[signature] = answer
        store_project_text_size(doc, signature, answer)

    resolved = {}
    for ri in row_signature:
        resolved[ri] = sizes.get(row_signature[ri],
                                 DEFAULT_SCHEDULE_TEXT_SIZE_MM)
    return resolved


# ── Autofit: size cells to the text they actually hold ───────────────────────
# Excel's own column widths and row heights are no use once the user has
# chosen their own text height - a column sized for 10pt Arial is far too
# wide for 2mm Revit text. So the cells are measured from the real text at
# the real size instead, with WPF doing the measuring.
#
# Padding is deliberately a little generous: text that overflows a Revit
# schedule cell is clipped, and a slightly loose column is much easier to
# live with than a truncated one. Tighten these if the result is airy.
AUTOFIT_PAD_W_MM = 2.0      # added to every column
AUTOFIT_PAD_H_MM = 1.2      # added to every row
AUTOFIT_MIN_COL_MM = 6.0
AUTOFIT_MIN_ROW_MM = 4.0

# Refit only. A rotated label needs a row as long as its text, which for
# a header like "Number: Support (npsup) [No]" runs away to something
# that swamps the sheet. Past this the row stops growing and the text is
# allowed to clip - raise it if you'd rather have the height.
REFIT_MAX_ROW_MM = 45.0
REFIT_MAX_COL_MM = 60.0

# What a line break inside a schedule cell looks like. Revit accepts one -
# Shift+Enter puts it there by hand - and that is the only way to get two
# lines out of a cell, because a schedule does NOT reflow text to fit the
# column the way a spreadsheet does: it draws one line and ellipsises the
# rest, however tall the row is. So anything that has to wrap is broken
# here, before it is written, at the point the spreadsheet broke it.
CELL_LINE_BREAK = u'\n'


def split_cell_lines(text):
    """The lines of a cell, however its breaks were written."""
    if text is None:
        return [u'']
    return unicode(text).replace(u'\r\n', u'\n').replace(u'\r', u'\n').split(u'\n')


def cell_line_count(text):
    """How many lines a cell's text already carries."""
    return len(split_cell_lines(text))


def wrap_text_to_lines(text, font_name, size_mm, available_mm,
                       bold=False, italic=False):
    """Break text into lines that each fit available_mm.

    Greedy and space-only, which is how a spreadsheet wraps: words fill a
    line until the next will not fit. A word wider than the space still
    gets a line to itself rather than being split, because breaking
    inside a word is what turns "DESIGN CAPACITY" into "DESI GN CAP".

    Any break already in the text is honoured and wrapped within.
    """
    out = []
    for para in split_cell_lines(text):
        words = para.split()
        if not words:
            out.append(u'')
            continue
        line = u''
        for word in words:
            trial = word if not line else line + u' ' + word
            width, _h = measure_text_mm(trial, font_name, size_mm, bold, italic)
            if line and width > available_mm:
                out.append(line)
                line = word
            else:
                line = trial
        out.append(line)
    return out


_MEASURE_CACHE = {}


def _typeface(font_name, bold, italic):
    from System.Windows.Media import Typeface, FontFamily
    from System.Windows import FontStyles, FontWeights, FontStretches
    return Typeface(
        FontFamily(font_name or 'Arial'),
        FontStyles.Italic if italic else FontStyles.Normal,
        FontWeights.Bold if bold else FontWeights.Normal,
        FontStretches.Normal)


def measure_text_mm(text, font_name, size_mm, bold=False, italic=False):
    """(width, line height) of one line of text, in mm.

    Measured with WPF's FormattedText at an em size of size_mm, so the
    numbers come back in the same units the caller is working in. Falls
    back to a character-count estimate if the measurement isn't
    available for any reason - a loose column beats a crash.
    """
    key = (text, font_name, round(float(size_mm), 3), bool(bold), bool(italic))
    cached = _MEASURE_CACHE.get(key)
    if cached:
        return cached

    # A cell that carries its own breaks is as wide as its WIDEST LINE,
    # not as wide as the whole string run together. Height stays one
    # line's worth - callers multiply by the line count themselves, so
    # that a cell measured here and a cell counted with cell_line_count
    # agree.
    if text and (u'\n' in text or u'\r' in text):
        widest, line_h = 0.0, 0.0
        for line in split_cell_lines(text):
            w, h = measure_text_mm(line, font_name, size_mm, bold, italic)
            widest = max(widest, w)
            line_h = max(line_h, h)
        result = (widest, line_h)
        _MEASURE_CACHE[key] = result
        return result

    result = None
    try:
        from System.Windows.Media import FormattedText, Brushes
        from System.Windows import FlowDirection
        from System.Globalization import CultureInfo
        face = _typeface(font_name, bold, italic)
        try:
            # .NET 4.6+ overload; the older one has no pixelsPerDip.
            ft = FormattedText(text, CultureInfo.InvariantCulture,
                               FlowDirection.LeftToRight, face,
                               float(size_mm), Brushes.Black, 1.0)
        except TypeError:
            ft = FormattedText(text, CultureInfo.InvariantCulture,
                               FlowDirection.LeftToRight, face,
                               float(size_mm), Brushes.Black)
        result = (float(ft.Width), float(ft.Height))
    except Exception as ex:
        logger.debug('text measure failed, estimating instead: {}'.format(ex))

    if result is None:
        result = (len(text) * size_mm * 0.6, size_mm * 1.35)

    _MEASURE_CACHE[key] = result
    return result


def longest_word_mm(text, font_name, size_mm, bold=False, italic=False):
    """Width of the widest unbreakable run in the text.

    Wrapping can only break at spaces, so a column narrower than this
    clips the word instead of wrapping it - which is exactly how
    "DESIGN CAPACITY" ends up rendered as "DESIGN CAPA...". Any width
    worked out by dividing a string across N lines has to be floored by
    this, or the arithmetic promises a wrap that can't happen.
    """
    widest = 0.0
    for word in text.split():
        width, _line_h = measure_text_mm(word, font_name, size_mm,
                                         bold, italic)
        if width > widest:
            widest = width
    return widest


def autofit_table(all_rows, cell_styles, row_size_mm, n_cols, merges=(),
                  default_font='Arial'):
    """Column widths and row heights in mm, measured from the real text.

    Three kinds of cell contribute differently:

      * rotated  - the text runs up the cell, so its LENGTH drives the
        row height and only its line height drives the column width.
        This is what keeps a rotated header column narrow.
      * wrapping - doesn't stretch its column to fit the whole string;
        it wraps onto more lines instead, and those lines drive the row
        height. It still can't go narrower than its longest WORD though,
        since wrapping only breaks at spaces.
      * merged   - spans several columns, so it can't stretch just one.
        Any shortfall is shared out across the columns it covers.

    Everything else drives its column's width and its row's height
    directly.
    """
    col_mm, row_mm = {}, {}
    merge_span, merge_need, wrap_cells = {}, {}, {}

    spanned = set()
    for r1, c1, r2, c2 in merges:
        merge_span[(r1, c1)] = (r2, c2)
        for rr in range(r1, r2 + 1):
            for cc in range(c1, c2 + 1):
                if (rr, cc) != (r1, c1):
                    spanned.add((rr, cc))

    for ri, row in enumerate(all_rows):
        size = row_size_mm.get(ri, DEFAULT_SCHEDULE_TEXT_SIZE_MM)
        for ci in range(min(len(row), n_cols)):
            if (ri, ci) in spanned:
                continue
            cell = row[ci]
            text = u'' if cell is None else unicode(cell).strip()
            if not text:
                continue

            style = cell_styles.get((ri, ci), {})
            width, line_h = measure_text_mm(
                text, style.get('font_name', default_font), size,
                style.get('bold'), style.get('italic'))

            rotated = bool(style.get('rotation'))
            n_lines = cell_line_count(text)
            if rotated:
                need_w, need_h = line_h * n_lines, width
            else:
                need_w, need_h = width, line_h * n_lines

            span_cols = None
            if (ri, ci) in merge_span:
                last = min(merge_span[(ri, ci)][1], n_cols - 1)
                span_cols = list(range(ci, last + 1))

            # Wrapping wins over merging: a wrapping title spread across
            # every column should use more lines, not force the whole
            # table wider to fit on one.
            #
            # Rotated text is excluded because it can't wrap in a Revit
            # schedule (see refit_schedule): its length has already set
            # the row height above, and sending it down this path would
            # divide that length by the column WIDTH and ask for a row
            # of that many lines.
            if style.get('wrap') and not rotated and n_lines == 1:
                cols = span_cols or [ci]
                wrap_cells[(ri, ci)] = (width, line_h, cols)
                # Wrapping still can't split a word, so the columns it
                # covers have to total at least the longest one.
                merge_need[('wrap', ri, ci)] = (
                    longest_word_mm(text, style.get('font_name',
                                                    default_font),
                                    size, style.get('bold'),
                                    style.get('italic')),
                    cols)
            elif span_cols:
                merge_need[(ri, ci)] = (need_w, span_cols)
            else:
                col_mm[ci] = max(col_mm.get(ci, 0.0), need_w)

            row_mm[ri] = max(row_mm.get(ri, 0.0), need_h)

    for ci in range(n_cols):
        col_mm[ci] = max(AUTOFIT_MIN_COL_MM,
                         col_mm.get(ci, 0.0) + AUTOFIT_PAD_W_MM)

    # A merged header still has to fit somewhere: if the columns it
    # covers don't add up, widen them all a little rather than one a lot.
    for need_w, span in merge_need.values():
        shortfall = need_w + AUTOFIT_PAD_W_MM - sum(col_mm[c] for c in span)
        if shortfall > 0:
            share = shortfall / len(span)
            for c in span:
                col_mm[c] += share

    # Now the columns are known, a wrapping cell's line count is too.
    # A merged wrapping cell wraps across everything it spans.
    for (ri, _ci), (width, line_h, span) in wrap_cells.items():
        available = sum(col_mm.get(c, AUTOFIT_MIN_COL_MM)
                        for c in span) - AUTOFIT_PAD_W_MM
        if available <= 0:
            lines = 1
        else:
            lines = int(width / available)
            if width % available:
                lines += 1
        row_mm[ri] = max(row_mm.get(ri, 0.0), max(1, lines) * line_h)

    for ri in range(len(all_rows)):
        row_mm[ri] = max(AUTOFIT_MIN_ROW_MM,
                         row_mm.get(ri, 0.0) + AUTOFIT_PAD_H_MM)

    return col_mm, row_mm


MM_PER_FT = 304.8

# TableCellStyle.TextSize is NOT millimetres. create_schedule.py writes it
# as  size_pt * 1.1812, and size_pt is size_mm * 72/25.4 - so what goes in
# (and comes back out) is the text height in mm multiplied by this:
REVIT_TEXT_SIZE_FACTOR = (72.0 / 25.4) * 1.1812      # ~3.348
#
# Reading it back as if it were mm makes every string measure ~3.3x too
# wide, which blows every column out to the cap. Undo the factor on read.
# Realistic schedule text is a few mm, so the result is sanity-checked
# and the raw value used instead if dividing gives something absurd -
# that way a Revit version whose getter already converts still works.
PLAUSIBLE_TEXT_MM = (0.8, 12.0)


def _resolve_text_size_mm(raw):
    """Turn a TableCellStyle.TextSize reading into millimetres."""
    try:
        raw = float(raw)
    except (TypeError, ValueError):
        return None
    if raw <= 0:
        return None
    low, high = PLAUSIBLE_TEXT_MM
    scaled = raw / REVIT_TEXT_SIZE_FACTOR
    if low <= scaled <= high:
        return scaled
    if low <= raw <= high:
        return raw
    return None


def _cell_style_info(sec, r, c, default_font='Arial'):
    """Font, size (mm) and rotation of one Revit schedule cell.

    Every read is guarded: a cell that has never been styled returns a
    TableCellStyle with unset members, and asking for one of those
    throws rather than returning a default.
    """
    info = {'font_name': default_font, 'bold': False, 'italic': False,
            'size_mm': DEFAULT_SCHEDULE_TEXT_SIZE_MM, 'rotated': False}
    try:
        style = sec.GetTableCellStyle(r, c)
    except Exception:
        return info
    try:
        if style.FontName:
            info['font_name'] = style.FontName
    except Exception:
        pass
    try:
        info['bold'] = bool(style.IsFontBold)
    except Exception:
        pass
    try:
        info['italic'] = bool(style.IsFontItalic)
    except Exception:
        pass
    try:
        size = _resolve_text_size_mm(style.TextSize)
        if size:
            info['size_mm'] = size
    except Exception:
        pass
    try:
        info['rotated'] = int(style.TextOrientation) in (90, 270)
    except Exception:
        pass
    return info


REFIT_RESET = 'reset'
REFIT_CURRENT = 'current'


def refit_schedule(doc, schedule, default_font='Arial', mode=REFIT_RESET,
                   row_size_mm=None, wrap_cells=None):
    """Size a schedule's cells so nothing is truncated.

    Two modes, because there are two things you might want protected:

    REFIT_RESET ("Reset Cell Sizes") ignores whatever the cells are now
    and sizes them from the text alone. Horizontal text WIDENS ITS
    COLUMN to fit on one line, deliberately rather than wrapping down to
    a sliver - a group header over one narrow column is what turns
    "DESIGN CAPACITY" into "DESI GN CAP ACI...", and a wider column is
    the fix. Use it to start again from a clean layout.

    REFIT_CURRENT ("Update Cells After Resizing") keeps the sizes you set
    in Revit and adjusts the other dimension to suit them. Each row's
    height decides how many lines it can hold, so a row you made taller
    lets its columns become narrower. Row heights are kept as you left
    them and only grow if something still won't fit.

    Rotated text can't wrap in either mode, so it always needs a column
    one line wide and a row as long as the text - which runs away on a
    long label, hence REFIT_MAX_ROW_MM.

    A cell the SPREADSHEET wrapped is the exception to RESET's widening:
    it only has to fit its longest word, and takes as many lines as that
    leaves it needing. Without it a merged "DETAILING CONSTANTS", or the
    title across the whole table, held every column it spanned open at
    its full one-line width, and the schedule came out much wider than
    the sheet it was imported from. Revit's TableCellStyle doesn't carry
    a wrap flag, so the set comes from the caller or, failing that, from
    what the import wrote down.

    Must be called inside a transaction. Returns (n_cols, n_rows).
    """
    section = schedule.GetTableData().GetSectionData(DB.SectionType.Header)
    n_rows = section.NumberOfRows
    n_cols = section.NumberOfColumns

    # What the build wrote down beats anything read back off the cells.
    if row_size_mm is None:
        try:
            row_size_mm = get_schedule_row_sizes(doc, schedule.Name)
        except Exception:
            row_size_mm = {}
    if wrap_cells is None:
        try:
            wrap_cells = get_schedule_wrap_cells(doc, schedule.Name)
        except Exception:
            wrap_cells = set()

    current_row_mm = {}
    for r in range(n_rows):
        try:
            current_row_mm[r] = section.GetRowHeight(r) * MM_PER_FT
        except Exception:
            current_row_mm[r] = AUTOFIT_MIN_ROW_MM

    cells = []
    sizes_seen = set()
    for r in range(n_rows):
        for c in range(n_cols):
            try:
                text = section.GetCellText(r, c)
            except Exception:
                continue
            text = unicode(text).strip() if text else u''
            if not text:
                continue

            span = 1
            try:
                merged = section.GetMergedCell(r, c)
                if merged and merged.Right > merged.Left:
                    if r != merged.Top or c != merged.Left:
                        continue            # only the anchor carries the text
                    span = merged.Right - merged.Left + 1
            except Exception:
                pass

            info = _cell_style_info(section, r, c, default_font)
            size_mm = row_size_mm.get(r) or info['size_mm']
            width, line_h = measure_text_mm(
                text, info['font_name'], size_mm,
                info['bold'], info['italic'])
            word_w = longest_word_mm(text, info['font_name'], size_mm,
                                     info['bold'], info['italic'])
            sizes_seen.add(round(size_mm, 2))
            cells.append((r, c, span, width, line_h, info['rotated'], word_w,
                          cell_line_count(text)))

    # If these aren't a few millimetres, REVIT_TEXT_SIZE_FACTOR is wrong
    # for this Revit version and every column will come out proportionally
    # off - this line is the first thing to check when that happens.
    logger.debug('refit "{}" mode={} recorded={} text sizes (mm): {}'.format(
        schedule.Name, mode, bool(row_size_mm), sorted(sizes_seen)))

    # ── Column widths ──
    # RESET: wide enough for horizontal text on one line.
    # CURRENT: only as wide as the text needs given the lines the row's
    # existing height already allows, so a taller row buys a narrower
    # column.
    col_mm, merged_needs = {}, []
    for r, c, span, width, line_h, rotated, word_w, n_lines in cells:
        if rotated:
            need = line_h * n_lines
        elif n_lines > 1:
            # Already broken into lines, so `width` is the widest of
            # them: fit that and the cell is whole. Nothing here has to
            # hope Revit will reflow it, because it will not.
            need = width
        elif (r, c) in wrap_cells:
            # The spreadsheet wrapped it but it arrived on one line, so
            # the longest word is the floor - anything narrower clips.
            need = word_w
        elif mode == REFIT_CURRENT and line_h > 0:
            lines = int(current_row_mm.get(r, 0.0) / line_h)
            # Never narrower than the longest word, or the wrap this
            # division assumes can't actually happen and Revit clips.
            need = max(width / max(1, lines), word_w)
        else:
            need = width
        if span > 1:
            merged_needs.append((c, span, need))
        else:
            col_mm[c] = max(col_mm.get(c, 0.0), need)

    for c in range(n_cols):
        col_mm[c] = min(REFIT_MAX_COL_MM,
                        max(AUTOFIT_MIN_COL_MM,
                            col_mm.get(c, 0.0) + AUTOFIT_PAD_W_MM))

    for c, span, need in merged_needs:
        span_cols = list(range(c, min(c + span, n_cols)))
        if not span_cols:
            continue
        shortfall = need + AUTOFIT_PAD_W_MM - sum(col_mm[x] for x in span_cols)
        if shortfall > 0:
            share = shortfall / len(span_cols)
            for x in span_cols:
                col_mm[x] += share

    # ── Row heights at those widths ──
    row_mm = {}
    for r, c, span, width, line_h, rotated, _word_w, n_lines in cells:
        if rotated:
            # Rotated text runs up the cell, so its LENGTH is the height
            # it needs - per line, since the lines sit side by side.
            need = width
        elif n_lines > 1:
            # Broken text takes exactly the lines it was broken into.
            # No division, no guessing: the height is the line count
            # times the height of one line at this row's text size.
            need = n_lines * line_h
        else:
            span_cols = list(range(c, min(c + max(1, span), n_cols)))
            available = sum(col_mm[x] for x in span_cols) - AUTOFIT_PAD_W_MM
            if available <= 0:
                lines = 1
            else:
                lines = int(width / available)
                if width % available:
                    lines += 1
            need = max(1, lines) * line_h
        row_mm[r] = max(row_mm.get(r, 0.0), need)

    if mode == REFIT_RESET:
        # "Reset" means back to how the import left it, so a row height
        # the spreadsheet asked for is part of the target, not something
        # to recompute away.
        try:
            overrides = get_schedule_row_overrides(doc, schedule.Name)
        except Exception:
            overrides = {}
    else:
        overrides = {}

    for r in range(n_rows):
        fitted = min(REFIT_MAX_ROW_MM,
                     max(AUTOFIT_MIN_ROW_MM,
                         row_mm.get(r, 0.0) + AUTOFIT_PAD_H_MM))
        if mode == REFIT_CURRENT:
            # Your height is the input, not something to overrule - only
            # grow past it when the text genuinely doesn't fit.
            fitted = max(current_row_mm.get(r, 0.0), fitted)
        elif r in overrides:
            # The spreadsheet's height is a target, never a ceiling. It
            # was measured around 10pt Excel text and scaled down for
            # the smaller Revit text, so it lands close but not exact -
            # and a row pinned a fraction under what the text needs is a
            # row Revit truncates: "DESIGN CAPACITY" came back as
            # "DESIGN...". Floor it at the fit, which is the same rule
            # create_schedule.py applies when it writes the height in
            # the first place.
            fitted = max(fitted, overrides[r])
        row_mm[r] = fitted

    for c in range(n_cols):
        try:
            section.SetColumnWidth(c, col_mm[c] / MM_PER_FT)
        except Exception as ex:
            logger.debug('refit SetColumnWidth({}): {}'.format(c, ex))
    for r in range(n_rows):
        try:
            section.SetRowHeight(r, row_mm[r] / MM_PER_FT)
        except Exception as ex:
            logger.debug('refit SetRowHeight({}): {}'.format(r, ex))

    return n_cols, n_rows


def get_installed_font_names():
    """Every font family name Windows/WPF actually has installed,
    sorted - used both to validate a table's own font before using it,
    and to populate the fallback-font picker dropdown."""
    names = []
    try:
        from System.Windows.Media import Fonts
        for ff in Fonts.SystemFontFamilies:
            try:
                # FontFamily.Source is "Family Name", or occasionally
                # "file://path/#Family Name" for a font referenced by
                # file - either way the real name is after any '#'.
                src = str(ff.Source)
                names.append(src.split('#')[-1].strip())
            except Exception:
                continue
    except Exception as ex:
        logger.warning('Reading installed fonts failed: {}'.format(ex))
    return sorted(set(n for n in names if n))


def _is_font_installed(font_name):
    """Whether font_name is actually installed on this machine. Revit
    silently substitutes an uninstalled font rather than erroring, so
    without this check a table styled in a font the drafter doesn't
    have (e.g. 'Aptos Narrow') would render in whatever Revit happened
    to pick instead - unpredictable rather than a deliberate, user-
    chosen fallback."""
    if not font_name:
        return False
    target = font_name.strip().lower()
    for n in get_installed_font_names():
        if n.strip().lower() == target:
            return True
    return False


def _alert(message, title='', exitscript=False):
    """Themed popup via the shared Snippets dialog lib, falls back to
    pyRevit's default forms.alert if the shared lib isn't available."""
    if sdlg:
        sdlg.message(message, title=title)
    else:
        forms.alert(message, title=title)
    if exitscript:
        script.exit()

def _confirm(message, title='', yes='Yes', no='No'):
    """Themed yes/no popup, returns True on yes."""
    if sdlg:
        return sdlg.confirm(message, title=title, yes=yes, no=no)
    return bool(forms.alert(message, title=title, ok=False, yes=True, no=True))

# ── Constants ──
VIEW_TYPE_LEGEND   = 'Legend View'
VIEW_TYPE_SCHEDULE = 'Schedule View'
VIEW_TYPE_DRAFTING = 'Drafting View'


# ── Data class ──

def _run_export_script(script_name, payload):
    """
    Run one of the Export/ scripts via exec() with PYLINK_PAYLOAD injected.
    Wraps execution in a transaction since export scripts modify the document.
    Legend script manages its own transactions internally so is run without
    the outer wrapper to avoid nesting.
    """
    export_dir  = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'Export')
    script_path = os.path.join(export_dir, script_name)

    if not os.path.exists(script_path):
        raise Exception(
            'Export script not found: {}'.format(script_path)
        )

    ns = {
        '__name__':        script_name,
        '__file__':        script_path,
        '__builtins__':    __builtins__,
        'PYLINK_PAYLOAD': payload,
    }

    src = open(script_path, 'r').read()

    # These scripts manage their own transactions internally
    if script_name in ('create_legend.py', 'create_notes.py'):
        exec(src, ns)
    else:
        with revit.Transaction(
            'pyLink - {}'.format(
                payload.get('view_name', script_name)
            )
        ):
            exec(src, ns)

    # Export scripts may leave a PYLINK_RESULT dict in their own
    # namespace (e.g. {'view_id': view.Id.IntegerValue}) so the caller
    # can tag the actual view that was created — used for the
    # ElementId-based ownership tracking.
    return ns.get('PYLINK_RESULT')

def _ensure_pylink_param_bound():
    """
    Ensure the pyLink shared parameter is bound to BOTH Project
    Information (the punch list of everything pyLink manages) and
    Views (Drafting/Legend/Schedule all fall under this one binding
    category) — same parameter definition, one independent value slot
    per element instance. Expands an existing binding in place if it
    was only ever bound to Project Information from an earlier version.
    Returns the Definition object, or None on failure.
    """
    try:
        app = revit.HOST_APP.app
        orig_file = app.SharedParametersFilename
        try:
            app.SharedParametersFilename = PYLINK_PARAM_FILE
            sp_file = app.OpenSharedParameterFile()
            if sp_file is None:
                return None
            grp = None
            for g in sp_file.Groups:
                if g.Name == 'Seed43':
                    grp = g
                    break
            if grp is None:
                return None
            defn = None
            for d in grp.Definitions:
                if d.Name == PYLINK_PARAM_NAME:
                    defn = d
                    break
            if defn is None:
                return None

            wanted_cats = [
                doc.Settings.Categories.get_Item(
                    DB.BuiltInCategory.OST_ProjectInformation),
                doc.Settings.Categories.get_Item(
                    DB.BuiltInCategory.OST_Views),
            ]

            existing_binding = doc.ParameterBindings.get_Item(defn)
            with revit.Transaction('pyLink - bind parameter'):
                if existing_binding is None:
                    cats = DB.CategorySet()
                    for c in wanted_cats:
                        cats.Insert(c)
                    binding = DB.InstanceBinding(cats)
                    doc.ParameterBindings.Insert(defn, binding)
                else:
                    existing_cats = existing_binding.Categories
                    changed = False
                    for c in wanted_cats:
                        if not existing_cats.Contains(c):
                            existing_cats.Insert(c)
                            changed = True
                    if changed:
                        doc.ParameterBindings.ReInsert(defn, existing_binding)
            return defn
        finally:
            app.SharedParametersFilename = orig_file
    except Exception as ex:
        logger.warning('pyLink param bind failed: {}'.format(ex))
        return None

def get_view_pylink_data(view):
    """
    Read a view's own pyLink record (sheet/range/hash/scale/etc) from
    its per-view parameter — the authoritative record for that one
    view, as opposed to Project Information's punch list, which is
    just an index of which views to go look up. Returns a dict of
    whatever tokens are present, or None if the view has no record
    (never touched by pyLink, or the parameter can't be read).
    """
    try:
        p = view.LookupParameter(PYLINK_PARAM_NAME)
        if p is None:
            return None
        text = p.AsString()
        if not text:
            return None
        parts = {}
        for seg in text.split('|'):
            if '-' in seg:
                k, v = seg.split('-', 1)
                parts[k] = v
        return parts
    except Exception:
        return None

def set_view_pylink_data(view, **kwargs):
    """
    Write this view's own pyLink record. Keyword args become K-V
    tokens on the view's per-view parameter, e.g.
    set_view_pylink_data(view, SH='Sheet1', RG='Temp', H=hash_str).
    Creates/expands the parameter binding on first use. Returns True
    on success.
    """
    try:
        p = view.LookupParameter(PYLINK_PARAM_NAME)
        if p is None:
            _ensure_pylink_param_bound()
            p = view.LookupParameter(PYLINK_PARAM_NAME)
        if p is None:
            return False
        text = '|'.join(
            '{}-{}'.format(k, v) for k, v in kwargs.items() if v is not None
        )
        with revit.Transaction('pyLink - tag view'):
            p.Set(text)
        return True
    except Exception as ex:
        logger.warning('pyLink view tag failed: {}'.format(ex))
        return False

def _get_pylink_param():
    """
    Get or create the pyLink shared parameter on ProjectInfo — this
    is the punch list: an index of what pyLink manages, not the
    per-view records themselves (see get_view_pylink_data /
    set_view_pylink_data for those). Returns the Parameter object or
    None.
    """
    try:
        proj_info = doc.ProjectInformation
        if proj_info is None:
            return None
        p = proj_info.LookupParameter(PYLINK_PARAM_NAME)
        if p is not None:
            return p
        _ensure_pylink_param_bound()
        return proj_info.LookupParameter(PYLINK_PARAM_NAME)
    except Exception as ex:
        logger.warning('pyLink param get failed: {}'.format(ex))
        return None

def _doc_base_dir():
    """Folder the current Revit document lives in, or None if the
    document has never been saved (nothing to be relative to yet)."""
    try:
        pn = doc.PathName
        if pn:
            return os.path.dirname(pn)
    except Exception:
        pass
    return None

def _to_relative(p, base_dir):
    """Best-effort convert an absolute path to relative-to-base_dir.
    Falls back to the original absolute path if that's not possible
    (e.g. different drive, or no base_dir yet)."""
    if not p or not base_dir:
        return p
    try:
        return os.path.relpath(p, base_dir)
    except Exception:
        return p

def _to_absolute(p, base_dir):
    """Resolve a possibly-relative stored path back to absolute using
    the given base_dir. Absolute paths pass through unchanged. A
    relative path with no base_dir available (doc never saved, or
    was saved somewhere pyLink can't see) is returned as-is — the
    caller's file-exists check will simply fail, same as a genuinely
    missing file."""
    if not p or os.path.isabs(p):
        return p
    if not base_dir:
        return p
    try:
        return os.path.normpath(os.path.join(base_dir, p))
    except Exception:
        return p

def save_pylink_state(file_data):
    """
    Serialise pyLink UI state to the shared parameter on ProjectInfo.

    file_data: {path: {rows: [Row, ...], ...}}

    Format:
        #card 01
        C:\\path\\to\\file.xlsx
        VN-name|S-sheet|R-range|VT-viewtype
        ...
        #card 02
        ...

    Cards with path_mode == 'relative' are stored relative to the
    current .rvt's folder, so the link survives the project folder
    (rvt + source files together) being moved or copied elsewhere.
    """
    base_dir = _doc_base_dir()
    lines = []
    for i, (path, fd) in enumerate(file_data.items(), 1):
        pm = fd.get('path_mode', 'absolute')
        rel_ok = pm == 'relative' and base_dir
        lines.append('#card {:02d}'.format(i))
        lines.append(_to_relative(path, base_dir) if rel_ok else path)
        for row in fd.get('rows', []):
            mt = ''
            try:
                if row._applied_mtime:
                    mt = str(int(row._applied_mtime))
            except Exception:
                pass
            h = ''
            try:
                if row._applied_hash:
                    h = row._applied_hash
            except Exception:
                pass
            at = ''
            try:
                if row._applied_at:
                    at = str(int(row._applied_at))
            except Exception:
                pass
            cn = getattr(row, 'ColNo', 1)
            vs = getattr(row, 'ViewScale', 1)
            pr = getattr(row, 'Priority', 'Medium')
            gr = getattr(row, 'Group', '')
            avn = getattr(row, '_applied_view_name', '') or ''
            vid = getattr(row, '_applied_view_id', None)
            vid = '' if vid is None else str(vid)
            en = '1' if getattr(row, 'Enabled', True) else '0'
            lines.append('VN-{}|S-{}|R-{}|VT-{}|MT-{}|H-{}|CN-{}|PR-{}|GR-{}|AT-{}|AVN-{}|VS-{}|VID-{}|EN-{}'.format(
                row.ViewName, row.Sheet, row.NamedRange,
                row.ViewType, mt, h, cn, pr, gr, at, avn, vs, vid, en))
        # Card-level word settings
        ss = fd.get('sheet_size', '')
        cc = fd.get('col_count', '')
        vn = fd.get('view_name', '')
        vt = fd.get('view_type', '')
        rp = fd.get('real_path', path)
        rp_stored = _to_relative(rp, base_dir) if rel_ok else rp
        ul = '1' if fd.get('unlinked') else '0'
        lm = fd.get('layout_mode', 'manual')
        avn = fd.get('_applied_view_name', '') or ''
        card_at = ''
        try:
            if fd.get('_applied_at'):
                card_at = str(int(fd['_applied_at']))
        except Exception:
            pass
        cl = '1' if fd.get('collapsed') else '0'
        cv = '1' if fd.get('combine_views') else '0'
        if ss:
            lines.append('CARD_SS-{}|CC-{}|VN-{}|VT-{}|RP-{}|UL-{}|PM-{}|LM-{}|AVN-{}|CAT-{}|CL-{}|CV-{}'.format(
                ss, cc, vn, vt, rp_stored, ul, pm, lm, avn, card_at, cl, cv))
        elif (rp != path or fd.get('unlinked') or pm != 'absolute'
                or lm != 'manual' or avn or fd.get('collapsed')
                or fd.get('combine_views')):
            # Excel duplicate/unlinked cards have no sheet_size line,
            # still need real_path (and now unlink/path-mode/layout
            # mode/applied view name) recorded so they round-trip on
            # reload.
            lines.append('CARD_RP-{}|UL-{}|PM-{}|LM-{}|AVN-{}|CAT-{}|CL-{}|CV-{}'.format(
                rp_stored, ul, pm, lm, avn, card_at, cl, cv))
    text = '\n'.join(lines)
    try:
        p = _get_pylink_param()
        if p is not None:
            with revit.Transaction('pyLink - save state'):
                p.Set(text)
            # Nothing linked means nothing stored: an empty list writes
            # an empty parameter rather than leaving the last list that
            # happened to be saved sitting in the project.
            logger.debug('pyLink state {} ({} chars)'.format(
                'cleared' if not text else 'saved', len(text)))
    except Exception as ex:
        logger.warning('pyLink save failed: {}'.format(ex))

def clear_pylink_state():
    """Blank pyLink's stored link list on Project Information.

    ONLY the record is cleared. Nothing made from it is touched: the
    schedules, legends and drafting views pyLink generated stay exactly
    as they are, their own per-view records stay with them (see
    get_view_pylink_data), and whatever is open in the pyLink window is
    left alone. This is the punch list, not the work.

    Worth knowing: the window writes its state back whenever it changes,
    so clearing while cards are still open only empties the parameter
    until the next edit. It is meant for a project whose links are gone
    or finished with.

    Returns True if there was something stored to clear.
    """
    try:
        p = _get_pylink_param()
        if p is None:
            return False
        had = bool(p.AsString())
        if had:
            with revit.Transaction('pyLink - clear stored links'):
                p.Set('')
            logger.debug('pyLink state cleared')
        return had
    except Exception as ex:
        logger.warning('pyLink clear failed: {}'.format(ex))
        return False


def pylink_state_summary():
    """What the project has stored, without rebuilding any of it.

    Returns {'cards': n, 'rows': n, 'missing': [path, ...]} - missing
    being the stored files that are no longer where they were, which is
    what fills the log with "restore: file not found" on open.
    """
    summary = {'cards': 0, 'rows': 0, 'missing': []}
    try:
        cards = load_pylink_state()
    except Exception as ex:
        logger.warning('pyLink summary failed: {}'.format(ex))
        return summary
    for card in cards:
        summary['cards'] += 1
        summary['rows'] += len(card.get('rows') or ())
        path = card.get('real_path') or card.get('path') or ''
        if not path or not os.path.exists(path):
            summary['missing'].append(path)
    return summary


def load_pylink_state():
    """
    Read pyLink state from the shared parameter.

    Returns list of dicts:
        [{'path': str, 'rows': [{'view_name', 'sheet', 'named_range', 'view_type'}]}]

    Any card whose path (or real_path) was stored relative gets
    resolved back to absolute here, against the current .rvt's
    folder, before the caller ever sees it.
    """
    try:
        p = _get_pylink_param()
        if p is None:
            return []
        text = p.AsString()
        if not text:
            return []
        base_dir = _doc_base_dir()
        cards = []
        current = None
        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue
            if line.startswith('#card'):
                current = {'path': '', 'rows': []}
                cards.append(current)
            elif current is not None and not current['path']:
                current['path'] = line
            elif current is not None and '|' in line:
                parts = {}
                for seg in line.split('|'):
                    if '-' in seg:
                        k, v = seg.split('-', 1)
                        parts[k] = v
                if 'CARD_SS' in parts:
                    # Card-level word settings line
                    current['sheet_size'] = parts.get('CARD_SS', 'A3 Landscape')
                    try:
                        current['col_count'] = int(parts.get('CC', 2))
                    except Exception:
                        current['col_count'] = 2
                    current['view_name'] = parts.get('VN', '')
                    current['view_type'] = parts.get('VT', '')
                    current['real_path'] = parts.get('RP', '')
                    current['unlinked']  = parts.get('UL') == '1'
                    current['path_mode'] = parts.get('PM', 'absolute')
                    current['layout_mode'] = parts.get('LM', 'manual')
                    # Backfill: state saved before this tracker existed has no
                    # AVN token, so assume the view hasn't been renamed since
                    # the last apply. Safest guess for legacy state, and
                    # harmless if wrong - the rename lookup just falls back to
                    # search-then-create.
                    current['applied_view_name'] = (
                        parts.get('AVN') or current['view_name'])
                    cat = parts.get('CAT', '')
                    current['applied_at'] = float(cat) if cat else None
                    current['collapsed'] = parts.get('CL') == '1'
                    current['combine_views'] = parts.get('CV') == '1'
                    continue
                if 'CARD_RP' in parts:
                    # Excel duplicate card, no sheet_size line, just the
                    # real underlying file path (and unlink/path-mode)
                    current['real_path'] = parts.get('CARD_RP', '')
                    current['unlinked']  = parts.get('UL') == '1'
                    current['path_mode'] = parts.get('PM', 'absolute')
                    current['layout_mode'] = parts.get('LM', 'manual')
                    current['applied_view_name'] = parts.get('AVN', '')
                    cat = parts.get('CAT', '')
                    current['applied_at'] = float(cat) if cat else None
                    current['collapsed'] = parts.get('CL') == '1'
                    current['combine_views'] = parts.get('CV') == '1'
                    continue
                mt = parts.get('MT', '')
                at = parts.get('AT', '')
                current['rows'].append({
                    'view_name':    parts.get('VN', ''),
                    'sheet':        parts.get('S',  ''),
                    'named_range':  parts.get('R',  ''),
                    'view_type':    parts.get('VT', 'Schedule View'),
                    'applied_mtime': float(mt) if mt else None,
                    'applied_hash':  parts.get('H') or None,
                    'applied_at':   float(at) if at else None,
                    'col_no':       int(parts.get('CN', 1)),
                    'priority':     parts.get('PR', 'Medium'),
                    'group':        parts.get('GR', ''),
                    'applied_view_name': parts.get('AVN') or None,
                    'view_scale':   int(parts.get('VS', 1) or 1),
                    'applied_view_id': (
                        int(parts['VID']) if parts.get('VID') else None
                    ),
                    # Older saved state has no EN token at all - default
                    # True so existing projects don't suddenly need every
                    # row rechecked; going forward this preserves an
                    # actual deliberate uncheck across a reopen.
                    'enabled': parts.get('EN', '1') == '1',
                })
        # Resolve any relative path/real_path back to absolute now
        # that we know the current doc's location.
        for card in cards:
            if card.get('path_mode') == 'relative':
                card['path'] = _to_absolute(card['path'], base_dir)
                if card.get('real_path'):
                    card['real_path'] = _to_absolute(card['real_path'], base_dir)
        return cards
    except Exception as ex:
        logger.warning('pyLink load failed: {}'.format(ex))
        return []

def hb(h):
    h = h.lstrip('#')
    return SolidColorBrush(Color.FromRgb(
        int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)))


def format_applied_at(ts):
    """Format a row/view's _applied_at (wall-clock time it was last
    successfully synced into Revit) for display. None (never applied
    yet) shows as a plain dash rather than a blank, so it reads as
    'nothing here yet' rather than looking broken."""
    if not ts:
        return u'\u2014'
    try:
        return _time.strftime('%d/%m/%Y %H:%M', _time.localtime(ts))
    except Exception:
        return u'\u2014'


class Row(object):
    _counter = [0]

    def __init__(self, file_path, source_type,
                 sheets=None, sheet_range_map=None):
        Row._counter[0] += 1
        self._id              = Row._counter[0]
        self.FilePath         = file_path
        self.SourceType       = source_type
        self.Enabled          = False
        self.ViewName         = ''
        self.Sheet            = ''
        self.NamedRange       = ''
        self.ViewType         = VIEW_TYPES[0]
        self.Status           = 'pending'
        self.LastModified     = self._mtime(file_path)
        self._sheets          = sheets or []
        self._sheet_range_map = sheet_range_map or {}
        self._dot             = None   # WPF Ellipse, set by _make_row_ui
        self._refresh_btn     = None   # per-row refresh button, shown on sync
        self._vn_textbox      = None   # TextBox ref for ViewName — update when auto-filling
        self._error_label     = None   # inline error pill Border
        self._error_text      = None   # TextBlock inside error pill
        self._applied_mtime   = None   # source file's mtime when last applied (staleness check)
        self._applied_hash    = None   # MD5 of range content at last apply
        self._applied_at      = None   # wall-clock time this row was last synced into Revit
        self._applied_view_name = None # the Revit view name this row actually created/owns —
                                        # lets _view_name_taken() recognise "this is my own
                                        # view" instead of flagging it as a conflict with itself
        self._applied_view_id   = None # ElementId (int) of the same view — the authoritative
                                        # ownership proof, survives the view being renamed
        self._modified_label  = None   # TextBlock showing _applied_at, live-updated on sync
        # Word-specific
        self.ColNo            = 1        # column assignment (1-based)
        self.ViewScale        = 1        # Excel Legend/Drafting view scale
        self.Priority          = 'Medium' # High/Medium/Low - layout algo reorder freedom
        self.Group             = ''       # section group name, packs with same-group rows
        self._col_textbox     = None   # TextBox for col number
        self._enabled_cb      = None   # CheckBox ref, set by row-UI builders
        self._priority_combo  = None   # ComboBox ref, set by _make_word_row_ui
        self._group_combo     = None   # ComboBox ref, set by _make_word_row_ui
        self._drag_origin     = None   # mouse Y when drag started
        self._drag_panel_ref  = None   # card_panel ref for reorder
        if self._sheets:
            self.Sheet = self._sheets[0]

    def _mtime(self, path):
        try:
            dt = DateTime.FromFileTime(
                int(os.path.getmtime(path) * 10000000) + 116444736000000000)
            return dt.ToString('dd/MM/yyyy HH:mm')
        except Exception:
            return ''

    def ranges_for(self, sheet=None):
        s = sheet if sheet is not None else self.Sheet
        return self._sheet_range_map.get(s, [])

    @property
    def SourceLabel(self):
        ext = os.path.splitext(self.FilePath or '')[1].lower()
        if ext == '.ods':
            return 'ODS'
        if ext == '.odt':
            return 'ODT'
        return {'xl': 'XL', 'word': 'W'}.get(self.SourceType, '?')

    @property
    def SourceColour(self):
        ext = os.path.splitext(self.FilePath or '')[1].lower()
        if ext == '.ods':
            return hb(SRC_COLOURS.get('ods', '#555'))
        if ext == '.odt':
            return hb(SRC_COLOURS.get('odt', '#555'))
        return hb(SRC_COLOURS.get(self.SourceType, '#555'))
