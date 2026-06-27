# -*- coding: utf-8 -*-
"""
create_notes.py  —  pyTable Word notes export.

One TextNote per column. All sections in a column are assembled into a
single plain-text string, then FormattedText overrides apply bold
headings and native Revit bullet lists.

Text rules applied to every paragraph:
  - Leading 2+ spaces           -> single tab
  - Internal 2+ spaces before ' - '  -> tab (alignment tables)
  - Bullet body paragraphs wrapped by Revit use \v for continuation
    lines so they stay within the same bullet paragraph.

Sheet widths (mm):
    Landscape: A4=297 A3=420 A2=594 A1=841 A0=1189
    Portrait:  A4=210 A3=297 A2=420 A1=594 A0=841
Column width = sheet_width / col_count, 5 mm gap between columns.
"""

import re as _re

_p = globals().get('PYTABLE_PAYLOAD', {})

from pyrevit import revit, script, DB
from Autodesk.Revit.DB import (
    FilteredElementCollector, XYZ,
    TextNote, TextNoteType, TextNoteOptions,
    HorizontalTextAlignment,
    FormattedText, TextRange, ListType,
    ViewDrafting, ViewFamilyType, ViewFamily,
    CurveElement, ImageInstance,
)

logger = script.get_logger()
doc    = revit.doc
MM     = 1.0 / 304.8

SHEET_WIDTHS_MM = {
    'A4 Landscape':  297.0, 'A4 Portrait':  210.0,
    'A3 Landscape':  420.0, 'A3 Portrait':  297.0,
    'A2 Landscape':  594.0, 'A2 Portrait':  420.0,
    'A1 Landscape':  841.0, 'A1 Portrait':  594.0,
    'A0 Landscape': 1189.0, 'A0 Portrait':  841.0,
    'A4': 297.0, 'A3': 420.0, 'A2': 594.0,
    'A1': 841.0, 'A0': 1189.0,
}
GAP_MM       = 5.0
DEFAULT_SIZE = 2.3   # mm — body text size in Revit


# ── Text normalisation ────────────────────────────────────────────────

# Arial character width proportions (fraction of cap height)
_ARIAL_W = {
    ' ':0.38,'f':0.38,'i':0.38,'j':0.38,'l':0.38,'r':0.42,'t':0.45,
    'I':0.42,'(':0.45,')':0.45,'[':0.45,']':0.45,'!':0.38,'.':0.38,
    ',':0.38,':':0.38,';':0.38,'s':0.58,'z':0.55,'x':0.60,
    'a':0.65,'b':0.65,'c':0.60,'d':0.65,'e':0.65,'g':0.65,'h':0.65,
    'k':0.62,'n':0.65,'o':0.68,'p':0.65,'q':0.65,'u':0.65,'v':0.60,
    'y':0.60,'A':0.72,'B':0.70,'C':0.70,'D':0.75,'E':0.65,'F':0.60,
    'G':0.75,'H':0.75,'J':0.50,'K':0.70,'L':0.62,'N':0.75,'O':0.78,
    'P':0.65,'Q':0.78,'R':0.72,'S':0.65,'T':0.65,'U':0.75,'V':0.72,
    'X':0.70,'Y':0.68,'Z':0.65,'M':0.85,'W':0.90,'m':0.95,'w':0.85,
    '/':0.45,'-':0.45,'_':0.65,
}
_ARIAL_W_DEF = 0.65
_TAB_MM      = 3.0
_SPACE_MM    = 0.63   # space width at 2.3mm Arial 85%


def _label_width_mm(label, cap_mm=DEFAULT_SIZE, scale=0.85):
    """Width of a label string in mm using Arial proportional metrics."""
    return sum(_ARIAL_W.get(c, _ARIAL_W_DEF) * cap_mm * scale
               for c in label)


def _tabs_to_clear(width_mm):
    """Minimum tab count to move past a label of given width."""
    import math
    return max(1, int(math.ceil(width_mm / _TAB_MM)))


def _spaces_to_tabs(n):
    """Convert a run of n spaces to the nearest tab count."""
    return max(1, int(round(n * _SPACE_MM / _TAB_MM)))


def _normalise(text):
    """
    Convert space runs to tabs using the space-count formula:
        tabs = round(n_spaces * 0.63mm / 3.0mm)
    where 0.63mm is one space width at 2.3mm Arial 85% scale
    and 3.0mm is the Revit tab stop size.

    Results:
        6  spaces = round(1.26) = 1 tab
        10 spaces = round(2.10) = 2 tabs
        12 spaces = round(2.52) = 3 tabs

    Leading spaces -> single tab (indent level only).
    Internal alignment spaces -> tab count via formula.
    Real w:tab characters from XML pass through unchanged.
    """
    if not text:
        return text

    # Leading spaces -> single tab (always just one indent level)
    stripped  = text.lstrip(u' ')
    n_leading = len(text) - len(stripped)
    if n_leading >= 2:
        text = u'\t' + stripped

    # Internal space runs of 2+ -> tabs via formula
    def _replace(m):
        return u'\t' * _spaces_to_tabs(len(m.group(0)))

    text = _re.sub(r' {2,}', _replace, text)
    return text


def _normalise_align_run(texts, cap_mm=DEFAULT_SIZE, scale=0.85):
    """
    Post-process a list of alignment-table paragraph texts so all labels
    land at the same dash column, regardless of individual label widths.

    Algorithm:
      1. Detect alignment pattern in each line:
             [optional leading tab] + label + [spaces or tabs] + rest
      2. Measure each label width with Arial metrics.
      3. Find max tabs needed to clear the widest label.
      4. Replace all space/tab padding with that uniform tab count.

    Example — S/S (3.4mm) and FSBW (5.6mm) in the same run:
      FSBW needs ceil(5.6/3) = 2 tabs to clear.
      S/S  needs ceil(3.4/3) = 2 tabs to clear.
      Both get 2 tabs -> both land at 6mm. Aligned.
    """
    PAT = _re.compile(r'^(\t*)(\S+)([ \t]+)(.*)')

    labels = []
    for text in texts:
        m = PAT.match(text)
        if m:
            labels.append(m.group(2))

    if not labels:
        return texts

    max_tabs = max(_tabs_to_clear(_label_width_mm(lbl, cap_mm, scale))
                   for lbl in labels)

    result = []
    for text in texts:
        m = PAT.match(text)
        if m:
            text = m.group(1) + m.group(2) + u'\t' * max_tabs + m.group(4)
        result.append(text)
    return result


# ── TextNoteType ──────────────────────────────────────────────────────

def _get_or_create_text_type(name, size_mm):
    size_ft = size_mm * MM
    all_tt  = list(FilteredElementCollector(doc)
                   .OfClass(TextNoteType).ToElements())
    existing = None
    for tt in all_tt:
        try:
            n = tt.get_Parameter(
                DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            if n == name:
                existing = tt
                break
        except Exception:
            pass
    with revit.Transaction('pyTable Notes - type: {}'.format(name)):
        new_tt = existing if existing else all_tt[0].Duplicate(name)
        for bip, val in [
            (DB.BuiltInParameter.TEXT_SIZE,        size_ft),
            (DB.BuiltInParameter.TEXT_FONT,        'Arial'),
            (DB.BuiltInParameter.TEXT_STYLE_BOLD,  0),
            (DB.BuiltInParameter.TEXT_STYLE_ITALIC, 0),
            (DB.BuiltInParameter.TEXT_BACKGROUND,  1),
            (DB.BuiltInParameter.TEXT_WIDTH_SCALE, 0.85),
            # Tab size = 3 mm
            (DB.BuiltInParameter.TEXT_TAB_SIZE,    3.0 * MM),
        ]:
            try:
                p = new_tt.get_Parameter(bip)
                if p and not p.IsReadOnly:
                    p.Set(val)
            except Exception:
                pass
        try:
            ap = new_tt.get_Parameter(
                DB.BuiltInParameter.LEADER_ARROWHEAD)
            if ap and not ap.IsReadOnly:
                ap.Set(DB.ElementId.InvalidElementId)
        except Exception:
            pass
    return new_tt


# ── View helpers ──────────────────────────────────────────────────────

def _get_or_create_drafting_view(name):
    for v in FilteredElementCollector(doc).OfClass(ViewDrafting):
        try:
            if v.IsValidObject and v.Name == name:
                return v
        except Exception:
            pass
    vft = None
    for t in FilteredElementCollector(doc).OfClass(ViewFamilyType):
        if t.ViewFamily == ViewFamily.Drafting:
            vft = t
            break
    if not vft:
        raise RuntimeError('No Drafting ViewFamilyType found.')
    with revit.Transaction('pyTable Notes - create view'):
        v = ViewDrafting.Create(doc, vft.Id)
        v.Name = name
        try:
            v.Scale = 1
        except Exception:
            pass
    return v


def _get_or_create_legend_view(name):
    for v in FilteredElementCollector(doc).OfClass(DB.View):
        try:
            if (v.IsValidObject
                    and v.ViewType == DB.ViewType.Legend
                    and v.Name == name):
                return v
        except Exception:
            pass
    src = None
    for v in FilteredElementCollector(doc).OfClass(DB.View):
        try:
            if v.IsValidObject and v.ViewType == DB.ViewType.Legend:
                src = v
                break
        except Exception:
            pass
    if not src:
        raise RuntimeError('No existing Legend view to duplicate from.')
    with revit.Transaction('pyTable Notes - create legend view'):
        new_id = src.Duplicate(DB.ViewDuplicateOption.Duplicate)
        new_v  = doc.GetElement(new_id)
        new_v.Name = name
        try:
            new_v.Scale = 1
        except Exception:
            pass
    return new_v


def _clear_view(view):
    with revit.Transaction('pyTable Notes - clear view'):
        for cls in (CurveElement, TextNote, ImageInstance):
            for el in list(FilteredElementCollector(doc, view.Id)
                           .OfClass(cls).ToElements()):
                try:
                    doc.Delete(el.Id)
                except Exception:
                    pass


# ── Column text builder ───────────────────────────────────────────────

def _build_column_text(sections):
    """
    Assemble one column's full text and segment descriptors.

    Returns:
        plain_text (str)   — joined with \\r (Revit paragraph breaks)
        segments   (list)  — [{start, length, bold, is_bullet}]

    \\r = new Revit paragraph (new bullet item in a list)
    \\v = new line without new paragraph (continuation within bullet)

    Structure:
        HEADING\\r
        bullet body\\r
        bullet body\\r
        \\r            <- single blank separator between sections
        HEADING\\r
        ...
    """
    parts    = []
    segments = []
    pos      = 0

    for sec_idx, sec in enumerate(sections):
        heading = sec.get('heading', '').strip()
        paras   = sec.get('paragraphs', [])

        # ── Heading ──────────────────────────────────────────────────
        if pos > 0:
            # Single blank line before each heading (except very first)
            parts.append(u'\r')
            pos += 1

        if heading:
            seg_start = pos
            parts.append(heading)
            pos += len(heading)
            segments.append({
                'start':     seg_start,
                'length':    len(heading),
                'bold':      True,
                'is_bullet': False,
            })
            parts.append(u'\r')
            pos += 1

        # ── Body paragraphs ──────────────────────────────────────────
        bullet_run_start  = None
        bullet_run_end    = None

        def _flush_bullet_run():
            if bullet_run_start is not None:
                segments.append({
                    'start':     bullet_run_start,
                    'length':    bullet_run_end - bullet_run_start,
                    'bold':      False,
                    'is_bullet': True,
                })

        prev_was_bullet   = False   # track continuation after bullet
        prev_was_indented = False   # track continuation of indented run

        for para in paras:
            # rstrip only — preserve leading tabs (alignment info)
            raw_text = para.get('text', '').rstrip()
            bullet   = para.get('bullet', '')
            bold     = para.get('bold', False)
            text     = _normalise(raw_text)

            # For indented non-bullet lines: ensure exactly one leading tab.
            # Covers: abbrev table lines (raw starts with \t),
            #         continuation after bullet (prev_was_bullet),
            #         continuation of indented run (prev_was_indented).
            if not bullet and (raw_text.startswith(u'\t')
                               or prev_was_bullet
                               or prev_was_indented):
                if not text.startswith(u'\t'):
                    text = u'\t' + text

            if not text:
                _flush_bullet_run()
                bullet_run_start  = bullet_run_end = None
                prev_was_bullet   = False
                prev_was_indented = False   # blank line breaks indent run
                continue

            # Detect indent type for non-bullet paragraphs:
            # a) has leading \t in raw text  -> abbrev table line
            # b) immediately follows a bullet or another indented line
            #    -> continuation or abbrev run
            # The indented run continues until a blank line or new heading.
            is_indented = (not bullet and
                           (raw_text.startswith(u'\t')
                            or prev_was_bullet
                            or prev_was_indented))

            seg_start = pos
            parts.append(text)
            pos += len(text)

            if bullet:
                if bullet_run_start is None:
                    bullet_run_start = seg_start
                bullet_run_end = pos
                parts.append(u'\r')
                pos += 1
                prev_was_bullet   = True
                prev_was_indented = False
            else:
                _flush_bullet_run()
                bullet_run_start = bullet_run_end = None
                segments.append({
                    'start':       seg_start,
                    'length':      len(text),
                    'bold':        bold,
                    'is_bullet':   False,
                    'is_indented': is_indented,
                })
                parts.append(u'\r')
                pos += 1
                prev_was_bullet   = False
                prev_was_indented = is_indented

        _flush_bullet_run()
        bullet_run_start = bullet_run_end = None

    plain_text = u''.join(parts).rstrip(u'\r')
    return plain_text, segments


# ── Main ──────────────────────────────────────────────────────────────

def run():
    view_name  = _p.get('view_name', 'pyTable Notes')
    view_type  = _p.get('view_type', 'Drafting View')
    sections   = _p.get('sections', [])
    sheet_size = _p.get('sheet_size', 'A3 Landscape')
    col_count  = max(1, int(_p.get('col_count', 2)))
    size_mm    = float(_p.get('size_mm', DEFAULT_SIZE))

    sheet_w_mm = SHEET_WIDTHS_MM.get(sheet_size, 420.0)
    MARGIN_MM  = 10.0   # left and right margin inside the view
    usable_mm  = sheet_w_mm - 2.0 * MARGIN_MM
    total_gap  = GAP_MM * (col_count - 1)
    col_w_mm   = (usable_mm - total_gap) / col_count
    col_w_ft   = col_w_mm * MM
    gap_ft     = GAP_MM * MM
    margin_ft  = MARGIN_MM * MM

    if 'Legend' in view_type:
        view = _get_or_create_legend_view(view_name)
    else:
        view = _get_or_create_drafting_view(view_name)
    _clear_view(view)

    type_name = 'pyTable Notes {:.1f} Arial'.format(size_mm)
    tt = _get_or_create_text_type(type_name, size_mm)

    by_col = {}
    for sec in sections:
        col_no = max(1, min(col_count, int(sec.get('col', 1))))
        by_col.setdefault(col_no, []).append(sec)

    with revit.Transaction('pyTable Notes - place'):
        for col_no in range(1, col_count + 1):
            col_sections = by_col.get(col_no, [])
            if not col_sections:
                continue

            x_off = margin_ft + (col_no - 1) * (col_w_ft + gap_ft)
            plain_text, segments = _build_column_text(col_sections)

            if not plain_text.strip():
                continue

            opts = TextNoteOptions(tt.Id)
            opts.HorizontalAlignment = HorizontalTextAlignment.Left
            try:
                tn = TextNote.Create(
                    doc, view.Id,
                    XYZ(float(x_off), 0.0, 0.0),
                    float(col_w_ft),
                    str(plain_text),
                    opts
                )
            except Exception as ex:
                logger.error('Notes col {} create: {}'.format(col_no, ex))
                continue

            # Apply FormattedText: bold headings + native bullet lists
            try:
                fmt = tn.GetFormattedText()
                for seg in segments:
                    r = TextRange()
                    r.Start  = seg['start']
                    r.Length = seg['length']
                    if seg['is_bullet']:
                        try:
                            fmt.SetListType(r, ListType.Bullet)
                        except Exception as bex:
                            logger.debug('SetListType bullet: {}'.format(bex))
                    elif seg.get('is_indented'):
                        try:
                            fmt.SetListType(r, ListType.None_)
                        except Exception:
                            try:
                                fmt.SetListType(r, getattr(ListType, 'None'))
                            except Exception as bex:
                                logger.debug('SetListType none: {}'.format(bex))
                    if seg['bold']:
                        try:
                            fmt.SetBoldStatus(r, True)
                        except Exception as bex:
                            logger.debug('SetBoldStatus: {}'.format(bex))
                tn.SetFormattedText(fmt)
            except Exception as fex:
                logger.debug('FormattedText col {}: {}'.format(col_no, fex))


run()
