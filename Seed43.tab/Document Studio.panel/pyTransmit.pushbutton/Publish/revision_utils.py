# -*- coding: utf-8 -*-
# revision_utils.py

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Revision


def build_rev_sheet_lookup(doc):
    """
    Return a dict mapping each revision ElementId to the first sheet
    that carries it. Used to resolve revision marks in Per Sheet mode,
    where the mark is stored on the sheet rather than the revision.
    """
    lookup = {}
    try:
        sheets = list(FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_Sheets)
            .WhereElementIsNotElementType().ToElements())
        for sheet in sheets:
            for rid in sheet.GetAllRevisionIds():
                if rid not in lookup:
                    lookup[rid] = sheet
    except Exception:
        pass
    return lookup


def safe_rev_num(rev, sheet=None):
    """
    Return the revision mark string for a revision, working in both
    Per Project and Per Sheet numbering modes.

    Tries in order:
    1. Per Sheet mark from the supplied sheet (if given)
    2. The revision-level RevisionNumber property
    3. Empty string, so the caller can apply its own fallback
    """
    if sheet is not None:
        try:
            n = sheet.GetRevisionNumberForRevision(rev.Id)
            if n:
                return n.strip()
        except Exception:
            pass
    try:
        return (rev.RevisionNumber or '').strip()
    except Exception:
        return ''


def rev_letter(seq):
    """
    Convert a Revit revision sequence number to a letter (1=A, 2=B, etc.).
    Used as a last-resort fallback when no mark can be read from Revit.
    """
    n = seq - 1
    if n < 0:
        return '?'
    result = ''
    while True:
        result = chr(65 + (n % 26)) + result
        n = n // 26 - 1
        if n < 0:
            break
    return result


def get_rev_mark(rev, sheet=None, seq_fallback=True):
    """
    Return the best available revision mark for display.
    Reads ptransmit_rev parameter first (written by write_rev_param.py),
    then falls back to RevisionNumber, then to a calculated letter.
    """
    # Read from ptransmit_rev parameter on the sheet first
    if sheet is not None:
        marks = get_sheet_rev_marks(sheet)
        if marks is not None:
            # Parameter exists, trust what is stored, including blank marks
            if not marks:
                # Parameter is blank, revision has no mark (None numbering)
                return ''
            try:
                all_revs = sorted(
                    [r for r in FilteredElementCollector(sheet.Document)
                        .OfClass(Revision).ToElements() if r.Issued],
                    key=lambda r: r.SequenceNumber
                )
                rev_ids = [r.Id for r in all_revs]
                if rev.Id in rev_ids:
                    idx = rev_ids.index(rev.Id)
                    if idx < len(marks):
                        return marks[idx]
            except Exception:
                pass
    # Fall back to API-based mark only when no ptransmit_rev data exists
    mark = safe_rev_num(rev, sheet=sheet)
    if not mark and seq_fallback:
        mark = rev_letter(rev.SequenceNumber)
    return mark


def get_sheet_rev_marks(sheet, param_name='ptransmit_rev'):
    """
    Return the list of revision marks stored on the sheet in the
    ptransmit_rev parameter, split on pipe separator.
    Returns None if the parameter does not exist on the sheet.
    Returns [] if the parameter exists but is blank (no issued revisions).
    Returns ['Draft1', '', 'A'] etc when marks are present.
    """
    try:
        p = sheet.LookupParameter(param_name)
        if p is not None:
            val = (p.AsString() or '').strip()
            return val.split('|') if val else []
    except Exception:
        pass
    return None
