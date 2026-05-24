# -*- coding: utf-8 -*-
# write_rev_param.py

import os

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory,
    Transaction, Revision,
)

PARAM_NAME = 'ptransmit_rev'


def _ensure_param(doc, spf_path):
    app = doc.Application
    # Already bound?
    it = doc.ParameterBindings.ForwardIterator()
    while it.MoveNext():
        try:
            if it.Key and it.Key.Name == PARAM_NAME:
                return True
        except Exception:
            pass
    # Load shared parameter file and bind
    orig = app.SharedParametersFilename
    try:
        app.SharedParametersFilename = spf_path
        spf = app.OpenSharedParameterFile()
        if not spf:
            return False
        defn = None
        for grp in spf.Groups:
            for d in grp.Definitions:
                if d.Name == PARAM_NAME:
                    defn = d
                    break
            if defn:
                break
        if defn is None:
            return False
        cat_set = app.Create.NewCategorySet()
        sheets_cat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Sheets)
        cat_set.Insert(sheets_cat)
        binding = app.Create.NewInstanceBinding(cat_set)
        t = Transaction(doc, 'pyTransmit: Add ptransmit_rev')
        t.Start()
        try:
            from Autodesk.Revit.DB import SharedParameterElement as _SPE
            try:
                if not _SPE.Lookup(doc, defn.GUID):
                    _SPE.Create(doc, defn)
            except Exception:
                pass
            doc.ParameterBindings.Insert(defn, binding)
            t.Commit()
            return True
        except Exception:
            t.RollBack()
            return False
    except Exception:
        return False
    finally:
        try:
            app.SharedParametersFilename = orig
        except Exception:
            pass


def _calc_mark(rev, doc, issued_revs):
    """Calculate the revision mark using the Revit 2022+ numbering sequence API."""
    try:
        seq_id   = rev.RevisionNumberingSequenceId
        seq_elem = doc.GetElement(seq_id)
        if seq_elem is None:
            return ''  # None numbering type has no sequence element
        same_seq = [r for r in issued_revs if r.RevisionNumberingSequenceId == seq_id]
        pos = same_seq.index(rev)
        num_type_str = str(seq_elem.NumberType)
        if 'None' in num_type_str:
            return ''
        elif 'Numeric' in num_type_str:
            ns = seq_elem.GetNumericRevisionSettings()
            return (ns.Prefix or '') + str(ns.StartNumber + pos) + (ns.Suffix or '')
        else:
            als = seq_elem.GetAlphanumericRevisionSettings()
            seq_list = list(als.GetSequence())
            if seq_list:
                return (als.Prefix or '') + seq_list[pos % len(seq_list)] + (als.Suffix or '')
    except Exception:
        pass
    # Last resort: letter from sequence number
    n = rev.SequenceNumber - 1
    if n < 0:
        return '?'
    result = ''
    while True:
        result = chr(65 + (n % 26)) + result
        n = n // 26 - 1
        if n < 0:
            break
    return result


def _get_mark(rev, sheet, issued_revs, doc):
    # Check sequence type first, return empty for None numbering
    mark = _calc_mark(rev, doc, issued_revs)
    if mark == '':
        return ''
    # Per Sheet mode: use GetRevisionNumberOnSheet (Revit 2022+)
    try:
        n = sheet.GetRevisionNumberOnSheet(rev.Id)
        if n and n.strip():
            return n.strip()
    except Exception:
        pass
    # Per Project mode: RevisionNumber property
    try:
        n = rev.RevisionNumber
        if n and n.strip():
            return n.strip()
    except Exception:
        pass
    return mark


def write_rev_param(doc, publish_dir):
    spf_path = os.path.join(publish_dir, 'ptransmit_rev.txt')
    if not os.path.isfile(spf_path):
        return
    if not _ensure_param(doc, spf_path):
        return
    all_revs = list(FilteredElementCollector(doc).OfClass(Revision).ToElements())
    issued_revs = sorted([r for r in all_revs if r.Issued], key=lambda r: r.SequenceNumber)
    if not issued_revs:
        return
    all_sheets = list(FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_Sheets)
        .WhereElementIsNotElementType().ToElements())
    sheet_marks = {}
    for sheet in all_sheets:
        sheet_rev_ids = set(sheet.GetAllRevisionIds())
        marks = []
        for rev in issued_revs:
            if rev.Id in sheet_rev_ids:
                marks.append(_get_mark(rev, sheet, issued_revs, doc))
            else:
                marks.append('')
        sheet_marks[sheet.Id] = marks
    t = Transaction(doc, 'pyTransmit: Write ptransmit_rev')
    t.Start()
    try:
        for sheet in all_sheets:
            p = sheet.LookupParameter(PARAM_NAME)
            if p and not p.IsReadOnly:
                p.Set('|'.join(sheet_marks.get(sheet.Id, [])))
        t.Commit()
    except Exception:
        t.RollBack()
