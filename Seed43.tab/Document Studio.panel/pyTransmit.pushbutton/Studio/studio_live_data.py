# -*- coding: utf-8 -*-
# studio_live_data.py
#
# Pulls real Revit + pyTransmit-settings data into a dict shaped exactly like
# Layout/LayoutSettings.py's DUMMY (proj_org, proj_client, distribution,
# revisions, reasons, methods, docs) so studio_blocks.render_block() can
# render live data with the same code path it uses for an empty/no-data
# state. Read-only: never writes to any Settings/*.json.

import os
import re
import json
import imp


def _get_param(el, name):
    try:
        p = el.LookupParameter(name)
        if p and p.HasValue:
            return (p.AsString() or p.AsValueString() or '').strip()
    except Exception:
        pass
    return ''


def _natural_sort_key(s):
    parts = re.split(r'(\d+)', str(s))
    return [int(p) if p.isdigit() else p.lower() for p in parts]


def _load_json(path, default):
    try:
        with open(path, 'r') as f:
            return json.load(f)
    except Exception:
        return default


def _parse_copies(issued_to_str, recipient_label, recipient_index=0):
    """Copies-per-recipient encoded in a Revision's IssuedTo string.

    Ported from Publish/script_create_excel.py so the builder preview and
    the real published document agree on how the tags are read.
    """
    block = ''
    for part in (issued_to_str or '').split(' | '):
        part = part.strip()
        if part.startswith('DL:') or part.startswith('CL:'):
            block = part[3:].strip()
            break
    if not block and ' | ' in (issued_to_str or ''):
        block = issued_to_str.split(' | ', 1)[1].strip()
    m = re.search(r'{}[A-Za-z]\.\[([^\]]*)\](\d*)'.format(recipient_index + 1), block)
    if m:
        return m.group(2)
    first = recipient_label[0].upper() if recipient_label else ''
    if first:
        m2 = re.search(r'(?:^| )' + first + r'\.\[([^\]]*)\](\d*)', block)
        if m2:
            return m2.group(2)
    if recipient_label:
        m3 = re.search(r'\[' + re.escape(recipient_label[:6]) + r'[^\]]*\](\d+)', issued_to_str or '')
        if m3:
            return m3.group(1)
    return ''


def empty_data():
    """Shape with no content - used when there's no active Revit document."""
    return {
        'proj_org': '', 'proj_client': '', 'proj_number': '', 'proj_name': '',
        'doc_type': '', 'print_size': '',
        'distribution': [], 'revisions': [], 'reasons': [], 'methods': [], 'docs': [],
    }


def get_live_data(settings_dir, max_revs=12):
    """settings_dir = the pyTransmit.pushbutton/Settings folder, so this can
    read the same recipients/distribution/reason/method JSON pyTransmit
    itself uses, without touching pyTransmit.py's own payload-building code."""
    try:
        from pyrevit import revit, DB
        doc = revit.doc
    except Exception:
        return empty_data()

    if doc is None:
        return empty_data()

    try:
        proj_info = doc.ProjectInformation
        data = {
            'proj_org':    _get_param(proj_info, 'Organization Name'),
            'proj_client': _get_param(proj_info, 'Client Name'),
            'proj_number': _get_param(proj_info, 'Project Number'),
            'proj_name':   _get_param(proj_info, 'Project Name') or doc.Title or '',
            'doc_type': '', 'print_size': '',
        }
    except Exception:
        data = empty_data()

    # -- Distribution (from Settings/distribution.json + recipients.json) ------
    dist_rows = _load_json(os.path.join(settings_dir, 'distribution.json'), [])
    recip_rows = _load_json(os.path.join(settings_dir, 'recipients.json'), [])
    distribution = []
    for row in sorted(dist_rows, key=lambda r: r.get('display_order', 0)):
        label = row.get('distribution', '')
        attn = ', '.join(
            r.get('attention_to', '') for r in recip_rows
            if r.get('company') == label and r.get('attention_to')
        )
        distribution.append({'to': label, 'attn': attn, 'copies': '',
                             'copies_by_rev': []})
    data['distribution'] = distribution

    # -- Reasons / Methods (from Settings/reason.json, method.json) -----------
    data['reasons'] = [{'code': r.get('code', ''), 'label': r.get('description', '')}
                       for r in _load_json(os.path.join(settings_dir, 'reason.json'), [])]
    data['methods'] = [{'code': r.get('code', ''), 'label': r.get('description', '')}
                       for r in _load_json(os.path.join(settings_dir, 'method.json'), [])]

    # -- Revisions + sheets (real Revit elements) -------------------------------
    try:
        all_revs = list(revit.query.get_elements_by_class(DB.Revision, doc=doc))
        issued_revs = sorted([r for r in all_revs if r.Issued], key=lambda r: r.SequenceNumber)
        issued_revs = issued_revs[-max_revs:]

        ru_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                                'Publish', 'revision_utils.py')
        get_rev_mark = None
        rev_sheet_lookup = {}
        if os.path.isfile(ru_path):
            ru = imp.load_source('studio_revision_utils', ru_path)
            get_rev_mark = ru.get_rev_mark
            rev_sheet_lookup = ru.build_rev_sheet_lookup(doc)

        revisions = []
        for rev in issued_revs:
            ito = rev.IssuedTo or ''

            def _f(key, raw=ito):
                m = re.search(r'\[' + key + r':([^\]]*)\]', raw)
                return m.group(1).strip() if m else ''

            letter = ''
            if get_rev_mark:
                try:
                    letter = get_rev_mark(rev, sheet=rev_sheet_lookup.get(rev.Id))
                except Exception:
                    letter = ''
            revisions.append({
                'rev':       letter or str(rev.SequenceNumber),
                'date':      _get_param(rev, 'Revision Date') or str(rev.RevisionDate),
                'initials':  (rev.IssuedBy or '').strip() or _f('I'),
                'reason':    _f('R'),
                'method':    _f('M'),
                # Same [F:]/[P:] tags script_create_excel.py reads out of
                # the revision's IssuedTo string.
                'doc_format': _f('F'),
                'paper_size': _f('P'),
            })
        data['revisions'] = revisions
        # Copies are per (recipient, revision) - one entry per revision so a
        # Number of Copies block in revision column N can read index N.
        for di, dist in enumerate(distribution):
            per_rev = []
            for rev_el in issued_revs:
                try:
                    per_rev.append(_parse_copies(rev_el.IssuedTo or '', dist.get('to', ''), di))
                except Exception:
                    per_rev.append('')
            dist['copies_by_rev'] = per_rev
            dist['copies'] = per_rev[-1] if per_rev else ''

        issued_ids = set(r.Id for r in issued_revs)
        sheets = sorted(
            [s for s in revit.query.get_elements_by_class(DB.ViewSheet, doc=doc)
             if any(rid in set(s.GetAllRevisionIds()) for rid in issued_ids)],
            key=lambda s: _natural_sort_key(s.SheetNumber)
        )
        docs = []
        for s in sheets:
            sheet_revs = set(s.GetAllRevisionIds())
            docs.append({
                'sheet': s.SheetNumber,
                'desc': s.Name,
                'revs': [r['rev'] if r_obj.Id in sheet_revs else ''
                         for r, r_obj in zip(revisions, issued_revs)],
            })
        data['docs'] = docs
    except Exception:
        data.setdefault('revisions', [])
        data.setdefault('docs', [])

    return data
