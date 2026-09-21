# -*- coding: utf-8 -*-
# studio_live_data.py
#
# Pulls real Revit + pyTransmit-settings data into a dict shaped like
# Layout/LayoutSettings.py's DUMMY (proj_org, proj_client, distribution,
# revisions, reasons, methods, docs), so studio_blocks.render_block() renders
# live data through the same path as the empty state. Read-only: never writes
# to the user's Settings/*.json under .user.

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


def _parse_recipient(issued_to_str, recipient_label, recipient_index=0):
    """(attention_to, copies) for one recipient, out of a Revision's IssuedTo.

    Ported from Publish/script_create_excel.py so the builder preview and
    the real published document agree on how the tags are read.

    Both values come out of the SAME tag - the distribution block is written
    as "DL: 1A.[BECA]1 2O.[] 3C.[Structa]1", where the bracket is the contact
    and the digits after it are the copy count. This used to return only the
    count and throw the contact away, which is why Attention To printed blank
    on transmittals whose copies printed fine: the name was in the model the
    whole time, just never read out.
    """
    block = ''
    for part in (issued_to_str or '').split(' | '):
        part = part.strip()
        if part.startswith('DL:') or part.startswith('CL:'):
            block = part[3:].strip()
            break
    if not block and ' | ' in (issued_to_str or ''):
        block = issued_to_str.split(' | ', 1)[1].strip()
    # Distribution mode, new format: position-numbered, e.g. "3C.[Structa]1".
    m = re.search(r'{}[A-Za-z]\.\[([^\]]*)\](\d*)'.format(recipient_index + 1), block)
    if m:
        return m.group(1).strip(), m.group(2)
    # Distribution mode, old format: keyed by the role's initial only.
    first = recipient_label[0].upper() if recipient_label else ''
    if first:
        m2 = re.search(r'(?:^| )' + first + r'\.\[([^\]]*)\](\d*)', block)
        if m2:
            return m2.group(1).strip(), m2.group(2)
    # Client mode: the bracket holds "Company — Contact" rather than a
    # bare contact, so the contact is the half after the em dash - the same
    # split pyTransmit.py makes when it rebuilds recipients from IssuedTo.
    if recipient_label:
        m3 = re.search(r'\[(' + re.escape(recipient_label[:6]) + r'[^\]]*)\](\d+)',
                       issued_to_str or '')
        if m3:
            inner = m3.group(1)
            attn = inner.split(u'—', 1)[1].strip() if u'—' in inner else ''
            return attn, m3.group(2)
    return '', ''


def empty_data():
    """Shape with no content - used when there's no active Revit document."""
    return {
        'proj_org': '', 'proj_client': '', 'proj_number': '', 'proj_name': '',
        'doc_type': '', 'print_size': '',
        'distribution': [], 'revisions': [], 'reasons': [], 'methods': [], 'docs': [],
        'sheet_params': [],
    }


def _groupable_sheet_params(sheets):
    """Sheet parameters that can be grouped on, by the same rule pyTransmit's
    own Sheet Parameters picker uses (pyTransmit.py get_sheet_parameters):
    text-backed parameters plus ElementId ones, which is how Revit 2026's
    Dropdown List parameters (e.g. Sheet Collection) are stored. Integer and
    Double are left out - a Scale of 100 is not a meaningful group heading.

    Sampled from one sheet, again like pyTransmit, because sheets in a project
    share a parameter set and reading every sheet's full parameter list is
    needlessly slow on a 2000-sheet model.
    """
    from pyrevit import DB
    sample = next(iter(sheets), None)
    if sample is None:
        return []

    names = set()
    for bip in (DB.BuiltInParameter.SHEET_NUMBER, DB.BuiltInParameter.SHEET_NAME):
        try:
            p = sample.get_Parameter(bip)
            if p and p.StorageType == DB.StorageType.String:
                names.add(p.Definition.Name)
        except Exception:
            pass

    usable = (DB.StorageType.String, DB.StorageType.ElementId)
    try:
        ordered = sample.GetOrderedParameters()
    except Exception:
        ordered = []
    for p in ordered:
        try:
            if not p.Definition or p.StorageType not in usable:
                continue
            if p.StorageType == DB.StorageType.ElementId and not p.AsValueString():
                continue
            names.add(p.Definition.Name)
        except Exception:
            continue
    return sorted(names)


def get_live_data(settings_dir, max_revs=12, group_params=None, recipients=None):
    """settings_dir = the user's Settings folder (pytransmit_paths.SETTINGS_DIR),
    passed in rather than resolved here, so this can
    read the same recipients/distribution/reason/method JSON pyTransmit
    itself uses, without touching pyTransmit.py's own payload-building code.

    group_params = the sheet parameters the caller is about to GROUP by.
    They are read off every sheet whether or not they pass the groupable
    test below - see the note in _groupable_sheet_params(). Grouping by a
    parameter that was never read produces a blank group key for every
    sheet, one nameless group, and therefore no group header rows at all -
    silently, and identically in the Studio preview and in every export.

    recipients = the caller's PYTRANSMIT_PAYLOAD['recipients'] - the
    [{label, attn, copies}] list the pyTransmit window built from whichever
    recipient mode is active. Passed in because it is the only place the
    contact names the user typed exist before they are written to a revision,
    and because in client mode the rows are companies, not the fixed roles in
    distribution.json. Publish scripts have it; the Studio canvas does not,
    and falls back to reading the last issued revision."""
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

    # -- Distribution ---------------------------------------------------------
    # Three sources, most authoritative first. Attention To is filled from the
    # first one that has a name; blanks fall through to the next.
    #
    #   1. the payload      what the pyTransmit window is publishing RIGHT NOW,
    #                       including contacts typed but not yet written to any
    #                       revision, and the right ROWS in client mode.
    #   2. the revision     parsed out of IssuedTo further down, once the
    #                       issued revisions have been read.
    #   3. the JSON files   recipients.json joined to distribution.json.
    #
    # (3) only ever fires by coincidence and is kept as a last resort: the two
    # files hold different vocabularies - distribution.json has fixed ROLES
    # ("Architect/Designer"), recipients.json has client COMPANIES - so the
    # names match only if a company happens to be spelt like a role. Relying
    # on it alone is what left Attention To blank on every transmittal.
    distribution = []
    for row in (recipients or []):
        label = (row.get('label') or '').strip()
        if label:
            distribution.append({'to': label,
                                 'attn': (row.get('attn') or '').strip(),
                                 'copies': '', 'copies_by_rev': []})

    if not distribution:
        dist_rows = _load_json(os.path.join(settings_dir, 'distribution.json'), [])
        recip_rows = _load_json(os.path.join(settings_dir, 'recipients.json'), [])
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
        # Attention To is not per-revision: it is one name per recipient, so
        # the latest issued revision that names one wins, and anything the
        # payload already supplied wins over all of them.
        for di, dist in enumerate(distribution):
            per_rev = []
            parsed_attn = ''
            for rev_el in issued_revs:
                try:
                    r_attn, r_copies = _parse_recipient(
                        rev_el.IssuedTo or '', dist.get('to', ''), di)
                except Exception:
                    r_attn, r_copies = '', ''
                per_rev.append(r_copies)
                if r_attn:
                    parsed_attn = r_attn
            dist['attn'] = dist.get('attn') or parsed_attn
            dist['copies_by_rev'] = per_rev
            dist['copies'] = per_rev[-1] if per_rev else ''

        issued_ids = set(r.Id for r in issued_revs)
        sheets = sorted(
            [s for s in revit.query.get_elements_by_class(DB.ViewSheet, doc=doc)
             if any(rid in set(s.GetAllRevisionIds()) for rid in issued_ids)],
            key=lambda s: _natural_sort_key(s.SheetNumber)
        )
        # Values for every groupable parameter, read once per sheet here so
        # studio_blocks.sheet_row_plan() can group without touching Revit -
        # the renderer runs on every repaint, the model does not.
        param_names = _groupable_sheet_params(sheets)
        # Whatever the caller means to group by is read too, even if the
        # sample sheet made it look unusable. That test samples ONE sheet, so
        # a Dropdown List parameter left empty on it - or any parameter the
        # sample happens not to carry - is dropped from the list, and the
        # grouping the user set in pyTransmit then quietly does nothing.
        for pn in (group_params or []):
            if pn and pn not in param_names:
                param_names.append(pn)
                param_names.sort()
        data['sheet_params'] = param_names

        docs = []
        for s in sheets:
            sheet_revs = set(s.GetAllRevisionIds())
            docs.append({
                'sheet': s.SheetNumber,
                'desc': s.Name,
                'revs': [r['rev'] if r_obj.Id in sheet_revs else ''
                         for r, r_obj in zip(revisions, issued_revs)],
                'params': dict((pn, _get_param(s, pn)) for pn in param_names),
            })
        data['docs'] = docs
    except Exception:
        data.setdefault('revisions', [])
        data.setdefault('docs', [])
        data.setdefault('sheet_params', [])

    return data
