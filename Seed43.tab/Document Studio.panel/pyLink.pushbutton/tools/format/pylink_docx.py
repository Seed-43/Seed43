# -*- coding: utf-8 -*-
from pyrevit import script

logger = script.get_logger()

"""
pylink_docx.py -- docx-specific reading for pyLink: extracts headings
and paragraphs (with bold/italic/underline/bullet formatting) directly
from a .docx zip's word/document.xml - no COM, no third-party
libraries. Returns the same section-list shape pylink_odt.py returns,
so pylink_word.py's read_word_sections dispatcher and the rest of
pyLink don't need to know which word-processor format a document
actually came from.
"""


def read_docx_sections(file_path):
    """
    Parse a .docx file and extract sections as a list of dicts:
        [{'heading': str, 'paragraphs': [{'text': str, 'bold': bool,
          'italic': bool, 'underline': bool}]}, ...]

    A section starts when a paragraph is detected as a heading:
    - Word heading styles (Heading1, Heading2, etc.)
    - Bold-only paragraphs with all-caps or short text (<= 60 chars)

    Uses zipfile + XmlDocument — no COM, no third-party libraries.
    """
    import zipfile
    try:
        import clr as _clr
        _clr.AddReference('System.Xml')
    except Exception:
        pass

    from System.Xml import XmlDocument

    def _load_xml(text):
        xd = XmlDocument()
        xd.LoadXml(text)
        return xd

    NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    def _attr(node, local):
        """Get a w: attribute value by local name."""
        try:
            return node.GetAttribute(local, NS) or node.GetAttribute(local)
        except Exception:
            return ''

    def _text_of(para_node):
        """
        Concatenate all w:t and w:tab content inside a paragraph node,
        preserving tab characters in document order.
        w:tab elements are emitted as \t — this is essential for
        alignment tables where Word uses tab stops to align columns.
        """
        parts = []
        # Walk all descendant nodes looking for w:t and w:tab in order
        # We use a simple recursive walk since XmlNodeList ordering
        # is document order for GetElementsByTagName.
        # Strategy: get all runs (w:r) in order, then within each run
        # get child w:t and w:tab nodes.
        def _walk_run(run):
            child = run.FirstChild
            while child is not None:
                local = child.LocalName
                if local == 't':
                    parts.append(child.InnerText)
                elif local == 'tab':
                    parts.append(u'\t')
                child = child.NextSibling

        # First try namespaced runs
        runs = para_node.GetElementsByTagName('r', NS)
        if not runs.Count:
            runs = para_node.GetElementsByTagName('r')
        for i in range(runs.Count):
            _walk_run(runs.Item(i))

        # Fallback: no runs — grab w:t directly (old behaviour)
        if not parts:
            for t in para_node.GetElementsByTagName('t', NS):
                parts.append(t.InnerText)
            if not parts:
                for t in para_node.GetElementsByTagName('t'):
                    parts.append(t.InnerText)
        return u''.join(parts)

    def _is_heading_style(style_id):
        sid = (style_id or '').lower()
        return (sid.startswith('heading') or
                sid in ('title', 'subtitle', 'caption'))

    def _run_props(run_node):
        """Return (bold, italic, underline) for a w:r run node."""
        bold = italic = underline = False
        rpr_list = run_node.GetElementsByTagName('rPr', NS)
        if not rpr_list.Count:
            rpr_list = run_node.GetElementsByTagName('rPr')
        if rpr_list.Count:
            rpr = rpr_list.Item(0)
            bold      = bool(rpr.GetElementsByTagName('b',  NS).Count or
                             rpr.GetElementsByTagName('b').Count)
            italic    = bool(rpr.GetElementsByTagName('i',  NS).Count or
                             rpr.GetElementsByTagName('i').Count)
            underline = bool(rpr.GetElementsByTagName('u',  NS).Count or
                             rpr.GetElementsByTagName('u').Count)
        return bold, italic, underline

    # ── Bullet character map from numbering.xml ──
    _bullet_chars = {}   # numId (str) -> bullet char string
    try:
        with zipfile.ZipFile(file_path, 'r') as _zf:
            if 'word/numbering.xml' in _zf.namelist():
                _nxml = _zf.read('word/numbering.xml').decode(
                    'utf-8', errors='replace')
                _ndoc = _load_xml(_nxml)
                # abstractNum entries carry the bullet format
                for _an in list(_ndoc.GetElementsByTagName(
                        'abstractNum', NS)) + list(
                        _ndoc.GetElementsByTagName('abstractNum')):
                    for _lvl in list(_an.GetElementsByTagName(
                            'lvl', NS)) + list(
                            _an.GetElementsByTagName('lvl')):
                        # Only ilvl 0 (first level)
                        ilvl = (_attr(_lvl, 'ilvl') or
                                _lvl.GetAttribute('w:ilvl') or '0')
                        if ilvl != '0':
                            continue
                        _nfmt_els = (list(_lvl.GetElementsByTagName(
                            'numFmt', NS)) or list(
                            _lvl.GetElementsByTagName('numFmt')))
                        _ltxt_els = (list(_lvl.GetElementsByTagName(
                            'lvlText', NS)) or list(
                            _lvl.GetElementsByTagName('lvlText')))
                        if _nfmt_els and _ltxt_els:
                            fmt = (_attr(_nfmt_els[0], 'val') or
                                   _nfmt_els[0].GetAttribute('w:val') or '')
                            txt = (_attr(_ltxt_els[0], 'val') or
                                   _ltxt_els[0].GetAttribute('w:val') or
                                   u'·')
                            if fmt == 'bullet':
                                # map abstractNumId -> char
                                _an_id = (_attr(_an, 'abstractNumId') or
                                          _an.GetAttribute('w:abstractNumId') or
                                          '0')
                                _bullet_chars[_an_id] = txt
                # num->abstractNum mapping
                _num_map = {}  # numId -> bullet char
                for _num in list(_ndoc.GetElementsByTagName(
                        'num', NS)) + list(
                        _ndoc.GetElementsByTagName('num')):
                    _nid = (_attr(_num, 'numId') or
                            _num.GetAttribute('w:numId') or '')
                    _anid_els = (list(_num.GetElementsByTagName(
                        'abstractNumId', NS)) or list(
                        _num.GetElementsByTagName('abstractNumId')))
                    if _anid_els and _nid:
                        _anid = (_attr(_anid_els[0], 'val') or
                                 _anid_els[0].GetAttribute('w:val') or '')
                        if _anid in _bullet_chars:
                            _num_map[_nid] = _bullet_chars[_anid]
                _bullet_chars.update(_num_map)
    except Exception as _bex:
        logger.debug('bullet parse: {}'.format(_bex))

    def _get_bullet_char(para_node):
        """Return bullet prefix string if paragraph is a list item, else ''."""
        ppr = None
        ppr_list = para_node.GetElementsByTagName('pPr', NS)
        if not ppr_list.Count:
            ppr_list = para_node.GetElementsByTagName('pPr')
        if ppr_list.Count:
            ppr = ppr_list.Item(0)
        if ppr is None:
            return ''
        num_pr = (list(ppr.GetElementsByTagName('numPr', NS)) or
                  list(ppr.GetElementsByTagName('numPr')))
        if not num_pr:
            return ''
        num_id_els = (list(num_pr[0].GetElementsByTagName('numId', NS)) or
                      list(num_pr[0].GetElementsByTagName('numId')))
        if not num_id_els:
            return u'§ '   # fallback § if numPr exists but no numId
        nid = (_attr(num_id_els[0], 'val') or
               num_id_els[0].GetAttribute('w:val') or '')
        char = _bullet_chars.get(nid, u'§')
        # Normalise common bullet chars to § to match doc style
        if char in (u'•', u'·', u'', '-', '*', u'–'):
            char = u'§'
        return char + u' '

    def _para_is_heading(para_node, style_id):
        if _is_heading_style(style_id):
            return True
        text = _text_of(para_node).strip()
        if not text:
            return False
        # Never treat parenthesised text as a heading
        if text.startswith('('):
            return False
        # List items are never headings
        if _get_bullet_char(para_node):
            return False
        # Heuristic: bold AND all-uppercase
        runs = list(para_node.GetElementsByTagName('r', NS))
        if not runs:
            runs = list(para_node.GetElementsByTagName('r'))
        if not runs:
            return False
        all_bold = all(_run_props(r)[0] for r in runs if _text_of(r).strip())
        is_upper = text == text.upper() and any(c.isalpha() for c in text)
        return all_bold and is_upper and len(text) <= 80

    sections = []
    current  = None

    try:
        with zipfile.ZipFile(file_path, 'r') as zf:
            doc_xml = zf.read('word/document.xml').decode('utf-8', errors='replace')
    except Exception as ex:
        logger.error('read_docx_sections: cannot open {}: {}'.format(file_path, ex))
        return []

    try:
        xdoc = _load_xml(doc_xml)
    except Exception as ex:
        logger.error('read_docx_sections: XML parse failed: {}'.format(ex))
        return []

    paras = xdoc.GetElementsByTagName('p', NS)
    if not paras.Count:
        paras = xdoc.GetElementsByTagName('p')

    for i in range(paras.Count):
        p = paras.Item(i)

        # Get paragraph style id
        style_id = ''
        ppr_list = p.GetElementsByTagName('pPr', NS)
        if not ppr_list.Count:
            ppr_list = p.GetElementsByTagName('pPr')
        if ppr_list.Count:
            ppr = ppr_list.Item(0)
            pstyle = ppr.GetElementsByTagName('pStyle', NS)
            if not pstyle.Count:
                pstyle = ppr.GetElementsByTagName('pStyle')
            if pstyle.Count:
                style_id = (_attr(pstyle.Item(0), 'val') or
                            pstyle.Item(0).GetAttribute('w:val') or '')

        text = _text_of(p).strip()

        if _para_is_heading(p, style_id):
            if current is not None:
                sections.append(current)
            current = {'heading': text, 'paragraphs': []}
        else:
            if current is None:
                # Text before any heading — create anonymous section
                if text:
                    current = {'heading': '', 'paragraphs': []}
            if current is not None:
                # Collect run-level formatting for the paragraph
                runs = list(p.GetElementsByTagName('r', NS))
                if not runs:
                    runs = list(p.GetElementsByTagName('r'))
                bullet_prefix = _get_bullet_char(p)
                if runs:
                    bold_any = italic_any = underline_any = False
                    for r in runs:
                        b, it, ul = _run_props(r)
                        if b:  bold_any      = True
                        if it: italic_any    = True
                        if ul: underline_any = True
                    current['paragraphs'].append({
                        'text':      text,
                        'bold':      bold_any,
                        'italic':    italic_any,
                        'underline': underline_any,
                        'bullet':    bullet_prefix,
                    })
                elif text:
                    current['paragraphs'].append({
                        'text': text, 'bold': False,
                        'italic': False, 'underline': False,
                        'bullet': bullet_prefix,
                    })

    if current is not None:
        sections.append(current)

    return sections
