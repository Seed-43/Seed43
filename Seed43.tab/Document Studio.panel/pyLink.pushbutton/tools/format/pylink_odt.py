# -*- coding: utf-8 -*-
from pyrevit import script

logger = script.get_logger()

"""
pylink_odt.py -- ODT (OpenDocument Text)-specific reading for pyLink:
extracts headings and paragraphs (with bold/italic/underline
formatting) directly from an .odt zip's content.xml - no COM, no
third-party libraries. Returns the same section-list shape
pylink_docx.py returns, so pylink_word.py's read_word_sections
dispatcher and the rest of pyLink don't need to know which
word-processor format a document actually came from.
"""


def read_odt_sections(file_path):
    """ODT equivalent of read_docx_sections. ODT doesn't inline bold/
    italic/underline the way docx's w:rPr does — text runs (text:span)
    reference a named style instead, resolved here against
    office:automatic-styles' style:text-properties. Formatting is
    aggregated per paragraph (own style OR any nested span's style),
    matching docx's own per-paragraph bold_any/italic_any/underline_any
    aggregation rather than tracking per-run.
    Not implemented yet for ODT: numbered/bulleted list markers (list
    items still extract as plain paragraph text, no bullet prefix) —
    docx's numbering.xml bullet-character lookup has no ODT equivalent
    here yet."""
    import zipfile
    try:
        import clr as _clr
        _clr.AddReference('System.Xml')
    except Exception:
        pass
    from System.Xml import XmlDocument, XmlNodeType

    try:
        with zipfile.ZipFile(file_path, 'r') as zf:
            content_xml = zf.read('content.xml').decode('utf-8', errors='replace')
    except Exception as ex:
        logger.error('read_odt_sections: cannot open {}: {}'.format(file_path, ex))
        return []

    try:
        xdoc = XmlDocument()
        xdoc.LoadXml(content_xml)
    except Exception as ex:
        logger.error('read_odt_sections: XML parse failed: {}'.format(ex))
        return []

    # style-name -> (bold, italic, underline), from every style:style
    # with text-properties (covers both paragraph and character styles)
    style_props = {}
    style_nodes = xdoc.GetElementsByTagName('style:style')
    for i in range(style_nodes.Count):
        st = style_nodes.Item(i)
        name = st.GetAttribute('style:name')
        if not name:
            continue
        tp_nodes = st.GetElementsByTagName('style:text-properties')
        if tp_nodes.Count == 0:
            continue
        tp = tp_nodes.Item(0)
        weight = tp.GetAttribute('fo:font-weight')
        fstyle = tp.GetAttribute('fo:font-style')
        uline  = tp.GetAttribute('style:text-underline-style')
        style_props[name] = (
            weight == 'bold', fstyle == 'italic',
            bool(uline) and uline != 'none')

    def _text_of(node):
        """Full text of a text:h/text:p node, preserving line breaks
        and tabs as \\n / \\t — ODT stores these as empty elements,
        not literal characters, unlike docx's w:tab (which the docx
        reader here does the equivalent for)."""
        parts = []
        def _walk(n):
            for child in list(n.ChildNodes):
                if child.NodeType == XmlNodeType.Text:
                    parts.append(child.Value)
                elif child.LocalName == 'line-break':
                    parts.append(u'\n')
                elif child.LocalName == 'tab':
                    parts.append(u'\t')
                elif child.LocalName == 's':
                    parts.append(u' ')
                else:
                    _walk(child)
        _walk(node)
        return u''.join(parts)

    def _para_props(node):
        own_style = node.GetAttribute('text:style-name')
        bold, italic, underline = style_props.get(own_style, (False, False, False))
        span_nodes = node.GetElementsByTagName('text:span')
        for i in range(span_nodes.Count):
            sp_style = span_nodes.Item(i).GetAttribute('text:style-name')
            b, it, ul = style_props.get(sp_style, (False, False, False))
            bold = bold or b
            italic = italic or it
            underline = underline or ul
        return bold, italic, underline

    def _is_heading(node, text, bold):
        if node.LocalName == 'h':
            return True
        if not text or text.startswith('('):
            return False
        is_upper = text == text.upper() and any(c.isalpha() for c in text)
        return bold and is_upper and len(text) <= 80

    # Collect text:h / text:p blocks under office:text in document
    # order via a single depth-first walk (GetElementsByTagName on two
    # separate tag names would need re-sorting into document order -
    # this walk avoids that entirely).
    blocks = []
    def _collect(n):
        for child in list(n.ChildNodes):
            if child.NodeType != XmlNodeType.Element:
                continue
            if child.LocalName in ('h', 'p'):
                blocks.append(child)
            else:
                _collect(child)
    body_nodes = xdoc.GetElementsByTagName('office:text')
    if body_nodes.Count:
        _collect(body_nodes.Item(0))

    sections = []
    current  = None
    for node in blocks:
        text = _text_of(node).strip()
        bold, italic, underline = _para_props(node)
        if _is_heading(node, text, bold):
            if current is not None:
                sections.append(current)
            current = {'heading': text, 'paragraphs': []}
        else:
            if current is None:
                if text:
                    current = {'heading': '', 'paragraphs': []}
            if current is not None and text:
                current['paragraphs'].append({
                    'text': text, 'bold': bold,
                    'italic': italic, 'underline': underline,
                    'bullet': '',
                })
    if current is not None:
        sections.append(current)

    return sections
