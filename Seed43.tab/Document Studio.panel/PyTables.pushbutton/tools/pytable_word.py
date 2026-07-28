# -*- coding: utf-8 -*-
from pyrevit import revit, DB
from pyrevit import forms
from pyrevit import script

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
    GridLength, GridUnitType, TextAlignment
)
from System import DateTime
from System.Windows.Controls import (
    StackPanel, Border, CheckBox, TextBlock, TextBox,
    ComboBox, Button, Orientation, ScrollViewer,
    Grid, ColumnDefinition
)
from System.Windows.Controls.Primitives import Popup, ToggleButton
from System.Windows.Shapes import Ellipse

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
pytable_word.py -- everything specific to the Word side of pyTable:
parsing .docx headings/paragraphs directly from the zip, building
native Legend/Drafting views of TextNotes from that data, the Strict-
layout packing algorithm, section-group settings (userdata/
section_groups.json), and WordCardMixin providing the Word-specific
parts of the row/card UI (mixed into PyTableWindow alongside
ExcelCardMixin in PyTable.py).
"""

import sys as _sys
_sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pytable_shared import (
    hb, Row, WORD_VIEW_TYPES, SHEET_SIZES, SRC_COLOURS, STATUS_COLOURS,
    _run_export_script, _confirm, _alert, format_applied_at,
)


SECTION_GROUPS_PATH = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    'userdata', 'section_groups.json')

DEFAULT_SECTION_GROUPS = {'groups': []}

# -- Strict-layout height-estimate constants --
# Duplicated from Export/create_notes.py rather than imported -- that
# module reads PYTABLE_PAYLOAD from globals() at import time and is
# only ever loaded via exec() with the payload injected, so a normal
# import here would break. Keep these in sync by hand if the export
# geometry ever changes.
SHEET_WIDTHS_MM = {
    'A4 Landscape':  297.0, 'A4 Portrait':  210.0,
    'A3 Landscape':  420.0, 'A3 Portrait':  297.0,
    'A2 Landscape':  594.0, 'A2 Portrait':  420.0,
    'A1 Landscape':  841.0, 'A1 Portrait':  594.0,
    'A0 Landscape': 1189.0, 'A0 Portrait':  841.0,
    'A4': 297.0, 'A3': 420.0, 'A2': 594.0,
    'A1': 841.0, 'A0': 1189.0,
}
GAP_MM             = 5.0
NOTE_TEXT_SIZE_MM  = 2.3
NOTE_MARGIN_MM     = 10.0
NOTE_LINE_HEIGHT_MM  = NOTE_TEXT_SIZE_MM * 1.6
NOTE_BULLET_GAP_MM   = 1.0
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


def match_section_group(text, groups_data=None):
    """First group whose keyword appears (case-insensitive substring)
    in the section text, or '' if nothing matches."""
    data = groups_data if groups_data is not None else load_section_groups()
    up = (text or '').upper()
    for g in data.get('groups', []):
        for kw in g.get('keywords', []):
            if kw and kw.upper() in up:
                return g.get('name', '')
    return ''

WORD_TEXT_SETTINGS_PATH = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    'userdata', 'word_text_settings.json')

DEFAULT_WORD_TEXT_SETTINGS = {
    'mode': 'auto', 'size_mm': 2.0, 'text_type_name': 'pyTable Notes 2.0 Arial',
}


def load_word_text_settings():
    """'auto' (default): pyTable reuses whichever existing project
    TextNoteType is closest to 2mm or 3mm, rather than always making
    its own. 'manual': always use the given size_mm, creating/reusing
    a pyTable-named type at that exact size."""
    try:
        if os.path.exists(WORD_TEXT_SETTINGS_PATH):
            with open(WORD_TEXT_SETTINGS_PATH, 'r') as f:
                data = _json.load(f)
            if isinstance(data, dict) and 'mode' in data:
                return data
    except Exception as ex:
        logger.warning('word_text_settings.json load failed: {}'.format(ex))
    return dict(DEFAULT_WORD_TEXT_SETTINGS)


def save_word_text_settings(data):
    try:
        folder = os.path.dirname(WORD_TEXT_SETTINGS_PATH)
        if not os.path.exists(folder):
            os.makedirs(folder)
        with open(WORD_TEXT_SETTINGS_PATH, 'w') as f:
            _json.dump(data, f, indent=2)
    except Exception as ex:
        logger.warning('word_text_settings.json save failed: {}'.format(ex))

def save_section_groups(data):
    try:
        folder = os.path.dirname(SECTION_GROUPS_PATH)
        if not os.path.exists(folder):
            os.makedirs(folder)
        with open(SECTION_GROUPS_PATH, 'w') as f:
            _json.dump(data, f, indent=2)
    except Exception as ex:
        logger.warning('section_groups.json save failed: {}'.format(ex))

def load_section_groups():
    """Read the user-editable section-grouping settings. Missing file
    or bad JSON both fall back to an empty group list rather than
    crashing pyTable — grouping is a convenience feature, never a
    hard dependency."""
    try:
        if os.path.exists(SECTION_GROUPS_PATH):
            with open(SECTION_GROUPS_PATH, 'r') as f:
                data = _json.load(f)
            if isinstance(data, dict) and isinstance(data.get('groups'), list):
                return data
    except Exception as ex:
        logger.warning('section_groups.json load failed: {}'.format(ex))
    return dict(DEFAULT_SECTION_GROUPS)

def read_word_sections(file_path):
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
    clr_ref = False
    try:
        import clr as _clr
        _clr.AddReference('System.Xml')
        clr_ref = True
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
                    parts.append(u'	')
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
        if char in (u'•', u'·', u'', '-', '*', u'–'):
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
        logger.error('read_word_sections: cannot open {}: {}'.format(file_path, ex))
        return []

    try:
        xdoc = _load_xml(doc_xml)
    except Exception as ex:
        logger.error('read_word_sections: XML parse failed: {}'.format(ex))
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

def get_word_headings(file_path):
    """
    Return display labels for the Section combo in the UI.

    When a heading appears more than once (e.g. EXTERIOR STEELWORK),
    append the first parenthesised subtitle from its body paragraphs
    so each entry is unique and meaningful:
        EXTERIOR STEELWORK (Zinc Metal Spray Only)
        EXTERIOR STEELWORK (Inorganic zinc and Top Coats)
    The label stored in row.NamedRange is this display string so we can
    look up the section at Apply time.
    """
    try:
        sections = [s for s in read_word_sections(file_path)
                    if s.get('heading')]
        # Count how many times each raw heading occurs
        from collections import Counter as _Counter
        counts = _Counter(s['heading'] for s in sections)
        labels = []
        for s in sections:
            heading = s['heading']
            if counts[heading] > 1:
                # Find first parenthesised paragraph to disambiguate
                subtitle = ''
                for p in s.get('paragraphs', []):
                    t = p.get('text', '').strip()
                    if t.startswith('(') and t.endswith(')'):
                        subtitle = ' ' + t
                        break
                labels.append(heading + subtitle)
            else:
                labels.append(heading)
        return labels
    except Exception as ex:
        logger.error('get_word_headings: {}'.format(ex))
        return []


# ── Notes row apply ──

def apply_notes_row(rows, view_name, view_type, sheet_size,
                    col_count, file_path, size_mm=None,
                    old_view_name=None, text_mode=None,
                    text_type_name=None):
    """
    Apply a set of Word notes rows to a single Drafting/Legend view.

    rows is a list of dicts:
        [{'heading': str, 'paragraphs': [...], 'col': int}, ...]

    old_view_name, if given and different from view_name, tells the
    export script to rename the previously-applied view in place
    rather than searching for a view under the new name (which would
    never find one and silently create a duplicate, orphaning the old
    view with its stale content).

    text_mode/size_mm/text_type_name default to the persisted
    word_text_settings.json when not explicitly passed — 'auto'
    reuses whichever existing project TextNoteType is closest to
    2mm/3mm; 'manual' reuses the exact TextNoteType named
    text_type_name (picked from the project's own list), falling
    back to size_mm only if that type can no longer be found.

    Returns {'view_name', 'status', 'message'}.
    """
    logger.debug('pyTable Notes: {}'.format(view_name))

    if text_mode is None or size_mm is None or text_type_name is None:
        settings = load_word_text_settings()
        if text_mode is None:
            text_mode = settings.get('mode', 'auto')
        if size_mm is None:
            size_mm = settings.get('size_mm', 2.3)
        if text_type_name is None:
            text_type_name = settings.get('text_type_name')

    result = {'view_name': view_name, 'status': 'error', 'message': ''}

    if not rows:
        result['message'] = 'No sections to place.'
        return result

    payload = {
        'view_name':      view_name,
        'old_view_name':  old_view_name,
        'view_type':      view_type,
        'sections':       rows,
        'sheet_size':     sheet_size,
        'col_count':      col_count,
        'size_mm':        size_mm,
        'text_mode':      text_mode,
        'text_type_name': text_type_name,
    }

    try:
        _run_export_script('create_notes.py', payload)
        result['status']  = 'success'
        result['message'] = 'Created'
    except Exception as ex:
        import traceback
        result['message'] = str(ex)
        logger.error(traceback.format_exc())

    return result

def _hash_word_section(file_path, heading):
    """
    Compute a hash of a single Word section (heading + body paragraphs).
    Used for per-row sync detection on Word cards, equivalent to
    _hash_range for Excel rows.

    Same fix as _hash_range: plain repr(dict) is not safe here, dict key
    iteration order can differ between separate Revit sessions (hash
    randomisation), which was causing every Word row to show as
    "changed" even when the source file was untouched. Sort every
    paragraph dict's keys before repr(); the paragraph list's own order
    is left alone since that's semantically meaningful (it's the order
    they appear in the document), unlike dict key order which isn't.
    """
    import hashlib
    try:
        sections = read_word_sections(file_path)
        for sec in sections:
            if sec.get('heading', '') == heading:
                paras_sorted = [
                    sorted(p.items()) for p in sec.get('paragraphs', [])
                ]
                other_sorted = sorted(
                    (k, v) for k, v in sec.items() if k != 'paragraphs'
                )
                content = repr((other_sorted, paras_sorted))
                return hashlib.md5(
                    content.encode('utf-8', errors='replace')).hexdigest()
    except Exception:
        pass
    return None


class WordCardMixin(object):
    """Word-specific card/row UI methods, mixed into
    PyTableWindow. Everything here assumes fd.get('source_type') == 'word'."""

    # ── Grouped-card rendering ──
    # All fd entries sharing the same real_path (the original card
    # plus every one added via '+ Add Row') render inside ONE outer
    # card with ONE shared header (collapse / badge / path+date /
    # + Add Row / reload / close). Each fd gets its own nested "view
    # block": its own View Name/Sheet size/Columns/View Type, its own
    # + Add Section, its own sections-collapse toggle, its own
    # Strict/Manual + Batch, its own section rows. self._card_groups
    # (set up in __init__) tracks {real_path: {...}} for this.

    def _make_word_card(self, path):
        fd = self._file_data[path]
        real_path = fd.get('real_path', path)

        group = self._card_groups.get(real_path)
        if group is None:
            group = self._build_word_group_header(real_path, path)
            self._card_groups[real_path] = group

        view_block = self._build_word_view_block(path)
        group['views_panel'].Children.Add(view_block)
        group['view_keys'].append(path)
        fd['card_border'] = view_block
        fd['group_real_path'] = real_path

        self._update_word_group_reload_indicator(real_path)
        self._update_card_link_badge(path)
        self._update_tri_select_state(path)

    def _build_word_group_header(self, real_path, template_path):
        """One-time outer card + header for a document. template_path
        is the fd that triggered this group's creation — the header's
        '+ Add Row' button stays tagged to it so _card_add_view (which
        reads default settings from a specific fd) keeps working
        unchanged; it always appends into this same group regardless,
        since _make_word_card looks the group up by real_path."""
        outer = Border()
        try:
            outer.Style = self.FindResource('CardStyle')
        except Exception as e:
            logger.warning('Failed to apply CardStyle: {}'.format(e))

        inner = StackPanel()
        inner.Orientation = Orientation.Vertical

        header_row = Grid()
        header_row.Margin = Thickness(0, 0, 0, 8)
        _hcol_l = ColumnDefinition(); _hcol_l.Width = GridLength(1, GridUnitType.Star)
        _hcol_r = ColumnDefinition(); _hcol_r.Width = GridLength.Auto
        header_row.ColumnDefinitions.Add(_hcol_l)
        header_row.ColumnDefinitions.Add(_hcol_r)

        header_left = StackPanel()
        header_left.Orientation = Orientation.Horizontal
        Grid.SetColumn(header_left, 0)
        header_row.Children.Add(header_left)

        header_right = StackPanel()
        header_right.Orientation = Orientation.Horizontal
        header_right.HorizontalAlignment = HorizontalAlignment.Right
        Grid.SetColumn(header_right, 1)
        header_row.Children.Add(header_right)

        collapse_btn = ToggleButton()
        collapse_btn.IsChecked = True
        try:
            collapse_btn.Style = self.FindResource('PrimarySecondaryToggleButtonStyle')
        except Exception as e:
            logger.warning('Failed to apply PrimarySecondaryToggleButtonStyle: {}'.format(e))
        collapse_btn.FocusVisualStyle = None
        collapse_btn.HorizontalContentAlignment = HorizontalAlignment.Center
        collapse_btn.VerticalContentAlignment   = VerticalAlignment.Center
        collapse_btn.VerticalAlignment = VerticalAlignment.Center
        collapse_btn.Margin = Thickness(0, 0, 8, 0)
        collapse_btn.Tag     = real_path
        collapse_btn.ToolTip = 'Collapse'
        collapse_btn.Click  += self._toggle_word_group_collapse
        self._set_collapse_icon(collapse_btn, False)
        header_left.Children.Add(collapse_btn)

        src_badge = Border()
        src_badge.Width        = 24
        src_badge.Height       = 24
        src_badge.CornerRadius = CornerRadius(4)
        src_badge.Background   = hb(SRC_COLOURS.get('word', '#2B579A'))
        src_badge.Margin       = Thickness(0, 0, 8, 0)
        src_badge.VerticalAlignment = VerticalAlignment.Center
        src_lbl = TextBlock()
        src_lbl.Text                = 'W'
        src_lbl.FontSize            = 9
        src_lbl.FontWeight          = FontWeights.Bold
        src_lbl.Foreground          = hb('#FFFFFF')
        src_lbl.HorizontalAlignment = HorizontalAlignment.Center
        src_lbl.VerticalAlignment   = VerticalAlignment.Center
        src_badge.Child = src_lbl
        header_left.Children.Add(src_badge)

        heading = TextBlock()
        heading.Text         = real_path
        heading.TextTrimming = TextTrimming.CharacterEllipsis
        heading.ToolTip      = real_path
        heading.VerticalAlignment = VerticalAlignment.Center
        heading.Foreground   = hb('#6B7280')
        heading.FontWeight   = FontWeights.SemiBold
        heading.FontSize     = 13
        header_left.Children.Add(heading)

        lm_text = TextBlock()
        try:
            dt = DateTime.FromFileTime(
                int(os.path.getmtime(real_path) * 10000000) + 116444736000000000)
            lm_text.Text = dt.ToString('dd/MM/yyyy HH:mm')
        except Exception:
            lm_text.Text = ''
        lm_text.FontSize          = 10
        lm_text.Foreground        = hb('#F4FAFF')
        lm_text.Opacity           = 0.55
        lm_text.VerticalAlignment = VerticalAlignment.Center
        lm_text.Margin            = Thickness(10, 0, 0, 0)
        header_left.Children.Add(lm_text)

        header_add_btn = self._green_btn(u'+ Add Row')
        header_add_btn.Tag     = template_path
        header_add_btn.Click  += self._card_add_view
        header_add_btn.ToolTip = (
            'Add a new, independent view for this document — its own '
            'View Name, sections, and layout, without needing to '
            'browse for the file again.')
        header_add_btn.VerticalAlignment = VerticalAlignment.Center
        header_add_btn.Margin = Thickness(6, 0, 0, 0)
        header_right.Children.Add(header_add_btn)

        # Batch — one per card, not one per view. Its actions (Open
        # File/Folder, Absolute/Relative Path, Unlink, Remove view,
        # Delete selected) are all document-wide, so this uses
        # template_path as the reference view, same as + Add Row above.
        batch_btn = Button()
        batch_btn.Content    = u'Batch \u25be'
        try:
            batch_btn.Style = self.FindResource('SecondaryButtonStyle')
        except Exception as e:
            logger.warning('Failed to apply SecondaryButtonStyle: {}'.format(e))
        batch_btn.FocusVisualStyle = None
        batch_btn.Height      = 24
        batch_btn.Padding     = Thickness(12, 0, 12, 0)
        batch_btn.FontSize    = 11
        batch_btn.VerticalAlignment = VerticalAlignment.Center
        batch_btn.Margin      = Thickness(6, 0, 0, 0)
        batch_btn.Tag         = template_path
        batch_btn.Click      += self._card_batch_menu
        header_right.Children.Add(batch_btn)

        reload_btn = Button()
        if _mi is not None:
            try:
                reload_btn.Content = _mi('reload', size=14, color='#FFFFFF')
            except Exception:
                reload_btn.Content = u'\u21bb'
        else:
            reload_btn.Content = u'\u21bb'
        try:
            reload_btn.Style = self.FindResource('RoundPrimaryButtonStyle')
        except Exception as e:
            logger.warning('Failed to apply RoundPrimaryButtonStyle: {}'.format(e))
        reload_btn.FocusVisualStyle = None
        reload_btn.Width       = 28
        reload_btn.Height      = 28
        reload_btn.HorizontalContentAlignment = HorizontalAlignment.Center
        reload_btn.VerticalContentAlignment   = VerticalAlignment.Center
        reload_btn.VerticalAlignment          = VerticalAlignment.Center
        reload_btn.Margin  = Thickness(6, 0, 0, 0)
        reload_btn.Tag     = real_path
        reload_btn.ToolTip = 'Reload views that need updating'
        reload_btn.Click  += self._group_reload_click
        header_right.Children.Add(reload_btn)

        del_card_btn = Button()
        del_card_btn.Content         = u'\u2715'
        del_card_btn.FontSize        = 11
        try:
            del_card_btn.Style = self.FindResource('DeleteButtonStyle')
        except Exception as e:
            logger.warning('Failed to apply DeleteButtonStyle: {}'.format(e))
        del_card_btn.FocusVisualStyle = None
        del_card_btn.Width            = 28
        del_card_btn.Height           = 28
        del_card_btn.HorizontalContentAlignment = HorizontalAlignment.Center
        del_card_btn.VerticalContentAlignment   = VerticalAlignment.Center
        del_card_btn.VerticalAlignment   = VerticalAlignment.Center
        del_card_btn.Margin          = Thickness(6, 0, 0, 0)
        del_card_btn.ToolTip         = 'Close (remove this card and all its views)'
        del_card_btn.Tag             = real_path
        del_card_btn.Click          += self._group_close_click
        header_right.Children.Add(del_card_btn)

        inner.Children.Add(header_row)

        # Column headers for the compact view rows below — shown once
        # per card, not repeated on every view, mirroring how Excel's
        # column headers work. Same two-column Grid structure as
        # value_row below it, so the right-aligned Layout label lines
        # up with its dropdown instead of just floating on the left.
        views_col_hdr = Grid()
        views_col_hdr.Margin = Thickness(0, 0, 0, 4)
        _vhcol_l = ColumnDefinition(); _vhcol_l.Width = GridLength(1, GridUnitType.Star)
        _vhcol_r = ColumnDefinition(); _vhcol_r.Width = GridLength.Auto
        views_col_hdr.ColumnDefinitions.Add(_vhcol_l)
        views_col_hdr.ColumnDefinitions.Add(_vhcol_r)

        def _vch(text, width, pad_left=4):
            tb = TextBlock()
            tb.Text              = text
            tb.Width             = width
            tb.FontSize          = 10
            tb.Foreground        = hb('#F4FAFF')
            tb.Opacity           = 0.45
            tb.VerticalAlignment = VerticalAlignment.Center
            tb.Padding           = Thickness(pad_left, 0, 0, 0)
            return tb

        views_hdr_left = StackPanel()
        views_hdr_left.Orientation = Orientation.Horizontal
        Grid.SetColumn(views_hdr_left, 0)
        views_hdr_left.Children.Add(_vch('', 32, 0))    # collapse-toggle placeholder
        views_hdr_left.Children.Add(_vch('View Name',  154))
        views_hdr_left.Children.Add(_vch('Sheet size', 119))
        views_hdr_left.Children.Add(_vch('Columns',     54))
        views_hdr_left.Children.Add(_vch('View Type',  134))
        views_hdr_left.Children.Add(_vch('Modified',   110))
        views_col_hdr.Children.Add(views_hdr_left)

        views_hdr_right = StackPanel()
        views_hdr_right.Orientation = Orientation.Horizontal
        views_hdr_right.HorizontalAlignment = HorizontalAlignment.Right
        Grid.SetColumn(views_hdr_right, 1)
        views_hdr_right.Children.Add(_vch('Layout',  98))
        views_hdr_right.Children.Add(_vch('',        118, 0))  # Add Section placeholder
        views_hdr_right.Children.Add(_vch('',         28, 0))  # reload placeholder
        views_hdr_right.Children.Add(_vch('',         24, 0))  # close placeholder
        views_col_hdr.Children.Add(views_hdr_right)

        inner.Children.Add(views_col_hdr)

        views_panel = StackPanel()
        views_panel.Orientation = Orientation.Vertical
        inner.Children.Add(views_panel)

        outer.Child = inner
        self.CardsPanel.Children.Add(outer)

        return {
            'outer':        outer,
            'inner':        inner,
            'views_panel':  views_panel,
            'real_path':    real_path,
            'view_keys':    [],
            'collapse_btn': collapse_btn,
            'reload_btn':   reload_btn,
            'heading_label': heading,
        }

    def _build_word_view_block(self, path):
        """Compact single-line view row nested inside a shared group:
        + Add Section, sections-collapse toggle, View Name/Sheet
        size/Columns/View Type/Modified (values only — the labels are
        the one-time views_col_hdr built in the group header), then
        Strict/Manual + Batch + this view's own reload/close, right-
        aligned. Section rows sit below, same as before."""
        fd = self._file_data[path]

        view_wrap = StackPanel()
        view_wrap.Orientation = Orientation.Vertical
        view_wrap.Margin = Thickness(0, 0, 0, 10)

        value_row = Grid()
        value_row.Margin = Thickness(0, 0, 0, 8)
        _vcol_l = ColumnDefinition(); _vcol_l.Width = GridLength(1, GridUnitType.Star)
        _vcol_r = ColumnDefinition(); _vcol_r.Width = GridLength.Auto
        value_row.ColumnDefinitions.Add(_vcol_l)
        value_row.ColumnDefinitions.Add(_vcol_r)

        value_left = StackPanel()
        value_left.Orientation = Orientation.Horizontal
        Grid.SetColumn(value_left, 0)
        value_row.Children.Add(value_left)

        # Sections collapse toggle
        sec_toggle_btn = ToggleButton()
        sec_toggle_btn.IsChecked = True
        try:
            sec_toggle_btn.Style = self.FindResource('PrimarySecondaryToggleButtonStyle')
        except Exception as e:
            logger.warning('Failed to apply PrimarySecondaryToggleButtonStyle: {}'.format(e))
        sec_toggle_btn.FocusVisualStyle = None
        sec_toggle_btn.Tag    = path
        sec_toggle_btn.VerticalAlignment = VerticalAlignment.Center
        sec_toggle_btn.Margin = Thickness(0, 0, 8, 0)
        sec_toggle_btn.Click  += self._toggle_sections_collapse
        self._set_collapse_icon(sec_toggle_btn, False)
        value_left.Children.Add(sec_toggle_btn)
        fd['sections_toggle_btn'] = sec_toggle_btn

        # View name — defaults to filename stem, user can edit
        default_vname = fd.get(
            'view_name',
            os.path.splitext(os.path.basename(fd.get('real_path', path)))[0])
        fd['view_name'] = fd.get('view_name', default_vname)

        vn_box = TextBox()
        vn_box.Text          = fd['view_name']
        vn_box.Width         = 150
        try:
            vn_box.Style = self.FindResource('TextBoxStyle')
        except Exception as e:
            logger.warning('Failed to apply TextBoxStyle: {}'.format(e))
        vn_box.VerticalAlignment = VerticalAlignment.Center
        vn_box.Margin        = Thickness(0, 0, 4, 0)
        vn_box.Tag           = path
        vn_box.LostFocus    += self._word_view_name_changed
        vn_box.TextChanged  += self._word_view_name_live_check
        fd['view_name_box'] = vn_box
        value_left.Children.Add(vn_box)
        self._style_view_name_conflict(
            vn_box, self._view_name_taken(fd['view_name'], exclude_word_path=path))

        sz_combo = ComboBox()
        self._combo_style(sz_combo, 115)
        for sz in SHEET_SIZES:
            sz_combo.Items.Add(sz)
        sz_combo.SelectedItem = fd.get('sheet_size', 'A3')
        sz_combo.Tag          = path
        sz_combo.SelectionChanged += self._word_sheet_size_changed
        value_left.Children.Add(sz_combo)

        cc_box = TextBox()
        cc_box.Text             = str(fd.get('col_count', 2))
        cc_box.Width            = 50
        cc_box.TextAlignment    = TextAlignment.Center
        try:
            cc_box.Style = self.FindResource('TextBoxStyle')
        except Exception as e:
            logger.warning('Failed to apply TextBoxStyle: {}'.format(e))
        cc_box.VerticalAlignment = VerticalAlignment.Center
        cc_box.Margin           = Thickness(0, 0, 4, 0)
        cc_box.Tag              = path
        cc_box.LostFocus       += self._word_col_count_changed
        value_left.Children.Add(cc_box)

        vt_combo = ComboBox()
        self._combo_style(vt_combo, 130)
        for vt in WORD_VIEW_TYPES:
            vt_combo.Items.Add(vt)
        fd['view_type'] = fd.get('view_type', WORD_VIEW_TYPES[0])
        vt_combo.SelectedItem = fd['view_type']
        vt_combo.Tag          = path
        vt_combo.SelectionChanged += self._card_view_type_changed
        value_left.Children.Add(vt_combo)

        # Synced — when this view was last successfully applied to
        # Revit, not the source file's own mtime (that only shows at
        # the card header level now).
        mod_text = TextBlock()
        mod_text.Text             = format_applied_at(fd.get('_applied_at'))
        mod_text.Width            = 106
        mod_text.FontSize         = 10
        mod_text.Foreground       = hb('#F4FAFF')
        mod_text.Opacity          = 0.55
        mod_text.VerticalAlignment = VerticalAlignment.Center
        mod_text.Margin           = Thickness(0, 0, 4, 0)
        fd['modified_label'] = mod_text
        value_left.Children.Add(mod_text)

        # Right side: Strict/Manual, Batch, this view's own reload/close
        right_group = StackPanel()
        right_group.Orientation = Orientation.Horizontal
        right_group.HorizontalAlignment = HorizontalAlignment.Right
        Grid.SetColumn(right_group, 1)

        layout_combo = ComboBox()
        self._combo_style(layout_combo, 90)
        layout_combo.Margin = Thickness(0, 0, 8, 0)
        for lm in ('Manual', 'Strict'):
            layout_combo.Items.Add(lm)
        layout_combo.SelectedItem = (
            'Strict' if fd.get('layout_mode', 'manual') == 'strict'
            else 'Manual')
        layout_combo.Tag = path
        layout_combo.ToolTip = (
            'Manual: set each row\'s column yourself.\n'
            'Strict: the layout algorithm assigns columns '
            'automatically, balancing section heights across '
            'the sheet. Sections never split across columns.')
        layout_combo.SelectionChanged += self._card_layout_mode_changed
        fd['layout_mode_combo'] = layout_combo
        right_group.Children.Add(layout_combo)

        add_row_btn = self._green_btn(u'+ Add Section', width=112)
        add_row_btn.Tag    = path
        add_row_btn.Click += self._add_row_for_card
        add_row_btn.Margin = Thickness(0, 0, 6, 0)
        add_row_btn.VerticalAlignment = VerticalAlignment.Center
        right_group.Children.Add(add_row_btn)

        # Per-view reload — grey/blue exactly like the group's
        # aggregate one, just scoped to this one view's own sections.
        view_reload_btn = Button()
        if _mi is not None:
            try:
                view_reload_btn.Content = _mi('reload', size=14, color='#FFFFFF')
            except Exception:
                view_reload_btn.Content = u'\u21bb'
        else:
            view_reload_btn.Content = u'\u21bb'
        try:
            view_reload_btn.Style = self.FindResource('RoundPrimaryButtonStyle')
        except Exception as e:
            logger.warning('Failed to apply RoundPrimaryButtonStyle: {}'.format(e))
        view_reload_btn.FocusVisualStyle = None
        view_reload_btn.Width  = 28
        view_reload_btn.Height = 28
        view_reload_btn.VerticalAlignment = VerticalAlignment.Center
        view_reload_btn.Margin = Thickness(0, 0, 4, 0)
        view_reload_btn.Tag    = path
        view_reload_btn.ToolTip = 'Reload sections that need updating'
        view_reload_btn.Click  += self._card_reload_click
        right_group.Children.Add(view_reload_btn)
        fd['reload_btn'] = view_reload_btn

        # Per-view close — removes just this view, not the whole card.
        view_close_btn = Button()
        view_close_btn.Content         = u'\u2715'
        view_close_btn.FontSize        = 11
        try:
            view_close_btn.Style = self.FindResource('DeleteButtonStyle')
        except Exception as e:
            logger.warning('Failed to apply DeleteButtonStyle: {}'.format(e))
        view_close_btn.FocusVisualStyle = None
        view_close_btn.Width   = 28
        view_close_btn.Height  = 28
        view_close_btn.VerticalAlignment = VerticalAlignment.Center
        view_close_btn.Tag     = path
        view_close_btn.ToolTip = 'Remove this view'
        view_close_btn.Click  += self._view_close_click
        right_group.Children.Add(view_close_btn)

        value_row.Children.Add(right_group)
        view_wrap.Children.Add(value_row)

        # Column headers + section rows
        col_hdr = StackPanel()
        col_hdr.Orientation = Orientation.Horizontal
        col_hdr.Margin      = Thickness(0, 0, 0, 8)

        def _ch(text, width, pad_left=4):
            tb = TextBlock()
            tb.Text             = text
            tb.Width            = width
            tb.FontSize         = 10
            tb.Foreground       = hb('#F4FAFF')
            tb.Opacity          = 0.45
            tb.VerticalAlignment = VerticalAlignment.Center
            tb.Padding          = Thickness(pad_left, 0, 0, 0)
            return tb

        tri_cb = CheckBox()
        tri_cb.IsThreeState      = True
        tri_cb.Margin            = Thickness(0, 0, 6, 0)
        tri_cb.VerticalAlignment = VerticalAlignment.Center
        tri_cb.Tag               = path
        tri_cb.ToolTip           = 'Select all / none rows in this view'
        tri_cb.Click            += self._card_select_all_click
        fd['select_all_cb'] = tri_cb

        col_hdr.Children.Add(_ch('',        24, 0))   # drag handle
        col_hdr.Children.Add(_ch('',        14, 0))   # status dot
        col_hdr.Children.Add(tri_cb)                  # select-all
        col_hdr.Children.Add(_ch('Section', 204))
        col_hdr.Children.Add(_ch('Priority', 82))
        col_hdr.Children.Add(_ch('Group',   104))
        col_hdr.Children.Add(_ch('Col',      40))

        row_panel = StackPanel()
        row_panel.Orientation = Orientation.Vertical

        view_wrap.Children.Add(col_hdr)
        view_wrap.Children.Add(row_panel)

        fd['card_panel']   = row_panel
        fd['sections_hdr'] = col_hdr
        fd['card_inner']   = view_wrap

        return view_wrap

    def _view_close_click(self, sender, e):
        """Remove just this one view — its own fd entry, its own UI
        block — leaving the rest of the card's other views intact.
        If it was the last view left, the whole card closes too."""
        path = sender.Tag
        fd = self._file_data.get(path)
        if fd is None:
            return
        real_path = fd.get('real_path', path)
        name = fd.get('view_name') or os.path.basename(real_path)
        if not _confirm(
                'Remove view "{}"?'.format(name),
                title='Confirm Remove'):
            return
        group = self._card_groups.get(real_path)
        if group is not None:
            view_block = fd.get('card_border')
            if view_block is not None:
                try:
                    group['views_panel'].Children.Remove(view_block)
                except Exception:
                    pass
            if path in group['view_keys']:
                group['view_keys'].remove(path)
        del self._file_data[path]
        if group is not None and not group['view_keys']:
            try:
                self.CardsPanel.Children.Remove(group['outer'])
            except Exception:
                pass
            del self._card_groups[real_path]
        if self._active_file not in self._file_data:
            paths = list(self._file_data.keys())
            self._active_file = paths[0] if paths else None
        self._update_file_combo()
        self._update_footer()
        self._refresh_status_card()
        self._refresh_drop_zone()
        self._save_persisted_state()

    def _toggle_word_group_collapse(self, sender, e):
        """Collapse/expand every view in this document's shared card.
        sender is a ToggleButton (PrimarySecondaryToggleButtonStyle) -
        WPF flips IsChecked natively before Click fires, so IsChecked
        already reflects the new state: True = now expanded."""
        real_path = sender.Tag
        group = self._card_groups.get(real_path)
        if group is None:
            return
        expanded = bool(sender.IsChecked)
        group['views_panel'].Visibility = (
            Visibility.Visible if expanded else Visibility.Collapsed)
        self._set_collapse_icon(sender, not expanded)

    def _group_reload_click(self, sender, e):
        """Reapply every stale section across every view in this
        document's card at once."""
        real_path = sender.Tag
        group = self._card_groups.get(real_path)
        if group is None:
            return
        for path in list(group['view_keys']):
            fd = self._file_data.get(path)
            if fd is None:
                continue
            stale_rows = [r for r in fd.get('rows', []) if r.Status == 'sync']
            for row in stale_rows:
                class _FakeSender(object):
                    def __init__(self, tag):
                        self.Tag = tag
                self._refresh_row_click(_FakeSender(row), None)
        self._update_word_group_reload_indicator(real_path)

    def _update_word_group_reload_indicator(self, real_path):
        """Enable (accent green) when at least one section anywhere in
        this card needs reapplying, disable (canonical disabled look)
        when every view is in sync - same IsEnabled-driven
        RoundPrimaryButtonStyle as the row/card-level reload buttons."""
        group = self._card_groups.get(real_path)
        if group is None or group.get('reload_btn') is None:
            return
        needs_reload = False
        for path in group['view_keys']:
            fd = self._file_data.get(path)
            if fd and any(r.Status == 'sync' for r in fd.get('rows', [])):
                needs_reload = True
                break
        btn = group['reload_btn']
        btn.IsEnabled = needs_reload
        btn.ToolTip = ('Click to reload views that need updating'
                        if needs_reload else 'All views up to date')

    def _group_close_click(self, sender, e):
        """Remove this document's whole card — every view, every
        section, gone. Individual views are removed via that view's
        own Batch menu, not this button."""
        real_path = sender.Tag
        group = self._card_groups.get(real_path)
        if group is None:
            return
        name = os.path.basename(real_path)
        n = len(group['view_keys'])
        if not _confirm(
                'Remove card for {}?\nThis deletes all {} view(s) and '
                'their sections.'.format(name, n),
                title='Confirm Remove'):
            return
        for path in list(group['view_keys']):
            if path in self._file_data:
                del self._file_data[path]
        try:
            self.CardsPanel.Children.Remove(group['outer'])
        except Exception:
            pass
        del self._card_groups[real_path]
        if self._active_file not in self._file_data:
            paths = list(self._file_data.keys())
            self._active_file = paths[0] if paths else None
        self._update_file_combo()
        self._update_footer()
        self._refresh_status_card()
        self._refresh_drop_zone()
        self._save_persisted_state()

    def _make_word_row_ui(self, row):
        """Build the row StackPanel for a Word section row."""
        sp = StackPanel()
        sp.Orientation = Orientation.Horizontal
        sp.Height      = 32
        sp.Margin      = Thickness(0, 0, 0, 4)
        sp.Tag         = row
        # Store panel ref on row for drag reorder
        row._drag_panel_ref = None  # set by caller after insertion

        # Drag handle :::
        drag_lbl = TextBlock()
        drag_lbl.Text             = u'∷'   # ⠿ grid dots
        drag_lbl.FontSize         = 14
        drag_lbl.Foreground       = hb('#F4FAFF')
        drag_lbl.Opacity          = 0.35
        drag_lbl.Width            = 20
        drag_lbl.VerticalAlignment = VerticalAlignment.Center
        drag_lbl.Margin           = Thickness(0, 0, 4, 0)
        drag_lbl.Tag              = row
        drag_lbl.Cursor           = __import__(
            'System.Windows.Input', fromlist=['Cursors']).Cursors.SizeNS
        drag_lbl.PreviewMouseLeftButtonDown += self._drag_start
        sp.Children.Add(drag_lbl)

        # Status dot
        dot = Ellipse()
        dot.Width             = 8
        dot.Height            = 8
        dot.Fill              = hb(STATUS_COLOURS.get(row.Status, '#3A4A3A'))
        dot.VerticalAlignment = VerticalAlignment.Center
        dot.Margin            = Thickness(0, 0, 6, 0)
        dot.Tag               = row
        row._dot              = dot
        sp.Children.Add(dot)

        # Checkbox
        cb = CheckBox()
        cb.IsChecked         = row.Enabled
        cb.VerticalAlignment = VerticalAlignment.Center
        cb.Margin            = Thickness(0, 0, 6, 0)
        cb.Tag               = row
        cb.Click            += self._cb_click
        row._enabled_cb      = cb
        sp.Children.Add(cb)

        # Section picker combo (headings from the docx)
        sc = ComboBox()
        self._combo_style(sc, 200)
        for h in row.ranges_for():
            sc.Items.Add(h)
        if row.NamedRange and row.NamedRange in [
                sc.Items.GetItemAt(i) for i in range(sc.Items.Count)]:
            sc.SelectedItem = row.NamedRange
        elif sc.Items.Count > 0:
            sc.SelectedIndex = 0
            row.NamedRange   = sc.Items.GetItemAt(0)
        if not row.Group and row.NamedRange:
            row.Group = match_section_group(row.NamedRange)
        sc.Tag               = row
        sc.SelectionChanged += self._word_section_changed
        sp.Children.Add(sc)

        # Priority — how much the layout algorithm can reorder this
        # row relative to its neighbours. Only relevant (and only
        # shown) once the card is in Strict layout mode.
        _owner_path = self._find_card_path_for_row(row)
        fd_for_mode = self._file_data.get(_owner_path, {}) if _owner_path else {}
        strict_mode = fd_for_mode.get('layout_mode', 'manual') == 'strict'
        combo_vis   = Visibility.Visible if strict_mode else Visibility.Collapsed

        pr_combo = ComboBox()
        self._combo_style(pr_combo, 78)
        for p in ('High', 'Medium', 'Low'):
            pr_combo.Items.Add(p)
        pr_combo.SelectedItem = row.Priority
        pr_combo.Tag          = row
        pr_combo.Margin       = Thickness(0, 0, 4, 0)
        pr_combo.Visibility   = combo_vis
        pr_combo.ToolTip      = (
            'High: fixed order, never reordered.\n'
            'Medium: can reorder within its own run of Medium rows.\n'
            'Low: placed wherever best balances the columns.')
        pr_combo.SelectionChanged += self._row_priority_changed
        row._priority_combo   = pr_combo
        sp.Children.Add(pr_combo)

        # Group — sections sharing a group are kept together and move
        # as one block. Options come from userdata/section_groups.json.
        gr_combo = ComboBox()
        self._combo_style(gr_combo, 100)
        self._populate_group_combo(gr_combo, row.Group)
        gr_combo.Tag          = row
        gr_combo.Margin       = Thickness(0, 0, 4, 0)
        gr_combo.Visibility   = combo_vis
        gr_combo.ToolTip      = 'Sections in the same group are kept together by the layout algorithm.'
        gr_combo.SelectionChanged += self._row_group_changed
        row._group_combo      = gr_combo
        sp.Children.Add(gr_combo)

        # Col number textbox
        col_tb = TextBox()
        col_tb.Text          = str(row.ColNo)
        col_tb.Width         = 36
        try:
            col_tb.Style = self.FindResource('TextBoxStyle')
        except Exception as e:
            logger.warning('Failed to apply TextBoxStyle: {}'.format(e))
        col_tb.Margin        = Thickness(0, 0, 4, 0)
        col_tb.VerticalAlignment = VerticalAlignment.Center
        col_tb.Tag           = row
        col_tb.LostFocus    += self._word_col_no_changed
        col_tb.IsEnabled     = not strict_mode
        col_tb.Opacity       = 0.5 if strict_mode else 1.0
        row._col_textbox     = col_tb
        sp.Children.Add(col_tb)

        # View type now lives on the card (View Type dropdown next to
        # Columns), not per row. Pull whatever the card is currently
        # set to.
        fd_for_row = self._file_data.get(_owner_path, {}) if _owner_path else {}
        card_view_type = fd_for_row.get('view_type', WORD_VIEW_TYPES[0])
        if row.ViewType not in WORD_VIEW_TYPES:
            row.ViewType = card_view_type

        rb = self._make_sync_btn(row)
        row._refresh_btn    = rb
        sp.Children.Add(rb)

        # Delete button — always-red DeleteButtonStyle, not the transparent-
        # until-hover CloseButtonStyle (this removes the row, it isn't a
        # window/card close action). Sized smaller than the card-level
        # delete button (24 vs the style's own 30px token) since this one
        # sits inline in a dense row, not a card header.
        db = Button()
        db.Content          = u'✕'
        db.FontSize         = 10
        try:
            db.Style = self.FindResource('DeleteButtonStyle')
        except Exception as e:
            logger.warning('Failed to apply DeleteButtonStyle: {}'.format(e))
        db.FocusVisualStyle = None
        db.Width            = 24
        db.Height           = 24
        db.Cursor           = __import__(
            'System.Windows.Input',
            fromlist=['Cursors']).Cursors.Hand
        db.VerticalAlignment = VerticalAlignment.Center
        db.Margin           = Thickness(4, 0, 0, 0)
        db.Tag              = row
        db.Click           += self._del_click
        sp.Children.Add(db)

        # Error pill
        err_pill = Border()
        err_pill.Background      = hb('#3B1515')
        err_pill.BorderBrush     = hb('#DC2626')
        err_pill.BorderThickness = Thickness(1)
        err_pill.CornerRadius    = CornerRadius(4)
        err_pill.Padding         = Thickness(8, 2, 8, 2)
        err_pill.Margin          = Thickness(6, 0, 0, 0)
        err_pill.VerticalAlignment = VerticalAlignment.Center
        err_pill.Visibility      = Visibility.Collapsed
        err_txt = TextBlock()
        err_txt.FontSize         = 10
        err_txt.Foreground       = hb('#F87171')
        err_txt.VerticalAlignment = VerticalAlignment.Center
        err_txt.Text             = ''
        err_pill.Child           = err_txt
        row._error_label         = err_pill
        row._error_text          = err_txt
        sp.Children.Add(err_pill)

        return sp

    def _word_view_name_live_check(self, sender, e):
        """Live (as-you-type) conflict check for the Word card's View
        Name box."""
        path = sender.Tag
        name = sender.Text.strip()
        taken = self._view_name_taken(name, exclude_word_path=path)
        self._style_view_name_conflict(sender, taken)

    def _parse_word(self, path):
        real_path = path
        if path in self._file_data:
            path = self._next_duplicate_key(real_path)
        # Extract headings from docx so section picker is populated
        headings = []
        try:
            headings = get_word_headings(real_path)
        except Exception as ex:
            logger.warning('pyTable word parse: {}'.format(ex))
        if not headings:
            headings = ['(no headings found)']
        srmap = {'Document': headings}
        self._file_data[path] = {
            'sheets':          ['Document'],
            'sheet_range_map': srmap,
            'source_type':     'word',
            'rows':            [],
            'card_panel':      None,
            'card_border':     None,
            'real_path':       real_path,
            'sheet_size':      'A3 Landscape',
            'col_count':       2,
        }
        self._active_file = path
        self._make_card(path)

    # ── Toolbar handlers ──

    def _word_view_name_changed(self, sender, e):
        path = sender.Tag
        if path and path in self._file_data:
            self._file_data[path]['view_name'] = sender.Text.strip()
            self._auto_check_card(path)
            self._revalidate_all_view_name_boxes()
            self._save_persisted_state()

    def _word_section_changed(self, sender, e):
        row = sender.Tag
        if row is None:
            return
        row.NamedRange = sender.SelectedItem or ''
        if not row.Group and row.NamedRange:
            matched = match_section_group(row.NamedRange)
            if matched:
                row.Group = matched
                if row._group_combo is not None:
                    row._group_combo.SelectedItem = matched
        self._auto_check_row(row)
        self._save_persisted_state()

    def _word_col_no_changed(self, sender, e):
        row = sender.Tag
        if row is None:
            return
        try:
            row.ColNo = max(1, int(sender.Text.strip()))
        except Exception:
            row.ColNo = 1
        sender.Text = str(row.ColNo)
        self._auto_check_row(row)
        self._save_persisted_state()

    def _populate_group_combo(self, combo, current):
        """Fill a row's Group combo from userdata/section_groups.json
        plus the fixed '(none)' and '+ New group…' options."""
        combo.Items.Clear()
        combo.Items.Add('(none)')
        data = load_section_groups()
        names = [g.get('name', '') for g in data.get('groups', []) if g.get('name')]
        for n in names:
            combo.Items.Add(n)
        combo.Items.Add(u'+ New group\u2026')
        if current and current in names:
            combo.SelectedItem = current
        else:
            combo.SelectedItem = '(none)'

    def _refresh_all_group_combos(self):
        """After a group is added/renamed, refresh every row's Group
        combo across every card so the new option shows up everywhere."""
        for fd in self._file_data.values():
            for row in fd.get('rows', []):
                if row._group_combo is not None:
                    self._populate_group_combo(row._group_combo, row.Group)

    def _open_word_text_settings_editor(self):
        """Default: pyTable uses its own named type, 'pyTable Notes
        2.0 Arial', creating it if the project doesn't have one yet.
        Manual: pick an exact existing TextNoteType from a dropdown
        of everything already in the project."""
        from System.Windows import (
            Window, SizeToContent, WindowStartupLocation, ResizeMode,
            TextWrapping)
        from System.Windows.Controls import RadioButton

        settings = load_word_text_settings()
        default_name = 'pyTable Notes 2.0 Arial'
        default_size = 2.0

        # Fetch existing project TextNoteTypes live, sorted by size —
        # same list Revit's own Properties panel type-picker shows.
        existing_types = []
        try:
            for tt in DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType).ToElements():
                try:
                    name = tt.get_Parameter(
                        DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                    size_ft = tt.get_Parameter(
                        DB.BuiltInParameter.TEXT_SIZE).AsDouble()
                    size_mm = size_ft / (1.0 / 304.8)
                    existing_types.append((name, size_mm))
                except Exception:
                    continue
        except Exception:
            pass
        existing_types.sort(key=lambda t: t[1])

        w = Window()
        w.Title = 'Word Text Size'
        w.Width = 380
        w.SizeToContent = SizeToContent.Height
        w.WindowStartupLocation = WindowStartupLocation.CenterOwner
        try:
            w.Owner = self
        except Exception:
            pass
        w.Background = hb('#2B3340')
        w.ResizeMode = ResizeMode.NoResize

        root = StackPanel()
        root.Margin = Thickness(16)
        w.Content = root

        title = TextBlock()
        title.Text       = 'Word Text Size'
        title.FontSize   = 14
        title.FontWeight = FontWeights.Bold
        title.Foreground = hb('#F4FAFF')
        title.Margin     = Thickness(0, 0, 0, 4)
        root.Children.Add(title)

        hint = TextBlock()
        hint.Text = 'Which TextNoteType Word notes get placed with.'
        hint.Foreground   = hb('#F4FAFF')
        hint.Opacity      = 0.6
        hint.FontSize     = 10
        hint.TextWrapping = TextWrapping.Wrap
        hint.Margin       = Thickness(0, 0, 0, 12)
        root.Children.Add(hint)

        auto_rb = RadioButton()
        auto_rb.GroupName   = 'word_text_size_mode'
        auto_rb.Content     = u'Default: {}'.format(default_name)
        auto_rb.Foreground  = hb('#F4FAFF')
        auto_rb.FontSize    = 12
        auto_rb.Margin      = Thickness(0, 0, 0, 6)
        auto_rb.IsChecked   = settings.get('mode', 'auto') == 'auto'
        root.Children.Add(auto_rb)

        manual_row = StackPanel()
        manual_row.Orientation = Orientation.Horizontal
        manual_row.Margin = Thickness(0, 0, 0, 14)

        manual_rb = RadioButton()
        manual_rb.GroupName  = 'word_text_size_mode'
        manual_rb.Content    = 'Manual, type:'
        manual_rb.Foreground = hb('#F4FAFF')
        manual_rb.FontSize   = 12
        manual_rb.VerticalAlignment = VerticalAlignment.Center
        manual_rb.IsChecked  = settings.get('mode', 'auto') == 'manual'
        manual_row.Children.Add(manual_rb)

        type_combo = ComboBox()
        self._combo_style(type_combo, 180)
        type_combo.Margin = Thickness(8, 0, 0, 0)
        if not existing_types:
            type_combo.IsEnabled = False
            type_combo.Items.Add('(no text types in this project)')
            type_combo.SelectedIndex = 0
        else:
            for name, size_mm in existing_types:
                type_combo.Items.Add(u'{}  ({:.1f}mm)'.format(name, size_mm))
            current_name = settings.get('text_type_name')
            names = [n for n, _ in existing_types]
            if current_name and current_name in names:
                type_combo.SelectedIndex = names.index(current_name)
            else:
                type_combo.SelectedIndex = 0
        manual_row.Children.Add(type_combo)
        root.Children.Add(manual_row)

        btn_row = StackPanel()
        btn_row.Orientation = Orientation.Horizontal
        btn_row.HorizontalAlignment = HorizontalAlignment.Right

        def _save(s, ev):
            if manual_rb.IsChecked:
                mode = 'manual'
                idx = type_combo.SelectedIndex
                if existing_types and 0 <= idx < len(existing_types):
                    chosen_name, chosen_size = existing_types[idx]
                else:
                    chosen_name, chosen_size = (
                        settings.get('text_type_name'),
                        settings.get('size_mm', default_size))
            else:
                mode = 'auto'
                chosen_name, chosen_size = default_name, default_size
            save_word_text_settings({
                'mode': mode,
                'text_type_name': chosen_name,
                'size_mm': chosen_size,
            })
            w.Close()

        save_btn = self._green_btn(u'Save')
        save_btn.Click += _save
        btn_row.Children.Add(save_btn)

        cancel_btn = Button()
        cancel_btn.Content    = u'Cancel'
        cancel_btn.Height     = 24
        cancel_btn.Padding    = Thickness(10, 0, 10, 0)
        cancel_btn.FontSize   = 11
        try:
            cancel_btn.Style = self.FindResource('SecondaryButtonStyle')
        except Exception as e:
            logger.warning('Failed to apply SecondaryButtonStyle: {}'.format(e))
        cancel_btn.FocusVisualStyle = None
        cancel_btn.VerticalAlignment = VerticalAlignment.Center
        cancel_btn.Margin = Thickness(8, 0, 0, 0)
        cancel_btn.Click += lambda s, ev: w.Close()
        btn_row.Children.Add(cancel_btn)

        root.Children.Add(btn_row)
        w.ShowDialog()

    def _open_group_settings_editor(self):
        """Master-detail list editor for userdata/section_groups.json:
        search/select a group from the list, then add/remove
        keywords for just that group, plus add/delete whole groups.
        This is the deliberate 'go to settings' path; the per-row
        Group dropdown's '+ New group…' is the convenience path for
        tagging one section on the fly."""
        from System.Windows import (
            Window, SizeToContent, WindowStartupLocation, ResizeMode,
            TextWrapping)
        from System.Windows.Controls import ListBox

        data = load_section_groups()
        # Working copy — Cancel just closes without touching the file.
        work = [{'name': g.get('name', ''), 'keywords': list(g.get('keywords', []))}
                for g in data.get('groups', [])]
        selected_idx = [None]

        w = Window()
        w.Title = 'Section Groups'
        w.Width = 460
        w.SizeToContent = SizeToContent.Height
        w.WindowStartupLocation = WindowStartupLocation.CenterOwner
        try:
            w.Owner = self
        except Exception:
            pass
        w.Background = hb('#2B3340')
        w.ResizeMode = ResizeMode.NoResize

        root = StackPanel()
        root.Margin = Thickness(16)
        w.Content = root

        title = TextBlock()
        title.Text       = 'Section Groups'
        title.FontSize   = 14
        title.FontWeight = FontWeights.Bold
        title.Foreground = hb('#F4FAFF')
        title.Margin     = Thickness(0, 0, 0, 4)
        root.Children.Add(title)

        hint = TextBlock()
        hint.Text = ('Sections whose text contains one of a group\'s '
                     'keywords are auto-tagged into it. Select a '
                     'group below to see and edit its keywords.')
        hint.Foreground   = hb('#F4FAFF')
        hint.Opacity      = 0.6
        hint.FontSize     = 10
        hint.TextWrapping = TextWrapping.Wrap
        hint.Margin       = Thickness(0, 0, 0, 12)
        root.Children.Add(hint)

        def _list_box(height):
            lb = ListBox()
            lb.Height          = height
            lb.Background      = hb('#232933')
            lb.Foreground      = hb('#F4FAFF')
            lb.BorderBrush     = hb('#404553')
            lb.BorderThickness = Thickness(1)
            return lb

        # ── Groups: search + list + add/delete ──
        search_label = TextBlock()
        search_label.Text       = 'Search'
        search_label.Foreground = hb('#F4FAFF')
        search_label.Opacity    = 0.7
        search_label.FontSize   = 11
        search_label.Margin     = Thickness(0, 0, 0, 4)
        root.Children.Add(search_label)

        search_tb = TextBox()
        search_tb.Margin = Thickness(0, 0, 0, 6)
        try:
            search_tb.Style = self.FindResource('TextBoxStyle')
        except Exception as e:
            logger.warning('Failed to apply TextBoxStyle: {}'.format(e))
        root.Children.Add(search_tb)

        groups_list = _list_box(120)
        root.Children.Add(groups_list)

        def _refresh_groups_list():
            filt = search_tb.Text.strip().lower()
            groups_list.Items.Clear()
            for g in work:
                if filt and filt not in g['name'].lower():
                    continue
                groups_list.Items.Add(g['name'])
        _refresh_groups_list()
        search_tb.TextChanged += lambda s, ev: _refresh_groups_list()

        group_btn_row = StackPanel()
        group_btn_row.Orientation = Orientation.Horizontal
        group_btn_row.Margin      = Thickness(0, 6, 0, 16)

        add_group_btn = self._green_btn(u'+ Add Group')
        add_group_btn.Margin = Thickness(0, 0, 8, 0)

        # "Delete" text button - no canonical rectangular-danger style
        # exists yet, so this stays SecondaryButtonStyle with its
        # Background overridden to the canonical danger brush (not a
        # hand-picked hex) rather than inventing a whole new style.
        del_group_btn = Button()
        del_group_btn.Content    = u'Delete'
        del_group_btn.Height     = 24
        del_group_btn.Padding    = Thickness(12, 0, 12, 0)
        del_group_btn.FontSize   = 11
        try:
            del_group_btn.Style = self.FindResource('SecondaryButtonStyle')
        except Exception as e:
            logger.warning('Failed to apply SecondaryButtonStyle: {}'.format(e))
        try:
            del_group_btn.Background = self.FindResource('BrushDanger')
        except Exception as e:
            logger.warning('Failed to apply BrushDanger: {}'.format(e))
        del_group_btn.FocusVisualStyle = None

        group_btn_row.Children.Add(add_group_btn)
        group_btn_row.Children.Add(del_group_btn)
        root.Children.Add(group_btn_row)

        # ── Keywords for whichever group is selected above ──
        kw_label = TextBlock()
        kw_label.Text       = 'Keywords in this group:'
        kw_label.Foreground = hb('#F4FAFF')
        kw_label.Opacity    = 0.7
        kw_label.FontSize   = 11
        kw_label.Margin     = Thickness(0, 0, 0, 4)
        root.Children.Add(kw_label)

        keywords_list = _list_box(110)
        root.Children.Add(keywords_list)

        def _refresh_keywords_list():
            keywords_list.Items.Clear()
            idx = selected_idx[0]
            if idx is None:
                return
            for kw in work[idx]['keywords']:
                keywords_list.Items.Add(kw)
        _refresh_keywords_list()

        def _on_group_selected(s, ev):
            name = groups_list.SelectedItem
            selected_idx[0] = None
            if name is not None:
                for i, g in enumerate(work):
                    if g['name'] == name:
                        selected_idx[0] = i
                        break
            _refresh_keywords_list()
        groups_list.SelectionChanged += _on_group_selected

        kw_add_row = StackPanel()
        kw_add_row.Orientation = Orientation.Horizontal
        kw_add_row.Margin      = Thickness(0, 6, 0, 6)

        new_kw_tb = TextBox()
        new_kw_tb.Width  = 220
        new_kw_tb.Margin = Thickness(0, 0, 8, 0)
        try:
            new_kw_tb.Style = self.FindResource('TextBoxStyle')
        except Exception as e:
            logger.warning('Failed to apply TextBoxStyle: {}'.format(e))
        kw_add_row.Children.Add(new_kw_tb)

        add_kw_btn = self._green_btn(u'+ Add')
        kw_add_row.Children.Add(add_kw_btn)
        root.Children.Add(kw_add_row)

        del_kw_btn = Button()
        del_kw_btn.Content    = u'Remove selected keyword'
        del_kw_btn.Height     = 24
        del_kw_btn.Padding    = Thickness(12, 0, 12, 0)
        del_kw_btn.FontSize   = 11
        try:
            del_kw_btn.Style = self.FindResource('SecondaryButtonStyle')
        except Exception as e:
            logger.warning('Failed to apply SecondaryButtonStyle: {}'.format(e))
        del_kw_btn.FocusVisualStyle = None
        del_kw_btn.Margin = Thickness(0, 0, 0, 16)
        root.Children.Add(del_kw_btn)

        def _on_add_group(s, ev):
            name = None
            if sdlg is not None:
                try:
                    name = sdlg.ask_string(
                        'Name for the new group:', title='Add Group')
                except Exception:
                    name = None
            if not name:
                return
            name = name.strip()
            if any(g['name'] == name for g in work):
                _alert('A group named "{}" already exists.'.format(name),
                       title='Add Group')
                return
            work.append({'name': name, 'keywords': []})
            _refresh_groups_list()
        add_group_btn.Click += _on_add_group

        def _on_delete_group(s, ev):
            idx = selected_idx[0]
            if idx is None:
                _alert('Select a group to delete first.', title='Delete Group')
                return
            if not _confirm(
                    'Delete group "{}"?'.format(work[idx]['name']),
                    title='Confirm Delete'):
                return
            del work[idx]
            selected_idx[0] = None
            _refresh_groups_list()
            _refresh_keywords_list()
        del_group_btn.Click += _on_delete_group

        def _on_add_kw(s, ev):
            idx = selected_idx[0]
            if idx is None:
                _alert('Select a group first.', title='Add Keyword')
                return
            text = new_kw_tb.Text.strip()
            if not text:
                return
            if text not in work[idx]['keywords']:
                work[idx]['keywords'].append(text)
            new_kw_tb.Text = ''
            _refresh_keywords_list()
        add_kw_btn.Click += _on_add_kw

        def _on_del_kw(s, ev):
            idx = selected_idx[0]
            if idx is None:
                return
            sel_kw = keywords_list.SelectedItem
            if sel_kw is None:
                _alert('Select a keyword to remove first.', title='Remove Keyword')
                return
            try:
                work[idx]['keywords'].remove(sel_kw)
            except ValueError:
                pass
            _refresh_keywords_list()
        del_kw_btn.Click += _on_del_kw

        btn_row = StackPanel()
        btn_row.Orientation = Orientation.Horizontal
        btn_row.HorizontalAlignment = HorizontalAlignment.Right

        def _save(s, ev):
            cleaned = [g for g in work if g['name'].strip()]
            save_section_groups({'groups': cleaned})
            self._refresh_all_group_combos()
            w.Close()

        save_btn = self._green_btn(u'Save')
        save_btn.Click += _save
        btn_row.Children.Add(save_btn)

        cancel_btn = Button()
        cancel_btn.Content    = u'Cancel'
        cancel_btn.Height     = 24
        cancel_btn.Padding    = Thickness(10, 0, 10, 0)
        cancel_btn.FontSize   = 11
        try:
            cancel_btn.Style = self.FindResource('SecondaryButtonStyle')
        except Exception as e:
            logger.warning('Failed to apply SecondaryButtonStyle: {}'.format(e))
        cancel_btn.FocusVisualStyle = None
        cancel_btn.VerticalAlignment = VerticalAlignment.Center
        cancel_btn.Margin = Thickness(8, 0, 0, 0)
        cancel_btn.Click += lambda s, ev: w.Close()
        btn_row.Children.Add(cancel_btn)

        root.Children.Add(btn_row)
        w.ShowDialog()

    def _row_priority_changed(self, sender, e):
        row = sender.Tag
        if row is None:
            return
        row.Priority = sender.SelectedItem or 'Medium'
        self._auto_check_row(row)
        self._save_persisted_state()
        path = self._find_card_path_for_row(row)
        if path:
            self._maybe_run_strict_layout(path)

    def _row_group_changed(self, sender, e):
        row = sender.Tag
        if row is None:
            return
        sel = sender.SelectedItem

        if sel == u'+ New group\u2026':
            name = None
            if sdlg is not None:
                try:
                    name = sdlg.ask_string(
                        'Name for the new section group:',
                        title='New Group')
                except Exception:
                    name = None
            if not name:
                # Cancelled — revert the combo to whatever it was.
                self._populate_group_combo(sender, row.Group)
                return
            name = name.strip()
            data = load_section_groups()
            existing = [g.get('name') for g in data.get('groups', [])]
            if name not in existing:
                data.setdefault('groups', []).append({
                    'name':     name,
                    'keywords': [row.NamedRange] if row.NamedRange else [],
                })
                save_section_groups(data)
            row.Group = name
            self._refresh_all_group_combos()

        elif sel == '(none)':
            row.Group = ''

        else:
            row.Group = sel or ''
            # Learn this section's exact phrase into the chosen group,
            # so a similarly-worded section auto-matches next time —
            # no guessed tokenization, just the literal text the user
            # confirmed belongs to this group.
            if row.NamedRange:
                data = load_section_groups()
                for g in data.get('groups', []):
                    if g.get('name') == row.Group:
                        kws = g.setdefault('keywords', [])
                        if row.NamedRange not in kws:
                            kws.append(row.NamedRange)
                            save_section_groups(data)
                        break

        self._auto_check_row(row)
        self._save_persisted_state()
        path = self._find_card_path_for_row(row)
        if path:
            self._maybe_run_strict_layout(path)

    def _card_layout_mode_changed(self, sender, e):
        path = sender.Tag
        fd = self._file_data.get(path)
        if fd is None:
            return
        fd['layout_mode'] = 'strict' if sender.SelectedItem == 'Strict' else 'manual'
        self._apply_layout_mode_ui(path)
        if fd['layout_mode'] == 'strict':
            self._run_strict_layout(path)
        self._auto_check_card(path)
        self._save_persisted_state()

    def _apply_layout_mode_ui(self, path):
        """Show/hide Priority+Group and lock/unlock Col for every row
        in a card, matching its current layout_mode."""
        fd = self._file_data.get(path)
        if fd is None:
            return
        strict = fd.get('layout_mode', 'manual') == 'strict'
        vis = Visibility.Visible if strict else Visibility.Collapsed
        for row in fd.get('rows', []):
            if row._priority_combo is not None:
                row._priority_combo.Visibility = vis
            if row._group_combo is not None:
                row._group_combo.Visibility = vis
            if row._col_textbox is not None:
                row._col_textbox.IsEnabled = not strict
                row._col_textbox.Opacity   = 0.5 if strict else 1.0

    def _maybe_run_strict_layout(self, path):
        """Re-run the layout algorithm only if this card is actually
        in Strict mode — a no-op otherwise, safe to call after any
        Priority/Group/row change without checking mode first."""
        fd = self._file_data.get(path)
        if fd is not None and fd.get('layout_mode') == 'strict':
            self._run_strict_layout(path)

    def _run_strict_layout(self, path):
        """Assign Col numbers automatically for every row in a Strict-
        mode card. Sections are atomic (never split across columns —
        that's Split mode, not yet built). Same-Group rows are merged
        into one packing unit and placed together. Priority governs
        how much a section can be reordered from document order:
        High = fixed sequential waterfall, Medium = reorderable within
        its own run (bounded by neighbouring High sections), Low =
        placed last, wherever balances the columns best.
        """
        fd = self._file_data.get(path)
        if fd is None or fd.get('source_type') != 'word':
            return
        rows = [r for r in fd.get('rows', []) if r.NamedRange]
        if not rows:
            return
        col_count = max(1, int(fd.get('col_count', 2) or 2))
        sheet_size = fd.get('sheet_size', 'A3 Landscape')
        real_path  = fd.get('real_path', path)

        # Pull the section's real paragraph text so the height estimate
        # reflects actual word-wrap, not just a bullet count.
        try:
            sections = read_word_sections(real_path)
        except Exception as ex:
            logger.warning('Strict layout: could not read {}: {}'.format(
                real_path, ex))
            sections = []
        para_map = {}
        for s in sections:
            para_map.setdefault(s.get('heading', ''), []).extend(
                p.get('text', '') for p in s.get('paragraphs', []))

        sheet_w = SHEET_WIDTHS_MM.get(sheet_size, 420.0)
        usable  = sheet_w - 2.0 * NOTE_MARGIN_MM
        col_w   = (usable - (col_count - 1) * GAP_MM) / col_count
        chars_per_line = max(
            10, int(col_w / (_ARIAL_W_DEF * NOTE_TEXT_SIZE_MM * 0.85)))

        def section_height(row):
            h = NOTE_LINE_HEIGHT_MM * 1.3  # heading line
            texts = para_map.get(row.NamedRange, [])
            if not texts:
                # No parsed content available (e.g. heading renamed,
                # file not yet re-scanned) — fall back to a one-line
                # placeholder so it still gets placed somewhere.
                return h + NOTE_LINE_HEIGHT_MM
            for text in texts:
                n_lines = max(1, -(-len(text) // chars_per_line))  # ceil
                h += n_lines * NOTE_LINE_HEIGHT_MM + NOTE_BULLET_GAP_MM
            return h

        # ── Step 1: collapse same-Group rows into single packing units ──
        priority_rank = {'High': 0, 'Medium': 1, 'Low': 2}
        units = []
        seen = set()
        for row in rows:
            if id(row) in seen:
                continue
            if row.Group:
                group_rows = [r for r in rows if r.Group == row.Group]
            else:
                group_rows = [row]
            for r in group_rows:
                seen.add(id(r))
            unit_priority = min(
                (r.Priority for r in group_rows),
                key=lambda p: priority_rank.get(p, 1))
            units.append({
                'rows':     group_rows,
                'height':   sum(section_height(r) for r in group_rows),
                'priority': unit_priority,
            })

        # ── Step 2: partition into runs, preserving document order ──
        runs = []
        medium_buffer = []

        def _flush_medium():
            if medium_buffer:
                runs.append(('medium_run', list(medium_buffer)))
                del medium_buffer[:]

        low_units = []
        for u in units:
            if u['priority'] == 'High':
                _flush_medium()
                runs.append(('high', u))
            elif u['priority'] == 'Medium':
                medium_buffer.append(u)
            else:
                low_units.append(u)
        _flush_medium()

        # ── Step 3: placement ──
        total_height = sum(u['height'] for u in units) or 1.0
        target = total_height / col_count
        col_heights = [0.0] * col_count
        assignment = {}
        high_col = [0]   # mutable pointer for the High waterfall

        def _place_high(u):
            col = high_col[0]
            if col_heights[col] >= target and col < col_count - 1:
                col += 1
                high_col[0] = col
            for r in u['rows']:
                assignment[id(r)] = col
            col_heights[col] += u['height']

        def _place_balanced(unit_list):
            # Longest-processing-time greedy bin packing: biggest
            # units first, always into the currently-shortest column.
            for u in sorted(unit_list, key=lambda x: -x['height']):
                col = min(range(col_count), key=lambda c: col_heights[c])
                for r in u['rows']:
                    assignment[id(r)] = col
                col_heights[col] += u['height']
                if col_heights[col] >= target and col >= high_col[0]:
                    high_col[0] = min(col + 1, col_count - 1)

        for kind, payload in runs:
            if kind == 'high':
                _place_high(payload)
            else:
                _place_balanced(payload)
        _place_balanced(low_units)

        # ── Step 4: write results back (1-based Col) ──
        for row in rows:
            col0 = assignment.get(id(row))
            if col0 is not None:
                row.ColNo = col0 + 1
                if row._col_textbox is not None:
                    row._col_textbox.Text = str(row.ColNo)

        self._save_persisted_state()

    def _word_sheet_size_changed(self, sender, e):
        path = sender.Tag
        if path and path in self._file_data:
            self._file_data[path]['sheet_size'] = (
                sender.SelectedItem or 'A3')
            self._auto_check_card(path)
            self._save_persisted_state()
            self._maybe_run_strict_layout(path)

    def _word_col_count_changed(self, sender, e):
        path = sender.Tag
        if path and path in self._file_data:
            try:
                n = max(1, int(sender.Text.strip()))
            except Exception:
                n = 2
            self._file_data[path]['col_count'] = n
            sender.Text = str(n)
            self._auto_check_card(path)
            self._save_persisted_state()
            self._maybe_run_strict_layout(path)

    def _card_view_type_changed(self, sender, e):
        """View Type now lives on the card, changing it applies to
        every row currently in that card."""
        if sender.SelectedItem is None:
            return
        path = sender.Tag
        fd = self._file_data.get(path)
        if fd is None:
            return
        fd['view_type'] = sender.SelectedItem
        for row in fd.get('rows', []):
            row.ViewType = sender.SelectedItem
        self._auto_check_card(path)
        self._save_persisted_state()

    # ── Drag reorder (Word rows) ──

    def _drag_indicator_show(self, panel, idx):
        """Position or create the blue drop-line indicator in *panel*."""
        # Create indicator Border on first call
        if getattr(self, '_drag_indicator', None) is None:
            ind = Border()
            ind.Height          = 2
            ind.Background      = hb('#3B82F6')
            ind.IsHitTestVisible = False   # don't interfere with mouse events
            ind.Margin          = Thickness(0)
            self._drag_indicator       = ind
            self._drag_indicator_panel = None

        ind = self._drag_indicator

        # Remove from previous panel if different
        if (getattr(self, '_drag_indicator_panel', None) is not None
                and self._drag_indicator_panel is not panel):
            try:
                self._drag_indicator_panel.Children.Remove(ind)
            except Exception:
                pass

        # Insert/move in current panel
        self._drag_indicator_panel = panel
        try:
            panel.Children.Remove(ind)
        except Exception:
            pass
        insert_at = min(idx, panel.Children.Count)
        try:
            panel.Children.Insert(insert_at, ind)
        except Exception:
            pass

    def _drag_indicator_remove(self):
        """Remove the drop-line indicator from whichever panel it's in."""
        ind = getattr(self, '_drag_indicator', None)
        if ind is None:
            return
        panel = getattr(self, '_drag_indicator_panel', None)
        if panel is not None:
            try:
                panel.Children.Remove(ind)
            except Exception:
                pass
        self._drag_indicator       = None
        self._drag_indicator_panel = None

    def _drag_start(self, sender, e):
        """Record drag origin and wire move/up events on the window."""
        row = sender.Tag
        if row is None:
            return
        self._dragging_row    = row
        self._drag_start_y    = e.GetPosition(self).Y
        self._drag_indicator  = None
        self._drag_target_idx = None
        sender.CaptureMouse()
        sender.PreviewMouseMove             += self._drag_move
        sender.PreviewMouseLeftButtonUp     += self._drag_drop
        e.Handled = True

    def _drag_move(self, sender, e):
        """Update drop-line position as the user drags."""
        if not getattr(self, '_dragging_row', None):
            return
        row = self._dragging_row
        for path, fd in self._file_data.items():
            if row not in fd['rows']:
                continue
            panel   = fd['card_panel']
            mouse_y = e.GetPosition(panel).Y
            idx     = 0
            for i in range(panel.Children.Count):
                child = panel.Children[i]
                # Skip the indicator itself when measuring
                if child is getattr(self, '_drag_indicator', None):
                    continue
                try:
                    pt = child.TransformToAncestor(panel).Transform(
                        __import__('System.Windows', fromlist=['Point'])
                        .Point(0, child.ActualHeight / 2.0))
                    if mouse_y > pt.Y:
                        idx = i + 1
                except Exception:
                    pass
            self._drag_target_idx = idx
            self._drag_indicator_show(panel, idx)
            break
        e.Handled = True

    def _drag_drop(self, sender, e):
        """Reorder the row to the computed insertion index."""
        row = getattr(self, '_dragging_row', None)
        if row is None:
            return
        sender.ReleaseMouseCapture()
        sender.PreviewMouseMove         -= self._drag_move
        sender.PreviewMouseLeftButtonUp -= self._drag_drop
        self._drag_indicator_remove()
        target_idx = getattr(self, '_drag_target_idx', None)
        self._dragging_row = None
        if target_idx is None:
            return

        for path, fd in self._file_data.items():
            if row not in fd['rows']:
                continue
            rows  = fd['rows']
            panel = fd['card_panel']
            cur   = rows.index(row)
            if cur == target_idx or cur + 1 == target_idx:
                return
            # Find the row's StackPanel child
            row_sp = None
            for child in list(panel.Children):
                if getattr(child, 'Tag', None) is row:
                    row_sp = child
                    break
            if row_sp is None:
                return
            # Reorder in data list
            rows.remove(row)
            ins = target_idx if target_idx <= cur else target_idx - 1
            rows.insert(ins, row)
            # Reorder in UI panel
            panel.Children.Remove(row_sp)
            panel.Children.Insert(ins, row_sp)
            self._save_persisted_state()
            break
        e.Handled = True

    def _toggle_sections_collapse(self, sender, e):
        """Show/hide just this view's section rows (col_hdr + the row
        list), leaving the header and View Name/Sheet size/Columns/
        View Type row untouched — narrower than the whole-card
        collapse, for tucking away a long section list while still
        seeing/editing the view's own settings. sender is a ToggleButton
        (PrimarySecondaryToggleButtonStyle) - WPF flips IsChecked
        natively before Click fires, so IsChecked already reflects the
        new state: True = now expanded."""
        path = sender.Tag
        fd = self._file_data.get(path)
        if fd is None:
            return
        col_hdr   = fd.get('sections_hdr')
        row_panel = fd.get('card_panel')
        if col_hdr is None or row_panel is None:
            return
        expanded = bool(sender.IsChecked)
        col_hdr.Visibility   = Visibility.Visible if expanded else Visibility.Collapsed
        row_panel.Visibility = Visibility.Visible if expanded else Visibility.Collapsed
        self._set_collapse_icon(sender, not expanded)

    def _card_add_view(self, sender, e):
        """'+ Add Row' at the card header, Word cards only: create a
        new independent VIEW for this same document — same underlying
        file, its own View Name/Sheet size/Columns/View Type, and its
        own empty section list ready for '+ Add Section'. A View, in
        pyTable terms, is the layout (one Legend/Drafting view); a
        Section is a heading placed into it. This reuses the exact
        same real_path-keyed duplicate-card mechanism as the Batch
        menu's Duplicate action — a second independent card IS a
        second View, just without cloning the first one's sections."""
        path = sender.Tag
        fd = self._file_data.get(path)
        if fd is None or fd.get('source_type') != 'word':
            return
        real_path = fd.get('real_path', path)
        new_key = self._next_duplicate_key(real_path)

        new_fd = {
            'sheets':          list(fd.get('sheets', [])),
            'sheet_range_map': dict(fd.get('sheet_range_map', {})),
            'source_type':     'word',
            'rows':            [],
            'card_panel':      None,
            'card_border':     None,
            'real_path':       real_path,
            'unlinked':        False,
            'path_mode':       'absolute',
            'sheet_size':      fd.get('sheet_size', 'A3 Landscape'),
            'col_count':       fd.get('col_count', 2),
            'view_name':       '',
            'view_type':       fd.get('view_type', WORD_VIEW_TYPES[0]),
        }
        self._file_data[new_key] = new_fd
        self._make_card(new_key)
        self._update_file_combo()
        self._update_footer()
        self._refresh_status_card()
        self._revalidate_all_view_name_boxes()
        self._save_persisted_state()
        self._set_status('Added new row')

