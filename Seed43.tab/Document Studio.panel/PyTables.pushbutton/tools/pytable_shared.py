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
pytable_shared.py -- genuinely cross-cutting pieces used by BOTH the
Excel and Word sides of pyTable: the Row class, colour/status
constants, the hb() colour helper, and Revit shared-parameter state
persistence (save/load). Deliberately has NO dependency on
pytable_excel.py or pytable_word.py, so both of those (and the main
PyTable.py) can import from here with zero circularity risk.
"""


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

PYTABLE_PARAM_GUID = 'f0a46d4c-c148-4ff4-95c8-9750eec5d480'
PYTABLE_PARAM_NAME = 'pyTable'
PYTABLE_PARAM_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), '..', 'pyTable.txt')


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

def _find_seed43_version():
    """Walk up from this pushbutton to Seed43.extension/version.txt and
    return just the version string (its first line). Returns 'unknown'
    if the file can't be found or read."""
    # tools/ is one level deeper than the pushbutton root the original
    # walk started from, so start one level higher to compensate.
    folder = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    for _ in range(6):
        candidate = os.path.join(folder, 'version.txt')
        if os.path.isfile(candidate):
            try:
                with open(candidate, 'r') as f:
                    return f.readline().strip()
            except Exception:
                return 'unknown'
        folder = os.path.dirname(folder)
    return 'unknown'


# ── Constants ──
VIEW_TYPE_LEGEND   = 'Legend View'
VIEW_TYPE_SCHEDULE = 'Schedule View'
VIEW_TYPE_DRAFTING = 'Drafting View'


# ── Data class ──

def _run_export_script(script_name, payload):
    """
    Run one of the Export/ scripts via exec() with PYTABLE_PAYLOAD injected.
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
        'PYTABLE_PAYLOAD': payload,
    }

    src = open(script_path, 'r').read()

    # These scripts manage their own transactions internally
    if script_name in ('create_legend.py', 'create_notes.py'):
        exec(src, ns)
    else:
        with revit.Transaction(
            'pyTable - {}'.format(
                payload.get('view_name', script_name)
            )
        ):
            exec(src, ns)

    # Export scripts may leave a PYTABLE_RESULT dict in their own
    # namespace (e.g. {'view_id': view.Id.IntegerValue}) so the caller
    # can tag the actual view that was created — used for the
    # ElementId-based ownership tracking.
    return ns.get('PYTABLE_RESULT')

def _ensure_pytable_param_bound():
    """
    Ensure the pyTable shared parameter is bound to BOTH Project
    Information (the punch list of everything pyTable manages) and
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
            app.SharedParametersFilename = PYTABLE_PARAM_FILE
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
                if d.Name == PYTABLE_PARAM_NAME:
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
            with revit.Transaction('pyTable - bind parameter'):
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
        logger.warning('pyTable param bind failed: {}'.format(ex))
        return None

def get_view_pytable_data(view):
    """
    Read a view's own pyTable record (sheet/range/hash/scale/etc) from
    its per-view parameter — the authoritative record for that one
    view, as opposed to Project Information's punch list, which is
    just an index of which views to go look up. Returns a dict of
    whatever tokens are present, or None if the view has no record
    (never touched by pyTable, or the parameter can't be read).
    """
    try:
        p = view.LookupParameter(PYTABLE_PARAM_NAME)
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

def set_view_pytable_data(view, **kwargs):
    """
    Write this view's own pyTable record. Keyword args become K-V
    tokens on the view's per-view parameter, e.g.
    set_view_pytable_data(view, SH='Sheet1', RG='Temp', H=hash_str).
    Creates/expands the parameter binding on first use. Returns True
    on success.
    """
    try:
        p = view.LookupParameter(PYTABLE_PARAM_NAME)
        if p is None:
            _ensure_pytable_param_bound()
            p = view.LookupParameter(PYTABLE_PARAM_NAME)
        if p is None:
            return False
        text = '|'.join(
            '{}-{}'.format(k, v) for k, v in kwargs.items() if v is not None
        )
        with revit.Transaction('pyTable - tag view'):
            p.Set(text)
        return True
    except Exception as ex:
        logger.warning('pyTable view tag failed: {}'.format(ex))
        return False

def _get_pytable_param():
    """
    Get or create the pyTable shared parameter on ProjectInfo — this
    is the punch list: an index of what pyTable manages, not the
    per-view records themselves (see get_view_pytable_data /
    set_view_pytable_data for those). Returns the Parameter object or
    None.
    """
    try:
        proj_info = doc.ProjectInformation
        if proj_info is None:
            return None
        p = proj_info.LookupParameter(PYTABLE_PARAM_NAME)
        if p is not None:
            return p
        _ensure_pytable_param_bound()
        return proj_info.LookupParameter(PYTABLE_PARAM_NAME)
    except Exception as ex:
        logger.warning('pyTable param get failed: {}'.format(ex))
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
    was saved somewhere pyTable can't see) is returned as-is — the
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

def save_pytable_state(file_data):
    """
    Serialise pyTable UI state to the shared parameter on ProjectInfo.

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
            lines.append('VN-{}|S-{}|R-{}|VT-{}|MT-{}|H-{}|CN-{}|PR-{}|GR-{}|AT-{}|AVN-{}|VS-{}|VID-{}'.format(
                row.ViewName, row.Sheet, row.NamedRange,
                row.ViewType, mt, h, cn, pr, gr, at, avn, vs, vid))
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
        if ss:
            lines.append('CARD_SS-{}|CC-{}|VN-{}|VT-{}|RP-{}|UL-{}|PM-{}|LM-{}|AVN-{}|CAT-{}'.format(
                ss, cc, vn, vt, rp_stored, ul, pm, lm, avn, card_at))
        elif rp != path or fd.get('unlinked') or pm != 'absolute' or lm != 'manual' or avn:
            # Excel duplicate/unlinked cards have no sheet_size line,
            # still need real_path (and now unlink/path-mode/layout
            # mode/applied view name) recorded so they round-trip on
            # reload.
            lines.append('CARD_RP-{}|UL-{}|PM-{}|LM-{}|AVN-{}|CAT-{}'.format(
                rp_stored, ul, pm, lm, avn, card_at))
    text = '\n'.join(lines)
    try:
        p = _get_pytable_param()
        if p is not None:
            with revit.Transaction('pyTable - save state'):
                p.Set(text)
            logger.debug('pyTable state saved ({} chars)'.format(len(text)))
    except Exception as ex:
        logger.warning('pyTable save failed: {}'.format(ex))

def load_pytable_state():
    """
    Read pyTable state from the shared parameter.

    Returns list of dicts:
        [{'path': str, 'rows': [{'view_name', 'sheet', 'named_range', 'view_type'}]}]

    Any card whose path (or real_path) was stored relative gets
    resolved back to absolute here, against the current .rvt's
    folder, before the caller ever sees it.
    """
    try:
        p = _get_pytable_param()
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
                    # Backfill: state saved before this tracker existed
                    # has no AVN token. Assume the view hasn't been
                    # renamed since the last apply — the safest guess
                    # for legacy state, and harmless if wrong (the
                    # rename lookup just falls through to the old
                    # search-then-create behaviour).
                    current['applied_view_name'] = (
                        parts.get('AVN') or current['view_name'])
                    cat = parts.get('CAT', '')
                    current['applied_at'] = float(cat) if cat else None
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
        logger.warning('pyTable load failed: {}'.format(ex))
        return []

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
