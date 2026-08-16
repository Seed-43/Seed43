# -*- coding: utf-8 -*-
# pyfilter_save.py
# Seed43 Filter Manager - grid rows, popup dialogs, save logic
# pylint: disable=import-error,invalid-name,broad-except

import datetime
import io
import json
import os

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System import Int64
from System.Windows import (
    Window, WindowStartupLocation, Thickness,
    GridLength, GridUnitType, HorizontalAlignment, VerticalAlignment,
    Style, Setter, Trigger, CornerRadius,
    FontWeights, TextTrimming, TextWrapping,
)
from System.Windows.Controls import (
    Grid, ColumnDefinition, RowDefinition,
    StackPanel, DockPanel, ScrollViewer, ScrollBarVisibility,
    TextBox, TextBlock, CheckBox, ComboBox, ComboBoxItem,
    Border, Button, Label, Separator,
)
from System.Windows.Media import SolidColorBrush, Color, Colors
from System.Windows.Input import Cursors

from pyrevit import revit, DB, forms
from Snippets._revisions import safe_str

doc = revit.doc

# ── COLUMN LAYOUT ─────────────────────────────────────────────────────────────

# Module-level store for deferred preview callbacks, keyed by id(grid).
# WPF Grid doesn't allow arbitrary Python attribute assignment, so we keep
# the list here and let _open_template look it up by id(grid) after build.
_deferred_preview_store = {}

# ── COLUMN WIDTH PERSISTENCE ───────────────────────────────────────────────────
# Stored in a local settings.json next to script.py (see pyfilter_localcfg.py)
# instead of pyRevit's shared global config file.

import pyfilter_localcfg as _localcfg

# pyfilter_save.py lives in <pushbutton_root>/lib/, so the pushbutton root
# (where settings.json should live, next to script.py) is one level up.
_SETTINGS_ANCHOR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Default widths indexed to match COLUMNS order (no assign column).
_DEFAULT_WIDTHS = [200, 80, 80, 80, 80, 80, 80, 80, 80]
_CFG_SECTION = "columns"
_CFG_OPT     = "col_widths"

def get_col_widths():
    """Return saved column widths list, falling back to defaults."""
    try:
        data = _localcfg.get_option(_SETTINGS_ANCHOR, _CFG_SECTION, _CFG_OPT)
        if isinstance(data, list) and len(data) == len(_DEFAULT_WIDTHS):
            return [max(20, int(w)) for w in data]
    except Exception:
        pass
    return list(_DEFAULT_WIDTHS)

def set_col_widths(widths):
    """Persist column widths."""
    try:
        _localcfg.set_option(_SETTINGS_ANCHOR, _CFG_SECTION, _CFG_OPT,
                             [int(w) for w in widths])
    except Exception:
        pass

def reset_col_widths():
    """Restore defaults and persist."""
    set_col_widths(_DEFAULT_WIDTHS)
    return list(_DEFAULT_WIDTHS)

# Registry of live ColumnDefinition objects across all visible rows.
# Each entry is a list of ColumnDefinition objects (one per column) for one row.
# Updated by build_display_row; cleared when FilterRowsPanel is cleared.
_live_row_col_defs = []

def register_row_col_defs(col_defs):
    """Called by build_display_row to register its ColumnDefinition list."""
    _live_row_col_defs.append(col_defs)

def clear_row_col_defs():
    """Called by _open_template before rebuilding rows."""
    del _live_row_col_defs[:]

def update_all_row_widths(widths):
    """Push new widths into every live row's ColumnDefinitions in-place."""
    for col_defs in _live_row_col_defs:
        for i, w in enumerate(widths):
            if i < len(col_defs):
                col_defs[i].Width = GridLength(w)
# Matches Revit VG column structure exactly
# (key, header, width, type)
# type: label | check | lines_btn | pats_btn | transp | halftone
COLUMNS = [
    ("name",             "Name",             180, "label"),
    ("enabled",          "Enable Filter",     80, "check"),
    ("visible",          "Visibility",        70, "check"),
    ("proj_lines",       "Lines",            120, "lines_btn"),   # proj line overrides
    ("proj_patterns",    "Patterns",         120, "pats_btn"),    # proj surface patterns
    ("transparency",     "Transparency",      90, "transp"),
    ("cut_lines",        "Lines",            120, "lines_btn"),   # cut line overrides
    ("cut_patterns",     "Patterns",         120, "pats_btn"),    # cut patterns
    ("halftone",         "Halftone",          70, "halftone"),
]

# Group header spans: (label, start_col, span)
GROUPS = [
    ("",                    0, 3),
    ("Projection/Surface",  3, 3),
    ("Cut",                 6, 2),
    ("",                    8, 1),
]

# Assignment mode prepends an "Add" checkbox column at index 0.
# Used when the sidebar is listing live Views / View Templates instead of
# saved filter templates.
ASSIGN_COLUMN = ("assign", "Add", 46, "assign")

def _cols_for(assign_mode):
    return ([ASSIGN_COLUMN] + COLUMNS) if assign_mode else COLUMNS

def _groups_for(assign_mode):
    # Only the rendered (non-blank) groups need shifting; blanks are skipped.
    if not assign_mode:
        return GROUPS
    return [(lbl, start + 1, span) for (lbl, start, span) in GROUPS]

# ── COLOURS ───────────────────────────────────────────────────────────────────
C_PANEL    = Color.FromRgb(43,  51,  64)
C_PANEL2   = Color.FromRgb(50,  61,  77)
C_HDR      = Color.FromRgb(30,  37,  48)
C_FG       = Color.FromRgb(244, 250, 255)
C_DIM      = Color.FromRgb(122, 138, 154)
C_BORDER   = Color.FromRgb(64,  69,  83)
C_GREEN    = Color.FromRgb(32,  138,  60)
C_SEL      = Color.FromRgb(26,  74,  46)
C_INPUT_BG = Color.FromRgb(244, 250, 255)
C_INPUT_FG = Color.FromRgb(43,  51,  64)
C_DANGER   = Color.FromRgb(197,  48,  48)

def br(c):  return SolidColorBrush(c)

# ── HELPERS ───────────────────────────────────────────────────────────────────

def _hex_to_brush(hex_str):
    if not hex_str or not hex_str.startswith("#") or len(hex_str) != 7:
        return SolidColorBrush(Color.FromArgb(0, 0, 0, 0))
    try:
        r = int(hex_str[1:3], 16)
        g = int(hex_str[3:5], 16)
        b = int(hex_str[5:7], 16)
        return SolidColorBrush(Color.FromRgb(r, g, b))
    except Exception:
        return SolidColorBrush(Color.FromArgb(0, 0, 0, 0))

def _combo_style():
    s = Style()
    s.TargetType = ComboBoxItem
    s.Setters.Add(Setter(ComboBoxItem.BackgroundProperty, br(C_PANEL)))
    s.Setters.Add(Setter(ComboBoxItem.ForegroundProperty, br(C_FG)))
    s.Setters.Add(Setter(ComboBoxItem.PaddingProperty, Thickness(6, 3, 6, 3)))
    h = Trigger()
    h.Property = ComboBoxItem.IsMouseOverProperty
    h.Value    = True
    h.Setters.Add(Setter(ComboBoxItem.BackgroundProperty, br(C_GREEN)))
    s.Triggers.Add(h)
    sel = Trigger()
    sel.Property = ComboBoxItem.IsSelectedProperty
    sel.Value    = True
    sel.Setters.Add(Setter(ComboBoxItem.BackgroundProperty,
        SolidColorBrush(Color.FromRgb(26, 110, 46))))
    s.Triggers.Add(sel)
    return s

_COMBO_STYLE = None
def get_combo_style():
    global _COMBO_STYLE
    if _COMBO_STYLE is None:
        _COMBO_STYLE = _combo_style()
    return _COMBO_STYLE

# ── REVIT DATA ────────────────────────────────────────────────────────────────

def get_all_filters():
    return sorted(
        DB.FilteredElementCollector(doc).OfClass(DB.ParameterFilterElement),
        key=lambda f: f.Name)

def get_fill_patterns():
    result = [("<No Override>", None)]
    for fp in DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement):
        result.append((safe_str(fp.Name), fp.Id))
    return result

def get_line_patterns():
    result = [("<No Override>", None)]
    for lp in DB.FilteredElementCollector(doc).OfClass(DB.LinePatternElement):
        result.append((safe_str(lp.Name), lp.Id))
    return result

def color_to_hex(c):
    """Convert a DB.Color to a hex string.
    Revit uses InvalidColorValue (255,255,255 with IsValid=False) for 'no colour'.
    In 2026 IsValid can return False even for valid colours so we read channels
    directly and treat (0,0,0) with IsValid=False as no-colour.
    """
    if c is None:
        return None
    try:
        def _ch(v):
            try:   return int(v) & 0xFF
            except Exception: return int(v.__int__()) & 0xFF
        r, g, b = _ch(c.Red), _ch(c.Green), _ch(c.Blue)
        # If IsValid is explicitly False AND all channels are zero, treat as unset.
        try:
            if not c.IsValid and r == 0 and g == 0 and b == 0:
                return None
        except Exception:
            pass
        return "#{:02X}{:02X}{:02X}".format(r, g, b)
    except Exception:
        return None

def pat_id_to_str(pid):
    try:
        if hasattr(pid, "Value"):        return str(pid.Value)
        if hasattr(pid, "IntegerValue"): return str(pid.IntegerValue)
    except Exception: pass
    return None

def pat_id_to_dict(pid, doc_ref=None):
    """Serialise a pattern ElementId as {"id": "3", "name": "Solid"} so it
    can be resolved by name in files where the ID differs."""
    id_str = pat_id_to_str(pid)
    if id_str is None:
        return None
    name = None
    if doc_ref is not None:
        try:
            from pyrevit import DB as _DB
            from System import Int64 as _I64
            eid = _DB.ElementId(_I64(int(id_str)))
            el  = doc_ref.GetElement(eid)
            if el is not None:
                name = safe_str(el.Name)
        except Exception:
            pass
    return {"id": id_str, "name": name}

def serialise_filter_def(f):
    try:
        cat_ids = sorted([
            str(getattr(cid, "Value", None) or getattr(cid, "IntegerValue", 0))
            for cid in f.GetCategories()])
    except Exception:
        cat_ids = []
    try:   rules_str = safe_str(f.GetElementFilter().ToString())
    except Exception: rules_str = ""
    return {"name": safe_str(f.Name), "cats": cat_ids, "rules": rules_str}

def get_filter_settings_from_view(f, view):
    from pyrevit import revit as _rv
    _doc = _rv.doc
    fid = f.Id
    s   = {}
    try: s["enabled"] = view.GetIsFilterEnabled(fid)
    except Exception: s["enabled"] = True
    try: s["visible"] = view.GetFilterVisibility(fid)
    except Exception: s["visible"] = True
    try:
        ogs = view.GetFilterOverrides(fid)
        def sc(a):
            try: return color_to_hex(getattr(ogs, a))
            except Exception: return None
        def sp(a):
            try: return pat_id_to_dict(getattr(ogs, a), _doc)
            except Exception: return None
        def si(a):
            try: return getattr(ogs, a)
            except Exception: return None
        def sb(a):
            try: return getattr(ogs, a)
            except Exception: return None
        s.update({
            "proj_line_color":      sc("ProjectionLineColor"),
            "proj_line_weight":     si("ProjectionLineWeight"),
            "proj_line_pattern_id": sp("ProjectionLinePatternId"),
            "surf_fg_color":        sc("SurfaceForegroundPatternColor"),
            "surf_fg_pat":          sp("SurfaceForegroundPatternId"),
            "surf_fg_visible":      sb("IsSurfaceForegroundPatternVisible"),
            "surf_bg_color":        sc("SurfaceBackgroundPatternColor"),
            "surf_bg_pat":          sp("SurfaceBackgroundPatternId"),
            "surf_bg_visible":      sb("IsSurfaceBackgroundPatternVisible"),
            "cut_line_color":       sc("CutLineColor"),
            "cut_line_weight":      si("CutLineWeight"),
            "cut_line_pattern_id":  sp("CutLinePatternId"),
            "cut_fg_color":         sc("CutForegroundPatternColor"),
            "cut_fg_pat":           sp("CutForegroundPatternId"),
            "cut_fg_visible":       sb("IsCutForegroundPatternVisible"),
            "cut_bg_color":         sc("CutBackgroundPatternColor"),
            "cut_bg_pat":           sp("CutBackgroundPatternId"),
            "cut_bg_visible":       sb("IsCutBackgroundPatternVisible"),
            "halftone":             sb("Halftone"),
            "transparency":         si("Transparency"),
        })
    except Exception: pass
    return s

# ── POPUP DIALOGS ─────────────────────────────────────────────────────────────

class LineGraphicsDialog(Window):
    """
    Popup for editing line overrides (Pattern, Color, Weight).
    Matches Revit's Line Graphics dialog layout.
    """
    def __init__(self, title, color_key, weight_key, pattern_key,
                 settings, line_patterns, owner=None):
        self.Title  = title
        self.Width  = 340
        self.Height = 240
        self.ResizeMode            = _resize_none()
        self.WindowStartupLocation = WindowStartupLocation.CenterOwner
        self.Background            = br(C_PANEL)
        if owner:
            self.Owner = owner

        self._color_key   = color_key
        self._weight_key  = weight_key
        self._pattern_key = pattern_key
        self._settings    = dict(settings)
        self._line_pats   = line_patterns
        self.accepted     = False

        root = Grid()
        root.Margin = Thickness(16)
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, _star())))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(40)))
        self.Content = root

        form = Grid()
        form.RowDefinitions.Add(RowDefinition(Height=GridLength(34)))
        form.RowDefinitions.Add(RowDefinition(Height=GridLength(34)))
        form.RowDefinitions.Add(RowDefinition(Height=GridLength(34)))
        form.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(90)))
        form.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, _star())))
        Grid.SetRow(form, 0)
        root.Children.Add(form)

        # Pattern
        self._add_label(form, "Pattern:", 0, 0)
        self._pat_combo = self._make_combo(line_patterns, settings.get(pattern_key))
        Grid.SetRow(self._pat_combo, 0)
        Grid.SetColumn(self._pat_combo, 1)
        form.Children.Add(self._pat_combo)

        # Color
        self._add_label(form, "Color:", 1, 0)
        color_row = DockPanel()
        color_row.LastChildFill = True
        hex_val = settings.get(color_key) or ""
        self._swatch = Border()
        self._swatch.Width           = 22
        self._swatch.Height          = 22
        self._swatch.CornerRadius    = CornerRadius(3)
        self._swatch.BorderThickness = Thickness(1)
        self._swatch.BorderBrush     = br(C_BORDER)
        self._swatch.Background      = _hex_to_brush(hex_val)
        self._swatch.Margin          = Thickness(0, 0, 6, 0)
        self._swatch.VerticalAlignment = VerticalAlignment.Center
        from System.Windows.Controls import Dock as WpfDock
        DockPanel.SetDock(self._swatch, WpfDock.Left)
        self._color_tb = self._make_textbox(hex_val)
        def on_color_change(sender, e):
            self._swatch.Background = _hex_to_brush(sender.Text.strip())
        self._color_tb.TextChanged += on_color_change
        color_row.Children.Add(self._swatch)
        color_row.Children.Add(self._color_tb)
        color_row.VerticalAlignment = VerticalAlignment.Center
        Grid.SetRow(color_row, 1)
        Grid.SetColumn(color_row, 1)
        form.Children.Add(color_row)

        # Weight
        self._add_label(form, "Weight:", 2, 0)
        w = settings.get(weight_key)
        self._weight_tb = self._make_textbox(str(w) if w is not None else "-1")
        Grid.SetRow(self._weight_tb, 2)
        Grid.SetColumn(self._weight_tb, 1)
        form.Children.Add(self._weight_tb)

        # Buttons
        btn_row = StackPanel()
        btn_row.Orientation          = _horizontal()
        btn_row.HorizontalAlignment  = HorizontalAlignment.Right
        btn_row.VerticalAlignment    = VerticalAlignment.Center
        ok_btn  = self._make_btn("OK",     self._on_ok,     green=True)
        cl_btn  = self._make_btn("Clear",  self._on_clear)
        ca_btn  = self._make_btn("Cancel", self._on_cancel)
        btn_row.Children.Add(cl_btn)
        btn_row.Children.Add(ca_btn)
        btn_row.Children.Add(ok_btn)
        Grid.SetRow(btn_row, 1)
        root.Children.Add(btn_row)

    def _add_label(self, parent, text, row, col):
        tb = TextBlock()
        tb.Text      = text
        tb.Foreground = br(C_FG)
        tb.FontSize  = 12
        tb.VerticalAlignment = VerticalAlignment.Center
        tb.Margin    = Thickness(0, 0, 8, 0)
        Grid.SetRow(tb, row)
        Grid.SetColumn(tb, col)
        parent.Children.Add(tb)

    def _make_textbox(self, text):
        tb = TextBox()
        tb.Text       = text
        tb.Background = br(C_INPUT_BG)
        tb.Foreground = br(C_INPUT_FG)
        tb.BorderBrush = br(C_GREEN)
        tb.BorderThickness = Thickness(1)
        tb.FontSize   = 11
        tb.Padding    = Thickness(6, 3, 6, 3)
        tb.Height     = 26
        tb.Margin     = Thickness(0, 4, 0, 4)
        tb.VerticalContentAlignment = VerticalAlignment.Center
        return tb

    def _make_combo(self, pats, stored_id):
        cb = ComboBox()
        cb.Background  = br(C_INPUT_BG)
        cb.Foreground  = br(C_INPUT_FG)
        cb.BorderBrush = br(C_GREEN)
        cb.FontSize    = 11
        cb.Height      = 26
        cb.Margin      = Thickness(0, 4, 0, 4)
        cb.ItemContainerStyle = get_combo_style()
        sel = 0
        for idx, (pname, pid) in enumerate(pats):
            item = ComboBoxItem()
            item.Content = pname
            cb.Items.Add(item)
            if stored_id and pid is not None:
                pid_str = str(pid.Value if hasattr(pid, "Value") else pid.IntegerValue)
                if pid_str == str(stored_id):
                    sel = idx
        cb.SelectedIndex = sel
        return cb

    def _make_btn(self, text, handler, green=False):
        btn = Button()
        btn.Content    = text
        btn.Background = br(C_GREEN if green else Color.FromRgb(64, 69, 83))
        btn.Foreground = br(C_FG)
        btn.BorderThickness = Thickness(0)
        btn.Padding    = Thickness(14, 6, 14, 6)
        btn.FontSize   = 12
        btn.Margin     = Thickness(4, 0, 0, 0)
        btn.Cursor     = Cursors.Hand
        btn.Click     += handler
        return btn

    def _on_ok(self, sender, e):
        self.accepted = True
        # Read color
        hex_val = self._color_tb.Text.strip()
        self._settings[self._color_key] = hex_val if hex_val else None
        # Read weight
        try:   self._settings[self._weight_key] = int(self._weight_tb.Text.strip())
        except Exception: self._settings[self._weight_key] = -1
        # Read pattern
        idx  = self._pat_combo.SelectedIndex
        pats = self._line_pats
        if idx <= 0 or idx >= len(pats):
            self._settings[self._pattern_key] = None
        else:
            pname, pid = pats[idx]
            self._settings[self._pattern_key] = (
                str(pid.Value if hasattr(pid, "Value") else pid.IntegerValue)
                if pid else None)
        self.Close()

    def _on_clear(self, sender, e):
        self.accepted = True
        self._settings[self._color_key]   = None
        self._settings[self._weight_key]  = -1
        self._settings[self._pattern_key] = None
        self.Close()

    def _on_cancel(self, sender, e):
        self.Close()

    def get_settings(self):
        return self._settings


class FillPatternDialog(Window):
    """
    Popup for editing surface/cut pattern overrides.
    Matches Revit's Fill Pattern Graphics dialog.
    """
    def __init__(self, title, fg_color_key, fg_pat_key, fg_vis_key,
                 bg_color_key, bg_pat_key, bg_vis_key,
                 settings, fill_patterns, owner=None):
        self.Title  = title
        self.Width  = 360
        self.Height = 340
        self.ResizeMode            = _resize_none()
        self.WindowStartupLocation = WindowStartupLocation.CenterOwner
        self.Background            = br(C_PANEL)
        if owner:
            self.Owner = owner

        self._keys     = (fg_color_key, fg_pat_key, fg_vis_key,
                          bg_color_key, bg_pat_key, bg_vis_key)
        self._settings = dict(settings)
        self._fill_pats = fill_patterns
        self.accepted  = False

        root = Grid()
        root.Margin = Thickness(16)
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, _star())))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(40)))
        self.Content = root

        form = StackPanel()
        Grid.SetRow(form, 0)
        root.Children.Add(form)

        # Foreground section
        fg_hdr = self._section_label("Foreground")
        form.Children.Add(fg_hdr)
        self._fg_vis   = self._add_check(form, "Visible",  settings.get(fg_vis_key, True))
        self._fg_combo = self._add_combo(form, "Pattern",  settings.get(fg_pat_key))
        self._fg_color_tb, self._fg_swatch = self._add_color(form, "Color", settings.get(fg_color_key))

        sep = Border()
        sep.Height = 1
        sep.Background = br(C_BORDER)
        sep.Margin = Thickness(0, 8, 0, 8)
        form.Children.Add(sep)

        # Background section
        bg_hdr = self._section_label("Background")
        form.Children.Add(bg_hdr)
        self._bg_vis   = self._add_check(form, "Visible",  settings.get(bg_vis_key, True))
        self._bg_combo = self._add_combo(form, "Pattern",  settings.get(bg_pat_key))
        self._bg_color_tb, self._bg_swatch = self._add_color(form, "Color", settings.get(bg_color_key))

        # Buttons
        btn_row = StackPanel()
        btn_row.Orientation         = _horizontal()
        btn_row.HorizontalAlignment = HorizontalAlignment.Right
        btn_row.VerticalAlignment   = VerticalAlignment.Center
        ok_btn = self._make_btn("OK",     self._on_ok,     green=True)
        cl_btn = self._make_btn("Clear",  self._on_clear)
        ca_btn = self._make_btn("Cancel", self._on_cancel)
        btn_row.Children.Add(cl_btn)
        btn_row.Children.Add(ca_btn)
        btn_row.Children.Add(ok_btn)
        Grid.SetRow(btn_row, 1)
        root.Children.Add(btn_row)

    def _section_label(self, text):
        tb = TextBlock()
        tb.Text      = text
        tb.Foreground = br(C_FG)
        tb.FontSize  = 12
        tb.FontWeight = FontWeights.SemiBold
        tb.Margin    = Thickness(0, 4, 0, 4)
        return tb

    def _add_check(self, parent, label, value):
        row = StackPanel()
        row.Orientation = _horizontal()
        row.Margin      = Thickness(0, 2, 0, 2)
        cb = CheckBox()
        cb.IsChecked = bool(value) if value is not None else True
        cb.Foreground = br(C_FG)
        cb.VerticalAlignment = VerticalAlignment.Center
        tb = TextBlock()
        tb.Text      = label
        tb.Foreground = br(C_DIM)
        tb.FontSize  = 11
        tb.Margin    = Thickness(6, 0, 0, 0)
        tb.VerticalAlignment = VerticalAlignment.Center
        row.Children.Add(cb)
        row.Children.Add(tb)
        parent.Children.Add(row)
        return cb

    def _add_combo(self, parent, label, stored_id):
        row = DockPanel()
        row.LastChildFill = True
        row.Margin        = Thickness(0, 2, 0, 2)
        from System.Windows.Controls import Dock as WpfDock
        lbl = TextBlock()
        lbl.Text      = label + ":"
        lbl.Foreground = br(C_DIM)
        lbl.FontSize  = 11
        lbl.Width     = 70
        lbl.VerticalAlignment = VerticalAlignment.Center
        DockPanel.SetDock(lbl, WpfDock.Left)
        cb = ComboBox()
        cb.Background  = br(C_INPUT_BG)
        cb.Foreground  = br(C_INPUT_FG)
        cb.BorderBrush = br(C_GREEN)
        cb.FontSize    = 11
        cb.Height      = 24
        cb.ItemContainerStyle = get_combo_style()
        sel = 0
        for idx, (pname, pid) in enumerate(self._fill_pats):
            item = ComboBoxItem()
            item.Content = pname
            cb.Items.Add(item)
            if stored_id and pid is not None:
                pid_str = str(pid.Value if hasattr(pid, "Value") else pid.IntegerValue)
                if pid_str == str(stored_id):
                    sel = idx
        cb.SelectedIndex = sel
        row.Children.Add(lbl)
        row.Children.Add(cb)
        parent.Children.Add(row)
        return cb

    def _add_color(self, parent, label, hex_val):
        row = DockPanel()
        row.LastChildFill = True
        row.Margin        = Thickness(0, 2, 0, 2)
        from System.Windows.Controls import Dock as WpfDock
        lbl = TextBlock()
        lbl.Text      = label + ":"
        lbl.Foreground = br(C_DIM)
        lbl.FontSize  = 11
        lbl.Width     = 70
        lbl.VerticalAlignment = VerticalAlignment.Center
        DockPanel.SetDock(lbl, WpfDock.Left)

        swatch = Border()
        swatch.Width           = 20
        swatch.Height          = 20
        swatch.CornerRadius    = CornerRadius(3)
        swatch.BorderThickness = Thickness(1)
        swatch.BorderBrush     = br(C_BORDER)
        swatch.Background      = _hex_to_brush(hex_val or "")
        swatch.VerticalAlignment = VerticalAlignment.Center
        swatch.Margin          = Thickness(0, 0, 4, 0)
        DockPanel.SetDock(swatch, WpfDock.Left)

        tb = TextBox()
        tb.Text       = hex_val or ""
        tb.Background = br(C_INPUT_BG)
        tb.Foreground = br(C_INPUT_FG)
        tb.BorderBrush = br(C_GREEN)
        tb.BorderThickness = Thickness(1)
        tb.FontSize   = 11
        tb.Padding    = Thickness(4, 2, 4, 2)
        tb.Height     = 24
        tb.VerticalContentAlignment = VerticalAlignment.Center

        def make_handler(sw):
            def on_change(sender, e):
                sw.Background = _hex_to_brush(sender.Text.strip())
            return on_change
        tb.TextChanged += make_handler(swatch)

        row.Children.Add(lbl)
        row.Children.Add(swatch)
        row.Children.Add(tb)
        parent.Children.Add(row)
        return tb, swatch

    def _make_btn(self, text, handler, green=False):
        btn = Button()
        btn.Content    = text
        btn.Background = br(C_GREEN if green else Color.FromRgb(64, 69, 83))
        btn.Foreground = br(C_FG)
        btn.BorderThickness = Thickness(0)
        btn.Padding    = Thickness(14, 6, 14, 6)
        btn.FontSize   = 12
        btn.Margin     = Thickness(4, 0, 0, 0)
        btn.Cursor     = Cursors.Hand
        btn.Click     += handler
        return btn

    def _read_combo(self, cb):
        idx  = cb.SelectedIndex
        pats = self._fill_pats
        if idx <= 0 or idx >= len(pats): return None
        pname, pid = pats[idx]
        if pid is None: return None
        return str(pid.Value if hasattr(pid, "Value") else pid.IntegerValue)

    def _on_ok(self, sender, e):
        self.accepted = True
        fg_ck, fg_pk, fg_vk, bg_ck, bg_pk, bg_vk = self._keys
        hex_fg = self._fg_color_tb.Text.strip()
        hex_bg = self._bg_color_tb.Text.strip()
        self._settings[fg_ck] = hex_fg if hex_fg else None
        self._settings[fg_pk] = self._read_combo(self._fg_combo)
        self._settings[fg_vk] = bool(self._fg_vis.IsChecked) if self._fg_vis.IsChecked is not None else True
        self._settings[bg_ck] = hex_bg if hex_bg else None
        self._settings[bg_pk] = self._read_combo(self._bg_combo)
        self._settings[bg_vk] = bool(self._bg_vis.IsChecked) if self._bg_vis.IsChecked is not None else True
        self.Close()

    def _on_clear(self, sender, e):
        self.accepted = True
        fg_ck, fg_pk, fg_vk, bg_ck, bg_pk, bg_vk = self._keys
        for k in self._keys:
            self._settings[k] = None
        self._settings[fg_vk] = True
        self._settings[bg_vk] = True
        self.Close()

    def _on_cancel(self, sender, e):
        self.Close()

    def get_settings(self):
        return self._settings


# ── HEADER BUILDERS ───────────────────────────────────────────────────────────

def build_group_header(group_grid, assign_mode=False):
    group_grid.Children.Clear()
    group_grid.ColumnDefinitions.Clear()
    cols   = _cols_for(assign_mode)
    groups = _groups_for(assign_mode)
    saved  = get_col_widths()
    for i, (key, label, width, _) in enumerate(cols):
        cd = ColumnDefinition()
        # assign column keeps its fixed width; data columns use saved widths
        if key == "assign":
            cd.Width = GridLength(width)
        else:
            col_idx = i - (1 if assign_mode else 0)
            w = saved[col_idx] if 0 <= col_idx < len(saved) else width
            cd.Width = GridLength(w)
        group_grid.ColumnDefinitions.Add(cd)
    for group_label, start_col, span in groups:
        if not group_label:
            continue
        b = Border()
        b.Background     = br(C_HDR)
        b.BorderBrush    = br(C_BORDER)
        b.BorderThickness = Thickness(0, 0, 1, 0)
        b.Margin         = Thickness(1, 0, 1, 0)
        tb = TextBlock()
        tb.Text      = group_label
        tb.Foreground = br(C_FG)
        tb.FontSize  = 10
        tb.FontWeight = FontWeights.SemiBold
        tb.HorizontalAlignment = HorizontalAlignment.Center
        tb.VerticalAlignment   = VerticalAlignment.Center
        b.Child = tb
        Grid.SetColumn(b, start_col)
        Grid.SetColumnSpan(b, span)
        group_grid.Children.Add(b)


def build_header(header_grid, assign_mode=False, on_widths_changed=None):
    """Build the column header row with draggable GridSplitters.
    on_widths_changed(widths_list) is called after a drag completes."""
    from System.Windows.Controls import GridSplitter, GridResizeDirection, GridResizeBehavior
    from System.Windows.Input import Cursors as _Cur

    header_grid.Children.Clear()
    header_grid.ColumnDefinitions.Clear()
    cols  = _cols_for(assign_mode)
    saved = get_col_widths()

    # Build ColumnDefinitions, saving refs for the splitter drag callback
    col_defs = []
    for i, (key, label, width, _) in enumerate(cols):
        cd = ColumnDefinition()
        cd.MinWidth = 20
        if key == "assign":
            cd.Width = GridLength(width)
        else:
            col_idx = i - (1 if assign_mode else 0)
            w = saved[col_idx] if 0 <= col_idx < len(saved) else width
            cd.Width = GridLength(w)
        header_grid.ColumnDefinitions.Add(cd)
        col_defs.append(cd)

    # Add a star-width filler at the end so the last splitter always has
    # a column to push into (without it the Halftone splitter can't resize).
    filler_cd = ColumnDefinition()
    filler_cd.Width = GridLength(1, GridUnitType.Star)
    filler_cd.MinWidth = 0
    header_grid.ColumnDefinitions.Add(filler_cd)
    filler_col_idx = len(cols)   # index of the filler column

    # Label cells
    for i, (key, label, width, _) in enumerate(cols):
        tb = TextBlock()
        tb.Text      = label
        tb.Foreground = br(C_FG)
        tb.FontSize  = 10
        tb.FontWeight = FontWeights.SemiBold
        tb.VerticalAlignment   = VerticalAlignment.Center
        tb.HorizontalAlignment = HorizontalAlignment.Center
        tb.Margin    = Thickness(2, 0, 2, 0)
        Grid.SetColumn(tb, i)
        header_grid.Children.Add(tb)

    # GridSplitters — one per column boundary, including the last data col / filler boundary
    def _on_drag_complete(s, e):
        """Read current widths (exclude filler) and persist + push to all live rows."""
        new_widths = []
        offset = 1 if assign_mode else 0
        for cd in col_defs[offset:]:   # col_defs excludes the filler
            new_widths.append(int(cd.ActualWidth) if cd.ActualWidth > 0
                              else int(cd.Width.Value))
        set_col_widths(new_widths)
        update_all_row_widths(new_widths)
        if on_widths_changed:
            on_widths_changed(new_widths)

    # Place splitters at columns 1..len(cols) inclusive (last one sits between
    # the Halftone column and the filler, so it can always resize freely).
    for i in range(1, filler_col_idx + 1):
        spl = GridSplitter()
        spl.Width = 4
        spl.HorizontalAlignment = HorizontalAlignment.Left
        spl.VerticalAlignment   = VerticalAlignment.Stretch
        spl.Background          = br(C_BORDER)
        spl.Cursor              = _Cur.SizeWE
        spl.ResizeDirection     = GridResizeDirection.Columns
        spl.ResizeBehavior      = GridResizeBehavior.PreviousAndNext
        spl.ShowsPreview        = False
        spl.DragCompleted      += _on_drag_complete
        Grid.SetColumn(spl, i)
        header_grid.Children.Add(spl)

# ── DISPLAY ROW ───────────────────────────────────────────────────────────────

def _override_btn_label(settings, color_keys, pat_keys):
    """Return 'Override...' with colour dot if any override is set, else 'Override...'"""
    has_override = any(settings.get(k) for k in color_keys + pat_keys)
    return "Override..." if True else "Override..."


def build_display_row(filter_name, settings, row_index,
                      line_patterns, fill_patterns, owner_window,
                      on_settings_changed, on_sidebar_open=None,
                      assign_mode=False, assigned=False,
                      on_assign_changed=None, skip_preview=False):
    """
    Build one display row matching Revit VG layout.
    Override... buttons open popup dialogs.
    on_settings_changed(new_settings) is called when user edits via dialog.

    assign_mode prepends an "Add" checkbox column. assigned sets its initial
    state, and on_assign_changed(filter_name, is_assigned) fires on toggle.
    """
    g = Grid()
    g.Height     = 28
    g.Background = br(C_PANEL if row_index % 2 == 0 else C_PANEL2)
    g.Tag        = filter_name
    g.Cursor     = Cursors.Arrow

    cols  = _cols_for(assign_mode)
    saved = get_col_widths()
    row_col_defs = []
    for i, (key, label, width, _) in enumerate(cols):
        cd = ColumnDefinition()
        cd.MinWidth = 20
        if key == "assign":
            cd.Width = GridLength(width)
        else:
            col_idx = i - (1 if assign_mode else 0)
            w = saved[col_idx] if 0 <= col_idx < len(saved) else width
            cd.Width = GridLength(w)
        g.ColumnDefinitions.Add(cd)
        row_col_defs.append(cd)
    register_row_col_defs(row_col_defs)

    # Mutable settings wrapper so closures see updates
    state = [dict(settings)]
    # Collects rebuild callables when skip_preview=True, for deferred rendering
    deferred_previews = []

    def refresh_row():
        """Rebuild the row background colours after settings change."""
        pass  # handled by on_settings_changed rebuilding the row

    for i, (key, label, width, ctype) in enumerate(cols):

        if ctype == "assign":
            cb = CheckBox()
            cb.IsChecked = bool(assigned)
            cb.VerticalAlignment   = VerticalAlignment.Center
            cb.HorizontalAlignment = HorizontalAlignment.Center
            cb.Margin    = Thickness(4)
            cb.Foreground = br(C_FG)

            def make_assign_handler(checkbox):
                def on_check(sender, e):
                    if on_assign_changed:
                        on_assign_changed(
                            filter_name, bool(checkbox.IsChecked))
                return on_check
            cb.Checked   += make_assign_handler(cb)
            cb.Unchecked += make_assign_handler(cb)

            Grid.SetColumn(cb, i)
            g.Children.Add(cb)

        elif ctype == "label":
            tb = TextBlock()
            tb.Text      = filter_name
            tb.Foreground = br(C_FG)
            tb.FontSize  = 11
            tb.VerticalAlignment = VerticalAlignment.Center
            tb.TextTrimming      = TextTrimming.CharacterEllipsis
            tb.Margin    = Thickness(6, 0, 4, 0)
            Grid.SetColumn(tb, i)
            g.Children.Add(tb)

        elif ctype == "check":
            is_enabled = key == "enabled"
            is_visible = key == "visible"

            cb = CheckBox()
            val = state[0].get(key)
            cb.IsChecked = bool(val) if val is not None else True
            cb.VerticalAlignment   = VerticalAlignment.Center
            cb.HorizontalAlignment = HorizontalAlignment.Center
            cb.Margin  = Thickness(4)
            cb.Foreground = br(C_FG)

            def make_check_handler(k, checkbox):
                def on_check(sender, e):
                    state[0][k] = bool(checkbox.IsChecked) if checkbox.IsChecked is not None else True
                    on_settings_changed(state[0])
                return on_check
            cb.Checked   += make_check_handler(key, cb)
            cb.Unchecked += make_check_handler(key, cb)

            Grid.SetColumn(cb, i)
            g.Children.Add(cb)

        elif ctype == "halftone":
            cb = CheckBox()
            val = state[0].get("halftone")
            cb.IsChecked = bool(val) if val is not None else False
            cb.VerticalAlignment   = VerticalAlignment.Center
            cb.HorizontalAlignment = HorizontalAlignment.Center
            cb.Margin  = Thickness(4)
            cb.Foreground = br(C_FG)

            def make_ht_handler(checkbox):
                def on_check(sender, e):
                    state[0]["halftone"] = bool(checkbox.IsChecked) if checkbox.IsChecked is not None else False
                    on_settings_changed(state[0])
                return on_check
            cb.Checked   += make_ht_handler(cb)
            cb.Unchecked += make_ht_handler(cb)

            Grid.SetColumn(cb, i)
            g.Children.Add(cb)

        elif ctype == "transp":
            # Small editable textbox for transparency %
            tb = TextBox()
            val = state[0].get("transparency")
            tb.Text       = str(val) if val is not None else "0"
            tb.Background = br(C_INPUT_BG)
            tb.Foreground = br(C_INPUT_FG)
            tb.BorderBrush = br(C_GREEN)
            tb.BorderThickness = Thickness(1)
            tb.FontSize   = 10
            tb.Padding    = Thickness(4, 2, 4, 2)
            tb.Height     = 22
            tb.Margin     = Thickness(4)
            tb.VerticalContentAlignment = VerticalAlignment.Center
            tb.HorizontalContentAlignment = HorizontalAlignment.Center

            def make_transp_handler():
                def on_change(sender, e):
                    try:
                        state[0]["transparency"] = int(sender.Text.strip())
                        on_settings_changed(state[0])
                    except Exception:
                        pass
                return on_change
            tb.TextChanged += make_transp_handler()

            Grid.SetColumn(tb, i)
            g.Children.Add(tb)

        elif ctype == "lines_btn":
            is_proj = (key == "proj_lines")
            if is_proj:
                c_key = "proj_line_color"
                w_key = "proj_line_weight"
                p_key = "proj_line_pattern_id"
                dlg_title = "Projection Line Graphics"
            else:
                c_key = "cut_line_color"
                w_key = "cut_line_weight"
                p_key = "cut_line_pattern_id"
                dlg_title = "Cut Line Graphics"

            btn, rb = _make_override_btn(
                state, c_key, w_key, p_key,
                dlg_title, line_patterns, owner_window,
                on_settings_changed, on_sidebar_open=on_sidebar_open,
                skip_preview=skip_preview)
            if skip_preview:
                deferred_previews.append(rb)
            Grid.SetColumn(btn, i)
            g.Children.Add(btn)

        elif ctype == "pats_btn":
            is_proj = (key == "proj_patterns")
            if is_proj:
                fgc = "surf_fg_color"; fgp = "surf_fg_pat"; fgv = "surf_fg_visible"
                bgc = "surf_bg_color"; bgp = "surf_bg_pat"; bgv = "surf_bg_visible"
                dlg_title = "Projection/Surface Fill Pattern Graphics"
            else:
                fgc = "cut_fg_color";  fgp = "cut_fg_pat";  fgv = "cut_fg_visible"
                bgc = "cut_bg_color";  bgp = "cut_bg_pat";  bgv = "cut_bg_visible"
                dlg_title = "Cut Fill Pattern Graphics"

            btn, rb = _make_pat_btn(
                state, fgc, fgp, fgv, bgc, bgp, bgv,
                dlg_title, fill_patterns, owner_window,
                on_settings_changed, on_sidebar_open=on_sidebar_open,
                skip_preview=skip_preview)
            if skip_preview:
                deferred_previews.append(rb)
            Grid.SetColumn(btn, i)
            g.Children.Add(btn)

    # Store deferred rebuild list in module dict keyed by grid id
    if deferred_previews:
        _deferred_preview_store[id(g)] = deferred_previews
    return g


def _make_line_preview(color_hex, pattern_id_str, width=100, height=16):
    """Render a line preview using pattern_preview module."""
    try:
        from pyfilter_pattern_preview import make_line_preview, find_line_pattern_element
        from pyrevit import revit
        lpe = find_line_pattern_element(revit.doc, pattern_id_str)
        _preview_log("line preview: color={} pat_id={} lpe={}".format(
            color_hex, pattern_id_str, lpe))
        return make_line_preview(lpe, color_hex, width, height)
    except Exception as ex:
        _preview_log("line preview FAILED: {}".format(ex), "ERR")
        from System.Windows.Controls import Canvas
        from System.Windows.Shapes import Line
        canvas = Canvas()
        canvas.Width  = width
        canvas.Height = height
        canvas.Background = SolidColorBrush(Color.FromArgb(0, 0, 0, 0))
        line = Line()
        line.X1 = 4
        line.Y1 = height / 2.0
        line.X2 = width - 4
        line.Y2 = height / 2.0
        line.Stroke          = br(C_DIM)
        line.StrokeThickness = 1.5
        canvas.Children.Add(line)
        return canvas


def _make_pat_preview(fg_color, fg_pat_id, bg_color, bg_pat_id,
                      width=100, height=16):
    """Render a fill pattern preview using pattern_preview module."""
    try:
        from pyfilter_pattern_preview import make_fill_preview, find_fill_pattern_element
        from pyrevit import revit
        fgpe = find_fill_pattern_element(revit.doc, fg_pat_id)
        bgpe = find_fill_pattern_element(revit.doc, bg_pat_id)
        _preview_log("pat preview: fg_id={} fgpe={} bg_id={} bgpe={} fg_col={} bg_col={}".format(
            fg_pat_id, fgpe, bg_pat_id, bgpe, fg_color, bg_color))
        return make_fill_preview(fgpe, fg_color, bgpe, bg_color, width, height)
    except Exception as ex:
        _preview_log("pat preview FAILED: {}".format(ex), "ERR")
        from System.Windows.Controls import Canvas
        canvas = Canvas()
        canvas.Width  = width
        canvas.Height = height
        canvas.Background = SolidColorBrush(Color.FromArgb(0, 0, 0, 0))
        return canvas


def _preview_log(msg, level="INFO"):
    """Preview diagnostic messages — disabled. print() triggers pyRevit's
    console popup, so this is now a no-op. Re-enable only for debugging."""
    pass


def _make_override_btn(state, c_key, w_key, p_key,
                       section_title, line_patterns, owner,
                       on_settings_changed, on_sidebar_open=None,
                       skip_preview=False):
    """
    Build a line override cell button.
    Shows a visual line preview. On click, calls on_sidebar_open
    to populate the sidebar panel rather than opening a popup.
    """
    from System.Windows.Controls import Canvas

    panel = StackPanel()
    panel.Orientation         = _horizontal()
    panel.HorizontalAlignment = HorizontalAlignment.Center
    panel.VerticalAlignment   = VerticalAlignment.Center

    preview_holder = Canvas()
    preview_holder.Width  = 100
    preview_holder.Height = 16

    def rebuild_preview():
        preview_holder.Children.Clear()
        color   = state[0].get(c_key)
        pat_id  = state[0].get(p_key)
        preview = _make_line_preview(color, pat_id)
        Canvas.SetLeft(preview, 0)
        Canvas.SetTop(preview, 0)
        preview_holder.Children.Add(preview)

    if not skip_preview:
        rebuild_preview()
    panel.Children.Add(preview_holder)

    btn = Button()
    btn.Content             = panel
    btn.Background          = br(C_PANEL)
    btn.BorderThickness     = Thickness(0)
    btn.Margin              = Thickness(1)
    btn.Cursor              = Cursors.Hand
    btn.HorizontalContentAlignment = HorizontalAlignment.Center
    btn.VerticalContentAlignment   = VerticalAlignment.Center

    def on_click(sender, e):
        if on_sidebar_open:
            on_sidebar_open(
                section_title, "lines",
                c_key, w_key, p_key, None, None, None, None, None, None,
                state, line_patterns, None,
                lambda: rebuild_preview(),
                on_settings_changed)

    btn.Click += on_click
    return btn, rebuild_preview


def _make_pat_btn(state, fgc, fgp, fgv, bgc, bgp, bgv,
                  section_title, fill_patterns, owner,
                  on_settings_changed, on_sidebar_open=None,
                  skip_preview=False):
    """
    Build a pattern override cell button.
    Shows a visual pattern preview. On click, calls on_sidebar_open
    to populate the sidebar panel.
    """
    from System.Windows.Controls import Canvas

    panel = StackPanel()
    panel.Orientation         = _horizontal()
    panel.HorizontalAlignment = HorizontalAlignment.Center
    panel.VerticalAlignment   = VerticalAlignment.Center

    preview_holder = Canvas()
    preview_holder.Width  = 100
    preview_holder.Height = 16

    def rebuild_preview():
        preview_holder.Children.Clear()
        fg_color = state[0].get(fgc)
        bg_color = state[0].get(bgc)
        fg_pat   = state[0].get(fgp)
        bg_pat   = state[0].get(bgp)
        _preview_log("rebuild_preview: fg_pat={} bg_pat={}".format(fg_pat, bg_pat))
        preview  = _make_pat_preview(fg_color, fg_pat, bg_color, bg_pat)
        Canvas.SetLeft(preview, 0)
        Canvas.SetTop(preview, 0)
        preview_holder.Children.Add(preview)

    if not skip_preview:
        rebuild_preview()
    panel.Children.Add(preview_holder)

    btn = Button()
    btn.Content             = panel
    btn.Background          = br(C_PANEL)
    btn.BorderThickness     = Thickness(0)
    btn.Margin              = Thickness(1)
    btn.Cursor              = Cursors.Hand
    btn.HorizontalContentAlignment = HorizontalAlignment.Center
    btn.VerticalContentAlignment   = VerticalAlignment.Center

    def on_click(sender, e):
        if on_sidebar_open:
            on_sidebar_open(
                section_title, "patterns",
                None, None, None, fgc, fgp, fgv, bgc, bgp, bgv,
                state, None, fill_patterns,
                lambda: rebuild_preview(),
                on_settings_changed)

    btn.Click += on_click
    return btn, rebuild_preview


def build_sidebar_editor(options_panel, section_title, editor_type,
                         c_key, w_key, p_key,
                         fgc, fgp, fgv, bgc, bgp, bgv,
                         state, line_patterns, fill_patterns,
                         on_preview_rebuild, on_settings_changed,
                         on_close=None):
    """
    Populate the sidebar OptionsPanel with an inline editor
    for either a lines cell or a patterns cell.
    Called instead of opening a popup dialog.
    """
    options_panel.Children.Clear()

    # Section heading
    heading = TextBlock()
    heading.Text       = section_title
    heading.Foreground = br(C_FG)
    heading.FontSize   = 12
    heading.FontWeight = FontWeights.SemiBold
    heading.Margin     = Thickness(0, 0, 0, 2)
    options_panel.Children.Add(heading)

    hint = TextBlock()
    hint.Text      = "Click away to close"
    hint.Foreground = br(C_DIM)
    hint.FontSize  = 9
    hint.Margin    = Thickness(0, 0, 0, 10)
    options_panel.Children.Add(hint)

    def add_label(text):
        tb = TextBlock()
        tb.Text      = text
        tb.Foreground = br(C_DIM)
        tb.FontSize  = 10
        tb.Margin    = Thickness(0, 6, 0, 2)
        options_panel.Children.Add(tb)

    def add_color_field(key, label):
        add_label(label)
        row = DockPanel()
        row.LastChildFill = True
        row.Margin        = Thickness(0, 0, 0, 4)
        from System.Windows.Controls import Dock as WpfDock

        hex_val = state[0].get(key) or ""
        swatch  = Border()
        swatch.Width           = 22
        swatch.Height          = 22
        swatch.CornerRadius    = CornerRadius(3)
        swatch.BorderThickness = Thickness(1)
        swatch.BorderBrush     = br(C_BORDER)
        swatch.Background      = _hex_to_brush(hex_val)
        swatch.VerticalAlignment = VerticalAlignment.Center
        swatch.Margin          = Thickness(0, 0, 6, 0)
        DockPanel.SetDock(swatch, WpfDock.Left)

        tb = TextBox()
        tb.Text       = hex_val
        tb.Background = br(C_INPUT_BG)
        tb.Foreground = br(C_INPUT_FG)
        tb.BorderBrush = br(C_GREEN)
        tb.BorderThickness = Thickness(1)
        tb.FontSize   = 11
        tb.Padding    = Thickness(6, 3, 6, 3)
        tb.Height     = 26
        tb.VerticalContentAlignment = VerticalAlignment.Center

        def make_color_handler(sw, k):
            def on_change(sender, e):
                sw.Background = _hex_to_brush(sender.Text.strip())
                state[0][k]   = sender.Text.strip() or None
                on_settings_changed(state[0])
                on_preview_rebuild()
            return on_change
        tb.TextChanged += make_color_handler(swatch, key)

        row.Children.Add(swatch)
        row.Children.Add(tb)
        options_panel.Children.Add(row)

    def add_weight_field(key, label):
        add_label(label)
        tb = TextBox()
        val = state[0].get(key)
        tb.Text       = str(val) if val is not None else "-1"
        tb.Background = br(C_INPUT_BG)
        tb.Foreground = br(C_INPUT_FG)
        tb.BorderBrush = br(C_GREEN)
        tb.BorderThickness = Thickness(1)
        tb.FontSize   = 11
        tb.Padding    = Thickness(6, 3, 6, 3)
        tb.Height     = 26
        tb.Margin     = Thickness(0, 0, 0, 4)
        tb.VerticalContentAlignment = VerticalAlignment.Center

        def make_weight_handler(k):
            def on_change(sender, e):
                try:
                    state[0][k] = int(sender.Text.strip())
                    on_settings_changed(state[0])
                except Exception:
                    pass
            return on_change
        tb.TextChanged += make_weight_handler(key)
        options_panel.Children.Add(tb)

    def add_combo_field(key, label, pats):
        add_label(label)
        cb = ComboBox()
        cb.Background  = br(C_INPUT_BG)
        cb.Foreground  = br(C_INPUT_FG)
        cb.BorderBrush = br(C_GREEN)
        cb.FontSize    = 11
        cb.Height      = 26
        cb.Margin      = Thickness(0, 0, 0, 4)
        cb.ItemContainerStyle = get_combo_style()
        stored_id = state[0].get(key)
        sel = 0
        for idx, (pname, pid) in enumerate(pats):
            item = ComboBoxItem()
            item.Content = pname
            cb.Items.Add(item)
            if stored_id and pid is not None:
                pid_str = str(pid.Value if hasattr(pid, "Value") else pid.IntegerValue)
                if pid_str == str(stored_id):
                    sel = idx
        cb.SelectedIndex = sel

        def make_combo_handler(k, p):
            def on_change(sender, e):
                if sender.Tag == "loading":
                    return
                idx2  = sender.SelectedIndex
                if idx2 <= 0 or idx2 >= len(p):
                    state[0][k] = None
                else:
                    pname2, pid2 = p[idx2]
                    state[0][k] = (
                        str(pid2.Value if hasattr(pid2, "Value") else pid2.IntegerValue)
                        if pid2 else None)
                on_settings_changed(state[0])
                on_preview_rebuild()
            return on_change
        cb.Tag = "loading"
        cb.SelectionChanged += make_combo_handler(key, pats)
        cb.SelectedIndex = sel
        cb.Tag = None
        options_panel.Children.Add(cb)

    def add_check_field(key, label):
        row = StackPanel()
        row.Orientation = _horizontal()
        row.Margin      = Thickness(0, 6, 0, 4)
        cb = CheckBox()
        val = state[0].get(key)
        cb.IsChecked  = bool(val) if val is not None else True
        cb.Foreground = br(C_FG)
        cb.VerticalAlignment = VerticalAlignment.Center
        lbl = TextBlock()
        lbl.Text      = label
        lbl.Foreground = br(C_DIM)
        lbl.FontSize  = 11
        lbl.Margin    = Thickness(6, 0, 0, 0)
        lbl.VerticalAlignment = VerticalAlignment.Center

        def make_check_handler(k):
            def on_check(sender, e):
                state[0][k] = bool(sender.IsChecked) if sender.IsChecked is not None else True
                on_settings_changed(state[0])
            return on_check
        cb.Checked   += make_check_handler(key)
        cb.Unchecked += make_check_handler(key)

        row.Children.Add(cb)
        row.Children.Add(lbl)
        options_panel.Children.Add(row)

    def add_clear_btn(keys_to_clear):
        btn = Button()
        btn.Content    = "Clear Overrides"
        btn.Background = br(Color.FromRgb(64, 69, 83))
        btn.Foreground = br(C_FG)
        btn.BorderThickness = Thickness(0)
        btn.Padding    = Thickness(10, 6, 10, 6)
        btn.FontSize   = 11
        btn.Margin     = Thickness(0, 12, 0, 0)
        btn.Cursor     = Cursors.Hand
        btn.HorizontalAlignment = HorizontalAlignment.Stretch

        def on_clear(sender, e):
            for k in keys_to_clear:
                state[0][k] = None
            # Rebuild the sidebar with cleared values
            build_sidebar_editor(
                options_panel, section_title, editor_type,
                c_key, w_key, p_key, fgc, fgp, fgv, bgc, bgp, bgv,
                state, line_patterns, fill_patterns,
                on_preview_rebuild, on_settings_changed)
            on_settings_changed(state[0])
            on_preview_rebuild()

        btn.Click += on_clear
        options_panel.Children.Add(btn)

    if editor_type == "lines":
        add_color_field(c_key,  "Colour")
        add_weight_field(w_key, "Weight (-1 = no override)")
        add_combo_field(p_key,  "Pattern", line_patterns)
        add_clear_btn([c_key, w_key, p_key])
    else:
        # Patterns editor
        fg_section = TextBlock()
        fg_section.Text      = "Foreground"
        fg_section.Foreground = br(C_GREEN)
        fg_section.FontSize  = 11
        fg_section.FontWeight = FontWeights.SemiBold
        fg_section.Margin    = Thickness(0, 4, 0, 4)
        options_panel.Children.Add(fg_section)
        add_check_field(fgv,  "Visible")
        add_combo_field(fgp,  "Pattern", fill_patterns)
        add_color_field(fgc,  "Colour")

        sep = Border()
        sep.Height     = 1
        sep.Background = br(C_BORDER)
        sep.Margin     = Thickness(0, 8, 0, 8)
        options_panel.Children.Add(sep)

        bg_section = TextBlock()
        bg_section.Text      = "Background"
        bg_section.Foreground = br(C_GREEN)
        bg_section.FontSize  = 11
        bg_section.FontWeight = FontWeights.SemiBold
        bg_section.Margin    = Thickness(0, 4, 0, 4)
        options_panel.Children.Add(bg_section)
        add_check_field(bgv,  "Visible")
        add_combo_field(bgp,  "Pattern", fill_patterns)
        add_color_field(bgc,  "Colour")
        add_clear_btn([fgc, fgp, fgv, bgc, bgp, bgv])


# ── SAVE ──────────────────────────────────────────────────────────────────────

def merge_rows_into_template_data(template_data, new_rows, overwrite_names):
    """
    Merge new_rows (list of {name, settings, definition}) into template_data
    (the dict loaded from a .json file). overwrite_names is a set of filter
    names that should overwrite the existing entry; others are appended only
    if not already present. Returns (merged_dict, added_count, updated_count,
    skipped_count).
    """
    existing = list(template_data.get("filters", []))
    by_name  = {f.get("name"): i for i, f in enumerate(existing)}

    added = updated = skipped = 0
    for row in new_rows:
        fname = row.get("name")
        if not fname:
            continue
        entry = {
            "name":       fname,
            "definition": row.get("definition", {}),
            "settings":   row.get("settings", {}),
        }
        if fname in by_name:
            if fname in overwrite_names:
                existing[by_name[fname]] = entry
                updated += 1
            else:
                skipped += 1
        else:
            existing.append(entry)
            by_name[fname] = len(existing) - 1
            added += 1

    out = dict(template_data)
    out["filters"] = existing
    return out, added, updated, skipped


def write_template_data(templates_folder, name, template_data):
    """Write a template dict directly. Used when merging rather than rebuilding
    from live filter elements."""
    out = dict(template_data)
    out["name"] = name
    if "created" not in out:
        out["created"] = datetime.datetime.now().strftime("%Y-%m-%d")
    path = os.path.join(templates_folder, name + ".json")
    try:
        with io.open(path, "w", encoding="utf-8") as fh:
            json.dump(out, fh, indent=2, ensure_ascii=False)
        return True, path
    except Exception as ex:
        return False, str(ex)


def save_template(templates_folder, name, filter_rows):
    filter_map   = {f.Name: f for f in get_all_filters()}
    filters_data = []
    for row in filter_rows:
        fname = row["name"]
        f     = filter_map.get(fname)
        defn  = row.get("definition") or (serialise_filter_def(f) if f else {})
        filters_data.append({
            "name":       fname,
            "definition": defn,
            "settings":   row.get("settings", {}),
        })
    tpl = {
        "name":    name,
        "created": datetime.datetime.now().strftime("%Y-%m-%d"),
        "filters": filters_data,
    }
    path = os.path.join(templates_folder, name + ".json")
    try:
        with io.open(path, "w", encoding="utf-8") as fh:
            json.dump(tpl, fh, indent=2, ensure_ascii=False)
        return True, path
    except Exception as ex:
        return False, str(ex)

# ── MISC HELPERS ──────────────────────────────────────────────────────────────

def _horizontal():
    from System.Windows.Controls import Orientation
    return Orientation.Horizontal

def _star():
    return _gl_star()

def _gl_star():
    from System.Windows import GridUnitType
    return GridUnitType.Star

def _resize_none():
    from System.Windows import ResizeMode
    return ResizeMode.NoResize
