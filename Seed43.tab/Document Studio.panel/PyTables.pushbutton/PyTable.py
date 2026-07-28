# -*- coding: utf-8 -*-
# PyTable.py
"""
pyTable - Import Excel named ranges and Word section notes into
Revit as native views.

This is now the slim entry point + shared window plumbing only.
Excel-specific logic (xlsx parsing, Schedule/Legend/Drafting
creation, the Excel side of the card/row UI) lives in
tools/pytable_excel.py. Word-specific logic (docx parsing, the
Strict-layout algorithm, section groups, the Word side of the
card/row UI) lives in tools/pytable_word.py. Both are mixed into
PyTableWindow below via multiple inheritance, so every method still
just uses `self.` exactly as before -- only the file each one lives
in changed, not how they're called.
"""
from pyrevit import revit, DB
from pyrevit import forms
from pyrevit import script

import os
import sys
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
    GridLength, GridUnitType,
    Window, SizeToContent, WindowStartupLocation,
    WindowStyle, ResizeMode
)
from System import DateTime
from System.Windows.Controls import (
    StackPanel, Border, CheckBox, TextBlock, TextBox,
    ComboBox, Button, Orientation, ScrollViewer,
    Grid, ColumnDefinition
)
from System.Windows.Controls.Primitives import Popup, PlacementMode
from System.Windows.Shapes import Ellipse, Rectangle
from System.Windows.Media import Color, Brushes
from System.Windows.Media.Effects import DropShadowEffect

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

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'tools'))
from pytable_shared import (
    hb, Row, VIEW_TYPES, WORD_VIEW_TYPES, SHEET_SIZES, SRC_COLOURS,
    STATUS_COLOURS, _alert, _confirm, _find_seed43_version,
    _get_pytable_param, _doc_base_dir, _to_relative, _to_absolute,
    save_pytable_state, load_pytable_state, _run_export_script,
    PYTABLE_PARAM_GUID, PYTABLE_PARAM_NAME, PYTABLE_PARAM_FILE,
    format_applied_at,
)
from pytable_excel import (
    TableRow, get_named_ranges_from_workbook, apply_row, _hash_range,
    ExcelCardMixin,
)
from pytable_word import (
    read_word_sections, get_word_headings, apply_notes_row,
    _hash_word_section, WordCardMixin,
)


class PyTableWindow(forms.WPFWindow, ExcelCardMixin, WordCardMixin):
    def __init__(self):
        forms.WPFWindow.__init__(self, 'PyTable.xaml')

        # -- Apply Seed43 theme (colours + sizing) ---------------------------------
        # This window never had this wired up at all before now - every colour
        # and size in PyTable.xaml was a hardcoded literal, disconnected from
        # the shared palette entirely. Must run AFTER LoadComponent (so our
        # injected brushes win over the XAML's own Setters) and BEFORE
        # anything below that builds dynamic UI or calls TryFindResource -
        # otherwise those lookups return None. See seed43-pyrevit-ui skill,
        # gotchas #3/#4, and pyTransmit.py/about.py for the same pattern.
        try:
            _pt_script_dir = os.path.dirname(os.path.abspath(__file__))
            from Snippets.seed43_theme import apply_seed43_palette, apply_seed43_dimensions
            apply_seed43_palette(self, _pt_script_dir)
            apply_seed43_dimensions(self, _pt_script_dir)
        except Exception as ex:
            logger.warning('Seed43 theme apply failed: {}'.format(ex))

        self.Closing += self._on_window_closing
        try:
            from Snippets._icons import make_icon as _mi
            self.CloseBtn.Content = _mi('close', size=14, color='#F4FAFF')
        except Exception as ex:
            logger.warning('Close icon load failed: {}'.format(ex))
        # _file_data: {path: {sheets, sheet_range_map, source_type, rows, card_panel}}
        self._file_data   = {}
        # _card_groups: {real_path: {group_border, views_panel, header
        # refs}} — Word fd entries sharing the same real_path render
        # as one card with a shared header, each fd as its own nested
        # "view block" inside it. Excel cards never use this — each
        # is always its own flat, standalone card.
        self._card_groups = {}
        self._active_file = None
        self._status_card  = None
        self._progress_bar = None
        self._lbl_pending  = None
        self._lbl_success  = None
        self._lbl_error    = None
        self._lbl_skipped  = None
        self._sync_panel   = None
        self._sync_label   = None
        self._setup_filter_toggles()
        self._update_file_combo()
        self._build_status_card()
        self._refresh_status_card()
        self._load_persisted_state()
        self._refresh_drop_zone()

    # ── File combo ──

    def _update_file_combo(self):
        """Keep _active_file pointed at something valid. The File:
        dropdown this used to also drive is gone, per-card controls
        (Add Row / Reload / Close) no longer need a globally 'active'
        file, but a few internal helpers (_add_row_for_card's
        temporary-redirect pattern) still read _active_file, so keep
        it sane here rather than ripping it out everywhere."""
        if self._active_file not in self._file_data:
            paths = list(self._file_data.keys())
            self._active_file = paths[0] if paths else None

    # ── Drop zone visibility ──

    def _refresh_drop_zone(self):
        if self._file_data:
            self.DropZonePanel.Visibility = Visibility.Collapsed
        else:
            self.DropZonePanel.Visibility = Visibility.Visible
        self._update_footer()

    def _update_footer(self):
        total = sum(len(fd['rows']) for fd in self._file_data.values())
        xl    = sum(1 for fd in self._file_data.values()
                    if fd.get('source_type') == 'xl')
        word  = len(self._file_data) - xl
        self.FooterCounts.Text = (
            'Total {} | Excel {} | Word {}'.format(total, xl, word)
            if self._file_data else 'No files loaded')

    def _set_status(self, text):
        pass  # status shown via dots and progress bar only

    # ── Card builder ──

    def _make_card(self, path):
        """Create a card Border for a file and add it to CardsPanel.
        Word cards render grouped (see _make_word_card in
        WordCardMixin) — multiple fd entries sharing the same file
        nest as separate 'views' inside one shared card. Excel cards
        stay flat, one fd per card, unchanged."""
        fd = self._file_data[path]
        is_word = fd.get('source_type') == 'word'
        if is_word:
            self._make_word_card(path)
            return

        self._make_excel_card(path)

    # ── Row builder ──

    def _short_error(self, message):
        """Map a full error message to a short inline label (max 5 words)."""
        m = message.lower()
        if 'no data'      in m or 'empty'     in m: return 'Empty range'
        if 'header'       in m:                      return 'No header row'
        if 'view type'    in m:                      return 'Invalid view type'
        if 'already exist' in m or 'name'     in m: return 'View name exists'
        if 'sheet'        in m:                      return 'Sheet not found'
        if 'ref'          in m:                      return 'Broken reference'
        if 'transaction'  in m or 'revit'     in m: return 'Revit error'
        return 'Failed'

    def _combo_style(self, combo, width):
        try:
            combo.Style = self.FindResource('ComboBoxStyle')
            combo.DropDownOpened += self._animate_combo_open
        except Exception as e:
            logger.warning('Failed to apply ComboBoxStyle: {}'.format(e))
        combo.Width              = width
        combo.Margin             = Thickness(0, 0, 4, 0)
        combo.VerticalAlignment  = VerticalAlignment.Center

    def _animate_combo_open(self, sender, e):
        """Play a slow expanding-open animation on the dropdown popup,
        every time it opens. Driven from the real DropDownOpened event
        rather than an XAML trigger bound across the Popup boundary,
        that binding style is known to be unreliable in WPF (Popup
        content isn't part of the normal visual tree), a real .NET
        event handler is the dependable way to do this."""
        try:
            sender.ApplyTemplate()
            scale = sender.Template.FindName('PopupScale', sender)
            if scale is None:
                return
            anim_mod = __import__(
                'System.Windows.Media.Animation',
                fromlist=['DoubleAnimation', 'QuadraticEase', 'EasingMode'])
            media_mod = __import__(
                'System.Windows.Media', fromlist=['ScaleTransform'])
            sys_mod = __import__('System', fromlist=['TimeSpan'])
            windows_mod = __import__('System.Windows', fromlist=['Duration'])

            anim = anim_mod.DoubleAnimation()
            anim.From = 0.0
            anim.To = 1.0
            anim.Duration = windows_mod.Duration(
                sys_mod.TimeSpan.FromSeconds(0.22))
            ease = anim_mod.QuadraticEase()
            ease.EasingMode = anim_mod.EasingMode.EaseOut
            anim.EasingFunction = ease
            scale.BeginAnimation(media_mod.ScaleTransform.ScaleYProperty, anim)
        except Exception:
            pass

    # ── Hand-built dropdown popups (MenuPopup, BatchPopup, and per-card
    #    batch popups). The toolbar ones (MenuPopup/BatchPopup) bind
    #    Popup.IsOpen directly to their anchor ToggleButton's own
    #    IsChecked in XAML - the same proven pattern the Source/Type
    #    filter dropdowns already use in this file - so there's no
    #    manual open/close/outside-click bookkeeping needed for those at
    #    all. Per-card popups (built fresh in code, anchored to a plain
    #    Button with no IsChecked to bind) still open/close imperatively,
    #    but StaysOpen="False" lets WPF's own native auto-dismiss handle
    #    the outside-click case for them too. ──

    def _build_styled_dialog(self, title, width=380):
        """Build a secondary dialog window (Section Groups, Word Text
        Size, etc.) that actually matches the app's theme instead of a
        plain Window with the OS's own white title bar and default
        control chrome. A brand-new Window has none of PyTable's own
        Resources - two things are needed, matching the main window's
        own setup in __init__ exactly:
          1. apply_seed43_palette/apply_seed43_dimensions injected into
             THIS window's own Resources (a Style's DynamicResource
             setters resolve against the element they're applied to's
             own resource chain, not wherever the Style itself was
             declared - so canonical styles pulled from self.FindResource
             below won't render correctly without this).
          2. WindowStyle=None + AllowsTransparency, with a hand-built
             title bar (drag via DragMove, a real CloseButtonStyle 'X'),
             since that's how the main window avoids the OS chrome too -
             just via WindowChrome there; plain AllowsTransparency here
             is simpler and entirely sufficient for a fixed-size dialog.
        Returns (window, content_panel) - add the dialog's own controls
        to content_panel, then set window.Content indirectly by just
        showing it (already wired) and call window.ShowDialog()."""
        w = Window()
        w.Title = title
        w.Width = width
        w.SizeToContent = SizeToContent.Height
        w.WindowStartupLocation = WindowStartupLocation.CenterOwner
        w.WindowStyle = getattr(WindowStyle, 'None')
        w.AllowsTransparency = True
        w.Background = None
        w.ResizeMode = ResizeMode.NoResize
        try:
            w.Owner = self
        except Exception:
            pass

        try:
            _pt_script_dir = os.path.dirname(os.path.abspath(__file__))
            from Snippets.seed43_theme import apply_seed43_palette, apply_seed43_dimensions
            apply_seed43_palette(w, _pt_script_dir)
            apply_seed43_dimensions(w, _pt_script_dir)
        except Exception as ex:
            logger.warning('Seed43 theme apply failed for dialog: {}'.format(ex))

        outer = Border()
        outer.Background   = w.TryFindResource('BrushWindowBg') or hb('#2B3340')
        outer.BorderBrush   = w.TryFindResource('BrushBorderDefault') or hb('#208A3C')
        outer.BorderThickness = Thickness(1)
        outer.CornerRadius = CornerRadius(10)
        shadow = DropShadowEffect()
        shadow.Color = Color.FromRgb(0, 0, 0)
        shadow.Opacity = 0.5
        shadow.ShadowDepth = 4
        shadow.BlurRadius = 14
        outer.Effect = shadow
        w.Content = outer

        root = StackPanel()
        outer.Child = root

        # Hand-built title bar — drag to move, real close button
        bar = Grid()
        bar.Margin = Thickness(14, 12, 10, 8)
        bar.MouseLeftButtonDown += lambda s, ev: w.DragMove()
        col1 = ColumnDefinition()
        col2 = ColumnDefinition()
        col2.Width = GridLength(0, GridUnitType.Auto)
        bar.ColumnDefinitions.Add(col1)
        bar.ColumnDefinitions.Add(col2)

        title_tb = TextBlock()
        title_tb.Text       = title
        title_tb.FontSize   = 14
        title_tb.FontWeight = FontWeights.Bold
        title_tb.Foreground = w.TryFindResource('BrushTextPrimary') or hb('#F4FAFF')
        title_tb.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(title_tb, 0)
        bar.Children.Add(title_tb)

        close_btn = Button()
        close_btn.Content = u'\u2715'
        try:
            close_btn.Style = self.FindResource('CloseButtonStyle')
        except Exception as ex:
            logger.warning('Failed to apply CloseButtonStyle: {}'.format(ex))
        close_btn.Click += lambda s, ev: w.Close()
        Grid.SetColumn(close_btn, 1)
        bar.Children.Add(close_btn)
        root.Children.Add(bar)

        content = StackPanel()
        content.Margin = Thickness(16, 0, 16, 16)
        root.Children.Add(content)

        return w, content

    def _build_dropdown_popup(self, anchor):
        """Build a Popup+Border+StackPanel matching MenuPopup/BatchPopup's
        XAML structure, for a dropdown anchored to a dynamically-built
        control (e.g. a per-card Batch button) that has no static XAML
        slot of its own. Adds the Popup as a sibling of the anchor in its
        parent panel so DynamicResource lookups resolve and the Popup
        positions itself correctly. Returns (popup, content_panel)."""
        popup = Popup()
        popup.PlacementTarget = anchor
        popup.Placement = PlacementMode.Bottom
        popup.AllowsTransparency = True
        popup.StaysOpen = False

        border = Border()
        try:
            border.Background  = self.FindResource('BrushCardBg')
            border.BorderBrush = self.FindResource('BrushBorderDefault')
        except Exception as e:
            logger.warning('Failed to resolve popup chrome brushes: {}'.format(e))
        border.BorderThickness = Thickness(1)
        try:
            border.CornerRadius = self.FindResource('CornerRadiusDropdownPopup')
        except Exception as e:
            logger.warning('Failed to resolve CornerRadiusDropdownPopup: {}'.format(e))
        border.MinWidth = 180
        shadow = DropShadowEffect()
        shadow.Color = Color.FromRgb(0, 0, 0)
        shadow.Opacity = 0.4
        shadow.ShadowDepth = 3
        shadow.BlurRadius = 8
        border.Effect = shadow

        panel = StackPanel()
        border.Child = panel
        popup.Child = border

        parent = anchor.Parent
        if parent is not None and hasattr(parent, 'Children'):
            parent.Children.Add(popup)
        return popup, panel

    def _make_menu_item(self, label, fn, popup):
        """Build a MenuItemStyle Button for a hand-built dropdown popup
        (not a native MenuItem/ContextMenu). StaysOpen="False" only
        auto-dismisses the popup on a click OUTSIDE it, so a click on
        one of its own items needs to close it explicitly here."""
        item = Button()
        item.Content = label
        try:
            item.Style = self.FindResource('MenuItemStyle')
        except Exception as e:
            logger.warning('Failed to apply MenuItemStyle: {}'.format(e))
        def _click(sender, ev):
            popup.IsOpen = False
            fn(sender, ev)
        item.Click += _click
        return item

    def _make_menu_separator(self):
        sep = Rectangle()
        sep.Height = 1
        sep.Margin = Thickness(6, 4, 6, 4)
        try:
            sep.Fill = self.FindResource('LocalBrushMenuBorder')
        except Exception as e:
            logger.warning('Failed to apply LocalBrushMenuBorder: {}'.format(e))
        return sep

    # ── Word row builder ──

    def _make_row_ui(self, row):
        if row.SourceType == 'word':
            return self._make_word_row_ui(row)

        sp = StackPanel()
        sp.Orientation = Orientation.Horizontal
        row._row_ui = sp
        sp.Height      = 32
        sp.Margin      = Thickness(0, 0, 0, 4)

        # Status dot — store ref on row so Apply can update it
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
        cb.Click             += self._cb_click
        row._enabled_cb      = cb
        sp.Children.Add(cb)

        # View Name
        vn = TextBox()
        vn.Text  = row.ViewName
        vn.Width = 120
        try:
            vn.Style = self.FindResource('TextBoxStyle')
        except Exception as e:
            logger.warning('Failed to apply TextBoxStyle: {}'.format(e))
        vn.Margin    = Thickness(0, 0, 4, 0)
        row._vn_textbox = vn
        vn.Tag       = row
        vn.LostFocus += self._vn_lost
        vn.TextChanged += self._excel_view_name_live_check
        sp.Children.Add(vn)

        # Sheet combo
        sc = ComboBox()
        self._combo_style(sc, 120)
        sc.Margin = Thickness(0, 0, 4, 0)
        for s in row._sheets:
            sc.Items.Add(s)
        if row.Sheet:
            sc.SelectedItem = row.Sheet
        elif sc.Items.Count > 0:
            sc.SelectedIndex = 0
            row.Sheet = sc.Items[0]
        sp.Children.Add(sc)

        # Named Range combo
        rc = ComboBox()
        self._combo_style(rc, 130)
        rc.Margin = Thickness(0, 0, 4, 0)
        for r in row.ranges_for():
            rc.Items.Add(r)
        if row.NamedRange:
            rc.SelectedItem = row.NamedRange
        elif rc.Items.Count > 0:
            rc.SelectedIndex = 0
            row.NamedRange = rc.Items[0]
        # Auto-fill ViewName on first row creation if still blank
        if not row.ViewName and row.NamedRange:
            row.ViewName = row.NamedRange
            if row._vn_textbox is not None:
                row._vn_textbox.Text = row.ViewName
        rc.Tag              = row
        rc.SelectionChanged += self._rc_changed
        sp.Children.Add(rc)

        sc.Tag              = (row, rc)
        sc.SelectionChanged += self._sc_changed

        # Synced — when this row was last successfully applied to
        # Revit, not the source file's own mtime (that only shows at
        # the card header level now).
        lm = TextBlock()
        lm.Text             = format_applied_at(row._applied_at)
        lm.Width            = 96
        lm.FontSize         = 10
        lm.Foreground       = hb('#F4FAFF')
        lm.Opacity          = 0.55
        lm.VerticalAlignment = VerticalAlignment.Center
        lm.Padding          = Thickness(4, 0, 4, 0)
        row._modified_label = lm
        sp.Children.Add(lm)

        # View Type combo
        vtc = ComboBox()
        self._combo_style(vtc, 130)
        vtc.Margin = Thickness(0, 0, 4, 0)
        for vt in VIEW_TYPES:
            vtc.Items.Add(vt)
        vtc.SelectedItem    = row.ViewType
        vtc.Tag             = row
        vtc.SelectionChanged += self._vt_changed
        sp.Children.Add(vtc)

        rb = self._make_sync_btn(row)
        row._refresh_btn    = rb
        sp.Children.Add(rb)

        # Delete button — always-red DeleteButtonStyle, not the transparent-
        # until-hover CloseButtonStyle (this removes the row, it isn't a
        # window/card close action). Sized smaller than the card-level
        # delete button (24 vs the style's own 30px token) since this one
        # sits inline in a dense row, not a card header.
        db = Button()
        db.Content          = u'\u2715'
        db.FontSize         = 10
        try:
            db.Style = self.FindResource('DeleteButtonStyle')
        except Exception as e:
            logger.warning('Failed to apply DeleteButtonStyle: {}'.format(e))
        db.FocusVisualStyle = None
        db.Width            = 24
        db.Height           = 24
        db.Cursor           = __import__('System.Windows.Input',
                                  fromlist=['Cursors']).Cursors.Hand
        db.VerticalAlignment = VerticalAlignment.Center
        db.Margin           = Thickness(4, 0, 0, 0)
        db.Tag              = row
        db.Click            += self._del_click
        sp.Children.Add(db)

        # Error pill — appears after ✕ as a small card
        err_pill = Border()
        err_pill.Background     = hb('#3B1515')
        err_pill.BorderBrush    = hb('#DC2626')
        err_pill.BorderThickness = Thickness(1)
        err_pill.CornerRadius   = CornerRadius(4)
        err_pill.Padding        = Thickness(8, 2, 8, 2)
        err_pill.Margin         = Thickness(6, 0, 0, 0)
        err_pill.VerticalAlignment = VerticalAlignment.Center
        err_pill.Visibility     = Visibility.Collapsed
        err_txt = TextBlock()
        err_txt.FontSize        = 10
        err_txt.Foreground      = hb('#F87171')
        err_txt.VerticalAlignment = VerticalAlignment.Center
        err_txt.Text            = ''
        err_pill.Child          = err_txt
        row._error_label        = err_pill
        row._error_text         = err_txt
        sp.Children.Add(err_pill)

        return sp

    # ── Row event handlers ──

    def _cb_click(self, sender, e):
        row = sender.Tag
        if row:
            row.Enabled = (sender.IsChecked == True)
            path = self._find_card_path_for_row(row)
            if path:
                self._update_tri_select_state(path)
            self._update_footer()
            self._save_persisted_state()

    def _find_card_path_for_row(self, row):
        for path, fd in self._file_data.items():
            if row in fd.get('rows', []):
                return path
        return None

    def _auto_check_row(self, row):
        """Tick a row's checkbox because the user just changed
        something that affects its output. Rows default unchecked so
        Apply doesn't silently pick up half-configured rows — but once
        something is actually edited, it's clearly meant to be
        included, so require unchecking to exclude it instead."""
        if row is None or row.Enabled:
            return
        row.Enabled = True
        if row._enabled_cb is not None:
            row._enabled_cb.IsChecked = True
        path = self._find_card_path_for_row(row)
        if path:
            self._update_tri_select_state(path)

    def _auto_check_card(self, path):
        """Same as _auto_check_row, but for a card-level change (view
        name, sheet size, columns, view type, layout mode) that
        affects every row in the card at once."""
        fd = self._file_data.get(path)
        if fd is None:
            return
        changed = False
        for row in fd.get('rows', []):
            if not row.Enabled:
                row.Enabled = True
                if row._enabled_cb is not None:
                    row._enabled_cb.IsChecked = True
                changed = True
        if changed:
            self._update_tri_select_state(path)

    def _update_tri_select_state(self, path):
        """Set the per-card header checkbox to checked/unchecked/
        indeterminate based on how many rows in that card are enabled."""
        fd = self._file_data.get(path)
        if fd is None:
            return
        cb = fd.get('select_all_cb')
        if cb is None:
            return
        rows = fd.get('rows', [])
        if not rows:
            cb.IsChecked = False
            return
        on = sum(1 for r in rows if r.Enabled)
        if on == 0:
            cb.IsChecked = False
        elif on == len(rows):
            cb.IsChecked = True
        else:
            cb.IsChecked = None

    def _card_select_all_click(self, sender, e):
        """Header tri-state checkbox: always resolves to a definite
        all-on or all-off for every row in this card, regardless of
        the indeterminate display state it was showing."""
        path = sender.Tag
        fd = self._file_data.get(path)
        if fd is None:
            return
        rows = fd.get('rows', [])
        if not rows:
            sender.IsChecked = False
            return
        new_state = not all(r.Enabled for r in rows)
        for r in rows:
            r.Enabled = new_state
            if r._enabled_cb is not None:
                r._enabled_cb.IsChecked = new_state
        self._update_tri_select_state(path)
        self._update_footer()
        self._save_persisted_state()

    def _revalidate_all_view_name_boxes(self):
        """Re-check every View Name box after any one of them commits
        a change — a rename can create or resolve a conflict for a
        completely different card/row, not just the one being edited."""
        for path, fd in self._file_data.items():
            if fd.get('source_type') == 'word':
                box = fd.get('view_name_box')
                if box is not None:
                    taken = self._view_name_taken(
                        fd.get('view_name', ''), exclude_word_path=path)
                    self._style_view_name_conflict(box, taken)
            for row in fd.get('rows', []):
                box = row._vn_textbox
                if box is not None:
                    taken = self._view_name_taken(row.ViewName, exclude_row=row)
                    self._style_view_name_conflict(box, taken)

    def _view_name_taken(self, name, exclude_word_path=None, exclude_row=None):
        """True if `name` is already used by another view pyTable
        would create — another Word card's View Name, another Excel
        row's View Name, or an existing Revit view this specific box
        doesn't already own (so re-typing your own current name isn't
        flagged as a conflict with yourself)."""
        name = (name or '').strip()
        if not name:
            return False
        for path, fd in self._file_data.items():
            if fd.get('source_type') == 'word':
                if path == exclude_word_path:
                    continue
                if (fd.get('view_name') or '').strip() == name:
                    return True
            for r in fd.get('rows', []):
                if r is exclude_row:
                    continue
                if fd.get('source_type') != 'word' and (r.ViewName or '').strip() == name:
                    return True
        owned = None
        if exclude_word_path is not None:
            owned = self._file_data.get(exclude_word_path, {}).get('_applied_view_name')
        elif exclude_row is not None:
            owned = getattr(exclude_row, '_applied_view_name', None)
        if name != owned:
            try:
                from pyrevit.revit import query
                for v in query.get_elements_by_class(DB.View, doc=doc):
                    try:
                        if v.IsValidObject and v.Name == name:
                            return True
                    except Exception:
                        continue
            except Exception:
                pass
        return False

    def _style_view_name_conflict(self, box, is_conflict):
        """Red border when the typed name collides with an existing view,
        normal border otherwise — swaps the whole Style rather than
        setting BorderBrush directly: TextBoxStyle's IsFocused/IsMouseOver
        triggers hardcode the normal border, which would silently mask
        a red border set from code the whole time the box has focus —
        exactly when a typed conflict most needs to be visible."""
        try:
            box.Style = self.FindResource(
                'TextBoxErrorStyle' if is_conflict else 'TextBoxStyle')
        except Exception as e:
            logger.warning('Failed to apply TextBox(Error)Style: {}'.format(e))

    def _remove_row_ui(self, row):
        """Remove a single row's data + UI element from whichever card
        owns it. Shared by the single-row X button and the per-card
        'Delete selected' batch action."""
        for path, fd in self._file_data.items():
            if row in fd['rows']:
                fd['rows'].remove(row)
                panel = fd['card_panel']
                to_remove = None
                for child in panel.Children:
                    if getattr(child, 'Tag', None) is row:
                        to_remove = child
                        break
                if to_remove is not None:
                    panel.Children.Remove(to_remove)
                break

    def _del_click(self, sender, e):
        row = sender.Tag
        name = row.ViewName or row.NamedRange or 'this row'
        if not _confirm(
                'Remove row: {}?'.format(name),
                title='Confirm Remove'):
            return
        path = self._find_card_path_for_row(row)
        self._remove_row_ui(row)
        self._update_footer()
        self._refresh_status_card()
        self._save_persisted_state()
        if path:
            self._maybe_run_strict_layout(path)

    # ── File loading ──

    def _load_files(self, paths):
        if not paths:
            return
        for path in paths:
            ext = os.path.splitext(path)[1].lower()
            if ext in ('.xlsx', '.xls', '.ods'):
                self._parse_excel(path)
            elif ext in ('.docx', '.doc', '.odt'):
                self._parse_word(path)
        self._update_file_combo()
        self._refresh_drop_zone()
        self._revalidate_all_view_name_boxes()

    def OnAddTables(self, sender, e):
        """+ Add Tables. For now this opens the same browse-and-load
        flow as before, files still get auto-populated the same way.
        The full DiRoots-style picker (choose file, then check which
        sheets/ranges become rows, set defaults, one Ok) is a bigger,
        separate piece and hasn't been built yet."""
        self.OnBrowse(sender, e)

    def _setup_filter_toggles(self):
        """Build the mini toggle switches inside the Source/Type accordion
        bodies - same bare-Border-track + Border-knob technique pyTransmit
        uses for its own on/off switches (no animation, knob position set
        directly - reliable in IronPython 2). All default ON, matching
        the equivalent checkboxes' previous default."""
        self._filter_state = {
            'excel': True, 'word': True, 'ods': True, 'odt': True,
            'schedule': True, 'legend': True, 'drafting': True,
        }
        self._filter_tracks = {
            'excel':    self.SrcToggleExcel,
            'word':     self.SrcToggleWord,
            'ods':      self.SrcToggleOds,
            'odt':      self.SrcToggleOdt,
            'schedule': self.TypeToggleSchedule,
            'legend':   self.TypeToggleLegend,
            'drafting': self.TypeToggleDrafting,
        }
        self._filter_knobs = {}
        source_keys = ('excel', 'word', 'ods', 'odt')
        for key, track in self._filter_tracks.items():
            self._build_mini_toggle(track, key)
            changed_fn = (self._apply_source_filter if key in source_keys
                          else self._apply_type_filter)
            track.MouseLeftButtonUp += (
                lambda s, ev, k=key, fn=changed_fn: self._toggle_filter(k, fn))

    def _build_mini_toggle(self, track, key):
        """Build one mini toggle switch's knob and set its initial
        position/colour to match self._filter_state[key]."""
        on = self._filter_state[key]
        on_brush = (self.TryFindResource('BrushPrimaryGreen') or
                    SolidColorBrush(Color.FromRgb(0x20, 0x8A, 0x3C)))
        off_brush = (self.TryFindResource('BrushToggleOffBg') or
                     SolidColorBrush(Color.FromRgb(0x5E, 0x5C, 0x64)))
        knob_brush = self.TryFindResource('BrushToggleKnob') or Brushes.White
        track_w = self.TryFindResource('WidthToggle') or 40.0
        knob_size = self.TryFindResource('SizeToggleKnob') or 16.0
        knob_margin = self.TryFindResource('MarginToggleKnob') or 2.0
        knob_radius = (self.TryFindResource('CornerRadiusToggleKnob') or
                        CornerRadius(knob_size / 2.0))

        track.Background = on_brush if on else off_brush
        knob = Border()
        knob.Width = knob_size
        knob.Height = knob_size
        knob.CornerRadius = knob_radius
        knob.Background = knob_brush
        knob.HorizontalAlignment = HorizontalAlignment.Left
        on_offset = track_w - knob_size - knob_margin
        knob.Margin = (
            Thickness(on_offset, knob_margin, 0, knob_margin) if on
            else Thickness(knob_margin, knob_margin, 0, knob_margin))
        track.Child = knob
        self._filter_knobs[key] = knob

    def _toggle_filter(self, key, changed_fn):
        """Flip one filter toggle's on/off state, reposition its knob
        directly (no animation - same as pyTransmit's own toggle re-read
        of live resources here, not the values captured at setup time),
        then re-run the corresponding filter."""
        on_brush = (self.TryFindResource('BrushPrimaryGreen') or
                    SolidColorBrush(Color.FromRgb(0x20, 0x8A, 0x3C)))
        off_brush = (self.TryFindResource('BrushToggleOffBg') or
                     SolidColorBrush(Color.FromRgb(0x5E, 0x5C, 0x64)))
        track_w = self.TryFindResource('WidthToggle') or 40.0
        knob_size = self.TryFindResource('SizeToggleKnob') or 16.0
        knob_margin = self.TryFindResource('MarginToggleKnob') or 2.0

        self._filter_state[key] = not self._filter_state[key]
        on = self._filter_state[key]
        track = self._filter_tracks[key]
        knob = self._filter_knobs[key]
        on_offset = track_w - knob_size - knob_margin
        knob.Margin = (
            Thickness(on_offset, knob_margin, 0, knob_margin) if on
            else Thickness(knob_margin, knob_margin, 0, knob_margin))
        track.Background = on_brush if on else off_brush
        changed_fn()

    def _source_filter_key(self, path, fd=None):
        """Which Source accordion toggle governs this file: 'ods' for a
        real .ods file (even though it's routed through the Excel code
        path), 'excel' for .xlsx/.xls, 'odt' for a real .odt file (routed
        through the Word code path), 'word' for .docx/.doc."""
        real_path = (fd.get('real_path', path) if fd else path) or path
        ext = os.path.splitext(real_path)[1].lower()
        if ext == '.ods':
            return 'ods'
        if ext == '.odt':
            return 'odt'
        return 'word' if ext in ('.docx', '.doc') else 'excel'

    def _apply_source_filter(self):
        """Show/hide cards by source type (Excel/ODS/Word/ODT). Word
        cards share one outer group across multiple views (see
        _make_word_card) — hiding the group's own outer Border is
        what actually hides the whole card; toggling each fd's
        individual view block alone would leave the shared header
        (path/date/+Add Row/reload/close) visible with nothing under
        it, which is exactly what looked broken here."""
        for path, fd in self._file_data.items():
            if fd.get('source_type') == 'word':
                continue
            border = fd.get('card_border')
            if border is None:
                continue
            key = self._source_filter_key(path, fd)
            border.Visibility = (
                Visibility.Visible if self._filter_state[key]
                else Visibility.Collapsed)
        for real_path, group in self._card_groups.items():
            outer = group.get('outer')
            if outer is None:
                continue
            key = self._source_filter_key(real_path)
            outer.Visibility = (
                Visibility.Visible if self._filter_state[key]
                else Visibility.Collapsed)

    def _apply_type_filter(self):
        """Show/hide by output View Type (Schedule/Legend/Drafting).
        Excel: each row has its own ViewType, so individual rows show
        or hide. Word: each view (not each section) has one view_type,
        so this hides/shows the whole view block via fd['card_border'],
        same reference the Source filter and view-close use."""
        checked = set()
        if self._filter_state['schedule']:
            checked.add('Schedule View')
        if self._filter_state['legend']:
            checked.add('Legend View')
        if self._filter_state['drafting']:
            checked.add('Drafting View')

        for path, fd in self._file_data.items():
            if fd.get('source_type') == 'word':
                border = fd.get('card_border')
                if border is None:
                    continue
                vt = fd.get('view_type', '')
                border.Visibility = (
                    Visibility.Visible if vt in checked else Visibility.Collapsed)
            else:
                for row in fd.get('rows', []):
                    row_ui = getattr(row, '_row_ui', None)
                    if row_ui is None:
                        continue
                    row_ui.Visibility = (
                        Visibility.Visible if row.ViewType in checked
                        else Visibility.Collapsed)

    def OnBrowse(self, sender, e):
        import clr
        clr.AddReference('System.Windows.Forms')
        import System.Windows.Forms as WF
        dlg            = WF.OpenFileDialog()
        dlg.Title      = 'Select Excel or Word files'
        dlg.Filter     = 'Supported|*.xlsx;*.xls;*.docx;*.doc;*.ods;*.odt|Excel|*.xlsx;*.xls|Word|*.docx;*.doc|LibreOffice|*.ods;*.odt'
        dlg.Multiselect = True
        if dlg.ShowDialog() == WF.DialogResult.OK:
            self._load_files(list(dlg.FileNames))

    def OnImportFolder(self, sender, e):
        import clr
        clr.AddReference('System.Windows.Forms')
        import System.Windows.Forms as WF
        dlg             = WF.FolderBrowserDialog()
        dlg.Description = 'Select folder'
        if dlg.ShowDialog() == WF.DialogResult.OK:
            supported = ('.xlsx', '.xls', '.docx', '.doc', '.ods', '.odt')
            files = [os.path.join(dlg.SelectedPath, f)
                     for f in os.listdir(dlg.SelectedPath)
                     if os.path.splitext(f)[1].lower() in supported]
            self._load_files(files)

    def OnDrop(self, sender, e):
        import System
        if e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop):
            self._load_files(list(
                e.Data.GetData(System.Windows.DataFormats.FileDrop)))

    def OnDragOver(self, sender, e):
        import System
        if e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop):
            e.Effects = System.Windows.DragDropEffects.Copy
        else:
            e.Effects = getattr(System.Windows.DragDropEffects, 'None')
        e.Handled = True

    def OnAddRow(self, sender, e):
        if not self._active_file:
            self.OnBrowse(sender, e)
            return
        fd  = self._file_data.get(self._active_file, {})
        row = Row(
            file_path      = fd.get('real_path', self._active_file),
            source_type    = fd.get('source_type', 'xl'),
            sheets         = fd.get('sheets', []),
            sheet_range_map= fd.get('sheet_range_map', {}),
        )
        fd['rows'].append(row)
        row_ui = self._make_row_ui(row)
        row_ui.Tag = row
        fd['card_panel'].Children.Add(row_ui)
        self._update_footer()
        self._refresh_status_card()
        try:
            self.RowsScroll.ScrollToBottom()
        except Exception:
            pass

    def OnRefresh(self, sender, e):
        for fd in self._file_data.values():
            for row in fd['rows']:
                row.LastModified = row._mtime(row.FilePath)
        self._refresh_status_card()
        self._set_status('Refreshed')

    def OnBatchPopupOpened(self, sender, e):
        """BatchPopup's own Opened event - fires right as the dropdown
        becomes visible (Popup.IsOpen is bound to BatchBtn.IsChecked in
        XAML, so opening/closing itself needs no code here at all)."""
        panel = self.BatchPopupPanel
        panel.Children.Clear()

        def item(label, fn):
            return self._make_menu_item(label, fn, self.BatchPopup)

        def all_rows():
            return [r for fd in self._file_data.values() for r in fd['rows']]

        panel.Children.Add(item('Select all',
            lambda s, ev: [setattr(r, 'Enabled', True) for r in all_rows()]))
        panel.Children.Add(item('Deselect all',
            lambda s, ev: [setattr(r, 'Enabled', False) for r in all_rows()]))
        panel.Children.Add(item(u'Set all \u2192 Schedule View',
            lambda s, ev: [setattr(r, 'ViewType', 'Schedule View')
                           for r in all_rows()]))
        panel.Children.Add(item(u'Set all \u2192 Legend View',
            lambda s, ev: [setattr(r, 'ViewType', 'Legend View')
                           for r in all_rows()]))

    def OnMenuOpened(self, sender, e):
        """MenuPopup's own Opened event - fires right as the hamburger
        dropdown becomes visible (Popup.IsOpen is bound to MenuBtn.IsChecked
        in XAML, same pattern as the Source/Type filter dropdowns already
        proven working in this file - no manual open/close code needed)."""
        panel = self.MenuPopupPanel
        panel.Children.Clear()

        def item(label, fn):
            return self._make_menu_item(label, fn, self.MenuPopup)

        panel.Children.Add(item(u'Edit Section Groups\u2026',
            lambda s, ev: self._open_group_settings_editor()))
        panel.Children.Add(item(u'Word Text Size\u2026',
            lambda s, ev: self._open_word_text_settings_editor()))
        panel.Children.Add(self._make_menu_separator())
        panel.Children.Add(item(u'\u2753  Support', self._menu_support_click))
        panel.Children.Add(item(u'\u2139  About pyTable', self._menu_about_click))
        panel.Children.Add(self._make_menu_separator())
        panel.Children.Add(item(u'\u2615  Support this project and help us grow',
                             self._menu_donate_click))

    _last_url_open_time = 0.0

    def _open_url(self, url, title=''):
        """Open a URL in the default browser without blocking the UI
        thread. subprocess.Popen('cmd /c start ...') spawns cmd.exe as a
        shell wrapper, and that first launch can hang for a long time
        from inside Revit's process, running synchronously on the UI
        thread would freeze the whole window for that entire time.
        os.startfile skips the shell wrapper entirely, and running it on
        a background thread means even a slow launch can never block
        the UI."""
        now_t = _time.time()
        if now_t - self._last_url_open_time < 2.0:
            return
        self._last_url_open_time = now_t

        def _launch():
            try:
                os.startfile(url)
            except Exception:
                try:
                    import subprocess
                    subprocess.Popen(['cmd', '/c', 'start', '', url])
                except Exception as ex:
                    def _show():
                        _alert("Could not open browser:\n{}".format(str(ex)), title=title)
                    try:
                        self.Dispatcher.Invoke(_Action(_show))
                    except Exception:
                        pass

        _threading.Thread(target=_launch).start()

    def _menu_support_click(self, sender, e):
        """☰ → Support: open a pre-filled support email in the default
        mail client, addressed to Seed43 support, with the extension
        version and which app it came from already filled in."""
        version = _find_seed43_version()
        subject = "pyTable Support Ticket"
        body = (
            "Hi Seed43 Team,\n\n"
            "Support Request\n\n"
            "App: pyTable\n"
            "Seed43 Version: {}\n\n"
            "Please describe your issue below:\n\n"
        ).format(version)
        import urllib
        mailto = "mailto:{}?subject={}&body={}".format(
            'support@seed43.org', urllib.quote(subject), urllib.quote(body))
        self._open_url(mailto, title="Support")

    def _menu_about_click(self, sender, e):
        _alert(
            'pyTable\n'
            'Part of the Seed43 pyRevit Extension.\n\n'
            'Links Excel and Word documents to Revit views.',
            title='About pyTable'
        )

    def _menu_donate_click(self, sender, e):
        self._open_url('https://buymeacoffee.com/seed43', title='Support')

    def OnCloseWindow(self, sender, e):
        self.Close()

    def OnApply(self, sender, e):
        all_rows = [r for fd in self._file_data.values()
                    for r in fd['rows'] if r.Enabled]
        # Cards that have rows but none of them checked — these get
        # silently skipped below with no per-row error to explain why,
        # so surface them in the status message instead.
        unchecked_cards = [
            os.path.basename(fd.get('real_path', p))
            for p, fd in self._file_data.items()
            if fd.get('rows') and not any(r.Enabled for r in fd['rows'])
        ]
        if not all_rows:
            msg = 'No rows enabled.'
            if unchecked_cards:
                msg += ' Nothing checked in: {}'.format(
                    ', '.join(unchecked_cards[:3])
                    + ('…' if len(unchecked_cards) > 3 else ''))
            self._set_status(msg)
            return

        # ── Reset all dots to pending before re-running ──
        for r in all_rows:
            if r.Status not in ('pending',):
                r.Status = 'pending'
                if r._dot:
                    r._dot.Fill = hb('#6B7280')
            if r._error_label is not None:
                r._error_label.Visibility = Visibility.Collapsed
                if r._error_text is not None:
                    r._error_text.Text = ''

        # ── Validation — mark invalid rows, skip them but run valid ones ──
        # Word rows use NamedRange (section heading) as identifier —
        # they have no ViewName so bypass that check entirely.
        seen_names = {}
        skip_rows  = set()
        for r in all_rows:
            if r.SourceType == 'word':
                # Word row is valid as long as it has a section selected
                if not r.NamedRange.strip():
                    r.Status = 'skipped'
                    if r._dot:
                        r._dot.Fill = hb('#CA8A04')
                    skip_rows.add(id(r))
                continue
            name = r.ViewName.strip()
            if not name:
                r.Status = 'skipped'
                if r._dot:
                    r._dot.Fill = hb('#CA8A04')
                skip_rows.add(id(r))
            elif name in seen_names:
                r.Status = 'error'
                if r._dot:
                    r._dot.Fill = hb('#DC2626')
                skip_rows.add(id(r))
            else:
                seen_names[name] = r

        # Re-check at run time — Word rows pass if NamedRange set,
        # Excel rows pass if ViewName non-blank
        run_rows = [r for r in all_rows
                    if id(r) not in skip_rows
                    and (r.SourceType == 'word'
                         or r.ViewName.strip())]
        if skip_rows:
            self._refresh_status_card()
        if not run_rows:
            self._set_status('{} row{} need attention — nothing to apply.'.format(
                len(skip_rows), 's' if len(skip_rows) > 1 else ''))
            return
        self._set_status('Applying...')
        # Show progress bar
        if self._progress_bar is not None:
            self._progress_bar.Value      = 0
            self._progress_bar.Visibility = Visibility.Visible
        results = []
        total_run = len(run_rows)

        # Group word rows by owning CARD, not by file — two duplicate
        # cards can point at the same underlying .docx but must still
        # produce two independent views. Grouping by real file path
        # would silently merge them into one and break the fd lookup
        # below (self._file_data is keyed by card, not by file).
        import collections as _col
        _word_by_file = _col.OrderedDict()
        _xl_rows      = []
        for r in run_rows:
            if r.SourceType == 'word':
                card_path = self._find_card_path_for_row(r) or r.FilePath
                _word_by_file.setdefault(card_path, []).append(r)
            else:
                _xl_rows.append(r)

        # Apply word files first (one view per docx)
        for card_path, wrows in _word_by_file.items():
            fd = self._file_data.get(card_path, {})
            real_path  = fd.get('real_path', card_path)
            view_name  = (fd.get('view_name') or
                          os.path.splitext(os.path.basename(real_path))[0])
            view_type  = wrows[0].ViewType
            sheet_size = fd.get('sheet_size', 'A3')
            col_count  = fd.get('col_count', 2)
            raw_sections = read_word_sections(real_path)
            labels       = get_word_headings(real_path)
            # Map display label -> section dict
            label_to_sec = {}
            for lbl, sec in zip(labels, [s for s in raw_sections
                                          if s.get('heading')]):
                label_to_sec[lbl] = sec
            sections_payload = []
            for wr in wrows:
                sec = label_to_sec.get(wr.NamedRange)
                if sec:
                    sections_payload.append({
                        'heading':    sec['heading'],
                        'paragraphs': sec['paragraphs'],
                        'col':        wr.ColNo,
                    })
            result = apply_notes_row(
                sections_payload, view_name, view_type,
                sheet_size, col_count, real_path,
                old_view_name=fd.get('_applied_view_name'))
            results.append(result)
            if result.get('status') == 'success':
                fd['_applied_view_name'] = view_name
                fd['_applied_at'] = _time.time()
                if fd.get('modified_label') is not None:
                    fd['modified_label'].Text = format_applied_at(fd['_applied_at'])
            for wr in wrows:
                wr.Status = result.get('status', 'error')
                wr.ViewName = view_name
                if wr.Status == 'success':
                    try:
                        wr._applied_mtime = os.path.getmtime(real_path)
                        wr._applied_hash = _hash_word_section(
                            real_path, wr.NamedRange)
                        wr._applied_at = _time.time()
                    except Exception:
                        pass
                if wr._dot is not None:
                    wr._dot.Fill = hb(STATUS_COLOURS.get(
                        wr.Status, '#3A4A3A'))
                self._set_sync_btn_state(wr._refresh_btn, wr.Status == 'sync')
                if wr._error_label is not None:
                    if wr.Status == 'error':
                        msg = result.get('message', '')
                        if wr._error_text is not None:
                            wr._error_text.Text = self._short_error(msg)
                        wr._error_label.Visibility = Visibility.Visible
                    else:
                        if wr._error_text is not None:
                            wr._error_text.Text = ''
                        wr._error_label.Visibility = Visibility.Collapsed
            self._refresh_status_card()

        for i, row in enumerate(_xl_rows):
            self._set_status('Processing {} of {}...'.format(
                i + 1, len(_xl_rows)))
            if self._progress_bar is not None:
                self._progress_bar.Value = int(
                    (i / float(total_run)) * 100)
            tr             = TableRow()
            tr.view_name   = row.ViewName
            tr.named_range = row.NamedRange
            tr.sheet_name  = row.Sheet
            tr.view_type   = row.ViewType
            tr.view_scale  = 1
            tr.file_path   = row.FilePath
            tr.auto_sync   = False
            result = apply_row(tr)
            results.append(result)
            row.Status = result.get('status', 'error')
            if row.Status == 'success':
                try:
                    row._applied_mtime = os.path.getmtime(row.FilePath)
                    row._applied_hash = _hash_range(
                        row.FilePath, row.NamedRange, row.Sheet)
                    row._applied_at = _time.time()
                    if row._modified_label is not None:
                        row._modified_label.Text = format_applied_at(row._applied_at)
                except Exception:
                    pass
            if row._dot is not None:
                row._dot.Fill = hb(STATUS_COLOURS.get(
                    row.Status, '#3A4A3A'))
            self._set_sync_btn_state(row._refresh_btn, row.Status == 'sync')
            # Show/hide inline error label
            if row._error_label is not None:
                if row.Status == 'error':
                    msg = result.get('message', '')
                    if row._error_text is not None:
                        row._error_text.Text    = self._short_error(msg)
                    row._error_label.Visibility = Visibility.Visible
                else:
                    if row._error_text is not None:
                        row._error_text.Text    = ''
                    row._error_label.Visibility = Visibility.Collapsed
            self._refresh_status_card()
        ok  = sum(1 for r in results if r['status'] == 'success')
        sk  = sum(1 for r in results if r['status'] == 'skipped')
        err = sum(1 for r in results if r['status'] == 'error')
        summary = '{} ok  {} skip  {} err'.format(ok, sk, err)
        if unchecked_cards:
            summary += '  ({} card(s) had nothing checked: {})'.format(
                len(unchecked_cards),
                ', '.join(unchecked_cards[:3])
                + ('…' if len(unchecked_cards) > 3 else ''))
        self._set_status(summary)
        # Save state to Revit parameter
        self._save_persisted_state()
        # Complete and hide progress bar
        if self._progress_bar is not None:
            self._progress_bar.Value      = 100
            import threading, System
            def _hide():
                import time
                time.sleep(1.0)
                self._progress_bar.Dispatcher.Invoke(
                    System.Action(lambda: setattr(
                        self._progress_bar, 'Visibility', Visibility.Collapsed)))
            threading.Thread(target=_hide).start()
        self._refresh_status_card()

    # ── Status card ──

    def _all_rows(self):
        return [r for fd in self._file_data.values() for r in fd['rows']]

    def _build_status_card(self):
        """Create the status summary card — always visible above file cards."""
        card = Border()
        card.Background   = hb('#2B3340')
        card.CornerRadius = CornerRadius(8)
        card.Padding      = Thickness(14, 10, 14, 10)
        card.Margin       = Thickness(0, 0, 0, 0)
        try:
            from System.Windows.Media.Effects import DropShadowEffect
            fx = DropShadowEffect()
            fx.Color       = Color.FromRgb(0, 0, 0)
            fx.Opacity     = 0.2
            fx.ShadowDepth = 1
            fx.BlurRadius  = 4
            card.Effect    = fx
        except Exception:
            pass

        # Progress bar — hidden when idle, fills left-to-right during Apply
        from System.Windows.Controls import ProgressBar
        self._progress_bar = ProgressBar()
        self._progress_bar.Minimum   = 0
        self._progress_bar.Maximum   = 100
        self._progress_bar.Value     = 0
        self._progress_bar.Height    = 3
        self._progress_bar.Margin    = Thickness(0, 0, 0, 8)
        self._progress_bar.Visibility = Visibility.Collapsed
        try:
            self._progress_bar.Foreground = hb('#208A3C')
            self._progress_bar.Background = hb('#404553')
            self._progress_bar.BorderThickness = Thickness(0)
        except Exception:
            pass

        row = StackPanel()
        row.Orientation      = Orientation.Horizontal
        row.VerticalAlignment = VerticalAlignment.Center

        def _item(dot_colour, label_text):
            """One dot + label pair."""
            sp = StackPanel()
            sp.Orientation       = Orientation.Horizontal
            sp.VerticalAlignment = VerticalAlignment.Center
            sp.Margin            = Thickness(0, 0, 24, 0)
            d = Ellipse()
            d.Width              = 10
            d.Height             = 10
            d.Fill               = hb(dot_colour)
            d.VerticalAlignment  = VerticalAlignment.Center
            d.Margin             = Thickness(0, 0, 7, 0)
            lbl = TextBlock()
            lbl.FontSize         = 12
            lbl.Foreground       = hb('#F4FAFF')
            lbl.VerticalAlignment = VerticalAlignment.Center
            sp.Children.Add(d)
            sp.Children.Add(lbl)
            lbl.Text = label_text
            return sp, lbl

        sp_p, self._lbl_pending = _item('#6B7280', 'Pending')
        sp_s, self._lbl_success = _item('#16A34A', 'Success')
        sp_e, self._lbl_error   = _item('#DC2626', 'Error')
        sp_k, self._lbl_skipped = _item('#CA8A04', 'Skipped')

        for sp in (sp_p, sp_s, sp_e, sp_k):
            row.Children.Add(sp)

        # Divider
        div = Border()
        div.Width             = 1
        div.Height            = 16
        div.Background        = hb('#404553')
        div.VerticalAlignment = VerticalAlignment.Center
        div.Margin            = Thickness(0, 0, 24, 0)
        row.Children.Add(div)

        # Sync indicator (hidden until needed)
        self._sync_panel = StackPanel()
        self._sync_panel.Orientation      = Orientation.Horizontal
        self._sync_panel.VerticalAlignment = VerticalAlignment.Center
        self._sync_panel.Visibility       = Visibility.Visible

        sync_dot = Ellipse()
        sync_dot.Width             = 10
        sync_dot.Height            = 10
        sync_dot.Fill              = hb('#3B82F6')
        sync_dot.VerticalAlignment = VerticalAlignment.Center
        sync_dot.Margin            = Thickness(0, 0, 7, 0)

        self._sync_label = TextBlock()
        self._sync_label.FontSize          = 12
        self._sync_label.Foreground        = hb('#F4FAFF')
        self._sync_label.VerticalAlignment = VerticalAlignment.Center

        self._sync_label.Text    = '0 update required'
        self._sync_label.Opacity = 0.9
        self._sync_panel.Children.Add(sync_dot)
        self._sync_panel.Children.Add(self._sync_label)
        row.Children.Add(self._sync_panel)

        # Blurb
        blurb = TextBlock()
        blurb.Text = (
            'View Name auto-fills from the selected range. '
            'Clear the View Name to allow it to update when the range changes.'
        )
        blurb.FontSize    = 10
        blurb.Foreground  = hb('#F4FAFF')
        blurb.Opacity     = 0.35
        from System.Windows import TextWrapping
        blurb.TextWrapping = TextWrapping.Wrap
        blurb.Margin      = Thickness(0, 8, 0, 0)

        outer_stack = StackPanel()
        outer_stack.Orientation = Orientation.Vertical
        outer_stack.Children.Add(self._progress_bar)
        outer_stack.Children.Add(row)
        outer_stack.Children.Add(blurb)

        card.Child            = outer_stack
        self._status_card     = card
        self.StatusCardHost.Child = card

    def _refresh_status_card(self):
        """Update counts and sync indicator on the status card."""
        if self._status_card is None or self._lbl_pending is None:
            return

        rows = self._all_rows()
        counts = {'pending': 0, 'success': 0, 'error': 0, 'skipped': 0}
        for r in rows:
            s = r.Status
            if s == 'sync':
                counts['success'] += 1  # view exists, counts as success
            elif s in counts:
                counts[s] += 1
            else:
                counts['pending'] += 1

        # Format: "2 Pending" — number first, muted if zero
        def _fmt(n, label):
            return '{} {}'.format(n, label)

        self._lbl_pending.Text    = _fmt(counts['pending'], 'Pending')
        self._lbl_success.Text    = _fmt(counts['success'], 'Success')
        self._lbl_error.Text      = _fmt(counts['error'],   'Error')
        self._lbl_skipped.Text    = _fmt(counts['skipped'], 'Skipped')

        # Dim zero-count items
        self._lbl_pending.Opacity  = 0.4 if counts['pending']  == 0 else 0.9
        self._lbl_success.Opacity  = 0.4 if counts['success']  == 0 else 0.9
        self._lbl_error.Opacity    = 0.4 if counts['error']     == 0 else 0.9
        self._lbl_skipped.Opacity  = 0.4 if counts['skipped']   == 0 else 0.9

        # Out-of-sync check — use content hash if available, else mtime
        stale = []
        seen_rows = set()
        for r in rows:
            if getattr(r, '_unlinked', False):
                continue
            key = '{}|{}'.format(r.FilePath, r.NamedRange)
            if key in seen_rows:
                continue
            seen_rows.add(key)
            try:
                if r.SourceType == 'word':
                    # Per-section hash — same pattern as Excel
                    if r._applied_hash:
                        cur = _hash_word_section(r.FilePath, r.NamedRange)
                        if cur and cur != r._applied_hash:
                            stale.append(os.path.basename(r.FilePath))
                    elif r._applied_mtime:
                        if os.path.getmtime(r.FilePath) > r._applied_mtime + 1:
                            stale.append(os.path.basename(r.FilePath))
                elif r._applied_hash:
                    current_hash = _hash_range(r.FilePath, r.NamedRange, r.Sheet)
                    if current_hash and current_hash != r._applied_hash:
                        stale.append(os.path.basename(r.FilePath))
                elif r._applied_mtime:
                    if os.path.getmtime(r.FilePath) > r._applied_mtime + 1:
                        stale.append(os.path.basename(r.FilePath))
            except Exception:
                pass

        # Sync indicator is always blue — acts as legend + counter
        n_stale = len(stale)
        self._sync_label.Text       = '{} update required'.format(n_stale)
        self._sync_label.Foreground = hb('#F4FAFF')
        self._sync_label.Opacity    = 0.4 if n_stale == 0 else 0.9

        # Keep every card's own reload indicator (grey/green) in sync too
        for card_path in list(self._file_data.keys()):
            self._update_card_reload_indicator(card_path)

    # ── Persistence ──

    def _save_persisted_state(self):
        """Write current UI state to the pyTable shared parameter."""
        try:
            save_pytable_state(self._file_data)
        except Exception as ex:
            logger.warning('pyTable save: {}'.format(ex))

    def _load_persisted_state(self):
        """
        Read pyTable shared parameter and rebuild cards + rows.
        After rebuilding, check if any source files have been modified
        since the parameter was last written (triggers sync indicator).
        """
        try:
            cards = load_pytable_state()
        except Exception as ex:
            logger.warning('pyTable load: {}'.format(ex))
            return
        if not cards:
            return

        for card in cards:
            path = card.get('path', '')
            # For duplicate cards, path is a synthetic key (real file
            # path + a " (copy N)" suffix) and will never exist on
            # disk itself — check the real underlying file instead.
            check_path = card.get('real_path') or path
            if not path or not os.path.exists(check_path):
                logger.warning('pyTable restore: file not found: {}'.format(check_path))
                continue
            # Parsing below reads the actual file content, so use the
            # real path there too — 'path' may be the synthetic key.
            parse_path = card.get('real_path') or path
            # Parse the file to get current sheets/ranges
            if path not in self._file_data:
                ext     = os.path.splitext(parse_path)[1].lower()
                is_word = ext not in ('.xlsx', '.xls', '.ods')
                if is_word:
                    # Word file — use heading parser, not workbook parser
                    try:
                        headings = get_word_headings(parse_path)
                    except Exception:
                        headings = []
                    if not headings:
                        headings = ['(no headings found)']
                    sheets = ['Document']
                    srmap  = {'Document': headings}
                else:
                    try:
                        wb = get_named_ranges_from_workbook(parse_path)
                    except Exception:
                        wb = {}
                    sheets = wb.get('sheets', [])
                    srmap  = wb.get('sheet_ranges', {})
                    if not srmap:
                        all_r = wb.get('named_ranges', [])
                        srmap = {s: all_r for s in sheets}
                saved_rows = card.get('rows', [])
                migrated_view_type = (
                    saved_rows[0].get('view_type', WORD_VIEW_TYPES[0])
                    if saved_rows else WORD_VIEW_TYPES[0]
                )
                self._file_data[path] = {
                    'sheets':          sheets,
                    'sheet_range_map': srmap,
                    'source_type':     'word' if is_word else 'xl',
                    'rows':            [],
                    'card_panel':      None,
                    'card_border':     None,
                    'sheet_size':      card.get('sheet_size', 'A3 Landscape'),
                    'col_count':       card.get('col_count', 2),
                    'view_name':       card.get('view_name', ''),
                    'view_type':       card.get('view_type') or migrated_view_type,
                    'real_path':       card.get('real_path') or path,
                    'unlinked':        card.get('unlinked', False),
                    'path_mode':       card.get('path_mode', 'absolute'),
                    'layout_mode':     card.get('layout_mode', 'manual'),
                    '_applied_view_name': card.get('applied_view_name', ''),
                    '_applied_at':     card.get('applied_at'),
                }
                self._make_card(path)

            # Restore rows
            fd = self._file_data[path]
            for rdata in card.get('rows', []):
                row = Row(
                    file_path      = fd.get('real_path', path),
                    source_type    = fd['source_type'],
                    sheets         = fd['sheets'],
                    sheet_range_map= fd['sheet_range_map'],
                )
                row._unlinked       = fd.get('unlinked', False)
                row.ViewName        = rdata.get('view_name', '')
                row.Sheet           = rdata.get('sheet', '')
                row.ColNo           = int(rdata.get('col_no', 1))
                row.NamedRange      = rdata.get('named_range', '')
                row.ViewType        = rdata.get('view_type', 'Schedule View')
                row.Priority        = rdata.get('priority', 'Medium')
                row.Group           = rdata.get('group', '')
                row._applied_mtime  = rdata.get('applied_mtime', None)
                row._applied_hash   = rdata.get('applied_hash', None)
                row._applied_at     = rdata.get('applied_at', None)

                # ── Check Revit view existence and sync status ──
                row.Status = self._check_row_status(row, path)

                fd['rows'].append(row)
                row_ui     = self._make_row_ui(row)
                row_ui.Tag = row
                # Sync ViewName textbox with restored value
                if row._vn_textbox is not None:
                    row._vn_textbox.Text = row.ViewName
                # Apply dot colour from restored status
                if row._dot is not None:
                    row._dot.Fill = hb(STATUS_COLOURS.get(row.Status, '#6B7280'))
                self._set_sync_btn_state(row._refresh_btn, row.Status == 'sync')
                fd['card_panel'].Children.Add(row_ui)

        self._update_file_combo()
        self._update_footer()
        self._refresh_status_card()
        self._revalidate_all_view_name_boxes()
        logger.debug('pyTable: restored {} card(s)'.format(len(cards)))

    def _check_row_status(self, row, card_key):
        """
        On load, determine row status by checking:
        1. Does the Revit view exist with that name?
        2. If yes, has the source content changed since last apply?

        card_key looks up the card's own settings (view_name etc), all
        actual file reads use row.FilePath, the real path, since a card's
        dict key and its underlying file path can differ once duplicate
        cards exist for the same file.

        Returns: 'success', 'pending', or 'sync'.
        """
        if getattr(row, '_unlinked', False):
            # Card was explicitly unlinked from its source — leave
            # whatever status it already had, don't check the file.
            return row.Status
        try:
            from pyrevit import revit, DB
            from pyrevit.revit import query

            file_path = row.FilePath

            # Word rows use the card-level view name (filename stem)
            if row.SourceType == 'word':
                fd = self._file_data.get(card_key, {})
                view_name = (fd.get('view_name') or
                             os.path.splitext(os.path.basename(file_path))[0])
            else:
                view_name = row.ViewName.strip()
            if not view_name:
                return 'pending'

            # Check if a view with this name exists
            found = False
            for v in query.get_elements_by_class(DB.View, doc=revit.doc):
                try:
                    if v.Name == view_name:
                        found = True
                        break
                except Exception:
                    continue

            if not found:
                return 'pending'

            # View exists — compare content hash to detect changes
            try:
                if row.SourceType == 'word':
                    # Per-section hash for Word rows
                    current_hash = _hash_word_section(file_path, row.NamedRange)
                    if current_hash and row._applied_hash:
                        if current_hash != row._applied_hash:
                            return 'sync'
                        else:
                            row._applied_hash = current_hash
                            return 'success'
                    elif current_hash and not row._applied_hash:
                        row._applied_hash = current_hash
                        return 'success'
                    return 'success'
                current_hash = _hash_range(file_path, row.NamedRange, row.Sheet)

                if current_hash and row._applied_hash:
                    if current_hash != row._applied_hash:
                        # Content changed — blue dot
                        return 'sync'
                    else:
                        # Up to date
                        row._applied_mtime = os.path.getmtime(file_path)
                        row._applied_hash  = current_hash
                        return 'success'
                elif current_hash and not row._applied_hash:
                    # First open after hash tracking added — store now
                    row._applied_hash  = current_hash
                    row._applied_mtime = os.path.getmtime(file_path)
                    return 'success'
                else:
                    # Hash failed — mtime fallback
                    file_mtime = os.path.getmtime(file_path)
                    if row._applied_mtime and file_mtime > row._applied_mtime + 1:
                        return 'sync'
                    row._applied_mtime = file_mtime
                    return 'success'
            except Exception as ex:
                logger.warning('pyTable hash check failed: {}'.format(ex))
                return 'success'

        except Exception as ex:
            logger.debug('_check_row_status failed: {}'.format(ex))
            return 'pending'

    def _refresh_row_click(self, sender, e):
        """Re-apply a single stale (blue) row without touching others."""
        row = sender.Tag
        if row is None:
            return
        # Clear error pill
        if row._error_label is not None:
            row._error_label.Visibility = Visibility.Collapsed
        if row._error_text is not None:
            row._error_text.Text = ''
        # Set dot to pending while running
        if row._dot is not None:
            row._dot.Fill = hb('#6B7280')
        self._set_sync_btn_state(row._refresh_btn, False)
        self._set_status('Refreshing {}...'.format(
            row.NamedRange if row.SourceType == 'word' else row.ViewName))
        try:
            if row.SourceType == 'word':
                # Word row: re-apply just this section via apply_notes_row
                card_path  = self._find_card_path_for_row(row) or row.FilePath
                fd         = self._file_data.get(card_path, {})
                real_path  = fd.get('real_path', card_path)
                view_name  = (fd.get('view_name') or
                              os.path.splitext(os.path.basename(real_path))[0])
                raw_secs   = read_word_sections(real_path)
                labels     = get_word_headings(real_path)
                label_map  = {lbl: s for lbl, s in zip(
                    labels, [s for s in raw_secs if s.get('heading')])}
                sec = label_map.get(row.NamedRange)
                sections_payload = []
                if sec:
                    sections_payload.append({
                        'heading':    sec['heading'],
                        'paragraphs': sec['paragraphs'],
                        'col':        row.ColNo,
                    })
                # Collect all rows for same file to rebuild view —
                # include success, sync, and the current row being refreshed
                all_word_rows = [r for r in fd.get('rows', [])
                                 if r.Status in ('success', 'sync')
                                 or r is row]
                full_payload  = []
                for wr in all_word_rows:
                    s = label_map.get(wr.NamedRange)
                    if s:
                        full_payload.append({
                            'heading':    s['heading'],
                            'paragraphs': s['paragraphs'],
                            'col':        wr.ColNo,
                        })
                result = apply_notes_row(
                    full_payload, view_name,
                    row.ViewType,
                    fd.get('sheet_size', 'A3 Landscape'),
                    fd.get('col_count', 2),
                    real_path,
                    old_view_name=fd.get('_applied_view_name'))
                if result.get('status') == 'success':
                    fd['_applied_view_name'] = view_name
                    fd['_applied_at'] = _time.time()
                    if fd.get('modified_label') is not None:
                        fd['modified_label'].Text = format_applied_at(fd['_applied_at'])
                row.Status = result.get('status', 'error')
                if row.Status == 'success':
                    try:
                        row._applied_mtime = os.path.getmtime(real_path)
                        row._applied_hash  = _hash_word_section(
                            real_path, row.NamedRange)
                        row._applied_at = _time.time()
                    except Exception:
                        pass
            else:
                tr             = TableRow()
                tr.view_name   = row.ViewName
                tr.named_range = row.NamedRange
                tr.sheet_name  = row.Sheet
                tr.view_type   = row.ViewType
                tr.view_scale  = 1
                tr.file_path   = row.FilePath
                tr.auto_sync   = False
                result = apply_row(tr)
                row.Status = result.get('status', 'error')
                if row.Status == 'success':
                    try:
                        row._applied_mtime = os.path.getmtime(row.FilePath)
                        row._applied_hash  = _hash_range(
                            row.FilePath, row.NamedRange, row.Sheet)
                        row._applied_at = _time.time()
                        if row._modified_label is not None:
                            row._modified_label.Text = format_applied_at(row._applied_at)
                    except Exception:
                        pass
            if row._dot is not None:
                row._dot.Fill = hb(STATUS_COLOURS.get(
                    row.Status, '#3A4A3A'))
            self._set_sync_btn_state(row._refresh_btn, row.Status == 'sync')
            if row._error_label is not None:
                if row.Status == 'error':
                    msg = result.get('message', '')
                    if row._error_text is not None:
                        row._error_text.Text = self._short_error(msg)
                    row._error_label.Visibility = Visibility.Visible
            self._set_status('{} refreshed.'.format(row.ViewName))
        except Exception as ex:
            row.Status = 'error'
            if row._dot is not None:
                row._dot.Fill = hb(STATUS_COLOURS.get('error', '#DC2626'))
            self._set_status('Refresh failed: {}'.format(ex))
        self._save_persisted_state()
        self._refresh_status_card()

    # ── Word row event handlers ──

    def _on_window_closing(self, sender, e):
        """Save state when window is closed so deletions are persisted."""
        try:
            self._save_persisted_state()
        except Exception:
            pass

    def _make_sync_btn(self, row):
        """Round reload button shown at the end of every row - accent
        green + enabled when the source has changed and needs reloading,
        the canonical disabled look when already in sync. Uses
        RoundPrimaryButtonStyle exactly as defined in the lib HTML
        palette editor (IsEnabled drives the colour, not a manual
        Background assignment)."""
        rb = Button()
        if _mi is not None:
            try:
                rb.Content = _mi('reload', size=13, color='#FFFFFF')
            except Exception:
                rb.Content = u'\u21bb'
        else:
            rb.Content = u'\u21bb'
        try:
            rb.Style = self.FindResource('RoundPrimaryButtonStyle')
        except Exception as e:
            logger.warning('Failed to apply RoundPrimaryButtonStyle: {}'.format(e))
        rb.FocusVisualStyle = None
        rb.Width       = 24
        rb.Height      = 24
        rb.HorizontalContentAlignment = HorizontalAlignment.Center
        rb.VerticalContentAlignment   = VerticalAlignment.Center
        rb.VerticalAlignment = VerticalAlignment.Center
        rb.Margin  = Thickness(4, 0, 0, 0)
        rb.Tag     = row
        rb.Click  += self._refresh_row_click
        self._set_sync_btn_state(rb, row.Status == 'sync')
        return rb

    def _set_sync_btn_state(self, rb, needs_reload):
        """Enabled state + tooltip for a row's reload button.
        RoundPrimaryButtonStyle handles the colour switch itself via
        IsEnabled (accent green when enabled, canonical disabled look
        when not) - no manual colour assignment needed."""
        if rb is None:
            return
        rb.Visibility = Visibility.Visible
        rb.IsEnabled = needs_reload
        rb.ToolTip = ('Source changed - click to reload' if needs_reload
                      else 'Up to date')

    def _green_btn(self, content, height=24, padding=(10,0,10,0),
                   font_size=11, width=None):
        """Rounded green button matching SmallButtonStyle."""
        btn = Button()
        btn.Content = content
        btn.Height  = height
        btn.Padding = Thickness(padding[0], padding[1], padding[2], padding[3])
        btn.FontSize = font_size
        if width is not None:
            btn.Width = width
        try:
            btn.Style = self.FindResource('SmallButtonStyle')
        except Exception as e:
            logger.warning('Failed to apply SmallButtonStyle: {}'.format(e))
        btn.FocusVisualStyle = None
        return btn

    def _add_row_for_card(self, sender, e):
        """Add a new row to a specific card regardless of active file."""
        path = sender.Tag
        if path not in self._file_data:
            return
        prev_active      = self._active_file
        self._active_file = path
        self._update_file_combo()
        self.OnAddRow(None, None)
        self._active_file = prev_active
        self._update_file_combo()
        self._maybe_run_strict_layout(path)

    def _del_card_click(self, sender, e):
        """Remove an entire file card and all its rows."""
        path = sender.Tag
        if path not in self._file_data:
            return
        name = os.path.basename(path)
        if not _confirm(
                'Remove card for {}\nThis will delete all rows for this file.'.format(name),
                title='Confirm Remove'):
            return
        # Remove card Border from CardsPanel
        card_border = self._file_data[path].get('card_border')
        if card_border is not None:
            try:
                self.CardsPanel.Children.Remove(card_border)
            except Exception:
                pass
        # Remove from data
        del self._file_data[path]
        # Update active file
        if self._active_file == path:
            paths = list(self._file_data.keys())
            self._active_file = paths[0] if paths else None
        self._update_file_combo()
        self._update_footer()
        self._refresh_status_card()
        self._refresh_drop_zone()
        self._save_persisted_state()

    def _toggle_card_collapse(self, sender, e):
        """Show/hide everything in the card except its header row.
        sender is a ToggleButton (PrimarySecondaryToggleButtonStyle) -
        WPF flips IsChecked natively before Click fires, so IsChecked
        already reflects the new state: True = now expanded, False =
        now collapsed."""
        path = sender.Tag
        fd = self._file_data.get(path)
        if fd is None or fd.get('card_inner') is None:
            return
        inner = fd['card_inner']
        children = list(inner.Children)
        if not children:
            return
        expanded = bool(sender.IsChecked)
        for child in children[1:]:
            child.Visibility = Visibility.Visible if expanded else Visibility.Collapsed
        # header_row (children[0]) carries its own bottom margin to
        # space it from the rows below. When everything below is
        # hidden, that margin is left stacking on top of the card's
        # own bottom padding, making the collapsed card look bottom-
        # heavy — zero it out while collapsed, restore on expand.
        header_row = children[0]
        header_row.Margin = Thickness(0, 0, 0, 8 if expanded else 0)
        self._set_collapse_icon(sender, not expanded)

    def _set_collapse_icon(self, btn, collapsed):
        """Set the collapse toggle's chevron glyph + tooltip to match
        state. Colour comes entirely from PrimarySecondaryToggleButtonStyle's
        own IsChecked trigger now (accent green while collapsed, secondary
        grey while expanded) - the glyph itself stays a plain white/
        style-Foreground chevron rather than a baked-colour compound icon,
        since the button's own round background already carries the
        colour and a circle-backed icon glyph on top of it would double
        up. No canonical plain chevron_right icon exists yet (only the
        _circle compound variants), so this uses a plain unicode
        triangle both ways for visual consistency."""
        btn.Content  = u'\u25B6' if collapsed else u'\u25BC'
        btn.FontSize = 11
        btn.ToolTip  = 'Expand' if collapsed else 'Collapse'

    def _update_card_link_badge(self, path):
        """Refresh the card heading text to reflect unlinked state and
        absolute/relative path display mode."""
        fd = self._file_data.get(path)
        if fd is None:
            return
        lbl = fd.get('heading_label')
        if lbl is None:
            return
        real_path = fd.get('real_path', path)
        display = real_path
        if fd.get('path_mode') == 'relative':
            try:
                base_dir = os.path.dirname(doc.PathName) if doc.PathName else None
                if base_dir:
                    display = os.path.relpath(real_path, base_dir)
            except Exception:
                pass
        if fd.get('unlinked'):
            display = display + u'  \u2014 unlinked'
        lbl.Text    = display
        lbl.ToolTip = real_path

    # ── Per-card Batch menu ──

    def _card_batch_menu(self, sender, e):
        """Open the per-card Batch dropdown, separate from the global
        toolbar Batch dropdown — everything here acts on this one card.
        A real Popup+MenuItemStyle dropdown built once per card and
        reused (stored on fd), same visual pattern as the toolbar
        Batch/hamburger dropdowns, not a native ContextMenu. sender is a
        plain Button (no IsChecked to bind Popup.IsOpen to), so this one
        still opens/closes itself imperatively - StaysOpen="False" on
        the popup still gets WPF's native outside-click dismissal."""
        path = sender.Tag
        fd = self._file_data.get(path, {})

        popup_panel = fd.get('batch_popup_panel')
        if popup_panel is None:
            popup, popup_panel = self._build_dropdown_popup(sender)
            fd['batch_popup']       = popup
            fd['batch_popup_panel'] = popup_panel
        else:
            popup = fd['batch_popup']

        if popup.IsOpen:
            popup.IsOpen = False
            return
        popup_panel.Children.Clear()

        def item(label, fn):
            return self._make_menu_item(label, fn, popup)

        popup_panel.Children.Add(item('Delete selected',
            lambda s, ev: self._card_delete_selected(path)))
        popup_panel.Children.Add(self._make_menu_separator())
        popup_panel.Children.Add(item('Duplicate',
            lambda s, ev: self._card_duplicate(path)))
        popup_panel.Children.Add(item('Absolute/Relative Path',
            lambda s, ev: self._card_toggle_path_mode(path)))
        popup_panel.Children.Add(self._make_menu_separator())
        popup_panel.Children.Add(item('Open File',
            lambda s, ev: self._card_open_file(path)))
        popup_panel.Children.Add(item('Open Folder',
            lambda s, ev: self._card_open_folder(path)))
        if fd.get('unlinked'):
            popup_panel.Children.Add(item('Re-link View',
                lambda s, ev: self._card_relink_view(path)))
        else:
            popup_panel.Children.Add(item('Unlink View',
                lambda s, ev: self._card_unlink_view(path)))
        popup_panel.Children.Add(item('Remove view',
            lambda s, ev: self._card_remove_views(path)))
        popup.IsOpen = True

    def _card_delete_selected(self, path):
        """Delete every checked row across every view in this card
        (Batch is one per document now, not one per view — so this
        spans the whole real_path group for word docs)."""
        fd = self._file_data.get(path)
        if fd is None:
            return
        if fd.get('source_type') == 'word':
            real_path = fd.get('real_path', path)
            group = self._card_groups.get(real_path)
            paths = list(group['view_keys']) if group else [path]
        else:
            paths = [path]

        selected = []
        for p in paths:
            f = self._file_data.get(p)
            if f:
                selected.extend(r for r in f.get('rows', []) if r.Enabled)
        if not selected:
            _alert('No rows are checked in this card.',
                   title='Nothing selected')
            return
        if not _confirm(
                'Delete {} checked row(s) from this card?'.format(len(selected)),
                title='Confirm Delete'):
            return
        for row in list(selected):
            self._remove_row_ui(row)
        self._update_footer()
        self._refresh_status_card()
        self._save_persisted_state()
        for p in paths:
            self._maybe_run_strict_layout(p)

    def _next_duplicate_key(self, real_path):
        """Generate a dict key for a second/third/... card pointed at
        the same underlying file — used both by the explicit Duplicate
        batch action and by re-picking an already-loaded file."""
        n = 2
        new_key = '{} (copy {})'.format(real_path, n)
        while new_key in self._file_data:
            n += 1
            new_key = '{} (copy {})'.format(real_path, n)
        return new_key

    def _card_duplicate(self, path):
        """Clone this card (same source file, same rows) into a new
        card with its own dict key, so it can be set up as a second,
        independent table from the same file. Cloned rows start blank
        on ViewName/status so Apply doesn't collide with the original
        card's Revit views."""
        fd = self._file_data.get(path)
        if fd is None:
            return
        real_path = fd.get('real_path', path)
        new_key = self._next_duplicate_key(real_path)

        new_fd = {
            'sheets':          list(fd.get('sheets', [])),
            'sheet_range_map': dict(fd.get('sheet_range_map', {})),
            'source_type':     fd.get('source_type', 'xl'),
            'rows':            [],
            'card_panel':      None,
            'card_border':     None,
            'real_path':       real_path,
            'unlinked':        False,
            'path_mode':       'absolute',
        }
        if fd.get('source_type') == 'word':
            new_fd['sheet_size'] = fd.get('sheet_size', 'A3 Landscape')
            new_fd['col_count']  = fd.get('col_count', 2)
            new_fd['view_name']  = (fd.get('view_name', '') + ' Copy').strip()
            new_fd['view_type']  = fd.get('view_type', WORD_VIEW_TYPES[0])

        self._file_data[new_key] = new_fd
        self._make_card(new_key)

        for row in fd.get('rows', []):
            nr = Row(
                file_path       = real_path,
                source_type     = row.SourceType,
                sheets          = row._sheets,
                sheet_range_map = row._sheet_range_map,
            )
            nr.ViewName   = ''   # forces a fresh view on next Apply
            nr.Sheet      = row.Sheet
            nr.NamedRange = row.NamedRange
            nr.ViewType   = row.ViewType
            nr.ColNo      = row.ColNo
            nr.Enabled    = row.Enabled
            new_fd['rows'].append(nr)
            row_ui = self._make_row_ui(nr)
            row_ui.Tag = nr
            new_fd['card_panel'].Children.Add(row_ui)

        self._update_file_combo()
        self._update_footer()
        self._refresh_status_card()
        self._revalidate_all_view_name_boxes()
        self._save_persisted_state()
        self._set_status('Duplicated card')

    def _card_toggle_path_mode(self, path):
        """Flip between storing this card's link as an absolute path
        or relative to the current .rvt's folder. Relative mode means
        the link survives the project folder (rvt + source files
        together) being moved or copied to a new machine — on next
        load, the path is re-resolved against wherever the .rvt is
        now, rather than the original absolute location."""
        fd = self._file_data.get(path)
        if fd is None:
            return
        current = fd.get('path_mode', 'absolute')
        going_relative = current == 'absolute'
        if going_relative and not _doc_base_dir():
            _alert(
                'Save this Revit file first \u2014 relative paths need '
                'somewhere to be relative to.',
                title='Absolute/Relative Path')
            return
        fd['path_mode'] = 'relative' if going_relative else 'absolute'
        self._update_card_link_badge(path)
        self._save_persisted_state()
        self._set_status('Path storage: {}'.format(fd['path_mode']))

    def _card_open_file(self, path):
        fd = self._file_data.get(path)
        real_path = fd.get('real_path', path) if fd else path
        try:
            os.startfile(real_path)
        except Exception as ex:
            _alert('Could not open file:\n{}'.format(ex), title='Open File')

    def _card_open_folder(self, path):
        fd = self._file_data.get(path)
        real_path = fd.get('real_path', path) if fd else path
        folder = os.path.dirname(real_path)
        try:
            os.startfile(folder)
        except Exception as ex:
            _alert('Could not open folder:\n{}'.format(ex), title='Open Folder')

    def _card_unlink_view(self, path):
        """Detach this card from its source file. The card and its
        rows stay exactly as they are, but pyTable stops checking the
        file for changes — no more blue 'update required' dots from
        this card. Does NOT touch the Revit view itself."""
        fd = self._file_data.get(path)
        if fd is None:
            return
        if fd.get('unlinked'):
            _alert('This card is already unlinked.', title='Unlink View')
            return
        if not _confirm(
                'Unlink this card from its source file?\n'
                'The card and its rows stay, but changes to the file '
                'will no longer be detected.',
                title='Confirm Unlink'):
            return
        fd['unlinked'] = True
        for row in fd.get('rows', []):
            row._unlinked = True
        self._update_card_link_badge(path)
        self._refresh_status_card()
        self._save_persisted_state()

    def _card_relink_view(self, path):
        """Undo Unlink View — resume tracking this card's source file
        for changes. No confirm dialog needed, unlike Unlink, since
        this is purely re-enabling detection, not discarding anything."""
        fd = self._file_data.get(path)
        if fd is None:
            return
        if not fd.get('unlinked'):
            _alert('This card is already linked.', title='Link View')
            return
        fd['unlinked'] = False
        for row in fd.get('rows', []):
            row._unlinked = False
        self._update_card_link_badge(path)
        self._refresh_status_card()
        self._save_persisted_state()

    def _card_remove_views(self, path):
        """Delete the Revit view(s) this card previously created,
        without touching the card or its link to the source file.
        Rows fall back to 'pending' so the next Apply recreates them."""
        fd = self._file_data.get(path)
        if fd is None:
            return
        rows = [r for r in fd.get('rows', []) if r.ViewName.strip()]
        if not rows:
            _alert('No applied views to remove in this card.',
                   title='Nothing to remove')
            return
        if not _confirm(
                'Remove {} Revit view(s) created by this card?\n'
                'The card and its file link are kept.'.format(len(rows)),
                title='Confirm Remove Views'):
            return
        from pyrevit.revit import query
        removed = 0
        with revit.Transaction('pyTable - Remove views'):
            for row in rows:
                view_name = row.ViewName.strip()
                for v in query.get_elements_by_class(DB.View, doc=doc):
                    try:
                        if v.Name == view_name:
                            doc.Delete(v.Id)
                            removed += 1
                            break
                    except Exception:
                        continue
                row.Status          = 'pending'
                row._applied_hash   = None
                row._applied_mtime  = None
                if row._dot is not None:
                    row._dot.Fill = hb(STATUS_COLOURS.get('pending', '#6B7280'))
                self._set_sync_btn_state(row._refresh_btn, False)
        self._set_status('Removed {} view(s)'.format(removed))
        self._refresh_status_card()
        self._save_persisted_state()

    def _card_reload_click(self, sender, e):
        """Reapply every row in this card that's currently flagged as
        needing an update (the same thing each row's own per-row reload
        button does, just for the whole card at once)."""
        path = sender.Tag
        fd = self._file_data.get(path)
        if fd is None:
            return
        stale_rows = [r for r in fd.get('rows', []) if r.Status == 'sync']
        if not stale_rows:
            return

        class _FakeSender(object):
            def __init__(self, tag):
                self.Tag = tag

        for row in stale_rows:
            self._refresh_row_click(_FakeSender(row), None)
        self._update_card_reload_indicator(path)

    def _update_card_reload_indicator(self, path):
        """Enable (accent green) when at least one row in this card
        needs reapplying, disable (canonical disabled look) when
        everything's in sync - same IsEnabled-driven RoundPrimaryButtonStyle
        as the per-row reload button, so card-level and row-level agree.
        Word cards also update the group-level aggregate (multiple views
        can share one card)."""
        fd = self._file_data.get(path)
        if fd is None:
            return
        needs_reload = any(r.Status == 'sync' for r in fd.get('rows', []))
        btn = fd.get('reload_btn')
        if btn is not None:
            btn.IsEnabled = needs_reload
            btn.ToolTip = ('Click to reload rows that need updating'
                            if needs_reload else 'All rows up to date')
        if fd.get('source_type') == 'word':
            real_path = fd.get('real_path', path)
            self._update_word_group_reload_indicator(real_path)


# ── ENTRY POINT ──


def main():
    """Launch the pyTable UI."""
    if doc is None:
        _alert(
            'pyTable needs an open Revit project to run.',
            title='pyTable — Not Available')
        return
    if getattr(doc, 'IsFamilyDocument', False) or doc.ProjectInformation is None:
        _alert(
            'pyTable isn\'t available in Family documents (.rfa) — it '
            'needs a Project document (.rvt) to store its state and '
            'create Schedule/Legend/Drafting views.\n\n'
            'Open a project and try again.',
            title='pyTable — Not Supported Here')
        return
    window = PyTableWindow()
    window.show_dialog()


if __name__ == '__main__':
    main()
