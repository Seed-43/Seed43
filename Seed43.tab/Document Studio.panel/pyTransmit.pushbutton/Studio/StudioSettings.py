# -*- coding: utf-8 -*-
# StudioSettings.py
#
# pyTransmit Studio - an Excel/LibreOffice-Calc-style alternative to the
# Layout Builder (../Layout/LayoutSettings.py, untouched by this file): a
# free grid (any rows/columns), drag-select + merge/unmerge, a formatting
# ribbon, a formula bar showing the block in the active cell, and a preview
# fed by live Revit data instead of dummy placeholders.

import os

# pytransmit_paths lives in the pushbutton root, which is not guaranteed to be
# on sys.path - pyTransmit loads this module by inserting only its own folder.
import sys as _sys
_PT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PT_ROOT not in _sys.path:
    _sys.path.insert(0, _PT_ROOT)
from pytransmit_paths import SETTINGS_DIR, STUDIO_LAYOUTS_DIR, STUDIO_CONFIG
import json
import string

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System.Windows as _SW
import System.Windows.Controls as _SWC
import System.Windows.Media as _SWM
import System.Windows.Input as _SWI
import System.Windows.Shapes as _SWS
import System.Windows.Forms as _WF
import System.Drawing as _SD

from pyrevit.forms import WPFWindow

try:
    from Snippets.seed43_theme import (apply_seed43_palette, apply_seed43_dimensions,
                                       get_color as _get_color)
except Exception:
    apply_seed43_palette = None
    apply_seed43_dimensions = None
    _get_color = None

try:
    from Snippets import _dialogs as sdlg
except Exception:
    sdlg = None

import studio_grid
import studio_blocks
import studio_live_data
from format_cells_dialog import open_format_cells_dialog


def _alert(message, title='pyTransmit Studio'):
    if sdlg:
        try:
            sdlg.message(message, title=title)
            return
        except Exception:
            pass
    _WF.MessageBox.Show(message, title)


def _confirm(message, title='pyTransmit Studio'):
    if sdlg:
        try:
            return sdlg.confirm(message, title=title, no='Cancel')
        except Exception:
            pass
    return _WF.MessageBox.Show(message, title, _WF.MessageBoxButtons.YesNo) == _WF.DialogResult.Yes


def _col_letter(c):
    """0-indexed column -> Excel-style letters (0->A, 25->Z, 26->AA...)."""
    s = ''
    c += 1
    while c > 0:
        c, rem = divmod(c - 1, 26)
        s = string.ascii_uppercase[rem] + s
    return s


def _cell_ref(r, c):
    return '{}{}'.format(_col_letter(c), r + 1)


SCALE = 3.2          # px per mm for the grid canvas
HEADER_H = 22
HEADER_W = 40
HEADER_RESIZE_MARGIN = 5   # px from a header's far edge that starts a drag-resize
# px band around the selection's outline where a drag moves the content
# instead of extending the selection - same distinction Excel makes.
MOVE_EDGE_MARGIN = 5

# Ctrl + mouse wheel zoom, same range/step as asked for: 25% out, 200% in,
# 100% normal, 5% per wheel notch.
ZOOM_MIN = 0.25
ZOOM_MAX = 2.00
ZOOM_STEP = 0.05
ZOOM_DEFAULT = 1.00


def _brush(hexcolor):
    h = (hexcolor or '#000000').lstrip('#')
    if len(h) == 3:
        h = h[0] * 2 + h[1] * 2 + h[2] * 2
    return _SWM.SolidColorBrush(_SWM.Color.FromRgb(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)))


# -- Themed context menus ----------------------------------------------------
# Background/Foreground only go so far: MenuItem's stock template draws a
# fixed icon-and-checkmark gutter down the left that stays light whatever
# colours are set. That strip is native chrome, so removing it needs a real
# ControlTemplate - same reason the ComboBox needed one.
_MENU_XAML = (
    '<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" '
    'xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" TargetType="ContextMenu">'
    '<Border Background="#2B3340" BorderBrush="#404E62" BorderThickness="1" '
    'CornerRadius="4" Padding="4">'
    '<StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Cycle"/>'
    '</Border>'
    '</ControlTemplate>'
)

_MENUITEM_XAML = (
    '<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" '
    'xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" TargetType="MenuItem">'
    '<Border x:Name="Bd" Background="Transparent" CornerRadius="3" Padding="12,5">'
    '<ContentPresenter ContentSource="Header" VerticalAlignment="Center"/>'
    '</Border>'
    '<ControlTemplate.Triggers>'
    '<Trigger Property="IsHighlighted" Value="True">'
    '<Setter TargetName="Bd" Property="Background" Value="#3A4760"/>'
    '</Trigger>'
    '<Trigger Property="IsEnabled" Value="False">'
    '<Setter Property="Opacity" Value="0.45"/>'
    '</Trigger>'
    '</ControlTemplate.Triggers>'
    '</ControlTemplate>'
)


def _themed_menu():
    menu = _SWC.ContextMenu()
    menu.Foreground = _SWM.Brushes.White
    try:
        import System.Windows.Markup as _Markup
        menu.Template = _Markup.XamlReader.Parse(_MENU_XAML)
    except Exception:
        # Fall back to plain colours if template parsing ever fails - the
        # menu still works, it just keeps the stock gutter.
        menu.Background = _brush('#2B3340')
        menu.BorderBrush = _brush('#404E62')
    return menu


def _themed_menu_item(text, handler, dot_color=None):
    """dot_color: a hex string draws a filled colour dot before the label,
    'outline' draws a hollow one. Used so the row-section menu items carry
    the same colours as the gutter stripes and the palette legend."""
    mi = _SWC.MenuItem()
    if dot_color:
        row = _SWC.StackPanel()
        row.Orientation = _SWC.Orientation.Horizontal
        dot = _SWC.Border()
        dot.Width = 9
        dot.Height = 9
        dot.CornerRadius = _SW.CornerRadius(5)
        dot.Margin = _SW.Thickness(0, 0, 8, 0)
        dot.VerticalAlignment = _SW.VerticalAlignment.Center
        if dot_color == 'outline':
            dot.Background = _SWM.Brushes.Transparent
            dot.BorderBrush = _brush('#8A96A8')
            dot.BorderThickness = _SW.Thickness(1)
        else:
            dot.Background = _brush(dot_color)
        lbl = _SWC.TextBlock()
        lbl.Text = text
        lbl.Foreground = _SWM.Brushes.White
        lbl.VerticalAlignment = _SW.VerticalAlignment.Center
        row.Children.Add(dot)
        row.Children.Add(lbl)
        mi.Header = row
    else:
        mi.Header = text
    mi.Foreground = _SWM.Brushes.White
    mi.FontSize = 12
    try:
        import System.Windows.Markup as _Markup
        mi.Template = _Markup.XamlReader.Parse(_MENUITEM_XAML)
    except Exception:
        pass
    if handler is not None:
        mi.Click += handler
    return mi


def _themed_separator():
    sep = _SWC.Separator()
    sep.Background = _brush('#404E62')
    sep.Height = 1
    sep.Margin = _SW.Thickness(6, 4, 6, 4)
    return sep


class StudioSettingsWindow(WPFWindow):

    def __init__(self, script_dir=None):
        if script_dir is None:
            script_dir = os.path.dirname(os.path.abspath(__file__))
        self._script_dir = script_dir
        self._layouts_dir = STUDIO_LAYOUTS_DIR
        self._settings_dir = SETTINGS_DIR
        if not os.path.isdir(self._layouts_dir):
            try:
                os.makedirs(self._layouts_dir)
            except Exception:
                pass

        xaml = os.path.join(script_dir, 'StudioSettings.xaml')
        WPFWindow.__init__(self, xaml)
        self._fit_to_screen()
        if apply_seed43_palette:
            try:
                apply_seed43_palette(self, script_dir)
            except Exception:
                pass
        if apply_seed43_dimensions:
            try:
                apply_seed43_dimensions(self, script_dir)
            except Exception:
                pass

        # -- state ---------------------------------------------------------
        self._active_path = None
        self._sel = (0, 0, 0, 0)        # anchor_r, anchor_c, active_r, active_c
        # How the selection was made. Merge-expansion applies to cell
        # selections only - picking a column header means "this column",
        # even if merged cells span across it.
        self._sel_mode = 'cell'          # 'cell' | 'row' | 'col'
        self._header_drag = None         # 'colhdr'/'rowhdr' while dragging across headers
        self._dragging = False
        self._resizing = None    # None, or ('col', idx) / ('row', idx) while drag-resizing a header
        self._row_px = []               # cumulative row boundaries (content-only, header excluded)
        self._col_px = []

        self._grid = self._load_last_or_default()
        self._data = studio_live_data.empty_data()
        self._refresh_live_data(silent=True)

        self._blocks_expanded = True
        self._block_drag_start = None
        self._block_drag_btn = None
        self._zoom = ZOOM_DEFAULT
        self._pending_move = None   # set on mouse-down over the selection edge
        self._move_start = None
        # Set while the ribbon is being updated to match the selected cell,
        # so the combos' own SelectionChanged handlers don't turn a passive
        # UI refresh into an edit and recurse back through _render_all().
        self._syncing_ui = False

        self._build_blocks_panel()
        self._wire_home_ribbon()
        self._wire_header_buttons()
        self._wire_blocks_panel()
        self._wire_templates()

        # Window-level: accept the drag everywhere so it's never cancelled
        # before reaching a cell's own drop handler (same fix the Layout
        # Builder needed for its palette drag - WPF otherwise shows "no drop"
        # over any gap between drop targets).
        self.AllowDrop = True
        self.DragOver += self._window_drag_over
        self.KeyDown += self._on_window_key_down
        # Preview- so it runs before the ScrollViewer consumes the wheel for
        # scrolling; without Ctrl held we leave it alone and normal scrolling
        # still works.
        self.grid_scroll.PreviewMouseWheel += self._on_grid_mouse_wheel
        # Keep the frozen header panes lined up with the cell area.
        self.grid_scroll.ScrollChanged += self._on_grid_scroll_changed
        # A ScrollViewer with a hidden scrollbar still scrolls on the mouse
        # wheel, so scrolling over the headers used to move them out of step
        # with the cells. Forward the wheel to the cell pane instead, which
        # then re-syncs both headers through ScrollChanged.
        self.rowhdr_scroll.PreviewMouseWheel += self._on_header_wheel
        self.colhdr_scroll.PreviewMouseWheel += self._on_header_wheel
        self.zoom_in_btn.Click += lambda s, a: self._set_zoom(self._zoom + ZOOM_STEP)
        self.zoom_out_btn.Click += lambda s, a: self._set_zoom(self._zoom - ZOOM_STEP)
        self.zoom_reset_btn.Click += lambda s, a: self._set_zoom(ZOOM_DEFAULT)

        self._render_all()
        self._update_zoom_label()

    def _on_grid_scroll_changed(self, sender, args):
        self._sync_header_scroll()

    def _sync_header_scroll(self):
        try:
            self.colhdr_scroll.ScrollToHorizontalOffset(self.grid_scroll.HorizontalOffset)
            self.rowhdr_scroll.ScrollToVerticalOffset(self.grid_scroll.VerticalOffset)
        except Exception:
            pass

    def _on_header_wheel(self, sender, args):
        """Wheel over a frozen header scrolls the cell pane, never the header
        on its own - otherwise the row numbers drift out of alignment with
        the rows they label."""
        mods = _SWI.Keyboard.Modifiers
        if (mods & _SWI.ModifierKeys.Control) == _SWI.ModifierKeys.Control:
            self._on_grid_mouse_wheel(sender, args)
            return
        try:
            self.grid_scroll.ScrollToVerticalOffset(
                self.grid_scroll.VerticalOffset - args.Delta)
        except Exception:
            pass
        args.Handled = True

    # ======================================================================
    # Zoom (Ctrl + mouse wheel)
    # ======================================================================

    def _on_grid_mouse_wheel(self, sender, args):
        mods = _SWI.Keyboard.Modifiers
        if (mods & _SWI.ModifierKeys.Control) != _SWI.ModifierKeys.Control:
            return
        step = ZOOM_STEP if args.Delta > 0 else -ZOOM_STEP
        self._set_zoom(self._zoom + step)
        args.Handled = True

    def _set_zoom(self, zoom):
        # Round to the step so repeated wheel notches can't accumulate float
        # drift (0.9999... instead of a clean 100%).
        zoom = round(zoom / ZOOM_STEP) * ZOOM_STEP
        zoom = max(ZOOM_MIN, min(ZOOM_MAX, zoom))
        if abs(zoom - self._zoom) < 1e-9:
            return
        self._zoom = zoom
        self._apply_zoom()
        self._update_zoom_label()

    def _update_zoom_label(self):
        try:
            self.zoom_label.Text = '{}%'.format(int(round(self._zoom * 100)))
        except Exception:
            pass

    def _apply_zoom(self):
        """LayoutTransform (not RenderTransform) so the ScrollViewers see the
        scaled size and their scrollbars stay correct. Applied to all three
        panes so the frozen headers scale with the cells and stay aligned.

        Mouse maths elsewhere uses GetPosition(<the scaled grid>), which
        reports that grid's own untransformed coordinates, so hit-testing
        needs no zoom correction.
        """
        at_default = abs(self._zoom - 1.0) < 1e-9
        for host in (self.grid_root, self.colhdr_root, self.rowhdr_root):
            try:
                # A fresh Transform per host rather than one shared instance -
                # avoids any question of Freezable ownership across elements.
                host.LayoutTransform = None if at_default else \
                    _SWM.ScaleTransform(self._zoom, self._zoom)
            except Exception:
                pass

    def _fit_to_screen(self):
        """Clamp the window to the screen's usable area.

        The XAML's default size is a preference, not a guarantee - on a
        smaller display (e.g. a 1280x800 laptop) a taller default pushes the
        bottom of the window off the screen entirely, taking the status bar
        and its zoom controls with it. The window is borderless with a
        custom chrome, so there's no OS titlebar to drag it back into view
        either. Clamp against SystemParameters.WorkArea, which already
        excludes the taskbar.
        """
        try:
            work = _SW.SystemParameters.WorkArea
            margin = 40
            max_h = work.Height - margin
            max_w = work.Width - margin
            if self.Height > max_h:
                self.Height = max(self.MinHeight, max_h)
            if self.Width > max_w:
                self.Width = max(self.MinWidth, max_w)
        except Exception:
            pass

    def _window_drag_over(self, sender, args):
        try:
            args.Effects = _SW.DragDropEffects.Copy
            args.Handled = True
        except Exception:
            pass

    def _on_window_key_down(self, sender, args):
        if args.Key != _SWI.Key.Delete:
            return
        # Don't hijack Delete while the user is actually editing text
        # somewhere (formula bar, Row H/Col W boxes, the font combos) -
        # only treat it as "clear the selected cell(s)" when focus is on
        # the grid itself.
        focused = _SWI.Keyboard.FocusedElement
        if isinstance(focused, (_SWC.TextBox, _SWC.ComboBox)):
            return
        for (r, c) in self._sel_origins():
            self._grid.clear_cell(r, c)
        self._render_all()
        args.Handled = True

    # ======================================================================
    # Persistence
    # ======================================================================

    def _last_file_marker(self):
        return STUDIO_CONFIG

    def _load_last_or_default(self):
        marker = self._last_file_marker()
        try:
            with open(marker, 'r') as f:
                cfg = json.load(f)
            last = cfg.get('last_file')
            if last and os.path.isfile(last):
                self._active_path = last
                return studio_grid.load_layout(last)
        except Exception:
            pass
        return studio_grid.default_grid()

    def _remember_last_file(self, path):
        try:
            with open(self._last_file_marker(), 'w') as f:
                json.dump({'last_file': path}, f)
        except Exception:
            pass

    def new_click(self, sender, args):
        if not _confirm('Start a new blank layout? Unsaved changes will be lost.'):
            return
        self._active_path = None
        self._grid = studio_grid.default_grid()
        self._sel = (0, 0, 0, 0)
        self._render_all()

    # -- Templates (named layouts in studio_layouts/) --------------------------
    def _template_names(self):
        try:
            return sorted(os.path.splitext(f)[0] for f in os.listdir(self._layouts_dir)
                          if f.lower().endswith('.json'))
        except Exception:
            return []

    def _template_path(self, name):
        safe = str(name).replace('/', '_').replace('\\', '_').replace(':', '_').strip()
        return os.path.join(self._layouts_dir, safe + '.json')

    def _wire_templates(self):
        self.template_combo.SelectionChanged += self._on_template_changed
        self.tmpl_delete_btn.Click += self._on_template_delete
        self._reload_template_list()

    def _reload_template_list(self, select=None):
        """Refill the dropdown. Guarded by _syncing_ui so repopulating it
        doesn't fire SelectionChanged and load a template as a side effect."""
        self._syncing_ui = True
        try:
            self.template_combo.Items.Clear()
            for name in self._template_names():
                self.template_combo.Items.Add(name)
            target = select if select is not None else self._active_name()
            if target and target in list(self.template_combo.Items):
                self.template_combo.SelectedItem = target
        except Exception:
            pass
        finally:
            self._syncing_ui = False

    def _active_name(self):
        if not self._active_path:
            return None
        return os.path.splitext(os.path.basename(self._active_path))[0]

    def _on_template_changed(self, sender, args):
        if self._syncing_ui:
            return
        name = self.template_combo.SelectedItem
        if not name:
            return
        path = self._template_path(name)
        if not os.path.isfile(path):
            return
        try:
            self._grid = studio_grid.load_layout(path)
            self._active_path = path
            self._sel = (0, 0, 0, 0)
            self._remember_last_file(path)
            self._render_all()
            self.status_label.Text = 'Loaded template "{}"'.format(name)
        except Exception as e:
            _alert('Could not load template "{}":\n{}'.format(name, e))

    def _on_template_delete(self, sender, args):
        name = self.template_combo.SelectedItem
        if not name:
            return
        if not _confirm('Delete the template "{}"? This cannot be undone.'.format(name)):
            return
        try:
            os.remove(self._template_path(name))
        except Exception as e:
            _alert('Could not delete template:\n{}'.format(e))
            return
        if self._active_name() == name:
            self._active_path = None
        self._reload_template_list()
        self.status_label.Text = 'Deleted template "{}"'.format(name)
        self._render_all()

    def save_click(self, sender, args):
        if not self._active_path:
            self.save_as_click(sender, args)
            return
        self._do_save(self._active_path)
        self._reload_template_list()

    def save_as_click(self, sender, args):
        name = None
        if sdlg:
            try:
                name = sdlg.ask_string('Template name:', title='Save Template As',
                                       default=self._active_name() or 'Untitled')
            except Exception:
                name = None
        if name is None:
            dlg = _WF.SaveFileDialog()
            dlg.InitialDirectory = self._layouts_dir
            dlg.Filter = 'pyTransmit Studio layouts (*.json)|*.json'
            dlg.FileName = (self._active_name() or 'Untitled') + '.json'
            if dlg.ShowDialog() != _WF.DialogResult.OK:
                return
            path = dlg.FileName
        else:
            name = name.strip()
            if not name:
                return
            path = self._template_path(name)
            if os.path.isfile(path) and not _confirm(
                    'Template "{}" already exists. Overwrite it?'.format(name)):
                return
        self._do_save(path)
        self._reload_template_list(select=os.path.splitext(os.path.basename(path))[0])

    def _do_save(self, path):
        try:
            name = os.path.splitext(os.path.basename(path))[0]
            studio_grid.save_layout(path, self._grid, name)
            self._active_path = path
            self._remember_last_file(path)
            self.active_file_label.Text = '- {}'.format(os.path.basename(path))
            self.status_label.Text = 'Saved template "{}"'.format(
                os.path.splitext(os.path.basename(path))[0])
        except Exception as e:
            _alert('Could not save layout:\n{}'.format(e))

    def close_click(self, sender, args):
        self.Close()

    def refresh_click(self, sender, args):
        self._refresh_live_data(silent=False)
        self._render_all()

    def _refresh_live_data(self, silent=True):
        try:
            self._data = studio_live_data.get_live_data(self._settings_dir)
            if not silent:
                n_rev = len(self._data.get('revisions', []))
                self.status_label.Text = 'Live data refreshed - {} issued revision(s)'.format(n_rev)
        except Exception as e:
            self._data = studio_live_data.empty_data()
            if not silent:
                self.status_label.Text = 'Could not read live Revit data: {}'.format(e)

    def _wire_header_buttons(self):
        pass  # Click handlers wired via XAML inline Click attrs

    # ======================================================================
    # Blocks panel (collapsible, right-hand side)
    # ======================================================================

    def _wire_blocks_panel(self):
        self.blocks_collapse_btn.Click += self._on_blocks_collapse_click
        self.blocks_collapse_btn.IsChecked = False

    def _on_blocks_collapse_click(self, sender, args):
        self._blocks_expanded = not self.blocks_collapse_btn.IsChecked
        if self._blocks_expanded:
            self.blocks_col.Width = _SW.GridLength(230)
            self.blocks_scroll.Visibility = _SW.Visibility.Visible
            self.blocks_title_label.Visibility = _SW.Visibility.Visible
            self.blocks_collapse_btn.Content = '»'  # »
        else:
            self.blocks_col.Width = _SW.GridLength(28)
            self.blocks_scroll.Visibility = _SW.Visibility.Collapsed
            self.blocks_title_label.Visibility = _SW.Visibility.Collapsed
            self.blocks_collapse_btn.Content = '«'  # «

    def _build_section_legend(self, panel):
        """Colour key for the row-section stripes, so the gutter colours
        aren't something the user has to decode by trial and error."""
        box = _SWC.StackPanel()
        box.Margin = _SW.Thickness(0, 0, 0, 12)
        hdr = _SWC.TextBlock()
        hdr.Text = 'ROW SECTIONS'
        hdr.FontSize = 9
        hdr.Foreground = _brush('#8A96A8')
        hdr.Margin = _SW.Thickness(2, 0, 0, 4)
        box.Children.Add(hdr)

        entries = [
            (studio_grid.SECTION_REPEAT_TOP,
             studio_grid.SECTION_LABELS[studio_grid.SECTION_REPEAT_TOP]),
            (studio_grid.SECTION_REPEAT_BOTTOM,
             studio_grid.SECTION_LABELS[studio_grid.SECTION_REPEAT_BOTTOM]),
            (studio_grid.SECTION_BODY,
             studio_grid.SECTION_LABELS[studio_grid.SECTION_BODY]),
        ]
        for key, text in entries:
            row = _SWC.StackPanel()
            row.Orientation = _SWC.Orientation.Horizontal
            row.Margin = _SW.Thickness(2, 0, 0, 3)
            swatch = _SWC.Border()
            swatch.Width = 10
            swatch.Height = 10
            swatch.CornerRadius = _SW.CornerRadius(2)
            swatch.Margin = _SW.Thickness(0, 0, 6, 0)
            swatch.VerticalAlignment = _SW.VerticalAlignment.Center
            colour = studio_grid.SECTION_COLORS.get(key)
            if colour:
                swatch.Background = _brush(colour)
            else:
                swatch.Background = _SWM.Brushes.Transparent
                swatch.BorderBrush = _brush('#55606F')
                swatch.BorderThickness = _SW.Thickness(1)
            lbl = _SWC.TextBlock()
            lbl.Text = text
            lbl.FontSize = 10
            lbl.TextWrapping = _SW.TextWrapping.Wrap
            lbl.Foreground = _brush('#C3CCD8')
            lbl.VerticalAlignment = _SW.VerticalAlignment.Center
            row.Children.Add(swatch)
            row.Children.Add(lbl)
            box.Children.Add(row)

        hint = _SWC.TextBlock()
        hint.Text = 'Right-click a row number to set.'
        hint.FontSize = 9
        hint.FontStyle = _SW.FontStyles.Italic
        hint.TextWrapping = _SW.TextWrapping.Wrap
        hint.Foreground = _brush('#7C8899')
        hint.Margin = _SW.Thickness(2, 2, 0, 0)
        box.Children.Add(hint)

        sep = _SWC.Border()
        sep.Height = 1
        sep.Background = _brush('#404E62')
        sep.Margin = _SW.Thickness(0, 8, 0, 0)
        box.Children.Add(sep)
        panel.Children.Add(box)

    def _build_blocks_panel(self):
        panel = self.blocks_panel_content
        panel.Children.Clear()
        self._build_section_legend(panel)
        group_box = None
        group_wrap = None
        built_group = None   # plain Python variable - WPF controls (.NET objects)
                              # can't carry extra Python attributes like _group
        for entry in studio_blocks.PALETTE:
            t, label, icon, group, subtext = entry
            if t == '__grp__':
                continue
            if group != built_group:
                built_group = group
                group_box = _SWC.StackPanel()
                group_box.Orientation = _SWC.Orientation.Vertical
                group_box.Margin = _SW.Thickness(0, 0, 0, 10)
                hdr = _SWC.TextBlock()
                hdr.Text = studio_blocks.GROUP_LABELS.get(group, group.upper())
                hdr.FontSize = 9
                hdr.Foreground = _brush(studio_blocks.GROUP_COLORS.get(group, '#8A96A8'))
                hdr.Margin = _SW.Thickness(2, 0, 0, 4)
                group_box.Children.Add(hdr)
                group_wrap = _SWC.WrapPanel()
                group_box.Children.Add(group_wrap)
                panel.Children.Add(group_box)
            btn = _SWC.Button()
            btn.Content = label
            btn.Tag = t
            btn.Padding = _SW.Thickness(8, 4, 8, 4)
            btn.Margin = _SW.Thickness(0, 0, 4, 4)
            btn.FontSize = 11
            try:
                btn.Background = _brush(studio_blocks.GROUP_COLORS.get(group, '#8A96A8'))
            except Exception:
                pass
            btn.Foreground = _SWM.Brushes.White
            btn.Cursor = _SWI.Cursors.Hand
            # Both a plain click (inserts into the currently selected cell)
            # and a drag onto any cell are supported - PreviewMouseMove
            # decides which one it turns into, same threshold-based pattern
            # the Layout Builder's own palette uses.
            btn.PreviewMouseLeftButtonDown += self._on_block_btn_down
            btn.PreviewMouseMove += self._on_block_btn_move
            btn.PreviewMouseLeftButtonUp += self._on_block_btn_up
            group_wrap.Children.Add(btn)

    def _on_block_btn_down(self, sender, args):
        self._block_drag_start = args.GetPosition(sender)
        self._block_drag_btn = sender

    def _on_block_btn_move(self, sender, args):
        if args.LeftButton != _SWI.MouseButtonState.Pressed:
            return
        start = getattr(self, '_block_drag_start', None)
        if start is None or getattr(self, '_block_drag_btn', None) is not sender:
            return
        pos = args.GetPosition(sender)
        if abs(pos.X - start.X) < 4 and abs(pos.Y - start.Y) < 4:
            return
        block_type = sender.Tag
        self._block_drag_start = None
        self._block_drag_btn = None
        data = _SW.DataObject('studio_block_type', block_type)
        try:
            _SW.DragDrop.DoDragDrop(sender, data, _SW.DragDropEffects.Copy)
        except Exception:
            pass

    def _on_block_btn_up(self, sender, args):
        # Only reached if PreviewMouseMove never crossed the drag threshold -
        # i.e. this was a plain click, not a drag. Insert into the active cell.
        if getattr(self, '_block_drag_start', None) is None:
            return
        self._block_drag_start = None
        self._block_drag_btn = None
        block_type = sender.Tag
        ar, ac, _, _ = self._sel
        self._grid.set_block(ar, ac, studio_blocks.new_block_from_palette(block_type))
        self._render_all()

    # ======================================================================
    # Home ribbon wiring
    # ======================================================================

    def _wire_home_ribbon(self):
        self.page_size_combo.Items.Clear()
        for name in studio_grid.PAGE_SIZES:
            self.page_size_combo.Items.Add(name)
        self.page_size_combo.SelectedItem = self._grid.page_size_name
        self.page_size_combo.SelectionChanged += self._on_page_size_changed

        self.page_orientation_combo.Items.Clear()
        self.page_orientation_combo.Items.Add('Portrait')
        self.page_orientation_combo.Items.Add('Landscape')
        self.page_orientation_combo.SelectedItem = self._grid.orientation.capitalize()
        self.page_orientation_combo.SelectionChanged += self._on_page_orientation_changed

        self.font_family_combo.Items.Clear()
        for fam in ['Arial', 'Calibri', 'Segoe UI', 'Times New Roman', 'Verdana']:
            self.font_family_combo.Items.Add(fam)
        self.font_family_combo.SelectedIndex = 0
        self.font_family_combo.SelectionChanged += self._on_font_family_changed

        # Points, like Excel - the model stores mm, the UI converts.
        self.font_size_combo.Items.Clear()
        for sz in studio_blocks.FONT_SIZES_PT:
            self.font_size_combo.Items.Add(str(sz))
        self.font_size_combo.SelectedItem = '9'
        self.font_size_combo.SelectionChanged += self._on_font_size_changed

        self.bold_btn.Click += self._on_bold_click
        self.italic_btn.Click += self._on_italic_click
        self.underline_btn.Click += self._on_underline_click
        self.font_color_btn.Click += self._on_font_color_click
        self.fill_color_btn.Click += self._on_fill_color_click

        self.align_left_btn.Click += lambda s, a: self._on_align_click('left')
        self.align_center_btn.Click += lambda s, a: self._on_align_click('center')
        self.align_right_btn.Click += lambda s, a: self._on_align_click('right')
        self.valign_top_btn.Click += lambda s, a: self._on_valign_click('top')
        self.valign_mid_btn.Click += lambda s, a: self._on_valign_click('middle')
        self.valign_bot_btn.Click += lambda s, a: self._on_valign_click('bottom')

        for _btn_name, _kind in (('align_left_btn', 'left'), ('align_center_btn', 'center'),
                                 ('align_right_btn', 'right'), ('valign_top_btn', 'top'),
                                 ('valign_mid_btn', 'middle'), ('valign_bot_btn', 'bottom')):
            try:
                getattr(self, _btn_name).Content = self._build_align_icon(_kind)
            except Exception:
                pass

        self._build_borders_popup()
        self.borders_btn.Checked += lambda s, a: setattr(self.borders_popup, 'IsOpen', True)
        self.borders_btn.Unchecked += lambda s, a: setattr(self.borders_popup, 'IsOpen', False)
        self.borders_popup.Closed += lambda s, a: setattr(self.borders_btn, 'IsChecked', False)

        self.merge_toggle_btn.Click += self._on_merge_toggle_click

        self.fit_width_btn.Click += self._on_fit_width_click
        self.distribute_btn.Click += self._on_distribute_click
        self.margin_box.LostFocus += self._on_margin_commit
        self.margin_box.Text = str(self._grid.margin_mm)

        self.row_h_box.LostFocus += self._on_row_h_commit
        self.col_w_box.LostFocus += self._on_col_w_commit

        self.formula_bar_tb.LostFocus += self._on_formula_commit
        self.formula_bar_tb.KeyDown += self._on_formula_keydown
        self.formula_change_btn.Click += self._on_formula_change_click

    # -- Borders: Excel/Calc-style icon-grid dropdown --------------------------
    BORDER_PRESETS = [
        ('none', 'No Border'), ('all', 'All Borders'), ('outside', 'Outside Borders'),
        ('top', 'Top Border'), ('bottom', 'Bottom Border'), ('left', 'Left Border'),
        ('right', 'Right Border'), ('top_bottom', 'Top && Bottom'), ('left_right', 'Left && Right'),
    ]

    def _build_align_icon(self, kind):
        """Excel-style alignment glyph drawn from bars, rather than a unicode
        character - the arrow/box glyphs previously used aren't present in
        every UI font and were rendering as the wrong symbol entirely."""
        bar_colour = _brush('#F4FAFF')
        panel = _SWC.StackPanel()
        panel.Width = 16
        panel.Height = 14
        horizontal = kind in ('left', 'center', 'right')
        widths = [14, 9, 14, 9] if horizontal else [14, 14, 14]
        align_map = {
            'left': _SW.HorizontalAlignment.Left,
            'center': _SW.HorizontalAlignment.Center,
            'right': _SW.HorizontalAlignment.Right,
        }
        if not horizontal:
            panel.VerticalAlignment = {
                'top': _SW.VerticalAlignment.Top,
                'middle': _SW.VerticalAlignment.Center,
                'bottom': _SW.VerticalAlignment.Bottom,
            }.get(kind, _SW.VerticalAlignment.Center)
            panel.Height = 9
        for w in widths:
            bar = _SWS.Rectangle()
            bar.Height = 2
            bar.Width = w
            bar.Fill = bar_colour
            bar.Margin = _SW.Thickness(0, 1, 0, 1)
            if horizontal:
                bar.HorizontalAlignment = align_map.get(kind, _SW.HorizontalAlignment.Left)
            else:
                bar.HorizontalAlignment = _SW.HorizontalAlignment.Center
            panel.Children.Add(bar)
        # Vertical-align icons show the bar group pinned top/middle/bottom
        # inside a fixed-height box, which is what conveys the alignment.
        if not horizontal:
            host = _SWC.Grid()
            host.Width = 16
            host.Height = 14
            host.Children.Add(panel)
            return host
        return panel

    def _accent_hex(self):
        """Live accent colour straight from the palette JSON (falls back to
        the Seed43 green). Used for the border-preset icons so the edges
        being previewed read clearly against the dark popup instead of
        disappearing as thin black-on-grey lines."""
        if _get_color:
            try:
                return _get_color(self._script_dir, 'accent', fallback='#20A344')
            except Exception:
                pass
        return '#20A344'

    def _build_border_icon(self, preset):
        size = 26
        accent = _brush(self._accent_hex())
        g = _SWC.Grid()
        g.Width = size
        g.Height = size
        bg = _SWC.Border()
        bg.Background = _brush('#1E2530')
        bg.BorderBrush = _brush('#55606F')
        bg.BorderThickness = _SW.Thickness(1)
        g.Children.Add(bg)
        edge_map = {
            'top': (0, 1, 0, 0), 'bottom': (0, 0, 0, 1), 'left': (1, 0, 0, 0), 'right': (0, 0, 1, 0),
            'top_bottom': (0, 1, 0, 1), 'left_right': (1, 0, 1, 0),
            'outside': (1, 1, 1, 1), 'all': (1, 1, 1, 1), 'none': (0, 0, 0, 0),
        }
        l, t, r, b = edge_map.get(preset, (0, 0, 0, 0))
        fg = _SWC.Border()
        fg.BorderBrush = accent
        fg.BorderThickness = _SW.Thickness(l * 2, t * 2, r * 2, b * 2)
        g.Children.Add(fg)
        if preset == 'all':
            vline = _SWS.Rectangle()
            vline.Fill = accent
            vline.Width = 1
            vline.HorizontalAlignment = _SW.HorizontalAlignment.Center
            vline.VerticalAlignment = _SW.VerticalAlignment.Stretch
            vline.Margin = _SW.Thickness(0, 4, 0, 4)
            g.Children.Add(vline)
            hline = _SWS.Rectangle()
            hline.Fill = accent
            hline.Height = 1
            hline.VerticalAlignment = _SW.VerticalAlignment.Center
            hline.HorizontalAlignment = _SW.HorizontalAlignment.Stretch
            hline.Margin = _SW.Thickness(4, 0, 4, 0)
            g.Children.Add(hline)
        return g

    def _build_borders_popup(self):
        container = self.borders_grid
        container.Children.Clear()
        seen = set()
        for key, label in self.BORDER_PRESETS:
            if key in seen:
                continue
            seen.add(key)
            btn = _SWC.Button()
            btn.Tag = key
            btn.Margin = _SW.Thickness(3)
            btn.Padding = _SW.Thickness(3)
            btn.ToolTip = label
            btn.Background = _SWM.Brushes.Transparent
            btn.BorderThickness = _SW.Thickness(0)
            btn.Content = self._build_border_icon(key)
            btn.Cursor = _SWI.Cursors.Hand
            btn.Click += self._on_border_preset_click
            container.Children.Add(btn)

    def _on_border_preset_click(self, sender, args):
        preset = sender.Tag
        self.borders_popup.IsOpen = False
        self.borders_btn.IsChecked = False
        self._apply_border_preset(preset)

    def _apply_border_preset(self, preset):
        """Mirrors Excel's border presets: 'all'/'none' touch every cell in
        the selection; the edge-specific presets ('top'/'outside'/etc.) only
        touch the cells that actually sit on that edge of the selection
        rectangle, not every cell - e.g. 'Top Border' draws one line along
        the top of the whole selection, not a line under every cell."""
        r0, c0, r1, c1 = self._sel_rect()
        for (r, c) in self._sel_origins():
            block = self._grid.block_at(r, c)
            if block is None:
                block = studio_blocks.new_block('text', content='')
                self._grid.set_block(r, c, block)
            row_span, col_span = self._grid.span_of(r, c)
            r_end, c_end = r + row_span - 1, c + col_span - 1
            is_top, is_bottom = (r <= r0), (r_end >= r1)
            is_left, is_right = (c <= c0), (c_end >= c1)
            borders = dict(block.get('borders') or {})
            if preset == 'none':
                borders = {'t': False, 'b': False, 'l': False, 'r': False}
            elif preset == 'all':
                borders = {'t': True, 'b': True, 'l': True, 'r': True}
            else:
                if preset in ('outside', 'top', 'top_bottom') and is_top: borders['t'] = True
                if preset in ('outside', 'bottom', 'top_bottom') and is_bottom: borders['b'] = True
                if preset in ('outside', 'left', 'left_right') and is_left: borders['l'] = True
                if preset in ('outside', 'right', 'left_right') and is_right: borders['r'] = True
            block['borders'] = borders
        self._render_all()

    # -- selection helpers -----------------------------------------------------
    def _sel_rect(self):
        """The selected rectangle, expanded so it always fully contains any
        merged cell it touches - a merged cell behaves as one cell, so a
        selection can never cover just part of one. Excel does the same:
        touch any part of a merge and the whole thing comes with it.

        Expansion repeats until stable, because growing the rectangle to
        swallow one merge can bring it into contact with another.
        """
        ar, ac, br, bc = self._sel
        r0, c0 = min(ar, br), min(ac, bc)
        r1, c1 = max(ar, br), max(ac, bc)

        # Row/column header selections are taken literally. Expanding them
        # would be useless in practice: one full-width merge anywhere in the
        # sheet would drag every column into the selection, making it
        # impossible to select or resize a single column.
        if getattr(self, '_sel_mode', 'cell') != 'cell':
            return (r0, c0, r1, c1)

        for _ in range(64):   # bounded; a stable result normally lands in 1-2 passes
            n_r0, n_c0, n_r1, n_c1 = r0, c0, r1, c1
            for r in range(r0, r1 + 1):
                for c in range(c0, c1 + 1):
                    orig_r, orig_c = self._grid.origin_of(r, c)
                    row_span, col_span = self._grid.span_of(orig_r, orig_c)
                    n_r0 = min(n_r0, orig_r)
                    n_c0 = min(n_c0, orig_c)
                    n_r1 = max(n_r1, orig_r + row_span - 1)
                    n_c1 = max(n_c1, orig_c + col_span - 1)
            if (n_r0, n_c0, n_r1, n_c1) == (r0, c0, r1, c1):
                break
            r0, c0, r1, c1 = n_r0, n_c0, n_r1, n_c1

        return (r0, c0, r1, c1)

    def _selection_merges(self):
        """Origins of every merged cell inside the current selection."""
        r0, c0, r1, c1 = self._sel_rect()
        seen = set()
        merges = []
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                orig = self._grid.origin_of(r, c)
                if orig in seen:
                    continue
                seen.add(orig)
                row_span, col_span = self._grid.span_of(orig[0], orig[1])
                if row_span > 1 or col_span > 1:
                    merges.append(orig)
        return merges

    def _sel_origins(self):
        """Every unique merge-origin cell touched by the current selection."""
        r0, c0, r1, c1 = self._sel_rect()
        seen = set()
        origins = []
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                orig = self._grid.origin_of(r, c)
                if orig not in seen:
                    seen.add(orig)
                    origins.append(orig)
        return origins

    def _apply_to_selection(self, fn):
        for (r, c) in self._sel_origins():
            block = self._grid.block_at(r, c)
            if block is None:
                block = studio_blocks.new_block('text', content='')
                self._grid.set_block(r, c, block)
            fn(block)
        self._render_all()

    def _anchor_block(self):
        ar, ac, _, _ = self._sel
        return self._grid.block_at(ar, ac)

    # -- Page Layout actions ---------------------------------------------------
    def _on_page_size_changed(self, sender, args):
        name = self.page_size_combo.SelectedItem
        if not name:
            return
        self._grid.set_page_size(name)
        self._render_all()

    def _on_page_orientation_changed(self, sender, args):
        val = self.page_orientation_combo.SelectedItem
        if not val:
            return
        self._grid.set_orientation(val.lower())
        self._render_all()

    # -- Home ribbon actions -----------------------------------------------------
    def _on_font_family_changed(self, sender, args):
        if self._syncing_ui:
            return
        fam = self.font_family_combo.SelectedItem
        if not fam:
            return
        self._apply_to_selection(lambda b: b.__setitem__('font', fam))

    def _on_font_size_changed(self, sender, args):
        if self._syncing_ui:
            return
        val = self.font_size_combo.SelectedItem
        if not val:
            return
        try:
            size_mm = studio_blocks.pt_to_mm(float(val))
        except Exception:
            return
        self._apply_to_selection(lambda b: b.__setitem__('size_mm', size_mm))

    def _on_bold_click(self, sender, args):
        cur = bool((self._anchor_block() or {}).get('bold'))
        self._apply_to_selection(lambda b: b.__setitem__('bold', not cur))

    def _on_italic_click(self, sender, args):
        cur = bool((self._anchor_block() or {}).get('italic'))
        self._apply_to_selection(lambda b: b.__setitem__('italic', not cur))

    def _on_underline_click(self, sender, args):
        cur = bool((self._anchor_block() or {}).get('underline'))
        self._apply_to_selection(lambda b: b.__setitem__('underline', not cur))

    def _pick_color(self, current_hex):
        dlg = _WF.ColorDialog()
        try:
            h = (current_hex or '#000000').lstrip('#')
            if len(h) == 3:
                h = h[0] * 2 + h[1] * 2 + h[2] * 2
            dlg.Color = _SD.Color.FromArgb(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
        except Exception:
            pass
        if dlg.ShowDialog() == _WF.DialogResult.OK:
            c = dlg.Color
            return '#{:02X}{:02X}{:02X}'.format(c.R, c.G, c.B)
        return None

    def _on_font_color_click(self, sender, args):
        cur = (self._anchor_block() or {}).get('color', '#000000')
        picked = self._pick_color(cur)
        if picked:
            self._apply_to_selection(lambda b: b.__setitem__('color', picked))

    def _on_fill_color_click(self, sender, args):
        cur = (self._anchor_block() or {}).get('bg_color') or '#FFFFFF'
        picked = self._pick_color(cur)
        if picked:
            self._apply_to_selection(lambda b: b.__setitem__('bg_color', picked))

    def _on_align_click(self, which):
        for name, val in (('align_left_btn', 'left'), ('align_center_btn', 'center'), ('align_right_btn', 'right')):
            getattr(self, name).IsChecked = (val == which)
        self._apply_to_selection(lambda b: b.__setitem__('just', which))

    def _on_valign_click(self, which):
        for name, val in (('valign_top_btn', 'top'), ('valign_mid_btn', 'middle'), ('valign_bot_btn', 'bottom')):
            getattr(self, name).IsChecked = (val == which)
        self._apply_to_selection(lambda b: b.__setitem__('v_just', which))


    def _on_merge_toggle_click(self, sender, args):
        """Single on/off toggle, Excel-style. ToggleButton flips IsChecked
        before Click fires, so it already holds the requested new state.

        Unchecking unmerges *every* merge in the selection, not just the
        anchor's - so a mixed selection (some merged cells, some plain) can
        be flattened in one go.
        """
        r0, c0, r1, c1 = self._sel_rect()
        if self.merge_toggle_btn.IsChecked:
            if (r0, c0) == (r1, c1):
                self.merge_toggle_btn.IsChecked = False  # nothing to merge
                return
            try:
                self._grid.merge(r0, c0, r1, c1)
            except ValueError as e:
                _alert(str(e))
                self.merge_toggle_btn.IsChecked = False
                return
            # Keep the whole merged area selected - collapsing to the origin
            # made a fresh merge look like only its first cell was selected.
            self._sel = (r0, c0, r1, c1)
        else:
            merges = self._selection_merges()
            for (mr, mc) in merges:
                self._grid.unmerge(mr, mc)
            if merges:
                self.status_label.Text = 'Unmerged {} cell(s)'.format(len(merges))
        self._render_all()

    # -- Row/column insert & delete, via right-click on a header --------------
    def _ctx_insert_row(self, idx, before):
        at = idx if before else idx + 1
        self._grid.insert_row(at)
        self._sel = (at, self._sel[1], at, self._sel[1])
        self._render_all()

    def _ctx_delete_row(self, idx):
        self._grid.delete_row(idx)
        ar = min(idx, self._grid.n_rows - 1)
        self._sel = (ar, self._sel[1], ar, self._sel[1])
        self._render_all()

    def _ctx_insert_col(self, idx, before):
        at = idx if before else idx + 1
        self._grid.insert_col(at)
        self._sel = (self._sel[0], at, self._sel[0], at)
        self._render_all()

    def _ctx_delete_col(self, idx):
        self._grid.delete_col(idx)
        ac = min(idx, self._grid.n_cols - 1)
        self._sel = (self._sel[0], ac, self._sel[0], ac)
        self._render_all()

    def _build_header_context_menu(self, kind, idx):
        menu = _themed_menu()

        def _item(text, handler, dot=None):
            menu.Items.Add(_themed_menu_item(text, handler, dot))

        if kind == 'colhdr':
            _item('Insert Column Left',  lambda s, a, i=idx: self._ctx_insert_col(i, True))
            _item('Insert Column Right', lambda s, a, i=idx: self._ctx_insert_col(i, False))
            menu.Items.Add(_themed_separator())
            _item('Delete Column',       lambda s, a, i=idx: self._ctx_delete_col(i))
        else:
            _item('Insert Row Above', lambda s, a, i=idx: self._ctx_insert_row(i, True))
            _item('Insert Row Below', lambda s, a, i=idx: self._ctx_insert_row(i, False))
            menu.Items.Add(_themed_separator())
            _item('Delete Row',       lambda s, a, i=idx: self._ctx_delete_row(i))
            menu.Items.Add(_themed_separator())
            # Print Titles (Excel's term): which rows reappear on every page.
            _item(studio_grid.SECTION_LABELS[studio_grid.SECTION_REPEAT_TOP],
                  lambda s, a, i=idx: self._ctx_set_section(i, studio_grid.SECTION_REPEAT_TOP),
                  studio_grid.SECTION_COLORS.get(studio_grid.SECTION_REPEAT_TOP))
            _item(studio_grid.SECTION_LABELS[studio_grid.SECTION_REPEAT_BOTTOM],
                  lambda s, a, i=idx: self._ctx_set_section(i, studio_grid.SECTION_REPEAT_BOTTOM),
                  studio_grid.SECTION_COLORS.get(studio_grid.SECTION_REPEAT_BOTTOM))
            _item(studio_grid.SECTION_LABELS[studio_grid.SECTION_BODY],
                  lambda s, a, i=idx: self._ctx_set_section(i, studio_grid.SECTION_BODY),
                  'outline')
        return menu

    def _ctx_set_section(self, idx, section):
        """Apply to the whole selected row range when the right-clicked row
        is part of it, otherwise just to that one row - the usual
        right-click-inside-a-selection convention."""
        r0, _c0, r1, _c1 = self._sel_rect()
        rows = range(r0, r1 + 1) if r0 <= idx <= r1 else [idx]
        for r in rows:
            self._grid.set_row_section(r, section)
        self.status_label.Text = '{}: rows {}'.format(
            studio_grid.SECTION_LABELS.get(section, section),
            ', '.join(str(r + 1) for r in rows))
        self._render_all()

    def _on_fit_width_click(self, sender, args):
        before = sum(self._grid.col_widths)
        self._grid.fit_columns_to_width()
        after = sum(self._grid.col_widths)
        self.status_label.Text = (
            'Columns scaled {:.0f}mm -> {:.0f}mm to fill the printable width '
            '({:.0f}mm page - {:.0f}mm margins)'.format(
                before, after, self._grid.page_w_mm, 2 * self._grid.margin_mm))
        self._render_all()

    def _on_distribute_click(self, sender, args):
        _r0, c0, _r1, c1 = self._sel_rect()
        # A whole-sheet / single-cell selection means "all columns".
        if c0 == c1:
            c0, c1 = 0, self._grid.n_cols - 1
        self._grid.distribute_columns(c0, c1)
        self.status_label.Text = 'Columns {}-{} distributed evenly'.format(
            _col_letter(c0), _col_letter(c1))
        self._render_all()

    def _on_margin_commit(self, sender, args):
        try:
            val = float(self.margin_box.Text)
        except Exception:
            self.margin_box.Text = str(self._grid.margin_mm)
            return
        # Leave room for at least a sliver of printable area.
        val = max(0.0, min(val, min(self._grid.page_w_mm, self._grid.page_h_mm) / 2.0 - 5))
        self._grid.margin_mm = round(val, 2)
        self.margin_box.Text = str(self._grid.margin_mm)
        self._render_all()

    def _on_row_h_commit(self, sender, args):
        try:
            val = float(self.row_h_box.Text)
        except Exception:
            return
        r0, _, r1, _ = self._sel_rect()
        for r in range(r0, r1 + 1):
            self._grid.row_heights[r] = val
        self._render_all()

    def _on_col_w_commit(self, sender, args):
        try:
            val = float(self.col_w_box.Text)
        except Exception:
            return
        _, c0, _, c1 = self._sel_rect()
        for c in range(c0, c1 + 1):
            self._grid.col_widths[c] = val
        self._render_all()

    # -- Formula bar ---------------------------------------------------------
    def _on_formula_commit(self, sender, args):
        ar, ac, _, _ = self._sel
        block = self._grid.block_at(ar, ac)
        if block is not None and block.get('type') != 'text':
            return  # read-only for data-driven blocks
        text = self.formula_bar_tb.Text
        if block is None:
            if text:
                self._grid.set_block(ar, ac, studio_blocks.new_block('text', content=text))
                self._render_all()
        else:
            if block.get('content', '') != text:
                block['content'] = text
                self._render_all()

    def _on_formula_keydown(self, sender, args):
        if args.Key == _SWI.Key.Enter:
            self._on_formula_commit(sender, args)
            self.grid_scroll.Focus()

    def _on_formula_change_click(self, sender, args):
        ar, ac, _, _ = self._sel
        if _confirm('Clear this cell so a different block can be inserted from the Blocks panel?'):
            self._grid.clear_cell(ar, ac)
            self._render_all()

    def _selection_kind(self):
        """'row' / 'col' / 'cell' - taken from how the selection was made
        rather than inferred from its shape, so a cell drag that happens to
        cover every column isn't mistaken for a row selection (and vice
        versa once merges are involved)."""
        return getattr(self, '_sel_mode', 'cell')

    def _update_formula_bar(self):
        ar, ac, br, bc = self._sel
        r0, c0, r1, c1 = self._sel_rect()
        kind = self._selection_kind()

        # Excel-style reference: 1:1 for a row, A:A for a column.
        if kind == 'row':
            self.formula_cell_label.Text = ('{}:{}'.format(r0 + 1, r1 + 1)
                                            if r1 != r0 else '{}:{}'.format(r0 + 1, r0 + 1))
        elif kind == 'col':
            self.formula_cell_label.Text = '{}:{}'.format(_col_letter(c0), _col_letter(c1))
        elif (ar, ac) == (br, bc):
            self.formula_cell_label.Text = _cell_ref(ar, ac)
        else:
            self.formula_cell_label.Text = '{}:{}'.format(_cell_ref(ar, ac), _cell_ref(br, bc))

        if kind in ('row', 'col', 'sheet'):
            # A row/column selection isn't a cell, so don't report a block -
            # describe what was actually selected instead.
            if kind == 'row':
                n = r1 - r0 + 1
                label = ('Row {}'.format(r0 + 1) if n == 1
                         else 'Rows {}-{}'.format(r0 + 1, r1 + 1))
                # Append the applied section inline, worded exactly as the
                # row-header menu words it.
                secs = set(self._grid.section_of(r) for r in range(r0, r1 + 1))
                if len(secs) == 1:
                    only = list(secs)[0]
                    if only != studio_grid.SECTION_BODY:
                        label += ' - {}'.format(studio_grid.SECTION_LABELS.get(only, only))
                elif len(secs) > 1:
                    label += ' - Mixed sections'
                self.formula_bar_tb.Text = label
            elif kind == 'col':
                n = c1 - c0 + 1
                self.formula_bar_tb.Text = ('Column {}'.format(_col_letter(c0)) if n == 1
                                            else 'Columns {}-{}'.format(_col_letter(c0), _col_letter(c1)))
            else:
                self.formula_bar_tb.Text = 'Entire sheet'
            self.formula_bar_tb.IsReadOnly = True
            self.formula_change_btn.Visibility = _SW.Visibility.Collapsed
            self._update_section_label()
            self._sync_format_buttons(self._grid.block_at(ar, ac) or {})
            return

        block = self._grid.block_at(ar, ac)
        if block is None:
            self.formula_bar_tb.Text = ''
            self.formula_bar_tb.IsReadOnly = False
            self.formula_change_btn.Visibility = _SW.Visibility.Collapsed
        elif block.get('type') == 'text':
            self.formula_bar_tb.Text = block.get('content', '')
            self.formula_bar_tb.IsReadOnly = False
            self.formula_change_btn.Visibility = _SW.Visibility.Collapsed
        else:
            summary = studio_blocks.block_display_summary(block)
            if str(block.get('type', '')).startswith('spine_'):
                # Show which revision this cell is linked to, and which
                # revision that actually resolves to in the live model.
                idx = self._rev_index_for(block, ac, self._rev_column_map())
                revs = self._data.get('revisions', [])
                pinned = isinstance(block.get('rev_index'), int)
                mark = revs[idx].get('rev', '') if idx < len(revs) else None
                summary += '   -> Revision {}{}{}'.format(
                    idx + 1,
                    ' "{}"'.format(mark) if mark else ' (not issued yet)',
                    ' [pinned]' if pinned else '')
            self.formula_bar_tb.Text = summary
            self.formula_bar_tb.IsReadOnly = True
            self.formula_change_btn.Visibility = _SW.Visibility.Visible

        self._update_section_label()
        self._sync_format_buttons(block or {})

    def _update_section_label(self):
        """Compact row-section indicator, colour-matched to the gutter stripe
        and the palette legend. Hidden for whole-row selections, where the
        formula bar already spells the section out in full."""
        try:
            r0, _c0, r1, _c1 = self._sel_rect()
            secs = set(self._grid.section_of(r) for r in range(r0, r1 + 1))
            secs.discard(studio_grid.SECTION_BODY)
            if self._selection_kind() == 'row':
                secs = set()
            if not secs:
                self.section_label.Text = ''
                self.section_label.Visibility = _SW.Visibility.Collapsed
            else:
                names = [studio_grid.SECTION_SHORT.get(s, s) for s in sorted(secs)]
                self.section_label.Text = ' / '.join(names)
                self.section_label.Foreground = _brush(
                    studio_grid.SECTION_COLORS.get(sorted(secs)[0], '#8A96A8'))
                self.section_label.Visibility = _SW.Visibility.Visible
        except Exception:
            pass

    def _sync_format_buttons(self, b):
        """Reflect a block's formatting on the ribbon controls."""
        # Font family/size are ComboBoxes, so guard against their
        # SelectionChanged handlers re-applying what we're only displaying.
        self._syncing_ui = True
        try:
            fam = b.get('font')
            if fam and fam in list(self.font_family_combo.Items):
                self.font_family_combo.SelectedItem = fam
            size_mm = b.get('size_mm')
            if size_mm:
                pt_txt = str(int(round(studio_blocks.mm_to_pt(size_mm))))
                if pt_txt in list(self.font_size_combo.Items):
                    self.font_size_combo.SelectedItem = pt_txt
        except Exception:
            pass
        finally:
            self._syncing_ui = False

        self.bold_btn.IsChecked = bool(b.get('bold'))
        self.italic_btn.IsChecked = bool(b.get('italic'))
        self.underline_btn.IsChecked = bool(b.get('underline'))
        just = b.get('just', 'left')
        self.align_left_btn.IsChecked = (just == 'left')
        self.align_center_btn.IsChecked = (just == 'center')
        self.align_right_btn.IsChecked = (just == 'right')
        vjust = b.get('v_just', 'middle')
        self.valign_top_btn.IsChecked = (vjust == 'top')
        self.valign_mid_btn.IsChecked = (vjust == 'middle')
        self.valign_bot_btn.IsChecked = (vjust == 'bottom')

        # Lit whenever the selection contains any merge, so a mixed
        # selection offers "unmerge everything" on the next click.
        self.merge_toggle_btn.IsChecked = bool(self._selection_merges())

    # ======================================================================
    # Grid rendering
    # ======================================================================

    def _render_all(self):
        try:
            self.margin_box.Text = str(self._grid.margin_mm)
        except Exception:
            pass
        self._render_grid()
        self._update_formula_bar()
        if self._active_path:
            self.active_file_label.Text = '- {}'.format(os.path.basename(self._active_path))
        else:
            self.active_file_label.Text = '- (unsaved)'

    def _rev_column_map(self):
        """Map column index -> revision index, for the "smart link" between
        revision blocks and revisions.

        Every column that contains at least one revision block counts as a
        revision column; they're numbered left to right. So dropping a
        revision block into the next column over automatically binds it to
        the next revision, and a date / mark / initials stacked in the same
        column all describe the same revision. A block can opt out by
        setting an explicit integer 'rev_index'.
        """
        grid = self._grid
        rev_cols = set()
        for (r, c), cell in grid.cells.items():
            if 'covered_by' in cell:
                continue
            block = cell.get('block') or {}
            if str(block.get('type', '')).startswith('spine_'):
                rev_cols.add(c)
        return {c: i for i, c in enumerate(sorted(rev_cols))}

    def _rev_index_for(self, block, col, rev_col_map):
        pinned = block.get('rev_index', 'auto')
        if isinstance(pinned, int):
            return max(0, pinned)
        return rev_col_map.get(col, 0)

    def _render_grid(self):
        grid = self._grid
        rev_col_map = self._rev_column_map()

        row_px = [int(h * SCALE) for h in grid.row_heights]
        col_px = [int(w * SCALE) for w in grid.col_widths]
        self._row_px = row_px
        self._col_px = col_px

        # Three separate grids so the headers can live in their own
        # ScrollViewers and stay frozen while the cell area scrolls (see the
        # four-quadrant layout in StudioSettings.xaml). They share the same
        # column widths / row heights, so everything stays lined up.
        cells_grid = _SWC.Grid()
        cells_grid.Background = _SWM.Brushes.Transparent
        colhdr_grid = _SWC.Grid()
        rowhdr_grid = _SWC.Grid()

        for h in row_px:
            rd = _SWC.RowDefinition(); rd.Height = _SW.GridLength(max(h, 4))
            cells_grid.RowDefinitions.Add(rd)
            rd2 = _SWC.RowDefinition(); rd2.Height = _SW.GridLength(max(h, 4))
            rowhdr_grid.RowDefinitions.Add(rd2)
        for w in col_px:
            cd = _SWC.ColumnDefinition(); cd.Width = _SW.GridLength(max(w, 4))
            cells_grid.ColumnDefinitions.Add(cd)
            cd2 = _SWC.ColumnDefinition(); cd2.Width = _SW.GridLength(max(w, 4))
            colhdr_grid.ColumnDefinitions.Add(cd2)

        chr_rd = _SWC.RowDefinition(); chr_rd.Height = _SW.GridLength(HEADER_H)
        colhdr_grid.RowDefinitions.Add(chr_rd)
        rhc_cd = _SWC.ColumnDefinition(); rhc_cd.Width = _SW.GridLength(HEADER_W)
        rowhdr_grid.ColumnDefinitions.Add(rhc_cd)

        # Column headers
        for c in range(grid.n_cols):
            hdr = _SWC.Border()
            hdr.Background = _brush('#2B3340')
            hdr.BorderBrush = _brush('#404E62')
            hdr.BorderThickness = _SW.Thickness(0, 0, 1, 1)
            tb = _SWC.TextBlock()
            tb.Text = _col_letter(c)
            tb.Foreground = _SWM.Brushes.White
            tb.FontSize = 10
            tb.HorizontalAlignment = _SW.HorizontalAlignment.Center
            tb.VerticalAlignment = _SW.VerticalAlignment.Center
            hdr.Child = tb
            hdr.Cursor = _SWI.Cursors.Hand
            hdr.Tag = ('colhdr', c)
            hdr.ContextMenu = self._build_header_context_menu('colhdr', c)
            hdr.MouseLeftButtonDown += self._on_header_down
            hdr.MouseMove += self._on_header_mouse_move
            hdr.MouseLeftButtonUp += self._on_header_up
            _SWC.Grid.SetRow(hdr, 0); _SWC.Grid.SetColumn(hdr, c)
            colhdr_grid.Children.Add(hdr)

        # Row headers
        for r in range(grid.n_rows):
            hdr = _SWC.Border()
            hdr.Background = _brush('#2B3340')
            section = grid.section_of(r)
            if section != studio_grid.SECTION_BODY:
                # Coloured stripe down the left of the row number marks a
                # repeating band at a glance, without stealing header width.
                hdr.BorderBrush = _brush(studio_grid.SECTION_COLORS.get(section, '#404E62'))
                hdr.BorderThickness = _SW.Thickness(4, 0, 1, 1)
                hdr.ToolTip = studio_grid.SECTION_SHORT.get(section, section)
            else:
                hdr.BorderBrush = _brush('#404E62')
                hdr.BorderThickness = _SW.Thickness(0, 0, 1, 1)
            tb = _SWC.TextBlock()
            tb.Text = str(r + 1)
            tb.Foreground = _SWM.Brushes.White
            tb.FontSize = 10
            tb.HorizontalAlignment = _SW.HorizontalAlignment.Center
            tb.VerticalAlignment = _SW.VerticalAlignment.Center
            hdr.Child = tb
            hdr.Cursor = _SWI.Cursors.Hand
            hdr.Tag = ('rowhdr', r)
            hdr.ContextMenu = self._build_header_context_menu('rowhdr', r)
            hdr.MouseLeftButtonDown += self._on_header_down
            hdr.MouseMove += self._on_header_mouse_move
            hdr.MouseLeftButtonUp += self._on_header_up
            _SWC.Grid.SetRow(hdr, r); _SWC.Grid.SetColumn(hdr, 0)
            rowhdr_grid.Children.Add(hdr)

        # Content cells (skip covered cells; draw merge origins with spans)
        for r in range(grid.n_rows):
            for c in range(grid.n_cols):
                cell = grid.cells.get((r, c))
                if cell is None or 'covered_by' in cell:
                    continue
                row_span = cell.get('row_span', 1)
                col_span = cell.get('col_span', 1)
                border = _SWC.Border()
                border.BorderBrush = _brush('#404E62')
                b = cell.get('block') or {}
                borders = b.get('borders', {}) if b else {}
                border.BorderThickness = _SW.Thickness(
                    1 if borders.get('l') else 0.4,
                    1 if borders.get('t') else 0.4,
                    1 if borders.get('r') else 0.4,
                    1 if borders.get('b') else 0.4)
                border.Background = _SWM.Brushes.White
                # Clip to the cell: a block taller than its row would
                # otherwise render straight over the rows below, which looks
                # like the grid has broken when a row is made shorter.
                border.ClipToBounds = True
                content = studio_blocks.render_block(
                    cell.get('block'), self._data, SCALE, grid.logo_path,
                    rev_index=self._rev_index_for(b, c, rev_col_map))
                if content:
                    border.Child = content
                border.Tag = ('cell', r, c)
                border.MouseLeftButtonDown += self._on_cell_down
                border.MouseMove += self._on_cell_move
                border.MouseLeftButtonUp += self._on_cell_up
                border.AllowDrop = True
                border.DragEnter += self._on_cell_drag_over
                border.DragOver += self._on_cell_drag_over
                border.Drop += self._on_cell_drop
                border.ContextMenu = self._build_cell_context_menu()
                _SWC.Grid.SetRow(border, r)
                _SWC.Grid.SetColumn(border, c)
                if row_span > 1:
                    _SWC.Grid.SetRowSpan(border, row_span)
                if col_span > 1:
                    _SWC.Grid.SetColumnSpan(border, col_span)
                cells_grid.Children.Add(border)

        # Selection overlay - a non-hit-testable Border spanning the selected
        # rectangle, drawn as the last (topmost) child of the cell Grid so
        # its Grid.Row/Column attached properties line up with the cells.
        r0, c0, r1, c1 = self._sel_rect()
        overlay = _SWC.Border()
        overlay.BorderBrush = _brush('#20A344')
        overlay.BorderThickness = _SW.Thickness(2)
        overlay.Background = _SWM.SolidColorBrush(_SWM.Color.FromArgb(30, 0x20, 0xA3, 0x44))
        overlay.IsHitTestVisible = False
        _SWC.Grid.SetRow(overlay, r0)
        _SWC.Grid.SetColumn(overlay, c0)
        _SWC.Grid.SetRowSpan(overlay, r1 - r0 + 1)
        _SWC.Grid.SetColumnSpan(overlay, c1 - c0 + 1)
        cells_grid.Children.Add(overlay)

        self._cells_grid = cells_grid
        self._colhdr_grid = colhdr_grid
        self._rowhdr_grid = rowhdr_grid
        self._overlay = overlay

        # Page split lines, layered over the cell area only.
        page_overlay = self._build_page_boundary_overlay(
            grid, col_px, row_px, sum(col_px), sum(row_px))

        outer = _SWC.Grid()
        outer.Children.Add(cells_grid)
        if page_overlay is not None:
            outer.Children.Add(page_overlay)

        self.grid_root.Children.Clear()
        self.grid_root.Children.Add(outer)
        self.colhdr_root.Children.Clear()
        self.colhdr_root.Children.Add(colhdr_grid)
        self.rowhdr_root.Children.Clear()
        self.rowhdr_root.Children.Add(rowhdr_grid)
        self._apply_zoom()
        # Rebuilding the panes resets their offsets - re-align them with the
        # cell pane so a re-render never leaves the headers adrift.
        self._sync_header_scroll()

    def _page_break_positions(self, px_list, page_px):
        """Offsets (from the start of the content area) where a page split
        falls along one axis, snapped to cell boundaries.

        Walks the cells accumulating size; when the next cell would not fit
        on the current page, a break is emitted at the boundary *before* it
        and the next page starts measuring from there. Measuring each page
        from the previous break is the part that matters: page N's split is
        NOT at N * page_px, because every page loses whatever slack was left
        over when its last cell didn't fit. Computing them as fixed
        multiples is what made the lines drift out of place.

        Only interior splits are returned - the far edge of the last page
        isn't a split, it's just where the content ends, and drawing a line
        there was the stray line that appeared at the bottom of the grid.
        """
        if page_px <= 0:
            return []
        breaks = []
        pos = 0          # running offset of the current cell
        page_start = 0   # offset where the current page begins
        for size in px_list:
            # pos > page_start guarantees at least one cell per page, so a
            # single cell taller/wider than a whole page just overflows on
            # its own page instead of emitting endless zero-height pages.
            if pos > page_start and (pos + size) - page_start > page_px:
                breaks.append(pos)
                page_start = pos
            pos += size
        return breaks

    def _row_break_positions(self, row_px, sections, page_px):
        """Vertical page breaks, accounting for repeating rows.

        Rows marked repeat-at-top / repeat-at-bottom appear on *every* page,
        so they reduce how much body content each page can hold. Only body
        rows flow across pages; the break is recorded at the row's real y
        offset in the sheet so the line still lands in the right place.
        """
        top_h = sum(px for px, s in zip(row_px, sections)
                    if s == studio_grid.SECTION_REPEAT_TOP)
        bottom_h = sum(px for px, s in zip(row_px, sections)
                       if s == studio_grid.SECTION_REPEAT_BOTTOM)
        capacity = page_px - top_h - bottom_h
        if capacity <= 0:
            # Repeating bands alone fill the page - nothing sensible to draw.
            return []

        breaks = []
        pos = 0
        used = 0
        for px, sec in zip(row_px, sections):
            if sec != studio_grid.SECTION_BODY:
                pos += px
                continue
            if used > 0 and used + px > capacity:
                breaks.append(pos)
                used = 0
            used += px
            pos += px
        return breaks

    def _dashed_line(self, x1, y1, x2, y2):
        line = _SWS.Line()
        line.X1, line.Y1, line.X2, line.Y2 = x1, y1, x2, y2
        line.Stroke = _brush('#2980B9')
        line.StrokeThickness = 1.5
        dash = _SWM.DoubleCollection()
        dash.Add(4.0)
        dash.Add(2.0)
        line.StrokeDashArray = dash
        return line

    def _build_page_boundary_overlay(self, grid, col_px, row_px, total_w_px, total_h_px):
        """One dashed line per page split - the plain Excel Page Break
        Preview look. Returns None when the content fits on a single page
        in both directions, since then there is nothing to split.

        Coordinates are relative to the cell area (the headers are separate
        frozen panes now), so no header offset is applied.
        """
        # Printable area, not the raw paper size - the line marks where
        # content must stop, which is the margin, not the paper edge.
        page_w_px = int(grid.printable_w_mm() * SCALE)
        page_h_px = int(grid.printable_h_mm() * SCALE)
        content_w_px = sum(col_px)

        x_breaks = self._page_break_positions(col_px, page_w_px)
        y_breaks = self._row_break_positions(row_px, grid.row_sections, page_h_px)

        # The width line is a ruler, not just a break - it answers "do my
        # columns fit the page?", so it must show even when content is
        # narrower than the page. Past the content there's no column edge to
        # snap to, so draw it at the true page width.
        if page_w_px > content_w_px:
            x_breaks = list(x_breaks) + [page_w_px]

        if not x_breaks and not y_breaks:
            return None

        canvas = _SWC.Canvas()
        # Stretch far enough right for a page edge that sits past the content.
        canvas.Width = max(total_w_px, (max(x_breaks) + 2) if x_breaks else total_w_px)
        canvas.Height = total_h_px
        canvas.IsHitTestVisible = False
        # Pin to the same origin as the cell Grid this is layered over. A
        # Canvas with an explicit Width inside a Grid otherwise centres
        # itself in whatever space it's given, which shifted every line
        # sideways from the column it was supposed to mark.
        canvas.HorizontalAlignment = _SW.HorizontalAlignment.Left
        canvas.VerticalAlignment = _SW.VerticalAlignment.Top

        for x in x_breaks:
            canvas.Children.Add(self._dashed_line(x, 0, x, total_h_px))
        for y in y_breaks:
            canvas.Children.Add(self._dashed_line(0, y, total_w_px, y))

        return canvas

    def _hit_cell(self, x, y):
        """Translate a point (relative to self._cells_grid's origin, i.e. the
        same coordinate space MouseEventArgs.GetPosition(self._cells_grid)
        returns) into a (row, col) index, clamped to the grid's bounds. Used
        during drag-select, where mouse capture pins every event to the cell
        that was originally clicked - args/sender no longer reflect where
        the pointer currently is, only this pixel math does."""
        xi = x
        yi = y
        c = 0
        acc = 0
        for i, w in enumerate(self._col_px):
            if xi < acc + w:
                c = i
                break
            acc += w
            c = i
        r = 0
        acc = 0
        for i, h in enumerate(self._row_px):
            if yi < acc + h:
                r = i
                break
            acc += h
            r = i
        r = max(0, min(r, self._grid.n_rows - 1))
        c = max(0, min(c, self._grid.n_cols - 1))
        return r, c

    def _move_overlay_to_selection(self):
        r0, c0, r1, c1 = self._sel_rect()
        _SWC.Grid.SetRow(self._overlay, r0)
        _SWC.Grid.SetColumn(self._overlay, c0)
        _SWC.Grid.SetRowSpan(self._overlay, r1 - r0 + 1)
        _SWC.Grid.SetColumnSpan(self._overlay, c1 - c0 + 1)

    # -- mouse interaction: row/column headers (select, or drag-resize) --------
    def _index_from_pos(self, px_list, value):
        acc = 0
        for i, size in enumerate(px_list):
            if value < acc + size:
                return i
            acc += size
        return max(0, len(px_list) - 1)

    def _on_header_mouse_move(self, sender, args):
        kind, idx = sender.Tag
        if self._resizing:
            self._apply_resize_drag(args)
            return
        if self._header_drag and args.LeftButton == _SWI.MouseButtonState.Pressed:
            anchor = getattr(self, '_header_anchor', idx)
            if self._header_drag == 'colhdr':
                pos = args.GetPosition(self._colhdr_grid)
                cur = self._index_from_pos(self._col_px, pos.X)
                self._sel = (0, anchor, self._grid.n_rows - 1, cur)
            else:
                pos = args.GetPosition(self._rowhdr_grid)
                cur = self._index_from_pos(self._row_px, pos.Y)
                self._sel = (anchor, 0, cur, self._grid.n_cols - 1)
            self._move_overlay_to_selection()
            self._update_formula_bar()
            return
        # Hover feedback: only show a resize cursor near the header's far edge
        pos = args.GetPosition(sender)
        if kind == 'colhdr':
            near_edge = pos.X >= sender.ActualWidth - HEADER_RESIZE_MARGIN
            sender.Cursor = _SWI.Cursors.SizeWE if near_edge else _SWI.Cursors.Hand
        else:
            near_edge = pos.Y >= sender.ActualHeight - HEADER_RESIZE_MARGIN
            sender.Cursor = _SWI.Cursors.SizeNS if near_edge else _SWI.Cursors.Hand

    def _on_header_down(self, sender, args):
        kind, idx = sender.Tag
        pos = args.GetPosition(sender)
        if kind == 'colhdr' and pos.X >= sender.ActualWidth - HEADER_RESIZE_MARGIN:
            self._start_resize('col', idx, sender)
            return
        if kind == 'rowhdr' and pos.Y >= sender.ActualHeight - HEADER_RESIZE_MARGIN:
            self._start_resize('row', idx, sender)
            return
        # Not near the resize edge - normal select-whole-row/column
        if kind == 'colhdr':
            self._sel_mode = 'col'
            self._sel = (0, idx, self._grid.n_rows - 1, idx)
        else:
            self._sel_mode = 'row'
            self._sel = (idx, 0, idx, self._grid.n_cols - 1)
        # Dragging along the headers extends to a range, so several rows or
        # columns can be sized or distributed in one go.
        self._header_drag = kind
        self._header_anchor = idx
        try:
            sender.CaptureMouse()
            self._resize_owner = sender
        except Exception:
            pass
        self._move_overlay_to_selection()
        self._update_formula_bar()

    def _on_header_up(self, sender, args):
        if self._header_drag:
            self._header_drag = None
            try:
                sender.ReleaseMouseCapture()
            except Exception:
                pass
        if self._resizing:
            try:
                getattr(self, '_resize_owner', sender).ReleaseMouseCapture()
            except Exception:
                pass
            self._resizing = None
            self._resize_group = None
            # Row/col sizes changed, so the page-break lines, the selection
            # overlay and any content that reflows are all stale - rebuild
            # rather than leaving the drag's live edits half-applied.
            self._render_all()

    def _resize_ref(self, kind):
        """Measure resize drags against the header grid being dragged - its
        columns/rows mirror the cell grid's, so offsets match either way."""
        return self._colhdr_grid if kind == 'col' else self._rowhdr_grid

    def _resize_group_for(self, kind, idx):
        """Which columns/rows a drag on `idx` should size.

        If the dragged header is part of a multi-header selection, the whole
        selection is sized to the dragged width - Excel's behaviour. A drag
        outside the selection sizes just that one.
        """
        r0, c0, r1, c1 = self._sel_rect()
        if kind == 'col' and self._sel_mode == 'col' and c0 <= idx <= c1 and c1 > c0:
            return list(range(c0, c1 + 1))
        if kind == 'row' and self._sel_mode == 'row' and r0 <= idx <= r1 and r1 > r0:
            return list(range(r0, r1 + 1))
        return [idx]

    def _start_resize(self, kind, idx, sender):
        self._resizing = (kind, idx)
        ref_pos = _SWI.Mouse.GetPosition(self._resize_ref(kind))
        self._resize_start_screen = ref_pos.X if kind == 'col' else ref_pos.Y
        self._resize_start_size = (self._grid.col_widths if kind == 'col' else self._grid.row_heights)[idx]
        self._resize_group = self._resize_group_for(kind, idx)
        try:
            sender.CaptureMouse()
        except Exception:
            pass
        self._resize_owner = sender

    def _apply_resize_drag(self, args):
        kind, idx = self._resizing
        pos = args.GetPosition(self._resize_ref(kind))
        # Every header in the group takes the dragged size, so a selected
        # run of columns/rows resizes together.
        group = getattr(self, '_resize_group', None) or [idx]
        if kind == 'col':
            delta = pos.X - self._resize_start_screen
            new_mm = max(5.0, self._resize_start_size + delta / SCALE)
            px = _SW.GridLength(max(int(new_mm * SCALE), 4))
            for i in group:
                self._grid.col_widths[i] = new_mm
                # Header and cell grids are separate now - resize both so
                # they stay aligned during the drag.
                self._colhdr_grid.ColumnDefinitions[i].Width = px
                self._cells_grid.ColumnDefinitions[i].Width = px
        else:
            delta = pos.Y - self._resize_start_screen
            new_mm = max(3.0, self._resize_start_size + delta / SCALE)
            px = _SW.GridLength(max(int(new_mm * SCALE), 4))
            for i in group:
                self._grid.row_heights[i] = new_mm
                self._rowhdr_grid.RowDefinitions[i].Height = px
                self._cells_grid.RowDefinitions[i].Height = px
        if len(group) > 1:
            self.status_label.Text = '{} {}s resized to {:.1f}mm'.format(
                len(group), 'column' if kind == 'col' else 'row', new_mm)

    # -- moving cell content by dragging the selection edge --------------------
    def _sel_pixel_bounds(self):
        r0, c0, r1, c1 = self._sel_rect()
        x0 = sum(self._col_px[:c0])
        x1 = x0 + sum(self._col_px[c0:c1 + 1])
        y0 = sum(self._row_px[:r0])
        y1 = y0 + sum(self._row_px[r0:r1 + 1])
        return x0, y0, x1, y1

    def _on_selection_edge(self, x, y):
        """True when (x, y) sits in the band around the selection outline.
        Dragging from there moves the content; dragging from inside the
        selection extends it, exactly as Excel behaves."""
        try:
            x0, y0, x1, y1 = self._sel_pixel_bounds()
        except Exception:
            return False
        m = MOVE_EDGE_MARGIN
        if not (x0 - m <= x <= x1 + m and y0 - m <= y <= y1 + m):
            return False
        near_v = abs(x - x0) <= m or abs(x - x1) <= m
        near_h = abs(y - y0) <= m or abs(y - y1) <= m
        return near_v or near_h

    def _move_selection_to(self, target_r, target_c):
        """Move every block in the selection so the selection's top-left
        lands on (target_r, target_c). Formatting travels with the block,
        since a block dict carries its own formatting."""
        r0, c0, r1, c1 = self._sel_rect()
        dr, dc = target_r - r0, target_c - c0
        if dr == 0 and dc == 0:
            return

        moving = []
        for (r, c) in self._sel_origins():
            blk = self._grid.block_at(r, c)
            if blk is not None:
                moving.append((r, c, blk))
        if not moving:
            return

        for (r, c, _blk) in moving:
            nr, nc = r + dr, c + dc
            if not (0 <= nr < self._grid.n_rows and 0 <= nc < self._grid.n_cols):
                _alert('That would move content off the edge of the sheet.')
                return

        # Only warn about targets that aren't themselves part of the moving set.
        sources = set((r, c) for (r, c, _b) in moving)
        clashes = [(r + dr, c + dc) for (r, c, _b) in moving
                   if (r + dr, c + dc) not in sources
                   and self._grid.block_at(r + dr, c + dc) is not None]
        if clashes and not _confirm(
                'There is already content in {} target cell(s). Replace it?'.format(len(clashes))):
            return

        # Clear all sources first, then write - otherwise an overlapping
        # move would clear cells it had just written.
        for (r, c, _blk) in moving:
            self._grid.clear_cell(r, c)
        for (r, c, blk) in moving:
            self._grid.set_block(r + dr, c + dc, blk)

        self._sel = (r0 + dr, c0 + dc, r1 + dr, c1 + dc)
        self.status_label.Text = 'Moved {} block(s)'.format(len(moving))
        self._render_all()

    # -- mouse interaction: grid cells (click/drag to select) ------------------
    def _on_cell_down(self, sender, args):
        _, r, c = sender.Tag
        # Pressing on the selection's outline starts a content move rather
        # than a new selection (Excel's drag-to-move).
        try:
            pos = args.GetPosition(self._cells_grid)
            if self._on_selection_edge(pos.X, pos.Y):
                self._pending_move = True
                self._move_start = pos
                return
        except Exception:
            pass

        # Just anchor on the clicked cell - _sel_rect() expands the rectangle
        # out to cover any merge it touches, so clicking a merged cell
        # selects the whole merged area.
        self._sel_mode = 'cell'
        self._sel = (r, c, r, c)
        self._dragging = True
        try:
            sender.CaptureMouse()
        except Exception:
            pass
        self._drag_owner = sender
        self._update_formula_bar()
        self._move_overlay_to_selection()

    def _on_cell_move(self, sender, args):
        # Waiting to see whether a press on the selection edge becomes a move
        if self._pending_move:
            if args.LeftButton != _SWI.MouseButtonState.Pressed:
                self._pending_move = None
                return
            pos = args.GetPosition(self._cells_grid)
            start = self._move_start
            if start is not None and (abs(pos.X - start.X) > 4 or abs(pos.Y - start.Y) > 4):
                self._pending_move = None
                self._move_start = None
                data = _SW.DataObject('studio_move_cells', True)
                try:
                    _SW.DragDrop.DoDragDrop(sender, data, _SW.DragDropEffects.Move)
                except Exception:
                    pass
            return

        if not self._dragging:
            # Hover feedback: show the move cursor over the selection outline
            try:
                pos = args.GetPosition(self._cells_grid)
                sender.Cursor = (_SWI.Cursors.SizeAll
                                 if self._on_selection_edge(pos.X, pos.Y)
                                 else _SWI.Cursors.Arrow)
            except Exception:
                pass
            return
        if args.LeftButton != _SWI.MouseButtonState.Pressed:
            return
        pos = args.GetPosition(self._cells_grid)
        r, c = self._hit_cell(pos.X, pos.Y)
        ar, ac, _, _ = self._sel
        # No merge handling needed here - _sel_rect() expands the rectangle
        # to whole merges, so a drag ending mid-merge still selects it fully.
        if (r, c) == (self._sel[2], self._sel[3]):
            return
        self._sel = (ar, ac, r, c)
        self._move_overlay_to_selection()

    def _on_cell_up(self, sender, args):
        # A press on the selection edge that never became a drag is just a
        # click - fall through to selecting that cell.
        if self._pending_move:
            self._pending_move = None
            self._move_start = None
            try:
                _, r, c = sender.Tag
                self._sel = (r, c, r, c)
                self._move_overlay_to_selection()
            except Exception:
                pass
        if self._dragging:
            self._dragging = False
            try:
                getattr(self, '_drag_owner', sender).ReleaseMouseCapture()
            except Exception:
                pass
        self._update_formula_bar()

    # -- drag-and-drop: a block dragged from the Blocks panel onto a cell -----
    def _on_cell_drag_over(self, sender, args):
        if args.Data.GetDataPresent('studio_block_type'):
            args.Effects = _SW.DragDropEffects.Copy
        elif args.Data.GetDataPresent('studio_move_cells'):
            args.Effects = _SW.DragDropEffects.Move
        else:
            args.Effects = getattr(_SW.DragDropEffects, 'None')
        args.Handled = True

    def _on_cell_drop(self, sender, args):
        _, r, c = sender.Tag
        if args.Data.GetDataPresent('studio_move_cells'):
            # The drop cell becomes the new top-left of what was selected.
            self._move_selection_to(r, c)
            args.Handled = True
            return
        if not args.Data.GetDataPresent('studio_block_type'):
            return
        block_type = args.Data.GetData('studio_block_type')
        self._grid.set_block(r, c, studio_blocks.new_block_from_palette(block_type))
        self._sel_mode = 'cell'
        self._sel = (r, c, r, c)
        self._render_all()
        args.Handled = True

    # -- right-click on a cell: Format Cells… ----------------------------------
    def _build_cell_context_menu(self):
        menu = _themed_menu()
        menu.Items.Add(_themed_menu_item('Format Cells…', self._on_format_cells_click))
        return menu

    def _on_format_cells_click(self, sender, args):
        open_format_cells_dialog(self)
