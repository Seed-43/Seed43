# -*- coding: utf-8 -*-
# StudioSettings.py
#
# pyTransmit Studio - an Excel/LibreOffice-Calc-style alternative to the
# Layout Builder (../Layout/LayoutSettings.py, untouched by this file): a
# free grid (any rows/columns), drag-select + merge/unmerge, a tabbed
# formatting ribbon, a formula bar showing the block in the active cell, and a
# preview fed by live Revit data instead of dummy placeholders.
#
# The Sheet Data tab's Sheet Grouping and Sheet Rows groups control how the
# documentation table is previewed rather than how the layout is built:
# grouping mirrors pyTransmit's "Grouping Drawing Sheets by Selected
# Parameters" option (including its Text On/Off toggle) so the preview shows
# what will really be published, and Condense Rows shortens a 2000-sheet list
# to 11 rows so the layout stays workable. Both are preview state, saved in
# studio_config.json - see _load_preview_options() and
# studio_blocks.sheet_row_plan().

import os

# pytransmit_paths lives in the pushbutton root, which is not guaranteed to be
# on sys.path - pyTransmit loads this module by inserting only its own folder.
import sys as _sys
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_PT_ROOT = os.path.dirname(_SCRIPT_DIR)
if _PT_ROOT not in _sys.path:
    _sys.path.insert(0, _PT_ROOT)
from pytransmit_paths import (SETTINGS_DIR, STUDIO_LAYOUTS_DIR, STUDIO_CONFIG,
                              SETUP_FILE, SYNC_FILE, LOGOS_DIR)
import shutil
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
import studio_rows
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
MIN_CELL_PX = 4      # smallest row/column the grid will lay out
MIN_COL_MM = 5.0     # drag-resize floors, in the model's own units
MIN_ROW_MM = 3.0
HEADER_H = 22
HEADER_W = 40
HEADER_RESIZE_MARGIN = 5   # px from a header's far edge that starts a drag-resize
# Excel's own Page Break Preview blue. Sits on the white sheet, so it is a
# fixed print-preview colour rather than a palette token (see Theme below).
PAGE_BREAK_COLOR = '#2980B9'
# Documentation rows above which Condense Rows switches itself on the first
# time - see StudioSettingsWindow._auto_condense_if_large().
AUTO_CONDENSE_ABOVE = 60
# px band around the selection's outline where a drag moves the content
# instead of extending the selection - same distinction Excel makes.
MOVE_EDGE_MARGIN = 5

# Ctrl + mouse wheel zoom, same range/step as asked for: 25% out, 200% in,
# 100% normal, 5% per wheel notch.
ZOOM_MIN = 0.25
ZOOM_MAX = 2.00
ZOOM_STEP = 0.05
ZOOM_DEFAULT = 1.00


# IronPython 2 formats numbers through .NET, and .NET rejects a precision on
# an INTEGER format ("Precision not allowed in integer format specifier"). So
# '{:.0f}'.format(210) - fine in CPython - throws here, and Fit Width did
# exactly that with page_w_mm, which is a plain int. Anything handed to a
# {:.Nf} is coerced with float() first.


def _px(mm):
    """Row height / column width in mm -> laid-out pixels.

    One definition on purpose. The GridLength actually applied and the
    _row_px / _col_px map used for hit-testing, the selection outline and the
    page-break lines must agree exactly; when they were computed separately
    (int() here, max(px, 4) there) a click could resolve to a different cell
    than the one it landed on.
    """
    return max(int(round(mm * SCALE)), MIN_CELL_PX)


def _brush(hexcolor, fallback='#000000'):
    """Hex string -> brush, never raising.

    A colour that cannot be parsed falls back rather than taking the whole
    window down with it: these values come from saved layouts, and one bad
    entry used to stop Studio opening at all with
    "Error initializing window: invalid literal for int() with base 16".
    """
    for candidate in (hexcolor, fallback):
        h = str(candidate or '').lstrip('#')
        if len(h) == 3:
            h = h[0] * 2 + h[1] * 2 + h[2] * 2
        try:
            return _SWM.SolidColorBrush(_SWM.Color.FromRgb(
                int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)))
        except Exception:
            continue
    return _SWM.Brushes.Black


# ============================================================================
# Theme
# ============================================================================
# The XAML gets its colours as DynamicResource brushes injected from
# lib/Snippets/seed43_palette.json. The window chrome drawn here in Python -
# row/column headers, the selection outline, context menus, the Blocks panel
# legend - reads the same palette file directly rather than carrying a second
# set of literals that stops matching the moment the palette is re-exported.
#
# What is NOT themed: anything representing the printed page. The sheet is
# white with black text and grey rules because that is what comes out of the
# printer, not because of the current theme, so studio_blocks.py's colours and
# PAGE_BREAK_COLOR below stay fixed whatever palette is loaded.
#
# Fallbacks are the palette's own current values, so a missing or unreadable
# seed43_palette.json degrades to what this window already looked like rather
# than to black.
_THEME_FALLBACK = {
    'card_bg':             '#2B3340',
    'header_bg':           '#232933',
    'window_bg':           '#3B4553',
    'input_bg':            '#1d1d20',
    'pill_off_border':     '#404553',
    'dropdown_item_hover': '#3f3f41',
    'accent':              '#208A3C',
    'text_primary':        '#FFFFFF',
    'text_muted':          '#9CA3AF',
}

_theme_cache = {}
_theme_brush_cache = {}


def theme(key):
    """Hex string for a palette key. Cached: the palette cannot change while
    the window is open, and _render_grid() asks for these once per header."""
    if key not in _theme_cache:
        value = _THEME_FALLBACK.get(key, '#FFFFFF')
        if _get_color is not None:
            try:
                value = _get_color(_SCRIPT_DIR, key, fallback=value)
            except Exception:
                pass
        _theme_cache[key] = value
    return _theme_cache[key]


def theme_brush(key):
    """Frozen brush for a palette key - one instance shared by every element
    that uses it, rather than a fresh SolidColorBrush per cell."""
    if key not in _theme_brush_cache:
        b = _brush(theme(key))
        try:
            b.Freeze()
        except Exception:
            pass
        _theme_brush_cache[key] = b
    return _theme_brush_cache[key]


# -- Themed context menus ----------------------------------------------------
# Background/Foreground only go so far: MenuItem's stock template draws a
# fixed icon-and-checkmark gutter down the left that stays light whatever
# colours are set. That strip is native chrome, so removing it needs a real
# ControlTemplate - same reason the ComboBox needed one.
def _menu_xaml():
    return (
        '<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" '
        'xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" TargetType="ContextMenu">'
        '<Border Background="{}" BorderBrush="{}" BorderThickness="1" '
        'CornerRadius="4" Padding="4">'
        '<StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Cycle"/>'
        '</Border>'
        '</ControlTemplate>'
    ).format(theme('card_bg'), theme('pill_off_border'))


def _menuitem_xaml():
    return (
        '<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" '
        'xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" TargetType="MenuItem">'
        '<Border x:Name="Bd" Background="Transparent" CornerRadius="3" Padding="12,5">'
        '<ContentPresenter ContentSource="Header" VerticalAlignment="Center"/>'
        '</Border>'
        '<ControlTemplate.Triggers>'
        '<Trigger Property="IsHighlighted" Value="True">'
        '<Setter TargetName="Bd" Property="Background" Value="{}"/>'
        '</Trigger>'
        '<Trigger Property="IsEnabled" Value="False">'
        '<Setter Property="Opacity" Value="0.45"/>'
        '</Trigger>'
        '</ControlTemplate.Triggers>'
        '</ControlTemplate>'
    ).format(theme('dropdown_item_hover'))


def _themed_menu():
    menu = _SWC.ContextMenu()
    menu.Foreground = theme_brush('text_primary')
    try:
        import System.Windows.Markup as _Markup
        menu.Template = _Markup.XamlReader.Parse(_menu_xaml())
    except Exception:
        # Fall back to plain colours if template parsing ever fails - the
        # menu still works, it just keeps the stock gutter.
        menu.Background = theme_brush('card_bg')
        menu.BorderBrush = theme_brush('pill_off_border')
    return menu


def _themed_menu_item(text, handler, dot_color=None, enabled=True):
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
            dot.BorderBrush = theme_brush('text_muted')
            dot.BorderThickness = _SW.Thickness(1)
        else:
            dot.Background = _brush(dot_color)
        lbl = _SWC.TextBlock()
        lbl.Text = text
        lbl.Foreground = theme_brush('text_primary')
        lbl.VerticalAlignment = _SW.VerticalAlignment.Center
        row.Children.Add(dot)
        row.Children.Add(lbl)
        mi.Header = row
    else:
        mi.Header = text
    mi.Foreground = theme_brush('text_primary')
    mi.FontSize = 12
    # The template dims a disabled item, so "can't do that here" reads as
    # greyed-out rather than as a menu entry that does nothing when clicked.
    mi.IsEnabled = bool(enabled)
    try:
        import System.Windows.Markup as _Markup
        mi.Template = _Markup.XamlReader.Parse(_menuitem_xaml())
    except Exception:
        pass
    if handler is not None:
        mi.Click += handler
    return mi


def _themed_separator():
    sep = _SWC.Separator()
    sep.Background = theme_brush('pill_off_border')
    sep.Height = 1
    sep.Margin = _SW.Thickness(6, 4, 6, 4)
    return sep


class StudioSettingsWindow(WPFWindow):

    # Label -> revision field, matching script_create_excel.py's _LABEL_TO_KEY
    # so the canvas and both writers agree on what a meta row means.
    META_LABEL_TO_KEY = {
        'issued by': 'initials', 'initials': 'initials',
        'reason for issue': 'reason', 'method of issue': 'method',
        'document format': 'doc_format', 'paper size': 'paper_size',
    }

    def __init__(self, script_dir=None, meta_rows=None):
        if script_dir is None:
            script_dir = os.path.dirname(os.path.abspath(__file__))
        self._script_dir = script_dir
        # Issue metadata from the pyTransmit window - see _apply_meta_rows().
        self._meta_rows = list(meta_rows or [])
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
        try:
            # _PT_ROOT, not script_dir: icon.png belongs to the pyTransmit
            # pushbutton, and Studio is a folder inside it.
            from Snippets._icons import set_header_icon
            set_header_icon(self, _PT_ROOT)
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
        self._row_px = []               # per GRID row (see _visual_rows)
        self._col_px = []
        self._vrows = []                # [(model_row, item_or_None)]
        self._vspans = {}               # model_row -> (first_grid_row, count)
        self._vdomains = {}             # model_row -> 'sheet'/'recipient'/None

        # -- documentation-table preview options ---------------------------
        # These describe how the sheet list is PREVIEWED, not what the layout
        # is, so they live in studio_config.json rather than in the layout
        # template - a template opened on another machine should not silently
        # re-group that project's sheets.
        #
        # Grouping is pyTransmit's setting, owned by its Sheet Parameters
        # panel; Studio seeds itself from there so it opens showing what will
        # actually be published, and treats a change made here as a preview
        # override that is never written back.
        self._group_params = []
        self._group_label = True
        self._condense = False

        self._grid = self._load_last_or_default()
        self._normalise_repeat_row_heights()
        self._load_preview_options()
        self._data = studio_live_data.empty_data()
        self._refresh_live_data(silent=True)
        self._auto_condense_if_large()

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
        self._wire_ribbon_tabs()
        self._wire_home_ribbon()
        self._wire_doc_table_ribbon()
        self._wire_header_buttons()
        self._wire_blocks_panel()
        self._wire_templates()
        self._wire_logos()

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
        if getattr(self, '_startup_notice', None):
            self.status_label.Text = self._startup_notice
        # ContentRendered, not __init__: the notice is about this window, so
        # it should appear over a drawn Studio rather than in front of an
        # empty rectangle. Fires once - it is not re-raised on redraws.
        self.ContentRendered += self._on_first_shown

    def _on_first_shown(self, sender, args):
        self.ContentRendered -= self._on_first_shown
        _alert(
            'pyTransmit Studio is under development.\n\n'
            'It is safe to use - it only ever reads the Revit model, and '
            'layouts are saved as templates of their own - but expect rough '
            'edges, and check anything you publish from a new layout before '
            'sending it out.',
            title='pyTransmit Studio - under development')

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

    def _normalise_repeat_row_heights(self):
        """Bring a pre-expansion layout's list rows down to per-item height.

        Before rows were expanded, a row holding Sheet Number had to be tall
        enough for the WHOLE list stacked inside one cell - 30mm or more. That
        same number now means the height of ONE sheet's row, so an old
        template would open as a sheet tens of metres long. Any repeating row
        noticeably taller than one line of its own font is therefore reset to
        that natural height. Rows already at a sensible height are untouched,
        so this is a no-op for anything saved since.

        Returns how many rows it changed, for the status bar.
        """
        grid = self._grid
        row_blocks = {}
        for (r, c), cell in grid.cells.items():
            if 'covered_by' in cell:
                continue
            row_blocks.setdefault(r, []).append(cell.get('block'))
        heights, fixed = studio_rows.normalise_row_heights(
            grid.row_heights, row_blocks)
        if fixed:
            grid.row_heights = heights
        return fixed

    def _read_config(self):
        try:
            with open(self._last_file_marker(), 'r') as f:
                cfg = json.load(f)
            return cfg if isinstance(cfg, dict) else {}
        except Exception:
            return {}

    def _write_config(self, **kw):
        """Merge keys into studio_config.json. Read-modify-write rather than
        overwrite, so remembering the last file cannot drop the preview
        options and vice versa."""
        cfg = self._read_config()
        cfg.update(kw)
        try:
            with open(self._last_file_marker(), 'w') as f:
                json.dump(cfg, f, indent=2)
        except Exception:
            pass

    def _remember_last_file(self, path):
        self._write_config(last_file=path)

    # -- documentation-table preview options -----------------------------------
    def _load_preview_options(self):
        """Studio's own saved options, falling back to pyTransmit's setup.

        First run has nothing saved, and defaulting to "no grouping" would
        show a preview that disagrees with what pyTransmit would publish, so
        the fallback reads pyTransmit's own files: group_params from
        pytransmit_setup.json (written by SetupSettings) and group_label_on
        from pytransmit_sync.json (written by the Text On/Off toggle).
        """
        cfg = self._read_config()
        self._condense = bool(cfg.get('condense_rows', False))
        # Whether the user has ever set this themselves - see
        # _auto_condense_if_large().
        self._condense_is_default = 'condense_rows' not in cfg

        if 'group_params' in cfg:
            self._group_params = list(cfg.get('group_params') or [])
        else:
            self._group_params = list(self._read_json(SETUP_FILE).get('group_params') or [])

        if 'group_label' in cfg:
            self._group_label = bool(cfg.get('group_label'))
        else:
            self._group_label = bool(self._read_json(SYNC_FILE).get('group_label_on', True))

    def _read_json(self, path):
        try:
            with open(path, 'r') as f:
                d = json.load(f)
            return d if isinstance(d, dict) else {}
        except Exception:
            return {}

    def _save_preview_options(self):
        self._condense_is_default = False
        self._write_config(group_params=list(self._group_params),
                           group_label=bool(self._group_label),
                           condense_rows=bool(self._condense))

    def _auto_condense_if_large(self):
        """Switch Condense Rows on by itself the first time a big model is
        opened.

        Every sheet now gets a real grid row, so a 2000-sheet transmittal is
        2000 rows of WPF elements: the window would spend a long time drawing
        a layout the user cannot work with anyway, before they had any chance
        to reach the toggle. Only applies while the setting is still at its
        default - once the user picks either way, their choice sticks whatever
        the model size.
        """
        self._startup_notice = None
        if self._condense or not self._condense_is_default:
            return
        total = studio_blocks.plan_summary(
            self._data, self._group_params, self._group_label, False)[1]
        if total > AUTO_CONDENSE_ABOVE:
            self._condense = True
            self._startup_notice = (
                'Condense Rows switched on - {} sheet rows would be slow to '
                'draw. Turn it off on the Sheet Data tab.'.format(total))

    def _row_plan(self):
        """The one row plan every per-sheet block in this render shares."""
        first_gap, between_gap = self._grid.gaps_for(self._group_label)
        return studio_blocks.sheet_row_plan(
            self._data, self._group_params, self._group_label, self._condense,
            space_first_group=first_gap, space_between_groups=between_gap)

    def _on_space_first_click(self, sender, args):
        # Edits the pair for the group-text state currently in force, so the
        # template can say "gap when the name is hidden, none when it shows".
        on = bool(self.space_first_btn.IsChecked)
        self._grid.set_gap(self._group_label, 'first', on)
        self._render_all()
        self.status_label.Text = '{} above the first group header (group text {})'.format(
            'Blank row' if on else 'No gap', 'on' if self._group_label else 'off')

    def _on_space_between_click(self, sender, args):
        on = bool(self.space_between_btn.IsChecked)
        self._grid.set_gap(self._group_label, 'between', on)
        self._render_all()
        self.status_label.Text = '{} above every later group header (group text {})'.format(
            'Blank row' if on else 'No gap', 'on' if self._group_label else 'off')

    def new_click(self, sender, args):
        if not _confirm('Start a new blank layout? Unsaved changes will be lost.'):
            return
        self._active_path = None
        self._grid = studio_grid.default_grid()
        self._normalise_repeat_row_heights()
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

    def _confirm_by_name(self, name, what):
        """Delete confirmation that asks the user to type the name back.

        Same wording and rule as LayoutSettings.del_template_click(). Falls
        back to a plain Yes/No only if the string prompt is unavailable, so
        the guard degrades rather than disappearing.
        """
        if sdlg is None:
            return _confirm('Delete the {} "{}"? This cannot be undone.'.format(what, name))
        try:
            typed = sdlg.ask_string(
                'To delete, type the {} name exactly:\n"{}"'.format(what, name),
                title='Delete {}'.format(what), default='')
        except Exception:
            return _confirm('Delete the {} "{}"? This cannot be undone.'.format(what, name))
        if not typed or typed.strip() != name:
            self.status_label.Text = 'Delete cancelled - name did not match'
            return False
        return True

    # ======================================================================
    # Logo library (ribbon: Layout tab)
    # ======================================================================
    # Logos are COPIED into the user's Logos folder rather than referenced
    # where they were found. A layout stores the path of the logo it uses, so
    # a file left on someone's Desktop would break the layout for anyone else
    # - and for the same person once they tidied up.

    NO_LOGO = '(No logo)'
    LOGO_EXTS = ('.png', '.jpg', '.jpeg', '.bmp', '.gif')
    # Where each logo was loaded from, so a logo kept on a server can be
    # re-pulled when it changes. Sits beside the images, not in the layout,
    # because the library is shared by every layout.
    LOGO_SOURCES_FILE = 'logo_sources.json'

    def _logo_names(self):
        try:
            return sorted(f for f in os.listdir(LOGOS_DIR)
                          if os.path.splitext(f)[1].lower() in self.LOGO_EXTS)
        except Exception:
            return []

    def _logo_path(self, name):
        return os.path.join(LOGOS_DIR, name)

    def _logo_sources(self):
        return self._read_json(os.path.join(LOGOS_DIR, self.LOGO_SOURCES_FILE))

    def _remember_logo_source(self, name, source):
        sources = self._logo_sources()
        sources[name] = source
        try:
            with open(os.path.join(LOGOS_DIR, self.LOGO_SOURCES_FILE), 'w') as f:
                json.dump(sources, f, indent=2)
        except Exception:
            pass

    def _forget_logo_source(self, name):
        sources = self._logo_sources()
        if sources.pop(name, None) is None:
            return
        try:
            with open(os.path.join(LOGOS_DIR, self.LOGO_SOURCES_FILE), 'w') as f:
                json.dump(sources, f, indent=2)
        except Exception:
            pass

    # Timestamp slack when comparing a cached logo with its source. copy2
    # carries the source's mtime across verbatim, so an untouched source
    # should compare exactly equal and this only absorbs filesystem rounding.
    # Deliberately tight: a wider window silently misses a source edited soon
    # after it was last pulled. Like every mtime-based sync, a same-size edit
    # made within this window is not detected; the Refresh Live Data button is
    # the way out of that.
    LOGO_MTIME_SLACK_S = 1.0

    @classmethod
    def _same_file(cls, a, b):
        """Is the cached copy still identical to its source? Size plus
        timestamp - the usual test, and all that is available without
        reading both files on every open."""
        try:
            sa, sb = os.stat(a), os.stat(b)
        except Exception:
            return False
        return (sa.st_size == sb.st_size
                and abs(sa.st_mtime - sb.st_mtime) <= cls.LOGO_MTIME_SLACK_S)

    def _sync_logos_from_source(self):
        """Re-pull any logo whose source has changed, same idea as Branding's
        logo_source sync.

        The copy in the user's Logos folder is a cache, not the original: a
        logo kept on a server is re-fetched when it changes there, so nobody
        has to re-load it by hand. If the source is missing - server down, VPN
        off, file moved - the local copy stands and nothing is said about it.
        That is the whole point of caching it: the layout has to keep working
        away from the office.

        Returns (refreshed, offline) counts for the status bar.
        """
        refreshed = offline = 0
        for name, source in self._logo_sources().items():
            local = self._logo_path(name)
            if not os.path.isfile(local):
                # Cache gone (deleted from the library) - leave it deleted
                # rather than resurrecting it from the source.
                continue
            if not source or not os.path.isfile(source):
                offline += 1
                continue
            if self._same_file(source, local):
                continue
            try:
                shutil.copy2(source, local)
                refreshed += 1
            except Exception:
                offline += 1
        return refreshed, offline

    def _wire_logos(self):
        if not os.path.isdir(LOGOS_DIR):
            try:
                os.makedirs(LOGOS_DIR)
            except Exception:
                pass
        refreshed, _offline = self._sync_logos_from_source()
        # Don't tread on the auto-condense notice, which matters more.
        if refreshed and not getattr(self, '_startup_notice', None):
            self._startup_notice = '{} logo(s) updated from their source'.format(refreshed)
        self._reload_logo_combo()
        self.logo_combo.SelectionChanged += self._on_logo_changed
        self.logo_load_btn.Click += self._on_logo_load
        self.logo_delete_btn.Click += self._on_logo_delete

    def _reload_logo_combo(self):
        self._syncing_ui = True
        try:
            self.logo_combo.Items.Clear()
            self.logo_combo.Items.Add(self.NO_LOGO)
            for name in self._logo_names():
                self.logo_combo.Items.Add(name)
        except Exception:
            pass
        finally:
            self._syncing_ui = False
        self._sync_logo_combo()

    def _sync_logo_combo(self):
        """Select whichever library entry the layout's logo_path points at."""
        self._syncing_ui = True
        try:
            current = os.path.basename(self._grid.logo_path or '')
            target = current if current in list(self.logo_combo.Items) else self.NO_LOGO
            self.logo_combo.SelectedItem = target
            self.logo_delete_btn.IsEnabled = (target != self.NO_LOGO)
        except Exception:
            pass
        finally:
            self._syncing_ui = False

    def _set_logo(self, name):
        self._grid.logo_path = '' if name == self.NO_LOGO else self._logo_path(name)
        self._render_all()

    def _on_logo_changed(self, sender, args):
        if self._syncing_ui:
            return
        name = self.logo_combo.SelectedItem
        if name is None:
            return
        self._set_logo(name)
        self.status_label.Text = ('Logo cleared' if name == self.NO_LOGO
                                  else 'Logo set to "{}"'.format(name))

    def _on_logo_load(self, sender, args):
        dlg = _WF.OpenFileDialog()
        dlg.Title = 'Select a logo image'
        dlg.Filter = ('Image files (*.png;*.jpg;*.jpeg;*.bmp;*.gif)'
                      '|*.png;*.jpg;*.jpeg;*.bmp;*.gif|All files (*.*)|*.*')
        if dlg.ShowDialog() != _WF.DialogResult.OK:
            return
        source = dlg.FileName
        name = os.path.basename(source)
        if os.path.splitext(name)[1].lower() not in self.LOGO_EXTS:
            _alert('That is not an image file Studio can show.\n\n'
                   'Supported: {}'.format(', '.join(self.LOGO_EXTS)))
            return
        target = self._logo_path(name)
        # Picking a file that is already the library copy is a no-op, not a
        # replace - copying a file onto itself would truncate it.
        already_in_library = (os.path.abspath(source) == os.path.abspath(target))
        if not already_in_library:
            # Same name, different file: ask, rather than quietly swapping a
            # logo that other saved layouts are pointing at.
            if os.path.isfile(target) and not _confirm(
                    'A logo called "{}" is already in the library. Replace it?\n\n'
                    'Any layout using that name will show the new image.'.format(name)):
                return
            try:
                # copy2, not copyfile: it carries the timestamp across, which
                # is what _same_file() later compares to decide whether the
                # source has changed.
                shutil.copy2(source, target)
            except Exception as e:
                _alert('Could not add that logo:\n{}'.format(e))
                return
        # Remember where it came from, so a logo kept on a server is re-pulled
        # whenever it changes there - and still works when it is unreachable.
        self._remember_logo_source(name, source)
        self._reload_logo_combo()
        self._set_logo(name)
        self._sync_logo_combo()
        on_server = not os.path.abspath(source).lower().startswith(
            os.path.abspath(LOGOS_DIR).lower())
        self.status_label.Text = (
            'Logo "{}" added - kept in step with {}'.format(name, source)
            if on_server else 'Logo "{}" added to the library'.format(name))

    def _on_logo_delete(self, sender, args):
        name = self.logo_combo.SelectedItem
        if not name or name == self.NO_LOGO:
            return
        if not self._confirm_by_name(name, 'logo'):
            return
        try:
            os.remove(self._logo_path(name))
        except Exception as e:
            _alert('Could not delete that logo:\n{}'.format(e))
            return
        # Drop the source too, so the next open does not treat the deletion
        # as a stale cache and pull it straight back in.
        self._forget_logo_source(name)
        # Layouts referencing it are left alone on disk; this one drops it so
        # the canvas stops showing a file that no longer exists.
        if os.path.basename(self._grid.logo_path or '') == name:
            self._grid.logo_path = ''
        self._reload_logo_combo()
        self._render_all()
        self.status_label.Text = 'Deleted logo "{}"'.format(name)

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
            rescaled = self._normalise_repeat_row_heights()
            self._active_path = path
            self._sel = (0, 0, 0, 0)
            self._remember_last_file(path)
            self._render_all()
            self.status_label.Text = (
                'Loaded template "{}"'.format(name) if not rescaled else
                'Loaded template "{}" - {} list row(s) resized to one row per '
                'item; Save to keep'.format(name, rescaled))
        except Exception as e:
            _alert('Could not load template "{}":\n{}'.format(name, e))

    def _on_template_delete(self, sender, args):
        name = self.template_combo.SelectedItem
        if not name:
            return
        # Typing the name, not a Yes/No - the same bar the Layout Builder
        # sets for deleting one of its templates. A template can be a lot of
        # work and there is no undo, so a mis-click must not be enough.
        if not self._confirm_by_name(name, 'template'):
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
        # The obvious "re-read everything" button, so it re-pulls changed
        # server logos too rather than making the user reopen Studio.
        refreshed, offline = self._sync_logos_from_source()
        if refreshed:
            self._reload_logo_combo()
            self.status_label.Text += ' - {} logo(s) updated'.format(refreshed)
        elif offline:
            self.status_label.Text += ' - {} logo source(s) unreachable, using local'.format(offline)
        self._render_all()

    def _apply_meta_rows(self):
        """Overlay the issue metadata pyTransmit passed in.

        studio_live_data reads Reason / Method / Document Format / Page Size
        out of a revision's IssuedTo tags ([R:] [M:] [F:] [P:]), and those are
        only written once a transmittal has been published - so on a first
        issue they are empty and the canvas showed ghost placeholders for rows
        the user had already filled in. The same override both Excel writers
        apply, so preview and output agree.
        """
        if not self._meta_rows:
            return
        values = {}
        for label, value in self._meta_rows:
            key = self.META_LABEL_TO_KEY.get(str(label).lower().strip())
            if key and str(value or '').strip():
                values[key] = value
        if not values:
            return
        for rev in self._data.get('revisions', []) or []:
            rev.update(values)

    def _refresh_live_data(self, silent=True):
        try:
            self._data = studio_live_data.get_live_data(self._settings_dir)
            self._apply_meta_rows()
            if not silent:
                n_rev = len(self._data.get('revisions', []))
                n_sheets = len(self._data.get('docs', []))
                self.status_label.Text = (
                    'Live data refreshed - {} issued revision(s), {} sheet(s)'.format(
                        n_rev, n_sheets))
        except Exception as e:
            self._data = studio_live_data.empty_data()
            if not silent:
                self.status_label.Text = 'Could not read live Revit data: {}'.format(e)
        if not silent:
            # A different model has a different parameter set, so the grouping
            # dropdown is rebuilt rather than left listing the old one's.
            self._reload_group_param_combo()

    # ======================================================================
    # Documentation-table ribbon (Sheet Grouping + Sheet Rows)
    # ======================================================================

    NO_GROUPING = '(No grouping)'

    def _wire_doc_table_ribbon(self):
        self._reload_group_param_combo()
        self.group_param_combo.SelectionChanged += self._on_group_param_changed
        self.group_label_btn.Click += self._on_group_label_click
        self.condense_btn.Click += self._on_condense_click
        self.space_first_btn.Click += self._on_space_first_click
        self.space_between_btn.Click += self._on_space_between_click
        self._sync_doc_table_ribbon()

    def _reload_group_param_combo(self):
        """Rebuild the dropdown and the display -> parameter-list map behind
        it.

        A map rather than parsing the selected string back into parameters,
        because pyTransmit groups by up to five parameters at once and a
        single dropdown cannot express that. Its saved combination is offered
        as one entry, so the preview can show exactly what will be published,
        with the individual parameters listed after it for trying something
        else.
        """
        options = [(self.NO_GROUPING, [])]

        def _add_combo(params):
            if len(params) > 1:
                label = 'pyTransmit: ' + ' + '.join(params)
                if label not in [lbl for lbl, _p in options]:
                    options.append((label, list(params)))

        _add_combo(self._read_json(SETUP_FILE).get('group_params') or [])
        # The current selection may be a combination pyTransmit has since
        # changed away from - keep it listed so it doesn't silently reset.
        _add_combo(self._group_params)

        names = list(self._data.get('sheet_params') or [])
        # A parameter this model doesn't have (a template carried over from
        # another project) would otherwise drop out of the list and look like
        # grouping had been switched off.
        for name in self._group_params:
            if len(self._group_params) == 1 and name not in names:
                names.append(name)
        for name in names:
            if name not in [lbl for lbl, _p in options]:
                options.append((name, [name]))

        self._group_options = options
        self._syncing_ui = True
        try:
            self.group_param_combo.Items.Clear()
            for label, _params in options:
                self.group_param_combo.Items.Add(label)
        finally:
            self._syncing_ui = False

    def _selected_group_item(self):
        """Which dropdown entry represents the current _group_params."""
        for label, params in getattr(self, '_group_options', []):
            if params == list(self._group_params):
                return label
        return self.NO_GROUPING

    def _on_group_param_changed(self, sender, args):
        if self._syncing_ui:
            return
        choice = self.group_param_combo.SelectedItem
        if choice is None:
            return
        for label, params in getattr(self, '_group_options', []):
            if label == choice:
                self._group_params = list(params)
                break
        else:
            return
        self._save_preview_options()
        self._render_all()
        self.status_label.Text = (
            'Sheet grouping off - one row per sheet' if not self._group_params
            else 'Grouping sheets by {}'.format(' + '.join(self._group_params)))

    def _on_group_label_click(self, sender, args):
        # ToggleButton flips IsChecked before Click fires, so it already
        # holds the requested state.
        self._group_label = bool(self.group_label_btn.IsChecked)
        self._save_preview_options()
        self._render_all()
        self.status_label.Text = (
            'Group headers show the group name'
            if self._group_label else
            'Group headers are blank separator rows (Text Off)')

    def _on_condense_click(self, sender, args):
        self._condense = bool(self.condense_btn.IsChecked)
        self._save_preview_options()
        self._render_all()

    def _sync_doc_table_ribbon(self):
        """Push state onto the three controls and refresh the row counter."""
        self._syncing_ui = True
        try:
            target = self._selected_group_item()
            if target in list(self.group_param_combo.Items):
                self.group_param_combo.SelectedItem = target
            else:
                self.group_param_combo.SelectedItem = self.NO_GROUPING
        except Exception:
            pass
        finally:
            self._syncing_ui = False

        self.group_label_btn.IsChecked = bool(self._group_label)
        _first_gap, _between_gap = self._grid.gaps_for(self._group_label)
        self.space_first_btn.IsChecked = bool(_first_gap)
        self.space_between_btn.IsChecked = bool(_between_gap)
        self.space_first_btn.IsEnabled = bool(self._group_params)
        self.space_between_btn.IsEnabled = bool(self._group_params)
        # The captions say which set is being edited, so it is obvious that
        # flipping Group Text swaps to the other pair rather than losing it.
        _state = 'text on' if self._group_label else 'text off'
        self.space_first_btn.ToolTip = (
            'Blank row above the FIRST group header, with group {}. '
            'The other state keeps its own setting.'.format(_state))
        self.space_between_btn.ToolTip = (
            'Blank row above every LATER group header, with group {}. '
            'The other state keeps its own setting.'.format(_state))
        self.condense_btn.IsChecked = bool(self._condense)
        # Group text only means anything once there are groups to label.
        self.group_label_btn.IsEnabled = bool(self._group_params)
        self._update_rows_count_label()

    def _update_rows_count_label(self):
        shown, total, n_groups = studio_blocks.plan_summary(
            self._data, self._group_params, self._group_label, self._condense)
        if not total:
            self.rows_count_label.Text = 'no sheets'
            self.condense_btn.IsEnabled = False
            return
        # Condensing a table that already fits inside the window is a no-op,
        # so say so rather than leaving a toggle that appears to do nothing.
        self.condense_btn.IsEnabled = total > studio_blocks.CONDENSE_MAX_ROWS
        parts = ['{} of {} rows'.format(shown, total) if shown != total
                 else '{} rows'.format(total)]
        if n_groups:
            parts.append('{} group{}'.format(n_groups, '' if n_groups == 1 else 's'))
        self.rows_count_label.Text = '  ·  '.join(parts)

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
        hdr.Foreground = theme_brush('text_muted')
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
                swatch.BorderBrush = theme_brush('pill_off_border')
                swatch.BorderThickness = _SW.Thickness(1)
            lbl = _SWC.TextBlock()
            lbl.Text = text
            lbl.FontSize = 10
            lbl.TextWrapping = _SW.TextWrapping.Wrap
            lbl.Foreground = theme_brush('text_primary')
            lbl.VerticalAlignment = _SW.VerticalAlignment.Center
            row.Children.Add(swatch)
            row.Children.Add(lbl)
            box.Children.Add(row)

        hint = _SWC.TextBlock()
        hint.Text = 'Right-click a row number to set.'
        hint.FontSize = 9
        hint.FontStyle = _SW.FontStyles.Italic
        hint.TextWrapping = _SW.TextWrapping.Wrap
        hint.Foreground = theme_brush('text_muted')
        hint.Margin = _SW.Thickness(2, 2, 0, 0)
        box.Children.Add(hint)

        sep = _SWC.Border()
        sep.Height = 1
        sep.Background = theme_brush('pill_off_border')
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
            btn.Margin = _SW.Thickness(0, 0, 4, 4)
            # Shape from the lib's primary-button tokens, fill from the
            # block's own group colour. See BlockButtonStyle in the XAML for
            # why the lib style cannot be used as-is. Falls back to the old
            # hand-set geometry if the resource is somehow missing, so the
            # panel still builds.
            try:
                btn.Style = self.FindResource('BlockButtonStyle')
            except Exception:
                btn.Padding = _SW.Thickness(8, 4, 8, 4)
                btn.FontSize = 11
                btn.Cursor = _SWI.Cursors.Hand
            try:
                btn.Background = _brush(studio_blocks.GROUP_COLORS.get(group, '#8A96A8'))
            except Exception:
                pass
            btn.Foreground = theme_brush('text_primary')
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
    # Ribbon tabs
    # ======================================================================

    # (key, tab button, tab panel) - the order they appear in the strip.
    RIBBON_TABS = [
        ('home',   'tab_home_btn',   'tab_home_panel'),
        ('layout', 'tab_layout_btn', 'tab_layout_panel'),
        ('data',   'tab_data_btn',   'tab_data_panel'),
        ('list',   'tab_list_btn',   'tab_list_panel'),
    ]

    # Tabs that only appear when the selection calls for them, the way Excel
    # shows Table Design only while a table cell is selected. They start
    # hidden and _set_tab_available() decides from then on.
    CONTEXTUAL_TABS = ('list',)

    def _wire_ribbon_tabs(self):
        self._tab_available = dict(
            (key, key not in self.CONTEXTUAL_TABS)
            for key, _b, _p in self.RIBBON_TABS)
        for key, btn_name, _panel in self.RIBBON_TABS:
            btn = getattr(self, btn_name)
            btn.Click += self._on_ribbon_tab_click
            btn.Tag = key
        # The tab the user last chose, which is not always the tab on show -
        # a contextual tab drops back to Home when it stops applying, and
        # this is what brings it back when it applies again.
        self._ribbon_tab_pref = self._read_config().get('ribbon_tab', 'home')
        self._show_ribbon_tab(self._ribbon_tab_pref)

    def _on_ribbon_tab_click(self, sender, args):
        # A ToggleButton flips itself on click, so clicking the tab already
        # showing would otherwise switch it off and leave no tab selected.
        self._ribbon_tab_pref = sender.Tag
        self._show_ribbon_tab(sender.Tag)
        self._write_config(ribbon_tab=sender.Tag)

    def _show_ribbon_tab(self, key):
        # A contextual tab that isn't currently offered falls back to Home -
        # including on startup, where the saved tab may be one that this
        # selection has no use for.
        if not self._tab_available.get(key, False):
            key = 'home'
        for tab_key, btn_name, panel_name in self.RIBBON_TABS:
            active = (tab_key == key)
            getattr(self, btn_name).IsChecked = active
            getattr(self, panel_name).Visibility = (
                _SW.Visibility.Visible if active else _SW.Visibility.Collapsed)
        self._ribbon_tab = key

    def _set_tab_available(self, key, available):
        """Show or hide a contextual tab in the strip.

        The saved 'ribbon_tab' is left alone on purpose: a tab that comes and
        goes with the selection should come back selected when the user
        returns to a cell it applies to, rather than having been silently
        forgotten the first time they clicked a title cell.
        """
        available = bool(available)
        if self._tab_available.get(key) == available:
            return
        self._tab_available[key] = available
        for tab_key, btn_name, _panel in self.RIBBON_TABS:
            if tab_key != key:
                continue
            getattr(self, btn_name).Visibility = (
                _SW.Visibility.Visible if available else _SW.Visibility.Collapsed)
        if available:
            # Re-offer it if it is the tab the user last chose.
            if getattr(self, '_ribbon_tab_pref', 'home') == key:
                self._show_ribbon_tab(key)
        elif getattr(self, '_ribbon_tab', 'home') == key:
            self._show_ribbon_tab('home')

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
        self.strike_btn.Click += self._on_strike_click
        self._build_color_button(self.font_color_btn, 'A', '#000000')
        self._build_color_button(self.fill_color_btn, '■', '#FFFFFF')
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

        self._wire_border_pickers()

        self.merge_toggle_btn.Click += self._on_merge_toggle_click
        self._build_merge_menu()
        self.merge_menu_btn.Checked += lambda s, a: setattr(self.merge_popup, 'IsOpen', True)
        self.merge_menu_btn.Unchecked += lambda s, a: setattr(self.merge_popup, 'IsOpen', False)
        self.merge_popup.Closed += lambda s, a: setattr(self.merge_menu_btn, 'IsChecked', False)
        self.band_btn.Click += self._on_band_click
        self.band_color_btn.Click += self._on_band_color_click
        self.group_color_btn.Click += self._on_group_color_click

        self.lock_width_btn.Click += self._on_lock_width_click
        self.auto_row_btn.Click += self._on_auto_row_click
        self.margin_box.LostFocus += self._on_margin_commit
        self.margin_box.Text = str(self._grid.margin_mm)


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
        bar_colour = theme_brush('text_primary')
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
        """Live accent colour straight from the palette JSON. Used for the
        border-preset icons so the edges being previewed read clearly against
        the dark popup instead of disappearing as thin black-on-grey lines."""
        return theme('accent')

    def _build_border_icon(self, preset):
        size = 26
        accent = _brush(self._accent_hex())
        g = _SWC.Grid()
        g.Width = size
        g.Height = size
        bg = _SWC.Border()
        bg.Background = theme_brush('input_bg')
        bg.BorderBrush = theme_brush('pill_off_border')
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

    # One picker per thing that can be ruled: the cell itself on Home, and on
    # the List Rows tab the data rows and the group headers separately. They
    # differ only in which buttons open them and which field they write, so
    # they share one build / open / apply path rather than three copies.
    #
    #   (opener buttons, popup, preset grid, block field)
    BORDER_PICKERS = [
        (('borders_btn', 'borders_menu_btn'), 'borders_popup',
         'borders_grid', 'borders'),
        (('data_borders_btn',), 'data_borders_popup',
         'data_borders_grid', 'borders'),
        (('group_borders_btn',), 'group_borders_popup',
         'group_borders_grid', 'group_borders'),
    ]

    def _wire_border_pickers(self):
        for btn_names, popup_name, grid_name, field in self.BORDER_PICKERS:
            buttons = [getattr(self, n) for n in btn_names]
            popup = getattr(self, popup_name)
            self._build_borders_popup(getattr(self, grid_name), field)
            for btn in buttons:
                # Either half of a split opens the same palette - the glyph is
                # not a separate default action, since there is no sensible
                # "last used border" to repeat.
                btn.Tag = popup_name
                btn.Checked += self._on_borders_open
                btn.Unchecked += self._on_borders_close
            popup.Closed += self._on_borders_popup_closed
            popup.Tag = tuple(btn_names)
        self.group_borders_inherit_btn.Click += self._on_group_borders_inherit_click

    def _picker_buttons(self, popup_name):
        for btn_names, name, _grid, _field in self.BORDER_PICKERS:
            if name == popup_name:
                return [getattr(self, n) for n in btn_names]
        return []

    def _on_borders_open(self, sender, args):
        popup_name = sender.Tag
        # Keep every half of the control in step so it lights up together.
        for b in self._picker_buttons(popup_name):
            if not b.IsChecked:
                b.IsChecked = True
        getattr(self, popup_name).IsOpen = True

    def _on_borders_close(self, sender, args):
        popup_name = sender.Tag
        for b in self._picker_buttons(popup_name):
            if b.IsChecked:
                b.IsChecked = False
        getattr(self, popup_name).IsOpen = False

    def _on_borders_popup_closed(self, sender, args):
        for name in (sender.Tag or ()):
            getattr(self, name).IsChecked = False

    def _build_borders_popup(self, container, field):
        container.Children.Clear()
        seen = set()
        for key, label in self.BORDER_PRESETS:
            if key in seen:
                continue
            seen.add(key)
            btn = _SWC.Button()
            btn.Tag = (key, field)
            btn.Margin = _SW.Thickness(3)
            btn.Padding = _SW.Thickness(3)
            btn.ToolTip = label
            btn.Background = _SWM.Brushes.Transparent
            btn.BorderThickness = _SW.Thickness(0)
            btn.Content = self._build_border_icon(key)
            btn.Cursor = _SWI.Cursors.Hand
            btn.Click += self._on_border_preset_click
            container.Children.Add(btn)

    def _close_border_popups(self):
        for btn_names, popup_name, _grid, _field in self.BORDER_PICKERS:
            getattr(self, popup_name).IsOpen = False
            for name in btn_names:
                getattr(self, name).IsChecked = False

    def _on_border_preset_click(self, sender, args):
        preset, field = sender.Tag
        self._close_border_popups()
        self._apply_border_preset(preset, field)

    def _on_group_borders_inherit_click(self, sender, args):
        """Drop a block's separate group-header rules, so the headers follow
        the data rows again - the state a layout starts in."""
        self._close_border_popups()
        self._apply_to_selection(lambda b: b.__setitem__('group_borders', None))
        self.status_label.Text = 'Group headers follow the data rows again'

    def _apply_border_preset(self, preset, field='borders'):
        """Mirrors Excel's border presets: 'all'/'none' touch every cell in
        the selection; the edge-specific presets ('top'/'outside'/etc.) only
        touch the cells that actually sit on that edge of the selection
        rectangle, not every cell - e.g. 'Top Border' draws one line along
        the top of the whole selection, not a line under every cell.

        field is which set of rules is being edited - the cell's own, or the
        group headers' (see studio_rows.borders_for).
        """
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
            # Group headers start from whatever the data rows say, since that
            # is what they were printing until now - the first edit adjusts
            # the rules on show rather than starting from a blank cell.
            current = block.get(field)
            if current is None:
                current = block.get('borders')
            borders = dict(current or {})
            if preset == 'none':
                borders = {'t': False, 'b': False, 'l': False, 'r': False}
            elif preset == 'all':
                borders = {'t': True, 'b': True, 'l': True, 'r': True}
            else:
                if preset in ('outside', 'top', 'top_bottom') and is_top: borders['t'] = True
                if preset in ('outside', 'bottom', 'top_bottom') and is_bottom: borders['b'] = True
                if preset in ('outside', 'left', 'left_right') and is_left: borders['l'] = True
                if preset in ('outside', 'right', 'left_right') and is_right: borders['r'] = True
            block[field] = borders
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

    def _build_color_button(self, btn, glyph, colour):
        """Glyph over a bar of the colour it applies - Excel's A / highlighter
        buttons. Rebuilt rather than tinted so the swatch is a real element
        the sync pass can update as the selection moves."""
        panel = _SWC.StackPanel()
        panel.VerticalAlignment = _SW.VerticalAlignment.Center
        tb = _SWC.TextBlock()
        tb.Text = glyph
        tb.FontSize = 12
        tb.FontWeight = _SW.FontWeights.Bold
        tb.HorizontalAlignment = _SW.HorizontalAlignment.Center
        tb.Foreground = theme_brush('text_primary')
        bar = _SWC.Border()
        bar.Height = 4
        bar.Width = 16
        bar.CornerRadius = _SW.CornerRadius(1)
        bar.Margin = _SW.Thickness(0, 1, 0, 0)
        bar.Background = _brush(colour, '#000000')
        bar.BorderBrush = theme_brush('pill_off_border')
        bar.BorderThickness = _SW.Thickness(1)
        panel.Children.Add(tb)
        panel.Children.Add(bar)
        btn.Content = panel
        btn.Tag = bar          # the swatch, for _sync_color_buttons()

    def _sync_color_buttons(self, b):
        """Point each swatch at the colour the selected cell actually uses."""
        for btn, key, default in ((self.font_color_btn, 'color', '#000000'),
                                  (self.fill_color_btn, 'bg_color', '#FFFFFF')):
            bar = getattr(btn, 'Tag', None)
            if bar is None:
                continue
            try:
                bar.Background = _brush(b.get(key) or default, default)
            except Exception:
                pass

    def _on_strike_click(self, sender, args):
        cur = bool((self._anchor_block() or {}).get('strike'))
        self._apply_to_selection(lambda blk: blk.__setitem__('strike', not cur))

    def _on_underline_click(self, sender, args):
        cur = bool((self._anchor_block() or {}).get('underline'))
        self._apply_to_selection(lambda b: b.__setitem__('underline', not cur))

    def _win32_owner(self):
        """This window as something WinForms will accept as a dialog owner.

        A WinForms dialog opened with no owner is not modal to the WPF window,
        so the colour picker dropped BEHIND Studio and the whole thing looked
        frozen. Handing it the HWND keeps it in front and modal.
        """
        try:
            from System.Windows.Interop import WindowInteropHelper

            class _Owner(_WF.IWin32Window):
                def __init__(self, hwnd):
                    self._hwnd = hwnd

                @property
                def Handle(self):
                    return self._hwnd

            return _Owner(WindowInteropHelper(self).Handle)
        except Exception:
            return None

    def _pick_color(self, current_hex):
        dlg = _WF.ColorDialog()
        dlg.FullOpen = True     # custom-colour panel open, so any hex is reachable
        try:
            h = (current_hex or '#000000').lstrip('#')
            if len(h) == 3:
                h = h[0] * 2 + h[1] * 2 + h[2] * 2
            dlg.Color = _SD.Color.FromArgb(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
        except Exception:
            pass
        owner = self._win32_owner()
        result = dlg.ShowDialog(owner) if owner is not None else dlg.ShowDialog()
        if result == _WF.DialogResult.OK:
            c = dlg.Color
            # int() around each channel is essential. c.R/G/B are .NET Bytes,
            # and '{:02X}'.format() on a .NET Byte defers to .NET's own
            # formatter, which reads "02X" as a CUSTOM numeric format rather
            # than "two hex digits" - producing text like 'X1' that is not
            # hex at all. That is what raised
            #     invalid literal for int() with base 16: 'X1'
            # the moment the colour was read back to paint a cell.
            return '#{:02X}{:02X}{:02X}'.format(int(c.R), int(c.G), int(c.B))
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


    # -- Row banding and group-header colour -----------------------------------
    # Both live on the block, so they travel with the layout and are read
    # identically by the canvas and by script_create_excel_studio.py.
    DEFAULT_BAND_COLOR = '#F5F7FA'
    DEFAULT_GROUP_COLOR = '#E8E8E8'

    def _on_band_click(self, sender, args):
        """Banded rows on/off. ToggleButton flips IsChecked before Click, so
        it already holds the requested state."""
        on = bool(self.band_btn.IsChecked)
        self._apply_to_selection(lambda b: b.__setitem__('alt_rows', on))
        self.status_label.Text = (
            'Banded rows on - every other row takes the band colour'
            if on else 'Banded rows off')

    def _on_band_color_click(self, sender, args):
        cur = (self._anchor_block() or {}).get('alt_color') or self.DEFAULT_BAND_COLOR
        picked = self._pick_color(cur)
        if not picked:
            return
        # Picking white is a real choice, not "no colour" - it is how you get
        # white / grey / white banding with the band itself left plain.
        self._apply_to_selection(lambda b: b.__setitem__('alt_color', picked))
        self.status_label.Text = 'Band colour set to {}'.format(picked)

    def _on_group_color_click(self, sender, args):
        cur = (self._anchor_block() or {}).get('group_color') or self.DEFAULT_GROUP_COLOR
        picked = self._pick_color(cur)
        if not picked:
            return
        self._apply_to_selection(lambda b: b.__setitem__('group_color', picked))
        self.status_label.Text = 'Sheet-grouping header colour set to {}'.format(picked)

    # -- Merge split button ----------------------------------------------------
    # The four operations Excel offers. "Merge & Center" is the default action
    # on the button face; the rest live under the chevron.
    MERGE_ACTIONS = [
        ('Merge && Center', 'center'),
        ('Merge Across', 'across'),
        ('Merge Cells', 'plain'),
        (None, None),                 # separator
        ('Unmerge Cells', 'unmerge'),
    ]

    def _build_merge_menu(self):
        panel = self.merge_menu_panel
        panel.Children.Clear()
        for label, action in self.MERGE_ACTIONS:
            if label is None:
                panel.Children.Add(_themed_separator())
                continue
            btn = _SWC.Button()
            # && is the XAML escape for a literal ampersand; in a plain string
            # here it would show as typed.
            btn.Content = label.replace('&&', '&')
            btn.Tag = action
            btn.HorizontalContentAlignment = _SW.HorizontalAlignment.Left
            btn.Foreground = theme_brush('text_primary')
            btn.Background = _SWM.Brushes.Transparent
            btn.BorderThickness = _SW.Thickness(0)
            btn.Padding = _SW.Thickness(12, 5, 12, 5)
            btn.FontSize = 12
            btn.Cursor = _SWI.Cursors.Hand
            btn.Click += self._on_merge_menu_click
            panel.Children.Add(btn)

    def _on_merge_menu_click(self, sender, args):
        self.merge_popup.IsOpen = False
        self.merge_menu_btn.IsChecked = False
        self._do_merge(sender.Tag)

    def _do_merge(self, action):
        r0, c0, r1, c1 = self._sel_rect()
        grid = self._grid

        if action == 'unmerge':
            n = grid.unmerge_all(r0, c0, r1, c1)
            self.status_label.Text = ('Unmerged {} cell(s)'.format(n) if n
                                      else 'Nothing merged in the selection')
            self._render_all()
            return

        if action == 'across':
            n = grid.merge_across(r0, c0, r1, c1)
            self.status_label.Text = (
                'Merged {} row(s) across'.format(n) if n else
                'Nothing to merge across - rows already merged, or one column wide')
            self._sel = (r0, c0, r1, c1)
            self._render_all()
            return

        # 'plain' and 'center' both merge the whole rectangle; centring is the
        # only difference, applied after so it lands on the merged block.
        if (r0, c0) == (r1, c1):
            self.status_label.Text = 'Select more than one cell to merge'
            return
        try:
            grid.merge(r0, c0, r1, c1)
        except ValueError as e:
            _alert(str(e))
            return
        self._sel = (r0, c0, r1, c1)
        if action == 'center':
            block = grid.block_at(r0, c0)
            if block is None:
                block = studio_blocks.new_block('text', content='')
                grid.set_block(r0, c0, block)
            block['just'] = 'center'
            block['v_just'] = 'middle'
        self.status_label.Text = ('Merged and centred' if action == 'center'
                                  else 'Merged')
        self._render_all()

    def _on_merge_toggle_click(self, sender, args):
        """The split button's face: Merge & Center on, unmerge off.

        ToggleButton flips IsChecked before Click fires, so it already holds
        the requested new state. Unchecking unmerges *every* merge in the
        selection, not just the anchor's - so a mixed selection (some merged
        cells, some plain) can be flattened in one go.

        Both directions go through _do_merge() so the button face and the
        dropdown cannot drift apart. IsChecked is put back on a no-op, since
        the toggle reflects "is the selection merged", not what was clicked.
        """
        r0, c0, r1, c1 = self._sel_rect()
        if self.merge_toggle_btn.IsChecked:
            if (r0, c0) == (r1, c1) or not self._grid.can_merge(r0, c0, r1, c1):
                self.merge_toggle_btn.IsChecked = False
                self.status_label.Text = (
                    'Select more than one cell to merge' if (r0, c0) == (r1, c1)
                    else 'Selection overlaps an existing merge - unmerge it first')
                return
            self._do_merge('center')
        else:
            self._do_merge('unmerge')

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

    def _ctx_rows_for(self, idx):
        """Which rows a row-header command applies to: the whole selected
        range when the right-clicked row is inside it, otherwise just that
        row - the usual right-click-inside-a-selection convention."""
        r0, _c0, r1, _c1 = self._sel_rect()
        if self._sel_mode == 'row' and r0 <= idx <= r1:
            return r0, r1
        return idx, idx

    def _ctx_move_rows(self, idx, delta):
        r0, r1 = self._ctx_rows_for(idx)
        if not self._grid.move_rows(r0, r1, delta):
            return
        self._sel_mode = 'row'
        self._sel = (r0 + delta, 0, r1 + delta, self._grid.n_cols - 1)
        n = r1 - r0 + 1
        self.status_label.Text = (
            'Moved row {} to row {}'.format(r0 + 1, r0 + delta + 1) if n == 1
            else 'Moved {} rows to rows {}-{}'.format(
                n, r0 + delta + 1, r1 + delta + 1))
        self._render_all()

    def _on_header_ctx_opening(self, sender, args):
        """Swap in a menu built against the selection as it is right now.
        WPF reads ContextMenu after this event, so reassigning it here is
        what makes the menu current rather than as-of-the-last-render."""
        try:
            kind, idx = sender.Tag
            self._select_header_for_menu(kind, idx)
            sender.ContextMenu = self._build_header_context_menu(kind, idx)
        except Exception:
            pass

    def _select_header_for_menu(self, kind, idx):
        """Move the selection to the header being right-clicked.

        Only the LEFT button used to change the selection, so right-clicking a
        different row or column left the green outline sitting on the old one -
        the menu acted on the row you clicked, but the highlight said
        otherwise, which reads as though the wrong row is about to change.

        Right-clicking INSIDE an existing multi-row or multi-column selection
        keeps it, which is the usual convention and the one Move Rows, Row
        Height and the section commands already rely on to act on the run.
        """
        mode = 'col' if kind == 'colhdr' else 'row'
        r0, c0, r1, c1 = self._sel_rect()
        inside = (self._selection_kind() == mode and
                  (c0 <= idx <= c1 if mode == 'col' else r0 <= idx <= r1))
        if inside:
            return
        self._sel_mode = mode
        if mode == 'col':
            self._sel = (0, idx, self._grid.n_rows - 1, idx)
        else:
            self._sel = (idx, 0, idx, self._grid.n_cols - 1)
        self._move_overlay_to_selection()
        self._update_formula_bar()

    def _build_header_context_menu(self, kind, idx):
        menu = _themed_menu()

        def _item(text, handler, dot=None, enabled=True):
            menu.Items.Add(_themed_menu_item(text, handler, dot, enabled))

        if kind == 'colhdr':
            _item('Insert Column Left',  lambda s, a, i=idx: self._ctx_insert_col(i, True))
            _item('Insert Column Right', lambda s, a, i=idx: self._ctx_insert_col(i, False))
            menu.Items.Add(_themed_separator())
            menu.Items.Add(_themed_separator())
            _item('Column Width…', lambda s, a, i=idx: self._ctx_col_width(i))
            _item('Distribute Evenly', lambda s, a: self._on_distribute_click(None, None))
            _item('Fit Columns to Page Width',
                  lambda s, a: self._on_fit_width_click(None, None))
            menu.Items.Add(_themed_separator())
            _item('Delete Column',       lambda s, a, i=idx: self._ctx_delete_col(i))
        else:
            _item('Insert Row Above', lambda s, a, i=idx: self._ctx_insert_row(i, True))
            _item('Insert Row Below', lambda s, a, i=idx: self._ctx_insert_row(i, False))
            menu.Items.Add(_themed_separator())
            # Reorder. Greyed out at the ends of the sheet, and wherever a
            # cell merged across rows would have to be split to make room -
            # see studio_grid.Grid.can_move_rows().
            mr0, mr1 = self._ctx_rows_for(idx)
            _item('Move Row Up' if mr0 == mr1 else 'Move Rows Up',
                  lambda s, a, i=idx: self._ctx_move_rows(i, -1),
                  enabled=self._grid.can_move_rows(mr0, mr1, -1))
            _item('Move Row Down' if mr0 == mr1 else 'Move Rows Down',
                  lambda s, a, i=idx: self._ctx_move_rows(i, 1),
                  enabled=self._grid.can_move_rows(mr0, mr1, 1))
            menu.Items.Add(_themed_separator())
            _item('Row Height…', lambda s, a, i=idx: self._ctx_row_height(i))
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
                float(before), float(after), float(self._grid.page_w_mm),
                float(2 * self._grid.margin_mm)))
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

    def _on_lock_width_click(self, sender, args):
        """Pin the columns to the page guide, and true them up now.

        Turning it on immediately fits the current columns to the printable
        width, so the sheet starts from the guide rather than only holding it
        from the next drag onwards.
        """
        self._grid.lock_width = bool(self.lock_width_btn.IsChecked)
        if self._grid.lock_width:
            # Only reports; turning the lock on does not resize anything. Fit
            # Columns to Page Width, on the column header menu, is the
            # deliberate way to make everything add up.
            over = sum(self._grid.col_widths) - self._grid.printable_w_mm()
            self.status_label.Text = (
                'Width locked at {:.0f}mm - columns cannot grow past the page '
                'guide{}'.format(float(self._grid.printable_w_mm()),
                                 ' (currently {:.2f}mm over)'.format(float(over))
                                 if over > 0.005 else ''))
        else:
            self.status_label.Text = 'Width unlocked - columns may exceed the page guide'
        self._render_all()

    def _on_auto_row_click(self, sender, args):
        """Hand the selected rows back to auto height, or pin them as they are."""
        auto = bool(self.auto_row_btn.IsChecked)
        r0, _c0, r1, _c1 = self._sel_rect()
        for r in range(r0, r1 + 1):
            if 0 <= r < len(self._grid.row_auto):
                self._grid.row_auto[r] = auto
        if auto:
            self._apply_auto_row_heights()
        n = r1 - r0 + 1
        self.status_label.Text = '{} row(s) {}'.format(
            n, 'sized to their content' if auto else 'set to a fixed height')
        self._render_all()

    def _apply_auto_row_heights(self):
        """Recompute every auto row's height from its content.

        Only rows still marked auto are touched, so a height the user typed or
        dragged is never overwritten - that is what makes the auto behaviour
        safe to run on every render.
        """
        grid = self._grid
        row_blocks = {}
        for (r, c), cell in grid.cells.items():
            if 'covered_by' in cell:
                continue
            row_blocks.setdefault(r, []).append(cell.get('block'))
        changed = False
        for r in range(grid.n_rows):
            if r >= len(grid.row_auto) or not grid.row_auto[r]:
                continue
            # A row that repeats is already sized per item by
            # studio_rows.natural_row_height_mm via the normaliser; anything
            # else takes the tallest block's own line height.
            want = grid.auto_row_height_mm(r, row_blocks.get(r, []))
            if abs(grid.row_heights[r] - want) >= 0.01:
                grid.row_heights[r] = want
                changed = True
        return changed

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

    # -- Row height / column width, from the header right-click ---------------
    # These used to be two boxes in the ribbon. They were the only ribbon
    # controls that acted on a row or column rather than on the selected
    # cells, and being write-only they read as though they did nothing. The
    # header right-click already owns everything else about a row or column -
    # insert, delete, move, print section - so the size belongs there too.

    def _prompt_size(self, title, prompt, current, minimum):
        """Ask for a size in mm. Returns a float, or None if cancelled or the
        answer was not a number."""
        if sdlg is None:
            return None
        try:
            typed = sdlg.ask_string(prompt, title=title,
                                    default='{:g}'.format(round(float(current), 2)))
        except Exception:
            return None
        if typed is None or not str(typed).strip():
            return None
        try:
            return max(minimum, float(str(typed).strip()))
        except ValueError:
            self.status_label.Text = '"{}" is not a number'.format(typed)
            return None

    def _ctx_row_height(self, idx):
        r0, r1 = self._ctx_rows_for(idx)
        n = r1 - r0 + 1
        val = self._prompt_size(
            'Row Height',
            'Height in mm for {}:'.format(
                'row {}'.format(r0 + 1) if n == 1 else '{} rows'.format(n)),
            self._grid.row_heights[r0], MIN_ROW_MM)
        if val is None:
            return
        for r in range(r0, r1 + 1):
            self._grid.row_heights[r] = val
            # An explicit height is an explicit choice - stop auto-fitting it.
            if 0 <= r < len(self._grid.row_auto):
                self._grid.row_auto[r] = False
        self.status_label.Text = '{} row(s) set to {:g}mm'.format(n, val)
        self._render_all()

    def _ctx_col_width(self, idx):
        _r0, c0, _r1, c1 = self._sel_rect()
        if not (self._sel_mode == 'col' and c0 <= idx <= c1):
            c0 = c1 = idx
        n = c1 - c0 + 1
        val = self._prompt_size(
            'Column Width',
            'Width in mm for {}:'.format(
                'column {}'.format(_col_letter(c0)) if n == 1
                else '{} columns'.format(n)),
            self._grid.col_widths[c0], MIN_COL_MM)
        if val is None:
            return
        for c in range(c0, c1 + 1):
            self._grid.col_widths[c] = val

        # A typed width is an explicit instruction and is honoured exactly,
        # even with Lock Width on. The lock caps DRAGS, where you are feeling
        # your way and an accidental overshoot is unwelcome; silently capping
        # a number you typed just looks like the box is broken - which is
        # precisely how it read. Going over the guide is reported instead.
        over = sum(self._grid.col_widths) - self._grid.printable_w_mm()
        if self._grid.lock_width and over > 0.005:
            self.status_label.Text = (
                '{} column(s) set to {:g}mm - the layout is now {:.2f}mm past '
                'the {:g}mm page guide. Narrow another column, or use Fit '
                'Columns to Page Width.'.format(
                    n, val, float(over),
                    round(float(self._grid.printable_w_mm()), 2)))
        else:
            self.status_label.Text = '{} column(s) set to {:g}mm'.format(n, val)
        self._render_all()

    # -- Formula bar ---------------------------------------------------------
    def _on_formula_commit(self, sender, args):
        # The formula bar doubles as a DESCRIPTION of the selection: pick a
        # column and it reads "Column B", pick a row and it reads "Row 8 -
        # Repeat at Top of Every Page". Those are labels, not cell contents,
        # and _update_formula_bar marks them read-only. Without this guard the
        # next LostFocus wrote the label straight into the anchor cell - so
        # clicking a column header and then clicking away silently replaced
        # the title of the transmittal with the words "Column B".
        if self.formula_bar_tb.IsReadOnly:
            return
        if self._selection_kind() != 'cell':
            return
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
            self._update_block_chip(None)
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

        self._update_block_chip(block)
        self._update_section_label()
        self._sync_format_buttons(block or {})

    def _update_block_chip(self, block):
        """Name the block occupying the active cell, in its group colour.

        Same wording and colour as the Blocks panel, so a cell can be
        identified from the formula bar without hunting for it in the
        palette: A1 [Text] Transmittal Document.
        """
        try:
            t = (block or {}).get('type')
            if not t:
                self.block_type_chip.Visibility = _SW.Visibility.Collapsed
                return
            name = studio_blocks.TYPE_NAMES.get(t, t)
            if t in ('reason_list', 'method_list'):
                # These appear twice in the palette; say which placement.
                name = '{} ({})'.format(
                    name.split(' (')[0],
                    'Horizontal' if (block or {}).get('list_style') == 'row'
                    else 'Vertical')
            self.block_type_label.Text = '[{}]'.format(name)
            self.block_type_chip.Background = _brush(
                studio_blocks.GROUP_COLORS.get(
                    studio_blocks.TYPE_GROUP.get(t, ''), '#8A96A8'))
            self.block_type_chip.Visibility = _SW.Visibility.Visible
        except Exception:
            pass

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

    @staticmethod
    def _fmt_pt(pt):
        """9.0 -> '9', 6.5 -> '6.5'. Sizes are per cell and often come from a
        converted Layout Builder template, where they are not whole points."""
        return (str(int(round(pt))) if abs(pt - round(pt)) < 0.05
                else '{:.1f}'.format(float(pt)))

    def _show_font_size(self, size_mm):
        """Show the selected cell's ACTUAL size in the ribbon.

        This used to round to the nearest whole point and only select it if it
        happened to be one of the preset sizes - so a 6.5pt cell left the box
        reading whatever it said before, usually 9. The ribbon claimed 9pt on
        cells that were really 6.5pt, which is why per-cell sizes looked like
        they were being ignored on export when they were in fact correct.

        A size that is not one of the presets is added to the list for as long
        as it is selected, the way Excel shows a non-standard size in its own
        box, and removed again when the selection moves on.
        """
        items = self.font_size_combo.Items
        previous = getattr(self, '_adhoc_font_size', None)

        if not size_mm:
            self.font_size_combo.SelectedItem = None
        else:
            txt = self._fmt_pt(studio_blocks.mm_to_pt(size_mm))
            if txt not in list(items):
                items.Insert(0, txt)
                self._adhoc_font_size = txt
            else:
                self._adhoc_font_size = None
            self.font_size_combo.SelectedItem = txt

        # Drop the entry added for the cell we were showing before, so the
        # list does not accumulate one row per odd size the user clicks on.
        if previous is not None and previous != getattr(self, '_adhoc_font_size', None):
            try:
                items.Remove(previous)
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
            self._show_font_size(b.get('size_mm'))
        except Exception:
            pass
        finally:
            self._syncing_ui = False

        self.bold_btn.IsChecked = bool(b.get('bold'))
        self.italic_btn.IsChecked = bool(b.get('italic'))
        self.underline_btn.IsChecked = bool(b.get('underline'))
        self.strike_btn.IsChecked = bool(b.get('strike'))
        self._sync_color_buttons(b)
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

        self.lock_width_btn.IsChecked = bool(self._grid.lock_width)
        r0, _c0, r1, _c1 = self._sel_rect()
        auto = [self._grid.row_auto[r] for r in range(r0, r1 + 1)
                if 0 <= r < len(self._grid.row_auto)]
        # Lit only when every selected row is auto, so a mixed selection reads
        # as "not all auto" and one click makes them all auto.
        self.auto_row_btn.IsChecked = bool(auto) and all(auto)

        self.band_btn.IsChecked = bool(b.get('alt_rows'))
        # The colour buttons stay enabled whatever is selected: they apply to
        # every cell in the selection, and the anchor cell is a poor guide to
        # what the rest of it holds - greying them out on the anchor made the
        # colours unreachable when a whole row was selected.
        self._sync_list_tab()

    def _selection_repeat_domains(self):
        """What the blocks under the selection repeat over - 'sheet',
        'recipient', or neither. A selection can hold both."""
        domains = set()
        for (r, c) in self._sel_origins():
            domain = studio_blocks.repeat_domain(self._grid.block_at(r, c))
            if domain:
                domains.add(domain)
        return domains

    def _sync_list_tab(self):
        """Offer the List Rows tab only where it means something.

        Banding and group headers describe rows that REPEAT - one per sheet,
        one per recipient. On a title cell or a project-info cell there are no
        such rows, so the tab is not shown at all rather than shown doing
        nothing.

        Group headers narrow it further: only the documentation table is
        grouped, so a recipient list gets the tab with that half greyed out.
        """
        domains = self._selection_repeat_domains()
        self._set_tab_available('list', bool(domains))
        grouped = 'sheet' in domains
        for name in ('group_color_btn', 'group_borders_btn'):
            getattr(self, name).IsEnabled = grouped
        self.group_header_caption.Text = (
            'Group Headers' if grouped else 'Group Headers - sheet lists only')

    # ======================================================================
    # Grid rendering
    # ======================================================================

    def _render_all(self):
        try:
            self.margin_box.Text = str(self._grid.margin_mm)
        except Exception:
            pass
        self._apply_auto_row_heights()
        self._render_grid()
        self._update_formula_bar()
        self._sync_doc_table_ribbon()
        self._sync_logo_combo()
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

    # ----------------------------------------------------------------------
    # Model rows vs grid rows
    # ----------------------------------------------------------------------
    # A row holding a repeating block stands for a list, and the grid gives
    # every item its own row: one grid row per sheet, one per recipient. The
    # LAYOUT still stores one row - that is what a template is, and what the
    # published document repeats - so the expansion lives here, at render
    # time, and everything the user manipulates (selection, heights, sections,
    # insert/delete/move) stays in model rows.
    #
    # _vrows  [(model_row, item_or_None)] - one entry per grid row
    # _vspans {model_row: (first_grid_row, count)}
    # _row_px is per GRID row, since that is what hit-testing measures.

    def _visual_rows(self, row_plan):
        grid = self._grid
        n_sheet = len(row_plan)
        n_recipient = len(self._data.get('distribution', []) or [])

        vrows = []
        spans = {}
        domains = {}
        for r in range(grid.n_rows):
            domain = None
            for c in range(grid.n_cols):
                cell = grid.cells.get((r, c))
                if cell is None or 'covered_by' in cell:
                    continue
                d = studio_blocks.repeat_domain(cell.get('block'))
                if d == 'sheet':
                    # A row with both kinds in it is ambiguous; the
                    # documentation table wins, since that is the one whose
                    # length actually varies.
                    domain = 'sheet'
                    break
                if d == 'recipient':
                    domain = 'recipient'
            count = n_sheet if domain == 'sheet' else (
                n_recipient if domain == 'recipient' else 0)
            # A repeating block with nothing to show still gets one row, so
            # its "(no data)" ghost has somewhere to live - a row that
            # vanished entirely would look like the block had been deleted.
            items = list(range(count)) if count else [None]
            spans[r] = (len(vrows), len(items))
            # Only a sheet-driven row's items index the sheet row plan; a
            # recipient row's item is a position in the distribution list, so
            # looking it up in the plan would band the wrong rows.
            domains[r] = domain
            for it in items:
                vrows.append((r, it))
        return vrows, spans, domains

    def _vrow_range(self, r0, r1):
        """Model row range -> (first_grid_row, grid_row_count)."""
        first = self._vspans.get(r0, (r0, 1))[0]
        last_start, last_n = self._vspans.get(r1, (r1, 1))
        return first, (last_start + last_n) - first

    @staticmethod
    def _band_kind(row_plan, item):
        """'group', 'more', or None for an ordinary sheet row."""
        if item is None or not (0 <= item < len(row_plan)):
            return None
        kind = row_plan[item][0]
        return kind if kind in ('group', 'more') else None

    def _model_row_of(self, vrow):
        if not self._vrows:
            return 0
        vrow = max(0, min(vrow, len(self._vrows) - 1))
        return self._vrows[vrow][0]

    def _render_grid(self):
        grid = self._grid
        rev_col_map = self._rev_column_map()
        # Built once per render and handed to every block: the per-sheet
        # blocks sit side by side in one printed table, so they must all draw
        # the same rows (see studio_blocks.sheet_row_plan).
        row_plan = self._row_plan()

        vrows, vspans, vdomains = self._visual_rows(row_plan)
        self._vrows = vrows
        self._vspans = vspans
        self._vdomains = vdomains

        row_px = [_px(grid.row_heights[mr]) for mr, _it in vrows]
        col_px = [_px(w) for w in grid.col_widths]
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
            rd = _SWC.RowDefinition(); rd.Height = _SW.GridLength(h)
            cells_grid.RowDefinitions.Add(rd)
            rd2 = _SWC.RowDefinition(); rd2.Height = _SW.GridLength(h)
            rowhdr_grid.RowDefinitions.Add(rd2)
        for w in col_px:
            cd = _SWC.ColumnDefinition(); cd.Width = _SW.GridLength(w)
            cells_grid.ColumnDefinitions.Add(cd)
            cd2 = _SWC.ColumnDefinition(); cd2.Width = _SW.GridLength(w)
            colhdr_grid.ColumnDefinitions.Add(cd2)

        chr_rd = _SWC.RowDefinition(); chr_rd.Height = _SW.GridLength(HEADER_H)
        colhdr_grid.RowDefinitions.Add(chr_rd)
        rhc_cd = _SWC.ColumnDefinition(); rhc_cd.Width = _SW.GridLength(HEADER_W)
        rowhdr_grid.ColumnDefinitions.Add(rhc_cd)

        # Column headers
        for c in range(grid.n_cols):
            hdr = _SWC.Border()
            hdr.Background = theme_brush('card_bg')
            hdr.BorderBrush = theme_brush('pill_off_border')
            hdr.BorderThickness = _SW.Thickness(0, 0, 1, 1)
            tb = _SWC.TextBlock()
            tb.Text = _col_letter(c)
            tb.Foreground = theme_brush('text_primary')
            tb.FontSize = 10
            tb.HorizontalAlignment = _SW.HorizontalAlignment.Center
            tb.VerticalAlignment = _SW.VerticalAlignment.Center
            hdr.Child = tb
            hdr.Cursor = _SWI.Cursors.Hand
            hdr.Tag = ('colhdr', c)
            # Built fresh on each open, not here: Move Rows Up/Down
            # is enabled or greyed out according to the CURRENT
            # selection, and selecting rows does not re-render.
            hdr.ContextMenu = _SWC.ContextMenu()
            hdr.ContextMenuOpening += self._on_header_ctx_opening
            hdr.MouseLeftButtonDown += self._on_header_down
            hdr.MouseMove += self._on_header_mouse_move
            hdr.MouseLeftButtonUp += self._on_header_up
            _SWC.Grid.SetRow(hdr, 0); _SWC.Grid.SetColumn(hdr, c)
            colhdr_grid.Children.Add(hdr)

        # Row headers
        for r in range(grid.n_rows):
            hdr = _SWC.Border()
            hdr.Background = theme_brush('card_bg')
            section = grid.section_of(r)
            if section != studio_grid.SECTION_BODY:
                # Coloured stripe down the left of the row number marks a
                # repeating band at a glance, without stealing header width.
                hdr.BorderBrush = _brush(studio_grid.SECTION_COLORS.get(section, '#404E62'))
                hdr.BorderThickness = _SW.Thickness(4, 0, 1, 1)
                hdr.ToolTip = studio_grid.SECTION_SHORT.get(section, section)
            else:
                hdr.BorderBrush = theme_brush('pill_off_border')
                hdr.BorderThickness = _SW.Thickness(0, 0, 1, 1)
            first_v, n_v = vspans.get(r, (r, 1))
            tb = _SWC.TextBlock()
            tb.Foreground = theme_brush('text_primary')
            tb.FontSize = 10
            tb.HorizontalAlignment = _SW.HorizontalAlignment.Center
            tb.VerticalAlignment = _SW.VerticalAlignment.Center
            if n_v > 1:
                # One header for the whole run, labelled with how many rows
                # this one template row is currently producing - the number
                # still identifies a row of the LAYOUT, which is what insert,
                # delete, move and Row H all act on.
                tb.Text = '{}\n×{}'.format(r + 1, n_v)
                tb.TextAlignment = _SW.TextAlignment.Center
                tb.LineHeight = 11
                tb.LineStackingStrategy = _SW.LineStackingStrategy.BlockLineHeight
                hdr.ToolTip = 'Row {} - one layout row, drawing {} rows'.format(r + 1, n_v)
            else:
                tb.Text = str(r + 1)
            hdr.Child = tb
            hdr.Cursor = _SWI.Cursors.Hand
            hdr.Tag = ('rowhdr', r)
            # Built fresh on each open, not here: Move Rows Up/Down
            # is enabled or greyed out according to the CURRENT
            # selection, and selecting rows does not re-render.
            hdr.ContextMenu = _SWC.ContextMenu()
            hdr.ContextMenuOpening += self._on_header_ctx_opening
            hdr.MouseLeftButtonDown += self._on_header_down
            hdr.MouseMove += self._on_header_mouse_move
            hdr.MouseLeftButtonUp += self._on_header_up
            _SWC.Grid.SetRow(hdr, first_v); _SWC.Grid.SetColumn(hdr, 0)
            if n_v > 1:
                _SWC.Grid.SetRowSpan(hdr, n_v)
            rowhdr_grid.Children.Add(hdr)

        # Content cells (skip covered cells; draw merge origins with spans)
        for r in range(grid.n_rows):
            for c in range(grid.n_cols):
                cell = grid.cells.get((r, c))
                if cell is None or 'covered_by' in cell:
                    continue
                row_span = cell.get('row_span', 1)
                col_span = cell.get('col_span', 1)
                b = cell.get('block') or {}
                rev_index = self._rev_index_for(b, c, rev_col_map)
                first_v, n_v = vspans.get(r, (r, 1))
                # If the ROW repeats, every column in it repeats - not just
                # the columns holding a list block. A static cell beside the
                # sheet list is a cell on each of those rows, not one tall box
                # spanning them, which is both what the printed table looks
                # like and what script_create_excel_studio.py writes.
                #
                # Tested on the item index rather than on n_v > 1, so a
                # one-sheet project still gets the per-item rendering while
                # an EMPTY list (item None) falls through to the block's own
                # "(no data)" ghost. A block caught inside a vertical merge
                # also falls through - the merge, not the list, decides how
                # many rows it covers.
                repeats = (row_span == 1
                           and self._vrows[first_v][1] is not None)
                if repeats:
                    placements = [(first_v + i, 1, self._vrows[first_v + i][1])
                                  for i in range(n_v)]
                else:
                    last_v, last_n = vspans.get(r + row_span - 1, (r + row_span - 1, 1))
                    placements = [(first_v, (last_v + last_n) - first_v, None)]

                for (v_row, v_span, item) in placements:
                    border = _SWC.Border()
                    border.BorderBrush = theme_brush('pill_off_border')
                    borders = b.get('borders', {}) if b else {}
                    border.BorderThickness = _SW.Thickness(
                        1 if borders.get('l') else 0.4,
                        1 if borders.get('t') else 0.4,
                        1 if borders.get('r') else 0.4,
                        1 if borders.get('b') else 0.4)
                    border.Background = _SWM.Brushes.White
                    # Clip to the cell: a block taller than its row would
                    # otherwise render straight over the rows below, which
                    # looks like the grid has broken when a row is made
                    # shorter.
                    border.ClipToBounds = True
                    # A group header or condensed marker spans the whole
                    # table: the Sheet Number column writes it, every other
                    # column just carries the band, which is how the exporter
                    # merges that row across. Without this a static cell in a
                    # repeating row would repeat its text down the group rows
                    # too, and the preview would stop matching the workbook.
                    band = (self._band_kind(row_plan, item)
                            if repeats and vdomains.get(r) == 'sheet' else None)
                    if band and studio_blocks.repeat_domain(b) is None:
                        border.Background = _brush(
                            '#E8E8E8' if band == 'group' else '#EEF2F7')
                        content = None
                    else:
                        content = studio_blocks.render_block(
                            cell.get('block'), self._data, SCALE, grid.logo_path,
                            rev_index=rev_index, row_plan=row_plan, item=item)
                    if content:
                        border.Child = content
                    # Every repeated Border reports the same model cell, so a
                    # click anywhere down the column selects that one cell of
                    # the layout - which is the thing the user can edit.
                    border.Tag = ('cell', r, c)
                    border.MouseLeftButtonDown += self._on_cell_down
                    border.MouseMove += self._on_cell_move
                    border.MouseLeftButtonUp += self._on_cell_up
                    border.AllowDrop = True
                    border.DragEnter += self._on_cell_drag_over
                    border.DragOver += self._on_cell_drag_over
                    border.Drop += self._on_cell_drop
                    border.ContextMenu = self._build_cell_context_menu()
                    _SWC.Grid.SetRow(border, v_row)
                    _SWC.Grid.SetColumn(border, c)
                    if v_span > 1:
                        _SWC.Grid.SetRowSpan(border, v_span)
                    if col_span > 1:
                        _SWC.Grid.SetColumnSpan(border, col_span)
                    cells_grid.Children.Add(border)

        # Selection overlay - a non-hit-testable Border spanning the selected
        # rectangle, drawn as the last (topmost) child of the cell Grid so
        # its Grid.Row/Column attached properties line up with the cells.
        r0, c0, r1, c1 = self._sel_rect()
        overlay = _SWC.Border()
        overlay.BorderBrush = theme_brush('accent')
        overlay.BorderThickness = _SW.Thickness(2)
        # Same accent at 12% over the white sheet, derived from the palette
        # rather than a second hardcoded copy of the green.
        accent = theme_brush('accent').Color
        overlay.Background = _SWM.SolidColorBrush(
            _SWM.Color.FromArgb(30, accent.R, accent.G, accent.B))
        overlay.IsHitTestVisible = False
        v_first, v_count = self._vrow_range(r0, r1)
        _SWC.Grid.SetRow(overlay, v_first)
        _SWC.Grid.SetColumn(overlay, c0)
        _SWC.Grid.SetRowSpan(overlay, v_count)
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
        line.Stroke = _brush(PAGE_BREAK_COLOR)
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
        # row_px is per GRID row, so the section list has to be expanded the
        # same way - a repeating row's section applies to every row it draws.
        v_sections = [grid.section_of(mr) for mr, _it in self._vrows]
        y_breaks = self._row_break_positions(row_px, v_sections, page_h_px)

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
        v = 0
        acc = 0
        for i, h in enumerate(self._row_px):
            if yi < acc + h:
                v = i
                break
            acc += h
            v = i
        # _row_px counts GRID rows; the caller works in model rows.
        r = self._model_row_of(v)
        r = max(0, min(r, self._grid.n_rows - 1))
        c = max(0, min(c, self._grid.n_cols - 1))
        return r, c

    def _move_overlay_to_selection(self):
        r0, c0, r1, c1 = self._sel_rect()
        v_first, v_count = self._vrow_range(r0, r1)
        _SWC.Grid.SetRow(self._overlay, v_first)
        _SWC.Grid.SetColumn(self._overlay, c0)
        _SWC.Grid.SetRowSpan(self._overlay, v_count)
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
                cur = self._model_row_of(self._index_from_pos(self._row_px, pos.Y))
                self._sel = (anchor, 0, cur, self._grid.n_cols - 1)
            self._move_overlay_to_selection()
            self._update_formula_bar()
            return
        # Hover feedback: only show a resize cursor near the header's far edge
        pos = args.GetPosition(sender)
        if kind == 'colhdr':
            near_edge = self._near_resize_edge(pos.X, sender.ActualWidth)
            sender.Cursor = _SWI.Cursors.SizeWE if near_edge else _SWI.Cursors.Hand
        else:
            near_edge = self._near_resize_edge(pos.Y, sender.ActualHeight)
            sender.Cursor = _SWI.Cursors.SizeNS if near_edge else _SWI.Cursors.Hand

    def _near_resize_edge(self, value, extent):
        """Is this point inside the header's drag-resize band?

        The band is capped at a third of the header, so a short row or narrow
        column still has somewhere to click that means "select" - a fixed 5px
        band covers half of a 3mm row, and every attempt to select it started
        a resize instead.
        """
        if extent <= 0:
            return False
        band = min(HEADER_RESIZE_MARGIN, extent / 3.0)
        return value >= extent - band

    def _on_header_down(self, sender, args):
        self._commit_pending_formula()
        kind, idx = sender.Tag
        pos = args.GetPosition(sender)
        if kind == 'colhdr' and self._near_resize_edge(pos.X, sender.ActualWidth):
            self._start_resize('col', idx, sender)
            return
        if kind == 'rowhdr' and self._near_resize_edge(pos.Y, sender.ActualHeight):
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
            kind, idx = self._resizing
            if kind == 'col' and self._grid.lock_width:
                # Applied on release, not during the drag: rescaling the other
                # columns on every mouse-move would make them jitter under the
                # pointer and the drag impossible to aim.
                group = getattr(self, '_resize_group', None) or [idx]
                if self._grid.clamp_to_width(group):
                    self.status_label.Text = (
                        'Width locked - capped at the {:g}mm page guide'.format(
                            round(float(self._grid.printable_w_mm()), 2)))
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
        """Live resize. Everything that describes the grid's geometry is
        updated together - the model (mm), both Grids' definitions (px) and
        the _row_px / _col_px map - because the selection outline, the
        move-cursor hit band and _hit_cell() all read that map on the very
        next mouse move. Leaving it on the pre-drag sizes was what made the
        selection and the cursor jump to the wrong cell mid-drag.
        """
        kind, idx = self._resizing
        pos = args.GetPosition(self._resize_ref(kind))
        # Every header in the group takes the dragged size, so a selected
        # run of columns/rows resizes together.
        group = getattr(self, '_resize_group', None) or [idx]
        if kind == 'col':
            delta = pos.X - self._resize_start_screen
            new_mm = round(max(MIN_COL_MM, self._resize_start_size + delta / SCALE), 2)
            px = _px(new_mm)
            for i in group:
                self._grid.col_widths[i] = new_mm
                self._col_px[i] = px
                # Header and cell grids are separate now - resize both so
                # they stay aligned during the drag.
                self._colhdr_grid.ColumnDefinitions[i].Width = _SW.GridLength(px)
                self._cells_grid.ColumnDefinitions[i].Width = _SW.GridLength(px)
        else:
            delta = pos.Y - self._resize_start_screen
            new_mm = round(max(MIN_ROW_MM, self._resize_start_size + delta / SCALE), 2)
            px = _px(new_mm)
            for i in group:
                self._grid.row_heights[i] = new_mm
                if 0 <= i < len(self._grid.row_auto):
                    self._grid.row_auto[i] = False
                # The height is per ITEM, so a repeating row applies it to
                # every grid row it currently draws - dragging the header of
                # a 2000-sheet row resizes all 2000 together, the way Excel
                # resizes a selected run.
                first_v, n_v = self._vspans.get(i, (i, 1))
                for v in range(first_v, first_v + n_v):
                    self._row_px[v] = px
                    self._rowhdr_grid.RowDefinitions[v].Height = _SW.GridLength(px)
                    self._cells_grid.RowDefinitions[v].Height = _SW.GridLength(px)
        # The outline spans Grid rows/columns, so it follows the new sizes on
        # its own - but the page-break lines are drawn on a Canvas in absolute
        # pixels and only move on a full re-render, which happens on mouse-up.
        self.status_label.Text = (
            '{} {}s resized to {:.1f}mm'.format(
                len(group), 'column' if kind == 'col' else 'row', float(new_mm))
            if len(group) > 1 else
            '{} {} resized to {:.1f}mm'.format(
                'Column' if kind == 'col' else 'Row',
                _col_letter(idx) if kind == 'col' else idx + 1, float(new_mm)))

    # -- moving cell content by dragging the selection edge --------------------
    def _sel_pixel_bounds(self):
        r0, c0, r1, c1 = self._sel_rect()
        v0, v_count = self._vrow_range(r0, r1)
        x0 = sum(self._col_px[:c0])
        x1 = x0 + sum(self._col_px[c0:c1 + 1])
        y0 = sum(self._row_px[:v0])
        y1 = y0 + sum(self._row_px[v0:v0 + v_count])
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
    def _commit_pending_formula(self):
        """Save an in-progress formula-bar edit before the selection moves.

        LostFocus alone is not enough: the grid cells are Borders, which are
        not focusable, so clicking one never takes focus off the text box -
        no LostFocus fires, and _update_formula_bar then overwrites what was
        typed with the newly selected cell's content. Typing then required
        Enter, and anything else silently discarded the edit.

        Committing here means clicking away saves, exactly like Excel.
        """
        try:
            if self.formula_bar_tb.IsFocused and not self.formula_bar_tb.IsReadOnly:
                self._on_formula_commit(None, None)
                self.grid_scroll.Focus()
        except Exception:
            pass

    def _on_cell_down(self, sender, args):
        self._commit_pending_formula()
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
