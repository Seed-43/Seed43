# -*- coding: utf-8 -*-
"""
pylink_table_builder.py -- pyLink's Table Builder: type a table into an
Excel-style grid instead of linking a spreadsheet.

The window is pyTransmit Studio's frame (TableBuilder.xaml copies its
styles, title bar, ribbon, formula bar and status bar) around a grid that
works like Excel: click to select, drag / Shift+click / Shift+arrows to
select a block, drag across column letters or row numbers for whole
columns or rows, Ctrl+A for all; type to replace the active cell, F2 or a
double-click to edit it, Enter / Tab / arrows to move, Alt+Enter for a new
line, Delete to clear, Ctrl+C / X / V (blocks from and to Excel), Ctrl+Z / Y.

The table is saved as an ordinary .xlsx (pylink_xlsx_write.py) in
.user/pyLink/tables, with a JSON copy of the editor's model beside it, and
pyLink loads it like any other Excel file: Apply makes the schedule
through the same writer, sync works the same way, and "Edit table" on the
card reopens this window on the JSON.

IronPython 2 (pyRevit).
"""
import os
import json
import re

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from pyrevit import forms
from System.Windows import (
    Thickness, GridLength, HorizontalAlignment, VerticalAlignment, TextWrapping,
    FontWeights, FontStyles, TextDecorations, TextAlignment, WindowStartupLocation,
    Clipboard, Visibility
)
from System.Windows.Controls import (
    Grid, ColumnDefinition, RowDefinition, Border, TextBox, TextBlock, Button
)
from System.Windows.Controls.Primitives import Thumb, ToggleButton
from System.Windows.Media import SolidColorBrush, Color, Brushes, FontFamily
from System.Windows.Input import Key, Keyboard, ModifierKeys, Cursors, MouseButtonState

HERE = os.path.dirname(os.path.abspath(__file__))
PYLINK_DIR = os.path.dirname(HERE)
PX_PER_MM = 3.5                  # screen pixels per mm of the finished table, at 100% zoom
MM_PER_PT = 0.352778
DEFAULT_COL_MM = 30.0
DEFAULT_ROW_MM = 7.0
MIN_MM = 3.0
SIZES_PT = (6, 7, 8, 9, 10, 11, 12, 14, 16, 18, 20, 24)
BORDER_PX = {'thin': 1.0, 'medium': 2.0, 'thick': 3.0}
FILLS = ((u'No fill', None), (u'Light grey', '#F2F2F2'), (u'Grey', '#D9D9D9'), (u'Dark grey', '#A6A6A6'),
         (u'Yellow', '#FFF2CC'), (u'Blue', '#DDEBF7'), (u'Green', '#E2EFDA'), (u'Red', '#FCE4D6'))


def _rgb(hexs):
    s = hexs.lstrip('#')
    return Color.FromRgb(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))


def _brush(hexs):
    return SolidColorBrush(_rgb(hexs))


def _col_letter(i):
    s = ''
    i += 1
    while i:
        i, rem = divmod(i - 1, 26)
        s = chr(65 + rem) + s
    return s


# ── MODEL ───────────────────────────────────────────────────────────────────

class TableModel(object):
    """The table: sizes in mm, text sizes in pt, merges as
    (r0, c0, r1, c1) inclusive, cells as {(r, c): dict}. Only a merge's
    top-left cell holds text and style."""

    def __init__(self, n_rows=6, n_cols=4):
        self.name = u''
        self.font = 'Arial'
        self.n_rows, self.n_cols = n_rows, n_cols
        self.col_w_mm = [DEFAULT_COL_MM] * n_cols
        self.row_h_mm = [DEFAULT_ROW_MM] * n_rows
        self.merges = []
        self.cells = {}

    # cells
    def get(self, r, c):
        return self.cells.get((r, c)) or {}

    def cell(self, r, c):
        return self.cells.setdefault((r, c), {})

    def merge_at(self, r, c):
        for m in self.merges:
            if m[0] <= r <= m[2] and m[1] <= c <= m[3]:
                return m
        return None

    def origin(self, r, c):
        m = self.merge_at(r, c)
        return (m[0], m[1]) if m else (r, c)

    def covered(self, r, c):
        return self.origin(r, c) != (r, c)

    # rows / columns
    def _shift(self, axis, at, delta):
        """Move cells and merges at index >= `at` on `axis` (0 rows, 1
        columns) by `delta`; a merge running across `at` grows/shrinks."""
        moved = {}
        for (r, c), v in self.cells.items():
            k = (r, c)[axis]
            if delta < 0 and k == at:
                continue
            if k >= at + (1 if delta < 0 else 0):
                k += delta
            moved[(k, c) if axis == 0 else (r, k)] = v
        self.cells = moved
        out = []
        for m in self.merges:
            a, b = (m[0], m[2]) if axis == 0 else (m[1], m[3])
            if delta > 0:
                if a >= at:
                    a, b = a + delta, b + delta
                elif b >= at:
                    b += delta
            else:
                if a > at:
                    a, b = a - 1, b - 1
                elif b >= at:
                    b -= 1
                if b < a:
                    continue
            m2 = (a, m[1], b, m[3]) if axis == 0 else (m[0], a, m[2], b)
            if m2[0] != m2[2] or m2[1] != m2[3]:
                out.append(m2)
        self.merges = out

    def insert_row(self, at):
        self._shift(0, at, 1)
        self.row_h_mm.insert(at, self.row_h_mm[min(at, self.n_rows - 1)] if self.n_rows else DEFAULT_ROW_MM)
        self.n_rows += 1

    def delete_row(self, at):
        if self.n_rows <= 1:
            return
        self._shift(0, at, -1)
        del self.row_h_mm[at]
        self.n_rows -= 1

    def insert_col(self, at):
        self._shift(1, at, 1)
        self.col_w_mm.insert(at, self.col_w_mm[min(at, self.n_cols - 1)] if self.n_cols else DEFAULT_COL_MM)
        self.n_cols += 1

    def delete_col(self, at):
        if self.n_cols <= 1:
            return
        self._shift(1, at, -1)
        del self.col_w_mm[at]
        self.n_cols -= 1

    def grow_to(self, n_rows, n_cols):
        while self.n_rows < n_rows:
            self.insert_row(self.n_rows)
        while self.n_cols < n_cols:
            self.insert_col(self.n_cols)

    # merges
    def merge(self, r0, c0, r1, c1):
        if r0 == r1 and c0 == c1:
            return
        self.unmerge(r0, c0, r1, c1)
        texts = []
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                t = self.get(r, c).get('text')
                if t:
                    texts.append(t)
                if (r, c) != (r0, c0) and (r, c) in self.cells:
                    self.cells[(r, c)].pop('text', None)
        if texts:
            self.cell(r0, c0)['text'] = u' '.join(texts)
        self.merges.append((r0, c0, r1, c1))

    def unmerge(self, r0, c0, r1, c1):
        self.merges = [m for m in self.merges
                       if m[2] < r0 or m[0] > r1 or m[3] < c0 or m[1] > c1]

    # persistence
    def to_dict(self):
        return {'name': self.name, 'font': self.font, 'n_rows': self.n_rows, 'n_cols': self.n_cols,
                'col_w_mm': self.col_w_mm, 'row_h_mm': self.row_h_mm,
                'merges': [list(m) for m in self.merges],
                'cells': dict(('%d,%d' % k, v) for k, v in self.cells.items() if v)}

    @classmethod
    def from_dict(cls, d):
        m = cls(int(d['n_rows']), int(d['n_cols']))
        m.name = d.get('name', u'')
        m.font = d.get('font', 'Arial')
        m.col_w_mm = [float(x) for x in d['col_w_mm']]
        m.row_h_mm = [float(x) for x in d['row_h_mm']]
        m.merges = [tuple(x) for x in d.get('merges', [])]
        for k, v in d.get('cells', {}).items():
            r, c = k.split(',')
            m.cells[(int(r), int(c))] = dict(v)
        return m


# ── WINDOW ──────────────────────────────────────────────────────────────────

class TableBuilderWindow(forms.WPFWindow):
    """Excel-style editing: cells are display tiles; a click selects, a drag
    or Shift+click selects a block, typing replaces the active cell (F2,
    double-click or the formula bar edit it), Enter/Tab/arrows move."""

    def __init__(self, owner, model, taken_names=(), rename=True):
        forms.WPFWindow.__init__(self, os.path.join(HERE, 'TableBuilder.xaml'))
        try:
            from Snippets.seed43_theme import apply_seed43_palette, apply_seed43_dimensions
            apply_seed43_palette(self, PYLINK_DIR)
            apply_seed43_dimensions(self, PYLINK_DIR)
            from Snippets._icons import set_header_icon
            set_header_icon(self, PYLINK_DIR)
        except Exception:
            pass
        if owner is not None:
            try:
                self.Owner = owner
                self.WindowStartupLocation = WindowStartupLocation.CenterOwner
            except Exception:
                pass
        self.model = model
        self.saved = False
        self._taken = set(n.lower() for n in taken_names)
        self._sel = (0, 0, 0, 0)            # anchor r, c, active r, c
        self._editing = None                # None, 'enter' (typing over) or 'edit' (F2 / double-click)
        self._drag = None                   # None, 'cells', 'cols' or 'rows'
        self._zoom = 1.0
        self._undo, self._redo = [], []
        self._suspend = False
        self._last_border = 'all'
        self._fg = self.TryFindResource('BrushTextPrimary') or _brush('#F4FAFF')
        self._muted = self.TryFindResource('BrushTextMuted') or _brush('#9AA4B2')
        self._hdr_bg = self.TryFindResource('BrushCardBg') or _brush('#323A47')
        self._hdr_line = self.TryFindResource('BrushPillOffBorder') or _brush('#4A5262')
        acc = self.TryFindResource('BrushPrimaryGreen')
        self._accent = acc.Color if acc is not None else _rgb('#208A3C')

        self.name_tb.Text = model.name or u''
        self.name_tb.IsEnabled = rename
        self.name_tb.ToolTip = (u'Also the schedule name and the named range in the .xlsx' if rename
                                else u'Fixed once the table is linked: Duplicate the card for a renamed copy')
        self._wire()
        self.sheet.MouseLeftButtonDown += self._sheet_down
        self.sheet.MouseMove += self._sheet_move
        self.sheet.MouseLeftButtonUp += self._sheet_up
        self.PreviewKeyDown += self._keys
        self.PreviewTextInput += self._text_input
        self._editor = TextBox()
        self._editor.BorderThickness = Thickness(0)
        self._editor.Padding = Thickness(2, 0, 2, 0)
        self._editor.AcceptsReturn = True
        self._editor.Visibility = Visibility.Collapsed
        self._editor.PreviewKeyDown += self._editor_keys
        self._editor.TextChanged += self._editor_changed
        self.rebuild()
        self.Loaded += lambda s, e: self.sheet.Focus()

    # -- setup --
    def _wire(self):
        icon = None
        try:
            from Snippets._icons import make_icon as icon
        except Exception:
            pass
        col = '#F4FAFF'
        for btn, key, txt in ((self.align_left_btn, 'h_align_left', u'L'), (self.align_center_btn, 'h_align_center', u'C'),
                              (self.align_right_btn, 'h_align_right', u'R'), (self.valign_top_btn, 'v_align_top', u'T'),
                              (self.valign_mid_btn, 'v_align_center', u'M'), (self.valign_bot_btn, 'v_align_bottom', u'B')):
            btn.Content = icon(key, size=16, color=col) if icon else txt
        self.fill_color_btn.Content = icon('format_color_fill', size=16, color=col) if icon else u'■'
        for s in SIZES_PT:
            self.font_size_combo.Items.Add(unicode(s))
        self.font_size_combo.SelectionChanged += self._size_changed

        self.bold_btn.Click += lambda s, e: self._toggle('bold')
        self.italic_btn.Click += lambda s, e: self._toggle('italic')
        self.underline_btn.Click += lambda s, e: self._toggle('underline')
        self.wrap_btn.Click += lambda s, e: self._toggle('wrap')
        self.align_left_btn.Click += lambda s, e: self._set('halign', 'left')
        self.align_center_btn.Click += lambda s, e: self._set('halign', 'center')
        self.align_right_btn.Click += lambda s, e: self._set('halign', 'right')
        self.valign_top_btn.Click += lambda s, e: self._set('valign', 'top')
        self.valign_mid_btn.Click += lambda s, e: self._set('valign', 'center')
        self.valign_bot_btn.Click += lambda s, e: self._set('valign', 'bottom')
        self.merge_toggle_btn.Click += lambda s, e: self._merge_toggle()
        self.clear_btn.Click += lambda s, e: self._clear()

        self.fill_color_btn.Click += lambda s, e: setattr(self.fill_popup, 'IsOpen', True)
        for label, value in FILLS:
            b = Button()
            b.Width, b.Height = 26, 22
            b.Margin = Thickness(2)
            b.ToolTip = label
            b.BorderBrush = self._hdr_line
            b.BorderThickness = Thickness(1)
            b.Background = _brush(value) if value else Brushes.White
            if not value:
                b.Content = u'∅'
            b.Click += lambda s, e, v=value: self._pick_fill(v)
            self.fill_grid.Children.Add(b)

        self.borders_btn.Click += lambda s, e: self._borders_main()
        self.borders_menu_btn.Click += lambda s, e: self._open_borders()
        self.borders_popup.Closed += lambda s, e: setattr(self.borders_menu_btn, 'IsChecked', False)
        for key, mode, tip in (('border_all', 'all', u'All borders'), ('border_outside', 'outside', u'Outside borders'),
                               (None, 'thick', u'Thick outside borders'), ('border_top', 'top', u'Top border'),
                               ('border_bottom', 'bottom', u'Bottom border'), ('border_inside', 'inside', u'Inside borders'),
                               ('border_left', 'left', u'Left border'), ('border_right', 'right', u'Right border'),
                               ('border_none', 'none', u'No borders')):
            b = ToggleButton()
            self._style(b, 'RibbonBtn')
            b.Width, b.Height = 34, 30
            b.Margin = Thickness(2)
            b.ToolTip = tip
            if key and icon:
                b.Content = icon(key, size=18, color=col)
            else:
                t = TextBlock()
                t.Text = u'▣'
                t.FontSize = 18
                t.FontWeight = FontWeights.Bold
                b.Content = t
            b.Click += lambda s, e, m=mode: self._border_pick(m)
            self.borders_grid.Children.Add(b)

        self.row_above_btn.Click += lambda s, e: self._insert_row(False)
        self.row_below_btn.Click += lambda s, e: self._insert_row(True)
        self.row_delete_btn.Click += lambda s, e: self._delete_rows()
        self.col_left_btn.Click += lambda s, e: self._insert_col(False)
        self.col_right_btn.Click += lambda s, e: self._insert_col(True)
        self.col_delete_btn.Click += lambda s, e: self._delete_cols()
        for tb, fn in ((self.width_tb, self._apply_width), (self.height_tb, self._apply_height)):
            tb.PreviewKeyDown += lambda s, e, fn=fn: (fn(s.Text), self.sheet.Focus()) if e.Key == Key.Enter else None

        self.tab_home_btn.Click += lambda s, e: self._tab('home')
        self.tab_layout_btn.Click += lambda s, e: self._tab('layout')

        self.formula_bar_tb.PreviewKeyDown += self._formula_keys
        self.formula_bar_tb.LostKeyboardFocus += lambda s, e: self._formula_commit()

        self.undo_btn.Click += lambda s, e: self._undo_step()
        self.redo_btn.Click += lambda s, e: self._redo_step()
        self.save_btn.Click += lambda s, e: self._save()
        self.close_btn.Click += lambda s, e: self.Close()
        self.zoom_in_btn.Click += lambda s, e: self._set_zoom(self._zoom + 0.1)
        self.zoom_out_btn.Click += lambda s, e: self._set_zoom(self._zoom - 0.1)
        self.zoom_reset_btn.Click += lambda s, e: self._set_zoom(1.0)

    def _style(self, ctrl, key):
        st = self.TryFindResource(key)
        if st is not None:
            ctrl.Style = st

    def _tab(self, which):
        home = which == 'home'
        self.tab_home_btn.IsChecked = home
        self.tab_layout_btn.IsChecked = not home
        self.tab_home_panel.Visibility = Visibility.Visible if home else Visibility.Collapsed
        self.tab_layout_panel.Visibility = Visibility.Collapsed if home else Visibility.Visible

    def _open_borders(self):
        self.borders_popup.IsOpen = True

    def _px(self):
        return PX_PER_MM * self._zoom

    def _set_zoom(self, z):
        self._zoom = max(0.5, min(2.5, round(z, 1)))
        self.zoom_label.Text = u'%d%%' % int(round(self._zoom * 100))
        self.rebuild()

    # -- undo --
    def _push(self):
        self._undo.append(self.model.to_dict())
        if len(self._undo) > 200:
            del self._undo[0]
        self._redo = []

    def _restore(self, d):
        name = self.model.name
        self.model.__dict__.update(TableModel.from_dict(d).__dict__)
        self.model.name = name
        self.rebuild()

    def _undo_step(self):
        self._commit()
        if self._undo:
            self._redo.append(self.model.to_dict())
            self._restore(self._undo.pop())

    def _redo_step(self):
        self._commit()
        if self._redo:
            self._undo.append(self.model.to_dict())
            self._restore(self._redo.pop())

    # -- drawing --
    def rebuild(self):
        m = self.model
        px = self._px()
        self._suspend = True
        sh = self.sheet
        sh.Children.Clear()
        sh.ColumnDefinitions.Clear()
        sh.RowDefinitions.Clear()
        self._hdr_w, self._hdr_h = 40.0, 22.0
        cd = ColumnDefinition()
        cd.Width = GridLength(self._hdr_w)
        sh.ColumnDefinitions.Add(cd)
        self._col_defs = []
        for c in range(m.n_cols):
            cd = ColumnDefinition()
            cd.Width = GridLength(m.col_w_mm[c] * px)
            sh.ColumnDefinitions.Add(cd)
            self._col_defs.append(cd)
        rd = RowDefinition()
        rd.Height = GridLength(self._hdr_h)
        sh.RowDefinitions.Add(rd)
        self._row_defs = []
        for r in range(m.n_rows):
            rd = RowDefinition()
            rd.Height = GridLength(m.row_h_mm[r] * px)
            sh.RowDefinitions.Add(rd)
            self._row_defs.append(rd)

        sh.Children.Add(self._header(u'', -1, -1))
        self._col_hdrs, self._row_hdrs = [], []
        for c in range(m.n_cols):
            h = self._header(_col_letter(c), -1, c)
            self._col_hdrs.append(h)
            sh.Children.Add(h)
        for r in range(m.n_rows):
            h = self._header(unicode(r + 1), r, -1)
            self._row_hdrs.append(h)
            sh.Children.Add(h)
        for r in range(m.n_rows):
            for c in range(m.n_cols):
                if not m.covered(r, c):
                    sh.Children.Add(self._cell(r, c))

        a = self._accent
        self._sel_rect = Border()
        self._sel_rect.IsHitTestVisible = False
        self._sel_rect.Background = SolidColorBrush(Color.FromArgb(34, a.R, a.G, a.B))
        self._sel_rect.BorderBrush = SolidColorBrush(a)
        self._sel_rect.BorderThickness = Thickness(1)
        sh.Children.Add(self._sel_rect)
        self._active_rect = Border()
        self._active_rect.IsHitTestVisible = False
        self._active_rect.BorderBrush = SolidColorBrush(a)
        self._active_rect.BorderThickness = Thickness(2)
        sh.Children.Add(self._active_rect)
        if self._editor.Parent is not None:
            self._editor.Parent.Children.Remove(self._editor)
        sh.Children.Add(self._editor)
        self._suspend = False
        self._clamp()
        self._show_sel()

    def _header(self, text, r, c):
        b = Border()
        b.Background = self._hdr_bg
        b.BorderBrush = self._hdr_line
        b.BorderThickness = Thickness(0, 0, 1, 1)
        g = Grid()
        t = TextBlock()
        t.Text = text
        t.FontSize = 11
        t.Foreground = self._muted
        t.HorizontalAlignment = HorizontalAlignment.Center
        t.VerticalAlignment = VerticalAlignment.Center
        g.Children.Add(t)
        b.Child = g
        Grid.SetRow(b, r + 1)
        Grid.SetColumn(b, c + 1)
        if r < 0 and c >= 0:
            th = Thumb()
            th.Width = 6
            th.HorizontalAlignment = HorizontalAlignment.Right
            th.Cursor = Cursors.SizeWE
            th.Opacity = 0.01
            th.DragStarted += lambda s, e: self._push()
            th.DragDelta += lambda s, e, c=c: self._drag_col(c, e.HorizontalChange)
            th.DragCompleted += lambda s, e: self._show_sel()
            g.Children.Add(th)
            b.MouseLeftButtonDown += lambda s, e, c=c: self._header_down(e, 'cols', c)
            b.ToolTip = u'Click or drag across to select columns; drag the right edge to resize'
        elif c < 0 and r >= 0:
            th = Thumb()
            th.Height = 5
            th.VerticalAlignment = VerticalAlignment.Bottom
            th.Cursor = Cursors.SizeNS
            th.Opacity = 0.01
            th.DragStarted += lambda s, e: self._push()
            th.DragDelta += lambda s, e, r=r: self._drag_row(r, e.VerticalChange)
            th.DragCompleted += lambda s, e: self._show_sel()
            g.Children.Add(th)
            b.MouseLeftButtonDown += lambda s, e, r=r: self._header_down(e, 'rows', r)
            b.ToolTip = u'Click or drag down to select rows; drag the bottom edge to resize'
        else:
            b.MouseLeftButtonDown += lambda s, e: self._select_all(e)
            b.ToolTip = u'Select the whole table (Ctrl+A)'
        return b

    def _cell(self, r, c):
        m = self.model
        st = m.get(r, c)
        mg = m.merge_at(r, c)
        r1, c1 = (mg[2], mg[3]) if mg else (r, c)
        z = self._zoom
        grid_line = Border()
        grid_line.BorderBrush = _brush('#D4D8DE')
        grid_line.BorderThickness = Thickness(0, 0, 1, 1)
        line = Border()
        line.BorderBrush = Brushes.Black
        # Each shared edge is drawn once, by the cell above / to the left, at
        # the heavier of the two cells' settings (as Excel shows it); a cell
        # draws its own top / left only along the table's outer edge.
        def side(rr, cc, key):
            o = m.origin(rr, cc)
            return BORDER_PX.get(m.get(o[0], o[1]).get(key), 0)
        bt = BORDER_PX.get(st.get('b_top'), 0) if r == 0 else 0
        bl = BORDER_PX.get(st.get('b_left'), 0) if c == 0 else 0
        bb = max(side(r1, c, 'b_bottom'), side(r1 + 1, c, 'b_top') if r1 + 1 < m.n_rows else 0)
        br = max(side(r, c1, 'b_right'), side(r, c1 + 1, 'b_left') if c1 + 1 < m.n_cols else 0)
        line.BorderThickness = Thickness(bl * z, bt * z, br * z, bb * z)
        fill = st.get('fill')
        line.Background = _brush(fill) if fill else Brushes.White
        t = TextBlock()
        t.Text = st.get('text') or u''
        t.Foreground = Brushes.Black
        t.FontFamily = FontFamily(m.font or 'Arial')
        t.FontSize = float(st.get('size') or 10) * MM_PER_PT * self._px()
        t.FontWeight = FontWeights.Bold if st.get('bold') else FontWeights.Normal
        t.FontStyle = FontStyles.Italic if st.get('italic') else FontStyles.Normal
        if st.get('underline'):
            t.TextDecorations = TextDecorations.Underline
        t.TextWrapping = TextWrapping.Wrap if st.get('wrap') else TextWrapping.NoWrap
        h = st.get('halign') or 'left'
        t.TextAlignment = {'center': TextAlignment.Center, 'right': TextAlignment.Right}.get(h, TextAlignment.Left)
        v = st.get('valign') or 'bottom'
        t.VerticalAlignment = {'top': VerticalAlignment.Top, 'center': VerticalAlignment.Center}.get(v, VerticalAlignment.Bottom)
        t.Margin = Thickness(3, 1, 3, 1)
        line.Child = t
        grid_line.Child = line
        Grid.SetRow(grid_line, r + 1)
        Grid.SetColumn(grid_line, c + 1)
        if mg:
            Grid.SetRowSpan(grid_line, r1 - r + 1)
            Grid.SetColumnSpan(grid_line, c1 - c + 1)
        return grid_line

    # -- selection --
    def _rect(self):
        ar, ac, fr, fc = self._sel
        r0, r1 = min(ar, fr), max(ar, fr)
        c0, c1 = min(ac, fc), max(ac, fc)
        changed = True
        while changed:
            changed = False
            for mg in self.model.merges:
                if not (mg[2] < r0 or mg[0] > r1 or mg[3] < c0 or mg[1] > c1):
                    n = (min(r0, mg[0]), min(c0, mg[1]), max(r1, mg[2]), max(c1, mg[3]))
                    if n != (r0, c0, r1, c1):
                        r0, c0, r1, c1 = n
                        changed = True
        return r0, c0, r1, c1

    def _clamp(self):
        m = self.model
        self._sel = tuple(max(0, min(v, (m.n_rows if i % 2 == 0 else m.n_cols) - 1))
                          for i, v in enumerate(self._sel))

    def _active(self):
        return self.model.origin(self._sel[2], self._sel[3])

    def _place(self, el, r0, c0, r1, c1):
        Grid.SetRow(el, r0 + 1)
        Grid.SetColumn(el, c0 + 1)
        Grid.SetRowSpan(el, r1 - r0 + 1)
        Grid.SetColumnSpan(el, c1 - c0 + 1)

    def _show_sel(self):
        m = self.model
        r0, c0, r1, c1 = self._rect()
        self._place(self._sel_rect, r0, c0, r1, c1)
        ar, ac = self._active()
        mg = m.merge_at(ar, ac)
        self._place(self._active_rect, ar, ac, mg[2] if mg else ar, mg[3] if mg else ac)
        a = self._accent
        for i, h in enumerate(self._col_hdrs):
            h.Background = SolidColorBrush(Color.FromArgb(90, a.R, a.G, a.B)) if c0 <= i <= c1 else self._hdr_bg
        for i, h in enumerate(self._row_hdrs):
            h.Background = SolidColorBrush(Color.FromArgb(90, a.R, a.G, a.B)) if r0 <= i <= r1 else self._hdr_bg
        st = m.get(ar, ac)
        self._suspend = True
        try:
            self.formula_cell_label.Text = (u'%s%d' % (_col_letter(c0), r0 + 1) if (r0, c0, r1, c1) == (ar, ac, ar, ac) or mg and (r0, c0, r1, c1) == mg
                                            else u'%s%d:%s%d' % (_col_letter(c0), r0 + 1, _col_letter(c1), r1 + 1))
            if not self.formula_bar_tb.IsKeyboardFocused:
                self.formula_bar_tb.Text = st.get('text') or u''
            size = int(round(float(st.get('size') or 10)))
            self.font_size_combo.SelectedIndex = SIZES_PT.index(size) if size in SIZES_PT else -1
            self.bold_btn.IsChecked = bool(st.get('bold'))
            self.italic_btn.IsChecked = bool(st.get('italic'))
            self.underline_btn.IsChecked = bool(st.get('underline'))
            self.wrap_btn.IsChecked = bool(st.get('wrap'))
            h = st.get('halign') or 'left'
            v = st.get('valign') or 'bottom'
            self.align_left_btn.IsChecked = h == 'left'
            self.align_center_btn.IsChecked = h == 'center'
            self.align_right_btn.IsChecked = h == 'right'
            self.valign_top_btn.IsChecked = v == 'top'
            self.valign_mid_btn.IsChecked = v == 'center'
            self.valign_bot_btn.IsChecked = v == 'bottom'
            self.merge_toggle_btn.IsChecked = bool(mg) and (r0, c0, r1, c1) == tuple(mg)
            ws = set(round(m.col_w_mm[c], 1) for c in range(c0, c1 + 1))
            hs = set(round(m.row_h_mm[r], 1) for r in range(r0, r1 + 1))
            if not self.width_tb.IsKeyboardFocused:
                self.width_tb.Text = (u'%g' % list(ws)[0]) if len(ws) == 1 else u''
            if not self.height_tb.IsKeyboardFocused:
                self.height_tb.Text = (u'%g' % list(hs)[0]) if len(hs) == 1 else u''
        finally:
            self._suspend = False
        n = (r1 - r0 + 1) * (c1 - c0 + 1)
        self.status_info_label.Text = u'%d x %d table' % (m.n_rows, m.n_cols) + (u'   |   %d cells selected' % n if n > 1 else u'')
        self.undo_btn.IsEnabled = bool(self._undo)
        self.redo_btn.IsEnabled = bool(self._redo)

    def _hit(self, pos, clamp=True):
        """(r, c) under a point on the sheet; -1 for the header strips."""
        m = self.model
        px = self._px()
        x, y = pos.X - self._hdr_w, pos.Y - self._hdr_h
        c = -1 if x < 0 else m.n_cols
        acc = 0.0
        for i in range(m.n_cols):
            acc += m.col_w_mm[i] * px
            if x >= 0 and x < acc:
                c = i
                break
        r = -1 if y < 0 else m.n_rows
        acc = 0.0
        for i in range(m.n_rows):
            acc += m.row_h_mm[i] * px
            if y >= 0 and y < acc:
                r = i
                break
        if clamp:
            r = max(0, min(m.n_rows - 1, r))
            c = max(0, min(m.n_cols - 1, c))
        return r, c

    def _sheet_down(self, sender, e):
        if self._editor.IsKeyboardFocusWithin and e.OriginalSource is not self.sheet and self._editor.IsMouseOver:
            return
        self._commit()
        pos = e.GetPosition(self.sheet)
        if pos.X < self._hdr_w or pos.Y < self._hdr_h:
            return
        r, c = self._hit(pos)
        if Keyboard.Modifiers & ModifierKeys.Shift:
            self._sel = (self._sel[0], self._sel[1], r, c)
        else:
            self._sel = (r, c, r, c)
        self._show_sel()
        self.sheet.Focus()
        if e.ClickCount == 2:
            self._begin_edit(keep=True)
            e.Handled = True
            return
        self._drag = 'cells'
        self.sheet.CaptureMouse()
        e.Handled = True

    def _header_down(self, e, mode, idx):
        self._commit()
        m = self.model
        shift = bool(Keyboard.Modifiers & ModifierKeys.Shift)
        if mode == 'cols':
            a = self._sel[1] if shift else idx
            self._sel = (0, a, m.n_rows - 1, idx)
        else:
            a = self._sel[0] if shift else idx
            self._sel = (a, 0, idx, m.n_cols - 1)
        self._drag = mode
        self._drag_anchor = a
        self._show_sel()
        self.sheet.Focus()
        self.sheet.CaptureMouse()
        e.Handled = True

    def _select_all(self, e):
        self._commit()
        m = self.model
        self._sel = (0, 0, m.n_rows - 1, m.n_cols - 1)
        self._show_sel()
        self.sheet.Focus()
        e.Handled = True

    def _sheet_move(self, sender, e):
        if not self._drag or e.LeftButton != MouseButtonState.Pressed:
            return
        r, c = self._hit(e.GetPosition(self.sheet))
        m = self.model
        if self._drag == 'cells':
            new = (self._sel[0], self._sel[1], r, c)
        elif self._drag == 'cols':
            new = (0, self._drag_anchor, m.n_rows - 1, c)
        else:
            new = (self._drag_anchor, 0, r, m.n_cols - 1)
        if new != self._sel:
            self._sel = new
            self._show_sel()

    def _sheet_up(self, sender, e):
        if self._drag:
            self._drag = None
            self.sheet.ReleaseMouseCapture()

    def _move(self, dr, dc, extend=False, to_edge=False):
        m = self.model
        ar, ac, fr, fc = self._sel
        r, c = fr, fc
        mg = m.merge_at(r, c)
        if mg and not extend:
            r, c = (mg[2] if dr > 0 else mg[0]), (mg[3] if dc > 0 else mg[1])
        if to_edge:
            r = (m.n_rows - 1 if dr > 0 else 0) if dr else r
            c = (m.n_cols - 1 if dc > 0 else 0) if dc else c
        else:
            r, c = r + dr, c + dc
        r = max(0, min(m.n_rows - 1, r))
        c = max(0, min(m.n_cols - 1, c))
        self._sel = (ar, ac, r, c) if extend else (r, c, r, c)
        self._show_sel()
        self._scroll_into_view(r, c)

    def _scroll_into_view(self, r, c):
        m = self.model
        px = self._px()
        x = self._hdr_w + sum(m.col_w_mm[:c]) * px
        y = self._hdr_h + sum(m.row_h_mm[:r]) * px
        sv = self.grid_scroll
        w, h = m.col_w_mm[c] * px, m.row_h_mm[r] * px
        if x < sv.HorizontalOffset + self._hdr_w:
            sv.ScrollToHorizontalOffset(max(0, x - self._hdr_w))
        elif x + w > sv.HorizontalOffset + sv.ViewportWidth:
            sv.ScrollToHorizontalOffset(x + w - sv.ViewportWidth)
        if y < sv.VerticalOffset + self._hdr_h:
            sv.ScrollToVerticalOffset(max(0, y - self._hdr_h))
        elif y + h > sv.VerticalOffset + sv.ViewportHeight:
            sv.ScrollToVerticalOffset(y + h - sv.ViewportHeight)

    # -- in-cell editing --
    def _begin_edit(self, keep, text=None):
        m = self.model
        ar, ac = self._active()
        self._sel = (ar, ac, ar, ac)
        self._show_sel()
        mg = m.merge_at(ar, ac)
        self._place(self._editor, ar, ac, mg[2] if mg else ar, mg[3] if mg else ac)
        st = m.get(ar, ac)
        ed = self._editor
        ed.FontFamily = FontFamily(m.font or 'Arial')
        ed.FontSize = float(st.get('size') or 10) * MM_PER_PT * self._px()
        ed.FontWeight = FontWeights.Bold if st.get('bold') else FontWeights.Normal
        ed.FontStyle = FontStyles.Italic if st.get('italic') else FontStyles.Normal
        ed.TextWrapping = TextWrapping.Wrap if st.get('wrap') else TextWrapping.NoWrap
        ed.Background = _brush(st.get('fill')) if st.get('fill') else Brushes.White
        ed.Foreground = Brushes.Black
        self._suspend = True
        ed.Text = (st.get('text') or u'') if keep else (text or u'')
        self._suspend = False
        self._edit_cell = (ar, ac)
        self._editing = 'edit' if keep else 'enter'
        ed.Visibility = Visibility.Visible
        ed.Focus()
        Keyboard.Focus(ed)
        ed.CaretIndex = len(ed.Text)

    def _editor_changed(self, sender, e):
        if not self._suspend and self._editing:
            self.formula_bar_tb.Text = sender.Text

    def _commit(self):
        if not self._editing:
            return
        r, c = self._edit_cell
        new = self._editor.Text
        if (self.model.get(r, c).get('text') or u'') != new:
            self._push()
            self.model.cell(r, c)['text'] = new
        self._end_edit()
        self.rebuild()

    def _cancel(self):
        self._end_edit()
        self._show_sel()

    def _end_edit(self):
        self._editing = None
        self._editor.Visibility = Visibility.Collapsed
        self.sheet.Focus()

    def _editor_keys(self, sender, e):
        k = e.Key
        mods = Keyboard.Modifiers
        shift = bool(mods & ModifierKeys.Shift)
        if k == Key.System and e.SystemKey == Key.Enter:            # Alt+Enter: new line
            i = sender.CaretIndex
            sender.Text = sender.Text[:i] + u'\n' + sender.Text[i:]
            sender.CaretIndex = i + 1
            e.Handled = True
        elif k == Key.Enter:
            self._commit()
            self._move(-1 if shift else 1, 0)
            e.Handled = True
        elif k == Key.Tab:
            self._commit()
            self._move(0, -1 if shift else 1)
            e.Handled = True
        elif k == Key.Escape:
            self._cancel()
            e.Handled = True
        elif self._editing == 'enter' and k in (Key.Up, Key.Down, Key.Left, Key.Right):
            self._commit()
            self._move({Key.Up: -1, Key.Down: 1}.get(k, 0), {Key.Left: -1, Key.Right: 1}.get(k, 0))
            e.Handled = True
        elif k == Key.F2:
            self._editing = 'edit'
            e.Handled = True

    def _formula_keys(self, sender, e):
        if e.Key == Key.Enter and not (Keyboard.Modifiers & ModifierKeys.Alt):
            self._formula_commit()
            self._move(1, 0)
            self.sheet.Focus()
            e.Handled = True
        elif e.Key == Key.Escape:
            self._show_sel()
            self.sheet.Focus()
            e.Handled = True

    def _formula_commit(self):
        if self._suspend or self._editing:
            return
        r, c = self._active()
        new = self.formula_bar_tb.Text
        if (self.model.get(r, c).get('text') or u'') != new:
            self._push()
            self.model.cell(r, c)['text'] = new
            self.rebuild()

    # -- keyboard on the sheet --
    def _sheet_has_focus(self):
        return self.sheet.IsKeyboardFocused

    def _text_input(self, sender, e):
        if not self._sheet_has_focus() or self._editing:
            return
        if e.Text and e.Text not in (u'\r', u'\n', u'\t', u'\x1b', u'\b') and ord(e.Text[0]) >= 32:
            self._begin_edit(keep=False, text=e.Text)
            e.Handled = True

    def _keys(self, sender, e):
        if not self._sheet_has_focus() or self._editing:
            return
        k = e.Key
        mods = Keyboard.Modifiers
        shift = bool(mods & ModifierKeys.Shift)
        ctrl = bool(mods & ModifierKeys.Control)
        arrows = {Key.Up: (-1, 0), Key.Down: (1, 0), Key.Left: (0, -1), Key.Right: (0, 1)}
        if k in arrows:
            dr, dc = arrows[k]
            self._move(dr, dc, extend=shift, to_edge=ctrl)
        elif k == Key.Tab:
            self._move(0, -1 if shift else 1)
        elif k == Key.Enter:
            self._move(-1 if shift else 1, 0)
        elif k == Key.F2:
            self._begin_edit(keep=True)
        elif k == Key.Delete:
            self._clear()
        elif k == Key.Back:
            self._begin_edit(keep=False, text=u'')
        elif k == Key.Home:
            self._sel = (self._sel[2], 0, self._sel[2], 0) if not ctrl else (0, 0, 0, 0)
            self._show_sel()
        elif ctrl and k == Key.A:
            m = self.model
            self._sel = (0, 0, m.n_rows - 1, m.n_cols - 1)
            self._show_sel()
        elif ctrl and k == Key.C:
            self._copy()
        elif ctrl and k == Key.X:
            self._copy()
            self._clear()
        elif ctrl and k == Key.V:
            self._paste()
        elif ctrl and k == Key.Z:
            self._undo_step()
        elif ctrl and k == Key.Y:
            self._redo_step()
        elif ctrl and k in (Key.B, Key.I, Key.U):
            self._toggle({Key.B: 'bold', Key.I: 'italic', Key.U: 'underline'}[k])
        else:
            return
        e.Handled = True

    # -- editing commands --
    def _cells(self):
        r0, c0, r1, c1 = self._rect()
        return [(r, c) for r in range(r0, r1 + 1) for c in range(c0, c1 + 1) if not self.model.covered(r, c)]

    def _after(self):
        self.rebuild()
        self.sheet.Focus()

    def _toggle(self, key):
        self._commit()
        cells = self._cells()
        self._push()
        on = not all(self.model.get(r, c).get(key) for r, c in cells)
        for r, c in cells:
            self.model.cell(r, c)[key] = on
        self._after()

    def _set(self, key, value):
        self._commit()
        self._push()
        for r, c in self._cells():
            self.model.cell(r, c)[key] = value
        self._after()

    def _size_changed(self, sender, e):
        if self._suspend or sender.SelectedIndex < 0:
            return
        self._set('size', SIZES_PT[sender.SelectedIndex])

    def _pick_fill(self, value):
        self.fill_popup.IsOpen = False
        self._set('fill', value)

    def _border_pick(self, mode):
        self.borders_popup.IsOpen = False
        self._last_border = mode
        self._borders(mode)

    def _borders_main(self):
        self.borders_btn.IsChecked = False
        self._borders(self._last_border)

    def _borders(self, mode):
        self._commit()
        self._push()
        m = self.model
        r0, c0, r1, c1 = self._rect()
        sides = ('b_top', 'b_bottom', 'b_left', 'b_right')
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                cl = m.cell(r, c)
                if mode == 'none':
                    for k in sides:
                        cl.pop(k, None)
                elif mode == 'all':
                    for k in sides:
                        cl[k] = 'thin'
                elif mode in ('outside', 'thick'):
                    w = 'thin' if mode == 'outside' else 'medium'
                    if r == r0:
                        cl['b_top'] = w
                    if r == r1:
                        cl['b_bottom'] = w
                    if c == c0:
                        cl['b_left'] = w
                    if c == c1:
                        cl['b_right'] = w
                elif mode == 'inside':
                    if r > r0:
                        cl['b_top'] = 'thin'
                    if r < r1:
                        cl['b_bottom'] = 'thin'
                    if c > c0:
                        cl['b_left'] = 'thin'
                    if c < c1:
                        cl['b_right'] = 'thin'
                elif mode == 'top' and r == r0:
                    cl['b_top'] = 'thin'
                elif mode == 'bottom' and r == r1:
                    cl['b_bottom'] = 'thin'
                elif mode == 'left' and c == c0:
                    cl['b_left'] = 'thin'
                elif mode == 'right' and c == c1:
                    cl['b_right'] = 'thin'
        self._after()

    def _clear(self):
        self._commit()
        cells = [rc for rc in self._cells() if self.model.get(*rc).get('text')]
        if not cells:
            return
        self._push()
        for r, c in cells:
            self.model.cell(r, c).pop('text', None)
        self._after()

    def _merge_toggle(self):
        self._commit()
        r0, c0, r1, c1 = self._rect()
        m = self.model
        whole = m.merge_at(r0, c0)
        self._push()
        if whole and tuple(whole) == (r0, c0, r1, c1):
            m.unmerge(r0, c0, r1, c1)
        elif r0 == r1 and c0 == c1:
            self._undo.pop()
            self._show_sel()
            return
        else:
            m.merge(r0, c0, r1, c1)
            cl = m.cell(r0, c0)
            cl.setdefault('halign', 'center')
            cl.setdefault('valign', 'center')
        self._after()

    def _insert_row(self, below):
        self._commit()
        self._push()
        r0, c0, r1, c1 = self._rect()
        at = r1 + 1 if below else r0
        self.model.insert_row(at)
        self._sel = (at, c0, at, c0)
        self._after()

    def _delete_rows(self):
        self._commit()
        self._push()
        r0, c0, r1, c1 = self._rect()
        for _ in range(r1 - r0 + 1):
            if self.model.n_rows <= 1:
                break
            self.model.delete_row(r0)
        self._sel = (r0, c0, r0, c0)
        self._after()

    def _insert_col(self, right):
        self._commit()
        self._push()
        r0, c0, r1, c1 = self._rect()
        at = c1 + 1 if right else c0
        self.model.insert_col(at)
        self._sel = (r0, at, r0, at)
        self._after()

    def _delete_cols(self):
        self._commit()
        self._push()
        r0, c0, r1, c1 = self._rect()
        for _ in range(c1 - c0 + 1):
            if self.model.n_cols <= 1:
                break
            self.model.delete_col(c0)
        self._sel = (r0, c0, r0, c0)
        self._after()

    def _num(self, text):
        try:
            v = float(text)
        except Exception:
            return None
        return v if v >= MIN_MM else None

    def _apply_width(self, text):
        v = self._num(text)
        if v is None:
            return
        self._push()
        r0, c0, r1, c1 = self._rect()
        for c in range(c0, c1 + 1):
            self.model.col_w_mm[c] = v
        self._after()

    def _apply_height(self, text):
        v = self._num(text)
        if v is None:
            return
        self._push()
        r0, c0, r1, c1 = self._rect()
        for r in range(r0, r1 + 1):
            self.model.row_h_mm[r] = v
        self._after()

    def _drag_col(self, c, dx):
        v = max(MIN_MM, self.model.col_w_mm[c] + dx / self._px())
        self.model.col_w_mm[c] = round(v, 1)
        self._col_defs[c].Width = GridLength(v * self._px())

    def _drag_row(self, r, dy):
        v = max(MIN_MM, self.model.row_h_mm[r] + dy / self._px())
        self.model.row_h_mm[r] = round(v, 1)
        self._row_defs[r].Height = GridLength(v * self._px())

    # -- clipboard --
    def _copy(self):
        r0, c0, r1, c1 = self._rect()
        lines = []
        for r in range(r0, r1 + 1):
            vals = []
            for c in range(c0, c1 + 1):
                t = self.model.get(r, c).get('text') or u''
                if u'\n' in t or u'\t' in t or u'"' in t:
                    t = u'"' + t.replace(u'"', u'""') + u'"'
                vals.append(t)
            lines.append(u'\t'.join(vals))
        Clipboard.SetText(u'\r\n'.join(lines))
        self.status_label.Text = u'Copied %d x %d' % (r1 - r0 + 1, c1 - c0 + 1)

    def _paste(self):
        """Text copied from Excel (or from here) fills the table from the
        top-left of the selection, adding rows and columns as needed. A
        single value copied into a block selection fills every cell."""
        if not Clipboard.ContainsText():
            return False
        rows = _parse_tsv(Clipboard.GetText().rstrip(u'\r\n'))
        if not rows:
            return False
        self._commit()
        self._push()
        m = self.model
        r0, c0, r1, c1 = self._rect()
        if len(rows) == 1 and len(rows[0]) == 1 and (r0, c0) != (r1, c1):
            for r, c in self._cells():
                m.cell(r, c)['text'] = rows[0][0]
        else:
            n_c = max(len(x) for x in rows)
            m.grow_to(r0 + len(rows), c0 + n_c)
            for i, vals in enumerate(rows):
                for j, v in enumerate(vals):
                    if not m.covered(r0 + i, c0 + j):
                        m.cell(r0 + i, c0 + j)['text'] = v
            self._sel = (r0, c0, r0 + len(rows) - 1, c0 + n_c - 1)
        self._after()
        return True

    # -- save --
    def _save(self):
        self._commit()
        name = (self.name_tb.Text or u'').strip()
        if not name:
            self.status_label.Text = u'Give the table a name first: it becomes the schedule name.'
            self.name_tb.Focus()
            return
        if self.name_tb.IsEnabled and name.lower() in self._taken:
            self.status_label.Text = u'A table called "%s" already exists. Pick another name.' % name
            self.name_tb.Focus()
            return
        if not any((v.get('text') or u'').strip() for v in self.model.cells.values()):
            self.status_label.Text = u'The table is empty.'
            return
        self.model.name = name
        self.saved = True
        self.Close()


def _parse_tsv(text):
    """Excel's clipboard text: rows by line, cells by tab, a cell holding
    a line break or tab wrapped in quotes with "" for a quote."""
    rows, row, cell, i, n, quoted = [], [], [], 0, len(text), False
    at_start = True
    while i < n:
        ch = text[i]
        if quoted:
            if ch == u'"':
                if i + 1 < n and text[i + 1] == u'"':
                    cell.append(u'"')
                    i += 1
                else:
                    quoted = False
            else:
                cell.append(ch)
        elif ch == u'"' and at_start:
            quoted = True
        elif ch == u'\t':
            row.append(u''.join(cell))
            cell, at_start = [], True
            i += 1
            continue
        elif ch in (u'\r', u'\n'):
            if ch == u'\r' and i + 1 < n and text[i + 1] == u'\n':
                i += 1
            row.append(u''.join(cell))
            rows.append(row)
            row, cell, at_start = [], [], True
            i += 1
            continue
        else:
            cell.append(ch)
        at_start = False
        i += 1
    if cell or row:
        row.append(u''.join(cell))
        rows.append(row)
    return rows


# ── ENTRY POINTS (used by pyLink.py) ────────────────────────────────────────

def tables_dir():
    from Snippets import _userdata
    return _userdata.user_dir('pyLink', 'tables')


def _safe_file(name):
    s = re.sub(r'[\\/:*?"<>|]+', '_', name).strip().rstrip('.')
    return s or 'Table'


def model_path(xlsx_path):
    return os.path.splitext(xlsx_path)[0] + '.pylinktable.json'


def is_built_table(xlsx_path):
    return bool(xlsx_path) and os.path.exists(model_path(xlsx_path))


def existing_names():
    d = tables_dir()
    out = []
    for f in os.listdir(d):
        if f.lower().endswith('.pylinktable.json'):
            try:
                with open(os.path.join(d, f)) as fh:
                    out.append(json.load(fh).get('name') or f)
            except Exception:
                out.append(f)
    return out


def _write(model, xlsx_path):
    import sys
    here = os.path.dirname(os.path.abspath(__file__))
    fmt = os.path.join(here, 'format')
    if fmt not in sys.path:
        sys.path.insert(0, fmt)
    import pylink_xlsx_write
    d = model.to_dict()
    rname = pylink_xlsx_write.write_table_xlsx(xlsx_path, d)
    with open(model_path(xlsx_path), 'w') as fh:
        json.dump(d, fh, indent=1)
    return rname


def build_new(owner):
    """Open an empty table. Returns (xlsx path, named range, name) once it
    is saved, or None."""
    win = TableBuilderWindow(owner, TableModel(), existing_names(), rename=True)
    win.ShowDialog()
    if not win.saved:
        return None
    m = win.model
    path = os.path.join(tables_dir(), _safe_file(m.name) + '.xlsx')
    k = 2
    while os.path.exists(path):
        path = os.path.join(tables_dir(), '%s (%d).xlsx' % (_safe_file(m.name), k))
        k += 1
    return path, _write(m, path), m.name


def edit_existing(owner, xlsx_path):
    """Reopen a built table on its saved model and rewrite its .xlsx.
    True when saved."""
    with open(model_path(xlsx_path)) as fh:
        m = TableModel.from_dict(json.load(fh))
    win = TableBuilderWindow(owner, m, (), rename=False)
    win.ShowDialog()
    if not win.saved:
        return False
    _write(win.model, xlsx_path)
    return True
