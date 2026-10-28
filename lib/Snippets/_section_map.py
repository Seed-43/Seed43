# -*- coding: utf-8 -*-
"""Section Map: the window that ties steel sections to Revit family types.

    from Snippets._section_map import ask_section_map, edit_section_map

Two uses of one window (SectionMap.xaml, next to this file):

  ask_section_map   during an import, for the sections that found no type:
                    pick a type for each (type to search) or import them as
                    shapes, optionally remembering the choices.
  edit_section_map  from 3D Import's "Section map..." button: the member
                    sections of the file chosen there, with their type now
                    (saved or automatic), to change or reset.

The rows are built in Python after the Seed43 palette is applied (load
order matters, see the seed43-pyrevit-ui notes), from styles the XAML
declares. IronPython 2.
"""
import os

from pyrevit import forms

from Snippets.seed43_theme import apply_seed43_dimensions, apply_seed43_palette

__all__ = ["ask_section_map", "edit_section_map"]

HERE = os.path.dirname(os.path.abspath(__file__))
XAML = os.path.join(HERE, "SectionMap.xaml")
FORGET = u"(forget this choice)"

try:
    TEXT = unicode      # IronPython 2
except NameError:
    TEXT = str


def _text(v):
    try:
        return TEXT(v)
    except Exception:
        return u""


class SectionMapWindow(forms.WPFWindow):
    """rows: dicts with section, cat, text, count (or None), options
    (labels), default (label), valid (set of labels)."""

    def __init__(self, rows, mode, specials, file_name=u""):
        forms.WPFWindow.__init__(self, XAML)
        apply_seed43_palette(self, HERE)
        apply_seed43_dimensions(self, HERE)
        self.rows = rows
        self.mode = mode
        self.specials = list(specials)     # [(label, map value)]
        self.file_name = file_name
        self.shape_label = self.specials[0][0]
        self.combos = []
        self.forgets = []
        self.ok = False
        self.cleared = False
        self._setup()
        self._build()

    # -- layout -----------------------------------------------------------------

    def _setup(self):
        from System.Windows import Visibility
        editing = self.mode == "edit"
        self.subtitle_tb.Text = self.file_name if editing else u"Sections with no match"
        self.hint_tb.Text = (
            u"Every member section in this file and the Revit type it becomes: "
            u"your saved choice, or the automatic match by size and name. "
            u"Change any (type to search) and Save; only choices that differ "
            u"from the automatic match are saved, for this and later "
            u"imports." if editing else
            u"These sections have no matching Revit family type. Pick a type "
            u"for each (type to search), or import them as shapes. Library "
            u"types are loaded only if you pick them.")
        self.ok_btn.Content = u"Save" if editing else u"Import"
        self.remember_cb.Visibility = Visibility.Collapsed if editing \
            else Visibility.Visible
        self.clear_btn.Visibility = Visibility.Visible if editing \
            else Visibility.Collapsed
        if editing:
            self.clear_btn.Content = u"Reset to automatic"
            self.clear_btn.ToolTip = u"Give every section here its automatic match"
            self.all_shape_btn.Visibility = Visibility.Collapsed
            if not self.rows:
                self.empty_tb.Text = u"This file has no member sections to map."
                self.empty_tb.Visibility = Visibility.Visible
        self.ok_btn.Click += self._on_ok
        self.cancel_btn.Click += lambda s, a: self.Close()
        self.win_close_btn.Click += lambda s, a: self.Close()
        self.all_shape_btn.Click += self._on_all_shapes
        self.clear_btn.Click += self._on_clear

    def _columns(self, grid):
        from System.Windows import GridLength, GridUnitType
        from System.Windows.Controls import ColumnDefinition
        widths = [(230, False), (250, False), (70, False), (1, True)]
        if self.mode == "edit":
            widths = [(200, False), (210, False), (60, False), (1, True),
                      (100, False)]
        for w, star in widths:
            c = ColumnDefinition()
            c.Width = GridLength(w, GridUnitType.Star if star else GridUnitType.Pixel)
            grid.ColumnDefinitions.Add(c)

    def _text_cell(self, grid, text, row, col, style="Cell"):
        from System.Windows.Controls import TextBlock, Grid
        tb = TextBlock()
        tb.Text = text
        tb.Style = self.FindResource(style)
        tb.ToolTip = text
        Grid.SetRow(tb, row)
        Grid.SetColumn(tb, col)
        grid.Children.Add(tb)

    def _build(self):
        from System.Windows import Thickness, GridLength
        from System.Windows.Controls import (RowDefinition, ComboBox, CheckBox,
                                             Grid)

        def auto_row(grid):
            # rows as tall as their content; the default (*) shares out the
            # whole height, so a short list spread over the window
            rd = RowDefinition()
            rd.Height = GridLength.Auto
            grid.RowDefinitions.Add(rd)
        self._columns(self.head_grid)
        self._columns(self.rows_grid)
        heads = [u"SECTION", u"SHAPE AND SIZE", u"COUNT", u"REVIT TYPE"]
        if self.mode == "edit":
            heads = [u"SECTION", u"SHAPE AND SIZE", u"COUNT", u"REVIT TYPE",
                     u"FROM"]
        auto_row(self.head_grid)
        for c, h in enumerate(heads):
            self._text_cell(self.head_grid, h, 0, c, "ColHead")
        for i, row in enumerate(self.rows):
            auto_row(self.rows_grid)
            label = row["section"] or u"(no name)"
            if row["cat"] == "column":
                label += u"  (column)"
            self._text_cell(self.rows_grid, label, i, 0)
            self._text_cell(self.rows_grid, row.get("text") or u"", i, 1, "CellMuted")
            self._text_cell(self.rows_grid, _text(row.get("count") or u""), i, 2,
                            "CellMuted")
            type_col = 3
            if self.mode == "edit":
                self._text_cell(self.rows_grid, row.get("how") or u"", i, 4,
                                "CellMuted")
            cb = ComboBox()
            cb.Style = self.FindResource("TypeCombo")
            for label, value in self.specials:
                cb.Items.Add(label)
            for opt in row["options"]:
                cb.Items.Add(opt)
            cb.Text = row["default"]
            cb.ToolTip = row["default"]
            Grid.SetRow(cb, i)
            Grid.SetColumn(cb, type_col)
            self.rows_grid.Children.Add(cb)
            self.combos.append(cb)

    # -- buttons ----------------------------------------------------------------

    def _on_all_shapes(self, sender, args):
        for cb in self.combos:
            cb.Text = self.shape_label

    def _on_clear(self, sender, args):
        """Every row back to its automatic match (saved on Save)."""
        for row, cb in zip(self.rows, self.combos):
            cb.Text = row.get("auto") or self.shape_label
            row["reset"] = True

    def _on_ok(self, sender, args):
        from Snippets._dialogs import message
        bad = []
        for i, (row, cb) in enumerate(zip(self.rows, self.combos)):
            if self.forgets and self.forgets[i].IsChecked:
                continue
            text = _text(cb.Text).strip()
            if text not in dict(self.specials) and text not in row["valid"]:
                bad.append(row["section"] or u"(no name)")
        if bad:
            message(u"Not a known family type:\n\n" + u"\n".join(bad[:20]) +
                    u"\n\nPick one from its list, or import it as a shape "
                    u"or plate.",
                    title=u"Section Map")
            return
        self.ok = True
        self.Close()

    def choices(self):
        """{(cat, section): type label, a special's map value, or None to
        forget}."""
        special = dict(self.specials)
        out = {}
        for i, (row, cb) in enumerate(zip(self.rows, self.combos)):
            key = (row["cat"], row["section"])
            if self.forgets and self.forgets[i].IsChecked:
                out[key] = None
                continue
            text = _text(cb.Text).strip()
            out[key] = special.get(text, text)
        return out


def ask_section_map(rows, specials):
    """The sections with no match. specials: [(label, map value)] offered
    before the family types (import as shape, as plate). Returns
    ({(cat, section): type label or map value}, remember) or None when
    cancelled."""
    win = SectionMapWindow(rows, "import", specials)
    win.ShowDialog()
    if not win.ok:
        return None
    return win.choices(), bool(win.remember_cb.IsChecked)


def edit_section_map(rows, specials, file_name=u""):
    """The member sections of one file (see _steel_import.file_section_rows)
    to change. Returns {(cat, section): type label or map value}, or None
    when cancelled; Reset to automatic gives every row its automatic match."""
    win = SectionMapWindow(rows, "edit", specials, file_name)
    win.ShowDialog()
    if not win.ok:
        return None
    return win.choices()
