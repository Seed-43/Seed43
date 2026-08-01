# -*- coding: utf-8 -*-
# ManageColumns.py
"""Manage Columns window — pick which sheet/view parameters show up as
grid columns, alongside the built-in Revision/Size/Collection columns
(sheets-only, listed here too so they can be toggled off).

The parameter scan only runs when this dialog is opened (on demand),
never during normal grid reloads — sample_sheet/sample_view are looked
up by the caller right before showing this window.
"""
import os.path as op
from pyrevit import forms
from pyrevit.framework import Windows

from Snippets.seed43_theme import apply_seed43_palette


BUILTIN_SHEET_COLUMNS = [
    ('revision',   'Revision'),
    ('size',       'Size'),
    ('collection', 'Collection'),
]
_BUILTIN_KEYS = set(k for k, _ in BUILTIN_SHEET_COLUMNS)


class ManageColumnsWindow(forms.WPFWindow):
    def __init__(self, xaml_file_name, sample_sheet, sample_view,
                 current_sheet_cols, current_view_cols, builtin_visible):
        forms.WPFWindow.__init__(self, xaml_file_name)
        apply_seed43_palette(self, op.dirname(xaml_file_name))
        self.result = None

        self._sheet_boxes = {}   # key -> CheckBox
        self._view_boxes  = {}   # key -> CheckBox

        for key, label in BUILTIN_SHEET_COLUMNS:
            self._add_row(self.sheet_list_panel, label,
                          (builtin_visible or {}).get(key, True),
                          self._sheet_boxes, key, builtin=True)

        sheet_params = self._scan_params(sample_sheet)
        if not sheet_params:
            self._add_note(self.sheet_list_panel,
                           'No sheet found in the model to scan parameters from.')
        for name in sheet_params:
            self._add_row(self.sheet_list_panel, name,
                          name in set(current_sheet_cols or []),
                          self._sheet_boxes, name, builtin=False)

        view_params = self._scan_params(sample_view)
        if not view_params:
            self._add_note(self.view_list_panel,
                           'No view found in the model to scan parameters from.')
        for name in view_params:
            self._add_row(self.view_list_panel, name,
                          name in set(current_view_cols or []),
                          self._view_boxes, name, builtin=False)

    @staticmethod
    def _scan_params(element):
        """Parameter names visible on one representative sheet/view."""
        if element is None:
            return []
        try:
            params = list(element.GetOrderedParameters())
        except Exception:
            try:
                params = list(element.Parameters)
            except Exception:
                params = []
        names = set()
        for p in params:
            try:
                if p.Definition and p.Definition.Name:
                    names.add(p.Definition.Name)
            except Exception:
                pass
        return sorted(names)

    def _add_note(self, panel, text):
        tb = Windows.Controls.TextBlock()
        tb.Text = text
        tb.Foreground = self.FindResource('BrushTextPrimary')
        tb.Opacity = 0.5
        tb.FontSize = 12
        tb.TextWrapping = Windows.TextWrapping.Wrap
        tb.Margin = Windows.Thickness(0, 0, 0, 6)
        panel.Children.Add(tb)

    def _add_row(self, panel, label, checked, box_dict, key, builtin):
        cb = Windows.Controls.CheckBox()
        cb.Content = label
        cb.IsChecked = checked
        cb.Foreground = self.FindResource('BrushTextPrimary')
        cb.FontSize = 12
        cb.Margin = Windows.Thickness(0, 0, 0, 6)
        if builtin:
            cb.FontWeight = Windows.FontWeights.SemiBold
        panel.Children.Add(cb)
        box_dict[key] = cb

    def save_clicked(self, sender, args):
        builtin_visible = {key: bool(self._sheet_boxes[key].IsChecked)
                           for key, _ in BUILTIN_SHEET_COLUMNS
                           if key in self._sheet_boxes}
        sheet_columns = sorted(key for key, cb in self._sheet_boxes.items()
                               if key not in _BUILTIN_KEYS and cb.IsChecked)
        view_columns = sorted(key for key, cb in self._view_boxes.items()
                              if cb.IsChecked)
        self.result = {
            'builtin_visible': builtin_visible,
            'sheet_columns':   sheet_columns,
            'view_columns':    view_columns,
        }
        self.Close()

    def cancel_clicked(self, sender, args):
        self.Close()

    def win_close_clicked(self, sender, args):
        self.Close()


def show_manager(sample_sheet, sample_view, current_sheet_cols,
                  current_view_cols, builtin_visible):
    """Show the manager modally. Returns a result dict, or None if cancelled."""
    xaml_path = op.join(op.dirname(__file__), 'ManageColumns.xaml')
    win = ManageColumnsWindow(xaml_path, sample_sheet, sample_view,
                              current_sheet_cols, current_view_cols,
                              builtin_visible)
    win.ShowDialog()
    return win.result
