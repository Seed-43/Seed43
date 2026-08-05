# -*- coding: utf-8 -*-
# FolderPresetEditor.py
"""Export-folder preset editor window (real XAML, see FolderPresetEditor.xaml).

Path-resolution (tokens, bucket matching, wildcard) lives in
folder_preset_resolve.py so it stays independent of the UI.
"""
import os.path as op
from pyrevit import forms
from pyrevit.framework import Windows

import folder_preset_resolve as resolve
from Snippets.seed43_theme import apply_seed43_palette


class FolderPresetEditorWindow(forms.WPFWindow):
    def __init__(self, xaml_file_name, project_info, username, revit_version,
                 existing=None):
        forms.WPFWindow.__init__(self, xaml_file_name)
        apply_seed43_palette(self, op.dirname(xaml_file_name))
        self._project_info = project_info
        self._username = username
        self._revit_version = revit_version
        self._saved = False
        self.result = None

        if existing:
            self.name_tb.Text = existing.get('name', '')
            self.root_tb.Text = existing.get('root', '')
            self.template_tb.Text = existing.get('template',
                                                  resolve.DEFAULT_TEMPLATE)
        else:
            self.template_tb.Text = resolve.DEFAULT_TEMPLATE

        self.wildcard_hint_tb.Text = resolve.WILDCARD_HELP
        self._build_token_chips()
        self._refresh_preview()

    def _build_token_chips(self):
        for token, color in resolve.TOKENS:
            chip = Windows.Controls.Button()
            chip.Content = token
            chip.Tag = token
            chip.Style = self.FindResource('Chip')
            chip.Background = Windows.Media.SolidColorBrush(
                Windows.Media.ColorConverter.ConvertFromString(color))
            chip.Click += self._chip_clicked
            chip.PreviewMouseLeftButtonDown += self._chip_drag_start
            self.token_panel.Children.Add(chip)

    def _chip_clicked(self, sender, args):
        self._insert_token(sender.Tag)

    def _chip_drag_start(self, sender, args):
        data = Windows.DataObject('token', sender.Tag)
        Windows.DragDrop.DoDragDrop(sender, data, Windows.DragDropEffects.Copy)

    def _insert_token(self, token):
        pos = self.template_tb.CaretIndex
        text = self.template_tb.Text or ''
        self.template_tb.Text = text[:pos] + token + text[pos:]
        self.template_tb.CaretIndex = pos + len(token)
        self.template_tb.Focus()

    def template_drag_over(self, sender, args):
        point = args.GetPosition(self.template_tb)
        idx = self.template_tb.GetCharacterIndexFromPoint(point, True)
        if idx < 0:
            idx = len(self.template_tb.Text or '')
        self.template_tb.CaretIndex = idx
        args.Effects = Windows.DragDropEffects.Copy
        args.Handled = True

    def template_drop(self, sender, args):
        token = args.Data.GetData('token')
        if token:
            self._insert_token(token)
        args.Handled = True

    def browse_clicked(self, sender, args):
        from pyrevit.framework import Forms
        fbd = Forms.FolderBrowserDialog()
        if fbd.ShowDialog() == Forms.DialogResult.OK:
            self.root_tb.Text = fbd.SelectedPath
            self._refresh_preview()

    def field_changed(self, sender, args):
        self._refresh_preview()

    def _refresh_preview(self):
        try:
            self.preview_tb.Text = resolve.resolve_path(
                self.template_tb.Text or '', self.root_tb.Text or '',
                self._project_info, self._username, self._revit_version)
        except Exception:
            self.preview_tb.Text = ''

    def save_clicked(self, sender, args):
        name = (self.name_tb.Text or '').strip()
        if not name:
            self.name_tb.Focus()
            return
        self.result = {
            'name': name,
            'root': (self.root_tb.Text or '').strip(),
            'template': (self.template_tb.Text or '').strip(),
        }
        self._saved = True
        self.Close()

    def cancel_clicked(self, sender, args):
        self.Close()

    def win_close_clicked(self, sender, args):
        self.Close()


def edit_preset(project_info, username, revit_version, existing=None):
    """Show the editor modally. Returns the preset dict, or None if cancelled."""
    xaml_path = op.join(op.dirname(__file__), 'FolderPresetEditor.xaml')
    win = FolderPresetEditorWindow(xaml_path, project_info, username,
                                   revit_version, existing=existing)
    win.ShowDialog()
    return win.result
