# -*- coding: utf-8 -*-
# FolderPresetManager.py
"""Manage Folder Presets window — create / edit / delete list."""
import os.path as op
from pyrevit import forms
from pyrevit.framework import Windows

import FolderPresetEditor as fpe_win
from Snippets import _dialogs as dlg
from Snippets.seed43_theme import apply_seed43_palette


class FolderPresetManagerWindow(forms.WPFWindow):
    def __init__(self, xaml_file_name, project_info, username, revit_version,
                 load_presets, save_presets):
        forms.WPFWindow.__init__(self, xaml_file_name)
        apply_seed43_palette(self, op.dirname(xaml_file_name))
        self._project_info = project_info
        self._username = username
        self._revit_version = revit_version
        self._load_presets = load_presets   # callable -> dict
        self._save_presets = save_presets   # callable(dict) -> None
        self.changed = False
        self._rebuild_list()

    def _rebuild_list(self):
        panel = self.preset_list_panel
        panel.Children.Clear()
        presets = self._load_presets()
        if not presets:
            tb = Windows.Controls.TextBlock()
            tb.Text = 'No presets yet — click "Create Preset" to add one.'
            tb.Foreground = self.FindResource('BrushTextPrimary')
            tb.Opacity = 0.5
            tb.FontSize = 12
            tb.TextWrapping = Windows.TextWrapping.Wrap
            panel.Children.Add(tb)
            return
        for name in sorted(presets):
            card = Windows.Controls.Border()
            card.Style = self.FindResource('RowCard')
            card.Margin = Windows.Thickness(0, 0, 0, 6)

            row = Windows.Controls.DockPanel()

            del_btn = Windows.Controls.Button()
            del_btn.Content = u'\u2715'
            del_btn.Style = self.FindResource('DeleteIconBtnFlat')
            del_btn.Tag = name
            del_btn.ToolTip = 'Delete preset'
            del_btn.Margin = Windows.Thickness(0, 4, 4, 4)
            del_btn.Click += self._delete_clicked
            Windows.Controls.DockPanel.SetDock(del_btn, Windows.Controls.Dock.Right)

            name_btn = Windows.Controls.Button()
            name_btn.Content = name
            name_btn.Style = self.FindResource('RowBtnFlat')
            name_btn.Tag = name
            name_btn.ToolTip = 'Edit preset'
            name_btn.Click += self._edit_clicked

            row.Children.Add(del_btn)
            row.Children.Add(name_btn)
            card.Child = row
            panel.Children.Add(card)

    def create_clicked(self, sender, args):
        result = fpe_win.edit_preset(self._project_info, self._username,
                                     self._revit_version)
        if not result:
            return
        presets = self._load_presets()
        presets[result['name']] = {'root': result['root'],
                                   'template': result['template']}
        self._save_presets(presets)
        self.changed = True
        self._rebuild_list()

    def _edit_clicked(self, sender, args):
        name = sender.Tag
        presets = self._load_presets()
        existing = presets.get(name)
        if existing:
            existing = dict(existing)
            existing['name'] = name
        result = fpe_win.edit_preset(self._project_info, self._username,
                                     self._revit_version, existing=existing)
        if not result:
            return
        if result['name'] != name:
            presets.pop(name, None)
        presets[result['name']] = {'root': result['root'],
                                   'template': result['template']}
        self._save_presets(presets)
        self.changed = True
        self._rebuild_list()

    def _delete_clicked(self, sender, args):
        name = sender.Tag
        if not dlg.confirm('Delete preset "{}"?'.format(name), yes='Delete'):
            return
        presets = self._load_presets()
        presets.pop(name, None)
        self._save_presets(presets)
        self.changed = True
        self._rebuild_list()

    def win_close_clicked(self, sender, args):
        self.Close()


def show_manager(project_info, username, revit_version, load_presets, save_presets):
    """Show the manager modally. Returns True if anything changed."""
    xaml_path = op.join(op.dirname(__file__), 'FolderPresetManager.xaml')
    win = FolderPresetManagerWindow(xaml_path, project_info, username,
                                    revit_version, load_presets, save_presets)
    win.ShowDialog()
    return win.changed
