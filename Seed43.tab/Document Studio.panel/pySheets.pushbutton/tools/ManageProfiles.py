# -*- coding: utf-8 -*-
# ManageProfiles.py
"""Profiles window — view / delete saved pySheets profiles.
Creating and saving profiles still happens from the header (+ / save);
this window is delete-only, mirroring FolderPresetManager's row style."""
import os.path as op
from pyrevit import forms
from pyrevit.framework import Windows

from Snippets import _dialogs as dlg
from Snippets.seed43_theme import apply_seed43_palette


class ManageProfilesWindow(forms.WPFWindow):
    def __init__(self, xaml_file_name, list_profiles, delete_profile):
        forms.WPFWindow.__init__(self, xaml_file_name)
        apply_seed43_palette(self, op.dirname(xaml_file_name))
        self._list_profiles  = list_profiles    # callable -> list of names
        self._delete_profile = delete_profile   # callable(name) -> None
        self.changed = False
        self._rebuild_list()

    def _rebuild_list(self):
        panel = self.profile_list_panel
        panel.Children.Clear()
        names = self._list_profiles()
        if not names:
            tb = Windows.Controls.TextBlock()
            tb.Text = 'No profiles yet — use the + button in the header to save one.'
            tb.Foreground = self.FindResource('BrushTextPrimary')
            tb.Opacity = 0.5
            tb.FontSize = 12
            tb.TextWrapping = Windows.TextWrapping.Wrap
            panel.Children.Add(tb)
            return
        for name in sorted(names):
            card = Windows.Controls.Border()
            card.Style = self.FindResource('RowCard')
            card.Margin = Windows.Thickness(0, 0, 0, 6)

            row = Windows.Controls.DockPanel()

            del_btn = Windows.Controls.Button()
            del_btn.Content = u'\u2715'
            del_btn.Style = self.FindResource('DeleteIconBtnFlat')
            del_btn.Tag = name
            del_btn.ToolTip = 'Delete profile'
            del_btn.Margin = Windows.Thickness(0, 4, 4, 4)
            del_btn.Click += self._delete_clicked
            Windows.Controls.DockPanel.SetDock(del_btn, Windows.Controls.Dock.Right)

            name_tb = Windows.Controls.TextBlock()
            name_tb.Text = name
            name_tb.Foreground = self.FindResource('BrushTextPrimary')
            name_tb.FontSize = 12
            name_tb.VerticalAlignment = Windows.VerticalAlignment.Center
            name_tb.Margin = Windows.Thickness(10, 8, 8, 8)

            row.Children.Add(del_btn)
            row.Children.Add(name_tb)
            card.Child = row
            panel.Children.Add(card)

    def _delete_clicked(self, sender, args):
        name = sender.Tag
        if not dlg.confirm('Delete profile "{}"?'.format(name), yes='Delete'):
            return
        self._delete_profile(name)
        self.changed = True
        self._rebuild_list()

    def win_close_clicked(self, sender, args):
        self.Close()


def show_manager(list_profiles, delete_profile):
    """Show the manager modally. Returns True if anything changed."""
    xaml_path = op.join(op.dirname(__file__), 'ManageProfiles.xaml')
    win = ManageProfilesWindow(xaml_path, list_profiles, delete_profile)
    win.ShowDialog()
    return win.changed
