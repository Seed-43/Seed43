# -*- coding: utf-8 -*-
"""Dialogs.py - pyTransmit's local dialog names, now a thin wrapper around
the shared Snippets._dialogs lib (the "lib" and "transmit" dialog code used
to be two separate, independently-styled implementations - this file used
to duplicate all of _dialogs.py's styling logic itself. It's consolidated
now: _dialogs.py is the one real implementation, this just keeps
pyTransmit's existing call sites (Dialogs.save_discard(...), etc.) working
unchanged so nothing else in pyTransmit needed to be touched.
"""
from Snippets import _dialogs as sdlg


class Dialogs(object):
    """Themed WPF dialog helpers for pyTransmit - see Snippets._dialogs for
    the actual implementation and styling; every method here is a thin
    wrapper mapping pyTransmit's existing call signatures onto it."""

    @staticmethod
    def save_discard(panel_label):
        """Prompt to save or discard changes to a named panel.
        Returns True if the user chose Save, False if Discard."""
        return sdlg.confirm(
            u"Do you want to save your changes to {}?".format(panel_label),
            title="Save Changes", yes="Save", no="Discard")

    @staticmethod
    def file_save(title, filename, ext, initial_folder=None):
        """Themed file save dialog with folder browser and editable filename.
        Returns the chosen full path, or None if cancelled."""
        return sdlg.save_file_as(title, filename, ext, initial_folder)

    @staticmethod
    def open_file(title, message):
        """Themed 'file saved, open it?' dialog.
        Returns True if the user chose Open, False if No."""
        return sdlg.confirm(message, title=title, yes="Open", no="No")

    @staticmethod
    def settings_mismatch(diff_text=''):
        """Prompt the user about a settings mismatch vs the last issued revision.
        Returns: 'update' | 'session' | 'ignore'"""
        result = sdlg.choice(
            "This project was previously issued with different settings.",
            [('ignore', 'Ignore'), ('session', 'This Issue Only'),
             ('update', 'Update Settings')],
            title="Settings Mismatch", detail_text=diff_text)
        return result or 'ignore'

    @staticmethod
    def save_log(default_folder, filename):
        """Save-log dialog, purpose-built rather than the generic
        save_file_as: the filename is fixed (pyTransmit_log_<date>_<time>,
        computed by the caller) and shown read-only, not editable - only the
        folder is a real choice here, with a Browse button, and the caller
        is expected to remember it (see menu_log_click/_save_log_setting in
        pyTransmit.py) so it defaults to wherever was used last time instead
        of always starting at the Desktop.

        Returns the chosen folder path, or None if cancelled.
        """
        w = sdlg._window(width=420)
        root = sdlg._card(w)
        result = {'folder': None}

        def _display_box(text):
            """Simple Border+TextBlock styled like an input box, for the two
            read-only fields below - neither is ever typed into directly
            (folder is set via Browse, filename is fixed), so there's no
            need for a real templated TextBox here. Using plain primitives
            instead of _themed_textbox's custom ControlTemplate sidesteps
            whatever was causing that template to render with no visible
            text."""
            b = sdlg.Windows.Controls.Border()
            b.Background = sdlg._theme_brush(w, 'BrushInputBg', sdlg.INPUTBG)
            b.BorderBrush = sdlg._theme_brush(w, 'BrushBorderDefault', sdlg.GREEN)
            b.BorderThickness = sdlg.Windows.Thickness(1)
            b.CornerRadius = sdlg._theme_dim(w, 'CornerRadiusInput', sdlg.Windows.CornerRadius(6))
            b.Padding = sdlg.Windows.Thickness(10, 0, 10, 0)
            b.Height = 34.0
            t = sdlg.Windows.Controls.TextBlock()
            t.Text = text or ''
            t.Foreground = sdlg._theme_brush(w, 'BrushTextInput', sdlg.TEXT)
            t.FontSize = 12
            t.VerticalAlignment = sdlg.Windows.VerticalAlignment.Center
            b.Child = t
            return b

        root.Children.Add(sdlg._textblock(w, 'Save Log File', title=True))

        folder_label = sdlg._textblock(w, 'Folder')
        folder_label.FontSize = 11
        folder_label.Opacity = 0.7
        folder_label.Margin = sdlg.Windows.Thickness(0, 0, 0, 4)
        root.Children.Add(folder_label)

        folder_row = sdlg.Windows.Controls.Grid()
        c0 = sdlg.Windows.Controls.ColumnDefinition(); c0.Width = sdlg.Windows.GridLength(1, sdlg.Windows.GridUnitType.Star)
        c1 = sdlg.Windows.Controls.ColumnDefinition(); c1.Width = sdlg.Windows.GridLength.Auto
        folder_row.ColumnDefinitions.Add(c0)
        folder_row.ColumnDefinitions.Add(c1)

        folder_box = _display_box(default_folder)
        folder_box.Margin = sdlg.Windows.Thickness(0, 0, 6, 12)
        sdlg.Windows.Controls.Grid.SetColumn(folder_box, 0)

        browse_btn = sdlg._button(w, 'Browse')
        browse_btn.Margin = sdlg.Windows.Thickness(0, 0, 0, 12)
        sdlg.Windows.Controls.Grid.SetColumn(browse_btn, 1)

        folder_row.Children.Add(folder_box)
        folder_row.Children.Add(browse_btn)
        root.Children.Add(folder_row)

        current_folder = {'path': default_folder or ''}

        def on_browse(s, a):
            try:
                from System.Windows.Forms import FolderBrowserDialog, DialogResult
                fb = FolderBrowserDialog()
                fb.SelectedPath = current_folder['path'] or default_folder or ''
                if fb.ShowDialog() == DialogResult.OK:
                    current_folder['path'] = fb.SelectedPath
                    folder_box.Child.Text = fb.SelectedPath
            except Exception:
                pass
        browse_btn.Click += on_browse

        name_label = sdlg._textblock(w, 'File name')
        name_label.FontSize = 11
        name_label.Opacity = 0.7
        name_label.Margin = sdlg.Windows.Thickness(0, 0, 0, 4)
        root.Children.Add(name_label)

        # Read-only, fixed - a lower opacity signals at a glance that this
        # one can't be changed, the way it's greyed out in most save dialogs.
        name_box = _display_box(filename)
        name_box.Opacity = 0.75
        root.Children.Add(name_box)

        row = sdlg.Windows.Controls.StackPanel()
        row.Orientation = sdlg.Windows.Controls.Orientation.Horizontal
        row.HorizontalAlignment = sdlg.Windows.HorizontalAlignment.Right

        def on_save(s, a):
            result['folder'] = (current_folder['path'] or '').strip()
            w.Close()
        cancel_btn = sdlg._button(w, 'Cancel', primary=False)
        cancel_btn.Click += lambda s, a: w.Close()
        save_btn = sdlg._button(w, 'Save')
        save_btn.Click += on_save
        row.Children.Add(cancel_btn)
        row.Children.Add(save_btn)
        root.Children.Add(row)

        w.ShowDialog()
        return result['folder']

    @staticmethod
    def alert(title, message, ok_label="OK"):
        """Themed alert dialog with a title, message, and single OK button.
        Use instead of forms.alert to match the pyTransmit dark theme."""
        sdlg.message(message, title=title, ok_label=ok_label)
