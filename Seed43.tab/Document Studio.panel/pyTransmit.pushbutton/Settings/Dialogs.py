# -*- coding: utf-8 -*-
# Dialogs.py
import os
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

import System.Windows.Markup as Markup
from System.Windows import Visibility
from System.IO import Path

try:
    from Snippets import _dialogs as sdlg
except Exception:
    sdlg = None


# ── SHARED XAML FRAGMENTS ─────────────────────────────────────────────────────

_WINDOW_OPEN = (
    '<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"'
    ' xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"'
    ' Title="" SizeToContent="Height"'
    ' WindowStyle="None" ResizeMode="NoResize"'
    ' WindowStartupLocation="CenterScreen"'
    ' Background="Transparent" FontFamily="Segoe UI" AllowsTransparency="True">'
    '<Border Background="#2B3340" CornerRadius="10" Margin="12" Padding="24,20,24,20">'
    '<Border.Effect>'
    '<DropShadowEffect Color="Black" Opacity="0.5" ShadowDepth="4" BlurRadius="16"/>'
    '</Border.Effect>'
    '<StackPanel>'
    '<Border Background="#208A3C" Height="3" CornerRadius="2" Margin="0,0,0,16"/>'
)
_WINDOW_CLOSE = '</StackPanel></Border></Window>'

_BTN_SECONDARY = (
    '<Button.Template><ControlTemplate TargetType="Button">'
    '<Border x:Name="Bd" Background="#404553" CornerRadius="6" Padding="20,8">'
    '<ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>'
    '</Border>'
    '<ControlTemplate.Triggers>'
    '<Trigger Property="IsMouseOver" Value="True">'
    '<Setter TargetName="Bd" Property="Background" Value="#4E5566"/></Trigger>'
    '<Trigger Property="IsPressed" Value="True">'
    '<Setter TargetName="Bd" Property="Background" Value="#333B48"/></Trigger>'
    '</ControlTemplate.Triggers>'
    '</ControlTemplate></Button.Template>'
)
_BTN_PRIMARY = (
    '<Button.Template><ControlTemplate TargetType="Button">'
    '<Border x:Name="Bd" Background="#208A3C" CornerRadius="6" Padding="20,8">'
    '<ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>'
    '</Border>'
    '<ControlTemplate.Triggers>'
    '<Trigger Property="IsMouseOver" Value="True">'
    '<Setter TargetName="Bd" Property="Background" Value="#2B933F"/></Trigger>'
    '<Trigger Property="IsPressed" Value="True">'
    '<Setter TargetName="Bd" Property="Background" Value="#1A6E2E"/></Trigger>'
    '</ControlTemplate.Triggers>'
    '</ControlTemplate></Button.Template>'
)


# ── DIALOG CLASS ──────────────────────────────────────────────────────────────

class Dialogs(object):
    """
    Themed WPF dialog helpers for pyTransmit.
    All dialogs match the dark theme defined in Seed43Styles.
    """

    # ── Save / Discard ────────────────────────────────────────────────────────

    @staticmethod
    def save_discard(panel_label):
        """
        Prompt to save or discard changes to a named panel.
        Returns True if the user chose Save, False if Discard.
        """
        xaml = (
            _WINDOW_OPEN.replace('SizeToContent="Height"',
                                 'Width="340" SizeToContent="Height"') +
            '<TextBlock Text="Save Changes"'
            ' Foreground="#F4FAFF" FontSize="15" FontWeight="Bold" Margin="0,0,0,8"/>'
            '<TextBlock x:Name="msg_tb"'
            ' Foreground="#F4FAFF" FontSize="12" Opacity="0.85"'
            ' TextWrapping="Wrap" Margin="0,0,0,24"/>'
            '<StackPanel Orientation="Horizontal" HorizontalAlignment="Right">'
            '<Button x:Name="discard_btn" Content="Discard"'
            ' Foreground="#F4FAFF" FontSize="12" FontWeight="Bold"'
            ' BorderThickness="0" Padding="20,8" Margin="0,0,8,0" Cursor="Hand">' +
            _BTN_SECONDARY +
            '</Button>'
            '<Button x:Name="save_btn" Content="Save"'
            ' Foreground="#F4FAFF" FontSize="12" FontWeight="Bold"'
            ' BorderThickness="0" Padding="20,8" Cursor="Hand">' +
            _BTN_PRIMARY +
            '</Button>'
            '</StackPanel>' +
            _WINDOW_CLOSE
        )
        dlg = Markup.XamlReader.Parse(xaml)
        dlg.FindName("msg_tb").Text = (
            u"Do you want to save your changes to {}?".format(panel_label))
        result = [False]

        def on_save(s, e):    result[0] = True;  dlg.Close()
        def on_discard(s, e): result[0] = False; dlg.Close()

        dlg.FindName("save_btn").Click    += on_save
        dlg.FindName("discard_btn").Click += on_discard
        dlg.ShowDialog()
        return result[0]

    # ── File Save ─────────────────────────────────────────────────────────────

    @staticmethod
    def file_save(title, filename, ext, initial_folder=None):
        """
        Themed file save dialog with folder browser and editable filename.
        Returns the chosen full path, or None if cancelled.
        """
        xaml = (
            _WINDOW_OPEN.replace('SizeToContent="Height"',
                                 'Width="480" SizeToContent="Height"') +
            '<TextBlock x:Name="title_tb"'
            ' Foreground="#F4FAFF" FontSize="15" FontWeight="Bold" Margin="0,0,0,12"/>'
            '<TextBlock Text="Folder"'
            ' Foreground="#F4FAFF" FontSize="11" Opacity="0.7" Margin="0,0,0,4"/>'
            '<Grid Margin="0,0,0,12">'
            '<Grid.ColumnDefinitions>'
            '<ColumnDefinition Width="*"/>'
            '<ColumnDefinition Width="Auto"/>'
            '</Grid.ColumnDefinitions>'
            '<Border Grid.Column="0" Background="#F4FAFF" CornerRadius="6,0,0,6"'
            ' BorderBrush="#208A3C" BorderThickness="1,1,0,1">'
            '<TextBlock x:Name="folder_tb" Foreground="#2B3340" FontSize="12"'
            ' Padding="8,4" VerticalAlignment="Center" TextTrimming="CharacterEllipsis"/>'
            '</Border>'
            '<Button x:Name="browse_btn" Grid.Column="1" Content="Browse"'
            ' Foreground="#F4FAFF" FontSize="11" FontWeight="SemiBold"'
            ' BorderThickness="0" Padding="12,0" Height="28" Cursor="Hand">'
            '<Button.Template><ControlTemplate TargetType="Button">'
            '<Border x:Name="Bd" Background="#208A3C" CornerRadius="0,6,6,0"'
            ' Padding="12,0">'
            '<ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>'
            '</Border>'
            '<ControlTemplate.Triggers>'
            '<Trigger Property="IsMouseOver" Value="True">'
            '<Setter TargetName="Bd" Property="Background" Value="#2B933F"/></Trigger>'
            '<Trigger Property="IsPressed" Value="True">'
            '<Setter TargetName="Bd" Property="Background" Value="#1A6E2E"/></Trigger>'
            '</ControlTemplate.Triggers>'
            '</ControlTemplate></Button.Template>'
            '</Button>'
            '</Grid>'
            '<TextBlock Text="File name"'
            ' Foreground="#F4FAFF" FontSize="11" Opacity="0.7" Margin="0,0,0,4"/>'
            '<Border Background="#F4FAFF" CornerRadius="6"'
            ' BorderBrush="#208A3C" BorderThickness="1" Margin="0,0,0,20">'
            '<TextBox x:Name="filename_tb" Background="Transparent" Foreground="#2B3340"'
            ' FontSize="12" Padding="8,4" BorderThickness="0"/>'
            '</Border>'
            '<StackPanel Orientation="Horizontal" HorizontalAlignment="Right">'
            '<Button x:Name="cancel_btn" Content="Cancel"'
            ' Foreground="#F4FAFF" FontSize="12" FontWeight="Bold"'
            ' BorderThickness="0" Padding="20,8" Margin="0,0,8,0" Cursor="Hand">' +
            _BTN_SECONDARY +
            '</Button>'
            '<Button x:Name="save_btn" Content="Save"'
            ' Foreground="#F4FAFF" FontSize="12" FontWeight="Bold"'
            ' BorderThickness="0" Padding="20,8" Cursor="Hand">' +
            _BTN_PRIMARY +
            '</Button>'
            '</StackPanel>' +
            _WINDOW_CLOSE
        )
        dlg         = Markup.XamlReader.Parse(xaml)
        dlg.FindName("title_tb").Text    = title
        dlg.FindName("filename_tb").Text = Path.GetFileNameWithoutExtension(filename)

        desktop   = os.path.expanduser("~\\Desktop")
        folder_tb   = dlg.FindName("folder_tb")
        filename_tb = dlg.FindName("filename_tb")
        folder_tb.Text = (
            initial_folder if (initial_folder and os.path.isdir(initial_folder))
            else desktop
        )
        result = [None]

        def on_browse(s, e):
            try:
                from System.Windows.Forms import FolderBrowserDialog, DialogResult
                fb = FolderBrowserDialog()
                fb.SelectedPath = folder_tb.Text or desktop
                if fb.ShowDialog() == DialogResult.OK:
                    folder_tb.Text = fb.SelectedPath
            except Exception:
                pass

        def on_save(s, e):
            folder = folder_tb.Text.strip()
            name   = filename_tb.Text.strip()
            if not name.lower().endswith('.' + ext):
                name = name + '.' + ext
            result[0] = os.path.join(folder, name)
            dlg.Close()

        def on_cancel(s, e):
            result[0] = None
            dlg.Close()

        dlg.FindName("browse_btn").Click += on_browse
        dlg.FindName("save_btn").Click   += on_save
        dlg.FindName("cancel_btn").Click += on_cancel
        dlg.ShowDialog()
        return result[0]

    # ── Open File ─────────────────────────────────────────────────────────────

    @staticmethod
    def open_file(title, message):
        """
        Themed 'file saved, open it?' dialog.
        Returns True if the user chose Open, False if No.
        """
        xaml = (
            _WINDOW_OPEN.replace('SizeToContent="Height"',
                                 'Width="360" SizeToContent="Height"') +
            '<TextBlock x:Name="title_tb"'
            ' Foreground="#F4FAFF" FontSize="15" FontWeight="Bold" Margin="0,0,0,8"/>'
            '<TextBlock x:Name="msg_tb"'
            ' Foreground="#F4FAFF" FontSize="12" Opacity="0.85"'
            ' TextWrapping="Wrap" Margin="0,0,0,24"/>'
            '<StackPanel Orientation="Horizontal" HorizontalAlignment="Right">'
            '<Button x:Name="no_btn" Content="No"'
            ' Foreground="#F4FAFF" FontSize="12" FontWeight="Bold"'
            ' BorderThickness="0" Padding="20,8" Margin="0,0,8,0" Cursor="Hand">' +
            _BTN_SECONDARY +
            '</Button>'
            '<Button x:Name="open_btn" Content="Open"'
            ' Foreground="#F4FAFF" FontSize="12" FontWeight="Bold"'
            ' BorderThickness="0" Padding="20,8" Cursor="Hand">' +
            _BTN_PRIMARY +
            '</Button>'
            '</StackPanel>' +
            _WINDOW_CLOSE
        )
        dlg = Markup.XamlReader.Parse(xaml)
        dlg.FindName("title_tb").Text = title
        dlg.FindName("msg_tb").Text   = message
        result = [False]

        def on_open(s, e): result[0] = True;  dlg.Close()
        def on_no(s, e):   result[0] = False; dlg.Close()

        dlg.FindName("open_btn").Click += on_open
        dlg.FindName("no_btn").Click   += on_no
        dlg.ShowDialog()
        return result[0]

    # ── Settings Mismatch ─────────────────────────────────────────────────────

    @staticmethod
    def settings_mismatch(diff_text=''):
        """
        Prompt the user about a settings mismatch vs the last issued revision.
        Returns: 'update' | 'session' | 'ignore'
        """
        xaml = (
            _WINDOW_OPEN.replace('SizeToContent="Height"',
                                 'Width="420" SizeToContent="Height"') +
            '<TextBlock Text="Settings Mismatch"'
            ' Foreground="#F4FAFF" FontSize="15" FontWeight="Bold" Margin="0,0,0,8"/>'
            '<TextBlock'
            ' Text="This project was previously issued with different settings."'
            ' Foreground="#F4FAFF" FontSize="12" Opacity="0.85"'
            ' TextWrapping="Wrap" Margin="0,0,0,24"/>'
            '<StackPanel Orientation="Horizontal" HorizontalAlignment="Right">'
            '<Button x:Name="ignore_btn" Content="Ignore"'
            ' Foreground="#F4FAFF" FontSize="12" FontWeight="Bold"'
            ' BorderThickness="0" Padding="20,8" Margin="0,0,8,0" Cursor="Hand">' +
            _BTN_SECONDARY +
            '</Button>'
            '<Button x:Name="session_btn" Content="This Issue Only"'
            ' Foreground="#F4FAFF" FontSize="12" FontWeight="Bold"'
            ' BorderThickness="0" Padding="20,8" Margin="0,0,8,0" Cursor="Hand">' +
            _BTN_SECONDARY +
            '</Button>'
            '<Button x:Name="update_btn" Content="Update Settings"'
            ' Foreground="#F4FAFF" FontSize="12" FontWeight="Bold"'
            ' BorderThickness="0" Padding="20,8" Cursor="Hand">' +
            _BTN_PRIMARY +
            '</Button>'
            '</StackPanel>'
            '<Border x:Name="diff_border" Background="#3B4553" CornerRadius="6"'
            ' Padding="12" Margin="0,16,0,0" Visibility="Collapsed">'
            '<TextBlock x:Name="diff_tb" Foreground="#F4FAFF" FontSize="11"'
            ' Opacity="0.7" TextWrapping="Wrap" FontFamily="Consolas"/>'
            '</Border>' +
            _WINDOW_CLOSE
        )
        dlg = Markup.XamlReader.Parse(xaml)
        if diff_text:
            dlg.FindName("diff_tb").Text = diff_text
            dlg.FindName("diff_border").Visibility = Visibility.Visible

        result = ['ignore']
        def on_update(s, e):  result[0] = 'update';  dlg.Close()
        def on_session(s, e): result[0] = 'session'; dlg.Close()
        def on_ignore(s, e):  result[0] = 'ignore';  dlg.Close()

        dlg.FindName("update_btn").Click  += on_update
        dlg.FindName("session_btn").Click += on_session
        dlg.FindName("ignore_btn").Click  += on_ignore
        dlg.ShowDialog()
        return result[0]

    # ── Alert ─────────────────────────────────────────────────────────────────

    @staticmethod
    def alert(title, message, ok_label="OK"):
        """
        Themed alert dialog with a title, message, and single OK button.
        Use instead of forms.alert to match the pyTransmit dark theme.

        Delegates to the shared Snippets._dialogs module when available and
        no custom ok_label is needed, falls back to the local XAML card
        below otherwise (custom button label, or shared lib not on path).
        """
        if sdlg and ok_label == "OK":
            sdlg.message(message, title=title)
            return
        xaml = (
            _WINDOW_OPEN.replace('SizeToContent="Height"',
                                 'Width="420" SizeToContent="Height"') +
            '<TextBlock x:Name="title_tb"'
            ' Foreground="#F4FAFF" FontSize="14" FontWeight="Bold" Margin="0,0,0,10"/>'
            '<TextBlock x:Name="msg_tb"'
            ' Foreground="#F4FAFF" FontSize="12" Opacity="0.85"'
            ' TextWrapping="Wrap" Margin="0,0,0,24"/>'
            '<StackPanel Orientation="Horizontal" HorizontalAlignment="Right">'
            '<Button x:Name="ok_btn"'
            ' Foreground="#F4FAFF" FontSize="12" FontWeight="Bold"'
            ' BorderThickness="0" Padding="20,8" Cursor="Hand">' +
            _BTN_PRIMARY +
            '</Button>'
            '</StackPanel>' +
            _WINDOW_CLOSE
        )
        dlg = Markup.XamlReader.Parse(xaml)
        dlg.FindName("title_tb").Text  = title
        dlg.FindName("msg_tb").Text    = message
        dlg.FindName("ok_btn").Content = ok_label

        def on_ok(s, e): dlg.Close()
        dlg.FindName("ok_btn").Click += on_ok
        dlg.ShowDialog()
