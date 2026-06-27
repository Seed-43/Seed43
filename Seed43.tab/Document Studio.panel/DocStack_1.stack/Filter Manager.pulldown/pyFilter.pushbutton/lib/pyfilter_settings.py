# -*- coding: utf-8 -*-
# settings_dialog.py
# Seed43 Filter Manager - sync settings window. XAML embedded inline so there
# is no separate file to locate at runtime.
# pylint: disable=import-error,invalid-name,broad-except

import os

from pyrevit import forms

import wpf  # noqa: F401  IronPython only
from System.Windows import Window
from System.Windows.Controls import CheckBox, TextBlock
from System.Windows.Markup import XamlReader

from pyfilter_sync import (
    get_server_path, set_server_path,
    get_synced_templates, set_synced_templates,
    list_server_templates, server_ready, sync_all,
)

# ── XAML ──────────────────────────────────────────────────────────────────────

XAML = u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Filter Manager - Sync Settings"
        Height="560" Width="640"
        WindowStartupLocation="CenterOwner"
        Background="#3B4553" Foreground="#F4FAFF"
        TextElement.FontFamily="Segoe UI" FontSize="12">
    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="#F4FAFF"/>
            <Setter Property="FontSize"   Value="12"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Background"  Value="#404553"/>
            <Setter Property="Foreground"  Value="#F4FAFF"/>
            <Setter Property="BorderBrush" Value="#5A6273"/>
            <Setter Property="Padding"     Value="10,4"/>
            <Setter Property="MinHeight"   Value="28"/>
            <Setter Property="FontSize"    Value="12"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background"   Value="#2A303B"/>
            <Setter Property="Foreground"   Value="#F4FAFF"/>
            <Setter Property="BorderBrush"  Value="#5A6273"/>
            <Setter Property="Padding"      Value="6,4"/>
            <Setter Property="FontSize"     Value="12"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="#F4FAFF"/>
            <Setter Property="FontSize"   Value="12"/>
            <Setter Property="Margin"     Value="2,3"/>
        </Style>
    </Window.Resources>

    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Text="Sync Settings"
                   FontSize="16" FontWeight="SemiBold" Foreground="#208A3C"
                   Margin="0,0,0,12"/>

        <TextBlock Grid.Row="1" Text="Server path"
                   FontWeight="SemiBold" Margin="0,0,0,4"/>

        <Grid Grid.Row="2" Margin="0,0,0,4">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBox x:Name="TxtServerPath" Grid.Column="0"/>
            <Button  x:Name="BtnBrowse" Grid.Column="1" Content="Browse..."
                     Margin="6,0,0,0"/>
        </Grid>

        <TextBlock x:Name="TxtServerStatus" Grid.Row="3"
                   Text="" FontSize="11" Foreground="#9CA3AF"
                   Margin="0,0,0,12"/>

        <Grid Grid.Row="4" Margin="0,0,0,8">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <TextBlock Grid.Row="0" Text="Templates to sync"
                       FontWeight="SemiBold" Margin="0,0,0,4"/>
            <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,6">
                <Button x:Name="BtnSelectAll"  Content="Select all" Margin="0,0,6,0"/>
                <Button x:Name="BtnSelectNone" Content="Clear"      Margin="0,0,6,0"/>
                <TextBlock x:Name="TxtCount" Text="" VerticalAlignment="Center"
                           Foreground="#9CA3AF" FontSize="11" Margin="6,0,0,0"/>
            </StackPanel>
            <Border Grid.Row="2" Background="#2A303B" BorderBrush="#404553"
                    BorderThickness="1" Padding="8">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel x:Name="TemplateList"/>
                </ScrollViewer>
            </Border>
        </Grid>

        <StackPanel Grid.Row="5" Orientation="Horizontal" Margin="0,0,0,10">
            <Button x:Name="BtnSyncNow" Content="Sync now" Background="#208A3C"
                    Padding="14,4"/>
            <TextBlock x:Name="TxtSyncStatus" Text="" VerticalAlignment="Center"
                       Margin="10,0,0,0" Foreground="#9CA3AF" FontSize="11"/>
        </StackPanel>

        <StackPanel Grid.Row="6" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="BtnCancel" Content="Cancel" Margin="0,0,8,0" Padding="16,4"/>
            <Button x:Name="BtnSave"   Content="Save"   Background="#208A3C" Padding="20,4"/>
        </StackPanel>
    </Grid>
</Window>
"""

# ── HELPERS ───────────────────────────────────────────────────────────────────

def _list_local_templates(folder):
    if not folder or not os.path.isdir(folder):
        return []
    out = []
    for f in os.listdir(folder):
        if f.lower().endswith(".json"):
            out.append(os.path.splitext(f)[0])
    return sorted(out)


# ── DIALOG ────────────────────────────────────────────────────────────────────

class _SettingsDialog(object):
    def __init__(self, templates_folder, on_changed=None):
        self.templates_folder = templates_folder
        self.on_changed       = on_changed
        self._checks          = {}

        # Parse the embedded XAML directly from string.
        self.win = XamlReader.Parse(XAML)

        # Resolve named controls.
        for name in ("TxtServerPath", "BtnBrowse", "TxtServerStatus",
                     "TemplateList", "BtnSelectAll", "BtnSelectNone",
                     "TxtCount", "BtnSyncNow", "TxtSyncStatus",
                     "BtnSave", "BtnCancel"):
            setattr(self, name, self.win.FindName(name))

        self.TxtServerPath.Text = get_server_path()
        self.TxtServerPath.TextChanged += self._on_path_changed

        self.BtnBrowse.Click     += self._on_browse
        self.BtnSelectAll.Click  += lambda s, e: self._set_all(True)
        self.BtnSelectNone.Click += lambda s, e: self._set_all(False)
        self.BtnSyncNow.Click    += self._on_sync_now
        self.BtnSave.Click       += self._on_save
        self.BtnCancel.Click     += lambda s, e: self.win.Close()

        self._populate_template_list()
        self._refresh_server_status()
        self._refresh_count()

    # ── template list ─────────────────────────────────────────────────────────

    def _populate_template_list(self):
        self.TemplateList.Children.Clear()
        self._checks = {}

        local_names  = set(_list_local_templates(self.templates_folder))
        server_names = set(list_server_templates(self.TxtServerPath.Text))
        synced       = set(get_synced_templates())

        all_names = sorted(local_names | server_names | synced)
        if not all_names:
            tb = TextBlock()
            tb.Text = "No templates found locally or on server."
            from System.Windows.Media import SolidColorBrush, Color
            tb.Foreground = SolidColorBrush(Color.FromRgb(156, 163, 175))
            self.TemplateList.Children.Add(tb)
            return

        for name in all_names:
            cb = CheckBox()
            tag = []
            if name in local_names:  tag.append("local")
            if name in server_names: tag.append("server")
            label = name if not tag else "{}   ({})".format(name, ", ".join(tag))
            cb.Content   = label
            cb.Tag       = name
            cb.IsChecked = name in synced
            cb.Checked   += lambda s, e: self._refresh_count()
            cb.Unchecked += lambda s, e: self._refresh_count()
            self.TemplateList.Children.Add(cb)
            self._checks[name] = cb

    def _refresh_count(self):
        n = sum(1 for cb in self._checks.values() if cb.IsChecked)
        self.TxtCount.Text = "{} selected of {}".format(n, len(self._checks))

    def _set_all(self, value):
        for cb in self._checks.values():
            cb.IsChecked = bool(value)

    def _selected_names(self):
        return [n for n, cb in self._checks.items() if cb.IsChecked]

    # ── server path ───────────────────────────────────────────────────────────

    def _on_browse(self, sender, e):
        folder = forms.pick_folder(title="Pick sync folder")
        if folder:
            self.TxtServerPath.Text = folder

    def _on_path_changed(self, sender, e):
        self._refresh_server_status()
        self._populate_template_list()
        self._refresh_count()

    def _refresh_server_status(self):
        path = self.TxtServerPath.Text
        if not path:
            self.TxtServerStatus.Text = "No path set. Sync is disabled."
            return
        ok, msg = server_ready(path)
        if ok:
            n = len(list_server_templates(path))
            self.TxtServerStatus.Text = "Reachable. {} template file(s) on server.".format(n)
        else:
            self.TxtServerStatus.Text = "Warning: {}".format(msg)

    # ── actions ───────────────────────────────────────────────────────────────

    def _persist(self):
        set_server_path(self.TxtServerPath.Text.strip())
        set_synced_templates(self._selected_names())

    def _on_sync_now(self, sender, e):
        try:
            self._persist()
            messages = []
            result = sync_all(self.templates_folder,
                              logger=lambda m: messages.append(m))
            if not result.get("ok"):
                self.TxtSyncStatus.Text = "Skipped: {}".format(result.get("reason"))
                forms.alert("Sync did not run.\n\n{}".format(result.get("reason")),
                            title="Sync now")
                return
            counts = result.get("counts", {})
            self.TxtSyncStatus.Text = ", ".join(
                "{}:{}".format(k, v) for k, v in sorted(counts.items())) or "Nothing to sync."
            self._populate_template_list()
            self._refresh_count()
            if self.on_changed:
                self.on_changed()
        except Exception as ex:
            forms.alert("Sync failed: {}".format(ex), title="Sync now")

    def _on_save(self, sender, e):
        try:
            self._persist()
            if self.on_changed:
                self.on_changed()
            self.win.Close()
        except Exception as ex:
            forms.alert("Save failed: {}".format(ex), title="Settings")


def show(owner_window, templates_folder, on_changed=None):
    dlg = _SettingsDialog(templates_folder, on_changed=on_changed)
    if owner_window is not None:
        try:
            dlg.win.Owner = owner_window
        except Exception:
            pass
    dlg.win.ShowDialog()



# ── SHARED STYLES ─────────────────────────────────────────────────────────────

_COMMON_STYLES = u"""
    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="#F4FAFF"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Background"      Value="#208A3C"/>
            <Setter Property="Foreground"      Value="#F4FAFF"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FontFamily"      Value="Segoe UI"/>
            <Setter Property="FontSize"        Value="13"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="Bd" Background="{TemplateBinding Background}"
                                CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center"
                                              VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#2B933F"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#1A6E2E"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="GrayBtn" TargetType="Button">
            <Setter Property="Background"      Value="#3B4553"/>
            <Setter Property="Foreground"      Value="#F4FAFF"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FontFamily"      Value="Segoe UI"/>
            <Setter Property="FontSize"        Value="13"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="Bd" Background="{TemplateBinding Background}"
                                CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center"
                                              VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#4E5566"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background"         Value="#F4FAFF"/>
            <Setter Property="Foreground"         Value="#1E2530"/>
            <Setter Property="BorderBrush"        Value="#208A3C"/>
            <Setter Property="BorderThickness"    Value="2"/>
            <Setter Property="Padding"            Value="8,6"/>
            <Setter Property="FontFamily"         Value="Segoe UI"/>
            <Setter Property="FontSize"           Value="12"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6">
                            <ScrollViewer x:Name="PART_ContentHost" Margin="0"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="#F4FAFF"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontSize"   Value="13"/>
            <Setter Property="Margin"     Value="0,5"/>
        </Style>
        <Style TargetType="RadioButton">
            <Setter Property="Foreground" Value="#F4FAFF"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontSize"   Value="12"/>
        </Style>
        <!-- Mac-style green scrollbar (Seed43Styles spec) -->
        <Style x:Key="MacScrollBarThumb" TargetType="Thumb">
            <Setter Property="Background" Value="#208A3C"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Thumb">
                        <Border x:Name="thumb"
                                Background="{TemplateBinding Background}"
                                CornerRadius="3"
                                Margin="2,2,2,2"/>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="thumb" Property="Background" Value="#27AE60"/>
                            </Trigger>
                            <Trigger Property="IsDragging" Value="True">
                                <Setter TargetName="thumb" Property="Background" Value="#2ECC71"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="DarkSB" TargetType="ScrollBar">
            <Setter Property="Orientation" Value="Vertical"/>
            <Setter Property="Width"       Value="8"/>
            <Setter Property="MinWidth"    Value="8"/>
            <Setter Property="Background"  Value="Transparent"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ScrollBar">
                        <Grid x:Name="GridRoot" Width="8">
                            <Track x:Name="PART_Track"
                                   Orientation="Vertical"
                                   IsDirectionReversed="True">
                                <Track.Thumb>
                                    <Thumb Style="{StaticResource MacScrollBarThumb}"/>
                                </Track.Thumb>
                            </Track>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="GridRoot" Property="Width" Value="10"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="DarkSV" TargetType="ScrollViewer">
            <Setter Property="VerticalScrollBarVisibility"   Value="Auto"/>
            <Setter Property="HorizontalScrollBarVisibility" Value="Disabled"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ScrollViewer">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <ScrollContentPresenter Grid.Column="0"
                                                    Margin="{TemplateBinding Padding}"
                                                    CanContentScroll="{TemplateBinding CanContentScroll}"/>
                            <ScrollBar x:Name="PART_VerticalScrollBar"
                                       Grid.Column="1"
                                       Orientation="Vertical"
                                       Style="{StaticResource DarkSB}"
                                       Value="{TemplateBinding VerticalOffset}"
                                       Maximum="{TemplateBinding ScrollableHeight}"
                                       ViewportSize="{TemplateBinding ViewportHeight}"
                                       Visibility="{TemplateBinding ComputedVerticalScrollBarVisibility}"/>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>"""


# ── EXPORT DIALOG ─────────────────────────────────────────────────────────────

EXPORT_XAML = u"""<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="pyFilter | Export" Height="620" Width="680"
    WindowStartupLocation="CenterOwner"
    Background="#1E2530" Foreground="#F4FAFF">
""" + _COMMON_STYLES + u"""
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="60"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Background="#161C24">
            <Grid Margin="16,0">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="py" FontWeight="Bold" FontSize="22"
                               Foreground="#208A3C"/>
                    <TextBlock Text="Filter" FontSize="22" FontWeight="SemiBold"
                               Margin="3,0,0,0"/>
                    <TextBlock Text=" | " FontSize="18" Foreground="#4A5568" Margin="6,0"/>
                    <TextBlock Text="Export" FontSize="18" FontWeight="SemiBold"/>
                </StackPanel>
                <Button x:Name="BtnExportNow" Grid.Column="1"
                        Content="Export Now" Padding="20,8" Margin="0,0,10,0"/>
                <Button x:Name="BtnClose" Grid.Column="2"
                        Content="X" Style="{StaticResource GrayBtn}"
                        Width="36" Height="36" FontWeight="Bold"/>
            </Grid>
        </Border>

        <!-- Body -->
        <ScrollViewer Grid.Row="1" Style="{StaticResource DarkSV}" Padding="16,12">
            <StackPanel>
                <!-- Export Settings card -->
                <Border Background="#2B3340" CornerRadius="8" Padding="16" Margin="0,0,0,12">
                    <StackPanel>
                        <TextBlock Text="Export Settings" Foreground="#208A3C"
                                   FontSize="15" FontWeight="SemiBold" Margin="0,0,0,6"/>
                        <TextBlock Text="Each selected template is saved as its own .json file in the chosen folder."
                                   Foreground="#9CA3AF" FontSize="12" TextWrapping="Wrap"
                                   Margin="0,0,0,14"/>
                        <CheckBox x:Name="ChkAutoExport" Content="Auto update on save"
                                  Margin="0,0,0,4"/>
                        <TextBlock Text="When enabled, saving a template automatically copies it to the export folder."
                                   Foreground="#9CA3AF" FontSize="11" TextWrapping="Wrap"
                                   Margin="20,0,0,14"/>
                        <TextBlock Text="Export Folder" FontWeight="SemiBold" Margin="0,0,0,6"/>
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBox x:Name="TxtDest" Grid.Column="0" Height="36"/>
                            <Button x:Name="BtnBrowse" Grid.Column="1"
                                    Content="Browse" Padding="16,8" Margin="8,0,0,0"/>
                        </Grid>
                    </StackPanel>
                </Border>

                <!-- Template list card -->
                <Border Background="#2B3340" CornerRadius="8" Padding="16">
                    <StackPanel>
                        <TextBlock Text="Select templates to export" Foreground="#208A3C"
                                   FontSize="15" FontWeight="SemiBold" Margin="0,0,0,10"/>
                        <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                            <Button x:Name="BtnAll"  Content="All"  Padding="14,6"
                                    Margin="0,0,8,0"/>
                            <Button x:Name="BtnNone" Content="None" Padding="14,6"
                                    Style="{StaticResource GrayBtn}"/>
                        </StackPanel>
                        <Border Background="#1E2530" CornerRadius="6" Padding="8"
                                Height="200">
                            <ScrollViewer Style="{StaticResource DarkSV}">
                                <StackPanel x:Name="TemplateList"/>
                            </ScrollViewer>
                        </Border>
                    </StackPanel>
                </Border>
            </StackPanel>
        </ScrollViewer>

        <!-- Footer -->
        <Border Grid.Row="2" Background="#161C24" Padding="16,10">
            <TextBlock x:Name="TxtStatus" Text="Settings are saved automatically when you close this panel."
                       Foreground="#4A5568" FontSize="11"/>
        </Border>
    </Grid>
</Window>"""


def _list_local_templates_internal(folder):
    if not folder or not os.path.isdir(folder):
        return []
    out = []
    for f in os.listdir(folder):
        if f.lower().endswith(".json"):
            out.append(os.path.splitext(f)[0])
    return sorted(out)


class _ExportDialog(object):
    def __init__(self, templates_folder):
        self.templates_folder = templates_folder
        self._checks          = {}
        self.win              = XamlReader.Parse(EXPORT_XAML)
        for n in ("TxtDest", "BtnBrowse", "BtnAll", "BtnNone", "ChkAutoExport",
                  "TemplateList", "TxtStatus", "BtnClose", "BtnExportNow"):
            setattr(self, n, self.win.FindName(n))

        self.BtnBrowse.Click    += self._on_browse
        self.BtnAll.Click       += lambda s, e: self._set_all(True)
        self.BtnNone.Click      += lambda s, e: self._set_all(False)
        self.BtnExportNow.Click += self._on_export
        self.BtnClose.Click     += lambda s, e: self.win.Close()
        self._populate()

    def _populate(self):
        self.TemplateList.Children.Clear()
        self._checks = {}
        from System.Windows.Controls import CheckBox as WpfCheckBox
        from System.Windows import Thickness
        names = _list_local_templates_internal(self.templates_folder)
        if not names:
            from System.Windows.Controls import TextBlock as TB
            from System.Windows.Media import SolidColorBrush, Color
            tb = TB()
            tb.Text = "No local templates found."
            tb.Foreground = SolidColorBrush(Color.FromRgb(156, 163, 175))
            self.TemplateList.Children.Add(tb)
            return
        for name in names:
            cb = WpfCheckBox()
            cb.Content   = name
            cb.Tag       = name
            cb.IsChecked = True
            cb.Margin    = Thickness(0, 3, 0, 3)
            self.TemplateList.Children.Add(cb)
            self._checks[name] = cb

    def _set_all(self, value):
        for cb in self._checks.values():
            cb.IsChecked = bool(value)

    def _on_browse(self, sender, e):
        folder = forms.pick_folder(title="Pick export destination")
        if folder:
            self.TxtDest.Text = folder

    def _on_export(self, sender, e):
        try:
            dest = (self.TxtDest.Text or "").strip()
            if not dest:
                forms.alert("Pick a destination folder first.")
                return
            if not os.path.isdir(dest):
                try:
                    os.makedirs(dest)
                except Exception as ex:
                    forms.alert("Cannot create destination: {}".format(ex))
                    return
            if not os.access(dest, os.W_OK):
                forms.alert("Destination is not writable.")
                return
            selected = [n for n, cb in self._checks.items() if cb.IsChecked]
            if not selected:
                forms.alert("Tick at least one template to export.")
                return
            import shutil
            written = skipped = 0
            errors = []
            for name in selected:
                src = os.path.join(self.templates_folder, name + ".json")
                tgt = os.path.join(dest, name + ".json")
                if not os.path.isfile(src):
                    skipped += 1
                    continue
                try:
                    shutil.copy2(src, tgt)
                    written += 1
                except Exception as ex:
                    errors.append("{}: {}".format(name, ex))
            msg = "Exported {} template(s) to:\n{}".format(written, dest)
            if skipped: msg += "\n\nSkipped (missing): {}".format(skipped)
            if errors:  msg += "\n\nErrors:\n" + "\n".join(errors)
            self.TxtStatus.Text = "Exported {}.".format(written)
            forms.alert(msg, title="Export Templates")
        except Exception as ex:
            forms.alert("Export failed: {}".format(ex))


def show_export(owner_window, templates_folder):
    dlg = _ExportDialog(templates_folder)
    if owner_window is not None:
        try: dlg.win.Owner = owner_window
        except Exception: pass
    dlg.win.ShowDialog()


# ── IMPORT DIALOG ─────────────────────────────────────────────────────────────

IMPORT_XAML = u"""<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="pyFilter | Import" Height="620" Width="680"
    WindowStartupLocation="CenterOwner"
    Background="#1E2530" Foreground="#F4FAFF">
""" + _COMMON_STYLES + u"""
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="60"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Background="#161C24">
            <Grid Margin="16,0">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="py" FontWeight="Bold" FontSize="22"
                               Foreground="#208A3C"/>
                    <TextBlock Text="Filter" FontSize="22" FontWeight="SemiBold"
                               Margin="3,0,0,0"/>
                    <TextBlock Text=" | " FontSize="18" Foreground="#4A5568" Margin="6,0"/>
                    <TextBlock Text="Import" FontSize="18" FontWeight="SemiBold"/>
                </StackPanel>
                <Button x:Name="BtnImportNow" Grid.Column="1"
                        Content="Import Now" Padding="20,8" Margin="0,0,10,0"/>
                <Button x:Name="BtnClose" Grid.Column="2"
                        Content="X" Style="{StaticResource GrayBtn}"
                        Width="36" Height="36" FontWeight="Bold"/>
            </Grid>
        </Border>

        <!-- Body -->
        <ScrollViewer Grid.Row="1" Style="{StaticResource DarkSV}" Padding="16,12">
            <StackPanel>
                <!-- Import Settings card -->
                <Border Background="#2B3340" CornerRadius="8" Padding="16" Margin="0,0,0,12">
                    <StackPanel>
                        <TextBlock Text="Import Settings" Foreground="#208A3C"
                                   FontSize="15" FontWeight="SemiBold" Margin="0,0,0,6"/>
                        <TextBlock Text="Browse to a folder containing pyFilter template .json files."
                                   Foreground="#9CA3AF" FontSize="12" TextWrapping="Wrap"
                                   Margin="0,0,0,14"/>
                        <CheckBox x:Name="ChkAutoImport" Content="Auto update on startup"
                                  Margin="0,0,0,4"/>
                        <TextBlock Text="When enabled, templates are automatically imported from this folder each time pyFilter opens."
                                   Foreground="#9CA3AF" FontSize="11" TextWrapping="Wrap"
                                   Margin="20,0,0,14"/>
                        <TextBlock Text="Import Folder" FontWeight="SemiBold" Margin="0,0,0,6"/>
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBox x:Name="TxtSource" Grid.Column="0" Height="36"/>
                            <Button x:Name="BtnBrowse" Grid.Column="1"
                                    Content="Browse" Padding="16,8" Margin="8,0,0,0"/>
                        </Grid>
                    </StackPanel>
                </Border>

                <!-- Template list card -->
                <Border Background="#2B3340" CornerRadius="8" Padding="16">
                    <StackPanel>
                        <TextBlock Text="Select templates to import" Foreground="#208A3C"
                                   FontSize="15" FontWeight="SemiBold" Margin="0,0,0,6"/>
                        <TextBlock Text="On name conflict:" Foreground="#9CA3AF"
                                   FontSize="11" Margin="0,0,0,8"/>
                        <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                            <RadioButton x:Name="RdoOverwrite" Content="Overwrite"
                                         IsChecked="True" GroupName="conflict" Margin="0,0,16,0"/>
                            <RadioButton x:Name="RdoSkip" Content="Skip"
                                         GroupName="conflict" Margin="0,0,16,0"/>
                            <RadioButton x:Name="RdoRename" Content="Rename"
                                         GroupName="conflict"/>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                            <Button x:Name="BtnAll"  Content="All"  Padding="14,6"
                                    Margin="0,0,8,0"/>
                            <Button x:Name="BtnNone" Content="None" Padding="14,6"
                                    Style="{StaticResource GrayBtn}"/>
                        </StackPanel>
                        <Border Background="#1E2530" CornerRadius="6" Padding="8"
                                Height="200">
                            <ScrollViewer Style="{StaticResource DarkSV}">
                                <StackPanel x:Name="TemplateList"/>
                            </ScrollViewer>
                        </Border>
                    </StackPanel>
                </Border>
            </StackPanel>
        </ScrollViewer>

        <!-- Footer -->
        <Border Grid.Row="2" Background="#161C24" Padding="16,10">
            <TextBlock x:Name="TxtStatus"
                       Text="Settings are saved automatically when you close this panel."
                       Foreground="#4A5568" FontSize="11"/>
        </Border>
    </Grid>
</Window>"""


class _ImportDialog(object):
    def __init__(self, templates_folder, on_changed=None):
        self.templates_folder = templates_folder
        self.on_changed = on_changed
        self._checks   = {}
        self.win       = XamlReader.Parse(IMPORT_XAML)
        for n in ("TxtSource", "BtnBrowse", "BtnAll", "BtnNone", "ChkAutoImport",
                  "TemplateList", "TxtStatus", "RdoOverwrite", "RdoSkip", "RdoRename",
                  "BtnClose", "BtnImportNow"):
            setattr(self, n, self.win.FindName(n))

        self.BtnBrowse.Click     += self._on_browse
        self.BtnAll.Click        += lambda s, e: self._set_all(True)
        self.BtnNone.Click       += lambda s, e: self._set_all(False)
        self.BtnImportNow.Click  += self._on_import
        self.BtnClose.Click      += lambda s, e: self.win.Close()
        self.TxtSource.TextChanged += lambda s, e: self._populate()
        self._populate()

    def _list_source(self):
        src = (self.TxtSource.Text or "").strip()
        if not src or not os.path.isdir(src):
            return []
        out = []
        for f in os.listdir(src):
            if f.lower().endswith(".json"):
                out.append(os.path.splitext(f)[0])
        return sorted(out)

    def _populate(self):
        self.TemplateList.Children.Clear()
        self._checks = {}
        from System.Windows.Controls import CheckBox as WpfCheckBox
        from System.Windows.Media import SolidColorBrush, Color
        from System.Windows import Thickness
        names = self._list_source()
        local = set(_list_local_templates_internal(self.templates_folder))
        if not names:
            from System.Windows.Controls import TextBlock as TB
            tb = TB()
            tb.Text = "No .json files found in source folder."
            tb.Foreground = SolidColorBrush(Color.FromRgb(156, 163, 175))
            self.TemplateList.Children.Add(tb)
            return
        for name in names:
            cb = WpfCheckBox()
            exists = name in local
            cb.Content   = u"{}{}".format(name, u"  \u2713 exists" if exists else u"  (new)")
            cb.Tag       = name
            cb.IsChecked = True
            cb.Margin    = Thickness(0, 3, 0, 3)
            self.TemplateList.Children.Add(cb)
            self._checks[name] = cb

    def _set_all(self, value):
        for cb in self._checks.values():
            cb.IsChecked = bool(value)

    def _on_browse(self, sender, e):
        folder = forms.pick_folder(title="Pick source folder")
        if folder:
            self.TxtSource.Text = folder

    def _conflict_mode(self):
        if self.RdoOverwrite.IsChecked: return "overwrite"
        if self.RdoSkip.IsChecked:      return "skip"
        return "rename"

    def _unique_name(self, base, taken):
        i = 2
        while True:
            cand = u"{} ({})".format(base, i)
            if cand not in taken:
                return cand
            i += 1

    def _on_import(self, sender, e):
        try:
            src = (self.TxtSource.Text or "").strip()
            if not src or not os.path.isdir(src):
                forms.alert("Pick a valid source folder.")
                return
            selected = [n for n, cb in self._checks.items() if cb.IsChecked]
            if not selected:
                forms.alert("Tick at least one template to import.")
                return
            import shutil
            local = set(_list_local_templates_internal(self.templates_folder))
            mode  = self._conflict_mode()
            taken = set(local)
            imported = skipped = renamed = overwritten = 0
            errors = []
            for name in selected:
                src_path = os.path.join(src, name + ".json")
                if not os.path.isfile(src_path):
                    continue
                target_name = name
                if name in local:
                    if mode == "skip":
                        skipped += 1
                        continue
                    if mode == "rename":
                        target_name = self._unique_name(name, taken)
                tgt_path = os.path.join(self.templates_folder, target_name + ".json")
                try:
                    shutil.copy2(src_path, tgt_path)
                    taken.add(target_name)
                    if target_name != name:
                        renamed += 1
                    elif name in local:
                        overwritten += 1
                    else:
                        imported += 1
                except Exception as ex:
                    errors.append(u"{}: {}".format(name, ex))
            msg = u"Imported: {} new, {} overwritten, {} renamed, {} skipped.".format(
                imported, overwritten, renamed, skipped)
            if errors:
                msg += u"\n\nErrors:\n" + u"\n".join(errors)
            self.TxtStatus.Text = msg
            forms.alert(msg, title="Import Templates")
            if self.on_changed:
                self.on_changed()
        except Exception as ex:
            forms.alert("Import failed: {}".format(ex))


def show_import(owner_window, templates_folder, on_changed=None):
    dlg = _ImportDialog(templates_folder, on_changed=on_changed)
    if owner_window is not None:
        try: dlg.win.Owner = owner_window
        except Exception: pass
    dlg.win.ShowDialog()


# ── INLINE SYNC SETTINGS PANEL ───────────────────────────────────────────────

def build_sync_settings_panel(container, templates_folder, on_changed=None):
    """
    Build the sync settings UI directly into a WPF StackPanel (container).
    No separate Window needed — matches pyTransmit inline panel pattern.
    """
    from System.Windows.Controls import (Button, CheckBox as WpfCheckBox,
                                         TextBox, StackPanel, ScrollViewer,
                                         TextBlock as TB)
    from System.Windows import Thickness, Visibility
    from System.Windows.Media import SolidColorBrush, Color

    container.Children.Clear()

    fg  = SolidColorBrush(Color.FromRgb(244, 250, 255))
    muted = SolidColorBrush(Color.FromRgb(156, 163, 175))

    def make_tb(text, size=12, opacity=1.0, weight="Normal", margin=None, color=None):
        t = TB()
        t.Text = text
        t.FontSize = size
        t.Opacity  = opacity
        t.Foreground = color or fg
        if margin:
            t.Margin = Thickness(*margin)
        from System.Windows import TextWrapping
        t.TextWrapping = TextWrapping.Wrap
        return t

    def section_label(text):
        t = TB()
        t.Text = text
        t.FontSize = 14
        t.FontWeight = bold_weight()
        from System.Windows.Media import SolidColorBrush, Color
        t.Foreground = SolidColorBrush(Color.FromRgb(32, 138, 60))
        t.Margin = Thickness(0, 0, 0, 8)
        return t

    def bold_weight():
        from System.Windows import FontWeights
        return FontWeights.SemiBold

    def card(children):
        from System.Windows.Controls import Border
        from System.Windows import CornerRadius
        b = Border()
        b.Background = SolidColorBrush(Color.FromRgb(43, 51, 64))
        b.CornerRadius = CornerRadius(8)
        b.Padding = Thickness(16)
        b.Margin  = Thickness(0, 0, 0, 12)
        sp = StackPanel()
        for c in children: sp.Children.Add(c)
        b.Child = sp
        return b

    # Sync server path card
    server_path_box = TextBox()
    server_path_box.Text = get_server_path()
    server_path_box.Height = 28
    server_path_box.Background = SolidColorBrush(Color.FromRgb(244, 250, 255))
    server_path_box.Foreground = SolidColorBrush(Color.FromRgb(43, 51, 64))
    from System.Windows.Media import SolidColorBrush as SB2, Color as C2
    server_path_box.BorderBrush = SB2(C2.FromRgb(32, 138, 60))
    server_path_box.BorderThickness = Thickness(1)
    server_path_box.Padding = Thickness(8, 4, 8, 4)
    server_path_box.FontSize = 12

    status_tb = make_tb("", 11, margin=(0, 4, 0, 0), color=muted)

    def refresh_status(path=None):
        if path is None:
            path = server_path_box.Text
        ok, msg = server_ready(path)
        if ok:
            n = len(list_server_templates(path))
            status_tb.Text = u"Reachable. {} template file(s) on server.".format(n)
        else:
            status_tb.Text = u"Warning: {}".format(msg) if path else "No path set."

    refresh_status(server_path_box.Text)
    server_path_box.TextChanged += lambda s, e: refresh_status(s.Text)

    browse_btn = Button()
    browse_btn.Content = "Browse"
    browse_btn.Height = 24
    browse_btn.Padding = Thickness(12, 0, 12, 0)
    browse_btn.FontSize = 11
    browse_btn.Margin = Thickness(8, 0, 0, 0)
    browse_btn.Background = SolidColorBrush(Color.FromRgb(32, 138, 60))
    browse_btn.Foreground = fg
    browse_btn.BorderThickness = Thickness(0)

    def on_browse(s, e):
        folder = forms.pick_folder(title="Pick sync folder")
        if folder:
            server_path_box.Text = folder
    browse_btn.Click += on_browse

    from System.Windows.Controls import DockPanel
    from System.Windows.Controls import DockPanel as DP
    path_row = DockPanel()
    DP.SetDock(browse_btn, 1)  # Right
    path_row.Children.Add(browse_btn)
    path_row.Children.Add(server_path_box)

    # Synced templates card
    checks = {}
    tpl_panel = StackPanel()

    def refresh_tpl_list():
        tpl_panel.Children.Clear()
        checks.clear()
        local_names  = set(_list_local_templates_internal(templates_folder))
        server_names = set(list_server_templates(server_path_box.Text))
        synced       = set(get_synced_templates())
        all_names    = sorted(local_names | server_names | synced)
        if not all_names:
            t = TB(); t.Text = "No templates found."
            t.Foreground = muted; t.FontSize = 12
            tpl_panel.Children.Add(t)
            return
        for name in all_names:
            cb = WpfCheckBox()
            tag = []
            if name in local_names:  tag.append("local")
            if name in server_names: tag.append("server")
            cb.Content   = u"{}   ({})".format(name, u", ".join(tag)) if tag else name
            cb.Tag       = name
            cb.IsChecked = name in synced
            cb.Foreground = fg
            cb.FontSize  = 12
            cb.Margin    = Thickness(0, 3, 0, 3)
            tpl_panel.Children.Add(cb)
            checks[name] = cb

    server_path_box.TextChanged += lambda s, e: refresh_tpl_list()
    refresh_tpl_list()

    sel_all_btn = Button()
    sel_all_btn.Content = "Select all"; sel_all_btn.Height = 24
    sel_all_btn.Padding = Thickness(10, 0, 10, 0); sel_all_btn.FontSize = 11
    sel_all_btn.Background = SolidColorBrush(Color.FromRgb(32, 138, 60))
    sel_all_btn.Foreground = fg; sel_all_btn.BorderThickness = Thickness(0)
    sel_all_btn.Margin = Thickness(0, 0, 6, 8)
    sel_all_btn.Click += lambda s, e: [setattr(cb, "IsChecked", True)
                                        for cb in checks.values()]

    clear_btn = Button()
    clear_btn.Content = "Clear"; clear_btn.Height = 24
    clear_btn.Padding = Thickness(10, 0, 10, 0); clear_btn.FontSize = 11
    clear_btn.Background = SolidColorBrush(Color.FromRgb(64, 69, 83))
    clear_btn.Foreground = fg; clear_btn.BorderThickness = Thickness(0)
    clear_btn.Margin = Thickness(0, 0, 0, 8)
    clear_btn.Click += lambda s, e: [setattr(cb, "IsChecked", False)
                                      for cb in checks.values()]

    sel_row = StackPanel(); sel_row.Orientation = 1  # Horizontal
    sel_row.Children.Add(sel_all_btn); sel_row.Children.Add(clear_btn)

    sync_now_btn = Button()
    sync_now_btn.Content = "Sync now"; sync_now_btn.Height = 28
    sync_now_btn.Padding = Thickness(16, 0, 16, 0); sync_now_btn.FontSize = 12
    sync_now_btn.Background = SolidColorBrush(Color.FromRgb(32, 138, 60))
    sync_now_btn.Foreground = fg; sync_now_btn.BorderThickness = Thickness(0)
    sync_now_btn.Margin = Thickness(0, 12, 0, 0)

    sync_status_tb = make_tb("", 11, margin=(8, 8, 0, 0), color=muted)

    def on_sync_now(s, e):
        set_server_path(server_path_box.Text.strip())
        set_synced_templates([n for n, cb in checks.items() if cb.IsChecked])
        msgs = []
        result = sync_all(templates_folder, logger=lambda m: msgs.append(m))
        if not result.get("ok"):
            sync_status_tb.Text = u"Skipped: {}".format(result.get("reason"))
        else:
            counts = result.get("counts", {})
            sync_status_tb.Text = u", ".join(
                u"{}:{}".format(k, v) for k, v in sorted(counts.items())) or "Nothing to sync."
        refresh_tpl_list()
        if on_changed: on_changed()

    sync_now_btn.Click += on_sync_now

    save_btn = Button()
    save_btn.Content = "Save"; save_btn.Height = 28
    save_btn.Padding = Thickness(20, 0, 20, 0); save_btn.FontSize = 12
    save_btn.Background = SolidColorBrush(Color.FromRgb(32, 138, 60))
    save_btn.Foreground = fg; save_btn.BorderThickness = Thickness(0)
    save_btn.Margin = Thickness(8, 12, 0, 0)

    def on_save(s, e):
        set_server_path(server_path_box.Text.strip())
        set_synced_templates([n for n, cb in checks.items() if cb.IsChecked])
        if on_changed: on_changed()

    save_btn.Click += on_save

    action_row = StackPanel(); action_row.Orientation = 1
    action_row.Children.Add(sync_now_btn)
    action_row.Children.Add(save_btn)
    action_row.Children.Add(sync_status_tb)

    container.Children.Add(card([
        section_label("Sync Settings"),
        make_tb("Server path", 12, 0.9, margin=(0, 0, 0, 4)),
        path_row,
        status_tb,
    ]))
    container.Children.Add(card([
        section_label("Templates to sync"),
        sel_row,
        tpl_panel,
        action_row,
    ]))
