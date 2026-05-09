# -*- coding: utf-8 -*-
"""
ExportSettings.py  —  pyTransmit Export Settings panel controller
=================================================================
Manages the Export Settings panel embedded in the main pyTransmit window.

Handles:
  - Export folder selection
  - Per-item checkboxes (Recipients, Reason, Method, Document, Print Size)
  - Auto-update toggle (silently re-exports on every Save)
  - Execute export (copies selected JSON files to <folder>/pyTransmit Settings/)

Config is persisted in pytransmit_sync.json alongside the main script.

Place this file in the  Settings  subfolder next to script.py.

Usage in script.py:
    from ExportSettings import ExportSettingsController
    self.export_ctrl = ExportSettingsController(script_dir)
    self.export_ctrl.attach(self)
    self.export_ctrl.load_config()
"""

import os
import json
import shutil

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

# ══════════════════════════════════════════════════════════════════════════════
# CONTROLLER
# ══════════════════════════════════════════════════════════════════════════════

class ExportSettingsController(object):
    """
    Drives the Export Settings panel embedded in pyTransmit.
    Reads/writes the 'export_*' and 'exp_*' keys of pytransmit_sync.json.
    """

    SYNC_FILE = 'pytransmit_sync.json'

    SETTINGS_FILENAMES = {
        'recipients': 'recipients.json',
        'reason':     'reason.json',
        'method':     'method.json',
        'format':     'format.json',
        'printsize':  'printsize.json',
    }

    def __init__(self, script_dir):
        self._script_dir  = script_dir
        self._sync_path   = os.path.join(script_dir, self.SYNC_FILE)
        self._settings_dir = os.path.join(script_dir, 'Settings')
        self._host        = None

    # ── Attach ────────────────────────────────────────────────────────────────

    def attach(self, host):
        """Attach to the host WPFWindow and wire button events."""
        self._host = host
        self._wire_events()

    def load_config(self):
        """Read saved config from disk and push values into the panel controls."""
        try:
            with open(self._sync_path, 'r') as f:
                cfg = json.load(f)
        except:
            cfg = {}
        h = self._host
        if h is None:
            return
        try: h.export_path_tb.Text              = cfg.get('export_path', '')
        except: pass
        try: h.export_auto_update_cb.IsChecked  = bool(cfg.get('export_auto', False))
        except: pass
        try: h.export_recipients_cb.IsChecked   = bool(cfg.get('exp_recipients', True))
        except: pass
        try: h.export_reason_cb.IsChecked       = bool(cfg.get('exp_reason', True))
        except: pass
        try: h.export_method_cb.IsChecked       = bool(cfg.get('exp_method', True))
        except: pass
        try: h.export_format_cb.IsChecked       = bool(cfg.get('exp_format', True))
        except: pass
        try: h.export_printsize_cb.IsChecked    = bool(cfg.get('exp_printsize', True))
        except: pass

    def save_config(self):
        """Read the current panel state and merge into the shared sync config file."""
        # Load existing (may have import keys we must preserve)
        try:
            with open(self._sync_path, 'r') as f:
                cfg = json.load(f)
        except:
            cfg = {}
        h = self._host
        if h is None:
            return
        def cb(name, default=True):
            el = getattr(h, name, None)
            if el is None: return default
            try:
                v = el.IsChecked
                return bool(v) if v is not None else default
            except: return default
        def tb(name):
            el = getattr(h, name, None)
            return el.Text if el is not None else ''
        cfg.update({
            'export_path':    tb('export_path_tb'),
            'export_auto':    cb('export_auto_update_cb', False),
            'exp_recipients': cb('export_recipients_cb'),
            'exp_reason':     cb('export_reason_cb'),
            'exp_method':     cb('export_method_cb'),
            'exp_format':     cb('export_format_cb'),
            'exp_printsize':  cb('export_printsize_cb'),
        })
        try:
            with open(self._sync_path, 'w') as f:
                json.dump(cfg, f, indent=2)
        except:
            pass

    # ── Core export logic ─────────────────────────────────────────────────────

    def do_export(self, folder_path, selections):
        """
        Copy selected JSON files from Settings/ into <folder_path>/pyTransmit Settings/.
        Returns (dest_folder, list_of_copied_keys).
        """
        dest = os.path.join(folder_path, 'pyTransmit Settings')
        if not os.path.exists(dest):
            os.makedirs(dest)
        copied = []
        for key, filename in self.SETTINGS_FILENAMES.items():
            if selections.get(key):
                src = os.path.join(self._settings_dir, filename)
                if os.path.exists(src):
                    shutil.copy2(src, os.path.join(dest, filename))
                    copied.append(key)
        return dest, copied

    def auto_export_if_enabled(self):
        """Silently re-export if auto-update is enabled and path is valid."""
        h = self._host
        if h is None:
            return
        try:
            auto = getattr(h, 'export_auto_update_cb', None)
            if auto is None or not auto.IsChecked:
                return
            path = getattr(h, 'export_path_tb', None)
            if path is None or not path.Text or not os.path.exists(path.Text):
                return
            selections = self._read_selections()
            self.do_export(path.Text, selections)
        except:
            pass

    # ── Event handlers ────────────────────────────────────────────────────────

    def _wire_events(self):
        # Browse and Execute buttons are wired in XAML — don't bind here again (double-fire).
        # Wire LostFocus on the path textbox so typed/pasted paths are saved immediately.
        h = self._host
        if h is None:
            return
        tb = getattr(h, 'export_path_tb', None)
        if tb is not None:
            try: tb.LostFocus.__iadd__(self._on_path_changed)
            except: pass

    def _on_path_changed(self, sender, args):
        """Save config whenever the user finishes editing the path textbox."""
        self.save_config()

    def on_browse(self, sender, args):
        try:
            from System.Windows.Forms import FolderBrowserDialog, DialogResult
            dlg = FolderBrowserDialog()
            dlg.Description = "Select export destination folder"
            dlg.ShowNewFolderButton = True
            h = self._host
            current = getattr(h, 'export_path_tb', None)
            if current and current.Text and os.path.exists(current.Text):
                dlg.SelectedPath = current.Text
            if dlg.ShowDialog() == DialogResult.OK:
                h.export_path_tb.Text = dlg.SelectedPath
                self.save_config()
        except Exception as e:
            self._alert("Browse Error", str(e))

    def on_execute(self, sender, args):
        h = self._host
        if h is None:
            return
        try:
            path = getattr(h, 'export_path_tb', None)
            if path is None or not path.Text:
                self._alert("Export", "Please choose an export folder first.")
                return
            if not os.path.exists(path.Text):
                self._alert("Export", "Folder does not exist:\n{}".format(path.Text))
                return
            selections = self._read_selections()
            if not any(selections.values()):
                self._alert("Export", "Please select at least one item to export.")
                return
            dest, copied = self.do_export(path.Text, selections)
            self.save_config()
            self._alert("Export Complete",
                        "Exported {} item(s) to:\n{}".format(len(copied), dest))
        except Exception as e:
            self._alert("Export Error", str(e))

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _read_selections(self):
        h = self._host
        def cb(name):
            el = getattr(h, name, None)
            if el is None: return False
            try:
                v = el.IsChecked
                return bool(v) if v is not None else False
            except: return False
        return {
            'recipients': cb('export_recipients_cb'),
            'reason':     cb('export_reason_cb'),
            'method':     cb('export_method_cb'),
            'format':     cb('export_format_cb'),
            'printsize':  cb('export_printsize_cb'),
        }

    def _alert(self, title, message):
        """Themed info dialog; falls back to forms.alert."""
        try:
            import System.Windows.Markup as Markup
            xaml = (
                '<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"'
                ' xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"'
                ' Title="" Width="360" SizeToContent="Height"'
                ' WindowStyle="None" ResizeMode="NoResize"'
                ' WindowStartupLocation="CenterScreen"'
                ' Background="Transparent" FontFamily="Segoe UI" AllowsTransparency="True">'
                '<Border Background="#2B3340" CornerRadius="10" Margin="12" Padding="24,20,24,20">'
                '<Border.Effect><DropShadowEffect Color="Black" Opacity="0.5" ShadowDepth="4" BlurRadius="16"/></Border.Effect>'
                '<StackPanel>'
                '<Border Background="#208A3C" Height="3" CornerRadius="2" Margin="0,0,0,16"/>'
                '<TextBlock x:Name="t" Foreground="#F4FAFF" FontSize="15" FontWeight="Bold" Margin="0,0,0,8"/>'
                '<TextBlock x:Name="m" Foreground="#F4FAFF" FontSize="12" Opacity="0.85" TextWrapping="Wrap" Margin="0,0,0,24"/>'
                '<StackPanel Orientation="Horizontal" HorizontalAlignment="Right">'
                '<Button x:Name="ok" Content="OK" Foreground="#F4FAFF" FontSize="12" FontWeight="Bold"'
                ' BorderThickness="0" Padding="28,8" Cursor="Hand">'
                '<Button.Template><ControlTemplate TargetType="Button">'
                '<Border x:Name="Bd" Background="#208A3C" CornerRadius="6" Padding="{TemplateBinding Padding}">'
                '<ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>'
                '</Border></ControlTemplate></Button.Template>'
                '</Button></StackPanel></StackPanel></Border></Window>'
            )
            dlg = Markup.XamlReader.Parse(xaml)
            dlg.FindName("t").Text = title
            dlg.FindName("m").Text = message
            def on_ok(s, e): dlg.Close()
            dlg.FindName("ok").Click += on_ok
            dlg.ShowDialog()
        except:
            try:
                from pyrevit import forms
                forms.alert("{}\n\n{}".format(title, message))
            except:
                pass
