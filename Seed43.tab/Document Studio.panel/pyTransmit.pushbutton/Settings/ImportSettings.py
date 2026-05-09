# -*- coding: utf-8 -*-
"""
ImportSettings.py  —  pyTransmit Import Settings panel controller
=================================================================
Manages the Import Settings panel embedded in the main pyTransmit window.

Handles:
  - Import folder selection (expects a  pyTransmit Settings  folder)
  - Per-item checkboxes (Recipients, Reason, Method, Document, Print Size)
  - Auto-update toggle (silently re-imports from the folder on startup)
  - Execute import (copies selected JSON files from the folder into Settings/)
  - Reloading live controller data after import

Config is persisted in pytransmit_sync.json alongside the main script.

Place this file in the  Settings  subfolder next to script.py.

Usage in script.py:
    from ImportSettings import ImportSettingsController
    self.import_ctrl = ImportSettingsController(script_dir)
    self.import_ctrl.attach(self)
    self.import_ctrl.load_config()
    self.import_ctrl.run_auto_import()   # call after controllers are ready
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

class ImportSettingsController(object):
    """
    Drives the Import Settings panel embedded in pyTransmit.
    Reads/writes the 'import_*' and 'imp_*' keys of pytransmit_sync.json.
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
        self._script_dir   = script_dir
        self._sync_path    = os.path.join(script_dir, self.SYNC_FILE)
        self._settings_dir = os.path.join(script_dir, 'Settings')
        self._host         = None

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
        try: h.import_path_tb.Text              = cfg.get('import_path', '')
        except: pass
        try: h.import_auto_update_cb.IsChecked  = bool(cfg.get('import_auto', False))
        except: pass
        try: h.import_recipients_cb.IsChecked   = bool(cfg.get('imp_recipients', True))
        except: pass
        try: h.import_reason_cb.IsChecked       = bool(cfg.get('imp_reason', True))
        except: pass
        try: h.import_method_cb.IsChecked       = bool(cfg.get('imp_method', True))
        except: pass
        try: h.import_format_cb.IsChecked       = bool(cfg.get('imp_format', True))
        except: pass
        try: h.import_printsize_cb.IsChecked    = bool(cfg.get('imp_printsize', True))
        except: pass

    def save_config(self):
        """Read the current panel state and merge into the shared sync config file."""
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
            'import_path':    tb('import_path_tb'),
            'import_auto':    cb('import_auto_update_cb', False),
            'imp_recipients': cb('import_recipients_cb'),
            'imp_reason':     cb('import_reason_cb'),
            'imp_method':     cb('import_method_cb'),
            'imp_format':     cb('import_format_cb'),
            'imp_printsize':  cb('import_printsize_cb'),
        })
        try:
            with open(self._sync_path, 'w') as f:
                json.dump(cfg, f, indent=2)
        except:
            pass

    # ── Core import logic ─────────────────────────────────────────────────────

    def do_import(self, folder_path, selections):
        """
        Copy selected JSON files from folder_path into Settings/.
        folder_path should be (or contain) a  pyTransmit Settings  folder.
        Returns list of copied keys.
        """
        if not os.path.exists(self._settings_dir):
            try: os.makedirs(self._settings_dir)
            except: pass
        copied = []
        for key, filename in self.SETTINGS_FILENAMES.items():
            if selections.get(key):
                src = os.path.join(folder_path, filename)
                if os.path.exists(src):
                    shutil.copy2(src, os.path.join(self._settings_dir, filename))
                    copied.append(key)
        return copied

    def run_auto_import(self):
        """
        If auto-import is enabled and the path is valid, silently import on startup.
        Call after controllers (rec_ctrl, opt_ctrl) are fully initialised.
        """
        h = self._host
        if h is None:
            return
        try:
            auto = getattr(h, 'import_auto_update_cb', None)
            if auto is None or not auto.IsChecked:
                return
            path = getattr(h, 'import_path_tb', None)
            if path is None or not path.Text:
                return
            folder = path.Text
            if not os.path.exists(folder):
                return
            selections = self._read_selections()
            copied = self.do_import(folder, selections)
            if copied:
                self._reload_controllers(selections)
        except:
            pass

    # ── Event handlers ────────────────────────────────────────────────────────

    def _wire_events(self):
        # Browse and Execute buttons are wired in XAML — don't bind here again (double-fire).
        # Wire LostFocus on the path textbox so typed/pasted paths are saved immediately.
        h = self._host
        if h is None:
            return
        tb = getattr(h, 'import_path_tb', None)
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
            dlg.Description = "Select the 'pyTransmit Settings' folder to import from"
            dlg.ShowNewFolderButton = False
            h = self._host
            current = getattr(h, 'import_path_tb', None)
            if current and current.Text and os.path.exists(current.Text):
                dlg.SelectedPath = current.Text
            if dlg.ShowDialog() == DialogResult.OK:
                h.import_path_tb.Text = dlg.SelectedPath
                self.save_config()
        except Exception as e:
            self._alert("Browse Error", str(e))

    def on_execute(self, sender, args):
        h = self._host
        if h is None:
            return
        try:
            path = getattr(h, 'import_path_tb', None)
            if path is None or not path.Text:
                self._alert("Import", "Please choose an import folder first.")
                return
            folder = path.Text
            if not os.path.exists(folder):
                self._alert("Import", "Folder does not exist:\n{}".format(folder))
                return
            selections = self._read_selections()
            if not any(selections.values()):
                self._alert("Import", "Please select at least one item to import.")
                return
            copied = self.do_import(folder, selections)
            if not copied:
                self._alert("Import",
                            "No matching files found in:\n{}".format(folder))
                return
            self._reload_controllers(selections)
            self.save_config()
            self._alert("Import Complete",
                        "Imported {} item(s) successfully.".format(len(copied)))
        except Exception as e:
            self._alert("Import Error", str(e))

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
            'recipients': cb('import_recipients_cb'),
            'reason':     cb('import_reason_cb'),
            'method':     cb('import_method_cb'),
            'format':     cb('import_format_cb'),
            'printsize':  cb('import_printsize_cb'),
        }

    def _reload_controllers(self, selections):
        """Reload live controller data from disk for any imported items."""
        h = self._host
        if h is None:
            return
        try:
            if selections.get('recipients'):
                rc = getattr(h, 'rec_ctrl', None)
                if rc:
                    rc.data.Clear()
                    for r in rc.db.load_all():
                        rc.data.Add(r)
                    rc.dist_data.Clear()
                    for r in rc.dist_db.load_all():
                        rc.dist_data.Add(r)
        except: pass
        try:
            opt = getattr(h, 'opt_ctrl', None)
            if opt:
                if selections.get('reason'):   opt.reason_data.Clear()
                if selections.get('method'):   opt.method_data.Clear()
                if selections.get('format'):   opt.format_data.Clear()
                if selections.get('printsize'):opt.printsize_data.Clear()
                opt._load_all()
        except: pass

    def _alert(self, title, message):
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
