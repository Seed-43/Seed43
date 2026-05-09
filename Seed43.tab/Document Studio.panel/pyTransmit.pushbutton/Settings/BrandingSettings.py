# -*- coding: utf-8 -*-
"""
BrandingSettings.py  —  pyTransmit Branding & Styling panel controller
=======================================================================
Manages the Branding & Styling panel embedded in the main pyTransmit window.

Handles:
  - Logo source path  (network / shared drive — authoritative copy)
    Auto-synced to Settings/logo.<ext> on startup and each time panel opens.
    The user sets the source path only. The local copy location is automatic
    (always Settings/ next to this script). No manual local path or save button.
  - Title bar background colour  (hex #RRGGBB)
  - Column header background colour (hex #RRGGBB)

Config is persisted in Settings/branding.json.

Startup behaviour
-----------------
  auto_sync_logo()  is called early in script.py __init__.
  It silently copies from source to Settings/ if reachable.
  If unreachable, the existing Settings/logo file is kept and used.

Save behaviour (matches SetupSettings / RecipientSettings pattern)
-----------------
  styling_back_click in script.py calls  brand_ctrl.save_and_back()
  which saves config then returns to main panel — no Save button needed.
"""

import os
import json
import shutil

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

import System.Windows.Media as _SWM


# ══════════════════════════════════════════════════════════════════════════════
# CONTROLLER
# ══════════════════════════════════════════════════════════════════════════════

class BrandingSettingsController(object):
    """
    Drives the Branding & Styling panel embedded in pyTransmit.
    Config is stored in  Settings/branding.json.
    """

    CONFIG_FILE = 'branding.json'

    DEFAULT_CFG = {
        'logo_source':     '',          # network / shared path to authoritative logo
        'title_bg_color':  '#FFFFFF',   # white for title bar
        'title_fg_color':  '#000000',   # black text for title bar
        'header_bg_color': '#FFFFFF',   # white background for column headers
        'header_fg_color': '#000000',   # black text for column headers
    }

    def __init__(self, script_dir):
        self._script_dir   = script_dir
        self._settings_dir = os.path.join(script_dir, 'Settings')
        self._config_path  = os.path.join(self._settings_dir, self.CONFIG_FILE)
        self._cfg          = dict(self.DEFAULT_CFG)
        self._host         = None
        self._load_config()

    # ── Attach ────────────────────────────────────────────────────────────────

    def attach(self, host):
        """Attach to the host WPFWindow. Call once after WPFWindow.__init__."""
        self._host = host
        self._wire_events()

    # ── Auto-sync on startup ──────────────────────────────────────────────────

    def auto_sync_logo(self):
        """
        Called once on startup (before window shown).
        If source is configured and reachable, copies to Settings/logo.<ext>.
        Silent — never raises.  Returns local logo path or ''.
        """
        try:
            self._do_sync_logo()
        except Exception:
            pass
        return self.get_logo_path()

    def _do_sync_logo(self):
        """Internal: copy source → Settings/logo.<ext> if source is reachable."""
        src = self._cfg.get('logo_source', '').strip()
        if not src or not os.path.exists(src):
            return
        ext = os.path.splitext(src)[1].lower() or '.png'
        dst = os.path.join(self._settings_dir, 'logo' + ext)
        if not os.path.exists(self._settings_dir):
            os.makedirs(self._settings_dir)
        shutil.copy2(src, dst)

    # ── Public API used by script.py at generate-time ────────────────────────

    def get_logo_path(self):
        """Return the best available local logo path, or '' if none."""
        # Prefer Settings/ folder first
        for fname in ('logo.png', 'logo.PNG', 'logo.jpg', 'logo.JPG',
                      'logo.jpeg', 'logo.JPEG', 'Logo.png', 'Logo.jpg'):
            p = os.path.join(self._settings_dir, fname)
            if os.path.exists(p):
                return p
        # Fallback: script root folder
        for fname in ('logo.png', 'logo.PNG', 'logo.jpg', 'logo.JPG',
                      'logo.jpeg', 'logo.JPEG', 'Logo.png', 'Logo.jpg'):
            p = os.path.join(self._script_dir, fname)
            if os.path.exists(p):
                return p
        return ''

    def get_title_bg_color(self):
        return self._cfg.get('title_bg_color',  self.DEFAULT_CFG['title_bg_color'])

    def get_title_fg_color(self):
        return self._cfg.get('title_fg_color',  self.DEFAULT_CFG['title_fg_color'])

    def get_header_bg_color(self):
        return self._cfg.get('header_bg_color', self.DEFAULT_CFG['header_bg_color'])

    def get_header_fg_color(self):
        return self._cfg.get('header_fg_color', self.DEFAULT_CFG['header_fg_color'])

    # ── Save on panel close (called by styling_back_click in script.py) ───────

    def save_and_back(self):
        """
        Read current panel UI state → save to branding.json → attempt logo sync.
        Returns True; caller then calls _show_panel("main").
        """
        self._read_ui_to_cfg()
        self._save_config()
        try:
            self._do_sync_logo()
        except Exception:
            pass
        return True

    def _read_ui_to_cfg(self):
        """Pull current control values into _cfg."""
        h = self._host
        if h is None:
            return
        try: self._cfg['logo_source']     = (h.logo_source_tb.Text  or '').strip()
        except: pass
        try: self._cfg['title_bg_color']  = (h.title_color_tb.Text  or self.DEFAULT_CFG['title_bg_color']).strip()
        except: pass
        try: self._cfg['title_fg_color']  = (h.title_fg_color_tb.Text or self.DEFAULT_CFG['title_fg_color']).strip()
        except: pass
        try: self._cfg['header_bg_color'] = (h.header_color_tb.Text or self.DEFAULT_CFG['header_bg_color']).strip()
        except: pass
        try: self._cfg['header_fg_color'] = (h.header_fg_color_tb.Text or self.DEFAULT_CFG['header_fg_color']).strip()
        except: pass

    # ── Persistence ───────────────────────────────────────────────────────────

    def _load_config(self):
        try:
            with open(self._config_path, 'r') as f:
                loaded = json.load(f)
            cfg = dict(self.DEFAULT_CFG)
            cfg.update(loaded)
            self._cfg = cfg
        except Exception:
            self._cfg = dict(self.DEFAULT_CFG)

    def _save_config(self):
        try:
            if not os.path.exists(self._settings_dir):
                os.makedirs(self._settings_dir)
            with open(self._config_path, 'w') as f:
                json.dump(self._cfg, f, indent=2)
        except Exception:
            pass

    # ── UI population ─────────────────────────────────────────────────────────

    def load_panel(self):
        """Push _cfg values into panel controls. Called when panel opens."""
        if self._host is None:
            return
        h = self._host
        try: h.logo_source_tb.Text = self._cfg.get('logo_source', '')
        except: pass
        try:
            tc = self._cfg.get('title_bg_color',  self.DEFAULT_CFG['title_bg_color'])
            h.title_color_tb.Text = tc
            self._update_swatch(h.title_color_swatch, tc)
        except: pass
        try:
            tfc = self._cfg.get('title_fg_color', self.DEFAULT_CFG['title_fg_color'])
            h.title_fg_color_tb.Text = tfc
            self._update_swatch(h.title_fg_color_swatch, tfc)
        except: pass
        try:
            hc = self._cfg.get('header_bg_color', self.DEFAULT_CFG['header_bg_color'])
            h.header_color_tb.Text = hc
            self._update_swatch(h.header_color_swatch, hc)
        except: pass
        try:
            hfc = self._cfg.get('header_fg_color', self.DEFAULT_CFG['header_fg_color'])
            h.header_fg_color_tb.Text = hfc
            self._update_swatch(h.header_fg_color_swatch, hfc)
        except: pass
        try: h.logo_status_lbl.Text = ''
        except: pass

    def _update_swatch(self, swatch_border, hex_color):
        try:
            c = _SWM.ColorConverter.ConvertFromString(hex_color)
            swatch_border.Background = _SWM.SolidColorBrush(c)
        except Exception:
            pass

    # ── Event wiring ──────────────────────────────────────────────────────────

    def _wire_events(self):
        h = self._host
        if h is None:
            return

        def bind(name, event, handler):
            el = getattr(h, name, None)
            if el is not None:
                try: getattr(el, event).__iadd__(handler)
                except: pass

        bind('title_color_tb',          'TextChanged', self._on_title_color_changed)
        bind('title_fg_color_tb',       'TextChanged', self._on_title_fg_color_changed)
        bind('header_color_tb',         'TextChanged', self._on_header_color_changed)
        bind('header_fg_color_tb',      'TextChanged', self._on_header_fg_color_changed)
        bind('logo_source_browse_btn',  'Click',       self.on_logo_source_browse)
        bind('title_color_reset_btn',   'Click',       self.on_title_color_reset)
        bind('title_fg_color_reset_btn','Click',       self.on_title_fg_color_reset)
        bind('header_color_reset_btn',  'Click',       self.on_header_color_reset)
        bind('header_fg_color_reset_btn','Click',      self.on_header_fg_color_reset)

    # ── Event handlers ────────────────────────────────────────────────────────

    def _on_title_color_changed(self, sender, args):
        try: self._update_swatch(self._host.title_color_swatch, sender.Text)
        except: pass

    def _on_title_fg_color_changed(self, sender, args):
        try: self._update_swatch(self._host.title_fg_color_swatch, sender.Text)
        except: pass

    def _on_header_color_changed(self, sender, args):
        try: self._update_swatch(self._host.header_color_swatch, sender.Text)
        except: pass

    def _on_header_fg_color_changed(self, sender, args):
        try: self._update_swatch(self._host.header_fg_color_swatch, sender.Text)
        except: pass

    def on_title_color_reset(self, sender, args):
        try: self._host.title_color_tb.Text = self.DEFAULT_CFG['title_bg_color']
        except: pass

    def on_title_fg_color_reset(self, sender, args):
        try: self._host.title_fg_color_tb.Text = self.DEFAULT_CFG['title_fg_color']
        except: pass

    def on_header_color_reset(self, sender, args):
        try: self._host.header_color_tb.Text = self.DEFAULT_CFG['header_bg_color']
        except: pass

    def on_header_fg_color_reset(self, sender, args):
        try: self._host.header_fg_color_tb.Text = self.DEFAULT_CFG['header_fg_color']
        except: pass

    def on_logo_source_browse(self, sender, args):
        """Browse for the authoritative logo source file."""
        try:
            clr.AddReference('System.Windows.Forms')
            from System.Windows.Forms import OpenFileDialog, DialogResult
            dlg = OpenFileDialog()
            dlg.Title  = "Select Logo File"
            dlg.Filter = "Image files (*.png;*.jpg;*.jpeg)|*.png;*.jpg;*.jpeg|All files (*.*)|*.*"
            if dlg.ShowDialog() == DialogResult.OK:
                self._host.logo_source_tb.Text = dlg.FileName
                # Update status to show where it will be cached
                try:
                    ext = os.path.splitext(dlg.FileName)[1].lower() or '.png'
                    dst = os.path.join(self._settings_dir, 'logo' + ext)
                    self._host.logo_status_lbl.Text = u'\u2139\uFE0F  Will sync to: {}'.format(dst)
                except: pass
        except Exception, ex:
            try: self._host.logo_status_lbl.Text = u'\u274C  Browse error: {}'.format(str(ex))
            except: pass
