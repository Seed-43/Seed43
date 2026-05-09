# -*- coding: utf-8 -*-
"""
SetupSettings.py  —  pyTransmit Setup panel controller
=======================================================
Manages the Setup panel embedded in the main pyTransmit window.

Controls:
  - Project Settings: which optional fields appear in the main window
    (Reason for Issue, Method of Issue, Document Type, Print Size)
  - Recipient mode: Distribution List (fixed rows) vs Client List (cascading dropdowns)

Config is persisted to pytransmit_setup.json next to the main script.

Place this file in the  Settings  subfolder next to script.py.

Usage in script.py:
    from SetupSettings import SetupSettingsController
    self.setup_ctrl = SetupSettingsController(script_dir)
    self.setup_ctrl.attach(self)          # pass the WPFWindow (self)
    self.setup_ctrl.load_and_apply()      # call after window is fully initialised
"""

import os
import json

# ── WPF imports ────────────────────────────────────────────────────────────────

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

import System.Windows as _SW
import System.Windows.Media as _SWM

# ══════════════════════════════════════════════════════════════════════════════
# CONTROLLER
# ══════════════════════════════════════════════════════════════════════════════

class SetupSettingsController(object):
    """
    Drives the Setup panel (embedded in pyTransmit main window).

    The controller deliberately does NOT call _apply_setup during
    _load_config  — it only populates the UI controls. _apply_setup
    is called explicitly by the host (script.py) after full init, and
    again by the Checked/Unchecked event handlers so changes are
    immediate and visible behind the open Setup panel.
    """

    CONFIG_FILE = 'pytransmit_setup.json'

    DEFAULT_CFG = {
        'show_from':       True,
        'show_client':     True,
        'show_projno':     True,
        'show_projname':   True,
        'show_initials':   True,
        'show_reason':     True,
        'show_method':     True,
        'show_format':     True,
        'show_printsize':  True,
        'recipient_mode':  'dist',
        'group_params':    [],
        'out_schedule':    True,
        'out_excel':       False,
        'out_pdf':         False,
        'out_drafting':    False,
        'out_legend':      False,
        'page_height_mode': 'a4',
        'page_height_mm':   287,
    }

    def __init__(self, script_dir):
        """
        Parameters
        ----------
        script_dir : str
            Absolute path of the folder that contains script.py.
            Config file is written here.
        """
        self._script_dir    = script_dir
        self._config_path   = os.path.join(script_dir, self.CONFIG_FILE)
        self._cfg           = dict(self.DEFAULT_CFG)
        self._host          = None      # the WPFWindow (set by attach())
        self._applying      = False     # re-entrancy guard

    # ── Attach ────────────────────────────────────────────────────────────────

    def attach(self, host):
        """
        Attach to the host WPFWindow.
        Call this once after WPFWindow.__init__ has registered all named elements.
        Does NOT call load_and_apply — call that separately after full init.
        """
        self._host = host
        self._wire_events()

    def load_and_apply(self):
        """Load config from disk, populate UI controls, then apply to main window."""
        self._load_config()
        self._populate_controls()
        self._restore_group_params()
        self.apply()

    def _restore_group_params(self):
        """Restore sheet grouping parameters onto the host window.

        Source priority:
          1. GP: tag in the last issued revision's IssuedTo  — project-specific,
             always correct for this model.
          2. pytransmit_setup.json group_params              — fallback for
             projects issued before GP: tag was introduced, filtered to params
             that exist in the current model to prevent cross-project bleed.
        """
        if self._host is None:
            return
        h = self._host
        try:
            all_params = getattr(h, 'sheet_params', [])
            if not all_params:
                return

            # ── Try GP: tag from last issued revision first ────────────────
            saved = []
            try:
                import clr as _clr
                _clr.AddReference('RevitAPI')
                from Autodesk.Revit.DB import FilteredElementCollector, Revision
                from pyrevit import revit as _rv
                import re as _re2
                _all = sorted(
                    FilteredElementCollector(_rv.doc).OfClass(Revision).ToElements(),
                    key=lambda r: r.SequenceNumber)
                _issued = [r for r in _all if r.Issued]
                if _issued:
                    _ito = (_issued[-1].IssuedTo or '').strip()
                    # GP: tag ends at next ' | ' or end of string
                    _gp_m = _re2.search(r'(?:^| \| )GP:(.*?)(?= \| |$)', _ito)
                    if _gp_m:
                        _raw = _gp_m.group(1).strip()
                        saved = [p for p in _raw.split('~~') if p and p in all_params]
            except Exception:
                pass

            # ── Fallback: pytransmit_setup.json (filtered to this model) ──
            if not saved:
                saved = [p for p in self._cfg.get('group_params', []) if p in all_params]

            if not saved:
                return

            h.selected_params = list(saved)
            # Rebuild the sheet param combos to reflect saved selection
            combos_stack = getattr(h, 'formatting_stack', None)
            if combos_stack is None:
                return
            # Clear existing combos and rebuild from saved list
            import System.Windows.Controls as _SWC2
            import System.Windows as _SW2
            combos_stack.Children.Clear()
            h.sheet_param_combos = []
            h.param_counter = 0
            for i, param_val in enumerate(saved + [None]):
                h.param_counter += 1
                cb2 = _SWC2.ComboBox()
                cb2.Name = "sheet_param_cb_{}".format(h.param_counter)
                available = [p for p in all_params
                             if p not in saved[:i]]
                cb2.ItemsSource = ["(None)"] + available
                cb2.Margin = _SW2.Thickness(0, 0, 0, 4)
                try: cb2.Style = h.FindResource("ModernComboBoxStyle")
                except: pass
                if param_val and param_val in all_params:
                    try: cb2.SelectedItem = param_val
                    except: pass
                else:
                    cb2.SelectedIndex = 0
                cb2.SelectionChanged += h.sheet_param_selection_changed
                combos_stack.Children.Add(cb2)
                h.sheet_param_combos.append(cb2)
            # Keep sheet_param_cb_1 in sync — it was cleared from the stack
            # so update the host reference to point to the new first combo
            if h.sheet_param_combos:
                try: h.sheet_param_cb_1 = h.sheet_param_combos[0]
                except: pass
        except Exception:
            pass

    # ── Persistence ───────────────────────────────────────────────────────────

    def _load_config(self):
        """Read config from JSON; fall back to defaults on any error."""
        try:
            with open(self._config_path, 'r') as f:
                loaded = json.load(f)
            # Merge with defaults so new keys always exist
            cfg = dict(self.DEFAULT_CFG)
            cfg.update(loaded)
            self._cfg = cfg
        except:
            self._cfg = dict(self.DEFAULT_CFG)

    def _read_page_height_mm(self):
        try:
            tb = getattr(self._host, 'setup_page_height_tb', None)
            if tb:
                return int(float(tb.Text or '287'))
        except Exception:
            pass
        return 287

    def save(self):
        """Read current UI state → _cfg → write to disk."""
        if self._host is None:
            return
        h = self._host

        def cb(name, default=True):
            el = getattr(h, name, None)
            if el is None:
                return default
            try:
                v = el.IsChecked
                # IronPython: IsChecked can be a Nullable[bool]
                if v is None:
                    return default
                return bool(v)
            except:
                return default

        self._cfg = {
            'show_from':      cb('setup_from_cb'),
            'show_client':    cb('setup_client_info_cb'),
            'show_projno':    cb('setup_projno_cb'),
            'show_projname':  cb('setup_projname_cb'),
            'show_initials':  cb('setup_initials_cb'),
            'show_reason':    cb('setup_reason_cb'),
            'show_method':    cb('setup_method_cb'),
            'show_format':    cb('setup_format_cb'),
            'show_printsize': cb('setup_printsize_cb'),
            'recipient_mode': 'client' if cb('setup_client_rb', False) else 'dist',
            'group_params':   list(getattr(h, 'selected_params', None) or []),
            'out_schedule':   cb('setup_output_schedule_cb', True),
            'out_excel':      cb('setup_output_excel_cb', False),
            'out_pdf':        cb('setup_output_pdf_cb', False),
            'out_drafting':   cb('setup_output_drafting_cb', False),
            'out_legend':     cb('setup_output_legend_cb', False),
            'page_height_mode': 'none' if cb('setup_height_none_rb', False) else ('custom' if cb('setup_height_custom_rb', False) else 'a4'),
            'page_height_mm':   self._read_page_height_mm(),
        }
        try:
            with open(self._config_path, 'w') as f:
                json.dump(self._cfg, f, indent=2)
        except:
            pass

    # ── UI population ─────────────────────────────────────────────────────────

    def _populate_controls(self):
        """
        Push _cfg values into the Setup panel UI controls.
        Uses a re-entrancy guard so setting IsChecked here does NOT
        trigger the event handlers (which would call save() and apply()
        in a loop during startup).
        """
        if self._host is None:
            return
        h = self._host

        self._applying = True   # <-- guard ON: block event handlers during init
        try:
            def set_cb(name, key):
                el = getattr(h, name, None)
                if el is not None:
                    try:
                        el.IsChecked = bool(self._cfg.get(key, True))
                    except:
                        pass

            set_cb('setup_from_cb',      'show_from')
            set_cb('setup_client_info_cb','show_client')
            set_cb('setup_projno_cb',    'show_projno')
            set_cb('setup_projname_cb',  'show_projname')
            set_cb('setup_initials_cb',  'show_initials')
            set_cb('setup_reason_cb',    'show_reason')
            set_cb('setup_method_cb',    'show_method')
            set_cb('setup_format_cb',    'show_format')
            set_cb('setup_printsize_cb', 'show_printsize')
            set_cb('setup_output_schedule_cb',  'out_schedule')
            set_cb('setup_output_excel_cb',     'out_excel')
            set_cb('setup_output_pdf_cb',       'out_pdf')
            set_cb('setup_output_drafting_cb',  'out_drafting')
            set_cb('setup_output_legend_cb',    'out_legend')

            # Page height
            mode = self._cfg.get('page_height_mode', 'a4')
            for rb_name, rb_val in [('setup_height_none_rb', 'none'), ('setup_height_a4_rb', 'a4'), ('setup_height_custom_rb', 'custom')]:
                el = getattr(h, rb_name, None)
                if el is not None:
                    try: el.IsChecked = (mode == rb_val)
                    except: pass
            try:
                tb = getattr(h, 'setup_page_height_tb', None)
                if tb:
                    tb.Text = str(self._cfg.get('page_height_mm', 287))
            except: pass

            mode = self._cfg.get('recipient_mode', 'dist')
            for rb_name, rb_val in [('setup_dist_rb', 'dist'), ('setup_client_rb', 'client')]:
                el = getattr(h, rb_name, None)
                if el is not None:
                    try:
                        el.IsChecked = (mode == rb_val)
                    except:
                        pass
        finally:
            self._applying = False   # <-- guard OFF

    # ── Apply ─────────────────────────────────────────────────────────────────

    def apply(self):
        """
        Show/hide main window elements based on _cfg.
        Also rebuilds the recipient section (dist rows or client dropdowns)
        and repopulates option combo boxes.
        Safe to call at any time.
        """
        if self._host is None:
            return
        h   = self._host
        V   = _SW.Visibility
        cfg = self._cfg

        def vis(name, show):
            el = getattr(h, name, None)
            if el is not None:
                try:
                    el.Visibility = V.Visible if show else V.Collapsed
                except:
                    pass

        # Optional project fields
        _show_from     = cfg.get('show_from',      True)
        _show_client   = cfg.get('show_client',    True)
        _show_projno   = cfg.get('show_projno',    True)
        _show_projname = cfg.get('show_projname',  True)
        vis('from_row',         _show_from)
        vis('client_row',       _show_client)
        vis('projno_row',       _show_projno)
        vis('projname_row',     _show_projname)
        # Hide the entire card if no info rows are visible
        vis('project_info_card', _show_from or _show_client or _show_projno or _show_projname)
        vis('initials_row',  cfg.get('show_initials',  True))
        vis('reason_row',    cfg.get('show_reason',    True))
        vis('method_row',    cfg.get('show_method',    True))
        vis('format_row',    cfg.get('show_format',    True))
        vis('printsize_row', cfg.get('show_printsize', True))

        # Recipient mode sections
        mode = cfg.get('recipient_mode', 'dist')
        vis('dist_list_section',   mode == 'dist')
        vis('client_list_section', mode == 'client')

        # Rebuild the active section
        try:
            if mode == 'dist':
                h._build_dist_rows()
            else:
                h._build_client_rows()
        except:
            pass

        # Repopulate option combo boxes from Options Manager data
        try:
            h._populate_option_combos()
        except:
            pass

    # ── Event wiring ──────────────────────────────────────────────────────────

    def _wire_events(self):
        """Wire all Setup panel control events."""
        h = self._host
        if h is None:
            return

        for name in ['setup_from_cb', 'setup_client_info_cb', 'setup_projno_cb',
                     'setup_projname_cb', 'setup_initials_cb', 'setup_reason_cb',
                     'setup_method_cb', 'setup_format_cb', 'setup_printsize_cb',
                     'setup_output_schedule_cb', 'setup_output_excel_cb',
                     'setup_output_drafting_cb', 'setup_output_legend_cb']:
            el = getattr(h, name, None)
            if el is not None:
                try:
                    el.Checked   += self._on_field_changed
                    el.Unchecked += self._on_field_changed
                except:
                    pass

        for name in ['setup_dist_rb', 'setup_client_rb']:
            el = getattr(h, name, None)
            if el is not None:
                try:
                    el.Checked += self._on_mode_changed
                except:
                    pass

        for name in ['setup_height_none_rb', 'setup_height_a4_rb', 'setup_height_custom_rb']:
            el = getattr(h, name, None)
            if el is not None:
                try:
                    el.Checked += self._on_field_changed
                except:
                    pass

    def _on_field_changed(self, sender, args):
        """A Project Settings checkbox was toggled."""
        if self._applying:          # guard: ignore events fired during init
            return
        self.save()
        self.apply()

    def _on_mode_changed(self, sender, args):
        """A recipient mode radio button was selected."""
        if self._applying:
            return
        self.save()
        self.apply()

    # ── Public helpers ────────────────────────────────────────────────────────

    @property
    def cfg(self):
        """Read-only access to current config dict."""
        return dict(self._cfg)

    def get_recipient_mode(self):
        """Return 'dist' or 'client'."""
        return self._cfg.get('recipient_mode', 'dist')
