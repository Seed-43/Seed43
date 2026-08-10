# -*- coding: utf-8 -*-
# pyTransmit.py
from pyrevit import revit, forms, script, DB
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, XYZ, Line, TextNote, TextNoteType, CurveElement,
    ViewFamilyType, ViewFamily, ViewDrafting, TextNoteOptions, HorizontalTextAlignment, Color, GraphicsStyle, Category,
    OverrideGraphicSettings, ImageType, ImageTypeOptions, ImageInstance, ImageTypeSource, ImagePlacementOptions, BoxPlacement
)
import math
import re
from itertools import groupby
from pyrevit.forms import WPFWindow  # kept for other uses
import wpf
from System.Windows import Window
from System import Uri, UriKind
from System.Windows.Media.Imaging import BitmapImage
import clr
clr.AddReference("PresentationFramework")
from System.Windows.Controls import ComboBox
from System.Windows import Thickness
import System.Windows
import System.Windows.Media
import os
import sys

from pytransmit_paths import (
    SETTINGS_DIR, LAYOUTS_DIR, STUDIO_LAYOUTS_DIR, LAYOUT_CONFIG, SYNC_FILE,
    USER_DIR, settings_file,
)

_SCRIPT_DIR_MAIN = os.path.dirname(os.path.abspath(__file__))

try:
    from Snippets import _dialogs as sdlg
except Exception:
    sdlg = None

def _alert(message, title='', exitscript=False):
    """Themed popup via the shared Snippets dialog lib, falls back to
    pyRevit's default forms.alert if the shared lib isn't available."""
    if sdlg:
        sdlg.message(message, title=title)
    else:
        forms.alert(message, title=title)
    if exitscript:
        script.exit()

def _confirm(message, title='', no='No'):
    """Themed yes/no popup, returns True on yes."""
    if sdlg:
        return sdlg.confirm(message, title=title, no=no)
    return bool(forms.alert(message, title=title, ok=False, yes=True, no=True))

# ── EXTERNAL URLS ──
# Update these to change where Help and About point
ABOUT_URL = "https://seed43.org/pytransmit/"
SUPPORT_URL = "https://buymeacoffee.com/seed43"

import json as _json

# --- PARAMETERS ---
max_revs = 8
recipients = [
    "Architect/Designer",
    "Owner/Developer",
    "Contractor",
    "Local Authority"
]
table_title = "Distribution List"
matrix_table_title = "Transmittal List"
columns = ["Sent To", "Attention To"]
copies_header = "Number of Copies"
copies_count = max_revs
copies_width_ft = 0.032808  # 10mm in feet
column_gap = 0.0492126  # 15mm in feet
matrix_table_origin = XYZ(0, 0, 0)
first_row_height = 0.0656
other_row_height = 0.0164042  # 5mm in feet
param_row_height = 0.019685   # 6mm in feet for parameter headers and heading rows
distribution_gap = 0.0656168  # 20mm in feet
short_curve_tolerance = 0.002083333  # Revit's ShortCurveTolerance in feet (~0.635 mm)
reason_row_height = 0.019685  # 6mm to match param_row_height
text_lift = 0.00328084  # 1mm in feet for lifting Reason for Issue text and sheet data
text_lift_dist = 0.00328084  # 1mm in feet for lifting Distribution and Matrix header text
key_lift_reason = 0  # 0mm
key_lift_additional = -0.00814044  # -2.48mm for Distribution Table
description_width_ft = 0.328084  # 100mm in feet (updated from 80mm)
attention_to_width_ft = 0.295276  # 90mm in feet (90 / 304.8, updated from 70mm)
logo_space_ft = 0.131234  # 40mm in feet for logo space above Sheet/Description headers

# --- XAML UI CLASS ---
class RevTableWindow(Window):
    def __init__(self):
        
        try:
            xaml_path = os.path.join(_SCRIPT_DIR_MAIN, "pyTransmit.xaml")
            wpf.LoadComponent(self, xaml_path)
        except Exception as e:
            # Build full exception chain including inner exceptions
            msg = str(e)
            try:
                inner = e.InnerException
                depth = 0
                while inner and depth < 5:
                    msg += '\n\nINNER[{}]: {}'.format(depth, str(inner))
                    inner = inner.InnerException
                    depth += 1
            except Exception:
                pass
            _alert("Failed to load pyTransmit.xaml:\n\n{}".format(msg), exitscript=True)

        # -- Apply Seed43 theme (colours + sizing) ---------------------------------
        # Must run AFTER LoadComponent (so injected brushes beat the XAML's own
        # Setters) and BEFORE anything below that builds dynamic UI or calls
        # TryFindResource, which would otherwise return None.
        #
        # The styles themselves (PrimaryButtonStyle, ComboBoxStyle, ...) are
        # declared locally in this window's XAML, named after Seed43Styles.xaml
        # rather than loaded from it. Only the colour/size VALUES they pull via
        # DynamicResource come from seed43_palette.json - what these calls inject.
        try:
            from Snippets.seed43_theme import apply_seed43_palette, apply_seed43_dimensions
            apply_seed43_palette(self, _SCRIPT_DIR_MAIN)
            apply_seed43_dimensions(self, _SCRIPT_DIR_MAIN)
        except Exception:
            pass

        # ── Load icon ─────────────────────────────────────────────────────────
        # Shared loader: it closes the file after reading, so the icon
        # rebuilder can still overwrite icon.png while this window is open,
        # and skips WPF's bitmap cache so a rebuilt icon actually shows.
        try:
            from Snippets._icons import set_header_icon
            set_header_icon(self, _SCRIPT_DIR_MAIN)
        except Exception:
            pass

        # Initialise panel controllers (panels are now in pyTransmit.xaml directly)
        self._init_controllers()
        self.Closing += self._on_closing

        # ── Set close button icons ─────────────────────────────────────────────
        try:
            from Snippets._icons import make_icon as _mi
            _icon_color = self._theme_hex('text_primary', '#F4FAFF')
            _close_icon_names = [
                'setup_close_btn', 'styling_close_btn', 'export_format_close_btn',
                'recipient_close_btn', 'options_close_btn', 'export_settings_close_btn',
                'import_settings_close_btn', 'filenaming_close_btn', 'win_close_btn',
            ]
            for _btn_name in _close_icon_names:
                _btn = self.FindName(_btn_name)
                if _btn:
                    _btn.Content = _mi('close', size=14, color=_icon_color)
            self.options_btn.Content = _mi('menu', size=18, color=_icon_color)
            # GitHub mark on the ☰ menu, built here for the same reason the
            # close icons are: make_icon bakes its colour in at build time.
            from Snippets._icons import make_icon_with_label as _mil
            self.issue_btn.Content = _mil(
                'github', u'Report an issue on GitHub', icon_size=14,
                color=_icon_color)
            # ── Log menu warning icon ─────────────────────────────────────────
            _log_holder = self.FindName('log_icon_holder')
            if _log_holder:
                _log_holder.Child = _mi('warning', size=14, color=_icon_color)
        except Exception:
            pass

        # ── Log state, always starts off ─────────────────────────────────────
        self._log_enabled  = False
        self._log_zip_path = ''
        try:
            import json as _lj
            _lcfg_path = settings_file('pytransmit_setup.json')
            with open(_lcfg_path, 'r') as _lf:
                _lcfg = _lj.load(_lf)
            self._log_zip_path = _lcfg.get('log_zip_path', '')
        except Exception:
            pass

        if not hasattr(self, 'execute_btn'):
            _alert("Button 'execute_btn' not found in XAML.", exitscript=True)
        
        try:
            self.execute_btn.Click += self.execute_btn_click
        except Exception as e:
            _alert("Failed to bind button Click events: {}".format(str(e)), exitscript=True)
        
        self.doc = revit.doc
        all_revs = list(revit.query.get_elements_by_class(DB.Revision, doc=self.doc))
        self.non_issued_revs = [rev for rev in all_revs if not rev.Issued]
        self.issued_revs = [rev for rev in all_revs if rev.Issued]
        self.issued_revs = sorted(self.issued_revs, key=lambda r: r.SequenceNumber)
        self._rev_numbering_type = ''
        self._init_rev_type_selector()

        try:
            self.revision_cb.ItemsSource = ["{} - {}".format(rev.SequenceNumber, rev.Description) for rev in self.non_issued_revs]
        except Exception as e:
            _alert("Failed to populate revision ComboBox: {}".format(str(e)), exitscript=True)
        
        try:
            self.reason_cb.SelectedIndex = 0
        except:
            pass
        
        self.sheet_param_combos = [self.sheet_param_cb_1]
        self.selected_params = []
        self.param_counter = 1
        self.group_label_on = self._load_sync().get('group_label_on', True)
        self._setup_group_label_toggle()
        self.PreviewMouseDown += self._options_popup_outside_click
        
        self.sheet_params = self.get_sheet_parameters()
        
        if not self.sheet_params:
            _alert("No suitable sheet parameters found.", exitscript=True)
        
        try:
            self.sheet_param_cb_1.ItemsSource = ["(None)"] + self.sheet_params
            self.sheet_param_cb_1.SelectionChanged += self.sheet_param_selection_changed
        except Exception as e:
            _alert("Failed to populate or bind sheet parameter ComboBox: {}".format(str(e)), exitscript=True)
        
        # (Excel export path is now configured in Setup Settings)

        # Load Export/Import config panels
        try:
            if self.export_ctrl:
                self.export_ctrl.load_config()
        except:
            pass
        try:
            if self.import_ctrl:
                self.import_ctrl.load_config()
        except:
            pass

        # Load Setup config, apply to main window, then run auto-import
        try:
            if self.setup_ctrl:
                self.setup_ctrl.load_and_apply()
        except:
            pass
        try:
            if self.import_ctrl:
                self.import_ctrl.run_auto_import()
        except:
            pass

        # Populate layout template dropdowns from Layout/Layouts/ folder
        try:
            self._populate_layout_combos()
        except:
            pass

        # Wire green scrollbars to ContentRendered (visual tree exists after render)
        try:
            self.ContentRendered += self._on_content_rendered
        except:
            pass

        # Check for settings mismatch with last issued revision
        try:
            self._check_settings_mismatch()
        except:
            pass

        # Pre-fill project info textboxes from Revit
        try:
            self._prefill_project_info()
        except:
            pass
        
    def _theme_hex(self, semantic_key, fallback_hex):
        """Current palette colour for a semantic key (e.g. 'text_primary'),
        read straight from the palette JSON, for make_icon() calls (its
        colour param wants a hex string, not a Brush, and is baked in at
        build time - see seed43-pyrevit-ui gotchas #3). Deliberately not
        based on self.TryFindResource(...) - going straight to disk avoids
        any question of whether the WPF resource dictionary has actually
        resolved by the time this runs, which was producing icons baked
        with the wrong (dark-profile) colour when the active profile is
        light."""
        try:
            from Snippets.seed43_theme import get_color
            return get_color(_SCRIPT_DIR_MAIN, semantic_key, fallback=fallback_hex)
        except Exception:
            return fallback_hex

    def _init_rev_type_selector(self):
        """
        Detect distinct revision numbering types from the issued revision list.
        If only one type exists, auto-select it and hide the combo.
        If multiple types exist, populate the combo and show it.
        """
        import System.Windows as _SW

        def _num_type(rev):
            try:
                _se = self.doc.GetElement(rev.RevisionNumberingSequenceId)
                return str(_se.NumberType).split('.')[-1] if _se else 'None'
            except Exception:
                return 'Unknown'

        # Collect distinct types preserving order of first appearance
        _seen = []
        for _r in self.issued_revs:
            _t = _num_type(_r)
            if _t not in _seen:
                _seen.append(_t)

        _cb  = getattr(self, 'rev_type_cb',  None)
        _row = getattr(self, 'rev_type_row', None)

        if len(_seen) <= 1:
            # Auto-select the only type (or blank if no revisions)
            self._rev_numbering_type = _seen[0] if _seen else ''
            if _row:
                _row.Visibility = _SW.Visibility.Collapsed
        else:
            # Multiple types, show the combo
            self._rev_numbering_type = _seen[0]
            if _cb:
                _cb.ItemsSource    = _seen
                _cb.SelectedIndex  = 0
                _cb.SelectionChanged += self._on_rev_type_changed
            if _row:
                _row.Visibility = _SW.Visibility.Visible

    def _on_rev_type_changed(self, sender, args):
        """Update _rev_numbering_type when the combo selection changes."""
        if sender.SelectedItem is not None:
            self._rev_numbering_type = str(sender.SelectedItem)

    def _setup_group_label_toggle(self):
        """Build the on/off toggle switch for group label display."""
        import System.Windows.Media as _SWM
        import System
        _on_brush  = self.TryFindResource('BrushPrimaryGreen') or _SWM.SolidColorBrush(_SWM.Color.FromRgb(0x20, 0x8A, 0x3C))
        _off_brush = self.TryFindResource('BrushToggleOffBg')  or _SWM.SolidColorBrush(_SWM.Color.FromRgb(0xA0, 0xAA, 0xBB))
        _knob_brush = self.TryFindResource('BrushToggleKnob')  or _SWM.Brushes.White
        _track_w    = self.TryFindResource('WidthToggle')      or 40.0
        _knob_size  = self.TryFindResource('SizeToggleKnob')   or 16.0
        _knob_margin = self.TryFindResource('MarginToggleKnob') or 2.0
        _knob_radius = self.TryFindResource('CornerRadiusToggleKnob') or System.Windows.CornerRadius(_knob_size / 2.0)
        try:
            sw = self.group_label_toggle
            sw.Background = _on_brush if self.group_label_on else _off_brush
            knob                     = System.Windows.Controls.Border()
            knob.Width               = _knob_size
            knob.Height              = _knob_size
            knob.CornerRadius        = _knob_radius
            knob.Background          = _knob_brush
            knob.HorizontalAlignment = System.Windows.HorizontalAlignment.Left
            _on_offset = _track_w - _knob_size - _knob_margin
            knob.Margin = (
                System.Windows.Thickness(_on_offset, _knob_margin, 0, _knob_margin) if self.group_label_on
                else System.Windows.Thickness(_knob_margin, _knob_margin, 0, _knob_margin))
            sw.Child = knob
            self._group_label_knob = knob
            sw.MouseLeftButtonUp += self._on_group_label_toggle
            # Set initial text colour
            self.group_label_toggle_tb.Foreground = (
                self.TryFindResource('BrushTextPrimary')
                or _SWM.SolidColorBrush(_SWM.Color.FromRgb(0xF4, 0xFA, 0xFF)))
        except Exception:
            pass

    def _on_group_label_toggle(self, sender, args):
        """Toggle group label on/off, updating knob position and colour directly."""
        import System.Windows.Media as _SWM
        import System
        # Fresh lookups here (not the values captured in _setup_group_label_toggle) -
        # this handler can fire long after window load, so re-read live (gotcha #3).
        _on_brush  = self.TryFindResource('BrushPrimaryGreen') or _SWM.SolidColorBrush(_SWM.Color.FromRgb(0x20, 0x8A, 0x3C))
        _off_brush = self.TryFindResource('BrushToggleOffBg')  or _SWM.SolidColorBrush(_SWM.Color.FromRgb(0xA0, 0xAA, 0xBB))
        _on_text_brush = self.TryFindResource('BrushTextPrimary') or _SWM.SolidColorBrush(_SWM.Color.FromRgb(0xF4, 0xFA, 0xFF))
        _track_w    = self.TryFindResource('WidthToggle')    or 40.0
        _knob_size  = self.TryFindResource('SizeToggleKnob') or 16.0
        _knob_margin = self.TryFindResource('MarginToggleKnob') or 2.0
        self.group_label_on = not self.group_label_on
        try:
            # Set knob position directly, no animation (reliable in IronPython 2)
            _on_offset = _track_w - _knob_size - _knob_margin
            self._group_label_knob.Margin = (
                System.Windows.Thickness(_on_offset, _knob_margin, 0, _knob_margin) if self.group_label_on
                else System.Windows.Thickness(_knob_margin, _knob_margin, 0, _knob_margin))
            self.group_label_toggle.Background = _on_brush if self.group_label_on else _off_brush
            self.group_label_toggle_tb.Text = "Text On" if self.group_label_on else "Text Off"
            self.group_label_toggle_tb.Foreground = (
                _on_text_brush if self.group_label_on else _off_brush)
        except Exception:
            pass
        self._save_sync()

    def get_sheet_parameters(self):
        sheets = FilteredElementCollector(self.doc).OfCategory(BuiltInCategory.OST_Sheets).WhereElementIsNotElementType().ToElements()
        if not sheets:
            return []
        param_names = set()
        sample_sheet = next(iter(sheets), None)
        if not sample_sheet:
            return []
        built_in_params = [
            DB.BuiltInParameter.SHEET_NUMBER,
            DB.BuiltInParameter.SHEET_NAME
        ]
        # String params (Sheet Number/Name) plus ElementId ones - Revit's newer
        # "Dropdown List" type, e.g. the 2026 Sheet Collection parameter, is
        # backed by ElementId, not String. Integer/Double stay excluded as
        # numeric fields like Scale aren't meaningful text to group by.
        # Grouping code already reads values via AsString()-or-AsValueString(),
        # which handles ElementId, so these just become selectable here too.
        _USABLE_STORAGE_TYPES = (DB.StorageType.String, DB.StorageType.ElementId)

        def _has_usable_value(param):
            if param.StorageType == DB.StorageType.String:
                return True
            try:
                return bool(param.AsValueString())
            except Exception:
                return False

        for bip in built_in_params:
            param = sample_sheet.get_Parameter(bip)
            if param and param.StorageType == DB.StorageType.String:
                param_names.add(param.Definition.Name)
        for param in sample_sheet.GetOrderedParameters():
            if (param.Definition
                    and param.StorageType in _USABLE_STORAGE_TYPES
                    and _has_usable_value(param)):
                param_names.add(param.Definition.Name)
        return sorted(list(param_names))
    
    def sheet_param_selection_changed(self, sender, args):
        
        # Check if "(None)" was selected
        if sender.SelectedItem == "(None)":
            # Find the index of the sender combo box
            sender_index = self.sheet_param_combos.index(sender)
            
            # Remove all combo boxes after this one
            from System.Windows import Application
            combos_to_remove = self.sheet_param_combos[sender_index + 1:]
            for combo in combos_to_remove:
                Application.Current.Dispatcher.Invoke(lambda c=combo: self.formatting_stack.Children.Remove(c))
            
            # Update the list of combo boxes
            self.sheet_param_combos = self.sheet_param_combos[:sender_index + 1]
            
            # Reset the sender's selection to blank (first item)
            sender.SelectedIndex = 0
            
            # Update selected params
            self.selected_params = [cb.SelectedItem for cb in self.sheet_param_combos if cb.SelectedItem and cb.SelectedItem != "(None)"]
            return
        
        self.selected_params = [cb.SelectedItem for cb in self.sheet_param_combos if cb.SelectedItem and cb.SelectedItem != "(None)"]
        if sender.SelectedItem and sender.SelectedItem != "(None)" and sender == self.sheet_param_combos[-1]:
            self.param_counter += 1
            new_combo = ComboBox()
            new_combo.Name = "sheet_param_cb_{}".format(self.param_counter)
            try:
                new_combo.Style = self.FindResource("ComboBoxStyle")
            except Exception as e:
                print("Failed to apply ComboBoxStyle: {}".format(str(e)))
            available_params = [p for p in self.sheet_params if p not in self.selected_params]
            new_combo.ItemsSource = ["(None)"] + available_params
            new_combo.SelectionChanged += self.sheet_param_selection_changed
            from System.Windows import Application
            Application.Current.Dispatcher.Invoke(lambda: self.formatting_stack.Children.Add(new_combo))
            self.sheet_param_combos.append(new_combo)
    
    def safe(self, val):
        return val if val is not None else ""

    def safeint(self, val):
        try:
            v = str(val).strip()
            if not v:
                return ''
            return str(int(v))
        except:
            return ''

    def execute_btn_click(self, sender, args):
        try:
            selected_rev = None
            reason_code = ""

            # Update revision if selected
            if self.non_issued_revs:
                selected_rev_index = self.revision_cb.SelectedIndex
                if selected_rev_index != -1:
                    selected_rev = self.non_issued_revs[selected_rev_index]
                    # Build issued-to string from current recipient mode + setup field toggles
                    data_str, initials_str = self._build_issued_to_string()
                    try:
                        with revit.Transaction("Set Issued To Data and Mark Issued"):
                            selected_rev.IssuedTo = data_str
                            if initials_str:
                                selected_rev.IssuedBy = initials_str
                            selected_rev.Issued = True
                    except Exception as e:
                        _alert("Failed to update revision data: {}".format(str(e)), exitscript=True)
                        return

            # Determine what to run based on Revit Export Type selection in Setup
            output_type = 'schedule'  # safe default
            try:
                if getattr(self, 'setup_output_drafting_rb', None) and self.setup_output_drafting_rb.IsChecked:
                    output_type = 'drafting'
                elif getattr(self, 'setup_output_legend_rb', None) and self.setup_output_legend_rb.IsChecked:
                    output_type = 'legend'
                elif getattr(self, 'setup_output_excel_rb', None) and self.setup_output_excel_rb.IsChecked:
                    output_type = 'excel'
                elif getattr(self, 'setup_output_schedule_rb', None) and self.setup_output_schedule_rb.IsChecked:
                    output_type = 'schedule'
            except:
                pass

            # All output types now dispatch through run_revit_export
            self.run_revit_export()

            self.Close()

        except Exception as e:
            try:
                self.Close()
            except Exception:
                pass
            _alert("Error in execute_btn_click: {}".format(str(e)), exitscript=True)

    def _build_issued_to_string(self):
        """
        Build the IssuedTo string written into the Revit revision record.

        Always records every enabled field and the full recipient list, no
        user-configurable toggles.  The schedule generator reads this string
        back to populate the transmittal, so it must be complete and consistent.

        Format:
          R:<code> M:<code> F:<value> S:<value> I:<initials> | <recipients>

        Recipients (Distribution List mode):
          A.[Attention To]<copies>  O.[...]<copies>  ...
          (first letter of the role label, attention to in brackets, copies integer)

        Recipients (Client List mode):
          [Company, Attention To]<copies>  ...

        Example:
          R:C M:E F:PDF S:A3 I:JD | A.[Jane Smith]3 O.[Bob Jones]1
        """
        cfg  = self.setup_ctrl.cfg if self.setup_ctrl else {}
        mode = cfg.get('recipient_mode', 'dist')

        meta_parts = []

        # Reason for Issue
        try:
            if cfg.get('show_reason'):
                cb = getattr(self, 'reason_cb', None)
                if cb and cb.SelectedIndex > 0:  # 0 = (none)
                    rows = list(self.opt_ctrl.reason_data)
                    idx  = cb.SelectedIndex - 1   # offset for (none)
                    if idx < len(rows):
                        code = getattr(rows[idx], 'Code', '') or ''
                        if code:
                            meta_parts.append('R:{}'.format(code))
        except:
            pass

        # Method of Issue
        try:
            if cfg.get('show_method'):
                cb = getattr(self, 'method_cb', None)
                if cb and cb.SelectedIndex > 0:  # 0 = (none)
                    rows = list(self.opt_ctrl.method_data)
                    idx  = cb.SelectedIndex - 1
                    if idx < len(rows):
                        code = getattr(rows[idx], 'Code', '') or ''
                        if code:
                            meta_parts.append('M:{}'.format(code))
        except:
            pass

        # Document Format
        try:
            if cfg.get('show_format'):
                cb = getattr(self, 'format_cb', None)
                if cb and cb.SelectedIndex > 0:  # 0 = (none)
                    rows = list(self.opt_ctrl.format_data)
                    idx  = cb.SelectedIndex - 1
                    if idx < len(rows):
                        val = getattr(rows[idx], 'Value', '') or ''
                        if val:
                            meta_parts.append('F:{}'.format(val))
        except:
            pass

        # Print Size
        try:
            if cfg.get('show_printsize'):
                cb = getattr(self, 'printsize_cb', None)
                if cb and cb.SelectedIndex > 0:  # 0 = (none)
                    rows = list(self.opt_ctrl.printsize_data)
                    idx  = cb.SelectedIndex - 1
                    if idx < len(rows):
                        val = getattr(rows[idx], 'Value', '') or ''
                        if val:
                            meta_parts.append('S:{}'.format(val))
        except:
            pass

        # Issued By, written to rev.IssuedBy, NOT included in IssuedTo string
        _initials_val = ''
        try:
            if cfg.get('show_initials', True):
                tb = getattr(self, 'initials_tb', None)
                if tb and tb.Text:
                    _initials_val = tb.Text.strip()
        except:
            pass

        # Recipients, always record the full list from whichever mode is active
        recipient_parts = []
        try:
            if mode == 'dist':
                for _ri, row in enumerate(self._dist_rows):
                    label  = row.get('label', '')
                    attn   = row['attn_tb'].Text   if row.get('attn_tb')   else ''
                    copies = row['copies_tb'].Text if row.get('copies_tb') else ''
                    code   = '{}{}.'.format(_ri + 1, label[:1].upper() if label else '?')
                    recipient_parts.append('{}[{}]{}'.format(
                        code, self.safe(attn), self.safeint(copies)))
            else:  # client mode, new structure: {company_cb, contact_cb, copies_tb}
                for row in self._client_rows:
                    comp_cb  = row.get('company_cb')
                    cont_cb  = row.get('contact_cb')
                    copies_tb = row.get('copies_tb')
                    if comp_cb is None or comp_cb.SelectedIndex <= 0:
                        continue
                    company = str(comp_cb.SelectedItem or '')
                    attn    = (str(cont_cb.SelectedItem or '') if cont_cb
                               and cont_cb.SelectedIndex > 0 else '')
                    copies  = copies_tb.Text if copies_tb else ''
                    label   = u'{} \u2014 {}'.format(company, attn) if attn else company
                    recipient_parts.append('[{}]{}'.format(
                        self.safe(label), self.safeint(copies)))
        except:
            pass

        # Combine: meta block | recipients block | VIS tag | EX tag
        parts = []
        if meta_parts:
            parts.append(' '.join(meta_parts))
        if recipient_parts:
            _rec_prefix = 'DL' if mode == 'dist' else 'CL'
            parts.append('{}: {}'.format(_rec_prefix, ' '.join(recipient_parts)))

        # Rev Numbering Type tag, records which type was selected for this issue
        _rnt = getattr(self, '_rev_numbering_type', '')
        if _rnt:
            parts.append('RT:{}'.format(_rnt))

        # VIS tag, snapshot which info rows are visible.
        # Always written (even if empty) so the mismatch checker can distinguish
        # "all fields off" from "old revision with no VIS tag".
        _vis_parts = []
        if cfg.get('show_from',     True): _vis_parts.append('FR')
        if cfg.get('show_client',   True): _vis_parts.append('CL')
        if cfg.get('show_projno',   True): _vis_parts.append('PN')
        if cfg.get('show_projname', True): _vis_parts.append('PJ')
        parts.append('VIS:{}'.format(','.join(_vis_parts)))

        # EX tag, snapshot which export formats are enabled
        _ex_parts = []
        if cfg.get('out_schedule',  True):  _ex_parts.append('RS')
        if cfg.get('out_drafting',  False): _ex_parts.append('RD')
        if cfg.get('out_legend',    False): _ex_parts.append('RL')
        if cfg.get('out_excel',     False): _ex_parts.append('Excl')
        if cfg.get('out_pdf',       False): _ex_parts.append('PDF')
        if _ex_parts:
            parts.append('EX:{}'.format(','.join(_ex_parts)))

        # RPG tag, snapshot page break setting
        _phm = cfg.get('page_height_mode', 'a4')
        _phv = cfg.get('page_height_mm', 287)
        if _phm == 'none':
            parts.append('RPG:0')
        elif _phm == 'custom':
            parts.append('RPG:{}'.format(int(_phv)))
        else:
            parts.append('RPG:1')  # A4

        # GP tag, snapshot active sheet grouping parameters
        _gp_list = getattr(self, 'selected_params', None) or []
        if _gp_list:
            # Use ~~ as separator, safe, won't appear in Revit parameter names
            parts.append(u'GP:{}'.format(u'~~'.join(_gp_list)))

        return ' | '.join(parts), _initials_val
    
    def open_recipient_manager(self, sender, args):
        """Open Recipient panel (called from XAML if needed)."""
        self._show_panel("recipient")

    def open_options_panel(self, sender, args):
        """Open Options panel (called from XAML if needed)."""
        self._show_panel("options")

    def open_settings_manager(self, sender, args):
        self.open_options_panel(sender, args)

    # ── Data models (mirrors standalone managers) ─────────────────────

    # ═══════════════════════════════════════════════════════════════════════
    # PANEL CONTROLLERS, panels are now inline in pyTransmit.xaml
    # ═══════════════════════════════════════════════════════════════════════

    def _init_controllers(self):
        """
        Import controller classes from Settings/ and attach them to the window.
        All named XAML elements are already on self via WPFWindow.__init__.
        """
        script_dir   = os.path.dirname(os.path.abspath(__file__))
        settings_dir = os.path.join(script_dir, 'Settings')

        if settings_dir not in sys.path:
            sys.path.insert(0, settings_dir)

        # ── Recipient Manager ──────────────────────────────────────────────
        try:
            from RecipientSettings import RecipientSettingsController
            self.rec_ctrl = RecipientSettingsController()
            self.rec_ctrl.attach(self)
        except Exception as ex:
            _alert("Failed to init RecipientManager:\n{}".format(str(ex)))
            self.rec_ctrl = None

        # ── Options Manager ────────────────────────────────────────────────
        try:
            from OptionsSettings import OptionsSettingsController
            self.opt_ctrl = OptionsSettingsController()
            self.opt_ctrl.attach(self)
        except Exception as ex:
            _alert("Failed to init OptionsManager:\n{}".format(str(ex)))
            self.opt_ctrl = None

        # ── Setup Settings ─────────────────────────────────────────────────
        try:
            from SetupSettings import SetupSettingsController
            self.setup_ctrl = SetupSettingsController(script_dir)
            self.setup_ctrl.attach(self)
            # load_and_apply() is called later, after full window init
        except Exception as ex:
            _alert("Failed to init SetupSettings:\n{}".format(str(ex)))
            self.setup_ctrl = None

        # ── Export Settings ────────────────────────────────────────────────
        try:
            from ExportSettings import ExportSettingsController
            self.export_ctrl = ExportSettingsController(script_dir)
            self.export_ctrl.attach(self)
        except Exception as ex:
            _alert("Failed to init ExportSettings:\n{}".format(str(ex)))
            self.export_ctrl = None

        # ── Import Settings ────────────────────────────────────────────────
        try:
            from ImportSettings import ImportSettingsController
            self.import_ctrl = ImportSettingsController(script_dir)
            self.import_ctrl.attach(self)
        except Exception as ex:
            _alert("Failed to init ImportSettings:\n{}".format(str(ex)))
            self.import_ctrl = None

        # ── Branding & Styling ─────────────────────────────────────────────
        # Initialised first among visual controllers so logo is synced before
        # anything else runs.  auto_sync_logo() silently copies from source if
        # the network path is reachable; does nothing if it is not.
        try:
            from BrandingSettings import BrandingSettingsController
            self.brand_ctrl = BrandingSettingsController(script_dir)
            self.brand_ctrl.attach(self)
            self.brand_ctrl.auto_sync_logo()
        except Exception as ex:
            _alert("Failed to init BrandingSettings:\n{}".format(str(ex)))
            self.brand_ctrl = None

        try:
            from FileNamingSettings import FileNamingSettingsController
            self.filenaming_ctrl = FileNamingSettingsController(script_dir)
            self.filenaming_ctrl.attach(self)
            self.filenaming_ctrl.load_config()
        except Exception as ex:
            _alert("Failed to init FileNamingSettings:\n{}".format(str(ex)))
            self.filenaming_ctrl = None

    # ── Panel visibility ──────────────────────────────────────────────────

    def _show_panel(self, panel_name):
        import System.Windows as _SW
        V = _SW.Visibility

        def hide(name):
            el = getattr(self, name, None)
            if el is not None:
                try: el.Visibility = V.Collapsed
                except: pass

        def show(name):
            el = getattr(self, name, None)
            if el is not None:
                try: el.Visibility = V.Visible
                except: pass

        # Hide everything first
        for n in ['SetupPanel', 'RecipientPanel', 'OptionsPanel',
                  'ExportSettingsPanel', 'ImportSettingsPanel', 'StylingPanel',
                  'FileNamingPanel', 'ExportFormatPanel',
                  'main_content',
                  'header_normal_btns', 'setup_close_btn', 'styling_close_btn',
                  'export_format_close_btn',
                  'recipient_header_btns', 'options_header_btns',
                  'export_settings_header_btns', 'import_settings_header_btns',
                  'filenaming_header_btns',
                  'setup_header_lbl', 'recipient_header_lbl', 'options_header_lbl',
                  'export_settings_header_lbl', 'import_settings_header_lbl',
                  'styling_header_lbl', 'filenaming_header_lbl', 'export_format_header_lbl']:
            hide(n)

        try:
            self.win_close_btn.Visibility = (
                V.Visible if panel_name == "main" else V.Collapsed)
        except Exception:
            pass

        if panel_name == "main":
            show('header_normal_btns')
            show('main_content')
        elif panel_name == "setup":
            show('SetupPanel')
            show('setup_header_lbl')
            show('setup_close_btn')
        elif panel_name == "recipient":
            show('RecipientPanel')
            show('recipient_header_lbl')
            show('recipient_header_btns')
            if self.rec_ctrl:
                self.rec_ctrl._take_snapshot()
        elif panel_name == "options":
            show('OptionsPanel')
            show('options_header_lbl')
            show('options_header_btns')
            if self.opt_ctrl:
                self.opt_ctrl._take_snapshot()
        elif panel_name == "export_settings":
            show('ExportSettingsPanel')
            show('export_settings_header_lbl')
            show('export_settings_header_btns')
        elif panel_name == "import_settings":
            show('ImportSettingsPanel')
            show('import_settings_header_lbl')
            show('import_settings_header_btns')
        elif panel_name == "styling":
            show('StylingPanel')
            show('styling_header_lbl')
            show('styling_close_btn')
        elif panel_name == "file_naming":
            show('FileNamingPanel')
            show('filenaming_header_lbl')
            show('filenaming_header_btns')
        elif panel_name == "export_format":
            show('ExportFormatPanel')
            show('export_format_header_lbl')
            show('export_format_close_btn')

        # One shared footer, its text swapped to match the panel now showing -
        # this is what the footer is for, instead of duplicating a note inside
        # every panel's own content.
        _FOOTER_TEXT = {
            'main':            u"Visit Setup and Export Format (\u2630 Menu) to configure how pyTransmit works for your projects.",
            'setup':           u"This panel saves automatically when you close it.",
            'export_format':   u"This panel saves automatically when you close it.",
            'export_settings': u"This panel saves automatically when you close it.",
            'import_settings': u"This panel saves automatically when you close it.",
            'recipient':       u"You'll be asked to save or discard changes when you close this panel.",
            'options':         u"You'll be asked to save or discard changes when you close this panel.",
            'file_naming':     u"This panel saves automatically when you close it. Thanks to Ryan McCullough for his printFromIndex tool, which this naming system is built on.",
        }
        try:
            self.footer_text.Text = _FOOTER_TEXT.get(panel_name, _FOOTER_TEXT['main'])
        except Exception:
            pass

    # ── Back / close handlers ─────────────────────────────────────────────

    def _show_save_dialog(self, panel_label):
        from Dialogs import Dialogs as _D
        return _D.save_discard(panel_label)

    def recipient_back_click(self, sender, args):
        try:
            if self.rec_ctrl:
                try:
                    save = self._show_save_dialog("Recipients")
                except Exception:
                    save = _confirm("Save changes to Recipients?", title="Recipients")
                if save:
                    self.rec_ctrl.save()
                    self._auto_export_if_enabled()
                else:
                    self.rec_ctrl.discard()
                self.rec_ctrl.clear_selections()
        except Exception:
            pass
        self._show_panel("main")

    def options_back_click(self, sender, args):
        try:
            if self.opt_ctrl:
                try:
                    save = self._show_save_dialog("Options")
                except Exception:
                    save = _confirm("Save changes to Options?", title="Options")
                if save:
                    self.opt_ctrl.save_all()
                    self._auto_export_if_enabled()
                else:
                    self.opt_ctrl.discard()
                self.opt_ctrl.clear_selections()
        except Exception:
            pass
        self._show_panel("main")

    # ── Header export buttons, delegate to controllers ───────────────────

    def recipient_export_click(self, sender, args):
        if self.rec_ctrl:
            self.rec_ctrl.export_data(sender, args)

    def options_export_click(self, sender, args):
        if self.opt_ctrl:
            self.opt_ctrl.export_data(sender, args)

    def menu_export_format_click(self, sender, args):
        """☰ → Export Format: open the export format configuration panel."""
        self.OptionsPopup.IsOpen = False
        self.options_btn.IsChecked = False
        self._show_panel("export_format")

    def export_format_back_click(self, sender, args):
        """Export Format X, save config and return to main."""
        if self.setup_ctrl:
            self.setup_ctrl.save()
            self.setup_ctrl.apply()
        self._save_layout_assignments()
        self._save_sync()
        self._show_panel("main")

    def menu_setup_click(self, sender, args):
        """☰ → Setup: open the Setup configuration panel."""
        self.OptionsPopup.IsOpen = False
        self.options_btn.IsChecked = False
        self._show_panel("setup")

    def setup_back_click(self, sender, args):
        """Setup X, save config, apply, return to main."""
        if self.setup_ctrl:
            self.setup_ctrl.save()
            self.setup_ctrl.apply()
        self._show_panel("main")

    def win_close_clicked(self, sender, args):
        self.Close()

    # ── Styling / Branding panel ──────────────────────────────────────────────
    # All logic lives in Settings/BrandingSettings.py (BrandingSettingsController)

    def menu_styling_click(self, sender, args):
        """Legacy, Branding panel removed; redirects to Document Layout."""
        self.menu_layout_click(sender, args)

    def menu_layout_click(self, sender, args):
        """☰ → Document Layout: open the Layout Builder in a separate window."""
        self.OptionsPopup.IsOpen = False
        try: self.options_btn.IsChecked = False
        except: pass
        try:
            _layout_dir = os.path.join(
                os.path.dirname(os.path.abspath(__file__)), 'Layout')
            if not os.path.isdir(_layout_dir):
                os.makedirs(_layout_dir)
            if _layout_dir not in sys.path:
                sys.path.insert(0, _layout_dir)
            from LayoutSettings import LayoutSettingsWindow
            win = LayoutSettingsWindow(_layout_dir)
            win.ShowDialog()
        except Exception as e:
            _alert(
                'Could not open Layout Builder:\n{}'.format(str(e)),
                title='Document Layout')

    def styling_back_click(self, sender, args):
        """Styling X, legacy handler kept for safety."""
        self._show_panel("main")

    # ── Layout Studio (Excel-style layout builder) ────────────────────────────
    # Separate from Document Layout above: its own module in Studio/, its own
    # JSON format under Studio/studio_layouts/ - does not read or write
    # anything the Layout Builder (Layout/LayoutSettings.py) uses.

    def menu_layout_studio_click(self, sender, args):
        """☰ → Layout Studio: open the Excel-style layout builder in a separate window."""
        self.OptionsPopup.IsOpen = False
        try: self.options_btn.IsChecked = False
        except: pass
        try:
            _studio_dir = os.path.join(
                os.path.dirname(os.path.abspath(__file__)), 'Studio')
            if not os.path.isdir(_studio_dir):
                os.makedirs(_studio_dir)
            if _studio_dir not in sys.path:
                sys.path.insert(0, _studio_dir)
            from StudioSettings import StudioSettingsWindow
            # Reason / Method / Document Format / Page Size live in this
            # window's combo boxes and are only written onto a revision once a
            # transmittal has been published, so Studio cannot read them from
            # the model. Handed over here, the canvas previews the issue the
            # user is actually about to send instead of ghost placeholders.
            win = StudioSettingsWindow(_studio_dir,
                                       meta_rows=self._current_meta_rows())
            win.ShowDialog()
        except Exception as e:
            _alert(
                'Could not open Layout Studio:\n{}'.format(str(e)),
                title='Layout Studio')

    def _on_closing(self, sender, args):
        """Save all controller state when the window is closed."""
        for ctrl in [
            getattr(self, 'filenaming_ctrl', None),
            getattr(self, 'opt_ctrl',        None),
            getattr(self, 'recipient_ctrl',  None),
        ]:
            try:
                if ctrl:
                    ctrl.save_config()
            except Exception:
                pass

    def filenaming_back_click(self, sender, args):
        """File Naming X, auto-save and return to main."""
        try:
            if self.filenaming_ctrl:
                self.filenaming_ctrl.save_and_back()
        except Exception:
            pass
        self._show_panel("main")

    def filenaming_start_drag(self, sender, args):
        if self.filenaming_ctrl:
            self.filenaming_ctrl._on_start_drag(sender, args)

    def filenaming_preview_drag(self, sender, args):
        if self.filenaming_ctrl:
            self.filenaming_ctrl._on_preview_drag(sender, args)

    def filenaming_stop_drag(self, sender, args):
        if self.filenaming_ctrl:
            self.filenaming_ctrl._on_stop_drag(sender, args)

    def filenaming_path_start_drag(self, sender, args):
        if self.filenaming_ctrl:
            self.filenaming_ctrl._on_path_start_drag(sender, args)

    def filenaming_path_preview_drag(self, sender, args):
        if self.filenaming_ctrl:
            self.filenaming_ctrl._on_path_preview_drag(sender, args)

    def filenaming_path_stop_drag(self, sender, args):
        if self.filenaming_ctrl:
            self.filenaming_ctrl._on_path_stop_drag(sender, args)

    def setup_mode_changed(self, sender, args):
        """Radio button toggled, delegate to SetupSettingsController."""
        if self.setup_ctrl:
            self.setup_ctrl._on_mode_changed(sender, args)

    def setup_fields_changed(self, sender, args):
        """Checkbox/RadioButton toggled, delegate to SetupSettingsController, then handle local UI."""
        if self.setup_ctrl:
            self.setup_ctrl._on_field_changed(sender, args)
        # Show/hide custom height TextBox based on page height radio selection
        try:
            custom_rb  = getattr(self, 'setup_height_custom_rb', None)
            custom_row = getattr(self, 'custom_height_row', None)
            if custom_rb is not None and custom_row is not None:
                import System.Windows as _SW
                custom_row.Visibility = (
                    _SW.Visibility.Visible if custom_rb.IsChecked
                    else _SW.Visibility.Collapsed
                )
        except:
            pass

    def menu_recipient_click(self, sender, args):
        """☰ → Recipient Manager."""
        self.OptionsPopup.IsOpen = False
        self.options_btn.IsChecked = False
        self._show_panel("recipient")

    def menu_options_click(self, sender, args):
        """☰ → Options Manager."""
        self.OptionsPopup.IsOpen = False
        self.options_btn.IsChecked = False
        self._show_panel("options")

    def _open_url(self, url, title=''):
        """Open a URL in the default browser. The launch itself lives in
        Snippets._support.open_url; this only supplies pyTransmit's error
        reporting."""
        from Snippets._support import open_url
        open_url(url, window=self,
                 on_error=lambda msg: _alert(msg, title=title))

    def options_btn_click(self, sender, args):
        """Toggle the options menu open/closed. With the Popup's own
        StaysOpen="True" (see XAML), it never captures the mouse or
        auto-dismisses itself, so there's no race left to guard against here
        - this is genuinely just "closed -> open it, open -> close it".
        Closing on an outside click is handled separately, by
        _options_popup_outside_click below."""
        new_state = not self.OptionsPopup.IsOpen
        self.OptionsPopup.IsOpen = new_state
        self.options_btn.IsChecked = new_state

    def _options_popup_outside_click(self, sender, args):
        """Wired to the window's own PreviewMouseDown. Closes the options
        popup on a click anywhere outside it - except on the hamburger
        button itself, which already toggles it via options_btn_click above;
        if this handler also reacted to a click there, the popup would close
        and then immediately reopen (or the reverse), which is exactly the
        bug this whole approach was designed to avoid."""
        if not self.OptionsPopup.IsOpen:
            return
        try:
            source = args.OriginalSource
            if self._is_visual_descendant(source, self.options_btn):
                return
            popup_content = self.OptionsPopup.Child
            if popup_content is not None and self._is_visual_descendant(source, popup_content):
                return
            self.OptionsPopup.IsOpen = False
            self.options_btn.IsChecked = False
        except Exception:
            pass

    def _is_visual_descendant(self, element, ancestor):
        """True if element is ancestor itself, or nested anywhere inside it."""
        import System.Windows.Media as _M
        node = element
        while node is not None:
            if node == ancestor:
                return True
            try:
                node = _M.VisualTreeHelper.GetParent(node)
            except Exception:
                return False
        return False

    def menu_issue_click(self, sender, args):
        """☰ → Report an issue: open a new GitHub issue, pre-filled with the
        app name, Seed43 version and Revit version."""
        from Snippets._support import github_issue_url
        self.OptionsPopup.IsOpen   = False
        self.options_btn.IsChecked = False
        self._open_url(github_issue_url("pyTransmit", _SCRIPT_DIR_MAIN),
                       title="Report an issue")

    def menu_about_click(self, sender, args):
        """☰ → About: open ABOUT_URL in the default browser."""
        self.OptionsPopup.IsOpen   = False
        self.options_btn.IsChecked = False
        self._open_url(ABOUT_URL, title="About")

    def menu_support_click(self, sender, args):
        """☰ → Support: open SUPPORT_URL in the default browser."""
        self.OptionsPopup.IsOpen   = False
        self.options_btn.IsChecked = False
        self._open_url(SUPPORT_URL, title="Support")

    def close_about_click(self, sender, args):
        """Legacy close handler, modal no longer used but kept for safety."""
        try:
            self.AboutModal.Visibility = System.Windows.Visibility.Collapsed
            self.Overlay.Visibility    = System.Windows.Visibility.Collapsed
        except: pass

    def menu_file_naming_click(self, sender, args):
        """☰ → File Naming Settings panel."""
        self.OptionsPopup.IsOpen = False
        try: self.options_btn.IsChecked = False
        except: pass
        if self.filenaming_ctrl:
            self.filenaming_ctrl.load_config()
            self.filenaming_ctrl.refresh_live_values()
        self._show_panel('file_naming')

    def menu_export_click(self, sender, args):
        """☰ → Export Settings: open the Export Settings panel."""
        self.OptionsPopup.IsOpen = False
        self.options_btn.IsChecked = False
        self._populate_layout_combos()
        self._show_panel("export_settings")

    def menu_import_click(self, sender, args):
        """☰ → Import Settings: open the Import Settings panel."""
        self.OptionsPopup.IsOpen = False
        self.options_btn.IsChecked = False
        self._show_panel("import_settings")

    # ── Sync config (persists export/import preferences) ─────────────────

    # ── Recipient row builders (called by SetupSettingsController.apply) ─────

    _dist_rows   = []   # list of {label, attn_tb, copies_tb}
    _client_rows = []   # list of {dp, company_cb, contact_cb, copies_tb}
    brand_ctrl       = None # BrandingSettingsController (set in _init_controllers)
    filenaming_ctrl  = None # FileNamingSettingsController (set in _init_controllers)

    def _populate_option_combos(self):
        """Fill reason/method/format/printsize dropdowns from OptionsSettings data.
        Always inserts '(none)' as item 0. After filling, pre-fills selection
        from the last issued revision's IssuedTo string."""
        try:
            if self.opt_ctrl is None:
                return

            def format_coded(row):
                code = getattr(row, 'Code', '') or ''
                sep  = getattr(row, 'Separator', '=') or '='
                desc = getattr(row, 'Description', '') or ''
                if code and desc:
                    return u"{} {} {}".format(code, sep, desc)
                elif code:
                    return code
                return desc

            def format_simple(row):
                return getattr(row, 'Value', '') or str(row)

            def fill_coded(cb_name, data_list):
                cb = getattr(self, cb_name, None)
                if cb is None:
                    return
                items = ['(none)'] + [format_coded(r) for r in data_list]
                cb.ItemsSource = items
                cb.SelectedIndex = 0  # default: (none)

            def fill_simple(cb_name, data_list):
                cb = getattr(self, cb_name, None)
                if cb is None:
                    return
                items = ['(none)'] + [format_simple(r) for r in data_list]
                cb.ItemsSource = items
                cb.SelectedIndex = 0  # default: (none)

            fill_coded('reason_cb',     self.opt_ctrl.reason_data)
            fill_coded('method_cb',     self.opt_ctrl.method_data)
            fill_simple('format_cb',    self.opt_ctrl.format_data)
            fill_simple('printsize_cb', self.opt_ctrl.printsize_data)

            # Pre-fill all combos + recipient fields from last issued revision
            self._prefill_from_last_revision()
        except:
            pass

    def _prefill_from_last_revision(self):
        """
        Pre-fill the main window fields from the most recently issued revision.
        - Dropdowns (reason/method/format/printsize): match by code/value
        - Initials textbox: from IssuedBy
        - Distribution rows: attn + copies from the IssuedTo recipients block
        - Client rows: not pre-filled (client mode uses free-form selection)

        If no issued revisions exist, all fields stay at (none)/blank.
        """
        try:
            if not self.issued_revs:
                return

            last = self.issued_revs[-1]
            issued_to  = last.IssuedTo  or ''
            issued_by  = (last.IssuedBy or '').strip()

            import re as _re

            def parse_tag(s, tag):
                """Parse TAG:value from IssuedTo string."""
                m = _re.search(r'(?:^| )' + _re.escape(tag) + r':([^ |]+)', s)
                return m.group(1).strip() if m else ''

            reason_code  = parse_tag(issued_to, 'R')
            method_code  = parse_tag(issued_to, 'M')
            format_val   = parse_tag(issued_to, 'F')
            size_val     = parse_tag(issued_to, 'S')
            rev_type_val = parse_tag(issued_to, 'RT')

            # ── Initials ──────────────────────────────────────────────────
            try:
                tb = getattr(self, 'initials_tb', None)
                if tb and issued_by:
                    tb.Text = issued_by
            except: pass

            # ── Coded dropdowns (reason, method) ─────────────────────────
            def select_coded(cb_name, data_list, code):
                if not code:
                    return
                cb = getattr(self, cb_name, None)
                if cb is None:
                    return
                # items[0] is '(none)', items[1..] correspond to data_list[0..]
                for i, row in enumerate(data_list):
                    if (getattr(row, 'Code', '') or '').strip().lower() == code.lower():
                        try: cb.SelectedIndex = i + 1  # +1 for (none)
                        except: pass
                        return

            # ── Simple dropdowns (format, printsize) ─────────────────────
            def select_simple(cb_name, data_list, val):
                if not val:
                    return
                cb = getattr(self, cb_name, None)
                if cb is None:
                    return
                for i, row in enumerate(data_list):
                    if (getattr(row, 'Value', '') or str(row)).strip().lower() == val.lower():
                        try: cb.SelectedIndex = i + 1  # +1 for (none)
                        except: pass
                        return

            # Restore rev numbering type selection
            try:
                if rev_type_val:
                    _cb = getattr(self, 'rev_type_cb', None)
                    if _cb and rev_type_val in list(_cb.ItemsSource or []):
                        _cb.SelectedItem = rev_type_val
                        self._rev_numbering_type = rev_type_val
            except Exception: pass

            if self.opt_ctrl:
                select_coded ('reason_cb',    self.opt_ctrl.reason_data,    reason_code)
                select_coded ('method_cb',    self.opt_ctrl.method_data,    method_code)
                select_simple('format_cb',    self.opt_ctrl.format_data,    format_val)
                select_simple('printsize_cb', self.opt_ctrl.printsize_data, size_val)

            # ── Distribution List rows, attn + copies ────────────────────
            try:
                # Recipients block is after " | DL: " or " | CL: "
                recip_block = ''
                _saved_as_client = False
                for _part in issued_to.split(' | '):
                    _part = _part.strip()
                    if _part.startswith('DL:'):
                        recip_block = _part[3:].strip()
                        break
                    elif _part.startswith('CL:'):
                        # Last revision was saved in client mode, skip dist row fill
                        _saved_as_client = True
                        break
                # Fallback: old format (no DL:/CL: prefix), second pipe-block
                if not recip_block and not _saved_as_client:
                    _blocks = issued_to.split(' | ')
                    if len(_blocks) > 1:
                        recip_block = _blocks[1].strip()

                # Never populate dist rows from client-mode data
                if _saved_as_client:
                    recip_block = ''

                # Parse new format: 1A.[attn]copies  or old format: A.[attn]copies
                tokens = _re.findall(r'(\d*)([A-Za-z]+)\.\[([^\]]*)\](\d*)', recip_block)
                # Build map: index (1-based) → (attn, copies), fallback to letter
                index_map  = {}
                letter_map = {}
                for num, letters, attn, copies in tokens:
                    if num:
                        index_map[int(num)] = (attn, copies)
                    letter_map[letters[0].upper()] = (attn, copies)

                for _ri, row in enumerate(self._dist_rows):
                    label = row.get('label', '')
                    # Prefer index match, fall back to first-letter match
                    if (_ri + 1) in index_map:
                        attn, copies = index_map[_ri + 1]
                    elif label and label[:1].upper() in letter_map:
                        attn, copies = letter_map[label[:1].upper()]
                    else:
                        continue
                    try: row['attn_tb'].Text   = attn
                    except: pass
                    try: row['copies_tb'].Text = copies
                    except: pass
            except: pass

        except:
            pass

    def _prefill_project_info(self):
        """Pre-fill Organisation/Client/Project No/Project textboxes from Revit."""
        try:
            pi = self.doc.ProjectInformation
            def _gp(name):
                try:
                    p = pi.LookupParameter(name)
                    if p and p.HasValue:
                        return (p.AsString() or p.AsValueString() or '').strip()
                except: pass
                return ''
            _vals = {
                'org_name_tb':    _gp('Organization Name'),
                'client_name_tb': _gp('Client Name'),
                'proj_number_tb': _gp('Project Number'),
                'proj_name_tb':   _gp('Project Name') or self.doc.Title or '',
            }
            for _tb_name, _val in _vals.items():
                tb = getattr(self, _tb_name, None)
                if tb is not None:
                    try: tb.Text = _val
                    except: pass
        except:
            pass

    def _check_settings_mismatch(self):
        """
        Compare current Setup settings against the VIS/EX tags stored in the
        last issued revision's IssuedTo field. If they differ, show a styled
        prompt giving the user three choices:
          - Update Settings  : permanently apply the project's snapshotted settings
          - This Issue Only  : temporarily apply for this session only
          - Ignore           : proceed with current settings unchanged
        """
        if not self.issued_revs:
            return
        if not self.setup_ctrl:
            return

        import re as _re

        last_ito = (self.issued_revs[-1].IssuedTo or '').strip()

        # Parse VIS tag, if present use it directly.
        # If absent, check whether this looks like a pyTransmit revision (has R:/M:/EX: tags).
        # If it does, VIS was omitted because all fields were off → treat as empty set.
        # If it doesn't, it's a genuinely old pre-pyTransmit revision → assume all-on.
        _vis_m = _re.search(r'\|?\s*VIS:([\w,]*)', last_ito)
        if _vis_m:
            _vis_val = _vis_m.group(1).strip()
            _proj_vis = set(_vis_val.split(',')) if _vis_val else set()
        else:
            _is_pytransmit = bool(_re.search(r'\b(?:R:|M:|EX:|DL:|CL:|RPG:)', last_ito))
            if _is_pytransmit:
                # pyTransmit revision issued with all info rows off, VIS tag was skipped
                _proj_vis = set()
            else:
                # Genuine old revision with no pyTransmit tags, assume all rows were on
                _proj_vis = {'FR', 'CL', 'PN', 'PJ'}

        # Parse EX tag, if absent infer from context (old revisions assumed schedule only)
        _ex_m = _re.search(r'\|?\s*EX:([\w,]+)', last_ito)
        if _ex_m:
            _proj_ex = set(_ex_m.group(1).split(','))
        else:
            _proj_ex = {'RS'}  # old revisions assumed schedule only

        # Build current user's VIS set
        cfg = self.setup_ctrl.cfg
        _cur_vis = set()
        if cfg.get('show_from',     True): _cur_vis.add('FR')
        if cfg.get('show_client',   True): _cur_vis.add('CL')
        if cfg.get('show_projno',   True): _cur_vis.add('PN')
        if cfg.get('show_projname', True): _cur_vis.add('PJ')

        # Build current user's EX set
        _cur_ex = set()
        if cfg.get('out_schedule',  True):  _cur_ex.add('RS')
        if cfg.get('out_drafting',  False): _cur_ex.add('RD')
        if cfg.get('out_legend',    False): _cur_ex.add('RL')
        if cfg.get('out_excel',     False): _cur_ex.add('Excl')
        if cfg.get('out_pdf',       False): _cur_ex.add('PDF')

        if _cur_vis == _proj_vis and _cur_ex == _proj_ex:
            return  # All good, no mismatch

        # Build human-readable diff
        _VIS_LABELS = {'FR': 'Organisation', 'CL': 'Client', 'PN': 'Project No.', 'PJ': 'Project'}
        _EX_LABELS  = {'RS': 'Revit Schedule', 'RD': 'Drafting View', 'RL': 'Legend',
                       'Excl': 'Excel', 'PDF': 'PDF'}
        _diff_lines = []
        for _code in sorted((_proj_vis | _cur_vis)):
            _lbl = _VIS_LABELS.get(_code, _code)
            if _code in _proj_vis and _code not in _cur_vis:
                _diff_lines.append(u'  {} — was ON, now OFF'.format(_lbl))
            elif _code not in _proj_vis and _code in _cur_vis:
                _diff_lines.append(u'  {} — was OFF, now ON'.format(_lbl))
        for _code in sorted((_proj_ex | _cur_ex)):
            _lbl = _EX_LABELS.get(_code, _code)
            if _code in _proj_ex and _code not in _cur_ex:
                _diff_lines.append(u'  {} — was ON, now OFF'.format(_lbl))
            elif _code not in _proj_ex and _code in _cur_ex:
                _diff_lines.append(u'  {} — was OFF, now ON'.format(_lbl))
        _diff_text = u'\n'.join(_diff_lines)

        # Show styled mismatch dialog
        _result = self._show_mismatch_dialog(_diff_text)

        if _result == 'update':
            # Permanently update Setup settings to match project snapshot
            self._apply_vis_ex_to_setup(_proj_vis, _proj_ex, permanent=True)
        elif _result == 'session':
            # Apply for this session only, don't save to disk
            self._apply_vis_ex_to_setup(_proj_vis, _proj_ex, permanent=False)
        # 'ignore', do nothing

    def _apply_vis_ex_to_setup(self, vis_set, ex_set, permanent=False):
        """Apply a VIS+EX snapshot to the Setup controller."""
        if not self.setup_ctrl:
            return
        h = self

        # Engage the re-entrancy guard so checkbox Checked/Unchecked events
        # do NOT call save() while we're programmatically setting values
        self.setup_ctrl._applying = True
        try:
            def set_cb(name, val):
                el = getattr(h, name, None)
                if el is not None:
                    try: el.IsChecked = val
                    except: pass

            set_cb('setup_from_cb',              'FR'   in vis_set)
            set_cb('setup_client_info_cb',       'CL'   in vis_set)
            set_cb('setup_projno_cb',            'PN'   in vis_set)
            set_cb('setup_projname_cb',          'PJ'   in vis_set)
            set_cb('setup_output_schedule_cb',   'RS'   in ex_set)
            set_cb('setup_output_drafting_cb',   'RD'   in ex_set)
            set_cb('setup_output_legend_cb',     'RL'   in ex_set)
            set_cb('setup_output_excel_cb',      'Excl' in ex_set)
            set_cb('setup_output_pdf_cb',        'PDF'  in ex_set)
        finally:
            self.setup_ctrl._applying = False

        if permanent:
            self.setup_ctrl.save()
        self.setup_ctrl.apply()

    def _show_mismatch_dialog(self, diff_text=''):
        from Dialogs import Dialogs as _D
        return _D.settings_mismatch(diff_text)

    def _build_dist_rows(self):
        """Build fixed Distribution List rows in the main window from distribution.json."""
        import json as _json
        import System.Windows as _SW
        import System.Windows.Controls as _SWC
        import System.Windows.Media as _SWM

        stack = getattr(self, 'dist_rows_stack', None)
        if stack is None:
            return
        stack.Children.Clear()
        self._dist_rows = []

        dist_file = settings_file('distribution.json')
        rows = []
        try:
            with open(dist_file, 'r') as f:
                rows = _json.load(f)
        except:
            rows = [{'distribution': 'Architect/Designer'},
                    {'distribution': 'Owner/Developer'},
                    {'distribution': 'Contractor'},
                    {'distribution': 'Local Authority'}]

        white = self.TryFindResource('BrushTextPrimary') or _SWM.SolidColorBrush(_SWM.Color.FromRgb(0xF4, 0xFA, 0xFF))

        for item in rows:
            label = (item.get('distribution', '')
                     or item.get('Distribution', '')
                     or str(item))

            # Row: DockPanel, label stretches, copies fixed right, attn fills middle
            dp = _SWC.DockPanel()
            dp.Margin = _SW.Thickness(0, 0, 0, 4)
            dp.LastChildFill = True

            # Label, left side, fixed width
            lbl = _SWC.TextBlock()
            lbl.Text = label
            lbl.Foreground = white
            lbl.FontSize = 12
            lbl.Width = 140
            lbl.VerticalAlignment = _SW.VerticalAlignment.Center
            lbl.Margin = _SW.Thickness(0, 0, 6, 0)
            _SWC.DockPanel.SetDock(lbl, _SWC.Dock.Left)
            dp.Children.Add(lbl)

            # Copies, right side, fixed width
            copies_tb = _SWC.TextBox()
            copies_tb.Width = 54
            copies_tb.HorizontalContentAlignment = _SW.HorizontalAlignment.Center
            copies_tb.Margin = _SW.Thickness(4, 0, 0, 0)
            _SWC.DockPanel.SetDock(copies_tb, _SWC.Dock.Right)
            try: copies_tb.Style = self.FindResource("TextBoxStyle")
            except: pass
            dp.Children.Add(copies_tb)

            # Attention To, fills remaining space
            attn_tb = _SWC.TextBox()
            attn_tb.HorizontalAlignment = _SW.HorizontalAlignment.Stretch
            try: attn_tb.Style = self.FindResource("TextBoxStyle")
            except: pass
            dp.Children.Add(attn_tb)

            stack.Children.Add(dp)
            self._dist_rows.append({
                'label':     label,
                'attn_tb':  attn_tb,
                'copies_tb': copies_tb,
            })

        # Pre-fill attn/copies from last issued revision
        self._prefill_from_last_revision()

    def _build_client_rows(self):
        """
        Build client recipient rows in the main window.
        Layout per row: [Company v] [Contact v] [Copies]
        Selecting a Company filters the Contact dropdown to that company's contacts.
        Selecting any value in the last row adds a new blank row below it.
        """
        import System.Windows as _SW
        import System.Windows.Controls as _SWC
        import System.Windows.Media as _SWM

        stack = getattr(self, 'client_rows_stack', None)
        if stack is None:
            return
        stack.Children.Clear()
        self._client_rows = []

        # Load full recipients data: [{company, attention_to}]
        self._client_data = self._load_client_data()

        # Column header row: Company | Contact | Copies
        try:
            white = self.TryFindResource('BrushTextPrimary') or _SWM.SolidColorBrush(_SWM.Color.FromRgb(0xF4, 0xFA, 0xFF))
            hdr_dp = _SWC.DockPanel()
            hdr_dp.Margin = _SW.Thickness(0, 0, 0, 2)

            # Copies label, fixed width, docked right
            copies_lbl = _SWC.TextBlock()
            copies_lbl.Text = "Copies"
            copies_lbl.Width = 46
            copies_lbl.Foreground = white
            copies_lbl.FontSize = 11
            copies_lbl.HorizontalAlignment = _SW.HorizontalAlignment.Center
            copies_lbl.Margin = _SW.Thickness(4, 0, 0, 0)
            _SWC.DockPanel.SetDock(copies_lbl, _SWC.Dock.Right)
            hdr_dp.Children.Add(copies_lbl)

            # Company / Contact labels, split evenly
            hdr_grid = _SWC.Grid()
            hc0 = _SWC.ColumnDefinition()
            hc0.Width = _SW.GridLength(1, _SW.GridUnitType.Star)
            hc1 = _SWC.ColumnDefinition()
            hc1.Width = _SW.GridLength(1, _SW.GridUnitType.Star)
            hdr_grid.ColumnDefinitions.Add(hc0)
            hdr_grid.ColumnDefinitions.Add(hc1)

            for _ci, _txt in enumerate(["Company", "Contact"]):
                _lbl = _SWC.TextBlock()
                _lbl.Text = _txt
                _lbl.Foreground = white
                _lbl.FontSize = 11
                _lbl.VerticalAlignment = _SW.VerticalAlignment.Center
                _lbl.Margin = _SW.Thickness(0 if _ci == 0 else 4, 0, 4, 0)
                _SWC.Grid.SetColumn(_lbl, _ci)
                hdr_grid.Children.Add(_lbl)
            hdr_dp.Children.Add(hdr_grid)
            stack.Children.Add(hdr_dp)
        except Exception:
            pass   # header labels optional, don't let them block row building

        # Pre-fill rows from last issued revision if it was saved in client mode
        _prefilled = False
        try:
            if self.issued_revs:
                import re as _re_cl
                _last_ito = (self.issued_revs[-1].IssuedTo or '')
                _cl_m = _re_cl.search(r'CL:\s*(.*?)(?:\s*\|[^|]|$)', _last_ito)
                if _cl_m:
                    _cl_block = _cl_m.group(1).strip()
                    # Format: [Company, Contact]copies  or  [Company]copies
                    _tokens = _re_cl.findall(r'\[([^\]]+)\](\d*)', _cl_block)
                    for _label_full, _copies in _tokens:
                        if u'\u2014' in _label_full:
                            _parts = _label_full.split(u'\u2014', 1)
                            _company = _parts[0].strip()
                            _attn    = _parts[1].strip()
                        else:
                            _company = _label_full.strip()
                            _attn    = ''
                        if _company:
                            self._add_client_row(
                                preset_company=_company,
                                preset_attn=_attn,
                                preset_copies=_copies)
                            _prefilled = True
        except Exception:
            pass

        # Always end with one blank row for new input
        self._add_client_row()

    def _load_client_data(self):
        """Return list of {'company': str, 'attn': str} dicts from recipients.json."""
        import json as _json
        rec_file = settings_file('recipients.json')
        try:
            with open(rec_file, 'r') as f:
                raw = _json.load(f)
            result = []
            for r in raw:
                company = (r.get('company') or r.get('Company') or
                           r.get('recipient') or '').strip()
                attn    = (r.get('attention_to') or r.get('AttentionTo') or
                           r.get('Attention') or '').strip()
                if company:
                    result.append({'company': company, 'attn': attn})
            return result
        except Exception:
            return []

    def _get_companies(self):
        """Return sorted unique company names from client data."""
        seen = []
        for r in getattr(self, '_client_data', []):
            c = r['company']
            if c not in seen:
                seen.append(c)
        return seen

    def _get_contacts_for_company(self, company):
        """Return list of attention_to values for a given company."""
        return [r['attn'] for r in getattr(self, '_client_data', [])
                if r['company'] == company and r['attn']]

    def _add_client_row(self, preset_company='', preset_attn='', preset_copies=''):
        """
        Add one Company/Contact/Copies row to client_rows_stack.
        When company is selected, Contact dropdown is populated.
        When any column changes in the last row, a new blank row is added.
        """
        import System.Windows as _SW
        import System.Windows.Controls as _SWC
        import System.Windows.Media as _SWM

        stack = getattr(self, 'client_rows_stack', None)
        if stack is None:
            return

        white = self.TryFindResource('BrushTextPrimary') or _SWM.SolidColorBrush(_SWM.Color.FromRgb(0xF4, 0xFA, 0xFF))

        # Outer DockPanel: copies fixed right, company+contact fill left
        dp = _SWC.DockPanel()
        dp.Margin = _SW.Thickness(0, 0, 0, 4)
        dp.LastChildFill = True

        # Copies textbox, docked right
        copies_tb = _SWC.TextBox()
        copies_tb.Width = 46
        copies_tb.HorizontalContentAlignment = _SW.HorizontalAlignment.Center
        copies_tb.Margin = _SW.Thickness(4, 0, 0, 0)
        copies_tb.Text = preset_copies
        _SWC.DockPanel.SetDock(copies_tb, _SWC.Dock.Right)
        try: copies_tb.Style = self.FindResource("TextBoxStyle")
        except: pass
        dp.Children.Add(copies_tb)

        # Inner Grid: two equal columns for Company | Contact
        grid = _SWC.Grid()
        col0 = _SWC.ColumnDefinition()
        col0.Width = _SW.GridLength(1, _SW.GridUnitType.Star)
        col1 = _SWC.ColumnDefinition()
        col1.Width = _SW.GridLength(1, _SW.GridUnitType.Star)
        grid.ColumnDefinitions.Add(col0)
        grid.ColumnDefinitions.Add(col1)
        dp.Children.Add(grid)

        # Company dropdown
        companies = self._get_companies()
        company_cb = _SWC.ComboBox()
        company_cb.ItemsSource = ['(Select Company)'] + companies
        company_cb.SelectedIndex = 0
        company_cb.Margin = _SW.Thickness(0, 0, 4, 0)
        _SWC.Grid.SetColumn(company_cb, 0)
        try: company_cb.Style = self.FindResource("ComboBoxStyle")
        except: pass
        grid.Children.Add(company_cb)

        # Contact dropdown, left margin matches the gap between company and grid edge
        contact_cb = _SWC.ComboBox()
        contact_cb.ItemsSource = ['(Select Contact)']
        contact_cb.SelectedIndex = 0
        contact_cb.IsEnabled = False
        contact_cb.Margin = _SW.Thickness(4, 0, 0, 0)
        _SWC.Grid.SetColumn(contact_cb, 1)
        try: contact_cb.Style = self.FindResource("ComboBoxStyle")
        except: pass
        grid.Children.Add(contact_cb)

        row_ref = {
            'dp':        dp,
            'company_cb': company_cb,
            'contact_cb': contact_cb,
            'copies_tb':  copies_tb,
        }
        self._client_rows.append(row_ref)
        stack.Children.Add(dp)

        # Pre-fill if preset values given (restoring saved state)
        if preset_company and preset_company in companies:
            try:
                company_cb.SelectedItem = preset_company
                contacts = self._get_contacts_for_company(preset_company)
                contact_cb.ItemsSource = ['(Select Contact)'] + contacts
                contact_cb.IsEnabled = True
                if preset_attn and preset_attn in contacts:
                    contact_cb.SelectedItem = preset_attn
                elif contacts:
                    contact_cb.SelectedIndex = 1
                else:
                    contact_cb.SelectedIndex = 0
            except Exception:
                pass

        def _is_last_row(rr):
            return rr is self._client_rows[-1]

        def _maybe_add_row(rr=row_ref):
            """Add a new blank row if this is the last one and it now has a company."""
            if (_is_last_row(rr)
                    and rr['company_cb'].SelectedIndex > 0):
                self._add_client_row()

        def _prune_trailing_empty(rr=row_ref):
            """Remove trailing empty rows when this row is cleared."""
            idx = next((i for i, r in enumerate(self._client_rows) if r is rr), -1)
            if idx < 0:
                return
            to_remove = []
            for r in reversed(self._client_rows[idx + 1:]):
                if r['company_cb'].SelectedIndex == 0:
                    to_remove.append(r)
                else:
                    break
            for r in to_remove:
                try: stack.Children.Remove(r['dp'])
                except: pass
                try: self._client_rows.remove(r)
                except: pass

        def on_company_changed(s, e, rr=row_ref):
            if s.SelectedIndex == 0:
                rr['contact_cb'].ItemsSource = ['(Select Contact)']
                rr['contact_cb'].SelectedIndex = 0
                rr['contact_cb'].IsEnabled = False
                _prune_trailing_empty(rr)
                return
            company = str(s.SelectedItem or '')
            contacts = self._get_contacts_for_company(company)
            rr['contact_cb'].ItemsSource = ['(Select Contact)'] + contacts
            rr['contact_cb'].IsEnabled = True
            if contacts:
                rr['contact_cb'].SelectedIndex = 1
            else:
                rr['contact_cb'].SelectedIndex = 0
            _maybe_add_row(rr)

        def on_contact_changed(s, e, rr=row_ref):
            _maybe_add_row(rr)

        company_cb.SelectionChanged += on_company_changed
        contact_cb.SelectionChanged += on_contact_changed

    # ── Export / Import panel handlers (delegate to controllers) ──────────
    def export_browse_click(self, sender, args):
        if self.export_ctrl: self.export_ctrl.on_browse(sender, args)

    def export_execute_click(self, sender, args):
        if self.export_ctrl: self.export_ctrl.on_execute(sender, args)
        # Also export Layout templates if checkbox is checked
        try:
            cb = getattr(self, 'export_layouts_cb', None)
            if cb is None or cb.IsChecked:  # default True if checkbox missing
                self._export_layouts()
        except Exception:
            self._export_layouts()

    def _export_layouts(self):
        """Copy Layout/Layouts/*.json to the export folder alongside other settings."""
        try:
            export_path = ''
            tb = getattr(self, 'export_path_tb', None)
            if tb and tb.Text: export_path = tb.Text.strip()
            if not export_path or not os.path.isdir(export_path):
                return  # no valid export path, skip silently

            script_dir = os.path.dirname(os.path.abspath(__file__))
            src_layouts = LAYOUTS_DIR
            if not os.path.isdir(src_layouts):
                return  # no layouts to export

            dest_settings = os.path.join(export_path, 'pyTransmit Settings')
            dest_layouts  = os.path.join(dest_settings, 'Layouts')
            if not os.path.isdir(dest_layouts):
                os.makedirs(dest_layouts)

            # Also copy layout_config.json
            src_config = LAYOUT_CONFIG
            if os.path.isfile(src_config):
                import shutil
                shutil.copy2(src_config, os.path.join(dest_settings, 'layout_config.json'))

            # Copy all template JSONs
            import shutil
            for fn in os.listdir(src_layouts):
                if fn.lower().endswith('.json'):
                    shutil.copy2(os.path.join(src_layouts, fn),
                                 os.path.join(dest_layouts, fn))
        except Exception:
            pass  # silent, don't break the main export if layouts fail

    def export_settings_back_click(self, sender, args):
        if self.export_ctrl: self.export_ctrl.save_config()
        self._save_layout_assignments()
        self._show_panel("main")

    def layout_assignment_changed(self, sender, args):
        """Called when any layout combo selection changes, save immediately."""
        self._save_layout_assignments()

    def _layouts_dir(self):
        return LAYOUTS_DIR

    # Both builders keep their own templates, in their own folders, and both
    # can hold one called "Excel". The dropdowns therefore say which builder
    # each entry came from rather than listing two identical-looking names.
    LB_PREFIX = 'LB - '
    STUDIO_PREFIX = 'Studio - '

    def _layout_template_map(self):
        """OrderedDict of {dropdown label: full path to the template JSON},
        covering the Layout Builder's folder and Studio's.

        A plain directory listing of each, so saving a new layout in either
        builder makes it appear in the format dropdowns with nothing else to
        update.
        """
        from collections import OrderedDict
        out = OrderedDict()
        for prefix, folder in ((self.LB_PREFIX, self._layouts_dir()),
                               (self.STUDIO_PREFIX, STUDIO_LAYOUTS_DIR)):
            if not os.path.isdir(folder):
                continue
            for f in sorted(os.listdir(folder)):
                if f.lower().endswith('.json'):
                    out[prefix + os.path.splitext(f)[0]] = os.path.join(folder, f)
        return out

    def _layout_templates(self):
        """Dropdown labels for every template, both builders."""
        return list(self._layout_template_map().keys())

    def _resolve_layout_label(self, saved, default):
        """Which dropdown entry a stored assignment means.

        Assignments saved before Studio templates were listed hold a bare
        name ("Excel") - those were always Layout Builder templates, so they
        resolve to the LB entry. Falls back to the format's default name,
        then to '(none)'.
        """
        labels = self._layout_templates()
        for candidate in (saved, self.LB_PREFIX + str(saved or ''),
                          self.STUDIO_PREFIX + str(saved or ''),
                          self.LB_PREFIX + default, self.STUDIO_PREFIX + default):
            if candidate and candidate in labels:
                return candidate
        return '(none)'

    _LAYOUT_COMBOS = {
        'layout_schedule_cb': 'Revit Schedule',
        'layout_drafting_cb': 'Revit Drafting View',
        'layout_legend_cb':   'Revit Legend',
        'layout_excel_cb':    'Excel',
        'layout_pdf_cb':      'PDF',
    }

    def _on_content_rendered(self, sender, args):
        """Called after window is fully rendered, visual tree is available."""
        try:
            self._apply_green_scrollbars()
        except:
            pass

    def _apply_green_scrollbars(self):
        """Apply green scrollbar style to all vertical ScrollBars via VisualTreeHelper."""
        try:
            import System.Windows
            import System.Windows.Controls
            import System.Windows.Controls.Primitives as Prim
            import System.Windows.Media as Media
            from System.Windows.Controls import ControlTemplate
            from System.Windows import FrameworkElementFactory as FEF, Style, Setter, Thickness, CornerRadius

            # ── Thumb template ────────────────────────────────────────────
            thumb_border = FEF(System.Windows.Controls.Border)
            thumb_border.SetValue(
                System.Windows.Controls.Border.BackgroundProperty,
                self.TryFindResource('BrushPrimaryGreen') or Media.SolidColorBrush(Media.Color.FromRgb(0x20, 0x8A, 0x3C)))
            thumb_border.SetValue(
                System.Windows.Controls.Border.CornerRadiusProperty,
                CornerRadius(3))
            thumb_border.SetValue(
                System.Windows.FrameworkElement.MarginProperty,
                Thickness(2))
            thumb_tpl = ControlTemplate(Prim.Thumb)
            thumb_tpl.VisualTree = thumb_border

            thumb_style = Style()
            thumb_style.TargetType = Prim.Thumb
            thumb_style.Setters.Add(Setter(
                System.Windows.Controls.Control.TemplateProperty, thumb_tpl))

            # ── Track inside a Grid ───────────────────────────────────────
            thumb_fac = FEF(Prim.Thumb)
            thumb_fac.SetValue(
                System.Windows.FrameworkElement.StyleProperty, thumb_style)

            track_fac = FEF(Prim.Track)
            track_fac.SetValue(
                Prim.Track.OrientationProperty,
                System.Windows.Controls.Orientation.Vertical)
            track_fac.SetValue(Prim.Track.IsDirectionReversedProperty, True)
            track_fac.Name = 'PART_Track'
            track_fac.AppendChild(thumb_fac)

            grid_fac = FEF(System.Windows.Controls.Grid)
            grid_fac.SetValue(System.Windows.FrameworkElement.WidthProperty, 8.0)
            grid_fac.AppendChild(track_fac)

            # ── ScrollBar style ───────────────────────────────────────────
            sb_tpl = ControlTemplate(Prim.ScrollBar)
            sb_tpl.VisualTree = grid_fac

            sb_style = Style()
            sb_style.TargetType = Prim.ScrollBar
            sb_style.Setters.Add(Setter(
                System.Windows.Controls.Control.TemplateProperty, sb_tpl))
            sb_style.Setters.Add(Setter(
                System.Windows.FrameworkElement.WidthProperty, 8.0))
            sb_style.Setters.Add(Setter(
                System.Windows.FrameworkElement.MinWidthProperty, 8.0))
            sb_style.Setters.Add(Setter(
                System.Windows.Controls.Control.BackgroundProperty,
                Media.Brushes.Transparent))

            # ── Walk visual tree ──────────────────────────────────────────
            def walk(el):
                try:
                    n = Media.VisualTreeHelper.GetChildrenCount(el)
                    for i in range(n):
                        child = Media.VisualTreeHelper.GetChild(el, i)
                        if (isinstance(child, Prim.ScrollBar) and
                                child.Orientation ==
                                System.Windows.Controls.Orientation.Vertical):
                            child.Style = sb_style
                        else:
                            walk(child)
                except Exception:
                    pass

            walk(self)
        except Exception:
            pass  # Never crash the window over cosmetics

    def _populate_layout_combos(self):
        """Populate all layout dropdowns with available templates, set saved selection."""
        templates = self._layout_templates()
        if not templates: return
        saved = self._load_layout_assignments()
        for cb_name, default in self._LAYOUT_COMBOS.items():
            cb = getattr(self, cb_name, None)
            if not cb: continue
            cb.Items.Clear()
            cb.Items.Add('(none)')
            for t in templates:
                cb.Items.Add(t)
            cb.SelectedItem = self._resolve_layout_label(
                saved.get(cb_name), default)

    def _save_layout_assignments(self):
        """Persist layout combo selections to pytransmit_sync.json."""
        try:
            sync_path = SYNC_FILE
            cfg = {}
            if os.path.isfile(sync_path):
                with open(sync_path, 'r') as f: cfg = _json.load(f)
            for cb_name in self._LAYOUT_COMBOS:
                cb = getattr(self, cb_name, None)
                if cb and cb.SelectedItem:
                    cfg['layout_assign_{}'.format(cb_name)] = str(cb.SelectedItem)
            with open(sync_path, 'w') as f: _json.dump(cfg, f, indent=2)
        except Exception: pass

    def _load_sync(self):
        """Load pytransmit_sync.json as a dict."""
        try:
            sync_path = SYNC_FILE
            if os.path.isfile(sync_path):
                with open(sync_path, 'r') as f: return _json.load(f)
        except Exception: pass
        return {}

    def _save_sync(self):
        """Persist non-layout settings (group_label_on, etc.) to pytransmit_sync.json."""
        try:
            sync_path = SYNC_FILE
            cfg = self._load_sync()
            cfg['group_label_on'] = getattr(self, 'group_label_on', True)
            with open(sync_path, 'w') as f: _json.dump(cfg, f, indent=2)
        except Exception: pass

    def _load_layout_assignments(self):
        """Load layout combo selections from pytransmit_sync.json."""
        try:
            cfg = self._load_sync()
            return {k.replace('layout_assign_', ''): v
                    for k, v in cfg.items() if k.startswith('layout_assign_')}
        except Exception: pass
        return {}

    def get_layout_for_output(self, output_type):
        """Return the full path to the selected layout JSON for a given output type.
        output_type: 'excel', 'pdf', 'schedule', 'drafting', 'legend'

        Resolved through _layout_template_map() rather than by joining a name
        onto the Layouts folder, since the label now says which builder the
        template belongs to and Studio's live somewhere else entirely.
        """
        cb_map = {'excel': 'layout_excel_cb', 'pdf': 'layout_pdf_cb',
                  'schedule': 'layout_schedule_cb', 'drafting': 'layout_drafting_cb',
                  'legend': 'layout_legend_cb'}
        cb = getattr(self, cb_map.get(output_type, ''), None)
        label = str(cb.SelectedItem) if cb and cb.SelectedItem else ''
        if not label or label == '(none)':
            return None
        path = self._layout_template_map().get(label)
        if not path:
            # Legacy assignment, or a template deleted since it was chosen.
            path = os.path.join(self._layouts_dir(), label + '.json')
        return path if os.path.isfile(path) else None

    def _current_meta_rows(self):
        """The issue metadata currently selected in this window.

        Same (label, value) pairs the export payload carries as 'meta_rows',
        read straight from the combos. Used to seed pyTransmit Studio's
        preview; every lookup is guarded because a combo may not be built yet
        depending on which panel the user has opened.
        """
        rows = []

        def _code(combo_name, data_name, attr):
            try:
                combo = getattr(self, combo_name, None)
                data = list(getattr(self.opt_ctrl, data_name, []) or [])
                idx = combo.SelectedIndex if combo else 0
                if idx > 0 and idx - 1 < len(data):
                    return getattr(data[idx - 1], attr, '') or ''
            except Exception:
                pass
            return ''

        try:
            initials = getattr(self, 'initials_tb', None)
            if initials is not None and (initials.Text or '').strip():
                rows.append(('Issued By', initials.Text.strip()))
        except Exception:
            pass
        for label, combo_name, data_name, attr in (
                ('Reason for Issue', 'reason_cb', 'reason_data', 'Code'),
                ('Method of Issue', 'method_cb', 'method_data', 'Code'),
                ('Document Format', 'format_cb', 'format_data', 'Value'),
                ('Paper Size', 'printsize_cb', 'printsize_data', 'Value')):
            value = _code(combo_name, data_name, attr)
            if value:
                rows.append((label, value))
        return rows

    def is_studio_layout(self, path):
        """Is this a Studio template rather than a Layout Builder one?

        The two schemas are unmistakable - Layout Builder has 'rows', Studio
        has 'cells' - and each has its own writer, so the caller uses this to
        pick between them. Unreadable files answer False so the Layout
        Builder writer reports the read error, which it already does well.
        """
        if not path or not os.path.isfile(path):
            return False
        try:
            with open(path, 'r') as f:
                data = _json.load(f)
        except Exception:
            return False
        return 'cells' in data and 'rows' not in data

    # ── Import panel handlers ─────────────────────────────────────────────

    def import_browse_click(self, sender, args):
        if self.import_ctrl: self.import_ctrl.on_browse(sender, args)

    def import_execute_click(self, sender, args):
        if self.import_ctrl: self.import_ctrl.on_execute(sender, args)
        # Also import Layout templates if checkbox is checked
        try:
            cb = getattr(self, 'import_layouts_cb', None)
            if cb is None or cb.IsChecked:
                self._import_layouts()
        except Exception:
            pass

    def _import_layouts(self):
        """Copy Layouts/*.json from import folder into Layout/Layouts/."""
        try:
            import_path = ''
            tb = getattr(self, 'import_path_tb', None)
            if tb and tb.Text: import_path = tb.Text.strip()
            if not import_path or not os.path.isdir(import_path):
                return
            src_layouts = os.path.join(import_path, 'pyTransmit Settings', 'Layouts')
            if not os.path.isdir(src_layouts):
                return
            script_dir = os.path.dirname(os.path.abspath(__file__))
            dest_layouts = LAYOUTS_DIR
            if not os.path.isdir(dest_layouts):
                os.makedirs(dest_layouts)
            import shutil as _shutil
            for fn in os.listdir(src_layouts):
                if fn.lower().endswith('.json'):
                    _shutil.copy2(os.path.join(src_layouts, fn),
                                  os.path.join(dest_layouts, fn))
            # Also import layout_config.json
            src_cfg = os.path.join(import_path, 'pyTransmit Settings', 'layout_config.json')
            if os.path.isfile(src_cfg):
                _shutil.copy2(src_cfg, LAYOUT_CONFIG)
        except Exception:
            pass

    def import_settings_back_click(self, sender, args):
        if self.import_ctrl: self.import_ctrl.save_config()
        self._show_panel("main")

    def _update_log_menu_label(self):
        """Flip the log menu label to show whether logging is on or off."""
        try:
            lbl = self.FindName('log_menu_label')
            if lbl:
                lbl.Text = 'Log: ON  (click to cancel)' if self._log_enabled else 'Enable Log'
        except Exception:
            pass

    def _save_log_setting(self):
        """Write log_zip_path and log_folder into pytransmit_setup.json,
        leaving all other keys intact."""
        try:
            import json as _lj
            _lcfg_path = settings_file('pytransmit_setup.json')
            try:
                with open(_lcfg_path, 'r') as _lf:
                    _lcfg = _lj.load(_lf)
            except Exception:
                _lcfg = {}
            _lcfg['log_zip_path'] = self._log_zip_path
            _lcfg['log_folder'] = os.path.dirname(self._log_zip_path)
            with open(_lcfg_path, 'w') as _lf:
                _lj.dump(_lcfg, _lf, indent=2)
        except Exception:
            pass

    def _load_log_folder(self):
        """Last folder the user saved a log to, or None if never set."""
        try:
            import json as _lj
            _lcfg_path = settings_file('pytransmit_setup.json')
            with open(_lcfg_path, 'r') as _lf:
                return _lj.load(_lf).get('log_folder')
        except Exception:
            return None

    def menu_log_click(self, sender, args):
        """☰ Enable Log: ask where to save, turn on, or cancel if already on."""
        self.OptionsPopup.IsOpen   = False
        self.options_btn.IsChecked = False

        if self._log_enabled:
            self._log_enabled = False
            self._update_log_menu_label()
            return

        try:
            import datetime as _ldt
            _now = _ldt.datetime.now()
            _filename = 'pyTransmit Log Report {} {}.zip'.format(
                _now.strftime('%Y-%m-%d'), _now.strftime('%H-%M-%S'))
            _default_folder = self._load_log_folder() or os.path.expanduser("~\\Desktop")
            from Dialogs import Dialogs as _D
            _folder = _D.save_log(_default_folder, _filename)
            if not _folder:
                return
            self._log_zip_path = os.path.join(_folder, _filename)
            self._log_enabled  = True
            self._save_log_setting()
            self._update_log_menu_label()
        except Exception as e:
            _alert('Could not set log path: ' + str(e))

    def menu_help_click(self, sender, args):
        """☰ → Email support: open a pre-filled support email in the default
        mail client, addressed to Seed43 support, with the extension version
        and which app it came from already filled in."""
        from Snippets._support import support_mailto
        self.OptionsPopup.IsOpen   = False
        self.options_btn.IsChecked = False
        self._open_url(support_mailto("pyTransmit", _SCRIPT_DIR_MAIN),
                       title="Support")

    def run_excel_export(self):
        """Run Excel export, loads script_excel.py in a clean namespace."""
        try:
            # Read export path from the Export Settings panel (export_path_tb)
            export_path = r'C:\Temp'
            tb = getattr(self, 'export_path_tb', None)
            if tb and tb.Text:
                export_path = tb.Text
            if not export_path or not os.path.exists(export_path):
                _alert("Excel export path does not exist:\n{}\n\nSet the path in  ☰ → Export Settings.".format(export_path))
                return

            script_dir   = os.path.dirname(os.path.abspath(__file__))
            excel_script = os.path.join(script_dir, "script_excel.py")
            if not os.path.exists(excel_script):
                _alert("script_excel.py not found in:\n{}".format(script_dir))
                return

            # Pass export path and grouping params via environment variables
            os.environ['PYTRANSMIT_EXCEL_PATH'] = export_path
            group_params = getattr(self, 'selected_params', []) or []
            os.environ['PYTRANSMIT_GROUP_PARAMS'] = ','.join(group_params)
            os.environ['PYTRANSMIT_GROUP_LABEL'] = '1' if getattr(self, 'group_label_on', True) else '0'

            # Pass method/format/printsize from current UI selections
            try:
                opt = self.opt_ctrl
                cb_method = getattr(self, 'method_cb', None)
                cb_format = getattr(self, 'format_cb', None)
                cb_print  = getattr(self, 'printsize_cb', None)
                if opt and cb_method and cb_method.SelectedIndex > 0:
                    rows = list(opt.method_data)
                    idx  = cb_method.SelectedIndex - 1
                    if idx < len(rows):
                        os.environ['PYTRANSMIT_METHOD'] = \
                            getattr(rows[idx], 'Code', '') or ''
                if opt and cb_format and cb_format.SelectedIndex > 0:
                    rows = list(opt.format_data)
                    idx  = cb_format.SelectedIndex - 1
                    if idx < len(rows):
                        os.environ['PYTRANSMIT_FORMAT'] = \
                            getattr(rows[idx], 'Value', '') or ''
                if opt and cb_print and cb_print.SelectedIndex > 0:
                    rows = list(opt.printsize_data)
                    idx  = cb_print.SelectedIndex - 1
                    if idx < len(rows):
                        os.environ['PYTRANSMIT_PRINTSIZE'] = \
                            getattr(rows[idx], 'Value', '') or ''
            except:
                pass

            # Execute in isolated namespace, __name__ != '__main__' so no
            # entry-point guards fire; IronPython 2 compatible (no compile())
            ns = {'__name__': 'excel_export', '__file__': excel_script,
                  '__builtins__': __builtins__}
            with open(excel_script, 'r') as f:
                src = f.read()
            exec(src, ns)

        except Exception as e:
            _alert("Error exporting to Excel:\n{}".format(str(e)))
    
    def _show_file_save_dialog(self, title, filename, ext_label, ext, initial_folder=None):
        from Dialogs import Dialogs as _D
        return _D.file_save(title, filename, ext, initial_folder)

    def _show_open_file_dialog(self, title, message):
        from Dialogs import Dialogs as _D
        return _D.open_file(title, message)

    def run_revit_export(self):
        """Run Revit export, dispatches to the correct Publish script based on output_type."""
        try:
            script_dir   = os.path.dirname(os.path.abspath(__file__))
            publish_dir  = os.path.join(script_dir, 'Publish')
            settings_dir = os.path.join(script_dir, 'Settings')
            if settings_dir not in sys.path:
                sys.path.insert(0, settings_dir)

            # ── Determine output types from Setup panel checkboxes ────────────
            cfg = self.setup_ctrl.cfg if self.setup_ctrl else {}

            output_types = []
            try:
                if getattr(self, 'setup_output_schedule_cb', None) and self.setup_output_schedule_cb.IsChecked:
                    output_types.append('schedule')
                if getattr(self, 'setup_output_excel_cb', None) and self.setup_output_excel_cb.IsChecked:
                    output_types.append('excel')
                if getattr(self, 'setup_output_pdf_cb', None) and self.setup_output_pdf_cb.IsChecked:
                    output_types.append('pdf')
                if getattr(self, 'setup_output_drafting_cb', None) and self.setup_output_drafting_cb.IsChecked:
                    output_types.append('drafting')
                if getattr(self, 'setup_output_legend_cb', None) and self.setup_output_legend_cb.IsChecked:
                    output_types.append('legend')
                if not output_types:
                    output_types = ['schedule']
            except:
                output_types = ['schedule']

            # ── Build page height config ───────────────────────────────────────
            page_height_mode = cfg.get('page_height_mode', 'a4')
            page_height_mm   = cfg.get('page_height_mm', 287)
            try:
                if getattr(self, 'setup_height_none_rb', None) and self.setup_height_none_rb.IsChecked:
                    page_height_mode = 'none'
                elif getattr(self, 'setup_height_custom_rb', None) and self.setup_height_custom_rb.IsChecked:
                    page_height_mode = 'custom'
                    raw = getattr(self, 'setup_page_height_tb', None)
                    if raw:
                        page_height_mm = int(float(raw.Text or '287'))
                elif getattr(self, 'setup_height_a4_rb', None) and self.setup_height_a4_rb.IsChecked:
                    page_height_mode = 'a4'
                    page_height_mm   = 287
            except:
                pass

            # ── Build meta rows from enabled Setup fields ──────────────────────
            # If no revision is being issued, fall back to last issued revision's stored data
            _last_issued = None
            try:
                _all_revs = sorted(
                    revit.query.get_elements_by_class(DB.Revision, doc=self.doc),
                    key=lambda r: r.SequenceNumber)
                _issued = [r for r in _all_revs if r.Issued]
                if _issued: _last_issued = _issued[-1]
            except: pass

            def _parse_tag(ito, tag):
                try:
                    import re as _re2
                    m = _re2.search(r'\b' + tag + r':([^\s|]+)', ito or '')
                    return m.group(1).strip() if m else ''
                except: return ''

            _ito = (_last_issued.IssuedTo or '') if _last_issued else ''
            _iby = ((_last_issued.IssuedBy or '').strip()) if _last_issued else ''

            meta_rows = []
            try:
                if cfg.get('show_initials', True):
                    initials_val = ''
                    initials_tb  = getattr(self, 'initials_tb', None)
                    if initials_tb:
                        initials_val = initials_tb.Text or ''
                    if not initials_val:
                        initials_val = _iby
                    meta_rows.append(('Issued By', initials_val))
            except:
                pass
            try:
                if cfg.get('show_reason', True):
                    idx  = self.reason_cb.SelectedIndex
                    rows = list(self.opt_ctrl.reason_data)
                    code = getattr(rows[idx - 1], 'Code', '') if idx > 0 and idx - 1 < len(rows) else ''
                    if not code: code = _parse_tag(_ito, 'R')
                    meta_rows.append(('Reason for Issue', code))
            except:
                pass
            try:
                if cfg.get('show_method', True):
                    idx  = self.method_cb.SelectedIndex
                    rows = list(self.opt_ctrl.method_data)
                    code = getattr(rows[idx - 1], 'Code', '') if idx > 0 and idx - 1 < len(rows) else ''
                    if not code: code = _parse_tag(_ito, 'M')
                    meta_rows.append(('Method of Issue', code))
            except:
                pass
            try:
                if cfg.get('show_format', True):
                    idx = self.format_cb.SelectedIndex
                    val = ''
                    if idx > 0:
                        rows = list(self.opt_ctrl.format_data)
                        val  = getattr(rows[idx - 1], 'Value', '') if idx - 1 < len(rows) else ''
                    if not val: val = _parse_tag(_ito, 'F')
                    meta_rows.append(('Document Format', val))
            except:
                pass
            try:
                if cfg.get('show_printsize', True):
                    idx = self.printsize_cb.SelectedIndex
                    val = ''
                    if idx > 0:
                        rows = list(self.opt_ctrl.printsize_data)
                        val  = getattr(rows[idx - 1], 'Value', '') if idx - 1 < len(rows) else ''
                    if not val: val = _parse_tag(_ito, 'S')
                    meta_rows.append(('Paper Size', val))
            except:
                pass

            # ── Build legend strings from live OptionsSettings data ────────────
            reason_legend = ''
            method_legend = ''
            try:
                reason_lines = []
                for r in self.opt_ctrl.reason_data:
                    _code = getattr(r, 'Code', '') or ''
                    _sep  = getattr(r, 'Separator', '') or ''
                    _desc = getattr(r, 'Description', '') or ''
                    if _sep:
                        reason_lines.append('{} {} {}'.format(_code, _sep, _desc).strip())
                    else:
                        reason_lines.append('{} {}'.format(_code, _desc).strip())
                reason_legend = '\n'.join(l for l in reason_lines if l)
            except:
                pass
            try:
                method_lines = []
                for r in self.opt_ctrl.method_data:
                    _code = getattr(r, 'Code', '') or ''
                    _sep  = getattr(r, 'Separator', '') or ''
                    _desc = getattr(r, 'Description', '') or ''
                    if _sep:
                        method_lines.append('{} {} {}'.format(_code, _sep, _desc).strip())
                    else:
                        method_lines.append('{} {}'.format(_code, _desc).strip())
                method_legend = '\n'.join(l for l in method_lines if l)
            except:
                pass

            # ── Build recipients list from active mode ─────────────────────────
            mode = cfg.get('recipient_mode', 'dist')
            recipients = []
            try:
                if mode == 'dist':
                    for r in getattr(self, '_dist_rows', []):
                        recipients.append({
                            'label':  r['label'],
                            'attn':   r['attn_tb'].Text   or '',
                            'copies': r['copies_tb'].Text or '',
                        })
                else:
                    for r in getattr(self, '_client_rows', []):
                        comp_cb   = r.get('company_cb')
                        cont_cb   = r.get('contact_cb')
                        copies_tb = r.get('copies_tb')
                        if comp_cb is None or comp_cb.SelectedIndex <= 0:
                            continue
                        company = str(comp_cb.SelectedItem or '').strip()
                        attn    = (str(cont_cb.SelectedItem or '').strip()
                                   if cont_cb and cont_cb.SelectedIndex > 0 else '')
                        copies  = copies_tb.Text.strip() if copies_tb else ''
                        if company:
                            recipients.append({'label': company, 'attn': attn, 'copies': copies})
            except:
                pass

            # If recipients empty, fall back to parsing the last issued revision's IssuedTo
            if not recipients or (mode == 'dist' and not any(r.get('attn') or r.get('copies') for r in recipients)):
                try:
                    import re as _re3
                    # Extract DL: or CL: block, or fall back to second pipe-block
                    _dl_m = _re3.search(r'DL:\s*(.*?)(?:\s*\|[^|]|$)', _ito)
                    _cl_m = _re3.search(r'CL:\s*(.*?)(?:\s*\|[^|]|$)', _ito)
                    if _dl_m:
                        _recip_block = _dl_m.group(1).strip()
                    elif _cl_m:
                        _recip_block = _cl_m.group(1).strip()
                    elif ' | ' in _ito:
                        _recip_block = _ito.split(' | ', 1)[1].strip()
                    else:
                        _recip_block = ''

                    if _recip_block:
                        if mode == 'dist':
                            # New format: 1A.[attn]copies, old format: A.[attn]copies
                            _tokens = _re3.findall(r'(\d*)([A-Za-z]+)\.\[([^\]]*)\](\d*)', _recip_block)
                            _imap = {}  # index → (attn, copies)
                            _lmap = {}  # letter → (attn, copies)
                            for _num, _lets, _attn, _copies in _tokens:
                                if _num: _imap[int(_num)] = (_attn, _copies)
                                _lmap[_lets[0].upper()] = (_attn, _copies)
                            for _ri, r in enumerate(recipients):
                                if (_ri + 1) in _imap:
                                    r['attn'], r['copies'] = _imap[_ri + 1]
                                elif (r.get('label', '') or '')[:1].upper() in _lmap:
                                    r['attn'], r['copies'] = _lmap[(r.get('label', '') or '')[:1].upper()]
                        else:
                            # client format: [Company, Contact]copies
                            _tokens = _re3.findall(r'\[([^\]]+)\](\d*)', _recip_block)
                            recipients = []
                            for _label_full, _copies in _tokens:
                                if u'\u2014' in _label_full:
                                    _parts = _label_full.split(u'\u2014', 1)
                                    _company = _parts[0].strip()
                                    _attn    = _parts[1].strip()
                                else:
                                    _company = _label_full.strip()
                                    _attn    = ''
                                if _company:
                                    recipients.append({'label': _company, 'attn': _attn, 'copies': _copies})
                except:
                    pass

            # ── Assemble payload ───────────────────────────────────────────────
            # Branding values come from BrandingSettingsController
            _bc = getattr(self, 'brand_ctrl', None)
            payload = {
                'page_height_mode':  page_height_mode,
                'page_height_mm':    page_height_mm,
                'meta_rows':         meta_rows,
                'reason_legend':     reason_legend,
                'method_legend':     method_legend,
                'recipients':        recipients,
                'group_params':      getattr(self, 'selected_params', []) or [],
                'group_label':       getattr(self, 'group_label_on', False),
                'logo_path':         _bc.get_logo_path()       if _bc else '',
                'title_bg_color':    _bc.get_title_bg_color()  if _bc else '#FFFFFF',
                'title_fg_color':    _bc.get_title_fg_color()  if _bc else '#000000',
                'header_bg_color':   _bc.get_header_bg_color() if _bc else '#FFFFFF',
                'header_fg_color':   _bc.get_header_fg_color() if _bc else '#000000',
                # The USER's Settings folder, not the code one above: the only
                # consumer is script_create_schedule.py, which reads
                # reason.json/method.json from it. legend and pdf resolve the
                # code folder themselves, for sys.path.
                '_settings_dir':     SETTINGS_DIR,
                # So the Publish scripts' convention-search fallback
                # looks in .user too, not just beside the tool.
                '_layouts_dir':      LAYOUTS_DIR,
                '_user_dir':         USER_DIR,
                'script_dir':        script_dir,
                'rev_numbering_type': getattr(self, '_rev_numbering_type', ''),
                '_open_file_dialog':  self._show_open_file_dialog,
            }

            # ── Pre-collect save paths for file outputs ────────────────────────
            # Ask for paths upfront. Cancelling one skips only that output.
            import re as _re
            import datetime as _dt

            # Resolve project info
            try:
                _pi        = revit.doc.ProjectInformation
                _proj_num  = (_pi.get_Parameter(DB.BuiltInParameter.PROJECT_NUMBER).AsString() or '') if _pi else ''
                _proj_name = (_pi.get_Parameter(DB.BuiltInParameter.PROJECT_NAME).AsString() or revit.doc.Title or '') if _pi else revit.doc.Title or ''
            except Exception:
                _proj_num  = ''
                _proj_name = ''

            # Resolve output folder via filenaming_ctrl
            _fn_ctrl       = getattr(self, 'filenaming_ctrl', None)
            _output_folder = None
            try:
                if _fn_ctrl and _proj_num:
                    # Resolve full folder name on disk (e.g. '4285 - Alterations...')
                    _job_folder_name = _proj_num
                    try:
                        from FileNamingSettings import find_project_folder
                        _roots = [r for r in [_fn_ctrl._projects_root,
                                              _fn_ctrl._projects_older_root] if r]
                        _found = find_project_folder(_proj_num, _roots)
                        if _found:
                            import os as _os
                            _job_folder_name = _os.path.basename(_found)
                    except Exception:
                        pass
                    _resolved = _fn_ctrl.resolve_output_path(
                        _proj_num,
                        values={'job_number': _job_folder_name,
                                'proj_number': _job_folder_name}
                    )
                    if _resolved and os.path.isdir(_resolved):
                        _output_folder = _resolved
            except Exception:
                pass

            # Resolve filename from transmittal naming template
            try:
                _today  = _dt.date.today()
                _tmpl   = _fn_ctrl.get_template() if _fn_ctrl else ''
                _subs   = {
                    '{proj_number}':  _proj_num,
                    '{proj_name}':    _proj_name,
                    '{current_date}': _today.strftime('%Y-%m-%d'),
                    '{issue_date}':   _today.strftime('%Y-%m-%d'),
                    '{date_cc}':      _today.strftime('%Y')[:2],
                    '{date_yy}':      _today.strftime('%y'),
                    '{date_mm}':      _today.strftime('%m'),
                    '{date_dd}':      _today.strftime('%d'),
                }
                for _k, _v in _subs.items():
                    _tmpl = _tmpl.replace(_k, _v)
                _tmpl      = re.sub(r'\{[^}]+\}', '', _tmpl).strip('_- ')
                _safe_base = re.sub(r'[\\/*?:"<>|]', '_', _tmpl)[:80] if _tmpl else ''
            except Exception:
                _safe_base = ''
            if not _safe_base:
                _safe_base = re.sub(r'[\\/*?:"<>|]', '_',
                    'Document_Transmittal_{}_{}'.format(_proj_num, _proj_name))[:60]

            for output_type in output_types:
                if output_type == 'excel' and not payload.get('_pdf_temp_xlsx_path'):
                    _path = self._show_file_save_dialog(
                        title          = u'Save Transmittal — Excel',
                        filename       = '{}.xlsx'.format(_safe_base),
                        ext_label      = 'Excel Workbook (*.xlsx)',
                        ext            = 'xlsx',
                        initial_folder = _output_folder,
                    )
                    payload['_excel_save_path'] = _path  # None = skip this output
                elif output_type == 'pdf':
                    _path = self._show_file_save_dialog(
                        title          = u'Save Transmittal — PDF',
                        filename       = '{}.pdf'.format(_safe_base),
                        ext_label      = 'PDF File (*.pdf)',
                        ext            = 'pdf',
                        initial_folder = _output_folder,
                    )
                    payload['_pdf_save_path'] = _path  # None = skip this output

            # ── Write ptransmit_rev parameter before export ──────────────────
            try:
                import imp as _wrev_imp, os as _wrev_os
                _wrev_path = _wrev_os.path.join(publish_dir, 'write_rev_param.py')
                if _wrev_os.path.isfile(_wrev_path):
                    _wrev = _wrev_imp.load_source('write_rev_param', _wrev_path)
                    _wrev.write_rev_param(revit.doc, publish_dir)
            except Exception as _wrev_err:
                import traceback as _wrev_tb
                _alert('ptransmit_rev write failed: ' + str(_wrev_err))

            # ── Set up log capture if logging is enabled ──────────────────────
            _log_lines        = []
            _log_layout_paths = []
            _log_enabled      = getattr(self, '_log_enabled', False)
            _log_zip_path     = getattr(self, '_log_zip_path', '')
            payload['_log_lines']   = _log_lines
            payload['_log_enabled'] = _log_enabled
            payload['output_types'] = output_types

            # ── Dispatch each selected output type ────────────────────────────
            for output_type in output_types:

                # Each layout schema has its own writer, chosen from the
                # assigned template rather than from a setting - the two
                # cannot read each other's files, so the template itself is
                # the only honest thing to decide on.
                _assigned = self.get_layout_for_output(output_type)
                _is_studio = self.is_studio_layout(_assigned)

                if output_type == 'excel':
                    target_script = os.path.join(
                        publish_dir, 'script_create_excel_studio.py' if _is_studio
                        else 'script_create_excel.py')
                    script_name   = 'excel_export'
                    err_label     = 'Excel script'
                elif output_type == 'pdf':
                    target_script = os.path.join(publish_dir, 'script_create_pdf.py')
                    script_name   = 'pdf_export'
                    err_label     = 'PDF script'
                elif output_type == 'drafting':
                    target_script = os.path.join(publish_dir, 'script_create_drafting_view.py')
                    script_name   = 'drafting_export'
                    err_label     = 'Drafting View script'
                elif output_type == 'legend':
                    target_script = os.path.join(publish_dir, 'script_create_legend.py')
                    script_name   = 'legend_export'
                    err_label     = 'Legend script'
                else:
                    target_script = os.path.join(publish_dir, 'script_create_schedule.py')
                    script_name   = 'revit_schedule_export'
                    err_label     = 'Schedule script'

                if not os.path.exists(target_script):
                    _alert("{} not found at:\n{}".format(err_label, target_script))
                    _log_lines.append('[SKIP] {} not found: {}'.format(err_label, target_script))
                    continue

                payload['output_type']      = output_type
                payload['layout_json_path'] = _assigned
                # PDF builds its workbook by running the Excel writer, so it
                # has to be told which one to run.
                payload['_layout_is_studio'] = _is_studio
                ns = {
                    '__name__':    script_name,
                    '__file__':    target_script,
                    '__builtins__': __builtins__,
                    'PYTRANSMIT_PAYLOAD': payload,
                }
                with open(target_script, 'r') as f:
                    src = f.read()
                try:
                    # Hide the output window before running so it never appears
                    try:
                        _out = script.get_output()
                        if _out: _out.hide()
                    except Exception: pass
                    _log_lines.append('[START] {}'.format(output_type.upper()))
                    exec(src, ns)
                    _log_lines.append('[DONE]  {}'.format(output_type.upper()))
                    _ljp = payload.get('layout_json_path', '')
                    if _ljp and _ljp not in _log_layout_paths:
                        _log_layout_paths.append(_ljp)
                except SystemExit:
                    # User cancelled inside the script, skip this output only
                    _log_lines.append('[CANCELLED] {}'.format(output_type.upper()))
                    continue
                except Exception as exec_e:
                    import traceback as _tb
                    tb_str = _tb.format_exc() or str(exec_e) or repr(exec_e)
                    _log_lines.append('[ERROR] {} : {}'.format(output_type.upper(), tb_str))
                    _alert("Error running {}:\n{}".format(err_label, tb_str))

            # ── Build and write log if enabled ────────────────────────────────
            if _log_enabled and _log_zip_path:
                try:
                    import imp as _limp
                    _log_script = os.path.join(publish_dir, 'script_create_log.py')
                    _lmod = _limp.load_source('script_create_log', _log_script)
                    # Add project info to payload for log
                    payload['proj_number'] = _proj_num
                    payload['proj_name']   = _proj_name
                    _ok, _result = _lmod.build_log(
                        payload, _log_lines, _log_zip_path, _log_layout_paths)
                    if _ok:
                        _alert('Log saved to:\n{}'.format(_result))
                    else:
                        _alert('Log could not be saved:\n{}'.format(_result))
                except Exception as _le:
                    _alert('Log error:\n{}'.format(str(_le)))
                finally:
                    # Always turn log off after the run
                    self._log_enabled = False
                    self._update_log_menu_label()

        except Exception as e:
            import traceback
            tb_str = traceback.format_exc() or str(e) or repr(e)
            _alert("Error exporting Revit data:\n{}".format(tb_str))

# --- generate_tables removed - this script only updates revision data ---

# --- MAIN EXECUTION ---
def main():
    try:
        window = RevTableWindow()
        window.ShowDialog()
    except Exception as ex:
        _alert("Error initializing window: {}".format(str(ex)), exitscript=True)

if __name__ == "__main__":
    main()
