# -*- coding: utf-8 -*-
__title__  = "pyTransmit"
__author__ = "Nagel Consultants"
__doc__    = """
VERSION 250507
_____________________________________________________________________
Description:
Main window for pyTransmit. Collects revision and issue details,
manages recipients, options, layout and export settings, then
publishes the transmittal document to the selected output formats.

_____________________________________________________________________
How-to:
1. Fill in the revision details and select the issue reason and method.
2. Choose who the transmittal is being sent to under Recipients.
3. Click Publish to generate the selected output formats.

_____________________________________________________________________
Notes:
Output formats and their layout templates are configured under
Settings, then Layout. Each format can be assigned its own layout
independently.

_____________________________________________________________________
Last update:
250507 - Added group label toggle for sheet grouping. Fixed group
boundary borders in Excel and Schedule exporters. Custom page size
now saves correctly per template.
_____________________________________________________________________
"""

from pyrevit import revit, forms, script, DB
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, XYZ, Line, TextNote, TextNoteType, Transaction, CurveElement,
    ViewFamilyType, ViewFamily, ViewDrafting, TextNoteOptions, HorizontalTextAlignment, Color, GraphicsStyle, Category,
    OverrideGraphicSettings, ImageType, ImageTypeOptions, ImageInstance, ImageTypeSource, ImagePlacementOptions, BoxPlacement
)
import math
import re
from itertools import groupby
from pyrevit.forms import WPFWindow  # kept for other uses
import wpf
from System.Windows import Window, ResourceDictionary
from System import Uri
import clr
clr.AddReference("PresentationFramework")
from System.Windows.Controls import ComboBox
from System.Windows import Thickness
import System.Windows
import System.Windows.Media
import os
import sys

_SCRIPT_DIR_MAIN = os.path.dirname(os.path.abspath(__file__))

def _find_seed43_styles():
    """Walk up from the pushbutton folder to find Seed43Styles.xaml. Returns path or None."""
    folder = _SCRIPT_DIR_MAIN
    for _ in range(6):
        candidate = os.path.join(folder, 'Seed43Styles.xaml')
        if os.path.isfile(candidate):
            return candidate
        folder = os.path.dirname(folder)
    return None

# ── EXTERNAL URLS ───────────────────────────────────────────────────────────────
# Update these to change where Help and About point
HELP_URL  = "https://seed43.org/your-thoughts/"
ABOUT_URL = "https://seed43.org/"

# ── AUTO-UPDATE CHECK ────────────────────────────────────────────────────────────
import json as _json

_PYTRANSMIT_CHANGELOG_URL = (
    "https://raw.githubusercontent.com/Seed-43/Seed43-PyTransmit/main/bundle.yaml"
)
_PYTRANSMIT_ZIP_URL = (
    "https://github.com/Seed-43/Seed43-PyTransmit/archive/refs/heads/main.zip"
)
_SCRIPT_DIR         = os.path.dirname(os.path.abspath(__file__))
_YAML_FILE          = os.path.join(_SCRIPT_DIR, "bundle.yaml")


def _pt_parse_yaml(path):
    """Parse a simple yaml file into a dict."""
    result = {}
    changelog = []
    in_changelog = False
    try:
        with open(path, "r") as f:
            for line in f.read().splitlines():
                stripped = line.strip()
                if not stripped or stripped.startswith("#"):
                    continue
                if stripped.startswith("- ") and in_changelog:
                    changelog.append(stripped[2:].strip())
                    continue
                if ":" in stripped:
                    in_changelog = False
                    key, _, val = stripped.partition(":")
                    key = key.strip()
                    val = val.strip()
                    if key == "changelog":
                        in_changelog = True
                    elif val:
                        result[key] = val
        if changelog:
            result["changelog"] = changelog
    except Exception:
        pass
    return result


def _pt_write_version_to_yaml(version):
    """Write installed version into local pytransmit.yaml version field."""
    try:
        if not os.path.exists(_YAML_FILE):
            with open(_YAML_FILE, "w") as f:
                f.write("version: {}\n".format(version))
            return
        with open(_YAML_FILE, "r") as f:
            lines = f.readlines()
        new_lines = []
        replaced = False
        for line in lines:
            if line.strip().startswith("version:") and not replaced:
                new_lines.append("version: {}\n".format(version))
                replaced = True
            else:
                new_lines.append(line)
        if not replaced:
            new_lines.append("version: {}\n".format(version))
        with open(_YAML_FILE, "w") as f:
            f.writelines(new_lines)
    except Exception:
        pass


def _pt_get_local_version():
    data = _pt_parse_yaml(_YAML_FILE)
    return data.get("version", None)


def _pt_fetch_remote_version():
    try:
        import time as _time
        from System.Net import WebClient
        _bust = int(_time.time())
        _url  = "{}?cb={}".format(_PYTRANSMIT_CHANGELOG_URL, _bust)
        wc = WebClient()
        wc.Headers.Add("Cache-Control", "no-cache, no-store, must-revalidate")
        wc.Headers.Add("Pragma", "no-cache")
        wc.Headers.Add("User-Agent", "pyTransmit-UpdateCheck")
        raw = wc.DownloadString(_url)
        result = {}
        changelog = []
        in_changelog = False
        for line in raw.splitlines():
            stripped = line.strip()
            if not stripped or stripped.startswith("#"):
                continue
            if stripped.startswith("- ") and in_changelog:
                changelog.append(stripped[2:].strip())
                continue
            if ":" in stripped:
                in_changelog = False
                key, _, val = stripped.partition(":")
                key = key.strip()
                val = val.strip()
                if key == "changelog":
                    in_changelog = True
                elif val:
                    result[key] = val
        return result.get("version", ""), changelog
    except Exception:
        return None, []


def _pt_do_update(remote_ver):
    """Download and install the latest PyTransmit files."""
    import shutil as _sh
    import zipfile as _zf
    try:
        tmp_zip = os.path.join(os.environ.get("TEMP", ""), "pytransmit_update.zip")
        tmp_dir = os.path.join(os.environ.get("TEMP", ""), "pytransmit_update_extracted")

        from System.Net import WebClient
        wc = WebClient()
        wc.Headers.Add("Cache-Control", "no-cache, no-store")
        wc.DownloadFile(_PYTRANSMIT_ZIP_URL, tmp_zip)

        if os.path.exists(tmp_dir):
            _sh.rmtree(tmp_dir)
        os.makedirs(tmp_dir)
        with _zf.ZipFile(tmp_zip, "r") as z:
            z.extractall(tmp_dir)

        # Find extracted root
        extracted_root = None
        for item in os.listdir(tmp_dir):
            full = os.path.join(tmp_dir, item)
            if os.path.isdir(full):
                extracted_root = full
                break
        if not extracted_root:
            raise Exception("Could not find extracted folder.")

        # Smart sync — skip user JSON and image files (except icon.png), delete removed files
        SKIP_EXT     = {".json", ".png", ".jpg", ".jpeg"}
        ALWAYS_UPDATE = {"icon.png"}

        def sync_tree(src_dir, dst_dir):
            if not os.path.exists(dst_dir):
                os.makedirs(dst_dir)
            src_items = set(os.listdir(src_dir))
            dst_items = set(os.listdir(dst_dir)) if os.path.exists(dst_dir) else set()
            for item in src_items:
                s    = os.path.join(src_dir, item)
                d    = os.path.join(dst_dir, item)
                ext  = os.path.splitext(item)[1].lower()
                name = item.lower()
                if os.path.isdir(s):
                    sync_tree(s, d)
                else:
                    if ext in SKIP_EXT and name not in ALWAYS_UPDATE:
                        continue  # skip user data files
                    _sh.copy2(s, d)
            for item in dst_items - src_items:
                ext  = os.path.splitext(item)[1].lower()
                name = item.lower()
                if ext in SKIP_EXT and name not in ALWAYS_UPDATE:
                    continue  # never delete user data files
                d = os.path.join(dst_dir, item)
                if os.path.isdir(d):
                    _sh.rmtree(d)
                else:
                    os.remove(d)

        sync_tree(extracted_root, _SCRIPT_DIR)

        # Write version into local yaml file
        _pt_write_version_to_yaml(remote_ver)

        # Cleanup
        if os.path.exists(tmp_zip):
            os.remove(tmp_zip)
        if os.path.exists(tmp_dir):
            _sh.rmtree(tmp_dir)

        return True, None
    except Exception as ex:
        return False, str(ex)


def _pt_check_and_notify():
    """Check for update and show orange ribbon if found — called after window init."""
    try:
        local_ver            = _pt_get_local_version()
        remote_ver, changes  = _pt_fetch_remote_version()
        if not remote_ver:
            return
        if local_ver and remote_ver == local_ver:
            return
        # Store remote version for use in update handler
        _pt_check_and_notify._remote_ver = remote_ver
        _pt_check_and_notify._changes    = changes
        _pt_check_and_notify._local_ver  = local_ver or "unknown"
    except Exception:
        pass


# Run update check on startup
_pt_check_and_notify()

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
            styles_path = _find_seed43_styles()
            if styles_path:
                rd = ResourceDictionary()
                rd.Source = Uri(styles_path)
                self.Resources = rd
            xaml_path = os.path.join(_SCRIPT_DIR_MAIN, "pyTransmit.xaml")
            wpf.LoadComponent(self, xaml_path)
        except Exception, e:
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
            forms.alert("Failed to load pyTransmit.xaml:\n\n{}".format(msg), exitscript=True)

        # Initialise panel controllers (panels are now in pyTransmit.xaml directly)
        self._init_controllers()
        
        if not hasattr(self, 'execute_btn'):
            forms.alert("Button 'execute_btn' not found in XAML.", exitscript=True)
        
        try:
            self.execute_btn.Click += self.execute_btn_click
        except Exception, e:
            forms.alert("Failed to bind button Click events: {}".format(str(e)), exitscript=True)
        
        self.doc = revit.doc
        all_revs = list(FilteredElementCollector(self.doc).OfClass(DB.Revision))
        self.non_issued_revs = [rev for rev in all_revs if not rev.Issued]
        self.issued_revs = [rev for rev in all_revs if rev.Issued]
        self.issued_revs = sorted(self.issued_revs, key=lambda r: r.SequenceNumber)

        try:
            self.revision_cb.ItemsSource = ["{} - {}".format(rev.SequenceNumber, rev.Description) for rev in self.non_issued_revs]
        except Exception, e:
            forms.alert("Failed to populate revision ComboBox: {}".format(str(e)), exitscript=True)
        
        try:
            self.reason_cb.SelectedIndex = 0
        except:
            pass
        
        self.sheet_param_combos = [self.sheet_param_cb_1]
        self.selected_params = []
        self.param_counter = 1
        self.group_label_on = False
        self._setup_group_label_toggle()
        
        self.sheet_params = self.get_sheet_parameters()
        
        if not self.sheet_params:
            forms.alert("No suitable sheet parameters found.", exitscript=True)
        
        try:
            self.sheet_param_cb_1.ItemsSource = ["(None)"] + self.sheet_params
            self.sheet_param_cb_1.SelectionChanged += self.sheet_param_selection_changed
        except Exception, e:
            forms.alert("Failed to populate or bind sheet parameter ComboBox: {}".format(str(e)), exitscript=True)
        
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
        
    
    def _setup_group_label_toggle(self):
        """Build the on/off toggle switch for group label display."""
        import System.Windows.Media as _SWM
        import System
        _ON_COLOR  = _SWM.Color.FromRgb(0x20, 0x8A, 0x3C)
        _OFF_COLOR = _SWM.Color.FromRgb(0xA0, 0xAA, 0xBB)
        try:
            sw = self.group_label_toggle
            sw.Background = _SWM.SolidColorBrush(
                _ON_COLOR if self.group_label_on else _OFF_COLOR)
            knob                     = System.Windows.Controls.Border()
            knob.Width               = 16
            knob.Height              = 16
            knob.CornerRadius        = System.Windows.CornerRadius(8)
            knob.Background          = _SWM.Brushes.White
            knob.HorizontalAlignment = System.Windows.HorizontalAlignment.Left
            knob.Margin = System.Windows.Thickness(22, 2, 0, 2) if self.group_label_on else System.Windows.Thickness(2, 2, 0, 2)
            sw.Child = knob
            self._group_label_knob = knob
            sw.MouseLeftButtonUp += self._on_group_label_toggle
            # Set initial text colour
            self.group_label_toggle_tb.Foreground = _SWM.SolidColorBrush(
                _SWM.Color.FromRgb(0xF4, 0xFA, 0xFF))
        except Exception:
            pass

    def _on_group_label_toggle(self, sender, args):
        """Toggle group label on/off, updating knob position and colour directly."""
        import System.Windows.Media as _SWM
        import System
        _ON_COLOR  = _SWM.Color.FromRgb(0x20, 0x8A, 0x3C)
        _OFF_COLOR = _SWM.Color.FromRgb(0xA0, 0xAA, 0xBB)
        self.group_label_on = not self.group_label_on
        try:
            # Set knob position directly, no animation (reliable in IronPython 2)
            self._group_label_knob.Margin = (
                System.Windows.Thickness(22, 2, 0, 2) if self.group_label_on
                else System.Windows.Thickness(2, 2, 0, 2))
            self.group_label_toggle.Background = _SWM.SolidColorBrush(
                _ON_COLOR if self.group_label_on else _OFF_COLOR)
            self.group_label_toggle_tb.Text = "Text On" if self.group_label_on else "Text Off"
            self.group_label_toggle_tb.Foreground = _SWM.SolidColorBrush(
                _SWM.Color.FromRgb(0xF4, 0xFA, 0xFF) if self.group_label_on
                else _SWM.Color.FromRgb(0xA0, 0xAA, 0xBB))
        except Exception:
            pass

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
        for bip in built_in_params:
            param = sample_sheet.get_Parameter(bip)
            if param and param.StorageType == DB.StorageType.String:
                param_names.add(param.Definition.Name)
        for param in sample_sheet.GetOrderedParameters():
            if param.Definition and param.StorageType == DB.StorageType.String:
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
                new_combo.Style = self.FindResource("ModernComboBoxStyle")
            except Exception, e:
                print("Failed to apply ModernComboBoxStyle: {}".format(str(e)))
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
                    t = Transaction(self.doc, "Set Issued To Data and Mark Issued")
                    t.Start()
                    try:
                        selected_rev.IssuedTo = data_str
                        if initials_str:
                            selected_rev.IssuedBy = initials_str
                        selected_rev.Issued = True
                        t.Commit()
                    except Exception, e:
                        t.RollBack()
                        forms.alert("Failed to update revision data: {}".format(str(e)), exitscript=True)
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

        except Exception, e:
            forms.alert("Error in execute_btn_click: {}".format(str(e)), exitscript=True)

    def _build_issued_to_string(self):
        """
        Build the IssuedTo string written into the Revit revision record.

        Always records every enabled field and the full recipient list — no
        user-configurable toggles.  The schedule generator reads this string
        back to populate the transmittal, so it must be complete and consistent.

        Format:
          R:<code> M:<code> F:<value> S:<value> I:<initials> | <recipients>

        Recipients (Distribution List mode):
          A.[Attention To]<copies>  O.[...]<copies>  ...
          (first letter of the role label, attention to in brackets, copies integer)

        Recipients (Client List mode):
          [Company — Attention To]<copies>  ...

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

        # Issued By — written to rev.IssuedBy, NOT included in IssuedTo string
        _initials_val = ''
        try:
            if cfg.get('show_initials', True):
                tb = getattr(self, 'initials_tb', None)
                if tb and tb.Text:
                    _initials_val = tb.Text.strip()
        except:
            pass

        # Recipients — always record the full list from whichever mode is active
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
            else:  # client mode — new structure: {company_cb, contact_cb, copies_tb}
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

        # VIS tag — snapshot which info rows are visible.
        # Always written (even if empty) so the mismatch checker can distinguish
        # "all fields off" from "old revision with no VIS tag".
        _vis_parts = []
        if cfg.get('show_from',     True): _vis_parts.append('FR')
        if cfg.get('show_client',   True): _vis_parts.append('CL')
        if cfg.get('show_projno',   True): _vis_parts.append('PN')
        if cfg.get('show_projname', True): _vis_parts.append('PJ')
        parts.append('VIS:{}'.format(','.join(_vis_parts)))

        # EX tag — snapshot which export formats are enabled
        _ex_parts = []
        if cfg.get('out_schedule',  True):  _ex_parts.append('RS')
        if cfg.get('out_drafting',  False): _ex_parts.append('RD')
        if cfg.get('out_legend',    False): _ex_parts.append('RL')
        if cfg.get('out_excel',     False): _ex_parts.append('Excl')
        if cfg.get('out_pdf',       False): _ex_parts.append('PDF')
        if _ex_parts:
            parts.append('EX:{}'.format(','.join(_ex_parts)))

        # RPG tag — snapshot page break setting
        _phm = cfg.get('page_height_mode', 'a4')
        _phv = cfg.get('page_height_mm', 287)
        if _phm == 'none':
            parts.append('RPG:0')
        elif _phm == 'custom':
            parts.append('RPG:{}'.format(int(_phv)))
        else:
            parts.append('RPG:1')  # A4

        # GP tag — snapshot active sheet grouping parameters
        _gp_list = getattr(self, 'selected_params', None) or []
        if _gp_list:
            # Use ~~ as separator — safe, won't appear in Revit parameter names
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
    # PANEL CONTROLLERS  — panels are now inline in pyTransmit.xaml
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
            forms.alert("Failed to init RecipientManager:\n{}".format(str(ex)))
            self.rec_ctrl = None

        # ── Options Manager ────────────────────────────────────────────────
        try:
            from OptionsSettings import OptionsSettingsController
            self.opt_ctrl = OptionsSettingsController()
            self.opt_ctrl.attach(self)
        except Exception as ex:
            forms.alert("Failed to init OptionsManager:\n{}".format(str(ex)))
            self.opt_ctrl = None

        # ── Setup Settings ─────────────────────────────────────────────────
        try:
            from SetupSettings import SetupSettingsController
            self.setup_ctrl = SetupSettingsController(script_dir)
            self.setup_ctrl.attach(self)
            # load_and_apply() is called later, after full window init
        except Exception as ex:
            forms.alert("Failed to init SetupSettings:\n{}".format(str(ex)))
            self.setup_ctrl = None

        # ── Export Settings ────────────────────────────────────────────────
        try:
            from ExportSettings import ExportSettingsController
            self.export_ctrl = ExportSettingsController(script_dir)
            self.export_ctrl.attach(self)
        except Exception as ex:
            forms.alert("Failed to init ExportSettings:\n{}".format(str(ex)))
            self.export_ctrl = None

        # ── Import Settings ────────────────────────────────────────────────
        try:
            from ImportSettings import ImportSettingsController
            self.import_ctrl = ImportSettingsController(script_dir)
            self.import_ctrl.attach(self)
        except Exception as ex:
            forms.alert("Failed to init ImportSettings:\n{}".format(str(ex)))
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
            forms.alert("Failed to init BrandingSettings:\n{}".format(str(ex)))
            self.brand_ctrl = None

        try:
            from FileNamingSettings import FileNamingSettingsController
            self.filenaming_ctrl = FileNamingSettingsController(script_dir)
            self.filenaming_ctrl.attach(self)
            self.filenaming_ctrl.load_config()
        except Exception as ex:
            forms.alert("Failed to init FileNamingSettings:\n{}".format(str(ex)))
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
                  'FileNamingPanel',
                  'main_content',
                  'header_normal_btns', 'setup_close_btn', 'styling_close_btn',
                  'recipient_header_btns', 'options_header_btns',
                  'export_settings_header_btns', 'import_settings_header_btns',
                  'filenaming_header_btns',
                  'setup_header_lbl', 'recipient_header_lbl', 'options_header_lbl',
                  'export_settings_header_lbl', 'import_settings_header_lbl',
                  'styling_header_lbl', 'filenaming_header_lbl']:
            hide(n)

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

    # ── Back / close handlers ─────────────────────────────────────────────

    def _show_save_dialog(self, panel_label):
        """Show a themed Save / Discard dialog matching the pyTransmit XAML theme.
        Returns True if the user chose Save, False if Discard."""
        import clr
        clr.AddReference("PresentationFramework")
        clr.AddReference("PresentationCore")
        clr.AddReference("WindowsBase")
        import System.Windows.Markup as Markup

        xaml = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="" Width="340" SizeToContent="Height"
        WindowStyle="None" ResizeMode="NoResize"
        WindowStartupLocation="CenterScreen"
        Background="Transparent"
        FontFamily="Segoe UI"
        AllowsTransparency="True">
    <Border Background="#2B3340" CornerRadius="10" Margin="12" Padding="24,20,24,20">
        <Border.Effect>
            <DropShadowEffect Color="Black" Opacity="0.5" ShadowDepth="4" BlurRadius="16"/>
        </Border.Effect>
        <StackPanel>
            <!-- green accent bar -->
            <Border Background="#208A3C" Height="3" CornerRadius="2" Margin="0,0,0,16"/>
            <!-- title -->
            <TextBlock Text="Save Changes"
                       Foreground="#F4FAFF" FontSize="15" FontWeight="Bold"
                       Margin="0,0,0,8"/>
            <!-- message -->
            <TextBlock x:Name="msg_tb"
                       Foreground="#F4FAFF" FontSize="12" Opacity="0.85"
                       TextWrapping="Wrap" Margin="0,0,0,24"/>
            <!-- buttons -->
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="discard_btn" Content="Discard"
                        Foreground="#F4FAFF" FontSize="12" FontWeight="Bold"
                        BorderThickness="0" Padding="20,8" Margin="0,0,8,0"
                        Cursor="Hand">
                    <Button.Template>
                        <ControlTemplate TargetType="Button">
                            <Border x:Name="Bd" Background="#404553" CornerRadius="6"
                                    Padding="{TemplateBinding Padding}">
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Border>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter TargetName="Bd" Property="Background" Value="#4E5566"/>
                                </Trigger>
                                <Trigger Property="IsPressed" Value="True">
                                    <Setter TargetName="Bd" Property="Background" Value="#333B48"/>
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Button.Template>
                </Button>
                <Button x:Name="save_btn" Content="Save"
                        Foreground="#F4FAFF" FontSize="12" FontWeight="Bold"
                        BorderThickness="0" Padding="20,8"
                        Cursor="Hand">
                    <Button.Template>
                        <ControlTemplate TargetType="Button">
                            <Border x:Name="Bd" Background="#208A3C" CornerRadius="6"
                                    Padding="{TemplateBinding Padding}">
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
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
                    </Button.Template>
                </Button>
            </StackPanel>
        </StackPanel>
    </Border>
</Window>
"""

        dlg = Markup.XamlReader.Parse(xaml)
        dlg.FindName("msg_tb").Text = "Do you want to save your changes to {}?".format(panel_label)

        result = [False]

        def on_save(s, e):
            result[0] = True
            dlg.Close()

        def on_discard(s, e):
            result[0] = False
            dlg.Close()

        dlg.FindName("save_btn").Click    += on_save
        dlg.FindName("discard_btn").Click += on_discard

        dlg.ShowDialog()
        return result[0]

    def recipient_back_click(self, sender, args):
        try:
            if self.rec_ctrl:
                try:
                    save = self._show_save_dialog("Recipients")
                except Exception:
                    save = forms.alert("Save changes to Recipients?",
                                       title="Recipients", ok=False, yes=True, no=True)
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
                    save = forms.alert("Save changes to Options?",
                                       title="Options", ok=False, yes=True, no=True)
                if save:
                    self.opt_ctrl.save_all()
                    self._auto_export_if_enabled()
                else:
                    self.opt_ctrl.discard()
                self.opt_ctrl.clear_selections()
        except Exception:
            pass
        self._show_panel("main")

    # ── Header export buttons — delegate to controllers ───────────────────

    def recipient_export_click(self, sender, args):
        if self.rec_ctrl:
            self.rec_ctrl.export_data(sender, args)

    def options_export_click(self, sender, args):
        if self.opt_ctrl:
            self.opt_ctrl.export_data(sender, args)

    def menu_setup_click(self, sender, args):
        """☰ → Setup: open the Setup configuration panel."""
        self.OptionsPopup.IsOpen = False
        self.options_btn.IsChecked = False
        self._show_panel("setup")

    def setup_back_click(self, sender, args):
        """Setup X — save config, apply, return to main."""
        if self.setup_ctrl:
            self.setup_ctrl.save()
            self.setup_ctrl.apply()
        self._show_panel("main")

    # ── Styling / Branding panel ──────────────────────────────────────────────
    # All logic lives in Settings/BrandingSettings.py (BrandingSettingsController)

    def menu_styling_click(self, sender, args):
        """Legacy — Branding panel removed; redirects to Document Layout."""
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
            forms.alert(
                'Could not open Layout Builder:\n{}'.format(str(e)),
                title='Document Layout')

    def styling_back_click(self, sender, args):
        """Styling X — legacy handler kept for safety."""
        self._show_panel("main")

    def filenaming_back_click(self, sender, args):
        """File Naming X — auto-save and return to main."""
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
        """Radio button toggled — delegate to SetupSettingsController."""
        if self.setup_ctrl:
            self.setup_ctrl._on_mode_changed(sender, args)

    def setup_fields_changed(self, sender, args):
        """Checkbox/RadioButton toggled — delegate to SetupSettingsController, then handle local UI."""
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

    def menu_about_click(self, sender, args):
        """☰ → About: open ABOUT_URL in the default browser."""
        self.OptionsPopup.IsOpen   = False
        self.options_btn.IsChecked = False
        try:
            import subprocess
            subprocess.Popen(['cmd', '/c', 'start', '', ABOUT_URL])
        except Exception, e:
            forms.alert("Could not open browser:\n{}".format(str(e)), title="About")

    def close_about_click(self, sender, args):
        """Legacy close handler — modal no longer used but kept for safety."""
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

            if self.opt_ctrl:
                select_coded ('reason_cb',    self.opt_ctrl.reason_data,    reason_code)
                select_coded ('method_cb',    self.opt_ctrl.method_data,    method_code)
                select_simple('format_cb',    self.opt_ctrl.format_data,    format_val)
                select_simple('printsize_cb', self.opt_ctrl.printsize_data, size_val)

            # ── Distribution List rows — attn + copies ────────────────────
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
                        # Last revision was saved in client mode — skip dist row fill
                        _saved_as_client = True
                        break
                # Fallback: old format (no DL:/CL: prefix) — second pipe-block
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
            from Autodesk.Revit.DB import FilteredElementCollector
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

        # Parse VIS tag — if present use it directly.
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
                # pyTransmit revision issued with all info rows off — VIS tag was skipped
                _proj_vis = set()
            else:
                # Genuine old revision with no pyTransmit tags — assume all rows were on
                _proj_vis = {'FR', 'CL', 'PN', 'PJ'}

        # Parse EX tag — if absent infer from context (old revisions assumed schedule only)
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
            return  # All good — no mismatch

        # Show styled mismatch dialog
        _result = self._show_mismatch_dialog()

        if _result == 'update':
            # Permanently update Setup settings to match project snapshot
            self._apply_vis_ex_to_setup(_proj_vis, _proj_ex, permanent=True)
        elif _result == 'session':
            # Apply for this session only — don't save to disk
            self._apply_vis_ex_to_setup(_proj_vis, _proj_ex, permanent=False)
        # 'ignore' — do nothing

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

    def _show_mismatch_dialog(self):
        """
        Styled dialog matching the Save/Discard pattern.
        Returns: 'update' | 'session' | 'ignore'
        """
        try:
            import System.Windows.Markup as Markup
            xaml = (
                '<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"'
                ' xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"'
                ' Title="" Width="420" SizeToContent="Height"'
                ' WindowStyle="None" ResizeMode="NoResize"'
                ' WindowStartupLocation="CenterScreen"'
                ' Background="Transparent" FontFamily="Segoe UI" AllowsTransparency="True">'
                '<Border Background="#2B3340" CornerRadius="10" Margin="12" Padding="24,20,24,20">'
                '<Border.Effect><DropShadowEffect Color="Black" Opacity="0.5" ShadowDepth="4" BlurRadius="16"/></Border.Effect>'
                '<StackPanel>'
                '<Border Background="#208A3C" Height="3" CornerRadius="2" Margin="0,0,0,16"/>'
                '<TextBlock Text="Settings Mismatch"'
                ' Foreground="#F4FAFF" FontSize="15" FontWeight="Bold" Margin="0,0,0,8"/>'
                '<TextBlock Text="This project was previously issued with different settings."'
                ' Foreground="#F4FAFF" FontSize="12" Opacity="0.85" TextWrapping="Wrap" Margin="0,0,0,24"/>'
                '<StackPanel Orientation="Horizontal" HorizontalAlignment="Right">'
                '<Button x:Name="ignore_btn" Content="Ignore"'
                ' Foreground="#F4FAFF" FontSize="12" FontWeight="Bold"'
                ' BorderThickness="0" Padding="20,8" Margin="0,0,8,0" Cursor="Hand">'
                '<Button.Template><ControlTemplate TargetType="Button">'
                '<Border x:Name="Bd" Background="#404553" CornerRadius="6" Padding="{TemplateBinding Padding}">'
                '<ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>'
                '</Border>'
                '<ControlTemplate.Triggers>'
                '<Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="#4E5566"/></Trigger>'
                '<Trigger Property="IsPressed" Value="True"><Setter TargetName="Bd" Property="Background" Value="#333B48"/></Trigger>'
                '</ControlTemplate.Triggers>'
                '</ControlTemplate></Button.Template>'
                '</Button>'
                '<Button x:Name="session_btn" Content="This Issue Only"'
                ' Foreground="#F4FAFF" FontSize="12" FontWeight="Bold"'
                ' BorderThickness="0" Padding="20,8" Margin="0,0,8,0" Cursor="Hand">'
                '<Button.Template><ControlTemplate TargetType="Button">'
                '<Border x:Name="Bd" Background="#404553" CornerRadius="6" Padding="{TemplateBinding Padding}">'
                '<ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>'
                '</Border>'
                '<ControlTemplate.Triggers>'
                '<Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="#4E5566"/></Trigger>'
                '<Trigger Property="IsPressed" Value="True"><Setter TargetName="Bd" Property="Background" Value="#333B48"/></Trigger>'
                '</ControlTemplate.Triggers>'
                '</ControlTemplate></Button.Template>'
                '</Button>'
                '<Button x:Name="update_btn" Content="Update Settings"'
                ' Foreground="#F4FAFF" FontSize="12" FontWeight="Bold"'
                ' BorderThickness="0" Padding="20,8" Cursor="Hand">'
                '<Button.Template><ControlTemplate TargetType="Button">'
                '<Border x:Name="Bd" Background="#208A3C" CornerRadius="6" Padding="{TemplateBinding Padding}">'
                '<ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>'
                '</Border>'
                '<ControlTemplate.Triggers>'
                '<Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="#2B933F"/></Trigger>'
                '<Trigger Property="IsPressed" Value="True"><Setter TargetName="Bd" Property="Background" Value="#1A6E2E"/></Trigger>'
                '</ControlTemplate.Triggers>'
                '</ControlTemplate></Button.Template>'
                '</Button>'
                '</StackPanel></StackPanel></Border></Window>'
            )
            dlg = Markup.XamlReader.Parse(xaml)
            result = ['ignore']
            def on_update(s, e):  result[0] = 'update';  dlg.Close()
            def on_session(s, e): result[0] = 'session'; dlg.Close()
            def on_ignore(s, e):  result[0] = 'ignore';  dlg.Close()
            dlg.FindName("update_btn").Click  += on_update
            dlg.FindName("session_btn").Click += on_session
            dlg.FindName("ignore_btn").Click  += on_ignore
            dlg.ShowDialog()
            return result[0]
        except:
            return 'ignore'

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

        dist_file = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                 'Settings', 'distribution.json')
        rows = []
        try:
            with open(dist_file, 'r') as f:
                rows = _json.load(f)
        except:
            rows = [{'distribution': 'Architect/Designer'},
                    {'distribution': 'Owner/Developer'},
                    {'distribution': 'Contractor'},
                    {'distribution': 'Local Authority'}]

        white = _SWM.SolidColorBrush(_SWM.Color.FromRgb(0xF4, 0xFA, 0xFF))

        for item in rows:
            label = (item.get('distribution', '')
                     or item.get('Distribution', '')
                     or str(item))

            # Row: DockPanel — label stretches, copies fixed right, attn fills middle
            dp = _SWC.DockPanel()
            dp.Margin = _SW.Thickness(0, 0, 0, 4)
            dp.LastChildFill = True

            # Label — left side, fixed width
            lbl = _SWC.TextBlock()
            lbl.Text = label
            lbl.Foreground = white
            lbl.FontSize = 12
            lbl.Width = 140
            lbl.VerticalAlignment = _SW.VerticalAlignment.Center
            lbl.Margin = _SW.Thickness(0, 0, 6, 0)
            _SWC.DockPanel.SetDock(lbl, _SWC.Dock.Left)
            dp.Children.Add(lbl)

            # Copies — right side, fixed width
            copies_tb = _SWC.TextBox()
            copies_tb.Width = 54
            copies_tb.HorizontalContentAlignment = _SW.HorizontalAlignment.Center
            copies_tb.Margin = _SW.Thickness(4, 0, 0, 0)
            _SWC.DockPanel.SetDock(copies_tb, _SWC.Dock.Right)
            try: copies_tb.Style = self.FindResource("ModernTextBoxStyle")
            except: pass
            dp.Children.Add(copies_tb)

            # Attention To — fills remaining space
            attn_tb = _SWC.TextBox()
            attn_tb.HorizontalAlignment = _SW.HorizontalAlignment.Stretch
            try: attn_tb.Style = self.FindResource("ModernTextBoxStyle")
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
            white = _SWM.SolidColorBrush(_SWM.Color.FromRgb(0xF4, 0xFA, 0xFF))
            hdr_dp = _SWC.DockPanel()
            hdr_dp.Margin = _SW.Thickness(0, 0, 0, 2)

            # Copies label — fixed width, docked right
            copies_lbl = _SWC.TextBlock()
            copies_lbl.Text = "Copies"
            copies_lbl.Width = 46
            copies_lbl.Foreground = white
            copies_lbl.FontSize = 11
            copies_lbl.HorizontalAlignment = _SW.HorizontalAlignment.Center
            copies_lbl.Margin = _SW.Thickness(4, 0, 0, 0)
            _SWC.DockPanel.SetDock(copies_lbl, _SWC.Dock.Right)
            hdr_dp.Children.Add(copies_lbl)

            # Company / Contact labels — split evenly
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
            pass   # header labels optional — don't let them block row building

        # Pre-fill rows from last issued revision if it was saved in client mode
        _prefilled = False
        try:
            if self.issued_revs:
                import re as _re_cl
                _last_ito = (self.issued_revs[-1].IssuedTo or '')
                _cl_m = _re_cl.search(r'CL:\s*(.*?)(?:\s*\|[^|]|$)', _last_ito)
                if _cl_m:
                    _cl_block = _cl_m.group(1).strip()
                    # Format: [Company — Contact]copies  or  [Company]copies
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
        rec_file = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                'Settings', 'recipients.json')
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

        white = _SWM.SolidColorBrush(_SWM.Color.FromRgb(0xF4, 0xFA, 0xFF))

        # Outer DockPanel: copies fixed right, company+contact fill left
        dp = _SWC.DockPanel()
        dp.Margin = _SW.Thickness(0, 0, 0, 4)
        dp.LastChildFill = True

        # Copies textbox — docked right
        copies_tb = _SWC.TextBox()
        copies_tb.Width = 46
        copies_tb.HorizontalContentAlignment = _SW.HorizontalAlignment.Center
        copies_tb.Margin = _SW.Thickness(4, 0, 0, 0)
        copies_tb.Text = preset_copies
        _SWC.DockPanel.SetDock(copies_tb, _SWC.Dock.Right)
        try: copies_tb.Style = self.FindResource("ModernTextBoxStyle")
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
        try: company_cb.Style = self.FindResource("ModernComboBoxStyle")
        except: pass
        grid.Children.Add(company_cb)

        # Contact dropdown — left margin matches the gap between company and grid edge
        contact_cb = _SWC.ComboBox()
        contact_cb.ItemsSource = ['(Select Contact)']
        contact_cb.SelectedIndex = 0
        contact_cb.IsEnabled = False
        contact_cb.Margin = _SW.Thickness(4, 0, 0, 0)
        _SWC.Grid.SetColumn(contact_cb, 1)
        try: contact_cb.Style = self.FindResource("ModernComboBoxStyle")
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
                return  # no valid export path — skip silently

            script_dir = os.path.dirname(os.path.abspath(__file__))
            src_layouts = os.path.join(script_dir, 'Layout', 'Layouts')
            if not os.path.isdir(src_layouts):
                return  # no layouts to export

            dest_settings = os.path.join(export_path, 'pyTransmit Settings')
            dest_layouts  = os.path.join(dest_settings, 'Layouts')
            if not os.path.isdir(dest_layouts):
                os.makedirs(dest_layouts)

            # Also copy layout_config.json
            src_config = os.path.join(script_dir, 'Layout', 'layout_config.json')
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
            pass  # silent — don't break the main export if layouts fail

    def export_settings_back_click(self, sender, args):
        if self.export_ctrl: self.export_ctrl.save_config()
        self._save_layout_assignments()
        self._show_panel("main")

    def layout_assignment_changed(self, sender, args):
        """Called when any layout combo selection changes — save immediately."""
        self._save_layout_assignments()

    def _layouts_dir(self):
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Layout', 'Layouts')

    def _layout_templates(self):
        """Return sorted list of JSON template names from Layout/Layouts/."""
        d = self._layouts_dir()
        if not os.path.isdir(d): return []
        return sorted([os.path.splitext(f)[0] for f in os.listdir(d)
                       if f.lower().endswith('.json')])

    _LAYOUT_COMBOS = {
        'layout_schedule_cb': 'Revit Schedule',
        'layout_drafting_cb': 'Revit Drafting View',
        'layout_legend_cb':   'Revit Legend',
        'layout_excel_cb':    'Excel',
        'layout_pdf_cb':      'PDF',
    }

    def _on_content_rendered(self, sender, args):
        """Called after window is fully rendered — visual tree is available."""
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
                Media.SolidColorBrush(Media.Color.FromRgb(0x20, 0x8A, 0x3C)))
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
            sel = saved.get(cb_name, default)
            cb.SelectedItem = sel if sel in templates else \
                (default if default in templates else templates[0] if templates else '(none)')

    def _save_layout_assignments(self):
        """Persist layout combo selections to pytransmit_sync.json."""
        try:
            sync_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pytransmit_sync.json')
            cfg = {}
            if os.path.isfile(sync_path):
                with open(sync_path, 'r') as f: cfg = json.load(f)
            for cb_name in self._LAYOUT_COMBOS:
                cb = getattr(self, cb_name, None)
                if cb and cb.SelectedItem:
                    cfg['layout_assign_{}'.format(cb_name)] = str(cb.SelectedItem)
            with open(sync_path, 'w') as f: json.dump(cfg, f, indent=2)
        except Exception: pass

    def _load_layout_assignments(self):
        """Load layout combo selections from pytransmit_sync.json."""
        try:
            sync_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pytransmit_sync.json')
            if os.path.isfile(sync_path):
                with open(sync_path, 'r') as f: cfg = json.load(f)
                return {k.replace('layout_assign_', ''): v
                        for k, v in cfg.items() if k.startswith('layout_assign_')}
        except Exception: pass
        return {}

    def get_layout_for_output(self, output_type):
        """Return the full path to the selected layout JSON for a given output type.
        output_type: 'excel', 'pdf', 'schedule', 'drafting', 'legend'"""
        cb_map = {'excel': 'layout_excel_cb', 'pdf': 'layout_pdf_cb',
                  'schedule': 'layout_schedule_cb', 'drafting': 'layout_drafting_cb',
                  'legend': 'layout_legend_cb'}
        cb = getattr(self, cb_map.get(output_type, ''), None)
        name = str(cb.SelectedItem) if cb and cb.SelectedItem else output_type.capitalize()
        if name == '(none)': return None
        path = os.path.join(self._layouts_dir(), name + '.json')
        return path if os.path.isfile(path) else None

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
            dest_layouts = os.path.join(script_dir, 'Layout', 'Layouts')
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
                _shutil.copy2(src_cfg, os.path.join(script_dir, 'Layout', 'layout_config.json'))
        except Exception:
            pass

    def import_settings_back_click(self, sender, args):
        if self.import_ctrl: self.import_ctrl.save_config()
        self._show_panel("main")

    def menu_help_click(self, sender, args):
        """☰ → Help: open HELP_URL in the default browser."""
        self.OptionsPopup.IsOpen   = False
        self.options_btn.IsChecked = False
        try:
            import subprocess
            subprocess.Popen(['cmd', '/c', 'start', '', HELP_URL])
        except Exception, e:
            forms.alert("Could not open browser:\n{}".format(str(e)), title="Help")

    def run_excel_export(self):
        """Run Excel export — loads script_excel.py in a clean namespace."""
        try:
            # Read export path from the Export Settings panel (export_path_tb)
            export_path = r'C:\Temp'
            tb = getattr(self, 'export_path_tb', None)
            if tb and tb.Text:
                export_path = tb.Text
            if not export_path or not os.path.exists(export_path):
                forms.alert("Excel export path does not exist:\n{}\n\nSet the path in  ☰ → Export Settings.".format(export_path))
                return

            script_dir   = os.path.dirname(os.path.abspath(__file__))
            excel_script = os.path.join(script_dir, "script_excel.py")
            if not os.path.exists(excel_script):
                forms.alert("script_excel.py not found in:\n{}".format(script_dir))
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

            # Execute in isolated namespace — __name__ != '__main__' so no
            # entry-point guards fire; IronPython 2 compatible (no compile())
            ns = {'__name__': 'excel_export', '__file__': excel_script,
                  '__builtins__': __builtins__}
            with open(excel_script, 'r') as f:
                src = f.read()
            exec(src, ns)

        except Exception, e:
            forms.alert("Error exporting to Excel:\n{}".format(str(e)))
    
    def run_revit_export(self):
        """Run Revit export — dispatches to the correct Publish script based on output_type."""
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
                from Autodesk.Revit.DB import FilteredElementCollector, Revision
                _all_revs = sorted(
                    FilteredElementCollector(self.doc).OfClass(Revision).ToElements(),
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
                            # client format: [Company — Contact]copies
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
                '_settings_dir':     settings_dir,
                'script_dir':        script_dir,
            }

            # ── Dispatch each selected output type ────────────────────────────
            for output_type in output_types:

                if output_type == 'excel':
                    target_script = os.path.join(publish_dir, 'script_create_excel.py')
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
                    forms.alert("{} not found at:\n{}".format(err_label, target_script))
                    continue

                payload['output_type']      = output_type
                payload['layout_json_path'] = self.get_layout_for_output(output_type)
                ns = {
                    '__name__':    script_name,
                    '__file__':    target_script,
                    '__builtins__': __builtins__,
                    'PYTRANSMIT_PAYLOAD': payload,
                }
                with open(target_script, 'r') as f:
                    src = f.read()
                try:
                    exec(src, ns)
                    # Suppress the pyRevit output window after each script runs
                    try:
                        _out = script.get_output()
                        if _out: _out.hide()
                    except Exception: pass
                except Exception, exec_e:
                    import traceback as _tb
                    tb_str = _tb.format_exc() or str(exec_e) or repr(exec_e)
                    forms.alert("Error running {}:\n{}".format(err_label, tb_str))

        except Exception, e:
            import traceback
            tb_str = traceback.format_exc() or str(e) or repr(e)
            forms.alert("Error exporting Revit data:\n{}".format(tb_str))


# --- generate_tables removed - this script only updates revision data ---

# --- MAIN EXECUTION ---
def main():
    try:
        # Check for update before opening window
        _pt_check_and_notify()

        window = RevTableWindow()

        # Show ribbon if update was found
        remote_ver = getattr(_pt_check_and_notify, '_remote_ver', None)
        if remote_ver:
            local_ver = getattr(_pt_check_and_notify, '_local_ver', 'unknown')
            try:
                ribbon  = window.FindName("update_ribbon")
                ver_lbl = window.FindName("update_ribbon_version")
                if ribbon:
                    import System.Windows as _SW
                    ribbon.Visibility = _SW.Visibility.Visible
                if ver_lbl:
                    ver_lbl.Text = u"v{0}  \u2192  v{1}".format(local_ver, remote_ver)

                def _on_ribbon_click(s, e):
                    changes     = getattr(_pt_check_and_notify, '_changes', [])
                    change_text = u"\n".join([u"  \u2022 " + c for c in changes[:3]]) if changes else ""
                    msg = u"Update PyTransmit to v{0}?\n\n{1}".format(remote_ver, change_text)
                    result = forms.alert(msg, title="PyTransmit Update",
                                         ok=False, yes=True, no=True)
                    if not result:
                        return
                    ok, err = _pt_do_update(remote_ver)
                    if ok:
                        forms.alert(
                            u"PyTransmit updated to v{0}.\n\nReloading PyRevit...".format(remote_ver),
                            title="PyTransmit Updated"
                        )
                        window.Close()
                        try:
                            from pyrevit.loader import sessionmgr
                            sessionmgr.reload_pyrevit()
                        except Exception:
                            pass
                    else:
                        forms.alert(u"Update failed:\n\n{}".format(err),
                                    title="PyTransmit Update")

                if ribbon:
                    ribbon.MouseLeftButtonUp += _on_ribbon_click
            except Exception:
                pass

        window.ShowDialog()
    except Exception as ex:
        forms.alert("Error initializing window: {}".format(str(ex)), exitscript=True)


if __name__ == "__main__":
    main()