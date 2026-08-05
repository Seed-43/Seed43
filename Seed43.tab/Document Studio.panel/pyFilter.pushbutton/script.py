# -*- coding: utf-8 -*-
# script.py
# Seed43 Filter Manager - Main entry point
# pylint: disable=import-error,invalid-name,broad-except

import io
import os
import sys
import json
import traceback
import datetime

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import Visibility
from System.Windows.Controls import Button, CheckBox

from pyrevit import revit, DB, forms, script
from pyrevit.forms import WPFWindow

_lib_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "lib")
if _lib_path not in sys.path:
    sys.path.insert(0, _lib_path)

import pyfilter_save as fs
import pyfilter_apply as fa
import pyfilter_view_assign as va
import pyfilter_sync as sm
import pyfilter_settings as settings_dialog
from Snippets._revisions import safe_str
from Snippets.seed43_theme import apply_seed43_palette
from Snippets import _userdata

doc    = revit.doc
output = None  # pyRevit output panel disabled

# ── LOGGING ───────────────────────────────────────────────────────────────────
# IMPORTANT: pyRevit auto-opens its console window on the first print() call
# from a script, so logging must NOT use print()/print_html()/get_output().
# Instead we write to a local log file only — silent, no popup ever.

_LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pyfilter.log")

def log(msg, level="INFO"):
    ts = datetime.datetime.now().strftime("%H:%M:%S")
    line = "[{t}] [{l}] {m}".format(t=ts, l=level, m=msg)
    try:
        with open(_LOG_FILE, "a") as f:
            f.write(line + "\n")
    except Exception:
        pass

def log_exc(context):
    log("{} -- {}".format(context, traceback.format_exc()), "ERR")

# ── PATHS ─────────────────────────────────────────────────────────────────────

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Shipped demo templates. Read-only, never what the tool reads from, so an
# update can refresh them freely without anything reappearing for the user.
DEFAULTS_DIR = os.path.join(SCRIPT_DIR, "defaults")


def get_templates_folder():
    """
    The user's templates folder, under .user so updates cannot touch it.

    Does three things, each at most once:
      1. resolves .user/pyFilter/templates/
      2. carries across templates saved beside the tool by older versions
      3. seeds the shipped demos, first run only

    Step 3 is marker-based on purpose: a demo the user deletes must stay
    deleted through every future update. See _userdata.seed_once. "Load demo
    data" in Settings is the way back if they change their mind.
    """
    folder = _userdata.user_dir("pyFilter", "templates")
    _userdata.migrate_dir(os.path.join(SCRIPT_DIR, "templates"), folder)
    _userdata.seed_once(
        _userdata.user_path("pyFilter", ".seeded"), DEFAULTS_DIR, folder)
    return folder

def list_templates(folder):
    try:
        return sorted([os.path.splitext(f)[0]
                       for f in os.listdir(folder) if f.endswith(".json")])
    except Exception:
        return []

def load_template_file(folder, name):
    path = os.path.join(folder, name + ".json")
    try:
        with io.open(path, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return None

# ── MAIN WINDOW ───────────────────────────────────────────────────────────────

class pyFilterWindow(WPFWindow):

    def __init__(self):
        xaml_path = os.path.join(SCRIPT_DIR, "pyFilter.xaml")
        log("Loading XAML: {}".format(xaml_path))
        WPFWindow.__init__(self, xaml_path)
        apply_seed43_palette(self, SCRIPT_DIR)

        self.templates_folder  = get_templates_folder()
        self.active_template   = None
        self.filter_rows       = []
        self._template_dirty   = False
        self._loading_template = False
        # Template mode row selection
        self._tpl_all_check    = None
        self._tpl_row_checks   = []
        self._tpl_last_clicked = None
        self._suppress_tpl_all = False
        self.fill_patterns     = fs.get_fill_patterns()
        self.line_patterns     = fs.get_line_patterns()
        self._template_buttons = {}

        # mode: "templates" | "views" | "viewtemplates"
        self.mode            = "templates"
        self.live_view_map   = {}
        self.source_template_name = None
        self.source_template_data = None
        # Multi-select destinations (checked view/template display names).
        self.selected_views  = set()
        self._view_checks    = {}
        # Within assign mode the toolbar swaps grid contents:
        #   "push" = template -> destination views
        #   "pull" = source view -> template
        self.direction       = "push"
        self.push_filter_state = {}
        self.pull_source_view  = None
        self.pull_filter_state = {}
        # Search query applied to the sidebar list.
        self._search_text    = ""

        log("Patterns -- fill:{} line:{}".format(
            len(self.fill_patterns), len(self.line_patterns)))

        # Toolbar wiring -- templates mode
        self.BtnNewTemplate.Click    += self._on_new_template
        self.BtnLoadDemo.Click       += self._on_load_demo
        self.BtnAddFilter.Click      += self._on_add_filters
        self.BtnDeleteFilter.Click   += self._on_remove_filter

        # Toolbar wiring -- assign mode
        self.BtnPushFilters.Click         += lambda s, e: self._set_direction("push")
        self.BtnPullFilters.Click         += lambda s, e: self._set_direction("pull")
        self.BtnDeleteFilterAssign.Click  += self._on_remove_filter

        # Mode tabs + permanent right-side actions
        self.BtnModeTemplates.Click  += lambda s, e: self._set_mode("templates")
        self.BtnModeViews.Click      += lambda s, e: self._set_mode("views")
        self.BtnModeVTemplates.Click += lambda s, e: self._set_mode("viewtemplates")
        self.BtnApply.Click          += self._on_apply
        self.TxtSearch.TextChanged   += self._on_search_changed

        # Hamburger menu items (ToggleButton popup, no manual open needed)
        self.MenuExport.Click   += self._on_menu_export
        self.MenuImport.Click   += self._on_menu_import
        self.MenuResetColumns.Click += self._on_menu_reset_columns
        self.MenuHelp.Click     += self._on_menu_help
        self.MenuAbout.Click    += self._on_menu_about

        # Inline panel close + action buttons
        self.BtnExportNow.Click      += self._on_export_execute
        self.BtnExportClose.Click    += self._on_export_close
        self.BtnExportBrowse.Click   += self._on_export_browse
        self.BtnExportAll.Click      += lambda s, e: self._set_export_all(True)
        self.BtnExportNone.Click     += lambda s, e: self._set_export_all(False)
        self.BtnImportNow.Click      += self._on_import_execute
        self.BtnImportClose.Click    += self._on_import_close
        self.BtnImportBrowse.Click   += self._on_import_browse
        self.BtnImportAll.Click      += lambda s, e: self._set_import_all(True)
        self.BtnImportNone.Click     += lambda s, e: self._set_import_all(False)
        self.TxtImportPath.TextChanged += self._on_import_path_changed

        # Export/import inline state
        self._export_checks = {}
        self._import_checks = {}
        # Filter grid tri-state "All" (assign mode)
        self._filter_all_check   = None
        self._filter_row_checks  = []
        self._suppress_filter_all = False
        self._filter_last_clicked = None

        # Load saved config into inline panel fields.
        try:
            self.TxtExportPath.Text  = sm.get_export_path()
            self.ExportAutoUpdate.IsChecked = sm.get_export_auto()
            self.TxtImportPath.Text  = sm.get_import_path()
            self.ImportAutoUpdate.IsChecked = sm.get_import_auto()
        except Exception:
            pass

        # Row context menu state (which template the menu is open for)
        self._row_menu_target = None
        self.RowMenuRename.Click += self._on_row_rename
        self.RowMenuDelete.Click += self._on_row_delete

        # Startup sync (silent if no server path configured).
        self._run_startup_sync()
        self._show_panel("main")
        self._rebuild_headers()
        self._set_mode("templates")

    def _run_startup_sync(self):
        try:
            result = sm.sync_all(self.templates_folder, logger=log)
            if result.get("ok") and result.get("counts"):
                counts = result["counts"]
                if any(k in counts for k in ("uploaded", "downloaded", "restored")):
                    log("Startup sync: {}".format(counts), "OK")
        except Exception:
            log_exc("_run_startup_sync")

    # ── PANEL SHOW/HIDE ───────────────────────────────────────────────────────

    def _show_panel(self, panel_name):
        panels = ["main_content", "ExportPanel", "ImportPanel"]
        header_normal = ["HeaderNormalBtns"]
        header_export = ["ExportHeaderBtns", "ExportHeaderLbl"]
        header_import = ["ImportHeaderBtns", "ImportHeaderLbl"]
        # Toolbar contextual buttons to hide in export/import mode
        toolbar_btns  = ["BtnAddFilter", "BtnDeleteFilter",
                         "BtnPushFilters", "BtnPullFilters", "BtnDeleteFilterAssign",
                         "BtnNewTemplate"]

        all_header = header_normal + header_export + header_import

        def vis(name, show):
            el = self.FindName(name)
            if el is not None:
                el.Visibility = (Visibility.Visible if show else Visibility.Collapsed)

        for p in panels + all_header:
            vis(p, False)

        if panel_name == "main":
            vis("main_content", True)
            vis("HeaderNormalBtns", True)
            # Restore toolbar buttons — _set_mode handles which ones are visible
            self._update_toolbar_visibility()
        elif panel_name == "export":
            self._populate_export_list()
            vis("ExportPanel", True)
            for n in header_export: vis(n, True)
            for n in toolbar_btns: vis(n, False)
        elif panel_name == "import":
            self._populate_import_list()
            vis("ImportPanel", True)
            for n in header_import: vis(n, True)
            for n in toolbar_btns: vis(n, False)

    def _update_toolbar_visibility(self):
        """Restore toolbar buttons to the correct state for the current mode."""
        assign = self.mode in ("views", "viewtemplates")
        for btn_name in ("BtnAddFilter", "BtnDeleteFilter", "BtnNewTemplate"):
            el = self.FindName(btn_name)
            if el: el.Visibility = (Visibility.Collapsed if assign else Visibility.Visible)
        for btn_name in ("BtnPushFilters", "BtnPullFilters", "BtnDeleteFilterAssign"):
            el = self.FindName(btn_name)
            if el: el.Visibility = (Visibility.Visible if assign else Visibility.Collapsed)

    # ── HAMBURGER MENU ────────────────────────────────────────────────────────

    def _on_hamburger(self, sender, e):
        try:
            self.MenuPopup.PlacementTarget = sender
            self.MenuPopup.IsOpen = True
        except Exception:
            log_exc("_on_hamburger")

    def _on_menu_export(self, sender, e):
        try:
            self.options_btn.IsChecked = False
            self._show_panel("export")
        except Exception:
            log_exc("_on_menu_export")

    def _on_menu_import(self, sender, e):
        try:
            self.options_btn.IsChecked = False
            self._show_panel("import")
        except Exception:
            log_exc("_on_menu_import")

    def _on_menu_reset_columns(self, sender, e):
        try:
            self.options_btn.IsChecked = False
            new_widths = fs.reset_col_widths()
            fs.update_all_row_widths(new_widths)
            self._rebuild_headers()
            log("Column widths reset to defaults.", "OK")
        except Exception:
            log_exc("_on_menu_reset_columns")

    def _on_menu_help(self, sender, e):
        try:
            self.options_btn.IsChecked = False
            import webbrowser
            webbrowser.open("https://seed43.org/pyfilter/")
        except Exception:
            log_exc("_on_menu_help")

    def _on_menu_about(self, sender, e):
        try:
            self.options_btn.IsChecked = False
            forms.alert(
                "pyFilter\n"
                "Part of the Seed43 extension suite\n\n"
                "Manage, save, push and pull Revit view filters "
                "across templates and views.\n\n"
                "seed43.org",
                title="About pyFilter")
        except Exception:
            log_exc("_on_menu_about")

    # ── EXPORT PANEL ──────────────────────────────────────────────────────────

    def _populate_export_list(self):
        from System.Windows.Controls import CheckBox as WpfCheckBox
        from System.Windows import Thickness
        self.ExportTemplateList.Children.Clear()
        self._export_checks = {}
        for name in list_templates(self.templates_folder):
            cb = WpfCheckBox()
            cb.Content   = name
            cb.Tag       = name
            cb.IsChecked = True
            cb.Foreground = self._white_brush()
            cb.FontSize  = 12
            cb.Margin    = Thickness(0, 3, 0, 3)
            self.ExportTemplateList.Children.Add(cb)
            self._export_checks[name] = cb

    def _set_export_all(self, value):
        for cb in self._export_checks.values():
            cb.IsChecked = bool(value)

    def _on_export_browse(self, sender, e):
        try:
            folder = forms.pick_folder(title="Pick export destination")
            if folder:
                self.TxtExportPath.Text = folder
        except Exception:
            log_exc("_on_export_browse")

    def _on_export_execute(self, sender, e):
        try:
            dest = (self.TxtExportPath.Text or "").strip()
            if not dest:
                forms.alert("Pick a destination folder first.")
                return
            if not os.path.isdir(dest):
                try: os.makedirs(dest)
                except Exception as ex:
                    forms.alert("Cannot create destination: {}".format(ex))
                    return
            if not os.access(dest, os.W_OK):
                forms.alert("Destination is not writable.")
                return
            selected = [n for n, cb in self._export_checks.items() if cb.IsChecked]
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
            self._set_status("Exported {}.".format(written))
            forms.alert(msg, title="Export Templates")
        except Exception:
            log_exc("_on_export_execute")

    def _on_export_close(self, sender, e):
        try:
            self._save_export_config()
            self._show_panel("main")
        except Exception:
            log_exc("_on_export_close")

    def _save_export_config(self):
        try:
            sm.set_export_path((self.TxtExportPath.Text or "").strip())
            sm.set_export_auto(bool(self.ExportAutoUpdate.IsChecked))
        except Exception:
            pass

    # ── IMPORT PANEL ──────────────────────────────────────────────────────────

    def _populate_import_list(self):
        from System.Windows.Controls import CheckBox as WpfCheckBox
        from System.Windows import Thickness
        self.ImportTemplateList.Children.Clear()
        self._import_checks = {}
        src = (self.TxtImportPath.Text or "").strip()
        local = set(list_templates(self.templates_folder))
        names = []
        if src and os.path.isdir(src):
            for f in os.listdir(src):
                if f.lower().endswith(".json"):
                    names.append(os.path.splitext(f)[0])
        for name in sorted(names):
            cb = WpfCheckBox()
            exists = name in local
            cb.Content   = u"{}  {}".format(
                name, u"(\u2713 exists)" if exists else u"(new)")
            cb.Tag       = name
            cb.IsChecked = True
            cb.Foreground = self._white_brush()
            cb.FontSize  = 12
            cb.Margin    = Thickness(0, 3, 0, 3)
            self.ImportTemplateList.Children.Add(cb)
            self._import_checks[name] = cb

    def _on_import_path_changed(self, sender, e):
        try:
            self._populate_import_list()
        except Exception:
            pass

    def _set_import_all(self, value):
        for cb in self._import_checks.values():
            cb.IsChecked = bool(value)

    def _on_import_browse(self, sender, e):
        try:
            folder = forms.pick_folder(title="Pick source folder")
            if folder:
                self.TxtImportPath.Text = folder
        except Exception:
            log_exc("_on_import_browse")

    def _on_import_execute(self, sender, e):
        try:
            src = (self.TxtImportPath.Text or "").strip()
            if not src or not os.path.isdir(src):
                forms.alert("Pick a valid source folder.")
                return
            selected = [n for n, cb in self._import_checks.items() if cb.IsChecked]
            if not selected:
                forms.alert("Tick at least one template to import.")
                return
            if self.RdoOverwrite.IsChecked:  mode = "overwrite"
            elif self.RdoSkip.IsChecked:     mode = "skip"
            else:                            mode = "rename"
            import shutil
            local = set(list_templates(self.templates_folder))
            taken = set(local)
            imported = skipped = renamed = overwritten = 0
            errors = []
            for name in selected:
                src_path = os.path.join(src, name + ".json")
                if not os.path.isfile(src_path):
                    continue
                target = name
                if name in local:
                    if mode == "skip":
                        skipped += 1
                        continue
                    if mode == "rename":
                        i = 2
                        while True:
                            cand = u"{} ({})".format(name, i)
                            if cand not in taken:
                                target = cand
                                break
                            i += 1
                tgt_path = os.path.join(self.templates_folder, target + ".json")
                try:
                    shutil.copy2(src_path, tgt_path)
                    taken.add(target)
                    if target != name:         renamed += 1
                    elif name in local:        overwritten += 1
                    else:                      imported += 1
                except Exception as ex:
                    errors.append(u"{}: {}".format(name, ex))
            msg = u"Imported: {} new, {} overwritten, {} renamed, {} skipped.".format(
                imported, overwritten, renamed, skipped)
            if errors: msg += u"\n\nErrors:\n" + u"\n".join(errors)
            self._set_status(msg)
            forms.alert(msg, title="Import Templates")
            self._refresh_sidebar()
        except Exception:
            log_exc("_on_import_execute")

    def _on_import_close(self, sender, e):
        try:
            self._save_import_config()
            self._show_panel("main")
        except Exception:
            log_exc("_on_import_close")

    def _save_import_config(self):
        try:
            sm.set_import_path((self.TxtImportPath.Text or "").strip())
            sm.set_import_auto(bool(self.ImportAutoUpdate.IsChecked))
        except Exception:
            pass

    def _white_brush(self):
        from System.Windows.Media import SolidColorBrush, Color
        return SolidColorBrush(Color.FromRgb(244, 250, 255))

    # ── ROW CONTEXT MENU (template ⋮) ────────────────────────────────────────

    def _on_template_dots_click(self, sender, e):
        try:
            self._row_menu_target = sender.Tag
            self.RowMenuPopup.PlacementTarget = sender
            self.RowMenuPopup.IsOpen = True
            e.Handled = True
        except Exception:
            log_exc("_on_template_dots_click")

    def _on_row_rename(self, sender, e):
        try:
            self.RowMenuPopup.IsOpen = False
            old = self._row_menu_target
            if not old:
                return
            new = forms.ask_for_string(
                default=old, prompt="New name for '{}':".format(old),
                title="Rename Template")
            if not new:
                return
            new = new.strip()
            illegal = set('/\\:*?"<>|')
            if not new or any(c in illegal for c in new):
                forms.alert("Invalid template name.")
                return
            if new == old:
                return
            existing = set(list_templates(self.templates_folder))
            if new in existing:
                forms.alert("Template '{}' already exists.".format(new))
                return
            old_path = os.path.join(self.templates_folder, old + ".json")
            new_path = os.path.join(self.templates_folder, new + ".json")
            try:
                os.rename(old_path, new_path)
                # Update the "name" inside the JSON too.
                tpl = load_template_file(self.templates_folder, new)
                if tpl is not None:
                    tpl["name"] = new
                    fs.write_template_data(self.templates_folder, new, tpl)
            except Exception as ex:
                log("Rename failed: {}".format(ex), "ERR")
                forms.alert("Rename failed: {}".format(ex))
                return
            log("Renamed: {} -> {}".format(old, new), "OK")
            sm.sync_after_save(new, self.templates_folder, logger=log)
            if self.active_template == old:
                self.active_template = new
                self.source_template_name = new
            self._refresh_sidebar()
            self._set_status("Renamed '{}' to '{}'.".format(old, new))
        except Exception:
            log_exc("_on_row_rename")

    def _on_row_delete(self, sender, e):
        try:
            self.RowMenuPopup.IsOpen = False
            name = self._row_menu_target
            if not name:
                return
            if not self._confirm(
                    "Delete template '{}'?".format(name),
                    title="Delete Template"):
                return
            path = os.path.join(self.templates_folder, name + ".json")
            try:
                os.remove(path)
                log("Deleted: {}".format(name), "WARN")
            except Exception as ex:
                log("Delete failed: {}".format(ex), "ERR")
                forms.alert("Could not delete: {}".format(ex))
                return
            if self.active_template == name:
                self.active_template = None
                self.source_template_name = None
                self.source_template_data = None
                self.filter_rows = []
                self.FilterRowsPanel.Children.Clear()
                self.TxtTargetLabel.Text = ""
                self._set_empty_state(True)
            self._refresh_sidebar()
            self._set_status("Deleted '{}'.".format(name))
        except Exception:
            log_exc("_on_row_delete")

    # ── MODE SWITCHING ────────────────────────────────────────────────────────

    def _set_mode(self, mode):
        try:
            if mode == self.mode and self._template_buttons:
                return  # already in this mode; don't wipe the grid
            log("Mode: {}".format(mode))
            self.mode = mode

            assign = mode in ("views", "viewtemplates")

            # Toolbar visibility - templates-mode actions vs assign-mode actions
            for btn in (self.BtnAddFilter, self.BtnDeleteFilter):
                btn.Visibility = (Visibility.Collapsed if assign else Visibility.Visible)
            for btn in (self.BtnPushFilters, self.BtnPullFilters,
                        self.BtnDeleteFilterAssign):
                btn.Visibility = (Visibility.Visible if assign else Visibility.Collapsed)

            # Sidebar title + search reset
            self.TxtSidebarTitle.Text = (
                "VIEW TEMPLATES" if mode == "viewtemplates"
                else ("VIEWS" if mode == "views" else "TEMPLATES"))
            self.TxtSearch.Text = ""
            self._search_text = ""

            # Apply button label
            self.BtnApply.Content = "Apply" if not assign else (
                "Push" if self.direction == "push" else "Pull")

            self._highlight_mode_buttons()
            self._rebuild_headers()

            # Reset grid + selections for the new mode.
            self.filter_rows = []
            self.FilterRowsPanel.Children.Clear()
            self.OptionsPanel.Children.Clear()
            self.selected_views = set()
            self.push_filter_state = {}
            self.pull_filter_state = {}
            self.pull_source_view  = None
            if not assign:
                self.active_template = None

            self._refresh_sidebar()

            if assign:
                # Default to push; _set_direction populates grid + labels.
                self._set_direction("push", force=True)
            else:
                templates = list_templates(self.templates_folder)
                if templates:
                    self._open_template(templates[0])
                else:
                    self._set_empty_state(True)
        except Exception:
            log_exc("_set_mode")

    # ── DIRECTION (Push / Pull) ───────────────────────────────────────────────

    def _set_direction(self, direction, force=False):
        try:
            if self.mode not in ("views", "viewtemplates"):
                return
            if direction == self.direction and not force:
                return
            self.direction = direction
            log("Direction: {}".format(direction))

            # Apply button label tracks direction.
            self.BtnApply.Content = "Push" if direction == "push" else "Pull"

            self._highlight_direction_buttons()
            self._rebuild_headers()

            self.filter_rows = []
            self.FilterRowsPanel.Children.Clear()
            self.OptionsPanel.Children.Clear()

            if direction == "push":
                if not self.source_template_data:
                    self.TxtTargetLabel.Text = "Pick a filter template first (Templates tab)"
                    self._set_empty_state(True)
                    self._set_status("Push needs a source template. Choose one under "
                                     "Templates, then come back.")
                else:
                    self._load_template_into_assign_grid()
                    self._update_assign_status()
            else:  # pull
                self.push_filter_state = {}
                if self.selected_views:
                    self._rebuild_pull_grid_from_selected(last_added=None)
                else:
                    dest = self.source_template_name or "(new template)"
                    self.TxtTargetLabel.Text = "Pull: tick a {} on the left  ->  {}".format(
                        "view template" if self.mode == "viewtemplates" else "view",
                        dest)
                    self._set_empty_state(True)
                    self._set_status("Pull reads filters off a view. Tick one on the "
                                     "left to load its filters into '{}'.".format(dest))
        except Exception:
            log_exc("_set_direction")

    def _highlight_direction_buttons(self):
        from System.Windows.Media import SolidColorBrush, Color
        def brush(hex_str):
            r = int(hex_str[1:3], 16); g = int(hex_str[3:5], 16); b = int(hex_str[5:7], 16)
            return SolidColorBrush(Color.FromRgb(r, g, b))
        active, inactive = "#208A3C", "#404553"
        self.BtnPushFilters.Background = brush(active if self.direction == "push" else inactive)
        self.BtnPullFilters.Background = brush(active if self.direction == "pull" else inactive)

    def _highlight_mode_buttons(self):
        active_bg   = "#208A3C"
        inactive_bg = "#404553"
        from System.Windows.Media import SolidColorBrush, Color
        def brush(hex_str):
            r = int(hex_str[1:3], 16); g = int(hex_str[3:5], 16); b = int(hex_str[5:7], 16)
            return SolidColorBrush(Color.FromRgb(r, g, b))
        mapping = {
            "templates":     self.BtnModeTemplates,
            "views":         self.BtnModeViews,
            "viewtemplates": self.BtnModeVTemplates,
        }
        for key, btn in mapping.items():
            btn.Background = brush(active_bg if key == self.mode else inactive_bg)

    def _rebuild_headers(self):
        assign = self.mode in ("views", "viewtemplates")
        from System.Windows import Thickness as _Th

        def _on_widths_changed(new_widths):
            # Push new widths into the group header ColumnDefinitions too
            try:
                from System.Windows import GridLength as _GL
                cdefs  = self.GroupHeaderGrid.ColumnDefinitions
                offset = 1 if assign else 0
                for i, w in enumerate(new_widths):
                    cd_idx = i + offset
                    if cd_idx < cdefs.Count:
                        cdefs[cd_idx].Width = _GL(w)
            except Exception:
                pass

        fs.build_group_header(self.GroupHeaderGrid, assign_mode=assign)
        fs.build_header(self.HeaderGrid, assign_mode=assign,
                        on_widths_changed=_on_widths_changed)
        # In templates mode every data row has a ~26px selection checkbox on the left.
        # Offset the headers by the same amount so columns line up.
        offset = 0 if assign else 26
        self.HeaderGrid.Margin      = _Th(offset, 0, 0, 0)
        self.GroupHeaderGrid.Margin = _Th(offset, 0, 0, 0)

    # ── SIDEBAR ───────────────────────────────────────────────────────────────

    def _on_search_changed(self, sender, e):
        try:
            self._search_text = (sender.Text or "").strip().lower()
            self._refresh_sidebar()
        except Exception:
            log_exc("_on_search_changed")

    def _matches_search(self, name):
        if not self._search_text:
            return True
        return self._search_text in name.lower()

    def _refresh_sidebar(self):
        if self.mode in ("views", "viewtemplates"):
            self._refresh_sidebar_views()
        else:
            self._refresh_sidebar_templates()

    def _refresh_sidebar_templates(self):
        from System.Windows import Thickness, GridLength, GridUnitType
        from System.Windows.Controls import Grid, ColumnDefinition
        self.TemplateListPanel.Children.Clear()
        self._template_buttons = {}
        for name in list_templates(self.templates_folder):
            if not self._matches_search(name):
                continue
            row = Grid()
            cd1 = ColumnDefinition()
            cd1.Width = GridLength(1, GridUnitType.Star)
            cd2 = ColumnDefinition()
            cd2.Width = GridLength.Auto
            row.ColumnDefinitions.Add(cd1)
            row.ColumnDefinitions.Add(cd2)

            btn = Button()
            btn.Content = name
            btn.Tag     = name
            btn.Style   = (self.Resources["SideBtnActive"]
                           if name == self.active_template
                           else self.Resources["SideBtn"])
            btn.Click += self._on_template_click
            Grid.SetColumn(btn, 0)
            row.Children.Add(btn)

            dots = Button()
            dots.Content = unichr(0x22EE)  # ⋮
            dots.Tag     = name
            dots.Style   = self.Resources["SideBtn"]
            dots.Padding = Thickness(8, 4, 8, 4)
            dots.FontSize = 16
            dots.Click += self._on_template_dots_click
            Grid.SetColumn(dots, 1)
            row.Children.Add(dots)

            self.TemplateListPanel.Children.Add(row)
            self._template_buttons[name] = btn

    def _refresh_sidebar_views(self):
        from System.Windows import Thickness
        from System.Windows.Media import SolidColorBrush, Color
        white = SolidColorBrush(Color.FromRgb(244, 250, 255))

        self.TemplateListPanel.Children.Clear()
        self._template_buttons = {}
        self._view_checks = {}
        # Ordered list of names as displayed (post-search-filter). Range
        # select uses this so "between A and B" means visible rows only.
        self._visible_view_names = []
        # Index of the last single-click for shift-click range select.
        self._last_clicked_index = None
        # Reentrancy guard so programmatic IsChecked changes during a range
        # toggle don't re-fire the row handler.
        self._suppress_check_handler = False

        templates = (self.mode == "viewtemplates")
        self.live_view_map = va.get_live_views(templates)
        log("Sidebar views: {}".format(len(self.live_view_map)))

        # Header: tri-state "all" check (works for both push and pull).
        header = CheckBox()
        from System.Windows.Media import SolidColorBrush, Color
        header.Content = "All"
        header.Foreground = SolidColorBrush(Color.FromRgb(244, 250, 255))
        header.FontSize  = 11
        header.Margin    = Thickness(2, 2, 2, 6)
        header.IsThreeState = True
        header.Click += self._on_select_all_click
        self.TemplateListPanel.Children.Add(header)
        self._select_all_check = header

        for name in sorted(self.live_view_map.keys()):
            if not self._matches_search(name):
                continue
            cb = CheckBox()
            cb.Content   = name
            cb.Tag       = name
            cb.Foreground = white
            cb.FontSize  = 12
            cb.Margin    = Thickness(2, 3, 2, 3)
            cb.IsChecked = (name in self.selected_views)
            # PreviewMouseLeftButtonDown fires before the box toggles, so we
            # can read Shift and handle range-select before the default click.
            cb.PreviewMouseLeftButtonDown += self._on_view_row_mouse_down
            cb.Checked   += self._on_view_check
            cb.Unchecked += self._on_view_check
            self.TemplateListPanel.Children.Add(cb)
            self._view_checks[name] = cb
            self._visible_view_names.append(name)

        self._refresh_select_all_state()

    def _highlight_active(self):
        # Only the template list uses single-select highlighting.
        if self.mode in ("views", "viewtemplates"):
            return
        for name, btn in self._template_buttons.items():
            btn.Style = (self.Resources["SideBtnActive"]
                         if name == self.active_template
                         else self.Resources["SideBtn"])

    # ── TEMPLATE OPEN ─────────────────────────────────────────────────────────

    def _open_template(self, name):
        log("Opening: {}".format(name))
        tpl = load_template_file(self.templates_folder, name)
        if not tpl:
            log("Could not load: {}".format(name), "ERR")
            return

        # Update state and clear UI immediately — highlight changes at once
        self.active_template = name
        self.source_template_name = name
        self.source_template_data = tpl
        self.TxtTargetLabel.Text = "Template: {}".format(name)
        self.filter_rows = []
        self._template_dirty = False
        self._tpl_all_check   = None
        self._tpl_row_checks  = []
        self._tpl_last_clicked = None
        self.FilterRowsPanel.Children.Clear()
        self._set_empty_state(True)
        fs.clear_row_col_defs()            # clear stale ColumnDefinition refs
        self._refresh_sidebar_templates()  # highlight updates immediately

        # Build rows without previews so UI appears instantly
        self._loading_template = True
        filters = sorted(tpl.get("filters", []), key=lambda f: f.get("name", "").lower())
        for fdata in filters:
            self._append_row(fdata["name"],
                             fdata.get("settings", {}),
                             fdata.get("definition", {}),
                             skip_preview=True)
        self._loading_template = False
        self._template_dirty = False
        self._set_empty_state(len(self.filter_rows) == 0)
        log("Loaded '{}' with {} filter(s)".format(name, len(self.filter_rows)), "OK")
        self._set_status("Loaded '{}'".format(name))

        # Collect all deferred preview callbacks and run them in the background
        # after the UI has painted — one per Dispatcher tick so scrolling stays smooth
        all_rebuilds = []
        for row in self.filter_rows:
            grid = row.get("grid")
            if grid is not None:
                deferred = fs._deferred_preview_store.pop(id(grid), None)
                if deferred:
                    all_rebuilds.extend(deferred)

        if all_rebuilds:
            # Snapshot the active template so stale callbacks abort if user switches
            target = name
            def _run_previews(rebuilds=all_rebuilds, tpl_name=target):
                if self.active_template != tpl_name:
                    return   # user already switched — discard
                try:
                    from System.Windows.Threading import Dispatcher, DispatcherPriority
                    import System
                    def _render_one(rb=rebuilds[0], rest=rebuilds[1:], t=tpl_name):
                        if self.active_template != t:
                            return
                        try:
                            rb()
                        except Exception:
                            pass
                        if rest:
                            def _next(r=rest, tn=t):
                                _run_previews_list(r, tn)
                            Dispatcher.CurrentDispatcher.BeginInvoke(
                                DispatcherPriority.Background,
                                System.Action(_next))
                    _render_one()
                except Exception:
                    # Fallback: run all synchronously
                    for rb in rebuilds:
                        try: rb()
                        except Exception: pass

            def _run_previews_list(rebuilds, tpl_name):
                if not rebuilds or self.active_template != tpl_name:
                    return
                try:
                    from System.Windows.Threading import Dispatcher, DispatcherPriority
                    import System
                    rb = rebuilds[0]
                    rest = rebuilds[1:]
                    try:
                        rb()
                    except Exception:
                        pass
                    if rest:
                        def _next(r=rest, t=tpl_name):
                            _run_previews_list(r, t)
                        Dispatcher.CurrentDispatcher.BeginInvoke(
                            DispatcherPriority.Background,
                            System.Action(_next))
                except Exception:
                    for rb in rebuilds:
                        try: rb()
                        except Exception: pass

            try:
                from System.Windows.Threading import Dispatcher, DispatcherPriority
                import System
                def _start(r=all_rebuilds, t=name):
                    _run_previews_list(r, t)
                Dispatcher.CurrentDispatcher.BeginInvoke(
                    DispatcherPriority.Background,
                    System.Action(_start))
            except Exception:
                for rb in all_rebuilds:
                    try: rb()
                    except Exception: pass

    def _append_row(self, filter_name, settings, definition=None, skip_preview=False):
        idx = len(self.filter_rows)

        # On first row, insert the tri-state All header
        if idx == 0:
            self._tpl_all_check  = self._build_template_grid_header()
            self._tpl_row_checks = []
            self._tpl_last_clicked = None

        row_data = {
            "name":       filter_name,
            "settings":   dict(settings),
            "definition": definition or {},
            "grid":       None,
        }
        self.filter_rows.append(row_data)
        if not self._loading_template:
            self._template_dirty = True

        def on_settings_changed(new_settings):
            row_data["settings"] = new_settings
            if not self._loading_template:
                self._template_dirty = True
                log("Updated: {}".format(filter_name))

        def on_sidebar_open(section_title, editor_type,
                            c_key, w_key, p_key,
                            fgc, fgp, fgv, bgc, bgp, bgv,
                            state, line_pats, fill_pats,
                            on_preview_rebuild, on_settings_cb):
            try:
                fs.build_sidebar_editor(
                    self.OptionsPanel,
                    section_title, editor_type,
                    c_key, w_key, p_key,
                    fgc, fgp, fgv, bgc, bgp, bgv,
                    state,
                    line_pats or self.line_patterns,
                    fill_pats or self.fill_patterns,
                    on_preview_rebuild,
                    on_settings_cb)
                self.OptionsPopup.IsOpen = True
            except Exception:
                log_exc("on_sidebar_open")

        grid = fs.build_display_row(
            filter_name, settings, idx,
            self.line_patterns, self.fill_patterns,
            self, on_settings_changed,
            on_sidebar_open=on_sidebar_open,
            skip_preview=skip_preview)
        row_data["grid"] = grid

        # Wrap in a DockPanel with a selection checkbox on the left
        from System.Windows.Controls import DockPanel as _DP2
        from System.Windows.Controls import Dock as _Dk2
        from System.Windows.Controls import CheckBox as _CB2
        from System.Windows import Thickness as _T2, VerticalAlignment as _VA2
        from System.Windows.Media import SolidColorBrush as _SB2, Color as _C2

        wrap = _DP2()
        wrap.LastChildFill = True

        sel_cb = _CB2()
        sel_cb.Tag       = filter_name
        sel_cb.IsChecked = False
        sel_cb.Margin    = _T2(2, 0, 6, 0)
        sel_cb.Foreground = _SB2(_C2.FromRgb(244, 250, 255))
        sel_cb.VerticalAlignment = _VA2.Center
        sel_cb.PreviewMouseLeftButtonDown += self._on_tpl_row_mouse_down
        _DP2.SetDock(sel_cb, _Dk2.Left)
        wrap.Children.Add(sel_cb)
        wrap.Children.Add(grid)

        if not hasattr(self, '_tpl_row_checks') or self._tpl_row_checks is None:
            self._tpl_row_checks = []
        self._tpl_row_checks.append(sel_cb)
        row_data["sel_cb"] = sel_cb
        row_data["wrap"]   = wrap

        self.FilterRowsPanel.Children.Add(wrap)
        self._set_empty_state(False)

    # ── TEMPLATE GRID TRI-STATE "ALL" + SHIFT-CLICK ──────────────────────────

    def _build_template_grid_header(self):
        from System.Windows.Controls import CheckBox as _CB3
        from System.Windows import Thickness as _T3
        from System.Windows.Media import SolidColorBrush as _SB3, Color as _C3
        h = _CB3()
        h.Content     = "All"
        h.IsThreeState = True
        h.Foreground  = _SB3(_C3.FromRgb(244, 250, 255))
        h.FontSize    = 11
        h.Margin      = _T3(4, 2, 4, 6)
        h.Click       += self._on_tpl_all_click
        self.FilterRowsPanel.Children.Add(h)
        return h

    def _on_tpl_all_click(self, sender, e):
        try:
            checks = getattr(self, '_tpl_row_checks', [])
            any_ticked = any(cb.IsChecked for cb in checks)
            target = not any_ticked
            self._suppress_tpl_all = True
            try:
                for cb in checks:
                    cb.IsChecked = target
            finally:
                self._suppress_tpl_all = False
            sender.IsChecked = target
            e.Handled = True
            self._refresh_tpl_all_state()
        except Exception:
            log_exc("_on_tpl_all_click")

    def _refresh_tpl_all_state(self):
        h = getattr(self, '_tpl_all_check', None)
        checks = getattr(self, '_tpl_row_checks', [])
        if h is None or not checks:
            return
        ticked = sum(1 for cb in checks if cb.IsChecked)
        self._suppress_tpl_all = True
        try:
            if ticked == 0:               h.IsChecked = False
            elif ticked == len(checks):   h.IsChecked = True
            else:                         h.IsChecked = None
        finally:
            self._suppress_tpl_all = False

    def _on_tpl_row_mouse_down(self, sender, e):
        try:
            from System.Windows.Input import Keyboard, ModifierKeys
            checks = getattr(self, '_tpl_row_checks', [])
            try:   idx = checks.index(sender)
            except ValueError: return

            shift = (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift
            last  = getattr(self, '_tpl_last_clicked', None)
            if not shift or last is None:
                self._tpl_last_clicked = idx
                return

            new_state = not bool(sender.IsChecked)
            lo, hi = min(last, idx), max(last, idx)
            for i in range(lo, hi + 1):
                checks[i].IsChecked = new_state
            self._tpl_last_clicked = idx
            e.Handled = True
            self._refresh_tpl_all_state()
        except Exception:
            log_exc("_on_tpl_row_mouse_down")

    # ── ASSIGN MODE (Views / View Templates) ──────────────────────────────────

    def _on_view_check(self, sender, e):
        try:
            if self._suppress_check_handler:
                return
            name = sender.Tag
            checked = bool(sender.IsChecked)

            if checked:
                self.selected_views.add(name)
            else:
                self.selected_views.discard(name)

            if self.direction == "pull":
                self._rebuild_pull_grid_from_selected(last_added=name if checked else None)
            self._refresh_select_all_state()
            self._update_assign_status()
        except Exception:
            log_exc("_on_view_check")

    def _rebuild_pull_grid_from_selected(self, last_added=None):
        """Multi-source pull: union of filters across all ticked views.

        For filter-name collisions across views, the most recently ticked
        view's overrides win. If last_added is given, its filters overwrite
        anything already in the grid for the same name."""
        self.filter_rows = []
        self.FilterRowsPanel.Children.Clear()
        self.OptionsPanel.Children.Clear()
        self.pull_filter_state = {}
        self._filter_row_checks = []
        self._filter_all_check  = None

        if not self.selected_views:
            self.pull_source_view = None
            self._set_empty_state(True)
            return

        # Read filters from each ticked view. Iteration order: previously
        # ticked first, then last_added last so its overrides win on collision.
        ordered = [n for n in self.selected_views if n != last_added]
        if last_added is not None and last_added in self.selected_views:
            ordered.append(last_added)

        merged = {}   # fname -> {settings, definition, source_view_name}
        for vname in ordered:
            view = self.live_view_map.get(vname)
            if view is None:
                continue
            for r in va.pull_view_filters(view):
                fname = r.get("name")
                if not fname:
                    continue
                merged[fname] = {
                    "settings":   dict(r.get("settings", {})),
                    "definition": r.get("definition", {}),
                    "source":     vname,
                }

        # Track the "primary" source for label display: last-added if any, else
        # the first ticked view alphabetically.
        if last_added:
            self.pull_source_view = self.live_view_map.get(last_added)
        else:
            first = sorted(self.selected_views)[0]
            self.pull_source_view = self.live_view_map.get(first)

        log("Pull grid (multi): {} unique filter(s) across {} view(s)".format(
            len(merged), len(self.selected_views)))

        if not merged:
            self._set_empty_state(True)
            return

        self._filter_all_check = self._build_filter_grid_header()

        for fname, data in sorted(merged.items()):
            st = {
                "assigned":   True,
                "settings":   data["settings"],
                "definition": data["definition"],
            }
            self.pull_filter_state[fname] = st
            self._append_assign_row(fname, st)

        self._refresh_filter_all_state()
        self._set_empty_state(False)

    # ── tri-state "All" + shift-click range select ───────────────────────────

    def _on_select_all_click(self, sender, e):
        """Click fires BEFORE WPF cycles the tri-state. We bypass the cycle
        and just set a definite state based on current ticks: anything ticked
        → untick all; nothing ticked → tick all."""
        try:
            any_ticked = any(self._view_checks[n].IsChecked
                             for n in self._visible_view_names)
            target = not any_ticked  # if anything ticked, clear; else select all
            self._suppress_check_handler = True
            try:
                for n in self._visible_view_names:
                    cb = self._view_checks[n]
                    cb.IsChecked = target
                    if target:
                        self.selected_views.add(n)
                    else:
                        self.selected_views.discard(n)
            finally:
                self._suppress_check_handler = False
            # Set the header to the resolved state, suppress the click toggle.
            sender.IsChecked = target
            e.Handled = True
            if self.direction == "pull":
                self._rebuild_pull_grid_from_selected(last_added=None)
            self._refresh_select_all_state()
            self._update_assign_status()
        except Exception:
            log_exc("_on_select_all_click")

    def _refresh_select_all_state(self):
        if self._select_all_check is None or not self._visible_view_names:
            return
        ticked = sum(1 for n in self._visible_view_names
                     if self._view_checks[n].IsChecked)
        self._suppress_check_handler = True
        try:
            if ticked == 0:
                self._select_all_check.IsChecked = False
            elif ticked == len(self._visible_view_names):
                self._select_all_check.IsChecked = True
            else:
                self._select_all_check.IsChecked = None  # indeterminate
        finally:
            self._suppress_check_handler = False

    def _on_view_row_mouse_down(self, sender, e):
        """Intercept clicks so Shift+click sets a range. Plain click just
        records the anchor index and lets WPF toggle the box normally."""
        try:
            from System.Windows.Input import Keyboard, ModifierKeys
            name = sender.Tag
            try:
                idx = self._visible_view_names.index(name)
            except ValueError:
                return

            shift_held = (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift

            if not shift_held or self._last_clicked_index is None:
                self._last_clicked_index = idx
                return  # let WPF handle the normal toggle

            # Shift-click: apply the clicked checkbox's new state to the range.
            # Current IsChecked is still pre-toggle here; the new state will
            # be its inverse.
            new_state = not bool(sender.IsChecked)
            lo = min(self._last_clicked_index, idx)
            hi = max(self._last_clicked_index, idx)

            self._suppress_check_handler = True
            try:
                for i in range(lo, hi + 1):
                    n = self._visible_view_names[i]
                    cb = self._view_checks[n]
                    if bool(cb.IsChecked) != new_state:
                        cb.IsChecked = new_state
                    if new_state:
                        self.selected_views.add(n)
                    else:
                        self.selected_views.discard(n)
            finally:
                self._suppress_check_handler = False

            self._last_clicked_index = idx
            e.Handled = True  # stop the original toggle, we've handled it
            if self.direction == "pull":
                # Use the highest-index newly-ticked view as the "primary".
                last_name = None
                for i in range(hi, lo - 1, -1):
                    n = self._visible_view_names[i]
                    if n in self.selected_views:
                        last_name = n
                        break
                self._rebuild_pull_grid_from_selected(last_added=last_name)
            self._refresh_select_all_state()
            self._update_assign_status()
        except Exception:
            log_exc("_on_view_row_mouse_down")

    def _load_template_into_assign_grid(self):
        """Push direction: load the source template's filters into the grid,
        all ticked by default. These are the filters Push sends to the checked
        views."""
        self.filter_rows = []
        self.FilterRowsPanel.Children.Clear()
        self.OptionsPanel.Children.Clear()
        self.push_filter_state = {}
        self._filter_row_checks = []
        self._filter_all_check  = None

        src = self.source_template_data
        template_filters = src.get("filters", []) if src else []
        if not template_filters:
            self._set_empty_state(True)
            self._set_status("Template '{}' has no filters.".format(
                self.source_template_name))
            return

        self._filter_all_check = self._build_filter_grid_header()

        for fdata in sorted(template_filters, key=lambda f: f.get("name", "").lower()):
            fname = fdata.get("name")
            if not fname:
                continue
            st = {
                "assigned":   True,
                "settings":   dict(fdata.get("settings", {})),
                "definition": fdata.get("definition", {}),
            }
            self.push_filter_state[fname] = st
            self._append_assign_row(fname, st)

        self._refresh_filter_all_state()
        self._set_empty_state(False)

    # ── FILTER GRID TRI-STATE "ALL" ───────────────────────────────────────────

    def _build_filter_grid_header(self):
        """Add a tri-state 'All' CheckBox above the filter rows panel and return it."""
        from System.Windows.Controls import CheckBox as WpfCheckBox
        from System.Windows import Thickness
        from System.Windows.Media import SolidColorBrush, Color

        header = WpfCheckBox()
        header.Content    = "All"
        header.Foreground = SolidColorBrush(Color.FromRgb(244, 250, 255))
        header.FontSize   = 11
        header.Margin     = Thickness(4, 2, 4, 6)
        header.IsThreeState = True
        header.Click += self._on_filter_all_click
        self.FilterRowsPanel.Children.Add(header)
        return header

    def _on_filter_all_click(self, sender, e):
        try:
            any_ticked = any(cb.IsChecked for cb in self._filter_row_checks)
            target = not any_ticked
            self._suppress_filter_all = True
            try:
                for cb in self._filter_row_checks:
                    cb.IsChecked = target
            finally:
                self._suppress_filter_all = False
            sender.IsChecked = target
            e.Handled = True
            # Update underlying state dicts
            state = self.pull_filter_state if self.direction == "pull" \
                    else self.push_filter_state
            for st in state.values():
                st["assigned"] = target
            self._update_assign_status()
        except Exception:
            log_exc("_on_filter_all_click")

    def _refresh_filter_all_state(self):
        if self._filter_all_check is None or not self._filter_row_checks:
            return
        ticked = sum(1 for cb in self._filter_row_checks if cb.IsChecked)
        self._suppress_filter_all = True
        try:
            if ticked == 0:
                self._filter_all_check.IsChecked = False
            elif ticked == len(self._filter_row_checks):
                self._filter_all_check.IsChecked = True
            else:
                self._filter_all_check.IsChecked = None  # indeterminate
        finally:
            self._suppress_filter_all = False

    def _append_assign_row(self, filter_name, st):
        idx = len(self.filter_rows)
        row_data = {"name": filter_name, "state": st, "grid": None}
        self.filter_rows.append(row_data)

        def on_settings_changed(new_settings):
            st["settings"] = new_settings

        def on_assign_changed(fname, is_assigned):
            st["assigned"] = bool(is_assigned)
            self._refresh_filter_all_state()
            self._update_assign_status()

        def on_sidebar_open(section_title, editor_type,
                            c_key, w_key, p_key,
                            fgc, fgp, fgv, bgc, bgp, bgv,
                            state, line_pats, fill_pats,
                            on_preview_rebuild, on_settings_cb):
            try:
                fs.build_sidebar_editor(
                    self.OptionsPanel,
                    section_title, editor_type,
                    c_key, w_key, p_key,
                    fgc, fgp, fgv, bgc, bgp, bgv,
                    state,
                    line_pats or self.line_patterns,
                    fill_pats or self.fill_patterns,
                    on_preview_rebuild,
                    on_settings_cb)
                self.OptionsPopup.IsOpen = True
            except Exception:
                log_exc("on_sidebar_open(assign)")

        grid = fs.build_display_row(
            filter_name, st.get("settings", {}), idx,
            self.line_patterns, self.fill_patterns,
            self, on_settings_changed,
            on_sidebar_open=on_sidebar_open,
            assign_mode=True,
            assigned=bool(st.get("assigned")),
            on_assign_changed=on_assign_changed)
        row_data["grid"] = grid
        self.FilterRowsPanel.Children.Add(grid)

        # Register the Add checkbox so the tri-state header can track it,
        # and wire shift-click range select.
        try:
            from System.Windows.Controls import CheckBox as _CB
            for child in grid.Children:
                if isinstance(child, _CB):
                    child.PreviewMouseLeftButtonDown += self._on_filter_row_mouse_down
                    self._filter_row_checks.append(child)
                    break
        except Exception:
            pass

    def _on_filter_row_mouse_down(self, sender, e):
        """Shift-click range select for assign grid filter rows."""
        try:
            from System.Windows.Input import Keyboard, ModifierKeys
            checks = self._filter_row_checks
            try:   idx = checks.index(sender)
            except ValueError: return

            shift = (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift
            last  = getattr(self, '_filter_last_clicked', None)
            if not shift or last is None:
                self._filter_last_clicked = idx
                return

            new_state = not bool(sender.IsChecked)
            lo, hi = min(last, idx), max(last, idx)
            state_dict = (self.pull_filter_state
                          if self.direction == "pull" else self.push_filter_state)
            self._suppress_filter_all = True
            try:
                for i in range(lo, hi + 1):
                    checks[i].IsChecked = new_state
                    # Update underlying state
                    fname = checks[i].Tag if hasattr(checks[i], 'Tag') else None
                    if fname and fname in state_dict:
                        state_dict[fname]["assigned"] = new_state
            finally:
                self._suppress_filter_all = False

            self._filter_last_clicked = idx
            e.Handled = True
            self._refresh_filter_all_state()
            self._update_assign_status()
        except Exception:
            log_exc("_on_filter_row_mouse_down")

    def _update_assign_status(self):
        if self.mode not in ("views", "viewtemplates"):
            return
        noun = "view template" if self.mode == "viewtemplates" else "view"

        if self.direction == "push":
            ticked = sum(1 for v in self.push_filter_state.values()
                         if v.get("assigned"))
            total  = len(self.push_filter_state)
            nsel   = len(self.selected_views)
            self.TxtTargetLabel.Text = "Push {}  ->  {} {}{} ticked".format(
                self.source_template_name, nsel, noun, "s" if nsel != 1 else "")
            self._set_status("{} of {} filters ticked. Push sends them to the "
                             "{} ticked {}{}.".format(
                                 ticked, total, nsel, noun,
                                 "s" if nsel != 1 else ""))
        else:  # pull
            nsel  = len(self.selected_views)
            dest  = self.source_template_name or "(new template)"
            ticked = sum(1 for v in self.pull_filter_state.values()
                         if v.get("assigned"))
            total  = len(self.pull_filter_state)
            if nsel == 0:
                src = "(no source)"
            elif nsel == 1:
                src = sorted(self.selected_views)[0]
            else:
                src = "{} {}s".format(nsel, noun)
            self.TxtTargetLabel.Text = "Pull {}  ->  {}".format(src, dest)
            self._set_status("{} of {} filters ticked. Saves to '{}'.".format(
                ticked, total, dest))

    def _on_push_to_views(self):
        try:
            if not self.source_template_data:
                forms.alert("Pick a filter template first under the Templates tab.",
                            title="Filter Manager")
                return
            if not self.selected_views:
                forms.alert("Tick one or more views on the left to push to.",
                            title="Filter Manager")
                return
            ticked = [n for n, v in self.push_filter_state.items()
                      if v.get("assigned")]
            if not ticked:
                forms.alert("No filters ticked. Tick at least one filter to push.",
                            title="Filter Manager")
                return

            # Build a per-view state dict; all views share the same ticked set.
            view_state = {}
            for name in self.selected_views:
                v = self.live_view_map.get(name)
                if v is None:
                    continue
                view_state[v.UniqueId] = {
                    "view":     v,
                    "template": self.source_template_name,
                    "filters":  self.push_filter_state,
                }

            summary = va.commit_assignments(
                view_state, additive=True,
                status_callback=lambda m: (log(m), self._set_status(m)))
            if summary:
                log(summary.replace("\n", " | "), "OK")
                forms.alert(summary, title="Push to Views")
        except Exception:
            log_exc("_on_push_to_views")

    def _on_pull_to_template(self):
        """Save the ticked filters from the pull grid into the currently
        selected template (the one chosen on the Templates tab). If none is
        selected, prompt for a new template name on the spot. Name conflicts
        are resolved OS-style: Overwrite / Rename / Skip, with apply-to-all."""
        try:
            view = self.pull_source_view
            if view is None:
                forms.alert("Tick a view on the left to pull filters from.",
                            title="Pull to Template")
                return

            rows = []
            for fname, st in self.pull_filter_state.items():
                if not st.get("assigned"):
                    continue
                rows.append({
                    "name":       fname,
                    "settings":   st.get("settings", {}),
                    "definition": st.get("definition", {}),
                })
            if not rows:
                forms.alert("No filters ticked to pull.", title="Pull to Template")
                return

            src_name = safe_str(view.Name)
            tpl_name = self.source_template_name

            # No template selected -> prompt for a new one on the spot.
            if not tpl_name:
                tpl_name = forms.ask_for_string(
                    default=src_name,
                    prompt="No template selected. Enter a name for a new template:",
                    title="Pull to Template")
                if not tpl_name:
                    return
                tpl_name = tpl_name.strip()
                illegal = set('/\\:*?"<>|')
                if not tpl_name or any(c in illegal for c in tpl_name):
                    forms.alert("Invalid template name.")
                    return
                ok, result = fs.save_template(self.templates_folder, tpl_name, rows)
                if not ok:
                    log("Pull save failed: {}".format(result), "ERR")
                    forms.alert("Save failed: {}".format(result))
                    return
                self._after_pull_save(tpl_name, src_name, len(rows),
                                      "Saved '{}' with {} filter(s).".format(
                                          tpl_name, len(rows)))
                return

            # Merge into the selected template.
            tpl_data = load_template_file(self.templates_folder, tpl_name)
            if not tpl_data:
                # Template was selected but its file is gone - just save fresh.
                ok, result = fs.save_template(self.templates_folder, tpl_name, rows)
                if not ok:
                    forms.alert("Save failed: {}".format(result))
                    return
                self._after_pull_save(tpl_name, src_name, len(rows),
                                      "Saved '{}' with {} filter(s).".format(
                                          tpl_name, len(rows)))
                return

            existing_names = set(f.get("name") for f in tpl_data.get("filters", []))

            # Resolve name conflicts OS-style.
            resolved_rows, overwrite_names, cancelled = self._resolve_pull_conflicts(
                rows, existing_names, tpl_name)
            if cancelled:
                self._set_status("Pull cancelled.")
                return

            merged, added, updated, skipped = fs.merge_rows_into_template_data(
                tpl_data, resolved_rows, overwrite_names)
            ok, result = fs.write_template_data(
                self.templates_folder, tpl_name, merged)
            if not ok:
                log("Pull merge save failed: {}".format(result), "ERR")
                forms.alert("Save failed: {}".format(result))
                return

            msg = ("Merged into '{}': {} added, {} updated, {} skipped.".format(
                tpl_name, added, updated, skipped))
            self._after_pull_save(tpl_name, src_name, len(rows), msg)
        except Exception:
            log_exc("_on_pull_to_template")

    def _resolve_pull_conflicts(self, rows, existing_names, tpl_name):
        """OS-style conflict resolution.

        For each row whose name collides with an existing filter in the
        destination template, prompt: Overwrite / Rename / Skip / Cancel,
        with an 'Apply to all remaining' option that reuses the last choice.

        Returns (final_rows, overwrite_names_set, cancelled_bool). Renamed
        rows have their 'name' key updated; skipped rows are dropped.
        """
        final_rows = []
        overwrite  = set()
        bulk_choice = None  # remembers "apply to all" decision

        # Build a "taken" set we extend as we rename, so two renames can't
        # collide with each other.
        taken = set(existing_names)
        # Also reserve the non-conflicting incoming names.
        for r in rows:
            if r["name"] not in existing_names:
                taken.add(r["name"])

        remaining_conflicts = sum(1 for r in rows if r["name"] in existing_names)

        for r in rows:
            fname = r["name"]
            if fname not in existing_names:
                final_rows.append(r)
                continue

            decision = bulk_choice
            new_name = None

            if decision is None:
                decision, new_name = self._ask_conflict(
                    fname, tpl_name, remaining_conflicts, taken)
                if decision is None:
                    return [], set(), True  # cancelled

            if decision == "overwrite":
                overwrite.add(fname)
                final_rows.append(r)
            elif decision == "rename":
                # If bulk-applied, auto-derive unique names instead of asking.
                if new_name is None:
                    new_name = self._auto_unique_name(fname, taken)
                taken.add(new_name)
                renamed = dict(r)
                renamed["name"] = new_name
                final_rows.append(renamed)
            elif decision == "skip":
                pass  # drop the row
            # else: shouldn't happen

            remaining_conflicts -= 1

            # Re-confirm bulk if user picked it during the dialog.
            # _ask_conflict returns (decision, new_name) but signals bulk via
            # a sentinel; we read it from an attribute set there.
            if getattr(self, "_pull_apply_all", False):
                bulk_choice = decision
                self._pull_apply_all = False

        return final_rows, overwrite, False

    def _auto_unique_name(self, base, taken):
        i = 2
        while True:
            candidate = "{} ({})".format(base, i)
            if candidate not in taken:
                return candidate
            i += 1

    def _ask_conflict(self, fname, tpl_name, remaining, taken):
        """OS-style conflict dialog. Returns (decision, new_name_or_None) or
        (None, None) if cancelled. Sets self._pull_apply_all if the user
        ticks 'apply to all remaining'."""
        from System.Windows import Window, Thickness, HorizontalAlignment, TextWrapping
        from System.Windows.Controls import (StackPanel, TextBlock, Button,
                                             TextBox, CheckBox, Orientation)
        from System.Windows.Media import SolidColorBrush, Color

        bg = SolidColorBrush(Color.FromRgb(43, 51, 64))
        fg = SolidColorBrush(Color.FromRgb(244, 250, 255))
        muted = SolidColorBrush(Color.FromRgb(156, 163, 175))

        win = Window()
        win.Title = "Filter exists"
        win.Width = 480
        win.SizeToContent = 2  # Height (enum value 2 = SizeToContent.Height)
        win.MinHeight = 240
        win.ResizeMode = 1  # NoResize
        win.Background = bg
        win.Foreground = fg
        try:
            win.Owner = self
        except Exception:
            pass

        root = StackPanel()
        root.Margin = Thickness(18)

        title = TextBlock()
        title.Text = "Filter '{}' already exists in template '{}'.".format(
            fname, tpl_name)
        title.Foreground = fg
        title.FontSize = 13
        title.TextWrapping = TextWrapping.Wrap
        title.Margin = Thickness(0, 0, 0, 10)
        root.Children.Add(title)

        if remaining > 1:
            sub = TextBlock()
            sub.Text = "{} conflict(s) remaining.".format(remaining)
            sub.Foreground = muted
            sub.FontSize = 11
            sub.Margin = Thickness(0, 0, 0, 8)
            root.Children.Add(sub)

        rename_label = TextBlock()
        rename_label.Text = "Rename to:"
        rename_label.Foreground = fg
        rename_label.FontSize = 11
        rename_label.Margin = Thickness(0, 4, 0, 2)
        root.Children.Add(rename_label)

        rename_box = TextBox()
        rename_box.Text = self._auto_unique_name(fname, taken)
        rename_box.Height = 26
        rename_box.Margin = Thickness(0, 0, 0, 10)
        try:
            rename_box.Style = self.Resources["ModernTextBoxStyle"]
        except Exception:
            pass
        root.Children.Add(rename_box)

        apply_all = CheckBox()
        apply_all.Content = "Apply to all remaining conflicts"
        apply_all.Foreground = fg
        apply_all.FontSize = 11
        apply_all.Margin = Thickness(0, 4, 0, 14)
        apply_all.IsEnabled = (remaining > 1)
        root.Children.Add(apply_all)

        btn_row = StackPanel()
        btn_row.Orientation = Orientation.Horizontal
        btn_row.HorizontalAlignment = HorizontalAlignment.Right

        result = {"decision": None, "new_name": None}

        def make_btn(text, decision_value, style_key="SecondaryButtonStyle"):
            b = Button()
            b.Content = text
            b.Margin = Thickness(6, 0, 0, 0)
            b.MinWidth = 88
            try:
                b.Style = self.Resources[style_key]
            except Exception:
                pass
            def on_click(s, e):
                if decision_value == "rename":
                    name = (rename_box.Text or "").strip()
                    illegal = set('/\\:*?"<>|')
                    if not name or any(c in illegal for c in name) \
                       or name == fname or name in taken:
                        from pyrevit import forms as _f
                        _f.alert("Pick a different, valid name.")
                        return
                    result["new_name"] = name
                result["decision"] = decision_value
                self._pull_apply_all = bool(apply_all.IsChecked)
                win.DialogResult = True
                win.Close()
            b.Click += on_click
            return b

        btn_row.Children.Add(make_btn("Skip",      "skip"))
        btn_row.Children.Add(make_btn("Rename",    "rename"))
        btn_row.Children.Add(make_btn("Overwrite", "overwrite", "ModernButtonStyle"))

        cancel_btn = Button()
        cancel_btn.Content = "Cancel all"
        cancel_btn.Margin = Thickness(6, 0, 0, 0)
        cancel_btn.MinWidth = 88
        try:
            cancel_btn.Style = self.Resources["SecondaryButtonStyle"]
        except Exception:
            pass
        def on_cancel(s, e):
            result["decision"] = None
            win.DialogResult = False
            win.Close()
        cancel_btn.Click += on_cancel
        btn_row.Children.Add(cancel_btn)

        root.Children.Add(btn_row)
        win.Content = root
        win.ShowDialog()
        return result["decision"], result["new_name"]

    def _after_pull_save(self, tpl_name, src_name, count, alert_msg):
        """Shared cleanup after a successful pull save: set as active push
        source, trigger sync, refresh sidebar, alert."""
        log("Pull save OK: {}".format(alert_msg), "OK")
        self.source_template_name = tpl_name
        self.source_template_data = load_template_file(
            self.templates_folder, tpl_name)
        sm.sync_after_save(tpl_name, self.templates_folder, logger=log)
        self._set_status(alert_msg)
        forms.alert(alert_msg, title="Pull to Template")

    # ── EVENT HANDLERS ────────────────────────────────────────────────────────

    def _on_template_click(self, sender, e):
        try:
            name = sender.Tag
            log("Template: {}".format(name))
            # Clicking the already-active template does nothing
            if name == self.active_template:
                return
            if self._template_dirty:
                choice = self._unsaved_changes_dialog(
                    self.active_template or "template")
                if choice == "Cancel" or choice is None:
                    return
                if choice == "Save":
                    ok, result = fs.save_template(
                        self.templates_folder, self.active_template, self.filter_rows)
                    if ok:
                        self._template_dirty = False
                        log("Saved: {}".format(result), "OK")
                    else:
                        forms.alert("Save failed: {}".format(result))
                        return
                # "Discard" falls through
            self._open_template(name)
        except Exception:
            log_exc("_on_template_click")

    def _on_new_template(self, sender, e):
        try:
            log("New template")
            name = forms.ask_for_string(
                default="",
                prompt="Name for the new template:",
                title="New Template")
            if not name:
                return
            name = name.strip()
            illegal = set('/\\:*?"<>|')
            if not name or any(c in illegal for c in name):
                forms.alert("Invalid template name.")
                return
            existing = set(list_templates(self.templates_folder))
            if name in existing:
                forms.alert("Template '{}' already exists.".format(name),
                            title="New Template")
                return
            # Create an empty template file on disk.
            ok, result = fs.save_template(self.templates_folder, name, [])
            if not ok:
                log("Create failed: {}".format(result), "ERR")
                forms.alert("Could not create template: {}".format(result))
                return
            log("Created: {}".format(result), "OK")
            sm.sync_after_save(name, self.templates_folder, logger=log)
            self._refresh_sidebar()
            self._open_template(name)
            self._set_status("New template '{}' created. Use + Add Filter.".format(name))
        except Exception:
            log_exc("_on_new_template")

    def _on_load_demo(self, sender, e):
        """Copy the shipped example templates in, on request.

        The way back from seed_once(): those only ever install on first run,
        so a demo the user deleted is otherwise gone for good. Additive and
        never overwrites, so pressing it twice, or after editing a demo, is
        harmless.
        """
        try:
            copied = _userdata.seed_from_defaults(DEFAULTS_DIR,
                                                  self.templates_folder)
            log("Load demo: copied {}".format(copied))
            self._refresh_sidebar()
            if copied:
                self._set_status(
                    "Added {} example template{}.".format(
                        copied, "" if copied == 1 else "s"))
            else:
                self._set_status("Example templates are already loaded.")
        except Exception:
            log_exc("_on_load_demo")

    def _on_add_filters(self, sender, e):
        try:
            existing = {r["name"] for r in self.filter_rows}
            seen     = set()
            available = []
            for f in fs.get_all_filters():
                n = f.Name
                if n not in existing and n not in seen:
                    seen.add(n)
                    available.append(n)
            available.sort()
            log("Add: {} available".format(len(available)))
            if not available:
                forms.alert("All filters already in template.")
                return

            selected = self._show_add_filter_dialog(available)
            log("Selected: {}".format(selected))
            if not selected:
                return

            target_template = self.active_template  # snapshot before dialog
            filter_map = {f.Name: f for f in fs.get_all_filters()}
            for name in selected:
                # If the user switched templates while the dialog was open, stop
                if self.active_template != target_template:
                    log("Template switched during add — stopping.", "WARN")
                    break
                f    = filter_map.get(name)
                defn = fs.serialise_filter_def(f) if f else {}
                self._append_row(name, {}, defn)
                log("Added: {}".format(name))
            self._set_status("Added {}.".format(len(selected)))
        except Exception:
            log_exc("_on_add_filters")

    def _confirm(self, message, title="pyFilter"):
        """Seed43-styled yes/no confirm. Returns True for Yes."""
        from System.Windows import Window, Thickness, HorizontalAlignment, \
            VerticalAlignment, TextWrapping, WindowStartupLocation, ResizeMode
        from System.Windows.Controls import StackPanel, TextBlock, Button
        from System.Windows.Controls import Orientation as _Or
        from System.Windows.Media import SolidColorBrush, Color

        bg    = SolidColorBrush(Color.FromRgb(43,  51, 64))
        fg    = SolidColorBrush(Color.FromRgb(244, 250, 255))
        green = SolidColorBrush(Color.FromRgb(32,  138, 60))
        grey  = SolidColorBrush(Color.FromRgb(64,  69,  83))

        win = Window()
        win.Title      = title
        win.Width      = 360
        win.Height     = 160
        win.Background = bg
        win.Foreground = fg
        win.ResizeMode = ResizeMode.NoResize
        win.WindowStartupLocation = WindowStartupLocation.CenterOwner
        try: win.Owner = self
        except Exception: pass

        root = StackPanel()
        root.Margin = Thickness(24, 20, 24, 20)

        msg_tb = TextBlock()
        msg_tb.Text         = message
        msg_tb.Foreground   = fg
        msg_tb.FontSize     = 13
        msg_tb.TextWrapping = TextWrapping.Wrap
        msg_tb.Margin       = Thickness(0, 0, 0, 20)
        root.Children.Add(msg_tb)

        btn_row = StackPanel()
        btn_row.Orientation         = _Or.Horizontal
        btn_row.HorizontalAlignment = HorizontalAlignment.Right

        result = [False]

        for label, color, answer in [("Cancel", grey, False), ("Yes", green, True)]:
            b = Button()
            b.Content         = label
            b.Foreground      = fg
            b.Background      = color
            b.BorderThickness = Thickness(0)
            b.Padding         = Thickness(20, 8, 20, 8)
            b.FontSize        = 12
            b.Margin          = Thickness(8, 0, 0, 0)
            def _click(s, e, ans=answer):
                result[0] = ans
                win.Close()
            b.Click += _click
            btn_row.Children.Add(b)

        root.Children.Add(btn_row)
        win.Content = root
        win.ShowDialog()
        return result[0]

    def _unsaved_changes_dialog(self, template_name):
        """Three-button: Save / Discard / Cancel. Returns the chosen string."""
        from System.Windows import Window, Thickness, HorizontalAlignment, \
            TextWrapping, WindowStartupLocation, ResizeMode
        from System.Windows.Controls import StackPanel, TextBlock, Button
        from System.Windows.Controls import Orientation as _Or2
        from System.Windows.Media import SolidColorBrush, Color

        bg    = SolidColorBrush(Color.FromRgb(43,  51, 64))
        fg    = SolidColorBrush(Color.FromRgb(244, 250, 255))
        green = SolidColorBrush(Color.FromRgb(32,  138, 60))
        grey  = SolidColorBrush(Color.FromRgb(64,  69,  83))
        red   = SolidColorBrush(Color.FromRgb(176, 55,  55))

        win = Window()
        win.Title     = "Unsaved Changes"
        win.Width     = 380
        win.Height    = 170
        win.Background = bg
        win.Foreground = fg
        win.ResizeMode = ResizeMode.NoResize
        win.WindowStartupLocation = WindowStartupLocation.CenterOwner
        try: win.Owner = self
        except Exception: pass

        root = StackPanel()
        root.Margin = Thickness(24, 20, 24, 20)

        msg = TextBlock()
        msg.Text        = u"Save changes to '{}' before switching?".format(template_name)
        msg.Foreground  = fg
        msg.FontSize    = 13
        msg.TextWrapping = TextWrapping.Wrap
        msg.Margin      = Thickness(0, 0, 0, 20)
        root.Children.Add(msg)

        btn_row = StackPanel()
        btn_row.Orientation         = _Or2.Horizontal
        btn_row.HorizontalAlignment = HorizontalAlignment.Right

        result = ["Cancel"]

        for label, color, val in [
                ("Cancel",  grey,  "Cancel"),
                ("Discard", red,   "Discard"),
                ("Save",    green, "Save")]:
            b = Button()
            b.Content         = label
            b.Foreground      = fg
            b.Background      = color
            b.BorderThickness = Thickness(0)
            b.Padding         = Thickness(18, 8, 18, 8)
            b.FontSize        = 12
            b.Margin          = Thickness(8, 0, 0, 0)
            def _click(s, e, v=val):
                result[0] = v
                win.Close()
            b.Click += _click
            btn_row.Children.Add(b)

        root.Children.Add(btn_row)
        win.Content = root
        win.ShowDialog()
        return result[0]


    def _show_add_filter_dialog(self, available):
        """Add-filter picker — styled to match the main window."""
        from System.Windows.Markup import XamlReader
        from System.Windows import Window
        import System

        DIALOG_XAML = u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Add Filters"
        Width="440" Height="600"
        MinWidth="320" MinHeight="300"
        Background="#3B4553"
        WindowStartupLocation="CenterOwner"
        ResizeMode="CanResizeWithGrip"
        ShowInTaskbar="False">
  <Window.Resources>

    <!-- Mac-style thin scrollbar thumb -->
    <Style x:Key="MacScrollBarThumb" TargetType="Thumb">
      <Setter Property="Background" Value="#208A3C"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Thumb">
            <Border x:Name="thumb" Background="{TemplateBinding Background}"
                    CornerRadius="3" Margin="2,2,2,2"/>
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

    <!-- Mac-style thin scrollbar -->
    <Style x:Key="MacScrollBar" TargetType="ScrollBar">
      <Setter Property="Orientation" Value="Vertical"/>
      <Setter Property="Width"       Value="8"/>
      <Setter Property="MinWidth"    Value="8"/>
      <Setter Property="Background"  Value="Transparent"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ScrollBar">
            <Grid x:Name="GridRoot" Width="8">
              <Track x:Name="PART_Track" Orientation="Vertical" IsDirectionReversed="True">
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

    <!-- Mac-style ScrollViewer -->
    <Style x:Key="MacScrollViewer" TargetType="ScrollViewer">
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
                                      CanContentScroll="{TemplateBinding CanContentScroll}"/>
              <ScrollBar x:Name="PART_VerticalScrollBar"
                         Grid.Column="1" Orientation="Vertical"
                         Style="{StaticResource MacScrollBar}"
                         Value="{TemplateBinding VerticalOffset}"
                         Maximum="{TemplateBinding ScrollableHeight}"
                         ViewportSize="{TemplateBinding ViewportHeight}"
                         Visibility="{TemplateBinding ComputedVerticalScrollBarVisibility}"/>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- White rounded search box with icon -->
    <Style x:Key="SearchBoxStyle" TargetType="TextBox">
      <Setter Property="Background"               Value="#F4FAFF"/>
      <Setter Property="Foreground"               Value="#2B3340"/>
      <Setter Property="BorderBrush"              Value="#208A3C"/>
      <Setter Property="BorderThickness"          Value="1"/>
      <Setter Property="Padding"                  Value="8,6"/>
      <Setter Property="FontSize"                 Value="12"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border x:Name="Bd"
                    Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="6"
                    Padding="{TemplateBinding Padding}">
              <ScrollViewer x:Name="PART_ContentHost" Focusable="False"
                            HorizontalScrollBarVisibility="Hidden"
                            VerticalScrollBarVisibility="Hidden"
                            VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsFocused" Value="True">
                <Setter TargetName="Bd" Property="BorderBrush"     Value="#2B933F"/>
                <Setter TargetName="Bd" Property="BorderThickness" Value="2"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Green rounded button -->
    <Style x:Key="GreenBtn" TargetType="Button">
      <Setter Property="Background"      Value="#208A3C"/>
      <Setter Property="Foreground"      Value="#F4FAFF"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding"         Value="18,5"/>
      <Setter Property="FontSize"        Value="12"/>
      <Setter Property="FontWeight"      Value="SemiBold"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Bd" Background="{TemplateBinding Background}"
                    CornerRadius="6" Padding="{TemplateBinding Padding}">
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
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Grey rounded button -->
    <Style x:Key="GreyBtn" TargetType="Button">
      <Setter Property="Background"      Value="#404553"/>
      <Setter Property="Foreground"      Value="#F4FAFF"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding"         Value="18,5"/>
      <Setter Property="FontSize"        Value="12"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Bd" Background="{TemplateBinding Background}"
                    CornerRadius="6" Padding="{TemplateBinding Padding}">
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
        </Setter.Value>
      </Setter>
    </Style>

    <!-- CheckBox style matching main window -->
    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="#F4FAFF"/>
      <Setter Property="FontSize"   Value="12"/>
      <Setter Property="Margin"     Value="0,3,0,3"/>
      <Setter Property="Cursor"     Value="Hand"/>
    </Style>

  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- Main content -->
    <DockPanel Grid.Row="0" Margin="16,16,16,0" LastChildFill="True">

      <!-- Heading -->
      <TextBlock DockPanel.Dock="Top"
                 Text="Add Filters"
                 Foreground="#208A3C"
                 FontSize="15" FontWeight="Bold"
                 Margin="0,0,0,10"/>

      <!-- Search box with icon -->
      <Border DockPanel.Dock="Top"
              Background="#F4FAFF"
              BorderBrush="#208A3C"
              BorderThickness="1"
              CornerRadius="6"
              Padding="8,0"
              Margin="0,0,0,8">
        <DockPanel>
          <TextBlock DockPanel.Dock="Left"
                     Text="&#x1F50D;"
                     Foreground="#9CA3AF"
                     FontSize="12"
                     VerticalAlignment="Center"
                     Margin="0,0,6,0"/>
          <TextBox x:Name="SearchBox"
                   Background="Transparent"
                   Foreground="#2B3340"
                   BorderThickness="0"
                   Padding="0,6"
                   FontSize="12"
                   VerticalContentAlignment="Center"/>
        </DockPanel>
      </Border>

      <!-- All / None row -->
      <StackPanel DockPanel.Dock="Top"
                  Orientation="Horizontal"
                  Margin="0,0,0,10">
        <Button x:Name="BtnAll"  Content="All"
                Style="{StaticResource GreenBtn}"
                Padding="16,5" Margin="0,0,8,0"/>
        <Button x:Name="BtnNone" Content="None"
                Style="{StaticResource GreyBtn}"
                Padding="16,5"/>
      </StackPanel>

      <!-- Filter list -->
      <Border Background="#1E2530"
              BorderBrush="#404553"
              BorderThickness="1"
              CornerRadius="6"
              Padding="8">
        <ScrollViewer x:Name="ListScroll"
                      Style="{StaticResource MacScrollViewer}">
          <StackPanel x:Name="ListPanel"/>
        </ScrollViewer>
      </Border>
    </DockPanel>

    <!-- Bottom button bar with resize grip -->
    <Grid Grid.Row="1">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="10"/>
      </Grid.RowDefinitions>

      <!-- Buttons -->
      <StackPanel Grid.Row="0"
                  Orientation="Horizontal"
                  HorizontalAlignment="Right"
                  Margin="0,8,16,0">
        <Button x:Name="BtnCancel" Content="Cancel"
                Style="{StaticResource GreyBtn}"
                Margin="0,0,8,0"/>
        <Button x:Name="BtnAdd" Content="Add"
                Style="{StaticResource GreenBtn}"/>
      </StackPanel>

    </Grid>
  </Grid>
</Window>"""

        try:
            win = XamlReader.Parse(DIALOG_XAML)
        except Exception as ex:
            log("Add filter dialog XAML error: {}".format(ex), "ERR")
            return []

        try:
            win.Owner = self
        except Exception:
            pass

        from System.Windows import Visibility as _Vis
        from System.Windows.Controls import CheckBox as WpfCB

        search_box  = win.FindName("SearchBox")
        list_panel  = win.FindName("ListPanel")
        btn_all     = win.FindName("BtnAll")
        btn_none    = win.FindName("BtnNone")
        btn_cancel  = win.FindName("BtnCancel")
        btn_add     = win.FindName("BtnAdd")

        # Populate checkboxes
        cb_checks = {}
        for name in available:
            cb = WpfCB()
            cb.Content   = name
            cb.Tag       = name
            cb.IsChecked = False
            list_panel.Children.Add(cb)
            cb_checks[name] = cb

        # Wire search
        def _on_search(s, e):
            txt = (search_box.Text or "").lower().strip()
            for n, cb in cb_checks.items():
                cb.Visibility = _Vis.Visible if txt in n.lower() else _Vis.Collapsed
        search_box.TextChanged += _on_search

        # Wire All / None
        def _all(s, e):
            for cb in cb_checks.values():
                if cb.Visibility == _Vis.Visible:
                    cb.IsChecked = True
        def _none(s, e):
            for cb in cb_checks.values():
                cb.IsChecked = False
        btn_all.Click  += _all
        btn_none.Click += _none

        # Wire Cancel / Add
        result = {"names": []}
        def _ok(s, e):
            result["names"] = [n for n, cb in cb_checks.items() if cb.IsChecked]
            win.DialogResult = True
            win.Close()
        def _cancel(s, e):
            win.DialogResult = False
            win.Close()
        btn_add.Click    += _ok
        btn_cancel.Click += _cancel

        win.ShowDialog()
        return result["names"]

    def _on_pull_from_view(self, sender, e):
        try:
            if self.mode in ("views", "viewtemplates"):
                self._on_pull_to_template()
                return
            SUPPORTED = {
                DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan,
                DB.ViewType.Elevation, DB.ViewType.Section,
                DB.ViewType.Detail,    DB.ViewType.ThreeD,
                DB.ViewType.EngineeringPlan, DB.ViewType.AreaPlan,
            }
            view_map = {}
            for v in DB.FilteredElementCollector(doc).OfClass(DB.View):
                try:
                    if v.IsTemplate:
                        view_map["[Template] " + safe_str(v.Name)] = v
                    elif v.ViewType in SUPPORTED:
                        view_map[safe_str(v.Name)] = v
                except Exception:
                    continue

            log("Pull: {} views".format(len(view_map)))
            chosen = forms.SelectFromList.show(
                sorted(view_map.keys()), title="Pull VG Settings From View",
                button_name="Pull", multiselect=False, owner=self)
            log("Chosen: {}".format(chosen))
            if not chosen:
                return

            source     = view_map[chosen]
            all_filters = fs.get_all_filters()
            fid_to_f   = {}
            for f in all_filters:
                val = f.Id.Value if hasattr(f.Id, "Value") else f.Id.IntegerValue
                fid_to_f[val] = f

            try:   source_fids = list(source.GetFilters())
            except Exception: source_fids = []
            log("Filters on view: {}".format(len(source_fids)))

            existing   = {r["name"] for r in self.filter_rows}
            added = updated = 0

            for fid in source_fids:
                fid_val = fid.Value if hasattr(fid, "Value") else fid.IntegerValue
                f = fid_to_f.get(fid_val)
                if not f:
                    log("Unknown fid: {}".format(fid_val), "WARN")
                    continue
                settings = fs.get_filter_settings_from_view(f, source)
                log("Filter: {} en={} vis={}".format(
                    f.Name, settings.get("enabled"), settings.get("visible")))

                if f.Name not in existing:
                    defn = fs.serialise_filter_def(f)
                    self._append_row(f.Name, settings, defn)
                    existing.add(f.Name)
                    added += 1
                    log("Added from view: {}".format(f.Name), "OK")
                else:
                    for row in self.filter_rows:
                        if row["name"] == f.Name:
                            row["settings"] = settings
                            # Rebuild display row with new settings
                            old_grid = row["grid"]
                            idx = self.filter_rows.index(row)

                            def make_cb(r):
                                def cb(new_s):
                                    r["settings"] = new_s
                                return cb

                            new_grid = fs.build_display_row(
                                f.Name, settings, idx,
                                self.line_patterns, self.fill_patterns,
                                self, make_cb(row),
                                on_sidebar_open=None)
                            row["grid"] = new_grid
                            panel_idx = self.FilterRowsPanel.Children.IndexOf(old_grid)
                            if panel_idx >= 0:
                                self.FilterRowsPanel.Children.RemoveAt(panel_idx)
                                self.FilterRowsPanel.Children.Insert(panel_idx, new_grid)
                            updated += 1
                            log("Updated: {}".format(f.Name), "OK")
                            break

            log("Done: added={} updated={}".format(added, updated), "OK")
            self._set_status("Pulled '{}': {} added, {} updated.".format(
                chosen, added, updated))
        except Exception:
            log_exc("_on_pull_from_view")

    def _on_remove_filter(self, sender, e):
        try:
            if not self.filter_rows:
                return

            if self.mode in ("views", "viewtemplates"):
                # Assign mode: show picker (no checkboxes in assign grid)
                chosen = forms.SelectFromList.show(
                    [r["name"] for r in self.filter_rows],
                    title="Remove Filters", button_name="Remove",
                    multiselect=True, owner=self)
                if not chosen:
                    return
            else:
                # Templates mode: use checked rows
                checks = getattr(self, '_tpl_row_checks', [])
                chosen = [r["name"] for r, cb in zip(self.filter_rows, checks)
                          if cb.IsChecked]
                if not chosen:
                    forms.alert("Tick the filters you want to delete first.")
                    return

            if not self._confirm(
                    "Delete {} filter(s)?".format(len(chosen)),
                    title="Delete Filters"):
                return

            chosen_set = set(chosen)
            kept_rows  = []
            kept_cbs   = []
            self.FilterRowsPanel.Children.Clear()

            # Re-add All header
            if self.mode not in ("views", "viewtemplates"):
                self._tpl_all_check  = self._build_template_grid_header()
                self._tpl_row_checks = []

            checks = getattr(self, '_tpl_row_checks', [])
            orig_checks = list(getattr(self, '_tpl_row_checks', []))

            # Rebuild from filter_rows
            for row in self.filter_rows:
                if row["name"] not in chosen_set:
                    kept_rows.append(row)
                    wrap = row.get("wrap") or row.get("grid")
                    if wrap is not None:
                        self.FilterRowsPanel.Children.Add(wrap)
                    sel_cb = row.get("sel_cb")
                    if sel_cb is not None:
                        kept_cbs.append(sel_cb)

            self.filter_rows = kept_rows
            self._tpl_row_checks = kept_cbs
            self._tpl_last_clicked = None
            self._set_empty_state(len(kept_rows) == 0)
            self._template_dirty = True
            log("Removed: {}".format(", ".join(chosen_set)), "WARN")
            self._set_status("Removed {}.".format(len(chosen_set)))
        except Exception:
            log_exc("_on_remove_filter")

    def _on_apply(self, sender, e):
        try:
            if self.mode in ("views", "viewtemplates"):
                if self.direction == "pull":
                    self._on_pull_to_template()
                else:
                    self._on_push_to_views()
                return
            # Templates mode: Save only — apply to views is a separate operation.
            if not self.active_template:
                forms.alert("Pick a template first (or use + New Template).",
                            title="Save")
                return
            ok, result = fs.save_template(
                self.templates_folder, self.active_template, self.filter_rows)
            if not ok:
                log("Save failed: {}".format(result), "ERR")
                forms.alert("Save failed: {}".format(result))
                return
            log("Saved: {}".format(result), "OK")
            self._template_dirty = False
            sm.sync_after_save(self.active_template, self.templates_folder, logger=log)
            self._set_status("Saved '{}'.".format(self.active_template))
        except Exception:
            log_exc("_on_apply")

    # ── HELPERS ───────────────────────────────────────────────────────────────

    def _set_status(self, msg):
        try: self.StatusText.Text = msg
        except Exception: pass

    def _set_empty_state(self, empty):
        try:
            self.EmptyState.Visibility = (
                Visibility.Visible if empty else Visibility.Collapsed)
        except Exception: pass

    def RowsScroll_ScrollChanged(self, sender, e):
        try:
            offset = self.RowsScroll.HorizontalOffset
            self.HeaderScroll.ScrollToHorizontalOffset(offset)
            self.GroupHeaderScroll.ScrollToHorizontalOffset(offset)
        except Exception: pass

# ── ENTRY ─────────────────────────────────────────────────────────────────────

def main():
    log("Filter Manager starting", "OK")
    try:
        win = pyFilterWindow()
        win.ShowDialog()
        log("Closed")
    except Exception:
        log_exc("startup")

if __name__ == "__main__":
    main()
