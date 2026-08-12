# -*- coding: utf-8 -*-
# "Batch Upgrade"
# "Seed43"
# """
# Save upgraded copies of a pile of Revit files, in one or more target Revit
# versions - either right now, or once, at a scheduled date and time.
#
# Revit can only ever save in its own version, so anything other than the
# running version means launching that Revit and having it do the work
# itself, then close - see tools/job_io.py and tools/headless_batch.py,
# and Seed43.extension/startup.py for the pickup side. The actual batch
# loop (tools/batch_runner.py) is shared with startup.py's scheduled-run
# handler, so a scheduled run behaves identically to clicking Upgrade -
# just without anyone watching.
# """

# ── IMPORTS ────────────────────────────────────────────────────────────────

import os
import sys
import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from pyrevit import forms, HOST_APP
from pyrevit.framework import Windows

from Snippets import _dialogs as dlg
from Snippets._icons import set_header_icon
from Snippets.seed43_theme import apply_seed43_palette, apply_seed43_dimensions

from tools import batch_runner
from tools import file_scan
from tools import revit_versions
from tools import schedule_io
from tools import templates
from tools import upgrade_core


# ── CONSTANTS ──────────────────────────────────────────────────────────────

XAML_FILE = os.path.join(SCRIPT_DIR, "BatchUpgrade.xaml")

FILE_FILTER = ("Revit Files (*.rvt, *.rfa, *.rte, *.rft)|"
               "*.rvt;*.rfa;*.rte;*.rft")


# ── HELPERS ────────────────────────────────────────────────────────────────

def running_year():
    """Return the Revit version number this script is running inside."""
    return int(str(HOST_APP.app.VersionNumber)[:4])


# ── CLASSES ────────────────────────────────────────────────────────────────

class BatchUpgradeWindow(forms.WPFWindow):
    """Picks the files, the output folder and the target Revit versions."""

    # --- construction ---
    def __init__(self, xaml_path, host_year):
        forms.WPFWindow.__init__(self, xaml_path)
        self.host_year = host_year
        self.records = []
        self.out_dir = None
        self.targets = []
        self.confirmed = False
        self._boxes = {}
        self._loading = False

        self._bind()
        apply_seed43_palette(self, SCRIPT_DIR)
        apply_seed43_dimensions(self, SCRIPT_DIR)
        set_header_icon(self, SCRIPT_DIR)
        self._build_version_grid()
        self._setup_templates()
        self._setup_schedule_time()
        self._refresh_schedule_status()
        self._refresh()

    # --- public methods ---
    def show(self):
        self.ShowDialog()
        return self.confirmed

    # --- private helpers: wiring ---
    def _bind(self):
        self.FindName("win_close_btn").Click += self._on_cancel
        self.FindName("cancel_btn").Click += self._on_cancel
        self.FindName("upgrade_btn").Click += self._on_upgrade
        self.FindName("add_files_btn").Click += self._on_add_files
        self.FindName("add_folder_btn").Click += self._on_add_folder
        self.FindName("clear_btn").Click += self._on_clear
        self.FindName("browse_btn").Click += self._on_browse
        self.FindName("template_cb").SelectionChanged += self._on_template_selected
        self.FindName("template_save_btn").Click += self._on_template_save
        self.FindName("template_delete_btn").Click += self._on_template_delete
        self.FindName("schedule_btn").Click += self._on_schedule
        self.FindName("schedule_cancel_btn").Click += self._on_schedule_cancel

    def _build_version_grid(self):
        """One checkbox per Revit year, greyed out when it isn't installed."""
        self.running_lbl.Text = "Running: {}".format(self.host_year)
        panel = self.version_panel
        panel.Children.Clear()
        self._boxes = {}

        for year, installed, is_running in revit_versions.version_grid(self.host_year):
            box = Windows.Controls.CheckBox()
            box.Style = self.FindResource("VersionCheck")
            label = ("Revit {}".format(year) if installed
                     else "Revit {} (not installed)".format(year))
            box.Content = label
            box.IsEnabled = installed
            box.IsChecked = bool(is_running and installed)
            box.Tag = year
            box.Checked += self._on_target_changed
            box.Unchecked += self._on_target_changed
            panel.Children.Add(box)
            self._boxes[year] = box

    # --- private helpers: state ---
    def _selected_targets(self):
        return sorted([year for year, box in self._boxes.items()
                       if box.IsEnabled and box.IsChecked])

    def _add_paths(self, paths):
        known = set(r["path"].lower() for r in self.records)
        for path in paths:
            if path.lower() in known:
                continue
            self.records.append(file_scan.scan_path(path))
            known.add(path.lower())
        self.records.sort(key=lambda r: r["name"].lower())

    def _refresh(self):
        self._refresh_list()
        self._refresh_warning()
        self.count_lbl.Text = "{} file(s)".format(len(self.records))
        self.upgrade_btn.IsEnabled = bool(
            self.records and self.out_dir and self._selected_targets())

    def _refresh_list(self):
        panel = self.file_list_panel
        panel.Children.Clear()
        if not self.records:
            panel.Children.Add(self._muted_text(
                u"No files yet — use Add Files... or Add Folder..."))
            return

        targets = self._selected_targets()
        for record in self.records:
            block = Windows.Controls.TextBlock()
            block.Text = "{}   [{}]".format(record["name"],
                                            file_scan.describe(record))
            block.FontSize = 12
            block.Margin = Windows.Thickness(0, 2, 0, 2)
            block.TextTrimming = Windows.TextTrimming.CharacterEllipsis
            block.ToolTip = record["path"]
            block.Foreground = self._brush_for(record, targets)
            panel.Children.Add(block)

    def _brush_for(self, record, targets):
        """Red for files nothing can do, muted for ones every target skips."""
        if record.get("error") or record.get("workshared"):
            return self.FindResource("BrushDanger")
        if targets:
            usable = [t for t in targets
                      if upgrade_core.plan_for(record, t, "x")[0]
                      == upgrade_core.ACTION_UPGRADE]
            if not usable:
                return self.FindResource("BrushTextMuted")
        return self.FindResource("BrushTextPrimary")

    def _refresh_warning(self):
        # NOTE: these land on a WPF Text property, so they're unicode literals -
        # a utf-8 str would come through as mojibake under IronPython 2.
        notes = []
        workshared = len([r for r in self.records if r.get("workshared")])
        unreadable = len([r for r in self.records if r.get("error")])
        if workshared:
            notes.append(u"{} workshared file(s) will be skipped — a central "
                         u"model can't be batch-upgraded safely.".format(workshared))
        if unreadable:
            notes.append(u"{} file(s) couldn't be read and will be "
                         u"skipped.".format(unreadable))

        if notes:
            self.warn_lbl.Text = u"\n".join(notes)
            self.warn_lbl.Visibility = Windows.Visibility.Visible
        else:
            self.warn_lbl.Visibility = Windows.Visibility.Collapsed

    def _muted_text(self, text):
        block = Windows.Controls.TextBlock()
        block.Text = text
        block.FontSize = 12
        block.Opacity = 0.55
        block.Foreground = self.FindResource("BrushTextPrimary")
        block.TextWrapping = Windows.TextWrapping.Wrap
        return block

    # --- private helpers: templates ---
    def _setup_templates(self, select=None):
        """Fill the template dropdown from disk."""
        was_loading = self._loading
        self._loading = True
        try:
            names = templates.list_templates()
            self.template_cb.ItemsSource = names
            if select and select in names:
                self.template_cb.SelectedItem = select
            elif names and not self.template_cb.SelectedItem:
                self.template_cb.SelectedIndex = -1
        finally:
            self._loading = was_loading

    def _selected_template_name(self):
        return self.template_cb.SelectedItem

    def _load_template(self, name):
        data = templates.load_template(name)
        if not data:
            dlg.message(u'Could not load template "{}".'.format(name))
            return
        self.records = []
        self._add_paths([f["path"] for f in data.get("files") or []
                         if os.path.isfile(f.get("path", ""))])
        out_dir = data.get("out_dir") or ""
        if out_dir and os.path.isdir(out_dir):
            self.out_dir = out_dir
            self.out_tb.Text = out_dir
        wanted = set(data.get("targets") or [])
        for year, box in self._boxes.items():
            if box.IsEnabled:
                box.IsChecked = year in wanted
        self._refresh()

    # --- private helpers: schedule ---
    def _setup_schedule_time(self):
        """Default the date/time pickers to an hour from now."""
        default = datetime.datetime.now() + datetime.timedelta(hours=1)
        try:
            import System
            self.sched_date_dp.SelectedDate = System.DateTime(
                default.year, default.month, default.day)
        except Exception:
            pass
        self.sched_hour_cb.ItemsSource = ["{:02d}".format(h) for h in range(24)]
        self.sched_hour_cb.SelectedItem = "{:02d}".format(default.hour)
        self.sched_minute_cb.ItemsSource = ["00", "15", "30", "45"]
        self.sched_minute_cb.SelectedItem = "{:02d}".format(
            (default.minute // 15) * 15)

    def _picked_datetime(self):
        """The date/time currently set in the pickers, or None if incomplete
        or in the past - callers use None to mean "can't schedule right now"."""
        try:
            raw_date = self.sched_date_dp.SelectedDate
            hour = int(self.sched_hour_cb.SelectedItem)
            minute = int(self.sched_minute_cb.SelectedItem)
        except Exception:
            return None
        if raw_date is None:
            return None
        when = datetime.datetime(raw_date.Year, raw_date.Month, raw_date.Day,
                                 hour, minute)
        if when <= datetime.datetime.now():
            return None
        return when

    def _refresh_schedule_status(self):
        entry = schedule_io.armed_entry()
        if entry:
            when = entry.get("next_run", "").replace("T", " ")
            self.schedule_status_lbl.Text = (
                u'"{}" scheduled for {}'.format(entry.get("profile_name"), when))
            self.schedule_status_lbl.Visibility = Windows.Visibility.Visible
            self.schedule_cancel_btn.Visibility = Windows.Visibility.Visible
        else:
            self.schedule_status_lbl.Visibility = Windows.Visibility.Collapsed
            self.schedule_cancel_btn.Visibility = Windows.Visibility.Collapsed

    # --- private helpers: handlers ---
    def _on_add_files(self, sender, args):
        picked = forms.pick_file(files_filter=FILE_FILTER, multi_file=True)
        if not picked:
            return
        # pick_file hands back a bare string for a single selection in some
        # pyRevit versions; list() on that would split it into characters.
        if isinstance(picked, basestring):
            picked = [picked]
        self._add_paths(list(picked))
        self._refresh()

    def _on_add_folder(self, sender, args):
        folder = forms.pick_folder(title="Folder to scan (includes subfolders)")
        if folder:
            self._add_paths(file_scan.collect_folder(folder, recursive=True))
            self._refresh()

    def _on_clear(self, sender, args):
        self.records = []
        self._refresh()

    def _on_browse(self, sender, args):
        folder = forms.pick_folder(title="Where should the upgraded copies go?")
        if folder:
            self.out_dir = folder
            self.out_tb.Text = folder
            self._refresh()

    def _on_target_changed(self, sender, args):
        self._refresh()

    def _on_template_selected(self, sender, args):
        if self._loading:
            return
        name = self._selected_template_name()
        if name:
            self._load_template(name)

    def _on_template_save(self, sender, args):
        if not self.records:
            dlg.message(u"Add some files before saving a template.")
            return
        default = self._selected_template_name() or ""
        name = dlg.ask_string(u"Name for this template:",
                              title=u"Save Template", default=default)
        if not name:
            return
        templates.save_template(
            name, self.records, self.out_dir or "",
            self._selected_targets(), audit=False, compact=True)
        self._setup_templates(select=name)

    def _on_template_delete(self, sender, args):
        name = self._selected_template_name()
        if not name:
            return
        if not dlg.confirm(u'Delete template "{}"?'.format(name), yes=u"Delete"):
            return
        templates.delete_template(name)
        # A deleted template's schedule (if any) can't be re-armed later
        # by name, so it has to go too - a scheduled run with nothing to
        # load would otherwise fire and silently do nothing.
        entry = schedule_io.armed_entry()
        if entry and entry.get("profile_name") == name:
            schedule_io.disarm()
        self._setup_templates()
        self._refresh_schedule_status()

    def _on_schedule(self, sender, args):
        name = self._selected_template_name()
        if not name:
            dlg.message(u"Save this configuration as a template first, "
                        u"then schedule it.")
            return
        when = self._picked_datetime()
        if not when:
            dlg.message(u"Pick a date and time in the future to schedule.")
            return
        ok = schedule_io.arm(name, templates.template_path(name), when)
        if not ok:
            dlg.message(u"Could not save the schedule.")
            return
        self._refresh_schedule_status()
        dlg.message(u'"{}" will run at {}.'.format(
            name, when.strftime("%Y-%m-%d %H:%M")))

    def _on_schedule_cancel(self, sender, args):
        schedule_io.disarm()
        self._refresh_schedule_status()

    def _on_cancel(self, sender, args):
        self.confirmed = False
        self.Close()

    def _on_upgrade(self, sender, args):
        targets = self._selected_targets()
        if not (self.records and self.out_dir and targets):
            return
        self.targets = targets
        self.confirmed = True
        self.Close()


# ── UI / ENTRY POINT ───────────────────────────────────────────────────────

def main():
    host_year = running_year()
    window = BatchUpgradeWindow(XAML_FILE, host_year)
    if not window.show():
        return

    # Nothing below here runs while a modal dialog is up, which matters:
    # opening documents from inside a modal window is not a valid API
    # context and throws.
    by_year = batch_runner.run_batch(
        window.records, window.out_dir, window.targets,
        audit=False, compact=True,
        app=HOST_APP.app, uiapp=HOST_APP.uiapp, host_year=host_year)
    batch_runner.write_report(by_year, window.out_dir)


main()
