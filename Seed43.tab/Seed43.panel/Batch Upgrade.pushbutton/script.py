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
from Snippets import _timepicker
from Snippets._icons import make_icon, set_header_icon
from Snippets._support import github_issue_url, open_url, support_mailto
from Snippets.seed43_theme import (apply_seed43_palette, apply_seed43_dimensions,
                                   get_color)

from tools import batch_runner
from tools import file_scan
from tools import revit_versions
from tools import schedule_io
from tools import upgrade_core


# ── CONSTANTS ──────────────────────────────────────────────────────────────

TOOL_NAME = "Batch Upgrade"
ABOUT_URL = "https://seed43.org/"
SUPPORT_URL = "https://buymeacoffee.com/seed43"

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
        # make_icon bakes its colour in, so this is built after the palette
        # has been applied rather than declared in XAML.
        self.settings_toggle_btn.Content = make_icon(
            "menu", size=18,
            color=get_color(SCRIPT_DIR, "text_primary", fallback="#F4FAFF"))
        self._build_version_grid()
        self._setup_schedule()
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
        self.FindName("sched_enable_cb").Checked += self._on_sched_enable
        self.FindName("sched_enable_cb").Unchecked += self._on_sched_enable
        self.FindName("sched_date_dp").SelectedDateChanged += self._on_sched_field
        self.FindName("sched_remove_btn").Click += self._on_sched_remove
        self.FindName("sched_time_popup").Opened += self._on_tp_opened
        self.FindName("tp_ok_btn").Click += self._on_tp_ok
        self.FindName("tp_cancel_btn").Click += self._on_tp_cancel
        self.FindName("settings_toggle_btn").PreviewMouseLeftButtonDown += \
            self._on_menu_preview_down
        self.FindName("help_btn").Click += self._on_menu_help
        self.FindName("issue_btn").Click += self._on_menu_issue
        self.FindName("about_btn").Click += self._on_menu_about
        self.FindName("support_btn").Click += self._on_menu_support

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
        self._update_sched_gates()
        self.count_lbl.Text = "{} file(s)".format(len(self.records))
        ready = bool(self.records and self.out_dir and self._selected_targets())
        # Scheduling additionally needs a time that hasn't already passed,
        # otherwise the button would offer to arm something that can never fire.
        if ready and self._scheduling():
            ready = self._picked_datetime() is not None
        self.upgrade_btn.IsEnabled = ready

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

    # --- private helpers: schedule ---
    def _setup_schedule(self):
        """Build the wheel time picker and default the date to an hour out."""
        import System

        default = datetime.datetime.now() + datetime.timedelta(hours=1)
        try:
            self.sched_date_dp.SelectedDate = System.DateTime(
                default.year, default.month, default.day)
        except Exception:
            pass

        # Wheel text colour is baked in when the wheels are built, so it's
        # read from the palette rather than left at the module default.
        self.picker = _timepicker.WheelTimePicker(
            self.tp_hours, self.tp_minutes, self.tp_ampm,
            button=self.sched_time_btn,
            foreground=get_color(SCRIPT_DIR, "text_primary", fallback="#F4FAFF"))
        self.picker.set_time(default.hour, (default.minute // 5) * 5)

        # A schedule armed in an earlier session is still pending - reopen it
        # in full so it can be edited rather than rebuilt from scratch.
        entry = schedule_io.armed_entry()
        if entry:
            self._load_armed_entry(entry)

    def _load_armed_entry(self, entry):
        """Reopen a pending schedule: its frozen job, then its timing."""
        was_loading = self._loading
        self._loading = True
        try:
            snapshot = schedule_io.read_snapshot(entry.get("profile_path"))
            if snapshot:
                self._apply_job(snapshot)
            parsed = self._parse_next_run(entry.get("next_run"))
            if parsed:
                import System
                self.sched_date_dp.SelectedDate = System.DateTime(
                    parsed.year, parsed.month, parsed.day)
                self.picker.set_time(parsed.hour, parsed.minute)
            self.sched_enable_cb.IsChecked = True
        finally:
            self._loading = was_loading

    def _apply_job(self, data):
        """Load a frozen job's files, output folder and targets into the UI.

        Files that have been moved or deleted since it was armed are dropped
        rather than carried as broken rows - the schedule is being reopened
        to be edited, so it should show what actually still exists.
        """
        self.records = []
        self._add_paths([f["path"] for f in data.get("files") or []
                         if f.get("path") and os.path.isfile(f["path"])])
        out_dir = data.get("out_dir") or ""
        if out_dir and os.path.isdir(out_dir):
            self.out_dir = out_dir
            self.out_tb.Text = out_dir
        wanted = set(data.get("targets") or [])
        for year, box in self._boxes.items():
            if box.IsEnabled:
                box.IsChecked = year in wanted

    @staticmethod
    def _parse_next_run(raw):
        """Parse an armed entry's next_run stamp, or None if unusable."""
        if not raw:
            return None
        sched = schedule_io.schedule_mod()
        if sched:
            try:
                return sched.parse_ts(raw)
            except Exception:
                pass
        return None

    def _picked_datetime(self):
        """The date/time currently set in the card, or None if incomplete or
        in the past - callers use None to mean "can't arm this"."""
        try:
            raw_date = self.sched_date_dp.SelectedDate
        except Exception:
            return None
        if raw_date is None:
            return None
        chosen = self.picker.get_time()
        if not chosen:
            return None
        hour, minute = chosen
        when = datetime.datetime(raw_date.Year, raw_date.Month, raw_date.Day,
                                 hour, minute)
        if when <= datetime.datetime.now():
            return None
        return when

    def _scheduling(self):
        """True while the window is in 'arm it for later' mode."""
        return bool(self.sched_enable_cb.IsChecked)

    def _update_sched_gates(self):
        """Enable unlocks the timing controls and retargets the main button.

        There's one action button for both jobs: it reads Upgrade normally,
        and Schedule once Enable is ticked, so it's always obvious which of
        the two pressing it will do.
        """
        on = self._scheduling()
        self.sched_time_btn.IsEnabled = on
        self.sched_date_dp.IsEnabled = on
        self.upgrade_btn.Content = u"Schedule" if on else u"Upgrade"

    def _refresh_schedule_status(self):
        self._update_sched_gates()
        entry = schedule_io.armed_entry()
        self.sched_remove_btn.Visibility = (
            Windows.Visibility.Visible if entry else Windows.Visibility.Collapsed)

        if entry:
            when = self._parse_next_run(entry.get("next_run"))
            shown = when.strftime("%d %b %Y, %I:%M %p") if when else \
                (entry.get("next_run") or "").replace("T", " ")
            text = u"Scheduled for {}".format(shown)
            # With Enable off, what's on screen is no longer what's booked in -
            # say so, or the green line reads as though it tracks the window.
            if not self._scheduling():
                text += u" (saved separately from what's shown)"
            self.sched_status_tb.Text = text
            self.sched_status_tb.Foreground = self.FindResource("BrushPrimaryGreen")
        elif self._scheduling():
            # Ticked but not armed yet - say what the button is waiting for,
            # so an un-pressed Schedule doesn't read as a saved one.
            self.sched_status_tb.Text = u"Not armed yet — press Schedule"
            self.sched_status_tb.Foreground = self.FindResource("BrushTextMuted")
        else:
            self.sched_status_tb.Text = u"Schedule off"
            self.sched_status_tb.Foreground = self.FindResource("BrushTextMuted")

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

    def _on_tp_opened(self, sender, args):
        """Re-centre the wheels on the current time each time the popup opens."""
        self.picker.refresh()

    def _on_tp_ok(self, sender, args):
        self.picker.commit()
        self.sched_time_btn.IsChecked = False    # closes the popup
        self._on_sched_field(sender, args)

    def _on_tp_cancel(self, sender, args):
        self.sched_time_btn.IsChecked = False

    # --- private helpers: hamburger menu ---
    def _on_menu_preview_down(self, sender, args):
        """Explicit close-on-reclick.

        StaysOpen=False already closes the popup on any outside click,
        including one on this toggle - and that same click's Click event
        would flip IsChecked straight back to True, reopening it. Intercept
        first so a re-click only ever closes.
        """
        if self.settings_popup.IsOpen:
            self.settings_popup.IsOpen = False
            self.settings_toggle_btn.IsChecked = False
            args.Handled = True

    def _open_url(self, url, title=""):
        open_url(url, window=self,
                 on_error=lambda msg: dlg.message(msg, title=title))

    def _on_menu_help(self, sender, args):
        self.settings_toggle_btn.IsChecked = False
        self._open_url(support_mailto(TOOL_NAME, SCRIPT_DIR), title=u"Support")

    def _on_menu_issue(self, sender, args):
        self.settings_toggle_btn.IsChecked = False
        self._open_url(github_issue_url(TOOL_NAME, SCRIPT_DIR),
                       title=u"Report an issue")

    def _on_menu_about(self, sender, args):
        self.settings_toggle_btn.IsChecked = False
        self._open_url(ABOUT_URL, title=u"About")

    def _on_menu_support(self, sender, args):
        self.settings_toggle_btn.IsChecked = False
        self._open_url(SUPPORT_URL, title=u"Support")

    def _on_sched_enable(self, sender, args):
        """Ticking only switches the button to Schedule - it doesn't arm.

        Unticking deliberately leaves a pending schedule alone, so you can
        untick, clear the list and run something else now without losing
        what's already booked in. Remove is what deletes it.
        """
        if self._loading:
            return
        self._refresh_schedule_status()
        self._refresh()

    def _on_sched_remove(self, sender, args):
        """Delete the pending schedule and its frozen job."""
        if not schedule_io.armed_entry():
            return
        if not dlg.confirm(u"Delete the saved schedule?", yes=u"Delete"):
            return
        schedule_io.disarm()
        self.sched_enable_cb.IsChecked = False
        self._refresh_schedule_status()
        self._refresh()

    def _on_sched_field(self, sender, args):
        """Date or time changed - only the button's enabled state moves."""
        if self._loading:
            return
        self._refresh()

    def _arm(self):
        """Freeze what's set up in this window and arm it for later.

        The whole configuration is snapshotted now, so editing the file list
        afterwards can't quietly change what the scheduled run does.
        Returns True once armed.
        """
        targets = self._selected_targets()
        if not (self.records and self.out_dir and targets):
            dlg.message(u"Add files, pick an output folder and tick at least "
                        u"one Revit version before scheduling.")
            return False
        when = self._picked_datetime()
        if not when:
            dlg.message(u"Pick a date and time in the future to schedule.")
            return False
        if not schedule_io.arm(when, self.records, self.out_dir, targets,
                               audit=False, compact=True):
            dlg.message(u"Could not save the schedule.")
            return False
        return True

    def _on_cancel(self, sender, args):
        self.confirmed = False
        self.Close()

    def _on_upgrade(self, sender, args):
        """The single action button - upgrades now, or arms it for later."""
        targets = self._selected_targets()
        if not (self.records and self.out_dir and targets):
            return

        if self._scheduling():
            if not self._arm():
                return
            # confirmed stays False: the run happens later, from startup.py's
            # timer, so main() must not also kick it off right now.
            self.Close()
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
