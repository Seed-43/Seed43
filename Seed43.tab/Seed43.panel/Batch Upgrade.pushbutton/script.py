# -*- coding: utf-8 -*-
# "Batch Upgrade"
# "Seed43"
# """
# Save upgraded copies of a pile of Revit files, in one or more target Revit
# versions.
#
# Revit can only ever save in its own version, so anything other than the
# running version is done by handing a job file to that Revit through the
# pyRevit CLI (see tools/job_io.py and worker.py). The running version is
# done here, in-session, because relaunching the Revit you are already in
# would be slower for no gain.
# """

# ── IMPORTS ────────────────────────────────────────────────────────────────

import io
import os
import sys
import time
import datetime
import subprocess

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from pyrevit import forms, script, HOST_APP
from pyrevit.framework import Windows

from Snippets._icons import set_header_icon
from Snippets.seed43_theme import apply_seed43_palette, apply_seed43_dimensions

from tools import file_scan
from tools import job_io
from tools import revit_versions
from tools import upgrade_core


# ── CONSTANTS ──────────────────────────────────────────────────────────────

XAML_FILE = os.path.join(SCRIPT_DIR, "BatchUpgrade.xaml")
WORKER_FILE = os.path.join(SCRIPT_DIR, "worker.py")

FILE_FILTER = ("Revit Files (*.rvt, *.rfa, *.rte, *.rft)|"
               "*.rvt;*.rfa;*.rte;*.rft")

# A worker Revit that hasn't finished in this long is assumed wedged - one
# stuck file must not hold the whole batch open forever.
WORKER_TIMEOUT_SECS = 3 * 60 * 60

CREATE_NO_WINDOW = 0x08000000


# ── HELPERS ────────────────────────────────────────────────────────────────

def running_year():
    """Return the Revit version number this script is running inside."""
    return int(str(HOST_APP.app.VersionNumber)[:4])


# ── CORE LOGIC ─────────────────────────────────────────────────────────────

# --- Running one target version ---

def _run_in_session(records, out_dir, year, audit, compact, progress):
    """Upgrade every file to the running version, in this Revit."""
    app = HOST_APP.app
    uiapp = HOST_APP.uiapp
    results = []

    with upgrade_core.DialogSuppressor(uiapp, app):
        for index, record in enumerate(records):
            if progress.cancelled:
                break
            progress.update_progress(index + 1, len(records))

            action, dst, reason = upgrade_core.plan_for(record, year, out_dir)
            if action == upgrade_core.ACTION_SKIP:
                results.append({"path": record["path"], "name": record["name"],
                                "ok": False, "skipped": True,
                                "dst": None, "message": reason})
                continue

            ok, message = upgrade_core.upgrade_one(
                app, record["path"], dst, audit=audit, compact=compact)
            results.append({"path": record["path"], "name": record["name"],
                            "ok": ok, "skipped": False,
                            "dst": dst if ok else None, "message": message})
    return results


def _run_via_cli(records, out_dir, year, audit, compact, cli_path, progress):
    """Launch Revit <year> through the pyRevit CLI and let worker.py save.

    Returns (results, error). error is set when the worker never produced a
    result file at all, which is different from it running and failing.
    """
    if not cli_path:
        return [], "pyRevit CLI not found - Revit {} can't be driven.".format(year)

    job_io.clear(year)
    job_file = job_io.write_job(year, out_dir, records, SCRIPT_DIR,
                                audit=audit, compact=compact)

    # Tell the worker exactly where its job is rather than making it re-derive
    # a temp path and hope the launched Revit inherited the same TEMP.
    env = os.environ.copy()
    env[job_io.JOB_ENV_VAR] = job_file

    command = [cli_path, "run", WORKER_FILE, "--revit={}".format(year), "--purge"]
    cli_log = job_io.cli_log_path(year)
    try:
        # Straight to a file, not subprocess.PIPE: the CLI reports real
        # failures on stdout and still exits 0, and an unread PIPE can fill
        # its buffer and deadlock the whole batch.
        cli_handle = open(cli_log, "w")
    except Exception:
        cli_handle = None
    try:
        proc = subprocess.Popen(command,
                                stdout=cli_handle or subprocess.PIPE,
                                stderr=subprocess.STDOUT,
                                env=env,
                                creationflags=CREATE_NO_WINDOW)
    except Exception as err:
        if cli_handle:
            cli_handle.close()
        return [], "Could not start the pyRevit CLI: {}".format(err)

    started = time.time()
    while proc.poll() is None:
        elapsed = int(time.time() - started)
        # update_progress is also what pumps the progress window, so it has
        # to be called on every pass or Cancel would never become clickable.
        # NOTE: ProgressBar.title only accepts str, silently ignoring unicode,
        # so every title here stays plain ASCII - no em-dashes.
        progress.title = ("Revit {} is running - {} file(s), "
                          "{}s elapsed".format(year, len(records), elapsed))
        progress.update_progress(1, 1)
        if progress.cancelled:
            try:
                proc.kill()
            except Exception:
                pass
            return [], "Cancelled while Revit {} was running.".format(year)
        if elapsed > WORKER_TIMEOUT_SECS:
            try:
                proc.kill()
            except Exception:
                pass
            return [], "Revit {} timed out after {} minutes.".format(
                year, WORKER_TIMEOUT_SECS // 60)
        time.sleep(1.0)

    if cli_handle:
        try:
            cli_handle.close()
        except Exception:
            pass

    payload = job_io.read_result(year)
    if payload is not None:
        _drop_job_file(year)
        return payload.get("results") or [], payload.get("error")

    return [], _explain_silent_exit(year, proc.returncode)


def _tail(path, limit=500):
    """Return the end of a log file, or '' if there isn't one."""
    if not os.path.isfile(path):
        return ""
    try:
        with open(path, "r") as handle:
            return handle.read().strip()[-limit:]
    except Exception:
        return ""


def _drop_job_file(year):
    """Remove a finished job so the worker's single-leftover fallback holds."""
    try:
        path = job_io.job_path(year)
        if os.path.isfile(path):
            os.remove(path)
    except Exception:
        pass


def _explain_silent_exit(year, returncode):
    """Turn 'Revit opened and closed and nothing happened' into a real reason.

    The three cases look identical from here but have completely different
    fixes, so they're told apart by which breadcrumbs the worker left behind.
    """
    cli_output = _tail(job_io.cli_log_path(year))
    worker_log = _tail(job_io.log_path(year))
    booted = os.path.isfile(job_io.boot_log_path(year))

    if worker_log:
        return ("Revit {} ran the worker but it failed:\n{}"
                .format(year, worker_log))

    if booted:
        return ("Revit {} started the worker but it stopped before writing "
                "results.\nBoot log: {}\nCLI output: {}"
                .format(year, _tail(job_io.boot_log_path(year)),
                        cli_output or "(none)"))

    if cli_output:
        return ("Revit {} never ran the worker. The pyRevit CLI said:\n{}"
                .format(year, cli_output))

    return ("Revit {} exited (code {}) without running the worker and without "
            "the CLI reporting anything. Check that pyRevit is attached to "
            "{} - run: pyrevit attached".format(year, returncode, year))


# --- Running the whole batch ---

def run_batch(records, out_dir, targets, audit, compact):
    """Run every selected target version. Returns {year: (results, error)}."""
    host_year = running_year()
    cli_path = revit_versions.find_pyrevit_cli()
    by_year = {}

    with forms.ProgressBar(title="Batch Upgrade", cancellable=True) as progress:
        for year in sorted(targets):
            progress.title = ("Upgrading to Revit {} - "
                              "{{value}} of {{max_value}}".format(year))
            if year == host_year:
                results = _run_in_session(records, out_dir, year,
                                          audit, compact, progress)
                by_year[year] = (results, None)
            else:
                results, error = _run_via_cli(records, out_dir, year, audit,
                                              compact, cli_path, progress)
                by_year[year] = (results, error)
            if progress.cancelled:
                break
    return by_year


# --- Reporting ---

def write_report(by_year, out_dir):
    """Print a summary to the output window and drop a log beside the copies."""
    output = script.get_output()
    output.set_title(u"Seed43 — Batch Upgrade")
    output.print_md(u"# Batch Upgrade")

    # Log lines are kept unicode throughout and written through io.open with
    # an explicit encoding: file names coming back from a worker's JSON are
    # unicode, and mixing them into byte strings blows up on the first
    # non-ASCII model name.
    lines = [u"Seed43 Batch Upgrade - {}".format(
        datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))]

    for year in sorted(by_year):
        results, error = by_year[year]
        done = [r for r in results if r.get("ok")]
        skipped = [r for r in results if r.get("skipped")]
        failed = [r for r in results if not r.get("ok") and not r.get("skipped")]

        output.print_md(u"## Revit {} — {} upgraded, {} skipped, {} failed"
                        .format(year, len(done), len(skipped), len(failed)))
        lines.append(u"")
        lines.append(u"Revit {}: {} upgraded, {} skipped, {} failed".format(
            year, len(done), len(skipped), len(failed)))

        if error:
            output.print_md(u"**{}**".format(error))
            lines.append(u"  ERROR: {}".format(error))

        rows = []
        for record in results:
            if record.get("ok"):
                state = u"upgraded"
            elif record.get("skipped"):
                state = u"skipped"
            else:
                state = u"FAILED"
            rows.append([record.get("name") or u"", state,
                         record.get("message") or u""])
            lines.append(u"  [{}] {} {}".format(
                state, record.get("name"), record.get("message") or u""))
        if rows:
            output.print_table(rows, columns=["File", "Result", "Detail"])

    if out_dir and os.path.isdir(out_dir):
        log_file = os.path.join(out_dir, "Seed43_BatchUpgrade_{}.log".format(
            datetime.datetime.now().strftime("%Y%m%d_%H%M%S")))
        try:
            with io.open(log_file, "w", encoding="utf-8") as handle:
                handle.write(u"\n".join(lines))
            output.print_md(u"Log written to `{}`".format(log_file))
        except Exception:
            pass


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

        self._bind()
        apply_seed43_palette(self, SCRIPT_DIR)
        apply_seed43_dimensions(self, SCRIPT_DIR)
        set_header_icon(self, SCRIPT_DIR)
        self._build_version_grid()
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

    def _build_version_grid(self):
        """One checkbox per Revit year, greyed out when it can't be used."""
        self.running_lbl.Text = "Running: {}".format(self.host_year)
        self.cli_path = revit_versions.find_pyrevit_cli()
        panel = self.version_panel
        panel.Children.Clear()
        self._boxes = {}

        for year, installed, is_running in revit_versions.version_grid(self.host_year):
            box = Windows.Controls.CheckBox()
            box.Style = self.FindResource("VersionCheck")
            # Anything other than the running Revit has to be driven by the
            # CLI, so without it only the running version is selectable.
            usable = installed and (is_running or self.cli_path is not None)
            if not installed:
                label = "Revit {} (not installed)".format(year)
            elif not usable:
                label = "Revit {} (needs pyRevit CLI)".format(year)
            else:
                label = "Revit {}".format(year)
            box.Content = label
            box.IsEnabled = usable
            box.IsChecked = bool(is_running and usable)
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
        if not self.cli_path:
            notes.append(u"pyRevit CLI not found — only Revit {} "
                         u"can be targeted.".format(self.host_year))

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
    by_year = run_batch(window.records, window.out_dir, window.targets,
                        audit=False, compact=True)
    write_report(by_year, window.out_dir)


main()
