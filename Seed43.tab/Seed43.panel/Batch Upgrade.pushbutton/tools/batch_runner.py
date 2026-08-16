# -*- coding: utf-8 -*-
# "Batch Upgrade / batch_runner"
# "Seed43"
# """
# Runs a batch upgrade across one or more target Revit versions. Used by
# script.py's interactive window (a real progress bar, a real output
# window) and by Seed43.extension/startup.py's scheduled-run handler (no
# window at all - nobody's watching a run that fires while you're doing
# something else in Revit). Pulled out into its own module specifically
# so those two callers can't quietly drift apart.
# """

# ── IMPORTS ────────────────────────────────────────────────────────────────

import io
import os
import time
import datetime
import subprocess

from tools import job_io
from tools import revit_versions
from tools import upgrade_core

__all__ = ["WORKER_TIMEOUT_SECS", "NullProgress", "run_batch", "write_report"]


# ── CONSTANTS ──────────────────────────────────────────────────────────────

# A launched Revit that hasn't finished in this long is assumed wedged -
# one stuck file must not hold the whole batch open forever.
WORKER_TIMEOUT_SECS = 3 * 60 * 60

# tools/ -> Batch Upgrade.pushbutton - this tool's own folder, handed to
# job_io.write_job as tool_dir so headless_batch.py can find tools/
# again once a launched Revit picks the job up.
_PUSHBUTTON_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


# ── HELPERS (private to this file) ────────────────────────────────────────

class _NullProgress(object):
    """Stand-in for pyrevit.forms.ProgressBar when nothing is watching.

    The scheduled/headless path has no window to update and nobody to
    offer Cancel to - this just quietly no-ops everything run_batch
    would otherwise call on a real progress bar. Exposed as NullProgress
    (see __all__) since callers that want it explicitly - rather than
    relying on run_batch's own default, which creates a real, visible
    ProgressBar - need to pass an instance in.
    """
    cancelled = False

    def __init__(self):
        self.title = ""

    def update_progress(self, value, max_value):
        pass


NullProgress = _NullProgress


# ── CORE LOGIC ─────────────────────────────────────────────────────────────

def _run_in_session(records, out_dir, year, audit, compact, app, uiapp, progress):
    """Upgrade every file to the running version, in this Revit."""
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


def _run_via_direct_launch(records, out_dir, year, audit, compact, progress):
    """Launch Revit <year> directly and let startup.py's Idling hook save.

    No pyRevit CLI involved at all (see tools/headless_batch.py's module
    docstring for why). Revit.exe opens normally, pyRevit attaches the
    same way it does for a human opening it by hand, startup.py notices
    the waiting job and does the work, then closes Revit itself.

    Returns (results, error). error is set when Revit closed without ever
    writing a result file - could mean it crashed, someone closed it by
    hand before it finished, or the pickup in startup.py didn't fire.
    """
    exe = revit_versions.revit_exe(year)
    if not exe:
        return [], "Revit {} is not installed.".format(year)

    job_io.clear(year)
    job_io.write_job(year, out_dir, records, _PUSHBUTTON_DIR,
                     audit=audit, compact=compact)

    try:
        proc = subprocess.Popen([exe])
    except Exception as err:
        return [], "Could not start Revit {}: {}".format(year, err)

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

    payload = job_io.read_result(year)
    if payload is None:
        return [], ("Revit {} closed without writing results (code {}). "
                    "startup.py may not have picked up the job - see "
                    "tools/headless_batch.py.".format(year, proc.returncode))

    job_io.delete_result(year)
    return payload.get("results") or [], payload.get("error")


def run_batch(records, out_dir, targets, audit, compact,
              app, uiapp, host_year, progress=None):
    """Run every selected target version. Returns {year: (results, error)}.

    app/uiapp/host_year are passed in explicitly rather than pulled from
    pyrevit.HOST_APP, so this works the same whether it's called from an
    interactive script.py session or from startup.py's headless schedule
    handler (which has its own app/uiapp from the ExternalEvent it fired
    on). progress defaults to a real pyrevit.forms.ProgressBar when None
    and a UI is expected - the headless caller passes a _NullProgress
    instead.
    """
    by_year = {}
    owns_progress = progress is None
    if owns_progress:
        from pyrevit import forms
        progress_cm = forms.ProgressBar(title="Batch Upgrade", cancellable=True)
    else:
        progress_cm = None

    def _run(progress):
        for year in sorted(targets):
            progress.title = ("Upgrading to Revit {} - "
                              "{{value}} of {{max_value}}".format(year))
            if year == host_year:
                results = _run_in_session(records, out_dir, year, audit,
                                          compact, app, uiapp, progress)
                by_year[year] = (results, None)
            else:
                results, error = _run_via_direct_launch(
                    records, out_dir, year, audit, compact, progress)
                by_year[year] = (results, error)
            if progress.cancelled:
                break

    if owns_progress:
        with progress_cm as progress:
            _run(progress)
    else:
        _run(progress)

    return by_year


def write_report(by_year, out_dir, show_output=True):
    """Print a summary to the output window and drop a log beside the copies.

    show_output=False skips pyrevit.script.get_output() entirely - used
    by the headless scheduled path, where popping an output window while
    nobody's watching (possibly mid-way through something else in Revit)
    would be more surprising than helpful. The log file is still written
    either way.
    """
    if show_output:
        from pyrevit import script
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

        if show_output:
            output.print_md(u"## Revit {} — {} upgraded, {} skipped, {} failed"
                            .format(year, len(done), len(skipped), len(failed)))
        lines.append(u"")
        lines.append(u"Revit {}: {} upgraded, {} skipped, {} failed".format(
            year, len(done), len(skipped), len(failed)))

        if error:
            if show_output:
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
        if rows and show_output:
            output.print_table(rows, columns=["File", "Result", "Detail"])

    if out_dir and os.path.isdir(out_dir):
        log_file = os.path.join(out_dir, "Seed43_BatchUpgrade_{}.log".format(
            datetime.datetime.now().strftime("%Y%m%d_%H%M%S")))
        try:
            with io.open(log_file, "w", encoding="utf-8") as handle:
                handle.write(u"\n".join(lines))
            if show_output:
                output.print_md(u"Log written to `{}`".format(log_file))
        except Exception:
            pass
