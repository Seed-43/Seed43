# -*- coding: utf-8 -*-
# "Batch Upgrade / job_io"
# "Seed43"
# """
# The hand-off between the tool and a Revit launched for a specific year.
# Revit.exe takes no argv of ours to speak of, so the file list and
# options travel through a JSON file at a fixed path instead - the job
# file's mere existence for a given year is also the signal that
# startup.py's Idling hook (see tools/headless_batch.py) has work to do.
#
# Lives under .user, same convention pySheets already uses for its own
# settings (see Seed43.extension/startup.py's _SCHEDULE_FILE_NEW) - the
# extension's own updater explicitly preserves .json files when syncing
# Seed43.tab, so anything here survives a Seed43 update automatically.
# """

# ── IMPORTS ────────────────────────────────────────────────────────────────

import os
import json

__all__ = ["job_dir", "job_path", "result_path", "log_path",
           "write_job", "read_job", "delete_job",
           "write_result", "read_result", "delete_result", "clear"]


# ── CONSTANTS ──────────────────────────────────────────────────────────────

# tools/ -> Batch Upgrade.pushbutton -> Seed43.panel -> Seed43.tab ->
# Seed43.extension. Walked explicitly rather than hardcoded as a string so
# this keeps working if the extension folder itself ever gets renamed.
_TOOLS_DIR = os.path.dirname(os.path.abspath(__file__))
_PUSHBUTTON_DIR = os.path.dirname(_TOOLS_DIR)
_PANEL_DIR = os.path.dirname(_PUSHBUTTON_DIR)
_TAB_DIR = os.path.dirname(_PANEL_DIR)
_EXTENSION_DIR = os.path.dirname(_TAB_DIR)

_ROOT = os.path.join(_EXTENSION_DIR, ".user", "BatchUpgrade", "settings")


# ── HELPERS ────────────────────────────────────────────────────────────────

def job_dir():
    """Return the exchange folder, creating it if needed."""
    if not os.path.isdir(_ROOT):
        try:
            os.makedirs(_ROOT)
        except OSError:
            pass
    return _ROOT


def job_path(year):
    return os.path.join(job_dir(), "job_{}.json".format(year))


def result_path(year):
    return os.path.join(job_dir(), "result_{}.json".format(year))


def log_path(year):
    """Where the worker dumps a traceback if it dies before writing results."""
    return os.path.join(job_dir(), "worker_{}.log".format(year))


# ── CORE LOGIC ─────────────────────────────────────────────────────────────

def write_job(year, out_dir, records, tool_dir, audit=False, compact=True):
    """Write the job file a worker Revit will pick up. Returns its path.

    tool_dir is the pushbutton folder, carried in the payload because the
    worker can't rely on __file__ to find its way back to tools/.
    """
    payload = {
        "target": int(year),
        "out_dir": out_dir,
        "audit": bool(audit),
        "compact": bool(compact),
        "tool_dir": tool_dir,
        "files": records,
    }
    path = job_path(year)
    with open(path, "w") as handle:
        json.dump(payload, handle, indent=2)
    return path


def read_job(year):
    """Read the job file for a year, or None if it isn't there."""
    path = job_path(year)
    if not os.path.isfile(path):
        return None
    try:
        with open(path, "r") as handle:
            return json.load(handle)
    except Exception:
        return None


def delete_job(year):
    """Remove just the job file, leaving any result/log alone.

    Used by headless_batch to consume a job the instant it's picked up,
    before doing any actual work - separate from clear() (which wipes
    job/result/log together to reset before a new run) because this needs
    to be surgical: only the job file, only right now, so a stray second
    Idling tick (or Revit being opened normally later) can never see it.
    """
    path = job_path(year)
    try:
        if os.path.isfile(path):
            os.remove(path)
    except OSError:
        pass


def write_result(year, results, error=None):
    """Write what the worker managed to do, for the launching tool to read."""
    payload = {"target": int(year), "results": results, "error": error}
    path = result_path(year)
    with open(path, "w") as handle:
        json.dump(payload, handle, indent=2)
    return path


def read_result(year):
    """Read a worker's result file, or None if it never wrote one."""
    path = result_path(year)
    if not os.path.isfile(path):
        return None
    try:
        with open(path, "r") as handle:
            return json.load(handle)
    except Exception:
        return None


def delete_result(year):
    """Remove just the result file, once script.py has finished reading it.

    Mirrors delete_job's "consume the instant you're done with it"
    discipline - without this, a successful run's result file just sits
    in .user/BatchUpgrade/settings until the *next* run for that year
    (when clear() would eventually wipe it), rather than being cleaned up
    right away.
    """
    path = result_path(year)
    try:
        if os.path.isfile(path):
            os.remove(path)
    except OSError:
        pass


def clear(year):
    """Remove any stale job/result/log left by a previous run of this year.

    Without this a crashed worker's old result file would be read back as if
    it belonged to the current run.
    """
    for path in (job_path(year), result_path(year), log_path(year)):
        try:
            if os.path.isfile(path):
                os.remove(path)
        except OSError:
            pass
