# -*- coding: utf-8 -*-
# "Batch Upgrade / job_io"
# "Seed43"
# """
# The hand-off between the tool and a Revit launched by the pyRevit CLI.
# `pyrevit run` takes a script path and nothing else, so the file list and
# options travel through a JSON file at a fixed temp path instead of argv.
# """

# ── IMPORTS ────────────────────────────────────────────────────────────────

import os
import json

__all__ = ["job_dir", "job_path", "result_path", "log_path",
           "write_job", "read_job", "write_result", "read_result", "clear"]


# ── CONSTANTS ──────────────────────────────────────────────────────────────

# Deliberately a fixed, predictable location rather than a random temp dir:
# the worker has no way to be told where to look, so both sides have to
# derive the same path independently.
_ROOT = os.path.join(
    os.environ.get("TEMP") or os.environ.get("TMP") or ".",
    "Seed43", "batch_upgrade")


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


def boot_log_path(year):
    """Proof-of-life the worker writes before it touches anything else.

    If this file is missing after a run, the worker never executed at all -
    which is a completely different problem from it executing and failing,
    and the two are otherwise indistinguishable from the launching side.
    """
    return os.path.join(job_dir(), "boot_{}.log".format(year))


def cli_log_path(year):
    """Everything the pyRevit CLI printed while driving that Revit.

    The CLI reports real errors here and still exits 0 - e.g. "pyRevit is not
    attached to Revit 2024" - so this output is the only place some failures
    are ever explained.
    """
    return os.path.join(job_dir(), "cli_{}.log".format(year))


# Passed to the worker so it never has to guess where its job file is.
JOB_ENV_VAR = "SEED43_BATCH_UPGRADE_JOB"


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


def clear(year):
    """Remove any stale job/result/log left by a previous run of this year.

    Without this a crashed worker's old result file would be read back as if
    it belonged to the current run.
    """
    for path in (job_path(year), result_path(year), log_path(year),
                 boot_log_path(year), cli_log_path(year)):
        try:
            if os.path.isfile(path):
                os.remove(path)
        except OSError:
            pass
