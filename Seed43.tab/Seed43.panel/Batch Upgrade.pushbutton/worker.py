# -*- coding: utf-8 -*-
# "Batch Upgrade / worker"
# "Seed43"
# """
# Runs inside a Revit that the pyRevit CLI launched for us:
#
#     pyrevit run worker.py --revit=2024
#
# Picks up the job file written for its own version, saves an upgraded copy
# of every file in it, and writes the results back for the launching tool to
# read. Not a pushbutton - it has no UI and is never shown in the ribbon.
#
# Everything here is written defensively: this code runs in a Revit with no
# visible UI and no output window, so an unhandled exception would otherwise
# vanish completely and look identical to "Revit opened and closed again".
# """

# ── IMPORTS ────────────────────────────────────────────────────────────────

import os
import sys
import json
import traceback


# ── BOOTSTRAP ──────────────────────────────────────────────────────────────

# The exchange folder has to be re-derived here with the standard library
# alone: tools/ can't be imported until the job tells us where it lives, and
# the job is what we're trying to find. Keep in step with tools/job_io.py.
_ROOT = os.path.join(
    os.environ.get("TEMP") or os.environ.get("TMP") or ".",
    "Seed43", "batch_upgrade")

_JOB_ENV_VAR = "SEED43_BATCH_UPGRADE_JOB"

# Filled in as soon as the job is located, so failure logs can be named after
# the right version even when the Revit version lookup is what blew up.
_year = "unknown"


def _write(path, text):
    """Best-effort write - never raises, because callers are error paths."""
    try:
        folder = os.path.dirname(path)
        if folder and not os.path.isdir(folder):
            os.makedirs(folder)
        with open(path, "w") as handle:
            handle.write(text)
        return True
    except Exception:
        return False


def _boot_log(note):
    """Prove the worker actually started, before anything can go wrong.

    Without this, "the script never ran" and "the script ran and died on its
    first import" produce byte-identical evidence: an empty folder.
    """
    _write(os.path.join(_ROOT, "boot_{}.log".format(_year)), note)


def _fail(message):
    """Record a run that died before it could process anything."""
    _write(os.path.join(_ROOT, "worker_{}.log".format(_year)), message)
    _write(os.path.join(_ROOT, "result_{}.json".format(_year)),
           json.dumps({"target": _year, "results": [], "error": message},
                      indent=2))


def _find_job():
    """Return (job_dict, path). The env var wins; the derived path is backup.

    Relying on the derived path alone assumes the launched Revit inherits the
    same TEMP, which is exactly the sort of thing that silently isn't true.
    """
    candidates = []
    from_env = os.environ.get(_JOB_ENV_VAR)
    if from_env:
        candidates.append(from_env)

    # Fallback for when the env var doesn't survive the CLI -> Revit hop. The
    # launcher deletes each job file once that version is done, so a single
    # leftover job_*.json is unambiguously the one meant for this run.
    try:
        pending = [os.path.join(_ROOT, name) for name in sorted(os.listdir(_ROOT))
                   if name.startswith("job_") and name.endswith(".json")]
        if len(pending) == 1:
            candidates.append(pending[0])
    except Exception:
        pass

    for path in candidates:
        if path and os.path.isfile(path):
            try:
                with open(path, "r") as handle:
                    return json.load(handle), path
            except Exception as err:
                _fail("Job file {} unreadable: {}".format(path, err))
                return None, path

    # Nothing found - say precisely what was looked for and what is actually
    # sitting in the exchange folder, so the report can explain itself.
    listing = []
    try:
        listing = sorted(os.listdir(_ROOT))
    except Exception:
        listing = ["<could not list {}>".format(_ROOT)]
    _fail("No job file found.\n  {}={}\n  exchange folder: {}\n  contains: {}"
          .format(_JOB_ENV_VAR, from_env, _ROOT, ", ".join(listing) or "(empty)"))
    return None, None


# ── CORE LOGIC ─────────────────────────────────────────────────────────────

def main():
    global _year

    _boot_log("worker started\n  TEMP={}\n  {}={}\n".format(
        os.environ.get("TEMP"), _JOB_ENV_VAR, os.environ.get(_JOB_ENV_VAR)))

    job, job_file = _find_job()
    if job is None:
        return

    _year = job.get("target", "unknown")
    _boot_log("worker started, job loaded from {}\n  target={}\n  files={}\n"
              .format(job_file, _year, len(job.get("files") or [])))

    tool_dir = job.get("tool_dir")
    if tool_dir and tool_dir not in sys.path:
        sys.path.insert(0, tool_dir)

    try:
        from pyrevit import HOST_APP
        from tools import job_io
        from tools import upgrade_core
    except Exception:
        _fail("Import failed inside Revit (tool_dir={}):\n{}".format(
            tool_dir, traceback.format_exc()))
        return

    app = HOST_APP.app
    uiapp = HOST_APP.uiapp

    # The job names the version it was written for; trust that over anything
    # derived here, so results always land where the launcher is watching.
    host_year = None
    try:
        host_year = int(str(app.VersionNumber)[:4])
    except Exception:
        pass
    target = int(_year) if str(_year).isdigit() else host_year
    if target is None:
        _fail("Could not determine the target Revit version.")
        return

    out_dir = job.get("out_dir") or ""
    audit = bool(job.get("audit"))
    compact = bool(job.get("compact", True))

    results = []
    try:
        with upgrade_core.DialogSuppressor(uiapp, app):
            for record in job.get("files") or []:
                action, dst, reason = upgrade_core.plan_for(
                    record, target, out_dir)
                if action == upgrade_core.ACTION_SKIP:
                    results.append({"path": record.get("path"),
                                    "name": record.get("name"),
                                    "ok": False, "skipped": True,
                                    "dst": None, "message": reason})
                    continue
                ok, message = upgrade_core.upgrade_one(
                    app, record["path"], dst, audit=audit, compact=compact)
                results.append({"path": record.get("path"),
                                "name": record.get("name"),
                                "ok": ok, "skipped": False,
                                "dst": dst if ok else None,
                                "message": message})
    except Exception:
        # Still write out whatever completed before the blow-up.
        _write(os.path.join(_ROOT, "worker_{}.log".format(_year)),
               traceback.format_exc())

    try:
        job_io.write_result(target, results)
    except Exception:
        _write(os.path.join(_ROOT, "result_{}.json".format(target)),
               json.dumps({"target": target, "results": results,
                           "error": None}, indent=2))


try:
    main()
except Exception:
    _fail(traceback.format_exc())
