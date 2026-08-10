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
# """

# ── IMPORTS ────────────────────────────────────────────────────────────────

import os
import sys
import json
import traceback


# ── BOOTSTRAP ──────────────────────────────────────────────────────────────

# The job path has to be re-derived here with the standard library alone:
# tools/ can't be imported until the job tells us where it lives, and the
# job is what we're trying to find. Keep this in step with tools/job_io.py.
_ROOT = os.path.join(
    os.environ.get("TEMP") or os.environ.get("TMP") or ".",
    "Seed43", "batch_upgrade")


def _host_year():
    """Return the Revit version number this worker is running inside."""
    from pyrevit import HOST_APP
    return int(str(HOST_APP.app.VersionNumber)[:4])


def _fail(year, message):
    """Record a run that died before it could process anything."""
    try:
        with open(os.path.join(_ROOT, "worker_{}.log".format(year)), "w") as handle:
            handle.write(message)
    except Exception:
        pass
    try:
        with open(os.path.join(_ROOT, "result_{}.json".format(year)), "w") as handle:
            json.dump({"target": year, "results": [], "error": message}, handle, indent=2)
    except Exception:
        pass


# ── CORE LOGIC ─────────────────────────────────────────────────────────────

def main():
    year = _host_year()

    job_file = os.path.join(_ROOT, "job_{}.json".format(year))
    if not os.path.isfile(job_file):
        _fail(year, "No job file at {} - nothing to do.".format(job_file))
        return
    with open(job_file, "r") as handle:
        job = json.load(handle)

    tool_dir = job.get("tool_dir")
    if tool_dir and tool_dir not in sys.path:
        sys.path.insert(0, tool_dir)

    from pyrevit import HOST_APP
    from tools import job_io
    from tools import upgrade_core

    app = HOST_APP.app
    uiapp = HOST_APP.uiapp
    out_dir = job.get("out_dir") or ""
    audit = bool(job.get("audit"))
    compact = bool(job.get("compact", True))

    results = []
    with upgrade_core.DialogSuppressor(uiapp, app):
        for record in job.get("files") or []:
            action, dst, reason = upgrade_core.plan_for(record, year, out_dir)
            if action == upgrade_core.ACTION_SKIP:
                results.append({"path": record.get("path"), "name": record.get("name"),
                                "ok": False, "skipped": True,
                                "dst": None, "message": reason})
                continue
            ok, message = upgrade_core.upgrade_one(
                app, record["path"], dst, audit=audit, compact=compact)
            results.append({"path": record.get("path"), "name": record.get("name"),
                            "ok": ok, "skipped": False,
                            "dst": dst if ok else None, "message": message})

    job_io.write_result(year, results)


try:
    main()
except Exception:
    try:
        _fail(_host_year(), traceback.format_exc())
    except Exception:
        _fail(0, traceback.format_exc())
