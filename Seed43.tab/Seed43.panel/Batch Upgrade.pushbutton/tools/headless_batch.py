# -*- coding: utf-8 -*-
# "Batch Upgrade / headless_batch"
# "Seed43"
# """
# The half of the batch-upgrade flow that runs inside a Revit session
# launched for a specific year, called from Seed43.extension/startup.py's
# existing Idling handler (the same one that already drives the pySheets
# scheduler) rather than through the pyRevit CLI. This exists because
# `pyrevit run` drives Revit through a separate, precompiled runner addin
# (PyRevitRunner.dll) that can end up version-mismatched against a
# clone's own engine assemblies - see job_io.py's module docstring for
# the full story. A normal Revit launch with pyRevit attached doesn't go
# through that addin at all, so this sidesteps the problem entirely
# rather than working around it.
#
# Flow: script.py writes a job file and launches Revit.exe directly for
# the target year. startup.py's Idling handler does one cheap file check
# per tick; if a job is waiting, it calls check_and_run() here. Nothing
# in this module subscribes to Idling itself - startup.py's existing
# subscription already handles the host-specific bootstrapping (Idling
# lives in different places depending on host) and the "call on every
# tick, cheaply" discipline, so there's no reason to duplicate either.
# """

# ── IMPORTS ────────────────────────────────────────────────────────────────

import sys

from Autodesk.Revit.UI import RevitCommandId, PostableCommand

from tools import job_io

__all__ = ["check_and_run"]


# ── HELPERS (private to this file) ────────────────────────────────────────

def _host_year(app):
    return int(str(app.VersionNumber)[:4])


def _close_revit(uiapp):
    """Ask Revit to exit, the same way a user closing the window would.

    PostCommand rather than a hard process kill so any save-changes
    prompt still goes through the normal dialog machinery - which
    DialogSuppressor (used while processing the job, below) has already
    configured to answer "don't save" automatically.
    """
    exit_cmd = RevitCommandId.LookupPostableCommandId(PostableCommand.ExitRevit)
    uiapp.PostCommand(exit_cmd)


# ── CORE LOGIC ─────────────────────────────────────────────────────────────

def check_and_run(app, uiapp):
    """If a job is waiting for this Revit year, process it and close Revit.

    Cheap when there's nothing to do - job_io.read_job() is just an
    os.path.isfile() plus, only if that hits, a small JSON read. Safe to
    call on every Idling tick: the job file is deleted the instant it's
    picked up (before any of the actual work), so a second tick before
    Revit has actually finished exiting just finds nothing there.

    Returns True if a job was found and handled (Revit is now closing),
    False if there was nothing to do this tick.
    """
    year = _host_year(app)
    job = job_io.read_job(year)
    if job is None:
        return False

    job_io.delete_job(year)

    tool_dir = job.get("tool_dir")
    if tool_dir and tool_dir not in sys.path:
        sys.path.insert(0, tool_dir)
    from tools import upgrade_core  # local: needs tool_dir on sys.path first

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
    _close_revit(uiapp)
    return True
