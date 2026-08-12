# -*- coding: utf-8 -*-
# "Batch Upgrade / upgrade_core"
# "Seed43"
# """
# The actual open-and-save-a-copy work, plus the rules deciding which files a
# given target version can accept. Imported both by the in-session run (when
# the target is the Revit you're sitting in) and by tools/headless_batch.py
# (when a different Revit has been launched to do the saving itself).
# """

# ── IMPORTS ────────────────────────────────────────────────────────────────

import os

from Autodesk.Revit.DB import (ModelPathUtils, OpenOptions, SaveAsOptions,
                               FailureProcessingResult)

__all__ = ["plan_for", "upgrade_one", "DialogSuppressor",
           "ACTION_UPGRADE", "ACTION_SKIP"]


# ── CONSTANTS ──────────────────────────────────────────────────────────────

ACTION_UPGRADE = "upgrade"
ACTION_SKIP = "skip"


# ── CORE LOGIC ─────────────────────────────────────────────────────────────

# --- What a target version can and can't accept ---

def plan_for(record, target_year, out_dir):
    """Decide what a target version should do with one scanned file.

    Returns (action, dst_path, reason). dst_path is None when skipping.
    Copies land in a per-version subfolder so several targets selected at
    once can't overwrite each other's output.

    Args:
        record: a dict from file_scan.scan_path.
        target_year (int): the Revit version that will do the saving.
        out_dir (str): the user's chosen output folder.
    """
    target_year = int(target_year)

    if record.get("error"):
        return ACTION_SKIP, None, record["error"]

    # A central model can't be copied-and-upgraded safely - doing so detaches
    # it and orphans every local, so it's refused rather than half-handled.
    if record.get("workshared"):
        return ACTION_SKIP, None, "workshared - a central model can't be batch-upgraded"

    year = record.get("year")
    if year and year > target_year:
        return (ACTION_SKIP, None,
                "saved in {} - Revit {} can't open a newer file".format(year, target_year))
    if year and year == target_year:
        return ACTION_SKIP, None, "already saved in {}".format(target_year)

    dst = os.path.join(out_dir, str(target_year), record["name"])
    return ACTION_UPGRADE, dst, ""


# --- Opening and saving ---

def _open_document(app, path, audit):
    """Open a Revit file as a background document.

    Prefers the ModelPath overload so Audit can be requested; family files
    fall back to the plain string overload, which is the one guaranteed to
    accept them across releases.
    """
    try:
        model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(path)
        opts = OpenOptions()
        opts.Audit = bool(audit)
        return app.OpenDocumentFile(model_path, opts)
    except Exception:
        return app.OpenDocumentFile(path)


def upgrade_one(app, src_path, dst_path, audit=False, compact=True):
    """Open one file in this Revit and save an upgraded copy at dst_path.

    Returns (ok, message). Never raises - a failure on one file must not
    stop the rest of the batch.
    """
    doc = None
    try:
        dst_dir = os.path.dirname(dst_path)
        if dst_dir and not os.path.isdir(dst_dir):
            os.makedirs(dst_dir)

        doc = _open_document(app, src_path, audit)
        if doc is None:
            return False, "Revit returned no document"

        save_opts = SaveAsOptions()
        save_opts.OverwriteExistingFile = True
        save_opts.Compact = bool(compact)
        try:
            save_opts.MaximumBackups = 1
        except Exception:
            pass  # not settable on every document type; harmless to skip

        doc.SaveAs(
            ModelPathUtils.ConvertUserVisiblePathToModelPath(dst_path),
            save_opts)
        return True, ""
    except Exception as err:
        return False, str(err)
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass


# ── CLASSES ────────────────────────────────────────────────────────────────

class DialogSuppressor(object):
    """Auto-dismisses Revit dialogs and warnings for the length of a batch.

    An unattended run stalls forever on the first modal dialog (missing
    links, "this file was created in an earlier version", and so on), so
    every dialog is answered OK and every warning swallowed while the batch
    is running. Both hooks are removed again on exit.
    """

    # --- construction ---
    def __init__(self, uiapp, app):
        self._uiapp = uiapp
        self._app = app
        self.dialogs_dismissed = 0

    # --- public methods ---
    def __enter__(self):
        try:
            self._uiapp.DialogBoxShowing += self._on_dialog
        except Exception:
            self._uiapp = None
        try:
            self._app.FailuresProcessing += self._on_failure
        except Exception:
            self._app = None
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        if self._uiapp is not None:
            try:
                self._uiapp.DialogBoxShowing -= self._on_dialog
            except Exception:
                pass
        if self._app is not None:
            try:
                self._app.FailuresProcessing -= self._on_failure
            except Exception:
                pass
        return False

    # --- private helpers ---
    def _on_dialog(self, sender, args):
        self.dialogs_dismissed += 1
        try:
            args.OverrideResult(1)   # IDOK
        except Exception:
            pass

    def _on_failure(self, sender, args):
        try:
            accessor = args.GetFailuresAccessor()
            accessor.DeleteAllWarnings()
            args.SetProcessingResult(FailureProcessingResult.Continue)
        except Exception:
            pass
