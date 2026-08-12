# -*- coding: utf-8 -*-
# "Batch Upgrade / templates"
# "Seed43"
# """
# Named, saved Batch Upgrade configurations - what pySheets calls a
# "profile". Everything needed to rebuild the window's state, or to run
# the batch headlessly from a schedule: the file list, output folder,
# target Revit versions, and the audit/compact switches.
#
# Lives under .user, same convention as job_io.py's exchange files and
# pySheets' own profiles - see job_io.py's module docstring for why.
# """

# ── IMPORTS ────────────────────────────────────────────────────────────────

import os
import json

__all__ = ["templates_dir", "template_path", "list_templates",
           "save_template", "load_template", "delete_template"]


# ── CONSTANTS ──────────────────────────────────────────────────────────────

# tools/ -> Batch Upgrade.pushbutton -> Seed43.panel -> Seed43.tab ->
# Seed43.extension. Walked explicitly rather than hardcoded, same
# reasoning as job_io.py's _ROOT.
_TOOLS_DIR = os.path.dirname(os.path.abspath(__file__))
_PUSHBUTTON_DIR = os.path.dirname(_TOOLS_DIR)
_PANEL_DIR = os.path.dirname(_PUSHBUTTON_DIR)
_TAB_DIR = os.path.dirname(_PANEL_DIR)
_EXTENSION_DIR = os.path.dirname(_TAB_DIR)

_ROOT = os.path.join(_EXTENSION_DIR, ".user", "BatchUpgrade",
                     "settings", "templates")


# ── HELPERS (private to this file) ────────────────────────────────────────

def _safe_name(name):
    """A filename-safe version of a template name.

    Deliberately not pyrevit.coreutils.cleanup_filename: this file has to
    stay importable from startup.py's headless schedule path, which can't
    assume pyRevit's own libraries are on sys.path yet, the same reason
    job_io.py avoids importing pyrevit too. Just the characters Windows
    actually forbids in a filename.
    """
    bad = '<>:"/\\|?*'
    cleaned = "".join("_" if ch in bad else ch for ch in name).strip()
    return cleaned or "template"


# ── CORE LOGIC ─────────────────────────────────────────────────────────────

def templates_dir():
    if not os.path.isdir(_ROOT):
        try:
            os.makedirs(_ROOT)
        except OSError:
            pass
    return _ROOT


def template_path(name):
    return os.path.join(templates_dir(), _safe_name(name) + ".json")


def list_templates():
    """Every saved template name, alphabetical."""
    try:
        names = [os.path.splitext(f)[0] for f in os.listdir(templates_dir())
                 if f.lower().endswith(".json")]
        return sorted(names)
    except OSError:
        return []


def save_template(name, files, out_dir, targets, audit, compact):
    """Write (or overwrite) a template. Returns the path written."""
    data = {
        "files": files,
        "out_dir": out_dir,
        "targets": sorted(targets),
        "audit": bool(audit),
        "compact": bool(compact),
    }
    path = template_path(name)
    with open(path, "w") as handle:
        json.dump(data, handle, indent=2)
    return path


def load_template(name):
    """A saved template's data, or None if it can't be read."""
    path = template_path(name)
    if not os.path.isfile(path):
        return None
    try:
        with open(path, "r") as handle:
            return json.load(handle)
    except Exception:
        return None


def delete_template(name):
    path = template_path(name)
    try:
        if os.path.isfile(path):
            os.remove(path)
    except OSError:
        pass
