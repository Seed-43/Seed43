# -*- coding: utf-8 -*-
# "Batch Upgrade / file_scan"
# "Seed43"
# """
# Collects Revit files from folders and reads what version each was saved in,
# without opening them - BasicFileInfo only parses the file header.
# """

# ── IMPORTS ────────────────────────────────────────────────────────────────

import os
import re

from Autodesk.Revit.DB import BasicFileInfo

__all__ = ["REVIT_EXTENSIONS", "scan_path", "collect_folder", "describe"]


# ── CONSTANTS ──────────────────────────────────────────────────────────────

REVIT_EXTENSIONS = (".rvt", ".rfa", ".rte", ".rft")

# Revit writes its own backups as "name.0001.rvt" next to the original.
# Upgrading those is never wanted and would triple the run time.
_BACKUP_RE = re.compile(r"\.\d{4}\.(rvt|rfa|rte|rft)$", re.IGNORECASE)

_VERSION_RE = re.compile(r"(20\d{2})")


# ── HELPERS (private to this file) ────────────────────────────────────────

def _saved_in_year(info):
    """Pull a 4-digit year out of BasicFileInfo.

    The property carrying it moved around across releases (Format on current
    versions, SavedInVersion on older ones, and either can be a bare "2022"
    or a full "Autodesk Revit 2022"), so probe both and regex the year out
    rather than trusting one attribute.
    """
    for attr in ("Format", "SavedInVersion"):
        try:
            raw = getattr(info, attr, None)
        except Exception:
            continue
        if not raw:
            continue
        match = _VERSION_RE.search(str(raw))
        if match:
            return int(match.group(1))
    return None


def _flag(info, attr):
    """Read a bool off BasicFileInfo without caring if it exists."""
    try:
        return bool(getattr(info, attr, False))
    except Exception:
        return False


# ── CORE LOGIC ─────────────────────────────────────────────────────────────

def scan_path(path):
    """Return a dict describing one Revit file.

    Keys: path, name, ext, year (int or None), workshared, central,
    error (str or None). Never raises - an unreadable file comes back with
    error set so the UI can list it as skipped instead of failing the scan.
    """
    record = {
        "path": path,
        "name": os.path.basename(path),
        "ext": os.path.splitext(path)[1].lower(),
        "year": None,
        "workshared": False,
        "central": False,
        "error": None,
    }
    try:
        info = BasicFileInfo.Extract(path)
    except Exception as err:
        record["error"] = "Not a readable Revit file ({})".format(err)
        return record
    if info is None:
        record["error"] = "Not a readable Revit file"
        return record

    record["year"] = _saved_in_year(info)
    record["workshared"] = _flag(info, "IsWorkshared")
    record["central"] = _flag(info, "IsCentral")
    return record


def collect_folder(folder, recursive=True):
    """Return sorted Revit file paths under a folder, skipping backups."""
    hits = []
    if not os.path.isdir(folder):
        return hits
    if recursive:
        for root, dirs, files in os.walk(folder):
            # Revit's own backup folders are never upgrade targets.
            dirs[:] = [d for d in dirs if not d.lower().endswith("_backup")]
            for name in files:
                if _is_candidate(name):
                    hits.append(os.path.join(root, name))
    else:
        for name in os.listdir(folder):
            full = os.path.join(folder, name)
            if os.path.isfile(full) and _is_candidate(name):
                hits.append(full)
    return sorted(hits)


def _is_candidate(name):
    if _BACKUP_RE.search(name):
        return False
    return name.lower().endswith(REVIT_EXTENSIONS)


def describe(record):
    """Return the short bracketed tag the file list shows after the name."""
    if record["error"]:
        return "unreadable"
    if record["workshared"]:
        return "workshared"
    if record["year"]:
        return str(record["year"])
    return "unknown version"
