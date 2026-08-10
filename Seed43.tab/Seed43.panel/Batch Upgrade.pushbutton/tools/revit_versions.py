# -*- coding: utf-8 -*-
# "Batch Upgrade / revit_versions"
# "Seed43"
# """
# Discovers which Revit versions are installed and locates the pyRevit CLI,
# which is what actually drives a Revit version other than the running one.
# """

# ── IMPORTS ────────────────────────────────────────────────────────────────

import os
import re

__all__ = ["installed_versions", "version_grid", "revit_exe", "find_pyrevit_cli"]


# ── CONSTANTS ──────────────────────────────────────────────────────────────

_YEAR_RE = re.compile(r"Revit\s+(\d{4})\s*$")

# Oldest year ever shown in the picker. Revit files older than this still
# upgrade fine - this only bounds how far back the checkbox grid goes.
_GRID_FLOOR = 2022

_PROGRAM_FILES = os.environ.get("ProgramFiles", r"C:\Program Files")
_AUTODESK_DIR = os.path.join(_PROGRAM_FILES, "Autodesk")

_CLI_CANDIDATES = [
    os.path.join(os.environ.get("APPDATA", ""), "pyRevit-Master", "bin", "pyrevit.exe"),
    os.path.join(os.environ.get("PROGRAMFILES", ""), "pyRevit CLI", "bin", "pyrevit.exe"),
    os.path.join(os.environ.get("PROGRAMW6432", ""), "pyRevit CLI", "bin", "pyrevit.exe"),
]


# ── HELPERS (private to this file) ────────────────────────────────────────

def _from_program_files():
    """Year -> Revit.exe for every 'Autodesk\\Revit YYYY' folder on disk."""
    found = {}
    if not os.path.isdir(_AUTODESK_DIR):
        return found
    try:
        entries = os.listdir(_AUTODESK_DIR)
    except OSError:
        return found
    for name in entries:
        match = _YEAR_RE.match(name)
        if not match:
            continue
        exe = os.path.join(_AUTODESK_DIR, name, "Revit.exe")
        if os.path.isfile(exe):
            found[int(match.group(1))] = exe
    return found


def _from_registry():
    """Year -> Revit.exe read from HKLM, for installs outside Program Files.

    Best-effort only: the layout under SOFTWARE\\Autodesk\\Revit has changed
    shape across releases, so anything unexpected is skipped rather than
    raised - the Program Files scan is the primary source.
    """
    found = {}
    try:
        from Microsoft.Win32 import Registry
    except ImportError:
        return found
    try:
        root = Registry.LocalMachine.OpenSubKey(r"SOFTWARE\Autodesk\Revit")
        if root is None:
            return found
        for key_name in root.GetSubKeyNames():
            match = re.search(r"(\d{4})", key_name or "")
            if not match:
                continue
            year = int(match.group(1))
            if year in found:
                continue
            sub = root.OpenSubKey(key_name)
            if sub is None:
                continue
            # The install path sits either on the year key itself or one
            # level down under a product GUID, depending on the release.
            probes = [sub]
            for child_name in sub.GetSubKeyNames():
                child = sub.OpenSubKey(child_name)
                if child is not None:
                    probes.append(child)
            for probe in probes:
                location = probe.GetValue("InstallationLocation")
                if not location:
                    continue
                exe = os.path.join(str(location), "Revit.exe")
                if os.path.isfile(exe):
                    found[year] = exe
                    break
    except Exception:
        return found
    return found


# ── CORE LOGIC ─────────────────────────────────────────────────────────────

def installed_versions():
    """Return {year (int): path to Revit.exe} for every install found."""
    found = _from_registry()
    found.update(_from_program_files())   # disk scan wins over a stale registry
    return found


def revit_exe(year):
    """Return the Revit.exe path for a year, or None if not installed."""
    return installed_versions().get(int(year))


def version_grid(running_year):
    """Return [(year, is_installed, is_running)] for the target picker.

    Spans _GRID_FLOOR (or the oldest install, if older) up to one year past
    the running version, so a not-yet-installed next release still shows as
    a greyed-out row rather than silently missing.
    """
    running_year = int(running_year)
    found = installed_versions()
    years = set(found)
    years.add(running_year)
    low = min([_GRID_FLOOR] + list(years))
    high = max([running_year + 1] + list(years))
    return [(y, y in found or y == running_year, y == running_year)
            for y in range(low, high + 1)]


def find_pyrevit_cli():
    """Return the pyrevit.exe path used to drive other Revit versions.

    Returns None if the CLI isn't installed, which is what makes targeting
    anything other than the running version impossible.
    """
    for candidate in _CLI_CANDIDATES:
        if candidate and os.path.isfile(candidate):
            return candidate
    # Fall back to PATH.
    for folder in (os.environ.get("PATH") or "").split(os.pathsep):
        if not folder:
            continue
        candidate = os.path.join(folder.strip('"'), "pyrevit.exe")
        if os.path.isfile(candidate):
            return candidate
    return None
