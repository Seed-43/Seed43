# -*- coding: utf-8 -*-
"""Per-tool user data under a single .user folder at the extension root.

Settings used to live beside each tool's script, which meant a shipped default
and a user's edited copy were the same file. The updater had to special-case
every .json to avoid destroying settings, and there was no way to tell a
pristine file from a customised one.

Everything now lives in:

    Seed43.extension/.user/<Tool>/<file>

The updater only touches Seed43.tab and lib, so .user is never overwritten,
deleted or merged - no special-casing needed.

Typical use in a tool, replacing a "next to the script" path:

    from Snippets import _userdata

    JSON_PATH = _userdata.migrate(
        os.path.join(os.path.dirname(__file__), "project_units.json"),
        _userdata.user_path("Units", "project_units.json"))

migrate() moves the old file across once and deletes the original, so reads
and writes afterwards only ever touch .user. It returns the new path, so the
rest of the tool is unchanged.

snippets.yaml entry:
  _userdata.py:
    description: >
      Per-tool user settings under a single .user folder at the extension
      root, so the updater never has to protect scattered .json files.
    functions:
      user_root: Return the .user folder, creating it if needed.
      user_path: Return a path inside .user/<tool>/, creating parent folders.
      migrate:   Move a legacy file into .user once, then delete the original.
"""

import os
import shutil

__all__ = ["user_root", "user_path", "migrate"]


# ── LOCATION ────────────────────────────────────────────────────────────────

# lib/Snippets/_userdata.py -> up three levels is Seed43.extension
_EXTENSION_ROOT = os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.abspath(__file__))))

USER_DIRNAME = ".user"


def user_root():
    """Return the .user folder at the extension root, creating it if needed."""
    root = os.path.join(_EXTENSION_ROOT, USER_DIRNAME)
    if not os.path.isdir(root):
        os.makedirs(root)
    return root


def user_path(tool, *parts):
    """
    Return a path inside .user/<tool>/, creating the parent folders.

    tool is the folder name for the tool, normally its pushbutton name without
    the suffix ("pyTransmit", "PySheets"). parts are further path segments, so
    a tool with its own structure can keep it:

        user_path("PySheets", "settings", "lastsession.json")

    Creates directories but never the file itself, so callers can still test
    os.path.isfile() to detect first run.
    """
    full = os.path.join(user_root(), tool, *parts)
    parent = os.path.dirname(full)
    if parent and not os.path.isdir(parent):
        os.makedirs(parent)
    return full


# ── MIGRATION ───────────────────────────────────────────────────────────────

def migrate(legacy_path, new_path):
    """
    Move a legacy file to new_path once, and return new_path.

    Does nothing if new_path already exists - the user's current data always
    wins over whatever the updater may have restored to the old location.
    After a successful copy the original is deleted, so the old location
    drains and cannot drift out of sync.

    Deliberately never raises: a tool must still open if its settings cannot
    be moved. A failed copy leaves the legacy file untouched, and a failed
    delete just leaves a harmless stale copy that the next run retries.
    """
    try:
        if os.path.isfile(new_path):
            return new_path
        if not os.path.isfile(legacy_path):
            return new_path

        parent = os.path.dirname(new_path)
        if parent and not os.path.isdir(parent):
            os.makedirs(parent)
        shutil.copy2(legacy_path, new_path)
    except Exception:
        return new_path

    # Only remove the original once the copy is definitely in place.
    try:
        if os.path.isfile(new_path):
            os.remove(legacy_path)
    except Exception:
        pass

    return new_path
