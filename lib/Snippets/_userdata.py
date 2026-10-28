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

Typical use in a tool:

    from Snippets import _userdata

    JSON_PATH = _userdata.user_path("Units", "project_units.json")

snippets.yaml entry:
  _userdata.py:
    description: >
      Per-tool user settings under a single .user folder at the extension
      root, so the updater never has to protect scattered .json files.
    functions:
      user_root: Return the .user folder, creating it if needed.
      user_path: Return a path inside .user/<tool>/, creating parent folders.
"""

import os
import shutil

__all__ = ["user_root", "user_path", "user_dir", "seed_from_defaults"]


# ── LOCATION ────────────────────────────────────────────────────────────────

# lib/Snippets/_userdata.py -> up three levels is Seed43.extension
_EXTENSION_ROOT = os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.abspath(__file__))))

USER_DIRNAME = ".user"


def _try_makedirs(path):
    """Create path if missing. Return True if it exists afterwards."""
    try:
        if path and not os.path.isdir(path):
            os.makedirs(path)
        return os.path.isdir(path)
    except Exception:
        return False


def user_root():
    """
    Return the .user folder at the extension root, creating it if needed.

    Never raises. Tools resolve their settings path at module scope, so an
    unwritable or blocked location must not stop the tool loading - the path
    comes back either way and the failure surfaces later, at save time.
    """
    root = os.path.join(_EXTENSION_ROOT, USER_DIRNAME)
    _try_makedirs(root)
    return root


def user_path(tool, *parts):
    """
    Return a path inside .user/<tool>/, creating the parent folders.

    tool is the folder name for the tool, normally its pushbutton name without
    the suffix ("pyTransmit", "pySheets"). parts are further path segments, so
    a tool with its own structure can keep it:

        user_path("pySheets", "settings", "lastsession.json")

    Creates directories but never the file itself, so callers can still test
    os.path.isfile() to detect first run. Never raises, for the same reason as
    user_root().
    """
    full = os.path.join(user_root(), tool, *parts)
    _try_makedirs(os.path.dirname(full))
    return full


def user_dir(tool, *parts):
    """
    Like user_path(), but the result is itself a folder, so it is created
    rather than just its parent. For tools that keep a directory of files:

        user_dir("pyFilter", "templates")
    """
    full = os.path.join(user_root(), tool, *parts)
    _try_makedirs(full)
    return full


# ── SHIPPED DEFAULTS ────────────────────────────────────────────────────────

def seed_from_defaults(defaults_dir, target_dir, suffix=".json", names=None):
    """
    Copy shipped defaults into target_dir, skipping names already present.

    Unconditional, with no marker involved, so it can back a "load demo data"
    button. Never overwrites what the user already has, so pressing it twice
    (or after editing a demo) is harmless.

    names limits the copy to specific filenames. Needed where the defaults
    folder is a live tool folder holding more than just templates - pyTransmit
    ships branding.json beside its vocabularies, and one person's branding is
    not a sensible starting point for everyone else.

    Returns how many files were copied. Never raises.
    """
    copied = 0
    wanted = set(names) if names else None
    try:
        if not os.path.isdir(defaults_dir):
            return 0
        if not _try_makedirs(target_dir):
            return 0
        for name in sorted(os.listdir(defaults_dir)):
            if suffix and not name.lower().endswith(suffix):
                continue
            if wanted is not None and name not in wanted:
                continue
            src = os.path.join(defaults_dir, name)
            dst = os.path.join(target_dir, name)
            if not os.path.isfile(src) or os.path.isfile(dst):
                continue
            try:
                shutil.copy2(src, dst)
                copied += 1
            except Exception:
                continue
    except Exception:
        pass
    return copied
