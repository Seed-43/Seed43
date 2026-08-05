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

__all__ = ["user_root", "user_path", "user_dir", "migrate", "migrate_dir",
           "seed_once", "seed_from_defaults"]


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


# ── MIGRATION ───────────────────────────────────────────────────────────────

def migrate(legacy_path, new_path, once_marker=None):
    """
    Move a legacy file to new_path once, and return new_path.

    Does nothing if new_path already exists - the user's current data always
    wins over whatever the updater may have restored to the old location.

    once_marker additionally stops the move ever happening a second time, so
    a shipped default the user deleted is not silently reinstated on the next
    update. Use it for a shipped file the user is allowed to remove entirely
    (pyTransmit's logo.png); omit it for a config that should simply be
    recreated from the shipped copy if it goes missing.
    After a successful copy the original is deleted, so the old location
    drains and cannot drift out of sync.

    Deliberately never raises: a tool must still open if its settings cannot
    be moved. A failed copy leaves the legacy file untouched, and a failed
    delete just leaves a harmless stale copy that the next run retries.

    If .user cannot be created at all - blocked, read-only, a file sitting
    where the folder should be - this falls back to legacy_path, so the tool
    keeps reading and writing exactly where it always did rather than failing.
    """
    try:
        if once_marker and os.path.isfile(once_marker):
            return new_path
        if os.path.isfile(new_path):
            return new_path

        # Bail out to the old location if the destination is unusable.
        if not _try_makedirs(os.path.dirname(new_path)):
            return legacy_path

        if not os.path.isfile(legacy_path):
            return new_path
        shutil.copy2(legacy_path, new_path)
    except Exception:
        return legacy_path if os.path.isfile(legacy_path) else new_path

    # Only remove the original once the copy is definitely in place.
    try:
        if os.path.isfile(new_path):
            os.remove(legacy_path)
    except Exception:
        pass

    if once_marker:
        _write_marker(once_marker)
    return new_path


def migrate_dir(legacy_dir, new_dir, suffix=".json", once_marker=None):
    """
    Move every matching file out of legacy_dir into new_dir, once each.

    The folder equivalent of migrate(), for tools keeping a directory of files
    rather than one settings file. A name already present in new_dir is left
    alone - the user's copy always wins - and only files actually copied are
    removed from legacy_dir.

    once_marker makes the whole migration fire only once, ever. Needed when
    legacy_dir is also the shipped folder: without it, the updater restores a
    template the user deleted, migration finds it missing from new_dir and
    faithfully moves it back - so deletions never stick. With the marker,
    later runs ignore the folder entirely, which is what makes a deleted
    template stay deleted through every future update.

    Returns how many files moved. Never raises.
    """
    moved = 0
    try:
        if once_marker and os.path.isfile(once_marker):
            return 0
        if not os.path.isdir(legacy_dir):
            return 0
        if not _try_makedirs(new_dir):
            return 0
        for name in os.listdir(legacy_dir):
            if suffix and not name.lower().endswith(suffix):
                continue
            src = os.path.join(legacy_dir, name)
            dst = os.path.join(new_dir, name)
            if not os.path.isfile(src) or os.path.isfile(dst):
                continue
            try:
                shutil.copy2(src, dst)
                os.remove(src)
                moved += 1
            except Exception:
                continue
        # Stamped even when nothing moved: "this collection has been dealt
        # with" is the fact worth recording, not "files were copied".
        if once_marker:
            _write_marker(once_marker)
    except Exception:
        pass
    return moved


def migrate_tree(legacy_dir, new_dir):
    """
    Move a whole folder tree into .user, once, preserving its structure.

    For tools that keep a userdata/ folder with subfolders rather than a flat
    list. Files already present in new_dir are left alone - the user's copy
    always wins - and only files actually copied are removed, so anything that
    fails to move stays put and is retried next run.

    Empty folders left behind are pruned, but legacy_dir itself is kept: on an
    existing install the updater may recreate it anyway, and removing it buys
    nothing.

    Returns how many files moved. Never raises.
    """
    moved = 0
    try:
        if not os.path.isdir(legacy_dir):
            return 0
        if not _try_makedirs(new_dir):
            return 0

        for root, dirs, files in os.walk(legacy_dir):
            rel = os.path.relpath(root, legacy_dir)
            target = new_dir if rel == "." else os.path.join(new_dir, rel)
            if not _try_makedirs(target):
                continue
            for name in files:
                src = os.path.join(root, name)
                dst = os.path.join(target, name)
                if os.path.isfile(dst):
                    continue
                try:
                    shutil.copy2(src, dst)
                    os.remove(src)
                    moved += 1
                except Exception:
                    continue

        # Prune folders that are now empty, deepest first.
        for root, dirs, files in os.walk(legacy_dir, topdown=False):
            if root == legacy_dir:
                continue
            try:
                if not os.listdir(root):
                    os.rmdir(root)
            except Exception:
                pass
    except Exception:
        pass
    return moved


# ── SHIPPED DEFAULTS ────────────────────────────────────────────────────────

def seed_once(marker_path, defaults_dir, target_dir, suffix=".json", names=None):
    """
    Copy shipped defaults into target_dir, but only the very first time.

    marker_path records an EVENT - "seeding happened" - not a state. Checking
    "is target_dir empty?" instead cannot tell a brand new user apart from one
    who deleted the demo deliberately, so it would keep resurrecting files
    they threw away. Once the marker exists the answer is permanently no,
    whatever they have since done to the folder.

    So a deleted default never comes back on update, which is the point - but
    it also cannot be recovered. Pair this with an explicit "load demo data"
    action via seed_from_defaults() to give a way back.

    Returns True if seeding ran. Never raises.
    """
    try:
        if os.path.isfile(marker_path):
            return False
        seed_from_defaults(defaults_dir, target_dir, suffix, names)
        _write_marker(marker_path)
        return True
    except Exception:
        return False


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


def _write_marker(marker_path):
    """Stamp the marker with the extension version that seeded.

    Costs nothing now, and lets a later version seed a newly added default for
    existing users without resurrecting one they already deleted.
    """
    version = ""
    try:
        vf = os.path.join(_EXTENSION_ROOT, "version.txt")
        if os.path.isfile(vf):
            with open(vf, "r") as f:
                version = f.read().strip()
    except Exception:
        pass
    try:
        _try_makedirs(os.path.dirname(marker_path))
        with open(marker_path, "w") as f:
            f.write('{"seeded_version": "%s"}' % version)
    except Exception:
        pass
