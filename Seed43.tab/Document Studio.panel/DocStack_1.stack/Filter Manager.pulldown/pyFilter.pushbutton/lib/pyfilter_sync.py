# -*- coding: utf-8 -*-
# sync.py
# Seed43 Filter Manager - per-template two-way sync to a configurable server
# path. Each synced template is one JSON file at server_path/<name>.json.
# Sync rule: newer mtime wins. A locally-missing file that is on the sync list
# is re-downloaded (deletion treated as accidental). Server-missing files are
# uploaded.
# pylint: disable=import-error,invalid-name,broad-except

import os
import json
import shutil

import pyfilter_localcfg as _localcfg

EXT = ".json"

# ── CONFIG ────────────────────────────────────────────────────────────────────
# Settings live in a local settings.json next to script.py (see
# pyfilter_localcfg.py) instead of pyRevit's shared global config file.

# pyfilter_sync.py lives in <pushbutton_root>/lib/, so the pushbutton root
# (where settings.json should live, next to script.py) is one level up.
_SETTINGS_ANCHOR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_CFG_SECTION = "sync"


def get_server_path():
    try:
        return _localcfg.get_option(_SETTINGS_ANCHOR, _CFG_SECTION, "server_path", "")
    except Exception:
        return ""


def set_server_path(path):
    _localcfg.set_option(_SETTINGS_ANCHOR, _CFG_SECTION, "server_path", path or "")


def get_synced_templates():
    """List of template names (without .json) the user has opted to sync."""
    try:
        data = _localcfg.get_option(_SETTINGS_ANCHOR, _CFG_SECTION, "synced", [])
        if isinstance(data, list):
            return [n for n in data if isinstance(n, str) and n]
    except Exception:
        pass
    return []


def set_synced_templates(names):
    _localcfg.set_option(_SETTINGS_ANCHOR, _CFG_SECTION, "synced",
                         sorted(set(names or [])))


def is_synced(name):
    return name in set(get_synced_templates())


def add_synced(name):
    s = set(get_synced_templates())
    s.add(name)
    set_synced_templates(list(s))


def remove_synced(name):
    s = set(get_synced_templates())
    s.discard(name)
    set_synced_templates(list(s))

# ── PATHS ─────────────────────────────────────────────────────────────────────

def server_file(server_path, name):
    return os.path.join(server_path, name + EXT)


def local_file(templates_folder, name):
    return os.path.join(templates_folder, name + EXT)


def server_ready(server_path):
    """Returns (ok, message). True only if the path exists and is writable."""
    if not server_path:
        return False, "No server path set."
    if not os.path.isdir(server_path):
        return False, "Server path does not exist: {}".format(server_path)
    if not os.access(server_path, os.W_OK):
        return False, "Server path is not writable: {}".format(server_path)
    return True, "OK"


def list_server_templates(server_path):
    if not server_path or not os.path.isdir(server_path):
        return []
    out = []
    try:
        for f in os.listdir(server_path):
            if f.lower().endswith(EXT):
                out.append(os.path.splitext(f)[0])
    except Exception:
        pass
    return sorted(out)

# ── SYNC ──────────────────────────────────────────────────────────────────────

def _safe_mtime(path):
    try:
        return os.path.getmtime(path)
    except Exception:
        return None


def _copy(src, dst):
    """Copy file preserving mtime so future syncs compare correctly."""
    shutil.copy2(src, dst)


def sync_one(name, templates_folder, server_path, logger=None):
    """
    Sync one template name. Returns a short status string:
        "uploaded" / "downloaded" / "restored" / "in-sync" / "skip:<reason>"
    """
    def log(msg):
        if logger:
            logger(msg)

    lpath = local_file(templates_folder, name)
    spath = server_file(server_path, name)
    lmt   = _safe_mtime(lpath)
    smt   = _safe_mtime(spath)

    if lmt is None and smt is None:
        return "skip:missing-both"

    # Locally missing -> deletion treated as accidental, re-download.
    if lmt is None and smt is not None:
        try:
            _copy(spath, lpath)
            log("Restored '{}' from server".format(name))
            return "restored"
        except Exception as ex:
            log("Restore failed for '{}': {}".format(name, ex))
            return "skip:restore-error"

    # Server missing -> upload local copy.
    if smt is None and lmt is not None:
        try:
            _copy(lpath, spath)
            log("Uploaded '{}' to server".format(name))
            return "uploaded"
        except Exception as ex:
            log("Upload failed for '{}': {}".format(name, ex))
            return "skip:upload-error"

    # Both exist -> newer mtime wins. 2-second tolerance for FS quirks.
    if abs(lmt - smt) < 2:
        return "in-sync"
    try:
        if lmt > smt:
            _copy(lpath, spath)
            log("Uploaded '{}' (local newer)".format(name))
            return "uploaded"
        else:
            _copy(spath, lpath)
            log("Downloaded '{}' (server newer)".format(name))
            return "downloaded"
    except Exception as ex:
        log("Sync failed for '{}': {}".format(name, ex))
        return "skip:copy-error"


def sync_all(templates_folder, logger=None):
    """
    Run sync for every name in the synced list. Returns a summary dict
    with counts and a per-name result map. Silent no-op if no server path
    is configured or it isn't reachable.
    """
    server_path = get_server_path()
    ok, msg = server_ready(server_path)
    if not ok:
        if logger:
            logger("Sync skipped: {}".format(msg))
        return {"ok": False, "reason": msg, "results": {}}

    names = get_synced_templates()
    results = {}
    for name in names:
        results[name] = sync_one(name, templates_folder, server_path, logger=logger)

    counts = {}
    for r in results.values():
        key = r.split(":")[0]
        counts[key] = counts.get(key, 0) + 1
    return {"ok": True, "reason": "", "results": results, "counts": counts}


def sync_after_save(name, templates_folder, logger=None):
    """Sync a single template after the user saves it, if it's on the list."""
    if not is_synced(name):
        return None
    server_path = get_server_path()
    ok, msg = server_ready(server_path)
    if not ok:
        if logger:
            logger("Save-sync skipped: {}".format(msg))
        return None
    return sync_one(name, templates_folder, server_path, logger=logger)

# ── EXPORT / IMPORT CONFIG ────────────────────────────────────────────────────

def get_export_path():
    try: return _localcfg.get_option(_SETTINGS_ANCHOR, _CFG_SECTION, "export_path", "")
    except Exception: return ""

def set_export_path(path):
    _localcfg.set_option(_SETTINGS_ANCHOR, _CFG_SECTION, "export_path", path or "")

def get_export_auto():
    try: return bool(_localcfg.get_option(_SETTINGS_ANCHOR, _CFG_SECTION, "export_auto", False))
    except Exception: return False

def set_export_auto(value):
    _localcfg.set_option(_SETTINGS_ANCHOR, _CFG_SECTION, "export_auto", bool(value))

def get_import_path():
    try: return _localcfg.get_option(_SETTINGS_ANCHOR, _CFG_SECTION, "import_path", "")
    except Exception: return ""

def set_import_path(path):
    _localcfg.set_option(_SETTINGS_ANCHOR, _CFG_SECTION, "import_path", path or "")

def get_import_auto():
    try: return bool(_localcfg.get_option(_SETTINGS_ANCHOR, _CFG_SECTION, "import_auto", False))
    except Exception: return False

def set_import_auto(value):
    _localcfg.set_option(_SETTINGS_ANCHOR, _CFG_SECTION, "import_auto", bool(value))
