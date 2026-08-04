# -*- coding: utf-8 -*-
# pyfilter_localcfg.py
# Seed43 Filter Manager - local JSON settings store.
#
# Replaces pyrevit.script.get_config()/save_config(), which write to the
# global pyRevit_config.ini in AppData that every extension on the machine
# shares. pyFilter's settings instead live in one JSON file inside this
# extension's folder - self-contained, easy to find, back up, or delete.
#
# File layout (settings.json, next to script.py):
# {
#     "columns": {"col_widths": [200, 80, ...]},
#     "sync":    {"server_path": "...", "synced": [...], "export_path": "...",
#                 "export_auto": false, "import_path": "...", "import_auto": false}
# }
# pylint: disable=import-error,invalid-name,broad-except

import os
import json

_SETTINGS_FILENAME = "settings.json"

# Cache of the loaded settings dict, keyed by the resolved file path, so
# repeated get/set calls in one session don't re-read the file every time.
_cache = {}


def _settings_path(anchor_dir):
    """Resolve the settings.json path. anchor_dir is any folder this
    extension already knows about (e.g. SCRIPT_DIR or templates_folder);
    we walk up to find the pyFilter.pushbutton root and place the file
    there, next to script.py."""
    d = os.path.abspath(anchor_dir)
    # If anchor_dir is the templates folder itself, its parent is the
    # pushbutton root where script.py lives.
    if os.path.basename(d).lower() == "templates":
        d = os.path.dirname(d)
    return os.path.join(d, _SETTINGS_FILENAME)


def _load(anchor_dir):
    path = _settings_path(anchor_dir)
    if path in _cache:
        return _cache[path], path
    data = {}
    try:
        if os.path.isfile(path):
            with open(path, "r") as f:
                data = json.load(f)
            if not isinstance(data, dict):
                data = {}
    except Exception:
        data = {}
    _cache[path] = data
    return data, path


def _save(anchor_dir, data):
    path = _settings_path(anchor_dir)
    _cache[path] = data
    try:
        tmp = path + ".tmp"
        with open(tmp, "w") as f:
            json.dump(data, f, indent=2, sort_keys=True)
        # Atomic-ish replace: write to temp then rename over the original.
        if os.path.isfile(path):
            os.remove(path)
        os.rename(tmp, path)
    except Exception:
        pass


def get_section(anchor_dir, section):
    """Return the dict stored under `section`, or {} if absent."""
    data, _ = _load(anchor_dir)
    val = data.get(section)
    return dict(val) if isinstance(val, dict) else {}


def get_option(anchor_dir, section, option, default=None):
    sec = get_section(anchor_dir, section)
    return sec.get(option, default)


def set_option(anchor_dir, section, option, value):
    data, path = _load(anchor_dir)
    sec = data.get(section)
    if not isinstance(sec, dict):
        sec = {}
    sec[option] = value
    data[section] = sec
    _save(anchor_dir, data)
