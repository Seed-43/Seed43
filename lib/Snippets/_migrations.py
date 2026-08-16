# -*- coding: utf-8 -*-
"""Rename-safety net for the Seed43 auto-updater.

When a tool or panel folder is renamed, seed43_migrations.yaml maps the old
path to the new one so the updater can carry the user's .json settings across
before syncing. Without an entry, the sync leaves the old config orphaned at
its old path and the renamed tool starts on defaults.

Shared by startup.py (auto-update) and About.pushbutton (manual update).

snippets.yaml entry:
  _migrations.py:
    description: >
      Reads seed43_migrations.yaml and copies user .json config from renamed
      tool folders to their new location before the updater syncs.
      Used by startup.py and About.
    functions:
      read_migrations:  Parse seed43_migrations.yaml into a list of migration dicts.
      apply_migrations: Copy .json config from each old folder path to its new one.
"""

import os
import shutil

__all__ = ["read_migrations", "apply_migrations"]


# ── PARSING ─────────────────────────────────────────────────────────────────

def read_migrations(extracted_root):
    """
    Parse seed43_migrations.yaml from the extracted repo root.

    Returns a list of dicts shaped:
      [{'from': 'rel/path', 'to': 'rel/path',
        'subfolders': [{'from': 'x', 'to': 'y'}, ...]}, ...]

    Paths are relative to Seed43.tab. Returns [] if the file is missing or
    unparseable - a failed migration must never block an update.

    Hand-rolled rather than using a yaml module: IronPython 2 has no yaml in
    the stdlib, and the file's shape is fixed and tiny. Indent <= 4 marks a
    top-level entry, deeper lines belong to the current subfolders: block.
    """
    yaml_path = os.path.join(extracted_root, "seed43_migrations.yaml")
    if not os.path.exists(yaml_path):
        return []

    migrations = []
    current    = None
    in_sub     = False
    sub_from   = None
    try:
        with open(yaml_path, "r") as f:
            lines = f.readlines()
        for line in lines:
            stripped = line.strip()
            indent   = len(line) - len(line.lstrip())
            if not stripped or stripped.startswith("#"):
                continue
            if stripped.startswith("migrations:"):
                continue
            if indent <= 4 and stripped.startswith("- from:"):
                if current:
                    migrations.append(current)
                current  = {'from': stripped[len("- from:"):].strip(),
                            'to': None, 'subfolders': []}
                in_sub   = False
                sub_from = None
            elif indent <= 4 and stripped.startswith("to:") and current and current['to'] is None:
                current['to'] = stripped[len("to:"):].strip()
            elif stripped == "subfolders:":
                in_sub = True
            elif in_sub and stripped.startswith("- from:"):
                sub_from = stripped[len("- from:"):].strip()
            elif in_sub and stripped.startswith("to:") and sub_from:
                current['subfolders'].append({
                    'from': sub_from,
                    'to':   stripped[len("to:"):].strip()
                })
                sub_from = None
        if current:
            migrations.append(current)
    except Exception:
        return []

    return [m for m in migrations if m.get('from') and m.get('to')]


# ── APPLYING ────────────────────────────────────────────────────────────────

def apply_migrations(migrations, tab_dst):
    """
    Copy .json config from each old folder path to its new one, then do the
    same for any subfolders. Paths are relative to tab_dst (Seed43.tab).

    Never overwrites an existing file at the destination, so re-running an
    update is harmless. Old folders are left for sync_tree to clean up.
    """
    for m in migrations:
        old_dir = os.path.join(tab_dst, m['from'])
        new_dir = os.path.join(tab_dst, m['to'])
        if not os.path.isdir(old_dir):
            continue
        if not os.path.isdir(new_dir):
            os.makedirs(new_dir)
        _copy_json(old_dir, new_dir)

        for sub in m.get('subfolders', []):
            old_sub = os.path.join(old_dir, sub['from'])
            new_sub = os.path.join(new_dir, sub['to'])
            if not os.path.isdir(old_sub):
                continue
            if not os.path.isdir(new_sub):
                os.makedirs(new_sub)
            _copy_json(old_sub, new_sub)


def _copy_json(src_dir, dst_dir):
    for fname in os.listdir(src_dir):
        if not fname.lower().endswith(".json"):
            continue
        src = os.path.join(src_dir, fname)
        dst = os.path.join(dst_dir, fname)
        if not os.path.exists(dst):
            shutil.copy2(src, dst)
