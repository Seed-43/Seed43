# -*- coding: utf-8 -*-
# folder_preset_resolve.py
"""Pure path-resolution logic for export-folder presets — no WPF/UI code.
Kept separate so it's easy to test and reuse (e.g. from pyTransmit).
"""
import os
import os.path as op
import re
from datetime import datetime

WILDCARD = '{*}'
WILDCARD_HELP = ('{*} auto-completes a path segment: put it after a known '
                 'token (e.g. {job_number}{*}) and it is replaced by the '
                 'full matching folder name found on disk.')

DEFAULT_TEMPLATE = '{projects_root}\\#JOB-{bucket_min}-{bucket_max}\\{job_number}'

TOKENS = [
    ('{projects_root}', '#208A3C'),
    ('{bucket_folder}',  '#2a7a8a'),
    ('{bucket}',         '#2a7a8a'),
    ('{bucket_min}',     '#2a7a8a'),
    ('{bucket_max}',     '#2a7a8a'),
    ('{job_number}',     '#665dba'),
    ('{proj_number}',    '#665dba'),
    ('{proj_name}',      '#665dba'),
    ('{proj_building_name}', '#665dba'),
    ('{proj_org_name}',  '#665dba'),
    ('{proj_status}',    '#665dba'),
    ('{current_date}',   '#1a6b7a'),
    ('{issue_date}',     '#1a6b7a'),
    ('{date_cc}',        '#1a6b7a'),
    ('{date_yy}',        '#1a6b7a'),
    ('{date_mm}',        '#1a6b7a'),
    ('{date_dd}',        '#1a6b7a'),
    ('{proj_param:PARAM_NAME}', '#404E60'),
    ('{*}', '#C0392B'),
]


def _extract_numbers(name):
    return [int(x) for x in re.findall(r'\d+', name)]


def resolve_bucket(root, job_int):
    """Find the range-bucket folder in *root* containing *job_int*.

    Folders with 2+ numbers use the first two as [min, max]. Folders with
    one number match exactly, else fall back to the nearest lower number.
    Returns (folder_name, bucket_min_str, bucket_max_str) or None.
    """
    if not root or not op.isdir(root):
        return None
    try:
        entries = [d for d in os.listdir(root) if op.isdir(op.join(root, d))]
    except OSError:
        return None
    floor_match = None
    for folder in entries:
        nums = _extract_numbers(folder)
        if not nums:
            continue
        if len(nums) >= 2:
            lo, hi = nums[0], nums[1]
            if lo <= job_int <= hi:
                return (folder, str(lo), str(hi))
        else:
            n = nums[0]
            if n == job_int:
                return (folder, str(n), str(n))
            if n < job_int:
                dist = job_int - n
                if floor_match is None or dist < floor_match[0]:
                    floor_match = (dist, folder, str(n), str(n))
    if floor_match:
        return (floor_match[1], floor_match[2], floor_match[3])
    return None


def match_wildcard(parent_dir, prefix, suffix):
    """Find a folder in parent_dir starting with prefix and ending with
    suffix (case-insensitive). Returns the matched name, or None."""
    try:
        if not parent_dir or not op.isdir(parent_dir):
            return None
        entries = [d for d in os.listdir(parent_dir)
                  if op.isdir(op.join(parent_dir, d))]
    except OSError:
        return None
    pl, sl = prefix.lower(), suffix.lower()
    candidates = [d for d in entries
                 if d.lower().startswith(pl) and d.lower().endswith(sl)]
    if not candidates:
        return None
    candidates.sort(key=len)   # prefer the tightest match
    return candidates[0]


def token_values(root, project_info, username, revit_version):
    """Build the token → value dict for one resolution pass."""
    now = datetime.now()
    sub = {
        '{projects_root}': root or '',
        '{job_number}': project_info.number or '',
        '{proj_number}': project_info.number or '',
        '{proj_name}': project_info.name or '',
        '{proj_building_name}': project_info.building_name or '',
        '{proj_org_name}': project_info.org_name or '',
        '{proj_status}': project_info.status or '',
        '{current_date}': now.strftime('%Y-%m-%d'),
        '{issue_date}': project_info.issue_date or '',
        '{date_cc}': now.strftime('%Y')[:2],
        '{date_yy}': now.strftime('%y'),
        '{date_mm}': now.strftime('%m'),
        '{date_dd}': now.strftime('%d'),
        '{username}': username or '',
        '{revit_version}': revit_version or '',
    }
    bucket_name = bmin = bmax = ''
    try:
        job_int = int(re.findall(r'\d+', project_info.number or '')[0])
        result = resolve_bucket(root, job_int)
        if result:
            bucket_name, bmin, bmax = result
    except Exception:
        pass
    sub['{bucket_folder}'] = bucket_name
    sub['{bucket}']        = '{}-{}'.format(bmin, bmax) if bmin else ''
    sub['{bucket_min}']    = bmin
    sub['{bucket_max}']    = bmax
    return sub


def resolve_path(template, root, project_info, username, revit_version):
    """Resolve *template* into a real path, walking segment by segment so
    {*} can see what's already on disk in the resolved parent folder."""
    sub = token_values(root, project_info, username, revit_version)
    segments = re.split(r'[\\/]', template)
    resolved = []
    current_path = ''
    for seg in segments:
        text = seg
        for tok, val in sub.items():
            text = text.replace(tok, val)
        if WILDCARD in text:
            prefix, suffix = text.split(WILDCARD, 1)
            match = match_wildcard(current_path, prefix, suffix)
            text = match if match is not None else (prefix + suffix)
        resolved.append(text)
        current_path = op.join(current_path, text) if current_path else text
    return op.sep.join(resolved)
