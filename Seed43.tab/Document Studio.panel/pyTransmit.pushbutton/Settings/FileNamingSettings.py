# -*- coding: utf-8 -*-
"""
FileNamingSettings.py  —  pyTransmit File Naming Settings panel controller
===========================================================================
Manages the File Naming Settings panel embedded in the main pyTransmit window.

Handles:
  - Transmittal PDF file naming template (drag-and-drop formatter tokens)
  - Live preview of the resolved filename
  - Projects root path + optional Older Jobs subfolder
  - resolve_project_folder(job_number) — range-bucket folder lookup
    Works for any folder structure where buckets are named  #XXX-{start}-{end}
    or any prefix, as long as the range is the last two numeric segments.

Config is persisted in  Settings/pytransmit_setup.json  under keys:
    transmittal_naming_template
    projects_root
    projects_older_root

Place this file in the  Settings  subfolder next to script.py.

Usage in script.py:
    from FileNamingSettings import FileNamingSettingsController
    self.filenaming_ctrl = FileNamingSettingsController(script_dir)
    self.filenaming_ctrl.attach(self)
    self.filenaming_ctrl.load_config()

Resolving a folder at publish time:
    folder = self.filenaming_ctrl.resolve_project_folder('6041')
    # Returns full path to e.g.  ...\\#JOB-6001-6100\\6041 - Project Burgundy
    # or None if not found.
"""

import os
import re
import json

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

import System.Windows as _SW

# ── Formatter token definitions (mirrors EditNamingFormats.py) ────────────────

FORMATTER_TOKENS = [
    {'template': '{proj_number}',           'desc': "Project Number e.g. 'PR2019.12'",          'color': '#665dba'},
    {'template': '{proj_name}',             'desc': "Project Name e.g. 'MY_PROJECT'",            'color': '#665dba'},
    {'template': '{proj_building_name}',    'desc': "Project Building Name e.g. 'BLDG01'",       'color': '#665dba'},
    {'template': '{proj_issue_date}',       'desc': "Project Issue Date e.g. '2019-10-12'",      'color': '#665dba'},
    {'template': '{proj_org_name}',         'desc': "Project Organization Name e.g. 'MYCOMP'",   'color': '#665dba'},
    {'template': '{proj_status}',           'desc': "Project Status e.g. 'CD100'",               'color': '#665dba'},
    {'template': '{current_date}',          'desc': "Today's Date e.g. '2019-10-12'",            'color': '#2a7a8a'},
    {'template': '{issue_date}',            'desc': "Issue Date e.g. '2019-10-12'",              'color': '#2a7a8a'},
    {'template': '{date_cc}',               'desc': "Century e.g. '20'",                         'color': '#1a6b7a'},
    {'template': '{date_yy}',               'desc': "Year (2-digit) e.g. '26'",                  'color': '#1a6b7a'},
    {'template': '{date_mm}',               'desc': "Month (2-digit) e.g. '04'",                 'color': '#1a6b7a'},
    {'template': '{date_dd}',               'desc': "Day (2-digit) e.g. '20'",                   'color': '#1a6b7a'},
    {'template': '{rev_number}',            'desc': "Revision Number e.g. '01'",                 'color': '#d84936'},
    {'template': '{rev_desc}',              'desc': "Revision Description e.g. 'ASI01'",         'color': '#d84936'},
    {'template': '{rev_date}',              'desc': "Revision Date e.g. '2019-10-12'",           'color': '#d84936'},
    {'template': '{username}',              'desc': "Active User e.g. 'jsmith'",                 'color': '#404E60'},
    {'template': '{revit_version}',         'desc': "Active Revit Version e.g. '2024'",          'color': '#404E60'},
    {'template': '{proj_param:PARAM_NAME}', 'desc': "Value of Given Project Information Parameter", 'color': '#232933'},
    {'template': '{glob_param:PARAM_NAME}', 'desc': "Value of Given Global Parameter",           'color': '#232933'},
]

DEFAULT_TEMPLATE = '{proj_number}-TRA-{current_date}.pdf'
DEFAULT_PATH_TEMPLATE = '{projects_root}\\#JOB-{bucket_min}-{bucket_max}\\{job_number}'

PATH_TOKENS = [
    {'template': '{projects_root}',         'desc': "Projects Root folder (configured above)",                                                                   'color': '#208A3C'},
    {'template': '{older_jobs_root}',       'desc': "Older Jobs Root — used only if project is not found in Projects Root",                                      'color': '#208A3C'},
    {'template': '{bucket_folder}',         'desc': "Full bucket folder name as found on disk e.g. '#JOB-4251-4300'. Simplest option — use as: {projects_root}\\{bucket_folder}\\{job_number}",  'color': '#2a7a8a'},
    {'template': '{bucket}',                'desc': "Bucket number range only e.g. '4251-4300' (no prefix). Use when you need just the numbers: {projects_root}\\PREFIX-{bucket}\\{job_number}", 'color': '#2a7a8a'},
    {'template': '{bucket_min}',            'desc': "Bucket range start number e.g. '4251'. Use with bucket_max to build folder name: #JOB-{bucket_min}-{bucket_max}",                            'color': '#2a7a8a'},
    {'template': '{bucket_max}',            'desc': "Bucket range end number e.g. '4300'. Use with bucket_min to build folder name: #JOB-{bucket_min}-{bucket_max}",                             'color': '#2a7a8a'},
    {'template': '{job_number}',            'desc': "Job/Project Number e.g. '6041'",            'color': '#665dba'},
    {'template': '{proj_number}',           'desc': "Project Number (alias for job_number)",     'color': '#665dba'},
    {'template': '{proj_name}',             'desc': "Project Name e.g. 'MY_PROJECT'",            'color': '#665dba'},
    {'template': '{current_date}',          'desc': "Today's Date e.g. '2019-10-12'",            'color': '#2a7a8a'},
    {'template': '{issue_date}',            'desc': "Issue Date e.g. '2019-10-12'",              'color': '#2a7a8a'},
    {'template': '{date_cc}',               'desc': "Century e.g. '20'",                         'color': '#1a6b7a'},
    {'template': '{date_yy}',               'desc': "Year (2-digit) e.g. '26'",                  'color': '#1a6b7a'},
    {'template': '{date_mm}',               'desc': "Month (2-digit) e.g. '04'",                 'color': '#1a6b7a'},
    {'template': '{date_dd}',               'desc': "Day (2-digit) e.g. '20'",                   'color': '#1a6b7a'},
    {'template': '{proj_param:PARAM_NAME}', 'desc': "Value of Given Project Information Parameter", 'color': '#232933'},
]

# ── Smart bucket finder ───────────────────────────────────────────────────────

def _extract_numbers(name):
    """Return all integers found in folder name."""
    return [int(x) for x in re.findall(r'\d+', name)]


def _resolve_bucket(root, job_int):
    """
    Find the best bucket folder in *root* for *job_int*.

    For folders with 2+ numbers: uses the FIRST two numbers as [min, max] range.
    e.g. '#JOB-4251-4300' → min=4251, max=4300
         '#JOB-4251-4300-ARCHIVE' → min=4251, max=4300 (trailing numbers ignored)

    For folders with 1 number: exact match first, then nearest floor.

    Returns (folder_name, bucket_min_str, bucket_max_str) or None.
    """
    if not root or not os.path.isdir(root):
        return None
    try:
        entries = [d for d in os.listdir(root)
                   if os.path.isdir(os.path.join(root, d))]
    except OSError:
        return None

    floor_match = None  # (distance, folder, min_str, max_str)

    for folder in entries:
        nums = _extract_numbers(folder)
        if not nums:
            continue

        if len(nums) >= 2:
            lo, hi = nums[0], nums[1]  # always first two only
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


# ══════════════════════════════════════════════════════════════════════════════
# PROJECT FOLDER RESOLVER  (standalone — usable without a UI)
# ══════════════════════════════════════════════════════════════════════════════

def find_project_folder(job_number, roots):
    """
    Locate the project folder for *job_number* inside one or more *roots*.

    Works with any folder naming convention — no wildcards or config needed.
    See _resolve_bucket for matching logic.

    Returns the absolute path to the project folder, or None.
    """
    job_str = str(job_number).strip()
    try:
        job_int = int(re.sub(r'\D', '', job_str))
    except ValueError:
        return None

    for root in roots:
        result = _resolve_bucket(root, job_int)
        if not result:
            continue
        bucket_name, bmin, bmax = result

        # Exact/floor match — the bucket folder IS the project folder
        if bmin == bmax:
            return os.path.join(root, bucket_name)

        # Range match — look for job subfolder inside the bucket
        bucket_path = os.path.join(root, bucket_name)
        try:
            candidates = [
                d for d in os.listdir(bucket_path)
                if os.path.isdir(os.path.join(bucket_path, d))
                and d.startswith(job_str)
            ]
        except OSError:
            continue
        if candidates:
            return os.path.join(bucket_path, candidates[0])

    return None

    return None


# ══════════════════════════════════════════════════════════════════════════════
# CONTROLLER
# ══════════════════════════════════════════════════════════════════════════════

class FileNamingSettingsController(object):
    """
    Drives the File Naming Settings panel embedded in pyTransmit.
    Config is stored in  Settings/pytransmit_setup.json.
    """

    CONFIG_FILE = 'pytransmit_setup.json'

    def __init__(self, script_dir):
        self._script_dir   = script_dir
        self._settings_dir = os.path.join(script_dir, 'Settings')
        self._config_path  = os.path.join(self._settings_dir, self.CONFIG_FILE)
        self._host         = None
        self._drop_pos      = 0
        self._path_drop_pos = 0
        self._template      = DEFAULT_TEMPLATE
        self._projects_root        = ''
        self._projects_older_root  = ''
        self._path_template        = DEFAULT_PATH_TEMPLATE
        self._live              = {}   # populated from Revit at runtime

    # ── Attach ────────────────────────────────────────────────────────────────

    def attach(self, host):
        """Attach to the host WPFWindow. Call once after WPFWindow.__init__."""
        self._host = host
        self._wire_events()

    # ── Config persistence ────────────────────────────────────────────────────

    def load_config(self):
        """Read saved config from disk and push into the panel."""
        try:
            with open(self._config_path, 'r') as f:
                cfg = json.load(f)
        except Exception:
            cfg = {}
        self._template             = cfg.get('transmittal_naming_template', DEFAULT_TEMPLATE)
        self._projects_root        = cfg.get('projects_root', '')
        self._projects_older_root  = cfg.get('projects_older_root', '')
        self._path_template        = cfg.get('output_path_template', DEFAULT_PATH_TEMPLATE)
        self._push_to_ui()
        self.refresh_live_values()

    def save_config(self):
        """Merge file naming config into pytransmit_setup.json (preserving all other keys)."""
        try:
            with open(self._config_path, 'r') as f:
                cfg = json.load(f)
        except Exception:
            cfg = {}
        cfg['transmittal_naming_template'] = self._get_from_tb('filenaming_template_tb',
                                                                DEFAULT_TEMPLATE)
        cfg['projects_root']               = self._get_from_tb('filenaming_projects_root_tb')
        cfg['projects_older_root']         = self._get_from_tb('filenaming_older_root_tb')
        cfg['output_path_template']        = self._get_from_tb('filenaming_path_template_tb',
                                                                DEFAULT_PATH_TEMPLATE)
        try:
            if not os.path.exists(self._settings_dir):
                os.makedirs(self._settings_dir)
            with open(self._config_path, 'w') as f:
                json.dump(cfg, f, indent=2)
        except Exception:
            pass
        # Keep internal state in sync
        self._template             = cfg['transmittal_naming_template']
        self._projects_root        = cfg['projects_root']
        self._projects_older_root  = cfg['projects_older_root']
        self._path_template        = cfg['output_path_template']

    # ── Public helpers ────────────────────────────────────────────────────────

    def get_template(self):
        """Return the saved naming template string."""
        return self._template

    def refresh_live_values(self):
        """
        Read project info and resolve the path preview from live Revit data.
        Call this when the panel opens (load_config calls it automatically).
        """
        import datetime
        today     = datetime.date.today()
        today_str = today.strftime('%Y-%m-%d')
        live  = {
            'current_date': today_str,
            'issue_date':   today_str,
            'date_cc':      today.strftime('%Y')[:2],
            'date_yy':      today.strftime('%y'),
            'date_mm':      today.strftime('%m'),
            'date_dd':      today.strftime('%d'),
        }

        # Read from Revit ProjectInformation via host.doc
        h = self._host
        doc = getattr(h, 'doc', None) if h else None
        if doc is not None:
            try:
                pi = doc.ProjectInformation
                def _gp(name):
                    try:
                        p = pi.LookupParameter(name)
                        if p and p.HasValue:
                            return (p.AsString() or p.AsValueString() or '').strip()
                    except Exception:
                        pass
                    return ''
                live['proj_number'] = _gp('Project Number')
                live['job_number']  = live['proj_number']
                live['proj_name']   = _gp('Project Name') or getattr(doc, 'Title', '') or ''
            except Exception:
                pass

        # Resolve bucket from actual folder structure
        job_str = live.get('proj_number', '')
        if job_str:
            try:
                job_int = int(re.sub(r'\D', '', job_str))
            except ValueError:
                job_int = None
            if job_int is not None:
                roots = [r for r in [self._projects_root, self._projects_older_root] if r]
                for root in roots:
                    result = _resolve_bucket(root, job_int)
                    if result:
                        bname, bmin, bmax = result
                        live['bucket_folder'] = bname
                        live['bucket']        = '{}-{}'.format(bmin, bmax)
                        live['bucket_min']    = bmin
                        live['bucket_max']    = bmax
                        live['active_root']   = root
                        break

        self._live = live
        self._update_preview()
        self._update_path_preview()

    def resolve_project_folder(self, job_number):
        """
        Return the absolute path to the project folder for *job_number*, or None.
        Searches projects_root first, then projects_older_root.
        """
        roots = [r for r in [self._projects_root, self._projects_older_root] if r]
        return find_project_folder(job_number, roots)

    def get_path_template(self):
        """Return the saved output path template string."""
        return self._path_template

    def resolve_output_path(self, job_number, values=None):
        """
        Resolve the output path template for *job_number*.

        *values* is an optional dict of extra token substitutions, e.g.:
            {'proj_name': 'PROJECT_BURGUNDY', 'issue_date': '2026-04-20'}

        Tokens resolved automatically:
            {projects_root}   → self._projects_root
            {older_jobs_root} → self._projects_older_root
            {bucket}          → range-bucket folder name (found by scanning roots)
            {job_number}      → job_number
            {proj_number}     → job_number  (alias)
        All other tokens are substituted from *values* if provided.

        Returns the resolved path string (tokens that couldn't be resolved
        are left as-is so the caller can handle them).
        """
        import datetime
        job_str = str(job_number).strip()

        # Try projects_root first, fall back to older_jobs_root only if not found there
        roots_ordered = [r for r in [self._projects_root, self._projects_older_root] if r]

        bucket_name = ''
        active_root = self._projects_root
        try:
            job_int = int(re.sub(r'\D', '', job_str))
        except ValueError:
            job_int = None

        if job_int is not None:
            for root in roots_ordered:
                result = _resolve_bucket(root, job_int)
                if result:
                    bucket_name, bucket_min, bucket_max = result
                    active_root = root
                    break

        # If bucket not found via _resolve_bucket, defaults are empty strings
        if not bucket_name:
            bucket_min   = ''
            bucket_max   = ''
            bucket_range = ''
        else:
            bucket_range = '{}-{}'.format(bucket_min, bucket_max)

        subs = {
            '{projects_root}':   self._projects_root,
            '{older_jobs_root}': active_root,
            '{bucket_folder}':   bucket_name,
            '{bucket}':          bucket_range,
            '{bucket_min}':      bucket_min,
            '{bucket_max}':      bucket_max,
            '{job_number}':      job_str,
            '{proj_number}':     job_str,
            '{current_date}':    datetime.date.today().strftime('%Y-%m-%d'),
            '{date_cc}':         datetime.date.today().strftime('%Y')[:2],
            '{date_yy}':         datetime.date.today().strftime('%y'),
            '{date_mm}':         datetime.date.today().strftime('%m'),
            '{date_dd}':         datetime.date.today().strftime('%d'),
        }
        if values:
            for k, v in values.items():
                key = k if k.startswith('{') else '{' + k + '}'
                subs[key] = v

        tmpl = self._path_template
        for token, val in subs.items():
            tmpl = tmpl.replace(token, val)
        return tmpl
        """Save config then navigate back to the main panel."""
        self.save_config()
        h = self._host
        if h is None:
            return
        try:
            h._show_panel('main')
        except Exception:
            pass

    # ── Internal ──────────────────────────────────────────────────────────────

    def _push_to_ui(self):
        h = self._host
        if h is None:
            return
        self._set_tb('filenaming_template_tb',       self._template)
        self._set_tb('filenaming_projects_root_tb',  self._projects_root)
        self._set_tb('filenaming_older_root_tb',     self._projects_older_root)
        self._set_tb('filenaming_path_template_tb',  self._path_template)
        self._update_preview()
        self._update_path_preview()

    def _get_from_tb(self, name, default=''):
        h = self._host
        if h is None:
            return default
        el = getattr(h, name, None)
        if el is None:
            return default
        try:
            return el.Text or default
        except Exception:
            return default

    def _set_tb(self, name, value):
        h = self._host
        if h is None:
            return
        el = getattr(h, name, None)
        if el is not None:
            try:
                el.Text = value or ''
            except Exception:
                pass

    def _update_preview(self):
        h = self._host
        if h is None:
            return
        preview_tb = getattr(h, 'filenaming_preview_tb', None)
        if preview_tb is None:
            return
        import datetime
        today = datetime.date.today()
        today_str = today.strftime('%Y-%m-%d')
        live = self._live
        sample = {
            '{proj_number}':           live.get('proj_number')       or 'J6041',
            '{proj_name}':             live.get('proj_name')         or 'PROJECT_BURGUNDY',
            '{proj_building_name}':    'BLDG01',
            '{proj_issue_date}':       today_str,
            '{proj_org_name}':         'NAGEL',
            '{proj_status}':           'CD100',
            '{current_date}':          live.get('current_date')      or today_str,
            '{issue_date}':            live.get('issue_date')        or today_str,
            '{date_cc}':               today.strftime('%Y')[:2],
            '{date_yy}':               today.strftime('%y'),
            '{date_mm}':               today.strftime('%m'),
            '{date_dd}':               today.strftime('%d'),
            '{rev_number}':            '01',
            '{rev_desc}':              'ASI01',
            '{rev_date}':              today_str,
            '{username}':              'jsmith',
            '{revit_version}':         '2024',
            '{proj_param:PARAM_NAME}': 'PARAM_VALUE',
            '{glob_param:PARAM_NAME}': 'PARAM_VALUE',
        }
        tmpl = self._get_from_tb('filenaming_template_tb', DEFAULT_TEMPLATE)
        preview = tmpl
        for token, val in sample.items():
            preview = preview.replace(token, val)
        try:
            preview_tb.Text = preview
        except Exception:
            pass

    def _update_path_preview(self):
        h = self._host
        if h is None:
            return
        preview_tb = getattr(h, 'filenaming_path_preview_tb', None)
        if preview_tb is None:
            return

        import datetime
        import System.Windows.Media as _SWM

        today      = datetime.date.today()
        today_str  = today.strftime('%Y-%m-%d')
        live       = self._live
        proj_root  = self._projects_root or r'N:\JOBS'
        older_root = self._projects_older_root or r'N:\JOBS\Older Jobs'
        active_root = live.get('active_root', proj_root)

        # Use real live values where available, fallback to illustrative dummy
        has_real_job    = bool(live.get('job_number'))
        has_real_bucket = bool(live.get('bucket_folder'))

        sample = {
            '{projects_root}':          proj_root,
            '{older_jobs_root}':        active_root,
            '{bucket_folder}':          live.get('bucket_folder') or '#JOB-4251-4300',
            '{bucket}':                 live.get('bucket')        or '4251-4300',
            '{bucket_min}':             live.get('bucket_min')    or '4251',
            '{bucket_max}':             live.get('bucket_max')    or '4300',
            '{job_number}':             live.get('job_number')    or '4286',
            '{proj_number}':            live.get('proj_number')   or '4286',
            '{proj_name}':              live.get('proj_name')     or 'PROJECT_BURGUNDY',
            '{current_date}':           today_str,
            '{issue_date}':             today_str,
            '{date_cc}':                today.strftime('%Y')[:2],
            '{date_yy}':                today.strftime('%y'),
            '{date_mm}':                today.strftime('%m'),
            '{date_dd}':                today.strftime('%d'),
            '{proj_param:PARAM_NAME}':  'PARAM_VALUE',
        }

        tmpl    = self._get_from_tb('filenaming_path_template_tb', DEFAULT_PATH_TEMPLATE)
        preview = tmpl
        for token, val in sample.items():
            preview = preview.replace(token, val)

        # Check if the resolved path actually exists on disk
        path_exists = os.path.isdir(preview)

        # Colour feedback: green = found, amber = resolved but missing, grey = not configured
        roots_set = bool(self._projects_root)
        if path_exists:
            colour = '#27AE60'   # green  — real path found on disk
            suffix = '  \u2713'  # ✓
        elif has_real_job and has_real_bucket:
            colour = '#E67E22'   # amber  — resolved but folder missing on disk
            suffix = '  \u26A0 Folder not found — check path template or create folder'
        elif has_real_job and not has_real_bucket and roots_set:
            colour = '#E67E22'   # amber  — job number found but no matching bucket
            suffix = '  \u26A0 No bucket folder found for job {} in {}'.format(
                live.get('job_number', ''), self._projects_root)
        elif not roots_set:
            colour = '#E05555'   # red    — roots not set at all
            suffix = '  \u2715 Folder scan failed — Projects Root not configured'
        else:
            colour = '#E05555'   # red    — something else wrong
            suffix = '  \u2715 Folder scan failed — check Projects Root path'

        try:
            preview_tb.Text = preview + suffix
            preview_tb.Foreground = _SWM.SolidColorBrush(
                _SWM.ColorConverter.ConvertFromString(colour))
        except Exception:
            try:
                preview_tb.Text = preview
            except Exception:
                pass

    def _wire_events(self):
        h = self._host
        if h is None:
            return

        # Build token chips directly into WrapPanel (IronPython 2 safe — no DataTemplate/binding)
        tokens_wp = getattr(h, 'filenaming_formatters_wp', None)
        if tokens_wp is not None:
            try:
                import System.Windows.Controls as _SWC
                import System.Windows.Media as _SWM
                import System.Windows as _SWW

                def _hex(s):
                    c = _SWM.ColorConverter.ConvertFromString(s)
                    return _SWM.SolidColorBrush(c)

                for t in FORMATTER_TOKENS:
                    tb = _SWC.TextBlock()
                    tb.Text               = t['template']
                    tb.FontFamily         = _SWM.FontFamily('Consolas')
                    tb.FontSize           = 11
                    tb.Foreground         = _hex('#F4FAFF')
                    tb.VerticalAlignment  = _SWW.VerticalAlignment.Center
                    tb.Margin             = _SWW.Thickness(6, 0, 6, 0)

                    chip = _SWC.Border()
                    chip.Height           = 22
                    chip.Margin           = _SWW.Thickness(0, 4, 6, 0)
                    chip.Background       = _hex(t['color'])
                    chip.CornerRadius     = _SWW.CornerRadius(5)
                    chip.Cursor           = _SWW.Input.Cursors.Hand
                    chip.ToolTip          = t['desc']
                    chip.DataContext      = _TokenItem(t['template'], t['desc'], t['color'])
                    chip.Child            = tb
                    tokens_wp.Children.Add(chip)

                tokens_wp.PreviewMouseLeftButtonDown += self._on_start_drag
            except Exception:
                pass

        # Template TextBox — live preview + drag-drop target
        tb = getattr(h, 'filenaming_template_tb', None)
        if tb is not None:
            try:
                tb.TextChanged     += self._on_template_changed
                tb.AllowDrop        = True
                tb.PreviewDrop     += self._on_stop_drag
                tb.PreviewDragOver += self._on_preview_drag
            except Exception:
                pass

        # Browse buttons + TextChanged on root fields to trigger live rescan
        for btn_name, tb_name in [
            ('filenaming_projects_root_browse_btn', 'filenaming_projects_root_tb'),
            ('filenaming_older_root_browse_btn',    'filenaming_older_root_tb'),
        ]:
            btn = getattr(h, btn_name, None)
            if btn is not None:
                def _make_handler(t=tb_name):
                    def _handler(s, e): self._browse_folder(t)
                    return _handler
                try:
                    btn.Click += _make_handler()
                except Exception:
                    pass
            # Wire TextChanged so typing a root path triggers a rescan + preview update
            root_tb = getattr(h, tb_name, None)
            if root_tb is not None:
                def _make_root_changed(t=tb_name):
                    def _on_root_changed(s, e):
                        try:
                            if t == 'filenaming_projects_root_tb':
                                self._projects_root = s.Text.strip()
                            else:
                                self._projects_older_root = s.Text.strip()
                            self.refresh_live_values()
                        except Exception:
                            pass
                    return _on_root_changed
                try:
                    root_tb.TextChanged += _make_root_changed()
                except Exception:
                    pass

        # Back button
        back_btn = getattr(h, 'filenaming_back_btn', None)
        if back_btn is not None:
            try:
                back_btn.Click += self._on_back_click
            except Exception:
                pass

        # Path template TextBox
        path_tb = getattr(h, 'filenaming_path_template_tb', None)
        if path_tb is not None:
            try:
                path_tb.TextChanged     += self._on_path_template_changed
                path_tb.AllowDrop        = True
                path_tb.PreviewDrop     += self._on_path_stop_drag
                path_tb.PreviewDragOver += self._on_path_preview_drag
            except Exception:
                pass

        # Path token chips
        path_wp = getattr(h, 'filenaming_path_formatters_wp', None)
        if path_wp is not None:
            try:
                import System.Windows.Controls as _SWC
                import System.Windows.Media as _SWM
                import System.Windows as _SWW

                def _hex2(s):
                    c = _SWM.ColorConverter.ConvertFromString(s)
                    return _SWM.SolidColorBrush(c)

                for t in PATH_TOKENS:
                    tb2 = _SWC.TextBlock()
                    tb2.Text              = t['template']
                    tb2.FontFamily        = _SWM.FontFamily('Consolas')
                    tb2.FontSize          = 11
                    tb2.Foreground        = _hex2('#F4FAFF')
                    tb2.VerticalAlignment = _SWW.VerticalAlignment.Center
                    tb2.Margin            = _SWW.Thickness(6, 0, 6, 0)

                    chip2 = _SWC.Border()
                    chip2.Height      = 22
                    chip2.Margin      = _SWW.Thickness(0, 4, 6, 0)
                    chip2.Background  = _hex2(t['color'])
                    chip2.CornerRadius = _SWW.CornerRadius(5)
                    chip2.Cursor      = _SWW.Input.Cursors.Hand
                    chip2.ToolTip     = t['desc']
                    chip2.DataContext  = _TokenItem(t['template'], t['desc'], t['color'])
                    chip2.Child       = tb2
                    path_wp.Children.Add(chip2)

                path_wp.PreviewMouseLeftButtonDown += self._on_path_start_drag
            except Exception:
                pass

    # ── Event handlers ────────────────────────────────────────────────────────

    def _on_template_changed(self, sender, args):
        self._update_preview()

    def _on_path_template_changed(self, sender, args):
        self._update_path_preview()

    def _on_start_drag(self, sender, args):
        # OriginalSource is the TextBlock inside the chip Border
        src = args.OriginalSource
        chip = getattr(src, 'Parent', None) or src
        token_item = getattr(chip, 'DataContext', None)
        if not isinstance(token_item, _TokenItem):
            # try one more level up (click on border itself)
            token_item = getattr(src, 'DataContext', None)
        if not isinstance(token_item, _TokenItem):
            return
        _SW.DragDrop.DoDragDrop(
            sender,
            _SW.DataObject('token_item', token_item),
            _SW.DragDropEffects.Copy
        )

    def _on_preview_drag(self, sender, args):
        try:
            from System.Windows.Forms import Cursor as _Cursor
            mp  = _Cursor.Position
            pt  = _SW.Point(mp.X, mp.Y)
            tb  = sender
            self._drop_pos = tb.GetCharacterIndexFromPoint(
                tb.PointFromScreen(pt), snapToText=True
            )
            tb.SelectionStart  = self._drop_pos
            tb.SelectionLength = 0
            tb.Focus()
        except Exception:
            pass
        args.Effects = _SW.DragDropEffects.Copy
        args.Handled = True

    def _on_stop_drag(self, sender, args):
        token_item = args.Data.GetData('token_item')
        if not token_item:
            return
        tb = sender
        try:
            pos     = self._drop_pos
            current = str(tb.Text)
            tb.Text = current[:pos] + token_item.template + current[pos:]
            tb.Focus()
        except Exception:
            pass
        args.Handled = True

    def _on_back_click(self, sender, args):
        self.save_and_back()

    def _on_path_start_drag(self, sender, args):
        src = args.OriginalSource
        chip = getattr(src, 'Parent', None) or src
        token_item = getattr(chip, 'DataContext', None)
        if not isinstance(token_item, _TokenItem):
            token_item = getattr(src, 'DataContext', None)
        if not isinstance(token_item, _TokenItem):
            return
        _SW.DragDrop.DoDragDrop(
            sender,
            _SW.DataObject('token_item', token_item),
            _SW.DragDropEffects.Copy
        )

    def _on_path_preview_drag(self, sender, args):
        try:
            from System.Windows.Forms import Cursor as _Cursor
            mp  = _Cursor.Position
            pt  = _SW.Point(mp.X, mp.Y)
            tb  = sender
            self._path_drop_pos = tb.GetCharacterIndexFromPoint(
                tb.PointFromScreen(pt), snapToText=True
            )
            tb.SelectionStart  = self._path_drop_pos
            tb.SelectionLength = 0
            tb.Focus()
        except Exception:
            pass
        args.Effects = _SW.DragDropEffects.Copy
        args.Handled = True

    def _on_path_stop_drag(self, sender, args):
        token_item = args.Data.GetData('token_item')
        if not token_item:
            return
        tb = sender
        try:
            pos     = self._path_drop_pos
            current = str(tb.Text)
            tb.Text = current[:pos] + token_item.template + current[pos:]
            tb.Focus()
        except Exception:
            pass
        args.Handled = True

    def _browse_folder(self, tb_name):
        try:
            from System.Windows.Forms import FolderBrowserDialog, DialogResult
            dlg = FolderBrowserDialog()
            dlg.Description      = 'Select folder'
            dlg.ShowNewFolderButton = False
            current = self._get_from_tb(tb_name)
            if current and os.path.exists(current):
                dlg.SelectedPath = current
            if dlg.ShowDialog() == DialogResult.OK:
                self._set_tb(tb_name, dlg.SelectedPath)
                self.save_config()
        except Exception:
            pass


# ── Simple data object for formatter token chips ──────────────────────────────

class _TokenItem(object):
    def __init__(self, template, desc, color='#404E60'):
        self.template = template
        self.desc     = desc
        self.color    = color
