# -*- coding: utf-8 -*-
# pySheets.py

# ── IMPORTS ──
import io          # schedule dumps come back as UTF-16, so open() won't do
import re
import os
import os.path as op
import json
from collections import namedtuple
from datetime import datetime, timedelta

from pyrevit import HOST_APP, USER_DESKTOP, framework
from pyrevit.framework import Windows, Drawing, ObjectModel, Forms, List
from pyrevit import coreutils, forms, revit, DB, script
from pyrevit.api import UI
from pyrevit.compat import get_elementid_value_func
import clr
clr.AddReference('System')
from System.ComponentModel import ListSortDirection

# Supporting tool windows now live in tools/, add it to path before importing
import sys as _sys
_tools_path = op.join(op.dirname(__file__), 'tools')
if _tools_path not in _sys.path:
    _sys.path.insert(0, _tools_path)


def _find_seed43_lib():
    """Walk up from this pushbutton to Seed43.extension/lib. pyRevit puts
    lib/ on sys.path automatically for every tool in the extension, this
    is just a fallback in case that hasn't happened yet (extension not
    reloaded since a lib file was added). Returns the path, or None."""
    folder = op.dirname(__file__)
    for _ in range(6):
        candidate = op.join(folder, 'lib')
        if op.isdir(candidate):
            return candidate
        folder = op.dirname(folder)
    return None


ABOUT_URL = "https://seed43.org/pysheets/"


_lib_path = _find_seed43_lib()
if _lib_path and _lib_path not in _sys.path:
    _sys.path.insert(0, _lib_path)

from EditNamingFormats import EditNamingFormatsWindow, NamingFormat
from Snippets import _dialogs as dlg
from Snippets import _userdata
from Snippets import _schedule
from Snippets._icons import make_icon, make_icon_with_label, set_header_icon
from Snippets._spreadsheet import write_workbook
from Snippets._support import (github_issue_url, open_folder, open_url,
                               support_mailto)
from Snippets.seed43_theme import (apply_seed43_palette, apply_seed43_dimensions,
                                   get_color)
import FolderPresetManager as fpm_win
import folder_preset_resolve as fpe_resolve
import ManageProfiles as profiles_win
import ImportExportSettings as ies_win
import ManageColumns as mc_win


# Per-format export settings modules (from settings/ subfolder)
_settings_path = op.join(op.dirname(__file__), 'settings')
if _settings_path not in _sys.path:
    _sys.path.insert(0, _settings_path)

try:
    import pdfexport as _pdf_settings
except ImportError:
    _pdf_settings = None
try:
    import dwgexport as _dwg_settings
except ImportError:
    _dwg_settings = None
try:
    import dgnexport as _dgn_settings
except ImportError:
    _dgn_settings = None
try:
    import nwcexport as _nwc_settings
except ImportError:
    _nwc_settings = None
try:
    import ifcexport as _ifc_settings
except ImportError:
    _ifc_settings = None
try:
    import imgexport as _img_settings
except ImportError:
    _img_settings = None

get_elementid_value = get_elementid_value_func()
logger = script.get_logger()
config = script.get_config()


# ── CONSTANTS ──
# Format pill names — must match XAML Button x:Name suffixes
ALL_FORMATS   = ['PDF', 'DWG', 'DGN', 'NWC', 'IFC', 'IMG', 'XLS']
VIEWS_ONLY_FORMATS = {'NWC', 'IFC'}   # these can only export 3D views, never sheets
# XLS is the mirror image: schedules and nothing else. Revit refuses to
# print or export a schedule to any of the other formats ("some of the views
# are not printable (exportable)"), and on a combined PDF one schedule fails
# the whole batch - so schedules are kept out of every other format's list
# entirely rather than being offered and then rejected.
SCHEDULES_ONLY_FORMATS = {'XLS'}
SCHEDULE_TYPE_LABEL    = 'Schedule'   # ViewItem.paper_size for a schedule
FILE_EXT      = {'PDF': '.pdf', 'DWG': '.dwg', 'DGN': '.dgn',
                 'NWC': '.nwc', 'IFC': '.ifc', 'IMG': '.png',
                 'XLS': '.xlsx'}

# Tab indices — tabs renamed: Select | Settings | Export
TAB_SELECT   = 0
TAB_EXPORT   = 1   # "Settings" tab
TAB_PRINT    = 2   # "Export" tab

# Subfolder for per-format export settings scripts
SETTINGS_FOLDER = op.join(op.dirname(__file__), 'settings')

# All user data (profiles, naming formats, settings) lives in .user, where the
# updater cannot reach it. Everything below derives from this one constant, and
# the whole legacy userdata/ tree beside the script is carried across on first
# run - structure intact, so the subpaths below are unchanged.
#
# NOTE: startup.py reads scheduled_print.json directly for the background
# scheduler, so it resolves this same location. Keep the two in step.
USERDATA_DIR = _userdata.user_dir('pySheets')
_userdata.migrate_tree(op.join(op.dirname(__file__), 'userdata'), USERDATA_DIR)
SCHEDULE_FILE = op.join(USERDATA_DIR, 'settings', 'scheduled_print.json')
CUSTOM_COLUMNS_FILE = op.join(USERDATA_DIR, 'settings', 'custom_columns.json')


# Each armed profile ("punch card") gets one entry in this file. The card's
# own timing lives in its profile JSON under "schedule"; what lands here is
# only the runtime state the scheduler needs - which document the card was
# armed against, and when it next comes due.
#
# The rules themselves live in Snippets/_schedule.py because startup.py needs
# exactly the same ones to fire a card once this window has closed. Aliased
# here so the rest of this module reads as it always did.
TS_FMT                 = _schedule.TS_FMT
SCHEDULE_GRACE_MINUTES = _schedule.GRACE_MINUTES
compute_next_run       = _schedule.compute_next_run
_parse_ts              = _schedule.parse_ts


def _xaml_path(name):
    """Absolute path to one of this tool's .xaml files.

    Bare names are resolved against this script's folder so a window opens
    the same way however it was launched - by hand, or by the scheduler from
    startup.py where the working script is somewhere else entirely."""
    if op.isabs(name) or op.isfile(name):
        return name
    return op.join(op.dirname(__file__), name)


class ExportBlocked(Exception):
    """An export could not start (no formats, no destination, missing
    folder). Raised only during a scheduled run, where the equivalent
    dialog would block instead."""


def _read_armed_file():
    """The armed cards, with every key defaulted."""
    return _schedule.read_armed_file(SCHEDULE_FILE)


def _write_armed_file(data):
    """Persist the armed cards. Silent on failure - scheduling should never
    crash the UI over a disk error."""
    if not _schedule.write_armed_file(SCHEDULE_FILE, data):
        logger.warning('Could not write schedule file: %s', SCHEDULE_FILE)


PROFILES_DIR = op.join(USERDATA_DIR, 'profiles')
NAMING_FILE  = op.join(USERDATA_DIR, 'settings', 'naming_memory.json')
COMBINED_NAMING_DIR = op.join(USERDATA_DIR, 'naming_combined')
COMBINED_DEFAULT_FORMATS = [
    NamingFormat(
        name='pySheets Combined Sheets',
        template='{current_date}_{proj_number} {proj_name}',
        builtin=True
    ),
]

# Folders backing each exportable/importable settings category —
# shared by the Export/Import Settings dialog and startup auto-import.
SETTINGS_FOLDERS = {
    'naming':          op.join(USERDATA_DIR, 'naming'),
    'naming_combined': COMBINED_NAMING_DIR,
    'profiles':        PROFILES_DIR,
}
SETTINGS_SYNC_FILE = op.join(USERDATA_DIR, 'settings_sync.json')

DEFAULT_SETUP = '<Revit Default>'

IMG_TYPE_EXT = {'PNG': '.png', 'JPEG': '.jpg', 'TIFF': '.tif'}


# ── NAMED TUPLES ──
# lightweight immutable data bags
AvailableDoc = namedtuple('AvailableDoc', ['name', 'hash', 'linked'])

SheetRevision = namedtuple('SheetRevision',
                           ['number', 'desc', 'date', 'is_set'])
UNSET_REVISION = SheetRevision(number='', desc='', date='', is_set=False)

TitleBlockPrintSettings = namedtuple('TitleBlockPrintSettings',
                                     ['psettings', 'set_by_param'])


# ── HEADER FILTER OPTION ──
# One checkbox row in a Revision/Size/Collection header filter popup
class FilterOption(forms.Reactive):
    def __init__(self, label):
        self.label = label
        self._is_checked = False   # unchecked = no filter (same as all-checked)

    @forms.reactive
    def is_checked(self):
        return self._is_checked

    @is_checked.setter
    def is_checked(self, value):
        self._is_checked = value


# ── SELECTION CONTAINER ──
# Keeps checkbox state independent of filtering / sorting / tab switches
class SelectionContainer(object):
    def __init__(self):
        self._ids = set()

    def add(self, sheet_id):
        self._ids.add(sheet_id)

    def remove(self, sheet_id):
        self._ids.discard(sheet_id)

    def clear(self):
        self._ids.clear()

    def contains(self, sheet_id):
        return sheet_id in self._ids

    def update(self, sheet_id, selected):
        if selected:
            self._ids.add(sheet_id)
        else:
            self._ids.discard(sheet_id)

    def count(self):
        return len(self._ids)

    def select_all(self, sheets):
        for s in sheets:
            self._ids.add(get_elementid_value(s.revit_sheet.Id))

    def deselect_all(self, sheets):
        for s in sheets:
            self._ids.discard(get_elementid_value(s.revit_sheet.Id))


# ── SHEET LIST ITEM ──
# One row in the Select tab DataGrid
class SheetItem(forms.Reactive):
    """Represents a single sheet row in the grid."""

    def __init__(self, revit_sheet, tblock, tblock_psettings,
                 rev_settings, selection_container):
        self._sheet      = revit_sheet
        self._tblock     = tblock
        self._container  = selection_container

        # Titleblock type (for parameter lookup)
        self._tblock_type = (revit_sheet.Document.GetElement(tblock.GetTypeId())
                             if tblock else None)

        # Basic sheet properties
        self.number     = revit_sheet.SheetNumber
        self.name       = revit_sheet.Name
        self.issue_date = (revit_sheet.Parameter[DB.BuiltInParameter.SHEET_ISSUE_DATE]
                           .AsString() or '')

        # Paper size — read from titleblock if available
        self.paper_size = self._read_paper_size()

        # Print settings
        self._tblock_psettings = tblock_psettings
        self.all_print_settings = tblock_psettings.psettings if tblock_psettings else []
        self._print_settings    = self.all_print_settings[0] if self.all_print_settings else None
        self.read_only          = tblock_psettings.set_by_param if tblock_psettings else False

        # Revision
        self.revision = self._read_revision(rev_settings)

        # Sheet Collection (Revit 2025+ native feature, blank on older)
        self.sheet_collection = self._read_sheet_collection()

        # Custom parameter columns (Manage Columns), populated on reload
        self.custom_params = {}

        # Reactive state
        self._is_selected    = False
        self._is_highlighted = False   # active multi-select group, for group drag-reorder
        self._print_filename = ''

    # ── Helpers ──
    def _read_paper_size(self):
        try:
            if self._tblock:
                w = self._tblock.get_Parameter(
                    DB.BuiltInParameter.SHEET_WIDTH).AsDouble()
                h = self._tblock.get_Parameter(
                    DB.BuiltInParameter.SHEET_HEIGHT).AsDouble()
                # crude size detection (feet)
                if   w >= 3.27 and h >= 2.30: return 'A0'
                elif w >= 2.30 and h >= 1.64: return 'A1'
                elif w >= 1.64 and h >= 1.15: return 'A2'
                elif w >= 1.15 and h >= 0.82: return 'A3'
                else:                          return 'A4'
        except Exception:
            pass
        return ''

    def _read_revision(self, rev_settings):
        try:
            per_sheet = (rev_settings.RevisionNumbering
                         == DB.RevisionNumbering.PerSheet
                         if rev_settings else False)
            cur_rev = revit.query.get_current_sheet_revision(self._sheet)
            if cur_rev:
                on_sheet = self._sheet if per_sheet else None
                return SheetRevision(
                    number=revit.query.get_rev_number(cur_rev, sheet=on_sheet),
                    desc=cur_rev.Description,
                    date=cur_rev.RevisionDate,
                    is_set=True)
        except Exception:
            pass
        return UNSET_REVISION

    def _read_sheet_collection(self):
        """Sheet Collections were added in the Revit 2025 API. Older hosts
        simply have no SheetCollectionId attribute, so this reads back ''."""
        try:
            coll_id = getattr(self._sheet, 'SheetCollectionId', None)
            if coll_id and coll_id != DB.ElementId.InvalidElementId:
                coll = self._sheet.Document.GetElement(coll_id)
                if coll:
                    return coll.Name
        except Exception:
            pass
        return ''

    # ── Read-only Revit accessors ──
    @property
    def revit_sheet(self):
        return self._sheet

    @property
    def revit_tblock(self):
        return self._tblock

    @property
    def revit_tblock_type(self):
        return self._tblock_type

    # ── Reactive properties ──
    @forms.reactive
    def is_selected(self):
        return self._is_selected

    @is_selected.setter
    def is_selected(self, value):
        self._is_selected = value
        if self._container:
            self._container.update(get_elementid_value(self._sheet.Id), value)

    @forms.reactive
    def is_highlighted(self):
        return self._is_highlighted

    @is_highlighted.setter
    def is_highlighted(self, value):
        self._is_highlighted = value

    @forms.reactive
    def print_filename(self):
        return self._print_filename

    @print_filename.setter
    def print_filename(self, value):
        self._print_filename = coreutils.cleanup_filename(value, windows_safe=True)

    @forms.reactive
    def print_settings(self):
        return self._print_settings

    @print_settings.setter
    def print_settings(self, value):
        self._print_settings = value



# ── VIEW LIST ITEM ──
# One row in the Select tab DataGrid when in Views mode
# Reuses SheetItem fields so the same DataGrid columns work
class ViewItem(forms.Reactive):
    """Represents a single non-sheet view row in the grid."""

    # Map Revit ViewType enum to a readable short string
    _VIEW_TYPE_LABELS = {
        DB.ViewType.FloorPlan:       'Floor Plan',
        DB.ViewType.CeilingPlan:     'RCP',
        DB.ViewType.Elevation:       'Elevation',
        DB.ViewType.Section:         'Section',
        DB.ViewType.Detail:          'Detail',
        DB.ViewType.ThreeD:          '3D',
        DB.ViewType.DraftingView:    'Drafting',
        DB.ViewType.Legend:          'Legend',
        DB.ViewType.Schedule:        'Schedule',
        DB.ViewType.AreaPlan:        'Area Plan',
        DB.ViewType.EngineeringPlan: 'Structural',
        DB.ViewType.Rendering:       'Rendering',
        DB.ViewType.Walkthrough:     'Walkthrough',
    }

    def __init__(self, revit_view, selection_container):
        self._view      = revit_view
        self._container = selection_container

        self.number     = ''   # views have no sheet number
        self.name       = revit_view.Name
        self.issue_date = ''
        self.paper_size = self._VIEW_TYPE_LABELS.get(
            revit_view.ViewType, str(revit_view.ViewType))
        self.revision   = UNSET_REVISION
        self.sheet_collection = ''   # views have no Sheet Collection
        self.custom_params = {}
        self.all_print_settings = []
        self.read_only          = False

        self._is_selected    = False
        self._is_highlighted = False   # active multi-select group, for group drag-reorder
        self._print_filename = ''
        self._print_settings = None

    @property
    def revit_sheet(self):
        """Alias so export code that calls .revit_sheet still works."""
        return self._view

    @property
    def revit_tblock(self):
        return None

    @property
    def revit_tblock_type(self):
        return None

    @forms.reactive
    def is_selected(self):
        return self._is_selected

    @is_selected.setter
    def is_selected(self, value):
        self._is_selected = value
        if self._container:
            self._container.update(get_elementid_value(self._view.Id), value)

    @forms.reactive
    def is_highlighted(self):
        return self._is_highlighted

    @is_highlighted.setter
    def is_highlighted(self, value):
        self._is_highlighted = value

    @forms.reactive
    def print_filename(self):
        return self._print_filename

    @print_filename.setter
    def print_filename(self, value):
        self._print_filename = coreutils.cleanup_filename(value, windows_safe=True)

    @forms.reactive
    def print_settings(self):
        return self._print_settings

    @print_settings.setter
    def print_settings(self, value):
        self._print_settings = value



class PrintSettingItem(forms.TemplateListItem):
    def __init__(self, psetting=None):
        super(PrintSettingItem, self).__init__(psetting)
        self.is_compatible = isinstance(self.item, DB.InSessionPrintSetting)

    @property
    def name(self):
        return ('<In Session>' if isinstance(self.item, DB.InSessionPrintSetting)
                else self.item.Name)

    @property
    def print_settings(self):
        return self.item

    @property
    def print_params(self):
        return self.item.PrintParameters if self.item else None

    @property
    def paper_size(self):
        try:
            return self.print_params.PaperSize if self.print_params else None
        except Exception:
            return None

    @property
    def allows_variable_paper(self):
        return False


class VariablePaperItem(PrintSettingItem):
    def __init__(self):
        PrintSettingItem.__init__(self, None)
        self.is_compatible = True

    @property
    def name(self):
        return '<Variable Paper Size>'

    @property
    def allows_variable_paper(self):
        return True


# ── Browser organization order ──
def _get_browser_order_ids(doc, for_sheets=True):
    """Return element-id-values (sheets or views) in the order the Project
    Browser currently shows them, honoring whatever Browser Organization
    (grouping/sorting rules) is active for that document. Returns None if
    it can't be determined, callers should fall back to another order."""
    try:
        if for_sheets:
            bo = DB.BrowserOrganization.GetCurrentBrowserOrganizationForSheets(doc)
        else:
            bo = DB.BrowserOrganization.GetCurrentBrowserOrganizationForViews(doc)
        if bo is None:
            return None
        ids = []

        def _walk(parent_id):
            for item in bo.GetFolderItems(parent_id):
                if item.IsFolder:
                    _walk(item.ElementId)
                else:
                    ids.append(get_elementid_value(item.ElementId))

        _walk(DB.ElementId.InvalidElementId)
        return ids
    except Exception:
        return None


# ── SHEET SET WRAPPERS ──
class AllSheetsSet(object):
    @property
    def name(self):
        return '<All Sheets>'

    def get_sheets(self, doc):
        return list(DB.FilteredElementCollector(doc)
                    .OfClass(framework.get_type(DB.ViewSheet))
                    .WhereElementIsNotElementType()
                    .ToElements())


class NamedSheetSet(object):
    def __init__(self, vss):
        self.name = vss.Name
        self._vss = vss

    def get_sheets(self, doc):
        if doc == self._vss.Document:
            return [v for v in self._vss.Views if isinstance(v, DB.ViewSheet)]
        return []


# The grid is always built from this one, whatever the dropdown is showing -
# a named set ticks rows rather than deciding which rows exist.
ALL_SHEETS = AllSheetsSet()


# ── PROFILE DROPDOWN ITEM  (both profile combos) ──
class ProfileItem(object):
    """One saved profile as shown in a dropdown. `armed` drives the accent
    dot in the item template, so the armed cards are visible without having
    to select each one in turn.

    Plain object, not forms.Reactive: the dropdowns are rebuilt wholesale
    whenever the armed set changes, so there is nothing to notify about."""

    def __init__(self, name, armed=False):
        self.name  = name
        self.armed = armed

    def __str__(self):
        return self.name


# ── PRINT QUEUE ITEM  (used in Print tab DataGrid) ──
class QueueItem(forms.Reactive):
    def __init__(self, number, filename, fmt, paper_size,
                 source=None, fname_noext=''):
        self.number      = number
        self.filename    = filename
        self.format      = fmt
        self.paper_size  = paper_size
        self.source      = source        # SheetItem / ViewItem
        self.fname_noext = fname_noext   # filename without extension
        self._status     = 'Waiting'
        self._progress   = 0

    @forms.reactive
    def status(self):
        return self._status

    @status.setter
    def status(self, value):
        self._status = value

    @forms.reactive
    def progress(self):
        return self._progress

    @progress.setter
    def progress(self, value):
        self._progress = value


# ── MAIN WINDOW ──
class PrintSheetsWindow(forms.WPFWindow):

    def __init__(self, xaml_file_name):
        self._loading = True   # suppress event handlers during init
        # Resolved against this file, not the caller's folder. pyRevit takes a
        # bare name as relative to whatever script is running, which is this
        # pushbutton when the user clicks it - but the extension root when a
        # scheduled run opens the window from startup.py, and the XAML is not
        # there.
        forms.WPFWindow.__init__(self, _xaml_path(xaml_file_name))
        apply_seed43_palette(self, op.dirname(__file__))
        # Sizing comes from the same palette as the colours. Both have to run
        # after the XAML has loaded, or the window's own Setters win.
        apply_seed43_dimensions(self, op.dirname(__file__))

        # ── Core state ──
        self._current_tab     = TAB_SELECT
        self._export_folder   = USER_DESKTOP
        self._init_psettings  = None     # to restore on close
        # Which document that setting was read from. Restoring has to target
        # that same document, not whatever the dropdown is showing when the
        # window closes - see _restore_print_settings.
        self._init_psettings_doc = None

        # ── Per-format selection containers ──
        # Sheets and Views each have their own independent containers
        # per format, so switching PDF→DWG or Sheets→Views never loses state
        self._sheet_selections = {f: SelectionContainer() for f in ALL_FORMATS}
        self._view_selections  = {f: SelectionContainer() for f in ALL_FORMATS}
        # _selection always points to the active mode+format container
        self._selection = self._sheet_selections['PDF']

        self._all_sheets      = []       # all items currently shown (sheets or views)
        self._ordered_sheets  = []       # filtered/ordered version for DataGrid
        self._all_sheet_items = []       # permanent cache of SheetItem objects
        self._all_view_items  = []       # permanent cache of ViewItem objects
        self._last_row_index  = -1       # for shift-range anchor
        self._sv_mode         = 'sheets' # 'sheets' | 'views'
        self._type_filter     = None     # None = all view types visible

        # ── Header multi-select filters (sheets mode) ──
        self._rev_filter          = None   # None = no filter, else set of labels
        self._size_filter         = None
        self._coll_filter         = None
        self._rev_filter_options  = []
        self._size_filter_options = []
        self._coll_filter_options = []

        # ── Row ordering ──
        self._sheet_order_mode   = 'ascending'  # 'browser' | 'ascending' | 'manual'
        self._view_order_mode    = 'ascending'
        self._manual_sheet_order = []    # element-id-values, user's drag-reorder sequence
        self._manual_view_order  = []
        self._updating_order_cb  = False # guard so syncing the dropdown doesn't re-trigger a sort
        self._drag_row       = None      # DataGridRow a plain click started on, pending drag/click decision
        self._drag_start_pt  = None
        self._drag_active     = False
        self._drag_preserve_grp = False  # click started on an already-highlighted row

        # ── Schedule state ──
        self._sched_timer = None
        # Timing edited since the card was last committed. Nothing is written
        # until the Schedule button is pressed, so an armed card keeps running
        # to its old time while you fiddle with the fields.
        self._sched_dirty = False
        # True only while a scheduled run is driving the window. Nothing may
        # block on a dialog then - there is no one sitting in front of it.
        self._unattended  = False
        self._running_sched = False   # re-entrancy guard, see _sched_tick

        # ── Format state ──
        # enabled: set of format strings the user has toggled on
        # viewing: which format's filename column is showing
        self._fmt_enabled = set()
        self._fmt_enabled.add('PDF')     # PDF on by default
        self._fmt_viewing = 'PDF'

        # Per-format naming format selection (name string)
        self._fmt_naming  = {f: None for f in ALL_FORMATS}

        # ── Project info ──
        # The host model, captured once. Everything that names a file or
        # builds an output path reads from here, never from the document
        # picked in the dropdown - a link's project number, name and global
        # parameters belong to whoever issued it, and filing this project's
        # output under a consultant's number is how drawings get lost.
        # A link's *sheets* are still the link's, so sheet_param and
        # tblock_param deliberately keep reading the selected document.
        self._host_doc     = revit.doc
        self._project_info = revit.query.get_project_info(doc=self._host_doc)
        self._active_folder_preset = None

        # ── Populate UI ──
        try:
            ies_win.run_auto_import(SETTINGS_FOLDERS, SETTINGS_SYNC_FILE)
        except Exception as ex:
            logger.warning('Auto-import settings failed: %s', ex)
        self._setup_custom_columns()
        self._setup_documents()
        self._setup_sheet_sets()
        self._setup_naming_formats()
        self._setup_printers()
        self._setup_print_settings()
        self._setup_export_setups()
        self._setup_profiles()
        self._setup_schedule()
        self._update_format_buttons()
        self._update_fmt_column_header()

        # Export folder display
        self.export_folder_tb.Text = self._export_folder

        # ── Extra event wiring ──
        self.sheets_dg.PreviewMouseLeftButtonDown += self._row_mouse_down
        self.sheets_dg.PreviewMouseMove           += self._row_mouse_move
        self.sheets_dg.PreviewMouseLeftButtonUp   += self._row_mouse_up
        self.sheets_dg.DragOver                   += self._row_drag_over
        self.sheets_dg.DragLeave                  += self._row_drag_leave
        self.sheets_dg.Drop                       += self._row_drop
        self.sheets_dg.AllowDrop                   = True
        self.PreviewMouseLeftButtonDown           += self._window_click
        self._setup_scroll_prevention()

        # ── Initialise tab/step state ──
        # Must run after all controls are wired so step indicators
        # and tab button Tags are set correctly from the start
        self._show_tab(TAB_SELECT)
        self._load_last_session()
        self._rebuild_folder_presets()
        self._dest_enabled = {'file'}
        self._dest_viewing = 'file'
        self._update_dest_buttons()
        self._update_dest_gates()
        self._update_dest_label()
        self.settings_toggle_btn.Content = make_icon('menu', size=18, color='#F4FAFF')
        self.win_close_btn.Content = make_icon('close', size=14, color='#F4FAFF')
        set_header_icon(self, op.dirname(__file__))
        # Search glyph sitting inside the search field. Built here rather than
        # in XAML because make_icon bakes the colour in at build time.
        self.search_icon.Content = make_icon(
            'search', size=14,
            color=get_color(op.dirname(__file__), 'text_muted',
                            fallback='#9CA3AF'))
        # Same reason: the GitHub mark on the ☰ menu is a vector icon, so it has
        # to be built here rather than declared as text in the XAML.
        self.issue_btn.Content = make_icon_with_label(
            'github', u'Report an issue on GitHub', icon_size=14,
            color=get_color(op.dirname(__file__), 'text_primary',
                            fallback='#F4FAFF'))
        self._loading = False   # init complete, events now active

    # ── PROPERTIES ──
    @property
    def _selected_doc(self):
        item = self.documents_cb.SelectedItem
        if item is None:
            return revit.doc
        for d in revit.docs:
            if d.GetHashCode() == item.hash:
                return d
        return revit.doc

    @property
    def _selected_sheetset(self):
        return self.sheetset_cb.SelectedItem

    @property
    def _selected_printer(self):
        return self.printer_cb.SelectedItem

    @property
    def _selected_printsetting(self):
        return self.printsetting_cb.SelectedItem

    @property
    def _selected_naming_format(self):
        return self.namingformat_cb.SelectedItem

    @property
    def _visible_sheets(self):
        """Sheets currently shown in the grid, in the order they're actually
        displayed. Uses Items (the DataGrid's live CollectionView) rather
        than ItemsSource, sorting by clicking a column header reorders
        Items but leaves the raw ItemsSource list untouched, so anything
        relying on visual row order (shift-click range select) needs Items."""
        items = self.sheets_dg.Items
        return list(items) if items else []

    @property
    def _checked_sheets(self):
        return [s for s in self._visible_sheets if s.is_selected]

    # ── SETUP METHODS ──
    def _setup_documents(self):
        docs = [AvailableDoc(name=revit.doc.Title,
                             hash=revit.doc.GetHashCode(),
                             linked=False)]
        for d in revit.query.get_all_linkeddocs(doc=revit.doc):
            docs.append(AvailableDoc(name=d.Title,
                                     hash=d.GetHashCode(),
                                     linked=True))
        self.documents_cb.ItemsSource = docs
        self.documents_cb.SelectedIndex = 0

    def _setup_sheet_sets(self):
        sets = [AllSheetsSet()]
        vss_col = (DB.FilteredElementCollector(self._selected_doc)
                   .OfClass(framework.get_type(DB.ViewSheetSet))
                   .WhereElementIsNotElementType()
                   .ToElements())
        sets.extend([NamedSheetSet(v) for v in vss_col])
        self.sheetset_cb.ItemsSource = sets
        self.sheetset_cb.SelectedIndex = 0
        self._reload_sheet_list()

    def _setup_naming_formats(self):
        fmts = EditNamingFormatsWindow.get_naming_formats()
        self.namingformat_cb.ItemsSource = fmts
        if not fmts:
            return
        # Per-format memory: PDF keeps x1, DWG keeps x2, until changed
        saved = {}
        try:
            if op.isfile(NAMING_FILE):
                with open(NAMING_FILE, 'r') as f:
                    saved = json.load(f).get('per_format', {})
        except Exception as ex:
            logger.warning('Naming memory load failed: %s', ex)
        names = [x.name for x in fmts]
        for f in ALL_FORMATS:
            self._fmt_naming[f] = (saved.get(f) if saved.get(f) in names
                                   else fmts[0].name)
        # Combined-PDF naming dropdown — its own independent naming store
        try:
            combined_fmts = EditNamingFormatsWindow.get_naming_formats(
                COMBINED_NAMING_DIR, COMBINED_DEFAULT_FORMATS)
            self.combined_naming_cb.ItemsSource = combined_fmts
            want = saved.get('PDF_COMBINED')
            match = next((x for x in combined_fmts if x.name == want), None)
            was_loading, self._loading = self._loading, True
            try:
                if match:
                    self.combined_naming_cb.SelectedItem = match
                else:
                    self.combined_naming_cb.SelectedIndex = 0
            finally:
                self._loading = was_loading
        except Exception as ex:
            logger.warning('Combined naming setup failed: %s', ex)
        # Show the remembered format for the current viewing format
        was_loading, self._loading = self._loading, True
        try:
            nfname = self._fmt_naming.get(self._fmt_viewing)
            match  = next((x for x in fmts if x.name == nfname), None)
            if match:
                self.namingformat_cb.SelectedItem = match
            else:
                self.namingformat_cb.SelectedIndex = 0
        finally:
            self._loading = was_loading

    def _save_naming_memory(self):
        try:
            naming_dir = op.dirname(NAMING_FILE)
            if not op.isdir(naming_dir):
                os.makedirs(naming_dir)
            data = {}
            if op.isfile(NAMING_FILE):
                with open(NAMING_FILE, 'r') as f:
                    data = json.load(f)
            data['per_format'] = dict(self._fmt_naming)
            with open(NAMING_FILE, 'w') as f:
                json.dump(data, f, indent=2)
            logger.debug('Naming memory saved: %s', data['per_format'])
        except Exception as ex:
            logger.warning('Naming memory save failed: %s', ex)

    def _get_printmanager(self):
        """The selected document's PrintManager, or None if it hasn't got one.

        Deliberately NOT refused for a linked document: pyRevit's own Print
        Sheets tool (which this one grew out of) reads PaperSizes and
        PrintSetup off a link's print manager, and drives SubmitPrint through
        it to print linked sheets. Blocking links here would take a capability
        away rather than fix anything. Callers that write to it are the ones
        that have to be careful - a link will not accept a transaction."""
        try:
            pm = self._selected_doc.PrintManager
            if pm:
                return pm
        except Exception:
            pass
        return None

    def _setup_printers(self):
        printers = list(Drawing.Printing.PrinterSettings.InstalledPrinters)
        self.printer_cb.ItemsSource = printers
        pm = self._get_printmanager()
        if pm and pm.PrinterName in printers:
            self.printer_cb.SelectedItem = pm.PrinterName
        elif printers:
            self.printer_cb.SelectedIndex = 0

    def _setup_print_settings(self):
        items = []
        if not self._selected_doc.IsLinked:
            items.append(VariablePaperItem())
        all_ps = revit.query.get_all_print_settings(doc=self._selected_doc)
        ps_items = sorted([PrintSettingItem(p) for p in all_ps],
                           key=lambda i: i.name.lower())
        items.extend(ps_items)

        # Mark compatibility. Both PrintManager reads below are wrapped: they
        # are only there to decorate and preselect the list, and a document
        # that will not answer them - a link is the case that turns up - must
        # leave you with a usable dropdown rather than take the window down.
        pm = self._get_printmanager()
        if pm:
            try:
                compat = {p.Name for p in pm.PaperSizes if p}
                for item in items:
                    if (not item.allows_variable_paper and
                            item.paper_size and item.paper_size.Name in compat):
                        item.is_compatible = True
            except Exception as ex:
                logger.warning('Could not read paper sizes from this '
                               'document: %s', ex)

        self.printsetting_cb.ItemsSource = items
        if items:
            # Try to match current Revit setting
            pm = self._get_printmanager()
            cur = None
            if pm:
                try:
                    cur = pm.PrintSetup.CurrentPrintSetting
                except Exception as ex:
                    logger.warning('Could not read the current print setting '
                                   'from this document: %s', ex)
            if cur is None:
                self.printsetting_cb.SelectedIndex = 0
            elif isinstance(cur, DB.InSessionPrintSetting):
                self.printsetting_cb.SelectedIndex = 0
            else:
                # Captured once, and only from a document that can actually
                # take it back on close. Without the "is None" guard, picking
                # a link part-way through would overwrite the host's original
                # setting and it would never be restored.
                if (self._init_psettings is None and
                        not getattr(self._selected_doc, 'IsLinked', False)):
                    self._init_psettings = cur
                    self._init_psettings_doc = self._selected_doc
                for i, item in enumerate(items):
                    if item.name == cur.Name:
                        self.printsetting_cb.SelectedIndex = i
                        break

    def _setup_export_setups(self):
        """Populate DWG/DGN export setup dropdowns from the document."""
        doc = self._selected_doc
        try:
            dwg = [DEFAULT_SETUP] + sorted(
                s.Name for s in DB.FilteredElementCollector(doc)
                .OfClass(framework.get_type(DB.ExportDWGSettings)))
            self.dwg_setup_cb.ItemsSource = dwg
            self.dwg_setup_cb.SelectedIndex = 0
        except Exception as ex:
            logger.warning('DWG setups: %s', ex)
        try:
            dgn = [DEFAULT_SETUP] + sorted(
                s.Name for s in DB.FilteredElementCollector(doc)
                .OfClass(framework.get_type(DB.ExportDGNSettings)))
            self.dgn_setup_cb.ItemsSource = dgn
            self.dgn_setup_cb.SelectedIndex = 0
        except Exception as ex:
            logger.warning('DGN setups: %s', ex)

    def _setup_scroll_prevention(self):
        """Stop combo boxes from changing value on mouse wheel."""
        for cb in [self.documents_cb, self.sheetset_cb,
                   self.printer_cb, self.printsetting_cb,
                   self.namingformat_cb, self.papersize_cb,
                   self.raster_quality_cb, self.colors_cb,
                   self.dwg_setup_cb, self.dgn_setup_cb,
                   self.ifc_version_cb, self.img_type_cb,
                   self.img_res_cb, self.profile_cb,
                   self.sched_repeat_cb,
                   self.sched_profile_cb]:
            try:
                cb.PreviewMouseWheel += self._block_cb_scroll
            except Exception:
                pass

    @staticmethod
    def _block_cb_scroll(sender, args):
        if not sender.IsDropDownOpen:
            args.Handled = True

    # ── Row ordering ──
    def _apply_order(self, items, mode, manual_order, doc, for_sheets, key_fn):
        """Return items sorted per mode. key_fn is the ascending-mode sort
        key (sheet number for sheets, name for views)."""
        if mode == 'manual' and manual_order:
            order_index = {eid: i for i, eid in enumerate(manual_order)}
            known   = [it for it in items
                       if get_elementid_value(it.revit_sheet.Id) in order_index]
            unknown = [it for it in items
                       if get_elementid_value(it.revit_sheet.Id) not in order_index]
            known.sort(key=lambda it: order_index[get_elementid_value(it.revit_sheet.Id)])
            unknown.sort(key=key_fn)
            return known + unknown

        if mode == 'browser':
            browser_ids = _get_browser_order_ids(doc, for_sheets=for_sheets)
            if browser_ids:
                order_index = {eid: i for i, eid in enumerate(browser_ids)}
                return sorted(items, key=lambda it: order_index.get(
                    get_elementid_value(it.revit_sheet.Id), len(browser_ids)))
            # Browser order unavailable (API failure), fall through to ascending

        return sorted(items, key=key_fn)

    def _detect_order_mode(self, items, doc, for_sheets, key_fn):
        """Return 'ascending', 'browser', or 'manual' depending on which
        preset order (if any) the current item sequence already matches."""
        current_ids = [get_elementid_value(it.revit_sheet.Id) for it in items]

        ascending_ids = [get_elementid_value(it.revit_sheet.Id)
                          for it in sorted(items, key=key_fn)]
        if current_ids == ascending_ids:
            return 'ascending'

        browser_ids_full = _get_browser_order_ids(doc, for_sheets=for_sheets)
        if browser_ids_full:
            present = set(current_ids)
            browser_ids = [eid for eid in browser_ids_full if eid in present]
            if current_ids == browser_ids:
                return 'browser'

        return 'manual'

    def _sync_order_dropdown(self):
        """Set order_cb to match the current mode without triggering a re-sort."""
        mode = (self._sheet_order_mode if self._sv_mode == 'sheets'
                else self._view_order_mode)
        self._updating_order_cb = True
        try:
            for item in self.order_cb.Items:
                if item.Tag == mode:
                    self.order_cb.SelectedItem = item
                    break
        finally:
            self._updating_order_cb = False

    def _order_ascending_key(self):
        """Sort key for 'ascending' mode: sheet number for sheets, name for views."""
        if self._sv_mode == 'sheets':
            return lambda it: it.number
        return lambda it: it.name

    # ── SHEET LIST ──
    def _reload_sheet_list(self):
        """Rebuild _all_sheets/_ordered_sheets from the selected doc + mode.

        Always every sheet in the document, whatever set is picked. The set
        dropdown ticks rows, it does not decide which rows exist - see
        sheetset_changed."""
        doc = self._selected_doc

        if self._sv_mode == 'views':
            self._reload_view_list(doc)
            return

        if doc is None:
            return

        tblocks = list(revit.query.get_elements_by_categories(
            [DB.BuiltInCategory.OST_TitleBlocks], doc=doc))
        rev_cfg = DB.RevisionSettings.GetRevisionSettings(doc)

        all_ps = revit.query.get_all_print_settings(doc=doc)
        sheet_ps = {}
        # Never for a link: variable paper is a printing feature and a link
        # has no printing context. The dropdown does not offer the variable
        # paper item for a link, but the host's choice is still selected at
        # the moment the document changes, so this has to be checked here
        # too rather than relying on the selection alone.
        if (not getattr(doc, 'IsLinked', False) and
                self._selected_printsetting is not None and
                self._selected_printsetting.allows_variable_paper):
            sheet_ps = self._build_sheet_ps_map(tblocks, all_ps)

        sheets = []
        for rs in ALL_SHEETS.get_sheets(doc):
            if getattr(rs, 'IsPlaceholder', False):
                continue
            tb  = self._find_tblock(rs, tblocks)
            tps = sheet_ps.get(rs.SheetNumber)
            if tps is None:
                ps = (self._selected_printsetting.print_settings
                      if (self._selected_printsetting is not None and
                          not self._selected_printsetting.allows_variable_paper)
                      else (all_ps[0] if all_ps else None))
                tps = TitleBlockPrintSettings(
                    psettings=[ps] if ps else [], set_by_param=False)
            sheets.append(SheetItem(rs, tb, tps, rev_cfg, self._selection))

        sheets = self._apply_order(sheets, self._sheet_order_mode,
                                    self._manual_sheet_order, doc, True,
                                    lambda s: s.number)
        self._all_sheets      = sheets
        self._ordered_sheets  = list(sheets)
        self._all_sheet_items = list(sheets)  # permanent cache for cross-mode printing
        for s in sheets:
            s.custom_params = self._read_custom_params(s, self._custom_sheet_columns)
        self._rebuild_rev_filter()
        self._rebuild_size_filter()
        self._rebuild_coll_filter()
        self._rebuild_extra_filters_for(self._custom_sheet_columns, self._all_sheet_items)
        self._sync_order_dropdown()
        self._apply_filter()

    def _reload_view_list(self, doc):
        """Rebuild _all_sheets/_ordered_sheets with non-sheet views."""
        # Collect all non-sheet, non-template views
        all_views = list(
            DB.FilteredElementCollector(doc)
            .OfClass(framework.get_type(DB.View))
            .WhereElementIsNotElementType()
            .ToElements()
        )
        excluded = {
            DB.ViewType.Internal,
            DB.ViewType.ProjectBrowser,
            DB.ViewType.SystemBrowser,
            DB.ViewType.Undefined,
            DB.ViewType.DrawingSheet,  # sheets handled separately
        }
        # Map view id → sheet number(s) it is placed on
        placed = {}
        try:
            for vp in (DB.FilteredElementCollector(doc)
                       .OfClass(framework.get_type(DB.Viewport))):
                vid = get_elementid_value(vp.ViewId)
                sht = doc.GetElement(vp.SheetId)
                num = sht.SheetNumber if sht else ''
                placed.setdefault(vid, []).append(num)
        except Exception as ex:
            logger.warning('Viewport scan failed: %s', ex)

        items = []
        for v in all_views:
            if v.IsTemplate:
                continue
            if v.ViewType in excluded:
                continue
            item = ViewItem(v, self._selection)
            item.placed_on = ', '.join(
                sorted(placed.get(get_elementid_value(v.Id), [])))
            items.append(item)

        items = self._apply_order(items, self._view_order_mode,
                                   self._manual_view_order, doc, False,
                                   lambda v: v.name)
        self._all_sheets     = items
        self._ordered_sheets = list(items)
        self._all_view_items = list(items)  # permanent cache for cross-mode printing
        for v in items:
            v.custom_params = self._read_custom_params(v, self._custom_view_columns)
        self._rebuild_type_filter()
        # The dropdown is fresh, so whatever lock the current format implies
        # has to be put back on it. Without this, reloading the list while
        # XLS was showing left the filter wide open and every view appeared.
        self._apply_type_lock()
        self._rebuild_extra_filters_for(self._custom_view_columns, self._all_view_items)
        self._sync_order_dropdown()
        self._apply_filter()

    def _rebuild_type_filter(self):
        """Build the Type header dropdown — single-select, one entry per view type."""
        try:
            cb = Windows.Controls.ComboBox()
            try:
                cb.Style = self.FindResource('DarkComboBox')
            except Exception:
                pass
            cb.Width = 120
            # Same list _apply_type_lock maintains, so the dropdown is built
            # and rebuilt from one rule rather than two that can disagree.
            cb.ItemsSource   = self._type_labels()
            cb.SelectedIndex = 0
            cb.SelectionChanged += self.type_filter_changed
            self._type_cb     = cb
            self._type_filter = None

            st = Windows.Controls.ComboBox()
            try:
                st.Style = self.FindResource('DarkComboBox')
            except Exception:
                pass
            st.Width = 120
            st.ItemsSource   = ['All views', 'Placed views', 'Unplaced views']
            st.SelectedIndex = 0
            st.SelectionChanged += self.view_state_changed
            self._state_cb     = st
            self._state_filter = None
            if self._sv_mode == 'views':
                self._set_size_col_header(cb)
                self._set_views_headers()
        except Exception as ex:
            logger.warning('Type filter rebuild failed: %s', ex)

    def _rebuild_header_filter(self, list_ctrl, opts_attr, labels):
        """Shared builder for the Revision/Size/Collection header filters.
        Preserves prior checked state for labels still present."""
        try:
            prior = {o.label: o.is_checked for o in getattr(self, opts_attr)}
            opts  = [FilterOption(l) for l in labels]
            for o in opts:
                if o.label in prior:
                    o.is_checked = prior[o.label]
            setattr(self, opts_attr, opts)
            list_ctrl.ItemsSource = ObjectModel.ObservableCollection[object](opts)
        except Exception as ex:
            logger.warning('Header filter rebuild failed: %s', ex)

    def _rebuild_rev_filter(self):
        labels = sorted({s.revision.number for s in self._all_sheet_items
                          if s.revision.is_set})
        if any(not s.revision.is_set for s in self._all_sheet_items):
            labels.append('<None>')
        self._rebuild_header_filter(self.rev_filter_cb, '_rev_filter_options', labels)
        self._recompute_rev_filter(apply=False)

    def _rebuild_size_filter(self):
        labels = sorted({s.paper_size for s in self._all_sheet_items if s.paper_size})
        if any(not s.paper_size for s in self._all_sheet_items):
            labels.append('<None>')
        self._rebuild_header_filter(self.size_filter_cb, '_size_filter_options', labels)
        self._recompute_size_filter(apply=False)

    def _rebuild_coll_filter(self):
        labels = sorted({s.sheet_collection for s in self._all_sheet_items
                          if s.sheet_collection})
        if any(not s.sheet_collection for s in self._all_sheet_items):
            labels.append('<None>')
        self._rebuild_header_filter(self.coll_filter_cb, '_coll_filter_options', labels)
        self._recompute_coll_filter(apply=False)

    @staticmethod
    def _compute_filter_set(opts):
        """None = no filter active (all or none checked = pass-through)."""
        checked = {o.label for o in opts if o.is_checked}
        return None if (not checked or len(checked) == len(opts)) else checked

    def _recompute_rev_filter(self, apply=True):
        self._rev_filter = self._compute_filter_set(self._rev_filter_options)
        if apply:
            self._apply_filter()

    def _recompute_size_filter(self, apply=True):
        self._size_filter = self._compute_filter_set(self._size_filter_options)
        if apply:
            self._apply_filter()

    def _recompute_coll_filter(self, apply=True):
        self._coll_filter = self._compute_filter_set(self._coll_filter_options)
        if apply:
            self._apply_filter()

    def _filter_checkbox_preview_down(self, sender, recompute_fn):
        """Standard fix for multi-select ComboBoxes: toggle the checkbox
        ourselves and mark the event handled here, before it can bubble up
        to the ComboBoxItem (which would otherwise treat the click as a
        selection and close the dropdown)."""
        sender.IsChecked = not sender.IsChecked
        recompute_fn()

    def rev_filter_checkbox_preview_down(self, sender, args):
        self._filter_checkbox_preview_down(sender, self._recompute_rev_filter)
        args.Handled = True

    def size_filter_checkbox_preview_down(self, sender, args):
        self._filter_checkbox_preview_down(sender, self._recompute_size_filter)
        args.Handled = True

    def coll_filter_checkbox_preview_down(self, sender, args):
        self._filter_checkbox_preview_down(sender, self._recompute_coll_filter)
        args.Handled = True

    def _recompute_for_combo(self, combo):
        """Route an All/None click to the right recompute — combo is
        whichever FilterComboBox instance the button's ControlTemplate
        belongs to (rev/size/collection, or one of the dynamic ones)."""
        if combo is self.rev_filter_cb:
            self._recompute_rev_filter()
        elif combo is self.size_filter_cb:
            self._recompute_size_filter()
        elif combo is self.coll_filter_cb:
            self._recompute_coll_filter()
        else:
            for key, c in self._extra_filter_combos.items():
                if c is combo:
                    self._recompute_extra_filter(key)
                    break

    def filter_all_clicked(self, sender, args):
        combo = sender.TemplatedParent
        for o in combo.ItemsSource:
            o.is_checked = True
        self._recompute_for_combo(combo)

    def filter_none_clicked(self, sender, args):
        combo = sender.TemplatedParent
        for o in combo.ItemsSource:
            o.is_checked = False
        self._recompute_for_combo(combo)

    # ── Custom parameter columns (Manage Columns) ──
    # Generalised version of the rev/size/coll filter machinery above, keyed
    # by parameter name instead of hard-coded attributes, since the set of
    # columns here is user-chosen at runtime.

    def _rebuild_extra_filter(self, key, combo, labels):
        try:
            prior = {o.label: o.is_checked
                     for o in self._extra_filter_options.get(key, [])}
            opts = [FilterOption(l) for l in labels]
            for o in opts:
                if o.label in prior:
                    o.is_checked = prior[o.label]
            self._extra_filter_options[key] = opts
            combo.ItemsSource = ObjectModel.ObservableCollection[object](opts)
            self._recompute_extra_filter(key, apply=False)
        except Exception as ex:
            logger.warning('Extra filter rebuild failed (%s): %s', key, ex)

    def _recompute_extra_filter(self, key, apply=True):
        opts = self._extra_filter_options.get(key, [])
        self._extra_filters[key] = self._compute_filter_set(opts)
        if apply:
            self._apply_filter()

    def extra_filter_checkbox_preview_down(self, sender, args):
        """Shared handler for every dynamic column's filter checkbox — the
        checkbox's Tag is bound (in ExtraColumnHeaderTemplate) to its
        ComboBox ancestor's Tag, which _build_extra_column sets to the
        parameter key, so one handler can route to the right filter."""
        key = sender.Tag
        sender.IsChecked = not sender.IsChecked
        self._recompute_extra_filter(key)
        args.Handled = True

    def _build_extra_column(self, key):
        """Create one DataGridColumn for a custom parameter, with the same
        sortable-header + multi-select-filter as Revision/Size/Collection,
        via DataTemplate.LoadContent() (safe runtime instantiation — no
        FrameworkElementFactory)."""
        template = self.FindResource('ExtraColumnHeaderTemplate')
        content  = template.LoadContent()
        title_tb = content.FindName('title_tb')
        combo    = content.FindName('filter_cb')
        title_tb.Text = key
        combo.Tag     = key

        col = Windows.Controls.DataGridTextColumn()
        col.Header      = content
        col.Binding     = Windows.Data.Binding("custom_params[{}]".format(key))
        col.SortMemberPath = "custom_params[{}]".format(key)
        col.Width       = Windows.Controls.DataGridLength(110)
        col.MinWidth    = 70
        col.IsReadOnly  = True
        self._extra_filter_combos[key] = combo
        return col

    def _sync_dynamic_columns(self):
        """Rebuild the dynamic column set to match self._custom_sheet_columns
        / self._custom_view_columns. Existing columns for keys that are
        still selected are left in place (preserving position/width)."""
        for key in list(self._dynamic_sheet_cols):
            if key not in self._custom_sheet_columns:
                self.sheets_dg.Columns.Remove(self._dynamic_sheet_cols.pop(key))
                self._extra_filter_options.pop(key, None)
                self._extra_filters.pop(key, None)
                self._extra_filter_combos.pop(key, None)
        for key in list(self._dynamic_view_cols):
            if key not in self._custom_view_columns:
                self.sheets_dg.Columns.Remove(self._dynamic_view_cols.pop(key))
                self._extra_filter_options.pop(key, None)
                self._extra_filters.pop(key, None)
                self._extra_filter_combos.pop(key, None)

        for key in self._custom_sheet_columns:
            if key not in self._dynamic_sheet_cols:
                col = self._build_extra_column(key)
                idx = list(self.sheets_dg.Columns).index(self.filename_col)
                self.sheets_dg.Columns.Insert(idx, col)
                self._dynamic_sheet_cols[key] = col
        for key in self._custom_view_columns:
            if key not in self._dynamic_view_cols:
                col = self._build_extra_column(key)
                idx = list(self.sheets_dg.Columns).index(self.filename_col)
                self.sheets_dg.Columns.Insert(idx, col)
                self._dynamic_view_cols[key] = col

        self._refresh_dynamic_column_visibility()

    @staticmethod
    def _read_custom_params(item, keys):
        result = {}
        for key in keys:
            try:
                p = item.revit_sheet.LookupParameter(key)
                result[key] = (p.AsValueString() or p.AsString() or '') if p else ''
            except Exception:
                result[key] = ''
        return result

    def _rebuild_extra_filters_for(self, keys, items):
        for key in keys:
            combo = self._extra_filter_combos.get(key)
            if combo is None:
                continue
            labels = sorted({it.custom_params.get(key, '') for it in items
                              if it.custom_params.get(key, '')})
            if any(not it.custom_params.get(key, '') for it in items):
                labels.append('<None>')
            self._rebuild_extra_filter(key, combo, labels)

    def _load_custom_columns(self):
        try:
            if op.isfile(CUSTOM_COLUMNS_FILE):
                with open(CUSTOM_COLUMNS_FILE, 'r') as f:
                    return json.load(f)
        except Exception as ex:
            logger.warning('Custom columns load failed: %s', ex)
        return {}

    def _save_custom_columns(self):
        try:
            d = op.dirname(CUSTOM_COLUMNS_FILE)
            if not op.isdir(d):
                os.makedirs(d)
            with open(CUSTOM_COLUMNS_FILE, 'w') as f:
                json.dump({
                    'builtin_visible': self._builtin_visible,
                    'sheet_columns':   self._custom_sheet_columns,
                    'view_columns':    self._custom_view_columns,
                }, f, indent=2)
        except Exception as ex:
            logger.warning('Custom columns save failed: %s', ex)

    def _setup_custom_columns(self):
        cc = self._load_custom_columns()
        self._builtin_visible       = cc.get(
            'builtin_visible', {'revision': True, 'size': True, 'collection': True})
        self._custom_sheet_columns  = cc.get('sheet_columns', [])
        self._custom_view_columns   = cc.get('view_columns', [])
        self._dynamic_sheet_cols    = {}   # param name -> DataGridColumn
        self._dynamic_view_cols     = {}   # param name -> DataGridColumn
        self._extra_filter_options  = {}   # param name -> [FilterOption]
        self._extra_filters         = {}   # param name -> set(labels) | None
        self._extra_filter_combos   = {}   # param name -> ComboBox
        self._sync_dynamic_columns()
        self._refresh_builtin_column_visibility()

    def manage_columns_clicked(self, sender, args):
        self.settings_toggle_btn.IsChecked = False
        doc = self._selected_doc
        sample_sheet = None
        sample_view  = None
        excluded_view_types = {
            DB.ViewType.Internal,
            DB.ViewType.ProjectBrowser,
            DB.ViewType.SystemBrowser,
            DB.ViewType.Undefined,
            DB.ViewType.DrawingSheet,
        }
        try:
            if doc is not None:
                sample_sheet = next(iter(
                    revit.query.get_elements_by_categories(
                        [DB.BuiltInCategory.OST_Sheets], doc=doc)), None)
                sample_view = next((
                    v for v in DB.FilteredElementCollector(doc)
                        .OfClass(framework.get_type(DB.View))
                        .WhereElementIsNotElementType()
                    if not v.IsTemplate and v.ViewType not in excluded_view_types),
                    None)
        except Exception as ex:
            logger.warning('Sample element scan failed: %s', ex)
        result = mc_win.show_manager(
            sample_sheet, sample_view,
            self._custom_sheet_columns, self._custom_view_columns,
            self._builtin_visible)
        if result is None:
            return
        self._builtin_visible      = result['builtin_visible']
        self._custom_sheet_columns = result['sheet_columns']
        self._custom_view_columns  = result['view_columns']
        self._save_custom_columns()
        self._sync_dynamic_columns()
        self._refresh_builtin_column_visibility()
        self._reload_sheet_list()

    def _refresh_builtin_column_visibility(self):
        """Revision/Collection are sheets-only and follow the user's Manage
        Columns toggle. Size doubles as the views-mode Type filter, so it
        always shows there regardless of the toggle."""
        sheets_mode = (self._sv_mode == 'sheets')
        self.revision_col.Visibility = (
            Windows.Visibility.Visible
            if sheets_mode and self._builtin_visible.get('revision', True)
            else Windows.Visibility.Collapsed)
        self.collection_col.Visibility = (
            Windows.Visibility.Visible
            if sheets_mode and self._builtin_visible.get('collection', True)
            else Windows.Visibility.Collapsed)
        self.size_col.Visibility = (
            Windows.Visibility.Visible
            if (not sheets_mode) or self._builtin_visible.get('size', True)
            else Windows.Visibility.Collapsed)

    def _refresh_dynamic_column_visibility(self):
        sheets_mode = (self._sv_mode == 'sheets')
        for col in self._dynamic_sheet_cols.values():
            col.Visibility = (Windows.Visibility.Visible if sheets_mode
                              else Windows.Visibility.Collapsed)
        for col in self._dynamic_view_cols.values():
            col.Visibility = (Windows.Visibility.Visible if not sheets_mode
                              else Windows.Visibility.Collapsed)

    def view_state_changed(self, sender, args):
        """All / Placed / Unplaced dropdown changed."""
        try:
            sel = sender.SelectedItem
            self._state_filter = (None if not sel or sel == 'All views'
                                  else sel)
            self._apply_filter()
        except Exception as ex:
            logger.warning('View state filter failed: %s', ex)

    def _set_views_headers(self):
        """Swap grid headers for views mode."""
        try:
            self.number_col.Header  = getattr(self, '_state_cb', None) or 'View'
            self.number_col.Width   = Windows.Controls.DataGridLength(136)
            self.number_col.Binding = Windows.Data.Binding('placed_on')
            self.name_col.Header    = 'View Name'
        except Exception as ex:
            logger.warning('Views headers failed: %s', ex)

    def _set_sheets_headers(self):
        """Restore grid headers for sheets mode."""
        try:
            self.number_col.Header  = 'Sheet No.'
            self.number_col.Width   = Windows.Controls.DataGridLength(88)
            self.number_col.Binding = Windows.Data.Binding('number')
            self.name_col.Header    = 'Sheet Name'
        except Exception as ex:
            logger.warning('Sheets headers failed: %s', ex)

    def type_filter_changed(self, sender, args):
        """Type dropdown changed — show only the selected view type."""
        try:
            sel = sender.SelectedItem
            self._type_filter = None if (not sel or sel == 'All types') else {sel}
            self._apply_filter()
        except Exception as ex:
            logger.warning('Type filter failed: %s', ex)

    @staticmethod
    def _find_tblock(revit_sheet, tblocks):
        for tb in tblocks:
            if (revit_sheet.Document.GetElement(tb.OwnerViewId).Id
                    == revit_sheet.Id):
                return tb
        return None

    def _build_sheet_ps_map(self, tblocks, all_ps):
        """Return {sheet_number: TitleBlockPrintSettings}."""
        cache = {}
        result = {}
        for tb in tblocks:
            sheet = self._selected_doc.GetElement(tb.OwnerViewId)
            tf    = tb.GetTotalTransform()
            key   = (get_elementid_value(tb.GetTypeId()) * 100
                     + tf.BasisX.X * 10 + tf.BasisX.Y)
            if key not in cache:
                tb_type  = self._selected_doc.GetElement(tb.GetTypeId())
                tps = None
                if tb_type:
                    p = tb_type.LookupParameter('Print Setting')
                    if p:
                        match = next((x for x in all_ps if x.Name == p.AsString()), None)
                        if match:
                            tps = TitleBlockPrintSettings(
                                psettings=[match], set_by_param=True)
                if tps is None:
                    tps = TitleBlockPrintSettings(
                        psettings=revit.query.get_titleblock_print_settings(
                            tb, self._selected_printer, all_ps),
                        set_by_param=False)
                cache[key] = tps
            result[sheet.SheetNumber] = cache[key]
        return result

    def _apply_filter(self):
        """Filter + sort _ordered_sheets, push to DataGrid, restore selections."""
        text = self.search_tb.Text.lower() if self.search_tb.Text else ''
        sheets = self._ordered_sheets

        # Filter by "Appears In Sheet List" — sheets mode only
        try:
            if self._sv_mode == 'sheets' and self.active_only_cb.IsChecked:
                def _appears_in_sheet_list(s):
                    try:
                        p = s.revit_sheet.get_Parameter(
                            DB.BuiltInParameter.SHEET_SCHEDULED)
                        return p is not None and p.AsInteger() == 1
                    except Exception:
                        return True  # include if parameter unreadable
                sheets = [s for s in sheets if _appears_in_sheet_list(s)]
        except Exception:
            pass

        # Filter by search
        if text:
            sheets = [s for s in sheets
                      if text in s.number.lower() or text in s.name.lower()]

        # Filter by view type — views mode only
        if self._sv_mode == 'views' and self._type_filter is not None:
            sheets = [s for s in sheets if s.paper_size in self._type_filter]

        # Schedules belong to XLS alone. Revit cannot print or export one to
        # any other format, so rather than listing them and failing at export
        # - which on a combined PDF took the whole batch down with it - they
        # are simply not shown unless XLS is the format being viewed. The
        # type lock above already narrows XLS to schedules, so this is the
        # other half of the same rule.
        if (self._sv_mode == 'views' and
                self._fmt_viewing not in SCHEDULES_ONLY_FORMATS):
            sheets = [s for s in sheets
                      if s.paper_size != SCHEDULE_TYPE_LABEL]

        # Filter by revision / size / sheet collection — sheets mode only
        if self._sv_mode == 'sheets':
            if self._rev_filter is not None:
                sheets = [s for s in sheets if
                          (s.revision.number if s.revision.is_set else '<None>')
                          in self._rev_filter]
            if self._size_filter is not None:
                sheets = [s for s in sheets if
                          (s.paper_size or '<None>') in self._size_filter]
            if self._coll_filter is not None:
                sheets = [s for s in sheets if
                          (s.sheet_collection or '<None>') in self._coll_filter]

        # Filter by placement state — views mode only
        state = getattr(self, '_state_filter', None)
        if self._sv_mode == 'views' and state:
            if state == 'Placed views':
                sheets = [s for s in sheets if getattr(s, 'placed_on', '')]
            else:
                sheets = [s for s in sheets if not getattr(s, 'placed_on', '')]

        # Filter by custom parameter columns (Manage Columns)
        active_keys = (self._custom_sheet_columns if self._sv_mode == 'sheets'
                       else self._custom_view_columns)
        for key in active_keys:
            fset = self._extra_filters.get(key)
            if fset is not None:
                sheets = [s for s in sheets if
                          (s.custom_params.get(key, '') or '<None>') in fset]

        # Restore checkbox states from THIS FORMAT's selection container
        sel_ids = set(self._selection._ids)
        for s in sheets:
            s._is_selected = get_elementid_value(s.revit_sheet.Id) in sel_ids

        # Regenerate filenames for viewing format
        self._update_filenames(sheets)

        self.sheets_dg.ItemsSource = ObjectModel.ObservableCollection[object](sheets)
        self._update_sel_count()
        self._update_select_all_cb()

    def _update_filenames(self, sheets):
        """Update print_filename on each sheet for the current viewing format."""
        fmt    = self._fmt_viewing
        nfname = self._fmt_naming.get(fmt)
        nfmts  = list(self.namingformat_cb.ItemsSource or [])
        nf     = next((x for x in nfmts if x.name == nfname), None)
        if nf is None and nfmts:
            nf = nfmts[0]
        if nf is None:
            return

        template = self._resolve_template(nf.template)

        for s in sheets:
            self._set_sheet_filename(s, template, fmt)

    def _replace_param(self, template, ptype, getter):
        pattern = r'{' + ptype + r':(.*?)}'
        for name in re.findall(pattern, template):
            val = getter(name) or ''
            template = re.sub(r'{' + ptype + ':' + name + r'}', str(val), template)
        return template

    def _resolve_template(self, template):
        """Replace project/global params — same for every sheet.

        Read from the host model, not the document picked in the dropdown.
        These are project-level values: exporting a link's sheets is still
        this project's issue, so it is this project's number, parameters and
        globals that name the file."""
        doc = self._host_doc
        template = self._replace_param(template, 'proj_param',
            lambda x: revit.query.get_param_value(
                doc.ProjectInformation.LookupParameter(x)))
        template = self._replace_param(template, 'glob_param',
            lambda x: revit.query.get_param_value(
                revit.query.get_global_parameter(x, doc=doc)))
        return template

    def _set_sheet_filename(self, sheet, template, fmt):
        """Generate and set print_filename (with extension) for one sheet."""
        fname = self._make_filename(sheet, template)
        sheet.print_filename = fname + self._ext_for(fmt)

    def _make_filename(self, sheet, template):
        """Return the filename (no extension) for one sheet/view item."""
        t = self._replace_param(template, 'tblock_param',
            lambda x: (revit.query.get_param_value(
                revit.query.get_param(sheet.revit_tblock, x))
                or revit.query.get_param_value(
                revit.query.get_param(sheet.revit_tblock_type, x))))
        t = self._replace_param(t, 'sheet_param',
            lambda x: revit.query.get_param_value(
                revit.query.get_param(sheet.revit_sheet, x)))
        try:
            fname = t.format(
                number         = sheet.number,
                name           = sheet.name,
                name_dash      = sheet.name.replace(' ', '-'),
                name_underline = sheet.name.replace(' ', '_'),
                current_date   = coreutils.current_date(),
                issue_date     = sheet.issue_date,
                rev_number     = sheet.revision.number,
                rev_desc       = sheet.revision.desc,
                rev_date       = sheet.revision.date,
                proj_name      = self._project_info.name,
                proj_number    = self._project_info.number,
                proj_building_name = self._project_info.building_name,
                proj_issue_date    = self._project_info.issue_date,
                proj_org_name      = self._project_info.org_name,
                proj_status        = self._project_info.status,
                username       = HOST_APP.username,
                revit_version  = HOST_APP.version,
            )
        except Exception:
            fname = ((sheet.number + ' ') if sheet.number else '') + sheet.name
        return coreutils.cleanup_filename(fname, windows_safe=True)

    def _ext_for(self, fmt):
        """File extension for a format — IMG and XLS depend on their type
        dropdown, the way IMG covers png/jpeg/tiff and XLS covers xlsx/ods."""
        if fmt == 'IMG':
            try:
                itype = self.img_type_cb.SelectedItem.Content
                return IMG_TYPE_EXT.get(itype, '.png')
            except Exception:
                return '.png'
        if fmt == 'XLS':
            return self._xls_ext()
        return FILE_EXT.get(fmt, '')

    XLS_TYPE_EXT = {'XLSX': '.xlsx', 'ODS': '.ods', 'CSV': '.csv'}

    def _xls_ext(self):
        """'.xlsx', '.ods' or '.csv', from the Table options dropdown."""
        try:
            return self.XLS_TYPE_EXT.get(
                self.xls_type_cb.SelectedItem.Content, '.xlsx')
        except Exception:
            return '.xlsx'

    # ── FORMAT PILL LOGIC ──
    def _fmt_btn(self, fmt):
        """Return the Button element for a format string."""
        name_map = {
            'PDF': self.fmt_pdf_btn,
            'DWG': self.fmt_dwg_btn,
            'DGN': self.fmt_dgn_btn,
            'NWC': self.fmt_nwc_btn,
            'IFC': self.fmt_ifc_btn,
            'IMG': self.fmt_img_btn,
            'XLS': self.fmt_xls_btn,
        }
        return name_map.get(fmt)

    def _exp_btn(self, fmt):
        """Return the Settings-tab format pill Button for a format string."""
        name_map = {
            'PDF': self.exp_pdf_btn,
            'DWG': self.exp_dwg_btn,
            'DGN': self.exp_dgn_btn,
            'NWC': self.exp_nwc_btn,
            'IFC': self.exp_ifc_btn,
            'IMG': self.exp_img_btn,
            'XLS': self.exp_xls_btn,
        }
        return name_map.get(fmt)

    @staticmethod
    def _pill_tooltip(tag):
        """Tooltip text matching a pill's on/off/active tri-state."""
        if tag == 'Viewing':
            return 'Deactivate'
        if tag == 'Enabled':
            return 'Select'
        return 'Activate'

    def _update_format_buttons(self):
        """Sync all format pill Tags to current _fmt_enabled / _fmt_viewing —
        both the Select-tab pills and the Settings-tab sub-tab pills share
        the same viewing format."""
        for fmt in ALL_FORMATS:
            btn = self._fmt_btn(fmt)
            if btn is not None:
                if fmt == self._fmt_viewing and fmt in self._fmt_enabled:
                    btn.Tag = 'Viewing'
                elif fmt in self._fmt_enabled:
                    btn.Tag = 'Enabled'
                else:
                    btn.Tag = ''
                btn.ToolTip = self._pill_tooltip(btn.Tag)

            exp_btn = self._exp_btn(fmt)
            if exp_btn is not None:
                enabled = fmt in self._fmt_enabled
                exp_btn.IsEnabled = enabled
                if fmt == self._fmt_viewing and enabled:
                    exp_btn.Tag = 'Viewing'
                elif enabled:
                    exp_btn.Tag = 'Enabled'
                else:
                    exp_btn.Tag = ''
                exp_btn.ToolTip = self._pill_tooltip(exp_btn.Tag)

    def _update_fmt_column_header(self):
        """Update the green column label text to show current viewing format."""
        try:
            lbl = self.fmt_col_label
            lbl.Text = self._fmt_viewing
        except Exception:
            pass
        # Sync naming format dropdown to this format's saved selection
        fmts    = list(self.namingformat_cb.ItemsSource or [])
        nfname  = self._fmt_naming.get(self._fmt_viewing)
        match   = next((x for x in fmts if x.name == nfname), None)
        if match:
            self.namingformat_cb.SelectedItem = match
        elif fmts:
            self.namingformat_cb.SelectedIndex = 0

    def _fmt_pill_logic(self, fmt):
        """
        Click logic for a format pill:
          - If clicking the currently-viewing format → toggle enabled off/on
          - If clicking any other format → enable it + switch view to it
          Each format has its own independent selection container.
        """
        if fmt == self._fmt_viewing:
            if fmt in self._fmt_enabled:
                self._fmt_enabled.discard(fmt)
                # Auto-jump view to another enabled format if any
                others = [f for f in ALL_FORMATS
                          if f in self._fmt_enabled and f != fmt]
                if others:
                    self._switch_viewing_format(others[0])
            else:
                self._fmt_enabled.add(fmt)
                self._switch_viewing_format(fmt)
        else:
            self._fmt_enabled.add(fmt)
            self._switch_viewing_format(fmt)

        self._update_format_buttons()
        self._update_fmt_column_header()
        self._apply_filter()

    def _switch_viewing_format(self, fmt):
        """Switch which format's column is shown and swap selection container.
        Keeps the Settings-tab sub-panel in sync too, since both tabs share
        one 'viewing' format. NWC/IFC only ever export 3D views, so they
        force Views mode and lock the Type filter to 3D."""
        self._fmt_viewing = fmt
        self._show_exp_panel(fmt)
        # XLS is views-only for the same reason NWC/IFC are: a schedule is
        # a view, and there is no sheet it could ever apply to.
        if fmt in VIEWS_ONLY_FORMATS or fmt in SCHEDULES_ONLY_FORMATS:
            self.sv_sheets_btn.IsEnabled = False
            self.sv_sheets_btn.Opacity   = 0.35
            if self._sv_mode != 'views':
                self.sv_views_clicked(None, None)
        else:
            self.sv_sheets_btn.IsEnabled = True
            self.sv_sheets_btn.Opacity   = 1.0
        # Point _selection to the right container for current mode + format
        containers = self._view_selections if self._sv_mode == 'views' else self._sheet_selections
        self._selection = containers[fmt]
        # Re-attach the new container to all existing items and restore their state
        for s in self._all_sheets:
            s._container = self._selection
            sheet_id = get_elementid_value(s.revit_sheet.Id)
            s._is_selected = self._selection.contains(sheet_id)
        if self._sv_mode == 'views':
            self._apply_type_lock()

    def _type_labels(self):
        """Entries for the Type header dropdown, for the format being viewed.

        Schedules belong to XLS and nowhere else, so 'Schedule' is not
        offered under any other format - and under XLS it is the only entry.
        Leaving it in every list meant you could pick a type that the format
        filters out anyway, and switching away from XLS left the dropdown
        holding 'Schedule' against a list that could never contain one."""
        types = set(v.paper_size for v in self._all_view_items)
        if self._fmt_viewing in SCHEDULES_ONLY_FORMATS:
            return ([SCHEDULE_TYPE_LABEL]
                    if SCHEDULE_TYPE_LABEL in types else [])
        return ['All types'] + sorted(t for t in types
                                      if t != SCHEDULE_TYPE_LABEL)

    def _apply_type_lock(self):
        """Lock the Type filter dropdown: 3D when viewing NWC/IFC, Schedule
        when viewing XLS. Both formats accept exactly one kind of view, so
        the filter is fixed rather than left to be got wrong."""
        cb = getattr(self, '_type_cb', None)
        if self._fmt_viewing in SCHEDULES_ONLY_FORMATS:
            locked_type = SCHEDULE_TYPE_LABEL
        elif self._fmt_viewing in VIEWS_ONLY_FORMATS:
            locked_type = '3D'
        else:
            locked_type = None
        lock = locked_type is not None
        if cb is not None:
            try:
                # Rebuilt on every format switch, not just when the view list
                # reloads - which format is showing decides what belongs here.
                cb.ItemsSource = self._type_labels()
                if lock:
                    cb.SelectedItem = locked_type
                else:
                    cb.SelectedIndex = 0
                cb.IsEnabled = not lock
            except Exception:
                pass
        # Last, deliberately: assigning ItemsSource and the selection above
        # fires type_filter_changed, which would otherwise overwrite this
        # with whatever the combo happened to land on mid-rebuild.
        self._type_filter = {locked_type} if lock else None

    # ── TAB NAVIGATION ──
    def _show_tab(self, tab_index):
        self._current_tab = tab_index

        # Visibility
        self.tab_select.Visibility = (Windows.Visibility.Visible
                                      if tab_index == TAB_SELECT
                                      else Windows.Visibility.Collapsed)
        self.tab_export.Visibility = (Windows.Visibility.Visible
                                      if tab_index == TAB_EXPORT
                                      else Windows.Visibility.Collapsed)
        self.tab_print.Visibility  = (Windows.Visibility.Visible
                                      if tab_index == TAB_PRINT
                                      else Windows.Visibility.Collapsed)

        # Next / Print icon button — always visible; icon changes with the tab
        try:
            on_export = (tab_index == TAB_PRINT)
            self.header_printer_icon.Visibility = (
                Windows.Visibility.Visible if on_export
                else Windows.Visibility.Collapsed)
            self.header_chevrons_icon.Visibility = (
                Windows.Visibility.Collapsed if on_export
                else Windows.Visibility.Visible)
            self.header_print_btn.ToolTip = 'Print / Export' if on_export else 'Next'
        except Exception:
            pass

        # Tab button Tags
        tab_states = ['', '', '']
        tab_states[tab_index] = 'Active'
        for i in range(tab_index):
            tab_states[i] = 'Done'

        self.tab_select_btn.Tag = tab_states[TAB_SELECT]
        self.tab_export_btn.Tag = tab_states[TAB_EXPORT]
        self.tab_print_btn.Tag  = tab_states[TAB_PRINT]

        # Step indicators
        self._update_step_indicators(tab_index)

        # Header subtitle — matches new tab names
        subtitles = ['| Select', '| Settings', '| Export']
        self.header_subtitle_tb.Text = subtitles[tab_index]

    def _update_step_indicators(self, current):
        """Accent fill for done/active, muted border for future steps.

        Resources are looked up fresh on every call (rather than cached
        hex constants) so this always reflects whatever accent colour is
        current in the palette - see seed43-pyrevit-ui skill gotcha #3
        (Python-built UI colours are a one-time snapshot unless re-looked
        up on every call that can run after a theme/accent change)."""
        accent  = self.TryFindResource('BrushPrimaryGreen') or _brush('#208A3C')
        grey_bg = self.TryFindResource('BrushSecondaryDisabledBg') or _brush('#555555')
        grey_fg = self.TryFindResource('BrushSecondaryDisabledFg') or _brush('#888888')
        white   = _brush('White')

        circles  = [self.step1_circle, self.step2_circle, self.step3_circle]
        lines    = [self.step_line_1,  self.step_line_2]

        for i, circle in enumerate(circles):
            if i < current:       # done
                circle.Background      = accent
                circle.BorderBrush     = accent
                circle.BorderThickness = Windows.Thickness(0)
                _set_child_text(circle, '✓', white)
            elif i == current:    # active
                circle.Background      = accent
                circle.BorderBrush     = accent
                circle.BorderThickness = Windows.Thickness(0)
                _set_child_text(circle, str(i + 1), white)
            else:                 # future
                circle.Background      = _brush('Transparent')
                circle.BorderBrush     = grey_bg
                circle.BorderThickness = Windows.Thickness(2)
                _set_child_text(circle, str(i + 1), grey_fg)

        for i, line in enumerate(lines):
            line.Fill = accent if i < current else grey_bg

    # ── SELECTION HELPERS ──
    def _update_sel_count(self):
        visible = self._visible_sheets
        total   = len(visible)
        checked = sum(1 for s in visible if s.is_selected)
        self.sel_count_tb.Text = '{} of {} selected'.format(checked, total)

        fmt = self._fmt_viewing if self._fmt_viewing in self._fmt_enabled else None
        if fmt:
            n_sheets = self._sheet_selections[fmt].count()
            n_views  = self._view_selections[fmt].count()
            self.footer_status_tb.Text = (
                '{} {} sheet{} | {} View{} selected'
                .format(fmt, n_sheets, 's' if n_sheets != 1 else '',
                        n_views, 's' if n_views != 1 else ''))
        else:
            self.footer_status_tb.Text = 'No format selected'

    def _update_select_all_cb(self):
        sheets = self._visible_sheets
        if not sheets:
            self.select_all_cb.IsChecked = False
            return
        n_sel = sum(1 for s in sheets if s.is_selected)
        if n_sel == 0:
            self.select_all_cb.IsChecked = False
        elif n_sel == len(sheets):
            self.select_all_cb.IsChecked = True
        else:
            self.select_all_cb.IsChecked = None  # indeterminate

    # ── PRINT TAB QUEUE ──
    def _get_checked_for_format(self, fmt):
        """
        Return all selected items (sheets + views) for a given format.
        Reads from the persistent selection containers so selections are
        honoured regardless of which mode (sheets/views) is currently showing.
        """
        sheet_container = self._sheet_selections[fmt]
        view_container  = self._view_selections[fmt]

        checked = []
        for item in self._all_sheet_items:
            eid = get_elementid_value(item.revit_sheet.Id)
            if sheet_container.contains(eid):
                checked.append(item)
        for item in self._all_view_items:
            eid = get_elementid_value(item.revit_sheet.Id)
            if view_container.contains(eid):
                checked.append(item)
        return checked

    def _build_queue(self):
        """Populate Print tab queue DataGrid with selected items × enabled formats."""
        items = []
        fmts_list = list(self.namingformat_cb.ItemsSource or [])

        for fmt in ALL_FORMATS:
            if fmt not in self._fmt_enabled:
                continue
            checked = self._get_checked_for_format(fmt)
            nf_name = self._fmt_naming.get(fmt)
            nobj    = next((x for x in fmts_list if x.name == nf_name), None)
            if nobj is None and fmts_list:
                nobj = fmts_list[0]
            tmpl = self._resolve_template(
                nobj.template if nobj else '{number} {name}')

            for s in checked:
                fname = self._make_filename(s, tmpl)
                size_display = s.paper_size
                if fmt == 'PDF':
                    chosen = (self.papersize_cb.Text or 'From Titleblock').strip()
                    if chosen and chosen != 'From Titleblock':
                        size_display = chosen
                items.append(QueueItem(
                    s.number, fname + self._ext_for(fmt), fmt,
                    size_display, source=s, fname_noext=fname))

        self.queue_dg.ItemsSource = ObjectModel.ObservableCollection[object](items)
        self.overall_progress.Value = 0
        self.overall_pct_tb.Text = 'Ready'
    # ── EXPORT ENGINE ──
    def _stop(self, message):
        """Refuse to export. Unattended runs raise instead of showing a
        dialog - a modal box on a scheduled run would sit there blocking
        Revit until someone came back and clicked it."""
        if self._unattended:
            raise ExportBlocked(message)
        dlg.message(message)

    def _do_export(self):
        """Run the export queue across all enabled formats."""
        if not self._fmt_enabled:
            self._stop('No export formats selected.')
            return
        if self._get_print_destination() == 'none':
            self._stop('Please select a print destination (file or printer).')
            return
        base_folder = self.export_folder_tb.Text
        if not op.isdir(base_folder):
            self._stop('Export folder does not exist.')
            return

        queue = list(self.queue_dg.ItemsSource or [])
        if not queue:
            self._build_queue()
            queue = list(self.queue_dg.ItemsSource or [])
        if not queue:
            self._stop('No sheets or views selected.')
            return

        if not self.auto_overwrite_cb.IsChecked:
            existing = self._existing_output_files(queue, base_folder)
            if existing:
                shown = existing[:10]
                more  = len(existing) - len(shown)
                msg = '{} file(s) already exist and will be overwritten:\n\n{}'.format(
                    len(existing), '\n'.join(shown))
                if more > 0:
                    msg += '\n... and {} more'.format(more)
                if self._unattended:
                    # Overwrite rather than stall on the prompt, but say so.
                    logger.warning('Scheduled export overwriting %d existing '
                                   'file(s) in %s', len(existing), base_folder)
                elif not dlg.confirm(msg, yes='Overwrite'):
                    return

        total = len(queue)
        done  = [0]

        def tick(qi, status):
            qi.status   = status
            qi.progress = 100 if status in ('Done', 'Failed', 'Skipped') else 50
            if status in ('Done', 'Failed', 'Skipped'):
                done[0] += 1
                pct = int(done[0] * 100.0 / total)
                self.overall_progress.Value = pct
                self.overall_pct_tb.Text = '{}%'.format(pct)
            self._pump()

        by_fmt = {}
        for qi in queue:
            by_fmt.setdefault(qi.format, []).append(qi)



        exporters = {'PDF': self._export_pdf, 'DWG': self._export_dwg,
                     'DGN': self._export_dgn, 'NWC': self._export_nwc,
                     'IFC': self._export_ifc, 'IMG': self._export_img,
                     'XLS': self._export_xls}

        for fmt in ALL_FORMATS:
            qitems = by_fmt.get(fmt)
            if not qitems:
                continue
            folder = self._fmt_folder(base_folder, fmt)
            try:
                exporters[fmt](qitems, folder, tick)
            except Exception as ex:
                logger.error('%s export failed: %s', fmt, ex)
                for qi in qitems:
                    if qi.status in ('Waiting', 'Exporting'):
                        tick(qi, 'Failed')

        self.overall_pct_tb.Text = 'Complete'
        if self.open_folder_cb.IsChecked:
            # Not coreutils.open_folder_in_explorer: that launches Explorer
            # unconditionally, so every export left another window for the
            # same folder stacked on the last one.
            try:
                open_folder(base_folder)
            except Exception:
                pass

    def _fmt_folder(self, base, fmt):
        """Return (and create) the per-format subfolder if enabled."""
        if self.subfolder_cb.IsChecked:
            p = op.join(base, fmt)
            if not op.isdir(p):
                os.makedirs(p)
            return p
        return base

    def _existing_output_files(self, queue, base_folder):
        """File-destination output paths that already exist on disk.
        Read-only — unlike _fmt_folder, this never creates folders, since
        it just powers the pre-export overwrite prompt."""
        if self._get_print_destination() not in ('file', 'both'):
            return []
        by_fmt = {}
        for qi in queue:
            by_fmt.setdefault(qi.format, []).append(qi)
        found = []
        for fmt, qitems in by_fmt.items():
            folder = (op.join(base_folder, fmt) if self.subfolder_cb.IsChecked
                      else base_folder)
            if (fmt == 'PDF' and bool(self.file_combine_rb.IsChecked)
                    and len(qitems) > 1):
                path = op.join(folder, self._combined_pdf_name(qitems) + '.pdf')
                if op.isfile(path):
                    found.append(path)
            else:
                for qi in qitems:
                    path = op.join(folder, qi.filename)
                    if op.isfile(path):
                        found.append(path)
        return found

    @staticmethod
    def _pump():
        """Let WPF repaint mid-export so queue statuses update live."""
        try:
            Forms.Application.DoEvents()
        except Exception:
            pass

    # ── PDF (native Revit exporter, Revit 2022+) ──
    def _pdf_options(self):
        """Build native PDFExportOptions from the Settings tab controls."""
        o = DB.PDFExportOptions()
        try:
            # Was hardcoded to Default (= always size-to-sheet), so
            # papersize_cb's selection never reached the native exporter —
            # this is the primary PDF path, the PrintManager one is a
            # fallback, so this alone explains every sheet printing at its
            # own titleblock size regardless of the Size dropdown.
            paper_name = (self.papersize_cb.Text or 'From Titleblock').strip()
            paper_format_map = {
                'From Titleblock': DB.ExportPaperFormat.Default,
                'A0': DB.ExportPaperFormat.ISO_A0,
                'A1': DB.ExportPaperFormat.ISO_A1,
                'A2': DB.ExportPaperFormat.ISO_A2,
                'A3': DB.ExportPaperFormat.ISO_A3,
                'A4': DB.ExportPaperFormat.ISO_A4,
            }
            o.PaperFormat = paper_format_map.get(paper_name, DB.ExportPaperFormat.Default)
        except Exception as ex:
            logger.warning('Could not set PDF paper format "%s": %s', paper_name, ex)
            try:
                o.PaperFormat = DB.ExportPaperFormat.Default
            except Exception:
                pass
        try:
            if self.orient_portrait_btn.Tag == 'Viewing':
                o.PaperOrientation = DB.PageOrientationType.Portrait
            elif self.orient_landscape_btn.Tag == 'Viewing':
                o.PaperOrientation = DB.PageOrientationType.Landscape
            else:
                o.PaperOrientation = DB.PageOrientationType.Auto
        except Exception:
            pass
        try:
            if self.zoom_fit_rb.IsChecked:
                o.ZoomType = DB.ZoomType.FitToPage
            else:
                o.ZoomType = DB.ZoomType.Zoom
                o.ZoomPercentage = int(float(self.zoom_pct_tb.Text or 100))
        except Exception:
            pass
        try:
            if self.placement_center_rb.IsChecked:
                o.PaperPlacement = DB.PaperPlacementType.Center
            else:
                o.PaperPlacement = DB.PaperPlacementType.LowerLeft
                # UI offsets are millimetres — API wants feet
                o.OriginOffsetX = float(self.offset_x_tb.Text or 0) / 304.8
                o.OriginOffsetY = float(self.offset_y_tb.Text or 0) / 304.8
        except Exception:
            pass
        try:
            o.AlwaysUseRaster = bool(self.hlv_raster_rb.IsChecked)
        except Exception:
            pass
        try:
            rq_map = {'High':   DB.RasterQualityType.High,
                      'Medium': DB.RasterQualityType.Medium,
                      'Low':    DB.RasterQualityType.Low}
            rq = (self.raster_quality_cb.SelectedItem.Content
                  if self.raster_quality_cb.SelectedItem else 'High')
            o.RasterQuality = rq_map.get(rq, DB.RasterQualityType.High)
        except Exception:
            pass
        try:
            col_map = {'Color':           DB.ColorDepthType.Color,
                       'Black and White': DB.ColorDepthType.BlackLine,
                       'Grayscale':       DB.ColorDepthType.GrayScale}
            col = (self.colors_cb.SelectedItem.Content
                   if self.colors_cb.SelectedItem else 'Color')
            o.ColorDepth = col_map.get(col, DB.ColorDepthType.Color)
        except Exception:
            pass
        try:
            o.ViewLinksInBlue           = bool(self.opt_links_blue_cb.IsChecked)
            o.HideReferencePlane        = bool(self.opt_hide_refplanes_cb.IsChecked)
            o.HideUnreferencedViewTags  = bool(self.opt_hide_unreftags_cb.IsChecked)
            o.HideScopeBoxes            = bool(self.opt_hide_scopeboxes_cb.IsChecked)
            o.HideCropBoundaries        = bool(self.opt_hide_cropbounds_cb.IsChecked)
            o.ReplaceHalftoneWithThinLines = bool(self.opt_halftone_thin_cb.IsChecked)
            o.MaskCoincidentLines       = bool(self.opt_region_edges_cb.IsChecked)
        except Exception:
            pass
        try:
            o.StopOnError = False
        except Exception:
            pass
        return o

    def _export_pdf(self, qitems, folder, tick):
        doc  = self._selected_doc
        dest = self._get_print_destination()

        if dest in ('file', 'both'):
            opts    = self._pdf_options()
            combine = bool(self.file_combine_rb.IsChecked)
            if combine and len(qitems) > 1:
                for qi in qitems:
                    qi.status = 'Exporting'
                self._pump()
                ids   = List[DB.ElementId](
                    [qi.source.revit_sheet.Id for qi in qitems])
                opts.Combine  = True
                opts.FileName = self._combined_pdf_name(qitems)
                try:
                    ok = doc.Export(folder, ids, opts)
                    for qi in qitems:
                        tick(qi, 'Done' if ok else 'Failed')
                except Exception as ex:
                    logger.error('Combined PDF export failed: %s', ex)
                    for qi in qitems:
                        tick(qi, 'Failed')
            else:
                for qi in qitems:
                    qi.status = 'Exporting'
                    self._pump()
                    try:
                        opts.Combine  = True   # one file per call
                        opts.FileName = qi.fname_noext
                        ok = doc.Export(
                            folder,
                            List[DB.ElementId]([qi.source.revit_sheet.Id]),
                            opts)
                        tick(qi, 'Done' if ok else 'Failed')
                    except Exception as ex:
                        logger.error('PDF %s failed: %s', qi.filename, ex)
                        tick(qi, 'Failed')

        if dest in ('printer', 'both'):
            self._print_to_physical([qi.source for qi in qitems])
            if dest == 'printer':
                for qi in qitems:
                    tick(qi, 'Done')

    def _combined_pdf_name(self, qitems):
        """Name for the single combined PDF from the chosen naming format."""
        try:
            nf = self.combined_naming_cb.SelectedItem
            if nf:
                tmpl  = self._resolve_template(nf.template)
                first = qitems[0].source
                return self._make_filename(first, tmpl)
        except Exception as ex:
            logger.warning('Combined name failed: %s', ex)
        now   = datetime.now().strftime('%Y-%m-%d %H.%M')
        pnum  = self._project_info.number or 'NoNumber'
        pname = self._project_info.name or 'NoName'
        return coreutils.cleanup_filename(
            '{} - {} {}'.format(pnum, pname, now), windows_safe=True)

    def combined_naming_changed(self, sender, args):
        if self._loading:
            return
        nf = self.combined_naming_cb.SelectedItem
        if nf:
            self._fmt_naming['PDF_COMBINED'] = nf.name
            self._save_naming_memory()

    def _print_to_physical(self, sheets):
        """Send items to the selected physical printer (PrintManager path).

        The inner blocks each catch their own failures, but the transactions
        themselves have to be guarded too: opening one can throw before the
        body ever runs (a document that refuses transactions is the case that
        turns up), and an exception escaping a handler unwinds out of
        ShowDialog with a transaction open - which is how this takes Revit
        down rather than just the window."""
        pm      = self._get_printmanager()
        printer = self._selected_printer
        if not pm or not printer:
            dlg.message('No printer selected for physical printing.')
            return
        try:
            self._print_to_physical_inner(pm, printer, sheets)
        except Exception as ex:
            logger.error('Physical print failed: %s', ex)
            dlg.message('Could not send to the printer.\n\n' + str(ex))

    def _print_to_physical_inner(self, pm, printer, sheets):
        with revit.DryTransaction('Apply Print Settings',
                                  doc=self._selected_doc):
            try:
                ps_item = self._selected_printsetting
                if ps_item is not None and not ps_item.allows_variable_paper:
                    pm.PrintSetup.CurrentPrintSetting = ps_item.print_settings
                self._read_ui_print_settings(pm)
            except Exception as ex:
                logger.warning('Could not set print settings: %s', ex)
        with revit.DryTransaction('Send To Printer',
                                  doc=self._selected_doc):
            try:
                pm.PrintRange = DB.PrintRange.Current
                pm.SelectNewPrintDriver(printer)
                try:
                    pm.PrintToFile = False
                except Exception:
                    # Virtual printers (PDF24 etc.) must print to file —
                    # the driver will pop its own save dialog
                    try:
                        pm.PrintToFile = True
                    except Exception:
                        pass
                ps_item = self._selected_printsetting
                for s in sheets:
                    if (ps_item is not None and ps_item.allows_variable_paper
                            and s.print_settings):
                        pm.PrintSetup.CurrentPrintSetting = s.print_settings
                    try:
                        pm.SubmitPrint(s.revit_sheet)
                    except Exception:
                        logger.error('Failed to print %s', s.number)
            except Exception as ex:
                dlg.message('Printer setup failed.\n\n' + str(ex))

    # ── DWG ──
    def _export_dwg(self, qitems, folder, tick):
        doc  = self._selected_doc
        opts = None
        try:
            name = self.dwg_setup_cb.SelectedItem
            if name and name != DEFAULT_SETUP:
                for s in (DB.FilteredElementCollector(doc)
                          .OfClass(framework.get_type(DB.ExportDWGSettings))):
                    if s.Name == name:
                        opts = s.GetDWGExportOptions()
                        break
        except Exception:
            pass
        if opts is None:
            opts = DB.DWGExportOptions()
        try:
            # Unchecked = merge views into the sheet DWG (no xrefs)
            opts.MergedViews = not bool(self.dwg_xrefs_cb.IsChecked)
        except Exception:
            pass
        for qi in qitems:
            qi.status = 'Exporting'
            self._pump()
            try:
                ok = doc.Export(
                    folder, qi.fname_noext,
                    List[DB.ElementId]([qi.source.revit_sheet.Id]), opts)
                tick(qi, 'Done' if ok else 'Failed')
            except Exception as ex:
                logger.error('DWG %s failed: %s', qi.filename, ex)
                tick(qi, 'Failed')

    # ── DGN ──
    def _export_dgn(self, qitems, folder, tick):
        doc  = self._selected_doc
        opts = None
        try:
            name = self.dgn_setup_cb.SelectedItem
            if name and name != DEFAULT_SETUP:
                for s in (DB.FilteredElementCollector(doc)
                          .OfClass(framework.get_type(DB.ExportDGNSettings))):
                    if s.Name == name:
                        opts = s.GetDGNExportOptions()
                        break
        except Exception:
            pass
        if opts is None:
            opts = DB.DGNExportOptions()
        for qi in qitems:
            qi.status = 'Exporting'
            self._pump()
            try:
                ok = doc.Export(
                    folder, qi.fname_noext,
                    List[DB.ElementId]([qi.source.revit_sheet.Id]), opts)
                tick(qi, 'Done' if ok else 'Failed')
            except Exception as ex:
                logger.error('DGN %s failed: %s', qi.filename, ex)
                tick(qi, 'Failed')

    # ── NWC ──
    def _export_nwc(self, qitems, folder, tick):
        doc = self._selected_doc
        if not self._check_nwc_available():
            dlg.message('Navisworks exporter add-in is not installed.\n'
                        'NWC items will be skipped.')
            for qi in qitems:
                tick(qi, 'Skipped')
            return
        opts = DB.NavisworksExportOptions()
        opts.ExportScope = DB.NavisworksExportScope.View
        try:
            opts.Coordinates = (
                DB.NavisworksCoordinates.Shared
                if self.nwc_shared_coords_cb.IsChecked
                else DB.NavisworksCoordinates.Internal)
        except Exception:
            pass
        try:
            opts.ConvertElementProperties = bool(
                self.nwc_props_cb.IsChecked)
        except Exception:
            pass
        for qi in qitems:
            qi.status = 'Exporting'
            self._pump()
            try:
                opts.ViewId = qi.source.revit_sheet.Id
                doc.Export(folder, qi.fname_noext, opts)
                tick(qi, 'Done')
            except Exception as ex:
                logger.error('NWC %s failed: %s', qi.filename, ex)
                tick(qi, 'Failed')

    # ── IFC ──
    def _export_ifc(self, qitems, folder, tick):
        doc  = self._selected_doc
        opts = DB.IFCExportOptions()
        vmap = {'IFC 2x3 Coordination View 2.0': 'IFC2x3CV2',
                'IFC 2x3':                       'IFC2x3',
                'IFC 4 Reference View':          'IFC4RV',
                'IFC 4 Design Transfer View':    'IFC4DTV'}
        try:
            sel = (self.ifc_version_cb.SelectedItem.Content
                   if self.ifc_version_cb.SelectedItem else '')
            opts.FileVersion = getattr(
                DB.IFCVersion, vmap.get(sel, 'IFC2x3CV2'))
        except Exception:
            pass
        for qi in qitems:
            qi.status = 'Exporting'
            self._pump()
            try:
                opts.FilterViewId = qi.source.revit_sheet.Id
                with revit.Transaction('pySheets IFC Export',
                                       doc=doc):
                    doc.Export(folder, qi.fname_noext, opts)
                tick(qi, 'Done')
            except Exception as ex:
                logger.error('IFC %s failed: %s', qi.filename, ex)
                tick(qi, 'Failed')

    # ── IMG ──
    def _export_img(self, qitems, folder, tick):
        doc   = self._selected_doc
        ftmap = {'PNG':  DB.ImageFileType.PNG,
                 'JPEG': DB.ImageFileType.JPEGLossless,
                 'TIFF': DB.ImageFileType.TIFF}
        rmap  = {'72 DPI':  DB.ImageResolution.DPI_72,
                 '150 DPI': DB.ImageResolution.DPI_150,
                 '300 DPI': DB.ImageResolution.DPI_300,
                 '600 DPI': DB.ImageResolution.DPI_600}
        try:
            itype = (self.img_type_cb.SelectedItem.Content
                     if self.img_type_cb.SelectedItem else 'PNG')
        except Exception:
            itype = 'PNG'
        try:
            ires = (self.img_res_cb.SelectedItem.Content
                    if self.img_res_cb.SelectedItem else '150 DPI')
        except Exception:
            ires = '150 DPI'
        try:
            pixels = int(self.img_pixels_tb.Text or 2048)
        except Exception:
            pixels = 2048

        for qi in qitems:
            # Revit cannot rasterise a schedule. ExportImage rejects the call
            # outright with "some of the views is not exportable", so a
            # Revision Schedule in the queue failed as an error rather than
            # being recognised as something that was never exportable. Same
            # treatment as a missing Navisworks exporter: Skipped, not Failed.
            if isinstance(qi.source.revit_sheet, DB.ViewSchedule):
                logger.info('IMG %s skipped: Revit cannot export a schedule '
                            'as an image', qi.filename)
                tick(qi, 'Skipped')
                continue
            qi.status = 'Exporting'
            self._pump()
            try:
                o = DB.ImageExportOptions()
                o.ExportRange = DB.ExportRange.SetOfViews
                o.SetViewsAndSheets(
                    List[DB.ElementId]([qi.source.revit_sheet.Id]))
                o.FilePath = op.join(folder, qi.fname_noext)
                o.HLRandWFViewsFileType = ftmap.get(
                    itype, DB.ImageFileType.PNG)
                o.ShadowViewsFileType = ftmap.get(
                    itype, DB.ImageFileType.PNG)
                o.ImageResolution = rmap.get(
                    ires, DB.ImageResolution.DPI_150)
                o.ZoomType      = DB.ZoomFitType.FitToPage
                o.PixelSize     = pixels
                o.FitDirection  = DB.FitDirectionType.Horizontal
                doc.ExportImage(o)
                tick(qi, 'Done')
            except Exception as ex:
                logger.error('IMG %s failed: %s', qi.filename, ex)
                tick(qi, 'Failed')

    # ── XLS  (schedules → one workbook, a tab each) ──
    @staticmethod
    def _decode_export(data):
        """Decode whatever ViewSchedule.Export just wrote.

        Revit is not consistent about the encoding, and Python's 'utf-16'
        codec refuses outright when there is no BOM to tell it the byte
        order - "UTF-16 stream does not start with BOM", which is exactly
        what a schedule dump hit. So: trust a BOM if there is one, otherwise
        recognise BOM-less UTF-16LE by its NUL padding, and only then fall
        back through the single-byte encodings. The last step cannot raise,
        because failing to decode must never be the reason an export dies."""
        if data.startswith(b'\xff\xfe') or data.startswith(b'\xfe\xff'):
            return data.decode('utf-16')
        if data.startswith(b'\xef\xbb\xbf'):
            return data.decode('utf-8-sig')
        # A UTF-16LE dump of mostly-ASCII text is close to half NUL bytes;
        # none of the single-byte encodings below can look like that.
        if data.count(b'\x00') > len(data) / 4:
            return data.decode('utf-16-le', 'replace')
        for enc in ('utf-8', 'mbcs', 'latin-1'):
            try:
                return data.decode(enc)
            except Exception:
                continue
        return data.decode('latin-1', 'replace')

    def _read_schedule_rows(self, schedule, tmp_dir):
        """One schedule as (header list, list of row lists).

        Goes through Revit's own ViewSchedule.Export - the same thing the
        Export Schedule command uses - rather than walking table cells. It
        already resolves formatting, units, grouping and totals exactly as
        the schedule shows them; re-deriving that from the API would be a
        different-looking table for no gain."""
        name = coreutils.cleanup_filename(schedule.Name, windows_safe=True)
        tmp_name = name + '.txt'
        opts = DB.ViewScheduleExportOptions()
        try:
            # The whole schedule, as the schedule shows it - group headers,
            # footers and blank rows included. No options: what you get is
            # the schedule.
            opts.HeadersFootersBlanks = True
            # Except the title. Not a preference: the first line of the
            # export is read as the column headers, and Revit puts the title
            # there when this is on - one cell, so every column would be
            # named after the schedule and every row cut to its first value.
            # The tab is already named after the schedule anyway.
            opts.Title = False
        except Exception:
            pass
        try:
            opts.FieldDelimiter = '\t'
            # Revit quotes every field by default. Nothing downstream needs
            # the quoting - the workbook writer escapes its own XML - and
            # stripping it here beats unpicking it per cell later.
            # 'None' is a Python keyword, so this enum member cannot be
            # reached with normal attribute access.
            opts.TextQualifier = getattr(DB.ExportTextQualifier, 'None')
        except Exception:
            pass
        schedule.Export(tmp_dir, tmp_name, opts)

        path = op.join(tmp_dir, tmp_name)
        with open(path, 'rb') as f:
            raw = self._decode_export(f.read())
        lines = [l for l in raw.replace('\r\n', '\n').split('\n') if l.strip()]
        table = [l.split('\t') for l in lines]
        if not table:
            return [], []
        return table[0], table[1:]

    # ── Reading a schedule as a styled grid ──
    # The text export above carries no formatting at all, so anything that
    # has to look like the schedule reads the table itself instead. Only CSV
    # still uses the text path, having nowhere to put a colour.
    @staticmethod
    def _revit_color_hex(color):
        """A Revit Color as 'RRGGBB', or None if it isn't actually set.

        An unset colour comes back as an object that either reports
        IsValid False or answers 0,0,0 - and a real black would be
        indistinguishable, so InvalidColorValue is checked first where the
        API offers it."""
        if color is None:
            return None
        try:
            if hasattr(color, 'IsValid') and not color.IsValid:
                return None
            return u'{0:02X}{1:02X}{2:02X}'.format(
                color.Red, color.Green, color.Blue)
        except Exception:
            return None

    @staticmethod
    def _has_border(style, edge):
        """True when that edge of the cell carries a line style.

        Revit gives the edge as a GraphicsStyle ElementId rather than a
        weight and colour. Resolving that to a real pen is a job of its own,
        so this only answers "is there a line here" and the writer draws a
        thin one - present or absent is the part that reads as the
        schedule's grid."""
        try:
            eid = getattr(style, 'Border{0}LineStyle'.format(edge))
            return eid is not None and eid != DB.ElementId.InvalidElementId
        except Exception:
            return False

    def _cell_style_dict(self, sd, r, c):
        """One cell's formatting, in the shape the workbook writer wants."""
        out = {}
        try:
            st = sd.GetTableCellStyle(r, c)
        except Exception:
            return out
        if st is None:
            return out
        bg = self._revit_color_hex(getattr(st, 'BackgroundColor', None))
        fg = self._revit_color_hex(getattr(st, 'TextColor', None))
        # White on white is Revit's "no fill" in practice; carrying it
        # through would paint every cell and lose the ones that are
        # deliberately filled.
        if bg and bg != u'FFFFFF':
            out['bg'] = bg
        if fg and fg != u'000000':
            out['fg'] = fg
        for key, attr in (('bold', 'IsFontBold'),
                          ('italic', 'IsFontItalic'),
                          ('underline', 'IsFontUnderline')):
            try:
                if bool(getattr(st, attr)):
                    out[key] = True
            except Exception:
                pass
        try:
            size = float(st.TextSize)
            if size > 0:
                # Revit reports the size in feet; Excel and ODF want points.
                out['size'] = round(size * 72.0, 1)
        except Exception:
            pass
        try:
            align = unicode(st.FontHorizontalAlignment)
            if 'Center' in align:
                out['align'] = 'center'
            elif 'Right' in align:
                out['align'] = 'right'
        except Exception:
            pass
        borders = [e[0].lower() for e in ('Left', 'Right', 'Top', 'Bottom')
                   if self._has_border(st, e)]
        if borders:
            out['borders'] = borders
        return out

    def _read_schedule_grid(self, schedule):
        """The whole schedule as {'rows': [[cell, ...]], 'merges': [...]}.

        A cell is {'text': ..., plus whatever formatting it carries}. Walks
        every section the table has, in order, so header rows, body rows,
        group headers and totals all arrive as the rows they appear as in
        the schedule."""
        rows, merges = [], []
        try:
            tbl = schedule.GetTableData()
        except Exception as ex:
            logger.warning('Could not read table data for %s: %s',
                           schedule.Name, ex)
            return {'rows': [], 'merges': []}

        sections = []
        for sec_name in ('Header', 'Body', 'Summary', 'Footer'):
            try:
                sec = getattr(DB.SectionType, sec_name)
            except Exception:
                continue
            try:
                sd = tbl.GetSectionData(sec)
            except Exception:
                continue
            if sd is not None:
                sections.append(sd)

        seen_merges = set()
        for sd in sections:
            try:
                n_rows = sd.NumberOfRows
                n_cols = sd.NumberOfColumns
            except Exception:
                continue
            row_offset = len(rows)
            for r in range(n_rows):
                row = []
                for c in range(n_cols):
                    try:
                        text = sd.GetCellText(r, c)
                    except Exception:
                        text = u''
                    cell = {'text': text or u''}
                    cell.update(self._cell_style_dict(sd, r, c))
                    row.append(cell)
                    # Merges are reported per member cell, so the same block
                    # comes back once for every cell it covers.
                    try:
                        m = sd.GetMergedCell(r, c)
                    except Exception:
                        m = None
                    if m is not None:
                        box = (m.Top + row_offset, m.Left,
                               m.Bottom + row_offset, m.Right)
                        if (box[0] != box[2] or box[1] != box[3]) and \
                                box not in seen_merges:
                            seen_merges.add(box)
                            merges.append(box)
                rows.append(row)
        return {'rows': rows, 'merges': merges}

    @staticmethod
    def _csv_line(cells):
        """One CSV record, quoted per RFC 4180.

        Hand-rolled rather than using the csv module: Python 2's csv writes
        bytes and mangles non-ASCII, and schedule text is full of it."""
        out = []
        for cell in cells:
            text = cell if isinstance(cell, unicode) else unicode(cell or '')
            if any(ch in text for ch in (u',', u'"', u'\n', u'\r')):
                text = u'"' + text.replace(u'"', u'""') + u'"'
            out.append(text)
        return u','.join(out)

    def _export_xls_csv(self, qitems, folder, tick, tmp_dir):
        """CSV holds exactly one table, so this is one file per schedule -
        named by the ordinary naming format, like every other per-item
        format. The workbook path below is the one that merges."""
        for qi in qitems:
            qi.status = 'Exporting'
            self._pump()
            sched = qi.source.revit_sheet
            try:
                header, body = self._read_schedule_rows(sched, tmp_dir)
                path = op.join(folder, qi.fname_noext + '.csv')
                with io.open(path, 'w', encoding='utf-8') as f:
                    f.write(self._csv_line(header) + u'\r\n')
                    for line in body:
                        f.write(self._csv_line(line) + u'\r\n')
                tick(qi, 'Done')
            except Exception as ex:
                logger.error('CSV %s failed: %s', qi.filename, ex)
                tick(qi, 'Failed')

    @staticmethod
    def _schedule_table(header, body, tab):
        """One schedule as (columns, rows) in the shape write_workbook wants.

        Rows carry their tab name in "_category" - that is the key the writer
        groups on, whatever the grouping means to the caller."""
        columns, rows, seen = [], [], set()
        keys = []
        for i, head in enumerate(header):
            key = head.strip() or 'Column {0}'.format(i + 1)
            keys.append(key)
            if key not in seen:
                seen.add(key)
                columns.append({'key': key, 'kind': 'instance',
                                'readonly': True})
        for line in body:
            row = {'_category': tab}
            for i, cell in enumerate(line):
                if i < len(keys):
                    row[keys[i]] = cell
            rows.append(row)
        return columns, rows

    def _export_xls(self, qitems, folder, tick):
        """Selected schedules to a spreadsheet.

        Two independent choices, the same pair PDF already offers: which
        format, and one file or one per schedule. CSV holds a single table
        by definition, so it is always separate files however the radio is
        set - xls_type_changed keeps the radios honest about that."""
        import shutil
        import tempfile

        ext    = self._xls_ext()
        single = ext != '.csv' and bool(self.xls_single_rb.IsChecked)
        tmp_dir = tempfile.mkdtemp(prefix='pysheets_xls_')
        try:
            if single:
                self._export_xls_single(qitems, folder, tick, tmp_dir, ext)
            elif ext == '.csv':
                self._export_xls_csv(qitems, folder, tick, tmp_dir)
            else:
                self._export_xls_separate(qitems, folder, tick, tmp_dir, ext)
        finally:
            try:
                shutil.rmtree(tmp_dir, ignore_errors=True)
            except Exception:
                pass

    def _export_xls_single(self, qitems, folder, tick, tmp_dir, ext):
        """All selected schedules into one workbook, a tab each."""
        for qi in qitems:
            qi.status = 'Exporting'
        self._pump()

        columns, rows, exported = [], [], []
        seen_cols = set()
        for qi in qitems:
            sched = qi.source.revit_sheet
            try:
                header, body = self._read_schedule_rows(sched, tmp_dir)
            except Exception as ex:
                logger.error('XLS %s failed: %s', qi.filename, ex)
                tick(qi, 'Failed')
                continue
            cols, rws = self._schedule_table(header, body, sched.Name)
            for col in cols:
                if col['key'] not in seen_cols:
                    seen_cols.add(col['key'])
                    columns.append(col)
            rows.extend(rws)
            exported.append(qi)

        if not exported:
            return
        path = op.join(folder, self._xls_workbook_name(qitems) + ext)
        try:
            # No Legend tab and no sheet protection: those exist for pyTable's
            # edit-and-reimport round trip. This is a plain dump of the
            # schedule, so there is no colour coding to explain and nothing
            # to guard against being edited.
            write_workbook(path, columns, rows, locked=False, legend=False)
        except Exception as ex:
            logger.error('XLS workbook failed: %s', ex)
            for qi in exported:
                tick(qi, 'Failed')
            return
        for qi in exported:
            tick(qi, 'Done')

    def _export_xls_separate(self, qitems, folder, tick, tmp_dir, ext):
        """One workbook per schedule, a single tab in each - named by the
        ordinary naming format, like every other per-item format."""
        for qi in qitems:
            qi.status = 'Exporting'
            self._pump()
            sched = qi.source.revit_sheet
            try:
                header, body = self._read_schedule_rows(sched, tmp_dir)
                columns, rows = self._schedule_table(header, body, sched.Name)
                write_workbook(op.join(folder, qi.fname_noext + ext),
                               columns, rows, locked=False, legend=False)
                tick(qi, 'Done')
            except Exception as ex:
                logger.error('XLS %s failed: %s', qi.filename, ex)
                tick(qi, 'Failed')

    def _xls_workbook_name(self, qitems):
        """Name for the single workbook, from the combined naming format -
        the same one the combined PDF uses, since both are one file made of
        many items."""
        return self._combined_pdf_name(qitems)

    def xls_type_changed(self, sender, args):
        """Grey out what CSV cannot do rather than let the control claim
        otherwise: a CSV holds one table, so it is always separate files."""
        if self._loading:
            return
        try:
            is_csv = self._xls_ext() == '.csv'
            if is_csv:
                self.xls_separate_rb.IsChecked = True
            self.xls_single_rb.IsEnabled = not is_csv
            self.xls_single_rb.Opacity   = 0.55 if is_csv else 1.0
        except Exception as ex:
            logger.warning('Table options: could not sync file mode: %s', ex)

    def _restore_print_settings(self):
        """Put back the print setting that was current before pySheets ran.

        Against the document it was read from, not whatever the dropdown
        happens to be showing at close. With a link selected, the old code
        opened a transaction on the link - Revit refuses that ("Transactions
        can only be used in primary documents") and the exception escaped
        window_closing, so the window died on the way out.

        Never raises: this runs while the window is closing, and there is
        nothing useful a failure here can do except be logged."""
        doc = self._init_psettings_doc
        if not self._init_psettings or doc is None:
            return
        if getattr(doc, 'IsLinked', False):
            return          # nothing was ever changed there to put back
        try:
            pm = doc.PrintManager
            if not pm:
                return
            with revit.Transaction('Restore Print Settings', doc=doc):
                pm.PrintSetup.CurrentPrintSetting = self._init_psettings
        except Exception as ex:
            logger.warning('Could not restore the original print setting: %s', ex)

    # ── ROW CLICK / MULTI-SELECT  (Shift / Ctrl highlight) ──
    # ── ROW CLICK — Ctrl/Shift directly checks/unchecks the checkboxes ──
    def _find_row_and_index(self, hit):
        """Walk up the visual tree from a click's OriginalSource to find
        the DataGridRow (and its index in the current visible order), or
        (None, -1) if the click didn't land on a row, or landed on a
        checkbox (which handles its own clicks)."""
        obj = hit
        while obj is not None:
            if isinstance(obj, Windows.Controls.CheckBox):
                return None, -1
            if isinstance(obj, Windows.Controls.DataGridRow):
                break
            obj = Windows.Media.VisualTreeHelper.GetParent(obj)

        row = None
        obj = hit
        while obj is not None:
            if isinstance(obj, Windows.Controls.DataGridRow):
                row = obj
                break
            obj = Windows.Media.VisualTreeHelper.GetParent(obj)
        if row is None:
            return None, -1

        sheets = self._visible_sheets
        idx = -1
        for i, s in enumerate(sheets):
            if s is row.Item:
                idx = i
                break
        if idx == -1:
            return None, -1
        return row, idx

    def _toggle_row_select(self, row, idx, shift, ctrl):
        """Handles both is_selected (checkbox/export state) and
        is_highlighted (the active multi-select group used for dragging
        several rows at once, a lighter green, separate from the
        checkbox-selected highlight)."""
        sheets = self._visible_sheets
        item = row.Item

        if shift and self._last_row_index >= 0:
            lo = min(self._last_row_index, idx)
            hi = max(self._last_row_index, idx)
            for s in sheets:
                s.is_highlighted = False
            for i in range(lo, hi + 1):
                sheets[i].is_selected = True
                sheets[i].is_highlighted = True
        elif shift:
            item.is_selected = True
            for s in sheets:
                s.is_highlighted = False
            item.is_highlighted = True
            self._last_row_index = idx
        elif ctrl:
            item.is_selected = not item.is_selected
            item.is_highlighted = not item.is_highlighted
            self._last_row_index = idx
        else:
            item.is_selected = not item.is_selected
            for s in sheets:
                s.is_highlighted = False
            item.is_highlighted = True
            self._last_row_index = idx

        self.sheets_dg.Items.Refresh()
        self._update_sel_count()
        self._update_select_all_cb()

    def _row_mouse_down(self, sender, args):
        """Shift/Ctrl clicks select immediately (no drag possible with a
        modifier held). A plain click is deferred: MouseMove decides
        whether it becomes a drag-reorder, MouseUp applies it as a normal
        toggle-select if it never moved far enough to count as a drag.
        If the clicked row is already part of the active highlighted
        group, the group is preserved in case this turns into a group
        drag, MouseUp collapses it to just this row otherwise."""
        try:
            row, idx = self._find_row_and_index(args.OriginalSource)
            if row is None:
                return

            ctrl  = (Windows.Input.Keyboard.IsKeyDown(Windows.Input.Key.LeftCtrl) or
                     Windows.Input.Keyboard.IsKeyDown(Windows.Input.Key.RightCtrl))
            shift = (Windows.Input.Keyboard.IsKeyDown(Windows.Input.Key.LeftShift) or
                     Windows.Input.Keyboard.IsKeyDown(Windows.Input.Key.RightShift))

            if shift or ctrl:
                self._toggle_row_select(row, idx, shift, ctrl)
                args.Handled = True
                return

            # Plain click/drag-start
            self._drag_row          = row
            self._drag_start_pt     = args.GetPosition(self.sheets_dg)
            self._drag_active       = False
            self._drag_preserve_grp = bool(row.Item.is_highlighted)
        except Exception as ex:
            logger.error('Row mouse down failed: %s', ex)

    def _row_mouse_move(self, sender, args):
        """Once a plain click has moved far enough, start a real drag
        instead of letting it resolve as a click on MouseUp. Drags the
        whole active-highlight group if the click started on a row that
        was already part of it, otherwise just this one row."""
        try:
            if self._drag_row is None or self._drag_active:
                return
            if args.LeftButton != Windows.Input.MouseButtonState.Pressed:
                return
            pt = args.GetPosition(self.sheets_dg)
            dx = abs(pt.X - self._drag_start_pt.X)
            dy = abs(pt.Y - self._drag_start_pt.Y)
            if dx < 6 and dy < 6:
                return

            self._drag_active = True
            row = self._drag_row
            self._drag_row = None

            if self._drag_preserve_grp:
                dragged = [s for s in self._visible_sheets if s.is_highlighted]
                if not any(s is row.Item for s in dragged):
                    dragged = [row.Item]
            else:
                for s in self._visible_sheets:
                    s.is_highlighted = False
                row.Item.is_highlighted = True
                dragged = [row.Item]
            self.sheets_dg.Items.Refresh()

            data = Windows.DataObject('pySheetsRows', dragged)
            Windows.DragDrop.DoDragDrop(row, data, Windows.DragDropEffects.Move)
            self._drag_active = False
            self._hide_drop_indicator()
        except Exception as ex:
            logger.error('Row drag start failed: %s', ex)
            self._drag_active = False
            self._hide_drop_indicator()

    def _row_mouse_up(self, sender, args):
        """If a plain click never turned into a drag, apply it as a
        normal toggle-select now."""
        try:
            row = self._drag_row
            self._drag_row = None
            if row is None or self._drag_active:
                return
            sheets = self._visible_sheets
            idx = -1
            for i, s in enumerate(sheets):
                if s is row.Item:
                    idx = i
                    break
            if idx == -1:
                return
            self._toggle_row_select(row, idx, False, False)
            args.Handled = True
        except Exception as ex:
            logger.error('Row mouse up failed: %s', ex)

    def _find_target_row(self, pos):
        """Hit-test the DataGrid at pos, return the DataGridRow under it, or None."""
        hit = Windows.Media.VisualTreeHelper.HitTest(self.sheets_dg, pos)
        if hit is None:
            return None
        obj = hit.VisualHit
        while obj is not None:
            if isinstance(obj, Windows.Controls.DataGridRow):
                return obj
            obj = Windows.Media.VisualTreeHelper.GetParent(obj)
        return None

    def _row_top(self, row):
        transform = row.TransformToAncestor(self.sheets_dg)
        return transform.Transform(Windows.Point(0, 0)).Y

    def _update_drop_indicator(self, pos):
        """Show the green line at the row boundary nearest the cursor,
        above the target row if dropping in its top half, below it
        otherwise."""
        try:
            target_row = self._find_target_row(pos)
            if target_row is None:
                self._hide_drop_indicator()
                return
            top = self._row_top(target_row)
            height = target_row.ActualHeight
            y = top + height if pos.Y > top + height / 2.0 else top
            self.drop_indicator.Margin = Windows.Thickness(0, y - 1.5, 0, 0)
            self.drop_indicator.Visibility = Windows.Visibility.Visible
        except Exception as ex:
            logger.error('Drop indicator update failed: %s', ex)

    def _hide_drop_indicator(self):
        try:
            self.drop_indicator.Visibility = Windows.Visibility.Collapsed
        except Exception:
            pass

    def _row_drag_over(self, sender, args):
        try:
            if not args.Data.GetDataPresent('pySheetsRows'):
                args.Effects = getattr(Windows.DragDropEffects, 'None')
                args.Handled = True
                return
            args.Effects = Windows.DragDropEffects.Move
            self._update_drop_indicator(args.GetPosition(self.sheets_dg))
            args.Handled = True
        except Exception as ex:
            logger.error('Row drag over failed: %s', ex)

    def _row_drag_leave(self, sender, args):
        self._hide_drop_indicator()

    def _row_drop(self, sender, args):
        """Drop one or more dragged rows at the row under the cursor.
        Reorders _ordered_sheets, marks the order mode as manual, and
        syncs the dropdown to reflect it."""
        try:
            self._hide_drop_indicator()
            if not args.Data.GetDataPresent('pySheetsRows'):
                return
            dragged_list = list(args.Data.GetData('pySheetsRows'))
            if not dragged_list:
                return
            dragged_ids = set(id(s) for s in dragged_list)

            pos = args.GetPosition(self.sheets_dg)
            target_row = self._find_target_row(pos)

            visible = list(self._visible_sheets)
            remaining = [s for s in visible if id(s) not in dragged_ids]
            # Preserve the dragged items' own relative order within the group
            dragged_in_order = [s for s in visible if id(s) in dragged_ids]

            target_idx = None
            if target_row is not None and id(target_row.Item) not in dragged_ids:
                for i, s in enumerate(remaining):
                    if s is target_row.Item:
                        target_idx = i
                        break
                if target_idx is not None:
                    top = self._row_top(target_row)
                    height = target_row.ActualHeight
                    if pos.Y > top + height / 2.0:
                        target_idx += 1

            if target_idx is None:
                new_visible = remaining + dragged_in_order
            else:
                new_visible = remaining[:target_idx] + dragged_in_order + remaining[target_idx:]

            # A search filter may be hiding some sheets, only reorder the
            # visible subset in place, keep every hidden sheet exactly
            # where it already was relative to everything else.
            full = self._ordered_sheets
            visible_ids = set(id(s) for s in new_visible)
            new_full = []
            vis_iter = iter(new_visible)
            for s in full:
                if id(s) in visible_ids:
                    new_full.append(next(vis_iter))
                else:
                    new_full.append(s)

            self._ordered_sheets = new_full
            ids = [get_elementid_value(s.revit_sheet.Id) for s in new_full]
            if self._sv_mode == 'sheets':
                self._manual_sheet_order = ids
                self._sheet_order_mode   = 'manual'
            else:
                self._manual_view_order = ids
                self._view_order_mode   = 'manual'

            if self._sv_mode == 'sheets':
                self._all_sheet_items = list(self._ordered_sheets)
            else:
                self._all_view_items = list(self._ordered_sheets)

            self._sync_order_dropdown()
            self._apply_filter()
            args.Handled = True
        except Exception as ex:
            logger.error('Row drop failed: %s', ex)

    def _sort_value(self, item, path):
        """Resolve a DataGrid SortMemberPath against a row item, including
        custom_params[key] indexer paths."""
        if path.endswith(']') and '[' in path:
            base, idx = path.split('[', 1)
            idx = idx.rstrip(']').strip("'\"")
            d = getattr(item, base, None) or {}
            try:
                return d.get(idx, '')
            except Exception:
                return ''
        return getattr(item, path, '')

    def sheets_dg_sorting(self, sender, args):
        """Native DataGrid header-click sort. We sort _ordered_sheets
        ourselves (instead of letting the DataGrid sort only its view)
        so the Print Queue and export stay in sync. Marks order as manual."""
        try:
            col = args.Column
            path = col.SortMemberPath
            if not path:
                return
            args.Handled = True

            ascending = col.SortDirection != ListSortDirection.Ascending
            for c in self.sheets_dg.Columns:
                c.SortDirection = None
            col.SortDirection = (ListSortDirection.Ascending if ascending
                                  else ListSortDirection.Descending)

            def key_fn(s):
                v = self._sort_value(s, path)
                return (v is None, v)

            sheets = sorted(self._ordered_sheets, key=key_fn, reverse=not ascending)
            self._ordered_sheets = sheets
            ids = [get_elementid_value(s.revit_sheet.Id) for s in sheets]
            if self._sv_mode == 'sheets':
                self._manual_sheet_order = ids
                self._sheet_order_mode   = 'manual'
                self._all_sheet_items    = list(sheets)
            else:
                self._manual_view_order  = ids
                self._view_order_mode    = 'manual'
                self._all_view_items     = list(sheets)

            self._sync_order_dropdown()
            self._apply_filter()
        except Exception as ex:
            logger.error('Header sort failed: %s', ex)

    def _window_click(self, sender, args):
        # Do NOT reset _last_row_index here — it must survive between clicks
        # so Shift+click range selection works correctly.
        pass

    # ── XAML EVENT HANDLERS TAB NAV ──
    def tab_select_clicked(self, sender, args):
        self._show_tab(TAB_SELECT)

    def tab_export_clicked(self, sender, args):
        missing = self._missing_selection_formats()
        if missing:
            self._warn_missing_selection(missing)
            return
        self._show_tab(TAB_EXPORT)

    def tab_print_clicked(self, sender, args):
        missing = self._missing_selection_formats()
        if missing:
            self._warn_missing_selection(missing)
            return
        self._build_queue()
        self._show_tab(TAB_PRINT)

    def _missing_selection_formats(self):
        """Enabled formats that have zero sheets/views selected."""
        return [f for f in ALL_FORMATS
                if f in self._fmt_enabled and not self._get_checked_for_format(f)]

    def _warn_missing_selection(self, missing):
        word = 'views' if self._sv_mode == 'views' else 'sheets'
        dlg.message('Please select {} for {}.'.format(word, ', '.join(missing)))

    def do_print_clicked(self, sender, args):
        """Export button on the Export tab."""
        self._do_export()

    def header_action_clicked(self, sender, args):
        """Header button: advances a tab on Select/Settings, runs the export
        on the Export tab. Same validation as the Select/Settings tab clicks."""
        if self._current_tab == TAB_SELECT:
            self.tab_export_clicked(sender, args)
        elif self._current_tab == TAB_EXPORT:
            self.tab_print_clicked(sender, args)
        else:
            self.do_print_clicked(sender, args)

    def help_clicked(self, sender, args):
        """☰ → Email support: open a pre-filled support email in the default
        mail client, addressed to Seed43 support, with the extension version
        and which app it came from already filled in."""
        self._open_url(support_mailto("pySheets", op.dirname(__file__)),
                       title="Support")

    def settings_toggle_preview_down(self, sender, args):
        """Explicit close-on-reclick. Popup StaysOpen=False already auto-closes
        on any click outside it, including a second click on this same toggle
        button, and that same click's Click event would then flip IsChecked
        back to True, reopening it. Intercept here first so a re-click always
        just closes, instead of closing and instantly reopening."""
        if self.settings_popup.IsOpen:
            self.settings_popup.IsOpen = False
            self.settings_toggle_btn.IsChecked = False
            args.Handled = True

    def issue_clicked(self, sender, args):
        """☰ → Report an issue: open a new GitHub issue, pre-filled with the
        app name, Seed43 version and Revit version."""
        self._open_url(github_issue_url("pySheets", op.dirname(__file__)),
                       title="Report an issue")

    def about_clicked(self, sender, args):
        """☰ → About: open ABOUT_URL in the default browser."""
        self._open_url(ABOUT_URL, title="About")

    def _open_url(self, url, title=''):
        """Open a URL in the default browser. The launch itself lives in
        Snippets._support.open_url; this only supplies pySheets' error
        reporting."""
        open_url(url, window=self,
                 on_error=lambda msg: dlg.message(msg, title=title))

    def support_clicked(self, sender, args):
        """☰ → Support: open the Buy Me a Coffee page in the default browser."""
        self._open_url('https://buymeacoffee.com/seed43', title='Support')

    def footer_pyrevit_clicked(self, sender, args):
        self._open_url('https://github.com/pyrevitlabs/pyRevit')

    def footer_ryan_clicked(self, sender, args):
        self._open_url('https://github.com/McCulloughRT/PrintFromIndex')

    # ── XAML EVENT HANDLERS SELECT TAB ──
    def doc_changed(self, sender, args):
        # _project_info deliberately stays on the host model (set once in
        # __init__ from revit.doc) and is NOT re-read from the selected
        # document. Picking a link changes which sheets you are exporting,
        # not whose job it is: a link's own project number and name belong to
        # somebody else's model, and letting them through would file the
        # output under the consultant's number instead of this project's.
        # Clear ALL containers — document changed so all selections are stale
        for c in self._sheet_selections.values(): c.clear()
        for c in self._view_selections.values():  c.clear()
        # Print settings first: _setup_sheet_sets rebuilds the sheet list, and
        # that reads the selected print setting. Left in the old order it
        # would still be holding the previous document's choice - which is
        # how switching to a link used to reach the variable-paper path with
        # a link as the document.
        self._setup_print_settings()
        self._setup_sheet_sets()
        self._setup_export_setups()

    def sheetset_changed(self, sender, args):
        """Picking a set ticks its sheets. It does not hide the others.

        A set is a starting selection, not a filter: the grid still holds
        every sheet in the document, so sheets can be added to or taken out
        of what the set gave you. <All Sheets> ticks nothing - it means "no
        set chosen", the blank slate the window opens on, not select-all."""
        # Replaces the selection rather than adding to it, and does so for
        # every format - picking a set has always reset all of them, back
        # when it rebuilt the list instead.
        for c in self._sheet_selections.values():
            c.clear()
        self._select_sheetset(self._selected_sheetset)

        # The rows themselves no longer depend on the set, so re-filtering is
        # enough - _apply_filter restores every tick from the containers, and
        # rebuilding the whole list here would re-read title blocks and custom
        # parameters for every sheet in the model for nothing.
        if self._all_sheet_items:
            self._apply_filter()
        else:
            self._reload_sheet_list()

    def _select_sheetset(self, sset):
        """Tick every sheet belonging to `sset`, in all formats' containers."""
        if not isinstance(sset, NamedSheetSet):
            return
        for rs in sset.get_sheets(self._selected_doc):
            if getattr(rs, 'IsPlaceholder', False):
                continue
            sid = get_elementid_value(rs.Id)
            for c in self._sheet_selections.values():
                c.update(sid, True)

    def search_changed(self, sender, args):
        self._apply_filter()

    def active_filter_changed(self, sender, args):
        self._apply_filter()

    def placed_only_changed(self, sender, args):
        self._state_filter = 'Placed views' if self.placed_only_cb.IsChecked else None
        self._apply_filter()

    def save_sheet_set(self, sender, args):
        """Save current selection as a new ViewSheetSet in Revit."""
        checked = self._checked_sheets
        if not checked:
            dlg.message('Select at least one sheet to save a set.')
            return
        if getattr(self._selected_doc, 'IsLinked', False):
            dlg.message('Sheet sets cannot be saved in a linked model.\n'
                        'Switch to the main document first.')
            return
        name = dlg.ask_string('Name for the new sheet set:',
                              title='Save Sheet Set')
        if not name:
            return
        pm = self._get_printmanager()
        if not pm:
            dlg.message('This document has no print manager, so a sheet set '
                        'cannot be saved from it.')
            return
        try:
            with revit.Transaction('Save Sheet Set', doc=self._selected_doc):
                # ViewSheetSetting is only reachable while PrintRange is
                # Select - on Current or All, Revit refuses with "This
                # property is only available when user choose Select of Print
                # Range". Printing sets it to Current, so whether this worked
                # depended on what had been done in the window beforehand.
                pm.PrintRange = DB.PrintRange.Select
                vss = pm.ViewSheetSetting
                # DB.ViewSet is required — List[View] causes a TypeError
                view_set = DB.ViewSet()
                for s in checked:
                    view_set.Insert(s.revit_sheet)
                vss.CurrentViewSheetSet.Views = view_set
                vss.SaveAs(name)
        except Exception as ex:
            # Never let this escape: it runs from a button handler, and an
            # exception there unwinds out of ShowDialog and closes pySheets.
            logger.error('Save Sheet Set failed: %s', ex)
            dlg.message('Could not save the sheet set.\n\n' + str(ex))
            return
        self._setup_sheet_sets()

    def _set_size_col_header(self, content):
        """Header is either the string 'Size' or the view-type ComboBox."""
        try:
            col = self.size_col
        except Exception:
            col = None
            try:
                for c in self.sheets_dg.Columns:
                    if c.Header in ('Size', 'Type') or isinstance(
                            c.Header, Windows.Controls.ComboBox):
                        col = c
                        break
            except Exception as ex:
                logger.warning('Header swap failed: %s', ex)
        if col is None:
            return
        col.Header = content
        col.Width = Windows.Controls.DataGridLength(
            136 if isinstance(content, Windows.Controls.ComboBox) else 86)

    def sv_sheets_clicked(self, sender, args):
        """Switch to Sheets mode — restores previously saved sheet selections."""
        if self._sv_mode == 'sheets':
            return
        self._sv_mode = 'sheets'
        self.sv_sheets_btn.Tag = 'Viewing'
        self.sv_views_btn.Tag  = ''
        self.active_only_cb.Visibility = Windows.Visibility.Visible
        self.active_only_cb.IsEnabled  = True
        self.placed_only_cb.Visibility = Windows.Visibility.Collapsed
        self.placed_only_cb.IsChecked  = False
        self._state_filter = None
        self._set_size_col_header(self.size_filter_header)
        self._refresh_builtin_column_visibility()
        self._refresh_dynamic_column_visibility()
        self._set_sheets_headers()
        self.order_ascending_item.Content = 'Sheet Number (Ascending)'
        # Point to sheets container for current format — do NOT clear it
        self._selection = self._sheet_selections[self._fmt_viewing]
        self._reload_sheet_list()

    def sv_views_clicked(self, sender, args):
        """Switch to Views mode — restores previously saved view selections."""
        if self._sv_mode == 'views':
            return
        self._sv_mode = 'views'
        self.sv_views_btn.Tag  = 'Viewing'
        self.sv_sheets_btn.Tag = ''
        self.active_only_cb.IsChecked  = False
        self.active_only_cb.Visibility = Windows.Visibility.Collapsed
        self.placed_only_cb.Visibility = Windows.Visibility.Visible
        self._set_size_col_header(getattr(self, '_type_cb', None) or 'Type')
        self._refresh_builtin_column_visibility()
        self._refresh_dynamic_column_visibility()
        self._set_views_headers()
        self.order_ascending_item.Content = 'Name (Ascending)'
        # Point to views container for current format — do NOT clear it
        self._selection = self._view_selections[self._fmt_viewing]
        self._reload_sheet_list()

    def order_changed(self, sender, args):
        """Order dropdown explicitly changed by the user. 'Manual order' has
        no preset to sort to, it's a detected state, not an action, so
        selecting it is a no-op. Browser/Ascending actively re-sort."""
        if self._updating_order_cb or self._loading:
            return
        item = self.order_cb.SelectedItem
        if item is None:
            return
        mode = item.Tag
        if mode == 'manual':
            return

        doc = self._selected_doc
        for_sheets = (self._sv_mode == 'sheets')
        key_fn = self._order_ascending_key()

        if for_sheets:
            self._sheet_order_mode = mode
            self._ordered_sheets = self._apply_order(
                self._all_sheets, mode, self._manual_sheet_order, doc, True, key_fn)
        else:
            self._view_order_mode = mode
            self._ordered_sheets = self._apply_order(
                self._all_sheets, mode, self._manual_view_order, doc, False, key_fn)

        if self._sv_mode == 'sheets':
            self._all_sheet_items = list(self._ordered_sheets)
        else:
            self._all_view_items = list(self._ordered_sheets)

        self._apply_filter()

    def fmt_pill_clicked(self, sender, args):
        """Any format pill clicked — read which one from button Content."""
        fmt = sender.Content
        self._fmt_pill_logic(fmt)

    def include_checkbox_clicked(self, sender, args):
        # Set the shift-anchor to this row so Shift+click from here works.
        # sender is the CheckBox; Tag="{Binding}" gives us the data item.
        # Use identity (is) not equality (==) to find index reliably in IronPython.
        try:
            item = sender.Tag
            if item is not None:
                sheets = self._visible_sheets
                for i, s in enumerate(sheets):
                    if s is item:
                        self._last_row_index = i
                        break
                # If this row is part of the active highlighted group, the
                # checkbox click applies to the whole group (same behavior
                # as select_all_clicked when a group is active).
                if item.is_highlighted:
                    new_state = item.is_selected
                    for s in sheets:
                        if s.is_highlighted:
                            s.is_selected = new_state
                    self.sheets_dg.Items.Refresh()
        except Exception:
            pass
        self._update_sel_count()
        self._update_select_all_cb()

    def grid_selection_changed(self, sender, args):
        try:
            self.sheets_dg.UnselectAll()
        except Exception:
            pass

    def select_all_clicked(self, sender, args):
        # Master checkbox always toggles every visible sheet, regardless of
        # any active highlighted (multi-select) group.
        sheets = self._visible_sheets

        checked = [s for s in sheets if s.is_selected]
        new_state = len(checked) < len(sheets)
        for s in sheets:
            s.is_selected = new_state
        self.sheets_dg.Items.Refresh()
        self._update_sel_count()
        self._update_select_all_cb()

    def select_all_sheets(self, sender, args):
        for s in self._visible_sheets:
            s.is_selected = True
        self.sheets_dg.Items.Refresh()
        self._update_sel_count()
        self._update_select_all_cb()

    def deselect_all_sheets(self, sender, args):
        for s in self._visible_sheets:
            s.is_selected = False
        self.sheets_dg.Items.Refresh()
        self._update_sel_count()
        self._update_select_all_cb()

    def naming_format_changed(self, sender, args):
        """User changed the naming format dropdown in the column header."""
        if self._loading:
            return
        nf = self._selected_naming_format
        if nf:
            self._fmt_naming[self._fmt_viewing] = nf.name
            self._save_naming_memory()
            self._apply_filter()

    def edit_naming_formats(self, sender, args):
        """Open EditNamingFormats dialog."""
        EditNamingFormatsWindow(
            _xaml_path(op.join('tools', 'EditNamingFormats.xaml')),
            start_with=self._selected_naming_format
        ).show_dialog()
        # Refresh dropdown and filenames
        self._setup_naming_formats()
        self._apply_filter()

    def edit_combined_naming_formats(self, sender, args):
        """Open EditNamingFormats dialog against the Combined PDF Name's
        own independent naming store."""
        EditNamingFormatsWindow(
            _xaml_path(op.join('tools', 'EditNamingFormats.xaml')),
            start_with=self.combined_naming_cb.SelectedItem,
            naming_dir=COMBINED_NAMING_DIR,
            default_formats=COMBINED_DEFAULT_FORMATS
        ).show_dialog()
        self._setup_naming_formats()

    # ── XAML EVENT HANDLERS — EXPORT SETTINGS TAB ──
    def _read_ui_print_settings(self, pm):
        """
        Apply all UI control values to the PrintManager's InSession print setting.
        Call this before printing when the user has not selected a named setting,
        OR to apply any UI overrides on top of a loaded setting.
        """
        try:
            ps = pm.PrintSetup
            try:
                ips = ps.InSessionPrintSetting   # pre-2026
            except Exception:
                ips = ps.CurrentPrintSetting     # Revit 2026+

            # ── Orientation ──
            if self.orient_portrait_btn.Tag == 'Viewing':
                ips.PrintParameters.PageOrientation = DB.PageOrientationType.Portrait
            elif self.orient_landscape_btn.Tag == 'Viewing':
                ips.PrintParameters.PageOrientation = DB.PageOrientationType.Landscape
            # 'From Sheet' = leave as-is (Revit default)

            # ── Paper Size ──
            # papersize_cb was never wired to PrintParameters — selecting
            # "A3" here did nothing, so export silently kept whatever size
            # was already active (usually each sheet's own titleblock size).
            paper_name = (self.papersize_cb.Text or 'From Titleblock').strip()
            if paper_name and paper_name != 'From Titleblock':
                try:
                    match = next((p for p in pm.PaperSizes
                                  if p and p.Name == paper_name), None)
                    if match is None:
                        # Driver-reported names vary ("A3", "ISO A3", etc.)
                        match = next((p for p in pm.PaperSizes if p and
                                      paper_name.lower() in p.Name.lower()), None)
                    if match is not None:
                        ips.PrintParameters.PaperSize = match
                    else:
                        logger.warning(
                            'No paper size on the selected printer matches "%s"',
                            paper_name)
                except Exception as ex:
                    logger.warning('Could not set paper size "%s": %s', paper_name, ex)
            # 'From Titleblock' = leave as-is (Revit sizes to the sheet)

            # ── Zoom ──
            if self.zoom_fit_rb.IsChecked:
                ips.PrintParameters.ZoomType = DB.ZoomType.FitPage
                ips.PrintParameters.Zoom = 100
            else:
                ips.PrintParameters.ZoomType = DB.ZoomType.Zoom
                try:
                    ips.PrintParameters.Zoom = int(float(self.zoom_pct_tb.Text))
                except Exception:
                    ips.PrintParameters.Zoom = 100

            # ── Paper Placement ──
            if self.placement_center_rb.IsChecked:
                ips.PrintParameters.PaperPlacement = DB.PaperPlacementType.Center
            else:
                ips.PrintParameters.PaperPlacement = DB.PaperPlacementType.LowerLeft
                try:
                    ips.PrintParameters.UserDefinedMarginX = float(self.offset_x_tb.Text)
                    ips.PrintParameters.UserDefinedMarginY = float(self.offset_y_tb.Text)
                except Exception:
                    pass

            # ── Hidden Lines ──
            if self.hlv_raster_rb.IsChecked:
                ips.PrintParameters.HiddenLineViews = DB.HiddenLineViewsType.RasterProcessing
            else:
                ips.PrintParameters.HiddenLineViews = DB.HiddenLineViewsType.VectorProcessing

            # ── Raster Quality ──
            rq_map = {'High':   DB.RasterQualityType.High,
                      'Medium': DB.RasterQualityType.Medium,
                      'Low':    DB.RasterQualityType.Low}
            rq_text = (self.raster_quality_cb.SelectedItem.Content
                       if self.raster_quality_cb.SelectedItem else 'High')
            ips.PrintParameters.RasterQuality = rq_map.get(rq_text, DB.RasterQualityType.High)

            # ── Colors ──
            col_map = {'Color':          DB.ColorDepthType.Color,
                       'Black and White': DB.ColorDepthType.BlackLine,
                       'Grayscale':      DB.ColorDepthType.GrayScale}
            col_text = (self.colors_cb.SelectedItem.Content
                        if self.colors_cb.SelectedItem else 'Color')
            ips.PrintParameters.ColorDepth = col_map.get(col_text, DB.ColorDepthType.Color)

            # ── Options ──
            ips.PrintParameters.ViewLinksinBlue    = bool(self.opt_links_blue_cb.IsChecked)
            ips.PrintParameters.HideReforWorkPlanes = bool(self.opt_hide_refplanes_cb.IsChecked)
            ips.PrintParameters.HideUnreferencedViewTags = bool(self.opt_hide_unreftags_cb.IsChecked)
            ips.PrintParameters.HideScopeBoxes     = bool(self.opt_hide_scopeboxes_cb.IsChecked)
            ips.PrintParameters.HideCropBoundaries = bool(self.opt_hide_cropbounds_cb.IsChecked)
            ips.PrintParameters.ReplaceHalftoneWithThinLines = bool(self.opt_halftone_thin_cb.IsChecked)
            ips.PrintParameters.MaskCoincidentLines = bool(self.opt_region_edges_cb.IsChecked)

            ps.CurrentPrintSetting = ips
        except Exception as ex:
            logger.warning('Could not apply UI print settings: %s', ex)

    DEST_OPTIONS = ('file', 'printer')

    def _dest_btn(self, name):
        return self.dest_file_btn if name == 'file' else self.dest_printer_btn

    def dest_pill_clicked(self, sender, args):
        """Print-destination pills — same on/off + active-focus logic as the
        Select-tab format pills:
          - clicking the currently-focused pill toggles it off (focus jumps
            to the other one if it's still on)
          - clicking the other pill turns it on (if needed) and focuses it,
            without turning the first one off
        """
        name = 'file' if sender is self.dest_file_btn else 'printer'
        printer_was_on = 'printer' in self._dest_enabled
        if name == self._dest_viewing:
            if name in self._dest_enabled:
                self._dest_enabled.discard(name)
                others = [n for n in self.DEST_OPTIONS
                          if n in self._dest_enabled and n != name]
                if others:
                    self._dest_viewing = others[0]
            else:
                self._dest_enabled.add(name)
                self._dest_viewing = name
        else:
            self._dest_enabled.add(name)
            self._dest_viewing = name
        self._update_dest_buttons()
        self._update_dest_gates()
        self._update_dest_label()
        if 'printer' in self._dest_enabled and not printer_was_on:
            # Printer destination just turned on — discard any unsaved edits
            # and reload the fields from the currently-selected Print Setting.
            self._apply_printsetting_to_ui()
            self._reload_sheet_list()

    def _update_dest_buttons(self):
        for name in self.DEST_OPTIONS:
            btn = self._dest_btn(name)
            if name == self._dest_viewing and name in self._dest_enabled:
                btn.Tag = 'Viewing'
            elif name in self._dest_enabled:
                btn.Tag = 'Enabled'
            else:
                btn.Tag = ''
            btn.ToolTip = self._pill_tooltip(btn.Tag)

    def _set_print_destination(self, dest):
        self._dest_enabled = set()
        if dest in ('file', 'both'):
            self._dest_enabled.add('file')
        if dest in ('printer', 'both'):
            self._dest_enabled.add('printer')
        self._dest_viewing = 'printer' if dest == 'printer' else 'file'
        self._update_dest_buttons()
        self._update_dest_gates()
        self._update_dest_label()

    def _update_dest_gates(self):
        to_printer = 'printer' in self._dest_enabled
        self.printer_fields_panel.IsEnabled = to_printer
        self.printer_fields_panel.Opacity = 1.0 if to_printer else 0.4

    def _update_dest_label(self):
        if not self._dest_enabled:
            self.settings_scope_label.Text = 'Select destination'
        elif self._dest_viewing == 'printer':
            self.settings_scope_label.Text = 'Send to printer'
        else:
            self.settings_scope_label.Text = 'Export to file (PDF)'

    def _get_print_destination(self):
        """Return 'file', 'printer', 'both', or 'none'."""
        file_on    = 'file' in self._dest_enabled
        printer_on = 'printer' in self._dest_enabled
        if file_on and printer_on:
            return 'both'
        if printer_on:
            return 'printer'
        if file_on:
            return 'file'
        return 'none'

    # ── Print Setting management actions ──
    def ps_save_clicked(self, sender, args):
        pm = self._get_printmanager()
        if not pm:
            return
        cur = pm.PrintSetup.CurrentPrintSetting
        if isinstance(cur, DB.InSessionPrintSetting):
            dlg.message('Use "Save As" to save the in-session setting with a name.')
            return
        try:
            self._read_ui_print_settings(pm)
            with revit.Transaction('Save Print Setting', doc=self._selected_doc):
                pm.PrintSetup.Save()
            dlg.message('Print setting saved.')
        except Exception as ex:
            dlg.message('Save failed.\n\n' + str(ex))

    def ps_saveas_clicked(self, sender, args):
        pm = self._get_printmanager()
        if not pm:
            return
        name = dlg.ask_string('Name for the new print setting:',
                                    default='My Print Setting')
        if not name:
            return
        try:
            self._read_ui_print_settings(pm)
            with revit.Transaction('Save As Print Setting', doc=self._selected_doc):
                pm.PrintSetup.SaveAs(name)
            self._setup_print_settings()
            dlg.message('Saved as "{}".'.format(name))
        except Exception as ex:
            dlg.message('Save As failed.\n\n' + str(ex))

    def ps_revert_clicked(self, sender, args):
        pm = self._get_printmanager()
        if not pm:
            return
        try:
            with revit.Transaction('Revert Print Setting', doc=self._selected_doc):
                pm.PrintSetup.Revert()
            self._setup_print_settings()
        except Exception as ex:
            dlg.message('Revert failed.\n\n' + str(ex))

    def ps_rename_clicked(self, sender, args):
        pm = self._get_printmanager()
        if not pm:
            return
        cur = pm.PrintSetup.CurrentPrintSetting
        if isinstance(cur, DB.InSessionPrintSetting):
            dlg.message('Cannot rename the in-session setting. Use "Save As" instead.')
            return
        new_name = dlg.ask_string('New name:', default=cur.Name)
        if not new_name:
            return
        try:
            with revit.Transaction('Rename Print Setting', doc=self._selected_doc):
                pm.PrintSetup.Rename(new_name)
            self._setup_print_settings()
        except Exception as ex:
            dlg.message('Rename failed.\n\n' + str(ex))

    def ps_delete_clicked(self, sender, args):
        pm = self._get_printmanager()
        if not pm:
            return
        cur = pm.PrintSetup.CurrentPrintSetting
        if isinstance(cur, DB.InSessionPrintSetting):
            dlg.message('Cannot delete the in-session setting.')
            return
        if not dlg.confirm(
                'Delete print setting "{}"?\nThis cannot be undone.'.format(cur.Name),
                yes='Delete'):
            return
        try:
            with revit.Transaction('Delete Print Setting', doc=self._selected_doc):
                pm.PrintSetup.Delete()
            self._setup_print_settings()
        except Exception as ex:
            dlg.message('Delete failed.\n\n' + str(ex))

    def printer_changed(self, sender, args):
        """Called when the user selects a different PDF printer."""
        if self._loading:
            return
        self._reload_sheet_list()

    def printsetting_scroll(self, sender, args):
        """Mouse wheel over the closed dropdown cycles print settings."""
        if self.printsetting_cb.IsDropDownOpen:
            return   # let the open list scroll normally
        count = self.printsetting_cb.Items.Count
        if count == 0:
            return
        args.Handled = True
        idx = self.printsetting_cb.SelectedIndex
        if args.Delta > 0:
            idx = max(0, idx - 1)
        else:
            idx = min(count - 1, idx + 1)
        self.printsetting_cb.SelectedIndex = idx

    def printsetting_changed(self, sender, args):
        if self._loading:
            return
        self._apply_printsetting_to_ui()
        self._reload_sheet_list()

    def _apply_printsetting_to_ui(self):
        """Load the selected print setting's saved values into the panel."""
        item = self._selected_printsetting
        if item is None:
            # Never a fault, so never a warning. Two ordinary things land
            # here: a linked model, which has no print settings of its own so
            # the dropdown is legitimately empty; and the moment _setup_print_
            # settings assigns ItemsSource, which clears the selection and
            # fires this handler before the new item is picked. There is
            # simply nothing to apply either way.
            logger.debug('printsetting update skipped: nothing selected')
            return
        if item.allows_variable_paper:
            logger.debug('printsetting update skipped: variable paper item')
            return
        if item.print_settings is None:
            logger.warning('printsetting update skipped: item.print_settings is empty')
            return
        try:
            pp = item.print_settings.PrintParameters
        except Exception as ex:
            logger.warning('printsetting update failed reading PrintParameters: %s', ex)
            return
        self._loading = True
        try:
            try:
                if pp.PaperSize:
                    for it in self.papersize_cb.Items:
                        if getattr(it, 'Content', it) == pp.PaperSize.Name:
                            self.papersize_cb.SelectedItem = it
                            break
            except Exception as ex:
                logger.warning('printsetting: paper size field failed: %s', ex)
            try:
                po = pp.PageOrientation
                self.orient_portrait_btn.Tag = (
                    'Viewing' if po == DB.PageOrientationType.Portrait else '')
                self.orient_landscape_btn.Tag = (
                    'Viewing' if po == DB.PageOrientationType.Landscape else '')
            except Exception as ex:
                logger.warning('printsetting: orientation field failed: %s', ex)
            try:
                if pp.PaperPlacement == DB.PaperPlacementType.Center:
                    self.placement_center_rb.IsChecked = True
                else:
                    self.placement_offset_rb.IsChecked = True
                    # UserDefinedMarginX/Y only exist while MarginType is
                    # UserDefined. On No Margin or Printer Limit - both
                    # perfectly ordinary settings - reading them throws
                    # "Current PaperPlacement is NOT Margins and Current
                    # MarginType is NOT User defined", which is the warning
                    # this used to log on every print. Nothing to read then,
                    # so the offset boxes keep what they had.
                    if pp.MarginType == DB.MarginType.UserDefined:
                        self.offset_x_tb.Text = '{:.2f}'.format(
                            pp.UserDefinedMarginX)
                        self.offset_y_tb.Text = '{:.2f}'.format(
                            pp.UserDefinedMarginY)
            except Exception as ex:
                logger.warning('printsetting: paper placement field failed: %s', ex)
            try:
                raster = (pp.HiddenLineViews ==
                          DB.HiddenLineViewsType.RasterProcessing)
                self.hlv_raster_rb.IsChecked = raster
                self.hlv_vector_rb.IsChecked = not raster
            except Exception as ex:
                logger.warning('printsetting: hidden line views field failed: %s', ex)
            try:
                if pp.ZoomType == DB.ZoomType.FitToPage:
                    self.zoom_fit_rb.IsChecked = True
                else:
                    self.zoom_pct_rb.IsChecked = True
                    self.zoom_pct_tb.Text = str(pp.Zoom)
            except Exception as ex:
                logger.warning('printsetting: zoom field failed: %s', ex)
            try:
                rq = {DB.RasterQualityType.High: 'High',
                      DB.RasterQualityType.Medium: 'Medium',
                      DB.RasterQualityType.Low: 'Low'}.get(
                          pp.RasterQuality, 'High')
                for it in self.raster_quality_cb.Items:
                    if it.Content == rq:
                        self.raster_quality_cb.SelectedItem = it
                        break
            except Exception as ex:
                logger.warning('printsetting: raster quality field failed: %s', ex)
            try:
                cd = {DB.ColorDepthType.Color: 'Color',
                      DB.ColorDepthType.BlackLine: 'Black and White',
                      DB.ColorDepthType.GrayScale: 'Grayscale'}.get(
                          pp.ColorDepth, 'Color')
                for it in self.colors_cb.Items:
                    if it.Content == cd:
                        self.colors_cb.SelectedItem = it
                        break
            except Exception as ex:
                logger.warning('printsetting: colors field failed: %s', ex)
            try:
                self.opt_links_blue_cb.IsChecked      = pp.ViewLinksinBlue
                self.opt_hide_refplanes_cb.IsChecked  = pp.HideReforWorkPlanes
                self.opt_hide_unreftags_cb.IsChecked  = pp.HideUnreferencedViewTags
                self.opt_hide_scopeboxes_cb.IsChecked = pp.HideScopeBoxes
                self.opt_hide_cropbounds_cb.IsChecked = pp.HideCropBoundaries
                self.opt_halftone_thin_cb.IsChecked   = pp.ReplaceHalftoneWithThinLines
                self.opt_region_edges_cb.IsChecked    = pp.MaskCoincidentLines
            except Exception as ex:
                logger.warning('printsetting: options checkboxes field failed: %s', ex)
        finally:
            self._loading = False

    def orient_clicked(self, sender, args):
        """Portrait / Landscape toggles — neither on = from each sheet."""
        was_on = sender.Tag == 'Viewing'
        self.orient_portrait_btn.Tag  = ''
        self.orient_landscape_btn.Tag = ''
        if not was_on:
            sender.Tag = 'Viewing'

    def _show_exp_panel(self, fmt):
        """Show the Settings panel for `fmt`, hide the others."""
        pnl_map = {'PDF': self.pdf_settings_panel,
                   'DWG': self.dwg_settings_panel,
                   'DGN': self.dgn_settings_panel,
                   'NWC': self.nwc_settings_panel,
                   'IFC': self.ifc_settings_panel,
                   'IMG': self.img_settings_panel,
                   'XLS': self.xls_settings_panel}
        for f, pnl in pnl_map.items():
            pnl.Visibility = (Windows.Visibility.Visible if f == fmt
                              else Windows.Visibility.Collapsed)
        if fmt == 'NWC':
            self._check_nwc_available()

    def exp_subtab_clicked(self, sender, args):
        """Export Settings sub-tab clicked — show that format's panel.
        Also switches the Select tab's viewing format, since both tabs
        share one 'viewing' state."""
        fmt = sender.Content
        if fmt not in self._fmt_enabled:
            return   # not enabled on the Select tab — nothing to show
        self._switch_viewing_format(fmt)
        self._update_format_buttons()
        self._update_fmt_column_header()
        self._apply_filter()

    def _check_nwc_available(self):
        """Grey out NWC options + show overlay until the exporter is found."""
        if getattr(self, '_nwc_ok', False):
            return True
        ok = False
        try:
            ok = DB.OptionalFunctionalityUtils.IsNavisworksExporterAvailable()
        except Exception as ex:
            logger.warning('NWC check failed: %s', ex)
        self._nwc_ok = ok   # cache only when True; False re-checks next click
        if not ok:
            self._nwc_ok = False
        self.nwc_overlay.Visibility = (
            Windows.Visibility.Collapsed if ok else Windows.Visibility.Visible)
        self.nwc_options_panel.IsEnabled = ok
        self.nwc_options_panel.Opacity = 1.0 if ok else 0.35
        return ok

    # ── Open native Revit dialogs (closes pySheets first) ──
    def _post_revit_command(self, postable_name):
        """Queue a built-in Revit command and close this window so it runs."""
        try:
            pc  = getattr(UI.PostableCommand, postable_name)
            cid = UI.RevitCommandId.LookupPostableCommandId(pc)
            HOST_APP.uiapp.PostCommand(cid)
            self.Close()
        except Exception as ex:
            dlg.message('Could not open the Revit dialog.\n\n' + str(ex))

    def dwg_setups_clicked(self, sender, args):
        self._post_revit_command('ExportOptionsExportSetupsDWGOrDXF')

    def dgn_setups_clicked(self, sender, args):
        self._post_revit_command('ExportOptionsExportSetupsDGN')

    def ifc_dialog_clicked(self, sender, args):
        self._post_revit_command('ExportIFC')

    def img_dialog_clicked(self, sender, args):
        self._post_revit_command('ExportImagesandAnimationsImage')

    # ── XAML EVENT HANDLERS — PRINT TAB ──
    def browse_export_folder(self, sender, args):
        fbd = Forms.FolderBrowserDialog()
        fbd.ShowNewFolderButton = True
        if fbd.ShowDialog() == Forms.DialogResult.OK:
            self._export_folder = fbd.SelectedPath
            self.export_folder_tb.Text = self._export_folder
            self._active_folder_preset = None

    # ── Export folder presets ──
    FOLDER_PRESETS_FILE = op.join(USERDATA_DIR, 'folder_presets', 'folder_presets.json')

    def _load_folder_presets(self):
        try:
            if op.isfile(self.FOLDER_PRESETS_FILE):
                with open(self.FOLDER_PRESETS_FILE, 'r') as f:
                    data = json.load(f)
                # Migrate old {name: path} format
                changed = False
                for k, v in list(data.items()):
                    if isinstance(v, str):
                        data[k] = {'root': '', 'template': v}
                        changed = True
                if changed:
                    self._save_folder_presets(data)
                return data
        except Exception:
            pass
        return {}

    def _save_folder_presets(self, presets):
        try:
            fp_dir = op.dirname(self.FOLDER_PRESETS_FILE)
            if not op.isdir(fp_dir):
                os.makedirs(fp_dir)
            with open(self.FOLDER_PRESETS_FILE, 'w') as f:
                json.dump(presets, f, indent=2)
        except Exception as ex:
            logger.warning('Folder preset save failed: %s', ex)

    def _rebuild_folder_presets(self):
        """Export tab popup — a plain pick-list, no add/edit/delete here."""
        panel = self.folder_preset_panel
        panel.Children.Clear()
        presets = self._load_folder_presets()
        if not presets:
            tb = Windows.Controls.TextBlock()
            tb.Text = 'No presets saved yet — see \u2630 Export Location Presets'
            tb.Foreground = self.FindResource('BrushTextPrimary')
            tb.Opacity = 0.5
            tb.FontSize = 11
            tb.TextWrapping = Windows.TextWrapping.Wrap
            tb.MaxWidth = 200
            panel.Children.Add(tb)
            return
        for name in sorted(presets):
            btn = Windows.Controls.Button()
            btn.Content = name
            btn.Style = self.FindResource('SmallSecBtn')
            btn.HorizontalContentAlignment = Windows.HorizontalAlignment.Left
            btn.Margin = Windows.Thickness(0, 0, 0, 4)
            btn.Tag = name
            btn.Click += self.folder_preset_pick_clicked
            panel.Children.Add(btn)

    def folder_preset_pick_clicked(self, sender, args):
        presets = self._load_folder_presets()
        preset = presets.get(sender.Tag)
        self.folder_preset_btn.IsChecked = False
        if not preset:
            return
        path = fpe_resolve.resolve_path(
                preset.get('template', ''), preset.get('root', ''),
                self._project_info, HOST_APP.username, HOST_APP.version)
        try:
            if not op.isdir(path):
                os.makedirs(path)
            self._export_folder = path
            self.export_folder_tb.Text = path
            self._active_folder_preset = sender.Tag
        except Exception as ex:
            dlg.message('Could not create/access folder:\n' + path
                        + '\n\n' + str(ex))

    def manage_presets_clicked(self, sender, args):
        self.settings_toggle_btn.IsChecked = False
        changed = fpm_win.show_manager(
            self._project_info, HOST_APP.username, HOST_APP.version,
            self._load_folder_presets, self._save_folder_presets)
        if changed:
            self._rebuild_folder_presets()

    def manage_profiles_clicked(self, sender, args):
        self.settings_toggle_btn.IsChecked = False
        changed = profiles_win.show_manager(
            self._list_profile_names, self._delete_profile_by_name)
        if changed:
            self._setup_profiles()

    def export_settings_clicked(self, sender, args):
        self.settings_toggle_btn.IsChecked = False
        ies_win.show_export(SETTINGS_FOLDERS, SETTINGS_SYNC_FILE)

    def import_settings_clicked(self, sender, args):
        self.settings_toggle_btn.IsChecked = False
        if ies_win.show_import(SETTINGS_FOLDERS, SETTINGS_SYNC_FILE):
            self._setup_naming_formats()
            self._setup_profiles()

    # ── PROFILES  (save / load / import export settings) ──
    @staticmethod
    def _cb_name(combo):
        """Selected profile name from either profile dropdown, or None.

        Both dropdowns hold ProfileItem objects rather than plain strings, so
        every read of .SelectedItem goes through here."""
        try:
            item = combo.SelectedItem
        except Exception:
            return None
        return item.name if item is not None else None

    @staticmethod
    def _select_cb_name(combo, name):
        """Select the ProfileItem called `name`, or clear the selection."""
        try:
            for item in (combo.ItemsSource or []):
                if item.name == name:
                    combo.SelectedItem = item
                    return True
            combo.SelectedIndex = -1
        except Exception:
            pass
        return False

    def _setup_profiles(self, select=None):
        """Scan the profiles folder and fill the header + schedule dropdowns.

        Rebuilt in full whenever the armed set changes, since that is what
        repaints each item's accent dot."""
        try:
            if not op.isdir(PROFILES_DIR):
                os.makedirs(PROFILES_DIR)
            names = sorted(op.splitext(f)[0]
                           for f in os.listdir(PROFILES_DIR)
                           if f.lower().endswith('.json'))
            armed = self._armed_profile_names()

            keep_header = select or self._cb_name(self.profile_cb)
            keep_sched  = self._cb_name(self.sched_profile_cb)

            # Rebuilding drops the selection and puts it back, which would
            # otherwise fire sched_profile_changed twice and reload the card
            # mid-arm. Nothing here is a user edit, so hold the guard.
            was_loading = self._loading
            self._loading = True
            try:
                # Separate item objects per dropdown: one ProfileItem can only
                # sit in one ComboBox at a time.
                self.profile_cb.ItemsSource = [
                    ProfileItem(n, n in armed) for n in names]
                self._select_cb_name(self.profile_cb, keep_header)
                try:
                    self.sched_profile_cb.ItemsSource = [
                        ProfileItem(n, n in armed) for n in names]
                    self._select_cb_name(self.sched_profile_cb, keep_sched)
                except Exception:
                    pass
            finally:
                self._loading = was_loading
        except Exception as ex:
            logger.warning('Profiles setup failed: %s', ex)

    # ── Per-profile schedule block ──
    # The card's timing lives in its own profile JSON, so it travels with the
    # profile (import/export, copy to another machine). Deliberately kept out
    # of _gather_profile/_apply_profile: those also drive lastsession.json,
    # and restoring a session must never arm anything by itself.
    def _read_profile_schedule(self, name):
        """The `schedule` block of profile `name`, or None."""
        try:
            with open(self._profile_path(name), 'r') as f:
                block = json.load(f).get('schedule')
            return block if isinstance(block, dict) else None
        except Exception:
            return None

    def _write_profile_schedule(self, name, block):
        """Merge `block` into profile `name` as its `schedule` key, or drop
        the key when block is None. Rewrites nothing else in the file."""
        try:
            path = self._profile_path(name)
            with open(path, 'r') as f:
                data = json.load(f)
            if block is None:
                data.pop('schedule', None)
            else:
                data['schedule'] = block
            with open(path, 'w') as f:
                json.dump(data, f, indent=2)
        except Exception as ex:
            logger.warning('Could not save schedule on profile "%s": %s',
                           name, ex)

    @staticmethod
    def _armed_profile_names():
        """Names of every card currently armed, from the runtime file."""
        return set(e.get('profile_name')
                   for e in _read_armed_file().get('entries', [])
                   if e.get('profile_name'))

    def _gather_profile(self):
        """Collect all current UI settings into a serialisable dict."""
        data = {
            'formats': sorted(self._fmt_enabled),
            'viewing': self._fmt_viewing,
            'naming':  dict(self._fmt_naming),
            'folder':  self.export_folder_tb.Text,
            'folder_preset': self._active_folder_preset,
            'subfolders': bool(self.subfolder_cb.IsChecked),
            'open_after': bool(self.open_folder_cb.IsChecked),
            'auto_overwrite': bool(self.auto_overwrite_cb.IsChecked),
            'dest': self._get_print_destination(),
            'sheet_sel': {f: sorted(int(i) for i in c._ids)
                          for f, c in self._sheet_selections.items()},
            'view_sel':  {f: sorted(int(i) for i in c._ids)
                          for f, c in self._view_selections.items()},
        }
        for key, mod in (('pdf', _pdf_settings), ('dwg', _dwg_settings),
                         ('dgn', _dgn_settings), ('nwc', _nwc_settings),
                         ('ifc', _ifc_settings), ('img', _img_settings)):
            try:
                data[key] = mod.read_from_window(self)._asdict()
            except Exception:
                pass
        return data

    def _apply_profile(self, data):
        """Push a profile dict back into the UI."""
        self._loading = True
        try:
            self._fmt_enabled = set(
                f for f in data.get('formats', ['PDF']) if f in ALL_FORMATS)
            viewing = data.get('viewing', 'PDF')
            if viewing in ALL_FORMATS:
                self._fmt_viewing = viewing
                self._show_exp_panel(viewing)
            for f, n in (data.get('naming') or {}).items():
                if f in self._fmt_naming:
                    self._fmt_naming[f] = n

            for key, mod, cls_name in (
                    ('pdf', _pdf_settings, 'PDFSettings'),
                    ('dwg', _dwg_settings, 'DWGSettings'),
                    ('dgn', _dgn_settings, 'DGNSettings'),
                    ('nwc', _nwc_settings, 'NWCSettings'),
                    ('ifc', _ifc_settings, 'IFCSettings'),
                    ('img', _img_settings, 'IMGSettings')):
                try:
                    if mod and key in data:
                        cls = getattr(mod, cls_name)
                        mod.apply_to_window(self, cls(**data[key]))
                except Exception as ex:
                    logger.warning('Profile %s apply failed: %s', key, ex)

            preset_name = data.get('folder_preset')
            preset = (self._load_folder_presets().get(preset_name)
                      if preset_name else None)
            if preset:
                path = fpe_resolve.resolve_path(
                        preset.get('template', ''), preset.get('root', ''),
                        self._project_info, HOST_APP.username, HOST_APP.version)
                self._export_folder = path
                self.export_folder_tb.Text = path
                self._active_folder_preset = preset_name
            else:
                folder = data.get('folder', '')
                if folder and op.isdir(folder):
                    self._export_folder = folder
                    self.export_folder_tb.Text = folder
            self.subfolder_cb.IsChecked      = data.get('subfolders', True)
            self.open_folder_cb.IsChecked    = data.get('open_after', False)
            self.auto_overwrite_cb.IsChecked = data.get('auto_overwrite', False)
            self._set_print_destination(data.get('dest', 'file'))

            # Restore checked sheets/views (ids that still exist apply)
            for key, containers in (('sheet_sel', self._sheet_selections),
                                    ('view_sel',  self._view_selections)):
                saved = data.get(key) or {}
                for f, ids in saved.items():
                    if f in containers:
                        containers[f]._ids = set(int(i) for i in ids)
        finally:
            self._loading = False
        self._update_format_buttons()
        self._update_fmt_column_header()
        self._switch_viewing_format(self._fmt_viewing)
        self._apply_filter()

    def _load_card_into_ui(self, name):
        """Push profile `name`'s saved schedule into the Schedule tab controls.

        Called whenever the schedule dropdown changes, and once at window
        open for the card due next - so an armed schedule is visible again
        instead of the tab coming up blank."""
        block = self._read_profile_schedule(name) or {}
        # Restore rather than clear: _setup_schedule calls this from __init__,
        # where _loading is already True and must stay that way.
        was_loading = self._loading
        self._loading = True
        try:
            self._select_cb_name(self.sched_profile_cb, name)
            # What is on screen is now exactly what is on the profile.
            self._sched_dirty = False

            hour, minute = block.get('hour'), block.get('minute')
            if hour is None or minute is None:
                hour, minute = self._current_time_rounded()
            self._sched_time = (hour, minute)
            self._apply_time_to_wheels(hour, minute)
            self.sched_time_btn.Content = self._time_label(hour, minute)

            self.sched_repeat_cb.SelectedIndex = (
                1 if block.get('repeat_mode') == 'Repeat' else 0)
            days = set(block.get('days') or [])
            for i, n in enumerate(self.SCHED_DAYS):
                getattr(self, n).IsChecked = (i in days)

            raw = block.get('start_date')
            if raw:
                try:
                    import System
                    d = datetime.strptime(raw, '%Y-%m-%d')
                    self.sched_date_dp.SelectedDate = System.DateTime(
                        d.year, d.month, d.day)
                except Exception:
                    pass
        finally:
            self._loading = was_loading
        self._update_sched_gates()
        self._update_sched_status()

    def _profile_path(self, name):
        return op.join(PROFILES_DIR,
                       coreutils.cleanup_filename(name, windows_safe=True)
                       + '.json')

    def profile_changed(self, sender, args):
        # Loading happens in profile_dropdown_closed so re-selecting the
        # same profile also works; SelectionChanged alone can't see that.
        pass

    def profile_dropdown_closed(self, sender, args):
        """Re-selecting the same profile also reloads it (reset to saved)."""
        if self._loading:
            return
        name = self._cb_name(self.profile_cb)
        if name:
            self._load_profile(name)

    def _load_profile(self, name):
        try:
            with open(self._profile_path(name), 'r') as f:
                self._apply_profile(json.load(f))
        except Exception as ex:
            dlg.message('Could not load profile "{}".\n\n{}'.format(name, ex))

    def profile_new_clicked(self, sender, args):
        name, err = '', ''
        while True:
            name = dlg.ask_string('Name for the new profile:',
                                  title='New Profile',
                                  default=name, error=err)
            if not name:
                return
            if op.isfile(self._profile_path(name)):
                err = '"{}" already exists — choose another name.'.format(name)
                continue
            break
        try:
            with open(self._profile_path(name), 'w') as f:
                json.dump(self._gather_profile(), f, indent=2)
            self._setup_profiles(select=name)
        except Exception as ex:
            dlg.message('Could not save profile.\n\n' + str(ex))

    def profile_save_clicked(self, sender, args):
        name = self._cb_name(self.profile_cb)
        if not name:
            self.profile_new_clicked(sender, args)
            return
        try:
            data = self._gather_profile()
            # _gather_profile knows nothing about scheduling, so carry the
            # card's timing across by hand - otherwise saving a profile
            # silently throws away the schedule attached to it.
            block = self._read_profile_schedule(name)
            if block:
                data['schedule'] = block
            with open(self._profile_path(name), 'w') as f:
                json.dump(data, f, indent=2)
            dlg.message('Profile "{}" saved.'.format(name))
        except Exception as ex:
            dlg.message('Could not save profile.\n\n' + str(ex))

    def _list_profile_names(self):
        """All saved profile names — used by the Profiles window."""
        try:
            if not op.isdir(PROFILES_DIR):
                return []
            return [op.splitext(f)[0] for f in os.listdir(PROFILES_DIR)
                    if f.lower().endswith('.json')]
        except Exception:
            return []

    def _delete_profile_by_name(self, name):
        """Delete one profile file — used by the Profiles window."""
        try:
            os.remove(self._profile_path(name))
            # Its schedule died with it; leaving the card armed would have the
            # background scheduler chasing a profile that no longer exists.
            self._disarm_card(name)
        except Exception as ex:
            dlg.message('Could not delete profile.\n\n' + str(ex))

    # ── SCHEDULING  (Revit and the project stay open, this window need not) ──
    # ── Chain: Profile → Time → Repeat (→ Date + weekdays) → Enable ──
    SCHED_DAYS = ['sched_day_mon', 'sched_day_tue', 'sched_day_wed',
                  'sched_day_thu', 'sched_day_fri', 'sched_day_sat',
                  'sched_day_sun']

    WHEEL_ROW_H   = 30
    WHEEL_VISIBLE = 5                       # rows shown in the wheel viewport
    WHEEL_PAD     = WHEEL_VISIBLE // 2

    def _setup_schedule(self):
        """Build the hour / minute / AM-PM scrolling wheels of the time picker."""
        self._tp_sel = {}
        self._wheel = {}
        import System
        self.sched_date_dp.SelectedDate = System.DateTime.Today

        self._build_wheel('h',  self.tp_hours,   [str(h) for h in range(1, 13)])
        self._build_wheel('m',  self.tp_minutes, ['{:02d}'.format(m) for m in range(0, 60, 5)])
        self._build_wheel('ap', self.tp_ampm,    ['AM', 'PM'])

        h24, m = self._current_time_rounded()
        self._sched_time = (h24, m)          # default shown immediately, like the date
        self._apply_time_to_wheels(h24, m)
        self.sched_time_btn.Content = self._time_label(h24, m)

        # Cards armed in an earlier session are still armed - show the one due
        # next, and take ownership of this document's cards while open.
        entries = _schedule.sort_entries(
            [e for e in _read_armed_file().get('entries', []) if e.get('next_run')])
        if entries:
            # Prefer a card armed against this project: a card belonging to
            # another model is still armed, but it is not what you opened
            # this window to look at.
            try:
                here = op.normcase(revit.doc.PathName)
            except Exception:
                here = None
            mine = [e for e in entries
                    if here and op.normcase(e.get('document_path') or '') == here]
            self._load_card_into_ui((mine or entries)[0].get('profile_name'))
            self._start_timer()
        else:
            self._update_sched_gates()
            self._update_sched_status()

    @staticmethod
    def _time_label(h24, minute):
        """'08:05 PM' — the time picker button's caption."""
        h = ((h24 - 1) % 12) + 1
        return '{:02d}:{:02d} {}'.format(h, minute, 'PM' if h24 >= 12 else 'AM')

    @staticmethod
    def _current_time_rounded():
        """Now, rounded up to the next 5-minute mark, as (hour24, minute)."""
        now = datetime.now()
        h24, m = now.hour, now.minute
        m = ((m // 5) + (1 if m % 5 else 0)) * 5
        if m == 60:
            m = 0
            h24 = (h24 + 1) % 24
        return h24, m

    def _apply_time_to_wheels(self, h24, m):
        h = ((h24 - 1) % 12) + 1
        ap = 'PM' if h24 >= 12 else 'AM'
        self._wheel_center('h', h - 1, animate=False)
        self._wheel_center('m', m // 5, animate=False)
        self._wheel_center('ap', 0 if ap == 'AM' else 1, animate=False)

    def _build_wheel(self, group, container, items):
        """Build one scrollable, snapping wheel column inside `container`."""
        rowh = self.WHEEL_ROW_H
        sv = Windows.Controls.ScrollViewer()
        sv.Height = rowh * self.WHEEL_VISIBLE
        sv.Width = 46
        sv.VerticalScrollBarVisibility = Windows.Controls.ScrollBarVisibility.Hidden
        sv.HorizontalScrollBarVisibility = Windows.Controls.ScrollBarVisibility.Disabled
        sv.PanningMode = Windows.Controls.PanningMode.VerticalOnly
        sv.Focusable = False

        panel = Windows.Controls.StackPanel()
        blocks = []
        for _ in range(self.WHEEL_PAD):
            panel.Children.Add(self._wheel_spacer(rowh))
        for text in items:
            tb = Windows.Controls.TextBlock()
            tb.Text = text
            tb.Height = rowh
            tb.FontSize = 15
            tb.HorizontalAlignment = Windows.HorizontalAlignment.Center
            tb.VerticalAlignment = Windows.VerticalAlignment.Center
            tb.Foreground = _brush('#F4FAFF')
            panel.Children.Add(tb)
            blocks.append(tb)
        for _ in range(self.WHEEL_PAD):
            panel.Children.Add(self._wheel_spacer(rowh))

        sv.Content = panel
        container.Children.Add(sv)

        self._wheel[group] = {
            'items': items, 'blocks': blocks, 'sv': sv, 'rowh': rowh,
            'dragging': False, 'start_y': 0.0, 'start_off': 0.0,
        }
        sv.PreviewMouseWheel          += lambda s, a: self._wheel_wheel(group, a)
        sv.PreviewMouseLeftButtonDown += lambda s, a: self._wheel_down(group, a)
        sv.PreviewMouseMove           += lambda s, a: self._wheel_move(group, a)
        sv.PreviewMouseLeftButtonUp   += lambda s, a: self._wheel_up(group, a)

    def _wheel_spacer(self, rowh):
        sp = Windows.Controls.TextBlock()
        sp.Height = rowh
        return sp

    def _wheel_wheel(self, group, args):
        idx = self._wheel_index(group) + (-1 if args.Delta > 0 else 1)
        st = self._wheel[group]
        self._wheel_center(group, max(0, min(len(st['items']) - 1, idx)))
        args.Handled = True

    def _wheel_down(self, group, args):
        st = self._wheel[group]
        st['dragging'] = True
        st['start_y'] = args.GetPosition(st['sv']).Y
        st['start_off'] = st['sv'].VerticalOffset
        st['sv'].CaptureMouse()

    def _wheel_move(self, group, args):
        st = self._wheel[group]
        if not st['dragging']:
            return
        y = args.GetPosition(st['sv']).Y
        offset = st['start_off'] - (y - st['start_y'])
        offset = max(0.0, min((len(st['items']) - 1) * st['rowh'], offset))
        st['sv'].ScrollToVerticalOffset(offset)
        self._wheel_fade(group)

    def _wheel_up(self, group, args):
        st = self._wheel[group]
        if not st['dragging']:
            return
        st['dragging'] = False
        st['sv'].ReleaseMouseCapture()
        self._wheel_center(group, self._wheel_index(group))

    def _wheel_index(self, group):
        st = self._wheel[group]
        idx = int(round(st['sv'].VerticalOffset / float(st['rowh'])))
        return max(0, min(len(st['items']) - 1, idx))

    def _wheel_center(self, group, idx, animate=True):
        """Snap the wheel so item `idx` sits centred, and record the selection."""
        st = self._wheel[group]
        idx = max(0, min(len(st['items']) - 1, idx))
        value = st['items'][idx]
        self._tp_sel[group] = int(value) if group != 'ap' else value
        target = idx * st['rowh']
        if not animate:
            st['sv'].ScrollToVerticalOffset(target)
            self._wheel_fade(group)
            return
        self._wheel_animate(group, target)

    def _wheel_animate(self, group, target, steps=6):
        """Ease the wheel toward `target` over a few timer ticks."""
        st = self._wheel[group]
        start = st['sv'].VerticalOffset
        state = {'n': 0}
        timer = Windows.Threading.DispatcherTimer()
        timer.Interval = framework.System.TimeSpan.FromMilliseconds(15)

        def tick(sender, args):
            state['n'] += 1
            t = state['n'] / float(steps)
            if t >= 1.0:
                st['sv'].ScrollToVerticalOffset(target)
                self._wheel_fade(group)
                timer.Stop()
                return
            eased = 1 - (1 - t) ** 2
            st['sv'].ScrollToVerticalOffset(start + (target - start) * eased)
            self._wheel_fade(group)

        timer.Tick += tick
        timer.Start()

    def _wheel_fade(self, group):
        """Fade rows based on distance from the centred (selected) row."""
        st = self._wheel[group]
        pos = st['sv'].VerticalOffset / float(st['rowh'])
        for i, tb in enumerate(st['blocks']):
            tb.Opacity = max(0.18, 1.0 - abs(i - pos) * 0.4)

    def tp_popup_opened(self, sender, args):
        """Re-centre the wheels on the last picked/default time each time it opens."""
        h24, m = self._sched_time or self._current_time_rounded()
        self._apply_time_to_wheels(h24, m)

    def tp_ok_clicked(self, sender, args):
        h, m, ap = self._tp_sel['h'], self._tp_sel['m'], self._tp_sel['ap']
        h24 = (h % 12) + (12 if ap == 'PM' else 0)
        self._sched_time = (h24, m)
        self.sched_time_btn.Content = self._time_label(h24, m)
        self.sched_time_btn.IsChecked = False   # close popup
        self.sched_field_changed(sender, args)

    def tp_cancel_clicked(self, sender, args):
        self.sched_time_btn.IsChecked = False

    def sched_profile_changed(self, sender, args):
        """Schedule dropdown changed — load that card, wholesale.

        Picking a card here loads the profile itself as well as its timing,
        so the window shows exactly what that card will print: its sheets,
        its views, its export settings. The header dropdown is moved to
        match, since the two are now showing the same profile."""
        if self._loading:
            return
        name = self._cb_name(self.sched_profile_cb)
        if not name:
            self._update_sched_gates()
            return
        self._load_profile(name)
        was_loading = self._loading
        self._loading = True
        try:
            self._select_cb_name(self.profile_cb, name)
        finally:
            self._loading = was_loading
        self._load_card_into_ui(name)

    def sched_action_clicked(self, sender, args):
        """The Schedule button — the one place a card is ever committed.

        It is per card: it arms or drops whichever profile the schedule
        dropdown is showing, and leaves every other card alone. Pressed on a
        card that is armed but has edits pending it re-arms rather than
        drops, which is what its label says by then."""
        if self._loading:
            return
        name = self._cb_name(self.sched_profile_cb)
        if not name:
            self.sched_status_tb.Text = 'Pick a profile…'
            return
        if self._card_armed(name) and not self._sched_dirty:
            self._disarm_card(name)
        else:
            self._arm_card(name)

    def sched_field_changed(self, sender, args):
        """Time / repeat / weekday / start date edited.

        Nothing is written here — not to the profile, not to the run list.
        The edit stays pending until the Schedule button commits it, so an
        armed card cannot quietly move to a new time under you while you are
        still deciding what that time should be."""
        if self._loading:
            return
        self._sched_dirty = True
        self._update_sched_gates()
        self._update_sched_status()

    def _sched_repeat_mode(self):
        try:
            item = self.sched_repeat_cb.SelectedItem
            return item.Content if item else 'Once'
        except Exception:
            return 'Once'

    def _parse_sched_time(self):
        return self._sched_time

    def _sched_checked_days(self):
        """Checked weekdays as python weekday() ints (Mon=0)."""
        return [i for i, n in enumerate(self.SCHED_DAYS)
                if getattr(self, n).IsChecked]

    def _sched_start_date(self):
        d = self.sched_date_dp.SelectedDate
        if d is None:
            return datetime.now().date()
        return datetime(d.Year, d.Month, d.Day).date()

    def _update_sched_gates(self):
        """Unlock each control only when the previous step is set.

        Chain: pick a profile, and its timing opens up. The profile has to
        come first because it decides which card everything else is
        editing."""
        prof   = bool(self._cb_name(self.sched_profile_cb))
        repeat = prof and self._sched_repeat_mode() == 'Repeat'

        self.sched_profile_cb.IsEnabled = True
        self.sched_time_btn.IsEnabled   = prof
        self.sched_repeat_cb.IsEnabled  = prof
        self.sched_date_dp.IsEnabled    = repeat
        self.sched_days_panel.Visibility = (
            Windows.Visibility.Visible if repeat
            else Windows.Visibility.Collapsed)
        for ctrl in (self.sched_time_btn, self.sched_repeat_cb):
            ctrl.Opacity = 1.0 if prof else 0.55
        self._update_sched_button()

    def _card_armed(self, name):
        """True when profile `name` is in the run list right now."""
        return bool(name) and name in self._armed_profile_names()

    def _update_sched_button(self):
        """Label and colour of the Schedule button.

        It always names what pressing it will do, so an armed card with edits
        pending reads Schedule again rather than Unschedule - the press
        applies the new timing instead of dropping the card."""
        name  = self._cb_name(self.sched_profile_cb)
        # Armed and untouched is the only state where the press cancels.
        drop  = self._card_armed(name) and not self._sched_dirty
        self.sched_action_btn.Content   = 'Unschedule' if drop else 'Schedule'
        self.sched_action_btn.IsEnabled = bool(name)
        self.sched_action_btn.ToolTip = (
            'Drop this card from the run list — its time stays on the profile'
            if drop else
            'Arm this profile — it will print at the time set here, with '
            'pySheets closed')
        style = self.TryFindResource('SmallSecBtn' if drop else 'SmallPrimBtn')
        if style:
            self.sched_action_btn.Style = style

    def _sched_block_from_ui(self):
        """The Schedule tab's current fields as a profile schedule block.

        No `enabled` key: arming is what sets it, in _arm_card."""
        t = self._parse_sched_time()
        if t is None:
            return None
        return {
            'hour':        t[0],
            'minute':      t[1],
            'repeat_mode': self._sched_repeat_mode(),
            'days':        self._sched_checked_days(),
            'start_date':  self._sched_start_date().strftime('%Y-%m-%d'),
        }

    # ── Arming ──
    def _arm_card(self, name):
        """Add (or move) profile `name` in the run list, at the time now
        showing in the Schedule tab."""
        block = self._sched_block_from_ui()
        if block is None:
            self.sched_status_tb.Text = 'Pick a time…'
            return
        block['enabled'] = True
        nxt = compute_next_run(block)
        self._write_profile_schedule(name, block)
        if nxt is None:
            # Only reachable on Repeat with no weekday ticked. Drop the card
            # rather than leave it armed - it still holds a next_run from
            # before the days were cleared, and would fire on it.
            self._disarm_card(name)
            self.sched_status_tb.Text = 'Pick at least one day…'
            return

        # An unsaved project has no path to match on later, so there would be
        # nothing for the background scheduler to find when the time came.
        doc_path = ''
        try:
            doc_path = revit.doc.PathName or ''
        except Exception:
            pass
        if not doc_path:
            self._disarm_card(name)
            self.sched_status_tb.Text = 'Save the project first'
            return

        # The card is bound to the host project, not to whatever is picked in
        # the documents dropdown - a linked document cannot be activated on
        # its own, and it is the host that has to be left open anyway.
        data = _read_armed_file()
        entries = [e for e in data.get('entries', [])
                   if e.get('profile_name') != name]
        entries.append({
            'profile_name':   name,
            # Stored so the background scheduler can read this card's timing
            # without having to guess how the name maps to a filename.
            'profile_path':   self._profile_path(name),
            'document_path':  doc_path,
            'document_title': revit.doc.Title,
            'next_run':       nxt.strftime(TS_FMT),
        })
        data['entries'] = entries
        _write_armed_file(data)

        self._sched_dirty = False    # committed - the button goes to Unschedule
        self._start_timer()
        self._setup_profiles()      # repaint the armed dots
        self._update_sched_gates()
        self._update_sched_status()

    def _disarm_card(self, name):
        """Drop profile `name` from the run list. Its timing stays on the
        profile, so re-arming it later brings the same time back."""
        data = _read_armed_file()
        data['entries'] = [e for e in data.get('entries', [])
                           if e.get('profile_name') != name]
        _write_armed_file(data)

        block = self._read_profile_schedule(name)
        if block:
            block['enabled'] = False
            self._write_profile_schedule(name, block)

        if not data['entries']:
            self._stop_timer()
        self._sched_dirty = False
        self._setup_profiles()
        self._update_sched_gates()
        self._update_sched_status()

    def _update_sched_status(self):
        """Green status line: what is armed, and when the next one fires.

        Pending edits are called out here, because the armed time it is
        showing is the old one until the Schedule button is pressed."""
        entries = [e for e in _read_armed_file().get('entries', [])
                   if e.get('next_run')]
        if not entries:
            self.sched_status_tb.Text = 'Schedule off'
            return
        hint = ''
        if self._sched_dirty and self._card_armed(
                self._cb_name(self.sched_profile_cb)):
            hint = ' — press Schedule to apply changes'
        entries.sort(key=lambda e: (e['next_run'],
                                    (e.get('profile_name') or '').lower()))
        first = entries[0]
        try:
            when = datetime.strptime(first['next_run'], TS_FMT).strftime(
                '%a %d %b %H:%M')
        except Exception:
            when = first['next_run']
        if len(entries) == 1:
            self.sched_status_tb.Text = 'Next run: {}{}'.format(when, hint)
        else:
            self.sched_status_tb.Text = '{} armed — next {} ({}){}'.format(
                len(entries), when, first.get('profile_name'), hint)

    def _start_timer(self):
        """One timer for the window, however many cards are armed."""
        if self._sched_timer:
            self._beat()
            return
        self._sched_timer = Windows.Threading.DispatcherTimer()
        self._sched_timer.Interval = framework.System.TimeSpan.FromSeconds(20)
        self._sched_timer.Tick += self._sched_tick
        self._sched_timer.Start()
        self._beat()

    def _beat(self):
        """Claim this document's cards while the window is open.

        startup.py skips any card whose document matches a fresh heartbeat,
        so the window and the background handler can never both fire one.
        Cards armed against a different project stay with the background
        handler, since this window cannot switch documents."""
        try:
            data = _read_armed_file()
            data['heartbeat']     = datetime.now().strftime(TS_FMT)
            data['heartbeat_doc'] = revit.doc.PathName
            _write_armed_file(data)
        except Exception:
            pass

    def _release_heartbeat(self):
        """Hand this document's cards back to the background handler."""
        try:
            data = _read_armed_file()
            data['heartbeat']     = None
            data['heartbeat_doc'] = None
            _write_armed_file(data)
        except Exception:
            pass

    def _stop_timer(self):
        if self._sched_timer:
            try:
                self._sched_timer.Stop()
            except Exception:
                pass
        self._sched_timer = None

    # ── Running due cards ──
    def _sched_tick(self, sender, args):
        """Every 20s while the window is open: renew the claim on this
        document's cards, and run any that have come due."""
        self._beat()
        # _do_export pumps the dispatcher, which lets this timer tick again
        # mid-export. Without this guard a long export starts itself twice.
        if self._running_sched:
            return
        entries = self._due_entries()
        if entries:
            self._run_due_cards(entries)

    def _due_entries(self):
        """Armed cards for this document whose time has come.

        respect_heartbeat is off because the heartbeat is this window's own
        claim - honouring it here would mean never running anything."""
        try:
            doc_path = revit.doc.PathName
        except Exception:
            return []
        return _schedule.due_entries(_read_armed_file(), doc_path=doc_path,
                                     respect_heartbeat=False)

    def _run_due_cards(self, entries):
        """Run each due card in turn — one export at a time, never two at once.

        Ordered by due time, then profile name (the order the dropdowns list
        them in). A card that fails is logged and the batch carries on, so one
        bad profile cannot stop the other four."""
        data    = _read_armed_file()
        entries = _schedule.sort_entries(entries)
        failed, missed, ran = [], [], []

        # Show the Export tab: a scheduled run should come up on the queue and
        # progress bar, not on whatever tab the window happens to open on.
        try:
            self._show_tab(TAB_PRINT)
        except Exception:
            pass

        self._running_sched = True
        was_unattended = self._unattended
        self._unattended = True
        try:
            for entry in entries:
                name = entry.get('profile_name')
                if not name or not op.isfile(self._profile_path(name)):
                    logger.warning('Scheduled card "%s" has no profile, dropped', name)
                    self._disarm_card(name)
                    continue
                if _schedule.is_stale(data, entry):
                    when = _parse_ts(entry.get('next_run'))
                    logger.warning('Scheduled export "%s" missed its slot (%s) '
                                   'by more than the grace window — skipped',
                                   name,
                                   when.strftime('%d %b %H:%M') if when else '?')
                    missed.append(name)
                else:
                    try:
                        self._run_one_card(name)
                        ran.append(name)
                    except Exception as ex:
                        logger.error('Scheduled export "%s" failed: %s', name, ex)
                        failed.append('{}: {}'.format(name, ex))
                self._advance_card(name)
        finally:
            self._unattended = was_unattended
            self._running_sched = False

        self._setup_profiles()
        self._update_sched_status()
        sel = self._cb_name(self.sched_profile_cb)
        if sel:
            self._load_card_into_ui(sel)
        self._report_sched_run(ran, missed, failed)

    def _run_one_card(self, name):
        """Load one profile and export it, exactly as pressing Print would."""
        with open(self._profile_path(name), 'r') as f:
            self._apply_profile(json.load(f))
        self._build_queue()
        self._do_export()

    def _advance_card(self, name):
        """Re-arm a repeating card for its next occurrence; retire a spent one."""
        block = self._read_profile_schedule(name)
        if block and block.get('repeat_mode') == 'Repeat':
            nxt = compute_next_run(block)
            if nxt:
                data = _read_armed_file()
                for entry in data.get('entries', []):
                    if entry.get('profile_name') == name:
                        entry['next_run'] = nxt.strftime(TS_FMT)
                _write_armed_file(data)
                return
        self._disarm_card(name)

    def _report_sched_run(self, ran, missed, failed):
        """One summary at the end of the batch, never one dialog per card."""
        if not (ran or missed or failed):
            return
        parts = []
        if ran:
            parts.append('Exported: {}'.format(', '.join(ran)))
        if missed:
            parts.append('Skipped (too late): {}'.format(', '.join(missed)))
        if failed:
            parts.append('Failed:\n  {}'.format('\n  '.join(failed)))
        summary = '\n'.join(parts)
        logger.info('Scheduled run finished. %s', summary.replace('\n', ' '))
        if failed:
            dlg.message('Scheduled export finished with errors.\n\n' + summary,
                        title='pySheets')

    # ── XAML EVENT HANDLERS WINDOW ──
    LAST_SESSION = op.join(USERDATA_DIR, 'settings', 'lastsession.json')

    # ── Column order persistence ──
    def _column_keys(self):
        """Ordered (key, column) pairs for the sheets grid columns whose
        DisplayIndex should persist across sessions."""
        return [
            ('check',      self.check_col),
            ('number',     self.number_col),
            ('name',       self.name_col),
            ('revision',   self.revision_col),
            ('size',       self.size_col),
            ('collection', self.collection_col),
            ('filename',   self.filename_col),
        ]

    def _get_column_order(self):
        try:
            pairs = self._column_keys()
            pairs.sort(key=lambda kv: kv[1].DisplayIndex)
            return [k for k, _ in pairs]
        except Exception as ex:
            logger.warning('Column order read failed: %s', ex)
            return None

    def _set_column_order(self, order):
        if not order:
            return
        try:
            by_key = dict(self._column_keys())
            idx = 0
            for key in order:
                col = by_key.get(key)
                if col is not None:
                    col.DisplayIndex = idx
                    idx += 1
            # Columns not present in a saved older session (e.g. newly
            # added 'collection') just get appended at the end.
            for key, col in self._column_keys():
                if key not in order:
                    col.DisplayIndex = idx
                    idx += 1
        except Exception as ex:
            logger.warning('Column order restore failed: %s', ex)

    def _get_column_widths(self):
        """Pixel widths for user-resized columns. 'name' is excluded — it's
        star-sized (Width="*") and should keep filling remaining space
        rather than being pinned to a saved pixel value."""
        try:
            widths = {}
            for key, col in self._column_keys():
                if key == 'name':
                    continue
                try:
                    if col.ActualWidth > 0:
                        widths[key] = col.ActualWidth
                except Exception:
                    pass
            return widths
        except Exception as ex:
            logger.warning('Column width read failed: %s', ex)
            return None

    def _set_column_widths(self, widths):
        if not widths:
            return
        try:
            by_key = dict(self._column_keys())
            for key, w in widths.items():
                col = by_key.get(key)
                if col is not None and w:
                    col.Width = Windows.Controls.DataGridLength(float(w))
        except Exception as ex:
            logger.warning('Column width restore failed: %s', ex)

    def _save_last_session(self):
        try:
            session_dir = op.dirname(self.LAST_SESSION)
            if not op.isdir(session_dir):
                os.makedirs(session_dir)
            data = self._gather_profile()
            data.pop('sheet_sel', None)   # selections are per-session
            data.pop('view_sel', None)
            data['column_order']  = self._get_column_order()
            data['column_widths'] = self._get_column_widths()
            with open(self.LAST_SESSION, 'w') as f:
                json.dump(data, f, indent=2)
        except Exception as ex:
            logger.warning('Last session save failed: %s', ex)

    def _load_last_session(self):
        try:
            if op.isfile(self.LAST_SESSION):
                with open(self.LAST_SESSION, 'r') as f:
                    data = json.load(f)
                self._apply_profile(data)
                self._set_column_order(data.get('column_order'))
                self._set_column_widths(data.get('column_widths'))
        except Exception as ex:
            logger.warning('Last session load failed: %s', ex)

    def window_closing(self, sender, args):
        self._save_last_session()
        self._save_naming_memory()
        # Only the in-window timer stops here. The armed cards stay armed -
        # startup.py picks them up once the heartbeat is released, which is
        # the whole point of scheduling from a window you then close.
        self._stop_timer()
        self._release_heartbeat()
        self._restore_print_settings()

    def close_window(self, sender, args):
        self.Close()

    def win_close_clicked(self, sender, args):
        self.Close()


# ── HELPERS  (module-level, no access to self) ──
def _brush(color_hex):
    """Return a SolidColorBrush from a hex colour string."""
    try:
        color = Windows.Media.ColorConverter.ConvertFromString(color_hex)
        b = Windows.Media.SolidColorBrush(color)
        b.Freeze()
        return b
    except Exception:
        return Windows.Media.Brushes.Transparent


def _set_child_text(border, text, foreground_brush):
    """Set the text and foreground of the first TextBlock child of a Border."""
    try:
        tb = border.Child
        if hasattr(tb, 'Text'):
            tb.Text       = text
            tb.Foreground = foreground_brush
    except Exception:
        pass


# ── ENTRY POINT ──
def launch_scheduled(entries):
    """Called by startup.py's Idling-based scheduler when armed cards come
    due and this window isn't already open. Opens the window normally
    (steals focus like any manual launch), runs the given cards one at a
    time, and leaves it open.

    Every entry must belong to the document that is active right now -
    startup.py activates it first and groups the batch by document, because
    a window is built around whichever project was active when it opened
    and cannot be pointed at another one afterwards."""
    win = PrintSheetsWindow('pySheets.xaml')
    win.Show()   # non-modal: returns immediately so the export can follow
    try:
        win._run_due_cards(entries)
    except Exception as ex:
        logger.error('Scheduled run failed: %s', ex)
        dlg.message('Scheduled export failed.\n\n' + str(ex), title='pySheets')
    return win


if __name__ == '__main__':
    try:
        PrintSheetsWindow('pySheets.xaml').show_dialog()
    except Exception as e:
        import traceback
        logger.error(traceback.format_exc())
        dlg.message('Error starting pySheets\n\n' + str(e))
