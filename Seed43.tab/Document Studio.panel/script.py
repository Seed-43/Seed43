# -*- coding: utf-8 -*-
__title__     = "pySheets"
__author__    = "Seed43"
__version__   = "2.0"
__doc__ = """Version = 2.0
__highlight__ = ""
Date    = 2026-03-24
_____________________________________________________________________
Description:
Print / export Revit sheets and views to PDF and other formats.
Three-tab workflow: Select → Export Settings → Print.
_____________________________________________________________________
"""

# ═══════════════════════════════════════════════════════════════════
# IMPORTS
# ═══════════════════════════════════════════════════════════════════
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

from EditNamingFormats import EditNamingFormatsWindow, NamingFormat
import seed43dialogs as dlg
import FolderPresetManager as fpm_win
import folder_preset_resolve as fpe_resolve


# Per-format export settings modules (from settings/ subfolder)
import sys as _sys
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


# ═══════════════════════════════════════════════════════════════════
# CONSTANTS
# ═══════════════════════════════════════════════════════════════════
# Format pill names — must match XAML Button x:Name suffixes
ALL_FORMATS   = ['PDF', 'DWG', 'DGN', 'NWC', 'IFC', 'IMG']
FILE_EXT      = {'PDF': '.pdf', 'DWG': '.dwg', 'DGN': '.dgn',
                 'NWC': '.nwc', 'IFC': '.ifc', 'IMG': '.png'}

# Tab indices — tabs renamed: Select | Settings | Export
TAB_SELECT   = 0
TAB_EXPORT   = 1   # "Settings" tab
TAB_PRINT    = 2   # "Export" tab

# Subfolder for per-format export settings scripts
SETTINGS_FOLDER = op.join(op.dirname(__file__), 'settings')

# All user data (profiles, naming formats) lives next to the script
USERDATA_DIR = op.join(op.dirname(__file__), 'userdata')
PROFILES_DIR = op.join(USERDATA_DIR, 'profiles')
NAMING_FILE  = op.join(USERDATA_DIR, 'naming.json')

DEFAULT_SETUP = '<Revit Default>'

IMG_TYPE_EXT = {'PNG': '.png', 'JPEG': '.jpg', 'TIFF': '.tif'}


# ═══════════════════════════════════════════════════════════════════
# NAMED TUPLES — lightweight immutable data bags
# ═══════════════════════════════════════════════════════════════════
AvailableDoc = namedtuple('AvailableDoc', ['name', 'hash', 'linked'])

SheetRevision = namedtuple('SheetRevision',
                           ['number', 'desc', 'date', 'is_set'])
UNSET_REVISION = SheetRevision(number='', desc='', date='', is_set=False)

TitleBlockPrintSettings = namedtuple('TitleBlockPrintSettings',
                                     ['psettings', 'set_by_param'])


# ═══════════════════════════════════════════════════════════════════
# SELECTION CONTAINER
# Keeps checkbox state independent of filtering / sorting / tab switches
# ═══════════════════════════════════════════════════════════════════
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


# ═══════════════════════════════════════════════════════════════════
# SHEET LIST ITEM
# One row in the Select tab DataGrid
# ═══════════════════════════════════════════════════════════════════
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

        # Reactive state
        self._is_selected    = False
        self._print_filename = ''

    # ── Helpers ────────────────────────────────────────────────────
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

    # ── Read-only Revit accessors ───────────────────────────────────
    @property
    def revit_sheet(self):
        return self._sheet

    @property
    def revit_tblock(self):
        return self._tblock

    @property
    def revit_tblock_type(self):
        return self._tblock_type

    # ── Reactive properties ─────────────────────────────────────────
    @forms.reactive
    def is_selected(self):
        return self._is_selected

    @is_selected.setter
    def is_selected(self, value):
        self._is_selected = value
        if self._container:
            self._container.update(get_elementid_value(self._sheet.Id), value)

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



# ═══════════════════════════════════════════════════════════════════
# VIEW LIST ITEM
# One row in the Select tab DataGrid when in Views mode
# Reuses SheetItem fields so the same DataGrid columns work
# ═══════════════════════════════════════════════════════════════════
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
        self.all_print_settings = []
        self.read_only          = False

        self._is_selected    = False
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


# ═══════════════════════════════════════════════════════════════════
# SHEET SET WRAPPERS
# ═══════════════════════════════════════════════════════════════════
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


# ═══════════════════════════════════════════════════════════════════
# PRINT QUEUE ITEM  (used in Print tab DataGrid)
# ═══════════════════════════════════════════════════════════════════
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


# ═══════════════════════════════════════════════════════════════════
# MAIN WINDOW
# ═══════════════════════════════════════════════════════════════════
class PrintSheetsWindow(forms.WPFWindow):

    def __init__(self, xaml_file_name):
        self._loading = True   # suppress event handlers during init
        forms.WPFWindow.__init__(self, xaml_file_name)

        # ── Core state ──────────────────────────────────────────────
        self._current_tab     = TAB_SELECT
        self._export_folder   = USER_DESKTOP
        self._init_psettings  = None     # to restore on close

        # ── Per-format selection containers ─────────────────────────
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
        self._highlighted     = set()    # row indices highlighted via shift/ctrl
        self._last_row_index  = -1       # for shift-range anchor
        self._sv_mode         = 'sheets' # 'sheets' | 'views'
        self._type_filter     = None     # None = all view types visible

        # ── Schedule state ──────────────────────────────────────────
        self._sched_timer = None
        self._sched_next  = None

        # ── Format state ────────────────────────────────────────────
        # enabled: set of format strings the user has toggled on
        # viewing: which format's filename column is showing
        self._fmt_enabled = set()
        self._fmt_enabled.add('PDF')     # PDF on by default
        self._fmt_viewing = 'PDF'

        # Per-format naming format selection (name string)
        self._fmt_naming  = {f: None for f in ALL_FORMATS}

        # ── Project info ────────────────────────────────────────────
        self._project_info = revit.query.get_project_info(doc=revit.doc)
        self._active_folder_preset = None

        # ── Populate UI ─────────────────────────────────────────────
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

        # ── Extra event wiring ──────────────────────────────────────
        self.sheets_dg.PreviewMouseLeftButtonDown += self._row_click
        self.PreviewMouseLeftButtonDown           += self._window_click
        self._setup_scroll_prevention()

        # ── Initialise tab/step state ────────────────────────────────
        # Must run after all controls are wired so step indicators
        # and tab button Tags are set correctly from the start
        self._show_tab(TAB_SELECT)
        self._load_last_session()
        self._rebuild_folder_presets()
        self._update_dest_gates()
        self._loading = False   # init complete — events now active

    # ═══════════════════════════════════════════════════════════════
    # PROPERTIES
    # ═══════════════════════════════════════════════════════════════
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
        src = self.sheets_dg.ItemsSource
        return list(src) if src else []

    @property
    def _checked_sheets(self):
        return [s for s in self._visible_sheets if s.is_selected]

    # ═══════════════════════════════════════════════════════════════
    # SETUP METHODS
    # ═══════════════════════════════════════════════════════════════
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
        # Combined-PDF naming dropdown
        try:
            self.combined_naming_cb.ItemsSource = fmts
            want = saved.get('PDF_COMBINED')
            match = next((x for x in fmts if x.name == want), None)
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
            if not op.isdir(USERDATA_DIR):
                os.makedirs(USERDATA_DIR)
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

        # Mark compatibility
        pm = self._get_printmanager()
        if pm:
            compat = {p.Name for p in pm.PaperSizes if p}
            for item in items:
                if (not item.allows_variable_paper and
                        item.paper_size and item.paper_size.Name in compat):
                    item.is_compatible = True

        self.printsetting_cb.ItemsSource = items
        if items:
            # Try to match current Revit setting
            pm = self._get_printmanager()
            if pm:
                cur = pm.PrintSetup.CurrentPrintSetting
                if isinstance(cur, DB.InSessionPrintSetting):
                    self.printsetting_cb.SelectedIndex = 0
                else:
                    self._init_psettings = cur
                    for i, item in enumerate(items):
                        if item.name == cur.Name:
                            self.printsetting_cb.SelectedIndex = i
                            break
            else:
                self.printsetting_cb.SelectedIndex = 0

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

    # ═══════════════════════════════════════════════════════════════
    # SHEET LIST
    # ═══════════════════════════════════════════════════════════════
    def _reload_sheet_list(self):
        """Rebuild _all_sheets/_ordered_sheets from the selected doc + mode."""
        doc  = self._selected_doc
        sset = self._selected_sheetset

        if self._sv_mode == 'views':
            self._reload_view_list(doc)
            return

        if sset is None:
            return

        tblocks = list(revit.query.get_elements_by_categories(
            [DB.BuiltInCategory.OST_TitleBlocks], doc=doc))
        rev_cfg = DB.RevisionSettings.GetRevisionSettings(doc)

        all_ps = revit.query.get_all_print_settings(doc=doc)
        sheet_ps = {}
        if (self._selected_printsetting is not None and
                self._selected_printsetting.allows_variable_paper):
            sheet_ps = self._build_sheet_ps_map(tblocks, all_ps)

        sheets = []
        for rs in sset.get_sheets(doc):
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

        sheets.sort(key=lambda s: s.number)
        self._all_sheets      = sheets
        self._ordered_sheets  = list(sheets)
        self._all_sheet_items = list(sheets)  # permanent cache for cross-mode printing
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

        items.sort(key=lambda v: (str(v.paper_size), v.name))
        self._all_sheets     = items
        self._ordered_sheets = list(items)
        self._all_view_items = list(items)  # permanent cache for cross-mode printing
        self._rebuild_type_filter()
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
            labels = ['All types'] + sorted(
                {v.paper_size for v in self._all_view_items})
            cb.ItemsSource   = labels
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

        # Filter by placement state — views mode only
        state = getattr(self, '_state_filter', None)
        if self._sv_mode == 'views' and state:
            if state == 'Placed views':
                sheets = [s for s in sheets if getattr(s, 'placed_on', '')]
            else:
                sheets = [s for s in sheets if not getattr(s, 'placed_on', '')]

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
        """Replace project/global params — same for every sheet."""
        doc = self._selected_doc
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
        """File extension for a format — IMG depends on the image type."""
        if fmt == 'IMG':
            try:
                itype = self.img_type_cb.SelectedItem.Content
                return IMG_TYPE_EXT.get(itype, '.png')
            except Exception:
                return '.png'
        return FILE_EXT.get(fmt, '')

    # ═══════════════════════════════════════════════════════════════
    # FORMAT PILL LOGIC
    # ═══════════════════════════════════════════════════════════════
    def _fmt_btn(self, fmt):
        """Return the Button element for a format string."""
        name_map = {
            'PDF': self.fmt_pdf_btn,
            'DWG': self.fmt_dwg_btn,
            'DGN': self.fmt_dgn_btn,
            'NWC': self.fmt_nwc_btn,
            'IFC': self.fmt_ifc_btn,
            'IMG': self.fmt_img_btn,
        }
        return name_map.get(fmt)

    def _update_format_buttons(self):
        """Sync all format pill Tags to current _fmt_enabled / _fmt_viewing."""
        for fmt in ALL_FORMATS:
            btn = self._fmt_btn(fmt)
            if btn is None:
                continue
            if fmt == self._fmt_viewing and fmt in self._fmt_enabled:
                btn.Tag = 'Viewing'
            elif fmt in self._fmt_enabled:
                btn.Tag = 'Enabled'
            else:
                btn.Tag = ''

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
        """Switch which format's column is shown and swap selection container."""
        self._fmt_viewing = fmt
        # Point _selection to the right container for current mode + format
        containers = self._view_selections if self._sv_mode == 'views' else self._sheet_selections
        self._selection = containers[fmt]
        # Re-attach the new container to all existing items and restore their state
        for s in self._all_sheets:
            s._container = self._selection
            sheet_id = get_elementid_value(s.revit_sheet.Id)
            s._is_selected = self._selection.contains(sheet_id)

    # ═══════════════════════════════════════════════════════════════
    # TAB NAVIGATION
    # ═══════════════════════════════════════════════════════════════
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

        # Print button in header — only visible on Export tab (TAB_PRINT)
        try:
            self.header_print_btn.Visibility = (
                Windows.Visibility.Visible
                if tab_index == TAB_PRINT
                else Windows.Visibility.Collapsed)
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
        """Green fill for done/active, grey border for future steps."""
        GREEN  = '#208A3C'
        GREY   = '#555555'
        WHITE  = 'White'
        MUTED  = '#888888'

        circles  = [self.step1_circle, self.step2_circle, self.step3_circle]
        lines    = [self.step_line_1,  self.step_line_2]

        for i, circle in enumerate(circles):
            if i < current:       # done
                circle.Background      = _brush(GREEN)
                circle.BorderBrush     = _brush(GREEN)
                circle.BorderThickness = Windows.Thickness(0)
                _set_child_text(circle, '✓', WHITE)
            elif i == current:    # active
                circle.Background      = _brush(GREEN)
                circle.BorderBrush     = _brush(GREEN)
                circle.BorderThickness = Windows.Thickness(0)
                _set_child_text(circle, str(i + 1), WHITE)
            else:                 # future
                circle.Background      = _brush('Transparent')
                circle.BorderBrush     = _brush(GREY)
                circle.BorderThickness = Windows.Thickness(2)
                _set_child_text(circle, str(i + 1), MUTED)

        for i, line in enumerate(lines):
            line.Fill = _brush(GREEN if i < current else GREY)

    # ═══════════════════════════════════════════════════════════════
    # SELECTION HELPERS
    # ═══════════════════════════════════════════════════════════════
    def _update_sel_count(self):
        total   = len(self._ordered_sheets)
        checked = self._selection.count()
        self.sel_count_tb.Text = '{} of {} selected'.format(checked, total)
        fmt_label = self._fmt_viewing if self._fmt_viewing in self._fmt_enabled else 'none'
        item_word = 'view' if self._sv_mode == 'views' else 'sheet'
        self.footer_status_tb.Text = (
            '{} {} {}{}'
            .format(checked, fmt_label, item_word, 's' if checked != 1 else '')
        )

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

    # ═══════════════════════════════════════════════════════════════
    # PRINT TAB — QUEUE
    # ═══════════════════════════════════════════════════════════════
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
                items.append(QueueItem(
                    s.number, fname + self._ext_for(fmt), fmt,
                    s.paper_size, source=s, fname_noext=fname))

        self.queue_dg.ItemsSource = ObjectModel.ObservableCollection[object](items)
        self.overall_progress.Value = 0
        self.overall_pct_tb.Text = 'Ready'
    # ═══════════════════════════════════════════════════════════════
    # EXPORT ENGINE
    # ═══════════════════════════════════════════════════════════════
    def _do_export(self):
        """Run the export queue across all enabled formats."""
        if not self._fmt_enabled:
            dlg.message('No export formats selected.')
            return
        base_folder = self.export_folder_tb.Text
        if not op.isdir(base_folder):
            dlg.message('Export folder does not exist.')
            return

        queue = list(self.queue_dg.ItemsSource or [])
        if not queue:
            self._build_queue()
            queue = list(self.queue_dg.ItemsSource or [])
        if not queue:
            dlg.message('No sheets or views selected.')
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
                     'IFC': self._export_ifc, 'IMG': self._export_img}

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
            try:
                coreutils.open_folder_in_explorer(base_folder)
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

    @staticmethod
    def _pump():
        """Let WPF repaint mid-export so queue statuses update live."""
        try:
            Forms.Application.DoEvents()
        except Exception:
            pass

    # ── PDF (native Revit exporter, Revit 2022+) ─────────────────────
    def _pdf_options(self):
        """Build native PDFExportOptions from the Settings tab controls."""
        o = DB.PDFExportOptions()
        try:
            # Paper size + orientation read from each sheet automatically
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
        """Send items to the selected physical printer (PrintManager path)."""
        pm      = self._get_printmanager()
        printer = self._selected_printer
        if not pm or not printer:
            dlg.message('No printer selected for physical printing.')
            return
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

    # ── DWG ──────────────────────────────────────────────────────────
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

    # ── DGN ──────────────────────────────────────────────────────────
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

    # ── NWC ──────────────────────────────────────────────────────────
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

    # ── IFC ──────────────────────────────────────────────────────────
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

    # ── IMG ──────────────────────────────────────────────────────────
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

    def _restore_print_settings(self):
        if self._init_psettings:
            pm = self._get_printmanager()
            if pm:
                with revit.Transaction('Restore Print Settings',
                                       doc=self._selected_doc):
                    pm.PrintSetup.CurrentPrintSetting = self._init_psettings

    # ═══════════════════════════════════════════════════════════════
    # ROW CLICK / MULTI-SELECT  (Shift / Ctrl highlight)
    # ═══════════════════════════════════════════════════════════════
    # ROW CLICK — Ctrl/Shift directly checks/unchecks the checkboxes
    # ═══════════════════════════════════════════════════════════════
    def _row_click(self, sender, args):
        """Tri-select: plain click anchors, Ctrl toggles, Shift range-checks."""
        try:
            hit = args.OriginalSource

            # Clicks on the checkbox itself are handled by the checkbox
            obj = hit
            while obj is not None:
                if isinstance(obj, Windows.Controls.CheckBox):
                    return
                if isinstance(obj, Windows.Controls.DataGridRow):
                    break
                obj = Windows.Media.VisualTreeHelper.GetParent(obj)

            # Find the DataGridRow that was clicked
            row = None
            obj = hit
            while obj is not None:
                if isinstance(obj, Windows.Controls.DataGridRow):
                    row = obj
                    break
                obj = Windows.Media.VisualTreeHelper.GetParent(obj)
            if row is None:
                return

            sheets = self._visible_sheets
            item   = row.Item

            # Locate row index by identity (== is unreliable in IronPython)
            idx = -1
            for i, s in enumerate(sheets):
                if s is item:
                    idx = i
                    break
            if idx == -1:
                return

            ctrl  = (Windows.Input.Keyboard.IsKeyDown(Windows.Input.Key.LeftCtrl) or
                     Windows.Input.Keyboard.IsKeyDown(Windows.Input.Key.RightCtrl))
            shift = (Windows.Input.Keyboard.IsKeyDown(Windows.Input.Key.LeftShift) or
                     Windows.Input.Keyboard.IsKeyDown(Windows.Input.Key.RightShift))

            if shift and self._last_row_index >= 0:
                lo = min(self._last_row_index, idx)
                hi = max(self._last_row_index, idx)
                for i in range(lo, hi + 1):
                    sheets[i].is_selected = True
                args.Handled = True
            elif shift:
                item.is_selected = True
                self._last_row_index = idx
                args.Handled = True
            elif ctrl:
                item.is_selected = not item.is_selected
                self._last_row_index = idx
                args.Handled = True
            else:
                item.is_selected = not item.is_selected
                self._last_row_index = idx
                args.Handled = True

            self.sheets_dg.Items.Refresh()
            self._update_sel_count()
            self._update_select_all_cb()
        except Exception as ex:
            logger.error('Row click failed: %s', ex)

    def _window_click(self, sender, args):
        # Do NOT reset _last_row_index here — it must survive between clicks
        # so Shift+click range selection works correctly.
        pass

    # ═══════════════════════════════════════════════════════════════
    # XAML EVENT HANDLERS — TAB NAV
    # ═══════════════════════════════════════════════════════════════
    def tab_select_clicked(self, sender, args):
        self._show_tab(TAB_SELECT)

    def tab_export_clicked(self, sender, args):
        self._show_tab(TAB_EXPORT)

    def tab_print_clicked(self, sender, args):
        self._build_queue()
        self._show_tab(TAB_PRINT)

    def do_print_clicked(self, sender, args):
        """Export button on the Export tab."""
        self._do_export()

    def help_clicked(self, sender, args):
        dlg.message(
            'pySheets — Sheet Export Tool\n\n'
            '1. Select tab: choose your document, sheet set, '
            'enable formats (PDF/DWG/IFC...) and select sheets per format.\n\n'
            '2. Settings tab: configure export options for each enabled format.\n\n'
            '3. Export tab: set output folder and click Print / Export.',
            title='pySheets Help'
        )

    def about_clicked(self, sender, args):
        dlg.message(
            'pySheets v2.0\n'
            'Part of the Seed pyRevit Extension.\n\n'
            'Multi-format sheet export tool built on pyRevit.\n\n'
            'Special thanks to Ryan McCullough for pyPrintFromIndex.\n'
            'github.com/McCulloughRT/PrintFromIndex',
            title='About pySheets'
        )

    def footer_pyrevit_clicked(self, sender, args):
        try:
            import subprocess
            subprocess.Popen(['cmd', '/c', 'start', '',
                              'https://github.com/pyrevitlabs/pyRevit'])
        except Exception:
            pass

    def footer_ryan_clicked(self, sender, args):
        try:
            import subprocess
            subprocess.Popen(['cmd', '/c', 'start', '',
                              'https://github.com/McCulloughRT/PrintFromIndex'])
        except Exception:
            pass

    # ═══════════════════════════════════════════════════════════════
    # XAML EVENT HANDLERS — SELECT TAB
    # ═══════════════════════════════════════════════════════════════
    def doc_changed(self, sender, args):
        self._project_info = revit.query.get_project_info(doc=self._selected_doc)
        # Clear ALL containers — document changed so all selections are stale
        for c in self._sheet_selections.values(): c.clear()
        for c in self._view_selections.values():  c.clear()
        self._setup_sheet_sets()
        self._setup_print_settings()
        self._setup_export_setups()

    def sheetset_changed(self, sender, args):
        # Clear only sheet containers — sheet set changed so sheet selections are stale
        for c in self._sheet_selections.values(): c.clear()
        self._reload_sheet_list()

    def search_changed(self, sender, args):
        self._apply_filter()

    def active_filter_changed(self, sender, args):
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
        with revit.Transaction('Save Sheet Set', doc=self._selected_doc):
            pm = self._get_printmanager()
            if pm:
                vss = pm.ViewSheetSetting
                # DB.ViewSet is required — List[View] causes a TypeError
                view_set = DB.ViewSet()
                for s in checked:
                    view_set.Insert(s.revit_sheet)
                vss.CurrentViewSheetSet.Views = view_set
                vss.SaveAs(name)
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
        self.active_only_cb.IsEnabled = True
        self._set_size_col_header('Size')
        self._set_sheets_headers()
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
        self.active_only_cb.IsChecked = False
        self.active_only_cb.IsEnabled = False
        self._set_size_col_header(getattr(self, '_type_cb', None) or 'Type')
        self._set_views_headers()
        # Point to views container for current format — do NOT clear it
        self._selection = self._view_selections[self._fmt_viewing]
        self._reload_sheet_list()

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
        except Exception:
            pass
        self._update_sel_count()
        self._update_select_all_cb()

    def grid_selection_changed(self, sender, args):
        try:
            self.sheets_dg.UnselectAll()
        except Exception:
            pass

    def grid_mouse_down(self, sender, args):
        pass  # reserved for future drag-and-drop

    def select_all_clicked(self, sender, args):
        sheets = self._visible_sheets
        # If rows are highlighted via Shift/Ctrl, operate only on those
        if self._highlighted:
            targets = [sheets[i] for i in sorted(self._highlighted)
                       if 0 <= i < len(sheets)]
        else:
            targets = sheets

        checked = [s for s in targets if s.is_selected]
        new_state = len(checked) < len(targets)
        for s in targets:
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
            'EditNamingFormats.xaml',
            start_with=self._selected_naming_format
        ).show_dialog()
        # Refresh dropdown and filenames
        self._setup_naming_formats()
        self._apply_filter()

    # ═══════════════════════════════════════════════════════════════
    # XAML EVENT HANDLERS — EXPORT SETTINGS TAB
    # ═══════════════════════════════════════════════════════════════
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

            # ── Orientation ──────────────────────────────────────────
            if self.orient_portrait_btn.Tag == 'Viewing':
                ips.PrintParameters.PageOrientation = DB.PageOrientationType.Portrait
            elif self.orient_landscape_btn.Tag == 'Viewing':
                ips.PrintParameters.PageOrientation = DB.PageOrientationType.Landscape
            # 'From Sheet' = leave as-is (Revit default)

            # ── Zoom ──────────────────────────────────────────────────
            if self.zoom_fit_rb.IsChecked:
                ips.PrintParameters.ZoomType = DB.ZoomType.FitPage
                ips.PrintParameters.Zoom = 100
            else:
                ips.PrintParameters.ZoomType = DB.ZoomType.Zoom
                try:
                    ips.PrintParameters.Zoom = int(float(self.zoom_pct_tb.Text))
                except Exception:
                    ips.PrintParameters.Zoom = 100

            # ── Paper Placement ───────────────────────────────────────
            if self.placement_center_rb.IsChecked:
                ips.PrintParameters.PaperPlacement = DB.PaperPlacementType.Center
            else:
                ips.PrintParameters.PaperPlacement = DB.PaperPlacementType.LowerLeft
                try:
                    ips.PrintParameters.UserDefinedMarginX = float(self.offset_x_tb.Text)
                    ips.PrintParameters.UserDefinedMarginY = float(self.offset_y_tb.Text)
                except Exception:
                    pass

            # ── Hidden Lines ──────────────────────────────────────────
            if self.hlv_raster_rb.IsChecked:
                ips.PrintParameters.HiddenLineViews = DB.HiddenLineViewsType.RasterProcessing
            else:
                ips.PrintParameters.HiddenLineViews = DB.HiddenLineViewsType.VectorProcessing

            # ── Raster Quality ────────────────────────────────────────
            rq_map = {'High':   DB.RasterQualityType.High,
                      'Medium': DB.RasterQualityType.Medium,
                      'Low':    DB.RasterQualityType.Low}
            rq_text = (self.raster_quality_cb.SelectedItem.Content
                       if self.raster_quality_cb.SelectedItem else 'High')
            ips.PrintParameters.RasterQuality = rq_map.get(rq_text, DB.RasterQualityType.High)

            # ── Colors ────────────────────────────────────────────────
            col_map = {'Color':          DB.ColorDepthType.Color,
                       'Black and White': DB.ColorDepthType.BlackLine,
                       'Grayscale':      DB.ColorDepthType.GrayScale}
            col_text = (self.colors_cb.SelectedItem.Content
                        if self.colors_cb.SelectedItem else 'Color')
            ips.PrintParameters.ColorDepth = col_map.get(col_text, DB.ColorDepthType.Color)

            # ── Options ───────────────────────────────────────────────
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

    def dest_pill_clicked(self, sender, args):
        """Print-destination pills — exclusive single select."""
        for btn in (self.dest_file_btn, self.dest_printer_btn):
            btn.Tag = ''
        sender.Tag = 'Viewing'
        self._update_dest_gates()

    def _set_print_destination(self, dest):
        self.dest_file_btn.Tag    = '' if dest == 'printer' else 'Viewing'
        self.dest_printer_btn.Tag = 'Viewing' if dest == 'printer' else ''
        self._update_dest_gates()

    def _update_dest_gates(self):
        to_printer = self.dest_printer_btn.Tag == 'Viewing'
        self.printer_fields_panel.IsEnabled = to_printer
        self.printer_fields_panel.Opacity = 1.0 if to_printer else 0.4

    def _get_print_destination(self):
        """Return 'file' or 'printer'."""
        if self.dest_printer_btn.Tag == 'Viewing':
            return 'printer'
        return 'file'

    # ── Print Setting management actions ─────────────────────────────
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
            logger.warning('printsetting update skipped: no item selected')
            return
        if item.allows_variable_paper:
            logger.warning('printsetting update skipped: variable paper item')
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

    def exp_subtab_clicked(self, sender, args):
        """Export Settings sub-tab clicked — show that format's panel."""
        fmt = sender.Content
        btn_map = {'PDF': self.exp_pdf_btn, 'DWG': self.exp_dwg_btn,
                   'DGN': self.exp_dgn_btn, 'NWC': self.exp_nwc_btn,
                   'IFC': self.exp_ifc_btn, 'IMG': self.exp_img_btn}
        pnl_map = {'PDF': self.pdf_settings_panel,
                   'DWG': self.dwg_settings_panel,
                   'DGN': self.dgn_settings_panel,
                   'NWC': self.nwc_settings_panel,
                   'IFC': self.ifc_settings_panel,
                   'IMG': self.img_settings_panel}
        for f, btn in btn_map.items():
            btn.Tag = 'Viewing' if f == fmt else ''
        for f, pnl in pnl_map.items():
            pnl.Visibility = (Windows.Visibility.Visible if f == fmt
                              else Windows.Visibility.Collapsed)
        if fmt == 'NWC':
            self._check_nwc_available()

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

    # ── Open native Revit dialogs (closes pySheets first) ────────────
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

    # ═══════════════════════════════════════════════════════════════
    # XAML EVENT HANDLERS — PRINT TAB
    # ═══════════════════════════════════════════════════════════════
    def browse_export_folder(self, sender, args):
        fbd = Forms.FolderBrowserDialog()
        fbd.ShowNewFolderButton = True
        if fbd.ShowDialog() == Forms.DialogResult.OK:
            self._export_folder = fbd.SelectedPath
            self.export_folder_tb.Text = self._export_folder
            self._active_folder_preset = None

    # ── Export folder presets ──────────────────────────────────────
    FOLDER_PRESETS_FILE = op.join(USERDATA_DIR, 'folder_presets.json')

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
            if not op.isdir(USERDATA_DIR):
                os.makedirs(USERDATA_DIR)
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
            tb.Text = 'No presets saved yet — see \u2630 Manage Folder Presets'
            tb.Foreground = self.FindResource('LightText')
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

    # ═══════════════════════════════════════════════════════════════
    # PROFILES  (save / load / import export settings)
    # ═══════════════════════════════════════════════════════════════
    def _setup_profiles(self, select=None):
        """Scan the profiles folder and fill the header + schedule dropdowns."""
        try:
            if not op.isdir(PROFILES_DIR):
                os.makedirs(PROFILES_DIR)
            names = sorted(op.splitext(f)[0]
                           for f in os.listdir(PROFILES_DIR)
                           if f.lower().endswith('.json'))
            self.profile_cb.ItemsSource = names
            if select and select in names:
                self.profile_cb.SelectedItem = select
            try:
                cur = self.sched_profile_cb.SelectedItem
                self.sched_profile_cb.ItemsSource = names
                if cur in names:
                    self.sched_profile_cb.SelectedItem = cur
            except Exception:
                pass
        except Exception as ex:
            logger.warning('Profiles setup failed: %s', ex)

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
            self.subfolder_cb.IsChecked   = data.get('subfolders', True)
            self.open_folder_cb.IsChecked = data.get('open_after', False)
            dest = data.get('dest', 'file')
            self._set_print_destination('printer' if dest == 'printer' else 'file')

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
        name = self.profile_cb.SelectedItem
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
        name = self.profile_cb.SelectedItem
        if not name:
            self.profile_new_clicked(sender, args)
            return
        try:
            with open(self._profile_path(name), 'w') as f:
                json.dump(self._gather_profile(), f, indent=2)
            dlg.message('Profile "{}" saved.'.format(name))
        except Exception as ex:
            dlg.message('Could not save profile.\n\n' + str(ex))

    def profile_delete_clicked(self, sender, args):
        name = self.profile_cb.SelectedItem
        if not name:
            return
        if not dlg.confirm('Delete profile "{}"?'.format(name),
                           yes='Delete'):
            return
        try:
            os.remove(self._profile_path(name))
            self._setup_profiles()
        except Exception as ex:
            dlg.message('Could not delete profile.\n\n' + str(ex))

    def profile_import_clicked(self, sender, args):
        """Import a profile JSON from any location (shared drive, etc.)."""
        dlg = Forms.OpenFileDialog()
        dlg.Filter = 'pySheets profile (*.json)|*.json'
        dlg.Title  = 'Import Profile'
        if dlg.ShowDialog() != Forms.DialogResult.OK:
            return
        try:
            with open(dlg.FileName, 'r') as f:
                data = json.load(f)
            name = op.splitext(op.basename(dlg.FileName))[0]
            with open(self._profile_path(name), 'w') as f:
                json.dump(data, f, indent=2)
            self._setup_profiles(select=name)
            self._apply_profile(data)
        except Exception as ex:
            dlg.message('Could not import profile.\n\n' + str(ex))

    # ═══════════════════════════════════════════════════════════════
    # SCHEDULING  (runs while this window and Revit stay open)
    # Chain: Enable → Profile → Time → Repeat (→ Date + weekdays)
    # ═══════════════════════════════════════════════════════════════
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
        h = ((h24 - 1) % 12) + 1
        ap = 'PM' if h24 >= 12 else 'AM'
        self.sched_time_btn.Content = '{:02d}:{:02d} {}'.format(h, m, ap)

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
        self.sched_time_btn.Content = '{:02d}:{:02d} {}'.format(h, m, ap)
        self.sched_time_btn.IsChecked = False   # close popup
        self._sched_refresh()

    def tp_cancel_clicked(self, sender, args):
        self.sched_time_btn.IsChecked = False

    def sched_enable_changed(self, sender, args):
        if self._loading:
            return
        if self.sched_enable_cb.IsChecked:
            pass
        else:
            # Reset: profile must be re-picked next time
            self._loading = True
            self.sched_profile_cb.SelectedIndex = -1
            h24, m = self._current_time_rounded()
            self._sched_time = (h24, m)
            self._apply_time_to_wheels(h24, m)
            h = ((h24 - 1) % 12) + 1
            ap = 'PM' if h24 >= 12 else 'AM'
            self.sched_time_btn.Content = '{:02d}:{:02d} {}'.format(h, m, ap)
            self.sched_repeat_cb.SelectedIndex = 0
            for n in self.SCHED_DAYS:
                getattr(self, n).IsChecked = False
            self._loading = False
        self._sched_refresh()

    def sched_field_changed(self, sender, args):
        if self._loading:
            return
        self._sched_refresh()

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
        """Unlock each control only when the previous step is set."""
        en     = bool(self.sched_enable_cb.IsChecked)
        prof   = en and bool(self.sched_profile_cb.SelectedItem)
        t_ok   = prof and self._parse_sched_time() is not None
        repeat = t_ok and self._sched_repeat_mode() == 'Repeat'

        self.sched_profile_cb.IsEnabled = en
        self.sched_time_btn.IsEnabled   = prof
        self.sched_repeat_cb.IsEnabled  = t_ok
        self.sched_date_dp.IsEnabled    = repeat
        self.sched_days_panel.Visibility = (
            Windows.Visibility.Visible if repeat
            else Windows.Visibility.Collapsed)
        for ctrl, on in ((self.sched_profile_cb, en),
                         (self.sched_time_btn, prof),
                         (self.sched_repeat_cb, t_ok)):
            ctrl.Opacity = 1.0 if on else 0.55

    def _compute_next_run(self):
        """Next datetime the schedule should fire, or None."""
        t = self._parse_sched_time()
        if t is None:
            return None
        hh, mm = t
        now = datetime.now()
        if self._sched_repeat_mode() == 'Once':
            nxt = now.replace(hour=hh, minute=mm, second=0, microsecond=0)
            if nxt <= now:
                nxt += timedelta(days=1)
            return nxt
        days = self._sched_checked_days()
        if not days:
            return None
        base = max(self._sched_start_date(), now.date())
        for i in range(8):
            d = base + timedelta(days=i)
            cand = datetime(d.year, d.month, d.day, hh, mm)
            if cand.weekday() in days and cand > now:
                return cand
        return None

    def _sched_refresh(self):
        """Re-evaluate gates and arm/disarm the timer."""
        self._update_sched_gates()
        self._stop_timer()
        if not self.sched_enable_cb.IsChecked:
            self.sched_status_tb.Text = 'Schedule off'
            return
        if not self.sched_profile_cb.SelectedItem:
            self.sched_status_tb.Text = 'Select a profile…'
            return
        if self._parse_sched_time() is None:
            self.sched_status_tb.Text = 'Pick a time…'
            return
        if (self._sched_repeat_mode() == 'Repeat'
                and not self._sched_checked_days()):
            self.sched_status_tb.Text = 'Pick at least one day…'
            return
        self._sched_next = self._compute_next_run()
        if not self._sched_next:
            self.sched_status_tb.Text = 'No valid run time'
            return
        self._sched_timer = Windows.Threading.DispatcherTimer()
        self._sched_timer.Interval = framework.System.TimeSpan.FromSeconds(20)
        self._sched_timer.Tick += self._sched_tick
        self._sched_timer.Start()
        self.sched_status_tb.Text = 'Next run: {}'.format(
            self._sched_next.strftime('%a %d %b %H:%M'))

    def _stop_timer(self):
        if self._sched_timer:
            try:
                self._sched_timer.Stop()
            except Exception:
                pass
        self._sched_timer = None
        self._sched_next  = None

    def _stop_schedule(self):
        """Full stop — also unticks Enable (used on close/errors)."""
        self._stop_timer()
        try:
            if self.sched_enable_cb.IsChecked:
                self._loading = True
                self.sched_enable_cb.IsChecked = False
                self._loading = False
        except Exception:
            pass
        self.sched_status_tb.Text = 'Schedule off'

    def _sched_tick(self, sender, args):
        if not self._sched_next or datetime.now() < self._sched_next:
            return
        try:
            name = self.sched_profile_cb.SelectedItem
            if name:
                with open(self._profile_path(name), 'r') as f:
                    self._apply_profile(json.load(f))
            self._build_queue()
            self._do_export()
        except Exception as ex:
            logger.error('Scheduled export failed: %s', ex)
        if self._sched_repeat_mode() == 'Once':
            self._stop_schedule()
        else:
            self._sched_refresh()

    # ═══════════════════════════════════════════════════════════════
    # XAML EVENT HANDLERS — WINDOW
    # ═══════════════════════════════════════════════════════════════
    LAST_SESSION = op.join(USERDATA_DIR, 'lastsession.json')

    def _save_last_session(self):
        try:
            if not op.isdir(USERDATA_DIR):
                os.makedirs(USERDATA_DIR)
            data = self._gather_profile()
            data.pop('sheet_sel', None)   # selections are per-session
            data.pop('view_sel', None)
            with open(self.LAST_SESSION, 'w') as f:
                json.dump(data, f, indent=2)
        except Exception as ex:
            logger.warning('Last session save failed: %s', ex)

    def _load_last_session(self):
        try:
            if op.isfile(self.LAST_SESSION):
                with open(self.LAST_SESSION, 'r') as f:
                    self._apply_profile(json.load(f))
        except Exception as ex:
            logger.warning('Last session load failed: %s', ex)

    def window_closing(self, sender, args):
        self._save_last_session()
        self._save_naming_memory()
        self._stop_schedule()
        self._restore_print_settings()

    def close_window(self, sender, args):
        self.Close()


# ═══════════════════════════════════════════════════════════════════
# HELPERS  (module-level, no access to self)
# ═══════════════════════════════════════════════════════════════════
def _brush(color_hex):
    """Return a SolidColorBrush from a hex colour string."""
    try:
        color = Windows.Media.ColorConverter.ConvertFromString(color_hex)
        b = Windows.Media.SolidColorBrush(color)
        b.Freeze()
        return b
    except Exception:
        return Windows.Media.Brushes.Transparent


def _set_child_text(border, text, color_hex):
    """Set the text and foreground of the first TextBlock child of a Border."""
    try:
        tb = border.Child
        if hasattr(tb, 'Text'):
            tb.Text       = text
            tb.Foreground = _brush(color_hex)
    except Exception:
        pass


# ═══════════════════════════════════════════════════════════════════
# ENTRY POINT
# ═══════════════════════════════════════════════════════════════════
if __name__ == '__main__':
    try:
        PrintSheetsWindow('pySheets.xaml').show_dialog()
    except Exception as e:
        import traceback
        logger.error(traceback.format_exc())
        forms.alert('Error starting pySheets', expanded=str(e))
