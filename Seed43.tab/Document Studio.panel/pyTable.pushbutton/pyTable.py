# -*- coding: utf-8 -*-
"""
pyTable - export model/schedule parameters to Excel or ODS, edit externally,
re-import with a diff preview before anything is written back to the model.

𝐒𝐄𝐄𝐃𝟒𝟑
"""


import os
import sys
import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System import Int64
from System.Windows import Thickness, Visibility
from System.Windows.Controls import ListBoxItem, Button
from System.Windows.Shapes import Rectangle
from System.Collections.ObjectModel import ObservableCollection


from Autodesk.Revit.DB import Transaction, FilteredElementCollector

from pyrevit import revit, DB, script, forms

from Snippets.seed43_theme import (apply_seed43_palette, apply_seed43_dimensions,
                                   get_color)
from Snippets import _dialogs as dlg
from Snippets._icons import make_icon, make_icon_with_label
from Snippets._selection import get_element_type
from Snippets._support import github_issue_url, open_url, support_mailto
from Snippets._parameters import (
    get_param_display_value,
    set_param_from_display_value,
    get_categories_with_elements,
    get_elements_by_category,
    get_schedules,
    get_param_dropdown_options,
)

# Was tools/pytable_io.py, local to this tool. Moved to the lib when pySheets
# needed the same xlsx/ods writer for its schedule export - two consumers, so
# it belongs in one place rather than being copied.
from Snippets._spreadsheet import write_workbook, read_workbook, HIDDEN_COLUMNS

SCRIPT_DIR = os.path.dirname(__file__)


doc = revit.doc
logger = script.get_logger()

TYPE_SUFFIX = u" (Type)"


def _encode_display(raw_name, kind):
    return raw_name + TYPE_SUFFIX if kind == "type" else raw_name


def _element_id(value):
    """ElementId from a spreadsheet cell or stored id. ElementId(int) is
    ambiguous in IronPython against Revit 2024+ (Int64 vs BuiltInParameter
    vs BuiltInCategory), so build it from an explicit Int64."""
    return DB.ElementId(Int64(long(float(value))))


def _decode_display(display_name):
    if display_name.endswith(TYPE_SUFFIX):
        return display_name[:-len(TYPE_SUFFIX)], "type"
    return display_name, "instance"


LENGTH_PARAM = u"Length"


def _editable_length(el, raw_name):
    """True when this is the Length of a straight line-based element (wall,
    beam, line). Revit makes Length read-only because the element's line
    drives it, but pyTable can set it by moving the end of that line."""
    if raw_name != LENGTH_PARAM or el is None:
        return False
    loc = getattr(el, "Location", None)
    return isinstance(loc, DB.LocationCurve) and isinstance(loc.Curve, DB.Line)


def _set_curve_length(el, text):
    """Make a straight element `text` long (project units), keeping its start
    point and direction: the end point moves along the line."""
    ok, length = DB.UnitFormatUtils.TryParse(
        doc.GetUnits(), DB.SpecTypeId.Length, text or u"")
    if not ok:
        raise ValueError(u"'Length' could not take '{0}': not a valid length "
                         u"in project units".format(text))
    if length <= doc.Application.ShortCurveTolerance:
        raise ValueError(u"'Length' must be more than zero, not '{0}'".format(text))
    # A wall joined at its end snaps back to the join when Revit
    # regenerates (8000 set, 6000 again after regen). Release that end's
    # join, move it, then allow joins again so it rejoins anything it now
    # meets. The other wall at the old corner is not moved.
    is_wall = isinstance(el, DB.Wall)
    if is_wall:
        DB.WallUtils.DisallowWallJoinAtEnd(el, 1)
    loc = el.Location
    line = loc.Curve
    start = line.GetEndPoint(0)
    direction = (line.GetEndPoint(1) - start).Normalize()
    loc.Curve = DB.Line.CreateBound(start, start + direction * length)
    doc.Regenerate()
    if is_wall:
        DB.WallUtils.AllowWallJoinAtEnd(el, 1)
        doc.Regenerate()
    # Never report a length as applied if something pulled it back.
    if abs(el.Location.Curve.Length - length) > doc.Application.ShortCurveTolerance:
        raise ValueError(u"'Length' set to '{0}' but Revit held the element at "
                         u"its old length (joined or constrained)".format(text))


class _FailureCollector(DB.IFailuresPreprocessor):
    """Handles what Revit reports when the apply commits: warnings are
    recorded and dismissed, errors are recorded and roll the whole
    transaction back (nothing half-applied, no Revit dialog)."""
    def __init__(self):
        self.warnings = []
        self.errors = []

    def PreprocessFailures(self, accessor):
        for msg in accessor.GetFailureMessages():
            text = msg.GetDescriptionText()
            if msg.GetSeverity() == DB.FailureSeverity.Warning:
                self.warnings.append(text)
                accessor.DeleteWarning(msg)
            else:
                self.errors.append(text)
        if self.errors:
            return DB.FailureProcessingResult.ProceedWithRollBack
        return DB.FailureProcessingResult.Continue


class DiffRow(object):
    """Bindable row for the Preview/Edit DataGrid."""
    def __init__(self, include, category, element_label, param_name, old_value, new_value,
                 write_eid, kind):
        self.Include = include
        self.Category = category
        self.ElementLabel = element_label
        self.ParamName = param_name
        self.OldValue = old_value
        self.NewValue = new_value
        self._write_eid = write_eid   # element to actually Set() on - instance id or type id
        self.Kind = "Type" if kind == "type" else "Instance"


class MainWindow(forms.WPFWindow):

    def __init__(self):
        forms.WPFWindow.__init__(self, "pyTable.xaml")

        # Rule #2: apply appearance immediately after load, before any
        # dynamic UI construction touches TryFindResource.
        apply_seed43_palette(self, SCRIPT_DIR)
        apply_seed43_dimensions(self, SCRIPT_DIR)
        from Snippets._icons import set_header_icon
        set_header_icon(self, SCRIPT_DIR)

        self._pending_export_columns = []   # list of {"key","kind","readonly"}
        self._pending_export_rows = []       # list of dicts keyed by column + hidden keys
        self._diff_rows = ObservableCollection[object]()
        # The window is modal, so the selection can't change while it is
        # open: read it once for the Current Selection scope.
        self._selection_ids = set(
            i.Value for i in revit.uidoc.Selection.GetElementIds())

        self._set_transfer_icons()
        self._bind_events()
        self._populate_categories()
        self._populate_schedules()

    # ── wiring ──────────────────────────────────────────────────────────

    def _bind_events(self):
        self.FindName("close_btn").Click += self.on_close
        self.FindName("menu_popup").Opened += self.on_menu_opened
        self.FindName("export_btn").Click += self.on_export
        self.FindName("import_btn").Click += self.on_import
        self.FindName("confirm_btn").Click += self.on_confirm
        self.FindName("reset_btn").Click += self.on_reset

        self.FindName("categories_list").SelectionChanged += self.on_categories_selection_changed
        self.FindName("schedules_list").SelectionChanged += self.on_schedules_selection_changed

        self.FindName("cat_params_add_btn").Click += self._make_transfer_handler(
            "available_params_categories", "selected_params_categories")
        self.FindName("cat_params_remove_btn").Click += self._make_transfer_handler(
            "selected_params_categories", "available_params_categories")
        self.FindName("sched_params_add_btn").Click += self._make_transfer_handler(
            "available_params_schedules", "selected_params_schedules")
        self.FindName("sched_params_remove_btn").Click += self._make_transfer_handler(
            "selected_params_schedules", "available_params_schedules")

        for name in ("scope_whole_model", "scope_active_view",
                     "scope_current_selection"):
            self.FindName(name).Checked += self.on_scope_changed
        self.FindName("schedules_search").TextChanged += self.on_schedules_search

        diff_grid = self.FindName("diff_grid")
        diff_grid.ItemsSource = self._diff_rows

    def _set_transfer_icons(self):
        """Arrow icons on the add/remove buttons, from the shared icon set.
        make_icon bakes its colour in; the window has no theme toggle, so
        building them once after the palette is applied is enough."""
        color = get_color(SCRIPT_DIR, 'text_primary', fallback='#F4FAFF')
        for name, key in (("cat_params_add_btn", "arrow_right"),
                          ("sched_params_add_btn", "arrow_right"),
                          ("cat_params_remove_btn", "arrow_left"),
                          ("sched_params_remove_btn", "arrow_left")):
            try:
                self.FindName(name).Content = make_icon(key, size=14, color=color)
            except Exception:
                pass
        # Header buttons: icons from the same file instead of text glyphs.
        for name, key, size in (("menu_btn", "menu", 18), ("close_btn", "close", 14)):
            try:
                self.FindName(name).Content = make_icon(key, size=size, color=color)
            except Exception:
                pass

    def _make_transfer_handler(self, source_name, target_name):
        """Return a Click handler that moves selected items from one
        ListBox to another, skipping names already present in the target
        (so re-clicking Add after a Remove doesn't create duplicates)."""
        def _handler(sender, args):
            source = self.FindName(source_name)
            target = self.FindName(target_name)
            existing = set(i.Content for i in target.Items)
            for item in list(source.SelectedItems):
                if item.Content not in existing:
                    new_item = ListBoxItem()
                    new_item.Content = item.Content
                    new_item.Tag = item.Tag
                    target.Items.Add(new_item)
                    existing.add(item.Content)
                source.Items.Remove(item)
        return _handler

    # ── population (Model Categories / Schedules tabs) ────────────────

    # -- scope (Whole Model / Active View / Current Selection) ----------

    def _scope(self):
        if self.FindName("scope_active_view").IsChecked:
            return "view"
        if self.FindName("scope_current_selection").IsChecked:
            return "selection"
        return "model"

    def _scope_elements(self):
        """Every placed element in the active view or the selection; None
        for the whole model (categories come from the model instead)."""
        scope = self._scope()
        if scope == "view":
            try:
                return list(FilteredElementCollector(doc, doc.ActiveView.Id)
                            .WhereElementIsNotElementType())
            except Exception:
                return []
        if scope == "selection":
            out = []
            for v in self._selection_ids:
                el = doc.GetElement(_element_id(v))
                if el is not None and not isinstance(el, DB.ElementType):
                    out.append(el)
            return out
        return None

    def _scoped_elements(self, cat_id):
        """Elements of one category, limited to the chosen scope."""
        if self._scope() == "model":
            return get_elements_by_category(doc, cat_id)
        cat_value = cat_id.Value
        return [e for e in (self._scope_elements() or [])
                if e.Category is not None and e.Category.Id.Value == cat_value]

    def on_scope_changed(self, sender, args):
        """Scope decides which categories are listed and which elements are
        exported, so the category list is rebuilt for it."""
        self._populate_categories()
        self._set_status({
            "model": "Scope: whole model.",
            "view": "Scope: elements in the active view.",
            "selection": "Scope: {0} selected elements.".format(
                len(self._selection_ids)),
        }[self._scope()])

    def _populate_categories(self):
        """List every category that has at least one element in scope."""
        categories_list = self.FindName("categories_list")
        categories_list.Items.Clear()
        elements = self._scope_elements()
        if elements is None:
            cats = get_categories_with_elements(doc)
        else:
            by_name = {}
            for el in elements:
                cat = el.Category
                if cat is not None and cat.Name not in by_name:
                    by_name[cat.Name] = cat
            cats = [by_name[n] for n in sorted(by_name)]
        for cat in cats:
            item = ListBoxItem()
            item.Content = cat.Name
            item.Tag = cat.Id
            categories_list.Items.Add(item)
        # Clearing the list doesn't always raise SelectionChanged, so drop
        # the parameters offered for the old category set explicitly.
        self.on_categories_selection_changed(None, None)

    def on_schedules_search(self, sender, args):
        """Filter the schedule list as you type (the selection is kept)."""
        text = (sender.Text or u"").strip().lower()
        for item in self.FindName("schedules_list").Items:
            match = not text or text in (item.Content or u"").lower()
            item.Visibility = Visibility.Visible if match else Visibility.Collapsed

    def _populate_schedules(self):
        schedules_list = self.FindName("schedules_list")
        for sched in get_schedules(doc):
            item = ListBoxItem()
            item.Content = sched.Name
            item.Tag = sched.Id
            schedules_list.Items.Add(item)

    def _get_checked_category_ids(self):
        categories_list = self.FindName("categories_list")
        return [i.Tag for i in categories_list.SelectedItems]

    def _get_checked_schedule_ids(self):
        schedules_list = self.FindName("schedules_list")
        return [i.Tag for i in schedules_list.SelectedItems]

    def on_categories_selection_changed(self, sender, args):
        """Rebuild Available Parameters from every distinct instance AND
        type param name found on elements/types in the newly-checked
        categories, minus anything already moved to Selected."""
        names = {}   # display_name -> (raw_name, kind)
        for cat_id in self._get_checked_category_ids():
            elements = self._scoped_elements(cat_id)
            seen_type_ids = set()
            for el in elements:
                for param in el.Parameters:
                    raw_name = param.Definition.Name
                    names[_encode_display(raw_name, "instance")] = (raw_name, "instance")

                type_el = get_element_type(el, doc)
                if type_el is not None and type_el.Id.Value not in seen_type_ids:
                    seen_type_ids.add(type_el.Id.Value)
                    for param in type_el.Parameters:
                        raw_name = param.Definition.Name
                        names[_encode_display(raw_name, "type")] = (raw_name, "type")

        self._refresh_available_list("available_params_categories", "selected_params_categories", names)

    def on_schedules_selection_changed(self, sender, args):
        """Rebuild Available Parameters from the ScheduleDefinition fields
        of every newly-checked schedule. Schedule fields are treated as
        instance-kind (the schedule field API doesn't cleanly expose
        Type-vs-Instance the way Element.Parameters does)."""
        names = {}
        for sched_id in self._get_checked_schedule_ids():
            schedule = doc.GetElement(sched_id)
            definition = schedule.Definition
            sample = self._schedule_elements(sched_id)[:20]
            for i in range(definition.GetFieldCount()):
                field = definition.GetField(i)
                raw_name = field.GetName()
                kind = self._field_kind(raw_name, sample)
                names[_encode_display(raw_name, kind)] = (raw_name, kind)

        self._refresh_available_list("available_params_schedules", "selected_params_schedules", names)

    @staticmethod
    def _schedule_elements(sched_id):
        """The elements a schedule lists (its own filters applied)."""
        try:
            return list(FilteredElementCollector(doc, sched_id)
                        .WhereElementIsNotElementType())
        except Exception:
            return []

    @staticmethod
    def _field_kind(raw_name, sample):
        """'instance' or 'type' for a schedule field: the field API doesn't
        say, so look the name up on a few of the schedule's elements, then
        on their types. Defaults to instance."""
        for el in sample:
            if el.LookupParameter(raw_name) is not None:
                return "instance"
            type_el = get_element_type(el, doc)
            if type_el is not None and type_el.LookupParameter(raw_name) is not None:
                return "type"
        return "instance"

    def _refresh_available_list(self, available_name, selected_name, names):
        """names: dict of display_text -> (raw_name, kind)."""
        available_list = self.FindName(available_name)
        selected_list = self.FindName(selected_name)
        already_selected = set(i.Content for i in selected_list.Items)

        available_list.Items.Clear()
        for display_text in sorted(names.keys()):
            if display_text in already_selected:
                continue
            raw_name, kind = names[display_text]
            item = ListBoxItem()
            item.Content = display_text
            item.Tag = (raw_name, kind)
            available_list.Items.Add(item)

    # ── export ──────────────────────────────────────────────────────────

    def _get_selected_params(self, selected_list_name):
        """Return list of (display_text, raw_name, kind) for a Selected list."""
        selected_list = self.FindName(selected_list_name)
        result = []
        for i in selected_list.Items:
            raw_name, kind = i.Tag
            result.append((i.Content, raw_name, kind))
        return result

    def _collect_export_data(self):
        """
        Build the shared (columns, rows) shape from whatever's checked on
        the Model Categories / Schedules tabs, restricted to the parameters
        moved into each tab's Selected Parameters list. Type parameters
        (suffixed " (Type)" in the picker) are pulled from the element's
        Type element rather than the instance, and repeat per-instance since
        every instance of the same type shares that value - matches how
        SheetLink itself displays type columns.

        TODO (needs live Revit testing): schedule-sourced rows currently
        reuse the same per-element param lookup as categories - once
        ScheduleDefinition sort/group/filter support matters, this should
        walk the schedule's own row set instead of re-collecting elements
        from scratch.
        """
        wanted = self._get_selected_params("selected_params_categories")
        wanted += self._get_selected_params("selected_params_schedules")

        if not wanted:
            return [], []

        # display_text -> (raw_name, kind), deduped
        wanted_map = {}
        for display_text, raw_name, kind in wanted:
            wanted_map[display_text] = (raw_name, kind)

        columns_seen = {
            dt: {"key": dt, "kind": kind, "readonly": False, "options": None}
            for dt, (raw_name, kind) in wanted_map.items()
        }
        # Yes/No and known-enum detection don't depend on the current value, so
        # they succeed on the first attempt. Only the ElementId-category path
        # needs a non-empty value, and plenty of ElementId parameters (Scope
        # Box, View Template, Design Option) are empty on most elements. Retry
        # across elements up to the cap rather than giving up on the first.
        MAX_OPTION_DISCOVERY_ATTEMPTS = 25
        options_attempts = {}
        rows = []
        seen_eids = set()

        # Category rows follow the scope; a schedule brings exactly the
        # elements it lists (its own filters already decide those).
        sources = [self._scoped_elements(c) for c in self._get_checked_category_ids()]
        sources += [self._schedule_elements(s) for s in self._get_checked_schedule_ids()]

        for elements in sources:
            for el in elements:
                if el.Id.Value in seen_eids:
                    continue
                seen_eids.add(el.Id.Value)
                row = {"_eid": el.Id.Value, "_category": el.Category.Name if el.Category else ""}

                type_el = get_element_type(el, doc)

                for display_text, (raw_name, kind) in wanted_map.items():
                    if kind == "type":
                        param = type_el.LookupParameter(raw_name) if type_el is not None else None
                    else:
                        param = el.LookupParameter(raw_name)

                    if param is None:
                        # Sentinel: this parameter doesn't exist on this
                        # element, so its cell is locked in the workbook.
                        row[display_text] = None
                        continue

                    columns_seen[display_text]["readonly"] = (
                        param.IsReadOnly and not _editable_length(el, raw_name))
                    # A parameter that exists but was never filled in
                    # (HasValue False: empty Comments, Mark, ...) exports as
                    # an empty, editable cell. Treating it as missing locked
                    # it, so a blank parameter could never be filled in.
                    row[display_text] = (get_param_display_value(param)
                                         if param.HasValue else u"")

                    attempts = options_attempts.get(display_text, 0)
                    if columns_seen[display_text]["options"] is None and attempts < MAX_OPTION_DISCOVERY_ATTEMPTS:
                        options_attempts[display_text] = attempts + 1
                        try:
                            found = get_param_dropdown_options(param, doc)
                            if found:
                                columns_seen[display_text]["options"] = found
                        except Exception:
                            pass

                rows.append(row)

        columns = list(columns_seen.values())
        return columns, rows

    def on_export(self, sender, args):
        columns, rows = self._collect_export_data()
        if not rows:
            self._set_status("Nothing selected to export - check a category or schedule first.")
            return

        fmt_index = self.FindName("format_combo").SelectedIndex
        ext = "xlsx" if fmt_index == 0 else "ods"
        filter_str = "Excel Workbook (*.xlsx)|*.xlsx" if ext == "xlsx" else "OpenDocument Spreadsheet (*.ods)|*.ods"

        filepath = forms.save_file(file_ext=ext, files_filter=filter_str, default_name="pyTable Export")
        if not filepath:
            return

        MAX_SAVE_RETRIES = 5
        saved = False
        last_ex = None
        for attempt in range(MAX_SAVE_RETRIES):
            try:
                write_workbook(filepath, columns, rows)
                saved = True
                break
            except Exception as ex:
                last_ex = ex
                if attempt < MAX_SAVE_RETRIES - 1:
                    # Matches pyTransmit's own file-locked handling: don't
                    # fail outright, give the user a chance to close the
                    # file (most likely already open in Excel/LibreOffice)
                    # and retry, or cancel.
                    retry = dlg.confirm(
                        "Could not save, the file may already be open in "
                        "Excel or LibreOffice.\n\nClose it, then click "
                        "Retry.\n\nFile: {0}\n\nError: {1}".format(filepath, ex),
                        title="File Locked", yes="Retry", no="Cancel")
                    if not retry:
                        self._set_status("Export cancelled.")
                        return
                else:
                    logger.error("pyTable export failed after {0} attempts: {1}".format(
                        MAX_SAVE_RETRIES, ex))
                    dlg.message(
                        "Could not save after {0} attempts:\n{1}".format(MAX_SAVE_RETRIES, ex),
                        title="pyTable")
                    self._set_status("Export failed after {0} attempts - see log.".format(MAX_SAVE_RETRIES))
                    return

        if not saved:
            return

        if not os.path.exists(filepath):
            dlg.message(
                "write_workbook() returned without error but no file was "
                "found at:\n{0}".format(filepath), title="pyTable")
            self._set_status("Export reported success but file was not found.")
            return

        self._set_status("Exported {0} rows to {1}".format(len(rows), filepath))

    # ── import / diff ──────────────────────────────────────────────────

    def on_import(self, sender, args):
        filepath = forms.pick_file(files_filter="Supported Files (*.xlsx;*.ods)|*.xlsx;*.ods")
        if not filepath:
            return
        self.load_diff(filepath)

    def load_diff(self, filepath):
        """Read an edited workbook and fill the Preview/Edit grid with every
        value that differs from the model. Split from on_import so it can be
        driven without the file picker."""
        try:
            headers, imported_rows = read_workbook(filepath)
        except Exception as ex:
            logger.error("pyTable import (read) failed: {0}".format(ex))
            dlg.message("Could not read file:\n{0}".format(ex), title="pyTable")
            self._set_status("Import failed - see log.")
            return

        self._diff_rows.Clear()

        changed_count = 0
        seen_type_diffs = set()   # (type_eid, raw_name) - only show one row per type param

        for row in imported_rows:
            eid = row.get("_eid")
            if eid is None:
                continue
            try:
                element = doc.GetElement(_element_id(eid))
            except Exception:
                continue
            if element is None:
                continue

            # Iterate the full header set, not row.items() - a real Excel/
            # LibreOffice save can omit a cell entirely when the user
            # clears it (rather than writing an empty string), so relying
            # on row.items() would silently miss that as a change.
            for display_text in headers:
                if display_text in HIDDEN_COLUMNS:
                    continue
                new_value = row.get(display_text, u"")
                raw_name, kind = _decode_display(display_text)

                if kind == "type":
                    type_id = element.GetTypeId()
                    if type_id is None or type_id.Value == -1:
                        continue
                    write_target = doc.GetElement(type_id)
                    dedupe_key = (type_id.Value, raw_name)
                    if dedupe_key in seen_type_diffs:
                        continue
                    write_eid = type_id.Value
                    # .Name isn't reachable on an ElementType from IronPython
                    # (it read "?"); the Element.Name getter is.
                    try:
                        type_name = DB.Element.Name.GetValue(write_target)
                    except Exception:
                        type_name = u"?"
                    element_label = u"Type: {0} (Id {1})".format(type_name, write_eid)
                else:
                    write_target = element
                    write_eid = eid
                    element_label = "Id {0}".format(eid)

                param = write_target.LookupParameter(raw_name) if write_target is not None else None
                if param is None or (param.IsReadOnly and not (
                        kind == "instance" and _editable_length(write_target, raw_name))):
                    continue

                old_value = get_param_display_value(param)
                new_value_str = u"" if new_value is None else unicode(new_value)
                if old_value == new_value_str:
                    continue

                if kind == "type":
                    seen_type_diffs.add(dedupe_key)

                self._diff_rows.Add(DiffRow(
                    include=True,
                    category=row.get("_category", ""),
                    element_label=element_label,
                    param_name=raw_name,
                    old_value=old_value,
                    new_value=new_value_str,
                    write_eid=write_eid,
                    kind=kind,
                ))
                changed_count += 1

        self.FindName("diff_summary_label").Text = "{0} changes found".format(changed_count)
        self.FindName("preview_tab").IsEnabled = changed_count > 0
        self.FindName("confirm_btn").IsEnabled = changed_count > 0
        if changed_count > 0:
            self.FindName("main_tabs").SelectedItem = self.FindName("preview_tab")
        self._set_status("Loaded {0} rows, {1} changed values".format(len(imported_rows), changed_count))

    def _start_checked(self, name):
        """A Transaction whose commit problems are collected, not shown.

        Revit only checks new values at commit (duplicate marks, a zero wall
        height, a sheet number in use). Left alone, warnings and errors pop
        Revit's own failure dialogs over this modal window. Here warnings
        are logged and dismissed, and any error rolls the transaction back.
        """
        t = Transaction(doc, name)
        failures = _FailureCollector()
        opts = t.GetFailureHandlingOptions()
        opts.SetFailuresPreprocessor(failures)
        opts.SetClearAfterRollback(True)
        t.SetFailureHandlingOptions(opts)
        t.Start()
        return t, failures

    @staticmethod
    def _write_row(row):
        """Write one preview row to the model. False when there is nothing
        to write to (element or parameter gone, now read-only); raises
        ValueError when the value itself is refused."""
        target = doc.GetElement(_element_id(row._write_eid))
        if target is None:
            return False
        if row.Kind == "Instance" and _editable_length(target, row.ParamName):
            _set_curve_length(target, row.NewValue)
            return True
        param = target.LookupParameter(row.ParamName)
        if param is None or param.IsReadOnly:
            return False
        set_param_from_display_value(param, row.NewValue)
        return True

    def _apply_rows(self, rows):
        """Write every row in ONE transaction, so the model is never left
        half-written. Returns (status, applied, skipped, failures)."""
        t, failures = self._start_checked("pyTable: Apply parameter changes")
        applied = 0
        skipped = []
        try:
            for row in rows:
                try:
                    if self._write_row(row):
                        applied += 1
                except Exception as row_ex:
                    skipped.append(u"{0}, {1}: {2}".format(
                        row.ElementLabel, row.ParamName, row_ex))
            status = t.Commit()
        except Exception:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
            raise
        return status, applied, skipped, failures

    def _find_rejected(self, rows):
        """Which rows Revit refuses at commit, and why.

        Only run after the whole apply was rolled back. Each row is tried in
        its own transaction inside a group that is rolled back at the end,
        so nothing here stays in the model. Some refusals (a wall height of
        0) roll back with no message at all, hence the fallback reason.
        """
        rejected = []
        group = DB.TransactionGroup(doc, "pyTable: check changes")
        group.Start()
        try:
            for row in rows:
                t, failures = self._start_checked("pyTable: check")
                try:
                    if not self._write_row(row):
                        t.RollBack()
                        continue
                except Exception:
                    # Refused while setting: _apply_rows skips these anyway.
                    if not t.HasEnded():
                        t.RollBack()
                    continue
                if t.Commit() != DB.TransactionStatus.Committed:
                    rejected.append((row, u"; ".join(failures.errors)
                                     or u"Revit refused the value"))
        finally:
            group.RollBack()
        return rejected

    def _refresh_diff_grid(self):
        """Re-add the rows so the grid shows changed Include ticks (IronPython
        objects don't notify WPF of attribute changes)."""
        rows = list(self._diff_rows)
        self._diff_rows.Clear()
        for row in rows:
            self._diff_rows.Add(row)

    def on_confirm(self, sender, args):
        rows_to_apply = [r for r in self._diff_rows if r.Include]
        if not rows_to_apply:
            self._set_status("Nothing checked to apply.")
            return

        try:
            status, applied, skipped, failures = self._apply_rows(rows_to_apply)
        except Exception as ex:
            logger.error("pyTable import failed: {0}".format(ex))
            dlg.message("Applying changes failed, nothing was written:\n{0}".format(ex),
                        title="pyTable")
            self._set_status("Import failed, no changes were made, see log.")
            return

        if status != DB.TransactionStatus.Committed:
            # Nothing was written. Find the rows Revit refuses, untick them
            # and keep the preview, so Confirm Changes applies the rest.
            rejected = self._find_rejected(rows_to_apply)
            for row, reason in rejected:
                row.Include = False
                logger.error(u"pyTable: Revit refused {0}, {1} = '{2}': {3}".format(
                    row.ElementLabel, row.ParamName, row.NewValue, reason))
            self._refresh_diff_grid()
            lines = [u"{0}, {1} = '{2}': {3}".format(
                r.ElementLabel, r.ParamName, r.NewValue[:40], reason)
                for r, reason in rejected[:10]]
            dlg.message(
                u"Revit refused some of these changes, so nothing was written."
                u"\n\n" + (u"\n".join(lines) or u"(Revit gave no reason)")
                + (u"\n..." if len(rejected) > 10 else u"")
                + u"\n\nThose rows are now unticked. Click Confirm Changes "
                  u"again to apply the rest.",
                title="pyTable")
            self._set_status("Nothing written: {0} rows refused by Revit and unticked.".format(
                len(rejected)))
            return

        for msg in failures.warnings:
            logger.warning(u"pyTable: Revit warning - {0}".format(msg))
        notes = []
        if skipped:
            for msg in skipped:
                logger.warning(u"pyTable: skipped - {0}".format(msg))
            notes.append(u"{0} could not be written and were skipped:\n\n{1}{2}".format(
                len(skipped), u"\n".join(skipped[:10]),
                u"\n..." if len(skipped) > 10 else u""))
        if failures.warnings:
            notes.append(u"Revit warnings (dismissed):\n\n" + u"\n".join(
                sorted(set(failures.warnings))[:10]))
        if notes:
            dlg.message(u"{0} changes applied.\n\n{1}".format(
                applied, u"\n\n".join(notes)), title="pyTable")

        self._diff_rows.Clear()
        self.FindName("confirm_btn").IsEnabled = False
        self.FindName("diff_summary_label").Text = "0 changes found"
        self.FindName("main_tabs").SelectedIndex = 0
        self.FindName("preview_tab").IsEnabled = False
        if skipped:
            self._set_status("Applied {0} changes, {1} skipped, see log.".format(
                applied, len(skipped)))
        else:
            self._set_status("Applied {0} parameter changes.".format(applied))

    def on_close(self, sender, args):
        self.Close()

    def on_menu_opened(self, sender, args):
        """menu_popup's own Opened event - fires as the hamburger dropdown
        becomes visible (Popup.IsOpen is bound to menu_btn.IsChecked in
        XAML, no manual open/close code needed)."""
        panel = self.FindName("menu_popup_panel")
        panel.Children.Clear()

        def item(label, fn):
            return self._make_menu_item(label, fn, self.FindName("menu_popup"))

        panel.Children.Add(item(u'\u2709  Email support', self._menu_support_click))
        # Vector GitHub mark rather than a glyph. make_icon bakes its colour in
        # at build time, which is fine here - this menu is rebuilt every time it
        # opens, so a theme change is picked up on the next open.
        panel.Children.Add(item(
            make_icon_with_label(
                'github', u'Report an issue on GitHub', icon_size=14,
                color=get_color(SCRIPT_DIR, 'text_primary', fallback='#F4FAFF')),
            self._menu_issue_click))
        panel.Children.Add(item(u'\u2139  About pyTable', self._menu_about_click))
        panel.Children.Add(self._make_menu_separator())
        panel.Children.Add(item(u'\u2615  Support this project and help us grow',
                                 self._menu_donate_click))

    def _make_menu_item(self, label, fn, popup):
        """StaysOpen='False' only auto-dismisses the popup on a click
        OUTSIDE it, so a click on one of its own items needs to close it
        explicitly here."""
        item = Button()
        item.Content = label
        try:
            item.Style = self.FindResource('MenuItemStyle')
        except Exception as ex:
            logger.warning('Failed to apply MenuItemStyle: {0}'.format(ex))

        def _click(sender, ev):
            popup.IsOpen = False
            fn(sender, ev)
        item.Click += _click
        return item

    def _make_menu_separator(self):
        sep = Rectangle()
        sep.Height = 1
        sep.Margin = Thickness(6, 4, 6, 4)
        try:
            sep.Fill = self.FindResource('LocalBrushMenuBorder')
        except Exception as ex:
            logger.warning('Failed to apply LocalBrushMenuBorder: {0}'.format(ex))
        return sep

    def _open_url(self, url, title=u''):
        """Open a URL in the default browser. The launch itself lives in
        Snippets._support.open_url; this only supplies pyTable's error
        reporting."""
        open_url(url, window=self, on_error=lambda msg: logger.error(msg))

    def _menu_support_click(self, sender, args):
        """Email support: open a pre-filled support email in the default mail
        client, addressed to Seed43 support, with the extension version and
        which app it came from already filled in."""
        self._open_url(support_mailto('pyTable', SCRIPT_DIR), title="Support")

    def _menu_issue_click(self, sender, args):
        """Report an issue: open a new GitHub issue, pre-filled with the app
        name, Seed43 version and Revit version."""
        self._open_url(github_issue_url('pyTable', SCRIPT_DIR),
                       title='Report an issue')

    ABOUT_URL = 'https://seed43.org/pytable/'

    def _menu_about_click(self, sender, args):
        """Open this tool's own page, the way pyTransmit and pyFilter do."""
        self._open_url(self.ABOUT_URL, title='About')

    def _menu_donate_click(self, sender, args):
        self._open_url('https://buymeacoffee.com/seed43', title='Support')

    def on_reset(self, sender, args):
        self._diff_rows.Clear()
        self.FindName("confirm_btn").IsEnabled = False
        self.FindName("main_tabs").SelectedIndex = 0
        self.FindName("preview_tab").IsEnabled = False
        self.FindName("diff_summary_label").Text = "0 changes found"
        self._set_status("Ready")

    # ── helpers ─────────────────────────────────────────────────────────

    def _set_status(self, text):
        self.FindName("status_label").Text = text


# ── DEVELOPMENT DISCLAIMER ──────────────────────────────────────────────
# Shown before the window opens, every launch. Deliberately not a "don't
# show again" tick: this tool writes changes back to the model, and the
# warning stops being a warning the moment it can be dismissed for good.
# Delete this block (and the call in __main__) once pyTable is production
# ready.
DISCLAIMER_TITLE = 'pyTable is under development'
DISCLAIMER_TEXT = (
    u'pyTable is a proof of concept and is still being built. It is not '
    u'ready for full production use - treat it as a testing tool for now.\n\n'
    u'It writes changes back to your model, so read the preview carefully '
    u'before confirming anything, and work on a model you can afford to '
    u'roll back.'
)


def show_disclaimer():
    """Never blocks the tool from opening - a failed dialog must not be the
    reason you cannot run pyTable at all."""
    try:
        dlg.message(DISCLAIMER_TEXT, title=DISCLAIMER_TITLE, ok_label='I understand')
    except Exception:
        pass


if __name__ == "__main__":
    show_disclaimer()
    window = MainWindow()
    window.ShowDialog()
