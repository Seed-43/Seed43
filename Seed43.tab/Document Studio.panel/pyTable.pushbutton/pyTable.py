# -*- coding: utf-8 -*-
"""
pyTable - export model/schedule parameters to Excel or ODS, edit externally,
re-import with a diff preview before anything is written back to the model.

𝐒𝐄𝐄𝐃𝟒𝟑
"""

__title__ = "pyTable"
__author__ = "𝐒𝐄𝐄𝐃𝟒𝟑"
__doc__ = "Export Revit parameters to Excel/ODS and import edits back with a preview step."

import os
import sys
import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import MessageBoxButton, MessageBoxImage, Thickness
from System.Windows.Controls import ListBoxItem, Button
from System.Windows.Shapes import Rectangle
from System.Collections.ObjectModel import ObservableCollection

import threading

from Autodesk.Revit.DB import Transaction

from pyrevit import revit, DB, script, forms

from Snippets.seed43_theme import apply_seed43_palette, apply_seed43_dimensions
from Snippets._selection import get_element_type
from Snippets._parameters import (
    get_param_display_value,
    set_param_from_display_value,
    get_categories_with_elements,
    get_elements_by_category,
    get_schedules,
    get_param_dropdown_options,
)

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'tools'))
from pytable_io import write_workbook, read_workbook, HIDDEN_COLUMNS

SCRIPT_DIR = os.path.dirname(__file__)


def _find_seed43_version():
    """Walk up from this pushbutton to Seed43.extension/version.txt and
    return just the version string (its first line). Returns 'unknown'
    if the file can't be found or read."""
    folder = SCRIPT_DIR
    for _ in range(6):
        candidate = os.path.join(folder, 'version.txt')
        if os.path.isfile(candidate):
            try:
                with open(candidate, 'r') as f:
                    return f.readline().strip()
            except Exception:
                return 'unknown'
        folder = os.path.dirname(folder)
    return 'unknown'
doc = revit.doc
logger = script.get_logger()

TYPE_SUFFIX = u" (Type)"


def _encode_display(raw_name, kind):
    return raw_name + TYPE_SUFFIX if kind == "type" else raw_name


def _decode_display(display_name):
    if display_name.endswith(TYPE_SUFFIX):
        return display_name[:-len(TYPE_SUFFIX)], "type"
    return display_name, "instance"


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

        self._pending_export_columns = []   # list of {"key","kind","readonly"}
        self._pending_export_rows = []       # list of dicts keyed by column + hidden keys
        self._diff_rows = ObservableCollection[object]()

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

        diff_grid = self.FindName("diff_grid")
        diff_grid.ItemsSource = self._diff_rows

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

    def _populate_categories(self):
        """List every category that has at least one element in the model."""
        categories_list = self.FindName("categories_list")
        for cat in get_categories_with_elements(doc):
            item = ListBoxItem()
            item.Content = cat.Name
            item.Tag = cat.Id
            categories_list.Items.Add(item)

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
            elements = get_elements_by_category(doc, cat_id)
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
            for i in range(definition.GetFieldCount()):
                field = definition.GetField(i)
                raw_name = field.GetName()
                names[_encode_display(raw_name, "instance")] = (raw_name, "instance")

        self._refresh_available_list("available_params_schedules", "selected_params_schedules", names)

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

        for cat_id in self._get_checked_category_ids():
            elements = get_elements_by_category(doc, cat_id)
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

                    if param is None or not param.HasValue:
                        # Sentinel: this parameter doesn't exist / has no
                        # value on this particular element - distinct from
                        # a genuinely blank string value.
                        row[display_text] = None
                        continue

                    columns_seen[display_text]["readonly"] = param.IsReadOnly
                    row[display_text] = get_param_display_value(param)

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
                    retry = forms.alert(
                        "Could not save - the file may already be open in "
                        "Excel or LibreOffice.\n\nPlease close it, then "
                        "click Yes to retry, or No to cancel.\n\n"
                        "File: {0}\n\nError: {1}".format(filepath, ex),
                        title="File Locked: Please Close the File",
                        ok=False, yes=True, no=True)
                    if not retry:
                        self._set_status("Export cancelled.")
                        return
                else:
                    logger.error("pyTable export failed after {0} attempts: {1}".format(
                        MAX_SAVE_RETRIES, ex))
                    forms.alert(
                        "Could not save after {0} attempts:\n{1}".format(MAX_SAVE_RETRIES, ex),
                        title="pyTable")
                    self._set_status("Export failed after {0} attempts - see log.".format(MAX_SAVE_RETRIES))
                    return

        if not saved:
            return

        if not os.path.exists(filepath):
            forms.alert(
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

        try:
            headers, imported_rows = read_workbook(filepath)
        except Exception as ex:
            logger.error("pyTable import (read) failed: {0}".format(ex))
            forms.alert("Could not read file:\n{0}".format(ex), title="pyTable")
            self._set_status("Import failed - see log.")
            return

        self._diff_rows.Clear()

        changed_count = 0
        seen_type_diffs = set()   # (type_eid, raw_name) - only show one row per type param

        for row in imported_rows:
            eid = row.get("_eid")
            if eid is None:
                continue
            element = doc.GetElement(DB.ElementId(int(eid)))
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
                    element_label = "Type: {0} (Id {1})".format(
                        getattr(write_target, "Name", "?"), write_eid)
                else:
                    write_target = element
                    write_eid = eid
                    element_label = "Id {0}".format(eid)

                param = write_target.LookupParameter(raw_name) if write_target is not None else None
                if param is None or param.IsReadOnly:
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

    def on_confirm(self, sender, args):
        rows_to_apply = [r for r in self._diff_rows if r.Include]
        if not rows_to_apply:
            self._set_status("Nothing checked to apply.")
            return

        t = Transaction(doc, "pyTable: Apply parameter changes")
        t.Start()
        applied = 0
        failed = []
        try:
            for row in rows_to_apply:
                target = doc.GetElement(DB.ElementId(int(row._write_eid)))
                if target is None:
                    continue
                param = target.LookupParameter(row.ParamName)
                if param is None or param.IsReadOnly:
                    continue
                try:
                    set_param_from_display_value(param, row.NewValue)
                    applied += 1
                except Exception as row_ex:
                    failed.append("{0} ({1}): {2}".format(row.ElementLabel, row.ParamName, row_ex))
            t.Commit()
        except Exception as ex:
            t.RollBack()
            logger.error("pyTable import failed: {0}".format(ex))
            forms.alert("Applying changes failed, nothing was written:\n{0}".format(ex), title="pyTable")
            self._set_status("Import failed, no changes were made - see log.")
            return

        if failed:
            for msg in failed:
                logger.warning("pyTable: skipped - {0}".format(msg))

        self._diff_rows.Clear()
        self.FindName("confirm_btn").IsEnabled = False
        self.FindName("diff_summary_label").Text = "0 changes found"
        if failed:
            self._set_status("Applied {0} changes, {1} skipped - see log.".format(applied, len(failed)))
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

        panel.Children.Add(item(u'\u2753  Support', self._menu_support_click))
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

    _last_url_open_time = 0.0

    def _open_url(self, url, title=u''):
        """Open a URL in the default browser without blocking the UI
        thread - os.startfile skips the cmd.exe shell-wrapper subprocess
        launch that can otherwise hang for a while from inside Revit's
        process, and running it on a background thread means even a slow
        launch can never freeze the window."""
        import time
        now_t = time.time()
        if now_t - self._last_url_open_time < 2.0:
            return
        self._last_url_open_time = now_t

        def _launch():
            try:
                os.startfile(url)
            except Exception:
                try:
                    import subprocess
                    subprocess.Popen(['cmd', '/c', 'start', '', url])
                except Exception as ex:
                    logger.error("Could not open browser: {0}".format(ex))

        threading.Thread(target=_launch).start()

    def _menu_support_click(self, sender, args):
        """Support: open a pre-filled support email in the default mail
        client, addressed to Seed43 support, with the extension version
        and which app it came from already filled in."""
        version = _find_seed43_version()
        subject = "pyTable Support Ticket"
        body = (
            "Hi Seed43 Team,\n\n"
            "Support Request\n\n"
            "App: pyTable\n"
            "Seed43 Version: {0}\n\n"
            "Please describe your issue below:\n\n"
        ).format(version)
        try:
            import urllib
            quote = urllib.quote
        except ImportError:
            import urllib.parse
            quote = urllib.parse.quote
        mailto = "mailto:{0}?subject={1}&body={2}".format(
            'support@seed43.org', quote(subject), quote(body))
        self._open_url(mailto, title="Support")

    def _menu_about_click(self, sender, args):
        forms.alert(
            'pyTable\n'
            'Part of the Seed43 pyRevit Extension.\n\n'
            'Exports model/schedule parameters to Excel or ODS, edit '
            'externally, reimport with a diff preview before writing '
            'changes back to Revit.',
            title='About pyTable')

    def _menu_donate_click(self, sender, args):
        self._open_url('https://buymeacoffee.com/seed43', title='Support')

    def on_reset(self, sender, args):
        self._diff_rows.Clear()
        self.FindName("confirm_btn").IsEnabled = False
        self.FindName("preview_tab").IsEnabled = False
        self.FindName("diff_summary_label").Text = "0 changes found"
        self._set_status("Ready")

    # ── helpers ─────────────────────────────────────────────────────────

    def _set_status(self, text):
        self.FindName("status_label").Text = text


if __name__ == "__main__":
    window = MainWindow()
    window.ShowDialog()
