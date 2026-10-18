# -*- coding: utf-8 -*-
# """
# pyRameters' Parameters tab: the grid, its toolbar and its menus.
#
# The window owns the controls; this module drives the ones whose names
# start with p_. Reading and writing the model is doc_params' job; nothing
# here touches the Revit document except through it.
#
# Every edit is staged on a DocParam and shown in the first column. Apply
# writes them all in one transaction, so Revit's own Undo takes the whole
# batch back in one step.
#
# Target: IronPython 2 (pyRevit).
# """

from pyrevit import forms

from System.Collections.Generic import List
from System.ComponentModel import (GroupDescription, INotifyPropertyChanged,
                                   PropertyChangedEventArgs)
from System.Windows import Visibility
from System.Windows.Controls import ComboBox, GroupStyle
from System.Windows.Data import ListCollectionView

from Snippets._dialogs import choice, confirm, message

import doc_params as dp

__all__ = ["ParamsTab"]

# Column positions in pyRameters.xaml's p_grid. By index rather than x:Name,
# as the Shared file tab: DataGridColumn is not in the visual tree, so its
# name is not something to rely on.
P_COL_NAME = 1
P_COL_TYPE = 4
P_COL_GROUP = 5
P_COL_INST = 6
P_COL_VALUE = 7
P_COL_FORMULA = 8
P_COL_CATS = 9

INSTANCE = u"Instance"
TYPE = u"Type"

STATE_LABELS = {dp.NEW: u"New", dp.CHANGED: u"Changed", dp.DELETE: u"Delete",
                dp.EXISTING: u""}

FILTER_ALL = u"All parameters"
FILTER_CHANGES = u"Changes waiting"
FILTER_SHARED = u"Shared"
FILTER_BUILTIN = u"Built-in"


def _plural(count, word):
    return u"{0} {1}{2}".format(count, word, u"" if count == 1 else u"s")


# ── CHOICES FOR THE COMBO COLUMNS ───────────────────────────────────────────

class Option(object):
    """A dropdown entry: Token is the TypeId written, Label what is shown."""

    def __init__(self, token, label, group=u""):
        self._token = token
        self._label = label
        self._group = group

    @property
    def Token(self):
        return self._token

    @property
    def Label(self):
        return self._label

    @property
    def Group(self):
        return self._group

    def __str__(self):
        return self._label


class _ByGroup(GroupDescription):
    """Discipline headings in the type dropdown, read straight off Group."""

    def GroupNameFromItem(self, item, level, culture):
        return getattr(item, "Group", u"")


# ── ROW MODEL ───────────────────────────────────────────────────────────────

class DocParamRow(INotifyPropertyChanged):
    """One grid row over one DocParam.

    Setters write the DocParam and ask the tab to re-mark the row, so the
    first column always says what Apply would do. INotifyPropertyChanged by
    hand, as the Shared file tab's ParamRow: IronPython cannot declare a
    .NET event, and the add_/remove_ pair is what WPF binds against.
    """

    def __init__(self, param, tab):
        self.param = param
        self.tab = tab
        self._handlers = []

    def add_PropertyChanged(self, handler):
        self._handlers.append(handler)

    def remove_PropertyChanged(self, handler):
        if handler in self._handlers:
            self._handlers.remove(handler)

    def notify(self, *names):
        for name in names:
            args = PropertyChangedEventArgs(name)
            for handler in list(self._handlers):
                handler(self, args)

    def notify_all(self):
        self.notify("StateLabel", "Name", "Kind", "Discipline", "Spec", "Group",
                    "InstanceType", "Value", "Formula", "CategoriesText", "Guid")

    def _set(self, field, value):
        if getattr(self.param, field) == value:
            return
        setattr(self.param, field, value)
        self.tab.row_edited(self)

    @property
    def StateLabel(self):
        return STATE_LABELS.get(self.param.state, u"")

    @property
    def Kind(self):
        return self.param.kind

    @property
    def Discipline(self):
        return self.tab.discipline_for(self.param)

    @property
    def Guid(self):
        return self.param.guid

    @property
    def CategoriesText(self):
        return self.tab.categories_text(self.param.categories)

    def _get_name(self):
        return self.param.name

    def _set_name(self, value):
        self._set("name", (value or u"").strip())

    Name = property(_get_name, _set_name)

    def _get_spec(self):
        return self.param.spec

    def _set_spec(self, value):
        self._set("spec", value or u"")
        self.notify("Discipline")

    Spec = property(_get_spec, _set_spec)

    def _get_group(self):
        return self.param.group

    def _set_group(self, value):
        self._set("group", value or u"")

    Group = property(_get_group, _set_group)

    def _get_instance(self):
        return INSTANCE if self.param.is_instance else TYPE

    def _set_instance(self, value):
        self._set("is_instance", value == INSTANCE)

    InstanceType = property(_get_instance, _set_instance)

    def _get_value(self):
        return self.param.value

    def _set_value(self, value):
        self._set("value", value or u"")

    Value = property(_get_value, _set_value)

    def _get_formula(self):
        return self.param.formula

    def _set_formula(self, value):
        self._set("formula", (value or u"").strip())

    Formula = property(_get_formula, _set_formula)


# ── THE TAB ─────────────────────────────────────────────────────────────────

class ParamsTab(object):
    """Drives the p_ controls of the pyRameters window.

    window          the pyRameters window (owns the controls and the shared
                    file open in the other tab)
    doc             the Revit document whose parameters are shown
    rename_dialog   the window's RenameWindow class, passed in rather than
                    imported to keep this module free of the script
    token_label     the Shared file tab's DATATYPE token -> label function,
                    used for parameters added from the file before they
                    exist in the model
    """

    def __init__(self, window, doc, rename_dialog, token_label):
        self.w = window
        self.doc = doc
        self.is_family = doc.IsFamilyDocument
        self.rename_dialog = rename_dialog
        self.token_label = token_label
        self.rows = []
        self._loading = False
        self._categories = {}       # id -> name
        self._groups = []           # Option
        self._specs = []            # Option
        self._spec_by_token = {}
        self._new_labels = {}       # guid -> label, for rows from the file

        self._init_columns()
        self._init_filter()
        self._wire()
        self.load()

    # --- setup ---

    def _init_columns(self):
        g = self.w.p_grid
        g.Columns[P_COL_INST].ItemsSource = List[str]([INSTANCE, TYPE])
        family_only = Visibility.Visible if self.is_family else Visibility.Collapsed
        project_only = Visibility.Collapsed if self.is_family else Visibility.Visible
        g.Columns[P_COL_VALUE].Visibility = family_only
        g.Columns[P_COL_FORMULA].Visibility = family_only
        g.Columns[P_COL_CATS].Visibility = project_only
        # Menu entries that mean nothing in this kind of document go, rather
        # than sitting there refusing.
        for name in ("p_cats_btn", "p_ctx_cats_btn"):
            self.w.FindName(name).Visibility = project_only
        for name in ("p_add_family_btn", "p_to_shared_btn", "p_to_family_btn",
                     "p_rename_btn", "p_ctx_rename_btn", "p_dup_btn",
                     "p_ctx_dup_btn"):
            self.w.FindName(name).Visibility = family_only

    def _init_filter(self):
        own = dp.KIND_FAMILY if self.is_family else dp.KIND_PROJECT
        items = List[object]()
        for label in (FILTER_ALL, FILTER_CHANGES, FILTER_SHARED, own,
                      FILTER_BUILTIN if self.is_family else None):
            if label:
                items.Add(label)
        self._loading = True
        self.w.p_filter_cb.ItemsSource = items
        self.w.p_filter_cb.SelectedIndex = 0
        self._loading = False

    def _wire(self):
        w = self.w
        w.p_filter_cb.SelectionChanged += lambda s, a: self.refresh_view()
        w.p_search_tb.TextChanged += lambda s, a: self.refresh_view()
        w.p_clear_btn.Click += self.clear_filters
        w.p_rename_btn.Click += self.rename
        w.p_revert_btn.Click += self.revert
        w.p_grid.BeginningEdit += self._beginning_edit
        w.p_grid.PreparingCellForEdit += self._preparing_cell_for_edit
        w.p_grid.PreviewMouseRightButtonUp += self._right_click
        actions = (
            ("p_add_shared_btn", self.add_from_file),
            ("p_add_family_btn", self.add_family_parameter),
            ("p_dup_btn", self.duplicate), ("p_ctx_dup_btn", self.duplicate),
            ("p_ctx_rename_btn", self.rename),
            ("p_instance_btn", self.make_instance),
            ("p_ctx_instance_btn", self.make_instance),
            ("p_typemode_btn", self.make_type),
            ("p_ctx_typemode_btn", self.make_type),
            ("p_group_btn", self.set_group), ("p_ctx_group_btn", self.set_group),
            ("p_cats_btn", self.set_categories),
            ("p_ctx_cats_btn", self.set_categories),
            ("p_to_shared_btn", self.convert_to_shared),
            ("p_to_family_btn", self.convert_to_family),
            ("p_del_btn", self.delete), ("p_ctx_del_btn", self.delete),
        )
        for name, action in actions:
            w.FindName(name).Click += self._menu_action(action)
        for toggle in (w.p_add_menu_btn, w.p_batch_menu_btn):
            toggle.PreviewMouseLeftButtonDown += self._dropdown_preview_down

    def _menu_action(self, action):
        def run(sender, args):
            self.close_menus()
            action(sender, args)
        return run

    def close_menus(self):
        self.w.p_add_menu_btn.IsChecked = False
        self.w.p_batch_menu_btn.IsChecked = False
        self.w.p_row_menu_popup.IsOpen = False

    def _dropdown_preview_down(self, sender, args):
        """A re-click on an open dropdown only closes it, as the header menu."""
        if sender == self.w.p_add_menu_btn:
            popup = self.w.p_add_menu_popup
        else:
            popup = self.w.p_batch_menu_popup
        if popup.IsOpen:
            popup.IsOpen = False
            sender.IsChecked = False
            args.Handled = True

    # --- reading ---

    def load(self):
        """Read the document's parameters, dropping anything staged."""
        params = dp.read_parameters(self.doc)
        self._groups = [Option(tid, label)
                        for label, tid in dp.group_options(self.doc)]
        self._specs = [Option(tid, label, disc)
                       for disc, label, tid in dp.spec_options()]
        self._spec_by_token = dict((o.Token, o) for o in self._specs)
        # A group or type in use but not in the lists still has to show.
        known_groups = set(o.Token for o in self._groups)
        for p in params:
            if p.group not in known_groups:
                self._groups.append(Option(p.group, dp.group_label(p.group)))
                known_groups.add(p.group)
            if p.spec and p.spec not in self._spec_by_token:
                extra = Option(p.spec, dp.spec_label(p.spec),
                               dp.discipline_label(p.spec))
                self._specs.append(extra)
                self._spec_by_token[p.spec] = extra
        self._groups.sort(key=lambda o: o.Label.lower())
        if not self.is_family:
            self._categories = dict((cid, name) for name, cid, _
                                    in dp.bindable_categories(self.doc))

        self.w.p_grid.Columns[P_COL_GROUP].ItemsSource = List[object](self._groups)
        view = ListCollectionView(List[object](self._specs))
        view.GroupDescriptions.Add(_ByGroup())
        self.w.p_grid.Columns[P_COL_TYPE].ItemsSource = view

        self.rows = [DocParamRow(p, self) for p in params]
        self._new_labels = {}
        self.refresh_view()
        self.w.set_status(u"{0} in {1}.".format(
            _plural(len(params), u"parameter"), self.doc.Title))

    # --- the grid ---

    def visible_rows(self):
        needle = (self.w.p_search_tb.Text or u"").strip().lower()
        chosen = self.w.p_filter_cb.SelectedItem
        out = []
        for row in self.rows:
            p = row.param
            if chosen == FILTER_CHANGES and p.state == dp.EXISTING:
                continue
            if chosen in (FILTER_SHARED, dp.KIND_FAMILY, dp.KIND_PROJECT,
                          FILTER_BUILTIN) and p.kind != chosen:
                continue
            if needle:
                hay = u" ".join((p.name, dp.group_label(p.group),
                                 self.spec_label_for(p), p.guid)).lower()
                if needle not in hay:
                    continue
            out.append(row)
        return out

    def refresh_view(self, keep=None):
        if self._loading:
            return
        self.w.p_grid.ItemsSource = List[object](self.visible_rows())
        if keep:
            for row in keep:
                if row in self.rows:
                    self.w.p_grid.SelectedItems.Add(row)
        self.refresh_title()

    def refresh_title(self):
        if self.w.params_panel.Visibility != Visibility.Visible:
            return      # the Shared file tab owns the subtitle while shown
        waiting = self.pending()
        kind = u"family" if self.is_family else u"project"
        tail = u"  -  {0} waiting".format(_plural(len(waiting), u"change")) \
            if waiting else u""
        self.w.subtitle_tb.Text = u"  |  {0} ({1}){2}".format(
            self.doc.Title, kind, tail)

    def clear_filters(self, sender=None, args=None):
        self._loading = True
        self.w.p_search_tb.Text = u""
        self.w.p_filter_cb.SelectedIndex = 0
        self._loading = False
        self.refresh_view()

    def row_edited(self, row):
        row.param.refresh_state()
        row.notify_all()
        self.refresh_title()

    def pending(self):
        return [r for r in self.rows if r.param.state != dp.EXISTING]

    def selected(self):
        rows = [r for r in self.w.p_grid.SelectedItems]
        if not rows:
            self.w.set_status(u"Select one or more rows first.")
        return rows

    # --- labels for the row model ---

    def spec_label_for(self, p):
        if p.state == dp.NEW and p.kind == dp.KIND_SHARED:
            return self._new_labels.get(p.guid, u"")
        option = self._spec_by_token.get(p.spec)
        return option.Label if option else dp.spec_label(p.spec)

    def discipline_for(self, p):
        if p.state == dp.NEW and p.kind == dp.KIND_SHARED:
            return u""      # set by the shared definition once it is added
        option = self._spec_by_token.get(p.spec)
        return option.Group if option else dp.discipline_label(p.spec)

    def categories_text(self, ids):
        names = sorted(self._categories.get(i, u"?") for i in ids)
        if len(names) > 4:
            return u"{0}, +{1} more".format(u", ".join(names[:3]), len(names) - 3)
        return u", ".join(names)

    # --- editing guard ---

    def _beginning_edit(self, sender, args):
        """Refuse an edit the API cannot carry out, and say why."""
        row = args.Row.Item
        p = row.param
        index = self.w.p_grid.Columns.IndexOf(args.Column)
        reason = None
        if index == P_COL_NAME:
            reason = p.why_not_rename(self.is_family) if p.state != dp.NEW \
                else (None if p.kind == dp.KIND_FAMILY else
                      p.why_not_rename(self.is_family))
        elif index == P_COL_TYPE:
            reason = p.why_not_type()
        elif index in (P_COL_GROUP, P_COL_INST):
            reason = p.why_not_change()
        elif index == P_COL_VALUE:
            reason = p.why_not_value(self.is_family)
        elif index == P_COL_FORMULA:
            reason = p.why_not_formula(self.is_family)
        if reason:
            args.Cancel = True
            self.w.set_status(reason)

    def _preparing_cell_for_edit(self, sender, args):
        """Discipline headings in the type dropdown, as the Shared file tab."""
        editor = args.EditingElement
        if (self.w.p_grid.Columns.IndexOf(args.Column) == P_COL_TYPE
                and isinstance(editor, ComboBox)
                and editor.GroupStyle.Count == 0):
            style = self.w.TryFindResource("TypeGroupStyle")
            if isinstance(style, GroupStyle):
                editor.GroupStyle.Add(style)

    def _right_click(self, sender, args):
        from System.Windows.Controls import DataGridRow
        from System.Windows.Media import VisualTreeHelper
        node = args.OriginalSource
        while node is not None and not isinstance(node, DataGridRow):
            try:
                node = VisualTreeHelper.GetParent(node)
            except Exception:
                node = None
        if node is None:
            return
        if not node.IsSelected:
            self.w.p_grid.SelectedItems.Clear()
            node.IsSelected = True
        self.w.p_row_menu_popup.IsOpen = True
        args.Handled = True

    # --- batch helpers ---

    def _each(self, rows, check, apply, done_word):
        """Apply to the rows check allows; report what was done and refused."""
        done, refused = [], []
        for row in rows:
            reason = check(row.param)
            if reason:
                refused.append(reason)
                continue
            apply(row.param)
            row.param.refresh_state()
            row.notify_all()
            done.append(row)
        self.refresh_title()
        text = u"{0} {1}.".format(_plural(len(done), u"parameter"), done_word)
        if refused:
            text += u" {0} skipped: {1}".format(len(refused), refused[0])
        self.w.set_status(text)
        return done

    def _unique_name(self, base):
        taken = set(r.param.name.lower() for r in self.rows)
        name, n = base, 2
        while name.lower() in taken:
            name = u"{0} {1}".format(base, n)
            n += 1
        return name

    def _default_group(self):
        for o in self._groups:
            if o.Label.lower() == u"data":
                return o.Token
        return self._groups[0].Token if self._groups else u""

    def _add_rows(self, params):
        new_rows = [DocParamRow(p, self) for p in params]
        self.rows.extend(new_rows)
        self.clear_filters()
        for row in new_rows:
            self.w.p_grid.SelectedItems.Add(row)
        if new_rows:
            self.w.p_grid.ScrollIntoView(new_rows[-1])
        return new_rows

    # --- actions ---

    def add_family_parameter(self, sender=None, args=None):
        p = dp.DocParam(dp.KIND_FAMILY, self._unique_name(u"New parameter"),
                        group=self._default_group(), spec=dp.text_spec(),
                        is_instance=False, state=dp.NEW)
        self._add_rows([p])
        self.w.set_status(u"New family parameter added below. Name it and pick "
                          u"its type in the grid; nothing is written until "
                          u"Apply.")

    def add_from_file(self, sender=None, args=None):
        """Stage shared parameters from the file open in the Shared file tab."""
        path = self.w.shared_file_on_disk()
        if not path:
            return
        in_doc = set(r.param.guid.lower() for r in self.rows if r.param.guid)
        rows_by_label = {}
        for data in self.w.spf.params:
            guid = (data.get("GUID") or u"").strip()
            if not guid or guid.lower() in in_doc:
                continue
            label = u"{0}  -  {1}  ({2})".format(
                data.get("NAME", u""), self.w.group_name_for(data.get("GROUP")),
                self.token_label(data.get("DATATYPE", u"")))
            rows_by_label[label] = data
        if not rows_by_label:
            message(u"Every parameter in the file is already in {0}."
                    .format(self.doc.Title), title=u"Add from the shared file")
            return
        picked = forms.SelectFromList.show(
            sorted(rows_by_label, key=lambda s: s.lower()),
            title=u"Add from the shared file", button_name=u"Add",
            multiselect=True)
        if not picked:
            return

        is_instance = True
        categories = ()
        if not self.is_family:
            categories = self._pick_categories((), len(picked))
            if categories is None:
                return
        mode = choice(u"Add {0} as instance or type parameters?".format(
                      _plural(len(picked), u"parameter")),
                      [("cancel", u"Cancel"), ("type", TYPE),
                       ("instance", INSTANCE)], title=u"Instance or type")
        if mode not in ("type", "instance"):
            return
        is_instance = mode == "instance"

        params = []
        for label in picked:
            data = rows_by_label[label]
            guid = data.get("GUID", u"").strip()
            self._new_labels[guid] = self.token_label(data.get("DATATYPE", u""))
            params.append(dp.DocParam(
                dp.KIND_SHARED, data.get("NAME", u""), guid=guid,
                group=self._default_group(), is_instance=is_instance,
                categories=categories, state=dp.NEW))
        self._add_rows(params)
        self.w.set_status(u"{0} staged. Adjust their group in the grid, then "
                          u"Apply.".format(_plural(len(params), u"parameter")))

    def duplicate(self, sender=None, args=None):
        rows = self.selected()
        if not rows:
            return
        params = []
        for row in rows:
            p = row.param
            if p.kind == dp.KIND_BUILTIN:
                continue
            params.append(dp.DocParam(
                dp.KIND_FAMILY, self._unique_name(p.name + u" copy"),
                group=p.group, spec=p.spec, is_instance=p.is_instance,
                formula=p.formula, value=p.value, state=dp.NEW))
        self._add_rows(params)
        self.w.set_status(u"{0} staged as new family parameters. A shared "
                          u"parameter is copied as a family one: a copy would "
                          u"need a new GUID.".format(
                              _plural(len(params), u"copy")))

    def rename(self, sender=None, args=None):
        rows = self.selected()
        if not rows:
            return
        allowed = [r for r in rows if r.param.state == dp.NEW
                   and r.param.kind == dp.KIND_FAMILY
                   or r.param.why_not_rename(self.is_family) is None]
        if not allowed:
            self.w.set_status(rows[0].param.why_not_rename(self.is_family)
                              or u"Nothing selected can be renamed.")
            return
        dialog = self.rename_dialog(self.w, [r.param.name for r in allowed])
        dialog.ShowDialog()
        if not dialog.result:
            return
        prefix, suffix, find, with_ = dialog.result
        if not (prefix or suffix or find):
            return
        for row in allowed:
            name = row.param.name.replace(find, with_) if find else row.param.name
            row.param.name = prefix + name + suffix
            row.param.refresh_state()
            row.notify_all()
        self.refresh_title()
        skipped = len(rows) - len(allowed)
        self.w.set_status(u"{0} renamed{1}.".format(
            _plural(len(allowed), u"parameter"),
            u"; {0} skipped (shared or built-in)".format(skipped)
            if skipped else u""))

    def make_instance(self, sender=None, args=None):
        rows = self.selected()
        if rows:
            self._each(rows, lambda p: p.why_not_change(),
                       lambda p: setattr(p, "is_instance", True),
                       u"set to instance")

    def make_type(self, sender=None, args=None):
        rows = self.selected()
        if rows:
            self._each(rows, lambda p: p.why_not_change(),
                       lambda p: setattr(p, "is_instance", False),
                       u"set to type")

    def set_group(self, sender=None, args=None):
        rows = self.selected()
        if not rows:
            return
        by_label = dict((o.Label, o.Token) for o in self._groups)
        picked = forms.SelectFromList.show(
            sorted(by_label, key=lambda s: s.lower()),
            title=u"Group under", button_name=u"Apply", multiselect=False)
        if not picked:
            return
        token = by_label[picked]
        self._each(rows, lambda p: p.why_not_change(),
                   lambda p: setattr(p, "group", token),
                   u"moved to {0}".format(picked))

    def _pick_categories(self, current, count):
        """Tick-list of bindable categories. Returns a tuple of ids, or None."""
        items = []
        for cid, name in sorted(self._categories.items(),
                                key=lambda kv: kv[1].lower()):
            items.append(forms.TemplateListItem(
                _CategoryChoice(cid, name), checked=cid in current))
        picked = forms.SelectFromList.show(
            items, title=u"Categories for {0}".format(
                _plural(count, u"parameter")),
            button_name=u"Use these", multiselect=True)
        if picked is None:
            return None
        ids = tuple(sorted(c.cid for c in picked))
        if not ids:
            message(u"A project parameter needs at least one category.",
                    title=u"Categories")
            return None
        return ids

    def set_categories(self, sender=None, args=None):
        rows = self.selected()
        if not rows:
            return
        rows = [r for r in rows if r.param.why_not_categories(self.is_family)
                is None]
        if not rows:
            self.w.set_status(u"Nothing selected takes categories.")
            return
        # Pre-tick what every selected row already shares, so a single row
        # opens on its own categories.
        common = set(rows[0].param.categories)
        for row in rows[1:]:
            common &= set(row.param.categories)
        ids = self._pick_categories(common, len(rows))
        if ids is None:
            return
        self._each(rows, lambda p: p.why_not_categories(self.is_family),
                   lambda p: setattr(p, "categories", ids),
                   u"bound to {0}".format(_plural(len(ids), u"category")))

    def convert_to_shared(self, sender=None, args=None):
        rows = self.selected()
        if not rows:
            return
        if len(rows) != 1 or rows[0].param.kind != dp.KIND_FAMILY:
            self.w.set_status(u"Select one family parameter to convert: each "
                              u"needs its own shared definition.")
            return
        path = self.w.shared_file_on_disk()
        if not path:
            return
        by_label = {}
        for data in self.w.spf.params:
            label = u"{0}  -  {1}".format(data.get("NAME", u""),
                                          self.w.group_name_for(data.get("GROUP")))
            by_label[label] = data
        picked = forms.SelectFromList.show(
            sorted(by_label, key=lambda s: s.lower()),
            title=u"Convert to which shared parameter?", button_name=u"Convert",
            multiselect=False)
        if not picked:
            return
        data = by_label[picked]
        p = rows[0].param
        p.convert_to = dp.KIND_SHARED
        p.convert_guid = data.get("GUID", u"").strip()
        p.name = data.get("NAME", p.name)
        p.refresh_state()
        rows[0].notify_all()
        self.refresh_title()
        self.w.set_status(u"{0} will be replaced by the shared parameter on "
                          u"Apply; its values carry across.".format(p.name))

    def convert_to_family(self, sender=None, args=None):
        rows = self.selected()
        if rows:
            self._each(rows,
                       lambda p: None if p.kind == dp.KIND_SHARED and
                       p.state != dp.NEW else u"Only shared parameters convert "
                       u"to family ones.",
                       lambda p: setattr(p, "convert_to", dp.KIND_FAMILY),
                       u"set to convert to family parameters")

    def delete(self, sender=None, args=None):
        """Mark for deletion; a second Delete unmarks. New rows just go."""
        rows = self.selected()
        if not rows:
            return
        removed = marked = unmarked = 0
        for row in rows:
            p = row.param
            if p.kind == dp.KIND_BUILTIN:
                continue
            if p.state == dp.NEW:
                self.rows.remove(row)
                removed += 1
            elif p.state == dp.DELETE:
                p.state = dp.EXISTING
                p.refresh_state()
                unmarked += 1
            else:
                p.state = dp.DELETE
                marked += 1
            row.notify_all()
        self.refresh_view()
        parts = []
        if marked:
            parts.append(u"{0} marked for deletion".format(_plural(marked, u"parameter")))
        if unmarked:
            parts.append(u"{0} unmarked".format(unmarked))
        if removed:
            parts.append(u"{0} new row{1} removed".format(
                removed, u"" if removed == 1 else u"s"))
        self.w.set_status((u"; ".join(parts) or u"Built-in parameters cannot be "
                           u"deleted") + u".")

    def revert(self, sender=None, args=None):
        waiting = self.pending()
        if waiting and not confirm(
                u"Throw away {0} and read the parameters again?".format(
                    _plural(len(waiting), u"waiting change")),
                title=u"Revert", yes=u"Revert"):
            return
        self.load()

    # --- apply ---

    def commit_edit(self):
        """Finish a cell still being typed in, so Apply sees its value.

        The first CommitEdit ends the cell edit, the second the row's.
        """
        try:
            self.w.p_grid.CommitEdit()
            self.w.p_grid.CommitEdit()
        except Exception:
            pass

    def problems(self):
        """Reasons Apply would fail before it starts, by parameter."""
        out = []
        names = {}
        for row in self.rows:
            p = row.param
            if p.state == dp.DELETE:
                continue
            key = p.name.lower()
            names[key] = names.get(key, 0) + 1
            if p.state == dp.NEW:
                if not p.name:
                    out.append(u"A new parameter has no name.")
                if not self.is_family and not p.categories:
                    out.append(u"{0} has no categories.".format(p.name))
        for key, count in names.items():
            if count > 1:
                out.append(u"More than one parameter is called {0}.".format(key))
        return out

    def apply(self):
        waiting = self.pending()
        if not waiting:
            self.w.set_status(u"Nothing waiting to apply.")
            return
        issues = self.problems()
        if issues:
            message(u"Fix these first:\n\n" + u"\n".join(issues[:12]),
                    title=u"Apply")
            return
        if not confirm(u"Apply {0} to {1}? Revit's Undo takes them back in one "
                       u"step.".format(_plural(len(waiting), u"change"),
                                       self.doc.Title),
                       title=u"Apply", yes=u"Apply"):
            return
        try:
            report = dp.apply_changes(self.doc, [r.param for r in self.rows],
                                      shared_file=self.w.spf.path)
        except Exception as ex:
            message(u"Nothing was changed.\n\n{0}".format(ex), title=u"Apply")
            return
        self.load()
        text = u"{0} applied.".format(_plural(len(report.done), u"change"))
        if report.failed:
            lines = [u"{0}: {1}".format(n, why) for n, why in report.failed[:12]]
            message(text + u" {0} could not be:\n\n{1}".format(
                len(report.failed), u"\n".join(lines)), title=u"Apply")
        self.w.set_status(text)


class _CategoryChoice(object):
    """A category in the tick-list: shows its name, remembers its id."""

    def __init__(self, cid, name):
        self.cid = cid
        self.name = name

    def __str__(self):
        return self.name
