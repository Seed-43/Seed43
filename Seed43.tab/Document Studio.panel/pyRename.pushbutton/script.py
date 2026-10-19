# -*- coding: utf-8 -*-
# "pyRename"
# "𝐒𝐄𝐄𝐃𝟒𝟑"
# """
# Batch rename sheet numbers and sheet names, GNOME Files style: pick the
# sheets, set a rule, watch the whole list update live, rename once you like
# what you see.
#
# Sheet number and sheet name each get their own rule, so a renumber and a
# retitle can happen in the same pass.
# """

# ── IMPORTS ────────────────────────────────────────────────────────────────

import os
import re
import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Collections.ObjectModel import ObservableCollection

from Autodesk.Revit.DB import Transaction, ViewSheet, FilteredElementCollector
from Autodesk.Revit.UI import DockablePane, DockablePanes

from pyrevit import revit, script, forms

from Snippets.seed43_theme import apply_seed43_palette, apply_seed43_dimensions
from Snippets import _dialogs as dlg

__all__ = ["MainWindow", "expand_template", "apply_rule"]


# ── CONSTANTS ──────────────────────────────────────────────────────────────

SCRIPT_DIR = os.path.dirname(__file__)

doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()

# Revit rejects these in a sheet number or a sheet name. Validating up front
# is what lets the commit stay strict instead of retrying with mangled values
# the way the find/replace tools this one replaces had to.
INVALID_CHARS = u"\\:{}[]|;<>?`~"

# Rule modes, matching the ComboBox item order in pyRename.xaml.
MODE_KEEP = 0
MODE_FIND_REPLACE = 1
MODE_TEMPLATE = 2

# Case modes, matching the ComboBox item order in pyRename.xaml.
CASE_KEEP = 0
CASE_UPPER = 1
CASE_LOWER = 2
CASE_TITLE = 3

# Splits a template into literal text and tokens. `[[` / `]]` are escapes for
# a literal bracket - they survive expansion but will then fail validation,
# since brackets are not legal in a sheet number or name. Kept anyway so an
# escaped bracket reports a clear error instead of silently expanding as a
# malformed token.
_TOKEN_RE = re.compile(r"\[\[|\]\]|\[[^\[\]]*\]")


# ── HELPERS ────────────────────────────────────────────────────────────────

def safe_text(value):
    """Return value as unicode text, with None becoming an empty string.

    Revit hands back .NET strings and, for an unset name, None - both have to
    land as something the WPF grid and the rule engine can treat as text.
    """
    if value is None:
        return u""
    try:
        return unicode(value)
    except Exception:
        return u"{}".format(value)


def parse_int(text, default):
    """Return text as an int, falling back to default on anything unparseable.

    The counter fields are free-text rather than spinners, so a half-typed
    value like "-" or "" has to leave the preview working, not raise.
    """
    try:
        return int(safe_text(text).strip())
    except (ValueError, TypeError):
        return default


def title_case(text):
    """Capitalise the first letter of each word, lowercase the rest.

    Not str.title(): that also capitalises after an apostrophe, turning
    "Owner's Room" into "Owner'S Room".
    """
    out = []
    new_word = True
    for char in text:
        if char.isspace() or char in u"-_/":
            new_word = True
            out.append(char)
        elif new_word:
            out.append(char.upper())
            new_word = False
        else:
            out.append(char.lower())
    return u"".join(out)


def natural_key(text):
    """Sort key that orders TD-9 before TD-10 rather than after it."""
    parts = re.split(r"(\d+)", safe_text(text))
    return [int(p) if p.isdigit() else p.lower() for p in parts]


def invalid_chars_in(text):
    """Return the sorted set of Revit-illegal characters found in text."""
    return sorted(set(c for c in text if c in INVALID_CHARS))


# ── CORE LOGIC ─────────────────────────────────────────────────────────────

# --- Rule engine ---

def expand_template(template, number, name, counter):
    """Expand template tokens against one sheet.

    Tokens (case-insensitive):
        [N]                 the sheet's current number
        [NAME]              the sheet's current name
        [#], [##], [###]    the counter, zero-padded to the token's width

    Anything else in brackets is left exactly as typed, so an unrecognised
    token shows up in the preview instead of vanishing.

    Args:
        template (unicode): the raw template text.
        number (unicode): the sheet's current number.
        name (unicode): the sheet's current name.
        counter (int): this row's sequence value.

    Returns:
        unicode: the expanded result.
    """
    def _replace(match):
        token = match.group(0)
        if token == u"[[":
            return u"["
        if token == u"]]":
            return u"]"

        inner = token[1:-1]
        key = inner.upper()
        if key == u"N":
            return number
        if key == u"NAME":
            return name
        if inner and inner.strip(u"#") == u"":
            return safe_text(counter).rjust(len(inner), u"0")
        return token

    return _TOKEN_RE.sub(_replace, template)


def apply_case(text, case_mode):
    """Return text transformed by the chosen case mode."""
    if case_mode == CASE_UPPER:
        return text.upper()
    if case_mode == CASE_LOWER:
        return text.lower()
    if case_mode == CASE_TITLE:
        return title_case(text)
    return text


def apply_rule(rule, number, name, current, counter):
    """Return the new value for one field of one sheet.

    Args:
        rule (dict): keys mode, find, replace, template, case, match_case.
        number (unicode): the sheet's current number, for [N].
        name (unicode): the sheet's current name, for [NAME].
        current (unicode): the value of the field this rule targets.
        counter (int): this row's sequence value.

    Returns:
        unicode: the rewritten value.
    """
    if rule["mode"] == MODE_TEMPLATE:
        result = expand_template(rule["template"], number, name, counter)
    elif rule["mode"] == MODE_FIND_REPLACE and rule["find"]:
        if rule["match_case"]:
            result = current.replace(rule["find"], rule["replace"])
        else:
            # No str.replace() ignore-case in Python 2, so escape the needle
            # and let re carry the flag.
            pattern = re.compile(re.escape(rule["find"]), re.IGNORECASE)
            result = pattern.sub(lambda m: rule["replace"], current)
    else:
        result = current

    return apply_case(result, rule["case"])


# --- Preview rows ---

class RenameRow(object):
    """One sheet's before/after state, bound straight into the preview grid.

    Plain attributes with no change notification, so the window rebuilds the
    ObservableCollection on every recompute rather than mutating rows in
    place - the established pattern for WPF binding under IronPython 2.
    """

    # --- construction ---
    def __init__(self, seq, sheet):
        self.Seq = seq
        self.OldNumber = safe_text(sheet.SheetNumber)
        self.OldName = safe_text(sheet.Name)
        self.NewNumber = self.OldNumber
        self.NewName = self.OldName
        self.Status = u""
        self.RowState = u"unchanged"
        self.sheet = sheet

    # --- public methods ---
    @property
    def number_changed(self):
        return self.NewNumber != self.OldNumber

    @property
    def name_changed(self):
        return self.NewName != self.OldName

    def mark_error(self, message):
        self.Status = message
        self.RowState = u"error"


def build_rows(sheets):
    """Return RenameRows for sheets, ordered the way a drawing set reads."""
    ordered = sorted(sheets, key=lambda s: natural_key(s.SheetNumber))
    return [RenameRow(i + 1, sheet) for i, sheet in enumerate(ordered)]


def compute_preview(rows, number_rule, name_rule, seq_start, seq_step, other_numbers):
    """Fill in every row's new values and status, and report the error count.

    Args:
        rows (list of RenameRow): rows to update in place.
        number_rule (dict): rule for the sheet number field.
        name_rule (dict): rule for the sheet name field.
        seq_start (int): first counter value.
        seq_step (int): counter increment per row.
        other_numbers (set): sheet numbers already used by sheets outside the
            selection, which a new number is not allowed to land on.

    Returns:
        int: how many rows ended up in an error state.
    """
    counter = seq_start
    for row in rows:
        row.NewNumber = apply_rule(
            number_rule, row.OldNumber, row.OldName, row.OldNumber, counter)
        row.NewName = apply_rule(
            name_rule, row.OldNumber, row.OldName, row.OldName, counter)
        row.Status = u""
        row.RowState = u"unchanged"
        counter += seq_step

    # A number is only a duplicate against the *other* rows' new numbers, so
    # this has to be counted across the whole set before any row is judged.
    seen = {}
    for row in rows:
        seen[row.NewNumber] = seen.get(row.NewNumber, 0) + 1

    errors = 0
    for row in rows:
        problem = _row_problem(row, seen, other_numbers)
        if problem:
            row.mark_error(problem)
            errors += 1
        elif row.number_changed or row.name_changed:
            row.Status = u"Will be renamed"
            row.RowState = u"changed"

    return errors


def _row_problem(row, seen, other_numbers):
    """Return a human-readable reason this row cannot be written, or None."""
    if not row.NewNumber.strip():
        return u"Number is empty"
    if not row.NewName.strip():
        return u"Name is empty"

    bad = invalid_chars_in(row.NewNumber) or invalid_chars_in(row.NewName)
    if bad:
        return u"Illegal character: {}".format(u" ".join(bad))

    if seen.get(row.NewNumber, 0) > 1:
        return u"Duplicate number"
    if row.number_changed and row.NewNumber in other_numbers:
        return u"Number already used in project"
    return None


# --- Writing back ---

def temp_number_prefix(used_numbers):
    """Return a sheet-number prefix no existing sheet already starts with."""
    prefix = u"ZZTMP-"
    guard = 0
    while any(n.startswith(prefix) for n in used_numbers) and guard < 50:
        prefix = u"ZZTMP{}-".format(guard)
        guard += 1
    return prefix


def commit_renames(rows, prefix):
    """Write every pending change in one transaction, or none of them.

    Sheet numbers are unique in Revit, so renaming TD-100 to TD-101 while
    TD-101 still exists throws even when the whole set ends up conflict-free.
    Every changing sheet therefore parks on a temporary number first, and only
    then takes its final one - which makes any shift, swap or full reorder
    safe regardless of the order the rows happen to be in.

    Args:
        rows (list of RenameRow): the previewed rows.
        prefix (unicode): a sheet-number prefix known to be unused.

    Returns:
        tuple: (numbers_written, names_written)
    """
    number_rows = [r for r in rows if r.number_changed]
    name_rows = [r for r in rows if r.name_changed]

    transaction = Transaction(doc, "pyRename sheets")
    transaction.Start()
    try:
        for index, row in enumerate(number_rows):
            row.sheet.SheetNumber = u"{}{}".format(prefix, index)
        for row in number_rows:
            row.sheet.SheetNumber = row.NewNumber
        for row in name_rows:
            row.sheet.Name = row.NewName
        transaction.Commit()
    except Exception:
        # Rolling back matters more than usual here: a failure between the two
        # number passes would otherwise strand sheets on their temp numbers.
        if transaction.HasStarted():
            transaction.RollBack()
        raise

    return len(number_rows), len(name_rows)


def refresh_project_browser():
    """Close and reopen the Project Browser so new numbers show immediately.

    Revit does not re-sort the browser on a SheetNumber change on its own.
    Never allowed to be the reason a successful rename reports as failed.
    """
    try:
        pane = DockablePane(DockablePanes.BuiltInDockablePanes.ProjectBrowser)
        pane.Hide()
        pane.Show()
    except Exception:
        logger.debug("Could not refresh the Project Browser")


# --- Sheet selection ---

def selected_sheets_in_browser():
    """Return sheets currently selected in the Project Browser."""
    sheets = []
    for element_id in uidoc.Selection.GetElementIds():
        element = doc.GetElement(element_id)
        if isinstance(element, ViewSheet):
            sheets.append(element)
    return sheets


def pick_sheets():
    """Return the sheets to work on, and how they were chosen.

    Whatever is already highlighted in the Project Browser wins, so the tool
    picks up the normal "select sheets, run tool" habit; the picker only
    appears when nothing is selected.

    Returns:
        tuple: (list of ViewSheet, source label)
    """
    sheets = selected_sheets_in_browser()
    if sheets:
        return sheets, u"from Project Browser"

    chosen = forms.select_sheets(title="Select sheets to rename",
                                 button_name="Select")
    return list(chosen or []), u"selected"


def all_sheet_numbers():
    """Return every sheet number in the document, keyed by UniqueId.

    Keyed on UniqueId rather than ElementId so the dict does not depend on
    how ElementId hashes, which changed with the 64-bit migration in 2024.
    """
    numbers = {}
    for sheet in FilteredElementCollector(doc).OfClass(ViewSheet):
        numbers[safe_text(sheet.UniqueId)] = safe_text(sheet.SheetNumber)
    return numbers


# ── UI ─────────────────────────────────────────────────────────────────────

class MainWindow(forms.WPFWindow):
    """pyRename's window: rules on top, live preview below."""

    # --- construction ---
    def __init__(self, sheets, source_label):
        forms.WPFWindow.__init__(self, "pyRename.xaml")

        # Appearance immediately after load, before anything reads a resource.
        apply_seed43_palette(self, SCRIPT_DIR)
        apply_seed43_dimensions(self, SCRIPT_DIR)
        from Snippets._icons import set_header_icon
        set_header_icon(self, SCRIPT_DIR)

        self._sheets = sheets
        self._source_label = source_label
        self._rows = []
        self._grid_items = ObservableCollection[object]()
        self._errors = 0
        self._suspend_refresh = False

        self.FindName("preview_grid").ItemsSource = self._grid_items

        self._bind_events()
        self._load_sheets(sheets, source_label)

    # --- wiring ---
    def _bind_events(self):
        self.FindName("close_btn").Click += self.on_close
        self.FindName("rename_btn").Click += self.on_rename
        self.FindName("reset_btn").Click += self.on_reset
        self.FindName("reselect_btn").Click += self.on_reselect

        for name in ("number_mode_combo", "number_case_combo",
                     "name_mode_combo", "name_case_combo"):
            self.FindName(name).SelectionChanged += self.on_rule_changed

        for name in ("number_find_box", "number_replace_box", "number_template_box",
                     "name_find_box", "name_replace_box", "name_template_box",
                     "seq_start_box", "seq_step_box"):
            self.FindName(name).TextChanged += self.on_rule_changed

        match_case = self.FindName("match_case_chk")
        match_case.Checked += self.on_rule_changed
        match_case.Unchecked += self.on_rule_changed

    # --- reading the form ---
    def _rule_for(self, field):
        """Return the rule dict for 'number' or 'name'."""
        return {
            "mode": self.FindName("{}_mode_combo".format(field)).SelectedIndex,
            "case": self.FindName("{}_case_combo".format(field)).SelectedIndex,
            "find": safe_text(self.FindName("{}_find_box".format(field)).Text),
            "replace": safe_text(self.FindName("{}_replace_box".format(field)).Text),
            "template": safe_text(self.FindName("{}_template_box".format(field)).Text),
            "match_case": bool(self.FindName("match_case_chk").IsChecked),
        }

    def _sync_enabled_fields(self):
        """Grey out the inputs the current mode does not use."""
        for field in ("number", "name"):
            mode = self.FindName("{}_mode_combo".format(field)).SelectedIndex
            self.FindName("{}_find_box".format(field)).IsEnabled = \
                mode == MODE_FIND_REPLACE
            self.FindName("{}_replace_box".format(field)).IsEnabled = \
                mode == MODE_FIND_REPLACE
            self.FindName("{}_template_box".format(field)).IsEnabled = \
                mode == MODE_TEMPLATE
            self.FindName("{}_case_combo".format(field)).IsEnabled = \
                mode != MODE_KEEP

    # --- preview ---
    def _load_sheets(self, sheets, source_label):
        """Swap in a new sheet set and rebuild the preview from scratch."""
        self._sheets = sheets
        self._source_label = source_label
        self._rows = build_rows(sheets)

        selected_ids = set(safe_text(s.UniqueId) for s in sheets)
        self._other_numbers = set(
            number for uid, number in all_sheet_numbers().items()
            if uid not in selected_ids)

        self.FindName("subtitle_label").Text = u"{} sheet{} {}".format(
            len(sheets), u"" if len(sheets) == 1 else u"s", source_label)
        self._refresh_preview()

    def _refresh_preview(self):
        if self._suspend_refresh:
            return

        self._sync_enabled_fields()
        self._errors = compute_preview(
            self._rows,
            self._rule_for("number"),
            self._rule_for("name"),
            parse_int(self.FindName("seq_start_box").Text, 1),
            parse_int(self.FindName("seq_step_box").Text, 1),
            self._other_numbers,
        )

        # Plain Python attributes raise no change notification, so the grid
        # only picks up new values if the collection itself is rebuilt.
        self._grid_items.Clear()
        for row in self._rows:
            self._grid_items.Add(row)

        self._update_action_state()

    def _update_action_state(self):
        changed = len([r for r in self._rows
                       if r.RowState == u"changed"])

        self.FindName("preview_summary_label").Text = u"{} of {} changing".format(
            changed, len(self._rows))

        problem_label = self.FindName("problem_label")
        if self._errors:
            problem_label.Text = u"{} sheet{} cannot be renamed - fix the rule first".format(
                self._errors, u"" if self._errors == 1 else u"s")
        else:
            problem_label.Text = u""

        self.FindName("rename_btn").IsEnabled = bool(changed) and not self._errors
        self._set_status(u"Ready" if changed else u"No changes yet")

    def _set_status(self, text):
        self.FindName("status_label").Text = text

    # --- event handlers ---
    def on_rule_changed(self, sender, args):
        self._refresh_preview()

    def on_reset(self, sender, args):
        """Clear every rule back to leaving both fields untouched."""
        self._suspend_refresh = True
        try:
            for field in ("number", "name"):
                self.FindName("{}_mode_combo".format(field)).SelectedIndex = MODE_KEEP
                self.FindName("{}_case_combo".format(field)).SelectedIndex = CASE_KEEP
                self.FindName("{}_find_box".format(field)).Text = ""
                self.FindName("{}_replace_box".format(field)).Text = ""
                self.FindName("{}_template_box".format(field)).Text = ""
            self.FindName("seq_start_box").Text = "1"
            self.FindName("seq_step_box").Text = "1"
            self.FindName("match_case_chk").IsChecked = False
        finally:
            self._suspend_refresh = False
        self._refresh_preview()

    def on_reselect(self, sender, args):
        """Pick a different sheet set without losing the current rules."""
        chosen = forms.select_sheets(title="Select sheets to rename",
                                     button_name="Select")
        if chosen:
            self._load_sheets(list(chosen), u"selected")

    def on_rename(self, sender, args):
        if self._errors:
            return

        pending = [r for r in self._rows if r.RowState == u"changed"]
        if not pending:
            return

        if not dlg.confirm(
                u"Rename {} sheet{}?".format(
                    len(pending), u"" if len(pending) == 1 else u"s"),
                title="pyRename", yes="Rename", no="Cancel"):
            return

        used = set(all_sheet_numbers().values())
        self._set_status(u"Renaming...")
        try:
            numbers, names = commit_renames(self._rows, temp_number_prefix(used))
        except Exception as error:
            self._set_status(u"Nothing was changed")
            dlg.message(
                u"The rename was rolled back, so no sheet was changed.\n\n{}".format(error),
                title="pyRename")
            return

        refresh_project_browser()
        self._load_sheets(self._sheets, self._source_label)
        self._set_status(u"Renamed {} number{} and {} name{}".format(
            numbers, u"" if numbers == 1 else u"s",
            names, u"" if names == 1 else u"s"))

    def on_close(self, sender, args):
        self.Close()


# ── ENTRY POINT ────────────────────────────────────────────────────────────

if __name__ == "__main__":
    picked, label = pick_sheets()
    if not picked:
        script.exit()

    MainWindow(picked, label).ShowDialog()
