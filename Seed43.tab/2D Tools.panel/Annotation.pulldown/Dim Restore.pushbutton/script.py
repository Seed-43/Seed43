# -*- coding: utf-8 -*-
# "Dim Restore"
# "Seed43"
# """
# Find dimensions that have lost their override text and put it back.
#
# The record lives on each Dimension as an Extensible Storage entity, so it
# travels inside the .rvt and a deleted dimension takes its record with it -
# there is no index to prune and no id that can dangle. See
# lib/Snippets/_dimoverrides.py for the storage rules.
#
# Manual, never automatic. Nothing scans or restores until a button is
# pressed, so a model change can never trigger a silent rewrite.
#
# Target: Revit 2022-2026, IronPython 2.
# """

# ── IMPORTS ─────────────────────────────────────────────────────────────────

import os

from Autodesk.Revit.DB import (Dimension, ElementId, FilteredElementCollector,
                               Transaction)
from pyrevit import forms, revit
from pyrevit.framework import Windows
from System.Collections.Generic import List

from Snippets import _dimoverrides
from Snippets.seed43_theme import (apply_seed43_dimensions,
                                   apply_seed43_palette, get_color)

doc = revit.doc
uidoc = revit.uidoc

# ── CONSTANTS ───────────────────────────────────────────────────────────────

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TOOL_NAME = "Dim Restore"

# Column widths, mirrored in DimRestore.xaml's header row. The two must move
# together or the headings drift off the data.
COL_CHECK = 28
COL_STATUS = 96
COL_VIEW = 210
COL_ID = 78
COL_RECORDED = 230

# Only these reach the list. In-sync dimensions are counted in the summary
# but never listed - on a model with 500 tracked dimensions they would bury
# the handful that actually need attention.
ACTIONABLE = (_dimoverrides.DROPPED,
              _dimoverrides.CHANGED,
              _dimoverrides.RESEGMENTED)

STATUS_LABEL = {_dimoverrides.DROPPED:     "DROPPED",
                _dimoverrides.CHANGED:     "CHANGED",
                _dimoverrides.RESEGMENTED: "REVIEW"}

STATUS_COLOUR = {_dimoverrides.DROPPED:     "danger",
                 _dimoverrides.CHANGED:     "text_muted",
                 _dimoverrides.RESEGMENTED: "text_muted"}


# ── HELPERS ─────────────────────────────────────────────────────────────────

def eid(element_id):
    """ElementId as a plain number, across the 2024 64-bit change."""
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue


def view_name(dim):
    """The name of the view a dimension is drawn on, or a placeholder."""
    try:
        owner = doc.GetElement(dim.OwnerViewId)
        if owner is not None:
            return owner.Name
    except Exception:
        pass
    return u"(no view)"


def brush(key, fallback="#FFFFFF"):
    """A resolved brush for a palette key, for text built in Python."""
    return Windows.Media.SolidColorBrush(
        Windows.Media.ColorConverter.ConvertFromString(
            get_color(SCRIPT_DIR, key, fallback=fallback)))


# ── CORE LOGIC ──────────────────────────────────────────────────────────────

class Row(object):
    """One tracked dimension and how it currently differs from its record."""

    def __init__(self, dim, status, record, current):
        self.dim = dim
        self.status = status
        self.record = record
        self.current = current
        self.view = view_name(dim)
        self.id = eid(dim.Id)
        self.checkbox = None    # set when the list row is built

    @property
    def ticked(self):
        return bool(self.checkbox and self.checkbox.IsChecked)


def survey():
    """Compare every tracked dimension against its record.

    Returns (rows, in_sync_count). rows holds only the actionable ones.

    A tracked dimension cannot be missing here - the walk starts from live
    elements and the record is carried by the element itself, so a deleted
    dimension simply does not appear and its record went with it. That is
    the housekeeping, and it needs no pass of its own.
    """
    rows = []
    in_sync = 0
    for dim in _dimoverrides.tracked_dimensions(doc):
        try:
            status, record, current = _dimoverrides.compare(dim)
        except Exception:
            continue
        if status == _dimoverrides.IN_SYNC:
            in_sync += 1
        elif status in ACTIONABLE:
            rows.append(Row(dim, status, record, current))

    # Dropped first, since that is what the tool exists for, then by view so
    # a sheet's worth of damage reads as one block.
    order = {_dimoverrides.DROPPED: 0,
             _dimoverrides.RESEGMENTED: 1,
             _dimoverrides.CHANGED: 2}
    rows.sort(key=lambda r: (order.get(r.status, 9), r.view.lower(), r.id))
    return rows, in_sync


def scan_model():
    """Record every dimension that currently carries override text.

    Picks up overrides typed into Revit's own dialog, not just ones Dim
    Override wrote. Blank dimensions are skipped rather than recorded: a
    tracked dimension that has since gone blank is exactly the case this
    tool protects, and stamping it would erase the evidence.

    Returns (recorded, failed).
    """
    recorded = 0
    failed = 0
    t = Transaction(doc, "Record dimension overrides")
    t.Start()
    try:
        for dim in FilteredElementCollector(doc).OfClass(Dimension):
            if not _dimoverrides.is_target(dim):
                continue
            try:
                current = _dimoverrides.read_current(dim)
            except Exception:
                continue
            if not _dimoverrides.has_text(current):
                continue
            if _dimoverrides.store_record(dim, current):
                recorded += 1
            else:
                failed += 1
        t.Commit()
    except Exception:
        t.RollBack()
        raise
    return recorded, failed


def restore_rows(rows):
    """Write each row's record back onto its dimension. Returns failures."""
    failures = []
    t = Transaction(doc, "Restore dimension overrides")
    t.Start()
    try:
        for row in rows:
            try:
                _dimoverrides.write_segments(row.dim, row.record)
            except Exception as ex:
                failures.append(u"{}: {}".format(row.id, ex))
        t.Commit()
    except Exception:
        t.RollBack()
        raise
    return failures


def forget_rows(rows):
    """Delete the record from each row's dimension. Returns failures."""
    failures = []
    t = Transaction(doc, "Forget dimension overrides")
    t.Start()
    try:
        for row in rows:
            if not _dimoverrides.forget_record(row.dim):
                failures.append(unicode(row.id))
        t.Commit()
    except Exception:
        t.RollBack()
        raise
    return failures


# ── UI ──────────────────────────────────────────────────────────────────────

class DimRestoreWindow(forms.WPFWindow):
    """The drift report: what each tracked dimension held, and holds now."""

    # --- construction ---
    def __init__(self, xaml_name):
        forms.WPFWindow.__init__(self, xaml_name)
        apply_seed43_palette(self, SCRIPT_DIR)
        apply_seed43_dimensions(self, SCRIPT_DIR)

        # Resolved after the palette is applied, never before - a brush
        # built from an unresolved lookup comes back transparent and the
        # row text renders invisible with no error to explain it.
        self._muted = brush("text_muted", "#9CA3AF")
        self._primary = brush("text_primary", "#FFFFFF")
        self._colours = dict((status, brush(key, "#FFFFFF"))
                             for status, key in STATUS_COLOUR.items())

        self._rows = []
        self.refresh()

    # --- public methods ---
    def refresh(self, note=None):
        """Re-survey the model and rebuild the list."""
        self._rows, in_sync = survey()

        dropped = len([r for r in self._rows
                       if r.status == _dimoverrides.DROPPED])
        changed = len([r for r in self._rows
                       if r.status == _dimoverrides.CHANGED])
        review = len([r for r in self._rows
                      if r.status == _dimoverrides.RESEGMENTED])
        tracked = len(self._rows) + in_sync

        self.subtitle_tb.Text = u"|  {} tracked".format(tracked)
        if self._rows:
            self.summary_tb.Text = u"{} dropped, {} changed, {} to review".format(
                dropped, changed, review)
        else:
            self.summary_tb.Text = u"Nothing to restore"
        self.explain_tb.Text = (
            u"{} tracked dimensions still match their record. Only dropped "
            u"overrides are ticked. Re-scan after typing overrides into "
            u"Revit's own dialog, or they are not yet protected.".format(
                in_sync)
            if tracked else
            u"No dimensions are tracked yet, so nothing can be restored. "
            u"Scan Model records every one that currently carries override "
            u"text.")

        self.rows_lb.Items.Clear()
        for row in self._rows:
            self.rows_lb.Items.Add(self._build_row(row))

        self.restore_btn.IsEnabled = bool(self._rows)
        self.discard_btn.IsEnabled = bool(self._rows)
        self.status_tb.Text = note or u""

    # --- event handlers ---
    def scan_clicked(self, sender, args):
        recorded, failed = scan_model()
        note = u"Scan recorded {} dimension{}. Save the model to keep " \
               u"them.".format(recorded, "" if recorded == 1 else "s")
        if failed:
            note += u" {} could not be written: {}".format(
                failed, _dimoverrides.last_error() or u"unknown")
        self.refresh(note)

    def restore_clicked(self, sender, args):
        picked = [r for r in self._rows if r.ticked]
        if not picked:
            self.status_tb.Text = u"Nothing ticked."
            return
        failures = restore_rows(picked)
        note = u"Restored {} of {}.".format(len(picked) - len(failures),
                                            len(picked))
        if failures:
            note += u" Failed: " + u"; ".join(failures[:5])
        self.refresh(note)

    def discard_clicked(self, sender, args):
        picked = [r for r in self._rows if r.ticked]
        if not picked:
            self.status_tb.Text = u"Nothing ticked."
            return
        if not forms.alert(u"Forget the stored record for {} dimension{}?\n\n"
                           u"Their text is left exactly as it is now, but "
                           u"they stop being tracked and this tool will no "
                           u"longer offer to restore them.".format(
                               len(picked), "" if len(picked) == 1 else "s"),
                           title=TOOL_NAME, yes=True, no=True):
            return
        failures = forget_rows(picked)
        note = u"Forgot {} of {}.".format(len(picked) - len(failures),
                                          len(picked))
        self.refresh(note)

    def show_clicked(self, sender, args):
        """Open the highlighted row's own view and zoom to its dimension.

        ShowElements on its own is not enough. It only searches views that
        are already OPEN, so a dimension on a closed view comes back as
        Revit's own "no open view shows any of the highlighted elements"
        box, stacked behind this window where it reads as the tool having
        hung. Activating the owner view first turns that into a hit every
        time, and leaves ShowElements to do nothing but frame the element.

        The window stays modal deliberately. Revit still redraws behind a
        modal dialog - this is exactly how its own Select by ID > Show
        behaves - so the jump is visible without handing the model back
        mid-session, which would leave this list describing a state that
        has since moved on.
        """
        item = self.rows_lb.SelectedItem
        if item is None:
            self.status_tb.Text = u"Highlight a row first."
            return
        row = item.Tag

        try:
            owner = doc.GetElement(row.dim.OwnerViewId)
        except Exception:
            owner = None

        if owner is not None:
            try:
                if owner.Id != uidoc.ActiveView.Id:
                    uidoc.ActiveView = owner
            except Exception:
                # A view template, a view Revit refuses to activate, or a
                # context that will not take a view change. Fall through:
                # ShowElements may still find it among the open views, and
                # if it cannot it reports that itself.
                pass

        try:
            ids = List[ElementId]()
            ids.Add(row.dim.Id)
            uidoc.Selection.SetElementIds(ids)
            uidoc.ShowElements(row.dim.Id)
        except Exception as ex:
            self.status_tb.Text = u"Could not show {}: {}".format(row.id, ex)
            return

        self.status_tb.Text = u"Showing {} in {}.".format(row.id, row.view)

    def all_clicked(self, sender, args):
        self._tick_all(True)

    def none_clicked(self, sender, args):
        self._tick_all(False)

    def close_clicked(self, sender, args):
        self.Close()

    # --- private helpers ---
    def _tick_all(self, state):
        for row in self._rows:
            if row.checkbox is not None:
                row.checkbox.IsChecked = state

    def _build_row(self, row):
        """One list line, laid out to the same column widths as the header."""
        item = Windows.Controls.ListBoxItem()
        item.Tag = row

        panel = Windows.Controls.StackPanel()
        panel.Orientation = Windows.Controls.Orientation.Horizontal

        row.checkbox = Windows.Controls.CheckBox()
        row.checkbox.Width = COL_CHECK
        row.checkbox.VerticalAlignment = Windows.VerticalAlignment.Center
        # Only a dropped override is a safe default. A changed one means
        # somebody typed something else on purpose, and ticking that by
        # default would quietly undo their edit.
        row.checkbox.IsChecked = (row.status == _dimoverrides.DROPPED)
        panel.Children.Add(row.checkbox)

        panel.Children.Add(self._cell(STATUS_LABEL.get(row.status, row.status),
                                      COL_STATUS,
                                      self._colours.get(row.status),
                                      bold=True))
        panel.Children.Add(self._cell(row.view, COL_VIEW, self._primary))
        panel.Children.Add(self._cell(unicode(row.id), COL_ID, self._muted))
        panel.Children.Add(self._cell(_dimoverrides.summarise(row.record),
                                      COL_RECORDED, self._primary))
        panel.Children.Add(self._cell(_dimoverrides.summarise(row.current),
                                      None, self._muted))

        item.Content = panel
        return item

    def _cell(self, text, width, colour, bold=False):
        block = Windows.Controls.TextBlock()
        block.Text = text or u""
        block.TextTrimming = Windows.TextTrimming.CharacterEllipsis
        block.VerticalAlignment = Windows.VerticalAlignment.Center
        block.Margin = Windows.Thickness(0, 0, 10, 0)
        block.ToolTip = text or None
        if width:
            block.Width = width
        if colour is not None:
            block.Foreground = colour
        if bold:
            block.FontWeight = Windows.FontWeights.SemiBold
        return block


# ── ENTRY POINT ─────────────────────────────────────────────────────────────

def main():
    # Nothing is recoverable until a record exists, and the record has to be
    # written BEFORE Revit drops the text - afterwards the old value is gone
    # from the file and no tool can recover what was never written down.
    #
    # So an untracked model is not a state to open quietly into: the window
    # would show an empty list, which reads as "no problems found" when it
    # actually means "not watching anything yet". Ask up front instead.
    if not _dimoverrides.tracked_dimensions(doc):
        if forms.alert(u"No dimensions are being tracked in this model yet, "
                       u"so there is nothing to restore.\n\n"
                       u"Overrides have to be recorded BEFORE Revit drops "
                       u"them. Scan the model now to record every dimension "
                       u"that currently carries override text?",
                       title=TOOL_NAME, yes=True, no=True):
            recorded, failed = scan_model()
            message = u"Recorded {} dimension{}.\n\nThese are unsaved " \
                      u"changes - save the model to keep them.".format(
                          recorded, "" if recorded == 1 else "s")
            if failed:
                message += u"\n\n{} could not be written: {}".format(
                    failed, _dimoverrides.last_error() or u"unknown")
            forms.alert(message, title=TOOL_NAME)

    DimRestoreWindow("DimRestore.xaml").ShowDialog()


main()
