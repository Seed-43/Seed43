# -*- coding: utf-8 -*-
# "View Refs"
# "Seed43"
# """
# Find views on sheets whose mark no longer reports a Referencing Sheet,
# and put the reference back.
#
# Revit takes Referencing Sheet from the ONE view the section, detail or
# elevation mark was cut in, not from every view the mark is drawn on. Lose
# that hold - the parent view deleted, taken off its sheet, or the mark's
# extents pulled clear of it - and the bubble on the sheet reads blank, with
# nothing left in the model saying where it used to point.
#
# THE FIX, AND WHY IT LOOKS LIKE A NO-OP
#     Writing the far clip offset and writing it straight back rebuilds the
#     association. The view ends on the exact value it started on, so nothing
#     about the drawing changes. What changes is that Revit re-derives the
#     mark's parent on the way through, and the bubble fills in.
#
#     The two writes MUST land in separate COMMITTED transactions. The
#     rebuild fires on commit, not on Regenerate, so a bump and a restore
#     inside one transaction cancel out and nothing happens. That is the
#     whole reason a long run of earlier attempts read as dead ends: they
#     were tested inside a transaction that was rolled back, which is a
#     window the rebuild can never fire in. Do not tidy the two transactions
#     into one, and do not move the restore in alongside a read.
#
# WHAT IT CANNOT FIX
#     Only RE-CUT rows have anywhere to point. OFF SHEET means the mark
#     draws solely on views that are not on a sheet; NO MARK means no mark
#     survives anywhere. No amount of nudging invents a sheet number for
#     either, so those need a view placing or a section re-cutting by hand.
#
# Target: Revit 2022-2026, IronPython 2.
# """

# ── IMPORTS ─────────────────────────────────────────────────────────────────

import os

from Autodesk.Revit.DB import (BuiltInCategory, BuiltInParameter, ElementId,
                               FilteredElementCollector, Transaction, View,
                               ViewSheet, ViewType, Viewport)
from pyrevit import forms, revit
from pyrevit.framework import Windows
from System.Collections.Generic import List

from Snippets.seed43_theme import (apply_seed43_dimensions,
                                   apply_seed43_palette, get_color)

doc = revit.doc
uidoc = revit.uidoc

# ── CONSTANTS ───────────────────────────────────────────────────────────────

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TOOL_NAME = "View Refs"

# Column widths, mirrored in ViewRefs.xaml's header row. The two must move
# together or the headings drift off the data.
COL_STATUS = 92
COL_VIEW = 240
COL_ID = 74
COL_SHEET = 62

# Only the view types that carry a mark another view can point at. A plan or
# a schedule has no bubble to go blank, so listing them would pad the report
# with rows that can never be wrong.
MARKED_TYPES = (ViewType.Section, ViewType.Detail, ViewType.Elevation)


def _view_types(*names):
    """The named ViewType members that this Revit build actually has.

    The enum has gained and lost members across releases, so naming one that
    is missing would take the whole tool down on import rather than in the
    one place it is used.
    """
    found = []
    for name in names:
        member = getattr(ViewType, name, None)
        if member is not None:
            found.append(member)
    return tuple(found)


# Views that cannot draw a mark. Skipping them up front is what keeps the
# host sweep to one pass over the model rather than one pass per blank view.
NON_HOSTS = _view_types("DrawingSheet", "Schedule", "Legend", "Report",
                        "ProjectBrowser", "SystemBrowser", "Internal",
                        "Undefined")

RECUT = "recut"         # mark draws on a sheeted view, reference still gone
OFF_SHEET = "offsheet"  # mark draws, but on nothing that reaches a sheet
NO_MARK = "nomark"      # no mark found in any view at all

# How far the far clip offset is pushed before being put straight back.
# The size is irrelevant to the rebuild, which is triggered by the write and
# not by the magnitude - but a value this far from any real clip offset makes
# an interrupted run obvious in the Properties palette instead of passing for
# a plausible setting somebody chose.
NUDGE = 10.0

# Feet. Tighter than any offset Revit will round-trip through its own UI, so
# a restore that lands inside it is exact for every practical purpose.
EXACT = 1e-9

STATUS_LABEL = {RECUT: "RE-CUT", OFF_SHEET: "OFF SHEET", NO_MARK: "NO MARK"}
STATUS_COLOUR = {RECUT: "danger", OFF_SHEET: "text_muted",
                 NO_MARK: "text_muted"}


# ── HELPERS ─────────────────────────────────────────────────────────────────

def eid(element_id):
    """ElementId as a plain number, across the 2024 64-bit change."""
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue


def referencing_sheet(view):
    """The sheet Revit says the view's mark was cut on, or "".

    Read-only and derived. Revit recomputes it live from the mark, so it is
    the truth about what the bubble on the sheet prints, not a cached value
    that a stale read could be hiding.
    """
    try:
        param = view.get_Parameter(BuiltInParameter.VIEW_REFERENCING_SHEET)
    except Exception:
        return ""
    if param is None:
        return ""
    return param.AsString() or ""


def far_clip(view):
    """The writable far clip offset parameter of view, or None.

    Read-only ones are handed back as None rather than raising on the write.
    A view the fix cannot reach is reported as skipped, which is a smaller
    problem than a run that stops halfway with some offsets still bumped.
    """
    try:
        param = view.get_Parameter(BuiltInParameter.VIEWER_BOUND_OFFSET_FAR)
    except Exception:
        return None
    if param is None or param.IsReadOnly:
        return None
    return param


def brush(key, fallback="#FFFFFF"):
    """A resolved brush for a palette key, for text built in Python."""
    return Windows.Media.SolidColorBrush(
        Windows.Media.ColorConverter.ConvertFromString(
            get_color(SCRIPT_DIR, key, fallback=fallback)))


# ── CORE LOGIC ──────────────────────────────────────────────────────────────

def sheet_map():
    """Every view placed on a sheet, mapped to that sheet's number."""
    placed = {}
    for viewport in FilteredElementCollector(doc).OfClass(Viewport):
        sheet = doc.GetElement(viewport.SheetId)
        if isinstance(sheet, ViewSheet):
            placed[eid(viewport.ViewId)] = sheet.SheetNumber
    return placed


def host_map(views):
    """Which views draw each mark: mark id -> list of views.

    Built by sweeping the model once and inverting the result, rather than
    asking each blank view where its mark shows. The sweep costs the same
    either way, and inverting means a model with thirty blank references
    does not pay for thirty passes over every view.
    """
    hosts = {}
    for view in views:
        try:
            drawn = (FilteredElementCollector(doc, view.Id)
                     .OfCategory(BuiltInCategory.OST_Viewers)
                     .ToElementIds())
        except Exception:
            # A view Revit will not open a collector on. One missing host
            # costs a row some detail; it must not cost the whole report.
            continue
        for mark_id in drawn:
            hosts.setdefault(eid(mark_id), []).append(view)
    return hosts


def mark_map():
    """Mark name -> mark element.

    Revit enforces unique view names and a mark carries the name of the view
    it opens, so the name is a sound key. Neither side states the pairing
    outright, and the mark's own id has no fixed relationship to the view's.
    """
    marks = {}
    for mark in (FilteredElementCollector(doc)
                 .OfCategory(BuiltInCategory.OST_Viewers)
                 .WhereElementIsNotElementType()):
        marks[mark.Name] = mark
    return marks


class Row(object):
    """One view on a sheet whose mark reports no referencing sheet."""

    def __init__(self, view, sheet, mark, hosts, placed):
        self.view = view
        self.name = view.Name
        self.id = eid(view.Id)
        self.sheet = sheet
        self.mark = mark
        self.hosts = hosts      # the views that draw the mark
        self.placed = placed
        self.sheeted = [h for h in hosts if eid(h.Id) in placed]

        if self.sheeted:
            self.status = RECUT
        elif hosts:
            self.status = OFF_SHEET
        else:
            self.status = NO_MARK

    @property
    def target(self):
        """The view to open when Show is pressed.

        A sheeted host first: that is where the reference is meant to be
        read from, so it is where the repair gets made. Failing that any
        host at all, and failing that the view itself, which at least puts
        something on screen rather than reporting nothing to show.
        """
        if self.sheeted:
            return self.sheeted[0]
        if self.hosts:
            return self.hosts[0]
        return self.view

    @property
    def where(self):
        """The host views, each spelled out with the sheet it reaches."""
        if not self.hosts:
            return u"mark not drawn in any view"
        parts = []
        for host in self.hosts[:4]:
            number = self.placed.get(eid(host.Id))
            parts.append(u"{} [{}]".format(host.Name, number or u"no sheet"))
        if len(self.hosts) > 4:
            parts.append(u"+{} more".format(len(self.hosts) - 4))
        return u"   ".join(parts)


def survey():
    """Every sheeted view whose reference has gone blank.

    Returns (rows, checked) where checked is how many were examined, so the
    window can say what the report is silent about as well as what it lists.
    """
    placed = sheet_map()

    views = []
    for view in FilteredElementCollector(doc).OfClass(View):
        if view.IsTemplate or view.ViewType in NON_HOSTS:
            continue
        views.append(view)

    hosts = host_map(views)
    marks = mark_map()

    rows = []
    checked = 0
    for view in views:
        if view.ViewType not in MARKED_TYPES:
            continue
        number = placed.get(eid(view.Id))
        if number is None:
            continue
        checked += 1
        if referencing_sheet(view):
            continue

        mark = marks.get(view.Name)
        drawn = []
        if mark is not None:
            # A view always draws its own mark. That is not a reference to
            # anywhere, so it never counts as a host.
            drawn = [h for h in hosts.get(eid(mark.Id), [])
                     if eid(h.Id) != eid(view.Id)]
        rows.append(Row(view, number, mark, drawn, placed))

    # Re-cuttable first: those are the ones with somewhere to go and a fix
    # that takes seconds. Sheet order after that, so a run through the list
    # is a run through the drawing set.
    rows.sort(key=lambda r: (r.status != RECUT, r.sheet, r.name))
    return rows, checked


# ── REPAIR ──────────────────────────────────────────────────────────────────

class Repair(object):
    """What one pass of the fix did, split by what the user must do next."""

    def __init__(self):
        self.fixed = []      # reference now reads
        self.missed = []     # nudged cleanly, reference still blank
        self.skipped = []    # no writable far clip offset to nudge
        self.stranded = []   # (row, wanted, left_at) - a view left wrong

    @property
    def touched(self):
        return len(self.fixed) + len(self.missed) + len(self.stranded)


def _write_offsets(rows, value_for, label):
    """Set every row's far clip offset in one committed transaction.

    Its own transaction on purpose. The association is rebuilt on commit, so
    a caller that folds the bump and the restore together gets a pair of
    writes that cancel out and a fix that silently does nothing.
    """
    t = Transaction(doc, label)
    t.Start()
    try:
        for row in rows:
            param = far_clip(row.view)
            if param is not None:
                param.Set(value_for(row))
        t.Commit()
    except Exception:
        t.RollBack()
        raise


def repair(rows):
    """Rebuild the reference on every re-cuttable row. Returns a Repair.

    Bump the far clip offset, commit, put it back, commit. The view ends on
    the value it started on and the drawing is unchanged; the commit in
    between is what makes Revit re-derive the mark's parent view.

    The restore is verified and retried once, because the failure it guards
    against is the serious one. A reference that stays blank leaves the model
    exactly as this tool found it, but an offset left sitting at plus ten feet
    is damage the tool caused, on a view nobody is looking at.
    """
    report = Repair()

    originals = {}
    working = []
    for row in rows:
        if row.status != RECUT:
            continue
        param = far_clip(row.view)
        if param is None:
            report.skipped.append(row)
            continue
        originals[row.id] = param.AsDouble()
        working.append(row)

    if not working:
        return report

    _write_offsets(working, lambda r: originals[r.id] + NUDGE,
                   "Refresh view references")
    _write_offsets(working, lambda r: originals[r.id],
                   "Restore far clip offsets")

    drifted = [r for r in working
               if not _is_restored(r, originals[r.id])]
    if drifted:
        _write_offsets(drifted, lambda r: originals[r.id],
                       "Restore far clip offsets (retry)")

    for row in working:
        wanted = originals[row.id]
        if not _is_restored(row, wanted):
            param = far_clip(row.view)
            report.stranded.append(
                (row, wanted, param.AsDouble() if param is not None else None))
        elif referencing_sheet(row.view):
            report.fixed.append(row)
        else:
            report.missed.append(row)

    return report


def _is_restored(row, wanted):
    param = far_clip(row.view)
    if param is None:
        return False
    return abs(param.AsDouble() - wanted) <= EXACT


# ── UI ──────────────────────────────────────────────────────────────────────

class ViewRefsWindow(forms.WPFWindow):
    """The blank-reference report: which view, and where its mark still is."""

    # --- construction ---
    def __init__(self, xaml_name):
        forms.WPFWindow.__init__(self, xaml_name)
        apply_seed43_palette(self, SCRIPT_DIR)
        apply_seed43_dimensions(self, SCRIPT_DIR)

        # Resolved after the palette is applied, never before - a brush
        # built from an unresolved lookup comes back transparent and the row
        # text renders invisible with no error to explain it.
        self._muted = brush("text_muted", "#9CA3AF")
        self._primary = brush("text_primary", "#FFFFFF")
        self._colours = dict((status, brush(key, "#FFFFFF"))
                             for status, key in STATUS_COLOUR.items())

        self._rows = []
        self.refresh()

    # --- public methods ---
    def refresh(self, note=None):
        """Re-read every reference and rebuild the list."""
        self._rows, checked = survey()

        recut = len([r for r in self._rows if r.status == RECUT])
        off = len([r for r in self._rows if r.status == OFF_SHEET])
        bare = len([r for r in self._rows if r.status == NO_MARK])

        self.subtitle_tb.Text = u"|  {} on sheets".format(checked)
        self.summary_tb.Text = (
            u"{} blank: {} to re-cut, {} off sheet, {} with no mark".format(
                len(self._rows), recut, off, bare)
            if self._rows else u"Every reference reads")
        self.explain_tb.Text = (
            u"Revit takes the reference from the one view the mark was cut "
            u"in, not from every view it appears on. Fix rebuilds that hold "
            u"on the RE-CUT rows and leaves every view on the settings it "
            u"started with. OFF SHEET and NO MARK have no sheeted view to "
            u"point at, so those need placing or re-cutting: Show takes you "
            u"to them."
            if self._rows else
            u"All {} section, detail and elevation views on sheets report a "
            u"referencing sheet.".format(checked))

        self.rows_lb.Items.Clear()
        for row in self._rows:
            self.rows_lb.Items.Add(self._build_row(row))

        self.show_btn.IsEnabled = bool(self._rows)
        self.fix_btn.IsEnabled = bool(recut)
        self.fix_btn.Content = (u"Fix {}".format(recut) if recut
                                else u"Nothing to Fix")
        self.status_tb.Text = note or u""

    # --- event handlers ---
    def recheck_clicked(self, sender, args):
        self.refresh(u"Rechecked against the model as it stands now.")

    def fix_clicked(self, sender, args):
        """Rebuild the reference on every re-cuttable row.

        All of them at once rather than the highlighted one. The fix costs
        two transactions no matter how many views ride along, and a list of
        five that has to be worked one row at a time invites stopping at
        four without noticing.
        """
        recut = [r for r in self._rows if r.status == RECUT]
        if not recut:
            self.status_tb.Text = (
                u"Nothing here can be rebuilt: these rows have no sheeted "
                u"view for a reference to point at.")
            return

        try:
            report = repair(recut)
        except Exception as ex:
            forms.alert(u"The refresh stopped and was rolled back:\n\n"
                        u"{}".format(ex), title=TOOL_NAME)
            self.refresh(u"Refresh failed. Nothing was changed.")
            return

        # Loud, and before the list is rebuilt. A stranded view is the one
        # outcome where the tool has left the model worse than it found it,
        # and it must not be something the user has to notice in a status
        # line to find out about.
        if report.stranded:
            lines = []
            for row, wanted, left_at in report.stranded:
                lines.append(u"{}  [{}]  wants {:.6f} ft, left at {}".format(
                    row.name, row.id, wanted,
                    u"{:.6f} ft".format(left_at) if left_at is not None
                    else u"unreadable"))
            forms.alert(
                u"{} view(s) could not be put back on their original far "
                u"clip offset, after a retry. Set these by hand or undo "
                u"this run:\n\n{}".format(len(report.stranded),
                                           u"\n".join(lines)),
                title=TOOL_NAME)

        parts = [u"{} rebuilt".format(len(report.fixed))]
        if report.missed:
            parts.append(u"{} did not take".format(len(report.missed)))
        if report.skipped:
            parts.append(u"{} had no writable far clip".format(
                len(report.skipped)))
        if report.stranded:
            parts.append(u"{} LEFT ON THE WRONG OFFSET".format(
                len(report.stranded)))
        self.refresh(u", ".join(parts) + u".")

    def show_clicked(self, sender, args):
        """Open the view holding the mark and zoom to it.

        ShowElements on its own is not enough. It only searches views that
        are already OPEN, so a mark on a closed view comes back as Revit's
        own "no open view shows any of the highlighted elements" box,
        stacked behind this window where it reads as the tool having hung.
        Activating the host view first turns that into a hit every time.

        The window stays modal deliberately. Revit still redraws behind a
        modal dialog, which is how its own Select by ID > Show behaves, so
        the jump is visible without handing the model back mid-session and
        leaving this list describing a state that has since moved on.
        """
        item = self.rows_lb.SelectedItem
        if item is None:
            self.status_tb.Text = u"Highlight a row first."
            return
        row = item.Tag

        target = row.target
        try:
            if target.Id != uidoc.ActiveView.Id:
                uidoc.ActiveView = target
        except Exception:
            # A view Revit refuses to activate, or a context that will not
            # take a view change. Fall through: ShowElements may still find
            # it among the open views, and reports for itself if it cannot.
            pass

        subject = row.mark if row.mark is not None else row.view
        try:
            ids = List[ElementId]()
            ids.Add(subject.Id)
            uidoc.Selection.SetElementIds(ids)
            uidoc.ShowElements(subject.Id)
        except Exception as ex:
            self.status_tb.Text = u"Could not show {}: {}".format(row.id, ex)
            return

        if row.mark is None:
            self.status_tb.Text = (
                u"{} has no mark left to show, so the view itself is open "
                u"instead.".format(row.name))
        else:
            self.status_tb.Text = u"Showing the mark for {} in {}.".format(
                row.name, target.Name)

    def close_clicked(self, sender, args):
        self.Close()

    # --- private helpers ---
    def _build_row(self, row):
        """One list line, laid out to the same column widths as the header."""
        item = Windows.Controls.ListBoxItem()
        item.Tag = row

        panel = Windows.Controls.StackPanel()
        panel.Orientation = Windows.Controls.Orientation.Horizontal

        panel.Children.Add(self._cell(STATUS_LABEL.get(row.status, row.status),
                                      COL_STATUS,
                                      self._colours.get(row.status),
                                      bold=True))
        panel.Children.Add(self._cell(row.name, COL_VIEW, self._primary))
        panel.Children.Add(self._cell(unicode(row.id), COL_ID, self._muted))
        panel.Children.Add(self._cell(row.sheet, COL_SHEET, self._primary))
        panel.Children.Add(self._cell(row.where, None, self._muted))

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
    ViewRefsWindow("ViewRefs.xaml").ShowDialog()


main()
