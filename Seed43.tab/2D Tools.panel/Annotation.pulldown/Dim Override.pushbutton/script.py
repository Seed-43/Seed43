# -*- coding: utf-8 -*-
# "Dim Override"
# "Seed43"
# """
# Batch-write the dimension text fields - Above, Prefix, Value, Suffix,
# Below - across a picked set of dimensions.
#
# Selection is expressed in the model itself rather than in a list: every
# dimension picked goes halftone, picking it again takes the halftone off.
# The whole selection phase runs inside a TransactionGroup that is rolled
# back when picking ends, so the halftone never survives the tool and the
# user's own view overrides come back untouched.
#
# The Value field predicts from a per-machine catalogue of overrides used
# before, ranked by how often each has been applied. That catalogue is a
# habit, not project data, so it lives in .user rather than in the model.
#
# What gets WRITTEN is recorded into the model itself, on each dimension,
# so Dim Restore can put it back when Revit later drops it. See
# lib/Snippets/_dimoverrides.py for the storage rules.
#
# Target: Revit 2022-2026, IronPython 2.
# """

# ── IMPORTS ─────────────────────────────────────────────────────────────────

import json
import os

from Autodesk.Revit.DB import (OverrideGraphicSettings, Transaction,
                               TransactionGroup)
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from pyrevit import forms, revit
from pyrevit.framework import Windows

from Snippets import _dimoverrides, _userdata
from Snippets.seed43_theme import (apply_seed43_dimensions,
                                   apply_seed43_palette, get_color)

doc = revit.doc
uidoc = revit.uidoc

# ── CONSTANTS ───────────────────────────────────────────────────────────────

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TOOL_NAME = "Dim Override"

CATALOGUE_PATH = _userdata.user_path(TOOL_NAME, "overrides.json")

# Seeded on first run so the list is not empty on a fresh install. Count 0,
# so anything the user actually applies immediately outranks them.
SEED_OVERRIDES = ["LAP", "N.T.S", "TYP", "EQ", "VARIES", "MIN", "MAX", "CLR"]

FIELDS = _dimoverrides.FIELDS

HALFTONE = OverrideGraphicSettings()
HALFTONE.SetHalftone(True)

CLEAR_OVERRIDE = OverrideGraphicSettings()


# ── CATALOGUE ───────────────────────────────────────────────────────────────

def load_catalogue():
    """Return the {override text: times applied} map, seeded on first run.

    Never raises. An unreadable or corrupt file is treated as a fresh one -
    losing the usage history is a nuisance, refusing to open the tool is not
    an acceptable response to it.
    """
    data = None
    try:
        if os.path.isfile(CATALOGUE_PATH):
            # Read as bytes and decode explicitly. Overrides carry Ø and ²,
            # and leaving the codec to the default would make the load
            # depend on the machine's code page. See save_catalogue().
            with open(CATALOGUE_PATH, "rb") as f:
                data = json.loads(f.read().decode("utf-8"))
    except Exception:
        data = None

    if not isinstance(data, dict):
        data = {}
    if not data:
        for name in SEED_OVERRIDES:
            data[name] = 0

    # Guard against a hand-edited file carrying non-numeric counts.
    clean = {}
    for name, count in data.items():
        try:
            clean[unicode(name)] = int(count)
        except Exception:
            continue
    return clean


def save_catalogue(catalogue):
    """Write the catalogue back to .user as UTF-8. Never raises.

    NOTE: ensure_ascii=False and the explicit encode are both required.
    IronPython 2 makes str and unicode the same .NET String, so CPython 2's
    ascii encoder byte-decodes anything holding U+0080..U+00FF and .NET
    throws - and dimension overrides are full of it, Ø for diameter and ²
    for square metres. Writing the encoded bytes ourselves also keeps the
    file off the machine's code page. Same trap as _dimoverrides.
    """
    try:
        payload = json.dumps(catalogue, indent=2, sort_keys=True,
                             ensure_ascii=False)
        with open(CATALOGUE_PATH, "wb") as f:
            f.write(payload.encode("utf-8"))
    except Exception:
        pass


def ranked(catalogue):
    """Override texts most-used first, ties broken alphabetically.

    The alphabetical tie-break matters more than it looks: every seeded
    entry starts on zero, so without it the untouched part of the list would
    reshuffle on every load and the muscle memory for "LAP is second" would
    never form.
    """
    names = list(catalogue.keys())
    names.sort(key=lambda n: (-catalogue.get(n, 0), n.lower()))
    return names


def record(catalogue, text, times=1):
    """Credit text with times uses. Blank text is not catalogued."""
    if not text:
        return
    catalogue[text] = catalogue.get(text, 0) + times


# ── REVIT HELPERS ───────────────────────────────────────────────────────────

def eid(element_id):
    """ElementId as a plain number, across the 2024 64-bit change."""
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue


def seed_fields(dim, index):
    """The five text fields to open the dialog on.

    Read off the carrier the user actually picked, not carrier 0. On a
    dimension string the segments legitimately differ, so seeding from the
    first one would show the dialog text belonging to a segment nobody
    clicked.
    """
    current = _dimoverrides.read_current(dim)
    if 0 <= index < len(current):
        return current[index]
    return current[0]


def actual_value(dim, index):
    """The measured value of one carrier as Revit renders it, or "".

    Asking the Dimension itself is no good on a string: it reports an empty
    ValueString because there is no single number to give. The segment does
    know its own, so addressing the carrier gets a real number back for a
    picked segment where the old whole-element read returned nothing.
    """
    try:
        carriers = _dimoverrides.text_carriers(dim)
        if 0 <= index < len(carriers):
            return carriers[index].ValueString or ""
    except Exception:
        pass
    return ""


# ── SELECTION ───────────────────────────────────────────────────────────────

class DimensionFilter(ISelectionFilter):
    """Keep the pick on dimensions, so nothing else is ever halftoned."""

    def AllowElement(self, elem):
        try:
            return _dimoverrides.is_target(elem)
        except Exception:
            return False

    def AllowReference(self, ref, point):
        return False


class Target(object):
    """One picked dimension and which of its segments to write.

    segments is always a set of carrier indices, never a "means all"
    sentinel - a single-segment dimension is simply the set {0}. Keeping one
    shape means the writer never has to ask which kind of selection it is
    holding.
    """

    def __init__(self, element_id, carrier_count):
        self.element_id = element_id
        self.carrier_count = carrier_count
        self.segments = set()

    @property
    def is_string(self):
        return self.carrier_count > 1

    def toggle(self, index):
        """Add or drop one segment. index None means the whole dimension."""
        if index is None:
            every = set(range(self.carrier_count))
            self.segments = set() if self.segments == every else every
        elif index in self.segments:
            self.segments.discard(index)
        else:
            self.segments.add(index)


def clicked_segment(dim, ref):
    """Which segment the pick landed on, or None if it cannot be told.

    A dimension string is ONE element with ONE ElementId, so the reference
    alone cannot say which of its texts was meant - but the click point can.
    Reference.GlobalPoint is the model point under the cursor, and each
    segment publishes an Origin on the dimension line, so the nearest origin
    is the segment that was clicked.

    Returns None rather than guessing when the geometry cannot separate the
    segments. Collapsed or zero-length segments all report the SAME origin,
    and that is not a rare theoretical case - this model carries dimension
    strings whose nine segments all sit on one point. A wrong guess there
    would write the override onto a segment nobody touched, so an ambiguous
    read falls back to the whole dimension, which is at worst the behaviour
    that was there before.
    """
    carriers = _dimoverrides.text_carriers(dim)
    if len(carriers) < 2:
        return 0

    try:
        point = ref.GlobalPoint
    except Exception:
        point = None
    if point is None:
        return None

    spread = []
    for index, seg in enumerate(carriers):
        anchor = None
        for name in ("Origin", "TextPosition"):
            try:
                anchor = getattr(seg, name)
            except Exception:
                anchor = None
            if anchor is not None:
                break
        if anchor is None:
            return None
        spread.append((point.DistanceTo(anchor), index))

    spread.sort()
    # An exact tie means two segments share an origin, so the geometry has
    # nothing left to tell them apart. The tolerance is tight on purpose:
    # clicking the witness line between two real segments lands very nearly
    # equidistant from both, and that is still a legitimate read.
    if len(spread) > 1 and abs(spread[0][0] - spread[1][0]) < 1e-9:
        return None
    return spread[0][1]


def pick_dimensions(view):
    """Run the halftone pick loop until Escape, returning the chosen ids.

    The halftone is the selection UI, but it is not the selection RECORD -
    that is the order/lookup pair below. Reading the state back off the view
    would mean a dimension the user had already halftoned themselves reads
    as pre-selected, and their override would then be stripped on the first
    click. Tracking our own marks keeps the two apart.

    Everything the loop paints lives inside a TransactionGroup that is
    rolled back on the way out, so the marks disappear with it and any
    pre-existing override on a touched dimension is restored exactly. It
    also keeps the whole phase out of the undo stack, which is right: the
    user marked things up, they did not change the model.
    """
    sel = uidoc.Selection
    pick_filter = DimensionFilter()

    order = []      # ints, in pick order
    lookup = {}     # int -> Target

    group = TransactionGroup(doc, "Dim Override selection")
    group.Start()
    try:
        while True:
            try:
                with forms.WarningBar(title="Dim Override: Select dimensions "
                                            "to override, then press Escape "
                                            "to finish."):
                    ref = sel.PickObject(ObjectType.Element, pick_filter,
                                         "Select dimensions to override")
            except OperationCanceledException:
                break

            number = eid(ref.ElementId)
            dim = doc.GetElement(ref.ElementId)
            if dim is None:
                continue
            index = clicked_segment(dim, ref)

            t = Transaction(doc, "Mark dimension")
            t.Start()
            try:
                target = lookup.get(number)
                if target is None:
                    target = Target(ref.ElementId,
                                    len(_dimoverrides.text_carriers(dim)))
                    lookup[number] = target
                    order.append(number)
                    view.SetElementOverrides(ref.ElementId, HALFTONE)
                target.toggle(index)

                # The halftone marks the ELEMENT, which is the finest grain
                # a graphic override has - there is no way to halftone one
                # segment of a string. So it comes off only when the last
                # picked segment of that dimension is dropped, and while any
                # remain the whole dimension stays marked. The dialog is
                # what says how many segments are actually in play.
                if not target.segments:
                    view.SetElementOverrides(target.element_id, CLEAR_OVERRIDE)
                    order.remove(number)
                    del lookup[number]
                t.Commit()
            except Exception:
                t.RollBack()
    finally:
        group.RollBack()

    return [lookup[n] for n in order]


# ── DIALOG ──────────────────────────────────────────────────────────────────

class DimOverrideWindow(forms.WPFWindow):
    """The five text fields plus the frequency-ranked catalogue below them."""

    # --- construction ---
    def __init__(self, xaml_name, prefill, catalogue, scope, measured):
        forms.WPFWindow.__init__(self, xaml_name)
        apply_seed43_palette(self, SCRIPT_DIR)
        apply_seed43_dimensions(self, SCRIPT_DIR)

        self.result = None
        self.whole_dims = False
        self._catalogue = catalogue
        self._ranked = ranked(catalogue)
        self._prediction = None     # full text Tab would accept, or None
        self._deleting = False      # last keystroke was Back/Delete
        self._quiet = False         # suppress the change handlers

        # Baked at build time from the JSON, not TryFindResource - a Run's
        # Foreground is a real brush, not a DynamicResource target, so it
        # has to be a resolved colour either way.
        self._ghost_brush = Windows.Media.SolidColorBrush(
            Windows.Media.ColorConverter.ConvertFromString(
                get_color(SCRIPT_DIR, "text_muted", fallback="#9CA3AF")))

        self.count_tb.Text = scope.headline
        self.actual_tb.Text = (
            u'Measured value: {}   -   leave Value empty to keep it.'.format(
                measured) if measured else
            u"Leave Value empty to keep each dimension's measured value.")

        # Only offered when the selection actually contains a string. On a
        # set of plain dimensions the checkbox would be a permanently
        # meaningless control, and one that reads as if it did something.
        if scope.strings:
            self.whole_cb.Content = scope.whole_label
            self.whole_cb.Visibility = Windows.Visibility.Visible
        else:
            self.whole_cb.Visibility = Windows.Visibility.Collapsed

        self._quiet = True
        self.above_tb.Text = prefill["above"]
        self.prefix_tb.Text = prefill["prefix"]
        self.value_tb.Text = prefill["value"]
        self.suffix_tb.Text = prefill["suffix"]
        self.below_tb.Text = prefill["below"]
        self._quiet = False

        self._refresh_list(prefill["value"])
        self._set_ghost("", "")

        self.PreviewKeyDown += self._window_keydown
        self.Loaded += self._focus_value

    # --- public methods ---
    def show(self):
        """Run the dialog, returning the field dict or None if cancelled."""
        self.ShowDialog()
        return self.result

    # --- event handlers ---
    def value_changed(self, sender, args):
        if self._quiet:
            return
        typed = self.value_tb.Text or ""
        self._refresh_list(typed)

        # Predicting mid-delete fights the user: they backspace a character
        # and the prediction puts it straight back.
        if self._deleting:
            self._prediction = None
            self._set_ghost(typed, "")
            return

        match = self._predict(typed)
        self._prediction = match
        self._set_ghost(typed, match[len(typed):] if match else "")

    def value_keydown(self, sender, args):
        key = args.Key
        self._deleting = key in (Windows.Input.Key.Back,
                                 Windows.Input.Key.Delete)
        if key == Windows.Input.Key.Tab and self._prediction:
            accepted = self._prediction
            self._quiet = True
            self.value_tb.Text = accepted
            self.value_tb.CaretIndex = len(accepted)
            self._quiet = False
            self._prediction = None
            self._set_ghost(accepted, "")
            self._refresh_list(accepted)
            args.Handled = True

    def catalogue_selected(self, sender, args):
        if self._quiet:
            return
        item = self.catalogue_lb.SelectedItem
        if item is None:
            return
        chosen = item.Tag

        # Quiet, so picking from the list does not re-filter the list out
        # from under the cursor that just clicked it.
        self._quiet = True
        self.value_tb.Text = chosen
        self.value_tb.CaretIndex = len(chosen)
        self._quiet = False
        self._prediction = None
        self._set_ghost(chosen, "")

    def apply_clicked(self, sender, args):
        self.whole_dims = bool(self.whole_cb.IsChecked)
        self.result = self._collect()
        self.Close()

    def clear_clicked(self, sender, args):
        # Clearing respects the picked segments too. "Clear Override" means
        # clear what was selected, not blank the whole string - a segment
        # pick would otherwise be silently widened on the one action where
        # the damage cannot be seen in the dialog first.
        self.whole_dims = bool(self.whole_cb.IsChecked)
        self.result = dict((name, "") for name in FIELDS)
        self.Close()

    def cancel_clicked(self, sender, args):
        self.result = None
        self.Close()

    # --- private helpers ---
    def _focus_value(self, sender, args):
        self.value_tb.Focus()
        self.value_tb.CaretIndex = len(self.value_tb.Text or "")

    def _window_keydown(self, sender, args):
        if args.Key == Windows.Input.Key.Escape:
            self.cancel_clicked(sender, args)
        elif args.Key == Windows.Input.Key.Enter:
            self.apply_clicked(sender, args)

    def _collect(self):
        return {"above":  self.above_tb.Text or "",
                "prefix": self.prefix_tb.Text or "",
                "value":  self.value_tb.Text or "",
                "suffix": self.suffix_tb.Text or "",
                "below":  self.below_tb.Text or ""}

    def _predict(self, typed):
        """The highest-ranked catalogue entry typed is a prefix of."""
        if not typed:
            return None
        low = typed.lower()
        for name in self._ranked:
            if len(name) > len(typed) and name.lower().startswith(low):
                return name
        return None

    def _refresh_list(self, typed):
        """Show catalogue entries starting with typed, or all of them.

        Falling back to the whole list when nothing matches is deliberate:
        an empty panel while someone types a brand new override reads as a
        fault, and there is nothing useful to put in its place.
        """
        low = (typed or "").lower()
        names = [n for n in self._ranked if n.lower().startswith(low)]
        if not names:
            names = list(self._ranked)

        self._quiet = True
        self.catalogue_lb.Items.Clear()
        for name in names:
            self.catalogue_lb.Items.Add(self._row(name))
        self._quiet = False

    def _row(self, name):
        """One catalogue line: the text left, its usage count right."""
        item = Windows.Controls.ListBoxItem()
        item.Tag = name

        panel = Windows.Controls.DockPanel()
        count = Windows.Controls.TextBlock()
        count.Text = unicode(self._catalogue.get(name, 0))
        count.Opacity = 0.5
        # Docked children must be added before the filling one, or the
        # label takes the whole width and the count never appears.
        Windows.Controls.DockPanel.SetDock(count, Windows.Controls.Dock.Right)
        panel.Children.Add(count)

        label = Windows.Controls.TextBlock()
        label.Text = name
        panel.Children.Add(label)

        item.Content = panel
        return item

    def _set_ghost(self, typed, remainder):
        """Draw the prediction behind the field: typed part invisible, the
        predicted remainder in muted grey.

        The typed part is drawn transparent rather than omitted so the
        remainder starts at exactly the right x - the real text sits on top
        of it in the TextBox above.
        """
        self.value_ghost.Inlines.Clear()
        if not remainder:
            return
        head = Windows.Documents.Run(typed)
        head.Foreground = Windows.Media.Brushes.Transparent
        tail = Windows.Documents.Run(remainder)
        tail.Foreground = self._ghost_brush
        self.value_ghost.Inlines.Add(head)
        self.value_ghost.Inlines.Add(tail)


# ── ENTRY POINT ─────────────────────────────────────────────────────────────

class Scope(object):
    """What the picked targets add up to, in the words the dialog needs."""

    def __init__(self, targets):
        self.dims = len(targets)
        self.strings = len([t for t in targets if t.is_string])
        self.segments = sum(len(t.segments) for t in targets)
        self.partial = len([t for t in targets
                            if t.is_string
                            and len(t.segments) < t.carrier_count])

    @property
    def headline(self):
        text = u"|  {} dimension{}".format(self.dims,
                                           "" if self.dims == 1 else "s")
        # The segment count is only worth the space when it says something
        # the dimension count does not.
        if self.segments != self.dims:
            text += u", {} segment{}".format(self.segments,
                                             "" if self.segments == 1 else "s")
        return text

    @property
    def whole_label(self):
        return (u"Apply to every segment of the {} dimension string{}"
                .format(self.strings, "" if self.strings == 1 else "s"))


def apply_to_targets(targets, data, whole_dims):
    """Write data onto the picked segments, recording it. Returns failures.

    Only the carriers the user actually picked are written. An untouched
    segment of the same string is not read back and rewritten either, so
    there is no window in which its text passes through this tool at all.

    The record is stamped here, at the point of writing, rather than left
    for the tracker to discover later - this is the only moment anything
    knows for certain that the text was put there deliberately. It is
    stamped from the dimension's state AFTER the write, so a string with one
    overridden segment and eight measured ones records exactly that.

    Whether the record survives is decided by what the dimension ends up
    holding, not by what was typed into the dialog. Clearing the one segment
    that carried text drops the record; clearing one segment of three leaves
    the other two still worth tracking.
    """
    failures = []

    t = Transaction(doc, "Override dimension text")
    t.Start()
    try:
        for target in targets:
            dim = doc.GetElement(target.element_id)
            if dim is None:
                continue
            try:
                if whole_dims:
                    _dimoverrides.write_uniform(dim, data)
                else:
                    _dimoverrides.write_indexed(dim, target.segments, data)

                if _dimoverrides.has_text(_dimoverrides.read_current(dim)):
                    _dimoverrides.store_record(dim)
                else:
                    _dimoverrides.forget_record(dim)
            except Exception as ex:
                failures.append(u"{}: {}".format(eid(target.element_id), ex))
        t.Commit()
    except Exception:
        t.RollBack()
        raise
    return failures


def main():
    view = doc.ActiveView
    if not view.AreGraphicsOverridesAllowed():
        forms.alert("This view will not take graphic overrides, so the "
                    "halftone selection has nothing to draw with.\n\nRun the "
                    "tool from a plan, section, elevation or drafting view.",
                    title=TOOL_NAME, exitscript=True)

    targets = pick_dimensions(view)
    if not targets:
        return

    scope = Scope(targets)
    first = doc.GetElement(targets[0].element_id)
    first_index = min(targets[0].segments)
    catalogue = load_catalogue()

    # The measured value is only shown when one carrier is in play, since
    # it is a statement about that one number and means nothing spread
    # across several.
    single = scope.segments == 1

    window = DimOverrideWindow("DimOverride.xaml",
                               seed_fields(first, first_index),
                               catalogue,
                               scope,
                               actual_value(first, first_index) if single
                               else "")
    data = window.show()
    if data is None:
        return

    failures = apply_to_targets(targets, data, window.whole_dims)

    # Credited per dimension, not per press, so the ranking reflects how
    # much of the drawing carries each override rather than how often the
    # dialog happened to be opened.
    written = len(targets) - len(failures)
    if written > 0:
        record(catalogue, data["value"], written)
        save_catalogue(catalogue)

    if failures:
        forms.alert(u"{} of {} dimensions could not be "
                    u"written:\n\n{}".format(len(failures), len(targets),
                                             u"\n".join(failures[:10])),
                    title=TOOL_NAME)


main()
