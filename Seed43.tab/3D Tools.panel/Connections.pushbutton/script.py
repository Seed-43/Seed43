# -*- coding: utf-8 -*-
# "Connections"
# "Seed43"
# """
# Take a connection you have already configured and re-file it under a
# family and type name of your own.
#
# WHY THE TOOL HAS TO EXIST
#     Revit ships steel connections as system families named after the
#     algorithm - Front plate splice, Stiffener, Cope - and that family name
#     is read-only in every direction there is. FamilyName has no setter,
#     ALL_MODEL_FAMILY_NAME reports IsReadOnly, and Duplicate() keeps the
#     parent family. So a connection tuned for one job can only ever be a
#     type name sitting inside Autodesk's family, and the Project Browser
#     files your work under Autodesk's vocabulary instead of yours.
#
#     There is exactly one moment a family name can be set: at creation.
#     StructuralConnectionHandlerType.Create takes it as an argument. This
#     tool is that call plus the bookkeeping around it.
#
# WHAT CARRIES OVER, AND WHAT DOES NOT
#     ConnectionGuid carries, which is the one that matters: it selects the
#     Advance Steel algorithm, so a clone generates the same KIND of
#     connection as the type it came from. The identity text carries too -
#     Type Mark, Keynote, Description and the rest of CARRIED.
#
#     The Modify Parameters values DO NOT. Confirmed on a real clone: plate
#     sizes, bolt grades and layouts reset to the algorithm's defaults. They
#     live in the Advance Steel object store, which the Revit API exposes
#     neither to read nor to write - no property, no parameter, no
#     extensible storage entity - so the tool cannot carry them across and
#     cannot even report what was lost. Set the clone up once by hand.
#
# THE TRADE THIS TOOL MAKES, AND WHY THERE IS NO WAY ROUND IT
#     Two mechanisms can produce a new connection type, and each keeps the
#     half the other drops:
#
#       Create           names the family, resets the settings
#       CopyElements     keeps the settings, cannot change the family
#
#     A copy is Revit's own machinery, so the Advance Steel data rides along
#     intact - but the family name rides along with it, which is precisely
#     the thing being changed here, so a copy can only ever land back in the
#     family it came from. That is also what Duplicate() already does.
#
#     So a clone costs one trip through Modify Parameters. If that trade
#     ever stops being worth it, the answer is Duplicate, not a change to
#     this tool.
#
# Target: Revit 2022-2026, IronPython 2.
# """

# ── IMPORTS ─────────────────────────────────────────────────────────────────

import os

from Autodesk.Revit.DB import (ElementId, FilteredElementCollector,
                               Transaction)
from Autodesk.Revit.DB.Structure import (StructuralConnectionHandler,
                                         StructuralConnectionHandlerType)
from pyrevit import forms, revit
from pyrevit.framework import Windows
from System.Collections.Generic import List

from Snippets._connections import (connection_types, eid, element_name,
                                    family_name, param_text, placed_counts,
                                    read_param)
from Snippets.seed43_theme import (apply_seed43_dimensions,
                                   apply_seed43_palette, get_color)

doc = revit.doc
uidoc = revit.uidoc

# ── CONSTANTS ───────────────────────────────────────────────────────────────

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TOOL_NAME = "Connections"

# Column widths, mirrored in Connections.xaml's header row. The two must
# move together or the headings drift off the data.
COL_FAMILY = 210
COL_TYPE = 250
COL_PLACED = 62
COL_STYLE = 86
COL_ID = 74

# Identity text copied onto a clone, by parameter name.
#
# Type IfcGUID is deliberately absent. It has to stay unique per type and
# Revit issues one for the new type itself, so copying the source's would
# leave two types in the model claiming the same IFC identity - an export
# problem that would not surface until somebody opened the IFC downstream.
#
# "Detail Discription" is spelled the way the shared parameter is spelled.
# Correcting it here would simply fail to find the parameter.
CARRIED = ("Type Mark", "Keynote", "Description", "Assembly Code", "Model",
           "Manufacturer", "Type Comments", "URL", "Cost",
           "Detail Discription", "IfcExportType")


# ── HELPERS ─────────────────────────────────────────────────────────────────

def brush(key, fallback="#FFFFFF"):
    """A resolved brush for a palette key, for text built in Python."""
    return Windows.Media.SolidColorBrush(
        Windows.Media.ColorConverter.ConvertFromString(
            get_color(SCRIPT_DIR, key, fallback=fallback)))


# ── CORE LOGIC ──────────────────────────────────────────────────────────────

class Source(object):
    """One connection type in the model, as the list shows it."""

    def __init__(self, symbol, placed):
        self.symbol = symbol
        self.id = eid(symbol.Id)
        self.name = element_name(symbol)
        self.family = family_name(symbol)
        self.placed = placed
        self.mark = param_text(symbol, "Type Mark")
        self.style = style_of(symbol)


def style_of(symbol):
    """Which flavour of connection this type is, for the STYLE column.

    Generic is the placeholder Revit uses before a real connection is
    chosen, and a version of one inherits that emptiness, so it is worth
    seeing in the list before pressing the green button.
    """
    try:
        if symbol.IsGeneric():
            return u"Generic"
        if symbol.IsCustom():
            return u"Custom"
        if symbol.IsDetailed():
            return u"Detailed"
    except Exception:
        pass
    return u"-"


def survey():
    """Every connection type in the model, with how many are placed."""
    counts = placed_counts(doc)
    return [Source(symbol, counts.get(eid(symbol.Id), 0))
            for symbol in connection_types(doc)]


def instances_on(type_id):
    """Every placed connection using type_id."""
    wanted = eid(type_id)
    return [h for h in FilteredElementCollector(doc).OfClass(
        StructuralConnectionHandler) if eid(h.GetTypeId()) == wanted]


def suggest_names(row):
    """The family and type name to open the fields on.

    Type Mark is the best signal available. A type called
    "MEPS-F 100-50 460UB67.1" whose mark reads "MEPS-F 100-50" is already
    carrying the split in its own naming, so the mark becomes the family
    and the remainder becomes the type. Anything else falls back to the
    names as they stand, which is a harmless suggestion to type over.
    """
    mark = (row.mark or u"").strip()
    if mark and row.name.startswith(mark):
        rest = row.name[len(mark):].strip()
        if rest:
            return mark, rest
    return row.family, row.name


def validate(family_name, type_name, existing):
    """Why this version cannot be made, or None if it can.

    Revit publishes a validator for the type name but none for the family
    name, so the family is only checked for being non-empty and for not
    colliding with a type that already exists under it.
    """
    if not family_name:
        return u"Give the new family a name."
    if not type_name:
        return u"Give the new type a name."
    for row in existing:
        if row.family == family_name and row.name == type_name:
            return u"{} : {} already exists in this model.".format(
                family_name, type_name)
    if not StructuralConnectionHandlerType.IsTypeNameValidForCustomConnection(
            doc, type_name):
        return u"Revit will not accept '{}' as a connection type name.".format(
            type_name)
    return None


def copy_identity(source, target):
    """Copy the identity text across. Returns the names that would not take.

    Verbatim, Type Mark included. A clone that quietly rewrote the mark
    would be second-guessing the naming, and Revit's duplicate Type Mark
    warning is not a fault to design around: two types legitimately share a
    mark while a clone and the type it came from both exist.
    """
    missed = []
    for name in CARRIED:
        src = source.LookupParameter(name)
        dst = target.LookupParameter(name)
        if src is None or dst is None or dst.IsReadOnly:
            continue
        value = read_param(src)
        if value is None:
            continue
        try:
            if not dst.Set(value):
                missed.append(name)
        except Exception:
            missed.append(name)
    return missed


def clone_type(row, family_name, type_name):
    """Create the new family and type. Returns the new type.

    Its own transaction, committed before any placed connection is moved.
    A clone that exists with nobody on it is a harmless spare the user
    can delete; a half-finished move against a type that was never
    committed is not something the UI can recover from.

    The category comes off the source rather than being hardcoded. Ordinary
    connections sit in Structural Connections but copes and mitres sit in
    Sub-Connections. The list shows only the former, so this is belt and
    braces rather than a live case.
    """
    t = Transaction(doc, "Create connection version")
    t.Start()
    try:
        made = StructuralConnectionHandlerType.Create(
            doc, type_name, row.symbol.ConnectionGuid, family_name,
            row.symbol.Category.Id)
        missed = copy_identity(row.symbol, made)
        t.Commit()
    except Exception:
        t.RollBack()
        raise
    return made, missed


def move_instances(handlers, new_id):
    """Move handlers onto new_id. Returns (moved ids, refusal lines).

    GetValidTypes is asked per handler rather than once for the batch. A
    connection joins a specific set of members and Revit decides for each
    one which types can serve it, so a clone that suits the first
    connection is not thereby proven to suit the fifth.
    """
    moved = []
    refused = []
    wanted = eid(new_id)

    t = Transaction(doc, "Move connections to new version")
    t.Start()
    try:
        for handler in handlers:
            marker = eid(handler.Id)
            try:
                valid = [eid(i) for i in handler.GetValidTypes()]
            except Exception:
                valid = None
            if valid is not None and wanted not in valid:
                refused.append(
                    u"{}: Revit does not offer the new type for this "
                    u"connection".format(marker))
                continue
            try:
                handler.ChangeTypeId(new_id)
                moved.append(marker)
            except Exception as ex:
                refused.append(u"{}: {}".format(marker, ex))
        t.Commit()
    except Exception:
        t.RollBack()
        raise
    return moved, refused


def preselected():
    """The connection type id picked in Revit before launch, or None.

    Opening the tool with a connection selected is the obvious way of
    saying "this one", and it saves hunting a list for a type whose name is
    Autodesk's rather than yours.
    """
    try:
        ids = uidoc.Selection.GetElementIds()
    except Exception:
        return None
    for element_id in ids:
        element = doc.GetElement(element_id)
        if isinstance(element, StructuralConnectionHandler):
            return eid(element.GetTypeId())
        if isinstance(element, StructuralConnectionHandlerType):
            return eid(element.Id)
    return None


# ── UI ──────────────────────────────────────────────────────────────────────

class ConnectionsWindow(forms.WPFWindow):
    """The connection types in the model, and the version bench beneath."""

    # --- construction ---
    def __init__(self, xaml_name):
        forms.WPFWindow.__init__(self, xaml_name)
        apply_seed43_palette(self, SCRIPT_DIR)
        apply_seed43_dimensions(self, SCRIPT_DIR)

        # Resolved after the palette is applied, never before. A brush built
        # from an unresolved lookup comes back transparent, and the row text
        # then renders invisible with no error to explain it.
        self._muted = brush("text_muted", "#9CA3AF")
        self._primary = brush("text_primary", "#FFFFFF")
        self._green = brush("primary_green", "#208A3C")

        self._rows = []
        self.refresh(select_id=preselected())

    # --- public methods ---
    def refresh(self, note=None, select_id=None):
        """Re-read the model and rebuild the list."""
        self._rows = survey()
        families = len(set(r.family for r in self._rows))
        placed = sum(r.placed for r in self._rows)

        self.subtitle_tb.Text = u"|  {} types in {} families".format(
            len(self._rows), families)
        self.summary_tb.Text = u"{} placed connections".format(placed)
        self.explain_tb.Text = (
            u"Pick a connection, name the family and type you want, and "
            u"Clone makes one running the same algorithm. The family name "
            u"can only be set here, at creation: Revit has no way to rename "
            u"one afterwards. Modify Parameters values do NOT come across - "
            u"a clone starts on the algorithm defaults, so set its plate and "
            u"bolts once before relying on it.")

        self.rows_lb.Items.Clear()
        chosen = None
        for row in self._rows:
            item = self._build_row(row)
            self.rows_lb.Items.Add(item)
            if select_id is not None and row.id == select_id:
                chosen = item
        if chosen is not None:
            self.rows_lb.SelectedItem = chosen
            self.rows_lb.ScrollIntoView(chosen)

        self._sync_bench()
        self.status_tb.Text = note or u""

    # --- event handlers ---
    def row_selected(self, sender, args):
        self._sync_bench()

    def reload_clicked(self, sender, args):
        self.refresh(u"Re-read from the model as it stands now.")

    def show_clicked(self, sender, args):
        """Select the placed connections that use the highlighted type."""
        row = self._selected()
        if row is None:
            self.status_tb.Text = u"Highlight a connection type first."
            return
        handlers = instances_on(row.symbol.Id)
        if not handlers:
            self.status_tb.Text = (
                u"{} : {} is not placed anywhere, so there is nothing to "
                u"show.".format(row.family, row.name))
            return
        try:
            ids = List[ElementId]()
            for handler in handlers:
                ids.Add(handler.Id)
            uidoc.Selection.SetElementIds(ids)
            uidoc.ShowElements(ids)
        except Exception as ex:
            self.status_tb.Text = u"Could not show them: {}".format(ex)
            return
        self.status_tb.Text = u"Selected {} placed connection(s).".format(
            len(handlers))

    def clone_clicked(self, sender, args):
        """Make the clone, and optionally move the placed connections."""
        row = self._selected()
        if row is None:
            self.status_tb.Text = u"Highlight the connection to clone."
            return

        family_name = (self.family_tb.Text or u"").strip()
        type_name = (self.type_tb.Text or u"").strip()
        problem = validate(family_name, type_name, self._rows)
        if problem:
            self.status_tb.Text = problem
            return

        try:
            made, missed = clone_type(row, family_name, type_name)
        except Exception as ex:
            forms.alert(u"The clone was not created and nothing in the "
                        u"model was changed:\n\n{}".format(ex),
                        title=TOOL_NAME)
            return

        parts = [u"Created {} : {}.".format(family_name, type_name)]
        if missed:
            parts.append(u"{} did not copy.".format(u", ".join(missed)))

        # A second transaction on purpose: see clone_type. If the move
        # fails the clone still exists to be tried again against.
        if self.move_cb.IsChecked and row.placed:
            handlers = instances_on(row.symbol.Id)
            try:
                moved, refused = move_instances(handlers, made.Id)
            except Exception as ex:
                forms.alert(u"The clone was created, but moving the placed "
                            u"connections onto it failed and was rolled "
                            u"back:\n\n{}".format(ex), title=TOOL_NAME)
                moved, refused = [], []
            parts.append(u"Moved {} of {}.".format(len(moved), len(handlers)))
            if refused:
                forms.alert(
                    u"{} connection(s) stayed on the old type:\n\n{}".format(
                        len(refused), u"\n".join(refused[:10])),
                    title=TOOL_NAME)

        parts.append(u"The clone is on the algorithm defaults: set its "
                     u"Modify Parameters before relying on it.")
        self.refresh(u" ".join(parts), select_id=eid(made.Id))

    def close_clicked(self, sender, args):
        self.Close()

    # --- private helpers ---
    def _selected(self):
        item = self.rows_lb.SelectedItem
        return item.Tag if item is not None else None

    def _sync_bench(self):
        """Point the name fields and the move checkbox at the highlighted row."""
        row = self._selected()
        if row is None:
            self.family_tb.Text = u""
            self.type_tb.Text = u""
            self.move_cb.Content = u"Move the placed connections onto the clone"
            self.move_cb.IsEnabled = False
            self.move_cb.IsChecked = False
            self.clone_btn.IsEnabled = False
            return

        family_name, type_name = suggest_names(row)
        self.family_tb.Text = family_name
        self.type_tb.Text = type_name

        self.move_cb.Content = (
            u"Move the {} placed connection{} onto the clone".format(
                row.placed, u"" if row.placed == 1 else u"s")
            if row.placed else
            u"Nothing is placed on this type to move")
        self.move_cb.IsEnabled = bool(row.placed)
        if not row.placed:
            self.move_cb.IsChecked = False
        self.clone_btn.IsEnabled = True

    def _build_row(self, row):
        """One list row: a horizontal strip of fixed-width cells."""
        panel = Windows.Controls.StackPanel()
        panel.Orientation = Windows.Controls.Orientation.Horizontal
        panel.Children.Add(self._cell(row.family, COL_FAMILY, self._muted))
        panel.Children.Add(self._cell(row.name, COL_TYPE, self._primary,
                                      bold=True))
        panel.Children.Add(self._cell(
            unicode(row.placed) if row.placed else u"-", COL_PLACED,
            self._green if row.placed else self._muted))
        panel.Children.Add(self._cell(row.style, COL_STYLE, self._muted))
        panel.Children.Add(self._cell(unicode(row.id), COL_ID, self._muted))
        panel.Children.Add(self._cell(row.mark, 0, self._muted))

        item = Windows.Controls.ListBoxItem()
        item.Content = panel
        item.Tag = row
        return item

    def _cell(self, text, width, colour=None, bold=False):
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
    ConnectionsWindow("Connections.xaml").ShowDialog()


main()
