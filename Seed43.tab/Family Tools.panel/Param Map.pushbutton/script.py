# -*- coding: utf-8 -*-
# "Param Map"
# "Seed43"
# """
# Move placed instances from an old family onto its updated replacement when
# the update renamed the parameters.
#
# A TYPE CHANGE ACROSS FAMILIES KEEPS NOTHING
#     Measured in a live 2026 model, 39 detail items moved between two
#     versions of the same family: ChangeTypeId reset EVERY family-declared
#     instance parameter to the new family's default, including the ones
#     whose name had not changed at all. "Bar Diameter db" exists verbatim in
#     both families and still came across 0.03281 -> 0.06562. Only the
#     element's own built-ins - Mark, Comments, the IFC set - survived.
#
#     So a parameter is not safe just because nobody renamed it. Auto-match
#     pairs identical names with themselves for exactly this reason, and
#     "Hide identical" only takes them off the screen; they stay mapped.
#     Anything genuinely left unmapped comes out at the new family's default.
#
#     This tool therefore reads every mapped value off each instance BEFORE
#     the swap and writes them back afterwards.
#
# WHICH PARAMETERS ARE LISTED
#     The family's OWN parameters, read out of FamilyManager, instance and
#     type. Built-ins - Mark, Comments, the IFC set - are deliberately not
#     listed: they live on the element rather than the family and survive a
#     type change untouched, so mapping them would be noise.
#
# THE DAMAGE IS INVISIBLE UNTIL COMMIT
#     Swapping 259 detail items in the Nagel 2026 template raises about 180
#     failures - dimension references gone invalid, references no longer
#     parallel, alignment constraints unsatisfied, two dimension segments
#     removed along with their override text. Revit resolves those by
#     deleting dimensions and dropping overrides.
#
#     None of it can be seen before the commit. With the swap applied and the
#     document regenerated, inside the open transaction, every dimension
#     reference count, segment count and override string is unchanged and
#     Document.GetWarnings() returns zero. An invalid reference is still a
#     reference, and the text is not dropped until Revit resolves the
#     failure, which happens during Commit().
#
#     So the guard is an IFailuresPreprocessor on the commit, and nothing
#     else. A before-and-after comparison inside the transaction reports a
#     clean swap and then commits the damage. See the note above
#     SwapFailures.
#
# WHEN THE COMMIT IS REFUSED
#     All or nothing is the wrong answer to 188 failures: either nothing
#     moves, or 159 dimensions go in one commit and nobody sees which
#     drawings changed. So a refusal offers the one-at-a-time pass instead.
#
#     Every element is retried in its own transaction. The ones nothing is
#     dimensioned to commit clean and are left alone; the rest are queued
#     with the failures Revit raised against each, and walked with the model
#     open - view activated, element framed, cost listed - so the dimensions
#     and tags can be redrawn on the spot. That window is modeless, which is
#     why this script declares __persistentengine__.
#
# TYPE PARAMETERS WRITE ONCE
#     An instance parameter is carried per instance. A type parameter belongs
#     to the type, so it is read off the source type and written once onto the
#     target type, and that lands on every instance of the target - including
#     any that were already placed. Preview says so before it happens.
#
# Target: Revit 2022-2026, IronPython 2.
# """

# The one-at-a-time pass leaves a MODELESS window open, so the engine has
# to still be alive when a button on it is clicked minutes later. Without
# this the walk dies the moment the script returns.
__persistentengine__ = True


# ── IMPORTS ─────────────────────────────────────────────────────────────────

import json
import os

from Autodesk.Revit.DB import (BuiltInParameter, CurveElement, Dimension,
                               ElementId, Family, FailureProcessingResult,
                               FailureSeverity, FilteredElementCollector,
                               IFailuresPreprocessor, IndependentTag,
                               ReferencePlane, StorageType, Transaction,
                               TransactionStatus)
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from System.Collections.Generic import List
from pyrevit import forms, revit
from pyrevit.framework import Windows

# eid() and element_name() are general-purpose and already live in
# _connections.py. Reaching into a steel-connections module for them reads
# oddly, but duplicating either one here would be worse.
from Snippets._connections import eid, element_name
from Snippets._dialogs import confirm, message
# The read/write half of _dimoverrides only. Its Extensible Storage half is
# for Dim Override's long-term tracking; a swap has no business stamping a
# permanent record onto every dimension it passes.
from Snippets._dimoverrides import (has_text, is_target, read_current,
                                    summarise, write_segments)
from Snippets._userdata import user_path
from Snippets.seed43_theme import (apply_seed43_dimensions,
                                   apply_seed43_palette, get_color)

doc = revit.doc
uidoc = revit.uidoc
app = doc.Application


# ── CONSTANTS ───────────────────────────────────────────────────────────────

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TOOL_NAME = "Param Map"

MAPS_FILE = user_path(TOOL_NAME, "maps.json")
GUIDE_FILE = user_path(TOOL_NAME, "guide_shown.json")

SKIP = u"—"            # em dash, the "not mapped" marker in both lists

TITLE_WALK = u"Param Map – one at a time"

# Column widths, mirrored in ParamMap.xaml's two header rows. The two must
# move together or the headings drift off the data.
COL_KIND = 42
COL_NAME = 200

GUIDE = (
    u"Param Map carries values across a family swap.\n\n"
    u"Pick the old family on the left and its replacement on the right, "
    u"pair each type, then pair the parameters. Auto-match does most of it: "
    u"it reads a trailing dot as One Side and two dots as Other Side.\n\n"
    u"A parameter left on the dash is skipped, and a skipped parameter comes "
    u"out at the new family's default - a type change across families keeps "
    u"nothing of its own, not even a name that did not change.\n\n"
    u"Preview reports what would happen without touching the model.\n\n"
    u"Skip formulas is on by default: a parameter a formula drives is "
    u"left out of the mapping, because the new family works it out for "
    u"itself.\n\n"
    u"Dimensions are watched. A family swap can cost them - references "
    u"gone invalid, segments removed, override text dropped - and none "
    u"of that shows until Revit commits.\n\n"
    u"So Swap goes element by element. Each one is tried on its own, the "
    u"ones Revit has no objection to are swapped and kept, and the rest "
    u"are walked: each step opens the view holding the dimensions and "
    u"tags at risk, frames the element and says what taking it would "
    u"cost. That window stays open beside Revit, so you can redraw what "
    u"went before clicking Next. Stop whenever you like - everything "
    u"already swapped stays.\n\n"
    u"Tick Accept fixes to take the whole lot in one commit instead, with "
    u"Revit applying its own fix to every failure. That is one undo, and "
    u"all the redrawing is left until afterwards."
)


# ── SMALL HELPERS ───────────────────────────────────────────────────────────

def read_json(path, default):
    try:
        with open(path, "r") as handle:
            return json.load(handle)
    except Exception:
        return default


def write_json(path, data):
    """Never raises: losing a saved mapping is not worth an exception."""
    try:
        folder = os.path.dirname(path)
        if folder and not os.path.isdir(folder):
            os.makedirs(folder)
        with open(path, "w") as handle:
            json.dump(data, handle, indent=2)
    except Exception:
        pass


# ── PARAMETER NAMES ─────────────────────────────────────────────────────────

def canon(name):
    """(base, side) for a parameter name, with the side marker taken off.

    The naming convention this tool exists for: the updated family writes one
    side as a trailing "." and the other as "..", where the old family spelled
    them out as "One Side" and "Other Side". The words can sit anywhere in the
    name - "Flip One Side Other Way" became "Flip Other Way." - so they are
    cut out of the middle rather than off the end.

    side is 0 for neither, 1 for one side, 2 for the other.
    """
    text = u" ".join(unicode(name).split())

    side = 0
    if text.endswith(u".."):
        side, text = 2, text[:-2]
    elif text.endswith(u"."):
        side, text = 1, text[:-1]

    lowered = text.lower()
    for token, marker in ((u" other side", 2), (u" one side", 1)):
        at = lowered.find(token)
        if at >= 0:
            side = marker
            text = text[:at] + text[at + len(token):]
            break

    return (u" ".join(text.split()).lower(), side)


def auto_pairs(sources, targets, skip_formula=False):
    """{source name: target name} paired by name, exact first then canonical.

    One target is used once. Exact matches are claimed before canonical ones
    so a family that renamed only some of its parameters cannot have an
    untouched name stolen by a near miss somewhere else in the list.

    skip_formula leaves out anything a formula drives, on either side. It is
    on by default in the window: pairing one is at best a wasted write and at
    worst a stale number sitting on top of a live formula.
    """
    pairs = {}
    taken = set()

    if skip_formula:
        sources = [i for i in sources if not i.formula]
        targets = [i for i in targets if not i.formula]

    by_name = {}
    by_canon = {}
    for info in targets:
        by_name.setdefault(info.name.lower(), info.name)
        by_canon.setdefault((info.kind,) + canon(info.name), info.name)

    for info in sources:
        hit = by_name.get(info.name.lower())
        if hit and hit not in taken:
            pairs[info.name] = hit
            taken.add(hit)

    for info in sources:
        if info.name in pairs:
            continue
        hit = by_canon.get((info.kind,) + canon(info.name))
        if hit and hit not in taken:
            pairs[info.name] = hit
            taken.add(hit)

    return pairs


# ── READING A FAMILY ────────────────────────────────────────────────────────

class ParamInfo(object):
    """One parameter a family declares for itself."""

    def __init__(self, name, kind, storage, writable, note, formula=False):
        self.name = name
        self.kind = kind            # "inst" or "type"
        self.storage = storage      # StorageType
        self.writable = writable    # False for formula-driven and reporting
        self.note = note            # why it cannot be written, or u""
        # Held apart from writable because it is the one the user filters on.
        # A formula-driven parameter is worth skipping on BOTH sides: as a
        # target it cannot be written at all, and as a source its value was
        # computed rather than authored, so carrying it either duplicates
        # what the new family's own formula will work out or freezes a stale
        # number on top of it.
        self.formula = formula

    @property
    def label(self):
        return u"INST" if self.kind == "inst" else u"TYPE"


def open_family(family):
    """(family document, close it afterwards) for a loaded family.

    A family the user already has open in the editor is handed back by
    EditFamily as that same document. Closing it would take their window
    down with it, so it is found first and left alone.
    """
    for other in app.Documents:
        try:
            if other.IsFamilyDocument and other.Title.startswith(family.Name):
                return other, False
        except Exception:
            continue
    return doc.EditFamily(family), True


def family_params(family):
    """([ParamInfo], problem) for a family, its own parameters only."""
    try:
        fam_doc, close_after = open_family(family)
    except Exception as ex:
        return [], u"Could not open {}: {}".format(family.Name, ex)

    rows = []
    try:
        for fp in fam_doc.FamilyManager.Parameters:
            # One unreadable parameter must not cost the whole list. Every
            # property here is on FamilyParameter, but Formula and
            # IsReporting throw on some shared and built-in ones.
            try:
                name = fp.Definition.Name
                kind = "inst" if fp.IsInstance else "type"
                storage = fp.StorageType
            except Exception:
                continue

            note = u""
            for probe, wording in (("Formula", u"driven by a formula"),
                                   ("IsDeterminedByFormula",
                                    u"driven by a formula"),
                                   ("IsReporting", u"a reporting parameter")):
                try:
                    if getattr(fp, probe, None):
                        note = wording
                        break
                except Exception:
                    continue

            rows.append(ParamInfo(name=name, kind=kind, storage=storage,
                                  writable=not note, note=note,
                                  formula=note == u"driven by a formula"))
    finally:
        if close_after:
            try:
                fam_doc.Close(False)
            except Exception:
                pass

    rows.sort(key=lambda r: (r.kind != "inst", r.name.lower()))
    return rows, None


class RefInfo(object):
    """One named reference on a family - something a dimension can hold onto.

    kind separates a reference plane from a reference line so the matcher
    cannot pair one with the other, the same way it keeps instance and type
    parameters apart.
    """

    def __init__(self, name, kind, setting):
        self.name = name        # the user's name, e.g. "Bar Face One Side"
        self.kind = kind        # "plane" or "line"
        self.setting = setting  # Revit's Is Reference wording, if it gave one
        # auto_pairs reads this when asked to skip formula-driven entries.
        # A reference has no formula, so it is never skipped.
        self.formula = False

    @property
    def label(self):
        return u"PLANE" if self.kind == "plane" else u"LINE"


# The Is Reference setting. Which BuiltInParameter carries it has moved
# about, so every candidate is tried and the first that answers is used.
REF_SETTING_PARAMS = ("ELEM_REFERENCE_NAME", "ELEM_IS_REFERENCE",
                      "ELEM_REFERENCE_NAME_TEXT")

NOT_A_REFERENCE = u"not a reference"


def _ref_setting(element):
    """(setting wording, is it usable) for an element's Is Reference.

    Unreadable is treated as usable. Being unable to read the setting is not
    evidence that a dimension cannot attach, and dropping a reference on that
    basis would quietly shrink what can be rebuilt.
    """
    for name in REF_SETTING_PARAMS:
        builtin = getattr(BuiltInParameter, name, None)
        if builtin is None:
            continue
        try:
            param = element.get_Parameter(builtin)
            if param is None:
                continue
            wording = param.AsValueString() or u""
        except Exception:
            continue
        if not wording:
            continue
        return wording, wording.strip().lower() != NOT_A_REFERENCE
    return u"", True


def family_references(family):
    """([RefInfo], problem) - what a dimension on this family can hold onto.

    Read out of the family document rather than off a placed instance,
    because it is the NAMES that have to be paired. A dimension the update
    broke was holding a reference whose name the update changed, and the
    name is the only thing the two versions of the family have in common.

    Anything set to "not a reference" is left out: a dimension cannot be
    attached to one, so it cannot be rebuilt onto one either.
    """
    try:
        fam_doc, close_after = open_family(family)
    except Exception as ex:
        return [], u"Could not open {}: {}".format(family.Name, ex)

    rows = []
    try:
        for kind, cls in (("plane", ReferencePlane), ("line", CurveElement)):
            try:
                found = list(FilteredElementCollector(fam_doc).OfClass(cls))
            except Exception:
                continue
            for element in found:
                try:
                    name = element.Name
                except Exception:
                    continue
                if not name or not unicode(name).strip():
                    continue
                setting, usable = _ref_setting(element)
                if not usable:
                    continue
                rows.append(RefInfo(unicode(name).strip(), kind, setting))
    finally:
        if close_after:
            try:
                fam_doc.Close(False)
            except Exception:
                pass

    seen, unique = set(), []
    for row in rows:
        key = (row.kind, row.name.lower())
        if key in seen:
            continue
        seen.add(key)
        unique.append(row)
    unique.sort(key=lambda r: (r.kind, r.name.lower()))
    return unique, None


# ── READING THE MODEL ───────────────────────────────────────────────────────

class FamilyRow(object):
    """A loaded family, its types, and how many elements are placed on it."""

    def __init__(self, family):
        self.family = family
        self.name = family.Name
        try:
            self.category = family.FamilyCategory.Name
            # Held as an integer, not an ElementId: comparing two ElementIds
            # with == goes through .NET's operator overload, and the integer
            # is the one thing that reads the same on every Revit here.
            self.category_id = eid(family.FamilyCategory.Id)
        except Exception:
            self.category = u"(no category)"
            self.category_id = -1
        self.symbols = []       # [(name, FamilySymbol)], sorted by name
        self.placed = 0

    @property
    def caption(self):
        return u"{}   –   {}  ({} placed)".format(
            self.name, self.category, self.placed)


def loaded_families():
    """Every loadable family in the model, with its types and placed counts.

    One collector pass builds the counts. The families are read into a list
    first: building a second collector inside a loop over the first one makes
    the outer loop stop early, silently.
    """
    families = FilteredElementCollector(doc).OfClass(Family).ToElements()

    rows = {}
    symbol_owner = {}
    for family in families:
        try:
            row = FamilyRow(family)
        except Exception:
            continue
        for symbol_id in list(family.GetFamilySymbolIds()):
            symbol = doc.GetElement(symbol_id)
            if symbol is None:
                continue
            row.symbols.append((element_name(symbol), symbol))
            symbol_owner[eid(symbol_id)] = row
        if not row.symbols:
            continue
        row.symbols.sort(key=lambda pair: pair[0].lower())
        rows[eid(family.Id)] = row

    for element in FilteredElementCollector(doc).WhereElementIsNotElementType():
        try:
            type_id = element.GetTypeId()
        except Exception:
            continue
        if type_id is None or eid(type_id) <= 0:
            continue
        owner = symbol_owner.get(eid(type_id))
        if owner is not None:
            owner.placed += 1

    return sorted(rows.values(), key=lambda r: (r.category.lower(),
                                                r.name.lower()))


def placed_on(symbols, selected_ids=None):
    """Every element placed on the given symbols, by category to stay quick."""
    wanted = set(eid(symbol.Id) for symbol in symbols)
    if not wanted:
        return []

    # The categories are kept as the ElementId objects Revit handed over
    # rather than rebuilt from their integers: ElementId's constructor took an
    # int up to 2023 and a long from 2024, and there is no reason to pick.
    categories = {}
    for symbol in symbols:
        try:
            categories[eid(symbol.Category.Id)] = symbol.Category.Id
        except Exception:
            pass

    found = []
    for category_id in categories.values():
        collector = FilteredElementCollector(doc).OfCategoryId(
            category_id).WhereElementIsNotElementType()
        for element in collector:
            try:
                if eid(element.GetTypeId()) in wanted:
                    found.append(element)
            except Exception:
                continue

    if selected_ids is not None:
        found = [e for e in found if eid(e.Id) in selected_ids]
    return found


# ── VALUES ──────────────────────────────────────────────────────────────────

def read_value(param):
    """(kind, value) ready to write elsewhere, or None if it cannot be read."""
    if param is None:
        return None
    storage = param.StorageType
    try:
        if storage == StorageType.Double:
            return ("d", param.AsDouble())
        if storage == StorageType.Integer:
            return ("i", param.AsInteger())
        if storage == StorageType.String:
            return ("s", param.AsString() or u"")
        if storage == StorageType.ElementId:
            return ("e", param.AsElementId())
    except Exception:
        return None
    return None


def write_value(param, kind, value):
    """Write a read_value() pair into a parameter, converting if it has to.

    Doubles carry Revit's internal units on both sides, so a like-for-like
    write needs no conversion. A cross-storage write is the user's own
    pairing and is done as literally as it can be, or refused.
    """
    if param is None or param.IsReadOnly:
        return False
    storage = param.StorageType
    try:
        if storage == StorageType.Double:
            if kind == "d":
                return param.Set(value)
            if kind == "i":
                return param.Set(float(value))
            if kind == "s":
                try:
                    return param.Set(float(value))
                except ValueError:
                    return param.SetValueString(value)
            return False
        if storage == StorageType.Integer:
            if kind == "i":
                return param.Set(value)
            if kind == "d":
                return param.Set(int(round(value)))
            if kind == "s":
                try:
                    return param.Set(int(round(float(value))))
                except ValueError:
                    return False
            return False
        if storage == StorageType.String:
            if kind == "s":
                return param.Set(value)
            if kind == "e":
                return param.Set(unicode(eid(value)))
            return param.Set(unicode(value))
        if storage == StorageType.ElementId:
            if kind == "e":
                return param.Set(value)
            return False
    except Exception:
        return False
    return False


# ── UI ──────────────────────────────────────────────────────────────────────

def brush(key, fallback="#FFFFFF"):
    """A resolved brush for a palette key, for text built in Python."""
    return Windows.Media.SolidColorBrush(
        Windows.Media.ColorConverter.ConvertFromString(
            get_color(SCRIPT_DIR, key, fallback=fallback)))


class ParamMapWindow(forms.WPFWindow):
    """Old family on the left, new one on the right, pairings between."""

    # --- construction ---
    def __init__(self, xaml_name):
        forms.WPFWindow.__init__(self, xaml_name)
        apply_seed43_palette(self, SCRIPT_DIR)
        apply_seed43_dimensions(self, SCRIPT_DIR)

        # Set when the user asks for the one-at-a-time pass. main() picks it
        # up after this window has closed, because the walk's own window is
        # modeless and one opened behind a modal window is unreachable.
        self.walk = None

        # Resolved after the palette is applied, never before. A brush built
        # from an unresolved lookup comes back transparent, and the row text
        # then renders invisible with no error to explain it.
        self._muted = brush("text_muted", "#9CA3AF")
        self._primary = brush("text_primary", "#FFFFFF")
        self._green = brush("primary_green", "#208A3C")
        self._warn = brush("warning", "#E5A50A")

        self.families = []
        self.source = None          # FamilyRow
        self.target = None          # FamilyRow
        self.src_params = []
        self.tgt_params = []
        self.src_refs = []          # what a dimension can hold onto
        self.tgt_refs = []
        self.param_map = {}         # source param name -> target param name
        self.type_map = {}          # source type name -> target type name
        self._loading = False
        self._scope_count = 0       # recounted only when the scope moves

        self.families = loaded_families()
        self.fill_family_lists()
        self.refresh()
        self.status_tb.Text = (
            u"{} families loaded. Pick the old one on the left and its "
            u"replacement on the right.".format(len(self.families)))

    # --- the family pickers ---
    def fill_family_lists(self):
        """Fill both dropdowns, each narrowed by its own search box.

        The target list is held to the source's category: a type change
        across categories is not something Revit will do, so offering it
        would only produce a failure later.
        """
        self._loading = True
        try:
            self._fill_one(self.src_family_cb, self.src_search_tb.Text,
                           [r for r in self.families if r.placed],
                           self.source)
            allowed = self.families
            if self.source is not None:
                allowed = [r for r in self.families
                           if r.category_id == self.source.category_id
                           and r.name != self.source.name]
            self._fill_one(self.tgt_family_cb, self.tgt_search_tb.Text,
                           allowed, self.target)
        finally:
            self._loading = False

    def _fill_one(self, combo, search, rows, keep):
        needle = (search or u"").strip().lower()
        if needle:
            rows = [r for r in rows if needle in r.name.lower()
                    or needle in r.category.lower()]
        combo.Items.Clear()
        chosen = -1
        for index, row in enumerate(rows):
            item = Windows.Controls.ComboBoxItem()
            item.Content = row.caption
            item.Tag = row
            combo.Items.Add(item)
            if keep is not None and row.name == keep.name:
                chosen = index
        combo.SelectedIndex = chosen

    def picked(self, combo):
        item = combo.SelectedItem
        return item.Tag if item is not None else None

    # --- loading the two sides ---
    def load_sides(self):
        """Read both families' parameters and restore any saved mapping."""
        self.src_params, self.tgt_params = [], []
        self.src_refs, self.tgt_refs = [], []
        self.param_map, self.type_map = {}, {}

        problems = []
        if self.source is not None:
            self.src_params, problem = family_params(self.source.family)
            if problem:
                problems.append(problem)
            self.src_refs, _problem = family_references(self.source.family)
        if self.target is not None:
            self.tgt_params, problem = family_params(self.target.family)
            if problem:
                problems.append(problem)
            self.tgt_refs, _problem = family_references(self.target.family)

        if self.source is not None and self.target is not None:
            saved = read_json(MAPS_FILE, {}).get(self.map_key(), {})
            if saved:
                self.param_map = dict(saved.get("params", {}))
                self.type_map = dict(saved.get("types", {}))
                note = u"Restored the mapping saved for this pair."
            else:
                self.auto_match()
                note = u"Auto-matched by name. Check it before swapping."
            self.prune_mapping()
            problems.append(note)

        self.build_type_rows()
        self.count_scope()
        self.refresh(u"  ".join(problems) if problems else None)

    def map_key(self):
        return u"{}||{}".format(self.source.name, self.target.name)

    def prune_mapping(self):
        """Drop pairings naming a parameter or type that is no longer there."""
        src_names = set(p.name for p in self.src_params)
        tgt_names = set(p.name for p in self.tgt_params)
        self.param_map = dict(
            (k, v) for k, v in self.param_map.items()
            if k in src_names and v in tgt_names)

        src_types = set(name for name, _ in self.source.symbols) \
            if self.source else set()
        tgt_types = set(name for name, _ in self.target.symbols) \
            if self.target else set()
        self.type_map = dict(
            (k, v) for k, v in self.type_map.items()
            if k in src_types and v in tgt_types)

    def auto_match(self):
        self.param_map = auto_pairs(self.src_params, self.tgt_params,
                                    skip_formula=self.skip_formula)
        if self.source is None or self.target is None:
            return
        self.type_map = {}
        targets_by_name = dict((name.lower(), name)
                               for name, _ in self.target.symbols)
        taken = set()
        for name, _ in self.source.symbols:
            hit = targets_by_name.get(name.lower())
            if hit and hit not in taken:
                self.type_map[name] = hit
                taken.add(hit)
        # A single type on each side is the common case after an update, and
        # the names rarely survive it. Pair them anyway.
        if (not self.type_map and len(self.source.symbols) == 1
                and len(self.target.symbols) == 1):
            self.type_map[self.source.symbols[0][0]] = \
                self.target.symbols[0][0]

    # --- the type strip ---
    def build_type_rows(self):
        """One row per source type: its name, and a dropdown of targets."""
        self.type_rows_sp.Children.Clear()
        if self.source is None or self.target is None:
            self.type_rows_sp.Children.Add(
                self._text(u"Pick both families to pair their types.",
                           colour=self._muted))
            return

        target_names = [name for name, _ in self.target.symbols]
        for name, _symbol in self.source.symbols:
            row = Windows.Controls.StackPanel()
            row.Orientation = Windows.Controls.Orientation.Horizontal
            row.Margin = Windows.Thickness(0, 2, 0, 2)

            row.Children.Add(self._text(name, width=260,
                                        colour=self._primary))
            row.Children.Add(self._text(u"→", width=22,
                                        colour=self._green, bold=True))

            combo = Windows.Controls.ComboBox()
            combo.Width = 300
            combo.Tag = name
            style = self.TryFindResource("Combo")
            if style is not None:
                combo.Style = style
            combo.Items.Add(SKIP + u"  skip this type")
            for target_name in target_names:
                combo.Items.Add(target_name)
            chosen = self.type_map.get(name)
            combo.SelectedIndex = (target_names.index(chosen) + 1
                                   if chosen in target_names else 0)
            combo.SelectionChanged += self.type_row_changed
            row.Children.Add(combo)

            self.type_rows_sp.Children.Add(row)

    def type_row_changed(self, sender, args):
        if self._loading:
            return
        source_type = sender.Tag
        if sender.SelectedIndex <= 0:
            self.type_map.pop(source_type, None)
        else:
            self.type_map[source_type] = sender.SelectedItem
        self.count_scope()
        self.refresh()

    # --- the two parameter lists ---
    def refresh(self, note=None):
        """Rebuild both lists from the current mapping."""
        used = set(self.param_map.values())
        hide_same = bool(self.hide_matched_cb.IsChecked)

        self.src_lb.Items.Clear()
        shown = 0
        for info in self.src_params:
            paired = self.param_map.get(info.name)
            if hide_same and paired == info.name:
                continue
            shown += 1
            self.src_lb.Items.Add(self._param_item(info, paired))

        self.tgt_lb.Items.Clear()
        back = dict((v, k) for k, v in self.param_map.items())
        for info in self.tgt_params:
            paired = back.get(info.name)
            if hide_same and paired == info.name:
                continue
            self.tgt_lb.Items.Add(self._param_item(info, paired))

        self.src_head_tb.Text = (
            u"{}  ({} of {} shown, {} mapped)".format(
                self.source.name, shown, len(self.src_params),
                len(self.param_map))
            if self.source else u"Source  (none picked)")
        self.tgt_head_tb.Text = (
            u"{}  ({} parameters, {} taken)".format(
                self.target.name, len(self.tgt_params), len(used))
            if self.target else u"Target  (none picked)")
        self.subtitle_tb.Text = (
            u"|  {} instances".format(self._scope_count)
            if self.source else u"|  nothing picked")

        self.sync_buttons()
        if note:
            self.status_tb.Text = note

    def _param_item(self, info, paired):
        row = Windows.Controls.StackPanel()
        row.Orientation = Windows.Controls.Orientation.Horizontal
        row.Tag = info.name

        row.Children.Add(self._text(info.label, width=COL_KIND,
                                    colour=self._muted))
        row.Children.Add(self._text(info.name, width=COL_NAME,
                                    colour=self._primary,
                                    tip=self._param_tip(info)))
        if paired:
            row.Children.Add(self._text(paired, colour=self._green))
        else:
            row.Children.Add(self._text(
                SKIP, colour=self._muted,
                tip=u"Not mapped. Its value is not carried across."))
        if not info.writable:
            row.Children.Add(self._text(u"  ⚠ " + info.note,
                                        colour=self._warn))
        return row

    def _param_tip(self, info):
        kind = u"instance" if info.kind == "inst" else u"type"
        tip = u"{} parameter, stored as {}".format(kind, info.storage)
        if info.note:
            tip += u", " + info.note
        return tip

    def _text(self, text, width=None, colour=None, bold=False, tip=None):
        block = Windows.Controls.TextBlock()
        block.Text = text
        block.VerticalAlignment = Windows.VerticalAlignment.Center
        block.TextTrimming = Windows.TextTrimming.CharacterEllipsis
        block.Margin = Windows.Thickness(0, 0, 10, 0)
        block.ToolTip = tip or text or None
        if width:
            block.Width = width
        if colour is not None:
            block.Foreground = colour
        if bold:
            block.FontWeight = Windows.FontWeights.SemiBold
        return block

    def selected_name(self, listbox):
        item = listbox.SelectedItem
        return item.Tag if item is not None else None

    # --- options ---
    @property
    def skip_formula(self):
        return bool(self.skip_formula_cb.IsChecked)

    # --- scope ---
    def selection_ids(self):
        if not self.sel_only_cb.IsChecked:
            return None
        return set(eid(i) for i in uidoc.Selection.GetElementIds())

    def count_scope(self):
        """Recount the elements in scope. A collector pass, so not on every
        refresh - refresh() reads the cached number instead."""
        if self.source is None:
            self._scope_count = 0
            return 0
        symbols = [s for name, s in self.source.symbols
                   if name in self.type_map]
        if not symbols:
            symbols = [s for _name, s in self.source.symbols]
        self._scope_count = len(placed_on(symbols, self.selection_ids()))
        return self._scope_count

    def sync_buttons(self):
        ready = (self.source is not None and self.target is not None
                 and bool(self.type_map))
        self.swap_btn.IsEnabled = ready
        self.preview_btn.IsEnabled = ready
        both = self.source is not None and self.target is not None
        self.auto_btn.IsEnabled = both
        self.clear_btn.IsEnabled = both
        self.map_btn.IsEnabled = (
            self.selected_name(self.src_lb) is not None
            and self.selected_name(self.tgt_lb) is not None)
        self.skip_btn.IsEnabled = self.selected_name(self.src_lb) is not None

    # --- event handlers ---
    def src_search_changed(self, sender, args):
        if not self._loading:
            self.fill_family_lists()

    def tgt_search_changed(self, sender, args):
        if not self._loading:
            self.fill_family_lists()

    def src_family_changed(self, sender, args):
        if self._loading:
            return
        self.source = self.picked(self.src_family_cb)
        self.fill_family_lists()
        self.load_sides()

    def tgt_family_changed(self, sender, args):
        if self._loading:
            return
        self.target = self.picked(self.tgt_family_cb)
        self.load_sides()

    def selection_changed(self, sender, args):
        self.sync_buttons()

    def filter_changed(self, sender, args):
        self.refresh()

    def formula_changed(self, sender, args):
        """Turning the filter on drops the pairings it would have prevented."""
        if not self.skip_formula:
            self.refresh(u"Formula-driven parameters can be paired again. "
                         u"Auto-match to pick them up.")
            return
        driven_src = set(p.name for p in self.src_params if p.formula)
        driven_tgt = set(p.name for p in self.tgt_params if p.formula)
        dropped = [k for k, v in self.param_map.items()
                   if k in driven_src or v in driven_tgt]
        for name in dropped:
            del self.param_map[name]
        self.refresh(
            u"{} formula-driven pairings dropped.".format(len(dropped))
            if dropped else u"No formula-driven parameters were paired.")

    def scope_changed(self, sender, args):
        self.count_scope()
        self.refresh(u"Scope is {}.".format(
            u"the current Revit selection" if self.sel_only_cb.IsChecked
            else u"every placed instance"))

    def map_clicked(self, sender, args):
        source_name = self.selected_name(self.src_lb)
        target_name = self.selected_name(self.tgt_lb)
        if not source_name or not target_name:
            return

        source = next((p for p in self.src_params if p.name == source_name),
                      None)
        target = next((p for p in self.tgt_params if p.name == target_name),
                      None)
        if source is None or target is None:
            return
        if source.kind != target.kind:
            self.status_tb.Text = (
                u"{} is an {} parameter and {} is a {} one. A pairing has to "
                u"stay on one side of that line.".format(
                    source.name,
                    u"instance" if source.kind == "inst" else u"type",
                    target.name,
                    u"instance" if target.kind == "inst" else u"type"))
            return
        if not target.writable:
            self.status_tb.Text = u"{} is {} and cannot be written.".format(
                target.name, target.note)
            return
        if self.skip_formula and (source.formula or target.formula):
            driven = source if source.formula else target
            self.status_tb.Text = (
                u"{} is driven by a formula. Untick Skip formula-driven to "
                u"pair it anyway.".format(driven.name))
            return

        # A target takes from one source only, so whoever held it lets go.
        for held, taken in list(self.param_map.items()):
            if taken == target_name:
                del self.param_map[held]
        self.param_map[source_name] = target_name
        self.refresh(u"{}  →  {}".format(source_name, target_name))

    def skip_clicked(self, sender, args):
        source_name = self.selected_name(self.src_lb)
        if source_name and source_name in self.param_map:
            del self.param_map[source_name]
            self.refresh(u"{} is skipped. Its value is not carried "
                         u"across.".format(source_name))

    def auto_clicked(self, sender, args):
        self.auto_match()
        self.build_type_rows()
        self.count_scope()
        self.refresh(u"Auto-matched {} of {} parameters.".format(
            len(self.param_map), len(self.src_params)))

    def clear_clicked(self, sender, args):
        self.param_map = {}
        self.refresh(u"Every parameter is skipped.")

    def preview_clicked(self, sender, args):
        message(self.plan().report(), title=u"Param Map – preview")

    def swap_clicked(self, sender, args):
        """Swap element by element, stopping on each one that costs something.

        One at a time is the default rather than an escape hatch, because
        all or nothing is the wrong shape for this job: either nothing
        moves, or a hundred and fifty dimensions go at once and nobody
        sees which drawings changed. Element by element keeps everything
        Revit has no objection to and stops only where there is a
        decision to make, with the view open and the element on screen.

        Accept fixes is the way back to a single commit, for when taking
        the lot and redrawing afterwards is the faster trade.
        """
        plan = self.plan()
        if not plan.elements:
            message(u"Nothing to swap. " + plan.report(),
                    title=u"Param Map")
            return

        headline = u"{} elements on {} become {}.".format(
            len(plan.elements), plan.source.name, plan.target.name)

        if plan.accept_fixes:
            if not confirm(
                    headline +
                    u"\n\nAccept fixes is ticked, so this goes in one "
                    u"commit and Revit applies its own fixes as it meets "
                    u"them, which usually means deleting the dimension. "
                    u"Untick it to go element by element instead.\n\n"
                    u"Preview has the full breakdown.",
                    title=u"Param Map", yes=u"Swap"):
                return
            self.save_mapping()
            result = plan.run()
            self.families = loaded_families()
            self.fill_family_lists()
            self.count_scope()
            self.refresh(result)
            message(result, title=u"Param Map – done")
            return

        if not confirm(
                headline +
                u"\n\nEach one is tried on its own. The ones Revit has no "
                u"objection to are swapped and kept. Then you are walked "
                u"through the ones that would cost a dimension or a tag, one "
                u"at a time with the view open, so you can redraw as you "
                u"go.\n\nPreview has the full breakdown.",
                title=u"Param Map", yes=u"Swap"):
            return
        self.save_mapping()
        self.walk = plan
        self.Close()

    def close_clicked(self, sender, args):
        if self.source is not None and self.target is not None:
            self.save_mapping()
        self.Close()

    # --- saving ---
    def save_mapping(self):
        maps = read_json(MAPS_FILE, {})
        maps[self.map_key()] = {"params": self.param_map,
                                "types": self.type_map}
        write_json(MAPS_FILE, maps)

    # --- the plan ---
    def plan(self):
        return SwapPlan(
            source=self.source,
            target=self.target,
            src_params=self.src_params,
            tgt_params=self.tgt_params,
            param_map=dict(self.param_map),
            type_map=dict(self.type_map),
            src_refs=list(self.src_refs),
            tgt_refs=list(self.tgt_refs),
            selection=self.selection_ids(),
            accept_fixes=bool(self.accept_fixes_cb.IsChecked))


# ── WHAT REVIT DOES TO THE DRAWING ──────────────────────────────────────────

# READ THIS BEFORE CHANGING THE GUARD. It was got wrong once, in a way that
# looked thoroughly tested.
#
# Swapping 259 detail items between two versions of one family in the Nagel
# 2026 template raises roughly 180 failures: "One or more dimension references
# are or have become invalid", "The References of the highlighted Dimension are
# no longer parallel", "Constraints are not satisfied", and two dimension
# segments removed outright, taking their override text - "OR 375x375 PAD FOR
# POST" and "OR 450x450 PAD FOR POST" - with them. Revit resolves those
# warnings by deleting dimensions and dropping overrides.
#
# NONE OF IT IS VISIBLE BEFORE COMMIT. Measured on the live model with the
# swap applied and the document regenerated, inside the open transaction:
#
#     dimension reference counts     unchanged
#     dimension segment counts       unchanged
#     dimension override text        unchanged
#     Document.GetWarnings()         0
#
# A reference does not go away when it goes invalid, so counting references
# finds nothing. The text is not dropped until Revit resolves the failure,
# which happens during Commit. A before-and-after comparison inside the
# transaction therefore reports a clean swap and then commits the damage.
#
# The only thing that sees it is an IFailuresPreprocessor, which Revit calls
# during Commit with every posted failure in hand and before any of them is
# resolved. That is the gate. It also means the report is Revit's own wording
# rather than a guess assembled from a diff.
#
# The structural record below is still taken, because override text that
# Revit drops while the transaction is open can be written straight back, and
# because it gives the after-the-fact report something to compare against. It
# is not the guard. Do not make it the guard again.

GONE = "gone"                # the dimension is no longer in the model
WEAKENED = "weakened"        # it lost one or more references
RESEGMENTED = "resegmented"  # its segment count changed
RESTORED = "restored"        # its override text dropped and was put back
INTACT = "intact"

# The three that mean the drawing now says something it did not say before.
HARMFUL = (GONE, WEAKENED, RESEGMENTED)


CORRUPTION = u"corruption"
ERROR = u"error"
WARNING = u"warning"


class SwapFailures(IFailuresPreprocessor):
    """Catches what Revit is about to do to the drawing, before it does it.

    Revit calls this during Commit, possibly more than once, with every
    failure it has posted. Returning ProceedWithRollBack throws the whole
    transaction away - the type changes, the parameter writes and all - so
    nothing lands unless the drawing survives it.

    SEVERITY IS NOT THE LINE TO DRAW ON
        A first cut here refused Errors outright and let the user opt in to
        Warnings only. That made the opt-in useless, because the failures a
        family swap actually raises are Errors: all 188 of them on the Nagel
        template - "One or more dimension references are or have become
        invalid", "Constraints are not satisfied", "Can't form Angular
        Dimension". Revit is perfectly willing to commit through every one
        of them; its fix is simply to delete the dimension or drop the
        constraint, which is exactly the damage worth stopping for.

        So the line is whether the user has accepted Revit's own fixes, not
        how Revit graded the failure. Only DocumentCorruption is refused
        outright, because there is no version of that worth committing.

    Accepting means calling ResolveFailure on each message, which applies
    the resolution Revit would have applied itself, and committing. The
    captions are collected first, so the report can say what was accepted
    rather than leaving the user to find out from the drawing.
    """

    def __init__(self, accept_fixes=False):
        self.accept_fixes = accept_fixes
        self.captured = []      # [(severity, text, [ids], fix caption)]
        self._seen = set()
        self.rolled_back = False
        self.resolved = 0

    # --- the Revit callback ---
    def PreprocessFailures(self, accessor):
        corrupt = False
        resolvable = []
        # Revit is not obliged to call this only when something went wrong,
        # and a pass with nothing in it must not be read as a reason to throw
        # the transaction away - that would refuse every clean swap.
        posted = False

        for message in accessor.GetFailureMessages():
            posted = True
            try:
                severity = message.GetSeverity()
                text = message.GetDescriptionText() or u"(no description)"
                ids = sorted(eid(i) for i in message.GetFailingElementIds())
            except Exception:
                continue

            if severity == FailureSeverity.DocumentCorruption:
                grade = CORRUPTION
                corrupt = True
            elif severity == FailureSeverity.Error:
                grade = ERROR
            else:
                grade = WARNING

            fix = u""
            try:
                if message.HasResolutions():
                    fix = message.GetDefaultResolutionCaption() or u""
                    resolvable.append(message)
            except Exception:
                pass

            key = (text, tuple(ids))
            if key not in self._seen:
                self._seen.add(key)
                self.captured.append((grade, text, ids, fix))

        if not posted:
            return FailureProcessingResult.Continue

        if corrupt or not self.accept_fixes:
            self.rolled_back = True
            return FailureProcessingResult.ProceedWithRollBack

        for message in resolvable:
            try:
                accessor.ResolveFailure(message)
                self.resolved += 1
            except Exception:
                pass
        return FailureProcessingResult.ProceedWithCommit

    # --- reading back ---
    def of_grade(self, grade):
        return [c for c in self.captured if c[0] == grade]

    @property
    def unfixable(self):
        """Failures Revit offered no resolution for - nothing accepts those."""
        return [c for c in self.captured if not c[3]]

    def grouped(self):
        """[(text, count, elements, fix)] worst first."""
        counts, elements, fixes = {}, {}, {}
        for _grade, text, ids, fix in self.captured:
            counts[text] = counts.get(text, 0) + 1
            elements[text] = elements.get(text, 0) + len(ids)
            fixes.setdefault(text, fix)
        return sorted(((text, counts[text], elements[text], fixes[text])
                       for text in counts),
                      key=lambda row: (-row[1], row[0]))

    def summary(self, limit=8):
        """Revit's own wording, grouped, with the fix it would apply."""
        lines = []
        rows = self.grouped()
        for text, count, touched, fix in rows[:limit]:
            line = u"  {} x {}".format(count, text)
            if touched:
                line += u" ({} elements)".format(touched)
            if fix:
                line += u"\n      Revit's fix: {}".format(fix)
            lines.append(line)
        if len(rows) > limit:
            lines.append(u"  and {} more kinds".format(len(rows) - limit))
        return lines


class DimRecord(object):
    """One dimension as it stood before the swap, and what became of it.

    The text is held in memory for the length of the run rather than stamped
    into the model as an Extensible Storage entity the way _dimoverrides'
    own tools do. A swap should not leave a permanent record on 208
    dimensions as a side effect of carrying some parameters across.
    """

    def __init__(self, dim, view_name):
        self.id = eid(dim.Id)
        # The id Revit handed over, kept as it came. Rebuilding one from the
        # integer would mean picking an ElementId constructor, and that
        # signature changed between 2023 and 2024.
        self.element_id = dim.Id
        self.view = view_name
        self.references = _reference_count(dim)
        self.segments = read_current(dim)
        self.count = len(self.segments)
        self.summary = summarise(self.segments) if has_text(self.segments) else u""
        self.status = INTACT

    @property
    def had_text(self):
        return bool(self.summary)

    def describe(self):
        text = u" reading {}".format(self.summary) if self.summary else u""
        return u"{} in {}{}".format(self.id, self.view, text)


def _reference_count(dim):
    """How many references a dimension holds, or -1 if they cannot be read.

    A count, not a health check: a reference that has gone invalid is still
    counted, which is why this cannot stand in for the failure preprocessor.
    """
    try:
        return len(list(dim.References))
    except Exception:
        return -1


def dimension_index():
    """{element id: [Dimension]} for every dimension in the model.

    One pass over all dimensions builds the whole index, which is cheaper
    than asking per element and is the shape the caller needs anyway. Spot
    dimensions are left out: they subclass Dimension but hold their text in
    type parameters, which is the same reason _dimoverrides screens them.
    """
    index = {}
    for dim in FilteredElementCollector(doc).OfClass(Dimension):
        if not is_target(dim):
            continue
        try:
            references = list(dim.References)
        except Exception:
            continue
        for reference in references:
            try:
                target = reference.ElementId
            except Exception:
                continue
            if target is None or eid(target) <= 0:
                continue
            index.setdefault(eid(target), []).append(dim)
    return index


def record_dimensions(elements, index=None):
    """[DimRecord] for every dimension referencing any of the given elements.

    An index can be passed in when this is called per element: building one
    costs a pass over every dimension in the model, and the per-element pass
    would otherwise pay that 259 times over.
    """
    if not elements:
        return []
    if index is None:
        index = dimension_index()

    seen, records = set(), []
    view_names = {}
    for element in elements:
        for dim in index.get(eid(element.Id), []):
            # A shared index is built once and then goes stale as the walk
            # deletes dimensions. Reading the Id off a deleted element throws
            # rather than returning anything, so check before touching it.
            if not dim.IsValidObject:
                continue
            key = eid(dim.Id)
            if key in seen:
                continue
            seen.add(key)
            view_id = eid(dim.OwnerViewId)
            if view_id not in view_names:
                view = doc.GetElement(dim.OwnerViewId)
                view_names[view_id] = (element_name(view) if view is not None
                                       else u"(no view)")
            try:
                records.append(DimRecord(dim, view_names[view_id]))
            except Exception:
                continue
    return records


def check_dimensions(records, restore=True):
    """Re-read every recorded dimension and say what became of it.

    With restore on - inside the transaction - override text that has gone
    missing is written straight back. With it off, after the commit, this is
    a read-only account of what Revit's failure resolution actually did.

    Text is only ever restored when the segment count still matches.
    Restoring by index into a string that has been resegmented would write
    each value onto the wrong segment, which is worse than the blank it
    replaces.
    """
    for record in records:
        dim = doc.GetElement(record.element_id)
        if dim is None or not is_target(dim):
            record.status = GONE
            continue

        references = _reference_count(dim)
        if 0 <= references < record.references:
            record.status = WEAKENED
            continue

        try:
            current = read_current(dim)
        except Exception:
            record.status = WEAKENED
            continue

        if len(current) != record.count:
            record.status = RESEGMENTED
            continue

        if record.had_text and current != record.segments:
            if not restore:
                record.status = WEAKENED
                continue
            try:
                write_segments(dim, record.segments)
                record.status = RESTORED
            except Exception:
                record.status = WEAKENED
    return records


def tally(records):
    """{status: count} for a checked record list."""
    counts = {}
    for record in records:
        counts[record.status] = counts.get(record.status, 0) + 1
    return counts


def harmed(records):
    """The records whose drawing no longer says what it said before."""
    return [r for r in records if r.status in HARMFUL]


# ── THE SWAP ────────────────────────────────────────────────────────────────

class SwapPlan(object):
    """Everything the swap will touch, worked out before anything is written.

    Values are read off every instance BEFORE a single type is changed. Doing
    it instance by instance would work too, but reading the whole scope first
    means a failure part way through leaves a transaction that is rolled back
    whole rather than half a model carrying half its data.
    """

    def __init__(self, source, target, src_params, tgt_params,
                 param_map, type_map, selection, accept_fixes=False,
                 src_refs=(), tgt_refs=()):
        self.source = source
        self.target = target
        self.src_params = src_params
        self.tgt_params = tgt_params
        self.src_refs = list(src_refs)
        self.tgt_refs = list(tgt_refs)
        # Paired the same way the parameters are, and for the same reason:
        # the update renamed things and the name is all the two versions
        # have in common. A dimension holding a paired reference can be
        # drawn again on the new family. One holding an unpaired reference
        # cannot, and is the sort that has to be redrawn by hand.
        self.ref_map = auto_pairs(self.src_refs, self.tgt_refs)
        self.param_map = param_map
        self.type_map = type_map

        self.src_by_name = dict((name, s) for name, s in source.symbols)
        self.tgt_by_name = dict((name, s) for name, s in target.symbols)

        self.pairs = []             # [(source symbol, target symbol)]
        for src_name, tgt_name in sorted(type_map.items()):
            src_symbol = self.src_by_name.get(src_name)
            tgt_symbol = self.tgt_by_name.get(tgt_name)
            if src_symbol is not None and tgt_symbol is not None:
                self.pairs.append((src_symbol, tgt_symbol))

        self.elements = placed_on([p[0] for p in self.pairs], selection)
        self.unmapped_types = [name for name, _ in source.symbols
                               if name not in type_map]

        kinds = dict((p.name, p.kind) for p in src_params)
        self.inst_map = dict((k, v) for k, v in param_map.items()
                             if kinds.get(k) == "inst")
        self.type_param_map = dict((k, v) for k, v in param_map.items()
                                   if kinds.get(k) == "type")
        self.skipped = [p.name for p in src_params
                        if p.name not in param_map]

        # Read now, before anything is written, so the text is the text as
        # it stood. Preview builds a plan too and throws these away, which
        # is what lets it report the exposure without touching the model.
        self.accept_fixes = accept_fixes
        self.failures = None    # the preprocessor, once run() has run
        self.dim_records = record_dimensions(self.elements)

    # --- reporting ---
    def report(self):
        lines = []
        lines.append(u"{} elements on {} become {}.".format(
            len(self.elements), self.source.name, self.target.name))
        lines.append(u"")
        lines.append(u"{} instance parameters carried, {} skipped.".format(
            len(self.inst_map), len(self.skipped)))
        if self.type_param_map:
            lines.append(
                u"{} type parameters written onto the target types. That "
                u"lands on every instance of the target, including any "
                u"already placed.".format(len(self.type_param_map)))
        if self.unmapped_types:
            lines.append(
                u"{} source types are not paired and are left alone: "
                u"{}.".format(len(self.unmapped_types),
                              u", ".join(self.unmapped_types)))
        if self.skipped:
            shown = self.skipped[:8]
            lines.append(u"")
            lines.append(
                u"A skipped parameter comes out at the new family's default, "
                u"whatever it held before.")
            driven = len([p for p in self.src_params
                          if p.formula and p.name in self.skipped])
            if driven:
                lines.append(
                    u"{} of them are driven by a formula, and the new family "
                    u"works those out for itself.".format(driven))
            lines.append(u"Skipped: {}{}".format(
                u", ".join(shown),
                u" and {} more".format(len(self.skipped) - len(shown))
                if len(self.skipped) > len(shown) else u""))
        lines.append(u"")
        lines.append(self.reference_note())
        lines.append(u"")
        lines.append(self.dimension_note())
        lines.append(u"")
        lines.append(u"One transaction, so one undo puts it back.")
        return u"\n".join(lines)

    def reference_note(self):
        """What a dimension could be reattached to on the new family.

        This is the question behind every broken dimension: the update
        renamed or dropped a reference plane something was dimensioned to.
        Where the name still pairs, the dimension can be drawn again. Where
        it does not, nothing can put it back without inventing an answer.
        """
        if not self.src_refs and not self.tgt_refs:
            return (u"Neither family reports a named reference. Dimensions "
                    u"on these are held on plain geometry, which cannot be "
                    u"paired up by name.")
        if not self.tgt_refs:
            return (u"{} named references on {}, none on {}. Nothing a "
                    u"broken dimension could be moved onto.".format(
                        len(self.src_refs), self.source.name,
                        self.target.name))

        unmatched = [r.name for r in self.src_refs
                     if r.name not in self.ref_map]
        line = (u"{} named references on {}, {} on {}. {} pair by "
                u"name.".format(len(self.src_refs), self.source.name,
                                len(self.tgt_refs), self.target.name,
                                len(self.ref_map)))
        if unmatched:
            shown = unmatched[:6]
            line += u" {} do not: {}{}.".format(
                len(unmatched), u", ".join(shown),
                u" and {} more".format(len(unmatched) - len(shown))
                if len(unmatched) > len(shown) else u"")
        return line

    def dimension_note(self):
        """What the swap is exposing on the dimension side."""
        if not self.dim_records:
            base = u"No dimensions reference these elements."
        else:
            texts = len([r for r in self.dim_records if r.had_text])
            base = (u"{} dimensions reference these elements, {} of them "
                    u"carrying override text.").format(
                        len(self.dim_records), texts)
        if self.accept_fixes:
            return base + (u" Revit's own fixes are being accepted, so "
                           u"the swap will stand even where the fix is to "
                           u"delete a dimension.")
        return base + (u" The swap is rolled back whole if Revit raises a "
                       u"single failure committing it, which is the only "
                       u"point at which dimension damage shows.")

    # --- reading the model ---
    def cache_values(self):
        """Every mapped value, read off the model before anything is written.

        Returns ([(element, target symbol, {name: (kind, value)})],
        {source symbol id: {name: (kind, value)}}). Both the whole-model run
        and the per-element pass read the same cache, so the two cannot drift
        apart on what they think they are carrying.
        """
        target_of = {}
        for src_symbol, tgt_symbol in self.pairs:
            target_of[eid(src_symbol.Id)] = tgt_symbol

        cached = []
        for element in self.elements:
            tgt_symbol = target_of.get(eid(element.GetTypeId()))
            if tgt_symbol is None:
                continue
            values = {}
            for src_name, tgt_name in self.inst_map.items():
                got = read_value(element.LookupParameter(src_name))
                if got is not None:
                    values[tgt_name] = got
            cached.append((element, tgt_symbol, values))

        type_values = {}
        for src_symbol, tgt_symbol in self.pairs:
            values = {}
            for src_name, tgt_name in self.type_param_map.items():
                got = read_value(src_symbol.LookupParameter(src_name))
                if got is not None:
                    values[tgt_name] = got
            type_values[eid(src_symbol.Id)] = values
        return cached, type_values

    # --- doing it ---
    def run(self):
        """Swap, carry the values, and report what Revit let stand.

        The gate is the failure preprocessor on the commit, not anything
        measured here: the damage a family swap does to dimensions is
        invisible until Revit resolves its own failures, which happens
        inside Commit(). See the note above SwapFailures.
        """
        cached, type_values = self.cache_values()

        swapped = failed_swap = 0
        wrote = failed_write = 0
        problems = []

        self.failures = SwapFailures(self.accept_fixes)
        transaction = Transaction(doc, "Param Map: {} to {}".format(
            self.source.name, self.target.name))
        options = transaction.GetFailureHandlingOptions()
        options.SetFailuresPreprocessor(self.failures)
        options.SetClearAfterRollback(True)
        transaction.SetFailureHandlingOptions(options)
        transaction.Start()
        try:
            for _src_symbol, tgt_symbol in self.pairs:
                if not tgt_symbol.IsActive:
                    tgt_symbol.Activate()
            doc.Regenerate()

            done = []
            for element, tgt_symbol, values in cached:
                try:
                    element.ChangeTypeId(tgt_symbol.Id)
                    swapped += 1
                    done.append((element, values))
                except Exception as ex:
                    failed_swap += 1
                    if len(problems) < 4:
                        problems.append(u"{}: {}".format(eid(element.Id), ex))
            doc.Regenerate()

            for element, values in done:
                for name, (kind, value) in values.items():
                    if write_value(element.LookupParameter(name), kind, value):
                        wrote += 1
                    else:
                        failed_write += 1

            for src_symbol, tgt_symbol in self.pairs:
                for name, (kind, value) in \
                        type_values.get(eid(src_symbol.Id), {}).items():
                    if write_value(tgt_symbol.LookupParameter(name),
                                   kind, value):
                        wrote += 1
                    else:
                        failed_write += 1

            doc.Regenerate()

            # Any override text that has already gone is put back while the
            # transaction is still open. What Revit drops when it resolves
            # its failures happens after this, during the commit below, and
            # is reported rather than repaired.
            check_dimensions(self.dim_records, restore=True)
            doc.Regenerate()

            status = transaction.Commit()
        except Exception as ex:
            try:
                if transaction.HasStarted() and not transaction.HasEnded():
                    transaction.RollBack()
            except Exception:
                pass
            return (u"Nothing was changed. The swap was rolled back "
                    u"whole: {}".format(ex))

        if status != TransactionStatus.Committed:
            return self._refused()

        # Committed. Re-read without writing, so the report says what Revit
        # actually left behind rather than what was intended.
        check_dimensions(self.dim_records, restore=False)
        return self._verdict(swapped, failed_swap, wrote, failed_write,
                             problems)

    # --- doing it one at a time ---
    def write_types(self, type_values):
        """Write the mapped type parameters onto the target types.

        A type parameter belongs to the type, not to any one instance, so it
        is written once and is no part of what the per-element pass judges.
        Failing here is reported but does not stop the walk: the instances
        are still worth carrying.
        """
        if not self.type_param_map:
            return True
        guard = SwapFailures(True)
        transaction = Transaction(doc, "Param Map: type parameters")
        options = transaction.GetFailureHandlingOptions()
        options.SetFailuresPreprocessor(guard)
        options.SetClearAfterRollback(True)
        transaction.SetFailureHandlingOptions(options)
        transaction.Start()
        try:
            for src_symbol, tgt_symbol in self.pairs:
                if not tgt_symbol.IsActive:
                    tgt_symbol.Activate()
                for name, pair in type_values.get(eid(src_symbol.Id),
                                                  {}).items():
                    kind, value = pair
                    write_value(tgt_symbol.LookupParameter(name), kind, value)
            return transaction.Commit() == TransactionStatus.Committed
        except Exception:
            try:
                if transaction.HasStarted() and not transaction.HasEnded():
                    transaction.RollBack()
            except Exception:
                pass
            return False

    def triage(self):
        """Swap what Revit has no objection to, and queue the rest.

        One transaction per element, because a failure names the dimension
        that broke rather than the instance that broke it - see the note
        above the ONE AT A TIME section. A clean commit is left standing; a
        refused one is rolled back whole by the preprocessor, so that element
        is still on the old family with its parameters untouched, and it goes
        on the queue with the failures Revit raised against it.
        """
        cached, type_values = self.cache_values()
        dim_map = dimension_index()
        tag_map = tag_index()
        prepared = self.write_types(type_values)

        jobs = []
        clean = written = 0
        cancelled = False
        total = len(cached)
        with forms.ProgressBar(title="Trying one at a time: "
                                     "{value} of {max_value}",
                               cancellable=True) as bar:
            for index, row in enumerate(cached):
                if bar.cancelled:
                    cancelled = True
                    break
                bar.update_progress(index + 1, total)
                element, tgt_symbol, values = row
                dims = record_dimensions([element], dim_map)
                ok, guard, wrote = swap_one(
                    element.Id, tgt_symbol.Id, values, False, records=dims,
                    label=u"element {}".format(eid(element.Id)))
                if ok:
                    clean += 1
                    written += wrote
                else:
                    jobs.append(ElementJob(element, tgt_symbol, values, guard,
                                           dims, record_tags([element],
                                                             tag_map)))
        return TriageResult(clean, written, jobs, cancelled, prepared)

    def _verdict(self, swapped, failed_swap, wrote, failed_write, problems):
        lines = [u"{} of {} elements are now {}.".format(
            swapped, len(self.elements), self.target.name)]
        lines.append(u"{} parameter values written.".format(wrote))
        if failed_swap:
            lines.append(
                u"{} elements would not change type. An element inside a "
                u"group or a pinned one has to be dealt with by hand.".format(
                    failed_swap))
        if failed_write:
            lines.append(
                u"{} values would not write. A formula-driven or read-only "
                u"target parameter is the usual reason.".format(failed_write))

        captured = self.failures.captured if self.failures else []
        if captured:
            lines.append(u"")
            lines.append(
                u"Revit raised {} failures committing this, and its own fixes "
                u"were applied because Accept Revit's fixes is ticked:".format(
                    len(captured)))
            lines.extend(self.failures.summary())

        counts = tally(self.dim_records)
        if self.dim_records:
            lines.append(u"")
            lines.append(u"{} dimensions were watched. {} came through "
                         u"unchanged.".format(len(self.dim_records),
                                              counts.get(INTACT, 0)))
            if counts.get(RESTORED):
                lines.append(
                    u"{} lost their override text while the transaction was "
                    u"open and had it put back.".format(counts[RESTORED]))
            spoiled = sum(counts.get(k, 0) for k in HARMFUL)
            if spoiled:
                lines.append(u"{} did not survive the commit:".format(spoiled))
                for label, wording in ((GONE, u"gone"),
                                       (WEAKENED, u"changed"),
                                       (RESEGMENTED, u"resegmented")):
                    if counts.get(label):
                        lines.append(u"  {} {}".format(counts[label], wording))
                shown = harmed(self.dim_records)[:8]
                lines.append(u"  " + u"; ".join(r.describe() for r in shown))
                lines.append(
                    u"Those have to be redrawn by hand. Undo puts the whole "
                    u"swap back if that is not what you wanted.")
        if problems:
            lines.append(u"")
            lines.append(u"First few: " + u"; ".join(problems))
        return u"\n".join(lines)

    def _refused(self):
        """Why nothing was changed, in Revit's own words.

        Revit's failure preprocessor threw the transaction away during the
        commit, so the type changes and the parameter writes went with it.
        The messages are reported as they came, with the fix Revit would
        have applied to each, because that fix IS the damage: which new
        reference should stand in for one the update removed is not a
        question this can answer without inventing an answer.
        """
        captured = self.failures.captured if self.failures else []
        lines = [u"Nothing was changed.", u""]
        if not captured:
            lines.append(
                u"Revit refused the commit without saying why. Nothing was "
                u"written, so the model is as it was.")
            return u"\n".join(lines)

        lines.append(
            u"Revit raised {} failures during the commit, so the whole swap "
            u"was rolled back:".format(len(captured)))
        lines.extend(self.failures.summary())
        lines.append(u"")

        corruption = self.failures.of_grade(CORRUPTION)
        unfixable = self.failures.unfixable
        if corruption:
            lines.append(
                u"{} of those report document corruption. That is refused "
                u"whatever the settings say.".format(len(corruption)))
        else:
            lines.append(
                u"Revit would commit through every one of these. Its fixes "
                u"are listed above, and they are the damage: removing the "
                u"reference, deleting the dimension, dropping the "
                u"constraint.")
            if unfixable:
                # Not a reason the swap cannot be accepted - the other
                # failures still resolve and the commit still goes through.
                # Saying "nothing to accept" here read as the checkbox being
                # useless, which it is not.
                lines.append(
                    u"{} of them Revit offered no fix for. Those are not a "
                    u"blocker: accepting takes the other {} and Revit does "
                    u"whatever it does with the rest.".format(
                        len(unfixable), len(captured) - len(unfixable)))

        lines.append(u"")
        lines.append(
            u"It happens when the updated family renamed or removed a "
            u"reference plane something was dimensioned or constrained to. "
            u"Fix the references in the family and the swap goes through "
            u"clean, which is the outcome worth having.\n\n"
            u"Short of that: leave Accept fixes unticked and Swap goes "
            u"element by element, keeping everything Revit has no "
            u"objection to and stopping on the rest with their view open "
            u"so you redraw as you go.")
        if self.dim_records:
            lines.append(u"")
            lines.append(
                u"{} dimensions reference these elements, {} of them "
                u"carrying override text.".format(
                    len(self.dim_records),
                    len([r for r in self.dim_records if r.had_text])))
        return u"\n".join(lines)


# ── ONE AT A TIME ───────────────────────────────────────────────────────────

# WHY THIS EXISTS
#     The whole-model gate is all or nothing. Swapping the Longitudinal LB
#     family in the Nagel template raises about 188 failures, so either
#     nothing moves at all, or Revit deletes 159 dimensions in one commit and
#     the user finds out afterwards which drawings changed. Neither is the
#     outcome anybody wants.
#
#     Most elements are not the problem. Tried one at a time, every element
#     nothing is dimensioned to commits clean and silently, and only the ones
#     Revit actually objects to are left. Those are worth stopping on: open
#     the view, say what is about to go, and let the user take it and redraw
#     the dimension while standing in the right view with the element in
#     front of them.
#
# WHY ONE TRANSACTION PER ELEMENT
#     A failure names the DIMENSION that broke, not the instance that broke
#     it. A bulk rollback therefore says what went wrong but not who did it,
#     and dimensions are shared - one can reference two instances. One
#     element per commit is the only way to get an unambiguous answer, and it
#     costs a transaction each, which is why it only runs after the bulk
#     attempt has already been refused.
#
#     Judging each element in the state the model is actually in at that
#     moment makes the result order dependent: a dimension spanning two
#     instances objects to whichever is swapped first. That is honest rather
#     than tidy - the second one really is clean once the first has gone
#     through - and it is why the queue is walked in id order and never
#     re-triaged part way.
#
# WHY THE WINDOW IS MODELESS
#     The user has to be able to draw dimensions between steps, and a modal
#     dialog blocks Revit. So this window is shown with Show() rather than
#     ShowDialog(), every Revit call from it goes through an ExternalEvent,
#     and the script declares __persistentengine__ so the engine is still
#     alive when a button is clicked minutes later.

PENDING = "pending"
SWAPPED = "swapped"
LEFT = "left"
BLOCKED = "blocked"      # Revit refused it even with its own fixes accepted


def tag_text(tag):
    """A tag's text, flattened, or "" if it will not give one up."""
    try:
        return u" ".join(unicode(tag.TagText).split())
    except Exception:
        return u""


def tag_index():
    """{element id: [IndependentTag]} for every tag in the model.

    GetTaggedLocalElementIds is the only accessor that spans 2022-2026.
    TaggedLocalElementId and GetTaggedLocalElement were both removed in
    2023, so neither can be used here.
    """
    index = {}
    for tag in FilteredElementCollector(doc).OfClass(IndependentTag):
        try:
            targets = list(tag.GetTaggedLocalElementIds())
        except Exception:
            continue
        for target in targets:
            if target is None or eid(target) <= 0:
                continue
            index.setdefault(eid(target), []).append(tag)
    return index


class TagRecord(object):
    """One tag on a swapped element, and whether it still reads the same.

    A tag survives a type change - the instance id does not move, so the tag
    keeps its host - but its label can stop resolving: a label pointing at a
    parameter the new family does not declare comes back empty. That is a
    silent change on the drawing, which is the kind most worth reporting.
    """

    def __init__(self, tag, view_name):
        self.id = eid(tag.Id)
        self.element_id = tag.Id
        self.view = view_name
        self.text = tag_text(tag)
        self.status = INTACT

    def describe(self):
        reading = u" reading {}".format(self.text) if self.text else u" blank"
        return u"{} in {}{}".format(self.id, self.view, reading)


def record_tags(elements, index=None):
    """[TagRecord] for every tag hosted on any of the given elements."""
    if not elements:
        return []
    if index is None:
        index = tag_index()

    seen, records, view_names = set(), [], {}
    for element in elements:
        for tag in index.get(eid(element.Id), []):
            key = eid(tag.Id)
            if key in seen:
                continue
            seen.add(key)
            view_id = eid(tag.OwnerViewId)
            if view_id not in view_names:
                view = doc.GetElement(tag.OwnerViewId)
                view_names[view_id] = (element_name(view) if view is not None
                                       else u"(no view)")
            try:
                records.append(TagRecord(tag, view_names[view_id]))
            except Exception:
                continue
    return records


def check_tags(records):
    """Re-read every recorded tag and say whether it still reads the same.

    Read only. A tag whose label stopped resolving is not something this can
    put back - the label points at a parameter that is gone - so it is
    reported for the user to re-tag by hand.
    """
    for record in records:
        tag = doc.GetElement(record.element_id)
        if tag is None:
            record.status = GONE
        elif tag_text(tag) != record.text:
            record.status = WEAKENED
        else:
            record.status = INTACT
    return records


def swap_one(element_id, symbol_id, values, accept_fixes, records=None,
             label=u"one element"):
    """Change one element's type and write its values, in its own transaction.

    Returns (committed, guard, written). A refused commit leaves nothing
    behind: the preprocessor rolls the whole thing back, so the element is
    still on the old family with its parameters untouched.
    """
    element = doc.GetElement(element_id)
    symbol = doc.GetElement(symbol_id)
    if element is None or symbol is None:
        return (False, None, 0)

    guard = SwapFailures(accept_fixes)
    transaction = Transaction(doc, u"Param Map: {}".format(label))
    options = transaction.GetFailureHandlingOptions()
    options.SetFailuresPreprocessor(guard)
    options.SetClearAfterRollback(True)
    transaction.SetFailureHandlingOptions(options)
    transaction.Start()

    written = 0
    try:
        if not symbol.IsActive:
            symbol.Activate()
            doc.Regenerate()
        element.ChangeTypeId(symbol.Id)
        doc.Regenerate()
        for name, pair in values.items():
            kind, value = pair
            if write_value(element.LookupParameter(name), kind, value):
                written += 1
        doc.Regenerate()
        # Override text Revit blanked while the transaction was open can be
        # written straight back. What it drops resolving its own failures
        # happens during the commit below, and is reported, not repaired.
        if records:
            check_dimensions(records, restore=True)
            doc.Regenerate()
        status = transaction.Commit()
    except Exception:
        try:
            if transaction.HasStarted() and not transaction.HasEnded():
                transaction.RollBack()
        except Exception:
            pass
        return (False, guard, 0)

    return (status == TransactionStatus.Committed, guard, written)


class ElementJob(object):
    """One element the bulk swap could not take, queued for the walk."""

    def __init__(self, element, symbol, values, guard, dims, tags):
        self.element_id = element.Id
        self.id = eid(element.Id)
        self.symbol_id = symbol.Id
        self.symbol_name = element_name(symbol)
        self.values = values
        self.guard = guard
        self.dims = dims
        self.tags = tags
        self.state = PENDING
        self.written = 0
        self.note = u""

    def view_id(self):
        """The view to open: the one holding what is about to be lost.

        A dimension that is going to be deleted is the whole reason to stop
        here, so its view is where the user has to stand to redraw it.
        Falling back to a tag's view, and then to the element's own if it is
        view specific, covers the cases where no dimension is involved.
        """
        for record in self.dims:
            dim = doc.GetElement(record.element_id)
            if dim is not None:
                return dim.OwnerViewId
        for record in self.tags:
            tag = doc.GetElement(record.element_id)
            if tag is not None:
                return tag.OwnerViewId
        element = doc.GetElement(self.element_id)
        if element is not None and element.ViewSpecific:
            return element.OwnerViewId
        return None

    def cost_lines(self):
        """What swapping this one costs: Revit's words, then the drawing's.

        Anything already gone from the model is left out. The queue is built
        before the walk starts, and a dimension listed against element 40 may
        have been deleted swapping element 12.
        """
        lines = []
        if self.guard is not None:
            for text, count, touched, fix in self.guard.grouped():
                lines.append((u"fail", u"{} x {}".format(count, text)))
                if fix:
                    lines.append((u"fix", u"Revit's fix: {}".format(fix)))
        for record in self.dims:
            if doc.GetElement(record.element_id) is None:
                continue
            lines.append((u"dim", u"Dimension {}".format(record.describe())))
        for record in self.tags:
            if doc.GetElement(record.element_id) is None:
                continue
            lines.append((u"tag", u"Tag {}".format(record.describe())))
        if not lines:
            lines.append((u"fix", u"Nothing left to lose here. Worth taking."))
        return lines


class TriageResult(object):
    """What the per-element pass found."""

    def __init__(self, clean, written, jobs, cancelled, prepared):
        self.clean = clean          # committed on their own, nothing lost
        self.written = written      # parameter values carried on those
        self.jobs = jobs            # [ElementJob] Revit objected to
        self.cancelled = cancelled
        self.prepared = prepared    # type parameters written, or not

    def report(self):
        lines = [u"{} elements went through clean, carrying {} "
                 u"values.".format(self.clean, self.written)]
        if not self.prepared:
            lines.append(u"The type parameters could not be written.")
        if self.cancelled:
            lines.append(u"Stopped early. Everything already swapped stays, "
                         u"and one undo puts each of them back.")
        if self.jobs:
            lines.append(u"")
            lines.append(u"{} would cost Revit something. Those are walked "
                         u"one at a time.".format(len(self.jobs)))
        elif not self.cancelled:
            lines.append(u"Nothing was left over.")
        return u"\n".join(lines)


class RevitCaller(IExternalEventHandler):
    """Runs queued work in a Revit API context on behalf of a modeless window.

    A button click on a modeless window arrives outside Revit's API context,
    where a transaction or a view change throws. Raising an ExternalEvent
    hands the work back to Revit to run when it is next idle - on the same
    thread the window lives on, so the callback can update the UI directly.

    The queue is this tool's own rather than pyrevit.revit.events' shared
    handler, which keeps a single function on a module global: a walk that
    sits open for an hour has no business sharing that with whatever else
    the user runs in the meantime.
    """

    def __init__(self):
        self.queue = []

    def Execute(self, uiapp):
        while self.queue:
            work = self.queue.pop(0)
            try:
                work()
            except Exception as ex:
                forms.alert(u"Param Map could not finish that step:\n\n"
                            u"{}".format(ex), title=TOOL_NAME)

    def GetName(self):
        return "Param Map one at a time"


class GuidedWindow(forms.WPFWindow):
    """Walks the queue one element at a time, modeless, so Revit stays usable.

    Each step: open the view holding what is at risk, frame the element, and
    list what swapping it costs. Swap this one applies it with Revit's own
    fixes accepted - that element only - and then waits, so the dimensions
    and tags can be redrawn before Next moves on.
    """

    def __init__(self, xaml_name, jobs, source, target, triage):
        # handle_esc off deliberately. This window sits open beside Revit
        # while the user draws, and Escape is how you back out of Revit's own
        # commands - one stray press with focus here would end the walk.
        forms.WPFWindow.__init__(self, xaml_name, handle_esc=False)
        apply_seed43_palette(self, SCRIPT_DIR)
        apply_seed43_dimensions(self, SCRIPT_DIR)

        self.jobs = jobs
        self.source = source
        self.target = target
        self.triage = triage
        self.at = 0
        self.busy = False

        # Resolved once, after the palette is applied, the same way the
        # mapping window does it.
        self._ink = {u"fail": brush("danger", "#E01B24"),
                     u"fix": brush("warning", "#E5A50A"),
                     u"dim": brush("text_primary", "#FFFFFF"),
                     u"tag": brush("text_muted", "#9CA3AF"),
                     u"ok": brush("success", "#27AE60")}

        self.caller = RevitCaller()
        self.event = ExternalEvent.Create(self.caller)

        self.park()
        # The first step runs straight through. __init__ is still inside the
        # script's own execution, which IS a valid Revit API context, so there
        # is no reason to wait for an idle moment that has not come yet - and
        # it puts the view on screen before the window does.
        self.enter(direct=True)

    # --- placement ---
    def park(self):
        """Sit against the right edge, out of the way of the drawing."""
        try:
            area = Windows.SystemParameters.WorkArea
            self.Left = area.Right - self.Width - 24
            self.Top = area.Top + 60
        except Exception:
            pass

    # --- handing work to Revit ---
    def ask(self, work):
        """Queue work for Revit's next idle moment and return immediately."""
        self.caller.queue.append(work)
        self.event.Raise()

    # --- the current job ---
    @property
    def job(self):
        if 0 <= self.at < len(self.jobs):
            return self.jobs[self.at]
        return None

    def enter(self, direct=False):
        """Show the job at self.at and send Revit to its view.

        direct is only true for the first step, called from __init__ while
        still in a valid API context. Every later step comes from a button on
        a modeless window and has to go through the ExternalEvent.
        """
        job = self.job
        if job is None:
            self.finish()
            return

        self.busy = False
        self.step_tb.Text = u"Element {} of {}".format(self.at + 1,
                                                       len(self.jobs))
        self.what_tb.Text = u"Id {} · {} becomes {}".format(
            job.id, job.symbol_name, self.target.name)
        self.view_tb.Text = u"Opening the view…"
        self.fill(job.cost_lines())
        # The pass that got here reported nothing of its own - this is where
        # its result is said, once, on the step that follows it.
        opening = u""
        if self.at == 0 and self.triage.clean:
            opening = (
                u"{} went through clean already, carrying {} values. "
                u"These {} need a decision. ".format(
                    self.triage.clean, self.triage.written,
                    len(self.jobs)))
        self.status_tb.Text = opening + (
            u"Swap this one lets Revit apply the fixes above, to this "
            u"element only. Leave it moves on and changes nothing.")
        self.arm(swap=True)
        if direct:
            self.go_to_view()
        else:
            self.ask(self.go_to_view)

    def fill(self, rows):
        """Rebuild the detail list from [(kind, text)] rows."""
        self.detail_lb.Items.Clear()
        for kind, text in rows:
            block = Windows.Controls.TextBlock()
            block.Text = text
            block.TextWrapping = Windows.TextWrapping.Wrap
            block.Foreground = self._ink.get(kind, self._ink[u"dim"])
            # Revit's fix is indented under the failure it belongs to.
            block.Margin = (Windows.Thickness(14, 0, 0, 4)
                            if kind == u"fix" else Windows.Thickness(0, 0, 0, 2))
            self.detail_lb.Items.Add(block)

    def arm(self, swap):
        """Before a swap it is Swap/Leave it; after one it is only Next."""
        self.swap_btn.IsEnabled = swap
        self.skip_btn.IsEnabled = swap
        self.next_btn.IsEnabled = not swap
        self.show_btn.IsEnabled = True

    # --- the Revit side ---
    def go_to_view(self):
        """Activate the view that holds what is at risk, and frame the element.

        ShowElements alone only searches views that are already open, so a
        dimension on a closed view comes back as Revit's own "no open view"
        box stacked behind this window, where it reads as a hang. Activating
        the owner view first turns that into a hit every time. Dim Restore
        does the same thing for the same reason.
        """
        job = self.job
        if job is None:
            return
        try:
            self._go_to_view(job)
        except Exception as ex:
            # Whatever went wrong, the one thing that must not happen is
            # the status sitting on "Opening the view" for good - that
            # reads as the tool having hung when it has in fact finished
            # and failed.
            self.view_tb.Text = u"Could not open the view: {}".format(ex)

    def _go_to_view(self, job):
        element = doc.GetElement(job.element_id)
        if element is None:
            self.view_tb.Text = u"That element is no longer in the model."
            return

        name = u""
        view_id = job.view_id()
        if view_id is not None:
            view = doc.GetElement(view_id)
            if view is not None:
                name = element_name(view)
                try:
                    if view.Id != uidoc.ActiveView.Id:
                        uidoc.ActiveView = view
                except Exception:
                    # A view template, or a context that will not take a view
                    # change. ShowElements may still find it among the open
                    # views, and says so itself if it cannot.
                    pass
        try:
            ids = List[ElementId]()
            ids.Add(element.Id)
            uidoc.Selection.SetElementIds(ids)
            uidoc.ShowElements(element.Id)
        except Exception as ex:
            self.view_tb.Text = u"Could not show {}: {}".format(job.id, ex)
            return
        self.view_tb.Text = (u"Showing in {}".format(name) if name
                             else u"Showing {}".format(job.id))

    def do_swap(self):
        """Swap this one element, accepting Revit's fixes, and say what went."""
        job = self.job
        if job is None:
            return
        ok, guard, written = swap_one(
            job.element_id, job.symbol_id, job.values, True,
            records=job.dims, label=u"element {}".format(job.id))
        job.written = written

        if not ok:
            job.state = BLOCKED
            job.note = u"Revit refused it even with its fixes accepted."
            self.fill([(u"fail", job.note)] +
                      [(u"fail", u"{} x {}".format(c, t))
                       for t, c, _n, _f in (guard.grouped() if guard else [])])
            self.status_tb.Text = (
                u"Nothing changed on this one. Next moves on and leaves it "
                u"on the old family.")
            self.arm(swap=False)
            return

        job.state = SWAPPED
        check_dimensions(job.dims, restore=False)
        check_tags(job.tags)
        self.fill(self.aftermath(job))
        self.status_tb.Text = (
            u"Swapped. Redraw the dimensions and tags listed above now, in "
            u"this view, then click Next.")
        self.arm(swap=False)

    def aftermath(self, job):
        """What the drawing lost on this one, for redrawing before Next."""
        rows = [(u"ok", u"Swapped to {}, {} values carried.".format(
            self.target.name, job.written))]
        lost = harmed(job.dims)
        if lost:
            rows.append((u"fail", u"{} dimensions to redraw:".format(
                len(lost))))
            for record in lost[:12]:
                rows.append((u"dim", record.describe()))
            if len(lost) > 12:
                rows.append((u"dim", u"and {} more".format(len(lost) - 12)))
        changed = [t for t in job.tags if t.status in HARMFUL]
        if changed:
            rows.append((u"fail", u"{} tags no longer read the same:".format(
                len(changed))))
            for record in changed[:12]:
                rows.append((u"tag", record.describe()))
        restored = [d for d in job.dims if d.status == RESTORED]
        if restored:
            rows.append((u"ok", u"{} kept their override text.".format(
                len(restored))))
        if not lost and not changed:
            rows.append((u"ok", u"Nothing was lost on the drawing."))
        return rows

    # --- buttons ---
    def show_clicked(self, sender, args):
        self.ask(self.go_to_view)

    def swap_clicked(self, sender, args):
        if self.busy:
            return
        self.busy = True
        self.status_tb.Text = u"Swapping…"
        self.arm(swap=False)
        self.next_btn.IsEnabled = False
        self.ask(self.after_swap)

    def after_swap(self):
        try:
            self.do_swap()
        finally:
            self.busy = False

    def skip_clicked(self, sender, args):
        job = self.job
        if job is not None:
            job.state = LEFT
        self.advance()

    def next_clicked(self, sender, args):
        self.advance()

    def advance(self):
        self.at += 1
        self.enter()

    def stop_clicked(self, sender, args):
        self.Close()
        message(self.wrap_up(u"Stopped."), title=u"Param Map – one at a time")

    def finish(self):
        self.Close()
        message(self.wrap_up(u"That was the last one."),
                title=u"Param Map – one at a time")

    def wrap_up(self, opening):
        swapped = len([j for j in self.jobs if j.state == SWAPPED])
        left = len([j for j in self.jobs if j.state == LEFT])
        blocked = len([j for j in self.jobs if j.state == BLOCKED])
        waiting = len([j for j in self.jobs if j.state == PENDING])

        lines = [opening, u""]
        lines.append(u"{} went through with the crowd.".format(
            self.triage.clean))
        lines.append(u"{} were taken one at a time.".format(swapped))
        if left:
            lines.append(u"{} were left on {}.".format(left, self.source.name))
        if blocked:
            lines.append(u"{} Revit would not take at all.".format(blocked))
        if waiting:
            lines.append(u"{} were never reached.".format(waiting))

        lost = sum(len(harmed(j.dims)) for j in self.jobs
                   if j.state == SWAPPED)
        if lost:
            lines.append(u"")
            lines.append(u"{} dimensions went in the process. Anything not "
                         u"redrawn as you went is still missing.".format(lost))
        lines.append(u"")
        lines.append(u"Each element was its own transaction, so undo takes "
                     u"them back one at a time rather than all at once.")
        return u"\n".join(lines)


# ── ENTRY POINT ─────────────────────────────────────────────────────────────

def show_guide_once():
    """Explain the tool on first run, then stay quiet about it."""
    if read_json(GUIDE_FILE, {}).get("shown"):
        return
    message(GUIDE, title=u"Param Map")
    write_json(GUIDE_FILE, {"shown": True})


def walk(plan):
    """Run the per-element pass, then walk whatever it could not take.

    Called after the mapping window has closed, never before. The guided
    window is modeless so Revit stays usable between steps, and a modeless
    window opened behind a modal one cannot be reached.

    There is deliberately no dialog between the pass and the walk. What one
    would have said - how many went through, how many are left - is the first
    thing the walk's own window says, and a popup in front of it is a click
    for something already on screen.

    Nor is there a check that the engine will outlive the script. It needs to
    be persistent for the walk's buttons to work, but pyRevit exposes nothing
    that answers that question: __cachedengine__ reports rocket mode, and
    EXEC_PARAMS.cached_engine reads that name from the wrong scope and is
    therefore always False. A check that cannot be right is worse than none,
    because it stops the walk on a session where it would have worked.
    """
    outcome = plan.triage()
    if not outcome.jobs:
        message(outcome.report(), title=TITLE_WALK)
        return
    GuidedWindow("Guided.xaml", outcome.jobs, plan.source, plan.target,
                 outcome).show(modal=False)


def main():
    if doc.IsFamilyDocument:
        message(u"Param Map works on a project. Open the model holding the "
                u"placed instances and run it there.", title=u"Param Map")
        return
    show_guide_once()
    window = ParamMapWindow("ParamMap.xaml")
    window.ShowDialog()
    if window.walk is not None:
        walk(window.walk)


main()
