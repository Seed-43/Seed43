# -*- coding: utf-8 -*-
# "Sheet Scale - title block setup"
# "Seed43"
# """
# Gives every title block family placed on a sheet a Scale_Override label
# and an Override Scale switch:
#
#   - the Scale_Override label is copied from Sheet Scale.rfa (next to this
#     file) onto the exact spot of the family's scale label, in the scale
#     label's text type (the API cannot change which parameter a label
#     shows, so it is copied from a family that already has it)
#   - Override Scale (Yes/No, instance, Other) drives the new label's
#     visibility; Bool 01 (shared, from Scale_Override.txt) is
#     not(Override Scale) and drives the scale label's, so one shows
#   - the family is loaded back into the project and closed unsaved
#
# A family that already has Override Scale is skipped, and so is one with
# no scale label. Scale_Override is added to the project as a Sheets
# instance parameter when it is missing.
# """

# ── IMPORTS ────────────────────────────────────────────────────────────────

import math
import os.path as op
import re

from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    BuiltInCategory, BuiltInParameter, Category, CopyPasteOptions,
    DuplicateTypeAction, ElementId, ElementTransformUtils,
    FilteredElementCollector, FamilyInstance, ForgeTypeId,
    IDuplicateTypeNamesHandler, IFamilyLoadOptions, Line, SpecTypeId,
    TextElement, TextNote, Transaction, Transform, XYZ
)

# ── CONSTANTS ──────────────────────────────────────────────────────────────

HERE          = op.dirname(op.abspath(__file__))
DONOR_RFA     = op.join(HERE, "Sheet Scale.rfa")
SHARED_FILE   = op.join(HERE, "Scale_Override.txt")

OVERRIDE_TEXT = "Scale_Override"    # shared text parameter, on the sheet
SHARED_BOOL   = "Bool 01"           # shared Yes/No: shows the scale label
TOGGLE        = "Override Scale"    # family Yes/No: shows Scale_Override
DONOR_VERSION = 2026                # Sheet Scale.rfa is saved in 2026

# NOTE: the API cannot say which parameter a label shows, so the scale label
# is found by its sample value, as the family editor draws it.
SCALE_TEXT_RE = re.compile(
    r"^\s*(1\s*:\s*\d+|as indicated|scale|n\.?\s*t\.?\s*s\.?)\s*$", re.IGNORECASE)
# ...or, failing that, as the label beside a caption reading SCALE
SCALE_NOTE_RE = re.compile(r"^\s*:?\s*scale\s*:?\s*$", re.IGNORECASE)
SCALE_NOTE_REACH = 80.0 / 304.8     # gap to the label, 80 mm in feet
SCALE_ROW_TOL = 4.0 / 304.8         # vertical centres within 4 mm: one row


# ── HELPERS ────────────────────────────────────────────────────────────────

class _UseDestinationTypes(IDuplicateTypeNamesHandler):
    def OnDuplicateTypeNamesFound(self, args):
        return DuplicateTypeAction.UseDestinationTypes


class _LoadOptions(IFamilyLoadOptions):
    """Reload over the project's copy, keeping its parameter values."""

    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source,
                            overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True


def _other_group():
    """The parameter group Other (an empty ForgeTypeId)."""
    return ForgeTypeId()


def labels(fdoc):
    """The family's labels: text elements that are not text notes."""
    return [e for e in FilteredElementCollector(fdoc)
            .WhereElementIsNotElementType().ToElements()
            if isinstance(e, TextElement) and not isinstance(e, TextNote)]


def describe_text(fdoc, limit=12):
    """What the family holds that could be a label, for the report when no
    scale label is found: each element's class, category and text."""
    seen = []
    for e in (FilteredElementCollector(fdoc)
              .WhereElementIsNotElementType().ToElements()):
        cat = ""
        try:
            cat = e.Category.Name if e.Category is not None else ""
        except Exception:
            pass
        text = getattr(e, "Text", None) if isinstance(e, TextElement) else None
        textish = "label" in cat.lower() or "text" in cat.lower()
        if text is None and not textish:
            continue
        seen.append(u"{} / {} / '{}'".format(
            type(e).__name__, cat, (text or u"").strip()[:30]))
    if not seen:
        return u"no text or labels found"
    more = u" +{} more".format(len(seen) - limit) if len(seen) > limit else u""
    return u"; ".join(seen[:limit]) + more


def shared_definitions(app):
    """{name: ExternalDefinition} from Scale_Override.txt. The app's shared
    parameter file is switched to it; restore_shared() puts it back."""
    app.SharedParametersFilename = SHARED_FILE
    f = app.OpenSharedParameterFile()
    if f is None:
        raise RuntimeError("could not read " + SHARED_FILE)
    defs = {}
    for g in f.Groups:
        for d in g.Definitions:
            defs[d.Name] = d
    for name in (OVERRIDE_TEXT, SHARED_BOOL):
        if name not in defs:
            raise RuntimeError("{} is not in Scale_Override.txt".format(name))
    return defs


def scale_label(fdoc):
    """(the scale label, None) or (None, why)."""
    fm = fdoc.FamilyManager
    found = [lb for lb in labels(fdoc) if SCALE_TEXT_RE.match(lb.Text or "")]
    if len(found) > 1:
        # NOTE: older Seed43 borders lay an A1_Scale / A3_Scale label over
        # the scale label, switched by a family parameter - the real scale
        # label is the one whose visibility is not driven.
        free = []
        for lb in found:
            p = lb.get_Parameter(BuiltInParameter.IS_VISIBLE_PARAM)
            if p is None or fm.GetAssociatedFamilyParameter(p) is None:
                free.append(lb)
        if len(free) == 1:
            return free[0], None
        return None, "{} labels look like the scale".format(len(found))
    if not found:
        lb = label_beside_scale_note(fdoc)
        if lb is not None:
            return lb, None
        return None, "no scale label ({})".format(describe_text(fdoc))
    return found[0], None


def _box(fdoc, e):
    """Bounding box in the element's own view, or None."""
    try:
        return e.get_BoundingBox(fdoc.GetElement(e.OwnerViewId))
    except Exception:
        return None


def label_beside_scale_note(fdoc):
    """The label on the same row as a text note that says SCALE (':SCALE',
    'SCALE:'). Labels show a sample value in the family editor - on the
    Nagel borders the scale label's is '-' - so its caption finds it. Row
    first (vertical centres), then the gap between the two boxes: label
    boxes are wide, so plain centre distance picks the wrong row."""
    notes = [e for e in FilteredElementCollector(fdoc)
             .WhereElementIsNotElementType().ToElements()
             if isinstance(e, TextNote) and SCALE_NOTE_RE.match(e.Text or "")]
    best = None
    for note in notes:
        nb = _box(fdoc, note)
        if nb is None:
            continue
        ny = (nb.Min.Y + nb.Max.Y) / 2.0
        for lb in labels(fdoc):
            bb = _box(fdoc, lb)
            if bb is None:
                continue
            dy = abs((bb.Min.Y + bb.Max.Y) / 2.0 - ny)
            gap = max(0.0, nb.Min.X - bb.Max.X, bb.Min.X - nb.Max.X)
            if dy > SCALE_ROW_TOL or gap > SCALE_NOTE_REACH:
                continue
            score = dy * 10.0 + gap
            if best is None or score < best[0]:
                best = (score, lb)
    return best[1] if best else None


def place_like(fdoc, new, target):
    """Match the new label to the target's text type, alignment, width,
    rotation and position."""
    try:
        new.ChangeTypeId(target.GetTypeId())
    except Exception:
        pass
    for attr in ("HorizontalAlignment", "VerticalAlignment"):
        try:
            setattr(new, attr, getattr(target, attr))
        except Exception:
            pass
    try:
        new.Width = target.Width
    except Exception:
        pass
    a = new.BaseDirection
    b = target.BaseDirection
    angle = math.atan2(a.X * b.Y - a.Y * b.X, a.X * b.X + a.Y * b.Y)
    if abs(angle) > 1e-6:
        c = new.Coord
        ElementTransformUtils.RotateElement(
            fdoc, new.Id, Line.CreateBound(c, c + XYZ.BasisZ), angle)
    move = target.Coord - new.Coord
    if move.GetLength() > 1e-9:
        ElementTransformUtils.MoveElement(fdoc, new.Id, move)


# ── CORE LOGIC ─────────────────────────────────────────────────────────────

def upgrade_family(doc, family, donor_doc, donor_label, defs):
    """Add the Scale_Override label and switches to one title block family
    and load it back. Returns what happened, for the report."""
    fdoc = doc.EditFamily(family)
    try:
        fm = fdoc.FamilyManager
        if fm.get_Parameter(TOGGLE) is not None:
            return "skipped: already has {}".format(TOGGLE)
        scale, why = scale_label(fdoc)
        if scale is None:
            return "skipped: " + why

        t = Transaction(fdoc, "Scale_Override")
        t.Start()
        try:
            opts = CopyPasteOptions()
            opts.SetDuplicateTypeNamesHandler(_UseDestinationTypes())
            ids = ElementTransformUtils.CopyElements(
                donor_doc.GetElement(donor_label.OwnerViewId),
                List[ElementId]([donor_label.Id]),
                fdoc.GetElement(scale.OwnerViewId),
                Transform.Identity, opts)
            new = fdoc.GetElement(list(ids)[0])
            place_like(fdoc, new, scale)

            toggle = fm.AddParameter(TOGGLE, _other_group(),
                                     SpecTypeId.Boolean.YesNo, True)
            show_scale = fm.get_Parameter(SHARED_BOOL)
            if show_scale is None:
                show_scale = fm.AddParameter(defs[SHARED_BOOL],
                                             _other_group(), True)
            fm.SetFormula(show_scale, "not({})".format(TOGGLE))
            fm.AssociateElementParameterToFamilyParameter(
                scale.get_Parameter(BuiltInParameter.IS_VISIBLE_PARAM),
                show_scale)
            fm.AssociateElementParameterToFamilyParameter(
                new.get_Parameter(BuiltInParameter.IS_VISIBLE_PARAM), toggle)
            t.Commit()
        except Exception:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
            raise
        if fdoc.LoadFamily(doc, _LoadOptions()) is None:
            return "failed: Revit did not load the family back"
        return "updated"
    finally:
        fdoc.Close(False)       # never saved: only loaded into the project


def ensure_project_parameter(doc, defs):
    """Scale_Override as a Sheets instance parameter. Returns True when it
    was added."""
    it = doc.ParameterBindings.ForwardIterator()
    while it.MoveNext():
        if it.Key.Name == OVERRIDE_TEXT:
            return False
    app = doc.Application
    cats = app.Create.NewCategorySet()
    cats.Insert(Category.GetCategory(doc, BuiltInCategory.OST_Sheets))
    t = Transaction(doc, "Add Scale_Override")
    t.Start()
    try:
        doc.ParameterBindings.Insert(defs[OVERRIDE_TEXT],
                                     app.Create.NewInstanceBinding(cats))
        t.Commit()
    except Exception:
        t.RollBack()
        raise
    return True


def placed_title_block_families(doc):
    """{name: Family} of the title blocks placed on sheets."""
    fams = {}
    for tb in (FilteredElementCollector(doc)
               .OfCategory(BuiltInCategory.OST_TitleBlocks)
               .WhereElementIsNotElementType().ToElements()):
        if isinstance(tb, FamilyInstance):
            fam = tb.Symbol.Family
            fams[fam.Name] = fam
    return fams


def set_up_title_blocks(doc, progress=None):
    """Run the whole setup. Returns ([(family, result)], parameter added)."""
    app = doc.Application
    old_shared = app.SharedParametersFilename
    rows = []
    donor = None
    try:
        defs = shared_definitions(app)
        added = ensure_project_parameter(doc, defs)
        donor = app.OpenDocumentFile(DONOR_RFA)
        donor_labels = labels(donor)
        if not donor_labels:
            raise RuntimeError("Sheet Scale.rfa has no label to copy")
        # NOTE: by name, looked up again before each family: loading one
        # back replaces its Family element, so held references go invalid.
        families = sorted(placed_title_block_families(doc))
        for i, name in enumerate(families):
            try:
                fam = placed_title_block_families(doc).get(name)
                if fam is None:
                    result = "skipped: no longer in the project"
                else:
                    result = upgrade_family(doc, fam, donor, donor_labels[0],
                                            defs)
            except Exception as ex:
                result = "failed: {}".format(ex)
            rows.append((name, result))
            if progress is not None:
                try:
                    progress(i + 1, len(families))
                except Exception:
                    pass
        return rows, added
    finally:
        if donor is not None:
            donor.Close(False)
        try:
            app.SharedParametersFilename = old_shared or ""
        except Exception:
            pass
