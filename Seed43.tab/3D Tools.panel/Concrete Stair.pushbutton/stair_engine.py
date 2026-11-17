# -*- coding: utf-8 -*-
"""Concrete Stair geometry and Revit building, kept apart from the window.

All sizes in here are millimetres unless a name ends in _ft. The stair is a
native Revit Assembled Stair with a Monolithic Run, plus a short sketched
top landing of length `xyz`, sized so the soffit meets the underside of the
landing - the same as a cast stair poured against a 100 mm landing slab.
"""
import math

from System.Collections.Generic import List
from Autodesk.Revit import DB
from Autodesk.Revit.DB.Architecture import (Stairs, StairsRun, StairsLanding,
                                            StairsType, StairsRunType,
                                            StairsLandingType, StairsRunJustification)

__all__ = ["plan", "type_name", "ensure_type", "build", "remove_railings",
           "chain_loop", "landing_slab", "set_material", "stair_material"]

# ── CONSTANTS ───────────────────────────────────────────────────────────────

MM = 304.8
STAIR_FAMILY = u"Assembled Stair"
BASE_TYPE_NAMES = (u"Concrete Stair", u"Monolithic Stair", u"M_Monolithic Stair")


# ── HELPERS (private to this file) ──────────────────────────────────────────

def _ft(mm):
    return mm / MM


def _name(element):
    return DB.Element.Name.GetValue(element)


def _by_name(doc, cls, name):
    for el in DB.FilteredElementCollector(doc).OfClass(cls):
        if _name(el) == name:
            return el
    return None


class _Quiet(DB.IFailuresPreprocessor):
    """Let warnings through silently; errors still roll the stair back."""

    def PreprocessFailures(self, accessor):
        for msg in accessor.GetFailureMessages():
            if msg.GetSeverity() == DB.FailureSeverity.Warning:
                accessor.DeleteWarning(msg)
        return DB.FailureProcessingResult.Continue


# ── CORE LOGIC ──────────────────────────────────────────────────────────────

# --- Geometry ---
def plan(rise, riser_max, going, waist, landing, risers=None, top_landing=None):
    """Work out the stair from the sizes. Returns a dict (mm, degrees).

    rise       total height, base to top landing
    risers     number of risers (steps); None works it out from riser_max
    riser_max  largest riser allowed; risers are made equal and no taller
    going      tread depth
    waist      Structural Depth: square to the soffit, from the line through
               the inside corners of the steps
    landing    top landing thickness

    top_landing  the top landing length wanted (TL), or None for the least

    xyz is the top landing length: at least from the top riser face to where
    the soffit reaches the underside of the landing (xyz_min), longer when
    top_landing asks for more.
    """
    if rise <= 0:
        raise ValueError(u"The top must be above the base.")
    if min(riser_max, going, waist, landing) <= 0:
        raise ValueError(u"Riser, going, waist and landing must all be above 0.")
    if risers is None:
        risers = int(math.ceil(rise / riser_max - 1e-6))
    risers = int(risers)
    if risers < 1:
        raise ValueError(u"There must be at least one step.")
    riser = rise / risers
    slope = math.atan2(riser, going)
    treads = risers - 1                       # the top riser lands on the landing
    run_length = treads * going
    soffit_meets = (rise - landing + waist / math.cos(slope)) * going / riser
    exact = max(soffit_meets - run_length, 0.0)
    # Rounded UP to the next 5 mm. Not only for tidiness: with the landing
    # exactly that long the soffit lands on the landing's far bottom edge,
    # and Revit fails the whole stair with "Can't make stairs".
    xyz = math.ceil(exact / 5.0) * 5.0
    if xyz - exact < 0.5:
        xyz += 5.0
    xyz_min = xyz
    if top_landing is not None and top_landing > xyz_min:
        xyz = float(top_landing)
    return {
        "rise": rise, "risers": risers, "riser": riser, "treads": treads,
        "going": going, "waist": waist, "landing": landing,
        "run_length": run_length, "xyz": xyz, "xyz_exact": exact, "xyz_min": xyz_min,
        "angle": math.degrees(slope),
        "over_max": riser > riser_max + 1e-6,
        "soffit_foot": waist / math.cos(slope) * going / riser,
    }


# --- Types ---
def type_name(riser_max, going, waist, landing):
    return u"Concrete Stair - SR{0:g} SG{1:g} SD{2:g} LD{3:g}".format(
        round(riser_max, 1), round(going, 1), round(waist, 1), round(landing, 1))


def _base_stair_type(doc):
    """Any stair type will do as the starting point - its run and landing
    types are swapped for monolithic ones - but prefer a concrete one."""
    for name in BASE_TYPE_NAMES:
        t = _by_name(doc, StairsType, name)
        if t is not None:
            return t
    types = list(DB.FilteredElementCollector(doc).OfClass(StairsType))
    if not types:
        raise ValueError(u"This model has no stair types. Load one Assembled "
                         u"Stair type first.")
    return types[0]


def _monolithic(doc, cls, family, preferred_id):
    """A run or landing type of the monolithic family to copy: the base
    type's own when it is monolithic, else the first one in the model.
    A stair type named Concrete can still carry a non-monolithic run."""
    own = doc.GetElement(preferred_id) if preferred_id else None
    if own is not None and own.FamilyName == family:
        return own
    for t in DB.FilteredElementCollector(doc).OfClass(cls):
        if t.FamilyName == family:
            return t
    raise ValueError(u"This model has no {0} type to copy. Load an Assembled "
                     u"Stair with a {0} first.".format(family))


def ensure_type(doc, riser_max, going, waist, landing):
    """The stair type for these sizes, made if it does not exist yet.

    One type per size set, named like Concrete Stair - SR180 SG300 SD100
    LD100, with its own run type (SD) and landing type (LD). Needs an open
    transaction. Returns the StairsType.
    """
    name = type_name(riser_max, going, waist, landing)
    stair_type = _by_name(doc, StairsType, name)
    if stair_type is not None:
        return stair_type
    base = _base_stair_type(doc)

    run_name = u"Monolithic Run - SD{0:g}".format(round(waist, 1))
    run_type = _by_name(doc, StairsRunType, run_name)
    if run_type is None:
        source = _monolithic(doc, StairsRunType, u"Monolithic Run", base.RunType)
        run_type = source.Duplicate(run_name)
        run_type.StructuralDepth = _ft(waist)

    landing_name = u"Monolithic Landing - LD{0:g}".format(round(landing, 1))
    landing_type = _by_name(doc, StairsLandingType, landing_name)
    if landing_type is None:
        source = _monolithic(doc, StairsLandingType, u"Monolithic Landing", base.LandingType)
        landing_type = source.Duplicate(landing_name)
        landing_type.Thickness = _ft(landing)

    stair_type = base.Duplicate(name)
    stair_type.MaxRiserHeight = _ft(riser_max)
    stair_type.MinTreadDepth = _ft(going)
    stair_type.RunType = run_type.Id
    stair_type.LandingType = landing_type.Id
    return stair_type


def _material_params(element_type):
    """The writable material parameters of a run or landing type (the
    monolithic material), leaving out the tread/riser ones."""
    for p in element_type.Parameters:
        bip = str(p.Definition.BuiltInParameter)
        if (p.StorageType == DB.StorageType.ElementId and not p.IsReadOnly
                and u"MATERIAL" in bip and u"TRISER" not in bip):
            yield p


def stair_material(doc, stairs):
    """The monolithic material of a stair's run type, or None."""
    for run_id in stairs.GetStairsRuns():
        run_type = doc.GetElement(doc.GetElement(run_id).GetTypeId())
        for p in _material_params(run_type):
            if p.AsElementId() != DB.ElementId.InvalidElementId:
                return p.AsElementId()
    return None


def set_material(doc, stairs, material_id):
    """Set the monolithic material on the stair's run and landing types.
    Needs an open transaction. The types are shared by every stair of the
    same sizes, so those change too."""
    ids = [doc.GetElement(i).GetTypeId() for i in stairs.GetStairsRuns()]
    ids += [doc.GetElement(i).GetTypeId() for i in stairs.GetStairsLandings()]
    for type_id in set(ids):
        for p in _material_params(doc.GetElement(type_id)):
            if p.AsElementId() != material_id:
                p.Set(material_id)


# --- Building ---
def remove_railings(doc, stairs):
    """Delete the railings Revit adds to a new stair by itself. Needs an
    open transaction. Returns how many went."""
    from Autodesk.Revit.DB.Architecture import Railing
    ids = [r.Id for r in DB.FilteredElementCollector(doc).OfClass(Railing)
           if r.HostId == stairs.Id]
    if ids:
        doc.Delete(List[DB.ElementId](ids))
    return len(ids)


def chain_loop(curves, z):
    """Order picked curves into one closed CurveLoop flat at height z.

    Lines may be picked in any order and drawn either way round; ends
    within 1 mm count as joined. Raises ValueError if they do not close.
    """
    tol = _ft(1.0)
    flat = []
    for c in curves:
        a, b = c.GetEndPoint(0), c.GetEndPoint(1)
        flat.append((DB.XYZ(a.X, a.Y, z), DB.XYZ(b.X, b.Y, z)))
    if len(flat) < 3:
        raise ValueError(u"Pick at least three lines that close a shape.")
    loop_pts = [flat[0][0], flat[0][1]]
    left = flat[1:]
    while left:
        tip = loop_pts[-1]
        for i, (a, b) in enumerate(left):
            if a.DistanceTo(tip) < tol:
                loop_pts.append(b)
                break
            if b.DistanceTo(tip) < tol:
                loop_pts.append(a)
                break
        else:
            raise ValueError(u"The picked lines do not join end to end.")
        left.pop(i)
    if loop_pts[-1].DistanceTo(loop_pts[0]) > tol:
        raise ValueError(u"The picked lines do not close a shape.")
    loop = DB.CurveLoop()
    for i in range(len(loop_pts) - 1):
        loop.Append(DB.Line.CreateBound(loop_pts[i], loop_pts[i + 1] if i < len(loop_pts) - 2
                                        else loop_pts[0]))
    return loop


def landing_slab(doc, curves, floor_type, level, offset):
    """Floor from picked lines, top at level + offset (mm). Needs an open
    transaction. Returns the Floor."""
    z = level.ProjectElevation
    loop = chain_loop(curves, z)
    loops = List[DB.CurveLoop]()
    loops.Add(loop)
    floor = DB.Floor.Create(doc, loops, floor_type.Id, level.Id)
    floor.get_Parameter(DB.BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).Set(_ft(offset))
    return floor


def build(doc, stair_type, base_level, base_offset, top_level, top_offset,
          start, direction, width, sizes):
    """Build the stair. Returns the new Stairs element.

    start      XYZ at the foot of the first riser, on the stair centreline
               (only X and Y are used)
    direction  XYZ pointing up the stair, in plan
    sizes      the dict from plan()
    Opens its own StairsEditScope; call it with no transaction open.
    """
    d = DB.XYZ(direction.X, direction.Y, 0).Normalize()
    side = DB.XYZ(-d.Y, d.X, 0)
    base_z = base_level.ProjectElevation + _ft(base_offset)
    foot = DB.XYZ(start.X, start.Y, base_z)
    run_end = foot + d * _ft(sizes["run_length"])

    # The stair is made between two throwaway levels at exactly its base and
    # top heights, then pointed at the real levels and offsets. Two reasons:
    # StairsEditScope refuses the same level for base and top (the usual
    # entry stair, one level, two offsets), and a run takes its riser height
    # from the stair as it was when the scope started - changing offsets
    # inside the scope leaves the run a riser short of the landing.
    top_z = top_level.ProjectElevation + _ft(top_offset)
    t = DB.Transaction(doc, u"Concrete Stair - temporary levels")
    t.Start()
    temp_base = DB.Level.Create(doc, base_z)
    temp_top = DB.Level.Create(doc, top_z)
    t.Commit()

    scope = DB.StairsEditScope(doc, u"Concrete Stair")
    stairs_id = scope.Start(temp_base.Id, temp_top.Id)
    t = DB.Transaction(doc, u"Concrete Stair")
    t.Start()
    try:
        stairs = doc.GetElement(stairs_id)
        stairs.ChangeTypeId(stair_type.Id)
        stairs.DesiredRisersNumber = sizes["risers"]
        stairs.ActualTreadDepth = _ft(sizes["going"])

        run = StairsRun.CreateStraightRun(doc, stairs_id, DB.Line.CreateBound(foot, run_end),
                                          StairsRunJustification.Center)
        run.ActualRunWidth = _ft(width)
        run.BeginsWithRiser = True
        run.EndsWithRiser = True

        if sizes["xyz"] > 1.0:
            half = side * (_ft(width) / 2.0)
            far = run_end + d * _ft(sizes["xyz"])
            corners = [run_end - half, far - half, far + half, run_end + half]
            loop = DB.CurveLoop()
            for i in range(4):
                a, b = corners[i], corners[(i + 1) % 4]
                loop.Append(DB.Line.CreateBound(DB.XYZ(a.X, a.Y, base_z),
                                                DB.XYZ(b.X, b.Y, base_z)))
            StairsLanding.CreateSketchedLanding(doc, stairs_id, loop, _ft(sizes["rise"]))
        t.Commit()
        scope.Commit(_Quiet())

        # onto the real levels; the heights do not change
        t = DB.Transaction(doc, u"Concrete Stair - levels")
        t.Start()
        stairs = doc.GetElement(stairs_id)
        P = DB.BuiltInParameter
        stairs.get_Parameter(P.STAIRS_BASE_LEVEL_PARAM).Set(base_level.Id)
        stairs.get_Parameter(P.STAIRS_BASE_OFFSET).Set(_ft(base_offset))
        stairs.get_Parameter(P.STAIRS_TOP_LEVEL_PARAM).Set(top_level.Id)
        stairs.get_Parameter(P.STAIRS_TOP_OFFSET).Set(_ft(top_offset))
        stairs.DesiredRisersNumber = sizes["risers"]
        t.Commit()
    except Exception:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        if scope.IsActive:
            scope.Cancel()
        raise
    finally:
        t = DB.Transaction(doc, u"Concrete Stair - remove temporary levels")
        t.Start()
        doc.Delete(List[DB.ElementId]([temp_base.Id, temp_top.Id]))
        t.Commit()
    return doc.GetElement(stairs_id)
