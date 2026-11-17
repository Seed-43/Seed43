# -*- coding: utf-8 -*-
"""Concrete Stair reinforcing: Revit rebar hosted on the stair it built.

Everything is set out in the stair's own section, in mm:

    x  along the stair, from the face of the first riser, uphill
    y  across the stair, from its centreline
    z  up, from the base (bottom of the first riser)

and turned into model points with the stair's start point and direction.
The bars follow the typical detail: main bars down the soffit, bent into
the footing with a hook and turned into the landing at the top; bars
across the stair on top of them up the slope; an optional mesh at the top
landing; and an optional footing beam under the first step.
"""
import math

from System.Collections.Generic import List
from Autodesk.Revit import DB
from Autodesk.Revit.DB.Structure import (Rebar, RebarStyle, RebarHookOrientation,
                                         StructuralType, RebarHostData, RebarCoverType)

__all__ = ["Frame", "layout", "layout_names", "LAYOUTS", "numbers", "min_waist", "reinforce", "find_footings",
           "remove_footings", "from_stairs"]

# ── CONSTANTS ───────────────────────────────────────────────────────────────

MM = 304.8


# ── HELPERS (private to this file) ──────────────────────────────────────────

def _ft(mm):
    return mm / MM


def _dia(bar_type):
    return bar_type.BarNominalDiameter * MM


class Frame(object):
    """The stair's section: turns (x, y, z) in mm into model points."""

    # --- construction ---
    def __init__(self, foot, direction, sizes, width):
        self.foot = foot                                  # XYZ, model, at the base
        self.d = DB.XYZ(direction.X, direction.Y, 0).Normalize()
        self.side = DB.XYZ(-self.d.Y, self.d.X, 0)
        self.s = sizes
        self.width = width
        self.slope = math.atan2(sizes["riser"], sizes["going"])

    # --- public methods ---
    def p(self, x, y, z):
        return (self.foot + self.d * _ft(x) + self.side * _ft(y)
                + DB.XYZ.BasisZ * _ft(z))

    def up_slope(self):
        """Unit vector up the stair along the soffit, in the model."""
        return (self.d * math.cos(self.slope) + DB.XYZ.BasisZ * math.sin(self.slope)).Normalize()

    def soffit_line(self, inset):
        """(t, k) for the line z = t*x + k parallel to the soffit, `inset`
        mm inside it (measured square to the soffit)."""
        t = self.s["riser"] / self.s["going"]
        k = (inset - self.s["waist"]) / math.cos(self.slope)
        return t, k


def _line(a, b):
    return DB.Line.CreateBound(a, b)


def _curves(*curves):
    out = List[DB.Curve]()
    for c in curves:
        out.Add(c)
    return out


def _bar(doc, host, bar_type, normal, curves, end_hook=None, end_orient=None):
    return Rebar.CreateFromCurves(
        doc, RebarStyle.Standard, bar_type, None, end_hook, host, normal,
        _curves(*curves), RebarHookOrientation.Left,
        end_orient or RebarHookOrientation.Left, True, True)


def _spread(rebar, spacing, length):
    """Lay a single bar out as a set at a maximum spacing over a length."""
    rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(
        _ft(spacing), _ft(length), True, True, True)


def _cover_type(doc, mm):
    """A rebar cover type of this distance, made if the model has none."""
    for c in DB.FilteredElementCollector(doc).OfClass(RebarCoverType):
        if abs(c.CoverDistance * MM - mm) < 0.01:
            return c
    return RebarCoverType.Create(doc, u"Stair Rebar Cover {0:g}".format(round(mm, 1)), _ft(mm))


def _set_cover(doc, element, bip, mm):
    p = element.get_Parameter(bip)
    if p is not None and not p.IsReadOnly:
        cover = _cover_type(doc, mm)
        if p.AsElementId() != cover.Id:
            p.Set(cover.Id)


def set_covers(doc, stairs, footing, opts):
    """Give the stair and its footing the covers the bars are laid out to.

    Revit keeps bars out of a host's cover zone, and moves them when they
    are inside it: with the usual 50 mm the landing leg dropped below the
    mesh and the bar in the lower bend was pushed down. The stair takes the
    smaller of the top and side covers (it has one cover for every face).
    The footing's top is only the stair's base, so it gets a nominal cover;
    its bottom and ends take the footing cover. Needs an open transaction.
    """
    _set_cover(doc, stairs, DB.BuiltInParameter.CLEAR_COVER,
               min(opts["cover_top"], opts["cover_side"]))
    if footing is not None:
        _set_cover(doc, footing, DB.BuiltInParameter.CLEAR_COVER_TOP, 20.0)
        _set_cover(doc, footing, DB.BuiltInParameter.CLEAR_COVER_BOTTOM, opts["cover_bottom"])
        _set_cover(doc, footing, DB.BuiltInParameter.CLEAR_COVER_OTHER,
                   min(opts.get("cover_end", opts["cover_bottom"]), opts["cover_side"]))


def _footing_host(doc, stairs):
    """The footing beam under the stair if it can host bars, else None."""
    try:
        beams = find_footings(doc, stairs)
    except ValueError:
        return None
    for beam in beams:
        data = RebarHostData.GetRebarHostData(beam)
        if data is not None and data.IsValidHost():
            return beam
    return None


def _hook_tail(rebar, start):
    """The free end of the hook: of the bar's two ends, the one away from
    `start` (the far end of the top leg). Revit may hand the centreline back
    in either direction, so the last curve is not always the hooked end."""
    curves = list(rebar.GetCenterlineCurves(False, False, False,
                                            DB.Structure.MultiplanarOption.IncludeOnlyPlanarCurves, 0))
    ends = [curves[0].GetEndPoint(0), curves[-1].GetEndPoint(1)]
    return max(ends, key=lambda p: p.DistanceTo(start))


# ── CORE LOGIC ──────────────────────────────────────────────────────────────

def min_waist(o):
    """The thinnest slope (SD, mm) that fits the bars: the clear below the
    step corners, the bar across, the main bar under it and the soffit
    cover, all square to the slope, rounded up to the next 5 mm."""
    need = o["corner_clear"] + o["main_dm"] + o["dist_dm"] + o["cover_soffit"]
    return math.ceil(need / 5.0 - 1e-9) * 5.0


def _in_bend(vertex, u1, u2, bend, dm, dd):
    """Centre of a bar across sitting in the inside of a bend, touching the
    bent bar. vertex is where the two straight legs would meet, u1 and u2
    unit vectors along the legs away from it, bend the inside bend
    diameter, dm and dd the two bars' diameters."""
    bx, bz = u1[0] + u2[0], u1[1] + u2[1]
    n = math.hypot(bx, bz)
    bx, bz = bx / n, bz / n
    half = math.acos(max(-1.0, min(1.0, u1[0] * u2[0] + u1[1] * u2[1]))) / 2.0
    rc = bend / 2.0 + dm / 2.0              # centreline radius of the bend
    c = dm / 2.0 + dd / 2.0
    if rc >= c:
        dist = rc / math.sin(half) - rc + c   # against the curved part
    else:
        dist = c / math.sin(half)             # against both straight legs
    return (vertex[0] + bx * dist, vertex[1] + bz * dist)


def _layout_hooked(s, o):
    """Where every bar goes, in the stair's section (mm, x along, z up).

    Pure arithmetic, no Revit: the Rebar tab draws from this and reinforce()
    builds from it, so the picture and the model cannot disagree.

    s  the stair sizes (stair_engine.plan, or read off a stair)
    o  numbers(): covers, corner_clear, lap, main_d / main_dm (nominal and
       model), dist_d / dist_dm and mesh_d (0 when off), main_bend and
       hook_bend (inside diameters), spacings, footing (width, depth) or None

    Up the slope the bars across are the top layer, `corner_clear` below the
    steps' inside corners (measured to their top, square to the slope), with
    the main bar under them and the soffit cover below that. One more bar
    across sits in the inside of each bend, and one in the bottom of the
    hook.
    """
    slope = math.atan2(s["riser"], s["going"])
    t = s["riser"] / s["going"]
    cs, sn = math.cos(slope), math.sin(slope)

    def line_k(inset):
        """k of the line z = t*x + k, `inset` mm above the soffit (square)."""
        return (inset - s["waist"]) / cs

    dm, dd = o["main_dm"], o["dist_dm"]
    cb, ct = o["cover_bottom"], o["cover_top"]
    ce = o.get("cover_end", cb)
    out = {"footing": None, "hook": False, "across": [], "bends": [], "mesh": None}

    # main bars, under the bars across; the steps' inside corners lie on
    # z = t*x, `waist` above the soffit
    k = line_k(s["waist"] - o["corner_clear"] - dd - dm / 2.0)
    # top leg flush with the mesh's bottom bars (same direction), under its top bars
    z_top = s["rise"] - ct - o["mesh_d"] - dm / 2.0
    x_top = (z_top - k) / t
    main = [(x_top + o["lap"], z_top), (x_top, z_top)]
    bends = []
    if dd:
        bends.append(_in_bend((x_top, z_top), (1.0, 0.0), (-cs, -sn),
                              o["main_bend"], dm, dd))
    if o["footing"]:
        fw, fd = o["footing"]
        out["footing"] = (0.0, fw, -fd, 0.0)
        # bend just below the base, then straight down, hooked at the bottom
        x_k = max((-min(fd / 2.0, 50.0) - k) / t, ce + dm / 2.0)
        x_k = min(x_k, fw - ce - dm / 2.0)
        z_k = t * x_k + k
        # Revit takes the end of a hooked bar as the outside of the hook,
        # so the end sits on the bottom cover
        z_end = -(fd - cb)
        main += [(x_k, z_k), (x_k, z_end)]
        out["hook"] = True
        rh = o["hook_bend"] / 2.0 + dm / 2.0                   # centreline radius
        out["hook_radius"] = rh
        out["hook_centre"] = (x_k + rh, z_end + dm / 2.0 + rh)
        if dd:
            bends.append(_in_bend((x_k, z_k), (cs, sn), (0.0, -1.0),
                                  o["main_bend"], dm, dd))
            # in the bottom of the hook, which turns uphill
            hx, hz = out["hook_centre"]
            bends.append((hx, hz - (rh - dm / 2.0 - dd / 2.0)))
    else:
        # no footing: stop at the soffit, the bottom cover above the base
        x_b = max((cb - k) / t, 0.0)
        main.append((x_b, t * x_b + k))
    out["main"] = main
    out["z_top"] = z_top
    out["bends"] = bends

    out["main_k"] = k

    # bars across up the slope, on top of the main bars, between the bend bars
    if dd:
        k2 = line_k(s["waist"] - o["corner_clear"] - dd / 2.0)
        x_t = bends[0][0]                                   # the top bend bar
        x_b = bends[1][0] if o["footing"] else max((cb - k2) / t, 0.0)
        length = (x_t - x_b) / cs
        if length > 1.0:
            gaps = int(math.ceil(length / o["dist_spacing"] - 1e-6))
            first = 1 if o["footing"] else 0                # the bend bar is the first
            xs = [x_b + (x_t - x_b) * i / float(gaps) for i in range(first, gaps)]
            out["across"] = [(x, t * x + k2) for x in xs]

    # mesh in the stair's own top landing
    if o["mesh_d"] and s["xyz"] > 2 * o["cover_side"]:
        out["mesh"] = (s["run_length"] + o["cover_side"],
                       s["run_length"] + s["xyz"] - o["cover_side"],
                       s["rise"] - ct - o["mesh_d"] / 2.0)
    return out


# --- Layouts ---
# Each rebar layout is a function (s, o) -> plan, returning the same keys as
# _layout_hooked: footing, main, hook, hook_radius, hook_centre, main_k,
# z_top, across, bends, mesh. reinforce() and the Rebar tab's drawing work
# from the plan alone, so a new layout only needs a function and a line in
# LAYOUTS: (key saved in settings, name shown in the Layout list, function).
LAYOUTS = [
    (u"hooked", u"Main bars hooked into the footing", _layout_hooked),
]


def layout_names():
    """[(key, name)] of the layouts, in the order the Layout list shows them."""
    return [(k, n) for k, n, _ in LAYOUTS]


def layout(s, o):
    """Where every bar goes, for the layout named by o["layout"] (the first
    one when it is missing or unknown). See _layout_hooked for the units."""
    for key, _name, fn in LAYOUTS:
        if key == o.get("layout"):
            return fn(s, o)
    return LAYOUTS[0][2](s, o)


def _model_dia(bar_type):
    return bar_type.BarModelDiameter * MM


def numbers(opts):
    """The plain numbers layout() needs, from reinforce()'s options."""
    dist = opts.get("dist_bar")
    return {
        "cover_top": opts["cover_top"], "cover_soffit": opts["cover_soffit"],
        "cover_bottom": opts["cover_bottom"], "cover_side": opts["cover_side"],
        "layout": opts.get("layout"),
        "corner_clear": opts.get("corner_clear", 50.0),
        "cover_end": opts.get("cover_end", opts["cover_bottom"]),
        "lap": opts["lap"], "main_d": _dia(opts["main_bar"]),
        "main_dm": _model_dia(opts["main_bar"]),
        "main_bend": opts["main_bar"].StandardBendDiameter * MM,
        "hook_bend": opts["main_bar"].StandardHookBendDiameter * MM,
        "dist_d": _dia(dist) if dist else 0.0,
        "dist_dm": _model_dia(dist) if dist else 0.0,
        "mesh_d": _dia(opts["mesh_bar"]) if opts.get("mesh_bar") else 0.0,
        "dist_spacing": opts["dist_spacing"], "mesh_spacing": opts["mesh_spacing"],
        "footing": opts.get("footing_size") if opts.get("footing") else None,
    }


def reinforce(doc, stairs, frame, level, opts):
    """Add the bars (the footing is placed first, by the caller). Needs an
    open transaction.

    opts (mm, Revit types):
        main_bar, main_spacing, main_hook   main bars, down the soffit
        lap                                 top leg into the landing
        cover_top, cover_soffit, cover_side, cover_bottom, corner_clear
        dist_bar, dist_spacing              bars across, up the slope (or None)
        mesh_bar, mesh_spacing              top landing mesh (or None)
        footing, footing_size               True and (width, depth), or None
    Returns a dict of what was made.
    """
    made = {"footing": 0, "main": 0, "across": 0, "mesh": 0}
    w = frame.width
    cs = opts["cover_side"]
    db = _dia(opts["main_bar"])
    y0, y1 = -w / 2.0 + cs, w / 2.0 - cs
    plan = layout(frame.s, numbers(opts))

    footing_host = _footing_host(doc, stairs) if plan["footing"] else None
    set_covers(doc, stairs, footing_host, opts)
    doc.Regenerate()

    # --- main bars: landing leg, down the soffit, into the footing ---
    y = y0 + db / 2.0
    pts = [frame.p(x, y, z) for x, z in plan["main"]]
    curves = [_line(pts[i], pts[i + 1]) for i in range(len(pts) - 1)]
    hook = opts.get("main_hook") if plan["hook"] else None
    main = _bar(doc, stairs, opts["main_bar"], frame.side, curves, hook)
    if hook is not None:
        # The hook has to turn into the footing, uphill. Revit's Left/Right
        # depends on how the bar is drawn, so check and flip if it is wrong.
        doc.Regenerate()
        tail = _hook_tail(main, pts[0])
        if (tail - pts[-1]).DotProduct(frame.d) < 0:
            doc.Delete(main.Id)
            main = _bar(doc, stairs, opts["main_bar"], frame.side, curves, hook,
                        RebarHookOrientation.Right)
    _spread(main, opts["main_spacing"], w - 2 * cs - db)
    made["main"] = main.NumberOfBarPositions

    # --- bars across: a set up the slope under the main bars, then one in
    # each bend and one in the bottom of the hook ---
    if plan["across"]:
        pts_a = plan["across"]
        x_s, z_s = pts_a[0]
        across = _bar(doc, stairs, opts["dist_bar"], frame.up_slope(),
                      [_line(frame.p(x_s, y0, z_s), frame.p(x_s, y1, z_s))])
        if len(pts_a) > 1:
            x_e, z_e = pts_a[-1]
            across.GetShapeDrivenAccessor().SetLayoutAsFixedNumber(
                len(pts_a), _ft(math.hypot(x_e - x_s, z_e - z_s)), True, True, True)
        made["across"] += across.NumberOfBarPositions
    # the ones below the base are in the footing: host them on it when it can
    for x_c, z_c in plan["bends"]:
        host = footing_host if (z_c < 0 and footing_host is not None) else stairs
        _bar(doc, host, opts["dist_bar"], frame.d,
             [_line(frame.p(x_c, y0, z_c), frame.p(x_c, y1, z_c))])
        made["across"] += 1

    # --- top landing mesh, both ways ---
    if plan["mesh"]:
        xa, xb, z1 = plan["mesh"]
        md = _dia(opts["mesh_bar"])
        across = _bar(doc, stairs, opts["mesh_bar"], frame.d,
                      [_line(frame.p(xa, y0, z1), frame.p(xa, y1, z1))])
        _spread(across, opts["mesh_spacing"], xb - xa)
        along = _bar(doc, stairs, opts["mesh_bar"], frame.side,
                     [_line(frame.p(xa, y0, z1 - md), frame.p(xb, y0, z1 - md))])
        _spread(along, opts["mesh_spacing"], y1 - y0)
        made["mesh"] = across.NumberOfBarPositions + along.NumberOfBarPositions
    return made


def eid(element_id):
    """Integer value of an ElementId on any Revit (Value from 2024)."""
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue


def remove_bars_on(doc, host):
    """Delete every bar hosted on `host` (the footing beam, before its stair's
    bars are made again). Needs an open transaction."""
    for r in list(DB.FilteredElementCollector(doc).OfClass(Rebar)):
        if r.GetHostId() == host.Id:
            doc.Delete(r.Id)


def has_rebar(doc, stairs):
    """True if any rebar is hosted on the stair."""
    for r in DB.FilteredElementCollector(doc).OfClass(Rebar):
        if r.GetHostId() == stairs.Id:
            return True
    return False


def remove_footings(doc, stairs):
    """Delete the footing beam this tool put under a stair. Needs an open
    transaction. Returns how many went."""
    gone = find_footings(doc, stairs)
    for beam in gone:
        doc.Delete(beam.Id)
    return len(gone)


def find_footings(doc, stairs):
    """The footings under a stair, as placed by this tool: Stair Footing
    instances, and the framing beams earlier versions used.

    Matched by where they are: a Stair Footing whose origin (its top, on
    the face toward the stair) is on the stair's centre line just in from
    the first riser at the base; or a framing beam lying across the stair at
    its first step, top at the base, about as long as the stair is wide.
    Anything else is left alone.
    """
    frame, _level = from_stairs(doc, stairs)
    w = frame.width
    found = []
    for inst in DB.FilteredElementCollector(doc).OfCategory(
            DB.BuiltInCategory.OST_StructuralFoundation).WhereElementIsNotElementType():
        sym = getattr(inst, "Symbol", None)
        if sym is None or sym.Family.Name != u"Stair Footing":
            continue
        rel = inst.Location.Point - frame.foot
        if (abs(rel.Z) * MM < 5.0 and 0.0 <= rel.DotProduct(frame.d) * MM <= 1500.0
                and abs(rel.DotProduct(frame.side)) * MM < 50.0):
            found.append(inst)
    for beam in DB.FilteredElementCollector(doc).OfCategory(
            DB.BuiltInCategory.OST_StructuralFraming).WhereElementIsNotElementType():
        curve = getattr(beam.Location, "Curve", None)
        if curve is None:
            continue
        a, b = curve.GetEndPoint(0), curve.GetEndPoint(1)
        mid = (a + b) * 0.5
        rel = mid - frame.foot
        along = rel.DotProduct(frame.d) * MM
        across = rel.DotProduct(frame.side) * MM
        length = a.DistanceTo(b) * MM
        direction = (b - a).Normalize()
        if (abs(mid.Z - frame.foot.Z) * MM < 5.0
                and -50.0 <= along <= 1000.0 and abs(across) < 50.0
                and abs(length - w) < max(0.05 * w, 20.0)
                and abs(direction.DotProduct(frame.side)) > 0.99):
            found.append(beam)
    return found


def from_stairs(doc, stairs):
    """Read what reinforce() needs off an existing straight stair.

    Returns (frame, base_level). Raises ValueError for a stair that is not
    one straight monolithic run.
    """
    runs = [doc.GetElement(i) for i in stairs.GetStairsRuns()]
    if len(runs) != 1:
        raise ValueError(u"Pick a stair with one straight run.")
    run = runs[0]
    path = list(run.GetStairsPath())
    if len(path) != 1 or not isinstance(path[0], DB.Line):
        raise ValueError(u"Pick a stair with one straight run.")
    run_type = doc.GetElement(run.GetTypeId())
    waist = getattr(run_type, "StructuralDepth", 0.0) * MM
    if waist <= 0:
        raise ValueError(u"That stair has no waist (its run is not monolithic), "
                         u"so there is no slab to reinforce.")
    a, b = path[0].GetEndPoint(0), path[0].GetEndPoint(1)
    direction = DB.XYZ(b.X - a.X, b.Y - a.Y, 0).Normalize()
    foot = DB.XYZ(a.X, a.Y, stairs.BaseElevation)
    riser = stairs.ActualRiserHeight * MM
    going = stairs.ActualTreadDepth * MM
    risers = stairs.ActualRisersNumber
    run_length = DB.XYZ(b.X - a.X, b.Y - a.Y, 0).GetLength() * MM
    # top landing: how far it runs past the end of the run, and how thick
    xyz, landing = 0.0, 0.0
    for lid in stairs.GetStairsLandings():
        lnd = doc.GetElement(lid)
        landing = max(landing, lnd.Thickness * MM)
        for c in lnd.GetFootprintBoundary():
            for pt in (c.GetEndPoint(0), c.GetEndPoint(1)):
                along = DB.XYZ(pt.X - a.X, pt.Y - a.Y, 0).DotProduct(direction) * MM
                xyz = max(xyz, along - run_length)
    if landing <= 0:
        landing = waist
    sizes = {"riser": riser, "going": going, "risers": risers,
             "rise": riser * risers, "waist": waist, "landing": landing,
             "run_length": run_length, "xyz": xyz}
    level = doc.GetElement(stairs.get_Parameter(
        DB.BuiltInParameter.STAIRS_BASE_LEVEL_PARAM).AsElementId())
    return Frame(foot, direction, sizes, run.ActualRunWidth * MM), level
