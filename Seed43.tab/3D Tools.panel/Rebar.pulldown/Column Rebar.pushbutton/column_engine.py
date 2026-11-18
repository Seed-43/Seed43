# -*- coding: utf-8 -*-
"""Column / pile reinforcing: native Revit rebar hosted on a concrete column.

Everything is set out in the column's own frame, in mm:

    x, y  across the column, from its centre (x along the family's width b,
          y along its depth h)
    z     up, from the column's base (its underside)

and turned into model points with the column's centre, axes and base.

layout() is pure arithmetic (no Revit): the window draws from it and
reinforce() builds from it, so the drawing and the model cannot disagree.
Each rebar layout is a line in LAYOUTS: its key (saved in settings), the
name shown in the Rebar layout list, the column shape it is for and the
defaults that make it, read off a sample column in the model.
"""
import math

from System import Int64
from System.Collections.Generic import List
from Autodesk.Revit import DB
from Autodesk.Revit.DB.Structure import (Rebar, RebarStyle, RebarHookOrientation,
                                         RebarHostData, RebarCoverType, RebarShape,
                                         RebarShapeDefinitionByArc, RebarBarType,
                                         RebarHookType, MultiplanarOption)

__all__ = ["Frame", "LAYOUTS", "layout_names", "layout_info", "numbers", "layout",
           "tie_levels", "reinforce", "remove_bars", "from_host", "read_existing",
           "eid", "bar_length", "MM", "column_types", "type_shape", "type_size", "type_name",
           "ensure_type", "place_column", "update_column", "column_levels"]

# ── CONSTANTS ───────────────────────────────────────────────────────────────

MM = 304.8
ARRAY_STEP = 5.0            # mm between the bars of a set before each is moved to its place

# --- Layouts ---
# (key, name in the list, shape, defaults). The defaults are the sample
# column's settings; a layout's shape decides which section and elevation
# the window draws and which ties reinforce() makes. To add a layout, add a
# line here (and, if its bars differ, its case in layout() and its drawing
# in the window's DRAWERS).
LAYOUTS = [
    (u"pile_spiral", u"Bored pile - spiral, bars hooked in at the toe", u"round",
     {"tie_kind": u"spiral", "cover_side": u"75", "cover_top": u"50", "cover_bottom": u"130",
      "main_bar": u"HD20", "main_count": u"8", "arrangement": u"axes",
      "main_hook": u"Standard Hook - 90 deg.", "hook_dir": u"in", "hook_leg": u"0",
      "lap": u"0", "bottom_ext": u"0",
      "tie_bar": u"HR12", "tie_first": u"165", "tie_last": u"85",
      "zone_bottom_len": u"0", "zone_bottom_sp": u"125", "zone_mid_sp": u"125",
      "zone_top_len": u"0", "zone_top_sp": u"125", "turns": u"3",
      "starters": False}),
    (u"rect_ties", u"Square column - perimeter and inner ties, starters", u"rect",
     {"tie_kind": u"rect", "cover_side": u"50", "cover_top": u"50", "cover_bottom": u"50",
      "main_bar": u"HD20", "bars_b": u"4", "bars_h": u"4",
      "main_hook": u"(no hook)", "hook_dir": u"in", "hook_leg": u"0",
      "lap": u"0", "bottom_ext": u"0",
      "tie_bar": u"HR10", "tie_hook": u"Stirrup/Tie Seismic - 135 deg.",
      "tie_first": u"120", "tie_last": u"55", "inner_ties": True,
      "zone_bottom_len": u"1280", "zone_bottom_sp": u"75", "zone_mid_sp": u"150",
      "zone_top_len": u"0", "zone_top_sp": u"75",
      "starters": True, "starter_bar": u"HD20", "starter_embed": u"480",
      "starter_foot": u"320", "starter_lap": u"1300", "starter_dir": u"out"}),
    (u"round_ties", u"Round column - circular ties, bars hooked out into the cap", u"round",
     {"tie_kind": u"circle", "cover_side": u"50", "cover_top": u"55", "cover_bottom": u"50",
      "main_bar": u"HD16", "main_count": u"8", "arrangement": u"between",
      "main_hook": u"Standard Hook - 90 deg.", "hook_dir": u"out", "hook_leg": u"200",
      "lap": u"0", "bottom_ext": u"485",
      "tie_bar": u"HR10", "tie_first": u"-365", "tie_last": u"85",
      "zone_bottom_len": u"0", "zone_bottom_sp": u"150", "zone_mid_sp": u"150",
      "zone_top_len": u"0", "zone_top_sp": u"150",
      "starters": False}),
]


# ── HELPERS (private to this file) ──────────────────────────────────────────

def _ft(mm):
    return mm / MM


def _name(element):
    try:
        return DB.Element.Name.GetValue(element)
    except Exception:
        return u""


def eid(element_id):
    """Integer value of an ElementId on any Revit (Value from 2024)."""
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue


def _curves(curves):
    out = List[DB.Curve]()
    for c in curves:
        out.Add(c)
    return out


def _gaps(length, spacing):
    """How many gaps Revit's Maximum Spacing layout puts in `length`."""
    if length <= 1e-6 or spacing <= 0:
        return 0
    return int(math.ceil(length / spacing - 1e-6))


class Frame(object):
    """The column's frame: turns (x, y, z) in mm into model points."""

    # --- construction ---
    def __init__(self, centre, ux, uy, base_z, shape, size, length):
        self.centre = centre                    # XYZ at the base, on the axis
        self.ux = DB.XYZ(ux.X, ux.Y, 0).Normalize()
        # always right-handed (a mirrored family's BasisY points the other way):
        # angles in layout() then turn the same way as in the model
        self.uy = DB.XYZ.BasisZ.CrossProduct(self.ux)
        self.base_z = base_z                    # ft
        self.shape = shape                      # "round" or "rect"
        self.size = size                        # (D, D) or (b, h), mm
        self.length = length                    # mm

    # --- public methods ---
    def p(self, x, y, z):
        return DB.XYZ(self.centre.X, self.centre.Y, self.base_z) + \
            self.ux * _ft(x) + self.uy * _ft(y) + DB.XYZ.BasisZ * _ft(z)

    def v(self, x, y, z=0.0):
        return (self.ux * x + self.uy * y + DB.XYZ.BasisZ * z).Normalize()


# ── CORE LOGIC: the layout, no Revit ────────────────────────────────────────

def layout_names():
    """[(key, name)] of the layouts, in the order the list shows them."""
    return [(k, n) for k, n, _s, _d in LAYOUTS]


def layout_info(key):
    """(name, shape, defaults) of a layout; the first one when unknown."""
    for k, n, s, d in LAYOUTS:
        if k == key:
            return n, s, d
    return LAYOUTS[0][1], LAYOUTS[0][2], LAYOUTS[0][3]


def tie_levels(o):
    """The ties' zones as (z start, z end, spacing, include first): the
    bottom zone, the middle and the top zone, each laid out by Revit at a
    maximum spacing. A zone of no length is left out; the middle starts
    where the bottom zone ends without doubling the tie there."""
    z0 = o["tie_first"]
    z1 = o["L"] - o["tie_last"]
    if z1 <= z0:
        return []
    zb = min(z0 + max(o["zone_bottom_len"], 0.0), z1)
    zt = max(z1 - max(o["zone_top_len"], 0.0), zb)
    sets = []
    if zb > z0 + 1e-6:
        sets.append((z0, zb, o["zone_bottom_sp"], True))
    if zt > zb + 1e-6:
        sets.append((zb, zt, o["zone_mid_sp"], not sets))
    elif not sets:
        sets.append((z0, z0, o["zone_mid_sp"], True))
    if z1 > zt + 1e-6:
        sets.append((zt, z1, o["zone_top_sp"], not sets))
    return sets


def _zs(sets):
    """Every tie's height, as Revit spaces them."""
    out = []
    for za, zb, sp, first in sets:
        n = _gaps(zb - za, sp)
        zs = [za + (zb - za) * i / float(n) for i in range(n + 1)] if n else [za]
        out += zs if first else zs[1:]
    return out


def layout(o):
    """Where every bar goes, in the column's frame (mm). See numbers() for o.

    Returns:
        shape, D / b, h, L        the column
        main      pts [(x, y)], z0, z1 (the straight part; a hook is added
                  below z0), hook "in" / "out" / None, ang [radial angle of each]
        starters  None or pts [(x, y, foot dx, foot dy)], z_top, z_foot, foot
        ties      kind ("spiral", "circle", "rect"), r (spiral, circle: the
                  centreline radius), boxes [(x0, y0, x1, y1)] (rect: centrelines),
                  sets (see tie_levels), zs (every tie), pitch, turns
    """
    c, cb, ct = o["cover_side"], o["cover_bottom"], o["cover_top"]
    dm, td = o["main_dm"], o["tie_dm"]
    L = o["L"]
    out = {"shape": o["shape"], "L": L, "D": o["D"], "b": o["b"], "h": o["h"]}

    # --- main bars: their top stops at the top cover, or laps up past the top ---
    z1 = L + o["lap"] if o["lap"] > 0 else L - ct
    z0 = -o["bottom_ext"] if o["bottom_ext"] > 0 else cb
    pts, angs = [], []
    if o["shape"] == u"round":
        r = o["D"] / 2.0 - c - td - dm / 2.0
        n = max(int(o["main_count"]), 1)
        step = 360.0 / n
        start = -90.0 if o["arrangement"] == u"axes" else 90.0 - step / 2.0
        for k in range(n):
            a = math.radians(start - k * step)
            pts.append((r * math.cos(a), r * math.sin(a)))
            angs.append(a)
        out["main_r"] = r
    else:
        nb, nh = max(int(o["bars_b"]), 2), max(int(o["bars_h"]), 2)
        xa = o["b"] / 2.0 - c - td - dm / 2.0
        ya = o["h"] / 2.0 - c - td - dm / 2.0
        xs = [-xa + 2 * xa * i / float(nb - 1) for i in range(nb)]
        ys = [-ya + 2 * ya * i / float(nh - 1) for i in range(nh)]
        ring = ([(x, -ya) for x in xs] + [(xa, y) for y in ys[1:]]
                + [(x, ya) for x in reversed(xs[:-1])] + [(-xa, y) for y in reversed(ys[1:-1])])
        for x, y in ring:
            pts.append((x, y))
            angs.append(math.atan2(y, x))
        out["xs"], out["ys"] = xs, ys
    out["main"] = {"pts": pts, "ang": angs, "z0": z0, "z1": z1,
                   "hook": o["hook_dir"] if o["main_hook"] else None}

    # --- starter bars beside the main bars, an L with its foot in the footing ---
    out["starters"] = None
    if o.get("starters") and o["shape"] == u"rect":
        sd = o["starter_dm"]
        st = []
        for (x, y), a in zip(pts, angs):
            # beside the main bar, toward the middle, touching it
            ux, uy = -x, -y
            ln = math.hypot(ux, uy) or 1.0
            gap = dm / 2.0 + sd / 2.0
            sx, sy = x + ux / ln * gap, y + uy / ln * gap
            # the foot runs square to the face the bar is on, out or in
            if abs(abs(x) - out["xs"][-1]) < 1e-6 and abs(abs(y) - out["ys"][-1]) > 1e-6:
                fx, fy = (1.0 if x > 0 else -1.0), 0.0
            else:
                fx, fy = 0.0, (1.0 if y > 0 else -1.0)
            if o.get("starter_dir") == u"in":
                fx, fy = -fx, -fy
            st.append((sx, sy, fx, fy))
        out["starters"] = {"pts": st, "z_top": o["starter_lap"], "z_foot": -o["starter_embed"],
                           "foot": o["starter_foot"]}

    # --- ties ---
    sets = tie_levels(o)
    ties = {"kind": o["tie_kind"], "sets": sets, "zs": _zs(sets)}
    if o["tie_kind"] in (u"spiral", u"circle"):
        ties["r"] = o["D"] / 2.0 - c - td / 2.0
    if o["tie_kind"] == u"spiral":
        z_a, z_b = o["tie_first"], L - o["tie_last"]
        ties["sets"] = [(z_a, z_b, o["zone_mid_sp"], True)]
        ties["pitch"] = o["zone_mid_sp"]
        ties["turns"] = o["turns"]
        n = _gaps(z_b - z_a, o["zone_mid_sp"])
        ties["zs"] = [z_a + (z_b - z_a) * i / float(max(n, 1)) for i in range(n + 1)]
    if o["tie_kind"] == u"rect":
        xa = o["b"] / 2.0 - c - td / 2.0
        ya = o["h"] / 2.0 - c - td / 2.0
        boxes = [(-xa, -ya, xa, ya)]
        if o.get("inner_ties"):
            xs, ys = out["xs"], out["ys"]
            wrap = dm / 2.0 + td / 2.0
            if len(xs) >= 4:
                boxes.append((xs[1] - wrap, -ya, xs[-2] + wrap, ya))
            if len(ys) >= 4:
                boxes.append((-xa, ys[1] - wrap, xa, ys[-2] + wrap))
        ties["boxes"] = boxes
    out["ties"] = ties
    return out


def bar_length(plan, o):
    """Rough bar lengths (m) for the Calculations card: (main, ties, starters).
    Bends are taken as square corners, hooks at their leg length."""
    m = plan["main"]
    hook = (o["hook_leg"] or 12 * o["main_d"]) if m["hook"] else 0.0
    main = len(m["pts"]) * (m["z1"] - m["z0"] + hook)
    t = plan["ties"]
    if t["kind"] == u"spiral":
        z_a, z_b, pitch, _f = t["sets"][0]
        turns = (z_b - z_a) / max(pitch, 1.0) + 2 * t["turns"]
        ties = turns * math.hypot(2 * math.pi * t["r"], pitch)
    elif t["kind"] == u"circle":
        ties = len(t["zs"]) * (2 * math.pi * t["r"] + 12 * o["tie_d"])
    else:
        per = sum(2 * (x1 - x0) + 2 * (y1 - y0) + 2 * 10 * o["tie_d"] for x0, y0, x1, y1 in t["boxes"])
        ties = len(t["zs"]) * per
    st = 0.0
    if plan["starters"]:
        s = plan["starters"]
        st = len(s["pts"]) * (s["z_top"] - s["z_foot"] + s["foot"])
    return main / 1000.0, ties / 1000.0, st / 1000.0


# ── REVIT: reading the column ───────────────────────────────────────────────

def _solids(element):
    opts = DB.Options()
    opts.ComputeReferences = False
    opts.DetailLevel = DB.ViewDetailLevel.Fine
    out = []

    def walk(geo):
        for g in geo:
            if isinstance(g, DB.Solid) and g.Volume > 1e-9:
                out.append(g)
            elif isinstance(g, DB.GeometryInstance):
                walk(g.GetInstanceGeometry())
    walk(element.get_Geometry(opts))
    return out


def from_host(doc, host):
    """Read the frame of a vertical concrete column (round or rectangular).
    Raises ValueError for anything else."""
    if not isinstance(host, DB.FamilyInstance):
        raise ValueError(u"Pick a concrete column or pile.")
    data = RebarHostData.GetRebarHostData(host)
    if data is None or not data.IsValidHost():
        raise ValueError(u"That element cannot host rebar (is its material concrete?).")
    t = host.GetTotalTransform()
    ux, uy = t.BasisX, t.BasisY
    if abs(t.BasisZ.Z) < 0.999:
        raise ValueError(u"That column is slanted; only vertical columns are supported.")
    solids = _solids(host)
    if not solids:
        raise ValueError(u"That column has no solid geometry to read its size from.")
    radius = None
    pts = []
    for s in solids:
        for f in s.Faces:
            if isinstance(f, DB.CylindricalFace) and abs(f.Axis.Normalize().Z) > 0.999:
                r = f.get_Radius(0).GetLength()
                radius = max(radius or 0.0, r)
                centre = f.Origin
        for e in s.Edges:
            for p in e.Tessellate():
                pts.append(p)
    zs = [p.Z for p in pts]
    z0, z1 = min(zs), max(zs)
    o = t.Origin
    lx = [(p - o).DotProduct(DB.XYZ(ux.X, ux.Y, 0).Normalize()) for p in pts]
    ly = [(p - o).DotProduct(DB.XYZ(uy.X, uy.Y, 0).Normalize()) for p in pts]
    if radius is not None and abs((max(lx) - min(lx)) - 2 * radius) < 0.01:
        shape = u"round"
        size = (2 * radius * MM, 2 * radius * MM)
        cx, cy = centre.X, centre.Y
    else:
        shape = u"rect"
        size = ((max(lx) - min(lx)) * MM, (max(ly) - min(ly)) * MM)
        mid = (o + DB.XYZ(ux.X, ux.Y, 0).Normalize() * ((max(lx) + min(lx)) / 2.0)
               + DB.XYZ(uy.X, uy.Y, 0).Normalize() * ((max(ly) + min(ly)) / 2.0))
        cx, cy = mid.X, mid.Y
    return Frame(DB.XYZ(cx, cy, z0), ux, uy, z0, shape, size, (z1 - z0) * MM)


# ── REVIT: making the bars ──────────────────────────────────────────────────

def _cover_type(doc, mm):
    for c in DB.FilteredElementCollector(doc).OfClass(RebarCoverType):
        if abs(c.CoverDistance * MM - mm) < 0.01:
            return c
    return RebarCoverType.Create(doc, u"Column Rebar Cover {0:g}".format(round(mm, 1)), _ft(mm))


def _set_cover(doc, element, bip, mm):
    p = element.get_Parameter(bip)
    if p is not None and not p.IsReadOnly:
        cover = _cover_type(doc, mm)
        if p.AsElementId() != cover.Id:
            p.Set(cover.Id)


def set_covers(doc, host, o):
    """Give the column the covers the bars are set out to: Revit keeps bars
    out of a host's cover zone and moves the ones inside it."""
    _set_cover(doc, host, DB.BuiltInParameter.CLEAR_COVER_OTHER, o["cover_side"])
    _set_cover(doc, host, DB.BuiltInParameter.CLEAR_COVER_TOP, min(o["cover_top"], o["cover_side"]))
    _set_cover(doc, host, DB.BuiltInParameter.CLEAR_COVER_BOTTOM, o["cover_bottom"])


def _hook_length(rebar, mm):
    """Override Hook Lengths on, every hook's straight leg set to `mm`."""
    rebar.EnableHookLengthOverride(True)
    for ids in rebar.GetOverridableHookParameters():
        for pid in ids:
            for p in rebar.Parameters:
                if p.Id == pid and not p.IsReadOnly:
                    p.Set(_ft(mm))


def _hook_tail(rebar, start):
    """The free end of the hook: of the bar's two ends, the one away from
    `start` (Revit may hand the centreline back in either direction)."""
    curves = list(rebar.GetCenterlineCurves(False, False, False,
                                            MultiplanarOption.IncludeOnlyPlanarCurves, 0))
    ends = [curves[0].GetEndPoint(0), curves[-1].GetEndPoint(1)]
    return max(ends, key=lambda p: p.DistanceTo(start))


def _rigid(frame, x, y, ang):
    """Model transform: turn by `ang` (radians) about the column's axis, then
    move by (x, y) mm in the column's frame."""
    axis_pt = frame.p(0.0, 0.0, 0.0)
    rot = DB.Transform.CreateRotationAtPoint(DB.XYZ.BasisZ, ang, axis_pt)
    move = DB.Transform.CreateTranslation(frame.ux * _ft(x) + frame.uy * _ft(y))
    return move.Multiply(rot)


def _place_set(doc, rebar, places, normal):
    """Lay a single bar out as a set of len(places) and move each bar to its
    place: places[k] is the model transform from bar 0 to bar k. Falls back
    to separate bars if Revit will not move bars in the set. Returns the bars."""
    n = len(places)
    if n < 2:
        return [rebar]
    try:
        acc = rebar.GetShapeDrivenAccessor()
        acc.SetLayoutAsFixedNumber(n, _ft(ARRAY_STEP * (n - 1)), True, True, True)
        for k in range(1, n):
            back = DB.Transform.CreateTranslation(normal.Normalize() * (-_ft(ARRAY_STEP) * k))
            rebar.MoveBarInSet(k, places[k].Multiply(back))
        return [rebar]
    except Exception:
        try:
            rebar.GetShapeDrivenAccessor().SetLayoutAsSingle()
        except Exception:
            pass
        out = [rebar]
        for k in range(1, n):
            ids = DB.ElementTransformUtils.CopyElement(doc, rebar.Id, DB.XYZ.Zero)
            bar = doc.GetElement(list(ids)[0])
            DB.ElementTransformUtils.MoveElement(doc, bar.Id, DB.XYZ.Zero)
            _transform_element(doc, bar, places[k])
            out.append(bar)
        return out


def _transform_element(doc, element, t):
    """Move an element by a rigid transform (a turn about Z, then a move)."""
    ang = math.atan2(t.BasisX.Y, t.BasisX.X)
    if abs(ang) > 1e-9:
        axis = DB.Line.CreateBound(DB.XYZ.Zero, DB.XYZ.BasisZ)
        DB.ElementTransformUtils.RotateElement(doc, element.Id, axis, ang)
    if t.Origin.GetLength() > 1e-9:
        DB.ElementTransformUtils.MoveElement(doc, element.Id, t.Origin)


def _vertical_set(doc, host, frame, bar_type, pts, angs, z0, z1, hook_type, hook_dir, hook_leg):
    """The main bars: one straight bar (hooked at the bottom) made at the
    first position, then copied round to the others as one set."""
    x, y = pts[0]
    a = angs[0]
    radial = frame.v(math.cos(a), math.sin(a))
    normal = DB.XYZ.BasisZ.CrossProduct(radial)          # the hook's plane holds the radial
    curve = DB.Line.CreateBound(frame.p(x, y, z1), frame.p(x, y, z0))
    want = -radial if hook_dir == u"in" else radial

    straight = (_shape(doc, lambda s: _name(s) == u"S" and _is_straight(s))
                or _shape(doc, _is_straight))

    def make(orient):
        # the straight shape (S) with the hook added, as the bars are usually
        # scheduled; Revit would otherwise pick a shape with the hook built in
        if straight is not None:
            try:
                bar = Rebar.CreateFromCurvesAndShape(doc, straight, bar_type, None, hook_type, host,
                                                     normal, _curves([curve]),
                                                     RebarHookOrientation.Left, orient)
                if bar is not None:         # Revit hands back nothing when the shape will not take it
                    return bar
            except Exception:
                pass
        return Rebar.CreateFromCurves(doc, RebarStyle.Standard, bar_type, None, hook_type, host,
                                      normal, _curves([curve]), RebarHookOrientation.Left,
                                      orient, True, True)
    bar = make(RebarHookOrientation.Left)
    if hook_type is not None:
        doc.Regenerate()
        tail = _hook_tail(bar, frame.p(x, y, z1))
        foot = frame.p(x, y, z0)
        if (tail - foot).DotProduct(want) < 0:
            doc.Delete(bar.Id)
            bar = make(RebarHookOrientation.Right)
        # Revit may match a shape with the hook built in and give it the shape's
        # own hook type: put the chosen one back on whichever end is hooked
        for end in (0, 1):
            current = bar.GetHookTypeId(end)
            if current != DB.ElementId.InvalidElementId and current != hook_type.Id:
                bar.SetHookTypeId(end, hook_type.Id)
        if hook_leg > 0:
            _hook_length(bar, hook_leg)
    # every other bar: bar 0 turned about the axis to its own angle (so its hook
    # still points in or out), then moved to its point
    places = [DB.Transform.Identity]
    p0 = frame.p(x, y, 0.0)
    for (xk, yk), ak in zip(pts[1:], angs[1:]):
        turn = DB.Transform.CreateRotationAtPoint(DB.XYZ.BasisZ, ak - a, frame.p(0.0, 0.0, 0.0))
        moved = turn.OfPoint(p0)
        target = frame.p(xk, yk, 0.0)
        places.append(DB.Transform.CreateTranslation(target - moved).Multiply(turn))
    return _place_set(doc, bar, places, normal)


def _starter_set(doc, host, frame, bar_type, st):
    """L starter bars: the foot in the footing, the leg up beside a main bar."""
    sx, sy, fx, fy = st["pts"][0]
    foot_dir = frame.v(fx, fy)
    normal = DB.XYZ.BasisZ.CrossProduct(foot_dir)
    top = frame.p(sx, sy, st["z_top"])
    knee = frame.p(sx, sy, st["z_foot"])
    toe = knee + foot_dir * _ft(st["foot"])
    bar = Rebar.CreateFromCurves(doc, RebarStyle.Standard, bar_type, None, None, host, normal,
                                 _curves([DB.Line.CreateBound(top, knee), DB.Line.CreateBound(knee, toe)]),
                                 RebarHookOrientation.Left, RebarHookOrientation.Left, True, True)
    a0 = math.atan2(fy, fx)
    places = [DB.Transform.Identity]
    p0 = frame.p(sx, sy, 0.0)
    for x, y, dx, dy in st["pts"][1:]:
        turn = DB.Transform.CreateRotationAtPoint(DB.XYZ.BasisZ, math.atan2(dy, dx) - a0,
                                                  frame.p(0.0, 0.0, 0.0))
        moved = turn.OfPoint(p0)
        places.append(DB.Transform.CreateTranslation(frame.p(x, y, 0.0) - moved).Multiply(turn))
    return _place_set(doc, bar, places, normal)


def _spread_up(rebar, za, zb, spacing, include_first):
    if zb - za > 1e-3:
        rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(
            _ft(spacing), _ft(zb - za), True, include_first, True)


def _rect_tie(doc, host, frame, bar_type, hook_type, box, z):
    x0, y0, x1, y1 = box
    pts = [frame.p(x0, y0, z), frame.p(x1, y0, z), frame.p(x1, y1, z), frame.p(x0, y1, z)]
    curves = [DB.Line.CreateBound(pts[i], pts[(i + 1) % 4]) for i in range(4)]
    return Rebar.CreateFromCurves(doc, RebarStyle.StirrupTie, bar_type, hook_type, hook_type, host,
                                  DB.XYZ.BasisZ, _curves(curves), RebarHookOrientation.Left,
                                  RebarHookOrientation.Left, True, True)


def _shape(doc, test):
    for s in DB.FilteredElementCollector(doc).OfClass(RebarShape):
        try:
            if test(s):
                return s
        except Exception:
            continue
    return None


def _is_straight(s):
    """A one-segment standard shape with no hooks of its own (S)."""
    d = s.GetRebarShapeDefinition()
    return (isinstance(d, DB.Structure.RebarShapeDefinitionBySegments) and d.NumberOfSegments == 1
            and s.RebarStyle == RebarStyle.Standard
            and s.GetDefaultHookAngle(0) == 0 and s.GetDefaultHookAngle(1) == 0)


def _is_spiral(s):
    d = s.GetRebarShapeDefinition()
    return (isinstance(d, RebarShapeDefinitionByArc)
            and d.Type == DB.Structure.RebarShapeDefinitionByArcType.Spiral)


def _is_circle(s):
    d = s.GetRebarShapeDefinition()
    return (isinstance(d, RebarShapeDefinitionByArc)
            and d.Type == DB.Structure.RebarShapeDefinitionByArcType.LappedCircle)


def _circle_tie(doc, host, frame, bar_type, r, z, o):
    """A circular tie (a lapped circle shape, e.g. CT) of centreline radius r."""
    shape = _shape(doc, _is_circle)
    if shape is None:
        raise ValueError(u"The model has no lapped-circle rebar shape (e.g. CT) for the ties.")
    outer = r + o["tie_dm"] / 2.0
    bar = Rebar.CreateFromRebarShape(doc, shape, bar_type, host, frame.p(-outer, -outer, z),
                                     frame.ux, frame.uy)
    bar.GetShapeDrivenAccessor().ScaleToBox(frame.p(-outer, -outer, z),
                                           frame.ux * _ft(2 * outer), frame.uy * _ft(2 * outer))
    return bar


def _spiral(doc, host, frame, bar_type, ties, o):
    shape = _shape(doc, _is_spiral)
    if shape is None:
        raise ValueError(u"The model has no spiral rebar shape (e.g. SP).")
    z_a, z_b, pitch, _f = ties["sets"][0]
    outer = ties["r"] + o["tie_dm"] / 2.0
    origin = frame.p(-outer, -outer, z_a)
    bar = Rebar.CreateFromRebarShape(doc, shape, bar_type, host, origin, frame.ux, frame.uy)
    acc = bar.GetShapeDrivenAccessor()
    _set_spiral(bar, pitch, ties["turns"])
    acc.ScaleToBoxFor3D(origin, frame.ux * _ft(2 * outer), frame.uy * _ft(2 * outer), _ft(z_b - z_a))
    _set_spiral(bar, pitch, ties["turns"])
    return bar


def _set_spiral(bar, pitch, turns):
    for bip, value in ((DB.BuiltInParameter.REBAR_SHAPE_SPIRAL_PITCH, _ft(pitch)),
                       (DB.BuiltInParameter.REBAR_SHAPE_SPIRAL_BASE_FINISHING_TURNS, float(turns)),
                       (DB.BuiltInParameter.REBAR_SHAPE_SPIRAL_TOP_FINISHING_TURNS, float(turns))):
        p = bar.get_Parameter(bip)
        if p is not None and not p.IsReadOnly:
            if p.StorageType == DB.StorageType.Integer:
                p.Set(int(round(value)))
            else:
                p.Set(value)


def numbers(opts):
    """The plain numbers layout() needs, from reinforce()'s options (Revit
    types in, diameters in mm out)."""
    o = dict(opts)
    o["main_d"] = opts["main_bar"].BarNominalDiameter * MM
    o["main_dm"] = opts["main_bar"].BarModelDiameter * MM
    o["tie_d"] = opts["tie_bar"].BarNominalDiameter * MM
    o["tie_dm"] = opts["tie_bar"].BarModelDiameter * MM
    sb = opts.get("starter_bar") or opts["main_bar"]
    o["starter_dm"] = sb.BarModelDiameter * MM
    return o


def reinforce(doc, host, frame, opts):
    """Add the bars to a column. Needs an open transaction. opts are the
    window's (mm, Revit types); the column's size comes from `frame`.
    Returns a dict of what was made."""
    o = numbers(opts)
    o["shape"] = frame.shape
    o["D"], o["b"], o["h"], o["L"] = frame.size[0], frame.size[0], frame.size[1], frame.length
    plan = layout(o)
    set_covers(doc, host, o)
    doc.Regenerate()
    made = {"main": 0, "ties": 0, "starters": 0}

    m = plan["main"]
    bars = _vertical_set(doc, host, frame, opts["main_bar"], m["pts"], m["ang"], m["z0"], m["z1"],
                         opts.get("main_hook"), o["hook_dir"], o["hook_leg"])
    made["main"] = sum(b.NumberOfBarPositions for b in bars)

    if plan["starters"]:
        bars = _starter_set(doc, host, frame, opts.get("starter_bar") or opts["main_bar"],
                            plan["starters"])
        made["starters"] = sum(b.NumberOfBarPositions for b in bars)

    t = plan["ties"]
    if t["kind"] == u"spiral":
        _spiral(doc, host, frame, opts["tie_bar"], t, o)
        made["ties"] = 1
    else:
        later = []
        for za, zb, sp, first in t["sets"]:
            if t["kind"] == u"circle":
                ties = [_circle_tie(doc, host, frame, opts["tie_bar"], t["r"], za, o)]
            else:
                ties = [_rect_tie(doc, host, frame, opts["tie_bar"], opts.get("tie_hook"), box, za)
                        for box in t["boxes"]]
            for tie in ties:
                _spread_up(tie, za, zb, sp, first)
                later.append((tie, first))
        # Revit puts a left-out first tie back when it regenerates (the zone below
        # already has a tie there): leave it out again once everything is made
        doc.Regenerate()
        for tie, first in later:
            if not first and tie.NumberOfBarPositions > 1 and tie.IncludeFirstBar:
                tie.IncludeFirstBar = False
        doc.Regenerate()
        made["ties"] = sum(_count(tie) for tie, _f in later)
    return made


def _count(rebar):
    """Bars in a set that are there (an excluded first or last is not)."""
    return sum(1 for k in range(rebar.NumberOfBarPositions) if rebar.DoesBarExistAtPosition(k))


def remove_bars(doc, host):
    """Delete every bar hosted on the column. Needs an open transaction.
    Returns how many sets went."""
    data = RebarHostData.GetRebarHostData(host)
    gone = 0
    if data is not None:
        for bar in list(data.GetRebarsInHost()):
            doc.Delete(bar.Id)
            gone += 1
    return gone


# ── REVIT: the column itself ────────────────────────────────────────────────

# yes/no type parameters some concrete column families switch their shape with
ROUND_FLAGS = (u"Round Column", u"Round", u"Circular")
RECT_FLAGS = (u"Square Column", u"Rectangular Column", u"Square", u"Rectangular")
# size parameters: a round type's diameter, a rectangular type's width and depth
ROUND_SIZES = (u"Diameter", u"D", u"d", u"a", u"b")
WIDTH_SIZES = (u"b", u"Width", u"B")
DEPTH_SIZES = (u"h", u"a", u"Depth", u"H")


def _type_param(symbol, names):
    for n in names:
        p = symbol.LookupParameter(n)
        if p is not None:
            return p
    return None


def column_types(doc):
    """Concrete column and pile types the tool can place: structural columns
    (and foundations) of families whose material is concrete."""
    out = []
    for bic in (DB.BuiltInCategory.OST_StructuralColumns, DB.BuiltInCategory.OST_StructuralFoundation):
        for s in DB.FilteredElementCollector(doc).OfCategory(bic).WhereElementIsElementType():
            if not isinstance(s, DB.FamilySymbol):
                continue
            try:
                concrete = s.Family.StructuralMaterialType == DB.Structure.StructuralMaterialType.Concrete
            except Exception:
                concrete = False
            if concrete and bic == DB.BuiltInCategory.OST_StructuralColumns:
                out.append(s)
    return sorted(out, key=lambda s: (s.FamilyName, _name(s)))


def type_shape(symbol):
    """'round', 'rect' or None (cannot tell) for a column type."""
    for names, shape in ((ROUND_FLAGS, u"round"), (RECT_FLAGS, u"rect")):
        p = _type_param(symbol, names)
        if p is not None and p.StorageType == DB.StorageType.Integer and p.AsInteger() == 1:
            return shape
    low = _name(symbol).lower()
    if u"round" in low or u"circ" in low or u"pile" in low:
        return u"round"
    if u"square" in low or u"rect" in low or u" x " in low:
        return u"rect"
    return None


def type_size(symbol):
    """(D, D) or (b, h) of a column type, mm, from its size parameters, or None."""
    shape = type_shape(symbol)
    if shape == u"round":
        p = _type_param(symbol, ROUND_SIZES)
        return (p.AsDouble() * MM,) * 2 if p is not None else None
    pb, ph = _type_param(symbol, WIDTH_SIZES), _type_param(symbol, DEPTH_SIZES)
    if pb is None or ph is None:
        return None
    return pb.AsDouble() * MM, ph.AsDouble() * MM


def type_name(shape, size):
    """The name a new type gets, as the model's own: 600 Round Column, 500 x 500 Square Column."""
    if shape == u"round":
        return u"{0:g} Round Column".format(round(size[0], 1))
    b, h = round(size[0], 1), round(size[1], 1)
    return u"{0:g} x {1:g} {2} Column".format(b, h, u"Square" if b == h else u"Rectangular")


def ensure_type(doc, symbol, shape, size):
    """A type of `symbol`'s family with this shape and size: one already in
    the model, else `symbol` duplicated and set to it. Needs an open
    transaction (only used when a new type is made)."""
    for s in DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol):
        if s.Family.Id != symbol.Family.Id or type_shape(s) != shape:
            continue
        have = type_size(s)
        if have and abs(have[0] - size[0]) < 0.5 and abs(have[1] - size[1]) < 0.5:
            return s
    name = type_name(shape, size)
    taken = set(_name(s) for s in DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol)
                if s.Family.Id == symbol.Family.Id)
    base, k = name, 2
    while name in taken:
        name = u"{0} ({1})".format(base, k)
        k += 1
    new = symbol.Duplicate(name)
    for names, value in ((ROUND_FLAGS, 1 if shape == u"round" else 0),
                         (RECT_FLAGS, 0 if shape == u"round" else 1)):
        p = _type_param(new, names)
        if p is not None and not p.IsReadOnly and p.StorageType == DB.StorageType.Integer:
            p.Set(value)
    if shape == u"round":
        for n in ROUND_SIZES:
            p = new.LookupParameter(n)
            if p is not None and not p.IsReadOnly and p.StorageType == DB.StorageType.Double:
                p.Set(_ft(size[0]))
    else:
        for names, value in ((WIDTH_SIZES, size[0]), (DEPTH_SIZES, size[1])):
            p = _type_param(new, names)
            if p is not None and not p.IsReadOnly:
                p.Set(_ft(value))
    return new


def _set_levels(host, level, top_offset, length, mark):
    """Top on `level` + top_offset, bottom `length` below it, both on the
    same level (as the model's piles are), and the Mark."""
    for bip, value in ((DB.BuiltInParameter.FAMILY_BASE_LEVEL_PARAM, level.Id),
                       (DB.BuiltInParameter.FAMILY_TOP_LEVEL_PARAM, level.Id)):
        p = host.get_Parameter(bip)
        if p is not None and not p.IsReadOnly:
            p.Set(value)
    # the top first when it rises, the base first when it drops, so the two never cross
    top = host.get_Parameter(DB.BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM)
    base = host.get_Parameter(DB.BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM)
    pairs = [(top, top_offset), (base, top_offset - length)]
    if top is not None and _ft(top_offset) < top.AsDouble():
        pairs.reverse()
    for p, mm in pairs:
        if p is not None and not p.IsReadOnly:
            p.Set(_ft(mm))
    if mark:
        p = host.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK)
        if p is not None and not p.IsReadOnly:
            p.Set(mark)


def place_column(doc, symbol, point, level, top_offset, length, mark=None, angle=0.0):
    """A new vertical column of `symbol` at `point` (plan), its top on
    `level` + top_offset and `length` long, turned by `angle` (radians).
    Needs an open transaction."""
    if not symbol.IsActive:
        symbol.Activate()
        doc.Regenerate()
    at = DB.XYZ(point.X, point.Y, level.Elevation)
    host = doc.Create.NewFamilyInstance(at, symbol, level, DB.Structure.StructuralType.Column)
    _set_levels(host, level, top_offset, length, mark)
    if abs(angle) > 1e-9:
        axis = DB.Line.CreateBound(at, at + DB.XYZ.BasisZ)
        DB.ElementTransformUtils.RotateElement(doc, host.Id, axis, angle)
    return host


def update_column(doc, host, symbol, level, top_offset, length, mark=None):
    """Give a column a new type, levels and Mark where it stands. Needs an
    open transaction."""
    if host.Symbol.Id != symbol.Id:
        if not symbol.IsActive:
            symbol.Activate()
        host.Symbol = symbol
    _set_levels(host, level, top_offset, length, mark)


def column_levels(doc, host):
    """(top level, top offset mm) of a column, or (None, None)."""
    p = host.get_Parameter(DB.BuiltInParameter.FAMILY_TOP_LEVEL_PARAM)
    o = host.get_Parameter(DB.BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM)
    if p is None or o is None:
        return None, None
    return doc.GetElement(p.AsElementId()), o.AsDouble() * MM


# ── REVIT: reading bars already in a column ─────────────────────────────────

def _g(mm):
    return u"{0:g}".format(round(mm))


def _existing(rebar):
    return [k for k in range(rebar.NumberOfBarPositions) if rebar.DoesBarExistAtPosition(k)]


def _bar_curves(rebar, k, hooks=False):
    """Bar k of a set where it really is (moved bars included), in the model."""
    return list(rebar.GetTransformedCenterlineCurves(False, not hooks, True,
                                                    MultiplanarOption.IncludeOnlyPlanarCurves, k))


def _local(frame, p):
    """(x, y, z) mm of a model point in the column's frame."""
    rel = p - DB.XYZ(frame.centre.X, frame.centre.Y, frame.base_z)
    return rel.DotProduct(frame.ux) * MM, rel.DotProduct(frame.uy) * MM, rel.Z * MM


def _hook_name(doc, rebar):
    for end in (1, 0):
        hook = doc.GetElement(rebar.GetHookTypeId(end))
        if hook is not None:
            return _name(hook)
    return u"(no hook)"


def read_existing(doc, host, frame):
    """Settings read off the bars already in a column, as the window's
    settings strings (only what could be read), plus "layout": the layout
    that fits them. Empty when it has none."""
    data = RebarHostData.GetRebarHostData(host)
    bars = list(data.GetRebarsInHost()) if data is not None else []
    out = {}
    verticals, starters, ties = [], [], []
    for b in bars:
        ks = _existing(b)
        if not ks:
            continue
        try:
            curves = _bar_curves(b, ks[0])
        except Exception:
            continue
        shape = doc.GetElement(b.GetShapeId())
        if shape is not None and shape.RebarStyle == RebarStyle.StirrupTie:
            ties.append((b, ks, shape))
        elif len(curves) == 1 and isinstance(curves[0], DB.Line) and abs(curves[0].Direction.Z) > 0.99:
            verticals.append((b, ks))
        elif len(curves) == 2 and all(isinstance(c, DB.Line) for c in curves):
            starters.append((b, ks))
    if not verticals and not ties:
        return out
    L = frame.length

    # --- main bars: the bar type with the most vertical bars ---
    if verticals:
        # grouped by bar type and where they stop: the biggest group is the
        # main bars; a group from below the base to part way up is starters
        # (their foot and crank are hooks, so they read as straight here)
        groups = {}
        for b, ks in verticals:
            c = _bar_curves(b, ks[0])[0]
            zs_ = sorted(_local(frame, c.GetEndPoint(i))[2] for i in (0, 1))
            key = (eid(b.GetTypeId()), int(round(zs_[0] / 50.0)), int(round(zs_[1] / 50.0)))
            groups.setdefault(key, []).append((b, ks))
        main_key = max(groups, key=lambda k: sum(len(ks) for _b, ks in groups[k]))
        mains = groups[main_key]
        for key, members in groups.items():
            if key != main_key and key[1] * 50.0 < -1 and key[2] * 50.0 < 0.8 * L:
                starters += members
        main_type = mains[0][0].GetTypeId()
        out["main_bar"] = _name(doc.GetElement(main_type))
        pts, zs = [], []
        for b, ks in mains:
            for k in ks:
                c = _bar_curves(b, k)[0]
                x, y, z0 = _local(frame, c.GetEndPoint(0))
                _x, _y, z1 = _local(frame, c.GetEndPoint(1))
                pts.append((x, y))
                zs += [z0, z1]
        if frame.shape == u"round":
            out["main_count"] = unicode(len(pts))
            step = 360.0 / max(len(pts), 1)
            on_axis = any(min(abs(math.degrees(math.atan2(y, x)) % 90.0),
                              90.0 - abs(math.degrees(math.atan2(y, x)) % 90.0)) < min(2.0, step / 4.0)
                          for x, y in pts)
            out["arrangement"] = u"axes" if on_axis else u"between"
        else:
            out["bars_b"] = unicode(len(set(int(round(x / 5.0)) for x, _y in pts)))
            out["bars_h"] = unicode(len(set(int(round(y / 5.0)) for _x, y in pts)))
        top, bottom = max(zs), min(zs)
        if top > L + 1:
            out["lap"] = _g(top - L)
        else:
            out["lap"], out["cover_top"] = u"0", _g(L - top)
        if bottom < -1:
            out["bottom_ext"] = _g(-bottom)
        else:
            out["bottom_ext"], out["cover_bottom"] = u"0", _g(bottom)
        b, ks = mains[0]
        out["main_hook"] = _hook_name(doc, b)
        if out["main_hook"] != u"(no hook)":
            # which way the hook's tail runs from the bar: toward the middle or out
            straight = _bar_curves(b, ks[0])[0]
            top_pt = max((straight.GetEndPoint(0), straight.GetEndPoint(1)), key=lambda p: p.Z)
            hooked = _bar_curves(b, ks[0], True)
            tail = max((hooked[0].GetEndPoint(0), hooked[-1].GetEndPoint(1)),
                       key=lambda p: p.DistanceTo(top_pt))
            tx, ty, _z = _local(frame, tail)
            bx, by, _z = _local(frame, top_pt)
            out["hook_dir"] = u"in" if tx * bx + ty * by < bx * bx + by * by else u"out"
            out["hook_leg"] = u"0"

    # --- starters ---
    out["starters"] = bool(starters)
    if starters:
        b, ks = starters[0]
        c = _bar_curves(b, ks[0])
        vert = [x for x in c if abs(x.Direction.Z) > 0.99]
        flat = [x for x in c if abs(x.Direction.Z) < 0.01]
        out["starter_bar"] = _name(doc.GetElement(b.GetTypeId()))
        if vert and not flat:
            # the foot is a hook: read the leg's ends, and the foot off the hooked bar
            zs = [_local(frame, vert[0].GetEndPoint(i))[2] for i in (0, 1)]
            out["starter_lap"], out["starter_embed"] = _g(max(zs)), _g(-min(zs))
            hooked = [x for x in _bar_curves(b, ks[0], True) if isinstance(x, DB.Line)
                      and abs(x.Direction.Z) < 0.01]
            if hooked:
                out["starter_foot"] = _g(max(x.Length for x in hooked) * MM)
        elif vert and flat:
            zs = [_local(frame, vert[0].GetEndPoint(i))[2] for i in (0, 1)]
            out["starter_lap"], out["starter_embed"] = _g(max(zs)), _g(-min(zs))
            out["starter_foot"] = _g(flat[0].Length * MM)
            knee = vert[0].GetEndPoint(0) if vert[0].GetEndPoint(0).Z < vert[0].GetEndPoint(1).Z \
                else vert[0].GetEndPoint(1)
            far = max((flat[0].GetEndPoint(0), flat[0].GetEndPoint(1)), key=lambda p: p.DistanceTo(knee))
            kx, ky, _z = _local(frame, knee)
            fx, fy, _z = _local(frame, far)
            out["starter_dir"] = u"out" if fx * fx + fy * fy > kx * kx + ky * ky else u"in"

    # --- ties: one spiral, or sets in zones up the column ---
    if ties:
        b, ks, shape = ties[0]
        out["tie_bar"] = _name(doc.GetElement(b.GetTypeId()))
        d = shape.GetRebarShapeDefinition()
        spiral = (isinstance(d, RebarShapeDefinitionByArc)
                  and d.Type == DB.Structure.RebarShapeDefinitionByArcType.Spiral)
        circle = isinstance(d, RebarShapeDefinitionByArc) and not spiral
        tie_dm = doc.GetElement(b.GetTypeId()).BarModelDiameter * MM
        if spiral:
            out["tie_kind"], out["layout"] = u"spiral", u"pile_spiral"
            pitch = b.get_Parameter(DB.BuiltInParameter.REBAR_SHAPE_SPIRAL_PITCH)
            turns = b.get_Parameter(DB.BuiltInParameter.REBAR_SHAPE_SPIRAL_BASE_FINISHING_TURNS)
            if pitch is not None:
                sp = _g(pitch.AsDouble() * MM)
                out["zone_mid_sp"] = out["zone_bottom_sp"] = out["zone_top_sp"] = sp
            if turns is not None:
                out["turns"] = _g(turns.AsInteger() if turns.StorageType == DB.StorageType.Integer
                                  else turns.AsDouble())
            pts = [_local(frame, p) for c in _bar_curves(b, ks[0]) for p in (c.GetEndPoint(0), c.GetEndPoint(1))]
            out["tie_first"] = _g(min(p[2] for p in pts))
            out["tie_last"] = _g(L - max(p[2] for p in pts))
            r = max(math.hypot(p[0], p[1]) for p in pts)
            out["cover_side"] = _g(frame.size[0] / 2.0 - r - tie_dm / 2.0)
            out["zone_bottom_len"] = out["zone_top_len"] = u"0"
        else:
            out["tie_kind"] = u"circle" if circle else u"rect"
            out["layout"] = u"round_ties" if circle else u"rect_ties"
            if not circle:
                out["tie_hook"] = _hook_name(doc, b)
            # each set's lowest and highest tie, and its spacing
            ranges = {}
            for t, tks, _s in ties:
                z_lo = min(_local(frame, c.GetEndPoint(0))[2] for c in _bar_curves(t, tks[0]))
                z_hi = min(_local(frame, c.GetEndPoint(0))[2] for c in _bar_curves(t, tks[-1]))
                sp = t.MaxSpacing * MM if t.NumberOfBarPositions > 1 else 0.0
                key = (int(round(z_lo)), int(round(z_hi)))
                ranges.setdefault(key, []).append((t, sp))
            zones = sorted(ranges.items())
            out["inner_ties"] = any(len(v) > 1 for _k, v in zones)
            out["tie_first"] = _g(zones[0][0][0])
            out["tie_last"] = _g(L - zones[-1][0][1])
            spans = [(k[1] - k[0], max(s for _t, s in v)) for k, v in zones]
            # the middle is the widest spacing (the ends are the closer ones)
            mid = max(range(len(spans)), key=lambda i: (round(spans[i][1]), spans[i][0]))
            out["zone_mid_sp"] = _g(spans[mid][1])
            out["zone_bottom_len"], out["zone_top_len"] = u"0", u"0"
            out["zone_bottom_sp"] = out["zone_top_sp"] = out["zone_mid_sp"]
            if mid > 0:
                out["zone_bottom_len"] = _g(zones[mid][0][0] - zones[0][0][0])
                out["zone_bottom_sp"] = _g(spans[0][1])
            if mid < len(zones) - 1:
                out["zone_top_len"] = _g(zones[-1][0][1] - zones[mid][0][1])
                out["zone_top_sp"] = _g(spans[-1][1])
            # the cover, from the outermost tie
            curves = _bar_curves(ties[0][0], ties[0][1][0])
            pts = [_local(frame, c.GetEndPoint(i)) for c in curves for i in (0, 1)]
            if circle:
                r = max(math.hypot(p[0], p[1]) for p in pts)
                out["cover_side"] = _g(frame.size[0] / 2.0 - r - tie_dm / 2.0)
            else:
                xa = max(abs(p[0]) for t, tks, _s in ties
                         for c in _bar_curves(t, tks[0]) for p in [_local(frame, c.GetEndPoint(0))])
                out["cover_side"] = _g(frame.size[0] / 2.0 - xa - tie_dm / 2.0)
    elif frame.shape == u"rect":
        out["layout"] = u"rect_ties"
    return out
