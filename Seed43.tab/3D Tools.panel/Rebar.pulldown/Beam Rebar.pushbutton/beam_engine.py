# -*- coding: utf-8 -*-
"""Beam reinforcing: native Revit rebar hosted on a concrete beam.

Everything is set out in the beam's own frame, in mm:

    x  along the beam, from its start face (the concrete as cut, not the
       location line's end)
    y  across the beam, from its centreline, positive to the left looking
       from the start toward the end
    z  up, from the beam's underside

and turned into model points with the beam's start, axis and underside.

layout() is pure arithmetic (no Revit): the window draws from it and
reinforce() builds from it, so the drawing and the model cannot disagree.
Each rebar layout is a line in LAYOUTS: its key (saved in settings), the
name shown in the Rebar layout list, the kind of section it makes and the
defaults that make it, read off the foundation beams in the sample model
(02: FDN01, FDN05, FDN11) and the standard beam details.
"""
import math

from System.Collections.Generic import List
from Autodesk.Revit import DB
from Autodesk.Revit.DB.Structure import (Rebar, RebarStyle, RebarHookOrientation,
                                         RebarHostData, RebarCoverType, RebarShape,
                                         StructuralType)

__all__ = ["Frame", "LAYOUTS", "layout_names", "layout_info", "numbers", "layout",
           "stirrup_xs", "reinforce", "remove_bars", "from_host", "eid", "bar_length",
           "MM", "beam_types", "type_size", "type_name", "ensure_type", "place_beam",
           "beam_level"]

# ── CONSTANTS ───────────────────────────────────────────────────────────────

MM = 304.8

# --- Layouts ---
# (key, name in the list, kind, defaults). The kind decides the section's
# ties (see layout()); the defaults are the standard detail's settings.
LAYOUTS = [
    (u"closed", u"Closed stirrup - top, bottom and side bars", u"closed",
     {"top_count": u"3", "bottom_count": u"3", "side_count": u"1",
      "tie_hook": u"Stirrup/Tie Seismic - 135 deg.", "tie_count": u"1", "link_count": u"0"}),
    (u"links", u"Closed stirrup with centre links", u"links",
     {"top_count": u"5", "bottom_count": u"5", "side_count": u"1",
      "tie_hook": u"Stirrup/Tie Seismic - 135 deg.", "tie_count": u"1", "link_count": u"1",
      "link_hook": u"Standard - 180 deg."}),
    (u"multi", u"Overlapping stirrups - wide beam", u"multi",
     {"top_count": u"9", "bottom_count": u"9", "side_count": u"1",
      "tie_hook": u"Stirrup/Tie Seismic - 135 deg.", "tie_count": u"4", "link_count": u"0"}),
    (u"two_u", u"Two U bars lapped - top and bottom", u"two_u",
     {"top_count": u"5", "bottom_count": u"5", "side_count": u"0",
      "tie_hook": u"(no hook)", "tie_count": u"1", "link_count": u"0", "u_lap": u"200"}),
    (u"flat", u"Flat link round the bottom bars - slab thickening", u"flat",
     {"top_count": u"0", "bottom_count": u"2", "side_count": u"0",
      "tie_hook": u"Stirrup/Tie Seismic - 135 deg.", "tie_count": u"1", "link_count": u"0"}),
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


def _spread(a, b, n):
    """n values from a to b, evenly (one: a)."""
    if n <= 1:
        return [a]
    return [a + (b - a) * i / float(n - 1) for i in range(n)]


class Frame(object):
    """The beam's frame: turns (x, y, z) in mm into model points."""

    # --- construction ---
    def __init__(self, origin, ux, width, depth, length):
        self.origin = origin                    # XYZ: start face, centreline, underside
        self.ux = DB.XYZ(ux.X, ux.Y, 0).Normalize()
        self.uy = DB.XYZ.BasisZ.CrossProduct(self.ux)     # to the left
        self.width = width                      # mm
        self.depth = depth                      # mm
        self.length = length                    # mm

    # --- public methods ---
    def p(self, x, y, z):
        return self.origin + self.ux * _ft(x) + self.uy * _ft(y) + DB.XYZ.BasisZ * _ft(z)


# ── CORE LOGIC: the layout, no Revit ────────────────────────────────────────

def layout_names():
    """[(key, name)] of the layouts, in the order the list shows them."""
    return [(k, n) for k, n, _s, _d in LAYOUTS]


def layout_info(key):
    """(name, kind, defaults) of a layout; the first one when unknown."""
    for k, n, s, d in LAYOUTS:
        if k == key:
            return n, s, d
    return LAYOUTS[0][1], LAYOUTS[0][2], LAYOUTS[0][3]


def stirrup_sets(o):
    """The stirrups along the beam as (x start, x end, spacing, include
    first): an end zone at each end at its own spacing, the middle between
    them, each laid out by Revit at a maximum spacing."""
    x0 = o["tie_first"]
    x1 = o["L"] - o["tie_first"]
    if x1 <= x0:
        return [(x0, x0, o["tie_sp"], True)]
    zone = max(o["zone_len"], 0.0) if o["zone_sp"] > 0 else 0.0
    xa = min(x0 + zone, (x0 + x1) / 2.0)
    xb = max(x1 - zone, xa)
    sets = []
    if xa > x0 + 1e-6:
        sets.append((x0, xa, o["zone_sp"], True))
    if xb > xa + 1e-6:
        sets.append((xa, xb, o["tie_sp"], not sets))
    elif not sets:
        sets.append((x0, x0, o["tie_sp"], True))
    if x1 > xb + 1e-6:
        sets.append((xb, x1, o["zone_sp"], False))
    return sets


def stirrup_xs(sets):
    """Every stirrup's position along the beam, as Revit spaces them."""
    out = []
    for xa, xb, sp, first in sets:
        n = _gaps(xb - xa, sp)
        xs = [xa + (xb - xa) * i / float(n) for i in range(n + 1)] if n else [xa]
        out += xs if first else xs[1:]
    return out


def _row(n, half):
    """n bar positions across, from -half to half (one bar: the middle)."""
    n = int(n)
    if n <= 0:
        return []
    if n == 1:
        return [0.0]
    return _spread(-half, half, n)


def layout(o):
    """Where every bar goes, in the beam's frame (mm). See numbers() for o.

    Returns:
        kind, W, H, L        the section kind and the beam
        top, bottom          rows: {"ys": [y], "z": z}
        side                 {"zs": [z], "y": y} on each face (the right face
                             is -y), or None
        ties                 [{"pts": [(y, z)], "closed": bool, "role": "tie" /
                             "link" / "u"}] in the section, centrelines
        sets, xs             the stirrups along the beam (see stirrup_sets)
        x0, x1               where the longitudinal bars start and end
    """
    W, H, L = o["W"], o["H"], o["L"]
    c, cb, ct = o["cover_side"], o["cover_bottom"], o["cover_top"]
    dm, td = o["main_dm"], o["tie_dm"]
    kind = o["kind"]
    out = {"kind": kind, "W": W, "H": H, "L": L}

    # the outer tie's centreline
    ty0, ty1 = -W / 2.0 + c + td / 2.0, W / 2.0 - c - td / 2.0
    tz0, tz1 = cb + td / 2.0, H - ct - td / 2.0
    # the bars inside it
    half = W / 2.0 - c - td - dm / 2.0
    zb = cb + td + dm / 2.0
    zt = H - ct - td - dm / 2.0
    if kind == u"two_u":
        # the top U sits inside the bottom one: its bars one tie further in
        half_top = half - td
    else:
        half_top = half
    out["top"] = {"ys": _row(o["top_count"], half_top), "z": zt}
    out["bottom"] = {"ys": _row(o["bottom_count"], half), "z": zb}
    ns = int(o["side_count"])
    out["side"] = None
    if ns > 0 and kind != u"flat":
        zs = _spread(zb, zt, ns + 2)[1:-1]
        out["side"] = {"zs": zs, "y": half}

    wrap = dm / 2.0 + td / 2.0              # bar centre to the tie's centreline
    ties = []
    if kind in (u"closed", u"links"):
        ties.append({"pts": [(ty0, tz1), (ty0, tz0), (ty1, tz0), (ty1, tz1)], "closed": True,
                     "role": u"tie"})
    if kind == u"links":
        ys = out["bottom"]["ys"]
        inner = ys[1:-1] if len(ys) > 2 else []
        nl = min(int(o["link_count"]), len(inner))
        if nl > 0:
            # the inner bars nearest to evenly spread, the middle one first
            picks = []
            for k in range(nl):
                want = -half + 2 * half * (k + 1) / float(nl + 1)
                best = min((y for y in inner if y not in picks), key=lambda y: abs(y - want))
                picks.append(best)
            # the leg one hook radius beside the bars, so each 180 degree
            # hook turns round its bar, centred on it
            r = o["link_bend"] / 2.0 + o["link_dm"] / 2.0
            # Revit takes a hooked bar's end as the outside of its hook: the
            # leg's ends one radius (and half the bar) past the bars' centres
            reach = r + o["link_dm"] / 2.0
            for y in sorted(picks):
                yl = y - r
                ties.append({"pts": [(yl, zt + reach), (yl, zb - reach)], "closed": False,
                             "role": u"link", "bar_y": y, "r": r, "zs": (zt, zb)})
    if kind == u"multi":
        ys = out["bottom"]["ys"]
        n = max(int(o["tie_count"]), 1)
        gaps = max(len(ys) - 1, 1)
        n = min(n, gaps)
        cuts = [int(round(gaps * k / float(n))) for k in range(n + 1)]
        for k in range(n):
            a, b = ys[cuts[k]], ys[cuts[k + 1]]
            y0 = ty0 if k == 0 else a - wrap
            y1 = ty1 if k == n - 1 else b + wrap
            ties.append({"pts": [(y0, tz1), (y0, tz0), (y1, tz0), (y1, tz1)], "closed": True,
                         "role": u"tie"})
    if kind == u"two_u":
        lap = max(o["u_lap"], 0.0)
        mid = (tz0 + tz1) / 2.0
        # bottom U open upward, its legs up past the middle by half the lap
        ties.append({"pts": [(ty0, mid + lap / 2.0), (ty0, tz0), (ty1, tz0), (ty1, mid + lap / 2.0)],
                     "closed": False, "role": u"u"})
        # top U open downward, one tie inside the bottom one
        uy0, uy1 = ty0 + td, ty1 - td
        ties.append({"pts": [(uy0, mid - lap / 2.0), (uy0, tz1), (uy1, tz1), (uy1, mid - lap / 2.0)],
                     "closed": False, "role": u"u"})
    if kind == u"flat":
        # a flat link round the bottom row, as the thickening detail: from
        # the top middle out to the left end bar, round it, along under the
        # bars, round the right end bar and back along the top, its two ends
        # leaving a gap at the middle
        ys = out["bottom"]["ys"] or [0.0]
        # each end a half circle no tighter than the bar bends, its outside
        # snug on the end bar; the bars sit up off the bottom cover to suit
        R = max(wrap, o["tie_bend"] / 2.0 + td / 2.0)
        zc = cb + td / 2.0 + R
        out["bottom"]["z"] = zc
        ya, yb = min(ys) - wrap + R, max(ys) + wrap - R
        if yb < ya:
            ya = yb = (ya + yb) / 2.0
        gap = min(60.0, 0.2 * max(yb - ya, 1.0)) / 2.0
        mid = (ya + yb) / 2.0
        z0, z1 = zc - R, zc + R
        segs = [(u"line", (mid - gap, z1), (ya, z1)),
                (u"arc", (ya, z1), (ya, z0), (ya - R, zc)),
                (u"line", (ya, z0), (yb, z0)),
                (u"arc", (yb, z0), (yb, z1), (yb + R, zc)),
                (u"line", (yb, z1), (mid + gap, z1))]
        if yb - ya < 1.0:
            segs = [s for s in segs if s[0] == u"arc"]
        pts = [(mid - gap, z1)]
        for a in range(1, 9):
            t = math.pi / 2.0 + math.pi * a / 8.0
            pts.append((ya + R * math.cos(t), zc + R * math.sin(t)))
        for a in range(0, 9):
            t = -math.pi / 2.0 + math.pi * a / 8.0
            pts.append((yb + R * math.cos(t), zc + R * math.sin(t)))
        pts.append((mid + gap, z1))
        ties.append({"pts": pts, "segs": segs, "closed": False, "role": u"loop"})
    out["ties"] = ties

    sets = stirrup_sets(o)
    out["sets"] = sets
    out["xs"] = stirrup_xs(sets)
    # the longitudinal bars stop the end cover in from each end (below 0:
    # they run on past it, into a support or to lap)
    out["x0"] = o["end_cover"]
    out["x1"] = L - o["end_cover"]
    return out


def _poly_length(pts, closed):
    run = sum(math.hypot(b[0] - a[0], b[1] - a[1]) for a, b in zip(pts, pts[1:]))
    if closed:
        run += math.hypot(pts[0][0] - pts[-1][0], pts[0][1] - pts[-1][1])
    return run


def bar_length(plan, o):
    """Rough bar lengths (m) for the Calculations card: (longitudinal,
    stirrups). Bends are taken as square corners, hooks at 10 d each."""
    n_long = len(plan["top"]["ys"]) + len(plan["bottom"]["ys"])
    if plan["side"]:
        n_long += 2 * len(plan["side"]["zs"])
    long_m = n_long * (plan["x1"] - plan["x0"])
    per = 0.0
    for t in plan["ties"]:
        hooks = 2 * 10 * o["tie_d"] if t["role"] in (u"tie", u"link") else 0.0
        if t["role"] == u"link":
            hooks += 2 * math.pi * t["r"]
        per += _poly_length(t["pts"], t["closed"]) + hooks
    ties_m = per * len(plan["xs"])
    return long_m / 1000.0, ties_m / 1000.0


# ── REVIT: reading the beam ─────────────────────────────────────────────────

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
    """Read the frame of a straight, level concrete beam: its concrete as
    cut (joins at the ends taken off). Raises ValueError for anything else."""
    if not isinstance(host, DB.FamilyInstance):
        raise ValueError(u"Pick a concrete beam.")
    data = RebarHostData.GetRebarHostData(host)
    if data is None or not data.IsValidHost():
        raise ValueError(u"That element cannot host rebar (is its material concrete?).")
    loc = getattr(host, "Location", None)
    curve = getattr(loc, "Curve", None)
    if not isinstance(curve, DB.Line):
        raise ValueError(u"Only straight beams are supported.")
    a, b = curve.GetEndPoint(0), curve.GetEndPoint(1)
    if abs(a.Z - b.Z) > _ft(1.0):
        raise ValueError(u"That beam slopes; only level beams are supported.")
    ux = DB.XYZ(b.X - a.X, b.Y - a.Y, 0).Normalize()
    uy = DB.XYZ.BasisZ.CrossProduct(ux)
    pts = []
    for s in _solids(host):
        for e in s.Edges:
            for p in e.Tessellate():
                pts.append(p)
    if not pts:
        raise ValueError(u"That beam has no solid geometry to read its size from.")
    xs = [(p - a).DotProduct(ux) for p in pts]
    ys = [(p - a).DotProduct(uy) for p in pts]
    zs = [p.Z for p in pts]
    x0, x1 = min(xs), max(xs)
    y0, y1 = min(ys), max(ys)
    origin = DB.XYZ(a.X, a.Y, min(zs)) + ux * x0 + uy * ((y0 + y1) / 2.0)
    return Frame(origin, ux, (y1 - y0) * MM, (max(zs) - min(zs)) * MM, (x1 - x0) * MM)


# ── REVIT: making the bars ──────────────────────────────────────────────────

def _cover_type(doc, mm):
    for c in DB.FilteredElementCollector(doc).OfClass(RebarCoverType):
        if abs(c.CoverDistance * MM - mm) < 0.01:
            return c
    return RebarCoverType.Create(doc, u"Beam Rebar Cover {0:g}".format(round(mm, 1)), _ft(mm))


def _set_cover(doc, element, bip, mm):
    p = element.get_Parameter(bip)
    if p is not None and not p.IsReadOnly:
        cover = _cover_type(doc, mm)
        if p.AsElementId() != cover.Id:
            p.Set(cover.Id)


def set_covers(doc, host, o):
    """Give the beam the covers the bars are set out to: Revit keeps bars
    out of a host's cover zone and moves the ones inside it. The ends take
    the least, so bars running on past them are not pushed back."""
    _set_cover(doc, host, DB.BuiltInParameter.CLEAR_COVER_TOP, o["cover_top"])
    _set_cover(doc, host, DB.BuiltInParameter.CLEAR_COVER_BOTTOM, o["cover_bottom"])
    _set_cover(doc, host, DB.BuiltInParameter.CLEAR_COVER_OTHER,
               max(min(o["cover_side"], o["end_cover"]), 0.0))


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


def _straight_set(doc, host, bar_type, a, b, normal, n, span):
    """One straight bar from a to b laid out as n bars over `span` mm along
    `normal` (a fixed number, so a tag reads the count)."""
    curve = DB.Line.CreateBound(a, b)
    straight = (_shape(doc, lambda s: _name(s) == u"S" and _is_straight(s))
                or _shape(doc, _is_straight))
    bar = None
    if straight is not None:
        try:
            bar = Rebar.CreateFromCurvesAndShape(doc, straight, bar_type, None, None, host, normal,
                                                 _curves([curve]), RebarHookOrientation.Left,
                                                 RebarHookOrientation.Left)
        except Exception:
            bar = None
    if bar is None:
        bar = Rebar.CreateFromCurves(doc, RebarStyle.Standard, bar_type, None, None, host, normal,
                                     _curves([curve]), RebarHookOrientation.Left,
                                     RebarHookOrientation.Left, True, True)
    if n > 1 and span > 1e-3:
        bar.GetShapeDrivenAccessor().SetLayoutAsFixedNumber(int(n), _ft(span), True, True, True)
    return bar


def _section_curves(frame, x, pts, closed):
    ps = [frame.p(x, y, z) for y, z in pts]
    if closed:
        ps.append(ps[0])
    return [DB.Line.CreateBound(ps[i], ps[i + 1]) for i in range(len(ps) - 1)]


def _seg_curves(frame, x, segs):
    out = []
    for s in segs:
        if s[0] == u"line":
            out.append(DB.Line.CreateBound(frame.p(x, *s[1]), frame.p(x, *s[2])))
        else:
            out.append(DB.Arc.Create(frame.p(x, *s[1]), frame.p(x, *s[2]), frame.p(x, *s[3])))
    return out


def _tie(doc, host, frame, bar_type, hook_type, x, tie):
    """One stirrup, link, U or loop in the section at x along the beam."""
    if tie.get("segs"):
        curves = _curves(_seg_curves(frame, x, tie["segs"]))
    else:
        curves = _curves(_section_curves(frame, x, tie["pts"], tie["closed"]))
    normal = frame.ux
    if tie["role"] == u"tie":
        style = RebarStyle.StirrupTie
        hooks = (hook_type, hook_type)
    elif tie["role"] == u"link":
        style = RebarStyle.Standard
        hooks = (hook_type, hook_type)
    else:
        style = RebarStyle.Standard
        hooks = (None, None)

    def make(o0, o1):
        return Rebar.CreateFromCurves(doc, style, bar_type, hooks[0], hooks[1], host, normal, curves,
                                      o0, o1, True, True)
    if tie["role"] != u"link":
        return make(RebarHookOrientation.Left, RebarHookOrientation.Left)
    # a link: both 180 degree hooks wrap their bars, so both turn toward them
    # (+y); Revit's Left / Right depends on how the bar runs, so try each
    want = frame.uy
    mid = frame.p(x, tie["pts"][0][0], (tie["pts"][0][1] + tie["pts"][-1][1]) / 2.0)
    for o0, o1 in ((RebarHookOrientation.Left, RebarHookOrientation.Left),
                   (RebarHookOrientation.Right, RebarHookOrientation.Right),
                   (RebarHookOrientation.Left, RebarHookOrientation.Right),
                   (RebarHookOrientation.Right, RebarHookOrientation.Left)):
        bar = make(o0, o1)
        doc.Regenerate()
        if _hooks_toward(bar, mid, want):
            return bar
        doc.Delete(bar.Id)
    return make(RebarHookOrientation.Left, RebarHookOrientation.Left)


def _hooks_toward(bar, mid, want):
    """True when every hook of a bar turns toward `want` from its leg."""
    from Autodesk.Revit.DB.Structure import MultiplanarOption
    curves = list(bar.GetCenterlineCurves(False, False, False,
                                          MultiplanarOption.IncludeOnlyPlanarCurves, 0))
    ok = True
    for c in curves:
        for i in (0, 1):
            p = c.GetEndPoint(i)
            ok = ok and (p - mid).DotProduct(want) > -1e-6
    return ok


def _spread_along(rebar, xa, xb, spacing, include_first):
    if xb - xa > 1e-3:
        rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(
            _ft(spacing), _ft(xb - xa), True, include_first, True)


def numbers(opts):
    """The plain numbers layout() needs, from reinforce()'s options (Revit
    types in, diameters in mm out)."""
    o = dict(opts)
    o["main_d"] = opts["main_bar"].BarNominalDiameter * MM
    o["main_dm"] = opts["main_bar"].BarModelDiameter * MM
    o["tie_d"] = opts["tie_bar"].BarNominalDiameter * MM
    o["tie_dm"] = opts["tie_bar"].BarModelDiameter * MM
    lb = opts.get("link_bar") or opts["tie_bar"]
    o["link_dm"] = lb.BarModelDiameter * MM
    o["link_bend"] = lb.StandardHookBendDiameter * MM
    o["tie_bend"] = opts["tie_bar"].StandardBendDiameter * MM
    o["kind"] = layout_info(opts["layout"])[1]
    return o


def _count(rebar):
    """Bars in a set that are there (an excluded first or last is not)."""
    return sum(1 for k in range(rebar.NumberOfBarPositions) if rebar.DoesBarExistAtPosition(k))


def reinforce(doc, host, frame, opts):
    """Add the bars to a beam. Needs an open transaction. opts are the
    window's (mm, Revit types); the beam's size comes from `frame`.
    Returns a dict of what was made."""
    o = numbers(opts)
    o["W"], o["H"], o["L"] = frame.width, frame.depth, frame.length
    plan = layout(o)
    set_covers(doc, host, o)
    doc.Regenerate()
    made = {"long": 0, "ties": 0}
    x0, x1 = plan["x0"], plan["x1"]

    # --- longitudinal bars: each row a set across, the side bars a set up each face ---
    for row in (plan["top"], plan["bottom"]):
        ys = row["ys"]
        if not ys:
            continue
        z = row["z"]
        bar = _straight_set(doc, host, opts["main_bar"], frame.p(x0, ys[0], z), frame.p(x1, ys[0], z),
                            frame.uy if len(ys) < 2 or ys[-1] > ys[0] else -frame.uy,
                            len(ys), abs(ys[-1] - ys[0]))
        made["long"] += bar.NumberOfBarPositions
    if plan["side"]:
        s = plan["side"]
        zs = s["zs"]
        sb = opts.get("side_bar") or opts["main_bar"]
        for y in (s["y"], -s["y"]):
            bar = _straight_set(doc, host, sb, frame.p(x0, y, zs[0]), frame.p(x1, y, zs[0]),
                                DB.XYZ.BasisZ, len(zs), abs(zs[-1] - zs[0]))
            made["long"] += bar.NumberOfBarPositions

    # --- stirrups, links and U bars: each a set along the beam per zone ---
    later = []
    for xa, xb, sp, first in plan["sets"]:
        for tie in plan["ties"]:
            if tie["role"] == u"link":
                bt, hk = opts.get("link_bar") or opts["tie_bar"], opts.get("link_hook")
            else:
                bt, hk = opts["tie_bar"], opts.get("tie_hook")
            bar = _tie(doc, host, frame, bt, hk, xa, tie)
            _spread_along(bar, xa, xb, sp, first)
            later.append((bar, first))
    # Revit puts a left-out first bar back when it regenerates (the zone
    # before already has one there): leave it out again once all are made
    doc.Regenerate()
    for bar, first in later:
        if not first and bar.NumberOfBarPositions > 1 and bar.IncludeFirstBar:
            bar.IncludeFirstBar = False
    doc.Regenerate()
    made["ties"] = sum(_count(b) for b, _f in later)
    return made


def remove_bars(doc, host):
    """Delete every bar hosted on the beam. Needs an open transaction.
    Returns how many sets went."""
    data = RebarHostData.GetRebarHostData(host)
    gone = 0
    if data is not None:
        for bar in list(data.GetRebarsInHost()):
            doc.Delete(bar.Id)
            gone += 1
    return gone


# ── REVIT: the beam itself ──────────────────────────────────────────────────

WIDTH_NAMES = (u"Width", u"b", u"B")
DEPTH_NAMES = (u"Depth", u"h", u"d", u"H")


def _type_param(symbol, names):
    for n in names:
        p = symbol.LookupParameter(n)
        if p is not None and p.StorageType == DB.StorageType.Double:
            return p
    return None


def _is_concrete(symbol):
    p = symbol.get_Parameter(DB.BuiltInParameter.STRUCTURAL_MATERIAL_PARAM)
    mat = symbol.Document.GetElement(p.AsElementId()) if p is not None else None
    if mat is not None:
        return u"concrete" in (_name(mat) + u" " + (mat.MaterialClass or u"")).lower()
    return u"concrete" in (symbol.FamilyName or u"").lower()


def beam_types(doc):
    """Concrete structural framing types with a width and a depth."""
    out = []
    for s in DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol) \
            .OfCategory(DB.BuiltInCategory.OST_StructuralFraming):
        if _type_param(s, WIDTH_NAMES) and _type_param(s, DEPTH_NAMES) and _is_concrete(s):
            out.append(s)
    return sorted(out, key=lambda s: (s.FamilyName, _name(s)))


def type_size(symbol):
    """(width, depth) of a beam type in mm, or None."""
    w, d = _type_param(symbol, WIDTH_NAMES), _type_param(symbol, DEPTH_NAMES)
    if w is None or d is None:
        return None
    return w.AsDouble() * MM, d.AsDouble() * MM


def type_name(size):
    return u"{0:g} x {1:g}".format(round(size[0], 1), round(size[1], 1))


def ensure_type(doc, symbol, size):
    """A type of `symbol`'s family of this (width, depth): `symbol` itself
    when it is that size, another of the family that is, else a new one
    duplicated from it. Needs an open transaction."""
    have = type_size(symbol)
    if have and abs(have[0] - size[0]) < 0.5 and abs(have[1] - size[1]) < 0.5:
        return symbol
    family = symbol.Family
    for sid in family.GetFamilySymbolIds():
        s = doc.GetElement(sid)
        h = type_size(s)
        if h and abs(h[0] - size[0]) < 0.5 and abs(h[1] - size[1]) < 0.5:
            return s
    name = type_name(size)
    names = set(_name(doc.GetElement(i)) for i in family.GetFamilySymbolIds())
    k = 2
    base = name
    while name in names:
        name = u"{0} ({1})".format(base, k)
        k += 1
    new = symbol.Duplicate(name)
    _type_param(new, WIDTH_NAMES).Set(_ft(size[0]))
    _type_param(new, DEPTH_NAMES).Set(_ft(size[1]))
    return new


def place_beam(doc, symbol, a, b, level, top_offset, mark=None):
    """A beam from a to b (only X and Y are used) with its top at the level
    plus top_offset (mm). Needs an open transaction. Returns the beam."""
    if not symbol.IsActive:
        symbol.Activate()
        doc.Regenerate()
    z = level.ProjectElevation
    line = DB.Line.CreateBound(DB.XYZ(a.X, a.Y, z), DB.XYZ(b.X, b.Y, z))
    beam = doc.Create.NewFamilyInstance(line, symbol, level, StructuralType.Beam)
    P = DB.BuiltInParameter
    zj = beam.get_Parameter(P.Z_JUSTIFICATION)
    if zj is not None and not zj.IsReadOnly:
        zj.Set(0)                                   # Top
    # the top on the level plus the offset: both ends' offsets carry it, and
    # the family's own z offset (some come with one) is cleared
    for bip, value in ((P.STRUCTURAL_BEAM_END0_ELEVATION, top_offset),
                       (P.STRUCTURAL_BEAM_END1_ELEVATION, top_offset), (P.Z_OFFSET_VALUE, 0.0)):
        p = beam.get_Parameter(bip)
        if p is not None and not p.IsReadOnly:
            p.Set(_ft(value))
    if mark:
        m = beam.get_Parameter(P.ALL_MODEL_MARK)
        if m is not None and not m.IsReadOnly:
            m.Set(mark)
    return beam


def beam_level(doc, host, frame):
    """(level, top offset mm) of a beam: its reference level and its top
    (from its concrete, `frame`) as an offset from it, or (None, 0)."""
    p = host.get_Parameter(DB.BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM)
    level = doc.GetElement(p.AsElementId()) if p is not None else None
    if level is None:
        return None, 0.0
    top = frame.origin.Z + _ft(frame.depth)
    return level, (top - level.ProjectElevation) * MM
