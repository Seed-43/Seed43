# -*- coding: utf-8 -*-
# "Slab Slope"
# "Seed43"
# """
# Apply a constant fall to a floor/slab through shape editing. You give the
# slope (1:100, 1%, 1 deg), pick the TOP edges (they stay at the current slab
# level, 0) and then the LOW edges. The slab falls from the top edges to the
# low line at that slope, whatever its outline; anything behind a top edge
# stays flat. When top edges sit on both sides of the low line (a valley),
# you choose to keep the slope on both sides (one side's top moves off 0) or
# keep both tops at 0 (the other side's slope changes to suit).
# """
import re
import math

from pyrevit import revit, DB, forms
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

doc = revit.doc
uidoc = revit.uidoc

# ── CONSTANTS ───────────────────────────────────────────────────────────────

TITLE = u"Slab Slope"
DEFAULT_SLOPE = u"1:100"
TOL = 0.005                    # ft (~1.5 mm): vertex "is on a line"
MAX_SLOPE = 1.0                # tan(45 deg); anything steeper is a typo
MM = 304.8

# Text for the orange bar (forms.WarningBar) shown across the top while picking.
BAR_TOP = (u"Slab Slope, step 1 of 2: click every TOP edge (stays at the "
           u"current slab level), then press Finish. ESC cancels.")
BAR_LOW = (u"Slab Slope, step 2 of 2: click every LOW edge (the bottom of "
           u"the fall), then press Finish. ESC cancels.")
BAR_SIDE_SLOPE = (u"Slab Slope: click the edge that stays at slab level. "
                  u"Both sides fall at {0}; the other edge will move up or "
                  u"down to suit. ESC cancels.")
BAR_SIDE_LEVEL = (u"Slab Slope: click an edge that stays at {0}. The other "
                  u"side's slope changes to meet slab level. ESC cancels.")


# ── HELPERS (private to this file) ──────────────────────────────────────────

# --- Slope input ---
def parse_slope(text):
    """Return a slope as rise over run (0.01 = 1:100) from free text.

    Accepts 1:100, 1/100, 1%, 1 deg / 1 degree / 1° and a bare number,
    which is read as a percentage.
    """
    t = (text or u"").strip().lower().replace(unichr(176), u"deg")
    t = t.replace(u" ", u"").replace(u",", u".")
    if not t:
        raise ValueError(u"No slope entered.")
    try:
        m = re.match(r"^([0-9.]+)(deg|degree|degrees)$", t)
        if m:
            value = math.tan(math.radians(float(m.group(1))))
        elif t.endswith(u"%"):
            value = float(t[:-1]) / 100.0
        elif u":" in t or u"/" in t:
            a, b = re.split(r"[:/]", t, 1)
            value = float(a) / float(b)
        else:
            value = float(t) / 100.0
    except (ValueError, ZeroDivisionError):
        raise ValueError(u"Could not read '{0}' as a slope. Try 1:100, 1%, "
                         u"or 1 deg.".format(text))
    if value <= 0 or value > MAX_SLOPE:
        raise ValueError(u"Slope must be above 0 and no steeper than 45 deg.")
    return value


def ratio(slope):
    """Slope as 1:N text."""
    n = 1.0 / slope
    return u"1:{0:.0f}".format(n) if n >= 10 else u"1:{0:.1f}".format(n)


# --- Picking ---
class _FloorFilter(ISelectionFilter):
    """Only floors/slabs can be picked."""

    def AllowElement(self, element):
        return isinstance(element, DB.Floor)

    def AllowReference(self, reference, position):
        return True


def _edge_info(ref):
    """Floor and (start, end, straight) of a picked edge."""
    floor = doc.GetElement(ref)
    geo = floor.GetGeometryObjectFromReference(ref)
    curve = geo.AsCurve() if isinstance(geo, DB.Edge) else None
    if curve is None:
        raise ValueError(u"Pick an edge of the slab.")
    return floor, (curve.GetEndPoint(0), curve.GetEndPoint(1),
                   isinstance(curve, DB.Line))


def pick_edges(prompt, floor=None):
    """Pick several edges of one slab. Returns (floor, [(a, b, straight)])."""
    refs = uidoc.Selection.PickObjects(ObjectType.Edge, _FloorFilter(), prompt)
    edges = []
    for ref in refs:
        f, edge = _edge_info(ref)
        if floor is None:
            floor = f
        elif f.Id != floor.Id:
            raise ValueError(u"All edges must be on the same slab.")
        edges.append(edge)
    if not edges:
        raise ValueError(u"No edges picked.")
    return floor, edges


def _length(e):
    return math.hypot(e[1].X - e[0].X, e[1].Y - e[0].Y)


# ── CORE LOGIC ──────────────────────────────────────────────────────────────

class Fall(object):
    """The geometry of the fall: the low line and the top edges either side.

    u(p) is a point's distance in plan from the low line, positive on one
    side and negative on the other. `sides` maps +1 / -1 to the distance from
    the low line to the nearest top edge on that side; a side with no top
    edges is missing from it.
    """

    # --- construction ---
    def __init__(self, top_edges, low_edges):
        self.top_edges = top_edges
        self.low_edges = low_edges
        # The low line's direction comes from a straight low edge; with only
        # curved low edges, from a straight top edge instead.
        straight = ([e for e in low_edges if e[2]] or
                    [e for e in top_edges if e[2]])
        if not straight:
            raise ValueError(u"Pick at least one straight edge, it sets the "
                             u"direction of the fall.")
        a, b, _ = max(straight, key=_length)
        length = _length((a, b))
        if length < TOL:
            raise ValueError(u"That edge is vertical in plan, pick another.")
        self.origin = a
        self.nx = -(b.Y - a.Y) / length
        self.ny = (b.X - a.X) / length

        lows = [self._t(p) for e in low_edges for p in (e[0], e[1])]
        if max(lows) - min(lows) > 0.05:          # ~15 mm
            raise ValueError(u"The low edges must lie on one straight line, "
                             u"parallel to the top edges.")
        self.t_low = sum(lows) / len(lows)

        self.sides = {}
        self.top_points = []
        for e in top_edges:
            us = [self.u(e[0]), self.u(e[1])]
            if min(us) < -TOL and max(us) > TOL:
                raise ValueError(u"A top edge crosses the low line.")
            side = 1 if sum(us) > 0 else -1
            near = min(abs(x) for x in us)
            if near < TOL:
                raise ValueError(u"A top edge touches the low line.")
            self.sides[side] = min(self.sides.get(side, near), near)
            self.top_points.extend([(e[0], side), (e[1], side)])

    # --- public methods ---
    def u(self, p):
        """Signed plan distance of p from the low line (ft)."""
        return self._t(p) - self.t_low

    def two_sided(self):
        """True when top edges sit both sides of the low line at different
        distances, so slope and levels cannot both be kept."""
        return (len(self.sides) == 2 and
                abs(self.sides[1] - self.sides[-1]) > TOL)

    def side_of(self, edge):
        """Which side of the low line a picked edge is on (+1 / -1)."""
        us = [self.u(edge[0]), self.u(edge[1])]
        side = 1 if sum(us) > 0 else -1
        if side not in self.sides:
            raise ValueError(u"That edge is not on a side with top edges.")
        return side

    def two_sided_levels(self, slope, keep_slope, side):
        """Low level and top levels for a valley.

        `side` sets the fall at `slope` and its top stays at 0. With
        keep_slope the other side also falls at `slope`, so its top ends
        above or below 0; otherwise the other top stays at 0 and its slope
        changes to suit.
        """
        other = -side
        low = -slope * self.sides[side]
        top_other = low + slope * self.sides[other] if keep_slope else 0.0
        return low, {side: 0.0, other: top_other}

    def single_levels(self, slope):
        """Low level and top levels when one fit serves: top edges on one
        side only, or both sides the same distance from the low line."""
        side = list(self.sides.keys())[0]
        return -slope * self.sides[side], dict((s, 0.0) for s in self.sides)

    def level_at(self, p, low, tops, slope):
        """Target offset (ft) for a vertex at p."""
        for q, side in self.top_points:
            if abs(q.X - p.X) < TOL and abs(q.Y - p.Y) < TOL:
                return tops[side]
        u = self.u(p)
        side = 1 if u > 0 else -1
        dist = abs(u)
        if side in self.sides:
            d = self.sides[side]
            if dist >= d:
                return tops[side]           # behind the top edge: flat
            return low + (tops[side] - low) * dist / d
        return low - slope * dist           # no top edge this side: keep falling

    def lines(self):
        """u values that need a crease: the low line and each top line."""
        return [0.0] + [d * s for s, d in self.sides.items()]

    # --- private helpers ---
    def _t(self, p):
        return (p.X - self.origin.X) * self.nx + (p.Y - self.origin.Y) * self.ny


def _add_creases(editor, fall):
    """Split the slab along the low line and the top lines.

    Stops Revit triangulating across a change of slope. Skipped where the
    line is just one picked boundary edge, which is creased already.
    """
    def near(p, q):
        return abs(p.X - q.X) < TOL and abs(p.Y - q.Y) < TOL

    picked = fall.top_edges + fall.low_edges
    count = 0
    for target in fall.lines():
        verts = [v for v in editor.SlabShapeVertices
                 if abs(fall.u(v.Position) - target) < TOL]
        if len(verts) < 2:
            continue
        verts.sort(key=lambda v: v.Position.X * fall.ny - v.Position.Y * fall.nx)
        first, last = verts[0], verts[-1]
        if len(verts) == 2 and any(
                (near(first.Position, e[0]) and near(last.Position, e[1])) or
                (near(first.Position, e[1]) and near(last.Position, e[0]))
                for e in picked):
            continue
        try:
            editor.AddSplitLine(first, last)
            count += 1
        except Exception:
            pass                            # boundary already creased there
        doc.Regenerate()
    return count


def apply_slope(floor, fall, low, tops, slope):
    """Set every shape-edit vertex of `floor` to the fall.

    Must run in an open transaction. Returns a dict of what was done,
    including a read-back of the vertex heights from the model.
    """
    editor = floor.GetSlabShapeEditor()
    if editor.IsEnabled:
        editor.ResetSlabShape()
        doc.Regenerate()
    # NOTE: ResetSlabShape also switches shape editing off, leaving no
    # vertices, so Enable has to follow it too.
    if not editor.IsEnabled:
        editor.Enable()
    doc.Regenerate()
    creases = _add_creases(editor, fall)

    verts = list(editor.SlabShapeVertices)
    if not verts:
        raise ValueError(u"Revit gave no shape-edit points for this slab.")
    before, wanted = {}, {}
    for v in verts:
        p = v.Position
        key = (round(p.X, 3), round(p.Y, 3))
        before[key] = p.Z
        wanted[key] = fall.level_at(p, low, tops, slope)
        editor.ModifySubElement(v, wanted[key])
    doc.Regenerate()

    # Read the model back: each vertex must now sit at its old z + offset.
    worst, checked = 0.0, 0
    for v in editor.SlabShapeVertices:
        key = (round(v.Position.X, 3), round(v.Position.Y, 3))
        if key in before:
            worst = max(worst, abs(v.Position.Z - (before[key] + wanted[key])))
            checked += 1
    return {"vertices": len(verts), "checked": checked, "creases": creases,
            "worst_mm": worst * MM,
            "low_mm": low * MM,
            "levels_mm": sorted(set(round(w * MM, 1) for w in wanted.values()))}


# ── UI / ENTRY POINT ────────────────────────────────────────────────────────

def main():
    text = forms.ask_for_string(
        default=DEFAULT_SLOPE,
        prompt=u"Fall (examples: 1:100, 1%, 1 deg):",
        title=TITLE)
    if text is None:
        return
    try:
        slope = parse_slope(text)
        # the orange bar at the top of the window says what to do next
        with forms.WarningBar(title=BAR_TOP) as bar:
            floor, top_edges = pick_edges(u"Pick the TOP edges, then Finish")
            bar.message_tb.Text = BAR_LOW
            floor, low_edges = pick_edges(u"Pick the LOW edges, then Finish",
                                          floor)
        fall = Fall(top_edges, low_edges)
    except OperationCanceledException:
        return
    except ValueError as err:
        forms.alert(unicode(err), title=TITLE)
        return

    if fall.two_sided():
        # The two sides are different lengths, so the slope and the top
        # levels cannot all be kept: ask which to keep, then which side.
        keep_slope_text = u"Keep slope {0}, adjust levels one side".format(
            ratio(slope))
        keep_level_text = u"Keep levels, adjust slope one side"
        choice = forms.CommandSwitchWindow.show(
            [keep_slope_text, keep_level_text],
            message=u"The two sides are different lengths. Keep the slope "
                    u"or keep the top edges level?")
        if not choice:
            return
        keep_slope = choice == keep_slope_text
        bar_text = (BAR_SIDE_SLOPE if keep_slope else BAR_SIDE_LEVEL).format(
            ratio(slope))
        try:
            with forms.WarningBar(title=bar_text):
                ref = uidoc.Selection.PickObject(
                    ObjectType.Edge, _FloorFilter(),
                    u"Pick a TOP edge on the side that sets the levels")
            picked_floor, edge = _edge_info(ref)
            if picked_floor.Id != floor.Id:
                raise ValueError(u"Pick an edge on the same slab.")
            side = fall.side_of(edge)
        except OperationCanceledException:
            return
        except ValueError as err:
            forms.alert(unicode(err), title=TITLE)
            return
        low, tops = fall.two_sided_levels(slope, keep_slope, side)
    else:
        low, tops = fall.single_levels(slope)

    if floor.GetSlabShapeEditor().IsEnabled:
        if not forms.alert(
                u"This slab already has shape editing. It will be reset and "
                u"replaced with the new slope.",
                title=TITLE, yes=True, no=True):
            return

    tx = DB.Transaction(doc, u"Slab Slope")
    tx.Start()
    try:
        info = apply_slope(floor, fall, low, tops, slope)
        tx.Commit()
    except Exception as err:
        tx.RollBack()
        forms.alert(u"Nothing changed. Revit said: {0}".format(err),
                    title=TITLE)
        return

    # Quiet on success; only speak up if the model does not match what was set.
    if info["worst_mm"] >= 0.5:
        forms.alert(
            u"The slope was applied but Revit's vertex heights differ from "
            u"what was set (worst error {0:.2f} mm over {1} vertices). Check "
            u"the slab.".format(info["worst_mm"], info["checked"]),
            title=TITLE)


if __name__ == "__main__":
    main()
