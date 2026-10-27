# -*- coding: utf-8 -*-
"""Steel parts out of an IFC file: member lines, profiles, shapes, bolts.

Every product with a body becomes a Part. A part whose body is one profile
swept straight along a line (through any end cuts) also gets a Member: the
line through the middle of the profile's bounding box, the profile's own
x and y axes in the world, and the profile's kind, dimensions and outline.
The Revit side decides which members become native framing.

Everything is in millimetres, in the file's world coordinates. No Revit
types in here, so it runs under CPython for testing as well as IronPython 2.
"""
import math
import re

try:
    from ._ifc_step import (IfcFile, IDENTITY, mat_mul, mat_point,
                            mat_vector, frame)
except (ImportError, ValueError):       # run outside the package (tests)
    from _ifc_step import (IfcFile, IDENTITY, mat_mul, mat_point,
                           mat_vector, frame)

__all__ = ["Profile", "Member", "Part", "SteelModel", "read_steel",
           "name_key", "is_bolt_name"]

CIRCLE_SEGMENTS = 32
ARC_STEP = math.radians(10.0)

MEMBER_TYPES = {"IFCBEAM": "beam", "IFCCOLUMN": "column", "IFCMEMBER": "brace"}
FASTENER_TYPES = {"IFCMECHANICALFASTENER", "IFCFASTENER"}
NO_BODY_TYPES = {"IFCOPENINGELEMENT", "IFCSPACE", "IFCGRID", "IFCANNOTATION",
                 "IFCVIRTUALELEMENT", "IFCSITE", "IFCBUILDING",
                 "IFCBUILDINGSTOREY", "IFCPROJECT"}

_BOLT_NAME = re.compile(
    r"(^|[^A-Z])(NUT|WASHER|BOLT|SCREW)([^A-Z]|$)|^TR\(|^M\d+(\.\d+)?_", re.I)


def is_bolt_name(text):
    return bool(text) and bool(_BOLT_NAME.search(text))


# ── NAMES ───────────────────────────────────────────────────────────────────

_SHAPE_WORDS = ["PFC", "TFB", "SHS", "RHS", "CHS", "UB", "UC", "WB", "WC",
                "EA", "UA", "BT", "CT", "FL", "FB", "PL", "RB", "SB", "ROD",
                "DHS", "C", "Z"]
_NUM = re.compile(r"\d+(?:\.\d+)?")


def name_key(name):
    """(shape word, rounded numbers) from a section name, so Tekla's
    'UB180*18', '180UB18.1' and 'UB 180x18' all give ('UB', (180, 18))."""
    if not name:
        return None
    up = re.sub(r"\(.*?\)", " ", name.upper())
    word = None
    for w in _SHAPE_WORDS:
        if re.search(r"(?<![A-Z])" + w + r"(?![A-Z])", up) or \
                re.search(r"\d" + w + r"(?![A-Z])", up) or \
                re.search(r"(?<![A-Z])" + w + r"\d", up):
            word = w
            break
    nums = tuple(int(round(float(n))) for n in _NUM.findall(up))
    if word is None or not nums:
        return None
    if word in ("ROD", "RB"):
        word = "ROD"
    if word in ("FL", "FB"):
        word = "FL"
    if word in ("SHS", "EA") and len(nums) == 2:
        nums = (nums[0], nums[0], nums[1])     # '89x6.0 SHS' is 89x89x6.0
    return (word, nums)


# ── 2D HELPERS ──────────────────────────────────────────────────────────────

def _circle_pts(cx, cy, r, n=CIRCLE_SEGMENTS):
    return [(cx + r * math.cos(2 * math.pi * i / n),
             cy + r * math.sin(2 * math.pi * i / n)) for i in range(n)]


def loop_points(loop):
    """A loop as a plain point list (circles become polygons)."""
    if isinstance(loop, tuple) and loop and loop[0] == "circle":
        return _circle_pts(loop[1], loop[2], loop[3])
    return loop


def signed_area(pts):
    a = 0.0
    n = len(pts)
    for i in range(n):
        x0, y0 = pts[i]
        x1, y1 = pts[(i + 1) % n]
        a += x0 * y1 - x1 * y0
    return a / 2.0


def _centroid(loops):
    """Area centroid of an outer loop minus inner loops."""
    total = cx = cy = 0.0
    for k, loop in enumerate(loops):
        pts = loop_points(loop)
        a = signed_area(pts)
        if a == 0:
            continue
        sx = sy = 0.0
        n = len(pts)
        for i in range(n):
            x0, y0 = pts[i]
            x1, y1 = pts[(i + 1) % n]
            c = x0 * y1 - x1 * y0
            sx += (x0 + x1) * c
            sy += (y0 + y1) * c
        a = abs(a)
        sign = 1.0 if k == 0 else -1.0
        gx = sx / (6.0 * signed_area(pts))
        gy = sy / (6.0 * signed_area(pts))
        total += sign * a
        cx += sign * a * gx
        cy += sign * a * gy
    if abs(total) < 1e-9:
        return (0.0, 0.0), 0.0
    return (cx / total, cy / total), total


# ── PROFILES ────────────────────────────────────────────────────────────────

class Profile(object):
    """A cross-section in its own 2D frame (mm). kind is one of I, U, L, T,
    RH (rectangular hollow), CH (circular hollow), C (round bar),
    R (rectangle / flat), CS (lipped C), Z, ARB (any outline)."""

    def __init__(self, kind, name, loops, dims=None):
        self.kind = kind
        self.name = name or u""
        self.loops = loops
        self.dims = dims or {}
        xs, ys = [], []
        for loop in loops[:1]:
            for x, y in loop_points(loop):
                xs.append(x)
                ys.append(y)
        if xs:
            self.bbox = (min(xs), min(ys), max(xs), max(ys))
        else:
            self.bbox = (0.0, 0.0, 0.0, 0.0)
        self.centre = ((self.bbox[0] + self.bbox[2]) / 2.0,
                       (self.bbox[1] + self.bbox[3]) / 2.0)
        self.extent = (self.bbox[2] - self.bbox[0], self.bbox[3] - self.bbox[1])
        (gx, gy), self.area = _centroid(loops)
        # centroid relative to the bounding-box centre
        self.centroid = (gx - self.centre[0], gy - self.centre[1])

    def key_dims(self):
        """The dimensions the Revit type catalogues carry, for matching."""
        d = self.dims
        k = self.kind
        if k in ("I", "U", "T"):
            return {"Height": d.get("depth"), "Width": d.get("width"),
                    "Flange Thickness": d.get("tf"), "Web Thickness": d.get("tw")}
        if k == "L":
            return {"Height": d.get("depth"), "Width": d.get("width"),
                    "Flange Thickness": d.get("t")}
        if k == "RH":
            return {"Height": d.get("depth"), "Width": d.get("width"),
                    "Wall Nominal Thickness": d.get("t")}
        if k == "CH":
            return {"Diameter": d.get("diameter"),
                    "Wall Nominal Thickness": d.get("t")}
        if k == "C":
            return {"Diameter": d.get("diameter")}
        if k == "R":
            return {"Height": d.get("depth"), "Width": d.get("width")}
        return {}


def _apply2d(m, pts):
    """2D points through a 2D placement matrix (from axis_placement)."""
    out = []
    for x, y in pts:
        p = mat_point(m, (x, y, 0.0))
        out.append((p[0], p[1]))
    return out


def _place_loops(m, loops):
    out = []
    for loop in loops:
        if isinstance(loop, tuple) and loop and loop[0] == "circle":
            c = mat_point(m, (loop[1], loop[2], 0.0))
            out.append(("circle", c[0], c[1], loop[3]))
        else:
            out.append(_apply2d(m, loop))
    return out


def _rect(w, h):
    return [(-w / 2, -h / 2), (w / 2, -h / 2), (w / 2, h / 2), (-w / 2, h / 2)]


class ProfileReader(object):
    def __init__(self, ifc):
        self.ifc = ifc
        self._cache = {}

    def read(self, ref):
        if ref in self._cache:
            return self._cache[ref]
        try:
            p = self._read(ref)
        except Exception:
            p = None
        self._cache[ref] = p
        return p

    def _num(self, v):
        if v is None:
            return 0.0
        if isinstance(v, tuple):
            v = v[1]
        return float(v) * self.ifc.mm

    def _read(self, ref):
        ifc = self.ifc
        t = ifc.type_of(ref)
        a = ifc.args(ref)
        name = a[1] if len(a) > 1 and isinstance(a[1], type(u"")) else u""
        n = self._num
        if t == "IFCDERIVEDPROFILEDEF":
            parent = self.read(a[2])
            if parent is None:
                return None
            m = self._operator2d(a[3])
            return Profile(parent.kind, a[4] if len(a) > 4 and a[4] else parent.name,
                           _place_loops(m, parent.loops), parent.dims)
        if t in ("IFCARBITRARYCLOSEDPROFILEDEF", "IFCARBITRARYPROFILEDEFWITHVOIDS"):
            outer = self.curve(a[2])
            if not outer or len(outer) < 3:
                return None
            loops = [outer]
            if t == "IFCARBITRARYPROFILEDEFWITHVOIDS":
                for c in a[3] or []:
                    inner = self.curve(c)
                    if inner and len(inner) >= 3:
                        loops.append(inner)
            return Profile("ARB", name, loops)

        place = ifc.axis_placement(a[2]) if len(a) > 2 else IDENTITY
        if t == "IFCRECTANGLEPROFILEDEF":
            w, h = n(a[3]), n(a[4])
            return Profile("R", name, _place_loops(place, [_rect(w, h)]),
                           {"width": min(w, h), "depth": max(w, h)})
        if t == "IFCRECTANGLEHOLLOWPROFILEDEF":
            w, h, th = n(a[3]), n(a[4]), n(a[5])
            loops = [_rect(w, h), list(reversed(_rect(w - 2 * th, h - 2 * th)))]
            return Profile("RH", name, _place_loops(place, loops),
                           {"width": min(w, h), "depth": max(w, h), "t": th})
        if t == "IFCCIRCLEPROFILEDEF":
            r = n(a[3])
            return Profile("C", name, _place_loops(place, [("circle", 0.0, 0.0, r)]),
                           {"diameter": 2 * r})
        if t == "IFCCIRCLEHOLLOWPROFILEDEF":
            r, th = n(a[3]), n(a[4])
            loops = [("circle", 0.0, 0.0, r), ("circle", 0.0, 0.0, r - th)]
            return Profile("CH", name, _place_loops(place, loops),
                           {"diameter": 2 * r, "t": th})
        if t == "IFCISHAPEPROFILEDEF" or t == "IFCASYMMETRICISHAPEPROFILEDEF":
            w, d, tw, tf = n(a[3]), n(a[4]), n(a[5]), n(a[6])
            x, y = w / 2, d / 2
            pts = [(-x, -y), (x, -y), (x, -y + tf), (tw / 2, -y + tf),
                   (tw / 2, y - tf), (x, y - tf), (x, y), (-x, y),
                   (-x, y - tf), (-tw / 2, y - tf), (-tw / 2, -y + tf),
                   (-x, -y + tf)]
            return Profile("I", name, _place_loops(place, [pts]),
                           {"depth": d, "width": w, "tw": tw, "tf": tf})
        if t == "IFCUSHAPEPROFILEDEF":
            d, w, tw, tf = n(a[3]), n(a[4]), n(a[5]), n(a[6])
            x, y = w / 2, d / 2
            pts = [(-x, -y), (x, -y), (x, -y + tf), (-x + tw, -y + tf),
                   (-x + tw, y - tf), (x, y - tf), (x, y), (-x, y)]
            return Profile("U", name, _place_loops(place, [pts]),
                           {"depth": d, "width": w, "tw": tw, "tf": tf})
        if t == "IFCLSHAPEPROFILEDEF":
            d = n(a[3])
            w = n(a[4]) if a[4] is not None else d
            th = n(a[5])
            x, y = w / 2, d / 2
            pts = [(-x, -y), (x, -y), (x, -y + th), (-x + th, -y + th),
                   (-x + th, y), (-x, y)]
            return Profile("L", name, _place_loops(place, [pts]),
                           {"depth": d, "width": w, "t": th})
        if t == "IFCTSHAPEPROFILEDEF":
            d, w, tw, tf = n(a[3]), n(a[4]), n(a[5]), n(a[6])
            x, y = w / 2, d / 2
            pts = [(-tw / 2, -y), (tw / 2, -y), (tw / 2, y - tf), (x, y - tf),
                   (x, y), (-x, y), (-x, y - tf), (-tw / 2, y - tf)]
            return Profile("T", name, _place_loops(place, [pts]),
                           {"depth": d, "width": w, "tw": tw, "tf": tf})
        if t == "IFCCSHAPEPROFILEDEF":
            d, w, th, g = n(a[3]), n(a[4]), n(a[5]), n(a[6])
            x, y = w / 2, d / 2
            pts = [(-x, -y), (x, -y), (x, -y + g), (x - th, -y + g),
                   (x - th, -y + th), (-x + th, -y + th), (-x + th, y - th),
                   (x - th, y - th), (x - th, y - g), (x, y - g), (x, y), (-x, y)]
            return Profile("CS", name, _place_loops(place, [pts]),
                           {"depth": d, "width": w, "t": th, "lip": g})
        if t == "IFCZSHAPEPROFILEDEF":
            d, w, tw, tf = n(a[3]), n(a[4]), n(a[5]), n(a[6])
            y = d / 2
            pts = [(-w + tw / 2, -y), (tw / 2, -y), (tw / 2, y - tf),
                   (w - tw / 2, y - tf), (w - tw / 2, y), (-tw / 2, y),
                   (-tw / 2, -y + tf), (-w + tw / 2, -y + tf)]
            return Profile("Z", name, _place_loops(place, [pts]),
                           {"depth": d, "width": w, "tw": tw, "tf": tf})
        return None

    def _operator2d(self, ref):
        a = self.ifc.args(ref)
        x = self.ifc.direction(a[0]) or (1.0, 0.0, 0.0)
        o = self.ifc.point(a[2])
        return frame((0.0, 0.0, 1.0), x, (o[0], o[1], 0.0))

    # -- curves -----------------------------------------------------------------

    def curve(self, ref):
        """A closed 2D curve as a point list (mm), or None."""
        ifc = self.ifc
        t = ifc.type_of(ref)
        a = ifc.args(ref)
        if t == "IFCPOLYLINE":
            pts = [ifc.point(p)[:2] for p in a[0]]
            if len(pts) > 1 and _close(pts[0], pts[-1]):
                pts = pts[:-1]
            return pts
        if t == "IFCCOMPOSITECURVE":
            pts = []
            for seg in a[0]:
                sa = ifc.args(seg)
                part = self.open_curve(sa[2])
                if part is None:
                    return None
                if sa[1] in ("F", False):
                    part = list(reversed(part))
                if pts and part and _close(pts[-1], part[0]):
                    part = part[1:]
                pts.extend(part)
            if len(pts) > 1 and _close(pts[0], pts[-1]):
                pts = pts[:-1]
            return pts
        if t == "IFCCIRCLE":
            m = ifc.axis_placement(a[0])
            r = float(a[1]) * ifc.mm
            return _apply2d(m, _circle_pts(0.0, 0.0, r))
        return None

    def open_curve(self, ref):
        ifc = self.ifc
        t = ifc.type_of(ref)
        a = ifc.args(ref)
        if t == "IFCPOLYLINE":
            return [ifc.point(p)[:2] for p in a[0]]
        if t == "IFCTRIMMEDCURVE":
            basis = a[0]
            bt = ifc.type_of(basis)
            p1 = self._trim_point(a[1])
            p2 = self._trim_point(a[2])
            sense = a[3] in ("T", True)
            if bt == "IFCLINE" and p1 and p2:
                return [p1, p2]
            if bt == "IFCCIRCLE":
                ba = ifc.args(basis)
                m = ifc.axis_placement(ba[0])
                r = float(ba[1]) * ifc.mm
                cx, cy = m[3][0], m[3][1]
                x0 = math.atan2(m[0][1], m[0][0])
                a1 = self._trim_angle(a[1], p1, cx, cy, x0)
                a2 = self._trim_angle(a[2], p2, cx, cy, x0)
                if a1 is None or a2 is None:
                    return None
                if sense:
                    while a2 <= a1:
                        a2 += 2 * math.pi
                else:
                    while a2 >= a1:
                        a2 -= 2 * math.pi
                steps = max(2, int(abs(a2 - a1) / ARC_STEP) + 1)
                return [(cx + r * math.cos(a1 + (a2 - a1) * i / steps),
                         cy + r * math.sin(a1 + (a2 - a1) * i / steps))
                        for i in range(steps + 1)]
            return None
        return self.curve(ref)

    def _trim_point(self, trims):
        for v in trims or []:
            if not isinstance(v, tuple) and self.ifc.type_of(v) == "IFCCARTESIANPOINT":
                return self.ifc.point(v)[:2]
        return None

    def _trim_angle(self, trims, pt, cx, cy, x0):
        if pt is not None:
            return math.atan2(pt[1] - cy, pt[0] - cx)
        for v in trims or []:
            if isinstance(v, tuple):
                return x0 + float(v[1]) * self.ifc.rad
        return None


def _close(p, q):
    return abs(p[0] - q[0]) < 1e-6 and abs(p[1] - q[1]) < 1e-6


# ── BODIES ──────────────────────────────────────────────────────────────────

class Extrusion(object):
    """A profile swept along dir for depth (mm), in world coordinates.
    plane is (origin, x, y): the profile's 2D frame placed in the world."""

    def __init__(self, profile, plane, direction, depth):
        self.profile = profile
        self.plane = plane
        self.dir = direction
        self.depth = depth
        self.clips = []         # dicts, see BodyReader._clip


class Brep(object):
    """Planar faces: each a list of loops (outer first) of world points."""

    def __init__(self, faces):
        self.faces = faces


def _norm3(v):
    l = (v[0] * v[0] + v[1] * v[1] + v[2] * v[2]) ** 0.5
    return (v[0] / l, v[1] / l, v[2] / l) if l > 1e-12 else None


def _dot3(a, b):
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


class BodyReader(object):
    def __init__(self, ifc, profiles):
        self.ifc = ifc
        self.profiles = profiles
        self.unsupported = {}

    def body(self, product_args, world):
        """Solids of a product's body representation, in world mm."""
        ifc = self.ifc
        out = []
        if product_args[6] is None:
            return out
        reps = ifc.args(product_args[6])[2] or []
        chosen = []
        for r in reps:
            ra = ifc.args(r)
            if ra[1] == "Body":
                chosen.append(ra)
        if not chosen:
            for r in reps:
                ra = ifc.args(r)
                if ra[1] not in ("Axis", "Box", "FootPrint", "Annotation",
                                 "Profile", "Reference"):
                    chosen.append(ra)
        for ra in chosen:
            for item in ra[3] or []:
                out.extend(self.item(item, world))
        return out

    def item(self, ref, m):
        ifc = self.ifc
        t = ifc.type_of(ref)
        a = ifc.args(ref)
        if t == "IFCMAPPEDITEM":
            source = ifc.args(a[0])
            origin = ifc.axis_placement(source[0])
            target = ifc.transform_operator(a[1])
            mm = mat_mul(m, mat_mul(target, origin))
            out = []
            for item in ifc.args(source[1])[3] or []:
                out.extend(self.item(item, mm))
            return out
        if t in ("IFCBOOLEANCLIPPINGRESULT", "IFCBOOLEANRESULT"):
            first = self.item(a[1], m)
            if a[0] == "DIFFERENCE" and len(first) == 1:
                clip = self._clip(a[2], m)
                if clip is not None and isinstance(first[0], Extrusion):
                    first[0].clips.append(clip)
                    return first
            self._skip(t + " " + str(a[0]))
            return first
        if t == "IFCEXTRUDEDAREASOLID":
            prof = self.profiles.read(a[0])
            if prof is None:
                self._skip(ifc.type_of(a[0]) or "profile")
                return []
            pos = mat_mul(m, ifc.axis_placement(a[1]))
            d_local = ifc.direction(a[2]) or (0.0, 0.0, 1.0)
            direction = _norm3(mat_vector(pos, d_local))
            depth = float(a[3]) * ifc.mm
            scale = sum(c * c for c in pos[2]) ** 0.5   # mapped items may scale
            depth *= scale
            plane = (pos[3], _norm3(pos[0]), _norm3(pos[1]))
            if scale and abs(scale - 1.0) > 1e-6:
                prof = Profile(prof.kind, prof.name,
                               _scale_loops(prof.loops, scale), prof.dims)
            return [Extrusion(prof, plane, direction, depth)]
        if t in ("IFCFACETEDBREP", "IFCFACETEDBREPWITHVOIDS"):
            return [Brep(self._shell(a[0], m))]
        if t in ("IFCSHELLBASEDSURFACEMODEL", "IFCFACEBASEDSURFACEMODEL"):
            faces = []
            for shell in a[0] or []:
                faces.extend(self._shell(shell, m))
            return [Brep(faces)]
        self._skip(t)
        return []

    def _shell(self, shell_ref, m):
        ifc = self.ifc
        faces = []
        sa = ifc.args(shell_ref)
        for face in sa[0] or []:
            loops = []
            outer_first = []
            for bound in ifc.args(face)[0] or []:
                ba = ifc.args(bound)
                loop = ifc.args(ba[0])
                if ifc.type_of(ba[0]) != "IFCPOLYLOOP":
                    continue
                pts = [mat_point(m, ifc.point(p)) for p in loop[0]]
                if ba[1] in ("F", False):
                    pts.reverse()
                if ifc.type_of(bound) == "IFCFACEOUTERBOUND":
                    outer_first.append(pts)
                else:
                    loops.append(pts)
            if not outer_first and loops:
                outer_first.append(loops.pop(0))
            if outer_first:
                faces.append(outer_first + loops)
        return faces

    def _clip(self, ref, m):
        """A DIFFERENCE operand: what to cut away from the solid.
        'keep' is the normal of the side that stays."""
        ifc = self.ifc
        t = ifc.type_of(ref)
        a = ifc.args(ref)
        if t in ("IFCHALFSPACESOLID", "IFCPOLYGONALBOUNDEDHALFSPACE",
                 "IFCBOXEDHALFSPACE"):
            surf = ifc.args(a[0])
            if ifc.type_of(a[0]) != "IFCPLANE":
                return None
            pm = mat_mul(m, ifc.axis_placement(surf[0]))
            normal = _norm3(pm[2])
            # AgreementFlag TRUE: the normal points away from the half-space
            # material, so after the difference that side stays.
            agree = a[1] in ("T", True)
            keep = normal if agree else (-normal[0], -normal[1], -normal[2])
            clip = {"kind": "plane", "origin": pm[3], "keep": keep}
            if t == "IFCPOLYGONALBOUNDEDHALFSPACE":
                bm = mat_mul(m, ifc.axis_placement(a[2]))
                poly = self.profiles.curve(a[3])
                if poly:
                    clip = {"kind": "bounded", "origin": pm[3], "keep": keep,
                            "plane": (bm[3], _norm3(bm[0]), _norm3(bm[1])),
                            "axis": _norm3(bm[2]), "poly": poly}
            return clip
        if t == "IFCEXTRUDEDAREASOLID":
            solids = self.item(ref, m)
            if solids:
                return {"kind": "solid", "solid": solids[0]}
        return None

    def _skip(self, what):
        self.unsupported[what] = self.unsupported.get(what, 0) + 1


def _scale_loops(loops, s):
    out = []
    for loop in loops:
        if isinstance(loop, tuple) and loop and loop[0] == "circle":
            out.append(("circle", loop[1] * s, loop[2] * s, loop[3] * s))
        else:
            out.append([(x * s, y * s) for x, y in loop])
    return out


# ── PARTS AND MEMBERS ───────────────────────────────────────────────────────

class Member(object):
    """A straight profile member: line through the profile's bbox centre."""

    def __init__(self, start, end, xaxis, yaxis, profile):
        self.start = start
        self.end = end
        self.x = xaxis
        self.y = yaxis
        self.profile = profile

    @property
    def length(self):
        d = [self.end[i] - self.start[i] for i in range(3)]
        return (d[0] ** 2 + d[1] ** 2 + d[2] ** 2) ** 0.5

    @property
    def direction(self):
        d = [self.end[i] - self.start[i] for i in range(3)]
        return _norm3(d)


class Part(object):
    def __init__(self, eid, ifc_type):
        self.eid = eid
        self.ifc_type = ifc_type
        self.guid = u""
        self.name = u""
        self.object_type = u""
        self.tag = u""
        self.material = u""
        self.mark = u""
        self.assembly = u""
        self.bolt = False
        self.solids = []
        self.member = None

    @property
    def role(self):
        return MEMBER_TYPES.get(self.ifc_type)

    @property
    def section(self):
        """The section name to match: profile name, else ObjectType, else Name."""
        if self.member is not None and self.member.profile.name.strip():
            return self.member.profile.name.strip()
        for v in (self.object_type, self.name):
            if v and name_key(v):
                return v.strip()
        return (self.object_type or self.name or u"").strip()


def member_of(solids):
    """The Member of a body that is exactly one straight extrusion."""
    if len(solids) != 1 or not isinstance(solids[0], Extrusion):
        return None
    ex = solids[0]
    o, x, y = ex.plane
    z = _norm3((x[1] * y[2] - x[2] * y[1], x[2] * y[0] - x[0] * y[2],
                x[0] * y[1] - x[1] * y[0]))
    if z is None or ex.depth <= 0 or abs(_dot3(ex.dir, z)) < 0.9999:
        return None             # oblique extrusion: not a member line
    cx, cy = ex.profile.centre
    start = (o[0] + cx * x[0] + cy * y[0], o[1] + cx * x[1] + cy * y[1],
             o[2] + cx * x[2] + cy * y[2])
    end = (start[0] + ex.dir[0] * ex.depth, start[1] + ex.dir[1] * ex.depth,
           start[2] + ex.dir[2] * ex.depth)
    return Member(start, end, x, y, ex.profile)


class SteelModel(object):
    def __init__(self, ifc):
        self.ifc = ifc
        self.parts = []
        self.unsupported = {}

    def bounds(self):
        lo = [1e30] * 3
        hi = [-1e30] * 3
        for p in self.parts:
            if p.member is not None:
                for q in (p.member.start, p.member.end):
                    for i in range(3):
                        lo[i] = min(lo[i], q[i])
                        hi[i] = max(hi[i], q[i])
        if lo[0] > hi[0]:
            return None
        return lo, hi


def _text(v):
    if v is None:
        return u""
    if isinstance(v, tuple):
        v = v[1]
    try:
        return unicode(v)       # IronPython 2
    except NameError:
        return str(v)


MARK_NAMES = ["part mark", "piece mark", "part position", "position number",
              "mark", "assembly/cast unit mark", "assembly mark",
              "assembly position", "shipping number"]


def read_steel(path_or_ifc, progress=None):
    ifc = path_or_ifc if isinstance(path_or_ifc, IfcFile) \
        else IfcFile.load(path_or_ifc, progress)
    ifc.rad = _angle_unit(ifc)
    model = SteelModel(ifc)

    # products: whatever a storey contains, and whatever those aggregate
    parent = {}
    children = {}
    for rel in ifc.by_type("IFCRELAGGREGATES"):
        a = ifc.args(rel)
        for c in a[5] or []:
            parent[c] = a[4]
            children.setdefault(a[4], []).append(c)
    products = []
    seen = set()
    stack = []
    for rel in ifc.by_type("IFCRELCONTAINEDINSPATIALSTRUCTURE"):
        stack.extend(ifc.args(rel)[4] or [])
    while stack:
        eid = stack.pop()
        if eid in seen:
            continue
        seen.add(eid)
        products.append(eid)
        stack.extend(children.get(eid, ()))
    products.sort()

    materials = {}
    for rel in ifc.by_type("IFCRELASSOCIATESMATERIAL"):
        a = ifc.args(rel)
        name = _material_name(ifc, a[5])
        for o in a[4] or []:
            materials[o] = name
    psets = {}
    for rel in ifc.by_type("IFCRELDEFINESBYPROPERTIES"):
        a = ifc.args(rel)
        for o in a[4] or []:
            psets.setdefault(o, []).append(a[5])

    profiles = ProfileReader(ifc)
    bodies = BodyReader(ifc, profiles)
    total = len(products)
    for n, eid in enumerate(products):
        t = ifc.type_of(eid)
        if t in NO_BODY_TYPES:
            continue
        a = ifc.args(eid)
        if a is None or len(a) < 7:
            continue
        part = Part(eid, t)
        part.guid = _text(a[0])
        part.name = _text(a[2])
        part.object_type = _text(a[4])
        part.tag = _text(a[7]) if len(a) > 7 else u""
        part.material = materials.get(eid, u"")
        p = parent.get(eid)
        part.mark = _find_mark(ifc, psets.get(eid, ()))
        if p is not None:
            pa = ifc.args(p)
            part.assembly = _text(pa[2]) or _text(pa[7] if len(pa) > 7 else None)
            if not part.material:
                part.material = materials.get(p, u"")
            if not part.mark:
                part.mark = _find_mark(ifc, psets.get(p, ()))
        part.bolt = t in FASTENER_TYPES or \
            (p is not None and ifc.type_of(p) in FASTENER_TYPES)
        try:
            world = ifc.placement(a[5])
            part.solids = bodies.body(a, world)
        except Exception as ex:
            bodies._skip("error: {}".format(ex))
            part.solids = []
        if not part.solids and t == "IFCELEMENTASSEMBLY":
            continue
        part.member = member_of(part.solids)
        if not part.bolt:
            names = [part.name, part.object_type]
            if part.member is not None:
                names.append(part.member.profile.name)
            part.bolt = any(is_bolt_name(v) for v in names)
        if part.solids or part.bolt:
            model.parts.append(part)
        if progress is not None and n % 200 == 0:
            progress(n, total)
    model.unsupported = bodies.unsupported
    return model


def _angle_unit(ifc):
    for ua in ifc.by_type("IFCUNITASSIGNMENT"):
        for ref in ifc.args(ua)[0] or []:
            a = ifc.args(ref)
            t = ifc.type_of(ref)
            if len(a) > 1 and a[1] == "PLANEANGLEUNIT":
                if t == "IFCCONVERSIONBASEDUNIT":
                    return math.pi / 180.0
                return 1.0
    return 1.0


def _material_name(ifc, ref):
    t = ifc.type_of(ref)
    a = ifc.args(ref)
    if t == "IFCMATERIAL":
        return _text(a[0])
    if t == "IFCMATERIALLIST" and a[0]:
        return _material_name(ifc, a[0][0])
    if t == "IFCMATERIALLAYERSETUSAGE":
        return _material_name(ifc, a[0])
    if t == "IFCMATERIALLAYERSET" and a[0]:
        return _material_name(ifc, a[0][0])
    if t == "IFCMATERIALLAYER":
        return _material_name(ifc, a[0])
    if t == "IFCMATERIALPROFILESETUSAGE":
        return _material_name(ifc, a[0])
    if t == "IFCMATERIALPROFILESET" and a[2]:
        return _material_name(ifc, a[2][0])
    if t == "IFCMATERIALPROFILE":
        return _material_name(ifc, a[2])
    return u""


def _find_mark(ifc, pset_refs):
    found = {}
    for ref in pset_refs:
        if ifc.type_of(ref) != "IFCPROPERTYSET":
            continue
        for prop in ifc.args(ref)[4] or []:
            if ifc.type_of(prop) != "IFCPROPERTYSINGLEVALUE":
                continue
            pa = ifc.args(prop)
            key = _text(pa[0]).strip().lower()
            if key in MARK_NAMES and key not in found:
                value = _text(pa[2]).strip()
                if value and not (value.startswith("{") and value.endswith("}")):
                    found[key] = value
    for k in MARK_NAMES:
        if k in found:
            return found[k]
    return u""
