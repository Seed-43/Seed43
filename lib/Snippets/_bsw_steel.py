# -*- coding: utf-8 -*-
"""Steel parts out of a StruCad .bsw file: members, plates, bolts, grids.

The .bsw is StruCad's own binary model (the format reference is bsw_core.js
in Seed43 Studio, which reads the same records). Each record is one part in
its own local frame (origin, X and Y axes) with planar faces made of edge
loops. A member's local Z runs along its length.

read_bsw() reads the file into BswEntity records and grid lines and depends
on nothing else in Seed43. to_steel_model() turns those records into the
same parts _ifc_steel.read_steel() makes for an IFC (Part, Member, Profile,
Extrusion, Brep), so _steel_import can place them as native framing,
columns, plates and bolts the way it does for an IFC.

Everything is in millimetres, in the file's world coordinates. No Revit
types in here, so it runs under CPython for testing as well as IronPython 2.
"""
import math
import struct

__all__ = ["BswError", "BswEntity", "BswGrid", "BswModel", "read_bsw",
           "to_steel_model", "read_steel_bsw", "is_bsw"]

KINDS = {0x4D: "member", 0x50: "plate", 0x43: "component", 0x42: "bolt",
         0x57: "weld"}
TOP_TAGS = (0x42, 0x47, 0x67, 0x76)     # B record, grid set, grid line, end
HEADER = 14                             # int16 4, int32 0, float64 16.0
ARC_STEP = math.radians(10.0)
CIRCLE_SEGMENTS = 32
STRAIGHT_TOL = 2.0                      # mm a member may stray off its line
SLICES = (0.5, 0.4, 0.6, 0.3, 0.7)      # where a cut member's section is read


class BswError(Exception):
    pass


def is_bsw(path):
    """True when the file starts the way a StruCad .bsw does."""
    try:
        with open(path, "rb") as f:
            head = f.read(HEADER)
        return len(head) == HEADER and \
            struct.unpack("<hid", head) == (4, 0, 16.0)
    except (IOError, OSError, struct.error):
        return False


# ── READER ──────────────────────────────────────────────────────────────────

class _Reader(object):
    """Little-endian values off the raw file. data is the str (IronPython 2)
    or bytes (CPython 3) for struct and text; b is a bytearray of the same
    bytes, so one byte reads as an int under both."""

    def __init__(self, data, b, pos=0):
        self.data = data
        self.b = b
        self.n = len(b)
        self.pos = pos

    def peek(self):
        return self.b[self.pos] if self.pos < self.n else -1

    def u8(self):
        v = self.b[self.pos]
        self.pos += 1
        return v

    def _unpack(self, fmt, size):
        if self.pos + size > self.n:
            raise BswError("the file ends inside a record")
        v = struct.unpack_from(fmt, self.data, self.pos)
        self.pos += size
        return v

    def i16(self):
        return self._unpack("<h", 2)[0]

    def u16(self):
        return self._unpack("<H", 2)[0]

    def i32(self):
        return self._unpack("<i", 4)[0]

    def f32(self):
        return self._unpack("<f", 4)[0]

    def f32x3(self):
        return self._unpack("<3f", 12)

    def f64x3(self):
        return self._unpack("<3d", 24)

    def text(self):
        size = self.u16()
        if self.pos + size > self.n:
            raise BswError("the file ends inside a name")
        s = self.data[self.pos:self.pos + size].decode("latin-1")
        self.pos += size
        return s

    def expect(self, tag, what):
        got = self.u8()
        if got != tag:
            raise BswError("expected {} at 0x{:x}, found 0x{:x}".format(
                what, self.pos - 1, got))


# ── RECORDS ─────────────────────────────────────────────────────────────────

class BswEntity(object):
    """One StruCad part. verts are local (x, y, z); edges are (a, b, arc)
    with 0-based vertex indices and arc = (centre, axis, radius) or None;
    faces are (normal, [loop, ...]) where a loop is a list of signed 1-based
    edge numbers (the sign is the direction); holes are (centre, axis,
    radius), all local."""

    def __init__(self, kind, group):
        self.kind = kind
        self.group = group
        self.section = u""
        self.mark = u""
        self.grade = u""
        self.product = u""      # bolts: a chemical anchor's product code
        self.attrs = {}         # members: layer, phase, desc...
        self.size = 0.0         # plate thickness, bolt diameter, weld size
        self.length = None
        self.pattern = None     # bolts: (countA, countB, spacingA, spacingB)
        self.origin = (0.0, 0.0, 0.0)
        self.x = (1.0, 0.0, 0.0)
        self.y = (0.0, 1.0, 0.0)
        self.z = (0.0, 0.0, 1.0)
        self.verts = []
        self.edges = []
        self.faces = []
        self.holes = []
        self.instance = False   # geometry borrowed from an earlier record
        self.bolt_copy = False  # one of a bolt group's repeats
        self.offset = 0         # where the record starts in the file

    def key(self):
        k = (self.kind, self.section, self.mark)
        if self.kind == "bolt":
            k += (self.size, self.length)
        return k

    def world(self, p):
        o, x, y, z = self.origin, self.x, self.y, self.z
        return (o[0] + p[0] * x[0] + p[1] * y[0] + p[2] * z[0],
                o[1] + p[0] * x[1] + p[1] * y[1] + p[2] * z[1],
                o[2] + p[0] * x[2] + p[1] * y[2] + p[2] * z[2])

    def world_dir(self, d):
        x, y, z = self.x, self.y, self.z
        return (d[0] * x[0] + d[1] * y[0] + d[2] * z[0],
                d[0] * x[1] + d[1] * y[1] + d[2] * z[1],
                d[0] * x[2] + d[1] * y[2] + d[2] * z[2])


class BswGrid(object):
    def __init__(self, label, a, b):
        self.label = label
        self.a = a
        self.b = b


class BswModel(object):
    def __init__(self):
        self.entities = []
        self.grids = []
        self.warnings = []      # (offset, what, message)


def _geometry(r, ent):
    r.expect(0x4F, "the origin")
    ent.origin = r.f32x3()
    r.expect(0x58, "the X axis")
    ent.x = r.f32x3()
    r.expect(0x59, "the Y axis")
    ent.y = r.f32x3()
    ent.z = _unit(_cross(ent.x, ent.y)) or (0.0, 0.0, 1.0)
    face = None
    while True:
        t = r.peek()
        if t in TOP_TAGS or t == -1:
            break
        r.pos += 1
        if t == 0x56:                       # V vertex
            ent.verts.append(r.f64x3())
        elif t == 0x53:                     # S straight edge, 1-based
            a = r.i16()
            b = r.i16()
            ent.edges.append((a - 1, b - 1, None))
        elif t == 0x44:                     # D flags
            r.pos += 13
        elif t == 0x50:                     # P face normal: a new face
            face = (r.f32x3(), [])
            ent.faces.append(face)
        elif t == 0x4C:                     # L loop of signed edge numbers
            loop = []
            while True:
                v = r.i16()
                if v == 0:
                    break
                loop.append(v)
            if face is not None:
                face[1].append(loop)
        elif t == 0x45:                     # E arc edge
            c = r.f64x3()
            axis = r.f32x3()
            rad = r.f32()
            a = r.i16()
            b = r.i16()
            ent.edges.append((a - 1, b - 1, (c, axis, rad)))
        elif t == 0x43:                     # C circle: a hole
            c = r.f64x3()
            axis = r.f32x3()
            rad = r.f32()
            r.pos += 4                      # start angle
            ent.holes.append((c, axis, rad))
            face = None                     # its rings follow, not faces
        else:
            raise BswError("unknown geometry tag 0x{:x} at 0x{:x}".format(
                t, r.pos - 1))


def _item(r):
    """One top-level item: ("ent", BswEntity), ("grid", BswGrid), ("skip",)
    or ("end",). Raises when the bytes there do not read as one."""
    start = r.pos
    t = r.peek()
    if t == 0x42:
        r.pos += 1
        kb = r.u8()
        kind = KINDS.get(kb)
        if kind is None:
            raise BswError("unknown record B{} at 0x{:x}".format(
                chr(kb), start))
        ent = BswEntity(kind, r.i32())
        ent.offset = start
        if kind == "member":
            ent.section = r.text()
            ent.mark = r.text()
            ent.grade = r.text()
            while r.peek() == 0x41:         # A name value
                r.pos += 1
                k = r.text().lower()
                ent.attrs[k] = r.text()
        elif kind in ("plate", "component"):
            r.i32()
            ent.section = r.text()
            ent.mark = r.text()
            ent.grade = r.text()
            if kind == "plate":
                ent.size = r.f32()
        elif kind == "bolt":
            r.i32()
            ent.size = r.f32()
            r.f32()
            ent.length = r.f32()
            ent.section = r.text()
            ent.mark = r.text()
            ent.product = r.text()
            na = r.i16()
            nb = r.i16()
            ent.pattern = (na, nb, r.f32(), r.f32())
        elif kind == "weld":
            r.i32()
            ent.section = r.text()
            ent.mark = ent.section
            ent.size = r.f32()
        _geometry(r, ent)
        return ("ent", ent)
    if t == 0x47:                           # grid set header
        r.pos += 14
        return ("skip",)
    if t == 0x67:                           # grid line
        r.pos += 1
        label = r.text()
        return ("grid", BswGrid(label, r.f32x3(), r.f32x3()))
    if t == 0x76:                           # the viewer's trailer
        return ("end",)
    raise BswError("unknown record 0x{:x} at 0x{:x}".format(t, start))


def _resync(data, b, start):
    """After an unreadable record: the next offset where a whole item reads
    and is followed by another item (or the end). -1 when none does."""
    n = len(b)
    p = start
    while p < n - 1:
        t = b[p]
        if (t == 0x42 and b[p + 1] in KINDS) or t == 0x67:
            r = _Reader(data, b, p)
            try:
                _item(r)
                if r.peek() in TOP_TAGS or r.peek() == -1:
                    return p
            except Exception:
                pass
        p += 1
    return -1


def _describe(b, at):
    if b[at] == 0x42 and at + 1 < len(b) and b[at + 1] in KINDS:
        return KINDS[b[at + 1]]
    if b[at] == 0x67:
        return "grid line"
    return "record"


def read_bsw(path, progress=None):
    """Read a .bsw into a BswModel. A record that does not read is skipped
    with a warning and the rest still loads, as in Studio."""
    with open(path, "rb") as f:
        data = f.read()
    try:
        b = bytearray(data)
    except TypeError:                   # a str that wants an encoding
        b = bytearray(data, "latin-1")
    if len(b) < HEADER or struct.unpack_from("<hid", data, 0) != (4, 0, 16.0):
        raise BswError("not a StruCad .bsw file")
    model = BswModel()
    r = _Reader(data, b, HEADER)
    total = r.n
    count = 0
    while r.pos < r.n:
        start = r.pos
        try:
            item = _item(r)
        except Exception as ex:
            nxt = _resync(data, b, start + 1)
            msg = u"{}".format(ex)
            if nxt < 0:
                msg += u" (nothing after it could be read)"
            model.warnings.append((start, _describe(b, start), msg))
            if nxt < 0:
                break
            r.pos = nxt
            continue
        if item[0] == "end":
            break
        if item[0] == "ent":
            model.entities.append(item[1])
        elif item[0] == "grid":
            model.grids.append(item[1])
        count += 1
        if progress is not None and count % 200 == 0:
            progress(r.pos, total)
    data = b = None
    _resolve_instances(model.entities)
    model.entities = _repeat_bolts(model.entities)
    return model


def _resolve_instances(entities):
    """Records with no geometry reuse the first full record with the same
    kind, section and mark (bolts: and size and length)."""
    full = {}
    for e in entities:
        if e.verts and e.key() not in full:
            full[e.key()] = e
    for e in entities:
        if e.verts:
            continue
        s = full.get(e.key())
        if s is not None:
            e.verts, e.edges, e.faces, e.holes = s.verts, s.edges, s.faces, \
                s.holes
            e.instance = True
    for e in entities:
        if e.kind == "member" and e.verts and e.length is None:
            zs = [v[2] for v in e.verts]
            e.length = max(zs) - min(zs)


def _repeat_bolts(entities):
    """A bolt record is one bolt and a grid pattern (countA x countB at
    spacingA along local X and spacingB along local Y); StruCad draws the
    whole group from it, so the copies are made here."""
    out = []
    for e in entities:
        out.append(e)
        if e.kind != "bolt" or not e.pattern:
            continue
        na, nb, sa, sb = e.pattern
        na = na or 1
        nb = nb or 1
        if na * nb <= 1:
            continue
        for i in range(na):
            for j in range(nb):
                if i == 0 and j == 0:
                    continue
                c = BswEntity(e.kind, e.group)
                c.__dict__.update(e.__dict__)
                c.origin = _add(e.origin, _add(_mul(e.x, i * (sa or 0.0)),
                                               _mul(e.y, j * (sb or 0.0))))
                c.instance = True
                c.bolt_copy = True
                out.append(c)
    return out


# ── VECTORS ─────────────────────────────────────────────────────────────────

def _add(a, b):
    return (a[0] + b[0], a[1] + b[1], a[2] + b[2])


def _sub(a, b):
    return (a[0] - b[0], a[1] - b[1], a[2] - b[2])


def _mul(a, s):
    return (a[0] * s, a[1] * s, a[2] * s)


def _dot(a, b):
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


def _cross(a, b):
    return (a[1] * b[2] - a[2] * b[1], a[2] * b[0] - a[0] * b[2],
            a[0] * b[1] - a[1] * b[0])


def _unit(v):
    l = math.sqrt(_dot(v, v))
    return (v[0] / l, v[1] / l, v[2] / l) if l > 1e-12 else None


def _perp(n):
    """Two unit vectors square to n and to each other."""
    a = (1.0, 0.0, 0.0) if abs(n[0]) < 0.9 else (0.0, 1.0, 0.0)
    u = _unit(_cross(n, a))
    return u, _cross(n, u)


def _newell(pts):
    nx = ny = nz = 0.0
    m = len(pts)
    for i in range(m):
        p = pts[i]
        q = pts[(i + 1) % m]
        nx += (p[1] - q[1]) * (p[2] + q[2])
        ny += (p[2] - q[2]) * (p[0] + q[0])
        nz += (p[0] - q[0]) * (p[1] + q[1])
    return (nx, ny, nz)


def _area2d(pts):
    a = 0.0
    m = len(pts)
    for i in range(m):
        x0, y0 = pts[i]
        x1, y1 = pts[(i + 1) % m]
        a += x0 * y1 - x1 * y0
    return a / 2.0


# ── LOOPS ───────────────────────────────────────────────────────────────────

def _arc_points(p, q, arc):
    """Points strictly between p and q along an arc edge, the short way
    round the centre (local coordinates)."""
    c, axis, rad = arc
    n = _unit(axis)
    u = _sub(p, c)
    w = _sub(q, c)
    if n is None or _dot(u, u) < 1e-9:
        return []
    # turn from u to w about n, kept within half a turn
    ang = math.atan2(_dot(_cross(u, w), n), _dot(u, w))
    steps = int(abs(ang) / ARC_STEP)
    if steps < 1:
        return []
    v = _cross(n, u)
    out = []
    for k in range(1, steps + 1):
        t = ang * k / (steps + 1)
        ct, st = math.cos(t), math.sin(t)
        out.append(_add(c, _add(_mul(u, ct), _mul(v, st))))
    return out


class _Ring(list):
    """A loop that is one full circle (a hole ring): circle is (centre,
    radius), local."""
    circle = None


_TRIG = [(math.cos(2 * math.pi * k / CIRCLE_SEGMENTS),
          math.sin(2 * math.pi * k / CIRCLE_SEGMENTS))
         for k in range(CIRCLE_SEGMENTS)]


def _circle(c, axis, rad):
    nn = _unit(axis) or (0.0, 0.0, 1.0)
    u, v = _perp(nn)
    ux, uy, uz = u[0] * rad, u[1] * rad, u[2] * rad
    vx, vy, vz = v[0] * rad, v[1] * rad, v[2] * rad
    ring = _Ring((c[0] + ux * ct + vx * st, c[1] + uy * ct + vy * st,
                  c[2] + uz * ct + vz * st) for ct, st in _TRIG)
    ring.circle = (c, rad)
    return ring


def loop_points(ent, loop):
    """A face loop as local points, arcs followed round. A loop that is one
    full-circle arc (a hole ring) comes back as a circle polygon."""
    pts = []
    for e in loop:
        k = abs(e) - 1
        if k < 0 or k >= len(ent.edges):
            continue
        a, b, arc = ent.edges[k]
        if a == b:
            if arc is not None and len(loop) == 1:
                return _circle(arc[0], arc[1], arc[2])
            continue
        if not (0 <= a < len(ent.verts) and 0 <= b < len(ent.verts)):
            continue
        pa, pb = ent.verts[a], ent.verts[b]
        mid = _arc_points(pa, pb, arc) if arc is not None else []
        if e < 0:
            pa, pb = pb, pa
            mid.reverse()
        if not pts or _dot(_sub(pts[-1], pa), _sub(pts[-1], pa)) > 1e-6:
            pts.append(pa)
        pts.extend(mid)
        pts.append(pb)
    if len(pts) > 1 and _dot(_sub(pts[0], pts[-1]), _sub(pts[0], pts[-1])) < 1e-6:
        pts.pop()
    return pts


def faces_local(ent):
    """[(normal, [outer, inner...])] in local points, each loop turned so
    its own normal agrees with the face's (outer anticlockwise seen from
    outside, inner loops the other way)."""
    out = []
    for n, loops in ent.faces:
        rings = []
        for i, loop in enumerate(loops):
            pts = loop_points(ent, loop)
            if len(pts) < 3:
                continue
            d = _dot(_newell(pts), n)
            if (i == 0 or not rings) == (d < 0):
                pts.reverse()       # in place, so a _Ring stays one
            rings.append(pts)
        if rings:
            out.append((n, rings))
    return out


# ── SHAPES ──────────────────────────────────────────────────────────────────

def _prism(ent, faces, axis):
    """When the part is a straight prism along local axis 0, 1 or 2: (lo,
    depth, outer, inners) with the loops as 2D points in the other two
    axes (right-handed), else None."""
    i = axis
    j, k = (i + 1) % 3, (i + 2) % 3
    vals = [v[i] for v in ent.verts]
    lo, hi = min(vals), max(vals)
    depth = hi - lo
    if depth < 0.1:
        return None
    tol = max(0.05, depth * 1e-5)
    for s in vals:
        if s - lo > tol and hi - s > tol:
            return None
    bot = top = None
    nb = nt = 0
    for n, rings in faces:
        if n[i] < -0.999:
            nb += 1
            bot = rings
        elif n[i] > 0.999:
            nt += 1
            top = rings
        elif abs(n[i]) > 1e-3:
            return None                     # a sloping face: a cut end
    if nb != 1 or nt != 1 or len(bot) != len(top):
        return None

    def keys(rings):
        return set((int(round(p[j] * 10)), int(round(p[k] * 10)))
                   for ring in rings for p in ring)
    if keys(bot) != keys(top):
        return None
    flat = []
    for ring in bot:
        pts = _Ring((p[j], p[k]) for p in ring)
        pts.circle = getattr(ring, "circle", None)
        flat.append(pts)
    outer = flat[0]
    if _area2d(outer) < 0:
        outer.reverse()
    inners = []
    for pts in flat[1:]:
        if _area2d(pts) > 0:
            pts.reverse()
        inners.append(pts)
    return lo, depth, outer, inners


def _triangles(faces):
    """Fan-free ear-clipped triangles of each face's outer loop (local)."""
    tris = []
    for n, rings in faces:
        pts = rings[0]
        ax = [abs(n[0]), abs(n[1]), abs(n[2])]
        w = ax.index(max(ax))
        u, v = (w + 1) % 3, (w + 2) % 3
        flat = [(p[u], p[v]) for p in pts]
        for a, b, c in _earclip(flat):
            tris.append((pts[a], pts[b], pts[c]))
    return tris


def _earclip(pts):
    idx = list(range(len(pts)))
    if len(idx) < 3:
        return []
    if _area2d(pts) < 0:
        idx.reverse()
    tris = []
    guard = 0
    while len(idx) > 3 and guard < 5000:
        guard += 1
        clipped = False
        m = len(idx)
        for i in range(m):
            a, b, c = idx[i - 1], idx[i], idx[(i + 1) % m]
            pa, pb, pc = pts[a], pts[b], pts[c]
            cr = (pb[0] - pa[0]) * (pc[1] - pa[1]) - \
                (pb[1] - pa[1]) * (pc[0] - pa[0])
            if cr <= 1e-9:
                continue
            if any(_in_tri(pts[q], pa, pb, pc) for q in idx
                   if q not in (a, b, c)):
                continue
            tris.append((a, b, c))
            del idx[i]
            clipped = True
            break
        if not clipped:
            for i in range(1, len(idx) - 1):
                tris.append((idx[0], idx[i], idx[i + 1]))
            return tris
    if len(idx) == 3:
        tris.append(tuple(idx))
    return tris


def _in_tri(p, a, b, c):
    s1 = (b[0] - a[0]) * (p[1] - a[1]) - (b[1] - a[1]) * (p[0] - a[0])
    s2 = (c[0] - b[0]) * (p[1] - b[1]) - (c[1] - b[1]) * (p[0] - b[0])
    s3 = (a[0] - c[0]) * (p[1] - c[1]) - (a[1] - c[1]) * (p[0] - c[0])
    neg = s1 < 0 or s2 < 0 or s3 < 0
    pos = s1 > 0 or s2 > 0 or s3 > 0
    return not (neg and pos)


def _slice(tris, z):
    """The closed loops where the plane local z cuts the triangles, as 2D
    (x, y) points, outer (largest) first and anticlockwise."""
    segs = []
    for tri in tris:
        pts = []
        for m in range(3):
            p, q = tri[m], tri[(m + 1) % 3]
            if (p[2] - z) * (q[2] - z) < 0:
                t = (z - p[2]) / (q[2] - p[2])
                pts.append((p[0] + (q[0] - p[0]) * t, p[1] + (q[1] - p[1]) * t))
        if len(pts) == 2:
            segs.append(pts)
    if not segs:
        return []

    def key(p):
        return (int(round(p[0] * 100)), int(round(p[1] * 100)))
    ends = {}
    for s in segs:
        ka, kb = key(s[0]), key(s[1])
        if ka == kb:
            continue
        ends.setdefault(ka, []).append((kb, s[1]))
        ends.setdefault(kb, []).append((ka, s[0]))
    loops = []
    used = set()
    for start in list(ends):
        if start in used:
            continue
        loop = []
        prev = None
        cur = start
        while cur not in used:
            used.add(cur)
            nxt = [e for e in ends.get(cur, []) if e[0] != prev and
                   e[0] not in used]
            if not nxt:
                back = [e for e in ends.get(cur, []) if e[0] == start]
                if back and len(loop) >= 2:
                    loop.append(cur)
                break
            loop.append(cur)
            prev, cur = cur, nxt[0][0]
        if len(loop) >= 3:
            loops.append([(x / 100.0, y / 100.0) for x, y in loop])
    loops.sort(key=lambda l: -abs(_area2d(l)))
    out = []
    for i, l in enumerate(loops):
        a = _area2d(l)
        if (i == 0) == (a < 0):
            l.reverse()
        out.append(l)
    return out


def member_line(ent, faces):
    """(lo, depth, outer, inners) for a member: the section and its extent
    along local Z. A straight prism reads exactly; a member with cut ends
    (copes, notches, mitres) takes the fullest section across its middle,
    as long as nothing strays off that section's box (so a curved or
    haunched member is not read as straight). None when neither works."""
    p = _prism(ent, faces, 2)
    if p is not None:
        return p
    zs = [v[2] for v in ent.verts]
    lo, hi = min(zs), max(zs)
    depth = hi - lo
    if depth < 1.0:
        return None
    tris = _triangles(faces)
    best = None
    for f in SLICES:
        loops = _slice(tris, lo + depth * f)
        if loops:
            a = abs(_area2d(loops[0])) - sum(abs(_area2d(l)) for l in loops[1:])
            if best is None or a > best[0] + 1e-6:
                best = (a, loops)
    if best is None:
        return None
    loops = best[1]
    xs = [q[0] for q in loops[0]]
    ys = [q[1] for q in loops[0]]
    x0, x1, y0, y1 = min(xs), max(xs), min(ys), max(ys)
    tol = max(STRAIGHT_TOL, 0.01 * max(x1 - x0, y1 - y0))
    for v in ent.verts:
        if v[0] < x0 - tol or v[0] > x1 + tol or \
                v[1] < y0 - tol or v[1] > y1 + tol:
            return None
    # only loops inside the outer one are hollows (an open web hole is not)
    inners = [l for l in loops[1:] if all(x0 <= q[0] <= x1 and y0 <= q[1] <= y1
                                          for q in l)]
    return lo, depth, loops[0], inners


# ── STEEL MODEL ─────────────────────────────────────────────────────────────

def _steel_types():
    try:
        from Snippets import _ifc_steel as S
    except ImportError:                 # run outside the package (tests)
        import _ifc_steel as S
    return S


class _Source(object):
    """Stands in for the IfcFile the importer reports on."""
    schema = u"StruCad BSW"
    mm = 1.0


def _role(ent, direction):
    """IFCCOLUMN, IFCBEAM or IFCMEMBER (brace) for a member: its description
    when that says, else how it lies."""
    desc = u" ".join(ent.attrs.get(k, u"") for k in ("desc", "description",
                                                     "type", "layer"))
    d = desc.upper()
    if u"COL" in d or u"POST" in d:
        return "IFCCOLUMN"
    if u"BRAC" in d:
        return "IFCMEMBER"
    if u"BEAM" in d or u"RAFT" in d or u"PURLIN" in d or u"GIRT" in d:
        return "IFCBEAM"
    c = abs(direction[2])
    if c > 0.996:                       # within 5 degrees of vertical
        return "IFCCOLUMN"
    if c < 0.087:                       # within 5 degrees of level
        return "IFCBEAM"
    return "IFCMEMBER"


def _brep(S, ent, faces):
    return S.Brep([[[ent.world(p) for p in ring] for ring in rings]
                   for n, rings in faces])


def _member_part(S, part, ent, faces):
    line = member_line(ent, faces)
    if line is None:
        part.ifc_type = "IFCMEMBER"
        part.solids = [_brep(S, ent, faces)]
        return
    lo, depth, outer, inners = line
    prof = S.Profile("ARB", ent.section, [outer] + inners)
    plane = (ent.world((0.0, 0.0, lo)), ent.x, ent.y)
    exact = _prism(ent, faces, 2) is not None
    if exact:
        part.solids = [S.Extrusion(prof, plane, ent.z, depth)]
    else:
        # the real shape (copes and all) for when it is made as a shape;
        # the importer reads end cuts off .clips, and there are none to add
        body = _brep(S, ent, faces)
        body.clips = []
        part.solids = [body]
    cx, cy = prof.centre
    start = ent.world((cx, cy, lo))
    end = ent.world((cx, cy, lo + depth))
    part.member = S.Member(start, end, ent.x, ent.y, prof)
    part.ifc_type = _role(ent, ent.z)


def _hole_opening(S, ent, c, axis, rad, through):
    """A round hole as an opening: a circle swept along its axis right
    through the part (through = how far either side of the centre)."""
    n = _unit(ent.world_dir(axis))
    if n is None:
        return None
    u, v = _perp(n)
    prof = S.Profile("C", u"HOLE", [("circle", 0.0, 0.0, rad)],
                     {"diameter": 2 * rad})
    o = _sub(ent.world(c), _mul(n, through))
    return S.Extrusion(prof, (o, u, v), n, 2 * through)


def _plate_part(S, part, ent, faces):
    best = None
    for axis in (2, 0, 1):
        p = _prism(ent, faces, axis)
        if p is not None and (best is None or p[1] < best[1][1]):
            best = (axis, p)
    span = 1.0
    if ent.verts:
        span = max(max(v[i] for v in ent.verts) - min(v[i] for v in ent.verts)
                   for i in range(3)) or 1.0
    if best is None:
        part.solids = [_brep(S, ent, faces)]
    else:
        axis, (lo, depth, outer, inners) = best
        frame = (ent.x, ent.y, ent.z)
        d = frame[axis]
        u, v = frame[(axis + 1) % 3], frame[(axis + 2) % 3]
        local = [0.0, 0.0, 0.0]
        local[axis] = lo
        prof = S.Profile("R" if len(outer) == 4 else "ARB", ent.section,
                         [outer])
        part.solids = [S.Extrusion(prof, (ent.world(local), u, v), d, depth)]
        # inner loops through the thickness are cut-outs; hole rings are
        # left to the C hole records, which hold the true circle
        for pts in inners:
            if getattr(pts, "circle", None) is not None and ent.holes:
                continue
            ip = S.Profile("ARB", u"CUT", [list(reversed(pts))])
            part.openings.append(S.Extrusion(ip, (ent.world(local), u, v),
                                             d, depth))
    for c, axis, rad in ent.holes:
        op = _hole_opening(S, ent, c, axis, rad, span)
        if op is not None:
            part.openings.append(op)


def _bolt_part(S, part, ent, faces):
    """A bolt: a round shank along the bolt line (which the importer reads
    the bolt's axis from), plus the faceted head and nut when the record
    has them. StruCad's bolt runs along its local Z."""
    part.bolt = True
    name = ent.section or u""
    part.props = {u"Bolt Name": name,
                  u"_Description": u"M{:g} x {:g} {}".format(
                      ent.size or 0, ent.length or 0, name).strip(),
                  u"_NominalDiameter": ent.size or 0.0,
                  u"_NominalLength": ent.length or 0.0,
                  u"Bolt standard": name}
    if ent.product:
        part.props[u"Product"] = ent.product
        part.props[u"_Description"] += u" " + ent.product
    solids = []
    if ent.verts:
        zs = [v[2] for v in ent.verts]
        lo, hi = min(zs), max(zs)
    else:
        lo, hi = 0.0, ent.length or 0.0
    if hi - lo > 0.1 and ent.size:
        prof = S.Profile("C", name, [("circle", 0.0, 0.0, ent.size / 2.0)],
                         {"diameter": ent.size})
        solids.append(S.Extrusion(prof, (ent.world((0.0, 0.0, lo)), ent.x,
                                         ent.y), ent.z, hi - lo))
    if faces:
        solids.append(_brep(S, ent, faces))
    part.solids = solids


def to_steel_model(bsw, progress=None):
    """The BswModel as an _ifc_steel.SteelModel, for _steel_import.

    Part types follow the IFC ones the importer knows: members become
    IFCBEAM, IFCCOLUMN or IFCMEMBER, plates IFCPLATE, bolts
    IFCMECHANICALFASTENER and components IFCBUILDINGELEMENTPROXY. Welds
    carry no solid and are counted in model.unsupported. The grid lines
    are kept on model.grids."""
    S = _steel_types()
    model = S.SteelModel(_Source())
    shared = {}             # id(faces list): local faces, for instances
    model.grids = bsw.grids
    model.warnings = bsw.warnings
    unsupported = {}
    total = len(bsw.entities)
    for i, ent in enumerate(bsw.entities):
        if ent.kind == "weld":
            unsupported["weld"] = unsupported.get("weld", 0) + 1
            continue
        part = S.Part(i + 1, "IFCBUILDINGELEMENTPROXY")
        part.guid = u"BSW-{:x}-{}".format(ent.offset, i + 1)
        part.name = ent.mark or ent.section
        part.object_type = ent.section
        part.tag = ent.mark
        part.mark = ent.mark
        part.material = ent.grade
        part.assembly = u"{}".format(ent.group) if ent.group else u""
        part.assembly_id = ent.group or None
        part.props = None
        try:
            faces = []
            if ent.verts:
                faces = shared.get(id(ent.faces))
                if faces is None:
                    faces = shared[id(ent.faces)] = faces_local(ent)
            if ent.kind == "member" and faces:
                _member_part(S, part, ent, faces)
            elif ent.kind == "plate" and faces:
                part.ifc_type = "IFCPLATE"
                _plate_part(S, part, ent, faces)
            elif ent.kind == "bolt":
                part.ifc_type = "IFCMECHANICALFASTENER"
                _bolt_part(S, part, ent, faces)
            elif faces:
                part.solids = [_brep(S, ent, faces)]
        except Exception as ex:
            key = u"{} error: {}".format(ent.kind, ex)
            unsupported[key] = unsupported.get(key, 0) + 1
            continue
        if not part.solids and not part.bolt:
            unsupported["{} with no geometry".format(ent.kind)] = \
                unsupported.get("{} with no geometry".format(ent.kind), 0) + 1
            continue
        model.parts.append(part)
        if progress is not None and i % 200 == 0:
            progress(i, total)
    model.unsupported = unsupported
    return model


def read_steel_bsw(path, progress=None):
    """read_bsw then to_steel_model: what _steel_import gets from
    _ifc_steel.read_steel for an IFC."""
    return to_steel_model(read_bsw(path), progress)
