# -*- coding: utf-8 -*-
"""Repeated custom shapes out of a steel IFC as Generic Model families.

Parts the importer would make as shapes (posts, bends, cast pieces...) are
compared by geometry. Each distinct shape becomes one
face-based Generic Model family holding the IFC's own geometry, and every
copy is an instance of it, placed and turned to where that copy is. Users
can then open the family and remodel it once for all the copies.

    groups = repeated_shapes(parts)            # pure Python
    make_shape_families(doc, mapper, groups)   # Revit

Matching: each part is put in its own frame (centre of its corners and the
principal axes of their spread, signs fixed by which way they lean); two
parts are the same shape when their corners agree in that frame to within
MATCH_MM. Placement measured in Revit 2026: a face-based family placed on a
reference plane through origin O with normal N and reference direction X
has its X along X and its Z along N; Revit makes a reference plane's normal
cutVec x (freeEnd - bubbleEnd).

The matching uses no Revit types, so it runs under CPython for testing as
well as IronPython 2.
"""
import math
import os
import re

try:
    from Snippets._ifc_steel import Extrusion, Brep, loop_points
except ImportError:                     # run outside the package (tests)
    from _ifc_steel import Extrusion, Brep, loop_points

try:
    from Autodesk.Revit import DB
except ImportError:                     # CPython tests
    DB = None

__all__ = ["shape_frame", "repeated_shapes", "make_shape_families",
           "ShapeFamilyResult"]

MM_PER_FT = 304.8
MATCH_MM = 1.5          # corners this close in the part's own frame: same shape
MIN_COPIES = 1          # every custom shape gets a family (copies share it)
FAMILY_PREFIX = "Seed43"
CIRCLE_SIDES = 24


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


def _len(a):
    return _dot(a, a) ** 0.5


def _unit(a):
    l = _len(a)
    return (a[0] / l, a[1] / l, a[2] / l) if l > 1e-12 else None


def _text(v):
    try:
        return unicode(v or u"")        # IronPython 2
    except NameError:
        return str(v or "")


# ── GEOMETRY OF A PART ──────────────────────────────────────────────────────

def part_faces(part):
    """The part's surface as faces (each a list of loops of world points,
    mm), or None when it is not one uncut solid the family can copy."""
    if len(part.solids) != 1:
        return None
    s = part.solids[0]
    if isinstance(s, Brep):
        return s.faces
    if not isinstance(s, Extrusion) or s.profile is None or s.clips:
        return None                     # cut extrusions stay shapes
    o, x, y = s.plane
    top = _mul(s.dir, s.depth)
    loops = []
    for loop in s.profile.loops:
        if isinstance(loop, tuple) and loop and loop[0] == "circle":
            cx, cy, r = loop[1], loop[2], loop[3]
            pts2 = [(cx + r * math.cos(2 * math.pi * i / CIRCLE_SIDES),
                     cy + r * math.sin(2 * math.pi * i / CIRCLE_SIDES))
                    for i in range(CIRCLE_SIDES)]
        else:
            pts2 = list(loop)
        loops.append([_add(o, _add(_mul(x, a), _mul(y, b))) for a, b in pts2])
    faces = [[list(l) for l in loops],
             [[_add(q, top) for q in reversed(l)] for l in loops]]
    for l in loops:
        for i in range(len(l)):
            a, b = l[i], l[(i + 1) % len(l)]
            faces.append([[a, b, _add(b, top), _add(a, top)]])
    return faces


def _jacobi(a):
    a = [row[:] for row in a]
    v = [[1.0, 0.0, 0.0], [0.0, 1.0, 0.0], [0.0, 0.0, 1.0]]
    for _ in range(60):
        p, q = 0, 1
        big = abs(a[0][1])
        for i, j in ((0, 2), (1, 2)):
            if abs(a[i][j]) > big:
                p, q, big = i, j, abs(a[i][j])
        if big < 1e-12:
            break
        theta = (a[q][q] - a[p][p]) / (2.0 * a[p][q])
        t = (1.0 if theta >= 0 else -1.0) / (abs(theta) + (theta * theta + 1.0) ** 0.5)
        c = 1.0 / (t * t + 1.0) ** 0.5
        s = t * c
        for k in range(3):
            akp, akq = a[k][p], a[k][q]
            a[k][p] = c * akp - s * akq
            a[k][q] = s * akp + c * akq
        for k in range(3):
            apk, aqk = a[p][k], a[q][k]
            a[p][k] = c * apk - s * aqk
            a[q][k] = s * apk + c * aqk
        for k in range(3):
            vkp, vkq = v[k][p], v[k][q]
            v[k][p] = c * vkp - s * vkq
            v[k][q] = s * vkp + c * vkq
    return [a[i][i] for i in range(3)], [(v[0][i], v[1][i], v[2][i]) for i in range(3)]


def corners(faces):
    seen = {}
    for f in faces:
        for loop in f:
            for q in loop:
                seen[(round(q[0], 2), round(q[1], 2), round(q[2], 2))] = q
    return list(seen.values())


def shape_frame(pts):
    """(origin, x, y, z): the centre of the corners and their principal
    axes, largest spread first, each pointing the way the corners lean
    (z = x cross y, so the frame is never mirrored)."""
    n = float(len(pts))
    c = tuple(sum(p[i] for p in pts) / n for i in range(3))
    cov = [[0.0] * 3 for _ in range(3)]
    for p in pts:
        d = _sub(p, c)
        for i in range(3):
            for j in range(3):
                cov[i][j] += d[i] * d[j]
    vals, vecs = _jacobi(cov)
    order = sorted(range(3), key=lambda i: -vals[i])
    axes = []
    for k in order[:2]:
        a = _unit(vecs[k]) or (1.0, 0.0, 0.0)
        skew = sum(_dot(_sub(p, c), a) ** 3 for p in pts)
        if skew < 0:
            a = _mul(a, -1.0)
        axes.append(a)
    x, y = axes
    z = _unit(_cross(x, y)) or (0.0, 0.0, 1.0)
    y = _cross(z, x)
    return c, x, y, z


def _local(pts, frame):
    o, x, y, z = frame
    out = []
    for p in pts:
        d = _sub(p, o)
        out.append((_dot(d, x), _dot(d, y), _dot(d, z)))
    return out


def _flips(frame):
    """The frame and its three half-turns about its own axes."""
    o, x, y, z = frame
    m = lambda v: _mul(v, -1.0)
    return [frame, (o, x, m(y), m(z)), (o, m(x), y, m(z)), (o, m(x), m(y), z)]


def _key(local):
    return tuple(sorted((int(round(p[0] / MATCH_MM)), int(round(p[1] / MATCH_MM)),
                         int(round(p[2] / MATCH_MM))) for p in local))


def _same(a, b):
    """True when two corner sets agree to within MATCH_MM (any order)."""
    if len(a) != len(b):
        return False
    sa = sorted(a)
    sb = sorted(b)
    for p, q in zip(sa, sb):
        if abs(p[0] - q[0]) > MATCH_MM or abs(p[1] - q[1]) > MATCH_MM or \
                abs(p[2] - q[2]) > MATCH_MM:
            break
    else:
        return True
    # sorting can pair points differently near a tie; check the slow way
    tol2 = (MATCH_MM * 1.75) ** 2
    used = [False] * len(sb)
    for p in sa:
        hit = False
        for k, q in enumerate(sb):
            if not used[k] and (p[0] - q[0]) ** 2 + (p[1] - q[1]) ** 2 + \
                    (p[2] - q[2]) ** 2 <= tol2:
                used[k] = hit = True
                break
        if not hit:
            return False
    return True


def repeated_shapes(parts, min_copies=MIN_COPIES):
    """[(master part, master frame, master faces, [(part, frame), ...])]
    for shapes that occur at least min_copies times; the list of copies
    includes the master. Parts that are no copyable solid are left out."""
    by_count = {}
    for part in parts:
        faces = part_faces(part)
        if not faces:
            continue
        pts = corners(faces)
        if len(pts) < 4:
            continue
        frame = shape_frame(pts)
        local = _local(pts, frame)
        by_count.setdefault((len(faces), len(pts)), []).append(
            (part, frame, faces, pts, local))
    groups = []
    for key in sorted(by_count):
        pending = by_count[key]
        while pending:
            master = pending.pop(0)
            copies = [(master[0], master[1])]
            rest = []
            for cand in pending:
                match = None
                for fr in _flips(cand[1]):
                    if _same(master[4], _local(cand[3], fr)):
                        match = fr
                        break
                if match is not None:
                    copies.append((cand[0], match))
                else:
                    rest.append(cand)
            pending = rest
            if len(copies) >= min_copies:
                groups.append((master[0], master[1], master[2], copies))
    return groups


# ── REVIT ───────────────────────────────────────────────────────────────────

class ShapeFamilyResult(object):
    def __init__(self):
        self.families = 0
        self.instances = 0
        self.placed = set()         # id() of parts placed as family instances
        self.solid = 0              # families holding a free-form solid
        self.mesh = 0               # families holding a shape element
        self.errors = []


def _xyz(t):
    return DB.XYZ(t[0], t[1], t[2])


def _template(app):
    roots = [app.FamilyTemplatePath, os.path.join(
        os.environ.get("ProgramData", "C:\\ProgramData"), "Autodesk",
        "RVT {}".format(app.VersionNumber), "Family Templates")]
    for root in roots:
        if not root or not os.path.isdir(root):
            continue
        for folder, dirs, files in os.walk(root):
            for f in files:
                low = f.lower()
                if low.endswith(".rft") and "generic model" in low and \
                        "face based" in low:
                    return os.path.join(folder, f)
    return None


def _quiet(t):
    class Quiet(DB.IFailuresPreprocessor):
        def PreprocessFailures(self, accessor):
            for f in accessor.GetFailureMessages():
                if f.GetSeverity() == DB.FailureSeverity.Warning:
                    accessor.DeleteWarning(f)
            return DB.FailureProcessingResult.Continue
    opts = t.GetFailureHandlingOptions()
    opts.SetFailuresPreprocessor(Quiet())
    t.SetFailureHandlingOptions(opts)


def _geometry(faces, frame, solid):
    """The faces in the shape's own frame (feet) through the tessellated
    shape builder: a closed solid, or (solid=False) whatever it can make."""
    from System.Collections.Generic import List, IList
    o, x, y, z = frame
    b = DB.TessellatedShapeBuilder()
    if solid:
        b.Target = DB.TessellatedShapeBuilderTarget.Solid
        b.Fallback = DB.TessellatedShapeBuilderFallback.Abort
    else:
        b.Target = DB.TessellatedShapeBuilderTarget.AnyGeometry
        b.Fallback = DB.TessellatedShapeBuilderFallback.Mesh
    b.OpenConnectedFaceSet(solid)
    for face in faces:
        loops = List[IList[DB.XYZ]]()
        for loop in face:
            pts = List[DB.XYZ]()
            last = None
            for q in loop:
                d = _sub(q, o)
                l = (_dot(d, x) / MM_PER_FT, _dot(d, y) / MM_PER_FT,
                     _dot(d, z) / MM_PER_FT)
                if last is None or _len(_sub(l, last)) > 1e-6:
                    pts.Add(_xyz(l))
                    last = l
            if pts.Count >= 3:
                loops.Add(pts)
        if loops.Count:
            b.AddFace(DB.TessellatedFace(loops, DB.ElementId.InvalidElementId))
    b.CloseConnectedFaceSet()
    b.Build()
    return list(b.GetBuildResult().GetGeometricalObjects())


def _build(doc, template, name, faces, frame, result):
    app = doc.Application
    fd = app.NewFamilyDocument(template)
    try:
        t = DB.Transaction(fd, "shape")
        _quiet(t)
        t.Start()
        made = False
        try:
            solids = [g for g in _geometry(faces, frame, True)
                      if isinstance(g, DB.Solid)]
            if solids:
                DB.FreeFormElement.Create(fd, solids[0])
                result.solid += 1
                made = True
        except Exception:
            made = False
        if not made:
            from System.Collections.Generic import List
            geo = _geometry(faces, frame, False)
            if geo:
                ds = DB.DirectShape.CreateElement(
                    fd, DB.ElementId(DB.BuiltInCategory.OST_GenericModel))
                ds.SetShape(List[DB.GeometryObject](geo))
                result.mesh += 1
                made = True
        t.Commit()
        if not made:
            return None
        folder = os.path.join(os.environ.get("TEMP", "."), "Seed43")
        if not os.path.isdir(folder):
            os.makedirs(folder)
        path = os.path.join(folder, name + u".rfa")
        opts = DB.SaveAsOptions()
        opts.OverwriteExistingFile = True
        fd.SaveAs(path, opts)
        fam = fd.LoadFamily(doc)
    finally:
        fd.Close(False)
    ids = list(fam.GetFamilySymbolIds()) if fam is not None else []
    return doc.GetElement(ids[0]) if ids else None


def _unique_name(doc, base):
    base = re.sub(r"[\\:{}\[\]|;<>?`~/]", u"", base).strip() or FAMILY_PREFIX
    names = set(f.Name for f in DB.FilteredElementCollector(doc).OfClass(DB.Family))
    if base not in names:
        return base
    i = 2
    while u"{} ({})".format(base, i) in names:
        i += 1
    return u"{} ({})".format(base, i)


def _view3d(doc, start_transaction):
    for v in DB.FilteredElementCollector(doc).OfClass(DB.View3D):
        if not v.IsTemplate:
            return v
    t = start_transaction(doc, u"Steel IFC view")
    vft = [x for x in DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType)
           if x.ViewFamily == DB.ViewFamily.ThreeDimensional]
    v = DB.View3D.CreateIsometric(doc, vft[0].Id) if vft else None
    t.Commit()
    return v


def make_shape_families(doc, mapper, groups, start_transaction, set_marks,
                        progress=None):
    """One family per group and one instance per copy. Call between
    transactions (families load with none open). Returns a
    ShapeFamilyResult; parts not in .placed stay shapes."""
    result = ShapeFamilyResult()
    template = _template(doc.Application)
    if template is None:
        result.errors.append(u"no face-based Generic Model template found")
        return result
    view = _view3d(doc, start_transaction)
    if view is None:
        result.errors.append(u"no 3D view to host the families in")
        return result
    built = []
    for i, (master, frame, faces, copies) in enumerate(groups):
        if progress is not None:
            progress(i, len(groups), master)
        name = _unique_name(doc, u"{} {}".format(
            FAMILY_PREFIX, master.section or master.name or u"Shape"))
        try:
            sym = _build(doc, template, name, faces, frame, result)
        except Exception as ex:
            sym = None
            result.errors.append(u"{}: {}".format(name, ex))
        if sym is not None:
            result.families += 1
            built.append((sym, copies))
    t = start_transaction(doc, u"Steel IFC shape families")
    for sym, copies in built:
        if not sym.IsActive:
            sym.Activate()
        for part, (o, x, y, z) in copies:
            try:
                O = _xyz(mapper.point(o))
                X = _unit(mapper.vector(x))
                N = _unit(mapper.vector(z))
                rp = doc.Create.NewReferencePlane(
                    O, _xyz(_add(mapper.point(o), X)), _xyz(_cross(X, N)), view)
                inst = doc.Create.NewFamilyInstance(rp.GetReference(), O,
                                                    _xyz(X), sym)
                set_marks(inst, part)
                result.instances += 1
                result.placed.add(id(part))
            except Exception as ex:
                result.errors.append(u"placing {}: {}".format(
                    part.section or part.name, ex))
    t.Commit()
    return result
