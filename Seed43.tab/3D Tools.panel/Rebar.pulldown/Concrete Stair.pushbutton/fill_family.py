# -*- coding: utf-8 -*-
"""The concrete fill between a stair's soffit, its base and the footing face.

When the footing reaches past where the soffit meets the base, a triangle of
air is left under the stair. It is filled by an instance of the loadable
family 'Stair Footing Fill' (Structural Foundations), built here from the
family template the first time it is needed. Each stair gets its own type,
named by the stair's element id, sized by the type parameters Fill Run, Fill Rise and
Fill Width; updating the stair replaces it.

Family geometry, in its own coordinates (feet): the right angle at the
origin, the run along -X, the rise up +Z, the width across Y about the
centre. A plain triangle: the two sharp corners are driven by locking their
edges (which run across the width) to the Run and Rise reference planes.
"""
import math
import os

from Autodesk.Revit import DB
from Autodesk.Revit.DB.Structure import StructuralType

__all__ = ["FAMILY", "FOOTING", "ensure_family", "place", "remove", "fill_size", "fill_shape",
           "place_footing", "is_footing", "footing_name"]

# ── CONSTANTS ───────────────────────────────────────────────────────────────

FAMILY = u"Stair Footing Fill"
FOOTING = u"Stair Footing"
MM = 304.8


# ── HELPERS (private to this file) ──────────────────────────────────────────

def _name(element):
    return DB.Element.Name.GetValue(element)


def _template(app):
    """The best family template: Structural Foundation, else Generic Model,
    metric first. None when the templates are not installed."""
    best = None
    root = app.FamilyTemplatePath
    if not root or not os.path.isdir(root):
        return None
    for folder, _, files in os.walk(root):
        for f in files:
            low = f.lower()
            if not low.endswith(".rft"):
                continue
            if low.endswith("structural foundation.rft"):
                rank = 0
            elif low.endswith("generic model.rft"):
                rank = 2
            else:
                continue
            if not low.startswith("metric"):
                rank += 1
            if best is None or rank < best[0]:
                best = (rank, os.path.join(folder, f))
    return best[1] if best else None


def _length_param(fm, name, instance=False):
    if hasattr(DB, "GroupTypeId"):
        return fm.AddParameter(name, DB.GroupTypeId.Geometry, DB.SpecTypeId.Length, instance)
    return fm.AddParameter(name, DB.BuiltInParameterGroup.PG_GEOMETRY,
                           DB.ParameterType.Length, instance)


def _yes_no_param(fm, name):
    if hasattr(DB, "GroupTypeId"):
        return fm.AddParameter(name, DB.GroupTypeId.Geometry, DB.SpecTypeId.Boolean.YesNo, False)
    return fm.AddParameter(name, DB.BuiltInParameterGroup.PG_GEOMETRY,
                           DB.ParameterType.YesNo, False)


def _material_param(fm, name):
    if hasattr(DB, "GroupTypeId"):
        return fm.AddParameter(name, DB.GroupTypeId.Materials,
                               DB.SpecTypeId.Reference.Material, False)
    return fm.AddParameter(name, DB.BuiltInParameterGroup.PG_MATERIALS,
                           DB.ParameterType.Material, False)


def _centre_planes(fdoc):
    """The template's centre planes: (left/right, front/back)."""
    lr = fb = None
    for p in DB.FilteredElementCollector(fdoc).OfClass(DB.ReferencePlane):
        n, o = p.Normal, p.BubbleEnd
        if abs(abs(n.X) - 1) < 1e-6 and abs(o.X) < 1e-6:
            lr = lr or p
        elif abs(abs(n.Y) - 1) < 1e-6 and abs(o.Y) < 1e-6:
            fb = fb or p
    return lr, fb


def _dimension(fdoc, view, refs, p0, p1, label=None, equal=False):
    arr = DB.ReferenceArray()
    for r in refs:
        arr.Append(r)
    dim = fdoc.FamilyCreate.NewDimension(view, DB.Line.CreateBound(p0, p1), arr)
    if label is not None:
        dim.FamilyLabel = label
    if equal:
        dim.AreSegmentsEqual = True
    return dim


def _faces(element, view):
    opts = DB.Options()
    opts.ComputeReferences = True
    opts.View = view
    for g in element.get_Geometry(opts):
        if isinstance(g, DB.Solid):
            for f in g.Faces:
                if isinstance(f, DB.PlanarFace):
                    yield f


def _align(fdoc, view, element, plane, normal, at):
    """Lock the face of `element` with this outward normal, lying on `at`
    along that normal, to the reference plane."""
    for f in _faces(element, view):
        if f.FaceNormal.IsAlmostEqualTo(normal) and \
                abs(f.Origin.DotProduct(normal) - at) < 1e-4:
            fdoc.FamilyCreate.NewAlignment(view, plane.GetReference(), f.Reference)
            return
    raise ValueError(u"No face to lock to plane {0}".format(plane.Name))


def _align_edge(fdoc, view, element, plane, x, z):
    """Lock the straight edge of `element` running across the width (along
    Y) at (x, z) to the reference plane. The view must show that edge as a
    line: plan for an edge at a given x, a side elevation for a given z."""
    opts = DB.Options()
    opts.ComputeReferences = True
    opts.View = view
    for g in element.get_Geometry(opts):
        if not isinstance(g, DB.Solid):
            continue
        for e in g.Edges:
            c = e.AsCurve()
            if not isinstance(c, DB.Line) or abs(abs(c.Direction.Y) - 1) > 1e-6:
                continue
            a = c.GetEndPoint(0)
            if abs(a.X - x) < 1e-4 and abs(a.Z - z) < 1e-4:
                fdoc.FamilyCreate.NewAlignment(view, plane.GetReference(), e.Reference)
                return
    raise ValueError(u"No edge to lock to plane {0}".format(plane.Name))


# ── CORE LOGIC ──────────────────────────────────────────────────────────────

def build(app, run=1.0, rise=0.5, width=3.0):
    """Make the family document (unsaved, left open). Lengths in feet."""
    template = _template(app)
    if template is None:
        raise ValueError(u"No Structural Foundation or Generic Model family template "
                         u"was found, so the footing fill could not be made.")
    fdoc = app.NewFamilyDocument(template)
    t = DB.Transaction(fdoc, u"Build")
    try:
        t.Start()
        fam_cat = DB.Category.GetCategory(fdoc, DB.BuiltInCategory.OST_StructuralFoundation)
        if fdoc.OwnerFamily.FamilyCategory.Id != fam_cat.Id:
            fdoc.OwnerFamily.FamilyCategory = fam_cat
        fm = fdoc.FamilyManager
        p_run = _length_param(fm, u"Fill Run")
        p_rise = _length_param(fm, u"Fill Rise")
        p_width = _length_param(fm, u"Fill Width")
        # NOTE: tying the solid to the template's Structural Material fails
        # regeneration, so the family carries its own material parameter
        p_mat = _material_param(fm, u"Fill Material")
        if fm.CurrentType is None:
            fm.NewType(u"Default")
        fm.Set(p_run, run)
        fm.Set(p_rise, rise)
        fm.Set(p_width, width)

        plan = [v for v in DB.FilteredElementCollector(fdoc).OfClass(DB.ViewPlan)
                if not v.IsTemplate and v.ViewType == DB.ViewType.FloorPlan][0]
        front = [v for v in DB.FilteredElementCollector(fdoc).OfClass(DB.View)
                 if not v.IsTemplate and v.ViewType == DB.ViewType.Elevation
                 and v.ViewDirection.IsAlmostEqualTo(DB.XYZ(0, -1, 0))][0]
        side = [v for v in DB.FilteredElementCollector(fdoc).OfClass(DB.View)
                if not v.IsTemplate and v.ViewType == DB.ViewType.Elevation
                and v.ViewDirection.IsAlmostEqualTo(DB.XYZ(1, 0, 0))][0]
        level = DB.FilteredElementCollector(fdoc).OfClass(DB.Level).FirstElement()
        centre_lr, centre_fb = _centre_planes(fdoc)
        hw = width / 2.0
        X, Y, Z = DB.XYZ.BasisX, DB.XYZ.BasisY, DB.XYZ.BasisZ
        create = fdoc.FamilyCreate

        run_plane = create.NewReferencePlane(DB.XYZ(-run, -hw - 1, 0), DB.XYZ(-run, hw + 1, 0), Z, plan)
        run_plane.Name = u"Run"
        front_plane = create.NewReferencePlane(DB.XYZ(-run - 1, -hw, 0), DB.XYZ(1, -hw, 0), Z, plan)
        front_plane.Name = u"Front"
        back_plane = create.NewReferencePlane(DB.XYZ(-run - 1, hw, 0), DB.XYZ(1, hw, 0), Z, plan)
        back_plane.Name = u"Back"
        rise_plane = create.NewReferencePlane(DB.XYZ(-run - 1, 0, rise), DB.XYZ(1, 0, rise), Y, front)
        rise_plane.Name = u"Rise"
        fdoc.Regenerate()

        _dimension(fdoc, plan, [centre_lr.GetReference(), run_plane.GetReference()],
                   DB.XYZ(-run, -hw - 0.5, 0), DB.XYZ(0, -hw - 0.5, 0), p_run)
        _dimension(fdoc, plan, [front_plane.GetReference(), back_plane.GetReference()],
                   DB.XYZ(0.7, -hw, 0), DB.XYZ(0.7, hw, 0), p_width)
        _dimension(fdoc, plan, [front_plane.GetReference(), centre_fb.GetReference(),
                                back_plane.GetReference()],
                   DB.XYZ(0.4, -hw, 0), DB.XYZ(0.4, hw, 0), equal=True)
        _dimension(fdoc, front, [level.GetPlaneReference(), rise_plane.GetReference()],
                   DB.XYZ(0.5, 0, 0), DB.XYZ(0.5, 0, rise), p_rise)

        # the profile in the front plane, extruded across the width
        pts = [(-run, 0), (0, 0), (0, rise)]
        loop = DB.CurveArray()
        for i in range(len(pts)):
            a, b = pts[i], pts[(i + 1) % len(pts)]
            loop.Append(DB.Line.CreateBound(DB.XYZ(a[0], -hw, a[1]), DB.XYZ(b[0], -hw, b[1])))
        profile = DB.CurveArrArray()
        profile.Append(loop)
        sketch = DB.SketchPlane.Create(fdoc, DB.Plane.CreateByNormalAndOrigin(Y, DB.XYZ(0, -hw, 0)))
        solid = create.NewExtrusion(True, profile, sketch, width)
        fdoc.Regenerate()

        _align_edge(fdoc, plan, solid, run_plane, -run, 0.0)
        _align_edge(fdoc, side, solid, rise_plane, 0.0, rise)
        _align(fdoc, plan, solid, front_plane, -Y, hw)
        _align(fdoc, plan, solid, back_plane, Y, hw)

        # below: when the footing stops short of where the soffit meets the
        # base, the fill goes under the stair's overhang instead, down the
        # footing face: right angle at the origin, run along +X, rise down
        run_dn = create.NewReferencePlane(DB.XYZ(run, -hw - 1, 0), DB.XYZ(run, hw + 1, 0), Z, plan)
        run_dn.Name = u"Run Below"
        rise_dn = create.NewReferencePlane(DB.XYZ(-1, 0, -rise), DB.XYZ(run + 1, 0, -rise), Y, front)
        rise_dn.Name = u"Rise Below"
        fdoc.Regenerate()
        _dimension(fdoc, plan, [centre_lr.GetReference(), run_dn.GetReference()],
                   DB.XYZ(0, -hw - 0.8, 0), DB.XYZ(run, -hw - 0.8, 0), p_run)
        _dimension(fdoc, front, [level.GetPlaneReference(), rise_dn.GetReference()],
                   DB.XYZ(-0.5, 0, 0), DB.XYZ(-0.5, 0, -rise), p_rise)
        pts = [(0, 0), (run, 0), (0, -rise)]
        loop = DB.CurveArray()
        for i in range(len(pts)):
            a, b = pts[i], pts[(i + 1) % len(pts)]
            loop.Append(DB.Line.CreateBound(DB.XYZ(a[0], -hw, a[1]), DB.XYZ(b[0], -hw, b[1])))
        profile = DB.CurveArrArray()
        profile.Append(loop)
        below = create.NewExtrusion(True, profile, sketch, width)
        fdoc.Regenerate()
        _align_edge(fdoc, plan, below, run_dn, run, 0.0)
        _align_edge(fdoc, side, below, rise_dn, 0.0, -rise)
        _align(fdoc, plan, below, front_plane, -Y, hw)
        _align(fdoc, plan, below, back_plane, Y, hw)

        p_below = _yes_no_param(fm, u"Fill Below")
        p_above = _yes_no_param(fm, u"Fill Above")
        fm.Set(p_below, 0)
        fm.SetFormula(p_above, u"not(Fill Below)")
        for form, param in ((solid, p_above), (below, p_below)):
            fm.AssociateElementParameterToFamilyParameter(
                form.get_Parameter(DB.BuiltInParameter.IS_VISIBLE_PARAM), param)
            fm.AssociateElementParameterToFamilyParameter(
                form.get_Parameter(DB.BuiltInParameter.MATERIAL_ID_PARAM), p_mat)
        t.Commit()
    except Exception:
        if t.GetStatus() == DB.TransactionStatus.Started:
            t.RollBack()
        fdoc.Close(False)
        raise
    return fdoc


def build_footing(app, width=1.0, depth=1.0, length=3.0):
    """Make the Stair Footing family document (unsaved, left open): a
    concrete box across the stair under its first step. Type parameters
    Footing Width and Footing Depth (a type per size), instance parameter
    Footing Length (the stair's width). Origin on its top, on the face
    toward the stair, centred along its length; the width runs along -X,
    the depth down -Z. Lengths in feet."""
    template = _template(app)
    if template is None:
        raise ValueError(u"No Structural Foundation or Generic Model family template "
                         u"was found, so the stair footing could not be made.")
    fdoc = app.NewFamilyDocument(template)
    t = DB.Transaction(fdoc, u"Build")
    try:
        t.Start()
        fam_cat = DB.Category.GetCategory(fdoc, DB.BuiltInCategory.OST_StructuralFoundation)
        if fdoc.OwnerFamily.FamilyCategory.Id != fam_cat.Id:
            fdoc.OwnerFamily.FamilyCategory = fam_cat
        fm = fdoc.FamilyManager
        p_w = _length_param(fm, u"Footing Width")
        p_d = _length_param(fm, u"Footing Depth")
        p_l = _length_param(fm, u"Footing Length", instance=True)
        p_mat = _material_param(fm, u"Footing Material")
        if fm.CurrentType is None:
            fm.NewType(u"Default")
        fm.Set(p_w, width)
        fm.Set(p_d, depth)
        fm.Set(p_l, length)

        plan = [v for v in DB.FilteredElementCollector(fdoc).OfClass(DB.ViewPlan)
                if not v.IsTemplate and v.ViewType == DB.ViewType.FloorPlan][0]
        front = [v for v in DB.FilteredElementCollector(fdoc).OfClass(DB.View)
                 if not v.IsTemplate and v.ViewType == DB.ViewType.Elevation
                 and v.ViewDirection.IsAlmostEqualTo(DB.XYZ(0, -1, 0))][0]
        level = DB.FilteredElementCollector(fdoc).OfClass(DB.Level).FirstElement()
        centre_lr, centre_fb = _centre_planes(fdoc)
        hl = length / 2.0
        X, Y, Z = DB.XYZ.BasisX, DB.XYZ.BasisY, DB.XYZ.BasisZ
        create = fdoc.FamilyCreate

        w_plane = create.NewReferencePlane(DB.XYZ(-width, -hl - 1, 0), DB.XYZ(-width, hl + 1, 0), Z, plan)
        w_plane.Name = u"Width"
        front_plane = create.NewReferencePlane(DB.XYZ(-width - 1, -hl, 0), DB.XYZ(1, -hl, 0), Z, plan)
        front_plane.Name = u"Front"
        back_plane = create.NewReferencePlane(DB.XYZ(-width - 1, hl, 0), DB.XYZ(1, hl, 0), Z, plan)
        back_plane.Name = u"Back"
        d_plane = create.NewReferencePlane(DB.XYZ(-width - 1, 0, -depth), DB.XYZ(1, 0, -depth), Y, front)
        d_plane.Name = u"Depth"
        fdoc.Regenerate()

        _dimension(fdoc, plan, [centre_lr.GetReference(), w_plane.GetReference()],
                   DB.XYZ(-width, -hl - 0.5, 0), DB.XYZ(0, -hl - 0.5, 0), p_w)
        _dimension(fdoc, plan, [front_plane.GetReference(), back_plane.GetReference()],
                   DB.XYZ(0.7, -hl, 0), DB.XYZ(0.7, hl, 0), p_l)
        _dimension(fdoc, plan, [front_plane.GetReference(), centre_fb.GetReference(),
                                back_plane.GetReference()],
                   DB.XYZ(0.4, -hl, 0), DB.XYZ(0.4, hl, 0), equal=True)
        _dimension(fdoc, front, [level.GetPlaneReference(), d_plane.GetReference()],
                   DB.XYZ(0.5, 0, 0), DB.XYZ(0.5, 0, -depth), p_d)

        pts = [(-width, 0), (0, 0), (0, -depth), (-width, -depth)]
        loop = DB.CurveArray()
        for i in range(len(pts)):
            a, b = pts[i], pts[(i + 1) % len(pts)]
            loop.Append(DB.Line.CreateBound(DB.XYZ(a[0], -hl, a[1]), DB.XYZ(b[0], -hl, b[1])))
        profile = DB.CurveArrArray()
        profile.Append(loop)
        sketch = DB.SketchPlane.Create(fdoc, DB.Plane.CreateByNormalAndOrigin(Y, DB.XYZ(0, -hl, 0)))
        solid = create.NewExtrusion(True, profile, sketch, length)
        fdoc.Regenerate()

        _align(fdoc, plan, solid, w_plane, -X, width)
        _align(fdoc, plan, solid, front_plane, -Y, hl)
        _align(fdoc, plan, solid, back_plane, Y, hl)
        _align(fdoc, front, solid, d_plane, -Z, depth)
        fm.AssociateElementParameterToFamilyParameter(
            solid.get_Parameter(DB.BuiltInParameter.MATERIAL_ID_PARAM), p_mat)
        t.Commit()
    except Exception:
        if t.GetStatus() == DB.TransactionStatus.Started:
            t.RollBack()
        fdoc.Close(False)
        raise
    return fdoc


def ensure_family(doc, name=FAMILY):
    """The fill (or, with name=FOOTING, the footing) family in the project,
    built and loaded if it is not there. Must be called with no transaction
    open."""
    for f in DB.FilteredElementCollector(doc).OfClass(DB.Family):
        if f.Name == name:
            first = doc.GetElement(list(f.GetFamilySymbolIds())[0])
            if name == FOOTING or first.LookupParameter(u"Fill Below") is not None:
                return f
            # the first version of the fill had no Fill Below: replace it (its
            # fills come back when their stairs are updated)
            t = DB.Transaction(doc, u"Concrete Stair - replace old fill family")
            t.Start()
            doc.Delete(f.Id)
            t.Commit()
            break
    fdoc = (build_footing if name == FOOTING else build)(doc.Application)
    try:
        family = fdoc.LoadFamily(doc)
    finally:
        fdoc.Close(False)
    t = DB.Transaction(doc, u"Concrete Stair - {0} family".format(name))
    t.Start()
    family.Name = name
    for i in family.GetFamilySymbolIds():
        doc.GetElement(i).Name = u"Default"    # not the unsaved document's "Family6"
    t.Commit()
    return family


def soffit_foot(sizes):
    """Where the soffit meets the base, mm from the first riser."""
    foot = sizes.get("soffit_foot")
    if foot is None:
        slope = math.atan2(sizes["riser"], sizes["going"])
        foot = sizes["waist"] / math.cos(slope) * sizes["going"] / sizes["riser"]
    return foot


def fill_shape(sizes, footing_width, footing_depth):
    """(run, rise, below) in mm of the fill, or None when the footing face
    is where the soffit meets the base. Above: the footing reaches past
    that point, and the fill is the gap between the soffit, the base and the
    footing face. Below: the footing stops short of it, and the fill goes
    under the stair's overhang: the soffit carried on down to the footing
    face, at the stair's slope (no deeper than the footing)."""
    foot = soffit_foot(sizes)
    t = sizes["riser"] / float(sizes["going"])
    if footing_width - foot >= 2.0:
        run = footing_width - foot
        return run, run * t, False
    if foot - footing_width >= 2.0:
        run = foot - footing_width
        return run, min(run * t, footing_depth), True
    return None


def fill_size(sizes, footing_width):
    """(run, rise) in mm of the fill for these stair sizes and footing width,
    or None when the footing does not reach past the soffit."""
    foot = sizes.get("soffit_foot")
    if foot is None:
        # a stair read back from the model: where the soffit meets the base
        slope = math.atan2(sizes["riser"], sizes["going"])
        foot = sizes["waist"] / math.cos(slope) * sizes["going"] / sizes["riser"]
    run = footing_width - foot
    if run < 2.0:
        return None
    return run, run * sizes["riser"] / sizes["going"]


def _symbol(doc, name):
    """The fill type of that name, or None."""
    for f in DB.FilteredElementCollector(doc).OfClass(DB.Family):
        if f.Name == FAMILY:
            for i in f.GetFamilySymbolIds():
                s = doc.GetElement(i)
                if _name(s) == name:
                    return s
    return None


def _instances(doc, symbol):
    return [i for i in DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance)
            if i.Symbol.Id == symbol.Id]


def place(doc, stairs, frame, level, footing_width, material_id=None, old_id=None,
          footing_depth=None):
    """Place or update the fill for this stair (needs an open transaction).

    Its type is named by the stair's id. With `old_id` (the stair it was made
    for, before an update rebuilt the stair) that fill is kept: its type is
    renamed, resized and its instance moved to suit, so ids and tags stay.
    Returns the instance, or None when no fill is needed (an old one is then
    removed) or the family is not loaded."""
    name = unicode(_eid(stairs.Id))
    old_name = unicode(_eid(old_id)) if old_id is not None else name
    shape = fill_shape(frame.s, footing_width,
                       footing_depth if footing_depth is not None else 1e9)
    size = shape[:2] if shape else None
    if size is None:
        remove(doc, old_id if old_id is not None else stairs.Id)
        return None
    family = [f for f in DB.FilteredElementCollector(doc).OfClass(DB.Family) if f.Name == FAMILY]
    if not family:
        return None
    symbol = _symbol(doc, name) or _symbol(doc, old_name)
    if symbol is None:
        symbol = doc.GetElement(list(family[0].GetFamilySymbolIds())[0]).Duplicate(name)
    elif _name(symbol) != name:
        symbol.Name = name
    for pname, mm in ((u"Fill Run", size[0]), (u"Fill Rise", size[1]),
                      (u"Fill Width", frame.width)):
        symbol.LookupParameter(pname).Set(mm / MM)
    below = symbol.LookupParameter(u"Fill Below")
    if below is not None:
        below.Set(1 if shape[2] else 0)
    if material_id is not None and material_id != DB.ElementId.InvalidElementId:
        for p in (symbol.LookupParameter(u"Fill Material"),
                  symbol.get_Parameter(DB.BuiltInParameter.STRUCTURAL_MATERIAL_PARAM)):
            if p is not None and not p.IsReadOnly:
                p.Set(material_id)
    if not symbol.IsActive:
        symbol.Activate()
        doc.Regenerate()

    # origin on the footing face at the base, the run back toward the stair
    origin = frame.p(footing_width, 0, 0)
    angle = math.atan2(frame.d.Y, frame.d.X)
    existing = _instances(doc, symbol)
    if existing:
        inst = existing[0]
        for extra in existing[1:]:
            doc.Delete(extra.Id)
        p = inst.get_Parameter(DB.BuiltInParameter.FAMILY_LEVEL_PARAM)
        if p is not None and not p.IsReadOnly and p.AsElementId() != level.Id:
            p.Set(level.Id)
        doc.Regenerate()
        turn = angle - inst.Location.Rotation
    else:
        inst = doc.Create.NewFamilyInstance(origin, symbol, level, StructuralType.Footing)
        doc.Regenerate()
        turn = angle
    here = inst.Location.Point
    DB.ElementTransformUtils.MoveElement(doc, inst.Id, origin - here)
    if abs(math.sin(turn)) > 1e-9 or math.cos(turn) < 0:
        axis = DB.Line.CreateBound(origin, origin + DB.XYZ.BasisZ)
        DB.ElementTransformUtils.RotateElement(doc, inst.Id, axis, turn)
    doc.Regenerate()
    # foundations can snap to their level; lift to the base
    loc = inst.Location.Point
    if abs(loc.Z - origin.Z) > 1e-6:
        DB.ElementTransformUtils.MoveElement(doc, inst.Id, DB.XYZ(0, 0, origin.Z - loc.Z))
    return inst


def remove(doc, stairs_id):
    """Delete the fill type (and so its instance) named by this stair id.
    Needs an open transaction."""
    s = _symbol(doc, unicode(_eid(stairs_id)))
    if s is None:
        return
    if len(list(s.Family.GetFamilySymbolIds())) == 1:   # a family keeps one type
        for inst in _instances(doc, s):
            doc.Delete(inst.Id)
        s.Name = u"Default"
    else:
        doc.Delete(s.Id)


# --- the footing ---
def footing_name(width, depth):
    return u"{0:g}x{1:g}".format(round(width, 1), round(depth, 1))


def _footing_symbol(doc, width, depth):
    """The Stair Footing type of this size (named e.g. 250x250), made from
    the family's first type if it is new. Needs an open transaction."""
    family = [f for f in DB.FilteredElementCollector(doc).OfClass(DB.Family) if f.Name == FOOTING]
    if not family:
        raise ValueError(u"The Stair Footing family is not loaded.")
    name = footing_name(width, depth)
    symbols = [doc.GetElement(i) for i in family[0].GetFamilySymbolIds()]
    for s in symbols:
        if _name(s) == name:
            return s
    s = symbols[0].Duplicate(name)
    s.LookupParameter(u"Footing Width").Set(width / MM)
    s.LookupParameter(u"Footing Depth").Set(depth / MM)
    return s


def is_footing(element):
    """True for an instance of the Stair Footing family."""
    sym = getattr(element, "Symbol", None)
    return sym is not None and sym.Family.Name == FOOTING


def place_footing(doc, frame, level, size, material_id=None, existing=None):
    """Place the stair's footing (needs an open transaction): outer face on
    the first riser, top at the base, across the stair's full width. Given
    an `existing` Stair Footing it is moved, resized and retyped to suit
    instead, keeping its id. Returns the instance."""
    fw, fd = size
    symbol = _footing_symbol(doc, fw, fd)
    if material_id is not None and material_id != DB.ElementId.InvalidElementId:
        for p in (symbol.LookupParameter(u"Footing Material"),
                  symbol.get_Parameter(DB.BuiltInParameter.STRUCTURAL_MATERIAL_PARAM)):
            if p is not None and not p.IsReadOnly:
                p.Set(material_id)
    if not symbol.IsActive:
        symbol.Activate()
        doc.Regenerate()
    origin = frame.p(fw, 0, 0)
    angle = math.atan2(frame.d.Y, frame.d.X)
    if existing is not None and is_footing(existing):
        inst = existing
        if inst.Symbol.Id != symbol.Id:
            inst.Symbol = symbol
        p = inst.get_Parameter(DB.BuiltInParameter.FAMILY_LEVEL_PARAM)
        if p is not None and not p.IsReadOnly and p.AsElementId() != level.Id:
            p.Set(level.Id)
        doc.Regenerate()
        turn = angle - inst.Location.Rotation
    else:
        inst = doc.Create.NewFamilyInstance(origin, symbol, level, StructuralType.Footing)
        doc.Regenerate()
        turn = angle
    inst.LookupParameter(u"Footing Length").Set(frame.width / MM)
    DB.ElementTransformUtils.MoveElement(doc, inst.Id, origin - inst.Location.Point)
    if abs(math.sin(turn)) > 1e-9 or math.cos(turn) < 0:
        axis = DB.Line.CreateBound(origin, origin + DB.XYZ.BasisZ)
        DB.ElementTransformUtils.RotateElement(doc, inst.Id, axis, turn)
    doc.Regenerate()
    loc = inst.Location.Point
    if abs(loc.Z - origin.Z) > 1e-6:
        DB.ElementTransformUtils.MoveElement(doc, inst.Id, DB.XYZ(0, 0, origin.Z - loc.Z))
    return inst


def _eid(element_id):
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue
