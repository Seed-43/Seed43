# -*- coding: utf-8 -*-
# "Trim Column"
# "Seed43"
# """
# Trim or extend a structural column to a face, until Esc. Single mode picks
# a face then one column; Multiple picks a face once then any number of
# columns, like Revit's Trim/Extend Multiple. Revit's own Trim/Extend Single Element refuses vertical columns, so
# each column is switched to End Point Driven, trimmed, then set back to its
# original Column Style, all inside one transaction (one undo per column).
# """
import clr

from pyrevit import revit, DB, forms
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

doc = revit.doc
uidoc = revit.uidoc

# ── CONSTANTS ───────────────────────────────────────────────────────────────

# SLANTED_COLUMN_TYPE_PARAM values, as shown in the Column Style dropdown.
STYLE_VERTICAL = 0
STYLE_END_POINT = 2

MIN_LENGTH = 10.0 / 304.8      # 10 mm, anything shorter is a mis-pick
REACH = 1000.0                 # ft either side of the column for curved faces
ON_FACE_TOL = 1e-4             # ft, picked point vs face in world space
# In a 2D view a face counts as edge-on (pickable) when its normal is within
# about 1 degree of the view plane: |normal . view direction| < sin(1 deg).
EDGE_ON_TOL = 0.0175

COLUMN_CAT = DB.ElementId(DB.BuiltInCategory.OST_StructuralColumns)
_CANNOT_KEEP_JOINED = DB.BuiltInFailures.JoinElementsFailures.CannotKeepJoined

BAR_TITLE = u"Trim Column"
STEP_FACE = u"Step 1 of 2: pick the FACE to trim or extend to"
STEP_COLUMN = u"Step 2 of 2: click the COLUMN on the part to keep"
STEP_FACE_MULTI = u"Multiple: pick the FACE every column trims or extends to"
STEP_COLUMN_MULTI = (u"Multiple: click each COLUMN on the part to keep "
                     u"(ESC for a new face)")

MODE_SINGLE = u"Single"
MODE_MULTIPLE = u"Multiple"


class _Skip(Exception):
    """This column cannot be done; say why and carry on with the next."""


class _Failed(Exception):
    """Revit threw part way through; the message names the step it was on.

    Native faults (SEHException) say nothing useful about where they came
    from, so trim_column tracks its step and reports that instead.
    """


def _reason(error):
    """One readable line from an exception, never the whole .NET stack.

    str() of a .NET exception in IronPython includes the full stack trace,
    which fills the dialog without saying anything.
    """
    net = getattr(error, "clsException", None)
    text = net.Message if net is not None else unicode(error)
    text = text.strip()
    return text.splitlines()[0] if text else type(error).__name__


def _eid(element):
    """Element id as a number, across the 2024 ElementId 64-bit change."""
    try:
        return element.Id.Value
    except AttributeError:
        return element.Id.IntegerValue


# ── SELECTION ───────────────────────────────────────────────────────────────

class ColumnFilter(ISelectionFilter):
    """Structural columns that carry a Column Style parameter."""

    def AllowElement(self, element):
        cat = element.Category
        if cat is None or not cat.Id.Equals(COLUMN_CAT):
            return False
        return element.get_Parameter(
            DB.BuiltInParameter.SLANTED_COLUMN_TYPE_PARAM) is not None

    def AllowReference(self, reference, point):
        return False


class FaceFilter(ISelectionFilter):
    """Any face in a 3D view; only edge-on faces in a 2D view.

    In a plan, section or elevation the faces that make sense to trim to are
    the ones drawn as lines: the top, bottom, left and right as seen. A face
    turned toward the viewer is a surface you look at, not a line, so it is
    not offered. The normal is taken at the hovered point, so a curved face
    is judged by the part under the cursor.
    """

    def __init__(self, view):
        self.view_dir = None
        if not isinstance(view, DB.View3D):
            try:
                self.view_dir = view.ViewDirection
            except Exception:
                pass  # no direction to judge by, allow everything

    def AllowElement(self, element):
        return True

    def AllowReference(self, reference, point):
        if self.view_dir is None:
            return True
        try:
            face, xf = face_in_world(reference, point)
            if isinstance(face, DB.PlanarFace):
                local_normal = face.FaceNormal
            else:
                hit = face.Project(xf.Inverse.OfPoint(point))
                if hit is None:
                    return False
                local_normal = face.ComputeNormal(hit.UVPoint)
            normal = xf.OfVector(local_normal).Normalize()
            return abs(normal.DotProduct(self.view_dir)) < EDGE_ON_TOL
        except Exception:
            return False


def pick_face():
    """The face to trim or extend to, or None on Esc."""
    try:
        return uidoc.Selection.PickObject(
            ObjectType.Face, FaceFilter(doc.ActiveView),
            "Trim Column: pick the face to trim/extend to (Esc to finish)")
    except OperationCanceledException:
        return None


def pick_column():
    """The column to trim, clicked on the part to keep, or None on Esc."""
    try:
        return uidoc.Selection.PickObject(
            ObjectType.Element, ColumnFilter(),
            "Trim Column: pick the column on the part to keep (Esc to finish)")
    except OperationCanceledException:
        return None


# ── GEOMETRY ────────────────────────────────────────────────────────────────

def face_in_world(face_ref, point=None):
    """The picked face plus the transform that takes it into model space.

    GetGeometryObjectFromReference hands back family instance faces in the
    family's own coordinates when the instance shares its symbol geometry,
    so the picked point is tested against the face to find out which.
    Pass point when the reference carries no GlobalPoint of its own, as in
    a selection filter's AllowReference.
    """
    element = doc.GetElement(face_ref)
    face = element.GetGeometryObjectFromReference(face_ref)
    if not isinstance(face, DB.Face):
        raise _Skip("That pick was not a face.")

    xf = DB.Transform.Identity
    if isinstance(element, DB.FamilyInstance):
        gp = point if point is not None else face_ref.GlobalPoint
        hit = face.Project(gp)
        if hit is None or hit.XYZPoint.DistanceTo(gp) > ON_FACE_TOL:
            xf = element.GetTotalTransform()
    return face, xf


class Target(object):
    """What the columns trim to, captured the moment the face is picked.

    A flat face is kept as its plane: a normal and the picked point, plain
    XYZ values that no regenerate can invalidate. Found 2026-10-01 with
    column 3854366 to beam 3859361: once the first column (or the style
    switch) changes the joins, the picked Reference can stop resolving to a
    face at all, and a Face object held across a regenerate can point at
    freed geometry and crash Revit (SEHException). So the reference is read
    once, here, and a flat face never touches it again.

    A curved face has no plane to keep, so it holds the reference and is
    read afresh after each regenerate instead (see trim_column).
    """

    def __init__(self, face_ref):
        self.ref = face_ref
        self.element_id = face_ref.ElementId
        self.point = face_ref.GlobalPoint
        face, xf = face_in_world(face_ref)
        self.normal = None
        if isinstance(face, DB.PlanarFace):
            self.normal = xf.OfVector(face.FaceNormal).Normalize()


def axis_hits(target, face, xf, p0, p1):
    """Points where the column's axis, run out both ways, meets the target.

    A flat face is treated as its whole plane, the same as Revit's trim, so
    a beam soffit that stops short of the column still works; face and xf
    are unused then. A curved face is intersected as drawn.
    """
    axis = (p1 - p0).Normalize()

    if target.normal is not None:
        denom = target.normal.DotProduct(axis)
        if abs(denom) < 1e-9:
            return []
        t = target.normal.DotProduct(target.point - p0) / denom
        return [p0 + axis * t]

    inv = xf.Inverse
    line = DB.Line.CreateBound(inv.OfPoint(p0 - axis * REACH),
                               inv.OfPoint(p1 + axis * REACH))
    results = clr.Reference[DB.IntersectionResultArray]()
    if face.Intersect(line, results) != DB.SetComparisonResult.Overlap:
        return []
    arr = results.Value
    return [xf.OfPoint(arr.get_Item(i).XYZPoint) for i in range(arr.Size)]


def new_ends(p0, p1, hit, keep_point):
    """Base and top after trimming to hit, keeping the side that was clicked.

    Returns (base, top, moved) where moved is "base" or "top".
    """
    axis = (p1 - p0).Normalize()
    length = p0.DistanceTo(p1)
    t_hit = (hit - p0).DotProduct(axis)

    if t_hit >= length:
        return p0, hit, "top"          # extend up
    if t_hit <= 0.0:
        return hit, p1, "base"         # extend down
    # the face cuts the column: keep whichever side was clicked
    t_keep = (keep_point - p0).DotProduct(axis) if keep_point else length
    if t_keep < t_hit:
        return p0, hit, "top"
    return hit, p1, "base"


# ── CORE LOGIC ──────────────────────────────────────────────────────────────

def _is_attached(column, end):
    bip = (DB.BuiltInParameter.COLUMN_TOP_ATTACHED_PARAM if end == "top"
           else DB.BuiltInParameter.COLUMN_BASE_ATTACHED_PARAM)
    param = column.get_Parameter(bip)
    return param is not None and param.AsInteger() == 1


class _UnjoinWhereNeeded(DB.IFailuresPreprocessor):
    """Answer Revit's "Can't keep elements joined" with Unjoin Elements.

    Moving a column end can break a join Revit cannot keep; by hand the
    fix is the Unjoin Elements button on that error, which is the
    DetachElements resolution. Only that one failure is touched; anything
    else goes to Revit's normal dialogs. unjoined counts the fixes made.
    """

    def __init__(self):
        self.unjoined = 0

    def PreprocessFailures(self, accessor):
        resolved = False
        for msg in accessor.GetFailureMessages():
            if msg.GetFailureDefinitionId().Guid != _CANNOT_KEEP_JOINED.Guid:
                continue
            if msg.HasResolutionOfType(DB.FailureResolutionType.DetachElements):
                msg.SetCurrentResolutionType(
                    DB.FailureResolutionType.DetachElements)
            elif not msg.HasResolutions():
                continue
            accessor.ResolveFailure(msg)
            self.unjoined += 1
            resolved = True
        if resolved:
            return DB.FailureProcessingResult.ProceedWithCommit
        return DB.FailureProcessingResult.Continue


def trim_column(column, col_ref, target):
    """Switch to End Point Driven, move one end to the target, switch back.

    Returns how many joins Revit had to undo to keep the move.
    """
    style = column.get_Parameter(DB.BuiltInParameter.SLANTED_COLUMN_TYPE_PARAM)
    original = style.AsInteger()

    t = DB.Transaction(doc, "Trim/Extend Column")
    t.Start()
    unjoin = _UnjoinWhereNeeded()
    options = t.GetFailureHandlingOptions()
    options.SetFailuresPreprocessor(unjoin)
    t.SetFailureHandlingOptions(options)
    step = "starting"
    face = xf = None
    try:
        if original != STYLE_END_POINT:
            step = "switching the column to End Point Driven"
            style.Set(STYLE_END_POINT)
            step = "regenerating after the style switch"
            doc.Regenerate()

        # A curved face has to be read from the reference, and only after
        # the regenerate: one fetched before it can point at freed native
        # geometry and throw SEHException (column 4149497 to joist 4140675).
        # A flat face was captured as a plane at pick time, see Target.
        if target.normal is None:
            step = "re-reading the picked face after the regenerate"
            face, xf = face_in_world(target.ref, target.point)

        step = "reading the column axis"
        loc = column.Location
        if not isinstance(loc, DB.LocationCurve):
            raise _Skip("This column has no axis Revit will let us move.")
        p0 = loc.Curve.GetEndPoint(0)
        p1 = loc.Curve.GetEndPoint(1)

        step = "intersecting the axis with the face"
        hits = axis_hits(target, face, xf, p0, p1)
        if not hits:
            raise _Skip("The column's axis never meets that face.")
        keep = col_ref.GlobalPoint
        target = keep if keep else p1
        hit = min(hits, key=lambda h: h.DistanceTo(target))

        base, top, moved = new_ends(p0, p1, hit, keep)
        if base.DistanceTo(top) < MIN_LENGTH:
            raise _Skip("That would leave the column with no length.")
        if _is_attached(column, moved):
            raise _Skip("The {} of this column is attached. Detach it first, "
                        "or Revit will pull it straight back.".format(moved))

        step = "moving the {} of the column".format(moved)
        loc.Curve = DB.Line.CreateBound(base, top)
        step = "regenerating after the move"
        doc.Regenerate()

        if original != STYLE_END_POINT:
            step = "setting the Column Style back"
            style.Set(original)
        step = "committing (Revit regenerates and checks joins here)"
        status = t.Commit()
        if status != DB.TransactionStatus.Committed:
            raise _Skip("Revit rolled the trim back when committing it. "
                        "Nothing was changed.")
        return unjoin.unjoined
    except Exception as e:
        # A native fault can leave the transaction unable to roll back
        # cleanly; never let that second error hide the first.
        try:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
        except Exception:
            pass
        if isinstance(e, _Skip):
            raise
        raise _Failed(u"Failed while {}.\n\n{}".format(step, _reason(e)))


# ── UI / ENTRY POINT ────────────────────────────────────────────────────────

def _bar_text(step, done, unjoined=0):
    text = u"{}: {}. ESC to finish.".format(BAR_TITLE, step)
    if done:
        text += u"   ({} done".format(done)
        if unjoined:
            text += u", {} join{} undone".format(
                unjoined, u"" if unjoined == 1 else u"s")
        text += u")"
    return text


def choose_mode():
    """Single or Multiple, or None if the switch window is closed.

    CommandSwitchWindow filters as you type, so S or M then Enter picks a
    mode without the mouse. A key cannot be read during PickObject itself,
    which is why the choice is made up front.
    """
    return forms.CommandSwitchWindow.show(
        [MODE_SINGLE, MODE_MULTIPLE],
        message=u"Trim/Extend Column: type S or M, then Enter")


def capture_target(face_ref):
    """The Target for a fresh face pick, or None after saying why not."""
    try:
        return Target(face_ref)
    except _Skip as e:
        forms.alert(str(e), title="Trim Column")
    except Exception as e:
        forms.alert(u"Could not read that face.\n\n{}".format(_reason(e)),
                    title="Trim Column")
    return None


def run_one(column, col_ref, target):
    """Trim one column, reporting any problem.

    Returns the number of joins undone, or None if it was not trimmed.
    """
    if column.Id.Equals(target.element_id):
        forms.alert("Pick a face on a different element to the column.",
                    title="Trim Column")
        return None
    try:
        return trim_column(column, col_ref, target)
    except _Skip as e:
        forms.alert(str(e), title="Trim Column")
    except Exception as e:
        detail = unicode(e) if isinstance(e, _Failed) else _reason(e)
        forms.alert(u"Revit would not trim column {} ({}).\n\n{}"
                    .format(_eid(column), column.Name, detail),
                    title="Trim Column")
    return None


def main():
    mode = choose_mode()
    if not mode:
        return
    multiple = mode == MODE_MULTIPLE
    step_face = STEP_FACE_MULTI if multiple else STEP_FACE
    step_column = STEP_COLUMN_MULTI if multiple else STEP_COLUMN

    done = 0
    unjoined = 0
    with forms.WarningBar(title=_bar_text(step_face, done)) as bar:
        # Esc at a column pick goes back to picking a face; Esc at the face
        # pick ends the tool. Same in both modes, as Revit's own Trim/Extend.
        while True:
            # the bar is a live WPF window, its text can change between picks
            bar.message_tb.Text = _bar_text(step_face, done, unjoined)
            face_ref = pick_face()
            if face_ref is None:
                break
            # Read the face now, while the pick is fresh. A flat face is kept
            # as a plane from here on, so it stays good across every column
            # in Multiple mode however much each trim changes the model.
            target = capture_target(face_ref)
            if target is None:
                continue
            while True:
                bar.message_tb.Text = _bar_text(step_column, done, unjoined)
                col_ref = pick_column()
                if col_ref is None:
                    break
                result = run_one(doc.GetElement(col_ref), col_ref, target)
                if result is not None:
                    done += 1
                    unjoined += result
                if not multiple:
                    break


main()
