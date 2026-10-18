# -*- coding: utf-8 -*-
"""Export the active 3D view to a 3D PDF.

Walks the view with a CustomExporter, the same pipeline Revit's own FBX
and glTF exports use, so visibility, filters, section box and links come
through as the view shows them. Each element becomes a node in the PDF
model tree under its category, coloured by its Revit material. The PDF
opens on the Revit view's camera, with Top, Front, Back, Left, Right and
Isometric as preset views.
"""
import os
import datetime

from pyrevit import revit, DB, forms

from u3d import U3DScene
from pdf3d import View, standard_views, write_pdf

doc = revit.doc

TITLE = "3D PDF"

# ── CONSTANTS ───────────────────────────────────────────────────────────────

FT_TO_MM = 304.8
DEFAULT_RGB = (0.7, 0.7, 0.7)


# ── HELPERS ─────────────────────────────────────────────────────────────────

def id_value(eid):
    """ElementId as an int, Revit 2024 (IntegerValue) and 2025+ (Value)."""
    v = getattr(eid, "Value", None)
    return v if v is not None else eid.IntegerValue


try:
    TEXT = unicode      # IronPython 2
except NameError:
    TEXT = str          # IronPython 3


def safe_text(value):
    try:
        return TEXT(value)
    except Exception:
        return u"?"


def element_name(el):
    """Element.Name, which IronPython hides behind the class property."""
    try:
        return safe_text(DB.Element.Name.__get__(el))
    except Exception:
        try:
            return safe_text(el.Name)
        except Exception:
            return u""


def element_label(el, src_doc):
    """'Family: Type [id]', falling back to the element name."""
    name = u""
    try:
        etype = src_doc.GetElement(el.GetTypeId())
    except Exception:
        etype = None
    if etype is not None:
        try:
            fam = safe_text(etype.FamilyName or u"")
        except Exception:
            fam = u""
        tname = element_name(etype)
        name = u"{}: {}".format(fam, tname) if fam and tname else (fam or tname)
    if not name:
        name = element_name(el)
    return u"{} [{}]".format(name or u"Element", id_value(el.Id))


def is_model_element(el):
    """True for model categories; levels, grids and the like are skipped."""
    try:
        cat = el.Category
        return cat is None or cat.CategoryType == DB.CategoryType.Model
    except Exception:
        return True


def category_name(el):
    try:
        if el.Category is not None:
            return safe_text(el.Category.Name)
    except Exception:
        pass
    return u"Other"


# ── EXPORT CONTEXT ──────────────────────────────────────────────────────────

class MeshContext(DB.IExportContext):
    """Collects triangles per element into a U3DScene.

    Positions are model coordinates in millimetres. The bounding box is
    tracked so the caller can centre the model.
    """

    def __init__(self, document, scene, root, progress=None, total=0):
        self.scene = scene
        self.root = root
        self.progress = progress
        self.total = max(total, 1)
        self.count = 0
        self.cancelled = False
        self._docs = [document]
        self._links = [u""]
        self._xforms = [DB.Transform.Identity]
        self._elements = []     # [element, model or None, id]
        self._groups = {}
        self._shader = scene.shader(DEFAULT_RGB)
        self.bbox = [1e30, 1e30, 1e30, -1e30, -1e30, -1e30]
        self.triangles = 0

    # -- lifecycle ----------------------------------------------------------

    def Start(self):
        return True

    def Finish(self):
        pass

    def IsCanceled(self):
        if self.progress is not None and self.progress.cancelled:
            self.cancelled = True
        return self.cancelled

    def OnViewBegin(self, node):
        return DB.RenderNodeAction.Proceed

    def OnViewEnd(self, element_id):
        pass

    # -- elements -----------------------------------------------------------

    def OnElementBegin(self, element_id):
        el = self._docs[-1].GetElement(element_id)
        if el is not None and not is_model_element(el):
            return DB.RenderNodeAction.Skip
        self._elements.append([el, None, id_value(element_id)])
        if len(self._elements) == 1:
            self.count += 1
            if self.progress is not None and self.count % 20 == 0:
                self.progress.update_progress(min(self.count, self.total),
                                              self.total)
        return DB.RenderNodeAction.Proceed

    def OnElementEnd(self, element_id):
        # Matched by id: the API does not say whether End follows a Skip.
        if self._elements and self._elements[-1][2] == id_value(element_id):
            self._elements.pop()

    def OnInstanceBegin(self, node):
        self._xforms.append(self._xforms[-1].Multiply(node.GetTransform()))
        return DB.RenderNodeAction.Proceed

    def OnInstanceEnd(self, node):
        self._xforms.pop()

    def OnLinkBegin(self, node):
        self._xforms.append(self._xforms[-1].Multiply(node.GetTransform()))
        link_doc = node.GetDocument()
        self._docs.append(link_doc)
        self._links.append(safe_text(link_doc.Title) if link_doc else u"Link")
        return DB.RenderNodeAction.Proceed

    def OnLinkEnd(self, node):
        self._xforms.pop()
        self._docs.pop()
        self._links.pop()

    def OnFaceBegin(self, node):
        return DB.RenderNodeAction.Proceed

    def OnFaceEnd(self, node):
        pass

    def OnRPC(self, node):
        pass

    def OnLight(self, node):
        pass

    # -- geometry -----------------------------------------------------------

    def OnMaterial(self, node):
        rgb = DEFAULT_RGB
        try:
            c = node.Color
            if c is not None and c.IsValid:
                rgb = (c.Red / 255.0, c.Green / 255.0, c.Blue / 255.0)
        except Exception:
            pass
        opacity = 1.0
        try:
            opacity = 1.0 - float(node.Transparency)
        except Exception:
            pass
        self._shader = self.scene.shader(rgb, max(0.05, min(1.0, opacity)))

    def _group_for(self, el):
        cat = category_name(el)
        link = self._links[-1]
        key = (link, cat)
        name = self._groups.get(key)
        if name is None:
            parent = self.root
            if link:
                lkey = (link, None)
                parent = self._groups.get(lkey)
                if parent is None:
                    parent = self.scene.add_group(u"Link: {}".format(link),
                                                  self.root)
                    self._groups[lkey] = parent
            name = self.scene.add_group(cat, parent)
            self._groups[key] = name
        return name

    def _model(self):
        """Model node for the innermost element, made on first geometry."""
        if not self._elements:
            return None
        entry = self._elements[-1]
        if entry[1] is None:
            el = entry[0]
            if el is None:
                return None
            entry[1] = self.scene.add_model(
                element_label(el, self._docs[-1]), self._group_for(el))
        return entry[1]

    def OnPolymesh(self, node):
        model = self._model()
        if model is None:
            return
        xf = self._xforms[-1]
        identity = xf.IsIdentity
        bb = self.bbox
        pos = []
        for p in node.GetPoints():
            if not identity:
                p = xf.OfPoint(p)
            x = p.X * FT_TO_MM
            y = p.Y * FT_TO_MM
            z = p.Z * FT_TO_MM
            pos.append(x)
            pos.append(y)
            pos.append(z)
            if x < bb[0]: bb[0] = x
            if y < bb[1]: bb[1] = y
            if z < bb[2]: bb[2] = z
            if x > bb[3]: bb[3] = x
            if y > bb[4]: bb[4] = y
            if z > bb[5]: bb[5] = z
        tri = []
        for f in node.GetFacets():
            tri.append(f.V1)
            tri.append(f.V2)
            tri.append(f.V3)
        self.triangles += len(tri) // 3
        U3DScene.add_triangles(model, self._shader, pos, tri)


# ── MAIN ────────────────────────────────────────────────────────────────────

def pick_output(view):
    folder = os.path.dirname(doc.PathName) if doc.PathName else ""
    name = u"{} {}".format(safe_text(doc.Title), safe_text(view.Name))
    for ch in u'<>:"/\\|?*{}':
        name = name.replace(ch, u"_")
    return forms.save_file(file_ext="pdf", default_name=name.strip(),
                           init_dir=folder, title=TITLE)


def main():
    view = doc.ActiveView
    if not isinstance(view, DB.View3D) or view.IsTemplate:
        forms.alert("Open a 3D view first. The PDF matches what the active "
                    "3D view shows: its section box, filters and hidden "
                    "elements all apply.", title=TITLE)
        return

    path = pick_output(view)
    if not path:
        return

    total = DB.FilteredElementCollector(doc, view.Id) \
        .WhereElementIsNotElementType().GetElementCount()

    scene = U3DScene(units_m=0.001)
    root = scene.add_group(safe_text(doc.Title) or u"Model")

    with forms.ProgressBar(title="Reading model... ({value} of {max_value})",
                           cancellable=True) as pb:
        ctx = MeshContext(doc, scene, root, pb, total)
        exporter = DB.CustomExporter(doc, ctx)
        exporter.IncludeGeometricObjects = False
        exporter.ShouldStopOnError = False
        exporter.Export(view)

    if ctx.cancelled:
        return
    if ctx.triangles == 0:
        forms.alert("Nothing to export: the view has no visible model "
                    "geometry.", title=TITLE)
        return

    bb = ctx.bbox
    center = ((bb[0] + bb[3]) / 2.0, (bb[1] + bb[4]) / 2.0,
              (bb[2] + bb[5]) / 2.0)
    radius = 0.5 * ((bb[3] - bb[0]) ** 2 + (bb[4] - bb[1]) ** 2 +
                    (bb[5] - bb[2]) ** 2) ** 0.5
    # Centre on the origin: positions are 32-bit floats in the file, and a
    # model far from the internal origin would lose precision.
    scene.translate(-center[0], -center[1], -center[2])

    with forms.ProgressBar(title="Writing 3D PDF... ({value} of {max_value})") as pb:
        u3d_bytes = scene.to_bytes(
            progress=lambda i, n: pb.update_progress(i, n) if i % 25 == 0 else None)

    orient = view.GetOrientation()
    f = orient.ForwardDirection
    u = orient.UpDirection
    views = [View(u"Revit view: {}".format(safe_text(view.Name)),
                  (f.X, f.Y, f.Z), (u.X, u.Y, u.Z))]
    views.extend(standard_views())

    subtitle = u"{} | Exported {}".format(
        safe_text(view.Name), datetime.date.today().strftime("%d %b %Y"))
    write_pdf(path, u3d_bytes, views, (0.0, 0.0, 0.0), max(radius, 1.0),
              title=safe_text(doc.Title), subtitle=subtitle)


if __name__ == "__main__":
    main()
