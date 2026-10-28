# -*- coding: utf-8 -*-
"""Export the active 3D view to a Seed43 Studio .seed file.

Walks the view with a CustomExporter, as 3D PDF does, so visibility,
filters, section box and links come through as the view shows them. The
file is the one Navisworks' Export to Studio writes (seed_writer.py), with
the elements laid out the way an IFC export makes them: a Project > Site >
Building > Storey > element tree, each element named Family:Type:Id with
the IFC class and GlobalId Revit's IFC exporter would give it, and Mark,
Section, Grade, Assembly and Length in a StruCad property set (where
Studio reads its fixed fields on an IFC). Revit material colours, the
other element properties and the view's camera come through as before.
Coordinates are shared coordinates, the Navisworks NWC default, so a Revit
export and a Navisworks export of the same model line up in Studio.
"""
import os

from pyrevit import revit, DB, forms

from seed_writer import SeedScene, SeedItem, SeedPart, write_source_seed
from ifc_style import ifc_class, ifc_guid, kind_of, storey_of, STRUCAD_PSET

doc = revit.doc

TITLE = "3D Seed"

# ── CONSTANTS ───────────────────────────────────────────────────────────────

FT_TO_MM = 304.8
DEFAULT_RGB = (0.7, 0.7, 0.7)
NO_LEVEL = u"(no level)"

# IFC classes whose type name is a steel section.
SECTION_CLASSES = ("IfcBeam", "IfcColumn", "IfcMember")

SMART = "Smart model: an IFC with every Revit parameter"
WITH_PROPS = "Mesh with element properties"
NO_PROPS = "Mesh: geometry, names, materials and IFC fields only"

# Revit's IFC exporter, set up so nothing of the model is lost: every
# instance and type parameter (built-in, project and shared) as Revit
# property sets, the IFC common sets and quantities, steel elements, and
# only what the view shows, placed in shared coordinates.
IFC_OPTIONS = (
    ("ExportInternalRevitPropertySets", "true"),
    ("ExportIFCCommonPropertySets", "true"),
    ("ExportMaterialPsets", "true"),
    ("ExportSchedulesAsPsets", "false"),
    ("ExportUserDefinedPsets", "false"),
    ("VisibleElementsOfCurrentView", "true"),
    ("UseActiveViewGeometry", "true"),
    ("IncludeSteelElements", "true"),
    ("ExportPartsAsBuildingElements", "false"),
    ("ExportAnnotations", "false"),
    ("ExportLinkedFiles", "false"),
    ("SitePlacement", "0"),             # shared coordinates
)


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


def family_type(el, src_doc):
    """(family, type) names of the element's type, either may be empty."""
    try:
        etype = src_doc.GetElement(el.GetTypeId())
    except Exception:
        etype = None
    if etype is None:
        return u"", u""
    try:
        fam = safe_text(etype.FamilyName or u"")
    except Exception:
        fam = u""
    return fam, element_name(etype)


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


def param_text(p):
    try:
        if not p.HasValue:
            return u""
        s = p.AsValueString()
        if not s and p.StorageType == DB.StorageType.String:
            s = p.AsString()
        return safe_text(s) if s else u""
    except Exception:
        return u""


def group_label(p):
    """The parameter's group as Revit's Properties palette names it."""
    try:
        return safe_text(DB.LabelUtils.GetLabelForGroup(
            p.Definition.GetGroupTypeId())) or u"Other"
    except Exception:
        return u"Other"


def structural_material(el, src_doc):
    try:
        p = el.get_Parameter(DB.BuiltInParameter.STRUCTURAL_MATERIAL_PARAM)
        if p is not None and p.HasValue:
            mat = src_doc.GetElement(p.AsElementId())
            if mat is not None:
                return element_name(mat)
    except Exception:
        pass
    return u""


def element_props(el, src_doc, fam, tname):
    props = {u"Element": {u"Id": safe_text(id_value(el.Id)),
                          u"Category": category_name(el)}}
    try:
        props[u"Element"][u"UniqueId"] = safe_text(el.UniqueId)
    except Exception:
        pass
    if fam:
        props[u"Element"][u"Family"] = fam
    if tname:
        props[u"Element"][u"Type"] = tname
    for p in el.Parameters:
        value = param_text(p)
        if value:
            props.setdefault(group_label(p), {})[
                safe_text(p.Definition.Name)] = value
    try:
        etype = src_doc.GetElement(el.GetTypeId())
    except Exception:
        etype = None
    if etype is not None:
        tp = {}
        for p in etype.Parameters:
            value = param_text(p)
            if value:
                tp[safe_text(p.Definition.Name)] = value
        if tp:
            props[u"Type"] = tp
    return props


# ── IFC-STYLE FIELDS ────────────────────────────────────────────────────────

def bic_name(el):
    """'OST_StructuralFraming' for an element's category, or ''."""
    try:
        cat = el.Category
        if cat is None:
            return u""
        b = getattr(cat, "BuiltInCategory", None)       # Revit 2023+
        if b is not None:
            return safe_text(b)
        import System
        return safe_text(System.Enum.ToObject(DB.BuiltInCategory,
                                              int(id_value(cat.Id))))
    except Exception:
        return u""


def param_string(el, bip=None, name=None):
    """A parameter's text by built-in id (skipped when the Revit version
    has no such id) or by name, '' when absent."""
    if el is None:
        return u""
    try:
        p = None
        if bip is not None:
            b = getattr(DB.BuiltInParameter, bip, None)
            if b is not None:
                p = el.get_Parameter(b)
        elif name:
            p = el.LookupParameter(name)
        return param_text(p) if p is not None else u""
    except Exception:
        return u""


def export_as(el, etype):
    """The element's "Export to IFC As", else its type's."""
    for src, bips in ((el, ("IFC_EXPORT_ELEMENT_AS",)),
                      (etype, ("IFC_EXPORT_ELEMENT_TYPE_AS",
                               "IFC_EXPORT_ELEMENT_AS"))):
        for bip in bips:
            v = param_string(src, bip)
            if v:
                return v
        v = param_string(src, name="IfcExportAs")
        if v:
            return v
    return u""


def structural_type(el):
    try:
        return safe_text(el.StructuralType)
    except Exception:
        return u""


def global_id(el, src_doc):
    """The IFC GlobalId: the one a past IFC export stored, else the id
    Revit's exporter derives from the element."""
    v = param_string(el, "IFC_GUID")
    if len(v) == 22:
        return v
    try:
        return ifc_guid(safe_text(DB.ExportUtils.GetExportId(src_doc, el.Id)))
    except Exception:
        return u""


class Levels(object):
    """Each document's levels, to put elements in storeys."""

    def __init__(self):
        self._by_doc = {}

    def of(self, src_doc):
        key = id(src_doc)
        hit = self._by_doc.get(key)
        if hit is None:
            hit = {}
            for lv in DB.FilteredElementCollector(src_doc).OfClass(DB.Level):
                hit[id_value(lv.Id)] = (element_name(lv), lv.ProjectElevation)
            self._by_doc[key] = hit
        return hit

    def storey(self, el, src_doc):
        """(name, elevation) of the element's level, or None."""
        levels = self.of(src_doc)
        ids = []
        try:
            ids.append(el.LevelId)
        except Exception:
            pass
        for bip in ("INSTANCE_REFERENCE_LEVEL_PARAM",
                    "FAMILY_BASE_LEVEL_PARAM", "SCHEDULE_LEVEL_PARAM"):
            try:
                p = el.get_Parameter(getattr(DB.BuiltInParameter, bip))
                if p is not None and p.HasValue:
                    ids.append(p.AsElementId())
            except Exception:
                pass
        for lid in ids:
            try:
                hit = levels.get(id_value(lid))
            except Exception:
                hit = None
            if hit is not None:
                return hit
        try:
            bb = el.get_BoundingBox(None)
            if bb is not None:
                return storey_of(list(levels.values()), bb.Min.Z)
        except Exception:
            pass
        return None


def length_mm(el):
    for bip in ("STRUCTURAL_FRAME_CUT_LENGTH", "INSTANCE_LENGTH_PARAM",
                "CURVE_ELEM_LENGTH"):
        try:
            p = el.get_Parameter(getattr(DB.BuiltInParameter, bip))
            if p is not None and p.HasValue and \
                    p.StorageType == DB.StorageType.Double and p.AsDouble() > 0:
                return u"{}".format(int(round(p.AsDouble() * FT_TO_MM)))
        except Exception:
            pass
    return u""


def assembly_name(el, src_doc):
    try:
        aid = el.AssemblyInstanceId
        if aid is not None and aid != DB.ElementId.InvalidElementId:
            a = src_doc.GetElement(aid)
            if a is not None:
                return element_name(a)
    except Exception:
        pass
    return u""


def ifc_fields(c, cls):
    """The StruCad set: Studio's fixed card fields, only those with values."""
    el, src_doc = c.el, c.doc
    try:
        etype = src_doc.GetElement(el.GetTypeId())
    except Exception:
        etype = None
    f = [(u"Kind", kind_of(cls)),
         (u"Mark", param_string(el, "ALL_MODEL_MARK")),
         (u"Section", c.tname if cls in SECTION_CLASSES else u""),
         (u"Grade", structural_material(el, src_doc)),
         (u"Assembly", assembly_name(el, src_doc)),
         (u"Length", length_mm(el)),
         (u"Desc", param_string(etype, "ALL_MODEL_DESCRIPTION")),
         (u"Phase", param_string(el, "PHASE_CREATED"))]
    return dict((k, v) for k, v in f if v)


# ── EXPORT CONTEXT ──────────────────────────────────────────────────────────

class Collected(object):
    """One element as read: its parts by material, before the tree is built."""

    def __init__(self, el, src_doc, model):
        self.el = el
        self.doc = src_doc
        self.model = model
        self.parts = []
        self.by_key = {}


class SeedContext(DB.IExportContext):
    """Collects triangles per element and material, in shared coordinates (mm)."""

    def __init__(self, document, to_shared, progress=None, total=0):
        self.progress = progress
        self.total = max(total, 1)
        self.count = 0
        self.cancelled = False
        self._docs = [document]
        self._models = [0]
        self._xforms = [to_shared]
        self._elements = []     # [element, Collected or None, id]
        self._material = (DEFAULT_RGB, 1.0, u"")
        self._mat_names = {}
        self.models = [safe_text(document.Title) or u"Model"]
        self.collected = {}     # (model, element id) -> Collected
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
        # Every link instance is its own model, so two copies of one link
        # never merge their elements.
        name = safe_text(link_doc.Title) if link_doc else u"Link"
        taken = set(self.models)
        label, n = name, 2
        while label in taken:
            label = u"{} ({})".format(name, n)
            n += 1
        self.models.append(label)
        self._models.append(len(self.models) - 1)
        return DB.RenderNodeAction.Proceed

    def OnLinkEnd(self, node):
        self._xforms.pop()
        self._docs.pop()
        self._models.pop()

    def OnFaceBegin(self, node):
        return DB.RenderNodeAction.Proceed

    def OnFaceEnd(self, node):
        pass

    def OnRPC(self, node):
        pass

    def OnLight(self, node):
        pass

    # -- geometry -----------------------------------------------------------

    def _material_name(self, mat_id):
        key = (self._models[-1], id_value(mat_id))
        name = self._mat_names.get(key)
        if name is None:
            name = u""
            try:
                mat = self._docs[-1].GetElement(mat_id)
                if mat is not None:
                    name = element_name(mat)
            except Exception:
                pass
            self._mat_names[key] = name
        return name

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
        name = u""
        try:
            if node.MaterialId != DB.ElementId.InvalidElementId:
                name = self._material_name(node.MaterialId)
        except Exception:
            pass
        self._material = (rgb, max(0.05, min(1.0, opacity)), name)

    def _current(self):
        """Collected entry for the innermost element, made on first geometry."""
        if not self._elements:
            return None
        entry = self._elements[-1]
        if entry[1] is None:
            el = entry[0]
            if el is None:
                return None
            key = (self._models[-1], entry[2])
            item = self.collected.get(key)
            if item is None:
                item = Collected(el, self._docs[-1], self._models[-1])
                self.collected[key] = item
            entry[1] = item
        return entry[1]

    def OnPolymesh(self, node):
        item = self._current()
        if item is None:
            return
        part = item.by_key.get(self._material)
        if part is None:
            rgb, opacity, name = self._material
            part = SeedPart(rgb, opacity, name)
            item.by_key[self._material] = part
            item.parts.append(part)
        xf = self._xforms[-1]
        bb = self.bbox
        pos = []
        for p in node.GetPoints():
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
        part.add(pos, tri)


# ── BUILD THE TREE ──────────────────────────────────────────────────────────

def site_building(src_doc):
    """(site name, building name) as Revit's IFC export names them."""
    site = building = u""
    try:
        site = element_name(src_doc.ActiveProjectLocation)
    except Exception:
        pass
    try:
        building = safe_text(src_doc.ProjectInformation.BuildingName or u"")
    except Exception:
        pass
    return site or u"Default", building or u"Default"


def build_scene(ctx, with_props, progress=None):
    """IfcProject (one per model) > IfcSite > IfcBuilding > IfcBuildingStorey
    (by elevation) > element, numbered depth-first so Studio's tree-order
    rule holds (each branch and its items one run)."""
    app = doc.Application
    scene = SeedScene(document=safe_text(doc.Title) or u"Model",
                      product=u"Autodesk Revit {}".format(app.VersionNumber))
    levels = Levels()

    # model -> (storey elevation, storey name) -> [(sort key, Collected)]
    tree = {}
    model_docs = {}
    for c in ctx.collected.values():
        if not c.parts:
            continue
        fam, tname = family_type(c.el, c.doc)
        c.fam, c.tname = fam, tname
        c.ftype = u"{}: {}".format(fam, tname) if fam and tname \
            else (fam or tname)
        etype = None
        try:
            etype = c.doc.GetElement(c.el.GetTypeId())
        except Exception:
            pass
        c.cls = ifc_class(bic_name(c.el), structural_type(c.el),
                          export_as(c.el, etype))
        eid = id_value(c.el.Id)
        # Revit's IFC export names an element Family:Type:Id
        base = u"{}:{}".format(fam, tname) if fam and tname \
            else (fam or tname or element_name(c.el) or u"Element")
        c.label = u"{}:{}".format(base, eid)
        st = levels.storey(c.el, c.doc)
        key = (1, 0.0, NO_LEVEL) if st is None else (0, st[1], st[0])
        tree.setdefault(c.model, {}).setdefault(key, []).append(
            ((c.cls, c.label, eid), c))
        model_docs.setdefault(c.model, c.doc)

    done, total = 0, len(ctx.collected)
    for model in sorted(tree):
        site, building = site_building(model_docs[model])
        project = scene.add_node(ctx.models[model], u"IfcProject")
        site_node = scene.add_node(site, u"IfcSite", project)
        bldg = scene.add_node(building, u"IfcBuilding", site_node)
        storeys = tree[model]
        for key in sorted(storeys):
            storey_name = key[2]
            storey = scene.add_node(storey_name, u"IfcBuildingStorey", bldg)
            for _, c in sorted(storeys[key], key=lambda e: e[0]):
                item = SeedItem(scene.add_node(c.label, c.cls, storey))
                item.parts = c.parts
                item.cat = category_name(c.el)
                item.type = c.ftype
                item.mat = structural_material(c.el, c.doc)
                if not item.mat:
                    # the material with the most triangles stands in
                    best = max(c.parts, key=lambda p: len(p.triangles))
                    item.mat = best.mat
                props = {}
                ifc = {u"GlobalId": global_id(c.el, c.doc),
                       u"Class": c.cls, u"Storey": storey_name}
                props[u"IFC"] = dict((k, v) for k, v in ifc.items() if v)
                try:
                    fields = ifc_fields(c, c.cls)
                except Exception:
                    fields = {}
                if fields:
                    props[STRUCAD_PSET] = fields
                if with_props:
                    try:
                        props.update(element_props(c.el, c.doc,
                                                   c.fam, c.tname))
                    except Exception:
                        pass
                item.props = props
                scene.items.append(item)
                done += 1
                if progress is not None and done % 50 == 0:
                    progress(done, total)
    return scene


# ── MAIN ────────────────────────────────────────────────────────────────────

def export_smart(view, path):
    """The view through Revit's IFC exporter (IFC_OPTIONS), wrapped in a
    .seed Studio opens as an IFC. Returns the IFC's size in bytes."""
    import shutil
    import tempfile
    folder = tempfile.mkdtemp(prefix="seed43_")
    name = safe_text(doc.Title) or u"Model"
    for ch in u'<>:"/\\|?*':
        name = name.replace(ch, u"_")
    name += u".ifc"
    try:
        opts = DB.IFCExportOptions()
        opts.FileVersion = DB.IFCVersion.IFC2x3CV2
        opts.FilterViewId = view.Id
        opts.ExportBaseQuantities = True
        opts.WallAndColumnSplitting = False
        for key, value in IFC_OPTIONS:
            opts.AddOption(key, value)
        # IFC export needs an open transaction; nothing it does is kept
        t = DB.Transaction(doc, TITLE)
        t.Start()
        try:
            ok = doc.Export(folder, name, opts)
        finally:
            t.RollBack()
        ifc = os.path.join(folder, name)
        if not ok or not os.path.isfile(ifc):
            raise RuntimeError("Revit's IFC exporter wrote no file")
        write_source_seed(path, ifc, u"ifc", name)
        return os.path.getsize(ifc)
    finally:
        shutil.rmtree(folder, ignore_errors=True)


def pick_output(view):
    folder = os.path.dirname(doc.PathName) if doc.PathName else ""
    name = u"{} {}".format(safe_text(doc.Title), safe_text(view.Name))
    for ch in u'<>:"/\\|?*{}':
        name = name.replace(ch, u"_")
    return forms.save_file(file_ext="seed", default_name=name.strip(),
                           init_dir=folder, title=TITLE)


def grids_levels(to_shared):
    """The project's straight grids and its levels in shared mm, for Seed
    Import to make again."""
    grids = []
    for g in DB.FilteredElementCollector(doc).OfClass(DB.Grid):
        c = g.Curve
        if not isinstance(c, DB.Line):
            continue            # arc grids are left out
        a = to_shared.OfPoint(c.GetEndPoint(0))
        b = to_shared.OfPoint(c.GetEndPoint(1))
        grids.append((element_name(g),
                      (a.X * FT_TO_MM, a.Y * FT_TO_MM, a.Z * FT_TO_MM),
                      (b.X * FT_TO_MM, b.Y * FT_TO_MM, b.Z * FT_TO_MM)))
    levels = []
    for lv in DB.FilteredElementCollector(doc).OfClass(DB.Level):
        p = to_shared.OfPoint(DB.XYZ(0.0, 0.0, lv.ProjectElevation))
        levels.append((element_name(lv), p.Z * FT_TO_MM))
    return grids, levels


def main():
    view = doc.ActiveView
    if not isinstance(view, DB.View3D) or view.IsTemplate:
        forms.alert("Open a 3D view first. The file matches what the active "
                    "3D view shows: its section box, filters and hidden "
                    "elements all apply.", title=TITLE)
        return

    # always the smart model (an IFC with every Revit parameter); the mesh
    # export below stays for when a choice comes back
    choice = SMART

    path = pick_output(view)
    if not path:
        return

    if choice == SMART:
        try:
            with forms.ProgressBar(title="Exporting IFC with Revit's exporter...",
                                   indeterminate=True):
                size = export_smart(view, path)
        except Exception as ex:
            forms.alert(u"Could not export the IFC: {}".format(ex), title=TITLE)
            return
        if forms.alert(u"Saved a smart model ({:,.1f} MB IFC with every Revit "
                       u"parameter) to\n{}\n\nOpen Seed43 Studio to view "
                       u"it?".format(size / 1048576.0, path),
                       title=TITLE, yes=True, no=True):
            import webbrowser
            webbrowser.open("https://seed43.org/studio/")
        return

    # Shared coordinates, as Navisworks reads an NWC by default.
    to_shared = doc.ActiveProjectLocation.GetTotalTransform().Inverse

    total = DB.FilteredElementCollector(doc, view.Id) \
        .WhereElementIsNotElementType().GetElementCount()

    with forms.ProgressBar(title="Reading model... ({value} of {max_value})",
                           cancellable=True) as pb:
        ctx = SeedContext(doc, to_shared, pb, total)
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

    with forms.ProgressBar(title="Reading properties... ({value} of {max_value})") as pb:
        scene = build_scene(ctx, choice == WITH_PROPS,
                            lambda i, n: pb.update_progress(i, n))

    problem = scene.check_tree_order()
    if problem:
        forms.alert("Could not write the file: {}.".format(problem), title=TITLE)
        return

    # Float32 positions lose precision far from 0, so they are written
    # relative to a rounded centre that Studio adds back.
    bb = ctx.bbox
    scene.origin = (round((bb[0] + bb[3]) / 2.0), round((bb[1] + bb[4]) / 2.0),
                    round((bb[2] + bb[5]) / 2.0))

    scene.grids, scene.levels = grids_levels(to_shared)

    orient = view.GetOrientation()
    eye = to_shared.OfPoint(orient.EyePosition)
    fwd = to_shared.OfVector(orient.ForwardDirection)
    up = to_shared.OfVector(orient.UpDirection)
    scene.view = ((eye.X * FT_TO_MM, eye.Y * FT_TO_MM, eye.Z * FT_TO_MM),
                  (fwd.X, fwd.Y, fwd.Z), (up.X, up.Y, up.Z),
                  view.IsPerspective)

    source_name = u"{}.s43nw".format(safe_text(doc.Title) or u"Model")
    with forms.ProgressBar(title="Writing .seed... ({value} of {max_value})") as pb:
        scene.write(path, source_name,
                    progress=lambda i, n: pb.update_progress(i, n))

    vcount, tcount = scene.counts()
    if forms.alert(u"Saved {} elements, {:,} triangles to\n{}\n\nOpen Seed43 "
                   u"Studio to view it?".format(len(scene.items), tcount, path),
                   title=TITLE, yes=True, no=True):
        import webbrowser
        webbrowser.open("https://seed43.org/studio/")


if __name__ == "__main__":
    main()
