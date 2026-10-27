# -*- coding: utf-8 -*-
"""Import a Seed43 Studio .seed mesh into Revit as DirectShapes.

Reads the files 3D Seed and Navisworks' Export to Studio write. Every item
becomes one DirectShape in its category (Walls, Structural Framing...,
Generic Models when the category is unknown or not allowed), named as in
Studio, with Mark and Comments filled in and one Revit material per colour.
Placed in shared coordinates by default, so a model exported with 3D Seed
or from an NWC lands where it came from.
"""
import os

from System.Collections.Generic import List

from pyrevit import revit, DB, forms, script

from seed_reader import SeedError, read_seed
from Snippets import _steel_import

doc = revit.doc

TITLE = "Seed Import"

# ── CONSTANTS ───────────────────────────────────────────────────────────────

MM_TO_FT = 1.0 / 304.8
BATCH = 200             # items per transaction, so Revit stays responsive
MIN_AREA_FT2 = 1e-10    # smaller triangles make TessellatedShapeBuilder fail

PLACE_SHARED = "Shared coordinates (where it was exported from)"
PLACE_INTERNAL = "Internal coordinates"
PLACE_CENTRE = "Centred on the internal origin"

SHAPE_SOLID = "Solids where possible (sections cut them, slower)"
SHAPE_MESH = "Meshes only (faster, not cut in sections)"

# IFC classes from a Navisworks IFC model, to Revit categories.
IFC_CATEGORIES = {
    "IfcBeam": DB.BuiltInCategory.OST_StructuralFraming,
    "IfcMember": DB.BuiltInCategory.OST_StructuralFraming,
    "IfcColumn": DB.BuiltInCategory.OST_StructuralColumns,
    "IfcFooting": DB.BuiltInCategory.OST_StructuralFoundation,
    "IfcPile": DB.BuiltInCategory.OST_StructuralFoundation,
    "IfcWall": DB.BuiltInCategory.OST_Walls,
    "IfcWallStandardCase": DB.BuiltInCategory.OST_Walls,
    "IfcSlab": DB.BuiltInCategory.OST_Floors,
    "IfcRoof": DB.BuiltInCategory.OST_Roofs,
    "IfcCovering": DB.BuiltInCategory.OST_Ceilings,
    "IfcDoor": DB.BuiltInCategory.OST_Doors,
    "IfcWindow": DB.BuiltInCategory.OST_Windows,
    "IfcStair": DB.BuiltInCategory.OST_Stairs,
    "IfcStairFlight": DB.BuiltInCategory.OST_Stairs,
    "IfcRamp": DB.BuiltInCategory.OST_Ramps,
    "IfcRailing": DB.BuiltInCategory.OST_StairsRailing,
    "IfcCurtainWall": DB.BuiltInCategory.OST_Walls,
    "IfcFurnishingElement": DB.BuiltInCategory.OST_Furniture,
    "IfcPipeSegment": DB.BuiltInCategory.OST_PipeCurves,
    "IfcPipeFitting": DB.BuiltInCategory.OST_PipeFitting,
    "IfcDuctSegment": DB.BuiltInCategory.OST_DuctCurves,
    "IfcDuctFitting": DB.BuiltInCategory.OST_DuctFitting,
    "IfcCableCarrierSegment": DB.BuiltInCategory.OST_CableTray,
}

BAD_NAME_CHARS = u'\\:{}[]|;<>?`~'


# ── HELPERS ─────────────────────────────────────────────────────────────────

try:
    TEXT = unicode      # IronPython 2
except NameError:
    TEXT = str          # IronPython 3


def safe_text(value):
    try:
        return TEXT(value)
    except Exception:
        return u"?"


def clean_name(name):
    for ch in BAD_NAME_CHARS:
        name = name.replace(ch, u"_")
    return name.strip() or u"Seed43"


def parse_rgb(text):
    try:
        text = text.lstrip("#")
        return int(text[0:2], 16), int(text[2:4], 16), int(text[4:6], 16)
    except Exception:
        return 180, 180, 180


class SwallowWarnings(DB.IFailuresPreprocessor):
    """Warnings (overlaps, small shapes) would stop every batch with a
    dialog; they are dropped, errors still roll the batch back."""

    def PreprocessFailures(self, accessor):
        for f in accessor.GetFailureMessages():
            if f.GetSeverity() == DB.FailureSeverity.Warning:
                accessor.DeleteWarning(f)
        return DB.FailureProcessingResult.Continue


# ── CATEGORIES ──────────────────────────────────────────────────────────────

class CategoryMap(object):
    """Studio category name to a category a DirectShape may use."""

    def __init__(self, document):
        self.doc = document
        self.by_name = {}
        for cat in document.Settings.Categories:
            try:
                if cat.CategoryType == DB.CategoryType.Model:
                    self.by_name[safe_text(cat.Name)] = cat.Id
            except Exception:
                pass
        self.generic = DB.ElementId(DB.BuiltInCategory.OST_GenericModel)
        self._cache = {}

    def _valid(self, cat_id):
        try:
            return cat_id is not None and \
                DB.DirectShape.IsValidCategoryId(cat_id, self.doc)
        except Exception:
            return False

    def resolve(self, name):
        """(category id, True when it is the requested one)."""
        hit = self._cache.get(name)
        if hit is None:
            cat_id = self.by_name.get(name)
            if cat_id is None and name in IFC_CATEGORIES:
                cat_id = DB.ElementId(IFC_CATEGORIES[name])
            hit = (cat_id, True) if self._valid(cat_id) else (self.generic, False)
            self._cache[name] = hit
        return hit


# ── MATERIALS ───────────────────────────────────────────────────────────────

class MaterialMap(object):
    """One Revit material per name (or colour), reused when it exists."""

    def __init__(self, document):
        self.doc = document
        self.by_name = {}
        for m in DB.FilteredElementCollector(document).OfClass(DB.Material):
            self.by_name[safe_text(m.Name)] = m.Id
        self._cache = {}
        self.created = 0

    def get(self, part, item_mat):
        rgb = parse_rgb(part.get("rgb", ""))
        op = float(part.get("op", 1.0))
        name = part.get("mat") or item_mat or u""
        if not name:
            name = u"Seed43 #{:02x}{:02x}{:02x}".format(*rgb)
            if op < 0.999:
                name += u" {}%".format(int(round(op * 100)))
        name = clean_name(name)
        key = name
        mid = self._cache.get(key)
        if mid is not None:
            return mid
        mid = self.by_name.get(name)
        if mid is None:
            mid = DB.Material.Create(self.doc, name)
            mat = self.doc.GetElement(mid)
            mat.Color = DB.Color(rgb[0], rgb[1], rgb[2])
            mat.Transparency = int(round((1.0 - op) * 100))
            self.by_name[name] = mid
            self.created += 1
        self._cache[key] = mid
        return mid


# ── GEOMETRY ────────────────────────────────────────────────────────────────

def make_mapper(model, placement):
    """Function (x, y, z in file mm) -> internal XYZ in feet."""
    ox, oy, oz = [float(v) for v in model.origin]
    if placement == PLACE_CENTRE:
        ox = oy = oz = 0.0
    if placement == PLACE_SHARED:
        # GetTotalTransform maps shared coordinates to internal ones.
        t = doc.ActiveProjectLocation.GetTotalTransform()
        bx, by, bz, o = t.BasisX, t.BasisY, t.BasisZ, t.Origin
        m = (bx.X, by.X, bz.X, o.X, bx.Y, by.Y, bz.Y, o.Y,
             bx.Z, by.Z, bz.Z, o.Z)
    else:
        m = (1.0, 0.0, 0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 0.0, 1.0, 0.0)
    a, b, c, d, e, f, g, h, i, j, k, l = m

    def to_xyz(x, y, z):
        x = (x + ox) * MM_TO_FT
        y = (y + oy) * MM_TO_FT
        z = (z + oz) * MM_TO_FT
        return DB.XYZ(a * x + b * y + c * z + d,
                      e * x + f * y + g * z + h,
                      i * x + j * y + k * z + l)
    return to_xyz


def _area2(p, q, r):
    ux, uy, uz = q.X - p.X, q.Y - p.Y, q.Z - p.Z
    vx, vy, vz = r.X - p.X, r.Y - p.Y, r.Z - p.Z
    cx = uy * vz - uz * vy
    cy = uz * vx - ux * vz
    cz = ux * vy - uy * vx
    return cx * cx + cy * cy + cz * cz


def build_shape(model, item, mats, to_xyz, solids):
    """Geometry list for one item, or None when nothing usable is left."""
    pos = model.positions
    idx = model.indices
    builder = DB.TessellatedShapeBuilder()
    builder.Target = DB.TessellatedShapeBuilderTarget.AnyGeometry \
        if solids else DB.TessellatedShapeBuilderTarget.Mesh
    builder.Fallback = DB.TessellatedShapeBuilderFallback.Mesh
    xyz = {}
    faces = 0
    item_mat = item.get("mat")
    for part in item.get("parts", []):
        first, count = part["t"]
        mat_id = mats.get(part, item_mat)
        # One face set per part: a part that does not close into a solid
        # falls back to a mesh on its own, not the whole item.
        builder.OpenConnectedFaceSet(solids)
        for t in range(first, first + count):
            pts = []
            for v in (idx[3 * t], idx[3 * t + 1], idx[3 * t + 2]):
                p = xyz.get(v)
                if p is None:
                    p = to_xyz(pos[3 * v], pos[3 * v + 1], pos[3 * v + 2])
                    xyz[v] = p
                pts.append(p)
            if _area2(pts[0], pts[1], pts[2]) < MIN_AREA_FT2:
                continue
            builder.AddFace(DB.TessellatedFace(List[DB.XYZ](pts), mat_id))
            faces += 1
        builder.CloseConnectedFaceSet()
    if not faces:
        return None
    builder.Build()
    result = builder.GetBuildResult()
    if result.Outcome == DB.TessellatedShapeBuilderOutcome.Nothing:
        return None
    return result.GetGeometricalObjects()


def set_text(el, bip, value):
    try:
        p = el.get_Parameter(bip)
        if p is not None and not p.IsReadOnly and value:
            p.Set(value)
    except Exception:
        pass


# ── DIALOGS ─────────────────────────────────────────────────────────────────

def item_category(model, item):
    cat = item.get("cat")
    if cat:
        return safe_text(cat)
    node = model.nodes[item["node"]] if 0 <= item.get("node", -1) < len(model.nodes) else {}
    return safe_text(node.get("c") or u"Generic Models")


def pick_categories(model, items):
    counts = {}
    for it in items:
        name = item_category(model, it)
        counts[name] = counts.get(name, 0) + 1
    if len(counts) <= 1:
        return set(counts)
    labels = {}
    options = []
    for name in sorted(counts):
        label = u"{}  ({})".format(name, counts[name])
        labels[label] = name
        options.append(forms.TemplateListItem(label, checked=True))
    picked = forms.SelectFromList.show(
        options, title=TITLE + " - categories to import",
        multiselect=True, button_name="Import")
    if not picked:
        return None
    return set(labels[p] for p in picked)


# ── MAIN ────────────────────────────────────────────────────────────────────

def main():
    path = forms.pick_file(file_ext="seed", title=TITLE)
    if not path:
        return
    try:
        model = read_seed(path)
    except SeedError as ex:
        fmt = _steel_import.seed_format(path)
        if fmt == "ifc":
            import_ifc(path)
        elif fmt == "bsw":
            forms.alert(u"This .seed holds a StruCad model. Open it in "
                        u"Seed43 Studio and use Export IFC, then import the "
                        u"IFC with Steel IFC.", title=TITLE)
        else:
            forms.alert(safe_text(ex), title=TITLE)
        return
    except Exception as ex:
        forms.alert(u"Could not read the file: {}".format(ex), title=TITLE)
        return

    items = [it for it in model.items if it.get("parts") and not model.hidden(it)]
    if not items:
        forms.alert("The file has no visible geometry.", title=TITLE)
        return

    wanted = pick_categories(model, items)
    if not wanted:
        return
    items = [it for it in items if item_category(model, it) in wanted]

    placement = forms.CommandSwitchWindow.show(
        [PLACE_SHARED, PLACE_INTERNAL, PLACE_CENTRE],
        message="Place the model at")
    if not placement:
        return
    shapes = forms.CommandSwitchWindow.show(
        [SHAPE_SOLID, SHAPE_MESH], message="Make the shapes as")
    if not shapes:
        return

    to_xyz = make_mapper(model, placement)
    cats = CategoryMap(doc)
    source = os.path.basename(path)
    made, new_mats, failed, fallback = 0, 0, [], {}

    group = DB.TransactionGroup(doc, TITLE)
    group.Start()
    try:
        mats = MaterialMap(doc)
        with forms.ProgressBar(title="Importing... ({value} of {max_value})",
                               cancellable=True) as pb:
            for start in range(0, len(items), BATCH):
                if pb.cancelled:
                    break
                t = DB.Transaction(doc, TITLE)
                opts = t.GetFailureHandlingOptions()
                opts.SetFailuresPreprocessor(SwallowWarnings())
                t.SetFailureHandlingOptions(opts)
                t.Start()
                batch_made = 0
                for it in items[start:start + BATCH]:
                    node = model.nodes[it["node"]]
                    name = safe_text(node.get("n", u"")) or u"Item"
                    cat_name = item_category(model, it)
                    cat_id, exact = cats.resolve(cat_name)
                    if not exact:
                        fallback[cat_name] = fallback.get(cat_name, 0) + 1
                    ds = None
                    try:
                        geo = build_shape(model, it, mats, to_xyz,
                                          shapes == SHAPE_SOLID)
                        if not geo:
                            failed.append((name, u"no usable triangles"))
                            continue
                        ds = DB.DirectShape.CreateElement(doc, cat_id)
                        ds.ApplicationId = "Seed43 Studio"
                        ds.ApplicationDataId = source
                        ds.SetShape(geo)
                        ds.SetName(clean_name(name))
                        set_text(ds, DB.BuiltInParameter.ALL_MODEL_MARK, name)
                        set_text(ds, DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS,
                                 u" > ".join(model.path(it["node"])))
                        batch_made += 1
                    except Exception as ex:
                        failed.append((name, safe_text(ex)))
                        if ds is not None:
                            # an empty DirectShape would be left behind
                            try:
                                doc.Delete(ds.Id)
                            except Exception:
                                pass
                if t.Commit() == DB.TransactionStatus.Committed:
                    made += batch_made
                    new_mats += mats.created
                else:
                    failed.append((u"Items {} to {}".format(
                        start + 1, start + BATCH), u"Revit rolled the batch back"))
                # materials made in a rolled-back batch are gone with it
                mats = MaterialMap(doc)
                pb.update_progress(min(start + BATCH, len(items)), len(items))
            cancelled = pb.cancelled
        if cancelled:
            group.RollBack()
            return
        group.Assimilate()
    except Exception:
        if group.HasStarted():
            group.RollBack()
        raise

    report(path, model, made, failed, fallback, new_mats, placement)


def import_ifc(path):
    """A .seed made from an IFC: hand the IFC to the Steel IFC importer."""
    try:
        ifc_path, source = _steel_import.seed_ifc(path)
    except Exception as ex:
        forms.alert(u"Could not read the IFC in the .seed: {}".format(ex),
                    title=TITLE)
        return
    try:
        _steel_import.run(doc, ifc_path, source)
    finally:
        try:
            os.remove(ifc_path)
        except Exception:
            pass


def report(path, model, made, failed, fallback, new_mats, placement):
    out = script.get_output()
    out.print_md(u"## Seed Import\n`{}`".format(path))
    out.print_md(u"From **{}** ({}). Placed in: {}.".format(
        safe_text(model.meta.get("document", u"")),
        safe_text(model.meta.get("product", u"")), placement))
    out.print_md(u"**{}** DirectShapes made, **{}** new materials.".format(
        made, new_mats))
    if fallback:
        out.print_md(u"### Imported as Generic Models\nNot a category a "
                     u"DirectShape can use in this project:")
        for name in sorted(fallback):
            out.print_md(u"- {} ({})".format(name, fallback[name]))
    if failed:
        out.print_md(u"### Not imported ({})".format(len(failed)))
        for name, why in failed[:200]:
            out.print_md(u"- {}: {}".format(name, why))
        if len(failed) > 200:
            out.print_md(u"- ... and {} more".format(len(failed) - 200))


if __name__ == "__main__":
    main()
