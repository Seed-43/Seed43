# -*- coding: utf-8 -*-
"""3D Import: import a model into Revit, working out what to do from the file.

One window (SeedImport.xaml) for every format, its options remembered in
.user/Seed Import/options.json:

  steel IFC, StruCad .bsw, or a .seed holding one
      _steel_import.run: native members, steel plates, bolts, shape
      families, with the window's options.
  mesh .seed from 3D Seed or Navisworks' Export to Studio
      every item one DirectShape in its category (Walls, Structural
      Framing..., Generic Models when unknown or not allowed), named as in
      Studio, Mark and Comments filled in, one material per colour.

"Section map..." opens the saved section choices (Snippets._section_map).
"""
import json
import os

from System.Collections.Generic import List

from pyrevit import revit, DB, forms, script

from seed_reader import SeedError, read_seed
from Snippets import _steel_import
from Snippets._icons import set_header_icon
from Snippets._userdata import user_path
from Snippets.seed43_theme import apply_seed43_dimensions, apply_seed43_palette

doc = revit.doc

TITLE = "3D Import"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OPTIONS_PATH = user_path("Seed Import", "options.json")   # the old name, kept

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


# ── OPTIONS WINDOW ──────────────────────────────────────────────────────────

FILE_FILTER = ("Models (*.seed;*.ifc;*.bsw)|*.seed;*.ifc;*.bsw|"
               "Seed43 Studio (*.seed)|*.seed|IFC (*.ifc)|*.ifc|"
               "StruCad (*.bsw)|*.bsw")
STEEL = "steel"         # file kinds
MESH = "mesh"
OTHER = "other"


def file_kind(path):
    """(kind, description) of a file: steel (IFC, StruCad or a .seed holding
    one), a mesh .seed, or something 3D Import cannot take."""
    low = path.lower()
    if _steel_import.is_bsw(path):
        return STEEL, u"StruCad model"
    if low.endswith(".ifc"):
        return STEEL, u"IFC model"
    if low.endswith(".seed"):
        fmt = _steel_import.seed_format(path)
        if fmt == "ifc":
            return STEEL, u"Seed43 Studio file holding an IFC model"
        if fmt == "bsw":
            return STEEL, u"Seed43 Studio file holding a StruCad model"
        if fmt == "navis":
            return MESH, u"Seed43 mesh (3D Seed or Navisworks)"
        return OTHER, (u"Seed43 Studio file holding a {} model, which only "
                       u"Studio can open".format(fmt or u"unknown"))
    return OTHER, u"Not a file 3D Import reads"


def load_options():
    opts = dict(_steel_import.DEFAULT_OPTIONS)
    opts["path"] = u""
    try:
        with open(OPTIONS_PATH, "r") as f:
            opts.update(json.load(f))
    except Exception:
        pass
    return opts


def save_options(opts):
    try:
        with open(OPTIONS_PATH, "wb") as f:
            f.write(_steel_import.json_text(opts, indent=2,
                                            sort_keys=True).encode("utf-8"))
    except Exception:
        pass


class ImportWindow(forms.WPFWindow):
    CHECKS = (("natives", "natives_cb"), ("ask_map", "ask_map_cb"),
              ("shape_families", "families_cb"), ("plates", "plates_cb"),
              ("plate_cuts", "cuts_cb"), ("plate_holes", "holes_cb"),
              ("grids_levels", "grids_cb"),
              ("bolts", "bolts_cb"),
              ("skip_base_plate_bolts", "skip_base_cb"))
    # an indented option and the one it depends on
    FOLLOWS = (("natives_cb", "ask_map_cb"), ("plates_cb", "cuts_cb"),
               ("bolts_cb", "skip_base_cb"))
    PLACES = (("shared", "place_shared_rb"), ("internal", "place_internal_rb"),
              ("centre", "place_centre_rb"))

    def __init__(self):
        forms.WPFWindow.__init__(self, os.path.join(SCRIPT_DIR, "SeedImport.xaml"))
        apply_seed43_palette(self, SCRIPT_DIR)
        apply_seed43_dimensions(self, SCRIPT_DIR)
        set_header_icon(self, SCRIPT_DIR)
        self.go = False
        self.options = load_options()
        for key, name in self.CHECKS:
            getattr(self, name).IsChecked = bool(self.options.get(key, True))
        for key, name in self.PLACES:
            getattr(self, name).IsChecked = self.options.get("placement") == key
        # the file box starts empty every time; the last file only decides
        # where Browse opens
        last = self.options.get("path") or u""
        self.last_dir = os.path.dirname(last) if last else u""
        self._set_path(u"")
        for parent, child in self.FOLLOWS:
            getattr(self, parent).Checked += self._follow
            getattr(self, parent).Unchecked += self._follow
        self._follow(None, None)
        self.browse_btn.Click += self._on_browse
        self.map_btn.Click += self._on_map
        self.import_btn.Click += self._on_import
        self.cancel_btn.Click += lambda s, a: self.Close()
        self.win_close_btn.Click += lambda s, a: self.Close()

    def _follow(self, sender, args):
        """An indented option only applies when the one above it is on."""
        for parent, child in self.FOLLOWS:
            getattr(self, child).IsEnabled = bool(getattr(self, parent).IsChecked) \
                and self.kind != MESH

    def _set_path(self, path):
        from System.Windows import Visibility
        self.path_tb.Text = path
        self.subtitle_tb.Text = os.path.basename(path) if path else u""
        if path:
            self.kind, text = file_kind(path)
            self.kind_tb.Text = text
        else:
            self.kind = None
        steel = self.kind != MESH
        for key, name in self.CHECKS:
            # grids and levels come from every format
            getattr(self, name).IsEnabled = steel or name == "grids_cb"
        self.mesh_note_tb.Visibility = Visibility.Collapsed if steel \
            else Visibility.Visible
        self.import_btn.IsEnabled = self.kind in (STEEL, MESH)
        if steel:
            self._follow(None, None)

    def _on_browse(self, sender, args):
        start = os.path.dirname(self.path_tb.Text) if self.path_tb.Text \
            else self.last_dir
        path = forms.pick_file(files_filter=FILE_FILTER, init_dir=start or u"",
                               title=TITLE)
        if path:
            self._set_path(path)

    def _on_map(self, sender, args):
        """The Section Map for the file chosen above: its member sections and
        the type each becomes; changes are saved for later imports."""
        from Snippets._dialogs import message
        from Snippets._section_map import edit_section_map
        path = self.path_tb.Text
        if not path or self.kind != STEEL:
            message(u"Browse to a steel IFC, StruCad .bsw or a .seed holding "
                    u"one first: the Section Map shows that file's sections.",
                    title=TITLE)
            return
        with forms.ProgressBar(title="Reading the model and family types...") as pb:
            pb.update_progress(1, 3)
            try:
                model = _steel_import.read_model(path)
            except Exception as ex:
                message(u"Could not read the file: {}".format(ex), title=TITLE)
                return
            pb.update_progress(2, 3)
            entries = _steel_import.loaded_entries(doc)
            entries.extend(_steel_import.library_entries(doc.Application))
            rows = _steel_import.file_section_rows(
                model, entries, _steel_import.load_map())
            pb.update_progress(3, 3)
        choices = edit_section_map(rows, _steel_import.SPECIAL_CHOICES,
                                   os.path.basename(path))
        if choices is not None:
            _steel_import.save_section_edits(rows, choices)

    def _on_import(self, sender, args):
        opts = {"path": self.path_tb.Text}
        for key, name in self.CHECKS:
            opts[key] = bool(getattr(self, name).IsChecked)
        for key, name in self.PLACES:
            if getattr(self, name).IsChecked:
                opts["placement"] = key
        self.options = opts
        save_options(opts)
        self.go = True
        self.Close()


# ── MAIN ────────────────────────────────────────────────────────────────────

MESH_PLACES = {"shared": PLACE_SHARED, "internal": PLACE_INTERNAL,
               "centre": PLACE_CENTRE}


def main():
    win = ImportWindow()
    win.ShowDialog()
    if not win.go:
        return
    opts = win.options
    path = opts["path"]
    kind = file_kind(path)[0]
    if kind == STEEL and path.lower().endswith(".seed"):
        import_ifc(path, opts)
    elif kind == STEEL:
        _steel_import.run(doc, path, os.path.basename(path), opts)
    elif kind == MESH:
        import_mesh(path, MESH_PLACES.get(opts.get("placement"), PLACE_SHARED),
                    opts.get("grids_levels", True))


# ── MESH IMPORT ─────────────────────────────────────────────────────────────

class _Grid(object):
    def __init__(self, label, a, b):
        self.label = label
        self.a = a
        self.b = b


def seed_grids_levels(model):
    """The grids and levels a mesh .seed carries (3D Seed writes them; mm,
    less the origin like the vertices): ([_Grid], [(name, z)])."""
    grids = []
    for g in model.meta.get("grids") or []:
        try:
            grids.append(_Grid(safe_text(g.get("n", u"")),
                               tuple(float(v) for v in g["a"]),
                               tuple(float(v) for v in g["b"])))
        except Exception:
            pass
    levels = []
    for lv in model.meta.get("levels") or []:
        try:
            levels.append((safe_text(lv.get("n", u"")), float(lv["z"])))
        except Exception:
            pass
    return grids, levels


def import_mesh(path, placement, with_grids=True):
    try:
        model = read_seed(path)
    except SeedError as ex:
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

    shapes = forms.CommandSwitchWindow.show(
        [SHAPE_SOLID, SHAPE_MESH], message="Make the shapes as")
    if not shapes:
        return

    to_xyz = make_mapper(model, placement)
    grids, levels = seed_grids_levels(model) if with_grids else ([], [])
    cats = CategoryMap(doc)
    source = os.path.basename(path)
    made, new_mats, failed, fallback = 0, 0, [], {}

    group = DB.TransactionGroup(doc, TITLE)
    group.Start()
    grid_result = None
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
                    step = u"reading the triangles"
                    try:
                        solid = shapes == SHAPE_SOLID
                        geo = build_shape(model, it, mats, to_xyz, solid)
                        if not geo:
                            failed.append((name, u"no usable triangles"))
                            continue
                        step = u"making the DirectShape"
                        ds = DB.DirectShape.CreateElement(doc, cat_id)
                        ds.ApplicationId = "Seed43 Studio"
                        ds.ApplicationDataId = source
                        step = u"setting its shape"
                        try:
                            ds.SetShape(geo)
                        except Exception:
                            if not solid:
                                raise
                            # Revit refuses some solids: the mesh still goes in
                            step = u"setting its shape as a mesh"
                            ds.SetShape(build_shape(model, it, mats, to_xyz,
                                                    False))
                        step = u"naming it"
                        ds.SetName(clean_name(name))
                        set_text(ds, DB.BuiltInParameter.ALL_MODEL_MARK, name)
                        set_text(ds, DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS,
                                 u" > ".join(model.path(it["node"])))
                        batch_made += 1
                    except Exception as ex:
                        failed.append((name, u"{} ({}): {}".format(
                            step, type(ex).__name__, safe_text(ex))))
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
        # grids and levels last: nothing they do can stop the shapes
        if grids or levels:
            from Snippets._grids_levels import make_grids_levels
            try:
                grid_result = make_grids_levels(doc, grids, levels, to_xyz)
            except Exception as ex:
                failed.append((u"Grids and levels", safe_text(ex)))
        group.Assimilate()
    except Exception:
        if group.HasStarted():
            group.RollBack()
        raise

    report(path, model, made, failed, fallback, new_mats, placement,
           grid_result)


def import_ifc(path, options=None):
    """A .seed made from an IFC or a StruCad .bsw: hand that model to the
    steel importer."""
    try:
        ifc_path, source = _steel_import.seed_ifc(path)
    except Exception as ex:
        forms.alert(u"Could not read the IFC in the .seed: {}".format(ex),
                    title=TITLE)
        return
    try:
        _steel_import.run(doc, ifc_path, source, options)
    finally:
        try:
            os.remove(ifc_path)
        except Exception:
            pass


def report(path, model, made, failed, fallback, new_mats, placement,
           grid_result=None):
    out = script.get_output()
    out.print_md(u"## 3D Import\n`{}`".format(path))
    out.print_md(u"From **{}** ({}). Placed in: {}.".format(
        safe_text(model.meta.get("document", u"")),
        safe_text(model.meta.get("product", u"")), placement))
    out.print_md(u"**{}** DirectShapes made, **{}** new materials.".format(
        made, new_mats))
    if grid_result is not None:
        from Snippets._grids_levels import report_md
        for line in report_md(grid_result):
            out.print_md(line)
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
