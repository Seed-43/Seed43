# -*- coding: utf-8 -*-
"""pyRecent thumbnails: a model-only picture of a project file.

A temporary isometric 3D view is zoomed to the model's own elements (not
topography, links, grids or levels, which would set the extents far too
wide), then shown with everything on, exported to PNG and thrown away. The view is
made inside a TransactionGroup that is rolled back, so making a thumbnail
leaves the model exactly as it was - no stray view, nothing to save.

Raw Revit API only: this runs from startup.py's DocumentOpened handler,
where importing pyrevit modules would slow every file open.
"""

import os
import time
import shutil
import tempfile

from Autodesk.Revit import DB
from System.Collections.Generic import List

from Snippets import _recent

__all__ = ["generate"]

# ── CONSTANTS ───────────────────────────────────────────────────────────────

def _bics(*names):
    """BuiltInCategory values by name, skipping any this Revit lacks
    (OST_Toposolid only exists from 2024)."""
    return [getattr(DB.BuiltInCategory, n) for n in names
            if hasattr(DB.BuiltInCategory, n)]


# Hidden in the picture: view furniture only. Everything else - topography,
# links, grids, levels - stays on, as in a normal 3D view.
_HIDE = _bics("OST_SectionBox", "OST_VolumeOfInterest", "OST_Cameras",
              "OST_Viewers", "OST_ProjectBasePoint", "OST_SharedBasePoint",
              "OST_IOS_GeoSite")

# Shown, but never allowed to set the framing: ground and site run far past
# the building and would shrink it to a speck.
_NOT_EXTENTS = set(int(b) for b in _bics(
    "OST_Topography", "OST_Toposolid", "OST_Site", "OST_Planting",
    "OST_Parking", "OST_Entourage", "OST_Cameras", "OST_Rooms", "OST_Areas",
    "OST_MEPSpaces", "OST_ProjectBasePoint", "OST_SharedBasePoint",
    "OST_IOS_GeoSite", "OST_Levels", "OST_Grids", "OST_Lines",
    "OST_SketchLines", "OST_RvtLinks", "OST_PointClouds",
    "OST_VolumeOfInterest", "OST_SectionBox", "OST_Viewers", "OST_Sheets"))

_MAX_SIZE = 3000.0      # ft: anything bigger is a stray far-off element


# ── HELPERS (private to this file) ──────────────────────────────────────────

def _eid(element_id):
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue


def _boxes(doc):
    """Bounding boxes of the model's own building elements."""
    out = []
    for el in (DB.FilteredElementCollector(doc).WhereElementIsNotElementType()
               .WhereElementIsViewIndependent()):
        try:
            cat = el.Category
            if (cat is None or cat.CategoryType != DB.CategoryType.Model
                    or _eid(cat.Id) in _NOT_EXTENTS or cat.Name.startswith(u"<")
                    or isinstance(el, (DB.View, DB.RevitLinkInstance,
                                       DB.ImportInstance))):
                continue
            bb = el.get_BoundingBox(None)
            if bb is None or max(bb.Max.X - bb.Min.X, bb.Max.Y - bb.Min.Y,
                                 bb.Max.Z - bb.Min.Z) > _MAX_SIZE:
                continue
            out.append((bb.Min, bb.Max))
        except Exception:
            continue
    return out


def _link_boxes(doc):
    """Boxes of the loaded links, used only when the model has nothing else."""
    out = []
    for link in DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance):
        bb = link.get_BoundingBox(None)
        if bb is not None:
            out.append((bb.Min, bb.Max))
    return out


def _trim(boxes):
    """Drop strays: with enough elements, keep those whose centre sits in
    the middle 98% on every axis, so one element left far off at the
    survey point cannot shrink the building."""
    if len(boxes) < 20:
        return boxes
    keep = boxes
    for axis in ("X", "Y", "Z"):
        centres = sorted((getattr(a, axis) + getattr(b, axis)) / 2.0 for a, b in keep)
        lo = centres[int(len(centres) * 0.01)]
        hi = centres[int(len(centres) * 0.99) - 1]
        pad = (hi - lo) * 0.1
        keep = [(a, b) for a, b in keep
                if lo - pad <= (getattr(a, axis) + getattr(b, axis)) / 2.0 <= hi + pad]
    return keep or boxes


def _model_box(doc):
    """Box to frame: the model's own elements; never topography or links,
    unless the model has nothing else to show."""
    boxes = _trim(_boxes(doc)) or _link_boxes(doc)
    if not boxes:
        return None
    lo = DB.XYZ(min(a.X for a, b in boxes), min(a.Y for a, b in boxes),
                min(a.Z for a, b in boxes))
    hi = DB.XYZ(max(b.X for a, b in boxes), max(b.Y for a, b in boxes),
                max(b.Z for a, b in boxes))
    size = max(hi.X - lo.X, hi.Y - lo.Y, hi.Z - lo.Z)
    if size < 0.1:
        return None
    pad = max(size * 0.02, 0.5)
    box = DB.BoundingBoxXYZ()
    box.Min = DB.XYZ(lo.X - pad, lo.Y - pad, lo.Z - pad)
    box.Max = DB.XYZ(hi.X + pad, hi.Y + pad, hi.Z + pad)
    return box


def _aim(view, box):
    """Isometric camera from front-left-above (Revit's Front/Left corner of
    the ViewCube) at the box centre."""
    c = (box.Min + box.Max) * 0.5
    diag = max(box.Min.DistanceTo(box.Max), 1.0)
    eye = DB.XYZ(c.X - diag * 0.66, c.Y - diag * 0.66, c.Z + diag * 0.57)
    forward = (c - eye).Normalize()
    right = forward.CrossProduct(DB.XYZ.BasisZ).Normalize()
    up = right.CrossProduct(forward).Normalize()
    view.SetOrientation(DB.ViewOrientation3D(eye, up, forward))


def _hide_clutter(view, doc):
    for bic in _HIDE:
        try:
            cat = DB.Category.GetCategory(doc, bic)
            if cat is not None:
                view.SetCategoryHidden(cat.Id, True)
        except Exception:
            pass


def _fit_crop(view, box, aspect=4.0 / 3.0, pad=0.03):
    """Crop the view to the box as seen by the camera, at the card's 4:3.

    The same as zooming to the model's extents: nothing is cut away, so the
    topography and links run on to the edges of the picture, but only the
    model's own elements decide what is in frame.
    """
    crop = view.CropBox
    to_view = crop.Transform.Inverse
    xs, ys = [], []
    for x in (box.Min.X, box.Max.X):
        for y in (box.Min.Y, box.Max.Y):
            for z in (box.Min.Z, box.Max.Z):
                p = to_view.OfPoint(DB.XYZ(x, y, z))
                xs.append(p.X)
                ys.append(p.Y)
    w, h = max(xs) - min(xs), max(ys) - min(ys)
    w, h = w * (1 + 2 * pad), h * (1 + 2 * pad)
    if w / h < aspect:
        w = h * aspect
    else:
        h = w / aspect
    cx, cy = (max(xs) + min(xs)) / 2.0, (max(ys) + min(ys)) / 2.0
    fitted = DB.BoundingBoxXYZ()
    fitted.Transform = crop.Transform
    fitted.Min = DB.XYZ(cx - w / 2.0, cy - h / 2.0, crop.Min.Z)
    fitted.Max = DB.XYZ(cx + w / 2.0, cy + h / 2.0, crop.Max.Z)
    view.CropBox = fitted
    view.CropBoxActive = True
    view.CropBoxVisible = False


def _make_view(doc, key):
    """Create and set up the thumbnail view. Needs an open transaction."""
    type_id = doc.GetDefaultElementTypeId(DB.ElementTypeGroup.ViewType3D)
    view = DB.View3D.CreateIsometric(doc, type_id)
    view.Name = u"__pyRecent_{0}__".format(key[:10])
    try:
        view.DetailLevel = DB.ViewDetailLevel.Medium
        view.DisplayStyle = DB.DisplayStyle.ShadingWithEdges
    except Exception:
        pass
    _hide_clutter(view, doc)
    box = _model_box(doc)
    if box is not None:
        _aim(view, box)
        doc.Regenerate()            # the crop box follows the new camera
        _fit_crop(view, box)
    return view.Id


def _export(doc, view_id, base, pixel_size):
    opts = DB.ImageExportOptions()
    opts.ExportRange = DB.ExportRange.SetOfViews
    opts.SetViewsAndSheets(List[DB.ElementId]([view_id]))
    opts.FilePath = base
    opts.FitDirection = DB.FitDirectionType.Horizontal
    opts.HLRandWFViewsFileType = DB.ImageFileType.PNG
    opts.ShadowViewsFileType = DB.ImageFileType.PNG
    opts.ImageResolution = DB.ImageResolution.DPI_150
    opts.PixelSize = pixel_size
    opts.ZoomType = DB.ZoomFitType.FitToPage
    doc.ExportImage(opts)


# ── CORE LOGIC ──────────────────────────────────────────────────────────────

def generate(doc, path, pixel_size=560):
    """Make the thumbnail for `doc` and record it. Returns the file name.

    Returns None, without raising, when anything goes wrong - a missing
    thumbnail must never stop a file opening.
    """
    key = _recent.path_key(path)
    # A new name each time, so WPF never shows a cached copy of the old one.
    filename = u"{0}_{1}.png".format(key, int(time.time()))
    temp_dir = tempfile.mkdtemp(prefix="pyrecent_")
    try:
        group = DB.TransactionGroup(doc, "pyRecent thumbnail")
        group.Start()
        try:
            t = DB.Transaction(doc, "pyRecent thumbnail view")
            t.Start()
            view_id = _make_view(doc, key)
            t.Commit()
            _export(doc, view_id, os.path.join(temp_dir, "thumb"), pixel_size)
        finally:
            # Rolled back, not deleted: the model ends up exactly as it was.
            if group.HasStarted() and not group.HasEnded():
                group.RollBack()

        produced = [f for f in os.listdir(temp_dir) if f.lower().endswith(".png")]
        if not produced:
            return None
        shutil.copy2(os.path.join(temp_dir, produced[0]),
                     os.path.join(_recent.thumb_dir(), filename))
        _recent.set_thumbnail(path, filename)
        return filename
    except Exception:
        return None
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
