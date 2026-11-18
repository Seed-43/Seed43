# -*- coding: utf-8 -*-
"""Relink a DWG link with other colours (Black and White, Invert, Preserve).

Revit fixes a link's colours when it is linked: Reload has no colour
option. So the link is linked again with the colour mode wanted and put
back exactly where it was: every view it is placed in, the same position
and rotation, Draw Layer, cutting and pin, and each view's layer
visibility and overrides. Linking a file that is already linked reuses its
link (colours and all), so the old link goes first. Needs an open
transaction; one undo with whatever wraps it.
"""
import clr
from System.Collections.Generic import List
from Autodesk.Revit import DB

__all__ = ["COLOUR_MODES", "can_relink", "relink"]

COLOUR_MODES = {"bw": DB.ImportColorMode.BlackAndWhite,
                "invert": DB.ImportColorMode.Inverted,
                "preserve": DB.ImportColorMode.Preserved}

P = DB.BuiltInParameter


def _name(e):
    return DB.Element.Name.GetValue(e) if e else u""


def can_relink(doc, instance):
    """(True, "") when the instance's link can be relinked, else (False, why)."""
    t = doc.GetElement(instance.GetTypeId())
    if t is None or not t.IsExternalFileReference():
        return False, u"This is an import, not a link: there is no file to link again."
    ext = t.GetExternalFileReference()
    if ext.GetLinkedFileStatus() != DB.LinkedFileStatus.Loaded:
        return False, u"The link is not loaded (its file may have moved)."
    path = DB.ModelPathUtils.ConvertModelPathToUserVisiblePath(ext.GetAbsolutePath())
    if not path.lower().endswith(u".dwg"):
        return False, u"Only DWG links can be relinked with other colours."
    others = [i for i in DB.FilteredElementCollector(doc).OfClass(DB.ImportInstance)
              if i.GetTypeId() == t.Id]
    if any(not i.ViewSpecific for i in others):
        return False, u"The link is placed in 3D (all views); only view-specific links are handled."
    return True, u""


def _cat_state(view, cat):
    """{layer name: (hidden, overrides)}; "" is the link's own category."""
    out = {}
    for c in [cat] + list(cat.SubCategories):
        try:
            out[u"" if c.Id == cat.Id else c.Name] = (view.GetCategoryHidden(c.Id),
                                                       view.GetCategoryOverrides(c.Id))
        except Exception:
            pass
    return out


def _apply_state(view, cat, state, notes):
    by_name = dict((c.Name, c) for c in cat.SubCategories)
    for name, (hidden, ov) in state.items():
        c = cat if name == u"" else by_name.get(name)
        if c is None:
            continue
        try:
            if view.CanCategoryBeHidden(c.Id):
                view.SetCategoryHidden(c.Id, hidden)
            view.SetCategoryOverrides(c.Id, ov)
        except Exception as err:
            notes.append(u"{0}: {1} not restored ({2})".format(_name(view), name or u"(all)", err))


def _match(doc, inst, target):
    """Turn and move an instance so its transform matches `target`."""
    inst.Pinned = False                     # a new link comes in pinned
    tr = inst.GetTotalTransform()
    axis = tr.BasisX.CrossProduct(target.BasisX)
    if axis.GetLength() > 1e-9:
        DB.ElementTransformUtils.RotateElement(
            doc, inst.Id, DB.Line.CreateBound(tr.Origin, tr.Origin + axis.Normalize()),
            tr.BasisX.AngleTo(target.BasisX))
        tr = inst.GetTotalTransform()
    move = target.Origin - tr.Origin
    if move.GetLength() > 1e-9:
        DB.ElementTransformUtils.MoveElement(doc, inst.Id, move)


def _view_of(doc, inst):
    return doc.GetElement(inst.OwnerViewId)


def relink(doc, instance, mode="bw"):
    """Relink `instance`'s DWG link with colour `mode` ("bw", "invert",
    "preserve"), every placement of it kept. Needs an open transaction.
    Returns ({old view id: new instance}, notes)."""
    ok, why = can_relink(doc, instance)
    if not ok:
        raise ValueError(why)
    t = doc.GetElement(instance.GetTypeId())
    ext = t.GetExternalFileReference()
    path = DB.ModelPathUtils.ConvertModelPathToUserVisiblePath(ext.GetAbsolutePath())
    units = t.get_Parameter(P.IMPORT_DISPLAY_UNITS)
    milli = units is None or u"milli" in (units.AsValueString() or u"").lower()
    scale = t.get_Parameter(P.IMPORT_SCALE).AsDouble()
    keep = []
    for i in DB.FilteredElementCollector(doc).OfClass(DB.ImportInstance):
        if i.GetTypeId() != t.Id:
            continue
        view = _view_of(doc, i)
        keep.append({"view": view, "tr": i.GetTotalTransform(), "pinned": i.Pinned,
                     "layer": i.get_Parameter(P.IMPORT_BACKGROUND).AsInteger(),
                     "cut": i.get_Parameter(P.IMPORT_INSTANCE_CUTTING_IN_VIEW).AsInteger(),
                     "state": _cat_state(view, i.Category)})
    doc.Delete(t.Id)
    doc.Regenerate()

    notes, made = [], []
    for k in keep:
        view = k["view"]
        if not made:
            opts = DB.DWGImportOptions()
            opts.ColorMode = COLOUR_MODES.get(mode, DB.ImportColorMode.BlackAndWhite)
            opts.ThisViewOnly = True
            opts.Placement = DB.ImportPlacement.Origin
            opts.Unit = DB.ImportUnit.Millimeter if milli else DB.ImportUnit.Default
            opts.OrientToView = True
            new_id = clr.Reference[DB.ElementId]()
            if not doc.Link(path, opts, view, new_id):
                raise ValueError(u"Revit would not link {0}".format(path))
            inst = doc.GetElement(new_id.Value)
        else:
            ids = DB.ElementTransformUtils.CopyElements(
                made[0][1], List[DB.ElementId]([made[0][0].Id]), view,
                DB.Transform.Identity, DB.CopyPasteOptions())
            inst = doc.GetElement(list(ids)[0])
        doc.Regenerate()
        _match(doc, inst, k["tr"])
        for bip, value in ((P.IMPORT_BACKGROUND, k["layer"]),
                           (P.IMPORT_INSTANCE_CUTTING_IN_VIEW, k["cut"])):
            p = inst.get_Parameter(bip)
            if p is not None and not p.IsReadOnly and p.AsInteger() != value:
                p.Set(value)
        made.append((inst, view, k))
    new_type = doc.GetElement(made[0][0].GetTypeId())
    sp = new_type.get_Parameter(P.IMPORT_SCALE)
    if sp is not None and not sp.IsReadOnly and abs(sp.AsDouble() - scale) > 1e-9:
        sp.Set(scale)
    doc.Regenerate()
    for inst, view, k in made:
        _apply_state(view, inst.Category, k["state"], notes)
        inst.Pinned = k["pinned"]
    return dict((view.Id, inst) for inst, view, _k in made), notes
