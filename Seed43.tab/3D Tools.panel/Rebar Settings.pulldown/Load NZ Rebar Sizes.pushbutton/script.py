# -*- coding: utf-8 -*-
# "Load NZ Rebar Sizes"
# "Seed43"
# """
# Load the NZ rebar bar types (D, HD, HR, R and Mesh SE) into the open model,
# set up as in 3488 King and Queen: diameters, bend diameters, deformation,
# material, description and hook lengths. The definitions ship with the tool
# in nz_rebar_sizes.json, so no other model needs to be open.
# """
import io
import os.path as op
import json

from pyrevit import revit, DB, forms
from Autodesk.Revit.DB.Structure import (RebarBarType, RebarHookType,
                                         RebarDeformationType)

from Snippets import _dialogs as dlg

doc = revit.doc

# ── CONSTANTS ───────────────────────────────────────────────────────────────

TITLE = u"Load NZ Rebar Sizes"
DATA = op.join(op.dirname(__file__), u"nz_rebar_sizes.json")
MM = 304.8


# ── HELPERS (private to this file) ──────────────────────────────────────────

def _name(element):
    return DB.Element.Name.GetValue(element)


def _ft(mm):
    return mm / MM


def _material(name, cache):
    """The material of that name, made (plain) if the model lacks it."""
    if not name:
        return None
    if name not in cache:
        found = [m for m in DB.FilteredElementCollector(doc).OfClass(DB.Material)
                 if m.Name == name]
        cache[name] = found[0].Id if found else DB.Material.Create(doc, name)
    return cache[name]


class _Quiet(DB.IFailuresPreprocessor):
    """Drop warnings, such as 'rebar does not show surface patterns' raised
    once per bar for the NZ materials; they say nothing to act on."""

    def PreprocessFailures(self, accessor):
        for msg in accessor.GetFailureMessages():
            if msg.GetSeverity() == DB.FailureSeverity.Warning:
                accessor.DeleteWarning(msg)
        return DB.FailureProcessingResult.Continue


class _BarItem(forms.TemplateListItem):
    """A bar in the pick list, ticked to start with."""

    @property
    def name(self):
        bar = self.item
        return u"{0}   ({1:g} mm)".format(bar["name"], bar["nominal"])


# ── CORE LOGIC ──────────────────────────────────────────────────────────────

def apply(bar_type, bar, hooks, materials):
    """Set a bar type up from one entry of the data file."""
    bar_type.BarNominalDiameter = _ft(bar["nominal"])
    bar_type.BarModelDiameter = _ft(bar["model"])
    bar_type.StandardBendDiameter = _ft(bar["bend"])
    bar_type.StandardHookBendDiameter = _ft(bar["hook_bend"])
    bar_type.StirrupTieBendDiameter = _ft(bar["tie_bend"])
    bar_type.MaximumBendRadius = _ft(bar["max_bend_radius"])
    bar_type.DeformationType = (RebarDeformationType.Plain if bar["deformation"] == u"Plain"
                                else RebarDeformationType.Deformed)
    mat = _material(bar.get("material"), materials)
    if mat is not None:
        p = bar_type.get_Parameter(DB.BuiltInParameter.MATERIAL_ID_PARAM)
        if p is not None and not p.IsReadOnly:
            p.Set(mat)
    if bar.get("description"):
        p = bar_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_DESCRIPTION)
        if p is not None and not p.IsReadOnly:
            p.Set(bar["description"])
    # hooks: only the hook types this model has, by name
    for hook_name, h in bar["hooks"].items():
        hook = hooks.get(hook_name)
        if hook is None:
            continue
        bar_type.SetHookPermission(hook.Id, bool(h.get("allowed")))
        bar_type.SetAutoCalcHookLengths(hook.Id, bool(h.get("auto")))
        if not h.get("auto"):
            if "length" in h:
                bar_type.SetHookLength(hook.Id, _ft(h["length"]))
            if "offset" in h:
                try:
                    bar_type.SetHookOffsetLength(hook.Id, _ft(h["offset"]))
                except Exception:
                    pass                # not every hook angle has an offset


# ── UI / ENTRY POINT ────────────────────────────────────────────────────────

def main():
    with io.open(DATA, "r", encoding="utf-8") as f:
        data = json.loads(f.read())
    existing = dict((_name(b), b) for b in
                    DB.FilteredElementCollector(doc).OfClass(RebarBarType))
    bars = data["bars"]
    picked = forms.SelectFromList.show(
        [_BarItem(b, checked=True) for b in bars],
        title=TITLE, button_name=u"Load", multiselect=True, width=420, height=640)
    if not picked:
        return

    clash = [b["name"] for b in picked if b["name"] in existing]
    update = True
    if clash:
        key = dlg.choice(
            u"{0} of these are already in the model ({1}{2}).".format(
                len(clash), u", ".join(clash[:6]), u"..." if len(clash) > 6 else u""),
            [("cancel", u"Cancel"), ("skip", u"Leave them as they are"),
             ("update", u"Update them to match")],
            title=TITLE)
        if key is None or key == "cancel":
            return
        update = key == "update"

    hooks = dict((_name(h), h) for h in DB.FilteredElementCollector(doc).OfClass(RebarHookType))
    made, updated, failed = [], [], []
    materials = {}
    t = DB.Transaction(doc, u"Load NZ Rebar Sizes")
    options = t.GetFailureHandlingOptions()
    options.SetFailuresPreprocessor(_Quiet())
    t.SetFailureHandlingOptions(options)
    t.Start()
    try:
        for bar in picked:
            try:
                bar_type = existing.get(bar["name"])
                if bar_type is None:
                    bar_type = RebarBarType.Create(doc)
                    bar_type.Name = bar["name"]
                    made.append(bar["name"])
                elif update:
                    updated.append(bar["name"])
                else:
                    continue
                apply(bar_type, bar, hooks, materials)
            except Exception as err:
                failed.append(u"{0}: {1}".format(bar["name"], err))
        t.Commit()
    except Exception:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        raise

    lines = [u"{0} added, {1} updated.".format(len(made), len(updated))]
    if failed:
        lines.append(u"")
        lines.append(u"Not loaded:")
        lines.extend(failed)
    dlg.message(u"\n".join(lines), title=TITLE)


if __name__ == "__main__":
    main()
