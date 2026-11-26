# -*- coding: utf-8 -*-
"""Shared Remap: swap a family's numbered helper parameters for the shared
parameters of the same kind and number in a shared parameter file.

    Bool 001, NG_Gb Bool 004        ->  Bool 01, Bool 04
    NG_Gl Length 001, Length 12     ->  Length 01, Length 12
    Integer 01                      ->  Integer 01
    Angle 001, NG_Ga Angle 002      ->  Angle 01, Angle 02

Run in an open family. FamilyManager.ReplaceParameter keeps values,
formulas, dimension labels and nested associations, and renames the
parameter in every formula that uses it. A parameter that is already shared
from another file goes shared -> family -> new shared, since Revit has no
direct shared-to-shared swap. One undo.

Two things it leaves alone, and lists:
  * a parameter that drives something inside a group or array: Revit only
    allows that change in Edit Group mode, which a script cannot enter
    (unlink it there, remap, link it back)
  * a parameter with no shared partner, or one of another data type

Afterwards every value in every type and every formula is compared with
before (the renames applied) and any difference is listed.

The rules are PATTERNS below: a regex on the family parameter's name, the
shared parameter group to look in, and the shared name to make from the
number. Add a line to cover another kind.

Target: Revit 2022-2026, IronPython 2.
"""
import re

from pyrevit import revit, DB, forms, script

doc = revit.doc
TITLE = u"Shared Remap"

# (regex on the family name, shared group, shared name from the number)
PATTERNS = [
    (re.compile(r"^(?:[A-Z]{2}_Gb )?Bool (\d+)$"), u"Bool", u"Bool {0:02d}"),
    (re.compile(r"^(?:[A-Z]{2}_Gl )?Length (\d+)$"), u"Length", u"Length {0:02d}"),
    (re.compile(r"^(?:[A-Z]{2}_Gi )?Integer (\d+)$"), u"Integer", u"Integer {0:02d}"),
    (re.compile(r"^(?:[A-Z]{2}_Ga )?Angle (\d+)$"), u"Angle", u"Angle {0:02d}"),
]


# ── HELPERS ─────────────────────────────────────────────────────────────────

def _eid(element_id):
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue


def shared_definitions(app):
    df = app.OpenSharedParameterFile()
    if df is None:
        return None
    defs = {}
    for g in df.Groups:
        for d in g.Definitions:
            defs[(g.Name, d.Name)] = d
    return defs


def grouped_parameters(fm):
    """Names of family parameters that drive elements inside groups/arrays."""
    out = set()
    for g in DB.FilteredElementCollector(doc).OfClass(DB.Group):
        for i in g.GetMemberIds():
            m = doc.GetElement(i)
            for p in m.Parameters:
                a = fm.GetAssociatedFamilyParameter(p)
                if a is not None:
                    out.add(a.Definition.Name)
    return out


def plan(fm, defs):
    """([(family parameter, definition)], [(name, target, why)])."""
    names = set(p.Definition.Name for p in fm.Parameters)
    grouped = grouped_parameters(fm)
    pairs, skipped, taken = [], [], set()
    for p in sorted(fm.Parameters, key=lambda x: x.Definition.Name):
        if _eid(p.Id) < 0:
            continue
        name = p.Definition.Name
        for rx, group, fmt in PATTERNS:
            m = rx.match(name)
            if not m:
                continue
            target = fmt.format(int(m.group(1)))
            d = defs.get((group, target))
            if d is not None and p.IsShared and p.GUID == d.GUID:
                break                                   # already that one
            if d is None:
                skipped.append((name, target, u"no '{0}' in group {1} of the shared file".format(target, group)))
            elif d.GetDataType() != p.Definition.GetDataType():
                skipped.append((name, target, u"different data type"))
            elif (target in names and target != name) or target in taken:
                skipped.append((name, target, u"'{0}' is already in this family".format(target)))
            elif name in grouped:
                skipped.append((name, target, u"drives a group / array member: unlink it in "
                                              u"Edit Group, remap, link it back"))
            else:
                pairs.append((p, d))
                taken.add(target)
            break
    return pairs, skipped


def snapshot(fm):
    vals, forms_ = {}, {}
    types = list(fm.Types)
    for p in fm.Parameters:
        n = p.Definition.Name
        forms_[n] = p.Formula
        row = []
        for t in types:
            st = p.StorageType
            if st == DB.StorageType.Double:
                v = t.AsDouble(p)
                row.append(round(v, 6) if v is not None else None)
            elif st == DB.StorageType.Integer:
                row.append(t.AsInteger(p))
            elif st == DB.StorageType.String:
                row.append(t.AsString(p))
            else:
                row.append(None)
        vals[n] = row
    return vals, forms_


def compare(before, after, renames):
    bv, bf = before
    av, af = after
    diffs = []
    for old, row in bv.items():
        new = renames.get(old, old)
        if av.get(new) != row:
            diffs.append((old, new, u"values differ"))
    order = sorted(renames.items(), key=lambda kv: -len(kv[0]))
    for old, f in bf.items():
        new = renames.get(old, old)
        want = f
        if want:
            for a, b in order:
                want = re.sub(r"(?<![\w.])" + re.escape(a) + r"(?![\w.])", b, want)
        if (af.get(new) or None) != (want or None):
            diffs.append((old, new, u"formula: {0}".format(af.get(new))))
    return diffs


# ── MAIN ────────────────────────────────────────────────────────────────────

def main():
    if not doc.IsFamilyDocument:
        forms.alert(u"Open a family first: Shared Remap works on the family you are editing.",
                    title=TITLE)
        return
    app = doc.Application
    fm = doc.FamilyManager
    original = app.SharedParametersFilename
    path = original
    if path:
        choice = forms.alert(u"Remap to the shared parameter file:\n\n{0}".format(path),
                             title=TITLE, options=[u"Use this file", u"Pick another...", u"Cancel"])
        if choice in (None, u"Cancel"):
            return
        if choice == u"Pick another...":
            path = forms.pick_file(file_ext="txt", title=u"Shared parameter file")
    else:
        path = forms.pick_file(file_ext="txt", title=u"Shared parameter file")
    if not path:
        return

    try:
        app.SharedParametersFilename = path
        defs = shared_definitions(app)
        if defs is None:
            forms.alert(u"That file could not be read as a shared parameter file.", title=TITLE)
            return
        pairs, skipped = plan(fm, defs)
        if not pairs:
            forms.alert(u"Nothing to remap." + (u"\n\nLeft alone:\n" + u"\n".join(
                u"{0}: {1}".format(n, why) for n, _t, why in skipped) if skipped else u""), title=TITLE)
            return
        labels = [u"{0}  ->  {1}{2}".format(p.Definition.Name, d.Name,
                                            u"   (replaces an old shared one)" if p.IsShared else u"")
                  for p, d in pairs]
        items = [forms.TemplateListItem(lab, checked=True) for lab in labels]
        chosen = forms.SelectFromList.show(items, title=TITLE + u": untick any to leave",
                                           multiselect=True, button_name=u"Remap")
        if not chosen:
            return
        chosen = set(chosen)
        pairs = [pd for pd, lab in zip(pairs, labels) if lab in chosen]

        before = snapshot(fm)
        renames = {}
        t = DB.Transaction(doc, u"Shared Remap")
        t.Start()
        try:
            for p, d in pairs:
                old_name = p.Definition.Name
                group, inst = p.Definition.GetGroupTypeId(), p.IsInstance
                if p.IsShared:
                    p = fm.ReplaceParameter(p, u"zz_remap_" + d.Name, group, inst)
                fm.ReplaceParameter(p, d, group, inst)
                renames[old_name] = d.Name
            t.Commit()
        except Exception:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
            raise
        diffs = compare(before, snapshot(fm), renames)
    finally:
        app.SharedParametersFilename = original

    out = script.get_output()
    out.print_md(u"## Shared Remap: {0}".format(doc.Title))
    out.print_md(u"Shared file: `{0}`".format(path))
    out.print_table([[a, b] for a, b in sorted(renames.items())],
                    columns=[u"Was", u"Now (shared)"], title=u"Remapped ({0})".format(len(renames)))
    if skipped:
        out.print_table([[n, t, why] for n, t, why in skipped],
                        columns=[u"Parameter", u"Would be", u"Left alone because"], title=u"Left alone")
    if diffs:
        out.print_table([list(x) for x in diffs], columns=[u"Was", u"Now", u"Difference"],
                        title=u"Check these: something changed")
    else:
        out.print_md(u"**Check:** every value in every type and every formula is the same as before.")


if __name__ == "__main__":
    main()
