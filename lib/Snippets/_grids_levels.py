# -*- coding: utf-8 -*-
"""Grids and levels from an imported model, as Revit grids and levels.

    from Snippets._grids_levels import make_grids_levels

Used by 3D Import for every format: steel IFC (IfcGrid, storeys), StruCad
.bsw (grid lines; the format has no levels) and mesh .seed files that carry
them (3D Seed writes "grids" and "levels" into the file).

Nothing already in the project is moved or renamed:
  - a level at the same height (within 1 mm) as one in the file is kept and
    no new one is made; a new level whose name is taken gets " (2)" etc.
  - a grid whose name is taken is left out and reported.
IronPython 2.
"""
from pyrevit import DB

__all__ = ["make_grids_levels", "GridsLevelsResult"]

MM_PER_FT = 304.8
SAME_LEVEL_FT = 1.0 / MM_PER_FT     # levels this close are one level
MIN_GRID_FT = 0.01                  # shorter lines cannot be grids

try:
    TEXT = unicode      # IronPython 2
except NameError:
    TEXT = str


class GridsLevelsResult(object):
    def __init__(self):
        self.levels_made = []       # names
        self.levels_kept = []       # (file name, project level it matched)
        self.grids_made = []        # names
        self.grids_skipped = []     # (name, why)


def _name(el):
    try:
        return TEXT(DB.Element.Name.__get__(el))
    except Exception:
        return u""


def _free_name(name, taken):
    if name not in taken:
        return name
    n = 2
    while u"{} ({})".format(name, n) in taken:
        n += 1
    return u"{} ({})".format(name, n)


def make_levels(doc, levels, to_xyz, result):
    """levels: [(name, elevation mm)] in the file's coordinates."""
    have = list(DB.FilteredElementCollector(doc).OfClass(DB.Level))
    taken = set(_name(lv) for lv in have)
    for name, z in sorted(levels, key=lambda l: l[1]):
        want = to_xyz(0.0, 0.0, z).Z
        match = [lv for lv in have
                 if abs(lv.ProjectElevation - want) <= SAME_LEVEL_FT]
        if match:
            result.levels_kept.append((name, _name(match[0])))
            continue
        lv = DB.Level.Create(doc, want)
        # Level.Create's elevation follows the level type's base (project
        # base point or survey point): correct it to the internal height
        off = want - lv.ProjectElevation
        if abs(off) > 1e-9:
            lv.Elevation = lv.Elevation + off
        name = _free_name(TEXT(name or u"Level").strip() or u"Level", taken)
        try:
            lv.Name = name
        except Exception:
            name = _name(lv)
        taken.add(name)
        have.append(lv)
        result.levels_made.append(name)


def make_grids(doc, grids, to_xyz, result):
    """grids: objects with label, a, b (end points, mm, file coordinates)."""
    taken = set(_name(g) for g in
                DB.FilteredElementCollector(doc).OfClass(DB.Grid))
    for g in grids:
        label = TEXT(g.label or u"").strip()
        if label and label in taken:
            result.grids_skipped.append((label, u"a grid of that name is "
                                                u"already in the project"))
            continue
        p0 = to_xyz(*g.a)
        p1 = to_xyz(*g.b)
        p1 = DB.XYZ(p1.X, p1.Y, p0.Z)       # grids are drawn level
        if p0.DistanceTo(p1) < MIN_GRID_FT:
            result.grids_skipped.append((label or u"(no name)", u"too short"))
            continue
        try:
            grid = DB.Grid.Create(doc, DB.Line.CreateBound(p0, p1))
            if label:
                grid.Name = label
                taken.add(label)
            result.grids_made.append(label or _name(grid))
        except Exception as ex:
            result.grids_skipped.append((label or u"(no name)", TEXT(ex)))


def make_grids_levels(doc, grids, levels, to_xyz):
    """Make the file's levels, then its grids, in one transaction.
    to_xyz(x, y, z) maps file mm to a Revit internal DB.XYZ."""
    result = GridsLevelsResult()
    if not grids and not levels:
        return result
    t = DB.Transaction(doc, "Import grids and levels")
    t.Start()
    try:
        make_levels(doc, levels or [], to_xyz, result)
        make_grids(doc, grids or [], to_xyz, result)
        t.Commit()
    except Exception:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        raise
    return result


def report_md(result):
    """Markdown lines for an import report."""
    out = []
    if result.levels_made or result.levels_kept:
        out.append(u"**{}** levels made{}.".format(
            len(result.levels_made),
            u", {} already in the project".format(len(result.levels_kept))
            if result.levels_kept else u""))
    if result.grids_made or result.grids_skipped:
        out.append(u"**{}** grids made.".format(len(result.grids_made)))
    if result.grids_skipped:
        out.append(u"### Grids left out ({})".format(len(result.grids_skipped)))
        for name, why in result.grids_skipped[:100]:
            out.append(u"- {}: {}".format(name, why))
    return out
