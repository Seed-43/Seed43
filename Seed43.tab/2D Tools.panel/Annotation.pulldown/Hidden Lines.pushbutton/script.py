# -*- coding: utf-8 -*-
"""Toggle Show Hidden Lines on views, drafting views included.

Revit keeps a Show Hidden Lines setting on every view, but only offers it in
Properties for model views; a drafting view still prints by it (detail items
behind other detail items or masks come out grey and dashed) with no way to
change it by hand. This sets it directly.

Which views:
  a view is open        that view
  a sheet is open       the selected viewports' views, or every view on the
                        sheet when no viewport is selected (schedules skipped)

Which way: if every target is None it turns them on (By Discipline, so the
view's Discipline decides, as Revit's own default), otherwise it turns them
all off. A view whose setting a view template controls is changed on the
template. One undo per click. Runs silently; a short balloon says what it
did, and only a view it cannot change raises a dialog.

Target: Revit 2022-2026, IronPython 2.
"""
from pyrevit import revit, DB, forms

doc = revit.doc
uidoc = revit.uidoc

TITLE = u"Hidden Lines"
NONE, BY_DISCIPLINE = 0, 1
P = DB.BuiltInParameter.VIEW_SHOW_HIDDEN_LINES


def _name(e):
    return DB.Element.Name.GetValue(e) if e is not None else u""


def targets():
    view = doc.ActiveView
    if not isinstance(view, DB.ViewSheet):
        return [view]
    picked = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]
    vids = [e.ViewId for e in picked if isinstance(e, DB.Viewport)]
    if not vids:
        vids = list(view.GetAllPlacedViews())
    out = []
    for vid in vids:
        v = doc.GetElement(vid)
        if v is not None and not isinstance(v, DB.ViewSchedule):
            out.append(v)
    return out


def holder(view):
    """The element whose setting counts: the view, or its template when the
    template controls it. None when it cannot be changed."""
    p = view.get_Parameter(P)
    if p is None:
        return None
    if not p.IsReadOnly:
        return view
    tpl = doc.GetElement(view.ViewTemplateId)
    if tpl is not None:
        tp = tpl.get_Parameter(P)
        if tp is not None and not tp.IsReadOnly:
            return tpl
    return None


def main():
    views = targets()
    if not views:
        forms.alert(u"There are no views here to change.", title=TITLE)
        return
    holders, stuck = {}, []
    for v in views:
        h = holder(v)
        if h is None:
            stuck.append(_name(v))
        else:
            holders[h.Id.IntegerValue if hasattr(h.Id, "IntegerValue") else h.Id.Value] = h
    if not holders:
        forms.alert(u"Show Hidden Lines cannot be changed on:\n\n" + u"\n".join(stuck), title=TITLE)
        return
    all_off = all(h.get_Parameter(P).AsInteger() == NONE for h in holders.values())
    value = BY_DISCIPLINE if all_off else NONE
    with revit.Transaction(u"Show Hidden Lines " + (u"By Discipline" if all_off else u"None")):
        for h in holders.values():
            if h.get_Parameter(P).AsInteger() != value:
                h.get_Parameter(P).Set(value)
    word = u"ON (By Discipline)" if all_off else u"OFF"
    try:
        forms.show_balloon(TITLE, u"Hidden lines {0} on {1} view{2}".format(
            word, len(holders), u"" if len(holders) == 1 else u"s"))
    except Exception:
        pass
    if stuck:
        forms.alert(u"Changed {0}, but these could not be changed:\n\n{1}".format(
            len(holders), u"\n".join(stuck)), title=TITLE)


if __name__ == "__main__":
    main()
