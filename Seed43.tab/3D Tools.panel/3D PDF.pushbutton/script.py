# -*- coding: utf-8 -*-
"""Export the active 3D view to a 3D PDF (pdf3d_export does the work, and
pySheets' 3D PDF format uses it too)."""
import os

from pyrevit import revit, DB, forms

from pdf3d_export import export_view, safe_text, EMPTY, CANCELLED

doc = revit.doc

TITLE = "3D PDF"


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

    with forms.ProgressBar(title="Reading model... ({value} of {max_value})",
                           cancellable=True) as pb:
        result = export_view(doc, view, path, pb,
                             lambda i, n: pb.update_progress(i, n))
    if result == EMPTY:
        forms.alert("Nothing to export: the view has no visible model "
                    "geometry.", title=TITLE)
    elif result == CANCELLED:
        return


if __name__ == "__main__":
    main()
