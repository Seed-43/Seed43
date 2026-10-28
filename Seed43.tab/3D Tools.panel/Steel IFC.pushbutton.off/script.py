# -*- coding: utf-8 -*-
"""Import a steel fabrication IFC as native Revit framing and columns.

Tekla Structures and ProStructures IFC2x3 exports: members that are one
section swept along a straight line become structural framing, braces and
columns of a matching family type, plates and bent parts become
DirectShapes, and bolts are skipped. Sections with no matching type are
listed for the user to map, and the choices are remembered. Also takes a
Seed43 Studio .seed that carries an IFC. The work is in
lib/Snippets/_steel_import.py, which 3D Import uses too.
"""
import os

from pyrevit import revit, forms

from Snippets import _steel_import

doc = revit.doc

TITLE = "Steel IFC"


# ── MAIN ────────────────────────────────────────────────────────────────────

def main():
    path = forms.pick_file(
        files_filter="Steel IFC (*.ifc;*.seed)|*.ifc;*.seed|"
                     "IFC (*.ifc)|*.ifc|Seed43 Studio (*.seed)|*.seed",
        title=TITLE)
    if not path:
        return
    if path.lower().endswith(".seed"):
        fmt = _steel_import.seed_format(path)
        if fmt != "ifc":
            forms.alert(u"This .seed holds a {} model, not an IFC. Open it "
                        u"in Seed43 Studio and use Export IFC, then import "
                        u"that file here.".format(fmt or u"unknown"),
                        title=TITLE)
            return
        try:
            ifc_path, source = _steel_import.seed_ifc(path)
        except Exception as ex:
            forms.alert(u"Could not read the .seed: {}".format(ex), title=TITLE)
            return
        try:
            _steel_import.run(doc, ifc_path, source)
        finally:
            try:
                os.remove(ifc_path)
            except Exception:
                pass
        return
    _steel_import.run(doc, path, os.path.basename(path))


if __name__ == "__main__":
    main()
