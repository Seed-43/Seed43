# -*- coding: utf-8 -*-
# selection_box_level.py
# "Selection Box: Level"
# "Seed43"
# """
# Section-box a thin horizontal slice of the model at a chosen level.
# """
from pyrevit import revit, DB, forms, script

# ── [LIB] Snippets/_selection.py, Snippets/_sectionbox.py ────────────────────
from Snippets._selection import get_levels
from Snippets._sectionbox import (
    mm_to_ft, require_3d_view, model_extents, apply_section_box,
)

doc = revit.doc

# ── CONSTANTS ───────────────────────────────────────────────────────────────

# Cut this far above and below the level line. Enough to catch a floor
# build-up and the slab under it without pulling in the storey above.
OFFSET_MM = 500.0


# ── MAIN ────────────────────────────────────────────────────────────────────

def main():
    view = require_3d_view(doc)
    if view is None:
        forms.alert("Section boxes only exist in 3D views.\n\n"
                    "Open a 3D view and try again.",
                    title="Selection Box: Level")
        script.exit()

    levels = get_levels(doc)
    if not levels:
        forms.alert("No levels found in this model.",
                    title="Selection Box: Level")
        script.exit()

    options = {level.Name: level for level in levels}
    chosen = forms.SelectFromList.show(
        sorted(options.keys()), title="Section box around which level?",
        multiselect=False)
    if not chosen:
        script.exit()
    level = options[chosen]

    # Measured against the model, not the view: the view may already be
    # section-boxed, and measuring that would shrink the box a bit more on
    # every run.
    extents = model_extents(doc)
    if extents is None:
        forms.alert("Could not measure the model extents - nothing with "
                    "geometry was found.", title="Selection Box: Level")
        script.exit()

    offset = mm_to_ft(OFFSET_MM)
    box = DB.BoundingBoxXYZ()
    box.Min = DB.XYZ(extents.Min.X, extents.Min.Y, level.Elevation - offset)
    box.Max = DB.XYZ(extents.Max.X, extents.Max.Y, level.Elevation + offset)

    if not apply_section_box(view, box, "Selection Box: Level"):
        forms.alert("Could not set the section box on this view.",
                    title="Selection Box: Level")
        script.exit()

    forms.alert("Section box set {:.0f}mm either side of {}.".format(
        OFFSET_MM, level.Name), title="Selection Box: Level")


if __name__ == "__main__":
    main()
