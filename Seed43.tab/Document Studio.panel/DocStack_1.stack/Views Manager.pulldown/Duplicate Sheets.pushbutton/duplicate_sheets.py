# -*- coding: utf-8 -*-
# duplicate_sheets.py
#
# Duplicates sheets and/or views selected from the project browser.
# Sheets:  Duplicate Empty  |  Duplicate with Sheet Detailing  |  Duplicate with Views
# Views:   Duplicate  |  Duplicate with Detailing  |  Duplicate as Dependent
#
# When duplicating a sheet "with Views", each placed view is duplicated using
# whichever view-duplication option you choose, and the duplicate is placed on
# the new sheet with its model where the original's is, and the same rotation
# and title position.

import re
import time

from pyrevit import revit, DB, forms, script
from System.Collections.Generic import List

doc   = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()
output = script.get_output()

# Set True to re-enable the title block / viewport placement diagnostics.
DEBUG = False

# ── Output styles ─────────────────────────────────────────────────────────────
output.add_style("""
body {
    background-color: #232933;
    color: #F4FAFF;
    font-family: Consolas, Courier New, monospace;
    padding: 20px;
}
.header {
    color: #2B933F;
    font-weight: bold;
    font-size: 1.2em;
}
.sheet {
    color: #F4FAFF;
    padding-left: 15px;
}
.rev {
    color: #8B9199;
    padding-left: 30px;
}
.warn {
    color: #E0A040;
    padding-left: 30px;
}
.line {
    color: #3B4553;
}
""")

# Lines are buffered and flushed once per sheet/view rather than once per line.
# Each print_html call is a separate write into the output browser, so the old
# line-by-line logging was most of the wait -- but the log still needs to stream
# as the run progresses, not appear all at once at the end.
_out_buffer = []


def _write(css_class, text):
    _out_buffer.append("<div class='{}'>{}</div>".format(css_class, text))


def flush_output():
    """Push everything buffered so far to the output window."""
    if _out_buffer:
        output.print_html("".join(_out_buffer))
        del _out_buffer[:]


def print_header(text):
    _write("header", text)


def print_separator():
    _write("line", "------------------------------------")


def print_success(text):
    _write("sheet", text)


def print_warning(text):
    _write("warn", "WARNING: {}".format(text))


def print_error(text):
    _write("warn", "ERROR: {}".format(text))


def print_dim(text):
    _write("rev", "-> {}".format(text))


def print_debug(text):
    if DEBUG:
        _write("rev", "-> {}".format(text))


# ── Helpers ───────────────────────────────────────────────────────────────────

def eid_int(element_id):
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue


# Existing sheet numbers / view names are collected once and then kept up to
# date as elements are created. Re-collecting per duplicate meant a full model
# sweep for every single sheet.
_sheet_numbers = None
_view_names = None


def _existing_sheet_numbers():
    global _sheet_numbers
    if _sheet_numbers is None:
        _sheet_numbers = {
            s.SheetNumber
            for s in DB.FilteredElementCollector(doc)
                       .OfClass(DB.ViewSheet)
                       .ToElements()
        }
    return _sheet_numbers


def _existing_view_names():
    """Built lazily -- the sheet-only modes never need it."""
    global _view_names
    if _view_names is None:
        _view_names = {
            v.Name
            for v in DB.FilteredElementCollector(doc)
                       .OfClass(DB.View)
                       .WhereElementIsNotElementType()
                       .ToElements()
        }
    return _view_names


def unique_sheet_number(base_number):
    """Return base_number + suffix matching Revit's own pattern: (1), (2)..."""
    existing = _existing_sheet_numbers()
    if base_number not in existing:
        existing.add(base_number)
        return base_number
    counter = 1
    while True:
        candidate = "{} ({})".format(base_number, counter)
        if candidate not in existing:
            existing.add(candidate)
            return candidate
        counter += 1


def sequence_sheet_number(source_number, pattern=None):
    """Next free sheet number continuing the sequence of source_number.

    Without a pattern the trailing digits of the source number are counted
    on:  45.03.01 -> 45.03.02,  S2001 -> S2002,  A-09 -> A-10.
    With a pattern such as 45.03.xx or S20xx the text before the x's is the
    fixed part and the x's set the digit width; the count carries on from
    the source number when it starts with that fixed part, otherwise from 1.
    Returns None when no sequence can be read (the caller falls back to
    Revit's own numbering).
    """
    if pattern:
        m = re.match(r"^(.*?)([xX#]+)$", pattern.strip())
        if not m:
            return None
        prefix, width = m.group(1), len(m.group(2))
        tail = source_number[len(prefix):]
        if source_number.lower().startswith(prefix.lower()) and tail.isdigit():
            counter = int(tail)
        else:
            counter = 0
    else:
        m = re.match(r"^(.*?)(\d+)$", source_number)
        if not m:
            return None
        prefix, width = m.group(1), len(m.group(2))
        counter = int(m.group(2))

    existing = _existing_sheet_numbers()
    while True:
        counter += 1
        candidate = prefix + str(counter).zfill(width)
        if candidate not in existing:
            existing.add(candidate)
            return candidate


def unique_view_name(base_name):
    """Return base_name + suffix that does not clash with existing view names."""
    existing = _existing_view_names()
    if base_name not in existing:
        existing.add(base_name)
        return base_name
    counter = 2
    while True:
        candidate = "{} ({})".format(base_name, counter)
        if candidate not in existing:
            existing.add(candidate)
            return candidate
        counter += 1


def duplicate_view(view, dup_option):
    """Duplicate a single view and return the new ElementId."""
    if not view.CanViewBeDuplicated(dup_option):
        # Fall back to plain duplicate if requested option is unsupported
        for fallback in [
            DB.ViewDuplicateOption.Duplicate,
            DB.ViewDuplicateOption.WithDetailing,
        ]:
            if view.CanViewBeDuplicated(fallback):
                return view.Duplicate(fallback)
        return None
    return view.Duplicate(dup_option)


def match_viewport(source, target):
    """Make target sit on its sheet exactly as source does. Returns how.

    Rotation and the title label first, since both change the box. Then the
    position. A viewport's box takes in its annotations, so a duplicate made
    without detailing - tags and dimensions left behind - has a smaller box
    with a different centre, and matching box centres shifted the model
    itself across the sheet (issue #4). From Revit 2023 the view's
    projection-to-sheet transform lines up the model instead, whatever the
    box does. Schedules and legends have no such transform, nor has 2022,
    and fall back to the box centre.
    """
    try:
        if target.Rotation != source.Rotation:
            target.Rotation = source.Rotation
    except Exception:
        pass
    for prop in ("LabelOffset", "LabelLineLength"):
        try:
            setattr(target, prop, getattr(source, prop))
        except Exception:
            pass
    # Outlines and transforms only reflect those edits after a regenerate.
    doc.Regenerate()
    try:
        if source.HasViewportTransforms() and target.HasViewportTransforms():
            delta = (source.GetProjectionToSheetTransform().Origin
                     - target.GetProjectionToSheetTransform().Origin)
            if not delta.IsZeroLength():
                DB.ElementTransformUtils.MoveElement(
                    doc, target.Id, DB.XYZ(delta.X, delta.Y, 0))
            return "model"
    except Exception:
        pass        # 2022 has no viewport transforms; anything else: the box
    target.SetBoxCenter(source.GetBoxCenter())
    return "box"


def get_title_blocks(view_sheet):
    """Return the title block elements placed on a sheet (may be empty)."""
    return list(
        DB.FilteredElementCollector(doc, view_sheet.Id)
        .OfCategory(DB.BuiltInCategory.OST_TitleBlocks)
        .WhereElementIsNotElementType()
        .ToElements()
    )


def copy_writable_parameters(source, target, skip_names=None):
    """Copy all writable parameter values from source element to target."""
    skip_names = skip_names or set()

    # LookupParameter is a linear scan of the element's parameters, so calling
    # it once per source parameter is quadratic. Map the target's writable
    # parameters by name up front instead.
    target_params = {}
    for tp in target.Parameters:
        try:
            if tp.IsReadOnly or not tp.Definition:
                continue
            name = tp.Definition.Name
            if name not in target_params:
                target_params[name] = tp
        except Exception:
            pass

    for p in source.Parameters:
        try:
            if p.IsReadOnly:
                continue
            if not p.Definition:
                continue
            name = p.Definition.Name
            if name in skip_names:
                continue
            tp = target_params.get(name)
            if tp is None:
                continue
            st = p.StorageType
            if st == DB.StorageType.String:
                val = p.AsString()
                if val is not None:
                    tp.Set(val)
            elif st == DB.StorageType.Integer:
                tp.Set(p.AsInteger())
            elif st == DB.StorageType.Double:
                tp.Set(p.AsDouble())
            elif st == DB.StorageType.ElementId:
                tp.Set(p.AsElementId())
        except Exception:
            pass


def _element_key(element, view):
    """Identify an element by what it is and where it sits in the view, so the
    same element can be found again in a duplicate of that view."""
    try:
        box = element.get_BoundingBox(view)
    except Exception:
        box = None
    where = (round(box.Min.X, 4), round(box.Min.Y, 4),
             round(box.Max.X, 4), round(box.Max.Y, 4)) if box else None
    return (element.GetType().Name, eid_int(element.GetTypeId()), where)


def match_view_graphics(source, target):
    """Carry the source view's hidden-line and graphic settings to its copy.

    Show Hidden Lines is set outright (it can be missing from Properties in
    a drafting view, so the parameter copy can miss it). Per-element graphic
    overrides are not kept by Duplicate with Detailing, so each override is
    re-applied to the matching element of the copy.
    """
    prm = DB.BuiltInParameter.VIEW_SHOW_HIDDEN_LINES
    try:
        sp, tp = source.get_Parameter(prm), target.get_Parameter(prm)
        if sp is not None and tp is not None and not tp.IsReadOnly:
            if tp.AsInteger() != sp.AsInteger():
                tp.Set(sp.AsInteger())
    except Exception:
        pass

    try:
        src_els = [e for e in DB.FilteredElementCollector(doc, source.Id)
                              .WhereElementIsNotElementType()
                   if e.OwnerViewId == source.Id]
        overridden = []
        for e in src_els:
            ogs = source.GetElementOverrides(e.Id)
            if _has_overrides(ogs):
                overridden.append((e, ogs))
        if not overridden:
            return 0

        pool = {}
        for e in DB.FilteredElementCollector(doc, target.Id) \
                   .WhereElementIsNotElementType():
            if e.OwnerViewId == target.Id:
                pool.setdefault(_element_key(e, target), []).append(e)
        for lst in pool.values():
            lst.sort(key=lambda x: eid_int(x.Id))

        applied = 0
        for e, ogs in sorted(overridden, key=lambda p: eid_int(p[0].Id)):
            lst = pool.get(_element_key(e, source))
            if lst:
                target.SetElementOverrides(lst.pop(0).Id, ogs)
                applied += 1
        return applied
    except Exception:
        return 0


def _has_overrides(ogs):
    """True when an OverrideGraphicSettings changes anything."""
    def vid(x):
        return eid_int(x)
    return bool(
        ogs.Halftone
        or ogs.Transparency != 0
        or ogs.ProjectionLineWeight != -1
        or ogs.CutLineWeight != -1
        or vid(ogs.ProjectionLinePatternId) != -1
        or vid(ogs.CutLinePatternId) != -1
        or vid(ogs.SurfaceForegroundPatternId) != -1
        or vid(ogs.SurfaceBackgroundPatternId) != -1
        or (ogs.ProjectionLineColor is not None and ogs.ProjectionLineColor.IsValid)
        or (ogs.CutLineColor is not None and ogs.CutLineColor.IsValid)
        or (ogs.SurfaceForegroundPatternColor is not None
            and ogs.SurfaceForegroundPatternColor.IsValid)
        or (ogs.SurfaceBackgroundPatternColor is not None
            and ogs.SurfaceBackgroundPatternColor.IsValid))


# ── Collect selection ─────────────────────────────────────────────────────────

# A caller that runs this file (e.g. Nagel Working) can set PRESET in the
# globals to skip every dialog: {"sheets": [ViewSheet], "copies": "3"}. It means
# Duplicate with Views + with Detailing, continue the sequence, keep detail numbers.
PRESET = globals().get("PRESET")

selected_elements = PRESET["sheets"] if PRESET else revit.get_selection().elements

sheets = [e for e in selected_elements if isinstance(e, DB.ViewSheet)
          and not e.IsPlaceholder]
views  = [e for e in selected_elements
          if isinstance(e, DB.View)
          and not isinstance(e, DB.ViewSheet)
          and not e.IsTemplate]

if not sheets and not views:
    # Fall back to the pyRevit view picker if nothing is pre-selected
    picked = forms.select_views(
        title="Select Sheets and/or Views to Duplicate",
        use_selection=False,
        filterfunc=lambda v: (
            (isinstance(v, DB.ViewSheet) and not v.IsPlaceholder)
            or (not isinstance(v, DB.ViewSheet)
                and not v.IsTemplate
                and v.CanViewBeDuplicated(DB.ViewDuplicateOption.Duplicate))
        )
    )
    if not picked:
        script.exit()
    sheets = [v for v in picked if isinstance(v, DB.ViewSheet)]
    views  = [v for v in picked if not isinstance(v, DB.ViewSheet)]

# ── Choose duplication mode ───────────────────────────────────────────────────

options = []
if sheets:
    options += [
        "SHEETS: Duplicate Empty",
        "SHEETS: Duplicate with Sheet Detailing",
        "SHEETS: Duplicate with Views",
    ]
if views:
    options += [
        "VIEWS: Duplicate",
        "VIEWS: Duplicate with Detailing",
        "VIEWS: Duplicate as Dependent",
    ]

selected_option = "SHEETS: Duplicate with Views" if PRESET else forms.CommandSwitchWindow.show(
    options,
    message="Select duplication mode  "
            "({} sheet(s), {} view(s) selected):".format(len(sheets), len(views))
)

if not selected_option:
    script.exit()

# ── View duplication option (used for views AND for "with Views" sheets) ──────

# When sheets are duplicated with views, ask separately how views should be duped
view_dup_option = DB.ViewDuplicateOption.WithDetailing  # sensible default

if selected_option == "SHEETS: Duplicate with Views" and sheets:
    view_mode = "Duplicate with Detailing" if PRESET else forms.CommandSwitchWindow.show(
        ["Duplicate", "Duplicate with Detailing"],
        message="How should the views placed on the sheet(s) be duplicated?"
    )
    if not view_mode:
        script.exit()
    view_dup_option = (
        DB.ViewDuplicateOption.WithDetailing
        if view_mode == "Duplicate with Detailing"
        else DB.ViewDuplicateOption.Duplicate
    )

# ── How many copies ───────────────────────────────────────────────────────────

copies_text = PRESET["copies"] if PRESET else forms.ask_for_string(
    default="1",
    prompt="How many copies of each selected sheet / view?",
    title="Duplicate Sheets / Views"
)
if copies_text is None:
    script.exit()
try:
    copies = int(copies_text.strip())
except ValueError:
    copies = 0
if copies < 1:
    forms.alert("Enter a whole number of 1 or more.",
                title="Duplicate Sheets / Views", exitscript=True)

# ── Sheet numbering and detail numbers (sheet modes only) ─────────────────────

NUM_REVIT = "Revit default  (45.03.01 -> 45.03.01 (1))"
NUM_CONTINUE = "Continue the sequence  (45.03.01 -> 45.03.02)"
NUM_PATTERN = "Numbering pattern  (e.g. 45.03.xx or S20xx)"

number_mode = NUM_REVIT
number_pattern = None
keep_detail_numbers = False

if sheets and selected_option.startswith("SHEETS"):
    number_mode = NUM_CONTINUE if PRESET else forms.CommandSwitchWindow.show(
        [NUM_REVIT, NUM_CONTINUE, NUM_PATTERN],
        message="How should the new sheets be numbered?"
    )
    if not number_mode:
        script.exit()
    if number_mode == NUM_PATTERN:
        number_pattern = forms.ask_for_string(
            default="45.03.xx",
            prompt="Pattern: fixed text then x's for the digits "
                   "(45.03.xx, S20xx). Each new sheet continues from the "
                   "number of the sheet it is copied from.",
            title="Sheet numbering pattern"
        )
        if not number_pattern or not re.match(r"^(.*?)([xX#]+)$",
                                              number_pattern.strip()):
            forms.alert("The pattern must end in x's, for example 45.03.xx.",
                        title="Sheet numbering pattern", exitscript=True)

if selected_option == "SHEETS: Duplicate with Views" and sheets:
    dn_mode = "Keep the original detail numbers" if PRESET else forms.CommandSwitchWindow.show(
        ["Keep the original detail numbers",
         "Reset detail numbers (Revit numbers them on the new sheet)"],
        message="Detail numbers of the views placed on the new sheets:"
    )
    if not dn_mode:
        script.exit()
    keep_detail_numbers = dn_mode.startswith("Keep")


def new_sheet_number(source):
    """The number for a copy of source, by the chosen numbering mode."""
    if number_mode == NUM_CONTINUE:
        found = sequence_sheet_number(source.SheetNumber)
    elif number_mode == NUM_PATTERN:
        found = sequence_sheet_number(source.SheetNumber, number_pattern)
    else:
        found = None
    return found or unique_sheet_number(source.SheetNumber)


# One job per copy, kept in selection order so a copy of sheet A is
# numbered before the copies of sheet B.
sheet_jobs = [s for s in sheets for _i in range(copies)]
view_jobs = [v for v in views for _i in range(copies)]

# ── Execute ───────────────────────────────────────────────────────────────────

# Timed from here so the dialogs above are not counted against the run.
start_time = time.time()

new_element_ids = []
sheet_count = 0
view_count = 0

print_header("SHEET DUPLICATION")
print_separator()
flush_output()   # opens the output window before any work starts

try:
    with revit.Transaction("Duplicate Sheets / Views"):

        # ── SHEET modes ───────────────────────────────────────────────────────
        if selected_option in (
            "SHEETS: Duplicate Empty",
            "SHEETS: Duplicate with Sheet Detailing",
            "SHEETS: Duplicate with Views",
        ):
            for sheet in sheet_jobs:
                try:
                    src_tbs = get_title_blocks(sheet)
                    tb_type_id = (src_tbs[0].GetTypeId() if src_tbs
                                  else DB.ElementId.InvalidElementId)

                    # Create the new blank sheet
                    if tb_type_id != DB.ElementId.InvalidElementId:
                        new_sheet = DB.ViewSheet.Create(doc, tb_type_id)
                    else:
                        new_sheet = DB.ViewSheet.CreatePlaceholder(doc)

                    # New sheet number only -- all other params copied from original
                    new_sheet.SheetNumber = new_sheet_number(sheet)
                    print_dim("Source sheet: {} | {}".format(
                        sheet.SheetNumber, sheet.Name))

                    # Move the new sheet's title block to match the source sheet's
                    # title block position -- Revit always creates it at 0,0 which
                    # causes the header to appear offset against annotations
                    dst_tbs = get_title_blocks(new_sheet)
                    tb_offset = DB.XYZ.Zero
                    if src_tbs and dst_tbs:
                        src_tb_org = src_tbs[0].GetTransform().Origin
                        dst_tb_org = dst_tbs[0].GetTransform().Origin
                        tb_offset = DB.XYZ(
                            src_tb_org.X - dst_tb_org.X,
                            src_tb_org.Y - dst_tb_org.Y,
                            0)
                        if abs(tb_offset.X) > 1e-6 or abs(tb_offset.Y) > 1e-6:
                            DB.ElementTransformUtils.MoveElement(
                                doc, dst_tbs[0].Id, tb_offset)
                            print_debug("Moved dst title block by: {:.4f}, {:.4f}".format(
                                tb_offset.X, tb_offset.Y))

                    # Copy all writable parameters from original sheet
                    # Skip SheetNumber (already set) and built-in read-only identity params
                    copy_writable_parameters(sheet, new_sheet, skip_names={
                        "Sheet Number", "Sheet Name", "Sheet Issue Date",
                        "Revisions on Sheet",
                    })

                    # Sheet Name (view.Name) must be set explicitly as it maps to .Name
                    new_sheet.Name = sheet.Name

                    new_element_ids.append(new_sheet.Id)
                    sheet_count += 1

                    if selected_option == "SHEETS: Duplicate Empty":
                        print_success("Sheet {} duplicated (empty) -> {}".format(
                            sheet.SheetNumber, new_sheet.SheetNumber))
                        continue

                    viewport_ids = set(sheet.GetAllViewports())

                    # Duplicate with Sheet Detailing: copy all annotation elements
                    if selected_option in (
                        "SHEETS: Duplicate with Sheet Detailing",
                        "SHEETS: Duplicate with Views",
                    ):
                        # Copy all non-viewport, non-title-block elements from the
                        # original sheet. Title blocks are created by
                        # ViewSheet.Create and must not be copied as annotations.
                        exclude_ids = viewport_ids | {tb.Id for tb in src_tbs}

                        annotation_ids = List[DB.ElementId]([
                            eid for eid in
                            DB.FilteredElementCollector(doc, sheet.Id)
                              .WhereElementIsNotElementType()
                              .ToElementIds()
                            if eid not in exclude_ids
                        ])

                        if annotation_ids.Count > 0:
                            try:
                                if DEBUG:
                                    for tb in src_tbs:
                                        org = tb.GetTransform().Origin
                                        print_debug("SRC tb id:{} origin: {:.4f}, {:.4f}, {:.4f}".format(
                                            eid_int(tb.Id), org.X, org.Y, org.Z))
                                    for tb in dst_tbs:
                                        org = tb.GetTransform().Origin
                                        print_debug("DST tb id:{} origin: {:.4f}, {:.4f}, {:.4f}".format(
                                            eid_int(tb.Id), org.X, org.Y, org.Z))
                                    print_debug("annotation offset: {:.4f}, {:.4f}".format(
                                        tb_offset.X, tb_offset.Y))
                                    print_debug("annotation_ids count: {}".format(
                                        annotation_ids.Count))

                                # NOTE: no translation. The new title block was
                                # already moved onto the source's origin above, so
                                # both sheets share one coordinate frame. Moving the
                                # annotations by tb_offset as well shifted them by
                                # the offset twice (red sheet note landed ~85 mm up).
                                DB.ElementTransformUtils.CopyElements(
                                    sheet,
                                    annotation_ids,
                                    new_sheet,
                                    DB.Transform.Identity,
                                    DB.CopyPasteOptions()
                                )
                            except Exception as ce:
                                print_warning(
                                    "Could not copy annotations from sheet {}: {}".format(
                                        sheet.SheetNumber, str(ce)))
                                logger.warning(
                                    "Could not copy annotations from sheet {}: {}".format(
                                        sheet.SheetNumber, str(ce)))

                    if selected_option == "SHEETS: Duplicate with Sheet Detailing":
                        print_success("Sheet {} duplicated with detailing -> {}".format(
                            sheet.SheetNumber, new_sheet.SheetNumber))
                        continue

                    # Duplicate with Views: duplicate each placed view and re-place it
                    if selected_option == "SHEETS: Duplicate with Views":
                        placed_count = 0
                        for vp_id in viewport_ids:
                            vp = doc.GetElement(vp_id)
                            if vp is None:
                                continue
                            original_view = doc.GetElement(vp.ViewId)
                            if original_view is None:
                                continue

                            new_view_id = duplicate_view(original_view, view_dup_option)
                            if new_view_id is None:
                                print_warning(
                                    "Could not duplicate view '{}' - skipping".format(
                                        original_view.Name))
                                continue

                            new_view = doc.GetElement(new_view_id)

                            # Copy writable parameters from original view
                            copy_writable_parameters(original_view, new_view, skip_names={
                                "View Name", "Title on Sheet", "Sheet Number",
                                "Sheet Name", "Detail Number", "Dependency",
                            })

                            # Keep view name matching original (unique suffix only if needed)
                            try:
                                new_view.Name = unique_view_name(original_view.Name)
                            except Exception:
                                pass

                            match_view_graphics(original_view, new_view)

                            # Place on the new sheet, then line it up with the
                            # original: see match_viewport for why not by box.
                            try:
                                centre = vp.GetBoxCenter()
                                new_vp = DB.Viewport.Create(
                                    doc, new_sheet.Id, new_view_id,
                                    DB.XYZ(centre.X, centre.Y, 0))

                                # Copy viewport type (title style) from original
                                try:
                                    new_vp.ChangeTypeId(vp.GetTypeId())
                                except Exception:
                                    pass

                                if keep_detail_numbers:
                                    try:
                                        src_dn = vp.get_Parameter(
                                            DB.BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                                        new_vp.get_Parameter(
                                            DB.BuiltInParameter.VIEWPORT_DETAIL_NUMBER
                                        ).Set(src_dn.AsString())
                                    except Exception:
                                        pass    # leave Revit's own number

                                how = match_viewport(vp, new_vp)
                                print_debug("viewport {} aligned by {}".format(
                                    eid_int(new_vp.Id), how))

                                placed_count += 1
                                new_element_ids.append(new_view_id)
                                view_count += 1
                            except Exception as pe:
                                print_warning(
                                    "Could not place view '{}' on new sheet: {}".format(
                                        new_view.Name, str(pe)))
                                logger.warning(
                                    "Could not place view '{}' on new sheet: {}".format(
                                        new_view.Name, str(pe)))

                        print_success("Sheet {} duplicated with {} view(s) -> {}".format(
                            sheet.SheetNumber, placed_count, new_sheet.SheetNumber))

                except Exception as e:
                    print_error("Error duplicating sheet {}: {}".format(
                        sheet.SheetNumber, str(e)))
                    logger.error(
                        "Error duplicating sheet {}: {}".format(
                            sheet.SheetNumber, str(e)))
                finally:
                    # One write per sheet, so the log fills in as the run goes.
                    flush_output()

        # ── VIEW modes ────────────────────────────────────────────────────────
        elif selected_option in (
            "VIEWS: Duplicate",
            "VIEWS: Duplicate with Detailing",
            "VIEWS: Duplicate as Dependent",
        ):
            dup_map = {
                "VIEWS: Duplicate":               DB.ViewDuplicateOption.Duplicate,
                "VIEWS: Duplicate with Detailing": DB.ViewDuplicateOption.WithDetailing,
                "VIEWS: Duplicate as Dependent":   DB.ViewDuplicateOption.AsDependent,
            }
            dup_op = dup_map[selected_option]

            for view in view_jobs:
                try:
                    new_id = duplicate_view(view, dup_op)
                    if new_id is None:
                        print_warning(
                            "View '{}' could not be duplicated with option '{}'".format(
                                view.Name, selected_option))
                        continue

                    new_view = doc.GetElement(new_id)
                    # Copy writable parameters from original
                    copy_writable_parameters(view, new_view, skip_names={
                        "View Name", "Title on Sheet", "Sheet Number",
                        "Sheet Name", "Detail Number", "Dependency",
                    })
                    try:
                        new_view.Name = unique_view_name(view.Name)
                    except Exception:
                        pass

                    match_view_graphics(view, new_view)

                    new_element_ids.append(new_id)
                    view_count += 1
                    print_success("View {} duplicated -> {}".format(
                        view.Name, new_view.Name))

                except Exception as e:
                    print_error("Error duplicating view '{}': {}".format(
                        view.Name, str(e)))
                    logger.error(
                        "Error duplicating view '{}': {}".format(view.Name, str(e)))
                finally:
                    flush_output()
finally:
    flush_output()

# ── Select new elements ───────────────────────────────────────────────────────

# The transaction has committed by now, so the document is already regenerated
# -- an extra transaction just to call Regenerate cost more than it bought.
elapsed = time.time() - start_time

if new_element_ids:
    try:
        revit.get_selection().set_to(new_element_ids)
    except Exception:
        pass

    print_separator()
    print_header("DUPLICATION COMPLETE")
    print_success("Sheets created:    {}".format(sheet_count))
    print_success("Views created:     {}".format(view_count))
    print_success("Total elements:    {}".format(len(new_element_ids)))
    print_success("Processing time:   {:.2f} seconds".format(elapsed))
    print_separator()
    flush_output()
else:
    forms.alert("No elements were duplicated.", title="Duplicate Sheets / Views")
