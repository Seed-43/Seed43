# -*- coding: utf-8 -*-
# duplicate_sheets.py
#
# Duplicates sheets and/or views selected from the project browser.
# Sheets:  Duplicate Empty  |  Duplicate with Sheet Detailing  |  Duplicate with Views
# Views:   Duplicate  |  Duplicate with Detailing  |  Duplicate as Dependent
#
# When duplicating a sheet "with Views", each placed view is duplicated using
# whichever view-duplication option you choose, and the duplicate is placed on
# the new sheet in the same viewport position.

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


# ── Collect selection ─────────────────────────────────────────────────────────

selected_elements = revit.get_selection().elements

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

selected_option = forms.CommandSwitchWindow.show(
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
    view_mode = forms.CommandSwitchWindow.show(
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
            for sheet in sheets:
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
                    new_sheet.SheetNumber = unique_sheet_number(sheet.SheetNumber)
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

                                # Offset = src_origin - dst_origin so annotations
                                # land correctly relative to the new title block
                                copy_transform = (
                                    DB.Transform.CreateTranslation(tb_offset)
                                    if src_tbs and dst_tbs
                                    else DB.Transform.Identity)

                                DB.ElementTransformUtils.CopyElements(
                                    sheet,
                                    annotation_ids,
                                    new_sheet,
                                    copy_transform,
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

                            # Place on new sheet then move to exact same position
                            try:
                                centre = vp.GetBoxCenter()
                                print_debug("VP original GetBoxCenter: {:.1f}, {:.1f}, {:.1f}".format(
                                    centre.X, centre.Y, centre.Z))

                                # Create at origin
                                new_vp = DB.Viewport.Create(
                                    doc, new_sheet.Id, new_view_id,
                                    DB.XYZ(0, 0, 0))

                                print_debug("new_vp created at 0,0 | id: {}".format(
                                    eid_int(new_vp.Id)))

                                # Copy viewport type (title style) from original
                                try:
                                    new_vp.ChangeTypeId(vp.GetTypeId())
                                except Exception:
                                    pass

                                # Move viewport to match original position.
                                DB.ElementTransformUtils.MoveElement(
                                    doc, new_vp.Id, DB.XYZ(centre.X, centre.Y, 0))
                                print_debug("MoveElement translation applied: {:.1f}, {:.1f}".format(
                                    centre.X, centre.Y))

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

            for view in views:
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
