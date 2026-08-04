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

from pyrevit import revit, DB, forms, script
from System.Collections.Generic import List

doc   = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()

# ── Helpers ───────────────────────────────────────────────────────────────────

def eid_int(element_id):
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue


def unique_sheet_number(base_number):
    """Return base_number + suffix matching Revit's own pattern: (1), (2)..."""
    existing = {
        s.SheetNumber
        for s in DB.FilteredElementCollector(doc)
                   .OfClass(DB.ViewSheet)
                   .ToElements()
    }
    if base_number not in existing:
        return base_number
    counter = 1
    while True:
        candidate = "{} ({})".format(base_number, counter)
        if candidate not in existing:
            return candidate
        counter += 1


def unique_view_name(base_name):
    """Return base_name + suffix that does not clash with existing view names."""
    existing = {
        v.Name
        for v in DB.FilteredElementCollector(doc)
                   .OfClass(DB.View)
                   .WhereElementIsNotElementType()
                   .ToElements()
    }
    if base_name not in existing:
        return base_name
    counter = 2
    while True:
        candidate = "{} ({})".format(base_name, counter)
        if candidate not in existing:
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


def get_title_block_id(sheet):
    """Return the ElementId of the first title block placed on the sheet."""
    tbs = (
        DB.FilteredElementCollector(doc, sheet.Id)
        .OfCategory(DB.BuiltInCategory.OST_TitleBlocks)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    for tb in tbs:
        return tb.GetTypeId()
    return DB.ElementId.InvalidElementId


def copy_writable_parameters(source, target, skip_names=None):
    """Copy all writable parameter values from source element to target."""
    skip_names = skip_names or set()
    for p in source.Parameters:
        try:
            if p.IsReadOnly:
                continue
            if not p.Definition:
                continue
            name = p.Definition.Name
            if name in skip_names:
                continue
            tp = target.LookupParameter(name)
            if tp is None or tp.IsReadOnly:
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

new_element_ids = []
output = script.get_output()

with revit.Transaction("Duplicate Sheets / Views"):

    # ── SHEET modes ───────────────────────────────────────────────────────────
    if selected_option in (
        "SHEETS: Duplicate Empty",
        "SHEETS: Duplicate with Sheet Detailing",
        "SHEETS: Duplicate with Views",
    ):
        for sheet in sheets:
            try:
                tb_type_id = get_title_block_id(sheet)

                # Create the new blank sheet
                if tb_type_id != DB.ElementId.InvalidElementId:
                    new_sheet = DB.ViewSheet.Create(doc, tb_type_id)
                else:
                    new_sheet = DB.ViewSheet.CreatePlaceholder(doc)

                # New sheet number only -- all other params copied from original
                new_sheet.SheetNumber = unique_sheet_number(sheet.SheetNumber)
                print_dim = lambda msg: output.print_md("-> " + msg)
                print_dim("=== SHEET DUPLICATION LOG ===")
                print_dim("Source sheet: {} | {}".format(sheet.SheetNumber, sheet.Name))

                # Move the new sheet's title block to match the source sheet's
                # title block position -- Revit always creates it at 0,0 which
                # causes the header to appear offset against annotations
                src_tb_elems = list(
                    DB.FilteredElementCollector(doc, sheet.Id)
                    .OfCategory(DB.BuiltInCategory.OST_TitleBlocks)
                    .WhereElementIsNotElementType()
                    .ToElements()
                )
                dst_tb_elems = list(
                    DB.FilteredElementCollector(doc, new_sheet.Id)
                    .OfCategory(DB.BuiltInCategory.OST_TitleBlocks)
                    .WhereElementIsNotElementType()
                    .ToElements()
                )
                if src_tb_elems and dst_tb_elems:
                    src_tb_org = src_tb_elems[0].GetTransform().Origin
                    dst_tb_org = dst_tb_elems[0].GetTransform().Origin
                    tb_translation = DB.XYZ(
                        src_tb_org.X - dst_tb_org.X,
                        src_tb_org.Y - dst_tb_org.Y,
                        0)
                    if abs(tb_translation.X) > 1e-6 or abs(tb_translation.Y) > 1e-6:
                        DB.ElementTransformUtils.MoveElement(
                            doc, dst_tb_elems[0].Id, tb_translation)
                        print_dim("Moved dst title block by: {:.4f}, {:.4f}".format(
                            tb_translation.X, tb_translation.Y))

                # Copy all writable parameters from original sheet
                # Skip SheetNumber (already set) and built-in read-only identity params
                copy_writable_parameters(sheet, new_sheet, skip_names={
                    "Sheet Number", "Sheet Name", "Sheet Issue Date",
                    "Revisions on Sheet",
                })

                # Sheet Name (view.Name) must be set explicitly as it maps to .Name
                new_sheet.Name = sheet.Name

                new_element_ids.append(new_sheet.Id)

                if selected_option == "SHEETS: Duplicate Empty":
                    output.print_md(
                        "Sheet **{}** duplicated (empty) -> **{}**".format(
                            sheet.SheetNumber, new_sheet.SheetNumber))
                    continue

                # Duplicate with Sheet Detailing: copy all annotation elements
                if selected_option in (
                    "SHEETS: Duplicate with Sheet Detailing",
                    "SHEETS: Duplicate with Views",
                ):
                    # Copy all non-viewport, non-title-block elements from the original sheet
                    all_ids_on_sheet = list(
                        DB.FilteredElementCollector(doc, sheet.Id)
                        .WhereElementIsNotElementType()
                        .ToElementIds()
                    )
                    viewport_ids = set(sheet.GetAllViewports())

                    # Collect title block ids to exclude -- they are created by
                    # ViewSheet.Create and must not be copied as annotations
                    title_block_ids = set(
                        DB.FilteredElementCollector(doc, sheet.Id)
                        .OfCategory(DB.BuiltInCategory.OST_TitleBlocks)
                        .WhereElementIsNotElementType()
                        .ToElementIds()
                    )

                    exclude_ids = viewport_ids | title_block_ids

                    annotation_ids = List[DB.ElementId]([
                        eid for eid in all_ids_on_sheet
                        if eid not in exclude_ids
                    ])

                    if annotation_ids.Count > 0:
                        try:
                            # Collect title block origins on both sheets
                            src_tbs = list(
                                DB.FilteredElementCollector(doc, sheet.Id)
                                .OfCategory(DB.BuiltInCategory.OST_TitleBlocks)
                                .WhereElementIsNotElementType()
                                .ToElements()
                            )
                            dst_tbs = list(
                                DB.FilteredElementCollector(doc, new_sheet.Id)
                                .OfCategory(DB.BuiltInCategory.OST_TitleBlocks)
                                .WhereElementIsNotElementType()
                                .ToElements()
                            )

                            # Log origins for debugging
                            for tb in src_tbs:
                                try:
                                    org = tb.GetTransform().Origin
                                    output.print_md("-> SRC tb id:{} origin: {:.4f}, {:.4f}, {:.4f}".format(
                                        eid_int(tb.Id), org.X, org.Y, org.Z))
                                except Exception as le:
                                    output.print_md("-> SRC tb log error: {}".format(le))
                            for tb in dst_tbs:
                                try:
                                    org = tb.GetTransform().Origin
                                    output.print_md("-> DST tb id:{} origin: {:.4f}, {:.4f}, {:.4f}".format(
                                        eid_int(tb.Id), org.X, org.Y, org.Z))
                                except Exception as le:
                                    output.print_md("-> DST tb log error: {}".format(le))

                            # Offset = src_origin - dst_origin so annotations
                            # land correctly relative to the new sheet title block
                            copy_transform = DB.Transform.Identity
                            if src_tbs and dst_tbs:
                                src_org = src_tbs[0].GetTransform().Origin
                                dst_org = dst_tbs[0].GetTransform().Origin
                                offset = DB.XYZ(
                                    src_org.X - dst_org.X,
                                    src_org.Y - dst_org.Y,
                                    0)
                                output.print_md("-> annotation offset: {:.4f}, {:.4f}".format(
                                    offset.X, offset.Y))
                                copy_transform = DB.Transform.CreateTranslation(offset)

                            output.print_md("-> annotation_ids count: {}".format(annotation_ids.Count))

                            DB.ElementTransformUtils.CopyElements(
                                sheet,
                                annotation_ids,
                                new_sheet,
                                copy_transform,
                                DB.CopyPasteOptions()
                            )
                        except Exception as ce:
                            logger.warning(
                                "Could not copy annotations from sheet {}: {}".format(
                                    sheet.SheetNumber, str(ce)))

                if selected_option == "SHEETS: Duplicate with Sheet Detailing":
                    output.print_md(
                        "Sheet **{}** duplicated with detailing -> **{}**".format(
                            sheet.SheetNumber, new_sheet.SheetNumber))
                    continue

                # Duplicate with Views: duplicate each placed view and re-place it
                if selected_option == "SHEETS: Duplicate with Views":
                    placed_count = 0
                    for vp_id in sheet.GetAllViewports():
                        vp = doc.GetElement(vp_id)
                        if vp is None:
                            continue
                        original_view = doc.GetElement(vp.ViewId)
                        if original_view is None:
                            continue

                        new_view_id = duplicate_view(original_view, view_dup_option)
                        if new_view_id is None:
                            logger.warning(
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

                        # Place on new sheet then move to exact same position as original
                        try:
                            centre = vp.GetBoxCenter()
                            output.print_md("-> VP original GetBoxCenter: {:.1f}, {:.1f}, {:.1f}".format(
                                centre.X, centre.Y, centre.Z))

                            # Create at origin
                            new_vp = DB.Viewport.Create(
                                doc, new_sheet.Id, new_view_id,
                                DB.XYZ(0, 0, 0))

                            output.print_md("-> new_vp created at 0,0 | id: {}".format(eid_int(new_vp.Id)))

                            # Copy viewport type (title style) from original viewport
                            try:
                                new_vp.ChangeTypeId(vp.GetTypeId())
                            except Exception:
                                pass

                            # Move viewport to match original position.
                            translation = DB.XYZ(centre.X, centre.Y, 0)
                            DB.ElementTransformUtils.MoveElement(
                                doc, new_vp.Id, translation)
                            output.print_md("-> MoveElement translation applied: {:.1f}, {:.1f}".format(
                                centre.X, centre.Y))
                            after_centre = new_vp.GetBoxCenter()
                            output.print_md("-> new_vp GetBoxCenter after move: {:.1f}, {:.1f}".format(
                                after_centre.X, after_centre.Y))

                            placed_count += 1
                            new_element_ids.append(new_view_id)
                        except Exception as pe:
                            logger.warning(
                                "Could not place view '{}' on new sheet: {}".format(
                                    new_view.Name, str(pe)))

                    output.print_md(
                        "Sheet **{}** duplicated with {} view(s) -> **{}**".format(
                            sheet.SheetNumber, placed_count, new_sheet.SheetNumber))

            except Exception as e:
                logger.error(
                    "Error duplicating sheet {}: {}".format(
                        sheet.SheetNumber, str(e)))

    # ── VIEW modes ────────────────────────────────────────────────────────────
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
                    logger.warning(
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
                output.print_md(
                    "View **{}** duplicated -> **{}**".format(
                        view.Name, new_view.Name))

            except Exception as e:
                logger.error(
                    "Error duplicating view '{}': {}".format(view.Name, str(e)))

# ── Select new elements ───────────────────────────────────────────────────────

if new_element_ids:
    # Regenerate must be called inside a transaction
    with revit.Transaction("Duplicate Sheets / Views - regenerate"):
        doc.Regenerate()
    try:
        revit.get_selection().set_to(new_element_ids)
    except Exception:
        pass
    output.print_md(
        "---\n**Done.** {} element(s) created.".format(len(new_element_ids)))
else:
    forms.alert("No elements were duplicated.", title="Duplicate Sheets / Views")
