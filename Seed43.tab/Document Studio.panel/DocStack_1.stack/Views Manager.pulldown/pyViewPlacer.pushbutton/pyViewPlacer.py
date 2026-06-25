# -*- coding: utf-8 -*-
#pyViewPlacer
import os
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
from System.Windows import Visibility, Point
from System.Windows.Controls import ListBoxItem, ScrollViewer
from System.Windows.Markup import XamlReader
from System.Windows.Media import SolidColorBrush, Color, VisualTreeHelper

from pyrevit import revit, DB, UI, forms, script

# ── VARIABLES ────────────────────────────────────────────────────────
HELP_URL  = "https://example.com/help"
ABOUT_URL = "https://example.com/about"
XAML_PATH = os.path.join(os.path.dirname(__file__), "pyViewPlacer.xaml")


# ── HELPERS ───────────────────────────────────────────────────────────
def get_element_id_value(element_id):
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue

def create_element_id(value):
    try:
        return DB.ElementId(int(value))
    except Exception:
        return DB.ElementId(long(value))

def load_xaml(path):
    with open(path, "r") as f:
        return XamlReader.Parse(f.read())


# ── SELECTION FILTER ──────────────────────────────────────────────────
class ViewSelectionFilter(UI.Selection.ISelectionFilter):
    def AllowElement(self, element):
        return True
    def AllowReference(self, reference, position):
        return False


# ── VIEW MANAGER ──────────────────────────────────────────────────────
class ViewManager(object):
    def __init__(self, doc, uidoc):
        self.doc   = doc
        self.uidoc = uidoc

    def get_all_non_template_views(self):
        return [
            v for v in
            revit.query.get_elements_by_class(DB.View, doc=self.doc)
            if not v.IsTemplate
        ]

    def get_view_from_element(self, element):
        if hasattr(element, "ViewId"):
            view_id = element.ViewId
            if view_id and view_id != DB.ElementId.InvalidElementId:
                return self.doc.GetElement(view_id)
        if isinstance(element, DB.ElevationMarker):
            view_ids = []
            for i in range(4):
                vid = element.GetViewId(i)
                if vid != DB.ElementId.InvalidElementId:
                    view_ids.append(vid)
            if len(view_ids) == 1:
                return self.doc.GetElement(view_ids[0])
            elif len(view_ids) > 1:
                return [self.doc.GetElement(vid) for vid in view_ids]
        view_name = self._get_view_name_from_element(element)
        if view_name:
            for view in self.get_all_non_template_views():
                if view.Name == view_name:
                    return view
        return None

    def _get_view_name_from_element(self, element):
        try:
            param = element.LookupParameter("View Name")
            if param and param.HasValue:
                return param.AsString()
        except Exception:
            pass
        try:
            type_id = element.GetTypeId()
            if type_id != DB.ElementId.InvalidElementId:
                elem_type = self.doc.GetElement(type_id)
                if elem_type:
                    param = elem_type.get_Parameter(
                        DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                    if param:
                        return param.AsString()
        except Exception:
            pass
        return None

    def select_views_from_active_view(self):
        selected_views = []
        used_view_ids  = set()
        while True:
            try:
                picked_ref = self.uidoc.Selection.PickObject(
                    UI.Selection.ObjectType.Element,
                    ViewSelectionFilter(),
                    "Select view reference/callout/section (ESC to finish)")
                if not picked_ref:
                    break
                element     = self.doc.GetElement(picked_ref)
                view_result = self.get_view_from_element(element)
                if isinstance(view_result, list):
                    for view in view_result:
                        if view and view.Id not in used_view_ids:
                            selected_views.append(view)
                            used_view_ids.add(view.Id)
                elif view_result and view_result.Id not in used_view_ids:
                    selected_views.append(view_result)
                    used_view_ids.add(view_result.Id)
            except Exception as e:
                if "cancelled" in str(e).lower() or "aborted" in str(e).lower():
                    break
                break
        return selected_views


# ── DIALOG ────────────────────────────────────────────────────────────
class ViewPlacerDialog(object):
    def __init__(self):
        self.result         = None
        self._panel_mode    = None
        self._all_items     = []
        self._active_list   = None
        self._active_search = None
        # ── Load XAML, FindName while name scope is intact ────────────
        self._window = load_xaml(XAML_PATH)
        # ── Named elements ────────────────────────────────────────────
        self.header_subtitle  = self._window.FindName("header_subtitle")
        self.panel_template   = self._window.FindName("panel_template")
        self.panel_sheet      = self._window.FindName("panel_sheet")
        self.panel_viewport   = self._window.FindName("panel_viewport")
        self.template_search  = self._window.FindName("template_search")
        self.template_list    = self._window.FindName("template_list")
        self.template_ok_btn  = self._window.FindName("template_ok_btn")
        self.sheet_search     = self._window.FindName("sheet_search")
        self.sheet_list       = self._window.FindName("sheet_list")
        self.sheet_ok_btn     = self._window.FindName("sheet_ok_btn")
        self.viewport_search  = self._window.FindName("viewport_search")
        self.viewport_list    = self._window.FindName("viewport_list")
        self.viewport_ok_btn  = self._window.FindName("viewport_ok_btn")
        self.help_btn         = self._window.FindName("help_btn")
        self.about_btn        = self._window.FindName("about_btn")
        # ── Event handlers ────────────────────────────────────────────
        self.help_btn.Click              += self._on_help
        self.about_btn.Click             += self._on_about
        self.template_search.TextChanged += self._on_search_changed
        self.sheet_search.TextChanged    += self._on_search_changed
        self.viewport_search.TextChanged += self._on_search_changed
        self.template_list.MouseDoubleClick  += self._on_double_click
        self.sheet_list.MouseDoubleClick     += self._on_double_click
        self.viewport_list.MouseDoubleClick  += self._on_double_click
        self.template_ok_btn.Click  += self._on_ok
        self.sheet_ok_btn.Click     += self._on_ok
        self.viewport_ok_btn.Click  += self._on_ok

    # ── Panel switching ───────────────────────────────────────────────
    def _show_panel(self, mode, subtitle):
        self.panel_template.Visibility  = Visibility.Collapsed
        self.panel_sheet.Visibility     = Visibility.Collapsed
        self.panel_viewport.Visibility  = Visibility.Collapsed
        self._panel_mode                = mode
        self.header_subtitle.Text       = " | {}".format(subtitle)
        self.header_subtitle.Visibility = Visibility.Visible
        if mode == "template":
            self.panel_template.Visibility = Visibility.Visible
            self._active_list   = self.template_list
            self._active_search = self.template_search
        elif mode == "sheet":
            self.panel_sheet.Visibility = Visibility.Visible
            self._active_list   = self.sheet_list
            self._active_search = self.sheet_search
        elif mode == "viewport":
            self.panel_viewport.Visibility = Visibility.Visible
            self._active_list   = self.viewport_list
            self._active_search = self.viewport_search

    # ── Populate helpers ──────────────────────────────────────────────
    def _make_item(self, content, tag, highlight=False):
        item         = ListBoxItem()
        item.Content = content
        item.Tag     = tag
        if highlight:
            item.Background = SolidColorBrush(
                Color.FromArgb(40, 32, 138, 60))
        return item

    def _scroll_to_index(self, list_box, index):
        if index < 0:
            return
        list_box.SelectedIndex = index
        list_box.UpdateLayout()
        list_box.ScrollIntoView(list_box.Items[index])
        def center(*args):
            try:
                container = (list_box.ItemContainerGenerator
                             .ContainerFromIndex(index))
                if not container:
                    return
                transform     = container.TransformToAncestor(list_box)
                pos           = transform.Transform(Point(0, 0))
                target_offset = (pos.Y - (list_box.ActualHeight / 2)
                                 + (container.ActualHeight / 2))
                sv = self._find_scroll_viewer(list_box)
                if sv:
                    target_offset = max(0, min(target_offset,
                                               sv.ScrollableHeight))
                    sv.ScrollToVerticalOffset(target_offset)
            except Exception:
                pass
            list_box.LayoutUpdated -= center
        list_box.LayoutUpdated += center

    def _find_scroll_viewer(self, dep_obj):
        if dep_obj is None:
            return None
        if isinstance(dep_obj, ScrollViewer):
            return dep_obj
        count = VisualTreeHelper.GetChildrenCount(dep_obj)
        for i in range(count):
            child  = VisualTreeHelper.GetChild(dep_obj, i)
            result = self._find_scroll_viewer(child)
            if result is not None:
                return result
        return None

    # ── Public populate methods ───────────────────────────────────────
    def populate_templates(self, templates, last_template):
        self._show_panel("template", "View Template")
        self._all_items = []
        list_box        = self.template_list
        list_box.Items.Clear()
        selected_index  = 0
        none_item       = self._make_item("<None - No Template>", None)
        list_box.Items.Add(none_item)
        self._all_items.append(none_item)
        if last_template:
            last_item = self._make_item(
                "{} (Last Used)".format(last_template.Name), last_template)
            list_box.Items.Add(last_item)
            self._all_items.append(last_item)
            selected_index = 1
        for tmpl in sorted(templates, key=lambda t: t.Name):
            if not last_template or tmpl.Id != last_template.Id:
                item = self._make_item(tmpl.Name, tmpl)
                list_box.Items.Add(item)
                self._all_items.append(item)
        self._scroll_to_index(list_box, selected_index)

    def populate_sheets(self, sheets, last_sheet):
        self._show_panel("sheet", "Sheet Selection")
        self._all_items = []
        list_box        = self.sheet_list
        list_box.Items.Clear()
        selected_index  = -1
        for i, sheet in enumerate(sorted(sheets, key=lambda s: s.SheetNumber)):
            is_last = last_sheet and sheet.Id == last_sheet.Id
            label   = ("{} - {} (Last Used)".format(sheet.SheetNumber, sheet.Name)
                       if is_last else
                       "{} - {}".format(sheet.SheetNumber, sheet.Name))
            item = self._make_item(label, sheet, highlight=is_last)
            list_box.Items.Add(item)
            self._all_items.append(item)
            if is_last:
                selected_index = i
        self._scroll_to_index(list_box, selected_index)

    def populate_viewports(self, viewport_types, last_viewport, get_name_func):
        self._show_panel("viewport", "Viewport Type")
        self._all_items = []
        list_box        = self.viewport_list
        list_box.Items.Clear()
        selected_index  = 0
        default_item    = self._make_item("<Default - No Change>", None)
        list_box.Items.Add(default_item)
        self._all_items.append(default_item)
        if last_viewport:
            last_item = self._make_item(
                "{} (Last Used)".format(get_name_func(last_viewport)),
                last_viewport)
            list_box.Items.Add(last_item)
            self._all_items.append(last_item)
            selected_index = 1
        vp_dict = {}
        for vp in viewport_types:
            name = get_name_func(vp)
            if name in vp_dict:
                name = "{} (ID: {})".format(name, get_element_id_value(vp.Id))
            vp_dict[name] = vp
        for name in sorted(vp_dict.keys()):
            vp = vp_dict[name]
            if not last_viewport or vp.Id != last_viewport.Id:
                item = self._make_item(name, vp)
                list_box.Items.Add(item)
                self._all_items.append(item)
        self._scroll_to_index(list_box, selected_index)

    # ── Event handlers ────────────────────────────────────────────────
    def _on_search_changed(self, sender, args):
        if not self._active_list or not self._active_search:
            return
        search_text = self._active_search.Text.lower()
        self._active_list.Items.Clear()
        for item in self._all_items:
            if not search_text or search_text in str(item.Content).lower():
                self._active_list.Items.Add(item)

    def _on_double_click(self, sender, args):
        self._on_ok(sender, args)

    def _on_ok(self, sender, args):
        if self._active_list and self._active_list.SelectedItem:
            self.result = self._active_list.SelectedItem.Tag
        else:
            self.result = None
        self._window.Close()

    def _on_help(self, sender, args):
        try:
            import webbrowser
            webbrowser.open(HELP_URL)
        except Exception:
            pass

    def _on_about(self, sender, args):
        try:
            import webbrowser
            webbrowser.open(ABOUT_URL)
        except Exception:
            pass


# ── MAIN ──────────────────────────────────────────────────────────────
def main():
    try:
        if not revit.doc:
            forms.alert("No active Revit document found")
            return
        active_view = revit.uidoc.ActiveView
        if isinstance(active_view, DB.ViewSheet):
            forms.alert(
                "Cannot run on a sheet view. "
                "Please open a floor plan, section, or elevation.")
            return

        view_manager     = ViewManager(revit.doc, revit.uidoc)
        config           = script.get_config()
        last_sheet_id    = getattr(config, "last_sheet_id", None)
        last_viewport_id = getattr(config, "last_viewport_id", None)
        last_template_id = getattr(config, "last_template_id", None)

        # ── Pick views ────────────────────────────────────────────────
        with forms.WarningBar(
                title="Pick view references, callouts, or sections. "
                      "ESCAPE to finish."):
            selected_views = view_manager.select_views_from_active_view()
        if not selected_views:
            return

        views_to_place  = []
        views_on_sheets = []
        for view in selected_views:
            sheet_param = view.get_Parameter(
                DB.BuiltInParameter.VIEWER_SHEET_NUMBER)
            if sheet_param and sheet_param.AsString() == "---":
                views_to_place.append(view)
            else:
                views_on_sheets.append(view)
        if not views_to_place:
            forms.alert("All selected views are already on sheets.")
            return

        # ── Template selection ────────────────────────────────────────
        templates = [
            v for v in
            revit.query.get_elements_by_class(DB.View, doc=revit.doc)
            if v.IsTemplate
        ]
        last_template = None
        if last_template_id:
            try:
                last_template = revit.doc.GetElement(
                    create_element_id(last_template_id))
                if not last_template or not last_template.IsTemplate:
                    last_template = None
            except (ValueError, TypeError, AttributeError):
                last_template = None

        dialog = ViewPlacerDialog()
        dialog.populate_templates(templates, last_template)
        dialog._window.ShowDialog()
        selected_template = dialog.result
        if dialog.result is None and (
                not dialog._active_list or
                dialog._active_list.SelectedIndex == -1):
            return
        config.last_template_id = (
            str(get_element_id_value(selected_template.Id))
            if selected_template else None)
        script.save_config()

        # ── Sheet selection ───────────────────────────────────────────
        all_sheets = revit.query.get_elements_by_class(
            DB.ViewSheet, doc=revit.doc)
        last_sheet = None
        if last_sheet_id:
            try:
                last_sheet = revit.doc.GetElement(
                    create_element_id(last_sheet_id))
                if not last_sheet or not isinstance(last_sheet, DB.ViewSheet):
                    last_sheet = None
            except (ValueError, TypeError, AttributeError):
                last_sheet = None

        dialog = ViewPlacerDialog()
        dialog.populate_sheets(all_sheets, last_sheet)
        dialog._window.ShowDialog()
        selected_sheet = dialog.result
        if not selected_sheet:
            return
        config.last_sheet_id = str(get_element_id_value(selected_sheet.Id))
        script.save_config()

        # ── Place views ───────────────────────────────────────────────
        created_viewports = []
        with revit.Transaction("Place Views on Sheet"):
            for i, view in enumerate(views_to_place):
                if selected_template:
                    try:
                        view.ViewTemplateId = selected_template.Id
                    except Exception:
                        pass
                position = DB.XYZ(1.0 + (i * 2.0), 1.5, 0.0)
                try:
                    viewport = DB.Viewport.Create(
                        revit.doc, selected_sheet.Id, view.Id, position)
                    created_viewports.append(viewport)
                except Exception:
                    pass

        # ── Viewport type selection ───────────────────────────────────
        if not created_viewports:
            return

        viewport_types = list(
            DB.FilteredElementCollector(revit.doc)
            .OfClass(DB.ElementType)
            .OfCategory(DB.BuiltInCategory.OST_Viewports))
        if not viewport_types:
            type_ids = set()
            for vp in revit.query.get_elements_by_class(
                    DB.Viewport, doc=revit.doc):
                type_ids.add(vp.GetTypeId())
            viewport_types = [
                revit.doc.GetElement(tid)
                for tid in type_ids
                if tid != DB.ElementId.InvalidElementId
            ]
        if not viewport_types:
            return

        last_viewport = None
        if last_viewport_id:
            try:
                last_viewport = revit.doc.GetElement(
                    create_element_id(last_viewport_id))
            except (ValueError, TypeError, AttributeError):
                pass

        def get_vp_name(vp_type):
            for param_id in (DB.BuiltInParameter.ALL_MODEL_TYPE_NAME,
                             DB.BuiltInParameter.SYMBOL_NAME_PARAM):
                try:
                    param = vp_type.get_Parameter(param_id)
                    if param and param.HasValue:
                        return param.AsString()
                except Exception:
                    pass
            try:
                return vp_type.Name
            except Exception:
                pass
            return "Unnamed Viewport Type"

        dialog = ViewPlacerDialog()
        dialog.populate_viewports(viewport_types, last_viewport, get_vp_name)
        dialog._window.ShowDialog()
        selected_viewport = dialog.result

        if selected_viewport:
            with revit.Transaction("Change Viewport Types"):
                for viewport in created_viewports:
                    try:
                        viewport.ChangeTypeId(selected_viewport.Id)
                    except Exception:
                        pass
            config.last_viewport_id = str(
                get_element_id_value(selected_viewport.Id))
        else:
            config.last_viewport_id = None
        script.save_config()

    except Exception as ex:
        forms.alert("Unexpected error occurred: {}".format(str(ex)))


if __name__ == "__main__":
    main()
