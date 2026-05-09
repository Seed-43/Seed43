# -*- coding: utf-8 -*-
__title__     = "Section & Elevation Placer"
__author__    = "Seed-43"
__doc__       = """
𝐕𝐄𝐑𝐒𝐈𝐎𝐍 𝟐𝟔𝟎𝟓𝟎𝟏
_____________________________________________________________________
Description:

Select view references (callouts, sections, elevations) in the active view
and place their associated views on a chosen sheet.

Includes options for:
- Applying a view template
- Selecting a target sheet
- Choosing viewport type
- Reusing last-used settings for speed and consistency

Automatically detects whether views are already placed on sheets
and separates workflow accordingly.
_____________________________________________________________________
How-to:
-> Open a view containing callouts / sections / elevations
-> Run the script
-> Select view references (ESC to finish)
-> Choose view template (or None)
-> Select target sheet
-> Views are placed automatically
-> Choose viewport type (optional)
_____________________________________________________________________
Last update:
- Initial release
_____________________________________________________________________
"""

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('System.Windows')
from pyrevit import revit, DB, forms, script
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
import System.Windows
import System.Windows.Media
import System.Windows.Controls
import wpf
import os
import webbrowser

HELP_URL = "https://example.com/help"
ABOUT_URL = "https://example.com/about"
DIALOG_WIDTH = 450
DIALOG_HEIGHT = 660

def get_element_id_value(element_id):
    try:
        return element_id.IntegerValue
    except AttributeError:
        return element_id.Value

def create_element_id(value):
    try:
        return DB.ElementId(int(value))
    except:
        return DB.ElementId(long(value))

class ViewSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return True
    def AllowReference(self, reference, position):
        return False

class ViewManager:
    def __init__(self, doc, uidoc):
        self.doc = doc
        self.uidoc = uidoc

    def get_all_non_template_views(self):
        collector = DB.FilteredElementCollector(self.doc).OfClass(DB.View).WhereElementIsNotElementType()
        views = []
        for view in collector:
            if not view.IsTemplate:
                views.append(view)
        return views

    def get_view_from_element(self, element):
        if hasattr(element, 'ViewId'):
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
                views = [self.doc.GetElement(vid) for vid in view_ids]
                return views
        view_name = self._get_view_name_from_element(element)
        if view_name:
            all_views = self.get_all_non_template_views()
            for view in all_views:
                if view.Name == view_name:
                    return view
        return None

    def _get_view_name_from_element(self, element):
        try:
            view_name_param = element.LookupParameter("View Name")
            if view_name_param and view_name_param.HasValue:
                return view_name_param.AsString()
        except:
            pass
        try:
            type_id = element.GetTypeId()
            if type_id != DB.ElementId.InvalidElementId:
                elem_type = self.doc.GetElement(type_id)
                if elem_type:
                    type_name_param = elem_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                    if type_name_param:
                        return type_name_param.AsString()
        except:
            pass
        return None

    def select_views_from_active_view(self):
        selected_views = []
        used_view_ids = set()
        while True:
            try:
                picked_ref = self.uidoc.Selection.PickObject(ObjectType.Element, ViewSelectionFilter(), "Select view reference/callout/section (ESC to finish)")
                if not picked_ref:
                    break
                element = self.doc.GetElement(picked_ref)
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

class StyledDialog(System.Windows.Window):
    def __init__(self, title, subtitle):
        self.WindowStyle = System.Windows.WindowStyle.SingleBorderWindow
        self.ResizeMode = System.Windows.ResizeMode.CanResizeWithGrip
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.ShowInTaskbar = True
        self.Background = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(59, 69, 83))
        self.Width = DIALOG_WIDTH
        self.Height = DIALOG_HEIGHT
        self.Title = title
        dock_panel = System.Windows.Controls.DockPanel()
        dock_panel.Margin = System.Windows.Thickness(0, 0, 0, 1)
        self._create_header(dock_panel, subtitle)
        scroll = System.Windows.Controls.ScrollViewer()
        scroll.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
        scroll.Margin = System.Windows.Thickness(24, 24, 24, 24)
        self.content_panel = System.Windows.Controls.StackPanel()
        scroll.Content = self.content_panel
        dock_panel.Children.Add(scroll)
        self.Content = dock_panel
        self.result = None

    def _create_header(self, dock_panel, subtitle):
        header = System.Windows.Controls.Border()
        header.Background = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(35, 41, 51))
        header.Height = 70
        header.CornerRadius = System.Windows.CornerRadius(0, 0, 12, 12)
        System.Windows.Controls.DockPanel.SetDock(header, System.Windows.Controls.Dock.Top)
        shadow = System.Windows.Media.Effects.DropShadowEffect()
        shadow.Color = System.Windows.Media.Colors.Black
        shadow.Opacity = 0.3
        shadow.ShadowDepth = 2
        shadow.BlurRadius = 6
        header.Effect = shadow
        header_grid = System.Windows.Controls.Grid()
        header_grid.Margin = System.Windows.Thickness(24, 0, 24, 0)
        left_stack = System.Windows.Controls.StackPanel()
        left_stack.Orientation = System.Windows.Controls.Orientation.Horizontal
        left_stack.VerticalAlignment = System.Windows.VerticalAlignment.Center
        left_stack.HorizontalAlignment = System.Windows.HorizontalAlignment.Left
        py_text = System.Windows.Controls.TextBlock()
        py_text.Text = "py"
        py_text.FontWeight = System.Windows.FontWeights.Bold
        py_text.FontSize = 32
        py_text.Foreground = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(32, 138, 60))
        py_text.VerticalAlignment = System.Windows.VerticalAlignment.Center
        subtitle_text = System.Windows.Controls.TextBlock()
        subtitle_text.Text = subtitle
        subtitle_text.FontWeight = System.Windows.FontWeights.SemiBold
        subtitle_text.FontSize = 32
        subtitle_text.Foreground = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(244, 250, 255))
        subtitle_text.Margin = System.Windows.Thickness(4, 0, 0, 0)
        subtitle_text.VerticalAlignment = System.Windows.VerticalAlignment.Center
        left_stack.Children.Add(py_text)
        left_stack.Children.Add(subtitle_text)
        header_grid.Children.Add(left_stack)
        right_stack = System.Windows.Controls.StackPanel()
        right_stack.Orientation = System.Windows.Controls.Orientation.Horizontal
        right_stack.VerticalAlignment = System.Windows.VerticalAlignment.Center
        right_stack.HorizontalAlignment = System.Windows.HorizontalAlignment.Right
        options_grid = System.Windows.Controls.Grid()
        self.options_btn = System.Windows.Controls.Primitives.ToggleButton()
        self.options_btn.Content = "☰"
        self.options_btn.Width = 40
        self.options_btn.Height = 40
        self.options_btn.FontSize = 18
        self.options_btn.FontWeight = System.Windows.FontWeights.Bold
        self.options_btn.Foreground = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(244, 250, 255))
        self.options_btn.Background = System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Transparent)
        self.options_btn.BorderThickness = System.Windows.Thickness(0)
        self.options_btn.Cursor = System.Windows.Input.Cursors.Hand
        self.options_btn.Template = self._create_options_button_template()
        self.options_popup = System.Windows.Controls.Primitives.Popup()
        self.options_popup.PlacementTarget = self.options_btn
        self.options_popup.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom
        self.options_popup.AllowsTransparency = True
        self.options_popup.StaysOpen = False
        binding = System.Windows.Data.Binding("IsChecked")
        binding.Source = self.options_btn
        self.options_popup.SetBinding(System.Windows.Controls.Primitives.Popup.IsOpenProperty, binding)
        popup_border = System.Windows.Controls.Border()
        popup_border.Background = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(244, 250, 255))
        popup_border.BorderBrush = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(32, 138, 60))
        popup_border.BorderThickness = System.Windows.Thickness(1)
        popup_border.CornerRadius = System.Windows.CornerRadius(6)
        popup_border.MinWidth = 200
        popup_shadow = System.Windows.Media.Effects.DropShadowEffect()
        popup_shadow.Color = System.Windows.Media.Colors.Black
        popup_shadow.Opacity = 0.2
        popup_shadow.ShadowDepth = 2
        popup_shadow.BlurRadius = 6
        popup_border.Effect = popup_shadow
        popup_stack = System.Windows.Controls.StackPanel()
        help_btn = System.Windows.Controls.Button()
        help_btn.Content = "? Help"
        help_btn.Padding = System.Windows.Thickness(12, 8, 12, 8)
        help_btn.FontSize = 12
        help_btn.Foreground = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(43, 51, 64))
        help_btn.Background = System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Transparent)
        help_btn.BorderThickness = System.Windows.Thickness(0)
        help_btn.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch
        help_btn.HorizontalContentAlignment = System.Windows.HorizontalAlignment.Left
        help_btn.Cursor = System.Windows.Input.Cursors.Hand
        help_btn.Template = self._create_menu_item_template()
        help_btn.Click += self.show_help
        about_btn = System.Windows.Controls.Button()
        about_btn.Content = "? About View Placer"
        about_btn.Padding = System.Windows.Thickness(12, 8, 12, 8)
        about_btn.FontSize = 12
        about_btn.Foreground = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(43, 51, 64))
        about_btn.Background = System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Transparent)
        about_btn.BorderThickness = System.Windows.Thickness(0)
        about_btn.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch
        about_btn.HorizontalContentAlignment = System.Windows.HorizontalAlignment.Left
        about_btn.Cursor = System.Windows.Input.Cursors.Hand
        about_btn.Template = self._create_menu_item_template()
        about_btn.Click += self.show_about
        popup_stack.Children.Add(help_btn)
        popup_stack.Children.Add(about_btn)
        popup_border.Child = popup_stack
        self.options_popup.Child = popup_border
        options_grid.Children.Add(self.options_btn)
        options_grid.Children.Add(self.options_popup)
        right_stack.Children.Add(options_grid)
        header_grid.Children.Add(right_stack)
        header.Child = header_grid
        dock_panel.Children.Add(header)

    def _create_options_button_template(self):
        template = System.Windows.Controls.ControlTemplate(System.Windows.Controls.Primitives.ToggleButton)
        factory = System.Windows.FrameworkElementFactory(System.Windows.Controls.Border)
        factory.Name = "MainBorder"
        factory.SetValue(System.Windows.Controls.Border.BackgroundProperty, System.Windows.Data.Binding("Background"))
        factory.SetValue(System.Windows.Controls.Border.CornerRadiusProperty, System.Windows.CornerRadius(8))
        factory.SetValue(System.Windows.Controls.Border.PaddingProperty, System.Windows.Data.Binding("Padding"))
        content = System.Windows.FrameworkElementFactory(System.Windows.Controls.ContentPresenter)
        content.SetValue(System.Windows.Controls.ContentPresenter.HorizontalAlignmentProperty, System.Windows.HorizontalAlignment.Center)
        content.SetValue(System.Windows.Controls.ContentPresenter.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center)
        factory.AppendChild(content)
        template.VisualTree = factory
        hover_trigger = System.Windows.Trigger()
        hover_trigger.Property = System.Windows.Controls.Primitives.ToggleButton.IsMouseOverProperty
        hover_trigger.Value = True
        hover_setter = System.Windows.Setter()
        hover_setter.Property = System.Windows.Controls.Border.BackgroundProperty
        hover_setter.Value = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(64, 69, 83))
        hover_trigger.Setters.Add(hover_setter)
        template.Triggers.Add(hover_trigger)
        checked_trigger = System.Windows.Trigger()
        checked_trigger.Property = System.Windows.Controls.Primitives.ToggleButton.IsCheckedProperty
        checked_trigger.Value = True
        checked_setter = System.Windows.Setter()
        checked_setter.Property = System.Windows.Controls.Border.BackgroundProperty
        checked_setter.Value = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(32, 138, 60))
        checked_trigger.Setters.Add(checked_setter)
        template.Triggers.Add(checked_trigger)
        return template

    def _create_menu_item_template(self):
        template = System.Windows.Controls.ControlTemplate(System.Windows.Controls.Button)
        factory = System.Windows.FrameworkElementFactory(System.Windows.Controls.Border)
        factory.Name = "MainBorder"
        factory.SetValue(System.Windows.Controls.Border.BackgroundProperty, System.Windows.Data.Binding("Background"))
        factory.SetValue(System.Windows.Controls.Border.PaddingProperty, System.Windows.Data.Binding("Padding"))
        content = System.Windows.FrameworkElementFactory(System.Windows.Controls.ContentPresenter)
        content.SetValue(System.Windows.Controls.ContentPresenter.HorizontalAlignmentProperty, System.Windows.Data.Binding("HorizontalContentAlignment"))
        content.SetValue(System.Windows.Controls.ContentPresenter.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center)
        factory.AppendChild(content)
        template.VisualTree = factory
        trigger = System.Windows.Trigger()
        trigger.Property = System.Windows.Controls.Button.IsMouseOverProperty
        trigger.Value = True
        setter = System.Windows.Setter()
        setter.Property = System.Windows.Controls.Border.BackgroundProperty
        setter.Value = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(232, 245, 232))
        trigger.Setters.Add(setter)
        template.Triggers.Add(trigger)
        return template

    def create_card(self):
        card = System.Windows.Controls.Border()
        card.Background = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(43, 51, 64))
        card.CornerRadius = System.Windows.CornerRadius(8)
        card.Padding = System.Windows.Thickness(16)
        card.Margin = System.Windows.Thickness(0, 0, 0, 12)
        shadow = System.Windows.Media.Effects.DropShadowEffect()
        shadow.Color = System.Windows.Media.Colors.Black
        shadow.Opacity = 0.2
        shadow.ShadowDepth = 2
        shadow.BlurRadius = 4
        card.Effect = shadow
        stack = System.Windows.Controls.StackPanel()
        card.Child = stack
        return card, stack

    def add_section_label(self, parent, text):
        label = System.Windows.Controls.TextBlock()
        label.Text = text
        label.Foreground = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(32, 138, 60))
        label.FontWeight = System.Windows.FontWeights.SemiBold
        label.FontSize = 14
        label.Margin = System.Windows.Thickness(0, 0, 0, 8)
        parent.Children.Add(label)

    def show_help(self, sender, args):
        try:
            webbrowser.open(HELP_URL)
        except:
            pass

    def show_about(self, sender, args):
        try:
            webbrowser.open(ABOUT_URL)
        except:
            pass

    def _find_scroll_viewer(self, dep_obj):
        if dep_obj is None:
            return None
        if isinstance(dep_obj, System.Windows.Controls.ScrollViewer):
            return dep_obj
        for i in range(System.Windows.Media.VisualTreeHelper.GetChildrenCount(dep_obj)):
            child = System.Windows.Media.VisualTreeHelper.GetChild(dep_obj, i)
            result = self._find_scroll_viewer(child)
            if result is not None:
                return result
        return None
class TemplateSelectionDialog(StyledDialog):
    def __init__(self, templates, last_template):
        StyledDialog.__init__(self, "View Placer - View Template", "View Placer")
        self.templates = templates
        self.last_template = last_template
        card, card_stack = self.create_card()
        self.add_section_label(card_stack, "Select View Template")
        search_label = System.Windows.Controls.TextBlock()
        search_label.Text = "Search:"
        search_label.Foreground = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(244, 250, 255))
        search_label.FontSize = 12
        search_label.Margin = System.Windows.Thickness(0, 0, 0, 4)
        card_stack.Children.Add(search_label)
        self.search_box = System.Windows.Controls.TextBox()
        self.search_box.Height = 28
        self.search_box.Background = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(59, 69, 83))
        self.search_box.Foreground = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(244, 250, 255))
        self.search_box.BorderBrush = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(32, 138, 60))
        self.search_box.BorderThickness = System.Windows.Thickness(1)
        self.search_box.Padding = System.Windows.Thickness(8, 4, 8, 4)
        self.search_box.Margin = System.Windows.Thickness(0, 0, 0, 8)
        self.search_box.TextChanged += self.on_search_changed
        card_stack.Children.Add(self.search_box)
        self.list_box = System.Windows.Controls.ListBox()
        self.list_box.Height = 320
        self.list_box.Background = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(59, 69, 83))
        self.list_box.Foreground = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(244, 250, 255))
        self.list_box.BorderBrush = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(32, 138, 60))
        self.list_box.BorderThickness = System.Windows.Thickness(1)
        self.list_box.Padding = System.Windows.Thickness(4)
        self.list_box.MouseDoubleClick += self.on_double_click
        self.all_items = []
        selected_index = 0
        none_item = System.Windows.Controls.ListBoxItem()
        none_item.Content = "<None - No Template>"
        none_item.Tag = None
        none_item.Padding = System.Windows.Thickness(8, 4, 8, 4)
        self.list_box.Items.Add(none_item)
        self.all_items.append(none_item)
        if last_template:
            last_item = System.Windows.Controls.ListBoxItem()
            last_item.Content = "{} (Last Used)".format(last_template.Name)
            last_item.Tag = last_template
            last_item.Padding = System.Windows.Thickness(8, 4, 8, 4)
            self.list_box.Items.Add(last_item)
            self.all_items.append(last_item)
            selected_index = 1
        for template in sorted(templates, key=lambda t: t.Name):
            if not last_template or template.Id != last_template.Id:
                item = System.Windows.Controls.ListBoxItem()
                item.Content = template.Name
                item.Tag = template
                item.Padding = System.Windows.Thickness(8, 4, 8, 4)
                self.list_box.Items.Add(item)
                self.all_items.append(item)
        self.list_box.SelectedIndex = selected_index
        if selected_index > 0:
            self.list_box.UpdateLayout()
            self.list_box.ScrollIntoView(self.list_box.Items[selected_index])
            def center_template(*args):
                scroll_viewer = self._find_scroll_viewer(self.list_box)
                if not scroll_viewer:
                    return
                container = self.list_box.ItemContainerGenerator.ContainerFromIndex(selected_index)
                if not container:
                    return
                transform = container.TransformToAncestor(self.list_box)
                pos = transform.Transform(System.Windows.Point(0, 0))
                visible_height = self.list_box.ActualHeight
                target_offset = pos.Y - (visible_height / 2) + (container.ActualHeight / 2)
                max_offset = scroll_viewer.ScrollableHeight
                target_offset = max(0, min(target_offset, max_offset))
                scroll_viewer.ScrollToVerticalOffset(target_offset)
                self.list_box.LayoutUpdated -= center_template
            self.list_box.LayoutUpdated += center_template
        card_stack.Children.Add(self.list_box)
        ok_btn = System.Windows.Controls.Button()
        ok_btn.Content = "OK"
        ok_btn.Width = 100
        ok_btn.Height = 36
        ok_btn.Margin = System.Windows.Thickness(0, 12, 0, 0)
        ok_btn.Background = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(32, 138, 60))
        ok_btn.Foreground = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(244, 250, 255))
        ok_btn.BorderThickness = System.Windows.Thickness(0)
        ok_btn.FontWeight = System.Windows.FontWeights.SemiBold
        ok_btn.Cursor = System.Windows.Input.Cursors.Hand
        ok_btn.Click += self.on_ok
        card_stack.Children.Add(ok_btn)
        self.content_panel.Children.Add(card)

    def on_search_changed(self, sender, args):
        search_text = self.search_box.Text.lower()
        self.list_box.Items.Clear()
        for item in self.all_items:
            content = str(item.Content).lower()
            if not search_text or search_text in content:
                self.list_box.Items.Add(item)

    def on_double_click(self, sender, args):
        self.on_ok(sender, args)

    def on_ok(self, sender, args):
        if self.list_box.SelectedItem:
            self.result = self.list_box.SelectedItem.Tag
        else:
            self.result = None
        self.Close()

class SheetSelectionDialog(StyledDialog):
    def __init__(self, sheets, last_sheet):
        StyledDialog.__init__(self, "View Placer - Sheet Selection", "View Placer")
        self.sheets = sheets
        self.last_sheet = last_sheet
        card, card_stack = self.create_card()
        self.add_section_label(card_stack, "Select Target Sheet")
        search_label = System.Windows.Controls.TextBlock()
        search_label.Text = "Search:"
        search_label.Foreground = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(244, 250, 255))
        search_label.FontSize = 12
        search_label.Margin = System.Windows.Thickness(0, 0, 0, 4)
        card_stack.Children.Add(search_label)
        self.search_box = System.Windows.Controls.TextBox()
        self.search_box.Height = 28
        self.search_box.Background = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(59, 69, 83))
        self.search_box.Foreground = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(244, 250, 255))
        self.search_box.BorderBrush = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(32, 138, 60))
        self.search_box.BorderThickness = System.Windows.Thickness(1)
        self.search_box.Padding = System.Windows.Thickness(8, 4, 8, 4)
        self.search_box.Margin = System.Windows.Thickness(0, 0, 0, 8)
        self.search_box.TextChanged += self.on_search_changed
        card_stack.Children.Add(self.search_box)
        self.list_box = System.Windows.Controls.ListBox()
        self.list_box.Height = 320
        self.list_box.Background = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(59, 69, 83))
        self.list_box.Foreground = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(244, 250, 255))
        self.list_box.BorderBrush = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(32, 138, 60))
        self.list_box.BorderThickness = System.Windows.Thickness(1)
        self.list_box.Padding = System.Windows.Thickness(4)
        self.list_box.MouseDoubleClick += self.on_double_click
        System.Windows.Controls.ScrollViewer.SetHorizontalScrollBarVisibility(self.list_box, System.Windows.Controls.ScrollBarVisibility.Auto)
        System.Windows.Controls.ScrollViewer.SetCanContentScroll(self.list_box, False)
        self.all_items = []
        selected_index = -1
        sorted_sheets = sorted(sheets, key=lambda s: s.SheetNumber)
        for i, sheet in enumerate(sorted_sheets):
            item = System.Windows.Controls.ListBoxItem()
            if last_sheet and sheet.Id == last_sheet.Id:
                item.Content = "{} - {} (Last Used)".format(sheet.SheetNumber, sheet.Name)
                item.Background = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(40, 32, 138, 60))
                selected_index = i
            else:
                item.Content = "{} - {}".format(sheet.SheetNumber, sheet.Name)
            item.Tag = sheet
            item.Padding = System.Windows.Thickness(8, 4, 8, 4)
            self.list_box.Items.Add(item)
            self.all_items.append(item)
        if selected_index >= 0:
            self.list_box.SelectedIndex = selected_index
            self.list_box.UpdateLayout()
            self.list_box.ScrollIntoView(self.list_box.Items[selected_index])
            def center_sheet(*args):
                scroll_viewer = self._find_scroll_viewer(self.list_box)
                if not scroll_viewer:
                    return
                container = self.list_box.ItemContainerGenerator.ContainerFromIndex(selected_index)
                if not container:
                    return
                transform = container.TransformToAncestor(self.list_box)
                pos = transform.Transform(System.Windows.Point(0, 0))
                visible_height = self.list_box.ActualHeight
                target_offset = pos.Y - (visible_height / 2) + (container.ActualHeight / 2)
                max_offset = scroll_viewer.ScrollableHeight
                target_offset = max(0, min(target_offset, max_offset))
                scroll_viewer.ScrollToVerticalOffset(target_offset)
                self.list_box.LayoutUpdated -= center_sheet
            self.list_box.LayoutUpdated += center_sheet
        def force_left_horizontal(*args):
            scroll_viewer = self._find_scroll_viewer(self.list_box)
            if scroll_viewer:
                scroll_viewer.ScrollToHorizontalOffset(0)
            self.list_box.LayoutUpdated -= force_left_horizontal
        self.list_box.LayoutUpdated += force_left_horizontal
        card_stack.Children.Add(self.list_box)
        ok_btn = System.Windows.Controls.Button()
        ok_btn.Content = "OK"
        ok_btn.Width = 100
        ok_btn.Height = 36
        ok_btn.Margin = System.Windows.Thickness(0, 12, 0, 0)
        ok_btn.Background = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(32, 138, 60))
        ok_btn.Foreground = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(244, 250, 255))
        ok_btn.BorderThickness = System.Windows.Thickness(0)
        ok_btn.FontWeight = System.Windows.FontWeights.SemiBold
        ok_btn.Cursor = System.Windows.Input.Cursors.Hand
        ok_btn.Click += self.on_ok
        card_stack.Children.Add(ok_btn)
        self.content_panel.Children.Add(card)

    def on_search_changed(self, sender, args):
        search_text = self.search_box.Text.lower()
        self.list_box.Items.Clear()
        for item in self.all_items:
            content = str(item.Content).lower()
            if not search_text or search_text in content:
                self.list_box.Items.Add(item)
        def force_left_after_search(*args):
            scroll_viewer = self._find_scroll_viewer(self.list_box)
            if scroll_viewer:
                scroll_viewer.ScrollToHorizontalOffset(0)
            self.list_box.LayoutUpdated -= force_left_after_search
        self.list_box.LayoutUpdated += force_left_after_search

    def on_double_click(self, sender, args):
        self.on_ok(sender, args)

    def on_ok(self, sender, args):
        if self.list_box.SelectedItem:
            self.result = self.list_box.SelectedItem.Tag
        else:
            self.result = None
        self.Close()
class ViewportTypeDialog(StyledDialog):
    def __init__(self, viewport_types, last_viewport, get_name_func):
        StyledDialog.__init__(self, "View Placer - Viewport Type", "View Placer")
        self.viewport_types = viewport_types
        self.last_viewport = last_viewport
        self.get_name_func = get_name_func
        card, card_stack = self.create_card()
        self.add_section_label(card_stack, "Select Viewport Type")
        search_label = System.Windows.Controls.TextBlock()
        search_label.Text = "Search:"
        search_label.Foreground = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(244, 250, 255))
        search_label.FontSize = 12
        search_label.Margin = System.Windows.Thickness(0, 0, 0, 4)
        card_stack.Children.Add(search_label)
        self.search_box = System.Windows.Controls.TextBox()
        self.search_box.Height = 28
        self.search_box.Background = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(59, 69, 83))
        self.search_box.Foreground = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(244, 250, 255))
        self.search_box.BorderBrush = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(32, 138, 60))
        self.search_box.BorderThickness = System.Windows.Thickness(1)
        self.search_box.Padding = System.Windows.Thickness(8, 4, 8, 4)
        self.search_box.Margin = System.Windows.Thickness(0, 0, 0, 8)
        self.search_box.TextChanged += self.on_search_changed
        card_stack.Children.Add(self.search_box)
        self.list_box = System.Windows.Controls.ListBox()
        self.list_box.Height = 320
        self.list_box.Background = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(59, 69, 83))
        self.list_box.Foreground = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(244, 250, 255))
        self.list_box.BorderBrush = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(32, 138, 60))
        self.list_box.BorderThickness = System.Windows.Thickness(1)
        self.list_box.Padding = System.Windows.Thickness(4)
        self.list_box.MouseDoubleClick += self.on_double_click
        self.all_items = []
        selected_index = 0
        default_item = System.Windows.Controls.ListBoxItem()
        default_item.Content = "<Default - No Change>"
        default_item.Tag = None
        default_item.Padding = System.Windows.Thickness(8, 4, 8, 4)
        self.list_box.Items.Add(default_item)
        self.all_items.append(default_item)
        if last_viewport:
            last_item = System.Windows.Controls.ListBoxItem()
            last_item.Content = "{} (Last Used)".format(get_name_func(last_viewport))
            last_item.Tag = last_viewport
            last_item.Padding = System.Windows.Thickness(8, 4, 8, 4)
            self.list_box.Items.Add(last_item)
            self.all_items.append(last_item)
            selected_index = 1
        viewport_dict = {}
        for vp in viewport_types:
            vp_name = get_name_func(vp)
            if vp_name in viewport_dict:
                vp_name = "{} (ID: {})".format(vp_name, get_element_id_value(vp.Id))
            viewport_dict[vp_name] = vp
        for vp_name in sorted(viewport_dict.keys()):
            vp = viewport_dict[vp_name]
            if not last_viewport or vp.Id != last_viewport.Id:
                item = System.Windows.Controls.ListBoxItem()
                item.Content = vp_name
                item.Tag = vp
                item.Padding = System.Windows.Thickness(8, 4, 8, 4)
                self.list_box.Items.Add(item)
                self.all_items.append(item)
        self.list_box.SelectedIndex = selected_index
        if selected_index > 0:
            self.list_box.UpdateLayout()
            self.list_box.ScrollIntoView(self.list_box.Items[selected_index])
            def center_viewport(*args):
                scroll_viewer = self._find_scroll_viewer(self.list_box)
                if not scroll_viewer:
                    return
                container = self.list_box.ItemContainerGenerator.ContainerFromIndex(selected_index)
                if not container:
                    return
                transform = container.TransformToAncestor(self.list_box)
                pos = transform.Transform(System.Windows.Point(0, 0))
                visible_height = self.list_box.ActualHeight
                target_offset = pos.Y - (visible_height / 2) + (container.ActualHeight / 2)
                max_offset = scroll_viewer.ScrollableHeight
                target_offset = max(0, min(target_offset, max_offset))
                scroll_viewer.ScrollToVerticalOffset(target_offset)
                self.list_box.LayoutUpdated -= center_viewport
            self.list_box.LayoutUpdated += center_viewport
        card_stack.Children.Add(self.list_box)
        ok_btn = System.Windows.Controls.Button()
        ok_btn.Content = "OK"
        ok_btn.Width = 100
        ok_btn.Height = 36
        ok_btn.Margin = System.Windows.Thickness(0, 12, 0, 0)
        ok_btn.Background = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(32, 138, 60))
        ok_btn.Foreground = System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(244, 250, 255))
        ok_btn.BorderThickness = System.Windows.Thickness(0)
        ok_btn.FontWeight = System.Windows.FontWeights.SemiBold
        ok_btn.Cursor = System.Windows.Input.Cursors.Hand
        ok_btn.Click += self.on_ok
        card_stack.Children.Add(ok_btn)
        self.content_panel.Children.Add(card)

    def on_search_changed(self, sender, args):
        search_text = self.search_box.Text.lower()
        self.list_box.Items.Clear()
        for item in self.all_items:
            content = str(item.Content).lower()
            if not search_text or search_text in content:
                self.list_box.Items.Add(item)

    def on_double_click(self, sender, args):
        self.on_ok(sender, args)

    def on_ok(self, sender, args):
        if self.list_box.SelectedItem:
            self.result = self.list_box.SelectedItem.Tag
        else:
            self.result = None
        self.Close()

def main():
    try:
        if not revit.doc:
            forms.alert("No active Revit document found")
            return
        active_view = revit.uidoc.ActiveView
        if isinstance(active_view, DB.ViewSheet):
            forms.alert("Cannot run on a sheet view. Please open a floor plan, section, or elevation.")
            return
        view_manager = ViewManager(revit.doc, revit.uidoc)
        config = script.get_config()
        last_sheet_id = getattr(config, 'last_sheet_id', None)
        last_viewport_id = getattr(config, 'last_viewport_id', None)
        last_template_id = getattr(config, 'last_template_id', None)
        with forms.WarningBar(title='Pick view references, callouts, or sections. ESCAPE to finish.'):
            selected_views = view_manager.select_views_from_active_view()
        if not selected_views:
            return
        views_to_place = []
        views_on_sheets = []
        for view in selected_views:
            sheet_param = view.get_Parameter(DB.BuiltInParameter.VIEWER_SHEET_NUMBER)
            if sheet_param and sheet_param.AsString() == '---':
                views_to_place.append(view)
            else:
                views_on_sheets.append(view)
        if not views_to_place:
            forms.alert("All selected views are already on sheets.")
            return
        collector = DB.FilteredElementCollector(revit.doc).OfClass(DB.View).WhereElementIsNotElementType()
        templates = [view for view in collector if view.IsTemplate]
        last_template = None
        if last_template_id:
            try:
                template_id = create_element_id(last_template_id)
                last_template = revit.doc.GetElement(template_id)
                if not last_template or not last_template.IsTemplate:
                    last_template = None
            except (ValueError, TypeError, AttributeError):
                last_template = None
        template_dialog = TemplateSelectionDialog(templates, last_template)
        template_dialog.ShowDialog()
        selected_template = template_dialog.result
        if template_dialog.result is None and template_dialog.list_box.SelectedIndex == -1:
            return
        if selected_template:
            config.last_template_id = str(get_element_id_value(selected_template.Id))
        else:
            config.last_template_id = None
        script.save_config()
        sheets_collector = DB.FilteredElementCollector(revit.doc).OfClass(DB.ViewSheet).WhereElementIsNotElementType()
        all_sheets = list(sheets_collector)
        last_sheet = None
        if last_sheet_id:
            try:
                sheet_id = create_element_id(last_sheet_id)
                last_sheet = revit.doc.GetElement(sheet_id)
                if not last_sheet or not isinstance(last_sheet, DB.ViewSheet):
                    last_sheet = None
            except (ValueError, TypeError, AttributeError):
                last_sheet = None
        sheet_dialog = SheetSelectionDialog(all_sheets, last_sheet)
        sheet_dialog.ShowDialog()
        selected_sheet = sheet_dialog.result
        if not selected_sheet:
            return
        config.last_sheet_id = str(get_element_id_value(selected_sheet.Id))
        script.save_config()
        created_viewports = []
        with revit.Transaction('Place Views on Sheet'):
            for i, view in enumerate(views_to_place):
                if selected_template:
                    try:
                        view.ViewTemplateId = selected_template.Id
                    except:
                        pass
                position = DB.XYZ(1.0 + (i * 2.0), 1.5, 0.0)
                try:
                    viewport = DB.Viewport.Create(revit.doc, selected_sheet.Id, view.Id, position)
                    created_viewports.append(viewport)
                except:
                    pass
        if created_viewports:
            vp_collector = DB.FilteredElementCollector(revit.doc).OfClass(DB.ElementType).OfCategory(DB.BuiltInCategory.OST_Viewports)
            viewport_types = list(vp_collector)
            if not viewport_types:
                vp_collector2 = DB.FilteredElementCollector(revit.doc).OfClass(DB.Viewport)
                type_ids = set()
                for vp in vp_collector2:
                    type_ids.add(vp.GetTypeId())
                viewport_types = [revit.doc.GetElement(tid) for tid in type_ids if tid != DB.ElementId.InvalidElementId]
            if viewport_types:
                last_viewport = None
                if last_viewport_id:
                    try:
                        viewport_id = create_element_id(last_viewport_id)
                        last_viewport = revit.doc.GetElement(viewport_id)
                    except (ValueError, TypeError, AttributeError):
                        pass
                def get_vp_name(vp_type):
                    try:
                        name_param = vp_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
                        if name_param and name_param.HasValue:
                            return name_param.AsString()
                        name_param = vp_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                        if name_param and name_param.HasValue:
                            return name_param.AsString()
                        if hasattr(vp_type, 'Name'):
                            return vp_type.Name
                    except:
                        pass
                    return "Unnamed Viewport Type"
                viewport_dialog = ViewportTypeDialog(viewport_types, last_viewport, get_vp_name)
                viewport_dialog.ShowDialog()
                selected_viewport = viewport_dialog.result
                if viewport_dialog.result is None and viewport_dialog.list_box.SelectedIndex == -1:
                    pass
                elif selected_viewport:
                    with revit.Transaction('Change Viewport Types'):
                        for viewport in created_viewports:
                            try:
                                viewport.ChangeTypeId(selected_viewport.Id)
                            except:
                                pass
                    config.last_viewport_id = str(get_element_id_value(selected_viewport.Id))
                    script.save_config()
                else:
                    config.last_viewport_id = None
                    script.save_config()
    except Exception as ex:
        forms.alert("Unexpected error occurred: {}".format(str(ex)))

if __name__ == '__main__':
    main()