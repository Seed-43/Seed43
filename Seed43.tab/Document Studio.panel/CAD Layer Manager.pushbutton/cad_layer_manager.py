# -*- coding: utf-8 -*-
# cad_layer_manager.py
import os
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import Window
from System.Windows.Controls import CheckBox, ListBoxItem
from System.Windows.Markup import XamlReader
from System.IO import File

from pyrevit import revit, DB, forms, script

# ── [LIB] Snippets/_selection.py ─────────────────────────────────────────────
from Snippets._selection import resolve_cad_instance

doc         = revit.doc
uidoc       = revit.uidoc
active_view = revit.active_view

XAML_PATH = os.path.join(os.path.dirname(__file__), "LayerManager.xaml")

# ── RESOLVE TARGET VIEW ───────────────────────────────────────────────────────

template_id = active_view.ViewTemplateId
if template_id != DB.ElementId.InvalidElementId:
    target_view    = doc.GetElement(template_id)
    using_template = True
else:
    target_view    = active_view
    using_template = False

# ── RESOLVE CAD AND LAYERS ────────────────────────────────────────────────────

cad_instance = resolve_cad_instance(uidoc, doc, revit, forms, script)

if not isinstance(cad_instance, DB.ImportInstance):
    forms.alert("Selected element is not a CAD Import or Link.", exitscript=True)

cad_category  = cad_instance.Category
subcategories = list(cad_category.SubCategories)

if not subcategories:
    forms.alert("No layers found in this CAD file.", exitscript=True)

subcategories.sort(key=lambda sc: sc.Name)

original_visibility = {}
for sc in subcategories:
    try:
        original_visibility[sc.Id] = target_view.GetCategoryHidden(sc.Id)
    except Exception:
        original_visibility[sc.Id] = False

# ── LAYER DATA MODEL ──────────────────────────────────────────────────────────

class LayerItem(object):
    def __init__(self, subcat):
        self.name     = subcat.Name
        self.subcat   = subcat
        self.can_hide = True
        try:
            self.visible = not target_view.GetCategoryHidden(subcat.Id)
        except Exception:
            self.visible  = True
            self.can_hide = False

layer_items = [LayerItem(sc) for sc in subcategories]

# ── WINDOW CONTROLLER ─────────────────────────────────────────────────────────

class LayerManagerWindow(object):

    def __init__(self):
        xaml_text   = File.ReadAllText(XAML_PATH)
        self.window = XamlReader.Parse(xaml_text)

        self.cad_name_lbl       = self.window.FindName("cad_name_lbl")
        self.search_tb          = self.window.FindName("search_tb")
        self.search_placeholder = self.window.FindName("search_placeholder_lbl")
        self.layer_list         = self.window.FindName("layer_list")
        self.layer_count_lbl    = self.window.FindName("layer_count_lbl")
        self.show_all_btn       = self.window.FindName("show_all_btn")
        self.hide_all_btn       = self.window.FindName("hide_all_btn")
        self.apply_btn          = self.window.FindName("apply_btn")
        self.cancel_btn         = self.window.FindName("cancel_btn")
        self.header_close_btn   = self.window.FindName("header_close_btn")
        self.modal_overlay      = self.window.FindName("modal_overlay")
        self.alert_popup        = self.window.FindName("alert_popup")
        self.alert_title_lbl    = self.window.FindName("alert_title_lbl")
        self.alert_msg_lbl      = self.window.FindName("alert_msg_lbl")
        self.alert_ok_btn       = self.window.FindName("alert_ok_btn")
        self.confirm_popup      = self.window.FindName("confirm_popup")
        self.confirm_yes_btn    = self.window.FindName("confirm_yes_btn")
        self.confirm_no_btn     = self.window.FindName("confirm_no_btn")

        self._result = None
        self._bind_events()
        self._populate(layer_items)

        if using_template:
            self.cad_name_lbl.Text = "{} (via template: {})".format(
                cad_category.Name or "Unknown CAD file",
                target_view.Name)
        else:
            self.cad_name_lbl.Text = cad_category.Name or "Unknown CAD file"

    # ── Event binding ─────────────────────────────────────────────────────────

    def _bind_events(self):
        self.apply_btn.Click        += self._on_apply
        self.cancel_btn.Click       += self._on_cancel_request
        self.header_close_btn.Click += self._on_cancel_request
        self.window.Closing         += self._on_window_closing
        self.show_all_btn.Click     += self._on_show_all
        self.hide_all_btn.Click     += self._on_hide_all
        self.search_tb.TextChanged  += self._on_search
        self.search_tb.GotFocus     += lambda s, e: self._set_placeholder(False)
        self.search_tb.LostFocus    += lambda s, e: self._set_placeholder(
            not self.search_tb.Text)
        self.alert_ok_btn.Click    += self._close_alert
        self.confirm_yes_btn.Click += self._on_confirm_revert
        self.confirm_no_btn.Click  += self._close_confirm

    # ── Populate list ─────────────────────────────────────────────────────────

    def _make_row(self, layer):
        cb           = CheckBox()
        cb.Content   = layer.name
        cb.IsChecked = layer.visible
        cb.Tag       = layer
        cb.Style     = self.window.Resources["LayerCheckBoxStyle"]

        if not layer.can_hide:
            cb.IsEnabled = False
            cb.Opacity   = 0.4

        def on_check(sender, e, _layer=layer):
            if not _layer.can_hide:
                return
            visible        = bool(sender.IsChecked)
            _layer.visible = visible
            with revit.Transaction("Live CAD Layer Visibility", swallow_errors=True):
                try:
                    target_view.SetCategoryHidden(_layer.subcat.Id, not visible)
                except Exception:
                    _layer.can_hide = False
            self._refresh_count()

        cb.Checked   += on_check
        cb.Unchecked += on_check

        item         = ListBoxItem()
        item.Style   = self.window.Resources["LayerRowStyle"]
        item.Content = cb
        return item

    def _populate(self, items):
        self.layer_list.Items.Clear()
        for layer in items:
            self.layer_list.Items.Add(self._make_row(layer))
        self._refresh_count()

    def _refresh_count(self):
        total      = len(layer_items)
        visible    = sum(1 for l in layer_items if l.visible)
        locked     = sum(1 for l in layer_items if not l.can_hide)
        count_text = "{} layers, {} visible, {} hidden".format(
            total, visible, total - visible)
        if locked:
            count_text += ", {} locked".format(locked)
        self.layer_count_lbl.Text = count_text

    # ── Search ────────────────────────────────────────────────────────────────

    def _on_search(self, sender, e):
        query    = (self.search_tb.Text or "").strip().lower()
        filtered = [l for l in layer_items if not query or query in l.name.lower()]
        self._populate(filtered)

    def _set_placeholder(self, show):
        from System.Windows import Visibility
        self.search_placeholder.Visibility = (
            Visibility.Visible if show else Visibility.Collapsed)

    # ── Show and hide all ─────────────────────────────────────────────────────

    def _set_all(self, visible):
        with revit.Transaction("CAD Layer Visibility - Bulk", swallow_errors=True):
            for layer in layer_items:
                if not layer.can_hide:
                    continue
                layer.visible = visible
                try:
                    target_view.SetCategoryHidden(layer.subcat.Id, not visible)
                except Exception:
                    layer.can_hide = False
        query    = (self.search_tb.Text or "").strip().lower()
        filtered = [l for l in layer_items if not query or query in l.name.lower()]
        self._populate(filtered)

    def _on_show_all(self, sender, e):
        self._set_all(True)

    def _on_hide_all(self, sender, e):
        self._set_all(False)

    # ── Apply ─────────────────────────────────────────────────────────────────

    def _on_apply(self, sender, e):
        self._result = "apply"
        self.window.Close()

    # ── Cancel and revert ─────────────────────────────────────────────────────

    def _on_cancel_request(self, sender, e):
        self._show_confirm()

    def _on_window_closing(self, sender, e):
        if self._result is None:
            self._do_revert()

    def _do_revert(self):
        with revit.Transaction("Revert CAD Layer Visibility"):
            for sc_id, was_hidden in original_visibility.items():
                try:
                    target_view.SetCategoryHidden(sc_id, was_hidden)
                except Exception:
                    pass

    # ── Confirm popup ─────────────────────────────────────────────────────────

    def _show_confirm(self):
        from System.Windows import Visibility
        self.modal_overlay.Visibility = Visibility.Visible
        self.confirm_popup.Visibility = Visibility.Visible

    def _close_confirm(self, sender=None, e=None):
        from System.Windows import Visibility
        self.modal_overlay.Visibility = Visibility.Collapsed
        self.confirm_popup.Visibility = Visibility.Collapsed

    def _on_confirm_revert(self, sender, e):
        self._close_confirm()
        self._result = "cancel"
        self._do_revert()
        self.window.Close()

    # ── Alert popup ───────────────────────────────────────────────────────────

    def show_alert(self, message, title="Layer Manager"):
        from System.Windows import Visibility
        self.alert_title_lbl.Text     = title
        self.alert_msg_lbl.Text       = message
        self.modal_overlay.Visibility = Visibility.Visible
        self.alert_popup.Visibility   = Visibility.Visible

    def _close_alert(self, sender=None, e=None):
        from System.Windows import Visibility
        self.modal_overlay.Visibility = Visibility.Collapsed
        self.alert_popup.Visibility   = Visibility.Collapsed

    # ── Show ──────────────────────────────────────────────────────────────────

    def show(self):
        self._set_placeholder(True)
        self.window.ShowDialog()
        return self._result

# ── LAUNCH ────────────────────────────────────────────────────────────────────

ui = LayerManagerWindow()
ui.show()
