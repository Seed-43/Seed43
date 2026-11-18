# -*- coding: utf-8 -*-
# cad_layer_manager.py
import os
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from pyrevit import revit, DB, forms, script
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

from System.Windows.Controls import CheckBox, ListBoxItem
from System.Windows.Markup import XamlReader
from System.Windows import Visibility
from System.Windows.Media.Imaging import BitmapImage
from System import Uri, UriKind
from System.IO import File

# ── [LIB] Snippets/_selection.py ─────────────────────────────────────────────
from Snippets._selection import resolve_cad_instance
from Snippets.seed43_theme import apply_seed43_palette

doc         = revit.doc
uidoc       = revit.uidoc
active_view = revit.active_view

XAML_PATH = os.path.join(os.path.dirname(__file__), "LayerManager.xaml")
ICON_PATH = os.path.join(os.path.dirname(__file__), "icon.png")

# ── RESOLVE TARGET VIEW ───────────────────────────────────────────────────────

template_id = active_view.ViewTemplateId
if template_id != DB.ElementId.InvalidElementId:
    target_view    = doc.GetElement(template_id)
    using_template = True
else:
    target_view    = active_view
    using_template = False

# ── RESOLVE CAD AND LAYERS ────────────────────────────────────────────────────

import cad_relink

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


def _load(instance):
    """Point the window at a CAD instance: its layers and how the view
    shows them now. Called at start, and again after a relink (the new link
    has new layer ids)."""
    global cad_instance, cad_category, subcategories, original_visibility, layer_items
    cad_instance  = instance
    cad_category  = instance.Category
    subcategories = sorted(cad_category.SubCategories, key=lambda sc: sc.Name)
    original_visibility = {}
    for sc in subcategories:
        try:
            original_visibility[sc.Id] = target_view.GetCategoryHidden(sc.Id)
        except Exception:
            original_visibility[sc.Id] = False
    layer_items = [LayerItem(sc) for sc in subcategories]


cad_instance = resolve_cad_instance(uidoc, doc, revit, forms, script)

if not isinstance(cad_instance, DB.ImportInstance):
    forms.alert("Selected element is not a CAD Import or Link.", exitscript=True)

_load(cad_instance)

if not subcategories:
    forms.alert("No layers found in this CAD file.", exitscript=True)

# ── WINDOW CONTROLLER ─────────────────────────────────────────────────────────

class LayerManagerWindow(object):

    def __init__(self):
        xaml_text   = File.ReadAllText(XAML_PATH)
        self.window = XamlReader.Parse(xaml_text)
        apply_seed43_palette(self.window, os.path.dirname(__file__))

        self.cad_name_lbl       = self.window.FindName("cad_name_lbl")
        self.search_tb          = self.window.FindName("search_tb")
        self.search_placeholder = self.window.FindName("search_placeholder_lbl")
        self.layer_list         = self.window.FindName("layer_list")
        self.layer_count_lbl    = self.window.FindName("layer_count_lbl")
        self.show_all_btn       = self.window.FindName("show_all_btn")
        self.hide_all_btn       = self.window.FindName("hide_all_btn")
        self.pick_btn           = self.window.FindName("pick_btn")
        self.colour_btns        = {"bw": self.window.FindName("bw_btn"),
                                   "invert": self.window.FindName("invert_btn"),
                                   "preserve": self.window.FindName("preserve_btn")}
        self.colour_note_lbl    = self.window.FindName("colour_note_lbl")
        self.apply_btn          = self.window.FindName("apply_btn")
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
        self.colour_mode = None

        ok, why = cad_relink.can_relink(doc, cad_instance)
        for btn in self.colour_btns.values():
            btn.IsEnabled = ok
        self.colour_note_lbl.Text = (
            "Links this DWG again with these colours, in every view it is "
            "placed in, at the same place." if ok else why)

        if os.path.exists(ICON_PATH):
            img       = self.window.FindName("header_icon")
            bmp       = BitmapImage()
            bmp.BeginInit()
            bmp.UriSource = Uri(ICON_PATH, UriKind.Absolute)
            bmp.EndInit()
            img.Source = bmp

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
        self.header_close_btn.Click += self._on_cancel_request
        self.window.Closing         += self._on_window_closing
        self.show_all_btn.Click     += self._on_show_all
        self.hide_all_btn.Click     += self._on_hide_all
        self.pick_btn.Click         += self._on_pick
        for mode, btn in self.colour_btns.items():
            btn.Click += self._colour_handler(mode)
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
            with revit.Transaction("CAD Layer Manager - Live Update"):
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
        self.search_placeholder.Visibility = (
            Visibility.Visible if show else Visibility.Collapsed)

    # ── Show and hide all ─────────────────────────────────────────────────────

    def _set_all(self, visible):
        with revit.Transaction("CAD Layer Manager - Bulk Visibility"):
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

    # ── Pick on model ─────────────────────────────────────────────────────────

    def _on_pick(self, sender, e):
        # NOTE: picking can't happen while this modal window is up, so it
        # closes with "pick" (which skips the revert) and the launch loop
        # below runs pick_layers_to_hide, then opens the window again.
        self._result = "pick"
        self.window.Close()

    # ── Relink with other colours ─────────────────────────────────────────────

    def _colour_handler(self, mode):
        # NOTE: like Pick, the relink runs once this window is closed: the
        # link's layers are replaced, so the window reopens on the new link.
        # Layer changes made so far are kept (the relink carries each view's
        # layer visibility across).
        def clicked(sender, e):
            self.colour_mode = mode
            self._result = "relink"
            self.window.Close()
        return clicked

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
        with revit.Transaction("CAD Layer Manager - Revert"):
            for sc_id, was_hidden in original_visibility.items():
                try:
                    target_view.SetCategoryHidden(sc_id, was_hidden)
                except Exception:
                    pass

    # ── Confirm popup ─────────────────────────────────────────────────────────

    def _show_confirm(self):
        self.modal_overlay.Visibility = Visibility.Visible
        self.confirm_popup.Visibility = Visibility.Visible

    def _close_confirm(self, sender=None, e=None):
        self.modal_overlay.Visibility = Visibility.Collapsed
        self.confirm_popup.Visibility = Visibility.Collapsed

    def _on_confirm_revert(self, sender, e):
        self._close_confirm()
        self._result = "cancel"
        self._do_revert()
        self.window.Close()

    # ── Alert popup ───────────────────────────────────────────────────────────

    def show_alert(self, message, title="Layer Manager"):
        self.alert_title_lbl.Text     = title
        self.alert_msg_lbl.Text       = message
        self.modal_overlay.Visibility = Visibility.Visible
        self.alert_popup.Visibility   = Visibility.Visible

    def _close_alert(self, sender=None, e=None):
        self.modal_overlay.Visibility = Visibility.Collapsed
        self.alert_popup.Visibility   = Visibility.Collapsed

    # ── Show ──────────────────────────────────────────────────────────────────

    def show(self):
        self._set_placeholder(True)
        self.window.ShowDialog()
        return self._result

# ── PICK ON MODEL ─────────────────────────────────────────────────────────────

class _CadOnlyFilter(ISelectionFilter):
    """Only lines of the CAD file this window manages can be picked."""
    def AllowElement(self, element):
        return element.Id == cad_instance.Id

    def AllowReference(self, reference, position):
        return True


def _layer_of(reference):
    """The LayerItem a picked CAD line belongs to, or None. The picked
    geometry's GraphicsStyle carries its layer as a subcategory."""
    try:
        geo = cad_instance.GetGeometryObjectFromReference(reference)
        style = doc.GetElement(geo.GraphicsStyleId) if geo else None
        cat = style.GraphicsStyleCategory if style else None
    except Exception:
        return None
    if cat is None:
        return None
    for layer in layer_items:
        if layer.subcat.Id == cat.Id:
            return layer
    return None


def pick_layers_to_hide():
    """Hide the layer of each line picked, until Esc."""
    hidden = []
    with forms.WarningBar(title="Pick CAD lines to hide their layer. "
                                "Esc to go back to the Layer Manager.") as bar:
        while True:
            try:
                ref = uidoc.Selection.PickObject(
                    ObjectType.PointOnElement, _CadOnlyFilter(),
                    "Pick a line to hide its layer")
            except OperationCanceledException:
                break
            layer = _layer_of(ref)
            if layer is None or not layer.can_hide or not layer.visible:
                continue
            with revit.Transaction("CAD Layer Manager - Pick to Hide"):
                try:
                    target_view.SetCategoryHidden(layer.subcat.Id, True)
                    layer.visible = False
                    hidden.append(layer.name)
                except Exception:
                    layer.can_hide = False
            try:
                bar.message_tb.Text = "Hidden: {}. Pick the next line, "                                       "or Esc to go back.".format(layer.name)
            except Exception:
                pass
    return hidden

# ── LAUNCH ────────────────────────────────────────────────────────────────────

def relink_colours(mode):
    """Relink the CAD file with other colours and point the window at the
    new link placed in the active view. One undo."""
    view_id = cad_instance.OwnerViewId
    try:
        with revit.Transaction("CAD Layer Manager - Relink Colours"):
            made, notes = cad_relink.relink(doc, cad_instance, mode)
    except Exception as err:
        forms.alert("The CAD file could not be relinked, nothing was changed.\n\n{}".format(err))
        return
    new = made.get(view_id) or list(made.values())[0]
    _load(new)
    if notes:
        forms.alert("Relinked. Some layer settings could not be carried across:\n\n"
                    + "\n".join(notes[:12]))


while True:
    ui = LayerManagerWindow()
    result = ui.show()
    if result == "pick":
        pick_layers_to_hide()
    elif result == "relink":
        relink_colours(ui.colour_mode)
    else:
        break
