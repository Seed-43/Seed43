# -*- coding: utf-8 -*-
"""Blanket overrides of the active view's filter checkboxes, from one button.

The window offers each column by the action it will do next:

    Enable Filter  Filters Off   (untick every filter)
                   Filters On    (the Off override is on: put them back)
    Visibility     Filters Show  (tick Visibility on every filter)
                   Filters Hide  (the Show override is on: put them back)
    either on      Restore Filters  (undo every override on this view)

"Putting back" restores each filter to what it was before the override,
from the record _viewfilters keeps in the model, so a filter that was
already unticked stays unticked. Type part of a name and Enter. All the
real work lives in Snippets/_viewfilters.py. Replaces the separate Filters
Off and Filters Show buttons.
"""
from pyrevit import revit, forms

from Snippets import _viewfilters as vfilt

try:
    from Snippets import _dialogs as sdlg
except Exception:
    sdlg = None

TITLE = "Filters"

# Per column: the label when it is not overridden (apply), and when it is
# (restore). The label is also the transaction name, so Undo reads the same.
COLUMNS = (
    (vfilt.ENABLE, u"Filters Off", u"Filters On"),
    (vfilt.VISIBLE, u"Filters Show", u"Filters Hide"),
)
RESTORE_ALL = u"Restore Filters"


# ── DIALOGS ─────────────────────────────────────────────────────────────────

def _alert(message, title=TITLE):
    """Themed popup via the shared Snippets dialog lib, falling back to
    pyRevit's default forms.alert if the shared lib isn't available."""
    if sdlg:
        sdlg.message(message, title=title)
    else:
        forms.alert(message, title=title)


def choose_action(view):
    """(label, kinds) to toggle, or None if the window was closed.

    CommandSwitchWindow filters as you type, so a few letters then Enter
    picks without the mouse.
    """
    active = vfilt.active_kinds(view)
    options = []
    for kind, apply_label, restore_label in COLUMNS:
        label = restore_label if kind in active else apply_label
        options.append((label, [kind]))
    if active:
        options.append((RESTORE_ALL, list(active)))

    by_label = dict(options)
    chosen = forms.CommandSwitchWindow.show(
        [label for label, _kinds in options],
        message=u"View filters: type a name, then Enter")
    if not chosen:
        return None
    return chosen, by_label[chosen]


def _template_name(doc, view):
    template = vfilt.controlling_template(doc, view)
    return template.Name if template is not None else None


# ── ENTRY POINT ─────────────────────────────────────────────────────────────

def main():
    doc = revit.doc
    view = doc.ActiveView

    if not vfilt.supports_filters(view):
        _alert(u"Open a view that has a Filters tab in "
               u"Visibility/Graphics.\n\nSheets, schedules and view "
               u"templates have no filter checkboxes to override.")
        return

    # A recorded override can still be restored after its filters were
    # removed from the view, so only refuse when there is nothing either way.
    if not vfilt.applied_filters(view) and not vfilt.active_kinds(view):
        _alert(u"No filters are applied to this view, so there is nothing "
               u"to override.")
        return

    if vfilt.notice_state(view) == "foreign":
        _alert(u"This view is already in Temporary View Properties mode.\n\n"
               u"Restore its view properties first, otherwise Revit would "
               u"discard the override the moment that mode ends.")
        return

    action = choose_action(view)
    if action is None:
        return
    label, kinds = action

    # One transaction for the lot, so Restore Filters is a single undo.
    # toggle() on a column that is overridden restores it, which is what
    # every restore label here means.
    with revit.Transaction(label):
        results = [vfilt.toggle(view, kind) for kind in kinds]

    # A clean run only speaks up the first time each direction is used, see
    # _viewfilters.report(). Anything that went wrong always speaks up.
    template = _template_name(doc, view)
    messages = [vfilt.report(r, template) for r in results]
    messages = [m for m in messages if m]
    if messages:
        _alert(u"\n\n".join(messages), title=label)


main()
