# -*- coding: utf-8 -*-
# "Filters Show"
# "Seed43"
# """
# Tick Visibility on for every filter in the active view, and run again to
# put the column back exactly as it was. All the real work lives in
# Snippets/_viewfilters.py, shared with Filters Off.
# """

from pyrevit import revit, forms

from Snippets import _viewfilters as vfilt

try:
    from Snippets import _dialogs as sdlg
except Exception:
    sdlg = None

TITLE = "Filters Show"
KIND = vfilt.VISIBLE


# ── DIALOGS ─────────────────────────────────────────────────────────────────

def _alert(message):
    """Themed popup via the shared Snippets dialog lib, falling back to
    pyRevit's default forms.alert if the shared lib isn't available."""
    if sdlg:
        sdlg.message(message, title=TITLE)
    else:
        forms.alert(message, title=TITLE)


# ── ENTRY POINT ─────────────────────────────────────────────────────────────

def main():
    doc = revit.doc
    view = doc.ActiveView

    if not vfilt.supports_filters(view):
        _alert(u"Open a view that has a Filters tab in "
               u"Visibility/Graphics.\n\nSheets, schedules and view "
               u"templates have no filter checkboxes to override.")
        return

    if not vfilt.applied_filters(view):
        _alert(u"No filters are applied to this view, so there is nothing to "
               u"make visible.")
        return

    if vfilt.notice_state(view) == "foreign":
        _alert(u"This view is already in Temporary View Properties mode.\n\n"
               u"Restore its view properties first, otherwise Revit would "
               u"discard the override the moment that mode ends.")
        return

    with revit.Transaction(TITLE):
        result = vfilt.toggle(view, KIND)

    # A clean run only speaks up the first time each direction is used - see
    # _viewfilters.report(). Anything that went wrong always speaks up.
    # The template, when there is one, is only named in the guide: the
    # override is on this view alone, but reapplying the template clears it.
    template = vfilt.controlling_template(doc, view)
    message = vfilt.report(
        result, template.Name if template is not None else None)
    if message:
        _alert(message)


main()
