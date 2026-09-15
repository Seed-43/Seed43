# -*- coding: utf-8 -*-
# "Seed43"
# """
# Offer to clear any Seed43 filter override before Revit closes the model.
#
# Bound to the Close command rather than the doc-closing event on purpose:
# DocumentClosing is a Revit pre-event (RevitAPIPreDocEventArgs) and the API
# forbids modifying a document inside one, so a doc-closing hook could see
# the override but never undo it. This fires before the Close command runs,
# where a transaction is still legal.
#
# Asks rather than acting. Removing an override is the right default but it
# is still the user's model, and "remove and save" in particular writes
# their other unsaved work to disk, which nothing should do silently.
#
# Does nothing at all on a model with no override on it: one quick
# Extensible Storage filter, no dialog, no transaction, no modified flag.
# """

from pyrevit import revit

from Snippets import _viewfilters as vfilt

TITLE = "Seed43 Filter Overrides"


# ── DIALOG ──────────────────────────────────────────────────────────────────

def _ask(names):
    """Ask what to do with the leftover overrides. Returns a choice key.

    NOTE: the dialog libraries are imported here rather than at the top of
    the file. Closing a model with nothing overridden is by far the common
    case, and it should not pay for loading WPF.
    """
    listing = u"\n".join(u"  - " + name for name in names)
    text = (u"{} view{} in this model still have a Seed43 filter override "
            u"on:\n\n{}\n\n"
            u"Remove and save clears them and writes the model to disk now, "
            u"including any other unsaved changes.\n\n"
            u"Remove only clears them and lets Revit ask about saving as "
            u"usual.\n\n"
            u"Left in place they stay in the file, and the amber warning "
            u"frame comes back next time the model is opened."
            .format(len(names), u"" if len(names) == 1 else u"s", listing))

    options = [("leave", "Leave them"),
               ("remove", "Remove only"),
               ("save", "Remove and save")]
    try:
        from Snippets import _dialogs as sdlg
        return sdlg.choice(text, options, title=TITLE)
    except Exception:
        from pyrevit import forms
        labels = [label for _key, label in options]
        picked = forms.alert(text, title=TITLE, options=labels)
        for key, label in options:
            if label == picked:
                return key
        return None


def _save(doc):
    """Write the model to disk. Returns None on success, or why it could not.

    Deliberately Save() and not a sync: a workshared local saves locally,
    which is what "save my file before it closes" means. Pushing someone's
    work to a central model on the way out of the door is a different and
    much larger promise.
    """
    try:
        if doc.IsReadOnly:
            return u"the model is read-only"
        if not doc.PathName:
            return u"the model has never been saved"
        doc.Save()
        return None
    except Exception as ex:
        return u"{}: {}".format(type(ex).__name__, ex)


# ── ENTRY POINT ─────────────────────────────────────────────────────────────

def main():
    doc = revit.doc
    if doc is None or doc.IsFamilyDocument or doc.IsLinked:
        return

    names = vfilt.overridden_names(doc)
    if not names:
        return

    answer = _ask(names)
    if answer not in ("remove", "save"):
        return          # "leave", or the dialog was dismissed

    with revit.Transaction("Clear Seed43 Filter Overrides", doc=doc):
        vfilt.reset_document(doc)

    if answer == "save":
        problem = _save(doc)
        if problem:
            try:
                from Snippets import _dialogs as sdlg
                sdlg.message(
                    u"The overrides were removed, but the model could not "
                    u"be saved: {}.\n\nRevit will ask about saving as "
                    u"usual.".format(problem), title=TITLE)
            except Exception:
                pass


try:
    main()
except Exception:
    # Never let tidying up get between the user and closing their file.
    pass
