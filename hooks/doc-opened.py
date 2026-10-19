# -*- coding: utf-8 -*-
# "Seed43"
# """
# Put the amber override frame back on any view that is still overridden.
#
# The override is written into the model and survives a save and reopen.
# The frame does not - Temporary View Properties is session state - so
# without this a model comes back with its filters overridden and nothing on
# screen saying so, which is the one state these tools must never produce.
#
# It re-marks, it does not undo. Clearing an override is the close hook's
# job, where the user is asked; taking their saved work away behind their
# back on open is not the same thing.
# """

from pyrevit import EXEC_PARAMS, revit

from Snippets import _viewfilters as vfilt


def main():
    args = EXEC_PARAMS.event_args
    doc = getattr(args, "Document", None)
    if doc is None or doc.IsFamilyDocument or doc.IsLinked:
        return

    # Nothing overridden is the normal case, and it costs one quick
    # Extensible Storage filter and no transaction at all.
    if not vfilt.overridden_views(doc):
        return

    # NOTE: tried without a transaction first, on purpose. Revit's own
    # temporary view modes do not count as document edits, and if raising the
    # frame is free then a model opened with an override on it is not marked
    # as modified for the sake of a marker - no spurious "save changes?" on
    # the way back out. If Revit does demand a transaction it throws rather
    # than half-applying, notice_on swallows that, and pending tells us to
    # do it the expensive way.
    _raised, pending = vfilt.restore_notices(doc)
    if pending:
        with revit.Transaction("Seed43 Filter Override Warning", doc=doc):
            vfilt.restore_notices(doc)


try:
    main()
except Exception:
    # A hook that raises interrupts the user's file open with a traceback.
    # Failing to redraw a marker is not worth that.
    pass
