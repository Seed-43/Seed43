# -*- coding: utf-8 -*-
"""Keeps the taggable Timber Treatment text in step with the Treatment
dropdown on the (Timber) families.

Revit tags cannot read a Family Type parameter, so each timber member also
carries a shared text parameter, Timber Treatment, that the tags read. This
updater copies the dropdown's choice (the type name of the nested Treatment
family: H1.2, H3.2, H4, H5, Untreated) into that text whenever a detail item,
structural framing or structural column member is added or changed. A choice
from any other family (the 2D dropdown also lists the project's generic
annotations, such as North Point) is ignored.

Registered from startup.py. Optional, so a model opened on a machine without
Seed43 raises no warning; there the text just stops following the dropdown
until the member is next changed on a machine with Seed43.

Target: Revit 2022-2026, IronPython 2.
"""
import clr
from System import Guid
from System.Collections.Generic import List
from Autodesk.Revit import DB

UPDATER_GUID = Guid("6f0b9d0e-3c55-4a43-9c1e-5e3d43a1b7c2")
DROPDOWN = "Treatment"
TEXT = "Timber Treatment"
NESTED_FAMILY = "Treatment"

_CATEGORIES = (DB.BuiltInCategory.OST_DetailComponents,
               DB.BuiltInCategory.OST_StructuralFraming,
               DB.BuiltInCategory.OST_StructuralColumns)


def chosen_treatment(doc, element):
    """The dropdown's type name, or None when there is nothing to copy."""
    p = element.LookupParameter(DROPDOWN)
    if p is None or p.StorageType != DB.StorageType.ElementId:
        return None
    ref = doc.GetElement(p.AsElementId())
    # nested types the project does not load itself come back as
    # NestedFamilyTypeReference; project-loaded ones as FamilySymbol
    if isinstance(ref, DB.NestedFamilyTypeReference):
        return ref.TypeName if ref.FamilyName == NESTED_FAMILY else None
    if isinstance(ref, DB.FamilySymbol) and ref.FamilyName == NESTED_FAMILY:
        return DB.Element.Name.GetValue(ref)
    return None


def sync_element(doc, element):
    """Copy the dropdown into the text; True if it changed anything."""
    text = element.LookupParameter(TEXT)
    if text is None or text.IsReadOnly:
        return False
    want = chosen_treatment(doc, element)
    if want is None or text.AsString() == want:
        return False
    text.Set(want)
    return True


class TimberTreatmentUpdater(DB.IUpdater):
    def __init__(self, addin_id):
        self._id = DB.UpdaterId(addin_id, UPDATER_GUID)

    def Execute(self, data):
        try:
            doc = data.GetDocument()
            ids = list(data.GetAddedElementIds()) + list(data.GetModifiedElementIds())
            for i in ids:
                e = doc.GetElement(i)
                if isinstance(e, DB.FamilyInstance):
                    sync_element(doc, e)
        except Exception:
            pass                 # never break the user's edit over a tag text

    def GetUpdaterId(self):
        return self._id

    def GetUpdaterName(self):
        return "Seed43 Timber Treatment"

    def GetAdditionalInformation(self):
        return "Copies the timber Treatment dropdown into the taggable Timber Treatment text."

    def GetChangePriority(self):
        return DB.ChangePriority.DetailComponents


_registered = [None]


def register(addin_id):
    """(Re)register the updater; safe to call again after a pyRevit reload."""
    uid = DB.UpdaterId(addin_id, UPDATER_GUID)
    if DB.UpdaterRegistry.IsUpdaterRegistered(uid):
        DB.UpdaterRegistry.UnregisterUpdater(uid)
    updater = TimberTreatmentUpdater(addin_id)
    DB.UpdaterRegistry.RegisterUpdater(updater, True)
    cats = List[DB.BuiltInCategory](list(_CATEGORIES))
    flt = DB.ElementMulticategoryFilter(cats)
    DB.UpdaterRegistry.AddTrigger(uid, flt, DB.Element.GetChangeTypeAny())
    DB.UpdaterRegistry.AddTrigger(uid, flt, DB.Element.GetChangeTypeElementAddition())
    _registered[0] = updater          # keep the Python object alive
    return updater
