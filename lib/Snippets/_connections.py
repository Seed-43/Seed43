# -*- coding: utf-8 -*-
"""Steel connection types: reading them, and moving them between models.

WHAT THE REVIT API WILL AND WILL NOT DO HERE
    A StructuralConnectionHandlerType's family name is read-only everywhere
    except one place, the familyName argument of Create. ConnectionGuid is
    what selects the Advance Steel algorithm.

    Two mechanisms produce a new type and each keeps the half the other
    drops:

      Create        names the family, RESETS the Modify Parameters
      CopyElements  keeps the Modify Parameters, cannot change the family

    The Modify Parameters values - plate sizes, bolt grades, layouts - live
    in the Advance Steel object store. No property, no parameter and no
    extensible storage entity reaches them, so this module can neither carry
    them itself nor report what was lost. Anything that must preserve them
    has to go through transfer(), never through Create.

THE TRAP THIS MODULE EXISTS TO CONTAIN
    CopyElements RENAMES THE SOURCE. The type it was copied FROM comes back
    named "Ref" with its Type Mark cleared, in its own document, and the
    copy arrives under the same default name. Nothing warns about it. A tool
    that checks only what arrived will strip the names out of the model it
    was meant to be reading from and not notice.

    transfer() restores both ends. Do not copy connection types any other
    way.

snippets.yaml entry:
  _connections.py:
    description: >
      Reading steel connection types, and moving them between documents
      with CopyElements without letting it strip the names off the source.
"""

from Autodesk.Revit.DB import (BuiltInCategory, CopyPasteOptions,
                               DuplicateTypeAction, Element,
                               ElementTransformUtils, FilteredElementCollector,
                               IDuplicateTypeNamesHandler, StorageType,
                               Transaction)
from Autodesk.Revit.DB.Structure import (StructuralConnectionHandler,
                                         StructuralConnectionHandlerType)
from System.Collections.Generic import List

__all__ = ["CONNECTION_CATEGORY", "MARK", "eid", "element_name", "family_name",
           "param_text", "read_param", "is_connection", "connection_types",
           "placed_counts", "find_type", "identity", "restore_identity",
           "transfer", "Transferred"]


# ── CONSTANTS ───────────────────────────────────────────────────────────────

# Real connections only. Sub-Connections (-2009033) holds copes, mitres and
# saw cuts, which modify one member rather than joining two, and a Generic
# Connection is Revit's empty placeholder waiting for a real one to be
# chosen. Neither is worth listing, cloning or filing in a library.
CONNECTION_CATEGORY = int(BuiltInCategory.OST_StructConnections)

MARK = "Type Mark"


# ── READING ─────────────────────────────────────────────────────────────────

def eid(element_id):
    """The integer behind an ElementId, on old and new Revit alike.

    IntegerValue was removed in 2026 and Value did not arrive until 2024, so
    neither name alone spans the versions these tools support.
    """
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue


def element_name(element):
    """An element's Name, around IronPython's property binder.

    ElementType redeclares Name over the one it inherits from Element, and
    IronPython's binder can end up resolving neither - element.Name then
    raises AttributeError: Name on a property that plainly exists. Reaching
    the reflected property through the base class picks the right one.

    "Type Name" is the last resort: read-only, but it holds the same string
    and is not reached through the binder at all, so it survives whatever
    the other two routes hit.
    """
    try:
        return element.Name
    except AttributeError:
        pass
    try:
        return Element.Name.GetValue(element)
    except Exception:
        return param_text(element, "Type Name")


def family_name(symbol):
    """A type's FamilyName, with the same guard as element_name."""
    try:
        return symbol.FamilyName
    except AttributeError:
        return param_text(symbol, "Family Name")


def param_text(element, name):
    """A named string parameter's value, or "" if absent or empty."""
    param = element.LookupParameter(name)
    if param is None or param.StorageType != StorageType.String:
        return u""
    return param.AsString() or u""


def read_param(param):
    """A parameter's value ready to write elsewhere, or None to skip it.

    Empty and zero come back as None on purpose. Writing a blank over a
    field the target has already filled in gains nothing, and a Cost of zero
    is indistinguishable from a Cost nobody set.
    """
    if param.StorageType == StorageType.String:
        text = param.AsString()
        return text if text else None
    if param.StorageType == StorageType.Double:
        value = param.AsDouble()
        return value if value else None
    if param.StorageType == StorageType.Integer:
        value = param.AsInteger()
        return value if value else None
    return None


def is_connection(symbol):
    """True for a type worth listing: see CONNECTION_CATEGORY.

    A type with no category is treated as not a connection. That is not a
    state a model should reach, but reading it as one would be the wrong way
    round: better to leave an oddity out than to offer a row that cannot be
    served.
    """
    category = symbol.Category
    if category is None or eid(category.Id) != CONNECTION_CATEGORY:
        return False
    try:
        return not symbol.IsGeneric()
    except Exception:
        return True


def connection_types(document):
    """Every real connection type in document, family then type order."""
    found = [s for s in FilteredElementCollector(document).OfClass(
        StructuralConnectionHandlerType) if is_connection(s)]
    found.sort(key=lambda s: (family_name(s).lower(), element_name(s).lower()))
    return found


def placed_counts(document):
    """How many connections are placed on each type id."""
    counts = {}
    for handler in FilteredElementCollector(document).OfClass(
            StructuralConnectionHandler):
        key = eid(handler.GetTypeId())
        counts[key] = counts.get(key, 0) + 1
    return counts


def find_type(document, family, name):
    """The type in document with this family and name, or None."""
    for symbol in FilteredElementCollector(document).OfClass(
            StructuralConnectionHandlerType):
        if family_name(symbol) == family and element_name(symbol) == name:
            return symbol
    return None


# ── MOVING BETWEEN DOCUMENTS ────────────────────────────────────────────────

class _UseDestination(IDuplicateTypeNamesHandler):
    """Answer Revit's duplicate-types question without a modal dialog.

    Left to itself CopyElements puts a prompt on screen and blocks Revit
    until somebody clicks it, which for a tool driving copies in a loop
    means a stalled session and no way to tell it apart from a hang.
    UseDestinationTypes is the safe answer here because the caller checks
    that what came back is genuinely new; a reused type is reported rather
    than renamed, so answering this way can never overwrite something that
    was already in the destination.
    """

    def OnDuplicateTypeNamesFound(self, args):
        return DuplicateTypeAction.UseDestinationTypes


class Transferred(object):
    """What one transfer() pass did, split by what the user must do next."""

    def __init__(self):
        self.copied = []    # (family, name) that arrived and were named
        self.present = []   # (family, name) already in the destination
        self.reused = []    # (family, name) Revit matched to an existing type
        self.failed = []    # human-readable lines
        self.stranded = []  # sources left renamed - the serious one


def identity(symbol):
    """The (family, name, mark) that must survive a copy."""
    return (family_name(symbol), element_name(symbol), param_text(symbol, MARK))


def restore_identity(document, symbol, wanted):
    """Put a type's name and mark back. Returns True if it now matches.

    Called on the SOURCE after every copy, because CopyElements renames it.
    Also called on the arrival, which lands under Revit's internal default
    name rather than the one it had.
    """
    family, name, mark = wanted
    t = Transaction(document, "Restore connection name")
    t.Start()
    try:
        if element_name(symbol) != name:
            symbol.Name = name
        param = symbol.LookupParameter(MARK)
        if param is not None and not param.IsReadOnly:
            if (param.AsString() or u"") != mark:
                param.Set(mark)
        t.Commit()
    except Exception:
        t.RollBack()
        return False
    return element_name(symbol) == name


def transfer(source_doc, symbols, dest_doc):
    """Copy connection types into dest_doc, names intact. Returns Transferred.

    One type per CopyElements call on purpose. The call does not promise the
    order of the ids it hands back, and identity is the whole difficulty
    here: a copy arrives under a default name, so the only reliable way to
    know which source it came from is to have sent exactly one.

    Every pass restores the SOURCE as well as naming the arrival. Copying
    renames the original and clears its Type Mark, in the document it was
    copied from, with no warning - so a transfer that skipped this would
    quietly strip the names out of the model being read.
    """
    report = Transferred()
    options = CopyPasteOptions()
    options.SetDuplicateTypeNamesHandler(_UseDestination())

    for symbol in symbols:
        wanted = identity(symbol)
        family, name, _ = wanted

        if find_type(dest_doc, family, name) is not None:
            report.present.append((family, name))
            continue

        before = set(eid(i) for i in FilteredElementCollector(dest_doc)
                     .OfClass(StructuralConnectionHandlerType).ToElementIds())

        ids = List[type(symbol.Id)]()
        ids.Add(symbol.Id)

        made = None
        t = Transaction(dest_doc, "Import connection type")
        t.Start()
        try:
            arrived = ElementTransformUtils.CopyElements(
                source_doc, ids, dest_doc, None, options)
            t.Commit()
            fresh = [i for i in arrived if eid(i) not in before]
            made = dest_doc.GetElement(fresh[0]) if fresh else None
        except Exception as ex:
            t.RollBack()
            report.failed.append(u"{} : {} - {}".format(family, name, ex))

        # Before anything else, and whether or not the copy worked. The
        # source is damaged by the attempt, not by its success.
        if not restore_identity(source_doc, symbol, wanted):
            report.stranded.append((family, name))

        if made is None:
            if not report.failed or report.failed[-1].find(name) < 0:
                report.reused.append((family, name))
            continue

        if restore_identity(dest_doc, made, wanted):
            report.copied.append((family, name))
        else:
            report.failed.append(
                u"{} : {} - arrived but could not be named".format(
                    family, name))

    return report
