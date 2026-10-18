# -*- coding: utf-8 -*-
# """
# The open document's own parameters: read, stage, apply.
#
# The Parameters tab of pyRameters. A project's parameters are its
# ParameterBindings; a family's are its FamilyManager parameters. Both are
# read into plain DocParam records, edited in the grid as staged changes,
# and written in one transaction on Apply - the model is never touched
# while the user is still deciding.
#
# ONE CODE PATH FOR REVIT 2022-2026
#     Every call uses the ForgeTypeId overloads (GroupTypeId, SpecTypeId).
#     They exist from 2022, and the BuiltInParameterGroup overloads they
#     replace were removed in 2025, so there is nothing to branch on.
#     Records hold the TypeId strings rather than ForgeTypeId objects, so
#     they compare and copy as plain values.
#
# WHAT THE API CANNOT DO, SO THE TOOL DOES NOT OFFER IT
#     A non-shared project parameter cannot be created through the API in
#     any version, so a project's new parameters come from a shared file.
#     A shared parameter's name belongs to its definition and cannot be
#     renamed. A built-in parameter cannot be changed at all.
#
# The record and staging half imports nothing from Revit, so it can be
# tested under desktop CPython. Target: IronPython 2 (pyRevit).
# """

try:
    from Autodesk.Revit.DB import (ForgeTypeId, InstanceBinding, LabelUtils,
                                   ParameterUtils, SharedParameterElement,
                                   SpecTypeId, SpecUtils, StorageType,
                                   Transaction, UnitUtils)
except ImportError:     # desktop CPython, for the tests
    pass

try:
    unicode
except NameError:       # CPython 3
    unicode = str

__all__ = [
    "DocParam", "NEW", "CHANGED", "DELETE", "EXISTING",
    "KIND_SHARED", "KIND_PROJECT", "KIND_FAMILY", "KIND_BUILTIN",
    "read_parameters", "group_options", "spec_options", "bindable_categories",
    "group_label", "spec_label", "discipline_label", "apply_changes",
    "text_spec",
]


# ── RECORDS AND STAGING ─────────────────────────────────────────────────────

NEW = u"new"
CHANGED = u"changed"
DELETE = u"delete"
EXISTING = u""

KIND_SHARED = u"Shared"
KIND_PROJECT = u"Project"
KIND_FAMILY = u"Family"
KIND_BUILTIN = u"Built-in"


class DocParam(object):
    """One parameter as read, plus whatever the user has changed.

    The values as read are kept beside the edited ones, so a row knows
    whether it has really changed - edit a value and put it back, and the
    row is no longer marked - and Revert can restore it.
    """

    FIELDS = ("name", "group", "spec", "is_instance", "categories",
              "formula", "value")

    def __init__(self, kind, name, group=u"", spec=u"", is_instance=True,
                 categories=(), formula=u"", value=u"", guid=u"",
                 handle=None, storage=None, reporting=False, state=EXISTING):
        self.kind = kind
        self.guid = guid
        self.handle = handle            # InternalDefinition or FamilyParameter
        self.storage = storage
        self.reporting = reporting
        self.state = state
        self.convert_to = None          # family: KIND_SHARED or KIND_FAMILY
        self.convert_guid = u""         # the shared definition to convert to
        self.name = name
        self.group = group or u""
        self.spec = spec or u""
        self.is_instance = bool(is_instance)
        self.categories = tuple(sorted(categories))
        self.formula = formula or u""
        self.value = value or u""
        self.original = dict((f, getattr(self, f)) for f in self.FIELDS)

    def changed_fields(self):
        return [f for f in self.FIELDS if getattr(self, f) != self.original[f]]

    def refresh_state(self):
        """Recompute CHANGED/EXISTING. NEW and DELETE are set explicitly."""
        if self.state in (NEW, DELETE):
            return
        changed = self.changed_fields() or self.convert_to
        self.state = CHANGED if changed else EXISTING

    def revert(self):
        for f in self.FIELDS:
            setattr(self, f, self.original[f])
        self.convert_to = None
        self.convert_guid = u""
        if self.state != NEW:
            self.state = EXISTING

    # --- what may be edited, and why not ---
    # Each returns None when the edit is allowed, or the reason it is not,
    # so the grid can say why a cell refused instead of just ignoring it.

    def why_not_rename(self, is_family):
        if self.kind == KIND_BUILTIN:
            return u"Built-in parameters cannot be renamed."
        if self.kind == KIND_SHARED:
            return (u"A shared parameter's name belongs to its definition in "
                    u"the shared file, so it cannot be renamed here.")
        if not is_family:
            return u"Revit cannot rename a project parameter through its API."
        return None

    def why_not_type(self):
        if self.state != NEW or self.kind != KIND_FAMILY:
            return (u"The data type can only be chosen for a new family "
                    u"parameter. Revit fixes it once a parameter exists.")
        return None

    def why_not_change(self):
        if self.kind == KIND_BUILTIN:
            return u"Built-in parameters cannot be changed."
        if self.state == DELETE:
            return u"This parameter is marked for deletion."
        return None

    def why_not_formula(self, is_family):
        if not is_family:
            return u"Formulas belong to family parameters."
        if self.kind == KIND_BUILTIN:
            return u"Built-in parameters cannot take a formula here."
        if self.reporting:
            return u"A reporting parameter takes its value from geometry."
        return self.why_not_change()

    def why_not_value(self, is_family):
        if not is_family:
            return (u"A project parameter's value belongs to each element, "
                    u"not to the parameter.")
        if self.formula:
            return u"The value comes from the formula."
        if self.storage is not None and unicode(self.storage) == u"ElementId":
            return u"This value is a reference to an element; set it in Revit."
        if self.state == DELETE:
            return u"This parameter is marked for deletion."
        return None

    def why_not_categories(self, is_family):
        if is_family:
            return u"Categories belong to project parameters."
        return self.why_not_change()


# ── READING ─────────────────────────────────────────────────────────────────

def eid(element_id):
    """The integer behind an ElementId: Value from 2024, IntegerValue before."""
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue


def _tid(forge_id):
    try:
        return forge_id.TypeId or u""
    except Exception:
        return u""


def read_parameters(doc):
    """Every parameter of the open document, as DocParam records."""
    rows = _read_family(doc) if doc.IsFamilyDocument else _read_project(doc)
    return sorted(rows, key=lambda p: p.name.lower())


def _read_project(doc):
    rows = []
    it = doc.ParameterBindings.ForwardIterator()
    it.Reset()
    while it.MoveNext():
        definition = it.Key           # InternalDefinition at run time
        binding = it.Current
        element = doc.GetElement(definition.Id)
        shared = isinstance(element, SharedParameterElement)
        categories = []
        for category in binding.Categories:
            categories.append(eid(category.Id))
        rows.append(DocParam(
            KIND_SHARED if shared else KIND_PROJECT, definition.Name,
            group=_tid(definition.GetGroupTypeId()),
            spec=_tid(definition.GetDataType()),
            is_instance=isinstance(binding, InstanceBinding),
            categories=categories,
            guid=unicode(element.GuidValue) if shared else u"",
            handle=definition))
    return rows


def _read_family(doc):
    rows = []
    fm = doc.FamilyManager
    current = fm.CurrentType
    for fp in fm.GetParameters():
        definition = fp.Definition
        if eid(fp.Id) < 0:
            kind = KIND_BUILTIN
        elif fp.IsShared:
            kind = KIND_SHARED
        else:
            kind = KIND_FAMILY
        rows.append(DocParam(
            kind, definition.Name,
            group=_tid(definition.GetGroupTypeId()),
            spec=_tid(definition.GetDataType()),
            is_instance=fp.IsInstance,
            formula=fp.Formula or u"",
            value=_value_text(current, fp),
            guid=unicode(fp.GUID) if fp.IsShared else u"",
            handle=fp, storage=fp.StorageType, reporting=fp.IsReporting))
    return rows


def _value_text(family_type, fp):
    """The current type's value as Revit would show it, units included."""
    if family_type is None:
        return u""
    try:
        if not family_type.HasValue(fp):
            return u""
        if fp.StorageType == StorageType.String:
            return family_type.AsString(fp) or u""
        text = family_type.AsValueString(fp)
        if text is not None:
            return text
        if fp.StorageType == StorageType.Integer:
            return unicode(family_type.AsInteger(fp))
        if fp.StorageType == StorageType.Double:
            return unicode(family_type.AsDouble(fp))
        if fp.StorageType == StorageType.ElementId:
            return unicode(eid(family_type.AsElementId(fp)))
    except Exception:
        pass
    return u""


# ── LABELS AND CHOICES ──────────────────────────────────────────────────────

def group_label(type_id):
    if not type_id:
        return u"Other"
    try:
        return LabelUtils.GetLabelForGroup(ForgeTypeId(type_id))
    except Exception:
        return type_id


def spec_label(type_id):
    if not type_id:
        return u""
    try:
        return LabelUtils.GetLabelForSpec(ForgeTypeId(type_id))
    except Exception:
        return type_id


def discipline_label(type_id):
    """Revit's discipline heading for a data type; Common for unitless ones.

    UnitUtils.GetDiscipline only accepts measurable specs. Text, Yes/No,
    Integer, URL and the rest sit under Common in Revit's own dialog.
    """
    try:
        spec = ForgeTypeId(type_id)
        if UnitUtils.IsMeasurableSpec(spec):
            return LabelUtils.GetLabelForDiscipline(UnitUtils.GetDiscipline(spec))
    except Exception:
        pass
    return u"Common"


def text_spec():
    """The Text data type, the default for a new family parameter."""
    return _tid(SpecTypeId.String.Text)


def group_options(doc):
    """(label, type_id) for every group a parameter can be put under.

    A family only accepts the groups FamilyManager calls user-assignable.
    """
    fm = doc.FamilyManager if doc.IsFamilyDocument else None
    seen = {}
    for gid in ParameterUtils.GetAllBuiltInGroups():
        if fm is not None:
            try:
                if not fm.IsUserAssignableParameterGroup(gid):
                    continue
            except Exception:
                continue
        seen[_tid(gid)] = group_label(_tid(gid))
    seen.setdefault(u"", u"Other")
    return sorted(((label, tid) for tid, label in seen.items()),
                  key=lambda pair: pair[0].lower())


def spec_options():
    """(discipline, label, type_id) for every type a parameter can have."""
    out = []
    for spec in SpecUtils.GetAllSpecs():
        try:
            if not SpecUtils.IsValidDataType(spec):
                continue
        except Exception:
            continue
        tid = _tid(spec)
        out.append((discipline_label(tid), spec_label(tid), tid))
    return sorted(out, key=lambda o: (o[0].lower(), o[1].lower()))


def bindable_categories(doc):
    """(name, id, category type) for every category a parameter can bind to."""
    out = []
    for category in doc.Settings.Categories:
        try:
            if category.AllowsBoundParameters:
                out.append((category.Name, eid(category.Id),
                            unicode(category.CategoryType)))
        except Exception:
            pass
    return sorted(out, key=lambda c: c[0].lower())


# ── APPLYING ────────────────────────────────────────────────────────────────

class ApplyReport(object):
    def __init__(self):
        self.done = []
        self.failed = []        # (parameter name, reason)


def apply_changes(doc, rows, shared_file=None):
    """Write every staged change in one transaction. Returns an ApplyReport.

    Each change is attempted on its own, so one Revit refuses does not cost
    the rest; what failed is reported by name with Revit's reason. Shared
    definitions are read from shared_file by pointing Revit at it for the
    duration and pointing it back afterwards.
    """
    report = ApplyReport()
    pending = [r for r in rows if r.state != EXISTING]
    if not pending:
        return report

    app = doc.Application
    externals = {}
    wants_file = any(_needs_external(r) for r in pending)
    previous = app.SharedParametersFilename
    try:
        if wants_file:
            externals = _external_definitions(app, shared_file)
        t = Transaction(doc, u"pyRameters: apply parameter changes")
        t.Start()
        try:
            if doc.IsFamilyDocument:
                _apply_family(doc, pending, externals, report)
            else:
                _apply_project(doc, pending, externals, report)
            t.Commit()
        except Exception:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
            raise
    finally:
        if wants_file:
            app.SharedParametersFilename = previous
    return report


def _needs_external(row):
    return ((row.state == NEW and row.kind == KIND_SHARED)
            or row.convert_to == KIND_SHARED)


def _external_definitions(app, path):
    if not path:
        raise ValueError(u"Save the shared parameter file first: Revit reads "
                         u"the definitions from disk.")
    app.SharedParametersFilename = path
    definition_file = app.OpenSharedParameterFile()
    if definition_file is None:
        raise ValueError(u"Revit could not open {0} as a shared parameter "
                         u"file.".format(path))
    found = {}
    for group in definition_file.Groups:
        for definition in group.Definitions:
            found[unicode(definition.GUID).lower()] = definition
    return found


def _external_for(externals, guid):
    definition = externals.get((guid or u"").lower())
    if definition is None:
        raise ValueError(u"GUID {0} is not in the shared parameter file."
                         .format(guid))
    return definition


def _attempt(report, row, action, what):
    try:
        action()
        report.done.append(u"{0}: {1}".format(row.name, what))
        return True
    except Exception as ex:
        report.failed.append((row.name, u"{0} - {1}".format(what, ex)))
        return False


# --- projects ---

def _binding(doc, is_instance, category_ids):
    by_id = {}
    for category in doc.Settings.Categories:
        by_id[eid(category.Id)] = category
    category_set = doc.Application.Create.NewCategorySet()
    for cid in category_ids:
        if cid in by_id:
            category_set.Insert(by_id[cid])
    if category_set.IsEmpty:
        raise ValueError(u"no categories chosen")
    if is_instance:
        return doc.Application.Create.NewInstanceBinding(category_set)
    return doc.Application.Create.NewTypeBinding(category_set)


def _apply_project(doc, rows, externals, report):
    bindings = doc.ParameterBindings

    for row in [r for r in rows if r.state == DELETE]:
        def remove(row=row):
            if not bindings.Remove(row.handle):
                raise ValueError(u"Revit did not remove it")
        _attempt(report, row, remove, u"deleted")

    for row in [r for r in rows if r.state == CHANGED]:
        changed = row.changed_fields()
        if "is_instance" in changed or "categories" in changed:
            def rebind(row=row):
                binding = _binding(doc, row.is_instance, row.categories)
                if not bindings.ReInsert(row.handle, binding,
                                         ForgeTypeId(row.group)):
                    raise ValueError(u"Revit did not rebind it")
            _attempt(report, row, rebind, u"rebound")
        elif "group" in changed:
            _attempt(report, row,
                     lambda row=row: row.handle.SetGroupTypeId(
                         ForgeTypeId(row.group)),
                     u"moved to {0}".format(group_label(row.group)))

    for row in [r for r in rows if r.state == NEW]:
        def add(row=row):
            definition = _external_for(externals, row.guid)
            binding = _binding(doc, row.is_instance, row.categories)
            if not bindings.Insert(definition, binding, ForgeTypeId(row.group)):
                raise ValueError(u"Revit refused it - is it already in the "
                                 u"project?")
        _attempt(report, row, add, u"added")


# --- families ---

def _set_value(fm, row, fp):
    text = row.value or u""
    if fp.StorageType == StorageType.String:
        fm.Set(fp, text)
    elif row.spec == _tid(SpecTypeId.Boolean.YesNo):
        yes = text.strip().lower() in (u"yes", u"y", u"1", u"true")
        fm.Set(fp, 1 if yes else 0)
    elif fp.StorageType == StorageType.Integer:
        try:
            fm.Set(fp, int(text))
        except ValueError:
            fm.SetValueString(fp, text)
    else:
        fm.SetValueString(fp, text)     # units aware: "1200 mm", "45°"


def _apply_family(doc, rows, externals, report):
    fm = doc.FamilyManager

    for row in [r for r in rows if r.state == DELETE]:
        _attempt(report, row, lambda row=row: fm.RemoveParameter(row.handle),
                 u"deleted")

    changed = [r for r in rows if r.state == CHANGED]
    for row in [r for r in changed if r.convert_to]:
        def convert(row=row):
            group = ForgeTypeId(row.group)
            if row.convert_to == KIND_SHARED:
                definition = _external_for(externals, row.convert_guid)
                row.handle = fm.ReplaceParameter(row.handle, definition, group,
                                                 row.is_instance)
            else:
                row.handle = fm.ReplaceParameter(row.handle, row.name, group,
                                                 row.is_instance)
        _attempt(report, row, convert,
                 u"converted to a {0} parameter".format(row.convert_to.lower()))

    for row in changed:
        fields = row.changed_fields()
        fp = row.handle
        if "name" in fields and not row.convert_to:
            _attempt(report, row, lambda fp=fp, row=row:
                     fm.RenameParameter(fp, row.name), u"renamed")
        if "is_instance" in fields and not row.convert_to:
            _attempt(report, row, lambda fp=fp, row=row:
                     fm.MakeInstance(fp) if row.is_instance else fm.MakeType(fp),
                     u"made " + (u"instance" if row.is_instance else u"type"))
        if "group" in fields and not row.convert_to:
            _attempt(report, row, lambda fp=fp, row=row:
                     fp.Definition.SetGroupTypeId(ForgeTypeId(row.group)),
                     u"moved to {0}".format(group_label(row.group)))

    for row in [r for r in rows if r.state == NEW]:
        def add(row=row):
            group = ForgeTypeId(row.group)
            if row.kind == KIND_SHARED:
                definition = _external_for(externals, row.guid)
                row.handle = fm.AddParameter(definition, group, row.is_instance)
            else:
                row.handle = fm.AddParameter(row.name, group,
                                             ForgeTypeId(row.spec),
                                             row.is_instance)
        _attempt(report, row, add, u"added")

    # Formulas and values last: a formula may name a parameter added above.
    for row in [r for r in rows if r.state in (NEW, CHANGED) and r.handle]:
        fields = row.changed_fields() if row.state == CHANGED else (
            ["formula"] if row.formula else []) + (["value"] if row.value else [])
        fp = row.handle
        if "formula" in fields:
            _attempt(report, row, lambda fp=fp, row=row:
                     fm.SetFormula(fp, row.formula or None),
                     u"formula set" if row.formula else u"formula cleared")
        elif "value" in fields and not row.formula:
            if fm.CurrentType is None:
                report.failed.append((row.name, u"the family has no types, so "
                                                u"there is no value to set"))
            else:
                _attempt(report, row, lambda fp=fp, row=row:
                         _set_value(fm, row, fp), u"value set")
