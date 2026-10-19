# -*- coding: utf-8 -*-
"""Set RVT Link Display Settings on many linked models at once.

Revit only lets you change one link at a time - Visibility/Graphics, Revit
Links, Display Settings, then back out again for the next one. This applies
the same change to every link you pick, across every view you pick.
"""

from pyrevit import revit, DB, forms, script

# ── [LIB] Snippets/_revisions.py ────────────────────────────────────────────
from Snippets._revisions import safe_str

doc = revit.doc
uidoc = revit.uidoc

TITLE = "Link Display"


# ── CONSTANTS ───────────────────────────────────────────────────────────────

# What the API will actually let us write. Basics (LinkVisibilityType) takes
# Custom; the per-row settings underneath it do not - ObjectStyles, ViewRange,
# ColorFill and NestedLinks all throw "Disallowed LinkVisibility value" on
# Custom, verified against Revit 2026. That is why the category tabs, and so
# the annotation checkbox, cannot be driven directly, and why COPY_ACTION
# exists: one link set up by hand, pushed onto the rest.
SET_CUSTOM = "Set Basics to Custom"
COPY_ACTION = "Copy every setting from one link to the others"
RESET_ACTION = "Reset to By host view"

ACTIONS = [SET_CUSTOM, COPY_ACTION, RESET_ACTION]

ACTIVE_ONLY = "Active view only"
PICK_VIEWS = "Choose views..."


# ── HELPERS ─────────────────────────────────────────────────────────────────

def _id_value(element_id):
    """An ElementId's integer, whichever accessor this Revit version has.

    ElementId went 64-bit at 2024, moving IntegerValue to Value. Seed43 still
    targets back to 2022, so ask for the new one and fall back.
    """
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue


def link_name(link):
    """The linked file's name, rather than the instance's long internal one.

    RevitLinkInstance.Name reads "file.rvt : 30 : location <Not Shared>",
    which is unreadable in a list. The type name is just the file.

    Run through safe_str because a linked file name is a .NET string that can
    carry non-ASCII, and a bare str() on one throws under IronPython 2.
    """
    try:
        return safe_str(doc.GetElement(link.GetTypeId()).Name)
    except Exception:
        return safe_str(link.Name)


def build_labels(links):
    """Unique display label -> link instance, sorted by name.

    A model linked in more than once shares one type name, so repeats get a
    counter - otherwise the second instance silently overwrites the first in
    the lookup and the user picks a link they cannot see.
    """
    totals = {}
    for link in links:
        name = link_name(link)
        totals[name] = totals.get(name, 0) + 1

    labels = {}
    seen = {}
    for link in sorted(links, key=lambda l: link_name(l).lower()):
        name = link_name(link)
        if totals[name] > 1:
            seen[name] = seen.get(name, 0) + 1
            labels["{}  ({})".format(name, seen[name])] = link
        else:
            labels[name] = link
    return labels


# ── COLLECTION ──────────────────────────────────────────────────────────────

def preselected_links():
    """Link instances in the current selection, if any."""
    found = []
    for element_id in uidoc.Selection.GetElementIds():
        element = doc.GetElement(element_id)
        if isinstance(element, DB.RevitLinkInstance):
            found.append(element)
    return found


def all_links():
    """Every RVT link instance in the model."""
    return list(DB.FilteredElementCollector(doc)
                .OfClass(DB.RevitLinkInstance)
                .WhereElementIsNotElementType())


def graphics_owner(view):
    """The element whose link overrides actually drive this view.

    A view template that controls V/G RVT Links wins over the view itself, and
    Revit discards the write to the view without complaining. Redirecting to
    the template is the only way the change sticks - the same thing CAD Layer
    Manager does.
    """
    if _id_value(view.ViewTemplateId) == _id_value(DB.ElementId.InvalidElementId):
        return view

    template = doc.GetElement(view.ViewTemplateId)
    if template is None:
        return view

    links_param = _id_value(
        DB.ElementId(DB.BuiltInParameter.VIS_GRAPHICS_RVT_LINKS))
    try:
        for param_id in template.GetNonControlledTemplateParameterIds():
            if _id_value(param_id) == links_param:
                return view  # template leaves links alone, write to the view
    except Exception:
        pass  # older API or an odd template - the template still owns V/G
    return template


# ── SETTINGS ────────────────────────────────────────────────────────────────

def settings_for(view, link):
    """This link's current override settings in the view, or a fresh set.

    GetLinkOverrides returns None for a link that has never been customised,
    and SetLinkOverrides will not take None back.
    """
    existing = view.GetLinkOverrides(link.Id)
    if existing is not None:
        return existing
    return DB.RevitLinkGraphicsSettings()


def apply_custom(view, link):
    """Put the link's Basics onto Custom, leaving every other row alone."""
    settings = settings_for(view, link)
    settings.LinkVisibilityType = DB.LinkVisibility.Custom
    view.SetLinkOverrides(link.Id, settings)


def apply_copy(view, link, source_settings):
    """Write another link's whole settings object onto this one."""
    view.SetLinkOverrides(link.Id, source_settings)


def apply_reset(view, link):
    """Drop the override entirely, back to By host view."""
    view.RemoveLinkOverrides(link.Id)


# ── PROMPTS ─────────────────────────────────────────────────────────────────

def choose_links():
    """The links to change - the selection if there is one, else a list."""
    selected = preselected_links()
    if selected:
        return selected

    links = all_links()
    if not links:
        forms.alert("This model has no RVT links.", title=TITLE)
        script.exit()

    labels = build_labels(links)
    picked = forms.SelectFromList.show(
        sorted(labels.keys()),
        title="{} - which links?".format(TITLE),
        button_name="Use these links",
        multiselect=True,
    )
    if not picked:
        script.exit()  # cancelled
    return [labels[name] for name in picked]


def choose_action():
    """What to do to the chosen links."""
    action = forms.SelectFromList.show(
        ACTIONS,
        title="{} - do what?".format(TITLE),
        button_name="Apply",
        multiselect=False,
    )
    if not action:
        script.exit()
    return action


def choose_source():
    """The link whose settings get copied onto the rest.

    Read from the active view, since that is the one the user just had open
    and set up by hand.
    """
    labels = build_labels(all_links())
    picked = forms.SelectFromList.show(
        sorted(labels.keys()),
        title="Copy settings FROM which link?",
        button_name="Copy from this link",
        multiselect=False,
    )
    if not picked:
        script.exit()

    source = labels[picked]
    settings = doc.ActiveView.GetLinkOverrides(source.Id)
    if settings is None:
        forms.alert(
            "{} has no display settings of its own in this view - it is still "
            "on By host view, so there is nothing to copy.\n\n"
            "Set that link up by hand first, then run this again."
            .format(link_name(source)), title=TITLE)
        script.exit()
    return source, settings


def choose_views():
    """Which views to write to."""
    scope = forms.SelectFromList.show(
        [ACTIVE_ONLY, PICK_VIEWS],
        title="{} - apply to which views?".format(TITLE),
        button_name="Continue",
        multiselect=False,
    )
    if not scope:
        script.exit()

    if scope == ACTIVE_ONLY:
        return [doc.ActiveView]

    views = forms.select_views(title="Apply link settings to these views",
                               multiple=True)
    if not views:
        script.exit()
    return views


# ── APPLY ───────────────────────────────────────────────────────────────────

def problem_note(link, view, err):
    """One line for the skipped list, which must never throw itself.

    A Revit exception message is a .NET string, so it can carry characters
    that blow up on a bare str() under IronPython 2 - and losing the whole
    report to that would hide the very thing it is meant to explain.
    """
    try:
        detail = safe_str("{}".format(err))
    except Exception:
        detail = "unknown error"
    return "{} in {}: {}".format(link_name(link), safe_str(view.Name), detail)


def run(action, links, views, source=None, source_settings=None):
    """Do the work, returning (changed count, list of failure notes).

    Failures are collected rather than raised: a schedule or a legend in the
    view list cannot hold link overrides at all, and one of those should not
    abandon the other forty views.
    """
    changed = 0
    problems = []

    with revit.Transaction("Set link display settings"):
        for view in views:
            owner = graphics_owner(view)
            for link in links:
                if source is not None and _id_value(link.Id) == _id_value(source.Id):
                    continue  # copying a link onto itself does nothing
                try:
                    if action == SET_CUSTOM:
                        apply_custom(owner, link)
                    elif action == COPY_ACTION:
                        apply_copy(owner, link, source_settings)
                    else:
                        apply_reset(owner, link)
                    changed += 1
                except Exception as err:
                    problems.append(problem_note(link, view, err))
    return changed, problems


def report(action, changed, problems, views):
    """Tell the user what happened, and what Revit will not let us do."""
    lines = ["{} link/view changes applied across {} view{}.".format(
        changed, len(views), "s" if len(views) != 1 else "")]

    if action == SET_CUSTOM:
        # Worth saying every time. Basics is all the API exposes, so the user
        # will open the dialog expecting the annotation tick to have moved
        # too, and it will not have.
        lines.append(
            "\nBasics is now Custom. Revit's API will not set the Annotation "
            "Categories tab, so untick 'Show annotation categories in this "
            "view' by hand on ONE link, then run this again and choose "
            "'{}' to push that link's settings onto the rest.".format(
                COPY_ACTION))

    if problems:
        lines.append("\nSkipped {}:".format(len(problems)))
        lines.extend(problems[:10])
        if len(problems) > 10:
            lines.append("...and {} more.".format(len(problems) - 10))

    forms.alert("\n".join(lines), title=TITLE)


# ── MAIN ────────────────────────────────────────────────────────────────────

def main():
    links = choose_links()
    action = choose_action()

    source = None
    source_settings = None
    if action == COPY_ACTION:
        source, source_settings = choose_source()

    views = choose_views()
    changed, problems = run(action, links, views, source, source_settings)
    report(action, changed, problems, views)


if __name__ == "__main__":
    main()
