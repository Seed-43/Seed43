# -*- coding: utf-8 -*-
# studio_rows.py
#
# The documentation table's row plan, shared by the Studio canvas and the
# Studio publisher. Pure Python - no WPF, no Revit - so the publish scripts
# can import it without dragging the whole renderer in.
#
# This module exists to stop the two from drifting. The canvas draws the
# plan, the publisher writes the plan; if grouping or condensing lived in
# both places separately, "what Studio shows" and "what Studio exports"
# would quietly stop matching, which is the one thing a WYSIWYG layout tool
# cannot afford.
#
# One documentation row is not always one sheet: turning on sheet grouping
# inserts a header row before each group, and Condense Rows collapses a long
# run of sheets into a single marker row. Sheet Number, Sheet Description,
# Drawing Group and the per-sheet Revision Marks columns all sit side by side
# in the same printed table, so they must agree on that row list exactly or
# the columns stop lining up.
#
# Entries are (kind, value):
#     ('doc',   doc_dict)   one sheet
#     ('group', label)      group header; label is '' when group text is off
#     ('more',  n)          n rows collapsed away by Condense Rows

from collections import OrderedDict

# ---------------------------------------------------------------------------
# Repeating blocks
# ---------------------------------------------------------------------------
# These blocks stand for a LIST, not a single value, and each item gets its
# own row - one row per sheet, one per recipient - on the canvas and in the
# exported document alike.
#
# Two independent domains, because they repeat over different things: a row
# holding Sheet Number repeats per sheet, a row holding Sent To repeats per
# recipient.
SHEET_ROW_TYPES = ('sheet_number', 'sheet_desc', 'drawing_group', 'spine_rev')
RECIPIENT_ROW_TYPES = ('sent_to', 'attn_to', 'spine_copies')
REPEAT_TYPES = SHEET_ROW_TYPES + RECIPIENT_ROW_TYPES

# Condense keeps the first and last few rows so the top and bottom of the
# table - headers, group starts, the last sheet - stay visible, and replaces
# everything between them with one marker row. 5 + 1 + 5 = 11 rows maximum,
# whether the model has 30 sheets or 2000.
CONDENSE_HEAD = 5
CONDENSE_TAIL = 5
CONDENSE_MAX_ROWS = CONDENSE_HEAD + 1 + CONDENSE_TAIL

# Em dash between the parameter values of a multi-parameter group label (the
# same separator the published document uses), and the midline ellipsis on the
# condensed marker row.
GROUP_LABEL_SEP = u' — '
CONDENSE_GLYPH = u'⋯'

# Row sizing. Worked in POINTS, because that is the unit both fonts and Excel
# row heights use - going via mm was what made repeated rows come out roughly
# twice as tall as they should be, with the text looking lost inside them.
MM_PER_PT = 0.352778
ROW_LINE_FACTOR = 1.3     # line height as a multiple of the font size
MIN_ROW_PT = 12.75        # comfortable floor; Excel's own default row is 15pt

# Kept for callers that still ask for it. No longer used to size rows.
ROW_PADDING_MM = 4.0

# Fallback when a block carries no size of its own: 9pt in mm, the Studio
# default. Kept here rather than imported from studio_blocks, so this module
# stays free of the renderer.
DEFAULT_SIZE_MM = 9.0 * MM_PER_PT


def repeat_domain(block):
    """'sheet', 'recipient' or None - what one row of this block means."""
    t = (block or {}).get('type', '')
    if t in SHEET_ROW_TYPES:
        return 'sheet'
    if t in RECIPIENT_ROW_TYPES:
        return 'recipient'
    return None


def natural_row_height_mm(block):
    """Height one repeated row wants, for this block's font size.

    A line of text plus a little breathing room - 12.75pt for 9pt Arial,
    against Excel's 15pt default row. The earlier version added a flat 4mm of
    padding borrowed from the canvas's pixel maths, which made a 9pt row 27pt
    tall: nearly double Excel's default, and the reason exported rows looked
    stretched with small text marooned in them.
    """
    size_mm = (block or {}).get('size_mm') or DEFAULT_SIZE_MM
    pt = float(size_mm) / MM_PER_PT
    return round(max(pt * ROW_LINE_FACTOR, MIN_ROW_PT) * MM_PER_PT, 2)


def normalise_row_heights(row_heights, row_blocks, factor=1.6):
    """Bring stacked-list row heights down to per-item height.

    Before rows were expanded, a row holding Sheet Number had to be tall
    enough for the WHOLE list inside one cell - 34mm in the shipped layouts.
    That number now means the height of ONE sheet's row, so left alone it
    produces 96pt rows in Excel and a sheet metres long.

    Applied by BOTH the window (on load) and the exporter (on read), because
    the window's fix only reaches the file if the user saves, and a template
    that has not been re-saved must still export correctly.

    row_blocks: {row_index: [block, ...]} - only the blocks actually in that
    row, in any order. Returns (heights, n_changed).
    """
    heights = list(row_heights)
    changed = 0
    for r, blocks in row_blocks.items():
        natural = 0.0
        for b in blocks:
            if repeat_domain(b) is None:
                continue
            natural = max(natural, natural_row_height_mm(b))
        if natural and 0 <= r < len(heights) and heights[r] > natural * factor:
            heights[r] = natural
            changed += 1
    return heights, changed


def _sheet_param(doc, name):
    return str((doc.get('params') or {}).get(name, '') or '').strip()


def group_label_for(doc, group_params):
    """Display label for a sheet's group - the non-blank parameter values
    joined by GROUP_LABEL_SEP, the same form
    Publish/script_create_schedule.py's get_group_label() builds."""
    parts = [_sheet_param(doc, pn) for pn in group_params]
    return GROUP_LABEL_SEP.join([p for p in parts if p])


def _group_key(doc, group_params):
    """Grouping identity. Blank values are kept (unlike the label) so two
    sheets that differ only by a missing value land in different groups -
    matching script_create_excel.py's _get_sheet_gl()."""
    return u'|'.join(_sheet_param(doc, pn) for pn in group_params)


def condense_plan(plan, condense=True):
    """Collapse the middle of a row plan, leaving CONDENSE_MAX_ROWS at most."""
    if not condense or len(plan) <= CONDENSE_MAX_ROWS:
        return plan
    hidden = len(plan) - CONDENSE_HEAD - CONDENSE_TAIL
    return plan[:CONDENSE_HEAD] + [('more', hidden)] + plan[-CONDENSE_TAIL:]


def sheet_row_plan(data, group_params=None, group_label=True, condense=False,
                   space_first_group=False, space_between_groups=False):
    """Build the row plan for the documentation table.

    Groups appear in first-appearance order, sheets stay in the order
    data['docs'] already has them (natural sheet-number sort), which is the
    same ordering the published document uses.

    group_label=False still inserts the header rows - they're what separates
    one group from the next on the page - but with no text in them, and drops
    the very first one, since a blank spacer above the first group would just
    be a gap at the top of the table. That is exactly what pyTransmit's own
    "Text Off" toggle does when publishing (script_create_excel.py's
    _first_grp_xl).
    """
    docs = data.get('docs', []) or []
    group_params = [p for p in (group_params or []) if p]

    if not group_params:
        return condense_plan([('doc', d) for d in docs], condense)

    groups = OrderedDict()
    for doc in docs:
        groups.setdefault(_group_key(doc, group_params), []).append(doc)

    plan = []
    first_group = True
    for _key, members in groups.items():
        label = group_label_for(members[0], group_params)
        # The gap is decided ONLY by its own switch - not by whether the group
        # is named, and not by whether group text is on. The two switches are
        # separate because the positions are different jobs: the FIRST group
        # sits directly under the table's own column headings, every LATER one
        # sits under the previous group's last sheet.
        if (space_first_group if first_group else space_between_groups):
            plan.append(('space', None))
        # With group text off there is no header row at all; the blank row
        # above IS the separation. That is what makes the two independent:
        #   text on,  gap on  -> gap + header
        #   text on,  gap off -> header only
        #   text off, gap on  -> one blank row
        #   text off, gap off -> nothing between the groups
        if group_label and label:
            plan.append(('group', label))
        for doc in members:
            plan.append(('doc', doc))
        first_group = False
    return condense_plan(plan, condense)


def band_indices(row_plan):
    """Which stripe each row of the plan gets, parallel to row_plan.

    A sheet row's index counts DOC rows only, and restarts at every group
    header, so:

      * the first sheet of the table is index 0 - unbanded, i.e. white;
      * the first sheet under each group header is index 0 again, so a group
        always opens white however many rows the group before it had.

    Counting raw plan positions instead (item % 2) banded the first sheet,
    and left the stripe pattern running straight through the group headers -
    so whether a group started white or grey depended on how many sheets
    happened to precede it.

    Group and marker rows get None: they have their own fill and take no
    part in the alternation.
    """
    out = []
    n = 0
    for kind, _value in row_plan:
        if kind == 'doc':
            out.append(n)
            n += 1
        else:
            out.append(None)
            n = 0      # next group starts white again
    return out


def plan_summary(data, group_params=None, group_label=True, condense=False):
    """(shown_rows, total_rows, n_groups) - for the status bar / ribbon."""
    full = sheet_row_plan(data, group_params, group_label, condense=False)
    shown = condense_plan(full, condense)
    n_groups = len([1 for kind, _v in full if kind == 'group'])
    return len(shown), len(full), n_groups


def repeat_cell_text(block, data, row_plan, item, rev_index=0):
    """Text for ONE repeated row, plus how that row should be banded.

    Returns (text, kind) where kind is 'doc', 'group', 'more' or 'recipient'.

    Which column carries a group label follows the published document
    exactly: the Sheet Number block writes it and everything beside it stays
    blank (script_create_excel.py writes the merged group row from the
    sheet_number block and explicitly skips it for sheet_desc).
    """
    t = (block or {}).get('type', '')

    if t in RECIPIENT_ROW_TYPES:
        dist = data.get('distribution', []) or []
        if not (0 <= item < len(dist)):
            return '', 'recipient'
        d = dist[item]
        if t == 'sent_to':
            return str(d.get('to', '') or ''), 'recipient'
        if t == 'attn_to':
            return str(d.get('attn', '') or ''), 'recipient'
        # Blank past the last issued revision, NOT the latest figure again.
        # Falling back to d['copies'] repeated the same count across every
        # empty revision column, so a transmittal with one revision printed
        # its copy numbers ten times across the page.
        per_rev = d.get('copies_by_rev') or []
        value = per_rev[rev_index] if rev_index < len(per_rev) else ''
        return str(value or ''), 'recipient'

    if not (0 <= item < len(row_plan)):
        return '', 'doc'
    kind, value = row_plan[item]

    if kind == 'space':
        return '', 'space'
    if kind == 'group':
        return (str(value or '') if t == 'sheet_number' else ''), 'group'
    if kind == 'more':
        return (u'{0}  {1:,} more rows  {0}'.format(CONDENSE_GLYPH, value)
                if t == 'sheet_number' else CONDENSE_GLYPH), 'more'

    if t == 'sheet_number':
        return str(value.get('sheet', '') or ''), 'doc'
    if t == 'sheet_desc':
        return str(value.get('desc', '') or ''), 'doc'
    if t == 'spine_rev':
        marks = value.get('revs', []) or []
        return str((marks[rev_index] if rev_index < len(marks) else '') or ''), 'doc'
    # drawing_group: the group headers carry the labels, so the sheet rows
    # under them are deliberately blank.
    return '', 'doc'


def expand_rows(grid_dict, data, row_plan):
    """Expand a Studio layout's model rows into the rows actually drawn.

    Returns (vrows, spans):
        vrows = [(model_row, item_or_None)]      one entry per output row
        spans = {model_row: (first_output_row, count)}

    The same expansion StudioSettings does for the canvas, so the exported
    document has exactly the rows the canvas showed. Kept here rather than in
    the window so the publisher does not have to reimplement it.
    """
    n_rows = int(grid_dict.get('n_rows', 0))
    n_cols = int(grid_dict.get('n_cols', 0))
    n_sheet = len(row_plan)
    n_recipient = len(data.get('distribution', []) or [])

    # cells arrive as a flat list of {r, c, row_span, col_span, block}
    by_rc = {}
    for entry in grid_dict.get('cells', []) or []:
        by_rc[(entry.get('r'), entry.get('c'))] = entry

    vrows, spans = [], {}
    for r in range(n_rows):
        domain = None
        for c in range(n_cols):
            entry = by_rc.get((r, c))
            if entry is None:
                continue
            d = repeat_domain(entry.get('block'))
            if d == 'sheet':
                # A row with both kinds in it is ambiguous; the documentation
                # table wins, since that is the one whose length varies.
                domain = 'sheet'
                break
            if d == 'recipient':
                domain = 'recipient'
        count = n_sheet if domain == 'sheet' else (
            n_recipient if domain == 'recipient' else 0)
        # A repeating block with nothing to show still gets one row, so an
        # empty list does not silently remove a row from the document.
        items = list(range(count)) if count else [None]
        spans[r] = (len(vrows), len(items))
        for it in items:
            vrows.append((r, it))
    return vrows, spans
