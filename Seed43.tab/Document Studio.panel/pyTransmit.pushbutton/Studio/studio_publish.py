# -*- coding: utf-8 -*-
# studio_publish.py
#
# One reading of a pyTransmit Studio layout, shared by the Revit publishers.
#
# Studio's grid already IS the document: n_rows x n_cols of cells with free
# rectangular merges, mm row heights and column widths, one revision per
# column. Turning that into a Revit drafting view or a Revit schedule is
# almost entirely the same work - resolve the row plan, expand the repeating
# rows, work out each cell's text - and only the last step (draw a text note /
# set a schedule cell) differs. That shared part lives here so the two Revit
# writers cannot drift apart, in the same spirit as studio_rows.py.
#
# Pure Python: no WPF, no Revit. Imports studio_rows for the row plan, which
# is the SAME module the Studio canvas draws from, so what Studio shows and
# what Studio publishes are built from one description of the table.
#
# NOTE: Publish/script_create_excel_studio.py predates this module and still
# carries its own copy of static_value() / rev_column_map(). It is left alone
# deliberately - it is the writer everything else is checked against - but if
# a value ever needs changing, change it in both places until that script is
# moved onto this module.

import json

import studio_rows

__all__ = [
    'MM_PER_PT', 'SECTION_BODY', 'SECTION_REPEAT_TOP', 'SECTION_REPEAT_BOTTOM',
    'load_layout', 'is_studio_layout', 'StudioLayout',
]

MM_PER_PT = 0.352778

SECTION_BODY = 'body'
SECTION_REPEAT_TOP = 'repeat_top'
SECTION_REPEAT_BOTTOM = 'repeat_bottom'

DEFAULT_ROW_H_MM = 8.0
DEFAULT_COL_W_MM = 30.0

# Project-info blocks: type -> key in the live-data dict, and the label each
# one prints in front of its value when the block asks for one.
KV_FIELD = {
    'proj_org': 'proj_org', 'proj_client': 'proj_client',
    'proj_number': 'proj_number', 'proj_name': 'proj_name',
    'doc_type': 'doc_type', 'print_size': 'print_size',
}
KV_LABELS = {
    'proj_org': 'Organisation', 'proj_client': 'Client',
    'proj_number': 'Project No', 'proj_name': 'Project',
    'doc_type': 'Document Type', 'print_size': 'Print Size',
}
# Per-revision blocks: type -> key in one revisions[] entry.
SPINE_FIELD = {
    'spine_dates': 'date', 'spine_initials': 'initials',
    'spine_reason': 'reason', 'spine_method': 'method',
    'spine_doc_type': 'doc_format', 'spine_print_size': 'paper_size',
}

# Issue metadata chosen in the pyTransmit window overrides what the revision
# carries - see StudioLayout._apply_meta_override().
META_KEYS = {'issued by': 'initials', 'initials': 'initials',
             'reason for issue': 'reason', 'method of issue': 'method',
             'document format': 'doc_format', 'paper size': 'paper_size'}


def load_layout(path):
    """Read a Studio layout from disk.

    Raises ValueError when the file is a Layout Builder template instead -
    the two schemas share a folder and an extension but nothing else, and a
    Layout Builder template read as a Studio one produces a blank document
    rather than an error.
    """
    with open(path, 'r') as f:
        layout = json.load(f)
    if not is_studio_layout(layout):
        raise ValueError('not a pyTransmit Studio layout')
    return layout


def is_studio_layout(layout):
    """Studio layouts have 'cells', Layout Builder ones have 'rows'."""
    return isinstance(layout, dict) and 'cells' in layout


class StudioLayout(object):
    """A Studio template joined to the live model data, ready to publish.

    Everything a writer needs is worked out once here: the grid, the row plan,
    the expansion of repeating rows into output rows, and the text of every
    cell. A writer then only has to place what placements() hands it.
    """

    # --- construction ---
    def __init__(self, layout, data, payload=None, log=None):
        """layout: a Studio layout dict (see load_layout).
        data:   studio_live_data.get_live_data() output.
        payload: pyTransmit's PYTRANSMIT_PAYLOAD, for the issue metadata and
                 the grouping toggles. None = publish the layout as saved.
        log:    callable taking one message, for the pyTransmit run log.
        """
        self._log = log or (lambda _m: None)
        self.layout = layout
        self.data = data
        self._p = payload or {}

        self.n_rows = int(layout.get('n_rows', 0))
        self.n_cols = int(layout.get('n_cols', 0))
        self.col_widths = list(layout.get('col_widths') or [])
        self.page_w_mm = float(layout.get('page_w_mm', 210))
        self.page_h_mm = float(layout.get('page_h_mm', 297))
        self.margin_mm = float(layout.get('margin_mm', 10))
        self.orientation = layout.get('orientation', 'portrait')
        self.page_size_name = layout.get('page_size_name', '')

        # Older layouts predate row_sections; pad so the list always matches
        # n_rows rather than trusting what is on disk.
        sections = list(layout.get('row_sections') or [])
        self.row_sections = (sections + [SECTION_BODY] * self.n_rows)[:self.n_rows]

        # cells is a flat list of merge origins; covered cells are implied by
        # the spans, so a lookup keyed on (r, c) is all a writer ever needs.
        self.cells = {}
        for entry in layout.get('cells', []) or []:
            self.cells[(entry.get('r'), entry.get('c'))] = entry

        self.logo_path = (layout.get('logo_path', '') or '').strip()
        if not self.logo_path:
            # A layout saved before Studio grew its logo picker has no logo of
            # its own; Branding's is better than printing nothing.
            self.logo_path = (self._p.get('logo_path', '') or '').strip()
            if self.logo_path:
                self._log('Layout has no logo of its own; using the Branding logo.')

        self.row_heights = self._normalised_row_heights(layout)
        self._apply_meta_override()
        self._build_row_plan()

    def _normalised_row_heights(self, layout):
        """Bring stacked-list row heights down to per-item height.

        The same correction the Studio window applies when it opens a layout,
        done again here because the window's version only reaches the file if
        the user saves - a template carried over from before rows were
        expanded still holds the old "tall enough for the whole list" height.
        """
        heights = list(layout.get('row_heights') or [])
        row_blocks = {}
        for (r, _c), entry in self.cells.items():
            row_blocks.setdefault(r, []).append(entry.get('block'))
        heights, rescaled = studio_rows.normalise_row_heights(heights, row_blocks)
        if rescaled:
            self._log('{} list row(s) resized to one row per item'.format(rescaled))
        return heights

    def _apply_meta_override(self):
        """Let the pyTransmit window's issue metadata win over the revision's.

        studio_live_data reads Reason / Method / Document Type / Page Size out
        of the revision's IssuedTo tags, and those are only written once a
        transmittal has been published - so on a first issue they are empty
        and those cells printed blank while the Layout Builder writers showed
        them. script_create_excel.py and script_create_excel_studio.py apply
        the same override.

        Worth knowing: these are the CURRENT issue's settings, so on a
        multi-revision transmittal they also overwrite the older columns.
        """
        vals = {}
        for label, value in (self._p.get('meta_rows') or []):
            key = META_KEYS.get(str(label).lower().strip())
            if key and str(value or '').strip():
                vals[key] = value
        if not vals:
            return
        for rev in (self.data.get('revisions') or []):
            rev.update(vals)
        self._log('Issue metadata from the pyTransmit window: {}'.format(
            ', '.join(sorted(vals.keys()))))

    def _build_row_plan(self):
        """Row plan, expanded rows and stripe indices - all from studio_rows,
        so the published document has exactly the rows the canvas showed."""
        group_params = self._p.get('group_params') or []
        # pyTransmit's Text On/Off toggle decides this, always. The layout is
        # for LAYOUT, not for which grouping to use or whether to name it, so
        # a 'group_label' key written by an earlier build is ignored.
        group_label = bool(self._p.get('group_label', True))

        # The layout stores a gap pair for each group-text state; the state in
        # force chooses which pair applies, so one template covers both.
        legacy_first = bool(self.layout.get('space_first_group', False))
        legacy_between = bool(self.layout.get('space_between_groups', False))
        suffix = 'on' if group_label else 'off'
        gap_first = bool(self.layout.get('space_first_' + suffix, legacy_first))
        gap_between = bool(self.layout.get('space_between_' + suffix, legacy_between))

        # condense=False always: Condense Rows shortens the PREVIEW so a
        # 2000-sheet layout stays workable on screen. The transmittal must
        # list every sheet.
        self.row_plan = studio_rows.sheet_row_plan(
            self.data, group_params, group_label, condense=False,
            space_first_group=gap_first, space_between_groups=gap_between)
        self.vrows, self.vspans = studio_rows.expand_rows(
            self.layout, self.data, self.row_plan)
        # Stripe index per row: restarts at every group so each group opens
        # white.
        self.band_index = studio_rows.band_indices(self.row_plan)

        # Only sheet-driven rows carry group and condensed-marker rows, so
        # only they need those rows treated specially.
        self._sheet_driven = set()
        for (r, _c), entry in self.cells.items():
            if studio_rows.repeat_domain(entry.get('block')) == 'sheet':
                self._sheet_driven.add(r)

        self.rev_col_map = self._rev_column_map()

        self._log('{} revisions, {} sheets, {} recipients -> {} output rows'.format(
            len(self.data.get('revisions') or []),
            len(self.data.get('docs') or []),
            len(self.data.get('distribution') or []),
            len(self.vrows)))

    # --- public: geometry ---
    def col_w(self, c):
        """Width of one grid column, in mm."""
        widths = self.col_widths
        return float(widths[c]) if c < len(widths) else DEFAULT_COL_W_MM

    def col_x(self, c):
        """Left edge of one grid column, in mm from the left of the table."""
        return sum(self.col_w(i) for i in range(c))

    def span_w(self, c, col_span):
        """Width of a run of columns, in mm."""
        return sum(self.col_w(c + i) for i in range(max(1, col_span)))

    def table_w(self):
        """Full table width, in mm."""
        return self.span_w(0, self.n_cols)

    def printable_w(self):
        """Page width inside the margins, in mm."""
        return max(10.0, self.page_w_mm - 2 * self.margin_mm)

    def printable_h(self):
        """Page height inside the margins, in mm."""
        return max(10.0, self.page_h_mm - 2 * self.margin_mm)

    def row_h(self, v):
        """Height of one OUTPUT row, in mm."""
        r = self.vrows[v][0]
        heights = self.row_heights
        return float(heights[r]) if r < len(heights) else DEFAULT_ROW_H_MM

    def section(self, v):
        """Which page section one OUTPUT row belongs to."""
        r = self.vrows[v][0]
        return self.row_sections[r] if r < len(self.row_sections) else SECTION_BODY

    def rows_in_section(self, name):
        """Output rows belonging to a page section, in order."""
        return [v for v in range(len(self.vrows)) if self.section(v) == name]

    def rows_height(self, vs):
        """Total height of a list of output rows, in mm."""
        return sum(self.row_h(v) for v in vs)

    def page_rows(self, printable_h_mm=None, split=True):
        """Output rows dealt out into printed pages.

        Rows in Studio's "Repeat at Top of Every Page" section open every
        page and rows in "Anchor to Bottom of Every Page" close it - the same
        two settings the Excel writer turns into print titles and a page
        footer. Everything else is dealt out in order until the page is full.

        Returns a list of row lists, always at least one. What a "page" then
        means is the writer's business: a fresh column across the drafting
        view, or a second schedule view.
        """
        top = self.rows_in_section(SECTION_REPEAT_TOP)
        bottom = self.rows_in_section(SECTION_REPEAT_BOTTOM)
        fixed = set(top) | set(bottom)
        body = [v for v in range(len(self.vrows)) if v not in fixed]

        if not split:
            return [top + body + bottom]

        height = float(printable_h_mm or self.printable_h())
        budget = height - self.rows_height(top) - self.rows_height(bottom)
        if budget <= 0:
            # The repeated bands alone are taller than the page. Splitting on
            # that basis would put one body row per page and run off the
            # paper, so keep it in one piece and say why.
            self._log('The repeating rows are taller than the printable page '
                      'height ({:.0f}mm) - the table is kept in one '
                      'piece.'.format(height))
            return [top + body + bottom]

        pages = []
        current = []
        used = 0.0
        for v in body:
            h = self.row_h(v)
            if current and used + h > budget:
                pages.append(top + current + bottom)
                current = []
                used = 0.0
            current.append(v)
            used += h
        pages.append(top + current + bottom)
        if len(pages) > 1:
            self._log('{} rows split across {} pages of {:.0f}mm printable '
                      'height.'.format(len(body), len(pages), height))
        return pages

    # --- public: content ---
    def rev_index_for(self, block, col):
        """Which revision a block in this column reads.

        Studio's rule, copied exactly: every column holding at least one
        revision block is a revision column, numbered left to right, so a
        block dropped one column over reads the next revision. A block can pin
        itself with an integer rev_index.
        """
        pinned = (block or {}).get('rev_index', 'auto')
        if isinstance(pinned, int):
            return max(0, pinned)
        return self.rev_col_map.get(col, 0)

    def static_value(self, block, col):
        """Text for a non-repeating block.

        Mirrors studio_blocks.render_block()'s value choices and
        script_create_excel_studio.py's static_value(); only the drawing
        differs between the three.
        """
        b = block or {}
        t = b.get('type', '')
        data = self.data
        revisions = data.get('revisions') or []

        if t == 'text':
            return b.get('content', '') or ''
        if t == 'blank':
            return ''
        if t in KV_FIELD:
            value = str(data.get(KV_FIELD[t], '') or '')
            if not value:
                return ''
            label = b.get('label') or KV_LABELS.get(t, '')
            return u'{}: {}'.format(label, value) if label else value
        if t == 'page_count':
            total = len(data.get('docs') or [])
            fmt = b.get('page_format', 'Page X of Y')
            value = {'Page X': 'Page 1',
                     'Page X of Y': 'Page 1 of {}'.format(total),
                     'X of Y': '1 of {}'.format(total),
                     'X / Y': '1 / {}'.format(total)}.get(fmt, str(total))
            return u' '.join([p for p in (b.get('prefix', ''), value,
                                          b.get('suffix', '')) if p])
        if t == 'issue_date':
            value = revisions[-1].get('date', '') if revisions else ''
            return u' '.join([p for p in (b.get('prefix', ''), value,
                                          b.get('suffix', '')) if p])
        if t in ('reason_list', 'method_list'):
            items = data.get('reasons' if t == 'reason_list' else 'methods') or []
            pairs = [u'{}  {}'.format(i.get('code', ''), i.get('label', ''))
                     for i in items]
            # 'row' spreads across one row, 'list' stacks - in a single cell
            # the difference is the separator.
            return (u'    '.join(pairs) if b.get('list_style') == 'row'
                    else u'\n'.join(pairs))
        if t.startswith('spine_'):
            idx = self.rev_index_for(b, col)
            if idx >= len(revisions):
                return ''
            return str(revisions[idx].get(SPINE_FIELD.get(t, 'rev'), '') or '')
        if t == 'logo':
            return ''   # placed as a picture, not as text
        return ''

    def placements(self):
        """Every cell to draw, already expanded over the row plan.

        One placement is one rectangle of the output document:

            r, c        the model cell it came from (for logging)
            v, n_v      first output row and how many it covers
            col_span    how many columns it covers
            block       the Studio block dict (may be None for an empty cell)
            text        the resolved text
            kind        'doc', 'group', 'more', 'space' or 'recipient'
            alt         True when this row takes the banding colour
            full_width  True for a group header, which spans the whole table

        Rows anchored to the bottom of the page are included: a writer that
        has somewhere to put them (a page footer) can pick them out by
        section(), and one that hasn't just draws them in place.
        """
        out = []
        for r in range(self.n_rows):
            for c in range(self.n_cols):
                entry = self.cells.get((r, c))
                if entry is None:
                    continue
                out.extend(self._placements_for(r, c, entry))
        return out

    # --- private helpers ---
    def _rev_column_map(self):
        """Grid column index -> revision index."""
        rev_cols = set()
        for (_r, c), entry in self.cells.items():
            block = entry.get('block') or {}
            if str(block.get('type', '')).startswith('spine_'):
                rev_cols.add(c)
        return dict((c, i) for i, c in enumerate(sorted(rev_cols)))

    def _vrow_range(self, r0, r1):
        """Model row range -> (first output row, count)."""
        first = self.vspans.get(r0, (r0, 1))[0]
        last_start, last_n = self.vspans.get(r1, (r1, 1))
        return first, (last_start + last_n) - first

    def _placements_for(self, r, c, entry):
        block = entry.get('block')
        row_span = int(entry.get('row_span', 1) or 1)
        col_span = int(entry.get('col_span', 1) or 1)
        first_v, n_v = self.vspans.get(r, (r, 1))
        domain = studio_rows.repeat_domain(block)

        # If the ROW repeats, every column in it repeats - not just the
        # columns holding a list block. A static cell beside the sheet list (a
        # label, or an empty column) is a cell on each of those rows, not one
        # tall box spanning them: that is what the printed table looks like,
        # and a tall merge here would also collide with the full-width group
        # header rows below.
        repeats = (row_span == 1 and self.vrows[first_v][1] is not None)
        if not repeats:
            v_first, v_count = self._vrow_range(r, r + row_span - 1)
            return [self._placement(r, c, v_first, v_count, col_span, block,
                                    self.static_value(block, c), 'doc')]

        rev_idx = self.rev_index_for(block, c)
        static_text = None if domain else self.static_value(block, c)
        out = []
        for i in range(n_v):
            v = first_v + i
            item = self.vrows[v][1]
            if domain:
                text, kind = studio_rows.repeat_cell_text(
                    block, self.data, self.row_plan, item, rev_idx)
            else:
                text = static_text
                kind = self.row_plan[item][0] if (
                    domain is None and item is not None
                    and 0 <= item < len(self.row_plan)
                    and r in self._sheet_driven) else 'doc'

            if kind in ('group', 'more'):
                # The group label spans the whole table and is written by the
                # Sheet Number column alone - every other column stays out of
                # the way, exactly as the Excel writers do.
                if (block or {}).get('type') != 'sheet_number':
                    continue
                out.append(self._placement(r, c, v, 1, max(1, self.n_cols - c),
                                           block, text, kind, full_width=True))
                continue

            # Band every other repeated row. The index is the position in the
            # row plan, so all the columns of one sheet band together and
            # group headers do not shift the pattern.
            bi = self.band_index[item] if (domain == 'sheet' and item is not None
                                           and item < len(self.band_index)) else item
            alt = (bool((block or {}).get('alt_rows'))
                   and bi is not None and bi % 2 == 1)
            out.append(self._placement(r, c, v, 1, col_span, block, text, kind,
                                       alt=alt))
        return out

    def _placement(self, r, c, v, n_v, col_span, block, text, kind,
                   alt=False, full_width=False):
        return {'r': r, 'c': c, 'v': v, 'n_v': n_v, 'col_span': col_span,
                'block': block, 'text': text, 'kind': kind, 'alt': alt,
                'full_width': full_width}
