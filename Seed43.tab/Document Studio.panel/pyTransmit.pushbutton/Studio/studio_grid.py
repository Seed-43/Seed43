# -*- coding: utf-8 -*-
# studio_grid.py
#
# Free-grid data model for pyTransmit Studio: arbitrary rows/columns, free
# rectangular merge/unmerge (Excel-style), and JSON persistence. Pure Python,
# no WPF here - StudioSettings.py turns this model into UI.

import os
import json
import copy

from studio_blocks import DEFAULT_FONT, pt_to_mm

DEFAULT_ROW_H_MM = 8.0
DEFAULT_COL_W_MM = 30.0

# Base (portrait) page sizes in mm - same convention as Excel's Page Setup:
# pick a paper size + an orientation, the effective width/height is derived
# from the two rather than typed in directly.
PAGE_SIZES = {
    'A4':      (210, 297),
    'A3':      (297, 420),
    'A2':      (420, 594),
    'Letter':  (216, 279),
    'Tabloid': (279, 432),
}
DEFAULT_PAGE_SIZE = 'A4'
# Printable area = page minus a margin per side. Without a margin model
# "does my content fit?" is ambiguous: converted Layout Builder templates
# bake a 10mm side margin into their column widths, so 190mm of columns on
# a 210mm sheet is an exact fit, not 20mm of slack.
DEFAULT_MARGIN_MM = 10.0
MIN_COL_MM = 5.0        # narrowest a column may be squeezed to
MIN_AUTO_ROW_MM = 4.5   # floor for an auto-fitted row
# 'Custom' keeps whatever page_w_mm/page_h_mm are set rather than looking
# them up - the Revit Drafting View / Legend / Schedule templates use a
# 185x200 sheet that matches no standard paper size.
CUSTOM_PAGE_SIZE = 'Custom'

# Row sections - Excel's "Print Titles" model. A row is either ordinary
# body content, or part of a band that reappears on every printed page.
SECTION_BODY = 'body'
SECTION_REPEAT_TOP = 'repeat_top'        # Excel: "rows to repeat at top"
SECTION_REPEAT_BOTTOM = 'repeat_bottom'  # anchored to the foot of every page
# Full wording - used by the row-header menu, the formula bar and the
# palette legend so all three read identically. SECTION_SHORT is the
# compact form for tight spots (the status indicator, gutter tooltips).
SECTION_LABELS = {
    SECTION_BODY: 'Body (no repeat)',
    SECTION_REPEAT_TOP: 'Repeat at Top of Every Page',
    SECTION_REPEAT_BOTTOM: 'Anchor to Bottom of Every Page',
}
SECTION_SHORT = {
    SECTION_BODY: 'Body',
    SECTION_REPEAT_TOP: 'Repeats at top',
    SECTION_REPEAT_BOTTOM: 'Repeats at bottom',
}
SECTION_COLORS = {
    SECTION_REPEAT_TOP: '#2980B9',
    SECTION_REPEAT_BOTTOM: '#E67E22',
}


class Grid(object):
    """cells: dict {(r, c): cell}
    A plain cell:  {'block': dict|None}
    A merge origin: {'block': dict|None, 'row_span': int, 'col_span': int}
    A covered cell (inside someone else's merge): {'covered_by': (r, c)}
    Every (r, c) in range(n_rows) x range(n_cols) has an entry."""

    def __init__(self, n_rows=12, n_cols=6):
        self.n_rows = n_rows
        self.n_cols = n_cols
        self.row_heights = [DEFAULT_ROW_H_MM] * n_rows
        self.col_widths = [DEFAULT_COL_W_MM] * n_cols
        # Parallel to row_heights - kept in step by insert_row/delete_row.
        self.row_sections = [SECTION_BODY] * n_rows
        self.page_size_name = DEFAULT_PAGE_SIZE
        self.margin_mm = DEFAULT_MARGIN_MM
        self.orientation = 'portrait'
        self.page_w_mm = 210
        self.page_h_mm = 297
        self._recompute_page_size()
        self.logo_path = ''
        # Blank row before a sheet-group header.
        #
        # FOUR switches, not two: the gap you want depends on whether group
        # text is showing. With a named header band above each group a gap may
        # be unnecessary; with the text off the gap IS the only thing
        # separating one group from the next, so the same template usually
        # wants different answers. pyTransmit's own Text On/Off toggle chooses
        # which pair applies, so one template covers both.
        #
        #   *_first   : above the FIRST group, the one under the column headings
        #   *_between : above every LATER group header
        self.space_first_on = False
        self.space_between_on = False
        self.space_first_off = False
        self.space_between_off = False
        # Columns may not total more than the printable width. This is a
        # LIMIT: the column being changed is capped, and no other column is
        # touched. See clamp_to_width().
        self.lock_width = False
        # Per row: True = height follows the content, False = the user set it.
        # Parallel to row_heights, kept in step by insert/delete/move_rows.
        self.row_auto = [True] * n_rows
        # NOTE: group text on/off is deliberately NOT stored here. pyTransmit's
        # Text On/Off toggle owns that, and the layout owns the layout - the
        # gaps above, the colours, where things sit. A layout that also decided
        # whether to name its groups would silently override the toggle.
        self.cells = {}
        for r in range(n_rows):
            for c in range(n_cols):
                self.cells[(r, c)] = {'block': None}

    # -- page size / orientation ------------------------------------------------
    def _recompute_page_size(self):
        if self.page_size_name == CUSTOM_PAGE_SIZE:
            return   # explicit page_w_mm / page_h_mm win
        w, h = PAGE_SIZES.get(self.page_size_name, PAGE_SIZES[DEFAULT_PAGE_SIZE])
        if self.orientation == 'landscape':
            w, h = h, w
        self.page_w_mm, self.page_h_mm = w, h

    def printable_w_mm(self):
        return max(10.0, self.page_w_mm - 2 * self.margin_mm)

    def printable_h_mm(self):
        return max(10.0, self.page_h_mm - 2 * self.margin_mm)

    def fit_columns_to_width(self):
        """Scale every column by the same factor so the total exactly fills
        the printable width - relative proportions are preserved, which is
        what makes it a stretch rather than a redistribution."""
        total = float(sum(self.col_widths))
        if total <= 0:
            return
        factor = self.printable_w_mm() / total
        self.col_widths = [round(w * factor, 2) for w in self.col_widths]

    def clamp_to_width(self, changed):
        """Stop the columns exceeding the printable width, by capping the ones
        just changed. Every other column is left exactly as it was.

        This is a LIMIT, not a redistribution. An earlier version rebalanced
        all the other columns to keep the total pinned, which quietly undid
        widths the user had deliberately set: size A, B and C, then adjust the
        rest, and A, B and C moved too. Only what is being dragged gives way.

        `changed` is one column index or a list of them. Returns True if
        anything had to be capped.
        """
        if isinstance(changed, int):
            changed = [changed]
        changed = [i for i in changed if 0 <= i < len(self.col_widths)]
        if not changed:
            return False
        target = self.printable_w_mm()
        excess = round(sum(self.col_widths) - target, 4)
        if excess <= 0.005:
            return False

        # Take it off the changed columns evenly, but never below the minimum
        # - anything that cannot be absorbed is left over rather than pushed
        # onto columns the user did not touch.
        remaining = excess
        for _pass in range(2):      # second pass mops up what a floor blocked
            takers = [i for i in changed if self.col_widths[i] > MIN_COL_MM]
            if not takers or remaining <= 0.005:
                break
            share = remaining / float(len(takers))
            for i in takers:
                give = min(share, self.col_widths[i] - MIN_COL_MM)
                self.col_widths[i] = round(self.col_widths[i] - give, 2)
                remaining = round(remaining - give, 4)
        return True

    def auto_row_height_mm(self, idx, blocks):
        """Height row `idx` wants for its content - the tallest block's own
        line height. Single-line: wrapped text would need real text
        measurement, which the model has no access to."""
        natural = 0.0
        for b in blocks:
            if not b:
                continue
            size_mm = b.get('size_mm') or (9 * 0.352778)
            natural = max(natural, round(size_mm * 1.3 + 1.6, 2))
        return max(natural, MIN_AUTO_ROW_MM)

    def distribute_columns(self, first=None, last=None):
        """Equal widths across a column range (all columns when unspecified),
        keeping the range's combined width unchanged."""
        if first is None or last is None:
            first, last = 0, self.n_cols - 1
        first = max(0, first); last = min(self.n_cols - 1, last)
        n = last - first + 1
        if n < 1:
            return
        share = sum(self.col_widths[first:last + 1]) / float(n)
        for c in range(first, last + 1):
            self.col_widths[c] = round(share, 2)

    def set_page_size(self, name):
        if name == CUSTOM_PAGE_SIZE:
            self.page_size_name = name
            return
        if name in PAGE_SIZES:
            self.page_size_name = name
            self._recompute_page_size()

    def set_row_section(self, idx, section):
        if 0 <= idx < self.n_rows and section in SECTION_LABELS:
            self.row_sections[idx] = section

    def section_of(self, idx):
        try:
            return self.row_sections[idx]
        except Exception:
            return SECTION_BODY

    def gaps_for(self, group_label_on):
        """(first, between) for the given group-text state."""
        if group_label_on:
            return self.space_first_on, self.space_between_on
        return self.space_first_off, self.space_between_off

    def set_gap(self, group_label_on, which, value):
        """Set one gap switch for the given group-text state."""
        suffix = 'on' if group_label_on else 'off'
        setattr(self, 'space_{}_{}'.format(which, suffix), bool(value))

    def set_orientation(self, orientation):
        if orientation in ('portrait', 'landscape'):
            self.orientation = orientation
            self._recompute_page_size()

    # -- cell access -----------------------------------------------------------
    def cell(self, r, c):
        return self.cells.get((r, c))

    def origin_of(self, r, c):
        """Return (origin_r, origin_c) for whatever cell/merge occupies (r, c)."""
        cell = self.cells.get((r, c))
        if cell is None:
            return (r, c)
        cb = cell.get('covered_by')
        return cb if cb else (r, c)

    def span_of(self, r, c):
        """Return (row_span, col_span) of the merge/cell at its origin."""
        cell = self.cells.get((r, c), {})
        return (cell.get('row_span', 1), cell.get('col_span', 1))

    def block_at(self, r, c):
        orig = self.origin_of(r, c)
        return self.cells.get(orig, {}).get('block')

    def set_block(self, r, c, block):
        orig = self.origin_of(r, c)
        self.cells[orig]['block'] = block

    def clear_cell(self, r, c):
        orig = self.origin_of(r, c)
        self.cells[orig]['block'] = None

    # -- rows / columns ----------------------------------------------------------
    def insert_row(self, at, height_mm=DEFAULT_ROW_H_MM):
        at = max(0, min(at, self.n_rows))
        new_cells = {}
        for (r, c), cell in self.cells.items():
            nr = r + 1 if r >= at else r
            cb = cell.get('covered_by')
            if cb:
                cb = (cb[0] + 1 if cb[0] >= at else cb[0], cb[1])
                new_cells[(nr, c)] = {'covered_by': cb}
            else:
                new_cells[(nr, c)] = cell
        for c in range(self.n_cols):
            new_cells[(at, c)] = {'block': None}
        self.cells = new_cells
        self.row_heights.insert(at, height_mm)
        self.row_sections.insert(at, SECTION_BODY)
        self.row_auto.insert(at, True)
        self.n_rows += 1

    def delete_row(self, at):
        if self.n_rows <= 1 or not (0 <= at < self.n_rows):
            return
        new_cells = {}
        for (r, c), cell in self.cells.items():
            if r == at:
                continue
            nr = r - 1 if r > at else r
            cb = cell.get('covered_by')
            if cb:
                if cb[0] == at:
                    new_cells[(nr, c)] = {'block': None}  # merge origin removed, fall back to plain
                    continue
                cb = (cb[0] - 1 if cb[0] > at else cb[0], cb[1])
                new_cells[(nr, c)] = {'covered_by': cb}
            else:
                new_cells[(nr, c)] = cell
        self.cells = new_cells
        del self.row_heights[at]
        del self.row_sections[at]
        del self.row_auto[at]
        self.n_rows -= 1

    # -- reordering rows ---------------------------------------------------------
    def _swap_rows(self, a, b):
        """Exchange two rows outright - contents, height and section.

        covered_by holds absolute coordinates, so a horizontal merge's covered
        cells have to be re-pointed at their origin's new row; can_move_rows()
        has already ruled out vertical merges, whose origin and covered cells
        would otherwise be torn apart by the move.
        """
        remap = {a: b, b: a}
        new_cells = {}
        for (r, c), cell in self.cells.items():
            nr = remap.get(r, r)
            cb = cell.get('covered_by')
            if cb:
                new_cells[(nr, c)] = {'covered_by': (remap.get(cb[0], cb[0]), cb[1])}
            else:
                new_cells[(nr, c)] = cell
        self.cells = new_cells
        self.row_heights[a], self.row_heights[b] = self.row_heights[b], self.row_heights[a]
        self.row_sections[a], self.row_sections[b] = self.row_sections[b], self.row_sections[a]
        self.row_auto[a], self.row_auto[b] = self.row_auto[b], self.row_auto[a]

    def can_move_rows(self, r0, r1, delta):
        """Can rows r0..r1 swap places with the row above (delta -1) or below
        (delta +1)?

        No, if any row involved - the block or the row it would swap with -
        is part of a cell merged across rows. Half a merged cell cannot be
        somewhere else, and silently unmerging to allow the move would lose
        work the user did deliberately. Merges within a single row travel
        with it and are fine.
        """
        if delta not in (-1, 1):
            return False
        if not (0 <= r0 <= r1 < self.n_rows):
            return False
        target = r0 - 1 if delta < 0 else r1 + 1
        if not (0 <= target < self.n_rows):
            return False
        for r in range(min(r0, target), max(r1, target) + 1):
            for c in range(self.n_cols):
                orig = self.origin_of(r, c)
                if self.span_of(orig[0], orig[1])[0] > 1:
                    return False
        return True

    def move_rows(self, r0, r1, delta):
        """Move rows r0..r1 one place up or down. Returns False (changing
        nothing) when can_move_rows() says no.

        Walked as a run of adjacent swaps, in the order that keeps the block
        together: upwards from the top row, downwards from the bottom one.
        """
        if not self.can_move_rows(r0, r1, delta):
            return False
        if delta < 0:
            for r in range(r0, r1 + 1):
                self._swap_rows(r, r - 1)
        else:
            for r in range(r1, r0 - 1, -1):
                self._swap_rows(r, r + 1)
        return True

    def insert_col(self, at, width_mm=DEFAULT_COL_W_MM):
        at = max(0, min(at, self.n_cols))
        new_cells = {}
        for (r, c), cell in self.cells.items():
            nc = c + 1 if c >= at else c
            cb = cell.get('covered_by')
            if cb:
                cb = (cb[0], cb[1] + 1 if cb[1] >= at else cb[1])
                new_cells[(r, nc)] = {'covered_by': cb}
            else:
                new_cells[(r, nc)] = cell
        for r in range(self.n_rows):
            new_cells[(r, at)] = {'block': None}
        self.cells = new_cells
        self.col_widths.insert(at, width_mm)
        self.n_cols += 1

    def delete_col(self, at):
        if self.n_cols <= 1 or not (0 <= at < self.n_cols):
            return
        new_cells = {}
        for (r, c), cell in self.cells.items():
            if c == at:
                continue
            nc = c - 1 if c > at else c
            cb = cell.get('covered_by')
            if cb:
                if cb[1] == at:
                    new_cells[(r, nc)] = {'block': None}
                    continue
                cb = (cb[0], cb[1] - 1 if cb[1] > at else cb[1])
                new_cells[(r, nc)] = {'covered_by': cb}
            else:
                new_cells[(r, nc)] = cell
        self.cells = new_cells
        del self.col_widths[at]
        self.n_cols -= 1

    # -- merge / unmerge ---------------------------------------------------------
    def can_merge(self, r0, c0, r1, c1):
        """Simple rule: merge only allowed if every cell in the rectangle is
        currently a plain, unmerged 1x1 cell (no existing merge overlaps)."""
        r0, r1 = min(r0, r1), max(r0, r1)
        c0, c1 = min(c0, c1), max(c0, c1)
        if r0 == r1 and c0 == c1:
            return False  # nothing to merge
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                cell = self.cells.get((r, c))
                if cell is None:
                    return False
                if 'covered_by' in cell:
                    return False
                if cell.get('row_span', 1) > 1 or cell.get('col_span', 1) > 1:
                    return False
        return True

    def merge(self, r0, c0, r1, c1):
        """Merge the rectangle; raises ValueError if can_merge() is False."""
        if not self.can_merge(r0, c0, r1, c1):
            raise ValueError('Selection overlaps an existing merged cell - unmerge it first.')
        r0, r1 = min(r0, r1), max(r0, r1)
        c0, c1 = min(c0, c1), max(c0, c1)
        origin_block = self.cells[(r0, c0)].get('block')
        # Prefer a non-empty block from anywhere in the selection as the merged content
        if origin_block is None:
            for r in range(r0, r1 + 1):
                for c in range(c0, c1 + 1):
                    b = self.cells[(r, c)].get('block')
                    if b is not None:
                        origin_block = b
                        break
                if origin_block is not None:
                    break
        self.cells[(r0, c0)] = {'block': origin_block, 'row_span': r1 - r0 + 1, 'col_span': c1 - c0 + 1}
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                if (r, c) == (r0, c0):
                    continue
                self.cells[(r, c)] = {'covered_by': (r0, c0)}

    def merge_across(self, r0, c0, r1, c1):
        """Merge each ROW of the rectangle on its own, Excel's "Merge Across".

        Four rows of six columns become four merged cells, not one - which is
        what you want for a stack of full-width labels. Rows that cannot be
        merged (they already contain a merge, or are a single column) are
        skipped rather than aborting the lot, so one awkward row does not stop
        the rest. Returns how many rows were merged.
        """
        r0, r1 = min(r0, r1), max(r0, r1)
        c0, c1 = min(c0, c1), max(c0, c1)
        merged = 0
        for r in range(r0, r1 + 1):
            if self.can_merge(r, c0, r, c1):
                self.merge(r, c0, r, c1)
                merged += 1
        return merged

    def unmerge_all(self, r0, c0, r1, c1):
        """Unmerge every merge the rectangle touches. Returns the count."""
        r0, r1 = min(r0, r1), max(r0, r1)
        c0, c1 = min(c0, c1), max(c0, c1)
        origins = []
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                orig = self.origin_of(r, c)
                if orig in origins:
                    continue
                rs, cs = self.span_of(orig[0], orig[1])
                if rs > 1 or cs > 1:
                    origins.append(orig)
        for (mr, mc) in origins:
            self.unmerge(mr, mc)
        return len(origins)

    def unmerge(self, r, c):
        orig = self.origin_of(r, c)
        cell = self.cells.get(orig)
        if not cell or ('row_span' not in cell and 'col_span' not in cell):
            return
        row_span = cell.get('row_span', 1)
        col_span = cell.get('col_span', 1)
        block = cell.get('block')
        for rr in range(orig[0], orig[0] + row_span):
            for cc in range(orig[1], orig[1] + col_span):
                self.cells[(rr, cc)] = {'block': block if (rr, cc) == orig else None}

    # -- persistence ---------------------------------------------------------
    def to_dict(self, name='Untitled'):
        cell_list = []
        for (r, c), cell in self.cells.items():
            if 'covered_by' in cell:
                continue
            cell_list.append({
                'r': r, 'c': c,
                'row_span': cell.get('row_span', 1),
                'col_span': cell.get('col_span', 1),
                'block': cell.get('block'),
            })
        return {
            'name': name,
            'n_rows': self.n_rows, 'n_cols': self.n_cols,
            'row_heights': self.row_heights, 'col_widths': self.col_widths,
            'row_sections': self.row_sections,
            'row_auto': self.row_auto,
            'lock_width': self.lock_width,
            'page_size_name': self.page_size_name, 'orientation': self.orientation,
            'margin_mm': self.margin_mm,
            'page_w_mm': self.page_w_mm, 'page_h_mm': self.page_h_mm,
            'logo_path': self.logo_path,
            'space_first_on': self.space_first_on,
            'space_between_on': self.space_between_on,
            'space_first_off': self.space_first_off,
            'space_between_off': self.space_between_off,
            'cells': cell_list,
        }

    @classmethod
    def from_dict(cls, d):
        g = cls(n_rows=int(d.get('n_rows', 12)), n_cols=int(d.get('n_cols', 6)))
        g.row_heights = list(d.get('row_heights') or ([DEFAULT_ROW_H_MM] * g.n_rows))
        g.col_widths = list(d.get('col_widths') or ([DEFAULT_COL_W_MM] * g.n_cols))
        # Older saved layouts predate row_sections - pad/truncate so the
        # list always matches n_rows rather than trusting what's on disk.
        secs = list(d.get('row_sections') or [])
        secs = (secs + [SECTION_BODY] * g.n_rows)[:g.n_rows]
        g.row_sections = [s if s in SECTION_LABELS else SECTION_BODY for s in secs]
        auto = list(d.get('row_auto') or [])
        # Older layouts have no row_auto. Their heights were all set by
        # hand, so they are adopted as manual - auto-fitting them on load
        # would silently redesign someone's template.
        g.row_auto = (auto + [False] * g.n_rows)[:g.n_rows]
        g.lock_width = bool(d.get('lock_width', False))
        g.margin_mm = float(d.get('margin_mm', DEFAULT_MARGIN_MM))
        g.page_size_name = d.get('page_size_name', DEFAULT_PAGE_SIZE)
        g.orientation = d.get('orientation', 'portrait')
        g.page_w_mm = d.get('page_w_mm', g.page_w_mm)
        g.page_h_mm = d.get('page_h_mm', g.page_h_mm)
        g._recompute_page_size()
        g.logo_path = d.get('logo_path', '')
        # Layouts saved with the earlier single pair carry that setting
        # into BOTH states, so they keep the spacing they were designed
        # with whichever way the text toggle is set.
        _legacy_first = bool(d.get('space_first_group', False))
        _legacy_between = bool(d.get('space_between_groups', False))
        g.space_first_on = bool(d.get('space_first_on', _legacy_first))
        g.space_between_on = bool(d.get('space_between_on', _legacy_between))
        g.space_first_off = bool(d.get('space_first_off', _legacy_first))
        g.space_between_off = bool(d.get('space_between_off', _legacy_between))
        g.cells = {}
        for r in range(g.n_rows):
            for c in range(g.n_cols):
                g.cells[(r, c)] = {'block': None}
        for entry in d.get('cells', []):
            r, c = entry['r'], entry['c']
            row_span = entry.get('row_span', 1)
            col_span = entry.get('col_span', 1)
            if not (0 <= r < g.n_rows and 0 <= c < g.n_cols):
                continue
            g.cells[(r, c)] = {'block': entry.get('block')}
            if row_span > 1 or col_span > 1:
                g.cells[(r, c)]['row_span'] = row_span
                g.cells[(r, c)]['col_span'] = col_span
                for rr in range(r, min(r + row_span, g.n_rows)):
                    for cc in range(c, min(c + col_span, g.n_cols)):
                        if (rr, cc) == (r, c):
                            continue
                        g.cells[(rr, cc)] = {'covered_by': (r, c)}
        return g


def save_layout(path, grid, name='Untitled'):
    with open(path, 'w') as f:
        json.dump(grid.to_dict(name), f, indent=2)


def load_layout(path):
    with open(path, 'r') as f:
        d = json.load(f)
    return Grid.from_dict(d)


def default_grid():
    """Starter layout so a brand-new window isn't blank.

    Columns D/E/F are revision columns - one column per revision, the
    Excel way. The Date of Issue / Revision Marks blocks in them are all
    'auto', so each column picks up the next revision in the model; add
    another revision column and it just continues the sequence.
    """
    g = Grid(n_rows=8, n_cols=6)
    from studio_blocks import new_block
    g.col_widths = [40, 75, 30, 16, 16, 16]
    g.row_heights = [10, 9, 7, 34, 9, 7, 20, 34]

    LAST = 5   # rightmost column index

    g.set_block(0, 0, new_block('proj_org', bold=True, size_mm=pt_to_mm(11)))
    g.merge(0, 0, 0, 2)
    g.set_block(0, 3, new_block('proj_number'))
    g.merge(0, 3, 0, LAST)

    g.set_block(1, 0, new_block('text', content='Distribution List',
                                bold=True, size_mm=pt_to_mm(12)))
    g.merge(1, 0, 1, LAST)

    g.set_block(2, 0, new_block('text', content='Sent To', bold=True))
    g.set_block(2, 1, new_block('text', content='Attention To', bold=True))
    g.set_block(2, 2, new_block('text', content='Copies', bold=True))
    g.merge(2, 2, 2, LAST)

    g.set_block(3, 0, new_block('sent_to', alt_rows=True))
    g.set_block(3, 1, new_block('attn_to', alt_rows=True))

    g.set_block(4, 0, new_block('text', content='Documentation',
                                bold=True, size_mm=pt_to_mm(12)))
    g.merge(4, 0, 4, LAST)

    g.set_block(5, 0, new_block('text', content='Sheet', bold=True))
    g.set_block(5, 1, new_block('text', content='Description', bold=True))
    g.set_block(5, 2, new_block('text', content='Revisions', bold=True))
    g.merge(5, 2, 5, LAST)

    # Revision columns: rotated issue date on top, marks down the doc list.
    for col in (3, 4, 5):
        g.set_block(6, col, new_block('spine_dates', just='center', rotation=270))
        g.set_block(7, col, new_block('spine_rev', just='center'))

    g.set_block(7, 0, new_block('sheet_number', alt_rows=True))
    g.set_block(7, 1, new_block('sheet_desc', alt_rows=True))

    return g
