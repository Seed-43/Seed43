# -*- coding: utf-8 -*-
"""
layout_Algoruthum.py
====================
General row height distribution algorithm for pyTransmit layout exporters.

Used by: script_create_excel.py, script_create_schedule.py, and any future exporters.

The algorithm handles the general case of a layout grid where:
  - Rows are grouped via merge_down
  - Col A (and optionally B) may have content on some rows but not others
  - Col C/D blocks may span multiple rows vertically (dates, legends, spine data)
  - The total height of a spanning block must be distributed across the rows it spans

Rules
-----
1. A row whose col A has a single-line block → fixed height (SINGLE_LINE_H)
2. A row whose col A spans down (is None in this row because a block above merges
   through it) → free height (receives distributed remainder)
3. For each spanning block in cols B-D:
   - Total required height = block_height_fn(block)
   - Fixed rows within the span contribute their fixed height
   - Remainder is split equally among free rows in the span
4. A row's final height = max of all requirements placed on it

Usage
-----
    from layout_Algoruthum import compute_row_heights

    heights = compute_row_heights(
        rows          = ROWS,           # list of layout row dicts
        groups        = GROUPS,         # list of (start_ri, end_ri) tuples
        excel_starts  = _row_excel_start,  # {layout_ri: first_excel_row}
        excel_counts  = _row_excel_count,  # {layout_ri: n_excel_rows}
        block_height_fn = _row_height_pt,  # fn(block) -> float pts
        single_line_h = 18.0,           # height for a single-line row
        min_h         = 15.0,           # minimum row height
    )
    # heights: {excel_row_index: float}

    for er, h in heights.items():
        ws.set_row(er, h)
"""


def compute_row_heights(rows, groups, excel_starts, excel_counts,
                        block_height_fn, single_line_h=18.0, min_h=15.0):
    """
    Compute Excel row heights from a layout JSON row/group structure.

    Parameters
    ----------
    rows : list of dict
        Layout row dicts, each with 'blocks' (list of 4) and 'merge_down' bool.
    groups : list of (int, int)
        (start_ri, end_ri) pairs from _get_groups().
    excel_starts : dict
        {layout_row_index: first_excel_row_index}
    excel_counts : dict
        {layout_row_index: number_of_excel_rows_for_this_layout_row}
    block_height_fn : callable
        fn(block_dict) -> float  — returns required height in pts for a block.
    single_line_h : float
        Height for a row containing only single-line col A content.
    min_h : float
        Minimum row height for any row.

    Returns
    -------
    dict : {excel_row_index: float}
    """

    heights = {}  # excel_row_index -> float

    def _set(er, h):
        heights[er] = max(heights.get(er, min_h), h)

    def _force(er, h):
        """Override — used when we know the exact value (not a max)."""
        heights[er] = h

    # ── Pass 1: initial heights from single-row blocks ──────────────────────
    # Skip spanning blocks — their height is handled in Pass 2 distribution.
    # A block "spans" if the same column is None in the next row of its group.
    def _is_spanning(ri, ci):
        """True if this block spans into the next row (next row ci is None)."""
        row = rows[ri]
        if not row.get('merge_down', False):
            return False
        next_row = rows[ri + 1] if ri + 1 < len(rows) else None
        if next_row is None:
            return False
        nb = _get_block(next_row, ci)
        return nb is None

    for ri, row in enumerate(rows):
        if row.get('section') == 'footer':
            continue
        er_s = excel_starts.get(ri)
        if er_s is None:
            continue
        er_c = excel_counts.get(ri, 1)

        for ci, block in enumerate(row.get('blocks', [])):
            if not block:
                continue
            # Skip spanning blocks — Pass 2 handles their height distribution
            if _is_spanning(ri, ci):
                continue
            bh = block_height_fn(block)
            if er_c > 1:
                for sub in range(er_c):
                    _set(er_s + sub, bh)
            else:
                _set(er_s, bh)

    # ── Pass 2: group-level height distribution ──────────────────────────────
    for gs, ge in groups:
        if gs == ge:
            continue  # single layout row — no distribution needed

        # For each layout col (0-3), find what spans it vertically in this group
        # A "span" is a block at row R in col C that is None in rows R+1, R+2...
        # Those None rows are part of the merged vertical span.

        for ci in range(4):
            # Walk the group and identify spans for this column
            _ri = gs
            while _ri <= ge:
                block = _get_block(rows[_ri], ci)
                if block is None:
                    _ri += 1
                    continue

                # This block starts at _ri — find how far it spans (where col ci is None)
                span_end = _ri
                nri = _ri + 1
                while nri <= ge:
                    nb = _get_block(rows[nri], ci)
                    if nb is None:
                        span_end = nri
                        nri += 1
                    else:
                        break

                if span_end == _ri:
                    # Span of 1 — already handled in Pass 1
                    _ri += 1
                    continue

                # Multi-row span: distribute block height across _ri..span_end
                span_ris = list(range(_ri, span_end + 1))
                _distribute_span(
                    rows, span_ris, ci,
                    block, block_height_fn, single_line_h, min_h,
                    excel_starts, excel_counts, heights, _set, _force
                )

                _ri = span_end + 1

    # ── Pass 3: distribute legend content across its span rows ──────────────
    # For reason_list/method_list spanning N rows:
    #   per_row = max(single_line_h, legend_height / N)
    #   first row also gets max(per_row, date_remainder) if spine_dates is in prev row
    #   all other rows get just per_row
    for gs, ge in groups:
        if gs == ge:
            continue
        for ri in range(gs, ge + 1):
            row = rows[ri]
            # Check col A or B for reason_list/method_list
            legend_block = None
            for ci in range(2):
                bl = _get_block(row, ci)
                if bl and bl.get('type') in ('reason_list', 'method_list'):
                    legend_block = bl
                    break
            if legend_block is None:
                continue

            # Find span extent (where col A is None in subsequent rows)
            span_end = ri
            nri = ri + 1
            while nri <= ge:
                nb = _get_block(rows[nri], 0) or _get_block(rows[nri], 1)
                if nb is None:
                    span_end = nri
                    nri += 1
                else:
                    break

            n_rows = span_end - ri + 1
            legend_h = block_height_fn(legend_block)
            per_row = max(single_line_h, legend_h / n_rows)

            # Check if prev row in group had spine_dates → date remainder for first row
            date_remainder = 0.0
            if ri > gs:
                prev_blocks = rows[ri - 1].get('blocks', [])
                if any(b and b.get('type') == 'spine_dates' for b in prev_blocks):
                    prev_er = excel_starts.get(ri - 1, 0)
                    prev_h = heights.get(prev_er, single_line_h)
                    date_remainder = max(0.0, 65.0 - prev_h)

            # Apply heights: first row gets date_remainder boost, rest get per_row
            for idx, sri in enumerate(range(ri, span_end + 1)):
                ser_s = excel_starts.get(sri, 0)
                ser_c = excel_counts.get(sri, 1)
                row_h = max(per_row, date_remainder) if idx == 0 else per_row
                for sub in range(ser_c):
                    heights[ser_s + sub] = row_h

    # Ensure all excel rows have at least min_h
    all_excel_rows = set()
    for ri in range(len(rows)):
        er_s = excel_starts.get(ri, 0)
        er_c = excel_counts.get(ri, 1)
        for sub in range(er_c):
            all_excel_rows.add(er_s + sub)

    for er in all_excel_rows:
        if er not in heights:
            heights[er] = min_h
        else:
            heights[er] = max(heights[er], min_h)

    return heights


def _get_block(row, ci):
    """Safely get block at column index ci from a row dict."""
    blocks = row.get('blocks', [])
    if ci < len(blocks):
        return blocks[ci]
    return None


def _distribute_span(rows, span_ris, ci, block, block_height_fn,
                     single_line_h, min_h, excel_starts, excel_counts,
                     heights, _set, _force):
    """
    Distribute a spanning block's required height across the rows it spans.

    For each row in span_ris:
      - If another col (A or B) in that row has a single-line block → fixed row
      - If no other col has content → free row

    Algorithm:
      total_required = block_height_fn(block)
      fixed_rows contribute their current (or single_line_h) height
      remainder = total_required - sum(fixed heights)
      free rows each get: remainder / n_free_rows
      Each row's final height = max(its fixed height, distributed amount)
    """
    required_h = block_height_fn(block)

    # Classify each row in the span as fixed or free
    # "Fixed" = has a single-line block in col A (ci != the spanning col)
    # "Free"  = no other col A content (this row exists only to extend the span)
    fixed_ris = []   # (ri, fixed_height)
    free_ris  = []   # ri

    for ri in span_ris:
        row = rows[ri]
        # Check col A (and col B if spanning col is not 0 or 1)
        other_cols = [c for c in range(4) if c != ci]
        fixed_h = None
        for oc in other_cols:
            ob = _get_block(row, oc)
            if ob is None:
                continue
            ot = ob.get('type', '')
            # reason_list/method_list ARE the content driving total height
            # treat them as free rows (not fixed anchors)
            if ot in ('reason_list', 'method_list'):
                fixed_h = None
                break
            oh = _row_height_fn_single(ob, block_height_fn, single_line_h)
            if fixed_h is None:
                fixed_h = oh
            else:
                fixed_h = max(fixed_h, oh)

        if fixed_h is not None:
            fixed_ris.append((ri, fixed_h))
        else:
            free_ris.append(ri)

    # Calculate remainder for free rows
    fixed_total = sum(h for _, h in fixed_ris)
    remainder = max(0.0, required_h - fixed_total)

    if free_ris:
        per_free = remainder / len(free_ris)
    else:
        # No free rows — distribute evenly across all rows
        per_free = required_h / len(span_ris) if span_ris else required_h

    # Apply heights
    for ri, fh in fixed_ris:
        er_s = excel_starts.get(ri, 0)
        er_c = excel_counts.get(ri, 1)
        for sub in range(er_c):
            _set(er_s + sub, fh)

    for ri in free_ris:
        er_s = excel_starts.get(ri, 0)
        er_c = excel_counts.get(ri, 1)
        # If multiple excel rows in this layout row, distribute further
        per_excel = per_free / er_c if er_c > 0 else per_free
        for sub in range(er_c):
            _set(er_s + sub, max(min_h, per_excel))


def _row_height_fn_single(block, block_height_fn, single_line_h):
    """
    Return the height for a block when it is the "fixed" anchor in a span.
    All blocks are treated as single-line when acting as anchors.
    The spanning block's total height is what drives distribution.
    """
    if block is None:
        return single_line_h
    return single_line_h


def get_groups(rows):
    """
    Build group spans from merge_down flags.
    Returns list of (start_ri, end_ri) tuples.
    Mirrors the _get_groups() function used in exporters.
    """
    groups = []
    i = 0
    while i < len(rows):
        start = i
        while i < len(rows) - 1 and rows[i].get('merge_down', False):
            i += 1
        groups.append((start, i))
        i += 1
    return groups
