# -*- coding: utf-8 -*-
# script_create_log.py

import os
import json
import zipfile
import datetime

# ── PUBLIC API ────────────────────────────────────────────────────────

def build_log(payload, log_lines, zip_path, layout_paths):
    """
    Write the log text and zip it together with the layout JSON files used.

    payload      : the full export payload dict from pyTransmit.py
    log_lines    : list of strings collected during the export run
    zip_path     : full path to the output zip file chosen by the user
    layout_paths : list of full paths to the layout JSON files that were used
    """
    timestamp = datetime.datetime.now().strftime('%Y-%m-%d  %H:%M:%S')
    run_tag   = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')

    # ── Build text log ────────────────────────────────────────────────
    lines = []

    lines.append('=' * 68)
    lines.append('  pyTransmit Export Log')
    lines.append('  {}'.format(timestamp))
    lines.append('=' * 68)
    lines.append('')

    # ── Run summary ───────────────────────────────────────────────────
    lines.append('RUN SUMMARY')
    lines.append('-' * 68)
    output_types = payload.get('output_types', [])
    lines.append('Output types : {}'.format(', '.join(output_types) if output_types else 'none'))
    lines.append('Page height  : {} ({} mm)'.format(
        payload.get('page_height_mode', ''),
        payload.get('page_height_mm', '')))
    lines.append('')

    # ── Project info ──────────────────────────────────────────────────
    lines.append('PROJECT')
    lines.append('-' * 68)
    lines.append('Number : {}'.format(payload.get('proj_number', '')))
    lines.append('Name   : {}'.format(payload.get('proj_name', '')))
    lines.append('')

    # ── Revision meta rows ────────────────────────────────────────────
    lines.append('REVISION METADATA')
    lines.append('-' * 68)
    for label, value in (payload.get('meta_rows') or []):
        lines.append('  {:20s} {}'.format(label, value))
    lines.append('')

    # ── Recipients ────────────────────────────────────────────────────
    lines.append('RECIPIENTS')
    lines.append('-' * 68)
    for rec in (payload.get('recipients') or []):
        lines.append('  {:30s} Attn: {:20s}'.format(
            rec.get('label', ''), rec.get('attn', '')))
    lines.append('')

    # ── Reason legend ─────────────────────────────────────────────────
    lines.append('REASON FOR ISSUE LEGEND')
    lines.append('-' * 68)
    lines.append(payload.get('reason_legend', '') or '')
    lines.append('')

    # ── Output file paths ─────────────────────────────────────────────
    lines.append('OUTPUT FILES')
    lines.append('-' * 68)
    for key, path in (payload.get('output_paths') or {}).items():
        lines.append('  {:20s} {}'.format(key, path))
    lines.append('')

    # ── Export event log ──────────────────────────────────────────────
    lines.append('EXPORT LOG')
    lines.append('-' * 68)
    for entry in log_lines:
        lines.append(entry)
    lines.append('')

    # ── Full payload JSON ─────────────────────────────────────────────
    lines.append('FULL SETTINGS PAYLOAD')
    lines.append('-' * 68)
    try:
        # Strip non-serialisable items before dumping
        _safe = {k: v for k, v in payload.items()
                 if isinstance(v, (str, int, float, bool, list, dict, type(None)))}
        lines.append(json.dumps(_safe, indent=2))
    except Exception as e:
        lines.append('Could not serialise payload, {}'.format(str(e)))
    lines.append('')

    log_text = '\n'.join(lines)
    log_filename = 'pyTransmit_log_{}.txt'.format(run_tag)

    # ── Write zip ─────────────────────────────────────────────────────
    try:
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:

            # Text log
            zf.writestr(log_filename, log_text)

            # Layout JSON files used in this export
            for fpath in (layout_paths or []):
                if fpath and os.path.isfile(fpath):
                    try:
                        zf.write(fpath, os.path.join('Layouts', os.path.basename(fpath)))
                    except Exception:
                        pass

        return True, zip_path

    except Exception as e:
        return False, str(e)
