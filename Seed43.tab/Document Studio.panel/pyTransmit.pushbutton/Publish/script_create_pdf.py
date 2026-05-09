# -*- coding: utf-8 -*-
"""
pyTransmit — PDF Creator
Driven by PYTRANSMIT_PAYLOAD injected from script.py via exec().
"""

_p = globals().get('PYTRANSMIT_PAYLOAD', {})

from pyrevit import revit, script, DB, forms
import re, os, time, sys, tempfile

output = script.get_output()
doc    = revit.doc

# ── Build temp Excel filename ─────────────────────────────────────────────────
proj_info   = doc.ProjectInformation
proj_number = ''
proj_name   = ''
try:
    p = proj_info.LookupParameter("Project Number")
    if p and p.HasValue: proj_number = (p.AsString() or p.AsValueString() or '').strip()
    p = proj_info.LookupParameter("Project Name")
    if p and p.HasValue: proj_name = (p.AsString() or p.AsValueString() or '').strip()
    if not proj_name: proj_name = doc.Title or ''
except Exception: pass

_safe_base = re.sub(r'[\\/*?:"<>|]', '_',
    "Document_Transmittal_{}_{}".format(proj_number, proj_name))[:60]

_temp_dir  = tempfile.gettempdir()
_temp_xlsx = os.path.join(_temp_dir, "{}_TEMP.xlsx".format(_safe_base))

# ── Step 1: Generate temp Excel ───────────────────────────────────────────────
_script_dir = os.path.dirname(os.path.abspath(__file__))
_excel_path = os.path.join(_script_dir, 'script_create_excel.py')

if not os.path.exists(_excel_path):
    forms.alert("script_create_excel.py not found at:\n{}".format(_excel_path), exitscript=True)

_payload_for_excel = dict(_p)
_payload_for_excel['_pdf_temp_xlsx_path'] = _temp_xlsx

_ns = {
    '__name__':           'excel_for_pdf',
    '__file__':           _excel_path,
    '__builtins__':       __builtins__,
    'PYTRANSMIT_PAYLOAD': _payload_for_excel,
}
with open(_excel_path, 'r') as _f:
    _src = _f.read()
try:
    exec(_src, _ns)
except Exception as _e:
    import traceback as _tb
    forms.alert("Error generating temp Excel:\n{}".format(
        _tb.format_exc() or str(_e)), exitscript=True)

if not os.path.exists(_temp_xlsx):
    forms.alert("Temp Excel file was not created:\n{}".format(_temp_xlsx), exitscript=True)

output.print_md("Temp Excel created: `{}`".format(_temp_xlsx))

# ── Step 2: Ask user where to save PDF ───────────────────────────────────────
_pdf_path = None
try:
    from System.Windows.Forms import SaveFileDialog, DialogResult
    dlg = SaveFileDialog()
    dlg.Title            = "Save Transmittal - PDF"
    dlg.Filter           = "PDF File (*.pdf)|*.pdf"
    dlg.FileName         = "{}.pdf".format(_safe_base)
    dlg.InitialDirectory = os.path.expanduser("~\\Desktop")
    if dlg.ShowDialog() == DialogResult.OK:
        _pdf_path = dlg.FileName
    else:
        output.print_md("PDF save cancelled.")
        try: os.remove(_temp_xlsx)
        except Exception: pass
        sys.exit(0)
except Exception:
    _pdf_path = os.path.join(os.path.expanduser("~\\Desktop"), "{}.pdf".format(_safe_base))

output.print_md("PDF path: `{}`".format(_pdf_path))

# ── Step 3: Export via Shell (most reliable — no COM interop issues) ──────────
output.print_md("Exporting to PDF via VBScript...")

# Delete existing PDF first so VBScript can write cleanly
if os.path.exists(_pdf_path):
    try:
        os.remove(_pdf_path)
    except Exception as _del_existing:
        forms.alert(
            "Cannot overwrite existing PDF — it may be open in another program.\n\n"
            "Please close the file and try again.\n\n{}".format(_del_existing),
            exitscript=True
        )

# Write a VBScript helper — no Python needed, no COM interop issues in IronPython
_vbs_path = os.path.join(_temp_dir, "pytransmit_pdf.vbs")
_vbs_code = u'''
Dim xl, wb, xlPath, pdfPath
xlPath  = WScript.Arguments(0)
pdfPath = WScript.Arguments(1)

Set xl = CreateObject("Excel.Application")
xl.Visible       = False
xl.DisplayAlerts = False

Set wb = xl.Workbooks.Open(xlPath, False, True)
wb.ExportAsFixedFormat 0, pdfPath, 0, True, False, , , False
wb.Close False
xl.Quit

Set wb = Nothing
Set xl = Nothing
'''

with open(_vbs_path, 'w') as _vf:
    _vf.write(_vbs_code)

import subprocess as _sp
try:
    _proc = _sp.Popen(
        ['cscript', '//Nologo', _vbs_path, _temp_xlsx, _pdf_path],
        stdout=_sp.PIPE,
        stderr=_sp.PIPE
    )
    _stdout, _stderr = _proc.communicate()
    _out = (_stdout or b'').decode('utf-8', errors='replace').strip()
    _err = (_stderr or b'').decode('utf-8', errors='replace').strip()
    _rc  = _proc.returncode

    if _rc == 0:
        if os.path.exists(_pdf_path):
            output.print_md("PDF saved: `{}`".format(_pdf_path))
        else:
            forms.alert("PDF not created.\n\nOutput: {}\nError: {}".format(_out, _err))
    else:
        _msg = _err or _out or "returncode={}".format(_rc)
        output.print_md("VBScript failed: {}".format(_msg))
        forms.alert("PDF export failed:\n{}".format(_msg))

except Exception as _se:
    import traceback as _tb3
    forms.alert("PDF export error:\n{}".format(_tb3.format_exc() or str(_se)))

finally:
    try: os.remove(_vbs_path)
    except Exception: pass

# ── Step 4: Delete temp Excel ─────────────────────────────────────────────────
for _attempt in range(8):
    try: os.remove(_temp_xlsx); break
    except Exception: time.sleep(0.5)

output.print_md("## Done!")
if _pdf_path and os.path.exists(_pdf_path):
    output.print_md("PDF: `{}`".format(_pdf_path))
    if forms.alert("PDF saved!\n\nOpen the file?", yes=True, no=True):
        os.startfile(_pdf_path)
