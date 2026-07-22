# -*- coding: utf-8 -*-
# script_create_pdf.py

_p = globals().get('PYTRANSMIT_PAYLOAD', {})

from pyrevit import revit, script, DB, forms
import re, os, time, sys, tempfile

output = script.get_output()
output.hide()

try:
    from Snippets import _dialogs as sdlg
except Exception:
    sdlg = None

def _alert(message, title='', exitscript=False):
    """Themed popup via the shared Snippets dialog lib, falls back to
    pyRevit's default forms.alert if the shared lib isn't available."""
    if sdlg:
        sdlg.message(message, title=title)
    else:
        forms.alert(message, title=title)
    if exitscript:
        script.exit()

def _confirm(message, title='', no='No'):
    """Themed yes/no popup, returns True on yes."""
    if sdlg:
        return sdlg.confirm(message, title=title, no=no)
    return bool(forms.alert(message, title=title, ok=False, yes=True, no=True))

# ── Log capture via payload ───────────────────────────────────────────
_log_lines = PYTRANSMIT_PAYLOAD.get('_log_lines', [])
def _log(msg):
    try: _log_lines.append(str(msg))
    except Exception: pass

class _SilentOutput(object):
    def print_md(self, msg, *a, **kw): _log(msg)
    def hide(self): pass
output = _SilentOutput()
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

import sys as _sys
_settings_dir = os.path.join(os.path.dirname(_script_dir), 'Settings')
if _settings_dir not in _sys.path:
    _sys.path.insert(0, _settings_dir)

if not os.path.exists(_excel_path):
    _alert("script_create_excel.py not found at:\n{}".format(_excel_path), exitscript=True)

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
    _alert("Error generating temp Excel:\n{}".format(
        _tb.format_exc() or str(_e)), exitscript=True)

if not os.path.exists(_temp_xlsx):
    _alert("Temp Excel file was not created:\n{}".format(_temp_xlsx), exitscript=True)

output.print_md("Temp Excel created: `{}`".format(_temp_xlsx))

# ── Step 2: Ask user where to save PDF ───────────────────────────────────────
_pdf_path = None
if '_pdf_save_path' in _p:
    _pdf_path = _p['_pdf_save_path']
    if _pdf_path is None:
        output.print_md('PDF export skipped.')
        try: os.remove(_temp_xlsx)
        except Exception: pass
        import sys as _sys; _sys.exit(0)
else:
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

# ── Step 3: Export via Shell (most reliable, no COM interop issues) ──────────
output.print_md("Exporting to PDF...")

# Delete existing PDF first so converter can write cleanly
if os.path.exists(_pdf_path):
    try:
        os.remove(_pdf_path)
    except Exception as _del_existing:
        _alert(
            "Cannot overwrite the existing PDF, it may be open in another program. Please close it and try again.",
            title="File In Use"
        )
        sys.exit(0)

import subprocess as _sp

# ── Try LibreOffice (no Microsoft Office required) ────────────────────────────
_lo_exe = None
for _c in [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
]:
    if os.path.exists(_c):
        _lo_exe = _c
        break

_pdf_written = False

if _lo_exe:
    output.print_md("Converting via LibreOffice...")
    try:
        _proc = _sp.Popen(
            [_lo_exe, '--headless', '--convert-to', 'pdf',
             '--outdir', os.path.dirname(_pdf_path), _temp_xlsx],
            stdout=_sp.PIPE, stderr=_sp.PIPE
        )
        _stdout, _stderr = _proc.communicate()
        _lo_out = os.path.join(
            os.path.dirname(_pdf_path),
            os.path.splitext(os.path.basename(_temp_xlsx))[0] + '.pdf'
        )
        if _proc.returncode == 0 and os.path.exists(_lo_out):
            if _lo_out != _pdf_path:
                os.rename(_lo_out, _pdf_path)
            _pdf_written = True
            output.print_md("PDF saved via LibreOffice: `{}`".format(_pdf_path))
        else:
            _err = (_stderr or b'').decode('utf-8', errors='replace').strip()
            output.print_md("LibreOffice failed: {}".format(_err))
    except Exception as _lo_e:
        output.print_md("LibreOffice error: {}".format(str(_lo_e)))

# ── Fallback: VBScript via Microsoft Excel ────────────────────────────────────
if not _pdf_written:
    output.print_md("Trying Microsoft Excel...")
    _vbs_path = os.path.join(_temp_dir, "pytransmit_pdf.vbs")
    _vbs_code = (
        'Dim xl, wb, xlPath, pdfPath\n'
        'xlPath  = WScript.Arguments(0)\n'
        'pdfPath = WScript.Arguments(1)\n'
        'Set xl = CreateObject("Excel.Application")\n'
        'xl.Visible       = False\n'
        'xl.DisplayAlerts = False\n'
        'Set wb = xl.Workbooks.Open(xlPath, False, True)\n'
        'wb.ExportAsFixedFormat 0, pdfPath, 0, True, False,,, False\n'
        'wb.Close False\n'
        'xl.Quit\n'
        'Set wb = Nothing\n'
        'Set xl = Nothing\n'
    )
    with open(_vbs_path, 'w') as _vf:
        _vf.write(_vbs_code)
    try:
        _proc = _sp.Popen(
            ['cscript', '//Nologo', _vbs_path, _temp_xlsx, _pdf_path],
            stdout=_sp.PIPE, stderr=_sp.PIPE
        )
        _stdout, _stderr = _proc.communicate()
        _rc  = _proc.returncode
        _err = (_stderr or b'').decode('utf-8', errors='replace').strip()
        if _rc == 0 and os.path.exists(_pdf_path):
            _pdf_written = True
            output.print_md("PDF saved via Microsoft Excel: `{}`".format(_pdf_path))
        else:
            output.print_md("Excel VBScript failed: {}".format(_err))
    except Exception as _vbs_e:
        output.print_md("VBScript error: {}".format(str(_vbs_e)))
    finally:
        try: os.remove(_vbs_path)
        except Exception: pass

# ── No converter found ────────────────────────────────────────────────────────
if not _pdf_written:
    _alert(
        "No supported PDF converter was found. Install Microsoft Office (Excel) or LibreOffice (free) from libreoffice.org, then re-run pyTransmit.",
        title="PDF Export Failed"
    )
    try: os.remove(_temp_xlsx)
    except Exception: pass
    sys.exit(0)

# ── Step 4: Delete temp Excel ─────────────────────────────────────────────────
for _attempt in range(8):
    try: os.remove(_temp_xlsx); break
    except Exception: time.sleep(0.5)

output.print_md("## Done!")
if _pdf_path and os.path.exists(_pdf_path):
    output.print_md("PDF: `{}`".format(_pdf_path))
    _open_dlg = _p.get('_open_file_dialog')
    if _open_dlg:
        if _open_dlg(u'PDF Saved', u'Open the file?'):
            os.startfile(_pdf_path)
    elif _confirm("PDF saved!\n\nOpen the file?"):
        os.startfile(_pdf_path)
