# pyTable

Import Excel and Word documents into Revit as Legend, Drafting, or Schedule views.
Part of the Seed43 pyRevit extension.

---

## Folder structure

```
PyTables.pushbutton/
    script.py       Main script and WPF controller
    PyTables.xaml   Window layout (merges Seed43Styles.xaml)
    README.md       This file
```

Place the `PyTables.pushbutton` folder inside your panel stack as normal.
`Seed43Styles.xaml` must sit at the extension root. Adjust the `Source` path
in the `MergedDictionaries` block inside `PyTables.xaml` if your layout differs.

---

## What it does

pyTable links source documents to Revit views and tracks those links inside
the project file using Revit ExtensibleStorage so the data travels with
the model.

Three import modes are supported:

**Table** (Legend or Drafting view)
Reads cell values via openpyxl and places TextNote elements in a grid layout.
Bold cells map to the bold TextNoteType if one exists. Works without Pillow.

**Table** (Schedule view)
Reads the same cell data and writes it into a new schedule's header section,
using the same approach as pyTransmit's schedule header population.

**Image** (Legend or Drafting view)
Renders the Excel range to a PNG using System.Drawing (no Pillow needed),
then inserts it as an ImageInstance via ImageType.Create (Revit 2020+).
The temp PNG is deleted after insertion.

---

## Data storage

All linked document records are stored as JSON inside a single
`DataStorage` element in the active Revit project using ExtensibleStorage.
Schema GUID: `A3F7C2B1-4D56-4E89-B021-9C8E1F234567`

Records persist when the model is saved, copied, or sent to consultants.
They are not affected by Purge Unused.

Each record stores:
- Source file path (absolute)
- Worksheet name and named range / range address
- Revit view name and ElementId integer
- Import type, view type, DPI, scale, conflict policy
- Auto-sync flag
- Status (new / live / stale / error)
- Last synced timestamp

---

## openpyxl availability

pyTable calls openpyxl for Excel reading. If openpyxl is not available in the
pyRevit IronPython environment the tool will still open, but Table mode will
produce empty views and Image mode will fall back to a blank render.

To check availability open the pyRevit console and run:
```python
import openpyxl
print(openpyxl.__version__)
```

If it is missing, copy the openpyxl package folder from a standard CPython
installation into the pyRevit IronPython site-packages directory.

---

## Revit version notes

- Table mode (Legend/Drafting): Revit 2019+
- Table mode (Schedule header): Revit 2019+
- Image mode (ImageType.Create): Revit 2020+

For Revit 2019 with Image mode, the `insert_image_into_view` function will
raise a runtime error. Table mode works fully on 2019.

---

## Known limitations (V1)

- File watcher auto-sync is not yet implemented. The Auto-sync toggle is
  stored per entry but the polling mechanism is not wired up. Sync is
  manual via the Sync button or Sync all menu item.
- Named range detection requires openpyxl. Without it the range dropdown
  only shows `<Used Range>` and `<Print Area>`.
- Schedule view import populates the header section only. Body rows are
  not written in V1.
- Word (.docx) files only support Image mode. Table mode is greyed out
  when a .docx file is selected (to be enforced in V2 UI validation).

---

## Development notes

All .NET string operations use `safe_str()` via `System.Text.Encoding.UTF8`
to avoid IronPython 2 codec errors on non-ASCII paths.

All file system operations use `System.IO` (File, Directory, Path).
No `os.path` calls anywhere in the script.

ElementId integer values use `getattr(eid, 'Value', None) or getattr(eid, 'IntegerValue', None)`
for Revit 2024+ compatibility.

Row expand/collapse is handled entirely in Python. No WPF triggers or
DataTriggers are used for the accordion behaviour.

The two-step import flow (Step 1: source config, Step 2: import type and
render settings) lives inside the expanded row detail, not in a separate
dialog. This matches the design prototype exactly.
