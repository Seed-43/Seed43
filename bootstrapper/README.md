# Seed43 Bootstrapper

Standalone installer for the Seed43 PyRevit extension.
Downloads directly from GitHub — no Git installation required on the user's machine.

## What it does
1. Downloads `Seed43/archive/main.zip` from `github.com/Seed-43`
2. Extracts and copies the extension to `%APPDATA%\pyRevit\Extensions\seed43`
3. Writes a `version.txt` from `changelog.json` for future update checks
4. Prompts the user to reload PyRevit in Revit

## Files
| File | Purpose |
|------|---------|
| `main.py` | Full source — edit this |
| `Seed43_Setup.spec` | PyInstaller build spec |

## Building the EXE

```bash
pip install pyinstaller
pyinstaller Seed43_Setup.spec
```

Output: `dist/Seed43_Setup.exe`

## Adding an icon
Replace `icon=None` in `Seed43_Setup.spec` with `icon='seed43.ico'`
and place `seed43.ico` next to `main.py` before building.

## Updating the branch
If you rename `main` to something else, update `BRANCH` at the top of `main.py`.
