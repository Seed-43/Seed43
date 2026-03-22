# Seed43 Bootstrapper

Standalone installer for the Seed43 PyRevit extension.
Downloads directly from GitHub — no Git installation required on the user's machine.

## Requirements
- Python 3.x installed (included tkinter — default on Windows)
- Internet connection

## What it does
1. Downloads `Seed43/archive/main.zip` from `github.com/Seed-43`
2. Extracts and installs the full extension to:
   `%APPDATA%\pyRevit\Extensions\seed43`
3. Confirms `script.py` and `seed43.xaml` are present in the About pushbutton folder
4. Writes `version.txt` from `changelog.json` for future update checks
5. Prompts the user to reload PyRevit in Revit

## How to run
Double-click `Seed43_Setup.pyw`
(Python must be installed — python.org/downloads)

## Files
| File | Purpose |
|------|---------|
| `Seed43_Setup.pyw` | Run directly by users — no compile needed |
| `README.md` | This file |

## Updating the branch
If you rename `main` to something else, update `BRANCH` at the top of `Seed43_Setup.pyw`.
