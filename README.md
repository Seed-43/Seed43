# Seed43 | PyRevit Extension

**Seed43** is a free PyRevit extension: a set of documentation, drafting and model tools for Revit, with an in-app manager to update itself and toggle individual tools from GitHub.

---

## Tools

### Document Studio

| Tool | Description |
|------|-------------|
| **pySheets** | Prints and exports sheets and views to PDF, DWG, DGN, NWC, IFC and image formats. Saves reusable profiles, and can print on a schedule while you carry on working. |
| **pyTransmit** | Collects revision and issue details, then publishes a transmittal document in your chosen output formats. |
| **pyLink** | Links Excel and Word documents to Revit views — named ranges and document sections become legends, schedules or drafting views built from real Revit data, not images. |
| **pyTable** | Exports model categories and schedule parameters to Excel or LibreOffice Calc, edit them externally, then import the changes back into the model. |
| **pyFilter** | Creates and manages View Filter templates, then pushes them onto views or view templates — or pulls existing filters back into a template. |
| **CAD Layer Manager** | Shows or hides individual CAD layers from the view you are working in, without digging through Visibility/Graphics. |
| **Filter Manager** | Builds View Filters from a selected host or linked element by type or family, deletes filters with automatic backup, and restores them again. |
| **Text Formatting** | Text Note tools — batch formatting and type changes, find and replace, matching alignment, properties or width, merging notes, and exploding them. |
| **Views Manager** | Places view references onto sheets, moves viewports between sheets, duplicates sheets and views, and renames and organises them in bulk. |
| **Revision** | Permanently links revision clouds to their parent sheet, colours clouds and tags by issued status, and writes sheet and revision data to cloud parameters. |

### 3D Tools

| Tool | Description |
|------|-------------|
| **Isolate Levels** | Temporarily isolates everything tied to a chosen level, so you can see what deleting that level would take with it. |
| **Model Slice** | Sets the 3D view's section box to a thin slice around a chosen level or grid — 500 mm either side of the line. |
| **Rebar** | Assigns rebar to a workset, toggles visibility between obscured and unobscured, and updates a tag parameter across elements. |

### Dev Tools

| Tool | Description |
|------|-------------|
| **Find Element** | Finds which sheet or sheets the selected elements are placed on. |
| **Inspector** | Pick any element and get a full breakdown of what Revit knows about it — identity, type parameters, family info, location and geometry. |

### Family Tools

| Tool | Description |
|------|-------------|
| **Project Units** | Saves and restores all project unit settings via a JSON file, including accuracy, grouping, prefix, suffix and display options. |

---

## Appearance

Every Seed43 window shares one palette. Open **About → Appearance** to switch between dark and light, or pick an accent colour — the windows update immediately, and the toolbar icons are redrawn to match the next time Revit starts.

---

## Compatibility

- Revit 2025 / 2026
- [PyRevit](https://github.com/pyrevitlabs/pyRevit) required

---

## Installation

### Option 1 - Setup Installer

1. Download `Seed43_Setup.pyw` from this repo
2. Double-click to run
3. Click **Install**
4. Reload PyRevit inside Revit

### Option 2 - Batch Script

1. Download `install.bat` from this repo
2. Double-click to run
3. Reload PyRevit inside Revit

### Option 3 - PyRevit Extension Manager

1. Open PyRevit Settings → Extensions
2. Add custom extension source:
3. Find **Seed43** in the list and click Install

---

## After Installation

The **Seed43** tab will appear in Revit after reloading PyRevit.

Use the **About** button to check which version is installed, see what has changed, update to the latest version, and turn individual tools and panels on or off. Updating needs an internet connection, and PyRevit must be reloaded afterwards for the changes to take effect.

---

## Contact

- Website: [seed43.org](https://seed43.org)
- LinkedIn: [Seed43](https://www.linkedin.com/company/seed43/)
- GitHub: [Seed-43](https://github.com/Seed-43)
- Issues: [GitHub Issues](https://github.com/Seed-43/Seed43/issues)

---

## License

GNU General Public License v3 - free to use and modify.
