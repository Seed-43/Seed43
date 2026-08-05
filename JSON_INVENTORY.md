# Seed43 JSON Inventory

Every `.json` under `Seed43.tab` and `lib`, for deciding what moves into `.user`.

Generated 2026-08-05. **36 files, all currently git-tracked** - so each one ships as a
default *and* is overwritten by the user at runtime. That is the problem being solved.

Fill in **Verdict**. Suggested values:

| Verdict | Meaning | Where it should live |
|---|---|---|
| `USER` | personal, changes as they work | `.user/<Tool>/`, stop shipping |
| `COMPANY` | shared config you want to push updates to | stays shipped, but needs an override in `.user` |
| `SHIPPED` | pure app data, users never edit | stays exactly where it is |
| `MIXED` | both - needs splitting | decide per key |


## pyTransmit (22)

| File | Size | Modified | Top-level keys | My guess | Verdict |
|---|---:|---|---|---|---|
| `Layout/Layouts/Excel.json` | 31.8 KB | 2026-08-04 | `page_w_mm`, `rev_count`, `hlines`, `vlines`, ... |  | |
| `Layout/Layouts/PDF.json` | 31.9 KB | 2026-08-04 | `page_w_mm`, `rev_count`, `hlines`, `vlines`, ... |  | |
| `Layout/Layouts/Revit Drafting View.json` | 15.9 KB | 2026-08-04 | `page_w_mm`, `rev_count`, `hlines`, `vlines`, ... |  | |
| `Layout/Layouts/Revit Legend.json` | 19.6 KB | 2026-08-04 | `page_w_mm`, `rev_count`, `hlines`, `vlines`, ... |  | |
| `Layout/Layouts/Revit Schedule.json` | 19.6 KB | 2026-08-04 | `page_w_mm`, `rev_count`, `hlines`, `vlines`, ... |  | |
| `Layout/layout_config.json` | 0.7 KB | 2026-08-04 | `active_template`, `logo_path`, `text_styles`, `col_pct` | USER - active template + logo path | |
| `Settings/branding.json` | 0.3 KB | 2026-07-25 | `header_fg_color`, `logo_source`, `title_fg_color`, `header_bg_color`, ... | COMPANY - logo/colours | |
| `Settings/distribution.json` | 0.4 KB | 2026-07-25 | list of 4 items | COMPANY - distribution list | |
| `Settings/format.json` | 0.1 KB | 2026-07-25 | list of 2 items | COMPANY - dropdown vocabulary | |
| `Settings/method.json` | 0.5 KB | 2026-07-25 | list of 6 items | COMPANY - dropdown vocabulary | |
| `Settings/printsize.json` | 0.1 KB | 2026-07-25 | list of 4 items | COMPANY - dropdown vocabulary | |
| `Settings/pytransmit_setup.json` | 0.2 KB | 2026-08-04 | `output_path_template`, `projects_older_root`, `projects_root`, `transmittal_naming_template` | USER/COMPANY - paths + naming templates | |
| `Settings/reason.json` | 0.3 KB | 2026-07-25 | list of 4 items | COMPANY - dropdown vocabulary | |
| `Settings/recipients.json` | 1.2 KB | 2026-07-25 | list of 9 items | COMPANY - contact list | |
| `Studio/studio_config.json` | 0.2 KB | 2026-08-04 | `last_file` | USER - remembers last opened file | |
| `Studio/studio_layouts/Excel.json` | 111.2 KB | 2026-08-04 | `row_sections`, `n_cols`, `page_size_name`, `orientation`, ... |  | |
| `Studio/studio_layouts/PDF.json` | 105.8 KB | 2026-08-03 | `name`, `n_rows`, `n_cols`, `row_heights`, ... |  | |
| `Studio/studio_layouts/Revit Drafting View.json` | 54.3 KB | 2026-08-03 | `name`, `n_rows`, `n_cols`, `row_heights`, ... |  | |
| `Studio/studio_layouts/Revit Legend.json` | 54.3 KB | 2026-08-03 | `name`, `n_rows`, `n_cols`, `row_heights`, ... |  | |
| `Studio/studio_layouts/Revit Schedule.json` | 54.3 KB | 2026-08-03 | `name`, `n_rows`, `n_cols`, `row_heights`, ... |  | |
| `pytransmit_setup.json` | 0.5 KB | 2026-07-26 | `show_method`, `out_pdf`, `recipient_mode`, `page_height_mm`, ... | USER/COMPANY - paths + naming templates | |
| `pytransmit_sync.json` | 0.0 KB | 2026-08-04 | `group_label_on` | USER - single toggle | |

## PySheets (5)

| File | Size | Modified | Top-level keys | My guess | Verdict |
|---|---:|---|---|---|---|
| `userdata/folder_presets/folder_presets.json` | 0.2 KB | 2026-07-25 | `Project` | USER? - or company standard | |
| `userdata/profiles/test.json` | 2.6 KB | 2026-08-04 | `viewing`, `view_sel`, `auto_overwrite`, `formats`, ... | USER - a saved profile | |
| `userdata/settings/custom_columns.json` | 0.1 KB | 2026-07-25 | `sheet_columns`, `builtin_visible`, `view_columns` | USER - chosen columns | |
| `userdata/settings/lastsession.json` | 1.9 KB | 2026-08-04 | `dwg`, `open_after`, `column_widths`, `dgn`, ... | USER - session state | |
| `userdata/settings/naming_memory.json` | 0.2 KB | 2026-08-04 | `per_format` | USER - remembers last entries | |

## pyLink (3)

| File | Size | Modified | Top-level keys | My guess | Verdict |
|---|---:|---|---|---|---|
| `userdata/excel_font_settings.json` | 0.0 KB | 2026-07-31 | `fallback_font` | USER - font preference | |
| `userdata/section_groups.json` | 0.6 KB | 2026-07-31 | `groups` | USER - user-built groups | |
| `userdata/word_text_settings.json` | 0.1 KB | 2026-07-31 | `mode`, `size_mm`, `text_type_name` | USER - text preference | |

## lib (2)

| File | Size | Modified | Top-level keys | My guess | Verdict |
|---|---:|---|---|---|---|
| `Snippets/_icons.json` | 19.6 KB | 2026-07-25 | `about`, `arrow_down`, `arrow_left`, `arrow_right`, ... | SHIPPED - icon vector data | |
| `Snippets/seed43_palette.json` | 5.2 KB | 2026-08-01 | `dimensions`, `profiles`, `active_profile` | MIXED - tokens shipped, active_profile is user | |

## About (1)

| File | Size | Modified | Top-level keys | My guess | Verdict |
|---|---:|---|---|---|---|
| `tool_order.json` | 2.7 KB | 2026-07-25 | `groups`, `panels` | SHIPPED - panel layout | |

## Units (1)

| File | Size | Modified | Top-level keys | My guess | Verdict |
|---|---:|---|---|---|---|
| `project_units.json` | 53.7 KB | 2026-07-25 | `autodesk.spec.aec:temperature-2.0.0`, `autodesk.spec.aec.hvac:velocity-2.0.0`, `autodesk.spec.aec.hvac:density-2.0.0`, `autodesk.spec.aec.structural:reinforcementArea-2.0.0`, ... | USER - snapshot of one project's units | |

## View Organiser (1)

| File | Size | Modified | Top-level keys | My guess | Verdict |
|---|---:|---|---|---|---|
| `view_organiser_config.json` | 0.1 KB | 2026-07-25 | `title_on_sheet_case`, `view_folder_param`, `sheet_folder_param`, `sheet_name_case` | USER - per-user preferences | |

## pyFilter (1)

| File | Size | Modified | Top-level keys | My guess | Verdict |
|---|---:|---|---|---|---|
| `templates/Rebar.json` | 8.6 KB | 2026-07-31 | `created`, `filters`, `name` | USER? - looks user-created (has 'created' key) | |

---

## Things to resolve while classifying

**1. `pytransmit_setup.json` exists twice, with different contents.**

| Path | Keys |
|---|---|
| `pytransmit_setup.json` | `show_method`, `out_pdf`, `recipient_mode`, `page_height_mm`, ... |
| `Settings/pytransmit_setup.json` | `output_path_template`, `projects_older_root`, `projects_root`, `transmittal_naming_template` |

Same filename, different data. Which is live, and is the other dead?

**2. `Studio/studio_layouts/` schemas disagree.**

- `Excel.json` -> `row_sections`, `n_cols`, `page_size_name`, `orientation`, ...
- `PDF.json` -> `name`, `n_rows`, `n_cols`, `row_heights`, ...
- `Revit Drafting View.json` -> `name`, `n_rows`, `n_cols`, `row_heights`, ...
- `Revit Legend.json` -> `name`, `n_rows`, `n_cols`, `row_heights`, ...
- `Revit Schedule.json` -> `name`, `n_rows`, `n_cols`, `row_heights`, ...

`Excel.json` opens with `row_sections`, the others with `name`. Same folder,
different shape - worth knowing whether that is intentional before moving them.

**3. `Layout/Layouts/` vs `Studio/studio_layouts/` share all five filenames**
but hold different schemas, and the Studio ones are far larger (56-114 KB vs
16-33 KB). Two separate systems, or one superseding the other?

**4. `seed43_palette.json` is genuinely mixed** - `dimensions` and `profiles`
are design tokens that should ship, but `active_profile` is the user's
dark/light choice. Splitting may beat classifying.

**5. Largest files** (worth knowing what gets copied around):

- 111 KB  `Seed43.tab/Document Studio.panel/pyTransmit.pushbutton/Studio/studio_layouts/Excel.json`
- 106 KB  `Seed43.tab/Document Studio.panel/pyTransmit.pushbutton/Studio/studio_layouts/PDF.json`
- 54 KB  `Seed43.tab/Document Studio.panel/pyTransmit.pushbutton/Studio/studio_layouts/Revit Drafting View.json`
- 54 KB  `Seed43.tab/Document Studio.panel/pyTransmit.pushbutton/Studio/studio_layouts/Revit Schedule.json`
- 54 KB  `Seed43.tab/Document Studio.panel/pyTransmit.pushbutton/Studio/studio_layouts/Revit Legend.json`

