# Seed43 JSON Inventory

Every `.json` under `Seed43.tab`, `lib` and `.user`, and whether it ships or belongs
to the user.

Updated 2026-08-06. **38 files: 4 git-tracked, 34 in `.user/`.**

When this was first written (2026-08-05) there were 36 files and *all* were tracked, so
every one shipped as a default and was then overwritten at runtime. That is no longer
true. The migration is done bar one file.

| Verdict | Meaning | Where it lives |
|---|---|---|
| `USER` | personal, changes as they work | `.user/<Tool>/`, not tracked |
| `SHIPPED` | app data, users never edit | tracked, beside the tool or in `lib` |
| `DEFAULT` | shipped seed, copied into `.user` on first run | tracked, in a `defaults/` folder |
| `MIXED` | both — needs splitting | see "Still open" |

---

## Still tracked (4)

| File | Size | Modified | Top-level keys | Verdict |
|---|---:|---|---|---|
| `lib/Snippets/_icons.json` | 19.6 KB | 2026-07-25 | `about`, `arrow_down`, `arrow_left`, ... | `SHIPPED` — icon vector data |
| `lib/Snippets/seed43_palette.json` | 5.2 KB | 2026-08-06 | `dimensions`, `active_profile`, `profiles` | **`MIXED`** — see below |
| `Seed43.tab/.../About.pushbutton/tool_order.json` | 2.7 KB | 2026-07-25 | `groups`, `panels` | `SHIPPED` — panel layout |
| `Seed43.tab/.../pyFilter.pushbutton/defaults/Rebar.json` | 8.6 KB | 2026-07-31 | `created`, `filters`, `name` | `DEFAULT` — seeds `.user/pyFilter/templates/` |

---

## Moved to `.user` (34)

All untracked, so an update can never overwrite them.

### pyTransmit (20)

`Layouts/` (5), `studio_layouts/` (5), `Settings/` (8), plus `layout_config.json`,
`pytransmit_setup.json`, `pytransmit_sync.json` and `studio_config.json`.
The two largest files in the extension live here: `studio_layouts/Excel.json` (111 KB)
and `studio_layouts/PDF.json` (106 KB).

### pySheets (6)

| File | Size | Modified | Top-level keys |
|---|---:|---|---|
| `profiles/test.json` | 2.8 KB | 2026-08-06 | `img`, `nwc`, `view_sel`, `ifc`, ... and `schedule` |
| `settings/scheduled_print.json` | 0.1 KB | 2026-08-06 | `version`, `entries`, `grace_minutes`, `heartbeat_doc` |
| `settings/lastsession.json` | 1.8 KB | 2026-08-06 | `formats`, `pdf`, `auto_overwrite`, `column_order` |
| `settings/custom_columns.json` | 0.1 KB | 2026-07-25 | `sheet_columns`, `builtin_visible`, `view_columns` |
| `settings/naming_memory.json` | 0.2 KB | 2026-08-06 | `per_format` |
| `folder_presets/folder_presets.json` | 0.2 KB | 2026-07-25 | `Project` |

Two of these are new since the last revision, and both belong to scheduled printing:

- **`scheduled_print.json`** is runtime state only — which profiles are armed, the
  document each was armed against, and when each is next due. Read by `startup.py` as
  well as pySheets, via `lib/Snippets/_schedule.py`. Version `2`; a v1 file (a single
  flat schedule) is upgraded on read.
- **A `schedule` block inside each profile** holds that card's timing: hour, minute,
  repeat mode, weekdays, start date. It lives on the profile so it travels with it.
  Anything writing a profile must preserve this key — `_gather_profile()` does not
  know about it.

### pyLink (3), pyFilter (1), Units (1), View Organiser (1)

`excel_font_settings.json`, `section_groups.json`, `word_text_settings.json`;
`templates/Rebar.json`; `project_units.json` (53.7 KB); `view_organiser_config.json`.

---

## Resolved since the last revision

**Duplicate `pytransmit_setup.json`.** Both copies still exist, now at
`.user/pyTransmit/pytransmit_setup.json` (window state: `show_from`, `show_method`,
`out_schedule`) and `.user/pyTransmit/Settings/pytransmit_setup.json` (paths and
naming templates). Different data, so both are live — but the shared filename is still
a trap for anyone reading the code.

**`pyFilter/Rebar.json`.** Now correctly split: the shipped copy is a `DEFAULT` under
`defaults/`, seeded into `.user/pyFilter/templates/` on first run. This is the pattern
the other `MIXED` cases should follow.

**`Layout/Layouts/` vs `Studio/studio_layouts/`.** Both moved to `.user/pyTransmit/`
and kept their separate schemas, so they are two systems, not one superseding the
other. `Excel.json` still opens with `row_sections` where the others open with `name`.

---

## Still open

**`seed43_palette.json` is the last mixed file, and the only one an update can still
clobber.** `dimensions` and `profiles` are design tokens that should ship, but
`active_profile` is the user's dark/light choice, and `set_accent()` rewrites the
colour values in place. So a user picking an accent dirties a tracked file, and an
update overwrites their choice.

Splitting still beats classifying: ship the tokens, keep the user's `active_profile`
and accent override in `.user/`.

Worth knowing when weighing that up: the icon renderer
(`lib/Snippets/_svg_icons.py`) reads this file to recolour `lib/icons/*.svg` into each
tool's `icon.png`, and treats the palette's timestamp as a staleness signal. Splitting
the file means deciding which half the renderer watches.
