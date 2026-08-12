# Popcorn Labels — Modernization Plan

## Background

Forest Trail Elementary's popcorn volunteer crew prints classroom bag labels weekly.
Today this is done with [popcorn_labels.py](https://github.com/mkreddy007/PopcornScripts/blob/main/popcorn_labels.py),
a Python script that reads an Excel order export and produces a PDF of labels
(2" × 4", 10 per letter sheet — Avery 5163-style).

### How the current script works

1. Scans `../Downloads/` for any `.xlsx` file and loads it (last one found wins).
2. Reads rows from row 3: teacher last name (col A), item code (col C, must be
   `SE-POPCORN` or `SE-POPCORN-SPRING-ONLY`), student count (col G). Sums per teacher.
3. Emits one label per classroom grouped K→5: `3rd: Chamness` / `22 Students` /
   `+ 2 Teachers = 24 Bags`. Bag total = students + adults (adults come from
   hardcoded `TwoTeachers` / `ThreeTeachers` lists; default 1).
4. Appends 13 hardcoded staff labels (Office ×2, Library ×2, Bus Drivers ×2,
   Learning Lab ×2, Custodial ×2, CDC Staff, Cafeteria, Specials) with fixed bag counts.
5. Saves `popcorn_MM_DD_YY.pdf`.

### Pain points

- Teacher roster, grade mapping, helper counts, and staff labels are all **hardcoded
  in Python source** — staffing changes require editing code.
- Requires Python + `pylabels`, `reportlab`, `openpyxl` installed, run from a terminal.
- Fragile: wrong file picked if Downloads has multiple `.xlsx`; crashes (`NameError`)
  if none; crashes (`KeyError`) mid-run if the Excel contains a teacher not in the
  hardcoded dict (e.g. a new hire).

## Recommendation: single-file browser app

One self-contained `index.html` — hosted on GitHub Pages and also usable by
double-clicking the file locally. Excel parsing (SheetJS) and PDF generation
(pdf-lib or jsPDF) run entirely client-side; order data never leaves the machine.
Works on Mac/Windows/Chromebook/iPad with zero installation.

### Weekly workflow (after)

1. Open the bookmarked page.
2. Drag the weekly Excel export onto it.
3. Review the on-screen preview, click **Download PDF**, print.

### Settings screen (replaces all hardcoded config)

- **Teacher roster table**: name, grade (K–5), adults in classroom (number —
  replaces TwoTeachers/ThreeTeachers). Add/edit/remove rows.
- **Staff labels table**: title, second line, bag count, copies. Fully editable.
- **Advanced**: popcorn item codes, Excel column positions, header row count,
  label sheet dimensions.
- Settings auto-save to the browser (localStorage), with **Export settings** /
  **Import settings** buttons producing a small `.json` file so the crew can share
  one config. A shareable settings **link** (config encoded in the URL) is a
  nice-to-have for texting changes to the crew.

### Settings persistence & update flow

Settings live in three layers, from most to least authoritative for defaults:

1. **`settings.json` committed to the repo** — the source of truth. Because
   GitHub Pages serves the app from the same repo, the page fetches this file on
   load. Updating the roster = exporting settings from the app and committing
   the file (maintainer does this occasionally). Every volunteer's browser picks
   up committed changes automatically — no imports needed. Bonus: git history of
   this file becomes a year-over-year record of roster/staff-label changes.
2. **localStorage** — a volunteer's in-season local tweaks, layered on top of the
   repo defaults. Per-browser, per-device.
3. **Baked-in snapshot in the HTML** — last-resort fallback so the app still works
   when run as a downloaded local file (where fetching a sibling file is blocked)
   or when the fetch fails.

Exports carry a date/version stamp. When the repo `settings.json` is newer than
the version a volunteer's local tweaks were based on, the app shows a banner —
"Updated roster available (Aug 2026) — use it?" — rather than silently clobbering
or ignoring either side.

**Maintenance story:** day-to-day volunteers touch nothing; mid-year tweaks happen
in the app's settings screen (shared via export file or link); durable changes are
one occasional `settings.json` commit by the maintainer.

### Guardrails (new)

- Unknown teacher in the Excel → flagged with an "add to roster" prompt, not a crash.
- Rows with missing/"Unknown" teacher → surfaced with student counts so kids
  aren't silently skipped.
- Totals summary: total bags + label count, so the popping crew knows volume.

## Build phases

1. **Core engine** — parse Excel, reproduce current PDF output exactly
   (layout, Helvetica 36/24pt, 2 cols × 5 rows of 101.6 × 50.8 mm labels on
   letter), current hardcoded values as defaults.
2. **Settings UI** — roster editor, staff-labels editor, export/import config,
   `settings.json` fetch-with-fallback and "newer roster available" banner.
3. **Preview + validation** — on-screen label preview, warnings, totals,
   shareable settings link.
4. **Ship** — GitHub Pages deployment (app + `settings.json` in the repo) +
   one-page illustrated volunteer instructions.

## Alternatives considered

- **Keep Python, move config to a spreadsheet/CSV**: still requires installing
  Python (the repo README calls installation the hardest part). Rejected.
- **Google Sheets + Apps Script**: config-in-a-sheet is friendly, but PDF label
  positioning control is weak and it ties the tool to one Google account. Rejected.

## Open items

- Get a sample weekly Excel export (scrubbed is fine) to verify column layout
  against the script's assumptions (data from row 3; teacher=A, item=C, qty=G).
- Confirm the label stock is Avery 5163 (2" × 4", 10/sheet) so the PDF template
  can be named accordingly in the UI.
