# 🍿 Forest Trail Popcorn Labels

A zero-install web app that turns the weekly popcorn order Excel export into a
print-ready PDF of bag labels (Avery 5163 — 2″ × 4″, 10 per sheet). It replaces
the original [popcorn_labels.py](popcorn_labels_original.py) script so that no
Python, terminal, or code editing is needed.

Everything runs inside the browser — the Excel file is never uploaded anywhere.

## For volunteers (weekly)

1. Open the app page (bookmark it).
2. Drag the weekly `.xlsx` order export onto the dashed box.
3. Read any yellow warnings (new teachers, kids without a teacher listed).
4. Click **Download PDF for printing** and print on Avery 5163 stock at
   **Actual size / 100%** (never "Fit to page").

The **Help** tab inside the app has the full instructions, including how to
handle roster changes.

## For the maintainer

### How settings work

- [settings.json](settings.json) in this repo is the **source of truth** for the
  teacher roster, staff labels, and Excel parsing rules. The app fetches it on
  page load, so **committing an updated settings.json updates every volunteer's
  browser automatically.**
- Volunteers' in-season edits live in their browser (localStorage) on top of
  those defaults. When the committed version is newer than what their edits were
  based on, the app shows an "updated roster available" banner.
- A snapshot of settings is also baked into `index.html` as an offline fallback.

### Updating the roster for a new school year

1. Open the app, go to **Roster & Settings**, make the changes.
2. Click **Export settings** — it downloads a ready-to-commit `settings.json`
   (version-stamped with today's date).
3. Replace `settings.json` in this repo with it and commit.
4. Optionally run `node build.js` and commit `index.html` too, so the offline
   fallback matches (not required for the hosted app to pick up the change).

### Building the app

`index.html` is generated — don't edit it by hand. Sources:

| File | Purpose |
|---|---|
| `src/template.html` | Page structure, styles, Help content |
| `src/app.js` | All application logic |
| `settings.json` | Default roster / labels / parsing config |
| `vendor/xlsx.full.min.js` | SheetJS 0.18.5 (Excel parsing) |
| `vendor/jspdf.umd.min.js` | jsPDF 2.5.1 (PDF generation) |

```bash
node build.js
```

inlines everything into a single self-contained `index.html` (~1.3 MB) that
also works opened straight from disk (e.g. off a USB stick).

### Testing

`node scripts/make-sample.js` regenerates [sample_orders.xlsx](sample_orders.xlsx),
a fake weekly export that exercises the edge cases: a teacher with orders under
two item codes (quantities must sum), a teacher missing from the roster, rows
with no teacher specified, and a non-popcorn item row (must be ignored).
Drop it on the app and check the warnings and totals.

The PDF output was verified against the original Python script's output for the
same input: identical text, positions (baselines exact, centering within 0.5 pt),
fonts, and page count. One deliberate difference: classroom labels are sorted
alphabetically within each grade instead of by order of appearance in the Excel.

### Label geometry

Mirrors the original pylabels spec: 216 × 280 mm sheet, 2 columns × 5 rows of
101.6 × 50.8 mm labels, 13 mm top margin, 5 mm column gap, centered horizontally.
Text baselines sit 100 / 60 / 20 pt from each label's bottom edge, Helvetica
36 / 24 (or 16) pt, auto-shrunk when a line is too wide to fit.

## Hosting

Designed for GitHub Pages: serve this repo's root (only `index.html` and
`settings.json` are needed at runtime). Any static host works.
