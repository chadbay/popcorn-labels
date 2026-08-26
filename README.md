# 🍿 Forest Trail Popcorn Labels

A zero-install web app that turns the weekly popcorn order files into
print-ready PDFs. It replaces the original
[popcorn_labels.py](popcorn_labels_original.py) script so that no Python,
terminal, or code editing is needed. Two outputs:

- **Bag labels** (from the weekly Excel order export) — Avery 5163/5963,
  2″ × 4″, 10 per sheet. One per classroom plus the fixed staff labels.
- **Roster labels** (from the Booster Club *Packing List Report* PDF) — the
  combined label+roster format: grade/teacher header, `TOTAL BAGS`, the
  participating students' names, `N STUDENTS/M TEACHER(S)`, and a footer.
  Printed on full-sheet 8.5″ × 11″ label stock, three homeroom strips per page
  (two cuts along the printed dashed guides), staff labels two-across after.

Everything runs inside the browser — neither file is uploaded anywhere.

## For volunteers (weekly)

1. Open the app page (bookmark it).
2. Drag the weekly files onto the dashed box — the `.xlsx` order export and/or
   the Packing List Report `.pdf` (both at once is fine). Either file alone
   works: the Excel makes bag labels; the packing list makes both.
3. Read any yellow warnings (new teachers, kids without a homeroom, or the two
   files disagreeing on a room's count). An order with no homeroom
   ("UNSPECIFIED") can be filed on the spot: type the student's name, pick the
   room, and the app remembers the fix-up and applies it automatically every
   week until the Booster Club database is corrected. Fix-ups live in the
   browser (and in share links) but are never included in exported
   settings.json, so student names stay out of this public repo.
4. Download and print at **Actual size / 100%** (never "Fit to page").

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
| `vendor/pdf.min.js` + `vendor/pdf.worker.min.js` | pdf.js 3.11.174 (packing-list PDF parsing; worker inlined as a blob) |

```bash
node build.js
```

inlines everything into a single self-contained `index.html` (~2.6 MB) that
also works opened straight from disk (e.g. off a USB stick).

### Testing

`node scripts/make-sample.js` regenerates [sample_orders.xlsx](sample_orders.xlsx),
a fake weekly export that exercises the edge cases: a teacher with orders under
two item codes (quantities must sum), a teacher missing from the roster, rows
with no teacher specified, and a non-popcorn item row (must be ignored).
`python3 scripts/make-sample-packing.py` (needs reportlab) regenerates
[sample_packing_list.pdf](sample_packing_list.pdf), a count-matched fake
Packing List Report with invented student names — no real student data lives
in this repo. Drop both on the app and check the warnings and totals.

The bag-label PDF was verified against the original Python script's output for
the same input: identical text, positions (baselines exact, centering within
0.5 pt), fonts, and page count — and against a real production run (the
packing-list-only path reproduced that week's labels entry-for-entry). One
deliberate difference: classroom labels are sorted alphabetically within each
grade instead of by order of appearance in the Excel.

### Label geometry

Mirrors the original pylabels spec: 216 × 280 mm sheet, 2 columns × 5 rows of
101.6 × 50.8 mm labels, 13 mm top margin, 5 mm column gap, centered horizontally.
Text baselines sit 100 / 60 / 20 pt from each label's bottom edge, Helvetica
36 / 24 (or 16) pt, auto-shrunk when a line is too wide to fit.

## Hosting

Designed for GitHub Pages: serve this repo's root (only `index.html` and
`settings.json` are needed at runtime). Any static host works.
