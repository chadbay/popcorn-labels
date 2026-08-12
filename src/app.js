'use strict';

/* ============================================================
 * Forest Trail Popcorn Labels — browser app
 * Everything runs client-side: the Excel file never leaves
 * this computer.
 * ============================================================ */

const LS_KEY = 'ftPopcornSettingsV1';
const MM2PT = 72 / 25.4;

/* Label sheet geometry — Avery 5163-style, 2" x 4", 10 per sheet.
 * Mirrors the pylabels Specification in the original script:
 * 216x280mm sheet, 2 cols x 5 rows, 101.6x50.8mm labels,
 * 13mm top/bottom margin, 5mm column gap, left margin auto-centered. */
const LAYOUT = {
  sheetW: 216, sheetH: 280,
  cols: 2, rows: 5,
  labelW: 101.6, labelH: 50.8,
  topMargin: 13, columnGap: 5,
};
LAYOUT.leftMargin =
  (LAYOUT.sheetW - (LAYOUT.cols * LAYOUT.labelW + (LAYOUT.cols - 1) * LAYOUT.columnGap)) / 2;
const LABELS_PER_PAGE = LAYOUT.cols * LAYOUT.rows;
const LABEL_W_PT = LAYOUT.labelW * MM2PT; // 288pt
const LABEL_H_PT = LAYOUT.labelH * MM2PT; // 144pt
/* Text baselines measured from the bottom of each label, in points
 * (same positions as the original script: 100 / 60 / 20). */
const BASELINES_PT = [100, 60, 20];

const state = {
  baked: null,          // defaults compiled into this file
  serverDefaults: null, // defaults fetched from settings.json next to the app
  settings: null,       // active settings
  basedOnVersion: null, // which defaults version the active settings came from
  dismissedVersion: null,
  file: null,           // { name, buffer } of the last loaded Excel
  result: null,         // output of buildResult()
};

const $ = (sel) => document.querySelector(sel);
const clone = (o) => JSON.parse(JSON.stringify(o));
const esc = (s) =>
  String(s ?? '').replace(/[&<>"']/g, (c) =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

function defaults() {
  return state.serverDefaults || state.baked;
}

/* ============================ Settings persistence ============================ */

function loadLocal() {
  try {
    const raw = localStorage.getItem(LS_KEY);
    if (!raw) return null;
    const obj = JSON.parse(raw);
    if (!obj || !obj.settings || !Array.isArray(obj.settings.teachers)) return null;
    return obj;
  } catch (e) {
    return null;
  }
}

function saveLocal() {
  try {
    localStorage.setItem(LS_KEY, JSON.stringify({
      settings: state.settings,
      basedOnVersion: state.basedOnVersion,
      dismissedVersion: state.dismissedVersion,
    }));
  } catch (e) { /* private browsing — settings just won't persist */ }
}

function validateSettings(s) {
  if (!s || typeof s !== 'object') return 'Not a settings object';
  if (!Array.isArray(s.teachers)) return 'Missing teacher list';
  if (!Array.isArray(s.staffLabels)) return 'Missing staff labels list';
  if (!s.parsing || !Array.isArray(s.parsing.itemCodes)) return 'Missing parsing settings';
  if (!s.grades || !Array.isArray(s.grades.order)) return 'Missing grade settings';
  return null;
}

function adoptSettings(s, version) {
  state.settings = clone(s);
  state.basedOnVersion = version || s.version || 'unknown';
  saveLocal();
  renderSettingsTab();
  reprocess();
}

function todayISO() {
  const d = new Date();
  const p = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

function exportSettings() {
  const out = clone(state.settings);
  out.version = todayISO();
  const blob = new Blob([JSON.stringify(out, null, 2) + '\n'], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  // Named settings.json on purpose: committing it to the repo is a drop-in update.
  a.download = 'settings.json';
  a.click();
  URL.revokeObjectURL(a.href);
  toast('Settings exported. Committing this file to the repo makes it the new default for everyone.');
}

function importSettingsFile(file) {
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const s = JSON.parse(reader.result);
      const err = validateSettings(s);
      if (err) { toast('Could not import: ' + err, true); return; }
      adoptSettings(s, s.version);
      toast('Settings imported (version ' + (s.version || 'unknown') + ').');
    } catch (e) {
      toast('Could not import: not valid JSON.', true);
    }
  };
  reader.readAsText(file);
}

function b64EncodeUnicode(str) {
  return btoa(unescape(encodeURIComponent(str)));
}
function b64DecodeUnicode(str) {
  return decodeURIComponent(escape(atob(str)));
}

function copyShareLink() {
  const s = clone(state.settings);
  s.version = todayISO();
  const base = location.origin.startsWith('http')
    ? location.origin + location.pathname
    : location.href.split('#')[0];
  const url = base + '#s=' + b64EncodeUnicode(JSON.stringify(s));
  navigator.clipboard.writeText(url).then(
    () => toast('Settings link copied — text or email it to the crew. Opening it loads these exact settings.'),
    () => { prompt('Copy this link:', url); }
  );
}

function checkHashImport() {
  const m = location.hash.match(/^#s=(.+)$/);
  if (!m) return false;
  let s = null;
  try { s = JSON.parse(b64DecodeUnicode(m[1])); } catch (e) { /* fall through */ }
  history.replaceState(null, '', location.pathname + location.search);
  if (!s || validateSettings(s)) return false;
  showBanner(
    `This link contains shared settings (version ${esc(s.version || 'unknown')}, ` +
    `${s.teachers.length} teachers). Load them?`,
    [
      { label: 'Load these settings', primary: true, onClick: () => { adoptSettings(s, s.version); clearBanner(); } },
      { label: 'Ignore', onClick: clearBanner },
    ]
  );
  return true;
}

function maybeShowUpdateBanner() {
  const d = defaults();
  if (!d.version || !state.basedOnVersion) return;
  if (d.version <= state.basedOnVersion) return;
  if (d.version === state.dismissedVersion) return;
  showBanner(
    `An updated roster is available (${esc(d.version)}). You are using settings based on ` +
    `${esc(state.basedOnVersion)}. Switching replaces any local changes you made.`,
    [
      { label: 'Use updated roster', primary: true, onClick: () => { adoptSettings(d, d.version); clearBanner(); } },
      { label: 'Keep my version', onClick: () => { state.dismissedVersion = d.version; saveLocal(); clearBanner(); } },
    ]
  );
}

function resetToDefaults() {
  if (!confirm('Replace your current settings with the school defaults? Local changes will be lost.')) return;
  const d = defaults();
  state.dismissedVersion = null;
  adoptSettings(d, d.version);
  toast('Settings reset to school defaults (version ' + (d.version || 'unknown') + ').');
}

/* ============================ Excel parsing ============================ */

function colLetterToIndex(letter) {
  let n = 0;
  const s = String(letter || '').trim().toUpperCase();
  for (const ch of s) {
    if (ch < 'A' || ch > 'Z') return -1;
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n - 1;
}

function parseWorkbook(buffer) {
  const wb = XLSX.read(buffer, { type: 'array' });
  // Match openpyxl's wb.active: use the workbook's active tab if recorded.
  let idx = 0;
  try {
    const v = wb.Workbook && wb.Workbook.WorkbookViews && wb.Workbook.WorkbookViews[0];
    if (v && typeof v.activeTab === 'number' && wb.SheetNames[v.activeTab]) idx = v.activeTab;
  } catch (e) { /* first sheet */ }
  const ws = wb.Sheets[wb.SheetNames[idx]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  const p = state.settings.parsing;
  const tCol = colLetterToIndex(p.teacherColumn);
  const iCol = colLetterToIndex(p.itemColumn);
  const qCol = colLetterToIndex(p.quantityColumn);
  const codes = p.itemCodes.map((c) => String(c).trim());
  const first = Math.max(1, Number(p.firstDataRow) || 1);

  const orders = new Map(); // teacher name -> student count
  const missingTeacher = { students: 0, rows: [] };
  let matchedRows = 0;

  for (let r = first - 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const item = row[iCol] == null ? '' : String(row[iCol]).trim();
    if (!codes.includes(item)) continue;
    matchedRows++;
    const qty = Math.round(Number(row[qCol])) || 0;
    const teacher = row[tCol] == null ? '' : String(row[tCol]).trim();
    if (!teacher || teacher === 'Unknown') {
      missingTeacher.students += qty;
      missingTeacher.rows.push(r + 1);
    } else {
      orders.set(teacher, (orders.get(teacher) || 0) + qty);
    }
  }
  return { orders, missingTeacher, matchedRows, totalRows: rows.length };
}

/* ============================ Label building ============================ */

function classroomLines(teacher, gradePretty, students, adults) {
  const total = students + adults;
  const word = adults === 1 ? 'Teacher' : 'Teachers';
  return [
    { text: `${gradePretty}: ${teacher.name}`, size: 36 },
    { text: `${students} Students`, size: 24 },
    { text: `+ ${adults} ${word} = ${total} Bags`, size: 24 },
  ];
}

function staffLines(sl) {
  return [
    { text: sl.line1 || '', size: 36 },
    { text: sl.line2 || '', size: sl.smallMiddle ? 16 : 24 },
    { text: sl.line3 || '', size: 36 },
  ];
}

function buildResult(parsed) {
  const s = state.settings;
  const labels = [];
  const unknown = [];
  let classroomBags = 0;
  let classroomStudents = 0;

  const byName = new Map(s.teachers.map((t) => [t.name, t]));

  // Classroom labels: grade order, alphabetical within grade.
  for (const grade of s.grades.order) {
    const inGrade = s.teachers
      .filter((t) => t.grade === grade && parsed.orders.has(t.name))
      .sort((a, b) => a.name.localeCompare(b.name));
    for (const t of inGrade) {
      const students = parsed.orders.get(t.name);
      const adults = Math.max(0, Math.round(Number(t.adults)) || 0);
      classroomBags += students + adults;
      classroomStudents += students;
      labels.push({
        kind: 'classroom',
        lines: classroomLines(t, s.grades.pretty[grade] || grade, students, adults),
      });
    }
  }

  // Teachers in the Excel that aren't in the roster.
  for (const [name, students] of parsed.orders) {
    if (!byName.has(name)) unknown.push({ name, students });
  }

  // Staff labels.
  let staffBags = 0;
  for (const sl of s.staffLabels) {
    const copies = Math.max(0, Math.round(Number(sl.copies)) || 0);
    for (let i = 0; i < copies; i++) labels.push({ kind: 'staff', lines: staffLines(sl) });
    if (sl.bags != null && sl.bags !== '') staffBags += Number(sl.bags) || 0;
  }

  return {
    labels,
    unknown,
    missingTeacher: parsed.missingTeacher,
    matchedRows: parsed.matchedRows,
    classrooms: labels.filter((l) => l.kind === 'classroom').length,
    classroomStudents,
    classroomBags,
    staffBags,
    pages: Math.ceil(labels.length / LABELS_PER_PAGE) || 0,
  };
}

/* ============================ PDF generation ============================ */

function fittedSize(doc, text, baseSize, maxWidthPt) {
  // Shrink the font until the line fits the label (the original script let long
  // names run off the edge; we shrink instead).
  let size = baseSize;
  while (size > 8 && doc.getStringUnitWidth(text) * size > maxWidthPt) size -= 1;
  return size;
}

function generatePdf() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({
    unit: 'pt',
    format: [LAYOUT.sheetW * MM2PT, LAYOUT.sheetH * MM2PT],
    compress: true,
  });
  doc.setFont('helvetica', 'normal');

  const labels = state.result.labels;
  const maxTextWidth = LABEL_W_PT - 12; // small safety margin on each side

  labels.forEach((label, i) => {
    const pos = i % LABELS_PER_PAGE;
    if (i > 0 && pos === 0) doc.addPage();
    const row = Math.floor(pos / LAYOUT.cols);
    const col = pos % LAYOUT.cols;
    const xCenter = (LAYOUT.leftMargin + col * (LAYOUT.labelW + LAYOUT.columnGap) + LAYOUT.labelW / 2) * MM2PT;
    const yTop = (LAYOUT.topMargin + row * LAYOUT.labelH) * MM2PT;

    label.lines.forEach((line, li) => {
      if (!line.text) return;
      const size = fittedSize(doc, line.text, line.size, maxTextWidth);
      doc.setFontSize(size);
      // jsPDF y is the baseline measured from the page top.
      const y = yTop + (LABEL_H_PT - BASELINES_PT[li]);
      doc.text(line.text, xCenter, y, { align: 'center' });
    });
  });

  return doc;
}

function pdfFileName() {
  const d = new Date();
  const p = (n) => String(n).padStart(2, '0');
  return `popcorn_${p(d.getMonth() + 1)}_${p(d.getDate())}_${String(d.getFullYear()).slice(-2)}.pdf`;
}

function downloadPdf() {
  if (!state.result || !state.result.labels.length) return;
  generatePdf().save(pdfFileName());
}

/* ============================ Processing pipeline ============================ */

function processBuffer(buffer, name) {
  state.file = { name, buffer };
  reprocess();
}

function reprocess() {
  if (!state.file) { renderOutput(); return; }
  let parsed;
  try {
    parsed = parseWorkbook(state.file.buffer);
  } catch (e) {
    state.result = null;
    renderOutput('Could not read that file as an Excel spreadsheet. Make sure it is the weekly .xlsx order export.');
    return;
  }
  state.result = buildResult(parsed);
  renderOutput();
}

function handleFiles(files) {
  const file = files && files[0];
  if (!file) return;
  if (!/\.(xlsx|xlsm|xls)$/i.test(file.name)) {
    renderOutput('That does not look like an Excel file (.xlsx). Drop the weekly order export here.');
    return;
  }
  const reader = new FileReader();
  reader.onload = () => processBuffer(reader.result, file.name);
  reader.readAsArrayBuffer(file);
}

/* ============================ Rendering: banners & toast ============================ */

function showBanner(text, buttons) {
  const el = $('#banner');
  el.innerHTML = `<span>${text}</span><span class="banner-buttons"></span>`;
  const btns = el.querySelector('.banner-buttons');
  for (const b of buttons) {
    const btn = document.createElement('button');
    btn.textContent = b.label;
    btn.className = b.primary ? 'btn primary small' : 'btn small';
    btn.onclick = b.onClick;
    btns.appendChild(btn);
  }
  el.hidden = false;
}
function clearBanner() { const el = $('#banner'); el.hidden = true; el.innerHTML = ''; }

let toastTimer = null;
function toast(msg, isError) {
  const el = $('#toast');
  el.textContent = msg;
  el.className = isError ? 'toast error' : 'toast';
  el.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { el.hidden = true; }, 6000);
}

/* ============================ Rendering: output ============================ */

function renderOutput(errorMsg) {
  const status = $('#fileStatus');
  const warnings = $('#warnings');
  const totals = $('#totals');
  const preview = $('#preview');
  const dlBtn = $('#downloadBtn');

  warnings.innerHTML = '';
  totals.innerHTML = '';
  preview.innerHTML = '';
  dlBtn.hidden = true;

  if (errorMsg) {
    status.innerHTML = `<div class="warn error">${esc(errorMsg)}</div>`;
    return;
  }
  if (!state.file) { status.innerHTML = ''; return; }

  const r = state.result;
  status.innerHTML = `<div class="file-ok">Loaded <strong>${esc(state.file.name)}</strong></div>`;

  if (r.matchedRows === 0) {
    warnings.innerHTML =
      `<div class="warn error">No popcorn orders found in this file. ` +
      `Check that this is the right export — or that the item codes under ` +
      `Roster &amp; Settings → Advanced (${esc(state.settings.parsing.itemCodes.join(', '))}) still match.</div>`;
    return;
  }

  // Unknown teachers — offer to add them on the spot.
  for (const u of r.unknown) {
    const div = document.createElement('div');
    div.className = 'warn';
    const gradeOpts = state.settings.grades.order
      .map((g) => `<option value="${esc(g)}">${esc(state.settings.grades.pretty[g] || g)}</option>`).join('');
    div.innerHTML =
      `<strong>${esc(u.name)}</strong> has ${u.students} student${u.students === 1 ? '' : 's'} ordering popcorn ` +
      `but is not in the teacher roster, so no label was made. Add them: ` +
      `<label>Grade <select data-u="grade">${gradeOpts}</select></label> ` +
      `<label>Adults in room <input data-u="adults" type="number" min="1" max="9" value="1" style="width:3.5em"></label> ` +
      `<button class="btn small primary">Add to roster</button>`;
    div.querySelector('button').onclick = () => {
      state.settings.teachers.push({
        name: u.name,
        grade: div.querySelector('[data-u=grade]').value,
        adults: Number(div.querySelector('[data-u=adults]').value) || 1,
      });
      saveLocal();
      renderSettingsTab();
      reprocess();
      toast(`${u.name} added to the roster. Remember to Export settings if this change should be shared.`);
    };
    warnings.appendChild(div);
  }

  if (r.missingTeacher.students > 0) {
    const div = document.createElement('div');
    div.className = 'warn';
    div.innerHTML =
      `<strong>${r.missingTeacher.students} student${r.missingTeacher.students === 1 ? '' : 's'}</strong> ` +
      `(spreadsheet row${r.missingTeacher.rows.length === 1 ? '' : 's'} ${esc(r.missingTeacher.rows.join(', '))}) ` +
      `have no teacher specified in the export. They are <strong>not</strong> on any label — ` +
      `fix the export or hand-adjust a label after printing.`;
    warnings.appendChild(div);
  }

  const totalKnown = r.classroomBags + r.staffBags;
  totals.innerHTML =
    `<div class="totals-card">` +
    `<div><span class="big">${r.labels.length}</span> labels</div>` +
    `<div><span class="big">${r.pages}</span> page${r.pages === 1 ? '' : 's'}</div>` +
    `<div><span class="big">${r.classrooms}</span> classrooms</div>` +
    `<div><span class="big">${r.classroomBags}</span> classroom bags<br><small>${r.classroomStudents} students + ${r.classroomBags - r.classroomStudents} adults</small></div>` +
    `<div><span class="big">${r.staffBags}</span> staff bags</div>` +
    `<div><span class="big">≈ ${totalKnown}</span> bags to pop</div>` +
    `</div>`;

  dlBtn.hidden = false;
  renderPreview();
}

function renderPreview() {
  const preview = $('#preview');
  preview.innerHTML = '<h3>Preview</h3>';
  const labels = state.result.labels;
  const SCALE = 2.1; // px per mm

  for (let p = 0; p < state.result.pages; p++) {
    const page = document.createElement('div');
    page.className = 'page';
    page.style.width = LAYOUT.sheetW * SCALE + 'px';
    page.style.height = LAYOUT.sheetH * SCALE + 'px';

    for (let pos = 0; pos < LABELS_PER_PAGE; pos++) {
      const i = p * LABELS_PER_PAGE + pos;
      if (i >= labels.length) break;
      const row = Math.floor(pos / LAYOUT.cols);
      const col = pos % LAYOUT.cols;
      const box = document.createElement('div');
      box.className = 'label-box';
      box.style.left = (LAYOUT.leftMargin + col * (LAYOUT.labelW + LAYOUT.columnGap)) * SCALE + 'px';
      box.style.top = (LAYOUT.topMargin + row * LAYOUT.labelH) * SCALE + 'px';
      box.style.width = LAYOUT.labelW * SCALE + 'px';
      box.style.height = LAYOUT.labelH * SCALE + 'px';

      labels[i].lines.forEach((line, li) => {
        if (!line.text) return;
        const div = document.createElement('div');
        div.className = 'label-line';
        // Same proportions as the PDF: baseline from label bottom, font in pt of a 144pt-high label.
        const fontPx = (line.size / LABEL_H_PT) * LAYOUT.labelH * SCALE;
        const baselinePx = (BASELINES_PT[li] / LABEL_H_PT) * LAYOUT.labelH * SCALE;
        div.style.fontSize = fontPx + 'px';
        div.style.bottom = baselinePx - fontPx * 0.2 + 'px';
        div.textContent = line.text;
        box.appendChild(div);
      });
      page.appendChild(box);
    }
    const wrap = document.createElement('div');
    wrap.className = 'page-wrap';
    wrap.innerHTML = `<div class="page-num">Page ${p + 1} of ${state.result.pages}</div>`;
    wrap.prepend(page);
    preview.appendChild(wrap);
  }
}

/* ============================ Rendering: settings tab ============================ */

function renderSettingsTab() {
  renderVersionInfo();
  renderRoster();
  renderStaffLabels();
  renderAdvanced();
}

function renderVersionInfo() {
  const d = defaults();
  $('#versionInfo').innerHTML =
    `Your settings are based on version <strong>${esc(state.basedOnVersion || '?')}</strong>` +
    (state.serverDefaults
      ? ` · school defaults on the website: <strong>${esc(d.version || '?')}</strong>`
      : ` · <em>could not check the website for newer defaults (offline?)</em>`);
}

function gradeSelect(value) {
  const s = state.settings;
  return `<select data-f="grade">` + s.grades.order.map((g) =>
    `<option value="${esc(g)}" ${g === value ? 'selected' : ''}>${esc(s.grades.pretty[g] || g)}</option>`
  ).join('') + `</select>`;
}

function renderRoster() {
  const s = state.settings;
  const tbody = $('#rosterBody');
  tbody.innerHTML = '';

  const sorted = s.teachers
    .map((t, idx) => ({ t, idx }))
    .sort((a, b) =>
      s.grades.order.indexOf(a.t.grade) - s.grades.order.indexOf(b.t.grade) ||
      a.t.name.localeCompare(b.t.name));

  let lastGrade = null;
  for (const { t, idx } of sorted) {
    if (t.grade !== lastGrade) {
      lastGrade = t.grade;
      const tr = document.createElement('tr');
      tr.className = 'grade-row';
      tr.innerHTML = `<td colspan="4">${esc(s.grades.pretty[t.grade] || t.grade)} grade</td>`;
      tbody.appendChild(tr);
    }
    const tr = document.createElement('tr');
    tr.innerHTML =
      `<td><input data-f="name" value="${esc(t.name)}"></td>` +
      `<td>${gradeSelect(t.grade)}</td>` +
      `<td><input data-f="adults" type="number" min="0" max="9" value="${esc(t.adults)}" class="num"></td>` +
      `<td><button class="btn small danger" title="Remove">✕</button></td>`;
    tr.querySelector('[data-f=name]').onchange = (e) => { s.teachers[idx].name = e.target.value.trim(); settingsEdited(); };
    tr.querySelector('[data-f=grade]').onchange = (e) => { s.teachers[idx].grade = e.target.value; settingsEdited(); };
    tr.querySelector('[data-f=adults]').onchange = (e) => { s.teachers[idx].adults = Math.max(0, Number(e.target.value) || 0); settingsEdited(); };
    tr.querySelector('button').onclick = () => {
      if (!confirm(`Remove ${t.name} from the roster?`)) return;
      s.teachers.splice(idx, 1);
      settingsEdited();
    };
    tbody.appendChild(tr);
  }

  $('#addTeacherBtn').onclick = () => {
    const name = $('#newTeacherName').value.trim();
    if (!name) { toast('Type the teacher’s name first.', true); return; }
    if (s.teachers.some((t) => t.name === name)) { toast('That teacher is already in the roster.', true); return; }
    s.teachers.push({
      name,
      grade: $('#newTeacherGrade').value,
      adults: Math.max(0, Number($('#newTeacherAdults').value) || 1),
    });
    $('#newTeacherName').value = '';
    settingsEdited();
  };
}

function renderStaffLabels() {
  const s = state.settings;
  const tbody = $('#staffBody');
  tbody.innerHTML = '';

  s.staffLabels.forEach((sl, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML =
      `<td><input data-f="line1" value="${esc(sl.line1)}"></td>` +
      `<td><input data-f="line2" value="${esc(sl.line2)}"></td>` +
      `<td><input data-f="line3" value="${esc(sl.line3)}"></td>` +
      `<td><input data-f="copies" type="number" min="0" max="9" value="${esc(sl.copies)}" class="num"></td>` +
      `<td><input data-f="smallMiddle" type="checkbox" ${sl.smallMiddle ? 'checked' : ''}></td>` +
      `<td><input data-f="bags" type="number" min="0" value="${sl.bags == null ? '' : esc(sl.bags)}" class="num" placeholder="—"></td>` +
      `<td class="row-actions">` +
      `<button class="btn small" data-a="up" title="Move up" ${idx === 0 ? 'disabled' : ''}>↑</button>` +
      `<button class="btn small" data-a="down" title="Move down" ${idx === s.staffLabels.length - 1 ? 'disabled' : ''}>↓</button>` +
      `<button class="btn small danger" data-a="del" title="Remove">✕</button></td>`;

    const bind = (f, fn) => { tr.querySelector(`[data-f=${f}]`).onchange = (e) => { fn(e); settingsEdited(); }; };
    bind('line1', (e) => { sl.line1 = e.target.value; });
    bind('line2', (e) => { sl.line2 = e.target.value; });
    bind('line3', (e) => { sl.line3 = e.target.value; });
    bind('copies', (e) => { sl.copies = Math.max(0, Number(e.target.value) || 0); });
    bind('smallMiddle', (e) => { sl.smallMiddle = e.target.checked; });
    bind('bags', (e) => { sl.bags = e.target.value === '' ? null : Math.max(0, Number(e.target.value) || 0); });

    tr.querySelector('[data-a=up]').onclick = () => { s.staffLabels.splice(idx - 1, 0, s.staffLabels.splice(idx, 1)[0]); settingsEdited(); };
    tr.querySelector('[data-a=down]').onclick = () => { s.staffLabels.splice(idx + 1, 0, s.staffLabels.splice(idx, 1)[0]); settingsEdited(); };
    tr.querySelector('[data-a=del]').onclick = () => {
      if (!confirm(`Remove the "${sl.line1}" label?`)) return;
      s.staffLabels.splice(idx, 1);
      settingsEdited();
    };
    tbody.appendChild(tr);
  });

  $('#addStaffBtn').onclick = () => {
    s.staffLabels.push({ line1: 'New label', line2: '', line3: '', copies: 1, smallMiddle: false, bags: null });
    settingsEdited();
  };
}

function renderAdvanced() {
  const p = state.settings.parsing;
  $('#advItemCodes').value = p.itemCodes.join(', ');
  $('#advTeacherCol').value = p.teacherColumn;
  $('#advItemCol').value = p.itemColumn;
  $('#advQtyCol').value = p.quantityColumn;
  $('#advFirstRow').value = p.firstDataRow;

  $('#advItemCodes').onchange = (e) => {
    p.itemCodes = e.target.value.split(',').map((c) => c.trim()).filter(Boolean);
    settingsEdited();
  };
  const bindCol = (id, key) => {
    $(id).onchange = (e) => {
      const v = e.target.value.trim().toUpperCase();
      if (colLetterToIndex(v) < 0) { toast('Column must be a letter like A, C, or G.', true); e.target.value = p[key]; return; }
      p[key] = v;
      settingsEdited();
    };
  };
  bindCol('#advTeacherCol', 'teacherColumn');
  bindCol('#advItemCol', 'itemColumn');
  bindCol('#advQtyCol', 'quantityColumn');
  $('#advFirstRow').onchange = (e) => {
    p.firstDataRow = Math.max(1, Math.round(Number(e.target.value)) || 1);
    settingsEdited();
  };
}

function settingsEdited() {
  saveLocal();
  renderSettingsTab();
  reprocess();
}

/* ============================ Tabs, dropzone, init ============================ */

function setupTabs() {
  document.querySelectorAll('nav [data-tab]').forEach((btn) => {
    btn.onclick = () => {
      document.querySelectorAll('nav [data-tab]').forEach((b) => b.classList.toggle('active', b === btn));
      document.querySelectorAll('main > section').forEach((sec) => {
        sec.hidden = sec.id !== 'tab-' + btn.dataset.tab;
      });
    };
  });
}

function setupDropzone() {
  const dz = $('#dropzone');
  const input = $('#fileInput');
  dz.onclick = () => input.click();
  input.onchange = () => { handleFiles(input.files); input.value = ''; };
  dz.ondragover = (e) => { e.preventDefault(); dz.classList.add('drag'); };
  dz.ondragleave = () => dz.classList.remove('drag');
  dz.ondrop = (e) => {
    e.preventDefault();
    dz.classList.remove('drag');
    handleFiles(e.dataTransfer.files);
  };
}

async function init() {
  state.baked = window.BAKED_DEFAULTS;

  setupTabs();
  setupDropzone();
  $('#downloadBtn').onclick = downloadPdf;
  $('#exportBtn').onclick = exportSettings;
  $('#importBtn').onclick = () => $('#importInput').click();
  $('#importInput').onchange = (e) => { if (e.target.files[0]) importSettingsFile(e.target.files[0]); e.target.value = ''; };
  $('#shareBtn').onclick = copyShareLink;
  $('#resetBtn').onclick = resetToDefaults;

  // Layer 1: settings.json committed next to the app (source of truth for defaults).
  try {
    const resp = await fetch('settings.json', { cache: 'no-store' });
    if (resp.ok) {
      const s = await resp.json();
      if (!validateSettings(s)) state.serverDefaults = s;
    }
  } catch (e) { /* offline or file:// — baked defaults cover this */ }

  // Layer 2: this volunteer's saved settings.
  const local = loadLocal();
  if (local) {
    state.settings = local.settings;
    state.basedOnVersion = local.basedOnVersion;
    state.dismissedVersion = local.dismissedVersion || null;
  } else {
    const d = defaults();
    state.settings = clone(d);
    state.basedOnVersion = d.version;
  }

  renderSettingsTab();

  // A shared-settings link takes priority over the update banner.
  if (!checkHashImport()) maybeShowUpdateBanner();
}

document.addEventListener('DOMContentLoaded', init);

/* Hooks for automated testing — not used by the UI. */
window.__popcorn = {
  processArrayBuffer: (buf, name) => processBuffer(buf, name),
  getState: () => state,
  getSummary: () => state.result && {
    labels: state.result.labels.length,
    pages: state.result.pages,
    classrooms: state.result.classrooms,
    classroomBags: state.result.classroomBags,
    staffBags: state.result.staffBags,
    unknown: state.result.unknown,
    missingTeacher: state.result.missingTeacher,
    firstLabels: state.result.labels.slice(0, 6).map((l) => l.lines.map((x) => x.text)),
  },
  generatePdfBase64: () => generatePdf().output('datauristring').split(',')[1],
};
