/* ============================================================
   Excel Power Toolbar — Office JS Add-in
   Works in Excel Web (Office 365) & Excel Desktop
   ============================================================ */

'use strict';

// ── Init ─────────────────────────────────────────────────────
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById('statusDot').style.background = '#4ade80';
  } else {
    document.getElementById('statusDot').style.background = '#f87171';
  }
});

// ── Toast ─────────────────────────────────────────────────────
function toast(msg, isError = false) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.style.background = isError ? '#C00000' : '#1F3864';
  el.classList.add('show');
  setTimeout(() => el.classList.remove('show'), 2800);
}

// ── Section toggle ────────────────────────────────────────────
function toggleSection(header) {
  const body = header.nextElementSibling;
  header.classList.toggle('collapsed');
  body.classList.toggle('hidden');
}

// ── Search / Filter ───────────────────────────────────────────
function filterTools(query) {
  const q = query.toLowerCase().trim();
  const sections = document.querySelectorAll('.section[data-section]');
  let anyVisible = false;
  const noRes = document.getElementById('no-results');

  sections.forEach(sec => {
    const sectionName = sec.dataset.section;
    if (sectionName === 'close') { sec.style.display = q ? 'none' : ''; return; }

    const buttons = sec.querySelectorAll('button[data-label]');
    let secHasMatch = false;

    buttons.forEach(btn => {
      const label = btn.dataset.label.toLowerCase();
      const matches = !q || label.includes(q);
      btn.style.display = matches ? '' : 'none';
      if (matches) secHasMatch = true;
    });

    if (!q) {
      sec.style.display = '';
      const header = sec.querySelector('.section-header');
      const body = sec.querySelector('.section-body');
      if (header && body) { header.classList.remove('collapsed'); body.classList.remove('hidden'); }
    } else {
      sec.style.display = secHasMatch ? '' : 'none';
      if (secHasMatch) {
        const header = sec.querySelector('.section-header');
        const body = sec.querySelector('.section-body');
        if (header && body) { header.classList.remove('collapsed'); body.classList.remove('hidden'); }
      }
    }
    if (secHasMatch) anyVisible = true;
  });

  noRes.style.display = q && !anyVisible ? 'block' : 'none';
  document.getElementById('searchTerm').textContent = query;
}

// ── Modal helpers ─────────────────────────────────────────────
let _modalCallback = null;

function openModal(title, fields, onConfirm) {
  document.getElementById('modalTitle').textContent = title;
  const body = document.getElementById('modalBody');
  body.innerHTML = fields.map(f => `
    <label>${f.label}</label>
    ${f.type === 'select'
      ? `<select id="mf_${f.id}">${f.options.map(o=>`<option value="${o.v}">${o.l}</option>`).join('')}</select>`
      : `<input id="mf_${f.id}" type="text" placeholder="${f.placeholder || ''}" value="${f.default || ''}" />`
    }
  `).join('');
  _modalCallback = () => {
    const vals = {};
    fields.forEach(f => { vals[f.id] = document.getElementById('mf_' + f.id).value.trim(); });
    onConfirm(vals);
  };
  document.getElementById('modalConfirm').onclick = () => { closeModal(); _modalCallback(); };
  document.getElementById('modalOverlay').classList.add('show');
  // focus first input
  setTimeout(() => { const first = body.querySelector('input,select'); if (first) first.focus(); }, 100);
}

function closeModal() {
  document.getElementById('modalOverlay').classList.remove('show');
}

// close modal on overlay click
document.getElementById('modalOverlay').addEventListener('click', (e) => {
  if (e.target === document.getElementById('modalOverlay')) closeModal();
});

// ── Close panel ───────────────────────────────────────────────
function closePanel() {
  if (typeof Office !== 'undefined' && Office.context && Office.context.ui) {
    Office.context.ui.closeContainer();
  } else {
    toast('Close the taskpane using the ✕ in the Excel ribbon panel.');
  }
}

// ── Excel helper ──────────────────────────────────────────────
async function run(fn) {
  try {
    await Excel.run(fn);
  } catch (e) {
    toast('Error: ' + (e.message || e), true);
  }
}

// ════════════════════════════════════════════════════════════
// TEXT CASE
// ════════════════════════════════════════════════════════════
async function toUpperCase() {
  await run(async ctx => {
    const range = ctx.workbook.getSelectedRange();
    range.load('values');
    await ctx.sync();
    range.values = range.values.map(r => r.map(c => typeof c === 'string' ? c.toUpperCase() : c));
    await ctx.sync();
    toast('✅ Converted to UPPER CASE');
  });
}

async function toLowerCase() {
  await run(async ctx => {
    const range = ctx.workbook.getSelectedRange();
    range.load('values');
    await ctx.sync();
    range.values = range.values.map(r => r.map(c => typeof c === 'string' ? c.toLowerCase() : c));
    await ctx.sync();
    toast('✅ Converted to lower case');
  });
}

async function toProperCase() {
  await run(async ctx => {
    const range = ctx.workbook.getSelectedRange();
    range.load('values');
    await ctx.sync();
    range.values = range.values.map(r => r.map(c =>
      typeof c === 'string'
        ? c.replace(/\w\S*/g, t => t.charAt(0).toUpperCase() + t.slice(1).toLowerCase())
        : c
    ));
    await ctx.sync();
    toast('✅ Converted to Proper Case');
  });
}

async function toSentenceCase() {
  await run(async ctx => {
    const range = ctx.workbook.getSelectedRange();
    range.load('values');
    await ctx.sync();
    range.values = range.values.map(r => r.map(c =>
      typeof c === 'string' && c.length > 0
        ? c.charAt(0).toUpperCase() + c.slice(1).toLowerCase()
        : c
    ));
    await ctx.sync();
    toast('✅ Converted to Sentence case');
  });
}

async function trimSpaces() {
  await run(async ctx => {
    const range = ctx.workbook.getSelectedRange();
    range.load('values');
    await ctx.sync();
    range.values = range.values.map(r => r.map(c => typeof c === 'string' ? c.trim() : c));
    await ctx.sync();
    toast('✅ Spaces trimmed');
  });
}

async function removeExtraSpaces() {
  await run(async ctx => {
    const range = ctx.workbook.getSelectedRange();
    range.load('values');
    await ctx.sync();
    range.values = range.values.map(r => r.map(c =>
      typeof c === 'string' ? c.trim().replace(/\s+/g, ' ') : c
    ));
    await ctx.sync();
    toast('✅ Extra spaces removed');
  });
}

// ════════════════════════════════════════════════════════════
// SELECTION
// ════════════════════════════════════════════════════════════
async function selectAll() {
  await run(async ctx => {
    const sheet = ctx.workbook.getActiveWorksheet();
    sheet.getRange().select();
    await ctx.sync();
    toast('✅ All cells selected');
  });
}

async function selectUsedRange() {
  await run(async ctx => {
    const sheet = ctx.workbook.getActiveWorksheet();
    sheet.getUsedRange().select();
    await ctx.sync();
    toast('✅ Used range selected');
  });
}

async function selectColumn() {
  await run(async ctx => {
    const range = ctx.workbook.getSelectedRange();
    range.getEntireColumn().select();
    await ctx.sync();
    toast('✅ Entire column(s) selected');
  });
}

async function selectRow() {
  await run(async ctx => {
    const range = ctx.workbook.getSelectedRange();
    range.getEntireRow().select();
    await ctx.sync();
    toast('✅ Entire row(s) selected');
  });
}

function goToCell() {
  openModal('🎯 Go To Cell', [
    { id: 'addr', label: 'Cell Address or Range (e.g. B5 or A1:C10)', placeholder: 'B5', default: 'A1' }
  ], async vals => {
    await run(async ctx => {
      const sheet = ctx.workbook.getActiveWorksheet();
      sheet.getRange(vals.addr).select();
      await ctx.sync();
      toast('✅ Navigated to ' + vals.addr);
    });
  });
}

async function jumpToLastCell() {
  await run(async ctx => {
    const sheet = ctx.workbook.getActiveWorksheet();
    sheet.getUsedRange().getLastCell().select();
    await ctx.sync();
    toast('✅ Jumped to last used cell');
  });
}

// ════════════════════════════════════════════════════════════
// COPY / PASTE
// ════════════════════════════════════════════════════════════
async function copyValues() {
  await run(async ctx => {
    const range = ctx.workbook.getSelectedRange();
    range.load(['values', 'address']);
    await ctx.sync();
    toast('✅ Values copied — use Paste Values to paste without formulas');
  });
}

async function pasteValues() {
  await run(async ctx => {
    const range = ctx.workbook.getSelectedRange();
    range.load('formulas');
    await ctx.sync();
    // Convert formulas to their values
    range.values = range.values; // reads cached; assign to strip formulas
    await ctx.sync();
    toast('✅ Formulas replaced with values');
  });
}

async function fillDown() {
  await run(async ctx => {
    const range = ctx.workbook.getSelectedRange();
    range.load('rowCount');
    await ctx.sync();
    if (range.rowCount < 2) { toast('⚠️ Select 2+ rows to Fill Down', true); return; }
    range.getRow(0).copyFrom(range.getRow(0), Excel.RangeCopyType.all, false, false);
    // Use fill direction
    range.fill.down();
    await ctx.sync();
    toast('✅ Filled down');
  });
}

async function fillRight() {
  await run(async ctx => {
    const range = ctx.workbook.getSelectedRange();
    range.load('columnCount');
    await ctx.sync();
    if (range.columnCount < 2) { toast('⚠️ Select 2+ columns to Fill Right', true); return; }
    range.fill.right();
    await ctx.sync();
    toast('✅ Filled right');
  });
}

async function clearContents() {
  await run(async ctx => {
    ctx.workbook.getSelectedRange().clear(Excel.ClearApplyTo.contents);
    await ctx.sync();
    toast('✅ Contents cleared');
  });
}

async function clearFormats() {
  await run(async ctx => {
    ctx.workbook.getSelectedRange().clear(Excel.ClearApplyTo.formats);
    await ctx.sync();
    toast('✅ Formats cleared');
  });
}

// ════════════════════════════════════════════════════════════
// FORMULA BUILDERS
// ════════════════════════════════════════════════════════════
async function insertFormula(formula) {
  await run(async ctx => {
    const range = ctx.workbook.getSelectedRange();
    range.load('cellCount');
    await ctx.sync();
    range.getCell(0, 0).formulas = [[formula]];
    await ctx.sync();
    toast('✅ Formula inserted: ' + formula.substring(0, 40));
  });
}

function buildVLOOKUP() {
  openModal('🔍 VLOOKUP Builder', [
    { id: 'lv',    label: 'Lookup Value (e.g. A2)',           default: 'A2' },
    { id: 'ta',    label: 'Table Array (e.g. D:F)',           default: 'D:F' },
    { id: 'col',   label: 'Column Index Number',              default: '2' },
    { id: 'exact', label: 'Match Type', type: 'select', options: [
      { v: 'FALSE', l: 'FALSE — Exact Match' },
      { v: 'TRUE',  l: 'TRUE — Approximate Match' }
    ]}
  ], v => insertFormula(`=VLOOKUP(${v.lv},${v.ta},${v.col},${v.exact})`));
}

function buildIF() {
  openModal('❓ IF Builder', [
    { id: 'test',  label: 'Logical Test (e.g. A1>100)',   default: 'A1>0' },
    { id: 'yes',   label: 'Value if TRUE',               default: '"Yes"' },
    { id: 'no',    label: 'Value if FALSE',              default: '"No"' }
  ], v => insertFormula(`=IF(${v.test},${v.yes},${v.no})`));
}

function buildSUMIF() {
  openModal('➕ SUMIF Builder', [
    { id: 'range',    label: 'Criteria Range (e.g. A:A)', default: 'A:A' },
    { id: 'criteria', label: 'Criteria (e.g. "Apple")',  default: '"Apple"' },
    { id: 'sum',      label: 'Sum Range (e.g. B:B)',     default: 'B:B' }
  ], v => insertFormula(`=SUMIF(${v.range},${v.criteria},${v.sum})`));
}

function buildCOUNTIF() {
  openModal('🔢 COUNTIF Builder', [
    { id: 'range',    label: 'Range (e.g. A:A)',         default: 'A:A' },
    { id: 'criteria', label: 'Criteria (e.g. "Apple")',  default: '"Apple"' }
  ], v => insertFormula(`=COUNTIF(${v.range},${v.criteria})`));
}

function buildINDEXMATCH() {
  openModal('📌 INDEX/MATCH Builder', [
    { id: 'lv',   label: 'Lookup Value (e.g. A2)',      default: 'A2' },
    { id: 'lr',   label: 'Lookup Range (e.g. B:B)',     default: 'B:B' },
    { id: 'ret',  label: 'Return Range (e.g. C:C)',     default: 'C:C' }
  ], v => insertFormula(`=INDEX(${v.ret},MATCH(${v.lv},${v.lr},0))`));
}

function buildCONCAT() {
  openModal('🔗 CONCAT Builder', [
    { id: 'a', label: 'First (e.g. A2)',           default: 'A2' },
    { id: 'sep', label: 'Separator (e.g. " ")',    default: '" "' },
    { id: 'b', label: 'Second (e.g. B2)',          default: 'B2' }
  ], v => insertFormula(`=CONCAT(${v.a},${v.sep},${v.b})`));
}

function buildTEXT() {
  openModal('🔤 TEXT Builder', [
    { id: 'val', label: 'Value / Cell (e.g. A2)',            default: 'A2' },
    { id: 'fmt', label: 'Format (e.g. "DD/MM/YYYY")',       default: '"DD/MM/YYYY"' }
  ], v => insertFormula(`=TEXT(${v.val},${v.fmt})`));
}

function buildDATE() {
  openModal('📅 DATE Builder', [
    { id: 'y', label: 'Year (cell or value)',  default: 'A2' },
    { id: 'm', label: 'Month (cell or value)', default: 'B2' },
    { id: 'd', label: 'Day (cell or value)',   default: 'C2' }
  ], v => insertFormula(`=DATE(${v.y},${v.m},${v.d})`));
}

function buildLEFT() {
  openModal('◀ LEFT Builder', [
    { id: 'txt', label: 'Text / Cell', default: 'A2' },
    { id: 'n',   label: 'Num Chars',  default: '5' }
  ], v => insertFormula(`=LEFT(${v.txt},${v.n})`));
}

function buildMID() {
  openModal('↔ MID Builder', [
    { id: 'txt',   label: 'Text / Cell',     default: 'A2' },
    { id: 'start', label: 'Start Position', default: '1' },
    { id: 'n',     label: 'Num Chars',      default: '5' }
  ], v => insertFormula(`=MID(${v.txt},${v.start},${v.n})`));
}

function buildRIGHT() {
  openModal('▶ RIGHT Builder', [
    { id: 'txt', label: 'Text / Cell', default: 'A2' },
    { id: 'n',   label: 'Num Chars',  default: '5' }
  ], v => insertFormula(`=RIGHT(${v.txt},${v.n})`));
}

function buildSUMIFS() {
  openModal('∑ SUMIFS Builder', [
    { id: 'sum',   label: 'Sum Range (e.g. C:C)',      default: 'C:C' },
    { id: 'cr1',   label: 'Criteria Range 1 (e.g. A:A)', default: 'A:A' },
    { id: 'c1',    label: 'Criteria 1',                default: '"Apple"' },
    { id: 'cr2',   label: 'Criteria Range 2 (optional)', default: '' },
    { id: 'c2',    label: 'Criteria 2 (optional)',     default: '' }
  ], v => {
    let f = `=SUMIFS(${v.sum},${v.cr1},${v.c1}`;
    if (v.cr2 && v.c2) f += `,${v.cr2},${v.c2}`;
    f += ')';
    insertFormula(f);
  });
}

// ════════════════════════════════════════════════════════════
// FORMATTING
// ════════════════════════════════════════════════════════════
async function autoFitColumns() {
  await run(async ctx => {
    const range = ctx.workbook.getSelectedRange();
    range.format.autofitColumns();
    await ctx.sync();
    toast('✅ Columns auto-fitted');
  });
}

async function autoFitRows() {
  await run(async ctx => {
    const range = ctx.workbook.getSelectedRange();
    range.format.autofitRows();
    await ctx.sync();
    toast('✅ Rows auto-fitted');
  });
}

async function addBorders() {
  await run(async ctx => {
    const range = ctx.workbook.getSelectedRange();
    const fmt = range.format;
    ['EdgeTop','EdgeBottom','EdgeLeft','EdgeRight','InsideVertical','InsideHorizontal'].forEach(side => {
      const b = fmt.borders.getItem(Excel.BorderIndex[side] !== undefined
        ? Excel.BorderIndex[side] : side);
      b.style = Excel.BorderLineStyle.continuous;
      b.weight = Excel.BorderWeight.thin;
      b.color = '#000000';
    });
    await ctx.sync();
    toast('✅ Borders added');
  });
}

async function removeBorders() {
  await run(async ctx => {
    const range = ctx.workbook.getSelectedRange();
    ['EdgeTop','EdgeBottom','EdgeLeft','EdgeRight','InsideVertical','InsideHorizontal'].forEach(side => {
      range.format.borders.getItem(side).style = Excel.BorderLineStyle.none;
    });
    await ctx.sync();
    toast('✅ Borders removed');
  });
}

async function applyBold() {
  await run(async ctx => {
    const range = ctx.workbook.getSelectedRange();
    range.load('format/font/bold');
    await ctx.sync();
    range.format.font.bold = !range.format.font.bold;
    await ctx.sync();
    toast('✅ Bold toggled');
  });
}

async function applyItalic() {
  await run(async ctx => {
    const range = ctx.workbook.getSelectedRange();
    range.load('format/font/italic');
    await ctx.sync();
    range.format.font.italic = !range.format.font.italic;
    await ctx.sync();
    toast('✅ Italic toggled');
  });
}

async function highlightYellow() {
  await run(async ctx => {
    ctx.workbook.getSelectedRange().format.fill.color = '#FFFF00';
    await ctx.sync();
    toast('✅ Yellow highlight applied');
  });
}

async function clearHighlight() {
  await run(async ctx => {
    ctx.workbook.getSelectedRange().format.fill.clear();
    await ctx.sync();
    toast('✅ Fill color cleared');
  });
}

async function alignLeft() {
  await run(async ctx => {
    ctx.workbook.getSelectedRange().format.horizontalAlignment = Excel.HorizontalAlignment.left;
    await ctx.sync();
    toast('✅ Aligned left');
  });
}

async function alignCenter() {
  await run(async ctx => {
    ctx.workbook.getSelectedRange().format.horizontalAlignment = Excel.HorizontalAlignment.center;
    await ctx.sync();
    toast('✅ Aligned center');
  });
}

async function alignRight() {
  await run(async ctx => {
    ctx.workbook.getSelectedRange().format.horizontalAlignment = Excel.HorizontalAlignment.right;
    await ctx.sync();
    toast('✅ Aligned right');
  });
}

async function wrapText() {
  await run(async ctx => {
    const range = ctx.workbook.getSelectedRange();
    range.load('format/wrapText');
    await ctx.sync();
    range.format.wrapText = !range.format.wrapText;
    await ctx.sync();
    toast('✅ Wrap text toggled');
  });
}

// ════════════════════════════════════════════════════════════
// UTILITIES
// ════════════════════════════════════════════════════════════
async function sortAZ() {
  await run(async ctx => {
    const range = ctx.workbook.getSelectedRange();
    range.load('columnCount');
    await ctx.sync();
    range.sort.apply([{ key: 0, ascending: true }]);
    await ctx.sync();
    toast('✅ Sorted A → Z');
  });
}

async function sortZA() {
  await run(async ctx => {
    const range = ctx.workbook.getSelectedRange();
    range.sort.apply([{ key: 0, ascending: false }]);
    await ctx.sync();
    toast('✅ Sorted Z → A');
  });
}

async function deleteBlankRows() {
  await run(async ctx => {
    const sheet = ctx.workbook.getActiveWorksheet();
    const used = sheet.getUsedRange();
    used.load(['values', 'rowCount', 'columnCount', 'address']);
    await ctx.sync();
    let deleted = 0;
    // Go bottom-up
    for (let i = used.rowCount - 1; i >= 0; i--) {
      const row = used.getRow(i);
      row.load('values');
      await ctx.sync();
      if (row.values[0].every(v => v === null || v === '')) {
        row.getEntireRow().delete(Excel.DeleteShiftDirection.up);
        deleted++;
      }
    }
    await ctx.sync();
    toast(`✅ Deleted ${deleted} blank row(s)`);
  });
}

async function removeDuplicates() {
  await run(async ctx => {
    const range = ctx.workbook.getSelectedRange();
    range.load('columnCount');
    await ctx.sync();
    const cols = Array.from({ length: range.columnCount }, (_, i) => i);
    const result = range.removeDuplicates(cols, true);
    await ctx.sync();
    result.load('removed');
    await ctx.sync();
    toast(`✅ Removed ${result.removed} duplicate row(s)`);
  });
}

async function freezeTopRow() {
  await run(async ctx => {
    const sheet = ctx.workbook.getActiveWorksheet();
    sheet.freezePanes.freezeRows(1);
    await ctx.sync();
    toast('✅ Top row frozen');
  });
}

async function freezeFirstColumn() {
  await run(async ctx => {
    const sheet = ctx.workbook.getActiveWorksheet();
    sheet.freezePanes.freezeColumns(1);
    await ctx.sync();
    toast('✅ First column frozen');
  });
}

async function insertRowAbove() {
  await run(async ctx => {
    const range = ctx.workbook.getSelectedRange();
    range.getEntireRow().insert(Excel.InsertShiftDirection.down);
    await ctx.sync();
    toast('✅ Row inserted above');
  });
}

async function insertColumnLeft() {
  await run(async ctx => {
    const range = ctx.workbook.getSelectedRange();
    range.getEntireColumn().insert(Excel.InsertShiftDirection.right);
    await ctx.sync();
    toast('✅ Column inserted to the left');
  });
}

async function deleteRows() {
  await run(async ctx => {
    ctx.workbook.getSelectedRange().getEntireRow().delete(Excel.DeleteShiftDirection.up);
    await ctx.sync();
    toast('✅ Row(s) deleted');
  });
}

async function deleteColumns() {
  await run(async ctx => {
    ctx.workbook.getSelectedRange().getEntireColumn().delete(Excel.DeleteShiftDirection.left);
    await ctx.sync();
    toast('✅ Column(s) deleted');
  });
}

async function showCellInfo() {
  await run(async ctx => {
    const range = ctx.workbook.getSelectedRange();
    range.load(['address', 'values', 'formulas', 'rowCount', 'columnCount']);
    await ctx.sync();
    const val = range.values[0][0];
    const formula = range.formulas[0][0];
    const info = `📍 ${range.address}\n📦 ${range.rowCount}×${range.columnCount} cells\n💎 Value: ${val}\n𝑓𝑥 Formula: ${formula !== val ? formula : '(none)'}`;
    toast('📍 ' + range.address + ' | ' + range.rowCount + '×' + range.columnCount + ' | ' + val);
  });
}

async function countWords() {
  await run(async ctx => {
    const range = ctx.workbook.getSelectedRange();
    range.load(['values', 'cellCount']);
    await ctx.sync();
    let words = 0, chars = 0, cells = 0;
    range.values.forEach(row => row.forEach(c => {
      if (c !== null && c !== '') {
        const s = String(c).trim();
        chars += s.length;
        words += s ? s.split(/\s+/).length : 0;
        cells++;
      }
    }));
    toast(`${cells} cells | ${words} words | ${chars} chars`);
  });
}

// ════════════════════════════════════════════════════════════
// NUMBER FORMATS
// ════════════════════════════════════════════════════════════
async function applyFormat(fmt, label) {
  await run(async ctx => {
    ctx.workbook.getSelectedRange().numberFormat = [[fmt]];
    await ctx.sync();
    toast('✅ Format applied: ' + label);
  });
}

const formatCurrency = () => applyFormat('$#,##0.00', 'Currency');
const formatPercent  = () => applyFormat('0.00%',     'Percentage');
const formatNumber   = () => applyFormat('#,##0.00',  'Number (2dp)');
const formatDate     = () => applyFormat('DD/MM/YYYY','Date');
const formatGeneral  = () => applyFormat('General',   'General');
const formatText     = () => applyFormat('@',         'Text');
