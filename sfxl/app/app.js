/* global ej */
'use strict';

// ── Syncfusion license ────────────────────────────────────────────────────────
// For community / trial usage the component works without a key but will show
// a trial banner.  Register your key here to suppress it:
// ej.base.registerLicense('YOUR-SYNCFUSION-LICENSE-KEY');

// ── DOM refs ──────────────────────────────────────────────────────────────────
const fileInput     = document.getElementById('fileInput');
const importBtn     = document.getElementById('importBtn');
const newSheetBtn   = document.getElementById('newSheetBtn');
const exportCsvBtn  = document.getElementById('exportCsvBtn');
const exportXlsxBtn = document.getElementById('exportXlsxBtn');
const resetBtn      = document.getElementById('resetBtn');
const statusEl      = document.getElementById('status');

// ── Helpers ───────────────────────────────────────────────────────────────────
function setStatus(msg) {
  statusEl.textContent = msg;
}

// ── Syncfusion EJ2 Spreadsheet ────────────────────────────────────────────────
const spreadsheet = new ej.spreadsheet.Spreadsheet({
  // Start with one empty sheet
  sheets: [
    {
      name: 'Sheet1',
      rows: [
        // Friendly placeholder header row
        {
          cells: [
            { value: 'Name',       style: { fontWeight: 'bold', background: '#294d7a', color: '#fff' } },
            { value: 'Category',   style: { fontWeight: 'bold', background: '#294d7a', color: '#fff' } },
            { value: 'Amount',     style: { fontWeight: 'bold', background: '#294d7a', color: '#fff' } },
            { value: 'Date',       style: { fontWeight: 'bold', background: '#294d7a', color: '#fff' } },
            { value: 'Notes',      style: { fontWeight: 'bold', background: '#294d7a', color: '#fff' } },
          ]
        }
      ],
      columns: [
        { width: 180 },
        { width: 150 },
        { width: 120 },
        { width: 130 },
        { width: 260 },
      ]
    }
  ],

  // Allow formula bar, show ribbon, allow editing
  showFormulaBar: true,
  showRibbon: true,
  showSheetTabs: true,
  allowEditing: true,
  allowOpen: true,
  allowSave: true,
  allowSorting: true,
  allowFiltering: true,
  allowFindAndReplace: true,
  allowUndoRedo: true,
  allowFormatCells: true,
  allowMerge: true,
  allowInsert: true,
  allowDelete: true,
  allowHyperlink: true,
  allowImage: true,
  allowChart: true,
  allowConditionalFormat: true,
  allowDataValidation: true,
  allowNumberFormatting: true,
  allowWrap: true,
  allowAutoFill: true,
  allowScrolling: true,

  // Freeze the header row
  frozenRowCount: 1,

  // Hooks
  created() {
    setStatus('Ready — open a file or start typing.');
  },

  cellEdit(args) {
    setStatus(`Editing cell ${args.address}`);
  },

  cellSave(args) {
    setStatus(`Saved ${args.address}  ·  value: ${args.value ?? '(empty)'}`);
  },

  beforeSelect(args) {
    setStatus(`Selected: ${args.range}`);
  },
});

spreadsheet.appendTo('#spreadsheet');

// ── Import ────────────────────────────────────────────────────────────────────
importBtn.addEventListener('click', () => fileInput.click());

fileInput.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;

  setStatus(`Loading "${file.name}" …`);

  spreadsheet.open({ file });

  spreadsheet.openComplete = () => {
    setStatus(`Opened: ${file.name}`);
  };

  // Reset input so the same file can be re-opened
  fileInput.value = '';
});

// ── New sheet ─────────────────────────────────────────────────────────────────
newSheetBtn.addEventListener('click', () => {
  const idx   = spreadsheet.sheets.length;
  const name  = `Sheet${idx + 1}`;
  spreadsheet.insertSheet([{ index: idx, sheet: { name } }]);
  spreadsheet.activeSheetIndex = idx;
  setStatus(`New sheet added: ${name}`);
});

// ── Export CSV ────────────────────────────────────────────────────────────────
exportCsvBtn.addEventListener('click', () => {
  const sheetName = spreadsheet.sheets[spreadsheet.activeSheetIndex].name;
  spreadsheet.save({
    saveType: 'Csv',
    fileName: sheetName,
  });
  setStatus(`Exported "${sheetName}" as CSV.`);
});

// ── Export XLSX ───────────────────────────────────────────────────────────────
exportXlsxBtn.addEventListener('click', () => {
  const sheetName = spreadsheet.sheets[spreadsheet.activeSheetIndex].name;
  spreadsheet.save({
    saveType: 'Xlsx',
    fileName: sheetName,
  });
  setStatus(`Exported "${sheetName}" as XLSX.`);
});

// ── Reset (clear current sheet) ───────────────────────────────────────────────
resetBtn.addEventListener('click', () => {
  const idx   = spreadsheet.activeSheetIndex;
  const name  = spreadsheet.sheets[idx].name;

  if (!confirm(`Clear all data in "${name}"? This cannot be undone.`)) return;

  // Select entire used range and delete contents
  spreadsheet.clear({ type: 'Clear All', range: `A1:${getLastCell()}` });
  setStatus(`Sheet "${name}" has been cleared.`);
});

// Returns the last used cell address (e.g. "Z1000") of the active sheet
function getLastCell() {
  const sheet = spreadsheet.sheets[spreadsheet.activeSheetIndex];
  const rows  = sheet.rowCount  || 1000;
  const cols  = sheet.colCount  || 26;
  return `${colIndexToLabel(cols)}${rows}`;
}

function colIndexToLabel(n) {
  let label = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    label     = String.fromCharCode(65 + rem) + label;
    n         = Math.floor((n - 1) / 26);
  }
  return label || 'Z';
}
