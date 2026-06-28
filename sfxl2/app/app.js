/* global ej */
'use strict';

// ─── Syncfusion License ────────────────────────────────────────────────────
// Community / trial: the component works without a key but shows a banner.
// Paste your free community key here to suppress it:
// ej.base.registerLicense('YOUR_KEY_HERE');

// ─── Service URLs (Syncfusion's public demo endpoints) ────────────────────
const OPEN_URL = 'https://document.syncfusion.com/web-services/spreadsheet-editor/api/spreadsheet/open';
const SAVE_URL = 'https://document.syncfusion.com/web-services/spreadsheet-editor/api/spreadsheet/save';

// ─── DOM refs ──────────────────────────────────────────────────────────────
const loaderEl       = document.getElementById('loader');
const appFrameEl     = document.getElementById('appFrame');
const fileInputEl    = document.getElementById('fileInput');
const docNameEl      = document.getElementById('docName');
const tbUnsavedEl    = document.getElementById('tbUnsaved');
const openBtn        = document.getElementById('openBtn');
const newBtn         = document.getElementById('newBtn');
const exportXlsxBtn  = document.getElementById('exportXlsxBtn');
const exportCsvBtn   = document.getElementById('exportCsvBtn');

// ─── Helpers ───────────────────────────────────────────────────────────────
let _unsaved = false;

function markUnsaved() {
  if (_unsaved) return;
  _unsaved = true;
  tbUnsavedEl.hidden = false;
}

function markSaved() {
  _unsaved = false;
  tbUnsavedEl.hidden = true;
}

function hideLoader() {
  loaderEl.classList.add('fade-out');
  // Remove from DOM after transition so it doesn't block clicks
  loaderEl.addEventListener('transitionend', () => loaderEl.remove(), { once: true });
}

function docName() {
  return docNameEl.value.trim() || 'spreadsheet';
}

// ─── Syncfusion EJ2 Spreadsheet ────────────────────────────────────────────
const spreadsheet = new ej.spreadsheet.Spreadsheet({
  // Syncfusion public demo service — handles open/save server-side
  openUrl: OPEN_URL,
  saveUrl: SAVE_URL,

  // Initial data — a clean example table so the grid isn't empty on first load
  sheets: [
    {
      name: 'Sheet1',
      selectedRange: 'A1',
      frozenRows: 1,
      rows: [
        {
          cells: [
            { value: 'Product',    style: { fontWeight: 'bold', backgroundColor: '#0f172a', color: '#f1f5f9', textAlign: 'center' } },
            { value: 'Category',   style: { fontWeight: 'bold', backgroundColor: '#0f172a', color: '#f1f5f9', textAlign: 'center' } },
            { value: 'Unit Price', style: { fontWeight: 'bold', backgroundColor: '#0f172a', color: '#f1f5f9', textAlign: 'center' } },
            { value: 'Qty',        style: { fontWeight: 'bold', backgroundColor: '#0f172a', color: '#f1f5f9', textAlign: 'center' } },
            { value: 'Total',      style: { fontWeight: 'bold', backgroundColor: '#0f172a', color: '#f1f5f9', textAlign: 'center' } },
            { value: 'In Stock',   style: { fontWeight: 'bold', backgroundColor: '#0f172a', color: '#f1f5f9', textAlign: 'center' } },
          ]
        },
        { cells: [{ value: 'Widget Alpha' }, { value: 'Hardware' }, { value: 24.99 }, { value: 120 }, { formula: '=C2*D2' }, { value: true }] },
        { cells: [{ value: 'Gadget Beta'  }, { value: 'Software' }, { value: 59.95 }, { value: 45  }, { formula: '=C3*D3' }, { value: true }] },
        { cells: [{ value: 'Thingamajig'  }, { value: 'Hardware' }, { value: 8.49  }, { value: 300 }, { formula: '=C4*D4' }, { value: false }] },
        { cells: [{ value: 'Doohickey'    }, { value: 'Services' }, { value: 149.00 }, { value: 12 }, { formula: '=C5*D5' }, { value: true }] },
        { cells: [{ value: 'Whatsit'      }, { value: 'Hardware' }, { value: 3.75   }, { value: 980 }, { formula: '=C6*D6' }, { value: true }] },
        { cells: [] },
        {
          cells: [
            {}, {}, {},
            { value: 'Grand Total', style: { fontWeight: 'bold', textAlign: 'right' } },
            { formula: '=SUM(E2:E6)', style: { fontWeight: 'bold' } },
          ]
        },
      ],
      columns: [
        { width: 160 },
        { width: 120 },
        { width: 110 },
        { width: 80  },
        { width: 110 },
        { width: 90  },
      ],
    }
  ],

  // Feature flags — enable the full set Syncfusion supports
  showFormulaBar:          true,
  showRibbon:              true,
  showSheetTabs:           true,
  allowEditing:            true,
  allowOpen:               true,
  allowSave:               true,
  allowSorting:            true,
  allowFiltering:          true,
  allowFindAndReplace:     true,
  allowUndoRedo:           true,
  allowFormatCells:        true,
  allowMerge:              true,
  allowInsert:             true,
  allowDelete:             true,
  allowHyperlink:          true,
  allowImage:              true,
  allowChart:              true,
  allowConditionalFormat:  true,
  allowDataValidation:     true,
  allowNumberFormatting:   true,
  allowWrap:               true,
  allowAutoFill:           true,
  allowScrolling:          true,
  allowResizing:           true,

  // ── Lifecycle ────────────────────────────────────────────────────────────
  created() {
    // Format price and total columns as currency after creation
    spreadsheet.numberFormat('$#,##0.00', 'C2:C6');
    spreadsheet.numberFormat('$#,##0.00', 'E2:E8');
    hideLoader();
  },

  // ── Cell edit / save ─────────────────────────────────────────────────────
  cellEdit() {
    markUnsaved();
  },

  cellSave() {
    markUnsaved();
  },

  // ── File open complete ────────────────────────────────────────────────────
  openComplete() {
    markSaved();
    // Set document title to the first sheet name if available
    const firstName = spreadsheet.sheets?.[0]?.name;
    if (firstName) docNameEl.value = firstName;
  },
});

spreadsheet.appendTo('#spreadsheet');

// ─── New ───────────────────────────────────────────────────────────────────
newBtn.addEventListener('click', () => {
  if (_unsaved && !confirm('Discard unsaved changes and start a new spreadsheet?')) return;

  // Destroy and recreate — cleanest way to reset to blank
  spreadsheet.destroy();
  const blank = new ej.spreadsheet.Spreadsheet({
    openUrl: OPEN_URL,
    saveUrl: SAVE_URL,
    showFormulaBar:      true,
    showRibbon:          true,
    showSheetTabs:       true,
    allowEditing:        true,
    allowOpen:           true,
    allowSave:           true,
    allowSorting:        true,
    allowFiltering:      true,
    allowFindAndReplace: true,
    allowUndoRedo:       true,
    allowFormatCells:    true,
    allowMerge:          true,
    allowInsert:         true,
    allowDelete:         true,
    allowHyperlink:      true,
    allowImage:          true,
    allowChart:          true,
    allowConditionalFormat: true,
    allowDataValidation: true,
    allowNumberFormatting: true,
    allowWrap:           true,
    allowAutoFill:       true,
    allowScrolling:      true,
    allowResizing:       true,
  });
  blank.appendTo('#spreadsheet');
  docNameEl.value = 'Untitled spreadsheet';
  markSaved();
});

// ─── Open ──────────────────────────────────────────────────────────────────
openBtn.addEventListener('click', () => fileInputEl.click());

fileInputEl.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;

  // Update document name from the file
  const baseName = file.name.replace(/\.[^.]+$/, '');
  docNameEl.value = baseName;
  markSaved();

  // Pass the File to Syncfusion — it POSTs to openUrl and renders the result
  spreadsheet.open({ file });

  // Reset so same file can be reopened
  fileInputEl.value = '';
});

// ─── Export XLSX ────────────────────────────────────────────────────────────
exportXlsxBtn.addEventListener('click', () => {
  spreadsheet.save({
    url:      SAVE_URL,
    fileName: docName(),
    saveType: 'Xlsx',
  });
  markSaved();
});

// ─── Export CSV ─────────────────────────────────────────────────────────────
exportCsvBtn.addEventListener('click', () => {
  spreadsheet.save({
    url:      SAVE_URL,
    fileName: docName(),
    saveType: 'Csv',
  });
  markSaved();
});

// ─── Document rename via title bar input ────────────────────────────────────
docNameEl.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') docNameEl.blur();
  if (e.key === 'Escape') {
    docNameEl.value = docNameEl.defaultValue;
    docNameEl.blur();
  }
});

docNameEl.addEventListener('blur', () => {
  // Keep default if cleared
  if (!docNameEl.value.trim()) docNameEl.value = 'Untitled spreadsheet';
  docNameEl.defaultValue = docNameEl.value;
});

// ─── Keyboard shortcuts ──────────────────────────────────────────────────────
document.addEventListener('keydown', (e) => {
  if ((e.ctrlKey || e.metaKey) && !e.shiftKey) {
    if (e.key === 'o' || e.key === 'O') { e.preventDefault(); openBtn.click(); }
    if (e.key === 'n' || e.key === 'N') { e.preventDefault(); newBtn.click(); }
    if (e.key === 's' || e.key === 'S') { e.preventDefault(); exportXlsxBtn.click(); }
  }
});
