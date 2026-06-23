const spreadsheetEl = document.getElementById('spreadsheet');
const statusEl = document.getElementById('status');
const fileInput = document.getElementById('fileInput');

let sheetData = createSheet(10, 8);
let fileName = 'spreadsheet';

function createSheet(rows, cols) {
  return Array.from({ length: rows }, () => Array.from({ length: cols }, () => ''));
}

function normalizeSheet(data) {
  if (!Array.isArray(data) || !data.length) {
    return [['']];
  }

  const rows = data.map((row) => {
    if (!Array.isArray(row)) {
      return [row == null ? '' : String(row)];
    }

    return row.map((value) => (value == null ? '' : String(value)));
  });

  const maxCols = Math.max(...rows.map((row) => row.length), 1);
  return rows.map((row) => {
    const normalized = [...row];
    while (normalized.length < maxCols) {
      normalized.push('');
    }
    return normalized.slice(0, maxCols);
  });
}

function renderSheet() {
  spreadsheetEl.innerHTML = '';

  const table = document.createElement('table');
  table.className = 'sheet-table';

  const headerRow = document.createElement('tr');
  const corner = document.createElement('th');
  corner.className = 'corner-cell';
  headerRow.appendChild(corner);

  for (let col = 0; col < sheetData[0].length; col += 1) {
    const cell = document.createElement('th');
    cell.textContent = String.fromCharCode(65 + col);
    headerRow.appendChild(cell);
  }

  const thead = document.createElement('thead');
  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');

  sheetData.forEach((row, rowIndex) => {
    const tr = document.createElement('tr');
    const rowHeader = document.createElement('th');
    rowHeader.className = 'row-header';
    rowHeader.textContent = String(rowIndex + 1);
    tr.appendChild(rowHeader);

    row.forEach((cellValue, colIndex) => {
      const td = document.createElement('td');
      const input = document.createElement('input');
      input.type = 'text';
      input.value = cellValue;
      input.dataset.row = String(rowIndex);
      input.dataset.col = String(colIndex);
      input.addEventListener('input', (event) => {
        const currentRow = Number(event.target.dataset.row);
        const currentCol = Number(event.target.dataset.col);
        sheetData[currentRow][currentCol] = event.target.value;
      });
      td.appendChild(input);
      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });

  table.appendChild(tbody);
  spreadsheetEl.appendChild(table);
}

function updateStatus(message) {
  statusEl.textContent = message;
}

function triggerDownload(content, filename, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function loadWorkbookFromFile(file) {
  const reader = new FileReader();
  reader.onload = (event) => {
    try {
      const workbook = XLSX.read(event.target.result, { type: 'binary' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      sheetData = normalizeSheet(rows);
      fileName = file.name.replace(/\.[^.]+$/, '');
      updateStatus(`${file.name} loaded successfully.`);
      renderSheet();
    } catch (error) {
      console.error(error);
      updateStatus('Unable to load the selected file.');
    }
  };
  reader.readAsBinaryString(file);
}

function exportCsv() {
  const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
  const csvContent = XLSX.utils.sheet_to_csv(worksheet);
  triggerDownload(csvContent, `${fileName}.csv`, 'text/csv;charset=utf-8');
  updateStatus(`Exported ${fileName}.csv.`);
}

function exportXlsx() {
  const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
  XLSX.writeFile(workbook, `${fileName}.xlsx`);
  updateStatus(`Exported ${fileName}.xlsx.`);
}

document.getElementById('addRowBtn').addEventListener('click', () => {
  const newRow = Array.from({ length: sheetData[0].length }, () => '');
  sheetData = [...sheetData, newRow];
  renderSheet();
  updateStatus('Added a new row.');
});

document.getElementById('addColBtn').addEventListener('click', () => {
  sheetData = sheetData.map((row) => [...row, '']);
  renderSheet();
  updateStatus('Added a new column.');
});

document.getElementById('resetBtn').addEventListener('click', () => {
  sheetData = createSheet(10, 8);
  fileName = 'spreadsheet';
  renderSheet();
  updateStatus('Reset the sheet.');
});

document.getElementById('exportCsvBtn').addEventListener('click', exportCsv);
document.getElementById('exportXlsxBtn').addEventListener('click', exportXlsx);

fileInput.addEventListener('change', (event) => {
  const [file] = event.target.files;
  if (!file) {
    return;
  }
  loadWorkbookFromFile(file);
});

renderSheet();
