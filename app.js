// app.js

// Globals for workbook and schedule data
let workbookGlobal = null;
let dateRow = [];
let headerRow = [];
let rawRows = [];
let scheduleData = [];
let selectedHeaders = [];

// DOM elements
const fileInput = document.getElementById('fileInput');
const weekStartInput = document.getElementById('weekStart');
const generateBtn = document.getElementById('generateTemplate');
const downloadBtn = document.getElementById('downloadTemplate');
const sendBtn = document.getElementById('sendAll');
const previewContainer = document.getElementById('preview');

// Event listeners
fileInput.addEventListener('change', onFileLoad);
generateBtn.addEventListener('click', onGeneratePreview);
downloadBtn.addEventListener('click', onDownloadTemplate);
sendBtn.addEventListener('click', onSendAll);

// 1. Load the workbook and extract Schedule tab rows
defunction onFileLoad(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = evt => {
    const wb = XLSX.read(new Uint8Array(evt.target.result), { type: 'array', cellDates: true });
    workbookGlobal = wb;
    const sheetName = 'Schedule';
    if (!wb.SheetNames.includes(sheetName)) {
      alert('Schedule tab not found.');
      return;
    }
    const ws = wb.Sheets[sheetName];
    const arr = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

    // arr[1] = dates row; arr[2] = headers; arr[3...] = data rows
    dateRow = (arr[1] || []).map(cell => {
      const d = (cell instanceof Date) ? cell : new Date(cell);
      return !isNaN(d)
        ? d.toLocaleDateString('en-US', { day: '2-digit', month: 'short', year: '2-digit' }).replace(/,/g, '')
        : cell.toString().trim();
    });
    headerRow = arr[2] || [];
    rawRows = arr.slice(3);

    // Reset UI
    previewContainer.innerHTML = '<p>File loaded. Select Week Start and click Generate Preview.</p>';
    generateBtn.disabled = false;
    downloadBtn.disabled = true;
    sendBtn.disabled = true;
  };
  reader.readAsArrayBuffer(file);
}

// 2. Generate preview only (no download)
defunction onGeneratePreview() {
  const startVal = weekStartInput.value;
  if (!startVal) {
    alert('Please select a Week Start date.');
    return;
  }
  const startDate = new Date(startVal);

  // Build 5 consecutive dates
  const dates = [];
  for (let i = 0; i < 5; i++) {
    const d = new Date(startDate);
    d.setDate(d.getDate() + i);
    dates.push(
      d.toLocaleDateString('en-US', { day: '2-digit', month: 'short', year: '2-digit' }).replace(/,/g, '')
    );
  }

  // Dynamic indices based on headerRow
  const teamIdx = headerRow.findIndex(h => h.toString().trim().toLowerCase() === 'team');
  const emailIdx = headerRow.findIndex(h => h.toString().trim().toLowerCase() === 'email');
  const empIdx = headerRow.findIndex(h => h.toString().trim().toLowerCase() === 'employee');

  // Determine date column indices in dateRow
  const dateIndices = dates
    .map(dt => dateRow.findIndex(cell => cell === dt))
    .filter(idx => idx >= 0);

  // Build selectedHeaders: Email, Employee, then the 5 date strings
  selectedHeaders = [headerRow[emailIdx], headerRow[empIdx], ...dates];

  // Filter rawRows and construct scheduleData
  scheduleData = rawRows
    .filter(r => {
      const teamVal = r[teamIdx];
      return teamVal && teamVal !== 'X';
    })
    .map(r => {
      const obj = {};
      obj[headerRow[emailIdx]] = r[emailIdx];
      obj[headerRow[empIdx]] = r[empIdx];
      dateIndices.forEach((ci, i) => {
        obj[dates[i]] = r[ci] || '';
      });
      return obj;
    });

  // Render preview
  previewContainer.innerHTML = '';
  if (!scheduleData.length) {
    previewContainer.textContent = 'No matching rows for the selected week.';
    return;
  }
  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const tr = document.createElement('tr');
  selectedHeaders.forEach(h => {
    const th = document.createElement('th'); th.textContent = h; tr.appendChild(th);
  });
  thead.appendChild(tr);
  table.appendChild(thead);
  const tbody = document.createElement('tbody');
  scheduleData.forEach(row => {
    const rowEl = document.createElement('tr');
    selectedHeaders.forEach(h => {
      const td = document.createElement('td'); td.textContent = row[h] || ''; rowEl.appendChild(td);
    });
    tbody.appendChild(rowEl);
  });
  table.appendChild(tbody);
  previewContainer.appendChild(table);

  // Enable Download Template button
  downloadBtn.disabled = false;
  sendBtn.disabled = true;
}

// 3. Download the updated Weekly Template sheet
defunction onDownloadTemplate() {
  if (!workbookGlobal) return;
  const ws = XLSX.utils.json_to_sheet(scheduleData, { header: selectedHeaders });
  workbookGlobal.Sheets['Weekly Template'] = ws;
  if (!workbookGlobal.SheetNames.includes('Weekly Template')) {
    workbookGlobal.SheetNames.push('Weekly Template');
  }
  // Trigger download only when clicked
  XLSX.writeFile(workbookGlobal, 'WeeklyTemplate.xlsx');

  // Enable Send All
  sendBtn.disabled = false;
}

// Stub: send all emails
defunction onSendAll() {
  alert(`Would send ${scheduleData.length} emails for the selected week.`);
}
