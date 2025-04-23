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

/**
 * 1. Load the workbook and dynamically detect header and date rows
 */
function onFileLoad(e) {
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

    // Find the header row index by locating key columns
    const headerIndex = arr.findIndex(
      row => Array.isArray(row) && row.includes('Team') && row.includes('Email') && row.includes('Employee')
    );
    if (headerIndex < 1) {
      alert('Could not find header row with Team, Email, Employee.');
      return;
    }

    // The date row is immediately above header row
    dateRow = arr[headerIndex - 1].map(cell => {
      const d = cell instanceof Date ? cell : new Date(cell);
      return !isNaN(d)
        ? d.toLocaleDateString('en-US', { day: '2-digit', month: 'short', year: '2-digit' }).replace(/,/g, '')
        : cell.toString().trim();
    });

    // Set headerRow and rawRows
    headerRow = arr[headerIndex];
    rawRows = arr.slice(headerIndex + 1);

    // Reset UI
    previewContainer.innerHTML = '<p>File loaded. Select Week Start and click Generate Preview.</p>';
    generateBtn.disabled = false;
    downloadBtn.disabled = true;
    sendBtn.disabled = true;
  };
  reader.readAsArrayBuffer(file);
}

/**
 * 2. Generate a 7-column preview: Email, Employee, and 5-day range
 */
function onGeneratePreview() {
  const startVal = weekStartInput.value;
  if (!startVal) {
    alert('Please select a Week Start date.');
    return;
  }
  // Parse without timezone shift
  const [y, m, d] = startVal.split('-').map(Number);
  const startDate = new Date(y, m - 1, d);

  // Build 5 consecutive date strings matching dateRow
  const dates = Array.from({ length: 5 }, (_, i) => {
    const dt = new Date(startDate);
    dt.setDate(dt.getDate() + i);
    return dt
      .toLocaleDateString('en-US', { day: '2-digit', month: 'short', year: '2-digit' })
      .replace(/,/g, '');
  });

  // Find dynamic indices for Team, Email, Employee
  const teamIdx = headerRow.indexOf('Team');
  const emailIdx = headerRow.indexOf('Email');
  const empIdx = headerRow.indexOf('Employee');
  if (teamIdx < 0 || emailIdx < 0 || empIdx < 0) {
    alert('Missing Team/Email/Employee columns in header.');
    return;
  }

  // Map dates to indices in dateRow
  const dateIndices = dates.map(dt => dateRow.indexOf(dt)).filter(i => i >= 0);

  // Build selected header labels
  selectedHeaders = [headerRow[emailIdx], headerRow[empIdx], ...dates];

  // Filter and map rows
  scheduleData = rawRows
    .filter(r => {
      const t = r[teamIdx];
      return t && t !== 'X';
    })
    .map(r => {
      const obj = {
        [headerRow[emailIdx]]: r[emailIdx],
        [headerRow[empIdx]]: r[empIdx]
      };
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
  const headerTr = document.createElement('tr');
  selectedHeaders.forEach(h => {
    const th = document.createElement('th');
    th.textContent = h;
    headerTr.appendChild(th);
  });
  thead.appendChild(headerTr);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  scheduleData.forEach(row => {
    const tr = document.createElement('tr');
    selectedHeaders.forEach(h => {
      const td = document.createElement('td');
      td.textContent = row[h] || '';
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  previewContainer.appendChild(table);

  // Enable Download Template button
  downloadBtn.disabled = false;
  sendBtn.disabled = true;
}

/**
 * 3. Download the updated Weekly Template sheet on demand
 */
function onDownloadTemplate() {
  if (!workbookGlobal) return;
  const ws = XLSX.utils.json_to_sheet(scheduleData, { header: selectedHeaders });
  workbookGlobal.Sheets['Weekly Template'] = ws;
  if (!workbookGlobal.SheetNames.includes('Weekly Template')) {
    workbookGlobal.SheetNames.push('Weekly Template');
  }
  XLSX.writeFile(workbookGlobal, 'WeeklyTemplate.xlsx');
  sendBtn.disabled = false;
}

/**
 * 4. Stub for sending all emails
 */
function onSendAll() {
  alert(`Would send ${scheduleData.length} emails for the selected week.`);
}
