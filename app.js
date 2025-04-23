// app.js

// Helper: format a Date as 'MMM DD YY' (e.g., 'Apr 28 25')
function formatDateShort(d) {
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const day = String(d.getUTCDate()).padStart(2, '0');
  const mon = months[d.getUTCMonth()];
  const yy  = String(d.getUTCFullYear()).slice(-2);
  return `${mon} ${day} ${yy}`;
}

// Helper: format a Date as 'MMM DD YYYY' (e.g., 'Apr 28 2025')
function formatDateFull(d) {
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const day = String(d.getUTCDate()).padStart(2, '0');
  const mon = months[d.getUTCMonth()];
  const yyyy = d.getUTCFullYear();
  return `${mon} ${day} ${yyyy}`;
}

// Globals for workbook and schedule data
let workbookGlobal = null;
let dateRow = [];
let headerRow = [];
let rawRows = [];
let scheduleData = [];
let selectedHeaders = [];

// DOM elements
const fileInput      = document.getElementById('fileInput');
const weekStartInput = document.getElementById('weekStart');
const generateBtn    = document.getElementById('generateTemplate');
const downloadBtn    = document.getElementById('downloadTemplate');
const sendBtn        = document.getElementById('sendAll');
const previewContainer = document.getElementById('preview');

// Event listeners
fileInput.addEventListener('change', onFileLoad);
generateBtn.addEventListener('click', onGeneratePreview);
downloadBtn.addEventListener('click', onDownloadTemplate);
sendBtn.addEventListener('click', onSendAll);

// 1. Load the workbook and detect header/date rows
function onFileLoad(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = evt => {
    const wb = XLSX.read(new Uint8Array(evt.target.result), { type: 'array', cellDates: true });
    workbookGlobal = wb;
    const ws = wb.Sheets['Schedule'];
    if (!ws) {
      alert('Schedule tab not found.');
      return;
    }
    const arr = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const headerIndex = arr.findIndex(r => r.includes('Team') && r.includes('Email') && r.includes('Employee'));
    if (headerIndex < 1) {
      alert('Header row not detected.');
      return;
    }
    // dateRow is row above headerIndex
    dateRow = (arr[headerIndex - 1] || []).map(cell => {
      const d = cell instanceof Date ? cell : new Date(cell);
      return isNaN(d) ? String(cell).trim() : formatDateShort(d);
    });
    headerRow = arr[headerIndex] || [];
    rawRows   = arr.slice(headerIndex + 1);
    previewContainer.innerHTML = '<p>File loaded. Select Week Start and click Generate Preview.</p>';
    generateBtn.disabled = false;
    downloadBtn.disabled = true;
    sendBtn.disabled = true;
  };
  reader.readAsArrayBuffer(file);
}

// 2. Generate preview using correct date indexing
function onGeneratePreview() {
  const startVal = weekStartInput.value;
  if (!startVal) {
    alert('Please select a Week Start date.');
    return;
  }
  const [y, m, d] = startVal.split('-').map(Number);
  const startDate = new Date(Date.UTC(y, m - 1, d));

  // Build 5-day sequences in short format
  const labelsShort = Array.from({ length: 5 }, (_, i) => {
    const dt = new Date(startDate);
    dt.setUTCDate(dt.getUTCDate() + i);
    return formatDateShort(dt);
  });
  const startIdx = dateRow.indexOf(labelsShort[0]);
  if (startIdx < 0) {
    alert(`Date ${labelsShort[0]} not found in schedule date row.`);
    return;
  }
  const dateIndices = [0,1,2,3,4].map(i => startIdx + i).filter(i => i < dateRow.length);

  // Build full-year labels for headers
  const labelsFull = dateIndices.map((_, i) => {
    const dt = new Date(startDate);
    dt.setUTCDate(dt.getUTCDate() + i);
    return formatDateFull(dt);
  });

  // Find Team, Email, Employee columns in headerRow
  const teamIdx  = headerRow.indexOf('Team');
  const emailIdx = headerRow.indexOf('Email');
  const empIdx   = headerRow.indexOf('Employee');
  if (teamIdx < 0 || emailIdx < 0 || empIdx < 0) {
    alert('Missing Team/Email/Employee columns.');
    return;
  }

  // Set up selectedHeaders
  selectedHeaders = [headerRow[emailIdx], headerRow[empIdx], ...labelsFull];

  // Filter and map rawRows to scheduleData
  scheduleData = rawRows.filter(r => r[teamIdx] && r[teamIdx] !== 'X')
    .map(r => {
      const obj = {
        [headerRow[emailIdx]]: r[emailIdx],
        [headerRow[empIdx]]:   r[empIdx]
      };
      dateIndices.forEach((ci, j) => { obj[labelsFull[j]] = r[ci] || ''; });
      return obj;
    });

  // Render preview table
  previewContainer.innerHTML = '';
  if (scheduleData.length === 0) {
    previewContainer.textContent = 'No matching rows for the selected week.';
    return;
  }
  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const thr   = document.createElement('tr');
  selectedHeaders.forEach(h => {
    const th = document.createElement('th');
    th.textContent = h;
    thr.appendChild(th);
  });
  thead.appendChild(thr);
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

// 3. Download the updated Weekly Template sheet when clicked
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

// 4. Stub send all emails
function onSendAll() {
  alert(`Would send ${scheduleData.length} emails for the selected week.`);
}
