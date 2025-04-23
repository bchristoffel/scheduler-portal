// app.js

// Globals for parsed workbook data
let workbookGlobal = null;
let headerRow = [];
let dateRow = [];
let rawRows = [];
let scheduleData = [];
let selectedHeaders = [];

// DOM elements
const fileInput = document.getElementById('fileInput');
const weekStartInput = document.getElementById('weekStart');
const generateBtn = document.getElementById('generateTemplate');
const sendBtn = document.getElementById('sendAll');
const previewContainer = document.getElementById('preview');

// Wire up events
fileInput.addEventListener('change', onFileLoad);
generateBtn.addEventListener('click', onGenerateTemplate);
sendBtn.addEventListener('click', onSendAll);

// 1. Load the workbook and extract Schedule tab data
function onFileLoad(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = evt => {
    const wb = XLSX.read(new Uint8Array(evt.target.result), { type: 'array' });
    workbookGlobal = wb;
    const sheetName = 'Schedule';
    if (!wb.SheetNames.includes(sheetName)) {
      alert("Schedule tab not found.");
      return;
    }
    const ws = wb.Sheets[sheetName];
    const arr = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    headerRow = arr[0];   // row 1: Team, Email, Employee labels
    dateRow = arr[1];     // row 2: actual dates (e.g. 28-Apr-25)
    rawRows = arr.slice(2);

    // Enable the Generate button
    generateBtn.disabled = false;
    sendBtn.disabled = true;
    previewContainer.innerHTML = '<p>File loaded. Choose Week Start and click Generate Template.</p>';
  };
  reader.readAsArrayBuffer(file);
}

// 2. Build 5-day slice, filter rows, preview, and build template
function onGenerateTemplate() {
  const startVal = weekStartInput.value;
  if (!startVal) {
    alert('Please select a Week Start date.');
    return;
  }
  const startDate = new Date(startVal);

  // Build the list of 5 date strings matching row2 format
  const dates = [];
  for (let i = 0; i < 5; i++) {
    const d = new Date(startDate);
    d.setDate(d.getDate() + i);
    const str = d.toLocaleDateString('en-US', { day: '2-digit', month: 'short', year: '2-digit' }).replace(/,/g, '');
    dates.push(str);
  }

  // Find column indices
  const teamIdx = headerRow.findIndex(h => h === 'Team');
  const emailIdx = headerRow.findIndex(h => h === 'Email');
  const empIdx = headerRow.findIndex(h => h === 'Employee');
  const dateIndices = dates.map(dt => dateRow.findIndex(h => h === dt)).filter(idx => idx >= 0);

  // Selected headers for preview & sheet
  selectedHeaders = ['Email', 'Employee', ...dates];

  // Filter rows where Team (col D) is not empty and not 'X'
  scheduleData = rawRows
    .filter(r => {
      const team = r[teamIdx];
      return team && team !== 'X';
    })
    .map(r => {
      const obj = {};
      obj['Email'] = r[emailIdx];
      obj['Employee'] = r[empIdx];
      dateIndices.forEach((ci, i) => {
        obj[dates[i]] = r[ci] || '';
      });
      return obj;
    });

  // Preview the 7 columns: Email, Employee, and 5 dates
  renderPreview();

  // Update the Weekly Template tab in workbook and prompt download
  updateTemplateSheet();

  sendBtn.disabled = false;
}

// Render the preview table
function renderPreview() {
  previewContainer.innerHTML = '';
  if (scheduleData.length === 0) {
    previewContainer.textContent = 'No matching rows for the selected week.';
    return;
  }
  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const hr = document.createElement('tr');
  selectedHeaders.forEach(h => {
    const th = document.createElement('th');
    th.textContent = h;
    hr.appendChild(th);
  });
  thead.appendChild(hr);
  table.appendChild(thead);
  const tb = document.createElement('tbody');
  scheduleData.forEach(row => {
    const tr = document.createElement('tr');
    selectedHeaders.forEach(h => {
      const td = document.createElement('td');
      td.textContent = row[h] || '';
      tr.appendChild(td);
    });
    tb.appendChild(tr);
  });
  table.appendChild(tb);
  previewContainer.appendChild(table);
}

// Create or update 'Weekly Template' sheet and trigger file download
function updateTemplateSheet() {
  const sheetName = 'Weekly Template';
  const ws = XLSX.utils.json_to_sheet(scheduleData, { header: selectedHeaders });
  workbookGlobal.Sheets[sheetName] = ws;
  if (!workbookGlobal.SheetNames.includes(sheetName)) {
    workbookGlobal.SheetNames.push(sheetName);
  }
  XLSX.writeFile(workbookGlobal, 'WeeklyTemplate.xlsx');
}

// Stub: send all emails (to be replaced with real email logic)
function onSendAll() {
  alert(`Would send ${scheduleData.length} emails for the selected week.`);
}
