// app.js

// Globals
let workbookGlobal = null;
let rawHeaders = [];
let rawRows = [];
let scheduleData = [];
let selectedHeaders = [];

// DOM elements
const fileInput         = document.getElementById('fileInput');
const weekStartInput    = document.getElementById('weekStart');
const weekEndInput      = document.getElementById('weekEnd');
const generateBtn       = document.getElementById('generateTemplate');
const sendBtn           = document.getElementById('sendAll');
const previewContainer  = document.getElementById('preview');

// Event listeners
fileInput.addEventListener('change', handleFile, false);
generateBtn.addEventListener('click', generateTemplatePreview, false);
sendBtn.addEventListener('click', sendAllEmails, false);

// Step 1: Load workbook and raw data
function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function(evt) {
    const data = new Uint8Array(evt.target.result);
    const wb   = XLSX.read(data, { type: 'array' });
    workbookGlobal = wb;

    const sheetName = 'Schedule';
    if (!wb.SheetNames.includes(sheetName)) {
      alert(`Sheet named '${sheetName}' not found.`);
      return;
    }
    const ws     = wb.Sheets[sheetName];
    const arr    = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    rawHeaders  = arr[0];
    rawRows     = arr.slice(1);

    // Enable generate after file is loaded
    generateBtn.disabled = false;
    sendBtn.disabled = true;
    previewContainer.innerHTML = '<p>File loaded. Pick week range and click Generate Template.</p>';
  };
  reader.readAsArrayBuffer(file);
}

// Step 2: Generate preview and update Weekly Template sheet
function generateTemplatePreview() {
  const startVal = weekStartInput.value;
  const endVal   = weekEndInput.value;
  if (!startVal || !endVal) {
    alert('Please pick both Week Start and Week End.');
    return;
  }
  const startDate = new Date(startVal);
  const endDate   = new Date(endVal);

  // Identify date columns within range
  const dateIndices = rawHeaders
    .map((h, i) => {
      const d = new Date(h);
      return (!isNaN(d) && d >= startDate && d <= endDate) ? i : -1;
    })
    .filter(i => i >= 0);

  // Always include E (4) & F (5)
  const baseCols = [4,5];
  selectedHeaders = baseCols.concat(dateIndices).map(i => rawHeaders[i]);

  // Filter rows: D (3) not empty or 'X'
  scheduleData = rawRows
    .filter(r => {
      const v = r[3];
      return v !== '' && v !== 'X';
    })
    .map(r => {
      const obj = {};
      baseCols.concat(dateIndices).forEach(i => obj[rawHeaders[i]] = r[i]);
      return obj;
    });

  // Render preview
  renderPreview(scheduleData);

  // Update Weekly Template sheet in workbook and prompt download
  updateWeeklyTemplateSheet();

  // Enable send button
  sendBtn.disabled = false;
}

// Render preview function
function renderPreview(data) {
  previewContainer.innerHTML = '';
  if (data.length === 0) {
    previewContainer.textContent = 'No matching data for the selected range.';
    return;
  }
  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');
  selectedHeaders.forEach(h => {
    const th = document.createElement('th'); th.textContent = h; headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  data.forEach(row => {
    const tr = document.createElement('tr');
    selectedHeaders.forEach(h => {
      const td = document.createElement('td'); td.textContent = row[h] || ''; tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  previewContainer.appendChild(table);
}

// Update 'Weekly Template' sheet and download workbook
function updateWeeklyTemplateSheet() {
  const sheetName = 'Weekly Template';
  const ws = XLSX.utils.json_to_sheet(scheduleData, { header: selectedHeaders });
  workbookGlobal.Sheets[sheetName] = ws;
  if (!workbookGlobal.SheetNames.includes(sheetName)) workbookGlobal.SheetNames.push(sheetName);
  XLSX.writeFile(workbookGlobal, 'Updated_Schedule_with_Template.xlsx');
}

// Stub for sending emails
defunction sendAllEmails() {
  const start = weekStartInput.value;
  const end   = weekEndInput.value;
  console.log('Sending emails for', {start, end, rows: scheduleData.length});
  alert(`Would send ${scheduleData.length} emails for ${start} → ${end}`);
}
