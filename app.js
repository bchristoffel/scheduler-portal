// app.js

// Globals for parsed workbook data
let workbookGlobal = null;
let headerRow = [];
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
      alert('Schedule tab not found.');
      return;
    }
    const ws = wb.Sheets[sheetName];
    const arr = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    headerRow = arr[1];   // row 2: headers including Team, Email, Employee, dates
    rawRows = arr.slice(2); // data starts on row 3

    // Enable the Generate button
    generateBtn.disabled = false;
    sendBtn.disabled = true;
    previewContainer.innerHTML = '<p>File loaded. Choose Week Start and click Generate Template.</p>';
  };
  reader.readAsArrayBuffer(file);
}
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

  // Build the list of 5 date strings matching headerRow format
  const dates = [];
  for (let i = 0; i < 5; i++) {
    const d = new Date(startDate);
    d.setDate(d.getDate() + i);
    const str = d.toLocaleDateString('en-US', { day: '2-digit', month: 'short', year: '2-digit' }).replace(/,/g, '');
    dates.push(str);
  }

  // Fixed column indices relative to headerRow
  const teamIdx = 3;   // column D
  const emailIdx = 4;  // column E
  const empIdx = 5;    // column F

  // Identify date column indices in headerRow
  const dateIndices = dates
    .map(dt => headerRow.findIndex(h => h === dt))
    .filter(idx => idx >= 0);

  // Build selectedHeaders: Email, Employee, then the 5 dates
  selectedHeaders = [headerRow[emailIdx], headerRow[empIdx], ...dates];

  // Filter and map rows
  scheduleData = rawRows
    .filter(r => {
      const team = r[teamIdx];
      return team && team !== 'X';
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

  // Render preview and download template sheet
  renderPreview();
  updateTemplateSheet();
  sendBtn.disabled = false;
} and trigger file download
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
