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
const sendBtn = document.getElementById('sendAll');
const previewContainer = document.getElementById('preview');

// Event listeners
fileInput.addEventListener('change', onFileLoad);
generateBtn.addEventListener('click', onGenerateTemplate);
sendBtn.addEventListener('click', onSendAll);

// 1. Load the workbook and extract Schedule tab rows
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

    // Row indices based on your file layout:
    // arr[1] = second row with dates (e.g., '28-Apr-25', '29-Apr-25', ...)
    // arr[2] = third row with column headers: 'Schedule Order', ... 'Team', 'Email', 'Employee'
    // arr[3...] = data rows
    dateRow = arr[1] || [];
    headerRow = arr[2] || [];
    rawRows = arr.slice(3);

    // Enable generate button
    generateBtn.disabled = false;
    sendBtn.disabled = true;
    previewContainer.innerHTML = '<p>File loaded. Select Week Start and click Generate Template.</p>';
  };
  reader.readAsArrayBuffer(file);
}

// 2. Generate preview and Weekly Template
function onGenerateTemplate() {
  const startVal = weekStartInput.value;
  if (!startVal) {
    alert('Please select a Week Start date.');
    return;
  }
  const startDate = new Date(startVal);

  // Build five consecutive dates matching dateRow format
  const dates = [];
  for (let i = 0; i < 5; i++) {
    const d = new Date(startDate);
    d.setDate(d.getDate() + i);
    const str = d.toLocaleDateString('en-US', { day: '2-digit', month: 'short', year: '2-digit' }).replace(/,/g, '');
    dates.push(str);
  }

  // Fixed indices: Team(D)=3, Email(E)=4, Employee(F)=5
  const teamIdx = 3;
  const emailIdx = 4;
  const empIdx = 5;

  // Map dates to column indices in dateRow
  const dateIndices = dates.map(dt => dateRow.indexOf(dt)).filter(i => i >= 0);

  // Build selectedHeaders: Email, Employee + the five date strings
  selectedHeaders = [headerRow[emailIdx], headerRow[empIdx], ...dates];

  // Filter rawRows by Team and map to scheduleData
  scheduleData = rawRows
    .filter(r => {
      const teamVal = r[teamIdx];
      return teamVal && teamVal !== 'X';
    })
    .map(r => {
      const obj = {};
      obj[headerRow[emailIdx]] = r[emailIdx];
      obj[headerRow[empIdx]] = r[empIdx];
      dateIndices.forEach((colIdx, i) => {
        obj[dates[i]] = r[colIdx] || '';
      });
      return obj;
    });

  // Render preview
  renderPreview();

  // Create/Update Weekly Template sheet and prompt download
  updateTemplateSheet();

  // Enable send
  sendBtn.disabled = false;
}

// Render preview table
function renderPreview() {
  previewContainer.innerHTML = '';
  if (!scheduleData.length) {
    previewContainer.textContent = 'No matching rows for the selected week.';
    return;
  }
  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const headerTr = document.createElement('tr');
  selectedHeaders.forEach(h => {
    const th = document.createElement('th'); th.textContent = h; headerTr.appendChild(th);
  });
  thead.appendChild(headerTr);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  scheduleData.forEach(row => {
    const tr = document.createElement('tr');
    selectedHeaders.forEach(h => {
      const td = document.createElement('td'); td.textContent = row[h] || ''; tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  previewContainer.appendChild(table);
}

// Update Weekly Template sheet and download workbook
function updateTemplateSheet() {
  const sheetName = 'Weekly Template';
  const ws = XLSX.utils.json_to_sheet(scheduleData, { header: selectedHeaders });
  workbookGlobal.Sheets[sheetName] = ws;
  if (!workbookGlobal.SheetNames.includes(sheetName)) {
    workbookGlobal.SheetNames.push(sheetName);
  }
  XLSX.writeFile(workbookGlobal, 'WeeklyTemplate.xlsx');
}

// Stub send all
function onSendAll() {
  alert(`Would send ${scheduleData.length} emails for the selected week.`);
}
