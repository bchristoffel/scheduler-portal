// app.js

// Globals for workbook and schedule data\let workbookGlobal = null;
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

// 1. Load the workbook and process Schedule tab
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
    // Read all rows as array-of-arrays
    const arr = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

    // Interpret rows: rowIndex 1 = dates, rowIndex 2 = headers, data from rowIndex 3 onward
    const rawDateRow = arr[1] || [];
    headerRow = arr[2] || [];
    rawRows = arr.slice(3);

    // Normalize dates to strings matching input format
    dateRow = rawDateRow.map(cell => {
      const d = (cell instanceof Date) ? cell : new Date(cell);
      if (!isNaN(d)) {
        return d.toLocaleDateString('en-US', { day: '2-digit', month: 'short', year: '2-digit' }).replace(/,/g, '');
      }
      return cell.toString().trim();
    });

    // Enable Preview button
    generateBtn.disabled = false;
    downloadBtn.disabled = true;
    sendBtn.disabled = true;
    previewContainer.innerHTML = '<p>File loaded. Select Week Start and click Generate Preview.</p>';
  };
  reader.readAsArrayBuffer(file);
}

// 2. Generate preview only (no download)
function onGeneratePreview() {
  const startVal = weekStartInput.value;
  if (!startVal) {
    alert('Please select a Week Start date.');
    return;
  }
  const startDate = new Date(startVal);

  // Build five consecutive date strings
  const dates = [];
  for (let i = 0; i < 5; i++) {
    const d = new Date(startDate);
    d.setDate(d.getDate() + i);
    const str = d.toLocaleDateString('en-US', { day: '2-digit', month: 'short', year: '2-digit' }).replace(/,/g, '');
    dates.push(str);
  }

  // Fixed column positions in headerRow: D=Team(3), E=Email(4), F=Employee(5)
  const teamIdx = 3;
  const emailIdx = 4;
  const empIdx = 5;

  // Determine which indices in dateRow match our 5 dates
  const dateIndices = dates.map(dt => dateRow.indexOf(dt)).filter(i => i >= 0);

  // Prepare selectedHeaders for display and template
  selectedHeaders = [headerRow[emailIdx], headerRow[empIdx], ...dates];

  // Filter and map rawRows for scheduleData
  scheduleData = rawRows
    .filter(r => {
      const teamVal = r[teamIdx];
      return teamVal && teamVal !== 'X';
    })
    .map(r => {
      const obj = {};
      obj[headerRow[emailIdx]] = r[emailIdx];
      obj[headerRow[empIdx]] = r[empIdx];
      dateIndices.forEach((ci, i) => obj[dates[i]] = r[ci] || '');
      return obj;
    });

  // Render the preview table
  renderPreview();

  // Enable Download button
  downloadBtn.disabled = false;
  sendBtn.disabled = true;
}

// Render the preview table
function renderPreview() {
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
  scheduleData.forEach(r => {
    const row = document.createElement('tr');
    selectedHeaders.forEach(h => {
      const td = document.createElement('td'); td.textContent = r[h] || ''; row.appendChild(td);
    });
    tbody.appendChild(row);
  });
  table.appendChild(tbody);
  previewContainer.appendChild(table);
}

// 3. Download the updated Weekly Template sheet when prompted
function onDownloadTemplate() {
  if (!workbookGlobal) return;
  const sheetName = 'Weekly Template';
  const ws = XLSX.utils.json_to_sheet(scheduleData, { header: selectedHeaders });
  workbookGlobal.Sheets[sheetName] = ws;
  if (!workbookGlobal.SheetNames.includes(sheetName)) workbookGlobal.SheetNames.push(sheetName);
  // Trigger download
  XLSX.writeFile(workbookGlobal, 'WeeklyTemplate.xlsx');

  // Enable the Send All button
  sendBtn.disabled = false;
}

// Stub: send all emails
function onSendAll() {
  alert(`Would send ${scheduleData.length} emails for the selected week.`);
}
