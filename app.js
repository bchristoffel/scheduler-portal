// app.js

// Global variables for workbook and filtered schedule data
let workbookGlobal = null;
let scheduleData = [];
let selectedHeaders = [];

// DOM elements
const fileInput = document.getElementById('fileInput');
const weekStartInput = document.getElementById('weekStart');
const weekEndInput   = document.getElementById('weekEnd');
const sendBtn        = document.getElementById('sendAll');

// Event listeners
fileInput.addEventListener('change', handleFile, false);
sendBtn.addEventListener('click', () => {
  const start = weekStartInput.value;
  const end   = weekEndInput.value;
  if (!start || !end) {
    alert('Please pick both Week Start and Week End.');
    return;
  }
  console.log('=== SEND ALL CLICKED ===', { weekStart: start, weekEnd: end, rows: scheduleData });
  alert(`Would send ${scheduleData.length} emails for ${start} → ${end}`);
});

// Handle file upload and parse the Schedule sheet
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
    const worksheet = wb.Sheets[sheetName];

    // Convert sheet to header+rows array
    const arr     = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    const headers = arr[0];
    const rows    = arr.slice(1);

    const startDate = new Date(weekStartInput.value);
    const endDate   = new Date(weekEndInput.value);

    // Identify date columns within range
    const dateIndices = headers
      .map((h, i) => {
        const d = new Date(h);
        return (!isNaN(d) && d >= startDate && d <= endDate) ? i : -1;
      })
      .filter(i => i >= 0);

    // Always include columns E (index 4) and F (index 5)
    const baseCols = [4, 5];

    // Save selected headers
    selectedHeaders = baseCols.concat(dateIndices).map(i => headers[i]);

    // Filter rows: Column D (index 3) not empty or 'X'
    scheduleData = rows
      .filter(r => {
        const dVal = r[3];
        return dVal !== '' && dVal !== 'X';
      })
      .map(r => {
        const obj = {};
        baseCols.concat(dateIndices).forEach(i => {
          obj[headers[i]] = r[i];
        });
        return obj;
      });

    // Render preview and enable send button
    renderPreview(scheduleData);
    sendBtn.disabled = false;

    // Update Weekly Template and prompt download
    generateWeeklyTemplate(wb, scheduleData, selectedHeaders);
  };
  reader.readAsArrayBuffer(file);
}

// Render the filtered schedule preview table
function renderPreview(data) {
  const preview = document.getElementById('preview');
  preview.innerHTML = '';
  if (data.length === 0) {
    preview.textContent = 'No matching data found.';
    return;
  }

  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');

  selectedHeaders.forEach(key => {
    const th = document.createElement('th');
    th.textContent = key;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  data.forEach(row => {
    const tr = document.createElement('tr');
    selectedHeaders.forEach(key => {
      const td = document.createElement('td');
      td.textContent = row[key] || '';
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  preview.appendChild(table);
}

// Generate/update Weekly Template sheet and trigger download of updated workbook
function generateWeeklyTemplate(wb, data, headers) {
  const sheetName = 'Weekly Template';
  const ws = XLSX.utils.json_to_sheet(data, { header: headers });
  wb.Sheets[sheetName] = ws;
  if (!wb.SheetNames.includes(sheetName)) {
    wb.SheetNames.push(sheetName);
  }
  XLSX.writeFile(wb, 'Updated_Schedule.xlsx');
}
