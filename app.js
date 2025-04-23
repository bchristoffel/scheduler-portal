// app.js

let workbookGlobal = null;
let scheduleData = [];
let selectedHeaders = [];

// File input and date selectors
const fileInput = document.getElementById('fileInput');
const weekStartInput = document.getElementById('weekStart');
const weekEndInput   = document.getElementById('weekEnd');
fileInput.addEventListener('change', handleFile, false);

document.getElementById('sendAll').addEventListener('click', () => {
  const start = weekStartInput.value;
  const end   = weekEndInput.value;
  if (!start || !end) return alert('Please pick both Week Start and Week End.');
  console.log('=== SEND ALL CLICKED ===', { weekStart: start, weekEnd: end, rows: scheduleData });
  alert(`Would send ${scheduleData.length} emails for ${start} → ${end}`);
});

function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function(evt) {
    const data = new Uint8Array(evt.target.result);
    const wb = XLSX.read(data, { type: 'array' });
    workbookGlobal = wb;

    // Use the "Schedule" sheet explicitly
    const sheetName = 'Schedule';
    if (!wb.SheetNames.includes(sheetName)) {
      return alert(`Sheet named '${sheetName}' not found.`);
    }
    const worksheet = wb.Sheets[sheetName];

    // Read as array of arrays for header and rows
    const arr = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    const headers = arr[0];
    const rows = arr.slice(1);

    // Parse selected date range
    const start = new Date(weekStartInput.value);
    const end   = new Date(weekEndInput.value);

    // Determine date columns within range
    const dateIndices = headers
      .map((h, i) => {
        const d = new Date(h);
        return (!isNaN(d) && d >= start && d <= end) ? i : -1;
      })
      .filter(i => i >= 0);

    // Always include columns E (index 4) and F (index 5)
    const baseCols = [4, 5];

    // Build selectedHeaders array for template sheet
    selectedHeaders = baseCols.concat(dateIndices).map(i => headers[i]);

    // Filter rows: column D (index 3) not empty and not "X"
    const filtered = rows
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

    scheduleData = filtered;
    renderPreview(filtered);
    document.getElementById('sendAll').disabled = false;

    // Also update the "Weekly Template" tab and prompt download
    generateWeeklyTemplate(wb, filtered, selectedHeaders);
  };
  reader.readAsArrayBuffer(file);
}

// Render a preview table
function renderPreview(data) {
  const preview = document.getElementById('preview');
  preview.innerHTML = '';
  if (!data.length) {
    preview.textContent = 'No matching data found.';
    return;
  }
  const table = document.createElement('table');
  table.style.borderCollapse = 'collapse';
  table.style.marginTop = '1em';

  // Header
  const headerRow = document.createElement('tr');
  selectedHeaders.forEach(key => {
    const th = document.createElement('th');
    th.textContent = key;
    th.style.border = '1px solid #333';
    th.style.padding = '4px 8px';
    headerRow.appendChild(th);
  });
  table.appendChild(headerRow);

  // Rows
  data.forEach(row => {
    const tr = document.createElement('tr');
    selectedHeaders.forEach(key => {
      const td = document.createElement('td');
      td.textContent = row[key] || '';
      td.style.border = '1px solid #333';
      td.style.padding = '4px 8px';
      tr.appendChild(td);
    });
    table.appendChild(tr);
  });
  document.getElementById('preview').appendChild(table);
}

// Generate or update the "Weekly Template" sheet and trigger download
function generateWeeklyTemplate(wb, data, headers) {
  const sheetName = 'Weekly Template';
  // Create worksheet from JSON
  const ws = XLSX.utils.json_to_sheet(data, { header: headers });
  // Assign or replace
  wb.Sheets[sheetName] = ws;
  if (!wb.SheetNames.includes(sheetName)) {
    wb.SheetNames.push(sheetName);
  }
  // Prompt user to save updated workbook
  XLSX.writeFile(wb, 'Updated_Schedule.xlsx');
}
