// app.js

let scheduleData = [];

// File input handler
const fileInput = document.getElementById('fileInput');
fileInput.addEventListener('change', handleFile, false);

function handleFile(e) {
  console.log("handleFile fired, file list:", e.target.files);
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function(evt) {
    const data = new Uint8Array(evt.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    console.log("Workbook sheets:", workbook.SheetNames);
    const sheetName = workbook.SheetNames[0];
    console.log("Using sheet:", sheetName);
    const worksheet = workbook.Sheets[sheetName];

    // Convert sheet to array of arrays (header:1)
    const arr = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    const headers = arr[0];
    const rows = arr.slice(1);

    // Get selected week range
    const start = new Date(document.getElementById('weekStart').value);
    const end   = new Date(document.getElementById('weekEnd').value);

    // Determine which header columns are dates within range
    const dateIndices = headers
      .map((h, i) => {
        const d = new Date(h);
        return (!isNaN(d) && start && end && d >= start && d <= end) ? i : -1;
      })
      .filter(i => i >= 0);

    // Columns D, E, F (0-based indices 3,4,5)
    const baseCols = [3, 4, 5];

    // Filter rows: column D not empty and not "X"
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
  };
  reader.readAsArrayBuffer(file);
}

// Render a preview table from filtered data
def function renderPreview(data) {
  const preview = document.getElementById('preview');
  preview.innerHTML = '';
  if (!data.length) {
    preview.textContent = 'No matching data found.';
    return;
  }
  const table = document.createElement('table');
  table.style.borderCollapse = 'collapse';
  table.style.marginTop = '1em';

  // Header row
  const headerRow = document.createElement('tr');
  Object.keys(data[0]).forEach(key => {
    const th = document.createElement('th');
    th.textContent = key;
    th.style.border = '1px solid #333';
    th.style.padding = '4px 8px';
    headerRow.appendChild(th);
  });
  table.appendChild(headerRow);

  // Data rows
  data.forEach(row => {
    const tr = document.createElement('tr');
    Object.keys(row).forEach(key => {
      const td = document.createElement('td');
      td.textContent = row[key];
      td.style.border = '1px solid #333';
      td.style.padding = '4px 8px';
      tr.appendChild(td);
    });
    table.appendChild(tr);
  });

  preview.appendChild(table);
}

// Stub: Send All button handler
const sendBtn = document.getElementById('sendAll');
sendBtn.addEventListener('click', () => {
  const start = document.getElementById('weekStart').value;
  const end   = document.getElementById('weekEnd').value;
  if (!start || !end) {
    return alert('Please pick both Week Start and Week End.');
  }
  console.log('=== SEND ALL CLICKED ===');
  console.log('Week range:', start, '→', end);
  console.log('Parsed rows:', scheduleData);
  alert(`Would send ${scheduleData.length} emails for ${start} → ${end}`);
});
