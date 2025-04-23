// app.js

// Hold the parsed schedule rows
let scheduleData = [];

// When a file is selected…
document.getElementById('fileInput').addEventListener('change', handleFile, false);

function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(evt) {
    const data = new Uint8Array(evt.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    // Use the first sheet (you can change this if needed)
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    // Convert sheet to JSON, with headers from row 1
    const json = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
    scheduleData = json;
    renderPreview(json);
    document.getElementById('sendAll').disabled = false;
  };
  reader.readAsArrayBuffer(file);
}

// Render a simple HTML table of the schedule
function renderPreview(data) {
  const preview = document.getElementById('preview');
  preview.innerHTML = '';
  if (!data.length) {
    preview.textContent = 'No data found in the sheet.';
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
// When Send All is clicked, gather the data and just log it for now
document.getElementById('sendAll').addEventListener('click', () => {
  const start = document.getElementById('weekStart').value;
  const end   = document.getElementById('weekEnd'  ).value;

  if (!start || !end) {
    return alert('Please pick both Week Start and Week End.');
  }

  console.log('=== SEND ALL CLICKED ===');
  console.log('Week range:', start, '→', end);
  console.log('Parsed rows:', scheduleData);

  alert(`Would send ${scheduleData.length} emails for ${start} → ${end}`);
});

// (We’ll wire up sendAll next)
