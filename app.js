// app.js

// — Helpers —
// Format a Date as 'MMM DD yy' (e.g., 'Apr 28 25')
function formatDateShort(d) {
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const day  = String(d.getUTCDate()).padStart(2, '0');
  const mon  = months[d.getUTCMonth()];
  const yy   = String(d.getUTCFullYear()).slice(-2);
  return `${mon} ${day} ${yy}`;
}
// Format a Date as 'MMM DD yyyy' (e.g., 'Apr 28 2025')
function formatDateFull(d) {
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const day  = String(d.getUTCDate()).padStart(2, '0');
  const mon  = months[d.getUTCMonth()];
  const yyyy = d.getUTCFullYear();
  return `${mon} ${day} ${yyyy}`;
}

// — Globals —
let workbookGlobal;
let dateRow      = [];
let headerRow    = [];
let rawRows      = [];
let scheduleData = [];
let selectedHeaders = [];

// — Entry Point — wait until the HTML is fully parsed
document.addEventListener('DOMContentLoaded', () => {
  const fileInput       = document.getElementById('fileInput');
  const weekStartInput  = document.getElementById('weekStart');
  const generateBtn     = document.getElementById('generateTemplate');
  const copyBtn         = document.getElementById('copyAll');
  const previewContainer= document.getElementById('preview');

  // Initial UI state
  generateBtn.disabled = true;
  if (copyBtn) copyBtn.style.display = 'none';

  // Wire events
  fileInput.addEventListener('change', () => {
    onFileLoad(fileInput, generateBtn, copyBtn, previewContainer);
  });
  generateBtn.addEventListener('click', () => {
    onGeneratePreview(weekStartInput, generateBtn, copyBtn, previewContainer);
  });
  if (copyBtn) {
    copyBtn.addEventListener('click', () => {
      onCopyAll(previewContainer);
    });
  }
});

// — 1) Load file, detect header & date rows —
function onFileLoad(fileInput, generateBtn, copyBtn, previewContainer) {
  const file = fileInput.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = evt => {
    const data = new Uint8Array(evt.target.result);
    const wb = XLSX.read(data, { type: 'array', cellDates: true });
    workbookGlobal = wb;

    const ws = wb.Sheets['Schedule'];
    if (!ws) {
      alert('Sheet named "Schedule" not found.');
      return;
    }
    // Convert to array of rows
    const arr = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

    // Find the header row (must contain Team, Email, Employee)
    const hi = arr.findIndex(r => r.includes('Team') && r.includes('Email') && r.includes('Employee'));
    if (hi < 1) {
      alert('Could not detect header row (looking for Team, Email, Employee).');
      return;
    }

    // Build dateRow from the row immediately above the header
    dateRow = (arr[hi - 1] || []).map(cell => {
      const d = new Date(cell);
      return isNaN(d) ? String(cell).trim() : formatDateShort(new Date(d.toUTCString()));
    });

    headerRow = arr[hi] || [];
    rawRows   = arr.slice(hi + 1);

    // Reset UI
    previewContainer.innerHTML = '<p>File loaded. Select Week Start and click "Generate WeeklyTemplate Preview".</p>';
    generateBtn.disabled = false;
    if (copyBtn) copyBtn.style.display = 'none';
  };
  reader.readAsArrayBuffer(file);
}

// — 2) Generate the 7-column preview —
function onGeneratePreview(weekStartInput, generateBtn, copyBtn, previewContainer) {
  const startVal = weekStartInput.value;
  if (!startVal) {
    alert('Please pick a Week Start date.');
    return;
  }
  // Build arrays of 5 days
  const [y, m, d] = startVal.split('-').map(Number);
  const startDate = new Date(Date.UTC(y, m - 1, d));
  const labelsShort = [], labelsFull = [];
  for (let i = 0; i < 5; i++) {
    const dt = new Date(startDate);
    dt.setUTCDate(dt.getUTCDate() + i);
    labelsShort.push(formatDateShort(dt));
    labelsFull.push(formatDateFull(dt));
  }

  // Find starting index in dateRow
  const startIdx = dateRow.indexOf(labelsShort[0]);
  if (startIdx < 0) {
    alert(`Date ${labelsShort[0]} not found in the row above your headers.`);
    return;
  }
  // Map five consecutive columns
  const dateIndices = Array.from({ length: 5 }, (_, i) => startIdx + i)
    .filter(idx => idx >= 0 && idx < dateRow.length);

  // Locate key columns in headerRow
  const teamIdx  = headerRow.indexOf('Team');
  const emailIdx = headerRow.indexOf('Email');
  const empIdx   = headerRow.indexOf('Employee');
  if (teamIdx < 0 || emailIdx < 0 || empIdx < 0) {
    alert('Missing one of Team / Email / Employee columns.');
    return;
  }

  // Build selectedHeaders for table
  selectedHeaders = [
    headerRow[emailIdx],
    headerRow[empIdx],
    ...labelsFull
  ];

  // Filter + map your data rows
  scheduleData = rawRows
    .filter(r => r[teamIdx] && r[teamIdx] !== 'X')
    .map(r => {
      const obj = {};
      obj[ headerRow[emailIdx] ] = r[emailIdx] || '';
      obj[ headerRow[empIdx]   ] = r[empIdx]   || '';
      dateIndices.forEach((ci, j) => {
        obj[ labelsFull[j] ] = r[ci] || '';
      });
      return obj;
    });

  // Render the preview
  previewContainer.innerHTML = '';
  if (!scheduleData.length) {
    previewContainer.textContent = 'No matching rows for that week.';
    return;
  }
  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const trh = document.createElement('tr');
  selectedHeaders.forEach(h => {
    const th = document.createElement('th');
    th.textContent = h;
    trh.appendChild(th);
  });
  thead.appendChild(trh);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  scheduleData.forEach(row => {
    const tr = document.createElement('tr');
    selectedHeaders.forEach(h => {
      const td = document.createElement('td');
      td.textContent = row[h];
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  previewContainer.appendChild(table);

  // Reveal the Copy All button
  if (copyBtn) copyBtn.style.display = 'inline-block';
}

// — 3) Copy the table to clipboard —
function onCopyAll(previewContainer) {
  const tbl = previewContainer.querySelector('table');
  if (!tbl) return;
  const range = document.createRange();
  range.selectNode(tbl);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
  document.execCommand('copy');
  sel.removeAllRanges();
  alert('Preview table copied to clipboard!');
}
