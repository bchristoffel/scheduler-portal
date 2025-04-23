// app.js

// Helper: format a Date as 'MMM DD yy' (e.g., 'Apr 28 25')
function formatDateShort(d) {
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const day = String(d.getDate()).padStart(2, '0');
  const mon = months[d.getMonth()];
  const yy  = String(d.getFullYear()).slice(-2);
  return `${mon} ${day} ${yy}`;
}

// Helper: format a Date as 'MMM DD yyyy' (e.g., 'Apr 28 2025')
function formatDateFull(d) {
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const day = String(d.getDate()).padStart(2, '0');
  const mon = months[d.getMonth()];
  const yyyy = d.getFullYear();
  return `${mon} ${day} ${yyyy}`;
}

// Globals
let workbookGlobal = null;
let dateRow = [];
let headerRow = [];
let rawRows = [];
let scheduleData = [];
let selectedHeaders = [];

// DOM references
const fileInput = document.getElementById('fileInput');
const weekStartInput = document.getElementById('weekStart');
const generateBtn = document.getElementById('generateTemplate');
const copyBtn = document.getElementById('copyAll');
const previewContainer = document.getElementById('preview');

// DOM references (initialized after DOM is ready)
document.addEventListener('DOMContentLoaded', () => {
  const fileInput = document.getElementById('fileInput');
  const weekStartInput = document.getElementById('weekStart');
  const generateBtn = document.getElementById('generateTemplate');
  const copyBtn = document.getElementById('copyAll');
  const previewContainer = document.getElementById('preview');

  // Hide copy button initially
  if (copyBtn) copyBtn.style.display = 'none';

  // Wire event listeners
  if (fileInput) fileInput.addEventListener('change', onFileLoad);
  if (generateBtn) generateBtn.addEventListener('click', onGeneratePreview);
  if (copyBtn) copyBtn.addEventListener('click', onCopyAll);
});('DOMContentLoaded', () => {
  fileInput.addEventListener('change', onFileLoad);
  generateBtn.addEventListener('click', onGeneratePreview);
  copyBtn.addEventListener('click', onCopyAll);
});

// 1. Load workbook and extract data
function onFileLoad(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(e) {
    const data = new Uint8Array(e.target.result);
    const wb = XLSX.read(data, { type: 'array', cellDates: true });
    workbookGlobal = wb;
    const ws = wb.Sheets['Schedule'];
    if (!ws) {
      alert('Schedule tab not found.');
      return;
    }
    const arr = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const hi = arr.findIndex(row => row.includes('Team') && row.includes('Email') && row.includes('Employee'));
    if (hi < 1) {
      alert('Header row not detected.');
      return;
    }
    dateRow = (arr[hi - 1] || []).map(cell => {
      const d = new Date(cell);
      return isNaN(d) ? String(cell).trim() : formatDateShort(d);
    });
    headerRow = arr[hi];
    rawRows = arr.slice(hi + 1);
    previewContainer.innerHTML = '<p>File loaded. Select Week Start and click Generate WeeklyTemplate Preview.</p>';
    generateBtn.disabled = false;
    copyBtn.style.display = 'none';
  };
  reader.readAsArrayBuffer(file);
}

// 2. Generate WeeklyTemplate preview
function onGeneratePreview() {
  const startVal = weekStartInput.value;
  if (!startVal) {
    alert('Please select a Week Start date.');
    return;
  }
  const [y, m, d] = startVal.split('-').map(Number);
  const startDate = new Date(y, m - 1, d);
  const labelsShort = [];
  const labelsFull = [];
  for (let i = 0; i < 5; i++) {
    const dt = new Date(startDate);
    dt.setDate(dt.getDate() + i);
    labelsShort.push(formatDateShort(dt));
    labelsFull.push(formatDateFull(dt));
  }
  const startIdx = dateRow.indexOf(labelsShort[0]);
  if (startIdx < 0) {
    alert(`Date ${labelsShort[0]} not found in dates row.`);
    return;
  }
  const dateIndices = labelsShort.map((_, i) => startIdx + i).filter(idx => idx >= 0 && idx < dateRow.length);
  const teamIdx = headerRow.indexOf('Team');
  const emailIdx = headerRow.indexOf('Email');
  const empIdx = headerRow.indexOf('Employee');
  if (teamIdx < 0 || emailIdx < 0 || empIdx < 0) {
    alert('Missing Team/Email/Employee columns.');
    return;
  }
  selectedHeaders = [headerRow[emailIdx], headerRow[empIdx], ...labelsFull];
  scheduleData = rawRows
    .filter(r => r[teamIdx] && r[teamIdx] !== 'X')
    .map(r => {
      const obj = {
        [headerRow[emailIdx]]: r[emailIdx],
        [headerRow[empIdx]]: r[empIdx]
      };
      dateIndices.forEach((ci, j) => obj[labelsFull[j]] = r[ci] || '');
      return obj;
    });
  previewContainer.innerHTML = '';
  if (!scheduleData.length) {
    previewContainer.textContent = 'No matching rows for the selected week.';
    return;
  }
  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const thr = document.createElement('tr');
  selectedHeaders.forEach(h => {
    const th = document.createElement('th'); th.textContent = h; thr.appendChild(th);
  });
  thead.appendChild(thr);
  table.appendChild(thead);
  const tbody = document.createElement('tbody');
  scheduleData.forEach(r => {
    const tr = document.createElement('tr');
    selectedHeaders.forEach(h => {
      const td = document.createElement('td'); td.textContent = r[h] || ''; tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  previewContainer.appendChild(table);
  copyBtn.style.display = 'inline-block';
}

// 3. Copy table to clipboard
function onCopyAll() {
  const tbl = previewContainer.querySelector('table');
  if (!tbl) return;
  const range = document.createRange();
  range.selectNode(tbl);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
  document.execCommand('copy');
  sel.removeAllRanges();
  alert('Preview table copied!');
}
