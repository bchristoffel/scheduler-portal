// app.js

// Utility to format a Date as 'DD-MMM-YY'
function formatDate(d) {
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const day = String(d.getDate()).padStart(2, '0');
  const mon = months[d.getMonth()];
  const yy = String(d.getFullYear()).slice(-2);
  return `${mon} ${day} ${yy}`;
}

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
const downloadBtn = document.getElementById('downloadTemplate');
const sendBtn = document.getElementById('sendAll');
const previewContainer = document.getElementById('preview');

// Event listeners
fileInput.addEventListener('change', onFileLoad);
generateBtn.addEventListener('click', onGeneratePreview);
downloadBtn.addEventListener('click', onDownloadTemplate);
sendBtn.addEventListener('click', onSendAll);

// 1. Load the workbook and detect header/date rows
function onFileLoad(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = evt => {
    const wb = XLSX.read(new Uint8Array(evt.target.result), { type: 'array', cellDates: true });
    workbookGlobal = wb;
    const ws = wb.Sheets['Schedule'];
    if (!ws) return alert('Schedule tab not found.');

    const arr = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    // find header row (Team, Email, Employee)
    const headerIndex = arr.findIndex(r => r.includes('Team') && r.includes('Email') && r.includes('Employee'));
    if (headerIndex < 1) return alert('Header row not detected.');

    // dateRow is row above header
    dateRow = (arr[headerIndex - 1] || []).map(cell => {
      const d = (cell instanceof Date ? cell : new Date(cell));
      return isNaN(d) ? String(cell).trim() : formatDate(d);
    });
    headerRow = arr[headerIndex] || [];
    rawRows = arr.slice(headerIndex + 1);

    // reset UI
    previewContainer.innerHTML = '<p>File loaded. Select Week Start and click Generate Preview.</p>';
    generateBtn.disabled = false;
    downloadBtn.disabled = true;
    sendBtn.disabled = true;
  };
  reader.readAsArrayBuffer(file);
}

// 2. Generate a 7‑column preview correctly from selected date
function onGeneratePreview() {
  const startVal = weekStartInput.value;
  if (!startVal) return alert('Please select a Week Start date.');
  const [y,m,d] = startVal.split('-').map(Number);
  const startDate = new Date(y, m-1, d);

  // build 5-day labels using formatDate
  const dates = [];
  for (let i=0; i<5; i++) {
    const dt = new Date(startDate);
    dt.setDate(dt.getDate()+i);
    dates.push(formatDate(dt));
  }

  const teamIdx = headerRow.indexOf('Team');
  const emailIdx = headerRow.indexOf('Email');
  const empIdx = headerRow.indexOf('Employee');
  if (teamIdx<0||emailIdx<0||empIdx<0) return alert('Missing Team/Email/Employee columns.');

  const dateIndices = dates.map(dt=>dateRow.indexOf(dt)).filter(i=>i>=0);
  selectedHeaders = [headerRow[emailIdx], headerRow[empIdx], ...dates];

  scheduleData = rawRows
    .filter(r=>{ const t=r[teamIdx]; return t && t!=='X'; })
    .map(r=>{
      const obj = {};
      obj[headerRow[emailIdx]] = r[emailIdx];
      obj[headerRow[empIdx]] = r[empIdx];
      dateIndices.forEach((ci,i)=> obj[dates[i]] = r[ci] || '');
      return obj;
    });

  // render preview
  previewContainer.innerHTML = '';
  if (!scheduleData.length) return previewContainer.textContent = 'No matching rows for the selected week.';
  const tbl=document.createElement('table');
  const thead=document.createElement('thead'), thr=document.createElement('tr');
  selectedHeaders.forEach(h=>{const th=document.createElement('th'); th.textContent=h; thr.appendChild(th);});
  thead.appendChild(thr); tbl.appendChild(thead);
  const tb=document.createElement('tbody');
  scheduleData.forEach(r=>{
    const tr=document.createElement('tr');
    selectedHeaders.forEach(h=>{const td=document.createElement('td'); td.textContent=r[h]||''; tr.appendChild(td);});
    tb.appendChild(tr);
  });
  tbl.appendChild(tb); previewContainer.appendChild(tbl);

  downloadBtn.disabled=false; sendBtn.disabled=true;
}

// 3. Download the template when clicked
function onDownloadTemplate() {
  if (!workbookGlobal) return;
  const ws=XLSX.utils.json_to_sheet(scheduleData,{header:selectedHeaders});
  workbookGlobal.Sheets['Weekly Template']=ws;
  if(!workbookGlobal.SheetNames.includes('Weekly Template')) workbookGlobal.SheetNames.push('Weekly Template');
  XLSX.writeFile(workbookGlobal,'WeeklyTemplate.xlsx');
  sendBtn.disabled=false;
}

// 4. Stub send
function onSendAll(){ alert(`Would send ${scheduleData.length} emails for the selected week.`);}
