// app.js

// — Helpers —
// Format date helpers (as defined earlier)
function formatDateShort(d) {
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const day  = String(d.getUTCDate()).padStart(2,'0');
  const mon  = months[d.getUTCMonth()];
  const yy   = String(d.getUTCFullYear()).slice(-2);
  return `${mon} ${day} ${yy}`;
}
function formatDateFull(d) {
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const day  = String(d.getUTCDate()).padStart(2,'0');
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

// Pagination globals
let emailPage = 1;
const emailsPerPage = 10;

// Entry point
document.addEventListener('DOMContentLoaded', () => {
  const fileInput       = document.getElementById('fileInput');
  const weekStartInput  = document.getElementById('weekStart');
  const generateBtn     = document.getElementById('generateTemplate');
  const copyBtn         = document.getElementById('copyAll');
  const generateEmails  = document.getElementById('generateEmails');
  const sendAllBtn      = document.getElementById('sendAll');
  const previewContainer= document.getElementById('preview');
  const emailPreview    = document.getElementById('emailPreview');

  // Initial UI
  generateBtn.disabled    = true;
  if (copyBtn) copyBtn.style.display = 'none';
  generateEmails.disabled = true;
  sendAllBtn.disabled     = true;

  // Wire events
  fileInput.addEventListener('change', () => onFileLoad(fileInput, generateBtn, copyBtn, previewContainer));
  generateBtn.addEventListener('click', () => {
    onGeneratePreview(weekStartInput, generateBtn, copyBtn, previewContainer);
    generateEmails.disabled = scheduleData.length === 0;
  });
  if (copyBtn) copyBtn.addEventListener('click', () => onCopyAll(previewContainer));
  generateEmails.addEventListener('click', () => {
    emailPage = 1;
    renderEmailPage(emailPreview, sendAllBtn);
    // switch tab
    document.querySelector('.tablinks[data-tab="emails"]').click();
  });
  sendAllBtn.addEventListener('click', onSendAll);
});

// 1) Load file
function onFileLoad(fileInput, generateBtn, copyBtn, previewContainer) {
  const file = fileInput.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = evt => {
    const wb = XLSX.read(new Uint8Array(evt.target.result), { type:'array', cellDates:true });
    workbookGlobal = wb;
    const ws = wb.Sheets['Schedule'];
    if (!ws) { alert('Schedule sheet not found'); return; }
    const arr = XLSX.utils.sheet_to_json(ws, { header:1, defval:'' });
    const hi = arr.findIndex(r => r.includes('Team') && r.includes('Email') && r.includes('Employee'));
    if (hi < 1) { alert('Header row not detected'); return; }
    dateRow = (arr[hi-1]||[]).map(c=>{
      const d=new Date(c); return isNaN(d)?String(c).trim():formatDateShort(new Date(d.toUTCString()));
    });
    headerRow = arr[hi]||[];
    rawRows   = arr.slice(hi+1);
    previewContainer.innerHTML = '<p>Select Week Start and click Generate Weekly Preview.</p>';
    generateBtn.disabled = false;
    if (copyBtn) copyBtn.style.display='none';
  };
  reader.readAsArrayBuffer(file);
}

// 2) Generate weekly preview
function onGeneratePreview(weekStartInput, generateBtn, copyBtn, previewContainer) {
  const val = weekStartInput.value;
  if (!val) { alert('Pick a Week Start'); return; }
  const [y,m,d] = val.split('-').map(Number);
  const startDate = new Date(Date.UTC(y,m-1,d));
  const labelsShort=[], labelsFull=[];
  for(let i=0;i<5;i++){
    const dt=new Date(startDate); dt.setUTCDate(dt.getUTCDate()+i);
    labelsShort.push(formatDateShort(dt));
    labelsFull.push(formatDateFull(dt));
  }
  const startIdx = dateRow.indexOf(labelsShort[0]);
  if (startIdx<0){ alert(`Date ${labelsShort[0]} not found`); return; }
  const dateIndices = Array.from({length:5},(_,i)=>startIdx+i).filter(i=>i>=0&&i<dateRow.length);
  const teamIdx  = headerRow.indexOf('Team'),
        emailIdx = headerRow.indexOf('Email'),
        empIdx   = headerRow.indexOf('Employee');
  if (teamIdx<0||emailIdx<0||empIdx<0){ alert('Missing Team/Email/Employee'); return; }
  selectedHeaders = [ headerRow[emailIdx], headerRow[empIdx], ...labelsFull ];
  scheduleData = rawRows.filter(r=>r[teamIdx]&&r[teamIdx]!=='X')
    .map(r=>{
      const o={};
      o[ headerRow[emailIdx] ] = r[emailIdx]||'';
      o[ headerRow[empIdx]   ] = r[empIdx]  ||'';
      dateIndices.forEach((ci,j)=> o[ labelsFull[j] ] = r[ci]||'' );
      return o;
    });
  // render preview
  previewContainer.innerHTML = '';
  if (!scheduleData.length) {
    previewContainer.textContent = 'No matching rows.';
    return;
  }
  const tbl=document.createElement('table'),
        thead=document.createElement('thead'),
        thr=document.createElement('tr');
  selectedHeaders.forEach(h=>{
    const th=document.createElement('th'); th.textContent=h; thr.appendChild(th);
  });
  thead.appendChild(thr); tbl.appendChild(thead);
  const tbody=document.createElement('tbody');
  scheduleData.forEach(r=>{
    const tr=document.createElement('tr');
    selectedHeaders.forEach(h=>{
      const td=document.createElement('td'); td.textContent=r[h]||''; tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  tbl.appendChild(tbody);
  previewContainer.appendChild(tbl);
  if (copyBtn=document.getElementById('copyAll')) copyBtn.style.display='inline-block';
}

// 3) Copy preview
function onCopyAll(previewContainer) {
  const tbl = previewContainer.querySelector('table');
  if (!tbl) return;
  const range = document.createRange();
  range.selectNode(tbl);
  const sel = window.getSelection();
  sel.removeAllRanges(); sel.addRange(range);
  document.execCommand('copy'); sel.removeAllRanges();
  alert('Table copied!');
}

// 4) Render one page of email drafts
function renderEmailPage(emailPreview, sendAllBtn) {
  emailPreview.innerHTML = '';
  const total = scheduleData.length,
        pages = Math.ceil(total/emailsPerPage),
        start = (emailPage-1)*emailsPerPage,
        pageData = scheduleData.slice(start, start+emailsPerPage);
  if (!pageData.length) {
    emailPreview.textContent = 'No drafts on this page.';
  } else {
    pageData.forEach(row=>{
      const email = row[selectedHeaders[0]],
            name  = row[selectedHeaders[1]],
            subject = 'Schedule';
      // build table
      let tbl = '<table style="border-collapse:collapse;width:100%;margin:1em 0;">'
              + '<thead><tr><th></th>';
      selectedHeaders.slice(2).forEach(h=>{ tbl+=`<th style="border:1px solid #ddd;padding:6px;">${h}</th>`; });
      tbl+= '</tr><tr><th></th>';
      selectedHeaders.slice(2).forEach(h=>{
        const wd=new Date(h).toLocaleDateString('en-US',{weekday:'short'});
        tbl+=`<th style="border:1px solid #ddd;padding:6px;">${wd}</th>`;
      });
      tbl+= '</tr></thead><tbody><tr>'
          + `<td style="border:1px solid #ddd;padding:6px;font-weight:600;">${name}</td>`;
      selectedHeaders.slice(2).forEach(h=>{
        const v=row[h];
        if(v) tbl+=`<td style="border:1px solid #ddd;padding:6px;">${v}</td>`;
        else  tbl+=`<td style="border:1px solid #ddd;padding:6px;text-align:center;">
                     <img src="AW_DIMENSIONAL_BLACK_HOR_2024.png" alt="Logo"
                          style="max-height:24px;opacity:0.3;" />
                   </td>`;
      });
      tbl+= '</tr></tbody></table>';
      // body wrapper
      const body = `<div style="font-family:Segoe UI,Arial,sans-serif;color:#333;">
        <p>Hi Team &ndash;</p>
        <p>Please see your schedule for next week below. If you have any questions, let us know.</p>
        ${tbl}
        <p>Thank you!</p>
      </div>`;
      const card=document.createElement('div');
      card.className='email-card';
      card.innerHTML = `<h3>To: ${name}; ${email}</h3>
                        <p><strong>Subject:</strong> ${subject}</p>${body}`;
      emailPreview.appendChild(card);
    });
  }
  renderPaginationControls(pages);
  sendAllBtn.disabled = false;
}

// 5) Pagination controls
function renderPaginationControls(totalPages) {
  const emailPreview = document.getElementById('emailPreview');
  let pg = document.getElementById('emailPagination');
  if (pg) pg.remove();
  pg=document.createElement('div');
  pg.id='emailPagination';
  pg.style.textAlign='center';
  const prev=document.createElement('button');
  prev.className='button'; prev.textContent='← Prev';
  prev.disabled = emailPage===1;
  prev.onclick = ()=>{ emailPage--; renderEmailPage(emailPreview, document.getElementById('sendAll')); };
  const info=document.createElement('span');
  info.textContent=` Page ${emailPage} of ${totalPages} `;
  info.style.margin='0 1em';
  const next=document.createElement('button');
  next.className='button'; next.textContent='Next →';
  next.disabled = emailPage===totalPages;
  next.onclick = ()=>{ emailPage++; renderEmailPage(emailPreview, document.getElementById('sendAll')); };
  pg.append(prev, info, next);
  emailPreview.parentNode.insertBefore(pg, emailPreview);
}

// 6) Stub send all
function onSendAll() {
  if (!confirm(`Send all ${scheduleData.length} emails now?`)) return;
  alert(`(Stub) Would send ${scheduleData.length} emails.`);
}
