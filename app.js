// app.js

// — Helpers —
function formatDateShort(d) {
  const m = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return `${m[d.getUTCMonth()]} ${String(d.getUTCDate()).padStart(2,'0')} ${String(d.getUTCFullYear()).slice(-2)}`;
}
function formatDateFull(d) {
  const m = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return `${m[d.getUTCMonth()]} ${String(d.getUTCDate()).padStart(2,'0')} ${d.getUTCFullYear()}`;
}

// — Globals —
let workbookGlobal, dateRow = [], headerRow = [], rawRows = [];
let scheduleData = [], selectedHeaders = [];

// Pagination
let emailPage = 1, emailsPerPage = 10;

document.addEventListener('DOMContentLoaded', () => {
  const fileInput       = document.getElementById('fileInput');
  const weekStartInput  = document.getElementById('weekStart');
  const generateBtn     = document.getElementById('generateTemplate');
  const copyBtn         = document.getElementById('copyAll');
  const generateEmails  = document.getElementById('generateEmails');
  const sendAllBtn      = document.getElementById('sendAll');
  const previewContainer= document.getElementById('preview');

  // Initial UI
  generateBtn.disabled    = true;
  if (copyBtn) copyBtn.style.display = 'none';
  generateEmails.disabled = true;
  sendAllBtn.disabled     = true;

  // File load
  fileInput.addEventListener('change', () =>
    onFileLoad(fileInput, generateBtn, copyBtn, previewContainer)
  );

  // Weekly preview
  generateBtn.addEventListener('click', () => {
    onGeneratePreview(weekStartInput, generateBtn, copyBtn, previewContainer);
    generateEmails.disabled = scheduleData.length === 0;
  });

  // Copy preview
  if (copyBtn) copyBtn.addEventListener('click', () => onCopyAll(previewContainer));

  // Generate emails
  generateEmails.addEventListener('click', () => {
    emailPage = 1;
    renderEmailPage();
    document.querySelector('.tablinks[data-tab="emails"]').click();
  });

  // Send all stub
  sendAllBtn.addEventListener('click', onSendAll);
});

// 1) Load workbook & detect headers
function onFileLoad(fileInput, generateBtn, copyBtn, previewContainer) {
  const file = fileInput.files[0]; if (!file) return;
  const reader = new FileReader();
  reader.onload = evt => {
    const wb = XLSX.read(new Uint8Array(evt.target.result), { type:'array', cellDates:true });
    workbookGlobal = wb;
    const ws = wb.Sheets['Schedule'];
    if (!ws) return alert('Schedule tab missing');
    const arr = XLSX.utils.sheet_to_json(ws, { header:1, defval:'' });
    const hi = arr.findIndex(r => r.includes('Team') && r.includes('Email') && r.includes('Employee'));
    if (hi < 1) return alert('Could not find header row');
    dateRow = (arr[hi-1]||[]).map(c=>{
      const d=new Date(c);
      return isNaN(d)?String(c).trim():formatDateShort(new Date(d.toUTCString()));
    });
    headerRow = arr[hi]||[];
    rawRows   = arr.slice(hi+1);
    previewContainer.innerHTML = '<p>Select Week Start & click Generate Weekly Preview.</p>';
    generateBtn.disabled = false;
    if (copyBtn) copyBtn.style.display = 'none';
  };
  reader.readAsArrayBuffer(file);
}

// 2) Generate Weekly Preview
function onGeneratePreview(weekStartInput, generateBtn, copyBtn, previewContainer) {
  const val = weekStartInput.value || ''; if (!val) return alert('Pick a Week Start date');
  const [y,m,d] = val.split('-').map(Number);
  const start = new Date(Date.UTC(y,m-1,d));
  const labelsShort = [], labelsFull = [];
  for (let i=0;i<5;i++){
    const dt=new Date(start); dt.setUTCDate(dt.getUTCDate()+i);
    labelsShort.push(formatDateShort(dt));
    labelsFull.push(formatDateFull(dt));
  }
  const startIdx = dateRow.indexOf(labelsShort[0]);
  if (startIdx<0) return alert(`Date ${labelsShort[0]} not found`);
  const dateIndices = Array.from({length:5},(_,i)=>startIdx+i).filter(i=>i>=0&&i<dateRow.length);

  const teamIdx  = headerRow.indexOf('Team'),
        emailIdx = headerRow.indexOf('Email'),
        empIdx   = headerRow.indexOf('Employee');
  if (teamIdx<0||emailIdx<0||empIdx<0) return alert('Missing Team/Email/Employee');

  selectedHeaders = [headerRow[emailIdx], headerRow[empIdx], ...labelsFull];
  scheduleData = rawRows.filter(r=>r[teamIdx]&&r[teamIdx]!=='X')
    .map(r=>{
      const o = {
        [headerRow[emailIdx]]: r[emailIdx]||'',
        [headerRow[empIdx]]:   r[empIdx]  ||''
      };
      dateIndices.forEach((ci,j)=> o[labelsFull[j]] = r[ci]||'');
      return o;
    });

  // Render table
  previewContainer.innerHTML = '';
  if (!scheduleData.length) {
    previewContainer.textContent = 'No matching rows.';
    return;
  }
  const tbl=document.createElement('table');
  const thead=document.createElement('thead');
  const thr=document.createElement('tr');
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

// 3) Copy
function onCopyAll(previewContainer) {
  const tbl = previewContainer.querySelector('table');
  if (!tbl) return;
  const range=document.createRange(); range.selectNode(tbl);
  const sel=window.getSelection(); sel.removeAllRanges(); sel.addRange(range);
  document.execCommand('copy'); sel.removeAllRanges();
  alert('Preview copied!');
}

// 4) Render 10 email drafts/page
function renderEmailPage() {
  const emailPreview = document.getElementById('emailPreview');
  const sendAllBtn   = document.getElementById('sendAll');
  emailPreview.innerHTML='';

  const total= scheduleData.length;
  const pages= Math.ceil(total/emailsPerPage);
  const startIdx= (emailPage-1)*emailsPerPage;
  const pageData= scheduleData.slice(startIdx, startIdx+emailsPerPage);

  pageData.forEach(r=>{
    const email = r[selectedHeaders[0]];
    const subject = 'Schedule';

    // Build table header
    let tbl = '<table style="border-collapse:collapse;width:100%;margin:1em 0;">'
            + '<thead><tr><th></th>';
    selectedHeaders.slice(2).forEach(h=>{
      tbl+= `<th style="border:1px solid #ddd;padding:6px;">${h}</th>`;
    });
    tbl+= '</tr><tr><th></th>';
    selectedHeaders.slice(2).forEach(h=>{
      const wd=new Date(h).toLocaleDateString('en-US',{weekday:'short'});
      tbl+= `<th style="border:1px solid #ddd;padding:6px;">${wd}</th>`;
    });
    tbl+= '</tr></thead><tbody><tr>'
        + `<td style="border:1px solid #ddd;padding:6px;font-weight:600;"></td>`;
    // Data row (no logo now)
    selectedHeaders.slice(2).forEach(h=>{
      const v=r[h];
      tbl+= `<td style="border:1px solid #ddd;padding:6px;">${v||''}</td>`;
    });
    tbl+= '</tr></tbody></table>';

    const bodyHtml = `<div style="font-family:Segoe UI,Arial,sans-serif;color:#333;">
      <p>Hi Team &ndash;</p>
      <p>Please see your schedule for next week below. If you have any questions, let us know.</p>
      ${tbl}
      <p>Thank you!</p>
    </div>`;

    const card=document.createElement('div');
    card.className='email-card';
    card.innerHTML=`<h3>To: ${email}</h3>
      <p><strong>Subject:</strong> ${subject}</p>
      ${bodyHtml}`;
    emailPreview.appendChild(card);
  });

  renderPaginationControls(pages);
  sendAllBtn.disabled = false;
}

// 5) Pagination
function renderPaginationControls(totalPages) {
  const emailPreview=document.getElementById('emailPreview');
  let pg=document.getElementById('emailPagination');
  if(pg) pg.remove();
  pg=document.createElement('div');
  pg.id='emailPagination';
  pg.style.textAlign='center';

  const prev=document.createElement('button');
  prev.className='button'; prev.textContent='← Prev';
  prev.disabled = emailPage===1;
  prev.onclick = ()=>{ emailPage--; renderEmailPage(); };

  const info=document.createElement('span');
  info.textContent=` Page ${emailPage} of ${totalPages} `;
  info.style.margin='0 1em';

  const next=document.createElement('button');
  next.className='button'; next.textContent='Next →';
  next.disabled = emailPage===totalPages;
  next.onclick = ()=>{ emailPage++; renderEmailPage(); };

  pg.append(prev, info, next);
  emailPreview.parentNode.insertBefore(pg, emailPreview);
}

// 6) Send All stub
function onSendAll() {
  if(!confirm(`Send all ${scheduleData.length} emails now?`)) return;
  alert(`(Stub) Would send ${scheduleData.length} emails.`);
}
