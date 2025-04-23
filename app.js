// app.js

/**
 * Schedule Mailer Web App
 * Reads an Excel schedule, previews a weekly template,
 * copies the table, and previews email drafts.
 */

// -----------------------------
// Helper Functions
// -----------------------------

/**
 * Format a Date object as "MMM DD yy" (e.g. "Apr 28 25").
 */
function formatDateShort(date) {
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const d = date.getUTCDate();
  const m = date.getUTCMonth();
  const y = date.getUTCFullYear();
  const day = String(d).padStart(2, '0');
  const mon = months[m];
  const yy = String(y).slice(-2);
  return `${mon} ${day} ${yy}`;
}

/**
 * Format a Date object as "MMM DD yyyy" (e.g. "Apr 28 2025").
 */
function formatDateFull(date) {
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const d = date.getUTCDate();
  const m = date.getUTCMonth();
  const yyyy = date.getUTCFullYear();
  const day = String(d).padStart(2, '0');
  const mon = months[m];
  return `${mon} ${day} ${yyyy}`;
}

// -----------------------------
// Globals
// -----------------------------

let workbookGlobal = null;     // holds the loaded workbook
let dateRow        = [];       // array of date labels (short)
let headerRow      = [];       // array of header names
let rawRows        = [];       // array of data rows from Excel
let scheduleData   = [];       // mapped schedule entries for preview
let selectedHeaders= [];       // column labels to display

// -----------------------------
// DOM Ready Initialization
// -----------------------------

document.addEventListener('DOMContentLoaded', () => {
  // Element references
  const fileInput        = document.getElementById('fileInput');
  const weekStartInput   = document.getElementById('weekStart');
  const generateBtn      = document.getElementById('generateTemplate');
  const copyBtn          = document.getElementById('copyAll');
  const previewContainer = document.getElementById('preview');

  const generateEmailsBtn = document.getElementById('generateEmails');
  const sendAllBtn       = document.getElementById('sendAll');
  const emailPreview     = document.getElementById('emailPreview');

  // Initial UI state
  generateBtn.disabled      = true;
  if (copyBtn) copyBtn.style.display       = 'none';
  if (generateEmailsBtn) generateEmailsBtn.disabled = true;
  if (sendAllBtn) sendAllBtn.disabled     = true;

  // Event listeners
  fileInput.addEventListener('change', () =>
    onFileLoad(fileInput, generateBtn, copyBtn, previewContainer)
  );

  generateBtn.addEventListener('click', () =>
    onGeneratePreview(
      weekStartInput,
      generateBtn,
      copyBtn,
      previewContainer,
      generateEmailsBtn
    )
  );

  if (copyBtn) {
    copyBtn.addEventListener('click', () =>
      onCopyAll(previewContainer)
    );
  }

  if (generateEmailsBtn) {
    generateEmailsBtn.addEventListener('click', () =>
      onGenerateEmails(emailPreview, sendAllBtn)
    );
  }

  if (sendAllBtn) {
    sendAllBtn.addEventListener('click', onSendAll);
  }
});

// -----------------------------
// 1. Load the workbook and extract data
// -----------------------------

/**
 * Reads the selected Excel file, locates the "Schedule" sheet,
 * identifies header and date rows, and prepares for preview.
 */
function onFileLoad(fileInput, generateBtn, copyBtn, previewContainer) {
  const file = fileInput.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (evt) => {
    const data = new Uint8Array(evt.target.result);
    const wb = XLSX.read(data, { type: 'array', cellDates: true });
    workbookGlobal = wb;

    const ws = wb.Sheets['Schedule'];
    if (!ws) {
      alert('Sheet named "Schedule" not found.');
      return;
    }

    // Convert to array-of-arrays
    const arr = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

    // Find header row (must contain Team, Email, Employee)
    const headerIndex = arr.findIndex(row =>
      row.includes('Team') && row.includes('Email') && row.includes('Employee')
    );

    if (headerIndex < 1) {
      alert('Could not detect header row with Team, Email, Employee.');
      return;
    }

    // Build dateRow from the row above header
    dateRow = (arr[headerIndex - 1] || []).map(cell => {
      const d = new Date(cell);
      return isNaN(d) ? String(cell).trim() : formatDateShort(d);
    });

    // Set headerRow and rawRows for data
    headerRow = arr[headerIndex] || [];
    rawRows   = arr.slice(headerIndex + 1);

    // Reset preview UI
    previewContainer.innerHTML =
      '<p>File loaded. Select Week Start and click Generate Weekly Preview.</p>';

    // Enable preview button, hide copy
    generateBtn.disabled = false;
    if (copyBtn) copyBtn.style.display = 'none';
  };

  reader.readAsArrayBuffer(file);
}

// -----------------------------
// 2. Generate Weekly Template Preview
// -----------------------------

/**
 * Based on the selected week start, builds a 7-column table:
 * Email, Employee, and the five-day date range.
 */
function onGeneratePreview(
  weekStartInput,
  generateBtn,
  copyBtn,
  previewContainer,
  generateEmailsBtn
) {
  const startVal = weekStartInput.value;
  if (!startVal) {
    alert('Please pick a Week Start date.');
    return;
  }

  // Parse local date
  const [year, month, day] = startVal.split('-').map(Number);
  const startDate = new Date(year, month - 1, day);

  // Build arrays for five consecutive days
  const labelsShort = [];
  const labelsFull  = [];
  for (let i = 0; i < 5; i++) {
    const dt = new Date(startDate);
    dt.setDate(dt.getDate() + i);
    labelsShort.push(formatDateShort(dt));
    labelsFull.push(formatDateFull(dt));
  }

  // Locate start index in dateRow
  const startIdx = dateRow.indexOf(labelsShort[0]);
  if (startIdx < 0) {
    alert(`Date ${labelsShort[0]} not found in schedule.`);
    return;
  }

  // Five consecutive column indices
  const dateIndices = labelsShort
    .map((_, i) => startIdx + i)
    .filter(idx => idx >= 0 && idx < dateRow.length);

  // Identify key column indices
  const teamIdx  = headerRow.indexOf('Team');
  const emailIdx = headerRow.indexOf('Email');
  const empIdx   = headerRow.indexOf('Employee');
  if (teamIdx < 0 || emailIdx < 0 || empIdx < 0) {
    alert('Missing Team / Email / Employee columns.');
    return;
  }

  // Build selectedHeaders for table
  selectedHeaders = [
    headerRow[emailIdx],
    headerRow[empIdx],
    ...labelsFull
  ];

  // Map rawRows into scheduleData
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

  // Render preview table
  previewContainer.innerHTML = '';
  if (!scheduleData.length) {
    previewContainer.textContent = 'No matching rows for selected week.';
    return;
  }

  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const thr   = document.createElement('tr');
  selectedHeaders.forEach(h => {
    const th = document.createElement('th');
    th.textContent = h;
    thr.appendChild(th);
  });
  thead.appendChild(thr);
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

  // Show Copy button and enable email draft
  if (copyBtn) copyBtn.style.display = 'inline-block';
  if (generateEmailsBtn) generateEmailsBtn.disabled = false;
}

// -----------------------------
// 3. Copy entire table to clipboard
// -----------------------------

function onCopyAll(previewContainer) {
  const table = previewContainer.querySelector('table');
  if (!table) return;
  const range = document.createRange();
  range.selectNode(table);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
  document.execCommand('copy');
  sel.removeAllRanges();
  alert('Table copied to clipboard!');
}

// -----------------------------
// 4. Generate Email Drafts Preview
// -----------------------------

/**
 * Renders a draft email for each schedule entry.
 */
function onGenerateEmails(emailPreview, sendAllBtn) {
  emailPreview.innerHTML = '';

  scheduleData.forEach(entry => {
    const card = document.createElement('div');
    card.style.border   = '1px solid #ccc';
    card.style.padding  = '10px';
    card.style.margin   = '8px 0';

    // Subject line
    const subj = document.createElement('h3');
    subj.textContent = `Subject: Your Schedule (${selectedHeaders[2]} – ${selectedHeaders.slice(-1)})`;
    card.appendChild(subj);

    // To:
    const toLine = document.createElement('p');
    toLine.textContent = `To: ${entry[selectedHeaders[0]]}`;
    card.appendChild(toLine);

    // Body:
    const body = document.createElement('pre');
    let text = `Hello ${entry[selectedHeaders[1]]},\n\nHere is your schedule for the week:\n`;
    selectedHeaders.slice(2).forEach(day => {
      text += `- ${day}: ${entry[day] || 'OFF'}\n`;
    });
    text += '\nBest,\nYour Team';
    body.textContent = text;
    card.appendChild(body);

    emailPreview.appendChild(card);
  });

  if (sendAllBtn) sendAllBtn.disabled = false;
}

// -----------------------------
// 5. Send All Stub
// -----------------------------

function onSendAll() {
  alert(`(Stub) Would send ${scheduleData.length} emails now.`);
}
// ─────────────────────────────────────────────────────────────────────────────
// 4) Wiring up the Emails tab buttons (call this in your DOMContentLoaded block)
const generateEmailsBtn = document.getElementById('generateEmails');
const sendAllBtn        = document.getElementById('sendAll');
const emailPreview      = document.getElementById('emailPreview');

generateEmailsBtn.addEventListener('click', onGenerateEmails);
sendAllBtn.addEventListener('click', onSendAll);

// Enable the “Generate Email Drafts” button once we have a weekly preview:
function enableEmailDrafts() {
  generateEmailsBtn.disabled = scheduleData.length === 0;
}

// Call this at the end of onGeneratePreview():
enableEmailDrafts();


// ─────────────────────────────────────────────────────────────────────────────
// 5) Generate Email Drafts Preview
function onGenerateEmails() {
  emailPreview.innerHTML = '';
  if (!scheduleData.length) {
    emailPreview.textContent = 'No schedule data to generate emails.';
    return;
  }

  const subject = 'Schedule';
  // Loop each associate’s row
  scheduleData.forEach(row => {
    const to   = row[selectedHeaders[0]]; // Email
    const name = row[selectedHeaders[1]]; // Employee

    // Build table: first two header rows (dates, weekdays), then data row
    let tableHtml = `<table style="border-collapse:collapse;width:100%;margin:1em 0;">
      <thead>
        <tr><th style="border:1px solid #ddd;padding:6px;"></th>`;
    // Date header row
    selectedHeaders.slice(2).forEach(full => {
      tableHtml += `<th style="border:1px solid #ddd;padding:6px;">${full}</th>`;
    });
    tableHtml += `</tr><tr><th style="border:1px solid #ddd;padding:6px;"></th>`;
    // Weekday header row
    selectedHeaders.slice(2).forEach(full => {
      const dayName = new Date(full).toLocaleDateString('en-US',{weekday:'long'});
      tableHtml += `<th style="border:1px solid #ddd;padding:6px;">${dayName}</th>`;
    });
    tableHtml += `</tr>
      </thead>
      <tbody>
        <tr>
          <td style="border:1px solid #ddd;padding:6px;font-weight:600;">${name}</td>`;
    // Data row
    selectedHeaders.slice(2).forEach(full => {
      tableHtml += `<td style="border:1px solid #ddd;padding:6px;">${row[full]||''}</td>`;
    });
    tableHtml += `
        </tr>
      </tbody>
    </table>`;

    // Professional wrapper
    const bodyHtml = `
      <div style="font-family:Segoe UI,Arial,sans-serif; color:#333;">
        <p style="font-size:1rem; margin:0 0 1em 0;">
          Hi Team &ndash;
        </p>
        <p style="font-size:1rem; margin:0 0 1em 0;">
          Please see your schedule for next week below. If you have any questions, let us know.
        </p>
        ${tableHtml}
        <p style="font-size:1rem; margin:1em 0 0 0;">
          Thank you!
        </p>
      </div>`;

    // Render a preview card
    const card = document.createElement('div');
    card.className = 'email-card';
    card.innerHTML = `
      <h3 style="margin:0 0 .5em 0; font-size:1.1rem;">
        To: ${to}
      </h3>
      <p style="margin:0 0 .5em 0;">
        <strong>Subject:</strong> ${subject}
      </p>
      <div>${bodyHtml}</div>
    `;
    emailPreview.appendChild(card);
  });

  // Enable Send All
  sendAllBtn.disabled = false;
  // Switch to Emails tab
  document.querySelector('.tablinks[data-tab="emails"]').click();
}

// ─────────────────────────────────────────────────────────────────────────────
// 6) Send All (stub)
function onSendAll() {
  const count = scheduleData.length;
  if (!confirm(`Send all ${count} drafts now?`)) return;
  // TODO: integrate Microsoft Graph or backend send endpoint here
  alert(`(Stub) Would send ${count} emails.`);
}
