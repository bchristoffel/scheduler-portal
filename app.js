const msalConfig = {
  auth: {
    clientId: "5c90a7aa-6318-49bb-a0ab-aaccdca24ca6",
    authority: "https://login.microsoftonline.com/consumers",
    redirectUri: window.location.origin + window.location.pathname
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);
const loginRequest = { scopes: ["Mail.Send"] };

let workbookGlobal, dateRow = [], headerRow = [], rawRows = [];
let scheduleData = [], selectedHeaders = [];
let emailPage = 1, emailsPerPage = 10;

document.addEventListener("DOMContentLoaded", () => {
  const fileInput = document.getElementById("fileInput");
  const weekStartInput = document.getElementById("weekStart");
  const generateBtn = document.getElementById("generateTemplate");
  const copyBtn = document.getElementById("copyAll");
  const generateEmails = document.getElementById("generateEmails");
  const sendAllBtn = document.getElementById("sendAll");
  const previewContainer = document.getElementById("preview");

  generateBtn.disabled = true;
  if (copyBtn) copyBtn.style.display = "none";
  generateEmails.disabled = true;
  sendAllBtn.disabled = true;

  fileInput.addEventListener("change", () =>
    onFileLoad(fileInput, generateBtn, copyBtn, previewContainer)
  );

  generateBtn.addEventListener("click", () => {
    onGeneratePreview(weekStartInput, generateBtn, copyBtn, previewContainer);
    generateEmails.disabled = scheduleData.length === 0;
  });

  if (copyBtn)
    copyBtn.addEventListener("click", () => onCopyAll(previewContainer));

  generateEmails.addEventListener("click", () => {
    emailPage = 1;
    renderEmailPage();
    document.querySelector('.tablinks[data-tab="emails"]').click();
  });

  sendAllBtn.addEventListener("click", onSendAll);

  document.getElementById("closeConfirmation").addEventListener("click", () => {
    document.getElementById("confirmation").style.display = "none";
  });

  document.getElementById("refreshApp").addEventListener("click", () => {
    document.getElementById("confirmation").style.display = "none";
    window.location.reload();
  });

  document.getElementById("logoutBtn").addEventListener("click", () => {
    const account = msalInstance.getAllAccounts()[0];
    if (account) {
      msalInstance.logoutPopup({
        account,
        postLogoutRedirectUri: window.location.href
      });
    } else {
      alert("No active Microsoft login session.");
    }
  });
});

async function ensureToken() {
  const loginResponse = await msalInstance.loginPopup({ scopes: ["Mail.Send"] });
  const account = loginResponse.account;
  const tokenRes = await msalInstance.acquireTokenSilent({
    account,
    scopes: ["Mail.Send"]
  });
  return tokenRes.accessToken;
}
function formatDateShort(d) {
  const m = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  return `${m[d.getUTCMonth()]} ${String(d.getUTCDate()).padStart(2,"0")} ${String(d.getUTCFullYear()).slice(-2)}`;
}

function formatDateFull(d) {
  const m = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  return `${m[d.getUTCMonth()]} ${String(d.getUTCDate()).padStart(2,"0")} ${d.getUTCFullYear()}`;
}

function onFileLoad(fi, genBtn, copyBtn, preview) {
  const file = fi.files[0]; if (!file) return;
  const reader = new FileReader();
  reader.onload = evt => {
    const wb = XLSX.read(new Uint8Array(evt.target.result), { type:"array", cellDates:true });
    workbookGlobal = wb;
    const ws = wb.Sheets["Schedule"];
    if (!ws) return alert("Schedule tab not found");
    const arr = XLSX.utils.sheet_to_json(ws, { header:1, defval:"" });
    const hi = arr.findIndex(r => r.includes("Team") && r.includes("Email") && r.includes("Employee"));
    if (hi < 0) return alert("Header row not detected");
    dateRow = (arr[hi - 1] || []).map(c => {
      const d = new Date(c);
      return isNaN(d) ? String(c).trim() : formatDateShort(new Date(d.toUTCString()));
    });
    headerRow = arr[hi] || [];
    rawRows = arr.slice(hi + 1);
    preview.innerHTML = "<p>Select Week Start & click Generate Weekly Preview.</p>";
    genBtn.disabled = false;
    if (copyBtn) copyBtn.style.display = "none";
  };
  reader.readAsArrayBuffer(file);
}

function onGeneratePreview(wsInput, genBtn, copyBtn, preview) {
  const val = wsInput.value;
  if (!val) return alert("Select a Week Start date.");
  const [y, m, d] = val.split("-").map(Number);
  const start = new Date(Date.UTC(y, m - 1, d));
  const labelsShort = [], labelsFull = [];
  for (let i = 0; i < 5; i++) {
    const dt = new Date(start); dt.setUTCDate(dt.getUTCDate() + i);
    labelsShort.push(formatDateShort(dt));
    labelsFull.push(formatDateFull(dt));
  }
  const startIdx = dateRow.indexOf(labelsShort[0]);
  if (startIdx < 0) return alert(`Date ${labelsShort[0]} not found.`);
  const dateIndices = Array.from({ length: 5 }, (_, i) => startIdx + i).filter(i => i >= 0 && i < dateRow.length);

  const ti = headerRow.indexOf("Team"),
        ei = headerRow.indexOf("Email"),
        ui = headerRow.indexOf("Employee");
  if (ti < 0 || ei < 0 || ui < 0) return alert("Missing key columns.");

  selectedHeaders = [headerRow[ei], headerRow[ui], ...labelsFull];
  scheduleData = rawRows.filter(r => r[ti] && r[ti] !== "X").map(r => {
    const o = {
      [headerRow[ei]]: r[ei] || "",
      [headerRow[ui]]: r[ui] || ""
    };
    dateIndices.forEach((ci, j) => o[labelsFull[j]] = r[ci] || "");
    return o;
  });

  preview.innerHTML = "";
  if (!scheduleData.length) {
    preview.textContent = "No matching rows for the selected week.";
    return;
  }

  const table = document.createElement("table");
  const thead = document.createElement("thead");

  // Header row: Email, Name, Date labels
  const thr = document.createElement("tr");
  selectedHeaders.forEach(h => {
    const th = document.createElement("th");
    th.textContent = h;
    thr.appendChild(th);
  });
  thead.appendChild(thr);

  // Header row 2: Day of week labels
  const dayRow = document.createElement("tr");
  selectedHeaders.forEach((h, i) => {
    const th = document.createElement("th");
    if (i < 2) {
      th.textContent = "";
    } else {
      const dn = new Date(h).toLocaleDateString("en-US", { weekday: "long" });
      th.textContent = dn;
    }
    dayRow.appendChild(th);
  });
  thead.appendChild(dayRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  scheduleData.forEach(r => {
    const tr = document.createElement("tr");
    selectedHeaders.forEach(h => {
      const td = document.createElement("td");
      td.textContent = r[h] || "";
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  preview.appendChild(table);

  const cb = document.getElementById("copyAll");
  if (cb) cb.style.display = "inline-block";
}

function onCopyAll(preview) {
  const tbl = preview.querySelector("table");
  if (!tbl) return;
  const range = document.createRange(); range.selectNode(tbl);
  const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(range);
  document.execCommand("copy"); sel.removeAllRanges();
  alert("Preview table copied!");
}
function renderEmailPage() {
  const emailPreview = document.getElementById("emailPreview");
  const sendAllBtn = document.getElementById("sendAll");
  emailPreview.innerHTML = "";

  const total = scheduleData.length;
  const pages = Math.ceil(total / emailsPerPage);
  const startIdx = (emailPage - 1) * emailsPerPage;
  const pageData = scheduleData.slice(startIdx, startIdx + emailsPerPage);

  pageData.forEach(r => {
    const toAddr = r[selectedHeaders[0]];
    const name = r[selectedHeaders[1]];
    const subject = "Schedule";

    let tbl = `<table style="border-collapse:collapse;width:100%;margin:1em 0;"><thead><tr><th></th>`;
    selectedHeaders.slice(2).forEach(h => {
      tbl += `<th style="border:1px solid #ddd;padding:6px;">${h}</th>`;
    });
    tbl += `</tr>
<tr><th></th>`;
    selectedHeaders.slice(2).forEach(h => {
      const dn = new Date(h).toLocaleDateString("en-US", { weekday: "long" });
      tbl += `<th style="border:1px solid #ddd;padding:6px;">${dn}</th>`;
    });
    tbl += `</tr></thead><tbody>
<tr>
  <td rowspan="3" style="border:1px solid #ddd;padding:6px;font-weight:600;text-align:center;vertical-align:middle;">${name}</td>`;
    selectedHeaders.slice(2).forEach(h => {
      tbl += `<td style="border:1px solid #ddd;padding:6px;">${r[h] || ""}</td>`;
    });
    tbl += `</tr></tbody></table>`;

    const bodyHtml = `<div style="font-family:Segoe UI,Arial,sans-serif;color:#333;">
      <p>Hi Team &ndash;</p>
      <p>Please see your schedule for next week below. If you have any questions, let us know.</p>
      ${tbl}
      <p>Thank you!</p>
    </div>`;

    const card = document.createElement("div");
    card.className = "email-card";
    card.innerHTML = `
      <h3>To: ${toAddr}</h3>
      <p><strong>Subject:</strong> ${subject}</p>
      ${bodyHtml}
    `;
    emailPreview.appendChild(card);
  });

  renderPaginationControls(pages);
  sendAllBtn.disabled = false;
}

function renderPaginationControls(totalPages) {
  const emailPreview = document.getElementById("emailPreview");
  let pg = document.getElementById("emailPagination");
  if (pg) pg.remove();
  pg = document.createElement("div");
  pg.id = "emailPagination";
  pg.style.textAlign = "center";
  pg.style.marginTop = "1rem";

  const prev = document.createElement("button");
  prev.className = "button";
  prev.textContent = "← Prev";
  prev.disabled = emailPage === 1;
  prev.onclick = () => {
    emailPage--;
    renderEmailPage();
  };

  const info = document.createElement("span");
  info.textContent = ` Page ${emailPage} of ${totalPages} `;
  info.style.margin = "0 1em";

  const next = document.createElement("button");
  next.className = "button";
  next.textContent = "Next →";
  next.disabled = emailPage === totalPages;
  next.onclick = () => {
    emailPage++;
    renderEmailPage();
  };

  pg.append(prev, info, next);
  emailPreview.parentNode.insertBefore(pg, emailPreview.nextSibling);
}

async function onSendAll() {
  if (!scheduleData.length) return;
  const confirmSend = confirm(`Send all ${scheduleData.length} emails now?`);
  if (!confirmSend) return;

  let token;
  try {
    token = await ensureToken();
  } catch (err) {
    alert("Login failed. Cannot send emails.");
    console.error(err);
    return;
  }

  let failedCount = 0;

  for (const r of scheduleData) {
    const toAddr = r[selectedHeaders[0]];
    const name = r[selectedHeaders[1]];
    const subject = "Schedule";

    let tbl = `<table style="border-collapse:collapse;width:100%;margin:1em 0;"><thead><tr><th></th>`;
    selectedHeaders.slice(2).forEach(h => {
      tbl += `<th style="border:1px solid #ddd;padding:6px;">${h}</th>`;
    });
    tbl += `</tr>
<tr><th></th>`;
    selectedHeaders.slice(2).forEach(h => {
      const dn = new Date(h).toLocaleDateString("en-US", { weekday: "long" });
      tbl += `<th style="border:1px solid #ddd;padding:6px;">${dn}</th>`;
    });
    tbl += `</tr></thead><tbody>
<tr>
  <td rowspan="3" style="border:1px solid #ddd;padding:6px;font-weight:600;text-align:center;vertical-align:middle;">${name}</td>`;
    selectedHeaders.slice(2).forEach(h => {
      tbl += `<td style="border:1px solid #ddd;padding:6px;">${r[h] || ""}</td>`;
    });
    tbl += `</tr></tbody></table>`;

    const bodyHtml = `<div style="font-family:Segoe UI,Arial,sans-serif;color:#333;">
      <p>Hi Team &ndash;</p>
      <p>Please see your schedule for next week below. If you have any questions, let us know.</p>
      ${tbl}
      <p>Thank you!</p>
    </div>`;

    const message = {
      subject: subject,
      body: { contentType: "html", content: bodyHtml },
      toRecipients: [{ emailAddress: { address: toAddr } }]
    };

    const response = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ message })
    });

    if (!response.ok) {
      failedCount++;
      console.error(`❌ Failed to send to ${toAddr}`, await response.text());
    }
  }

  if (failedCount === 0) {
    document.querySelector(".tabcontent.active").style.display = "none";
    document.getElementById("confirmation").style.display = "flex";
  } else {
    alert(`⚠️ ${failedCount} of ${scheduleData.length} emails failed. See console.`);
  }
}
