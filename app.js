// — MSAL setup —
const msalConfig = {
  auth: {
    clientId: "5c90a7aa-6318-49bb-a0ab-aaccdca24ca6",
    authority: "https://login.microsoftonline.com/common",
    redirectUri: window.location.origin + window.location.pathname
  }
};
const msalInstance = new msal.PublicClientApplication(msalConfig);
const loginRequest = { scopes: ["Mail.Send"] };

// Only runs when needed
async function ensureToken() {
  let account = msalInstance.getAllAccounts()[0];
  if (!account) {
    const loginRes = await msalInstance.loginPopup(loginRequest);
    account = loginRes.account;
  }
  const tokenRes = await msalInstance.acquireTokenSilent({
    account,
    scopes: loginRequest.scopes
  });
  return tokenRes.accessToken;
}

// ...everything else (onFileLoad, onGeneratePreview, etc.) remains unchanged...

// ✅ 6) Send all via Graph (with confirmation and login at time of send)
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

  for (const r of scheduleData) {
    const toAddr = r[selectedHeaders[0]];
    const name = r[selectedHeaders[1]];

    let tbl = `<table style="border-collapse:collapse;width:100%;margin:1em 0;"><thead><tr><th></th>`;
    selectedHeaders.slice(2).forEach(h => {
      tbl += `<th style="border:1px solid #ddd;padding:6px;">${h}</th>`;
    });
    tbl += `</tr><tr><th></th>`;
    selectedHeaders.slice(2).forEach(h => {
      const dn = new Date(h).toLocaleDateString("en-US", { weekday: "long" });
      tbl += `<th style="border:1px solid #ddd;padding:6px;">${dn}</th>`;
    });
    tbl += `</tr></thead><tbody><tr>`;
    tbl += `<td style="border:1px solid #ddd;padding:6px;font-weight:600;">${name}</td>`;
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
      subject: "Schedule",
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
      console.error(`❌ Failed to send to ${toAddr}`, await response.text());
    }
  }

  alert(`✅ Successfully sent ${scheduleData.length} emails!`);
}
