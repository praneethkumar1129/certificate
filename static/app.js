const form = document.getElementById("certificateForm");
const excelInput = document.getElementById("excelFile");
const templateInput = document.getElementById("templateFile");
const message = document.getElementById("message");
const previewBtn = document.getElementById("previewBtn");
const confirmSection = document.getElementById("confirmSection");
const reportSection = document.getElementById("reportSection");
const confirmCheckbox = document.getElementById("confirmCheckbox");
const confirmSendBtn = document.getElementById("confirmSendBtn");
let pendingRecipients = [];
let lastJobPayload = null;

excelInput.addEventListener("change", () => {
  document.getElementById("excelLabel").textContent = excelInput.files[0]?.name || "Choose Excel or CSV file";
});
templateInput.addEventListener("change", () => {
  document.getElementById("templateLabel").textContent = templateInput.files[0]?.name || "Use the included participation certificate";
});

function showMessage(text, kind = "") {
  message.hidden = false;
  message.className = "message " + kind;
  message.textContent = text;
}
function hideMessage() { message.hidden = true; }
function makeFormData(includeConfirm = false) {
  const data = new FormData();
  data.append("event_name", document.getElementById("eventName").value.trim());
  data.append("excel", excelInput.files[0]);
  if (templateInput.files[0]) data.append("template", templateInput.files[0]);
  if (includeConfirm) data.append("confirm_send", "true");
  return data;
}
function validateForm() {
  if (!document.getElementById("eventName").value.trim()) {
    showMessage("Please enter the event name.", "error"); return false;
  }
  if (!excelInput.files[0]) {
    showMessage("Please choose your Excel or CSV registration file.", "error"); return false;
  }
  return true;
}

previewBtn.addEventListener("click", async () => {
  hideMessage();
  if (!validateForm()) return;
  previewBtn.disabled = true;
  previewBtn.textContent = "Preparing preview…";
  try {
    const response = await fetch("/api/preview", { method: "POST", body: makeFormData() });
    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      throw new Error(err.detail || "Could not generate preview.");
    }
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    window.open(url, "_blank", "noopener");
    showMessage("Preview opened in a new tab. No emails have been sent.", "good");
  } catch (err) {
    showMessage(err.message, "error");
  } finally {
    previewBtn.disabled = false;
    previewBtn.innerHTML = 'Preview certificate <span>↗</span>';
  }
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  hideMessage();
  if (!validateForm()) return;
  const sendBtn = document.getElementById("sendBtn");
  sendBtn.disabled = true;
  sendBtn.textContent = "Checking recipients…";
  try {
    const response = await fetch("/api/recipients", { method: "POST", body: makeFormData() });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(payload.detail || "Could not validate the spreadsheet.");
    pendingRecipients = payload.recipients || [];
    const note = `${payload.eligible} eligible recipient(s) found out of ${payload.total} student record(s). Review the list below. Rows with missing or invalid details will be skipped.`;
    document.getElementById("confirmText").textContent = note;
    const preview = document.getElementById("recipientPreview");
    preview.innerHTML = "";
    pendingRecipients.forEach(person => {
      const card = document.createElement("div");
      card.className = "recipient";
      card.innerHTML = `<strong>${escapeHtml(person.name || "Missing name")}</strong>
        <small>${escapeHtml(person.email || "Missing email")}${person.college ? " · " + escapeHtml(person.college) : ""}</small>
        ${person.issue ? `<small style="color:#b42318">${escapeHtml(person.issue)}</small>` : ""}`;
      preview.appendChild(card);
    });
    confirmCheckbox.checked = false;
    confirmSendBtn.disabled = true;
    confirmSection.hidden = false;
    confirmSection.scrollIntoView({behavior:"smooth", block:"start"});
  } catch (err) {
    showMessage(err.message, "error");
  } finally {
    sendBtn.disabled = false;
    sendBtn.innerHTML = 'Review &amp; send <span>→</span>';
  }
});

confirmCheckbox.addEventListener("change", () => {
  confirmSendBtn.disabled = !confirmCheckbox.checked;
});

confirmSendBtn.addEventListener("click", async () => {
  if (!confirmCheckbox.checked) return;
  confirmSendBtn.disabled = true;
  confirmSendBtn.textContent = "Sending certificates… please wait";
  showMessage("Sending is in progress. Keep this tab open until the report appears.");
  try {
    const response = await fetch("/api/send", { method: "POST", body: makeFormData(true) });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(payload.detail || "The send operation failed.");
    lastJobPayload = payload;
    renderReport(payload);
    confirmSection.hidden = true;
    reportSection.hidden = false;
    reportSection.scrollIntoView({behavior:"smooth", block:"start"});
    showMessage(`Finished: ${payload.sent} sent, ${payload.failed} failed, ${payload.skipped} skipped.`, payload.failed ? "" : "good");
  } catch (err) {
    showMessage(err.message, "error");
  } finally {
    confirmSendBtn.disabled = !confirmCheckbox.checked;
    confirmSendBtn.innerHTML = 'Send certificates now <span>→</span>';
  }
});

function renderReport(payload) {
  document.getElementById("jobDate").textContent = payload.created_at || "";
  document.getElementById("totalCount").textContent = payload.total;
  document.getElementById("sentCount").textContent = payload.sent;
  document.getElementById("failedCount").textContent = payload.failed;
  document.getElementById("skippedCount").textContent = payload.skipped;
  document.getElementById("zipLink").href = payload.download_url;
  document.getElementById("csvLink").href = payload.report_url;
  const tbody = document.getElementById("reportRows");
  tbody.innerHTML = "";
  (payload.report || []).forEach(row => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${escapeHtml(row.name || "—")}</td>
      <td>${escapeHtml(row.email || "—")}</td>
      <td><span class="badge ${escapeHtml(row.status)}">${escapeHtml(row.status)}</span></td>
      <td>${escapeHtml(row.message || "")}</td>`;
    tbody.appendChild(tr);
  });
}
function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, ch => ({
    "&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;", "'":"&#39;"
  }[ch]));
}
