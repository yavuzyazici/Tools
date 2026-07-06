/* =====================================================================
   Toplu Fatura Mail Gönderici — arayüz mantığı
   Python (pywebview) köprüsü: window.pywebview.api.*
   Backend'den arayüze itilen çağrılar: window.__onProgress/__onLog/...
   ===================================================================== */

const $ = (id) => document.getElementById(id);
let apiReady = false;
let running = false;
let headers = [];   // mevcut sayfanın başlık listesi

/* ---- Excel sütun harfi: 0->A, 1->B ... ---- */
function colLetter(i) {
  let s = ""; i += 1;
  while (i) { const r = (i - 1) % 26; s = String.fromCharCode(65 + r) + s; i = Math.floor((i - 1) / 26); }
  return s;
}

/* ---- Toast bildirimi ---- */
function toast(msg, type = "info", ms = 3200) {
  const wrap = $("toast-wrap");
  const el = document.createElement("div");
  el.className = "toast " + type;
  const icons = {
    error: '<path d="M12 9v4m0 4h.01M10.3 3.9L1.8 18a2 2 0 001.7 3h17a2 2 0 001.7-3L13.7 3.9a2 2 0 00-3.4 0z"/>',
    success: '<path d="M22 11.08V12a10 10 0 11-5.93-9.14"/><path d="M22 4L12 14.01l-3-3"/>',
    info: '<circle cx="12" cy="12" r="10"/><path d="M12 16v-4M12 8h.01"/>',
  };
  el.innerHTML = `<svg class="ic" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">${icons[type] || icons.info}</svg><span>${msg}</span>`;
  wrap.appendChild(el);
  setTimeout(() => { el.style.opacity = "0"; el.style.transition = "opacity .3s"; setTimeout(() => el.remove(), 300); }, ms);
}

/* ---- Onay modalı ---- */
function confirmModal(title, body, okLabel = "Başlat") {
  return new Promise((resolve) => {
    $("modal-title").textContent = title;
    $("modal-body").textContent = body;
    $("modal-ok").textContent = okLabel;
    $("modal").classList.remove("hidden");
    const done = (val) => { $("modal").classList.add("hidden"); $("modal-ok").onclick = null; $("modal-cancel").onclick = null; resolve(val); };
    $("modal-ok").onclick = () => done(true);
    $("modal-cancel").onclick = () => done(false);
  });
}

/* ---- Log ---- */
function logLine(level, msg) {
  const box = $("log");
  const div = document.createElement("div");
  div.className = "l-" + level;
  div.textContent = msg;
  box.appendChild(div);
  box.scrollTop = box.scrollHeight;
}

/* ---- Durum rozeti ---- */
function setBusy(busy) {
  running = busy;
  const s = $("status");
  if (busy) { s.className = "status-badge busy"; s.innerHTML = '<span class="dot"></span> Gönderiliyor...'; }
  else { s.className = "status-badge"; s.innerHTML = '<span class="dot"></span> Hazır'; }
  $("btn-start").disabled = busy;
  $("btn-check").disabled = busy;
  $("btn-stop").disabled = !busy;
}

/* ---- Sütun/sayfa açılır listelerini doldur ---- */
function fillColumns(hdrs) {
  headers = hdrs || [];
  const opts = headers.map((h, i) => `<option value="${i}">${colLetter(i)} — ${h || "(boş)"}</option>`).join("");
  $("email-col").innerHTML = opts;
  $("attach-col").innerHTML = opts;
}
function fillSheets(sheets, selected) {
  $("sheet").innerHTML = (sheets || []).map((s) => `<option value="${s}">${s}</option>`).join("");
  if (selected) $("sheet").value = selected;
}

/* ---- Alanları job nesnesine topla ---- */
function collectJob() {
  return {
    xlsx: $("xlsx").value.trim(),
    sheet: $("sheet").value,
    email_col: parseInt($("email-col").value, 10),
    attach_col: parseInt($("attach-col").value, 10),
    subject: $("subject").value,
    body: $("body").value,
    is_html: $("is-html").checked,
    method: currentMethod(),
    sender: $("sender").value.trim(),
    smtp_host: $("smtp-host").value.trim(),
    smtp_port: parseInt($("smtp-port").value, 10) || 587,
    smtp_security: $("smtp-security").value,
    smtp_user: $("smtp-user").value.trim(),
    smtp_password: $("smtp-pass").value,
    delay: parseFloat($("delay").value) || 0,
    limit: parseInt($("limit").value, 10) || 0,
    test: $("test").value.trim(),
  };
}

function currentMethod() {
  return document.querySelector("#method-seg button.on").dataset.method;
}

/* ---- Ayarları geri yükle ---- */
function applySettings(d) {
  if (!d) return;
  $("xlsx").value = d.xlsx || "";
  $("subject").value = d.subject || "";
  $("body").value = d.body || "";
  $("is-html").checked = !!d.is_html;
  $("sender").value = d.sender || "";
  $("smtp-host").value = d.smtp_host || "";
  $("smtp-port").value = d.smtp_port || "587";
  $("smtp-security").value = d.smtp_security || "starttls";
  $("smtp-user").value = d.smtp_user || "";
  $("smtp-pass").value = d.smtp_password || "";
  $("delay").value = d.delay != null ? d.delay : "1.0";
  $("limit").value = d.limit != null ? d.limit : "0";
  $("test").value = d.test || "";
  selectMethod(d.method || "outlook");
  return d; // sütunlar dosya okunduktan sonra ayarlanır
}

/* ---- Gönderim yöntemi segmenti ---- */
function selectMethod(m) {
  document.querySelectorAll("#method-seg button").forEach((b) => b.classList.toggle("on", b.dataset.method === m));
  $("smtp-panel").classList.toggle("hidden", m !== "smtp");
}

/* ===================================================================
   Olay çekme (pull) — worker thread'ten evaluate_js ile itmek yerine,
   arayüz olayları düzenli aralıkla çeker. pywebview'de donmayı önler.
   =================================================================== */
let pollTimer = null;
let polling = false;

function applyEvent(ev) {
  if (ev.t === "progress") {
    $("fill").style.width = ev.total ? (ev.done / ev.total * 100) + "%" : "0%";
    $("count").textContent = `${ev.done} / ${ev.total}  (✓${ev.ok} ✗${ev.fail})`;
  } else if (ev.t === "log") {
    logLine(ev.level, ev.msg);
  } else if (ev.t === "done") {
    const res = ev.result;
    stopPolling();
    setBusy(false);
    let msg = `Tamamlandı — Gönderilen: ${res.sent}, Hatalı: ${res.failed}`;
    if (res.stopped) msg = "Durduruldu — " + msg;
    logLine("info", "=== " + msg + " ===");
    toast(msg, res.failed ? "info" : "success", 5000);
  }
}

async function pump() {
  if (polling) return;            // önceki çağrı bitmeden yenisini başlatma
  polling = true;
  try {
    const evs = await window.pywebview.api.olaylari_al();
    if (evs && evs.length) evs.forEach(applyEvent);
  } catch (e) {}
  polling = false;
}

function startPolling() { if (!pollTimer) pollTimer = setInterval(pump, 200); }
function stopPolling() { if (pollTimer) { clearInterval(pollTimer); pollTimer = null; } setTimeout(pump, 0); }

/* ===================================================================
   API çağrıları
   =================================================================== */
async function readColumns(keepSel) {
  const path = $("xlsx").value.trim();
  if (!path) { toast("Önce bir Excel dosyası seçin.", "error"); return; }
  const emailSel = keepSel ? keepSel.email_col : null;
  const attachSel = keepSel ? keepSel.attach_col : null;
  const res = await window.pywebview.api.sutunlari_oku(path, $("sheet").value || null);
  if (res.error) { toast("Excel okunamadı: " + res.error, "error"); return; }
  fillSheets(res.sheets, res.sheet);
  fillColumns(res.headers);
  // Kayıtlı seçim varsa uygula, yoksa akıllı tahmin
  if (emailSel != null && emailSel < res.headers.length) $("email-col").value = emailSel;
  else guessColumn("email-col", ["email", "e-posta", "mail", "adres"]);
  if (attachSel != null && attachSel < res.headers.length) $("attach-col").value = attachSel;
  else guessColumn("attach-col", ["ek", "attach", "dosya", "pdf", "yol", "path", "fatura"]);
}

function guessColumn(selId, keywords) {
  for (let i = 0; i < headers.length; i++) {
    const h = (headers[i] || "").toLowerCase();
    if (keywords.some((k) => h.includes(k))) { $(selId).value = i; return; }
  }
}

async function saveSettings() {
  if (!apiReady) return;
  try { await window.pywebview.api.ayarlari_kaydet(collectJob()); } catch (e) {}
}

/* ===================================================================
   Olay bağlama
   =================================================================== */
function wire() {
  $("btn-browse").onclick = async () => {
    const p = await window.pywebview.api.sec_dosya();
    if (p) { $("xlsx").value = p; await readColumns(null); }
  };
  $("btn-read").onclick = () => readColumns(null);
  $("sheet").onchange = () => readColumns(collectJob());

  document.querySelectorAll("#method-seg button").forEach((b) => {
    b.onclick = () => selectMethod(b.dataset.method);
  });

  $("btn-load-html").onclick = async () => {
    const res = await window.pywebview.api.html_yukle();
    if (res.error) { toast(res.error, "error"); return; }
    if (res.content != null) { $("body").value = res.content; $("is-html").checked = true; toast("HTML içerik yüklendi.", "success"); }
  };
  $("btn-preview").onclick = async () => {
    const body = $("body").value;
    if (!body.trim()) { toast("Önizlenecek içerik yok.", "error"); return; }
    const res = await window.pywebview.api.onizle(body, $("is-html").checked);
    if (res && res.error) toast(res.error, "error");
  };

  $("btn-clear-log").onclick = () => { $("log").innerHTML = ""; };

  $("btn-check").onclick = async () => {
    const job = collectJob();
    if (!job.xlsx) { toast("Önce geçerli bir Excel dosyası seçin.", "error"); return; }
    logLine("info", "=== KONTROL (göndermeden tarama) ===");
    const rep = await window.pywebview.api.kontrol_et(job);
    if (rep.error) { toast(rep.error, "error"); return; }
    logLine("info", `Toplam satır: ${rep.total}  |  Sorunsuz: ${rep.ok}`);
    logLine("info", `Geçersiz e-posta: ${rep.bad_email.length}  |  Ek boş: ${rep.empty_attach.length}  |  Ek bulunamadı: ${rep.missing_attach.length}  |  Tekrarlı ek: ${rep.duplicate_attach.length}`);
    const dump = (title, items, level) => {
      if (!items.length) return;
      logLine(level, `— ${title} (${items.length}):`);
      items.slice(0, 50).forEach((it) => logLine(level, `    satır ${it[0]}: ${it[1] || "(boş)"} → ${it[3]}`));
      if (items.length > 50) logLine(level, `    ... ve ${items.length - 50} satır daha (tam liste CSV'de)`);
    };
    dump("Geçersiz e-posta", rep.bad_email, "error");
    dump("Ek yolu boş", rep.empty_attach, "error");
    dump("Ek dosyası bulunamadı", rep.missing_attach, "error");
    dump("Tekrarlı ek", rep.duplicate_attach, "info");
    const problems = rep.bad_email.length + rep.empty_attach.length + rep.missing_attach.length;
    if (problems === 0) toast(`✔ Her şey yolunda — ${rep.total} satırın tamamı geçerli.`, "success", 5000);
    else toast(`${problems} sorunlu satır bulundu. Ayrıntılar log alanında.`, "error", 5000);
  };

  $("btn-start").onclick = async () => {
    const job = collectJob();
    const problems = await window.pywebview.api.dogrula(job);
    if (problems && problems.length) { toast("Eksik/hatalı ayarlar: " + problems.join(", "), "error", 6000); return; }
    const n = job.limit ? String(job.limit) : "TÜMÜ";
    const mode = job.test ? `TEST → ${job.test}` : "GERÇEK gönderim";
    const ok = await confirmModal("Gönderimi başlat", `${mode}\nYöntem: ${job.method}\nGönderen: ${job.sender}\nKonu: ${job.subject}\nGönderilecek: ${n}`, "Başlat");
    if (!ok) return;
    await saveSettings();
    $("log").innerHTML = "";
    $("fill").style.width = "0%";
    logLine("info", "=== Gönderim başlatılıyor ===");
    setBusy(true);
    startPolling();
    const res = await window.pywebview.api.gonder(job);
    if (res && res.error) { stopPolling(); setBusy(false); toast(res.error, "error"); }
  };

  $("btn-stop").onclick = () => { window.pywebview.api.durdur(); logLine("info", "Durdurma istendi, mevcut mail bitince duracak..."); $("btn-stop").disabled = true; };

  // Kapanışta ayarları kaydet
  window.addEventListener("beforeunload", saveSettings);
}

/* ---- Başlangıç ---- */
async function init() {
  apiReady = true;
  wire();
  const d = await window.pywebview.api.ayarlari_yukle();
  const saved = applySettings(d);
  if (saved && saved.xlsx) {
    $("sheet").innerHTML = saved.sheet ? `<option>${saved.sheet}</option>` : "";
    if (saved.sheet) $("sheet").value = saved.sheet;
    await readColumns(saved);
  }
}

window.addEventListener("pywebviewready", init);
