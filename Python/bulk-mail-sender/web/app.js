/* =====================================================================
   Toplu Fatura Mail Gönderici — arayüz mantığı
   Python (pywebview) köprüsü: window.pywebview.api.*

   Akıcılık (kasmama) ilkeleri
   ---------------------------
   1. DOM'a satır satır yazılmaz. Log satırları bir diziye toplanır ve ekranın
      bir sonraki karesinde (requestAnimationFrame) tek seferde eklenir.
   2. Log alanı sınırlıdır (LOG_MAX). Binlerce <div> tarayıcıyı yavaşlatır;
      tam kayıt zaten gonderim_sonuclari.csv dosyasındadır.
   3. İlerleme çubuğu 'width' yerine 'transform' ile güncellenir: yerleşim
      (layout) değil yalnızca birleştirme (compositing) tetiklenir.
      Değer değişmediyse DOM'a hiç dokunulmaz.
   4. Köprü çağrıları uyarlanabilir aralıklıdır ve üst üste binmez: bir çağrı
      bitmeden yenisi başlamaz, olay yoksa aralık kendiliğinden uzar.
   5. Uzun işler Python tarafında arka planda çalışır; arayüz hiçbir zaman
      bir cevabı bekleyerek kilitlenmez.
   ===================================================================== */

const $ = (id) => document.getElementById(id);

const LOG_MAX = 800;          // log alanında tutulacak en fazla satır
const POLL_FAST = 120;        // olay akarken çekme aralığı (ms)
const POLL_SLOW = 400;        // sessizken çekme aralığı (ms)

let running = false;          // gönderim/kontrol sürüyor mu
let headers = [];             // mevcut sayfanın başlık listesi

/* Köprü hazır olmadan da çağrılabilen API vekili.
   Editör, pencere açılır açılmaz kurulur; köprü birkaç yüz ms sonra gelir.
   Vekil sayesinde editörün kodunda "hazır mı" kontrolü dolaşmaz. */
const api = new Proxy({}, {
  get(_t, ad) {
    return (...args) => {
      const a = window.pywebview && window.pywebview.api;
      if (!a || typeof a[ad] !== "function") {
        return Promise.reject(new Error("Python köprüsü henüz hazır değil"));
      }
      try { return Promise.resolve(a[ad](...args)); }
      catch (e) { return Promise.reject(e); }
    };
  },
});

/* ---- Excel sütun harfi: 0->A, 1->B ... ---- */
function colLetter(i) {
  let s = ""; i += 1;
  while (i) { const r = (i - 1) % 26; s = String.fromCharCode(65 + r) + s; i = Math.floor((i - 1) / 26); }
  return s;
}

/* ---- Metni HTML'e gömerken güvenli hale getir ---- */
function esc(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
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
  el.innerHTML = `<svg class="ic" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">${icons[type] || icons.info}</svg><span>${esc(msg)}</span>`;
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

/* ===================================================================
   Log — toplu (batch) yazım
   Satırlar önce diziye girer, bir sonraki karede tek DocumentFragment
   olarak eklenir. Böylece 500 satırlık bir akış 500 yerine 1 yerleşim
   (reflow) maliyeti çıkarır.
   =================================================================== */
let logQueue = [];
let logFlushQueued = false;

function logLine(level, msg) {
  logQueue.push([level, msg]);
  scheduleLogFlush();
}

function logLines(lines) {          // [[level, msg], ...]
  if (!lines || !lines.length) return;
  for (let i = 0; i < lines.length; i++) logQueue.push(lines[i]);
  scheduleLogFlush();
}

function scheduleLogFlush() {
  if (logFlushQueued) return;
  logFlushQueued = true;
  requestAnimationFrame(flushLog);
}

function flushLog() {
  logFlushQueued = false;
  if (!logQueue.length) return;

  const box = $("log");
  // Ekrana sığmayacak kadar çok satır geldiyse yalnızca son LOG_MAX tanesini
  // düğüme çevir: görünmeyecek satırlar için DOM üretmenin anlamı yok.
  let pending = logQueue;
  logQueue = [];
  if (pending.length > LOG_MAX) pending = pending.slice(-LOG_MAX);

  // Kullanıcı yukarı kaydırıp bir şey okuyorsa otomatik kaydırma yapmayız.
  const stick = box.scrollHeight - box.scrollTop - box.clientHeight < 40;

  const frag = document.createDocumentFragment();
  for (let i = 0; i < pending.length; i++) {
    const div = document.createElement("div");
    div.className = "l-" + pending[i][0];
    div.textContent = pending[i][1];
    frag.appendChild(div);
  }
  box.appendChild(frag);

  let over = box.childElementCount - LOG_MAX;
  while (over-- > 0) box.removeChild(box.firstChild);

  if (stick) box.scrollTop = box.scrollHeight;
}

function clearLog() {
  logQueue = [];
  $("log").textContent = "";
}

/* ===================================================================
   İlerleme çubuğu — transform ile, yalnızca değer değiştiğinde
   =================================================================== */
let lastRatio = -1;
let lastCount = "";

function setProgress(ratio, text) {
  const r = Math.max(0, Math.min(1, ratio || 0));
  if (Math.abs(r - lastRatio) >= 0.001) {
    $("fill").style.transform = "scaleX(" + r + ")";
    lastRatio = r;
  }
  if (text !== undefined && text !== lastCount) {
    $("count").textContent = text;
    lastCount = text;
  }
}

/* ---- Durum rozeti / buton durumları ---- */
function setBusy(busy, label) {
  running = busy;
  const s = $("status");
  if (busy) {
    s.className = "status-badge busy";
    s.innerHTML = '<span class="dot"></span> ' + esc(label || "Gönderiliyor...");
  } else {
    s.className = "status-badge";
    s.innerHTML = '<span class="dot"></span> Hazır';
  }
  $("btn-start").disabled = busy;
  $("btn-check").disabled = busy;
  $("btn-browse").disabled = busy;
  $("btn-read").disabled = busy;
  $("btn-stop").disabled = !busy;
}

/* Ek sütunu seçilmediğini belirten değer (core.NO_ATTACH_COL ile aynı). */
const EK_YOK = -1;

/* ---- Sütun/sayfa açılır listelerini doldur ---- */
function fillColumns(hdrs) {
  headers = hdrs || [];
  const opts = headers.map((h, i) => `<option value="${i}">${colLetter(i)} — ${esc(h || "(boş)")}</option>`).join("");
  $("email-col").innerHTML = opts;
  // Ek isteğe bağlıdır: listenin başına "ek yok" seçeneği konur.
  $("attach-col").innerHTML = `<option value="${EK_YOK}">— Ek gönderme —</option>` + opts;
}
function fillSheets(sheets, selected) {
  $("sheet").innerHTML = (sheets || []).map((s) => `<option value="${esc(s)}">${esc(s)}</option>`).join("");
  if (selected) $("sheet").value = selected;
}

/* ---- Alanları job nesnesine topla ---- */
function collectJob() {
  return {
    xlsx: $("xlsx").value.trim(),
    sheet: $("sheet").value,
    email_col: parseInt($("email-col").value, 10),
    attach_col: parseInt($("attach-col").value, 10),   // -1 = ek gönderme
    subject: $("subject").value,
    body: MailEditor.getValue(),
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
  if (!d) return null;
  $("xlsx").value = d.xlsx || "";
  $("subject").value = d.subject || "";
  $("is-html").checked = d.is_html !== false;      // varsayılan: HTML
  MailEditor.setHtmlMode($("is-html").checked);
  MailEditor.setValue(d.body || "");
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
   Olay çekme (pull)
   Worker thread'ten evaluate_js ile itmek pywebview'de kilitlenmeye yol
   açtığı için olaylar Python tarafında birikir, arayüz toplu çeker.
   Aralık uyarlanabilir: olay geldikçe hızlanır, sessizlikte yavaşlar.
   =================================================================== */
let pollTimer = null;
let polling = false;
let pollDelay = POLL_FAST;

/* Tek bir zamanlayıcı yaşar: zincir asla ne kopar ne de ikiye bölünür. */
function schedulePump(ms) {
  if (!running || pollTimer !== null) return;
  pollTimer = setTimeout(pump, ms);
}

function startPolling() {
  pollDelay = POLL_FAST;
  schedulePump(0);
}

function stopPolling() {
  if (pollTimer !== null) { clearTimeout(pollTimer); pollTimer = null; }
}

async function pump() {
  pollTimer = null;
  // Önceki çağrı hâlâ sürüyorsa üstüne binme; ama zinciri de koparma.
  if (polling) { schedulePump(POLL_FAST); return; }
  polling = true;
  let hadEvents = false;
  try {
    const ev = await window.pywebview.api.olaylari_al();
    hadEvents = applyEvents(ev);
  } catch (e) {
    /* pencere kapanıyor olabilir; sessizce geç */
  } finally {
    polling = false;
  }
  // Olay aktıkça sık, sessizlikte seyrek çek: boşa köprü çağrısı yapma.
  pollDelay = hadEvents ? POLL_FAST : Math.min(POLL_SLOW, pollDelay + 60);
  schedulePump(pollDelay);
}

/* Son bir kez çekip kalan olayları boşalt (iş bittikten sonra). */
function drainOnce() {
  setTimeout(async () => {
    try { applyEvents(await window.pywebview.api.olaylari_al()); } catch (e) {}
  }, 0);
}

function applyEvents(ev) {
  if (!ev) return false;
  let any = false;

  if (ev.logs && ev.logs.length) { logLines(ev.logs); any = true; }
  if (ev.dropped) { logLine("info", `(hızlı akış: ${ev.dropped} log satırı ekranda gösterilmedi — tam kayıt CSV'de)`); any = true; }

  if (ev.progress) {
    const [done, total, ok, fail] = ev.progress;
    setProgress(total ? done / total : 0, `${done} / ${total}  (✓${ok} ✗${fail})`);
    any = true;
  }
  if (ev.scan) {
    const [done, total] = ev.scan;
    setProgress(total ? done / total : 0, `Kontrol: ${done} / ${total}`);
    any = true;
  }
  if (ev.finals && ev.finals.length) {
    ev.finals.forEach(applyFinal);
    any = true;
  }
  return any;
}

function applyFinal(f) {
  stopPolling();
  setBusy(false);
  if (f.t === "done") {
    const res = f.result;
    let msg = `Tamamlandı — Gönderilen: ${res.sent}, Hatalı: ${res.failed}`;
    if (res.stopped) msg = "Durduruldu — " + msg;
    logLine("info", "=== " + msg + " ===");
    toast(msg, res.failed ? "info" : "success", 5000);
  } else if (f.t === "check_done") {
    renderCheckReport(f.report);
  } else if (f.t === "error") {
    logLine("error", "Beklenmeyen hata: " + f.msg);
    toast("Beklenmeyen hata: " + f.msg, "error", 6000);
  }
  flushLog();
}

/* ---- Ön kontrol raporu ---- */
function renderCheckReport(rep) {
  if (!rep) return;
  if (rep.stopped) logLine("info", "Kontrol kullanıcı tarafından durduruldu (kısmi sonuç).");
  logLine("info", `Toplam satır: ${rep.total}  |  Gönderilebilir: ${rep.ok}`);
  if (rep.attach_used === false)
    logLine("info", "Ek sütunu seçilmedi — mailler eksiz gönderilecek.");
  else
    logLine("info", `Geçersiz e-posta: ${rep.bad_email.length}  |  Eki olmayan: ${rep.empty_attach.length}  |  Ek bulunamadı: ${rep.missing_attach.length}  |  Tekrarlı ek: ${rep.duplicate_attach.length}`);

  const bosAlan = rep.empty_field || [];
  const bilinmeyen = rep.unknown_fields || [];
  if (rep.used_fields && rep.used_fields.length)
    logLine("info", `Kişiselleştirme: ${rep.used_fields.length} alan kullanılıyor → ${rep.used_fields.join(", ")}`);
  if (bilinmeyen.length)
    logLine("error", `— Tanınmayan alan (${bilinmeyen.length}): ${bilinmeyen.map((x) => "{" + x + "}").join(", ")}  ← bu adda sütun yok, metinde olduğu gibi kalır`);

  const dump = (title, items, level) => {
    if (!items.length) return;
    logLine(level, `— ${title} (${items.length}):`);
    const shown = Math.min(items.length, 50);
    for (let i = 0; i < shown; i++) {
      const it = items[i];
      logLine(level, `    satır ${it[0]}: ${it[1] || "(boş)"} → ${it[3]}`);
    }
    if (items.length > shown) logLine(level, `    ... ve ${items.length - shown} satır daha (tam liste CSV'de)`);
  };
  dump("Geçersiz e-posta", rep.bad_email, "error");
  // Ek isteğe bağlıdır: eksiz satır gönderimi ENGELLEMEZ, bu yüzden bilgi satırı.
  dump("Eki olmayan satır (eksiz gönderilecek)", rep.empty_attach, "info");
  dump("Ek dosyası bulunamadı", rep.missing_attach, "error");
  dump("Tekrarlı ek", rep.duplicate_attach, "info");
  dump("Yer tutucu değeri boş", bosAlan, "info");
  if (rep.csv) logLine("info", "Sorun listesi kaydedildi: " + rep.csv);

  setProgress(1, `${rep.total} satır tarandı`);
  // Yalnızca gönderimi engelleyenler "sorun" sayılır.
  const problems = rep.bad_email.length + rep.missing_attach.length + bilinmeyen.length;
  const notlar = bosAlan.length + rep.empty_attach.length;
  if (problems === 0 && !notlar) toast(`✔ Her şey yolunda — ${rep.total} satırın tamamı geçerli.`, "success", 5000);
  else if (problems === 0) toast(`✔ Gönderime engel yok — ${notlar} satır için not var, log alanına bakın.`, "info", 5000);
  else toast(`${problems} sorun bulundu. Ayrıntılar log alanında.`, "error", 5000);
}

/* ===================================================================
   API çağrıları
   =================================================================== */
let readingColumns = false;

async function readColumns(keepSel) {
  const path = $("xlsx").value.trim();
  if (!path) { toast("Önce bir Excel dosyası seçin.", "error"); return; }
  if (readingColumns) return;
  readingColumns = true;
  const s = $("status");
  const prev = s.innerHTML;
  s.innerHTML = '<span class="dot"></span> Excel okunuyor...';
  s.className = "status-badge busy";
  try {
    const emailSel = keepSel ? keepSel.email_col : null;
    const attachSel = keepSel ? keepSel.attach_col : null;
    const res = await window.pywebview.api.sutunlari_oku(path, $("sheet").value || null);
    if (res.error) { toast("Excel okunamadı: " + res.error, "error"); return; }
    fillSheets(res.sheets, res.sheet);
    fillColumns(res.headers);
    // Kayıtlı seçim varsa uygula, yoksa akıllı tahmin
    if (emailSel != null && !isNaN(emailSel) && emailSel < res.headers.length) $("email-col").value = emailSel;
    else guessColumn("email-col", ["email", "e-posta", "mail", "adres"]);
    // Ek sütunu: kayıtlı seçim (-1 "ek yok" dahil) korunur; yoksa tahmin edilir.
    // Tahmin tutmazsa "ek gönderme"de kalır — rastgele bir sütunu ek yolu sanıp
    // her maile yanlış dosya iliştirmektense eksiz göndermek doğrudur.
    if (attachSel != null && !isNaN(attachSel) && attachSel < res.headers.length) $("attach-col").value = attachSel;
    else if (!guessColumn("attach-col", ["ek", "attach", "dosya", "pdf", "yol", "path", "fatura"]))
      $("attach-col").value = EK_YOK;
    refreshEditorFields();
  } finally {
    readingColumns = false;
    if (!running) { s.className = "status-badge"; s.innerHTML = '<span class="dot"></span> Hazır'; }
    else s.innerHTML = prev;
  }
}

/* Editöre sütun listesini ve ilk satırın örnek değerlerini verir.
   Böylece "Alan Ekle" menüsü gerçek sütunları, rozetler de gerçek değerleri
   gösterir ({Ad Soyad} rozetinin üstüne gelince "→ Ahmet Yılmaz" yazar). */
async function refreshEditorFields() {
  try {
    const res = await api.alanlar(collectJob());
    MailEditor.setFields(res);
  } catch (e) { /* köprü yoksa sessiz geç */ }
  try {
    const s = await api.ornek_satirlar(collectJob(), 1);
    MailEditor.setSample(s && s.rows && s.rows.length ? s.rows[0].values : null);
  } catch (e) {}
}

/* Türkçe başlıkları karşılaştırılabilir hale getirir: "Ek Dosyası" -> "ek dosyasi" */
function normBaslik(s) {
  return String(s || "").toLowerCase()
    .replace(/ı/g, "i").replace(/ş/g, "s").replace(/ğ/g, "g")
    .replace(/ü/g, "u").replace(/ö/g, "o").replace(/ç/g, "c")
    .replace(/\s+/g, " ").trim();
}

/* Başlıklarda anahtar kelime arar. Bulduysa seçer ve true döner.
   Sıra ÖNEMLİ: önce tam eşleşme, sonra kelime eşleşmesi, en son içerme.
   Aksi halde "Fatura No" başlığı, 'fatura' anahtarı yüzünden asıl "Ek"
   sütunundan önce seçiliyor ve fatura NUMARASI dosya yolu sanılıyordu. */
function guessColumn(selId, keywords) {
  const hn = headers.map(normBaslik);
  const kn = keywords.map(normBaslik);
  const sec = (i) => { $(selId).value = i; return true; };

  for (const k of kn) { const i = hn.indexOf(k); if (i >= 0) return sec(i); }
  for (const k of kn) {
    const re = new RegExp("(^|[^a-z0-9])" + k.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "([^a-z0-9]|$)");
    const i = hn.findIndex((h) => re.test(h));
    if (i >= 0) return sec(i);
  }
  for (const k of kn) { const i = hn.findIndex((h) => h.indexOf(k) >= 0); if (i >= 0) return sec(i); }
  return false;
}

async function saveSettings() {
  try { await window.pywebview.api.ayarlari_kaydet(collectJob()); } catch (e) {}
}

/* ===================================================================
   Olay bağlama
   =================================================================== */
function wire() {
  // Editör köprüden ÖNCE kurulur: pencere ilk kareden itibaren dolu görünür.
  MailEditor.mount($("editor-mount"), {
    api: api,
    getJob: collectJob,
    onToast: toast,
    onChange: () => {},
  });

  $("btn-browse").onclick = async () => {
    const p = await window.pywebview.api.sec_dosya();
    if (p) { $("xlsx").value = p; $("sheet").innerHTML = ""; await readColumns(null); }
  };
  $("btn-read").onclick = () => readColumns(null);
  $("sheet").onchange = () => readColumns(collectJob());

  document.querySelectorAll("#method-seg button").forEach((b) => {
    b.onclick = () => selectMethod(b.dataset.method);
  });

  $("is-html").onchange = () => MailEditor.setHtmlMode($("is-html").checked);
  $("btn-subject-field").onclick = (e) => MailEditor.openFields(e.currentTarget, "subject");

  $("btn-clear-log").onclick = clearLog;

  $("btn-check").onclick = async () => {
    const job = collectJob();
    if (!job.xlsx) { toast("Önce geçerli bir Excel dosyası seçin.", "error"); return; }
    clearLog();
    logLine("info", "=== KONTROL (göndermeden tarama) ===");
    setProgress(0, "Kontrol ediliyor...");
    setBusy(true, "Kontrol ediliyor...");
    startPolling();
    const res = await window.pywebview.api.kontrol_et(job);
    if (res && res.error) { stopPolling(); setBusy(false); toast(res.error, "error"); }
  };

  $("btn-start").onclick = async () => {
    const job = collectJob();
    const problems = await window.pywebview.api.dogrula(job);
    if (problems && problems.length) { toast("Eksik/hatalı ayarlar: " + problems.join(", "), "error", 6000); return; }
    const n = job.limit ? String(job.limit) : "TÜMÜ";
    const mode = job.test ? `TEST → ${job.test}` : "GERÇEK gönderim";
    // Ek isteğe bağlı olduğu için onay kutusunda açıkça yazılır: kullanıcı
    // "ek gönderme" seçtiğini fark etmeden 1300 mail atmasın.
    const ekBilgi = job.attach_col >= 0
      ? `${colLetter(job.attach_col)} sütunundaki dosya`
      : "yok (eksiz gönderilecek)";
    const ok = await confirmModal(
      "Gönderimi başlat",
      `${mode}\nYöntem: ${job.method}\nGönderen: ${job.sender}\nKonu: ${job.subject}\nEk: ${ekBilgi}\nGönderilecek: ${n}`,
      "Başlat");
    if (!ok) return;
    await saveSettings();
    clearLog();
    setProgress(0, "0 / 0");
    logLine("info", "=== Gönderim başlatılıyor ===");
    setBusy(true, "Gönderiliyor...");
    startPolling();
    const res = await window.pywebview.api.gonder(job);
    if (res && res.error) { stopPolling(); setBusy(false); toast(res.error, "error"); }
  };

  $("btn-stop").onclick = () => {
    window.pywebview.api.durdur();
    logLine("info", "Durdurma istendi, mevcut işlem bitince duracak...");
    $("btn-stop").disabled = true;
    drainOnce();
  };

  // Kapanışta ayarları kaydet
  window.addEventListener("beforeunload", saveSettings);
}

/* ===================================================================
   Açılış
   Sıra önemlidir: önce perde + butonlar, sonra köprü, EN SON Excel.
   Hiçbir adım bir diğerini bekletmez; pencere ilk kareden itibaren canlıdır.
   =================================================================== */
function bootMsg(msg, isError) {
  const el = $("boot-msg");
  if (!el) return;
  el.textContent = msg;
  el.className = "boot-msg" + (isError ? " err" : "");
}

function bootKapat() {
  const b = $("boot");
  if (!b) return;
  b.classList.add("gone");
  setTimeout(() => { if (b.parentNode) b.parentNode.removeChild(b); }, 300);
}

function apiHazirMi() {
  return !!(window.pywebview && window.pywebview.api && window.pywebview.api.hazir);
}

/* Köprüyü bekler. 'pywebviewready' olayı bizden önce tetiklenmiş olabileceği
   için ayrıca yoklama da yapılır — yalnızca olaya güvenmek sonsuz bekleme riskidir. */
function apiBekle() {
  return new Promise((resolve, reject) => {
    if (apiHazirMi()) return resolve();
    let gecen = 0;
    const t = setInterval(() => {
      gecen += 100;
      if (apiHazirMi()) { clearInterval(t); resolve(); return; }
      if (gecen === 2500) bootMsg("Python köprüsü bekleniyor…");
      if (gecen === 8000) bootMsg("Hâlâ bekleniyor… (WebView2 ilk açılışta yavaş olabilir)");
      if (gecen >= 25000) { clearInterval(t); reject(new Error("zaman aşımı")); }
    }, 100);
    window.addEventListener("pywebviewready", () => {
      if (apiHazirMi()) { clearInterval(t); resolve(); }
    });
  });
}

async function boot() {
  wire();                                   // butonlar daha köprü gelmeden bağlanır
  try {
    await apiBekle();
  } catch (e) {
    bootMsg("Python köprüsü kurulamadı. Uygulamayı kapatıp yeniden açın. "
            + "Sorun sürerse %APPDATA%\\TopluFaturaMailer\\calisma.log dosyasına bakın.", true);
    return;                                 // perdeyi kapatmıyoruz: sahte arayüz göstermeyelim
  }
  try { await window.pywebview.api.hazir(); } catch (e) {}
  bootKapat();
  init();                                   // await YOK: arayüz beklemesin
}

async function init() {
  let d = null;
  try { d = await window.pywebview.api.ayarlari_yukle(); } catch (e) {}
  const saved = applySettings(d);
  if (saved && saved.xlsx) {
    $("sheet").innerHTML = saved.sheet ? `<option>${esc(saved.sheet)}</option>` : "";
    if (saved.sheet) $("sheet").value = saved.sheet;
    // Excel okuma arka planda: kopuk bir ağ yolu bile arayüzü kilitlemez,
    // durum rozetinde "Excel okunuyor..." görünür.
    readColumns(saved);
  } else {
    // Excel yoksa bile hazır alanlar ({SATIR}, {GONDERIM_TARIHI}…) ve
    // biçim listesi editöre gitsin.
    refreshEditorFields();
  }
}

if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", boot);
else boot();
