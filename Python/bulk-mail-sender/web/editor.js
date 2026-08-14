/* =====================================================================
   Mail HTML Editörü  —  window.MailEditor

   Tasarım kararları
   -----------------
   1. DÜZENLEME YÜZEYİ BİR IFRAME'DİR. Mailin kendi CSS'i (Word şablonlarındaki
      <style> blokları dahil) uygulamanın arayüzüne sızmamalıdır; ayrıca <head>,
      <style> ve doctype'ın korunabilmesi için tam bir belge gerekir.
   2. KAYNAK METİN ASIL DOĞRUDUR. Kullanıcı tasarım görünümüne hiç girmediyse
      HTML'i harfi harfine korunur — hiçbir "normalleştirme" yapılmaz. Yalnızca
      tasarımda düzenleme yapıldığında iframe'den yeniden üretilir.
   3. YER TUTUCU ROZETİ METNİ, YER TUTUCUNUN KENDİSİDİR. {Ad} rozeti
      <span>{Ad}</span> olarak durur; rozet temizliği bir sebeple atlanırsa bile
      çıktı yine doğrudur (bozulma riski sıfır).
   4. TANINMAYAN {..} YAZIMINA DOKUNULMAZ. CSS blokları da süslü parantez
      içerir; bu kural Python tarafındaki merge.py ile birebir aynıdır.
   5. ÖNİZLEME PYTHON'DA ÜRETİLİR. Önizlemede görülen metin, gönderilecek
      mailin metnidir (aynı merge.Template kodu çalışır).
   ===================================================================== */

window.MailEditor = (function () {
  "use strict";

  /* ---------------------------------------------------------------- ikonlar */
  const I = {
    undo: '<path d="M3 7v6h6"/><path d="M3.5 13a9 9 0 102.1-9.4L3 7"/>',
    redo: '<path d="M21 7v6h-6"/><path d="M20.5 13a9 9 0 11-2.1-9.4L21 7"/>',
    bold: '<path d="M6 4h8a4 4 0 010 8H6zM6 12h9a4 4 0 010 8H6z"/>',
    italic: '<path d="M19 4h-9M14 20H5M15 4L9 20"/>',
    underline: '<path d="M6 3v7a6 6 0 0012 0V3M4 21h16"/>',
    strike: '<path d="M4 12h16M17.5 7a4.5 4.5 0 00-4.5-3h-1a4 4 0 00-1.2 7.8M7 17a4.5 4.5 0 004.5 3h1a4 4 0 001.4-7.7"/>',
    alignL: '<path d="M3 6h18M3 12h11M3 18h15"/>',
    alignC: '<path d="M3 6h18M6 12h12M5 18h14"/>',
    alignR: '<path d="M3 6h18M10 12h11M6 18h15"/>',
    alignJ: '<path d="M3 6h18M3 12h18M3 18h18"/>',
    ul: '<circle cx="4" cy="6" r="1.4"/><circle cx="4" cy="12" r="1.4"/><circle cx="4" cy="18" r="1.4"/><path d="M9 6h12M9 12h12M9 18h12"/>',
    ol: '<path d="M4 5h1v4M4 9h2M4 14h2l-2 3h2M9 6h12M9 12h12M9 18h12"/>',
    indent: '<path d="M3 6h18M11 12h10M11 18h10M3 18V10l4 4z"/>',
    outdent: '<path d="M3 6h18M11 12h10M11 18h10M7 18V10l-4 4z"/>',
    link: '<path d="M10 13a5 5 0 007.5.5l3-3a5 5 0 00-7-7l-1.7 1.7"/><path d="M14 11a5 5 0 00-7.5-.5l-3 3a5 5 0 007 7l1.7-1.7"/>',
    unlink: '<path d="M15 7l3-3a5 5 0 017 7l-3 3M9 17l-3 3a5 5 0 01-7-7l3-3M3 3l18 18"/>',
    image: '<rect x="3" y="4" width="18" height="16" rx="2"/><circle cx="8.5" cy="9.5" r="1.5"/><path d="M21 16l-5-5L5 20"/>',
    table: '<rect x="3" y="4" width="18" height="16" rx="1.6"/><path d="M3 10h18M3 15h18M9 4v16M15 4v16"/>',
    hr: '<path d="M3 12h18M6 7h12M6 17h12" stroke-dasharray="0 0"/>',
    quote: '<path d="M7 7H4v5h4v-1c0 2-1 3-3 3v2c3 0 5-2 5-5V7zM18 7h-3v5h4v-1c0 2-1 3-3 3v2c3 0 5-2 5-5V7z"/>',
    clear: '<path d="M4 7h12M10 7l-1 13M14 7l-.5 6M3 3l18 18M20 12l-4 8"/>',
    color: '<path d="M5 20h14M7 16L12 4l5 12M8.5 13h7"/>',
    fill: '<path d="M19 11l-8-8-8.5 8.5a2 2 0 000 2.8l5.2 5.2a2 2 0 002.8 0z"/><path d="M2 15h17"/><path d="M20 14s2 2.6 2 4a2 2 0 11-4 0c0-1.4 2-4 2-4z"/>',
    field: '<path d="M8 3H6a2 2 0 00-2 2v4l-2 3 2 3v4a2 2 0 002 2h2M16 3h2a2 2 0 012 2v4l2 3-2 3v4a2 2 0 01-2 2h-2"/>',
    eye: '<path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7-10-7-10-7z"/><circle cx="12" cy="12" r="3"/>',
    design: '<path d="M12 19l7-7 3 3-7 7-3-3zM18 13l-1.5-7.5L2 2l3.5 14.5L13 18l5-5zM2 2l7.6 7.6"/><circle cx="11" cy="11" r="2"/>',
    code: '<path d="M16 18l6-6-6-6M8 6l-6 6 6 6"/>',
    split: '<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M12 4v16"/>',
    full: '<path d="M8 3H5a2 2 0 00-2 2v3M16 3h3a2 2 0 012 2v3M8 21H5a2 2 0 01-2-2v-3M16 21h3a2 2 0 002-2v-3"/>',
    exit: '<path d="M3 8V5a2 2 0 012-2h3M21 8V5a2 2 0 00-2-2h-3M3 16v3a2 2 0 002 2h3M21 16v3a2 2 0 01-2 2h-3"/>',
    tidy: '<path d="M4 6h16M4 12h10M4 18h13"/><path d="M18 15l3 3-3 3" transform="translate(0,-6)"/>',
    broom: '<path d="M19 3l-7 7M13 5l6 6M9 21l-4-4 5-6 5 5-6 5z"/>',
    save: '<path d="M19 21H5a2 2 0 01-2-2V5a2 2 0 012-2h11l5 5v11a2 2 0 01-2 2z"/><path d="M17 21v-8H7v8M7 3v5h8"/>',
    folder: '<path d="M3 7a2 2 0 012-2h4l2 2h8a2 2 0 012 2v8a2 2 0 01-2 2H5a2 2 0 01-2-2z"/>',
    open: '<path d="M12 15V3m0 0L8 7m4-4l4 4M4 17v2a2 2 0 002 2h12a2 2 0 002-2v-2"/>',
    shield: '<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><path d="M9 12l2 2 4-4"/>',
    trash: '<path d="M3 6h18M8 6V4h8v2m2 0v14a2 2 0 01-2 2H8a2 2 0 01-2-2V6"/>',
    left: '<path d="M15 18l-6-6 6-6"/>',
    right: '<path d="M9 18l6-6-6-6"/>',
    x: '<path d="M18 6L6 18M6 6l12 12"/>',
    phone: '<rect x="7" y="2" width="10" height="20" rx="2"/><path d="M11 18h2"/>',
    monitor: '<rect x="2" y="4" width="20" height="13" rx="2"/><path d="M8 21h8M12 17v4"/>',
    warn: '<path d="M12 9v4m0 4h.01M10.3 3.9L1.8 18a2 2 0 001.7 3h17a2 2 0 001.7-3L13.7 3.9a2 2 0 00-3.4 0z"/>',
    info: '<circle cx="12" cy="12" r="10"/><path d="M12 16v-4M12 8h.01"/>',
    check: '<path d="M20 6L9 17l-5-5"/>',
    caret: '<path d="M6 9l6 6 6-6"/>',
  };

  const svg = (d, w) =>
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="' +
    (w || 2) + '" stroke-linecap="round" stroke-linejoin="round">' + d + "</svg>";

  /* --------------------------------------------------------------- yardımcı */
  function esc(s) {
    return String(s).replace(/[&<>"']/g, (c) =>
      ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }
  function el(tag, cls, html) {
    const n = document.createElement(tag);
    if (cls) n.className = cls;
    if (html != null) n.innerHTML = html;
    return n;
  }

  /* ==================================================================
     Yer tutucu kuralları — merge.py ile BİREBİR aynı olmalıdır.
     ================================================================== */
  const PH_RE = /\{[^{}\r\n]{1,200}\}/g;

  function trLower(s) {
    return String(s).replace(/I/g, "ı").replace(/İ/g, "i").toLowerCase();
  }
  function trUpper(s) {
    return String(s).replace(/ı/g, "I").replace(/i/g, "İ").toUpperCase();
  }
  function normName(s) {
    return trLower(String(s || "").trim().replace(/\s+/g, " ")).replace(/ı/g, "i");
  }

  // merge.py: FORMATS + _FORMAT_ESLESME
  const FORMAT_KEYS = new Set([
    "buyuk", "kucuk", "baslik", "para", "sayi", "tamsayi", "tarih", "tarihsaat",
    "saat", "gun", "ay", "yil", "kirp", "tekhane", "ham",
    "upper", "lower", "title", "money", "currency", "number", "int",
    "date", "datetime", "time", "day", "month", "year", "trim", "oneline", "raw", "html",
  ]);

  function knownFormat(fmt) {
    const f = (fmt || "").trim();
    if (!f) return "";
    if (f.charAt(0) === "%") return f;
    const k = normName(f);
    return FORMAT_KEYS.has(k) ? k : null;
  }

  /** '{Tutar:para|0,00}' -> {name, fmt, def, key} ; ayrıştırılamazsa null */
  function parsePh(raw) {
    const inner = raw.slice(1, -1);
    const bar = inner.indexOf("|");
    const spec = bar < 0 ? inner : inner.slice(0, bar);
    const def = bar < 0 ? null : inner.slice(bar + 1);
    const colon = spec.indexOf(":");
    const name = (colon < 0 ? spec : spec.slice(0, colon)).trim();
    const fmt = colon < 0 ? "" : spec.slice(colon + 1).trim();
    if (!name) return null;
    return { name: name, fmt: fmt, def: def, key: normName(name), raw: raw };
  }

  /** merge.py:_looks_like_field — CSS parçalarını alan sanmamak için katı. */
  function looksLikeField(p) {
    if (!p) return false;
    if (p.name.indexOf(";") >= 0 || (p.fmt || "").indexOf(";") >= 0) return false;
    if (p.fmt && knownFormat(p.fmt) === null) return false;
    if (p.name.length > 64 || !/[a-zA-ZğĞüÜşŞıİöÖçÇ]/.test(p.name)) return false;
    if ("@#.!$&*/\\".indexOf(p.name.charAt(0)) >= 0) return false;
    if (/[<>=;{}()"'`]/.test(p.name)) return false;
    return true;
  }

  /* ================================================================== */
  const st = {
    mode: "design",      // design | code | split
    html: true,          // HTML mi, düz metin mi
    value: "",           // KAYNAK: tek doğruluk merkezi
    fullDoc: false,      // <html>/<body> içeren tam belge mi
    doctype: "",
    designLive: false,   // iframe güncel mi
    designDirty: false,  // tasarımda düzenleme yapıldı mı
    full: false,
    wrap: false,
    fields: { columns: [], builtins: [], formats: [] },
    keys: new Set(),     // bilinen alan anahtarları
    sample: null,        // {ad: "değer"} ilk satır
    api: null,
    getJob: null,
    onChange: null,
    onToast: null,
  };

  let root, barMain, barSecond, bodyBox, designPane, codePane, frame,
      srcArea, hlBox, gutter, codeBox, statusBar, backdrop, previewBox;
  let savedRange = null;
  let phTimer = null, hlTimer = null, changeTimer = null;

  const toast = (m, t) => { if (st.onToast) st.onToast(m, t || "info"); };

  /* ================================================================
     iframe (tasarım görünümü)
     ================================================================ */
  const EDITOR_CSS =
    "html{box-sizing:border-box}" +
    "body{margin:0;padding:14px 16px;font-family:Arial,Helvetica,sans-serif;font-size:14px;" +
    "line-height:1.6;color:#111;min-height:100%;outline:none;word-wrap:break-word}" +
    "body:empty:before{content:attr(data-bos);color:#aeaeb2}" +
    ".ed-ph{background:rgba(0,122,255,.14);color:#0056b3;border-radius:4px;padding:0 3px;" +
    "font-weight:600;white-space:nowrap;box-shadow:inset 0 0 0 1px rgba(0,122,255,.28)}" +
    ".ed-ph.bad{background:rgba(192,22,22,.12);color:#c01616;box-shadow:inset 0 0 0 1px rgba(192,22,22,.3)}" +
    "*[contenteditable=false]{user-select:all}";
  //
  // DİKKAT — buraya 'img{max-width:100%}' ya da 'table{border-collapse:collapse}'
  // GİBİ KURALLAR EKLEMEYİN. Chromium, insertHTML ile eklenen satır içi stilleri
  // "zaten geçerli" sayıp SİLER: editöre böyle bir kural koyduğumuzda eklediğimiz
  // <img style="max-width:100%"> ve <table style="border-collapse:collapse">
  // stilleri sessizce kayboluyordu ve mail, resmi taşmış / tablo çizgileri çift
  // olarak gidiyordu. Ayrıca bu kurallar önizlemeyi de yalancı yapar: burada
  // düzgün görünen şey e-posta istemcisinde bozuk çıkardı.
  //

  function doc() { return frame && frame.contentDocument; }

  function isFullDoc(s) {
    return /<html[\s>]/i.test(s) || /<body[\s>]/i.test(s) || /<!doctype/i.test(s);
  }

  /** Kaynağı iframe'e yazar. Undo geçmişi sıfırlanır; bu yüzden gereksiz çağrılmaz. */
  function writeFrame() {
    const d = doc();
    if (!d) return;
    let src = st.value || "";
    st.fullDoc = isFullDoc(src);
    const dm = /^\s*(<!doctype[^>]*>)/i.exec(src);
    st.doctype = dm ? dm[1] : "";

    if (!st.fullDoc) {
      src = '<!doctype html><html><head><meta charset="utf-8"></head><body>' + src + "</body></html>";
    } else if (!st.doctype) {
      src = "<!doctype html>" + src;
    }

    d.open();
    d.write(src);
    d.close();

    // Editöre özel stil — dışa aktarırken silinir.
    const style = d.createElement("style");
    style.setAttribute("data-ed-only", "1");
    style.textContent = EDITOR_CSS;
    (d.head || d.documentElement).appendChild(style);

    if (!d.body) d.documentElement.appendChild(d.createElement("body"));
    d.body.contentEditable = "true";
    d.body.spellcheck = false;
    d.body.setAttribute("data-bos", "Mail içeriğinizi buraya yazın…");

    decorateChips();
    bindFrame(d);
    st.designLive = true;
    st.designDirty = false;
  }

  function bindFrame(d) {
    d.addEventListener("input", onDesignInput);
    d.addEventListener("keydown", onFrameKey);
    d.addEventListener("paste", onPaste);
    d.addEventListener("drop", onDrop);
    d.addEventListener("dragover", (e) => e.preventDefault());
    d.addEventListener("selectionchange", () => { rememberSel(); syncToolbar(); });
    d.addEventListener("mouseup", rememberSel);
    d.addEventListener("focus", () => root.classList.add("focus"), true);
    d.addEventListener("blur", () => root.classList.remove("focus"), true);
    // İçerideki bağlantılar tıklanınca gezinmesin.
    d.addEventListener("click", (e) => {
      const a = e.target && e.target.closest && e.target.closest("a");
      if (a) e.preventDefault();
    });
  }

  function onDesignInput() {
    st.designDirty = true;
    scheduleChange();
    clearTimeout(phTimer);
    phTimer = setTimeout(() => decorateChips(true), 550);
  }

  /** iframe'den kaynak metni üretir. Rozetler ve editör stilleri temizlenir. */
  function serializeFrame() {
    const d = doc();
    if (!d || !d.documentElement) return st.value;
    const root2 = d.documentElement.cloneNode(true);

    root2.querySelectorAll("[data-ed-only]").forEach((n) => n.remove());
    root2.querySelectorAll("span.ed-ph").forEach((n) => {
      n.replaceWith(d.createTextNode(n.getAttribute("data-ph") || n.textContent));
    });
    root2.querySelectorAll("[contenteditable]").forEach((n) => n.removeAttribute("contenteditable"));
    root2.querySelectorAll("[spellcheck]").forEach((n) => n.removeAttribute("spellcheck"));
    const b = root2.querySelector("body");
    if (b) b.removeAttribute("data-bos");

    if (st.fullDoc) return (st.doctype ? st.doctype + "\n" : "") + root2.outerHTML;
    return b ? b.innerHTML : root2.innerHTML;
  }

  /* ---------------------------------------------------- yer tutucu rozetleri */
  function phClass(raw) {
    const p = parsePh(raw);
    if (!p) return "";
    if (st.keys.has(p.key)) return "ok";
    return looksLikeField(p) ? "bad" : "";
  }

  /**
   * Metin düğümlerindeki {Alan} yazımlarını rozete çevirir.
   * @param {boolean} keepCaret  true ise imlecin bulunduğu düğüme dokunulmaz
   *        (yazarken imleç kaybolmasın diye). O düğüm, imleç başka yere gidince
   *        bir sonraki turda rozetlenir.
   */
  function decorateChips(keepCaret) {
    const d = doc();
    if (!d || !d.body) return;

    let caretNode = null;
    if (keepCaret) {
      const sel = d.getSelection();
      if (sel && sel.anchorNode) caretNode = sel.anchorNode;
    }

    const targets = [];
    const walker = d.createTreeWalker(d.body, NodeFilter.SHOW_TEXT, {
      acceptNode(n) {
        if (!n.nodeValue || n.nodeValue.indexOf("{") < 0) return NodeFilter.FILTER_REJECT;
        if (n === caretNode) return NodeFilter.FILTER_REJECT;
        let p = n.parentNode;
        while (p && p !== d.body) {
          const t = p.nodeName;
          if (t === "STYLE" || t === "SCRIPT" || t === "TEXTAREA") return NodeFilter.FILTER_REJECT;
          if (p.classList && p.classList.contains("ed-ph")) return NodeFilter.FILTER_REJECT;
          p = p.parentNode;
        }
        return NodeFilter.FILTER_ACCEPT;
      },
    });
    let n;
    while ((n = walker.nextNode())) targets.push(n);

    targets.forEach((node) => {
      const text = node.nodeValue;
      PH_RE.lastIndex = 0;
      let m, last = 0, frag = null;
      while ((m = PH_RE.exec(text))) {
        const cls = phClass(m[0]);
        if (!cls) continue;
        frag = frag || d.createDocumentFragment();
        if (m.index > last) frag.appendChild(d.createTextNode(text.slice(last, m.index)));
        const chip = d.createElement("span");
        chip.className = "ed-ph" + (cls === "bad" ? " bad" : "");
        chip.setAttribute("contenteditable", "false");
        chip.setAttribute("data-ph", m[0]);
        chip.title = cls === "bad"
          ? "Bu adda bir sütun yok — metinde olduğu gibi kalır"
          : (sampleFor(m[0]) || "Excel sütunu");
        chip.textContent = m[0];          // ROZETİN METNİ = YER TUTUCUNUN KENDİSİ
        frag.appendChild(chip);
        last = m.index + m[0].length;
      }
      if (frag) {
        if (last < text.length) frag.appendChild(d.createTextNode(text.slice(last)));
        node.parentNode.replaceChild(frag, node);
      }
    });
  }

  /** Rozet ipucu: 'Ad Soyad → Ahmet Yılmaz' */
  function sampleFor(raw) {
    const p = parsePh(raw);
    if (!p || !st.sample) return "";
    const v = st.sample[p.key];
    if (v == null || v === "") return p.name + " → (bu satırda boş)";
    return p.name + " → " + v;
  }

  /* ================================================================
     Kaynak (kod) görünümü
     ================================================================ */
  function findTagEnd(s, i) {
    let q = 0;
    for (let j = i + 1; j < s.length; j++) {
      const c = s[j];
      if (q) { if (c === q) q = 0; continue; }
      if (c === '"' || c === "'") { q = c; continue; }
      if (c === ">") return j + 1;
    }
    return s.length;
  }

  function markText(s) {
    let out = "", i = 0, m;
    PH_RE.lastIndex = 0;
    while ((m = PH_RE.exec(s))) {
      out += esc(s.slice(i, m.index));
      const c = phClass(m[0]);
      out += c ? '<span class="' + (c === "bad" ? "t-ph-bad" : "t-ph") + '">' + esc(m[0]) + "</span>"
               : esc(m[0]);
      i = m.index + m[0].length;
    }
    return out + esc(s.slice(i));
  }

  function markAttrs(s) {
    let out = "", i = 0, m;
    const re = /([a-zA-Z_:@][\w:.-]*)(\s*=\s*)("[^"]*"|'[^']*'|[^\s"'>]+)?/g;
    while ((m = re.exec(s))) {
      out += esc(s.slice(i, m.index));
      out += '<span class="t-attr">' + esc(m[1]) + "</span>";
      if (m[2]) out += esc(m[2]);
      if (m[3] != null) out += '<span class="t-val">' + markText(m[3]) + "</span>";
      i = m.index + m[0].length;
    }
    return out + esc(s.slice(i));
  }

  function markTag(t) {
    const m = /^<\/?\s*([a-zA-Z][\w:-]*)/.exec(t);
    if (!m) return esc(t);
    const head = t.slice(0, m[0].length - m[1].length);
    return '<span class="t-tag">' + esc(head) + "</span>" +
           '<span class="t-name">' + esc(m[1]) + "</span>" +
           markAttrs(t.slice(m[0].length));
  }

  function highlight(src) {
    let out = "", i = 0;
    const n = src.length;
    while (i < n) {
      const lt = src.indexOf("<", i);
      if (lt < 0) { out += markText(src.slice(i)); break; }
      if (lt > i) out += markText(src.slice(i, lt));

      if (src.startsWith("<!--", lt)) {
        let e = src.indexOf("-->", lt + 4);
        e = e < 0 ? n : e + 3;
        out += '<span class="t-cmt">' + esc(src.slice(lt, e)) + "</span>";
        i = e; continue;
      }
      if (src.startsWith("<!", lt) || src.startsWith("<?", lt)) {
        let e = src.indexOf(">", lt);
        e = e < 0 ? n : e + 1;
        out += '<span class="t-doc">' + esc(src.slice(lt, e)) + "</span>";
        i = e; continue;
      }
      const m = /^<\/?\s*([a-zA-Z][\w:-]*)/.exec(src.slice(lt, lt + 60));
      if (!m) { out += esc("<"); i = lt + 1; continue; }

      const e = findTagEnd(src, lt);
      out += markTag(src.slice(lt, e));
      i = e;

      const name = m[1].toLowerCase();
      if ((name === "style" || name === "script") && src[lt + 1] !== "/") {
        const close = new RegExp("</\\s*" + name + "\\s*>", "i");
        const rest = src.slice(i);
        const cm = close.exec(rest);
        const end = cm ? i + cm.index : n;
        if (end > i) out += '<span class="t-css">' + esc(src.slice(i, end)) + "</span>";
        i = end;
      }
    }
    return out;
  }

  const HL_LIMIT = 150000;   // bu boyutun üstünde vurgulama kapanır (akıcılık)

  function refreshCode(force) {
    const src = srcArea.value;
    const big = src.length > HL_LIMIT;
    codeBox.classList.toggle("plain", big || !st.html);
    if (!big && st.html) hlBox.innerHTML = highlight(src) + "\n";
    // Satır numaraları
    if (!st.wrap) {
      const lines = src.split("\n").length;
      if (force || gutter._n !== lines) {
        let g = "";
        for (let k = 1; k <= lines; k++) g += k + "\n";
        gutter.textContent = g;
        gutter._n = lines;
      }
    }
    syncScroll();
  }

  function syncScroll() {
    hlBox.scrollTop = srcArea.scrollTop;
    hlBox.scrollLeft = srcArea.scrollLeft;
    gutter.scrollTop = srcArea.scrollTop;
  }

  function onCodeInput() {
    clearTimeout(hlTimer);
    hlTimer = setTimeout(() => refreshCode(false), 60);
    scheduleChange();
    if (st.mode === "split") {
      clearTimeout(onCodeInput._t);
      onCodeInput._t = setTimeout(() => {
        st.value = srcArea.value;
        writeFrame();
      }, 700);
    } else {
      st.designLive = false;
    }
  }

  function onCodeKey(e) {
    if (e.key === "Tab") {
      e.preventDefault();
      const a = srcArea.selectionStart, b = srcArea.selectionEnd;
      const v = srcArea.value;
      if (a === b && !e.shiftKey) {
        srcArea.setRangeText("  ", a, b, "end");
      } else {
        // Satır satır girinti ekle/çıkar
        let ls = v.lastIndexOf("\n", a - 1) + 1;
        const le = v.indexOf("\n", b) < 0 ? v.length : v.indexOf("\n", b);
        const block = v.slice(ls, le);
        const yeni = e.shiftKey
          ? block.replace(/^ {1,2}/gm, "")
          : block.replace(/^/gm, "  ");
        srcArea.setRangeText(yeni, ls, le, "select");
      }
      onCodeInput();
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "s") {
      e.preventDefault(); openTemplates(true);
    }
  }

  /* --------------------------------------------------------- güzelleştirme */
  const BLOCK = new Set(("html,head,body,div,p,table,thead,tbody,tfoot,tr,td,th,ul,ol,li," +
    "h1,h2,h3,h4,h5,h6,section,article,header,footer,nav,aside,blockquote,pre,form," +
    "fieldset,hr,style,script,meta,link,title,center,figure,figcaption,dl,dt,dd,main").split(","));
  const VOID = new Set(("area,base,br,col,embed,hr,img,input,link,meta,param,source,track,wbr").split(","));
  const RAW = new Set(["pre", "style", "script", "textarea"]);

  function tokenize(src) {
    const out = [];
    let i = 0;
    const n = src.length;
    while (i < n) {
      const lt = src.indexOf("<", i);
      if (lt < 0) { out.push({ t: "text", s: src.slice(i) }); break; }
      if (lt > i) out.push({ t: "text", s: src.slice(i, lt) });
      if (src.startsWith("<!--", lt)) {
        let e = src.indexOf("-->", lt + 4); e = e < 0 ? n : e + 3;
        out.push({ t: "cmt", s: src.slice(lt, e) }); i = e; continue;
      }
      if (src.startsWith("<!", lt) || src.startsWith("<?", lt)) {
        let e = src.indexOf(">", lt); e = e < 0 ? n : e + 1;
        out.push({ t: "decl", s: src.slice(lt, e) }); i = e; continue;
      }
      const m = /^<(\/?)\s*([a-zA-Z][\w:-]*)/.exec(src.slice(lt, lt + 60));
      if (!m) { out.push({ t: "text", s: "<" }); i = lt + 1; continue; }
      const e = findTagEnd(src, lt);
      const name = m[2].toLowerCase();
      const kapali = !!m[1];
      out.push({
        t: "tag", s: src.slice(lt, e), name: name, close: kapali,
        self: /\/>$/.test(src.slice(lt, e)) || VOID.has(name),
      });
      i = e;
      if (!kapali && RAW.has(name) && !VOID.has(name)) {
        const close = new RegExp("</\\s*" + name + "\\s*>", "i");
        const cm = close.exec(src.slice(i));
        const end = cm ? i + cm.index : n;
        if (end > i) out.push({ t: "raw", s: src.slice(i, end) });
        i = end;
      }
    }
    return out;
  }

  /**
   * HTML'i okunur hâle getirir.
   * Yalnızca BLOK etiketlerinin önüne/arkasına satır sonu koyar; satır içi
   * (inline) diziliminde boşluk ANLAMLIDIR, bu yüzden oralara dokunulmaz.
   */
  function beautify(src) {
    const tk = tokenize(src);
    const lines = [];
    let depth = 0;
    let cur = "";          // yazılmakta olan satır (girintisi dahil)
    let acildi = false;    // az önce bir blok etiketi açıldı mı

    const ind = () => "  ".repeat(Math.max(0, depth));
    const flush = () => { if (cur.trim()) lines.push(cur.replace(/\s+$/, "")); cur = ""; };
    const isBlock = (t) => t && t.t === "tag" && BLOCK.has(t.name);

    for (let i = 0; i < tk.length; i++) {
      const t = tk[i];

      if (t.t === "text") {
        if (!t.s.trim()) {
          // Yalnızca boşluk: blok sınırında görüntülenmez, at.
          // Satır içi öğeler arasındaysa TEK boşluğa indir (aynı şeyi çizer).
          const prev = tk[i - 1], next = tk[i + 1];
          if (isBlock(prev) || isBlock(next) || !prev || !next) continue;
          if (cur.trim()) cur += " ";
          continue;
        }
        let s = t.s.replace(/\s*\n\s*/g, " ");
        // Bir bloğun hemen başındaki boşluk HTML'de zaten görüntülenmez; atmak
        // "Biçimlendir"i KARARLI kılar (ikinci kez basmak belgeyi değiştirmez).
        if (acildi || !cur.trim()) s = s.replace(/^[ \t]+/, "");
        if (!cur) cur = ind();
        cur += s;
        acildi = false;
        continue;
      }

      if (t.t === "raw") { cur += t.s; acildi = false; continue; }

      if (t.t === "cmt" || t.t === "decl") {
        flush(); lines.push(ind() + t.s); acildi = false; continue;
      }

      // <style>/<script>/<pre>/<textarea>: açılış + içerik + kapanış TEK parça
      // olarak yazılır; içeriğine kesinlikle dokunulmaz.
      if (!t.close && RAW.has(t.name) && !t.self) {
        let s = t.s;
        if (tk[i + 1] && tk[i + 1].t === "raw") { s += tk[i + 1].s; i++; }
        if (tk[i + 1] && tk[i + 1].t === "tag" && tk[i + 1].close && tk[i + 1].name === t.name) {
          s += tk[i + 1].s; i++;
        }
        flush(); lines.push(ind() + s); acildi = false;
        continue;
      }

      // Satır içi etiket: satırı bölmeyiz, boşluk anlamlıdır.
      if (!BLOCK.has(t.name)) {
        if (!cur) cur = ind();
        cur += t.s;
        acildi = false;
        continue;
      }

      if (t.close) {
        depth--;
        cur = cur.replace(/[ \t]+$/, "");
        if (cur.trim()) { cur += t.s; flush(); }     // <p>metin</p> tek satırda
        else { flush(); lines.push(ind() + t.s); }   // </div> kendi satırında
        acildi = false;
        continue;
      }

      // Blok açılışı: yeni satır BAŞLATIR ama satırı bitirmez —
      // kısa içerik aynı satırda kalsın diye ( <p>Merhaba</p> ).
      flush();
      cur = ind() + t.s;
      if (!t.self) { depth++; acildi = true; }
    }
    flush();
    return lines.join("\n").replace(/\n{3,}/g, "\n\n");
  }

  /* ------------------------------------------------------------ Word temizle */
  function cleanWord(src) {
    let s = src;
    s = s.replace(/<!--\[if[\s\S]*?<!\[endif\]-->/gi, "");
    s = s.replace(/<!\[if[\s\S]*?<!\[endif\]>/gi, "");
    s = s.replace(/<\/?[ovwxmp]:[^>]*>/gi, "");
    s = s.replace(/<!--[\s\S]*?-->/g, (m) =>
      /mso|WordSection|Generator|StartFragment|EndFragment/i.test(m) ? "" : m);
    s = s.replace(/<meta[^>]*(Generator|ProgId|Originator)[^>]*>/gi, "");
    s = s.replace(/<link[^>]*File-List[^>]*>/gi, "");
    // class=MsoNormal  /  class="MsoListParagraphCxSpFirst"
    s = s.replace(/\sclass\s*=\s*(?:"[^"]*Mso[^"]*"|'[^']*Mso[^']*'|Mso[^\s>]*)/gi, "");
    // style içindeki mso-* bildirimleri
    s = s.replace(/\sstyle\s*=\s*(["'])([^"']*)\1/gi, (m, q, css) => {
      const kalan = css.split(";")
        .filter((d) => d.trim() && !/^\s*mso-/i.test(d) && !/^\s*tab-stops/i.test(d))
        .join(";").trim();
      return kalan ? " style=" + q + kalan + q : "";
    });
    s = s.replace(/<span\s*>([\s\S]*?)<\/span>/gi, "$1");
    s = s.replace(/<a\s+name\s*=\s*(["'])?_[^>]*>\s*<\/a>/gi, "");
    s = s.replace(/\sv:shapes\s*=\s*(["'])[^"']*\1/gi, "");
    s = s.replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n");
    return s;
  }

  /* ------------------------------------------------------- yapıştırma / bırakma */
  function sanitize(html) {
    let s = html;
    s = s.replace(/<script[\s\S]*?<\/script>/gi, "");
    s = s.replace(/\son[a-z]+\s*=\s*(["'][^"']*["']|[^\s>]+)/gi, "");
    s = s.replace(/(href|src)\s*=\s*(["'])\s*javascript:[^"']*\2/gi, '$1="#"');
    if (/mso-|MsoNormal|urn:schemas-microsoft-com|WordSection/i.test(s)) s = cleanWord(s);
    return s;
  }

  function onPaste(e) {
    const cd = e.clipboardData;
    if (!cd) return;
    // Panodaki resim dosyası (ekran görüntüsü) -> gömülü resim
    const files = cd.files;
    if (files && files.length && /^image\//.test(files[0].type)) {
      e.preventDefault();
      readImage(files[0]);
      return;
    }
    const html = cd.getData("text/html");
    if (html && st.html) {
      e.preventDefault();
      if (e.shiftKey) {
        exec("insertText", cd.getData("text/plain"));
      } else {
        exec("insertHTML", sanitize(html));
        setTimeout(() => decorateChips(true), 0);
      }
    }
  }

  function onDrop(e) {
    const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    if (f && /^image\//.test(f.type)) { e.preventDefault(); readImage(f); }
  }

  function readImage(file) {
    if (file.size > 3 * 1024 * 1024) {
      toast("Resim çok büyük (" + Math.round(file.size / 1024) + " KB). 3 MB altını tercih edin.", "error");
      return;
    }
    const fr = new FileReader();
    fr.onload = () => insertImage(fr.result, file.name);
    fr.readAsDataURL(file);
  }

  function insertImage(uri, name) {
    const html = '<img src="' + esc(uri) + '" alt="' + esc(name || "") +
                 '" style="max-width:100%;height:auto" />';
    insertHTML(html);
    toast("Resim gömüldü. Gönderimde otomatik olarak mail içine yerleştirilir.", "success");
  }

  /* ================================================================
     Komutlar
     ================================================================ */
  function focusFrame() {
    if (frame && frame.contentWindow) frame.contentWindow.focus();
  }

  function rememberSel() {
    const d = doc();
    if (!d) return;
    const sel = d.getSelection();
    if (sel && sel.rangeCount && d.body && d.body.contains(sel.anchorNode)) {
      savedRange = sel.getRangeAt(0).cloneRange();
    }
  }

  function restoreSel() {
    const d = doc();
    if (!d) return;
    focusFrame();
    if (!savedRange) {
      if (d.body) {
        const r = d.createRange();
        r.selectNodeContents(d.body);
        r.collapse(false);
        const s = d.getSelection();
        s.removeAllRanges(); s.addRange(r);
      }
      return;
    }
    const s = d.getSelection();
    s.removeAllRanges();
    s.addRange(savedRange);
  }

  function exec(cmd, val) {
    const d = doc();
    if (!d) return false;
    restoreSel();
    try { d.execCommand("styleWithCSS", false, true); } catch (e) {}
    let ok = false;
    try { ok = d.execCommand(cmd, false, val); } catch (e) { ok = false; }
    st.designDirty = true;
    rememberSel();
    scheduleChange();
    syncToolbar();
    return ok;
  }

  /** Yazı boyutu: execCommand yalnızca 1-7 kabul eder; <font size=7> üretip
      onu px'li <span>'e çeviriyoruz (tarayıcıda çalışan tek güvenilir yol). */
  function setFontSize(px) {
    const d = doc();
    if (!d) return;
    restoreSel();
    try { d.execCommand("styleWithCSS", false, false); } catch (e) {}
    d.execCommand("fontSize", false, "7");
    d.querySelectorAll('font[size="7"]').forEach((f) => {
      const s = d.createElement("span");
      s.style.fontSize = px + "px";
      while (f.firstChild) s.appendChild(f.firstChild);
      f.replaceWith(s);
    });
    try { d.execCommand("styleWithCSS", false, true); } catch (e) {}
    st.designDirty = true;
    scheduleChange();
  }

  function insertHTML(html) {
    if (st.mode === "code" || (!st.html && st.mode !== "design")) {
      insertAtCode(html);
      return;
    }
    exec("insertHTML", html);
    setTimeout(() => decorateChips(true), 0);
  }

  function insertAtCode(text) {
    const a = srcArea.selectionStart, b = srcArea.selectionEnd;
    srcArea.setRangeText(text, a, b, "end");
    srcArea.focus();
    onCodeInput();
  }

  /** Yer tutucuyu bulunulan görünüme uygun biçimde ekler. */
  function insertField(raw) {
    if (st.mode === "code" || !st.html) { insertAtCode(raw); return; }
    exec("insertHTML", esc(raw) + "&nbsp;");
    setTimeout(() => decorateChips(false), 0);
  }

  /* ================================================================
     Araç çubuğu durumu
     ================================================================ */
  const stateBtns = {};
  function syncToolbar() {
    const d = doc();
    if (!d || st.mode === "code" || !st.html) return;
    Object.keys(stateBtns).forEach((cmd) => {
      let on = false;
      try { on = d.queryCommandState(cmd); } catch (e) {}
      stateBtns[cmd].classList.toggle("on", !!on);
    });
    try {
      const blk = (d.queryCommandValue("formatBlock") || "").toLowerCase();
      const sel = root.querySelector(".w-block");
      if (sel) sel.value = ["h1", "h2", "h3", "blockquote", "pre"].indexOf(blk) >= 0 ? blk : "p";
    } catch (e) {}
  }

  /* ================================================================
     Değişiklik bildirimi
     ================================================================ */
  function scheduleChange() {
    clearTimeout(changeTimer);
    changeTimer = setTimeout(() => {
      updateStatus();
      if (st.onChange) st.onChange();
    }, 250);
  }

  /** Kaynağı, aktif görünümden toplayıp döndürür. */
  function value() {
    if (st.mode === "code") { st.value = srcArea.value; return st.value; }
    if (st.mode === "split") {
      // Bölünmüş görünümde kaynak kutusu asıldır (tasarım ondan beslenir).
      if (st.designDirty) { st.value = serializeFrame(); srcArea.value = st.value; refreshCode(true); }
      else st.value = srcArea.value;
      return st.value;
    }
    if (st.designDirty) st.value = serializeFrame();
    return st.value;
  }

  /* ================================================================
     Görünüm değiştirme
     ================================================================ */
  function setMode(m) {
    if (m === st.mode) return;
    // Ayrılmadan önce mevcut görünümdeki değişiklikleri kaynağa al
    if (st.mode === "design" && st.designDirty) st.value = serializeFrame();
    else if (st.mode === "code" || st.mode === "split") st.value = srcArea.value;

    st.mode = m;
    const dsn = m === "design" || m === "split";
    const cod = m === "code" || m === "split";
    designPane.classList.toggle("hidden", !dsn);
    codePane.classList.toggle("hidden", !cod);
    bodyBox.classList.toggle("split", m === "split");
    root.querySelectorAll(".ed-modes button").forEach((b) =>
      b.classList.toggle("on", b.dataset.mode === m));

    if (cod) { srcArea.value = st.value; refreshCode(true); }
    if (dsn) { writeFrame(); }
    // HTML'e özgü düğmeler yalnızca tasarımda anlamlı
    updateBarState();
    updateStatus();
  }

  function updateBarState() {
    const designAktif = (st.mode === "design" || st.mode === "split") && st.html;
    barMain.querySelectorAll("[data-need-design]").forEach((b) => { b.disabled = !designAktif; });
    barMain.classList.toggle("hidden", !st.html);
  }

  function setFull(on) {
    st.full = on;
    root.classList.toggle("full", on);
    backdrop.classList.toggle("hidden", !on);
    const b = root.querySelector('[data-act="full"]');
    if (b) {
      b.innerHTML = svg(on ? I.exit : I.full) + (on ? "Küçült" : "Tam ekran");
      b.title = on ? "Tam ekrandan çık (Esc)" : "Tam ekran (Ctrl+Shift+F)";
    }
  }

  /* ================================================================
     Açılır paneller
     ================================================================ */
  let openPop = null;
  function closePop() {
    if (openPop) {
      if (openPop._ro) openPop._ro.disconnect();
      openPop.remove();
      openPop = null;
    }
  }

  /** Paneli ankrajın altına (sığmazsa üstüne) ve HER ZAMAN ekranın içine koyar. */
  function yerlestir(p, anchor) {
    const ar = anchor.getBoundingClientRect();
    const VW = document.documentElement.clientWidth;
    const VH = document.documentElement.clientHeight;

    // Çok uzun panel ekrandan taşmasın: kendi içinde kaysın.
    p.style.maxHeight = (VH - 16) + "px";
    p.style.overflowY = "auto";
    p.style.left = Math.max(8, Math.min(ar.left, VW - p.offsetWidth - 8)) + "px";

    // Ankraj görünür alanın dışında bile olabilir; bu yüzden HER iki aday da
    // ayrı ayrı sınanır — tek bir "sığar mı" kontrolü yetmez.
    const h = p.offsetHeight;
    const sigar = (t) => t >= 8 && t + h <= VH - 8;
    let top = ar.bottom + 4;
    if (!sigar(top)) {
      const ust = ar.top - h - 4;
      top = sigar(ust) ? ust : Math.max(8, VH - h - 8);
    }
    p.style.top = top + "px";
  }
  function popup(anchor, builder, opts) {
    const acikAyni = openPop && openPop._src === anchor;
    closePop();
    if (acikAyni) return null;
    const p = el("div", "ed-pop");
    p._src = anchor;
    builder(p);
    // <body>'ye eklenir: editörün overflow:hidden kutusu paneli kırpamaz.
    document.body.appendChild(p);

    yerlestir(p, anchor);
    // Şablon listesi gibi içeriği SONRADAN gelen paneller büyüyünce ekrandan
    // taşabilir; boyut değişimini izleyip yeniden yerleştiriyoruz.
    if (window.ResizeObserver) {
      const ro = new ResizeObserver(() => yerlestir(p, anchor));
      ro.observe(p);
      p._ro = ro;
    }
    openPop = p;
    if (opts && opts.focus) {
      const f = p.querySelector("input,select");
      if (f) setTimeout(() => f.focus(), 0);
    }
    return p;
  }

  document.addEventListener("mousedown", (e) => {
    if (openPop && !openPop.contains(e.target) && !e.target.closest(".ed-btn")) closePop();
  });

  /* -------------------------------------------------------- alan menüsü */
  function openFields(anchor, target) {
    rememberSel();
    popup(anchor, (p) => {
      p.style.minWidth = "290px";
      p.appendChild(el("div", "ed-pop-title", "Excel sütunu ekle"));
      const search = el("input", "ed-pop-search");
      search.type = "text";
      search.placeholder = "Sütun ara…";
      p.appendChild(search);
      const list = el("div", "ed-pop-list");
      p.appendChild(list);

      const fmtSec = el("div", "ed-pop-row");
      fmtSec.innerHTML = '<label>Biçim</label>';
      const fmtSel = el("select");
      (st.fields.formats.length ? st.fields.formats : [["", "olduğu gibi"]]).forEach(([k, d]) => {
        const o = document.createElement("option");
        o.value = k; o.textContent = d;
        fmtSel.appendChild(o);
      });
      fmtSec.appendChild(fmtSel);
      p.appendChild(fmtSec);
      p.appendChild(el("div", "ed-pop-note",
        "Alan, gönderimde o satırın değeriyle değişir. Biçim seçerseniz " +
        "<b>{Tutar:para}</b> gibi yazılır. Değer boşsa yazılacak metni " +
        "<b>{Ad|Sayın Müşterimiz}</b> şeklinde ekleyebilirsiniz."));

      function ekle(ad) {
        const f = fmtSel.value;
        const raw = "{" + ad + (f ? ":" + f : "") + "}";
        closePop();
        if (target === "subject") {
          const inp = document.getElementById("subject");
          const a = inp.selectionStart != null ? inp.selectionStart : inp.value.length;
          const b = inp.selectionEnd != null ? inp.selectionEnd : a;
          inp.value = inp.value.slice(0, a) + raw + inp.value.slice(b);
          inp.focus();
          inp.setSelectionRange(a + raw.length, a + raw.length);
          inp.dispatchEvent(new Event("input", { bubbles: true }));
        } else {
          insertField(raw);
        }
      }

      function ciz() {
        const q = normName(search.value);
        list.innerHTML = "";
        const cols = st.fields.columns.filter((c) => !q || normName(c).indexOf(q) >= 0);
        if (!st.fields.columns.length) {
          list.appendChild(el("div", "ed-pop-empty",
            "Önce bir Excel dosyası seçip <b>Sütunları Oku</b>'ya basın."));
        } else if (!cols.length) {
          list.appendChild(el("div", "ed-pop-empty", "Eşleşen sütun yok."));
        }
        cols.forEach((c) => {
          const b = el("button", "ed-pop-item");
          b.type = "button";
          const ornek = st.sample ? st.sample[normName(c)] : "";
          b.innerHTML = '<span class="nm">' + esc(c) + "</span>" +
                        (ornek ? '<span class="ex">' + esc(ornek) + "</span>" : "");
          b.onclick = () => ekle(c);
          list.appendChild(b);
        });

        const bi = st.fields.builtins.filter(([ad]) => !q || normName(ad).indexOf(q) >= 0);
        if (bi.length) {
          list.appendChild(el("div", "ed-pop-title", "Hazır alanlar"));
          bi.forEach(([ad, aciklama]) => {
            const b = el("button", "ed-pop-item");
            b.type = "button";
            b.innerHTML = '<span class="nm">' + esc(ad) + "</span>" +
                          '<span class="ex">' + esc(aciklama) + "</span>";
            b.onclick = () => ekle(ad);
            list.appendChild(b);
          });
        }
      }
      search.oninput = ciz;
      search.onkeydown = (e) => {
        if (e.key === "Enter") {
          const ilk = list.querySelector(".ed-pop-item");
          if (ilk) ilk.click();
        } else if (e.key === "Escape") closePop();
      };
      ciz();
    }, { focus: true });
  }

  /* ----------------------------------------------------------- renkler */
  const PALET = [
    "#000000", "#111111", "#2c2c2e", "#48484a", "#636366", "#8e8e93", "#aeaeb2", "#c7c7cc", "#e5e5ea", "#ffffff",
    "#c01616", "#e03131", "#f76707", "#f59f00", "#2f9e44", "#34c759", "#0c8599", "#007AFF", "#5f3dc4", "#c2255c",
    "#ffe3e3", "#fff3bf", "#e6fcf5", "#d0ebff", "#e5dbff", "#ffdeeb", "#f8f9fa", "#f1f3f5", "#dee2e6", "#868e96",
  ];

  function openColor(anchor, cmd) {
    rememberSel();
    popup(anchor, (p) => {
      p.appendChild(el("div", "ed-pop-title", cmd === "foreColor" ? "Yazı rengi" : "Vurgu rengi"));
      const grid = el("div", "ed-swatches");
      if (cmd !== "foreColor") {
        const yok = el("button", "ed-swatch none");
        yok.type = "button";
        yok.title = "Vurguyu kaldır";
        yok.onclick = () => { exec("hiliteColor", "transparent"); closePop(); };
        grid.appendChild(yok);
      }
      PALET.forEach((c) => {
        const b = el("button", "ed-swatch");
        b.type = "button";
        b.style.background = c;
        b.title = c;
        b.onclick = () => { exec(cmd, c); closePop(); };
        grid.appendChild(b);
      });
      p.appendChild(grid);
      const row = el("div", "ed-pop-row");
      row.innerHTML = "<label>Özel</label>";
      const inp = el("input");
      inp.type = "color";
      inp.style.cssText = "height:30px;width:52px;padding:2px;border:1.5px solid var(--gray-200);border-radius:8px;background:#fff;cursor:pointer";
      inp.oninput = () => exec(cmd, inp.value);
      row.appendChild(inp);
      p.appendChild(row);
    });
  }

  /* ------------------------------------------------------------ tablo */
  function openTable(anchor) {
    rememberSel();
    popup(anchor, (p) => {
      const lbl = el("div", "ed-grid-label", "Tablo boyutu seçin");
      p.appendChild(lbl);
      const grid = el("div", "ed-grid");
      const R = 6, C = 8;
      const cells = [];
      for (let r = 0; r < R; r++) {
        for (let c = 0; c < C; c++) {
          const b = el("button", "ed-cell");
          b.type = "button";
          b.dataset.r = r + 1; b.dataset.c = c + 1;
          b.onmouseenter = () => {
            lbl.textContent = (r + 1) + " satır × " + (c + 1) + " sütun";
            cells.forEach((x) => x.classList.toggle("on",
              +x.dataset.r <= r + 1 && +x.dataset.c <= c + 1));
          };
          b.onclick = () => { closePop(); insertTable(r + 1, c + 1); };
          cells.push(b);
          grid.appendChild(b);
        }
      }
      p.appendChild(grid);
      p.appendChild(el("div", "ed-pop-note",
        "E-posta uyumlu tablo eklenir (border-collapse + cellpadding)."));
    });
  }

  function insertTable(rows, cols) {
    let h = '<table cellpadding="8" cellspacing="0" border="0" ' +
      'style="border-collapse:collapse;width:100%;margin:8px 0;font-size:14px">';
    for (let r = 0; r < rows; r++) {
      h += "<tr>";
      for (let c = 0; c < cols; c++) {
        const bas = r === 0;
        h += "<t" + (bas ? "h" : "d") + ' style="border:1px solid #e5e5ea;text-align:left' +
             (bas ? ";background:#f9f9fb;font-weight:600" : "") + '">&nbsp;</t' + (bas ? "h" : "d") + ">";
      }
      h += "</tr>";
    }
    h += "</table><p><br></p>";
    insertHTML(h);
  }

  /* ------------------------------------------------------------- bağlantı */
  function openLink(anchor) {
    rememberSel();
    const d = doc();
    const sel = d && d.getSelection();
    const mevcutMetin = sel ? String(sel) : "";
    let mevcutUrl = "";
    if (sel && sel.anchorNode) {
      const a = sel.anchorNode.parentElement && sel.anchorNode.parentElement.closest("a");
      if (a) mevcutUrl = a.getAttribute("href") || "";
    }
    popup(anchor, (p) => {
      p.style.minWidth = "300px";
      p.appendChild(el("div", "ed-pop-title", "Bağlantı"));
      const r1 = el("div", "ed-pop-row", "<label>Adres</label>");
      const url = el("input"); url.type = "text"; url.placeholder = "https://…"; url.value = mevcutUrl;
      r1.appendChild(url); p.appendChild(r1);
      const r2 = el("div", "ed-pop-row", "<label>Metin</label>");
      const txt = el("input"); txt.type = "text"; txt.placeholder = "Görünecek metin"; txt.value = mevcutMetin;
      r2.appendChild(txt); p.appendChild(r2);

      const act = el("div", "ed-pop-actions");
      const kaldir = el("button", "ed-btn txt", "Kaldır");
      kaldir.type = "button";
      kaldir.onclick = () => { exec("unlink"); closePop(); };
      const ok = el("button", "ed-btn txt on", "Ekle");
      ok.type = "button";
      ok.onclick = () => {
        let u = url.value.trim();
        if (!u) { closePop(); return; }
        if (!/^([a-z][\w+.-]*:|#|\/|\{)/i.test(u)) u = "https://" + u;
        const metin = txt.value.trim();
        closePop();
        if (metin && metin !== mevcutMetin) {
          insertHTML('<a href="' + esc(u) + '" target="_blank" rel="noopener">' + esc(metin) + "</a>");
        } else if (mevcutMetin) {
          exec("createLink", u);
        } else {
          insertHTML('<a href="' + esc(u) + '" target="_blank" rel="noopener">' + esc(u) + "</a>");
        }
      };
      act.appendChild(kaldir); act.appendChild(ok);
      p.appendChild(act);
      url.onkeydown = (e) => { if (e.key === "Enter") ok.click(); if (e.key === "Escape") closePop(); };
      txt.onkeydown = url.onkeydown;
    }, { focus: true });
  }

  /* ----------------------------------------------------------- şablonlar */
  const HAZIR = [
    {
      ad: "Sade fatura maili",
      html:
'<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.6;color:#111">\n' +
'  <p>Sayın {Ad Soyad|Yetkili},</p>\n' +
'  <p>{GONDERIM_AYI} dönemine ait <b>{Fatura No}</b> numaralı faturanız ektedir.</p>\n' +
'  <table cellpadding="8" cellspacing="0" border="0" style="border-collapse:collapse;margin:14px 0">\n' +
'    <tr><td style="border:1px solid #e5e5ea;background:#f9f9fb">Fatura No</td>' +
'<td style="border:1px solid #e5e5ea"><b>{Fatura No}</b></td></tr>\n' +
'    <tr><td style="border:1px solid #e5e5ea;background:#f9f9fb">Tutar</td>' +
'<td style="border:1px solid #e5e5ea"><b>{Tutar:para} TL</b></td></tr>\n' +
'  </table>\n' +
'  <p>Bilgilerinize sunarız.<br />İyi çalışmalar dileriz.</p>\n' +
'</div>',
    },
    {
      ad: "Duyuru / bilgilendirme",
      html:
'<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.6;color:#111;max-width:600px">\n' +
'  <h2 style="margin:0 0 12px;font-size:20px">Başlık buraya</h2>\n' +
'  <p>Merhaba {Ad Soyad|Değerli Müşterimiz},</p>\n' +
'  <p>Mesajınız buraya gelecek.</p>\n' +
'  <p style="margin-top:20px">\n' +
'    <a href="https://" style="background:#007AFF;color:#fff;text-decoration:none;' +
'padding:11px 22px;border-radius:6px;display:inline-block">Ayrıntılar</a>\n' +
'  </p>\n' +
'</div>',
    },
    { ad: "Boş sayfa", html: "<p></p>" },
  ];

  function openTemplates(kaydetOdakli) {
    const anchor = root.querySelector('[data-act="tpl"]');
    rememberSel();
    popup(anchor, (p) => {
      p.style.minWidth = "300px";
      p.appendChild(el("div", "ed-pop-title", "Şablonu kaydet"));
      const r = el("div", "ed-pop-row", "<label>Ad</label>");
      const ad = el("input"); ad.type = "text"; ad.placeholder = "Örn. Fatura maili";
      r.appendChild(ad);
      const kaydet = el("button", "ed-btn txt on", "Kaydet");
      kaydet.type = "button";
      kaydet.onclick = () => {
        const isim = ad.value.trim();
        if (!isim) { ad.focus(); return; }
        if (!st.api) return;
        st.api.sablon_kaydet(isim, value()).then((res) => {
          if (res && res.error) toast(res.error, "error");
          else { toast("Şablon kaydedildi: " + res.name, "success"); closePop(); }
        });
      };
      r.appendChild(kaydet);
      p.appendChild(r);

      p.appendChild(el("div", "ed-pop-title", "Kayıtlı şablonlar"));
      const list = el("div", "ed-pop-list");
      p.appendChild(list);
      list.appendChild(el("div", "ed-pop-empty", "Yükleniyor…"));

      const yukle = (html) => {
        closePop();
        setValue(html);
        if (st.onChange) st.onChange();
        toast("Şablon yüklendi.", "success");
      };

      const ciz = (items) => {
        list.innerHTML = "";
        (items || []).forEach((it) => {
          const row = el("div", "ed-tpl-item");
          const b = el("button", "ed-pop-item");
          b.type = "button";
          b.style.flex = "1";
          const t = new Date(it.time * 1000);
          b.innerHTML = '<span class="nm">' + esc(it.name) + "</span>" +
            '<span class="ex">' + t.toLocaleDateString("tr-TR") + "</span>";
          b.onclick = () => st.api.sablon_yukle(it.name).then((res) => {
            if (res && res.error) toast(res.error, "error");
            else yukle(res.content || "");
          });
          const del = el("button", "del", svg(I.trash));
          del.type = "button";
          del.title = "Sil";
          del.onclick = (e) => {
            e.stopPropagation();
            st.api.sablon_sil(it.name).then(() => { row.remove(); toast("Şablon silindi.", "info"); });
          };
          row.appendChild(b); row.appendChild(del);
          list.appendChild(row);
        });
        list.appendChild(el("div", "ed-pop-title", "Hazır şablonlar"));
        HAZIR.forEach((h) => {
          const b = el("button", "ed-pop-item");
          b.type = "button";
          b.innerHTML = '<span class="nm">' + esc(h.ad) + "</span>";
          b.onclick = () => yukle(h.html);
          list.appendChild(b);
        });
      };

      if (st.api && st.api.sablon_listesi) {
        st.api.sablon_listesi().then((res) => ciz(res && res.items)).catch(() => ciz([]));
      } else ciz([]);

      if (kaydetOdakli) setTimeout(() => ad.focus(), 0);
    }, { focus: !kaydetOdakli });
  }

  /* ------------------------------------------------------------- denetim */
  function lint(html) {
    const out = [];
    const add = (lvl, b, d) => out.push({ lvl: lvl, b: b, d: d });
    if (!html.trim()) return out;

    if (/<script[\s>]/i.test(html))
      add("err", "&lt;script&gt; etiketi", "Hiçbir e-posta istemcisi JavaScript çalıştırmaz; bu blok silinir ya da mail spam'e düşer.");
    if (/<(iframe|video|audio|form|input|button|object|embed)[\s>]/i.test(html))
      add("err", "Desteklenmeyen etiket", "iframe/video/form gibi etiketler e-postada çalışmaz. Bunun yerine bir bağlantı düğmesi kullanın.");
    if (/display\s*:\s*(flex|grid)/i.test(html))
      add("warn", "flex / grid düzeni", "Outlook masaüstü (Word motoru) bunları yok sayar; sütunlu düzen için &lt;table&gt; kullanın.");
    if (/position\s*:\s*(absolute|fixed)/i.test(html))
      add("warn", "position: absolute/fixed", "Çoğu istemcide çalışmaz, öğeler üst üste biner.");
    if (/float\s*:\s*(left|right)/i.test(html))
      add("warn", "float kullanımı", "Outlook'ta güvenilir değil; tablo hücreleri daha sağlamdır.");
    if (/background-image\s*:|<[^>]+\sbackground\s*=/i.test(html))
      add("warn", "Arka plan resmi", "Outlook arka plan resimlerini genelde göstermez; düz renk yedeği verin.");
    if (/\d\s*(vh|vw)\b/i.test(html))
      add("warn", "vh / vw birimleri", "E-postada desteklenmez; px kullanın.");
    if (/<img[^>]+src\s*=\s*["']?https?:/i.test(html))
      add("info", "Dışarıdan resim", "Alıcı 'resimleri göster' demeden görünmez. Editörden eklediğiniz resimler ise maile gömülür.");
    const imgs = html.match(/<img\b[^>]*>/gi) || [];
    const altsiz = imgs.filter((t) => !/\salt\s*=/i.test(t)).length;
    if (altsiz)
      add("info", altsiz + " resimde alt metni yok", "Resim yüklenmezse yerine ne yazacağını belirtmek teslimat ve erişilebilirlik için iyidir.");
    let bayt = 0;
    try { bayt = new Blob([html]).size; } catch (e) { bayt = html.length; }
    if (bayt > 102400)
      add("warn", "Boyut " + Math.round(bayt / 1024) + " KB", "Gmail 102 KB üstünü kırpar ve 'Tümünü göster' bağlantısı koyar. Gömülü resimleri küçültün.");
    if (/<style[\s>]/i.test(html) && !/style\s*=/i.test(html))
      add("info", "Yalnızca &lt;style&gt; bloğu", "Gmail &lt;style&gt;'ı çoğu durumda korur ama en güvenlisi satır içi style özniteliğidir.");

    // Bilinmeyen yer tutucular
    const bilinmeyen = [];
    let m;
    PH_RE.lastIndex = 0;
    while ((m = PH_RE.exec(html))) {
      const p = parsePh(m[0]);
      if (p && !st.keys.has(p.key) && looksLikeField(p) && bilinmeyen.indexOf(p.name) < 0)
        bilinmeyen.push(p.name);
    }
    if (bilinmeyen.length)
      add("err", "Tanınmayan alan: " + bilinmeyen.map((x) => "{" + x + "}").join(", "),
        "Bu adda bir Excel sütunu yok. Mail metninde olduğu gibi görünür — sütun adını düzeltin ya da varsayılan ekleyin: {Ad|Sayın Müşterimiz}");
    return out;
  }

  function openLint() {
    const anchor = root.querySelector('[data-act="lint"]') || statusBar;
    const items = lint(value());
    popup(anchor, (p) => {
      p.style.minWidth = "340px";
      p.style.maxWidth = "420px";
      p.appendChild(el("div", "ed-pop-title", "E-posta uyumluluk denetimi"));
      const box = el("div", "ed-lint");
      if (!items.length) {
        box.appendChild(el("div", "ed-lint-item info",
          '<span class="ic">' + svg(I.check) + '</span><span class="tx"><b>Sorun bulunamadı</b>' +
          "<span>İçerik yaygın e-posta istemcileriyle uyumlu görünüyor.</span></span>"));
      }
      items.forEach((it) => {
        box.appendChild(el("div", "ed-lint-item " + it.lvl,
          '<span class="ic">' + svg(it.lvl === "info" ? I.info : I.warn) + "</span>" +
          '<span class="tx"><b>' + it.b + "</b><span>" + it.d + "</span></span>"));
      });
      p.appendChild(box);
    });
  }

  /* ------------------------------------------------------------ önizleme */
  function openPreview() {
    if (!st.api || !st.getJob) { toast("Önizleme için Excel bağlantısı gerekiyor.", "error"); return; }
    let idx = 0;
    previewBox.classList.remove("hidden");
    const frameEl = previewBox.querySelector("iframe");
    const pos = previewBox.querySelector(".pos");
    const meta = previewBox.querySelector(".ed-prev-meta");

    function goster(res) {
      if (!res) return;
      idx = res.index || 0;
      pos.textContent = res.total ? (idx + 1) + " / " + res.total : "veri yok";
      const uyari = res.unknown && res.unknown.length
        ? '<div class="ln"><span class="k">Uyarı</span><span class="v bad">Tanınmayan alan: ' +
          res.unknown.map((x) => "{" + esc(x) + "}").join(", ") + "</span></div>"
        : "";
      meta.innerHTML =
        '<div class="ln"><span class="k">Kime</span><span class="v">' + esc(res.email || "—") + "</span></div>" +
        '<div class="ln"><span class="k">Konu</span><span class="v">' + esc(res.subject || "(boş)") + "</span></div>" +
        '<div class="ln"><span class="k">Ek</span><span class="v' + (res.attachment && !res.attachment_ok ? " bad" : "") + '">' +
          (res.attachment_name ? esc(res.attachment_name) + (res.attachment_ok ? "" : "  — dosya bulunamadı!") : "yok") +
        "</span></div>" +
        '<div class="ln"><span class="k">Excel satırı</span><span class="v">' + esc(String(res.row)) + "</span></div>" +
        uyari;

      const d = frameEl.contentDocument;
      d.open();
      if (res.is_html) {
        d.write(isFullDoc(res.body) ? res.body
          : '<!doctype html><meta charset="utf-8"><body style="margin:0;padding:16px;' +
            'font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.6">' + res.body + "</body>");
      } else {
        d.write('<!doctype html><meta charset="utf-8"><body style="margin:0;padding:16px;' +
          'font-family:Arial,Helvetica,sans-serif;font-size:14px;white-space:pre-wrap">' +
          esc(res.body) + "</body>");
      }
      d.close();
    }

    function yukle(i) {
      const job = st.getJob();
      job.body = value();
      st.api.onizleme(job, i).then(goster).catch((e) => toast("Önizleme hatası: " + e, "error"));
    }

    previewBox.querySelector('[data-p="prev"]').onclick = () => yukle(Math.max(0, idx - 1));
    previewBox.querySelector('[data-p="next"]').onclick = () => yukle(idx + 1);
    previewBox.querySelector('[data-p="close"]').onclick = () => previewBox.classList.add("hidden");
    previewBox.querySelector('[data-p="device"]').onclick = (e) => {
      const b = previewBox.querySelector(".ed-prev-body");
      const mobil = b.classList.toggle("mobil");
      e.currentTarget.innerHTML = svg(mobil ? I.monitor : I.phone) + (mobil ? "Masaüstü" : "Mobil");
    };
    yukle(0);
  }

  /* ================================================================
     Kurulum
     ================================================================ */
  function btn(icon, title, act, opts) {
    const b = el("button", "ed-btn" + (opts && opts.txt ? " txt" : ""));
    b.type = "button";
    b.title = title;
    b.innerHTML = (icon ? svg(icon) : "") + (opts && opts.label ? opts.label : "");
    b.dataset.act = act;
    if (opts && opts.needDesign) b.setAttribute("data-need-design", "1");
    // mousedown'da odağı çalmayız: iframe'deki seçim korunur.
    b.addEventListener("mousedown", (e) => e.preventDefault());
    return b;
  }

  const FONTS = ["Arial", "Helvetica", "Verdana", "Tahoma", "Trebuchet MS", "Georgia",
    "Times New Roman", "Courier New", "Segoe UI", "Calibri"];
  const SIZES = [10, 11, 12, 13, 14, 16, 18, 20, 24, 28, 32, 40];
  const BLOCKS = [["p", "Paragraf"], ["h1", "Başlık 1"], ["h2", "Başlık 2"],
    ["h3", "Başlık 3"], ["blockquote", "Alıntı"], ["pre", "Kod"]];

  function buildToolbar() {
    barMain = el("div", "ed-bar ed-bar-main");
    const add = (n) => barMain.appendChild(n);
    const sep = () => barMain.appendChild(el("div", "ed-sep"));

    add(btn(I.undo, "Geri al (Ctrl+Z)", "undo", { needDesign: true }));
    add(btn(I.redo, "Yinele (Ctrl+Y)", "redo", { needDesign: true }));
    sep();

    const blk = el("select", "ed-select w-block");
    BLOCKS.forEach(([v, t]) => {
      const o = document.createElement("option"); o.value = v; o.textContent = t; blk.appendChild(o);
    });
    blk.title = "Paragraf biçimi";
    blk.setAttribute("data-need-design", "1");
    blk.onchange = () => { exec("formatBlock", "<" + blk.value + ">"); };
    add(blk);

    const font = el("select", "ed-select w-font");
    font.innerHTML = '<option value="">Yazı tipi</option>' +
      FONTS.map((f) => '<option value="' + f + '" style="font-family:' + f + '">' + f + "</option>").join("");
    font.title = "Yazı tipi (e-posta güvenli listesi)";
    font.setAttribute("data-need-design", "1");
    font.onchange = () => { if (font.value) exec("fontName", font.value); font.selectedIndex = 0; };
    add(font);

    const size = el("select", "ed-select w-size");
    size.innerHTML = '<option value="">Punto</option>' +
      SIZES.map((s) => '<option value="' + s + '">' + s + "</option>").join("");
    size.title = "Yazı boyutu";
    size.setAttribute("data-need-design", "1");
    size.onchange = () => { if (size.value) setFontSize(size.value); size.selectedIndex = 0; };
    add(size);
    sep();

    ["bold:Kalın (Ctrl+B)", "italic:İtalik (Ctrl+I)", "underline:Altı çizili (Ctrl+U)",
     "strikeThrough:Üstü çizili"].forEach((x) => {
      const [cmd, t] = x.split(":");
      const icon = { bold: I.bold, italic: I.italic, underline: I.underline, strikeThrough: I.strike }[cmd];
      const b = btn(icon, t, cmd, { needDesign: true });
      stateBtns[cmd] = b;
      add(b);
    });
    add(btn(I.color, "Yazı rengi", "foreColor", { needDesign: true }));
    add(btn(I.fill, "Vurgu rengi", "hiliteColor", { needDesign: true }));
    sep();

    [["justifyLeft", I.alignL, "Sola hizala"], ["justifyCenter", I.alignC, "Ortala"],
     ["justifyRight", I.alignR, "Sağa hizala"], ["justifyFull", I.alignJ, "İki yana yasla"]]
      .forEach(([cmd, ic, t]) => {
        const b = btn(ic, t, cmd, { needDesign: true });
        stateBtns[cmd] = b;
        add(b);
      });
    sep();

    [["insertUnorderedList", I.ul, "Madde işaretli liste"],
     ["insertOrderedList", I.ol, "Numaralı liste"]].forEach(([cmd, ic, t]) => {
      const b = btn(ic, t, cmd, { needDesign: true });
      stateBtns[cmd] = b;
      add(b);
    });
    add(btn(I.outdent, "Girintiyi azalt", "outdent", { needDesign: true }));
    add(btn(I.indent, "Girintiyi artır", "indent", { needDesign: true }));
    sep();

    add(btn(I.link, "Bağlantı (Ctrl+K)", "link", { needDesign: true }));
    add(btn(I.image, "Resim ekle (maile gömülür)", "image", { needDesign: true }));
    add(btn(I.table, "Tablo ekle", "table", { needDesign: true }));
    add(btn(I.hr, "Yatay çizgi", "insertHorizontalRule", { needDesign: true }));
    add(btn(I.quote, "Alıntı bloğu", "quote", { needDesign: true }));
    add(btn(I.clear, "Biçimi temizle", "removeFormat", { needDesign: true }));

    /* ---- ikinci sıra ---- */
    barSecond = el("div", "ed-bar ed-bar-second");
    const add2 = (n) => barSecond.appendChild(n);

    const alan = btn(I.field, "Excel sütunu ekle (Ctrl+Shift+A)", "field",
      { txt: true, label: "Alan Ekle" });
    alan.classList.add("on");
    add2(alan);
    add2(el("div", "ed-sep"));

    add2(btn(I.eye, "Gerçek veriyle önizle", "preview", { txt: true, label: "Önizle" }));
    add2(btn(I.shield, "E-posta uyumluluğunu denetle", "lint", { txt: true, label: "Denetle" }));
    add2(btn(I.tidy, "HTML'i düzenli yaz", "beautify", { txt: true, label: "Biçimlendir" }));
    add2(btn(I.broom, "Word'den gelen fazlalıkları temizle", "clean", { txt: true, label: "Word Temizle" }));
    add2(el("div", "ed-sep"));
    add2(btn(I.folder, "Şablonlar", "tpl", { txt: true, label: "Şablon" }));
    add2(btn(I.open, "HTML dosyasından yükle", "import"));
    add2(btn(I.save, "HTML dosyası olarak kaydet", "export"));

    add2(el("div", "ed-spacer"));

    const modes = el("div", "ed-modes");
    [["design", I.design, "Tasarım"], ["code", I.code, "Kod"], ["split", I.split, "Bölünmüş"]]
      .forEach(([m, ic, t]) => {
        const b = el("button", m === "design" ? "on" : "", svg(ic) + t);
        b.type = "button";
        b.dataset.mode = m;
        b.title = t + " görünümü";
        b.onclick = () => setMode(m);
        modes.appendChild(b);
      });
    add2(modes);
    add2(btn(I.full, "Tam ekran (Ctrl+Shift+F)", "full", { txt: true, label: "Tam ekran" }));
  }

  function buildBody() {
    bodyBox = el("div", "ed-body");

    designPane = el("div", "ed-pane");
    frame = el("iframe", "ed-frame");
    frame.setAttribute("title", "Mail içeriği");
    designPane.appendChild(frame);
    bodyBox.appendChild(designPane);

    codePane = el("div", "ed-pane hidden");
    codeBox = el("div", "ed-code");
    gutter = el("div", "ed-gutter");
    const wrap = el("div", "ed-code-wrap");
    hlBox = el("pre", "ed-hl");
    srcArea = el("textarea", "ed-src");
    srcArea.spellcheck = false;
    srcArea.setAttribute("wrap", "off");
    srcArea.setAttribute("aria-label", "HTML kaynağı");
    wrap.appendChild(hlBox);
    wrap.appendChild(srcArea);
    codeBox.appendChild(gutter);
    codeBox.appendChild(wrap);
    codePane.appendChild(codeBox);
    bodyBox.appendChild(codePane);
  }

  function buildStatus() {
    statusBar = el("div", "ed-status");
    statusBar.innerHTML =
      '<span data-s="mode"></span><span data-s="size"></span>' +
      '<span data-s="fields"></span>' +
      '<span class="push" data-s="lint"></span>';
    statusBar.querySelector('[data-s="lint"]').onclick = openLint;
  }

  function updateStatus() {
    if (!statusBar) return;
    const v = st.mode === "code" || st.mode === "split" ? srcArea.value : st.value;
    const q = (s) => statusBar.querySelector('[data-s="' + s + '"]');
    let bayt = 0;
    try { bayt = new Blob([v || ""]).size; } catch (e) { bayt = (v || "").length; }
    q("mode").innerHTML = st.html ? "<b>HTML</b>" : "<b>Düz metin</b>";
    q("size").textContent = (v || "").length.toLocaleString("tr-TR") + " karakter · " +
      (bayt > 1024 ? (bayt / 1024).toFixed(1) + " KB" : bayt + " B");

    // Kullanılan alanlar
    const kullanilan = [], bilinmeyen = [];
    let m;
    PH_RE.lastIndex = 0;
    while ((m = PH_RE.exec(v || ""))) {
      const p = parsePh(m[0]);
      if (!p) continue;
      if (st.keys.has(p.key)) { if (kullanilan.indexOf(p.name) < 0) kullanilan.push(p.name); }
      else if (looksLikeField(p) && bilinmeyen.indexOf(p.name) < 0) bilinmeyen.push(p.name);
    }
    q("fields").innerHTML = kullanilan.length
      ? '<span class="good">' + kullanilan.length + " alan kullanılıyor: " +
        esc(kullanilan.slice(0, 4).join(", ")) + (kullanilan.length > 4 ? "…" : "") + "</span>"
      : "";

    const uyari = bilinmeyen.length;
    q("lint").className = "push " + (uyari ? "bad" : "warn");
    q("lint").innerHTML = uyari
      ? svg(I.warn) + " " + uyari + " tanınmayan alan — denetle"
      : svg(I.shield) + " Uyumluluğu denetle";
    q("lint").style.display = "inline-flex";
    q("lint").style.alignItems = "center";
    q("lint").style.gap = "5px";
  }

  /* ------------------------------------------------------------- eylemler */
  function handleAction(act, target) {
    switch (act) {
      case "undo": case "redo": case "removeFormat":
      case "insertUnorderedList": case "insertOrderedList":
      case "indent": case "outdent": case "insertHorizontalRule":
      case "bold": case "italic": case "underline": case "strikeThrough":
      case "justifyLeft": case "justifyCenter": case "justifyRight": case "justifyFull":
        exec(act); break;
      case "quote": exec("formatBlock", "<blockquote>"); break;
      case "foreColor": openColor(target, "foreColor"); break;
      case "hiliteColor": openColor(target, "hiliteColor"); break;
      case "link": openLink(target); break;
      case "table": openTable(target); break;
      case "field": openFields(target, "body"); break;
      case "image":
        if (st.api && st.api.resim_yukle) {
          st.api.resim_yukle().then((res) => {
            if (!res) return;
            if (res.error) { toast(res.error, "error"); return; }
            if (res.uri) {
              if (res.size > 2 * 1024 * 1024)
                toast("Resim " + Math.round(res.size / 1024) + " KB — büyük resimler maili yavaşlatır.", "info");
              insertImage(res.uri, res.name);
            }
          });
        } else toast("Resim seçme yalnızca uygulama içinde çalışır.", "error");
        break;
      case "preview": openPreview(); break;
      case "lint": openLint(); break;
      case "beautify": {
        const v = value();
        const yeni = beautify(v);
        setValue(yeni);
        toast("HTML düzenli yazıldı.", "success");
        break;
      }
      case "clean": {
        const v = value();
        const yeni = cleanWord(v);
        if (yeni === v) { toast("Temizlenecek Word fazlalığı bulunamadı.", "info"); break; }
        const kazanc = v.length - yeni.length;
        setValue(yeni);
        toast("Word fazlalıkları temizlendi (" + kazanc.toLocaleString("tr-TR") + " karakter).", "success");
        break;
      }
      case "tpl": openTemplates(false); break;
      case "import":
        if (st.api && st.api.html_yukle) {
          st.api.html_yukle().then((res) => {
            if (!res) return;
            if (res.error) { toast(res.error, "error"); return; }
            if (res.content != null) {
              setValue(res.content);
              if (!st.html) setHtmlMode(true);
              if (st.onChange) st.onChange();
              toast("HTML içerik yüklendi.", "success");
            }
          });
        }
        break;
      case "export":
        if (st.api && st.api.html_kaydet) {
          st.api.html_kaydet(value(), "mail-sablonu.html").then((res) => {
            if (res && res.error) toast(res.error, "error");
            else if (res && res.path) toast("Kaydedildi: " + res.path, "success");
          });
        }
        break;
      case "full": setFull(!st.full); break;
      default: break;
    }
  }

  function onFrameKey(e) {
    const c = e.ctrlKey || e.metaKey;
    if (c && e.shiftKey && e.key.toLowerCase() === "f") { e.preventDefault(); setFull(!st.full); return; }
    if (c && e.shiftKey && e.key.toLowerCase() === "a") {
      e.preventDefault();
      openFields(root.querySelector('[data-act="field"]'), "body");
      return;
    }
    if (c && !e.shiftKey && e.key.toLowerCase() === "k") { e.preventDefault(); openLink(root.querySelector('[data-act="link"]')); return; }
    if (e.key === "Escape" && st.full) { setFull(false); return; }
    // Ctrl+B/I/U tarayıcının kendi işidir; styleWithCSS için biz yapalım
    if (c && !e.shiftKey && "biu".indexOf(e.key.toLowerCase()) >= 0) {
      e.preventDefault();
      exec({ b: "bold", i: "italic", u: "underline" }[e.key.toLowerCase()]);
    }
  }

  /* ================================================================
     Genel API
     ================================================================ */
  function setValue(v) {
    st.value = v == null ? "" : String(v);
    st.designDirty = false;
    if (st.mode === "code" || st.mode === "split") { srcArea.value = st.value; refreshCode(true); }
    if (st.mode === "design" || st.mode === "split") writeFrame();
    updateStatus();
  }

  function setHtmlMode(on) {
    const yeni = !!on;
    if (yeni === st.html) return;
    st.html = yeni;
    value();                       // mevcut değeri topla
    root.querySelector(".ed-modes").style.display = yeni ? "" : "none";
    updateBarState();
    if (!yeni) {
      // Düz metin: yalnızca kaynak kutusu, vurgulamasız
      if (st.mode !== "code") setMode("code");
      else { srcArea.value = st.value; refreshCode(true); }
    } else {
      setMode("design");
    }
    updateStatus();
  }

  function setFields(f) {
    st.fields = {
      columns: (f && f.columns) || [],
      builtins: (f && f.builtins) || [],
      formats: (f && f.formats) || [],
    };
    st.keys = new Set();
    st.fields.columns.forEach((c) => st.keys.add(normName(c)));
    st.fields.builtins.forEach(([ad]) => st.keys.add(normName(ad)));
    st.fields.formats.forEach(([k]) => { if (k) FORMAT_KEYS.add(normName(k)); });
    if (st.mode !== "code") decorateChips();
    if (st.mode !== "design") refreshCode(true);
    updateStatus();
  }

  function setSample(row) {
    st.sample = row || null;
    if (st.mode !== "code") decorateChips();
  }

  function mount(el0, opts) {
    opts = opts || {};
    st.api = opts.api || null;
    st.getJob = opts.getJob || null;
    st.onChange = opts.onChange || null;
    st.onToast = opts.onToast || null;

    root = el("div", "ed");
    backdrop = el("div", "ed-backdrop hidden");
    backdrop.onclick = () => setFull(false);
    document.body.appendChild(backdrop);

    buildToolbar();
    buildBody();
    buildStatus();
    root.appendChild(barMain);
    root.appendChild(barSecond);
    root.appendChild(bodyBox);
    root.appendChild(statusBar);
    el0.appendChild(root);

    // Önizleme penceresi
    previewBox = el("div", "ed-prev-backdrop hidden");
    previewBox.innerHTML =
      '<div class="ed-prev">' +
        '<div class="ed-prev-head">' +
          "<h3>Gerçek veriyle önizleme</h3>" +
          '<button class="ed-btn txt" data-p="device">' + svg(I.phone) + "Mobil</button>" +
          '<div class="ed-prev-nav">' +
            '<button class="ed-btn" data-p="prev" title="Önceki satır">' + svg(I.left) + "</button>" +
            '<span class="pos">1 / 1</span>' +
            '<button class="ed-btn" data-p="next" title="Sonraki satır">' + svg(I.right) + "</button>" +
            '<button class="ed-btn" data-p="close" title="Kapat">' + svg(I.x) + "</button>" +
          "</div>" +
        "</div>" +
        '<div class="ed-prev-meta"></div>' +
        '<div class="ed-prev-body"><iframe title="Önizleme"></iframe></div>' +
      "</div>";
    document.body.appendChild(previewBox);
    previewBox.addEventListener("mousedown", (e) => {
      if (e.target === previewBox) previewBox.classList.add("hidden");
    });

    // Olaylar
    root.addEventListener("click", (e) => {
      const b = e.target.closest ? e.target.closest("[data-act]") : null;
      if (b && root.contains(b)) handleAction(b.dataset.act, b);
    });
    srcArea.addEventListener("input", onCodeInput);
    srcArea.addEventListener("keydown", onCodeKey);
    srcArea.addEventListener("scroll", syncScroll);
    srcArea.addEventListener("focus", () => root.classList.add("focus"));
    srcArea.addEventListener("blur", () => root.classList.remove("focus"));
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape") {
        if (openPop) { closePop(); return; }
        if (!previewBox.classList.contains("hidden")) { previewBox.classList.add("hidden"); return; }
        if (st.full) setFull(false);
      }
    });

    writeFrame();
    updateBarState();
    updateStatus();
    return API;
  }

  const API = {
    mount: mount,
    getValue: value,
    setValue: setValue,
    setHtmlMode: setHtmlMode,
    isHtmlMode: () => st.html,
    setFields: setFields,
    setSample: setSample,
    insertField: insertField,
    openFields: openFields,
    focus: () => { if (st.mode === "code") srcArea.focus(); else focusFrame(); },
    // test/hata ayıklama için
    _beautify: beautify,
    _cleanWord: cleanWord,
    _lint: lint,
    _parsePh: parsePh,
    _normName: normName,
    _state: st,
  };
  return API;
})();
