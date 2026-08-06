"""
Toplu Fatura Mail Gönderici - Grafik Arayüz (Tkinter)

Excel dosyasındaki her satır için (Email + Ek yolu), seçtiğiniz konu ve içerikle,
Outlook veya SMTP üzerinden otomatik mail gönderir.

Çalıştırmak için:  python app.py
"""

from __future__ import annotations

import os
import json
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

import core

APP_TITLE = "Toplu Fatura Mail Gönderici"

# Log kutusunda tutulacak en fazla satır. Tk Text widget'ı on binlerce satırda
# gözle görülür yavaşlar; tam kayıt zaten gonderim_sonuclari.csv dosyasındadır.
LOG_MAX_LINES = 1500


def _settings_file() -> str:
    """Ayar dosyasının tam yolu.

    Ayarlar, exe'nin/masaüstünün yanında DEĞİL, Windows'un uygulama-verisi
    klasöründe tutulur:  %APPDATA%\\TopluFaturaMailer\\ayarlar.json

    Böylece:
      - masaüstü / exe klasörü temiz kalır,
      - dosya kullanıcıya özeldir ve kalıcıdır (Temp gibi silinmez),
      - .exe ve normal (python app.py) çalıştırmada aynı yeri kullanır.
    """
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    folder = os.path.join(base, "TopluFaturaMailer")
    try:
        os.makedirs(folder, exist_ok=True)
    except Exception:  # noqa: BLE001
        folder = base  # klasör oluşturulamazsa en azından APPDATA köküne yaz
    return os.path.join(folder, "ayarlar.json")


SETTINGS_FILE = _settings_file()


def _obfuscate(text: str) -> str:
    """Şifreyi düz metin görünmesin diye hafifçe gizler (base64).

    DİKKAT: Bu GERÇEK şifreleme DEĞİLDİR; sadece dosyayı açan birinin şifreyi
    çıplak gözle okumasını engeller. Dosya .gitignore'da olduğu için yayınlanmaz.
    """
    import base64
    return base64.b64encode((text or "").encode("utf-8")).decode("ascii")


def _deobfuscate(text: str) -> str:
    import base64
    try:
        return base64.b64decode((text or "").encode("ascii")).decode("utf-8")
    except Exception:  # noqa: BLE001
        return ""


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("820x760")
        self.minsize(760, 640)

        self._queue = queue.Queue()   # worker -> GUI mesajları
        self._worker = None
        self._stop_flag = threading.Event()
        self._headers = []

        self._build_ui()
        self._load_settings()
        self.after(100, self._drain_queue)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------------------------------------------------------------- UI
    def _build_ui(self):
        pad = {"padx": 6, "pady": 3}
        root = ttk.Frame(self, padding=10)
        root.pack(fill="both", expand=True)

        # --- Excel dosyası ---
        f_file = ttk.LabelFrame(root, text="1) Excel dosyası", padding=8)
        f_file.pack(fill="x", **pad)
        self.var_xlsx = tk.StringVar()
        ttk.Entry(f_file, textvariable=self.var_xlsx).pack(side="left", fill="x", expand=True)
        ttk.Button(f_file, text="Gözat...", command=self._pick_xlsx).pack(side="left", padx=6)
        ttk.Button(f_file, text="Sütunları Oku", command=self._load_columns).pack(side="left")

        # --- Sütun eşleme ---
        f_cols = ttk.LabelFrame(root, text="2) Sütun eşleme", padding=8)
        f_cols.pack(fill="x", **pad)
        ttk.Label(f_cols, text="Sayfa:").grid(row=0, column=0, sticky="e", **pad)
        self.var_sheet = tk.StringVar()
        self.cmb_sheet = ttk.Combobox(f_cols, textvariable=self.var_sheet, width=18, state="readonly")
        self.cmb_sheet.grid(row=0, column=1, sticky="w", **pad)
        self.cmb_sheet.bind("<<ComboboxSelected>>", lambda e: self._load_columns())

        ttk.Label(f_cols, text="E-posta sütunu:").grid(row=0, column=2, sticky="e", **pad)
        self.var_email_col = tk.StringVar()
        self.cmb_email = ttk.Combobox(f_cols, textvariable=self.var_email_col, width=28, state="readonly")
        self.cmb_email.grid(row=0, column=3, sticky="w", **pad)

        ttk.Label(f_cols, text="Ek (dosya yolu) sütunu:").grid(row=1, column=2, sticky="e", **pad)
        self.var_attach_col = tk.StringVar()
        self.cmb_attach = ttk.Combobox(f_cols, textvariable=self.var_attach_col, width=28, state="readonly")
        self.cmb_attach.grid(row=1, column=3, sticky="w", **pad)

        # --- Gönderim yöntemi ---
        f_method = ttk.LabelFrame(root, text="3) Gönderim yöntemi", padding=8)
        f_method.pack(fill="x", **pad)
        self.var_method = tk.StringVar(value="outlook")
        ttk.Radiobutton(f_method, text="Outlook (masaüstü)", value="outlook",
                        variable=self.var_method, command=self._toggle_smtp).grid(row=0, column=0, **pad)
        ttk.Radiobutton(f_method, text="SMTP sunucusu", value="smtp",
                        variable=self.var_method, command=self._toggle_smtp).grid(row=0, column=1, **pad)
        ttk.Label(f_method, text="Gönderen adres:").grid(row=0, column=2, sticky="e", **pad)
        self.var_sender = tk.StringVar(value="muhasebe@atakonline.com")
        ttk.Entry(f_method, textvariable=self.var_sender, width=32).grid(row=0, column=3, sticky="w", **pad)

        # SMTP alt panel
        self.f_smtp = ttk.Frame(f_method)
        self.f_smtp.grid(row=1, column=0, columnspan=4, sticky="we", pady=(4, 0))
        ttk.Label(self.f_smtp, text="Sunucu:").grid(row=0, column=0, sticky="e", **pad)
        self.var_host = tk.StringVar()
        ttk.Entry(self.f_smtp, textvariable=self.var_host, width=26).grid(row=0, column=1, sticky="w", **pad)
        ttk.Label(self.f_smtp, text="Port:").grid(row=0, column=2, sticky="e", **pad)
        self.var_port = tk.StringVar(value="587")
        ttk.Entry(self.f_smtp, textvariable=self.var_port, width=7).grid(row=0, column=3, sticky="w", **pad)
        ttk.Label(self.f_smtp, text="Güvenlik:").grid(row=0, column=4, sticky="e", **pad)
        self.var_security = tk.StringVar(value="starttls")
        ttk.Combobox(self.f_smtp, textvariable=self.var_security, width=10, state="readonly",
                     values=["starttls", "ssl", "none"]).grid(row=0, column=5, sticky="w", **pad)
        ttk.Label(self.f_smtp, text="Kullanıcı:").grid(row=1, column=0, sticky="e", **pad)
        self.var_user = tk.StringVar()
        ttk.Entry(self.f_smtp, textvariable=self.var_user, width=26).grid(row=1, column=1, sticky="w", **pad)
        ttk.Label(self.f_smtp, text="Şifre:").grid(row=1, column=2, sticky="e", **pad)
        self.var_pass = tk.StringVar()
        ttk.Entry(self.f_smtp, textvariable=self.var_pass, width=20, show="•").grid(row=1, column=3, sticky="w", **pad)
        ttk.Label(self.f_smtp, text="(şifre de kaydedilir)").grid(row=1, column=4, columnspan=2, sticky="w", **pad)

        # --- Mail içeriği ---
        f_msg = ttk.LabelFrame(root, text="4) Mail içeriği", padding=8)
        f_msg.pack(fill="both", **pad)
        ttk.Label(f_msg, text="Konu:").grid(row=0, column=0, sticky="e", **pad)
        self.var_subject = tk.StringVar()
        ttk.Entry(f_msg, textvariable=self.var_subject, width=70).grid(row=0, column=1, columnspan=3, sticky="we", **pad)
        self.var_html = tk.BooleanVar(value=False)
        ttk.Checkbutton(f_msg, text="İçerik HTML", variable=self.var_html).grid(row=0, column=4, **pad)
        ttk.Label(f_msg, text="İçerik:").grid(row=1, column=0, sticky="ne", **pad)
        self.txt_body = scrolledtext.ScrolledText(f_msg, height=7, wrap="word")
        self.txt_body.grid(row=1, column=1, columnspan=4, sticky="we", **pad)
        # HTML gövde için yardımcı düğmeler
        f_body_btn = ttk.Frame(f_msg)
        f_body_btn.grid(row=2, column=1, columnspan=4, sticky="w", padx=6)
        ttk.Button(f_body_btn, text="HTML dosyasından yükle...", command=self._load_html).pack(side="left")
        ttk.Button(f_body_btn, text="Tarayıcıda önizle", command=self._preview_body).pack(side="left", padx=6)
        f_msg.columnconfigure(1, weight=1)

        # --- Çalışma ayarları ---
        f_run = ttk.LabelFrame(root, text="5) Çalışma ayarları", padding=8)
        f_run.pack(fill="x", **pad)
        ttk.Label(f_run, text="Mailler arası bekleme (sn):").grid(row=0, column=0, sticky="e", **pad)
        self.var_delay = tk.StringVar(value="1.0")
        ttk.Entry(f_run, textvariable=self.var_delay, width=7).grid(row=0, column=1, sticky="w", **pad)
        ttk.Label(f_run, text="Adet sınırı (0=hepsi):").grid(row=0, column=2, sticky="e", **pad)
        self.var_limit = tk.StringVar(value="0")
        ttk.Entry(f_run, textvariable=self.var_limit, width=7).grid(row=0, column=3, sticky="w", **pad)

        ttk.Label(f_run, text="TEST adresi (doluysa herkese yerine buraya):").grid(row=1, column=0, columnspan=2, sticky="e", **pad)
        self.var_test = tk.StringVar()
        ttk.Entry(f_run, textvariable=self.var_test, width=30).grid(row=1, column=2, columnspan=2, sticky="w", **pad)

        # --- Aksiyon butonları ---
        f_act = ttk.Frame(root)
        f_act.pack(fill="x", **pad)
        self.btn_check = ttk.Button(f_act, text="🔍 Kontrol Et (göndermeden)", command=self._check)
        self.btn_check.pack(side="left")
        self.btn_start = ttk.Button(f_act, text="▶ Gönderimi Başlat", command=self._start)
        self.btn_start.pack(side="left", padx=6)
        self.btn_stop = ttk.Button(f_act, text="■ Durdur", command=self._stop, state="disabled")
        self.btn_stop.pack(side="left", padx=6)
        self.progress = ttk.Progressbar(f_act, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=10)
        self.lbl_count = ttk.Label(f_act, text="0 / 0")
        self.lbl_count.pack(side="left")

        # --- Log ---
        f_log = ttk.LabelFrame(root, text="Kayıt / Log", padding=6)
        f_log.pack(fill="both", expand=True, **pad)
        self.txt_log = scrolledtext.ScrolledText(f_log, height=10, wrap="word", state="disabled")
        self.txt_log.pack(fill="both", expand=True)
        self.txt_log.tag_config("ok", foreground="#128a12")
        self.txt_log.tag_config("error", foreground="#c01616")
        self.txt_log.tag_config("info", foreground="#333333")

        self._toggle_smtp()

    # ---------------------------------------------------------------- olaylar
    def _pick_xlsx(self):
        path = filedialog.askopenfilename(
            title="Excel dosyası seç",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Tümü", "*.*")],
        )
        if path:
            self.var_xlsx.set(path)
            self._load_columns()

    def _load_columns(self):
        path = self.var_xlsx.get().strip()
        if not path or not os.path.isfile(path):
            return
        try:
            sheets, headers = core.get_headers(path, self.var_sheet.get() or None)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(APP_TITLE, f"Excel okunamadı:\n{exc}")
            return
        self.cmb_sheet["values"] = sheets
        if not self.var_sheet.get() and sheets:
            self.var_sheet.set(sheets[0])
        self._headers = headers
        labels = [f"{_col_letter(i)} — {h or '(boş)'}" for i, h in enumerate(headers)]
        self.cmb_email["values"] = labels
        self.cmb_attach["values"] = labels
        # Akıllı varsayılan: başlığa göre tahmin
        self._guess_column(self.var_email_col, labels, headers, ["email", "e-posta", "mail", "adres"])
        self._guess_column(self.var_attach_col, labels, headers, ["ek", "attach", "dosya", "pdf", "yol", "path"])

    @staticmethod
    def _guess_column(var, labels, headers, keywords):
        if var.get():
            return
        for i, h in enumerate(headers):
            hl = (h or "").lower()
            if any(k in hl for k in keywords):
                var.set(labels[i])
                return

    def _load_html(self):
        path = filedialog.askopenfilename(
            title="HTML içerik dosyası seç",
            filetypes=[("HTML", "*.html *.htm"), ("Metin", "*.txt"), ("Tümü", "*.*")],
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as fh:
                content = fh.read()
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(APP_TITLE, f"Dosya okunamadı:\n{exc}")
            return
        self.txt_body.delete("1.0", "end")
        self.txt_body.insert("1.0", content)
        self.var_html.set(True)  # HTML dosyası yüklendi -> HTML modunu aç

    def _preview_body(self):
        """Gövdeyi geçici bir dosyaya yazıp varsayılan tarayıcıda açar."""
        import tempfile
        import webbrowser
        body = self.txt_body.get("1.0", "end-1c")
        if not body.strip():
            messagebox.showinfo(APP_TITLE, "Önizlenecek içerik yok.")
            return
        if self.var_html.get():
            html = body
        else:
            # Düz metni güvenli şekilde HTML'e çevir (satır sonları korunur)
            from html import escape
            html = "<pre style='font-family:sans-serif;white-space:pre-wrap'>" + escape(body) + "</pre>"
        try:
            fd, tmp = tempfile.mkstemp(suffix=".html", prefix="mail_onizleme_")
            with os.fdopen(fd, "w", encoding="utf-8") as fh:
                fh.write("<!doctype html><meta charset='utf-8'>" + html)
            webbrowser.open("file:///" + tmp.replace("\\", "/"))
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(APP_TITLE, f"Önizleme açılamadı:\n{exc}")

    def _toggle_smtp(self):
        state = "normal" if self.var_method.get() == "smtp" else "disabled"
        for child in self.f_smtp.winfo_children():
            try:
                child.configure(state=state)
            except tk.TclError:
                pass

    # ---------------------------------------------------------------- başlat/durdur
    def _collect_job(self) -> core.SendJob | None:
        try:
            email_idx = self.cmb_email["values"].index(self.var_email_col.get())
            attach_idx = self.cmb_attach["values"].index(self.var_attach_col.get())
        except ValueError:
            messagebox.showerror(APP_TITLE, "Lütfen e-posta ve ek sütunlarını seçin.")
            return None

        xlsx = self.var_xlsx.get().strip()
        results_path = os.path.join(os.path.dirname(xlsx) or ".", "gonderim_sonuclari.csv")

        def _f(v, d):
            try:
                return float(v)
            except ValueError:
                return d

        def _i(v, d):
            try:
                return int(float(v))
            except ValueError:
                return d

        return core.SendJob(
            xlsx_path=xlsx,
            sheet=self.var_sheet.get(),
            email_col=email_idx,
            attach_col=attach_idx,
            subject=self.var_subject.get(),
            body=self.txt_body.get("1.0", "end-1c"),
            is_html=self.var_html.get(),
            method=self.var_method.get(),
            sender=self.var_sender.get().strip(),
            smtp_host=self.var_host.get().strip(),
            smtp_port=_i(self.var_port.get(), 587),
            smtp_security=self.var_security.get(),
            smtp_user=self.var_user.get().strip(),
            smtp_password=self.var_pass.get(),
            delay=_f(self.var_delay.get(), 1.0),
            start_row=2,  # her zaman ilk veri satırından başla
            limit=_i(self.var_limit.get(), 0),
            test_to=self.var_test.get().strip(),
            results_path=results_path,
        )

    def _check(self):
        """Hiçbir mail göndermeden tüm listeyi tarar ve rapor verir.

        Tarama binlerce dosya sorgusu içerebildiğinden ARKA PLANDA çalışır;
        aksi halde Tk penceresi tarama boyunca donar (yanıt vermiyor görünür).
        """
        if self._worker and self._worker.is_alive():
            return
        job = self._collect_job()
        if job is None:
            return
        if not os.path.isfile(job.xlsx_path):
            messagebox.showerror(APP_TITLE, "Önce geçerli bir Excel dosyası seçin.")
            return

        self._stop_flag.clear()
        self.progress["value"] = 0
        self.btn_check.configure(state="disabled")
        self.btn_start.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self._log("info", "=== KONTROL (göndermeden tarama) ===")

        def work():
            try:
                rep = core.check_job(
                    job,
                    on_progress=lambda done, total: self._queue.put(("scan", (done, total))),
                    should_stop=self._stop_flag.is_set,
                )
            except Exception as exc:  # noqa: BLE001
                self._queue.put(("check_error", str(exc)))
                return
            self._queue.put(("check_done", (job, rep)))

        self._worker = threading.Thread(target=work, daemon=True, name="kontrol")
        self._worker.start()

    def _show_check_report(self, job, rep):
        """Arka planda biten taramanın sonucunu log'a ve özet kutusuna döker."""
        self.btn_check.configure(state="normal")
        self.btn_start.configure(state="normal")
        self.btn_stop.configure(state="disabled")
        if rep.get("stopped"):
            self._log("info", "Kontrol durduruldu (kısmi sonuç).")
        self._log("info", f"Toplam satır: {rep['total']}  |  Sorunsuz: {rep['ok']}")
        self._log("info",
                  f"Geçersiz e-posta: {len(rep['bad_email'])}  |  "
                  f"Ek boş: {len(rep['empty_attach'])}  |  "
                  f"Ek bulunamadı: {len(rep['missing_attach'])}  |  "
                  f"Tekrarlı ek: {len(rep['duplicate_attach'])}")

        # Sorunlu satırları loga dök (ilk 50 kadarını)
        def dump(title, items, level="error"):
            if not items:
                return
            self._log(level, f"— {title} ({len(items)}):")
            for row, email, _tip, detay in items[:50]:
                self._log(level, f"    satır {row}: {email or '(boş)'}  →  {detay}")
            if len(items) > 50:
                self._log(level, f"    ... ve {len(items) - 50} satır daha (tam liste CSV'de)")

        dump("Geçersiz e-posta", rep["bad_email"])
        dump("Ek yolu boş", rep["empty_attach"])
        dump("Ek dosyası bulunamadı", rep["missing_attach"])
        dump("Tekrarlı ek (aynı dosya birden çok satırda)", rep["duplicate_attach"], "info")

        # İsteğe bağlı: sorunları CSV'ye yaz
        problems = rep["bad_email"] + rep["empty_attach"] + rep["missing_attach"] + rep["duplicate_attach"]
        if problems:
            out = os.path.join(os.path.dirname(job.xlsx_path) or ".", "kontrol_sorunlari.csv")
            try:
                import csv
                with open(out, "w", encoding="utf-8-sig", newline="") as fh:
                    w = csv.writer(fh)
                    w.writerow(["satir", "email", "sorun_tipi", "detay"])
                    w.writerows(problems)
                self._log("info", f"Sorun listesi kaydedildi: {out}")
            except Exception as exc:  # noqa: BLE001
                self._log("error", f"Sorun listesi yazılamadı: {exc}")

        # Özet kutusu
        total_problem = len(rep["bad_email"]) + len(rep["empty_attach"]) + len(rep["missing_attach"])
        if total_problem == 0:
            messagebox.showinfo(
                APP_TITLE,
                f"✔ Her şey yolunda.\n\nToplam {rep['total']} satırın tamamında "
                f"geçerli e-posta ve mevcut ek dosyası var.\n"
                f"(Tekrarlı ek: {len(rep['duplicate_attach'])} — bilgi amaçlı)",
            )
        else:
            messagebox.showwarning(
                APP_TITLE,
                f"Toplam satır: {rep['total']}\n"
                f"Sorunsuz: {rep['ok']}\n\n"
                f"Geçersiz e-posta: {len(rep['bad_email'])}\n"
                f"Ek boş: {len(rep['empty_attach'])}\n"
                f"Ek bulunamadı: {len(rep['missing_attach'])}\n"
                f"Tekrarlı ek (bilgi): {len(rep['duplicate_attach'])}\n\n"
                f"Ayrıntılar log alanında ve 'kontrol_sorunlari.csv' dosyasında.",
            )

    def _start(self):
        if self._worker and self._worker.is_alive():
            return
        job = self._collect_job()
        if job is None:
            return
        problems = core.validate_job(job)
        if problems:
            messagebox.showerror(APP_TITLE, "Eksik/hatalı ayarlar:\n\n- " + "\n- ".join(problems))
            return

        n = "TÜMÜ" if not job.limit else str(job.limit)
        mode = f"TEST → {job.test_to}" if job.test_to else "GERÇEK gönderim"
        if not messagebox.askyesno(
            APP_TITLE,
            f"{mode}\nYöntem: {job.method}\nGönderen: {job.sender}\n"
            f"Konu: {job.subject}\nGönderilecek: {n}\n\nBaşlatılsın mı?",
        ):
            return

        self._save_settings()
        self._stop_flag.clear()
        self.progress["value"] = 0
        self.btn_check.configure(state="disabled")
        self.btn_start.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self._log("info", "=== Gönderim başlatılıyor ===")

        self._worker = threading.Thread(target=self._run_worker, args=(job,), daemon=True)
        self._worker.start()

    def _run_worker(self, job):
        def on_progress(done, total, ok, fail):
            self._queue.put(("progress", (done, total, ok, fail)))

        def on_log(level, msg):
            self._queue.put(("log", (level, msg)))

        try:
            result = core.run_job(
                job,
                on_progress=on_progress,
                on_log=on_log,
                stop_event=self._stop_flag,      # bekleme Durdur'a anında yanıt versin
            )
        except Exception as exc:  # noqa: BLE001 - arayüz asılı kalmasın
            self._queue.put(("log", ("error", f"Beklenmeyen hata: {exc}")))
            result = {"total": 0, "sent": 0, "failed": 0, "stopped": True}
        self._queue.put(("done", result))

    def _stop(self):
        self._stop_flag.set()
        self.btn_stop.configure(state="disabled")
        self._log("info", "Durdurma istendi, mevcut işlem bitince duracak...")

    # ---------------------------------------------------------------- kuyruk
    def _drain_queue(self):
        """Worker mesajlarını TOPLU işler.

        Her mesaj için ayrı ayrı Text'e yazmak yerine satırlar biriktirilip tek
        insert ile eklenir; ilerleme çubuğu da yalnızca son değerle güncellenir.
        Böylece saniyede yüzlerce mesaj gelse bile arayüz akıcı kalır.
        """
        lines = []
        progress = None
        scan = None
        finals = []
        try:
            for _ in range(2000):                      # tek turda üst sınır
                kind, payload = self._queue.get_nowait()
                if kind == "progress":
                    progress = payload                 # sıkıştır: yalnızca sonuncusu
                elif kind == "scan":
                    scan = payload
                elif kind == "log":
                    lines.append(payload)
                else:
                    finals.append((kind, payload))
        except queue.Empty:
            pass

        if lines:
            self._log_many(lines)
        if progress is not None:
            done, total, ok, fail = progress
            self.progress["maximum"] = max(total, 1)
            self.progress["value"] = done
            self.lbl_count.configure(text=f"{done} / {total}  (✓{ok} ✗{fail})")
        elif scan is not None:
            done, total = scan
            self.progress["maximum"] = max(total, 1)
            self.progress["value"] = done
            self.lbl_count.configure(text=f"Kontrol: {done} / {total}")
        for kind, payload in finals:
            if kind == "done":
                self._on_done(payload)
            elif kind == "check_done":
                self._show_check_report(*payload)
            elif kind == "check_error":
                self.btn_check.configure(state="normal")
                self.btn_start.configure(state="normal")
                self.btn_stop.configure(state="disabled")
                messagebox.showerror(APP_TITLE, f"Kontrol sırasında hata:\n{payload}")

        self.after(120, self._drain_queue)

    def _on_done(self, result):
        self.btn_check.configure(state="normal")
        self.btn_start.configure(state="normal")
        self.btn_stop.configure(state="disabled")
        msg = (f"Tamamlandı.\n\nGönderilen: {result['sent']}\n"
               f"Hatalı: {result['failed']}")
        if result.get("stopped"):
            msg = "Durduruldu.\n\n" + msg
        self._log("info", "=== " + msg.replace("\n", " ") + " ===")
        messagebox.showinfo(APP_TITLE, msg)

    def _log(self, level, msg):
        self._log_many([(level, msg)])

    def _log_many(self, lines):
        """Birden çok log satırını tek seferde ekler ve toplam satırı sınırlar."""
        if not lines:
            return
        self.txt_log.configure(state="normal")
        # Aynı renk (tag) ardışık satırları tek insert'te birleştir.
        buf_level, buf = lines[0][0], []
        for level, msg in lines:
            if level != buf_level and buf:
                self.txt_log.insert("end", "".join(buf), buf_level)
                buf = []
            buf_level = level
            buf.append(msg + "\n")
        if buf:
            self.txt_log.insert("end", "".join(buf), buf_level)

        # Text widget'ı sınırsız büyürse kaydırma ve çizim yavaşlar.
        extra = int(self.txt_log.index("end-1c").split(".")[0]) - LOG_MAX_LINES
        if extra > 0:
            self.txt_log.delete("1.0", f"{extra + 1}.0")

        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    # ---------------------------------------------------------------- ayarlar
    def _save_settings(self):
        # Kullanıcı bir daha uğraşmasın diye TÜM alanlar kaydedilir (şifre dahil).
        data = {
            "xlsx": self.var_xlsx.get(),
            "sheet": self.var_sheet.get(),
            "email_col": self.var_email_col.get(),
            "attach_col": self.var_attach_col.get(),
            "subject": self.var_subject.get(),
            "body": self.txt_body.get("1.0", "end-1c"),
            "is_html": self.var_html.get(),
            "method": self.var_method.get(),
            "sender": self.var_sender.get(),
            "smtp_host": self.var_host.get(),
            "smtp_port": self.var_port.get(),
            "smtp_security": self.var_security.get(),
            "smtp_user": self.var_user.get(),
            # Şifre base64 ile hafifçe gizlenerek yazılır (bkz. _obfuscate).
            "smtp_password": _obfuscate(self.var_pass.get()),
            "delay": self.var_delay.get(),
            "limit": self.var_limit.get(),
            "test_to": self.var_test.get(),
        }
        try:
            with open(SETTINGS_FILE, "w", encoding="utf-8") as fh:
                json.dump(data, fh, ensure_ascii=False, indent=2)
        except Exception:  # noqa: BLE001
            pass

    def _load_settings(self):
        if not os.path.isfile(SETTINGS_FILE):
            return
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as fh:
                data = json.load(fh)
        except Exception:  # noqa: BLE001
            return
        # Önce dosya/sayfa, sonra sütunları oku ki kayıtlı sütun seçimleri geçerli olsun.
        self.var_xlsx.set(data.get("xlsx", ""))
        self.var_sheet.set(data.get("sheet", ""))
        if self.var_xlsx.get() and os.path.isfile(self.var_xlsx.get()):
            self._load_columns()  # combobox değerlerini doldurur (sütun seçimi için gerekli)
        self.var_email_col.set(data.get("email_col", ""))
        self.var_attach_col.set(data.get("attach_col", ""))
        self.var_subject.set(data.get("subject", ""))
        self.txt_body.insert("1.0", data.get("body", ""))
        self.var_html.set(data.get("is_html", False))
        self.var_method.set(data.get("method", "outlook"))
        self.var_sender.set(data.get("sender", "muhasebe@atakonline.com"))
        self.var_host.set(data.get("smtp_host", ""))
        self.var_port.set(data.get("smtp_port", "587"))
        self.var_security.set(data.get("smtp_security", "starttls"))
        self.var_user.set(data.get("smtp_user", ""))
        self.var_pass.set(_deobfuscate(data.get("smtp_password", "")))
        self.var_delay.set(data.get("delay", "1.0"))
        self.var_limit.set(data.get("limit", "0"))
        self.var_test.set(data.get("test_to", ""))
        self._toggle_smtp()

    def _on_close(self):
        if self._worker and self._worker.is_alive():
            if not messagebox.askyesno(APP_TITLE, "Gönderim sürüyor. Yine de çıkılsın mı?"):
                return
            self._stop_flag.set()
        self._save_settings()
        self.destroy()


def _col_letter(idx):
    """0->A, 1->B ... Excel sütun harfi."""
    s = ""
    idx += 1
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


if __name__ == "__main__":
    App().mainloop()
