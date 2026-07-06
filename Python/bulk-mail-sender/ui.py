"""
Toplu Fatura Mail Gönderici — Masaüstü arayüz (pywebview + HTML/CSS)

Arayüz yavuzyazici.com temasıyla web/ klasöründeki HTML ile çizilir; gönderim
mantığı değişmeden core.py / mailer.py üzerinden çalışır. Uygulama offline'dır;
tarayıcı/sunucu gerektirmez (Windows'ta gömülü Edge WebView2 kullanılır).

Çalıştırmak için:  python ui.py
"""

from __future__ import annotations

import os
import sys
import csv
import json
import base64
import tempfile
import threading
import webbrowser

import webview

import core


APP_TITLE = "Toplu Fatura Mail Gönderici"


# --------------------------------------------------------------------- yollar
def _app_web_dir() -> str:
    """web/ klasörünün yolu (derlenmişse PyInstaller geçici dizininde)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, "web")


def _settings_file() -> str:
    """Ayar dosyası: %APPDATA%\\TopluFaturaMailer\\ayarlar.json (exe'nin yanında değil)."""
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    folder = os.path.join(base, "TopluFaturaMailer")
    try:
        os.makedirs(folder, exist_ok=True)
    except Exception:  # noqa: BLE001
        folder = base
    return os.path.join(folder, "ayarlar.json")


def _obfuscate(text: str) -> str:
    """Şifreyi düz metin görünmesin diye base64 ile gizler (gerçek şifreleme DEĞİL)."""
    return base64.b64encode((text or "").encode("utf-8")).decode("ascii")


def _deobfuscate(text: str) -> str:
    try:
        return base64.b64decode((text or "").encode("ascii")).decode("utf-8")
    except Exception:  # noqa: BLE001
        return ""


class Api:
    """Arayüzün (JS) çağırdığı Python metotları. core.py'ye köprü kurar."""

    def __init__(self):
        self.window = None
        self._stop = threading.Event()
        self._worker = None
        # Worker -> arayüz olay kuyruğu (pull modeli). Thread'ler arası evaluate_js
        # çağrısı yapılmaz; arayüz olayları 'olaylari_al' ile çeker. Bu, pywebview'de
        # arka plandan GUI'ye erişimin yol açtığı donmayı/kilitlenmeyi önler.
        self._events = []
        self._lock = threading.Lock()

    # ---------------------------------------------------------- ayarlar
    def ayarlari_yukle(self):
        path = _settings_file()
        if not os.path.isfile(path):
            return {}
        try:
            with open(path, "r", encoding="utf-8") as fh:
                data = json.load(fh)
        except Exception:  # noqa: BLE001
            return {}
        data["smtp_password"] = _deobfuscate(data.get("smtp_password", ""))
        return data

    def ayarlari_kaydet(self, data):
        data = dict(data or {})
        data["smtp_password"] = _obfuscate(data.get("smtp_password", ""))
        try:
            with open(_settings_file(), "w", encoding="utf-8") as fh:
                json.dump(data, fh, ensure_ascii=False, indent=2)
        except Exception:  # noqa: BLE001
            pass

    # ---------------------------------------------------------- dosya
    def sec_dosya(self):
        res = self.window.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=("Excel (*.xlsx;*.xlsm)", "Tüm dosyalar (*.*)"),
        )
        if not res:
            return ""
        return res[0] if isinstance(res, (list, tuple)) else res

    def sutunlari_oku(self, path, sheet):
        if not path or not os.path.isfile(path):
            return {"error": "Dosya bulunamadı."}
        try:
            sheets, hdrs = core.get_headers(path, sheet or None)
        except Exception as exc:  # noqa: BLE001
            return {"error": str(exc)}
        return {"sheets": sheets, "headers": hdrs, "sheet": sheet or (sheets[0] if sheets else "")}

    def html_yukle(self):
        res = self.window.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=("HTML (*.html;*.htm)", "Metin (*.txt)", "Tüm dosyalar (*.*)"),
        )
        if not res:
            return {}
        p = res[0] if isinstance(res, (list, tuple)) else res
        try:
            with open(p, "r", encoding="utf-8") as fh:
                return {"content": fh.read()}
        except Exception as exc:  # noqa: BLE001
            return {"error": f"Dosya okunamadı: {exc}"}

    def onizle(self, body, is_html):
        if not (body or "").strip():
            return {"error": "Önizlenecek içerik yok."}
        if is_html:
            html = body
        else:
            from html import escape
            html = "<pre style='font-family:sans-serif;white-space:pre-wrap'>" + escape(body) + "</pre>"
        try:
            fd, tmp = tempfile.mkstemp(suffix=".html", prefix="mail_onizleme_")
            with os.fdopen(fd, "w", encoding="utf-8") as fh:
                fh.write("<!doctype html><meta charset='utf-8'>" + html)
            webbrowser.open("file:///" + tmp.replace("\\", "/"))
            return {}
        except Exception as exc:  # noqa: BLE001
            return {"error": f"Önizleme açılamadı: {exc}"}

    # ---------------------------------------------------------- iş kurma
    def _build_job(self, job):
        xlsx = (job.get("xlsx") or "").strip()
        results_path = os.path.join(os.path.dirname(xlsx) or ".", "gonderim_sonuclari.csv")
        return core.SendJob(
            xlsx_path=xlsx,
            sheet=job.get("sheet", ""),
            email_col=int(job.get("email_col") or 0),
            attach_col=int(job.get("attach_col") or 0),
            subject=job.get("subject", ""),
            body=job.get("body", ""),
            is_html=bool(job.get("is_html")),
            method=job.get("method", "outlook"),
            sender=(job.get("sender") or "").strip(),
            smtp_host=(job.get("smtp_host") or "").strip(),
            smtp_port=int(job.get("smtp_port") or 587),
            smtp_security=job.get("smtp_security", "starttls"),
            smtp_user=(job.get("smtp_user") or "").strip(),
            smtp_password=job.get("smtp_password", ""),
            delay=float(job.get("delay") or 0),
            start_row=2,
            limit=int(job.get("limit") or 0),
            test_to=(job.get("test") or "").strip(),
            results_path=results_path,
        )

    def dogrula(self, job):
        if job.get("email_col") is None or job.get("attach_col") is None:
            return ["Lütfen e-posta ve ek sütunlarını seçin."]
        return core.validate_job(self._build_job(job))

    def kontrol_et(self, job):
        try:
            return core.check_job(self._build_job(job))
        except Exception as exc:  # noqa: BLE001
            return {"error": f"Kontrol sırasında hata: {exc}"}

    # ---------------------------------------------------------- gönderim
    def _emit(self, ev):
        with self._lock:
            self._events.append(ev)

    def olaylari_al(self):
        """Arayüz düzenli aralıkla çağırır: biriken olayları döndürür ve temizler."""
        with self._lock:
            evs = self._events
            self._events = []
        return evs

    def gonder(self, job):
        if self._worker and self._worker.is_alive():
            return {"error": "Gönderim zaten sürüyor."}
        sendjob = self._build_job(job)
        problems = core.validate_job(sendjob)
        if problems:
            return {"error": "Eksik/hatalı ayarlar: " + ", ".join(problems)}
        self._stop.clear()
        with self._lock:
            self._events = []

        def on_progress(done, total, ok, fail):
            self._emit({"t": "progress", "done": done, "total": total, "ok": ok, "fail": fail})

        def on_log(level, msg):
            self._emit({"t": "log", "level": level, "msg": msg})

        def run():
            result = core.run_job(
                sendjob,
                on_progress=on_progress,
                on_log=on_log,
                should_stop=self._stop.is_set,
            )
            self._emit({"t": "done", "result": result})

        self._worker = threading.Thread(target=run, daemon=True)
        self._worker.start()
        return {"started": True}

    def durdur(self):
        self._stop.set()


def main():
    api = Api()
    index = os.path.join(_app_web_dir(), "index.html")
    window = webview.create_window(
        APP_TITLE,
        url=index,
        js_api=api,
        width=940,
        height=880,
        min_size=(820, 680),
        background_color="#f9f9fb",
    )
    api.window = window
    webview.start()


if __name__ == "__main__":
    main()
