"""
Toplu Fatura Mail Gönderici — Masaüstü arayüz (pywebview + HTML/CSS)

Arayüz yavuzyazici.com temasıyla web/ klasöründeki HTML ile çizilir; gönderim
mantığı değişmeden core.py / mailer.py üzerinden çalışır. Uygulama offline'dır;
tarayıcı/sunucu gerektirmez (Windows'ta gömülü Edge WebView2 kullanılır).

Çalıştırmak için:  python ui.py

Akıcılık (kasmama) ilkeleri
---------------------------
1. UZUN İŞ ANA İŞ PARÇACIĞINDA ÇALIŞMAZ. Excel okuma, ön kontrol ve gönderim
   arka planda çalışır; arayüz hiçbir zaman beklemez.
2. ARKA PLANDAN ARAYÜZE ÇAĞRI YAPILMAZ. Olaylar bir kuyruğa yazılır, arayüz
   'olaylari_al' ile toplu çeker (pywebview'de evaluate_js ile itmek kilitlenmeye
   yol açıyor).
3. OLAYLAR SIKIŞTIRILIR. İlerleme bildirimleri tek bir en güncel değere indirilir,
   log satırları listeye toplanır: 1300 mail = 1300 köprü çağrısı değil, ~30 çağrı.
4. HER İŞ BİR SONUÇ ÜRETİR. Worker'lar try/except ile sarılıdır; beklenmedik hata
   olsa bile arayüz sonsuza kadar "Gönderiliyor..." durumunda kalmaz.
5. HER ŞEY GÜNLÜĞE YAZILIR. Paketlenmiş (.exe) çalışmada konsol olmadığı için
   açılış aşamaları ve yavaş çağrılar %APPDATA%\\TopluFaturaMailer\\calisma.log
   dosyasına yazılır; "yanıt vermiyor" durumunda nerede takıldığı oradan görülür.
"""

from __future__ import annotations

import os
import sys
import csv
import json
import time
import base64
import tempfile
import threading
import traceback
import webbrowser
from collections import deque

import webview

import core
import merge


APP_TITLE = "Toplu Fatura Mail Gönderici"

# Arayüz çekene kadar bellekte tutulacak en fazla log satırı. Arayüz normalde
# 200 ms'de bir çektiği için bu sınıra ulaşılmaz; ulaşılırsa en eskiler düşer
# (tam kayıt zaten gonderim_sonuclari.csv dosyasında).
LOG_BUFFER = 4000

# Bu süreden uzun süren köprü çağrıları günlüğe "YAVAŞ" olarak yazılır.
YAVAS_ESIK = 0.30


# --------------------------------------------------------------------- yollar
def _app_web_dir() -> str:
    """web/ klasörünün yolu (derlenmişse PyInstaller geçici dizininde)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, "web")


def _veri_klasoru() -> str:
    """%APPDATA%\\TopluFaturaMailer — ayarlar ve günlük burada tutulur."""
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    folder = os.path.join(base, "TopluFaturaMailer")
    try:
        os.makedirs(folder, exist_ok=True)
    except Exception:  # noqa: BLE001
        folder = base
    return folder


def _settings_file() -> str:
    """Ayar dosyası: %APPDATA%\\TopluFaturaMailer\\ayarlar.json (exe'nin yanında değil)."""
    return os.path.join(_veri_klasoru(), "ayarlar.json")


def _sablon_klasoru() -> str:
    """Kullanıcının kaydettiği HTML şablonları: %APPDATA%\\TopluFaturaMailer\\sablonlar."""
    folder = os.path.join(_veri_klasoru(), "sablonlar")
    try:
        os.makedirs(folder, exist_ok=True)
    except OSError:
        pass
    return folder


def _sutun(deger, varsayilan: int) -> int:
    """Arayüzden gelen sütun indeksini güvenle çözer.

    JS tarafında seçim yapılmamışsa değer "" / null / NaN olabilir; bunların
    hiçbiri "0. sütun" anlamına GELMEZ.
    """
    try:
        n = int(deger)
    except (TypeError, ValueError):
        return varsayilan
    return varsayilan if n != n else n      # NaN koruması


def _guvenli_ad(ad: str) -> str:
    """YENİ şablon kaydederken adı dosya adına çevirir.

    Yol ayıracı, sürücü harfi ve '..' gibi parçaları eleyerek adın şablon
    klasörünün DIŞINA çıkamamasını garanti eder.

    DİKKAT: Bu yalnızca KAYDETMEK içindir. Var olan bir şablonu açarken/silerken
    adı yeniden buradan geçirmeyin — bkz. _sablon_yolu.
    """
    ad = " ".join(str(ad or "").split())[:60]
    for ch in '\\/:*?"<>|':
        ad = ad.replace(ch, "-")
    while ".." in ad:                  # '..' kalıntıları dosya adını çirkinleştirir
        ad = ad.replace("..", ".")
    return ad.strip(" .-") or "sablon"


def _sablon_yolu(name):
    """Listede görünen bir şablon adını GERÇEK dosya yoluna çevirir.

    Adı _guvenli_ad'dan geçirmeyiz. Sebep: klasöre dışarıdan kopyalanmış
    (ya da eski bir sürümden kalmış) dosyaların adı _guvenli_ad'ın üreteceği
    biçimde olmayabilir. Ad yeniden temizlenirse başka bir dosya adı hesaplanır,
    o dosya bulunamaz ve şablon arayüzden ASLA silinemez/açılamaz hale gelir.

    Bunun yerine ad, klasörün gerçek içeriğiyle eşleştirilir: listede ne
    görünüyorsa o silinebilir. Eşleşme yoksa None döner.
    """
    klasor = _sablon_klasoru()
    ham = str(name or "")
    istenen = ham.strip()
    if not istenen:
        return None
    try:
        dosyalar = [f for f in os.listdir(klasor) if f.lower().endswith(".html")]
    except OSError:
        return None

    # Eşleştirme kademeli: önce birebir, sonra boşluk/harf farkı hoş görülür.
    # (Windows dosya adlarının SONUNDAKİ boşlukları kendiliğinden atar; bu
    #  yüzden diskteki ad, kaydedilen adla birebir aynı olmayabilir.)
    kokler = [(f, os.path.splitext(f)[0]) for f in dosyalar]
    eslesme = None
    for aday in (
        lambda k: k == ham,
        lambda k: k == istenen,
        lambda k: k.strip() == istenen,
        lambda k: k.strip().lower() == istenen.lower(),
    ):
        for f, kok in kokler:
            if aday(kok):
                eslesme = f
                break
        if eslesme:
            break
    if eslesme is None:
        return None

    # Güvenlik ağı: çözülen yol gerçekten şablon klasörünün içinde mi?
    yol = os.path.join(klasor, eslesme)
    if os.path.dirname(os.path.realpath(yol)) != os.path.realpath(klasor):
        return None
    return yol


# ----------------------------------------------------------------- günlük
LOG_PATH = os.path.join(_veri_klasoru(), "calisma.log")
_log_lock = threading.Lock()


def kayit(mesaj: str) -> None:
    """Tanılama günlüğüne satır yazar. ASLA hata fırlatmaz.

    Paketlenmiş exe'de konsol yoktur; bir şey takılırsa tek kanıt bu dosyadır.
    """
    try:
        satir = f"{time.strftime('%Y-%m-%d %H:%M:%S')} [{threading.current_thread().name}] {mesaj}\n"
        with _log_lock:
            with open(LOG_PATH, "a", encoding="utf-8") as fh:
                fh.write(satir)
    except Exception:  # noqa: BLE001
        pass


class _GunlukAkisi:
    """sys.stdout/stderr yerine geçer.

    KRİTİK: PyInstaller'ın penceresel (console=False) derlemesinde sys.stderr
    None'dır; bir traceback.print_exc() çağrısı 'NoneType has no write' hatası
    verip iş parçacığını öldürür ve arayüz sonsuza kadar asılı kalır.
    """

    def write(self, s):
        if s and s.strip():
            kayit("cikti: " + s.rstrip())
        return len(s or "")

    def flush(self):
        pass

    def isatty(self):
        return False


def _gunluk_hazirla():
    if sys.stderr is None:
        sys.stderr = _GunlukAkisi()
    if sys.stdout is None:
        sys.stdout = _GunlukAkisi()
    # Günlük şişmesin: 1 MB üstünü baştan al.
    try:
        if os.path.isfile(LOG_PATH) and os.path.getsize(LOG_PATH) > 1_000_000:
            os.replace(LOG_PATH, LOG_PATH + ".eski")
    except OSError:
        pass


def _obfuscate(text: str) -> str:
    """Şifreyi düz metin görünmesin diye base64 ile gizler (gerçek şifreleme DEĞİL)."""
    return base64.b64encode((text or "").encode("utf-8")).decode("ascii")


def _deobfuscate(text: str) -> str:
    try:
        return base64.b64decode((text or "").encode("ascii")).decode("utf-8")
    except Exception:  # noqa: BLE001
        return ""


class _izle:
    """Bir köprü çağrısının süresini ölçer; eşiği aşarsa günlüğe yazar.

    Dekoratör KULLANILMIYOR: pywebview, JS'ten gelen parametreleri metodun
    imzasına bakarak eşliyor; sarmalanmış bir fonksiyon imzayı bozar.
    """

    __slots__ = ("ad", "t0")

    def __init__(self, ad):
        self.ad = ad

    def __enter__(self):
        self.t0 = time.monotonic()
        return self

    def __exit__(self, exc_type, exc, tb):
        dt = time.monotonic() - self.t0
        if exc is not None:
            kayit(f"HATA {self.ad} ({dt:.2f} sn): {exc}\n{traceback.format_exc()}")
        elif dt >= YAVAS_ESIK:
            kayit(f"YAVAŞ {self.ad}: {dt:.2f} sn")
        return False


def _write_problem_csv(job, rep) -> str:
    """Ön kontrolde bulunan sorunları Excel'in yanına 'kontrol_sorunlari.csv' olarak yazar.

    Arayüz 50 satırdan fazlası için "tam liste CSV'de" dediğinden bu dosya gereklidir.
    Dönen değer: yazılan dosyanın yolu (yazılamadıysa boş metin).
    """
    problems = (rep.get("bad_email", []) + rep.get("empty_attach", [])
                + rep.get("missing_attach", []) + rep.get("duplicate_attach", [])
                + rep.get("empty_field", []))
    if not problems:
        return ""
    out = os.path.join(os.path.dirname(job.xlsx_path) or ".", "kontrol_sorunlari.csv")
    try:
        with open(out, "w", encoding="utf-8-sig", newline="") as fh:
            w = csv.writer(fh)
            w.writerow(["satir", "email", "sorun_tipi", "detay"])
            w.writerows(problems)
    except OSError:
        return ""
    return out


class Api:
    """Arayüzün (JS) çağırdığı Python metotları. core.py'ye köprü kurar."""

    def __init__(self):
        # KRİTİK: alt çizgi ile başlamalı. pywebview köprüyü kurarken js_api
        # nesnesinin ALTI ÇİZGİSİZ tüm niteliklerini dolaşır ve nesne olanların
        # içine iner (webview/util.py:inject_pywebview). Bu nitelik "window"
        # olsaydı pywebview Window nesnesinin içine inip width/height/x/y
        # özelliklerini okur; bunların her biri events.shown.wait(15) yapar —
        # köprü 60 sn'ye kadar gecikir, arayüz 25 sn'de vazgeçip
        # "Python köprüsü kurulamadı" der. Ayrıca window.destroy /
        # window.evaluate_js gibi metotlar da JS'e açılmış olurdu.
        self._pencere = None
        self._stop = threading.Event()
        self._worker = None

        # ---- olay kuyruğu (bkz. modül başlığı) ----
        self._lock = threading.Lock()
        self._logs = deque(maxlen=LOG_BUFFER)   # [(level, msg), ...]
        self._progress = None                   # [done, total, ok, fail]  (sıkıştırılmış)
        self._scan = None                       # [done, total]            (sıkıştırılmış)
        self._finals = []                       # [{"t": "done"|"check_done"|"error", ...}]
        self._dropped = 0

    # ---------------------------------------------------------- ayarlar
    def hazir(self):
        """Arayüz köprünün kurulduğunu bu çağrıyla doğrular (ilk temas).

        Hangi motorun kullanıldığını da buradan günlüğe yazarız: 'edgechromium'
        beklenen değerdir. 'mshtml' görünüyorsa WebView2 kurulu değil demektir ve
        arayüz o motorda düzgün çalışmaz.
        """
        motor = getattr(getattr(webview, "guilib", None), "renderer", "?")
        kayit(f"köprü kuruldu (motor={motor})")
        return {"ok": True, "motor": motor}

    def ayarlari_yukle(self):
        with _izle("ayarlari_yukle"):
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
        kayit("dosya seçme penceresi açılıyor")
        res = self._pencere.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=("Excel (*.xlsx;*.xlsm)", "Tüm dosyalar (*.*)"),
        )
        kayit(f"dosya seçme penceresi kapandı: {'seçildi' if res else 'iptal'}")
        if not res:
            return ""
        return res[0] if isinstance(res, (list, tuple)) else res

    def sutunlari_oku(self, path, sheet):
        # DİKKAT: bu çağrı ağ/OneDrive yolunda uzun sürebilir (os.path.isfile bile
        # kopuk bir ağ sürücüsünde onlarca saniye bekleyebilir). Arka plan
        # iş parçacığında çalışır; arayüz bu sırada "Excel okunuyor..." gösterir.
        with _izle("sutunlari_oku"):
            if not path:
                return {"error": "Dosya seçilmedi."}
            if not os.path.isfile(path):
                return {"error": "Dosya bulunamadı (yol geçersiz ya da ağ sürücüsü kapalı olabilir)."}
            try:
                sheets, hdrs = core.get_headers(path, sheet or None)
            except Exception as exc:  # noqa: BLE001
                kayit(f"sutunlari_oku hata: {exc}")
                return {"error": str(exc)}
            return {"sheets": sheets, "headers": hdrs, "sheet": sheet or (sheets[0] if sheets else "")}

    def html_yukle(self):
        res = self._pencere.create_file_dialog(
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

    def resim_yukle(self):
        """Resim seçtirir ve gövdeye gömülecek data: URI olarak döndürür.

        Gönderim sırasında bu data: URI'ler otomatik olarak cid: (gömülü parça)
        hâline gelir — bkz. mailer.extract_inline_images.
        """
        res = self._pencere.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=("Resim (*.png;*.jpg;*.jpeg;*.gif;*.webp)", "Tüm dosyalar (*.*)"),
        )
        if not res:
            return {}
        p = res[0] if isinstance(res, (list, tuple)) else res
        try:
            with open(p, "rb") as fh:
                data = fh.read()
        except OSError as exc:
            return {"error": f"Resim okunamadı: {exc}"}

        import mimetypes
        tur = mimetypes.guess_type(p)[0] or ""
        if not tur.startswith("image/"):
            tur = "image/png"
        return {
            "uri": f"data:{tur};base64," + base64.b64encode(data).decode("ascii"),
            "name": os.path.basename(p),
            "size": len(data),
        }

    def html_kaydet(self, content, name="mail-sablonu.html"):
        """İçeriği kullanıcının seçtiği bir .html dosyasına yazar."""
        try:
            res = self._pencere.create_file_dialog(
                webview.SAVE_DIALOG,
                save_filename=_guvenli_ad(name) or "mail-sablonu.html",
                file_types=("HTML (*.html)", "Tüm dosyalar (*.*)"),
            )
        except Exception as exc:  # noqa: BLE001
            return {"error": f"Kaydetme penceresi açılamadı: {exc}"}
        if not res:
            return {}
        p = res[0] if isinstance(res, (list, tuple)) else res
        if not os.path.splitext(p)[1]:
            p += ".html"
        try:
            with open(p, "w", encoding="utf-8") as fh:
                fh.write(content or "")
        except OSError as exc:
            return {"error": f"Dosya yazılamadı: {exc}"}
        return {"path": p}

    # ------------------------------------------------------ yer tutucu alanlar
    def alanlar(self, job):
        """Editörün 'Alan Ekle' menüsünü besler.

        Döner: {columns: [...], builtins: [[ad, açıklama], ...],
                formats: [[ad, açıklama], ...], error: ""}
        """
        with _izle("alanlar"):
            path = (job or {}).get("xlsx") or ""
            columns, hata = [], ""
            if path and os.path.isfile(path):
                try:
                    _sheets, headers = core.get_headers(path, (job or {}).get("sheet") or None)
                    # Boş başlıklar ve tekrar edenler menüde işe yaramaz.
                    gorulen = set()
                    for h in headers:
                        ad = str(h or "").strip()
                        k = merge.normalize_name(ad)
                        if not k or k in gorulen:
                            continue
                        gorulen.add(k)
                        columns.append(ad)
                except Exception as exc:  # noqa: BLE001
                    hata = str(exc)
            return {
                "columns": columns,
                "builtins": [list(b) for b in merge.BUILTINS],
                "formats": [list(f) for f in merge.FORMATS],
                "error": hata,
            }

    def ornek_satirlar(self, job, adet=25):
        """Önizleme için ilk satırların okunabilir değerleri.

        Editör bunlarla alan uçlarında 'Ad Soyad → Ahmet Yılmaz' ipucu gösterir;
        gerçek üretim yine Python'da (onizleme) yapılır ki gönderilenle birebir
        aynı olsun.
        """
        with _izle("ornek_satirlar"):
            sendjob = self._build_job(job)
            if not os.path.isfile(sendjob.xlsx_path):
                return {"rows": [], "total": 0, "columns": []}
            try:
                recipients = core.read_recipients(
                    sendjob.xlsx_path, sendjob.sheet, sendjob.email_col, sendjob.attach_col
                )
                headers = core.read_header_names(sendjob.xlsx_path, sendjob.sheet)
            except Exception as exc:  # noqa: BLE001
                return {"error": str(exc), "rows": [], "total": 0, "columns": []}

            mapper = merge.RowMapper(headers)
            rows = []
            for r in recipients[: max(1, int(adet or 1))]:
                degerler = mapper.values(r.cells)
                rows.append({
                    "row": r.row,
                    "email": r.email,
                    "attachment": r.attachment,
                    "values": {ad: merge.cell_to_text(degerler.get(ad)) for ad in mapper.keys},
                })
            return {"rows": rows, "total": len(recipients), "columns": mapper.names}

    def onizleme(self, job, index=0):
        """Seçilen satırın verisiyle konu + gövdeyi ÜRETİLMİŞ hâlde döndürür.

        Üretim gönderimle AYNI kodu kullanır (merge.Template); önizlemede
        gördüğünüz metin, gidecek mailin metnidir.
        """
        with _izle("onizleme"):
            sendjob = self._build_job(job)
            headers, recipients = [], []
            if os.path.isfile(sendjob.xlsx_path):
                try:
                    headers = core.read_header_names(sendjob.xlsx_path, sendjob.sheet)
                    recipients = core.read_recipients(
                        sendjob.xlsx_path, sendjob.sheet, sendjob.email_col, sendjob.attach_col
                    )
                except Exception:  # noqa: BLE001 - veri yoksa da önizleme açılsın
                    recipients = []

            mapper = merge.RowMapper(headers)
            subject_tpl = merge.Template(sendjob.subject, is_html=False)
            body_tpl = merge.Template(sendjob.body, is_html=sendjob.is_html)
            known = mapper.keys | merge.BUILTIN_KEYS

            toplam = len(recipients)
            i = max(0, min(int(index or 0), toplam - 1)) if toplam else 0
            if toplam:
                rcp = recipients[i]
                values = merge.builtin_values(
                    row=rcp.row, email=rcp.email,
                    attachment=rcp.attachment, sender=sendjob.sender,
                )
                values.update(mapper.values(rcp.cells))
                satir, eposta, ek = rcp.row, rcp.email, rcp.attachment
            else:
                # Excel yoksa alanların yerine adlarını gösteririz; editör yine çalışır.
                values = merge.builtin_values(row="?", email="ornek@musteri.com", sender=sendjob.sender)
                values.update({k: f"«{k}»" for k in mapper.keys})
                satir, eposta, ek = "?", "ornek@musteri.com", ""

            return {
                "subject": subject_tpl.render(values),
                "body": body_tpl.render(values),
                "is_html": sendjob.is_html,
                "index": i,
                "total": toplam,
                "row": satir,
                "email": eposta,
                "attachment": ek,
                "attachment_name": os.path.basename(ek) if ek else "",
                "attachment_ok": bool(ek) and os.path.isfile(ek),
                "unknown": list(dict.fromkeys(
                    subject_tpl.unknown_names(known) + body_tpl.unknown_names(known)
                )),
                "used": list(dict.fromkeys(
                    a for a in (subject_tpl.used_names() + body_tpl.used_names())
                    if merge.normalize_name(a) in known
                )),
            }

    # ----------------------------------------------------------------- şablonlar
    def sablon_listesi(self):
        """Kaydedilmiş şablon adları (yeniden eskiye)."""
        klasor = _sablon_klasoru()
        try:
            dosyalar = [f for f in os.listdir(klasor) if f.lower().endswith(".html")]
        except OSError:
            return {"items": []}
        items = []
        for f in dosyalar:
            tam = os.path.join(klasor, f)
            try:
                st = os.stat(tam)
            except OSError:
                continue
            items.append({"name": os.path.splitext(f)[0], "time": st.st_mtime, "size": st.st_size})
        items.sort(key=lambda d: d["time"], reverse=True)
        return {"items": items}

    def sablon_kaydet(self, name, content):
        ad = _guvenli_ad(name)
        yol = os.path.join(_sablon_klasoru(), ad + ".html")
        try:
            with open(yol, "w", encoding="utf-8") as fh:
                fh.write(content or "")
        except OSError as exc:
            return {"error": f"Şablon kaydedilemedi: {exc}"}
        return {"name": ad}

    def sablon_yukle(self, name):
        yol = _sablon_yolu(name)
        if yol is None:
            return {"error": f"'{name}' adlı şablon bulunamadı."}
        try:
            with open(yol, "r", encoding="utf-8") as fh:
                return {"content": fh.read()}
        except OSError as exc:
            return {"error": f"Şablon okunamadı: {exc}"}

    def sablon_sil(self, name):
        yol = _sablon_yolu(name)
        if yol is None:
            return {"error": f"'{name}' adlı şablon bulunamadı."}
        try:
            os.remove(yol)
        except OSError as exc:
            return {"error": f"Şablon silinemedi: {exc}"}
        kayit(f"şablon silindi: {os.path.basename(yol)}")
        return {"ok": True}

    # ---------------------------------------------------------- iş kurma
    def _build_job(self, job):
        xlsx = (job.get("xlsx") or "").strip()
        results_path = os.path.join(os.path.dirname(xlsx) or ".", "gonderim_sonuclari.csv")
        return core.SendJob(
            xlsx_path=xlsx,
            sheet=job.get("sheet", ""),
            email_col=max(0, _sutun(job.get("email_col"), 0)),
            # Ek sütunu İSTEĞE BAĞLI: seçilmemişse (boş/None/-1) ek eklenmez.
            # DİKKAT: burada 'or 0' kullanılamaz — seçim yapılmamışken sessizce
            # A sütununu ek yolu sanıp her maile yanlış dosya iliştirirdi.
            attach_col=_sutun(job.get("attach_col"), core.NO_ATTACH_COL),
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
        # Yalnızca e-posta sütunu zorunludur; ek sütunu seçilmeyebilir.
        if _sutun(job.get("email_col"), -1) < 0:
            return ["Lütfen e-posta sütununu seçin."]
        return core.validate_job(self._build_job(job))

    # ---------------------------------------------------------- olay kuyruğu
    def _log(self, level, msg):
        with self._lock:
            if len(self._logs) == LOG_BUFFER:
                self._dropped += 1
            self._logs.append((level, msg))

    def _set_progress(self, done, total, ok, fail):
        # Sıkıştırma: yalnızca en güncel değer tutulur, ara değerler arayüze taşınmaz.
        with self._lock:
            self._progress = [done, total, ok, fail]

    def _set_scan(self, done, total):
        with self._lock:
            self._scan = [done, total]

    def _final(self, payload):
        with self._lock:
            self._finals.append(payload)

    def _reset_events(self):
        with self._lock:
            self._logs.clear()
            self._progress = None
            self._scan = None
            self._finals = []
            self._dropped = 0

    def olaylari_al(self):
        """Arayüz düzenli aralıkla çağırır: biriken olayları TEK paket olarak döndürür.

        Dönen paket JSON'a çevrilip köprüden geçtiği için olabildiğince küçük tutulur
        (sözlük yerine dizi, ara ilerleme değerleri atılmış).
        """
        with self._lock:
            logs = list(self._logs)
            self._logs.clear()
            progress, self._progress = self._progress, None
            scan, self._scan = self._scan, None
            finals, self._finals = self._finals, []
            dropped, self._dropped = self._dropped, 0
        return {
            "logs": logs,
            "progress": progress,
            "scan": scan,
            "finals": finals,
            "dropped": dropped,
        }

    # ---------------------------------------------------------- ortak worker
    def mesgul(self):
        return bool(self._worker and self._worker.is_alive())

    def _start_worker(self, name, target):
        """İşi arka planda başlatır; ne olursa olsun bir 'final' olayı üretir."""
        if self.mesgul():
            return {"error": "Şu anda başka bir işlem sürüyor."}
        self._stop.clear()
        self._reset_events()

        def run():
            kayit(f"{name}: başladı")
            t0 = time.monotonic()
            try:
                target()
            except BaseException as exc:  # noqa: BLE001 - arayüz asla asılı kalmasın
                # print_exc() KULLANMIYORUZ: penceresel exe'de sys.stderr None olabilir
                # ve bu satırın kendisi patlayıp işi sonuçsuz bırakır.
                kayit(f"{name}: HATA {exc}\n{traceback.format_exc()}")
                self._log("error", f"Beklenmeyen hata: {exc}")
                self._final({"t": "error", "msg": str(exc)})
            finally:
                kayit(f"{name}: bitti ({time.monotonic() - t0:.1f} sn)")

        self._worker = threading.Thread(target=run, name=name, daemon=True)
        self._worker.start()
        return {"started": True}

    # ---------------------------------------------------------- ön kontrol
    def kontrol_et(self, job):
        """Ön kontrolü arka planda başlatır (ilerleme ve Durdur destekli)."""
        sendjob = self._build_job(job)
        if not os.path.isfile(sendjob.xlsx_path):
            return {"error": "Önce geçerli bir Excel dosyası seçin."}

        def work():
            rep = core.check_job(
                sendjob,
                on_progress=self._set_scan,
                should_stop=self._stop.is_set,
            )
            rep["csv"] = _write_problem_csv(sendjob, rep)
            self._final({"t": "check_done", "report": rep})

        return self._start_worker("kontrol", work)

    # ---------------------------------------------------------- gönderim
    def gonder(self, job):
        sendjob = self._build_job(job)
        problems = core.validate_job(sendjob)
        if problems:
            return {"error": "Eksik/hatalı ayarlar: " + ", ".join(problems)}

        def work():
            result = core.run_job(
                sendjob,
                on_progress=self._set_progress,
                on_log=self._log,
                stop_event=self._stop,
            )
            self._final({"t": "done", "result": result})

        return self._start_worker("gonderim", work)

    def durdur(self):
        """Hem gönderimi hem ön kontrolü durdurur."""
        self._stop.set()


# --------------------------------------------------------------- açılış
def _mesaj_kutusu(baslik, metin):
    """Konsolsuz exe'de kullanıcıya bir şey söyleyebilmenin tek yolu."""
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, metin, baslik, 0x40)   # 0x40 = bilgi
    except Exception:  # noqa: BLE001
        pass


def _webview2_kurulu_mu():
    """Edge WebView2 Runtime kurulu mu?

    Kurulu değilse pywebview eski IE (MSHTML) motoruna düşer; arayüz modern
    JS/CSS kullandığı için orada çalışmaz. Bunu sessizce yaşamak yerine söylüyoruz.
    """
    if os.name != "nt":
        return True
    try:
        import winreg
    except ImportError:
        return True
    anahtar = r"SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
    yerler = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"),
        (winreg.HKEY_LOCAL_MACHINE, anahtar),
        (winreg.HKEY_CURRENT_USER, anahtar),
    ]
    for kok, yol in yerler:
        try:
            with winreg.OpenKey(kok, yol) as k:
                surum, _ = winreg.QueryValueEx(k, "pv")
                if surum and surum != "0.0.0.0":
                    return True
        except OSError:
            continue
    return False


def main():
    _gunluk_hazirla()
    tanilama = os.environ.get("TFM_TANILAMA") == "1" or "--tanilama" in sys.argv

    kayit("=" * 60)
    kayit(f"AÇILIŞ | donmuş(exe)={getattr(sys, 'frozen', False)} | python={sys.version.split()[0]}")
    kayit(f"çalışma yolu={os.path.abspath(sys.argv[0])}")
    kayit(f"web klasörü={_app_web_dir()}")
    try:
        kayit(f"pywebview={getattr(webview, '__version__', '?')}")
    except Exception:  # noqa: BLE001
        pass

    if tanilama:
        # Uygulama takılırsa 20 sn'de bir TÜM iş parçacıklarının yığınını günlüğe
        # döker; "nerede asılı kaldı" sorusunun kesin cevabı budur.
        try:
            import faulthandler
            _fh = open(LOG_PATH, "a", encoding="utf-8", buffering=1)
            faulthandler.enable(file=_fh)
            faulthandler.dump_traceback_later(20, repeat=True, exit=False, file=_fh)
            kayit("TANILAMA açık: 20 sn'de bir iş parçacığı yığınları dökülecek")
        except Exception as exc:  # noqa: BLE001
            kayit(f"faulthandler kurulamadı: {exc}")

    if not _webview2_kurulu_mu():
        kayit("UYARI: Edge WebView2 Runtime bulunamadı")
        _mesaj_kutusu(
            APP_TITLE,
            "Microsoft Edge WebView2 Runtime bulunamadı.\n\n"
            "Arayüz bu bileşenle çizilir; kurulu değilse pencere boş kalabilir "
            "veya yanıt vermeyebilir.\n\n"
            "Kurulum: https://developer.microsoft.com/microsoft-edge/webview2/\n"
            "(Evergreen Standalone Installer)",
        )

    api = Api()
    index = os.path.join(_app_web_dir(), "index.html")
    if not os.path.isfile(index):
        kayit(f"HATA: arayüz dosyası yok: {index}")
        _mesaj_kutusu(APP_TITLE, f"Arayüz dosyası bulunamadı:\n{index}")
        return

    kayit("pencere oluşturuluyor")
    window = webview.create_window(
        APP_TITLE,
        url=index,
        js_api=api,
        width=940,
        height=880,
        min_size=(820, 680),
        background_color="#f9f9fb",
    )
    api._pencere = window

    # WebView2'nin profil klasörü: exe'nin yanına değil, kullanıcı klasörüne.
    # (Exe ağ sürücüsünden/OneDrive'dan çalıştırıldığında profil yazamayıp
    #  açılışta kilitlenebiliyor.)
    depo = os.path.join(
        os.environ.get("LOCALAPPDATA") or _veri_klasoru(), "TopluFaturaMailer", "webview"
    )
    try:
        os.makedirs(depo, exist_ok=True)
    except OSError:
        depo = None

    kayit(f"webview.start() çağrılıyor (depo={depo}, tanilama={tanilama})")
    t0 = time.monotonic()
    try:
        if depo:
            webview.start(debug=tanilama, private_mode=False, storage_path=depo)
        else:
            webview.start(debug=tanilama)
    except TypeError:
        # Eski pywebview sürümleri private_mode/storage_path bilmez.
        kayit("start(): eski pywebview imzası, sade çağrıya düşülüyor")
        webview.start()
    except Exception as exc:  # noqa: BLE001
        kayit(f"webview.start() HATA: {exc}\n{traceback.format_exc()}")
        _mesaj_kutusu(APP_TITLE, f"Arayüz başlatılamadı:\n{exc}\n\nAyrıntı: {LOG_PATH}")
        return
    kayit(f"webview.start() döndü — pencere kapandı ({time.monotonic() - t0:.1f} sn açık kaldı)")


if __name__ == "__main__":
    try:
        main()
    except BaseException as exc:  # noqa: BLE001
        kayit(f"ÖLÜMCÜL HATA: {exc}\n{traceback.format_exc()}")
        _mesaj_kutusu(APP_TITLE, f"Uygulama başlatılamadı:\n{exc}\n\nAyrıntı: {LOG_PATH}")
        raise
