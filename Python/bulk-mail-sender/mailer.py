"""
Gönderim arka planı (backend) modülü.

İki gönderim yöntemi sağlar:
  - SmtpMailer   : Standart SMTP sunucusu üzerinden (Outlook kurulu olmasa da çalışır).
  - OutlookMailer: Bilgisayarda kurulu Masaüstü Outlook (win32com) üzerinden.

Her iki sınıf da aynı arayüze sahiptir:
    mailer.connect()                       -> bağlantı/hazırlık (gerekirse)
    mailer.prefetch(path)                  -> eki önden okur (isteğe bağlı hızlandırma)
    mailer.send(to, subject, body, attachment, is_html) -> tek bir mail gönderir
    mailer.close()                         -> kaynakları serbest bırakır

Böylece GUI tarafı hangi yöntemin seçildiğini bilmeden aynı şekilde kullanabilir.

Performans notları
------------------
* SMTP'de her mailden önce NOOP atılmaz. NOOP tam bir sunucu gidiş-dönüşüdür ve
  1300 maillik bir listede 1300 gereksiz tur demektir. Bunun yerine bağlantı
  KEEPALIVE saniyeden uzun süredir boştaysa kontrol edilir (bkz. _ensure).
* Ek dosyaları AttachmentCache ile bellekte tutulur; aynı dosya birden çok satırda
  geçiyorsa (veya core.py önden okuduysa) diskten/ağdan tekrar okunmaz.
* Outlook, arka plan iş parçacığında çalıştığı için COM açıkça başlatılır
  (pythoncom.CoInitialize); aksi halde "CoInitialize has not been called" hatası olur.
"""

from __future__ import annotations

import os
import time
import smtplib
import mimetypes
import threading
from collections import OrderedDict
from email.message import EmailMessage


class MailerError(Exception):
    """Gönderimle ilgili beklenen hataları tek tipte iletmek için."""


# ---------------------------------------------------------------------------
# Ek dosyası önbelleği
# ---------------------------------------------------------------------------
class AttachmentCache:
    """Ek dosyalarının içeriğini sınırlı boyutta (LRU) bellekte tutar.

    Amaç: aynı dosyanın birden çok satırda kullanılması ve core.py'nin ekleri
    önden okuması (prefetch) durumunda ağ sürücüsünden tekrar tekrar okumamak.

    Sınırlar aşıldığında en eski girdiler atılır; tek başına 'max_bytes'tan büyük
    dosyalar hiç önbelleğe alınmaz (bellek şişmesin diye doğrudan okunur).
    """

    def __init__(self, max_bytes: int = 64 * 1024 * 1024, max_items: int = 128):
        self.max_bytes = max_bytes
        self.max_items = max_items
        self._items = OrderedDict()   # key -> (data, maintype, subtype, filename)
        self._bytes = 0
        self._lock = threading.Lock()

    @staticmethod
    def _key(path):
        """Dosya kimliği: yol + değişiklik zamanı + boyut (dosya değişirse anahtar değişir)."""
        st = os.stat(path)
        return (os.path.normcase(os.path.abspath(path)), st.st_mtime_ns, st.st_size)

    def load(self, path):
        """(veri, maintype, subtype, dosya_adı) döndürür. Önbellekte yoksa okur."""
        if not path:
            raise MailerError("Ek dosyası yolu boş.")
        try:
            key = self._key(path)
        except OSError as exc:
            raise MailerError(f"Ek dosyası bulunamadı: {path}") from exc

        with self._lock:
            hit = self._items.get(key)
            if hit is not None:
                self._items.move_to_end(key)   # LRU: en son kullanılan sona
                return hit

        # Okuma kilit dışında yapılır; böylece bir dosya okunurken diğer
        # iş parçacıkları önbelleğe erişmeye devam edebilir.
        try:
            with open(path, "rb") as fh:
                data = fh.read()
        except OSError as exc:
            raise MailerError(f"Ek dosyası okunamadı: {path} ({exc})") from exc

        ctype, encoding = mimetypes.guess_type(path)
        if ctype is None or encoding is not None:
            ctype = "application/octet-stream"
        maintype, subtype = ctype.split("/", 1)
        item = (data, maintype, subtype, os.path.basename(path))

        size = len(data)
        if size <= self.max_bytes:
            with self._lock:
                if key not in self._items:
                    self._items[key] = item
                    self._bytes += size
                    self._evict()
        return item

    def _evict(self):
        """Sınırlar aşıldıysa en eski girdileri at (çağıran kilidi tutuyor olmalı)."""
        while self._items and (self._bytes > self.max_bytes or len(self._items) > self.max_items):
            _key, old = self._items.popitem(last=False)
            self._bytes -= len(old[0])

    def clear(self):
        with self._lock:
            self._items.clear()
            self._bytes = 0


def _attach_file(msg: EmailMessage, path: str, cache: AttachmentCache | None = None):
    if not os.path.isfile(path):
        raise MailerError(f"Ek dosyası bulunamadı: {path}")
    if cache is None:
        cache = AttachmentCache(max_bytes=0)   # önbelleksiz: sadece oku
    data, maintype, subtype, filename = cache.load(path)
    msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=filename)


# ---------------------------------------------------------------------------
# SMTP
# ---------------------------------------------------------------------------
class SmtpMailer:
    """Standart SMTP sunucusu ile gönderim yapar.

    security:
        'starttls' -> genellikle 587 portu (TLS ile şifreli)
        'ssl'      -> genellikle 465 portu (baştan SSL)
        'none'     -> şifresiz (önerilmez)
    """

    # Bağlantı bu süreden kısa süre önce başarıyla kullanıldıysa "canlı" sayılır
    # ve NOOP ile kontrol edilmez. Sunucuların boşta kalma zaman aşımı genelde
    # 60-300 sn olduğundan 15 sn güvenli bir pencere.
    KEEPALIVE = 15.0

    supports_prefetch = True

    def __init__(self, host, port, security, username, password, sender,
                 attachment_cache=None, timeout=30):
        self.host = host
        self.port = int(port)
        self.security = (security or "starttls").lower()
        self.username = username
        self.password = password
        self.sender = sender or username
        self.timeout = timeout
        self.attachments = attachment_cache or AttachmentCache()
        self._server = None
        self._last_ok = 0.0

    # ------------------------------------------------------------- bağlantı
    def connect(self):
        self._close_quietly()
        try:
            if self.security == "ssl":
                self._server = smtplib.SMTP_SSL(self.host, self.port, timeout=self.timeout)
            else:
                self._server = smtplib.SMTP(self.host, self.port, timeout=self.timeout)
                self._server.ehlo()
                if self.security == "starttls":
                    self._server.starttls()
                    self._server.ehlo()
            if self.username:
                self._server.login(self.username, self.password)
        except Exception as exc:  # noqa: BLE001 - kullanıcıya anlaşılır hata
            self._server = None
            raise MailerError(f"SMTP bağlantısı kurulamadı: {exc}") from exc
        self._last_ok = time.monotonic()

    def _ensure(self):
        """Bağlantının kullanılabilir olduğundan emin olur.

        Az önce başarıyla kullanılmış bir bağlantı için sunucuya soru sormayız;
        bu, mail başına bir gidiş-dönüşü ortadan kaldırır. Yalnızca bağlantı bir
        süredir boştaysa NOOP ile yoklanır.
        """
        if self._server is None:
            self.connect()
            return
        if (time.monotonic() - self._last_ok) < self.KEEPALIVE:
            return
        try:
            status = self._server.noop()[0]
        except Exception:  # noqa: BLE001
            status = -1
        if status != 250:
            self.connect()
        else:
            self._last_ok = time.monotonic()

    def prefetch(self, path):
        """Eki önden belleğe alır (core.py arka planda çağırır). Hata yutulur."""
        if not path:
            return
        try:
            self.attachments.load(path)
        except Exception:  # noqa: BLE001 - önden okuma başarısızsa gönderimde raporlanır
            pass

    # ------------------------------------------------------------- gönderim
    def _build(self, to, subject, body, attachment, is_html):
        msg = EmailMessage()
        msg["From"] = self.sender
        msg["To"] = to
        msg["Subject"] = subject
        if is_html:
            msg.set_content("Bu e-postayı görüntülemek için HTML destekli bir istemci kullanın.")
            msg.add_alternative(body, subtype="html")
        else:
            msg.set_content(body)
        if attachment:
            _attach_file(msg, attachment, self.attachments)
        return msg

    def send(self, to, subject, body, attachment=None, is_html=False):
        # Mesaj bağlantıdan önce hazırlanır: ek okunamıyorsa sunucuyu meşgul etmeyiz.
        msg = self._build(to, subject, body, attachment, is_html)
        self._ensure()
        try:
            self._server.send_message(msg)
        except smtplib.SMTPServerDisconnected as exc:
            # Bağlantı koptu. Mesajın sunucuya ulaşıp ulaşmadığı belli olmadığından
            # OTOMATİK TEKRAR GÖNDERMİYORUZ (mükerrer fatura riski). Bağlantıyı
            # düşürüyoruz; bir sonraki satır temiz bir bağlantıyla devam eder.
            self._server = None
            raise MailerError(f"Bağlantı koptu: {exc}") from exc
        except Exception as exc:  # noqa: BLE001
            raise MailerError(str(exc)) from exc
        self._last_ok = time.monotonic()

    # ------------------------------------------------------------- kapatma
    def _close_quietly(self):
        if self._server is not None:
            try:
                self._server.quit()
            except Exception:  # noqa: BLE001
                pass
            self._server = None

    def close(self):
        self._close_quietly()
        self.attachments.clear()


# ---------------------------------------------------------------------------
# Outlook (win32com)
# ---------------------------------------------------------------------------
class OutlookMailer:
    """Masaüstü Outlook üzerinden gönderim yapar (pywin32 gerekir).

    sender: Gönderimde kullanılacak hesabın SMTP adresi (örn. muhasebe@atakonline.com).
            Outlook'ta birden fazla hesap varsa doğru hesap seçilir. Boş bırakılırsa
            Outlook'un varsayılan hesabı kullanılır.

    Not: Ekler Outlook'a dosya YOLU olarak verilir; dosyayı Outlook kendisi okur.
    Bu yüzden ek içeriğini belleğe almanın (prefetch) faydası yoktur.
    """

    supports_prefetch = False

    def __init__(self, sender=None):
        self.sender = sender
        self._outlook = None
        self._account = None
        self._create_item = None
        self._com_ready = False

    def connect(self):
        try:
            import pythoncom  # noqa: PLC0415 - yerel import: SMTP kullanıcısı pywin32 istemesin
            import win32com.client
        except ImportError as exc:
            raise MailerError(
                "Outlook yöntemi için 'pywin32' gerekli. Kurulum: pip install pywin32"
            ) from exc

        # Gönderim arka plan iş parçacığında çalışır; COM'un bu iş parçacığı için
        # başlatılması şarttır, yoksa Dispatch "CoInitialize has not been called" der.
        if not self._com_ready:
            try:
                pythoncom.CoInitialize()
                self._com_ready = True
            except Exception:  # noqa: BLE001 - zaten başlatılmışsa sorun değil
                pass

        try:
            self._outlook = win32com.client.Dispatch("Outlook.Application")
        except Exception as exc:  # noqa: BLE001
            raise MailerError(
                "Outlook başlatılamadı. Masaüstü Outlook kurulu ve açık mı?"
            ) from exc

        # COM'da her nokta (.) bir arama demektir; sık kullanılanı bir kez çözüp saklıyoruz.
        self._create_item = self._outlook.CreateItem

        if self.sender:
            self._account = self._find_account(self.sender)
            if self._account is None:
                raise MailerError(
                    f"Outlook'ta '{self.sender}' hesabı bulunamadı. "
                    "Bu hesabın Outlook'a ekli olduğundan emin olun."
                )

    def _find_account(self, smtp_address):
        session = self._outlook.Session
        target = smtp_address.strip().lower()
        for acc in session.Accounts:
            try:
                if str(acc.SmtpAddress).strip().lower() == target:
                    return acc
            except Exception:  # noqa: BLE001
                continue
        return None

    def prefetch(self, path):
        """Outlook dosyayı kendisi okuduğu için önden okumaya gerek yok."""

    def send(self, to, subject, body, attachment=None, is_html=False):
        if self._outlook is None:
            self.connect()
        try:
            mail = self._create_item(0)  # 0 = olMailItem
            mail.To = to
            mail.Subject = subject
            if is_html:
                mail.HTMLBody = body
            else:
                mail.Body = body
            if attachment:
                if not os.path.isfile(attachment):
                    raise MailerError(f"Ek dosyası bulunamadı: {attachment}")
                mail.Attachments.Add(os.path.abspath(attachment))
            if self._account is not None:
                # Doğru hesaptan gönder.
                mail._oleobj_.Invoke(*(64209, 0, 8, 0, self._account))  # SendUsingAccount
            mail.Send()
        except MailerError:
            raise
        except Exception as exc:  # noqa: BLE001
            raise MailerError(str(exc)) from exc

    def close(self):
        # Outlook uygulaması kullanıcıya ait; kapatmıyoruz, sadece referansı bırakıyoruz.
        self._outlook = None
        self._account = None
        self._create_item = None
        if self._com_ready:
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:  # noqa: BLE001
                pass
            self._com_ready = False


def make_mailer(method, **kwargs):
    """Fabrika: 'smtp' veya 'outlook' için uygun mailer nesnesini döndürür."""
    method = (method or "").lower()
    if method == "smtp":
        return SmtpMailer(
            host=kwargs["host"],
            port=kwargs["port"],
            security=kwargs.get("security", "starttls"),
            username=kwargs.get("username", ""),
            password=kwargs.get("password", ""),
            sender=kwargs.get("sender", ""),
        )
    if method == "outlook":
        return OutlookMailer(sender=kwargs.get("sender") or None)
    raise MailerError(f"Bilinmeyen gönderim yöntemi: {method}")
