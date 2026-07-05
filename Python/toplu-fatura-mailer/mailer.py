"""
Gönderim arka planı (backend) modülü.

İki gönderim yöntemi sağlar:
  - SmtpMailer   : Standart SMTP sunucusu üzerinden (Outlook kurulu olmasa da çalışır).
  - OutlookMailer: Bilgisayarda kurulu Masaüstü Outlook (win32com) üzerinden.

Her iki sınıf da aynı arayüze sahiptir:
    mailer.connect()                       -> bağlantı/hazırlık (gerekirse)
    mailer.send(to, subject, body, attachment, is_html) -> tek bir mail gönderir
    mailer.close()                         -> kaynakları serbest bırakır

Böylece GUI tarafı hangi yöntemin seçildiğini bilmeden aynı şekilde kullanabilir.
"""

from __future__ import annotations

import os
import smtplib
import mimetypes
from email.message import EmailMessage


class MailerError(Exception):
    """Gönderimle ilgili beklenen hataları tek tipte iletmek için."""


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

    def __init__(self, host, port, security, username, password, sender):
        self.host = host
        self.port = int(port)
        self.security = (security or "starttls").lower()
        self.username = username
        self.password = password
        self.sender = sender or username
        self._server = None

    def connect(self):
        try:
            if self.security == "ssl":
                self._server = smtplib.SMTP_SSL(self.host, self.port, timeout=30)
            else:
                self._server = smtplib.SMTP(self.host, self.port, timeout=30)
                self._server.ehlo()
                if self.security == "starttls":
                    self._server.starttls()
                    self._server.ehlo()
            if self.username:
                self._server.login(self.username, self.password)
        except Exception as exc:  # noqa: BLE001 - kullanıcıya anlaşılır hata
            raise MailerError(f"SMTP bağlantısı kurulamadı: {exc}") from exc

    def _ensure(self):
        # Sunucu bağlantısı düşmüşse yeniden bağlan.
        if self._server is None:
            self.connect()
            return
        try:
            status = self._server.noop()[0]
        except Exception:  # noqa: BLE001
            status = -1
        if status != 250:
            self.connect()

    def send(self, to, subject, body, attachment=None, is_html=False):
        self._ensure()
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
            _attach_file(msg, attachment)

        try:
            self._server.send_message(msg)
        except Exception as exc:  # noqa: BLE001
            raise MailerError(str(exc)) from exc

    def close(self):
        if self._server is not None:
            try:
                self._server.quit()
            except Exception:  # noqa: BLE001
                pass
            self._server = None


def _attach_file(msg: EmailMessage, path: str):
    if not os.path.isfile(path):
        raise MailerError(f"Ek dosyası bulunamadı: {path}")
    ctype, encoding = mimetypes.guess_type(path)
    if ctype is None or encoding is not None:
        ctype = "application/octet-stream"
    maintype, subtype = ctype.split("/", 1)
    with open(path, "rb") as fh:
        data = fh.read()
    msg.add_attachment(
        data,
        maintype=maintype,
        subtype=subtype,
        filename=os.path.basename(path),
    )


# ---------------------------------------------------------------------------
# Outlook (win32com)
# ---------------------------------------------------------------------------
class OutlookMailer:
    """Masaüstü Outlook üzerinden gönderim yapar (pywin32 gerekir).

    sender: Gönderimde kullanılacak hesabın SMTP adresi (örn. muhasebe@atakonline.com).
            Outlook'ta birden fazla hesap varsa doğru hesap seçilir. Boş bırakılırsa
            Outlook'un varsayılan hesabı kullanılır.
    """

    def __init__(self, sender=None):
        self.sender = sender
        self._outlook = None
        self._account = None

    def connect(self):
        try:
            import win32com.client  # yerel import: SMTP kullanıcısı pywin32'ye ihtiyaç duymasın
        except ImportError as exc:
            raise MailerError(
                "Outlook yöntemi için 'pywin32' gerekli. Kurulum: pip install pywin32"
            ) from exc
        try:
            self._outlook = win32com.client.Dispatch("Outlook.Application")
        except Exception as exc:  # noqa: BLE001
            raise MailerError(
                "Outlook başlatılamadı. Masaüstü Outlook kurulu ve açık mı?"
            ) from exc

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

    def send(self, to, subject, body, attachment=None, is_html=False):
        if self._outlook is None:
            self.connect()
        try:
            mail = self._outlook.CreateItem(0)  # 0 = olMailItem
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
