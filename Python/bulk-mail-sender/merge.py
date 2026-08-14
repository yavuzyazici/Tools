"""
Yer tutucu (mail merge) motoru — {Sütun} alanları.

Konu ve içerikteki `{Sütun Adı}` yazımları, o satırın Excel değeriyle değiştirilir.
Böylece her alıcıya kişiye özel bir mail gider.

Yazım
-----
    {Ad}                       -> sütunun değeri
    {Ad|Değerli Müşterimiz}    -> değer boşsa bu yazılır (varsayılan)
    {Tutar:para}               -> biçimlendirilmiş değer  (1.234,50)
    {Tarih:%d %B %Y}           -> '%' ile başlayan biçim strftime kalıbıdır
    {Tutar:para|0,00}          -> biçim + varsayılan birlikte

Kurallar
--------
* Sütun adı BÜYÜK/küçük harf ve baştaki/sondaki boşluklar dikkate alınmadan eşleşir
  ("Fatura No", "fatura no" ve "FATURA NO" aynı alandır). Türkçe I/İ/ı/i ayrımı da
  normalleştirilir.
* TANINMAYAN bir ad ASLA değiştirilmez, olduğu gibi kalır. Bu kural bilinçlidir:
  HTML içindeki CSS blokları (`p { margin:0 }`) de süslü parantez içerir; yanlışlıkla
  silinmemeleri gerekir.
* HTML gövdede `<style>` ve `<script>` blokları hiç taranmaz (aynı sebep).
* HTML gövdede değerler kaçışlanır (& < > " ' -> &amp; ...) ve satır sonları <br />
  olur; `:ham` biçimi bunu kapatır (değeri HTML olarak gömmek isteyenler için).

Bu modül GUI'den ve core.py'den bağımsızdır; tek başına test edilebilir.
"""

from __future__ import annotations

import re
import os
import datetime as _dt
from decimal import Decimal, InvalidOperation
from html import escape as _html_escape


# ---------------------------------------------------------------------------
# Türkçe büyük/küçük harf (I/İ/ı/i)
# ---------------------------------------------------------------------------
_TR_TO_UPPER = str.maketrans("ıi", "Iİ")
_TR_TO_LOWER = str.maketrans("Iİ", "ıi")


def tr_upper(s: str) -> str:
    return s.translate(_TR_TO_UPPER).upper()


def tr_lower(s: str) -> str:
    return s.translate(_TR_TO_LOWER).lower()


def tr_title(s: str) -> str:
    """Her kelimenin ilk harfi büyük ('ahmet yılmaz' -> 'Ahmet Yılmaz')."""
    out = []
    yeni_kelime = True
    for ch in s:
        if ch.isalpha():
            out.append(tr_upper(ch) if yeni_kelime else tr_lower(ch))
            yeni_kelime = False
        else:
            out.append(ch)
            yeni_kelime = True
    return "".join(out)


def normalize_name(name) -> str:
    """Alan adını eşleştirme anahtarına çevirir.

    Boşluklar sadeleşir, Türkçe I/İ/ı/i tek harfe iner, hepsi küçük harf olur.
    'Fatura  No' , 'FATURA NO' , 'fatura no' -> 'fatura no'
    """
    s = " ".join(str(name or "").split())
    return tr_lower(s).replace("ı", "i")


# ---------------------------------------------------------------------------
# Hücre değerini metne çevirme
# ---------------------------------------------------------------------------
_AY_ADLARI = [
    "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran",
    "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık",
]
_GUN_ADLARI = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]

# Metin hücrelerinde tarih arayan kalıplar (Excel bazen tarihi metin olarak tutar).
_TARIH_KALIPLARI = (
    "%d.%m.%Y", "%d/%m/%Y", "%d-%m-%Y",
    "%Y-%m-%d", "%Y/%m/%d",
    "%d.%m.%Y %H:%M", "%d.%m.%Y %H:%M:%S",
    "%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S",
    "%d.%m.%y", "%m/%d/%Y",
)


def _float_to_text(value: float) -> str:
    """1234.0 -> '1234', 1234.5 -> '1234.5' (bilimsel gösterime kaçmadan)."""
    if value != value or value in (float("inf"), float("-inf")):   # NaN / sonsuz
        return ""
    if float(value).is_integer() and abs(value) < 1e15:
        return str(int(value))
    if abs(value) >= 1e15 or (value != 0 and abs(value) < 1e-4):
        return repr(value)
    return f"{value:.10f}".rstrip("0").rstrip(".")


def cell_to_text(value) -> str:
    """Excel hücresini kullanıcıya gösterilecek düz metne çevirir."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "Evet" if value else "Hayır"
    if isinstance(value, _dt.datetime):
        if (value.hour, value.minute, value.second) == (0, 0, 0):
            return value.strftime("%d.%m.%Y")
        return value.strftime("%d.%m.%Y %H:%M")
    if isinstance(value, _dt.date):
        return value.strftime("%d.%m.%Y")
    if isinstance(value, _dt.time):
        return value.strftime("%H:%M")
    if isinstance(value, _dt.timedelta):
        toplam = int(value.total_seconds())
        return f"{toplam // 3600:02d}:{(toplam % 3600) // 60:02d}"
    if isinstance(value, float):
        return _float_to_text(value)
    if isinstance(value, Decimal):
        return _float_to_text(float(value))
    if isinstance(value, int):
        return str(value)
    return str(value).strip()


def _as_datetime(value):
    """Değeri tarihe çevirmeye çalışır; olmuyorsa None."""
    if isinstance(value, _dt.datetime):
        return value
    if isinstance(value, _dt.date):
        return _dt.datetime(value.year, value.month, value.day)
    if isinstance(value, str):
        metin = value.strip()
        if not metin:
            return None
        for kalip in _TARIH_KALIPLARI:
            try:
                return _dt.datetime.strptime(metin, kalip)
            except ValueError:
                continue
    return None


def _as_decimal(value):
    """Değeri sayıya çevirmeye çalışır ('1.234,50' gibi Türkçe yazım dahil)."""
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float, Decimal)):
        try:
            return Decimal(str(value))
        except InvalidOperation:
            return None
    if isinstance(value, str):
        metin = value.strip()
        if not metin:
            return None
        metin = re.sub(r"[^\d,.\-+]", "", metin)
        if not metin:
            return None
        # '1.234,50' -> '1234.50' ;  '1,234.50' -> '1234.50'
        if "," in metin and "." in metin:
            if metin.rfind(",") > metin.rfind("."):
                metin = metin.replace(".", "").replace(",", ".")
            else:
                metin = metin.replace(",", "")
        elif "," in metin:
            metin = metin.replace(",", ".")
        try:
            return Decimal(metin)
        except InvalidOperation:
            return None
    return None


def _group(metin: str) -> str:
    """'1234567.89' -> '1.234.567,89' (binlik nokta, ondalık virgül)."""
    eksi = metin.startswith("-")
    metin = metin.lstrip("+-")
    tam, _, ondalik = metin.partition(".")
    parcalar = []
    while len(tam) > 3:
        parcalar.insert(0, tam[-3:])
        tam = tam[:-3]
    parcalar.insert(0, tam or "0")
    sonuc = ".".join(parcalar)
    if ondalik:
        sonuc += "," + ondalik
    return ("-" if eksi else "") + sonuc


# ---------------------------------------------------------------------------
# Biçimler
# ---------------------------------------------------------------------------
#  ad -> (açıklama)   — arayüzdeki "biçim" menüsü de bu listeden beslenir.
FORMATS = [
    ("", "olduğu gibi"),
    ("buyuk", "BÜYÜK HARF"),
    ("kucuk", "küçük harf"),
    ("baslik", "Her Kelime Büyük"),
    ("para", "Para (1.234,50)"),
    ("sayi", "Sayı (1.234,5)"),
    ("tamsayi", "Tam sayı (1.235)"),
    ("tarih", "Tarih (01.02.2026)"),
    ("tarihsaat", "Tarih + saat"),
    ("saat", "Saat (14:30)"),
    ("gun", "Gün adı (Pazartesi)"),
    ("ay", "Ay adı (Şubat)"),
    ("yil", "Yıl (2026)"),
    ("kirp", "Baş/son boşlukları at"),
    ("tekhane", "Tek satıra indir"),
    ("ham", "HTML olarak göm (kaçışlama yok)"),
]

_FORMAT_ADLARI = {ad for ad, _ in FORMATS if ad}
# Yaygın İngilizce karşılıklar da kabul edilir.
_FORMAT_ESLESME = {
    "upper": "buyuk", "lower": "kucuk", "title": "baslik",
    "money": "para", "currency": "para", "number": "sayi", "int": "tamsayi",
    "date": "tarih", "datetime": "tarihsaat", "time": "saat",
    "day": "gun", "month": "ay", "year": "yil",
    "trim": "kirp", "oneline": "tekhane", "raw": "ham", "html": "ham",
}


def _bilinen_bicim(fmt: str):
    """Biçim adını sadeleştirir; tanınmıyorsa None."""
    f = (fmt or "").strip()
    if not f:
        return ""
    if f.startswith("%"):
        return f                      # strftime kalıbı
    k = normalize_name(f)
    if k in _FORMAT_ADLARI:
        return k
    return _FORMAT_ESLESME.get(k)


def apply_format(value, fmt: str) -> str:
    """Ham hücre değerini istenen biçimde metne çevirir."""
    fmt = (fmt or "").strip()
    if not fmt:
        return cell_to_text(value)

    if fmt.startswith("%"):
        tarih = _as_datetime(value)
        if tarih is None:
            return cell_to_text(value)
        try:
            metin = tarih.strftime(fmt)
        except ValueError:
            return cell_to_text(value)
        # %B / %A yerel ayara bağlıdır; Türkçe adları biz koyuyoruz.
        metin = metin.replace(tarih.strftime("%B"), _AY_ADLARI[tarih.month - 1])
        metin = metin.replace(tarih.strftime("%A"), _GUN_ADLARI[tarih.weekday()])
        return metin

    ad = _bilinen_bicim(fmt)
    if ad is None:
        return cell_to_text(value)      # tanınmayan biçim: değeri bozma

    if ad in ("buyuk", "kucuk", "baslik", "kirp", "tekhane", "ham"):
        metin = cell_to_text(value)
        if ad == "buyuk":
            return tr_upper(metin)
        if ad == "kucuk":
            return tr_lower(metin)
        if ad == "baslik":
            return tr_title(metin)
        if ad == "kirp":
            return metin.strip()
        if ad == "tekhane":
            return " ".join(metin.split())
        return metin                     # 'ham' yalnızca kaçışlamayı kapatır

    if ad in ("para", "sayi", "tamsayi"):
        sayi = _as_decimal(value)
        if sayi is None:
            return cell_to_text(value)
        if ad == "para":
            sayi = sayi.quantize(Decimal("0.01"))
            return _group(f"{sayi:f}")
        if ad == "tamsayi":
            return _group(str(int(sayi.to_integral_value(rounding="ROUND_HALF_UP"))))
        metin = f"{sayi:f}"
        if "." in metin:
            metin = metin.rstrip("0").rstrip(".")
        return _group(metin or "0")

    tarih = _as_datetime(value)
    if tarih is None:
        return cell_to_text(value)
    if ad == "tarih":
        return tarih.strftime("%d.%m.%Y")
    if ad == "tarihsaat":
        return tarih.strftime("%d.%m.%Y %H:%M")
    if ad == "saat":
        return tarih.strftime("%H:%M")
    if ad == "gun":
        return _GUN_ADLARI[tarih.weekday()]
    if ad == "ay":
        return _AY_ADLARI[tarih.month - 1]
    if ad == "yil":
        return str(tarih.year)
    return cell_to_text(value)


# ---------------------------------------------------------------------------
# Hazır (dahili) alanlar
# ---------------------------------------------------------------------------
#  ad, açıklama
#
# DİKKAT — ÖNCELİK: Excel sütunları dahili alanları EZER. Fatura listelerinde
# "Tarih", "Tutar" gibi bir sütun bulunması çok olağandır; kullanıcının kendi
# verisi her zaman daha özeldir. Bu yüzden bir sütun aynı ada sahipse {TARIH}
# o sütunun değerini verir. Gönderim anının tarihi her hâlükârda
# {GONDERIM_TARIHI} ile alınabilir (bu adlar sütunlarla çakışmaz).
#
BUILTINS = [
    ("SATIR", "Excel satır numarası"),
    ("EPOSTA", "Satırdaki e-posta adresi"),
    ("EK", "Ek dosyasının tam yolu"),
    ("EK_ADI", "Ek dosyasının adı (Fatura 1.pdf)"),
    ("EK_ADI_SADE", "Ek adı, uzantısız (Fatura 1)"),
    ("GONDERIM_TARIHI", "Gönderim tarihi (01.02.2026)"),
    ("GONDERIM_SAATI", "Gönderim saati (14:30)"),
    ("GONDERIM_GUNU", "Gönderim günü (Pazartesi)"),
    ("TARIH", "Gönderim tarihi — aynı adlı sütun varsa o kazanır"),
    ("SAAT", "Gönderim saati — aynı adlı sütun varsa o kazanır"),
    ("AY", "İçinde bulunulan ay (Şubat)"),
    ("YIL", "İçinde bulunulan yıl (2026)"),
    ("GONDEREN", "Gönderen adres"),
]

BUILTIN_KEYS = {normalize_name(ad) for ad, _ in BUILTINS}


def builtin_values(row=None, email="", attachment="", sender="", now=None) -> dict:
    """Dahili alanların değerlerini üretir (anahtarlar normalleştirilmiştir).

    Sonuç sözlüğü, satır değerleriyle GÜNCELLENMEK üzere tasarlanmıştır:
        v = merge.builtin_values(...); v.update(mapper.values(cells))
    Böylece aynı adlı bir Excel sütunu dahili alanı ezer (bkz. BUILTINS notu).
    """
    now = now or _dt.datetime.now()
    ek_adi = os.path.basename(attachment) if attachment else ""
    tarih = now.strftime("%d.%m.%Y")
    saat = now.strftime("%H:%M")
    gun = _GUN_ADLARI[now.weekday()]
    return {
        normalize_name("SATIR"): row if row is not None else "",
        normalize_name("EPOSTA"): email or "",
        normalize_name("EK"): attachment or "",
        normalize_name("EK_ADI"): ek_adi,
        normalize_name("EK_ADI_SADE"): os.path.splitext(ek_adi)[0],
        normalize_name("GONDERIM_TARIHI"): tarih,
        normalize_name("GONDERIM_SAATI"): saat,
        normalize_name("GONDERIM_GUNU"): gun,
        normalize_name("TARIH"): tarih,
        normalize_name("SAAT"): saat,
        normalize_name("GUN"): gun,
        normalize_name("AY"): _AY_ADLARI[now.month - 1],
        normalize_name("YIL"): str(now.year),
        normalize_name("GONDEREN"): sender or "",
    }


class RowMapper:
    """Başlıkları bir kez çözüp her satır için hızlı sözlük üretir.

    Aynı ada sahip iki sütun varsa İLKİ geçerlidir (sonrakiler yok sayılır);
    sürpriz olmaması için kural sabittir.
    """

    __slots__ = ("pairs", "keys", "names")

    def __init__(self, headers):
        self.pairs = []     # [(normal_ad, sütun_indeksi), ...]
        self.names = []     # başlıklar, yazıldığı hâliyle
        gorulen = set()
        for i, h in enumerate(headers or ()):
            ad = normalize_name(h)
            if not ad or ad in gorulen:
                continue
            gorulen.add(ad)
            self.pairs.append((ad, i))
            self.names.append(str(h))
        self.keys = gorulen

    def values(self, cells) -> dict:
        n = len(cells or ())
        return {ad: (cells[i] if i < n else None) for ad, i in self.pairs}


def row_values(headers, cells) -> dict:
    """Başlık listesi + satır hücrelerinden {normal_ad: değer} sözlüğü üretir."""
    return RowMapper(headers).values(cells)


# ---------------------------------------------------------------------------
# Şablon ayrıştırma
# ---------------------------------------------------------------------------
# Süslü parantez içi: satır sonu ve iç içe parantez içermez (CSS blokları korunsun).
_ALAN_RE = re.compile(r"\{([^{}\r\n]{1,200})\}")

# HTML gövdede taranmayacak bölgeler.
_ATLA_RE = re.compile(r"(?is)<(style|script)\b[^>]*>.*?</\1\s*>")


class Field:
    """Şablonda geçen tek bir {alan} yazımı."""

    __slots__ = ("name", "fmt", "default", "raw", "key")

    def __init__(self, name, fmt, default, raw):
        self.name = name          # kullanıcının yazdığı ad
        self.fmt = fmt            # biçim ('' olabilir)
        self.default = default    # varsayılan (None = yok)
        self.raw = raw            # şablondaki tam yazım: '{Ad|x}'
        self.key = normalize_name(name)

    def __repr__(self):  # pragma: no cover - hata ayıklama kolaylığı
        return f"<Field {self.raw}>"

    @property
    def escapes(self) -> bool:
        """HTML gövdede değer kaçışlanacak mı? (':ham' kapatır)"""
        return _bilinen_bicim(self.fmt) != "ham"


def parse_field(inner: str):
    """'Tutar:para|0,00' -> Field. Ayrıştırılamazsa None."""
    spec, ayrac, default = inner.partition("|")
    name, _, fmt = spec.partition(":")
    name = name.strip()
    if not name:
        return None
    return Field(name, fmt.strip(), default if ayrac else None, "{" + inner + "}")


def _looks_like_field(f: Field) -> bool:
    """Bu yazım gerçekten bir alan denemesi mi, yoksa CSS/metin mi?

    'p { margin:0 }' gibi bir CSS parçası da parantez içerir. Tanınmayan adlar
    zaten değiştirilmez; bu işlev yalnızca KULLANICIYI UYARIP uyarmayacağımıza
    karar verir — yanlış uyarı vermemek için katı davranır.
    """
    if ";" in f.name or ";" in (f.fmt or ""):
        return False
    if f.fmt and _bilinen_bicim(f.fmt) is None:
        return False                       # 'margin:0cm' -> alan değil
    if len(f.name) > 64 or not any(ch.isalpha() for ch in f.name):
        return False
    if f.name[0] in "@#.!$&*/\\":
        return False
    # Alan adında olmayacak karakterler (CSS/JS parçaları buradan elenir).
    if re.search(r"[<>=;{}()\"'`]", f.name):
        return False
    return True


class Template:
    """Bir şablon metnini bir kez ayrıştırır, her satır için hızlıca üretir.

    1300 satırlık bir gönderimde şablon 1300 kez ayrıştırılmaz; parçalar
    (sabit metin + alanlar) baştan hazırlanır.
    """

    __slots__ = ("text", "is_html", "_parts", "fields")

    def __init__(self, text: str, is_html: bool = False):
        self.text = text or ""
        self.is_html = bool(is_html)
        self._parts = []      # str  -> sabit metin,  Field -> alan
        self.fields = []      # sırayla geçen alanlar
        self._parse()

    # -------------------------------------------------------------- ayrıştır
    def _parse(self):
        metin = self.text
        if not metin:
            return
        # HTML'de <style>/<script> içi olduğu gibi kalır.
        bloklar = []
        if self.is_html:
            son = 0
            for m in _ATLA_RE.finditer(metin):
                bloklar.append((son, m.start(), True))
                bloklar.append((m.start(), m.end(), False))
                son = m.end()
            bloklar.append((son, len(metin), True))
        else:
            bloklar.append((0, len(metin), True))

        for bas, bit, taranir in bloklar:
            if bas >= bit:
                continue
            parca = metin[bas:bit]
            if not taranir:
                self._add_text(parca)
                continue
            son = 0
            for m in _ALAN_RE.finditer(parca):
                f = parse_field(m.group(1))
                if f is None:
                    continue
                self._add_text(parca[son:m.start()])
                self._parts.append(f)
                self.fields.append(f)
                son = m.end()
            self._add_text(parca[son:])

    def _add_text(self, s):
        if not s:
            return
        if self._parts and isinstance(self._parts[-1], str):
            self._parts[-1] += s
        else:
            self._parts.append(s)

    # ---------------------------------------------------------------- üretim
    def render(self, values: dict) -> str:
        """Alanları 'values' sözlüğündeki değerlerle doldurur.

        Sözlükte OLMAYAN bir ad hiç değiştirilmez (şablondaki hâliyle kalır).
        """
        if not self._parts:
            return self.text
        out = []
        for p in self._parts:
            if isinstance(p, str):
                out.append(p)
                continue
            if p.key not in values:
                # Bilinmeyen ada dokunmayız — CSS blokları da parantez içerir.
                # TEK istisna: kullanıcı açıkça varsayılan yazmışsa ({Ad|Sayın Müşterimiz})
                # niyet bellidir ve CSS böyle görünmez; varsayılanı basarız.
                out.append(p.default if (p.default is not None and _looks_like_field(p)) else p.raw)
                continue
            metin = apply_format(values.get(p.key), p.fmt)
            if not metin and p.default is not None:
                metin = p.default
            if self.is_html and p.escapes:
                metin = _html_escape(metin, quote=True).replace("\n", "<br />")
            out.append(metin)
        return "".join(out)

    # -------------------------------------------------------------- çözümleme
    def used_names(self):
        """Şablonda alan olarak geçen benzersiz adlar (yazıldığı hâliyle).

        CSS parçaları ('p { margin:0 }') alan sayılmaz; bkz. _looks_like_field.
        """
        gorulen, sonuc = set(), []
        for f in self.fields:
            if f.key in gorulen or not _looks_like_field(f):
                continue
            gorulen.add(f.key)
            sonuc.append(f.name)
        return sonuc

    def unknown_names(self, known_keys):
        """Bilinen adlar arasında olmayan, ama alan gibi görünen yazımlar."""
        gorulen, sonuc = set(), []
        for f in self.fields:
            if f.key in known_keys or f.key in gorulen:
                continue
            if _looks_like_field(f):
                gorulen.add(f.key)
                sonuc.append(f.name)
        return sonuc

    def required_keys(self, known_keys):
        """Varsayılanı olmayan (boş kalırsa gözle görülür) alanların anahtarları."""
        return {f.key for f in self.fields if f.key in known_keys and not f.default}


def render(text: str, values: dict, is_html: bool = False) -> str:
    """Tek seferlik kullanım için kısa yol."""
    return Template(text, is_html).render(values)


def has_fields(*texts) -> bool:
    """Verilen metinlerin herhangi birinde alan yazımı var mı?"""
    return any(_ALAN_RE.search(t or "") for t in texts)
