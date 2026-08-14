# Toplu Fatura Mail Gönderici

Excel dosyasındaki her satır için (alıcı **e-posta adresi** + **ek dosya yolu**),
belirlediğiniz **konu** ve **içerik** ile toplu e-posta gönderen basit bir masaüstü
(GUI) uygulaması. Gönderimi **Outlook (masaüstü)** ya da **SMTP sunucusu** üzerinden
yapabilirsiniz.

> Örnek kullanım: 1300+ satırlık bir fatura listesindeki her müşteriye, kendi PDF
> faturasını `muhasebe@atakonline.com` adresinden otomatik göndermek.

## Özellikler

- 📄 **Excel'den okuma** — E-posta ve ek yolu sütunlarını arayüzden seçersiniz (otomatik tahmin eder).
- ✉️ **İki gönderim yöntemi** — Outlook masaüstü **veya** SMTP (TLS/SSL/şifresiz).
- 🧩 **Kişiselleştirme** — konu ve içerikte `{Sütun Adı}` yazarsınız, her maile o satırın
  değeri girer: *"Sayın Ahmet Yılmaz, 1.234,50 TL tutarındaki FTR-2026-001 numaralı faturanız…"*
- 📎 **Ek isteğe bağlı** — ek sütunu seçmeden de gönderebilirsiniz (duyuru/bilgilendirme
  maili); seçtiğinizde bazı satırların ek hücresi boş kalabilir, onlar eksiz gider.
- 🖋️ **Profesyonel HTML editörü** — tasarım (WYSIWYG) + renklendirilmiş kaynak kodu +
  bölünmüş görünüm, tablo/resim/bağlantı ekleme, Word temizleme, e-posta uyumluluk denetimi.
- 👁️ **Gerçek veriyle önizleme** — göndermeden önce her satırın mailini olduğu gibi görün.
- 🖼️ **Gövdeye gömülü resim** — eklediğiniz resimler maile gömülür (`cid:`), alıcı
  "resimleri göster" demeden görünür.
- 🧪 **Test modu** — tüm mailleri tek bir deneme adresine yönlendirip önce kendinize gönderin.
- 🔢 **Adet sınırı & başlangıç satırı** — önce 3 mail deneyin, sonra tümünü gönderin.
- ⏱️ **Hız sınırı** — mailler arasında bekleme süresi (sunucu limitlerine takılmamak için).
- 🧾 **Gönderim kaydı** — her gönderim `gonderim_sonuclari.csv`'ye (zaman, satır, e-posta, durum) yazılır.
- 📊 **Canlı log ve ilerleme çubuğu** — log alanı için **Temizle** düğmesi.
- 💾 **Ayarları hatırlar** — tüm alanlar (SMTP şifresi dahil) `%APPDATA%` altında saklanır.
- 🎨 **yavuzyazici.com teması** — arayüz siteyle aynı tasarım dili (Inter fontu, mavi vurgu) ile `web/` klasöründeki HTML/CSS'ten çizilir.

## Performans / akıcılık

Uygulama büyük listelerde (1000+ satır) donmayacak şekilde tasarlandı:

- **Arayüz hiçbir zaman beklemez.** Excel okuma, ön kontrol ve gönderim arka planda çalışır;
  pencere her an tıklanabilir, **Durdur** anında yanıt verir.
- **Ön kontrol paralel.** Ek dosyalarının varlığı iş parçacığı havuzunda aynı anda sorgulanır.
  Ağ sürücüsündeki 300 dosyalık ölçümde: **1,55 sn → 0,06 sn**.
- **Mail başına gereksiz sunucu turu yok.** SMTP bağlantısı her mailden önce NOOP ile
  yoklanmaz; yalnızca bir süredir boştaysa kontrol edilir. Aynı ölçümde gönderim
  **13,2 sn → 9,0 sn**.
- **Excel bir kez okunur.** Sonuç (dosya + değişiklik zamanı) anahtarıyla önbelleğe alınır;
  "Sütunları Oku → Kontrol Et → Gönder" zincirinde dosya tekrar ayrıştırılmaz.
- **Gönderim kaydı tamponlu.** CSV her satırda açılıp kapatılmaz. Dosya Excel'de açık
  (kilitli) olsa bile gönderim durmaz, yalnızca uyarı verir.
- **Ekler önden okunur.** Bir mail giderken sonraki eklerin içeriği arka planda belleğe alınır;
  aynı ek birden çok satırda geçiyorsa diskten tekrar okunmaz.
- **Log alanı sınırlıdır** (arayüzde son 800 satır) ve tek seferde toplu çizilir. Tam kayıt
  her zaman `gonderim_sonuclari.csv` dosyasındadır.

> **Not:** "Mailler arası bekleme" iki mailin *başlangıcı* arasındaki en az süredir.
> 1 sn'lik ayarda gönderimin kendisi 0,8 sn sürdüyse yalnızca 0,2 sn beklenir; sunucu
> hız sınırına uyulur ama boşa zaman harcanmaz.

## Kişiselleştirme — `{Sütun Adı}`

Konu ve içerikte süslü parantez içine bir **Excel sütun başlığı** yazın; gönderim
sırasında o satırın değeriyle değişir.

```
Konu   : {Fatura No} numaralı faturanız
İçerik : Sayın {Ad Soyad}, {Tutar:para} TL tutarındaki faturanız ektedir.
```

Editörde **Alan Ekle** düğmesi sütunlarınızı listeler (yanında ilk satırdaki örnek
değerle birlikte); tıklayınca imlecin olduğu yere ekler. Eklenen alanlar mavi bir
rozet olarak görünür, üzerine gelince o satırdaki değeri yazar.

### Yazım kuralları

| Yazım | Anlamı |
|---|---|
| `{Ad Soyad}` | Sütunun değeri |
| `{Ad Soyad\|Değerli Müşterimiz}` | Değer **boşsa** yazılacak metin (varsayılan) |
| `{Tutar:para}` | Biçimlendirilmiş değer |
| `{Tutar:para\|0,00}` | Biçim + varsayılan birlikte |

- Sütun adı **büyük/küçük harf ve boşluk farkı gözetmeden** eşleşir: `Fatura No`,
  `fatura no`, `FATURA NO` aynı alandır. Türkçe I/İ/ı/i ayrımı da normalleştirilir.
- **Tanınmayan bir ad asla değiştirilmez**, metinde olduğu gibi kalır. Bu kural
  bilinçlidir: HTML'deki CSS blokları (`p { margin:0 }`) da süslü parantez içerir ve
  bozulmamalıdır. `<style>` ve `<script>` blokları hiç taranmaz.
- HTML gövdede değerler **kaçışlanır** (`&`, `<`, `>` → `&amp;` …), satır sonları
  `<br />` olur. Bir hücrede HTML varsa ve olduğu gibi gömülsün istiyorsanız `:ham`
  biçimini kullanın.

### Biçimler

`buyuk` · `kucuk` · `baslik` (Her Kelime Büyük) · `para` (1.234,50) · `sayi` ·
`tamsayi` · `tarih` (01.02.2026) · `tarihsaat` · `saat` · `gun` (Pazartesi) ·
`ay` (Şubat) · `yil` · `kirp` · `tekhane` · `ham`

`%` ile başlayan bir biçim doğrudan tarih kalıbı sayılır: `{Tarih:%d %B %Y}` →
`01 Temmuz 2026` (ay ve gün adları Türkçe yazılır).

### Hazır alanlar

Excel'de olmayan, program tarafından üretilen alanlar:
`{SATIR}` · `{EPOSTA}` · `{EK}` · `{EK_ADI}` · `{EK_ADI_SADE}` ·
`{GONDERIM_TARIHI}` · `{GONDERIM_SAATI}` · `{GONDERIM_GUNU}` · `{AY}` · `{YIL}` ·
`{GONDEREN}`

> **Öncelik:** Aynı ada sahip bir Excel sütunu varsa **o kazanır**. Fatura
> listelerinde `Tarih` sütunu bulunmak çok olağandır; `{TARIH}` o sütunun değerini
> verir. Gönderim anının tarihi her zaman `{GONDERIM_TARIHI}` ile alınır.

### Kontrol Et ne söyler?

**Kontrol Et** taraması, yer tutucular için de rapor verir:

- **Tanınmayan alan** — o adda sütun yok, mailde `{Yok Boyle}` diye görünecek.
- **Değeri boş satırlar** — alan var ama o satırda hücre boş; mailde boşluk kalır.
  (`{Ad|Sayın Müşterimiz}` yazdıysanız uyarı verilmez.)

Tam liste, diğer sorunlarla birlikte `kontrol_sorunlari.csv` dosyasına yazılır.

## HTML editörü

**İçerik** alanı, e-posta için tasarlanmış tam bir HTML editörüdür.

### Görünümler

| Görünüm | Ne işe yarar |
|---|---|
| **Tasarım** | Yazdığınızı gördüğünüz düzenleme (WYSIWYG) |
| **Kod** | Renklendirilmiş HTML kaynağı, satır numaralı |
| **Bölünmüş** | İkisi yan yana; kodda yazdıkça tasarım güncellenir |

Alt kenardan sürükleyerek editörü büyütebilir, **Tam ekran** (Ctrl+Shift+F) ile
pencerenin tamamına yayabilirsiniz.

### Araç çubuğu

Geri al/yinele · paragraf biçimi (başlık, alıntı, kod) · yazı tipi (yalnızca
e-posta güvenli fontlar) · punto · kalın/italik/altı çizili/üstü çizili · yazı ve
vurgu rengi · hizalama · listeler · girinti · bağlantı (Ctrl+K) · resim · tablo ·
yatay çizgi · biçim temizleme.

### Özel düğmeler

- **Alan Ekle** — Excel sütunlarını ve hazır alanları listeler, biçim seçtirir.
- **Önizle** — Gerçek satır verisiyle üretilmiş maili gösterir; ◀ ▶ ile satırlar
  arasında gezinir, konu/alıcı/ek bilgisini ve mobil görünümü de verir.
  **Üretimi gönderimle aynı kod yapar**, yani burada gördüğünüz metin gidecek metindir.
- **Denetle** — E-posta istemcilerinde sorun çıkaracak şeyleri listeler:
  `<script>`, `flex`/`grid`, `position:absolute`, arka plan resmi, `vh/vw`,
  dışarıdan resim, alt metni eksikliği, 102 KB üstü boyut (Gmail maili kırpar) ve
  tanınmayan yer tutucular.
- **Biçimlendir** — HTML'i okunur şekilde girintiler. İçeriği değiştirmez;
  `<style>`, `<script>` ve `<pre>` bloklarına dokunmaz.
- **Word Temizle** — Word'den gelen HTML'deki `mso-*` bildirimlerini, `class=Mso…`
  özniteliklerini, koşullu yorumları ve `<o:p>` etiketlerini siler. Word'den
  **yapıştırdığınızda bu temizlik kendiliğinden** yapılır.
- **Şablon** — İçeriği adlandırıp kaydedin, sonra tek tıkla geri yükleyin
  (`%APPDATA%\TopluFaturaMailer\sablonlar`). Birkaç hazır şablon da vardır.
- **İçe/dışa aktarma** — HTML dosyasından yükleyin ya da dosyaya kaydedin.

### Resimler

Araç çubuğundaki resim düğmesiyle (veya panodan yapıştırıp sürükleyip bırakarak)
eklediğiniz resimler gövdeye gömülür. Gönderim sırasında bunlar otomatik olarak
**`cid:` gömülü parçaya** dönüştürülür — Gmail ve Outlook `data:` URI'li resimleri
göstermez, `cid:` ise her istemcide görünür. Aynı resim 1300 mailde de geçse yalnızca
bir kez okunur.

HTML mailler ayrıca okunabilir bir **düz metin alternatifiyle** gönderilir; bu hem
kibarlıktır hem de spam puanını düşürür.

## Kurulum

Python 3.9+ gereklidir.

```bash
pip install -r requirements.txt
```

- `openpyxl` — Excel okumak için.
- `pywebview` — masaüstü arayüz penceresi için (Windows'ta gömülü Edge WebView2'yi kullanır).
- `pywin32` — **yalnızca** Outlook yöntemi için (Windows). SMTP kullanacaksanız gerekmez.

## Çalıştırma

Ana (yeni) arayüz — yavuzyazici.com temalı:

```bash
python ui.py
```

Eski Tkinter arayüzü hâlâ yedek olarak durur (ek bağımlılık istemez):

```bash
python app.py
```

## Derleme (.exe)

```bat
build.bat              :: KLASÖR sürümü (varsayılan) -> dist\TopluFaturaMailer\TopluFaturaMailer.exe
build.bat tanilama     :: konsollu tanılama sürümü (hatalar ekranda görünür)
build.bat tekdosya     :: tek .exe -> dist\TopluFaturaMailer.exe
```

**Klasör sürümü önerilir.** Tek dosya (`onefile`) sürümü her çalıştırmada ~28 MB'lık
içeriği `%TEMP%` altına açar, Windows Defender de bu dosyaları her seferinde tarar;
bu sırada pencere görünür ama mesaj döngüsü çalışmadığı için Windows **"Yanıt vermiyor"**
der. Klasör sürümünde bu adım yoktur, uygulama anında açılır. Kod iki sürümde de aynıdır.
Klasör sürümünü dağıtırken **klasörün tamamını** kopyalayın (exe tek başına çalışmaz);
masaüstüne kısayol oluşturabilirsiniz.

`web/` klasörü (HTML/CSS + Inter fontu) çıktının içine gömülür; internet gerekmez.

## "Yanıt vermiyor" / açılmıyor ise

1. **Önce kaynaktan deneyin:** `python ui.py`. Burada sorun yoksa mesele paketlemededir
   (yukarıdaki klasör sürümünü kullanın).
2. **Tanılama günlüğü:** her çalıştırma `%APPDATA%\TopluFaturaMailer\calisma.log`
   dosyasına açılış aşamalarını ve 0,3 sn'den uzun süren her işlemi yazar.
   `YAVAŞ sutunlari_oku: 45.2 sn` gibi bir satır, örneğin kayıtlı Excel yolunun
   kopuk bir ağ sürücüsünde olduğunu gösterir.
3. **Konsollu sürüm:** `build.bat tanilama` ile derleyip çalıştırın; arkadaki siyah
   pencerede hata mesajları görünür. Bu sürüm ayrıca 20 saniyede bir tüm iş
   parçacıklarının nerede olduğunu günlüğe döker — takılma varsa yeri kesin belli olur.
4. **WebView2:** Arayüz Microsoft Edge WebView2 ile çizilir. Kurulu değilse uygulama
   açılışta uyarı verir; günlükte `motor=mshtml` görürseniz eksik demektir.
   [Evergreen Standalone Installer](https://developer.microsoft.com/microsoft-edge/webview2/)
   ile kurun (`motor=edgechromium` olmalı).
5. Uygulama açılırken pencerede **"Arayüz başlatılıyor…"** perdesi görünür; köprü
   kurulamazsa perde nedenini yazar. Boş/donmuş bir pencere görmezsiniz.

## Kullanım adımları

1. **Excel dosyası** seçin, ardından **Sütunları Oku**'ya basın.
2. **Sayfa** ve **E-posta sütunu**'nu seçin. **Ek (dosya yolu) sütunu isteğe
   bağlıdır** — ek göndermeyecekseniz **"— Ek gönderme —"** seçili kalsın.
   (Sütunlar başlıklardan otomatik tahmin edilir; gerekirse düzeltin.)
3. **Gönderim yöntemi**:
   - **Outlook**: Bilgisayarınızda kurulu ve ilgili hesapla oturum açmış Outlook kullanır.
     "Gönderen adres" alanına o hesabın adresini yazın (Outlook'ta ekli olmalı).
   - **SMTP**: Sunucu, port, güvenlik, kullanıcı ve şifreyi girin.
     Yaygın ayarlar: `587 + starttls` veya `465 + ssl`.
4. **Konu** ve **İçerik**'i yazın.
   - Kişiye özel yazmak için **Alan Ekle**'yi kullanın ya da doğrudan `{Sütun Adı}`
     yazın (bkz. [Kişiselleştirme](#kişiselleştirme--sütun-adı)).
   - Düz metin göndermek isterseniz **"İçerik HTML"** kutusunun işaretini kaldırın.
   - Hazır bir HTML şablonunuz varsa editördeki içe aktarma düğmesiyle alın.
   - **Önizle** ile göndermeden önce gerçek satır verisiyle nasıl görüneceğini kontrol edin.
5. **Önce test edin**: "TEST adresi" alanına kendi adresinizi yazın veya "Adet sınırı"nı `1`
   yapın. Doğru göründüğünde alanları temizleyip gerçek gönderimi başlatın.
6. **▶ Gönderimi Başlat**. İlerlemeyi log alanından izleyin; gerekirse **■ Durdur**.

## Excel biçimi

İlk satır başlık kabul edilir. **Zorunlu olan tek sütun e-posta adresidir.**

| Email Adresi          | Ad Soyad     | Fatura No     | Tutar   | Ek                            |
|-----------------------|--------------|---------------|---------|-------------------------------|
| info@ornek.com        | Ahmet Yılmaz | FTR-2026-001  | 1234,50 | C:\...\Faturalar\Fatura 1.pdf |
| destek@ornek.com      | Ayşe Öztürk  | FTR-2026-002  | 2469,00 | C:\...\Faturalar\Fatura 2.pdf |

- **Ek sütunu isteğe bağlıdır.** Seçmezseniz mailler eksiz gider — duyuru,
  bilgilendirme ya da hatırlatma maili göndermek için bu yeterlidir.
- Ek sütunu seçtiyseniz oradaki yol, o satırda gönderilecek dosyanın **tam yoludur**.
  Bazı satırlarda hücre **boş kalabilir**: o satırlar eksiz gönderilir, gönderim
  durmaz. **Kontrol Et** bu satırları bilgi olarak listeler.
- Diğer bütün sütunlar `{Sütun Adı}` ile konu ve içerikte kullanılabilir.

## Gönderim kaydı

Uygulama, her gönderimi Excel'in bulunduğu klasördeki `gonderim_sonuclari.csv` dosyasına
kaydeder (zaman, satır, e-posta, durum, detay). Bu dosya yalnızca **kayıt/rapor** amaçlıdır;
program her çalıştırmada listedeki **tüm** satırları gönderir.

## Güvenlik notları

- Tüm ayarlar (Excel yolu, sütunlar, konu/içerik, SMTP bilgileri **ve şifre** dahil)
  `%APPDATA%\TopluFaturaMailer\ayarlar.json` dosyasında saklanır — böylece her açılışta tekrar
  girmeniz gerekmez. `.exe` ve `python app.py` aynı yeri kullanır; masaüstü/exe klasörü temiz kalır.
- **Şifre uyarısı:** Şifre yalnızca base64 ile *gizlenir*, gerçek anlamda şifrelenmez. Bilgisayara
  fiziksel/uzaktan erişimi olan biri şifreyi geri çözebilir. `ayarlar.json` dosyasını paylaşmayın.
- `.gitignore`, `ayarlar.json` ile Excel/PDF/sonuç dosyalarını depoya eklemez — **müşteri
  verilerini ve şifreyi yanlışlıkla GitHub'a yüklemeyin.**
- Toplu gönderimde alıcı adreslerinin ve eklerin doğruluğundan siz sorumlusunuz; mutlaka
  önce **test modunda** deneyin.

## Dosya yapısı

```
toplu-fatura-mailer/
├── ui.py             # Ana arayüz (pywebview) + Python köprüsü
├── app.py            # Eski Tkinter arayüz (yedek)
├── core.py           # Excel okuma + gönderim döngüsü (arayüzden bağımsız)
├── merge.py          # {Sütun} yer tutucu motoru (bağımsız, test edilebilir)
├── mailer.py         # SMTP ve Outlook gönderim arka planları
├── web/
│   ├── index.html    # Arayüz iskeleti
│   ├── theme.css     # yavuzyazici.com teması
│   ├── app.js        # Arayüz mantığı
│   ├── editor.css    # HTML editörü stilleri
│   └── editor.js     # HTML editörü
├── requirements.txt
├── .gitignore
└── README.md
```

Yer tutucu motoru (`merge.py`) arayüzden ve `core.py`'den bağımsızdır; tek başına
içe aktarılıp test edilebilir:

```python
import merge
merge.render("Sayın {Ad Soyad}, {Tutar:para} TL", {"ad soyad": "ahmet yılmaz", "tutar": 1234.5})
# -> 'Sayın ahmet yılmaz, 1.234,50 TL'
```

## Lisans

MIT
