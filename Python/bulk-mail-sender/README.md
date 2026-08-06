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
- 🧩 **Sabit konu + içerik** — her maile aynı başlık/gövde, satıra özel ek eklenir.
- 🖋️ **Düz metin veya HTML gövde** — "İçerik HTML" kutusu; HTML'i dosyadan yükleyip tarayıcıda önizleyebilirsiniz.
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
2. **Sayfa**, **E-posta sütunu** ve **Ek (dosya yolu) sütunu**'nu seçin.
   (Başlıklardan otomatik tahmin edilir; gerekirse düzeltin.)
3. **Gönderim yöntemi**:
   - **Outlook**: Bilgisayarınızda kurulu ve ilgili hesapla oturum açmış Outlook kullanır.
     "Gönderen adres" alanına o hesabın adresini yazın (Outlook'ta ekli olmalı).
   - **SMTP**: Sunucu, port, güvenlik, kullanıcı ve şifreyi girin.
     Yaygın ayarlar: `587 + starttls` veya `465 + ssl`.
4. **Konu** ve **İçerik**'i yazın.
   - İçerik HTML ise **"İçerik HTML"** kutusunu işaretleyin (Outlook'ta `HTMLBody`, SMTP'de HTML olarak gönderilir).
   - Hazır bir HTML şablonunuz varsa **"HTML dosyasından yükle..."** ile içeri alın (kutu otomatik işaretlenir).
   - **"Tarayıcıda önizle"** ile göndermeden önce nasıl görüneceğini kontrol edin.
5. **Önce test edin**: "TEST adresi" alanına kendi adresinizi yazın veya "Adet sınırı"nı `1`
   yapın. Doğru göründüğünde alanları temizleyip gerçek gönderimi başlatın.
6. **▶ Gönderimi Başlat**. İlerlemeyi log alanından izleyin; gerekirse **■ Durdur**.

## Excel biçimi

İlk satır başlık kabul edilir. En az iki sütun olmalıdır:

| Email Adresi          | Ek                                              |
|-----------------------|-------------------------------------------------|
| info@ornek.com        | C:\...\Faturalar\Fatura 1.pdf                    |
| destek@ornek.com      | C:\...\Faturalar\Fatura 2.pdf                    |

Ek sütunundaki yol, o satırda gönderilecek dosyanın **tam yoludur**.

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
├── app.py            # Tkinter grafik arayüz
├── core.py           # Excel okuma + gönderim döngüsü (arayüzden bağımsız)
├── mailer.py         # SMTP ve Outlook gönderim arka planları
├── requirements.txt
├── .gitignore
└── README.md
```

## Lisans

MIT
