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

```bash
build.bat
```

`web/` klasörü (HTML/CSS + Inter fontu) exe'nin içine gömülür; çıktı `dist/TopluFaturaMailer.exe` tek dosyadır ve internet gerektirmez.

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
