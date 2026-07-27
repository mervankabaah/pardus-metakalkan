# Pardus MetaKalkan — Belge Gizlilik Denetçisi

Pardus MetaKalkan, bir dosyayı paylaşmadan önce içindeki gizlilik verilerini yerel bilgisayarınızda gösteren ve seçtiklerinizi temiz bir kopyadan kaldıran GTK 3 masaüstü uygulamasıdır. Ağ bağlantısı kullanmaz; dosya hiçbir sunucuya yüklenmez.

> Bu araç bir belgeyi göndermeden önceki son denetim içindir. Kopyayı göndermeden önce açıp içeriğin ve biçimin beklendiği gibi olduğunu doğrulayın.

## Neleri denetler?

| Dosya türü | Bulduğu veriler | Temizleme |
| --- | --- | --- |
| PDF | Yazar, oluşturucu, başlık, tarih, XMP ve PDF bilgi sözlüğü | Her alan ayrı ayrı seçilebilir |
| DOCX | Yazar, son düzenleyen, kurum, özel özellikler, yorumlar ve revizyon izleri | Yorumlar silinir; revizyonlar kabul edilerek izler kaldırılır |
| XLSX | Yazar, son düzenleyen, kurum, özel özellikler ve hücre yorumları | Seçilen alanlar/yorumlar silinir |
| JPEG, TIFF, WebP, PNG | EXIF yazar/kamera/tarih ve GPS konumu | EXIF etiketleri tek tek silinebilir |

Temizlenen dosya, kaynak dosyaya dokunulmadan aynı klasöre `dosya_adi_temizlenmis.uzanti` adıyla yazılır. Aynı ad önceden varsa uygulama numara ekler.

## Pardus / Debian kurulumu

Uygulama Python 3 ve GTK 3 kullanır. Aşağıdaki paketler GTK bağlarını ve sanal ortam desteğini sağlar:

```bash
sudo apt update
sudo apt install python3 python3-venv python3-gi gir1.2-gtk-3.0
```

Depoyu indirin ve klasöre girin:

```bash
git clone https://github.com/KULLANICI_ADINIZ/pardus-metakalkan.git
cd pardus-metakalkan
```

Sanal ortamı oluşturup etkinleştirin:

```bash
python3 -m venv .venv --system-site-packages
source .venv/bin/activate
```

`--system-site-packages` seçeneği, apt ile kurulan GTK/PyGObject bağlarının sanal ortamda da görünmesini sağlar. Python paketlerini yükleyin:

```bash
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

Uygulamayı çalıştırın:

```bash
python metakalkan.py
```

Sanal ortamdan çıkmak için `deactivate` yazabilirsiniz.

## Kullanım

1. Dosyayı pencerenin üst bölümüne sürükleyip bırakın veya **Gözat…** düğmesine basın.
2. MetaKalkan dosyayı çevrimdışı tarar; bulduğu her veriyi türü, adı ve değeriyle listeler.
3. Silmek istemediklerinizin solundaki işareti kaldırın. **Tümünü seç** ve **Seçimi kaldır** düğmeleri toplu seçim içindir.
4. **Seçilenleri sil ve kopya oluştur** düğmesine basın.
5. Temizlenmiş dosya, orijinal dosyanın yanına yazılır. Orijinal değişmez.

## Masaüstü kısayolu (isteğe bağlı)

Sistem geneline kurulum için önce proje dosyalarını `/opt/metakalkan` altına kopyalayın ve `.desktop` dosyasındaki `Exec` yolunun doğru olduğundan emin olun. Ardından:

```bash
sudo cp com.pardus.MetaKalkan.desktop /usr/share/applications/
```

Geliştirme aşamasında bu adım gerekli değildir; `python metakalkan.py` yeterlidir.

## Gizlilik ve sınırlamalar

- MetaKalkan internet bağlantısı açmaz ve seçilen dosyayı dışarı göndermez.
- Parolalı/şifreli veya hasarlı dosyalar açılmayabilir.
- Dijital imzalı PDF'lerde meta veri değişikliği imzayı geçersiz kılabilir.
- Görsellerin piksel içeriği değil, seçtiğiniz meta veri temizlenir. Bazı PNG metin parçaları veya format dışı üretici verileri uygulamanın gösteremediği konumlarda bulunabilir.
- Makrolu Office dosyaları (`.docm`, `.xlsm`) bu sürümde desteklenmez; makroları korumak için özellikle kapsam dışındadır.

## Geliştirici notu

Kod tek dosyada (`metakalkan.py`) tutulur. Arayüz GTK 3/PyGObject, PDF işlemleri `pikepdf`, görüntü EXIF işlemleri Pillow, Office Open XML paket işlemleri ise `lxml` ve Python standart kütüphanesiyle yapılır.
