# Bilgi Bankası (Knowledge Base)

## ADR (Architecture Decision Records)
Tarihsel alınan kararlar ve gerekçeleri.

### ADR 1: `uv` Kullanımı
**Neden?** pip/poetry/conda yerine modern, çok daha hızlı ve bağımlılık çözümlemeyi milisaniyeler içinde yapan Rust tabanlı `uv` seçildi. Kurulumu standardize eder.

### ADR 2: `typer` + `rich` ile CLI
**Neden?** Standart `argparse` veya `click` çok fazla boilerplate kod gerektiriyor. `typer`, Python tip ipuçlarını (type hints) otomatik CLI flag'lerine dönüştürür. `rich` ise konsol çıktılarının (loglar, tablolar) terminalde çok güzel görünmesini sağlar.

### ADR 3: `pandas` Kullanarak Excel İşleme
**Neden?** `openpyxl` veya `csv` modülü ile hücreleri teker teker gezmek büyük satırlarda performansı düşürür. `pandas` vectorized (vektörel) işlemlerle 100 binlerce satırı saniyeler içerisinde filtreleyip, hızlıca yeni bir excel olarak dışa aktarabilir.

### ADR 4: `uv` ve `venv` Hibrit Başlatıcı Kontrolü
**Neden?** Şirket bilgisayarları gibi bazı kısıtlı ortamlarda `uv` paket yöneticisinin kurulması engellenebilir veya zor olabilir. Tüm `.bat` dosyaları hem `uv` hem de standart `venv` (virtualenv) ortamlarını otomatik algılayıp çalıştıracak şekilde (fallback) güncellendi. Böylece ortam fark etmeksizin taşınabilirlik sağlandı.

### ADR 5: Reçete Ağacında Operasyon Kimliği ve Üretim Durumunun Hibrit Renklendirilmesi
**Neden?** Reçete ağacında operasyonel vurgular (Montaj Otomasyonu - Elektrik Mavisi, Final Montaj - Yeşil) ile üretim takip durumları (tamamlandı - yeşil, devam ediyor - sarı, başlanmadı - kırmızı) birbiriyle çakışıyordu. Çözüm olarak; düğüm metinleri ve daire çeperlerinin (stroke) her zaman operasyon rengini koruması, daire iç dolgularının (fill) ise üretim takip durumunu göstermesi kararlaştırıldı. Böylece iki bilgi katmanı aynı anda okunabilir hale getirildi.

---

## SOLVED (Çözülen Sorunlar)
### S1: Şirket Bilgisayarında `uv` Kurulum Kısıtı
* **Sorun:** Bazı şirket bilgisayarlarında yetki kısıtlaması nedeniyle `uv` kurulamaması ve bat dosyalarının hata vermesi.
* **Çözüm:** Tüm bat başlatıcı dosyalarına `where uv` sorgusu eklendi. `uv` varsa hızlı çalışma moduna geçilir; yoksa proje dizinindeki `.venv` sanal ortamı kullanılarak çalıştırılır. Ayrıca tek tıkla sanal ortamı oluşturan ve bağımlılıkları yükleyen `1_Kurulum_Yap.bat` betiği eklendi.

### S2: Arayüzdeki Geçici Excel Yükleme Karışıklığı
* **Sorun:** HTML ağaç arayüzündeki "Excel Yükle" butonunun sadece tarayıcı belleğinde geçici güncelleme yapması ve sayfa yenilendiğinde verilerin sıfırlanması nedeniyle son kullanıcılarda oluşan kafa karışıklığı.
* **Çözüm:** Kafa karışıklığını önlemek amacıyla arayüzdeki kesikli çizgili Excel yükleme kutusu tamamen kaldırıldı. Kullanıcıların kalıcı güncellemeler için `3_Agac_Olustur.bat` veya `4_Agac_Takip.bat` (canlı izleme) yöntemlerini kullanması zorunlu hale getirildi.

### S3: Pembe (Montaj Otomasyonu) ve Kırmızı (Başlanmadı) Renk Çakışması
* **Sorun:** Montaj Otomasyonu operasyonunun rengi pembe (`#ff1493`), tamamlanmamış/başlanmamış parçaların rengi ise kırmızı (`#ff3b30`) idi. Karanlık modda bu iki renk yan yana geldiğinde veya içiçe geçtiğinde birbirine karışıyordu.
* **Çözüm:** Montaj Otomasyonu'nun operasyonel vurgu rengi neon **Elektrik Mavisi / Turkuaz (`#00f0ff`)** olarak değiştirildi. Bu sayede kırmızı durum rengiyle olan görsel kontrast en üst düzeye çıkarılarak karışma engellendi.

### S4: Kapasite Ayarlarının Dinamikleştirilmesi
* **Sorun:** 15 adet 3D Yazıcı makine sayısı ve 22 saatlik günlük çalışma kapasiteleri kod içerisine sabit (hard-coded) olarak yazılmıştı. Bu durum, gelecekte makine sayısı veya çalışma saati değiştiğinde kodun yeniden derlenmesini gerektiriyordu.
* **Çözüm:** `veritabanlari/kapasite_ayarlari.json` yapılandırma dosyası oluşturuldu. `main.py` içerisine bu dosyayı okuyup tezgah bazlı makine sayısını ve günlük çalışma saatini override eden `get_station_capacity_settings` fonksiyonu eklendi.

### S5: Mükerrer Kaynak Dosya Birleştirmesinde Veri Kaybı
* **Sorun:** Birden fazla kaynak Excel dosyası tek bir toplu listede birleştirildiğinde, aynı parça koduna ait satırlar toplanıyordu. Miktar doğru toplanırken, `KAYNAK DOSYA` sütununda `'first'` toplayıcısı kullanıldığı için yalnızca ilk dosya yazılıyor, parça diğer dosyalarda gerekmiyormuş gibi bir algı oluşuyordu.
* **Çözüm:** `src/recipe_automation/services/matcher.py` dosyasındaki `'first'` toplayıcısı yerine, benzersiz kaynak isimlerini virgülle birleştiren `combine_sources` adlı yeni bir fonksiyon entegre edildi.

### S6: Üretim Takip Sayfası Wildcard Eşleşme Hatası
* **Sorun:** Üretilen kodlar girilmesine rağmen "Genel Tamamlanma (%)" oranları hep `%0.0` kalıyordu. Bunun nedeni, `KAYNAK DOSYA` ve `DOSYA BAZLI TAKİP` sütunlarındaki değerlerin Excel'e sayısal (int/float örn: `2254` veya `2059.0`) olarak yazılmasıydı. Excel'deki `SUMIF` formülü wildcard (`*`) ile arama yaparken sayısal hücreleri metinsel olarak eşleştiremiyordu.
* **Çözüm:** `main.py` içerisinde `Üretim Takip` sayfası oluşturma aşamasında `KAYNAK DOSYA` ve `DOSYA BAZLI TAKİP` sütun değerleri temiz metne (`str`) dönüştürüldü ve ondalık (`.0`) kısımları temizlendi.

### S7: Uzun Dosyalarda Performans Tıkanması
* **Sorun:** 900+ satırlı reçetelerde Excel çıktı işleminin dakikalar sürmesi.
* **Çözüm:** `ws_izleme` sayfasındaki canlı toplanabilirlik/set limit hesaplamaları sırasında `matched_tim_relations` listesinin iç içe döngülerle taranması engellendi. Liste, döngü başlamadan önce `(kaynak, parent, child)` anahtarlarıyla sözlüklere (pre-index) dönüştürülerek arama maliyeti O(1)'e düşürüldü. Ayrıca sütun genişliği otomatik ayarlanırken ilk 100 satır ve son 10 satır örneklenerek işlem hızı 9 kat artırıldı.

### S8: İşlem Durum Bildirimi Eksikliği
* **Sorun:** Veritabanı eşleştirmesi tamamlandıktan sonra arka planda yapılan Excel dosyası yazımı, formül enjeksiyonu ve reçete ağacı oluşturma gibi ağır işlemler sırasında kullanıcının ekranında herhangi bir geri bildirim gösterilmemesi ve programın donduğu hissi.
* **Çözüm:** Kodda ana adımların başladığı yerlere progress logları eklenerek kullanıcıya anlık bildirimler sunuldu.

### S9: Koşullu Biçimlendirme Üst Limit Hatası
* **Sorun:** Üretim Takip sayfasındaki "ÜRETİLEN KOD" sütununa yazılan kodlar, listede bulunmasına rağmen kırmızı renkli hata uyarısı olarak işaretleniyordu. Nedeni, eşleşmeyi kontrol eden `COUNTIF` formülünün üst limitinin statik olarak `len(matched_df) + 1000` (örneğin 1907 satır) şeklinde belirlenmesiydi. Ancak "Üretim Takip" veya "Rotasızlar" gibi sayfaların satır sayısı 2900+'ü bulabildiğinden, bu limitin altındaki geçerli kodlar eşleştirilemeyip hatalıymış gibi kırmızıya boyanıyordu.
* **Çözüm:** `global_max_row` hesaplaması yazılan tüm sayfalardaki maksimum satır sayısı tespit edilerek ve `2000` satırlık güvenli bir ekleme marjı yapılarak dinamikleştirildi (`max_df_len + 2000`).

### S10: Dinamik Hazır/Eksik Kalem Sayımları
* **Sorun:** Üretim Takip sayfasındaki "HAZIR (Kalem)" ve "EKSİK (Kalem)" sütunlarındaki değerler, Excel ilk oluşturulduğunda hesaplanan statik tam sayılar olarak kalıyordu. Kullanıcı parça kodu girdikçe bu sayılar güncellenmiyordu.
* **Çözüm:** `EKSİK (Kalem)` hücresi için dinamik `COUNTIFS` formülü entegre edildi. `HAZIR (Kalem)` hücresi için ise `MAX(0, N{i}-P{i})` çıkarma formülü entegre edildi.

### S11: Şirket Bilgisayarında `.bat` Dosyalarında Türkçe Karakter Sorunu
* **Sorun:** Şirket bilgisayarında `1_Kurulum_Yap.bat` çalıştırıldığında `echo` komutlarındaki Türkçe karakterler (`ş`, `ç`, `ğ`, `İ`, `ü`, `ö` vb.) bozuk görünüyor ve bazı durumlarda komut hata vererek duruyor. Kendi bilgisayarında sorunsuz çalışan dosya, şirket bilgisayarının farklı Windows dil/code page veya terminal font ayarları nedeniyle Türkçe karakterleri düzgün işleyemiyor.
* **Kök Neden:** Windows `cmd.exe`'nin varsayılan karakter kod sayfası (`code page`) şirket bilgisayarlarında farklı olabilir. `chcp 65001` komutu UTF-8'e geçiş sağlasa da bazı eski Windows sürümleri veya kısıtlı ortamlarda `echo` içindeki özel karakterler yine de hatalı render edilebilir.
* **Çözüm:** Tüm `.bat` dosyalarındaki `echo` komutlarında Türkçe özel karakterler ASCII karşılıklarıyla değiştirildi (örn: `Bilgisayarda` yerine `Bilgisayarda`, `İ` → `I`, `ş` → `s` vb.). Fonksiyonel içerik (Python kodu, dosya yolları) değiştirilmedi — yalnızca kullanıcıya görünen `echo` mesajları ASCII'ye indirildi.
* **Önlem:** Bundan sonra yazılacak tüm `.bat` `echo` satırlarında Türkçe özel karakter **kullanılmamalıdır**. Sadece `@echo off`, dosya yolları ve Python/uv komutlarında Türkçe karakter kabul edilebilir.

### S12: Dashboard Üretim Loglarının Sayfa Yenilenmesinde Silinmesi
* **Sorun:** Dashboard (Web UI) üzerinde girilen üretim miktarları (`productionLog`) yalnızca tarayıcı belleğinde (RAM) tutuluyordu. Sayfa yenilendiğinde (Ctrl+Shift+R) tüm üretim girişleri kayboluyor, istasyon tablolarındaki "Hazır" / "Fazla Üretim" durumu sıfırlanarak "Eksik" olarak görünüyordu.
* **Çözüm:** Üretim logları `localStorage`'a kaydedilecek şekilde güncellendi. Yüklenen Excel dosyasının adı anahtar (key) olarak kullanıldı. Böylece sayfa yenilense veya tarayıcı kapatılıp açılsa bile, aynı dosya yüklendiğinde eski kayıtlar geri getirildi. Ayrıca hedeflenenden fazla girilen üretimler için mavi tonlu "Fazla Üretim" (Over-Production) rozeti eklendi.

## FAILED (Başarısız Denemeler)
*(Boş)*

## RECIPES (Kod Şablonları)

### Typer CLI Komutu Örneği
```python
import typer
from rich.console import Console

app = typer.Typer(help="Recipe Automation CLI")
console = Console()

@app.command()
def scan(path: str = typer.Argument(..., help="Excel dosya veya klasör yolu")):
    \"\"\"Dosyadaki veya klasördeki benzersiz operasyonları tarar.\"\"\"
    console.print(f"[green]Taranıyor: {path}[/green]")
    # ... iş mantığı
```
