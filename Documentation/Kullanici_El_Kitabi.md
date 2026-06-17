# 🏭 Makine Reçete & Filtreleme Otomasyonu — Kullanıcı El Kitabı
> **Bu Rehber Kimin İçin?** Bu doküman, kodlama bilgisi olmayan, fabrikada planlama, üretim kontrol veya depo süreçlerini yöneten kişilerin projenin ne işe yaradığını, arka plandaki kuralları ve sistemi nasıl kullanacağını en sade ve detaylı şekilde anlaması için hazırlanmıştır.

---

## 1. Proje Nedir ve Ne İş Yapar? (Büyük Resim)

Bir makine fabrikasında, tek bir makineyi üretmek için binlerce farklı parça, vida, sac levha, motor ve elektronik bileşen bir araya gelir. Bu parçaların listesine ve montaj sırasına **BOM (Bill of Materials / Malzeme Reçetesi)** adı verilir. 

Elinizde binlerce satırlık ham bir reçete Excel dosyası olduğunda, şu soruların cevaplarını manuel olarak bulmak günler sürebilir:
* *"Talaşlı İmalat (TİM) departmanı olarak bizim bu makine için hangi parçaları sıfırdan üretmemiz gerekiyor?"*
* *"Bu parçaların hangileri zaten depoda hazır bulunuyor ve üretilmesine gerek yok?"*
* *"Üreteceğimiz parçalar hangi tezgahlarda (Torna, Freze, CNC Dik İşlem vb.) kaçar dakika işlem görecek?"*
* *"Bu üretim yükünü eritmek için tezgahlarımız toplamda kaç gün çalışmalı?"*
* *"Parçaların hammaddeleri nelerdir ve ne kadar sipariş vermeliyiz?"*

**Bu Otomasyon Sistemi;** ham reçeteleri alır, referans veritabanlarıyla eşleştirir, elenmesi gerekenleri (stokta olanları, iptal edilenleri vb.) eler ve saniyeler içinde **her makineye özel, canlı formüllerle çalışan, renk kodlu, çok sayfalı mükemmel bir Üretim Takip Excel'i** ve **makine ağacı görseli** oluşturur.

---

## 2. Genel Çalışma Mantığı (Nasıl Çalışır?)

Sistem, ardışık 3 temel aşamada çalışır:

```
[1. Adım: Filtreleme] ➔ [2. Adım: Depo & Hammadde Eşleştirme] ➔ [3. Adım: Excel & Görsel Ağaç Üretimi]
```

### A. 1. Aşama: Akıllı Filtreleme (Ayıklama)
Ham reçete Excel'i sisteme verildiğinde, program ilk olarak gereksiz satırları ayıklar:
1. **İptal Edilenler/Sıfır Miktarlılar:** Miktarı sıfır olan parçalar ve bunlara bağlı alt parçalar elenir.
2. **Stok Durumu:** Eğer bir parça zaten kullanılabilir stokta varsa, o parça ve onun altındaki tüm montaj parçaları listeden çıkarılır (çünkü zaten elimizde vardır, yeniden üretmeye gerek yoktur).
3. **Departman Filtresi:** Sadece hedef departmanın (örneğin **TİM - Talaşlı İmalat**) yapacağı operasyonları içeren parçalar seçilir.

### B. 2. Aşama: Depo & Bilgi Eşleştirme (Zenginleştirme)
Elenen listedeki parçalar sadece birer koddan ibarettir. Program, arka plandaki fabrika kataloglarını (`TumRotaBilgileri.xlsx` ve `ReceteTumRotaListe.xlsx`) tarayarak şu bilgileri otomatik olarak çeker ve parçaların yanına ekler:
* Parçanın fabrikadaki tüm rotası (hangi tezgahlara uğrayacağı).
* Her tezgahtaki hazırlık süresi (ayarlama süresi) ve birim işlem süresi (tek bir adedin işlenme süresi).
* Parçanın üretilmesi için gereken hammadde kodu, hammadde adı ve birim hammadde miktarı.

### C. 3. Aşama: Dağıtım ve Akıllı Excel Oluşturma
Zenginleştirilen veriler, dinamik olarak her iş istasyonuna (makineye) göre sayfalara ayrılır:
* **SN50, SN71, CY, QUASER, PLANYA, MATKAP** vb. makinelerin her biri için Excel'de ayrı birer sekme (sayfa) açılır.
* Parçalar hangi makinede işlenecekse sadece o makinenin sayfasına yazılır.
* Excel dosyasına **akıllı formüller (VLOOKUP, SUMIF vb.)** yazılır. Böylece siz ana sayfada bir miktarı değiştirdiğinizde, tüm alt sayfalardaki miktarlar, süreler ve hammadde ihtiyaçları kendiliğinden güncellenir.
* Aynı zamanda tarayıcıda açılan **interaktif 3D benzeri bir ürün ağacı (HTML)** üretilir. Hangi parçanın hangi parçaya bağlı olduğunu ve üretim durumunu görsel olarak görebilirsiniz.

---

## 3. Sistem Hangi Kriterlere ve Kurallara Dikkat Eder?

Programın hata yapmadan çalışmasını sağlayan ve fabrika verimliliğini artıran en kritik iş kuralları şunlardır:

### 1. Nokta-Ayrımlı Hiyerarşi (Ebeveyn - Çocuk İlişkisi)
Reçetelerde sıra numaraları `1`, `1.004`, `1.004.039` gibi noktalarla belirtilir.
* `1` ana montajdır.
* `1.004` onun içindeki bir alt montajdır.
* `1.004.039` ise en alttaki tekil parçadır.
* **Kritik Kural:** Eğer ana montaj (`1.004`) iptal edilmişse veya depoda hazır bulunuyorsa, program onun altındaki `1.004.039` gibi tüm çocuk parçaları da otomatik olarak eler. Böylece gereksiz üretim planlaması yapılmaz.

### 2. Hariç Tutulan Parçalar Filtresi (`haric_tutulacak_parcalar.json`)
Fabrikada cıvata, somun, kelepçe gibi hazır satın alınan veya üretimi TİM departmanını ilgilendirmeyen standart malzemeler bulunur. Bu malzemelerin kodları sistemde kayıtlıdır. Program bu kodları gördüğünde, reçeteden ve onun alt kırılımlarından tamamen siler.

### 3. Setup Yükü (%) ve Verimli Adet Uyarısı (🔴 Kırmızı Alarm)
Bir tezgahta parçayı işlemeye başlamadan önce makineyi ayarlamak (aparat bağlamak, program yüklemek vb.) belirli bir zaman alır. Buna **Hazırlık Süresi (Setup)** denir.
* Eğer 1 adet parça üretecekseniz ve makine ayarı 2 saat sürüyor, parçanın işlenmesi ise sadece 5 dakika sürüyorsa, bu çok verimsiz bir üretimdir.
* **Hesaplama:** `Setup Yükü = Hazırlık Süresi / (Üretilecek Adet × Birim İşlem Süresi)`
* **Kural:** Eğer bu oran **%15'in üzerindeyse**, program Excel'de o satırı otomatik olarak **🔴 kırmızıya boyar** ve yan sütunda **"Önerilen Verimli Adet"** hesaplayarak planlamacıya *"Bu parçadan 1 adet değil, en az 8 adet üretirsen makine hazırlık süresi boşa gitmez"* uyarısı verir.

### 4. Günlük Tezgah Kapasitesi ve İş Yükü Hesaplama
Her tezgah sayfasının altında otomatik bir özet tablosu yer alır. Bu tablo, o tezgaha atanan tüm parçaların toplam kaç günde biteceğini hesaplar.
* **Dinamik Tezgah Sayısı:** Program, ilgili istasyonda aktif çalışan kaç adet tezgah olduğunu arka plandaki ayarlardan (`istasyon_gruplari.json`) dinamik olarak okur (örneğin CY istasyonunda 3 makine, Quaser istasyonunda 6 makine vardır).
* **Çalışma Kapasitesi:** 
  * CY ve QUASER gibi çoklu ve yoğun çalışan CNC istasyonları için günlük kapasite **27 saat/gün** (9 saat × 3 makine) kabul edilir.
  * 3D Yazıcı grubu (15 makine) için **330 saat/gün** kabul edilir.
  * Diğer standart tezgahlar için **9 saat/gün** kabul edilir.
* Bu saatlere göre toplam iş yükü süresi bölünerek net **"Gereken İş Günü"** planlamacının önüne konur.

### 5. Canlı Üretim Takibi (Yeşil / Mavi Boyama)
Excel çıktısındaki **"Üretim Takip"** sayfası interaktif bir yönetim panelidir.
* Siz üretimi tamamlanan bir parçanın kodunu ve miktarını bu sayfadaki ilgili sütunlara girdiğinizde:
  * Program ilgili parçanın tüm tezgah sayfalarındaki satırını otomatik olarak **🟢 yeşile boyar** (üretim tamamlandı anlamında).
  * Kalan miktarı düşer ve ilerleme çubuğunu günceller.
  * Eğer listede olmayan yanlış/hatalı bir parça kodu girerseniz, o hücre anında **🔴 kırmızıya boyanarak** sizi uyarır.

---

## 4. Adım Adım Kullanım Kılavuzu

Programı kullanmak için hiçbir kod yazmanıza veya siyah ekran komutları ezberlemenize gerek yoktur. Klasördeki `.bat` dosyaları sizin için her şeyi yapar.

### Adım 1: Reçeteyi Filtreleme ve Depo Eşleştirme (Tek Adımda)
1. Ham Excel reçete dosyanızı fareyle sürükleyin.
2. Klasördeki **`2_Akilli_Filtre.bat`** dosyasının üzerine bırakın.
3. Açılan siyah ekranda program size hangi üretim grubunu (örneğin `TIM`) filtrelemek istediğinizi soracaktır. `TIM` yazıp Enter'a basın.
4. Program çalışacak, kütüphaneleri tarayacak, stokları düşecek, hammadde join işlemlerini yapacak ve doğrudan **`Filtered_TIM_2241.xlsx`** (reçete numaranıza göre) dosyasını tek adımda üretecektir.

### Adım 2: Görsel Makine Ağacını Üretme
1. Üretilen `Filtered_TIM_2241.xlsx` dosyasını fareyle sürükleyin.
2. Klasördeki **`3_Agac_Olustur.bat`** dosyasının üzerine bırakın.
3. Bilgisayarınızın tarayıcısında (Chrome/Edge vb.) otomatik olarak interaktif bir ağaç şeması açılacaktır. Buradan parçaların ilişkilerini tıklayarak inceleyebilirsiniz.

### Adım 3: Canlı Reçete Ağacı İzleme Sunucusu
1. Excel üzerinde üretim takibi yaptıkça ağacın canlı boyanmasını görmek için **`4_Agac_Takip.bat`** dosyasını çalıştırın.
2. Bu sunucu arka planda açık kaldığı sürece, Excel'e girdiğiniz her veri tarayıcıdaki ağaç ekranını canlı olarak güncelleyecektir.

### Adım 4: Kapasite Analizi (İsteğe Bağlı)
1. Elinizdeki stoklarla kaç adet komple makine üretebileceğinizi simüle etmek için ham Excel dosyasını **`5_Kapasite_Analizi.bat`** üzerine bırakın.
2. Program size kaç adet makine hedeflediğinizi soracak ve buna göre bir eksik listesi çıkaracaktır.

---

## 5. Excel Sayfa Yapısı ve Renklerin Anlamı

Oluşan **`Filtered_TIM_*.xlsx`** dosyasını açtığınızda karşılaşacağınız yapının açıklaması:

| Sayfa Adı | Ne İçin Kullanılır? | Dikkat Edilmesi Gerekenler |
|---|---|---|
| **ÜRETİM LİSTESİ** | Üretilecek tüm parçaların ana özet listesidir. | Miktarları buradan değiştirebilirsiniz. |
| **Üretim Takip** | Günlük olarak fabrikada üretilen miktarların girildiği takip ekranıdır. | **I (Üretilen Kod)** ve **J (Üretim Adedi)** sütunlarına veri girişi yapılır. |
| **HAMMADDE SİPARİŞ** | Hangi hammaddeden toplam ne kadar alınması gerektiğini gösterir. | Satın alma departmanına doğrudan gönderilebilir. |
| **İstasyon Sayfaları (SN50, CY vb.)** | Doğrudan tezgah başındaki operatörlere verilecek iş listeleridir. | Sağ alt köşedeki toplam iş günü ve iş saati verilerine bakılarak planlama yapılır. |

### Renk Kodları (Excel Dosyaları)
* **🟢 Yeşil Satırlar:** Üretimi tamamen bitmiş, depoya teslim edilmiş parçaları gösterir.
* **🟣 Mor Satırlar:** Özel olarak hariç tutulması gereken veya farklı bir durum içeren, sistem tarafından işaretlenmiş parça kodlarıdır.
* **🔴 Kırmızı Hücreler:** Setup yükü çok yüksek olan (verimsiz) parçaları veya Üretim Takip ekranına girilen hatalı kodları gösterir.
* **🟡 Sarı Hücreler/Veri Çubukları:** Setup yükünün büyüklüğünü görsel olarak gösteren derecelendirmelerdir.

### Renk Kodları (Reçete Ağacı / BOM Tree HTML)
Reçete ağacında iki farklı görsel katman (Operasyon Grubu ve Üretim Durumu) aynı anda renkler vasıtasıyla sunulur:

1. **Yazı ve Daire Çeperi (Border) Renkleri (Operasyonel Kimlik):**
   * **🔵 Elektrik Mavisi / Turkuaz (`#00f0ff`):** **Montaj Otomasyonu** operasyonuna ait montaj gruplarını ve alt parçalarını temsil eder. Kırmızı ile karışmaması için özel olarak neon mavi tonu seçilmiştir.
   * **🟢 Canlı Yeşil (`#32cd32`):** **Final Montaj** operasyonuna ait grupları ve alt parçalarını gösterir.
   * **⚪ Varsayılan Gri/Turkuaz:** Herhangi bir filtreleme veya özel montaj vurgusu olmayan standart reçete parçalarıdır.
   * **🟠 Turuncu (`#ff9500`):** Çoklu tekil parçaları gruplayan paket (`GROUP`) düğümleridir (Klasör mantığı).

2. **Daire İç Dolgu (Fill) Renkleri (Üretim Takip Durumu):**
   * **🟢 Yeşil Dolgu (`#32cd32`):** Alt TIM parçalarının tamamı veya kendisi **%100 tamamlanmış** montaj grupları.
   * **🟡 Sarı Dolgu (`#ffd700`):** Üretimi başlamış ancak **kısmen tamamlanmış (devam eden)** montaj grupları.
   * **🔴 Kırmızı Dolgu (`#ff3b30`):** Üretimine **henüz başlanmamış (bekleyen)** montaj grupları.
   * **⚫ Koyu Dolgu (`#1e1e1e`):** Üretim takip kapsamı dışındaki (stokta hazır olan veya TİM operasyonu olmayan) parçalar.
