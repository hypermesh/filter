# Recipe Automation - CLI Specification

## Komutlar ve Kullanım
Bu araç 6 komut üzerinden çalışır:

| Komut | Açıklama |
|---|---|
| `scan` | Operasyon sütunlarını tarar, yeni operasyonları kullanıcıya sorar ve `operasyon_gruplari.json`'a kaydeder. |
| `filter-id` | Rezerve miktarı 0 olanları (ve alt kırılımlarını) siler, seçilen grubun operasyonlarına göre filtreler. |
| `filter-stock` | Stok miktarı 0'dan büyük olanları (ve alt kırılımlarını) siler, seçilen grubun operasyonlarına göre filtreler. |
| `auto-filter` | Dosyayı inceleyerek `filter-id` mi yoksa `filter-stock` mu çalıştırılacağına otomatik karar verir. Ardından `match-depo` otomatik tetiklenir. **Bat dosyaları bu komutu kullanır.** |
| `match-depo` | Filtrelenmiş Excel dosyasını `TumRotaBilgileri.xlsx` ile eşleştirir. İstasyon sayfaları, Üretim Takip, Hammadde Sipariş oluşturur; Excel formülleri ve HTML ağacı çıktı üretir. |
| `agac` | Ham reçete Excel dosyası veya klasöründen interaktif BOM Tree HTML dosyası oluşturur. |

---

### 1. `scan`
**Kullanım:** `recipe-automation scan [DOSYA_VEYA_KLASOR_YOLU]`
**Açıklama:** Verilen yoldaki dosyalardan "Operasyon" adlı tüm sütunları tarar. Daha önce görülmemiş operasyonları kullanıcıya sorgular ve `operasyon_gruplari.json` dosyasına kaydeder. Taranan tüm operasyonlar `operasyon_gecmisi.txt` dosyasına kümülatif olarak birikir.

### 2. `filter-id`
**Kullanım:** `recipe-automation filter-id [DOSYA_VEYA_KLASOR_YOLU] --group [GRUP_ADI]`
**Açıklama:** `Sıra No` ve `Rezerve Edilecek Miktar` sütunlarına göre kontrol yapar. Değeri 0 olanlar ve o ID'ye bağlı alt kırılımlar temizlenir. Geriye kalan satırlar arasından seçilen grubun operasyonlarıyla eşleşen satırlar bulunur ve yeni bir Excel dosyası oluşturulur.

### 3. `filter-stock`
**Kullanım:** `recipe-automation filter-stock [DOSYA_VEYA_KLASOR_YOLU] --group [GRUP_ADI]`
**Açıklama:** `Sira` ve `Kullanilabilir Stok` sütunlarına göre çalışır. Stoğu > 0 olan parçalar ve alt kırılımları temizlenir. Yine sadece seçilen grubun operasyonlarından geçen satırlar kaydedilir.

### 4. `auto-filter`
**Kullanım:** `recipe-automation auto-filter [DOSYA_VEYA_KLASOR_YOLU] --group [GRUP_ADI]`
**Açıklama:** Dosyada `Sıra No` ve `Rezerve Edilecek Miktar` sütunları varsa `filter-id`, yoksa `filter-stock` komutunu otomatik olarak çalıştırır. İşlem tamamlandıktan sonra `match-depo` otomatik tetiklenir. **2_Akilli_Filtre.bat bu komutu kullanır.**

### 5. `match-depo`
**Kullanım:** `recipe-automation match-depo [FILTRELENMIS_EXCEL_YOLU] --group [GRUP_ADI]`
**Açıklama:** Filtrelenmiş Excel dosyasını `TumRotaBilgileri.xlsx` (Ana Depo) ile Kod bazlı eşleştirir. Çıktı Excel'de şu sayfalar oluşur: ÜRETİM LİSTESİ, Tüm Veriler, HAMMADDE, HAMMADDE SİPARİŞ, Üretim Takip, İstasyon sayfaları (istasyon_gruplari.json'a göre), İzleme sayfaları. Excel formülleri (VLOOKUP, SUMIF) ve koşullu biçimlendirme otomatik enjekte edilir. Ayrıca `Makine_Agaci_*.html` oluşturulur.

### 6. `agac`
**Kullanım:** `recipe-automation agac [DOSYA_VEYA_KLASOR_YOLU]`
**Açıklama:** Ham makine reçetesi Excel dosyası veya klasörü verilir. Her dosyanın BOM (Bill of Materials) ağacını JSON'a çevirip `template.html` şablonuna yerleştirerek `Makine_Agaci_*.html` dosyası oluşturur. Tarayıcıda otomatik açılır.

---

## Hata Mesajları Standardı
- Eğer yanlış yol verilirse: `[Hata] Dosya veya klasör bulunamadı: {yol}`
- Gerekli sütunlar yoksa: `[Uyarı] {dosya_adi} atlanıyor: Gerekli sütunlar bulunamadı.`

## Haric Tutulacak Parçalar (haric_tutulacak_parcalar.json)
Filtreleme sırasında kullanıcıya "Bu parçaları sil mi?" diye sorulur.
- **Evet (sil):** Parçalar Excel çıktısına hiç girmez.
- **Hayır (silme):** Parçalar Excel'e girer ama tüm sayfalarda **mor** (F2EBF9) rengiyle işaretlenir.
