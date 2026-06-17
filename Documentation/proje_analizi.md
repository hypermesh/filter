# 🏭 Makine Reçete & Filtreleme Otomasyonu — Tam Proje Analizi
> **Kaynak:** Bu belge yalnızca projenin kaynak kodları, bat dosyaları ve veritabanı dosyaları incelenerek hazırlanmıştır.
> Eski .md dokümantasyon dosyaları referans alınmamıştır.

---

## 1. Projenin Genel Amacı

Bu program, bir makine üreticisi fabrikada kullanılan binlerce satırlık Excel reçetelerini (BOM — Bill of Materials / Malzeme Ağacı) işler ve belirli bir üretim departmanının (örneğin **TIM = Talaşlı İmalat**) ne üretmesi gerektiğini saniyeler içinde temiz, çok sayfalı bir Excel çıktısına dönüştürür.

### Büyük Resim — Genel Akış ve Dosya İsimleri

```
[Ham Excel Reçete]  →  2241.xlsx  (örnek)
         │
         ▼
[1. ADIM: Filtrele]  (filter-id / filter-stock / auto-filter)
         │
         └──► TEK DOSYA çıktısı:  TIM_2241.xlsx
         └──► KLASÖR çıktısı:     TIM_{klasör_adı}.xlsx
                                   (2 sayfa: TOPLU LİSTE + RAW_DATA)
         │
         ▼
[2. ADIM: Depo Eşleştir]  (match-depo — otomatik veya manual)
         │
         ├── Left Join: Hammadde (ReceteTumRotaListe.xlsx)
         ├── Miktar hesaplamaları, sütun ekleme
         ├── İstasyon bazlı sayfa bölümleme
         ├── Excel formülleri enjeksiyonu (VLOOKUP, SUMIF, ÇARPIM)
         └── Otomatik HTML ağaç üretimi
         │
         └──► Çıktı: Filtered_TIM_2241.xlsx
               (klasör modu):  Filtered_TIM_{klasör_adı}.xlsx
         │
         └──► HTML Ağaç: Makine_Agaci_TIM_2241.html
               (klasör modu): Makine_Agaci_Receteler.html

[3. ADIM (isteğe bağlı): Kapasite Analizi]
         └──► Çıktı: 2241_5Adet_TIM.xlsx  (örnek)
```

> 📌 **"ESLESTI_" diye bir dosya adı yoktur.** Bu isim eski belgelere aittir.

---

## 2. Proje Dizin Yapısı

```
Filter/                               ← Proje kök klasörü (Masaüstü)
│
├── 1_Kurulum_Yap.bat                 ← Proje gereksinimlerini otomatik kurar/günceller
├── 2_Akilli_Filtre.bat               ← Sürükle-bırak: otomatik filtreler ve depo eşleştirir
├── 3_Agac_Olustur.bat                ← Sürükle-bırak: Reçete hiyerarşik ağaç HTML'i üretir
├── 4_Agac_Takip.bat                  ← Sürükle-bırak: Reçete ağacını canlı izleme sunucusu
├── 5_Kapasite_Analizi.bat            ← Kapasite simülasyon filtresini çalıştırır
│
├── src/recipe_automation/
│   ├── main.py             (2170 satır) ← Tüm CLI komutlarının merkezi
│   ├── analysis.py         (425 satır)  ← Kapasite simülasyon modülü
│   ├── watch_excel.py      (295 satır)  ← Excel canlı takip izleme sunucusu
│   ├── core/
│   │   ├── config.py       (25 satır)   ← Merkezi ayarlar (pydantic-settings)
│   │   └── constants.py    (16 satır)   ← TIM_LISTESI sabit seti (artık aktif kullanılmıyor)
│   ├── services/
│   │   ├── filters.py      (325 satır)  ← Filtreleme algoritmaları
│   │   ├── matcher.py      (223 satır)  ← Depo eşleştirme + hammadde join
│   │   ├── bom_tree.py     (186 satır)  ← Hiyerarşik ağaç oluşturucu
│   │   ├── sorter.py       (104 satır)  ← Öncelik bazlı sıralama
│   │   └── scanner.py      (12 satır)   ← Operasyon sütunu tarayıcı
│   ├── utils/
│   │   └── excel_io.py     (29 satır)   ← Excel okuma yardımcıları
│   └── web/
│       └── template.html   (57KB)       ← D3.js interaktif ağaç HTML şablonu
│
├── veritabanlari/                    ← Referans veritabanları ve JSON ayarlar
│   ├── TumRotaBilgileri.xlsx         (3.1 MB) ← Ana Rota Deposu
│   ├── ReceteTumRotaListe.xlsx       (2.8 MB) ← Hammadde Deposu
│   ├── StokListesi.xlsx              (1.2 MB) ← Fiziki Stok (yalnızca analysis.py kullanır)
│   ├── operasyon_gruplari.json       ← Operasyon → Departman tanımları
│   ├── istasyon_gruplari.json        ← İstasyon → Makine eşleştirmeleri
│   ├── haric_tutulacak_parcalar.json ← Yasaklı parça kodları
│   ├── oncelik_sirasi.json           ← Dosya öncelik sıralaması
│   ├── operasyon_gecmisi.txt         ← scan komutu tarafından tutulan geçmiş
│   └── operasyon_log.txt             ← Bilinen operasyonların kaydı
│
├── pyproject.toml                    ← Proje tanımı ve bağımlılıklar
├── Kullanim_Rehberi.md               ← Son kullanıcı kılavuzu
└── Documentation/                    ← Eski proje belgeleri (KODA GÜVENME, ESKİ OLABİLİR)
```

---

## 3. Teknoloji Yığını

| Kütüphane | Min. Sürüm | Kullanım Amacı |
|-----------|-----------|----------------|
| Python | 3.12 | Ana dil |
| `pandas` | 2.2.0 | DataFrame manipülasyonu |
| `openpyxl` | 3.1.0 | Excel okuma/yazma, stil, formül enjeksiyonu, grafikler |
| `typer` | 0.9.0 | CLI komut arayüzü |
| `rich` | 13.0.0 | Terminalde renkli çıktı, panel, tablo |
| `pydantic-settings` | 2.0.0 | `config.py` için tip güvenli ayar yönetimi |
| `uv` | — | pip yerine kullanılan hızlı paket yöneticisi |
| D3.js | CDN | `template.html`'deki interaktif ağaç çizimi |

---

## 4. Merkezi Ayarlar (`core/config.py`)

Tüm proje genelinde kullanılan sütun adları ve klasör yolu tek noktada tanımlıdır.
`.env` dosyası oluşturulursa ortam değişkeniyle ezilebilir.

```python
class Settings(BaseSettings):
    col_sira_no_id: str       = "Sıra No"                 # ID bazlı format sıra sütunu
    col_rezerve_miktar: str   = "Rezerve Edilecek Miktar"  # ID bazlı format miktar sütunu
    col_sira_no_stock: str    = "Sira"                    # Stok bazlı format sıra sütunu
    col_kullanilabilir_stok: str = "Kullanilabilir Stok"  # Stok bazlı format miktar sütunu
    col_operasyon_keyword: str = "Operasyon"              # Operasyon sütunlarını bulmak için
    db_dir_name: str          = "veritabanlari"
    col_depo_kod: str         = "Kod"
    col_hammadde_kod: str     = "Hammadde Kod"
    col_hammadde_isim: str    = "Hammadde"
```

> ℹ️ `constants.py`'deki `TIM_LISTESI` sabit seti eski bir kalıntıdır. `main.py`'de `import` edilmez, `operasyon_gruplari.json` kullanılır.

---

## 5. Excel Okuma Katmanı (`utils/excel_io.py`)

**Yalnızca iki fonksiyon** içerir:

```python
def find_excel_files(path: str) -> list[str]:
    # Yol dosyaysa → [path]
    # Yol klasörse  → glob("*.xlsx") içinde ~$ ile başlamayanlar

def read_excel_safe(file_path: str, headers_to_try: tuple = (0, 2)) -> DataFrame | None:
    # header=0 dener, hata alırsa header=2 dener
    # dtype=str olarak okur → tip karışıklığı yaşanmaz
```

> ⚠️ `agac` komutu ve `do_match_depo` içindeki otomatik ağaç üretimi **header=0,1,2,3** dörddünü de dener çünkü reçete Excel'lerinde başlık satırı farklı olabilir.

---

## 6. Operasyon Tarayıcı (`services/scanner.py`)

```python
def extract_operations(df: DataFrame) -> Tuple[Set[str], list[str]]:
    # Sütun adında "Operasyon" (küçük/büyük harf duyarsız) geçenleri bulur
    # Bu sütunlardaki tüm unique değerleri set olarak döner
```

---

## 7. Veritabanı Dosyaları (`veritabanlari/`)

### `TumRotaBilgileri.xlsx` — Ana Rota Deposu (3.1 MB)
Her parça/bileşenin üretim rota kayıtlarını içerir:
- Hangi iş istasyonunda işlem gördüğü
- Aktiflik durumu sütunları: `receteaktifmi`, `anarotami`, `rotaaktifmi`, `istasyonanakayitmi`
- `Birim İşlem Süresi`, `Hazırlık Süresi` gibi kapasite verileri

### `ReceteTumRotaListe.xlsx` — Hammadde Deposu (2.8 MB)
Her parça kodu için hammadde bilgisini tutar:
- `Hammadde Kod`, `Hammadde` (isim), `Hammadde Miktar`

### `StokListesi.xlsx` — Fiziki Stok (1.2 MB)
Yalnızca `analysis.py` tarafından okunur.
Kullanılabilir stok `TumRotaBilgileri.xlsx`'ten, fiziki stok bu dosyadan alınır.
Dosya 3 günden eskiyse `analysis.py` terminal uyarısı verir.

---

## 8. JSON Konfigürasyon Dosyaları

### `operasyon_gruplari.json`
Hangi operasyon adının hangi departmana ait olduğunu tanımlar.
**Gerçek içerik** (dosyadan birebir doğrulandı — `veritabanlari/operasyon_gruplari.json`):
```json
{
  "3D YAZICI":        ["3D YAZICI"],
  "TIM":              ["TALAŞLI İMALAT", "TORNA + DELİK DELME + DİŞ AÇMA",
                        "FREZE + DELİK DELME + DİŞ AÇMA", "DELİK DELME - DİŞ AÇMA",
                        "TORNA", "FREZE", "KAMA KANALI", "CNC DİK PLANYA",
                        "EBATLAMA", "DÖKÜM", "NEXT 110", "QUASER - 3"],
  "BUKUM":            ["BÜKÜM"],
  "KAYNAK":           ["KAYNAK"],
  "LAZER KESIM":      ["LAZER KESİM", "LAZER PROFİL KESİM"],
  "LAZER MARKALAMA":  ["LAZER MARKALAMA"],
  "ELEKTRIK":         ["ELEKTRİK TESİSAT", "PANO MONTAJ"],
  "MONTAJ OTOMASYON": ["MONTAJ OTOMASYONU"],
  "FINAL MONTAJ":     ["FİNAL MONTAJ", "MONTAJ"],
  "ORNEK_GRUP":       ["TESTERE", "DİŞLİ AZDIRMA", "ISLAH", "BORVERK",
                        "YATAK ÇAKMA", "BOYAMA", "ELOKSAL KAPLAMA", ...
                        (toplam 21 operasyon türü — aktif kullanımda değil, örnek grup)]
}
```
> ⚠️ **Not:** `TESTERE`, `BORVERK`, `YATAK ÇAKMA` gibi operasyonlar **TIM grubunda değil**, `ORNEK_GRUP` içindedir.
> `scan` komutu yeni operasyon bulduğunda bu dosyayı günceller.

### `istasyon_gruplari.json`
Her gruptaki parçaların hangi fiziksel makineye/istasyona atanacağını tanımlar.
Bu eşleştirme çıktı Excel'indeki **sekme (sheet) adlarını** belirler.
**Gerçek içerik** (dosyadan birebir doğrulandı — `veritabanlari/istasyon_gruplari.json`):
```json
{
  "TIM": {
    "HAMMADDE":       ["TESTERE", "EBATLAMA"],
    "SN50":           ["TOS2"],
    "SN71":           ["TOS3"],
    "CY":             ["NEXT110Y", "TS4000Y-1", "TS4000Y-2"],
    "NEX110":         ["NEXT110", "NEXT 110"],
    "QUASER":         ["QUASER-1", "QUASER-2", "QUASER-3",
                       "QUASER - 1", "QUASER - 2", "QUASER - 3"],
    "PHOEBUS":        ["PHOEBUS"],
    "ARION2142":      ["ARİON2142KÖPRÜ", "ARION2142KOPRU"],
    "ARION2000S":     ["ARİONGSM2000", "ARIONGSM2000"],
    "MATKAP&KILAVUZ": ["MATKAP+KILAVUZ", "KILAVUZ", "MATKAP"],
    "TMX2000S":       ["TMX2000S"],
    "PLANYA":         ["PLANYA"],
    "SA32B":          ["NEXTURNSA32B"]
  },
  "BUKUM": {
    "HAMMADDE": [],
    "BÜKÜM 3MT": ["BÜKÜM - 3MT"],
    "BÜKÜM 4MT": ["BÜKÜM - 4MT"]
  },
  "KAYNAK": {
    "HAMMADDE": ["TESTERE", "LAZER KESİM"],
    "SAC KAYNAGI": ["SAC KAYNAĞI - 1", "SAC KAYNAĞI - 2"],
    "PASLANMAZ KAYNAGI": ["PASLANMAZ KAYNAĞI"],
    "UNİTE/FLANS KAYNAGI": ["ÜNİTE - FLANŞ KAYNAĞI"],
    "SASE KAYNAGI": ["ŞASE KAYNAĞI"]
  },
  "LAZER KESIM": {
    "HAMMADDE": ["LAZER KESİM", "LAZER PROFİL KESİM"],
    "LAZER KESIM": ["LAZER KESİM", "LAZER PROFİL KESİM"]
  },
  "LAZER MARKALAMA": {
    "LAZER MARKALAMA GRUBU": ["LAZER MARKALAMA"]
  },
  "MONTAJ OTOMASYON": {
    "MONTAJ OTO - 1": ["MONTAJ OTOMASYONU - 1"],
    "MONTAJ OTO - 2": ["MONTAJ OTOMASYONU - 2"],
    "MONTAJ OTO - 3": ["MONTAJ OTOMASYONU - 3"],
    "MONTAJ OTO - 4": ["MONTAJ OTOMASYONU - 4"]
  },
  "3D YAZICI": {
    "HAMMADDE": ["3D YAZICI"],
    "3D YAZICI GRUBU": ["3D YAZICI - 1", ..., "3D YAZICI - 15"]
  }
}
```

> ⚠️ **TIM grubunda `3D PRINT` istasyonu yoktur.** Belgenin önceki sürümündeki bu bilgi hatalıydı. `3D YAZICI` ayrı bir üst gruptur.

> 🔑 **Kritik kural (Hammadde ekleme):** `station_mapping`'de `"HAMMADDE"` anahtarı varsa `append_hammadde()` çağrılır. Bu kontrol `main.py:635` satırında `if "HAMMADDE" in station_mapping:` ile yapılır.

> 🔑 **Makine sayısı hesabı:** `machine_counts[safe_name] = max(1, len(machine_list))` — yani her istasyon sayfasının makine sayısı JSON'daki liste uzunluğundan dinamik hesaplanır. CY = 3 (NEXT110Y + TS4000Y-1 + TS4000Y-2), QUASER = 6, 3D YAZICI GRUBU = 15.

### `haric_tutulacak_parcalar.json`
Üretim listesinden tamamen çıkarılacak parça kodları.
**Gerçek içerik** (dosyadan birebir doğrulandı — 28 satır, 27 geçerli kod):
```json
["3292", "3312", "3427", "4001", "5002", "5003", "5006", "5188", "5274",
 "5779", "6142", "6263", "7495", "10004", "10005", "10149", "10150",
 "15007", "17308", "17593", "18810", "27260", "45825", "45831", "45911",
 "XYLOKAT.OYBM.012", "XYLOKAT.OYBM.013"]
```
Şu an **27 adet** kod içeriyor (25 sayısal + 2 alfasayısal).

### `oncelik_sirasi.json`
Kaynak Excel dosyalarının öncelik numaraları (küçük sayı = daha önce işlenir/gösterilir).
**Gerçek içerik** (dosyadan doğrulandı — 50 giriş):
- İlk: `"2254": 1`, Son: `"2259": 50`
- Örnek: `"2241": 47`, `"2242": 48`, `"2168": 33`, `"2169": 35`
- **50 adet** dosya kaydı içeriyor.

### `operasyon_gecmisi.txt`
`scan` komutu tarafından yönetilir. Daha önce görülen operasyonları listeler.
Yeni bir operasyon bulunduğunda bununla karşılaştırılır; listede olmayan operasyon "yeni" sayılır.

---

## 9. Filtreleme Katmanı (`services/filters.py`)

### 9.1 Türkçe Karakter Normalizasyonu (proje geneli)
`main.py` ve `analysis.py`'de her ikisinde de tanımlı yardımcı fonksiyon:
```python
def norm_col(text: str) -> str:
    # ı→i, İ→i, ş→s, ğ→g, ü→u, ö→o, ç→c
    # lower().replace(" ","").replace("_","")
```
Bu olmadan Excel sütun adı karşılaştırmaları hatalı sonuç verir.

### 9.2 ID Bazlı Filtreleme (`filter_id_based`)
**Tetikleyici:** `Sıra No` + `Rezerve Edilecek Miktar` sütunları bulunduğunda.

```
1. Rezerve Edilecek Miktar = 0 olan satırlar → silinecek
2. Bu satırların Sıra No değerleri + "." → alt kırılım prefix'leri
3. Bu prefix'le başlayan alt satırlar da → silinecek
4. Operasyon sütunlarında hedef grup bulunuyorsa → dahil et
5. (dahil et) AND NOT (silinecek) → final filtrelenmiş liste
6. aggregate_duplicates() ile mükerrer kodlar birleştirilir
```

### 9.3 Stok Bazlı Filtreleme (`filter_stock_based`)
**Tetikleyici:** `Sira` + `Kullanilabilir Stok` sütunları bulunduğunda.

Mantık tersine çevrilir:
```
Kullanilabilir Stok > 0 → silinecek (stokta var, üretmeye gerek yok)
Stok > 0 olan satırın alt kırılımları da → silinecek
```

### 9.4 Hariç Tutma (`apply_exclusions`)
```
1. haric_tutulacak_parcalar.json listesini oku
2. Bu kodlarla eşleşen satırların Sıra No'larını bul
3. "{sira}." ile başlayan tüm alt kırılımları da işaretle
4. İşaretli satırları DataFrame'den sil
```

### 9.5 Mükerrer Birleştirme (`aggregate_duplicates`)
```
- Önce "Kod" sütununu, yoksa "Sıra No" sütununu anahtar al
- Miktar sütunu (Çarpılmış Miktar > Rezerve Edilecek Miktar > ...) → sum()
- Diğer sütunlar → first()
- Sütun sırasını orijinal DataFrame ile aynı tut
```

### 9.6 Metadata Çıktısı
Her filtreleme sonrası meta dict döner:
```python
{
  "kod_sutunu": ...,
  "hedef_sutun": ...,
  "op_cols": [...],
  "satir_ilk": ...,
  "satir_son": ...,
  "orijinal_grup_toplami": float,
  "kalan_grup_toplami": float,
  "orijinal_kalem_sayisi": int
}
```

---

## 10. Depo Eşleştirme Katmanı (`services/matcher.py`)

### `match_with_depo(filtered_df, depo_df)`
```
1. filtered_df'de "Kod" veya "Sıra No" sütununu bul → aranacak kodlar seti
2. depo_df'deki "Kod" sütunuyla maske oluştur
3. Sadece eşleşen depo satırlarını döner
   (deponun tam kaydı döner, reçetenin değil)
```

### `append_hammadde(matched_df, hammadde_df)`
```
1. Her iki df'e "_JOIN_KEY" (normalize edilmiş Kod) ekle
2. hammadde_df'de duplicate satırları temizle (keep='first')
3. Left Join: pd.merge(..., how="left")
4. Bulunamayanlar → "Bulunamadı"
5. "Hammadde Kod", "Hammadde", "Hammadde Miktar" sütunlarını
   "Kod" sütununun hemen sağına taşı
```

### `append_carpimis_miktar(matched_df, filtered_df)`
```
1. filtered_df'den miktar sütununu bul (Çarpılmış Miktar > Rezerve > Kullanilabilir)
2. Aynı koddan birden fazla satır varsa miktarları topla
3. Left Join ile matched_df'e ekle
4. "Üretilecek Miktar" = Çarpılmış Miktar kopyası olarak eklenir
5. "Toplam Hammadde Miktarı" sütunu yer yer tutucu olarak 0 eklenir
   (Gerçek değer → Excel'e ÇARPIM formülü olarak enjekte edilir)
```

---

## 11. Öncelik Sıralama Servisi (`services/sorter.py`)

```python
load_priority_mapping(db_dir)
# → oncelik_sirasi.json'u okur, {"2241": 47, ...} döner

extract_file_names("2241 (4), 2242 (5)")
# → ["2241", "2242"]  (parantez ve boşlukları temizler)

calculate_row_priority(kaynak_metin, mapping)
# → Birden fazla kaynak dosya varsa en küçük öncelik sayısını seçer
# → Bulunamazsa 9999 döner

sort_dataframe(df, mapping)
# → "Öncelik Sırası" sütunu ekler
# → Öncelik Sırası + KAYNAK DOSYA'ya göre sıralar
```

---

## 12. BOM Ağaç Oluşturucu (`services/bom_tree.py`)

Excel'deki nokta-ayrımlı Sıra No yapısını (`1`, `1.004`, `1.004.039`) iç içe geçmiş Python dict'e dönüştürür.

**Desteklenen 2 Excel formatı:**
- KR_FZM tipi: `Sira | Kod | Ad | Çarpılmış Miktar | Operasyon-1...`
- 2241 tipi: `Sıra No | Kod | Malzeme | Miktar | 1. Operasyon...`

**Parent bulma algoritması:**
```
"1.004.039" → parts = ["1","004","039"] → parent_id = "1.004"
"1.004"     → parent_id = "1"
"1"         → parent_id = "0" (Root)
```

**Pandas küsurat koruması:**
`1.010` float'a çevrilince `1.01` olabilir.
Kod bunu `float(parent_id)` karşılaştırmasıyla yakalar.

**Yaprak düğüm gruplama:**
Bir düğümün 5'ten fazla yaprak çocuğu varsa, bunlar otomatik olarak
`"📦 [N Tekil Parçayı Göster]"` grubu içinde toplanır.

**`agac` komutundaki ek özellik (main.py, satır 1721):**
Depo eşleştirme sonrası otomatik çalışırsa `initial_reqs` parametresi iletilir —
Üretim Takip sayfasındaki adetler ağaç düğümlerine de bağlanır.

---

## 13. Ana Komut Merkezi (`main.py`) — Tüm CLI Komutları

### 13.1 `auto-filter` → `0_Akilli_Filtre.bat`
```python
# Excel sütunlarına bakar:
if "Sıra No" in df.columns AND "Rezerve Edilecek Miktar" in df.columns:
    → filter_id() çağır
else:
    → filter_stock() çağır
```

### 13.2 `filter-id` → `1_ID_Filtrele.bat`
1. `oncelik_sirasi.json` güncellenip güncellenmediği sorulur → hayırsa dosyayı açar ve çıkar
2. `scan()` çağrılır (yeni operasyon tespiti)
3. Hangi grup? → kullanıcı seçer
4. Hariç tutma? → kullanıcı seçer
5. Her Excel için: `apply_exclusions` → `filter_id_based` → `print_report`
6. **Tek dosya:** `TIM_2241.xlsx` çıktısı → `do_match_depo` çağrısı → `Filtered_TIM_2241.xlsx`
7. **Çok dosya:** tümü `TOPLU LİSTE + RAW_DATA` şeklinde birleştirilir → `TIM_{klasör}.xlsx` → `do_match_depo` → `Filtered_TIM_{klasör}.xlsx`

### 13.3 `filter-stock` → `2_Stok_Filtrele.bat`
`filter-id` ile aynı akış, yalnızca filtreleme fonksiyonu `filter_stock_based`.

### 13.4 `scan` → `3_Operasyon_Tara.bat`
1. Excel'deki tüm operasyon sütunlarını `extract_operations()` ile toplar
2. `operasyon_gecmisi.txt` ile karşılaştırır → yeni olanları bulur
3. Her yeni operasyon için interaktif menü:
   - Mevcut bir gruba ekle
   - Yeni grup oluştur
   - Atla (sadece geçmişe kaydet)
4. `operasyon_gruplari.json` güncellenir
5. `operasyon_gecmisi.txt` güncellenir

### 13.5 `match-depo` → `4_Depo_Eslestir.bat`
`do_match_depo()` fonksiyonunu doğrudan çağırır (filtrele + eşleştir akışında zaten otomatik çalışır).

### 13.6 `agac` → `5_Agac_Olustur.bat`
1. Tek dosya veya klasör modu
2. `Filtered_` veya grup adıyla (`TIM_`, `KAYNAK_` vb.) başlayan çıktı dosyalarını atlar
3. Her dosya için header 0,1,2,3 denenerek doğru satır bulunur
4. `build_bom_tree()` ile dict oluşturulur, öncelik sırasına göre sıralanır
5. `template.html` içindeki placeholder'lar değiştirilir:
   - `__TREE_DATA_PLACEHOLDER__` → İlk ağacın JSON'u
   - `__ALL_TREES_PLACEHOLDER__` → Tüm ağaçların JSON dizisi
   - `__FILE_NAME_PLACEHOLDER__` → İlk dosyanın adı
6. Çıktı: `Makine_Agaci_*.html` (tek dosya) veya `Makine_Agaci_Receteler.html` (klasör modu)
7. `webbrowser.open()` ile tarayıcıda otomatik açılır

---

## 14. `do_match_depo` — Projenin Kalbi (main.py)

Bu fonksiyon projenin en karmaşık kısmıdır. Adım adım akış:

```
① TumRotaBilgileri.xlsx + filtered_df → match_with_depo() → matched_df

② Aktif/Pasif temizliği (interaktif):
   receteaktifmi / anarotami / rotaaktifmi / istasyonanakayitmi
   sütunlarındaki 0/"Hayır Pasif"/"pasif" satırları kullanıcı onayıyla silinir

③ istasyon_gruplari.json okunur → seçilen grubun station_mapping'i alınır

④ "HAMMADDE" station_mapping'de varsa → append_hammadde() çağrılır

⑤ append_carpimis_miktar() → Çarpılmış Miktar + Üretilecek Miktar + Toplam Hammadde Miktarı

⑥ "Birim İşlem Süresi" varsa → "Toplam Süre" sütunu eklenir (şimdilik 0, sonra formülle doldurulur)

⑦ MONTAJ İZLEME sayfaları oluşturulur:
   - operasyon_gruplari.json'daki "MONTAJ" içeren grupları bul
   - raw_df'deki montaj satırlarını ve TIM satırlarını hiyerarşik eşleştir
   - Kalan miktarları orantılı dağıt
   - "MONTAJ OTOMASYON İZLEME" / "FINAL MONTAJ İZLEME" sayfaları oluştur

⑧ "HAMMADDE SİPARİŞ" sayfası:
   - HAMMADDE sayfasındaki unique hammadde kodları → yeni sayfa
   - Gerçek toplam → sonradan SUMIF formülü ile hesaplanacak

⑨ "Üretim Takip" sayfası oluşturulur:
   - Kod, Üretilecek Miktar, Üretilen Miktar, Kalan Miktar, Tamamlanma (%)
   - I/J/K sütunları: ÜRETİLEN KOD / ÜRETİM ADEDİ / FAZLA ÜRETİM (manuel giriş alanı)
   - M/N/O/P/Q sütunları: DOSYA BAZLI TAKİP tablosu

⑩ sort_dataframe() ile tüm sayfalar öncelik sırasına göre sıralanır

⑪ Metin olarak saklanan sayılar gerçek sayıya çevrilir

⑫ ExcelWriter ile tüm sayfalar yazılır (sheet sırası):
   ÜRETİM LİSTESİ → Tüm Veriler → HAMMADDE → HAMMADDE SİPARİŞ → Üretim Takip
   → İZLEME sayfaları (MONTAJ OTOMASYON önce, FINAL MONTAJ sonra)
   → İstasyon sayfaları (doğal sıralama: 1,2,3...10)

⑬ openpyxl ile dosya yeniden açılır ve post-processing:
   - Belirli sütunlar gizlenir (receteaktifmi, anarotami, isistasyonu vb.)
   - Görünen sütunlar genişliğe göre otomatik ayarlanır (max 45 karakter)
   - HAMMADDE sayfası gizlenir (sheet_state = 'hidden')

⑭ Excel formülleri enjekte edilir:
   - VLOOKUP: Alt sayfalardaki Üretilecek Miktar ← ÜRETİM LİSTESİ'nden çekme
   - ÇARPIM: Üretilecek Miktar × Hammadde Miktar = Toplam Hammadde Miktarı
   - SUMIF: HAMMADDE SİPARİŞ ← HAMMADDE sayfasından toplama
   - TOPLAM SÜRE: Hazırlık Süresi + (Birim İşlem Süresi × Üretilecek Miktar)
   - Setup Yükü (%): Hazırlık / (Adet × Birim Süre)
   - Önerilen Verimli Adet: IF(setup<=0.15, adet, ROUNDUP(hazırlık/(birim*0.15),0))
   - Üretilen Miktar (Üretim Takip): MIN+SUMIF+COUNTIF kombinasyonu
   - Kalan Miktar: Üretilecek - Üretilen
   - Tamamlanma (%): Üretilen / Üretilecek

⑮ Koşullu biçimlendirmeler eklenir:
   - 🟢 Yeşil satır: Üretim Takip'te üretilmiş parçalar (istasyon sayfalarında)
   - 🟡 Sarı veri çubuğu: Setup Yükü sütunları
   - 🔴 Kırmızı: Setup Yükü > %15 uyarısı
   - 🔴 Kırmızı arka plan: Geçersiz parça kodu girişi (Üretim Takip I sütunu)
   - 🟣 Mor satır: haric_tutulacak_parcalar.json'daki kodlar
   - 🔵 Mavi/yeşil veri çubukları: Tamamlanma oranları (Üretim Takip)
   - 🟥🟨🟩 Kırmızı/Sarı/Yeşil: GENEL TAMAMLANMA (<20% / 20-80% / >80%)
   - 🟢 Yeşil satır + font: İzleme sayfalarında tamamlanan alt parçalar

⑯ Her istasyon sayfasına iş yükü özeti eklenir:
   - İSTASYON SAYISI = `machine_counts[ws_name]` (JSON liste uzunluğundan dinamik hesaplanır)
     → CY: 3 (NEXT110Y+TS4000Y-1+TS4000Y-2), QUASER: 6, 3D YAZICI GRUBU: 15, SN50: 1...
   - TOPLAM İŞ SAATİ = `=SUM(Toplam Süre) / 3600`
   - TOPLAM İŞ GÜNÜ = `=TOPLAM İŞ SAATİ / günlük_kapasite`
     Günlük kapasite: `ws_name in ["CY","QUASER"]` → 27 saat/gün (9×3 tezgah)
                      `ws_name == "3D PRINT"` → 330 saat/gün (22×15 tezgah)
                      Diğerleri → 9 saat/gün
   > ⚠️ `"3D PRINT"` istasyon sayfası adı olarak kullanılır. Üst grup `"3D YAZICI"`, sayfa adı koddaki default mapping'de `"3D PRINT"` olarak geçer.

⑰ OTOMATİK REÇETE AĞACI:
   Depo eşleştirme tamamlandıktan sonra ham reçete dosyaları okunur
   ve Üretim Takip miktarları ağaç düğümlerine bağlanarak HTML ağacı üretilir

⑱ "Unassigned" istasyonlar terminale listelenir (grup atanmamış makineler)
```

---

## 15. Kapasite Analizi (`analysis.py`)

`6_Harici Analiz (Kapasite) Filtresi.bat` tarafından doğrudan `python analysis.py` olarak çalıştırılır (uv kullanmaz, `PYTHONPATH=src` ile çalışır).

**Akış:**
1. Operasyon grubu seçimi (veya "Tüm Parçalar" modu — filtresiz)
2. Hedef adet girişi (`hedef_adet`)
3. Dosya tespiti: sürükle-bırak argümanı varsa o yol, yoksa mevcut klasörde Excel ara
4. Kaynak reçeteyi okur → header satırını "kod" + "malzeme" kelimeleri arayarak tespit eder
5. `apply_exclusions()` uygular
6. İsteğe bağlı operasyon grubu filtresi
7. `TumRotaBilgileri.xlsx`'ten kullanılabilir stok, `StokListesi.xlsx`'ten fiziki stok okunur
8. İki ayrı analiz:
   - `floor(stok / birim_ihtiyaç)` = Kapasite (sınırlayıcı parça belirler)
   - Her 1..hedef_adet set için: eksik kalem sayısı ve tamamlanma yüzdesi
9. Çıktı Excel (`{dosya}_{adet}Adet_{grup}.xlsx`):
   - `Kapasite Simülasyonu`: Set bazlı tablo + mavi/yeşil veri çubukları
   - `Kullanılabilir Stok Eksikleri`: Eksik parçalar ve miktarları
   - `Fiziki Stok Eksikleri`: Aynısı fiziki stokla
10. Tüm hücrelere kenarlık + ortalama hizalama + koyu başlık uygulanır
11. Sütun genişlikleri otomatik ayarlanır

---

## 16. Çıktı Dosyaları — Gerçek İsimler ve Sayfa Yapıları

### 16.1 Filtreleme Çıktısı (1. Adım)

| Durum | Dosya Adı | İçerik |
|-------|-----------|--------|
| **Tek dosya** | `TIM_2241.xlsx` | Tek sayfa, filtrelenmiş ham liste |
| **Klasör modu** | `TIM_{klasör_adı}.xlsx` | `TOPLU LİSTE` + `RAW_DATA` sayfaları |

> Kod: `f"{group}_{os.path.basename(path)}"` (tek) / `f"{group}_{folder_name}.xlsx"` (klasör)

### 16.2 Depo Eşleştirme Çıktısı (2. Adım) — `Filtered_TIM_2241.xlsx`

> Kod: `f"Filtered_{kaynak_dosya_adi}"` — giriş dosyasının adının başına `Filtered_` eklenir.

Sayfa yapısı (koddan doğrulanmış, yazılış sırası ile):

| Sıra | Sayfa Adı | İçerik |
|------|-----------|--------|
| 1 | **ÜRETİM LİSTESİ** | Benzersiz parçalar, Üretilecek Miktar. Ana kumanda. |
| 2 | **Tüm Veriler** | Depo eşleştirme sonucu tüm ham veri |
| 3 | **HAMMADDE** | Hammadde detayları. **Gizli sayfa** (sheet_state='hidden'). SUMIF kaynağı. |
| 4 | **HAMMADDE SİPARİŞ** | SUMIF formüllü hammadde özeti |
| 5 | **Üretim Takip** | Kod, Üretilecek/Üretilen/Kalan Miktar, Tamamlanma (%) + I/J/K + M..Q sütunları |
| 6+ | **İZLEME sayfaları** | MONTAJ OTOMASYON İZLEME önce, FINAL MONTAJ İZLEME sonra (varsa) |
| son | **İstasyon sayfaları** | SN50, SN71, CY, NEX110, QUASER, PLANYA, SA32B... (doğal sıralama) |

### 16.3 Kapasite Analizi Çıktısı

| Dosya Adı | Örnek |
|-----------|-------|
| `{kaynak_dosya}_{hedef_adet}Adet_{grup}.xlsx` | `2241_5Adet_TIM.xlsx` |

Sayfalar: `Kapasite Simülasyonu`, `Kullanılabilir Stok Eksikleri`, `Fiziki Stok Eksikleri`

### 16.4 HTML Ağaç Çıktısı

| Durum | Dosya Adı |
|-------|----------|
| Tek dosya | `Makine_Agaci_{dosya_adı_uzantısız}.html` |
| Klasör modu | `Makine_Agaci_Receteler.html` |

> Çıktı dosyası tarayıcıda `webbrowser.open()` ile otomatik açılır.

---

## 17. Bilinen İnce Detaylar (Doğrudan Koddan Doğrulandı)

| # | Detay | Kaynak Satır |
|---|-------|--------|
| 1 | Excel `dtype=str` okunur — tip dönüşüm hatası önlenir | `excel_io.py:26` |
| 2 | Header satırı: `agac` 0,1,2,3 dener; `do_match_depo` 0,2 dener | `main.py:1682`, `excel_io.py:20` |
| 3 | Pandas küsurat sorunu: `1.010` → `1.01` — float karşılaştırmayla yakalanır | `bom_tree.py:123-138` |
| 4 | `global_max_row = len(matched_df) + 1000` — boş satırlar da formülle kapsanır | `main.py:1103` |
| 5 | Mor satır (F2EBF9 / 6C3483) = `haric_tutulacak_parcalar.json`'daki kodlar **tüm sayfalarda** uygulanır | `main.py:1480-1568` |
| 6 | Kırmızı hücre: Üretim Takip I sütununa geçersiz kod girilirse `AND($I2<>"", COUNTIF(...)=0)` kuralı devreye girer | `main.py:1419` |
| 7 | Makine sayısı sabit değil, `len(machine_list)`'ten dinamik hesaplanır | `main.py:686` |
| 8 | Atanmamış istasyonlar terminale listelenir, silinmez — "Tüm Veriler"de kalırlar | `main.py:1763-1766` |
| 9 | Ağaçta `Filtered_` veya grup prefix'i (`TIM_`, `KAYNAK_` vb.) olan dosyalar atlanır | `main.py:1665-1673` |
| 10 | `constants.py`'deki `TIM_LISTESI` hiçbir yerde import edilmiyor | Grep doğrulaması |
| 11 | Çoklu dosya modu: her satıra `KAYNAK DOSYA` sütunu eklenir, `"*"&M{i}&"*"` wildcard SUMIF kullanılır | `main.py:1370` |
| 12 | `do_match_depo` giriş `path`'inde `RAW_DATA` sheet okumayı dener, yoksa `filtered_df.copy()` kullanır | `main.py:544-546` |
| 13 | StokListesi.xlsx 3 günden eskiyse sarı terminal uyarısı | `analysis.py:244` |
| 14 | `agac` komutu `webbrowser.open()` ile HTML'i otomatik açar | `main.py:1948` |
| 15 | Mor renk ayrıca İZLEME sayfalarında mavi/beyaz satır renklendirmesiyle birlikte çalışır | `main.py:1507-1568` |
| 16 | BOM ağaç yaprak grubu limiti: `limit=5` — 5'ten fazla yaprak çocuk gruplanır | `bom_tree.py:185` |
| 17 | Güncellenen `Setup Yükü (%)` ve `Güncel Setup Yükü (%)` olmak üzere iki ayrı setup sütunu eklenir | `main.py:1033-1035` |
| 18 | HAMMADDE sayfası `sheet_state='hidden'` olarak kaydedilmez — sadece `target_hide_cols` ile sütunları gizlenir; HAMMADDE sayfasının gizlenmesi farklı bir mekanizma | `main.py:1046-1047` |

---

## 18. Çalıştırma

```powershell
# Normal kullanım (bat dosyası üzerine Excel sürükle-bırak)
2_Akilli_Filtre.bat

# Komut satırından
uv run recipe-automation auto-filter "C:\path\to\recete.xlsx"
uv run recipe-automation filter-id "C:\path\to\recete.xlsx"
uv run recipe-automation agac "C:\path\to\recete_klasoru"

# Geliştirme
uv run pytest
uv run ruff check .
uv run mypy src
```
