# Milestone ve Görev Takibi

## ⚠️ Zorunlu Pratikler
- Her yeni şey denemeden önce `KNOWLEDGE_BASE.md` kontrol edilecek.
- Her çalışan özellik sonrası Git commit atılacak.
- Sayıları koda gömmek yerine (hardcoded), sabitler (`core/constants.py`) kullanılacak.
- Problem çözünce `KNOWLEDGE_BASE.md` -> SOLVED kaydı girilecek.
- `make check` (lint ve typecheck) geçmeden kod birleştirilmeyecek.
- Her public fonksiyona type annotation + docstring zorunludur.

---
## 📍 Progress Summary
**Aktif Phase:** Geliştirmeler ve Hata Düzeltmeleri (Phase 4) Tamamlandı
**Durum:** Dinamik kapasite ayarları, mükerrer kaynak birleştirme düzeltmesi, Excel formül wildcard ve limit hatalarının çözümü, terminal log iyileştirmeleri ve kolay öncelik güncelleme aracı başarıyla tamamlandı.

---
## Phase 0 — Kurulum
- [x] `uv` kurulumu ve projenin başlatılması (`pyproject.toml` başarıyla yüklendi).
- [x] Linter ve Formatter ayarları (`ruff` ve `mypy` sorunsuz çalışıyor).
- [x] `pre-commit` kurulumu sağlandı.
- [x] Örnek `pytest` çalıştırılarak test ortamı doğrulandı.
- [x] Klasör yapısı (src, tests, vb.) oluşturuldu.

## Phase 1 — MVP (Mevcut Scriptlerin Aktarımı)
- [x] `Operasyon_Tara.py` mantığı -> `services/scanner.py` olarak yazılacak.
- [x] `id_bazlı.py` mantığı -> `services/filters.py` (id_based_filter) olarak yazılacak.
- [x] `reçete_detay.py` mantığı -> `services/filters.py` (stock_based_filter) olarak yazılacak.
- [x] CLI (Komut Satırı) yapılandırması: `typer` kullanılarak `recipe-automation` komutları eklenecek.
- [x] Excel okuma ve veri işleme testleri eklenecek.

## Phase 2 — Excel Tablo ve Sayfa Biçimlendirmeleri
- [x] Sütun genişliklerinin içeriğe göre dinamik ayarlanması.
- [x] Tamamlanan satırların kalın font ve istasyon yeşili/koyu yeşil ile uyumlu boyanması.
- [x] Tüm parça kodlarının sağa hizalanması.
- [x] İzleme tablolarına "Toplam Parça Adedi" (Q sütunu) entegrasyonu.

## Phase 3 — Reçete Ağacı (BOM Tree) Geliştirmeleri
- [x] Elektrik Mavisi / Turkuaz (#00f0ff) neon operasyon renginin entegrasyonu.
- [x] Grup klasör düğümlerinin (GROUP) durum renklendirmesinden muaf tutulması.
- [x] Operasyonel vurgular (yazı/çeper) ile üretim dolgu durumlarının çakışmayacak şekilde hibrit tasarımı.
- [x] Otomatik testlerin (`pytest`) çalıştırılması ve doğrulanması.

## Phase 4 — Kapasite, Hata Düzeltmeleri ve Geliştirmeler
- [x] Kapasite ve makine sayılarını `kapasite_ayarlari.json` üzerinden dinamik okuma.
- [x] Mükerrer kaynak dosyaları toplayarak birleştirme mantığı (`combine_sources` entegrasyonu).
- [x] `Üretim Takip` ve `DOSYA BAZLI TAKİP` sütunlarını temiz metin (`str`) yaparak wildcard formüllerinin Excel'de çalışmasını sağlama.
- [x] Ağır işlem adımlarında terminale progress logları/durum bilgileri yazdırma.
- [x] Reçete ağacı çözümlerinde verbose logları gizleyerek terminal çıktısını sadeleştirme.
- [x] txt dosyasından json formatına öncelik listesini dönüştüren `7_Oncelik_Guncelle.bat` ve `update_priority.py` araçlarını hazırlama (sürükle-bırak desteği dahil).
- [x] Koşullu biçimlendirme üst limitinin (`global_max_row`) dinamik satır sayısına göre (`max_df_len + 2000`) ayarlanması.
- [x] Hazır ve Eksik kalem sayımlarının statik değerler yerine Excel formülleriyle (`COUNTIFS` ve `MAX`) dinamikleştirilmesi.
