# Teknik Tasarım Dokümanı (Technical Design Document)

## Sistem Mimarisi
Proje, katmanlı bir Python CLI aracı mimarisindedir.

### Katmanlar ve Akış (Request Lifecycle)
1. **CLI Layer (Typer):** Kullanıcıdan terminal üzerinden argümanı (Excel dosya yolu) alır. Giriş argümanlarının geçerliliğini doğrular.
2. **Utils Layer (Excel I/O):** Pandas ve Openpyxl kullanarak Excel dosyasını `DataFrame` nesnesine okur.
3. **Services Layer (İş Mantığı):** DataFrameler üzerinde vektörel filtrelemeler, NaN temizlemeleri ve `TIM_LISTESI` kontrollerini gerçekleştirir.
4. **Utils Layer (Excel I/O):** Filtrelenmiş veya işlem görmüş DataFrame'i yeni bir Excel olarak veya text dosyasına (`operasyon_log.txt`) kaydeder.

## Veri Akışı Diyagramı
`Kullanıcı -> [CLI: filter-id komutu] -> [Dosya Yolu Doğrulama] -> [Pandas Read] -> [Sıfır Stok & Alt Kırılım Filtresi] -> [TIM Filtresi] -> [Pandas Write] -> Başarı Mesajı`

## Performans ve Tip Güvenliği
- Tip güvenliği (Type Safety): Sistemdeki tüm argümanlar ve dönüşler `mypy` strict mode altında doğrulanacaktır.
- Performans Kısıtları: `pandas` kullanılarak büyük Excel dosyaları (>100k satır) saniyeler içerisinde işlenebilmektedir.
