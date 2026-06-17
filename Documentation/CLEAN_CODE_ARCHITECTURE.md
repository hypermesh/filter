# Mimari ve Kod Yazım Standartları

## Klasör Yapısı (Layered Architecture)
```text
src/recipe_automation/
├── __init__.py
├── main.py              # CLI Giriş Noktası (Typer App)
├── core/
│   ├── __init__.py
│   ├── constants.py     # TİM_LISTESI ve diğer sabitler
│   └── config.py        # Çevre değişkenleri tanımları
├── cli/
│   ├── __init__.py
│   └── commands.py      # scan, filter-id komut grupları
├── services/
│   ├── __init__.py
│   ├── scanner.py       # Operasyon tarama iş mantığı
│   └── filters.py       # Pandas DataFrame filtreleme mantıkları
└── utils/
    ├── __init__.py
    └── excel_io.py      # Dosya okuma/yazma yardımcı fonksiyonları
```

## Kurallar ve Naming Conventions
- Tip Bildirimi (Type Annotation): Projede `Any` kullanımı minimize edilmeli. Tüm public fonksiyonlar kesinlikle `-> X` tip dönüş değeri ve tipli argümanlar içermeli.
- Değişken ve Fonksiyon İsimlendirme: `snake_case` (Örn: `filter_by_id`)
- Sınıf İsimlendirme: `PascalCase`
- Sabit Değerler: `UPPER_SNAKE_CASE`
- Bağımlılık Ayrımı (Separation of Concerns): CLI katmanı I/O işleriyle, Services katmanı veri işleme (Pandas) ile, Utils katmanı ise disk I/O operasyonlarıyla ilgilenecektir.
