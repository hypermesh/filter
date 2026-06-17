# Recipe Automation
## Proje Özeti
Şirket içi üretim ve stok yönetimi süreçlerini kolaylaştırmak adına Excel (BOM / Reçete) dosyalarından belirli kurallara göre (Sıfır miktar eleme, stok fazlası eleme, TİM operasyonlarını süzme) veri ayıklayan modern bir CLI (Command Line Interface) otomasyon aracıdır.

## Doküman Navigasyonu ( /ctx )
| Doküman | İçerik |
|---|---|
| `MILESTONE_CHECKLISTS.md` | Proje ilerleme durumu, biten ve bekleyen görevler |
| `KNOWLEDGE_BASE.md` | Kararlar, karşılaşılan sorunların çözümleri ve kod şablonları |
| `PYTHON_PROJECT_SETUP.md` | Projenin yerelde nasıl kurulacağı ve çalıştırılacağı |
| `PRODUCT_SPEC_OR_GDD.md` | CLI komutları, flag'ler ve iş kuralları (Spesifikasyon) |

## Tech Stack
| Katman | Teknoloji | Notlar |
|---|---|---|
| **Dil** | Python 3.12 | En güncel, tip uyumlu ve performanslı |
| **Paket Yönetimi** | `uv` | pip yerine çok daha hızlı modern standart |
| **Linter & Formatter** | `ruff` | Çok hızlı rust tabanlı kod standartlaştırma |
| **Type Checker** | `mypy` | Tip güvenliği (strict mode) |
| **Test** | `pytest` | Birim ve entegrasyon testleri için |
| **CLI Framework** | `typer` + `rich` | Komut satırı arayüzü ve renkli çıktılar için |
| **Veri İşleme** | `pandas` + `openpyxl` | Excel okuma, DataFrame bazlı hızlı filtrelemeler |

## Temel Mimari Kararlar
- Proje `src/recipe_automation` klasör yapısında modüler bir kütüphane/uygulama (CLI) olarak tasarlanacaktır.
- Eski "standalone script" mantığındaki `.py` dosyaları tek bir CLI komut ağına (`recipe-automation run-filter` vb.) dönüştürülecektir.
- Type annotation zorunludur.

## Kapsam Dışı (Phase 1 İçin)
- Herhangi bir veritabanı (DB) entegrasyonu.
- API veya Web Arayüzü sunumu.
- Çoklu dil (i18n) desteği.

## Başarı Kriterleri (Phase 1)
- Tüm eski `.py` scriptleri silinip, iş mantıkları servis klasörüne modüler taşınmış olmalı.
- CLI üzerinden komut verilerek klasör veya tekil Excel okunabilmeli ve doğru çıktıyı hatasız vermeli.
- Test coverage en az %70 oranında olmalı.
