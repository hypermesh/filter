---
description: Initialize a professional project structure with all documentation and workflows for Game, Web, Mobile, Desktop or Python projects.
---

Kullanıcı yeni bir proje için profesyonel kurulum yapmak istediğinde bu workflow'u çalıştır.
Kullanım: Projeyi kısaca anlat, sonra `/setup-project` de.

---

## ⚠️ AŞAMA 1 — KAPSAMLI BİLGİ TOPLAMA (Zorunlu)

Dosya oluşturmadan önce eksik bilgileri tamamla.
Kullanıcının anlattıklarından cevabı belli olanları **SORMA**. Sadece eksik olanları sor.
Tüm soruları **tek seferde**, numaralı liste olarak sor.

---

### 🎯 Proje Kimliği

- [ ] Projenin tam adı nedir?
- [ ] Tür: `Oyun` / `Web App` / `Mobil App` / `Masaüstü Yazılım` / `Python Projesi`
  - Alt tür:
    - Oyun: RPG, Roguelike, FPS vb.
    - Web App: SaaS, Dashboard, E-ticaret vb.
    - Python: `REST API` / `CLI Tool` / `Data Science / ML` / `Library / Package` / `Automation / Scraper`
- [ ] Hedef platform: `Web` / `iOS` / `Android` / `Windows` / `MacOS` / `Console` / `Server` / Hepsi?
- [ ] Takım büyüklüğü: Solo mu, ekip mi?
- [ ] Proje kapsamı: `Jam Prototype` / `MVP` / `Bağımsız Ürün` / `Ticari Ürün`

---

### ⚙️ Teknik Altyapı

> **⚠️ KRİTİK KURAL — Web App:**
> Tür **"Web App"** seçilirse aşağıdaki yığın **zorunludur** — alternatif önerilmez:
>
> | Katman | Teknoloji |
> |---|---|
> | **Frontend** | Next.js (App Router), React, Tailwind CSS |
> | **Backend / DB** | Supabase (Postgres) |
> | **Auth** | Supabase Auth |
> | **Dil** | TypeScript (strict mode) |

> **⚠️ KRİTİK KURAL — Python Projesi:**
> Tür **"Python Projesi"** seçilirse aşağıdaki araç zinciri **zorunludur** — alternatif önerilmez:
>
> | Katman | Teknoloji | Notlar |
> |---|---|---|
> | **Paket & Ortam Yöneticisi** | `uv` | pip/poetry/conda yerine — 2025 standardı |
> | **Linter & Formatter** | `ruff` | flake8 + black + isort tek araçta |
> | **Type Checker** | `mypy` (strict) | tüm public API'lar annotated olmalı |
> | **Test Framework** | `pytest` + `pytest-cov` | — |
> | **Pre-commit Hooks** | `pre-commit` | ruff + mypy + testler commit'te çalışır |
> | **Config Yönetimi** | `pydantic-settings` | `.env` + tip güvenliği |
> | **Proje Tanımı** | `pyproject.toml` | requirements.txt yok |
>
> **Alt türe göre ek zorunlu stack:**
>
> | Alt Tür | Zorunlu Ekler |
> |---|---|
> | **REST API** | `FastAPI` + `uvicorn` + `httpx` (test client) |
> | **CLI Tool** | `typer` + `rich` |
> | **Data Science / ML** | `pandas` + `numpy` + `jupyter` + `matplotlib`/`plotly` |
> | **Library / Package** | `twine` + `build` + ReadTheDocs/mkdocs config |
> | **Automation / Scraper** | `httpx`/`playwright` + `tenacity` (retry) |

- [ ] *(Yazılım / Mobil ise)* Framework / Dil tercihi: `Flutter`, `Swift`, `.NET`, `React Native` vb.
- [ ] *(Oyun ise)* Motor / Render Pipeline: `Unity URP` / `Unity HDRP` / `Unreal` / `Godot` vb.
- [ ] *(Oyun ise)* Fizik sistemi: `Rigidbody` / `CharacterController` / `Custom`
- [ ] *(Oyun ise)* Input tipi: Klavye+fare / Gamepad / Dokunmatik / Hepsi?
- [ ] *(Python ise)* Python minimum sürümü: `3.11` / `3.12` / `3.13`?
- [ ] *(Python — REST API ise)* Database gerekli mi? `PostgreSQL` / `SQLite` / `MongoDB` / Hayır
- [ ] *(Python — REST API ise)* Auth sistemi gerekli mi? `JWT` / `OAuth2` / `API Key` / Hayır
- [ ] *(Python — ML ise)* Model servisleme gerekli mi? `ONNX` / `FastAPI endpoint` / Sadece notebook / Hayır
- [ ] *(Python — Library ise)* PyPI'ye publish edilecek mi?
- [ ] Auth sistemi gerekli mi? *(Web ise Supabase Auth, Python REST API ise JWT/OAuth2 varsayılandır)*
- [ ] Save / Persist sistemi: `Cloud DB` / `Local JSON` / `PlayerPrefs` / `Hybrid` / `Yok`
- [ ] Online / Multiplayer / Real-time özellik gerekli mi?
- [ ] *(Oyun ise)* Hedef performans: FPS hedefi ve minimum donanım?
- [ ] *(Python ise)* CI/CD hedefi: `GitHub Actions` / `GitLab CI` / `Yok`
- [ ] *(Python ise)* Docker ile deploy gerekli mi?

---

### 🏗️ Mimari & Kullanıcı / Oynanış Akışı

- [ ] *(Web / Yazılım ise)* Temel kullanıcı akışı: *"Kullanıcı girer → X yapar → Y sonucunu alır"*
- [ ] *(Python — REST API ise)* Temel endpoint grupları: *"Auth, Users, Products vb."*
- [ ] *(Python — CLI ise)* Temel komutlar: *"`run`, `build`, `deploy` vb."*
- [ ] *(Python — ML ise)* Veri kaynağı ve model türü: *"CSV → sklearn sınıflandırıcı vb."*
- [ ] *(Oyun ise)* Temel oynanış döngüsü: *(Oyuncu ne yapar, ne kazanır/kaybeder?)*
- [ ] *(Oyun ise)* Kamera tipi: `1. Şahıs` / `3. Şahıs` / `İzometrik` / `2D Side-scroll`
- [ ] *(Oyun ise)* Savaş / etkileşim sistemi var mı? *(Melee, Ranged, Büyü, Combo?)*
- [ ] *(Oyun ise)* Envanter veya progression sistemi var mı?
- [ ] Kod mimarisi tercihi: `Composition` / `ECS` / `MVC` / `MVVM` / `Layered (Service/Repo)` / Fark etmez
- [ ] Data yönetimi: `ScriptableObject` / `JSON Config` / `Supabase` / `SQLAlchemy ORM` / `Hardcoded`
- [ ] Olay sistemi: `Event-driven` / `Polling` / Fark etmez

---

### 🎨 Sanat, Ses & Görsel Stil

- [ ] Görsel stil: `Pixel Art` / `Low Poly` / `Realistic` / `Stylized` / `Cel-shaded` / `Modern Minimal`
- [ ] *(Oyun ise)* 2D sprite mi, 3D model mi, Hybrid mi?
- [ ] *(Oyun ise)* Asset üretim araçları: `Blender`, `Aseprite`, `Figma` vb.
- [ ] *(Oyun ise)* Ses/müzik: Kendi yapacak mısın, lisanssız mı kullanacaksın?
- [ ] *(Web / Yazılım ise)* UI kütüphanesi: `shadcn/ui`, `Radix`, `MUI`, `Custom` vb.

---

### 📅 Kapsam & Referanslar

- [ ] Phase 1 (MVP) için en kritik **3 özellik** nedir? *(Bunlar olmadan ürün çalışmaz)*
- [ ] Phase 1 kesin **dışında** ne var? *(Kapsam dışı liste)*
- [ ] Phase 1'i ne zamana tamamlamayı hedefliyorsun?
- [ ] İlham alınan referanslar: *(Görsel veya mekanik bazlı 2 örnek — ne almak istiyorsun?)*

---

Bilgiler toplandıktan sonra kullanıcıya şunu söyle:
> "Harika! Şimdi [X] dosya oluşturacağım. Başlıyorum..."

---

## 📁 AŞAMA 2 — KLASÖR YAPISI

Proje türüne göre kök dizinde oluştur:

```
[ProjeKökü]/
├── Documentation/           # Tüm tasarım ve teknik dökümanlar
├── .agent/
│   └── workflows/           # AI asistan iş akışları
│
├── src/                     # ✅ (Web / Yazılım ise)
│   ├── app/                 #    Next.js App Router sayfaları
│   ├── components/          #    Yeniden kullanılabilir UI bileşenleri
│   └── lib/                 #    Yardımcı fonksiyonlar & servisler
│
├── supabase/                # ✅ (Web ise)
│   └── migrations/          #    SQL şemaları ve migration'lar
│
├── src/[proje_adi]/         # ✅ (Python ise — src layout zorunlu)
│   ├── __init__.py
│   ├── main.py              #    Giriş noktası
│   ├── core/                #    Ayarlar, config, sabitler
│   │   ├── __init__.py
│   │   └── config.py        #    pydantic-settings Config sınıfı
│   ├── api/                 #    ✅ (REST API ise) route'lar & şemalar
│   │   ├── __init__.py
│   │   ├── routes/
│   │   └── schemas/
│   ├── services/            #    İş mantığı katmanı
│   ├── repositories/        #    ✅ (DB varsa) veri erişim katmanı
│   ├── models/              #    Domain modelleri / ORM modelleri
│   ├── cli/                 #    ✅ (CLI ise) typer komutları
│   └── utils/               #    Yardımcı fonksiyonlar
│
├── tests/                   # ✅ (Python ise)
│   ├── __init__.py
│   ├── conftest.py          #    Paylaşılan fixtures
│   ├── unit/                #    Birim testler (dış bağımlılık yok)
│   ├── integration/         #    Entegrasyon testleri (DB, API vb.)
│   └── e2e/                 #    ✅ (API / CLI ise) uçtan uca testler
│
├── notebooks/               # ✅ (Data Science / ML ise)
│   ├── 01_eda.ipynb
│   └── 02_modeling.ipynb
│
├── data/                    # ✅ (Data Science / ML ise)
│   ├── raw/                 #    Ham veri — asla değiştirme
│   ├── processed/           #    İşlenmiş veri
│   └── .gitkeep
│
├── scripts/                 # ✅ (Python ise) tek seferlik yardımcı scriptler
│
├── .github/
│   └── workflows/           # ✅ (Python + CI/CD ise)
│       ├── ci.yml           #    Test & lint pipeline
│       └── release.yml      #    ✅ (Library ise) PyPI publish
│
├── pyproject.toml           # ✅ (Python ise) ZORUNLU — tek config dosyası
├── .python-version          # ✅ (Python ise) uv için Python sürümü
├── .pre-commit-config.yaml  # ✅ (Python ise) ruff + mypy hooks
├── .env.example             # ✅ (Python ise) örnek ortam değişkenleri
├── Makefile                 # ✅ (Python ise) geliştirici komutları
├── Dockerfile               # ✅ (Python + Docker ise)
│
└── Assets/                  # ✅ (Oyun ise)
    ├── Art/
    ├── Audio/
    └── Prefabs/
```

---

## 📝 AŞAMA 3 — DOCUMENTATION DOSYALARINI OLUŞTUR

Her dosyayı **projeye özel** içerikle doldur — genel şablon değil.

---

### `PROJECT_OVERVIEW.md`
- Projenin tek cümlelik özeti
- `/ctx` komutuna referans ve belge navigasyon tablosu
- Tech Stack tablosu
  - *(Web ise **Next.js + Tailwind + Supabase + TypeScript** vurgusu)*
  - *(Oyun ise Motor + Pipeline + Dil vurgusu)*
  - *(Python ise **uv + ruff + mypy + pytest** ve alt türe özel stack vurgusu)*
- Temel mimari kararlar tablosu
- Phase 1 kapsam dışı özellikler
- Phase 1 başarı kriterleri (ölçülebilir)

---

### `PRODUCT_SPEC_OR_GDD.md`

**Yazılım / Web App ise — Product Spec:**
- Kullanıcı hikayeleri *(User Stories)*
- Ekran listesi ve akış diyagramı
- Fonksiyonel gereksinimler
- UI/UX prensipleri ve kullanılan kütüphane kararı

**Python — REST API ise — API Spec:**
- Tüm endpoint listesi: method, path, request/response şeması
- Auth akışı diyagramı
- Hata kodları ve mesajlar tablosu
- Rate limiting ve güvenlik kuralları

**Python — CLI Tool ise — CLI Spec:**
- Tüm komutlar ve flag'ler tablosu
- Her komut için örnek kullanım
- Hata mesajları standardı
- `--help` çıktı formatı

**Python — Data Science / ML ise — ML Spec:**
- Problem tanımı ve başarı metrikleri
- Veri kaynağı ve EDA planı
- Model seçim kriterleri
- Baseline → hedef metrik karşılaştırması
- Deployment / servis planı

**Python — Library / Package ise — Library Spec:**
- Public API tasarımı (sınıflar, fonksiyonlar, imzalar)
- Kullanım örnekleri (README'ye gidecek)
- Versiyonlama stratejisi (SemVer)
- Breaking change politikası

**Oyun ise — GDD (Game Design Document):**
- Vizyon ve 3–5 tasarım pillar'ı
- Temel oynanış döngüsü
- Tüm mekanikler detaylı
- Kontrol şeması
- Progression ve envanter sistemi
- UI yoğunluğu ve HUD tasarımı
- Görsel & ses stili

---

### `TECHNICAL_DESIGN_DOC.md`
- Mimari genel bakış
  - *(Web ise: Server vs Client Components stratejisi, API route yapısı)*
  - *(Oyun ise: Her sistemin nasıl çalıştığı + kod iskeleti)*
  - *(Python ise: Katmanlı mimari diyagramı — core / services / repositories / api)*
- Sistem bağımlılıkları diyagramı
- Veri akış diyagramı
- *(Python — REST API ise)* Request lifecycle: middleware → router → service → repository
- *(Python — ML ise)* Model training pipeline diyagramı
- Performans hedefleri ve kısıtları
- *(Python ise)* Type annotation stratejisi ve mypy config kararları

---

### `DATABASE_SCHEMA.md` *(Web / Cloud persist / Python REST API + DB içeriyorsa)*
- Supabase / Postgres / SQLite tablo yapıları
- Kolon tipleri, ilişkiler ve index'ler
- *(Web ise)* RLS *(Row Level Security)* politikaları ve güvenlik kuralları
- *(Python + SQLAlchemy ise)* ORM model örnekleri ve ilişki tanımları
- *(Python ise)* Alembic migration stratejisi
- Migration stratejisi

---

### `PYTHON_PROJECT_SETUP.md` *(Sadece Python projeleri için)*

Bu dosya projenin teknik kurulum rehberidir — her yeni geliştirici buradan başlar.

İçermesi gerekenler:

**Ortam Kurulumu:**
```bash
# uv ile kurulum
curl -LsSf https://astral.sh/uv/install.sh | sh
uv python install 3.12
uv sync --all-extras
source .venv/bin/activate   # Linux/Mac
# ya da
.venv\Scripts\activate      # Windows
```

**`pyproject.toml` Tam Yapısı:**
- `[project]` — metadata, dependencies, python sürümü
- `[project.optional-dependencies]` — `dev`, `test`, `docs` grupları
- `[tool.uv]` — ortam ayarları
- `[tool.ruff]` — linting kuralları (seçilen kurallar ve gerekçeleri)
- `[tool.ruff.format]` — format ayarları
- `[tool.mypy]` — strict mode + per-module overrides
- `[tool.pytest.ini_options]` — test dizini, coverage ayarları
- `[tool.coverage.report]` — minimum coverage hedefi (örn. %80)

**`.pre-commit-config.yaml` İçeriği:**
```yaml
repos:
  - repo: https://github.com/astral-sh/ruff-pre-commit
    rev: v0.x.x
    hooks:
      - id: ruff
        args: [--fix]
      - id: ruff-format
  - repo: https://github.com/pre-commit/mirrors-mypy
    rev: v1.x.x
    hooks:
      - id: mypy
```

**`Makefile` Komutları:**
```makefile
install:     uv sync --all-extras && pre-commit install
dev:         uv run uvicorn ...  (REST API ise)
test:        uv run pytest --cov --cov-report=term-missing
lint:        uv run ruff check . && uv run ruff format --check .
typecheck:   uv run mypy src/
check:       make lint && make typecheck && make test
clean:       rm -rf .venv dist build *.egg-info .pytest_cache .mypy_cache
```

**`.env.example` Yapısı** — tüm ortam değişkenleri, değersiz, açıklamalı

**Geliştirici Workflow'u:** Yeni özellik → dal aç → kod → `make check` → commit → PR

---

### `TESTING_STRATEGY.md` *(Sadece Python projeleri için)*

- Test piramidi ve kapsam hedefleri:
  - Unit: %70 — dış bağımlılık yok, mock kullan
  - Integration: %20 — gerçek DB/servis, test container'ı
  - E2E: %10 — tam uçtan uca senaryo

- `conftest.py` fixture yapısı:
  ```python
  # Scope hiyerarşisi
  # session → module → function (varsayılan)
  @pytest.fixture(scope="session")
  def db_engine(): ...

  @pytest.fixture
  def db_session(db_engine): ...
  ```

- *(REST API ise)* `httpx.AsyncClient` ile endpoint testi örneği
- *(CLI ise)* `typer.testing.CliRunner` ile komut testi örneği
- *(ML ise)* Model validasyon testleri — metrik threshold kontrolleri
- Mock stratejisi: ne zaman mock, ne zaman gerçek bağımlılık
- Coverage raporlama ve minimum eşik kuralları

---

### `CLEAN_CODE_ARCHITECTURE.md`
- Projeye özgü klasör yapısı ve dosya organizasyonu
- Naming conventions *(değişken, fonksiyon, dosya, bileşen)*
- Seçilen mimari pattern ve uygulanma kuralları
- *(Python ise)* Naming conventions:
  - Modüller: `snake_case.py`
  - Sınıflar: `PascalCase`
  - Sabitler: `UPPER_SNAKE_CASE`
  - Private: `_leading_underscore`
  - Type aliases: `PascalCase` (örn. `UserId = NewType("UserId", int)`)
- *(Python ise)* Import sırası (ruff'ın `I` kuralları zaten zorlar ama gerekçe belirt)
- *(Python ise)* Tip annotation zorunlulukları:
  - Tüm public fonksiyon imzaları annotated olmalı
  - `Any` kullanımı yasak — gerekirse `# type: ignore[...]` ile gerekçe yaz
  - `Optional[X]` yerine `X | None` kullan (Python 3.10+)
- *(Oyun ise)* Assembly Definition yapısı
- Code Review checklist

---

### `MILESTONE_CHECKLISTS.md`

**Zorunlu Pratikler** bölümü *(her projede sabit — değiştirme)*:
```
- Her yeni şey denemeden önce KNOWLEDGE_BASE.md kontrol
- Her çalışan özellik sonrası Git commit
- Her sistem için ayrı test sahnesi / test ortamı
- Sayıları koda gömmeme (config / SO / SerializeField / env var kullan)
- Her milestone sonunda performans / bundle size kontrolü
- Phase 1 bitmeden yeni özellik ekleme
- Problem çözünce KNOWLEDGE_BASE → SOLVED kaydı
- Başarısız denemeler KNOWLEDGE_BASE → FAILED kaydı
- (Web) OnMount/useEffect temizlik fonksiyonlarını unutma
- (Oyun) OnEnable subscribe → OnDisable unsubscribe
- (Python) `make check` geçmeden commit yapma
- (Python) Her public fonksiyona type annotation + docstring
- (Python) Yeni bağımlılık eklemeden önce alternatiflerini KNOWLEDGE_BASE'e yaz
- (Python) `.env` dosyasını asla commit'leme — .env.example güncelle
- (Python) Coverage düşerse yeni kod kabul edilmez
```

**Phase 0 — Kurulum** (somut, ölçülebilir görevler)

*(Python ise Phase 0 özel görevler):*
- [ ] `uv` kurulumu ve `pyproject.toml` oluşturma
  - Kabul kriteri: *`uv sync` hatasız çalışır*
- [ ] `ruff` + `mypy` konfigürasyonu
  - Kabul kriteri: *`make lint && make typecheck` temiz çıktı verir*
- [ ] `pre-commit` kurulumu
  - Kabul kriteri: *`pre-commit run --all-files` geçer*
- [ ] İlk `pytest` çalıştırma (boş test)
  - Kabul kriteri: *`make test` hatasız çalışır, coverage raporu üretilir*
- [ ] `.env.example` ve `pydantic-settings` Config sınıfı
  - Kabul kriteri: *Tüm config ortam değişkenlerinden okunur, hardcoded değer yok*
- [ ] *(REST API ise)* FastAPI app skeleton + `/health` endpoint
  - Kabul kriteri: *`GET /health` → `{"status": "ok"}` döner*
- [ ] *(CI/CD ise)* GitHub Actions CI pipeline
  - Kabul kriteri: *PR açıldığında lint + typecheck + test otomatik çalışır*
- [ ] *(Docker ise)* Multi-stage Dockerfile
  - Kabul kriteri: *`docker build` başarılı, image boyutu makul (<200MB)*

**Phase 1 — MVP** (somut, ölçülebilir görevler)
**Phase 2+ — Sonraki Aşamalar**

Her görev:
- [ ] Açık tanım
- Kabul kriteri: *"Bitti sayılır: ..."*

---

### `KNOWLEDGE_BASE.md`
- Kullanım açıklaması ve kayıt formatı
- **ADR Kayıtları** *(toplanan bilgilerden minimum 3 — proje başından itibaren)*:
  - Motor / Framework seçimi neden?
  - Mimari pattern neden?
  - Veri yönetimi / DB seçimi neden?
  - *(Python ise ek zorunlu ADR'lar):*
    - Neden `uv`? *(pip/poetry/conda'ya karşı)*
    - Neden `ruff`? *(flake8+black+isort kombinasyonuna karşı)*
    - Neden `FastAPI` / `typer` / seçilen framework? *(alternatiflerine karşı)*
- **RECIPES:** Projeye özel minimum 3 kod şablonu
  - *(Web ise: API route, Supabase query, Server Component örneği)*
  - *(Oyun ise: StateMachine, EventBus, ScriptableObject şablonu)*
  - *(Python — REST API ise: FastAPI route, repository pattern, pydantic şema örneği)*
  - *(Python — CLI ise: typer command, rich output, hata yönetimi örneği)*
  - *(Python — ML ise: veri yükleme, model eğitim, metrik kayıt şablonu)*
  - *(Python — Library ise: public API tasarımı, __init__.py export yapısı)*

---

### `VERSION_CONTROL_SETUP.md`
- Git + Git LFS kurulumu *(Oyun ise LFS zorunlu, Python — ML ise data/ için LFS önerilir)*
- Projeye özgü `.gitignore`
  - *(Python ise ek girişler):*
    ```
    .venv/
    __pycache__/
    *.pyc
    .mypy_cache/
    .pytest_cache/
    .ruff_cache/
    dist/
    build/
    *.egg-info/
    .env
    data/raw/     # büyük veri dosyaları
    notebooks/.ipynb_checkpoints/
    ```
- Branch stratejisi: `main` / `dev` / `feature/*` / `hotfix/*`
- Commit mesajı kuralları *(Conventional Commits)*
- *(Python — Library ise)* Tag stratejisi: `v1.2.3` SemVer tag'leri → PyPI release tetikler

---

### `CHARACTER_PIPELINE.md` *(Sadece 3D veya 2D sprite karakter içeriyorsa)*
- Karakter üretim ve import akışı
- Rig ve animasyon standartları
- Naming ve dosya organizasyonu kuralları

> ⚠️ Bu dosya yalnızca projede karakter asset pipeline'ı gerektiren bir yapı varsa oluşturulur. Yoksa atla.

---

## 🤖 AŞAMA 4 — WORKFLOW DOSYALARINI OLUŞTUR

### `.agent/workflows/ctx.md`

```markdown
---
description: Load full project context at the start of a new session
---

## Adım 1 — Proje Genel Bakış
[ProjeKökü]/Documentation/PROJECT_OVERVIEW.md dosyasını oku.

## Adım 2 — Neredeyiz?
[ProjeKökü]/Documentation/MILESTONE_CHECKLISTS.md dosyasını oku.
Progress Summary bölümünü güncelle.

## Adım 3 — Geçmiş & Kararlar
[ProjeKökü]/Documentation/KNOWLEDGE_BASE.md dosyasını oku.

## Adım 4 — Durum Raporu

📍 PROJE DURUMU
─────────────────────────────────
Phase   : [aktif phase]
Aktif   : [milestone adı]
Son ✅  : [son tamamlanan görev]
Sonraki : [sıradaki görev]
Blokaj  : [engelleyen durum / —]
─────────────────────────────────
Ne yapmak istiyorsun?

## ⚠️ Zorunlu Kurallar (Her oturumda hatırlat)
[MILESTONE_CHECKLISTS.md'deki Zorunlu Pratikler listesi buraya kopyalanır]
```

---

### `.agent/workflows/done.md`

```markdown
---
description: Mark current milestone task as complete, get next steps and professional feedback
---

## Adım 1 — Tamamlanan görevi işaretle
MILESTONE_CHECKLISTS.md dosyasında ilgili görevi [x] yap.
Progress Summary bölümünü güncelle.

## Adım 2 — KNOWLEDGE_BASE güncelle
Gerekiyorsa SOLVED / FAILED / yeni ADR kaydı ekle.

## Adım 3 — Tamamlama raporu ver
- Ne yapıldı (özet)
- Pro tip: Bu aşamada sık yapılan hata veya iyileştirme önerisi
- Sıradaki görev

## Adım 4 — Geliştirme sorusu sor
Bu görevi daha da ilerletmek ister misin, yoksa sıradaki göreve geçelim mi?

## Adım 5 — Milestone kutlaması
Eğer tüm milestone görevleri tamamlandıysa kutla ve bir sonraki milestone'u tanıt.
```

---

## ✅ AŞAMA 5 — KURULUM ÖZETİ (Dashboard)

```
🚀 KURULUM TAMAMLANDI: [Proje Adı]
════════════════════════════════════════
📁 Documentation/
   ✅ PROJECT_OVERVIEW.md
   ✅ PRODUCT_SPEC_OR_GDD.md
   ✅ TECHNICAL_DESIGN_DOC.md
   ✅ DATABASE_SCHEMA.md              [Web/Cloud/Python+DB ise]
   ✅ PYTHON_PROJECT_SETUP.md         [Python ise]
   ✅ TESTING_STRATEGY.md             [Python ise]
   ✅ CLEAN_CODE_ARCHITECTURE.md
   ✅ MILESTONE_CHECKLISTS.md
   ✅ KNOWLEDGE_BASE.md
   ✅ VERSION_CONTROL_SETUP.md
   ✅ CHARACTER_PIPELINE.md           [Karakter pipeline varsa]

⚙️  .agent/workflows/
   ✅ ctx.md   — Oturum başlangıcı
   ✅ done.md  — Görev tamamlama

🐍  Python Proje Dosyaları (kök dizin)
   ✅ pyproject.toml
   ✅ .python-version
   ✅ .pre-commit-config.yaml
   ✅ .env.example
   ✅ Makefile
   ✅ Dockerfile                       [Docker ise]
   ✅ .github/workflows/ci.yml         [CI/CD ise]
   ✅ .github/workflows/release.yml    [Library + PyPI ise]

📌 KULLANIM:
   Yeni oturum    → /ctx
   Görev bitti    → /done
   Yeni proje     → /setup-project

🐍  Python geliştirici hızlı başlangıç:
   Kurulum        → make install
   Geliştirme     → make dev
   Tüm kontroller → make check
════════════════════════════════════════
Phase 0'dan başlayalım mı?
```

---

## 🔍 KALİTE KONTROL LİSTESİ

Her dosya oluşturulduktan sonra kontrol et:

- [ ] İçerik gerçekten **projeye özel** mi? (Genel şablon dil yok)
- [ ] Dosya içi referanslar ve linkler doğru mu?
- [ ] `MILESTONE_CHECKLISTS.md`: Phase 1 görevleri somut ve ölçülebilir mi?
- [ ] `KNOWLEDGE_BASE.md`: En az **3 ADR** kaydı var mı? *(Python ise 5+)*
- [ ] `DATABASE_SCHEMA.md`: RLS politikaları tanımlanmış mı? *(Web ise)*
- [ ] `DATABASE_SCHEMA.md`: Alembic migration stratejisi var mı? *(Python + DB ise)*
- [ ] `ctx.md` ve `done.md`: Dosya yolları bu projeye özgü ve doğru mu?
- [ ] `PRODUCT_SPEC_OR_GDD.md`: Türe göre doğru format seçildi mi?
- [ ] `CHARACTER_PIPELINE.md`: Gereksiz yere oluşturulmadı mı?
- [ ] *(Python ise)* `PYTHON_PROJECT_SETUP.md`: `pyproject.toml` içeriği alt türe uygun mu?
- [ ] *(Python ise)* `TESTING_STRATEGY.md`: Test türleri ve coverage hedefi tanımlı mı?
- [ ] *(Python ise)* `Makefile`: `make check` komutu lint + typecheck + test'i zincirlediğinde hatasız çalışıyor mu?
- [ ] *(Python ise)* `.env.example`: Tüm ortam değişkenleri mevcut mu, gerçek değer yok mu?
- [ ] *(Python — Library ise)* PyPI publish pipeline tanımlanmış mı?
- [ ] *(Python — ML ise)* `notebooks/` ve `data/` dizinleri oluşturuldu mu?
