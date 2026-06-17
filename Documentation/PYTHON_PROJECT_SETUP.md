# Python Projesi Teknik Kurulum Rehberi

Bu proje, modern Python standartları kullanılarak yapılandırılmıştır.

## Ortam Kurulumu
Projede paket yöneticisi olarak `uv` kullanılmaktadır.

### 1. uv Yüklenmesi
Eğer sisteminizde yüklü değilse, PowerShell veya CMD üzerinden:
```bash
# Windows
powershell -c "irm https://astral.sh/uv/install.ps1 | iex"
```

### 2. Projenin Başlatılması
Proje dizininde (C:\Users\Mustafa\Desktop\Filter) terminal açıp aşağıdaki komutu girin:
```bash
make install
```
Eğer Windows'ta Make kurulu değilse komutları doğrudan çalıştırabilirsiniz:
```bash
uv sync --all-extras
uv run pre-commit install
```

### 3. Uygulamanın Çalıştırılması
Sanal ortam (virtual environment) otomatik olarak oluşturulur. Aracı çalıştırmak için:
```bash
uv run recipe-automation --help
```

## Geliştirici Komutları (`Makefile` alternatifleri)
- Linter çalıştır ve düzelt: `uv run ruff check . --fix` ve `uv run ruff format .`
- Tip kontrolü: `uv run mypy src/`
- Testler: `uv run pytest --cov=src`
