# Versiyon Kontrolü Kurulumu

Projeyi Git ile yönetirken aşağıdaki standartlara uyulacaktır.

## .gitignore Örneği
Excel dosyalarının (hassas veri içerebileceğinden) ve derlenmiş Python önbelleklerinin depoya (repository) gitmesini önlemek için:

```text
# Python Özel Klasörler
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

# Kullanıcı Verisi
data/
*.xlsx
~$*.xlsx
operasyon_log.txt

# Ancak Test için kullanılan örnekleri yoksayma:
!tests/data/*.xlsx
```

## Commit Mesajı Kuralları (Conventional Commits)
- `feat: [açıklama]` -> Yeni bir araç eklendiğinde.
- `fix: [açıklama]` -> Excel okurken oluşan vb. bir hata giderildiğinde.
- `docs: [açıklama]` -> README vb. dökümanlarda değişiklik yapıldığında.
- `refactor: [açıklama]` -> Fonksiyonların kod okunuşunu iyileştirme amaçlı düzenlendiğinde.
