# Test Stratejisi

## Kapsam ve Hedefler
Projede Pytest kullanılmaktadır ve hedef coverage (kapsama alanı) minimum %70 olmalıdır. CLI aracı olduğundan ve veriler Excel formatında işlendiğinden testler hem fonksiyonel hem entegrasyon seviyesinde olmalıdır.

### 1. Unit Testler (Birim Testleri)
Dış bağımlılığı (dosya okuma, yazma vb.) simüle ederek (mocking) sadece DataFrame filtreleme mantığını (örneğin stok filtreleme fonksiyonunu veya ID filtreleme fonksiyonunu) test eder.

### 2. Integration Testler (Entegrasyon Testleri)
Belirli senaryolar için hazırlanmış fiziksel test excel dosyaları (örn. `tests/data/test_verisi.xlsx`) üzerinden tüm pipeline'ın çalışıp, doğru Excel dosyasını üretip üretmediği kontrol edilir.

### 3. CLI Testleri
`typer.testing.CliRunner` kullanılarak terminal komutlarının (örn. `uv run recipe-automation scan ...`) doğru argümanlarla çalışıp çalışmadığı, stdout'a (konsola) doğru mesajları yazıp yazmadığı doğrulanır.
