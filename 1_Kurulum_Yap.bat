@echo off
chcp 65001 > nul
set PYTHONIOENCODING=utf-8
title Proje Kurulum Arayüzü
color 0E

echo ========================================================
echo        MAKİNE REÇETE OTOMASYONU - GEREKSİNİM KURULUMU
echo ========================================================
echo.

cd /d "%~dp0"

:: 1. Python kontrolü
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [HATA] Bilgisayarda Python bulunamadı!
    echo Lütfen önce Python kurun ve kurulum esnasında
    echo "Add Python to PATH" seçeneğini işaretleyin.
    echo.
    pause
    exit /b 1
)

:: 2. uv kontrolü
where uv >nul 2>nul
if %errorlevel% equ 0 (
    echo [BİLGİ] Bilgisayarda 'uv' bulundu.
    echo Paketler 'uv' kullanılarak kuruluyor/senkronize ediliyor...
    echo.
    uv pip install -e .
    goto SUCCESS
) else (
    echo [BİLGİ] 'uv' paket yöneticisi bulunamadı.
    echo Standart Python sanal ortamı (.venv) kurulumu deneniyor...
    echo.
)

:: 3. .venv kontrolü ve oluşturma
if not exist ".\.venv\Scripts\python.exe" (
    echo [BİLGİ] Yerel sanal ortam (.venv) bulunamadı. Oluşturuluyor...
    python -m venv .venv
    if %errorlevel% neq 0 (
        echo [HATA] Sanal ortam (.venv) oluşturulamadı!
        pause
        exit /b 1
    )
    echo [BASKI] Sanal ortam (.venv) başarıyla oluşturuldu.
)

:: 4. venv içine paketleri kurma
echo [BİLGİ] Kütüphaneler kuruluyor veya güncelleniyor.
echo Lütfen bekleyin, bu işlem internet hızınıza bağlı olarak birkaç dakika sürebilir...
echo.
.\.venv\Scripts\pip install --upgrade pip
.\.venv\Scripts\pip install pandas openpyxl typer rich pydantic-settings

:SUCCESS
echo.
echo ========================================================
echo [BAŞARILI] Kurulum tamamlandı! Proje çalışmaya hazır.
echo Artık diğer bat dosyalarını güvenle kullanabilirsiniz.
echo ========================================================
echo.
pause

