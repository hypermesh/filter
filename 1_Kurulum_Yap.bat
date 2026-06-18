@echo off
chcp 65001 > nul
set PYTHONIOENCODING=utf-8
title Proje Kurulum Arayuzu
color 0E

echo ========================================================
echo        MAKINE RECETE OTOMASYONU - GEREKSINIM KURULUMU
echo ========================================================
echo.
cd /d "%~dp0"

:: 1. Python kontrolu
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [HATA] Bilgisayarda Python bulunamadi!
    echo Lutfen once Python kurun ve kurulum esnasinda
    echo "Add Python to PATH" secenegini isaretleyin.
    echo.
    pause
    exit /b 1
)

:: 2. uv kontrolu
where uv >nul 2>nul
if %errorlevel% equ 0 (
    echo [BILGI] Bilgisayarda 'uv' bulundu.
    echo Paketler 'uv' kullanilarak kuruluyor...
    echo.
    uv pip install -e .
    goto SUCCESS
) else (
    echo [BILGI] 'uv' paket yoneticisi bulunamadi.
    echo Standart Python sanal ortami venv kurulumu deneniyor...
    echo.
)

:: 3. .venv kontrolu ve olusturma
if not exist ".\.venv\Scripts\python.exe" (
    echo [BILGI] Yerel sanal ortam venv bulunamadi. Olusturuluyor...
    python -m venv .venv
    if %errorlevel% neq 0 (
        echo [HATA] Sanal ortam venv olusturulamadi!
        pause
        exit /b 1
    )
    echo [BASKI] Sanal ortam venv basariyla olusturuldu.
)

:: 4. venv icine paketleri kurma
echo [BILGI] Kutuphaneler kuruluyor veya guncelleniyor.
echo Lutfen bekleyin, bu islem birkac dakika surebilir...
echo.
.\.venv\Scripts\pip install --upgrade pip
.\.venv\Scripts\pip install pandas openpyxl typer rich pydantic-settings

:SUCCESS
echo.
echo ========================================================
echo [BASARILI] Kurulum tamamlandi! Proje calismaya hazir.
echo Artik diger bat dosyalarini guvenle kullanabilirsiniz.
echo ========================================================
echo.
pause
