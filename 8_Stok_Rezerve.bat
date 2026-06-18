@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0"
set PYTHONPATH=src
color 0B
title Stok Rezervasyon Sistemi

echo ====================================================
echo   STOK REZERVASYON SİSTEMİ
echo   Klasoru bu dosyanin uzerine surukleme-birakma ile
echo   oncelik sirasina gore stok dagitimi yapilir.
echo ====================================================
echo.

if "%~1"=="" (
    echo [HATA] Lutfen islemek istediginiz klasoru bu
    echo        dosyanin uzerine SURUKLEYIP BIRAKIN.
    echo.
    pause
    exit /b
)

echo Klasor: %~1
echo.
echo Stok rezervasyonu baslatiliyor...
echo.

where uv >nul 2>nul
if %errorlevel% equ 0 (
    uv run recipe-automation reserve "%~1"
) else (
    if exist ".\.venv\Scripts\python.exe" (
        .\.venv\Scripts\python.exe src\recipe_automation\main.py reserve "%~1"
    ) else (
        echo [UYARI] uv veya yerel sanal ortam bulunamadi.
        echo Otomatik kurulum baslatiliyor...
        echo.
        call 1_Kurulum_Yap.bat
        if exist ".\.venv\Scripts\python.exe" (
            echo.
            echo [BILGI] Kurulum tamamlandi. Islem devam ediyor...
            echo.
            .\.venv\Scripts\python.exe src\recipe_automation\main.py reserve "%~1"
        ) else (
            echo [HATA] Kurulum basarisiz.
            pause
            exit /b 1
        )
    )
)

echo.
pause
