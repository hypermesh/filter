@echo off
chcp 65001 > nul
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"
set PYTHONPATH=src
color 0E
title Excel Birlestirme Otomasyonu

echo ========================================================
echo         EXCEL BIRLESTIRME VE LOG OTOMASYONU
echo ========================================================
echo.
echo [TALIMAT] Lutfen birlestirmek istediginiz Excel
echo dosyalarini barindiran KLASORU bu bat dosyasinin
echo uzerine surukleyip birakin (Drag-and-Drop).
echo ========================================================
echo.

if "%~1"=="" (
    echo.
    echo [HATA] Herhangi bir klasor surukleyip birakmadiniz!
    echo Lutfen Excel dosyalarinin bulundugu bir klasoru surukleyin.
    echo.
    pause
    exit /b 1
)

echo Hedef Klasor: "%~1"
echo.

where uv >nul 2>nul
if %errorlevel% equ 0 (
    uv run python src\recipe_automation\excel_merger.py "%~1"
) else (
    if exist ".\.venv\Scripts\python.exe" (
        .\.venv\Scripts\python.exe src\recipe_automation\excel_merger.py "%~1"
    ) else (
        echo [UYARI] uv veya yerel sanal ortam venv bulunamadi.
        echo Sistemdeki genel Python deneniyor...
        echo.
        python src\recipe_automation\excel_merger.py "%~1"
    )
)

echo.
echo ========================================================
echo Islem tamamlandi.
echo ========================================================
echo.
pause
