@echo off
chcp 65001 > nul
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"
set PYTHONPATH=src
color 0B

echo ----------------------------------------------------
echo [AKILLI FILTRE] Lutfen islemek istediginiz
echo Excel dosyasini veya klasorunu bu dosyanin uzerine
echo SURUKLEYIP BIRAKIN.
echo ----------------------------------------------------

if "%~1"=="" (
    echo.
    echo [HATA] Surukle-Birak yapmadiniz!
    pause
    exit /b
)

echo Dosyalar otomatik analiz ediliyor...
where uv >nul 2>nul
if %errorlevel% equ 0 (
    uv run recipe-automation auto-filter "%~1" --group TIM
) else (
    if exist ".\.venv\Scripts\python.exe" (
        .\.venv\Scripts\python.exe src\recipe_automation\main.py auto-filter "%~1" --group TIM
    ) else (
        echo [UYARI] uv veya yerel sanal ortam venv bulunamadi
        echo Otomatik kurulum baslatiliyor...
        echo.
        call 1_Kurulum_Yap.bat
        if exist ".\.venv\Scripts\python.exe" (
            echo.
            echo [BILGI] Kurulum basariyla tamamlandi. Islem devam ediyor...
            echo.
            .\.venv\Scripts\python.exe src\recipe_automation\main.py auto-filter "%~1" --group TIM
        ) else (
            echo [HATA] Kurulum basarisiz oldu
            pause
            exit /b 1
        )
    )
)
pause


