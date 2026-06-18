@echo off
chcp 65001 > nul
set PYTHONIOENCODING=utf-8
title Recete Agaci Canli Takip Sunucusu
color 0A

echo ========================================================
echo       RECETE AGACI CANLI TAKIP VE IZLEME SERVISI
echo ========================================================
echo.

cd /d "%~dp0"

where uv >nul 2>nul
if %errorlevel% equ 0 (
    if "%~1"=="" (
        echo [BILGI] Excel surukle-birak yapilmadi.
        echo En son olusturulan Filtrelenmis Excel otomatik bulunuyor...
        uv run recipe-automation agac --watch
    ) else (
        echo [BILGI] Excel: %~1
        uv run recipe-automation agac "%~1" --watch
    )
) else (
    if exist ".\.venv\Scripts\python.exe" (
        if "%~1"=="" (
            echo [BILGI] Excel surukle-birak yapilmadi.
            echo En son olusturulan Filtrelenmis Excel otomatik bulunuyor...
            .\.venv\Scripts\python.exe src\recipe_automation\main.py agac --watch
        ) else (
            echo [BILGI] Excel: %~1
            .\.venv\Scripts\python.exe src\recipe_automation\main.py agac "%~1" --watch
        )
    ) else (
        echo [UYARI] uv veya yerel sanal ortam venv bulunamadi.
        echo Otomatik kurulum baslatiliyor...
        echo.
        call 1_Kurulum_Yap.bat
        if exist ".\.venv\Scripts\python.exe" (
            echo.
            echo [BILGI] Kurulum tamamlandi. Islem devam ediyor...
            echo.
            if "%~1"=="" (
                .\.venv\Scripts\python.exe src\recipe_automation\main.py agac --watch
            ) else (
                .\.venv\Scripts\python.exe src\recipe_automation\main.py agac "%~1" --watch
            )
        ) else (
            echo [HATA] Kurulum basarisiz oldu.
            pause
            exit /b 1
        )
    )
)

pause
