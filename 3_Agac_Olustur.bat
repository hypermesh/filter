@echo off
chcp 65001 > nul
set PYTHONIOENCODING=utf-8
title Recete Agaci Olustur

if "%~1"=="" (
    echo ----------------------------------------------------
    echo [HATA] Lutfen islemek istediginiz Excel dosyasini
    echo veya recete klasorunu bu bat dosyasinin uzerine
    echo SURUKLEYIP BIRAKIN.
    echo ----------------------------------------------------
    pause
    exit /b
)

echo ==============================================
if exist "%~1\" (
    echo ISLEM: RECETE AGACI OLUSTUR ^(KLASOR MODU^)
) else (
    echo ISLEM: RECETE AGACI OLUSTUR
)
echo HEDEF: %~1
echo ==============================================
echo.

cd /d "%~dp0"
where uv >nul 2>nul
if %errorlevel% equ 0 (
    uv run recipe-automation agac "%~1"
) else (
    if exist ".\.venv\Scripts\python.exe" (
        .\.venv\Scripts\python.exe src\recipe_automation\main.py agac "%~1"
    ) else (
        echo [UYARI] uv veya yerel sanal ortam venv bulunamadi
        echo Otomatik kurulum baslatiliyor...
        echo.
        call 1_Kurulum_Yap.bat
        if exist ".\.venv\Scripts\python.exe" (
            echo.
            echo [BILGI] Kurulum basariyla tamamlandi. Islem devam ediyor...
            echo.
            .\.venv\Scripts\python.exe src\recipe_automation\main.py agac "%~1"
        ) else (
            echo [HATA] Kurulum basarisiz oldu
            pause
            exit /b 1
        )
    )
)

echo.
pause

