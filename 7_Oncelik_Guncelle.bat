@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion
set PYTHONIOENCODING=utf-8

echo ===================================================
echo   ONCELIK SIRASI GUNCELLEME ARACI
echo ===================================================
echo.
if "%~1"=="" (
    echo 'oncelik_sirasi.txt' dosyasindaki listeye gore
    echo 'veritabanlari/oncelik_sirasi.json' guncelleniyor...
) else (
    echo Suruklenen dosya: '%~1'
    echo Bu dosya okunarak 'veritabanlari/oncelik_sirasi.json' guncelleniyor...
)
echo.

where uv >nul 2>nul
if %errorlevel% equ 0 (
    uv run python src/recipe_automation/utils/update_priority.py %1
) else (
    if exist ".\.venv\Scripts\python.exe" (
        .\.venv\Scripts\python.exe src/recipe_automation/utils/update_priority.py %1
    ) else (
        python src/recipe_automation/utils/update_priority.py %1
    )
)

echo.
echo ===================================================
pause
