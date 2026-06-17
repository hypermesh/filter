@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

echo ===================================================
echo   ÖNCELİK SIRASI GÜNCELLEME ARACI
echo ===================================================
echo.
if "%~1"=="" (
    echo 'oncelik_sirasi.txt' dosyasındaki listeye göre 
    echo 'veritabanlari/oncelik_sirasi.json' güncelleniyor...
) else (
    echo Sürüklenen dosya: '%~1'
    echo Bu dosya okunarak 'veritabanlari/oncelik_sirasi.json' güncelleniyor...
)
echo.

python src/recipe_automation/utils/update_priority.py %1

echo.
echo ===================================================
pause
