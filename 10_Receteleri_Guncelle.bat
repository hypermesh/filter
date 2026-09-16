@echo off
chcp 65001 > nul
title RECETE VERITABANI GUNCELLEME SIHIRBAZI

echo ====================================================================
echo        HYPERMESH RECIPE & PART DATABASE AUTO-SYNC ENGINE
echo ====================================================================
echo.

if not "%~1"=="" (
    echo [BILGI] Suruklenen dosya/klasor isleniyor: "%~1"
    python -m recipe_automation.services.recipe_database_builder "%~1"
) else (
    if exist "Receteler" (
        echo [BILGI] "Receteler" klasorundeki tum Excel dosyalari isleniyor...
        python -m recipe_automation.services.recipe_database_builder "Receteler"
    ) else if exist "test.xlsx" (
        echo [BILGI] test.xlsx dosyasi isleniyor...
        python -m recipe_automation.services.recipe_database_builder "test.xlsx"
    ) else (
        echo [UYARI] Receteler klasoru veya test.xlsx bulunamadi!
        echo Lutfen Excel dosyanizi bu .bat dosyasinin uzerine surukleyip birakin.
    )
)

echo.
echo ====================================================================
echo  Islem tamamlandi! Veritabanlari basariyla guncellendi.
echo ====================================================================
pause
