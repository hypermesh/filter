@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"

echo ===================================================
echo   ONCELIK SIRASI GUNCELLEME ARACI
echo ===================================================
echo.

:: Hangi txt dosyasini kullanacagiz?
if "%~1"=="" (
    set TXT_PATH=%~dp0oncelik_sirasi.txt
) else (
    set TXT_PATH=%~1
)

:: Dosya yoksa olustur
if not exist "!TXT_PATH!" (
    echo Dosya bulunamadi, ornek olusturuluyor: !TXT_PATH!
    echo # Her satira tek bir kod yazin > "!TXT_PATH!"
    echo # En ustteki kod en yuksek oncelikte olur >> "!TXT_PATH!"
    echo 2254 >> "!TXT_PATH!"
    echo 2235 >> "!TXT_PATH!"
)

echo Oncelik dosyasi: !TXT_PATH!
for %%F in ("!TXT_PATH!") do echo Son kayit: %%~tF
echo.

:: Dosyayi Notepad ile ac, kullanici kapatana kadar bekle
echo [ADIM 1] Dosya Notepad ile aciliyor - degisikliklerinizi yapip KAYDEDIN ve KAPATIN...
echo.
notepad "!TXT_PATH!"

:: Notepad kapandiktan sonra otomatik devam et
echo [ADIM 2] Notepad kapatildi. JSON guncelleniyor...
echo.

for %%F in ("!TXT_PATH!") do echo Yeni kayit zamani: %%~tF
echo.

where uv >nul 2>nul
if %errorlevel% equ 0 (
    uv run python src/recipe_automation/utils/update_priority.py "!TXT_PATH!"
) else (
    if exist ".\.venv\Scripts\python.exe" (
        .\.venv\Scripts\python.exe src/recipe_automation/utils/update_priority.py "!TXT_PATH!"
    ) else (
        python src/recipe_automation/utils/update_priority.py "!TXT_PATH!"
    )
)

echo.
echo ===================================================
pause
