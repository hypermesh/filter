@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0"
color 0A
title Interaktif Uretim ve Montaj Takip Paneli

echo ====================================================
echo   INTERAKTIF URETIM VE MONTAJ TAKIP PANELİ
echo   Tarayici uzerinden cok hizli ve canli veri girisi
echo   ile Excel formullerinin anlik simulasyonu.
echo ====================================================
echo.

echo [BILGI] Takip Paneli sunucusu baslatiliyor...
echo [BILGI] Sunucu Adresi: http://localhost:8000
echo.
echo [IPUCU] Panel acildiginda olusturulan "Filtered_TIM_Haziran"
echo         dosyasini tarayiciya surukleyip birakarak baslayin.
echo.
echo Kapatmak icin bu pencereyi kapatabilir veya Ctrl+C tuslarina basabilirsiniz.
echo.

:: Tarayiciyi arka planda ac
start http://localhost:8000

:: Python basit HTTP sunucusunu baslat
if exist ".\.venv\Scripts\python.exe" (
    .\.venv\Scripts\python.exe -m http.server 8000 --directory dashboard
) else (
    python -m http.server 8000 --directory dashboard
)

if %errorlevel% neq 0 (
    echo [HATA] Sunucu baslatilamadi. Python yuklu oldugundan emin olun.
    pause
)
