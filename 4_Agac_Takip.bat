@echo off
chcp 65001 > nul
set PYTHONIOENCODING=utf-8
title Re�ete A�ac� Canl� Takip Sunucusu
color 0A

echo ========================================================
echo       RE�ETE A�ACI CANLI TAK�P VE �ZLEME SERV�S�
echo ========================================================
echo.

cd /d "%~dp0"

where uv >nul 2>nul
if %errorlevel% equ 0 (
    if "%~1"=="" (
        echo [B�LG�] Excel s�r�klenmedi. En son �retilen Filtrelenmi� Excel bulunuyor...
        uv run python src\recipe_automation\w
