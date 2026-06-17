@echo off
chcp 65001 >nul
title Harici Analiz (Kapasite) Filtresi
echo =======================================================
echo          HARICI ANALIZ VE KAPASITE FILTRESI
echo =======================================================
echo.
echo Modul yukleniyor...
echo.

set PYTHONPATH=%~dp0src
where uv >nul 2>nul
if %errorlevel% equ 0 (
    uv run python "%~dp0src\recipe_automation\analysis.py" %*
) else (
    if exist "%~dp0.venv\Scripts\python.exe" (
        "%~dp0.venv\Scripts\python.exe" "%~dp0src\recipe_automation\analysis.py" %*
    ) else (
        python "%~dp0src\recipe_automation\analysis.py" %*
    )
)

pause

