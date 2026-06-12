@echo off
setlocal enableextensions
cd /d "%~dp0"

set "VENV=.venv"
set "PY=%VENV%\Scripts\python.exe"
set "STAMP=%VENV%\pyproject.stamp"
set "NEED_INSTALL="

rem --- venv anlegen, falls nicht vorhanden ---
if not exist "%PY%" (
    echo [exceltool] Erstelle virtuelle Umgebung .venv ...
    py -3 -m venv "%VENV%" 2>nul || python -m venv "%VENV%" || goto :err
    set "NEED_INSTALL=1"
)

rem --- Neuinstallation nur, wenn pyproject.toml sich geaendert hat ---
if not exist "%STAMP%" (
    set "NEED_INSTALL=1"
) else (
    fc /b "pyproject.toml" "%STAMP%" >nul 2>&1 || set "NEED_INSTALL=1"
)

if defined NEED_INSTALL (
    echo [exceltool] Installiere Abhaengigkeiten aus pyproject.toml ...
    "%PY%" -m pip install --upgrade pip >nul
    "%PY%" -m pip install -e . || goto :err
    copy /y "pyproject.toml" "%STAMP%" >nul
)

echo [exceltool] Starte ...
"%PY%" exceltool.py
goto :end

:err
echo.
echo [exceltool] FEHLER: Einrichtung oder Start fehlgeschlagen.
pause
exit /b 1

:end
exit /b 0
