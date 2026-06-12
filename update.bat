@echo off
setlocal enableextensions
cd /d "%~dp0"

echo [exceltool] git pull ...
git pull || (
    echo [exceltool] FEHLER: git pull fehlgeschlagen.
    pause
    exit /b 1
)

echo [exceltool] Aktualisierung abgeschlossen.
pause
exit /b 0
