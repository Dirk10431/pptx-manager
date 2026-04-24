@echo off
REM =============================================================
REM pptx-manager - Start-Skript fuer Windows
REM Doppelklick auf diese Datei startet die App.
REM =============================================================

cd /d "%~dp0"

echo.
echo  =========================================
echo   pptx-manager wird gestartet...
echo  =========================================
echo.

REM Abhaengigkeiten pruefen
if not exist "node_modules\" (
    echo Erstinstallation: npm install laeuft...
    call npm install
    if errorlevel 1 (
        echo.
        echo FEHLER: npm install fehlgeschlagen.
        pause
        exit /b 1
    )
)

REM Browser oeffnen nach 2 Sekunden
start "" cmd /c "timeout /t 2 >nul && start http://127.0.0.1:3002"

REM Server starten (bleibt im Vordergrund)
node server.js

pause
