@echo off
REM =============================================================
REM make-shortcuts.bat - generiert pptx-thumbs.bat / pptx-open.bat
REM mit hardcoded Pfad fuer diesen Projekt-Ordner.
REM Doppelklick reicht.
REM =============================================================

cd /d "%~dp0"

echo.
echo Generiere Shortcut-Batches fuer pptx-manager...
echo.

node bin\generate-shortcuts.js
if errorlevel 1 (
    echo.
    echo FEHLER: Generator konnte nicht ausgefuehrt werden.
    echo Pruefen, ob node installiert ist: "node --version"
    pause
    exit /b 1
)

echo.
echo Fertig. Beliebige Taste zum Schliessen.
pause >nul
