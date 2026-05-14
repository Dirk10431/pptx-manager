#!/usr/bin/env node
// =============================================================
// bin/generate-shortcuts.js
// =============================================================
// Erzeugt 2 Batch-Dateien im Projekt-Root mit *hardcoded* Pfad.
// Damit funktionieren die .bat-Dateien auch dann, wenn man sie
// nach %USERPROFILE%\Desktop\ zieht oder ins Startmenue verschiebt.
//
// Erzeugte Dateien:
//   pptx-thumbs.bat  — startet die Thumbnail-Generierung
//   pptx-open.bat    — startet Server im Hintergrund (falls noetig)
//                      und oeffnet den Browser auf 127.0.0.1:3002
//
// Aufruf:
//   node bin/generate-shortcuts.js
// oder per package.json-Script:
//   npm run shortcuts
// oder per Doppelklick auf:
//   make-shortcuts.bat (im Projekt-Root)
// =============================================================

const fs = require('fs');
const path = require('path');

const projectRoot = path.resolve(__dirname, '..');

// .bat-Dateien werden mit cp1252 / OEM-konform gelesen — wir bleiben
// ASCII-sicher (keine Umlaute), damit es egal ist, ob Notepad das
// als UTF-8 oder ANSI speichert.
const thumbsBat =
`@echo off
title pptx-manager - Thumbnails
cd /d "${projectRoot}"
echo ==========================================
echo  pptx-manager - Thumbnail-Generierung
echo ==========================================
echo  Strg+C bricht ab; beim naechsten Start
echo  macht der Lauf an der Stelle weiter.
echo.
call npm run thumbs
echo.
echo Fertig. Beliebige Taste zum Schliessen.
pause >nul
`;

const openBat =
`@echo off
title pptx-manager - Launcher

REM Pruefen, ob Port 3002 schon belegt ist (= Server laeuft schon)
netstat -ano | findstr ":3002 " | findstr "LISTENING" >nul
if errorlevel 1 (
    REM Server laeuft nicht — Erstcheck Dependencies
    if not exist "${projectRoot}\\node_modules\\" (
        echo.
        echo node_modules fehlen. Bitte einmalig "start.bat" im Projekt-Ordner ausfuehren:
        echo    ${projectRoot}\\start.bat
        echo Beim ersten Start wird "npm install" automatisch ausgefuehrt.
        echo.
        pause
        exit /b 1
    )
    echo Server nicht aktiv. Starte im Hintergrund...
    start "pptx-manager Server" /min /D "${projectRoot}" node server.js

    REM Warten bis Server lauscht (max ca. 20s)
    set "READY="
    for /l %%i in (1,1,20) do (
        if not defined READY (
            timeout /t 1 /nobreak >nul
            netstat -ano | findstr ":3002 " | findstr "LISTENING" >nul && set "READY=1"
        )
    )
)

REM Browser oeffnen
start "" "http://127.0.0.1:3002"
exit
`;

const targets = [
    { name: 'pptx-thumbs.bat', content: thumbsBat },
    { name: 'pptx-open.bat',   content: openBat   },
];

console.log('Generiere Shortcut-Batches mit hardcoded Pfad:');
console.log('  ' + projectRoot);
console.log('');

for (const t of targets) {
    const outPath = path.join(projectRoot, t.name);
    fs.writeFileSync(outPath, t.content, { encoding: 'utf8' });
    console.log('  + ' + outPath);
}

console.log('');
console.log('Diese Dateien funktionieren von ueberall — Pfad ist eingebaut.');
console.log('Tipp: Rechtsklick auf eine .bat -> "Senden an" -> "Desktop (Verknuepfung erstellen)"');
console.log('oder einfach die Dateien auf den Desktop ziehen / kopieren.');
