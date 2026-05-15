// =============================================================
// thumbnailer.js - Einzelfolien-Thumbnails via PowerPoint
// =============================================================
// Oeffnet PPTX per PowerPoint-Automation und exportiert Folien
// als PNG. Zwei Plattform-Pfade:
//   - Windows: PowerPoint-COM via PowerShell, Folie-fuer-Folie-Export
//             auf Zielgroesse (THUMB_WIDTH x THUMB_HEIGHT).
//   - macOS:  PowerPoint via AppleScript (osascript). Mac-PowerPoint
//             exportiert immer das ganze Deck als PNG-Ordner; wir
//             pflueken die gewuenschten Folien raus und skalieren
//             per `sips` (macOS-Bordmittel) auf Zielgroesse.
// Voraussetzung in beiden Faellen: Microsoft PowerPoint installiert.
// =============================================================

const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');

const THUMB_WIDTH = 480;   // Anzeige in der Liste, hoch genug fuer Erkennbarkeit
const THUMB_HEIGHT = 270;  // 16:9

const IS_WIN = process.platform === 'win32';
const IS_MAC = process.platform === 'darwin';

// PowerShell-Helfer: PowerPoint-COM hat ein 255-Zeichen-Limit fuer Pfade
// (alte Win32-API). Wir vermeiden die GetShortPathName-API (kostet Add-Type
// und Norton-Heuristik feuert auf csc.exe) und gehen pragmatisch vor:
//   - Kurzer Pfad? Direkt nehmen.
//   - Langer Pfad? In %TEMP% mit GUID-Dateinamen kopieren, am Ende loeschen.
// Reine Standard-Kommandos, keine Laufzeit-Kompilierung.
const SHORTPATH_HELPER_PS = `
function Get-OpenablePath {
    param([Parameter(Mandatory)][string]$LongPath)
    if ($LongPath.Length -le 250) {
        return @{ Path = $LongPath; Temp = $null }
    }
    $ext = [System.IO.Path]::GetExtension($LongPath)
    if (-not $ext) { $ext = '.pptx' }
    $tempName = [System.Guid]::NewGuid().ToString('N') + $ext
    $tempPath = Join-Path $env:TEMP $tempName
    Copy-Item -LiteralPath $LongPath -Destination $tempPath -Force
    return @{ Path = $tempPath; Temp = $tempPath }
}
`;

// Eigener Skript-Cache-Ordner (nicht system-temp), damit Norton-Ausschluss
// auf den Projekt-Ordner greift.
const PS_SCRIPT_DIR = path.join(__dirname, '..', 'data', 'ps-scripts');

/**
 * Schreibt ein PowerShell-Skript in eine .ps1-Datei und gibt den Pfad zurueck.
 * UTF-8-BOM ist nicht noetig (Skript ist ASCII-only-PS); falls Umlaute drin
 * sind, sorgen wir per [Console]::OutputEncoding fuer korrekte Ausgabe.
 */
function writeTempPs1(scriptContents) {
    fs.mkdirSync(PS_SCRIPT_DIR, { recursive: true });
    const id = require('crypto').randomBytes(8).toString('hex');
    const file = path.join(PS_SCRIPT_DIR, `thumb_${id}.ps1`);
    // UTF-8-BOM, damit PowerShell die Datei ohne Encoding-Mismatch lesen kann.
    fs.writeFileSync(file, '\uFEFF' + scriptContents, 'utf8');
    return file;
}

function safeUnlink(file) {
    try { fs.unlinkSync(file); } catch (e) {}
}

/**
 * Laueft ein PowerShell-Skript via -File <pfad>.ps1.
 * Das vermeidet -EncodedCommand und damit Norton's IDP.HELU.PSE71-Erkennung.
 */
function runPowerShell(psScript, timeoutMs = 60_000) {
    return new Promise((resolve, reject) => {
        const scriptFile = writeTempPs1(psScript);
        const ps = spawn('powershell.exe', [
            '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
            '-STA', '-File', scriptFile,
        ]);

        let stdout = '';
        let stderr = '';
        const timer = setTimeout(() => {
            ps.kill();
            safeUnlink(scriptFile);
            reject(new Error(`PowerShell-Timeout nach ${timeoutMs}ms`));
        }, timeoutMs);

        ps.stdout.on('data', (d) => { stdout += d.toString('utf8'); });
        ps.stderr.on('data', (d) => { stderr += d.toString('utf8'); });
        ps.on('error', (err) => { clearTimeout(timer); safeUnlink(scriptFile); reject(err); });
        ps.on('close', (code) => {
            clearTimeout(timer);
            safeUnlink(scriptFile);
            if (code === 0) {
                resolve({ stdout, stderr });
            } else {
                reject(new Error(`PowerShell exit ${code}: ${stderr || stdout}`));
            }
        });
    });
}

/**
 * Eine einzelne Folie als PNG exportieren.
 * Oeffnet PPTX (ReadOnly), exportiert Slide N, schliesst alles.
 * Gibt das Ziel-PNG-Pfad zurueck.
 *
 * filePath    - absolute PPTX-Datei
 * slideIndex  - 1-basierter Folien-Index
 * outPath     - absolute Ziel-Datei (.png)
 */
async function generateThumbnail(filePath, slideIndex, outPath) {
    if (!IS_WIN && !IS_MAC) {
        throw new Error('Thumbnail-Generierung aktuell nur unter Windows und macOS verfuegbar.');
    }
    if (!fs.existsSync(filePath)) {
        throw new Error('PPTX-Datei nicht gefunden: ' + filePath);
    }

    // Zielordner anlegen
    fs.mkdirSync(path.dirname(outPath), { recursive: true });

    if (IS_MAC) {
        // Mac-Pfad: Batch fuer eine einzelne Folie nutzen (gleicher Code-Pfad).
        const { failed } = await generateThumbnailsBatch(filePath, [{ slideIndex, outPath }]);
        if (failed.length > 0) {
            throw new Error(failed[0].error || 'Mac-Thumbnail-Export fehlgeschlagen');
        }
        return outPath;
    }

    // PowerShell-Skript: PowerPoint unsichtbar oeffnen, Slide exportieren.
    // Achtung: PowerPoint verlangt, dass der PPT-Prozess in einer STA laeuft.
    // Wir haben -STA in den Args, das passt.
    const ps = `
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

${SHORTPATH_HELPER_PS}

$ppt = $null
$pres = $null
$tempCopy = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    try { $ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse } catch { $ppt.Visible = 1 }

    $openInfo = Get-OpenablePath '${filePath.replace(/'/g, "''")}'
    $tempCopy = $openInfo.Temp

    # Open(FileName, ReadOnly, Untitled, WithWindow)
    $pres = $ppt.Presentations.Open($openInfo.Path, $true, $true, $false)

    if ($pres.Slides.Count -lt ${slideIndex}) {
        throw "Folie ${slideIndex} existiert nicht (nur $($pres.Slides.Count) Folien)"
    }

    $slide = $pres.Slides.Item(${slideIndex})
    $slide.Export('${outPath.replace(/'/g, "''")}', 'PNG', ${THUMB_WIDTH}, ${THUMB_HEIGHT})

    Write-Output "OK"
} finally {
    if ($pres) { try { $pres.Close() } catch {} }
    if ($ppt)  { try { $ppt.Quit() } catch {} }
    if ($pres) { try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($pres) | Out-Null } catch {} }
    if ($ppt)  { try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt)  | Out-Null } catch {} }
    if ($tempCopy -and (Test-Path -LiteralPath $tempCopy)) { Remove-Item -LiteralPath $tempCopy -Force -ErrorAction SilentlyContinue }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
`;

    await runPowerShell(ps, 60_000);

    if (!fs.existsSync(outPath)) {
        throw new Error('Export lief durch, aber PNG nicht vorhanden: ' + outPath);
    }

    return outPath;
}

// =============================================================
// macOS-Implementierung
// =============================================================
// PowerPoint Mac unterstuetzt AppleScript. "save thePres in <ordner>
// as save as PNG" exportiert ALLE Folien in einen Ordner; die
// Dateinamen variieren je nach Sprache der PowerPoint-Installation
// ("Slide1.png" / "Folie1.png" / "Diapositive 1.png" ...). Wir
// sortieren die erzeugten PNGs deshalb nach Aenderungszeit-Reihenfolge
// PLUS Nummern-Extraktion aus dem Dateinamen — beides zusammen ist
// gegen Sprache und Sortier-Eigenheiten robust.
//
// Aus PowerPoints Vollformat-PNGs (z.B. 1280x720) machen wir mit `sips`
// (in macOS eingebaut) das Anzeige-Format THUMB_WIDTH x THUMB_HEIGHT.
// =============================================================

function runOsascript(scriptBody, timeoutMs = 90_000) {
    return new Promise((resolve, reject) => {
        const proc = spawn('osascript', ['-e', scriptBody]);
        let stdout = '';
        let stderr = '';
        const timer = setTimeout(() => {
            proc.kill();
            reject(new Error(`osascript-Timeout nach ${timeoutMs}ms`));
        }, timeoutMs);
        proc.stdout.on('data', (d) => { stdout += d.toString('utf8'); });
        proc.stderr.on('data', (d) => { stderr += d.toString('utf8'); });
        proc.on('error', (err) => { clearTimeout(timer); reject(err); });
        proc.on('close', (code) => {
            clearTimeout(timer);
            if (code === 0) resolve({ stdout, stderr });
            else reject(new Error(`osascript exit ${code}: ${stderr || stdout}`));
        });
    });
}

// PNGs in einem Ordner in PowerPoint-Slide-Reihenfolge bringen.
// Strategie:
//   1) Aus dem Dateinamen erste vorkommende Zahlen-Sequenz lesen
//      ("Folie 12.png" -> 12, "Diapositive 03.PNG" -> 3).
//   2) Als sekundaerer Sortierschluessel die mtime-Reihenfolge (PowerPoint
//      schreibt sie in Reihenfolge). Das fängt den Fall ab, dass jemand
//      Folien ohne Zahl benannt hat oder die Sprache exotisch ist.
function sortMacExportedPngs(files, stats) {
    return files.slice().sort((a, b) => {
        const numA = (a.match(/\d+/) || [Infinity])[0];
        const numB = (b.match(/\d+/) || [Infinity])[0];
        const da = parseInt(numA, 10);
        const db = parseInt(numB, 10);
        if (da !== db) return da - db;
        return (stats.get(a) || 0) - (stats.get(b) || 0);
    });
}

async function exportPptxAsPngFolderMac(filePath, exportDir, timeoutMs) {
    // POSIX-Pfade duerfen in AppleScript-String stehen, Quotes maskieren wir.
    const safeFile = filePath.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
    const safeDir  = exportDir.replace(/\\/g, '\\\\').replace(/"/g, '\\"');

    // Wichtig: PowerPoint Mac haengt den App-Namen an, blockiert Dialoge wenn
    // visible. Daher "without showing window" / im Hintergrund halten. Falls
    // PowerPoint nicht laeuft, startet es kurz auf. Wir lassen offene
    // Praesentationen anderer Nutzer in Ruhe und schliessen NUR unsere.
    const applescript = `
set thePosixFile to "${safeFile}"
set theExportDir to "${safeDir}"
set theFileAlias to POSIX file thePosixFile
set theExportFolder to POSIX file theExportDir
tell application "Microsoft PowerPoint"
    activate
    set thePres to open theFileAlias with read only
    -- "save as PNG" als Format-Konstante: Powerpoint Mac kennt enum
    -- "save as PNG" (15) -> ganzes Deck als PNG-Ordner.
    save thePres in theExportFolder as save as PNG
    close thePres saving no
end tell
return "OK"
`;
    await runOsascript(applescript, timeoutMs);
}

function resizePngWithSips(srcPath, dstPath) {
    return new Promise((resolve, reject) => {
        // sips -z H W <src> --out <dst>: Resampling, Aspect-Ratio wird ignoriert
        // (genau das Verhalten der Windows-Variante mit fester Zielgroesse).
        const proc = spawn('sips', ['-z', String(THUMB_HEIGHT), String(THUMB_WIDTH), srcPath, '--out', dstPath]);
        let stderr = '';
        proc.stderr.on('data', (d) => { stderr += d.toString('utf8'); });
        proc.on('error', reject);
        proc.on('close', (code) => {
            if (code === 0) resolve();
            else reject(new Error(`sips exit ${code}: ${stderr}`));
        });
    });
}

async function generateThumbnailsBatchMac(filePath, jobs, onProgress = null, options = {}) {
    if (!fs.existsSync(filePath)) {
        throw new Error('PPTX-Datei nicht gefunden: ' + filePath);
    }
    if (!Array.isArray(jobs) || jobs.length === 0) {
        return { exported: [], failed: [] };
    }
    for (const job of jobs) {
        fs.mkdirSync(path.dirname(job.outPath), { recursive: true });
    }

    // Temp-Ordner unter data/ (statt /tmp), damit der gleiche
    // Antivirus-/Ordner-Ausschluss greift wie fuer den Cache.
    const tempBase = path.join(__dirname, '..', 'data', 'tmp-mac-export');
    fs.mkdirSync(tempBase, { recursive: true });
    const exportDir = fs.mkdtempSync(path.join(tempBase, 'pptx-'));
    const timeoutMs = options.timeoutMs || Math.max(120_000, jobs.length * 3_000);

    const exported = [];
    const failed = [];
    let done = 0;

    try {
        await exportPptxAsPngFolderMac(filePath, exportDir, timeoutMs);

        // Alle PNGs im Export-Ordner aufsammeln.
        const allFiles = fs.readdirSync(exportDir).filter(f => /\.png$/i.test(f));
        if (allFiles.length === 0) {
            // Pauschal-Fehler: kein PNG entstanden
            for (const job of jobs) {
                failed.push({ slideIndex: job.slideIndex, error: 'PowerPoint hat keine PNGs erzeugt' });
                done++;
                if (onProgress) onProgress(done, jobs.length, job.slideIndex, null);
            }
            return { exported, failed };
        }

        // mtime-Map fuer Tie-Breaker
        const mtimeMap = new Map();
        for (const f of allFiles) {
            const st = fs.statSync(path.join(exportDir, f));
            mtimeMap.set(f, st.mtimeMs);
        }
        const orderedFiles = sortMacExportedPngs(allFiles, mtimeMap);

        for (const job of jobs) {
            const idx = job.slideIndex;
            if (!Number.isInteger(idx) || idx < 1 || idx > orderedFiles.length) {
                failed.push({ slideIndex: idx, error: `out-of-range (total=${orderedFiles.length})` });
                done++;
                if (onProgress) onProgress(done, jobs.length, idx, null);
                continue;
            }
            const srcPng = path.join(exportDir, orderedFiles[idx - 1]);
            try {
                await resizePngWithSips(srcPng, job.outPath);
                exported.push({ slideIndex: idx, outPath: job.outPath });
                done++;
                if (onProgress) onProgress(done, jobs.length, idx, job.outPath);
            } catch (err) {
                failed.push({ slideIndex: idx, error: err.message });
                done++;
                if (onProgress) onProgress(done, jobs.length, idx, null);
            }
        }
        return { exported, failed };
    } finally {
        // Temp-Export-Ordner wegraeumen
        try { fs.rmSync(exportDir, { recursive: true, force: true }); } catch (e) { /* egal */ }
    }
}

/**
 * Cache-Pfad fuer ein Thumbnail bestimmen.
 * Basis: <cacheDir>/<first2chars>/<textHash>.png  (Shard-Verzeichnis, damit
 * nicht alle PNGs in einem Ordner landen — je nach FS wuerde das langsam).
 *
 * Erwartet einen hex-Hash (SHA-256, 64 Zeichen). Fuer Rueckwaerts-Kompat
 * akzeptiert diese Funktion auch eine Zahl (Slide-ID) -> flacher Pfad ohne Shard.
 */
function thumbnailCachePath(cacheDir, key) {
    const s = String(key);
    if (/^[0-9a-f]{16,}$/i.test(s)) {
        // Hex-Hash: Shard nach ersten 2 Zeichen
        return path.join(cacheDir, s.slice(0, 2), `${s}.png`);
    }
    // Alt-Pfad (Slide-ID als Zahl) — flach
    return path.join(cacheDir, `${s}.png`);
}

/**
 * BATCH: Mehrere Folien derselben PPTX in einem Rutsch exportieren.
 * PowerPoint + Datei werden EINMAL geoeffnet, dann alle Jobs abgearbeitet.
 * Das ist der grosse Zeitgewinn gegenueber vielen Einzel-Aufrufen.
 *
 * filePath   - absolute PPTX-Datei
 * jobs       - Array von { slideIndex, outPath }
 * onProgress - optional: (done, total, slideIndex, outPath) => void
 *
 * Gibt zurueck: { exported: [{slideIndex, outPath}], failed: [{slideIndex, error}] }
 *
 * Protokoll mit dem PowerShell-Prozess:
 *   - PS liest Jobs als JSON-Liste via stdin
 *   - PS schreibt pro Job eine Zeile nach stdout:
 *       "OK <slideIndex>"   - Erfolg
 *       "ERR <slideIndex> <message>" - Fehler
 *     am Ende: "DONE"
 */
function generateThumbnailsBatch(filePath, jobs, onProgress = null, options = {}) {
    if (IS_MAC) {
        return generateThumbnailsBatchMac(filePath, jobs, onProgress, options);
    }
    return new Promise((resolve, reject) => {
        if (!IS_WIN) {
            return reject(new Error('Thumbnail-Generierung nur unter Windows und macOS verfuegbar.'));
        }
        if (!fs.existsSync(filePath)) {
            return reject(new Error('PPTX-Datei nicht gefunden: ' + filePath));
        }
        if (!Array.isArray(jobs) || jobs.length === 0) {
            return resolve({ exported: [], failed: [] });
        }

        // Zielverzeichnisse anlegen
        for (const job of jobs) {
            fs.mkdirSync(path.dirname(job.outPath), { recursive: true });
        }

        const timeoutMs = options.timeoutMs || Math.max(120_000, jobs.length * 3_000);

        // PowerShell-Skript: PPTX einmal oeffnen, alle Jobs abarbeiten.
        // Jobs werden als JSON per stdin uebergeben ($stdin = [Console]::In.ReadToEnd()).
        const psScript = `
$ErrorActionPreference = 'Continue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8

${SHORTPATH_HELPER_PS}

$jobsJson = [Console]::In.ReadToEnd()
$jobs = $jobsJson | ConvertFrom-Json

$ppt = $null
$pres = $null
$tempCopy = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    try { $ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse } catch { $ppt.Visible = 1 }

    $openInfo = Get-OpenablePath '${filePath.replace(/'/g, "''")}'
    $tempCopy = $openInfo.Temp

    $pres = $ppt.Presentations.Open($openInfo.Path, $true, $true, $false)
    $slideCount = $pres.Slides.Count
    Write-Output ("INFO slides=" + $slideCount)

    foreach ($job in $jobs) {
        $idx = [int]$job.slideIndex
        $out = [string]$job.outPath
        try {
            if ($idx -lt 1 -or $idx -gt $slideCount) {
                Write-Output ("ERR " + $idx + " out-of-range (total=" + $slideCount + ")")
                continue
            }
            $slide = $pres.Slides.Item($idx)
            $slide.Export($out, 'PNG', ${THUMB_WIDTH}, ${THUMB_HEIGHT})
            Write-Output ("OK " + $idx)
        } catch {
            $msg = $_.Exception.Message -replace "[\`r\`n]", " "
            Write-Output ("ERR " + $idx + " " + $msg)
        }
    }

    Write-Output "DONE"
} catch {
    $msg = $_.Exception.Message -replace "[\`r\`n]", " "
    Write-Output ("FATAL " + $msg)
} finally {
    if ($pres) { try { $pres.Close() } catch {} }
    if ($ppt)  { try { $ppt.Quit() } catch {} }
    if ($pres) { try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($pres) | Out-Null } catch {} }
    if ($ppt)  { try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt)  | Out-Null } catch {} }
    if ($tempCopy -and (Test-Path -LiteralPath $tempCopy)) { Remove-Item -LiteralPath $tempCopy -Force -ErrorAction SilentlyContinue }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
`;

        // Skript in .ps1 schreiben statt -EncodedCommand: Norton sieht keinen
        // Base64-Befehl mehr und beruhigt sich. Datei liegt im Projekt-Ordner,
        // den der User in Norton ausgeschlossen hat.
        const scriptFile = writeTempPs1(psScript);
        const ps = spawn('powershell.exe', [
            '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
            '-STA', '-File', scriptFile,
        ]);

        const exported = [];
        const failed = [];
        let done = 0;
        let fatal = null;
        let buffer = '';

        const timer = setTimeout(() => {
            ps.kill();
            safeUnlink(scriptFile);
            reject(new Error(`Batch-Timeout nach ${timeoutMs}ms (${done}/${jobs.length} fertig)`));
        }, timeoutMs);

        // JSON-Input pushen und Eingabe schliessen
        ps.stdin.write(JSON.stringify(jobs.map(j => ({
            slideIndex: j.slideIndex,
            outPath: j.outPath,
        }))));
        ps.stdin.end();

        ps.stdout.on('data', (chunk) => {
            buffer += chunk.toString('utf8');
            let nl;
            while ((nl = buffer.indexOf('\n')) >= 0) {
                const line = buffer.slice(0, nl).replace(/\r$/, '').trim();
                buffer = buffer.slice(nl + 1);
                if (!line) continue;

                if (line.startsWith('OK ')) {
                    const idx = parseInt(line.slice(3).trim(), 10);
                    const job = jobs.find(j => j.slideIndex === idx);
                    if (job) {
                        exported.push({ slideIndex: idx, outPath: job.outPath });
                        done++;
                        if (onProgress) onProgress(done, jobs.length, idx, job.outPath);
                    }
                } else if (line.startsWith('ERR ')) {
                    const rest = line.slice(4).trim();
                    const spaceIdx = rest.indexOf(' ');
                    const idx = parseInt(spaceIdx > 0 ? rest.slice(0, spaceIdx) : rest, 10);
                    const msg = spaceIdx > 0 ? rest.slice(spaceIdx + 1) : 'unknown';
                    failed.push({ slideIndex: idx, error: msg });
                    done++;
                    if (onProgress) onProgress(done, jobs.length, idx, null);
                } else if (line.startsWith('FATAL ')) {
                    fatal = line.slice(6).trim();
                } else if (line === 'DONE') {
                    // wird im close-Handler behandelt
                }
                // INFO-Zeilen ignorieren
            }
        });

        let stderrBuf = '';
        ps.stderr.on('data', (d) => { stderrBuf += d.toString('utf8'); });

        ps.on('error', (err) => {
            clearTimeout(timer);
            safeUnlink(scriptFile);
            reject(err);
        });
        ps.on('close', (code) => {
            clearTimeout(timer);
            safeUnlink(scriptFile);
            if (fatal) {
                return reject(new Error('PowerPoint-Fehler: ' + fatal));
            }
            if (code !== 0 && exported.length === 0 && failed.length === 0) {
                return reject(new Error(`PowerShell exit ${code}: ${stderrBuf || 'unbekannt'}`));
            }
            resolve({ exported, failed });
        });
    });
}

module.exports = {
    generateThumbnail,
    generateThumbnailsBatch,
    thumbnailCachePath,
    writeTempPs1,
    safeUnlink,
    THUMB_WIDTH,
    THUMB_HEIGHT,
};
