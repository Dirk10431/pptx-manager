// =============================================================
// thumbnailer.js - Einzelfolien-Thumbnails via PowerPoint-COM
// =============================================================
// Oeffnet PPTX per PowerPoint-Automation und exportiert eine
// bestimmte Folie als PNG. Nur Windows, nur mit installiertem
// Microsoft Office.
// =============================================================

const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');

const THUMB_WIDTH = 480;   // Anzeige in der Liste, hoch genug fuer Erkennbarkeit
const THUMB_HEIGHT = 270;  // 16:9

/**
 * Laueft ein PowerShell-Skript mit gesetzter UTF-8-Ausgabe aus.
 * Das Skript wird Base64-encoded uebergeben, damit Umlaute & Anfuehrungszeichen
 * keine Probleme machen.
 */
function runPowerShell(psScript, timeoutMs = 60_000) {
    return new Promise((resolve, reject) => {
        const encoded = Buffer.from(psScript, 'utf16le').toString('base64');
        const ps = spawn('powershell.exe', [
            '-NoProfile', '-NonInteractive', '-STA', '-EncodedCommand', encoded,
        ]);

        let stdout = '';
        let stderr = '';
        const timer = setTimeout(() => {
            ps.kill();
            reject(new Error(`PowerShell-Timeout nach ${timeoutMs}ms`));
        }, timeoutMs);

        ps.stdout.on('data', (d) => { stdout += d.toString('utf8'); });
        ps.stderr.on('data', (d) => { stderr += d.toString('utf8'); });
        ps.on('error', (err) => { clearTimeout(timer); reject(err); });
        ps.on('close', (code) => {
            clearTimeout(timer);
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
    if (process.platform !== 'win32') {
        throw new Error('Thumbnail-Generierung aktuell nur unter Windows verfuegbar.');
    }
    if (!fs.existsSync(filePath)) {
        throw new Error('PPTX-Datei nicht gefunden: ' + filePath);
    }

    // Zielordner anlegen
    fs.mkdirSync(path.dirname(outPath), { recursive: true });

    // PowerShell-Skript: PowerPoint unsichtbar oeffnen, Slide exportieren.
    // Achtung: PowerPoint verlangt, dass der PPT-Prozess in einer STA laeuft.
    // Wir haben -STA in den Args, das passt.
    const ps = `
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$ppt = $null
$pres = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    # Visible = msoFalse (0) geht meist, manche Versionen verlangen msoTrue (1) + WithWindow:$false beim Open.
    try { $ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse } catch { $ppt.Visible = 1 }

    # Open(FileName, ReadOnly, Untitled, WithWindow)
    $pres = $ppt.Presentations.Open('${filePath.replace(/'/g, "''")}', $true, $true, $false)

    if ($pres.Slides.Count -lt ${slideIndex}) {
        throw "Folie ${slideIndex} existiert nicht (nur $($pres.Slides.Count) Folien)"
    }

    $slide = $pres.Slides.Item(${slideIndex})
    $slide.Export('${outPath.replace(/'/g, "''")}', 'PNG', ${THUMB_WIDTH}, ${THUMB_HEIGHT})

    Write-Output "OK"
} finally {
    if ($pres) { try { $pres.Close() } catch {} }
    if ($ppt)  { try { $ppt.Quit() } catch {} }
    # COM-Objekte freigeben, damit PowerPoint-Prozess sauber beendet
    if ($pres) { try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($pres) | Out-Null } catch {} }
    if ($ppt)  { try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt)  | Out-Null } catch {} }
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
    return new Promise((resolve, reject) => {
        if (process.platform !== 'win32') {
            return reject(new Error('Thumbnail-Generierung nur unter Windows verfuegbar.'));
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

$jobsJson = [Console]::In.ReadToEnd()
$jobs = $jobsJson | ConvertFrom-Json

$ppt = $null
$pres = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    try { $ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse } catch { $ppt.Visible = 1 }

    $pres = $ppt.Presentations.Open('${filePath.replace(/'/g, "''")}', $true, $true, $false)
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
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
`;

        const encoded = Buffer.from(psScript, 'utf16le').toString('base64');
        const ps = spawn('powershell.exe', [
            '-NoProfile', '-NonInteractive', '-STA', '-EncodedCommand', encoded,
        ]);

        const exported = [];
        const failed = [];
        let done = 0;
        let fatal = null;
        let buffer = '';

        const timer = setTimeout(() => {
            ps.kill();
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
            reject(err);
        });
        ps.on('close', (code) => {
            clearTimeout(timer);
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
    THUMB_WIDTH,
    THUMB_HEIGHT,
};
