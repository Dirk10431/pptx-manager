// =============================================================
// pptx-manager - Lokaler Server
// =============================================================
// Verwaltet PPTX-Sammlungen, findet Duplikate, ermoeglicht
// Volltextsuche. Laeuft nur lokal auf 127.0.0.1.
// =============================================================

const express = require('express');
const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');
const {
    initDb,
    checkFingerprintCompatibility,
    setFingerprintVersion,
    clearScanData,
    listScanRoots,
    getScanRoot,
    updateScanRootLabel,
    deleteScanRoot,
    listPresentationsInRoot,
    deletePresentationsByIds,
    garbageCollectThumbnails,
} = require('./lib/db');
const { runScan, scanState, findPptxFiles } = require('./lib/scanner');
const { FINGERPRINT_VERSION } = require('./lib/fingerprint');
const { generateThumbnail, thumbnailCachePath, writeTempPs1, safeUnlink } = require('./lib/thumbnailer');
const { loadConfig } = require('./lib/config');

const config = loadConfig();
console.log('[CONFIG] geladen:', JSON.stringify(config));

const PORT = process.env.PORT || 3002;
const HOST = '127.0.0.1'; // NUR lokal, nicht im Netzwerk!

const app = express();
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// --- Datenbank initialisieren ---
const dbPath = path.join(__dirname, 'data', 'pptx-manager.db');
fs.mkdirSync(path.dirname(dbPath), { recursive: true });
const db = initDb(dbPath);
console.log(`[DB] Datenbank: ${dbPath}`);

// Thumbnail-Cache-Verzeichnis
const thumbCacheDir = path.join(__dirname, 'data', 'thumbnails');
fs.mkdirSync(thumbCacheDir, { recursive: true });
console.log(`[THUMB] Cache: ${thumbCacheDir}`);

// Fingerprint-Versions-Check beim Start
const compat = checkFingerprintCompatibility(db, FINGERPRINT_VERSION);
if (!compat.compatible) {
    console.warn(`[DB] WARNUNG: DB enthaelt ${compat.slideCount} Folien, gescannt mit Fingerprint-Version ${compat.storedVersion}.`);
    console.warn(`[DB] Aktuelle Code-Version ist ${compat.codeVersion}. Duplikatsuche ist deshalb unzuverlaessig.`);
    console.warn(`[DB] Bitte in der Weboberflaeche "Neu scannen (Reset)" klicken oder data/pptx-manager.db loeschen.`);
} else if (!compat.hasData) {
    // Leere DB -> Version direkt setzen
    setFingerprintVersion(db, FINGERPRINT_VERSION);
}

// --- API: Status/Statistik ---
// --- API: Datei oeffnen oder im Datei-Browser zeigen ---
// Sicherheit: filePath MUSS in der DB stehen, sonst koennten beliebige
// Dateien per HTTP-Request geoeffnet werden. Gilt nur lokal (127.0.0.1).
// Unterstuetzt Windows (cmd/explorer) und macOS (open).
app.post('/api/open', (req, res) => {
    if (process.platform !== 'win32' && process.platform !== 'darwin') {
        return res.status(501).json({ error: 'Datei oeffnen aktuell nur unter Windows und macOS verfuegbar.' });
    }

    const { filePath, mode } = req.body || {};
    if (!filePath || !mode) {
        return res.status(400).json({ error: 'filePath und mode (file|folder) erforderlich.' });
    }
    if (mode !== 'file' && mode !== 'folder') {
        return res.status(400).json({ error: 'Ungueltiger mode (file|folder).' });
    }

    // Validierung: Pfad muss in presentations stehen
    const known = db.prepare('SELECT 1 FROM presentations WHERE file_path = ?').get(filePath);
    if (!known) {
        return res.status(404).json({ error: 'Pfad nicht in der Datenbank.' });
    }
    if (!fs.existsSync(filePath)) {
        return res.status(404).json({ error: 'Datei nicht mehr auf der Platte gefunden.' });
    }

    try {
        let child;
        if (process.platform === 'win32') {
            if (mode === 'file') {
                // In Standard-Anwendung oeffnen (PowerPoint o.ae.)
                // cmd /c start "" "<datei>" -- start braucht einen leeren Titel als 1. Arg
                child = spawn('cmd', ['/c', 'start', '""', filePath], {
                    detached: true,
                    stdio: 'ignore',
                    windowsHide: true,
                });
            } else {
                // Uebergeordneten Ordner oeffnen.
                // Zwei Methoden:
                //   A) cmd /c start "" <folder> — bringt Fenster in den Vordergrund,
                //      scheitert aber bei Pfaden mit Klammern UND Kommas (cmd-Quoting)
                //   B) spawn('explorer.exe', [folder]) — immer zuverlaessig, aber
                //      Fenster oeffnet im Hintergrund (Taskleiste blinkt)
                // Wir nehmen A, ausser der Pfad ist "auffaellig" (Klammern + Komma).
                const folder = path.dirname(filePath);
                const tricky = /[()]/.test(folder) && /,/.test(folder);
                if (tricky) {
                    child = spawn('explorer.exe', [folder], {
                        detached: true,
                        stdio: 'ignore',
                    });
                } else {
                    child = spawn('cmd', ['/c', 'start', '""', folder], {
                        detached: true,
                        stdio: 'ignore',
                        windowsHide: true,
                    });
                }
            }
        } else {
            // macOS: open <datei> startet Standard-App (PowerPoint fuer .pptx).
            //        open -R <datei> oeffnet Finder und markiert die Datei darin.
            // -g vermeidet, dass Finder den Fokus klaut, falls Browser aktiv ist;
            //    bei "file" lassen wir Fokus auf PowerPoint wandern (-g weglassen).
            if (mode === 'file') {
                child = spawn('open', [filePath], {
                    detached: true,
                    stdio: 'ignore',
                });
            } else {
                child = spawn('open', ['-R', filePath], {
                    detached: true,
                    stdio: 'ignore',
                });
            }
        }
        child.unref();
        res.json({ ok: true, mode });
    } catch (err) {
        console.error('[OPEN] Fehler:', err.message);
        res.status(500).json({ error: 'Oeffnen fehlgeschlagen: ' + err.message });
    }
});

// --- API: Hauptpfade aus scan_roots-Tabelle ---
// Frueher per Heuristik aus file_path geraten — jetzt persistierte Roots,
// vom Nutzer beim Scan explizit gewaehlt und per Label benennbar.
app.get('/api/scan-roots', (req, res) => {
    const rows = listScanRoots(db);
    const roots = rows.map(r => ({
        id: r.id,
        fullPath: r.rootPath,
        label: r.label,
        count: r.presentationCount,
        slideCount: r.slideCount,
        lastScannedAt: r.lastScannedAt,
        createdAt: r.createdAt,
    })).sort((a, b) => a.label.localeCompare(b.label, 'de', { sensitivity: 'base' }));
    res.json({ roots });
});

// Label eines Roots aendern
app.patch('/api/scan-roots/:id', (req, res) => {
    const id = parseInt(req.params.id, 10);
    const { label } = req.body || {};
    if (!Number.isInteger(id) || id <= 0) {
        return res.status(400).json({ error: 'Ungueltige id.' });
    }
    if (!label || typeof label !== 'string' || !label.trim()) {
        return res.status(400).json({ error: 'label erforderlich (nicht leer).' });
    }
    const changed = updateScanRootLabel(db, id, label.trim());
    if (!changed) return res.status(404).json({ error: 'Root nicht gefunden.' });
    res.json({ ok: true });
});

// Kompletten Root inkl. aller Praesentationen / Folien loeschen.
// Thumbnails ohne zugehoerige Folie werden danach mit-eingesammelt.
app.delete('/api/scan-roots/:id', (req, res) => {
    if (scanState.running) {
        return res.status(409).json({ error: 'Scan laeuft gerade, bitte warten.' });
    }
    const id = parseInt(req.params.id, 10);
    if (!Number.isInteger(id) || id <= 0) {
        return res.status(400).json({ error: 'Ungueltige id.' });
    }
    const root = getScanRoot(db, id);
    if (!root) return res.status(404).json({ error: 'Root nicht gefunden.' });

    const before = listPresentationsInRoot(db, id).length;
    deleteScanRoot(db, id);
    const orphans = garbageCollectThumbnails(db);
    let thumbsDeleted = 0;
    for (const o of orphans) {
        try { fs.unlinkSync(o.pngPath); thumbsDeleted++; } catch (e) { /* schon weg */ }
    }
    console.log(`[ROOT-DELETE] "${root.label}" (${root.rootPath}): ${before} Praesentationen, ${thumbsDeleted}/${orphans.length} Thumbnails`);
    res.json({ ok: true, deletedPresentations: before, deletedThumbnails: thumbsDeleted });
});

// Einen einzelnen Root rescannen (gleicher Pfad, gleiches Label aus der DB)
app.post('/api/scan-roots/:id/rescan', (req, res) => {
    if (scanState.running) {
        return res.status(409).json({ error: 'Ein Scan laeuft bereits.' });
    }
    const id = parseInt(req.params.id, 10);
    if (!Number.isInteger(id) || id <= 0) {
        return res.status(400).json({ error: 'Ungueltige id.' });
    }
    const root = getScanRoot(db, id);
    if (!root) return res.status(404).json({ error: 'Root nicht gefunden.' });
    if (!fs.existsSync(root.rootPath)) {
        return res.status(400).json({ error: 'Pfad existiert nicht mehr auf der Platte.' });
    }
    runScan(db, root.rootPath, root.label).catch(err => {
        console.error('[RESCAN] Unerwarteter Fehler:', err);
    });
    res.json({ ok: true, message: `Rescan von "${root.label}" gestartet.` });
});

// Alle Roots nacheinander rescannen
let rescanAllPending = null;
app.post('/api/scan-roots/rescan-all', (req, res) => {
    if (scanState.running || rescanAllPending) {
        return res.status(409).json({ error: 'Es laeuft bereits ein Scan / Massen-Rescan.' });
    }
    const roots = listScanRoots(db);
    if (roots.length === 0) {
        return res.json({ ok: true, message: 'Keine Roots vorhanden.' });
    }
    rescanAllPending = (async () => {
        for (const r of roots) {
            if (!fs.existsSync(r.rootPath)) {
                console.warn(`[RESCAN-ALL] uebersprungen (Pfad weg): ${r.rootPath}`);
                continue;
            }
            try {
                await runScan(db, r.rootPath, r.label);
            } catch (err) {
                console.error(`[RESCAN-ALL] Fehler bei ${r.rootPath}:`, err.message);
            }
        }
    })().finally(() => { rescanAllPending = null; });
    res.json({ ok: true, message: `Rescan von ${roots.length} Roots gestartet.`, count: roots.length });
});

// Stale-Files in einem Root: in der DB unter dem Root, aber nicht mehr
// auf der Platte vorhanden. Live-Check via fs.existsSync.
app.get('/api/scan-roots/:id/stale', (req, res) => {
    const id = parseInt(req.params.id, 10);
    if (!Number.isInteger(id) || id <= 0) {
        return res.status(400).json({ error: 'Ungueltige id.' });
    }
    const root = getScanRoot(db, id);
    if (!root) return res.status(404).json({ error: 'Root nicht gefunden.' });
    const presentations = listPresentationsInRoot(db, id);
    const stale = presentations.filter(p => !fs.existsSync(p.filePath));
    res.json({ stale });
});

// Stale-Files in einem Root loeschen (mit explizit uebergebenen ids).
// IDs werden gegen den Root validiert — Sicherung gegen Cross-Root-Delete.
app.post('/api/scan-roots/:id/cleanup-stale', (req, res) => {
    if (scanState.running) {
        return res.status(409).json({ error: 'Scan laeuft, bitte warten.' });
    }
    const id = parseInt(req.params.id, 10);
    if (!Number.isInteger(id) || id <= 0) {
        return res.status(400).json({ error: 'Ungueltige id.' });
    }
    const root = getScanRoot(db, id);
    if (!root) return res.status(404).json({ error: 'Root nicht gefunden.' });

    const { ids } = req.body || {};
    if (!Array.isArray(ids) || ids.length === 0) {
        return res.status(400).json({ error: 'ids[] erforderlich.' });
    }
    const numericIds = ids.map(n => parseInt(n, 10)).filter(n => Number.isInteger(n) && n > 0);
    if (numericIds.length === 0) return res.json({ ok: true, deletedPresentations: 0, deletedThumbnails: 0 });

    // Sicherheits-Check: nur ids loeschen, die wirklich unter diesem Root liegen.
    const placeholders = numericIds.map(() => '?').join(',');
    const owned = db.prepare(
        `SELECT id FROM presentations WHERE scan_root_id = ? AND id IN (${placeholders})`
    ).all(id, ...numericIds).map(r => r.id);

    const deleted = deletePresentationsByIds(db, owned);
    const orphans = garbageCollectThumbnails(db);
    let thumbsDeleted = 0;
    for (const o of orphans) {
        try { fs.unlinkSync(o.pngPath); thumbsDeleted++; } catch (e) { /* schon weg */ }
    }
    console.log(`[CLEANUP-STALE] root="${root.label}": ${deleted} Praesentationen, ${thumbsDeleted}/${orphans.length} Thumbnails geloescht`);
    res.json({ ok: true, deletedPresentations: deleted, deletedThumbnails: thumbsDeleted });
});

// --- API: Projekt-Info (Pfad, Versionsnummer) fuer UI-Hinweise ---
// Wird genutzt, um in der Scan-Seite Bash-Befehle wie
//   cd "<projectRoot>"
//   npm run thumbs
// korrekt anzuzeigen — ohne dass das Frontend den Pfad raten muss.
app.get('/api/project-info', (req, res) => {
    res.json({
        projectRoot: __dirname,
        platform: process.platform,
    });
});

// --- API: Frontend-relevante Config-Werte ---
// Wird beim Laden der Seite einmal gefetched, damit Lightbox-Groesse,
// Page-Size etc. zentral aus config.json kommen.
app.get('/api/config', (req, res) => {
    res.json({
        search: {
            pageSize: config.search.pageSize,
            occurrencesShown: config.search.occurrencesShown,
            minQueryLength: config.search.minQueryLength,
            yearFilterOptions: buildYearFilterOptions(config.search.yearFilterYearsBack),
        },
        lightbox: {
            heightVh: config.lightbox.heightVh,
            maxWidthVw: config.lightbox.maxWidthVw,
        },
    });
});

app.get('/api/stats', (req, res) => {
    const presentationCount = db.prepare('SELECT COUNT(*) AS c FROM presentations').get().c;
    const slideCount = db.prepare('SELECT COUNT(*) AS c FROM slides').get().c;
    const uniqueSlides = db.prepare('SELECT COUNT(DISTINCT text_hash) AS c FROM slides').get().c;
    const thumbnailsDone = db.prepare('SELECT COUNT(*) AS c FROM thumbnails').get().c;

    const duplicateSlides = Math.max(0, slideCount - uniqueSlides);
    const thumbnailsMissing = Math.max(0, uniqueSlides - thumbnailsDone);
    const duplicateFactor = uniqueSlides > 0 ? slideCount / uniqueSlides : 0;
    const thumbnailCoverage = uniqueSlides > 0 ? thumbnailsDone / uniqueSlides : 0;

    const c = checkFingerprintCompatibility(db, FINGERPRINT_VERSION);
    res.json({
        presentations: presentationCount,
        slides: slideCount,
        uniqueSlides,
        duplicateSlides,
        thumbnailsDone,
        thumbnailsMissing,
        thumbnailCoverage,         // 0..1
        duplicateFactor,           // 1.0 = keine Duplikate, 8.1 = im Schnitt 8x kopiert
        dbPath: dbPath,
        fingerprintVersion: FINGERPRINT_VERSION,
        storedFingerprintVersion: c.storedVersion,
        compatible: c.compatible,
    });
});

// --- API: Alle Scan-Daten loeschen (bei Inkompatibilitaet oder auf Wunsch) ---
app.post('/api/reset', (req, res) => {
    if (scanState.running) {
        return res.status(409).json({ error: 'Scan laeuft gerade, bitte warten.' });
    }
    clearScanData(db);
    setFingerprintVersion(db, FINGERPRINT_VERSION);
    // Thumbnail-Cache-Dateien auch wegraeumen
    try {
        fs.rmSync(thumbCacheDir, { recursive: true, force: true });
        fs.mkdirSync(thumbCacheDir, { recursive: true });
    } catch (err) {
        console.warn('[RESET] Thumbnail-Cache konnte nicht geloescht werden:', err.message);
    }
    res.json({ ok: true });
});

// --- API: Thumbnail einer Folie (lazy erzeugt, dann geaecht) ---
// In-Memory-Dedupe: Wenn zwei Requests parallel dasselbe Thumbnail anfragen,
// wartet der zweite auf den Export des ersten statt PowerPoint zweimal zu starten.
const thumbPending = new Map();

app.get('/api/thumb/:slideId', async (req, res) => {
    const slideId = parseInt(req.params.slideId, 10);
    if (!Number.isInteger(slideId) || slideId <= 0) {
        return res.status(400).json({ error: 'Ungueltige Folien-ID.' });
    }

    // Folie nachschlagen: text_hash (Dedupe-Key) + Quell-Datei + Folien-Index
    const row = db.prepare(`
        SELECT s.id AS slide_id, s.slide_index, s.text_hash, p.file_path
        FROM slides s
        JOIN presentations p ON p.id = s.presentation_id
        WHERE s.id = ?
    `).get(slideId);

    if (!row) {
        return res.status(404).json({ error: 'Folie nicht gefunden.' });
    }

    const textHash = row.text_hash;
    const outPath = thumbnailCachePath(thumbCacheDir, textHash);

    // 1) Aus Cache ausliefern, falls vorhanden (alle Duplikate teilen sich dieses PNG)
    if (fs.existsSync(outPath)) {
        res.setHeader('Cache-Control', 'public, max-age=31536000, immutable');
        return res.sendFile(outPath);
    }

    // 2) Standardfall: kein Cache vorhanden -> 404. Nutzer erzeugt fehlende
    //    Thumbnails per `npm run thumbs` (Batch). Lazy-Render war frueher
    //    Default, hat aber pro Anfrage ~15s gekostet -> abgeschaltet.
    //    Opt-in fuer Einzelfall: ?generate=1 erzwingt sofortige Generierung.
    if (req.query.generate !== '1') {
        return res.status(404).json({ error: 'Thumbnail nicht im Cache. npm run thumbs ausfuehren.' });
    }

    // 3) Lazy-Generierung explizit angefragt
    try {
        let promise = thumbPending.get(textHash);
        if (!promise) {
            promise = generateThumbnail(row.file_path, row.slide_index, outPath)
                .then(() => {
                    try {
                        db.prepare(`
                            INSERT INTO thumbnails (text_hash, png_path, width, height, generated_at, source_slide_id, source_file_path, source_slide_index)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                            ON CONFLICT(text_hash) DO UPDATE SET
                                png_path=excluded.png_path,
                                width=excluded.width,
                                height=excluded.height,
                                generated_at=excluded.generated_at,
                                source_slide_id=excluded.source_slide_id,
                                source_file_path=excluded.source_file_path,
                                source_slide_index=excluded.source_slide_index
                        `).run(textHash, outPath, 480, 270, Date.now(), row.slide_id, row.file_path, row.slide_index);
                    } catch (dbErr) {
                        console.warn('[THUMB] DB-Eintrag fehlgeschlagen:', dbErr.message);
                    }
                })
                .finally(() => {
                    thumbPending.delete(textHash);
                });
            thumbPending.set(textHash, promise);
        }
        await promise;

        res.setHeader('Cache-Control', 'public, max-age=31536000, immutable');
        return res.sendFile(outPath);
    } catch (err) {
        console.error('[THUMB] Fehler bei Slide', slideId, '(hash', textHash?.slice(0, 8), '):', err.message);
        return res.status(500).json({ error: 'Thumbnail-Erzeugung fehlgeschlagen: ' + err.message });
    }
});

// --- API: Pfad validieren + Vorschau ---
app.post('/api/check-path', (req, res) => {
    const { folderPath } = req.body;
    if (!folderPath) {
        return res.status(400).json({ error: 'Kein Pfad angegeben.' });
    }
    try {
        const stat = fs.statSync(folderPath);
        if (!stat.isDirectory()) {
            return res.json({ valid: false, error: 'Pfad ist kein Ordner.' });
        }
        const files = findPptxFiles(folderPath);
        res.json({
            valid: true,
            fileCount: files.length,
            sample: files.slice(0, 5).map(f => path.basename(f)),
        });
    } catch (err) {
        res.json({ valid: false, error: err.message });
    }
});

// --- API: Nativen Ordner-Dialog oeffnen (Windows + macOS) ---
// Windows: PowerShell mit FolderBrowserDialog, per .ps1-Datei (Norton-vertraeglich).
// macOS:   AppleScript "choose folder" via osascript, gibt POSIX-Pfad zurueck.
app.get('/api/pick-folder', (req, res) => {
    if (process.platform === 'darwin') {
        // AppleScript "choose folder" bringt einen nativen Finder-Dialog hoch.
        // Wir laufen unter "System Events" um den Dialog vor andere Apps zu legen.
        const applescript = `
try
    tell application "System Events"
        activate
        set chosen to choose folder with prompt "Ordner mit PPTX-Dateien waehlen"
    end tell
    return POSIX path of chosen
on error errMsg number errNum
    -- -128 = User-Abbruch, leerer Output -> Cancel-Pfad
    if errNum is -128 then
        return ""
    else
        error errMsg number errNum
    end if
end try
`;
        const proc = spawn('osascript', ['-e', applescript]);
        let stdout = '';
        let stderr = '';
        proc.stdout.on('data', (d) => { stdout += d.toString('utf8'); });
        proc.stderr.on('data', (d) => { stderr += d.toString('utf8'); });
        proc.on('error', (err) => {
            res.status(500).json({ error: 'osascript fehlgeschlagen: ' + err.message });
        });
        proc.on('close', () => {
            const selected = stdout.trim().replace(/\/$/, ''); // Trailing-Slash entfernen
            if (selected) {
                res.json({ path: selected });
            } else {
                res.json({ path: null, cancelled: true });
            }
        });
        return;
    }

    if (process.platform !== 'win32') {
        return res.status(501).json({ error: 'Ordner-Dialog aktuell nur unter Windows und macOS verfuegbar.' });
    }

    // Norton-vertraeglich: Skript in .ps1-Datei schreiben und per -File aufrufen,
    // statt -EncodedCommand (das die IDP.HELU.PSE71-Heuristik triggert).
    const psScript = `
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.Windows.Forms
# Unsichtbare TopMost-Form, damit der Dialog vor dem Browser erscheint
$form = New-Object System.Windows.Forms.Form
$form.TopMost = $true
$form.ShowInTaskbar = $false
$form.Opacity = 0
$form.Size = New-Object System.Drawing.Size(1,1)
$form.StartPosition = 'Manual'
$form.Location = New-Object System.Drawing.Point(-2000,-2000)
$d = New-Object System.Windows.Forms.FolderBrowserDialog
$d.Description = 'Ordner mit PPTX-Dateien waehlen'
$d.ShowNewFolderButton = $false
$result = $d.ShowDialog($form)
if ($result -eq [System.Windows.Forms.DialogResult]::OK) { Write-Output $d.SelectedPath }
$form.Dispose()
`;
    const scriptFile = writeTempPs1(psScript);
    const ps = spawn('powershell.exe', [
        '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
        '-STA', '-File', scriptFile,
    ]);

    let stdout = '';
    let stderr = '';
    ps.stdout.on('data', (d) => { stdout += d.toString('utf8'); });
    ps.stderr.on('data', (d) => { stderr += d.toString('utf8'); });
    ps.on('error', (err) => {
        safeUnlink(scriptFile);
        res.status(500).json({ error: 'PowerShell-Dialog fehlgeschlagen: ' + err.message });
    });
    ps.on('close', () => {
        safeUnlink(scriptFile);
        const selected = stdout.trim();
        if (selected) {
            res.json({ path: selected });
        } else {
            res.json({ path: null, cancelled: true });
        }
    });
});

// --- API: Scan starten ---
// label optional — wenn weggelassen, nimmt der Scanner den letzten Ordnernamen.
app.post('/api/scan', async (req, res) => {
    const { folderPath, label } = req.body;
    if (!folderPath) {
        return res.status(400).json({ error: 'Kein Pfad angegeben.' });
    }
    if (scanState.running) {
        return res.status(409).json({ error: 'Ein Scan laeuft bereits.' });
    }
    // Scan asynchron starten, Response sofort zurueck
    runScan(db, folderPath, label).catch(err => {
        console.error('[SCAN] Unerwarteter Fehler:', err);
    });
    res.json({ ok: true, message: 'Scan gestartet.' });
});

// --- API: Scan-Status (Polling) ---
app.get('/api/scan/status', (req, res) => {
    res.json({
        ...scanState,
        duration: scanState.startedAt
            ? (scanState.finishedAt || Date.now()) - scanState.startedAt
            : 0,
    });
});

// --- API: Volltextsuche ---
// FTS5-Syntax: Woerter werden mit Leerzeichen AND-verknuepft.
// Wir maskieren Sonderzeichen indem wir jedes Wort in Quotes packen.
function buildFtsQuery(term) {
    const words = term.trim().split(/\s+/).filter(w => w.length > 0);
    if (words.length === 0) return null;
    return words.map(w => `"${w.replace(/"/g, '""')}"`).join(' ');
}

// Jahr-Parameter in mtime-Range umrechnen.
// Akzeptiert:
//   'all'          -> kein Filter
//   '<year>'       -> genau dieses Jahr
//   'le<year>'     -> Jahr und alles davor (inklusive)
function yearToMtimeRange(year) {
    if (!year || year === 'all') return null;
    if (typeof year === 'string' && year.startsWith('le')) {
        const y = parseInt(year.slice(2), 10);
        if (!Number.isFinite(y)) return null;
        return { min: null, max: Date.UTC(y + 1, 0, 1) };
    }
    const y = parseInt(year, 10);
    if (!Number.isFinite(y)) return null;
    return { min: Date.UTC(y, 0, 1), max: Date.UTC(y + 1, 0, 1) };
}

// Liste der Jahres-Pills aus Konfig + aktuellem Jahr ableiten.
// Beispiel: yearsBack=3 in 2026 -> ['Alle', '≤ 2023', '2024', '2025', '2026']
function buildYearFilterOptions(yearsBack) {
    const current = new Date().getFullYear();
    const cutoff = current - yearsBack;
    const out = [
        { value: 'all',          label: 'Alle' },
        { value: 'le' + cutoff,  label: '≤ ' + cutoff },
    ];
    for (let y = cutoff + 1; y <= current; y++) {
        out.push({ value: String(y), label: String(y) });
    }
    return out;
}

app.get('/api/search', (req, res) => {
    const q = (req.query.q || '').toString().trim();
    const uniqueOnly = req.query.uniqueOnly === '1' || req.query.uniqueOnly === 'true';
    const year       = (req.query.year || '').toString();
    const filename   = (req.query.filename || '').toString().trim();
    const rootsParam = (req.query.roots || '').toString();
    const sort       = (req.query.sort || 'relevance').toString();

    if (!q) return res.json({ query: '', groups: [], totalGroups: 0, totalOccurrences: 0 });

    const ftsQuery = buildFtsQuery(q);
    if (!ftsQuery) return res.json({ query: q, groups: [], totalGroups: 0, totalOccurrences: 0 });

    try {
        // SQL-Filter dynamisch zusammensetzen
        const where = ['slides_fts MATCH ?'];
        const params = [ftsQuery];

        const yr = yearToMtimeRange(year);
        if (yr) {
            if (yr.min !== null) { where.push('p.mtime >= ?'); params.push(yr.min); }
            if (yr.max !== null) { where.push('p.mtime < ?');  params.push(yr.max); }
        }

        if (filename) {
            // Case-insensitive Substring-Suche im Dateinamen
            where.push('LOWER(p.file_name) LIKE LOWER(?)');
            params.push('%' + filename + '%');
        }

        // Hauptpfad-Filter: Komma-separierte Liste voller Praefixe.
        // Wenn nichts angegeben -> kein Filter (alles zugelassen).
        const roots = rootsParam ? rootsParam.split('|').map(s => s.trim()).filter(Boolean) : [];
        if (roots.length > 0) {
            const placeholders = roots.map(() => 'LOWER(p.file_path) LIKE LOWER(?)').join(' OR ');
            where.push('(' + placeholders + ')');
            for (const r of roots) params.push(r + '%');
        }

        // Sortierung: Relevanz (rank) oder Datum
        let orderBy = 'ORDER BY rank';
        if (sort === 'date_desc') orderBy = 'ORDER BY p.mtime DESC, rank';
        else if (sort === 'date_asc') orderBy = 'ORDER BY p.mtime ASC, rank';

        const stmt = db.prepare(`
            SELECT
                s.id, s.slide_index, s.title, s.text_content,
                s.text_hash, s.exact_hash,
                p.id AS presentation_id, p.file_name, p.file_path, p.mtime,
                snippet(slides_fts, 1, '<mark>', '</mark>', '...', 20) AS snippet,
                rank AS fts_rank
            FROM slides_fts
            JOIN slides s ON s.id = slides_fts.rowid
            JOIN presentations p ON p.id = s.presentation_id
            WHERE ${where.join(' AND ')}
            ${orderBy}
            LIMIT ${config.search.maxFtsRows}
        `);
        const rows = stmt.all(...params);

        // 2) Nach text_hash gruppieren. Erste Zeile pro Gruppe (beste FTS-Rank)
        //    wird Repraesentant, die restlichen Vorkommen hinten dran.
        const groupsMap = new Map();
        for (const row of rows) {
            const key = row.text_hash;
            if (!groupsMap.has(key)) {
                groupsMap.set(key, {
                    textHash: row.text_hash,
                    representative: {
                        slideId: row.id,
                        slideIndex: row.slide_index,
                        title: row.title,
                        snippet: row.snippet,
                        presentationId: row.presentation_id,
                        fileName: row.file_name,
                        filePath: row.file_path,
                        mtime: row.mtime,
                    },
                    occurrences: [],
                    fileSet: new Set(),
                });
            }
            const g = groupsMap.get(key);
            g.occurrences.push({
                slideId: row.id,
                slideIndex: row.slide_index,
                title: row.title,
                presentationId: row.presentation_id,
                fileName: row.file_name,
                filePath: row.file_path,
                mtime: row.mtime,
            });
            g.fileSet.add(row.presentation_id);
        }

        // 3) In Array umwandeln, fileSet -> fileCount, evtl. filtern
        let groups = Array.from(groupsMap.values()).map(g => ({
            textHash: g.textHash,
            representative: g.representative,
            occurrences: g.occurrences,
            count: g.occurrences.length,
            fileCount: g.fileSet.size,
        }));

        const totalOccurrences = rows.length;
        const totalGroups = groups.length;

        if (uniqueOnly) {
            groups = groups.filter(g => g.count === 1);
        }

        // Pagination: Default aus config.search.pageSize.
        const offset = Math.max(0, parseInt(req.query.offset, 10) || 0);
        const limit  = Math.min(200, Math.max(1, parseInt(req.query.limit, 10) || config.search.pageSize));
        const limited = groups.slice(offset, offset + limit);

        res.json({
            query: q,
            uniqueOnly,
            groups: limited,
            totalGroups: groups.length,        // nach uniqueOnly-Filter
            totalGroupsBeforeFilter: totalGroups,
            totalOccurrences,
            shownGroups: limited.length,
            offset,
            limit,
            hasMore: offset + limited.length < groups.length,
        });
    } catch (err) {
        res.status(400).json({ error: 'Suchfehler: ' + err.message });
    }
});

// --- Server starten ---
app.listen(PORT, HOST, () => {
    console.log('');
    console.log('=========================================');
    console.log('  pptx-manager - lokal gestartet');
    console.log('=========================================');
    console.log(`  URL: http://${HOST}:${PORT}`);
    console.log('  Zum Beenden: Strg+C');
    console.log('=========================================');
    console.log('');
});
