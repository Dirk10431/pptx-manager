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
const { initDb, checkFingerprintCompatibility, setFingerprintVersion, clearScanData } = require('./lib/db');
const { runScan, scanState, findPptxFiles } = require('./lib/scanner');
const { FINGERPRINT_VERSION } = require('./lib/fingerprint');
const { generateThumbnail, thumbnailCachePath } = require('./lib/thumbnailer');
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
// --- API: Frontend-relevante Config-Werte ---
// Wird beim Laden der Seite einmal gefetched, damit Lightbox-Groesse,
// Page-Size etc. zentral aus config.json kommen.
app.get('/api/config', (req, res) => {
    res.json({
        search: {
            pageSize: config.search.pageSize,
            occurrencesShown: config.search.occurrencesShown,
            minQueryLength: config.search.minQueryLength,
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
    const c = checkFingerprintCompatibility(db, FINGERPRINT_VERSION);
    res.json({
        presentations: presentationCount,
        slides: slideCount,
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

// --- API: Nativen Windows-Ordner-Dialog oeffnen ---
// Spawnt PowerShell mit FolderBrowserDialog. Nur unter Windows nutzbar.
// Dialog wird per Dummy-Form mit TopMost nach vorn gebracht.
app.get('/api/pick-folder', (req, res) => {
    if (process.platform !== 'win32') {
        return res.status(501).json({ error: 'Ordner-Dialog aktuell nur unter Windows verfuegbar.' });
    }

    const psScript = `
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.Windows.Forms
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
    const encoded = Buffer.from(psScript, 'utf16le').toString('base64');
    const ps = spawn('powershell.exe', ['-NoProfile', '-STA', '-EncodedCommand', encoded]);

    let stdout = '';
    let stderr = '';
    ps.stdout.on('data', (d) => { stdout += d.toString('utf8'); });
    ps.stderr.on('data', (d) => { stderr += d.toString('utf8'); });
    ps.on('error', (err) => {
        res.status(500).json({ error: 'PowerShell-Dialog fehlgeschlagen: ' + err.message });
    });
    ps.on('close', () => {
        const selected = stdout.trim();
        if (selected) {
            res.json({ path: selected });
        } else {
            res.json({ path: null, cancelled: true });
        }
    });
});

// --- API: Scan starten ---
app.post('/api/scan', async (req, res) => {
    const { folderPath } = req.body;
    if (!folderPath) {
        return res.status(400).json({ error: 'Kein Pfad angegeben.' });
    }
    if (scanState.running) {
        return res.status(409).json({ error: 'Ein Scan laeuft bereits.' });
    }
    // Scan asynchron starten, Response sofort zurueck
    runScan(db, folderPath).catch(err => {
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

app.get('/api/search', (req, res) => {
    const q = (req.query.q || '').toString().trim();
    // Query-Parameter: uniqueOnly=1 blendet alle Gruppen mit >1 Vorkommen aus
    // (nur Folien, die genau einmal vorkommen — typisch: deine individuellen Edits).
    const uniqueOnly = req.query.uniqueOnly === '1' || req.query.uniqueOnly === 'true';

    if (!q) return res.json({ query: '', groups: [], totalGroups: 0, totalOccurrences: 0 });

    const ftsQuery = buildFtsQuery(q);
    if (!ftsQuery) return res.json({ query: q, groups: [], totalGroups: 0, totalOccurrences: 0 });

    try {
        // 1) Alle Treffer holen, nach FTS-Rank sortiert (beste zuerst).
        //    Limit vorm Gruppieren hoch angesetzt, damit wir pro Gruppe nicht
        //    abgeschnittene Vorkommen bekommen.
        const stmt = db.prepare(`
            SELECT
                s.id, s.slide_index, s.title, s.text_content,
                s.text_hash, s.exact_hash,
                p.id AS presentation_id, p.file_name, p.file_path,
                snippet(slides_fts, 1, '<mark>', '</mark>', '...', 20) AS snippet,
                rank AS fts_rank
            FROM slides_fts
            JOIN slides s ON s.id = slides_fts.rowid
            JOIN presentations p ON p.id = s.presentation_id
            WHERE slides_fts MATCH ?
            ORDER BY rank
            LIMIT ${config.search.maxFtsRows}
        `);
        const rows = stmt.all(ftsQuery);

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
