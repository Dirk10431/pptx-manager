// =============================================================
// pptx-manager - Lokaler Server
// =============================================================
// Verwaltet PPTX-Sammlungen, findet Duplikate, ermoeglicht
// Volltextsuche. Laeuft nur lokal auf 127.0.0.1.
// =============================================================

const express = require('express');
const path = require('path');
const fs = require('fs');
const { initDb } = require('./lib/db');
const { runScan, scanState, findPptxFiles } = require('./lib/scanner');

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

// --- API: Status/Statistik ---
app.get('/api/stats', (req, res) => {
    const presentationCount = db.prepare('SELECT COUNT(*) AS c FROM presentations').get().c;
    const slideCount = db.prepare('SELECT COUNT(*) AS c FROM slides').get().c;
    res.json({
        presentations: presentationCount,
        slides: slideCount,
        dbPath: dbPath
    });
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
    if (!q) return res.json({ query: '', results: [] });

    const ftsQuery = buildFtsQuery(q);
    if (!ftsQuery) return res.json({ query: q, results: [] });

    try {
        const stmt = db.prepare(`
            SELECT
                s.id, s.slide_index, s.title, s.text_content,
                p.id AS presentation_id, p.file_name, p.file_path,
                snippet(slides_fts, 1, '<mark>', '</mark>', '...', 20) AS snippet
            FROM slides_fts
            JOIN slides s ON s.id = slides_fts.rowid
            JOIN presentations p ON p.id = s.presentation_id
            WHERE slides_fts MATCH ?
            ORDER BY rank
            LIMIT 100
        `);
        const results = stmt.all(ftsQuery);
        res.json({ query: q, results });
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
