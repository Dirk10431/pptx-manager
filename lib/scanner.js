// =============================================================
// scanner.js - Ordner scannen, PPTX-Dateien indizieren
// =============================================================
// Inkrementell: Dateien mit unveraenderter mtime + Groesse werden
// nicht erneut analysiert. Lock-Dateien (~$xxx.pptx) ueberspringen.
// =============================================================

const fs = require('fs');
const path = require('path');
const { analyzePptx, hashFile, FINGERPRINT_VERSION } = require('./fingerprint');
const { setFingerprintVersion } = require('./db');

/**
 * Rekursiv alle .pptx-Dateien in einem Ordner finden.
 * Ignoriert:
 *   - Lock-Dateien (beginnen mit ~$)
 *   - Versteckte Ordner (.git, .venv, node_modules, _old, ...)
 */
function findPptxFiles(rootDir) {
    const results = [];
    const IGNORE_DIRS = new Set(['node_modules', '.git', '.venv', 'venv', '_old', 'backups', 'data']);

    function walk(dir) {
        let entries;
        try {
            entries = fs.readdirSync(dir, { withFileTypes: true });
        } catch (err) {
            // z.B. Berechtigungsfehler - ueberspringen
            return;
        }
        for (const entry of entries) {
            const fullPath = path.join(dir, entry.name);
            if (entry.isDirectory()) {
                if (entry.name.startsWith('.')) continue;
                if (IGNORE_DIRS.has(entry.name)) continue;
                walk(fullPath);
            } else if (entry.isFile()) {
                if (entry.name.startsWith('~$')) continue; // Office-Lock-Datei
                if (entry.name.toLowerCase().endsWith('.pptx')) {
                    results.push(fullPath);
                }
            }
        }
    }

    walk(rootDir);
    return results;
}

/**
 * Scan-Status fuer Live-Progress (wird in server.js geteilt).
 */
const scanState = {
    running: false,
    rootDir: null,
    total: 0,
    processed: 0,
    added: 0,
    updated: 0,
    skipped: 0,
    failed: 0,
    currentFile: null,
    startedAt: null,
    finishedAt: null,
    lastError: null,
};

function resetScanState() {
    scanState.running = false;
    scanState.rootDir = null;
    scanState.total = 0;
    scanState.processed = 0;
    scanState.added = 0;
    scanState.updated = 0;
    scanState.skipped = 0;
    scanState.failed = 0;
    scanState.currentFile = null;
    scanState.startedAt = null;
    scanState.finishedAt = null;
    scanState.lastError = null;
}

/**
 * Eine einzelne PPTX verarbeiten: Fingerprint berechnen, in DB speichern.
 * Gibt zurueck: 'added' | 'updated' | 'skipped' | 'failed'
 */
async function processPptx(db, filePath) {
    try {
        const stat = fs.statSync(filePath);
        const mtime = Math.floor(stat.mtimeMs);
        const fileName = path.basename(filePath);

        // Existiert die Datei schon in der DB?
        const existing = db.prepare('SELECT id, mtime, file_size FROM presentations WHERE file_path = ?').get(filePath);

        if (existing && existing.mtime === mtime && existing.file_size === stat.size) {
            return 'skipped'; // Unveraendert
        }

        // Fingerprint berechnen
        const fileHash = await hashFile(filePath);
        const { slides } = await analyzePptx(filePath);
        const now = Date.now();

        // In DB schreiben - in Transaktion
        const insertPres = db.prepare(`
            INSERT INTO presentations (file_path, file_name, file_hash, file_size, mtime, slide_count, scanned_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        `);
        const updatePres = db.prepare(`
            UPDATE presentations SET file_name=?, file_hash=?, file_size=?, mtime=?, slide_count=?, scanned_at=?
            WHERE id = ?
        `);
        const deleteSlides = db.prepare('DELETE FROM slides WHERE presentation_id = ?');
        const insertSlide = db.prepare(`
            INSERT INTO slides (presentation_id, slide_index, title, text_content, exact_hash, text_hash, structure_hash)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        `);

        db.exec('BEGIN');
        try {
            let presId;
            if (existing) {
                updatePres.run(fileName, fileHash, stat.size, mtime, slides.length, now, existing.id);
                deleteSlides.run(existing.id);
                presId = existing.id;
            } else {
                const result = insertPres.run(filePath, fileName, fileHash, stat.size, mtime, slides.length, now);
                presId = result.lastInsertRowid;
            }
            for (const slide of slides) {
                insertSlide.run(
                    presId,
                    slide.slideIndex,
                    slide.title,
                    slide.text,
                    slide.exactHash,
                    slide.textHash,
                    slide.structureHash
                );
            }
            db.exec('COMMIT');
        } catch (err) {
            db.exec('ROLLBACK');
            throw err;
        }

        return existing ? 'updated' : 'added';
    } catch (err) {
        scanState.lastError = `${path.basename(filePath)}: ${err.message}`;
        console.error(`[SCAN] Fehler bei ${filePath}:`, err.message);
        return 'failed';
    }
}

/**
 * Scan eines Ordners starten. Laeuft asynchron.
 * Fortschritt wird in scanState aktualisiert.
 */
async function runScan(db, rootDir) {
    if (scanState.running) {
        throw new Error('Ein Scan laeuft bereits.');
    }
    if (!fs.existsSync(rootDir)) {
        throw new Error(`Ordner nicht gefunden: ${rootDir}`);
    }

    resetScanState();
    scanState.running = true;
    scanState.rootDir = rootDir;
    scanState.startedAt = Date.now();

    try {
        const files = findPptxFiles(rootDir);
        scanState.total = files.length;
        console.log(`[SCAN] ${files.length} PPTX-Dateien gefunden in ${rootDir}`);

        for (const file of files) {
            scanState.currentFile = path.basename(file);
            const result = await processPptx(db, file);
            scanState[result] = (scanState[result] || 0) + 1;
            scanState.processed++;
        }

        // Nach erfolgreichem Scan aktuelle Fingerprint-Version in der DB festhalten
        setFingerprintVersion(db, FINGERPRINT_VERSION);

        scanState.finishedAt = Date.now();
        console.log(`[SCAN] Fertig: ${scanState.added} neu, ${scanState.updated} aktualisiert, ${scanState.skipped} unveraendert, ${scanState.failed} Fehler`);
    } catch (err) {
        scanState.lastError = err.message;
        scanState.finishedAt = Date.now();
        console.error('[SCAN] Abbruch:', err);
    } finally {
        scanState.running = false;
        scanState.currentFile = null;
    }
}

module.exports = {
    findPptxFiles,
    runScan,
    scanState,
};
