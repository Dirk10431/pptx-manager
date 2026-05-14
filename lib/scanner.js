// =============================================================
// scanner.js - Ordner scannen, PPTX-Dateien indizieren
// =============================================================
// Inkrementell: Dateien mit unveraenderter mtime + Groesse werden
// nicht erneut analysiert. Lock-Dateien (~$xxx.pptx) ueberspringen.
// =============================================================

const fs = require('fs');
const path = require('path');
const { analyzePptx, hashFile, FINGERPRINT_VERSION } = require('./fingerprint');
const {
    setFingerprintVersion,
    getOrCreateScanRoot,
    setScanRootLastScannedAt,
    listPresentationsInRoot,
} = require('./db');

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
    scanRootId: null,
    label: null,
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
    // Praesentationen, die im DB unter diesem Root stehen, aber beim Scan
    // nicht mehr auf der Platte gefunden wurden. Bestaetigungs-Workflow im UI.
    staleFiles: [],
};

function resetScanState() {
    scanState.running = false;
    scanState.rootDir = null;
    scanState.scanRootId = null;
    scanState.label = null;
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
    scanState.staleFiles = [];
}

/**
 * Eine einzelne PPTX verarbeiten: Fingerprint berechnen, in DB speichern.
 * Gibt zurueck: 'added' | 'updated' | 'skipped' | 'failed'
 */
async function processPptx(db, filePath, scanRootId) {
    try {
        const stat = fs.statSync(filePath);
        const mtime = Math.floor(stat.mtimeMs);
        const fileName = path.basename(filePath);

        // Existiert die Datei schon in der DB?
        const existing = db.prepare(
            'SELECT id, mtime, file_size, scan_root_id FROM presentations WHERE file_path = ?'
        ).get(filePath);

        if (existing && existing.mtime === mtime && existing.file_size === stat.size) {
            // Unveraendert. Aber falls die Datei jetzt unter einem anderen Root
            // gescannt wird (z.B. weil ihre Root-Zuordnung gewechselt hat),
            // korrigieren wir die Zuordnung still im Hintergrund.
            if (existing.scan_root_id !== scanRootId) {
                db.prepare('UPDATE presentations SET scan_root_id = ? WHERE id = ?').run(scanRootId, existing.id);
            }
            return 'skipped';
        }

        // Fingerprint berechnen
        const fileHash = await hashFile(filePath);
        const { slides } = await analyzePptx(filePath);
        const now = Date.now();

        // In DB schreiben - in Transaktion
        const insertPres = db.prepare(`
            INSERT INTO presentations (file_path, file_name, file_hash, file_size, mtime, slide_count, scanned_at, scan_root_id)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        `);
        const updatePres = db.prepare(`
            UPDATE presentations SET file_name=?, file_hash=?, file_size=?, mtime=?, slide_count=?, scanned_at=?, scan_root_id=?
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
                updatePres.run(fileName, fileHash, stat.size, mtime, slides.length, now, scanRootId, existing.id);
                deleteSlides.run(existing.id);
                presId = existing.id;
            } else {
                const result = insertPres.run(filePath, fileName, fileHash, stat.size, mtime, slides.length, now, scanRootId);
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
 * Bindet alle gefundenen Praesentationen an einen scan_root (mit Label).
 * Erkennt am Ende stale files: in der DB unter diesem Root, aber nicht mehr
 * auf der Platte. Diese werden NICHT auto-geloescht, sondern in scanState.staleFiles
 * gesammelt — das Frontend zeigt sie zur Bestaetigung.
 */
async function runScan(db, rootDir, label) {
    if (scanState.running) {
        throw new Error('Ein Scan laeuft bereits.');
    }
    if (!fs.existsSync(rootDir)) {
        throw new Error(`Ordner nicht gefunden: ${rootDir}`);
    }

    const effectiveLabel = (label && label.trim()) || path.basename(rootDir) || rootDir;
    const scanRootId = getOrCreateScanRoot(db, rootDir, effectiveLabel);

    resetScanState();
    scanState.running = true;
    scanState.rootDir = rootDir;
    scanState.scanRootId = scanRootId;
    scanState.label = effectiveLabel;
    scanState.startedAt = Date.now();

    try {
        const files = findPptxFiles(rootDir);
        scanState.total = files.length;
        console.log(`[SCAN] ${files.length} PPTX-Dateien gefunden in ${rootDir} (root #${scanRootId} "${effectiveLabel}")`);

        // Set der gefundenen Pfade (lowercase) fuer Stale-Erkennung am Ende.
        const seenPaths = new Set(files.map(f => f.toLowerCase()));

        for (const file of files) {
            scanState.currentFile = path.basename(file);
            const result = await processPptx(db, file, scanRootId);
            scanState[result] = (scanState[result] || 0) + 1;
            scanState.processed++;
        }

        // Stale-Check: alle Praesentationen unter diesem Root, deren Datei-Pfad
        // beim aktuellen Scan nicht aufgetaucht ist. Heisst entweder Datei
        // geloescht/verschoben oder OneDrive war kurz aus dem Tritt.
        const allInRoot = listPresentationsInRoot(db, scanRootId);
        scanState.staleFiles = allInRoot
            .filter(p => !seenPaths.has(p.filePath.toLowerCase()))
            .map(p => ({ id: p.id, filePath: p.filePath, fileName: p.fileName }));

        // Nach erfolgreichem Scan aktuelle Fingerprint-Version in der DB festhalten
        setFingerprintVersion(db, FINGERPRINT_VERSION);
        setScanRootLastScannedAt(db, scanRootId);

        scanState.finishedAt = Date.now();
        console.log(`[SCAN] Fertig: ${scanState.added} neu, ${scanState.updated} aktualisiert, ${scanState.skipped} unveraendert, ${scanState.failed} Fehler, ${scanState.staleFiles.length} stale`);
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
