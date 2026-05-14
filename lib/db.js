// =============================================================
// db.js - SQLite-Datenbank Setup
// =============================================================
// Nutzt das in Node.js 22+ eingebaute node:sqlite Modul.
// Keine externen Dependencies noetig, kein Build-Prozess.
//
// Schema:
//   presentations  - Eine Zeile pro PPTX-Datei
//   slides         - Eine Zeile pro Folie, mit 3 Hash-Ebenen
//   slides_fts     - Volltextindex (FTS5) fuer schnelle Suche
// =============================================================

const { DatabaseSync } = require('node:sqlite');

function initDb(dbPath) {
    const db = new DatabaseSync(dbPath);
    db.exec('PRAGMA journal_mode = WAL;');
    db.exec('PRAGMA foreign_keys = ON;');

    // --- Tabelle: Meta (key/value) ---
    // Speichert u.a. fingerprint_version (Hash-Algorithmus-Version).
    db.exec(`
        CREATE TABLE IF NOT EXISTS meta (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        );
    `);

    // --- Tabelle: Scan-Roots (vom Nutzer gewaehlte Top-Level-Pfade) ---
    // Pro Root ein editierbares Label fuer die UI. Loeschen kaskadiert in
    // presentations (und damit slides) ueber den FK weiter unten.
    db.exec(`
        CREATE TABLE IF NOT EXISTS scan_roots (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            root_path TEXT UNIQUE NOT NULL,
            label TEXT NOT NULL,
            created_at INTEGER NOT NULL,
            last_scanned_at INTEGER
        );
    `);

    // --- Tabelle: Praesentationen ---
    db.exec(`
        CREATE TABLE IF NOT EXISTS presentations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_path TEXT UNIQUE NOT NULL,
            file_name TEXT NOT NULL,
            file_hash TEXT NOT NULL,
            file_size INTEGER NOT NULL,
            mtime INTEGER NOT NULL,
            slide_count INTEGER NOT NULL DEFAULT 0,
            scanned_at INTEGER NOT NULL,
            scan_root_id INTEGER REFERENCES scan_roots(id) ON DELETE CASCADE
        );
    `);
    // Migration: scan_root_id-Spalte nachruesten, falls aelteres DB-Schema.
    const presCols = db.prepare("PRAGMA table_info(presentations)").all();
    if (presCols.length > 0 && !presCols.some(c => c.name === 'scan_root_id')) {
        db.exec('ALTER TABLE presentations ADD COLUMN scan_root_id INTEGER REFERENCES scan_roots(id) ON DELETE CASCADE;');
    }
    db.exec(`CREATE INDEX IF NOT EXISTS idx_pres_scan_root ON presentations(scan_root_id);`);

    // --- Tabelle: Folien ---
    db.exec(`
        CREATE TABLE IF NOT EXISTS slides (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            presentation_id INTEGER NOT NULL,
            slide_index INTEGER NOT NULL,
            title TEXT,
            text_content TEXT,
            exact_hash TEXT NOT NULL,
            text_hash TEXT NOT NULL,
            structure_hash TEXT NOT NULL,
            FOREIGN KEY(presentation_id) REFERENCES presentations(id) ON DELETE CASCADE
        );
    `);

    // --- Tabelle: Thumbnails ---
    // KEY ist text_hash, nicht slide_id: identische Folien (gleicher Inhalts-Hash)
    // teilen sich ein PNG. So wird nur einmal pro Inhalt gerendert und gespeichert.
    //
    // Migration: falls die alte Tabelle mit slide_id als PK existiert, droppen.
    // Das ist verlustfrei, weil Thumbnails jederzeit wieder erzeugbar sind.
    const thumbCols = db.prepare("PRAGMA table_info(thumbnails)").all();
    if (thumbCols.length > 0 && !thumbCols.some(c => c.name === 'text_hash')) {
        console.log('[DB] Alte thumbnails-Tabelle erkannt -> wird neu angelegt (text_hash-basiert).');
        db.exec('DROP TABLE thumbnails;');
    }

    db.exec(`
        CREATE TABLE IF NOT EXISTS thumbnails (
            text_hash TEXT PRIMARY KEY,
            png_path TEXT NOT NULL,
            width INTEGER NOT NULL,
            height INTEGER NOT NULL,
            generated_at INTEGER NOT NULL,
            source_slide_id INTEGER,
            source_file_path TEXT,
            source_slide_index INTEGER
        );
    `);

    // --- Indizes fuer schnelle Suche ---
    db.exec(`CREATE INDEX IF NOT EXISTS idx_slides_exact ON slides(exact_hash);`);
    db.exec(`CREATE INDEX IF NOT EXISTS idx_slides_text ON slides(text_hash);`);
    db.exec(`CREATE INDEX IF NOT EXISTS idx_slides_structure ON slides(structure_hash);`);
    db.exec(`CREATE INDEX IF NOT EXISTS idx_slides_pres ON slides(presentation_id);`);

    // --- Volltext-Suche (FTS5) ---
    db.exec(`
        CREATE VIRTUAL TABLE IF NOT EXISTS slides_fts USING fts5(
            title, text_content,
            content='slides',
            content_rowid='id',
            tokenize='unicode61 remove_diacritics 2'
        );
    `);

    // --- Trigger: FTS synchron halten ---
    db.exec(`
        CREATE TRIGGER IF NOT EXISTS slides_ai AFTER INSERT ON slides BEGIN
            INSERT INTO slides_fts(rowid, title, text_content)
            VALUES (new.id, new.title, new.text_content);
        END;
    `);
    db.exec(`
        CREATE TRIGGER IF NOT EXISTS slides_ad AFTER DELETE ON slides BEGIN
            INSERT INTO slides_fts(slides_fts, rowid, title, text_content)
            VALUES ('delete', old.id, old.title, old.text_content);
        END;
    `);
    db.exec(`
        CREATE TRIGGER IF NOT EXISTS slides_au AFTER UPDATE ON slides BEGIN
            INSERT INTO slides_fts(slides_fts, rowid, title, text_content)
            VALUES ('delete', old.id, old.title, old.text_content);
            INSERT INTO slides_fts(rowid, title, text_content)
            VALUES (new.id, new.title, new.text_content);
        END;
    `);

    return db;
}

/**
 * Liest die gespeicherte Fingerprint-Version aus der DB.
 * Gibt 0 zurueck, wenn noch keine gesetzt ist (leere DB).
 */
function getStoredFingerprintVersion(db) {
    const row = db.prepare("SELECT value FROM meta WHERE key = 'fingerprint_version'").get();
    return row ? parseInt(row.value, 10) : 0;
}

/**
 * Setzt die aktuelle Fingerprint-Version in der DB.
 */
function setFingerprintVersion(db, version) {
    db.prepare(`
        INSERT INTO meta (key, value) VALUES ('fingerprint_version', ?)
        ON CONFLICT(key) DO UPDATE SET value = excluded.value
    `).run(String(version));
}

/**
 * Prueft, ob die vorhandene DB mit der aktuellen Fingerprint-Version kompatibel ist.
 * Gibt { compatible, storedVersion, hasData } zurueck.
 * Wenn hasData=false, ist alles in Ordnung (nichts zu migrieren).
 * Wenn storedVersion < codeVersion UND hasData=true, ist ein Neu-Scan noetig.
 */
function checkFingerprintCompatibility(db, codeVersion) {
    const storedVersion = getStoredFingerprintVersion(db);
    const slideCount = db.prepare('SELECT COUNT(*) AS c FROM slides').get().c;
    const hasData = slideCount > 0;
    return {
        compatible: !hasData || storedVersion === codeVersion,
        storedVersion,
        codeVersion,
        hasData,
        slideCount,
    };
}

/**
 * Loescht alle gescannten Daten (slides + presentations) und setzt die
 * Fingerprint-Version neu. Wird gebraucht, wenn sich der Hash-Algorithmus
 * geaendert hat. FTS-Tabelle wird per Trigger automatisch geleert.
 */
function clearScanData(db) {
    // Thumbnails werden per text_hash referenziert und sind vom Scan-Zustand
    // unabhaengig — leeren ist aber sauber, weil bei Reset auch die PNG-Dateien
    // geloescht werden.
    db.exec('DELETE FROM thumbnails;');
    db.exec('DELETE FROM slides;');
    db.exec('DELETE FROM presentations;');
    db.exec('DELETE FROM scan_roots;');
    db.exec("DELETE FROM sqlite_sequence WHERE name IN ('slides','presentations','scan_roots');");
}

// --- Scan-Root-Verwaltung ---

function listScanRoots(db) {
    return db.prepare(`
        SELECT sr.id, sr.root_path AS rootPath, sr.label,
               sr.created_at AS createdAt, sr.last_scanned_at AS lastScannedAt,
               COUNT(p.id) AS presentationCount,
               COALESCE(SUM(p.slide_count), 0) AS slideCount
        FROM scan_roots sr
        LEFT JOIN presentations p ON p.scan_root_id = sr.id
        GROUP BY sr.id
        ORDER BY sr.label COLLATE NOCASE
    `).all();
}

function getScanRoot(db, id) {
    return db.prepare(`
        SELECT id, root_path AS rootPath, label, created_at AS createdAt, last_scanned_at AS lastScannedAt
        FROM scan_roots WHERE id = ?
    `).get(id);
}

// Liefert id eines existierenden Roots (case-insensitive Path-Match auf Windows)
// oder legt einen neuen mit dem gegebenen Label an.
function getOrCreateScanRoot(db, rootPath, label) {
    const existing = db.prepare(
        'SELECT id FROM scan_roots WHERE LOWER(root_path) = LOWER(?)'
    ).get(rootPath);
    if (existing) return existing.id;
    const result = db.prepare(`
        INSERT INTO scan_roots (root_path, label, created_at)
        VALUES (?, ?, ?)
    `).run(rootPath, label, Date.now());
    return Number(result.lastInsertRowid);
}

function updateScanRootLabel(db, id, label) {
    return db.prepare('UPDATE scan_roots SET label = ? WHERE id = ?').run(label, id).changes;
}

function setScanRootLastScannedAt(db, id) {
    db.prepare('UPDATE scan_roots SET last_scanned_at = ? WHERE id = ?').run(Date.now(), id);
}

function deleteScanRoot(db, id) {
    return db.prepare('DELETE FROM scan_roots WHERE id = ?').run(id).changes;
}

function listPresentationsInRoot(db, scanRootId) {
    return db.prepare(`
        SELECT id, file_path AS filePath, file_name AS fileName
        FROM presentations WHERE scan_root_id = ?
    `).all(scanRootId);
}

function deletePresentationsByIds(db, ids) {
    if (!ids || ids.length === 0) return 0;
    const placeholders = ids.map(() => '?').join(',');
    return db.prepare(`DELETE FROM presentations WHERE id IN (${placeholders})`).run(...ids).changes;
}

// Thumbs ohne zugehoerige Folien: Eintraege in DB loeschen + PNG-Pfade
// zurueckgeben, damit der Aufrufer die Dateien auf der Platte aufraeumen kann.
function garbageCollectThumbnails(db) {
    const orphans = db.prepare(`
        SELECT text_hash AS textHash, png_path AS pngPath
        FROM thumbnails
        WHERE text_hash NOT IN (SELECT DISTINCT text_hash FROM slides)
    `).all();
    if (orphans.length > 0) {
        db.prepare(`
            DELETE FROM thumbnails
            WHERE text_hash NOT IN (SELECT DISTINCT text_hash FROM slides)
        `).run();
    }
    return orphans;
}

module.exports = {
    initDb,
    getStoredFingerprintVersion,
    setFingerprintVersion,
    checkFingerprintCompatibility,
    clearScanData,
    listScanRoots,
    getScanRoot,
    getOrCreateScanRoot,
    updateScanRootLabel,
    setScanRootLastScannedAt,
    deleteScanRoot,
    listPresentationsInRoot,
    deletePresentationsByIds,
    garbageCollectThumbnails,
};
