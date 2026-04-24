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
            scanned_at INTEGER NOT NULL
        );
    `);

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

module.exports = { initDb };
