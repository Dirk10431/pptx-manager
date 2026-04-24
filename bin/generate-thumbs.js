#!/usr/bin/env node
// =============================================================
// bin/generate-thumbs.js - Batch-Generator fuer Folien-Thumbnails
// =============================================================
// Geht die DB durch, findet alle Folien ohne Thumbnail, gruppiert
// nach Praesentation und laesst PowerPoint pro Datei nur einmal
// hochfahren. Schreibt PNG-Dateien in data/thumbnails/<slideId>.png
// und pflegt die Tabelle `thumbnails`.
//
// Nutzung:
//   node bin/generate-thumbs.js              alle fehlenden erzeugen
//   node bin/generate-thumbs.js --limit 50   nur 50 Folien
//   node bin/generate-thumbs.js --force      auch vorhandene neu
//   node bin/generate-thumbs.js --dry-run    nur zeigen, was zu tun waere
//
// Abbruch mit Strg+C: PowerPoint wird sauber beendet, beim naechsten
// Lauf macht das Tool an der Abbruchstelle weiter (Resume).
// =============================================================

const fs = require('fs');
const path = require('path');

const { initDb } = require('../lib/db');
const { generateThumbnailsBatch, thumbnailCachePath, THUMB_WIDTH, THUMB_HEIGHT } = require('../lib/thumbnailer');

// -------- CLI-Parameter --------
const args = process.argv.slice(2);
const opts = {
    limit: null,
    force: false,
    dryRun: false,
    help: false,
};
for (let i = 0; i < args.length; i++) {
    const a = args[i];
    if (a === '--help' || a === '-h') opts.help = true;
    else if (a === '--force') opts.force = true;
    else if (a === '--dry-run') opts.dryRun = true;
    else if (a === '--limit') {
        opts.limit = parseInt(args[++i], 10);
        if (!Number.isFinite(opts.limit) || opts.limit <= 0) {
            console.error('Ungueltiger --limit Wert'); process.exit(1);
        }
    } else {
        console.error('Unbekannte Option: ' + a);
        process.exit(1);
    }
}

if (opts.help) {
    console.log(`
pptx-manager: Thumbnail-Batch-Generator

Nutzung:
  node bin/generate-thumbs.js              # alle fehlenden Thumbnails erzeugen
  node bin/generate-thumbs.js --limit 50   # nur 50 Folien bearbeiten
  node bin/generate-thumbs.js --force      # auch vorhandene neu erzeugen
  node bin/generate-thumbs.js --dry-run    # nur anzeigen, nichts erzeugen
  node bin/generate-thumbs.js --help       # diese Hilfe

Abbruch: Strg+C beendet sauber, beim naechsten Lauf geht es an der
Stelle weiter (Eintraege in der Tabelle 'thumbnails' werden uebersprungen).
`);
    process.exit(0);
}

// -------- DB oeffnen --------
const dbPath = path.join(__dirname, '..', 'data', 'pptx-manager.db');
if (!fs.existsSync(dbPath)) {
    console.error('Datenbank nicht gefunden: ' + dbPath);
    console.error('Zuerst einen Scan in der Weboberflaeche durchfuehren.');
    process.exit(1);
}
const db = initDb(dbPath);

// Thumbnail-Cache-Dir
const thumbCacheDir = path.join(__dirname, '..', 'data', 'thumbnails');
fs.mkdirSync(thumbCacheDir, { recursive: true });

// -------- Pro text_hash einen Repraesentanten holen --------
// Wir gruppieren slides nach text_hash und nehmen pro Gruppe die Folie mit
// kleinster slide_id als Quelle. So rendert das CLI jedes PNG nur einmal,
// auch wenn 42 Folien den gleichen Inhalt haben.
//
// Mit --force: ALLE text_hashes neu, auch wenn schon in thumbnails.
// Ohne:        nur die, deren text_hash noch nicht in thumbnails steht
//              ODER deren PNG-Datei auf der Platte fehlt.
const allRepresentativesSql = `
    WITH reps AS (
        SELECT
            s.text_hash,
            MIN(s.id) AS slide_id
        FROM slides s
        GROUP BY s.text_hash
    )
    SELECT
        reps.text_hash,
        reps.slide_id,
        s.slide_index,
        p.id AS presentation_id,
        p.file_name,
        p.file_path
    FROM reps
    JOIN slides s ON s.id = reps.slide_id
    JOIN presentations p ON p.id = s.presentation_id
`;

let slideRows;
if (opts.force) {
    slideRows = db.prepare(allRepresentativesSql + ' ORDER BY p.id, s.slide_index').all();
} else {
    // Links-join auf thumbnails; wir nehmen auch Eintraege mit fehlendem PNG mit
    slideRows = db.prepare(`
        ${allRepresentativesSql}
        LEFT JOIN thumbnails t ON t.text_hash = reps.text_hash
        ORDER BY p.id, s.slide_index
    `).all();
    // JS-Filter: wenn Eintrag existiert UND Datei vorhanden -> skippen
    const keep = [];
    const existingThumbs = new Map(db.prepare('SELECT text_hash, png_path FROM thumbnails').all()
        .map(r => [r.text_hash, r.png_path]));
    for (const r of slideRows) {
        const thumbPath = existingThumbs.get(r.text_hash);
        if (thumbPath && fs.existsSync(thumbPath)) continue; // schon da
        keep.push(r);
    }
    slideRows = keep;
}

if (opts.limit) {
    slideRows = slideRows.slice(0, opts.limit);
}

if (slideRows.length === 0) {
    console.log('Nichts zu tun: Alle Thumbnails sind aktuell.');
    process.exit(0);
}

// Nach Praesentation gruppieren (batch-effizient: PowerPoint oeffnet Datei nur einmal)
const byPresentation = new Map();
for (const row of slideRows) {
    if (!byPresentation.has(row.presentation_id)) {
        byPresentation.set(row.presentation_id, {
            fileName: row.file_name,
            filePath: row.file_path,
            slides: [],
        });
    }
    byPresentation.get(row.presentation_id).slides.push({
        slideId: row.slide_id,
        slideIndex: row.slide_index,
        textHash: row.text_hash,
    });
}

const totalFiles = byPresentation.size;
const totalSlides = slideRows.length;

// Hintergrund-Info: wie viele Folien-Duplikate liegen dahinter?
const totalSlideCount = db.prepare('SELECT COUNT(*) AS c FROM slides').get().c;
const uniqueHashes = db.prepare('SELECT COUNT(DISTINCT text_hash) AS c FROM slides').get().c;

console.log('='.repeat(60));
console.log(`pptx-manager - Thumbnail-Batch (Dedupe per text_hash)`);
console.log('='.repeat(60));
console.log(`DB gesamt: ${totalSlideCount} Folien in ${totalFiles > 0 ? '...' : '0'} Dateien -> ${uniqueHashes} eindeutige Inhalte`);
console.log(`Zu rendern: ${totalSlides} eindeutige Folien aus ${totalFiles} Praesentationen`);
console.log(`Ziel-Verzeichnis: ${thumbCacheDir}`);
console.log(`Aufloesung: ${THUMB_WIDTH} x ${THUMB_HEIGHT}`);
if (opts.dryRun) console.log('*** DRY RUN - es wird nichts erzeugt ***');
console.log('Abbrechen mit Strg+C (PowerPoint wird sauber beendet).');
console.log('='.repeat(60));
console.log('');

if (opts.dryRun) {
    let i = 0;
    for (const [presId, group] of byPresentation.entries()) {
        i++;
        console.log(`[${i}/${totalFiles}] ${group.fileName} — ${group.slides.length} Folien`);
        for (const s of group.slides.slice(0, 3)) {
            console.log(`    Folie ${s.slideIndex} -> ${thumbnailCachePath(thumbCacheDir, s.slideId)}`);
        }
        if (group.slides.length > 3) {
            console.log(`    ... und ${group.slides.length - 3} weitere`);
        }
    }
    process.exit(0);
}

// -------- SIGINT-Handling --------
let stopRequested = false;
process.on('SIGINT', () => {
    if (stopRequested) {
        console.log('\nZweites Strg+C — harter Abbruch.');
        process.exit(130);
    }
    stopRequested = true;
    console.log('\nStrg+C empfangen. Beende nach aktueller Datei... (nochmal fuer harten Abbruch)');
});

// -------- DB-Statement fuer Thumbnail-Eintraege --------
const insertThumb = db.prepare(`
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
`);

// -------- Hauptschleife --------
(async () => {
    const startTime = Date.now();
    let fileIndex = 0;
    let totalExported = 0;
    let totalFailed = 0;
    const missingFiles = [];

    for (const [presId, group] of byPresentation.entries()) {
        if (stopRequested) break;
        fileIndex++;

        const fileStart = Date.now();
        console.log(`[${fileIndex}/${totalFiles}] ${group.fileName}  (${group.slides.length} Folien)`);

        // Existiert die PPTX noch?
        if (!fs.existsSync(group.filePath)) {
            console.log(`  \u2716 Datei nicht gefunden: ${group.filePath}`);
            missingFiles.push(group.filePath);
            totalFailed += group.slides.length;
            continue;
        }

        // Jobs bauen — Cache-Pfad jetzt per text_hash (alle Duplikate teilen sich das PNG)
        const jobs = group.slides.map(s => ({
            slideIndex: s.slideIndex,
            outPath: thumbnailCachePath(thumbCacheDir, s.textHash),
            slideId: s.slideId,
            textHash: s.textHash,
        }));

        try {
            const result = await generateThumbnailsBatch(
                group.filePath,
                jobs,
                (done, total, slideIndex) => {
                    // Sparsamer Progress innerhalb der Datei
                    if (done === total || done % 10 === 0) {
                        process.stdout.write(`\r  ... ${done}/${total} Folien`);
                    }
                }
            );

            // Ergebnisse in die DB schreiben
            const now = Date.now();
            for (const exp of result.exported) {
                const job = jobs.find(j => j.slideIndex === exp.slideIndex);
                if (job) {
                    insertThumb.run(
                        job.textHash, exp.outPath, THUMB_WIDTH, THUMB_HEIGHT, now,
                        job.slideId, group.filePath, job.slideIndex
                    );
                }
            }

            const fileMs = Date.now() - fileStart;
            const msg = `\r  \u2713 ${result.exported.length}/${jobs.length} Folien in ${(fileMs/1000).toFixed(1)}s`;
            if (result.failed.length > 0) {
                process.stdout.write(msg + `  (${result.failed.length} Fehler)\n`);
                for (const f of result.failed.slice(0, 3)) {
                    console.log(`      Folie ${f.slideIndex}: ${f.error}`);
                }
                if (result.failed.length > 3) {
                    console.log(`      ... und ${result.failed.length - 3} weitere Fehler`);
                }
            } else {
                process.stdout.write(msg + '\n');
            }

            totalExported += result.exported.length;
            totalFailed += result.failed.length;
        } catch (err) {
            console.log(`\n  \u2716 Fehler: ${err.message}`);
            totalFailed += jobs.length;
        }
    }

    const totalMs = Date.now() - startTime;
    console.log('');
    console.log('='.repeat(60));
    console.log(`Fertig.`);
    console.log(`  Erzeugt:   ${totalExported}`);
    console.log(`  Fehler:    ${totalFailed}`);
    console.log(`  Dauer:     ${(totalMs/1000).toFixed(1)}s`);
    if (stopRequested) {
        console.log(`  (Abgebrochen nach Datei ${fileIndex}/${totalFiles} — einfach wieder starten, macht weiter.)`);
    }
    if (missingFiles.length > 0) {
        console.log(`  Fehlende Dateien: ${missingFiles.length}`);
        for (const f of missingFiles.slice(0, 5)) console.log(`    ${f}`);
        if (missingFiles.length > 5) console.log(`    ... und ${missingFiles.length - 5} weitere`);
    }
    console.log('='.repeat(60));

    process.exit(stopRequested ? 130 : 0);
})().catch(err => {
    console.error('Unerwarteter Fehler:', err);
    process.exit(1);
});
