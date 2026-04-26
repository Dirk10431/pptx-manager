#!/usr/bin/env node
// =============================================================
// bin/analyze-pdfs.js — PDF-Erkennung "Praesentation vs. Dokument"
// =============================================================
// Phase 0 / Studie: Liest Metadaten + Seitenmasse + Textdichte aus
// allen PDFs in einem Ordner und schreibt eine CSV mit Score und
// Verdikt. KEINE DB-Aenderung, kein Eingriff ins Tool. Reine Analyse.
//
// Nutzung:
//   npm run analyze-pdfs -- "C:\\Pfad\\zu\\Ordner"
//   npm run analyze-pdfs -- "C:\\Pfad" --limit 50
//   npm run analyze-pdfs -- "C:\\Pfad" --output mein-bericht.csv
//
// Punkte-Modell (siehe README/STYLEGUIDE oder Kommentar unten):
//   +3  Producer ist Praesentations-Software
//   +2  Querformat (Breite > Hoehe)
//   +2  Aspect-Ratio in {4:3, 16:10, 16:9} (±5%)
//   +2  Wenig Text pro Seite (< 200 Woerter im Schnitt)
//   +1  Filename matcht Praesentations-Pattern
//   -2  Hochformat
//   -2  Viel Text pro Seite (> 500 Woerter)
//
//   verdikt = score >= 3 ? PRAESENTATION
//           : score <=  0 ? DOKUMENT
//           : UNKLAR
// =============================================================

const fs = require('fs');
const path = require('path');
const pdf = require('pdf-parse');

// -------- Argumente --------
const args = process.argv.slice(2);
let folder = null;
const opts = { limit: null, output: null, help: false };
for (let i = 0; i < args.length; i++) {
    const a = args[i];
    if (a === '--help' || a === '-h') opts.help = true;
    else if (a === '--limit')   opts.limit  = parseInt(args[++i], 10);
    else if (a === '--output')  opts.output = args[++i];
    else if (!a.startsWith('--')) folder = a;
    else { console.error('Unbekannte Option: ' + a); process.exit(1); }
}

if (opts.help || !folder) {
    console.log(`
PDF-Analyse: Erkennt Praesentations-PDFs vs. normale Dokumente

Nutzung:
  npm run analyze-pdfs -- "<ordner>"
  npm run analyze-pdfs -- "<ordner>" --limit 50
  npm run analyze-pdfs -- "<ordner>" --output mein-bericht.csv

Es wird eine CSV in data/ geschrieben (Zeitstempel im Dateinamen),
die du in Excel oder VS Code anschauen kannst. Sortiere nach
'verdikt' und 'score', um Treffer und Grenzfaelle zu finden.

Phase 0: KEINE Aenderung am Tool. Reine Auswertung.
`);
    process.exit(opts.help ? 0 : 1);
}

if (!fs.existsSync(folder)) {
    console.error('Ordner nicht gefunden: ' + folder);
    process.exit(1);
}

// -------- Walker --------
function findPdfs(dir, results = []) {
    let entries;
    try { entries = fs.readdirSync(dir, { withFileTypes: true }); }
    catch (e) { return results; }
    for (const e of entries) {
        const full = path.join(dir, e.name);
        if (e.isDirectory()) {
            if (e.name.startsWith('.')) continue;
            if (['node_modules', 'data', 'backups', '_old'].includes(e.name)) continue;
            findPdfs(full, results);
        } else if (e.isFile() && e.name.toLowerCase().endsWith('.pdf')) {
            results.push(full);
        }
    }
    return results;
}

// -------- Heuristik / Scoring --------
const PRESENTATION_PRODUCERS = [
    'powerpoint', 'keynote', 'impress', 'slidev', 'reveal',
    'beamer', 'gamma', 'marp', 'pitch',
];
const DOCUMENT_PRODUCERS = [
    'microsoft word', 'ms word', 'word', 'pages',
    'libreoffice writer', 'openoffice writer',
    'abiword',
];
const PRESENTATION_FILENAME_HINTS = [
    'folien', 'slides', 'praesentation', 'präsentation',
    'presentation', 'vortrag', 'seminar', 'schulung',
    'workshop', 'training', 'lecture', 'unterweisung',
];
// Standard-Praesentations-Seitenverhaeltnisse
const PRESENTATION_RATIOS = [4/3, 16/10, 16/9];
const RATIO_TOLERANCE = 0.05;

function looksLikePresentationRatio(ratio) {
    return PRESENTATION_RATIOS.some(r => Math.abs(ratio - r) < RATIO_TOLERANCE);
}

function scorePdf(info) {
    let score = 0;
    const reasons = [];

    const producer = (info.producer || '').toLowerCase();
    if (PRESENTATION_PRODUCERS.some(p => producer.includes(p))) {
        score += 3;
        reasons.push('producer-praes');
    } else if (DOCUMENT_PRODUCERS.some(p => producer.includes(p))) {
        score -= 2;
        reasons.push('producer-doc');
    }

    if (info.pages > 0 && info.firstPageWidth > 0 && info.firstPageHeight > 0) {
        if (info.firstPageLandscape) {
            score += 2;
            reasons.push('querformat');
        } else {
            score -= 2;
            reasons.push('hochformat');
        }

        if (info.firstPageAspect && looksLikePresentationRatio(info.firstPageAspect)) {
            score += 2;
            reasons.push('praes-ratio');
        }
    }

    if (info.avgWordsPerPage > 0) {
        if (info.avgWordsPerPage < 200) {
            score += 2;
            reasons.push('wenig-text');
        } else if (info.avgWordsPerPage > 500) {
            score -= 2;
            reasons.push('viel-text');
        }
    }

    const fname = (info.fileName || '').toLowerCase();
    if (PRESENTATION_FILENAME_HINTS.some(h => fname.includes(h))) {
        score += 1;
        reasons.push('filename');
    }

    let verdikt;
    if (score >= 3)      verdikt = 'PRAESENTATION';
    else if (score <= 0) verdikt = 'DOKUMENT';
    else                 verdikt = 'UNKLAR';

    return { score, verdikt, reasons };
}

// -------- PDF-Analyse --------
async function analyzePdf(filePath) {
    const buffer = fs.readFileSync(filePath);
    const stat = fs.statSync(filePath);

    // pagerender-Callback: pdfjs gibt uns pro Seite das page-Objekt.
    // Statt getViewport() (API-Drift zwischen pdfjs-Versionen) nutzen wir
    // pageData.view = [x1, y1, x2, y2] — direkt aus dem PDF-MediaBox.
    const pageInfos = [];
    function render_page(pageData) {
        try {
            const view = pageData.view || [0, 0, 0, 0];
            const width = view[2] - view[0];
            const height = view[3] - view[1];
            return pageData.getTextContent().then(tc => {
                const text = tc.items.map(it => it.str).join(' ');
                pageInfos.push({
                    width,
                    height,
                    wordCount: text.split(/\s+/).filter(Boolean).length,
                });
                return text;
            });
        } catch (e) {
            return Promise.resolve('');
        }
    }

    const result = await pdf(buffer, { pagerender: render_page });

    const totalWords = pageInfos.reduce((s, p) => s + p.wordCount, 0);
    const avgWordsPerPage = pageInfos.length > 0 ? Math.round(totalWords / pageInfos.length) : 0;

    const firstPage = pageInfos[0] || { width: 0, height: 0 };
    const firstPageAspect    = firstPage.height > 0 ? firstPage.width / firstPage.height : 0;
    const firstPageLandscape = firstPage.width  >  firstPage.height;

    const fileName = path.basename(filePath);
    const info = {
        filePath,
        fileName,
        fileSizeKb: Math.round(stat.size / 1024),
        pages: result.numpages || pageInfos.length,
        firstPageWidth:  Math.round(firstPage.width),
        firstPageHeight: Math.round(firstPage.height),
        firstPageAspect,
        firstPageLandscape,
        producer: (result.info && result.info.Producer) || '',
        creator:  (result.info && result.info.Creator)  || '',
        title:    (result.info && result.info.Title)    || '',
        avgWordsPerPage,
    };

    const scoring = scorePdf(info);
    return { ...info, ...scoring };
}

// -------- CSV-Helfer --------
function csvField(s) {
    if (s === null || s === undefined) return '';
    s = String(s);
    if (/[",\r\n]/.test(s)) {
        s = '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
}

// -------- Haupt-Lauf --------
(async () => {
    console.log('Suche PDFs in: ' + folder);
    let pdfs = findPdfs(folder);
    if (opts.limit) pdfs = pdfs.slice(0, opts.limit);
    console.log(`Gefunden: ${pdfs.length} PDF-Datei${pdfs.length === 1 ? '' : 'en'}`);
    if (pdfs.length === 0) { console.log('Nichts zu tun.'); process.exit(0); }

    // Output-Pfad
    let output = opts.output;
    if (!output) {
        const stamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
        output = path.join(__dirname, '..', 'data', `pdf-analysis-${stamp}.csv`);
    } else if (!path.isAbsolute(output)) {
        output = path.resolve(output);
    }
    fs.mkdirSync(path.dirname(output), { recursive: true });

    const headers = [
        'file_path', 'file_name', 'pages', 'file_size_kb',
        'first_page_width', 'first_page_height', 'first_page_aspect',
        'querformat', 'producer', 'creator', 'title',
        'avg_words_per_page', 'score', 'verdikt', 'reasons',
    ];
    fs.writeFileSync(output, headers.map(csvField).join(',') + '\n');

    const counts = { PRAESENTATION: 0, DOKUMENT: 0, UNKLAR: 0, FEHLER: 0 };
    const start = Date.now();

    // SIGINT sauber abfangen
    let stopRequested = false;
    process.on('SIGINT', () => {
        stopRequested = true;
        console.log('\nStrg+C empfangen. Beende nach aktueller Datei.');
    });

    for (let i = 0; i < pdfs.length; i++) {
        if (stopRequested) break;
        const filePath = pdfs[i];
        const short = path.basename(filePath).slice(0, 60).padEnd(60);
        process.stdout.write(`\r[${i + 1}/${pdfs.length}] ${short}`);

        try {
            const r = await analyzePdf(filePath);
            const row = [
                r.filePath, r.fileName, r.pages, r.fileSizeKb,
                r.firstPageWidth, r.firstPageHeight, r.firstPageAspect.toFixed(3),
                r.firstPageLandscape ? 'ja' : 'nein',
                r.producer, r.creator, r.title,
                r.avgWordsPerPage, r.score, r.verdikt, r.reasons.join('|'),
            ];
            fs.appendFileSync(output, row.map(csvField).join(',') + '\n');
            counts[r.verdikt] = (counts[r.verdikt] || 0) + 1;
        } catch (err) {
            const row = [filePath, path.basename(filePath), '', '', '', '', '', '', '', '', '', '', '', 'FEHLER', err.message.slice(0, 200)];
            fs.appendFileSync(output, row.map(csvField).join(',') + '\n');
            counts.FEHLER++;
        }
    }

    process.stdout.write('\n\n');
    const ms = Date.now() - start;
    console.log('=================================================');
    console.log(`Fertig in ${(ms / 1000).toFixed(1)}s.`);
    console.log(`  PRÄSENTATION:  ${counts.PRAESENTATION}`);
    console.log(`  DOKUMENT:      ${counts.DOKUMENT}`);
    console.log(`  UNKLAR:        ${counts.UNKLAR}`);
    if (counts.FEHLER) console.log(`  FEHLER:        ${counts.FEHLER}`);
    if (stopRequested)  console.log('  (Vorzeitig abgebrochen.)');
    console.log('');
    console.log('CSV-Bericht:');
    console.log('  ' + output);
    console.log('=================================================');
    process.exit(0);
})().catch(err => {
    console.error('Unerwarteter Fehler:', err);
    process.exit(1);
});
