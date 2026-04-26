// =============================================================
// config.js - Zentrale Konfiguration
// =============================================================
// Laedt config.json aus dem Projekt-Root. Fehlende Werte werden
// aus den DEFAULTS aufgefuellt — die Config-Datei muss also nicht
// vollstaendig sein. Wenn config.json fehlt, gelten die Defaults.
// =============================================================

const fs = require('fs');
const path = require('path');

const DEFAULTS = {
    search: {
        pageSize: 30,           // Wie viele Gruppen pro Suchanfrage zurueck
        maxFtsRows: 5000,       // FTS5-Treffer-Limit, bevor gruppiert wird
        occurrencesShown: 30,   // Wie viele Vorkommen pro Gruppe in der UI
        minQueryLength: 2,      // Suche startet ab dieser Zeichenanzahl
    },
    lightbox: {
        heightVh: 85,           // Lightbox-Bildhoehe in % der Bildschirmhoehe
        maxWidthVw: 95,         // Maximale Lightbox-Breite in % der Bildschirmbreite
    },
};

let cached = null;

function deepMerge(a, b) {
    const out = { ...a };
    for (const k of Object.keys(b || {})) {
        const av = a[k];
        const bv = b[k];
        if (bv && typeof bv === 'object' && !Array.isArray(bv)) {
            out[k] = deepMerge(av && typeof av === 'object' ? av : {}, bv);
        } else {
            out[k] = bv;
        }
    }
    return out;
}

function loadConfig() {
    if (cached) return cached;
    const file = path.join(__dirname, '..', 'config.json');
    let userConfig = {};
    if (fs.existsSync(file)) {
        try {
            userConfig = JSON.parse(fs.readFileSync(file, 'utf8'));
        } catch (err) {
            console.warn(`[CONFIG] config.json ungueltig (${err.message}) — nutze Defaults.`);
        }
    } else {
        console.log('[CONFIG] config.json nicht vorhanden — nutze Defaults.');
    }
    cached = deepMerge(DEFAULTS, userConfig);
    return cached;
}

/**
 * Erzwingt Neu-Laden beim naechsten loadConfig() (z.B. fuer Tests).
 */
function resetCache() {
    cached = null;
}

module.exports = { loadConfig, resetCache, DEFAULTS };
