// =============================================================
// fingerprint.js - PPTX-Dateien analysieren
// =============================================================
// Eine PPTX ist eine ZIP-Datei mit XML-Inhalt. Wir:
//   1. Entpacken sie mit jszip
//   2. Lesen jede Folien-XML (ppt/slides/slide*.xml)
//   3. Extrahieren Text + Shape-Struktur
//   4. Berechnen 3 Hashes pro Folie:
//      - exact_hash: SHA-256 des Roh-XMLs (byte-identisch)
//      - text_hash: SHA-256 des normalisierten Texts
//      - structure_hash: SHA-256 der Shape-Typen + grobe Positionen
// =============================================================

const fs = require('fs');
const crypto = require('crypto');
const JSZip = require('jszip');
const { XMLParser } = require('fast-xml-parser');

// XML-Parser: behaelt Attribute, alle Werte als String
const xmlParser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    parseAttributeValue: false,
    parseTagValue: false,
    textNodeName: '#text',
    alwaysCreateTextNode: true,
});

/**
 * SHA-256-Hash eines Strings oder Buffers.
 */
function sha256(data) {
    return crypto.createHash('sha256').update(data).digest('hex');
}

/**
 * Hash einer ganzen Datei (Stream-basiert, auch fuer grosse Dateien schnell).
 */
function hashFile(filePath) {
    return new Promise((resolve, reject) => {
        const hash = crypto.createHash('sha256');
        const stream = fs.createReadStream(filePath);
        stream.on('data', chunk => hash.update(chunk));
        stream.on('end', () => resolve(hash.digest('hex')));
        stream.on('error', reject);
    });
}

/**
 * Rekursiv alle Text-Knoten aus einem geparsten XML-Objekt sammeln.
 * In PPTX sind Texte in <a:t>-Elementen.
 */
function collectText(node, out = []) {
    if (node === null || node === undefined) return out;
    if (typeof node === 'string') {
        out.push(node);
        return out;
    }
    if (typeof node !== 'object') return out;

    if (Array.isArray(node)) {
        for (const item of node) collectText(item, out);
        return out;
    }

    // a:t = Text-Element (Attribute sind '@_' prefixed, Text in '#text')
    for (const [key, value] of Object.entries(node)) {
        if (key.startsWith('@_')) continue;
        if (key === '#text') {
            if (value !== undefined && value !== null) out.push(String(value));
            continue;
        }
        collectText(value, out);
    }
    return out;
}

/**
 * Titel einer Folie aus dem XML extrahieren.
 * Heuristik: Erster Text in einer Shape mit placeholder type="title" oder "ctrTitle".
 * Falls nichts gefunden, die ersten 80 Zeichen des Gesamt-Texts.
 */
function extractTitle(slideObj, fullText) {
    try {
        const spTree = slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'];
        if (!spTree) return fullText.substring(0, 80);

        const shapes = Array.isArray(spTree['p:sp']) ? spTree['p:sp'] : (spTree['p:sp'] ? [spTree['p:sp']] : []);
        for (const sp of shapes) {
            const phType = sp?.['p:nvSpPr']?.['p:nvPr']?.['p:ph']?.['@_type'];
            if (phType === 'title' || phType === 'ctrTitle') {
                const titleText = collectText(sp).join(' ').trim();
                if (titleText) return titleText.substring(0, 200);
            }
        }
    } catch (err) {
        // ignorieren
    }
    return fullText.substring(0, 80);
}

/**
 * Shape-Typen und grobe Positionen extrahieren fuer structure_hash.
 * Positionen werden auf 100.000 EMU-Raster gerundet (~0.26 cm Toleranz),
 * damit Mini-Verschiebungen keine neue Signatur erzeugen.
 */
function extractStructure(slideObj) {
    const EMU_GRID = 100000;
    const shapes = [];
    try {
        const spTree = slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'];
        if (!spTree) return '[]';

        const sps = Array.isArray(spTree['p:sp']) ? spTree['p:sp'] : (spTree['p:sp'] ? [spTree['p:sp']] : []);
        for (const sp of sps) {
            const xfrm = sp?.['p:spPr']?.['a:xfrm'];
            const off = xfrm?.['a:off'];
            const ext = xfrm?.['a:ext'];
            shapes.push({
                t: 'sp',
                x: off ? Math.round(parseInt(off['@_x'] || '0') / EMU_GRID) : 0,
                y: off ? Math.round(parseInt(off['@_y'] || '0') / EMU_GRID) : 0,
                cx: ext ? Math.round(parseInt(ext['@_cx'] || '0') / EMU_GRID) : 0,
                cy: ext ? Math.round(parseInt(ext['@_cy'] || '0') / EMU_GRID) : 0,
            });
        }

        const pics = Array.isArray(spTree['p:pic']) ? spTree['p:pic'] : (spTree['p:pic'] ? [spTree['p:pic']] : []);
        for (const pic of pics) {
            const xfrm = pic?.['p:spPr']?.['a:xfrm'];
            const off = xfrm?.['a:off'];
            const ext = xfrm?.['a:ext'];
            shapes.push({
                t: 'pic',
                x: off ? Math.round(parseInt(off['@_x'] || '0') / EMU_GRID) : 0,
                y: off ? Math.round(parseInt(off['@_y'] || '0') / EMU_GRID) : 0,
                cx: ext ? Math.round(parseInt(ext['@_cx'] || '0') / EMU_GRID) : 0,
                cy: ext ? Math.round(parseInt(ext['@_cy'] || '0') / EMU_GRID) : 0,
            });
        }
    } catch (err) {
        // ignorieren
    }
    // Nach Position sortieren, damit Reihenfolge in XML keine Rolle spielt
    shapes.sort((a, b) => a.y - b.y || a.x - b.x || a.t.localeCompare(b.t));
    return JSON.stringify(shapes);
}

/**
 * Text normalisieren fuer text_hash.
 * - lowercase
 * - Whitespace kollabieren
 * - Fuehrende/nachgestellte Leerzeichen entfernen
 */
function normalizeText(text) {
    return text
        .toLowerCase()
        .replace(/\s+/g, ' ')
        .trim();
}

/**
 * Eine PPTX-Datei komplett analysieren.
 * Liefert: { slides: [{ slideIndex, title, text, exactHash, textHash, structureHash }] }
 */
async function analyzePptx(filePath) {
    const buffer = await fs.promises.readFile(filePath);
    const zip = await JSZip.loadAsync(buffer);

    // Alle Folien-XMLs finden (ppt/slides/slide1.xml, slide2.xml, ...)
    const slideFiles = Object.keys(zip.files)
        .filter(name => /^ppt\/slides\/slide\d+\.xml$/.test(name))
        .sort((a, b) => {
            const na = parseInt(a.match(/slide(\d+)\.xml$/)[1]);
            const nb = parseInt(b.match(/slide(\d+)\.xml$/)[1]);
            return na - nb;
        });

    const slides = [];
    for (let i = 0; i < slideFiles.length; i++) {
        const xmlContent = await zip.files[slideFiles[i]].async('string');
        const slideObj = xmlParser.parse(xmlContent);

        // Text extrahieren
        const textParts = collectText(slideObj);
        const fullText = textParts.join(' ').replace(/\s+/g, ' ').trim();
        const title = extractTitle(slideObj, fullText);

        // 3 Hashes
        const exactHash = sha256(xmlContent);
        const textHash = sha256(normalizeText(fullText));
        const structureHash = sha256(extractStructure(slideObj));

        slides.push({
            slideIndex: i + 1,
            title: title,
            text: fullText,
            exactHash,
            textHash,
            structureHash,
        });
    }

    return { slides };
}

module.exports = {
    analyzePptx,
    hashFile,
    sha256,
};
