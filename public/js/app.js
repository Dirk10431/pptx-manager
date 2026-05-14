// =============================================================
// app.js - Frontend-Logik pptx-manager
// =============================================================

// --- Config (wird beim DOMContentLoaded aus /api/config geholt) ---
// Defaults entsprechen den Server-Defaults; werden ueberschrieben, sobald
// /api/config antwortet.
const APP_CONFIG = {
    search: { pageSize: 30, occurrencesShown: 30, minQueryLength: 2 },
    lightbox: { heightVh: 85, maxWidthVw: 95 },
};

async function loadAppConfig() {
    try {
        const response = await fetch('/api/config');
        if (!response.ok) return;
        const cfg = await response.json();
        if (cfg.search)   Object.assign(APP_CONFIG.search, cfg.search);
        if (cfg.lightbox) Object.assign(APP_CONFIG.lightbox, cfg.lightbox);
        // CSS-Variablen fuer Lightbox setzen
        document.documentElement.style.setProperty('--lightbox-height-vh', APP_CONFIG.lightbox.heightVh + 'vh');
        document.documentElement.style.setProperty('--lightbox-max-width-vw', APP_CONFIG.lightbox.maxWidthVw + 'vw');

        // Jahres-Filter-Pills aus Config rendern (kommen vom Server, basierend
        // auf aktuellem Jahr + config.search.yearFilterYearsBack)
        renderYearButtons(cfg.search && cfg.search.yearFilterOptions);
    } catch (err) {
        console.warn('Config konnte nicht geladen werden, nutze Defaults:', err);
    }
}

function renderYearButtons(options) {
    const container = document.getElementById('year-buttons');
    if (!container) return;
    // Fallback, falls /api/config nicht antwortet: aktuelles Jahr und 3 davor
    if (!options || !Array.isArray(options) || options.length === 0) {
        const current = new Date().getFullYear();
        const cutoff = current - 3;
        options = [{ value: 'all', label: 'Alle' }, { value: 'le' + cutoff, label: '≤ ' + cutoff }];
        for (let y = cutoff + 1; y <= current; y++) {
            options.push({ value: String(y), label: String(y) });
        }
    }
    container.innerHTML = options.map((opt, i) => {
        const active = (i === 0) ? ' active' : '';
        return `<button class="btn-pill year-btn${active}" data-year="${escapeHtml(opt.value)}">${escapeHtml(opt.label)}</button>`;
    }).join('');

    // Click-Handler an die frisch erzeugten Buttons binden
    container.querySelectorAll('.year-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            container.querySelectorAll('.year-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            filterState.year = btn.dataset.year;
            triggerSearch();
        });
    });
}

// --- Lightbox: Thumbnail vergroessert anzeigen ---
// Bewusst hochskaliert (kein neuer Render); wird unscharf, reicht aber zur
// visuellen Identifikation. ESC oder Klick auf Hintergrund schliesst.
function openLightbox(imgSrc, caption) {
    const lb = document.getElementById('lightbox');
    if (!lb) return;
    lb.querySelector('.lightbox-img').src = imgSrc;
    lb.querySelector('.lightbox-caption').textContent = caption || '';
    lb.style.display = 'flex';
    document.body.style.overflow = 'hidden';
}
function closeLightbox() {
    const lb = document.getElementById('lightbox');
    if (!lb) return;
    lb.style.display = 'none';
    lb.querySelector('.lightbox-img').src = '';
    document.body.style.overflow = '';
}
document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') closeLightbox();
});
document.addEventListener('DOMContentLoaded', () => {
    const lb = document.getElementById('lightbox');
    if (!lb) return;
    lb.querySelector('.lightbox-backdrop').addEventListener('click', closeLightbox);
    lb.querySelector('.lightbox-close').addEventListener('click', closeLightbox);
});

// --- Statistik laden + Dashboard-Tacho rendern ---
// Helfer: Zahl mit Tausender-Trennzeichen (deutsch).
function fmtNum(n) {
    if (n === null || n === undefined || !Number.isFinite(n)) return '–';
    return n.toLocaleString('de-DE');
}
function fmtPct(p) {
    if (p === null || p === undefined || !Number.isFinite(p)) return '–';
    return (p * 100).toFixed(1).replace('.', ',');
}

// Tacho-Donut: Fortschritts-Ring per stroke-dasharray animieren.
// Perimeter eines Kreises mit r=50 ist 2*PI*50 ≈ 314.16.
function setTachoCoverage(coverage) {
    const ring = document.getElementById('tacho-ring');
    if (!ring) return;
    const perimeter = 2 * Math.PI * 50;
    const filled = Math.max(0, Math.min(1, coverage)) * perimeter;
    ring.setAttribute('stroke-dasharray', `${filled} ${perimeter}`);
}

function setText(id, value) {
    const el = document.getElementById(id);
    if (el) el.textContent = value;
}

async function loadStats() {
    try {
        const response = await fetch('/api/stats');
        const data = await response.json();

        // Detail-Stats (Meta-Grid)
        setText('stat-presentations', fmtNum(data.presentations));
        setText('stat-slides',        fmtNum(data.slides));
        setText('stat-unique',        fmtNum(data.uniqueSlides));
        setText('stat-duplicates',    fmtNum(data.duplicateSlides));
        setText('stat-dup-factor',    Number.isFinite(data.duplicateFactor) && data.duplicateFactor > 0
            ? data.duplicateFactor.toFixed(1).replace('.', ',') + '×'
            : '–');
        setText('stat-thumb-missing', fmtNum(data.thumbnailsMissing));

        // Tacho (Thumbnail-Coverage)
        setTachoCoverage(data.thumbnailCoverage || 0);
        setText('tacho-pct-val',  fmtPct(data.thumbnailCoverage));
        setText('tacho-done',     fmtNum(data.thumbnailsDone));
        setText('tacho-total',    fmtNum(data.uniqueSlides));

        const note = document.getElementById('tacho-missing-note');
        if (note) {
            if (data.thumbnailsMissing > 0) {
                note.innerHTML = `<strong>${fmtNum(data.thumbnailsMissing)}</strong> fehlen — auf der Scan-Seite findest du den Befehl, um sie zu rendern.`;
            } else if (data.uniqueSlides > 0) {
                note.textContent = 'Alle Vorschauen sind aktuell.';
            } else {
                note.textContent = 'Noch nichts indiziert. Auf der Scan-Seite einen Ordner hinzufügen.';
            }
        }

        // Versionshinweis ggf. anzeigen
        const warning = document.getElementById('version-warning');
        if (warning) {
            warning.style.display = data.compatible ? 'none' : 'block';
        }
    } catch (err) {
        console.error('Fehler beim Laden der Stats:', err);
    }
}

// --- Index komplett zuruecksetzen ---
// Nutzt das Modal-Confirm, falls verfuegbar (Scan-Seite hat es); fallback auf
// natives confirm() auf Seiten ohne Modal. Nach erfolgreichem Reset wird die
// Seite neu geladen, damit keine alten Listen / Filter-States haengenbleiben.
async function resetIndex() {
    const useModal = typeof showConfirm === 'function' && document.getElementById('confirm-modal');
    let ok;
    if (useModal) {
        ok = await showConfirm({
            title: 'Index komplett zurücksetzen',
            body: 'Alle gescannten Folien-Daten, Hauptpfade und Thumbnails werden gelöscht. Die PPTX-Dateien auf der Platte bleiben unberührt.<br><br><span style="color: var(--muted); font-size: 0.85rem;">Hinweis: Die Seite wird nach dem Reset automatisch neu geladen.</span>',
            okText: 'Endgültig zurücksetzen',
            danger: true,
        });
    } else {
        ok = window.confirm('Alle gescannten Folien-Daten wirklich loeschen? Die PPTX-Dateien bleiben unberuehrt.\n\nDie Seite wird nach dem Reset automatisch neu geladen.');
    }
    if (!ok) return;

    try {
        const response = await fetch('/api/reset', { method: 'POST' });
        if (!response.ok) {
            const err = await response.json();
            alert('Fehler: ' + (err.error || 'Unbekannter Fehler'));
            return;
        }
        // Sauberer Re-Init via Full-Reload statt manuellem DOM-Aufraeumen
        window.location.reload();
    } catch (err) {
        alert('Netzwerkfehler: ' + err.message);
    }
}

// --- Pfad-Eingabe im localStorage merken ---
const PATH_KEY = 'pptx-manager.lastFolder';

function saveLastPath(path) {
    try { localStorage.setItem(PATH_KEY, path); } catch (e) {}
}
function loadLastPath() {
    try { return localStorage.getItem(PATH_KEY) || ''; } catch (e) { return ''; }
}

// --- Letztes Path-Segment aus einem Pfad extrahieren (Default-Label) ---
function basenameOfPath(p) {
    if (!p) return '';
    return p.replace(/[\\/]+$/, '').split(/[\\/]+/).pop() || p;
}

// Label-Feld auto-fuellen, solange der Nutzer es nicht selbst geaendert hat.
// Ein einmal manuell befuelltes Feld lassen wir in Ruhe.
function autoFillLabel(folderPath) {
    const labelInput = document.getElementById('scan-label');
    if (!labelInput) return;
    if (labelInput.dataset.userEdited === '1') return;
    labelInput.value = basenameOfPath(folderPath);
}

// --- Ordner ueber nativen Windows-Dialog waehlen ---
async function pickFolder() {
    const btn = document.getElementById('btn-pick-folder');
    const input = document.getElementById('folder-path');
    const originalText = btn.textContent;
    btn.disabled = true;
    btn.textContent = 'Dialog geoeffnet...';
    try {
        const response = await fetch('/api/pick-folder');
        const data = await response.json();
        if (data.error) {
            alert('Fehler: ' + data.error);
            return;
        }
        if (data.path) {
            input.value = data.path;
            autoFillLabel(data.path);
            saveLastPath(data.path);
            checkPath();
        }
    } catch (err) {
        alert('Netzwerkfehler: ' + err.message);
    } finally {
        btn.disabled = false;
        btn.textContent = originalText;
    }
}

// --- Pfad pruefen ---
async function checkPath() {
    const folderPath = document.getElementById('folder-path').value.trim();
    const resultEl = document.getElementById('check-result');
    const scanBtn = document.getElementById('btn-scan');

    if (!folderPath) {
        resultEl.innerHTML = '<p style="color: var(--warn);">Bitte einen Pfad eingeben.</p>';
        scanBtn.disabled = true;
        return;
    }

    resultEl.innerHTML = '<p>Pruefe Pfad...</p>';
    scanBtn.disabled = true;

    try {
        const response = await fetch('/api/check-path', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ folderPath }),
        });
        const data = await response.json();

        if (!data.valid) {
            resultEl.innerHTML = `<p style="color: var(--warn);">Fehler: ${data.error}</p>`;
            scanBtn.disabled = true;
            return;
        }

        if (data.fileCount === 0) {
            resultEl.innerHTML = `<p style="color: var(--accent2);">Keine PPTX-Dateien in diesem Ordner gefunden.</p>`;
            scanBtn.disabled = true;
            return;
        }

        let html = `<p><span class="badge badge-success">${data.fileCount} PPTX-Dateien</span> gefunden.</p>`;
        if (data.sample && data.sample.length > 0) {
            html += `<p style="color: var(--muted); font-size: 0.9rem;">Beispiele: ${data.sample.join(', ')}${data.fileCount > data.sample.length ? ', ...' : ''}</p>`;
        }
        resultEl.innerHTML = html;
        scanBtn.disabled = false;
        saveLastPath(folderPath);
    } catch (err) {
        resultEl.innerHTML = `<p style="color: var(--warn);">Netzwerkfehler: ${err.message}</p>`;
        scanBtn.disabled = true;
    }
}

// --- Scan starten ---
let pollInterval = null;

async function startScan() {
    const folderPath = document.getElementById('folder-path').value.trim();
    if (!folderPath) return;
    const labelInput = document.getElementById('scan-label');
    const label = labelInput ? labelInput.value.trim() : '';

    document.getElementById('btn-scan').disabled = true;
    document.getElementById('btn-check').disabled = true;
    document.getElementById('progress-card').style.display = 'block';

    try {
        const response = await fetch('/api/scan', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ folderPath, label }),
        });
        if (!response.ok) {
            const err = await response.json();
            alert('Fehler: ' + (err.error || 'Unbekannter Fehler'));
            return;
        }
        // Polling starten
        pollInterval = setInterval(updateProgress, 500);
    } catch (err) {
        alert('Netzwerkfehler: ' + err.message);
        document.getElementById('btn-scan').disabled = false;
        document.getElementById('btn-check').disabled = false;
    }
}

async function updateProgress() {
    try {
        const response = await fetch('/api/scan/status');
        const state = await response.json();

        document.getElementById('progress-current').textContent = state.currentFile || '-';
        document.getElementById('progress-text').textContent = `${state.processed} / ${state.total}`;
        document.getElementById('progress-added').textContent = state.added;
        document.getElementById('progress-updated').textContent = state.updated;
        document.getElementById('progress-skipped').textContent = state.skipped;
        document.getElementById('progress-failed').textContent = state.failed;

        const pct = state.total > 0 ? Math.round((state.processed / state.total) * 100) : 0;
        document.getElementById('progress-fill').style.width = pct + '%';

        if (!state.running && state.finishedAt) {
            clearInterval(pollInterval);
            pollInterval = null;
            const progHdr = document.querySelector('#progress-card h2, #progress-card .section-label');
            if (progHdr) progHdr.textContent = 'Scan abgeschlossen';
            const btnScan = document.getElementById('btn-scan');
            const btnCheck = document.getElementById('btn-check');
            if (btnScan)  btnScan.disabled = false;
            if (btnCheck) btnCheck.disabled = false;
            document.getElementById('progress-current').textContent = '-';
            loadStats();
            // Index-Verwaltung neu rendern (Counts, last-scanned-at)
            loadAndRenderManageRoots();
            // Wenn der Scanner stale files erkannt hat: Bestaetigungs-Modal zeigen.
            if (state.scanRootId && Array.isArray(state.staleFiles) && state.staleFiles.length > 0) {
                showStaleModal(state.scanRootId, state.label || '', state.staleFiles);
            }
        }
    } catch (err) {
        console.error('Polling-Fehler:', err);
    }
}

// --- Volltextsuche (debounced) ---
let searchTimer = null;

function escapeHtml(str) {
    return String(str || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

// Aktueller Suchzustand fuer "Mehr anzeigen"
let searchState = { term: '', offset: 0, limit: 30, accumulated: [] };

// Aktuelle Filter-Einstellungen (werden beim Wechsel gesetzt; loesen Re-Search aus)
const filterState = {
    year: 'all',          // 'all' | 'le2023' | '2024' | '2025' | '2026'
    sort: 'relevance',    // 'relevance' | 'date_desc' | 'date_asc'
    filename: '',         // Substring im Dateinamen
    activeRoots: null,    // Set<fullPath> mit aktivierten Roots; null = alle
};

let allRoots = []; // Aus /api/scan-roots geladene Hauptpfade

// --- Hauptpfade laden + als Checkbox-Pills rendern ---
async function loadAndRenderRoots() {
    try {
        const response = await fetch('/api/scan-roots');
        const data = await response.json();
        allRoots = data.roots || [];

        const row = document.getElementById('filter-roots-row');
        const list = document.getElementById('filter-roots');
        if (!row || !list) return;

        if (allRoots.length === 0) {
            row.style.display = 'none';
            return;
        }

        // Standard: alle aktiv
        filterState.activeRoots = new Set(allRoots.map(r => r.fullPath));

        list.innerHTML = allRoots.map(r => `
            <label class="root-checkbox checked" data-root="${escapeHtml(r.fullPath)}" title="${escapeHtml(r.fullPath)} (${r.count} Dateien)">
                <input type="checkbox" checked>
                <span>${escapeHtml(r.label)}</span>
                <span style="color: var(--muted); font-size: 0.75rem;">(${r.count})</span>
            </label>
        `).join('');

        row.style.display = 'flex';

        // Klick-Handler je Pill
        list.querySelectorAll('.root-checkbox').forEach(el => {
            el.addEventListener('click', (e) => {
                if (e.target.tagName === 'INPUT') return; // doppelt vermeiden
                const cb = el.querySelector('input');
                cb.checked = !cb.checked;
                el.classList.toggle('checked', cb.checked);
                const root = el.dataset.root;
                if (cb.checked) filterState.activeRoots.add(root);
                else filterState.activeRoots.delete(root);
                triggerSearch();
            });
        });

        // "Alle abwaehlen / aktivieren"
        const toggleBtn = document.getElementById('btn-roots-toggle');
        if (toggleBtn) {
            toggleBtn.addEventListener('click', () => {
                const allActive = filterState.activeRoots.size === allRoots.length;
                if (allActive) {
                    filterState.activeRoots.clear();
                } else {
                    filterState.activeRoots = new Set(allRoots.map(r => r.fullPath));
                }
                // UI synchronisieren
                list.querySelectorAll('.root-checkbox').forEach(el => {
                    const cb = el.querySelector('input');
                    cb.checked = filterState.activeRoots.has(el.dataset.root);
                    el.classList.toggle('checked', cb.checked);
                });
                toggleBtn.textContent = allActive ? 'Alle aktivieren' : 'Alle abwaehlen';
                triggerSearch();
            });
        }
    } catch (err) {
        console.warn('Hauptpfade konnten nicht geladen werden:', err);
    }
}

// Wird von Filter-Aenderungen aufgerufen
function triggerSearch() {
    const term = document.getElementById('search-input').value.trim();
    runSearch(term);
}

async function runSearch(term, append = false) {
    const resultsEl = document.getElementById('search-results');
    const summaryEl = document.getElementById('search-summary');
    const pageSize = APP_CONFIG.search.pageSize;
    const minLen   = APP_CONFIG.search.minQueryLength;

    if (!append) {
        if (!term || term.length < minLen) {
            resultsEl.innerHTML = '';
            summaryEl.textContent = '';
            searchState = { term: '', offset: 0, limit: pageSize, accumulated: [] };
            return;
        }
        resultsEl.innerHTML = '<p style="color: var(--muted);">Suche...</p>';
        summaryEl.textContent = '';
        searchState = { term, offset: 0, limit: pageSize, accumulated: [] };
    }

    const uniqueOnly = document.getElementById('filter-unique-only').checked;

    // Aktive Hauptpfade als Pipe-getrennte Liste (kommt nur mit, wenn nicht alle aktiv)
    let rootsParam = '';
    if (filterState.activeRoots && allRoots.length > 0) {
        const active = Array.from(filterState.activeRoots);
        if (active.length === 0) {
            // Nichts aktiv -> sofort leeres Ergebnis, kein Server-Roundtrip noetig
            const resultsEl2 = document.getElementById('search-results');
            const summaryEl2 = document.getElementById('search-summary');
            resultsEl2.innerHTML = '<p style="color: var(--muted);">Kein Hauptpfad ausgewaehlt — Filter zu eng.</p>';
            summaryEl2.textContent = '';
            return;
        }
        if (active.length < allRoots.length) {
            rootsParam = '&roots=' + encodeURIComponent(active.join('|'));
        }
    }

    const url = '/api/search?q=' + encodeURIComponent(term)
              + (uniqueOnly ? '&uniqueOnly=1' : '')
              + (filterState.year && filterState.year !== 'all' ? '&year=' + encodeURIComponent(filterState.year) : '')
              + (filterState.sort && filterState.sort !== 'relevance' ? '&sort=' + encodeURIComponent(filterState.sort) : '')
              + (filterState.filename ? '&filename=' + encodeURIComponent(filterState.filename) : '')
              + rootsParam
              + `&offset=${searchState.offset}&limit=${searchState.limit}`;

    try {
        const response = await fetch(url);
        const data = await response.json();

        if (data.error) {
            resultsEl.innerHTML = `<p style="color: var(--warn);">${escapeHtml(data.error)}</p>`;
            return;
        }

        if (!append && (!data.groups || data.groups.length === 0)) {
            resultsEl.innerHTML = '<p style="color: var(--muted);">Keine Treffer.</p>';
            summaryEl.textContent = uniqueOnly
                ? `Gesamt ${data.totalGroupsBeforeFilter} Gruppen, davon 0 eindeutige.`
                : '';
            return;
        }

        searchState.accumulated.push(...data.groups);

        // Summary-Zeile
        summaryEl.textContent = uniqueOnly
            ? `${searchState.accumulated.length} von ${data.totalGroups} eindeutigen Folien angezeigt (${data.totalOccurrences} Treffer gesamt).`
            : `${searchState.accumulated.length} von ${data.totalGroups} Gruppen angezeigt (${data.totalOccurrences} Treffer gesamt).`;

        // Gruppen rendern (kompletter Re-Render bei accumulated)
        const groupsHtml = searchState.accumulated.map((g, idx) => renderGroup(g, idx)).join('');
        const moreBtnHtml = data.hasMore
            ? `<div style="text-align:center; margin: 1rem 0;"><button class="act-btn" id="btn-load-more">Naechste ${searchState.limit} laden (${data.totalGroups - searchState.accumulated.length} verbleibend)</button></div>`
            : '';
        resultsEl.innerHTML = `<div class="search-groups">${groupsHtml}</div>${moreBtnHtml}`;

        // "Mehr"-Button verdrahten
        const moreBtn = document.getElementById('btn-load-more');
        if (moreBtn) {
            moreBtn.addEventListener('click', () => {
                searchState.offset += searchState.limit;
                runSearch(searchState.term, true);
            });
        }

        // Klick-Handler fuer "Alle X Vorkommen zeigen" anhaengen
        resultsEl.querySelectorAll('.group-toggle').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const idx = e.currentTarget.dataset.groupIdx;
                const details = resultsEl.querySelector('.group-details[data-group-idx="' + idx + '"]');
                if (details) {
                    const open = details.style.display !== 'none';
                    details.style.display = open ? 'none' : 'block';
                    e.currentTarget.textContent = open ? 'Vorkommen zeigen' : 'Vorkommen ausblenden';
                }
            });
        });

        // Klick auf Thumbnail -> Lightbox (Event-Delegation)
        resultsEl.querySelectorAll('.thumb-wrap[data-slide-id]').forEach(wrap => {
            wrap.addEventListener('click', () => {
                const slideId = wrap.dataset.slideId;
                const img = wrap.querySelector('img');
                if (!img || img.style.display === 'none') return;
                const caption = wrap.dataset.caption || '';
                openLightbox(`/api/thumb/${slideId}`, caption);
            });
        });

        // Klick auf Aktions-Buttons (Ordner / Datei oeffnen)
        resultsEl.querySelectorAll('.btn-action').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const path = btn.dataset.path;
                const mode = btn.dataset.action; // 'folder' | 'file'
                openFileOrFolder(path, mode);
            });
        });
    } catch (err) {
        resultsEl.innerHTML = `<p style="color: var(--warn);">Netzwerkfehler: ${escapeHtml(err.message)}</p>`;
    }
}

// --- Aktions-Buttons: Datei oeffnen / im Explorer zeigen ---
function actionButtonsHtml(filePath) {
    const enc = escapeHtml(filePath);
    return `
        <span class="action-buttons" style="white-space: nowrap;">
            <button type="button" class="btn-action" data-action="folder" data-path="${enc}"
                    title="Im Explorer anzeigen">📂</button>
            <button type="button" class="btn-action" data-action="file" data-path="${enc}"
                    title="In PowerPoint oeffnen">▶</button>
        </span>
    `;
}

async function openFileOrFolder(filePath, mode) {
    try {
        const r = await fetch('/api/open', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ filePath, mode }),
        });
        if (!r.ok) {
            const err = await r.json();
            alert('Fehler: ' + (err.error || 'Unbekannt'));
        }
    } catch (e) {
        alert('Netzwerkfehler: ' + e.message);
    }
}

// --- Thumbnail-HTML fuer eine Folie (lazy, loading=lazy laesst Browser bestimmen wann) ---
// Wir geben ein festes Seitenverhaeltnis 16:9 vor, damit beim Laden nichts springt.
// onerror blendet Bild aus, wenn PowerPoint fehlt oder Export fehlgeschlagen ist.
function thumbnailHtml(slideId, opts = {}) {
    const width = opts.width || 200;
    const height = Math.round(width * 9 / 16);
    const caption = opts.caption ? ` data-caption="${escapeHtml(opts.caption)}"` : '';
    return `
        <div class="thumb-wrap" data-slide-id="${slideId}"${caption}
             style="width: ${width}px; height: ${height}px; flex-shrink: 0; background: #f5f5f5; border: 1px solid var(--border); border-radius: 4px; overflow: hidden; display: flex; align-items: center; justify-content: center;"
             title="Klicken fuer grosse Ansicht">
            <img
                src="/api/thumb/${slideId}"
                alt="Folien-Vorschau"
                loading="lazy"
                style="width: 100%; height: 100%; object-fit: contain;"
                onerror="this.style.display='none'; this.parentElement.style.cursor='default'; this.parentElement.innerHTML='<span style=\\'color:var(--muted);font-size:0.75rem;text-align:center;padding:0.25rem;\\'>Vorschau nicht verfuegbar</span>';"
            >
        </div>
    `;
}

// --- Eine Gruppe (identische Folien per text_hash) als HTML rendern ---
function renderGroup(g, idx) {
    const rep = g.representative;
    const isUnique = g.count === 1;
    const badgeClass = isUnique ? 'badge-success' : 'badge';
    const badgeText = isUnique
        ? 'Einzigartig'
        : `${g.count}&times; in ${g.fileCount} ${g.fileCount === 1 ? 'Datei' : 'Dateien'}`;

    // Vorkommen begrenzen: bei sehr grossen Gruppen (z.B. 8997 Kopien
    // einer Boilerplate-Folie) wuerde die Tabelle den Browser einfrieren.
    // Anzahl kommt aus config.json (search.occurrencesShown).
    const OCCURRENCES_INITIAL = APP_CONFIG.search.occurrencesShown;
    const visibleOcc = g.occurrences.slice(0, OCCURRENCES_INITIAL);
    const hiddenOccCount = g.occurrences.length - visibleOcc.length;

    const occurrencesHtml = visibleOcc.map(o => `
        <tr>
            <td style="width: 130px;">${thumbnailHtml(o.slideId, { width: 120, caption: `Folie ${o.slideIndex} — ${o.fileName}` })}</td>
            <td style="vertical-align: top;"><span class="badge">Folie ${o.slideIndex}</span></td>
            <td style="font-size: 0.85rem; vertical-align: top;" title="${escapeHtml(o.filePath)}">
                ${escapeHtml(o.fileName)}
                <div style="margin-top: 0.25rem;">${actionButtonsHtml(o.filePath)}</div>
            </td>
        </tr>
    `).join('');

    const occurrencesFooter = hiddenOccCount > 0
        ? `<tr><td colspan="3" style="text-align:center; padding:0.75rem; color:var(--muted); font-style: italic;">+ ${hiddenOccCount} weitere Vorkommen ausgeblendet (Performance-Schutz). Verfeinere die Suche, um sie zu sehen.</td></tr>`
        : '';

    return `
    <div class="group-card" style="border: 1px solid var(--border); border-radius: 6px; padding: 0.85rem 1rem; margin-bottom: 0.75rem;">
        <div style="display: flex; gap: 1rem; align-items: flex-start; flex-wrap: wrap;">
            ${thumbnailHtml(rep.slideId, { width: 200, caption: `Folie ${rep.slideIndex} — ${rep.fileName}` })}
            <div style="flex: 1; min-width: 260px;">
                <div style="margin-bottom: 0.25rem;">
                    <span class="badge ${badgeClass}">${badgeText}</span>
                    <span class="badge">Folie ${rep.slideIndex}</span>
                </div>
                <strong>${escapeHtml(rep.title || '(ohne Titel)')}</strong>
                <div style="font-size: 0.9rem; color: var(--muted); margin-top: 0.25rem;">${rep.snippet || ''}</div>
                <div style="font-size: 0.8rem; color: var(--muted); margin-top: 0.35rem;" title="${escapeHtml(rep.filePath)}">
                    Beispiel aus: ${escapeHtml(rep.fileName)}
                    ${actionButtonsHtml(rep.filePath)}
                </div>
            </div>
            ${!isUnique ? `
                <button class="act-btn group-toggle" data-group-idx="${idx}" style="flex-shrink: 0;">
                    Vorkommen zeigen
                </button>
            ` : ''}
        </div>
        ${!isUnique ? `
            <div class="group-details" data-group-idx="${idx}" style="display: none; margin-top: 0.75rem;">
                <table class="table" style="margin-top: 0.5rem;">
                    <thead>
                        <tr><th style="width: 130px;">Vorschau</th><th style="width: 7rem;">Folie</th><th>Datei</th></tr>
                    </thead>
                    <tbody>${occurrencesHtml}${occurrencesFooter}</tbody>
                </table>
            </div>
        ` : ''}
    </div>`;
}

// --- Thumb-Befehl rendern (permanenter Block zwischen Scan und Fortschritt) ---
// Holt den projectRoot vom Server und baut einen kopierbaren Code-Schnipsel.
async function renderThumbsCmd() {
    const box = document.getElementById('thumbs-cmd-box');
    if (!box) return;
    let projectRoot = '<projekt-pfad>';
    try {
        const r = await fetch('/api/project-info');
        if (r.ok) {
            const d = await r.json();
            projectRoot = d.projectRoot || projectRoot;
        }
    } catch {}
    // Git-Bash erwartet POSIX-Slashes; Win-Pfad mit Backslashes funktioniert
    // dort auch, aber Forward-Slashes sind freundlicher.
    const codeOneLine = `cd "${projectRoot.replace(/\\/g, '/')}" && npm run thumbs`;
    // Kompakte Variante: einzeilig, deshalb HTML auch ohne Einrueckung
    // schreiben (sonst expandiert white-space:pre-wrap die Leerzeichen).
    box.innerHTML =
        `<div class="prompt-box prompt-box-compact">` +
            `<button type="button" class="copy-btn" data-copy-target="thumbs-cmd-text">📋 Kopieren</button>` +
            `<code class="prompt-content" id="thumbs-cmd-text">${escapeHtml(codeOneLine)}</code>` +
        `</div>`;
    // Copy-Button verdrahten
    const cbtn = box.querySelector('.copy-btn');
    if (cbtn) {
        cbtn.addEventListener('click', () => {
            const target = document.getElementById(cbtn.dataset.copyTarget);
            if (!target) return;
            navigator.clipboard.writeText(target.textContent).then(() => {
                const orig = cbtn.innerHTML;
                cbtn.innerHTML = '✓ Kopiert';
                cbtn.style.color = 'var(--success)';
                setTimeout(() => { cbtn.innerHTML = orig; cbtn.style.color = ''; }, 1800);
            }).catch(() => {
                cbtn.innerHTML = '✗ Fehler';
                setTimeout(() => { cbtn.innerHTML = '📋 Kopieren'; }, 1800);
            });
        });
    }
}

// Pfad-Check-Hook: nach erfolgreichem Check Label auto-fuellen.
const __originalCheckPath = checkPath;
checkPath = async function patchedCheckPath() {
    await __originalCheckPath();
    const scanBtn = document.getElementById('btn-scan');
    if (scanBtn && !scanBtn.disabled) {
        const folderPath = document.getElementById('folder-path').value.trim();
        autoFillLabel(folderPath);
    }
};

// =============================================================
// Index-Verwaltung (scan-roots)
// =============================================================

// Generisches Bestaetigungs-Modal. Promise resolved zu true/false.
function showConfirm({ title = 'Bestätigen', body = 'Wirklich?', okText = 'OK', cancelText = 'Abbrechen', danger = false } = {}) {
    return new Promise((resolve) => {
        const modal = document.getElementById('confirm-modal');
        if (!modal) { resolve(window.confirm(body)); return; }
        document.getElementById('confirm-modal-title').textContent = title;
        document.getElementById('confirm-modal-body').innerHTML = body;
        const okBtn = document.getElementById('confirm-ok');
        const cancelBtn = document.getElementById('confirm-cancel');
        okBtn.textContent = okText;
        cancelBtn.textContent = cancelText;
        okBtn.classList.toggle('danger', !!danger);

        const close = (result) => {
            modal.style.display = 'none';
            okBtn.removeEventListener('click', onOk);
            cancelBtn.removeEventListener('click', onCancel);
            modal.querySelector('.modal-backdrop').removeEventListener('click', onCancel);
            modal.querySelector('.modal-close').removeEventListener('click', onCancel);
            resolve(result);
        };
        const onOk     = () => close(true);
        const onCancel = () => close(false);
        okBtn.addEventListener('click', onOk);
        cancelBtn.addEventListener('click', onCancel);
        modal.querySelector('.modal-backdrop').addEventListener('click', onCancel);
        modal.querySelector('.modal-close').addEventListener('click', onCancel);
        modal.style.display = 'flex';
    });
}

// Stale-Files-Modal nach Scan oder per Rescan-Button.
function showStaleModal(scanRootId, rootLabel, staleFiles) {
    const modal = document.getElementById('stale-modal');
    if (!modal) return;
    document.getElementById('stale-modal-title').textContent =
        `${staleFiles.length} Datei(en) nicht mehr gefunden — "${rootLabel}"`;
    const ul = document.getElementById('stale-modal-list');
    ul.innerHTML = staleFiles.map(f =>
        `<li><strong>${escapeHtml(f.fileName)}</strong><br><span style="color:var(--muted);">${escapeHtml(f.filePath)}</span></li>`
    ).join('');

    const confirmBtn = document.getElementById('stale-confirm');
    const cancelBtn  = document.getElementById('stale-cancel');
    const closeBtn   = modal.querySelector('.modal-close');
    const backdrop   = modal.querySelector('.modal-backdrop');

    const close = () => {
        modal.style.display = 'none';
        confirmBtn.removeEventListener('click', onConfirm);
        cancelBtn.removeEventListener('click', close);
        closeBtn.removeEventListener('click', close);
        backdrop.removeEventListener('click', close);
        confirmBtn.disabled = false;
        confirmBtn.textContent = 'Aus DB entfernen';
    };
    const onConfirm = async () => {
        confirmBtn.disabled = true;
        confirmBtn.textContent = 'Lösche...';
        try {
            const ids = staleFiles.map(f => f.id);
            const r = await fetch(`/api/scan-roots/${scanRootId}/cleanup-stale`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ ids }),
            });
            const data = await r.json();
            if (!r.ok) {
                alert('Fehler: ' + (data.error || 'unbekannt'));
            }
            loadStats();
            loadAndRenderManageRoots();
        } catch (err) {
            alert('Netzwerkfehler: ' + err.message);
        } finally {
            close();
        }
    };

    confirmBtn.addEventListener('click', onConfirm);
    cancelBtn.addEventListener('click', close);
    closeBtn.addEventListener('click', close);
    backdrop.addEventListener('click', close);
    modal.style.display = 'flex';
}

// Datum formatieren (Mio.-ms Timestamp -> "vor X Tagen" oder Datum)
function fmtTimestamp(ms) {
    if (!ms) return '–';
    const diff = Date.now() - ms;
    const days = Math.floor(diff / (1000*60*60*24));
    if (days < 1) return 'heute';
    if (days < 2) return 'gestern';
    if (days < 14) return `vor ${days} Tagen`;
    return new Date(ms).toLocaleDateString('de-DE');
}

// Liste aller scan_roots laden + rendern (Index-Verwaltung-Sektion).
async function loadAndRenderManageRoots() {
    const card = document.getElementById('roots-card');
    if (!card) return;
    const list = document.getElementById('roots-list');
    const empty = document.getElementById('roots-empty');
    const btnRescanAll = document.getElementById('btn-rescan-all');

    try {
        const r = await fetch('/api/scan-roots');
        const data = await r.json();
        const roots = data.roots || [];

        if (roots.length === 0) {
            if (empty) empty.style.display = 'block';
            if (list)  list.innerHTML = '';
            if (btnRescanAll) btnRescanAll.style.display = 'none';
            return;
        }
        if (empty) empty.style.display = 'none';
        if (btnRescanAll) btnRescanAll.style.display = '';

        list.innerHTML = roots.map(rt => `
            <div class="root-row" data-root-id="${rt.id}">
                <div class="root-row-main">
                    <div class="root-row-label-line">
                        <span class="root-row-label" contenteditable="true"
                              spellcheck="false"
                              data-original="${escapeHtml(rt.label)}"
                              title="Klicken zum Bearbeiten — Enter speichert">${escapeHtml(rt.label)}</span>
                        <span class="root-row-counts">${fmtNum(rt.count)} Präsentationen · ${fmtNum(rt.slideCount)} Folien · Letzter Scan: ${fmtTimestamp(rt.lastScannedAt)}</span>
                    </div>
                    <div class="root-row-path" title="${escapeHtml(rt.fullPath)}">${escapeHtml(rt.fullPath)}</div>
                </div>
                <div class="root-row-actions">
                    <button class="root-action-btn root-rescan" type="button" title="Diesen Pfad neu scannen (erkennt verschwundene Dateien)">↻ Aktualisieren</button>
                    <button class="root-action-btn danger root-delete" type="button" title="Pfad mit allen Folien aus der Datenbank entfernen">✕ Löschen</button>
                </div>
            </div>
        `).join('');

        // Label-Edit: Enter speichert, Esc verwirft, blur speichert.
        list.querySelectorAll('.root-row-label').forEach(el => {
            el.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') { e.preventDefault(); el.blur(); }
                if (e.key === 'Escape') { e.preventDefault(); el.textContent = el.dataset.original; el.blur(); }
            });
            el.addEventListener('blur', async () => {
                const newLabel = el.textContent.trim();
                if (!newLabel || newLabel === el.dataset.original) {
                    el.textContent = el.dataset.original;
                    return;
                }
                const id = el.closest('.root-row').dataset.rootId;
                try {
                    const r = await fetch(`/api/scan-roots/${id}`, {
                        method: 'PATCH',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ label: newLabel }),
                    });
                    const data = await r.json();
                    if (!r.ok) {
                        alert('Fehler beim Umbenennen: ' + (data.error || 'unbekannt'));
                        el.textContent = el.dataset.original;
                        return;
                    }
                    el.dataset.original = newLabel;
                } catch (err) {
                    alert('Netzwerkfehler: ' + err.message);
                    el.textContent = el.dataset.original;
                }
            });
        });

        // Rescan-Button pro Root
        list.querySelectorAll('.root-rescan').forEach(btn => {
            btn.addEventListener('click', async () => {
                const row = btn.closest('.root-row');
                const id = row.dataset.rootId;
                row.classList.add('is-busy');
                try {
                    const r = await fetch(`/api/scan-roots/${id}/rescan`, { method: 'POST' });
                    const data = await r.json();
                    if (!r.ok) {
                        alert('Fehler: ' + (data.error || 'unbekannt'));
                        row.classList.remove('is-busy');
                        return;
                    }
                    // Progress-Card aufmachen + Poller starten
                    document.getElementById('progress-card').style.display = 'block';
                    if (!pollInterval) pollInterval = setInterval(updateProgress, 500);
                } catch (err) {
                    alert('Netzwerkfehler: ' + err.message);
                    row.classList.remove('is-busy');
                }
            });
        });

        // Delete-Button pro Root (mit Bestaetigung)
        list.querySelectorAll('.root-delete').forEach(btn => {
            btn.addEventListener('click', async () => {
                const row = btn.closest('.root-row');
                const id = row.dataset.rootId;
                const label = row.querySelector('.root-row-label').textContent.trim();
                const ok = await showConfirm({
                    title: 'Hauptpfad löschen',
                    body: `Den Hauptpfad <strong>${escapeHtml(label)}</strong> wirklich aus der Datenbank entfernen? Alle dazugehörigen Folien werden gelöscht. <br><br><span style="color:var(--muted); font-size: 0.85rem;">Die PPTX-Dateien auf der Platte bleiben unberührt.</span>`,
                    okText: 'Endgültig löschen',
                    danger: true,
                });
                if (!ok) return;
                try {
                    const r = await fetch(`/api/scan-roots/${id}`, { method: 'DELETE' });
                    const data = await r.json();
                    if (!r.ok) {
                        alert('Fehler: ' + (data.error || 'unbekannt'));
                        return;
                    }
                    loadStats();
                    loadAndRenderManageRoots();
                } catch (err) {
                    alert('Netzwerkfehler: ' + err.message);
                }
            });
        });

        // "Alle aktualisieren"
        const btnAll = document.getElementById('btn-rescan-all');
        if (btnAll && !btnAll.dataset.bound) {
            btnAll.dataset.bound = '1';
            btnAll.addEventListener('click', async () => {
                const ok = await showConfirm({
                    title: 'Alle Hauptpfade aktualisieren',
                    body: 'Alle ' + roots.length + ' Hauptpfade nacheinander neu scannen? Das kann je nach Datenmenge dauern.',
                    okText: 'Los',
                });
                if (!ok) return;
                try {
                    const r = await fetch('/api/scan-roots/rescan-all', { method: 'POST' });
                    const data = await r.json();
                    if (!r.ok) {
                        alert('Fehler: ' + (data.error || 'unbekannt'));
                        return;
                    }
                    document.getElementById('progress-card').style.display = 'block';
                    if (!pollInterval) pollInterval = setInterval(updateProgress, 500);
                } catch (err) {
                    alert('Netzwerkfehler: ' + err.message);
                }
            });
        }
    } catch (err) {
        console.warn('Hauptpfad-Liste konnte nicht geladen werden:', err);
    }
}

// --- Init ---
document.addEventListener('DOMContentLoaded', () => {
    // Config laden (setzt CSS-Variablen + APP_CONFIG-Werte)
    loadAppConfig();
    loadStats();

    // --- Scan-Seite (Folder-Picker, Pfad-Check, Scan starten) ---
    const input = document.getElementById('folder-path');
    if (input) {
        input.value = loadLastPath();
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                checkPath();
            }
        });
        const btnPick = document.getElementById('btn-pick-folder');
        const btnCheck = document.getElementById('btn-check');
        const btnScan = document.getElementById('btn-scan');
        if (btnPick)  btnPick.addEventListener('click', pickFolder);
        if (btnCheck) btnCheck.addEventListener('click', checkPath);
        if (btnScan)  btnScan.addEventListener('click', startScan);

        // Label-Feld: sobald der Nutzer selbst tippt, kein Auto-Fuellen mehr.
        const labelInput = document.getElementById('scan-label');
        if (labelInput) {
            labelInput.addEventListener('input', () => {
                labelInput.dataset.userEdited = labelInput.value.trim() ? '1' : '';
            });
        }

        // Thumb-Befehl-Block immer rendern (zwischen Scan und Fortschritt)
        if (document.getElementById('thumbs-cmd-box')) {
            renderThumbsCmd();
        }

        // Index-Verwaltung-Sektion (nur auf scan.html vorhanden)
        if (document.getElementById('roots-card')) {
            loadAndRenderManageRoots();
        }

        // Falls bei Seitenaufruf bereits ein Pfad in der History steht und ein
        // Scan moeglich waere, koennte hier auto-validiert werden — wir machen
        // das aber nicht, damit der User die Kontrolle behaelt.
    }

    // --- Such-Seite (Suche + Filter) ---
    const searchInput = document.getElementById('search-input');
    if (searchInput) {
        searchInput.addEventListener('input', () => {
            clearTimeout(searchTimer);
            searchTimer = setTimeout(() => runSearch(searchInput.value.trim()), 250);
        });
        const uniqueCb = document.getElementById('filter-unique-only');
        if (uniqueCb) {
            uniqueCb.addEventListener('change', () => runSearch(searchInput.value.trim()));
        }

        // Jahres-Filter wird in renderYearButtons() (nach loadAppConfig)
        // dynamisch gefuellt — inkl. Click-Handlern.

        const sortSelect = document.getElementById('sort-select');
        if (sortSelect) {
            sortSelect.addEventListener('change', () => {
                filterState.sort = sortSelect.value;
                triggerSearch();
            });
        }

        const filenameInput = document.getElementById('filter-filename');
        let filenameTimer = null;
        if (filenameInput) {
            filenameInput.addEventListener('input', () => {
                clearTimeout(filenameTimer);
                filenameTimer = setTimeout(() => {
                    filterState.filename = filenameInput.value.trim();
                    triggerSearch();
                }, 250);
            });
        }

        // Hauptpfade laden + Pills rendern (nur auf Such-Seite)
        loadAndRenderRoots();
    }

    // Reset-Button — sowohl im Warn-Block (btn-reset) als auch permanent
    // in der Index-Verwaltung (btn-reset-all) verfuegbar.
    const resetBtn = document.getElementById('btn-reset');
    if (resetBtn) resetBtn.addEventListener('click', resetIndex);
    const resetAllBtn = document.getElementById('btn-reset-all');
    if (resetAllBtn) resetAllBtn.addEventListener('click', resetIndex);

    // Falls beim Aufruf der Scan-Seite bereits ein Scan laeuft → Polling starten
    if (document.getElementById('progress-card')) {
        fetch('/api/scan/status').then(r => r.json()).then(state => {
            if (state.running) {
                document.getElementById('progress-card').style.display = 'block';
                pollInterval = setInterval(updateProgress, 500);
            }
        }).catch(() => {});
    }
});
