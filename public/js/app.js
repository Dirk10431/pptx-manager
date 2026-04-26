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
    } catch (err) {
        console.warn('Config konnte nicht geladen werden, nutze Defaults:', err);
    }
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

// --- Statistik laden ---
async function loadStats() {
    try {
        const response = await fetch('/api/stats');
        const data = await response.json();
        document.getElementById('stat-presentations').textContent = data.presentations;
        document.getElementById('stat-slides').textContent = data.slides;

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
async function resetIndex() {
    if (!confirm('Alle gescannten Folien-Daten wirklich loeschen? Die PPTX-Dateien bleiben unberuehrt. Danach muss neu gescannt werden.')) {
        return;
    }
    try {
        const response = await fetch('/api/reset', { method: 'POST' });
        if (!response.ok) {
            const err = await response.json();
            alert('Fehler: ' + (err.error || 'Unbekannter Fehler'));
            return;
        }
        loadStats();
        document.getElementById('search-results').innerHTML = '';
        document.getElementById('search-summary').textContent = '';
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

    document.getElementById('btn-scan').disabled = true;
    document.getElementById('btn-check').disabled = true;
    document.getElementById('progress-card').style.display = 'block';

    try {
        const response = await fetch('/api/scan', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ folderPath }),
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
            document.querySelector('#progress-card h2').textContent = 'Scan abgeschlossen';
            document.getElementById('btn-scan').disabled = false;
            document.getElementById('btn-check').disabled = false;
            document.getElementById('progress-current').textContent = '-';
            loadStats();
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

// --- Bash-Hinweise auf der Scan-Seite ---
// Sobald ein Pfad erfolgreich validiert wurde, zeigen wir ein Code-Block
// mit den Befehlen, die der Nutzer in einem zweiten Bash-Fenster ausfuehren
// soll, um nach dem Scan die Thumbnails zu generieren.
async function renderBashHints() {
    const box = document.getElementById('bash-hints');
    if (!box) return;
    let projectRoot = '<projekt-pfad>';
    try {
        const r = await fetch('/api/project-info');
        if (r.ok) {
            const d = await r.json();
            projectRoot = d.projectRoot || projectRoot;
        }
    } catch {}
    const code = `cd "${projectRoot}"\nnpm run thumbs`;
    box.innerHTML = `
        <div class="prompt-box">
            <span class="prompt-label">▸ Nach dem Scan: Thumbnails generieren</span>
            <button type="button" class="copy-btn" data-copy-target="bash-cmd">📋 Kopieren</button>
            <span class="prompt-content" id="bash-cmd">${escapeHtml(code)}</span>
        </div>
        <p style="font-size:12px; color:var(--muted); margin-top:8px;">
            Oeffne ein zweites Git-Bash-Fenster, fuege die zwei Zeilen ein und druecke Enter. Der CLI-Lauf braucht je nach Datenmenge zwischen 10 Minuten und 2 Stunden — abbrechbar mit <strong>Strg+C</strong>, beim naechsten Start macht er an der Stelle weiter.
        </p>
    `;
    box.style.display = 'block';
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

// Bei jedem erfolgreichen Pfad-Check Bash-Hinweise (re-)rendern.
// Wird in checkPath() unten getriggert via Hook.
const __originalCheckPath = checkPath;
checkPath = async function patchedCheckPath() {
    await __originalCheckPath();
    // Wenn der Scan-Button nicht mehr disabled ist, war der Check erfolgreich
    const scanBtn = document.getElementById('btn-scan');
    if (scanBtn && !scanBtn.disabled) {
        renderBashHints();
    }
};

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

        // Jahres-Filter
        document.querySelectorAll('.year-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                document.querySelectorAll('.year-btn').forEach(b => b.classList.remove('active'));
                btn.classList.add('active');
                filterState.year = btn.dataset.year;
                triggerSearch();
            });
        });

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

    // Reset-Button (kann auf beiden Seiten vorkommen)
    const resetBtn = document.getElementById('btn-reset');
    if (resetBtn) resetBtn.addEventListener('click', resetIndex);

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
