// =============================================================
// app.js - Frontend-Logik pptx-manager
// =============================================================

// --- Statistik laden ---
async function loadStats() {
    try {
        const response = await fetch('/api/stats');
        const data = await response.json();
        document.getElementById('stat-presentations').textContent = data.presentations;
        document.getElementById('stat-slides').textContent = data.slides;
    } catch (err) {
        console.error('Fehler beim Laden der Stats:', err);
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

// --- Pfad pruefen ---
async function checkPath() {
    const folderPath = document.getElementById('folder-path').value.trim();
    const resultEl = document.getElementById('check-result');
    const scanBtn = document.getElementById('btn-scan');

    if (!folderPath) {
        resultEl.innerHTML = '<p style="color: var(--danger);">Bitte einen Pfad eingeben.</p>';
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
            resultEl.innerHTML = `<p style="color: var(--danger);">Fehler: ${data.error}</p>`;
            scanBtn.disabled = true;
            return;
        }

        if (data.fileCount === 0) {
            resultEl.innerHTML = `<p style="color: var(--warning);">Keine PPTX-Dateien in diesem Ordner gefunden.</p>`;
            scanBtn.disabled = true;
            return;
        }

        let html = `<p><span class="badge badge-success">${data.fileCount} PPTX-Dateien</span> gefunden.</p>`;
        if (data.sample && data.sample.length > 0) {
            html += `<p style="color: var(--text-light); font-size: 0.9rem;">Beispiele: ${data.sample.join(', ')}${data.fileCount > data.sample.length ? ', ...' : ''}</p>`;
        }
        resultEl.innerHTML = html;
        scanBtn.disabled = false;
        saveLastPath(folderPath);
    } catch (err) {
        resultEl.innerHTML = `<p style="color: var(--danger);">Netzwerkfehler: ${err.message}</p>`;
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

async function runSearch(term) {
    const resultsEl = document.getElementById('search-results');
    if (!term || term.length < 2) {
        resultsEl.innerHTML = '';
        return;
    }

    resultsEl.innerHTML = '<p style="color: var(--text-light);">Suche...</p>';

    try {
        const response = await fetch('/api/search?q=' + encodeURIComponent(term));
        const data = await response.json();

        if (data.error) {
            resultsEl.innerHTML = `<p style="color: var(--danger);">${escapeHtml(data.error)}</p>`;
            return;
        }

        if (!data.results || data.results.length === 0) {
            resultsEl.innerHTML = '<p style="color: var(--text-light);">Keine Treffer.</p>';
            return;
        }

        const rows = data.results.map(r => `
            <tr>
                <td><span class="badge">Folie ${r.slide_index}</span></td>
                <td><strong>${escapeHtml(r.title || '(ohne Titel)')}</strong></td>
                <td style="font-size: 0.9rem; color: var(--text-light);">${r.snippet || ''}</td>
                <td style="font-size: 0.85rem;" title="${escapeHtml(r.file_path)}">${escapeHtml(r.file_name)}</td>
            </tr>
        `).join('');

        resultsEl.innerHTML = `
            <p style="margin: 1rem 0 0.5rem;"><span class="badge badge-success">${data.results.length} Treffer</span></p>
            <table class="table">
                <thead>
                    <tr>
                        <th>Folie</th>
                        <th>Titel</th>
                        <th>Textauszug</th>
                        <th>Datei</th>
                    </tr>
                </thead>
                <tbody>${rows}</tbody>
            </table>
        `;
    } catch (err) {
        resultsEl.innerHTML = `<p style="color: var(--danger);">Netzwerkfehler: ${escapeHtml(err.message)}</p>`;
    }
}

// --- Init ---
document.addEventListener('DOMContentLoaded', () => {
    loadStats();

    const input = document.getElementById('folder-path');
    input.value = loadLastPath();

    document.getElementById('btn-check').addEventListener('click', checkPath);
    document.getElementById('btn-scan').addEventListener('click', startScan);

    // Enter im Pfad-Feld = Pfad pruefen
    input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            checkPath();
        }
    });

    // Suchfeld: Tippen triggert debounce-Suche
    const searchInput = document.getElementById('search-input');
    searchInput.addEventListener('input', () => {
        clearTimeout(searchTimer);
        searchTimer = setTimeout(() => runSearch(searchInput.value.trim()), 250);
    });
});
