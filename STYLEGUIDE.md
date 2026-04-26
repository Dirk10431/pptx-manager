# pptx-manager — Styleguide & Konventionen

Dieser Guide beschreibt die **Design- und Code-Entscheidungen**, die das pptx-manager-Projekt prägen. Er sitzt zwischen dem allgemeinen [IDDP-Designsystem](https://github.com/Dirk10431/) (lila Corporate-Farbe, gelber Akzent, weißes SVG-Logo) und der konkreten Umsetzung in dieser Anwendung. Wer hier ein neues Feature baut, hält sich an die Patterns aus diesem Dokument — neue Funktionen sehen dann automatisch aus „wie der Rest" und verhalten sich konsistent.

**Im Zweifel:** lieber einem bestehenden Pattern folgen als ein neues erfinden.

---

## 1. Architektur in zwei Sätzen

Lokaler Express-Server mit SQLite (`node:sqlite`-Modul, kein Build-Schritt) und Vanilla-HTML/JS-Frontend. Läuft ausschließlich auf `127.0.0.1:3002`, kein Netzwerkzugriff.

```
Browser  ─HTTP─▶  Express  ─SQL─▶  pptx-manager.db (SQLite)
                     │
                     ├── PowerPoint-COM (PowerShell)  → Thumbnails (PNG-Cache)
                     ├── cmd /c start                 → Datei in Default-App öffnen
                     └── explorer.exe                 → Ordner anzeigen
```

**Datei-Layout:**
```
pptx-manager/
├── server.js          # Express + alle Routes
├── config.json        # Frontend- und Backend-Defaults
├── package.json       # Scripts: start, thumbs, thumbs:dry, thumbs:help
├── lib/
│   ├── db.js          # SQLite-Schema + Migration + Versionierung
│   ├── fingerprint.js # PPTX-Parsing + 3 Hash-Ebenen
│   ├── scanner.js     # Ordner-Scan mit Inkrementalität
│   ├── thumbnailer.js # PowerPoint-COM via PowerShell-File
│   └── config.js      # config.json laden mit Deep-Merge auf Defaults
├── bin/
│   └── generate-thumbs.js   # CLI-Batch-Renderer
├── public/
│   ├── index.html     # Single-Page UI
│   ├── logo.svg       # IDDP-Logo (weiss, fuer lila Hintergrund)
│   ├── css/style.css
│   └── js/app.js
└── data/              # gitignored: db, thumbnails/, ps-scripts/
```

---

## 2. Designsystem (Frontend)

### 2.1 Farben — alle aus CSS-Variablen

Niemals Farben hart in Komponenten kodieren.

```css
:root {
  --bg:           #6667AB;    /* Hintergrund — IDDP-Lila */
  --surface:      #585a9a;    /* Karten, leicht dunkler als bg */
  --surface2:     #4e5089;    /* Inputs, Select-Felder, Sub-Container */
  --border:       #7d7fbe;    /* Rahmen aller Elemente */
  --ink:          #ffffff;    /* Primärtext */
  --muted:        #d4d5ec;    /* Sekundärtext, Hilfshinweise */
  --accent:       #ffffff;    /* Fokus, Highlight */
  --accent2:      #ffe88a;    /* Warmes Gelb — Section-Labels, aktive Pills, CTAs */
  --accent-light: rgba(255,255,255,0.15);
  --warn:         #ffb3b3;    /* Fehler, Warnungen (weiches Rosa) */
  --success:      #a5e8b0;    /* Erfolg (Mint) */
}
```

### 2.2 Layout-Breiten

Anders als das IDDP-Standard-Designsystem (das 720 px für alles vorsieht), nutzt pptx-manager **zwei Breiten**:

- **Header und Footer:** `--max-header: 720px` (kompakt, brand-fokussiert)
- **Inhalt (`.wrap`):** `--max-content: 1200px` (Suchergebnisse mit Tabellen + Thumbnails brauchen Platz)

Beibehalten — nicht in einer einheitlichen Breite zusammenführen.

### 2.3 Typografie

System-Font-Stack, kein Webfont:
```css
font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif;
```

Hierarchie:

| Element | Größe | Weight | Anwendung |
|---|---|---|---|
| Tool-Titel | `clamp(28px, 5vw, 44px)` | 700 | nur im Header, **ein** Akzent-Wort in Gelb + Italic 800 |
| Section-Label | 10 px, 2px letter-spacing, uppercase | 700 | Karten-Überschrift, gelb |
| Karten-Body | 13–14 px | 400 | Standard-Text |
| Stat-Zahl | 22 px | 700 | in `.meta-item` |
| Captions | 11–12 px | 400 | Footer, Tooltips |
| Mono | beliebig | – | `"SF Mono", Monaco, monospace` (Versionsnummer im Footer) |

---

## 3. Komponenten — wann wofür

### 3.1 `.card` für jeden Inhaltsbereich

```html
<section class="card">
  <div class="section-label">Titel des Bereichs</div>
  <p>Optionaler Beschreibungssatz (--muted).</p>
  <!-- ... -->
</section>
```

Jede Karte beginnt mit dem `.section-label`. Niemals `<h1>`/`<h2>` direkt in einer Karte — das Label ist der visuelle Anker.

### 3.2 Buttons — drei Klassen, klare Hierarchie

| Klasse | Wann | Beispiel |
|---|---|---|
| `.go-btn` | **Genau eine** Haupt-Aktion pro Karte | „Scan starten" |
| `.act-btn` | Sekundäre Aktionen | „Pfad prüfen", „Vorkommen zeigen", „Index zurücksetzen" |
| `.btn-pill` | Filter / Toggle | Jahres-Filter, Hauptpfade |
| `.btn-action` | Kleines Icon (📂/▶) | Datei/Ordner öffnen |

**Regel:** Kein neuer Button-Style. Wenn ein bestehender nicht passt, eher das UI umstellen als eine neue Klasse einführen.

### 3.3 Filter-Toolbar

Reihenfolge im pptx-manager (von links nach rechts), alle in der Such-Karte oberhalb der Ergebnisse:

```
┌─ Jahr-Pills ──┬─ Sort-Dropdown ──┬─ Datei-enthält-Input ──┬─ "Nur eindeutig" ─┐
└─ Hauptpfad-Pills (eigene Zeile, alle aktiv = Standard) ──────────────────────┘
```

Markup-Skelett:
```html
<div class="filter-toolbar">
  <div class="filter-row">
    <div class="filter-group">
      <span class="filter-label">Jahr</span>
      <button class="btn-pill active" data-year="all">Alle</button>
      <!-- ... -->
    </div>
    <!-- weitere filter-group -->
  </div>
  <div class="filter-row filter-roots-row">
    <span class="filter-label">Hauptpfade</span>
    <div id="filter-roots" class="filter-roots-list"></div>
  </div>
  <span id="search-summary" class="search-summary"></span>
</div>
```

Bei jeder Filter-Änderung: `triggerSearch()` aufrufen — kein „Filter anwenden"-Knopf.

### 3.4 Suchergebnis-Gruppen

Eine Gruppe = ein eindeutiger `text_hash`. Links Thumbnail, rechts Titel + Snippet + Datei-Quelle + Aktions-Buttons. Wenn die Gruppe mehrere Vorkommen hat, gibt es einen „Vorkommen zeigen"-Knopf, der eine Tabelle aufklappt.

```js
// 1 Repräsentant pro Gruppe gerendert
// Bei >1 Vorkommen: ausklappbare Tabelle
// In jeder Vorkommen-Zeile: Thumbnail (kleiner) + Folie-Index + Dateiname + Aktions-Buttons
```

**Performance-Schutz:** Standardmäßig 30 Vorkommen pro Gruppe sichtbar (Config `search.occurrencesShown`). Bei großen Boilerplate-Gruppen (8.000+ Kopien) würde der Browser sonst einfrieren.

### 3.5 Lightbox

Klick auf jedes Thumbnail → Vollbild-Overlay (85 vh / 95 vw, konfigurierbar).

```css
.thumb-wrap[data-slide-id] { cursor: zoom-in; }
```

Klick auf Hintergrund / Esc / × schließt. Bewusst hochskaliertes PNG (kein neuer Render) — die Unschärfe ist akzeptabel zur visuellen Identifikation.

### 3.6 Aktions-Buttons (📂 / ▶)

Nur bei Suchergebnissen. **Reihenfolge: erst 📂, dann ▶** (Anschauen vor Öffnen).

```js
function actionButtonsHtml(filePath) {
  return `
    <span class="action-buttons">
      <button class="btn-action" data-action="folder" data-path="...">📂</button>
      <button class="btn-action" data-action="file"   data-path="...">▶</button>
    </span>
  `;
}
```

### 3.7 Meta-Grid für Zahlen

Statistiken, Scan-Fortschritt, Mengen-Angaben:
```html
<div class="meta-grid">
  <div class="meta-item">Präsentationen <strong id="…">2491</strong></div>
  <div class="meta-item">Folien         <strong id="…">276635</strong></div>
</div>
```

Großer Wert (22 px) oben, kleines Label (12 px, muted) darunter. Auto-Fit-Grid passt sich an die Breite an.

### 3.8 Warn-Box

Für nicht-kritische Hinweise (Datenbank veraltet, Konfiguration fehlerhaft):
```html
<div class="warn-box">
  <strong>Achtung:</strong> Beschreibung des Problems.
  <div style="margin-top:10px;">
    <button class="act-btn">Aktion zur Behebung</button>
  </div>
</div>
```

Für echte Fehler (Server-Antwort 4xx/5xx) gibt es `.error-box`.

---

## 4. Backend-Patterns

### 4.1 Konfiguration: `config.json` + `/api/config`

Werte, die der Nutzer ohne Code-Änderung anpassen können soll, gehören in `config.json`. Defaults stehen in `lib/config.js`.

| Aktuelle Keys | Was sie tun |
|---|---|
| `search.pageSize` | Initiale Anzahl Gruppen pro Suche (30) |
| `search.maxFtsRows` | FTS5-Limit vor Gruppierung (5000) |
| `search.occurrencesShown` | Max. Vorkommen pro Gruppe in der UI (30) |
| `search.minQueryLength` | Suche startet ab so vielen Zeichen (2) |
| `lightbox.heightVh` | Lightbox-Höhe in % der Bildschirmhöhe (85) |
| `lightbox.maxWidthVw` | Lightbox-Breite-Maximum (95) |

**Beim Hinzufügen einer Nutzbarkeits-Konstante:** in `lib/config.js` als Default eintragen, in `config.json` dokumentieren, im Server an der Verbrauchsstelle aus `config.search.x` lesen.

Frontend lädt die Werte beim Start via `/api/config` und setzt CSS-Variablen entsprechend.

### 4.2 PowerShell-Aufrufe — Norton-vertraeglich

**Verbotene Patterns** (triggern Norton's IDP.HELU.PSE71):
- `-EncodedCommand <base64>` — Base64-PowerShell ist Malware-Pattern
- `Add-Type -MemberDefinition` — startet csc.exe (C#-Compiler) zur Laufzeit, klassischer Trojaner-Indikator
- Häufige PowerShell-Aufrufe in schneller Folge

**Erlaubt:**
- `.ps1`-Datei in `data/ps-scripts/` schreiben, dann `powershell.exe -File <pfad>` aufrufen
- Helfer in `lib/thumbnailer.js`: `writeTempPs1(scriptContent)` und `safeUnlink(file)` — nutzen, statt selbst Skripte zu erzeugen
- Nach Aufruf das Script-File aufräumen (`safeUnlink`, idealerweise mit Timeout für detached Calls)

### 4.3 Lange Pfade umgehen

Windows COM-APIs (PowerPoint, Office) haben ein 255-Zeichen-Limit. Bei tief geschachtelten Pfaden:

```powershell
function Get-OpenablePath {
  param([string]$LongPath)
  if ($LongPath.Length -le 250) { return @{ Path = $LongPath; Temp = $null } }
  # Temp-Kopie mit GUID-Namen, danach loeschen
  $tempName = [System.Guid]::NewGuid().ToString('N') + [System.IO.Path]::GetExtension($LongPath)
  $tempPath = Join-Path $env:TEMP $tempName
  Copy-Item -LiteralPath $LongPath -Destination $tempPath -Force
  return @{ Path = $tempPath; Temp = $tempPath }
}
```

Ist bereits in `lib/thumbnailer.js` als `SHORTPATH_HELPER_PS` enthalten — bei neuen PowerPoint-COM-Aufrufen einbauen.

### 4.4 Datei/Ordner öffnen — pragmatische Methoden

Die richtige Methode hängt vom Inhalt ab:

| Aufgabe | Methode | Warum |
|---|---|---|
| Datei in Default-App öffnen | `cmd /c start "" <pfad>` | Bringt das Fenster in den Vordergrund |
| Ordner anzeigen (normaler Pfad) | `cmd /c start "" <ordner>` | Vordergrund |
| Ordner mit Klammern + Kommas im Pfad | `spawn('explorer.exe', [folder])` | cmd-Quoting scheitert sonst, Fallback öffnet im Hintergrund |
| Ordner-Auswahl-Dialog | PowerShell `FolderBrowserDialog` per `-File` | Norton-konform |
| Datei vorausgewählt im Explorer (`/select`) | **NICHT NUTZEN** | Auf diesem System unzuverlässig — wir öffnen den Ordner stattdessen |

### 4.5 SQLite-Konventionen

- **Embedded:** `node:sqlite` (Node.js 22+ eingebaut, experimentell)
- **WAL-Modus** für gleichzeitiges Lesen/Schreiben (z.B. Scan + Suche parallel)
- **FTS5** für Volltextsuche, mit Triggern für synchrones Update
- **Versionierung:** `meta`-Tabelle mit `fingerprint_version`. Beim Code-Update mit neuer Hash-Logik diese Konstante in `fingerprint.js` erhöhen — der Server warnt dann beim Start, dass ein Reset nötig ist.
- **Migration light:** Bei Schema-Änderungen einer Helfertabelle (z.B. `thumbnails`) per `PRAGMA table_info` prüfen, falls inkompatibel: `DROP TABLE` und neu anlegen. Verlustfrei, weil Daten regenerierbar.

### 4.6 Pagination

Suchergebnisse: `?offset=0&limit=30`. Frontend akkumuliert Gruppen via „Mehr laden"-Button.

Niemals direkt 1000+ Items rendern — Browser-Hauptthread blockiert, Maus ruckelt, Maschine sieht aus als hätte sie sich aufgehängt.

---

## 5. Wenn du eine neue Karte/Funktion baust

Reihenfolge, die fast immer passt:

1. **HTML-Skelett** in `index.html`:
   ```html
   <section class="card">
     <div class="section-label">Mein neuer Bereich</div>
     <p>Kurzer Hinweis-Text.</p>
     <!-- meta-grid / form-group / action-row je nach Inhalt -->
   </section>
   ```

2. **Server-Endpoint** in `server.js`:
   - Pfade nach Schema `/api/<noun>` (z. B. `/api/scan-roots`, `/api/thumb/:slideId`, `/api/open`)
   - Validierung **immer** (DB-Lookup für übergebene Pfade, Range-Checks für Zahlen)
   - Antwort als JSON, Fehler mit `{ error: "..." }`
   - Bei Spawn-Aufrufen: `detached: true, stdio: 'ignore'`, danach `child.unref()` und sofort `res.json(...)` — nicht auf das Kind warten

3. **Frontend-Logik** in `app.js`:
   - Globale Funktion oben definieren (kein Modul, Vanilla JS)
   - Event-Listener im `DOMContentLoaded`-Block am Ende
   - Bei Such-relevanten Filtern: `triggerSearch()` aufrufen, Suchstate zurücksetzen

4. **Konfigurierbare Werte** in `config.json` + `lib/config.js`-Defaults eintragen, Frontend liest sie aus `APP_CONFIG.<bereich>.<key>`

5. **CSS** nur ergänzen, wenn keine bestehende Klasse passt. Neue Klassen folgen der bestehenden Naming-Logik (`btn-*`, `filter-*`, `meta-*`).

6. **Testen:**
   - Server `npm start`, Browser Strg+F5
   - Neuer Endpoint mit `curl` oder direkt aus dem UI
   - Bei UI-Änderungen mobile Breite (600 px) checken
   - Performance: bei großem Datensatz schnell tippen und scrollen — sollte flüssig bleiben

7. **Commit-Message:** kurz beschreibend, deutsch, Co-Author-Zeile am Ende.

---

## 6. Was wir bewusst NICHT haben (und warum)

| Nicht da | Begründung |
|---|---|
| Build-Pipeline (Webpack, Vite, ...) | Single-File-HTML, kein Bundling. Reload = sehen. |
| Frontend-Framework (React, Vue) | Vanilla reicht für 800 Zeilen JS. Frameworks bringen mehr Komplexität als sie hier sparen. |
| Externe CSS-/JS-Dependencies | Alles wird vom lokalen Server geliefert, kein CDN. Funktioniert offline. |
| Authentifizierung | Tool läuft auf `127.0.0.1`, nur lokal. Kein Bedarf. |
| Datenbank-Server | SQLite-Datei reicht — eine Datei, portabel, backup-fähig, bis ~10 Mio Zeilen kein Problem. |
| Worker-Threads / Queue | Scan und Batch-Thumbnail-Generation laufen sequenziell. Resume-fähig durch DB-State. Einfacher. |
| Live-Reload / WebSockets | Periodisches Polling (`/api/scan/status`) reicht für unsere Cadence. |

Wenn du eines davon einbauen willst, **vorher absprechen** — oft gibt es einen einfacheren Weg.

---

## 7. Checkliste für jede Änderung

- [ ] Klassennamen aus Abschnitt 3 verwendet, keine neuen erfunden
- [ ] Farben über CSS-Variablen (kein Hex hart im HTML/JS)
- [ ] `--ink` für Primärtext, `--muted` für Hilfshinweise, `--accent2` für Akzent
- [ ] Genau ein `.go-btn` pro Karte (oder gar keiner)
- [ ] Bei Spawn-Aufrufen: kein `-EncodedCommand`, kein `Add-Type`
- [ ] Bei Server-API: validiert (DB-Lookup für Pfade)
- [ ] Bei langen Listen: pagination/limit (nicht alles auf einmal rendern)
- [ ] Mobile Breite (600 px) sieht okay aus
- [ ] Versions-Marker im Footer hochzählen, falls UI-relevant
- [ ] Commit-Message beschreibt das Warum, nicht nur das Was

---

## Anhang: Wichtige Dateien zum Anschauen vor Änderungen

| Bevor du... | Schau in... |
|---|---|
| ...eine neue Such-Filter-Option baust | `server.js` → `/api/search`, `public/js/app.js` → `runSearch`, `filterState` |
| ...PowerPoint mit COM aufrufst | `lib/thumbnailer.js` → `SHORTPATH_HELPER_PS`, `writeTempPs1` |
| ...etwas im Browser öffnen lässt | `server.js` → `/api/open` (Datei vs. Ordner-Branchen) |
| ...eine neue Karte hinzufügst | `public/index.html` → bestehende Sections; `public/css/style.css` → `.card`, `.section-label` |
| ...eine Konfig-Konstante exponierst | `config.json`, `lib/config.js` (DEFAULTS), `server.js` (`/api/config`), `app.js` (APP_CONFIG) |
