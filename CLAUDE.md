# CLAUDE.md - pptx-manager

## Projektbeschreibung

Lokale Web-App zur **Verwaltung und Duplikatsuche in PPTX-Sammlungen** (ca. 500 Dateien von IDDP / Amperian). Ersetzt die Idee eines Python-CLI-Tools durch eine browser-basierte Oberflaeche mit gleicher Funktionalitaet, aber besserer Bedienbarkeit.

**Betreiber:** Ingenieurbuero Dr.-Ing. Dirk Peters
**Laeuft:** Nur lokal auf `127.0.0.1:3002` (kein Internet-Zugriff)

## Kernfunktionen (Zielzustand)

- Ordner scannen und Folien indizieren (inkrementell)
- Volltextsuche ueber alle Folien
- Duplikate finden auf 3 Ebenen:
  - **exact**: byte-identische Folien-XML
  - **text**: gleicher Text, andere Formatierung
  - **structure**: gleiches Layout, anderer Inhalt
- "Wo wird diese Folie verwendet?" (Where-Used)
- Varianten (aehnlich, nicht identisch) via MinHash
- **Phase 2**: Folie propagieren (neue Version in alle Dateien zurueckschreiben, mit Backup + Dry-Run)

## Designprinzipien

1. **PPTX-Dateien sind heilig.** Phase 1 schreibt NUR in die DB. Erst Phase 2 schreibt an Dateien und dann IMMER mit Backup + Dry-Run-Vorschau.
2. **Lokal und offline.** Kein Netzwerk, keine API-Calls. SQLite-Datei ist die Quelle der Wahrheit.
3. **Inkrementell.** Zweiter Scan muss schnell sein (mtime + file_hash).
4. **Drei Hash-Ebenen**: `exact_hash`, `text_hash`, `structure_hash` - nicht vermischen.
5. **Deutsche UI-Texte**, englische Variablennamen/Kommentare.

## Technologie-Stack

| Komponente | Technologie |
|---|---|
| Backend | Node.js + Express |
| Frontend | Vanilla HTML/CSS/JS |
| Datenbank | SQLite (`better-sqlite3`) mit FTS5-Volltextsuche |
| PPTX-Parsing | `jszip` + `fast-xml-parser` |
| Hosting | Lokal, nur 127.0.0.1 (kein Netzwerk) |

## Dateistruktur

```
pptx-manager/
├── server.js              # Express-Server
├── package.json           # Dependencies
├── lib/
│   ├── db.js              # SQLite-Schema + Init
│   ├── scanner.js         # (TODO) Ordner-Scan
│   ├── fingerprint.js     # (TODO) Hash-Berechnung
│   └── queries.js         # (TODO) Duplikate, Where-Used, Varianten
├── public/
│   ├── index.html
│   ├── css/style.css
│   └── js/app.js
├── data/                  # SQLite-Datei (gitignored)
├── backups/               # Backups vor propagate (gitignored)
├── _old/                  # Alte Python-Version (Referenz)
└── CLAUDE.md              # Dieses Dokument
```

## Start

```bash
npm install
npm start
# Browser: http://127.0.0.1:3002
```

Oder per Doppelklick auf `start.bat` (Windows) bzw. `start.sh` (Mac).

## Entwicklungsstand

- [x] Grundgeruest (Server, HTML/CSS, SQLite-Schema, Stats-API)
- [ ] PPTX-Parser (fingerprint.js)
- [ ] Scanner (scanner.js)
- [ ] Volltextsuche UI
- [ ] Duplikate-Ansicht
- [ ] Where-Used-Ansicht
- [ ] Varianten (MinHash)
- [ ] Phase 2: Propagate
