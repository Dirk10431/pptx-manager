# pptx-manager — Klonen und Starten auf dem Mac

Diese Anleitung beschreibt Schritt fuer Schritt, wie das Projekt **pptx-manager** auf einem Mac eingerichtet wird — vom Klonen aus GitHub bis zum laufenden Server im Browser.

---

## 1. Voraussetzungen pruefen

Oeffne das **Terminal** (Cmd + Leertaste, "Terminal" eingeben, Enter).

Pruefe der Reihe nach, ob die noetigen Werkzeuge installiert sind:

```bash
git --version
node --version
npm --version
```

Erwartete Ausgabe: Versionsnummern. Wenn ein Befehl unbekannt ist, fehlt das Programm.

### Falls Git fehlt
```bash
xcode-select --install
```
Ein Installations-Dialog erscheint — durchklicken, fertig.

### Falls Node.js fehlt
Empfohlen: Installiere Node.js ueber [nodejs.org](https://nodejs.org/) (LTS-Version).
Alternativ mit Homebrew:
```bash
brew install node
```

> **Node 22 oder neuer** ist Pflicht — das Tool nutzt das in Node 22+ eingebaute `node:sqlite`-Modul. Aeltere Node-Versionen starten den Server nicht.

### Microsoft PowerPoint (fuer Thumbnails)

Die Folien-Vorschaubilder werden vom Mac per AppleScript an Microsoft PowerPoint geschickt. Damit das funktioniert, muss **PowerPoint fuer Mac installiert** sein (egal ob Microsoft 365 oder Standalone).

Ohne PowerPoint laufen Scan, Volltextsuche und Duplikat-Anzeige normal — nur die Vorschau-Bilder fehlen.

Zusaetzlich genutzt: `sips` (zum Skalieren der PNGs auf 480×270). Das ist macOS-Bordmittel, nichts zu installieren.

---

## 2. Repository klonen

Lege einen Ordner fuer Projekte an (falls nicht vorhanden) und klone das Repo dort hinein:

```bash
mkdir -p ~/Projekte
cd ~/Projekte
git clone git@github.com:Dirk10431/pptx-manager.git
cd pptx-manager
```

> **Hinweis SSH-Key:** Der obige Befehl nutzt SSH. Wenn auf dem Mac noch kein SSH-Key bei GitHub hinterlegt ist, gibt es zwei Wege:
>
> **Variante A — HTTPS (einfach, kein Key noetig):**
> ```bash
> git clone https://github.com/Dirk10431/pptx-manager.git
> ```
> Beim ersten `git push` fragt GitHub nach Benutzername + Personal Access Token.
>
> **Variante B — SSH-Key auf dem Mac einrichten:**
> ```bash
> ssh-keygen -t ed25519 -C "dr.dirk.peters@gmail.com"
> cat ~/.ssh/id_ed25519.pub
> ```
> Den ausgegebenen Public-Key bei GitHub unter **Settings → SSH and GPG keys → New SSH key** einfuegen. Danach funktioniert `git@github.com:...`.

---

## 3. Abhaengigkeiten installieren

Im Projektordner (`pptx-manager/`):

```bash
npm install
```

Das laedt alle Pakete aus `package.json` herunter (Express, jszip, fast-xml-parser, pdf-parse).
Beim ersten Mal kann das 1–2 Minuten dauern.

---

## 4. Server starten

### Variante A — Komfort-Skript (empfohlen)

```bash
./start.sh
```

Falls das Skript nicht ausfuehrbar ist:
```bash
chmod +x start.sh
./start.sh
```

Das Skript prueft `node_modules`, startet den Server **und oeffnet automatisch den Browser** auf `http://127.0.0.1:3002`.

### Variante B — Direkt per npm
```bash
npm start
```
Danach manuell den Browser oeffnen: **http://127.0.0.1:3002**

---

## 5. Erste Schritte in der App

1. Im Browser: **http://127.0.0.1:3002**
2. Im Menue **Scan** auf der Subpage `/scan.html` einen Ordner mit `.pptx`-Dateien angeben
3. Scan starten — die Datenbank in `data/` wird angelegt
4. Anschliessend Volltextsuche, Duplikate und Status-Dashboard nutzen

---

## 6. Was wird **nicht** mit gecloned?

Per `.gitignore` ausgeschlossen — diese Ordner/Dateien entstehen erst lokal:

| Pfad | Zweck |
|---|---|
| `node_modules/` | Wird durch `npm install` neu erzeugt |
| `data/` | SQLite-Datenbank (deine lokalen Scan-Ergebnisse) |
| `backups/` | Automatische Sicherungen vor Phase-2-Aenderungen |
| `_old/` | Alte Python-Version (nur auf dem Windows-Rechner) |

Das ist **so gewollt**: Jeder Rechner hat seine eigene lokale Datenbank.

---

## 7. Aenderungen zurueck zu GitHub schicken

Nach Code-Aenderungen auf dem Mac:

```bash
git status                       # Was hat sich geaendert?
git add <dateien>                # Bestimmte Dateien stagen
git commit -m "Kurze Beschreibung"
git push
```

Vor jeder Arbeit den aktuellen Stand vom Server holen:
```bash
git pull
```

---

## 8. Troubleshooting

### "Permission denied" bei `./start.sh`
```bash
chmod +x start.sh
```

### Port 3002 bereits belegt
Anderer Prozess laeuft auf 3002. Pruefen mit:
```bash
lsof -i :3002
```
Notfalls den Prozess beenden:
```bash
kill <PID>
```

### `npm install` schlaegt fehl
- Internet-Verbindung pruefen
- Node-Version checken: `node --version` (muss 22 oder neuer sein)
- Cache loeschen: `npm cache clean --force` und erneut `npm install`

### Server startet, aber Fehler "node:sqlite is not defined"
Node-Version zu alt. `node --version` muss `v22.x` oder hoeher zeigen. Mit Homebrew aktualisieren: `brew upgrade node` oder das LTS-Paket von [nodejs.org](https://nodejs.org/) installieren.

### Ordner-Dialog ("Ordner waehlen") oeffnet sich nicht
Beim ersten Klick fragt macOS, ob das Terminal/Node "System Events" steuern darf. **Erlauben** → ab dann funktioniert der Dialog. Spaeter erreichbar unter *Systemeinstellungen → Datenschutz & Sicherheit → Automation*.

### Thumbnails werden nicht erzeugt
- PowerPoint installiert? `ls /Applications | grep -i powerpoint` muss `Microsoft PowerPoint.app` zeigen.
- Beim ersten Aufruf von `npm run thumbs` fragt macOS, ob Node Microsoft PowerPoint steuern darf. **Erlauben.**
- `sips`-Binary vorhanden? `which sips` muss einen Pfad liefern (auf jedem Mac vorinstalliert).

### Browser oeffnet sich nicht automatisch
Manuell oeffnen: **http://127.0.0.1:3002**

---

## 9. Projekt-Spezifika (siehe auch `CLAUDE.md`)

- **Lokal und offline** — keine Internet-Verbindung noetig im Betrieb
- **Port 3002** ist fest verdrahtet (nicht 3000/3001 wie bei den Hetzner-Projekten)
- **Nur 127.0.0.1** — der Server ist nicht im Netzwerk erreichbar (Sicherheit)
- **PPTX-Dateien werden in Phase 1 nur gelesen** — Schreibzugriffe kommen erst spaeter mit Backup + Dry-Run

---

**Fertig.** Bei Fragen: `CLAUDE.md` und `STYLEGUIDE.md` enthalten die Projekt-Konventionen.
