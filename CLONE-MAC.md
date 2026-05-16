# pptx-manager — Installation auf einem frischen Mac

Diese Anleitung beschreibt **Schritt fuer Schritt**, wie das Projekt auf einem komplett unkonfigurierten Mac eingerichtet wird. Voraussetzung ist nur:

- Mac mit aktuellem macOS (Apple Silicon empfohlen, Intel funktioniert genauso)
- Admin-Rechte (also dein normaler Benutzer-Account, mit dem du Software installieren kannst)
- Internet-Verbindung waehrend der Installation

Am Ende laeuft der pptx-manager unter [http://127.0.0.1:3002](http://127.0.0.1:3002) im Browser, scannt PPTX-Ordner, findet Duplikate und zeigt Folien-Vorschauen.

---

## 0. Terminal oeffnen

Alle folgenden Schritte geschehen im **Terminal**. Oeffnen:

1. Cmd + Leertaste → "Terminal" eintippen → Enter

Lass das Fenster die ganze Anleitung ueber offen.

---

## 1. Apple Command Line Tools installieren

Liefert `git`, `make`, Compiler — wird von Homebrew und npm gebraucht.

```bash
xcode-select --install
```

Es oeffnet sich ein Dialog → **Installieren** klicken → 5–10 Minuten warten. Wenn die Tools schon da sind, sagt der Befehl das und tut nichts.

---

## 2. Homebrew installieren

Homebrew ist der Paketmanager fuer macOS — ueber ihn installieren wir alles weitere.

```bash
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
```

Du wirst nach dem Mac-Passwort gefragt → eintippen (du siehst es nicht, das ist normal) → Enter.

**Nach der Installation** zeigt Homebrew zwei Befehle an, die du in dein Profil eintragen sollst, damit `brew` bei jedem neuen Terminal-Fenster verfuegbar ist. Der wichtigste sieht so aus (bei Apple Silicon):

```bash
echo >> ~/.zprofile
echo 'eval "$(/opt/homebrew/bin/brew shellenv)"' >> ~/.zprofile
eval "$(/opt/homebrew/bin/brew shellenv)"
```

> Bei Intel-Macs steht statt `/opt/homebrew` der Pfad `/usr/local`. Folge einfach genau dem, was Homebrew dir nach der Installation anzeigt.

Pruefen, ob es klappt:

```bash
brew --version
```

Sollte eine Versionsnummer zeigen (z.B. `Homebrew 4.x.x`).

---

## 3. Node.js und LibreOffice installieren

```bash
brew install node
brew install --cask libreoffice
```

- **Node.js** liefert `node` und `npm` (das Tool laeuft auf Node.js 22+).
- **LibreOffice** wird fuer die Folien-Vorschaubilder gebraucht. Es muss nicht gestartet werden — der pptx-manager spricht es im Hintergrund an.

Pruefen:

```bash
node --version    # mind. v22.x
npm --version     # irgendeine Versionsnummer
which soffice     # /opt/homebrew/bin/soffice (oder /usr/local/bin/soffice)
```

---

## 4. Repository klonen

Lege einen Ordner fuer das Projekt an und klone das Repo dort hinein.

### Variante A — HTTPS (einfach, ohne SSH-Key)

```bash
mkdir -p ~/Projekte
cd ~/Projekte
git clone https://github.com/Dirk10431/pptx-manager.git
cd pptx-manager
```

Beim ersten `git push` (also wenn du spaeter Aenderungen zurueckschicken willst) fragt GitHub nach Benutzername + **Personal Access Token** (kein Passwort mehr). Den Token erstellst du in GitHub unter *Settings → Developer settings → Personal access tokens*.

### Variante B — SSH-Key einrichten (komfortabler)

```bash
ssh-keygen -t ed25519 -C "deine-mail@beispiel.de"
# Bei den Fragen einfach 3x Enter druecken (Default-Pfad, kein Passwort)
cat ~/.ssh/id_ed25519.pub
```

Den ausgegebenen Public-Key (`ssh-ed25519 AAAA...`) **komplett kopieren** und in GitHub unter *Settings → SSH and GPG keys → New SSH key* einfuegen.

Danach:

```bash
mkdir -p ~/Projekte
cd ~/Projekte
git clone git@github.com:Dirk10431/pptx-manager.git
cd pptx-manager
```

---

## 5. Abhaengigkeiten installieren

Im Projektordner:

```bash
npm install
```

Laedt alle Pakete aus `package.json` (Express, jszip, fast-xml-parser, pdf-parse). Beim ersten Mal 1–2 Minuten.

---

## 6. Server starten

```bash
./start.sh
```

Falls das Skript nicht ausfuehrbar ist:

```bash
chmod +x start.sh
./start.sh
```

Das Skript prueft `node_modules`, startet den Server **und oeffnet automatisch den Browser** auf [http://127.0.0.1:3002](http://127.0.0.1:3002).

> Server beenden: **Ctrl + C** im Terminal-Fenster.

---

## 7. Ersten Scan durchfuehren

1. Im Browser auf **Scan** klicken
2. **Ordner waehlen** → einen Ordner mit `.pptx`-Dateien auswaehlen
   - Beim ersten Klick fragt macOS einmalig: *"Node moechte System Events steuern"* → **OK**
3. Optional ein Anzeige-Label vergeben → **Scan starten**
4. Der Scan-Fortschritt laeuft durch, danach steht das Ergebnis in der DB

---

## 8. Folien-Vorschaubilder generieren

Im Terminal (im Projektordner):

```bash
npm run thumbs
```

Was passiert:
- LibreOffice startet im Hintergrund (kein GUI-Fenster), konvertiert jede PPTX zu einem PDF
- Aus den PDF-Seiten werden 480×270-PNGs (Apples PDFKit, macOS-Bordmittel)
- PNGs landen in `data/thumbnails/` und werden in der Web-UI angezeigt

Nuetzliche Optionen:

```bash
npm run thumbs -- --limit 5     # erstmal nur 5 Folien testen
npm run thumbs:dry              # zeigt nur, was gemacht wuerde
npm run thumbs:help             # alle Optionen
```

> Strg+C bricht sauber ab; beim naechsten Lauf wird an der gleichen Stelle weitergemacht (Resume).

---

## 9. Was wird **nicht** mit gecloned

Per `.gitignore` ausgeschlossen — entstehen erst lokal beim Benutzen:

| Pfad | Zweck |
|---|---|
| `node_modules/` | wird durch `npm install` neu erzeugt |
| `data/` | SQLite-Datenbank + Thumbnail-Cache (deine Scan-Ergebnisse) |
| `backups/` | automatische Sicherungen vor Phase-2-Aenderungen |
| `_old/` | alte Python-Version (nur auf dem Windows-Rechner) |

Das ist **so gewollt**: Jeder Rechner hat seine eigene lokale Datenbank.

---

## 10. Updates vom Server holen

Wenn jemand etwas am Code geaendert hat:

```bash
cd ~/Projekte/pptx-manager
git pull
npm install        # nur falls package.json sich geaendert hat
```

Eigene Aenderungen zurueckschicken (nur fuer Entwickler):

```bash
git status                       # Was hat sich geaendert?
git add <dateien>                # Bestimmte Dateien stagen
git commit -m "Kurze Beschreibung"
git push
```

---

## 11. Troubleshooting

### "Permission denied" bei `./start.sh`
```bash
chmod +x start.sh
```

### Port 3002 bereits belegt
Anderer Prozess laeuft auf 3002. Pruefen:
```bash
lsof -nP -iTCP:3002 -sTCP:LISTEN
```
Notfalls den Prozess beenden:
```bash
kill <PID>
```

### `npm install` schlaegt fehl
- Internet-Verbindung pruefen
- Node-Version: `node --version` muss `v22.x` oder neuer sein
- Cache loeschen: `npm cache clean --force` und erneut `npm install`

### Server startet, aber Fehler "node:sqlite is not defined"
Node-Version zu alt. `node --version` muss `v22.x` oder hoeher zeigen. Mit Homebrew aktualisieren: `brew upgrade node`.

### Ordner-Dialog ("Ordner waehlen") oeffnet sich nicht
Beim ersten Klick fragt macOS, ob das Terminal/Node "System Events" steuern darf. **Erlauben** → ab dann funktioniert der Dialog. Spaeter erreichbar unter *Systemeinstellungen → Datenschutz & Sicherheit → Automation*.

### Thumbnails werden nicht erzeugt
- LibreOffice installiert? `which soffice` muss einen Pfad liefern.
- Falls nicht: `brew install --cask libreoffice` nochmal ausfuehren.

### `brew` nach Neustart nicht gefunden
Die `eval "$(/opt/homebrew/bin/brew shellenv)"`-Zeile in `~/.zprofile` fehlt. Schritt 2 nochmal anschauen.

### Browser oeffnet sich nicht automatisch
Manuell oeffnen: [http://127.0.0.1:3002](http://127.0.0.1:3002)

---

## 12. Projekt-Spezifika

- **Lokal und offline** — keine Internet-Verbindung im Betrieb noetig (nur fuer die Installation oben)
- **Port 3002** ist fest verdrahtet (nicht 3000/3001)
- **Nur 127.0.0.1** — der Server ist nicht im Netzwerk erreichbar (Sicherheit)
- **PPTX-Dateien werden in Phase 1 nur gelesen** — Schreibzugriffe kommen erst spaeter mit Backup + Dry-Run
- **Thumbnail-Pipeline auf Mac**: PPTX → PDF (LibreOffice headless) → PNG (Apple PDFKit) — ohne externe Tools wie pdftoppm/Ghostscript
- **PowerPoint fuer Mac** ist als Fallback unterstuetzt, aber **nicht empfohlen**: Microsoft Office laeuft sandboxed und fragt pro Datei nach einer Berechtigung — bei vielen PPTX untragbar. Das Tool nutzt automatisch LibreOffice, sobald `soffice` im System ist.

---

**Fertig.** Bei Detailfragen: `CLAUDE.md` und `STYLEGUIDE.md` enthalten die Projekt-Konventionen.
