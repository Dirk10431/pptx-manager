# pptx-manager — Installieren und Starten auf Windows

Diese Anleitung beschreibt Schritt fuer Schritt, wie das Projekt **pptx-manager** auf einem Windows-Rechner eingerichtet wird — von Null (frischer Windows-11-PC) bis zur laufenden App im Browser. **Kein Git, keine Bash, kein Programmiererwissen noetig.**

Ziel-Empfaenger dieser Anleitung: jemand, der die App benutzen will, nicht entwickeln. Wer am Code arbeiten moechte, schaut zusaetzlich in `CLAUDE.md` und `STYLEGUIDE.md`.

---

## 1. Voraussetzungen installieren

### Node.js (Pflicht)

1. Browser oeffnen, https://nodejs.org/ aufrufen
2. Den linken/groesseren Button mit der **LTS-Version** anklicken (aktuell Node 22.x oder neuer)
3. Die heruntergeladene Datei (`.msi`) doppelklicken und durch den Installer-Dialog gehen — alle Defaults uebernehmen
4. Pruefen, ob es geklappt hat: `Windows-Taste`, `cmd` eintippen, Enter. Im schwarzen Fenster:
   ```
   node --version
   ```
   Erwartete Ausgabe: `v22.x.x` oder hoeher.

> Falls "node ist nicht als interner oder externer Befehl erkannt" kommt: einmal **PC neu starten**. Der Installer setzt zwar den PATH, aber bestehende Fenster haben die Aenderung noch nicht.

### Microsoft PowerPoint (optional, fuer Folien-Vorschaubilder)

Wenn Office 365, 2021 oder 2019 mit PowerPoint bereits installiert ist, ist nichts zu tun. Ohne PowerPoint funktioniert die App vollstaendig — nur die Folien-Vorschaubilder (Thumbnails) lassen sich nicht erzeugen. Suche, Duplikat-Analyse, Where-Used etc. laufen unabhaengig davon.

---

## 2. pptx-manager herunterladen

1. Im Browser https://github.com/Dirk10431/pptx-manager oeffnen
2. Den gruenen Button **"Code"** anklicken → **"Download ZIP"**
3. ZIP irgendwohin entpacken — empfohlene Orte:
   - `C:\Tools\pptx-manager\`
   - oder `C:\Users\<DEINNAME>\Documents\pptx-manager\`
4. Im entpackten Ordner muessen u.a. diese Dateien liegen:
   - `start.bat`
   - `server.js`
   - `package.json`
   - Ordner `lib`, `public`, `bin`

> Heisst der entpackte Ordner `pptx-manager-main`? Einfach umbenennen zu `pptx-manager` (optional, nur kosmetisch).

---

## 3. Erststart

**Doppelklick auf `start.bat`** im Projekt-Ordner.

> **SmartScreen-Warnung beim ersten Mal:** Windows zeigt evtl. "Windows hat Ihren PC geschuetzt". Auf **"Weitere Informationen"** klicken → **"Trotzdem ausfuehren"**.

Beim allerersten Start passiert:

1. Ein schwarzes Konsolenfenster oeffnet sich
2. Meldung: **"Erstinstallation: npm install laeuft..."** — laedt ca. 80 MB Abhaengigkeiten herunter, dauert 1–3 Minuten beim ersten Mal
3. Server startet mit der Meldung **"pptx-manager — lokal gestartet — URL: http://127.0.0.1:3002"**
4. Der Browser oeffnet sich automatisch auf der App

**Beenden:** Konsolenfenster schliessen — oder `Strg + C` druecken, dann eine beliebige Taste.

> Beim zweiten Start ist das `npm install` schon erledigt — `start.bat` springt sofort zum Server-Start (binnen weniger Sekunden).

---

## 4. Komfort: Desktop-Shortcuts (optional)

Nach dem ersten erfolgreichen Start kannst du dir zwei kleine Helfer-Batches anlegen, die du dann auf den Desktop ziehst — damit startest du die App in Zukunft mit einem Doppelklick.

1. Im pptx-manager-Ordner: **Doppelklick auf `make-shortcuts.bat`**
2. Das Tool erzeugt zwei Dateien im Projekt-Ordner:
   - **`pptx-open.bat`** — Startet den Server im Hintergrund (falls er nicht schon laeuft) und oeffnet den Browser auf die App. **Das ist dein normaler Tagesstart.**
   - **`pptx-thumbs.bat`** — Startet die Thumbnail-Generierung in einem Konsolenfenster mit Fortschrittsanzeige.
3. Beide Dateien einfach **auf den Desktop ziehen** (oder kopieren). Der Pfad zum Projekt-Ordner ist eingebaut, deshalb funktionieren sie von ueberall.

> Falls du den Projekt-Ordner spaeter mal verschiebst: einfach `make-shortcuts.bat` nochmal doppelklicken, dann werden die zwei Dateien mit dem neuen Pfad neu erzeugt.

**Tipp:** Statt zu kopieren kannst du auch eine **Verknuepfung** auf den Desktop legen — Rechtsklick auf die `.bat`-Datei → **"Senden an"** → **"Desktop (Verknuepfung erstellen)"**.

---

## 5. Erste Schritte in der App

1. Im Browser: http://127.0.0.1:3002 (oeffnet sich von selbst)
2. Oben rechts: **"Index verwalten / Ordner scannen"** anklicken
3. **"Ordner waehlen…"** druecken — Windows-Dialog erscheint
4. Ordner mit den PPTX-Dateien aussuchen (z.B. `C:\Users\<DEINNAME>\OneDrive\Praesentationen\`). Unterordner werden automatisch mitgescannt
5. Optional ein eigenes **Anzeige-Label** vergeben (sonst nimmt die App den Ordnernamen)
6. **"Scan starten"** — Fortschrittsbalken zeigt Live-Status
7. Nach dem Scan: zurueck zur Hauptseite, im Suchfeld tippen — Treffer kommen sofort

### Folien-Vorschaubilder erzeugen (optional, dauert)

1. Im pptx-manager-Ordner Rechtsklick → **"Im Terminal oeffnen"** (Windows 11)
   - Alternativ Win 10: Shift+Rechtsklick → **"PowerShell-Fenster hier oeffnen"**
2. In der App auf der Scan-Seite gibt es den Block **"Thumbnails generieren / aktualisieren"** — Befehlszeile mit 📋-Kopierbutton
3. Im Terminal einfuegen (Rechtsklick = Einfuegen) und Enter
4. Der Lauf rendert pro eindeutiger Folie ein PNG via PowerPoint-Automation. Dauer: ca. 5–15 Sekunden pro Folie
5. Abbrechbar mit `Strg+C`; beim naechsten Start macht der Lauf an der Stelle weiter

---

## 6. Was die App **nicht** anfasst

- **PPTX-Dateien werden NICHT veraendert.** In Phase 1 wird ausschliesslich gelesen — die Software baut ihre eigene Datenbank in `data\pptx-manager.db` auf
- Keine Internet-Verbindung im Betrieb noetig. Der Server laeuft nur auf `127.0.0.1` (= dein eigener Computer, nicht im Netzwerk erreichbar)
- Die Ordner `data\` und `data\thumbnails\` sind die einzigen Stellen, an denen die App schreibt. Loescht man diese, ist die Indizierung weg — die PPTX-Dateien bleiben

---

## 7. Hauptpfade verwalten

Auf der Scan-Seite, unten in der Sektion **"Index-Verwaltung — gespeicherte Hauptpfade"**, kannst du:

- Mehrere Hauptpfade pflegen (z.B. einen pro Kunde oder Themengebiet)
- Labels frei vergeben (Klick auf den Label-Text in der Liste, tippen, Enter speichert)
- Einzelne Hauptpfade neu scannen (Button **↻ Aktualisieren**) — erkennt verschwundene Dateien und bietet Loeschung an
- Hauptpfade komplett aus der DB werfen (Button **✕ Loeschen** mit Bestaetigungsdialog)
- Alles auf einmal aktualisieren (Button **"Alle Pfade aktualisieren"**)
- Index komplett zuruecksetzen (Button **"Index komplett zuruecksetzen"** ganz rechts unten)

Auf der **Suchseite** (Hauptseite) tauchen die Hauptpfade als anklickbare Filter-Pills oben auf — so kannst du die Suche auf bestimmte Pfade einschraenken.

---

## 8. Updates einspielen

Da du ohne Git arbeitest, geht ein Update so:

1. Erneut https://github.com/Dirk10431/pptx-manager → Code → Download ZIP
2. **WICHTIG:** Vorher deinen `data\`-Ordner **kopieren** und beiseite legen — sonst ist die Indizierung weg
3. Den alten Projekt-Ordner umbenennen (z.B. `pptx-manager_alt`) oder loeschen
4. Neue ZIP entpacken
5. Den gesicherten `data\`-Ordner in den neuen Projekt-Ordner kopieren
6. Doppelklick auf `start.bat` — `node_modules\` wird neu aufgebaut, die Datenbank bleibt erhalten

---

## 9. Troubleshooting

### "Port 3002 already in use" / "EADDRINUSE"
Eine fruehere Instanz laeuft noch. PowerShell als Admin starten und:
```powershell
Get-Process node -ErrorAction SilentlyContinue | Stop-Process -Force
```
Dann `start.bat` erneut.

### "npm install" schlaegt fehl
- Internet-Verbindung pruefen
- Hinter Firewall/Proxy? `npm config set proxy <URL>` setzen (selten noetig)
- Antivirus (Norton, BitDefender, Avast) blockiert die Installation oder einzelne Dateien. **Quarantaene pruefen** und einen Ausschluss fuer den Projekt-Ordner ergaenzen

### Thumbnails: "PowerPoint-Application konnte nicht gestartet werden"
- Microsoft Office / PowerPoint installiert? Mindestens 2016 wird empfohlen
- PowerPoint einmal manuell oeffnen + schliessen (initialisiert die Office-Registrierung)
- Antivirus blockiert evtl. die PowerShell-Skripte unter `data\ps-scripts\`. Diesen Ordner in den AV-Ausschluessen ergaenzen

### Browser oeffnet sich nicht automatisch
Manuell oeffnen: http://127.0.0.1:3002

### "Windows protected your PC" beim Doppelklick auf start.bat
Klicke **"Weitere Informationen"** → **"Trotzdem ausfuehren"**. Das ist die normale SmartScreen-Warnung fuer aus dem Internet geladene Skripte.

---

## 10. Projekt-Spezifika

- **Lokal und offline** — keine Internet-Verbindung im Betrieb noetig
- **Port 3002** ist fest verdrahtet (nicht 3000/3001 wie andere Tools)
- **Nur 127.0.0.1** — der Server ist nicht im Netzwerk erreichbar (Sicherheits-Setting)
- **PPTX-Dateien werden in Phase 1 nur gelesen** — Schreibzugriffe kommen erst spaeter mit Backup + Dry-Run

---

**Fertig.** Bei Problemen: melde dich bei Dirk Peters (dr.dirk.peters@gmail.com).
