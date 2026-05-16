# pptx-manager auf Windows installieren

Kurzanleitung — nur die Befehle. Fuer einen frischen Windows-11-PC mit Internet-Verbindung.

> Bei jeder `winget`-Installation einmal **UAC-Dialog bestaetigen**. **PowerPoint** muss installiert sein, sonst werden keine Folien-Vorschaubilder erzeugt — Suche und Duplikat-Erkennung funktionieren auch ohne.

---

## 1. PowerShell oeffnen

Windows-Taste → "PowerShell" eintippen → Enter.

## 2. Node.js und Git installieren

```powershell
winget install OpenJS.NodeJS.LTS
winget install Git.Git
```

**PowerShell danach komplett schliessen und neu oeffnen** — sonst sieht sie die neuen Tools nicht im PATH.

## 3. Repository klonen

```powershell
mkdir C:\Tools -ErrorAction SilentlyContinue
cd C:\Tools
git clone https://github.com/Dirk10431/pptx-manager.git
cd pptx-manager
```

## 4. Abhaengigkeiten installieren

```powershell
npm install
```

## 5. Server starten

```powershell
.\start.bat
```

Browser oeffnet automatisch [http://127.0.0.1:3002](http://127.0.0.1:3002).
Beenden mit **Strg + C** im Konsolen-Fenster (zweimal druecken).

---

## Nutzung

**Scan starten** (im Browser):
Scan-Seite → "Ordner waehlen…" → Scan starten.

**Vorschaubilder erzeugen** (im zweiten PowerShell-Fenster):

```powershell
cd C:\Tools\pptx-manager
npm run thumbs
```

Erstmal nur 5 Folien testen:

```powershell
npm run thumbs -- --limit 5
```

---

## Desktop-Shortcuts (optional)

Erstellt `pptx-open.bat` (Server + Browser) und `pptx-thumbs.bat` (Vorschaubilder) mit hardcoded Pfad. Beide Dateien koennen anschliessend auf den Desktop gezogen werden.

```powershell
cd C:\Tools\pptx-manager
.\make-shortcuts.bat
```

---

## Updates ziehen

```powershell
cd C:\Tools\pptx-manager
git pull
npm install
```
