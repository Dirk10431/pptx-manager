# pptx-manager auf dem Mac installieren

Kurzanleitung — nur die Befehle. Fuer einen frischen Mac mit Internet-Verbindung.

> Mac-Passwort wird beim Homebrew-Schritt einmal abgefragt. Beim ersten Klick auf "Ordner waehlen" fragt macOS nach Berechtigung → **Erlauben**.

---

## 1. Terminal oeffnen

Cmd + Leertaste → "Terminal" eintippen → Enter.

## 2. Apple Command Line Tools

```bash
xcode-select --install
```

Dialog erscheint → **Installieren** → 5–10 Min warten.

## 3. Homebrew installieren

```bash
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
```

Nach der Installation **die zwei Zeilen aus dem Homebrew-Output ausfuehren** (Anpassung des PATH). Auf Apple Silicon sieht das so aus:

```bash
echo >> ~/.zprofile
echo 'eval "$(/opt/homebrew/bin/brew shellenv)"' >> ~/.zprofile
eval "$(/opt/homebrew/bin/brew shellenv)"
```

## 4. Node.js und LibreOffice

```bash
brew install node
brew install --cask libreoffice
```

## 5. Repository klonen

```bash
mkdir -p ~/Projekte
cd ~/Projekte
git clone https://github.com/Dirk10431/pptx-manager.git
cd pptx-manager
```

## 6. Abhaengigkeiten installieren

```bash
npm install
```

## 7. Server starten

```bash
chmod +x start.sh
./start.sh
```

Browser oeffnet automatisch [http://127.0.0.1:3002](http://127.0.0.1:3002).
Beenden mit **Ctrl + C** im Terminal.

---

## Nutzung

**Scan starten** (im Browser):
Scan-Seite → Ordner waehlen → Scan starten.

**Vorschaubilder erzeugen** (im Terminal):

```bash
cd ~/Projekte/pptx-manager
npm run thumbs
```

Erstmal nur 5 Folien testen:

```bash
npm run thumbs -- --limit 5
```

---

## Updates ziehen

```bash
cd ~/Projekte/pptx-manager
git pull
npm install
```
