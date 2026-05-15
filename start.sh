#!/bin/bash
# =============================================================
# pptx-manager - Start-Skript fuer Mac/Linux
# =============================================================

cd "$(dirname "$0")"

echo ""
echo " ========================================="
echo "  pptx-manager wird gestartet..."
echo " ========================================="
echo ""

# Abhaengigkeiten pruefen
if [ ! -d "node_modules" ]; then
    echo "Erstinstallation: npm install laeuft..."
    npm install || {
        echo ""
        echo "FEHLER: npm install fehlgeschlagen."
        exit 1
    }
fi

# Browser oeffnen nach 2 Sekunden
(sleep 2 && open http://127.0.0.1:3002 2>/dev/null || xdg-open http://127.0.0.1:3002 2>/dev/null) &

# Server starten
node server.js
