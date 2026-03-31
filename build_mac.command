#!/bin/bash
set -e

cd "$(dirname "$0")"

echo "============================================"
echo " Contract Generator -- Build macOS App"
echo "============================================"
echo
echo "[1/4] Installing dependencies..."
python3 -m pip install -r requirements.txt pyinstaller

echo
echo "[2/4] Building app bundle..."
pyinstaller \
  --noconfirm \
  --clean \
  --windowed \
  --name "ContractGenerator" \
  --add-data "Workflow Automation template.docx:." \
  app.py

echo
echo "[3/4] Creating shareable zip..."
rm -f "dist/ContractGenerator-mac.zip"
ditto -c -k --sequesterRsrc --keepParent "dist/ContractGenerator.app" "dist/ContractGenerator-mac.zip"

echo
echo "[4/4] Done!"
echo
echo "Your macOS app is at:  dist/ContractGenerator.app"
echo "Shareable zip is at:   dist/ContractGenerator-mac.zip"
echo
echo "The app bundle includes Python dependencies and the default template."
echo "DOCX generation/export works without Python on the end user's Mac."
echo "PDF export on macOS still requires LibreOffice with the current app."
echo
echo "If the app is unsigned, macOS may require right-click > Open on first launch."
echo "For frictionless distribution to other Macs, code signing and notarization are recommended."
echo
read -n 1 -s -r -p "Press any key to close..."
echo
