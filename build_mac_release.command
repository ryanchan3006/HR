#!/bin/bash
set -euo pipefail

cd "$(dirname "$0")"

APP_NAME="ContractGenerator"
APP_PATH="dist/${APP_NAME}.app"
ZIP_PATH="dist/${APP_NAME}-mac.zip"

if [[ -z "${MAC_SIGN_IDENTITY:-}" ]]; then
  echo "ERROR: MAC_SIGN_IDENTITY is not set."
  echo "Example:"
  echo '  export MAC_SIGN_IDENTITY="Developer ID Application: Your Name (TEAMID)"'
  exit 1
fi

if [[ -z "${APPLE_ID:-}" || -z "${APPLE_TEAM_ID:-}" || -z "${APPLE_APP_PASSWORD:-}" ]]; then
  echo "ERROR: APPLE_ID, APPLE_TEAM_ID, and APPLE_APP_PASSWORD must be set."
  echo "Example:"
  echo '  export APPLE_ID="you@example.com"'
  echo '  export APPLE_TEAM_ID="ABCDE12345"'
  echo '  export APPLE_APP_PASSWORD="xxxx-xxxx-xxxx-xxxx"'
  exit 1
fi

echo "============================================"
echo " Contract Generator -- Build macOS Release"
echo "============================================"
echo
echo "[1/7] Installing dependencies..."
python3 -m pip install -r requirements.txt pyinstaller

echo
echo "[2/7] Building app bundle..."
pyinstaller \
  --noconfirm \
  --clean \
  --windowed \
  --name "${APP_NAME}" \
  --add-data "Workflow Automation template.docx:." \
  app.py

echo
echo "[3/7] Code-signing app..."
codesign --force --deep --options runtime --timestamp \
  --sign "${MAC_SIGN_IDENTITY}" \
  "${APP_PATH}"

echo
echo "[4/7] Verifying signature..."
codesign --verify --deep --strict --verbose=2 "${APP_PATH}"
spctl --assess --type execute --verbose=2 "${APP_PATH}" || true

echo
echo "[5/7] Creating notarization zip..."
rm -f "${ZIP_PATH}"
ditto -c -k --sequesterRsrc --keepParent "${APP_PATH}" "${ZIP_PATH}"

echo
echo "[6/7] Submitting for notarization..."
xcrun notarytool submit "${ZIP_PATH}" \
  --apple-id "${APPLE_ID}" \
  --team-id "${APPLE_TEAM_ID}" \
  --password "${APPLE_APP_PASSWORD}" \
  --wait

echo
echo "[7/7] Stapling notarization ticket and refreshing zip..."
xcrun stapler staple "${APP_PATH}"
rm -f "${ZIP_PATH}"
ditto -c -k --sequesterRsrc --keepParent "${APP_PATH}" "${ZIP_PATH}"

echo
echo "Done!"
echo
echo "Signed app:        ${APP_PATH}"
echo "Notarized zip:     ${ZIP_PATH}"
echo
echo "This build is ready to move to another Mac with far fewer Gatekeeper issues."
echo
