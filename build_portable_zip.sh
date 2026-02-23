#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

RELEASE_DIR="$SCRIPT_DIR/release"
STAGE_DIR="$RELEASE_DIR/PrinterKeyCheckoutTracker"
ZIP_PATH="$RELEASE_DIR/PrinterKeyCheckoutTracker-python-portable.zip"

rm -rf "$STAGE_DIR"
mkdir -p "$STAGE_DIR/resources"

cp app.py "$STAGE_DIR/"
cp storage.py "$STAGE_DIR/"
cp models.py "$STAGE_DIR/"
cp utils.py "$STAGE_DIR/"
cp requirements.txt "$STAGE_DIR/"
cp README.md "$STAGE_DIR/"
cp run_mac_linux.sh "$STAGE_DIR/"
cp run_windows.bat "$STAGE_DIR/"
cp resources/AppIcon-256.png "$STAGE_DIR/resources/"

rm -f "$ZIP_PATH"
(
  cd "$RELEASE_DIR"
  zip -r "$(basename "$ZIP_PATH")" "$(basename "$STAGE_DIR")" >/dev/null
)

echo "Portable bundle created: $ZIP_PATH"
