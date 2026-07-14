#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

if ! command -v python3 >/dev/null 2>&1; then
    osascript -e 'display dialog "Python 3 was not found. Please install Python 3 from python.org or install Anaconda." buttons {"OK"} with icon stop with title "FAI IPQC Excel Extractor Build"'
    exit 1
fi

python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements-mac.txt

rm -rf build dist release
pyinstaller --clean --noconfirm FAI_IPQC_Excel_Extractor.spec
file "dist/FAI IPQC Excel Extractor.app/Contents/MacOS/FAI IPQC Excel Extractor" || true
if command -v lipo >/dev/null 2>&1; then
    lipo -archs "dist/FAI IPQC Excel Extractor.app/Contents/MacOS/FAI IPQC Excel Extractor" || true
fi
FAI_EXTRACTOR_SMOKE_TEST=1 "dist/FAI IPQC Excel Extractor.app/Contents/MacOS/FAI IPQC Excel Extractor"

mkdir -p release
cp -R "dist/FAI IPQC Excel Extractor.app" release/
cd release
zip -qr "FAI_IPQC_Excel_Extractor_macOS_Standalone.zip" "FAI IPQC Excel Extractor.app"
cd ..

osascript -e 'display dialog "Build complete: mac_package/release/FAI IPQC Excel Extractor.app" buttons {"OK"} with icon note with title "FAI IPQC Excel Extractor Build Complete"' || true

echo "Build complete: $SCRIPT_DIR/release/FAI IPQC Excel Extractor.app"
echo "Zip package: $SCRIPT_DIR/release/FAI_IPQC_Excel_Extractor_macOS_Standalone.zip"
