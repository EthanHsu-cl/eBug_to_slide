#!/bin/bash
# Build eBug_to_slide.app for macOS
set -e

source venv/bin/activate
pip install pyinstaller --quiet
pyinstaller eBug_to_slide.spec --clean

echo ""
echo "Done. Output: dist/eBug_to_slide.app"
