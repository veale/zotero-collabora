#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"

mkdir -p built-extensions

CV=$(python3 -c "import json; print(json.load(open('chromium-extension/manifest.json'))['version'])")
FV=$(python3 -c "import json; print(json.load(open('firefox-extension/manifest.json'))['version'])")

(cd chromium-extension && zip -r "../built-extensions/chromium-extension-v${CV}.zip" . -x '.*')
(cd firefox-extension && zip -r "../built-extensions/firefox-extension-v${FV}.zip" . -x '.*')

echo "Built: built-extensions/chromium-extension-v${CV}.zip"
echo "Built: built-extensions/firefox-extension-v${FV}.zip"
