#!/bin/bash
# build-firefox.sh — Build the Firefox extension from the Chrome source.
#
# Copies chromium-extension/ to firefox-extension/, patches the manifest for Firefox
# MV3, and strips ES module syntax (Firefox MV3 background scripts don't
# support "type": "module").
#
# Output: firefox-extension/ directory (load as temporary add-on in about:debugging)

set -euo pipefail

SRC="chromium-extension"
DEST="firefox-extension"

rm -rf "$DEST"
cp -r "$SRC" "$DEST"

# Patch manifest.json for Firefox
python3 -c "
import json, sys

with open('$DEST/manifest.json') as f:
    m = json.load(f)

# Firefox needs a gecko add-on ID
m['browser_specific_settings'] = {
    'gecko': {
        'id': 'zotero-collabora@example.com',
        'strict_min_version': '115.0'
    }
}

# Firefox MV3 uses 'scripts' array, not 'service_worker', and no 'type' field
m['background'] = {
    'scripts': ['background.js']
}

with open('$DEST/manifest.json', 'w') as f:
    json.dump(m, f, indent=2)
    f.write('\n')
"

# Strip 'export' keywords from background.js (Firefox background scripts
# are not ES modules)
sed -i.bak 's/^export //g' "$DEST/background.js"
rm -f "$DEST/background.js.bak"

echo "Firefox extension built in $DEST/"
echo "Load it in Firefox: about:debugging → This Firefox → Load Temporary Add-on → select $DEST/manifest.json"
