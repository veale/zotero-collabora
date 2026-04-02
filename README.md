# Cite with Local Zotero in Collabora

A browser extension and server-side plugin that brings **full Zotero desktop integration** to Collabora Online. Insert, edit, and manage citations and bibliographies in your web browser using the same Zotero connector protocol as the native LibreOffice plugin — producing documents that are fully interchangeable between Collabora Online and LibreOffice desktop, so you can work offline and online sa you wish.

Download as a [Firefox Add On](https://addons.mozilla.org/en-GB/firefox/addon/cite-with-zotero-collabora/) or a Chromium Add On. Currently requires two scripts to be placed server side too, see below. 

![2](https://github.com/user-attachments/assets/44255c78-282a-43ed-bd08-543232621d57)

Think how Google Docs integration works via the Zotero Connector, but for Collabora Online.

## Key features

- **Full Zotero connector protocol** — uses the same HTTP transaction API as the native LibreOffice plugin, not the limited Zotero Web API, which is rate-limited and can struggle with certain types of Zotero libraries
- **All citation operations** — insert, edit, delete, refresh, change citation style, convert between in-text and footnote/endnote
- **Bibliography management** — insert, update, and format bibliographies with proper per-entry paragraph styling
- **Native document format** — citations stored as LibreOffice ReferenceMark fields and text sections, fully compatible with the desktop Zotero plugin. Can be swapped to bookmarks just as with the LibreOffice plugin if you need to shift to other word processors like Microsoft Word
- **Round-trip editing** — edit the same `.odt` file in Collabora and LibreOffice desktop interchangeably, with all Zotero fields preserved
- **All citation styles** — any style available in Zotero works, including numbered styles (IEEE, Vancouver) and note-based styles (OSCOLA, Chicago)
- **Citation export** — a useful feature that in one click lets you export all document citations in formats including JSON directly from the document, without needing Zotero running
- **Toolbar buttons** — injected directly into the Collabora editor toolbar for one-click access (can be disabled if you wish)
- **Extension popup menu** — all operations also always accessible from the browser extension icon, in case you do not want them cluttering your Collabora interface.
- **Works with any WOPI host** — Nextcloud, ownCloud, Seafile, or any custom WOPI server

## How it compares to Nextcloud richdocuments' built-in Zotero

Nextcloud's [richdocuments](https://github.com/nextcloud/richdocuments) app includes a built-in Zotero integration. It uses the **Zotero Web API** with an API key to fetch citations from your online library. This project takes a fundamentally different approach: it talks directly to **Zotero desktop** running on your machine via the HTTP connector protocol — the same protocol the native LibreOffice plugin uses.

| Feature | This project | Nextcloud richdocuments |
|---|---|---|
| **Connection** | Zotero desktop (local HTTP connector) | Zotero Web API (cloud) |
| **Zotero desktop required** | Yes (running locally) | No (API key only) |
| **Internet required** | No | Yes |
| **Citation picker** | Native Zotero picker dialog | Custom web-based search UI |
| **Citation styles** | All styles, applied by Zotero engine | All styles, applied by Collabora's built-in CSL processor |
| **Style switching** | Full (with footnote conversion) | Limited (no footnote/endnote conversion) |
| **Document storage** | ReferenceMark + text sections | ReferenceMark, Field, or Bookmark |
| **Round-trip with LibreOffice desktop** | Full compatibility (identical format, Zotero renders the text) | Compatible (same field format; text is rendered by Collabora's CSL processor rather than Zotero's, so minor formatting differences are possible) |
| **Bibliography** | Stored as text:section with per-entry paragraph styling (matches LibreOffice) | Stored as field/reference mark (single block) |
| **Numbered styles (IEEE, Vancouver)** | Correct tab-aligned formatting | Correct formatting |
| **Note-based styles (OSCOLA, Chicago)** | Full support with footnote conversion | No automatic footnote conversion |
| **Edit existing Zotero fields** | Yes (reads existing fields, full edit cycle) | Yes |
| **Citation export** | CSL-JSON, BibTeX, RIS (no Zotero needed) | No |
| **Multi-user editing** | Document-level locking during transactions | No locking |
| **Sync with Zotero library** | Instant (local desktop) | Delayed (API sync) |
| **Works outside Nextcloud** | Yes (any WOPI host) | No (Nextcloud only, requires every user to have registered accounts, so no guests) |
| **Group libraries** | Yes (via Zotero desktop) | Yes (via API) |

### When to use which

**Use this project** if you:
- Already use Zotero desktop and want the same experience in the browser
- Need perfect round-tripping between Collabora and LibreOffice desktop
- Use note-based citation styles (OSCOLA, Chicago, etc.) that require footnotes
- Want to work offline or without sharing your Zotero API key with a server
- Use a WOPI host other than Nextcloud
- Have a large Zotero library

**Use the built-in richdocuments integration** if you:
- Don't have Zotero desktop installed
- Want a cloud-only workflow with no local software
- Only use author-date citation styles (Harvard, APA, etc.)
- Are on Nextcloud and want zero-install setup
- Have a small Zotero library

## Installation

### Prerequisites

- **Zotero desktop** (7.x or later) running on the same machine as your browser
- **Collabora Online** server (Docker recommended)
- **Chrome**, **Chromium**, or **Firefox** browser

### 1. Server setup (Collabora administrator)

The extension works by calling Python scripts that run inside Collabora's LibreOffice Kit process. These scripts use the LibreOffice UNO API to manipulate reference marks, text sections, footnotes, and document properties — operations that are not accessible from the browser. There is no way to avoid this server-side component; the browser extension cannot function without the scripts installed.

Two things need to be configured:

1. **Install the Python scripts** into Collabora's trusted script directory (`/opt/collaboraoffice/share/Scripts/python/`)
2. **Enable macro execution** so that LibreOffice Kit will run the scripts when called

#### Macro security levels

Collabora (like LibreOffice) has a macro security setting that controls which scripts are allowed to run. The default is level 4 (all macros disabled). We recommend **level 3**:

| Level | Name | Behaviour | Recommendation |
|---|---|---|---|
| 0 | Very Low | All macros run without any checks | Not recommended — would also run macros embedded in uploaded documents |
| 1 | Low | Trusted locations run silently; untrusted macros show a confirmation dialog | Not recommended — the dialog hangs Collabora's headless Kit process |
| 2 | Medium | Trusted locations run; everything else is silently blocked | Works, but level 3 is stricter |
| **3** | **High** | **Signed macros and trusted filesystem locations only** | **Recommended — most restrictive level that works** |
| 4 | Very High | All macros disabled | Default — will not work |

Scripts installed in `/opt/collaboraoffice/share/Scripts/python/` are in a trusted filesystem location, so they run at levels 0–3 regardless of signing. Level 3 is the most secure option: our scripts execute normally, while unsigned macros embedded in user-uploaded documents are blocked.

See the [Collabora Online SDK documentation](https://sdk.collaboraonline.com/docs/installation/Configuration.html) for more on `coolwsd.xml` configuration.

#### Option A: Bind mount (recommended)

Add volume mounts and environment variables to your existing Collabora `docker-compose.yml`:

```yaml
services:
  collabora:
    image: collabora/code:latest
    volumes:
      - ./scripts/zotero_fields.py:/opt/collaboraoffice/share/Scripts/python/zotero_fields.py:ro
      - ./scripts/zotero_export.py:/opt/collaboraoffice/share/Scripts/python/zotero_export.py:ro
    environment:
      - extra_params=--o:security.enable_macros_execution=true --o:security.macro_security_level=3
    ports:
      - "9980:9980"
```

This is the simplest approach — it layers the scripts on top of the official image, so you get Collabora updates automatically when you pull `latest`. Changes to the scripts take effect on the next document open.

> **Note:** The base `collabora/code` image may not include the Python scripting packages. If `CallPythonScript` fails silently, use Option B to install them.

#### Option B: Custom Docker image

If the base image is missing the Python scripting packages, or you want a fully self-contained image:

```dockerfile
FROM collabora/code:latest

USER root
RUN apt-get update && apt-get install -y --no-install-recommends \
        collaboraofficebasis-python-script-provider \
        collaboraofficebasis-pyuno \
    && rm -rf /var/lib/apt/lists/*

# Enable Python macro execution (level 3 = signed + trusted locations)
RUN COOLWSD=/etc/coolwsd/coolwsd.xml && \
    sed -i 's|<enable_macros_execution[^>]*>[^<]*</enable_macros_execution>|<enable_macros_execution desc="" type="bool" default="false">true</enable_macros_execution>|' "$COOLWSD" && \
    sed -i 's|<macro_security_level[^>]*>[^<]*</macro_security_level>|<macro_security_level desc="" type="int" default="4">3</macro_security_level>|' "$COOLWSD"

COPY scripts/zotero_fields.py scripts/zotero_export.py \
     /opt/collaboraoffice/share/Scripts/python/

USER cool
```

Build and run:

```bash
docker build -t collabora-zotero .
docker run -p 9980:9980 \
  -e "aliasgroup1=https://your-wopi-host.example.com" \
  collabora-zotero
```

> **Note:** You will need to rebuild this image whenever Collabora releases an update, so Option A is generally preferred.

#### WOPI host requirement

Your WOPI server's `CheckFileInfo` response must include `PostMessageOrigin` set to the origin of the host page. This allows the Collabora iframe to communicate with the extension via PostMessage.

#### Nextcloud

If using Nextcloud with the **Collabora Online** app (`richdocuments`), deploy Collabora as a separate Docker container using the instructions above, then point Nextcloud at it via **Settings > Administration > Collabora Online**.

The built-in CODE server (`richdocumentscode` app) does **not** include Python scripting support and will not work.

#### Verifying the server setup

```bash
# Check Python packages are installed
docker exec collabora ls /opt/collaboraoffice/program/ | grep python

# Check scripts are in place
docker exec collabora ls /opt/collaboraoffice/share/Scripts/python/zotero_*.py

# Check for macro errors in logs
docker logs collabora 2>&1 | grep -i macro
```

### 2. Browser extension (end users)

The extension will soon be available in the Chrome Web Store and Firefox Add-ons under the name **Cite with Local Zotero in Collabora**. Until then, install manually:

#### Chrome / Chromium / Edge

1. Download or clone this repository
2. Open `chrome://extensions/`
3. Enable **Developer mode** (toggle in the top right)
4. Click **Load unpacked** and select the `chromium-extension/` directory
5. The Zotero icon appears in your toolbar

#### Firefox

1. Run `./build-firefox.sh` to generate the Firefox extension (or use the pre-built `firefox-extension/` directory)
2. Open `about:debugging#/runtime/this-firefox`
3. Click **Load Temporary Add-on** and select any file inside `firefox-extension/`

> **Note:** Temporary add-ons are removed when Firefox restarts. For persistent installation, the extension needs to be signed via [addons.mozilla.org](https://addons.mozilla.org).

#### Using the extension

1. Make sure **Zotero desktop** is running
2. Open a document in Collabora Online
3. Use the toolbar buttons that appear in the Collabora editor, or click the extension icon for the popup menu
4. The first time you insert a citation, Zotero will open its citation picker dialog

## Architecture

```
Browser                      Local machine               Collabora server
┌─────────────────────┐        ┌──────────────────┐        ┌─────────────────────┐
│ Collabora iframe    │        │ Browser          │        │ LibreOffice Kit     │
│ (editor UI)         │◄──PM──►│ Extension        │        │ ┌─────────────────┐ │
│                     │        │                  │        │ │ zotero_fields   │ │
│ ┌─────────────────┐ │        │ background.js    │◄─HTTP─►│ │ .py             │ │
│ │ Toolbar buttons │ │        │ content.js       │   PM   │ │                 │ │
│ │ (injected)      │ │        │                  │        │ │ zotero_export   │ │
│ └─────────────────┘ │        └────────┬─────────┘        │ │ .py             │ │
└─────────────────────┘                 │                  │ └─────────────────┘ │
                                        │ HTTP             └─────────────────────┘
                                ┌───────▼────────┐
                                │ Zotero Desktop │
                                │ :23119         │
                                └────────────────┘
```

**PM** = PostMessage API between the host page and the Collabora iframe.

The extension acts as a bridge between Zotero desktop and Collabora's LibreOffice Kit process:

1. **Zotero desktop** runs the citation engine and drives the transaction (insert field, set code, set text, etc.)
2. **background.js** speaks Zotero's HTTP connector protocol, translating each command into a Python function call
3. **content.js** relays Python calls to the Collabora iframe via PostMessage, and injects toolbar buttons
4. **zotero_fields.py** executes inside LibreOffice Kit, manipulating ReferenceMark objects, text sections, and document properties using the UNO API

### Document storage format

Citations are stored in the same format as the native Zotero LibreOffice plugin:

| Element | Storage |
|---|---|
| Citation fields | `text:reference-mark` named `ZOTERO_<CSL_CITATION JSON> RND<13-char ID>` |
| Bibliography | `text:section` named `ZOTERO_BIBL <JSON> CSL_BIBLIOGRAPHY RND<ID>`, with one paragraph per entry using the `Bibliography 1` style |
| Document preferences | Custom document properties `ZOTERO_PREF_1`, `_2`, etc. (255-char chunks) |
| Field type | Stored as `ReferenceMark` in the document, sent as `Http` to Zotero's HTTP connector |

This means a document edited in Collabora with this extension can be opened in LibreOffice desktop with the Zotero plugin, and vice versa, with all citations and bibliography intact and editable.

## Supported operations

| Operation | Toolbar | Popup | Description |
|---|---|---|---|
| Add/Edit Citation | Yes | Yes | Opens the Zotero citation picker to insert or edit a citation |
| Bibliography | Yes | Yes | Insert or refresh the bibliography |
| Refresh | Yes | Yes | Refresh all citations and bibliography |
| Document Preferences | Yes | Yes | Change citation style, language, and other settings |
| Add Note | Yes | Yes | Insert a Zotero note |
| Unlink Citations | Yes | Yes | Remove Zotero field codes, leaving plain text |
| Export Citations | Yes | Yes | Export all cited items as CSL-JSON (no Zotero needed) |
| Force Cancel | — | Yes | Abort a stuck Zotero transaction |
| Toggle Toolbar | — | Yes | Show/hide the toolbar buttons in the editor |

## Development

### Local setup

```bash
# Start Collabora and the development WOPI server
docker compose up -d --build

# Open http://localhost:8080 to see available test documents
# Load the extension in your browser (see installation above)
```

The development WOPI server (`wopi/server.py`) serves `.odt` files from `wopi/docs/` and generates toolbar button icons dynamically.

### Project structure

```
chromium-extension/  Chromium extension (Manifest V3)
  background.js      Zotero HTTP connector transaction loop
  content.js         Toolbar injection, Python call relay, dialog UI
  popup.html/js      Extension popup menu

firefox-extension/   Firefox extension (generated by build-firefox.sh)

scripts/
  zotero_fields.py   Core field management (UNO API)
  zotero_export.py   Citation export (CSL-JSON, BibTeX, RIS)

wopi/
  server.py          Development WOPI server (Flask)
  docs/              Test documents
```

## Licence

This is an unofficial add-on for both Zotero and Collabora. Trademarked names are used on the basis of nominative fair use.

Mozilla Public License Version 2.0
