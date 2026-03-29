# Deployment Guide

## What needs to be deployed

Two components:

1. **Python scripts** — installed inside the Collabora server at
   `/opt/collaboraoffice/share/Scripts/python/`
   - `zotero_fields.py` — core field management (required)
   - `zotero_export.py` — citation export (optional)

2. **Browser extension** — installed in each user's Chrome/Chromium browser.
   Load unpacked from the `chromium-extension/` directory, or package as a `.crx`.

## 1. Docker Collabora (standalone)

This is the recommended deployment for both standalone setups and Nextcloud.

### Option A: Bind mount (simplest)

Add a volume mount to your `docker-compose.yml` or `docker run` command:

```yaml
services:
  collabora:
    image: collabora/code:latest
    environment:
      - extra_params=--o:security.capabilities=false --o:security.seccomp=false
    volumes:
      - ./scripts/zotero_fields.py:/opt/collaboraoffice/share/Scripts/python/zotero_fields.py:ro
      - ./scripts/zotero_export.py:/opt/collaboraoffice/share/Scripts/python/zotero_export.py:ro
    ports:
      - "9980:9980"
```

That's it. No custom image, no rebuild. Changes to the scripts take effect on
the next document open (or container restart).

### Option B: Custom image layer (more portable)

Create a one-line Dockerfile:

```dockerfile
FROM collabora/code:latest
COPY scripts/zotero_fields.py scripts/zotero_export.py \
     /opt/collaboraoffice/share/Scripts/python/
```

Build and run:

```bash
docker build -t collabora-zotero .
docker run -p 9980:9980 collabora-zotero
```

This is better for production — the image is self-contained and can be pushed
to a registry.

### Collabora configuration

Collabora must have Python macro execution enabled. The official
`collabora/code` Docker image includes this by default. Verify by checking
that the container has `/opt/collaboraoffice/program/python-core-*` present:

```bash
docker exec collabora ls /opt/collaboraoffice/program/ | grep python
```

## 2. Nextcloud with Collabora

### Recommended: External Collabora Docker (richdocuments app)

Nextcloud's **Collabora Online** integration app (`richdocuments`) connects to
an external Collabora server. Deploy Collabora as a Docker container using the
instructions above, then point Nextcloud at it:

1. Deploy Collabora Docker with the Python scripts (Option A or B above).
2. In Nextcloud, install the **Collabora Online** app (`richdocuments`).
3. Go to **Settings → Administration → Collabora Online** and set the URL to
   your Collabora server (e.g., `https://collabora.example.com`).
4. Install the browser extension for each user.

### Not recommended: Built-in CODE (richdocumentscode)

The **Nextcloud Office** built-in CODE server (`richdocumentscode`) is a
stripped-down Collabora package compiled into a Nextcloud app. It typically
does **not** include Python scripting support, which means
`CallPythonScript` PostMessages will fail silently.

Even if you manually copy scripts into the CODE app's directory structure,
the Python runtime is usually absent. There is no workaround short of
recompiling CODE with Python support.

**Use the external Docker container instead.**

## WOPI host requirements

Your WOPI server's `CheckFileInfo` response must include:

```json
{
  "PostMessageOrigin": "https://your-wopi-host.example.com"
}
```

This allows the Collabora iframe to send PostMessages back to the host page,
which the browser extension intercepts.

## Browser extension

The extension requires:

- **Zotero desktop** running on the same machine (listens on `127.0.0.1:23119`)
- **Chrome or Chromium-based browser** (Manifest V3)

### Install (development)

1. Open `chrome://extensions/`
2. Enable **Developer mode**
3. Click **Load unpacked** and select the `chromium-extension/` directory

### Install (production)

Package as a `.crx` file or distribute via Chrome Web Store / enterprise
policy. The extension needs `host_permissions` for `127.0.0.1:23119` (Zotero)
and the WOPI host origin.
