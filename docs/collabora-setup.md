# Installing Zotero scripts in a Collabora Online Docker container

## What you need

Two Python files from the `scripts/` directory:

- `zotero_fields.py` — citation and bibliography field management (required)
- `zotero_export.py` — standalone citation export (optional)

These go into Collabora's global Python script directory:
`/opt/collaboraoffice/share/Scripts/python/`

## Three things must be true

### 1. Python scripting packages are installed

The official `collabora/code` image does **not** include Python scripting
support by default. Two packages are needed:

```
collaboraofficebasis-python-script-provider
collaboraofficebasis-pyuno
```

Check if they're present:

```bash
docker exec <container> dpkg -l | grep -i pyuno
```

If missing, install them in a custom image (see below).

### 2. Macro execution is enabled

Collabora's `coolwsd.xml` has two settings that block script execution by
default:

| Setting                    | Default | Required |
|----------------------------|---------|----------|
| `enable_macros_execution`  | `false` | `true`   |
| `macro_security_level`     | `4`     | `3`      |

Security level `3` (High) is the most restrictive level that works. Scripts
in trusted filesystem locations (the share directory where ours are installed)
run normally, while unsigned macros embedded in documents are blocked.

These can be set via `sed` in a Dockerfile or passed as `extra_params`:

```
--o:security.enable_macros_execution=true
--o:security.macro_security_level=3
```

### 3. The script files are in the right directory

Scripts must be at `/opt/collaboraoffice/share/Scripts/python/`. This is the
"share" location that LibreOffice Kit scans for Python macros available to all
documents.

## Option A: Bind mount (simplest, good for development)

No custom image needed. Mount the scripts as read-only volumes:

```yaml
services:
  collabora:
    image: collabora/code:latest
    volumes:
      - ./scripts/zotero_fields.py:/opt/collaboraoffice/share/Scripts/python/zotero_fields.py:ro
      - ./scripts/zotero_export.py:/opt/collaboraoffice/share/Scripts/python/zotero_export.py:ro
    environment:
      - extra_params=--o:security.enable_macros_execution=true --o:security.macro_security_level=3
```

**Caveat:** this assumes the Python packages are already installed in the base
image. If `CallPythonScript` fails silently, you need Option B.

## Option B: Custom Dockerfile (recommended for production)

```dockerfile
FROM collabora/code:latest

USER root
RUN apt-get update && apt-get install -y --no-install-recommends \
        collaboraofficebasis-python-script-provider \
        collaboraofficebasis-pyuno \
    && rm -rf /var/lib/apt/lists/*

# Enable macro execution (level 3 = signed + trusted locations)
RUN COOLWSD=/etc/coolwsd/coolwsd.xml && \
    sed -i 's|<enable_macros_execution[^>]*>[^<]*</enable_macros_execution>|<enable_macros_execution desc="" type="bool" default="false">true</enable_macros_execution>|' "$COOLWSD" && \
    sed -i 's|<macro_security_level[^>]*>[^<]*</macro_security_level>|<macro_security_level desc="" type="int" default="4">3</macro_security_level>|' "$COOLWSD"

COPY scripts/zotero_fields.py scripts/zotero_export.py \
     /opt/collaboraoffice/share/Scripts/python/

USER cool
```

Build and use:

```bash
docker build -t collabora-zotero .
```

Then reference `collabora-zotero` in your compose file or `docker run`
instead of `collabora/code`.

## Verifying it works

1. **Packages installed:**
   ```bash
   docker exec <container> ls /opt/collaboraoffice/program/ | grep python
   ```
   You should see `python-core-*` directories.

2. **Scripts in place:**
   ```bash
   docker exec <container> ls /opt/collaboraoffice/share/Scripts/python/zotero_*.py
   ```

3. **Macros enabled** — open a document, open the browser console, and check
   for errors after clicking a Zotero button. A silent failure (no error, no
   result) usually means macros are disabled. Check the container logs:
   ```bash
   docker logs <container> 2>&1 | grep -i macro
   ```

## Nextcloud integration

If using Nextcloud with the **Collabora Online** app (`richdocuments`),
deploy Collabora as a separate Docker container using Option B above, then
point Nextcloud at it via **Settings → Administration → Collabora Online**.

The built-in CODE server (`richdocumentscode` Nextcloud app) ships without
Python scripting support and cannot run these scripts.
