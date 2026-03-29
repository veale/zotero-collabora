FROM collabora/code:latest

# Install Python scripting support for LibreOffice
USER root
RUN apt-get update && apt-get install -y --no-install-recommends \
        collaboraofficebasis-python-script-provider \
        collaboraofficebasis-pyuno \
    && rm -rf /var/lib/apt/lists/*

# Enable Python macro execution in coolwsd.xml.
# The default config ships with macro_security_level = 4 (disabled).
# Level 3 (High) requires signed macros, but scripts in trusted filesystem
# locations (the share directory where ours are installed) bypass signing.
# Document-embedded macros are blocked unless signed.
RUN COOLWSD=/etc/coolwsd/coolwsd.xml && \
    sed -i 's|<enable_macros_execution[^>]*>[^<]*</enable_macros_execution>|<enable_macros_execution desc="" type="bool" default="false">true</enable_macros_execution>|' "$COOLWSD" && \
    sed -i 's|<macro_security_level[^>]*>[^<]*</macro_security_level>|<macro_security_level desc="" type="int" default="4">3</macro_security_level>|' "$COOLWSD"

# Copy Zotero Python scripts into the global script directory
COPY scripts/zotero_fields.py  /opt/collaboraoffice/share/Scripts/python/zotero_fields.py
COPY scripts/zotero_export.py  /opt/collaboraoffice/share/Scripts/python/zotero_export.py

USER cool
