# zotero_fields.py — Zotero citation field storage for Collabora Online
#
# Installed in: /opt/collaboraoffice/share/Scripts/python/
# Called via the Collabora PostMessage CallPythonScript API.
#
# Storage design (compatible with the native Zotero LibreOffice plugin):
#   Field location : ReferenceMark wrapping visible citation text
#   Field code     : encoded in the reference mark name:
#                    ZOTERO_<code> RND<13 alphanum ID>
#   Document prefs : custom property ZOTERO_PREF_1, _2, … (255-char chunks)
#
# This matches the default ReferenceMark mode of the native Zotero LibreOffice
# plugin, so documents are fully interchangeable between Collabora and desktop.

import json
import random
import re
import string

# ── Constants ────────────────────────────────────────────────────────────────

CHUNK_SIZE = 255             # chars per document-property chunk
DOC_PREFS_PROP = "ZOTERO_PREF"
LOCK_PROP = "Zotero_Lock"
LOCK_TIMEOUT_SECS = 30
_ID_CHARS = string.ascii_letters + string.digits  # a-zA-Z0-9
_ID_LEN = 13                # native Zotero uses 13-char RND suffix
BREF_PREFIX = "ZOTERO_BREF_"  # for migrating old bookmark-based docs

# ── Internal helpers ─────────────────────────────────────────────────────────

def _doc():
    return XSCRIPTCONTEXT.getDocument()  # noqa: F821 – provided by LibreOffice runtime


def _udprops(doc):
    return doc.getDocumentProperties().getUserDefinedProperties()


def _prop_exists(props, name):
    return props.getPropertySetInfo().hasPropertyByName(name)


def _read_prop(props, name):
    return props.getPropertyValue(name) if _prop_exists(props, name) else None


def _write_prop(props, name, value):
    if _prop_exists(props, name):
        props.setPropertyValue(name, value)
    else:
        props.addProperty(name, 128, value)  # 128 = REMOVEABLE


def _del_prop(props, name):
    if _prop_exists(props, name):
        props.removeProperty(name)


# ── Reference mark helpers ───────────────────────────────────────────────────

def _new_id():
    """Generate a 13-char alphanumeric ID (matches native Zotero RND suffix)."""
    return "".join(random.choices(_ID_CHARS, k=_ID_LEN))


def _rm_name(code, fid):
    """Build a reference mark name: ZOTERO_<code> RND<id>"""
    return f"ZOTERO_{code} RND{fid}"


def _parse_rm(name):
    """Extract (code, fieldID) from a Zotero reference mark name.

    Name format: ZOTERO_<code> RND<13 alphanum>
    Returns (code, fid) or (None, None) if not a Zotero mark.
    """
    m = re.match(r'^ZOTERO_(.+)\s+RND([A-Za-z0-9]+)$', name)
    if m:
        return m.group(1), m.group(2)
    return None, None


def _find_rm(doc, fid):
    """Find a Zotero field by its RND field ID.

    Searches reference marks first, then text sections (LibreOffice stores
    bibliographies as sections after refresh).
    Returns (obj, name, kind) where kind is "rm" or "section", or (None, None, None).
    """
    suffix = f" RND{fid}"
    rms = doc.getReferenceMarks()
    for name in rms.getElementNames():
        if name.endswith(suffix):
            return rms.getByName(name), name, "rm"
    # LibreOffice's Zotero plugin stores bibliography as text:section
    sections = doc.getTextSections()
    for name in sections.getElementNames():
        if name.endswith(suffix) and name.startswith("ZOTERO_"):
            return sections.getByName(name), name, "section"
    return None, None, None


def _find_rm_name(doc, fid):
    """Find a reference mark name by field ID."""
    _, name, _ = _find_rm(doc, fid)
    return name


# ── Document-order reference mark enumeration ────────────────────────────────

def _zotero_rms_in_order(doc):
    """Return (code, fid, rm_name) tuples for all Zotero RMs in document order."""
    result = []
    seen = set()

    def visit_text(text_obj):
        en = text_obj.createEnumeration()
        while en.hasMoreElements():
            para = en.nextElement()
            if para.supportsService("com.sun.star.text.TextTable"):
                for cell_name in para.getCellNames():
                    visit_text(para.getCellByName(cell_name).getText())
            else:
                pe = para.createEnumeration()
                while pe.hasMoreElements():
                    portion = pe.nextElement()
                    ptype = portion.TextPortionType
                    if ptype == "ReferenceMark":
                        try:
                            rm = portion.ReferenceMark
                            name = rm.Name
                            if name.startswith("ZOTERO_") and name not in seen:
                                code, fid = _parse_rm(name)
                                if fid:
                                    seen.add(name)
                                    result.append((code, fid, name))
                        except Exception:
                            pass
                    elif ptype == "Footnote":
                        try:
                            visit_text(portion.Footnote.getText())
                        except Exception:
                            pass

    visit_text(doc.getText())
    enum_count = len(result)

    # Catch any Zotero RMs missed by text-portion enumeration.
    # This happens for reference marks that span multiple paragraphs
    # (e.g., bibliography entries separated by newlines).
    rms = doc.getReferenceMarks()
    all_names = list(rms.getElementNames())
    for name in all_names:
        if name in seen or not name.startswith("ZOTERO_"):
            continue
        code, fid = _parse_rm(name)
        if fid:
            seen.add(name)
            result.append((code, fid, name))

    fallback_count = len(result) - enum_count

    # LibreOffice's Zotero plugin stores bibliography as text:section after
    # refresh.  Include those so getFields sees them too.
    sections = doc.getTextSections()
    section_names = list(sections.getElementNames())
    section_count = 0
    for name in section_names:
        if name in seen or not name.startswith("ZOTERO_"):
            continue
        code, fid = _parse_rm(name)
        if fid:
            seen.add(name)
            result.append((code, fid, name))
            section_count += 1

    import sys
    all_zotero = [n[:80] for n in all_names if n.startswith("ZOTERO_")]
    all_zotero_sec = [n[:80] for n in section_names if n.startswith("ZOTERO_")]
    print(f"[ZOTERO] _zotero_rms_in_order: {enum_count} from enum, "
          f"{fallback_count} from fallback, {section_count} from sections, "
          f"all_rm_names={all_zotero}, all_section_names={all_zotero_sec}",
          file=sys.stderr, flush=True)

    return result


# ── Note index ───────────────────────────────────────────────────────────────

def _note_index_for_rm(doc, rm_name):
    """Return the 1-based footnote/endnote number containing the reference mark.

    noteIndex 0 = in-text, 1 = first footnote, 2 = second, etc.
    """
    try:
        rm = doc.getReferenceMarks().getByName(rm_name)
        anchor = rm.getAnchor()
        text = anchor.getText()

        if text.supportsService("com.sun.star.text.Footnote"):
            footnotes = doc.getFootnotes()
            for i in range(footnotes.getCount()):
                if footnotes.getByIndex(i).getText() == text:
                    return i + 1
            return 1

        if text.supportsService("com.sun.star.text.Endnote"):
            endnotes = doc.getEndnotes()
            for i in range(endnotes.getCount()):
                if endnotes.getByIndex(i).getText() == text:
                    return i + 1
            return 1

    except Exception:
        pass
    return 0


# ── HTML stripping ───────────────────────────────────────────────────────────

def _strip_html(html):
    """Strip HTML tags and decode entities for plain-text fallback."""
    import html as _html
    # IEEE/numbered styles use csl-left-margin + csl-right-inline divs within
    # a single entry: join them with a tab so "[1]\tText..." stays on one line.
    html = re.sub(
        r'</div>\s*<div\s+class="csl-right-inline">', "\t", html, flags=re.IGNORECASE
    )
    html = re.sub(r"</?(p|div|br|li|tr|h[1-6])[^>]*>", "\n", html, flags=re.IGNORECASE)
    html = re.sub(r"<[a-zA-Z/][^>]*>", "", html)
    html = _html.unescape(html)
    return re.sub(r"[ \t]*\n[ \t\n]*", "\n", html).strip()


# ── Bookmark → ReferenceMark migration ───────────────────────────────────────
#
# Documents created with the old bookmark-based version store codes in custom
# properties (ZOTERO_BREF_<id>_1, _2, …).  On first access we convert them
# to reference marks so everything uses one format going forward.

def _migrate_bookmarks(doc):
    """Convert old ZOTERO_BREF_* bookmarks to reference marks."""
    import sys
    bookmarks = doc.getBookmarks()
    bm_names = [n for n in bookmarks.getElementNames() if n.startswith(BREF_PREFIX)]
    if not bm_names:
        return False

    props = _udprops(doc)
    print(f"[ZOTERO] Migrating {len(bm_names)} bookmarks to reference marks",
          file=sys.stderr, flush=True)

    for bm_name in bm_names:
        try:
            old_fid = bm_name[len(BREF_PREFIX):]
            # Read code from custom properties (old format)
            code = _read_old_bookmark_code(props, old_fid)
            if not code:
                continue

            bm = bookmarks.getByName(bm_name)
            anchor = bm.getAnchor()
            parent_text = anchor.getText()
            visible = anchor.getString() or "\u200b"

            cursor = parent_text.createTextCursorByRange(anchor)
            anchor.setString("")
            parent_text.removeTextContent(bm)

            # Insert text and wrap in reference mark
            fid = _new_id()
            parent_text.insertString(cursor, visible, True)
            new_rm = doc.createInstance("com.sun.star.text.ReferenceMark")
            new_rm.Name = _rm_name(code, fid)
            parent_text.insertTextContent(cursor, new_rm, True)

            # Clean up old custom properties
            _erase_old_bookmark_code(props, old_fid)

            print(f"[ZOTERO] Migrated BM {old_fid} → RM {fid}",
                  file=sys.stderr, flush=True)
        except Exception as e:
            print(f"[ZOTERO] Failed to migrate BM {bm_name}: {e}",
                  file=sys.stderr, flush=True)

    return True


def _read_old_bookmark_code(props, field_id):
    """Read a field code from old bookmark-style custom properties."""
    key = BREF_PREFIX + field_id
    chunks = []
    i = 1
    while True:
        chunk = _read_prop(props, f"{key}_{i}")
        if chunk is None:
            break
        chunks.append(chunk)
        i += 1
    raw = "".join(chunks)
    if raw.startswith("ZOTERO_"):
        return raw[len("ZOTERO_"):]
    return raw


def _erase_old_bookmark_code(props, field_id):
    """Delete old bookmark-style custom properties."""
    key = BREF_PREFIX + field_id
    i = 1
    while _prop_exists(props, f"{key}_{i}"):
        _del_prop(props, f"{key}_{i}")
        i += 1


# ── RM create/update helpers ─────────────────────────────────────────────────

def _create_rm_at_cursor(doc, cursor, parent_text, code, fid, visible):
    """Insert text and wrap it in a new reference mark. Returns the RM name."""
    import sys
    content = visible or "\u200b"
    rm_name = _rm_name(code, fid)
    try:
        parent_text.insertString(cursor, content, True)
        new_rm = doc.createInstance("com.sun.star.text.ReferenceMark")
        new_rm.Name = rm_name
        parent_text.insertTextContent(cursor, new_rm, True)
    except Exception as e:
        print(f"[ZOTERO] _create_rm_at_cursor FAILED: fid={fid!r} error={e}",
              file=sys.stderr, flush=True)
        raise
    return rm_name


def _update_rm(doc, rm, name, new_code, new_visible):
    """Update a reference mark's code and/or visible text.

    Strategy: create a text cursor spanning the RM's full range, remove just
    the RM wrapper (text stays as plain text in the document), then replace
    the text and insert a new RM using the surviving cursor.  This avoids
    the need for bookmark/view-cursor position recovery which fails in
    Collabora when citations live inside footnotes.
    """
    import sys
    fid = _parse_rm(name)[1]
    anchor = rm.getAnchor()
    parent_text = anchor.getText()
    existing_text = anchor.getString()
    is_collapsed = (existing_text == "")

    code_to_write = new_code
    text_to_write = new_visible if new_visible is not None else (existing_text or "\u200b")

    if is_collapsed:
        # Inline mark: insert text at the mark position first so the mark
        # expands to wrap it, then use standard delete+recreate on the now-ranged mark.
        cur = parent_text.createTextCursorByRange(anchor.getStart())
        parent_text.insertString(cur, text_to_write, False)
        # Re-fetch: the mark should now span the inserted text
        rm2 = doc.getReferenceMarks().getByName(name)
        if rm2:
            rm, name, anchor, parent_text = rm2, rm2.Name, rm2.getAnchor(), rm2.getAnchor().getText()
            existing_text = anchor.getString()

    # Create a cursor spanning the full RM range BEFORE removing it.
    # Text cursors track positions in the underlying text, independent of
    # the RM content object, so they survive removeTextContent.
    cursor = parent_text.createTextCursorByRange(anchor)

    # Remove just the RM wrapper — the text it contained stays in the document
    # as plain (unwrapped) text, and `cursor` still spans it.
    try:
        parent_text.removeTextContent(rm)
    except Exception as e:
        print(f"[ZOTERO] _update_rm: removeTextContent failed: {e}",
              file=sys.stderr, flush=True)
        # Fallback: clear text and try again
        anchor.setString("")
        try:
            parent_text.removeTextContent(rm)
        except Exception:
            pass

    # Replace the old text with new content (cursor selects the range)
    parent_text.insertString(cursor, text_to_write, True)

    # Wrap the new text in a fresh reference mark
    new_rm = doc.createInstance("com.sun.star.text.ReferenceMark")
    rm_name = _rm_name(code_to_write, fid)
    new_rm.Name = rm_name
    try:
        parent_text.insertTextContent(cursor, new_rm, True)
    except Exception as e:
        print(f"[ZOTERO] _update_rm: insertTextContent failed: {e}",
              file=sys.stderr, flush=True)
        raise

    return rm_name


# ── Section helpers (bibliography stored as text:section by LibreOffice) ────

BIB_PARA_STYLE = "Bibliography 1"


def _get_section_text(section):
    """Get the concatenated plain text from all paragraphs in a section."""
    text = section.getAnchor().getString()
    return text or ""


def _create_bib_section(doc, cursor, parent_text, code, fid, visible):
    """Create a text:section for a bibliography field.

    Each newline-separated entry becomes its own paragraph with the
    'Bibliography 1' paragraph style, matching LibreOffice's native format.
    Returns the section name.
    """
    import sys
    from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK

    entries = [e.strip() for e in visible.split('\n') if e.strip()] if visible else []
    if not entries:
        entries = ["\u200b"]

    section_name = _rm_name(code, fid)

    # Insert first entry text, selected by cursor
    parent_text.insertString(cursor, entries[0], True)

    # Wrap it in a section
    section = doc.createInstance("com.sun.star.text.TextSection")
    section.Name = section_name
    parent_text.insertTextContent(cursor, section, True)

    # Style the first paragraph
    try:
        anchor = section.getAnchor()
        pc = parent_text.createTextCursorByRange(anchor.getStart())
        pc.gotoStartOfParagraph(False)
        pc.gotoEndOfParagraph(True)
        pc.setPropertyValue("ParaStyleName", BIB_PARA_STYLE)
    except Exception as e:
        print(f"[ZOTERO] _create_bib_section: style first para failed: {e}",
              file=sys.stderr, flush=True)

    # Add remaining entries as new paragraphs inside the section
    if len(entries) > 1:
        try:
            anchor = section.getAnchor()
            ec = parent_text.createTextCursorByRange(anchor.getEnd())
            ec.collapseToEnd()
            for entry in entries[1:]:
                parent_text.insertControlCharacter(ec, PARAGRAPH_BREAK, False)
                parent_text.insertString(ec, entry, False)
                # Style this paragraph
                try:
                    ec.gotoStartOfParagraph(False)
                    ec.gotoEndOfParagraph(True)
                    ec.setPropertyValue("ParaStyleName", BIB_PARA_STYLE)
                    ec.collapseToEnd()
                except Exception:
                    pass
        except Exception as e:
            print(f"[ZOTERO] _create_bib_section: add entries failed: {e}",
                  file=sys.stderr, flush=True)

    print(f"[ZOTERO] _create_bib_section: created section {section_name[:60]}... "
          f"with {len(entries)} entries", file=sys.stderr, flush=True)
    return section_name


def _update_section(doc, section, name, new_code, new_visible):
    """Update a section-based bibliography field.

    Replaces section content with one paragraph per bibliography entry
    and renames the section if the code changed.
    """
    import sys
    from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK

    fid = _parse_rm(name)[1]
    code_to_write = new_code

    # Update visible text: replace content with one paragraph per entry
    if new_visible is not None:
        entries = [e.strip() for e in new_visible.split('\n') if e.strip()]
        if not entries:
            entries = ["\u200b"]

        anchor = section.getAnchor()
        parent_text = anchor.getText()

        # Clear existing content
        anchor.setString("")

        # Write first entry into the now-empty section paragraph
        sc = parent_text.createTextCursorByRange(anchor.getStart())
        parent_text.insertString(sc, entries[0], False)
        try:
            sc.gotoStartOfParagraph(False)
            sc.gotoEndOfParagraph(True)
            sc.setPropertyValue("ParaStyleName", BIB_PARA_STYLE)
            sc.collapseToEnd()
        except Exception:
            pass

        # Add remaining entries
        for entry in entries[1:]:
            try:
                parent_text.insertControlCharacter(sc, PARAGRAPH_BREAK, False)
                parent_text.insertString(sc, entry, False)
                sc.gotoStartOfParagraph(False)
                sc.gotoEndOfParagraph(True)
                sc.setPropertyValue("ParaStyleName", BIB_PARA_STYLE)
                sc.collapseToEnd()
            except Exception:
                pass

    # Rename the section if the code changed
    new_name = _rm_name(code_to_write, fid)
    if new_name != name:
        try:
            section.Name = new_name
        except Exception as e:
            print(f"[ZOTERO] _update_section: rename failed: {e}",
                  file=sys.stderr, flush=True)

    return new_name


def _delete_section(doc, section):
    """Delete a section and its content from the document."""
    try:
        anchor = section.getAnchor()
        anchor.setString("")
        doc.getText().removeTextContent(section)
    except Exception:
        pass


def _rm_to_section(doc, rm, name, new_code):
    """Convert a reference mark bibliography field to a section.

    Called when setFieldCode receives a BIBL code for an RM-based field.
    Returns (section, new_name).
    """
    import sys
    fid = _parse_rm(name)[1]
    anchor = rm.getAnchor()
    parent_text = anchor.getText()
    visible = anchor.getString() or "\u200b"

    # Remove the RM (cursor spans the range, text stays as plain text)
    cursor = parent_text.createTextCursorByRange(anchor)
    parent_text.removeTextContent(rm)
    cursor.setString("")  # clear the leftover plain text

    # Create a section at the same position
    section_name = _create_bib_section(doc, cursor, parent_text, new_code, fid, visible)
    print(f"[ZOTERO] _rm_to_section: converted RM → section for fid={fid}",
          file=sys.stderr, flush=True)
    section = doc.getTextSections().getByName(section_name)
    return section, section_name


# ── Public API ───────────────────────────────────────────────────────────────

def getFields(fieldType=None, **_):
    """Return all Zotero fields in document order.

    Returns JSON: {"fieldIDs": [...], "fieldCodes": [...], "noteIndices": [...], "fieldTexts": [...]}
    """
    doc = _doc()
    _migrate_bookmarks(doc)
    fields = _zotero_rms_in_order(doc)
    rms = doc.getReferenceMarks()
    sections = doc.getTextSections()

    ids, codes, notes, texts = [], [], [], []
    for code, fid, name in fields:
        try:
            if rms.hasByName(name):
                vis = rms.getByName(name).getAnchor().getString() or ""
            elif sections.hasByName(name):
                vis = _get_section_text(sections.getByName(name))
            else:
                vis = ""
        except Exception:
            vis = ""
        if not code and not vis.strip():
            continue
        if code.startswith("BIBL") and not vis.strip():
            vis = "{Bibliography}"
        ids.append(fid)
        codes.append(code)
        notes.append(_note_index_for_rm(doc, name))
        texts.append(vis)

    import sys
    code_previews = [c[:60] for c in codes]
    print(f"[ZOTERO] getFields: {len(ids)} fields, ids={ids}, code_previews={code_previews}, notes={notes}",
          file=sys.stderr, flush=True)

    return json.dumps({
        "fieldIDs": ids,
        "fieldCodes": codes,
        "noteIndices": notes,
        "fieldTexts": texts,
    })


def insertField(fieldType=None, noteType=0, **_):
    """Insert a new field at the current cursor position.

    Returns JSON: {"fieldID": str, "fieldCode": str, "noteIndex": int}
    """
    doc = _doc()
    ctrl = doc.getCurrentController()
    vc = ctrl.getViewCursor()
    fid = _new_id()
    note_index = 0

    # Create a temporary RM with empty code — setFieldCode will rename it
    temp_code = "TEMP"

    if noteType in (1, 2):
        svc = ("com.sun.star.text.Footnote" if noteType == 1
               else "com.sun.star.text.Endnote")
        note = doc.createInstance(svc)
        doc.getText().insertTextContent(vc, note, False)
        note_text = note.getText()
        nc = note_text.createTextCursor()
        nc.gotoStart(False)
        nc.gotoEnd(True)
        _create_rm_at_cursor(doc, nc, note_text, temp_code, fid, "\u200b")
        note_index = noteType
    else:
        text = doc.getText()
        tc = text.createTextCursorByRange(vc)
        _create_rm_at_cursor(doc, tc, text, temp_code, fid, "\u200b")

    return json.dumps({"fieldID": fid, "fieldCode": "", "noteIndex": note_index})


def cursorInField(fieldType=None, **_):
    """Check whether the view cursor is inside a Zotero field.

    Returns JSON: {"fieldID": str, "fieldCode": str, "noteIndex": int} or null
    """
    doc = _doc()
    ctrl = doc.getCurrentController()
    vc = ctrl.getViewCursor()

    # Check reference marks
    rms = doc.getReferenceMarks()
    for name in rms.getElementNames():
        code, fid = _parse_rm(name)
        if not fid:
            continue
        rm = rms.getByName(name)
        anchor = rm.getAnchor()
        vis = anchor.getString() or ""
        if not code and not vis.strip():
            continue
        text = anchor.getText()
        try:
            s = text.compareRegionStarts(anchor, vc)
            e = text.compareRegionEnds(anchor, vc)
            if s >= 0 and e <= 0:
                return json.dumps({
                    "fieldID": fid,
                    "fieldCode": code,
                    "fieldText": vis,
                    "noteIndex": _note_index_for_rm(doc, name),
                })
        except Exception:
            pass

    # Check sections (LibreOffice bibliography)
    sections = doc.getTextSections()
    for name in sections.getElementNames():
        code, fid = _parse_rm(name)
        if not fid:
            continue
        section = sections.getByName(name)
        anchor = section.getAnchor()
        text = anchor.getText()
        try:
            s = text.compareRegionStarts(anchor, vc)
            e = text.compareRegionEnds(anchor, vc)
            if s >= 0 and e <= 0:
                vis = _get_section_text(section)
                return json.dumps({
                    "fieldID": fid,
                    "fieldCode": code,
                    "fieldText": vis,
                    "noteIndex": 0,
                })
        except Exception:
            pass

    return json.dumps(None)


def selectField(fieldID=None, **_):
    """Move the view cursor to select the named field."""
    doc = _doc()
    obj, _, kind = _find_rm(doc, fieldID)
    if obj:
        doc.getCurrentController().select(obj.getAnchor())
    return json.dumps(None)


def deleteField(fieldID=None, **_):
    """Delete a field and its visible text."""
    doc = _doc()
    obj, _, kind = _find_rm(doc, fieldID)
    if obj and kind == "section":
        _delete_section(doc, obj)
    elif obj:
        anchor = obj.getAnchor()
        anchor.setString("")
        anchor.getText().removeTextContent(obj)
    return json.dumps(None)


def removeFieldCode(fieldID=None, **_):
    """Remove the field code, leaving the visible text as plain text.

    Removes the reference mark or section wrapper; the text stays in the document.
    """
    import sys
    doc = _doc()
    obj, name, kind = _find_rm(doc, fieldID)
    vis = ""
    if obj and kind == "section":
        try:
            vis = _get_section_text(obj)
        except Exception:
            pass
        # Remove section wrapper, keeping its text content
        try:
            obj.getAnchor().getText().removeTextContent(obj)
        except Exception:
            pass
    elif obj:
        try:
            vis = obj.getAnchor().getString() or ""
        except Exception:
            pass
        try:
            obj.getAnchor().getText().removeTextContent(obj)
        except Exception:
            pass
    print(f"[ZOTERO] removeFieldCode: fieldID={fieldID!r} kind={kind!r} exists={obj is not None} visible_text_len={len(vis)}",
          file=sys.stderr, flush=True)
    return json.dumps(None)


def _is_bibl_code(code):
    """Check if a field code is for a bibliography."""
    return code and code.startswith("BIBL")


def setFields(updates=None, **_):
    """Batch-update field visible text and codes."""
    doc = _doc()
    if isinstance(updates, str):
        updates = json.loads(updates)

    for upd in updates:
        fid = upd["fieldID"]
        text = upd.get("text", "")
        code = upd.get("code", "")
        is_rich = upd.get("isRich", False)
        display = _strip_html(text) if is_rich else text

        obj, name, kind = _find_rm(doc, fid)
        if not obj:
            continue

        # Convert RM → section for bibliography fields
        if kind == "rm" and _is_bibl_code(code):
            obj, name = _rm_to_section(doc, obj, name, code)
            kind = "section"

        if kind == "section":
            _update_section(doc, obj, name, code, display or None)
        else:
            _update_rm(doc, obj, name, code, display or None)

    return json.dumps(None)


def setFieldCode(fieldID=None, code="", **_):
    """Set the field code for a reference mark or section.

    When a BIBL code is set on an RM-based field, converts it to a section
    to match LibreOffice's native Zotero bibliography format.
    """
    import sys
    print(f"[ZOTERO] setFieldCode: fieldID={fieldID!r} code={code!r}",
          file=sys.stderr, flush=True)
    if not fieldID or not isinstance(fieldID, str):
        return json.dumps(None)

    doc = _doc()
    obj, name, kind = _find_rm(doc, fieldID)
    if not obj:
        return json.dumps(None)

    # Convert RM → section for bibliography
    if kind == "rm" and _is_bibl_code(code):
        _rm_to_section(doc, obj, name, code)
        return json.dumps(None)

    if kind == "section":
        _update_section(doc, obj, name, code, None)
    else:
        _update_rm(doc, obj, name, code, None)
    return json.dumps(None)


def setFieldText(fieldID=None, text="", isRich=False, **_):
    """Set the visible text for a single field."""
    import sys
    preview = repr(text[:200]) if text else 'None'
    print(f"[ZOTERO] setFieldText: fieldID={fieldID!r} isRich={isRich!r} text_len={len(text) if text else 0} text_preview={preview}",
          file=sys.stderr, flush=True)

    doc = _doc()
    obj, name, kind = _find_rm(doc, fieldID)
    if not obj:
        print(f"[ZOTERO] setFieldText: field for {fieldID!r} NOT FOUND",
              file=sys.stderr, flush=True)
        return json.dumps(None)

    code, _ = _parse_rm(name)
    display_text = _strip_html(text) if isRich else text

    if kind == "section":
        _update_section(doc, obj, name, code, display_text or "\u200b")
    else:
        _update_rm(doc, obj, name, code, display_text or "\u200b")
    return json.dumps(None)


def getFieldText(fieldID=None, **_):
    """Return the visible text of a single field."""
    import sys
    print(f"[ZOTERO] getFieldText: fieldID={fieldID!r}", file=sys.stderr, flush=True)
    doc = _doc()
    obj, name, kind = _find_rm(doc, fieldID)
    vis = ""
    if obj:
        try:
            if kind == "section":
                vis = _get_section_text(obj)
            else:
                vis = obj.getAnchor().getString() or ""
        except Exception:
            pass
    if not vis.strip() and name:
        code, _ = _parse_rm(name)
        if code and code.startswith("BIBL"):
            vis = "{Bibliography}"
    return json.dumps(vis)


def getFieldCode(fieldID=None, **_):
    """Return the stored code for a single field."""
    import sys
    doc = _doc()
    _, name, _ = _find_rm(doc, fieldID)
    code = ""
    if name:
        code, _ = _parse_rm(name)
        code = code or ""
    print(f"[ZOTERO] getFieldCode: fieldID={fieldID!r} code_len={len(code)}",
          file=sys.stderr, flush=True)
    return json.dumps(code)


def getFieldNoteIndex(fieldID=None, **_):
    """Return the note index for a single field."""
    import sys
    doc = _doc()
    name = _find_rm_name(doc, fieldID)
    idx = _note_index_for_rm(doc, name) if name else 0
    print(f"[ZOTERO] getFieldNoteIndex: fieldID={fieldID!r} noteIndex={idx}",
          file=sys.stderr, flush=True)
    return json.dumps(idx)


def getDocumentData(**_):
    """Read the Zotero document preferences string.

    On disk: fieldType="ReferenceMark". HTTP connector expects "Http".
    """
    doc = _doc()
    props = _udprops(doc)
    chunks = []
    i = 1
    while True:
        chunk = _read_prop(props, f"{DOC_PREFS_PROP}_{i}")
        if chunk is None:
            break
        chunks.append(chunk)
        i += 1
    data = "".join(chunks)
    # Replace any fieldType with "Http" for the HTTP connector
    data = re.sub(r'(<pref\s+name="fieldType"\s+value=")[^"]*(")', r'\1Http\2', data)
    data = re.sub(r'"fieldType"\s*:\s*"[^"]*"', '"fieldType": "Http"', data)
    return json.dumps(data)


def getDocumentState(**_):
    """Return document prefs + all fields in one call."""
    doc = _doc()
    _migrate_bookmarks(doc)

    props = _udprops(doc)
    chunks = []
    i = 1
    while True:
        chunk = _read_prop(props, f"{DOC_PREFS_PROP}_{i}")
        if chunk is None:
            break
        chunks.append(chunk)
        i += 1
    data = "".join(chunks)
    data = re.sub(r'(<pref\s+name="fieldType"\s+value=")[^"]*(")', r'\1Http\2', data)
    data = re.sub(r'"fieldType"\s*:\s*"[^"]*"', '"fieldType": "Http"', data)

    fields = _zotero_rms_in_order(doc)
    rms = doc.getReferenceMarks()
    sections = doc.getTextSections()
    ids, codes, notes, texts = [], [], [], []
    for code, fid, name in fields:
        try:
            if rms.hasByName(name):
                vis = rms.getByName(name).getAnchor().getString() or ""
            elif sections.hasByName(name):
                vis = _get_section_text(sections.getByName(name))
            else:
                vis = ""
        except Exception:
            vis = ""
        if not code and not vis.strip():
            continue
        if code.startswith("BIBL") and not vis.strip():
            vis = "{Bibliography}"
        ids.append(fid)
        codes.append(code)
        notes.append(_note_index_for_rm(doc, name))
        texts.append(vis)

    return json.dumps({
        "documentData": data,
        "fieldIDs": ids,
        "fieldCodes": codes,
        "noteIndices": notes,
        "fieldTexts": texts,
    })


def setDocumentData(data="", **_):
    """Write the Zotero document preferences string.

    HTTP connector sends fieldType="Http". We store "ReferenceMark"
    so the native LibreOffice Zotero plugin recognises the document.
    """
    data = re.sub(r'(<pref\s+name="fieldType"\s+value=")Http(")', r'\1ReferenceMark\2', data)
    data = data.replace('"fieldType":"Http"', '"fieldType":"ReferenceMark"')
    data = data.replace('"fieldType": "Http"', '"fieldType": "ReferenceMark"')

    doc = _doc()
    props = _udprops(doc)
    i = 1
    while _prop_exists(props, f"{DOC_PREFS_PROP}_{i}"):
        _del_prop(props, f"{DOC_PREFS_PROP}_{i}")
        i += 1
    i = 1
    while (i - 1) * CHUNK_SIZE < len(data):
        _write_prop(props, f"{DOC_PREFS_PROP}_{i}",
                    data[(i - 1) * CHUNK_SIZE : i * CHUNK_SIZE])
        i += 1
    return json.dumps(None)


def flushUpdates(updates=None, **_):
    """Apply a batch of buffered operations from the extension.

    Each entry has a 'type' field:
      - 'field': {fieldID, code?, text?, isRich?} — update code and/or text
      - 'delete': {fieldID} — delete field
      - 'removeCode': {fieldID} — unlink field, keep text
      - 'setDocumentData': {data} — write doc prefs
      - 'setBibliographyStyle': {firstLineIndent, bodyIndent, lineSpacing, entrySpacing, tabStops, count}
    """
    import sys
    doc = _doc()
    if isinstance(updates, str):
        updates = json.loads(updates)

    print(f"[ZOTERO] flushUpdates: {len(updates)} operations", file=sys.stderr, flush=True)

    for op in updates:
        t = op.get("type")
        try:
            if t == "field":
                fid = op["fieldID"]
                code = op.get("code")
                text = op.get("text")
                is_rich = op.get("isRich", False)
                display = _strip_html(text) if (is_rich and text) else text

                obj, name, kind = _find_rm(doc, fid)
                if not obj:
                    continue

                new_code = code if code is not None else _parse_rm(name)[0]

                if kind == "rm" and _is_bibl_code(new_code):
                    obj, name = _rm_to_section(doc, obj, name, new_code)
                    kind = "section"

                if kind == "section":
                    _update_section(doc, obj, name, new_code, display or None)
                else:
                    _update_rm(doc, obj, name, new_code, display or None)

            elif t == "delete":
                obj, _, kind = _find_rm(doc, op["fieldID"])
                if obj and kind == "section":
                    _delete_section(doc, obj)
                elif obj:
                    anchor = obj.getAnchor()
                    anchor.setString("")
                    anchor.getText().removeTextContent(obj)

            elif t == "removeCode":
                obj, name, kind = _find_rm(doc, op["fieldID"])
                if obj and kind == "section":
                    try:
                        obj.getAnchor().getText().removeTextContent(obj)
                    except Exception:
                        pass
                elif obj:
                    try:
                        obj.getAnchor().getText().removeTextContent(obj)
                    except Exception:
                        pass

            elif t == "setDocumentData":
                data = op["data"]
                data = re.sub(r'(<pref\s+name="fieldType"\s+value=")Http(")', r'\1ReferenceMark\2', data)
                data = data.replace('"fieldType":"Http"', '"fieldType":"ReferenceMark"')
                data = data.replace('"fieldType": "Http"', '"fieldType": "ReferenceMark"')
                props = _udprops(doc)
                i = 1
                while _prop_exists(props, f"{DOC_PREFS_PROP}_{i}"):
                    _del_prop(props, f"{DOC_PREFS_PROP}_{i}")
                    i += 1
                i = 1
                while (i - 1) * CHUNK_SIZE < len(data):
                    _write_prop(props, f"{DOC_PREFS_PROP}_{i}",
                                data[(i - 1) * CHUNK_SIZE : i * CHUNK_SIZE])
                    i += 1

            elif t == "setBibliographyStyle":
                styles = doc.getStyleFamilies().getByName("ParagraphStyles")
                if styles.hasByName("Bibliography"):
                    style = styles.getByName("Bibliography")

                    def pt20_to_mm100(v):
                        return int(v * 0.353)

                    style.ParaFirstLineIndent = pt20_to_mm100(op.get("firstLineIndent", 0))
                    style.ParaLeftMargin = pt20_to_mm100(op.get("bodyIndent", 0))
                    ls_val = op.get("lineSpacing", 0)
                    if ls_val > 0:
                        from com.sun.star.style import LineSpacing, LineSpacingMode
                        ls = LineSpacing()
                        ls.Mode = LineSpacingMode.PROP
                        ls.Height = ls_val
                        style.ParaLineSpacing = ls
                    style.ParaBottomMargin = pt20_to_mm100(op.get("entrySpacing", 0))

        except Exception as e:
            print(f"[ZOTERO] flushUpdates: op {t} failed: {e}", file=sys.stderr, flush=True)

    return json.dumps(None)


def setBibliographyStyle(
    firstLineIndent=0, bodyIndent=0, lineSpacing=0,
    entrySpacing=0, tabStops=None, count=0, **_
):
    """Apply paragraph formatting to the bibliography paragraph style."""
    doc = _doc()
    try:
        styles = doc.getStyleFamilies().getByName("ParagraphStyles")
        if not styles.hasByName("Bibliography"):
            return json.dumps(None)
        style = styles.getByName("Bibliography")

        def pt20_to_mm100(v):
            return int(v * 0.353)

        style.ParaFirstLineIndent = pt20_to_mm100(firstLineIndent)
        style.ParaLeftMargin = pt20_to_mm100(bodyIndent)

        if lineSpacing > 0:
            from com.sun.star.style import LineSpacing, LineSpacingMode
            ls = LineSpacing()
            ls.Mode = LineSpacingMode.PROP
            ls.Height = lineSpacing
            style.ParaLineSpacing = ls

        style.ParaBottomMargin = pt20_to_mm100(entrySpacing)
    except Exception:
        pass
    return json.dumps(None)


def insertText(html="", **_):
    """Insert text at the current cursor position."""
    doc = _doc()
    vc = doc.getCurrentController().getViewCursor()
    text = vc.getText()
    tc = text.createTextCursorByRange(vc)
    text.insertString(tc, _strip_html(html), True)
    return json.dumps(None)


def convertPlaceholdersToFields(placeholderIDs=None, noteType=0, fieldType=None, **_):
    """Convert Zotero placeholder hyperlinks into reference mark fields."""
    if isinstance(placeholderIDs, str):
        placeholderIDs = json.loads(placeholderIDs)
    if not placeholderIDs:
        return json.dumps({"fieldIDs": [], "fieldCodes": [], "noteIndices": []})

    doc = _doc()
    placeholder_set = set(placeholderIDs)
    ids, codes, notes = [], [], []

    text_fields = doc.getTextFields()
    en = text_fields.createEnumeration()
    while en.hasMoreElements():
        tf = en.nextElement()
        try:
            if not tf.supportsService("com.sun.star.text.TextField.URL"):
                continue
            url = tf.URL
            matched_id = next((pid for pid in placeholder_set if pid in url), None)
            if not matched_id:
                continue

            anchor = tf.getAnchor()
            parent_text = anchor.getText()
            visible = anchor.getString() or "\u200b"
            fid = _new_id()

            tc = parent_text.createTextCursorByRange(anchor)
            parent_text.removeTextContent(tf)
            rm_name = _create_rm_at_cursor(doc, tc, parent_text, "", fid, visible)

            ids.append(fid)
            codes.append("")
            notes.append(_note_index_for_rm(doc, rm_name))
        except Exception:
            pass

    return json.dumps({"fieldIDs": ids, "fieldCodes": codes, "noteIndices": notes})


def convertFields(fieldIDs=None, toFieldType=None, toNoteTypes=None, count=0, **_):
    """Batch-convert field note types (e.g., in-text → footnote).

    With reference marks, footnote conversion is straightforward since
    reference marks work natively in footnotes (unlike bookmarks).
    """
    if not fieldIDs:
        return json.dumps(None)
    if isinstance(fieldIDs, str):
        fieldIDs = json.loads(fieldIDs)
    if isinstance(toNoteTypes, str):
        toNoteTypes = json.loads(toNoteTypes)
    if not toNoteTypes:
        toNoteTypes = [0] * len(fieldIDs)

    doc = _doc()
    main_text = doc.getText()

    for fid, target_nt in zip(fieldIDs, toNoteTypes):
        rm, name, kind = _find_rm(doc, fid)
        if not rm or kind == "section":
            continue  # sections (bibliography) don't convert to footnotes
        code, _ = _parse_rm(name)
        current_ni = _note_index_for_rm(doc, name)
        if (current_ni == 0 and target_nt == 0) or \
           (current_ni > 0 and target_nt > 0 and current_ni == target_nt):
            continue

        anchor = rm.getAnchor()
        container = anchor.getText()
        visible = anchor.getString() or "\u200b"

        try:
            if current_ni == 0 and target_nt in (1, 2):
                # In-text → footnote/endnote
                tc = main_text.createTextCursorByRange(anchor)
                anchor.setString("")
                main_text.removeTextContent(rm)
                svc = ("com.sun.star.text.Footnote" if target_nt == 1
                       else "com.sun.star.text.Endnote")
                note = doc.createInstance(svc)
                main_text.insertTextContent(tc, note, False)
                note_text = note.getText()
                nc = note_text.createTextCursor()
                nc.gotoStart(False)
                nc.gotoEnd(True)
                _create_rm_at_cursor(doc, nc, note_text, code, fid, visible)

            elif current_ni in (1, 2) and target_nt == 0:
                # Footnote/endnote → in-text
                note_anchor = container.getAnchor()
                pos = main_text.createTextCursorByRange(note_anchor)
                pos.collapseToStart()
                try:
                    container.removeTextContent(rm)
                except Exception:
                    pass
                _create_rm_at_cursor(doc, pos, main_text, code, fid, visible)
                main_text.removeTextContent(container)

            elif current_ni in (1, 2) and target_nt in (1, 2):
                # Footnote ↔ endnote
                note_anchor = container.getAnchor()
                pos = main_text.createTextCursorByRange(note_anchor)
                pos.collapseToStart()
                try:
                    container.removeTextContent(rm)
                except Exception:
                    pass
                main_text.removeTextContent(container)

                svc = ("com.sun.star.text.Footnote" if target_nt == 1
                       else "com.sun.star.text.Endnote")
                note = doc.createInstance(svc)
                main_text.insertTextContent(pos, note, False)
                note_text = note.getText()
                nc = note_text.createTextCursor()
                nc.gotoStart(False)
                nc.gotoEnd(True)
                _create_rm_at_cursor(doc, nc, note_text, code, fid, visible)

        except Exception:
            pass

    return json.dumps(None)


def exportDocument(fieldType=None, instructions=None, **_):
    """Convert all Zotero reference mark fields to transfer hyperlinks."""
    doc = _doc()
    rms = doc.getReferenceMarks()

    for name in list(rms.getElementNames()):
        code, fid = _parse_rm(name)
        if not fid:
            continue
        rm = rms.getByName(name)
        anchor = rm.getAnchor()
        text = anchor.getText()

        try:
            import urllib.parse
            encoded = urllib.parse.quote(code)
            url = f"zotero://transfer/{fid}?code={encoded}"

            tf = doc.createInstance("com.sun.star.text.TextField.URL")
            tf.URL = url
            tf.Representation = anchor.getString()
            tc = text.createTextCursorByRange(anchor)
            text.insertTextContent(tc, tf, True)
            text.removeTextContent(rm)
        except Exception:
            pass

    return json.dumps(None)


def importDocument(fieldType=None, **_):
    """Restore Zotero reference marks from transfer hyperlinks."""
    doc = _doc()
    text_fields = doc.getTextFields()
    en = text_fields.createEnumeration()

    to_convert = []
    while en.hasMoreElements():
        tf = en.nextElement()
        try:
            if not tf.supportsService("com.sun.star.text.TextField.URL"):
                continue
            url = tf.URL
            if not url.startswith("zotero://transfer/"):
                continue
            to_convert.append(tf)
        except Exception:
            pass

    for tf in to_convert:
        try:
            import urllib.parse
            url = tf.URL
            path = url[len("zotero://transfer/"):]
            if "?" in path:
                fid_part, query = path.split("?", 1)
                params = urllib.parse.parse_qs(query)
                code = urllib.parse.unquote(params.get("code", [""])[0])
            else:
                fid_part = path
                code = ""

            anchor = tf.getAnchor()
            parent_text = anchor.getText()
            visible = anchor.getString() or "\u200b"
            tc = parent_text.createTextCursorByRange(anchor)
            parent_text.removeTextContent(tf)

            fid = _new_id()
            _create_rm_at_cursor(doc, tc, parent_text, code, fid, visible)
        except Exception:
            pass

    return json.dumps(None)


def acquireLock(userID="", **_):
    """Try to acquire the document edit lock."""
    import time
    doc = _doc()
    props = _udprops(doc)
    raw = _read_prop(props, LOCK_PROP)
    if raw:
        try:
            lock = json.loads(raw)
            if time.time() - lock.get("ts", 0) < LOCK_TIMEOUT_SECS:
                if lock.get("userID") != userID:
                    return json.dumps({"acquired": False, "holder": lock.get("userID")})
        except Exception:
            pass
    _write_prop(props, LOCK_PROP, json.dumps({"userID": userID, "ts": time.time()}))
    return json.dumps({"acquired": True})


def releaseLock(userID="", **_):
    """Release the document edit lock."""
    doc = _doc()
    props = _udprops(doc)
    _del_prop(props, LOCK_PROP)
    return json.dumps(None)


# Required for LibreOffice UNO runtime to discover these functions
g_exportedScripts = (
    getFields, insertField, cursorInField, selectField, deleteField,
    removeFieldCode, setFields, setFieldCode, setFieldText, getFieldText,
    getFieldCode, getFieldNoteIndex,
    getDocumentData, setDocumentData, setBibliographyStyle, getDocumentState,
    insertText, convertPlaceholdersToFields, convertFields,
    exportDocument, importDocument, acquireLock, releaseLock,
    flushUpdates,
)
