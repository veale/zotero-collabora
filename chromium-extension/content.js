// content.js — Content script running on the host page that embeds Collabora.
//
// Responsibilities:
//   1. Locate the Collabora iframe via the first PostMessage it sends
//   2. Wait for App_LoadingStatus → Document_Loaded
//   3. Inject Zotero toolbar buttons via Insert_Button PostMessages
//   4. On button click: notify background worker to start a Zotero transaction
//   5. On CALL_PYTHON from background: call Python via PostMessage, return result

// ── Collabora iframe detection ────────────────────────────────────────────────
//
// The Collabora iframe may be navigated via form POST (no src attribute) or
// have its src set directly. We identify it lazily from the first PostMessage
// it sends, by matching event.source against all iframes on the page.

let collaboraIframe = null;

function sendToCollabora(messageId, values) {
  if (!collaboraIframe?.contentWindow) return;
  collaboraIframe.contentWindow.postMessage(
    JSON.stringify({ MessageId: messageId, SendTime: Date.now(), Values: values }),
    '*'
  );
}

// ── Toolbar button icons ──────────────────────────────────────────────────────
//
// Icons must be absolute HTTP URLs reachable from inside the Collabora iframe.
// chrome-extension:// URLs are cross-origin to the iframe and won't load.
// We serve SVG icons from the WOPI host instead.

const WOPI_ORIGIN = 'http://localhost:8080';

// ── Toolbar buttons ───────────────────────────────────────────────────────────

const BUTTONS = [
  { id: 'zotero-add-citation',     label: 'Add Citation', hint: 'Add or edit a Zotero citation',       imgurl: `${WOPI_ORIGIN}/icons/cite.svg`,   command: 'addEditCitation' },
  { id: 'zotero-add-bibliography', label: 'Bibliography', hint: 'Insert or refresh bibliography',      imgurl: `${WOPI_ORIGIN}/icons/bib.svg`,    command: 'addEditBibliography' },
  { id: 'zotero-refresh',          label: 'Refresh',      hint: 'Refresh all Zotero fields',           imgurl: `${WOPI_ORIGIN}/icons/ref.svg`,    command: 'refresh' },
  { id: 'zotero-set-prefs',        label: 'Doc Prefs',    hint: 'Change citation style and settings',  imgurl: `${WOPI_ORIGIN}/icons/pref.svg`,   command: 'setDocPrefs' },
  { id: 'zotero-add-note',         label: 'Add Note',     hint: 'Insert a Zotero note',                imgurl: `${WOPI_ORIGIN}/icons/note.svg`,   command: 'addNote' },
  { id: 'zotero-remove-codes',     label: 'Unlink',       hint: 'Convert citations to plain text',     imgurl: `${WOPI_ORIGIN}/icons/unlink.svg`, command: 'removeCodes' },
  { id: 'zotero-export',           label: 'Export',       hint: 'Export citations as CSL-JSON/BibTeX', imgurl: `${WOPI_ORIGIN}/icons/export.svg`, command: 'exportCitations' },
];

let toolbarButtonsInjected = false;

function injectToolbarButtons() {
  if (toolbarButtonsInjected) return;
  for (const btn of BUTTONS) {
    sendToCollabora('Insert_Button', {
      id:     btn.id,
      hint:   btn.hint,
      label:  btn.label,
      imgurl: btn.imgurl,
      // insertBefore is omitted: broken in notebookbar mode, buttons go to shortcuts bar
    });
  }
  toolbarButtonsInjected = true;
  console.log('[Zotero] Injected', BUTTONS.length, 'toolbar buttons');
}

function removeToolbarButtons() {
  if (!toolbarButtonsInjected) return;
  // Collabora has no Remove_Button API; hide each button via its .uno style command
  for (const btn of BUTTONS) {
    sendToCollabora('Remove_Button', { id: btn.id });
  }
  toolbarButtonsInjected = false;
  console.log('[Zotero] Removed toolbar buttons');
}

async function maybeInjectToolbarButtons() {
  const data = await chrome.storage.local.get({ showToolbarButtons: true });
  if (data.showToolbarButtons) {
    injectToolbarButtons();
  }
}

// ── Python script invocation via PostMessage ──────────────────────────────────
//
// Correct format per Collabora Map.WOPI.js:
//   ScriptFile, Function, Values are top-level (not nested in a Values wrapper).
//   Each argument: { type: "string"|"long"|"boolean", value: ... }
//   Result path:   msg.Values.result.value  (a JSON-encoded string)
//   Success check: msg.Values.success === true  (boolean, not string)

function toTypedArgs(obj) {
  if (!obj || Object.keys(obj).length === 0) return null;
  const out = {};
  for (const [k, v] of Object.entries(obj)) {
    if (typeof v === 'boolean')                       out[k] = { type: 'boolean', value: v };
    else if (typeof v === 'number' && v === (v | 0)) out[k] = { type: 'long',    value: v };
    else                                              out[k] = { type: 'string',  value: String(v) };
  }
  return out;
}

// Pending call: only one in flight at a time (transactions are sequential).
let pendingPythonCall = null;

function callPython(scriptFile, fn, args) {
  return new Promise((resolve, reject) => {
    const timeout = setTimeout(() => {
      pendingPythonCall = null;
      reject(new Error(`Python call timed out after 20s: ${fn}`));
    }, 20_000);

    pendingPythonCall = {
      resolve: (val) => { clearTimeout(timeout); resolve(val); },
      reject:  (err) => { clearTimeout(timeout); reject(err); },
    };

    // ScriptFile/Function/Values must be top-level alongside MessageId
    if (!collaboraIframe?.contentWindow) return;
    collaboraIframe.contentWindow.postMessage(JSON.stringify({
      MessageId:  'CallPythonScript',
      SendTime:   Date.now(),
      ScriptFile: scriptFile,
      Function:   fn,
      Values:     toTypedArgs(args),
    }), '*');
  });
}

// ── Handle export button locally (no Zotero desktop needed) ──────────────────

async function handleExportCitations() {
  try {
    const raw = await callPython('zotero_export.py', 'exportCitations', { format: 'csljson' });
    const data = typeof raw === 'string' ? JSON.parse(raw) : raw;
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'citations.json';
    a.click();
    URL.revokeObjectURL(url);
  } catch (err) {
    zoteroDialog(`Export failed: ${err.message}`, 0);
  }
}

// ── Zotero dialog (replaces browser alert/confirm) ───────────────────────────
//
// Zotero button constants:
//   0 = OK only
//   1 = OK / Cancel       → resolve 1 (OK) or 0 (Cancel)
//   2 = Yes / No          → resolve 1 (Yes) or 0 (No)
//   3 = Yes / No / Cancel → resolve 2 (Yes), 1 (No), or 0 (Cancel)

function zoteroDialog(text, buttons) {
  return new Promise((resolve) => {
    const overlay = document.createElement('div');
    overlay.style.cssText = [
      'position:fixed', 'inset:0', 'z-index:2147483647',
      'background:rgba(0,0,0,.45)',
      'display:flex', 'align-items:center', 'justify-content:center',
      'font-family:system-ui,sans-serif',
    ].join(';');

    const box = document.createElement('div');
    box.style.cssText = [
      'background:#fff', 'border-radius:8px', 'padding:24px 28px',
      'max-width:480px', 'width:90%',
      'box-shadow:0 8px 32px rgba(0,0,0,.35)',
    ].join(';');

    const title = document.createElement('div');
    title.style.cssText = 'font-weight:700;font-size:15px;color:#c1121f;margin-bottom:14px';
    title.textContent = 'Zotero';

    const body = document.createElement('p');
    body.style.cssText = 'margin:0 0 22px;color:#222;font-size:13.5px;line-height:1.55;white-space:pre-wrap';
    body.textContent = text;

    const row = document.createElement('div');
    row.style.cssText = 'display:flex;gap:8px;justify-content:flex-end';

    const close = (val) => { overlay.remove(); resolve(val); };

    const btn = (label, val, primary) => {
      const b = document.createElement('button');
      b.textContent = label;
      b.style.cssText = [
        'padding:7px 18px', 'border-radius:5px', 'border:none',
        'cursor:pointer', 'font-size:13px', 'font-weight:500',
        primary
          ? 'background:#c1121f;color:#fff'
          : 'background:#e8e8e8;color:#333',
      ].join(';');
      b.onclick = () => close(val);
      row.appendChild(b);
    };

    // Build buttons — rightmost is always the "primary" affirmative action
    switch (buttons) {
      case 0:
        btn('OK', 1, true);
        break;
      case 2:
        btn('No',  0, false);
        btn('Yes', 1, true);
        break;
      case 3:
        btn('Cancel', 0, false);
        btn('No',     1, false);
        btn('Yes',    2, true);
        break;
      default: // 1 = OK / Cancel
        btn('Cancel', 0, false);
        btn('OK',     1, true);
        break;
    }

    box.append(title, body, row);
    overlay.appendChild(box);
    document.body.appendChild(overlay);

    // Allow Escape to dismiss (treated as the lowest-value button)
    overlay.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') close(0);
    });
    // Focus the primary button immediately
    setTimeout(() => row.lastChild?.focus(), 0);
  });
}

// ── PostMessage listener (from Collabora iframe) ──────────────────────────────

window.addEventListener('message', (event) => {
  let msg;
  try { msg = JSON.parse(event.data); } catch { return; }

  // Identify the Collabora iframe from the first message it sends.
  if (!collaboraIframe && event.source) {
    for (const iframe of document.querySelectorAll('iframe')) {
      if (iframe.contentWindow === event.source) {
        collaboraIframe = iframe;
        break;
      }
    }
  }

  switch (msg.MessageId) {
    case 'App_LoadingStatus':
      // The host page's inline script (OPEN_HTML) already sends Host_PostmessageReady
      // in response to App_LoadingStatus, so WOPIPostmessageReady is already true by the
      // time Document_Loaded fires. We do NOT send it again here to avoid timing issues.
      if (msg.Values?.Status === 'Document_Loaded') {
        // Small delay to ensure WOPIPostmessageReady is set inside the iframe
        // before Insert_Button messages arrive.
        setTimeout(() => maybeInjectToolbarButtons(), 300);
      }
      break;

    case 'Clicked_Button': {
      const btnId = msg.Values?.Id ?? msg.Values?.id;  // Collabora sends capital Id
      const btn = BUTTONS.find((b) => b.id === btnId);
      if (!btn) return;

      if (btn.command === 'exportCitations') {
        handleExportCitations();
        return;
      }

      // All other commands go through the Zotero desktop HTTP transaction
      chrome.runtime.sendMessage({ type: 'ZOTERO_COMMAND', command: btn.command });
      break;
    }

    case 'CallPythonScript-Result': {
      if (pendingPythonCall) {
        const cb = pendingPythonCall;
        pendingPythonCall = null;
        const success = msg.Values?.success;
        if (success !== true && success !== 'true') {
          cb.reject(new Error(`Python error: ${JSON.stringify(msg.Values)}`));
        } else {
          cb.resolve(msg.Values.result?.value ?? null);
        }
      }
      break;
    }

    default:
      break;
  }
});

// ── Runtime message listener (from background worker) ────────────────────────

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  if (msg.type === 'CALL_PYTHON') {
    callPython('zotero_fields.py', msg.fn, msg.args)
      .then((result) => sendResponse({ success: true, result }))
      .catch((err)  => sendResponse({ success: false, error: err.message }));
    return true; // Keep the message channel open for async response
  }

  if (msg.type === 'DISPLAY_ALERT') {
    // Replace browser confirm()/alert() with a styled in-page modal so we can
    // show descriptive button labels instead of generic "OK" / "Cancel".
    zoteroDialog(msg.text ?? '', msg.buttons ?? 0)
      .then((result) => sendResponse({ success: true, result }));
    return true; // async
  }

  if (msg.type === 'TOGGLE_TOOLBAR_BUTTONS') {
    if (msg.enabled) {
      injectToolbarButtons();
    } else {
      removeToolbarButtons();
    }
    return;
  }

  if (msg.type === 'POPUP_COMMAND') {
    if (msg.command === 'exportCitations') {
      handleExportCitations();
    } else {
      chrome.runtime.sendMessage({ type: 'ZOTERO_COMMAND', command: msg.command });
    }
    return;
  }

  if (msg.type === 'TRANSACTION_ERROR') {
    console.error('[Zotero] Transaction error:', msg.message);
    zoteroDialog(`Zotero error: ${msg.message}`, 0);
    return;
  }
});

// The iframe is identified lazily from the first message it sends.
// The host page's OPEN_HTML handles the Host_PostmessageReady handshake.
// This script only needs to listen for Document_Loaded and inject buttons.
console.log('[Zotero] content script loaded');
