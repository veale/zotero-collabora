// background.js — Service worker managing the Zotero HTTP transaction loop.
//
// Architecture:
//   Content script  ──postMessage──►  Collabora iframe  (Python UNO)
//   Content script  ──runtime msg──►  Background (this file)
//   Background      ──fetch()──────►  Zotero desktop  :23119
//
// Because fetch() to 127.0.0.1 requires host_permissions and is blocked in
// content-script context (CORS), ALL Zotero HTTP calls go through this worker.
// The content script relays PostMessage results back here via chrome.runtime.

// ── Config ────────────────────────────────────────────────────────────────────

const ZOTERO_BASE = 'http://127.0.0.1:23119';
const ZOTERO_HEADERS = {
  'Content-Type': 'application/json',
  'X-Zotero-Version': '6.0',
  'X-Zotero-Connector-API-Version': '3',
};

// ── State ─────────────────────────────────────────────────────────────────────

// Map tabId → { docId, busy, cancelled }
const tabState = {};

// Pending resolve/reject for the content-script command round-trip.
// Only one transaction runs per tab at a time so a single slot is fine.
const pendingCommands = {};   // tabId → { resolve, reject }

// ── Zotero HTTP helpers ───────────────────────────────────────────────────────

// Hard timeout on every Zotero HTTP call so a non-responsive Zotero
// doesn't freeze the extension indefinitely.
const ZOTERO_FETCH_TIMEOUT_MS = 12_000;

async function zoteroPost(path, body) {
  const ctrl = new AbortController();
  const tid  = setTimeout(() => ctrl.abort(), ZOTERO_FETCH_TIMEOUT_MS);

  let resp;
  try {
    resp = await fetch(`${ZOTERO_BASE}${path}`, {
      method: 'POST',
      headers: ZOTERO_HEADERS,
      body: JSON.stringify(body),
      signal: ctrl.signal,
    });
  } catch (err) {
    clearTimeout(tid);
    if (err.name === 'AbortError')
      throw new Error('Zotero is not responding (timeout). Is Zotero running with HTTP integration enabled?');
    throw new Error(`Cannot reach Zotero: ${err.message}`);
  }
  clearTimeout(tid);

  if (resp.status === 503) {
    if (path === '/connector/document/execCommand') {
      // A previous transaction may be stuck on Zotero's side.
      // Send an abort signal to clear it, then retry once.
      console.warn('[Zotero] 503 on execCommand — aborting stuck transaction and retrying...');
      try {
        await fetch(`${ZOTERO_BASE}/connector/document/respond`, {
          method: 'POST', headers: ZOTERO_HEADERS,
          body: JSON.stringify({ error: 'ScriptError', message: 'Clearing stuck transaction' }),
          signal: AbortSignal.timeout(4000),
        });
      } catch {}
      await new Promise(r => setTimeout(r, 400));
      try {
        const retry = await fetch(`${ZOTERO_BASE}${path}`, {
          method: 'POST', headers: ZOTERO_HEADERS,
          body: JSON.stringify(body),
          signal: AbortSignal.timeout(8000),
        });
        if (retry.ok) return retry.json();
      } catch {}
    }
    throw new Error('Zotero is busy. Use "Force Cancel" in the extension popup if it stays stuck.');
  }
  if (resp.status === 404) throw new Error('Zotero version too old or HTTP integration disabled.');
  if (!resp.ok) throw new Error(`Zotero HTTP ${resp.status}`);
  return resp.json();
}

// Send an abort to Zotero without throwing — used by force-cancel.
function zoteroAbortSilently() {
  fetch(`${ZOTERO_BASE}/connector/document/respond`, {
    method: 'POST', headers: ZOTERO_HEADERS,
    body: JSON.stringify({ error: 'ScriptError', message: 'Transaction cancelled by user' }),
    signal: AbortSignal.timeout(4000),
  }).catch(() => {});
}

export async function checkZoteroAlive() {
  try {
    const resp = await fetch(`${ZOTERO_BASE}/connector/ping`);
    return resp.ok;
  } catch {
    return false;
  }
}

// ── Document-capability response (Application.getActiveDocument) ──────────────

function buildDocumentInfo(docId) {
  return {
    documentID: docId,
    outputFormat: 'html',
    supportedNotes: ['footnotes', 'endnotes'],
    supportsImportExport: true,
    supportsTextInsertion: true,
    supportsCitationMerging: true,
    processorName: 'Collabora Online',
  };
}

// ── WP-command → Python function mapping ─────────────────────────────────────

// Returns { pythonFn, args } or null for commands handled directly in JS.
function mapToPython(command, cmdArgs) {
  const [a0, a1, a2, a3] = cmdArgs || [];
  switch (command) {
    case 'Document.getDocumentData':
      return { fn: 'getDocumentData', args: {} };
    case 'Document.setDocumentData':
      return { fn: 'setDocumentData', args: { data: a0 } };
    case 'Document.cursorInField':
      return { fn: 'cursorInField', args: { fieldType: a0 } };
    case 'Document.insertField':
      return { fn: 'insertField', args: { fieldType: a0, noteType: a1 ?? 0 } };
    case 'Document.getFields':
      return { fn: 'getFields', args: { fieldType: a0 } };
    case 'Document.insertText':
      return { fn: 'insertText', args: { html: a0 } };
    case 'Document.setBibliographyStyle':
      return {
        fn: 'setBibliographyStyle',
        args: {
          firstLineIndent: a0, bodyIndent: a1, lineSpacing: a2,
          entrySpacing: a3, tabStops: cmdArgs[4], count: cmdArgs[5],
        },
      };
    case 'Document.convert':
      // fieldIDs and toNoteTypes are arrays; JSON.stringify so Python receives
      // valid JSON strings — toTypedArgs coerces raw arrays via String(), which
      // produces comma-separated values (not valid JSON).
      return {
        fn: 'convertFields',
        args: {
          fieldIDs:    JSON.stringify(a0),
          toFieldType: a1,
          toNoteTypes: JSON.stringify(a2),
          count: a3,
        },
      };
    case 'Document.convertPlaceholdersToFields':
      return {
        fn: 'convertPlaceholdersToFields',
        args: {
          placeholderIDs: JSON.stringify(a0),
          noteType: a1,
          fieldType: a2,
        },
      };
    case 'Field.setCode':
      // a0 may be null when Zotero means "the field just inserted" — caller substitutes
      return { fn: 'setFieldCode', args: { fieldID: a0, code: a1 } };
    case 'Field.setText':
      return { fn: 'setFieldText', args: { fieldID: a0, text: a1, isRich: a2 } };
    case 'Field.getText':
      // Read from getFields result; handle via a dedicated helper
      return { fn: 'getFieldText', args: { fieldID: a0 } };
    case 'Field.delete':
      return { fn: 'deleteField', args: { fieldID: a0 } };
    case 'Field.select':
      return { fn: 'selectField', args: { fieldID: a0 } };
    case 'Field.removeCode':
      return { fn: 'removeFieldCode', args: { fieldID: a0 } };
    case 'Field.getCode':
      return { fn: 'getFieldCode', args: { fieldID: a0 } };
    case 'Field.getNoteIndex':
      return { fn: 'getFieldNoteIndex', args: { fieldID: a0 } };
    default:
      return null;
  }
}

// ── Command result normalisation ──────────────────────────────────────────────

// Zotero expects specific return shapes for each command.
// The Python scripts return JSON strings; we parse and reshape here.
function normaliseResult(command, raw) {
  if (raw === null || raw === undefined) return null;

  let parsed;
  try {
    parsed = typeof raw === 'string' ? JSON.parse(raw) : raw;
  } catch {
    parsed = raw;
  }

  switch (command) {
    case 'Document.cursorInField': {
      if (!parsed) return null;
      return {
        id:        parsed.fieldID,
        code:      parsed.fieldCode ?? '',
        text:      parsed.fieldText ?? '',
        noteIndex: parsed.noteIndex ?? 0,
      };
    }
    case 'Document.insertField': {
      // Return \u200b so the field's cached getText() is non-empty.
      // Zotero's ignoreEmptyBibliography check does getText().trim() === ""
      // and JS trim() does NOT strip zero-width spaces, so the bibliography
      // field survives instead of being flagged for removal.
      return {
        id:        parsed.fieldID,
        code:      parsed.fieldCode ?? '',
        text:      '\u200b',
        noteIndex: parsed.noteIndex ?? 0,
      };
    }
    case 'Document.getFields': {
      return (parsed.fieldIDs || []).map((id, i) => ({
        id,
        code:      (parsed.fieldCodes  || [])[i] ?? '',
        text:      (parsed.fieldTexts  || [])[i] ?? '',
        noteIndex: (parsed.noteIndices || [])[i] ?? 0,
      }));
    }
    case 'Document.convertPlaceholdersToFields': {
      return (parsed.fieldIDs || []).map((id, i) => ({
        id,
        code:      (parsed.fieldCodes  || [])[i] ?? '',
        text:      (parsed.fieldTexts  || [])[i] ?? '',
        noteIndex: (parsed.noteIndices || [])[i] ?? 0,
      }));
    }
    case 'Document.canInsertField':
      return true;
    case 'Document.activate':
    case 'Document.setDocumentData':
    case 'Document.setBibliographyStyle':
    case 'Document.convert':
    case 'Document.insertText':
    case 'Field.setCode':
    case 'Field.setText':
    case 'Field.delete':
    case 'Field.select':
    case 'Field.removeCode':
      return null;
    case 'Field.getText':
    case 'Field.getCode':
      return typeof parsed === 'string' ? parsed : JSON.stringify(parsed);
    case 'Field.getNoteIndex':
      return typeof parsed === 'number' ? parsed : parseInt(parsed) || 0;
    case 'Document.getDocumentData':
      return typeof parsed === 'string' ? parsed : '';
    case 'Document.displayAlert':
      return 0;   // 0 = first button (OK) — alert shown by browser, not here
    default:
      return parsed;
  }
}

// ── Execute a single WP command ───────────────────────────────────────────────

async function executeCommand(tabId, docId, wpCommand) {
  // If the user force-cancelled, bail out immediately.
  if (tabState[tabId]?.cancelled) {
    throw new Error('Transaction cancelled by user.');
  }

  const { command, arguments: allArgs } = wpCommand;
  // Zotero always passes docId as arguments[0]; real args start at index 1.
  let cmdArgs = allArgs?.slice(1) ?? [];

  // Commands handled entirely in the background worker (no Python needed)
  if (command === 'Application.getActiveDocument') {
    return buildDocumentInfo(docId);
  }
  if (command === 'Document.activate') return null;
  if (command === 'Document.canInsertField') return true;
  if (command === 'Document.displayAlert') {
    // cmdArgs: [message, icon, buttons]  (docId already stripped from allArgs)
    // Must await user response — Zotero uses the return value to branch.
    return new Promise((resolve) => {
      chrome.tabs.sendMessage(tabId, {
        type: 'DISPLAY_ALERT',
        text:    cmdArgs[0] ?? '',
        icon:    cmdArgs[1] ?? 0,
        buttons: cmdArgs[2] ?? 0,
      }, (response) => {
        resolve(response?.result ?? 1);
      });
    });
  }

  // When Zotero sends Field.setCode with null fieldID it means the field just
  // inserted via insertField.  Substitute the tracked fieldID from tabState.
  if (command === 'Field.setCode' && cmdArgs[0] == null) {
    const fallback = tabState[tabId]?.lastInsertedFieldId;
    if (fallback) {
      cmdArgs = [fallback, ...cmdArgs.slice(1)];
    } else {
      console.warn('[Zotero] Field.setCode: null fieldID and no lastInsertedFieldId tracked');
    }
  }

  const mapped = mapToPython(command, cmdArgs);
  if (!mapped) {
    console.warn(`[Zotero] Unhandled WP command: ${command}`);
    return null;
  }

  // Delegate to content script → PostMessage → Python
  const raw = await sendToContentScript(tabId, mapped.fn, mapped.args);
  const result = normaliseResult(command, raw);

  // Track the fieldID returned by insertField so Field.setCode can use it
  if (command === 'Document.insertField' && result?.id) {
    tabState[tabId].lastInsertedFieldId = result.id;
  }

  return result;
}

// ── Content-script round-trip ─────────────────────────────────────────────────

// Timeout for a single Python UNO round-trip (content script → Collabora → Python).
const PYTHON_ROUNDTRIP_TIMEOUT_MS = 20_000;

function sendToContentScript(tabId, pythonFn, args) {
  return new Promise((resolve, reject) => {
    const tid = setTimeout(() => {
      delete pendingCommands[tabId];
      reject(new Error(`Python call timed out after ${PYTHON_ROUNDTRIP_TIMEOUT_MS / 1000}s: ${pythonFn}`));
    }, PYTHON_ROUNDTRIP_TIMEOUT_MS);

    const done = (fn, val) => { clearTimeout(tid); delete pendingCommands[tabId]; fn(val); };
    pendingCommands[tabId] = {
      resolve: (v) => done(resolve, v),
      reject:  (e) => done(reject,  e),
    };

    chrome.tabs.sendMessage(tabId, { type: 'CALL_PYTHON', fn: pythonFn, args }, (response) => {
      if (chrome.runtime.lastError) {
        done(reject, new Error(chrome.runtime.lastError.message));
        return;
      }
      if (response?.success) {
        done(resolve, response.result);
      } else {
        done(reject, new Error(response?.error ?? 'Unknown error from content script'));
      }
    });
  });
}

// ── Main transaction loop ─────────────────────────────────────────────────────

async function runTransaction(tabId, zoteroCommand) {
  if (tabState[tabId]?.busy) {
    console.warn('[Zotero] Transaction already in progress for tab', tabId);
    return;
  }

  const docId = tabState[tabId]?.docId ?? `collabora-${tabId}`;
  tabState[tabId] = { docId, busy: true, cancelled: false };

  try {
    // Step 1: Start the transaction
    let wpCommand = await zoteroPost('/connector/document/execCommand', {
      command: zoteroCommand,
      docId,
    });
    console.log('[Zotero] execCommand response:', JSON.stringify(wpCommand));

    // Step 2: Command loop
    while (true) {
      if (wpCommand.command === 'Document.complete') break;

      let result;
      try {
        result = await executeCommand(tabId, docId, wpCommand);
      } catch (err) {
        // Report failure to Zotero so it can exit the transaction cleanly
        try {
          await zoteroPost('/connector/document/respond', {
            error: 'ScriptError',
            message: err.message,
            stack: err.stack ?? '',
          });
        } catch (respondErr) {
          console.error('[Zotero] Failed to send error to Zotero:', respondErr);
        }
        throw err;
      }

      console.log('[Zotero] cmd:', wpCommand.command,
        'args:', JSON.stringify(wpCommand.arguments),
        '→ result:', JSON.stringify(result));

      try {
        wpCommand = await zoteroPost('/connector/document/respond', result ?? null);
      } catch (respondErr) {
        console.error('[Zotero] Failed to send respond:', respondErr);
        throw respondErr;
      }
      console.log('[Zotero] next command:', JSON.stringify(wpCommand));
    }
  } catch (err) {
    // Last-resort: if Zotero is still waiting, try to unstick it
    try {
      await zoteroPost('/connector/document/respond', {
        error: 'ScriptError',
        message: `Transaction aborted: ${err.message}`,
      });
    } catch {
      // Zotero may have already closed the transaction — that's fine
    }

    // Clean up any orphan bookmark left by a partially-completed insertField.
    // Without this, a cancelled citation leaves a zero-code/zero-text bookmark
    // that confuses subsequent getFields calls (especially bibliography).
    const orphanId = tabState[tabId]?.lastInsertedFieldId;
    if (orphanId) {
      chrome.tabs.sendMessage(tabId, {
        type: 'CALL_PYTHON',
        fn: 'deleteField',
        args: { fieldID: orphanId },
      }).catch(() => {});
    }

    throw err;
  } finally {
    if (tabState[tabId]) {
      tabState[tabId].busy = false;
      tabState[tabId].lastInsertedFieldId = null;
    }
  }
}

// ── Message handler (from content scripts) ────────────────────────────────────

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  const tabId = sender.tab?.id;
  if (!tabId) return;

  if (msg.type === 'ZOTERO_COMMAND') {
    // Content script button click → start a transaction
    runTransaction(tabId, msg.command).catch((err) => {
      console.error('[Zotero] Transaction failed:', err);
      chrome.tabs.sendMessage(tabId, {
        type: 'TRANSACTION_ERROR',
        message: err.message,
      }).catch(() => {});
    });
    return; // no sendResponse needed
  }

  if (msg.type === 'PYTHON_RESULT') {
    // Content script returns a Python call result to unblock sendToContentScript
    const pending = pendingCommands[tabId];
    if (pending) {
      delete pendingCommands[tabId];
      if (msg.success) {
        pending.resolve(msg.result);
      } else {
        pending.reject(new Error(msg.error ?? 'Python error'));
      }
    }
    return;
  }
});

// ── Force-cancel handler (from popup, no sender.tab) ─────────────────────────

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  if (msg.type !== 'FORCE_CANCEL') return;

  const tabId = msg.tabId;

  // Mark the transaction as cancelled so the next executeCommand bails out.
  if (tabState[tabId]) {
    tabState[tabId].cancelled = true;
    tabState[tabId].busy      = false;
  }

  // Unblock any Python round-trip that is currently waiting.
  const pending = pendingCommands[tabId];
  if (pending) {
    delete pendingCommands[tabId];
    pending.reject(new Error('Transaction cancelled by user.'));
  }

  // Tell Zotero to exit its transaction too.
  zoteroAbortSilently();

  sendResponse({ success: true });
  return true;
});