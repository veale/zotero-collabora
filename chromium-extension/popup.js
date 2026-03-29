// popup.js — Extension popup menu: sends commands to the active tab's content script.

// ── Command buttons ─────────────────────────────────────────────────────────

document.querySelectorAll('button[data-command]').forEach((btn) => {
  btn.addEventListener('click', async () => {
    const command = btn.dataset.command;
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab?.id) return;
    chrome.tabs.sendMessage(tab.id, { type: 'POPUP_COMMAND', command });
    window.close();
  });
});

// ── Force cancel ────────────────────────────────────────────────────────────

document.getElementById('force-cancel').addEventListener('click', async () => {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.id) return;
  chrome.runtime.sendMessage({ type: 'FORCE_CANCEL', tabId: tab.id });
  window.close();
});

// ── Toolbar button toggle ───────────────────────────────────────────────────

const toggle = document.getElementById('toolbar-toggle');

// Load saved preference (default: on)
chrome.storage.local.get({ showToolbarButtons: true }, (data) => {
  toggle.checked = data.showToolbarButtons;
});

toggle.addEventListener('change', async () => {
  const enabled = toggle.checked;
  chrome.storage.local.set({ showToolbarButtons: enabled });

  // Notify the active tab's content script so it can act immediately
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (tab?.id) {
    chrome.tabs.sendMessage(tab.id, {
      type: 'TOGGLE_TOOLBAR_BUTTONS',
      enabled,
    }).catch(() => {});
  }
});
