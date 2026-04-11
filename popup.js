// ── UI helpers ────────────────────────────────────────────────────────────────

function showStatus(message, type = 'info') {
  const el = document.getElementById('status');
  el.textContent = message;
  el.className = type;
}

function setRunning(isRunning) {
  document.getElementById('run').disabled = isRunning;
  document.getElementById('authenticate').disabled = isRunning;
}

// ── Status listener (from background service worker) ──────────────────────────

chrome.runtime.onMessage.addListener((msg) => {
  if (msg.action === 'statusUpdate') {
    showStatus(msg.message, msg.type);
    if (msg.final) setRunning(false);
  }
});

// ── Google OAuth (for manual sign-in button) ──────────────────────────────────

function getAuthToken(interactive = true) {
  return new Promise((resolve, reject) => {
    chrome.identity.getAuthToken({ interactive }, (token) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
      } else {
        resolve(token);
      }
    });
  });
}

// ── Persistence ───────────────────────────────────────────────────────────────

function saveSettings() {
  const hoursStr = document.getElementById('scheduleHours').value.trim();
  const scheduleHours = hoursStr
    ? hoursStr.split(',').map(h => parseInt(h.trim())).filter(h => !isNaN(h) && h >= 0 && h <= 23)
    : [];

  chrome.storage.local.set({
    targetSheetId: document.getElementById('targetSheetId').value,
    autoRearrange: document.getElementById('autoRearrange').checked,
    scheduleEnabled: document.getElementById('scheduleEnabled').checked,
    scheduleHours
  });

  chrome.runtime.sendMessage({ action: 'updateSchedule' });
}

function loadSettings() {
  chrome.storage.local.get(
    ['targetSheetId', 'autoRearrange', 'scheduleEnabled', 'scheduleHours', 'lastSyncTime', 'lastSyncResult'],
    (data) => {
      if (data.targetSheetId) document.getElementById('targetSheetId').value = data.targetSheetId;
      if (data.autoRearrange != null) document.getElementById('autoRearrange').checked = data.autoRearrange;
      if (data.scheduleEnabled != null) document.getElementById('scheduleEnabled').checked = data.scheduleEnabled;
      if (data.scheduleHours?.length) {
        document.getElementById('scheduleHours').value = data.scheduleHours.join(', ');
      }
      updateScheduleVisibility();
      showLastSync(data.lastSyncTime, data.lastSyncResult);
    }
  );
}

function updateScheduleVisibility() {
  const enabled = document.getElementById('scheduleEnabled').checked;
  document.getElementById('scheduleConfig').style.display = enabled ? 'block' : 'none';
}

function showLastSync(time, result) {
  const el = document.getElementById('lastSync');
  if (!time) { el.textContent = ''; return; }
  const d = new Date(time);
  const timeStr = d.toLocaleString();
  el.textContent = `Last sync: ${timeStr}`;
}

// ── Init ──────────────────────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', () => {
  loadSettings();

  document.getElementById('authenticate').addEventListener('click', async () => {
    try {
      await getAuthToken(true);
      showStatus('Signed in successfully.', 'success');
    } catch (e) {
      showStatus(`Sign-in failed: ${e.message}`, 'error');
    }
  });

  document.getElementById('run').addEventListener('click', () => {
    const targetSheetId = document.getElementById('targetSheetId').value.trim();
    if (!targetSheetId) {
      showStatus('Please enter the Target Google Sheet ID.', 'error');
      return;
    }
    setRunning(true);
    saveSettings();
    chrome.runtime.sendMessage({
      action: 'runAutomation',
      targetSheetId,
      autoRearrange: document.getElementById('autoRearrange').checked
    });
  });

  document.getElementById('scheduleEnabled').addEventListener('change', () => {
    updateScheduleVisibility();
    saveSettings();
  });

  document.getElementById('scheduleHours').addEventListener('change', saveSettings);
});
