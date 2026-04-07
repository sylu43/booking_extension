// ── Constants ─────────────────────────────────────────────────────────────────

const SHEETS_API = 'https://sheets.googleapis.com/v4/spreadsheets';
const DEFAULT_SOURCE_RANGE = 'Sheet1!A2:D';

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

// ── Google OAuth ──────────────────────────────────────────────────────────────

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

// ── Google Sheets helpers ─────────────────────────────────────────────────────

async function sheetsGet(spreadsheetId, range, token) {
  const url = `${SHEETS_API}/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(range)}`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || `Sheets GET failed: ${res.status}`);
  }
  return res.json();
}

async function sheetsUpdate(spreadsheetId, range, values, token) {
  const url = `${SHEETS_API}/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`;
  const res = await fetch(url, {
    method: 'PUT',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ range, majorDimension: 'ROWS', values })
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || `Sheets PUT failed: ${res.status}`);
  }
  return res.json();
}

/**
 * Clear a sheet tab and write the calendar data starting at A1.
 * The tab is addressed by name; we use "Calendar" and create it if absent.
 */
async function writeCalendarToSheet(spreadsheetId, calendarRows, year, month, token) {
  const tabName = `${year}-${String(month).padStart(2, '0')}`;

  // Ensure the target tab exists (add it if not)
  const metaRes = await fetch(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!metaRes.ok) {
    const err = await metaRes.json().catch(() => ({}));
    throw new Error(err.error?.message || `Could not read spreadsheet: ${metaRes.status}`);
  }
  const meta = await metaRes.json();
  const sheetExists = meta.sheets?.some(s => s.properties?.title === tabName);

  if (!sheetExists) {
    const addRes = await fetch(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}:batchUpdate`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        requests: [{ addSheet: { properties: { title: tabName } } }]
      })
    });
    if (!addRes.ok) {
      const err = await addRes.json().catch(() => ({}));
      throw new Error(err.error?.message || `Could not create sheet tab: ${addRes.status}`);
    }
  }

  // Write data
  const range = `${tabName}!A1`;
  await sheetsUpdate(spreadsheetId, range, calendarRows, token);
}

// ── Phone reservations (from Google Sheet) ────────────────────────────────────

/**
 * Fetch phone/manual reservations from the source Google Sheet.
 * Expected columns: A=Name, B=Check-in, C=Check-out, D=Room (optional)
 */
async function fetchPhoneReservations(sheetId, range, token) {
  const data = await sheetsGet(sheetId, range, token);
  const rows = data.values || [];
  return rows
    .filter(row => row[0] && row[1] && row[2])
    .map(row => ({
      name:      row[0].trim(),
      startDate: row[1].trim(),
      endDate:   row[2].trim(),
      room:      row[3] ? row[3].trim() : '',
      source:    'phone'
    }));
}

// ── booking.com reservations (via content script) ─────────────────────────────

async function extractBookingReservations(tabId) {
  return new Promise((resolve, reject) => {
    chrome.tabs.sendMessage(tabId, { action: 'extractReservations' }, (response) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
      } else {
        resolve(response?.reservations || []);
      }
    });
  });
}

// ── Calendar builder (mirrors logic in content.js, runs in popup context) ─────

function parseDate(dateStr) {
  if (!dateStr) return null;
  const stripped = dateStr.replace(/^\s*(Mon|Tue|Wed|Thu|Fri|Sat|Sun)[a-z]*\.?,?\s+/i, '').trim();
  const native = new Date(stripped);
  if (!isNaN(native)) return new Date(native.getFullYear(), native.getMonth(), native.getDate());

  let m = stripped.match(/^(\d{1,2})\s+([A-Za-z]{3,})\s+(\d{4})$/);
  if (m) {
    const d = new Date(`${m[2]} ${m[1]}, ${m[3]}`);
    if (!isNaN(d)) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  m = stripped.match(/^([A-Za-z]{3,})\s+(\d{1,2}),?\s+(\d{4})$/);
  if (m) {
    const d = new Date(`${m[1]} ${m[2]}, ${m[3]}`);
    if (!isNaN(d)) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }
  return null;
}

function buildCalendar(reservations, year, month) {
  const NUM_ROOMS = 10;
  const daysInMonth = new Date(year, month, 0).getDate();
  const monthStart  = new Date(year, month - 1, 1);
  const monthEnd    = new Date(year, month - 1, daysInMonth);

  const parsed = reservations
    .map(r => ({ ...r, startParsed: parseDate(r.startDate), endParsed: parseDate(r.endDate) }))
    .filter(r => r.startParsed && r.endParsed)
    .filter(r => r.startParsed <= monthEnd && r.endParsed > monthStart);

  parsed.sort((a, b) => a.startParsed - b.startParsed);

  const rooms = Array.from({ length: NUM_ROOMS }, () => []);
  for (const res of parsed) {
    for (let i = 0; i < NUM_ROOMS; i++) {
      const last = rooms[i][rooms[i].length - 1];
      if (!last || last.endParsed <= res.startParsed) {
        rooms[i].push(res);
        break;
      }
    }
  }

  const header = ['Room'];
  for (let d = 1; d <= daysInMonth; d++) header.push(String(d));

  const rows = [header];
  for (let i = 0; i < NUM_ROOMS; i++) {
    const row = [`Room ${i + 1}`];
    for (let d = 1; d <= daysInMonth; d++) {
      const date = new Date(year, month - 1, d);
      let cell = '';
      for (const res of rooms[i]) {
        if (date >= res.startParsed && date < res.endParsed) { cell = res.name || ''; break; }
      }
      row.push(cell);
    }
    rows.push(row);
  }
  return rows;
}

// ── Persistence (save / restore settings) ────────────────────────────────────

function saveSettings() {
  chrome.storage.local.set({
    sourceSheetId: document.getElementById('sourceSheetId').value,
    sourceRange:   document.getElementById('sourceRange').value,
    targetSheetId: document.getElementById('targetSheetId').value
  });
}

function loadSettings() {
  chrome.storage.local.get(['sourceSheetId', 'sourceRange', 'targetSheetId'], (data) => {
    if (data.sourceSheetId) document.getElementById('sourceSheetId').value = data.sourceSheetId;
    if (data.sourceRange)   document.getElementById('sourceRange').value   = data.sourceRange;
    if (data.targetSheetId) document.getElementById('targetSheetId').value = data.targetSheetId;
  });
}

// ── Main workflow ─────────────────────────────────────────────────────────────

async function runAutomation() {
  const year    = parseInt(document.getElementById('year').value, 10);
  const month   = parseInt(document.getElementById('month').value, 10);
  const sourceSheetId = document.getElementById('sourceSheetId').value.trim();
  const sourceRange   = document.getElementById('sourceRange').value.trim() || DEFAULT_SOURCE_RANGE;
  const targetSheetId = document.getElementById('targetSheetId').value.trim();

  if (!targetSheetId) {
    showStatus('Please enter the Target Google Sheet ID.', 'error');
    return;
  }

  setRunning(true);
  saveSettings();

  try {
    showStatus('Signing in to Google…', 'info');
    const token = await getAuthToken(true);

    // 1. Extract booking.com reservations from the active tab
    showStatus('Extracting booking.com reservations…', 'info');
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    let bookingRes = [];
    try {
      bookingRes = await extractBookingReservations(tab.id);
      showStatus(`Found ${bookingRes.length} booking.com reservation(s).`, 'info');
    } catch (e) {
      showStatus(`Warning: could not extract from booking.com page (${e.message}). Proceeding with phone reservations only.`, 'warning');
    }

    // 2. Fetch phone reservations (optional)
    let phoneRes = [];
    if (sourceSheetId) {
      showStatus('Fetching phone reservations from Google Sheet…', 'info');
      try {
        phoneRes = await fetchPhoneReservations(sourceSheetId, sourceRange, token);
        showStatus(`Found ${phoneRes.length} phone reservation(s).`, 'info');
      } catch (e) {
        showStatus(`Warning: could not load phone reservations: ${e.message}`, 'warning');
      }
    }

    const allReservations = [...bookingRes, ...phoneRes];
    if (allReservations.length === 0) {
      showStatus('No reservations found for the selected period.', 'warning');
      setRunning(false);
      return;
    }

    // 3. Build calendar
    showStatus('Building calendar…', 'info');
    const calendar = buildCalendar(allReservations, year, month);

    // 4. Write to Google Sheet
    showStatus('Writing calendar to Google Sheet…', 'info');
    await writeCalendarToSheet(targetSheetId, calendar, year, month, token);

    const monthName = new Date(year, month - 1, 1).toLocaleString('default', { month: 'long' });
    showStatus(`✅ Calendar written for ${monthName} ${year} (${allReservations.length} reservations).`, 'success');
  } catch (e) {
    showStatus(`Error: ${e.message}`, 'error');
  } finally {
    setRunning(false);
  }
}

// ── Init ──────────────────────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', () => {
  // Default month/year to current
  const now = new Date();
  document.getElementById('year').value  = now.getFullYear();
  document.getElementById('month').value = now.getMonth() + 1;

  loadSettings();

  document.getElementById('authenticate').addEventListener('click', async () => {
    try {
      await getAuthToken(true);
      showStatus('Signed in successfully.', 'success');
    } catch (e) {
      showStatus(`Sign-in failed: ${e.message}`, 'error');
    }
  });

  document.getElementById('run').addEventListener('click', runAutomation);
});
