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

const CALENDAR_TAB = 'Calendar';

/**
 * Clear all cells in a tab, then write new data starting at A1.
 */
async function clearAndWriteTab(spreadsheetId, tabName, rows, token) {
  await ensureTab(spreadsheetId, tabName, token);

  // Clear existing content
  const clearUrl = `${SHEETS_API}/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(tabName)}:clear`;
  await fetch(clearUrl, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}` }
  });

  // Write new data
  await sheetsUpdate(spreadsheetId, `${tabName}!A1`, rows, token);
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

// ── Google Sheet tab helpers ─────────────────────────────────────────────────

/**
 * Ensure a sheet tab exists.
 */
async function ensureTab(spreadsheetId, tabName, token) {
  const metaRes = await fetch(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!metaRes.ok) {
    const err = await metaRes.json().catch(() => ({}));
    throw new Error(err.error?.message || `Could not read spreadsheet: ${metaRes.status}`);
  }
  const meta = await metaRes.json();
  const exists = meta.sheets?.some(s => s.properties?.title === tabName);
  if (!exists) {
    const addRes = await fetch(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}:batchUpdate`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ requests: [{ addSheet: { properties: { title: tabName } } }] })
    });
    if (!addRes.ok) {
      const err = await addRes.json().catch(() => ({}));
      throw new Error(err.error?.message || `Could not create tab "${tabName}": ${addRes.status}`);
    }
  }
}

// ── Calendar builder ─────────────────────────────────────────────────────────

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

/**
 * Build a 2-D calendar grid for the given month.
 * Accepts reservation objects with { name, startDate, endDate }.
 */
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

/**
 * Get array of {year, month} objects from current month to 6 months ahead.
 */
function getMonthRange() {
  const months = [];
  const now = new Date();
  for (let i = 0; i < 7; i++) {
    const d = new Date(now.getFullYear(), now.getMonth() + i, 1);
    months.push({ year: d.getFullYear(), month: d.getMonth() + 1 });
  }
  return months;
}

async function runAutomation() {
  const sourceSheetId = document.getElementById('sourceSheetId').value.trim();
  const sourceRange   = document.getElementById('sourceRange').value.trim() || DEFAULT_SOURCE_RANGE;
  const targetSheetId = document.getElementById('targetSheetId').value.trim();

  if (!targetSheetId) {
    showStatus('Please enter the Target Google Sheet ID.', 'error');
    return;
  }

  if (!sourceSheetId) {
    showStatus('Please enter the Source Google Sheet ID.', 'error');
    return;
  }

  setRunning(true);
  saveSettings();

  try {
    showStatus('Signing in to Google…', 'info');
    const token = await getAuthToken(true);

    // 1. Fetch reservations from Google Sheet
    showStatus('Fetching reservations from Google Sheet…', 'info');
    let reservations = [];
    try {
      reservations = await fetchPhoneReservations(sourceSheetId, sourceRange, token);
      showStatus(`Found ${reservations.length} reservation(s).`, 'info');
    } catch (e) {
      showStatus(`Error: could not load reservations: ${e.message}`, 'error');
      setRunning(false);
      return;
    }

    if (reservations.length === 0) {
      showStatus('No reservations found in source sheet.', 'warning');
      setRunning(false);
      return;
    }

    // 2. Build calendars for all months and write to a single "Calendar" sheet
    const monthRange = getMonthRange();
    showStatus(`Building calendars for ${monthRange.length} months…`, 'info');

    const allCalendarRows = [];
    for (const { year, month } of monthRange) {
      const monthName = new Date(year, month - 1, 1).toLocaleString('default', { month: 'long' });
      // Month separator row
      allCalendarRows.push([`${monthName} ${year}`]);
      const calendar = buildCalendar(reservations, year, month);
      allCalendarRows.push(...calendar);
      allCalendarRows.push([]);  // blank row between months
    }

    showStatus('Writing calendar…', 'info');
    await clearAndWriteTab(targetSheetId, CALENDAR_TAB, allCalendarRows, token);

    showStatus(
      `✅ Done! ${reservations.length} reservations written to Calendar.`,
      'success'
    );
  } catch (e) {
    showStatus(`Error: ${e.message}`, 'error');
  } finally {
    setRunning(false);
  }
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

  document.getElementById('run').addEventListener('click', runAutomation);
});
