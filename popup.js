// ── Constants ─────────────────────────────────────────────────────────────────

const SHEETS_API = 'https://sheets.googleapis.com/v4/spreadsheets';
const DEFAULT_SOURCE_RANGE = 'Sheet1!A2:D';

/** Maximum time (ms) to wait for the reservation .xlsx download to start. */
const DOWNLOAD_TIMEOUT_MS = 60_000;

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

// ── booking.com reservations (download via content script) ───────────────────

/**
 * Ask the content script to navigate the current tab to the booking.com
 * reservation-statements page (with the correct session, date_from, and
 * date_to parameters) and wait for the .xlsx file to be downloaded.
 *
 * Resolves with { downloaded: true, filename } when the download completes,
 * or rejects on error / timeout.
 */
async function triggerReservationDownload(tabId) {
  // 1. Tell the content script to navigate to the download page.
  const navResult = await new Promise((resolve, reject) => {
    chrome.tabs.sendMessage(tabId, { action: 'extractReservations' }, (response) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
      } else {
        resolve(response);
      }
    });
  });

  if (!navResult?.navigating) {
    throw new Error('Content script did not confirm navigation.');
  }

  // 2. Wait 10 seconds for the page to load and settle.
  await new Promise(resolve => setTimeout(resolve, 10_000));

  // 3. Wait for the browser to start a new .xlsx download (up to DOWNLOAD_TIMEOUT_MS).
  //    The content script auto-clicks the download button once the
  //    reservation-statements page has loaded.
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      chrome.downloads.onCreated.removeListener(listener);
      reject(new Error('Timed out waiting for the .xlsx download to start.'));
    }, DOWNLOAD_TIMEOUT_MS);

    function listener(downloadItem) {
      const isXlsx =
        downloadItem.filename?.toLowerCase().endsWith('.xlsx') ||
        downloadItem.mime === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
      // Only match downloads that originate from the booking.com admin domain.
      const isFromBooking = downloadItem.url?.includes('admin.booking.com');
      if (isXlsx && isFromBooking) {
        clearTimeout(timer);
        chrome.downloads.onCreated.removeListener(listener);
        resolve({ downloaded: true, filename: downloadItem.filename });
      }
    }

    chrome.downloads.onCreated.addListener(listener);
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

  setRunning(true);
  saveSettings();

  try {
    showStatus('Signing in to Google…', 'info');
    const token = await getAuthToken(true);

    // 1. Navigate to the booking.com reservation-statements page and wait for
    //    the .xlsx file to download.  The content script handles the navigation
    //    and automatically clicks the download button once the page loads.
    showStatus('Navigating to booking.com reservation statements page…', 'info');
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    let downloadSucceeded = false;
    try {
      const { filename } = await triggerReservationDownload(tab.id);
      const shortName = filename ? filename.split(/[\\/]/).pop() : 'reservations.xlsx';
      showStatus(`✅ Reservation statement downloaded: ${shortName}`, 'success');
      downloadSucceeded = true;
    } catch (e) {
      showStatus(`Warning: could not download reservation statement (${e.message}). Proceeding with phone reservations only.`, 'warning');
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

    if (phoneRes.length === 0) {
      // No phone reservations to build the calendar from.
      // If the xlsx was also not downloaded, surface a clearer message.
      if (!downloadSucceeded) {
        showStatus('No reservations found: booking.com download failed and no phone reservations provided.', 'warning');
      }
      setRunning(false);
      return;
    }

    // 3. Build calendars for each month in range and write to Google Sheet
    const monthRange = getMonthRange();
    showStatus(`Building calendars for ${monthRange.length} months…`, 'info');

    for (const { year, month } of monthRange) {
      const calendar = buildCalendar(phoneRes, year, month);
      const monthName = new Date(year, month - 1, 1).toLocaleString('default', { month: 'short' });
      showStatus(`Writing ${monthName} ${year}…`, 'info');
      await writeCalendarToSheet(targetSheetId, calendar, year, month, token);
    }

    const firstMonth = monthRange[0];
    const lastMonth = monthRange[monthRange.length - 1];
    const rangeStr = `${new Date(firstMonth.year, firstMonth.month - 1).toLocaleString('default', { month: 'short' })} ${firstMonth.year} – ${new Date(lastMonth.year, lastMonth.month - 1).toLocaleString('default', { month: 'short' })} ${lastMonth.year}`;
    showStatus(`✅ Calendars written for ${rangeStr} (${phoneRes.length} phone reservation(s)).`, 'success');
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
