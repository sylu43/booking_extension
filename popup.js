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

// ── booking.com reservations (fetch report binary via content script) ─────────

/**
 * Ask the content script to fetch the latest report as binary XLS data.
 * Returns the base64-encoded file contents.
 */
async function fetchReportBinary(tabId) {
  return new Promise((resolve, reject) => {
    chrome.tabs.sendMessage(tabId, { action: 'fetchLatestReportBinary' }, (response) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
      } else if (!response?.success) {
        reject(new Error(response?.error || 'Failed to fetch report'));
      } else {
        resolve({ base64: response.base64, reportName: response.reportName });
      }
    });
  });
}

// ── XLS parsing (using SheetJS) ──────────────────────────────────────────────

/** Column indices in the booking.com XLS export. */
const XLS_COLS = {
  BOOK_NUMBER: 0,
  BOOKED_BY: 1,
  GUEST_NAME: 2,
  CHECK_IN: 3,
  CHECK_OUT: 4,
  BOOKED_ON: 5,
  STATUS: 6,
  ROOMS: 7,
  PEOPLE: 8,
  ADULTS: 9,
  CHILDREN: 10,
  CHILDREN_AGES: 11,
  PRICE: 12,
  COMMISSION_PCT: 13,
  COMMISSION_AMT: 14,
  PAYMENT_STATUS: 15,
  PAYMENT_METHOD: 16,
  REMARKS: 17,
  BOOKER_COUNTRY: 18,
  TRAVEL_PURPOSE: 19,
  DEVICE: 20,
  UNIT_TYPE: 21,
  DURATION: 22,
  CANCELLATION_DATE: 23,
  ADDRESS: 24,
  PHONE: 25
};

/**
 * Convert a cell value to a stable string (avoids scientific notation for large numbers).
 */
function cellToString(val) {
  if (val == null) return '';
  if (typeof val === 'number') {
    // Avoid scientific notation for large booking numbers
    return Number.isInteger(val) ? val.toFixed(0) : String(val);
  }
  return String(val).trim();
}

/**
 * Parse a base64-encoded XLS/XLSX file into { headers: string[], rows: string[][] }.
 * Each row is an array of string cell values.
 */
function parseXlsBase64(base64) {
  const binaryStr = atob(base64);
  const wb = XLSX.read(binaryStr, { type: 'binary' });
  const sheetName = wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  // Get raw 2D array (header: 1 means row 0 is included as data, not keys)
  const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  if (aoa.length === 0) return { headers: [], rows: [] };

  const headers = aoa[0].map(cellToString);
  const rows = aoa.slice(1)
    .filter(r => r.some(c => c !== '' && c != null))   // skip empty rows
    .map(r => r.map(cellToString));

  return { headers, rows };
}

/**
 * Convert parsed XLS rows into reservation objects for buildCalendar.
 * Filters out cancelled reservations.
 */
function xlsRowsToReservations(rows) {
  return rows
    .filter(r => r[XLS_COLS.STATUS] === 'ok')
    .map(r => ({
      orderNumber: r[XLS_COLS.BOOK_NUMBER],
      name:        r[XLS_COLS.GUEST_NAME] || r[XLS_COLS.BOOKED_BY],
      startDate:   r[XLS_COLS.CHECK_IN],
      endDate:     r[XLS_COLS.CHECK_OUT],
      unitType:    r[XLS_COLS.UNIT_TYPE],
      rooms:       r[XLS_COLS.ROOMS],
      source:      'booking'
    }));
}

// ── Google Sheet upsert (by order number) ────────────────────────────────────

const RESERVATIONS_TAB = 'Reservations';

/**
 * Ensure a sheet tab exists. Returns the sheet metadata.
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

/**
 * Upsert XLS reservation rows into the "Reservations" tab of the target sheet.
 * Uses the Book Number (column A) as the unique key.
 *
 * Strategy:
 *  1. Read all existing rows from the tab.
 *  2. Build a map: orderNumber → row data.
 *  3. Merge incoming XLS rows (insert new, update existing).
 *  4. Write the entire dataset back as a single PUT.
 */
async function upsertReservationsToSheet(spreadsheetId, headers, newRows, token) {
  await ensureTab(spreadsheetId, RESERVATIONS_TAB, token);

  // Read existing data (skip header)
  let existingRows = [];
  try {
    const data = await sheetsGet(spreadsheetId, `${RESERVATIONS_TAB}!A2:Z`, token);
    existingRows = data.values || [];
  } catch { /* tab is empty or doesn't have data yet */ }

  // Build map from order number → row data (preserve insertion order)
  const rowMap = new Map();
  for (const row of existingRows) {
    const key = (row[0] || '').toString().trim();
    if (key) rowMap.set(key, row);
  }

  // Merge new rows (upsert)
  let updatedCount = 0;
  let insertedCount = 0;
  for (const row of newRows) {
    const key = (row[0] || '').toString().trim();
    if (!key) continue;
    if (rowMap.has(key)) {
      updatedCount++;
    } else {
      insertedCount++;
    }
    rowMap.set(key, row);
  }

  // Convert map back to array and write with header
  const allRows = [headers, ...rowMap.values()];
  await sheetsUpdate(spreadsheetId, `${RESERVATIONS_TAB}!A1`, allRows, token);

  return { updatedCount, insertedCount, totalCount: rowMap.size };
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

  setRunning(true);
  saveSettings();

  try {
    showStatus('Signing in to Google…', 'info');
    const token = await getAuthToken(true);

    // 1. Fetch the booking.com XLS report as binary via content script.
    showStatus('Fetching booking.com report data…', 'info');
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    let xlsReservations = [];
    let xlsHeaders = [];
    let xlsRows = [];

    try {
      const { base64, reportName } = await fetchReportBinary(tab.id);
      showStatus(`Parsing report: ${reportName}…`, 'info');

      const parsed = parseXlsBase64(base64);
      xlsHeaders = parsed.headers;
      xlsRows = parsed.rows;

      // Upsert all XLS rows to the Reservations tab (keyed by Book Number).
      showStatus(`Upserting ${xlsRows.length} rows to Reservations sheet…`, 'info');
      const { updatedCount, insertedCount, totalCount } =
        await upsertReservationsToSheet(targetSheetId, xlsHeaders, xlsRows, token);
      showStatus(
        `Reservations upserted: ${insertedCount} new, ${updatedCount} updated (${totalCount} total).`,
        'info'
      );

      // Convert active (non-cancelled) XLS rows to reservation objects for the calendar.
      xlsReservations = xlsRowsToReservations(xlsRows);
      showStatus(`${xlsReservations.length} active booking.com reservation(s) for calendar.`, 'info');
    } catch (e) {
      showStatus(`Warning: could not fetch/parse booking.com report (${e.message}). Proceeding with phone reservations only.`, 'warning');
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

    // 3. Combine all reservations for calendar building
    const allReservations = [...xlsReservations, ...phoneRes];

    if (allReservations.length === 0) {
      showStatus('No active reservations found from booking.com or phone sheet.', 'warning');
      setRunning(false);
      return;
    }

    // 4. Build calendars for all months and write to a single "Calendar" sheet
    const monthRange = getMonthRange();
    showStatus(`Building calendars for ${monthRange.length} months…`, 'info');

    const allCalendarRows = [];
    for (const { year, month } of monthRange) {
      const monthName = new Date(year, month - 1, 1).toLocaleString('default', { month: 'long' });
      // Month separator row
      allCalendarRows.push([`${monthName} ${year}`]);
      const calendar = buildCalendar(allReservations, year, month);
      allCalendarRows.push(...calendar);
      allCalendarRows.push([]);  // blank row between months
    }

    showStatus('Writing calendar…', 'info');
    await clearAndWriteTab(targetSheetId, CALENDAR_TAB, allCalendarRows, token);

    showStatus(
      `✅ Done! ${xlsReservations.length} booking.com + ${phoneRes.length} phone reservations written to Calendar.`,
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
