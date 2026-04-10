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

// ── Booking.com API helpers ──────────────────────────────────────────────────

/**
 * Send a message to the content script on an active booking.com tab.
 * Returns the response from the content script.
 */
async function sendToBookingTab(message) {
  const tabs = await chrome.tabs.query({ url: 'https://admin.booking.com/*' });
  if (tabs.length === 0) {
    throw new Error('No booking.com admin tab found. Please open booking.com admin first.');
  }
  return new Promise((resolve, reject) => {
    chrome.tabs.sendMessage(tabs[0].id, message, (response) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
      } else if (!response?.success) {
        reject(new Error(response?.error || 'Unknown error'));
      } else {
        resolve(response);
      }
    });
  });
}

/**
 * Format a Date as YYYY-MM-DD (local time).
 */
function formatISODate(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

/**
 * Fetch reservations from booking.com via content script.
 * Uses the retrieve_list_v2 API endpoint.
 */
async function fetchBookingComReservations() {
  const today = new Date();
  const thirtyDaysLater = new Date(today);
  thirtyDaysLater.setDate(today.getDate() + 30);

  const dateFrom = formatISODate(today);
  const dateTo = formatISODate(thirtyDaysLater);

  const response = await sendToBookingTab({
    action: 'fetchBookingReservations',
    dateFrom,
    dateTo
  });

  return response.reservations || [];
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
const BOOKING_TAG = ' [B]'; // Suffix marks booking.com entries

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

// ── Read existing calendar ───────────────────────────────────────────────────

/**
 * Read the existing Calendar tab from Google Sheets.
 * Returns raw row data or empty array if the tab doesn't exist yet.
 */
async function readExistingCalendar(spreadsheetId, token) {
  await ensureTab(spreadsheetId, CALENDAR_TAB, token);
  try {
    const data = await sheetsGet(spreadsheetId, CALENDAR_TAB, token);
    return data.values || [];
  } catch {
    return [];
  }
}

/**
 * Parse existing calendar sheet rows into room assignments per month.
 * Returns a map: "YYYY-M" => { roomNumber: [{name, startDay, endDay}] }
 * where startDay/endDay are 1-based day-of-month numbers (endDay is inclusive).
 */
function parseExistingAssignments(sheetRows) {
  const result = {};
  let i = 0;
  while (i < sheetRows.length) {
    const row = sheetRows[i] || [];
    const monthMatch = (row[0] || '').match(/^([A-Za-z]+)\s+(\d{4})$/);
    if (monthMatch) {
      const monthName = monthMatch[1];
      const year = parseInt(monthMatch[2]);
      const monthDate = new Date(`${monthName} 1, ${year}`);
      if (isNaN(monthDate)) { i++; continue; }
      const month = monthDate.getMonth() + 1;
      const key = `${year}-${month}`;

      i++; // header row
      if (i >= sheetRows.length) break;
      const header = sheetRows[i] || [];
      const numDays = header.length - 1;

      const monthAssignments = {};
      i++; // first room row
      while (i < sheetRows.length) {
        const roomRow = sheetRows[i] || [];
        if (!roomRow[0] || roomRow.length <= 1) { i++; break; }

        const roomNumber = (roomRow[0] || '').match(/^(\d+)/)?.[1];
        if (!roomNumber) { i++; continue; }

        const blocks = [];
        let col = 1;
        while (col <= numDays) {
          const name = (roomRow[col] || '').trim();
          if (name) {
            const startDay = col;
            let endDay = col;
            while (endDay + 1 <= numDays && (roomRow[endDay + 1] || '').trim() === name) {
              endDay++;
            }
            const isBooking = name.endsWith(BOOKING_TAG);
            const cleanName = isBooking ? name.slice(0, -BOOKING_TAG.length) : name;
            blocks.push({ name: cleanName, startDay, endDay, isBooking });
            col = endDay + 1;
          } else {
            col++;
          }
        }
        if (blocks.length > 0) {
          monthAssignments[roomNumber] = blocks;
        }
        i++;
      }

      result[key] = monthAssignments;
    } else {
      i++;
    }
  }
  return result;
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

// Room definitions with number and type
const ROOMS = [
  { number: '201', type: 'double', floor: 2 },
  { number: '202', type: 'quadruple', floor: 2 },
  { number: '203', type: 'double', floor: 2 },
  { number: '205', type: 'double', floor: 2 },
  { number: '206', type: 'quadruple', floor: 2 },
  { number: '301', type: 'double', floor: 3 },
  { number: '302', type: 'quadruple', floor: 3 },
  { number: '303', type: 'family', floor: 3 },
  { number: '305', type: 'triple', floor: 3 },
  { number: '306', type: 'quadruple', floor: 3 },
];

// Map room names from booking.com to our types
function normalizeRoomType(roomName) {
  const name = (roomName || '').toLowerCase();
  if (name.includes('double')) return 'double';
  if (name.includes('quadruple')) return 'quadruple';
  if (name.includes('family')) return 'family';
  if (name.includes('triple')) return 'triple';
  return null;
}

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
 * Check if a room is available during the given date range.
 */
function isRoomAvailable(roomAssignments, roomIndex, startDate, endDate) {
  for (const assignment of roomAssignments[roomIndex]) {
    // Overlap check: new reservation overlaps if it starts before existing ends and ends after existing starts
    if (startDate < assignment.endParsed && endDate > assignment.startParsed) {
      return false;
    }
  }
  return true;
}

/**
 * Calculate how far apart a set of rooms are by room number.
 */
function roomSpread(indices) {
  if (indices.length <= 1) return 0;
  const numbers = indices.map(i => parseInt(ROOMS[i].number));
  return Math.max(...numbers) - Math.min(...numbers);
}

/**
 * Find the combination of available rooms that satisfies all needed types
 * with minimum spread (room numbers as close together as possible).
 * @param {Array}         roomAssignments - current room assignment arrays
 * @param {Array<string>} neededTypes     - flat list of types, e.g. ['double','quadruple']
 * @param {Object}        reservation     - must have startParsed / endParsed
 * @param {number|null}   floor           - restrict to this floor, or null for any
 * @returns {Array<number>|null} room indices, or null if impossible
 */
function findClosestRooms(roomAssignments, neededTypes, reservation, floor, excludeIndices = []) {
  const candidates = [];
  for (const type of neededTypes) {
    const typeCandidates = [];
    for (let i = 0; i < ROOMS.length; i++) {
      if (excludeIndices.includes(i)) continue;
      if (ROOMS[i].type !== type) continue;
      if (floor !== null && ROOMS[i].floor !== floor) continue;
      if (!isRoomAvailable(roomAssignments, i, reservation.startParsed, reservation.endParsed)) continue;
      typeCandidates.push(i);
    }
    if (typeCandidates.length === 0) return null;
    candidates.push(typeCandidates);
  }

  let bestCombo = null;
  let bestSpread = Infinity;

  function backtrack(depth, chosen) {
    if (depth === candidates.length) {
      const spread = roomSpread(chosen);
      if (spread < bestSpread) {
        bestSpread = spread;
        bestCombo = [...chosen];
      }
      return;
    }
    for (const idx of candidates[depth]) {
      if (chosen.includes(idx)) continue;
      chosen.push(idx);
      backtrack(depth + 1, chosen);
      chosen.pop();
    }
  }

  backtrack(0, []);
  return bestCombo;
}

/**
 * Assign rooms to a reservation based on its room requirements.
 * Priority order:
 * 1. If this guest already occupies room(s), keep them in the same room(s)
 * 2. If any room is triple/family → prefer floor 3, rooms close together
 * 3. If multi-room with no triple/family → try both floors, pick closest spread
 * 4. If single room → prefer floor 2 first
 * Falls back to cross-floor assignment if no single floor works.
 */
function assignRoomsToReservation(roomAssignments, reservation) {
  const roomReqs = reservation.rooms || [];
  if (roomReqs.length === 0) return [];

  // Flatten requirements into individual type entries
  const neededTypes = [];
  for (const req of roomReqs) {
    const type = normalizeRoomType(req.name);
    if (!type) continue;
    const qty = req.quantity || 1;
    for (let q = 0; q < qty; q++) neededTypes.push(type);
  }
  if (neededTypes.length === 0) return [];

  // ── Priority 1: Keep guest in rooms they already occupy ──────────
  // Find all rooms where this guest already has an assignment
  // (from existing calendar restore or earlier reservations in this build).
  // This ensures a guest doesn't move between rooms across date ranges
  // and that a reservation stays in the same physical room.
  const guestExistingRooms = [];
  for (let i = 0; i < ROOMS.length; i++) {
    if (roomAssignments[i].some(r => r.name === reservation.name)) {
      guestExistingRooms.push(i);
    }
  }

  if (guestExistingRooms.length > 0) {
    const pinned = [];
    const unfulfilled = [...neededTypes];

    // Match existing rooms to needed types
    for (const ri of guestExistingRooms) {
      const roomType = ROOMS[ri].type;
      const typeIdx = unfulfilled.indexOf(roomType);
      if (typeIdx !== -1 && isRoomAvailable(roomAssignments, ri, reservation.startParsed, reservation.endParsed)) {
        pinned.push(ri);
        unfulfilled.splice(typeIdx, 1);
      }
    }

    if (unfulfilled.length === 0 && pinned.length > 0) {
      // All needs satisfied by existing rooms
      for (const idx of pinned) {
        roomAssignments[idx].push(reservation);
      }
      return pinned;
    }

    // Some needs pinned, find closest rooms for the rest
    if (pinned.length > 0 && unfulfilled.length > 0) {
      const extra = findClosestRooms(roomAssignments, unfulfilled, reservation, null, pinned);
      if (extra) {
        const all = [...pinned, ...extra];
        for (const idx of all) {
          roomAssignments[idx].push(reservation);
        }
        return all;
      }
    }
    // Fall through to normal algorithm if pinning didn't work
  }

  // ── Priority 2+: Floor preference based on room types ────────────
  const isSingle = neededTypes.length === 1;
  const needsThirdFloor = neededTypes.some(t => t === 'triple' || t === 'family');

  let bestAssignment = null;

  if (needsThirdFloor) {
    // Prefer floor 3, then floor 2, then cross-floor
    bestAssignment = findClosestRooms(roomAssignments, neededTypes, reservation, 3)
                  || findClosestRooms(roomAssignments, neededTypes, reservation, 2)
                  || findClosestRooms(roomAssignments, neededTypes, reservation, null);
  } else if (isSingle) {
    // Single room: prefer floor 2, then floor 3
    bestAssignment = findClosestRooms(roomAssignments, neededTypes, reservation, 2)
                  || findClosestRooms(roomAssignments, neededTypes, reservation, 3);
  } else {
    // Multi-room, no triple/family: try both floors, pick tightest spread
    const floor2 = findClosestRooms(roomAssignments, neededTypes, reservation, 2);
    const floor3 = findClosestRooms(roomAssignments, neededTypes, reservation, 3);
    if (floor2 && floor3) {
      bestAssignment = roomSpread(floor2) <= roomSpread(floor3) ? floor2 : floor3;
    } else {
      bestAssignment = floor2 || floor3
                    || findClosestRooms(roomAssignments, neededTypes, reservation, null);
    }
  }

  if (!bestAssignment) return [];

  const assigned = [];
  for (const idx of bestAssignment) {
    roomAssignments[idx].push(reservation);
    assigned.push(idx);
  }
  return assigned;
}

/**
 * Build a 2-D calendar grid for the given month.
 * Accepts reservation objects with { name, startDate, endDate, rooms }.
 *
 * When existingMonthAssignments is provided (from parseExistingAssignments),
 * reservations that already appear in the existing calendar keep their room.
 * Cancelled reservations (in existing but not in new data) are removed.
 * New reservations are assigned to available rooms automatically.
 */
function buildCalendar(reservations, year, month, existingMonthAssignments = {}) {
  const daysInMonth = new Date(year, month, 0).getDate();
  const monthStart  = new Date(year, month - 1, 1);
  const monthEnd    = new Date(year, month - 1, daysInMonth);

  const parsed = reservations
    .map(r => ({ ...r, startParsed: parseDate(r.startDate), endParsed: parseDate(r.endDate) }))
    .filter(r => r.startParsed && r.endParsed)
    .filter(r => r.startParsed <= monthEnd && r.endParsed > monthStart);

  parsed.sort((a, b) => a.startParsed - b.startParsed);

  // Initialize room assignments
  const roomAssignments = Array.from({ length: ROOMS.length }, () => []);

  // Phase 1: Restore existing room assignments for reservations that still exist
  const placedIndices = new Set();
  const hasExisting = Object.keys(existingMonthAssignments).length > 0;

  if (hasExisting) {
    for (let i = 0; i < ROOMS.length; i++) {
      const roomNumber = ROOMS[i].number;
      const existingBlocks = existingMonthAssignments[roomNumber] || [];
      for (const block of existingBlocks) {
        // Find a matching reservation by name with overlapping dates
        const blockStart = new Date(year, month - 1, block.startDay);
        const blockEnd   = new Date(year, month - 1, block.endDay + 1); // exclusive

        const matchIdx = parsed.findIndex((r, idx) => {
          if (placedIndices.has(idx)) return false;
          if (r.name !== block.name) return false;
          // Check date overlap between existing block and new reservation
          return r.startParsed < blockEnd && r.endParsed > blockStart;
        });

        if (matchIdx === -1) {
          if (block.isBooking) continue; // cancelled booking → remove
          // Manual entry → preserve in its current room
          const fakeRes = {
            name: block.name,
            source: 'manual',
            startParsed: new Date(year, month - 1, block.startDay),
            endParsed:   new Date(year, month - 1, block.endDay + 1),
            rooms: []
          };
          if (isRoomAvailable(roomAssignments, i, fakeRes.startParsed, fakeRes.endParsed)) {
            roomAssignments[i].push(fakeRes);
          }
          continue;
        }

        // Verify no conflict with already-placed reservations in this room
        const res = parsed[matchIdx];
        if (!isRoomAvailable(roomAssignments, i, res.startParsed, res.endParsed)) continue;

        roomAssignments[i].push(res);
        placedIndices.add(matchIdx);
      }
    }
  }

  // Phase 2: Assign remaining (new) reservations to available rooms
  const unplaced = parsed.filter((_, idx) => !placedIndices.has(idx));
  for (const res of unplaced) {
    assignRoomsToReservation(roomAssignments, res);
  }

  // Build header row with room numbers
  const header = ['Room'];
  for (let d = 1; d <= daysInMonth; d++) header.push(String(d));

  const rows = [header];
  for (let i = 0; i < ROOMS.length; i++) {
    const room = ROOMS[i];
    const row = [`${room.number} (${room.type})`];
    for (let d = 1; d <= daysInMonth; d++) {
      const date = new Date(year, month - 1, d);
      let cell = '';
      for (const res of roomAssignments[i]) {
        if (date >= res.startParsed && date < res.endParsed) {
          cell = res.name ? res.name + (res.source !== 'manual' ? BOOKING_TAG : '') : '';
          break;
        }
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
    // 1. Sign in to Google for Sheets access
    showStatus('Signing in to Google…', 'info');
    const token = await getAuthToken(true);

    // 2. Fetch reservations from booking.com API
    showStatus('Fetching reservations from booking.com…', 'info');
    let reservations = [];
    try {
      reservations = await fetchBookingComReservations();
      showStatus(`Found ${reservations.length} reservation(s) from booking.com.`, 'info');
    } catch (e) {
      showStatus(`Error: could not load reservations from booking.com: ${e.message}`, 'error');
      setRunning(false);
      return;
    }

    if (reservations.length === 0) {
      showStatus('No reservations found from booking.com.', 'warning');
      setRunning(false);
      return;
    }

    // 3. Read existing calendar to preserve room arrangements
    showStatus('Reading existing calendar…', 'info');
    const existingRows = await readExistingCalendar(targetSheetId, token);
    const existingAssignments = parseExistingAssignments(existingRows);

    // 4. Build calendars, merging with existing room assignments
    const monthRange = getMonthRange();
    showStatus(`Building calendars for ${monthRange.length} months…`, 'info');

    const allCalendarRows = [];
    for (const { year, month } of monthRange) {
      const monthName = new Date(year, month - 1, 1).toLocaleString('default', { month: 'long' });
      const monthKey = `${year}-${month}`;
      // Month separator row
      allCalendarRows.push([`${monthName} ${year}`]);
      const calendar = buildCalendar(reservations, year, month, existingAssignments[monthKey] || {});
      allCalendarRows.push(...calendar);
      allCalendarRows.push([]);  // blank row between months
    }

    showStatus('Writing calendar…', 'info');
    await clearAndWriteTab(targetSheetId, CALENDAR_TAB, allCalendarRows, token);

    showStatus(
      `✅ Done! ${reservations.length} reservations from booking.com written to Calendar.`,
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
