// ── Background service worker ─────────────────────────────────────────────────

const SHEETS_API = 'https://sheets.googleapis.com/v4/spreadsheets';
const CALENDAR_TAB = 'Calendar';
const BOOKING_TAG = ' [B]';

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

// ── Status broadcast ──────────────────────────────────────────────────────────

function broadcastStatus(message, type = 'info', final = false) {
  chrome.runtime.sendMessage({ action: 'statusUpdate', message, type, final }).catch(() => {});
}

// ── Booking.com API helpers ───────────────────────────────────────────────────

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

function formatISODate(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

async function fetchBookingComReservations() {
  const today = new Date();
  const allReservations = [];
  const seen = new Set();

  // Query 180 days in 10-day chunks (10 rooms × 10 days ≤ 100 items per query)
  for (let offset = 0; offset < 180; offset += 10) {
    const from = new Date(today);
    from.setDate(today.getDate() + offset);
    const to = new Date(today);
    to.setDate(today.getDate() + offset + 10);

    const response = await sendToBookingTab({
      action: 'fetchBookingReservations',
      dateFrom: formatISODate(from),
      dateTo: formatISODate(to)
    });

    for (const res of (response.reservations || [])) {
      const key = res.bookingNumber || `${res.name}_${res.startDate}_${res.endDate}`;
      if (!seen.has(key)) {
        seen.add(key);
        allReservations.push(res);
      }
    }

    broadcastStatus(`Fetching reservations… (day ${offset + 10}/300)`, 'info');
    const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));
    await sleep(2000);
  }

  return allReservations;
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

async function clearAndWriteTab(spreadsheetId, tabName, rows, token) {
  await ensureTab(spreadsheetId, tabName, token);
  const clearUrl = `${SHEETS_API}/${encodeURIComponent(spreadsheetId)}/values/${encodeURIComponent(tabName)}:clear`;
  await fetch(clearUrl, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}` }
  });
  await sheetsUpdate(spreadsheetId, `${tabName}!A1`, rows, token);
}

// ── Read existing calendar ────────────────────────────────────────────────────

async function readExistingCalendar(spreadsheetId, token) {
  await ensureTab(spreadsheetId, CALENDAR_TAB, token);
  try {
    const data = await sheetsGet(spreadsheetId, CALENDAR_TAB, token);
    return data.values || [];
  } catch {
    return [];
  }
}

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
      i++;
      if (i >= sheetRows.length) break;
      const header = sheetRows[i] || [];
      const numDays = header.length - 1;
      const monthAssignments = {};
      i++;
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

// ── Calendar builder ──────────────────────────────────────────────────────────

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

function isRoomAvailable(roomAssignments, roomIndex, startDate, endDate) {
  for (const assignment of roomAssignments[roomIndex]) {
    if (startDate < assignment.endParsed && endDate > assignment.startParsed) {
      return false;
    }
  }
  return true;
}

function roomSpread(indices) {
  if (indices.length <= 1) return 0;
  const numbers = indices.map(i => parseInt(ROOMS[i].number));
  return Math.max(...numbers) - Math.min(...numbers);
}

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

function assignRoomsToReservation(roomAssignments, reservation) {
  const roomReqs = reservation.rooms || [];
  if (roomReqs.length === 0) return [];
  const neededTypes = [];
  for (const req of roomReqs) {
    const type = normalizeRoomType(req.name);
    if (!type) continue;
    const qty = req.quantity || 1;
    for (let q = 0; q < qty; q++) neededTypes.push(type);
  }
  if (neededTypes.length === 0) return [];

  const guestExistingRooms = [];
  for (let i = 0; i < ROOMS.length; i++) {
    if (roomAssignments[i].some(r => r.name === reservation.name)) {
      guestExistingRooms.push(i);
    }
  }
  if (guestExistingRooms.length > 0) {
    const pinned = [];
    const unfulfilled = [...neededTypes];
    for (const ri of guestExistingRooms) {
      const roomType = ROOMS[ri].type;
      const typeIdx = unfulfilled.indexOf(roomType);
      if (typeIdx !== -1 && isRoomAvailable(roomAssignments, ri, reservation.startParsed, reservation.endParsed)) {
        pinned.push(ri);
        unfulfilled.splice(typeIdx, 1);
      }
    }
    if (unfulfilled.length === 0 && pinned.length > 0) {
      for (const idx of pinned) {
        roomAssignments[idx].push(reservation);
      }
      return pinned;
    }
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
  }

  const isSingle = neededTypes.length === 1;
  const needsThirdFloor = neededTypes.some(t => t === 'triple' || t === 'family');
  let bestAssignment = null;
  if (needsThirdFloor) {
    bestAssignment = findClosestRooms(roomAssignments, neededTypes, reservation, 3)
                  || findClosestRooms(roomAssignments, neededTypes, reservation, 2)
                  || findClosestRooms(roomAssignments, neededTypes, reservation, null);
  } else if (isSingle) {
    bestAssignment = findClosestRooms(roomAssignments, neededTypes, reservation, 2)
                  || findClosestRooms(roomAssignments, neededTypes, reservation, 3);
  } else {
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

function tryRearrangeAndAssign(roomAssignments, reservation, neededTypes) {
  const typeToIndices = {};
  for (let i = 0; i < ROOMS.length; i++) {
    const t = ROOMS[i].type;
    if (!typeToIndices[t]) typeToIndices[t] = [];
    typeToIndices[t].push(i);
  }
  if (neededTypes.length === 1) {
    const type = neededTypes[0];
    const sameTypeRooms = typeToIndices[type] || [];
    if (sameTypeRooms.length < 2) return [];
    for (const targetIdx of sameTypeRooms) {
      const conflicts = roomAssignments[targetIdx].filter(
        a => reservation.startParsed < a.endParsed && reservation.endParsed > a.startParsed
      );
      if (conflicts.length === 0) continue;
      let canClearAll = true;
      const moves = [];
      for (const conflict of conflicts) {
        let moved = false;
        for (const altIdx of sameTypeRooms) {
          if (altIdx === targetIdx) continue;
          const altFree = !roomAssignments[altIdx].some(
            a => a !== conflict && conflict.startParsed < a.endParsed && conflict.endParsed > a.startParsed
          );
          const blockedByMove = moves.some(
            m => m.toIdx === altIdx && conflict.startParsed < m.res.endParsed && conflict.endParsed > m.res.startParsed
          );
          if (altFree && !blockedByMove) {
            moves.push({ res: conflict, fromIdx: targetIdx, toIdx: altIdx });
            moved = true;
            break;
          }
        }
        if (!moved) { canClearAll = false; break; }
      }
      if (canClearAll) {
        for (const { res, fromIdx, toIdx } of moves) {
          roomAssignments[fromIdx] = roomAssignments[fromIdx].filter(r => r !== res);
          roomAssignments[toIdx].push(res);
        }
        roomAssignments[targetIdx].push(reservation);
        return [targetIdx];
      }
    }
    return [];
  }
  for (const type of [...new Set(neededTypes)]) {
    const sameTypeRooms = typeToIndices[type] || [];
    for (const targetIdx of sameTypeRooms) {
      const conflicts = roomAssignments[targetIdx].filter(
        a => reservation.startParsed < a.endParsed && reservation.endParsed > a.startParsed
      );
      if (conflicts.length === 0) continue;
      let canClearAll = true;
      const moves = [];
      for (const conflict of conflicts) {
        let moved = false;
        for (const altIdx of sameTypeRooms) {
          if (altIdx === targetIdx) continue;
          const altFree = !roomAssignments[altIdx].some(
            a => a !== conflict && conflict.startParsed < a.endParsed && conflict.endParsed > a.startParsed
          );
          const blockedByMove = moves.some(
            m => m.toIdx === altIdx && conflict.startParsed < m.res.endParsed && conflict.endParsed > m.res.startParsed
          );
          if (altFree && !blockedByMove) {
            moves.push({ res: conflict, fromIdx: targetIdx, toIdx: altIdx });
            moved = true;
            break;
          }
        }
        if (!moved) { canClearAll = false; break; }
      }
      if (canClearAll) {
        for (const { res, fromIdx, toIdx } of moves) {
          roomAssignments[fromIdx] = roomAssignments[fromIdx].filter(r => r !== res);
          roomAssignments[toIdx].push(res);
        }
        const result = assignRoomsToReservation(roomAssignments, reservation);
        if (result.length > 0) return result;
        for (const { res, fromIdx, toIdx } of moves) {
          roomAssignments[toIdx] = roomAssignments[toIdx].filter(r => r !== res);
          roomAssignments[fromIdx].push(res);
        }
      }
    }
  }
  return [];
}

function buildCalendar(reservations, year, month, existingMonthAssignments = {}, autoRearrange = false) {
  const daysInMonth = new Date(year, month, 0).getDate();
  const monthStart  = new Date(year, month - 1, 1);
  const monthEnd    = new Date(year, month - 1, daysInMonth);

  const parsed = reservations
    .map(r => ({ ...r, startParsed: parseDate(r.startDate), endParsed: parseDate(r.endDate) }))
    .filter(r => r.startParsed && r.endParsed)
    .filter(r => r.startParsed <= monthEnd && r.endParsed > monthStart);

  parsed.sort((a, b) => a.startParsed - b.startParsed);

  const roomAssignments = Array.from({ length: ROOMS.length }, () => []);
  const unplaceable = [];
  const placedIndices = new Set();
  const hasExisting = Object.keys(existingMonthAssignments).length > 0;

  if (hasExisting) {
    for (let i = 0; i < ROOMS.length; i++) {
      const roomNumber = ROOMS[i].number;
      const existingBlocks = existingMonthAssignments[roomNumber] || [];
      for (const block of existingBlocks) {
        const blockStart = new Date(year, month - 1, block.startDay);
        const blockEnd   = new Date(year, month - 1, block.endDay + 1);
        const matchIdx = parsed.findIndex((r, idx) => {
          if (placedIndices.has(idx)) return false;
          if (r.name !== block.name) return false;
          return r.startParsed < blockEnd && r.endParsed > blockStart;
        });
        if (matchIdx === -1) {
          if (block.isBooking) continue;
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
        const res = parsed[matchIdx];
        if (!isRoomAvailable(roomAssignments, i, res.startParsed, res.endParsed)) continue;
        roomAssignments[i].push(res);
        placedIndices.add(matchIdx);
      }
    }
  }

  const unplaced = parsed.filter((_, idx) => !placedIndices.has(idx));
  for (const res of unplaced) {
    const result = assignRoomsToReservation(roomAssignments, res);
    if (result.length === 0) {
      if (autoRearrange) {
        const neededTypes = [];
        for (const req of (res.rooms || [])) {
          const type = normalizeRoomType(req.name);
          if (!type) continue;
          const qty = req.quantity || 1;
          for (let q = 0; q < qty; q++) neededTypes.push(type);
        }
        const swapResult = neededTypes.length > 0
          ? tryRearrangeAndAssign(roomAssignments, res, neededTypes)
          : [];
        if (swapResult.length === 0) {
          unplaceable.push(res);
        }
      } else {
        unplaceable.push(res);
      }
    }
  }

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

  const splitCells = [];
  const trulyUnplaceable = [];
  for (const res of unplaceable) {
    const cellName = res.name
      ? res.name + (res.source !== 'manual' ? BOOKING_TAG : '')
      : '';
    if (!cellName) { trulyUnplaceable.push(res); continue; }
    const neededTypes = [];
    for (const req of (res.rooms || [])) {
      const type = normalizeRoomType(req.name);
      if (type) neededTypes.push(type);
    }
    const preferredType = neededTypes[0] || null;
    let anyPlaced = false;
    for (let d = 1; d <= daysInMonth; d++) {
      const date = new Date(year, month - 1, d);
      if (date < res.startParsed || date >= res.endParsed) continue;
      let placedRowIdx = -1;
      if (preferredType) {
        for (let i = 0; i < ROOMS.length; i++) {
          if (ROOMS[i].type !== preferredType) continue;
          if (rows[i + 1][d] === '') {
            rows[i + 1][d] = cellName;
            placedRowIdx = i;
            break;
          }
        }
      }
      if (placedRowIdx === -1) {
        for (let i = 0; i < ROOMS.length; i++) {
          if (rows[i + 1][d] === '') {
            rows[i + 1][d] = cellName;
            placedRowIdx = i;
            break;
          }
        }
      }
      if (placedRowIdx >= 0) {
        anyPlaced = true;
        splitCells.push({ gridRow: placedRowIdx + 1, gridCol: d });
      }
    }
    if (!anyPlaced) {
      trulyUnplaceable.push(res);
    }
  }

  return { rows, unplaceable: trulyUnplaceable, splitCells };
}

// ── Highlight helpers ─────────────────────────────────────────────────────────

async function getTabSheetId(spreadsheetId, tabName, token) {
  const metaRes = await fetch(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!metaRes.ok) return null;
  const meta = await metaRes.json();
  const sheet = meta.sheets?.find(s => s.properties?.title === tabName);
  return sheet ? sheet.properties.sheetId : null;
}

async function highlightCells(spreadsheetId, cells, token) {
  const sheetId = await getTabSheetId(spreadsheetId, CALENDAR_TAB, token);
  if (sheetId == null || cells.length === 0) return;
  const requests = [
    {
      repeatCell: {
        range: { sheetId },
        cell: { userEnteredFormat: { backgroundColor: { red: 1, green: 1, blue: 1 } } },
        fields: 'userEnteredFormat.backgroundColor'
      }
    }
  ];
  for (const { row, col } of cells) {
    requests.push({
      repeatCell: {
        range: {
          sheetId,
          startRowIndex: row,
          endRowIndex: row + 1,
          startColumnIndex: col,
          endColumnIndex: col + 1
        },
        cell: {
          userEnteredFormat: {
            backgroundColor: { red: 1, green: 0.85, blue: 0.4 }
          }
        },
        fields: 'userEnteredFormat.backgroundColor'
      }
    });
  }
  await fetch(`${SHEETS_API}/${encodeURIComponent(spreadsheetId)}:batchUpdate`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ requests })
  });
}

// ── Month range ───────────────────────────────────────────────────────────────

function getMonthRange() {
  const months = [];
  const now = new Date();
  for (let i = 0; i < 11; i++) {
    const d = new Date(now.getFullYear(), now.getMonth() + i, 1);
    months.push({ year: d.getFullYear(), month: d.getMonth() + 1 });
  }
  return months;
}

// ── Core automation runner ────────────────────────────────────────────────────

let automationRunning = false;

async function runAutomationCore(targetSheetId, autoRearrange, interactive) {
  if (automationRunning) {
    broadcastStatus('Automation is already running.', 'warning', true);
    return;
  }
  automationRunning = true;
  try {
    broadcastStatus('Signing in to Google…', 'info');
    const token = await getAuthToken(interactive);

    broadcastStatus('Fetching reservations from booking.com…', 'info');
    let reservations;
    try {
      reservations = await fetchBookingComReservations();
      broadcastStatus(`Found ${reservations.length} reservation(s) from booking.com.`, 'info');
    } catch (e) {
      broadcastStatus(`Error: could not load reservations from booking.com: ${e.message}`, 'error', true);
      return;
    }

    if (reservations.length === 0) {
      broadcastStatus('No reservations found from booking.com.', 'warning', true);
      return;
    }

    broadcastStatus('Reading existing calendar…', 'info');
    const existingRows = await readExistingCalendar(targetSheetId, token);
    const existingAssignments = parseExistingAssignments(existingRows);

    const monthRange = getMonthRange();
    broadcastStatus(`Building calendars for ${monthRange.length} months…`, 'info');

    const allCalendarRows = [];
    const allUnplaceable = [];
    const allSplitCells = [];
    let rowOffset = 0;
    for (const { year, month } of monthRange) {
      const monthName = new Date(year, month - 1, 1).toLocaleString('default', { month: 'long' });
      const monthKey = `${year}-${month}`;
      allCalendarRows.push([`${monthName} ${year}`]);
      const gridStartRow = rowOffset + 1;
      const { rows: calendar, unplaceable, splitCells } = buildCalendar(
        reservations, year, month,
        existingAssignments[monthKey] || {},
        autoRearrange
      );
      allCalendarRows.push(...calendar);
      allCalendarRows.push([]);
      for (const res of unplaceable) {
        allUnplaceable.push({ ...res, monthName: `${monthName} ${year}` });
      }
      for (const { gridRow, gridCol } of splitCells) {
        allSplitCells.push({ row: gridStartRow + gridRow, col: gridCol });
      }
      rowOffset += 1 + calendar.length + 1;
    }

    broadcastStatus('Writing calendar…', 'info');
    await clearAndWriteTab(targetSheetId, CALENDAR_TAB, allCalendarRows, token);

    if (allSplitCells.length > 0) {
      broadcastStatus('Highlighting split reservations…', 'info');
      await highlightCells(targetSheetId, allSplitCells, token);
    }

    let finalMsg, finalType;
    if (allUnplaceable.length > 0) {
      const names = allUnplaceable.map(r => `• ${r.name} (${r.monthName})`).join('\n');
      finalMsg = `⚠️ Done with ${allUnplaceable.length} unplaceable reservation(s):\n${names}`;
      finalType = 'warning';
    } else if (allSplitCells.length > 0) {
      finalMsg = `✅ Done! ${reservations.length} reservations written. Some required split rooms (highlighted in orange).`;
      finalType = 'warning';
    } else {
      finalMsg = `✅ Done! ${reservations.length} reservations from booking.com written to Calendar.`;
      finalType = 'success';
    }

    chrome.storage.local.set({
      lastSyncTime: new Date().toISOString(),
      lastSyncResult: { message: finalMsg, type: finalType }
    });
    broadcastStatus(finalMsg, finalType, true);
  } catch (e) {
    const errMsg = `Error: ${e.message}`;
    chrome.storage.local.set({
      lastSyncTime: new Date().toISOString(),
      lastSyncResult: { message: errMsg, type: 'error' }
    });
    broadcastStatus(errMsg, 'error', true);
  } finally {
    automationRunning = false;
  }
}

// ── Schedule / Alarm management ───────────────────────────────────────────────

function setupScheduleAlarm() {
  chrome.alarms.clear('scheduledSync');
  chrome.storage.local.get(['scheduleEnabled', 'scheduleHours'], (data) => {
    if (!data.scheduleEnabled || !data.scheduleHours?.length) return;
    const now = new Date();
    const nextHour = new Date(now.getFullYear(), now.getMonth(), now.getDate(), now.getHours() + 1, 0, 0);
    chrome.alarms.create('scheduledSync', {
      when: nextHour.getTime(),
      periodInMinutes: 60
    });
  });
}

chrome.alarms.onAlarm.addListener(async (alarm) => {
  if (alarm.name !== 'scheduledSync') return;
  const data = await chrome.storage.local.get(['scheduleEnabled', 'scheduleHours', 'targetSheetId', 'autoRearrange']);
  if (!data.scheduleEnabled || !data.targetSheetId) return;
  const currentHour = new Date().getHours();
  if (!data.scheduleHours?.includes(currentHour)) return;
  await runAutomationCore(data.targetSheetId, !!data.autoRearrange, false);
});

// Restore alarms on service worker startup
setupScheduleAlarm();

// ── Message handling ──────────────────────────────────────────────────────────

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.action === 'runAutomation') {
    runAutomationCore(msg.targetSheetId, msg.autoRearrange, true);
    sendResponse({ started: true });
    return false;
  }
  if (msg.action === 'updateSchedule') {
    setupScheduleAlarm();
    sendResponse({ success: true });
    return false;
  }
  if (msg.action === 'downloadFile') {
    const { url, filename } = msg;
    chrome.downloads.download({
      url,
      filename: filename || 'reservations.xlsx',
      saveAs: false
    }, (downloadId) => {
      if (chrome.runtime.lastError) {
        sendResponse({ success: false, error: chrome.runtime.lastError.message });
      } else {
        sendResponse({ success: true, downloadId });
      }
    });
    return true;
  }
});
