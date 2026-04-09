// ── Constants ─────────────────────────────────────────────────────────────────

/** Hardcoded hotel ID for the booking.com property. */
const HOTEL_ID = '15204299';

/** Hardcoded hotel account ID for the booking.com property. */
const HOTEL_ACCOUNT_ID = '25086419';

/**
 * The path (relative to the booking.com admin host) for the reservation
 * search page where the xlsx download can be triggered.
 */
const RESERVATION_SEARCH_PATH =
  '/hotel/hoteladmin/extranet_ng/manage/search_reservations.html';

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
 * Extract the `ses` session token from the current page's query string.
 */
function getSessionId() {
  return new URLSearchParams(window.location.search).get('ses') || '';
}

/**
 * Generate array of date strings (YYYY-MM-DD) from today to 6 months later.
 */
function generateDateRange() {
  const dates = [];
  const today = new Date();
  const end = new Date(today);
  end.setMonth(end.getMonth() + 6);

  const current = new Date(today);
  while (current <= end) {
    dates.push(formatISODate(current));
    current.setDate(current.getDate() + 1);
  }
  return dates;
}

// ── Date helpers ──────────────────────────────────────────────────────────────

/**
 * Parse a date string from booking.com.
 * Handles formats like:
 *   "Sat 5 Apr 2025", "Sat, Apr 5, 2025", "5 Apr 2025", "Apr 5, 2025"
 * Returns a Date (midnight local time) or null.
 */
function parseBookingDate(dateStr) {
  if (!dateStr) return null;

  // Strip leading weekday token (e.g. "Sat," or "Sat")
  const stripped = dateStr.replace(/^\s*(Mon|Tue|Wed|Thu|Fri|Sat|Sun)[a-z]*\.?,?\s+/i, '').trim();

  // Try native parse on stripped string first
  const native = new Date(stripped);
  if (!isNaN(native)) {
    return new Date(native.getFullYear(), native.getMonth(), native.getDate());
  }

  // DD Mon YYYY  e.g.  5 Apr 2025
  let m = stripped.match(/^(\d{1,2})\s+([A-Za-z]{3,})\s+(\d{4})$/);
  if (m) {
    const d = new Date(`${m[2]} ${m[1]}, ${m[3]}`);
    if (!isNaN(d)) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  // Mon DD, YYYY  e.g.  Apr 5, 2025
  m = stripped.match(/^([A-Za-z]{3,})\s+(\d{1,2}),?\s+(\d{4})$/);
  if (m) {
    const d = new Date(`${m[1]} ${m[2]}, ${m[3]}`);
    if (!isNaN(d)) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  console.warn('[content.js] Could not parse date:', dateStr);
  return null;
}

// ── Calendar builder ─────────────────────────────────────────────────────────

const NUM_ROOMS = 10;

/**
 * Assign reservations to room rows (greedy, sorted by check-in).
 * Returns an array of NUM_ROOMS arrays, each containing the reservations
 * assigned to that room (with parsed Date objects attached).
 */
function assignRooms(reservations) {
  const rooms = Array.from({ length: NUM_ROOMS }, () => []);

  const sorted = [...reservations].sort((a, b) => a.startParsed - b.startParsed);

  for (const res of sorted) {
    let placed = false;
    for (let i = 0; i < NUM_ROOMS; i++) {
      const last = rooms[i][rooms[i].length - 1];
      // Check-out day is free for new arrivals (strict less-than)
      if (!last || last.endParsed <= res.startParsed) {
        rooms[i].push(res);
        placed = true;
        break;
      }
    }
    if (!placed) {
      console.warn('[content.js] No room available for reservation:', res.name, res.startDate);
    }
  }
  return rooms;
}

/**
 * Build a 2-D array representing the calendar for the given month.
 * Row 0:  ["Room", "1", "2", ..., daysInMonth]
 * Row 1–10: ["Room 1", guest | "", ...]
 *
 * @param {Array}  reservations  – combined from booking.com + phone sheet
 * @param {number} year
 * @param {number} month  – 1-based (January = 1)
 * @returns {Array<Array<string>>}
 */
function buildCalendar(reservations, year, month) {
  const daysInMonth = new Date(year, month, 0).getDate();
  const monthStart  = new Date(year, month - 1, 1);
  const monthEnd    = new Date(year, month - 1, daysInMonth);

  // Attach parsed dates and filter to reservations overlapping this month
  const parsed = reservations
    .map(r => ({
      ...r,
      startParsed: parseBookingDate(r.startDate),
      endParsed:   parseBookingDate(r.endDate)
    }))
    .filter(r => r.startParsed && r.endParsed)
    .filter(r => r.startParsed <= monthEnd && r.endParsed > monthStart);

  const rooms = assignRooms(parsed);

  // Header row
  const header = ['Room'];
  for (let d = 1; d <= daysInMonth; d++) header.push(String(d));

  const rows = [header];

  for (let i = 0; i < NUM_ROOMS; i++) {
    const row = [`Room ${i + 1}`];
    for (let d = 1; d <= daysInMonth; d++) {
      const date = new Date(year, month - 1, d);
      let cell = '';
      for (const res of rooms[i]) {
        // Guest occupies the room from startDate up to (not including) endDate
        if (date >= res.startParsed && date < res.endParsed) {
          cell = res.name || '';
          break;
        }
      }
      row.push(cell);
    }
    rows.push(row);
  }

  return rows;
}

// ── Fetch report as binary ────────────────────────────────────────────────────

/**
 * Fetch the latest report's binary data (XLS/XLSX) and return as base64.
 * Uses the booking.com API to list reports, picks the latest, then fetches it.
 */
async function fetchLatestReportBinary() {
  const reports = await fetchReportList();
  if (reports.length === 0) throw new Error('No reports available');

  const latestReport = reports.reduce((max, r) => (r.id > max.id ? r : max), reports[0]);
  console.log('[content.js] Fetching binary for report:', latestReport);

  const ses = getSessionId();
  const url = `https://admin.booking.com/fresa/extranet/reservations/download_request?lang=xu&hotel_account_id=${HOTEL_ACCOUNT_ID}&ses=${ses}&hotel_id=${HOTEL_ID}&report_id=${latestReport.id}`;

  const response = await fetch(url, { credentials: 'include' });
  if (!response.ok) throw new Error(`Download failed: ${response.status}`);

  const buf = await response.arrayBuffer();
  // Convert ArrayBuffer to base64 string for messaging
  const bytes = new Uint8Array(buf);
  let binary = '';
  for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
  const base64 = btoa(binary);

  return { base64, reportName: latestReport.name || `report_${latestReport.id}` };
}

// ── Fetch reservations from booking.com API ───────────────────────────────────

/**
 * Extract the CSRF token from the search_reservations.html form.
 */
function extractToken() {
  const tokenInput = document.querySelector('input[name="token"][type="hidden"]');
  if (!tokenInput) {
    throw new Error('Could not find token input on page');
  }
  return tokenInput.value;
}

/**
 * Fetch reservations from booking.com retrieve_list_v2 API.
 * @param {string} dateFrom - Start date in YYYY-MM-DD format
 * @param {string} dateTo - End date in YYYY-MM-DD format
 * @returns {Promise<Array>} - Array of reservation objects
 */
async function fetchBookingReservations(dateFrom, dateTo) {
  const token = extractToken();
  const ses = getSessionId();

  const params = new URLSearchParams({
    hotel_account_id: HOTEL_ACCOUNT_ID,
    hotel_id: HOTEL_ID,
    lang: 'xu',
    ses: ses,
    perpage: '100',
    page: '1',
    date_type: 'arrival',
    date_from: dateFrom,
    date_to: dateTo,
    token: token,
    user_triggered_search: '1'
  });

  const url = `https://admin.booking.com/fresa/extranet/reservations/retrieve_list_v2?${params.toString()}`;

  const response = await fetch(url, {
    method: 'POST',
    credentials: 'include'
  });

  if (!response.ok) {
    throw new Error(`API request failed: ${response.status}`);
  }

  const data = await response.json();
  console.log('[content.js] API response data:', data);

  // Parse the response and extract reservation data
  // The exact structure depends on booking.com's API response
  const reservations = [];
  if (data && data.data && data.data.reservations) {
    for (const res of data.data.reservations) {
      if (res.reservationStatus !== 'ok')
        continue;
      reservations.push({
        name: res.guestName || 'unknown',
        startDate: res.checkin || res.arrivalDate || '',
        endDate: res.checkout || res.departureDate || '',
        room: res.rooms || [],
        source: 'booking.com',
        bookingNumber: res.id || -1
      });
    }
  }

  return reservations;
}

// ── Message listener ──────────────────────────────────────────────────────────

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  // Fetch reservations from booking.com API
  if (msg.action === 'fetchBookingReservations') {
    const { dateFrom, dateTo } = msg;
    fetchBookingReservations(dateFrom, dateTo)
      .then(reservations => sendResponse({ success: true, reservations }))
      .catch(err => sendResponse({ success: false, error: err.message }));
    return true; // keep channel open for async
  }

  // Navigate to the reservation-statements download page.
  if (msg.action === 'extractReservations') {
    const url = navigateToDownloadPage();
    sendResponse({ navigating: true, url });
  }

  // Explicit request to click the download button.
  if (msg.action === 'clickDownloadButton') {
    const clicked = clickDownloadButton();
    sendResponse({ clicked });
  }

  if (msg.action === 'buildCalendar') {
    const { allReservations, year, month } = msg;
    sendResponse({ calendar: buildCalendar(allReservations, year, month) });
  }

  // Fetch latest report as base64 binary for XLS parsing in popup.
  if (msg.action === 'fetchLatestReportBinary') {
    fetchLatestReportBinary()
      .then(result => sendResponse({ success: true, ...result }))
      .catch(err => sendResponse({ success: false, error: err.message }));
  }

  // Legacy action: redirect to the download page.
  if (msg.action === 'run') {
    const url = navigateToDownloadPage();
    sendResponse({ navigating: true, url });
  }

  // Fetch room inventory from GraphQL API.
  if (msg.action === 'fetchRoomInventory') {
    fetchRoomInventory()
      .then(data => sendResponse({ success: true, data }))
      .catch(err => sendResponse({ success: false, error: err.message }));
  }

  return true; // keep channel open for async use
});
