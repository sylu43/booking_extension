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
 * Build the full URL for the reservation search page with download params:
 *   - ses        : current session token
 *   - hotel_id   : hardcoded property ID
 *   - date_from  : today (YYYY-MM-DD)
 *   - date_to    : six months from today (YYYY-MM-DD)
 *   - date_type  : 'arrival'
 *   - upcoming_reservations : 1
 */
function buildDownloadUrl() {
  const ses = getSessionId();
  const today = new Date();
  const sixMonthsLater = new Date(today);
  sixMonthsLater.setMonth(sixMonthsLater.getMonth() + 6);

  const params = new URLSearchParams({
    upcoming_reservations: '1',
    source: 'nav',
    hotel_id: HOTEL_ID,
    lang: 'xu',
    ses,
    date_from: formatISODate(today),
    date_to: formatISODate(sixMonthsLater),
    date_type: 'arrival'
  });

  return `https://admin.booking.com${RESERVATION_SEARCH_PATH}?${params}`;
}

/**
 * Navigate the current tab to the reservation-statements download page.
 * Returns the constructed URL so callers can report it.
 */
function navigateToDownloadPage() {
  const url = buildDownloadUrl();
  window.location.href = url;
  return url;
}

/**
 * Fetch the list of available reports from booking.com API.
 * Returns array of report objects with id, name, status, etc.
 */
async function fetchReportList() {
  const ses = getSessionId();
  const url = `https://admin.booking.com/fresa/extranet/reservations/list_request?hotel_id=${HOTEL_ID}&ses=${ses}&hotel_account_id=${HOTEL_ACCOUNT_ID}&lang=xu`;

  console.log('[content.js] Fetching report list from:', url);

  const response = await fetch(url, {
    credentials: 'include'
  });

  if (!response.ok) {
    throw new Error(`Failed to fetch report list: ${response.status}`);
  }

  const data = await response.json();

  if (!data.success || !data.data?.ok) {
    throw new Error('API returned error');
  }

  return data.data.reports || [];
}

/**
 * Download a report by its ID.
 * Navigates to the download URL which triggers the browser download.
 */
function downloadReport(reportId) {
  const ses = getSessionId();
  const url = `https://admin.booking.com/fresa/extranet/reservations/download_request?lang=xu&hotel_account_id=${HOTEL_ACCOUNT_ID}&ses=${ses}&hotel_id=${HOTEL_ID}&report_id=${reportId}`;

  console.log('[content.js] Downloading report:', url);

  // Create a temporary link and click it to trigger download
  const link = document.createElement('a');
  link.href = url;
  link.download = '';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

/**
 * Fetch reports and download the one with the largest ID.
 */
async function downloadLatestReport() {
  const reports = await fetchReportList();

  if (reports.length === 0) {
    throw new Error('No reports available');
  }

  // Find report with largest ID
  const latestReport = reports.reduce((max, report) =>
    report.id > max.id ? report : max
  , reports[0]);

  console.log('[content.js] Latest report:', latestReport);

  downloadReport(latestReport.id);

  return latestReport;
}

// ── Auto-download on the reservation search page ──────────────────────────────

/**
 * When the content script is injected into the reservation search page,
 * fetch the report list via API and download the latest report.
 */
if (window.location.pathname.includes('search_reservations')) {
  const startDownload = async () => {
    // Wait for page to settle
    await new Promise(r => setTimeout(r, 2000));

    try {
      const report = await downloadLatestReport();
      console.log('[content.js] Download initiated for report:', report.name);
    } catch (e) {
      console.warn('[content.js] Failed to download report:', e.message);
    }
  };

  if (document.readyState === 'loading') {
    console.log('[content.js] Waiting for DOMContentLoaded to start download');
    document.addEventListener('DOMContentLoaded', startDownload);
  } else {
    console.log('[content.js] Document already loaded, starting download');
    startDownload();
  }
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

// ── Message listener ──────────────────────────────────────────────────────────

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  // Navigate to the reservation-statements download page.
  // The content script will auto-click the download button once the page loads.
  if (msg.action === 'extractReservations') {
    const url = navigateToDownloadPage();
    sendResponse({ navigating: true, url });
  }

  // Explicit request to click the download button (sent after the download
  // page has finished loading and the auto-click did not fire).
  if (msg.action === 'clickDownloadButton') {
    const clicked = clickDownloadButton();
    sendResponse({ clicked });
  }

  if (msg.action === 'buildCalendar') {
    const { allReservations, year, month } = msg;
    sendResponse({ calendar: buildCalendar(allReservations, year, month) });
  }

  // Legacy action: redirect to the download page (same as 'extractReservations').
  if (msg.action === 'run') {
    const url = navigateToDownloadPage();
    sendResponse({ navigating: true, url });
  }

  return true; // keep channel open for async use
});
