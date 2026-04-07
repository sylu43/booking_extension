// ── Constants ─────────────────────────────────────────────────────────────────

/** Hardcoded hotel ID for the booking.com property. */
const HOTEL_ID = '15204299';

/**
 * The path (relative to the booking.com admin host) for the reservation
 * search page where the xlsx download can be triggered.
 */
const RESERVATION_SEARCH_PATH =
  '/hotel/hoteladmin/extranet_ng/manage/search_reservations.html';

/**
 * CSS selector that matches the "Download reservations statement" button.
 */
const DOWNLOAD_BUTTON_SELECTOR = 'button.bui-button--tertiary';

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
 * Find and click the "Download reservations statement" button.
 * Returns true when the button was found and clicked, false otherwise.
 */
function clickDownloadButton() {
  // Find button containing "Download reservations statement" text
  const buttons = document.querySelectorAll(DOWNLOAD_BUTTON_SELECTOR);
  for (const btn of buttons) {
    if (btn.textContent.includes('Download reservations statement')) {
      console.log('[content.js] Clicking download button:', btn);
      btn.click();
      return true;
    }
  }
  console.warn('[content.js] Download button not found');
  return false;
}

/**
 * Wait for the download list item matching our date range to appear,
 * then click it to start the download.
 */
function waitForDownloadItem() {
  const today = new Date();
  const sixMonthsLater = new Date(today);
  sixMonthsLater.setMonth(sixMonthsLater.getMonth() + 6);
  const expectedDateFrom = formatISODate(today);
  const expectedDateTo = formatISODate(sixMonthsLater);
  const expectedText = `Check-in ${expectedDateFrom} to ${expectedDateTo}`;

  return new Promise((resolve, reject) => {
    let attempts = 0;
    const maxAttempts = 30; // 30 seconds max

    const checkForItem = () => {
      const items = document.querySelectorAll('li.res-download-item');
      for (const item of items) {
        const text = item.textContent || '';
        if (text.includes(expectedDateFrom) && text.includes(expectedDateTo) && text.includes('Ready to download')) {
          console.log('[content.js] Found download item, clicking:', item);
          item.click();
          resolve(true);
          return;
        }
      }

      attempts++;
      if (attempts >= maxAttempts) {
        reject(new Error('Timeout waiting for download item'));
        return;
      }
      setTimeout(checkForItem, 1000);
    };

    checkForItem();
  });
}

// ── Auto-click on the reservation search page ──────────────────────────────

/**
 * When the content script is injected into the reservation search page,
 * automatically click the download button, then wait for the download
 * list item to appear and click it.
 */
if (window.location.pathname.includes('search_reservations')) {
  const startDownload = async () => {
    // First click "Download reservations statement" button
    await new Promise(r => setTimeout(r, 1000)); // Wait for page to settle
    const clicked = clickDownloadButton();
    if (clicked) {
      // Wait for download list to appear and click the matching item
      try {
        await waitForDownloadItem();
        console.log('[content.js] Download initiated successfully');
      } catch (e) {
        console.warn('[content.js] Failed to find download item:', e.message);
      }
    }
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', startDownload);
  } else {
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
