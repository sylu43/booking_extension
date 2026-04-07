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

// ── booking.com scraper ───────────────────────────────────────────────────────

function extractBookingReservations() {
  const block = document.querySelector(
    'div.homepage-blocks-wrapper.homepage-selectable-card-wrapper.bui-spacer--large'
  );
  if (!block) {
    console.warn('[content.js] Reservation block not found');
    return [];
  }

  const grid = block.querySelector('ul.bui-list.bui-list--divided.bui-list--text');
  if (!grid) {
    console.warn('[content.js] Reservation list not found');
    return [];
  }

  const reservations = [];

  grid.querySelectorAll('li.bui-list__item').forEach(item => {
    // Guest name
    const nameEl = item.querySelector('.bui-flag__text');
    const name = nameEl ? nameEl.textContent.trim() : '';

    // Order number
    const orderLinkEl = item.querySelector('a[title]');
    const orderNumber = orderLinkEl ? orderLinkEl.getAttribute('title') : '';

    // Room types – direct .bui-f-color-grayscale children of first spacer
    const spacers = item.querySelectorAll('.reservation-overview-item--spacer');
    const roomTypes = [];
    if (spacers[0]) {
      spacers[0].querySelectorAll(':scope > .bui-f-color-grayscale').forEach(div => {
        const text = div.textContent.trim();
        if (text) roomTypes.push(text);
      });
    }

    // Check-in / check-out dates – first .bui-f-color-grayscale in second spacer
    let startDate = '', endDate = '';
    if (spacers[1]) {
      const dateDiv = spacers[1].querySelector(':scope > .bui-f-color-grayscale');
      if (dateDiv) {
        const spans = dateDiv.querySelectorAll('span');
        startDate = spans[0] ? spans[0].textContent.trim() : '';
        endDate   = spans[2] ? spans[2].textContent.trim() : '';
      }
    }

    // Guest count
    let guestCount = '';
    if (spacers[1]) {
      const guestDiv = spacers[1].querySelector('[item]');
      if (guestDiv) {
        guestDiv.querySelectorAll('span').forEach(span => {
          if (span.textContent.includes('adult')) guestCount = span.textContent.trim();
        });
      }
    }

    // Order date – last non-empty span in third spacer
    let orderDate = '';
    if (spacers[2]) {
      const dateDiv = spacers[2].querySelector('.bui-f-color-grayscale');
      if (dateDiv) {
        dateDiv.querySelectorAll('span').forEach(span => {
          const text = span.textContent.trim();
          if (text) orderDate = text;
        });
      }
    }

    reservations.push({
      name,
      orderNumber,
      roomTypes,
      startDate,
      endDate,
      guestCount,
      orderDate,
      source: 'booking.com'
    });
  });

  console.log(`Extracted ${reservations.length} reservations from booking.com`);
  return reservations;
}

// ── Message listener ──────────────────────────────────────────────────────────

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  if (msg.action === 'extractReservations') {
    sendResponse({ reservations: extractBookingReservations() });
  }

  if (msg.action === 'buildCalendar') {
    const { allReservations, year, month } = msg;
    sendResponse({ calendar: buildCalendar(allReservations, year, month) });
  }

  // Legacy action kept for compatibility
  if (msg.action === 'run') {
    sendResponse({ reservations: extractBookingReservations() });
  }

  return true; // keep channel open for async use
});
