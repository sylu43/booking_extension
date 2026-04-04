// Example: click a button, fill a form, etc.
function doAutomation() {
  const block = document.querySelector('div.homepage-blocks-wrapper.homepage-selectable-card-wrapper.bui-spacer--large');
  if (!block) {
    console.warn("block not found");
    return;
  }
  const grid = block.querySelector('ul.bui-list.bui-list--divided.bui-list--text');
  if (!grid) {
    console.warn("Element ui.bui-list not found");
    return;
  }

  const reservations = [];

  grid.querySelectorAll('li.bui-list__item').forEach(item => {
    // Name
    const nameEl = item.querySelector('.bui-flag__text');
    const name = nameEl ? nameEl.textContent.trim() : '';

    // Order number (from the inner <a title="...">)
    const orderLinkEl = item.querySelector('a[title]');
    const orderNumber = orderLinkEl ? orderLinkEl.getAttribute('title') : '';

    // Room types — direct .bui-f-color-grayscale children of the first spacer
    const spacers = item.querySelectorAll('.reservation-overview-item--spacer');
    const roomTypes = [];
    if (spacers[0]) {
      spacers[0].querySelectorAll(':scope > .bui-f-color-grayscale').forEach(div => {
        const text = div.textContent.trim();
        if (text) roomTypes.push(text);
      });
    }

    // Start / end dates — first .bui-f-color-grayscale in second spacer
    let startDate = '', endDate = '';
    if (spacers[1]) {
      const dateDiv = spacers[1].querySelector(':scope > .bui-f-color-grayscale');
      if (dateDiv) {
        const spans = dateDiv.querySelectorAll('span');
        startDate = spans[0] ? spans[0].textContent.trim() : '';
        endDate   = spans[2] ? spans[2].textContent.trim() : '';
      }
    }

    // Guest count — div with [item] attribute in second spacer
    let guestCount = '';
    if (spacers[1]) {
      const guestDiv = spacers[1].querySelector('[item]');
      if (guestDiv) {
        guestDiv.querySelectorAll('span').forEach(span => {
          if (span.textContent.includes('adult')) guestCount = span.textContent.trim();
        });
      }
    }

    // Order date — last non-empty span inside .bui-f-color-grayscale in third spacer
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

    reservations.push({ name, orderNumber, roomTypes, startDate, endDate, guestCount, orderDate });
  });

  console.log("Reservations:", reservations);
  return reservations;
}

// Listen for a trigger from popup
chrome.runtime.onMessage.addListener((msg) => {
  if (msg.action === "run") doAutomation();
});
