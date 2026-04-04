// Example: click a button, fill a form, etc.
function doAutomation() {
  const grid = document.querySelector('.av-monthly__grid');
  if (grid) {
    console.log("Found element:", grid);
    // Do something with it, e.g.:
    // grid.style.border = "2px solid red"; // highlight it
    // console.log(grid.innerHTML);
  } else {
    console.warn("Element .av-monthly__grid not found");
  }
}

// Listen for a trigger from popup
chrome.runtime.onMessage.addListener((msg) => {
  if (msg.action === "run") doAutomation();
});
