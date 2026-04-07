// ── Background service worker for handling downloads ──────────────────────────

/**
 * Listen for download requests from content scripts.
 * Uses chrome.downloads.download() with saveAs: false to skip the dialog.
 */
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.action === 'downloadFile') {
    const { url, filename } = msg;
    
    console.log('[background.js] Downloading file:', url);
    
    chrome.downloads.download({
      url: url,
      filename: filename || 'reservations.xlsx',
      saveAs: false  // Skip the "Save As" dialog
    }, (downloadId) => {
      if (chrome.runtime.lastError) {
        console.error('[background.js] Download failed:', chrome.runtime.lastError.message);
        sendResponse({ success: false, error: chrome.runtime.lastError.message });
      } else {
        console.log('[background.js] Download started, ID:', downloadId);
        sendResponse({ success: true, downloadId });
      }
    });
    
    return true; // Keep channel open for async response
  }
});
