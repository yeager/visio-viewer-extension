// Service worker: intercept .vsdx file downloads and open in viewer

chrome.downloads.onDeterminingFilename.addListener((item, suggest) => {
  if (item.filename && item.filename.toLowerCase().endsWith('.vsdx')) {
    // Open viewer with download URL
    const viewerUrl = chrome.runtime.getURL('viewer.html') + '?url=' + encodeURIComponent(item.url);
    chrome.tabs.create({ url: viewerUrl });
    // Cancel the download
    chrome.downloads.cancel(item.id);
    suggest({ filename: item.filename });
  }
});

// Intercept navigation to .vsdx URLs
chrome.webNavigation.onBeforeNavigate.addListener((details) => {
  if (details.frameId !== 0) return; // Only top-level
  if (details.url && details.url.toLowerCase().endsWith('.vsdx')) {
    const viewerUrl = chrome.runtime.getURL('viewer.html') + '?url=' + encodeURIComponent(details.url);
    chrome.tabs.update(details.tabId, { url: viewerUrl });
  }
}, { url: [{ urlSuffix: '.vsdx' }] });

// Handle file:// fetch requests from viewer
// This works in the service worker when "Allow access to file URLs" is enabled
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.type === 'fetch-file' && message.url) {
    fetch(message.url)
      .then(resp => {
        if (!resp.ok) throw new Error('HTTP ' + resp.status);
        return resp.arrayBuffer();
      })
      .then(buf => {
        sendResponse({ data: Array.from(new Uint8Array(buf)) });
      })
      .catch(err => {
        sendResponse({ error: err.message || 'Failed to fetch file' });
      });
    return true; // Keep message channel open for async response
  }
});
