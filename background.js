// Service worker: intercept .vsdx file downloads and open in viewer

// Pending file data to pass to viewer tabs
let pendingFileData = null;
// Track tabs we've already redirected to avoid loops
const redirectedTabs = new Set();

// Handle .vsdx URL: redirect tab to viewer
function handleVsdxUrl(tabId, url) {
  if (redirectedTabs.has(tabId)) return;
  if (!url || !url.toLowerCase().endsWith('.vsdx')) return;
  if (url.startsWith('chrome-extension://')) return;

  redirectedTabs.add(tabId);
  setTimeout(() => redirectedTabs.delete(tabId), 10000);

  const viewerUrl = chrome.runtime.getURL('viewer.html') + '?url=' + encodeURIComponent(url);
  chrome.tabs.update(tabId, { url: viewerUrl });
}

// Use tabs.onUpdated to detect .vsdx navigation (no webNavigation permission needed)
chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
  if (changeInfo.url) {
    handleVsdxUrl(tabId, changeInfo.url);
  }
  if (changeInfo.status === 'complete' && tab.url) {
    handleVsdxUrl(tabId, tab.url);
  }
});

// Intercept .vsdx downloads
chrome.downloads.onDeterminingFilename.addListener((item, suggest) => {
  if (item.filename && item.filename.toLowerCase().endsWith('.vsdx')) {
    const viewerUrl = chrome.runtime.getURL('viewer.html') + '?url=' + encodeURIComponent(item.url);
    chrome.tabs.create({ url: viewerUrl });
    chrome.downloads.cancel(item.id);
    suggest({ filename: item.filename });
  }
});

// Handle messages from content script and viewer
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.type === 'file-data' && message.data) {
    pendingFileData = { data: message.data, filename: message.filename };
    const viewerUrl = chrome.runtime.getURL('viewer.html') + '?source=local';
    chrome.tabs.update(sender.tab.id, { url: viewerUrl });
    return;
  }

  if (message.type === 'get-pending-file') {
    const data = pendingFileData;
    pendingFileData = null;
    sendResponse(data || { error: 'No pending file data' });
    return;
  }

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
    return true;
  }
});
