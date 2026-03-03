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
