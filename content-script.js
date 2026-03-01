// Content script for file://*.vsdx URLs
// Content scripts CAN fetch file:// URLs (unlike service workers)
(async function() {
  try {
    const url = window.location.href;
    const resp = await fetch(url);
    const buf = await resp.arrayBuffer();
    const data = Array.from(new Uint8Array(buf));
    const filename = decodeURIComponent(url.split('/').pop() || 'file.vsdx');
    
    // Send file data to background, which will open the viewer
    chrome.runtime.sendMessage({
      type: 'file-data',
      data: data,
      filename: filename
    });
  } catch (e) {
    console.error('Visio Viewer: Failed to read file:', e);
  }
})();
