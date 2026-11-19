// Keep Chrome's download intact and open the completed local copy.
const viewer = chrome.runtime.getURL("viewer.html");

function isVisioFile(url) {
  return /\.vsdx$/i.test(url.pathname);
}

function viewerUrl(fileUrl) {
  return viewer + "?url=" + encodeURIComponent(fileUrl);
}

function localFileUrl(filename) {
  const path = filename.replace(/\\/g, "/");
  const encodePath = (value) => value.split("/").map(encodeURIComponent).join("/");
  if (path.startsWith("//")) {
    const firstSeparator = path.indexOf("/", 2);
    if (firstSeparator < 0) return null;
    const host = path.slice(2, firstSeparator);
    return new URL("file://" + host + encodePath(path.slice(firstSeparator)));
  }
  if (/^[A-Za-z]:\//.test(path)) {
    return new URL("file:///" + path.slice(0, 2) + encodePath(path.slice(2)));
  }
  return path.startsWith("/") ? new URL("file://" + encodePath(path)) : null;
}

chrome.webNavigation.onBeforeNavigate.addListener(({frameId, tabId, url}) => {
  if (frameId !== 0) return;
  const parsed = new URL(url);
  if (parsed.protocol === "file:" && isVisioFile(parsed)) {
    chrome.tabs.update(tabId, {url: viewerUrl(parsed.href)}).catch(console.warn);
  }
});

chrome.downloads.onChanged.addListener(async ({id, state}) => {
  if (state?.current !== "complete") return;
  try {
    const [item] = await chrome.downloads.search({id});
    if (!item?.filename) return;
    const url = localFileUrl(item.filename);
    if (url && isVisioFile(url)) await chrome.tabs.create({url: viewerUrl(url.href)});
  } catch (error) {
    console.warn("Could not open the completed Visio download:", error);
  }
});

// Each response goes back to its own viewer request, even after a worker restart.
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.type !== "fetch-file" || !sender.url || new URL(sender.url).pathname !== "/viewer.html" ||
      sender.id !== chrome.runtime.id || !sender.url.startsWith(viewer)) return false;
  try {
    const url = new URL(message.url);
    if (url.protocol !== "file:" || !isVisioFile(url)) throw new Error("Select a local Visio file.");
    fetch(url).then(async (response) => {
      if (!response.ok) throw new Error("HTTP " + response.status);
      sendResponse({data: Array.from(new Uint8Array(await response.arrayBuffer()))});
    }).catch((error) => sendResponse({error: error.message}));
  } catch (error) {
    sendResponse({error: error.message});
    return false;
  }
  return true;
});
