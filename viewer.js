/* Visio Viewer — Sandboxed Pyodide + libvisio-ng */
(function () {
  "use strict";

  // State
  let svgPages = [];
  let pageNames = [];
  let currentPage = 0;
  let scale = 1;
  let translateX = 0;
  let translateY = 0;
  let isDragging = false;
  let dragStartX = 0, dragStartY = 0;
  let dragStartTX = 0, dragStartTY = 0;
  let sandboxReady = false;
  let pendingCallbacks = {};
  let callId = 0;

  // DOM
  const loading = document.getElementById("loading");
  const loadingStatus = document.getElementById("loading-status");
  const errorOverlay = document.getElementById("error-overlay");
  const errorMessage = document.getElementById("error-message");
  const fileTitle = document.getElementById("file-title");
  const pageInfo = document.getElementById("page-info");
  const prevBtn = document.getElementById("prev-page");
  const nextBtn = document.getElementById("next-page");
  const zoomInBtn = document.getElementById("zoom-in");
  const zoomOutBtn = document.getElementById("zoom-out");
  const zoomFitBtn = document.getElementById("zoom-fit");
  const zoomLevel = document.getElementById("zoom-level");
  const viewport = document.getElementById("viewport");
  const svgContainer = document.getElementById("svg-container");
  const fileInput = document.getElementById("file-input");
  const dropZone = document.getElementById("drop-zone");

  // Sandbox iframe
  const sandbox = document.getElementById("sandbox-frame");

  function showError(msg) {
    errorMessage.textContent = msg;
    errorOverlay.classList.remove("hidden");
  }

  function setLoading(visible, status) {
    if (status) loadingStatus.textContent = status;
    if (visible) loading.classList.remove("hidden");
    else loading.classList.add("hidden");
  }

  // Send message to sandbox and return promise
  function sandboxCall(type, data) {
    return new Promise((resolve, reject) => {
      const id = ++callId;
      pendingCallbacks[id] = { resolve, reject };
      const msg = { type, id };
      if (data !== undefined) msg.data = data;
      sandbox.contentWindow.postMessage(msg, "*");
    });
  }

  // Listen for sandbox messages
  window.addEventListener("message", (event) => {
    const msg = event.data;
    if (!msg || !msg.type) return;

    if (msg.type === "sandbox-ready") {
      sandboxReady = true;
      return;
    }

    const cb = pendingCallbacks[msg.id];
    if (!cb) return;
    delete pendingCallbacks[msg.id];

    if (msg.type === "error") {
      cb.reject(new Error(msg.error));
    } else if (msg.type === "init-done") {
      cb.resolve();
    } else if (msg.type === "convert-done") {
      cb.resolve(msg.result);
    }
  });

  // Wait for sandbox ready
  function waitForSandbox() {
    if (sandboxReady) return Promise.resolve();
    return new Promise((resolve) => {
      const handler = (event) => {
        if (event.data && event.data.type === "sandbox-ready") {
          sandboxReady = true;
          window.removeEventListener("message", handler);
          resolve();
        }
      };
      window.addEventListener("message", handler);
    });
  }

  // ---- Init Pyodide via sandbox ----
  async function initPyodide() {
    setLoading(true, "Waiting for sandbox…");
    await waitForSandbox();

    setLoading(true, "Initializing Python runtime…");
    await sandboxCall("init");
    setLoading(false);
  }

  // ---- Convert .vsdx to SVG pages ----
  async function convertFile(arrayBuffer, filename) {
    setLoading(true, "Parsing Visio file…");
    try {
      // Transfer as ArrayBuffer
      const data = Array.from(new Uint8Array(arrayBuffer));
      const result = await sandboxCall("convert", data);
      svgPages = result.svgs;
      pageNames = result.names;

      if (svgPages.length === 0) {
        showError("No pages found in the Visio file.");
        setLoading(false);
        return;
      }

      fileTitle.textContent = filename || "Visio Viewer";
      currentPage = 0;
      renderPage(0);
      updatePageControls();
      zoomToFit();
      setLoading(false);
    } catch (e) {
      setLoading(false);
      showError("Failed to parse file: " + (e.message || e));
    }
  }

  // ---- Render SVG ----
  function renderPage(idx) {
    currentPage = idx;
    svgContainer.innerHTML = svgPages[idx];
    updatePageControls();
    applyTransform();
  }

  function updatePageControls() {
    const total = svgPages.length;
    if (total <= 1) {
      pageInfo.textContent = pageNames[currentPage] || "—";
      prevBtn.disabled = true;
      nextBtn.disabled = true;
    } else {
      pageInfo.textContent = `${pageNames[currentPage] || (currentPage + 1)} (${currentPage + 1}/${total})`;
      prevBtn.disabled = currentPage === 0;
      nextBtn.disabled = currentPage === total - 1;
    }
  }

  // ---- Transform ----
  function applyTransform() {
    svgContainer.style.transform = `translate(${translateX}px, ${translateY}px) scale(${scale})`;
    zoomLevel.textContent = Math.round(scale * 100) + "%";
  }

  function zoomToFit() {
    const svg = svgContainer.querySelector("svg");
    if (!svg) return;
    const vw = viewport.clientWidth;
    const vh = viewport.clientHeight;
    const natW = svg.viewBox?.baseVal?.width || svg.clientWidth || 800;
    const natH = svg.viewBox?.baseVal?.height || svg.clientHeight || 600;
    const fitScale = Math.min((vw - 40) / natW, (vh - 40) / natH, 2);
    scale = fitScale;
    translateX = (vw - natW * scale) / 2;
    translateY = (vh - natH * scale) / 2;
    applyTransform();
  }

  // ---- Zoom ----
  function zoomBy(factor, cx, cy) {
    if (cx === undefined) { cx = viewport.clientWidth / 2; cy = viewport.clientHeight / 2; }
    const newScale = Math.max(0.1, Math.min(10, scale * factor));
    const ratio = newScale / scale;
    translateX = cx - ratio * (cx - translateX);
    translateY = cy - ratio * (cy - translateY);
    scale = newScale;
    applyTransform();
  }

  // ---- Pan ----
  viewport.addEventListener("mousedown", (e) => {
    if (e.button !== 0) return;
    isDragging = true;
    dragStartX = e.clientX; dragStartY = e.clientY;
    dragStartTX = translateX; dragStartTY = translateY;
    viewport.classList.add("dragging");
  });
  window.addEventListener("mousemove", (e) => {
    if (!isDragging) return;
    translateX = dragStartTX + (e.clientX - dragStartX);
    translateY = dragStartTY + (e.clientY - dragStartY);
    applyTransform();
  });
  window.addEventListener("mouseup", () => {
    isDragging = false;
    viewport.classList.remove("dragging");
  });

  // ---- Scroll zoom ----
  viewport.addEventListener("wheel", (e) => {
    e.preventDefault();
    const factor = e.deltaY < 0 ? 1.15 : 1 / 1.15;
    const rect = viewport.getBoundingClientRect();
    zoomBy(factor, e.clientX - rect.left, e.clientY - rect.top);
  }, { passive: false });

  // ---- Button handlers ----
  zoomInBtn.addEventListener("click", () => zoomBy(1.25));
  zoomOutBtn.addEventListener("click", () => zoomBy(0.8));
  zoomFitBtn.addEventListener("click", zoomToFit);
  prevBtn.addEventListener("click", () => { if (currentPage > 0) { renderPage(currentPage - 1); zoomToFit(); } });
  nextBtn.addEventListener("click", () => { if (currentPage < svgPages.length - 1) { renderPage(currentPage + 1); zoomToFit(); } });

  // ---- File input ----
  fileInput.addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const buf = await file.arrayBuffer();
    await convertFile(buf, file.name);
  });

  // ---- Drag & drop ----
  viewport.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropZone.classList.remove("hidden");
  });
  viewport.addEventListener("dragleave", (e) => {
    if (!viewport.contains(e.relatedTarget)) dropZone.classList.add("hidden");
  });
  viewport.addEventListener("drop", async (e) => {
    e.preventDefault();
    dropZone.classList.add("hidden");
    const file = e.dataTransfer.files[0];
    if (!file) return;
    const buf = await file.arrayBuffer();
    await convertFile(buf, file.name);
  });

  // ---- Keyboard ----
  document.addEventListener("keydown", (e) => {
    if (e.key === "ArrowLeft") prevBtn.click();
    else if (e.key === "ArrowRight") nextBtn.click();
    else if (e.key === "=" || e.key === "+") zoomBy(1.25);
    else if (e.key === "-") zoomBy(0.8);
    else if (e.key === "0") zoomToFit();
  });

  // ---- Fetch URL (with file:// support via background service worker) ----
  async function fetchUrl(url) {
    if (url.startsWith("file://")) {
      // file:// URLs can't be fetched from extension pages directly
      // Route through background service worker
      return new Promise((resolve, reject) => {
        chrome.runtime.sendMessage({ type: "fetch-file", url: url }, (response) => {
          if (chrome.runtime.lastError) {
            reject(new Error("Could not access local file. Use the file picker or drag & drop instead."));
            return;
          }
          if (response && response.error) {
            reject(new Error(response.error));
            return;
          }
          if (response && response.data) {
            resolve(new Uint8Array(response.data).buffer);
          } else {
            reject(new Error("Empty response from background worker."));
          }
        });
      });
    }

    const resp = await fetch(url);
    if (!resp.ok) throw new Error("HTTP " + resp.status);
    return resp.arrayBuffer();
  }

  async function openFromUrl(url) {
    setLoading(true, "Downloading file…");
    try {
      const buf = await fetchUrl(url);
      const filename = decodeURIComponent(url.split("/").pop().split("?")[0] || "file.vsdx");
      await convertFile(buf, filename);
    } catch (e) {
      setLoading(false);
      if (url.startsWith("file://")) {
        // file:// fetch failed — show friendly file picker prompt
        const fname = decodeURIComponent(url.split("/").pop().split("?")[0] || "file.vsdx");
        showFilePrompt(fname);
      } else {
        let msg = "Failed to download file: " + e.message;
        msg += "\n\nTry drag & drop or the file picker for local files.";
        showError(msg);
      }
    }
  }

  // ---- Open URL button ----
  const openUrlBtn = document.getElementById("open-url");
  openUrlBtn.addEventListener("click", () => {
    const url = prompt("Enter URL to a .vsdx file:");
    if (url && url.trim()) {
      openFromUrl(url.trim());
    }
  });

  // ---- Show file picker prompt for local files ----
  function showFilePrompt(filename) {
    const overlay = document.createElement("div");
    overlay.className = "file-prompt-overlay";
    overlay.innerHTML = `
      <div class="file-prompt-box">
        <h2>📁 Local File Detected</h2>
        <p>Chrome extensions cannot directly read local files.<br>
        Please select <strong>${filename}</strong> using the file picker below.</p>
        <label class="file-prompt-btn">
          Select File
          <input type="file" accept=".vsdx,.vsd,.vsdm" style="display:none">
        </label>
        <p class="file-prompt-hint">Or drag & drop the file onto this page.</p>
      </div>
    `;
    document.body.appendChild(overlay);

    const input = overlay.querySelector("input[type=file]");
    input.addEventListener("change", async () => {
      const file = input.files[0];
      if (!file) return;
      overlay.remove();
      const buf = await file.arrayBuffer();
      await convertFile(buf, file.name);
    });
  }

  // ---- Load pending file data from background (for file:// URLs) ----
  async function loadPendingFile() {
    return new Promise((resolve) => {
      chrome.runtime.sendMessage({ type: "get-pending-file" }, (response) => {
        if (chrome.runtime.lastError || !response || response.error) {
          resolve(null);
          return;
        }
        resolve(response);
      });
    });
  }

  // ---- Init ----
  async function main() {
    await initPyodide();

    const params = new URLSearchParams(location.search);
    const source = params.get("source");
    const url = params.get("url");

    if (source === "local") {
      // File loaded via content script → get data from background
      const pending = await loadPendingFile();
      if (pending && pending.data) {
        const buf = new Uint8Array(pending.data).buffer;
        await convertFile(buf, pending.filename || "file.vsdx");
      } else {
        setLoading(false);
        showError("Failed to load local file. Try drag & drop or the file picker.");
      }
    } else if (source === "file") {
      // Navigated to a local file:// .vsdx — prompt user to select it
      setLoading(false);
      const fname = params.get("name") || "file.vsdx";
      showFilePrompt(fname);
    } else if (url) {
      await openFromUrl(url);
    } else {
      setLoading(false);
    }
  }

  main();
})();
