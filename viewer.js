/* Visio Viewer — Pure JavaScript + libvisio.js */
(function () {
  "use strict";

  // State
  let converter = null;
  let currentPage = 0;
  let scale = 1;
  let translateX = 0;
  let translateY = 0;
  let isDragging = false;
  let dragStartX = 0, dragStartY = 0;
  let dragStartTX = 0, dragStartTY = 0;

  // DOM
  const loading = document.getElementById("loading");
  const loadingStatus = document.getElementById("loading-status");
  const errorOverlay = document.getElementById("error-overlay");
  const errorMessage = document.getElementById("error-message");
  const welcome = document.getElementById("welcome");
  const welcomeOpen = document.getElementById("welcome-open");
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

  // ---- i18n ----
  function msg(key, ...subs) {
    if (typeof chrome !== "undefined" && chrome.i18n && chrome.i18n.getMessage) {
      const m = chrome.i18n.getMessage(key, subs);
      if (m) return m;
    }
    // Fallback: return key as-is (works outside extension context)
    return key;
  }

  // Apply i18n to data-i18n elements
  document.querySelectorAll("[data-i18n]").forEach(el => {
    const key = el.getAttribute("data-i18n");
    const translated = msg(key);
    if (translated && translated !== key) {
      el.textContent = translated;
    }
  });

  function showError(msgText) {
    errorMessage.textContent = msgText;
    errorOverlay.classList.remove("hidden");
  }

  function setLoading(visible, status) {
    if (status) loadingStatus.textContent = status;
    if (visible) loading.classList.remove("hidden");
    else loading.classList.add("hidden");
  }

  // ---- Convert .vsdx to SVG pages ----
  async function convertFile(arrayBuffer, filename) {
    if (welcome) welcome.classList.add("hidden");
    setLoading(true, msg("parsing"));
    try {
      // Create new converter instance
      converter = new VisioConverter();
      
      // Load and parse the VSDX file
      await converter.loadFromArrayBuffer(arrayBuffer);
      
      const pages = converter.getPages();
      if (pages.length === 0) {
        showError(msg("errorNoPages"));
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
      console.error("Conversion error:", e);
      showError("Failed to parse file: " + (e.message || e));
    }
  }

  // ---- Render SVG ----
  function renderPage(idx) {
    if (!converter) return;
    
    currentPage = idx;
    try {
      const svgContent = converter.renderPage(idx);
      svgContainer.innerHTML = svgContent;
      updatePageControls();
      applyTransform();
    } catch (e) {
      console.error("Render error:", e);
      showError("Failed to render page: " + e.message);
    }
  }

  function updatePageControls() {
    if (!converter) return;
    
    const pages = converter.getPages();
    const total = pages.length;
    
    if (total <= 1) {
      pageInfo.textContent = pages[currentPage]?.name || "—";
      prevBtn.disabled = true;
      nextBtn.disabled = true;
    } else {
      const pageName = pages[currentPage]?.name || (currentPage + 1);
      pageInfo.textContent = `${pageName} (${currentPage + 1}/${total})`;
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
  prevBtn.addEventListener("click", () => { 
    if (currentPage > 0) { 
      renderPage(currentPage - 1); 
      zoomToFit(); 
    } 
  });
  nextBtn.addEventListener("click", () => { 
    if (converter && currentPage < converter.getPages().length - 1) { 
      renderPage(currentPage + 1); 
      zoomToFit(); 
    } 
  });

  // ---- Welcome screen ----
  if (welcomeOpen) {
    welcomeOpen.addEventListener("click", () => fileInput.click());
  }
  if (welcome) {
    welcome.addEventListener("dragover", (e) => e.preventDefault());
    welcome.addEventListener("drop", async (e) => {
      e.preventDefault();
      const file = e.dataTransfer.files[0];
      if (!file) return;
      const buf = await file.arrayBuffer();
      await convertFile(buf, file.name);
    });
  }

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
    setLoading(true, "Initializing viewer…");
    
    // Check if JSZip is available
    if (typeof JSZip === 'undefined') {
      showError("JSZip library not loaded. Please check your installation.");
      setLoading(false);
      return;
    }
    
    // Check if VisioConverter is available
    if (typeof VisioConverter === 'undefined') {
      showError("VisioConverter library not loaded. Please check your installation.");
      setLoading(false);
      return;
    }

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
        showError(msg("errorLoadFailed"));
      }
    } else if (source === "file") {
      // Navigated to a local file:// .vsdx — prompt user to select it
      setLoading(false);
      const fname = params.get("name") || "file.vsdx";
      showFilePrompt(fname);
    } else if (url) {
      await openFromUrl(url);
    } else {
      // No source — show welcome screen
      setLoading(false);
      if (welcome) welcome.classList.remove("hidden");
    }
  }

  main();
})();