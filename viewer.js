/* Visio Viewer — Pyodide + libvisio-ng */
(function () {
  "use strict";

  const PYODIDE_CDN = "https://cdn.jsdelivr.net/pyodide/v0.26.4/full/";

  // State
  let pyodide = null;
  let svgPages = [];
  let pageNames = [];
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

  function showError(msg) {
    errorMessage.textContent = msg;
    errorOverlay.classList.remove("hidden");
  }

  function setLoading(visible, status) {
    if (status) loadingStatus.textContent = status;
    if (visible) {
      loading.classList.remove("hidden");
    } else {
      loading.classList.add("hidden");
    }
  }

  // ---- Pyodide setup ----
  async function initPyodide() {
    setLoading(true, "Downloading Pyodide runtime…");
    const script = document.createElement("script");
    script.src = PYODIDE_CDN + "pyodide.js";
    document.head.appendChild(script);
    await new Promise((res, rej) => { script.onload = res; script.onerror = rej; });

    setLoading(true, "Initializing Python runtime…");
    pyodide = await loadPyodide({ indexURL: PYODIDE_CDN });

    setLoading(true, "Installing libvisio-ng…");
    await pyodide.loadPackage("micropip");
    await pyodide.runPythonAsync(`
import micropip
await micropip.install("libvisio-ng")
`);
    setLoading(false);
  }

  // ---- Convert .vsdx to SVG pages ----
  async function convertFile(arrayBuffer, filename) {
    setLoading(true, "Parsing Visio file…");
    try {
      // Write file to Pyodide virtual FS
      const uint8 = new Uint8Array(arrayBuffer);
      pyodide.FS.writeFile("/tmp/input.vsdx", uint8);

      // Convert
      const result = await pyodide.runPythonAsync(`
import json, os
from libvisio_ng import convert, get_page_info

pages_info = get_page_info("/tmp/input.vsdx")
page_names = [p.get("name", f"Page {p['index']+1}") for p in pages_info]

svg_files = convert("/tmp/input.vsdx", output_dir="/tmp/svg_out")
svgs = []
for f in svg_files:
    with open(f, "r") as fh:
        svgs.append(fh.read())

json.dumps({"svgs": svgs, "names": page_names})
`);
      const data = JSON.parse(result);
      svgPages = data.svgs;
      pageNames = data.names;

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
    // Apply current transform
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
    const sw = svg.getBoundingClientRect().width / scale || svg.clientWidth || 800;
    const sh = svg.getBoundingClientRect().height / scale || svg.clientHeight || 600;
    // Get natural size
    const natW = svg.viewBox?.baseVal?.width || sw;
    const natH = svg.viewBox?.baseVal?.height || sh;

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

  // ---- Init ----
  async function main() {
    await initPyodide();

    // Check URL param
    const params = new URLSearchParams(location.search);
    const url = params.get("url");
    if (url) {
      setLoading(true, "Downloading file…");
      try {
        const resp = await fetch(url);
        if (!resp.ok) throw new Error("HTTP " + resp.status);
        const buf = await resp.arrayBuffer();
        const filename = url.split("/").pop().split("?")[0] || "file.vsdx";
        await convertFile(buf, decodeURIComponent(filename));
      } catch (e) {
        setLoading(false);
        showError("Failed to download file: " + e.message);
      }
    } else {
      setLoading(false);
    }
  }

  main();
})();
