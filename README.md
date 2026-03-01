# Visio Viewer — Chrome Extension

View Microsoft Visio (.vsdx) files directly in Chrome using [libvisio-ng](https://github.com/nicoptere/libvisio-ng) and [Pyodide](https://pyodide.org/).

## Features

- Open `.vsdx` files via drag & drop, file picker, or by clicking download links
- Multi-page navigation with page names
- Pan, zoom (mouse wheel), and zoom-to-fit
- Keyboard shortcuts: ← → for pages, +/- for zoom, 0 for fit
- **Fully offline** — Pyodide and all dependencies are bundled locally

## Architecture

The extension uses a **sandboxed iframe** approach to run Pyodide:

- `viewer.html` — Main UI with toolbar, viewport, and controls
- `sandbox.html` — Sandboxed page with relaxed CSP that loads Pyodide + libvisio-ng
- Communication between viewer and sandbox via `postMessage`
- This avoids MV3 Content Security Policy restrictions on `eval()` and WASM

## Installation

1. Clone or download this repository
2. Open `chrome://extensions/` in Chrome
3. Enable **Developer mode** (top right)
4. Click **Load unpacked** and select the extension folder
5. The extension is ~15MB due to bundled Pyodide runtime

## Usage

- **Drag & drop** a `.vsdx` file onto the viewer
- **Click the folder icon** in the toolbar to open a file picker
- **Click a `.vsdx` download link** — the extension intercepts it and opens the viewer

## Bundled Dependencies

- Pyodide v0.26.4 (Python runtime compiled to WebAssembly)
- micropip (Python package installer for Pyodide)
- libvisio-ng (Visio file parser)
- olefile (OLE2 file parser, dependency of libvisio-ng)

All dependencies are bundled in the `pyodide/` directory — no internet connection required.

## Development

To update Pyodide:
1. Download the release from https://github.com/pyodide/pyodide/releases
2. Extract only the core files: `pyodide.js`, `pyodide.asm.js`, `pyodide.asm.wasm`, `pyodide-lock.json`, `package.json`, `python_stdlib.zip`
3. Place in `pyodide/` directory along with the wheel files

## License

MIT
