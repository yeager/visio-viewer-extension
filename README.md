# Visio Viewer — Chrome Extension

View Microsoft Visio (.vsdx) files directly in Chrome without needing Visio installed.

Uses [libvisio-ng](https://github.com/yeager/libvisio-ng) running via [Pyodide](https://pyodide.org/) (Python in WebAssembly) for fully client-side rendering — no server required.

## Features

- **Open .vsdx files** via drag & drop, file picker, or direct URL
- **SVG rendering** — crisp vector output at any zoom level
- **Zoom & pan** — scroll to zoom, drag to pan, zoom-to-fit button
- **Multi-page support** — navigate between pages with arrow buttons or keyboard
- **Light/dark theme** — follows your system preference
- **Zero server dependencies** — everything runs in the browser

## Installation

1. Clone or download this repository
2. Open `chrome://extensions/` in Chrome
3. Enable **Developer mode** (top right toggle)
4. Click **Load unpacked** and select this folder
5. The Visio Viewer extension icon appears in your toolbar

## Usage

### Drag & drop
Open the extension page (`viewer.html`) and drag a `.vsdx` file onto the window.

### File picker
Click the folder icon in the toolbar to browse for a file.

### Direct URL
The extension intercepts `.vsdx` file downloads and opens them in the viewer automatically.

### Keyboard shortcuts

| Key | Action |
|-----|--------|
| `←` / `→` | Previous / next page |
| `+` / `-` | Zoom in / out |
| `0` | Zoom to fit |

## How It Works

1. Extension loads [Pyodide](https://pyodide.org/) (Python WASM runtime) from CDN
2. Installs `libvisio-ng` via `micropip`
3. Parses the `.vsdx` file (which is a ZIP of XML) using libvisio-ng's built-in parser
4. Renders each page as SVG
5. Displays SVG with interactive zoom/pan controls

First load takes a few seconds while Pyodide initializes. Subsequent file opens are fast.

## Technical Details

- Chrome Extension Manifest V3
- No external dependencies beyond Pyodide CDN
- All processing happens client-side
- Supports `.vsdx`, `.vsdm` formats (XML-based Visio files)

## Development

```bash
# Clone
git clone https://github.com/yeager/visio-viewer-extension
cd visio-viewer-extension

# Load in Chrome as unpacked extension
# Make changes, then reload extension in chrome://extensions/
```

## License

MIT
