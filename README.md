# Visio Viewer — Chrome Extension

View Microsoft Visio (.vsdx) files directly in Chrome. No plugins, no cloud services, no Visio license needed.

## Features

- **Pure JavaScript** — no Python, no WebAssembly, no server
- **Local processing** — bundled JavaScript dependencies; no server required
- **Rendering support** — shapes, connectors, text, colors, gradients, shadows, arrows
- **Multi-page** — navigate between pages
- **Zoom & pan** — scroll to zoom, drag to pan, keyboard shortcuts
- **Drag & drop** — drop a .vsdx file onto the viewer
- **Automatic opening** — opens completed .vsdx downloads and local .vsdx navigation in the viewer
- **i18n** — English and Swedish
- **Works everywhere** — Windows, macOS, Linux

When local file access is disabled for the extension in Chromium, select the saved file in the viewer or drag it onto the prompt.

## Rendering Quality

Tested against 16 reference files with 100% element parity on 14/16 files, ≥90% on all 16.

Supported:
- Shapes with geometry (MoveTo, LineTo, ArcTo, EllipticalArcTo, RelMoveTo, RelLineTo, NURBS)
- Master shapes / stencils with inheritance
- 1D connectors with arrows and text labels
- Themes, stylesheets, color resolution
- Group shapes with nested elements
- Embedded images (PNG/JPG)
- Rounded corners, drop shadows, gradients
- Auto text contrast (white text on dark backgrounds)
- Multi-line text with formatting
- Semi-transparent connector labels

## Install

### Chrome Web Store
*(Coming soon)*

### Manual
1. Download the latest release `.zip`
2. Go to `chrome://extensions`
3. Enable **Developer mode**
4. Click **Load unpacked** and select the extracted folder

## Usage

- **Drag & drop** a `.vsdx` file onto the viewer tab
- **Click the folder icon** in the toolbar to open a file
- Browse to a `.vsdx` URL — it opens automatically
- Navigate pages with ← → or arrow keys
- Zoom with +/- or scroll wheel
- Press 0 to fit to window

## Tech

| | v1 (Pyodide) | v2 (Pure JS) |
|---|---|---|
| Runtime | Python + WASM | Bundled JavaScript |
| Load time | Includes runtime startup | Depends on document size |
| Windows | ❌ | ✅ |
| Dependencies | Python + WASM | JSZip only |
| .vsdx support | ✅ | ✅ |
| .vsd (binary) | ✅ | ❌ |

Built by porting [libvisio-ng](https://github.com/nicoptere/libvisio-ng) (Python) to JavaScript.

## Development

```bash
# Test with Node.js
npm ci
npm test
```

## License

MIT

Tests cover file-opening races, error recovery, local-file prompts, zoom bounds, and a generated VSDX archive through the real converter. Browser UI tests are still recommended before Chrome Web Store submission.
