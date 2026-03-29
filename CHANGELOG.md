# Changelog

## 2.0.0 (2026-03-29)

Complete rewrite from Python/Pyodide to pure JavaScript.

### Changed
- **Architecture**: Replaced Pyodide (Python in WebAssembly) with pure JavaScript VSDX parser
- **Size**: 50 MB → 66 KB (750x smaller)
- **Load time**: 10-30s → <1s
- **Manifest**: Removed sandbox, wasm-unsafe-eval, scripting permission

### Added
- `libvisio.js` — complete VSDX→SVG converter ported from libvisio-ng Python
- Full geometry rendering (MoveTo, LineTo, ArcTo, EllipticalArcTo, RelMoveTo, RelLineTo, NURBS, Ellipse)
- Master shape / stencil support with cell inheritance
- Theme and stylesheet color resolution
- 1D connector rendering with arrows and text labels
- Drop shadows, gradients, rounded corners
- Auto text contrast (white on dark, dark on light)
- Connector text labels with semi-transparent backgrounds
- Group shape rendering with nested elements
- Embedded image support (PNG/JPG as data URI)
- Welcome screen with drag & drop prompt
- i18n support (English + Swedish via Chrome _locales)
- Keyboard shortcuts (arrows, +/-, 0)

### Removed
- Pyodide runtime (~50 MB)
- sandbox.html, sandbox-worker.js
- wasm-unsafe-eval CSP requirement
- scripting permission

### Fixed
- Windows compatibility (Pyodide crashed on Windows Chrome)
- Node.ELEMENT_NODE reference for cross-environment compatibility

## 1.2.1

- Initial Pyodide-based release
- Python libvisio-ng running in WebAssembly sandbox
