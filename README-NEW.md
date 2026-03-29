# Visio Viewer (Pure JavaScript Edition)

A Chrome extension for viewing Microsoft Visio (.vsdx) files directly in the browser using pure JavaScript.

## What's New

This version has been completely rewritten to eliminate Python/Pyodide/WASM dependencies, fixing Windows compatibility issues while maintaining VSDX viewing functionality.

### Changes in v2.0.0

- ✅ **Pure JavaScript** - No Python, no WASM, no Pyodide
- ✅ **Windows Compatible** - Works on all Chrome platforms  
- ✅ **Manifest V3** - No sandbox or unsafe-eval CSP requirements
- ✅ **Smaller Size** - Removed 50MB+ Pyodide runtime
- ✅ **Better Performance** - Direct JavaScript execution
- ✅ **JSZip Integration** - Reliable ZIP file handling

### Removed Features (from v1.x)
- ❌ Binary .vsd file support (XML-based .vsdx only)
- ❌ Advanced shape effects and complex gradients
- ❌ Master shape inheritance (simplified)

## Installation

### For Development
1. Clone this repository
2. Open Chrome and go to `chrome://extensions/`
3. Enable "Developer mode" (top right)
4. Click "Load unpacked" and select the extension folder

### For Production
1. Download the latest release
2. Follow the same steps as development installation

## Technical Details

### Architecture
- `libvisio.js` - Pure JavaScript VSDX parser and SVG renderer
- `jszip.min.js` - ZIP file handling for VSDX archives
- `viewer.js` - Main UI and viewer logic
- `background.js` - Service worker for file interception

### Supported VSDX Features
- ✅ Basic shapes (rectangles, ellipses, lines)
- ✅ Custom geometry paths
- ✅ Fill and stroke colors
- ✅ Text rendering
- ✅ Multi-page documents
- ✅ Shape positioning and sizing
- ⚠️ Partial theme support (basic colors)

### File Support
- ✅ `.vsdx` (XML-based Visio format)
- ✅ `.vsdm` (macro-enabled VSDX)
- ❌ `.vsd` (legacy binary format) - may be added in future

## Testing

Use the included `test.html` file for development testing:

```bash
cd /path/to/extension
open test.html
```

Test with the included `simple-test.vsdx` sample file.

## Development

### Key Files
- `libvisio.js` - Core VSDX parsing logic
- `viewer.js` - UI and rendering
- `manifest.json` - Extension manifest
- `test.html` - Development testing page

### Adding Features
1. Modify `libvisio.js` for VSDX parsing improvements
2. Update `viewer.js` for UI enhancements  
3. Test with `test.html` before full extension testing

### Known Limitations
- Binary .vsd files not supported
- Complex gradients simplified to solid fills
- Master shape inheritance limited
- Theme colors partially supported
- No connector routing algorithms

## Comparison with Previous Version

| Feature | v1.x (Pyodide) | v2.0 (Pure JS) |
|---------|----------------|----------------|
| Windows Support | ❌ Broken | ✅ Works |
| File Size | ~50MB | ~500KB |
| Load Time | 10-30s | <1s |
| Binary .vsd | ✅ | ❌ |
| Complex Shapes | ✅ | ⚠️ Basic |
| Themes | ✅ | ⚠️ Partial |
| Performance | Slow | Fast |

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make changes to `libvisio.js` or other files
4. Test with `test.html` and sample files
5. Submit a pull request

## License

Same license as the original project (see LICENSE file).

## Troubleshooting

**Extension won't load**: Check that all files are present and manifest.json is valid.

**VSDX files don't open**: Ensure the file is a valid .vsdx (XML-based) format, not legacy .vsd.

**Shapes don't render correctly**: Check browser console for errors. Some complex Visio features may not be fully supported yet.

**Text positioning issues**: Text rendering uses simplified positioning logic compared to full Visio.

## Migration from v1.x

No migration needed - just replace the extension files. Note that some advanced features may render differently or be simplified in the pure JavaScript version.