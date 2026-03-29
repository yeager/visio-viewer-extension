#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const { JSDOM } = require('jsdom');
const { DOMParser } = require('@xmldom/xmldom');

// Set up globals for browser environment
const dom = new JSDOM();
global.window = dom.window;
global.document = dom.window.document;
global.DOMParser = DOMParser;
global.JSZip = JSZip;

// Load the converter
const VisioConverter = require('./libvisio.js');

async function debugGeometry() {
    console.log('🔍 Debugging geometry parsing for test4_connectors.vsdx...\n');
    
    // Load test file
    const testFile = path.join(__dirname, 'test-fixtures', 'test4_connectors.vsdx');
    if (!fs.existsSync(testFile)) {
        console.error(`❌ Test file not found: ${testFile}`);
        return;
    }
    
    const buffer = fs.readFileSync(testFile);
    const conv = new VisioConverter();
    
    try {
        await conv.loadFromArrayBuffer(buffer);
        console.log(`✅ Loaded file, found ${conv.pages.length} page(s)`);
        
        if (conv.pages.length === 0) {
            console.error('❌ No pages found');
            return;
        }
        
        // Test page 0
        const page = conv.pages[0];
        const shapes = page.shapes;
        
        console.log(`📄 Page 0 has ${shapes.length} shapes\n`);
        
        // Dump shape data as requested
        shapes.forEach((s, idx) => {
            console.log(`Shape ${idx + 1}: ID=${s.id}, name_u="${s.name_u || ''}", name="${s.name || ''}", type="${s.type || 'Shape'}"`);
            console.log(`  sections:`, Object.keys(s.sections || {}));
            console.log(`  cells:`, Object.keys(s.cells || {}));
            console.log(`  geometry:`, s.geometry ? s.geometry.length : 0, 'sections');
            console.log(`  text: "${s.text || ''}"`);
            console.log(`  master: "${s.master || ''}", master_shape: "${s.master_shape || ''}"`);
            
            // Check for geometry details
            if (s.geometry && s.geometry.length > 0) {
                s.geometry.forEach((geo, geoIdx) => {
                    console.log(`    Geometry ${geoIdx}: ${geo.rows ? geo.rows.length : 0} rows`);
                    if (geo.rows) {
                        geo.rows.forEach((row, rowIdx) => {
                            console.log(`      Row ${rowIdx}: type="${row.type || ''}", cells=${Object.keys(row.cells || {})}`);
                        });
                    }
                });
            }
            
            // Check key cells that affect geometry rendering
            const keyCells = ['PinX', 'PinY', 'Width', 'Height', 'BeginX', 'BeginY', 'EndX', 'EndY'];
            const cellVals = {};
            keyCells.forEach(cell => {
                if (s.cells[cell]) {
                    cellVals[cell] = s.cells[cell].V;
                }
            });
            console.log(`  key cells:`, cellVals);
            console.log('');
        });
        
        // Try to render
        console.log('🎨 Attempting to render SVG...');
        try {
            const svg = conv.renderPage(0);
            console.log(`✅ SVG rendered, length: ${svg.length} characters`);
            
            // Count elements in rendered SVG
            const pathCount = (svg.match(/<path/g) || []).length;
            const lineCount = (svg.match(/<line/g) || []).length;
            const textCount = (svg.match(/<text/g) || []).length;
            const rectCount = (svg.match(/<rect/g) || []).length - 1; // Subtract background rect
            
            console.log(`📊 Elements in JS SVG:`);
            console.log(`  - paths: ${pathCount}`);
            console.log(`  - lines: ${lineCount}`);
            console.log(`  - texts: ${textCount}`);
            console.log(`  - rects: ${rectCount}`);
            console.log(`  - total shapes: ${pathCount + lineCount + rectCount}`);
            
            console.log('\n📊 Elements in Python reference:');
            console.log('  - paths: 5 (3 rectangles + 2 connectors)');
            console.log('  - texts: 3');
            console.log('  - total shapes: 5');
            
            // Save JS output for inspection
            const outputFile = path.join(__dirname, 'debug_test4_connectors_js.svg');
            fs.writeFileSync(outputFile, svg);
            console.log(`💾 JS SVG saved to: ${outputFile}`);
            
        } catch (renderError) {
            console.error('❌ Rendering failed:', renderError);
        }
        
    } catch (error) {
        console.error('❌ Loading failed:', error);
    }
}

debugGeometry().catch(console.error);