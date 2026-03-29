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

async function debugRendering() {
    console.log('🔍 Debugging rendering logic...\n');
    
    // Load test file
    const testFile = path.join(__dirname, 'test-fixtures', 'test4_connectors.vsdx');
    const buffer = fs.readFileSync(testFile);
    const conv = new VisioConverter();
    
    await conv.loadFromArrayBuffer(buffer);
    const page = conv.pages[0];
    const shapes = page.shapes;
    
    console.log('📦 Testing individual shape rendering:\n');
    
    // Test each shape individually
    for (let i = 0; i < shapes.length; i++) {
        const shape = shapes[i];
        console.log(`\n--- Shape ${i + 1} (ID=${shape.id}) ---`);
        console.log(`Type: ${shape.type}, Name: "${shape.name_u || shape.name}"`);
        console.log(`Has text: "${shape.text || ''}"`);
        console.log(`Geometry sections: ${shape.geometry?.length || 0}`);
        
        // Check for 1D connector
        const beginX = shape.cells.BeginX?.V;
        const endX = shape.cells.EndX?.V;
        const is1D = !!(beginX && endX);
        const objType = shape.cells.ObjType?.V;
        const isConnector = is1D || objType === "2";
        
        console.log(`BeginX: ${beginX}, EndX: ${endX}`);
        console.log(`Is 1D: ${is1D}, ObjType: ${objType}, Is Connector: ${isConnector}`);
        
        // Check dimensions
        const wInch = parseFloat(shape.cells.Width?.V || "0");
        const hInch = parseFloat(shape.cells.Height?.V || "0");
        const wPx = Math.abs(wInch) * 72;
        const hPx = Math.abs(hInch) * 72;
        console.log(`Dimensions: ${wInch}" x ${hInch}" (${wPx}px x ${hPx}px)`);
        
        // Test geometry conversion
        if (shape.geometry && shape.geometry.length > 0) {
            console.log('Geometry details:');
            for (let j = 0; j < shape.geometry.length; j++) {
                const geo = shape.geometry[j];
                console.log(`  Geometry ${j}: ${geo.rows?.length || 0} rows`);
                console.log(`    no_fill: ${geo.no_fill}, no_line: ${geo.no_line}, no_show: ${geo.no_show}`);
                
                if (geo.rows) {
                    geo.rows.forEach((row, rowIdx) => {
                        console.log(`    Row ${rowIdx}: type="${row.type}", cells=${JSON.stringify(row.cells)}`);
                    });
                }
                
                // Test path conversion
                try {
                    const pathD = conv.geometryToPath(geo, wInch, hInch, 0, 0);
                    console.log(`    Generated path: "${pathD}"`);
                } catch (error) {
                    console.log(`    ❌ Path conversion failed: ${error.message}`);
                }
            }
        }
        
        // Test render method call 
        try {
            const usedMarkers = new Set();
            const gradients = {};
            const hasShadow = new Set();
            const textLayer = [];
            const pageRels = {};
            
            const svgElements = conv.renderShapeToSvg(
                shape, 11.0, usedMarkers, gradients, hasShadow, textLayer, pageRels
            );
            
            console.log(`Rendered elements: ${svgElements.length}`);
            svgElements.forEach((elem, idx) => {
                console.log(`  Element ${idx}: ${elem.substring(0, 100)}${elem.length > 100 ? '...' : ''}`);
            });
            
        } catch (error) {
            console.log(`❌ Rendering failed: ${error.message}`);
            console.log(error.stack);
        }
    }
}

debugRendering().catch(console.error);