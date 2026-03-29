#!/usr/bin/env node

/**
 * Debug script för att inspektera pilar i VSDX
 */

const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');

// Setup JSDOM environment
const dom = new JSDOM('<!DOCTYPE html><html><body></body></html>');
global.window = dom.window;
global.document = dom.window.document;
global.DOMParser = dom.window.DOMParser;

// Läs JSZip
const JSZip = require('jszip');
global.JSZip = JSZip;

// Läs libvisio.js
const LibVisioCode = fs.readFileSync(path.join(__dirname, 'libvisio.js'), 'utf8');
eval(LibVisioCode);

async function debugArrows() {
    console.log('🔍 Debugging arrows in VSDX...\n');
    
    const testFile = '/tmp/complex-zscaler.vsdx';
    const buffer = fs.readFileSync(testFile);
    
    const converter = new dom.window.VisioConverter();
    await converter.loadFromArrayBuffer(buffer);
    
    const pages = converter.getPages();
    console.log(`Found ${pages.length} pages\n`);
    
    // Inspektera första sidan
    const page = converter.pages[0];
    console.log('=== PAGE 0 SHAPES ===');
    
    page.shapes.forEach((shape, i) => {
        console.log(`\nShape ${i}: ${shape.name || shape.id}`);
        console.log(`  Type: ${shape.type}`);
        
        // Kolla linjer specifikt
        const hasBeginX = shape.cells.BeginX?.V !== undefined;
        const hasEndX = shape.cells.EndX?.V !== undefined;
        
        if (hasBeginX && hasEndX) {
            console.log('  🔗 CONNECTOR/LINE:');
            console.log(`    BeginX: ${shape.cells.BeginX?.V}`);
            console.log(`    BeginY: ${shape.cells.BeginY?.V}`);
            console.log(`    EndX: ${shape.cells.EndX?.V}`);
            console.log(`    EndY: ${shape.cells.EndY?.V}`);
            console.log(`    BeginArrow: ${shape.cells.BeginArrow?.V || 'undefined'}`);
            console.log(`    EndArrow: ${shape.cells.EndArrow?.V || 'undefined'}`);
            console.log(`    LineColor: ${shape.cells.LineColor?.V || 'undefined'}`);
        }
        
        // Lista alla celler för debugging
        const cellNames = Object.keys(shape.cells);
        if (cellNames.length > 0) {
            console.log(`  Cells: ${cellNames.slice(0,10).join(', ')}${cellNames.length > 10 ? '...' : ''}`);
        }
    });
}

debugArrows().catch(console.error);