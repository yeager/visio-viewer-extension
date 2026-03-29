#!/usr/bin/env node

/**
 * Test script för libvisio.js rendering-kvalitet
 * Laddar complex-zscaler.vsdx och testar SVG-rendering
 */

const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');

// Setup JSDOM environment
const dom = new JSDOM('<!DOCTYPE html><html><body></body></html>', {
    url: "http://localhost",
    pretendToBeVisual: true,
    resources: "usable"
});

global.window = dom.window;
global.document = dom.window.document;
global.DOMParser = dom.window.DOMParser;

// Läs JSZip
const JSZip = require('jszip');
global.JSZip = JSZip;

// Läs libvisio.js
const LibVisioCode = fs.readFileSync(path.join(__dirname, 'libvisio.js'), 'utf8');
eval(LibVisioCode);

async function testRendering() {
    console.log('🔍 Testing libvisio.js rendering quality...\n');
    
    const testFile = '/tmp/complex-zscaler.vsdx';
    
    if (!fs.existsSync(testFile)) {
        console.error(`❌ Test file not found: ${testFile}`);
        process.exit(1);
    }
    
    try {
        // Läs VSDX filen
        const buffer = fs.readFileSync(testFile);
        console.log(`📁 Loaded ${testFile} (${buffer.length} bytes)`);
        
        // Skapa VisioConverter instans
        const converter = new dom.window.VisioConverter();
        
        // Ladda och parsa filen
        await converter.loadFromArrayBuffer(buffer);
        
        // Få lista av sidor
        const pages = converter.getPages();
        console.log(`📄 Found ${pages.length} page(s):`);
        pages.forEach((page, i) => {
            console.log(`  Page ${i}: "${page.name}" (${page.width}"×${page.height}")`);
        });
        
        // Rendera alla sidor
        for (let i = 0; i < pages.length; i++) {
            console.log(`\n🎨 Rendering page ${i} (${pages[i].name})...`);
            
            const svg = converter.renderPage(i);
            const outputPath = `/tmp/complex-zscaler-page${i}.svg`;
            
            fs.writeFileSync(outputPath, svg, 'utf8');
            console.log(`✅ SVG saved to: ${outputPath}`);
            
            // Analysera SVG-innehåll för kvalitetsindikatorer
            analyzeRendering(svg, i);
        }
        
        console.log('\n🏁 Test completed successfully!');
        console.log('📝 Open the SVG files to verify visual quality.');
        
    } catch (error) {
        console.error('❌ Test failed:', error);
        process.exit(1);
    }
}

function analyzeRendering(svg, pageIndex) {
    console.log(`   📊 Analyzing page ${pageIndex} rendering:`);
    
    // Räkna olika SVG-element
    const rectCount = (svg.match(/<rect/g) || []).length;
    const lineCount = (svg.match(/<line/g) || []).length;
    const pathCount = (svg.match(/<path/g) || []).length;
    const textCount = (svg.match(/<text/g) || []).length;
    const ellipseCount = (svg.match(/<ellipse/g) || []).length;
    const groupCount = (svg.match(/<g/g) || []).length;
    
    console.log(`      • Rectangles: ${rectCount}`);
    console.log(`      • Lines: ${lineCount}`);
    console.log(`      • Paths: ${pathCount}`);
    console.log(`      • Text elements: ${textCount}`);
    console.log(`      • Ellipses: ${ellipseCount}`);
    console.log(`      • Groups: ${groupCount}`);
    
    // Kontrollera qualitets-features
    const hasRoundedCorners = svg.includes('rx=') || svg.includes('ry=');
    const hasArrows = svg.includes('marker-');
    const hasShadows = svg.includes('filter') || svg.includes('drop-shadow');
    const hasGradients = svg.includes('gradient');
    const hasStyledText = svg.includes('font-weight') || svg.includes('font-style');
    
    console.log(`   🔍 Quality features detected:`);
    console.log(`      • Rounded corners: ${hasRoundedCorners ? '✅' : '❌'}`);
    console.log(`      • Arrows on connectors: ${hasArrows ? '✅' : '❌'}`);
    console.log(`      • Drop shadows: ${hasShadows ? '✅' : '❌'}`);
    console.log(`      • Gradients: ${hasGradients ? '✅' : '❌'}`);
    console.log(`      • Text styling: ${hasStyledText ? '✅' : '❌'}`);
}

// Kör testet
testRendering().catch(console.error);