#!/usr/bin/env node

const fs = require('fs');
const path = require('path');

// Direct Node.js test without JSDOM
async function testVisio() {
  console.log('🧪 Testing Visio converter directly in Node.js...\n');
  
  // Load and configure JSZip for Node.js
  const JSZip = require('jszip');
  
  // Set up minimal globals needed by libvisio
  global.DOMParser = require('@xmldom/xmldom').DOMParser;
  global.JSZip = JSZip;
  
  // Load the converter
  const converterCode = fs.readFileSync(path.join(__dirname, 'libvisio.js'), 'utf8');
  
  // Execute in Node context - need to extract the class properly
  let VisioConverter;
  
  // Create a minimal window-like object
  const mockWindow = {
    DOMParser: global.DOMParser,
    JSZip: global.JSZip
  };
  
  // Execute the libvisio code with our mock window
  eval(`
    const window = mockWindow;
    ${converterCode}
    VisioConverter = window.VisioConverter;
  `);
  
  if (!VisioConverter) {
    console.error('Failed to load VisioConverter');
    return;
  }
  
  await analyzeTestFile('test3_house.vsdx', VisioConverter);
}

async function analyzeTestFile(filename, VisioConverter) {
  console.log(`🔍 Analyzing ${filename}...`);
  
  const vsdxPath = path.join(__dirname, 'test-fixtures', filename);
  const refPath = path.join(__dirname, 'test-fixtures', 'reference', filename.replace('.vsdx', '_p0.svg'));
  
  try {
    // Load the file as buffer
    const buffer = fs.readFileSync(vsdxPath);
    console.log(`Loaded file: ${buffer.length} bytes`);
    
    const conv = new VisioConverter();
    await conv.loadFromArrayBuffer(buffer);
    
    console.log(`Masters:`, Object.keys(conv.masters || {}));
    console.log(`Pages:`, conv.pages.length);
    
    if (conv.pages.length > 0) {
      const page = conv.pages[0];
      console.log(`\n📋 Page 0 has ${page.shapes.length} shapes:`);
      
      page.shapes.forEach((shape, i) => {
        console.log(`\nShape ${i+1}:`);
        console.log(`  ID: ${shape.id}`);
        console.log(`  Name: ${shape.name_u || 'unnamed'}`);
        console.log(`  Type: ${shape.type || 'unknown'}`);
        console.log(`  Master: ${shape.master || 'none'}`);
        
        const cellKeys = Object.keys(shape.cells || {});
        console.log(`  Cells (${cellKeys.length}): [${cellKeys.slice(0, 10).join(', ')}${cellKeys.length > 10 ? '...' : ''}]`);
        
        const sectionKeys = Object.keys(shape.sections || {});
        console.log(`  Sections (${sectionKeys.length}): [${sectionKeys.join(', ')}]`);
        
        // Check key positioning cells
        const cells = shape.cells || {};
        const pinX = cells['PinX'];
        const pinY = cells['PinY'];
        const width = cells['Width'];
        const height = cells['Height'];
        console.log(`  Position: PinX=${pinX?.value}, PinY=${pinY?.value}, W=${width?.value}, H=${height?.value}`);
        
        // Show geometry details
        if (shape.sections && shape.sections['Geometry']) {
          const geom = shape.sections['Geometry'];
          console.log(`  Geometry rows: ${Object.keys(geom).length}`);
          Object.keys(geom).slice(0, 5).forEach(rowKey => {
            const row = geom[rowKey];
            console.log(`    ${rowKey}: ${row.type || 'unknown type'}`);
          });
          if (Object.keys(geom).length > 5) {
            console.log(`    ... and ${Object.keys(geom).length - 5} more rows`);
          }
        }
      });
      
      // Generate SVG and analyze
      console.log(`\n🎨 Generating SVG...`);
      const svg = conv.renderPage(0);
      
      // Save for inspection
      const outputPath = path.join(__dirname, filename.replace('.vsdx', '_node_debug_js.svg'));
      fs.writeFileSync(outputPath, svg);
      console.log(`💾 JS SVG saved as: ${path.basename(outputPath)}`);
      
      // Count elements
      const pathCount = (svg.match(/<path/g) || []).length;
      const lineCount = (svg.match(/<line/g) || []).length;
      const rectCount = Math.max(0, (svg.match(/<rect/g) || []).length - 1); // Exclude background rect
      const textCount = (svg.match(/<text/g) || []).length;
      const jsTotal = pathCount + lineCount + rectCount;
      
      console.log(`\n📊 JS Output: ${jsTotal} shapes (${pathCount} paths, ${lineCount} lines, ${rectCount} rects), ${textCount} texts`);
      
      // Compare with reference
      if (fs.existsSync(refPath)) {
        const refSvg = fs.readFileSync(refPath, 'utf8');
        
        const refPathCount = (refSvg.match(/<path/g) || []).length;
        const refLineCount = (refSvg.match(/<line/g) || []).length;
        const refRectCount = Math.max(0, (refSvg.match(/<rect/g) || []).length - 1);
        const refTextCount = (refSvg.match(/<text/g) || []).length;
        const pythonTotal = refPathCount + refLineCount + refRectCount;
        
        console.log(`🐍 Python Ref: ${pythonTotal} shapes (${refPathCount} paths, ${refLineCount} lines, ${refRectCount} rects), ${refTextCount} texts`);
        
        const accuracy = pythonTotal > 0 ? (jsTotal / pythonTotal * 100) : 0;
        console.log(`📈 Accuracy: ${accuracy.toFixed(1)}%`);
        
        if (accuracy < 90) {
          console.log(`\n⚠️ Missing ${pythonTotal - jsTotal} shapes!`);
          console.log(`Need to debug why ${Math.round((100-accuracy)/100*pythonTotal)} elements are not being rendered.`);
        } else {
          console.log(`\n✅ Good accuracy!`);
        }
      }
      
    } else {
      console.log('❌ No pages found!');
    }
    
  } catch (error) {
    console.log(`❌ Error: ${error.message}`);
    console.log(error.stack);
  }
}

// Install missing dependency and run
async function setup() {
  try {
    require('@xmldom/xmldom');
  } catch (e) {
    console.log('Installing @xmldom/xmldom...');
    require('child_process').execSync('npm install @xmldom/xmldom', { stdio: 'inherit' });
  }
  
  await testVisio();
}

setup().catch(console.error);