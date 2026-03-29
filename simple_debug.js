#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');

// Create a browser-like environment
const dom = new JSDOM(`
<!DOCTYPE html>
<html>
  <head><title>Test</title></head>
  <body>
    <script src="jszip.min.js"></script>
    <script src="libvisio.js"></script>
  </body>
</html>
`, {
  runScripts: "dangerously",
  resources: "usable",
  url: "file://" + __dirname + "/"
});

// Wait for scripts to load
setTimeout(async () => {
  try {
    const window = dom.window;
    
    console.log('JSZip available:', typeof window.JSZip !== 'undefined');
    console.log('VisioConverter available:', typeof window.VisioConverter !== 'undefined');
    
    if (typeof window.VisioConverter === 'undefined') {
      console.log('VisioConverter not found, trying to load manually...');
      
      // Load JSZip manually
      const jszipCode = fs.readFileSync(path.join(__dirname, 'jszip.min.js'), 'utf8');
      const libvisioCode = fs.readFileSync(path.join(__dirname, 'libvisio.js'), 'utf8');
      
      // Execute JSZip in window context  
      window.eval(jszipCode);
      console.log('JSZip loaded manually:', typeof window.JSZip !== 'undefined');
      
      // Execute libvisio in window context
      window.eval(libvisioCode);
      console.log('VisioConverter loaded manually:', typeof window.VisioConverter !== 'undefined');
    }
    
    if (typeof window.VisioConverter !== 'undefined') {
      await analyzeFile('test3_house.vsdx', window);
    }
    
  } catch (error) {
    console.error('Error:', error.message);
  }
}, 1000);

async function analyzeFile(filename, window) {
  console.log(`\n🔍 Analyzing ${filename}...`);
  
  const vsdxPath = path.join(__dirname, 'test-fixtures', filename);
  const refPath = path.join(__dirname, 'test-fixtures', 'reference', filename.replace('.vsdx', '_p0.svg'));
  
  try {
    // Load the file
    const buffer = fs.readFileSync(vsdxPath);
    const conv = new window.VisioConverter();
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
        console.log(`  Cells: [${Object.keys(shape.cells || {}).join(', ')}]`);
        console.log(`  Sections: [${Object.keys(shape.sections || {}).join(', ')}]`);
        
        // Check key positioning cells
        const pinX = shape.cells?.['PinX'];
        const pinY = shape.cells?.['PinY'];
        const width = shape.cells?.['Width'];
        const height = shape.cells?.['Height'];
        console.log(`  Position: PinX=${pinX?.value}, PinY=${pinY?.value}, W=${width?.value}, H=${height?.value}`);
        
        if (shape.sections && shape.sections['Geometry']) {
          const geom = shape.sections['Geometry'];
          console.log(`  Geometry rows: ${Object.keys(geom).length}`);
          Object.keys(geom).forEach(rowKey => {
            const row = geom[rowKey];
            console.log(`    ${rowKey}: ${row.type || 'unknown type'}`);
          });
        }
      });
      
      // Generate SVG and count elements  
      const svg = conv.renderPage(0);
      
      // Save JS output for inspection
      const outputPath = path.join(__dirname, filename.replace('.vsdx', '_debug_js.svg'));
      fs.writeFileSync(outputPath, svg);
      console.log(`\n💾 JS SVG saved as: ${outputPath}`);
      
      const pathCount = (svg.match(/<path/g) || []).length;
      const lineCount = (svg.match(/<line/g) || []).length;
      const rectCount = Math.max(0, (svg.match(/<rect/g) || []).length - 1);
      const textCount = (svg.match(/<text/g) || []).length;
      const jsTotal = pathCount + lineCount + rectCount;
      
      console.log(`\n🎨 JS Output: ${jsTotal} shapes (${pathCount} paths, ${lineCount} lines, ${rectCount} rects), ${textCount} texts`);
      
      // Compare with Python reference
      if (fs.existsSync(refPath)) {
        const refSvg = fs.readFileSync(refPath, 'utf8');
        const refPathCount = (refSvg.match(/<path/g) || []).length;
        const refLineCount = (refSvg.match(/<line/g) || []).length;
        const refRectCount = Math.max(0, (refSvg.match(/<rect/g) || []).length - 1);
        const refTextCount = (refSvg.match(/<text/g) || []).length;
        const pythonTotal = refPathCount + refLineCount + refRectCount;
        
        console.log(`🐍 Python Ref: ${pythonTotal} shapes (${refPathCount} paths, ${refLineCount} lines, ${refRectCount} rects), ${refTextCount} texts`);
        
        const accuracy = pythonTotal > 0 ? (jsTotal / pythonTotal * 100) : 0;
        console.log(`📊 Accuracy: ${accuracy.toFixed(1)}%`);
        
        if (accuracy < 90) {
          console.log(`⚠️ Missing ${pythonTotal - jsTotal} shapes - analyzing differences...`);
          
          // Show first few lines of each SVG for comparison
          console.log(`\n🔍 First 10 elements in Python reference:`);
          const pyElements = refSvg.match(/<(path|line|rect|text)[^>]*>/g) || [];
          pyElements.slice(0, 10).forEach((elem, i) => {
            console.log(`  ${i+1}: ${elem.substring(0, 80)}${elem.length > 80 ? '...' : ''}`);
          });
          
          console.log(`\n🔍 First 10 elements in JS output:`);
          const jsElements = svg.match(/<(path|line|rect|text)[^>]*>/g) || [];
          jsElements.slice(0, 10).forEach((elem, i) => {
            console.log(`  ${i+1}: ${elem.substring(0, 80)}${elem.length > 80 ? '...' : ''}`);
          });
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