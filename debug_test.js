#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');

// Set up the browser environment as instructed
const dom = new JSDOM('');
const win = dom.window;

// Load JSZip from the bundled file
const JSZip = new Function('module', 'exports', 'window', 'document', 
  'return ' + fs.readFileSync('./jszip.min.js', 'utf8')
    .replace('!function(e){if("object"==typeof exports&&"undefined"!=typeof module)module.exports=e();', 'function(){return ')
    .replace(/}\(\)\);?$/, '();}')
)({}, {}, win, win.document);

win.JSZip = JSZip;

// Load VisioConverter
new Function('window','JSZip','DOMParser','document',
  fs.readFileSync('libvisio.js','utf8')
)(win, JSZip, win.DOMParser, win.document);

async function analyzeTestFile(filename) {
  console.log(`\n🔍 Analyzing ${filename}...`);
  
  const vsdxPath = path.join(__dirname, 'test-fixtures', filename);
  const refPath = path.join(__dirname, 'test-fixtures', 'reference', filename.replace('.vsdx', '_p0.svg'));
  
  try {
    // Load the file
    const buffer = fs.readFileSync(vsdxPath);
    const conv = new win.VisioConverter();
    await conv.loadFromArrayBuffer(buffer);
    
    console.log(`Masters:`, Object.keys(conv.masters || {}));
    console.log(`Pages:`, conv.pages.length);
    
    if (conv.pages.length > 0) {
      const page = conv.pages[0];
      console.log(`\n📋 Page 0 has ${page.shapes.length} shapes:`);
      
      page.shapes.forEach((shape, i) => {
        console.log(`Shape ${i+1}:`);
        console.log(`  ID: ${shape.id}`);
        console.log(`  Name: ${shape.name_u || 'unnamed'}`);
        console.log(`  Type: ${shape.type || 'unknown'}`);
        console.log(`  Master: ${shape.master || 'none'}`);
        console.log(`  Cells: [${Object.keys(shape.cells).join(', ')}]`);
        console.log(`  Sections: [${Object.keys(shape.sections || {}).join(', ')}]`);
        
        // Check key positioning cells
        const pinX = shape.cells['PinX'];
        const pinY = shape.cells['PinY'];
        const width = shape.cells['Width'];
        const height = shape.cells['Height'];
        console.log(`  Position: PinX=${pinX?.value}, PinY=${pinY?.value}, W=${width?.value}, H=${height?.value}`);
        
        if (shape.sections && shape.sections['Geometry']) {
          const geom = shape.sections['Geometry'];
          console.log(`  Geometry rows: ${Object.keys(geom).length}`);
          Object.keys(geom).forEach(rowKey => {
            const row = geom[rowKey];
            console.log(`    ${rowKey}: ${row.type || 'unknown type'}`);
          });
        }
        console.log('');
      });
      
      // Generate SVG and count elements
      const svg = conv.renderPage(0);
      const pathCount = (svg.match(/<path/g) || []).length;
      const lineCount = (svg.match(/<line/g) || []).length;
      const rectCount = Math.max(0, (svg.match(/<rect/g) || []).length - 1);
      const textCount = (svg.match(/<text/g) || []).length;
      const jsTotal = pathCount + lineCount + rectCount;
      
      console.log(`🎨 JS Output: ${jsTotal} shapes (${pathCount} paths, ${lineCount} lines, ${rectCount} rects), ${textCount} texts`);
      
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
          console.log(`⚠️ Missing ${pythonTotal - jsTotal} shapes - debugging needed!`);
        }
      }
      
      // Save JS output for comparison
      const outputPath = filename.replace('.vsdx', '_js.svg');
      fs.writeFileSync(outputPath, svg);
      console.log(`💾 JS SVG saved as: ${outputPath}`);
      
    } else {
      console.log('❌ No pages found!');
    }
    
  } catch (error) {
    console.log(`❌ Error: ${error.message}`);
    console.log(error.stack);
  }
}

async function main() {
  // Analyze the 4 problem files
  const problemFiles = [
    'test3_house.vsdx',
    'test5_master.vsdx', 
    'bpmn-sample.vsdx',
    'media.vsdx'
  ];
  
  for (const file of problemFiles) {
    await analyzeTestFile(file);
    console.log('\n' + '='.repeat(80));
  }
}

main().catch(console.error);