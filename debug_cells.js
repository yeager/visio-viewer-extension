#!/usr/bin/env node

const fs = require('fs');
const path = require('path');

async function debugCells() {
  console.log('🐛 Debugging cell parsing and master inheritance...\n');
  
  // Set up globals
  global.DOMParser = require('@xmldom/xmldom').DOMParser;
  global.JSZip = require('jszip');
  
  const VisioConverter = require('./libvisio.js');
  
  const vsdxPath = path.join(__dirname, 'test-fixtures', 'test3_house.vsdx');
  const buffer = fs.readFileSync(vsdxPath);
  
  const conv = new VisioConverter();
  await conv.loadFromArrayBuffer(buffer);
  
  console.log('📁 Masters found:', Object.keys(conv.masters));
  console.log('');
  
  // Debug master shapes
  for (const [masterId, masterShapes] of Object.entries(conv.masters)) {
    console.log(`🎭 Master ${masterId}:`);
    for (const [shapeId, masterShape] of Object.entries(masterShapes)) {
      console.log(`  Shape ${shapeId}:`);
      console.log(`    Cells: ${Object.keys(masterShape.cells).length}`);
      console.log(`    Geometry sections: ${masterShape.geometry.length}`);
      
      // Show key cells
      const keyCells = ['PinX', 'PinY', 'Width', 'Height'];
      keyCells.forEach(cellName => {
        const cell = masterShape.cells[cellName];
        if (cell) {
          console.log(`    ${cellName}: V="${cell.V}", F="${cell.F}"`);
        }
      });
      
      if (masterShape.geometry.length > 0) {
        console.log(`    Geometry:`, masterShape.geometry.map(g => g.rows?.length || 0).join(','), 'rows');
      }
    }
    console.log('');
  }
  
  // Debug page shapes BEFORE master merging
  console.log('📄 Page shapes (before merging):');
  const page = conv.pages[0];
  for (const shape of page.shapes) {
    console.log(`\n🔷 Shape ${shape.id} (${shape.name || 'unnamed'}):`);
    console.log(`  Type: ${shape.type}, Master: ${shape.master}, MasterShape: ${shape.master_shape}`);
    console.log(`  Cells: ${Object.keys(shape.cells).length}`);
    console.log(`  Geometry sections: ${shape.geometry.length}`);
    
    // Show key cells
    const keyCells = ['PinX', 'PinY', 'Width', 'Height'];
    keyCells.forEach(cellName => {
      const cell = shape.cells[cellName];
      if (cell) {
        console.log(`    ${cellName}: V="${cell.V}", F="${cell.F}"`);
      }
    });
    
    // Test master merging manually
    if (shape.master) {
      console.log(`\n  🔗 Testing master merge for master ${shape.master}...`);
      const mergedShape = conv.mergeShapeWithMaster(shape, conv.masters);
      console.log(`  After merge - Cells: ${Object.keys(mergedShape.cells).length}`);
      console.log(`  After merge - Geometry: ${mergedShape.geometry.length}`);
      
      keyCells.forEach(cellName => {
        const cell = mergedShape.cells[cellName];
        if (cell && cell.V) {
          console.log(`    Merged ${cellName}: ${cell.V}`);
        }
      });
      
      // Check cell value parsing
      const testCells = ['PinX', 'PinY', 'Width', 'Height'];
      for (const cellName of testCells) {
        const cellVal = conv.getCellVal(mergedShape, cellName);
        console.log(`    getCellVal(${cellName}): ${cellVal}`);
      }
    }
  }
}

debugCells().catch(console.error);