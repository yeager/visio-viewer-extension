#!/usr/bin/env node

const fs = require('fs');
const path = require('path');

async function debugMasterSelection() {
  console.log('🎯 Debugging master selection logic...\n');
  
  global.DOMParser = require('@xmldom/xmldom').DOMParser;
  global.JSZip = require('jszip');
  
  const VisioConverter = require('./libvisio.js');
  
  const vsdxPath = path.join(__dirname, 'test-fixtures', 'test3_house.vsdx');
  const buffer = fs.readFileSync(vsdxPath);
  
  const conv = new VisioConverter();
  await conv.loadFromArrayBuffer(buffer);
  
  // Debug the master selection logic for shape 7
  const testShape = conv.pages[0].shapes.find(s => s.id === '7');
  console.log('🔍 Testing shape 7 (House):');
  console.log(`  Master: "${testShape.master}"`);
  console.log(`  MasterShape: "${testShape.master_shape}"`);
  console.log(`  Geometry length: ${testShape.geometry.length}`);
  console.log('');
  
  const masterId = testShape.master;
  const masterShapeId = testShape.master_shape;
  
  console.log(`🎭 Available masters: [${Object.keys(conv.masters).join(', ')}]`);
  
  if (conv.masters[masterId]) {
    const masterShapes = conv.masters[masterId];
    console.log(`📦 Master ${masterId} contains shapes: [${Object.keys(masterShapes).join(', ')}]`);
    
    let selectedMaster = null;
    
    if (masterShapeId && masterShapes[masterShapeId]) {
      selectedMaster = masterShapes[masterShapeId];
      console.log(`✅ Selected master by MasterShape ID: ${masterShapeId}`);
    } else if (Object.keys(masterShapes).length > 0) {
      selectedMaster = Object.values(masterShapes)[0];
      const firstKey = Object.keys(masterShapes)[0];
      console.log(`⚠️ Selected first available master shape: ${firstKey} (MasterShape "${masterShapeId}" not found)`);
    }
    
    if (selectedMaster) {
      console.log(`\n🔧 Selected master details:`);
      console.log(`  Geometry sections: ${selectedMaster.geometry?.length || 0}`);
      console.log(`  Cells: ${Object.keys(selectedMaster.cells || {}).length}`);
      console.log(`  Width: ${selectedMaster.cells?.Width?.V}`);
      console.log(`  Height: ${selectedMaster.cells?.Height?.V}`);
      
      if (selectedMaster.geometry?.length > 0) {
        console.log(`\n📐 Geometry details:`);
        selectedMaster.geometry.forEach((geo, i) => {
          console.log(`    Section ${i}: ${geo.rows?.length || 0} rows, no_show: ${geo.no_show}`);
          if (geo.rows?.length > 0) {
            console.log(`      Row types: ${geo.rows.map(r => r.type).join(', ')}`);
          }
        });
      }
    }
  } else {
    console.log(`❌ Master ${masterId} not found!`);
  }
  
  // Test the actual merging process step by step
  console.log(`\n🔄 Testing merge process...`);
  const mergedShape = conv.mergeShapeWithMaster(JSON.parse(JSON.stringify(testShape)), conv.masters);
  console.log(`After merge: geometry sections = ${mergedShape.geometry?.length || 0}`);
  
  if (mergedShape.geometry?.length > 0) {
    console.log(`Merged geometry details:`);
    mergedShape.geometry.forEach((geo, i) => {
      console.log(`  Section ${i}: ${geo.rows?.length || 0} rows`);
    });
  }
  
  // Test rendering with the merged shape
  console.log(`\n🎨 Testing rendering...`);
  const svgElements = conv.renderShapeToSvg(JSON.parse(JSON.stringify(testShape)), 11.0, new Set(), [], false, [], {});
  console.log(`Rendered elements: ${svgElements.length}`);
  svgElements.forEach((elem, i) => {
    console.log(`  ${i}: ${elem.substring(0, 80)}...`);
  });
}

debugMasterSelection().catch(console.error);