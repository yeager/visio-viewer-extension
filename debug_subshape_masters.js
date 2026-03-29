#!/usr/bin/env node

const fs = require('fs');
const path = require('path');

async function debugSubShapeMasters() {
  console.log('🔍 Debugging sub-shape master references...\n');
  
  global.DOMParser = require('@xmldom/xmldom').DOMParser;
  global.JSZip = require('jszip');
  
  const VisioConverter = require('./libvisio.js');
  
  const vsdxPath = path.join(__dirname, 'test-fixtures', 'test3_house.vsdx');
  const buffer = fs.readFileSync(vsdxPath);
  
  const conv = new VisioConverter();
  await conv.loadFromArrayBuffer(buffer);
  
  // Check page group sub-shapes in detail
  const houseShape = conv.pages[0].shapes.find(s => s.id === '7');
  console.log(`🏠 House group (ID=${houseShape.id}) analysis:`);
  console.log(`Master: ${houseShape.master}, MasterShape: ${houseShape.master_shape}`);
  console.log(`Sub-shapes: ${houseShape.sub_shapes.length}`);
  
  console.log(`\n📝 Page sub-shapes before merge:`);
  houseShape.sub_shapes.forEach((sub, i) => {
    console.log(`  Sub ${i}: ID=${sub.id}, Master="${sub.master}", MasterShape="${sub.master_shape}"`);
    console.log(`           Type=${sub.type}, Geometry=${sub.geometry?.length || 0}`);
    console.log(`           Cells: ${Object.keys(sub.cells || {}).length}`);
    
    // Show geometry details if any
    if (sub.geometry?.length > 0) {
      sub.geometry.forEach((geo, gi) => {
        console.log(`           Geo ${gi}: ${geo.rows?.length || 0} rows`);
      });
    }
  });
  
  // Check master sub-shapes
  console.log(`\n🎭 Master sub-shapes structure:`);
  const masterData = conv.masters[houseShape.master];
  const masterGroupShape = masterData['5']; // The group master
  
  console.log(`Master group shape ${masterGroupShape.id}: ${masterGroupShape.sub_shapes.length} sub-shapes`);
  masterGroupShape.sub_shapes.forEach((masterSub, i) => {
    console.log(`  Master Sub ${i}: ID=${masterSub.id}, Geometry=${masterSub.geometry?.length || 0}`);
    if (masterSub.geometry?.length > 0) {
      masterSub.geometry.forEach((geo, gi) => {
        console.log(`                   Geo ${gi}: ${geo.rows?.length || 0} rows`);
      });
    }
  });
  
  // Test sub-shape merging manually
  console.log(`\n🔧 Testing sub-shape merging:`);
  houseShape.sub_shapes.forEach((pageSub, i) => {
    console.log(`\nSub-shape ${i} (Page ID=${pageSub.id}):`);
    
    // Try to match with master sub-shape by position/index 
    const masterSub = masterGroupShape.sub_shapes[i];
    if (masterSub) {
      console.log(`  Matching with master sub ID=${masterSub.id}`);
      console.log(`  Master has ${masterSub.geometry?.length || 0} geometry sections`);
      
      // Test merging this specific sub-shape
      const mergedSub = conv.mergeShapeWithMaster(JSON.parse(JSON.stringify(pageSub)), conv.masters);
      console.log(`  After merge: ${mergedSub.geometry?.length || 0} geometry sections`);
      
      if (mergedSub.geometry?.length > 0) {
        mergedSub.geometry.forEach((geo, gi) => {
          console.log(`    Merged Geo ${gi}: ${geo.rows?.length || 0} rows`);
        });
      }
    }
  });
  
  // Test full group rendering
  console.log(`\n🎨 Testing group rendering:`);
  const groupElements = conv.renderGroupShape(houseShape, 11.0, new Set(), [], false, [], {});
  console.log(`Group rendered ${groupElements.length} elements`);
  groupElements.slice(0, 10).forEach((elem, i) => {
    console.log(`  ${i}: ${elem.substring(0, 100)}${elem.length > 100 ? '...' : ''}`);
  });
}

debugSubShapeMasters().catch(console.error);