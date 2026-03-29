#!/usr/bin/env node

const fs = require('fs');
const path = require('path');

async function debugSubShapes() {
  console.log('🔍 Debugging sub-shapes for groups...\n');
  
  global.DOMParser = require('@xmldom/xmldom').DOMParser;
  global.JSZip = require('jszip');
  
  const VisioConverter = require('./libvisio.js');
  
  const vsdxPath = path.join(__dirname, 'test-fixtures', 'test3_house.vsdx');
  const buffer = fs.readFileSync(vsdxPath);
  
  const conv = new VisioConverter();
  await conv.loadFromArrayBuffer(buffer);
  
  // Check sub-shapes in page shapes
  console.log('📄 Page shapes sub-shapes:');
  const page = conv.pages[0];
  for (const shape of page.shapes) {
    console.log(`\nShape ${shape.id} (${shape.name || 'unnamed'}):`)
    console.log(`  Type: ${shape.type}, Master: ${shape.master}`);
    console.log(`  Sub-shapes: ${shape.sub_shapes?.length || 0}`);
    
    if (shape.sub_shapes?.length > 0) {
      shape.sub_shapes.forEach((sub, i) => {
        console.log(`    Sub ${i}: ID=${sub.id}, Name=${sub.name}, Type=${sub.type}`);
      });
    }
  }
  
  console.log('\n🎭 Master shapes sub-shapes:');
  for (const [masterId, masterShapes] of Object.entries(conv.masters)) {
    console.log(`\nMaster ${masterId}:`);
    for (const [shapeId, masterShape] of Object.entries(masterShapes)) {
      console.log(`  Shape ${shapeId}:`);
      console.log(`    Type: ${masterShape.type}, Sub-shapes: ${masterShape.sub_shapes?.length || 0}`);
      console.log(`    Geometry: ${masterShape.geometry?.length || 0} sections`);
      
      if (masterShape.sub_shapes?.length > 0) {
        console.log(`    Sub-shapes:`);
        masterShape.sub_shapes.forEach((sub, i) => {
          console.log(`      Sub ${i}: ID=${sub.id}, Name=${sub.name}, Geom=${sub.geometry?.length || 0}`);
        });
      }
    }
  }
  
  // Test if sub-shapes should be inherited during master merge
  const testShape = page.shapes.find(s => s.id === '7'); // House group
  console.log(`\n🔬 Testing master merge for House group...`);
  
  // Create a copy to test merging
  const originalSubShapes = testShape.sub_shapes.length;
  const mergedShape = conv.mergeShapeWithMaster(JSON.parse(JSON.stringify(testShape)), conv.masters);
  const mergedSubShapes = mergedShape.sub_shapes.length;
  
  console.log(`Original sub-shapes: ${originalSubShapes}`);
  console.log(`After merge sub-shapes: ${mergedSubShapes}`);
  
  // If group doesn't have sub-shapes but master does, the merging should transfer them
  const masterId = testShape.master;
  if (conv.masters[masterId]) {
    console.log(`\n🎯 Master ${masterId} analysis:`);
    const masterShapes = Object.values(conv.masters[masterId]);
    
    const masterWithSubShapes = masterShapes.find(ms => ms.sub_shapes?.length > 0);
    if (masterWithSubShapes) {
      console.log(`Found master shape with ${masterWithSubShapes.sub_shapes.length} sub-shapes`);
      masterWithSubShapes.sub_shapes.forEach((sub, i) => {
        console.log(`  Sub ${i}: ID=${sub.id}, Geometry=${sub.geometry?.length || 0} sections`);
      });
    } else {
      console.log(`No master shapes have sub-shapes - they should have individual geometry instead`);
      console.log(`Master shapes should be rendered as separate geometry within the group`);
    }
  }
}

debugSubShapes().catch(console.error);