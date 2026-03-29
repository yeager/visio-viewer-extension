#!/usr/bin/env node

const fs = require('fs');
const path = require('path');

async function testFile(filename) {
  global.DOMParser = require('@xmldom/xmldom').DOMParser;
  global.JSZip = require('jszip');
  
  const VisioConverter = require('./libvisio.js');
  
  const vsdxPath = path.join(__dirname, 'test-fixtures', filename);
  const refPath = path.join(__dirname, 'test-fixtures', 'reference', filename.replace('.vsdx', '_p0.svg'));
  
  console.log(`🔍 Testing ${filename}...`);
  
  try {
    const buffer = fs.readFileSync(vsdxPath);
    const conv = new VisioConverter();
    await conv.loadFromArrayBuffer(buffer);
    
    console.log(`  Masters: [${Object.keys(conv.masters).join(', ')}]`);
    console.log(`  Pages: ${conv.pages.length}`);
    
    if (conv.pages.length > 0) {
      const page = conv.pages[0];
      console.log(`  Shapes: ${page.shapes.length}`);
      
      // Generate SVG
      const svg = conv.renderPage(0);
      
      // Save output
      const outputPath = path.join(__dirname, filename.replace('.vsdx', '_fixed_js.svg'));
      fs.writeFileSync(outputPath, svg);
      
      // Count elements
      const pathCount = (svg.match(/<path/g) || []).length;
      const lineCount = (svg.match(/<line/g) || []).length;
      const rectCount = Math.max(0, (svg.match(/<rect/g) || []).length - 1);
      const textCount = (svg.match(/<text/g) || []).length;
      const jsTotal = pathCount + lineCount + rectCount;
      
      console.log(`  JS Output: ${jsTotal} shapes (${pathCount} paths, ${lineCount} lines, ${rectCount} rects), ${textCount} texts`);
      
      // Compare with reference
      if (fs.existsSync(refPath)) {
        const refSvg = fs.readFileSync(refPath, 'utf8');
        
        const refPathCount = (refSvg.match(/<path/g) || []).length;
        const refLineCount = (refSvg.match(/<line/g) || []).length;
        const refRectCount = Math.max(0, (refSvg.match(/<rect/g) || []).length - 1);
        const refTextCount = (refSvg.match(/<text/g) || []).length;
        const pythonTotal = refPathCount + refLineCount + refRectCount;
        
        console.log(`  Python Ref: ${pythonTotal} shapes (${refPathCount} paths, ${refLineCount} lines, ${refRectCount} rects), ${refTextCount} texts`);
        
        const accuracy = pythonTotal > 0 ? (jsTotal / pythonTotal * 100) : 0;
        const status = accuracy >= 90 ? '✅' : (accuracy >= 50 ? '⚠️' : '❌');
        console.log(`  Accuracy: ${accuracy.toFixed(1)}% ${status}`);
        
        return { file: filename, jsTotal, pythonTotal, accuracy: accuracy.toFixed(1) };
      }
    }
    
  } catch (error) {
    console.log(`  ❌ Error: ${error.message}`);
    return { file: filename, error: error.message };
  }
}

async function testProblemFiles() {
  console.log('🧪 Testing the 4 problem files after fix...\n');
  
  const problemFiles = [
    'test3_house.vsdx',      // Was 32% - should be fixed now
    'test5_master.vsdx',     // Was 54%  
    'bpmn-sample.vsdx',      // Was 68%
    'media.vsdx'             // Was 63%
  ];
  
  const results = [];
  
  for (const file of problemFiles) {
    const result = await testFile(file);
    results.push(result);
    console.log('');
  }
  
  // Summary
  console.log('📊 SUMMARY AFTER FIX:\n');
  
  let totalAccuracy = 0;
  let validTests = 0;
  let passCount = 0;
  
  results.forEach(r => {
    if (r.error) {
      console.log(`${r.file.padEnd(25)} | ERROR: ${r.error}`);
    } else {
      const status = r.accuracy >= 90 ? '✅ PASS' : 
                    r.accuracy >= 50 ? '⚠️ PARTIAL' : '❌ FAIL';
      console.log(`${r.file.padEnd(25)} | ${r.jsTotal.toString().padStart(3)}/${r.pythonTotal.toString().padStart(3)} shapes | ${r.accuracy.padStart(5)}% | ${status}`);
      
      if (parseFloat(r.accuracy) >= 90) passCount++;
      totalAccuracy += parseFloat(r.accuracy);
      validTests++;
    }
  });
  
  if (validTests > 0) {
    const avgAccuracy = totalAccuracy / validTests;
    const passRate = passCount / validTests * 100;
    
    console.log('\n' + '='.repeat(80));
    console.log(`🎯 RESULTS:`);
    console.log(`   Files passing (≥90%): ${passCount}/${validTests}`);
    console.log(`   Pass rate: ${passRate.toFixed(1)}%`);
    console.log(`   Average accuracy: ${avgAccuracy.toFixed(1)}%`);
    
    if (passRate >= 75) {
      console.log('\n🎉 EXCELLENT! Major improvement achieved.');
    } else if (avgAccuracy >= 80) {
      console.log('\n✅ GOOD! Significant progress made.');
    } else {
      console.log('\n📈 PROGRESS! More work needed.');
    }
  }
}

testProblemFiles().catch(console.error);