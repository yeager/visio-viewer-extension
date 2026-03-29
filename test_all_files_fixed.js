#!/usr/bin/env node

const fs = require('fs');
const path = require('path');

// Direct Node.js test
async function testAllFiles() {
  console.log('🧪 Testing all 16 test files after fix...\n');
  
  global.DOMParser = require('@xmldom/xmldom').DOMParser;
  global.JSZip = require('jszip');
  
  const VisioConverter = require('./libvisio.js');
  
  const testDir = path.join(__dirname, 'test-fixtures');
  const refDir = path.join(__dirname, 'test-fixtures', 'reference');
  
  // Get all VSDX files
  const vsdxFiles = fs.readdirSync(testDir).filter(f => f.endsWith('.vsdx')).sort();
  
  console.log(`Found ${vsdxFiles.length} test files:\n`);
  
  const results = [];
  
  for (const vsdxFile of vsdxFiles) {
    const baseName = vsdxFile.replace('.vsdx', '');
    const vsdxPath = path.join(testDir, vsdxFile);
    const refSvgPath = path.join(refDir, `${baseName}_p0.svg`);
    
    console.log(`📄 Testing: ${vsdxFile}`);
    
    let jsElements = 0;
    let pythonElements = 0;
    let jsTexts = 0;
    let pythonTexts = 0;
    let success = false;
    let errorMsg = '';
    
    try {
      // Test JS conversion
      const buffer = fs.readFileSync(vsdxPath);
      const conv = new VisioConverter();
      await conv.loadFromArrayBuffer(buffer);
      
      if (conv.pages.length > 0) {
        const svg = conv.renderPage(0);
        
        // Save output for inspection
        const outputPath = path.join(__dirname, `${baseName}_final_js.svg`);
        fs.writeFileSync(outputPath, svg);
        
        // Count JS elements
        const pathCount = (svg.match(/<path/g) || []).length;
        const lineCount = (svg.match(/<line/g) || []).length;
        const rectCount = Math.max(0, (svg.match(/<rect/g) || []).length - 1); // Subtract background
        const textCount = (svg.match(/<text/g) || []).length;
        
        jsElements = pathCount + lineCount + rectCount;
        jsTexts = textCount;
        
        // Read Python reference
        if (fs.existsSync(refSvgPath)) {
          const refSvg = fs.readFileSync(refSvgPath, 'utf8');
          
          const refPathCount = (refSvg.match(/<path/g) || []).length;
          const refLineCount = (refSvg.match(/<line/g) || []).length;
          const refRectCount = Math.max(0, (refSvg.match(/<rect/g) || []).length - 1);
          const refTextCount = (refSvg.match(/<text/g) || []).length;
          
          pythonElements = refPathCount + refLineCount + refRectCount;
          pythonTexts = refTextCount;
        }
        
        success = true;
      }
    } catch (error) {
      errorMsg = error.message;
      console.log(`  ❌ Error: ${error.message}`);
    }
    
    const accuracy = pythonElements > 0 ? (jsElements / pythonElements * 100) : 0;
    const status = accuracy >= 90 ? '✅' : (accuracy >= 50 ? '⚠️' : '❌');
    
    if (success) {
      console.log(`  JS: ${jsElements} shapes, ${jsTexts} texts`);
      console.log(`  Python: ${pythonElements} shapes, ${pythonTexts} texts`);
      console.log(`  Accuracy: ${accuracy.toFixed(1)}% ${status}`);
    }
    console.log('');
    
    results.push({
      file: vsdxFile,
      jsElements,
      pythonElements,
      jsTexts,
      pythonTexts,
      accuracy: accuracy.toFixed(1),
      success,
      errorMsg
    });
  }
  
  // Summary
  console.log('📊 FINAL SUMMARY:\n');
  
  let passCount = 0;
  let totalAccuracy = 0;
  let validTests = 0;
  
  results.forEach(r => {
    if (!r.success) {
      const status = '❌ ERROR';
      console.log(`${r.file.padEnd(25)} | ${status.padEnd(12)} | ${r.errorMsg}`);
    } else {
      const status = r.accuracy >= 90 ? '✅ PASS' : 
                    r.accuracy >= 50 ? '⚠️ PARTIAL' : '❌ FAIL';
      console.log(`${r.file.padEnd(25)} | ${r.jsElements.toString().padStart(3)}/${r.pythonElements.toString().padStart(3)} shapes | ${r.accuracy.padStart(5)}% | ${status}`);
      
      if (parseFloat(r.accuracy) >= 90) passCount++;
      totalAccuracy += parseFloat(r.accuracy);
      validTests++;
    }
  });
  
  const avgAccuracy = validTests > 0 ? totalAccuracy / validTests : 0;
  const passRate = validTests > 0 ? passCount / validTests * 100 : 0;
  
  console.log('\n' + '='.repeat(80));
  console.log(`🎯 FINAL RESULTS:`);
  console.log(`   Files tested: ${validTests}`);
  console.log(`   Files passing (≥90%): ${passCount}`);
  console.log(`   Pass rate: ${passRate.toFixed(1)}%`);
  console.log(`   Average accuracy: ${avgAccuracy.toFixed(1)}%`);
  
  if (passRate >= 80) {
    console.log('\n🎉 EXCELLENT! Ready for production.');
  } else if (passRate >= 60) {
    console.log('\n✅ GOOD! Minor issues to address.');
  } else {
    console.log('\n📈 PROGRESS! More work needed.');
  }
}

testAllFiles().catch(console.error);