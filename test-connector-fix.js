const fs = require('fs');
const { JSDOM } = require('jsdom');
const JSZip = require('./jszip.min.js');

const dom = new JSDOM('');
const win = dom.window;
win.JSZip = JSZip;

// Load libvisio.js in window context
new Function('window','JSZip','DOMParser','document',fs.readFileSync('libvisio.js','utf8'))(win,JSZip,win.DOMParser,win.document);

async function test(name) {
    console.log(`Testing ${name}...`);
    
    const data = fs.readFileSync('/tmp/visio-tests/' + name + '.vsdx');
    const conv = new win.VisioConverter();
    await conv.loadFromArrayBuffer(data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength));
    const svg = conv.renderPage(0);
    
    // Load reference
    const ref = fs.readFileSync('/tmp/visio-ref/' + name + '_p0.svg', 'utf8');
    
    const tags = ['rect','path','text','line','ellipse','image','polyline'];
    let jsTotal = 0, pyTotal = 0;
    
    tags.forEach(tag => {
        const js = (svg.match(new RegExp('<' + tag + '[\\s/]', 'g')) || []).length;
        const py = (ref.match(new RegExp('<' + tag + '[\\s/]', 'g')) || []).length;
        jsTotal += js; 
        pyTotal += py;
        console.log(`  ${tag}: JS=${js} PY=${py}`);
    });
    
    const pct = Math.round(jsTotal/pyTotal*100);
    console.log(`${name}: JS=${jsTotal} PY=${pyTotal} (${pct}%)`);
    
    // Save output for debugging
    fs.writeFileSync(`/tmp/${name}_js_output.svg`, svg);
    console.log(`Output saved to /tmp/${name}_js_output.svg`);
    
    return pct;
}

async function main() {
    console.log('Testing connector text fixes...\n');
    
    const results = [];
    
    for (const name of ['bpmn-sample','media']) {
        const pct = await test(name);
        results.push({ name, pct });
        console.log('');
    }
    
    console.log('FINAL RESULTS:');
    results.forEach(r => {
        const status = r.pct >= 90 ? '✅ PASS' : '❌ FAIL';
        console.log(`${r.name}: ${r.pct}% ${status}`);
    });
    
    const allPassed = results.every(r => r.pct >= 90);
    console.log(`\nOverall: ${allPassed ? '✅ ALL PASSED' : '❌ SOME FAILED'}`);
}

main().catch(console.error);