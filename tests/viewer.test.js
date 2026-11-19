const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { JSDOM } = require('jsdom');
const root = path.join(__dirname, '..');
const tick = () => new Promise(resolve => setImmediate(resolve));

function viewer(search = '') {
  const dom = new JSDOM(fs.readFileSync(path.join(root, 'viewer.html'), 'utf8'), {
    url: 'https://viewer.example/' + search, runScripts: 'outside-only'
  });
  const { window } = dom;
  window.JSZip = {};
  window.console.error = () => {};
  window.VisioConverter = class {
    async loadFromArrayBuffer(data) { this.name = await data; }
    getPages() { return [{ name: this.name }]; }
    renderPage() { return `<svg><text>${this.name}</text></svg>`; }
  };
  window.eval(fs.readFileSync(path.join(root, 'viewer.js'), 'utf8'));
  const open = file => {
    const input = window.document.getElementById('file-input');
    Object.defineProperty(input, 'files', { value: [file], configurable: true });
    input.dispatchEvent(new window.Event('change'));
  };
  return { dom, window, document: window.document, open };
}

test('error dismissal uses an external listener and a new file clears errors', async () => {
  const v = viewer();
  v.open({ name: 'old.vsd', arrayBuffer: async () => 'wrong' });
  await tick();
  const overlay = v.document.getElementById('error-overlay');
  assert.equal(overlay.classList.contains('hidden'), false);
  v.document.getElementById('dismiss-error').click();
  assert.equal(overlay.classList.contains('hidden'), true);
  assert.equal(v.document.querySelector('[onclick]'), null);
  v.open({ name: 'good.vsdx', arrayBuffer: async () => 'good' });
  await tick();
  assert.equal(v.document.getElementById('file-title').textContent, 'good.vsdx');
  v.dom.window.close();
});

test('a slow earlier conversion cannot replace the latest document', async () => {
  const v = viewer();
  let finish;
  const pending = new Promise(resolve => { finish = resolve; });
  v.open({ name: 'slow.vsdx', arrayBuffer: async () => ({ then: pending.then.bind(pending) }) });
  await tick();
  v.open({ name: 'fast.vsdx', arrayBuffer: async () => 'fast' });
  await tick();
  finish('slow');
  await tick();
  assert.equal(v.document.getElementById('file-title').textContent, 'fast.vsdx');
  assert.match(v.document.getElementById('svg-container').textContent, /fast/);
  v.dom.window.close();
});

test('filenames in local file prompts are displayed as text', () => {
  const name = '<img src=x onerror=alert(1)>.vsdx';
  const v = viewer('?source=file&name=' + encodeURIComponent(name));
  assert.equal(v.document.querySelector('.file-prompt-overlay strong').textContent, name);
  assert.equal(v.document.querySelector('.file-prompt-overlay img'), null);
  v.dom.window.close();
});

test('a tiny viewport never produces a negative zoom', async () => {
  const v = viewer();
  v.open({ name: 'good.vsdx', arrayBuffer: async () => 'good' });
  await tick();
  assert.equal(v.document.getElementById('zoom-level').textContent, '10%');
  v.dom.window.close();
});

test('failed replacement preserves the previous document and catches read errors', async () => {
  const v = viewer();
  v.open({ name: 'good.vsdx', arrayBuffer: async () => 'good' });
  await tick();
  v.open({ name: 'broken.vsdx', arrayBuffer: async () => { throw Error('read failed'); } });
  await tick();
  assert.equal(v.document.getElementById('file-title').textContent, 'good.vsdx');
  assert.equal(v.document.getElementById('error-message').textContent, 'read failed');
  v.dom.window.close();
});


test('a slow converter cannot replace a later successful conversion', async () => {
  const v = viewer();
  let finish;
  const pending = new Promise(resolve => { finish = resolve; });
  v.window.VisioConverter = class {
    async loadFromArrayBuffer(data) {
      this.name = data;
      if (data === 'slow') await pending;
    }
    getPages() { return [{ name: this.name }]; }
    renderPage() { return `<svg><text>${this.name}</text></svg>`; }
  };
  v.open({ name: 'slow.vsdx', arrayBuffer: async () => 'slow' });
  await tick();
  v.open({ name: 'fast.vsdx', arrayBuffer: async () => 'fast' });
  await tick();
  finish();
  await tick();
  assert.equal(v.document.getElementById('file-title').textContent, 'fast.vsdx');
  assert.equal(v.document.getElementById('page-info').textContent, 'fast');
  v.dom.window.close();
});

test('local file prompt accepts drag and drop and closes after opening', async () => {
  const v = viewer('?source=file&name=diagram.vsdx');
  const event = new v.window.Event('drop', { cancelable: true });
  Object.defineProperty(event, 'dataTransfer', { value: {
    files: [{ name: 'diagram.vsdx', arrayBuffer: async () => 'diagram' }]
  }});
  v.document.querySelector('.file-prompt-overlay').dispatchEvent(event);
  await tick();
  assert.equal(v.document.querySelector('.file-prompt-overlay'), null);
  assert.equal(v.document.getElementById('file-title').textContent, 'diagram.vsdx');
  v.dom.window.close();
});

test('real ZIP parsing and SVG rendering integrate with file opening', async () => {
  const v = viewer();
  const JSZip = require('../jszip.min.js');
  v.window.JSZip = JSZip;
  v.window.eval(fs.readFileSync(path.join(root, 'libvisio.js'), 'utf8'));
  const zip = new JSZip();
  zip.file('visio/pages/page1.xml', `<PageContents xmlns="http://schemas.microsoft.com/office/visio/2012/main">
    <Shapes><Shape ID="1" Type="Shape"><Cell N="Width" V="2"/><Cell N="Height" V="1"/>
    <Cell N="PinX" V="2"/><Cell N="PinY" V="2"/><Text>Hello diagram</Text></Shape></Shapes>
    </PageContents>`);
  const data = await zip.generateAsync({ type: 'nodebuffer' });
  v.open({ name: 'real.vsdx', arrayBuffer: async () => data });
  for (let i = 0; i < 50 && v.document.getElementById('file-title').textContent !== 'real.vsdx'; i++) await tick();
  assert.equal(v.document.getElementById('file-title').textContent, 'real.vsdx');
  assert.match(v.document.getElementById('svg-container').textContent, /Hello diagram/);
  v.dom.window.close();
});


test('page order and labels follow catalog relationships, including empty pages', async () => {
  const v = viewer();
  const JSZip = require('../jszip.min.js');
  v.window.JSZip = JSZip;
  v.window.eval(fs.readFileSync(path.join(root, 'libvisio.js'), 'utf8'));
  const zip = new JSZip();
  const ns = 'http://schemas.microsoft.com/office/visio/2012/main';
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  zip.file('visio/pages/pages.xml', `<Pages xmlns="${ns}" xmlns:r="${rel}">
    <Page Name="Tenth file first"><Rel r:id="second"/></Page>
    <Page Name="Empty sheet"><Rel r:id="first"/></Page>
  </Pages>`);
  zip.file('visio/pages/_rels/pages.xml.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="first" Target="page1.xml"/><Relationship Id="second" Target="page10.xml"/>' +
    '</Relationships>');
  zip.file('visio/pages/page1.xml', `<PageContents xmlns="${ns}"/>`);
  zip.file('visio/pages/page10.xml', `<PageContents xmlns="${ns}"><Shapes><Shape ID="1">
    <Cell N="Width" V="2"/><Cell N="Height" V="1"/><Cell N="PinX" V="2"/><Cell N="PinY" V="2"/>
    <Text>Catalog order marker</Text></Shape></Shapes></PageContents>`);
  const data = await zip.generateAsync({ type: 'nodebuffer' });
  v.open({ name: 'ordered.vsdx', arrayBuffer: async () => data });
  for (let i = 0; i < 50 && v.document.getElementById('file-title').textContent !== 'ordered.vsdx'; i++) await tick();
  assert.match(v.document.getElementById('svg-container').textContent, /Catalog order marker/);
  assert.equal(v.document.getElementById('page-info').textContent, 'Tenth file first (1/2)');
  v.document.getElementById('next-page').click();
  assert.equal(v.document.getElementById('page-info').textContent, 'Empty sheet (2/2)');
  v.dom.window.close();
});
