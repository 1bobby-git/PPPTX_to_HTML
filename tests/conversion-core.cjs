'use strict';
const assert = require('node:assert/strict');
const test = require('node:test');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const {JSDOM} = require('jsdom');
const createDOMPurify = require('dompurify');
const core = require('../src/conversion-core.js');
const root = path.resolve(__dirname, '..');
function fake(infoValues = {}, resultValues = {}) {
  const state = {infoFreed: 0, resultFreed: 0, calls: []};
  const engine = {
    get_presentation_info() { return {widthPx: 1280, heightPx: 720, slideCount: 4, ...infoValues, free() {state.infoFreed++;}}; },
    convert_with_options_metadata(...args) {
      state.calls.push(args);
      return {html: '<div>ok</div>', slideCount: 3, diagnosticsJson: '[]', ...resultValues, free() {state.resultFreed++;}};
    }
  };
  return {engine, state};
}
function errorCode(fn, code) { assert.throws(fn, error => error.code === code); }
function exporter() {
  const dom = new JSDOM('', {runScripts: 'outside-only', url: 'about:blank'});
  let script = fs.readFileSync(path.join(root, 'src/exporter.js'), 'utf8');
  for (const [marker, value] of [['VIEWER_CSS', '.slide{}'], ['VIEWER_SCRIPT', 'void 0;'], ['VIEWER_HASH', 'a'.repeat(43) + '=']]) {
    script = script.replace('/* ' + marker + ' */ null', () => JSON.stringify(value));
  }
  dom.window.DOMPurify = createDOMPurify(dom.window);
  dom.window.eval(script);
  return {dom, prepare: options => dom.window.PptxExport.prepare(options)};
}
function shell(n, text = 'slide') { return '<div class="slide-shell" data-slide="' + n + '"><div class="slide" data-slide="' + n + '">' + text + '</div></div>'; }

test('Node and browser globals expose the same shared contract', () => {
  const context = vm.createContext({Uint32Array});
  vm.runInContext(fs.readFileSync(path.join(root, 'src/conversion-core.js'), 'utf8'), context);
  assert.equal(typeof context.PptxCore.render, 'function');
  const {engine, state} = fake();
  assert.equal(context.PptxCore.render(engine, new Uint8Array()).slideCount, 3);
  assert.deepEqual([state.infoFreed, state.resultFreed], [1, 1]);
});
test('hidden-inclusive empty ranges enumerate presentation indices and explicit ranges stay ordered', () => {
  const {engine, state} = fake({}, {slideCount: 4});
  assert.equal(core.render(engine, new Uint8Array(), {includeHidden: true}).slideCount, 4);
  assert.deepEqual([...state.calls[0][3]], [1, 2, 3, 4]);
  assert.equal(state.calls[0][2], true);
  const visible = fake();
  core.render(visible.engine, new Uint8Array());
  assert.deepEqual([...visible.state.calls[0][3]], []);
  const selected = fake({}, {slideCount: 2});
  core.render(selected.engine, new Uint8Array(), {slideRange: '3,1,1', includeHidden: true});
  assert.deepEqual([...selected.state.calls[0][3]], [1, 3]);
});
test('invalid selection and dimensions free metadata before conversion', () => {
  for (const range of [0, '0', '2-1', '5', '1,x', '-1', '9007199254740992']) {
    const {engine, state} = fake();
    errorCode(() => core.render(engine, new Uint8Array(), {slideRange: range}), 'INVALID_ARGUMENT');
    assert.equal(state.calls.length, 0);
    assert.equal(state.infoFreed, 1);
  }
  for (const values of [{slideCount: 1001}, {slideCount: 0}, {slideCount: 1.5}, {widthPx: Infinity}, {heightPx: 30001}]) {
    const {engine, state} = fake(values);
    errorCode(() => core.render(engine, new Uint8Array()), values.slideCount !== undefined ? 'SLIDE_LIMIT' : 'INVALID_DIMENSIONS');
    assert.equal(state.infoFreed, 1);
    assert.equal(state.calls.length, 0);
  }
});
test('empty and inconsistent engine results fail and release both resources', () => {
  for (const [values, code] of [[{slideCount: 0}, 'EMPTY_SELECTION'], [{slideCount: 5}, 'SLIDE_COUNT_MISMATCH'], [{slideCount: NaN}, 'SLIDE_COUNT_MISMATCH'], [{html: ''}, 'EMPTY_OUTPUT']]) {
    const {engine, state} = fake({}, values);
    errorCode(() => core.render(engine, new Uint8Array()), code);
    assert.deepEqual([state.infoFreed, state.resultFreed], [1, 1]);
  }
  const {engine, state} = fake();
  engine.convert_with_options_metadata = () => {throw new Error('engine failure');};
  assert.throws(() => core.render(engine, new Uint8Array()), /engine failure/);
  assert.equal(state.infoFreed, 1);
});
test('UTF-8 output budget rejects multibyte output smaller than the character limit', () => {
  const {engine, state} = fake({}, {html: '한'.repeat(Math.floor(256 * 1024 * 1024 / 3) + 1)});
  errorCode(() => core.render(engine, new Uint8Array()), 'OUTPUT_LIMIT');
  assert.deepEqual([state.infoFreed, state.resultFreed], [1, 1]);
});
test('diagnostics retain group controls but never disclose raw notes or unsupported text', () => {
  const notices = core.summarizeDiagnostics(JSON.stringify([
    null, {code: 'NOTE_SECRET', message: 'PRIVATE_NOTE'},
    {code: 'OOXML_ELEMENT_UNSUPPORTED', message: 'RAW_SOURCE', location: {part_name: 'ppt/slides/slide1.xml', qualified_element_name: 'a:chOff', slide_index: 0}},
    {code: 'OOXML_ELEMENT_UNSUPPORTED', location: {part_name: 'ppt/slideMasters/slideMaster1.xml', qualified_element_name: 'a:masterClrMapping', slide_index: 1}},
    {code: 'FONT_MISSING', location: {part_name: 'ppt/slides/slide1.xml', slide_index: 0}}
  ]));
  assert.equal(notices.length, 2);
  assert.doesNotMatch(notices[0].message, /슬라이드 \d/);
  assert.ok(!JSON.stringify(notices).includes('PRIVATE_NOTE'));
  assert.ok(!JSON.stringify(notices).includes('RAW_SOURCE'));
  assert.match(notices[1].message, /글꼴/);
});
test('exporter recovers orphan shells in order and retains original indices and safe links', () => {
  const {dom, prepare} = exporter();
  try {
    const result = prepare({html: '<div class="pptx-container">' + shell(2, '<a href="#slide-4">go</a>') + '</div>' + shell(4), slideCount: 2});
    assert.deepEqual(Array.from(result.sourceSlides), [2, 4]);
    assert.equal(result.slideCount, 2);
    assert.match(result.html, /href="#ptv-slide-2"/);
    assert.ok(result.html.indexOf('>go</a>') < result.html.indexOf('>slide</div>'));
    assert.match(result.html, /Content-Security-Policy/);
  } finally {dom.window.close();}
});
test('exporter rejects lost slides and marks ambiguous source indices unknown', () => {
  const {dom, prepare} = exporter();
  try {
    errorCode(() => prepare({html: '<div class="pptx-container">' + shell(1) + '<div class="slide-shell">invalid</div></div>', slideCount: 2}), 'SLIDE_COUNT_MISMATCH');
    const ambiguous = prepare({html: '<div class="pptx-container">' + shell(1) + shell(1) + '</div>', slideCount: 2});
    assert.equal(ambiguous.sourceSlides, null);
    const missing = prepare({html: '<div class="pptx-container"><div class="slide-shell"><div class="slide">ok</div></div></div>'});
    assert.equal(missing.sourceSlides, null);
    const hostile = prepare({html: '<div class="pptx-container">' + shell(1, '<script>bad()</script><img src="https://example.invalid/x">') + '</div>', slideCount: 1});
    assert.ok(!hostile.html.includes('bad()'));
    assert.ok(!hostile.html.includes('example.invalid'));
  } finally {dom.window.close();}
});
