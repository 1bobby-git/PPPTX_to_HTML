'use strict';

const {parentPort, workerData} = require('node:worker_threads');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const crypto = require('node:crypto');
const {pathToFileURL} = require('node:url');
const JSZip = require('jszip');
const {JSDOM} = require('jsdom');
const createDOMPurify = require('dompurify');
const {prepare} = require('./input-package.js');

const root = path.resolve(__dirname, '..');
const LIMIT = 256 * 1024 * 1024;
function progress(message) { parentPort.postMessage({type: 'progress', message}); }
function failure(code, message) { const e = new Error(message); e.code = code; return e; }
function selectedSlides(value, total) {
  if (!value) return new Uint32Array();
  const selected = new Set();
  for (const part of value.split(',')) {
    const m = /^\s*(\d+)(?:\s*-\s*(\d+))?\s*$/.exec(part);
    if (!m) throw failure('INVALID_ARGUMENT', 'Slide range must use 1,3-5 syntax.');
    const first = Number(m[1]);
    const last = m[2] ? Number(m[2]) : first;
    if (!Number.isSafeInteger(first) || !Number.isSafeInteger(last) || first < 1 || last < first || last > total) {
      throw failure('INVALID_ARGUMENT', `Slide numbers must be between 1 and ${total}.`);
    }
    for (let n = first; n <= last; n++) selected.add(n);
  }
  return new Uint32Array([...selected].sort((a, b) => a - b));
}
function diagnostics(json) {
  let items;
  try { items = JSON.parse(json || '[]'); } catch { return []; }
  if (!Array.isArray(items)) return [];
  const names = {
    CHART: 'Some charts may differ from the original.', TABLE: 'Some table styling may differ from the original.',
    SMARTART: 'Some SmartArt uses a preview or fallback.', DIAGRAM: 'Some diagrams use a preview or fallback.',
    OLE: 'Embedded objects are not executable.', MATH: 'Some equations use fallback rendering.',
    MEDIA: 'Some media may not play.', TIMING: 'Animations and transitions are static.',
    IMAGE: 'Some images may be unavailable.', DRAWINGML: 'Some drawing effects may differ.',
    ACTION: 'Some slide actions or links are omitted.', FONT: 'Fonts may be substituted.'
  };
  const grouped = new Map();
  for (const item of items) {
    const code = String(item?.code || 'CONVERSION_NOTICE');
    const part = String(item?.location?.part_name || '');
    if (/NOTE|COMMENT|AUTHOR/.test(code) || /(?:notes|comments|Authors)/i.test(part)) continue;
    if (code === 'OOXML_PART_UNSUPPORTED' && (part.endsWith('/') || !/\/(slides|charts|diagrams)\//.test(part))) continue;
    if (code === 'OOXML_ELEMENT_UNSUPPORTED' && !/\/(slides|charts|diagrams)\//.test(part)) continue;
    if (code === 'OOXML_ELEMENT_UNSUPPORTED' && /^(?:a:chOff|a:chExt|a:masterClrMapping)$/.test(item?.location?.qualified_element_name || '')) continue;
    const family = Object.keys(names).find(name => code.includes(name)) || 'OTHER';
    if (!grouped.has(family)) grouped.set(family, {code: family === 'OTHER' ? 'UNSUPPORTED_FEATURES' : code,
      message: names[family] || 'Some PPTX features may not reproduce fully.', slides: new Set()});
    const index = item?.location?.slide_index;
    if (Number.isInteger(index) && index >= 0) grouped.get(family).slides.add(index + 1);
  }
  return [...grouped.values()].map(({code, message, slides}) => ({code,
    message: message + (slides.size ? ` (slides ${[...slides].sort((a,b) => a-b).slice(0,12).join(', ')}${slides.size > 12 ? ', …' : ''})` : '')}));
}
function normalize(bytes) {
  const source = fs.readFileSync(path.join(root, 'src/normalize-package.js'), 'utf8');
  const context = vm.createContext({JSZip, Uint8Array, Error, Number, Object, String});
  vm.runInContext(source + '\nthis.normalizePptxPackage = normalizePptxPackage;', context, {timeout: 1000});
  return context.normalizePptxPackage(bytes);
}
function exportSafe(html, width, height, title, warnings) {
  const dom = new JSDOM('<!doctype html><html><body></body></html>', {
    runScripts: 'outside-only', resources: undefined, url: 'about:blank'
  });
  try {
    const template = dom.window.document.createElement('template');
    template.innerHTML = html;
    const container = template.content.querySelector('.pptx-container');
    if (!container) throw failure('EXPORTER_ERROR', 'Engine output has no slide container.');
    // The pinned engine can emit an unmatched closing div in a chart placeholder.
    // HTML parsing then leaves later slide shells outside the container.
    for (const shell of template.content.querySelectorAll('.slide-shell')) {
      if (shell.parentElement !== container && shell.querySelector(':scope > .slide')) container.appendChild(shell);
    }
    const sourceSlides = [...container.querySelectorAll(':scope > .slide-shell')]
      .map(shell => Number(shell.getAttribute('data-slide') || shell.querySelector('.slide')?.getAttribute('data-slide')))
      .map(n => Number.isSafeInteger(n) && n >= 1 ? n : null);
    const source = fs.readFileSync(path.join(root, 'src/exporter.js'), 'utf8');
    const viewerCss = fs.readFileSync(path.join(root, 'src/viewer.css'), 'utf8');
    const viewerScript = fs.readFileSync(path.join(root, 'src/viewer.js'), 'utf8');
    const hash = crypto.createHash('sha256').update(viewerScript).digest('base64');
    let script = source;
    for (const [marker, value] of [
      ['/* VIEWER_CSS */ null', viewerCss], ['/* VIEWER_SCRIPT */ null', viewerScript], ['/* VIEWER_HASH */ null', hash]
    ]) {
      if (script.split(marker).length !== 2) throw failure('EXPORTER_ERROR', `Exporter build marker is missing: ${marker}`);
      script = script.replace(marker, () => JSON.stringify(value));
    }
    dom.window.DOMPurify = createDOMPurify(dom.window);
    dom.window.eval(script); // Trusted local exporter only; source PPTX markup is parsed as an inert template.
    const result = dom.window.PptxExport.prepare({html: template.innerHTML, width, height, title, diagnostics: warnings});
    if (Buffer.byteLength(result.html, 'utf8') > LIMIT) throw failure('OUTPUT_LIMIT', 'HTML output exceeds 256 MiB.');
    return {...result, sourceSlides: sourceSlides.length === result.slideCount && sourceSlides.every(Boolean) ? sourceSlides : null};
  } finally { dom.window.close(); }
}
async function main() {
  let info, converted, password = workerData.password || '';
  workerData.password = '';
  try {
    if (workerData.mode === 'check') {
      const packageDir = path.join(root, 'node_modules/@briank-dev/pptx-to-html');
      const engine = await import(pathToFileURL(path.join(packageDir, 'pptx2html_wasm.js')).href);
      engine.initSync({module: fs.readFileSync(path.join(packageDir, 'pptx2html_wasm_bg.wasm'))});
      const sample = '<style>.slide{color:black}</style><div class="pptx-container"><div class="slide-shell" data-slide="1"><div class="slide" data-slide="1">check</div></div></div>';
      const safe = exportSafe(sample, 1280, 720, 'check', []);
      if (safe.slideCount !== 1 || !safe.html.includes('Content-Security-Policy')) throw failure('EXPORTER_ERROR', 'Exporter self-check failed.');
      parentPort.postMessage({type: 'check', ok: true});
      return;
    }
    progress('Checking the PowerPoint package…');
    let data = new Uint8Array(workerData.bytes);
    const prepared = await prepare(data, password);
    password = '';
    data = await normalize(prepared.bytes);
    progress('Loading the offline conversion engine…');
    const packageDir = path.join(root, 'node_modules/@briank-dev/pptx-to-html');
    const wasm = fs.readFileSync(path.join(packageDir, 'pptx2html_wasm_bg.wasm'));
    const engine = await import(pathToFileURL(path.join(packageDir, 'pptx2html_wasm.js')).href);
    engine.initSync({module: wasm});
    info = engine.get_presentation_info(data);
    const {widthPx: width, heightPx: height, slideCount: total} = info;
    if (!Number.isInteger(total) || total < 1 || total > 1000) throw failure('SLIDE_LIMIT', 'Slide count must be between 1 and 1000.');
    if (![width, height].every(v => Number.isFinite(v) && v > 0 && v <= 30000)) throw failure('INVALID_DIMENSIONS', 'Unsupported slide dimensions.');
    const range = selectedSlides(workerData.slides, total);
    // This engine treats an empty range as its default visible-slide set.
    // Enumerate all indices when hidden slides were explicitly requested.
    const effectiveRange = workerData.includeHidden && !range.length
      ? Uint32Array.from({length: total}, (_, index) => index + 1) : range;
    progress('Converting slides…');
    converted = engine.convert_with_options_metadata(data, true, !!workerData.includeHidden, effectiveRange, 1);
    if (!converted.slideCount) throw failure('EMPTY_SELECTION', 'No visible slides in the selected range.');
    if (Buffer.byteLength(converted.html, 'utf8') > LIMIT) throw failure('OUTPUT_LIMIT', 'Engine HTML exceeds 256 MiB.');
    const warnings = diagnostics(converted.diagnosticsJson);
    progress('Sanitizing and packaging HTML…');
    const safe = exportSafe(converted.html, width, height, workerData.title, warnings);
    if (!safe.slideCount) throw failure('EMPTY_SELECTION', 'No visible slides in the selected range.');
    if (safe.slideCount !== converted.slideCount) throw failure('SLIDE_COUNT_MISMATCH', 'Some converted slides could not be safely exported.');
    parentPort.postMessage({type: 'result', html: safe.html, slideCount: safe.slideCount, width, height,
      warnings: safe.diagnostics, sourceSlides: safe.sourceSlides});
  } catch (error) {
    let message = String(error?.message || 'Conversion failed.');
    if (error?.code === 'UNKNOWN_CONTAINER') message = 'Unrecognized file container. The file may use another format or protection scheme.';
    if (error?.code === 'DRM_PROTECTED') message = 'This document has recognized DRM/IRM protection. Export an authorized plain PPTX copy.';
    parentPort.postMessage({type: 'error', code: error?.code || 'CONVERSION_FAILED', message: message.slice(0, 1000)});
  } finally {
    password = '';
    converted?.free();
    info?.free();
  }
}
main();
