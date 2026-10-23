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
const {prepare, inspectChartData, resolveThemes} = require('./input-package.js');
const PptxCore = require('./conversion-core.js');

const root = path.resolve(__dirname, '..');
const LIMIT = 256 * 1024 * 1024;
function progress(message) { parentPort.postMessage({type: 'progress', message}); }
function failure(code, message) { const e = new Error(message); e.code = code; return e; }
function normalize(bytes) {
  const source = fs.readFileSync(path.join(root, 'src/normalize-package.js'), 'utf8');
  const context = vm.createContext({JSZip, Uint8Array, Error, Number, Object, String});
  vm.runInContext(source + '\nthis.normalizePptxPackage = normalizePptxPackage;', context, {timeout: 1000});
  return context.normalizePptxPackage(bytes);
}
function exportSafe(html, width, height, title, warnings, slideCount, chartData) {
  const dom = new JSDOM('<!doctype html><html><body></body></html>', {
    runScripts: 'outside-only', resources: undefined, url: 'about:blank'
  });
  try {
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
    const result = dom.window.PptxExport.prepare({html, width, height, title, diagnostics: warnings, slideCount, chartData});
    if (Buffer.byteLength(result.html, 'utf8') > LIMIT) throw failure('OUTPUT_LIMIT', 'HTML output exceeds 256 MiB.');
    return result;
  } finally { dom.window.close(); }
}
async function main() {
  let password = workerData.password || '';
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
    const themed = await resolveThemes(data, JSZip);
    data = themed.bytes;
    progress('Loading the offline conversion engine…');
    const packageDir = path.join(root, 'node_modules/@briank-dev/pptx-to-html');
    const wasm = fs.readFileSync(path.join(packageDir, 'pptx2html_wasm_bg.wasm'));
    const engine = await import(pathToFileURL(path.join(packageDir, 'pptx2html_wasm.js')).href);
    engine.initSync({module: wasm});
    const converted = PptxCore.render(engine, data, {
      slideRange: workerData.slides, includeHidden: workerData.includeHidden, onProgress: progress
    });
    converted.diagnostics.push(...themed.diagnostics);
    const {width, height} = converted;
    progress('Sanitizing and packaging HTML…');
    const chartData = await inspectChartData(data, JSZip, converted.chartFallbacks, converted.html);
    const safe = exportSafe(converted.html, width, height, workerData.title, converted.diagnostics, converted.slideCount, chartData);
    parentPort.postMessage({type: 'result', html: safe.html, slideCount: safe.slideCount, width, height,
      warnings: safe.diagnostics, sourceSlides: safe.sourceSlides});
  } catch (error) {
    let message = String(error?.message || 'Conversion failed.');
    if (error?.code === 'UNKNOWN_CONTAINER') message = 'Unrecognized file container. The file may use another format or protection scheme.';
    if (error?.code === 'DRM_PROTECTED') message = 'This document has recognized DRM/IRM protection. Export an authorized plain PPTX copy.';
    parentPort.postMessage({type: 'error', code: error?.code || 'CONVERSION_FAILED', message: message.slice(0, 1000)});
  } finally {
    password = '';
  }
}
main();
