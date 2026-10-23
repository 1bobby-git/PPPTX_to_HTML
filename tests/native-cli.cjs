#!/usr/bin/env node
'use strict';

// End-to-end contract for the Node/WASM CLI. Run after creating fixtures:
// node scripts/create-fixtures.cjs && node scripts/create-protected-fixtures.cjs
// node tests/native-cli.cjs
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const {spawnSync} = require('node:child_process');
const {pathToFileURL} = require('node:url');
const crypto = require('node:crypto');
const {JSDOM} = require('jsdom');
const {chromium} = require('playwright');
const {PASSWORD} = require('../scripts/create-protected-fixtures.cjs');

const root = path.resolve(__dirname, '..');
const cli = path.join(root, 'tools/convert-pptx.cjs');
const fixture = name => path.join(__dirname, 'fixtures', name);
const reference = name => path.join(__dirname, 'reference', name);
const sha = filename => crypto.createHash('sha256').update(fs.readFileSync(filename)).digest('hex');
const temporary = fs.mkdtempSync(path.join(os.tmpdir(), 'pptx-native-cli-test-'));
const output = name => path.join(temporary, name);
const guard = output('deny-subprocess-and-network.cjs');

// The guard is loaded in both the CLI and its worker. A network attempt or a
// LibreOffice/PowerPoint/PDF/PNG helper invocation will fail the conversion.
fs.writeFileSync(guard, `
'use strict';
const child = require('node:child_process');
for (const name of ['spawn', 'spawnSync', 'exec', 'execSync', 'execFile', 'execFileSync', 'fork']) {
  child[name] = () => { throw new Error('FORBIDDEN_SUBPROCESS_' + name); };
}
for (const name of ['node:net', 'node:http', 'node:https', 'node:dgram']) {
  const mod = require(name);
  for (const key of ['connect', 'createConnection', 'request', 'get', 'createSocket']) {
    if (typeof mod[key] === 'function') mod[key] = () => { throw new Error('FORBIDDEN_NETWORK_' + name + '_' + key); };
  }
}
globalThis.fetch = () => { throw new Error('FORBIDDEN_FETCH'); };
if (process.env.NATIVE_TEST_WORKER_MODE) {
  const {EventEmitter} = require('node:events');
  const threads = require('node:worker_threads');
  threads.Worker = class FakeWorker extends EventEmitter {
    constructor() {
      super();
      if (process.env.NATIVE_TEST_WORKER_MODE === 'exit') setImmediate(() => this.emit('exit', 17));
    }
    terminate() { return Promise.resolve(0); }
  };
}
`);

function run(args, {stdin, guarded = true, timeout = 180000, workerMode} = {}) {
  const env = {...process.env};
  if (guarded) env.NODE_OPTIONS = [env.NODE_OPTIONS, '--require=' + JSON.stringify(guard)].filter(Boolean).join(' ');
  if (workerMode) env.NATIVE_TEST_WORKER_MODE = workerMode;
  const result = spawnSync(process.execPath, [cli, ...args.map(String)], {
    cwd: root, env, input: stdin, encoding: 'utf8', timeout,
    maxBuffer: 16 * 1024 * 1024
  });
  assert.ifError(result.error);
  return result;
}
function json(result, expectedStatus) {
  assert.equal(result.status, expectedStatus, `status=${result.status}\nstdout=${result.stdout}\nstderr=${result.stderr}`);
  assert.ok(result.stdout.trim(), 'JSON mode needs one stdout object');
  let value;
  assert.doesNotThrow(() => { value = JSON.parse(result.stdout); }, `stdout must contain one JSON object: ${result.stdout}`);
  assert.ok(value && typeof value === 'object' && !Array.isArray(value));
  return value;
}
function html(filename) {
  assert.ok(fs.statSync(filename).size > 1000, `missing or empty HTML output: ${filename}`);
  const source = fs.readFileSync(filename, 'utf8');
  const document = new JSDOM(source).window.document;
  assert.equal(document.querySelectorAll('.pptx-container > .slide-shell').length,
    document.querySelectorAll('.slide').length);
  assert.ok(document.querySelector('meta[http-equiv="Content-Security-Policy"]'));
  assert.ok(!source.includes('PRIVATE_SPEAKER_NOTE_MUST_NOT_EXPORT'), 'speaker notes leaked into output');
  return {source, document, slides: [...document.querySelectorAll('.slide')]};
}
function errorCode(result, code, expectedStatus = 1) {
  const report = json(result, expectedStatus);
  assert.equal(report?.code, code, JSON.stringify(report));
  assert.ok(!result.stdout.includes(PASSWORD) && !result.stderr.includes(PASSWORD), 'password leaked into diagnostics');
  return report;
}
function browserExecutable() {
  return [process.env.CHROME_BIN, process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH,
    chromium.executablePath(), '/tmp/pptx-browsers/chrome-linux64/chrome',
    '/usr/bin/chromium', '/usr/bin/google-chrome']
    .find(filename => filename && fs.existsSync(filename));
}

async function browserCheck(files) {
  const executablePath = browserExecutable();
  assert.ok(executablePath, 'Chromium is required for the native CLI file:// check');
  const browser = await chromium.launch({headless: true, executablePath, args: ['--no-sandbox']});
  const requests = [], errors = [];
  try {
    const context = await browser.newContext({viewport: {width: 1280, height: 850}});
    await context.route(/^https?:\/\//, route => {
      requests.push(route.request().url());
      return route.abort('internetdisconnected');
    });
    const page = await context.newPage();
    page.on('pageerror', error => errors.push(error.message));
    async function open(filename) {
      await page.goto(pathToFileURL(filename).href);
      assert.equal(new URL(page.url()).protocol, 'file:');
      await page.locator('.pptx-container > .slide-shell').first().waitFor();
      return page.locator('.pptx-container > .slide-shell');
    }
    let shells = await open(files.regular);
    assert.equal(await shells.count(), 3);
    const slideText = await page.locator('.slide').allTextContents();
    assert.match(slideText[0], /TABLE_FILE_TWO/);
    assert.match(slideText[1], /INTRO_FILE_ONE/);
    assert.match(slideText[2], /IMAGE_CHART_FILE_THREE/);
    assert.ok(await page.locator('img').evaluateAll(images => images.some(img => img.complete && img.naturalWidth > 0)),
      'embedded image must decode from file://');
    await page.locator('#ptv-mode').click();
    assert.equal(await page.locator('#ptv-mode').getAttribute('aria-pressed'), 'true');
    assert.equal(await page.locator('.slide-shell:visible').count(), 1);
    await page.locator('#ptv-next').click();
    assert.equal(await page.locator('#ptv-select').inputValue(), '1');
    await page.setViewportSize({width: 390, height: 844});
    const mobile = await page.evaluate(() => ({scroll: document.documentElement.scrollWidth, viewport: innerWidth}));
    assert.ok(mobile.scroll <= mobile.viewport + 1, `mobile overflow: ${JSON.stringify(mobile)}`);
    shells = await open(files.hidden);
    assert.equal(await shells.count(), 4);
    shells = await open(files.selected);
    assert.equal(await shells.count(), 1);
    assert.match(await page.locator('.slide').first().textContent(), /INTRO_FILE_ONE/);
    await open(files.hostile);
    assert.equal(await page.evaluate(() => window.__pptxExecuted), undefined);
    assert.equal(await page.evaluate(() => window.__pptxTitleExecuted), undefined);
    assert.deepEqual(requests, [], `unexpected network requests: ${requests.join(', ')}`);
    assert.deepEqual(errors, [], `uncaught browser errors: ${errors.join(', ')}`);
    await context.close();
  } finally { await browser.close(); }
}

async function main() {
  assert.ok(fs.existsSync(cli), `missing ${cli}`);
  assert.ok(fs.existsSync(fixture('complex.pptx')), 'run node scripts/create-fixtures.cjs first');
  assert.ok(fs.existsSync(fixture('protected-standard.pptx')), 'run node scripts/create-protected-fixtures.cjs first');
  const source = output('발표 $(touch SHOULD_NOT_EXIST) space.pptx');
  fs.copyFileSync(fixture('complex.pptx'), source);
  const sourceHash = sha(source);

  const check = json(run(['--check', '--json']), 0);
  assert.equal(check.ok, true);
  assert.ok(!/libreoffice|soffice|powerpoint|pymupdf/i.test(JSON.stringify(check)), 'check reports an external converter');
  assert.match(run(['--help']).stdout, /--slides/);
  const regular = output('regular.html');
  assert.equal(json(run([source, '-o', regular, '--json']), 0).ok, true);
  const deck = html(regular);
  assert.equal(deck.slides.length, 3);
  assert.match(deck.slides[0].textContent, /TABLE_FILE_TWO/);
  assert.match(deck.slides[0].textContent, /사과/);
  assert.match(deck.slides[1].textContent, /INTRO_FILE_ONE/);
  assert.match(deck.slides[1].textContent, /안녕하세요 오프라인/);
  assert.match(deck.slides[2].textContent, /IMAGE_CHART_FILE_THREE/);
  assert.ok(deck.document.querySelector('img[src^="data:image/"]'), 'embedded image is absent');
  assert.ok(!deck.source.includes('HIDDEN_FILE_FOUR'));
  assert.equal(sha(source), sourceHash, 'source PPTX was modified');
  assert.ok(!fs.existsSync(path.join(root, 'SHOULD_NOT_EXIST')), 'filename triggered shell expansion');

  const hidden = output('hidden.html');
  const selected = output('selected.html');
  const hostile = output('hostile.html');
  const hiddenReport = json(run([source, '-o', hidden, '--include-hidden', '--json']), 0);
  assert.equal(hiddenReport.ok, true);
  assert.deepEqual(hiddenReport.sourceSlides, [1, 2, 3, 4]);
  assert.equal(html(hidden).slides.length, 4);
  assert.match(html(hidden).slides[3].textContent, /HIDDEN_FILE_FOUR/);
  const selectedReport = json(run([source, '-o', selected, '--slides', '2', '--json']), 0);
  assert.equal(selectedReport.ok, true);
  assert.deepEqual(selectedReport.sourceSlides, [2]);
  assert.equal(html(selected).slides.length, 1);
  assert.match(html(selected).slides[0].textContent, /INTRO_FILE_ONE/);
  const selectedHiddenDefault = output('selected-hidden-default.html');
  const selectedHiddenIncluded = output('selected-hidden-included.html');
  const omitted = json(run([source, '-o', selectedHiddenDefault, '--slides', '1,4', '--json']), 0);
  assert.deepEqual(omitted.sourceSlides, [1]);
  assert.equal(omitted.slideCount, 1);
  assert.equal(html(selectedHiddenDefault).slides.length, 1);
  const included = json(run([source, '-o', selectedHiddenIncluded, '--slides', '1,4', '--include-hidden', '--json']), 0);
  assert.deepEqual(included.sourceSlides, [1, 4]);
  assert.equal(included.slideCount, 2);
  assert.equal(html(selectedHiddenIncluded).slides.length, 2);
  const title = '<img src=x onerror="window.__pptxTitleExecuted=1">';
  assert.equal(json(run([fixture('hostile.pptx'), '-o', hostile, '--title', title, '--json']), 0).ok, true);
  const hostileDeck = html(hostile);
  assert.ok(!hostileDeck.document.querySelector('script[src], img[src^="http"], iframe, object, embed'));
  assert.ok(!hostileDeck.source.includes('https://example.invalid/pptx-leak.png'));
  assert.ok(!hostileDeck.document.querySelector('[onerror], [onclick]'));
  assert.equal(hostileDeck.document.title, title);
  const chart = output('chart.html');
  assert.equal(json(run([fixture('basic-chart.pptx'), '-o', chart, '--json']), 0).ok, true);
  assert.match(html(chart).slides[0].textContent, /PIE_CHART_VALUES_2_4_3/);
  const pie = html(chart).document;
  assert.ok(pie.querySelectorAll('.chart-direct svg path.chart-pie-slice').length >= 3,
    'pie chart must render direct SVG segments');
  for (const label of ['Mon', 'Tue', 'Wed', '22%', '44%', '33%']) {
    assert.ok(pie.querySelector('.chart-direct').textContent.includes(label), `pie value is missing: ${label}`);
  }

  const prefixed = output('prefixed.html');
  assert.equal(json(run([fixture('prefixed.pptx'), '-o', prefixed, '--json']), 0).ok, true);
  assert.match(html(prefixed).slides[0].textContent, /TABLE_FILE_TWO/);

  const original = sha(regular);
  errorCode(run([source, '-o', regular, '--json']), 'OUTPUT_EXISTS');
  assert.equal(sha(regular), original, 'refusal modified existing output');
  errorCode(run([fixture('unknown.pptx'), '-o', regular, '--overwrite', '--json']), 'UNKNOWN_CONTAINER');
  assert.equal(sha(regular), original, 'failed overwrite modified existing output');
  errorCode(run([source, '-o', source, '--overwrite', '--json']), 'PATH_COLLISION');
  assert.equal(sha(source), sourceHash, 'input/output collision modified source');
  const empty = output('empty.pptx');
  const oversized = output('oversized.pptx');
  fs.writeFileSync(empty, '');
  fs.closeSync(fs.openSync(oversized, 'w'));
  fs.truncateSync(oversized, 64 * 1024 * 1024 + 1);
  for (const input of [empty, oversized]) {
    const destination = input + '.html';
    errorCode(run([input, '-o', destination, '--json']), 'INPUT_LIMIT');
    assert.ok(!fs.existsSync(destination), 'invalid-size input created an output');
  }
  errorCode(run([fixture('drm-marker.pptx'), '-o', output('drm.html'), '--json']), 'DRM_PROTECTED');
  errorCode(run([fixture('malformed-cfb.pptx'), '-o', output('malformed.html'), '--json']), 'INVALID_CONTAINER');
  errorCode(run([fixture('protected-standard.pptx'), '-o', output('missing-password.html'), '--json']), 'PASSWORD_REQUIRED');
  errorCode(run([fixture('protected-standard.pptx'), '-o', output('wrong-password.html'), '--password-stdin', '--json'],
    {stdin: 'incorrect-password\n'}), 'INCORRECT_PASSWORD');
  const decrypted = output('decrypted.html');
  assert.equal(json(run([fixture('protected-standard.pptx'), '-o', decrypted, '--password-stdin', '--json'],
    {stdin: PASSWORD + '\n'}), 0).ok, true);
  assert.match(html(decrypted).slides[0].textContent, /TABLE_FILE_TWO/);
  const agile = output('agile.html');
  assert.equal(json(run([reference('complex-agile.pptx'), '-o', agile, '--password-stdin', '--json'],
    {stdin: PASSWORD + '\n'}), 0).ok, true);
  assert.match(html(agile).slides[0].textContent, /TABLE_FILE_TWO/);
  assert.ok(!fs.existsSync(output('wrong-password.html')));

  errorCode(run([source, '--slides', '0', '-o', output('bad-range.html'), '--json']), 'INVALID_ARGUMENT', 2);
  errorCode(run([source, '--timeout', '0', '-o', output('bad-timeout.html'), '--json']), 'INVALID_ARGUMENT', 2);
  errorCode(run([source, '-o', regular, '--overwrite', '--timeout', '1', '--json'],
    {workerMode: 'hang', timeout: 10000}), 'TIMEOUT');
  assert.equal(sha(regular), original, 'timeout modified existing output');
  errorCode(run([source, '-o', regular, '--overwrite', '--json'],
    {workerMode: 'exit'}), 'WORKER_EXIT');
  assert.equal(sha(regular), original, 'worker exit modified existing output');
  await browserCheck({regular, hidden, selected, hostile});
  console.log('PASS native CLI: offline WASM conversion, slide order/visibility/range, media/chart, password and container classification, atomic output, JSON/arguments, file:// viewer, zero HTTP(S)');
}

main().catch(error => { console.error(error); process.exitCode = 1; });
