#!/usr/bin/env node
'use strict';

// HTTP + browser acceptance for the web UI that runs the native CLI.
// Run after scripts/create-fixtures.cjs and create-protected-fixtures.cjs.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const http = require('node:http');
const {pathToFileURL} = require('node:url');
const {JSDOM} = require('jsdom');
const {chromium} = require('playwright');
const {PASSWORD} = require('../scripts/create-protected-fixtures.cjs');
const {createWebServer} = require('../tools/serve-web.cjs');

const root = path.resolve(__dirname, '..');
const fixture = name => path.join(__dirname, 'fixtures', name);
const temporary = fs.mkdtempSync(path.join(os.tmpdir(), 'pptx-web-test-'));
const instances = [];
let sequence = 0;

async function eventually(check, message, timeout = 10000) {
  const limit = Date.now() + timeout;
  do {
    if (await check()) return;
    await new Promise(resolve => setTimeout(resolve, 25));
  } while (Date.now() < limit);
  throw new Error(message);
}

async function openService(options = {}) {
  const tempRoot = path.join(temporary, 'jobs-' + ++sequence);
  fs.mkdirSync(tempRoot);
  const service = createWebServer({host: '127.0.0.1', port: 0, tempRoot, ...options});
  instances.push(service);
  await new Promise((resolve, reject) => {
    service.server.once('error', reject);
    service.server.listen(options.port || 0, '127.0.0.1', resolve);
  });
  service.origin = 'http://127.0.0.1:' + service.server.address().port;
  service.tempRoot = tempRoot;
  return service;
}

async function cleaned(service) {
  await eventually(() => !service.getState().busy && fs.readdirSync(service.tempRoot).length === 0,
    'service did not release its slot and delete temporary document files');
}

function requestHeaders(service, extra = {}) {
  return {Origin: service.origin, 'X-Pptx-Token': service.token, ...extra};
}

function form(filename = 'complex.pptx', options = {}) {
  const data = new FormData();
  data.set('file', new Blob([fs.readFileSync(fixture(filename))]), filename);
  data.set('options', JSON.stringify(options));
  return data;
}

async function convert(service, filename, options, extra = {}) {
  return fetch(service.origin + '/api/convert', {
    method: 'POST', headers: requestHeaders(service), body: form(filename, options), ...extra
  });
}

async function apiError(response, status, code) {
  assert.equal(response.status, status);
  assert.match(response.headers.get('content-type') || '', /application\/json/);
  const value = await response.json();
  assert.equal(value.ok, false);
  assert.equal(typeof value.code, 'string');
  if (code) assert.equal(value.code, code, JSON.stringify(value));
  assert.equal(typeof value.error, 'string');
  assert.ok(!JSON.stringify(value).includes(temporary), 'absolute temporary path leaked');
  assert.ok(!JSON.stringify(value).includes(PASSWORD), 'document password leaked');
  return value;
}

async function converted(response, count) {
  assert.equal(response.status, 200, await (response.status === 200 ? Promise.resolve('') : response.text()));
  assert.match(response.headers.get('content-type') || '', /text\/html/);
  assert.match(response.headers.get('cache-control') || '', /no-store/);
  const header = response.headers.get('x-pptx-metadata');
  assert.ok(header && header.length < 16384, 'bounded conversion metadata header is required');
  const metadata = JSON.parse(Buffer.from(header, 'base64url').toString('utf8'));
  assert.equal(metadata.slideCount, count);
  assert.equal(metadata.output, undefined, 'CLI filesystem path must not be returned');
  const html = await response.text();
  assert.ok(!html.includes(PASSWORD), 'password appeared in HTML');
  assert.ok(!html.includes('PRIVATE_SPEAKER_NOTE_MUST_NOT_EXPORT'), 'speaker notes appeared in HTML');
  const dom = new JSDOM(html);
  const document = dom.window.document;
  assert.equal(document.querySelectorAll('.pptx-container > .slide-shell').length, count);
  assert.ok(document.querySelector('meta[http-equiv="Content-Security-Policy"]'));
  assert.ok(!document.querySelector('script[src], img[src^="http"], iframe, object, embed'));
  return {metadata, html, dom, document};
}

function rawRequest(service, {method = 'GET', pathname = '/api/status', headers = {}, chunks = []} = {}) {
  return new Promise((resolve, reject) => {
    const req = http.request(service.origin + pathname, {method, headers}, res => {
      const buffers = [];
      res.on('data', chunk => buffers.push(chunk));
      res.on('error', reject);
      res.on('end', () => resolve({status: res.statusCode, headers: res.headers, body: Buffer.concat(buffers).toString('utf8')}));
    });
    req.on('error', reject);
    for (const chunk of chunks) req.write(chunk);
    req.end();
  });
}

function fakeCli(name) {
  const filename = path.join(temporary, name + '.cjs');
  fs.writeFileSync(filename, `
    'use strict';
    const fs = require('node:fs');
    fs.writeFileSync(__filename + '.starting', JSON.stringify({pid:process.pid,args:process.argv.slice(2)}));
    fs.renameSync(__filename + '.starting', __filename + '.started');
    process.stdin.resume();
    process.on('SIGTERM', () => {});
    setInterval(() => {}, 1000);
  `);
  return filename;
}

function alive(pid) {
  try { process.kill(pid, 0); return true; }
  catch (error) { if (error.code === 'ESRCH') return false; throw error; }
}

async function started(filename) {
  await eventually(() => fs.existsSync(filename + '.started'), 'test CLI was not started');
  const details = JSON.parse(fs.readFileSync(filename + '.started', 'utf8'));
  assert.ok(alive(details.pid), 'CLI process unexpectedly exited');
  assert.ok(!details.args.some(arg => arg.includes(PASSWORD)), 'password passed in CLI arguments');
  assert.ok(details.args.includes('--password-stdin'));
  return details;
}

async function protocolAndConversion() {
  const service = await openService();
  const page = await fetch(service.origin);
  assert.equal(page.status, 200);
  assert.match(page.headers.get('cache-control') || '', /no-store/);
  assert.match(page.headers.get('content-security-policy') || '', /connect-src 'self'/);
  const shell = new JSDOM(await page.text());
  assert.equal(shell.window.document.querySelector('meta[name="pptx-csrf"]').content, service.token);
  assert.equal(shell.window.document.querySelector('#engine-worker'), null, 'web service should not download the browser engine');
  shell.window.close();
  const state = await (await fetch(service.origin + '/api/status')).json();
  assert.equal(state.ok, true);
  assert.equal(state.engine, 'native');
  assert.equal(state.mode, 'local');
  assert.equal(state.busy, false);
  assert.equal(state.csrfToken, service.token, 'status must refresh the token for an already open page');
  const asset = await fetch(service.origin + '/assets/server-client.js');
  assert.equal(asset.status, 200);
  assert.match(await asset.text(), /\/api\/convert/);
  assert.equal((await fetch(service.origin + '/tools/convert-pptx.cjs')).status, 404);
  const wrongHost = await rawRequest(service, {headers: {Host: 'untrusted.example'}});
  assert.ok([403, 421].includes(wrongHost.status), 'untrusted Host was accepted');
  await apiError(await convert(service, 'complex.pptx', {}, {
    headers: requestHeaders(service, {Origin: 'https://untrusted.example'})
  }), 403);
  await apiError(await convert(service, 'complex.pptx', {}, {
    headers: {Origin: service.origin}
  }), 403);
  await apiError(await convert(service, 'complex.pptx', {}, {
    headers: requestHeaders(service, {'Sec-Fetch-Site': 'cross-site'})
  }), 403);
  await cleaned(service);

  // Only child Node processes load this guard. The web server itself must use
  // HTTP, but the conversion process must never start Office or use the network.
  const guard = path.join(temporary, 'native-offline-guard.cjs');
  fs.writeFileSync(guard, `
    const child = require('node:child_process');
    for (const name of ['spawn','spawnSync','exec','execSync','execFile','execFileSync','fork'])
      child[name] = () => { throw new Error('FORBIDDEN_CONVERTER_SUBPROCESS'); };
    for (const name of ['node:http','node:https','node:net','node:dgram']) {
      const api = require(name);
      for (const method of ['request','get','connect','createConnection','createSocket'])
        if (typeof api[method] === 'function') api[method] = () => { throw new Error('FORBIDDEN_CONVERTER_NETWORK'); };
    }
    globalThis.fetch = () => { throw new Error('FORBIDDEN_CONVERTER_FETCH'); };
  `);
  const previousNodeOptions = process.env.NODE_OPTIONS;
  process.env.NODE_OPTIONS = [previousNodeOptions, '--require=' + JSON.stringify(guard)].filter(Boolean).join(' ');
  try {
    const normal = await converted(await convert(service, 'complex.pptx', {title: '웹 변환 검증'}), 3);
    assert.equal(normal.document.title, '웹 변환 검증');
    assert.match(normal.document.querySelector('.slide').textContent, /TABLE_FILE_TWO/);
    assert.match(normal.document.body.textContent, /안녕하세요 오프라인/);
    assert.ok(normal.document.querySelector('img[src^="data:image/"]'));
    assert.deepEqual(normal.metadata.sourceSlides, [1, 2, 3]);
    normal.dom.window.close();
    await cleaned(service);
    const selected = await converted(await convert(service, 'complex.pptx', {slideRange: '1,4', includeHidden: true}), 2);
    assert.deepEqual(selected.metadata.sourceSlides, [1, 4]);
    assert.match(selected.document.body.textContent, /HIDDEN_FILE_FOUR/);
    selected.dom.window.close();
    await cleaned(service);
    const protectedResult = await converted(await convert(service, 'protected-standard.pptx', {password: PASSWORD}), 3);
    assert.match(protectedResult.document.body.textContent, /INTRO_FILE_ONE/);
    protectedResult.dom.window.close();
    await cleaned(service);
    await apiError(await convert(service, 'protected-standard.pptx', {password: 'incorrect-password'}), 422, 'INCORRECT_PASSWORD');
    await cleaned(service);
    const unknown = await apiError(await convert(service, 'unknown.pptx'), 422, 'UNKNOWN_CONTAINER');
    assert.ok(!unknown.error.includes('확장자만'), 'unknown containers must not be blamed on a renamed extension');
    await cleaned(service);
    await apiError(await convert(service, 'drm-marker.pptx'), 422, 'DRM_PROTECTED');
    await cleaned(service);
  } finally {
    if (previousNodeOptions === undefined) delete process.env.NODE_OPTIONS;
    else process.env.NODE_OPTIONS = previousNodeOptions;
  }
  return service;
}

async function lifecycleChecks() {
  const cliPath = fakeCli('cancel-cli');
  const service = await openService({cliPath, shutdownGraceMs: 100});
  const controller = new AbortController();
  const pending = convert(service, 'complex.pptx', {password: PASSWORD}, {signal: controller.signal});
  pending.catch(() => {});
  const child = await started(cliPath);
  await apiError(await convert(service, 'complex.pptx'), 429);
  controller.abort();
  await assert.rejects(pending, error => error.name === 'AbortError');
  await cleaned(service);
  await eventually(() => !alive(child.pid), 'cancelled CLI child was not reaped');

  const timeoutCli = fakeCli('timeout-cli');
  const timed = await openService({cliPath: timeoutCli, conversionTimeoutMs: 700, shutdownGraceMs: 100});
  const timeoutRequest = convert(timed, 'complex.pptx');
  const timeoutChild = await started(timeoutCli);
  await apiError(await timeoutRequest, 504);
  await cleaned(timed);
  assert.equal(alive(timeoutChild.pid), false, 'timed out CLI child remains alive');

  const limitCli = fakeCli('limit-cli');
  const limited = await openService({cliPath: limitCli, maxInputBytes: 1024});
  const boundary = 'pptx-web-test-boundary';
  const tooLarge = await rawRequest(limited, {
    method: 'POST', pathname: '/api/convert',
    headers: requestHeaders(limited, {'Content-Type': 'multipart/form-data; boundary=' + boundary}),
    chunks: [Buffer.from('--' + boundary + '\r\nContent-Disposition: form-data; name="file"; filename="large.pptx"\r\nContent-Type: application/octet-stream\r\n\r\n'),
      Buffer.alloc(8192, 65), Buffer.from('\r\n--' + boundary + '--\r\n')]
  });
  assert.equal(tooLarge.status, 413, tooLarge.body);
  assert.equal(fs.existsSync(limitCli + '.started'), false, 'oversized upload started the converter');
  await cleaned(limited);

  const uploadCli = fakeCli('upload-abort-cli');
  const uploading = await openService({cliPath: uploadCli});
  const partial = http.request(uploading.origin + '/api/convert', {
    method: 'POST', headers: requestHeaders(uploading, {'Content-Type': 'multipart/form-data; boundary=' + boundary})
  });
  partial.on('error', () => {}); // Deliberate client disconnect below.
  partial.on('response', response => response.resume());
  partial.write('--' + boundary + '\r\nContent-Disposition: form-data; name="file"; filename="partial.pptx"\r\n\r\n');
  partial.write(Buffer.alloc(128, 66));
  await eventually(() => uploading.getState().busy && fs.readdirSync(uploading.tempRoot).length > 0,
    'partial upload was not accepted');
  partial.destroy();
  await cleaned(uploading);
  assert.equal(fs.existsSync(uploadCli + '.started'), false, 'incomplete upload started the converter');

  // A client that stops reading must not retain the service's only slot or
  // its temporary output indefinitely after the CLI itself has completed.
  const outputCli = path.join(temporary, 'large-output-cli.cjs');
  fs.writeFileSync(outputCli, `
    const args = process.argv.slice(2);
    require('node:fs').writeFileSync(args[args.indexOf('-o') + 1], Buffer.alloc(32 * 1024 * 1024, 32));
    process.stdout.write(JSON.stringify({ok:true,slideCount:1,width:960,height:540,warnings:[],sourceSlides:[1]}));
  `);
  const slowReader = await openService({cliPath: outputCli, responseTimeoutMs: 150});
  const response = await new Promise((resolve, reject) => {
    const request = http.request(slowReader.origin + '/api/convert', {
      method: 'POST', headers: requestHeaders(slowReader, {'Content-Type': 'multipart/form-data; boundary=' + boundary})
    }, result => {
      result.on('error', () => {});
      result.pause();
      resolve(result);
    });
    request.on('error', reject);
    request.write('--' + boundary + '\r\nContent-Disposition: form-data; name="options"\r\n\r\n{}\r\n');
    request.write('--' + boundary + '\r\nContent-Disposition: form-data; name="file"; filename="test.pptx"\r\n\r\n');
    request.write(fs.readFileSync(fixture('complex.pptx')));
    request.end('\r\n--' + boundary + '--\r\n');
  });
  try {
    assert.equal(response.statusCode, 200);
    assert.equal(Number(response.headers['content-length']), 32 * 1024 * 1024);
    await cleaned(slowReader);
    assert.equal(response.complete, false, 'slow client unexpectedly received the complete output');
  } finally { response.destroy(); }
}

function browserExecutable() {
  return [process.env.CHROME_BIN, process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH,
    chromium.executablePath(), '/tmp/pptx-browsers/chrome-linux64/chrome', '/usr/bin/chromium', '/usr/bin/google-chrome']
    .find(filename => filename && fs.existsSync(filename));
}

async function waitForViewer(scope, expectedSlides) {
  // Markup may become visible before the script installs controls/listeners.
  // Script-owned DOM must exist; a CSP-blocked viewer still fails this gate.
  await scope.locator('#ptv-select option[value="' + (expectedSlides - 1) + '"]')
    .waitFor({state: 'attached', timeout: 30000});
  await scope.locator('#ptv-slides > .slide-shell[aria-current="true"]')
    .waitFor({state: 'attached', timeout: 30000});
  assert.equal(await scope.locator('#ptv-select option').count(), expectedSlides,
    'viewer must initialize one option per displayed slide');
  assert.equal(await scope.locator('#ptv-slides > .slide-shell[aria-current="true"]').count(), 1,
    'viewer must initialize exactly one current slide');
}

async function browserCheck(service) {
  const executablePath = browserExecutable();
  assert.ok(executablePath, 'Chromium is required for the web service UI check');
  const browser = await chromium.launch({headless: true, executablePath, args: ['--no-sandbox']});
  const externalRequests = [], errors = [], cspErrors = [];
  try {
    const context = await browser.newContext({viewport: {width: 1280, height: 900}, acceptDownloads: true});
    await context.route(/^https?:\/\//, route => {
      if (new URL(route.request().url()).origin === service.origin) return route.continue();
      externalRequests.push(route.request().url());
      return route.abort('internetdisconnected');
    });
    const page = await context.newPage();
    page.on('pageerror', error => errors.push(error.message));
    page.on('console', message => {
      if (message.type() === 'error' && /Content Security Policy|violates.*directive/i.test(message.text())) {
        cspErrors.push(message.text());
      }
    });
    await page.addInitScript(() => {
      window.Worker = class { constructor() { throw new Error('WEB_MODE_MUST_USE_SERVER_CLI'); } };
    });
    await page.goto(service.origin);
    await page.locator('#file-input').setInputFiles(fixture('complex.pptx'));
    await page.locator('#output-title').fill('사용자 웹 변환');
    await page.locator('#convert-button').click();
    await page.locator('#status-line[data-tone="success"]').waitFor({timeout: 120000});
    assert.match(await page.locator('#result-subtitle').textContent(), /3개/);
    const preview = page.frameLocator('#preview-frame');
    await preview.locator('.slide-shell').first().waitFor();
    assert.equal(await preview.locator('.slide-shell').count(), 3);
    await waitForViewer(preview, 3);
    await preview.locator('#ptv-mode').click();
    assert.equal(await preview.locator('#ptv-mode').getAttribute('aria-pressed'), 'true');
    await preview.locator('#ptv-next').click();
    assert.equal(await preview.locator('#ptv-select').inputValue(), '1');
    await preview.locator('img').first().evaluate(image => image.decode());
    assert.ok(await preview.locator('img').evaluateAll(images => images.some(image => image.complete && image.naturalWidth > 0)));
    const screenshots = path.join(root, 'test-results');
    fs.mkdirSync(screenshots, {recursive: true});
    await page.screenshot({path: path.join(screenshots, 'web-service-desktop.png'), fullPage: true});
    const downloadEvent = page.waitForEvent('download');
    await page.locator('#download-button').click();
    const download = await downloadEvent;
    assert.equal(download.suggestedFilename(), '사용자 웹 변환.html');
    const saved = path.join(temporary, 'web-result.html');
    await download.saveAs(saved);
    assert.ok(fs.statSync(saved).size > 1000);
    const outputPage = await context.newPage();
    await outputPage.goto(pathToFileURL(saved).href);
    await outputPage.locator('.slide-shell').first().waitFor();
    await waitForViewer(outputPage, 3);
    assert.equal(await outputPage.title(), '사용자 웹 변환');
    await outputPage.locator('#ptv-mode').click();
    await outputPage.locator('#ptv-next').click();
    assert.equal(await outputPage.locator('#ptv-select').inputValue(), '1');
    await outputPage.close();
    await page.bringToFront();
    await page.setViewportSize({width: 390, height: 844});
    await page.locator('#preview-frame').scrollIntoViewIfNeeded();
    await preview.locator('#ptv-next').click();
    assert.equal(await preview.locator('#ptv-select').inputValue(), '2');
    const size = await page.evaluate(() => ({scroll: document.documentElement.scrollWidth, viewport: innerWidth}));
    assert.ok(size.scroll <= size.viewport + 1, 'web UI overflowed a mobile viewport');
    await page.screenshot({path: path.join(screenshots, 'web-service-mobile.png'), fullPage: true});
    await cleaned(service);
    assert.deepEqual(externalRequests, [], 'web conversion requested an external origin');
    assert.deepEqual(errors, [], 'uncaught browser errors');
    assert.deepEqual(cspErrors, [], 'conversion page or Blob preview violated its CSP');

    // An already open tab must recover after a service restart. The old page
    // has the previous CSRF token, so reconnect must obtain the new one.
    const originalPort = service.server.address().port;
    await service.close();
    await page.locator('#file-input').setInputFiles(fixture('complex.pptx'));
    await page.locator('#convert-button').click();
    await page.locator('#retry-connection').waitFor({state: 'visible'});
    const restarted = await openService({port: originalPort});
    assert.equal(restarted.origin, service.origin);
    assert.notEqual(restarted.token, service.token);
    await page.locator('#retry-connection').click();
    await page.locator('#convert-button').click();
    await page.locator('#status-line[data-tone="success"]').waitFor({timeout: 120000});
    assert.match(await page.locator('#result-subtitle').textContent(), /3개/);
    await cleaned(restarted);

    // Verify that the visible Cancel control terminates the real server job,
    // rather than merely hiding its progress while the subprocess keeps running.
    const cancelCli = fakeCli('browser-cancel-cli');
    const cancelService = await openService({cliPath: cancelCli, shutdownGraceMs: 100});
    await context.unroute(/^https?:\/\//);
    await context.route(/^https?:\/\//, route => {
      if ([service.origin, cancelService.origin].includes(new URL(route.request().url()).origin)) return route.continue();
      externalRequests.push(route.request().url());
      return route.abort('internetdisconnected');
    });
    await page.goto(cancelService.origin);
    await page.locator('#file-input').setInputFiles(fixture('complex.pptx'));
    await page.locator('#convert-button').click();
    const cancelled = await started(cancelCli);
    await page.locator('#cancel-button').click();
    assert.match(await page.locator('#status-text').textContent(), /취소/);
    assert.equal(await page.locator('#download-button').isDisabled(), true);
    await cleaned(cancelService);
    assert.equal(alive(cancelled.pid), false, 'browser cancellation left a CLI process running');
    assert.deepEqual(externalRequests, [], 'browser cancellation made an external request');
    await context.close();
  } finally { await browser.close(); }
}

async function main() {
  assert.ok(fs.existsSync(fixture('complex.pptx')), 'generate presentation fixtures first');
  assert.ok(fs.existsSync(fixture('protected-standard.pptx')), 'generate protected fixtures first');
  const service = await protocolAndConversion();
  await lifecycleChecks();
  await browserCheck(service);
  console.log('PASS web service: native CLI behind web UI, password/range, Host/Origin/token checks, bounded uploads, cancellation/timeout cleanup, standalone download and no external requests');
}

main().catch(error => {
  console.error(error);
  process.exitCode = 1;
}).finally(async () => {
  await Promise.allSettled(instances.map(service => service.close()));
  fs.rmSync(temporary, {recursive: true, force: true});
});
