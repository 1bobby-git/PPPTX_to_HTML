#!/usr/bin/env node
'use strict';

// Run against the extracted portable application, not an Electron development
// runner. The caller creates the existing fixtures and supplies a display on
// Linux (for example, xvfb-run -a node tests/desktop.cjs).
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const crypto = require('node:crypto');
const {JSDOM} = require('jsdom');
const {_electron: electron} = require(process.env.PPTX_TEST_PLAYWRIGHT || 'playwright');
const {PASSWORD} = require('../scripts/create-protected-fixtures.cjs');

const root = path.resolve(__dirname, '..');
const version = require('../package.json').version;
const runtimePin = require('../desktop/runtime.json');
const engineVersion = require('../vendor/pptx2html-engine/BUILD-INFO.json').package.version;
const fixture = name => path.join(__dirname, 'fixtures', name);
const temporary = fs.mkdtempSync(path.join(os.tmpdir(), 'pptx-desktop-test-'));
const output = name => path.join(temporary, name);
const sha = filename => crypto.createHash('sha256').update(fs.readFileSync(filename)).digest('hex');
const timeout = 120000;

function inspectHtml(filename, expectedSlides) {
  const source = fs.readFileSync(filename, 'utf8');
  assert.ok(Buffer.byteLength(source) > 1000, 'saved HTML is unexpectedly empty');
  assert.ok(!source.includes(PASSWORD), 'document password leaked into HTML');
  assert.ok(!source.includes('PRIVATE_SPEAKER_NOTE_MUST_NOT_EXPORT'), 'speaker notes leaked into HTML');
  const dom = new JSDOM(source);
  try {
    const document = dom.window.document;
    const shells = [...document.querySelectorAll('.pptx-container > .slide-shell')];
    assert.equal(shells.length, expectedSlides, 'saved HTML has the wrong slide count');
    assert.equal(document.querySelectorAll('.slide').length, expectedSlides);
    const csp = document.querySelector('meta[http-equiv="Content-Security-Policy"]');
    assert.ok(csp, 'saved HTML is missing its content security policy');
    assert.match(csp.content, /connect-src 'none'/);
    assert.equal(document.querySelectorAll('script[src], link[href], iframe, object, embed').length, 0,
      'saved HTML loads external or embedded documents');
    assert.equal(document.querySelectorAll('[src^="http:"], [src^="https:"], [onerror], [onclick]').length, 0,
      'saved HTML contains an external resource or event handler');
    return {source, title: document.title, texts: shells.map(element => element.textContent)};
  } finally {
    dom.window.close();
  }
}

async function installDialogs(application) {
  // These substitutions live only in the test's main-process evaluation. The
  // shipped program has no environment-controlled file picker or test route.
  await application.evaluate(({dialog}) => {
    const state = {input: null, output: null, cancelOpen: false, cancelSave: false,
      overwriteResponse: 0, openCalls: 0, saveCalls: 0, messageCalls: 0, lastMessage: null};
    globalThis.__pptxDesktopTestDialogs = state;
    dialog.showOpenDialog = async () => {
      state.openCalls++;
      return state.cancelOpen
        ? {canceled: true, filePaths: []}
        : {canceled: false, filePaths: [state.input]};
    };
    dialog.showSaveDialog = async () => {
      state.saveCalls++;
      return state.cancelSave
        ? {canceled: true, filePath: undefined}
        : {canceled: false, filePath: state.output};
    };
    dialog.showMessageBox = async (_window, options) => {
      const message = options || _window;
      state.messageCalls++;
      state.lastMessage = {buttons: message.buttons, defaultId: message.defaultId,
        cancelId: message.cancelId, message: message.message, detail: message.detail};
      return {response: state.overwriteResponse, checkboxChecked: false};
    };
  });
}

async function main() {
  const executable = process.env.PPTX_DESKTOP_EXECUTABLE;
  assert.ok(executable && path.isAbsolute(executable) && fs.existsSync(executable),
    'PPTX_DESKTOP_EXECUTABLE must name an extracted portable application executable');
  for (const name of ['complex.pptx', 'hostile.pptx', 'unknown.pptx', 'protected-standard.pptx']) {
    assert.ok(fs.existsSync(fixture(name)), 'missing fixture: ' + name);
  }

  // Copy the actual distribution away from the repository. A missing packaged
  // dependency must not resolve accidentally from the developer node_modules.
  const distribution = path.dirname(executable);
  assert.ok(fs.existsSync(path.join(distribution, 'resources/app/package.json')),
    'expected the portable resources/app distribution');
  const isolated = output('portable-runtime');
  fs.cpSync(distribution, isolated, {recursive: true});
  const isolatedExecutable = path.join(isolated, path.basename(executable));
  const source = output('발표 자료 원본 (회사 테스트).pptx');
  fs.copyFileSync(fixture('complex.pptx'), source);
  const sourceHash = sha(source);
  const env = {...process.env};
  delete env.NODE_PATH;
  delete env.NODE_OPTIONS;
  delete env.ELECTRON_RUN_AS_NODE;
  const application = await electron.launch({
    executablePath: isolatedExecutable,
    args: process.platform === 'linux' ? ['--no-sandbox'] : [],
    cwd: temporary,
    env,
    timeout: 60000
  });
  const requests = [], pageErrors = [], policyErrors = [];
  let passed = false;
  try {
    const runtime = await application.evaluate(({app}) => ({
      appPath: app.getAppPath(), version: app.getVersion(), packaged: app.isPackaged,
      node: process.versions.node, electron: process.versions.electron
    }));
    assert.equal(runtime.version, version);
    assert.equal(runtime.packaged, true, 'the test launched a development runner');
    assert.equal(path.resolve(runtime.appPath), path.resolve(isolated, 'resources/app'));
    assert.equal(runtime.node, runtimePin.nodeVersion, 'the executable has an unexpected Node runtime');
    assert.equal(runtime.electron, runtimePin.electronVersion, 'the executable has an unexpected Electron runtime');
    await installDialogs(application);
    await application.context().route(/^https?:\/\//, route => {
      requests.push(route.request().url());
      return route.abort('internetdisconnected');
    });
    const page = await application.firstWindow();
    page.setDefaultTimeout(30000);
    page.on('pageerror', error => pageErrors.push(error.message));
    page.on('console', message => {
      if (message.type() === 'error' && /content security policy|refused to execute|refused to load/i.test(message.text())) {
        policyErrors.push(message.text());
      }
    });
    await page.locator('#select-file').waitFor({state: 'visible'});
    assert.equal(page.url(), 'pptx-app://desktop/index.html');
    const renderer = await page.evaluate(() => ({
      require: typeof window.require, process: typeof window.process, buffer: typeof window.Buffer,
      bridge: typeof window.pptxDesktop, genericIpc: typeof window.pptxDesktop?.send
    }));
    assert.deepEqual(renderer, {require: 'undefined', process: 'undefined', buffer: 'undefined',
      bridge: 'object', genericIpc: 'undefined'}, 'renderer exposes Node or generic IPC');
    const info = await page.evaluate(() => window.pptxDesktop.getInfo());
    assert.equal(info.version, version, 'packaged desktop metadata uses a different version');
    assert.equal(info.engineVersion, engineVersion, 'packaged desktop metadata uses a different engine');

    const labels = await page.evaluate(() => ['slides', 'include-hidden', 'title', 'password'].map(id => {
      const field = document.getElementById(id);
      return {id, labeled: !!(field && (field.labels?.length || field.getAttribute('aria-label') ||
        field.getAttribute('aria-labelledby')))};
    }));
    assert.ok(labels.every(item => item.labeled), 'desktop fields need accessible labels: ' + JSON.stringify(labels));
    for (const id of ['select-file', 'convert', 'cancel', 'preview', 'save-html']) {
      const control = page.locator('#' + id);
      const name = await control.getAttribute('aria-label') || await control.textContent();
      assert.ok(name?.trim(), id + ' has no accessible name');
    }

    async function dialogState(values) {
      return application.evaluate((_, values) => Object.assign(globalThis.__pptxDesktopTestDialogs, values), values);
    }
    async function choose(filename) {
      const before = await dialogState({input: filename, cancelOpen: false});
      await page.locator('#select-file').focus();
      await page.locator('#select-file').press('Enter');
      await page.waitForFunction(() => !document.getElementById('convert').disabled);
      const after = await dialogState({});
      assert.equal(after.openCalls, before.openCalls + 1, 'keyboard activation did not open the native file picker');
    }
    async function convertUi({slides = '', includeHidden = false, title = '', password = ''} = {}) {
      await page.locator('#slides').fill(slides);
      await page.locator('#include-hidden').setChecked(includeHidden);
      await page.locator('#title').fill(title);
      if (!(await page.locator('#password-options').evaluate(element => element.open))) {
        await page.locator('#password-options > summary').click();
      }
      await page.locator('#password').fill(password);
      await page.locator('#convert').click();
      await page.waitForFunction(() => {
        const cancel = document.getElementById('cancel');
        return !document.getElementById('convert').disabled &&
          (cancel.hidden || getComputedStyle(cancel).display === 'none');
      }, null, {timeout});
      assert.equal(await page.locator('#password').inputValue(), '', 'password input was retained after conversion');
    }
    async function save(filename, expectedSlides, {chosenPath = filename, overwriteResponse = 0} = {}) {
      await page.waitForFunction(() => !document.getElementById('save-html').disabled);
      const before = await dialogState({output: chosenPath, cancelSave: false, overwriteResponse});
      await page.locator('#save-html').click();
      await page.waitForFunction(() => !document.getElementById('save-html').disabled);
      const after = await dialogState({});
      assert.equal(after.saveCalls, before.saveCalls + 1, 'save did not use the native save dialog');
      // The button can re-enable before a slow filesystem is observable by the
      // test process. Poll the actual chosen destination, without another save.
      const deadline = Date.now() + 10000;
      while (!fs.existsSync(filename) && Date.now() < deadline) await new Promise(resolve => setTimeout(resolve, 50));
      assert.ok(fs.existsSync(filename), 'the chosen HTML file was not written');
      return inspectHtml(filename, expectedSlides);
    }
    async function preview(expectedSlides) {
      await page.locator('#preview').click();
      const frame = page.frameLocator('#preview-frame');
      await frame.locator('#ptv-select option[value="' + (expectedSlides - 1) + '"]')
        .waitFor({state: 'attached'});
      assert.equal(await frame.locator('#ptv-select option').count(), expectedSlides);
      assert.equal(await page.locator('#preview-frame').getAttribute('sandbox'), 'allow-scripts');
      const iframe = await page.locator('#preview-frame').elementHandle();
      const content = await iframe.contentFrame();
      assert.ok(content && /^pptx-app:\/\/preview\/[^/]+\/index\.html$/.test(content.url()),
        'preview is not served from its dedicated application URL');
      const exposure = await content.evaluate(() => {
        let parentBlocked = false;
        try { void parent.pptxDesktop; } catch { parentBlocked = true; }
        return {node: typeof window.require, bridge: typeof window.pptxDesktop, parentBlocked};
      });
      assert.deepEqual(exposure, {node: 'undefined', bridge: 'undefined', parentBlocked: true},
        'document preview can reach privileged APIs');
      return frame;
    }

    await choose(source);
    await convertUi({title: '로컬 EXE 변환 검증'});
    let frame = await preview(3);
    assert.match(await frame.locator('.slide').first().textContent(), /TABLE_FILE_TWO/);
    await frame.locator('#ptv-mode').click();
    await frame.locator('#ptv-next').click();
    assert.equal(await frame.locator('#ptv-select').inputValue(), '1', 'preview slide navigation failed');
    const normal = await save(output('normal.html'), 3);
    assert.equal(normal.title, '로컬 EXE 변환 검증');
    assert.match(normal.texts[0], /TABLE_FILE_TWO/);
    assert.match(normal.texts[1], /INTRO_FILE_ONE/);
    assert.match(normal.texts[2], /IMAGE_CHART_FILE_THREE/);
    const movedSource = source + '.moved';
    fs.renameSync(source, movedSource);
    try {
      const savedAfterMove = await save(output('source-moved.html'), 3);
      assert.equal(savedAfterMove.source, normal.source, 'moving the source prevents saving a completed result');
    } finally { fs.renameSync(movedSource, source); }
    if (process.env.PPTX_DESKTOP_SCREENSHOT) {
      await page.screenshot({path: process.env.PPTX_DESKTOP_SCREENSHOT, fullPage: true});
    }

    // The native dialog chose a name without a suffix. Adding .html afterward
    // must not overwrite an existing different path without confirmation.
    const overwriteBase = output('기존 문서 덮어쓰기');
    const overwriteTarget = overwriteBase + '.html';
    const existingContents = 'EXISTING_HTML_MUST_SURVIVE_CANCEL';
    fs.writeFileSync(overwriteTarget, existingContents, 'utf8');
    const beforeOverwrite = await dialogState({output: overwriteBase, cancelSave: false, overwriteResponse: 0});
    await page.locator('#save-html').click();
    await page.waitForFunction(() => !document.getElementById('save-html').disabled);
    const cancelledOverwrite = await dialogState({});
    assert.equal(cancelledOverwrite.saveCalls, beforeOverwrite.saveCalls + 1);
    assert.equal(cancelledOverwrite.messageCalls, beforeOverwrite.messageCalls + 1,
      'adding .html bypassed confirmation for an existing destination');
    assert.deepEqual(cancelledOverwrite.lastMessage.buttons, ['취소', '덮어쓰기']);
    assert.equal(fs.readFileSync(overwriteTarget, 'utf8'), existingContents,
      'cancelling overwrite changed the existing HTML file');
    assert.equal(fs.existsSync(overwriteBase), false, 'save wrote to the unnormalized destination');
    const overwritten = await save(overwriteTarget, 3, {chosenPath: overwriteBase, overwriteResponse: 1});
    const confirmedOverwrite = await dialogState({});
    assert.equal(confirmedOverwrite.messageCalls, cancelledOverwrite.messageCalls + 1,
      'retrying the same destination skipped overwrite confirmation');
    assert.equal(overwritten.source, normal.source, 'approved overwrite did not save the converted HTML');

    await convertUi({slides: '1,4', includeHidden: true});
    await preview(2);
    const selected = await save(output('selected.html'), 2);
    assert.match(selected.texts[0], /TABLE_FILE_TWO/);
    assert.match(selected.texts[1], /HIDDEN_FILE_FOUR/);

    await choose(fixture('unknown.pptx'));
    const unknown = await page.evaluate(() => window.pptxDesktop.convert({slides: '', includeHidden: false, title: '', password: ''}));
    assert.equal(unknown.ok, false);
    assert.equal(unknown.code, 'UNKNOWN_CONTAINER');
    await convertUi();
    const unknownMessage = await page.locator('#status').textContent();
    assert.match(unknownMessage, /형식|구조|인식/);
    assert.ok(!/NASCA(?:\+SD)?(?:로|에 의해) 보호|조직의 DRM·IRM으로 보호된 문서입니다/.test(unknownMessage),
      'unrecognized input was incorrectly identified as company DRM');

    await choose(fixture('protected-standard.pptx'));
    for (const [password, code] of [['', 'PASSWORD_REQUIRED'], ['incorrect-desktop-password', 'INCORRECT_PASSWORD']]) {
      const result = await page.evaluate(options => window.pptxDesktop.convert(options),
        {slides: '', includeHidden: false, title: '', password});
      assert.equal(result.ok, false);
      assert.equal(result.code, code);
      if (password) assert.ok(!JSON.stringify(result).includes(password), 'password leaked into an error response');
    }
    await convertUi({password: PASSWORD});
    await preview(3);
    const protectedOutput = await save(output('protected.html'), 3);
    assert.match(protectedOutput.texts[0], /TABLE_FILE_TWO/);
    assert.ok(!(await page.locator('#status').textContent()).includes(PASSWORD), 'password leaked into status text');

    await choose(source);
    const cancelled = await page.evaluate(async () => {
      const pending = window.pptxDesktop.convert({slides: '', includeHidden: false, title: '', password: ''});
      const cancel = await window.pptxDesktop.cancel();
      return {cancel, result: await pending};
    });
    assert.equal(cancelled.cancel.ok, true);
    assert.equal(cancelled.result.ok, false);
    assert.equal(cancelled.result.code, 'CANCELLED', 'immediate cancellation did not stop the active conversion');
    await convertUi({slides: '2'});
    const retried = await save(output('retry.html'), 1);
    assert.match(retried.texts[0], /INTRO_FILE_ONE/);

    await choose(fixture('hostile.pptx'));
    await convertUi({title: '<img src=x onerror="window.__desktopTitleExecuted=1">'});
    frame = await preview(1);
    const hostile = await save(output('hostile.html'), 1);
    assert.ok(!hostile.source.includes('https://example.invalid/pptx-leak.png'));
    assert.equal(await page.evaluate(() => window.__desktopTitleExecuted), undefined);
    assert.equal(await frame.locator('body').evaluate(() => window.__pptxExecuted), undefined);

    assert.equal(sha(source), sourceHash, 'desktop conversion modified the original PPTX');
    assert.deepEqual(requests, [], 'desktop conversion attempted an HTTP(S) request');
    assert.deepEqual(pageErrors, [], 'desktop renderer had an uncaught error');
    assert.deepEqual(policyErrors, [], 'desktop or preview scripts were blocked by CSP');
    passed = true;
    console.log('PASS packaged desktop: isolated runtime, native file/save dialogs, overwrite cancellation/approval, real conversion, range/hidden slides, protected Office input, cancellation/retry, preview navigation and isolation, zero HTTP(S)');
    console.log('Runtime: Electron ' + runtime.electron + ', Node ' + runtime.node + ', app ' + runtime.version);
  } finally {
    await application.close();
    if (passed) fs.rmSync(temporary, {recursive: true, force: true, maxRetries: 10, retryDelay: 100});
    else console.error('Desktop test evidence retained at ' + temporary);
  }
}

main().catch(error => { console.error(error); process.exitCode = 1; });
