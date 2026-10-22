#!/usr/bin/env node
'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
const { chromium } = require('playwright');
const JSZip = require('jszip');
const { generateFixtures, PASSWORD } = require('../scripts/create-protected-fixtures.cjs');

const root = path.resolve(__dirname, '..');
const fixture = name => path.join(__dirname, 'fixtures', name);
const reference = name => path.join(__dirname, 'reference', name);
const bytes = filename => new Uint8Array(fs.readFileSync(filename));

function browserExecutable() {
  const explicit = process.env.CHROME_BIN || process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH;
  return [explicit, chromium.executablePath(),
    '/tmp/pptx-browsers/chrome-linux64/chrome',
    '/tmp/pptx-browsers/chromium-1234/chrome-linux64/chrome',
    '/tmp/pptx-browsers/chromium_headless_shell-1234/chrome-headless-shell-linux64/chrome-headless-shell',
    '/usr/bin/chromium', '/usr/bin/google-chrome',
  ].find(filename => filename && fs.existsSync(filename));
}

async function moduleContract() {
  const { prepare } = require('../src/input-package.js');
  const plain = fs.readFileSync(reference('complex-plain.pptx'));
  const encrypted = bytes(reference('complex-agile.pptx'));

  async function rejectsCode(input, password, expected) {
    const allowed = Array.isArray(expected) ? expected : [expected];
    await assert.rejects(() => prepare(input, password), error => {
      assert.ok(allowed.includes(error.code), `expected ${allowed.join('/')}, got ${error.code}: ${error.message}`);
      assert.match(error.message, /[가-힣]/, 'format failures need a useful Korean message');
      assert.ok(!JSON.stringify({ message: error.message, details: error.details }).includes(PASSWORD),
        'error diagnostics must not expose the password');
      return true;
    });
  }

  const direct = await prepare(new Uint8Array(plain), '');
  assert.equal(direct.format, 'zip');
  assert.deepEqual(Buffer.from(direct.bytes), plain, 'ordinary ZIP bytes must pass unchanged');
  const unrelated = bytes(fixture('unknown.pptx'));
  const unknown = await prepare(unrelated, '');
  assert.equal(unknown.format, 'unknown', 'unrelated bytes cannot prove encryption or a renamed extension');
  assert.deepEqual(unknown.bytes, unrelated, 'the ZIP parser receives original unknown bytes');
  await rejectsCode(encrypted, '', 'PASSWORD_REQUIRED');
  await rejectsCode(encrypted, 'incorrect-password', 'INCORRECT_PASSWORD');
  await rejectsCode(encrypted, 'x'.repeat(256), 'INPUT_LIMIT');

  const agile = await prepare(encrypted, PASSWORD);
  assert.equal(agile.format, 'password-agile');
  assert.deepEqual(Buffer.from(agile.bytes), plain, 'independent Agile reference must decrypt byte-for-byte');
  const renamed = await prepare(bytes(fixture('agile-renamed-prefix.pptx')), PASSWORD);
  assert.deepEqual(Buffer.from(renamed.bytes), plain, 'XML namespace prefix spelling must not change decryption');
  const jsAgile = await prepare(bytes(fixture('protected-agile-js.pptx')), PASSWORD);
  assert.deepEqual(Buffer.from(jsAgile.bytes), fs.readFileSync(fixture('complex.pptx')),
    'JS-generated Agile remains compatible with the independent reference path');

  const standard = await prepare(bytes(fixture('protected-standard.pptx')), PASSWORD);
  assert.equal(standard.format, 'password-standard');
  assert.deepEqual(Buffer.from(standard.bytes), fs.readFileSync(fixture('complex.pptx')), 'Standard decrypted package bytes');
  await rejectsCode(bytes(fixture('protected-standard.pptx')), 'incorrect-password', 'INCORRECT_PASSWORD');

  await rejectsCode(bytes(fixture('tampered-package.pptx')), PASSWORD, 'INVALID_CONTAINER');
  await rejectsCode(bytes(fixture('tampered-hmac.pptx')), PASSWORD, 'INVALID_CONTAINER');
  await rejectsCode(bytes(fixture('declared-too-large.pptx')), '', 'INPUT_LIMIT');
  await rejectsCode(bytes(fixture('unsupported-cipher.pptx')), PASSWORD, 'UNSUPPORTED_ENCRYPTION');
  await rejectsCode(bytes(fixture('excessive-iterations.pptx')), PASSWORD, 'UNSUPPORTED_ENCRYPTION');
  await rejectsCode(bytes(fixture('drm-marker.pptx')), '', 'DRM_PROTECTED');
  await rejectsCode(bytes(fixture('drm-content-marker.pptx')), '', 'DRM_PROTECTED');
  await rejectsCode(bytes(fixture('legacy-marker.pptx')), '', 'LEGACY_PPT');
  await rejectsCode(bytes(fixture('unknown-cfb.pptx')), '', 'UNKNOWN_CONTAINER');
  await rejectsCode(bytes(fixture('malformed-cfb.pptx')), '', 'INVALID_CONTAINER');
  await rejectsCode(bytes(fixture('incomplete-encryption.pptx')), '', 'INVALID_CONTAINER');

  const prefix = fs.readFileSync(fixture('prefixed.pptx'));
  const preparedPrefix = await prepare(new Uint8Array(prefix), '');
  assert.deepEqual(Buffer.from(preparedPrefix.bytes), prefix, 'do not strip guessed ZIP headers');
  assert.ok((await JSZip.loadAsync(preparedPrefix.bytes)).file('ppt/presentation.xml'));
  console.log('PASS input package contract; independent Agile bytes + HMAC; Standard; namespaces; container classification');
}

async function browserContract({ noPowerPointUIOnly = false } = {}) {
  const executablePath = browserExecutable();
  assert.ok(executablePath, 'Chromium is unavailable; set CHROME_BIN to a compatible Chrome executable.');
  const browser = await chromium.launch({ headless: true, executablePath, args: ['--no-sandbox'] });
  const temporary = fs.mkdtempSync(path.join(os.tmpdir(), 'pptx-protection-test-'));
  const failedRequests = [];
  const errors = [];
  try {
    // file:// pages cannot register service workers. Playwright's explicit
    // serviceWorkers:'block' injects a navigator getter into sandboxed preview
    // frames and itself raises a SecurityError, so network isolation uses routes.
    const context = await browser.newContext({ acceptDownloads: true, viewport: { width: 1280, height: 850 } });
    await context.route(/^https?:\/\//, route => {
      failedRequests.push(route.request().url());
      return route.abort('internetdisconnected');
    });
    const page = await context.newPage();
    page.on('pageerror', error => errors.push(error.message));
    page.on('console', message => { if (message.type() === 'error') errors.push('console: ' + message.text()); });
    await page.goto(pathToFileURL(path.join(root, 'index.html')).href);
    assert.equal(new URL(page.url()).protocol, 'file:');
    assert.equal(await page.locator('#document-password').getAttribute('type'), 'password');
    await page.locator('#protection-options').evaluate(element => { element.open = true; });

    async function expectNoPowerPointIntegration() {
      assert.equal(await page.locator('#powerpoint-helper-button, #powerpoint-launcher-button').count(), 0,
        'the browser workflow must not offer a PowerPoint helper or launcher');
      assert.equal(await page.locator('#powerpoint-helper-source, #powerpoint-launcher-source').count(), 0,
        'the single-file app must not embed PowerPoint helper programs');
      assert.doesNotMatch(fs.readFileSync(path.join(root, 'index.html'), 'utf8'),
        /Convert-PptxWithPowerPoint|powerpoint-(?:helper|launcher)-(?:button|source)|downloadHelper/,
        'legacy PowerPoint filenames, embedded sources and download handlers must be absent from the app');
    }
    await expectNoPowerPointIntegration();

    async function select(filename) {
      await page.locator('#file-input').setInputFiles(filename);
      assert.equal(await page.locator('#download-button').isDisabled(), true, 'selection must invalidate previous results');
      assert.equal(await page.locator('#document-password').inputValue(), '', 'selection must clear the previous password');
      assert.equal(await page.locator('#convert-button').isEnabled(), true);
      await page.locator('#output-title').fill('보호 문서 오프라인 회귀');
    }

    async function convert(password) {
      if (password !== undefined) {
        await page.locator('#protection-options').evaluate(element => { element.open = true; });
        await page.locator('#document-password').fill(password);
      }
      await page.locator('#convert-button').click();
      await page.waitForFunction(() => !document.querySelector('#download-button').disabled ||
        document.querySelector('#status-line').dataset.tone === 'error', null, { timeout: 120000 });
    }

    async function expectError(codes) {
      const allowed = Array.isArray(codes) ? codes : [codes];
      assert.equal(await page.locator('#download-button').isDisabled(), true);
      const code = await page.locator('#status-line').getAttribute('data-error-code');
      const status = await page.locator('#status-text').innerText();
      assert.ok(allowed.includes(code), `expected ${allowed.join('/')}, got ${code}: ${status}`);
      assert.match(status, /[가-힣]/);
      assert.ok(!status.includes('확장자만 바꾼'), 'an unknown container is not evidence that the extension was renamed');
      const diagnostic = await page.locator('#diagnostics-text').textContent();
      assert.ok(diagnostic && diagnostic.includes(code), 'failure details should identify the actual error code');
      assert.ok(!diagnostic.includes(PASSWORD), 'diagnostics must not expose the password');
      assert.equal(await page.locator('#preview-frame').isHidden(), true, 'a failed conversion must not retain a stale preview');
    }

    async function expectDeck() {
      const status = await page.locator('#status-text').innerText();
      assert.equal(await page.locator('#download-button').isEnabled(), true,
        `conversion failed: ${status}; browser errors: ${errors.join(' | ')}`);
      assert.equal(await page.locator('#document-password').inputValue(), '', 'successful conversion must clear the password');
      const frame = page.frameLocator('#preview-frame');
      await frame.locator('.slide').first().waitFor({ timeout: 30000 });
      assert.equal(await frame.locator('.slide').count(), 3, 'the encrypted reference must preserve 3 visible slides');
      const slideText = await frame.locator('.slide').allTextContents();
      assert.ok(slideText[0].includes('TABLE_FILE_TWO'), 'presentation ordering must survive decryption');
      assert.ok(slideText[1].includes('INTRO_FILE_ONE'));
      assert.ok(slideText[1].includes('안녕하세요 오프라인'));
      assert.ok(slideText[2].includes('IMAGE_CHART_FILE_THREE'));
      const text = await frame.locator('body').innerText();
      assert.ok(!text.includes('HIDDEN_FILE_FOUR'));
      assert.ok(!text.includes('PRIVATE_SPEAKER_NOTE_MUST_NOT_EXPORT'));
      assert.ok(await frame.locator('img').evaluateAll(images => images.some(img => img.complete && img.naturalWidth > 0)),
        'the encrypted package image must decode');
      return {
        slideText,
        tableText: await frame.locator('table').allTextContents(),
        imageCount: await frame.locator('img').count(),
        svgCount: await frame.locator('svg').count(),
        chartFallbackCount: await frame.locator('.chart-placeholder').count(),
      };
    }

    async function capturePreview(filename) {
      await page.evaluate(() => scrollTo(0, 0));
      await page.locator('#preview-frame').scrollIntoViewIfNeeded();
      const frame = page.frameLocator('#preview-frame');
      let fit = false;
      for (let attempt = 0; attempt < 40; attempt++) {
        const outer = await page.locator('#preview-frame').boundingBox();
        const size = await frame.locator('body').evaluate(() => ({
          viewport: innerWidth, slide: document.querySelector('.slide-shell').getBoundingClientRect().width,
        }));
        if (outer && size.viewport >= outer.width - 4 && size.slide >= outer.width * 0.8) { fit = true; break; }
        await page.waitForTimeout(50);
      }
      assert.ok(fit, 'the preview must fill the iframe width after layout settles');
      // fullPage:true in this Chromium version temporarily resizes sandboxed
      // iframes while capturing. A viewport screenshot preserves actual layout.
      await page.screenshot({ path: path.join(temporary, filename) });
    }

    // Establish a baseline from the exact bytes encrypted by an independent
    // implementation, then compare slide contents, order, tables and media.
    await select(reference('complex-plain.pptx'));
    await convert();
    const baseline = await expectDeck();

    if (noPowerPointUIOnly) {
      await capturePreview('no-powerpoint-local-conversion.png');
      // These benign markers test the new stop-and-explain flow without
      // regenerating or rerunning the unchanged password encryption fixtures.
      const CFB = require('cfb');
      const cfb = CFB.utils.cfb_new();
      CFB.utils.cfb_add(cfb, '\u0006DataSpaces/DataSpaceInfo/DRMEncryptedDataSpace', Buffer.from('benign marker'));
      for (const [name, buffer, code] of [
        ['drm-marker.pptx', CFB.write(cfb, { type: 'buffer' }), 'DRM_PROTECTED'],
        ['unknown.pptx', Buffer.from('Benign unrelated bytes; no document format implied.'), 'UNKNOWN_CONTAINER'],
      ]) {
        await select({ name, mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', buffer });
        await convert();
        await expectError(code);
        const guidance = (await page.locator('#status-text').innerText()) + '\n' +
          (await page.locator('#protection-options').innerText());
        assert.doesNotMatch(guidance, /(?:변환|PowerPoint)\s*도우미|PowerPoint에서[^\n]*(?:저장|내보내|변환)|PowerPoint를\s*실행(?:해|하)|powershell|\.cmd/i,
          'unreadable protection must not trigger a PowerPoint-dependent fallback');
        assert.match(guidance, /승인|허용|SDK|공식\s*연동/, 'protected input must explain the approved-copy or official-integration requirement');
        await expectNoPowerPointIntegration();
      }
      assert.deepEqual(failedRequests, [], 'the no-PowerPoint file:// workflow must make zero HTTP(S) requests');
      assert.deepEqual(errors, [], 'the no-PowerPoint workflow must have no uncaught browser errors');
      console.log(`PASS no-PowerPoint UI contract; no helper UI/embedded programs; ordinary local conversion; DRM/unknown stop with guidance; zero HTTP(S); screenshots=${temporary}`);
      await context.close();
      return;
    }

    await page.locator('#protection-options').evaluate(element => { element.open = true; });
    await page.locator('#document-password').fill('must-clear-on-new-file');
    await select(reference('complex-agile.pptx'));
    await convert();
    await expectError('PASSWORD_REQUIRED');
    await convert('incorrect-password');
    await expectError('INCORRECT_PASSWORD');
    await convert(PASSWORD);
    assert.deepEqual(await expectDeck(), baseline, 'independently encrypted Agile deck must render as the exact plaintext deck');
    await capturePreview('encrypted-agile-success.png');

    const exportedPromise = page.waitForEvent('download');
    await page.locator('#download-button').click();
    const exportedDownload = await exportedPromise;
    const exportedFile = path.join(temporary, 'agile-export.html');
    await exportedDownload.saveAs(exportedFile);
    const exportedHtml = fs.readFileSync(exportedFile, 'utf8');
    assert.ok(!exportedHtml.includes(PASSWORD), 'the password must never be embedded in exported HTML');
    assert.ok(!exportedHtml.includes('incorrect-password'));
    assert.ok(!exportedHtml.includes('PRIVATE_SPEAKER_NOTE_MUST_NOT_EXPORT'));
    assert.match(exportedHtml, /TABLE_FILE_TWO/);
    const exportedPage = await context.newPage();
    exportedPage.on('pageerror', error => errors.push(error.message));
    await exportedPage.goto(pathToFileURL(exportedFile).href);
    await exportedPage.locator('.slide').first().waitFor();
    assert.equal(await exportedPage.locator('.slide').count(), 3);
    assert.deepEqual(await exportedPage.locator('.slide').allTextContents(), baseline.slideText);
    await exportedPage.close();

    await select(fixture('protected-standard.pptx'));
    await convert(PASSWORD);
    assert.deepEqual(await expectDeck(), baseline, 'Standard encrypted package must render the same slide contents');

    await select(fixture('prefixed.pptx'));
    await convert();
    assert.deepEqual(await expectDeck(), baseline, 'a JSZip-valid prefixed ZIP must convert without guessed header stripping');

    for (const [name, codes, password] of [
      ['unknown.pptx', 'UNKNOWN_CONTAINER'],
      ['unknown-cfb.pptx', 'UNKNOWN_CONTAINER'],
      ['broken-zip.pptx', 'INVALID_PACKAGE'],
      ['unrelated-zip.pptx', 'UNSUPPORTED_PACKAGE'],
      ['malformed-cfb.pptx', 'INVALID_CONTAINER'],
      ['drm-marker.pptx', 'DRM_PROTECTED'],
      ['legacy-marker.pptx', 'LEGACY_PPT'],
      ['tampered-package.pptx', 'INVALID_CONTAINER', PASSWORD],
    ]) {
      await select(fixture(name));
      await convert(password);
      await expectError(codes);
      if (name === 'drm-marker.pptx') {
        await page.locator('#diagnostics').evaluate(element => { element.open = true; });
        await page.locator('#diagnostics').scrollIntoViewIfNeeded();
        await page.screenshot({ path: path.join(temporary, 'drm-format-diagnostic.png') });
      }
    }

    await expectNoPowerPointIntegration();

    assert.deepEqual(failedRequests, [], 'the file:// app, decryption, previews, and export may not make any HTTP(S) requests');
    assert.deepEqual(errors, [], 'protection handling should have no uncaught browser errors');
    console.log(`PASS file:// protection flow; Agile + Standard + valid prefixed ZIP; truthful errors; password hygiene; no PowerPoint helper; browser=${executablePath}; screenshots=${temporary}`);
    await context.close();
  } finally {
    await browser.close();
  }
}

async function main() {
  const noPowerPointUIOnly = process.argv.includes('--no-powerpoint-ui-only');
  if (!noPowerPointUIOnly) {
    await generateFixtures();
    await moduleContract();
  }
  if (!process.argv.includes('--module-only')) await browserContract({ noPowerPointUIOnly });
}

main().catch(error => { console.error(error); process.exitCode = 1; });
