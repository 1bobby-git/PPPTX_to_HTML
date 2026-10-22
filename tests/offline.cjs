#!/usr/bin/env node
'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
const { chromium } = require('playwright');
const JSZip = require('jszip');

const root = path.resolve(__dirname, '..');
const index = path.join(root, 'index.html');
const fixture = name => path.join(__dirname, 'fixtures', name);
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));

function browserExecutable() {
  const explicit = process.env.CHROME_BIN || process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH;
  const candidates = [explicit, chromium.executablePath(),
    '/tmp/pptx-browsers/chrome-linux64/chrome',
    '/tmp/pptx-browsers/chromium-1234/chrome-linux64/chrome',
    '/tmp/pptx-browsers/chromium_headless_shell-1234/chrome-headless-shell-linux64/chrome-headless-shell',
    '/usr/bin/chromium', '/usr/bin/google-chrome'];
  return candidates.find(p => p && fs.existsSync(p));
}

async function fixtureContract() {
  const zip = await JSZip.loadAsync(fs.readFileSync(fixture('complex.pptx')));
  const presentation = await zip.file('ppt/presentation.xml').async('string');
  const relationships = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
  const ids = [...presentation.matchAll(/<p:sldId\b[^>]*r:id="([^"]+)"[^>]*\/>/g)].map(x => x[1]);
  const targets = new Map([...relationships.matchAll(/<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*\/>/g)].map(x => [x[1], x[2]]));
  assert.equal(ids.length, 4, 'fixture must contain four slides');
  assert.match(targets.get(ids[0]), /slide2\.xml$/, 'presentation order should differ from slide filename order');
  assert.match(targets.get(ids[1]), /slide1\.xml$/);
  assert.match(await zip.file('ppt/slides/slide4.xml').async('string'), /<p:sld\b[^>]*show="0"/);
  assert.ok(Object.keys(zip.files).some(name => /^ppt\/media\/[^/]+\.png$/.test(name)), 'image data URI must be embedded');
  assert.ok(zip.file('ppt/charts/chart1.xml'), 'chart must be embedded');
  assert.match(await zip.file('ppt/slides/_rels/slide3.xml.rels').async('string'),
    /Target="\/ppt\/charts\/chart1\.xml"/, 'fixture must preserve the legal absolute OPC chart relationship');
  assert.ok(Object.keys(zip.files).some(name => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(name)), 'speaker notes must be present in source');
  const hostile = await JSZip.loadAsync(fs.readFileSync(fixture('hostile.pptx')));
  assert.match(await hostile.file('ppt/slides/_rels/slide1.xml.rels').async('string'), /TargetMode="External"/);
  const pie = await JSZip.loadAsync(fs.readFileSync(fixture('basic-chart.pptx')));
  assert.ok(Object.keys(pie.files).some(name => /^ppt\/charts\/chart\d+\.xml$/.test(name)),
    'direct chart fixture must include chart XML');
}

async function main() {
  await fixtureContract();
  assert.ok(fs.existsSync(index), 'build index.html before running browser tests');
  const executablePath = browserExecutable();
  assert.ok(executablePath, 'Chromium is unavailable. Set CHROME_BIN to a compatible Chrome executable.');
  const browser = await chromium.launch({ headless: true, executablePath, args: ['--no-sandbox'] });
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'pptx-offline-test-'));
  const failedRequests = [];
  try {
    const context = await browser.newContext({ acceptDownloads: true, viewport: { width: 1280, height: 850 } });
    await context.route(/^https?:\/\//, route => {
      failedRequests.push(route.request().url());
      return route.abort('internetdisconnected');
    });
    const page = await context.newPage();
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    page.on('console', message => { if (message.type() === 'error') errors.push('console: ' + message.text()); });
    await page.goto(pathToFileURL(index).href);
    async function waitForConversion() {
      await page.waitForFunction(() => !document.querySelector('#download-button').disabled ||
        document.querySelector('#status-line').dataset.tone === 'error', null, { timeout: 30000 });
      const status = await page.locator('#status-text').innerText();
      assert.equal(await page.locator('#download-button').isEnabled(), true,
        `conversion failed: ${status}; browser errors: ${errors.join(' | ')}`);
    }
    assert.equal(new URL(page.url()).protocol, 'file:');
    assert.equal(await page.locator('#convert-button').isDisabled(), true);

    // A valid presentation must render in a blob iframe without an HTTP server.
    await page.locator('#file-input').setInputFiles(fixture('complex.pptx'));
    assert.equal(await page.locator('#convert-button').isEnabled(), true);
    await page.locator('#output-title').fill('오프라인 결과');
    await page.locator('#convert-button').click();
    await waitForConversion();
    await page.locator('#preview-frame').waitFor({ state: 'visible' });
    const frame = page.frameLocator('#preview-frame');
    await frame.locator('.slide').first().waitFor({ timeout: 30000 });
    assert.equal(await frame.locator('.slide').count(), 3, 'hidden slide excluded by default');
    assert.equal(await frame.locator('#ptv-select').inputValue(), '0', 'visible first slide should remain selected');
    const text = await frame.locator('body').innerText();
    assert.ok(text.includes('TABLE_FILE_TWO'));
    assert.ok(text.includes('안녕하세요 오프라인'));
    assert.ok(text.includes('IMAGE_CHART_FILE_THREE'));
    assert.ok(!text.includes('HIDDEN_FILE_FOUR'));
    assert.ok(!text.includes('PRIVATE_SPEAKER_NOTE_MUST_NOT_EXPORT'));
    assert.ok((await frame.locator('.slide').first().innerText()).includes('TABLE_FILE_TWO'), 'render should honor presentation.xml order');
    assert.ok(await frame.locator('table').count() > 0, 'table should retain cells');
    assert.ok(await frame.locator('.chart-placeholder').count() > 0, 'unsupported bar chart must show a fallback');
    assert.equal(await frame.locator('.chart-direct').count(), 0, 'unsupported bar chart must not claim direct rendering');
    assert.ok((await page.locator('#warning-list li').allTextContents()).some(message => message.includes('차트')),
      'bar fallback must have a chart warning');
    assert.ok(await frame.locator('img').count() > 0, 'embedded image should render');
    assert.ok(await frame.locator('svg').count() > 0, 'chart or vector shape should render as SVG');
    assert.ok(await frame.locator('img').evaluateAll(images => images.some(img => img.complete && img.naturalWidth > 0)),
      'an embedded image must decode in preview');
    assert.match(await page.locator('#result-subtitle').innerText(), /3/);
    let previewFit = false;
    for (let attempt = 0; attempt < 80; attempt++) {
      const [outer, inner, parentWidth, frameWidth] = await Promise.all([
        page.locator('#preview-frame').boundingBox(), frame.locator('.slide-shell').first().boundingBox(),
        page.locator('#preview-frame').evaluate(element => element.parentElement.clientWidth),
        frame.locator('body').evaluate(() => innerWidth),
      ]);
      previewFit = !!outer && !!inner && outer.width >= parentWidth * 0.9 &&
        frameWidth >= outer.width - 4 && inner.width >= outer.width * 0.8;
      if (previewFit) break;
      await delay(50);
    }
    assert.ok(previewFit, 'preview slide should fill at least 80% of iframe width after viewer layout');
    await page.screenshot({ path: path.join(tmp, 'desktop.png'), fullPage: true });

    const downloadPromise = page.waitForEvent('download');
    await page.locator('#download-button').click();
    const download = await downloadPromise;
    assert.equal(download.suggestedFilename(), '오프라인 결과.html');
    const exported = path.join(tmp, 'exported.html');
    await download.saveAs(exported);
    const exportedHtml = fs.readFileSync(exported, 'utf8');
    assert.match(exportedHtml, /TABLE_FILE_TWO/);
    assert.ok(!exportedHtml.includes('PRIVATE_SPEAKER_NOTE_MUST_NOT_EXPORT'), 'speaker notes must not leak into standalone HTML');
    const standalone = await context.newPage();
    await standalone.goto(pathToFileURL(exported).href);
    await standalone.locator('.slide').first().waitFor();
    assert.equal(await standalone.locator('.slide').count(), 3);
    assert.ok((await standalone.locator('body').innerText()).includes('안녕하세요 오프라인'));
    assert.ok(await standalone.locator('table').count() > 0);
    assert.equal(await standalone.locator('#ptv-select option').count(), 3, 'standalone viewer script must initialize');
    assert.ok(await standalone.locator('img').evaluateAll(images => images.some(img => img.complete && img.naturalWidth > 0)),
      'an embedded image must decode after export');
    assert.ok(await standalone.locator('svg').count() > 0);
    await standalone.locator('#ptv-mode').click();
    assert.equal(await standalone.locator('#ptv-mode').getAttribute('aria-pressed'), 'true');
    assert.equal(await standalone.locator('#ptv-slides > .slide-shell:visible').count(), 1);
    assert.equal(await standalone.locator('#ptv-prev').isDisabled(), true);
    await standalone.locator('#ptv-next').click();
    assert.equal(await standalone.locator('#ptv-select').inputValue(), '1');
    await standalone.locator('#ptv-select').selectOption('2');
    assert.equal(await standalone.locator('#ptv-next').isDisabled(), true);
    await standalone.screenshot({ path: path.join(tmp, 'standalone-desktop.png'), fullPage: true });
    await standalone.setViewportSize({ width: 390, height: 844 });
    await standalone.screenshot({ path: path.join(tmp, 'standalone-mobile.png'), fullPage: true });
    const standaloneWidth = await standalone.evaluate(() => ({ scroll: document.documentElement.scrollWidth, viewport: innerWidth }));
    assert.ok(standaloneWidth.scroll <= standaloneWidth.viewport + 1, `standalone mobile overflow: ${JSON.stringify(standaloneWidth)}`);
    await standalone.close();

    // A valid simple pie chart takes the direct renderer path. This is a
    // separate fixture so the preceding bar fallback remains representative.
    await page.locator('#file-input').setInputFiles(fixture('basic-chart.pptx'));
    await page.locator('#convert-button').click();
    await waitForConversion();
    const directChart = page.frameLocator('#preview-frame');
    await directChart.locator('.chart-direct svg').waitFor();
    assert.ok(await directChart.locator('.chart-direct svg path').count() >= 3, 'pie segments should be SVG paths');
    const pieText = await directChart.locator('.chart-direct').innerText();
    for (const label of ['Mon', 'Tue', 'Wed', '22%', '44%', '33%']) assert.ok(pieText.includes(label), `missing chart value: ${label}`);
    const pieDownloadPromise = page.waitForEvent('download');
    await page.locator('#download-button').click();
    const pieDownload = await pieDownloadPromise;
    const pieExport = path.join(tmp, 'chart.html');
    await pieDownload.saveAs(pieExport);
    const piePage = await context.newPage();
    await piePage.goto(pathToFileURL(pieExport).href);
    await piePage.locator('.chart-direct svg').waitFor();
    assert.ok(await piePage.locator('.chart-direct svg path').count() >= 3);
    await piePage.screenshot({ path: path.join(tmp, 'chart.png'), fullPage: true });
    await piePage.close();

    await page.locator('#file-input').setInputFiles(fixture('complex.pptx'));

    // Setting changes invalidate the previous result and filter by displayed order.
    await page.locator('#slide-range').fill('2');
    assert.equal(await page.locator('#download-button').isDisabled(), true);
    await page.locator('#convert-button').click();
    await waitForConversion();
    await frame.locator('.slide').first().waitFor();
    assert.equal(await frame.locator('.slide').count(), 1);
    assert.ok((await frame.locator('.slide').first().innerText()).includes('INTRO_FILE_ONE'));

    await page.locator('#slide-range').fill('4');
    await page.locator('#include-hidden').check();
    await page.locator('#convert-button').click();
    await waitForConversion();
    await frame.locator('.slide').first().waitFor();
    assert.equal(await frame.locator('.slide').count(), 1);
    assert.ok((await frame.locator('body').innerText()).includes('HIDDEN_FILE_FOUR'));

    // Selecting a new file must discard the previous blob and download action.
    await page.locator('#file-input').setInputFiles(fixture('invalid.pptx'));
    assert.equal(await page.locator('#download-button').isDisabled(), true);
    assert.equal(await page.locator('#preview-frame').isHidden(), true);
    await page.locator('#slide-range').fill('');
    await page.locator('#convert-button').click();
    await page.waitForFunction(() => document.querySelector('#status-line').dataset.tone === 'error', null, { timeout: 30000 });
    assert.match(await page.locator('#status-text').innerText(), /실패|오류|파일|읽/);
    assert.equal(await page.locator('#download-button').isDisabled(), true);

    // Malicious-looking slide text must stay text; external image references
    // must not cause an HTTP fetch even when the network is intentionally cut.
    await page.locator('#file-input').setInputFiles(fixture('hostile.pptx'));
    await page.locator('#convert-button').click();
    await page.waitForFunction(() => !document.querySelector('#download-button').disabled || document.querySelector('#status-line').dataset.tone === 'error', null, { timeout: 120000 });
    if (await page.locator('#download-button').isEnabled()) {
      const hostileFrame = page.frameLocator('#preview-frame');
      assert.ok((await hostileFrame.locator('body').innerText()).includes('<script>'));
      assert.equal(await hostileFrame.locator('body').evaluate(() => window.__pptxExecuted || 0), 0);
    }
    assert.deepEqual(failedRequests, [], 'no HTTP(S) request may occur from app or preview');
    assert.deepEqual(errors, [], 'app should have no uncaught page errors');

    // A deterministic delayed Worker proves that cancellation ignores even a
    // result posted after terminate(), instead of racing a fast real deck.
    await page.evaluate(() => {
      window.Worker = class DelayedWorker {
        postMessage(message) {
          setTimeout(() => this.onmessage?.({ data: {
            id: message.id, type: 'result', html: '<!doctype html><title>stale</title><div class="slide">STALE</div>', slideCount: 1,
          } }), 250);
        }
        terminate() {}
      };
    });
    await page.locator('#file-input').setInputFiles(fixture('complex.pptx'));
    await page.locator('#slide-range').fill('');
    await page.locator('#convert-button').click();
    await page.locator('#cancel-button').click();
    assert.match(await page.locator('#status-text').innerText(), /취소/);
    await delay(750);
    assert.equal(await page.locator('#download-button').isDisabled(), true);

    await page.setViewportSize({ width: 390, height: 844 });
    await page.screenshot({ path: path.join(tmp, 'mobile.png'), fullPage: true });
    const dimensions = await page.evaluate(() => ({ scroll: document.documentElement.scrollWidth, viewport: innerWidth }));
    assert.ok(dimensions.scroll <= dimensions.viewport + 1, `mobile horizontal overflow: ${JSON.stringify(dimensions)}`);
    await page.setViewportSize({ width: 320, height: 700 });
    const narrowDimensions = await page.evaluate(() => ({ scroll: document.documentElement.scrollWidth, viewport: innerWidth }));
    assert.ok(narrowDimensions.scroll <= narrowDimensions.viewport + 1, `320px horizontal overflow: ${JSON.stringify(narrowDimensions)}`);
    console.log(`PASS file:// browser flow; browser=${executablePath}; screenshots=${tmp}`);
    await context.close();
  } finally {
    await browser.close();
  }
}

main().catch(error => { console.error(error); process.exitCode = 1; });
