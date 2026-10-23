#!/usr/bin/env node
'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const crypto = require('node:crypto');
const {pathToFileURL} = require('node:url');
const {Worker} = require('node:worker_threads');
const {JSDOM} = require('jsdom');
const DOMPurify = require('dompurify');
const JSZip = require('jszip');
const PptxGenJS = require('pptxgenjs');
const {chromium, firefox, webkit} = require(process.env.PPTX_TEST_PLAYWRIGHT || 'playwright');

const root = path.resolve(__dirname, '..');
const browserName = process.env.PPTX_TEST_BROWSER || 'chromium';
const browserTypes = {chromium, firefox, webkit};
assert.ok(Object.hasOwn(browserTypes, browserName), 'PPTX_TEST_BROWSER must be chromium, firefox, or webkit.');
const browserType = browserTypes[browserName];
const actions = ['next', 'previous', 'first', 'last'];
const kinds = ['text', 'shape', 'image'];
const cases = [
  {name: 'visible', slides: '', includeHidden: false, originals: [1, 2, 4], parts: [3, 1, 2]},
  {name: 'all', slides: '', includeHidden: true, originals: [1, 2, 3, 4], parts: [3, 1, 4, 2]},
  {name: 'range', slides: '2,4', includeHidden: false, originals: [2, 4], parts: [1, 2]},
  {name: 'range-hidden', slides: '1,3', includeHidden: true, originals: [1, 3], parts: [3, 4]},
  {name: 'range-visible', slides: '1,3', includeHidden: false, originals: [1], parts: [3]}
];

function browserExecutable() {
  if (browserName !== 'chromium') {
    const explicit = process.env[browserName === 'firefox' ? 'FIREFOX_BIN' : 'WEBKIT_BIN'] ||
      process.env['PLAYWRIGHT_' + browserName.toUpperCase() + '_EXECUTABLE_PATH'];
    return [explicit, browserType.executablePath()].find(p => p && fs.existsSync(p));
  }
  return [process.env.CHROME_BIN || process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH,
    chromium.executablePath(), '/tmp/pptx-browsers/chrome-linux64/chrome',
    '/tmp/pptx-browsers/chromium-1234/chrome-linux64/chrome',
    '/tmp/pptx-browsers/chromium_headless_shell-1234/chrome-headless-shell-linux64/chrome-headless-shell',
    '/usr/bin/chromium', '/usr/bin/google-chrome'].find(p => p && fs.existsSync(p));
}

function linkSelector(kind, token) {
  return (kind === 'text' ? '.run' : '.shape-action-surface') + '[title="' + kind + '-' + token + '"]';
}

async function createFixture() {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  const png = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==';
  for (let part = 1; part <= 4; part++) {
    const slide = pptx.addSlide();
    slide.addText('PART_' + part, {x: 0.4, y: 0.1, w: 4, h: 0.5, fontSize: 22});
    slide.addNotes('PRIVATE_LINK_TEST_NOTE_MUST_NOT_EXPORT');
    for (const [row, token] of ['slide-1', 'slide-2', 'slide-3', 'slide-4', ...actions].entries()) {
      const target = token.startsWith('slide-') ? Number(token.slice(6)) : 2;
      const y = 0.9 + row * 0.7;
      slide.addText('text-' + token, {x: 0.4, y, w: 4, h: 0.45, fontSize: 18,
        hyperlink: {slide: target, tooltip: 'text-' + token}});
      slide.addShape(pptx.ShapeType.rect, {x: 4.8, y, w: 2.8, h: 0.45,
        fill: {color: '99BBEE'}, hyperlink: {slide: target, tooltip: 'shape-' + token}});
      slide.addImage({data: png, x: 8, y, w: 2.8, h: 0.45,
        hyperlink: {slide: target, tooltip: 'image-' + token}});
    }
    // A text hyperlink may override the containing shape's separate action.
    slide.addText([{text: 'mixed inner link', options: {hyperlink: {slide: 4, tooltip: 'mixed-inner'}}}],
      {x: 0.4, y: 6.6, w: 6, h: 0.45, fontSize: 18, hyperlink: {slide: 1, tooltip: 'mixed-surface'}});
  }
  const zip = await JSZip.loadAsync(await pptx.write({outputType: 'nodebuffer'}));
  const presentation = await zip.file('ppt/presentation.xml').async('string');
  const ids = [...presentation.matchAll(/<p:sldId\b[^>]*\/>/g)].map(m => m[0]);
  assert.equal(ids.length, 4);
  // Slide file numbers intentionally differ from presentation positions.
  zip.file('ppt/presentation.xml', presentation.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/,
    '<p:sldIdLst>' + [ids[2], ids[0], ids[3], ids[1]].join('') + '</p:sldIdLst>'));
  for (let part = 1; part <= 4; part++) {
    const slideName = 'ppt/slides/slide' + part + '.xml';
    let xml = await zip.file(slideName).async('string');
    xml = xml.replace(/<a:hlinkClick\b[^>]*>/g, tag => {
      // PptxGenJS omits the shape relationship when its text has a separate
      // run-level link. Supply that valid OOXML relationship explicitly below.
      if (tag.includes('tooltip="mixed-surface"')) return tag.replace(/r:id="[^"]*"/, 'r:id="rIdMixedSurface"');
      const match = /tooltip="(?:text|shape|image)-(next|previous|first|last)"/.exec(tag);
      return match ? tag.replace(/r:id="[^"]*"/, 'r:id=""')
        .replace('ppaction://hlinksldjump', 'ppaction://hlinkshowjump?jump=' + match[1] + 'slide') : tag;
    });
    if (part === 4) xml = xml.replace(/<p:sld\b/, '<p:sld show="0"');
    zip.file(slideName, xml);
    const relsName = 'ppt/slides/_rels/slide' + part + '.xml.rels';
    let rels = (await zip.file(relsName).async('string')).replace('</Relationships>',
      '<Relationship Id="rIdMixedSurface" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slide1.xml"/></Relationships>');
    // Exercise legal absolute OPC targets as well as relative targets.
    if (part % 2) rels = rels.replace(/Target="(?:\.\.\/slides\/)?slide(\d+)\.xml"/g, 'Target="/ppt/slides/slide$1.xml"');
    zip.file(relsName, rels);
  }
  return zip.generateAsync({type: 'nodebuffer', compression: 'DEFLATE'});
}

async function convertNative(bytes, options) {
  return new Promise((resolve, reject) => {
    const worker = new Worker(path.join(root, 'src/native-cli-worker.cjs'), {
      workerData: {bytes: Uint8Array.from(bytes).buffer, title: 'Internal navigation',
        slides: options.slides, includeHidden: options.includeHidden}
    });
    const timer = setTimeout(() => {worker.terminate(); reject(new Error('Native link conversion timed out'));}, 60000);
    worker.on('error', error => {clearTimeout(timer); reject(error);});
    worker.on('message', message => {
      if (!['result', 'error'].includes(message.type)) return;
      clearTimeout(timer);
      worker.terminate();
      if (message.type === 'error') reject(new Error(message.code + ': ' + message.message));
      else resolve(message);
    });
  });
}

function assertMarkup(html, scenario) {
  const dom = new JSDOM(html);
  try {
    const document = dom.window.document;
    const slides = [...document.querySelectorAll('#ptv-slides > .slide-shell')];
    assert.equal(slides.length, scenario.parts.length);
    assert.deepEqual(slides.map(slide => Number(slide.textContent.match(/PART_(\d)/)[1])), scenario.parts);
    assert.equal(document.querySelectorAll('#ptv-slides button, form, input, iframe, object').length, 0);
    assert.ok(!html.includes('PRIVATE_LINK_TEST_NOTE_MUST_NOT_EXPORT'));
    for (const node of document.querySelectorAll('#ptv-slides [data-action]')) assert.equal(node.dataset.action, 'slide');
    const scripts = [...document.querySelectorAll('script')];
    assert.equal(scripts.length, 1, 'only the trusted viewer script may survive');
    const csp = document.querySelector('meta[http-equiv="Content-Security-Policy"]').content;
    const hash = crypto.createHash('sha256').update(scripts[0].textContent).digest('base64');
    assert.ok(csp.includes("script-src 'sha256-" + hash + "'"), 'CSP must match the actual parsed viewer script');
    assert.ok(csp.includes("connect-src 'none'"));
    slides.forEach((slide, index) => {
      for (const kind of kinds) {
        for (let part = 1; part <= 4; part++) {
          const destination = scenario.parts.indexOf(part) + 1;
          const node = slide.querySelector(linkSelector(kind, 'slide-' + part));
          if (destination) {
            assert.ok(node, scenario.name + ': missing ' + kind + ' link');
            assert.equal(node.getAttribute('data-slide-target'), String(destination));
            assert.equal(node.getAttribute('role'), 'button');
            assert.equal(node.getAttribute('tabindex'), '0');
          } else if (kind === 'text') {
            assert.ok(node, 'excluded destinations retain the original text');
            assert.equal(node.hasAttribute('data-slide-target'), false);
            assert.equal(node.hasAttribute('role'), false);
            assert.equal(node.hasAttribute('tabindex'), false);
          } else assert.equal(node, null, 'excluded destinations must not leave an invisible clickable surface');
        }
        const destinations = {next: Math.min(slides.length, index + 2), previous: Math.max(1, index), first: 1, last: slides.length};
        for (const action of actions) {
          const node = slide.querySelector(linkSelector(kind, action));
          assert.ok(node, scenario.name + ': missing ' + kind + ' ' + action);
          assert.equal(node.getAttribute('data-slide-target'), String(destinations[action]),
            'relative actions use the selected result order, not a missing original neighbor');
        }
      }
      const surface = slide.querySelector('.shape-action-surface[title="mixed-surface"]');
      const inner = slide.querySelector('.run[title="mixed-inner"]');
      assert.equal(surface?.getAttribute('data-slide-target') || null,
        scenario.parts.includes(1) ? String(scenario.parts.indexOf(1) + 1) : null);
      assert.equal(inner.getAttribute('data-slide-target'),
        scenario.parts.includes(4) ? String(scenario.parts.indexOf(4) + 1) : null);
    });
  } finally {dom.window.close();}
}

function exporterContract() {
  const dom = new JSDOM('', {runScripts: 'outside-only'});
  try {
    dom.window.DOMPurify = DOMPurify(dom.window);
    const viewer = fs.readFileSync(path.join(root, 'src/viewer.js'), 'utf8').replace(/\r\n?/g, '\n');
    let exporter = fs.readFileSync(path.join(root, 'src/exporter.js'), 'utf8');
    for (const [marker, value] of [['VIEWER_CSS', fs.readFileSync(path.join(root, 'src/viewer.css'), 'utf8')],
      ['VIEWER_SCRIPT', viewer], ['VIEWER_HASH', crypto.createHash('sha256').update(viewer).digest('base64')]]) {
      exporter = exporter.replace('/* ' + marker + ' */ null', () => JSON.stringify(value));
    }
    dom.window.eval(exporter);
    const shell = (number, content) => '<div class="slide-shell" data-slide="' + number + '"><div class="slide" data-slide="' + number + '">' + content + '</div></div>';
    const hostile = '<button data-slide-target="4" onclick="alert(1)">arbitrary</button>' +
      '<div class="shape"><button class="shape-action-surface" type="button" data-action="slide" data-slide-target="4"><img src="https://example.invalid/x"></button>' +
      '<button class="shape-action-surface" type="button" data-action="media"></button>' +
      '<button class="shape-action-surface" type="submit" data-action="slide" data-slide-target="4"></button></div>' +
      '<button class="shape-action-surface" type="button" data-action="slide" data-slide-target="4"></button>' +
      '<form action="https://example.invalid/submit"><input name="leak"></form><script>alert(2)</script>';
    const result = dom.window.PptxExport.prepare({slideCount: 2, html: '<div class="pptx-container">' +
      shell(2, hostile + '<a href="#slide-4" title="native-anchor">native anchor</a>') + shell(4, 'destination') + '</div>'});
    const clean = new JSDOM(result.html);
    try {
      const slides = clean.window.document.getElementById('ptv-slides');
      assert.equal(slides.querySelectorAll('button, form, input, script, .shape-action-surface, [onclick]').length, 0);
      assert.equal(slides.querySelector('a').getAttribute('href'), '#ptv-slide-2');
      assert.ok(!result.html.includes('example.invalid'));
    } finally {clean.window.close();}
    return result.html;
  } finally {dom.window.close();}
}

async function verifyInteractions(scope, scenario) {
  await scope.locator('#ptv-select option[value="' + (scenario.parts.length - 1) + '"]').waitFor({state: 'attached'});
  assert.equal(await scope.locator('#ptv-select option').count(), scenario.parts.length);
  if (await scope.locator('#ptv-mode').getAttribute('aria-pressed') === 'false') await scope.locator('#ptv-mode').click();
  const source = Math.min(1, scenario.parts.length - 1);
  const destination = source ? 0 : scenario.parts.length - 1;
  async function activate(kind, token, index, expected, key) {
    await scope.locator('#ptv-select').selectOption(String(index));
    const link = scope.locator('#ptv-slides > .slide-shell').nth(index).locator(linkSelector(kind, token));
    if (key === 'click') await link.click();
    else {await link.focus(); await link.press(key);}
    assert.equal(await scope.locator('#ptv-select').inputValue(), String(expected),
      scenario.name + ': ' + kind + ' ' + token + ' by ' + key);
    if (key !== 'click') assert.equal(await scope.locator('#ptv-slides > .slide-shell').nth(expected)
      .evaluate(node => node.ownerDocument.activeElement === node), true, 'keyboard navigation must focus the destination slide');
  }
  for (const kind of kinds) for (const key of ['click', 'Enter', 'Space']) {
    await activate(kind, 'slide-' + scenario.parts[destination], source, destination, key);
  }
  // The middle hidden/omitted original slides must not interrupt next/previous.
  await activate('text', 'next', 0, Math.min(1, scenario.parts.length - 1), 'click');
  await activate('shape', 'previous', source, Math.max(0, source - 1), 'Enter');
  await activate('image', 'first', scenario.parts.length - 1, 0, 'Space');
  await activate('text', 'last', 0, scenario.parts.length - 1, 'Enter');
  await activate('shape', 'previous', 0, 0, 'click');
  await activate('image', 'next', scenario.parts.length - 1, scenario.parts.length - 1, 'Space');
  const missing = [1, 2, 3, 4].find(part => !scenario.parts.includes(part));
  if (missing) {
    await scope.locator('#ptv-select').selectOption('0');
    const text = scope.locator('#ptv-slides > .slide-shell').first().locator(linkSelector('text', 'slide-' + missing));
    await text.click();
    assert.equal(await scope.locator('#ptv-select').inputValue(), '0', 'an excluded target must not jump to a renumbered slide');
  }
  if (scenario.parts.includes(1) && !scenario.parts.includes(4) && scenario.parts.length > 1) {
    const index = scenario.parts.indexOf(1) === 0 ? 1 : 0;
    await scope.locator('#ptv-select').selectOption(String(index));
    const shell = scope.locator('#ptv-slides > .slide-shell').nth(index);
    assert.equal(await shell.locator('.shape-action-surface[title="mixed-surface"]').count(), 1);
    // Hit the text's coordinates even if a sibling surface intercepts them.
    await shell.locator('.run[title="mixed-inner"]').click({force: true});
    assert.equal(await scope.locator('#ptv-select').inputValue(), String(index),
      'an excluded text target must not fall through to its containing shape\'s different destination');
  }
  if (scenario.parts.includes(4)) {
    await scope.locator('#ptv-select').selectOption('0');
    await scope.locator('#ptv-slides > .slide-shell').first().locator('.run[title="mixed-inner"]').click();
    assert.equal(await scope.locator('#ptv-select').inputValue(), String(scenario.parts.indexOf(4)),
      'an included text link must override the containing shape\'s different destination');
  }
}

async function main() {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'pptx-internal-links-'));
  let browser;
  try {
    const bytes = await createFixture();
    const fixture = path.join(tmp, 'internal-links.pptx');
    fs.writeFileSync(fixture, bytes);
    const nativeAnchor = path.join(tmp, 'native-anchor.html');
    fs.writeFileSync(nativeAnchor, exporterContract());
    for (const scenario of cases) {
      const result = await convertNative(bytes, scenario);
      assert.deepEqual(result.sourceSlides, scenario.originals);
      assertMarkup(result.html, scenario);
      assert.equal(result.warnings.some(w => w.code === 'SLIDE_LINK_TARGET_EXCLUDED'), scenario.parts.length < 4);
      fs.writeFileSync(path.join(tmp, scenario.name + '.html'), result.html);
    }
    console.log('PASS native WASM navigation: reordered parts, hidden/range selection, targets, relative actions, sanitizer, CSP');
    const executablePath = browserExecutable();
    assert.ok(executablePath, 'No ' + browserName + ' executable is available; install it or set its executable path.');
    browser = await browserType.launch({headless: true, executablePath,
      ...(browserName === 'chromium' ? {args: ['--no-sandbox']} : {})});
    const context = await browser.newContext({acceptDownloads: true, viewport: {width: 1280, height: 900}});
    const requests = [], errors = [];
    await context.route(/^https?:\/\//, route => {requests.push(route.request().url()); return route.abort('internetdisconnected');});
    context.on('page', page => {
      page.on('pageerror', error => errors.push(error.message));
      page.on('console', message => {if (message.type() === 'error') errors.push(message.text());});
    });
    const page = await context.newPage();
    for (const scenario of cases) {
      await page.goto(pathToFileURL(path.join(tmp, scenario.name + '.html')).href);
      await verifyInteractions(page, scenario);
    }
    await page.goto(pathToFileURL(nativeAnchor).href);
    await page.locator('#ptv-select option[value="1"]').waitFor({state: 'attached'});
    await page.locator('#ptv-mode').click();
    const anchor = page.locator('[title="native-anchor"]');
    await anchor.evaluate(node => {node.dataset.clicks = '0'; node.addEventListener('click', () => {node.dataset.clicks = String(Number(node.dataset.clicks) + 1);});});
    await anchor.focus(); await anchor.press('Enter');
    assert.equal(await page.locator('#ptv-select').inputValue(), '1');
    assert.equal(await anchor.getAttribute('data-clicks'), '1', 'native anchor Enter must not synthesize a second click');
    assert.equal(await page.locator('#ptv-slides > .slide-shell').nth(1)
      .evaluate(node => node.ownerDocument.activeElement === node), true, 'native anchor keyboard activation must focus the destination');
    await page.locator('#ptv-select').selectOption('0');
    await anchor.focus(); await anchor.press('Space');
    assert.equal(await page.locator('#ptv-select').inputValue(), '0', 'native anchor Space keeps its normal behavior');
    assert.equal(await anchor.getAttribute('data-clicks'), '1');
    console.log('PASS ' + browserName + ' standalone click/Enter/Space navigation and native anchor behavior');

    // Browser conversion uses the same exporter in a Blob preview and download.
    await page.goto(pathToFileURL(path.join(root, 'index.html')).href);
    await page.locator('#file-input').setInputFiles(fixture);
    const scenario = cases.find(item => item.name === 'range-hidden');
    await page.locator('#slide-range').fill(scenario.slides);
    await page.locator('#include-hidden').check();
    await page.locator('#convert-button').click();
    await page.waitForFunction(() => !document.querySelector('#download-button').disabled ||
      document.querySelector('#status-line').dataset.tone === 'error', null, {timeout: 60000});
    assert.ok(await page.locator('#download-button').isEnabled(), await page.locator('#status-text').innerText());
    await verifyInteractions(page.frameLocator('#preview-frame'), scenario);
    const pending = page.waitForEvent('download');
    await page.locator('#download-button').click();
    const download = await pending;
    const downloaded = path.join(tmp, 'browser-export.html');
    await download.saveAs(downloaded);
    assertMarkup(fs.readFileSync(downloaded, 'utf8'), scenario);
    await page.goto(pathToFileURL(downloaded).href);
    await verifyInteractions(page, scenario);
    assert.deepEqual(requests, [], 'conversion, preview, and standalone navigation must make no external requests');
    assert.deepEqual(errors, [], 'viewer, worker, and CSP must produce no browser errors');
    console.log('PASS ' + browserName + ' browser WASM preview/download navigation; zero external requests or CSP errors');
  } finally {
    if (browser) await browser.close();
    fs.rmSync(tmp, {recursive: true, force: true});
  }
}

main().catch(error => {console.error(error); process.exitCode = 1;});
