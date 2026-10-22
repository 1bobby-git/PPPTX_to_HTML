#!/usr/bin/env node
'use strict';

// Run by cli.py against outputs made by the public Python CLI.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
const { chromium } = require('playwright');

function browserExecutable() {
  const candidates = [process.env.CHROME_BIN, process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH,
    chromium.executablePath(), '/tmp/pptx-browsers/chrome-linux64/chrome',
    '/usr/bin/chromium', '/usr/bin/google-chrome'];
  return candidates.find(candidate => candidate && fs.existsSync(candidate));
}

async function main() {
  const [regular, hidden, selected, hostileTitle] = process.argv.slice(2).map(x => path.resolve(x));
  assert.ok(regular && hidden && selected && hostileTitle, 'four generated HTML paths are required');
  const executablePath = browserExecutable();
  assert.ok(executablePath, 'Chromium is required for the file:// browser check');
  const browser = await chromium.launch({ headless: true, executablePath, args: ['--no-sandbox'] });
  const externalRequests = [];
  const errors = [];
  try {
    const context = await browser.newContext({ viewport: { width: 1280, height: 850 } });
    await context.route(/^https?:\/\//, route => {
      externalRequests.push(route.request().url());
      return route.abort('internetdisconnected');
    });
    const page = await context.newPage();
    page.on('pageerror', error => errors.push(error.message));

    async function openOutput(file) {
      await page.goto(pathToFileURL(file).href);
      assert.equal(new URL(page.url()).protocol, 'file:');
      const slides = page.locator('img[src^="data:image/svg+xml"]');
      await slides.first().waitFor();
      await page.waitForFunction(() => [...document.querySelectorAll('img[src^="data:image/svg+xml"]')]
        .every(image => image.complete && image.naturalWidth > 0));
      return slides;
    }

    const slides = await openOutput(regular);
    assert.equal(await slides.count(), 3, 'hidden slide should be excluded by default');
    const accessibleText = await slides.evaluateAll(images => images.map(image => image.closest('figure').textContent));
    assert.match(accessibleText[0], /TABLE_FILE_TWO/, 'presentation order must beat XML file order');
    assert.match(accessibleText[1], /INTRO_FILE_ONE/, 'second slide must be original intro');
    assert.match(accessibleText[2], /IMAGE_CHART_FILE_THREE/);
    assert.ok(accessibleText.some(text => text.includes('안녕하세요 오프라인')), 'Korean content is retained');
    assert.ok(!accessibleText.join('').includes('HIDDEN_FILE_FOUR'));
    assert.ok(!fs.readFileSync(regular, 'utf8').includes('PRIVATE_SPEAKER_NOTE_MUST_NOT_EXPORT'));

    const firstDetails = page.locator('main > figure details').first();
    await firstDetails.locator('summary').click();
    assert.equal(await firstDetails.getAttribute('open'), '');
    const selectedSlideText = await firstDetails.locator('pre').evaluate(pre => {
      const selection = window.getSelection();
      selection.removeAllRanges();
      const range = document.createRange();
      range.selectNodeContents(pre);
      selection.addRange(range);
      return selection.toString();
    });
    assert.match(selectedSlideText, /TABLE_FILE_TWO/, 'slide text can be selected');

    await page.locator('#mode').click();
    assert.equal(await page.locator('#mode').getAttribute('aria-pressed'), 'true');
    assert.equal(await page.locator('main > figure:visible').count(), 1);
    assert.equal(await page.locator('#prev').isDisabled(), true);
    await page.locator('#next').click();
    assert.equal(await page.locator('#slide-select').inputValue(), '1');
    await page.locator('#slide-select').selectOption('2');
    assert.equal(await page.locator('#next').isDisabled(), true);
    const keyboard = await page.evaluate(() => {
      document.activeElement.blur();
      const event = new KeyboardEvent('keydown', { key: 'ArrowLeft', bubbles: true, cancelable: true });
      const allowed = document.dispatchEvent(event);
      return { allowed, selected: document.querySelector('#slide-select').value };
    });
    assert.deepEqual(keyboard, { allowed: false, selected: '1' }, 'arrow navigation must prevent page scrolling');

    await page.setViewportSize({ width: 390, height: 844 });
    const mobile = await page.evaluate(() => ({ scroll: document.documentElement.scrollWidth, viewport: innerWidth }));
    assert.ok(mobile.scroll <= mobile.viewport + 1, `mobile horizontal overflow: ${JSON.stringify(mobile)}`);

    assert.equal(await (await openOutput(hidden)).count(), 4, '--include-hidden retains fourth slide');
    assert.equal(await (await openOutput(selected)).count(), 1, '--slides 2 emits one slide');
    const selectedText = await page.locator('main > figure').first().textContent();
    assert.match(selectedText, /INTRO_FILE_ONE/, '--slides 2 must select the original second slide');

    await openOutput(hostileTitle);
    assert.equal(await page.evaluate(() => window.__pptxTitleExecuted), undefined, 'title HTML executed');
    assert.equal(externalRequests.length, 0, `network requests: ${externalRequests.join(', ')}`);
    assert.deepEqual(errors, [], `browser errors: ${errors.join(', ')}`);
    await context.close();
  } finally {
    await browser.close();
  }
}

main().catch(error => { console.error(error); process.exitCode = 1; });
