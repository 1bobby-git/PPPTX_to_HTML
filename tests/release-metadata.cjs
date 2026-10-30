'use strict';
// Keeps user-visible version labels and download instructions on the released version.
const assert = require('node:assert/strict');
const test = require('node:test');
const fs = require('node:fs');
const path = require('node:path');

const root = path.resolve(__dirname, '..');
// Windows checkouts may convert text files to CRLF; compare LF content everywhere.
const read = file => fs.readFileSync(path.join(root, file), 'utf8').replace(/\r\n/g, '\n');
const json = file => JSON.parse(read(file));
const version = json('package.json').version;
const tag = 'v' + version;

test('package, lockfile and desktop app declare the same version', () => {
  const lock = json('package-lock.json');
  assert.equal(lock.version, version);
  assert.equal(lock.packages[''].version, version);
  assert.equal(json('desktop/package.json').version, version);
  assert.ok(read(`docs/releases/${tag}.md`).trim(), `docs/releases/${tag}.md must not be empty`);
});

test('converter footers show the current version', () => {
  for (const file of ['src/shell.html', 'src/web-shell.html', 'index.html']) {
    const footers = read(file).match(/<footer>[^<]*<\/footer>/g) || [];
    assert.equal(footers.length, 1, `${file} must have one footer`);
    const labels = footers[0].match(/v\d+\.\d+\.\d+/g) || [];
    assert.deepEqual(labels, [tag], `${file} footer must show ${tag}`);
  }
});

test('README and Native CLI guide point to the current release assets', () => {
  const readme = read('README.md');
  assert.match(readme, new RegExp(`최신 버전은 \\*\\*${tag.replace(/\./g, '\\.')}\\*\\*`));
  assert.ok(readme.includes(`/releases/tag/${tag})`), `README download link must use ${tag}`);
  assert.ok(readme.includes(`PPPTX_to_HTML_Desktop_${tag}_win32-x64.zip`));
  assert.ok(readme.includes(`PPPTX_to_HTML_Native_${tag}.zip`));
  const guide = read('docs/NATIVE_CLI.md');
  assert.ok(guide.startsWith(`# PPTX → HTML Native CLI · ${tag}\n`), `docs/NATIVE_CLI.md title must use ${tag}`);
  assert.ok(guide.includes(`PPPTX_to_HTML_Native_${tag}.zip`));
  for (const [file, text] of [['README.md', readme], ['docs/NATIVE_CLI.md', guide]]) {
    const assets = text.match(/PPPTX_to_HTML_(?:Desktop|Native)_v\d+\.\d+\.\d+/g) || [];
    assert.deepEqual([...new Set(assets.map(name => name.replace(/^.*_(v[\d.]+)$/, '$1')))], [tag], `${file} must not point to older release assets`);
  }
});
