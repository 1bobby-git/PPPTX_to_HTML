#!/usr/bin/env node
'use strict';

// Package the installed production dependency tree without fetching or executing it.
const fs = require('node:fs/promises');
const path = require('node:path');
const crypto = require('node:crypto');
const JSZip = require('jszip');

const root = path.resolve(__dirname, '..');
const read = relative => fs.readFile(path.join(root, relative));
const fixedDate = new Date('2000-01-01T00:00:00Z');

async function main() {
  const pkg = JSON.parse(await read('package.json'));
  const lock = JSON.parse(await read('package-lock.json'));
  if (pkg.version !== lock.version || pkg.version !== lock.packages?.['']?.version) {
    throw new Error('package.json and package-lock.json versions differ.');
  }
  const prefix = 'PPPTX_to_HTML_Native/';
  const zip = new JSZip();
  const seen = new Set();
  const packages = [];
  let sourceBytes = 0;
  async function add(relative, archiveName = relative, transform) {
    if (seen.has(archiveName)) return;
    if (path.isAbsolute(archiveName) || archiveName.split('/').includes('..')) throw new Error('Unsafe archive path.');
    const info = await fs.lstat(path.join(root, relative));
    if (!info.isFile()) throw new Error('Only regular runtime files may be packaged: ' + relative);
    const source = await read(relative);
    const bytes = transform ? Buffer.from(transform(source.toString('utf8')), 'utf8') : source;
    zip.file(prefix + archiveName, bytes, {date: fixedDate, createFolders: false, unixPermissions: 0o100644});
    seen.add(archiveName);
    sourceBytes += bytes.length;
  }
  async function visit(relative) {
    const entries = await fs.readdir(path.join(root, relative), {withFileTypes: true});
    for (const entry of entries.sort((a, b) => a.name.localeCompare(b.name, 'en'))) {
      // Nested dependency roots are enumerated from the lockfile separately.
      if (entry.name === 'node_modules' || entry.name === '.git' || entry.name === '.DS_Store') continue;
      const child = relative + '/' + entry.name;
      if (entry.isDirectory()) await visit(child);
      else if (entry.isFile()) await add(child);
      else throw new Error('Unexpected link or special file in runtime: ' + child);
    }
  }
  for (const relative of [
    'tools/convert-pptx.cjs', 'tools/Convert-PptxNative.cmd', 'tools/serve-web.cjs',
    'src/native-cli-worker.cjs', 'src/conversion-core.js', 'src/input-package.js', 'src/normalize-package.js', 'src/chart-data.js', 'src/theme-resolver.js',
    'src/exporter.js', 'src/viewer.js', 'src/viewer.css',
    'src/web-shell.html', 'src/server-client.js', 'src/app.js', 'src/styles.css',
    'docs/NATIVE_CLI.md', 'docs/WEB_SERVICE.md', 'docs/REFERENCE_REVIEW.md', 'package.json', 'package-lock.json', 'LICENSE'
  ]) await add(relative);
  await add('docs/NATIVE_CLI.md', 'README.md', text => text
    .replace('(../LICENSE)', '(LICENSE)').replace('(WEB_SERVICE.md)', '(docs/WEB_SERVICE.md)')
    .replace('(REFERENCE_REVIEW.md)', '(docs/REFERENCE_REVIEW.md)'));

  for (const [relative, record] of Object.entries(lock.packages).sort(([a], [b]) => a.localeCompare(b, 'en'))) {
    if (!relative.startsWith('node_modules/') || record.dev) continue;
    if (record.link) throw new Error('Linked dependency cannot be packaged: ' + relative);
    try { await fs.access(path.join(root, relative, 'package.json')); }
    catch (error) {
      if (record.optional && error.code === 'ENOENT') continue;
      throw new Error('Production dependency is missing; run npm ci: ' + relative);
    }
    const installed = JSON.parse(await read(relative + '/package.json'));
    if (installed.version !== record.version) throw new Error('Installed version does not match lockfile: ' + relative);
    await visit(relative);
    packages.push({name: installed.name, version: installed.version, license: installed.license || 'See package notices', path: relative});
  }
  for (const name of Object.keys(pkg.dependencies)) {
    if (!packages.some(item => item.path === 'node_modules/' + name)) throw new Error('Missing direct runtime dependency: ' + name);
  }
  const manifest = {version: pkg.version, node: pkg.engines.node, packages, files: seen.size, uncompressedBytes: sourceBytes,
    note: 'Node.js itself is not bundled. Package license and source files are preserved under node_modules.'};
  zip.file(prefix + 'RUNTIME-MANIFEST.json', JSON.stringify(manifest, null, 2) + '\n',
    {date: fixedDate, createFolders: false, unixPermissions: 0o100644});
  const bytes = await zip.generateAsync({type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: {level: 6}, platform: 'UNIX'});
  const output = path.join(root, 'dist', `PPPTX_to_HTML_Native_v${pkg.version}.zip`);
  await fs.mkdir(path.dirname(output), {recursive: true});
  await fs.writeFile(output, bytes);
  console.log(JSON.stringify({output, bytes: bytes.length, files: seen.size + 1, packages: packages.length,
    sha256: crypto.createHash('sha256').update(bytes).digest('hex')}, null, 2));
}
main().catch(error => { console.error(error.message); process.exitCode = 1; });
