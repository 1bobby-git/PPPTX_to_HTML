import { readFile, writeFile, mkdir, readdir } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { join, dirname } from 'node:path';
import { createHash } from 'node:crypto';
import vm from 'node:vm';
import browserify from 'browserify';
import browserifyBuiltins from 'browserify/lib/builtins.js';

const root = dirname(dirname(fileURLToPath(import.meta.url)));
// Keep embedded scripts/styles and generated artifacts independent of checkout EOLs.
const read = async (path) => (await readFile(join(root, path), 'utf8')).replace(/\r\n?/g, '\n');
const enginePath = 'node_modules/@briank-dev/pptx-to-html/';
const enginePackage = JSON.parse(await read(enginePath + 'package.json'));
const engineVersion = enginePackage.version;
if (enginePackage.name !== '@briank-dev/pptx-to-html' || typeof engineVersion !== 'string' ||
    !/^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?(?:\+[0-9A-Za-z.-]+)?$/.test(engineVersion)) {
  throw new Error('Unexpected engine package metadata.');
}
const engineBuildInfo = await read(enginePath + 'BUILD-INFO.json').catch(error => {
  if (error.code === 'ENOENT') return '';
  throw error;
});
const [shell, css, app, adapter, exporterTemplate, viewer, viewerCss, glueSource, wasm, purifierSource, engineLicense, purifierLicense, zipSource, zipLicense, normalizer, conversionCore] = await Promise.all([
  read('src/shell.html'), read('src/styles.css'), read('src/app.js'), read('src/worker.js'),
  read('src/exporter.js'), read('src/viewer.js').then(source => source.replace(/\r\n?/g, '\n')), read('src/viewer.css'),
  read(enginePath + 'pptx2html_wasm.js'), readFile(join(root, enginePath, 'pptx2html_wasm_bg.wasm')),
  read('node_modules/dompurify/dist/purify.min.js'), read(enginePath + 'LICENSE'), read('node_modules/dompurify/LICENSE'),
  read('node_modules/jszip/dist/jszip.min.js'), read('node_modules/jszip/LICENSE.markdown'), read('src/normalize-package.js'),
  read('src/conversion-core.js')
]);

// Bundle the password reader and all browser shims into the same offline worker.
const cryptoPackages = new Map();
const cryptoBundler = browserify(join(root, 'src/input-package.js'), {
  standalone: 'PptxInput', debug: false,
  builtins: {...browserifyBuiltins, crypto: join(root, 'src/browser-office-crypto.cjs')}
});
cryptoBundler.on('package', (pkg) => {
  if (pkg.__dirname && pkg.__dirname.includes('node_modules')) cryptoPackages.set(pkg.__dirname, pkg);
});
const cryptoSource = await new Promise((resolve, reject) => {
  cryptoBundler.bundle((error, buffer) => error ? reject(error) : resolve(buffer.toString('utf8').replace(/\r\n?/g, '\n')));
});
// Office decryption needs symmetric crypto only. Fail the build if a future
// dependency pulls the unused elliptic-curve/signature implementations back in.
const forbiddenCrypto = new Set(['crypto-browserify', 'browserify-sign', 'create-ecdh', 'elliptic']);
for (const pkg of cryptoPackages.values()) {
  if (forbiddenCrypto.has(pkg.name)) throw new Error('Unexpected browser crypto dependency: ' + pkg.name);
}
const cryptoLicenseBlocks = await Promise.all([...cryptoPackages.values()].sort((a, b) =>
  a.name.localeCompare(b.name, 'en') || a.version.localeCompare(b.version, 'en')).map(async (pkg) => {
  const names = await readdir(pkg.__dirname);
  const licenses = names.filter((name) => /^(?:licen[cs]e|copying)(?:[._-]|$)/i.test(name));
  let texts = await Promise.all(licenses.map((name) => readFile(join(pkg.__dirname, name), 'utf8')));
  if (!texts.length) {
    const readmeName = names.find((name) => /^readme(?:\.|$)/i.test(name));
    const readme = readmeName ? await readFile(join(pkg.__dirname, readmeName), 'utf8') : '';
    const section = /^(?:#{1,6}\s*)?licen[cs]e\b[^\n]*\n[\s\S]*/im.exec(readme);
    if (!section) throw new Error('Missing bundled dependency license: ' + pkg.name);
    texts = [section[0]];
  }
  return pkg.name + ' ' + pkg.version + ' (' + (pkg.license || 'see original notice') + ')\n' + texts.join('\n');
}));
const cryptoLicenses = cryptoLicenseBlocks.join('\n\n' + '-'.repeat(72) + '\n\n');

function replaceOnce(source, marker, replacement) {
  if (source.split(marker).length !== 2) throw new Error('Expected one build marker: ' + marker);
  return source.replace(marker, () => replacement);
}
function scriptSafe(source) { return source.replace(/<\/script/gi, '<\\/script'); }
// HTML parsing normalizes CRLF/CR before CSP checks. The viewer above uses the
// same LF text for both embedding and hashing, including on Windows checkouts.
function hash(source) { return createHash('sha256').update(source).digest('base64'); }

// Package is pinned. Transform its ESM export syntax into a classic Blob Worker.
const glue = glueSource
  .replace(/^export (?=(?:class|function|async function)\b)/gm, '')
  .replace(/^export \{[^\n]+\};?$/gm, '')
  .replace("module_or_path = new URL('pptx2html_wasm_bg.wasm', import.meta.url);", "throw new Error('Embedded WASM bytes are required.');");
if (/\bimport\.meta\b|^export\b/m.test(glue)) throw new Error('Unexpected upstream module format; review before bundling.');
const worker = zipSource + '\n' + cryptoSource + '\n' + normalizer + '\n' + conversionCore + '\n' +
  '/* pptx2html-turbo v' + engineVersion + ' — MIT — Copyright (c) 2026 Brian Kim */\n' +
  glue + '\nconst EMBEDDED_WASM_BASE64 = ' + JSON.stringify(wasm.toString('base64')) + ';\n' + adapter;
let exporter = replaceOnce(exporterTemplate, '/* VIEWER_CSS */ null', JSON.stringify(viewerCss));
exporter = replaceOnce(exporter, '/* VIEWER_SCRIPT */ null', JSON.stringify(viewer));
exporter = replaceOnce(exporter, '/* VIEWER_HASH */ null', JSON.stringify(hash(viewer)));
const purifier = purifierSource.replace(/\/\/# sourceMappingURL=.*$/gm, '');
for (const [filename, script] of [['worker', worker], ['purifier', purifier], ['exporter', exporter], ['app', app], ['viewer', viewer]]) {
  new vm.Script(script, {filename});
}
const csp = "default-src 'none'; script-src 'unsafe-inline' 'wasm-unsafe-eval'; style-src 'unsafe-inline'; img-src data: blob:; font-src data:; media-src data: blob:; worker-src blob:; frame-src blob:; connect-src 'none'; object-src 'none'; base-uri 'none'; form-action 'none'";
let html = shell.replace('<meta charset="utf-8">', '<meta charset="utf-8">\n  <meta http-equiv="Content-Security-Policy" content="' + csp + '">');
html = replaceOnce(html, '<!-- APP_STYLES -->', '<style>\n' + css + '\n</style>');
html = replaceOnce(html, '<!-- ENGINE_WORKER -->', '<script id="engine-worker" type="text/plain">\n' + scriptSafe(worker) + '\n</script>');
html = replaceOnce(html, '<!-- APP_SCRIPT -->', [purifier, exporter, app].map((code) => '<script>\n' + scriptSafe(code) + '\n</script>').join('\n'));
html = html.replace('</head>', '<!-- Bundled third-party licenses are in the inert block at the end of this file. -->\n</head>');
const licenseText = 'pptx2html-turbo ' + engineVersion + '\n' + engineLicense +
  (engineBuildInfo ? '\n\nPatched engine build information\n' + engineBuildInfo : '') +
  '\n\nDOMPurify (Apache-2.0 OR MPL-2.0)\n' + purifierLicense + '\n\nJSZip 3.10.1 (MIT option selected)\n' + zipLicense + '\n\nOffice password reader and bundled dependencies\n' + cryptoLicenses;
// Vendor code may contain literal HTML strings. Insert at the actual closing body,
// never the first occurrence of </body> inside an embedded JavaScript string.
const bodyEnd = html.lastIndexOf('</body>');
if (bodyEnd < 0) throw new Error('Missing closing body in app shell.');
html = html.slice(0, bodyEnd) + '<script type="text/plain" id="third-party-licenses">\n' + scriptSafe(licenseText) + '\n</script>\n' + html.slice(bodyEnd);
await writeFile(join(root, 'index.html'), html);
await mkdir(join(root, 'licenses'), {recursive: true});
await writeFile(join(root, 'licenses', 'pptx2html-turbo-MIT.txt'), engineLicense);
await writeFile(join(root, 'licenses', 'DOMPurify-LICENSE.txt'), purifierLicense);
await writeFile(join(root, 'licenses', 'JSZip-LICENSE.txt'), zipLicense);
await writeFile(join(root, 'licenses', 'OfficeCrypto-BUNDLED-LICENSES.txt'), cryptoLicenses);
console.log('Built index.html (' + (Buffer.byteLength(html) / 1024 / 1024).toFixed(2) + ' MiB), embedded WASM and dependencies.');
