import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { join, dirname } from 'node:path';
import { createHash } from 'node:crypto';
import vm from 'node:vm';

const root = dirname(dirname(fileURLToPath(import.meta.url)));
const read = (path) => readFile(join(root, path), 'utf8');
const enginePath = 'node_modules/@briank-dev/pptx-to-html/';
const [shell, css, app, adapter, exporterTemplate, viewer, viewerCss, glueSource, wasm, purifierSource, engineLicense, purifierLicense, zipSource, zipLicense, normalizer] = await Promise.all([
  read('src/shell.html'), read('src/styles.css'), read('src/app.js'), read('src/worker.js'),
  read('src/exporter.js'), read('src/viewer.js'), read('src/viewer.css'),
  read(enginePath + 'pptx2html_wasm.js'), readFile(join(root, enginePath, 'pptx2html_wasm_bg.wasm')),
  read('node_modules/dompurify/dist/purify.min.js'), read(enginePath + 'LICENSE'), read('node_modules/dompurify/LICENSE'),
  read('node_modules/jszip/dist/jszip.min.js'), read('node_modules/jszip/LICENSE.markdown'), read('src/normalize-package.js')
]);

function replaceOnce(source, marker, replacement) {
  if (source.split(marker).length !== 2) throw new Error('Expected one build marker: ' + marker);
  return source.replace(marker, () => replacement);
}
function scriptSafe(source) { return source.replace(/<\/script/gi, '<\\/script'); }
function hash(source) { return createHash('sha256').update(source).digest('base64'); }

// Package is pinned. Transform its ESM export syntax into a classic Blob Worker.
const glue = glueSource
  .replace(/^export (?=(?:class|function|async function)\b)/gm, '')
  .replace(/^export \{[^\n]+\};?$/gm, '')
  .replace("module_or_path = new URL('pptx2html_wasm_bg.wasm', import.meta.url);", "throw new Error('Embedded WASM bytes are required.');");
if (/\bimport\.meta\b|^export\b/m.test(glue)) throw new Error('Unexpected upstream module format; review before bundling.');
const worker = zipSource + '\n' + normalizer + '\n' +
  '/* pptx2html-turbo v2.1.0 — MIT — Copyright (c) 2026 Brian Kim */\n' +
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
const licenseText = 'pptx2html-turbo 2.1.0\n' + engineLicense + '\n\nDOMPurify (Apache-2.0 OR MPL-2.0)\n' + purifierLicense + '\n\nJSZip 3.10.1 (MIT option selected)\n' + zipLicense;
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
console.log('Built index.html (' + (Buffer.byteLength(html) / 1024 / 1024).toFixed(2) + ' MiB), embedded WASM and dependencies.');
