#!/usr/bin/env node
'use strict';

const fs = require('node:fs');
const fsp = require('node:fs/promises');
const path = require('node:path');
const crypto = require('node:crypto');
const {Worker} = require('node:worker_threads');

const VERSION = require('../package.json').version;
const MAX_INPUT = 64 * 1024 * 1024;
const MAX_OUTPUT = 256 * 1024 * 1024;
const HELP = `Usage: node tools/convert-pptx.cjs input.pptx [options]

Offline native PPTX to standalone HTML conversion. No office app or browser is required.

Options:
  -o, --output FILE       Output HTML path (default: same stem as input)
  --slides LIST           Slide numbers/ranges, e.g. 1,3-5
  --include-hidden        Include hidden slides
  --title TEXT            Override HTML title
  --timeout SECONDS       Time limit, 1-3600 (default: 120)
  --overwrite             Replace an existing output file
  --password-stdin        Read a document password from stdin once
  --json                  Print one JSON result on stdout
  --check                 Verify runtime, packages, WASM, and exporter
  -h, --help              Show this help

Exit status: 0 success; 1 conversion or system error; 2 argument error; 130 interrupted.
`;

function fail(code, message, status = 1) { const error = new Error(message); error.code = code; error.status = status; throw error; }
function supportedNode() {
  const [major, minor, patch] = process.versions.node.split('.').map(Number);
  return major >= 26 || (major === 24 && (minor > 15 || (minor === 15 && patch >= 0))) ||
    (major === 22 && (minor > 22 || (minor === 22 && patch >= 2)));
}
function parse(argv) {
  const opts = {timeout: 120, slides: '', includeHidden: false, overwrite: false, json: false, check: false, passwordStdin: false};
  for (let i = 0; i < argv.length; i++) {
    const item = argv[i];
    if (item === '-h' || item === '--help') opts.help = true;
    else if (item === '-o' || item === '--output' || item === '--slides' || item === '--title' || item === '--timeout') {
      if (i + 1 >= argv.length) fail('INVALID_ARGUMENT', `${item} requires a value.`, 2);
      const val = argv[++i];
      if (item === '-o' || item === '--output') opts.output = val;
      else if (item === '--slides') opts.slides = val;
      else if (item === '--title') opts.title = val;
      else opts.timeout = Number(val);
    } else if (item === '--include-hidden') opts.includeHidden = true;
    else if (item === '--overwrite') opts.overwrite = true;
    else if (item === '--json') opts.json = true;
    else if (item === '--check') opts.check = true;
    else if (item === '--password-stdin') opts.passwordStdin = true;
    else if (item.startsWith('-')) fail('INVALID_ARGUMENT', `Unknown option: ${item}`, 2);
    else if (!opts.input) opts.input = item;
    else fail('INVALID_ARGUMENT', 'Specify exactly one input file.', 2);
  }
  if (opts.help) return opts;
  if (!Number.isInteger(opts.timeout) || opts.timeout < 1 || opts.timeout > 3600) fail('INVALID_ARGUMENT', '--timeout must be an integer from 1 to 3600.', 2);
  if (opts.slides && !/^\s*\d+(?:\s*-\s*\d+)?(?:\s*,\s*\d+(?:\s*-\s*\d+)?)*\s*$/.test(opts.slides)) fail('INVALID_ARGUMENT', '--slides must use 1,3-5 syntax.', 2);
  if (!opts.check && !opts.input) fail('INVALID_ARGUMENT', 'Input PPTX path is required.', 2);
  if (opts.check && opts.input) fail('INVALID_ARGUMENT', '--check does not accept an input file.', 2);
  if (opts.title && opts.title.length > 300) fail('INVALID_ARGUMENT', '--title must be at most 300 characters.', 2);
  return opts;
}
function check() {
  if (!supportedNode()) fail('UNSUPPORTED_NODE', `Node.js ${process.version} is unsupported; use >=22.22.2 on 22.x, >=24.15.0 on 24.x, or >=26.`);
  for (const name of ['jszip', 'dompurify', 'jsdom', 'cfb', 'officecrypto-tool']) {
    try { require.resolve(name); }
    catch { fail('DEPENDENCY_MISSING', `Dependency ${name} is missing; run npm ci in the project directory.`); }
  }
  const engineDir = path.resolve(__dirname, '../node_modules/@briank-dev/pptx-to-html');
  let enginePackage;
  try { enginePackage = JSON.parse(fs.readFileSync(path.join(engineDir, 'package.json'), 'utf8')); }
  catch { fail('DEPENDENCY_MISSING', 'Native engine metadata is missing or invalid; run npm ci.'); }
  if (enginePackage.name !== '@briank-dev/pptx-to-html' || typeof enginePackage.version !== 'string')
    fail('DEPENDENCY_MISSING', 'Native engine metadata is invalid; run npm ci.');
  if (!fs.existsSync(path.join(engineDir, 'pptx2html_wasm.js')) ||
      !fs.statSync(path.join(engineDir, 'pptx2html_wasm.js')).size) fail('DEPENDENCY_MISSING', 'Native engine module is missing; run npm ci.');
  const wasm = path.join(engineDir, 'pptx2html_wasm_bg.wasm');
  if (!fs.existsSync(wasm)) fail('DEPENDENCY_MISSING', 'WASM engine is missing; run npm ci.');
  const wasmHeader = Buffer.alloc(8);
  const wasmFile = fs.openSync(wasm, 'r');
  try { fs.readSync(wasmFile, wasmHeader, 0, 8, 0); } finally { fs.closeSync(wasmFile); }
  if (wasmHeader.toString('hex') !== '0061736d01000000') fail('DEPENDENCY_MISSING', 'WASM engine header is invalid.');
  const root = path.resolve(__dirname, '..');
  for (const file of ['src/input-package.js', 'src/normalize-package.js', 'src/exporter.js', 'src/viewer.css', 'src/viewer.js', 'src/native-cli-worker.cjs']) {
    if (!fs.existsSync(path.join(root, file)) || !fs.statSync(path.join(root, file)).size)
      fail('DEPENDENCY_MISSING', `${file} is missing or empty; restore the converter package.`);
  }
  const exporter = fs.readFileSync(path.join(root, 'src/exporter.js'), 'utf8');
  for (const marker of ['/* VIEWER_CSS */ null', '/* VIEWER_SCRIPT */ null', '/* VIEWER_HASH */ null']) {
    if (exporter.split(marker).length !== 2) fail('EXPORTER_ERROR', 'Exporter build marker is missing.');
  }
  const jsdom = require('jsdom');
  if (typeof jsdom.JSDOM !== 'function' || typeof require('dompurify') !== 'function') fail('DEPENDENCY_MISSING', 'HTML sanitizer is unavailable.');
  return {ok: true, engine: 'native', version: VERSION, enginePackage: enginePackage.name,
    engineVersion: enginePackage.version, node: process.version, wasmBytes: fs.statSync(wasm).size,
    externalAppsRequired: false};
}
function deepCheck() {
  return new Promise((resolve, reject) => {
    const worker = new Worker(path.resolve(__dirname, '../src/native-cli-worker.cjs'), {workerData: {mode: 'check'}});
    let settled = false;
    const finish = (error) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      process.off('SIGINT', interrupt);
      worker.terminate().finally(() => error ? reject(error) : resolve());
    };
    const interrupt = () => { const e = new Error('Interrupted.'); e.code = 'CANCELLED'; e.status = 130; finish(e); };
    const timer = setTimeout(() => { const e = new Error('Native self-check timed out.'); e.code = 'CHECK_TIMEOUT'; finish(e); }, 15000);
    process.once('SIGINT', interrupt);
    worker.on('message', message => {
      if (message.type === 'check') finish();
      if (message.type === 'error') { const e = new Error(message.message); e.code = message.code; finish(e); }
    });
    worker.on('error', finish);
    worker.on('exit', code => { if (!settled) { const e = new Error(`Self-check worker stopped (${code}).`); e.code = 'WORKER_EXIT'; finish(e); } });
  });
}
async function readPassword(timeoutSeconds) {
  if (process.stdin.isTTY) fail('PASSWORD_INPUT', '--password-stdin requires piped input.');
  let timer;
  let interrupt;
  const abort = new Promise((_, reject) => {
    timer = setTimeout(() => {
      const e = new Error(`Password input exceeded ${timeoutSeconds} seconds.`);
      e.code = 'TIMEOUT';
      process.stdin.destroy();
      reject(e);
    }, timeoutSeconds * 1000);
    interrupt = () => {
      const e = new Error('Interrupted.'); e.code = 'CANCELLED'; e.status = 130;
      process.stdin.destroy();
      reject(e);
    };
    process.once('SIGINT', interrupt);
  });
  try {
    const reading = (async () => {
      const chunks = []; let length = 0;
      for await (const chunk of process.stdin) {
        length += chunk.length;
        if (length > 1024) fail('PASSWORD_INPUT', 'Password input exceeds 1024 bytes.');
        chunks.push(chunk);
      }
      let password = Buffer.concat(chunks).toString('utf8');
      password = password.replace(/\r?\n$/, '');
      if (password.length > 255) fail('PASSWORD_INPUT', 'Password must be at most 255 characters.');
      return password;
    })();
    return await Promise.race([reading, abort]);
  } finally { clearTimeout(timer); process.off('SIGINT', interrupt); }
}
async function paths(opts) {
  const input = path.resolve(opts.input);
  const output = path.resolve(opts.output || path.join(path.dirname(input), path.parse(input).name + '.html'));
  const source = await fsp.stat(input);
  if (!source.isFile()) fail('INVALID_INPUT', 'Input must be a regular file.');
  if (source.size < 1 || source.size > MAX_INPUT) fail('INPUT_LIMIT', 'Input must be nonempty and at most 64 MiB.');
  let target;
  try { target = await fsp.stat(output); } catch (e) { if (e.code !== 'ENOENT') throw e; }
  if (input === output || (target && source.dev === target.dev && source.ino === target.ino)) fail('PATH_COLLISION', 'Input and output refer to the same file.');
  // Resolve parent aliases even when the target does not exist yet.
  const inputReal = await fsp.realpath(input);
  const outputParentReal = await fsp.realpath(path.dirname(output));
  if (inputReal === path.join(outputParentReal, path.basename(output))) fail('PATH_COLLISION', 'Input and output refer to the same file.');
  if (target && !opts.overwrite) fail('OUTPUT_EXISTS', 'Output exists; use --overwrite to replace it.');
  return {input, output};
}
async function atomicWrite(output, html, overwrite) {
  const bytes = Buffer.from(html, 'utf8');
  if (bytes.length > MAX_OUTPUT) fail('OUTPUT_LIMIT', 'HTML output exceeds 256 MiB.');
  const temp = path.join(path.dirname(output), `.pptx-html-${process.pid}-${crypto.randomUUID()}.tmp`);
  try {
    await fsp.writeFile(temp, bytes, {flag: 'wx', mode: 0o600});
    if (overwrite) await fsp.rename(temp, output);
    else { await fsp.link(temp, output); await fsp.unlink(temp); }
  } finally { await fsp.unlink(temp).catch(() => {}); }
}
function convert(bytes, opts, title) {
  return new Promise((resolve, reject) => {
    const worker = new Worker(path.resolve(__dirname, '../src/native-cli-worker.cjs'), {
      workerData: {bytes: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
        password: opts.password || '', slides: opts.slides, includeHidden: opts.includeHidden, title}
    });
    opts.password = '';
    let settled = false;
    const end = (error, result) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      process.off('SIGINT', interrupt);
      worker.terminate().finally(() => error ? reject(error) : resolve(result));
    };
    const interrupt = () => { const e = new Error('Interrupted.'); e.code = 'CANCELLED'; e.status = 130; end(e); };
    const timer = setTimeout(() => { const e = new Error(`Conversion exceeded ${opts.timeout} seconds.`); e.code = 'TIMEOUT'; end(e); }, opts.timeout * 1000);
    process.once('SIGINT', interrupt);
    worker.on('message', message => {
      if (message.type === 'progress') process.stderr.write(message.message + '\n');
      else if (message.type === 'result') end(null, message);
      else if (message.type === 'error') { const e = new Error(message.message); e.code = message.code; end(e); }
    });
    worker.on('error', error => end(error));
    worker.on('exit', code => { if (!settled) { const e = new Error(`Conversion worker stopped (${code}).`); e.code = 'WORKER_EXIT'; end(e); } });
  });
}
async function main() {
  let opts = {json: process.argv.includes('--json')};
  try {
    opts = parse(process.argv.slice(2));
    if (opts.help) { process.stdout.write(HELP); return; }
    const health = check();
    if (opts.check) { await deepCheck(); process.stdout.write(opts.json ? JSON.stringify(health) + '\n' : `Native converter ready (${health.node}; WASM ${health.wasmBytes} bytes; no office app or browser required).\n`); return; }
    const {input, output} = await paths(opts);
    opts.password = opts.passwordStdin ? await readPassword(opts.timeout) : '';
    const bytes = await fsp.readFile(input);
    if (bytes.length > MAX_INPUT) fail('INPUT_LIMIT', 'Input exceeds 64 MiB.');
    const converted = await convert(bytes, opts, opts.title || path.parse(input).name);
    opts.password = '';
    await atomicWrite(output, converted.html, opts.overwrite);
    const result = {ok: true, engine: 'native', version: VERSION, output, slideCount: converted.slideCount,
      width: converted.width, height: converted.height, warnings: converted.warnings, sourceSlides: converted.sourceSlides};
    for (const warning of result.warnings || []) {
      process.stderr.write(`Warning [${warning.code}]: ${warning.message}\n`);
    }
    process.stdout.write(opts.json ? JSON.stringify(result) + '\n' : `Created ${output} (${result.slideCount} slides).\n`);
  } catch (error) {
    opts.password = '';
    const status = error.status || (error.code === 'INVALID_ARGUMENT' ? 2 : 1);
    const code = error.code || 'CONVERSION_FAILED';
    const message = String(error.message || 'Conversion failed.');
    if (opts.json) process.stdout.write(JSON.stringify({ok: false, engine: 'native', version: VERSION, code, error: message}) + '\n');
    else process.stderr.write(`${code}: ${message}\n`);
    process.exitCode = status;
  }
}
main();
