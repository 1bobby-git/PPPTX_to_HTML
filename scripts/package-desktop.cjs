#!/usr/bin/env node
'use strict';

// Build-time only: bundle the pinned Electron runtime and the already-installed
// production dependency tree. The resulting application never runs this script.
const fs = require('node:fs');
const fsp = require('node:fs/promises');
const path = require('node:path');
const os = require('node:os');
const crypto = require('node:crypto');
const {Readable, Transform} = require('node:stream');
const {pipeline} = require('node:stream/promises');
const JSZip = require('jszip');

const root = path.resolve(__dirname, '..');
const generatedBy = 'PPPTX_to_HTML/scripts/package-desktop.cjs';
const fixedDate = new Date('2000-01-01T00:00:00Z');
const desktopFiles = [
  'desktop/main.cjs', 'desktop/preload.cjs', 'desktop/controller.cjs',
  'desktop/index.html', 'desktop/styles.css', 'desktop/renderer.js',
  'desktop/runtime.json', 'desktop/package.json'
];
const conversionFiles = [
  'src/native-cli-worker.cjs', 'src/conversion-core.js', 'src/input-package.js',
  'src/normalize-package.js', 'src/chart-data.js', 'src/theme-resolver.js',
  'src/exporter.js', 'src/viewer.js', 'src/viewer.css'
];
const vendorFiles = [
  'BUILD-INFO.json', 'CHART-DEPTH.patch', 'LICENSE', 'README.md', 'SHA256SUMS',
  'briank-dev-pptx-to-html-2.1.0-ppptx.1.tgz'
].map(name => 'vendor/pptx2html-engine/' + name);
const readJson = async file => JSON.parse(await fsp.readFile(file, 'utf8'));
const sha256 = bytes => crypto.createHash('sha256').update(bytes).digest('hex');
const token = () => crypto.randomBytes(8).toString('hex');

function parse(argv) {
  const options = {platform: 'win32', arch: 'x64', outputDir: path.join(root, 'dist'), unpackedOnly: false};
  for (let i = 0; i < argv.length; i++) {
    const key = argv[i];
    if (key === '--unpacked-only') options.unpackedOnly = true;
    else if (key === '--help' || key === '-h') options.help = true;
    else if (['--platform', '--arch', '--runtime-archive', '--output-dir'].includes(key)) {
      const value = argv[++i];
      if (!value || value.startsWith('--')) throw new Error(key + ' requires a value.');
      if (key === '--platform') options.platform = value;
      if (key === '--arch') options.arch = value;
      if (key === '--runtime-archive') options.runtimeArchive = path.resolve(value);
      if (key === '--output-dir') options.outputDir = path.resolve(value);
    } else throw new Error('Unknown option: ' + key);
  }
  if (!['win32', 'linux'].includes(options.platform) || options.arch !== 'x64') {
    throw new Error('Supported desktop targets: --platform win32|linux --arch x64.');
  }
  return options;
}

function safeName(name) {
  if (!name || name.includes('\\') || name.includes('\0') || name.startsWith('/') || /^[A-Za-z]:/.test(name)) return false;
  return !name.replace(/\/$/, '').split('/').some(part => !part || part === '.' || part === '..' || part.includes(':'));
}

async function verifyArchive(file, runtime) {
  const stat = await fsp.stat(file);
  if (!stat.isFile() || stat.size !== runtime.bytes) throw new Error('Electron archive size mismatch: ' + file);
  const bytes = await fsp.readFile(file);
  if (sha256(bytes) !== runtime.sha256) throw new Error('Electron archive SHA256 mismatch: ' + file);
  return bytes;
}

async function downloadRuntime(config, runtime, supplied) {
  if (supplied) return supplied;
  const cache = path.join(os.tmpdir(), 'ppptx-electron-runtime', config.electronVersion);
  await fsp.mkdir(cache, {recursive: true});
  const cached = path.join(cache, runtime.archive);
  try {
    await fsp.access(cached);
    return cached; // The exact cached bytes are verified before extraction too.
  } catch (error) { if (error.code !== 'ENOENT') throw error; }
  const temp = cached + '.' + token() + '.download';
  const digest = crypto.createHash('sha256');
  let total = 0;
  try {
    console.error('Downloading the pinned Electron build: ' + runtime.archive);
    const response = await fetch(runtime.url, {signal: AbortSignal.timeout(180000)});
    if (!response.ok || !response.body) throw new Error('Electron download failed: HTTP ' + response.status);
    const counter = new Transform({
      transform(chunk, encoding, callback) {
        total += chunk.length;
        if (total > runtime.bytes) return callback(new Error('Electron download exceeds the pinned size.'));
        digest.update(chunk);
        callback(null, chunk);
      }
    });
    await pipeline(Readable.fromWeb(response.body), counter, fs.createWriteStream(temp, {flags: 'wx', mode: 0o600}));
    if (total !== runtime.bytes || digest.digest('hex') !== runtime.sha256) {
      throw new Error('Downloaded Electron archive did not match its pinned size/SHA256.');
    }
    await fsp.rename(temp, cached);
    return cached;
  } finally { await fsp.rm(temp, {force: true}); }
}

async function extractRuntime(bytes, directory) {
  const zip = await JSZip.loadAsync(bytes, {checkCRC32: false, createFolders: false});
  let total = 0;
  const entries = Object.values(zip.files).sort((a, b) => a.name.localeCompare(b.name, 'en'));
  if (entries.length > 10000) throw new Error('Unexpected number of Electron archive entries.');
  for (const entry of entries) {
    if (!safeName(entry.name) || (entry.unsafeOriginalName && entry.unsafeOriginalName !== entry.name)) {
      throw new Error('Unsafe Electron archive path.');
    }
    const permissions = Number(entry.unixPermissions || 0);
    const type = permissions & 0o170000;
    if (type && type !== 0o100000 && type !== 0o040000) throw new Error('Unexpected special file in Electron archive.');
    const target = path.join(directory, ...entry.name.split('/'));
    if (entry.dir) { await fsp.mkdir(target, {recursive: true}); continue; }
    const size = entry._data?.uncompressedSize;
    if (!Number.isSafeInteger(size) || size < 0 || size > 512 * 1024 * 1024) throw new Error('Unexpected Electron entry size.');
    total += size;
    if (total > 2 * 1024 * 1024 * 1024) throw new Error('Electron runtime exceeds its extraction budget.');
    await fsp.mkdir(path.dirname(target), {recursive: true});
    const contents = await entry.async('nodebuffer');
    await fsp.writeFile(target, contents, {flag: 'wx', mode: (permissions & 0o777) || 0o644});
  }
}

async function copyRegular(source, target) {
  const stat = await fsp.lstat(source);
  if (!stat.isFile()) throw new Error('Only regular application files may be packaged: ' + source);
  await fsp.mkdir(path.dirname(target), {recursive: true});
  await fsp.copyFile(source, target);
  await fsp.chmod(target, stat.mode & 0o777);
}

async function copyTree(source, target) {
  const entries = await fsp.readdir(source, {withFileTypes: true});
  for (const entry of entries.sort((a, b) => a.name.localeCompare(b.name, 'en'))) {
    // Nested dependencies are enumerated from package-lock.json independently.
    if (['node_modules', '.git', '.DS_Store'].includes(entry.name)) continue;
    const from = path.join(source, entry.name), to = path.join(target, entry.name);
    if (entry.isDirectory()) await copyTree(from, to);
    else if (entry.isFile()) await copyRegular(from, to);
    else throw new Error('Unexpected link or special file in a runtime package: ' + from);
  }
}

async function copyDependencies(lock, pkg, appRoot) {
  const packages = [];
  for (const [relative, record] of Object.entries(lock.packages).sort(([a], [b]) => a.localeCompare(b, 'en'))) {
    if (!relative.startsWith('node_modules/') || record.dev) continue;
    if (!safeName(relative) || record.link) throw new Error('Unsafe or linked runtime dependency: ' + relative);
    const source = path.join(root, relative);
    let installed;
    try { installed = await readJson(path.join(source, 'package.json')); }
    catch (error) {
      if (record.optional && error.code === 'ENOENT') continue;
      throw new Error('Production dependency is missing or invalid; run npm ci: ' + relative);
    }
    if (installed.version !== record.version) throw new Error('Dependency version differs from the lockfile: ' + relative);
    await copyTree(source, path.join(appRoot, relative));
    packages.push({name: installed.name, version: installed.version, license: installed.license || 'See package notices', path: relative});
  }
  for (const name of Object.keys(pkg.dependencies)) {
    if (!packages.some(item => item.path === 'node_modules/' + name)) throw new Error('Missing direct runtime dependency: ' + name);
  }
  return packages;
}

async function verifyEngine(appRoot) {
  const vendorRoot = path.join(appRoot, 'vendor/pptx2html-engine');
  const info = await readJson(path.join(vendorRoot, 'BUILD-INFO.json'));
  const engineRoot = path.join(appRoot, 'node_modules/@briank-dev/pptx-to-html');
  const pkg = await readJson(path.join(engineRoot, 'package.json'));
  if (pkg.name !== info.package.name || pkg.version !== info.package.version) throw new Error('Installed engine metadata differs from the pinned build.');
  for (const [name, record] of Object.entries(info.artifacts)) {
    if (!safeName(name)) throw new Error('Unsafe engine artifact name.');
    const bytes = await fsp.readFile(path.join(engineRoot, name));
    if (bytes.length !== record.patchedBytes || sha256(bytes) !== record.patchedSha256) throw new Error('Engine artifact hash mismatch: ' + name);
  }
  const sums = (await fsp.readFile(path.join(vendorRoot, 'SHA256SUMS'), 'utf8')).trim().split(/\r?\n/);
  for (const line of sums) {
    const match = /^([a-f0-9]{64})  ([^/\\]+)$/.exec(line);
    if (!match || !safeName(match[2]) || sha256(await fsp.readFile(path.join(vendorRoot, match[2]))) !== match[1]) {
      throw new Error('Vendored engine checksum mismatch.');
    }
  }
  return {name: pkg.name, version: pkg.version, license: pkg.license,
    sourceCommit: info.source.commit, wasmBytes: info.artifacts['pptx2html_wasm_bg.wasm'].patchedBytes,
    wasmSha256: info.artifacts['pptx2html_wasm_bg.wasm'].patchedSha256,
    buildInfo: 'resources/app/vendor/pptx2html-engine/BUILD-INFO.json'};
}

async function filesIn(directory, prefix = '') {
  const result = [];
  for (const entry of (await fsp.readdir(directory, {withFileTypes: true})).sort((a, b) => a.name.localeCompare(b.name, 'en'))) {
    const relative = prefix ? prefix + '/' + entry.name : entry.name;
    if (entry.isDirectory()) result.push(...await filesIn(path.join(directory, entry.name), relative));
    else if (entry.isFile()) result.push(relative);
    else throw new Error('Unexpected link or special file in the desktop distribution.');
  }
  return result;
}

async function publishDirectory(staging, target) {
  let backup;
  let stat;
  try {
    stat = await fsp.lstat(target);
  } catch (error) { if (error.code !== 'ENOENT') throw error; }
  if (stat) {
    if (!stat.isDirectory() || stat.isSymbolicLink()) throw new Error('Refusing to replace an unexpected output path: ' + target);
    let previous;
    try { previous = await readJson(path.join(target, 'DESKTOP-MANIFEST.json')); }
    catch { throw new Error('Refusing to replace a directory without a valid desktop build manifest.'); }
    if (previous.generatedBy !== generatedBy) throw new Error('Refusing to replace a directory not created by this packager.');
    backup = target + '.previous-' + token();
    await fsp.rename(target, backup);
  }
  try { await fsp.rename(staging, target); }
  catch (error) { if (backup) await fsp.rename(backup, target); throw error; }
  if (backup) await fsp.rm(backup, {recursive: true, force: true});
}

async function createArchive(directory, output, folderName) {
  const zip = new JSZip();
  for (const relative of await filesIn(directory)) {
    const source = path.join(directory, ...relative.split('/'));
    const stat = await fsp.stat(source);
    // Each file is opened lazily when JSZip starts consuming that entry, avoiding
    // thousands of open handles and keeping large Electron files off the heap.
    const content = Readable.from((async function* () {
      for await (const chunk of fs.createReadStream(source)) yield chunk;
    })());
    zip.file(folderName + '/' + relative, content, {
      binary: true, date: fixedDate, createFolders: false, unixPermissions: 0o100000 | (stat.mode & 0o777)
    });
  }
  const temp = output + '.' + token() + '.tmp';
  const digest = crypto.createHash('sha256');
  let bytes = 0;
  const counter = new Transform({transform(chunk, encoding, callback) {
    bytes += chunk.length; digest.update(chunk); callback(null, chunk);
  }});
  try {
    await pipeline(zip.generateNodeStream({streamFiles: true, compression: 'DEFLATE', compressionOptions: {level: 6}, platform: 'UNIX'}),
      counter, fs.createWriteStream(temp, {flags: 'wx'}));
    await fsp.rename(temp, output);
    return {bytes, sha256: digest.digest('hex')};
  } finally { await fsp.rm(temp, {force: true}); }
}

async function main(argv = process.argv.slice(2)) {
  const options = parse(argv);
  if (options.help) {
    console.log('Usage: node scripts/package-desktop.cjs [--platform win32|linux] [--arch x64] [--runtime-archive FILE] [--output-dir DIR] [--unpacked-only]');
    return;
  }
  const pkg = await readJson(path.join(root, 'package.json'));
  const lock = await readJson(path.join(root, 'package-lock.json'));
  const desktop = await readJson(path.join(root, 'desktop/package.json'));
  const config = await readJson(path.join(root, 'desktop/runtime.json'));
  if (!/^\d+\.\d+\.\d+(?:-[A-Za-z0-9.-]+)?$/.test(pkg.version) ||
      pkg.version !== lock.version || pkg.version !== lock.packages?.['']?.version || pkg.version !== desktop.version ||
      desktop.main !== 'desktop/main.cjs') {
    throw new Error('Root, lockfile and desktop package versions must match.');
  }
  const target = options.platform + '-' + options.arch;
  const runtime = config.platforms?.[target];
  if (config.schemaVersion !== 1 || !safeName(config.productName) || config.productName.includes('/') ||
      !runtime || !/^[a-f0-9]{64}$/.test(runtime.sha256) ||
      !Number.isSafeInteger(runtime.bytes) || runtime.bytes <= 0 || runtime.bytes > 250 * 1024 * 1024 ||
      !/^\d+\.\d+\.\d+$/.test(config.electronVersion) ||
      runtime.archive !== `electron-v${config.electronVersion}-${target}.zip` ||
      runtime.url !== `https://github.com/electron/electron/releases/download/v${config.electronVersion}/${runtime.archive}` ||
      !safeName(runtime.executable) || !safeName(runtime.packagedExecutable)) {
    throw new Error('Invalid pinned Electron runtime manifest.');
  }
  const stem = `PPPTX_to_HTML_Desktop_v${pkg.version}_${target}`;
  await fsp.mkdir(options.outputDir, {recursive: true});
  const staging = await fsp.mkdtemp(path.join(options.outputDir, stem + '.staging-'));
  const directory = path.join(options.outputDir, stem);
  try {
    const runtimeArchive = await downloadRuntime(config, runtime, options.runtimeArchive);
    await extractRuntime(await verifyArchive(runtimeArchive, runtime), staging);
    await fsp.rm(path.join(staging, 'resources/default_app.asar'), {force: true});
    await fsp.rename(path.join(staging, runtime.executable), path.join(staging, runtime.packagedExecutable));
    if (options.platform === 'win32') {
      await fsp.rename(path.join(staging, 'electron_wer.dll'), path.join(staging, 'PPPTX_to_HTML_wer.dll'));
    } else {
      await fsp.chmod(path.join(staging, runtime.packagedExecutable), 0o755);
    }
    const appRoot = path.join(staging, 'resources/app');
    for (const relative of [...desktopFiles, ...conversionFiles, ...vendorFiles, 'LICENSE', 'docs/DESKTOP.md']) {
      await copyRegular(path.join(root, relative), path.join(appRoot, relative));
    }
    const packages = await copyDependencies(lock, pkg, appRoot);
    await fsp.writeFile(path.join(appRoot, 'package.json'), JSON.stringify({...desktop, dependencies: pkg.dependencies, engines: pkg.engines}, null, 2) + '\n');
    const engine = await verifyEngine(appRoot);
    await copyRegular(path.join(root, 'docs/DESKTOP.md'), path.join(staging, 'README.md'));
    const manifest = {
      schemaVersion: 1, generatedBy, version: pkg.version, platform: options.platform, arch: options.arch,
      executable: runtime.packagedExecutable,
      runtime: {electronVersion: config.electronVersion, nodeVersion: config.nodeVersion, chromiumVersion: config.chromiumVersion,
        source: runtime.url, sourceArchive: runtime.archive, sourceBytes: runtime.bytes, sourceSha256: runtime.sha256,
        license: 'LICENSE', thirdPartyNotices: 'LICENSES.chromium.html'},
      engine, packages, files: (await filesIn(staging)).length + 1,
      note: 'Includes the Electron runtime and production dependencies. No separate Node.js, PowerPoint or CLI installation is required.'
    };
    await fsp.writeFile(path.join(staging, 'DESKTOP-MANIFEST.json'), JSON.stringify(manifest, null, 2) + '\n');
    await publishDirectory(staging, directory);
    const archive = options.unpackedOnly ? null : path.join(options.outputDir, stem + '.zip');
    const archiveInfo = archive ? await createArchive(directory, archive, config.productName) : null;
    console.log(JSON.stringify({directory, executable: path.join(directory, runtime.packagedExecutable), archive,
      version: pkg.version, electronVersion: config.electronVersion, nodeVersion: config.nodeVersion,
      engineVersion: engine.version, runtimeSha256: runtime.sha256, files: manifest.files, packages: packages.length,
      ...(archiveInfo || {})}, null, 2));
  } finally { await fsp.rm(staging, {recursive: true, force: true}); }
}

module.exports = {parse, main};
if (require.main === module) main().catch(error => {console.error(error.message); process.exitCode = 1;});
