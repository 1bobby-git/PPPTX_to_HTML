#!/usr/bin/env node
'use strict';

// Developer-only rebuild. All downloads, compiler files and Cargo caches stay in
// a new temporary directory; the end-user converter needs none of these tools.
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const crypto = require('node:crypto');
const {execFileSync} = require('node:child_process');
const vendor = path.resolve(__dirname, '../vendor/pptx2html-engine');
const info = JSON.parse(fs.readFileSync(path.join(vendor, 'BUILD-INFO.json'), 'utf8'));
const sha256 = bytes => crypto.createHash('sha256').update(bytes).digest('hex');

function verify(file, expected) {
  if (sha256(fs.readFileSync(file)) !== expected) throw new Error('SHA-256 mismatch: ' + file);
}
function run(command, args, cwd, env = process.env, capture = false) {
  return execFileSync(command, args, {cwd, env, maxBuffer: 32 * 1024 * 1024,
    stdio: capture ? ['ignore', 'pipe', 'pipe'] : 'inherit', encoding: capture ? 'utf8' : undefined});
}
async function download(item, cache) {
  const target = path.join(cache, item.file);
  if (!fs.existsSync(target)) {
    const response = await fetch(item.url, {signal: AbortSignal.timeout(120000)});
    if (!response.ok) throw new Error('Download failed: ' + item.url + ' (' + response.status + ')');
    const data = Buffer.from(await response.arrayBuffer());
    if (data.length > 200 * 1024 * 1024 || sha256(data) !== item.sha256) {
      throw new Error('Downloaded file failed its size/hash check: ' + item.file);
    }
    fs.writeFileSync(target, data, {flag: 'wx'});
  }
  verify(target, item.sha256);
  return target;
}
function extract(file, directory, strip = 0) {
  fs.mkdirSync(directory, {recursive: true});
  // Every archive is checked against its fixed official digest before extraction.
  run('tar', ['--extract', '--file', file, '--directory', directory,
    '--no-same-owner', '--strip-components=' + strip], directory);
}

function packageEngine(built, packageDirectory, outputDirectory) {
  for (const [name, record] of Object.entries(info.artifacts)) {
    verify(path.join(built, name), record.patchedSha256);
    fs.copyFileSync(path.join(built, name), path.join(packageDirectory, name));
  }
  verify(path.join(packageDirectory, 'LICENSE'), sha256(fs.readFileSync(path.join(vendor, 'LICENSE'))));
  const packageFile = path.join(packageDirectory, 'package.json');
  const pkg = JSON.parse(fs.readFileSync(packageFile, 'utf8'));
  if (pkg.name !== info.package.name || pkg.version !== info.upstreamPackage.version) {
    throw new Error('Expected the verified original npm package before packaging the patch.');
  }
  pkg.version = info.package.version;
  pkg.files = [...new Set([...pkg.files, 'BUILD-INFO.json', 'CHART-DEPTH.patch'])];
  pkg.ppptxPatch = 'Local chart namespace-depth fix; see BUILD-INFO.json.';
  fs.writeFileSync(packageFile, JSON.stringify(pkg, null, 2) + '\n');
  for (const name of ['BUILD-INFO.json', 'CHART-DEPTH.patch']) {
    fs.copyFileSync(path.join(vendor, name), path.join(packageDirectory, name));
  }
  fs.mkdirSync(outputDirectory, {recursive: true});
  const npmEnvironment = {...process.env, npm_config_cache: path.join(packageDirectory, '..', 'npm-cache'),
    npm_config_audit: 'false', npm_config_fund: 'false'};
  run('npm', ['pack', '--ignore-scripts', '--json', '--pack-destination', outputDirectory],
    packageDirectory, npmEnvironment);
  const archive = path.join(outputDirectory, info.package.archive);
  if (!fs.statSync(archive).isFile()) throw new Error('npm did not produce the expected archive.');
  const publishedSums = path.join(vendor, 'SHA256SUMS');
  if (fs.existsSync(publishedSums)) {
    const entry = fs.readFileSync(publishedSums, 'utf8').split('\n')
      .find(line => line.endsWith('  ' + info.package.archive));
    if (!entry) throw new Error('Recorded archive checksum is missing.');
    verify(archive, entry.split('  ')[0]);
  }
  for (const name of ['BUILD-INFO.json', 'CHART-DEPTH.patch', 'LICENSE']) {
    const source = path.join(vendor, name), target = path.join(outputDirectory, name);
    if (path.resolve(source) !== path.resolve(target)) fs.copyFileSync(source, target);
  }
  const entries = ['BUILD-INFO.json', 'CHART-DEPTH.patch', 'LICENSE', info.package.archive];
  const sums = entries.map(name => sha256(fs.readFileSync(path.join(outputDirectory, name))) + '  ' + name);
  fs.writeFileSync(path.join(outputDirectory, 'SHA256SUMS'), sums.join('\n') + '\n');
  return archive;
}

async function main() {
  const options = {};
  for (let index = 2; index < process.argv.length; index += 2) {
    const key = process.argv[index], value = process.argv[index + 1];
    if (!['--work-dir', '--cache-dir', '--output-dir'].includes(key) || !value || value.startsWith('--')) {
      throw new Error('Usage: node scripts/rebuild-engine.cjs [--work-dir NEW_DIRECTORY] [--cache-dir DIRECTORY] [--output-dir DIRECTORY]');
    }
    options[key] = path.resolve(value);
  }
  if (process.platform !== 'linux' || process.arch !== 'x64') {
    throw new Error('The recorded rebuild toolchain targets Linux x86_64. Runtime WASM remains platform independent.');
  }
  if (process.version !== info.build.nodeVersion) {
    throw new Error('Reproduce this archive with Node ' + info.build.nodeVersion + '; the recorded packager is npm ' + info.build.npmVersion + '.');
  }
  verify(path.join(vendor, info.patch.file), info.patch.sha256);
  const work = options['--work-dir'] || fs.mkdtempSync(path.join(os.tmpdir(), 'pptx-engine-rebuild-'));
  fs.mkdirSync(work, {recursive: true});
  if (fs.readdirSync(work).length) throw new Error('--work-dir must be a new or empty directory.');
  const cache = options['--cache-dir'] || path.join(work, 'downloads');
  const output = options['--output-dir'] || path.join(work, 'artifacts');
  fs.mkdirSync(cache, {recursive: true});
  const inputs = [info.source.archive, info.upstreamPackage, info.toolchain.wasmBindgen,
    ...info.toolchain.components];
  console.log('Downloading and verifying the pinned source and isolated compiler into ' + work);
  const files = await Promise.all(inputs.map(item => download(item, cache)));
  const source = path.join(work, 'source'), toolchain = path.join(work, 'toolchain');
  extract(files[0], source, 1);
  verify(path.join(source, 'Cargo.lock'), info.source.cargoLockSha256);
  for (const item of info.patch.files) verify(path.join(source, item.path), item.beforeSha256);
  run('patch', ['--batch', '--fuzz=0', '-p1', '-i', path.join(vendor, info.patch.file)], source);
  for (const item of info.patch.files) verify(path.join(source, item.path), item.afterSha256);
  for (let index = 0; index < info.toolchain.components.length; index++) {
    const archive = files[index + 3], unpacked = path.join(work, 'unpacked');
    extract(archive, unpacked);
    const directory = path.join(unpacked, path.basename(archive, '.tar.xz'));
    run('bash', [path.join(directory, 'install.sh'), '--prefix=' + toolchain, '--disable-ldconfig'], directory);
  }
  const bindingRoot = path.join(work, 'bindings');
  extract(files[2], bindingRoot, 1);
  const env = {...process.env, PATH: path.join(toolchain, 'bin') + path.delimiter + process.env.PATH,
    CARGO_HOME: path.join(work, 'cargo-home'), CARGO_TARGET_DIR: path.join(work, 'target'), CARGO_BUILD_JOBS: '2',
    RUSTC: path.join(toolchain, 'bin/rustc'), RUSTDOC: path.join(toolchain, 'bin/rustdoc'),
    RUSTFLAGS: info.build.rustFlags.map(flag => flag.replace('{workDirectory}', work)).join(' ')};
  for (const key of ['CARGO_ENCODED_RUSTFLAGS', 'RUSTC_WRAPPER', 'RUSTC_WORKSPACE_WRAPPER']) delete env[key];
  const cargo = path.join(toolchain, 'bin/cargo');
  if (run(cargo, ['--version'], source, env, true).trim() !== info.toolchain.cargo ||
      run(env.RUSTC, ['--version'], source, env, true).trim() !== info.toolchain.rustc ||
      run(path.join(bindingRoot, 'wasm-bindgen'), ['--version'], source, env, true).trim() !== 'wasm-bindgen ' + info.toolchain.wasmBindgen.version) {
    throw new Error('Unexpected compiler or binding tool version.');
  }
  console.log('Running the chart regression suite and building only the PPTX WASM adapter.');
  run(cargo, info.build.testArguments, source, env);
  run(cargo, info.build.cargoArguments, source, env);
  verify(path.join(source, 'Cargo.lock'), info.source.cargoLockSha256);
  const built = path.join(work, 'built');
  run(path.join(bindingRoot, 'wasm-bindgen'), [...info.build.wasmBindgenArguments, '--out-dir', built,
    path.join(work, 'target/wasm32-unknown-unknown/release/pptx2html_wasm.wasm')], source, env);
  const packageDirectory = path.join(work, 'package');
  extract(files[1], packageDirectory, 1);
  const archive = packageEngine(built, packageDirectory, output);
  console.log('Verified runtime hashes and wrote ' + archive);
  console.log('Compiler, source and caches remain isolated in ' + work);
}

module.exports = {packageEngine};
if (require.main === module) main().catch(error => {console.error(error.message); process.exitCode = 1;});
