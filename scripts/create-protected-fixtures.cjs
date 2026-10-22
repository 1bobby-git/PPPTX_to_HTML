#!/usr/bin/env node
'use strict';

// All generated inputs are benign local test data. The checked-in reference
// pair is encrypted by msoffcrypto-tool, independently of the JS implementation.
// Normal fixture generation needs Node only. To refresh the reference pair:
//   python -m pip install --target /tmp/pptx-reference-crypto msoffcrypto-tool==5.4.2
//   PYTHONPATH=/tmp/pptx-reference-crypto node scripts/create-protected-fixtures.cjs --refresh-independent
// Regeneration intentionally changes salts/IVs. The manifest records exact
// bytes and the independent password/integrity verification performed then.

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const { spawnSync } = require('node:child_process');
const CFB = require('cfb');
const JSZip = require('jszip');
const officeCrypto = require('officecrypto-tool');

const root = path.resolve(__dirname, '..');
const outDir = path.join(root, 'tests/fixtures');
const referenceDir = path.join(root, 'tests/reference');
const PASSWORD = 'Offline!암호-2026🔐';
const sha256 = bytes => crypto.createHash('sha256').update(bytes).digest('hex');

function writeCFB(name, entries) {
  const cfb = CFB.utils.cfb_new();
  for (const [entry, bytes] of entries) CFB.utils.cfb_add(cfb, entry, bytes);
  fs.writeFileSync(path.join(outDir, name), CFB.write(cfb, { type: 'buffer' }));
}

function mutateCFB(input, name, mutate) {
  const cfb = CFB.read(input, { type: 'buffer' });
  mutate(cfb);
  fs.writeFileSync(path.join(outDir, name), CFB.write(cfb, { type: 'buffer' }));
}

function refreshIndependent() {
  fs.mkdirSync(referenceDir, { recursive: true });
  fs.copyFileSync(path.join(outDir, 'complex.pptx'), path.join(referenceDir, 'complex-plain.pptx'));
  const python = `
import hashlib, importlib.metadata, io, json, pathlib, sys
import msoffcrypto

folder = pathlib.Path(sys.argv[1])
password = sys.argv[2]
plain = (folder / 'complex-plain.pptx').read_bytes()
output = io.BytesIO()
msoffcrypto.OfficeFile(io.BytesIO(plain)).encrypt(password, output)
encrypted = output.getvalue()
office = msoffcrypto.OfficeFile(io.BytesIO(encrypted))
office.load_key(password=password, verify_password=True)
decrypted = io.BytesIO()
office.decrypt(decrypted, verify_integrity=True)
assert decrypted.getvalue() == plain, 'Independent encryption/decryption must round-trip byte-for-byte'
(folder / 'complex-agile.pptx').write_bytes(encrypted)
manifest = {
    'generator': 'msoffcrypto-tool',
    'generatorVersion': importlib.metadata.version('msoffcrypto-tool'),
    'source': 'Benign complex.pptx generated locally by scripts/create-fixtures.cjs',
    'encryption': 'ECMA-376 Agile AES-256-CBC, SHA-512, 100000 password iterations',
    'password': password,
    'plaintextFile': 'complex-plain.pptx',
    'encryptedFile': 'complex-agile.pptx',
    'plaintextSha256': hashlib.sha256(plain).hexdigest(),
    'encryptedSha256': hashlib.sha256(encrypted).hexdigest(),
    'independentVerification': 'msoffcrypto.OfficeFile.load_key(verify_password=True), decrypt(verify_integrity=True), byte-for-byte equality',
    'refresh': 'PYTHONPATH=/tmp/pptx-reference-crypto node scripts/create-protected-fixtures.cjs --refresh-independent'
}
(folder / 'manifest.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + '\\n', encoding='utf-8')
print('Created and independently verified Agile reference: ' + manifest['encryptedSha256'])
`;
  const result = spawnSync(process.env.PYTHON || 'python', ['-c', python, referenceDir, PASSWORD], {
    encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'],
  });
  if (result.status !== 0) throw new Error(`Independent reference generation failed: ${result.error || result.stderr}`);
  process.stdout.write(result.stdout);
}

async function generateFixtures() {
  fs.mkdirSync(outDir, { recursive: true });
  const plain = fs.readFileSync(path.join(outDir, 'complex.pptx'));
  const manifest = JSON.parse(fs.readFileSync(path.join(referenceDir, 'manifest.json'), 'utf8'));
  const independentPlain = fs.readFileSync(path.join(referenceDir, manifest.plaintextFile));
  const independentAgile = fs.readFileSync(path.join(referenceDir, manifest.encryptedFile));
  assert.equal(sha256(independentPlain), manifest.plaintextSha256, 'independent reference plaintext hash');
  assert.equal(sha256(independentAgile), manifest.encryptedSha256, 'independent reference ciphertext hash');
  assert.equal(manifest.generator, 'msoffcrypto-tool');
  assert.equal(manifest.generatorVersion, '5.4.2');
  assert.equal(manifest.password, PASSWORD);

  // Standard generation is a compatibility fixture. Agile interoperability is
  // proved by the independent reference pair, not a same-library round-trip.
  const standard = officeCrypto.encrypt(plain, { password: PASSWORD, type: 'standard' });
  fs.writeFileSync(path.join(outDir, 'protected-standard.pptx'), standard);
  const agile = officeCrypto.encrypt(plain, { password: PASSWORD });
  fs.writeFileSync(path.join(outDir, 'protected-agile-js.pptx'), agile);

  // ZIP permits a prepended byte sequence. JSZip must validate the package and
  // compensate for its offsets; stripping at a guessed PK signature is unsafe.
  const prefixed = Buffer.concat([Buffer.from('BENIGN-PREFIX\r\n', 'ascii'), plain]);
  const zip = await JSZip.loadAsync(prefixed);
  assert.ok(zip.file('ppt/presentation.xml'), 'the prefixed test fixture must remain a valid ZIP');
  fs.writeFileSync(path.join(outDir, 'prefixed.pptx'), prefixed);
  fs.writeFileSync(path.join(outDir, 'broken-zip.pptx'), plain.subarray(0, plain.length - 100));
  const unrelatedZip = new JSZip();
  unrelatedZip.file('benign.txt', 'A valid ZIP archive containing no PowerPoint package.');
  fs.writeFileSync(path.join(outDir, 'unrelated-zip.pptx'), await unrelatedZip.generateAsync({ type: 'nodebuffer' }));
  fs.writeFileSync(path.join(outDir, 'unknown.pptx'), Buffer.from('Benign unrelated bytes. No document format is implied.'));
  fs.writeFileSync(path.join(outDir, 'malformed-cfb.pptx'), Buffer.from('d0cf11e0a1b11ae100000000', 'hex'));

  // Only container markers are required to exercise honest format reporting;
  // these are intentionally not real protected or legacy business documents.
  writeCFB('drm-marker.pptx', [['\u0006DataSpaces/DataSpaceInfo/DRMEncryptedDataSpace', Buffer.from('benign test marker')]]);
  writeCFB('drm-content-marker.pptx', [['\u0006DataSpaces/DataSpaceMap', Buffer.from('DRMEncryptedDataSpace', 'utf16le')]]);
  writeCFB('legacy-marker.pptx', [['PowerPoint Document', Buffer.alloc(32)]]);
  writeCFB('unknown-cfb.pptx', [['Unrelated Stream', Buffer.from('benign test marker')]]);
  writeCFB('incomplete-encryption.pptx', [['EncryptionInfo', Buffer.from([4, 0, 4, 0, 0x40, 0, 0, 0])]]);

  mutateCFB(independentAgile, 'tampered-package.pptx', cfb => {
    const entry = CFB.find(cfb, '/EncryptedPackage');
    assert.ok(entry && entry.content.length > 4096);
    entry.content = Buffer.from(entry.content);
    entry.content[100] ^= 1;
  });
  mutateCFB(independentAgile, 'declared-too-large.pptx', cfb => {
    const entry = CFB.find(cfb, '/EncryptedPackage');
    entry.content = Buffer.from(entry.content);
    entry.content.writeUInt32LE(64 * 1024 * 1024 + 1, 0);
  });
  mutateCFB(independentAgile, 'tampered-hmac.pptx', cfb => {
    const entry = CFB.find(cfb, '/EncryptionInfo');
    const prefix = Buffer.from(entry.content.subarray(0, 8));
    let xml = Buffer.from(entry.content.subarray(8)).toString('utf8');
    assert.match(xml, /encryptedHmacValue="[^"]+"/);
    xml = xml.replace(/encryptedHmacValue="([^"]+)"/, (_, value) => {
      const changed = Buffer.from(value, 'base64');
      changed[0] ^= 1;
      return `encryptedHmacValue="${changed.toString('base64')}"`;
    });
    entry.content = Buffer.concat([prefix, Buffer.from(xml)]);
    entry.size = entry.content.length;
  });
  mutateCFB(independentAgile, 'agile-renamed-prefix.pptx', cfb => {
    const entry = CFB.find(cfb, '/EncryptionInfo');
    const prefix = Buffer.from(entry.content.subarray(0, 8));
    const xml = Buffer.from(entry.content.subarray(8)).toString('utf8')
      .replace(/xmlns:p=/g, 'xmlns:passwordKey=').replace(/p:encryptedKey/g, 'passwordKey:encryptedKey');
    assert.match(xml, /passwordKey:encryptedKey/);
    entry.content = Buffer.concat([prefix, Buffer.from(xml)]);
    entry.size = entry.content.length;
  });
  mutateCFB(independentAgile, 'excessive-iterations.pptx', cfb => {
    const entry = CFB.find(cfb, '/EncryptionInfo');
    const prefix = Buffer.from(entry.content.subarray(0, 8));
    const xml = Buffer.from(entry.content.subarray(8)).toString('utf8').replace(/spinCount="\d+"/, 'spinCount="1000001"');
    assert.match(xml, /spinCount="1000001"/);
    entry.content = Buffer.concat([prefix, Buffer.from(xml)]);
    entry.size = entry.content.length;
  });
  mutateCFB(independentAgile, 'unsupported-cipher.pptx', cfb => {
    const entry = CFB.find(cfb, '/EncryptionInfo');
    const prefix = Buffer.from(entry.content.subarray(0, 8));
    const xml = Buffer.from(entry.content.subarray(8)).toString('utf8').replace(/cipherChaining="ChainingModeCBC"/g, 'cipherChaining="ChainingModeCFB"');
    assert.match(xml, /ChainingModeCFB/);
    entry.content = Buffer.concat([prefix, Buffer.from(xml)]);
    entry.size = entry.content.length;
  });
  console.log('Created protected, prefixed, malformed, DRM-marker, and legacy-marker PPTX fixtures');
}

if (require.main === module) {
  Promise.resolve().then(() => {
    if (process.argv.includes('--refresh-independent')) refreshIndependent();
    return generateFixtures();
  }).catch(error => { console.error(error); process.exitCode = 1; });
}

module.exports = { generateFixtures, PASSWORD };
