'use strict';

// Browserify embeds these modules in the offline worker. No file or network APIs are used.
const CFB = require('cfb');
const crypto = require('crypto');
const {Parser} = require('xml2js');
const agile = require('officecrypto-tool/src/crypto/ecma376_agile');
const standard = require('officecrypto-tool/src/crypto/ecma376_standard');

const MAX_INPUT_BYTES = 64 * 1024 * 1024;
const MAX_INFO_BYTES = 64 * 1024;
const MAX_SPIN_COUNT = 1000000;
const CFB_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const ENC_NS = 'http://schemas.microsoft.com/office/2006/encryption';
const PASSWORD_NS = 'http://schemas.microsoft.com/office/2006/keyEncryptor/password';
const HASH_SIZES = {sha1: 20, sha256: 32, sha384: 48, sha512: 64};
const ERROR_CODES = new Set([
  'PASSWORD_REQUIRED', 'INCORRECT_PASSWORD', 'DRM_PROTECTED', 'LEGACY_PPT',
  'UNSUPPORTED_ENCRYPTION', 'INVALID_CONTAINER', 'INPUT_LIMIT', 'UNKNOWN_CONTAINER'
]);

// MS-OFFCRYPTO 2.3.4.13 and 2.3.4.14; also used by officecrypto-tool's Agile writer.
const BLOCK_KEYS = {
  verifierInput: Buffer.from([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79]),
  verifierHash: Buffer.from([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e]),
  packageKey: Buffer.from([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6]),
  hmacKey: Buffer.from([0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6]),
  hmacValue: Buffer.from([0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33])
};

function failure(code, message, details) {
  const error = new Error(message);
  error.code = code;
  error.details = details;
  return error;
}

function invalid(details) {
  return failure('INVALID_CONTAINER', '보호된 문서의 구조 또는 무결성 정보를 확인하지 못했습니다. 원본을 다시 저장한 뒤 시도해 주세요.', details);
}

function unsupported(details, message) {
  return failure('UNSUPPORTED_ENCRYPTION', message || '이 문서의 암호화 방식은 지원하지 않습니다. 사용이 허용된 일반 PPTX 사본을 선택해 주세요.', details);
}

function checkPassword(password, details) {
  if (typeof password !== 'string' || !password.length) {
    throw failure('PASSWORD_REQUIRED', '암호로 보호된 Office 문서입니다. 문서를 열 때 사용하는 암호를 입력해 주세요.', details);
  }
  if (password.length > 255) {
    throw failure('INPUT_LIMIT', '암호는 255자 이하로 입력해 주세요.', details);
  }
}

function equalBytes(left, right) {
  if (left.length !== right.length) return false;
  let difference = 0;
  for (let i = 0; i < left.length; i += 1) difference |= left[i] ^ right[i];
  return difference === 0;
}

function wipe(...buffers) {
  for (const buffer of buffers) if (buffer && typeof buffer.fill === 'function') buffer.fill(0);
}

function parseContainer(bytes) {
  if (bytes.length < 512) throw invalid('CFB · 잘린 파일 헤더');
  const input = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const major = input.readUInt16LE(26);
  const sectorShift = input.readUInt16LE(30);
  if (input.readUInt16LE(28) !== 0xfffe || input.readUInt16LE(32) !== 6 ||
      !((major === 3 && sectorShift === 9) || (major === 4 && sectorShift === 12)) ||
      input.length % (1 << sectorShift) !== 0) throw invalid('CFB · 지원되지 않거나 손상된 헤더');
  let cfb;
  try { cfb = CFB.read(input, {type: 'buffer'}); } catch { throw invalid('CFB · 컨테이너 읽기 실패'); }
  if (!Array.isArray(cfb.FullPaths) || !Array.isArray(cfb.FileIndex) ||
      cfb.FullPaths.length !== cfb.FileIndex.length || cfb.FileIndex[0]?.type !== 5) {
    throw invalid('CFB · 디렉터리 구조 오류');
  }
  const root = cfb.FullPaths[0];
  const entries = cfb.FileIndex.map((entry, index) => ({entry, path: cfb.FullPaths[index]}));
  const drmName = /^(?:drmencrypteddataspace|drmencryptedtransform|\x09drmdataspace|\x09drmtransform|\x09lzxdrmdataspace)$/i;
  for (const {entry, path} of entries) {
    if (typeof path !== 'string') throw invalid('CFB · 디렉터리 이름 오류');
    if (path.split('/').some((name) => drmName.test(name))) {
      throw failure('DRM_PROTECTED', '조직의 DRM·IRM으로 보호된 문서입니다. 브라우저에서 읽을 수 있는 일반 PPTX 사본을 회사에서 허용한 절차로 준비하거나 공식 연동을 사용해야 합니다.', 'CFB · Microsoft IRM 데이터 공간');
    }
    // Some producers keep the transform reference only inside a DataSpaces stream.
    if (entry.type === 2 && /(?:^|\/)\x06DataSpaces\//i.test(path) && entry.content) {
      if (entry.content.length > MAX_INFO_BYTES) throw invalid('CFB · 데이터 공간 정보 크기 초과');
      const markerText = Buffer.from(entry.content).toString('utf16le');
      if (/DRMEncrypted(?:DataSpace|Transform)|Microsoft\.Metadata\.DRMTransform|\x09(?:LZX)?DRM(?:DataSpace|Transform)/i.test(markerText)) {
        throw failure('DRM_PROTECTED', '조직의 DRM·IRM으로 보호된 문서입니다. 브라우저에서 읽을 수 있는 일반 PPTX 사본을 회사에서 허용한 절차로 준비하거나 공식 연동을 사용해야 합니다.', 'CFB · Microsoft IRM 변환 정보');
      }
    }
  }
  if (entries.some(({entry}) => entry.type === 2 && /^PowerPoint Document$/i.test(entry.name))) {
    throw failure('LEGACY_PPT', '예전 PowerPoint PPT 형식입니다. 이 변환기에는 .pptx 형식으로 저장된 사본이 필요합니다.', 'CFB · PowerPoint Document');
  }
  const stream = (name) => {
    const matches = entries.filter(({path}) => path.toLowerCase() === (root + name).toLowerCase());
    if (matches.length > 1) throw invalid('CFB · 중복된 암호화 스트림');
    if (!matches.length) return null;
    const entry = matches[0].entry;
    if (entry.type !== 2 || !entry.content || !Number.isSafeInteger(entry.size) || entry.size !== entry.content.length) {
      throw invalid('CFB · 암호화 스트림 길이 오류');
    }
    return Buffer.isBuffer(entry.content) ? entry.content : Buffer.from(entry.content);
  };
  const info = stream('EncryptionInfo');
  const encryptedPackage = stream('EncryptedPackage');
  if (!info && !encryptedPackage) {
    throw failure('UNKNOWN_CONTAINER', '이 복합 문서 형식은 지원하지 않습니다. 다른 문서 형식이거나 보안 프로그램으로 보호된 파일일 수 있습니다.', 'CFB · 인식되지 않은 문서');
  }
  if (!info || !encryptedPackage) throw invalid('CFB · 필수 암호화 스트림 누락');
  if (info.length < 8 || info.length > MAX_INFO_BYTES) throw invalid('CFB · 암호화 정보 크기 오류');
  if (encryptedPackage.length < 24 || encryptedPackage.length > MAX_INPUT_BYTES) throw invalid('CFB · 암호화 데이터 크기 오류');
  const clearLength = encryptedPackage.readUInt32LE(0);
  if (encryptedPackage.readUInt32LE(4) !== 0 || clearLength > MAX_INPUT_BYTES) {
    throw failure('INPUT_LIMIT', '복호화한 파일의 크기가 64 MiB 한도를 넘습니다.', 'CFB · 복호화 크기 한도 초과');
  }
  if (!clearLength || encryptedPackage.length !== 8 + Math.ceil(clearLength / 16) * 16) {
    throw invalid('CFB · 선언된 크기와 암호화 데이터 길이 불일치');
  }
  return {info, encryptedPackage, clearLength};
}

function parseStandard(info) {
  const details = 'Office Standard · AES-128 · SHA-1';
  if (info.length < 12) throw invalid('Office Standard · 잘린 암호화 헤더');
  const headerSize = info.readUInt32LE(8);
  if (headerSize < 32 || headerSize > 1024 || headerSize % 2 || info.length !== 12 + headerSize + 72) {
    throw invalid('Office Standard · 암호화 헤더 길이 오류');
  }
  if (info.readUInt32LE(4) !== 0x24 || info.readUInt32LE(12) !== 0x24 ||
      info.readUInt32LE(20) !== 0x660e || info.readUInt32LE(24) !== 0x8004 ||
      info.readUInt32LE(28) !== 128 || info.readUInt32LE(32) !== 24) {
    throw unsupported('Office Standard · AES-128/SHA-1 외의 암호화 설정');
  }
  if (info.readUInt32LE(16) !== 0 || info.readUInt32LE(36) !== 0 || info.readUInt32LE(40) !== 0) {
    throw invalid('Office Standard · 예약 필드 오류');
  }
  const offset = 12 + headerSize;
  if (info.readUInt32LE(offset) !== 16 || info.readUInt32LE(offset + 36) !== 20) {
    throw invalid('Office Standard · 암호 확인 정보 크기 오류');
  }
  return {
    details,
    salt: info.subarray(offset + 4, offset + 20),
    verifier: info.subarray(offset + 20, offset + 36),
    verifierHash: info.subarray(offset + 40, offset + 72)
  };
}

function children(node, name, namespace = ENC_NS) {
  const found = [];
  for (const value of Object.values(node || {})) {
    if (!Array.isArray(value)) continue;
    for (const child of value) {
      if (child?.$ns?.local === name && child.$ns.uri === namespace) found.push(child);
    }
  }
  return found;
}

function oneChild(node, name, namespace = ENC_NS) {
  const matches = children(node, name, namespace);
  if (matches.length !== 1) throw invalid('Office Agile · 필수 암호화 요소 누락 또는 중복');
  return matches[0];
}

function attribute(node, name) {
  const item = node?.$?.[name];
  if (!item || item.uri !== '' || typeof item.value !== 'string') throw invalid('Office Agile · 필수 암호화 속성 누락');
  return item.value;
}

function numberAttribute(node, name) {
  const value = attribute(node, name);
  if (!/^\d{1,10}$/.test(value) || !Number.isSafeInteger(Number(value))) throw invalid('Office Agile · 숫자 속성 오류');
  return Number(value);
}

function binaryAttribute(node, name, expectedLengths) {
  const text = attribute(node, name).replace(/[\t\r\n ]/g, '');
  if (!text.length || text.length > 128 || text.length % 4 || !/^[A-Za-z0-9+/]*={0,2}$/.test(text)) {
    throw invalid('Office Agile · 암호화 값 인코딩 오류');
  }
  const bytes = Buffer.from(text, 'base64');
  if (bytes.toString('base64') !== text || !expectedLengths.includes(bytes.length)) {
    throw invalid('Office Agile · 암호화 값 길이 오류');
  }
  return bytes;
}

function cipherParameters(node) {
  const cipherAlgorithm = attribute(node, 'cipherAlgorithm');
  const cipherChaining = attribute(node, 'cipherChaining');
  const hashAlgorithm = attribute(node, 'hashAlgorithm').toLowerCase().replace(/-/g, '');
  const keyBits = numberAttribute(node, 'keyBits');
  const blockSize = numberAttribute(node, 'blockSize');
  const saltSize = numberAttribute(node, 'saltSize');
  const hashSize = numberAttribute(node, 'hashSize');
  if (cipherAlgorithm !== 'AES' || cipherChaining !== 'ChainingModeCBC' ||
      !Object.hasOwn(HASH_SIZES, hashAlgorithm) || ![128, 192, 256].includes(keyBits) ||
      blockSize !== 16 || saltSize !== 16) {
    throw unsupported('Office Agile · 지원되지 않는 알고리즘 또는 블록·솔트 크기');
  }
  if (hashSize !== HASH_SIZES[hashAlgorithm]) throw invalid('Office Agile · 해시 길이 불일치');
  return {cipherAlgorithm, cipherChaining, hashAlgorithm, keyBits, blockSize, saltSize, hashSize,
    saltValue: binaryAttribute(node, 'saltValue', [saltSize])};
}

async function parseAgile(info) {
  if (info.readUInt32LE(4) !== 0x40) throw invalid('Office Agile · 암호화 헤더 플래그 오류');
  const xml = info.subarray(8).toString('utf8');
  if (/<!DOCTYPE|<!ENTITY/i.test(xml) || xml.includes('\0')) throw invalid('Office Agile · 지원되지 않는 암호화 XML');
  let root;
  try {
    root = await new Parser({xmlns: true, explicitArray: true, explicitRoot: false, strict: true}).parseStringPromise(xml);
  } catch { throw invalid('Office Agile · 암호화 XML 읽기 실패'); }
  if (root?.$ns?.local !== 'encryption' || root.$ns.uri !== ENC_NS) throw invalid('Office Agile · 암호화 XML 네임스페이스 오류');
  const packageParams = cipherParameters(oneChild(root, 'keyData'));
  const encryptors = children(oneChild(root, 'keyEncryptors'), 'keyEncryptor');
  const passwordEncryptors = encryptors.filter((node) => node?.$?.uri?.value === PASSWORD_NS);
  if (passwordEncryptors.length !== 1) {
    throw unsupported('Office Agile · 단일 암호 키 없음', '인증서 전용 또는 지원되지 않는 키 방식으로 보호된 문서입니다. 권한이 있는 앱의 내보내기를 사용해 주세요.');
  }
  const passwordNode = oneChild(passwordEncryptors[0], 'encryptedKey', PASSWORD_NS);
  const key = cipherParameters(passwordNode);
  key.spinCount = numberAttribute(passwordNode, 'spinCount');
  if (key.spinCount > MAX_SPIN_COUNT) {
    throw unsupported('Office Agile · 암호 반복 횟수 한도 초과', '이 문서의 암호 반복 횟수가 지원 한도인 1,000,000회를 넘습니다. 사용이 허용된 일반 PPTX 사본을 선택해 주세요.');
  }
  key.encryptedVerifierHashInput = binaryAttribute(passwordNode, 'encryptedVerifierHashInput', [16]);
  key.encryptedVerifierHashValue = binaryAttribute(passwordNode, 'encryptedVerifierHashValue', [Math.ceil(key.hashSize / 16) * 16]);
  key.encryptedKeyValue = binaryAttribute(passwordNode, 'encryptedKeyValue', [Math.ceil(packageParams.keyBits / 8 / 16) * 16]);
  const integrityNodes = children(root, 'dataIntegrity');
  if (integrityNodes.length > 1) throw invalid('Office Agile · 중복된 무결성 정보');
  let integrity = null;
  if (integrityNodes.length) {
    integrity = {
      // The spec uses saltSize; Office and the upstream writer use hashSize.
      encryptedHmacKey: binaryAttribute(integrityNodes[0], 'encryptedHmacKey', [16, Math.ceil(packageParams.hashSize / 16) * 16]),
      encryptedHmacValue: binaryAttribute(integrityNodes[0], 'encryptedHmacValue', [Math.ceil(packageParams.hashSize / 16) * 16])
    };
  }
  return {package: packageParams, key, integrity,
    details: 'Office Agile · AES-' + packageParams.keyBits + '-CBC · ' + packageParams.hashAlgorithm.toUpperCase()};
}

function decryptAgile(parsed, encryptedPackage, password) {
  const keyInfo = parsed.key;
  const packageInfo = parsed.package;
  const secrets = [];
  const remember = (value) => { secrets.push(value); return value; };
  const passwordKey = (blockKey) => remember(agile.convertPasswordToKey(password, keyInfo.hashAlgorithm,
    keyInfo.saltValue, keyInfo.spinCount, keyInfo.keyBits, blockKey));
  const decryptKeyValue = (value, blockKey) => remember(agile.crypt(false, 'AES', 'ChainingModeCBC',
    passwordKey(blockKey), keyInfo.saltValue, value));
  try {
    const verifier = decryptKeyValue(keyInfo.encryptedVerifierHashInput, BLOCK_KEYS.verifierInput);
    const expectedHash = decryptKeyValue(keyInfo.encryptedVerifierHashValue, BLOCK_KEYS.verifierHash);
    const actualHash = remember(agile.hash(keyInfo.hashAlgorithm, verifier));
    if (!equalBytes(actualHash, expectedHash.subarray(0, keyInfo.hashSize))) {
      throw failure('INCORRECT_PASSWORD', '입력한 암호가 문서의 암호와 일치하지 않습니다. 다시 입력해 주세요.', parsed.details);
    }
    const packageKey = decryptKeyValue(keyInfo.encryptedKeyValue, BLOCK_KEYS.packageKey).subarray(0, packageInfo.keyBits / 8);
    if (parsed.integrity) {
      // MS-OFFCRYPTO 2.3.4.14: authenticate the entire stream, including its 8-byte size.
      const keyIV = agile.createIV(packageInfo.hashAlgorithm, packageInfo.saltValue, 16, BLOCK_KEYS.hmacKey);
      const valueIV = agile.createIV(packageInfo.hashAlgorithm, packageInfo.saltValue, 16, BLOCK_KEYS.hmacValue);
      const hmacKey = remember(agile.crypt(false, 'AES', 'ChainingModeCBC', packageKey, keyIV, parsed.integrity.encryptedHmacKey));
      const expectedHmac = remember(agile.crypt(false, 'AES', 'ChainingModeCBC', packageKey, valueIV, parsed.integrity.encryptedHmacValue));
      const actualHmac = remember(crypto.createHmac(packageInfo.hashAlgorithm, hmacKey).update(encryptedPackage).digest());
      if (!equalBytes(actualHmac, expectedHmac.subarray(0, packageInfo.hashSize))) {
        throw invalid('Office Agile · 암호화 데이터 무결성 검사 실패');
      }
    }
    return agile.cryptPackage(false, 'AES', 'ChainingModeCBC', packageInfo.hashAlgorithm,
      16, packageInfo.saltValue, packageKey, encryptedPackage);
  } finally {
    wipe(...secrets);
    password = undefined; // Strings and upstream internal copies cannot be reliably erased in JavaScript.
  }
}

async function prepare(bytes, password) {
  try {
    if (bytes instanceof ArrayBuffer) bytes = new Uint8Array(bytes);
    if (!(bytes instanceof Uint8Array)) throw invalid('입력 · 바이트 배열 필요');
    if (!bytes.length || bytes.length > MAX_INPUT_BYTES) {
      throw failure('INPUT_LIMIT', '파일 크기는 0보다 크고 64 MiB 이하여야 합니다.', '입력 · 파일 크기 한도');
    }
    if (!CFB_SIGNATURE.every((value, index) => bytes[index] === value)) {
      const zip = bytes[0] === 0x50 && bytes[1] === 0x4b;
      return {bytes, format: zip ? 'zip' : 'unknown', description: zip ? 'ZIP 문서 패키지' : '형식 확인이 필요한 파일'};
    }
    const {info, encryptedPackage, clearLength} = parseContainer(bytes);
    const major = info.readUInt16LE(0);
    const minor = info.readUInt16LE(2);
    let output;
    let format;
    let description;
    if ((major === 3 || major === 4) && minor === 2) {
      const parsed = parseStandard(info);
      checkPassword(password, parsed.details);
      let key;
      try {
        key = standard.convertPasswordToKey(password, 0x660e, 0x8004, 24, 128, 16, parsed.salt);
        if (!standard.verifyKey(key, parsed.verifier, parsed.verifierHash)) {
          throw failure('INCORRECT_PASSWORD', '입력한 암호가 문서의 암호와 일치하지 않습니다. 다시 입력해 주세요.', parsed.details);
        }
        output = standard.decrypt(key, encryptedPackage);
      } finally { wipe(key); }
      format = 'password-standard';
      description = 'Office Standard 암호 보호 · AES-128';
    } else if (major === 4 && minor === 4) {
      const parsed = await parseAgile(info);
      checkPassword(password, parsed.details);
      output = decryptAgile(parsed, encryptedPackage, password);
      format = 'password-agile';
      description = 'Office Agile 암호 보호 · AES-' + parsed.package.keyBits + (parsed.integrity ? ' · 무결성 확인' : ' · 무결성 정보 없음');
    } else {
      throw unsupported('Office · 지원되지 않는 EncryptionInfo 버전');
    }
    if (!output || output.length !== clearLength) throw invalid('Office · 복호화 결과 길이 불일치');
    return {bytes: new Uint8Array(output.buffer, output.byteOffset, output.byteLength), format, description};
  } catch (error) {
    if (ERROR_CODES.has(error?.code)) throw error;
    throw invalid('Office · 보호된 문서 처리 실패');
  } finally {
    password = undefined;
  }
}

module.exports = {prepare, inspectChartData: require('./chart-data.js').inspect,
  resolveThemes: require('./theme-resolver.js').resolveThemes};
