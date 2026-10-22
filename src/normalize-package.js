/* Normalize valid OPC absolute relationship targets for the PPTX converter. */
async function normalizePptxPackage(bytes) {
  'use strict';

  const MAX_INPUT_BYTES = 64 * 1024 * 1024;
  const MAX_UNCOMPRESSED_BYTES = 256 * 1024 * 1024;
  const MAX_RELS_BYTES = 4 * 1024 * 1024;
  const MAX_FILES = 20000;

  if (!(bytes instanceof Uint8Array) || bytes.byteLength === 0 || bytes.byteLength > MAX_INPUT_BYTES) {
    throw new Error('PPTX 입력은 64 MiB 이하의 유효한 바이트 배열이어야 합니다.');
  }
  if (typeof JSZip === 'undefined' || typeof JSZip.loadAsync !== 'function') {
    throw new Error('PPTX 압축 해제 모듈을 사용할 수 없습니다.');
  }

  let zip;
  try {
    zip = await JSZip.loadAsync(bytes, { checkCRC32: false, createFolders: false });
  } catch (error) {
    const startsZip = bytes[0] === 0x50 && bytes[1] === 0x4b;
    const encryptedZip = /encrypted zip/i.test(String(error && error.message));
    const failure = new Error(encryptedZip
      ? 'ZIP 자체에 암호가 설정된 형식은 지원하지 않습니다. 사용이 허용된 일반 PPTX 사본을 선택해 주세요.'
      : startsZip
        ? 'PPTX의 압축 구조를 끝까지 읽지 못했습니다. 파일이 불완전하거나 현재 엔진이 지원하지 않는 구조일 수 있습니다. 정상적인 PPTX 사본인지 확인해 주세요.'
        : '브라우저가 받은 데이터를 일반 PPTX 압축 구조로 읽을 수 없습니다. NASCA+SD 등 회사 DRM으로 보호된 파일일 수도 있습니다. 보호된 경우 회사에서 허용한 일반 PPTX 사본이나 공식 연동이 필요합니다.');
    failure.code = encryptedZip ? 'UNSUPPORTED_ENCRYPTION' : startsZip ? 'INVALID_PACKAGE' : 'UNKNOWN_CONTAINER';
    throw failure;
  }

  const names = Object.keys(zip.files);
  if (names.length > MAX_FILES) {
    throw new Error('PPTX 내부 파일이 20,000개를 초과합니다.');
  }

  let totalUncompressed = 0;
  const relationships = [];
  for (const name of names) {
    const entry = zip.files[name];
    if (entry.dir) continue;
    // JSZip 3.10.1 stores the ZIP central-directory size here after loadAsync.
    const declaredSize = entry._data && entry._data.uncompressedSize;
    if (!Number.isSafeInteger(declaredSize) || declaredSize < 0) {
      throw new Error('PPTX 내부 파일 크기를 확인할 수 없습니다.');
    }
    totalUncompressed += declaredSize;
    if (totalUncompressed > MAX_UNCOMPRESSED_BYTES) {
      throw new Error('PPTX 압축 해제 크기가 256 MiB를 초과합니다.');
    }
    if (/\.rels$/i.test(name)) {
      if (declaredSize > MAX_RELS_BYTES) {
        throw new Error('PPTX 관계 파일이 4 MiB를 초과합니다.');
      }
      relationships.push({ name: name, entry: entry });
    }
  }

  if (!zip.file('[Content_Types].xml') || !zip.file('_rels/.rels') || !zip.file('ppt/presentation.xml')) {
    const failure = new Error('압축 파일은 읽었지만 변환에 필요한 PPTX 문서 구성을 찾지 못했습니다. 지원되는 일반 PPTX 사본인지 확인해 주세요.');
    failure.code = 'UNSUPPORTED_PACKAGE';
    throw failure;
  }

  function relativeTarget(relsName, absoluteTarget) {
    const suffixAt = absoluteTarget.search(/[?#]/);
    const path = suffixAt < 0 ? absoluteTarget : absoluteTarget.slice(0, suffixAt);
    const suffix = suffixAt < 0 ? '' : absoluteTarget.slice(suffixAt);
    const targetParts = [];
    for (const part of path.slice(1).split('/')) {
      if (!part || part === '.') continue;
      if (part === '..') {
        if (!targetParts.length) throw new Error('PPTX 관계 경로가 패키지 밖을 가리킵니다.');
        targetParts.pop();
      } else {
        targetParts.push(part);
      }
    }
    if (!targetParts.length) throw new Error('PPTX 관계 대상 경로가 비어 있습니다.');

    const ownerMatch = /^(?:(.*)\/)?_rels\/([^/]+)\.rels$/i.exec(relsName);
    if (relsName !== '_rels/.rels' && !ownerMatch) throw new Error('PPTX 관계 파일의 위치가 올바르지 않습니다.');
    const ownerDir = ownerMatch && ownerMatch[1] ? ownerMatch[1].split('/').filter(Boolean) : [];
    let common = 0;
    while (common < ownerDir.length && common < targetParts.length && ownerDir[common] === targetParts[common]) {
      common++;
    }
    const relative = ownerDir.slice(common).map(function () { return '..'; }).concat(targetParts.slice(common));
    return relative.join('/') + suffix;
  }

  // JSZip validates the complete archive, including valid ZIPs with a prefix.
  // Re-emit such an archive for the WASM reader instead of guessing from two bytes.
  let changed = bytes[0] !== 0x50 || bytes[1] !== 0x4b || bytes[2] !== 0x03 || bytes[3] !== 0x04;
  for (const item of relationships) {
    let source;
    try {
      source = await item.entry.async('string');
    } catch (error) {
      throw new Error('PPTX 관계 파일을 읽을 수 없습니다: ' + item.name);
    }
    const normalized = source.replace(
      /<(?:[A-Za-z_][\w.-]*:)?Relationship\b(?:[^"'>]|"[^"]*"|'[^']*')*>/g,
      function (tag) {
        const mode = /\sTargetMode\s*=\s*(['"])(.*?)\1/s.exec(tag);
        if (mode && mode[2].trim().toLowerCase() === 'external') return tag;
        return tag.replace(/(\sTarget\s*=\s*)(['"])(.*?)\2/s, function (attribute, prefix, quote, target) {
          if (!target.startsWith('/') || target.startsWith('//')) return attribute;
          const relative = relativeTarget(item.name, target);
          changed = true;
          return prefix + quote + relative + quote;
        });
      }
    );
    if (normalized !== source) {
      zip.file(item.name, normalized, {
        date: item.entry.date,
        comment: item.entry.comment,
        unixPermissions: item.entry.unixPermissions,
        dosPermissions: item.entry.dosPermissions,
        createFolders: false
      });
    }
  }

  if (!changed) return bytes;

  let output;
  try {
    output = await zip.generateAsync({
      type: 'uint8array',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
      streamFiles: true
    });
  } catch (error) {
    throw new Error('정규화된 PPTX 파일을 만들 수 없습니다.');
  }
  if (output.byteLength > MAX_INPUT_BYTES) {
    throw new Error('정규화된 PPTX 크기가 64 MiB를 초과합니다.');
  }
  return output;
}
