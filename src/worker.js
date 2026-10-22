/* Runs only in the embedded Blob Worker. The build prepends the pinned WASM glue. */
'use strict';

const MAX_INPUT_BYTES = 64 * 1024 * 1024;

function selectedSlides(value, total) {
  if (!String(value || '').trim()) return new Uint32Array();
  const selected = new Set();
  for (const part of String(value).split(',')) {
    const match = /^\s*(\d+)(?:\s*-\s*(\d+))?\s*$/.exec(part);
    if (!match) throw new Error('슬라이드 범위는 1,3-5 형식으로 입력해 주세요.');
    const first = Number(match[1]);
    const last = match[2] ? Number(match[2]) : first;
    if (!Number.isSafeInteger(first) || !Number.isSafeInteger(last) || first < 1 || last < first || last > total) {
      throw new Error('슬라이드 번호는 1부터 ' + total + ' 사이에서 지정해 주세요.');
    }
    for (let index = first; index <= last; index += 1) selected.add(index);
  }
  return new Uint32Array([...selected].sort((a, b) => a - b));
}

function summarizeDiagnostics(json) {
  let items;
  try { items = JSON.parse(json || '[]'); } catch { return []; }
  if (!Array.isArray(items)) return [];
  const grouped = new Map();
  const names = {
    CHART: '일부 차트의 모양이나 데이터 표현이 원본과 다를 수 있습니다.',
    TABLE: '일부 표 스타일은 원본과 다르게 표시될 수 있습니다.',
    SMARTART: '일부 SmartArt는 미리보기 이미지 또는 대체 표시로 변환되었습니다.',
    DIAGRAM: '일부 다이어그램은 미리보기 이미지 또는 대체 표시로 변환되었습니다.',
    OLE: '삽입된 개체는 실행되지 않으며 미리보기 또는 대체 표시로 변환됩니다.',
    MATH: '일부 수식은 대체 표시로 변환되었습니다.',
    MEDIA: '일부 오디오·동영상은 브라우저에서 재생할 수 없습니다.',
    TIMING: '애니메이션과 전환 효과는 정적인 슬라이드로 저장됩니다.',
    IMAGE: '일부 이미지 형식 또는 외부 이미지를 표시하지 못할 수 있습니다.',
    DRAWINGML: '일부 도형·효과는 원본과 다르게 표시될 수 있습니다.',
    ACTION: '일부 슬라이드 동작 또는 링크는 변환되지 않습니다.',
    FONT: '일부 글꼴은 이 기기에 설치된 대체 글꼴로 표시될 수 있습니다.'
  };
  for (const item of items) {
    const code = String(item.code || 'CONVERSION_NOTICE');
    const part = String(item.location?.part_name || '');
    // Notes and package inventories are not rendering failures. Never copy their raw text.
    if (/NOTE|COMMENT|AUTHOR/.test(code) || /(?:notes|comments|Authors)/i.test(part)) continue;
    if (code === 'OOXML_PART_UNSUPPORTED' && (part.endsWith('/') || !/\/(slides|charts|diagrams)\//.test(part))) continue;
    if (code === 'OOXML_ELEMENT_UNSUPPORTED' && !/\/(slides|charts|diagrams)\//.test(part)) continue;
    if (code === 'OOXML_ELEMENT_UNSUPPORTED' && /^(?:a:chOff|a:chExt|a:masterClrMapping)$/.test(item.location?.qualified_element_name || '')) continue;
    const family = Object.keys(names).find((name) => code.includes(name)) || 'OTHER';
    if (!grouped.has(family)) grouped.set(family, {
      code: family === 'OTHER' ? 'UNSUPPORTED_FEATURES' : code,
      message: names[family] || '일부 PPTX 기능이 완전히 재현되지 않았습니다. 미리보기를 원본과 비교해 주세요.',
      slides: new Set()
    });
    const index = item.location?.slide_index;
    if (Number.isInteger(index) && index >= 0) grouped.get(family).slides.add(index + 1);
  }
  return [...grouped.values()].map(({code, message, slides}) => ({
    code,
    message: message + (slides.size ? ' (슬라이드 ' + [...slides].sort((a,b) => a-b).slice(0, 12).join(', ') + (slides.size > 12 ? ' 외' : '') + ')' : '')
  }));
}

function friendlyError(error) {
  const message = String(error?.message || error || '');
  if (/[가-힣]/.test(message)) return message;
  if (/encrypt|password|compound|OLE2/i.test(message)) return '암호가 설정된 PPTX는 지원하지 않습니다. PowerPoint에서 암호를 해제하고 다시 저장해 주세요.';
  if (/memory|allocation|too (?:large|many)|limit|out of bounds/i.test(message)) return '파일을 처리할 수 있는 메모리 또는 크기 한도를 넘었습니다. 슬라이드나 이미지 크기를 줄인 뒤 다시 시도해 주세요.';
  if (/zip|archive|invalid|corrupt|xml|missing|not found|parse/i.test(message)) return 'PPTX 파일을 읽지 못했습니다. 손상되었거나 지원되지 않는 형식입니다. PowerPoint에서 .pptx로 다시 저장해 주세요.';
  return '파일을 변환하지 못했습니다. PowerPoint에서 .pptx로 다시 저장하거나 다른 파일로 확인해 주세요.';
}

self.onmessage = async function (event) {
  const {type, id, bytes, options = {}} = event.data || {};
  if (type !== 'convert') return;
  let info;
  let result;
  try {
    if (!(bytes instanceof ArrayBuffer) || !bytes.byteLength || bytes.byteLength > MAX_INPUT_BYTES) throw new Error('파일 크기는 0보다 크고 64 MiB 이하여야 합니다.');
    let data = new Uint8Array(bytes);
    if (data[0] === 0xd0 && data[1] === 0xcf) throw new Error('암호가 설정된 PPTX 또는 예전 .ppt 형식입니다. 암호를 해제한 .pptx 파일로 다시 저장해 주세요.');
    if (data[0] !== 0x50 || data[1] !== 0x4b) throw new Error('올바른 PPTX 파일이 아닙니다. 확장자만 바꾼 파일은 변환할 수 없습니다.');
    self.postMessage({type: 'progress', id, message: 'PPTX 파일과 이미지·차트 연결을 확인하는 중입니다…'});
    data = await normalizePptxPackage(data);
    self.postMessage({type: 'progress', id, message: '내장 변환 엔진을 준비하는 중입니다…'});
    const binary = atob(EMBEDDED_WASM_BASE64);
    const wasmBytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) wasmBytes[index] = binary.charCodeAt(index);
    initSync({module: wasmBytes});
    self.postMessage({type: 'progress', id, message: '슬라이드 구성과 크기를 확인하는 중입니다…'});
    info = get_presentation_info(data);
    const width = info.widthPx;
    const height = info.heightPx;
    if (!info.slideCount) throw new Error('변환할 슬라이드가 없습니다.');
    if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0 || width > 30000 || height > 30000) throw new Error('지원하지 않는 슬라이드 크기입니다. PowerPoint에서 슬라이드 크기를 다시 지정해 주세요.');
    const range = selectedSlides(options.slideRange, info.slideCount);
    self.postMessage({type: 'progress', id, message: '텍스트·이미지·도형을 HTML로 변환하는 중입니다…'});
    result = convert_with_options_metadata(data, true, !!options.includeHidden, range, 1);
    const html = result.html;
    const slideCount = (html.match(/class="slide"/g) || []).length;
    if (!slideCount) throw new Error('선택한 범위에 표시할 슬라이드가 없습니다. 숨긴 슬라이드 포함 설정을 확인해 주세요.');
    self.postMessage({type: 'result', id, html, width, height, slideCount, diagnostics: summarizeDiagnostics(result.diagnosticsJson)});
  } catch (error) {
    self.postMessage({type: 'error', id, message: friendlyError(error)});
  } finally {
    result?.free();
    info?.free();
  }
};
