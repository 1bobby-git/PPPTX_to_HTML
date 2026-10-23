/* Runs only in the embedded Blob Worker. The build prepends the pinned WASM glue. */
'use strict';

const MAX_INPUT_BYTES = 64 * 1024 * 1024;

function friendlyError(error) {
  const message = String(error?.message || error || '');
  if (/[가-힣]/.test(message)) return message;
  if (/encrypt|password|compound|OLE2/i.test(message)) return '문서의 암호 또는 보호 방식을 확인하지 못했습니다. 문서 열기 암호를 확인해 주세요. 회사 DRM 문서는 허용된 일반 PPTX 사본이나 공식 연동이 필요합니다.';
  if (/memory|allocation|too (?:large|many)|limit|out of bounds/i.test(message)) return '파일을 처리할 수 있는 메모리 또는 크기 한도를 넘었습니다. 슬라이드나 이미지 크기를 줄인 뒤 다시 시도해 주세요.';
  if (/zip|archive|invalid|corrupt|xml|missing|not found|parse/i.test(message)) return 'PPTX 파일을 읽지 못했습니다. 손상되었거나 지원되지 않는 형식일 수 있습니다. 정상적인 PPTX 사본인지 확인해 주세요.';
  return '파일을 변환하지 못했습니다. 지원되는 PPTX 사본이나 다른 파일로 확인해 주세요.';
}

self.onmessage = async function (event) {
  const {type, id, bytes, options = {}} = event.data || {};
  if (type !== 'convert') return;
  let inputFormat = '확인되지 않은 파일 구조';
  try {
    if (!(bytes instanceof ArrayBuffer) || !bytes.byteLength || bytes.byteLength > MAX_INPUT_BYTES) throw new Error('파일 크기는 0보다 크고 64 MiB 이하여야 합니다.');
    let data = new Uint8Array(bytes);
    self.postMessage({type: 'progress', id, message: '문서 형식과 보호 방식을 확인하는 중입니다…'});
    const prepared = await PptxInput.prepare(data, String(options.password || ''));
    options.password = '';
    data = prepared.bytes;
    inputFormat = prepared.description || prepared.format;
    self.postMessage({type: 'progress', id, message: 'PPTX 파일과 이미지·차트 연결을 확인하는 중입니다…'});
    data = await normalizePptxPackage(data);
    const themed = await PptxInput.resolveThemes(data, JSZip);
    data = themed.bytes;
    self.postMessage({type: 'progress', id, message: '내장 변환 엔진을 준비하는 중입니다…'});
    const binary = atob(EMBEDDED_WASM_BASE64);
    const wasmBytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) wasmBytes[index] = binary.charCodeAt(index);
    initSync({module: wasmBytes});
    self.postMessage({type: 'progress', id, message: '슬라이드 구성과 크기를 확인하는 중입니다…'});
    const rendered = PptxCore.render({get_presentation_info, convert_with_options_metadata}, data, {
      slideRange: options.slideRange, includeHidden: options.includeHidden,
      onProgress: message => self.postMessage({type: 'progress', id, message})
    });
    rendered.diagnostics.push(...themed.diagnostics);
    rendered.chartData = await PptxInput.inspectChartData(data, JSZip, rendered.chartFallbacks, rendered.html);
    delete rendered.chartFallbacks;
    self.postMessage(Object.assign({type: 'result', id}, rendered));
  } catch (error) {
    const code = typeof error?.code === 'string' ? error.code : 'CONVERSION_FAILED';
    const initial = bytes instanceof ArrayBuffer ? new Uint8Array(bytes).slice(0, 8) : new Uint8Array();
    const cfb = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1].every((value, index) => initial[index] === value);
    const header = cfb ? 'D0 CF 11 E0 A1 B1 1A E1 (CFB)' : initial[0] === 0x50 && initial[1] === 0x4b ? '50 4B (ZIP)' : '알려진 ZIP/CFB 표식 없음';
    const details = ['검사 결과: ' + code, '파일 크기: ' + (bytes?.byteLength || 0) + ' bytes', '파일 구조: ' + inputFormat];
    if (typeof error?.details === 'string' && error.details) details.push(error.details);
    details.push('형식 표식: ' + header);
    self.postMessage({type: 'error', id, code, message: friendlyError(error), details: details.join('\n')});
  } finally {
    options.password = '';
  }
};
