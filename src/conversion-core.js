/* Shared browser/Node contract for the pinned PPTX engine. No DOM or I/O. */
(function (global) {
  'use strict';
  const MAX_OUTPUT_BYTES = 256 * 1024 * 1024;
  function failure(code, message) { const error = new Error(message); error.code = code; return error; }
  function selectedSlides(value, total) {
    const text = value == null ? '' : String(value);
    if (!text.trim()) return new Uint32Array();
    const selected = new Set();
    for (const part of text.split(',')) {
      const match = /^\s*(\d+)(?:\s*-\s*(\d+))?\s*$/.exec(part);
      if (!match) throw failure('INVALID_ARGUMENT', '슬라이드 범위는 1,3-5 형식으로 입력해 주세요.');
      const first = Number(match[1]);
      const last = match[2] ? Number(match[2]) : first;
      if (!Number.isSafeInteger(first) || !Number.isSafeInteger(last) || first < 1 || last < first || last > total) {
        throw failure('INVALID_ARGUMENT', '슬라이드 번호는 1부터 ' + total + ' 사이에서 지정해 주세요.');
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
      if (!item || typeof item !== 'object') continue;
      const code = String(item.code || 'CONVERSION_NOTICE');
      const part = String(item.location?.part_name || '');
      const isDrawingControl = /^(?:[^:]+:)?(?:chOff|chExt|masterClrMapping)$/.test(String(item.location?.qualified_element_name || ''));
      // Notes and package inventories are not rendering failures. Never copy their raw text.
      if (/NOTE|COMMENT|AUTHOR/.test(code) || /(?:notes|comments|Authors)/i.test(part)) continue;
      if (code === 'OOXML_PART_UNSUPPORTED' && (part.endsWith('/') || !/\/(slides|charts|diagrams)\//.test(part))) continue;
      if (code === 'OOXML_ELEMENT_UNSUPPORTED' && !isDrawingControl && !/\/(slides|charts|diagrams)\//.test(part)) continue;
      const family = isDrawingControl ? 'DRAWINGML' : Object.keys(names).find((name) => code.includes(name)) || 'OTHER';
      if (!grouped.has(family)) grouped.set(family, {
        code: family === 'OTHER' ? 'UNSUPPORTED_FEATURES' : code,
        message: names[family] || '일부 PPTX 기능이 완전히 재현되지 않았습니다. 미리보기를 원본과 비교해 주세요.',
        slides: new Set()
      });
    }
    return [...grouped.values()].map(({code, message, slides}) => ({
      code,
      // The pinned engine reports part indices here, which can differ from
      // presentation order. Do not present them as user-visible slide numbers.
      message
    }));
  }


  // Count UTF-8 without allocating a second full-sized output buffer.
  function utf8Length(value, limit) {
    let bytes = 0;
    for (let i = 0; i < value.length; i++) {
      const code = value.charCodeAt(i);
      if (code < 0x80) bytes++;
      else if (code < 0x800) bytes += 2;
      else if (code >= 0xd800 && code <= 0xdbff && i + 1 < value.length &&
          value.charCodeAt(i + 1) >= 0xdc00 && value.charCodeAt(i + 1) <= 0xdfff) { bytes += 4; i++; }
      else bytes += 3;
      if (bytes > limit) return bytes;
    }
    return bytes;
  }

  function chartFallbacks(json) {
    let items;
    try { items = JSON.parse(json || '[]'); } catch { return []; }
    if (!Array.isArray(items)) return [];
    const records = [];
    for (const item of items) {
      if (records.length >= 100) break;
      if (!['DRAWINGML_CHART_FALLBACK', 'CHART_STRUCTURE_UNSUPPORTED'].includes(item?.code) ||
          item.support_tier !== 'fallback' || item.stage !== 'rendered') continue;
      const loc = item.location;
      if (!loc || !/^ppt\/slides\/[^/]{1,200}\.xml$/.test(loc.part_name) ||
          typeof loc.relationship_id !== 'string' || !loc.relationship_id.length || loc.relationship_id.length > 256 ||
          !Number.isInteger(loc.slide_index) || loc.slide_index < 0 || loc.slide_index >= 1000) continue;
      const rect = {x: loc.position?.x, y: loc.position?.y, width: loc.size?.width, height: loc.size?.height};
      if (!Object.values(rect).every(n => Number.isFinite(n) && Math.abs(n) <= 30000 * 9525) || rect.width <= 0 || rect.height <= 0) continue;
      records.push({part: loc.part_name, relationshipId: loc.relationship_id, slideNumber: loc.slide_index + 1, rect});
    }
    return records;
  }

  function render(engine, data, options) {
    const settings = options || {};
    let info, converted;
    try {
      info = engine.get_presentation_info(data);
      const width = info.widthPx, height = info.heightPx, total = info.slideCount;
      if (!Number.isInteger(total) || total < 1 || total > 1000) {
        throw failure('SLIDE_LIMIT', '슬라이드 수는 1장 이상 1,000장 이하여야 합니다.');
      }
      if (![width, height].every(value => Number.isFinite(value) && value > 0 && value <= 30000)) {
        throw failure('INVALID_DIMENSIONS', '지원하지 않는 슬라이드 크기입니다.');
      }
      const range = selectedSlides(settings.slideRange, total);
      // Empty ranges select only visible slides in the pinned engine.
      const effectiveRange = settings.includeHidden && !range.length
        ? Uint32Array.from({ length: total }, (_, index) => index + 1) : range;
      if (typeof settings.onProgress === 'function') settings.onProgress('텍스트·이미지·도형을 HTML로 변환하는 중입니다…');
      converted = engine.convert_with_options_metadata(data, true, !!settings.includeHidden, effectiveRange, 1);
      const slideCount = converted.slideCount;
      if (slideCount === 0) throw failure('EMPTY_SELECTION', '선택한 범위에 표시할 슬라이드가 없습니다. 숨긴 슬라이드 포함 설정을 확인해 주세요.');
      if (!Number.isInteger(slideCount) || slideCount < 1 || slideCount > total ||
          (effectiveRange.length && slideCount > effectiveRange.length)) {
        throw failure('SLIDE_COUNT_MISMATCH', '변환 결과의 슬라이드 수가 올바르지 않습니다.');
      }
      const html = converted.html;
      if (typeof html !== 'string' || !html.length) throw failure('EMPTY_OUTPUT', '변환된 HTML이 비어 있습니다.');
      if (html.length > MAX_OUTPUT_BYTES || utf8Length(html, MAX_OUTPUT_BYTES) > MAX_OUTPUT_BYTES) throw failure('OUTPUT_LIMIT', 'HTML 출력이 256 MiB를 초과합니다.');
      if (!html.trim()) throw failure('EMPTY_OUTPUT', '변환된 HTML이 비어 있습니다.');
      return { html, width, height, slideCount, diagnostics: summarizeDiagnostics(converted.diagnosticsJson),
        chartFallbacks: chartFallbacks(converted.diagnosticsJson) };
    } finally {
      // Always free metadata as well, even if releasing a result throws.
      try { if (converted) converted.free(); } finally { if (info) info.free(); }
    }
  }
  const api = Object.freeze({ render, selectedSlides, summarizeDiagnostics });
  if (typeof module === 'object' && module.exports) module.exports = api;
  else global.PptxCore = api;
})(typeof globalThis !== 'undefined' ? globalThis : self);
