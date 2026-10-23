(function () {
  'use strict';

  const MAX_RESULT_BYTES = 256 * 1024 * 1024;
  const tokenElement = document.querySelector('meta[name="pptx-csrf"]');
  let token = tokenElement ? tokenElement.content : '';

  function failure(code, message) {
    const error = new Error(message);
    error.code = code;
    return error;
  }

  async function readError(response) {
    let result;
    try { result = await response.json(); } catch (_) { /* Use a safe fallback below. */ }
    if (result && typeof result.error === 'string' && result.error.length <= 1000) {
      return failure(typeof result.code === 'string' ? result.code : 'CONVERSION_FAILED', result.error);
    }
    if (response.status === 429) return failure('BUSY', '다른 파일을 변환하고 있습니다. 잠시 후 다시 시도해 주세요.');
    if (response.status === 413) return failure('FILE_TOO_LARGE', '파일 크기는 최대 64 MB(64 MiB)까지 지원합니다.');
    return failure('SERVER_ERROR', '변환 서비스가 요청을 처리하지 못했습니다. 잠시 후 다시 시도해 주세요.');
  }

  function readMetadata(value) {
    if (typeof value !== 'string' || value.length > 16384 || !/^[A-Za-z0-9_-]+$/.test(value)) {
      throw failure('INVALID_RESPONSE', '변환 결과 정보를 읽을 수 없습니다. 다시 변환해 주세요.');
    }
    try {
      const binary = atob(value.replace(/-/g, '+').replace(/_/g, '/'));
      const bytes = Uint8Array.from(binary, function (char) { return char.charCodeAt(0); });
      const data = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(bytes));
      if (!data || !Number.isInteger(data.slideCount) || data.slideCount < 1 || data.slideCount > 1000) throw new Error('Invalid slide count');
      return {
        slideCount: data.slideCount,
        diagnostics: Array.isArray(data.warnings) ? data.warnings.filter(function (warning) {
          return warning && typeof warning.message === 'string';
        }).slice(0, 100) : []
      };
    } catch (_) {
      throw failure('INVALID_RESPONSE', '변환 결과 정보를 읽을 수 없습니다. 다시 변환해 주세요.');
    }
  }

  async function readResult(response, controller) {
    const length = Number(response.headers.get('Content-Length'));
    if (Number.isFinite(length) && length > MAX_RESULT_BYTES) {
      controller.abort();
      throw failure('OUTPUT_TOO_LARGE', '변환 결과가 너무 큽니다. 슬라이드 범위를 나누어 변환해 주세요.');
    }
    if (!response.body || typeof response.body.getReader !== 'function') {
      const blob = await response.blob();
      if (!blob.size || blob.size > MAX_RESULT_BYTES) throw failure('INVALID_RESPONSE', '변환 결과가 비어 있거나 너무 큽니다.');
      return blob.slice(0, blob.size, 'text/html;charset=utf-8');
    }
    const reader = response.body.getReader();
    const parts = [];
    let total = 0;
    try {
      for (;;) {
        const item = await reader.read();
        if (item.done) break;
        total += item.value.byteLength;
        if (total > MAX_RESULT_BYTES) {
          controller.abort();
          throw failure('OUTPUT_TOO_LARGE', '변환 결과가 너무 큽니다. 슬라이드 범위를 나누어 변환해 주세요.');
        }
        parts.push(item.value);
      }
    } finally {
      reader.releaseLock();
    }
    if (!total) throw failure('INVALID_RESPONSE', '변환 결과가 비어 있습니다. 다시 변환해 주세요.');
    return new Blob(parts, { type: 'text/html;charset=utf-8' });
  }

  async function check() {
    if (!token || typeof fetch !== 'function' || typeof AbortController !== 'function') {
      throw failure('SERVER_UNAVAILABLE', '이 페이지에서 변환 서비스에 연결할 수 없습니다. 페이지를 새로 고쳐 주세요.');
    }
    const controller = new AbortController();
    const timer = setTimeout(function () { controller.abort(); }, 8000);
    try {
      const response = await fetch('/api/status', {
        cache: 'no-store', credentials: 'same-origin', mode: 'same-origin', redirect: 'error', signal: controller.signal
      });
      if (!response.ok) throw await readError(response);
      const status = await response.json();
      if (!status || status.ok !== true || status.engine !== 'native' || !['local', 'server'].includes(status.mode) ||
          typeof status.csrfToken !== 'string' || !/^[a-f0-9]{64}$/.test(status.csrfToken)) {
        throw failure('INVALID_RESPONSE', '변환 서비스의 상태를 확인할 수 없습니다.');
      }
      // A restarted service issues a new token even while this page stays open.
      token = status.csrfToken;
      return status;
    } catch (error) {
      if (error && error.code) throw error;
      throw failure('SERVER_UNAVAILABLE', '변환 서비스에 연결할 수 없습니다. 연결 상태를 확인한 뒤 다시 시도해 주세요.');
    } finally {
      clearTimeout(timer);
    }
  }

  function convert(file, options, onProgress) {
    const controller = new AbortController();
    const form = new FormData();
    form.append('file', file, file.name);
    form.append('options', JSON.stringify(options));
    // The request body owns the transient password from this point onward.
    options.password = '';
    const promise = (async function () {
      try {
        onProgress('파일을 변환 서비스에 전달하고 처리하는 중입니다…');
        const response = await fetch('/api/convert', {
          method: 'POST', body: form, headers: { 'X-Pptx-Token': token },
          cache: 'no-store', credentials: 'same-origin', mode: 'same-origin', redirect: 'error', signal: controller.signal
        });
        if (!response.ok) throw await readError(response);
        if (!/^text\/html(?:\s*;|\s*$)/i.test(response.headers.get('Content-Type') || '')) {
          throw failure('INVALID_RESPONSE', '변환 서비스에서 HTML 결과를 받지 못했습니다.');
        }
        const metadata = readMetadata(response.headers.get('X-Pptx-Metadata'));
        onProgress('변환된 HTML을 받는 중입니다…');
        const blob = await readResult(response, controller);
        return Object.assign({ blob: blob }, metadata);
      } catch (error) {
        if (error && (error.code || error.name === 'AbortError')) throw error;
        throw failure('SERVER_UNAVAILABLE', '변환 서비스와 연결이 끊겼습니다. 연결을 확인한 뒤 다시 시도해 주세요.');
      } finally {
        // Drop the form's references after completion or cancellation.
        form.delete('options');
        form.delete('file');
      }
    })();
    return { promise: promise, abort: function () { controller.abort(); } };
  }

  window.PptxServer = Object.freeze({ check: check, convert: convert, timeoutMs: 195000 });
})();
