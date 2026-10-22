(function () {
  'use strict';

  const MAX_BYTES = 64 * 1024 * 1024;
  const TIMEOUT_MS = 120000;
  const $ = (id) => document.getElementById(id);
  const ui = {
    input: $('file-input'), select: $('select-button'), change: $('change-button'),
    drop: $('drop-zone'), detail: $('file-detail'), fileName: $('file-name'),
    fileSize: $('file-size'), title: $('output-title'), range: $('slide-range'),
    hidden: $('include-hidden'), convert: $('convert-button'), cancel: $('cancel-button'),
    status: $('status-line'), statusText: $('status-text'), preview: $('preview-frame'),
    empty: $('preview-empty'), download: $('download-button'), open: $('open-button'),
    ready: $('result-ready'), subtitle: $('result-subtitle'), warnings: $('warnings'),
    warningCount: $('warning-count'), warningList: $('warning-list'),
    password: $('document-password'), protection: $('protection-options'),
    diagnostics: $('diagnostics'), diagnosticsText: $('diagnostics-text'),
    helper: $('powerpoint-helper-button'), launcher: $('powerpoint-launcher-button')
  };

  let selectedFile = null;
  let resultUrl = null;
  let worker = null;
  let deadline = null;
  let busy = false;
  let jobId = 0;
  let dragDepth = 0;

  function setStatus(message, tone) {
    ui.statusText.textContent = message;
    ui.status.dataset.tone = tone || 'neutral';
    ui.status.dataset.errorCode = '';
  }

  function humanSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KiB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MiB';
  }

  function baseName(name) {
    return name.replace(/\.pptx$/i, '') || 'presentation';
  }

  function safeDownloadName(title) {
    const safe = title.replace(/[\\/:*?"<>|\x00-\x1f]/g, '_').replace(/[.\s]+$/g, '').trim().slice(0, 120);
    return (safe || 'presentation') + '.html';
  }

  function clearResult() {
    ui.preview.removeAttribute('src');
    ui.preview.hidden = true;
    ui.empty.hidden = false;
    ui.download.disabled = true;
    ui.open.disabled = true;
    ui.ready.hidden = true;
    ui.subtitle.textContent = '변환 후 이곳에서 확인할 수 있습니다';
    ui.warnings.hidden = true;
    ui.warnings.open = false;
    ui.warningList.replaceChildren();
    ui.diagnostics.hidden = true;
    ui.diagnostics.open = false;
    ui.diagnosticsText.textContent = '';
    if (resultUrl) {
      URL.revokeObjectURL(resultUrl);
      resultUrl = null;
    }
  }

  function refreshControls() {
    ui.input.disabled = busy;
    ui.select.disabled = busy;
    ui.change.disabled = busy;
    ui.title.disabled = busy;
    ui.range.disabled = busy;
    ui.hidden.disabled = busy;
    ui.password.disabled = busy;
    ui.convert.disabled = busy || !selectedFile;
    ui.cancel.hidden = !busy;
    ui.drop.classList.toggle('is-busy', busy);
  }

  function stopWorker() {
    if (deadline !== null) {
      clearTimeout(deadline);
      deadline = null;
    }
    if (worker) {
      worker.terminate();
      worker = null;
    }
  }

  function finishJob(message, tone) {
    stopWorker();
    busy = false;
    refreshControls();
    setStatus(message, tone);
  }

  function validateFile(file) {
    if (!file || !/\.pptx$/i.test(file.name)) return 'PPTX(.pptx) 파일을 선택해 주세요.';
    if (file.size === 0) return '빈 파일은 변환할 수 없습니다.';
    if (file.size > MAX_BYTES) return '파일 크기는 최대 64 MB(64 MiB)까지 지원합니다.';
    return null;
  }

  function selectFile(file) {
    if (busy) return;
    clearResult();
    ui.password.value = '';
    const error = validateFile(file);
    if (error) {
      selectedFile = null;
      ui.detail.hidden = true;
      ui.drop.hidden = false;
      ui.input.value = '';
      refreshControls();
      setStatus(error, 'error');
      return;
    }
    selectedFile = file;
    ui.fileName.textContent = file.name;
    ui.fileSize.textContent = humanSize(file.size);
    ui.detail.hidden = false;
    ui.drop.hidden = true;
    ui.title.value = baseName(file.name);
    refreshControls();
    setStatus('파일을 선택했습니다. 설정을 확인한 뒤 변환해 주세요.', 'neutral');
  }

  function validateRange(value) {
    if (!value.trim()) return null;
    const items = value.split(',');
    for (const item of items) {
      const match = /^\s*(\d+)(?:\s*-\s*(\d+))?\s*$/.exec(item);
      if (!match) return '슬라이드 범위는 1,3-5 형식으로 입력해 주세요.';
      const start = Number(match[1]);
      const end = match[2] ? Number(match[2]) : start;
      if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 1 || end < start) {
        return '슬라이드 번호는 1 이상이어야 하고, 범위의 끝은 시작보다 작을 수 없습니다.';
      }
    }
    return null;
  }

  function showWarnings(diagnostics) {
    ui.warningList.replaceChildren();
    if (!Array.isArray(diagnostics) || diagnostics.length === 0) {
      ui.warnings.hidden = true;
      return;
    }
    for (const item of diagnostics) {
      const li = document.createElement('li');
      li.textContent = typeof item.message === 'string' ? item.message : String(item.code || '변환 알림');
      ui.warningList.append(li);
    }
    ui.warningCount.textContent = '변환 알림 ' + diagnostics.length + '건';
    ui.warnings.hidden = false;
  }

  function showResult(message) {
    if (typeof message.html !== 'string' || !message.html.trim()) {
      finishJob('변환 결과가 비어 있습니다. 다른 PPTX 파일로 다시 시도해 주세요.', 'error');
      return;
    }
    try {
      if (!window.PptxExport || typeof window.PptxExport.prepare !== 'function') {
        throw new Error('결과 처리 모듈을 사용할 수 없습니다.');
      }
      const prepared = window.PptxExport.prepare(Object.assign({}, message, {
        title: ui.title.value.trim() || baseName(selectedFile.name)
      }));
      if (!prepared || typeof prepared.html !== 'string' || !prepared.html.trim()) {
        throw new Error('결과 HTML이 비어 있습니다.');
      }
      clearResult();
      resultUrl = URL.createObjectURL(new Blob([prepared.html], { type: 'text/html;charset=utf-8' }));
      ui.preview.src = resultUrl;
      ui.preview.hidden = false;
      ui.empty.hidden = true;
      ui.download.disabled = false;
      ui.open.disabled = false;
      ui.ready.hidden = false;
      const count = Number(prepared.slideCount);
      ui.subtitle.textContent = Number.isFinite(count) && count >= 0 ? count + '개 슬라이드 변환 완료' : '변환 완료';
      showWarnings(prepared.diagnostics);
      finishJob('HTML 변환이 완료되었습니다. 미리보기를 확인하거나 파일을 저장하세요.', 'success');
    } catch (error) {
      clearResult();
      finishJob('결과를 표시하지 못했습니다: ' + (error && error.message ? error.message : '알 수 없는 오류'), 'error');
    }
  }

  async function convert() {
    if (busy || !selectedFile) return;
    const fileError = validateFile(selectedFile);
    if (fileError) {
      setStatus(fileError, 'error');
      return;
    }
    const rangeError = validateRange(ui.range.value);
    if (rangeError) {
      ui.range.focus();
      setStatus(rangeError, 'error');
      return;
    }
    const workerSource = $('engine-worker');
    if (!workerSource || !workerSource.textContent.trim() || typeof Worker === 'undefined') {
      setStatus('이 브라우저에서 로컬 변환 엔진을 사용할 수 없습니다.', 'error');
      return;
    }
    clearResult();
    busy = true;
    const id = ++jobId;
    const file = selectedFile;
    const options = {
      includeHidden: ui.hidden.checked,
      slideRange: ui.range.value.trim(),
      title: ui.title.value.trim() || baseName(file.name),
      password: ui.password.value
    };
    // Keep the password only in this conversion request, never in storage or output.
    ui.password.value = '';
    refreshControls();
    setStatus('파일을 읽는 중입니다…', 'working');
    deadline = setTimeout(function () {
      if (id !== jobId || !busy) return;
      ++jobId;
      finishJob('변환 시간이 2분을 초과해 중지했습니다. 파일 크기를 확인하고 다시 시도해 주세요.', 'error');
    }, TIMEOUT_MS);

    try {
      const bytes = await file.arrayBuffer();
      if (id !== jobId || !busy) return;
      if (bytes.byteLength === 0 || bytes.byteLength > MAX_BYTES) {
        finishJob('파일 크기가 올바르지 않습니다. 최대 64 MB(64 MiB)까지 지원합니다.', 'error');
        return;
      }
      const objectUrl = URL.createObjectURL(new Blob([workerSource.textContent], { type: 'text/javascript' }));
      try {
        worker = new Worker(objectUrl);
      } finally {
        URL.revokeObjectURL(objectUrl);
      }
      worker.onmessage = function (event) {
        const message = event.data;
        if (id !== jobId || !busy || !message || message.id !== id) return;
        if (message.type === 'progress') {
          setStatus(typeof message.message === 'string' && message.message ? message.message : '변환 중입니다…', 'working');
        } else if (message.type === 'result') {
          showResult(message);
        } else if (message.type === 'error') {
          finishJob('변환 실패: ' + (typeof message.message === 'string' ? message.message : '알 수 없는 오류'), 'error');
          ui.status.dataset.errorCode = typeof message.code === 'string' ? message.code : 'CONVERSION_FAILED';
          if (typeof message.details === 'string' && message.details) {
            ui.diagnosticsText.textContent = message.details;
            ui.diagnostics.hidden = false;
          }
          if (['PASSWORD_REQUIRED', 'INCORRECT_PASSWORD', 'DRM_PROTECTED', 'UNKNOWN_CONTAINER', 'INVALID_CONTAINER', 'LEGACY_PPT', 'UNSUPPORTED_ENCRYPTION', 'UNSUPPORTED_PACKAGE'].includes(message.code)) {
            ui.protection.open = true;
          }
          if (message.code === 'PASSWORD_REQUIRED' || message.code === 'INCORRECT_PASSWORD') ui.password.focus();
        }
      };
      worker.onerror = function (event) {
        if (id !== jobId || !busy) return;
        if (event && typeof event.preventDefault === 'function') event.preventDefault();
        finishJob('변환 엔진에서 오류가 발생했습니다. 파일을 확인하고 다시 시도해 주세요.', 'error');
      };
      setStatus('슬라이드를 변환하는 중입니다…', 'working');
      worker.postMessage({ type: 'convert', id: id, bytes: bytes, options: options }, [bytes]);
      options.password = '';
    } catch (error) {
      if (id !== jobId || !busy) return;
      finishJob('파일을 읽거나 변환할 수 없습니다: ' + (error && error.message ? error.message : '알 수 없는 오류'), 'error');
    }
  }

  ui.select.addEventListener('click', function () { if (!busy) ui.input.click(); });
  ui.change.addEventListener('click', function () { if (!busy) ui.input.click(); });
  ui.input.addEventListener('change', function () { if (ui.input.files && ui.input.files[0]) selectFile(ui.input.files[0]); });
  ui.input.addEventListener('click', function () { if (!busy) ui.input.value = ''; });
  ui.convert.addEventListener('click', convert);
  ui.password.addEventListener('keydown', function (event) {
    if (event.key === 'Enter' && !busy) {
      event.preventDefault();
      convert();
    }
  });
  function downloadHelper(sourceId, filename) {
    const source = $(sourceId);
    if (!source || !source.textContent.trim()) {
      setStatus('변환 도우미가 포함되지 않은 파일입니다. 최신 index.html을 다시 받아 주세요.', 'error');
      return;
    }
    const binary = atob(source.textContent.trim());
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    const url = URL.createObjectURL(new Blob([bytes], {type: 'application/octet-stream'}));
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 30000);
  }
  ui.helper.addEventListener('click', function () { downloadHelper('powerpoint-helper-source', 'Convert-PptxWithPowerPoint.ps1'); });
  ui.launcher.addEventListener('click', function () { downloadHelper('powerpoint-launcher-source', 'Convert-PptxWithPowerPoint.cmd'); });
  ui.cancel.addEventListener('click', function () {
    if (!busy) return;
    ++jobId;
    finishJob('변환을 취소했습니다.', 'neutral');
  });
  [ui.title, ui.range, ui.hidden].forEach(function (control) {
    control.addEventListener('input', function () {
      if (busy) return;
      if (resultUrl) {
        clearResult();
        setStatus('설정이 변경되었습니다. 다시 변환해 주세요.', 'neutral');
      }
    });
  });
  ui.download.addEventListener('click', function () {
    if (!resultUrl || !selectedFile) return;
    const link = document.createElement('a');
    link.href = resultUrl;
    link.download = safeDownloadName(ui.title.value.trim() || baseName(selectedFile.name));
    document.body.append(link);
    link.click();
    link.remove();
  });
  ui.open.addEventListener('click', function () {
    if (!resultUrl) return;
    const link = document.createElement('a');
    link.href = resultUrl;
    link.target = '_blank';
    link.rel = 'noopener noreferrer';
    document.body.append(link);
    link.click();
    link.remove();
  });

  [ui.drop, ui.detail].forEach(function (target) {
  ['dragenter', 'dragover'].forEach(function (type) {
    target.addEventListener(type, function (event) {
      event.preventDefault();
      if (type === 'dragenter') dragDepth++;
      if (!busy) target.classList.add('drag-over');
    });
  });
  target.addEventListener('dragleave', function (event) {
    event.preventDefault();
    dragDepth = Math.max(0, dragDepth - 1);
    if (dragDepth === 0) target.classList.remove('drag-over');
  });
  target.addEventListener('drop', function (event) {
    event.preventDefault();
    dragDepth = 0;
    target.classList.remove('drag-over');
    if (!busy && event.dataTransfer && event.dataTransfer.files.length) selectFile(event.dataTransfer.files[0]);
  });
  });
  document.addEventListener('dragover', function (event) { event.preventDefault(); });
  document.addEventListener('drop', function (event) { event.preventDefault(); });
  window.addEventListener('beforeunload', function () {
    stopWorker();
    if (resultUrl) URL.revokeObjectURL(resultUrl);
  });

  clearResult();
  refreshControls();
})();
