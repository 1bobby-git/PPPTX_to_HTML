(() => {
  'use strict';
  const api = window.pptxDesktop;
  const ids = ['workspace', 'select-file', 'file-prompt', 'file-detail', 'file-name', 'file-size',
    'title', 'slides', 'include-hidden', 'password', 'password-options', 'convert', 'cancel',
    'status', 'status-message', 'error-code', 'result', 'result-heading', 'result-summary',
    'ready-badge', 'preview-empty', 'preview-title', 'preview-description', 'preview-frame',
    'preview', 'save-html', 'save-status', 'warnings', 'warning-summary', 'warning-list', 'app-version'];
  const ui = Object.fromEntries(ids.map(id => [id, document.getElementById(id)]));
  const fields = ['title', 'slides', 'include-hidden', 'password'];
  const state = {ready: false, file: null, action: '', result: null, previewUrl: '', cancelRequested: false, operation: 0};
  let unsubscribe;

  function controls() {
    const busy = Boolean(state.action);
    ui.workspace.setAttribute('aria-busy', String(busy));
    ui.result.setAttribute('aria-busy', String(state.action === 'converting' || state.action === 'previewing'));
    ui['select-file'].disabled = !state.ready || busy;
    ui.convert.disabled = !state.ready || !state.file || busy;
    ui.cancel.hidden = state.action !== 'converting';
    ui.cancel.disabled = state.cancelRequested;
    ui.preview.disabled = !state.result || busy;
    ui['save-html'].disabled = !state.result || busy;
    for (const id of fields) ui[id].disabled = !state.ready || busy;
  }

  function status(message, kind = 'ready', code = '') {
    ui['status-message'].textContent = message;
    ui.status.dataset.kind = kind;
    ui['error-code'].textContent = code ? '오류 코드: ' + code : '';
    ui['error-code'].hidden = !code;
  }

  function showError(result, fallback) {
    const message = typeof result?.error === 'string' && result.error ? result.error : fallback;
    const code = typeof result?.code === 'string' ? result.code.slice(0, 100) : '';
    status(message, 'error', code);
    if (code === 'PASSWORD_REQUIRED' || code === 'INCORRECT_PASSWORD') ui['password-options'].open = true;
    ui.status.focus({preventScroll: true});
  }

  function clearPreview() {
    ui['preview-frame'].removeAttribute('src');
    ui['preview-frame'].hidden = true;
    ui['preview-empty'].hidden = false;
    state.previewUrl = '';
  }

  function clearResult() {
    state.result = null;
    clearPreview();
    ui['ready-badge'].hidden = true;
    ui['result-summary'].textContent = '변환이 끝나면 미리 보고 저장할 수 있습니다.';
    ui['preview-title'].textContent = 'HTML 미리보기';
    ui['preview-description'].textContent = '파일을 선택하고 변환을 시작해 주세요.';
    ui['warning-list'].replaceChildren();
    ui.warnings.hidden = true;
    ui.warnings.open = false;
    ui['save-status'].textContent = '';
    ui['save-status'].hidden = true;
  }

  function formatSize(size) {
    if (!Number.isFinite(size) || size < 0) return '';
    if (size < 1024) return size.toLocaleString('ko-KR') + ' B';
    if (size < 1024 * 1024) return (size / 1024).toLocaleString('ko-KR', {maximumFractionDigits: 1}) + ' KiB';
    return (size / (1024 * 1024)).toLocaleString('ko-KR', {maximumFractionDigits: 2}) + ' MiB';
  }

  function renderResult(result) {
    state.result = result;
    const count = Number.isInteger(result.slideCount) && result.slideCount >= 0 ? result.slideCount : 0;
    const dimensions = Number.isFinite(result.width) && Number.isFinite(result.height)
      ? ' · ' + Math.round(result.width) + ' × ' + Math.round(result.height) : '';
    ui['result-summary'].textContent = count.toLocaleString('ko-KR') + '개 슬라이드' + dimensions;
    ui['ready-badge'].hidden = false;
    ui['preview-title'].textContent = 'HTML 변환 완료';
    ui['preview-description'].textContent = '미리보기를 열거나 HTML 파일로 저장해 주세요.';
    const warnings = Array.isArray(result.warnings) ? result.warnings.filter(item => typeof item?.message === 'string' && item.message) : [];
    ui['warning-list'].replaceChildren();
    for (const warning of warnings.slice(0, 100)) {
      const item = document.createElement('li');
      item.textContent = warning.message;
      ui['warning-list'].append(item);
    }
    if (warnings.length > 100) {
      const item = document.createElement('li');
      item.textContent = '추가 알림 ' + (warnings.length - 100).toLocaleString('ko-KR') + '개가 있습니다.';
      ui['warning-list'].append(item);
    }
    ui['warning-summary'].textContent = '변환 알림 ' + warnings.length.toLocaleString('ko-KR') + '개';
    ui.warnings.hidden = warnings.length === 0;
    status(count.toLocaleString('ko-KR') + '개 슬라이드를 변환했습니다. HTML을 저장해 주세요.', 'success');
    ui['result-heading'].focus({preventScroll: true});
  }

  ui['select-file'].addEventListener('click', async () => {
    if (!state.ready || state.action) return;
    state.action = 'choosing';
    controls();
    try {
      const result = await api.chooseFile();
      if (result?.cancelled) return;
      if (!result?.ok || typeof result.file?.name !== 'string') {
        showError(result, '파일을 선택하지 못했습니다. 다시 시도해 주세요.');
        return;
      }
      state.file = result.file;
      state.operation += 1;
      clearResult();
      ui.password.value = '';
      ui.title.value = '';
      ui['file-detail'].hidden = false;
      ui['file-name'].textContent = state.file.name;
      ui['file-size'].textContent = formatSize(state.file.size);
      ui['file-prompt'].textContent = '변환할 파일을 선택했습니다.';
      ui['select-file'].textContent = '다른 파일 선택';
      ui.title.placeholder = state.file.name.replace(/\.pptx$/i, '') || '비워두면 파일 이름을 사용합니다';
      status('파일 선택을 완료했습니다. HTML로 변환을 눌러 주세요.');
    } catch {
      showError(null, '파일을 선택하지 못했습니다. 다시 시도해 주세요.');
    } finally {
      state.action = '';
      controls();
    }
  });

  ui.convert.addEventListener('click', async () => {
    if (!state.ready || !state.file || state.action) return;
    const slides = ui.slides.value.trim();
    if (slides && !/^\d+(?:\s*-\s*\d+)?(?:\s*,\s*\d+(?:\s*-\s*\d+)?)*$/.test(slides)) {
      ui.slides.setCustomValidity('슬라이드 범위를 1,3-5 형식으로 입력해 주세요.');
      ui.slides.reportValidity();
      return;
    }
    ui.slides.setCustomValidity('');
    const operation = ++state.operation;
    let passwordError = false;
    clearResult();
    state.action = 'converting';
    state.cancelRequested = false;
    controls();
    status('문서를 HTML로 변환하고 있습니다…', 'working');
    try {
      const result = await api.convert({slides, includeHidden: ui['include-hidden'].checked,
        title: ui.title.value.trim(), password: ui.password.value});
      if (operation !== state.operation) return;
      if (result?.ok) renderResult(result);
      else if (result?.cancelled || result?.code === 'CANCELLED') status('변환을 취소했습니다. 설정을 확인하고 다시 시작할 수 있습니다.');
      else {
        passwordError = result?.code === 'PASSWORD_REQUIRED' || result?.code === 'INCORRECT_PASSWORD';
        showError(result, '변환을 완료하지 못했습니다. 파일과 설정을 확인하고 다시 시도해 주세요.');
      }
    } catch {
      if (operation === state.operation) showError(null, '변환을 완료하지 못했습니다. 다시 시도해 주세요.');
    } finally {
      ui.password.value = '';
      if (operation === state.operation) {
        state.action = '';
        state.cancelRequested = false;
        controls();
        if (passwordError) ui.password.focus();
      }
    }
  });

  ui.cancel.addEventListener('click', async () => {
    if (state.action !== 'converting' || state.cancelRequested) return;
    const operation = state.operation;
    state.cancelRequested = true;
    status('변환을 취소하고 있습니다…', 'working');
    controls();
    try {
      const result = await api.cancel();
      if (operation === state.operation && state.action === 'converting' && !result?.ok) {
        state.cancelRequested = false;
        showError(result, '취소 요청을 처리하지 못했습니다. 다시 시도해 주세요.');
        controls();
      }
    } catch {
      if (operation === state.operation && state.action === 'converting') {
        state.cancelRequested = false;
        showError(null, '취소 요청을 처리하지 못했습니다. 다시 시도해 주세요.');
        controls();
      }
    }
  });

  ui.preview.addEventListener('click', async () => {
    if (!state.result || state.action) return;
    if (state.previewUrl) {
      ui['preview-frame'].focus();
      return;
    }
    state.action = 'previewing';
    controls();
    try {
      const result = await api.preview();
      if (!result?.ok || typeof result.url !== 'string') {
        showError(result, '미리보기를 열지 못했습니다. 다시 시도해 주세요.');
        return;
      }
      const url = new URL(result.url);
      if (url.protocol !== 'pptx-app:' || url.hostname !== 'preview' || url.username || url.password || url.port) throw new Error('Invalid preview URL');
      state.previewUrl = result.url;
      ui['preview-frame'].src = state.previewUrl;
      ui['preview-frame'].hidden = false;
      ui['preview-empty'].hidden = true;
      status('미리보기를 열었습니다. HTML 저장을 눌러 파일로 저장할 수 있습니다.', 'success');
      ui['preview-frame'].focus();
    } catch {
      showError(null, '미리보기를 열지 못했습니다. 다시 시도해 주세요.');
    } finally {
      state.action = '';
      controls();
    }
  });

  ui['save-html'].addEventListener('click', async () => {
    if (!state.result || state.action) return;
    state.action = 'saving';
    controls();
    try {
      const result = await api.saveHtml();
      if (result?.cancelled) return;
      if (!result?.ok) {
        showError(result, 'HTML 파일을 저장하지 못했습니다. 다시 시도해 주세요.');
        return;
      }
      const name = typeof result.name === 'string' && result.name ? result.name : 'HTML 파일';
      ui['save-status'].textContent = name + ' 저장 완료';
      ui['save-status'].hidden = false;
      status('HTML을 저장했습니다. 저장한 파일을 브라우저로 열어 보세요.', 'success');
    } catch {
      showError(null, 'HTML 파일을 저장하지 못했습니다. 다시 시도해 주세요.');
    } finally {
      state.action = '';
      controls();
    }
  });

  for (const id of fields) {
    ui[id].addEventListener(id === 'include-hidden' ? 'change' : 'input', () => {
      if (id === 'slides') ui.slides.setCustomValidity('');
      if (!state.action && state.result) {
        clearResult();
        controls();
        status('설정이 바뀌었습니다. HTML로 변환을 다시 눌러 주세요.');
      }
    });
  }

  window.addEventListener('beforeunload', () => {
    state.operation += 1;
    ui.password.value = '';
    if (typeof unsubscribe === 'function') unsubscribe();
  });

  async function initialize() {
    const methods = ['getInfo', 'chooseFile', 'convert', 'cancel', 'preview', 'saveHtml', 'onProgress'];
    if (!api || methods.some(method => typeof api[method] !== 'function')) {
      showError(null, '변환기를 시작하지 못했습니다. 앱을 닫고 다시 실행해 주세요.');
      return;
    }
    try {
      const info = await api.getInfo();
      if (!info || typeof info.version !== 'string') throw new Error('Invalid app information');
      ui['app-version'].textContent = 'PPTX → HTML v' + info.version;
      unsubscribe = api.onProgress(event => {
        if (state.action === 'converting' && !state.cancelRequested && typeof event?.message === 'string') status(event.message, 'working');
      });
      state.ready = true;
      controls();
      status('PPTX 파일을 선택해 주세요.');
    } catch {
      showError(null, '변환기를 시작하지 못했습니다. 앱을 닫고 다시 실행해 주세요.');
    }
  }
  initialize();
})();
