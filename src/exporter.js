/* Offline, static PPTX export. Engine scripts, metadata and animation state are deliberately discarded. */
(function (global) {
  'use strict';

  // The build replaces each complete `/* NAME */ null` token with a JSON literal.
  const VIEWER_CSS = /* VIEWER_CSS */ null;
  const VIEWER_SCRIPT = /* VIEWER_SCRIPT */ null;
  const VIEWER_HASH = /* VIEWER_HASH */ null;

  const FORBIDDEN_TAGS = [
    'script', 'noscript', 'iframe', 'frame', 'frameset', 'object', 'embed',
    'link', 'meta', 'base', 'form', 'input', 'button', 'select', 'textarea',
    'option', 'template', 'style', 'foreignObject'
  ];
  const FORBIDDEN_ATTRS = [
    'srcdoc', 'formaction', 'action', 'ping', 'download', 'autofocus',
    'contenteditable', 'xml:base', 'srcset', 'imagesrcset'
  ];
  const URL_ATTRS = new Set([
    'src', 'href', 'xlink:href', 'poster', 'background', 'cite', 'longdesc',
    'data', 'usemap', 'formaction', 'action'
  ]);
  const REF_ATTRS = new Set([
    'aria-labelledby', 'aria-describedby', 'for', 'headers'
  ]);
  const NAVIGATION_ACTIONS = new Set(['slide', 'next', 'previous', 'first', 'last']);

  function escapeHtml(value) {
    return String(value).replace(/[&<>"']/g, function (ch) {
      return ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[ch];
    });
  }

  function safeLength(value, fallback) {
    const number = Number(value);
    return Number.isFinite(number) && number > 0 && number <= 30000 ? number : fallback;
  }

  function isDataUrl(value) {
    // Native images, fonts and audio may be embedded. SVG data images stay under image CSP.
    return /^data:(?:image\/(?:png|jpe?g|gif|webp|svg\+xml|bmp|avif)|font\/(?:woff2?|ttf|otf)|application\/(?:font-woff|font-woff2|x-font-ttf|x-font-opentype)|audio\/(?:mpeg|mp4|ogg|wav|webm)|video\/(?:mp4|webm|ogg));base64,[a-z0-9+/=\s]+$/i.test(value);
  }

  function cleanCss(css, status) {
    let value = String(css || '');
    // CSS comments can conceal @import and external URLs from simple filtering.
    value = value.replace(/\/\*[\s\S]*?\*\//g, '');
    value = value.replace(/@import\b[^;{}]*(?:;|$)/gi, function () {
      status.external = true;
      return '';
    });
    value = value.replace(/url\(\s*([^)]*?)\s*\)/gi, function (whole, target) {
      const unquoted = target.trim().replace(/^(['"])([\s\S]*)\1$/, '$2').trim();
      if (unquoted.startsWith('#') || isDataUrl(unquoted)) return whole;
      status.external = true;
      return 'url("")';
    });
    // An HTML style element must never be closable by deck content.
    return value.replace(/</g, '\\3c ');
  }

  function retainNavigationSurfaces(root) {
    // The pinned engine represents shape/image links as empty buttons. Rebuild
    // only those navigation surfaces; arbitrary buttons and forms stay forbidden.
    root.querySelectorAll('button.shape-action-surface').forEach(function (button) {
      if (button.getAttribute('type') !== 'button' || button.childNodes.length ||
          !button.parentElement?.classList.contains('shape') ||
          !NAVIGATION_ACTIONS.has(button.getAttribute('data-action'))) return;
      const surface = document.createElement('span');
      surface.className = 'shape-action-surface';
      for (const name of ['data-action', 'data-slide-target', 'title', 'aria-label']) {
        if (button.hasAttribute(name)) surface.setAttribute(name, button.getAttribute(name));
      }
      button.replaceWith(surface);
    });
  }

  function cleanNodes(root, slideNumber, slideCount, slideTargets, status) {
    const idMap = new Map();
    const used = new Set();
    root.querySelectorAll('[id]').forEach(function (node) {
      const old = node.getAttribute('id');
      let next = 'ptc-' + slideNumber + '-' + String(old).replace(/[^A-Za-z0-9_-]/g, '-');
      if (!next || used.has(next)) next += '-' + used.size;
      used.add(next);
      idMap.set(old, next);
      node.setAttribute('id', next);
    });
    const all = [root].concat(Array.from(root.querySelectorAll('*')));
    all.forEach(function (node) {
      const wasTimed = node.hasAttribute('data-timing-initial');
      const action = node.getAttribute('data-action');
      const rawTarget = (node.getAttribute('data-slide-target') || node.getAttribute('data-target-slide') || '').trim();
      const hadNavigation = action !== null || node.hasAttribute('data-slide-target') || node.hasAttribute('data-target-slide');
      const fixedTarget = action === 'slide' || (!action && rawTarget);
      // Relative actions follow the exported selection, including its hidden
      // slide setting. Fixed links only resolve to an included original slide.
      let target = null;
      if (action === 'next') target = Math.min(slideCount, slideNumber + 1);
      else if (action === 'previous') target = Math.max(1, slideNumber - 1);
      else if (action === 'first') target = 1;
      else if (action === 'last') target = slideCount;
      else if (fixedTarget) target = slideTargets.get(rawTarget) || slideTargets.get('slide-' + rawTarget);
      for (const attr of Array.from(node.attributes)) {
        const name = attr.name.toLowerCase();
        const value = attr.value.trim();
        if (name.startsWith('on') || name === 'name' || name === 'xml:base') {
          node.removeAttribute(attr.name);
          continue;
        }
        if (name.startsWith('data-') && ![
          'data-slide', 'data-chart-legend-position',
          'data-slide-target', 'data-target-slide', 'data-timing-initial'
        ].includes(name)) {
          node.removeAttribute(attr.name);
          continue;
        }
        if (name === 'style') {
          const mappedStyle = attr.value.replace(/url\(\s*(['"]?)#([^)'"\s]+)\1\s*\)/gi, function (match, quote, id) {
            return idMap.has(id) ? 'url(#' + idMap.get(id) + ')' : 'url(#' + id + ')';
          });
          node.setAttribute(attr.name, cleanCss(mappedStyle, status));
          continue;
        }
        if (URL_ATTRS.has(name)) {
          if (name === 'href' && node.localName.toLowerCase() === 'a' && /^(?:https?:\/\/|mailto:)/i.test(value)) {
            node.setAttribute('target', '_blank');
            node.setAttribute('rel', 'noopener noreferrer');
          } else if (value.startsWith('#')) {
            const mapped = idMap.get(value.slice(1));
            if (slideTargets.has(value.slice(1))) {
              node.setAttribute(attr.name, '#ptv-slide-' + slideTargets.get(value.slice(1)));
            }
            else if (mapped) node.setAttribute(attr.name, '#' + mapped);
            else node.removeAttribute(attr.name);
          } else if (node.localName.toLowerCase() === 'a' || !isDataUrl(value)) {
            node.removeAttribute(attr.name);
            status.external = true;
          }
          continue;
        }
        if (REF_ATTRS.has(name)) {
          const mapped = value.split(/\s+/).map(function (part) { return idMap.get(part) || ''; }).filter(Boolean);
          if (mapped.length) node.setAttribute(attr.name, mapped.join(' '));
          else node.removeAttribute(attr.name);
          continue;
        }
        if (name === 'data-timing-initial') {
          node.removeAttribute(attr.name);
          status.timing = true;
          continue;
        }
        if (name === 'data-slide-target' || name === 'data-target-slide') {
          // Add the single, validated destination below after removing source
          // action metadata. This also avoids remapping relative actions twice.
          node.removeAttribute(attr.name);
          continue;
        }
        if (/url\s*\(/i.test(value)) node.setAttribute(attr.name, cleanCss(value.replace(/url\(\s*(['"]?)#([^)'"\s]+)\1\s*\)/gi, function (match, quote, id) {
          return 'url(#' + (idMap.get(id) || id) + ')';
        }), status));
      }
      if (target) {
        node.setAttribute('data-slide-target', String(target));
        // Keep the engine's pointer-event selector for linked text over a shape
        // surface, but only with a constant action understood by our own viewer.
        node.setAttribute('data-action', 'slide');
        if (node.localName.toLowerCase() !== 'a') {
          node.setAttribute('role', 'button');
          node.setAttribute('tabindex', '0');
        }
      } else if (hadNavigation) {
        if (fixedTarget) status.missingSlideTarget = true;
        if (node.classList.contains('shape-action-surface')) node.remove();
        else {
          if (node.getAttribute('role') === 'button') node.removeAttribute('role');
          node.removeAttribute('tabindex');
          // A text run may override a different shape-level link. Once its
          // target is excluded, keep clicks from falling through to that surface.
          if (node.classList.contains('run')) node.style.pointerEvents = 'auto';
        }
      }
      if (wasTimed) {
        node.style.removeProperty('visibility');
        node.style.removeProperty('opacity');
      }
      // The engine hides initial animation targets in style as well as via attribute.
      if (node.localName.toLowerCase() === 'a' && node.hasAttribute('target')) {
        node.setAttribute('rel', 'noopener noreferrer');
      }
    });
  }

  function addChartData(slide, sourceNumber, input) {
    if (!sourceNumber || !Array.isArray(input) || input.length > 100) return;
    function near(a, b) { return Number.isFinite(a) && Number.isFinite(b) && Math.abs(a - b) <= 0.11; }
    function sameRect(a, b) { return a && b && ['x', 'y', 'width', 'height'].every(k => near(a[k] / 9525, b[k] / 9525)); }
    const records = input.filter(r => r && r.slideNumber === sourceNumber && r.rect &&
      Object.values(r.rect).every(n => Number.isFinite(n) && Math.abs(n) <= 30000 * 9525));
    const placeholders = Array.from(slide.querySelectorAll('.chart-placeholder')).filter(p => p.parentElement?.classList.contains('shape'));
    for (const record of records) {
      if (records.filter(other => sameRect(record.rect, other.rect)).length !== 1) continue;
      const matched = placeholders.filter(p => {
        const shape = p.parentElement;
        if (p.closest('.slide') !== slide || shape.querySelector('.chart-direct')) return false;
        // Only untransformed top-level chart frames can be matched from these
        // engine diagnostics. Group-relative coordinates are intentionally skipped.
        for (let n = shape; n && n !== slide; n = n.parentElement) {
          if ((n !== shape && n.classList.contains('shape')) ||
              (n.style.transform && n.style.transform !== 'none') ||
              (n !== shape && (n.style.left || n.style.top))) return false;
        }
        return [['left', 'x'], ['top', 'y'], ['width', 'width'], ['height', 'height']].every(([css, key]) =>
          /^-?\d+(?:\.\d+)?px$/.test(shape.style[css]) && near(parseFloat(shape.style[css]), record.rect[key] / 9525));
      });
      if (matched.length !== 1) continue;
      const placeholder = matched[0];
      if (placeholder.querySelector('.chart-cache-data')) continue;
      const box = document.createElement('div');
      box.className = 'chart-cache-data';
      box.style.cssText = 'padding:12px;color:#334155;font-size:13px;line-height:1.4;text-align:left';
      const note = document.createElement('p');
      note.textContent = '원본 차트 대신 문서에 저장된 데이터 · Cached data, not the original chart';
      box.appendChild(note);
      if (record.status === 'cached' && Array.isArray(record.columns) && record.columns.length <= 32 &&
          Array.isArray(record.rows) && record.rows.length <= 100 && record.rows.length &&
          record.rows.every(row => Array.isArray(row) && row.length === record.columns.length && row.every(v => typeof v === 'string' && v.length <= 256)) &&
          record.columns.every(v => typeof v === 'string' && v.length <= 256)) {
        const table = document.createElement('table');
        table.style.cssText = 'border-collapse:collapse;width:100%;font-size:12px';
        const caption = document.createElement('caption');
        caption.textContent = '저장된 차트 캐시의 원시 값 (백분율·날짜 등 숫자 서식 미적용, 빈칸은 저장된 값 없음, 긴 텍스트는 256자까지 표시)';
        table.appendChild(caption);
        const head = document.createElement('thead'), heading = document.createElement('tr');
        ['행'].concat(record.columns).forEach(label => {
          const th = document.createElement('th'); th.scope = 'col'; th.textContent = label;
          th.style.cssText = 'border:1px solid #94a3b8;padding:4px;text-align:left'; heading.appendChild(th);
        });
        head.appendChild(heading); table.appendChild(head);
        const body = document.createElement('tbody');
        record.rows.forEach((row, index) => {
          const tr = document.createElement('tr'), th = document.createElement('th');
          th.scope = 'row'; th.textContent = String(index + 1); tr.appendChild(th);
          row.forEach(value => { const td = document.createElement('td'); td.textContent = value;
            td.style.cssText = 'border:1px solid #cbd5e1;padding:4px'; tr.appendChild(td); });
          body.appendChild(tr);
        });
        table.appendChild(body); box.appendChild(table);
      } else {
        const reason = document.createElement('p');
        reason.textContent = record.status === 'limit' ? '저장된 데이터가 표시 한도(100행·16계열)를 초과합니다.' :
          record.status === 'empty' ? '문서에 저장된 차트 캐시 값이 없습니다.' : '저장된 차트 데이터를 안전하게 읽을 수 없습니다.';
        box.appendChild(reason);
      }
      placeholder.replaceChildren(box);
      placeholder.style.cssText = 'display:block;overflow:auto;box-sizing:border-box;width:100%;height:100%;background:#f8fafc;border:1px dashed #94a3b8';
      placeholder.setAttribute('tabindex', '0');
      placeholder.setAttribute('role', 'region');
      placeholder.setAttribute('aria-label', '저장된 차트 데이터');
    }
  }

  function prepare(options) {
    if (!global.DOMPurify || typeof global.DOMPurify.sanitize !== 'function') {
      throw new Error('HTML 정화 모듈을 사용할 수 없습니다.');
    }
    const source = options || {};
    const width = safeLength(source.width, 1280);
    const height = safeLength(source.height, 720);
    const title = String(source.title || '프레젠테이션').slice(0, 300);
    const diagnostics = Array.isArray(source.diagnostics) ? source.diagnostics.map(function (entry) {
      return {
        code: String(entry && entry.code || 'NOTICE').slice(0, 100),
        message: String(entry && entry.message || '').slice(0, 1000),
        slideIndex: Number.isInteger(entry && entry.slideIndex) ? entry.slideIndex : null
      };
    }) : [];
    const status = { external: false, timing: false, missingSlideTarget: false };

    // A template is inert; only selected style text and slide markup ever leave it.
    const sourceTemplate = document.createElement('template');
    sourceTemplate.innerHTML = String(source.html || '');
    const sourceRoot = sourceTemplate.content;
    const engineCss = Array.from(sourceRoot.querySelectorAll('style')).map(function (style) {
      return cleanCss(style.textContent, status);
    }).join('\n');
    const container = sourceRoot.querySelector('.pptx-container');
    if (!container) throw new Error('변환 결과에서 슬라이드를 찾을 수 없습니다.');
    // A pinned-engine chart fallback can close its container too early. Recover
    // genuine later shells in document order, never shells inside slide content.
    const recoverable = Array.from(sourceRoot.querySelectorAll('.slide-shell')).filter(function (shell) {
      return !shell.parentElement?.closest('.slide, .slide-shell') &&
        Array.from(shell.children).some(function (node) { return node.classList?.contains('slide'); });
    });
    recoverable.forEach(function (shell) { container.appendChild(shell); });
    const shells = Array.from(container.children).filter(function (node) {
      return node.classList && node.classList.contains('slide-shell');
    });
    if (!shells.length) throw new Error('변환 결과에 슬라이드가 없습니다.');
    const records = shells.map(function (sourceShell) {
      const slide = Array.from(sourceShell.children).find(function (node) {
        return node.classList && node.classList.contains('slide');
      });
      const shellNumber = sourceShell.getAttribute('data-slide');
      const slideNumber = slide && slide.getAttribute('data-slide');
      const number = Number(shellNumber || slideNumber);
      const sourceNumber = Number.isSafeInteger(number) && number >= 1 && number <= 1000 &&
        (!shellNumber || !slideNumber || Number(shellNumber) === Number(slideNumber)) ? number : null;
      return { shell: sourceShell, slide: slide, sourceNumber: sourceNumber };
    }).filter(function (record) { return Boolean(record.slide); });
    const slideTargets = new Map();
    records.forEach(function (record, index) {
      const number = index + 1;
      [record.shell.getAttribute('data-slide'), record.slide.getAttribute('data-slide')].forEach(function (oldNumber) {
        if (oldNumber) {
          slideTargets.set(oldNumber, number);
          slideTargets.set('slide-' + oldNumber, number);
        }
      });
      if (record.slide.id) slideTargets.set(record.slide.id, number);
    });
    const sourceSlides = [];
    const slideMarkup = records.map(function (record, index) {
      const slide = record.slide;
      addChartData(slide, record.sourceNumber, source.chartData);
      retainNavigationSurfaces(slide);
      if (/data-timing-initial\s*=/i.test(slide.outerHTML)) status.timing = true;
      const sanitized = global.DOMPurify.sanitize(slide.outerHTML, {
        USE_PROFILES: { html: true, svg: true, svgFilters: true },
        FORBID_TAGS: FORBIDDEN_TAGS,
        FORBID_ATTR: FORBIDDEN_ATTRS,
        ALLOW_DATA_ATTR: true,
        RETURN_DOM_FRAGMENT: true
      });
      const cleanSlide = sanitized.querySelector('.slide');
      if (!cleanSlide) return '';
      cleanNodes(cleanSlide, index + 1, records.length, slideTargets, status);
      cleanSlide.id = 'ptv-slide-' + (index + 1);
      cleanSlide.dataset.slide = String(index + 1);
      cleanSlide.style.width = width + 'px';
      cleanSlide.style.height = height + 'px';
      const cleanShell = document.createElement('div');
      cleanShell.className = 'slide-shell';
      cleanShell.dataset.slide = String(index + 1);
      cleanShell.appendChild(cleanSlide);
      sourceSlides.push(record.sourceNumber);
      return cleanShell.outerHTML;
    }).filter(Boolean);
    if (!slideMarkup.length) throw new Error('안전하게 표시할 수 있는 슬라이드가 없습니다.');
    if (source.slideCount !== undefined && (!Number.isInteger(source.slideCount) ||
        source.slideCount < 1 || source.slideCount > 1000 || source.slideCount !== slideMarkup.length)) {
      const failure = new Error('변환된 슬라이드 일부를 안전하게 내보내지 못했습니다.');
      failure.code = 'SLIDE_COUNT_MISMATCH';
      throw failure;
    }

    if (status.timing) diagnostics.push({ code: 'STATIC_ANIMATION', message: '애니메이션은 정적 상태로 표시됩니다.', slideIndex: null });
    if (status.external) diagnostics.push({ code: 'EXTERNAL_RESOURCE_BLOCKED', message: '외부 리소스를 제거했습니다. 포함된 이미지와 미디어만 표시됩니다.', slideIndex: null });
    if (status.missingSlideTarget) diagnostics.push({ code: 'SLIDE_LINK_TARGET_EXCLUDED', message: '내보내기에 포함되지 않은 슬라이드로 이동하는 링크는 비활성화했습니다.', slideIndex: null });

    const viewerCss = typeof VIEWER_CSS === 'string' ? VIEWER_CSS : '';
    const viewerScript = typeof VIEWER_SCRIPT === 'string' ? VIEWER_SCRIPT : '';
    const viewerHash = typeof VIEWER_HASH === 'string' ? VIEWER_HASH : '';
    if (!viewerCss || !viewerScript || !/^[A-Za-z0-9+/]{43}=$/.test(viewerHash)) {
      throw new Error('내보내기 뷰어가 올바르게 빌드되지 않았습니다.');
    }
    const csp = "default-src 'none'; img-src data:; font-src data:; media-src data:; style-src 'unsafe-inline'; script-src 'sha256-" + viewerHash + "'; connect-src 'none'; object-src 'none'; base-uri 'none'; form-action 'none'; frame-src 'none'";
    const out = '<!doctype html>\n<html lang="ko"><head>' +
      '<meta http-equiv="Content-Security-Policy" content="' + escapeHtml(csp) + '">' +
      '<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<title>' + escapeHtml(title) + '</title>' +
      '<style>' + engineCss + '\n' + viewerCss.replace(/</g, '\\3c ') +
      '\n@page { size: ' + width + 'px ' + height + 'px; margin: 0; }</style>' +
      '</head><body data-ptv-width="' + width + '" data-ptv-height="' + height + '">' +
      '<header id="ptv-toolbar" role="toolbar" aria-label="슬라이드 도구"><strong id="ptv-title">' + escapeHtml(title) + '</strong>' +
      '<div class="ptv-controls"><button id="ptv-prev" type="button" aria-label="이전 슬라이드">‹</button>' +
      '<label for="ptv-select">슬라이드</label><select id="ptv-select" aria-label="슬라이드 선택"></select>' +
      '<button id="ptv-next" type="button" aria-label="다음 슬라이드">›</button>' +
      '<button id="ptv-mode" type="button" aria-pressed="false">한 장씩</button>' +
      '<button id="ptv-print" type="button">인쇄</button></div></header>' +
      '<main id="ptv-main"><div class="pptx-container" id="ptv-slides">' + slideMarkup.join('') + '</div></main>' +
      '<script>' + viewerScript.replace(/<\//g, '<\\/') + '</script></body></html>';
    return { html: out, diagnostics: diagnostics, slideCount: slideMarkup.length,
      sourceSlides: sourceSlides.every(Boolean) && new Set(sourceSlides).size === sourceSlides.length ? sourceSlides : null };
  }

  global.PptxExport = Object.freeze({ prepare: prepare });
})(window);
