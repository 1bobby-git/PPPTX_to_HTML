/* Export viewer: static slides only. PowerPoint animations and transitions are not replayed. */
(function () {
  'use strict';
  const body = document.body;
  const slidesHost = document.getElementById('ptv-slides');
  const shells = Array.from(slidesHost.querySelectorAll(':scope > .slide-shell'));
  const select = document.getElementById('ptv-select');
  const previous = document.getElementById('ptv-prev');
  const next = document.getElementById('ptv-next');
  const mode = document.getElementById('ptv-mode');
  const print = document.getElementById('ptv-print');
  const width = Number(body.dataset.ptvWidth) || 1280;
  const height = Number(body.dataset.ptvHeight) || 720;
  let current = 0;
  let oneAtATime = false;
  let programmaticScroll = false;
  let scrollTimer = null;
  let fitFrame = null;
  const visibleRatios = new Map(shells.map(function (shell) { return [shell, 0]; }));

  shells.forEach(function (shell, index) {
    const option = document.createElement('option');
    option.value = String(index);
    option.textContent = (index + 1) + ' / ' + shells.length;
    select.appendChild(option);
    shell.tabIndex = -1;
    shell.setAttribute('aria-label', (index + 1) + '번 슬라이드');
    shell.setAttribute('role', 'group');
  });

  function updateControls() {
    select.value = String(current);
    previous.disabled = current === 0;
    next.disabled = current >= shells.length - 1;
    shells.forEach(function (shell, index) {
      if (index === current) shell.setAttribute('aria-current', 'true');
      else shell.removeAttribute('aria-current');
      shell.hidden = oneAtATime && index !== current;
    });
  }

  function fitSlides() {
    const maxWidth = Math.max(1, slidesHost.clientWidth - 24);
    const toolbarHeight = document.getElementById('ptv-toolbar').getBoundingClientRect().height;
    const maxHeight = Math.max(100, window.innerHeight - toolbarHeight - 48);
    const scale = Math.min(1, maxWidth / width, oneAtATime ? maxHeight / height : 1);
    for (const shell of shells) {
      const slide = shell.querySelector(':scope > .slide');
      if (!slide) continue;
      shell.style.width = (width * scale) + 'px';
      shell.style.height = (height * scale) + 'px';
      slide.style.transform = 'scale(' + scale + ')';
      slide.style.transformOrigin = 'top left';
    }
  }

  function scheduleFitSlides() {
    if (fitFrame !== null) return;
    // Resizing the children also changes the observed host's height. Write
    // their dimensions in the next frame, outside ResizeObserver delivery.
    fitFrame = window.requestAnimationFrame(function () {
      fitFrame = null;
      fitSlides();
    });
  }

  function syncSelectionToViewport() {
    if (oneAtATime || programmaticScroll) return;
    let best = current;
    let bestRatio = 0;
    const toolbarBottom = document.getElementById('ptv-toolbar').getBoundingClientRect().bottom;
    for (let index = 0; index < shells.length; index++) {
      const ratio = visibleRatios.get(shells[index]) || 0;
      if (ratio <= 0) continue;
      if (ratio > bestRatio + 0.001) {
        best = index;
        bestRatio = ratio;
      } else if (Math.abs(ratio - bestRatio) <= 0.001 && index !== best) {
        if (index === current || (best !== current &&
            Math.abs(shells[index].getBoundingClientRect().top - toolbarBottom) <
            Math.abs(shells[best].getBoundingClientRect().top - toolbarBottom))) best = index;
      }
    }
    if (bestRatio > 0 && best !== current) {
      current = best;
      updateControls();
    }
  }

  function goTo(index, focusSlide) {
    if (!shells.length) return;
    current = Math.max(0, Math.min(shells.length - 1, index));
    updateControls();
    if (!oneAtATime) {
      programmaticScroll = true;
      if (scrollTimer !== null) window.clearTimeout(scrollTimer);
      shells[current].scrollIntoView({ block: 'start', behavior: 'smooth' });
      scrollTimer = window.setTimeout(function () {
        scrollTimer = null;
        programmaticScroll = false;
        syncSelectionToViewport();
      }, 600);
    }
    if (focusSlide) shells[current].focus({ preventScroll: true });
  }

  previous.addEventListener('click', function () { goTo(current - 1, false); });
  next.addEventListener('click', function () { goTo(current + 1, false); });
  select.addEventListener('change', function () { goTo(Number(select.value), false); });
  function navigationIndex(target) {
    const link = target.closest('a[href^="#ptv-slide-"]');
    const action = target.closest('[data-slide-target], [data-target-slide]');
    const raw = link ? link.getAttribute('href').slice('#ptv-slide-'.length) :
      action && (action.getAttribute('data-slide-target') || action.getAttribute('data-target-slide'));
    const number = Number(raw);
    return Number.isInteger(number) && number >= 1 && number <= shells.length ? number - 1 : null;
  }

  slidesHost.addEventListener('click', function (event) {
    const destination = navigationIndex(event.target);
    if (destination === null) return;
    event.preventDefault();
    // Firefox uses a nonzero click count for anchor Enter but no pointer type.
    // Keep keyboard focus on the destination when the original link is hidden.
    goTo(destination, event.detail === 0 || event.pointerType === '');
  });
  slidesHost.addEventListener('keydown', function (event) {
    if (event.defaultPrevented || event.altKey || event.ctrlKey || event.metaKey ||
        (event.key !== 'Enter' && event.key !== ' ')) return;
    // Native links already synthesize their own Enter click. Only emulate the
    // keyboard behavior of the engine's non-native navigation buttons.
    if (event.target.closest('a[href], button, input, select, textarea')) return;
    const action = event.target.closest('[role="button"][data-slide-target], [role="button"][data-target-slide]');
    if (!action) return;
    const destination = navigationIndex(action);
    if (destination === null) return;
    event.preventDefault();
    goTo(destination, true);
  });
  mode.addEventListener('click', function () {
    oneAtATime = !oneAtATime;
    if (scrollTimer !== null) window.clearTimeout(scrollTimer);
    scrollTimer = null;
    programmaticScroll = false;
    body.classList.toggle('ptv-single', oneAtATime);
    mode.setAttribute('aria-pressed', String(oneAtATime));
    mode.textContent = oneAtATime ? '전체 보기' : '한 장씩';
    updateControls();
    fitSlides();
    if (!oneAtATime) goTo(current, false);
  });
  print.addEventListener('click', function () { window.print(); });

  document.addEventListener('keydown', function (event) {
    if (event.altKey || event.ctrlKey || event.metaKey) return;
    const target = event.target;
    if (target && (target.closest('input, select, textarea, button, [contenteditable="true"]') || target.isContentEditable)) return;
    let destination;
    if (event.key === 'ArrowRight' || event.key === 'ArrowDown' || event.key === 'PageDown') destination = current + 1;
    else if (event.key === 'ArrowLeft' || event.key === 'ArrowUp' || event.key === 'PageUp') destination = current - 1;
    else if (event.key === 'Home') destination = 0;
    else if (event.key === 'End') destination = shells.length - 1;
    else return;
    event.preventDefault();
    goTo(destination, true);
  });

  if ('IntersectionObserver' in window) {
    const observer = new IntersectionObserver(function (entries) {
      entries.forEach(function (entry) {
        visibleRatios.set(entry.target, entry.isIntersecting ? entry.intersectionRatio : 0);
      });
      syncSelectionToViewport();
    }, { rootMargin: '-10% 0px -25% 0px', threshold: [0, 0.25, 0.5, 0.75, 1] });
    shells.forEach(function (shell) { observer.observe(shell); });
  }
  if ('ResizeObserver' in window) new ResizeObserver(scheduleFitSlides).observe(slidesHost);
  // In single-slide mode a height-only viewport resize may leave the host's
  // dimensions unchanged, so the window event is needed even with an observer.
  window.addEventListener('resize', scheduleFitSlides);
  updateControls();
  fitSlides();
})();
