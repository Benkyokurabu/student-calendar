// Share the display preference across the student and teacher calendars.
(() => {
  const key = 'benkyo_calendar_mobile_view_v1';
  const viewport = document.querySelector('.calendar-scroll');
  const controls = document.querySelector('.calendar-mode-toggle');
  const hint = document.querySelector('.calendar-scroll-hint');
  if (!viewport || !controls || !hint) return;
  const mobile = window.matchMedia('(max-width: 600px)');
  let mode = 'compact';
  try { if (localStorage.getItem(key) === 'scroll') mode = 'scroll'; } catch { /* Storage may be disabled. */ }

  function apply() {
    viewport.dataset.view = mode;
    hint.dataset.visible = String(mode === 'scroll');
    for (const button of controls.querySelectorAll('button[data-view]')) {
      button.setAttribute('aria-pressed', String(button.dataset.view === mode));
    }
    if (mobile.matches && mode === 'scroll') viewport.setAttribute('tabindex', '0');
    else {
      viewport.removeAttribute('tabindex');
      viewport.scrollLeft = 0;
    }
  }
  controls.addEventListener('click', event => {
    const button = event.target.closest('button[data-view]');
    if (!button || !controls.contains(button)) return;
    mode = button.dataset.view === 'scroll' ? 'scroll' : 'compact';
    try { localStorage.setItem(key, mode); } catch { /* Switching still works without saving. */ }
    apply();
  });
  window.addEventListener('storage', event => {
    if (event.key !== key && event.key !== null) return;
    mode = event.newValue === 'scroll' ? 'scroll' : 'compact';
    apply();
  });
  mobile.addEventListener('change', apply);
  apply();
})();
