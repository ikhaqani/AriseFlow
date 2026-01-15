const ICONS = Object.freeze({
  success: `<svg viewBox="0 0 24 24" width="20" height="20" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"></polyline></svg>`,
  error: `<svg viewBox="0 0 24 24" width="20" height="20" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="8" x2="12" y2="12"></line><line x1="12" y1="16" x2="12.01" y2="16"></line></svg>`,
  info: `<svg viewBox="0 0 24 24" width="20" height="20" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="16" x2="12" y2="12"></line><line x1="12" y1="8" x2="12.01" y2="8"></line></svg>`,
  save: `<svg viewBox="0 0 24 24" width="20" height="20" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"></path><polyline points="17 21 17 13 7 13 7 21"></polyline><polyline points="7 3 7 8 15 8"></polyline></svg>`,
  close: `<svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>`
});

function ensureContainer() {
  const existing = document.getElementById('toast-container');
  if (existing) return existing;

  const el = document.createElement('div');
  el.id = 'toast-container';
  document.body.appendChild(el);
  return el;
}

function removeToastEl(toast) {
  if (!toast) return;

  toast.classList.remove('visible');
  toast.classList.add('removing');

  const onDone = () => {
    toast.removeEventListener('transitionend', onDone);
    if (toast.parentElement) toast.remove();
  };

  toast.addEventListener('transitionend', onDone);

  setTimeout(() => {
    if (toast.parentElement) onDone();
  }, 600);
}

function createToastElement(message, type, duration) {
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;

  const content = document.createElement('div');
  content.className = 'toast-content';

  const icon = document.createElement('span');
  icon.className = 'toast-icon';
  icon.innerHTML = ICONS[type] || ICONS.info;

  const msg = document.createElement('span');
  msg.className = 'toast-msg';
  msg.textContent = String(message ?? '');

  content.appendChild(icon);
  content.appendChild(msg);

  const closeBtn = document.createElement('button');
  closeBtn.className = 'toast-close';
  closeBtn.type = 'button';
  closeBtn.setAttribute('aria-label', 'Sluiten');
  closeBtn.innerHTML = ICONS.close;

  const progress = document.createElement('div');
  progress.className = 'toast-progress';
  progress.style.animationDuration = `${duration}ms`;

  toast.appendChild(content);
  toast.appendChild(closeBtn);
  toast.appendChild(progress);

  return toast;
}

export const Toast = {
  container: null,
  maxToasts: 5,

  init() {
    this.container = ensureContainer();
  },

  show(message, type = 'info', duration = 3000) {
    if (!this.container) this.init();

    const t = String(type || 'info');
    const safeType = ICONS[t] ? t : 'info';
    const ms = Number.isFinite(Number(duration)) ? Math.max(400, Number(duration)) : 3000;

    if (this.container.childElementCount >= this.maxToasts) {
      const oldest = this.container.firstElementChild;
      if (oldest) removeToastEl(oldest);
    }

    const toast = createToastElement(message, safeType, ms);

    const closeBtn = toast.querySelector('.toast-close');
    if (closeBtn) closeBtn.onclick = () => removeToastEl(toast);

    let timerId = null;

    const startTimer = () => {
      clearTimeout(timerId);
      timerId = setTimeout(() => removeToastEl(toast), ms);
      const bar = toast.querySelector('.toast-progress');
      if (bar) bar.style.animationPlayState = 'running';
    };

    const stopTimer = () => {
      clearTimeout(timerId);
      const bar = toast.querySelector('.toast-progress');
      if (bar) bar.style.animationPlayState = 'paused';
    };

    toast.addEventListener('mouseenter', stopTimer);
    toast.addEventListener('mouseleave', startTimer);

    this.container.appendChild(toast);

    requestAnimationFrame(() => {
      toast.classList.add('visible');
      startTimer();
    });
  },

  removeToast(toast) {
    removeToastEl(toast);
  }
};