const ICONS = Object.freeze({
  success: `<svg viewBox="0 0 24 24" width="20" height="20" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"></polyline></svg>`,
  error: `<svg viewBox="0 0 24 24" width="20" height="20" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="8" x2="12" y2="12"></line><line x1="12" y1="16" x2="12.01" y2="16"></line></svg>`,
  info: `<svg viewBox="0 0 24 24" width="20" height="20" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="16" x2="12" y2="12"></line><line x1="12" y1="8" x2="12.01" y2="8"></line></svg>`,
  save: `<svg viewBox="0 0 24 24" width="20" height="20" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"></path><polyline points="17 21 17 13 7 13 7 21"></polyline><polyline points="7 3 7 8 15 8"></polyline></svg>`,
  close: `<svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>`
});

/** Ensures the toast container exists and returns it. */
const ensureContainer = () => {
  let el = document.getElementById('toast-container');
  if (el) return el;

  el = document.createElement('div');
  el.id = 'toast-container';
  document.body.appendChild(el);
  return el;
};

/** Creates a toast DOM element for the given payload. */
const createToastElement = ({ message, type, duration }) => {
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.innerHTML = `
    <div class="toast-content">
      <span class="toast-icon">${ICONS[type] || ICONS.info}</span>
      <span class="toast-msg">${message}</span>
    </div>
    <button class="toast-close" type="button" aria-label="Sluiten">${ICONS.close}</button>
    <div class="toast-progress" style="animation-duration:${duration}ms"></div>
  `;
  return toast;
};

/** Removes a toast element with transition cleanup. */
const removeToast = (toast) => {
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
  }, 500);
};

export const Toast = {
  container: null,
  maxToasts: 5,

  /** Initializes the toast container reference. */
  init() {
    this.container = ensureContainer();
  },

  /** Shows a toast message with type and auto-dismiss duration. */
  show(message, type = 'info', duration = 3000) {
    if (!this.container) this.init();

    if (this.container.childElementCount >= this.maxToasts) {
      const oldest = this.container.firstElementChild;
      if (oldest) removeToast(oldest);
    }

    const toast = createToastElement({ message, type, duration });

    const closeBtn = toast.querySelector('.toast-close');
    if (closeBtn) closeBtn.onclick = () => removeToast(toast);

    let timerId = null;

    /** Starts or restarts the auto-dismiss timer. */
    const startTimer = () => {
      clearTimeout(timerId);
      timerId = setTimeout(() => removeToast(toast), duration);
      const bar = toast.querySelector('.toast-progress');
      if (bar) bar.style.animationPlayState = 'running';
    };

    /** Pauses the auto-dismiss timer. */
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

  /** Removes a toast programmatically. */
  removeToast(toast) {
    removeToast(toast);
  }
};