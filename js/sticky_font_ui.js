/**
 * sticky_font_ui.js
 * UI control to set post-it font size via CSS variable --sticky-font-size
 * Persists in localStorage.
 */

const LS_KEY = "ariseflow.stickyFontSizePx";
const CSS_VAR = "--sticky-font-size";

function clamp(n, a, b) {
  const x = Number(n);
  if (!Number.isFinite(x)) return a;
  return Math.max(a, Math.min(b, x));
}

function getCurrentPx(fallbackPx = 18) {
  const raw = localStorage.getItem(LS_KEY);
  const v = raw != null ? Number(raw) : NaN;
  if (Number.isFinite(v)) return clamp(v, 12, 48);

  const css = getComputedStyle(document.documentElement).getPropertyValue(CSS_VAR).trim();
  const m = css.match(/^(\d+)\s*px$/i);
  if (m) return clamp(Number(m[1]), 12, 48);

  return fallbackPx;
}

function applyStickyFontSize(px) {
  const v = clamp(px, 12, 48);
  document.documentElement.style.setProperty(CSS_VAR, `${v}px`);
  localStorage.setItem(LS_KEY, String(v));
}

function findHost() {
  return (
    document.querySelector("#topbar .topbar-right") ||
    document.querySelector("#topbar") ||
    document.querySelector(".topbar") ||
    document.querySelector(".toolbar") ||
    document.querySelector("#toolbar") ||
    document.body
  );
}

function ensureOnce() {
  const existing = document.getElementById("stickyFontSizeSelect");
  return !existing;
}

export function initStickyFontSizeUI() {
  // always apply persisted size early
  const startPx = getCurrentPx(18);
  applyStickyFontSize(startPx);

  if (!ensureOnce()) return;

  const host = findHost();

  const wrap = document.createElement("div");
  wrap.className = "sticky-font-control";
  wrap.innerHTML = `
    <label class="sticky-font-label" for="stickyFontSizeSelect">Post-it tekst</label>
    <select id="stickyFontSizeSelect" class="sticky-font-select" aria-label="Post-it lettergrootte">
      ${[14,16,18,20,22,24,26,28,32,36].map(v => `<option value="${v}">${v}px</option>`).join("")}
    </select>
  `;

  host.appendChild(wrap);

  const sel = wrap.querySelector("#stickyFontSizeSelect");
  sel.value = String(startPx);
  sel.addEventListener("change", () => applyStickyFontSize(sel.value));
}

// optional: export for other modules/tests
export function _applyStickyFontSize(px) {
  applyStickyFontSize(px);
}
