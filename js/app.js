import { state } from './state.js';
import { renderBoard, setupDelegatedEvents, applyStateUpdate } from './dom.js';
import { openEditModal, saveModalDetails } from './modals.js';
import { saveToFile, loadFromFile, exportToCSV, exportHD } from './io.js';
import { Toast } from './toast.js';

const $ = (id) => document.getElementById(id);

const pickEl = (...selectors) => {
  for (const sel of selectors) {
    if (!sel) continue;
    const byId = document.getElementById(sel);
    if (byId) return byId;
    const bySel = document.querySelector(sel);
    if (bySel) return bySel;
  }
  return null;
};

const bindClickEl = (el, handler) => {
  if (!el) return;
  el.addEventListener('click', (e) => {
    e.preventDefault();
    handler(e);
  });
};

const bindClick = (id, handler) => {
  const el = $(id);
  if (!el) return;
  el.addEventListener('click', (e) => {
    e.preventDefault();
    handler(e);
  });
};

const safeToast = (msg, type = 'info', ms) => {
  if (!Toast || typeof Toast.show !== 'function') return;
  Toast.show(msg, type, ms);
};

const initToast = () => {
  try {
    if (Toast && typeof Toast.init === 'function') Toast.init();
  } catch (e) {
    console.warn("Toast init failed", e);
  }
};

const syncOpenModal = () => {
  const modal = $("editModal");
  if (!modal) return;
  const isOpen = modal.style.display && modal.style.display !== "none";
  if (!isOpen) return;

  try {
    saveModalDetails(false);
  } catch (e) {
    console.warn("Modal sync failed", e);
  }
};

const renameActiveSheet = () => {
  syncOpenModal();

  const currentName = state.activeSheet?.name || '';
  const newName = prompt("Hernoem proces:", currentName);
  if (!newName || !newName.trim() || newName.trim() === currentName) return;
  state.renameSheet(newName.trim());
  safeToast("Naam gewijzigd", "success");
};

/* ============================================================
   MERGE LAYER (for spanning/merged stickies)
   - Renders an overlay sticky that spans multiple columns.
   - Hides the underlying duplicated stickies.
   ============================================================ */

let mergeRefreshTimer = null;

const ensureMergeLayer = () => {
  const wrapper = $("board-content-wrapper");
  if (!wrapper) return null;

  // Ensure wrapper can be positioning context
  const wrapperStyle = getComputedStyle(wrapper);
  if (wrapperStyle.position === "static") {
    wrapper.style.position = "relative";
  }

  let layer = $("merge-layer");
  if (!layer) {
    layer = document.createElement("div");
    layer.id = "merge-layer";
    layer.setAttribute("aria-hidden", "true");
    wrapper.insertBefore(layer, $("cols") || null);
  }

  // Ensure overlay styling (works even if CSS not yet added)
  Object.assign(layer.style, {
    position: "absolute",
    inset: "0",
    pointerEvents: "none",
    zIndex: "50",
  });

  return layer;
};

const getColumnElements = () => {
  const colsRoot = $("cols");
  if (!colsRoot) return [];
  return Array.from(colsRoot.querySelectorAll(":scope > .col"));
};

const getStickyAt = (colEl, slotIdx) => {
  if (!colEl) return null;
  const slots = colEl.querySelector(".slots");
  if (!slots) return null;

  // SIPOC order: 0 Lev, 1 Sys, 2 Input, 3 Proces, 4 Output, 5 Klant
  const slot = slots.querySelector(`:scope > .slot:nth-child(${slotIdx + 1})`);
  if (!slot) return null;

  return slot.querySelector(":scope > .sticky");
};

const setUnderlyingVisibility = (stickyEls, visible) => {
  for (const el of stickyEls) {
    if (!el) continue;
    if (visible) {
      el.style.visibility = "";
      el.style.opacity = "";
    } else {
      // Keep layout (so bounding boxes remain stable if needed) but hide visually
      el.style.visibility = "hidden";
      el.style.opacity = "0";
    }
  }
};

const buildMergedStickyEl = (textValue) => {
  const merged = document.createElement("div");
  merged.id = "merged-approval-output";
  merged.className = "sticky merged-sticky";
  merged.style.position = "absolute";
  merged.style.pointerEvents = "auto"; // allow click forwarding

  // Inner structure matching existing sticky DOM
  merged.innerHTML = `
    <div class="sticky-grip"></div>
    <div class="sticky-content">
      <div class="text" spellcheck="false"></div>
    </div>
    <div class="sticky-badge label-tr">MERGED</div>
  `;

  const textEl = merged.querySelector(".text");
  if (textEl) textEl.textContent = textValue || "";

  return merged;
};

const syncMergedApprovalOutput = () => {
  const layer = ensureMergeLayer();
  if (!layer) return;

  // Clear previous merged overlays
  layer.innerHTML = "";

  const cols = getColumnElements();
  if (cols.length < 4) return; // need at least columns 1..4

  // Target: merge Output row for columns 2,3,4 (1-based); => indices 1..3 (0-based)
  const SLOT_OUTPUT = 4;
  const targetColIdxs = [1, 2, 3];

  const stickyEls = targetColIdxs.map((i) => getStickyAt(cols[i], SLOT_OUTPUT));
  if (stickyEls.some((el) => !el)) return;

  // Use data text if possible (preferred); fallback to DOM text.
  const dataTexts = targetColIdxs.map((i) => {
    const t = state.activeSheet?.columns?.[i]?.slots?.[SLOT_OUTPUT]?.text;
    return typeof t === "string" ? t.trim() : "";
  });

  const domTexts = stickyEls.map((el) => {
    const t = el?.querySelector(".text")?.textContent;
    return typeof t === "string" ? t.trim() : "";
  });

  const texts = dataTexts.every((t) => t) ? dataTexts : domTexts;
  const nonEmpty = texts.filter((t) => t.length > 0);
  if (nonEmpty.length === 0) return;

  // Only merge if all three are identical (after trim)
  const first = texts[0];
  const allSame = texts.every((t) => t === first);
  if (!allSame) {
    // if not identical, do not merge; ensure originals are visible
    setUnderlyingVisibility(stickyEls, true);
    return;
  }

  // Compute union rect across the 3 stickies
  const layerRect = layer.getBoundingClientRect();
  const rects = stickyEls.map((el) => el.getBoundingClientRect());

  const left = Math.min(...rects.map((r) => r.left));
  const top = Math.min(...rects.map((r) => r.top));
  const right = Math.max(...rects.map((r) => r.right));
  const bottom = Math.max(...rects.map((r) => r.bottom));

  const mergedEl = buildMergedStickyEl(first);

  // Style to match the stickies
  // Note: use small inset so it doesn’t overlap borders harshly
  const inset = 0;
  mergedEl.style.left = `${left - layerRect.left + inset}px`;
  mergedEl.style.top = `${top - layerRect.top + inset}px`;
  mergedEl.style.width = `${right - left - inset * 2}px`;
  mergedEl.style.height = `${bottom - top - inset * 2}px`;

  // Forward click to the first underlying sticky (so your delegated edit/modal behavior still works)
  mergedEl.addEventListener("click", (e) => {
    e.preventDefault();
    e.stopPropagation();
    const target = stickyEls[0];
    if (target) {
      target.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }));
    }
  });

  layer.appendChild(mergedEl);

  // Hide underlying duplicates
  setUnderlyingVisibility(stickyEls, false);
};

const scheduleMergeRefresh = (delay = 0) => {
  if (mergeRefreshTimer) clearTimeout(mergeRefreshTimer);
  mergeRefreshTimer = setTimeout(() => {
    try {
      syncMergedApprovalOutput();
    } catch (e) {
      console.warn("Merge refresh failed", e);
    }
  }, delay);
};

/* ============================================================
   UI SETUP
   ============================================================ */

const setupProjectTitle = () => {
  const titleInput = $("boardTitle");
  if (!titleInput) return;

  titleInput.addEventListener("input", (e) => {
    state.updateProjectTitle(e.target.value);
  });

  titleInput.addEventListener("blur", () => {
    state.updateProjectTitle(titleInput.value);
  });
};

const setupSheetControls = () => {
  const sheetSelect = $("sheetSelect");
  if (sheetSelect) {
    sheetSelect.addEventListener("change", (e) => {
      syncOpenModal();
      state.setActiveSheet(e.target.value);
      safeToast(`Gewisseld naar: ${state.activeSheet.name}`, "info", 1000);
      scheduleMergeRefresh(50);
    });
  }

  const btnRename = pickEl("btnRenameSheet", "#btnRenameSheet", "[data-action='rename-sheet']");
  bindClickEl(btnRename, renameActiveSheet);

  document.addEventListener("dblclick", (e) => {
    if (e.target && e.target.id === "board-header-display") renameActiveSheet();
  });

  const btnAdd = pickEl("btnAddSheet", "#btnAddSheet", "[data-action='add-sheet']", "#addSheetBtn");
  bindClickEl(btnAdd, () => {
    syncOpenModal();
    const name = prompt("Nieuw procesblad naam:", `Proces ${state.data.sheets.length + 1}`);
    if (!name || !name.trim()) return;
    state.addSheet(name.trim());
    safeToast("Procesblad toegevoegd", "success");
    scheduleMergeRefresh(50);
  });

  const btnDel = pickEl("btnDelSheet", "#btnDelSheet", "[data-action='delete-sheet']", "#deleteSheetBtn");
  bindClickEl(btnDel, () => {
    syncOpenModal();

    if (state.data.sheets.length <= 1) {
      safeToast("Laatste blad kan niet verwijderd worden", "error");
      return;
    }
    if (!confirm(`Weet je zeker dat je "${state.activeSheet.name}" wilt verwijderen?`)) return;
    state.deleteSheet();
    safeToast("Procesblad verwijderd", "info");
    scheduleMergeRefresh(50);
  });

  document.addEventListener("visibilitychange", () => {
    if (document.hidden) syncOpenModal();
  });
};

const setupToolbarActions = () => {
  bindClick("saveBtn", async () => {
    syncOpenModal();
    await saveToFile();
    safeToast("Project opgeslagen", "success");
  });

  bindClick("exportCsvBtn", () => {
    syncOpenModal();
    exportToCSV();
    safeToast("Excel export gereed", "success");
  });

  bindClick("exportBtn", async () => {
    syncOpenModal();
    await exportHD();
    safeToast("Screenshot gemaakt", "success");
  });

  bindClick("loadBtn", () => {
    const inp = $("fileInput");
    if (inp) inp.click();
  });

  const fileInput = $("fileInput");
  if (fileInput) {
    fileInput.addEventListener("change", (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      loadFromFile(file, () => {
        fileInput.value = "";
        safeToast("Project geladen", "success");
        scheduleMergeRefresh(100);
      });
    });
  }

  bindClick("clearBtn", () => {
    syncOpenModal();
    if (!confirm("⚠️ Pas op: Alles wissen?")) return;
    localStorage.clear();
    location.reload();
  });
};

const setupZoom = () => {
  let zoomLevel = 1;

  const updateZoom = () => {
    const boardEl = $("board");
    if (!boardEl) return;

    zoomLevel = Math.max(0.4, Math.min(2.0, zoomLevel));
    boardEl.style.transform = `scale(${zoomLevel})`;

    const zoomDisplay = $("zoomDisplay");
    if (zoomDisplay) zoomDisplay.textContent = `${Math.round(zoomLevel * 100)}%`;

    // Trigger layout recalculation for overlays
    setTimeout(() => {
      window.dispatchEvent(new Event("resize"));
      scheduleMergeRefresh(0);
    }, 50);
  };

  bindClick("zoomIn", () => {
    zoomLevel += 0.1;
    updateZoom();
  });

  bindClick("zoomOut", () => {
    zoomLevel -= 0.1;
    updateZoom();
  });

  updateZoom();
};

const setupMenuToggle = () => {
  bindClick("menuToggle", () => {
    const topbar = $("topbar");
    if (!topbar) return;

    const viewportEl = $("viewport");
    const icon = document.querySelector("#menuToggle .toggle-icon");

    const isCollapsed = topbar.classList.toggle("collapsed");
    if (viewportEl) viewportEl.classList.toggle("expanded-view");
    if (icon) icon.style.transform = isCollapsed ? "rotate(180deg)" : "rotate(0deg)";

    scheduleMergeRefresh(50);
  });
};

const setupColumnManager = () => {
  bindClick("btnManageCols", () => {
    syncOpenModal();

    const list = $("colManagerList");
    const modal = $("colManagerModal");
    if (!list || !modal) return;

    list.innerHTML = "";

    state.activeSheet.columns.forEach((col, idx) => {
      const raw = col.slots?.[3]?.text || "";
      const procText = raw
        ? `${raw.substring(0, 25)}${raw.length > 25 ? "..." : ""}`
        : "<i>(Leeg)</i>";

      const item = document.createElement("div");
      item.className = "col-manager-item";
      item.innerHTML = `
        <span style="font-size:13px; color:#ddd;">
          <strong>Kolom ${idx + 1}</strong>: ${procText}
        </span>
        <input type="checkbox" ${col.isVisible !== false ? "checked" : ""} style="cursor:pointer; transform:scale(1.2);">
      `;

      const checkbox = item.querySelector("input");
      checkbox.addEventListener("change", (e) => {
        state.setColVisibility(idx, e.target.checked);
        scheduleMergeRefresh(50);
      });

      list.appendChild(item);
    });

    modal.style.display = "grid";
  });

  bindClick("colManagerCloseBtn", () => {
    const modal = $("colManagerModal");
    if (modal) modal.style.display = "none";
    scheduleMergeRefresh(50);
  });
};

const setupModals = () => {
  bindClick("modalSaveBtn", () => {
    saveModalDetails(true);
    safeToast("Wijzigingen opgeslagen", "save");
    scheduleMergeRefresh(50);
  });

  bindClick("modalCancelBtn", () => {
    const m = $("editModal");
    if (m) m.style.display = "none";
  });
};

const setupStateSubscription = () => {
  const titleInput = $("boardTitle");

  state.subscribe((_, meta) => {
    applyStateUpdate(meta, openEditModal);

    if (titleInput && state.data.projectTitle && document.activeElement !== titleInput) {
      titleInput.value = state.data.projectTitle;
    }

    document.title = state.data.projectTitle ? `${state.data.projectTitle} - SIPOC` : "SIPOC Board";

    const header = $("board-header-display");
    if (header) {
      header.style.cursor = "pointer";
      header.title = "Dubbelklik om naam te wijzigen";
    }

    // After any state-driven render/update, resync merged overlay
    scheduleMergeRefresh(80);
  });
};

const setupGlobalHotkeys = () => {
  document.addEventListener("keydown", (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === "s") {
      e.preventDefault();
      syncOpenModal();
      saveToFile();
      safeToast("Quick Save", "save");
    }

    if (e.key === "Escape") {
      document.querySelectorAll(".modal-overlay").forEach((m) => (m.style.display = "none"));
    }
  });
};

const setupOverlayResyncOnResize = () => {
  window.addEventListener("resize", () => scheduleMergeRefresh(0));
  window.addEventListener("scroll", () => scheduleMergeRefresh(0), true);
};

const initApp = () => {
  initToast();
  setupDelegatedEvents();
  setupProjectTitle();
  setupSheetControls();
  setupToolbarActions();
  setupZoom();
  setupMenuToggle();
  setupColumnManager();
  setupModals();
  setupStateSubscription();
  setupGlobalHotkeys();
  setupOverlayResyncOnResize();

  renderBoard(openEditModal);

  const titleInput = $("boardTitle");
  if (titleInput) titleInput.value = state.data.projectTitle;

  // Initial merged overlay pass (after first paint)
  scheduleMergeRefresh(120);

  setTimeout(() => safeToast("Klaar voor gebruik", "info"), 500);
};

document.addEventListener("DOMContentLoaded", initApp);