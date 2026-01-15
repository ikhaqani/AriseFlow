// state.js  (VOLLEDIG, met isVariant + System-Merge + Multi-System support)
// -----------------------------------------------------------
import {
  createProjectState,
  createSheet,
  createColumn,
  createSticky,
  STORAGE_KEY,
  uid
} from './config.js';

class StateManager {
  constructor() {
    this.project = createProjectState();
    this.listeners = new Set();
    this.saveCallbacks = new Set();
    this.historyStack = [];
    this.redoStack = [];
    this.maxHistory = 20;

    this._lastNotifyTs = 0;
    this._notifyQueued = false;
    this._pendingMeta = null;
    this._suspendNotify = 0;

    this._lastHistoryTs = 0;

    this.loadFromStorage();
  }

  /* =========================================================
     ID helpers (for stable process IDs used by gates/checks)
     ========================================================= */

  _makeId(prefix = 'id') {
    try {
      if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
        return `${prefix}_${crypto.randomUUID()}`;
      }
    } catch (_) {}
    return `${prefix}_${Date.now()}_${Math.random().toString(16).slice(2)}`;
  }

  _ensureProcessIds(project) {
    const sheets = project?.sheets || [];
    sheets.forEach((sheet) => {
      (sheet.columns || []).forEach((col) => {
        const slot = col?.slots?.[3];
        if (!slot) return;
        if (!slot.id || String(slot.id).trim() === '') {
          slot.id = this._makeId('proc');
        }
      });
    });
  }

  /* =========================================================
     Multi-System helpers (SYSTEM row only)
     ========================================================= */

  _sanitizeSystemEntry(s) {
    const x = s && typeof s === 'object' ? s : {};
    return {
      id: String(x.id || uid()),
      name: String(x.name || '').trim(),
      isLegacy: !!x.isLegacy,
      futureSystem: String(x.futureSystem || '').trim(),
      // Answers per system (SYSTEM_QUESTIONS keys -> values)
      qa: x.qa && typeof x.qa === 'object' ? { ...x.qa } : {},
      // Per-system score (optional; modals.js will compute/update)
      calculatedScore:
        Number.isFinite(Number(x.calculatedScore)) ? Number(x.calculatedScore) : null
    };
  }

  _ensureMultiSystemShape(slot) {
    if (!slot || typeof slot !== 'object') return;

    if (!slot.systemData || typeof slot.systemData !== 'object') slot.systemData = {};

    const sd = slot.systemData;

    // Flag: does this process step use multiple systems?
    if (typeof sd.isMulti !== 'boolean') sd.isMulti = false;

    // Array of system entries
    if (!Array.isArray(sd.systems)) {
      // Backward compat: if old single-name exists, keep it as first system
      const legacyName = typeof sd.systemName === 'string' ? sd.systemName.trim() : '';
      sd.systems = legacyName
        ? [this._sanitizeSystemEntry({ name: legacyName })]
        : [this._sanitizeSystemEntry({})];
    } else {
      sd.systems = sd.systems.map((s) => this._sanitizeSystemEntry(s));
      if (sd.systems.length === 0) sd.systems = [this._sanitizeSystemEntry({})];
    }

    // Keep a stable overall score field (badge can show this)
    if (!Number.isFinite(Number(sd.calculatedScore))) sd.calculatedScore = null;

    // Optional convenience: keep a single systemName (first one) for older UI parts
    if (typeof sd.systemName !== 'string') {
      sd.systemName = sd.systems?.[0]?.name || '';
    }
  }

  /* =========================================================
     Merge helpers (OUTPUT row + SYSTEM row)
     ========================================================= */

  _normalizeMergeRanges(sheet, key, slotIdx) {
    const colsLen = sheet?.columns?.length ?? 0;
    if (!Array.isArray(sheet[key])) sheet[key] = [];

    let ranges = sheet[key]
      .filter((r) => r && r.slotIdx === slotIdx)
      .map((r) => ({
        id: r.id || uid(),
        slotIdx,
        startCol: Math.max(0, Math.min(colsLen - 1, Number(r.startCol))),
        endCol: Math.max(0, Math.min(colsLen - 1, Number(r.endCol)))
      }))
      .map((r) => {
        if (r.startCol > r.endCol) [r.startCol, r.endCol] = [r.endCol, r.startCol];
        return r;
      })
      .filter((r) => Number.isFinite(r.startCol) && Number.isFinite(r.endCol))
      .filter((r) => (r.endCol - r.startCol) >= 1); // must be >=2 columns

    // Union overlaps
    ranges.sort((a, b) => a.startCol - b.startCol);
    const merged = [];
    for (const r of ranges) {
      const last = merged[merged.length - 1];
      if (!last) merged.push(r);
      else if (r.startCol <= last.endCol) {
        last.endCol = Math.max(last.endCol, r.endCol);
      } else {
        merged.push(r);
      }
    }

    sheet[key] = merged;
  }

  /* ===== OUTPUT merges ===== */

  _normalizeOutputMerges(sheet) {
    this._normalizeMergeRanges(sheet, 'outputMerges', 4);
  }

  getOutputMergeForCol(colIdx) {
    const sheet = this.activeSheet;
    if (!sheet || !Array.isArray(sheet.outputMerges)) return null;
    return (
      sheet.outputMerges.find(
        (r) => r.slotIdx === 4 && colIdx >= r.startCol && colIdx <= r.endCol
      ) || null
    );
  }

  setOutputMergeRangeForCol(colIdx, startCol, endCol) {
    const sheet = this.activeSheet;
    if (!sheet) return;

    this.pushHistory();

    if (!Array.isArray(sheet.outputMerges)) sheet.outputMerges = [];

    const s = Number(startCol);
    const e = Number(endCol);

    // remove existing range(s) that include this col, or overlap the new one
    sheet.outputMerges = sheet.outputMerges.filter((r) => {
      if (!r || r.slotIdx !== 4) return true;
      const overlaps = !(e < r.startCol || s > r.endCol);
      const includesCol = colIdx >= r.startCol && colIdx <= r.endCol;
      return !(overlaps || includesCol);
    });

    // if start==end => no merge (remove only)
    if (Number.isFinite(s) && Number.isFinite(e) && s !== e) {
      sheet.outputMerges.push({
        id: uid(),
        slotIdx: 4,
        startCol: Math.min(s, e),
        endCol: Math.max(s, e)
      });
    }

    this._normalizeOutputMerges(sheet);
    this.notify({ reason: 'columns' }, { clone: false });
  }

  /* ===== SYSTEM merges (NEW) ===== */

  _normalizeSystemMerges(sheet) {
    this._normalizeMergeRanges(sheet, 'systemMerges', 1);
  }

  getSystemMergeForCol(colIdx) {
    const sheet = this.activeSheet;
    if (!sheet || !Array.isArray(sheet.systemMerges)) return null;
    return (
      sheet.systemMerges.find(
        (r) => r.slotIdx === 1 && colIdx >= r.startCol && colIdx <= r.endCol
      ) || null
    );
  }

  setSystemMergeRangeForCol(colIdx, startCol, endCol) {
    const sheet = this.activeSheet;
    if (!sheet) return;

    this.pushHistory();

    if (!Array.isArray(sheet.systemMerges)) sheet.systemMerges = [];

    const s = Number(startCol);
    const e = Number(endCol);

    sheet.systemMerges = sheet.systemMerges.filter((r) => {
      if (!r || r.slotIdx !== 1) return true;
      const overlaps = !(e < r.startCol || s > r.endCol);
      const includesCol = colIdx >= r.startCol && colIdx <= r.endCol;
      return !(overlaps || includesCol);
    });

    if (Number.isFinite(s) && Number.isFinite(e) && s !== e) {
      sheet.systemMerges.push({
        id: uid(),
        slotIdx: 1,
        startCol: Math.min(s, e),
        endCol: Math.max(s, e)
      });
    }

    this._normalizeSystemMerges(sheet);
    this.notify({ reason: 'columns' }, { clone: false });
  }

  /* =========================================================
     Subscribe/notify/history (unchanged)
     ========================================================= */

  subscribe(listenerFn) {
    this.listeners.add(listenerFn);
    return () => this.listeners.delete(listenerFn);
  }

  onSave(fn) {
    this.saveCallbacks.add(fn);
    return () => this.saveCallbacks.delete(fn);
  }

  beginBatch(meta = { reason: 'batch' }) {
    this._suspendNotify++;
    this._pendingMeta = this._mergeMeta(this._pendingMeta, meta);
  }

  endBatch(meta = null) {
    this._suspendNotify = Math.max(0, this._suspendNotify - 1);
    if (meta) this._pendingMeta = this._mergeMeta(this._pendingMeta, meta);
    if (this._suspendNotify === 0) this.notify(this._pendingMeta || { reason: 'batch' });
  }

  _mergeMeta(a, b) {
    if (!a) return b ? { ...b } : null;
    if (!b) return a ? { ...a } : null;
    return { ...a, ...b };
  }

  notify(meta = { reason: 'full' }, { clone = false, throttleMs = 0 } = {}) {
    if (this._suspendNotify > 0) {
      this._pendingMeta = this._mergeMeta(this._pendingMeta, meta);
      return;
    }

    const now = performance.now();

    if (throttleMs > 0 && now - this._lastNotifyTs < throttleMs) {
      this._pendingMeta = this._mergeMeta(this._pendingMeta, meta);
      if (this._notifyQueued) return;
      this._notifyQueued = true;

      requestAnimationFrame(() => {
        this._notifyQueued = false;
        const merged = this._pendingMeta || meta;
        this._pendingMeta = null;
        this._lastNotifyTs = performance.now();
        this._emit(merged, { clone });
      });

      return;
    }

    this._lastNotifyTs = now;
    this._emit(meta, { clone });
  }

  _emit(meta, { clone }) {
    const payload = clone
      ? typeof structuredClone === 'function'
        ? structuredClone(this.project)
        : JSON.parse(JSON.stringify(this.project))
      : this.project;

    this.listeners.forEach((fn) => fn(payload, meta));
  }

  pushHistory({ throttleMs = 0 } = {}) {
    if (throttleMs > 0) {
      const now = performance.now();
      if (this._lastHistoryTs && now - this._lastHistoryTs < throttleMs) return;
      this._lastHistoryTs = now;
    }

    const snapshot = JSON.stringify(this.project);
    const last = this.historyStack[this.historyStack.length - 1];
    if (last === snapshot) return;

    this.historyStack.push(snapshot);
    if (this.historyStack.length > this.maxHistory) this.historyStack.shift();
    this.redoStack = [];
  }

  undo() {
    if (this.historyStack.length === 0) return false;
    this.redoStack.push(JSON.stringify(this.project));
    this.project = JSON.parse(this.historyStack.pop());
    this.notify({ reason: 'full' }, { clone: false });
    return true;
  }

  redo() {
    if (this.redoStack.length === 0) return false;
    this.historyStack.push(JSON.stringify(this.project));
    this.project = JSON.parse(this.redoStack.pop());
    this.notify({ reason: 'full' }, { clone: false });
    return true;
  }

  loadFromStorage() {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (!saved) return;
      const parsed = JSON.parse(saved);
      this.project = this.sanitizeProjectData(parsed);
    } catch (e) {
      console.error('Critical: Failed to load project state.', e);
    }
  }

  sanitizeProjectData(data) {
    const fresh = createProjectState();
    const merged = { ...fresh, ...data };

    if (!Array.isArray(merged.sheets) || merged.sheets.length === 0) {
      merged.sheets = fresh.sheets;
      this._ensureProcessIds(merged);
      return merged;
    }

    merged.sheets.forEach((sheet) => {
      // Ensure merges exist + normalize
      if (!Array.isArray(sheet.outputMerges)) sheet.outputMerges = [];
      if (!Array.isArray(sheet.systemMerges)) sheet.systemMerges = [];
      this._normalizeOutputMerges(sheet);
      this._normalizeSystemMerges(sheet);

      if (!Array.isArray(sheet.columns) || sheet.columns.length === 0) {
        sheet.columns = [createColumn()];
      }

      sheet.columns.forEach((col) => {
        // Ensure isVariant exists
        if (typeof col.isVariant !== 'boolean') col.isVariant = false;

        if (!Array.isArray(col.slots) || col.slots.length !== 6) {
          col.slots = Array(6)
            .fill(null)
            .map(() => createSticky());
          return;
        }

        col.slots = col.slots.map((slot, slotIdx) => {
          const clean = createSticky();
          const s = slot || {};

          const cleanGate = clean.gate && typeof clean.gate === 'object' ? clean.gate : {};
          const sGate = s.gate && typeof s.gate === 'object' ? s.gate : {};

          const mergedSlot = {
            ...clean,
            ...s,

            qa: { ...clean.qa, ...(s.qa || {}) },

            // systemData stays object, we will enrich with multi-system defaults below
            systemData: { ...clean.systemData, ...(s.systemData || {}) },

            linkedSourceId: s.linkedSourceId ?? clean.linkedSourceId,
            inputDefinitions: Array.isArray(s.inputDefinitions)
              ? s.inputDefinitions
              : clean.inputDefinitions,
            disruptions: Array.isArray(s.disruptions) ? s.disruptions : clean.disruptions,

            workExp: s.workExp ?? clean.workExp,
            workExpNote: s.workExpNote ?? clean.workExpNote,

            id: s.id ?? clean.id,
            isGate: s.isGate ?? clean.isGate,
            gate: { ...cleanGate, ...sGate }
          };

          // Ensure process-id exists for Proces row
          if (slotIdx === 3 && (!mergedSlot.id || String(mergedSlot.id).trim() === '')) {
            mergedSlot.id = this._makeId('proc');
          }

          // Ensure gate structure
          if (mergedSlot.gate) {
            if (!Array.isArray(mergedSlot.gate.checkProcessIds)) mergedSlot.gate.checkProcessIds = [];
            mergedSlot.gate.onFailTargetProcessId = mergedSlot.gate.onFailTargetProcessId || null;
            mergedSlot.gate.onPassTargetProcessId = mergedSlot.gate.onPassTargetProcessId || null;
            mergedSlot.gate.rule = mergedSlot.gate.rule || 'ALL_OK';
            mergedSlot.gate.note = mergedSlot.gate.note || '';
          }

          // NEW: ensure multi-system shape on SYSTEM row only
          if (slotIdx === 1) {
            this._ensureMultiSystemShape(mergedSlot);
          }

          return mergedSlot;
        });
      });
    });

    this._ensureProcessIds(merged);
    return merged;
  }

  saveToStorage() {
    try {
      this.project.lastModified = new Date().toISOString();
      this.saveCallbacks.forEach((fn) => fn('saving'));
      localStorage.setItem(STORAGE_KEY, JSON.stringify(this.project));
      this.notify({ reason: 'saved' }, { clone: false });
      setTimeout(() => this.saveCallbacks.forEach((fn) => fn('saved')), 400);
    } catch (e) {
      console.error('Save error:', e);
      this.saveCallbacks.forEach((fn) => fn('error', 'Opslag fout'));
    }
  }

  get data() {
    return this.project;
  }

  get activeSheet() {
    return (
      this.project.sheets.find((s) => s.id === this.project.activeSheetId) ||
      this.project.sheets[0]
    );
  }

  setActiveSheet(id) {
    if (!this.project.sheets.some((s) => s.id === id)) return;
    this.project.activeSheetId = id;
    this.notify({ reason: 'sheet' }, { clone: false });
  }

  updateProjectTitle(title) {
    this.project.projectTitle = title;
    this.notify({ reason: 'title' }, { clone: false, throttleMs: 50 });
  }

  updateStickyText(colIdx, slotIdx, text) {
    const sheet = this.activeSheet;
    const slot = sheet.columns[colIdx]?.slots?.[slotIdx];
    if (!slot) return;
    slot.text = text;
    this.notify({ reason: 'text', colIdx, slotIdx }, { clone: false, throttleMs: 50 });
  }

  setTransition(colIdx, val) {
    const col = this.activeSheet.columns[colIdx];
    if (!col) return;

    if (val === null) {
      col.hasTransition = false;
      col.transitionNext = '';
    } else {
      col.hasTransition = true;
      col.transitionNext = val;
    }

    this.notify({ reason: 'transition', colIdx }, { clone: false, throttleMs: 50 });
  }

  addSheet(name) {
    this.pushHistory();
    const newSheet = createSheet(name || `Proces ${this.project.sheets.length + 1}`);
    this.project.sheets.push(newSheet);
    this.project.activeSheetId = newSheet.id;

    this._ensureProcessIds(this.project);

    this.notify({ reason: 'sheets' }, { clone: false });
  }

  renameSheet(newName) {
    const sheet = this.activeSheet;
    if (!sheet || !newName || sheet.name === newName) return;
    this.pushHistory();
    sheet.name = newName;
    this.notify({ reason: 'sheets' }, { clone: false });
  }

  deleteSheet() {
    if (this.project.sheets.length <= 1) return false;

    const idx = this.project.sheets.findIndex((s) => s.id === this.project.activeSheetId);
    if (idx === -1) return false;

    this.pushHistory();
    this.project.sheets.splice(idx, 1);

    const newIdx = Math.max(0, idx - 1);
    this.project.activeSheetId = this.project.sheets[newIdx].id;

    this.notify({ reason: 'sheets' }, { clone: false });
    return true;
  }

  addColumn(afterIndex) {
    this.pushHistory();
    const sheet = this.activeSheet;
    const newCol = createColumn();

    if (afterIndex === -1) sheet.columns.push(newCol);
    else sheet.columns.splice(afterIndex + 1, 0, newCol);

    this._ensureProcessIds(this.project);

    // normalize merges because column count changed
    this._normalizeOutputMerges(sheet);
    this._normalizeSystemMerges(sheet);

    this.notify({ reason: 'columns' }, { clone: false });
  }

  deleteColumn(index) {
    const sheet = this.activeSheet;
    if (sheet.columns.length <= 1) return false;

    this.pushHistory();
    sheet.columns.splice(index, 1);

    // normalize merges because column count changed
    this._normalizeOutputMerges(sheet);
    this._normalizeSystemMerges(sheet);

    this.notify({ reason: 'columns' }, { clone: false });
    return true;
  }

  moveColumn(index, direction) {
    const sheet = this.activeSheet;
    const targetIndex = index + direction;
    if (targetIndex < 0 || targetIndex >= sheet.columns.length) return;

    this.pushHistory();
    [sheet.columns[index], sheet.columns[targetIndex]] = [
      sheet.columns[targetIndex],
      sheet.columns[index]
    ];

    // merges become ambiguous when reordering: safest is to clear them
    sheet.outputMerges = [];
    sheet.systemMerges = [];

    this.notify({ reason: 'columns' }, { clone: false });
  }

  setColVisibility(index, isVisible) {
    const sheet = this.activeSheet;
    if (!sheet.columns[index]) return;

    this.pushHistory();
    sheet.columns[index].isVisible = isVisible;

    // keep merges but rendering will skip if hidden inside range
    this.notify({ reason: 'columns' }, { clone: false });
  }

  saveStickyDetails() {
    this.pushHistory();
    this.notify({ reason: 'details' }, { clone: false });
  }

  toggleParallel(colIdx) {
    this.pushHistory();
    const col = this.activeSheet.columns[colIdx];
    if (!col) return;

    col.isParallel = !col.isParallel;
    if (col.isParallel) col.isVariant = false; // Reset variant

    this.notify({ reason: 'columns' }, { clone: false });
  }

  toggleVariant(colIdx) {
    this.pushHistory();
    const col = this.activeSheet.columns[colIdx];
    if (!col) return;

    col.isVariant = !col.isVariant;
    if (col.isVariant) col.isParallel = false; // Reset parallel

    this.notify({ reason: 'columns' }, { clone: false });
  }

  getGlobalCountersBeforeActive() {
    let inCount = 0;
    let outCount = 0;

    for (const sheet of this.project.sheets) {
      if (sheet.id === this.project.activeSheetId) break;
      sheet.columns.forEach((col) => {
        if (col.isVisible !== false) {
          if (col.slots[2].text?.trim()) inCount++;
          if (col.slots[4].text?.trim()) outCount++;
        }
      });
    }

    return { inStart: inCount, outStart: outCount };
  }

  getAllOutputs() {
    const map = {};
    let counter = 0;

    this.project.sheets.forEach((sheet) => {
      sheet.columns.forEach((col) => {
        if (col.slots[4].text?.trim()) {
          counter++;
          map[`OUT${counter}`] = col.slots[4].text;
        }
      });
    });

    return map;
  }
}

export const state = new StateManager();