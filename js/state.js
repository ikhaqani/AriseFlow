import { createProjectState, createSheet, createColumn, createSticky, STORAGE_KEY, uid } from './config.js';

class StateManager {
  constructor() {
    /** Initializes the state manager with storage, listeners, and history. */
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

  _makeId(prefix = 'id') {
    /** Returns a reasonably unique id string for internal stable identifiers. */
    try {
      if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
        return `${prefix}_${crypto.randomUUID()}`;
      }
    } catch {
    }
    return `${prefix}_${Date.now()}_${Math.random().toString(16).slice(2)}`;
  }

  _ensureProcessIds(project) {
    /** Ensures each process slot has a stable id for referencing across features like gates. */
    const sheets = project?.sheets || [];
    sheets.forEach((sheet) => {
      (sheet.columns || []).forEach((col) => {
        const slot = col?.slots?.[3];
        if (!slot) return;
        if (!slot.id || String(slot.id).trim() === '') slot.id = this._makeId('proc');
      });
    });
  }

  _ensureOutputUids(project) {
    /** Ensures each output slot has a stable outputUid used for durable linking across reordering. */
    const sheets = project?.sheets || [];
    sheets.forEach((sheet) => {
      (sheet.columns || []).forEach((col) => {
        const outSlot = col?.slots?.[4];
        if (!outSlot) return;
        if (!outSlot.outputUid || String(outSlot.outputUid).trim() === '') outSlot.outputUid = this._makeId('out');
      });
    });
  }

  _buildLegacyOutIdToUidMap(project) {
    /** Builds a best-effort mapping OUTn -> outputUid using the current sheet/column order. */
    const map = {};
    let counter = 0;

    (project?.sheets || []).forEach((sheet) => {
      (sheet?.columns || []).forEach((col) => {
        if (col?.isVisible === false) return;

        const outSlot = col?.slots?.[4];
        if (!outSlot?.text?.trim()) return;

        if (!outSlot.outputUid || String(outSlot.outputUid).trim() === '') outSlot.outputUid = this._makeId('out');

        counter += 1;
        map[`OUT${counter}`] = outSlot.outputUid;
      });
    });

    return map;
  }

  _migrateLinkedSourceIdToUid(project) {
    /** Migrates legacy linkedSourceId (OUTn) into linkedSourceUid (stable) when possible. */
    const map = this._buildLegacyOutIdToUidMap(project);

    (project?.sheets || []).forEach((sheet) => {
      (sheet?.columns || []).forEach((col) => {
        const inSlot = col?.slots?.[2];
        if (!inSlot) return;

        const already = String(inSlot.linkedSourceUid || '').trim();
        if (already) return;

        const legacy = String(inSlot.linkedSourceId || '').trim();
        if (!/^OUT\d+$/.test(legacy)) return;

        const uidVal = map[legacy];
        if (!uidVal) return;

        inSlot.linkedSourceUid = uidVal;
      });
    });
  }

  _sanitizeSystemEntry(s) {
    /** Normalizes one system entry to a safe schema used by the System slot UI. */
    const x = s && typeof s === 'object' ? s : {};
    return {
      id: String(x.id || uid()),
      name: String(x.name || '').trim(),
      legacy: !!(x.legacy ?? x.isLegacy),
      future: String(x.future ?? x.futureSystem ?? '').trim(),
      qa: x.qa && typeof x.qa === 'object' ? { ...x.qa } : {},
      score: Number.isFinite(Number(x.score ?? x.calculatedScore)) ? Number(x.score ?? x.calculatedScore) : null
    };
  }

  _ensureMultiSystemShape(slot) {
    /** Ensures the System slot has a multi-system structure compatible with the current UI. */
    if (!slot || typeof slot !== 'object') return;

    if (!slot.systemData || typeof slot.systemData !== 'object') slot.systemData = {};
    const sd = slot.systemData;

    if (typeof sd.isMulti !== 'boolean') sd.isMulti = false;

    if (!Array.isArray(sd.systemsMeta?.systems)) {
      const legacyName = typeof sd.systemName === 'string' ? sd.systemName.trim() : '';
      const base = legacyName ? [this._sanitizeSystemEntry({ name: legacyName })] : [this._sanitizeSystemEntry({})];
      sd.systemsMeta = { multi: false, systems: base, activeSystemIdx: 0 };
    } else {
      const meta = sd.systemsMeta && typeof sd.systemsMeta === 'object' ? sd.systemsMeta : {};
      const systems = Array.isArray(meta.systems) ? meta.systems.map((x) => this._sanitizeSystemEntry(x)) : [];
      const safeSystems = systems.length ? systems : [this._sanitizeSystemEntry({})];

      const inferredMulti = safeSystems.length > 1;
      const multi = typeof meta.multi === 'boolean' ? meta.multi : inferredMulti;
      const activeSystemIdx = Number.isFinite(Number(meta.activeSystemIdx)) ? Number(meta.activeSystemIdx) : 0;

      sd.systemsMeta = {
        multi,
        systems: safeSystems,
        activeSystemIdx: Math.max(0, Math.min(safeSystems.length - 1, activeSystemIdx))
      };
    }

    if (!Number.isFinite(Number(sd.calculatedScore))) sd.calculatedScore = null;
    if (typeof sd.systemName !== 'string') sd.systemName = sd.systemsMeta?.systems?.[0]?.name || '';
  }

  _normalizeMergeRanges(sheet, key, slotIdx) {
    /** Normalizes merge range lists into non-overlapping contiguous ranges within sheet bounds. */
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
      .filter((r) => r.endCol - r.startCol >= 1);

    ranges.sort((a, b) => a.startCol - b.startCol);

    const merged = [];
    for (const r of ranges) {
      const last = merged[merged.length - 1];
      if (!last) merged.push(r);
      else if (r.startCol <= last.endCol) last.endCol = Math.max(last.endCol, r.endCol);
      else merged.push(r);
    }

    sheet[key] = merged;
  }

  _normalizeOutputMerges(sheet) {
    /** Normalizes output merges for the active sheet. */
    this._normalizeMergeRanges(sheet, 'outputMerges', 4);
  }

  getOutputMergeForCol(colIdx) {
    /** Returns the output merge range covering colIdx or null. */
    const sheet = this.activeSheet;
    if (!sheet || !Array.isArray(sheet.outputMerges)) return null;
    return sheet.outputMerges.find((r) => r.slotIdx === 4 && colIdx >= r.startCol && colIdx <= r.endCol) || null;
  }

  setOutputMergeRangeForCol(colIdx, startCol, endCol) {
    /** Sets or clears the output merge range covering colIdx. */
    const sheet = this.activeSheet;
    if (!sheet) return;

    this.pushHistory();

    if (!Array.isArray(sheet.outputMerges)) sheet.outputMerges = [];

    const s = Number(startCol);
    const e = Number(endCol);

    sheet.outputMerges = sheet.outputMerges.filter((r) => {
      if (!r || r.slotIdx !== 4) return true;
      const overlaps = !(e < r.startCol || s > r.endCol);
      const includesCol = colIdx >= r.startCol && colIdx <= r.endCol;
      return !(overlaps || includesCol);
    });

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

  _normalizeSystemMerges(sheet) {
    /** Normalizes system merges for the active sheet. */
    this._normalizeMergeRanges(sheet, 'systemMerges', 1);
  }

  getSystemMergeForCol(colIdx) {
    /** Returns the system merge range covering colIdx or null. */
    const sheet = this.activeSheet;
    if (!sheet || !Array.isArray(sheet.systemMerges)) return null;
    return sheet.systemMerges.find((r) => r.slotIdx === 1 && colIdx >= r.startCol && colIdx <= r.endCol) || null;
  }

  setSystemMergeRangeForCol(colIdx, startCol, endCol) {
    /** Sets or clears the system merge range covering colIdx. */
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

  subscribe(listenerFn) {
    /** Subscribes a listener to state updates and returns an unsubscribe function. */
    this.listeners.add(listenerFn);
    return () => this.listeners.delete(listenerFn);
  }

  onSave(fn) {
    /** Subscribes a callback to save lifecycle events and returns an unsubscribe function. */
    this.saveCallbacks.add(fn);
    return () => this.saveCallbacks.delete(fn);
  }

  beginBatch(meta = { reason: 'batch' }) {
    /** Starts a notification batch where updates are merged until endBatch is called. */
    this._suspendNotify += 1;
    this._pendingMeta = this._mergeMeta(this._pendingMeta, meta);
  }

  endBatch(meta = null) {
    /** Ends a notification batch and emits a merged update if this was the outermost batch. */
    this._suspendNotify = Math.max(0, this._suspendNotify - 1);
    if (meta) this._pendingMeta = this._mergeMeta(this._pendingMeta, meta);
    if (this._suspendNotify === 0) this.notify(this._pendingMeta || { reason: 'batch' });
  }

  _mergeMeta(a, b) {
    /** Merges two meta objects with later keys overriding earlier keys. */
    if (!a) return b ? { ...b } : null;
    if (!b) return a ? { ...a } : null;
    return { ...a, ...b };
  }

  notify(meta = { reason: 'full' }, { clone = false, throttleMs = 0 } = {}) {
    /** Notifies listeners with optional throttling and batching semantics. */
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
    /** Emits an update payload to all subscribed listeners. */
    const payload = clone
      ? typeof structuredClone === 'function'
        ? structuredClone(this.project)
        : JSON.parse(JSON.stringify(this.project))
      : this.project;

    this.listeners.forEach((fn) => fn(payload, meta));
  }

  pushHistory({ throttleMs = 0 } = {}) {
    /** Pushes a snapshot onto the undo history with optional throttling. */
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
    /** Restores the previous snapshot from history and returns whether it succeeded. */
    if (this.historyStack.length === 0) return false;
    this.redoStack.push(JSON.stringify(this.project));
    this.project = JSON.parse(this.historyStack.pop());
    this.notify({ reason: 'full' }, { clone: false });
    return true;
  }

  redo() {
    /** Reapplies the last undone snapshot and returns whether it succeeded. */
    if (this.redoStack.length === 0) return false;
    this.historyStack.push(JSON.stringify(this.project));
    this.project = JSON.parse(this.redoStack.pop());
    this.notify({ reason: 'full' }, { clone: false });
    return true;
  }

  loadFromStorage() {
    /** Loads project state from localStorage and sanitizes the result. */
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
    /** Normalizes project data to the latest schema while preserving user content. */
    const fresh = createProjectState();
    const merged = { ...fresh, ...data };

    if (!Array.isArray(merged.sheets) || merged.sheets.length === 0) {
      merged.sheets = fresh.sheets;
      this._ensureOutputUids(merged);
      this._ensureProcessIds(merged);
      return merged;
    }

    merged.sheets.forEach((sheet) => {
      if (!Array.isArray(sheet.outputMerges)) sheet.outputMerges = [];
      if (!Array.isArray(sheet.systemMerges)) sheet.systemMerges = [];
      this._normalizeOutputMerges(sheet);
      this._normalizeSystemMerges(sheet);

      if (!Array.isArray(sheet.columns) || sheet.columns.length === 0) sheet.columns = [createColumn()];

      sheet.columns.forEach((col) => {
        if (typeof col.isVariant !== 'boolean') col.isVariant = false;
        if (typeof col.isParallel !== 'boolean') col.isParallel = !!col.isParallel;
        if (typeof col.isVisible !== 'boolean') col.isVisible = col.isVisible !== false;

        if (!Array.isArray(col.slots) || col.slots.length !== 6) {
          col.slots = Array(6)
            .fill(null)
            .map(() => createSticky());
          return;
        }

        col.slots = col.slots.map((slot, slotIdx) => {
          const clean = createSticky();
          const s = slot && typeof slot === 'object' ? slot : {};

          const cleanGate = clean.gate && typeof clean.gate === 'object' ? clean.gate : {};
          const sGate = s.gate && typeof s.gate === 'object' ? s.gate : {};

          const mergedSlot = {
            ...clean,
            ...s,
            qa: { ...clean.qa, ...(s.qa || {}) },
            systemData: { ...clean.systemData, ...(s.systemData || {}) },
            linkedSourceId: s.linkedSourceId ?? clean.linkedSourceId,
            linkedSourceUid: s.linkedSourceUid ?? clean.linkedSourceUid,
            inputDefinitions: Array.isArray(s.inputDefinitions) ? s.inputDefinitions : clean.inputDefinitions,
            disruptions: Array.isArray(s.disruptions) ? s.disruptions : clean.disruptions,
            workExp: s.workExp ?? clean.workExp,
            workExpNote: s.workExpNote ?? clean.workExpNote,
            id: s.id ?? clean.id,
            isGate: s.isGate ?? clean.isGate,
            gate: { ...cleanGate, ...sGate }
          };

          if (slotIdx === 3 && (!mergedSlot.id || String(mergedSlot.id).trim() === '')) mergedSlot.id = this._makeId('proc');

          if (mergedSlot.gate) {
            if (!Array.isArray(mergedSlot.gate.checkProcessIds)) mergedSlot.gate.checkProcessIds = [];
            if (mergedSlot.gate.onFailTargetProcessId === undefined) mergedSlot.gate.onFailTargetProcessId = null;
            if (mergedSlot.gate.onPassTargetProcessId === undefined) mergedSlot.gate.onPassTargetProcessId = null;
            if (!mergedSlot.gate.rule) mergedSlot.gate.rule = 'ALL_OK';
            if (mergedSlot.gate.note === undefined) mergedSlot.gate.note = '';
          }

          if (slotIdx === 1) this._ensureMultiSystemShape(mergedSlot);

          return mergedSlot;
        });

        const outSlot = col?.slots?.[4];
        if (outSlot && (!outSlot.outputUid || String(outSlot.outputUid).trim() === '')) outSlot.outputUid = this._makeId('out');
      });
    });

    this._ensureOutputUids(merged);
    this._migrateLinkedSourceIdToUid(merged);
    this._ensureProcessIds(merged);

    return merged;
  }

  saveToStorage() {
    /** Persists the current project state into localStorage and emits save events. */
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
    /** Returns the current project state object. */
    return this.project;
  }

  get activeSheet() {
    /** Returns the currently active sheet or a safe fallback sheet. */
    const found = this.project.sheets.find((s) => s.id === this.project.activeSheetId);
    return found || this.project.sheets[0];
  }

  setActiveSheet(id) {
    /** Sets the active sheet by id and triggers a sheet update. */
    if (!this.project.sheets.some((s) => s.id === id)) return;
    this.project.activeSheetId = id;
    this.notify({ reason: 'sheet' }, { clone: false });
  }

  updateProjectTitle(title) {
    /** Updates the project title with throttled notifications. */
    this.project.projectTitle = title;
    this.notify({ reason: 'title' }, { clone: false, throttleMs: 50 });
  }

  updateStickyText(colIdx, slotIdx, text) {
    /** Updates a sticky text field in-place and emits a minimal text update. */
    const sheet = this.activeSheet;
    const slot = sheet?.columns?.[colIdx]?.slots?.[slotIdx];
    if (!slot) return;
    slot.text = text;
    this.notify({ reason: 'text', colIdx, slotIdx }, { clone: false, throttleMs: 50 });
  }

  setTransition(colIdx, val) {
    /** Sets or clears a column transition attribute used by the UI. */
    const col = this.activeSheet?.columns?.[colIdx];
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
    /** Adds a new sheet and activates it. */
    this.pushHistory();
    const newSheet = createSheet(name || `Proces ${this.project.sheets.length + 1}`);
    this.project.sheets.push(newSheet);
    this.project.activeSheetId = newSheet.id;
    this._ensureOutputUids(this.project);
    this._ensureProcessIds(this.project);
    this.notify({ reason: 'sheets' }, { clone: false });
  }

  renameSheet(newName) {
    /** Renames the active sheet if the name changed. */
    const sheet = this.activeSheet;
    if (!sheet || !newName || sheet.name === newName) return;
    this.pushHistory();
    sheet.name = newName;
    this.notify({ reason: 'sheets' }, { clone: false });
  }

  moveSheet(fromIndex, toIndex) {
    /** Moves a sheet within the sheets array to support custom process flow ordering. */
    const sheets = this.project.sheets || [];
    if (sheets.length <= 1) return;

    const a = Number(fromIndex);
    const b = Number(toIndex);
    if (!Number.isFinite(a) || !Number.isFinite(b)) return;
    if (a < 0 || b < 0 || a >= sheets.length || b >= sheets.length) return;
    if (a === b) return;

    this.pushHistory();
    const [moved] = sheets.splice(a, 1);
    sheets.splice(b, 0, moved);

    const active = this.project.activeSheetId;
    if (!sheets.some((s) => s.id === active)) this.project.activeSheetId = sheets[0]?.id || active;

    this.notify({ reason: 'sheets' }, { clone: false });
  }

  deleteSheet() {
    /** Deletes the active sheet if more than one exists and returns whether it succeeded. */
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
    /** Inserts a new column after the given index (or at end if -1). */
    this.pushHistory();
    const sheet = this.activeSheet;
    const newCol = createColumn();

    if (afterIndex === -1) sheet.columns.push(newCol);
    else sheet.columns.splice(afterIndex + 1, 0, newCol);

    this._ensureOutputUids(this.project);
    this._ensureProcessIds(this.project);

    this._normalizeOutputMerges(sheet);
    this._normalizeSystemMerges(sheet);

    this.notify({ reason: 'columns' }, { clone: false });
  }

  deleteColumn(index) {
    /** Deletes a column if at least one column remains and returns whether it succeeded. */
    const sheet = this.activeSheet;
    if (sheet.columns.length <= 1) return false;

    this.pushHistory();
    sheet.columns.splice(index, 1);

    this._normalizeOutputMerges(sheet);
    this._normalizeSystemMerges(sheet);

    this.notify({ reason: 'columns' }, { clone: false });
    return true;
  }

  moveColumn(index, direction) {
    /** Swaps a column with its neighbor and clears merges to avoid ambiguous ranges. */
    const sheet = this.activeSheet;
    const targetIndex = index + direction;
    if (targetIndex < 0 || targetIndex >= sheet.columns.length) return;

    this.pushHistory();
    [sheet.columns[index], sheet.columns[targetIndex]] = [sheet.columns[targetIndex], sheet.columns[index]];

    sheet.outputMerges = [];
    sheet.systemMerges = [];

    this.notify({ reason: 'columns' }, { clone: false });
  }

  setColVisibility(index, isVisible) {
    /** Sets a column visibility flag used by rendering and exports. */
    const sheet = this.activeSheet;
    if (!sheet.columns[index]) return;

    this.pushHistory();
    sheet.columns[index].isVisible = !!isVisible;

    this.notify({ reason: 'columns' }, { clone: false });
  }

  saveStickyDetails() {
    /** Creates a history point for non-text edits and triggers a details update. */
    this.pushHistory();
    this.notify({ reason: 'details' }, { clone: false });
  }

  toggleParallel(colIdx) {
    /** Toggles a column's parallel flag and disables variant when enabled. */
    this.pushHistory();
    const col = this.activeSheet?.columns?.[colIdx];
    if (!col) return;

    col.isParallel = !col.isParallel;
    if (col.isParallel) col.isVariant = false;

    this.notify({ reason: 'columns' }, { clone: false });
  }

  toggleVariant(colIdx) {
    /** Toggles a column's variant flag and disables parallel when enabled. */
    this.pushHistory();
    const col = this.activeSheet?.columns?.[colIdx];
    if (!col) return;

    col.isVariant = !col.isVariant;
    if (col.isVariant) col.isParallel = false;

    this.notify({ reason: 'columns' }, { clone: false });
  }

  getGlobalCountersBeforeActive() {
    /** Returns cumulative IN/OUT counts across sheets before the active sheet for stable numbering. */
    let inCount = 0;
    let outCount = 0;

    for (const sheet of this.project.sheets) {
      if (sheet.id === this.project.activeSheetId) break;

      (sheet.columns || []).forEach((col) => {
        if (col?.isVisible === false) return;

        const inSlot = col?.slots?.[2];
        const outSlot = col?.slots?.[4];

        if (inSlot?.text?.trim()) inCount += 1;
        if (outSlot?.text?.trim()) outCount += 1;
      });
    }

    return { inStart: inCount, outStart: outCount };
  }

  getAllOutputs() {
    /** Returns an OUTn->text map derived from the current sheet/column ordering. */
    const map = {};
    let counter = 0;

    (this.project.sheets || []).forEach((sheet) => {
      (sheet.columns || []).forEach((col) => {
        if (col?.isVisible === false) return;
        const outSlot = col?.slots?.[4];
        if (!outSlot?.text?.trim()) return;
        counter += 1;
        map[`OUT${counter}`] = outSlot.text;
      });
    });

    return map;
  }
}

export const state = new StateManager();