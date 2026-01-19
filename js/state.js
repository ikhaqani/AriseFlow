import { createProjectState, createSheet, createColumn, createSticky, STORAGE_KEY, uid } from './config.js';

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

  _makeId(prefix = 'id') {
    try {
      if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
        return `${prefix}_${crypto.randomUUID()}`;
      }
    } catch {}
    return `${prefix}_${Date.now()}_${Math.random().toString(16).slice(2)}`;
  }

  _ensureProcessIds(project) {
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
    const sheets = project?.sheets || [];
    sheets.forEach((sheet) => {
      (sheet.columns || []).forEach((col) => {
        const outSlot = col?.slots?.[4];
        if (!outSlot) return;
        if (!outSlot.outputUid || String(outSlot.outputUid).trim() === '') outSlot.outputUid = this._makeId('out');
      });
    });
  }

  _sanitizeOutputBundle(b) {
    const x = b && typeof b === 'object' ? b : {};
    const id = String(x.id || uid()).trim() || uid();
    const name = String(x.name || '').trim();
    const outputUids = Array.isArray(x.outputUids)
      ? [...new Set(x.outputUids.map((v) => String(v || '').trim()).filter(Boolean))]
      : [];
    return { id, name, outputUids };
  }

  _ensureOutputBundles(project) {
    if (!project || typeof project !== 'object') return;
    if (!Array.isArray(project.outputBundles)) project.outputBundles = [];
    project.outputBundles = project.outputBundles.map((b) => this._sanitizeOutputBundle(b));
  }

  _migrateInputBundleFields(project) {
    // Migrates legacy single linkedBundleId -> linkedBundleIds[]
    (project?.sheets || []).forEach((sheet) => {
      (sheet?.columns || []).forEach((col) => {
        const inSlot = col?.slots?.[2];
        if (!inSlot) return;

        if (!Array.isArray(inSlot.linkedBundleIds)) {
          const single = String(inSlot.linkedBundleId || '').trim();
          inSlot.linkedBundleIds = single ? [single] : [];
        } else {
          inSlot.linkedBundleIds = inSlot.linkedBundleIds.map((x) => String(x || '').trim()).filter(Boolean);
        }

        // Keep legacy field in sync (first item) for backward compatibility
        inSlot.linkedBundleId = String(inSlot.linkedBundleIds[0] || '').trim() || null;
      });
    });
  }

  _buildLegacyOutIdToUidMap(project) {
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
    // Migrates legacy linkedSourceId (OUTn) and/or legacy single linkedSourceUid -> new linkedSourceUids[]
    const outIdToUid = this._buildLegacyOutIdToUidMap(project);

    (project?.sheets || []).forEach((sheet) => {
      (sheet?.columns || []).forEach((col) => {
        const inSlot = col?.slots?.[2];
        if (!inSlot) return;

        // Ensure array field exists
        if (!Array.isArray(inSlot.linkedSourceUids)) {
          const singleUid = String(inSlot.linkedSourceUid || '').trim();
          inSlot.linkedSourceUids = singleUid ? [singleUid] : [];
        }

        // If array already has content, keep it (and sync single field)
        if (inSlot.linkedSourceUids.length > 0) {
          inSlot.linkedSourceUid = String(inSlot.linkedSourceUids[0] || '').trim() || null;
          return;
        }

        // Legacy OUTn -> uid
        const legacy = String(inSlot.linkedSourceId || '').trim();
        if (/^OUT\d+$/.test(legacy)) {
          const uidVal = outIdToUid[legacy];
          if (uidVal) {
            inSlot.linkedSourceUids = [uidVal];
            inSlot.linkedSourceUid = uidVal;
          }
        }
      });
    });
  }

  _sanitizeSystemEntry(s) {
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
    this._normalizeMergeRanges(sheet, 'outputMerges', 4);
  }

  getOutputMergeForCol(colIdx) {
    const sheet = this.activeSheet;
    if (!sheet || !Array.isArray(sheet.outputMerges)) return null;
    return sheet.outputMerges.find((r) => r.slotIdx === 4 && colIdx >= r.startCol && colIdx <= r.endCol) || null;
  }

  setOutputMergeRangeForCol(colIdx, startCol, endCol) {
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
    this._normalizeMergeRanges(sheet, 'systemMerges', 1);
  }

  getSystemMergeForCol(colIdx) {
    const sheet = this.activeSheet;
    if (!sheet || !Array.isArray(sheet.systemMerges)) return null;
    return sheet.systemMerges.find((r) => r.slotIdx === 1 && colIdx >= r.startCol && colIdx <= r.endCol) || null;
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

  subscribe(listenerFn) {
    this.listeners.add(listenerFn);
    return () => this.listeners.delete(listenerFn);
  }

  onSave(fn) {
    this.saveCallbacks.add(fn);
    return () => this.saveCallbacks.delete(fn);
  }

  beginBatch(meta = { reason: 'batch' }) {
    this._suspendNotify += 1;
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
      this._ensureOutputUids(merged);
      this._ensureProcessIds(merged);
      this._migrateLinkedSourceIdToUid(merged);
      this._ensureOutputBundles(merged);
      this._migrateInputBundleFields(merged);
      return merged;
    }

    merged.sheets.forEach((sheet) => {
      // === GROUPS: Ensure array exists ===
      if (!Array.isArray(sheet.groups)) sheet.groups = [];
      
      // Valideer bestaande groups op nieuwe structuur
      sheet.groups = sheet.groups.map(g => {
         if (!Array.isArray(g.cols)) {
             // Migratie van oude Start/End naar Array
             if (Number.isFinite(g.startCol) && Number.isFinite(g.endCol)) {
                 const newCols = [];
                 for(let i=g.startCol; i<=g.endCol; i++) newCols.push(i);
                 g.cols = newCols;
             } else {
                 g.cols = [];
             }
         }
         return g;
      }).filter(g => g.cols.length > 0);
      
      // === NIEUW: Variant Groups initialiseren ===
      if (!Array.isArray(sheet.variantGroups)) sheet.variantGroups = [];
      // Schoon lege groups op
      sheet.variantGroups = sheet.variantGroups.filter(vg => 
          Number.isFinite(vg.parentColIdx) && 
          Array.isArray(vg.variants) && 
          vg.variants.length > 0
      );
      // ===========================================

      if (!Array.isArray(sheet.outputMerges)) sheet.outputMerges = [];
      if (!Array.isArray(sheet.systemMerges)) sheet.systemMerges = [];
      this._normalizeOutputMerges(sheet);
      this._normalizeSystemMerges(sheet);

      if (!Array.isArray(sheet.columns) || sheet.columns.length === 0) sheet.columns = [createColumn()];

      sheet.columns.forEach((col) => {
        if (typeof col.isVariant !== 'boolean') col.isVariant = false;
        if (typeof col.isParallel !== 'boolean') col.isParallel = !!col.isParallel;
        if (typeof col.isQuestion !== 'boolean') col.isQuestion = !!col.isQuestion;
        
        if (typeof col.isConditional !== 'boolean') col.isConditional = !!col.isConditional; 
        
        if (col.logic && typeof col.logic === 'object') {
           col.logic = {
             condition: String(col.logic.condition || ''),
             ifTrue: (col.logic.ifTrue !== null && Number.isFinite(Number(col.logic.ifTrue))) ? Number(col.logic.ifTrue) : null,
             ifFalse: (col.logic.ifFalse !== null && Number.isFinite(Number(col.logic.ifFalse))) ? Number(col.logic.ifFalse) : null
           };
        } else {
           col.logic = null;
        }
        
        if (typeof col.isGroup !== 'boolean') col.isGroup = !!col.isGroup;

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
          linkedSourceUids: Array.isArray(s.linkedSourceUids)
            ? s.linkedSourceUids.map((x) => String(x || '').trim()).filter(Boolean)
            : Array.isArray(clean.linkedSourceUids)
              ? clean.linkedSourceUids
              : [],
          linkedBundleId: s.linkedBundleId ?? clean.linkedBundleId,
          linkedBundleIds: Array.isArray(s.linkedBundleIds)
            ? s.linkedBundleIds.map((x) => String(x || '').trim()).filter(Boolean)
            : Array.isArray(clean.linkedBundleIds)
              ? clean.linkedBundleIds
              : [],

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

          if (slotIdx === 2) {
            if (!Array.isArray(mergedSlot.linkedSourceUids)) mergedSlot.linkedSourceUids = [];
            mergedSlot.linkedSourceUids = mergedSlot.linkedSourceUids.map((x) => String(x || '').trim()).filter(Boolean);

            if (mergedSlot.linkedSourceUids.length === 0) {
              const singleUid = String(mergedSlot.linkedSourceUid || '').trim();
              if (singleUid) mergedSlot.linkedSourceUids = [singleUid];
            }

            mergedSlot.linkedSourceUid = String(mergedSlot.linkedSourceUids[0] || '').trim() || null;

            if (!Array.isArray(mergedSlot.linkedBundleIds)) mergedSlot.linkedBundleIds = [];
            mergedSlot.linkedBundleIds = mergedSlot.linkedBundleIds.map((x) => String(x || '').trim()).filter(Boolean);

            if (mergedSlot.linkedBundleIds.length === 0) {
              const singleB = String(mergedSlot.linkedBundleId || '').trim();
              if (singleB) mergedSlot.linkedBundleIds = [singleB];
            }

            mergedSlot.linkedBundleId = String(mergedSlot.linkedBundleIds[0] || '').trim() || null;
          }

          return mergedSlot;
        });

        const outSlot = col?.slots?.[4];
        if (outSlot && (!outSlot.outputUid || String(outSlot.outputUid).trim() === '')) outSlot.outputUid = this._makeId('out');
      });
    });

    this._ensureOutputUids(merged);
    this._migrateLinkedSourceIdToUid(merged);
    this._ensureProcessIds(merged);
    this._ensureOutputBundles(merged);
    this._migrateInputBundleFields(merged);

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
    const found = this.project.sheets.find((s) => s.id === this.project.activeSheetId);
    return found || this.project.sheets[0];
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
    const slot = sheet?.columns?.[colIdx]?.slots?.[slotIdx];
    if (!slot) return;
    slot.text = text;
    this.notify({ reason: 'text', colIdx, slotIdx }, { clone: false, throttleMs: 50 });
  }

  setTransition(colIdx, val) {
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
    this.pushHistory();
    const newSheet = createSheet(name || `Proces ${this.project.sheets.length + 1}`);
    this.project.sheets.push(newSheet);
    this.project.activeSheetId = newSheet.id;
    this._ensureOutputUids(this.project);
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

  moveSheet(fromIndex, toIndex) {
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

    this._ensureOutputUids(this.project);
    this._ensureProcessIds(this.project);

    this._normalizeOutputMerges(sheet);
    this._normalizeSystemMerges(sheet);

    this.notify({ reason: 'columns' }, { clone: false });
  }

  deleteColumn(index) {
    const sheet = this.activeSheet;
    if (sheet.columns.length <= 1) return false;

    this.pushHistory();
    sheet.columns.splice(index, 1);

    if (Array.isArray(sheet.groups)) {
        sheet.groups = sheet.groups.map(g => {
            const newCols = g.cols.filter(c => c !== index)
                .map(c => c > index ? c - 1 : c);
            return { ...g, cols: newCols };
        }).filter(g => g.cols.length > 0);
    }
    
    // NIEUW: Variant Groups updaten bij verwijderen kolom
    if (Array.isArray(sheet.variantGroups)) {
        sheet.variantGroups = sheet.variantGroups.map(vg => {
             // Als parent verwijderd wordt, is de groep ongeldig -> straks filteren
             if (vg.parentColIdx === index) return null;
             
             // Update parent index
             const newParent = vg.parentColIdx > index ? vg.parentColIdx - 1 : vg.parentColIdx;
             
             // Update variants
             const newVariants = vg.variants.filter(v => v !== index).map(v => v > index ? v - 1 : v);
             
             return { ...vg, parentColIdx: newParent, variants: newVariants };
        }).filter(vg => vg && vg.variants.length > 0);
    }

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
    [sheet.columns[index], sheet.columns[targetIndex]] = [sheet.columns[targetIndex], sheet.columns[index]];

    if (Array.isArray(sheet.groups)) {
        sheet.groups.forEach(g => {
            g.cols = g.cols.map(c => {
                if (c === index) return targetIndex;
                if (c === targetIndex) return index;
                return c;
            });
        });
    }

    // NIEUW: Variant Groups updaten bij verplaatsen
    if (Array.isArray(sheet.variantGroups)) {
        sheet.variantGroups.forEach(vg => {
            // Update Parent
            if (vg.parentColIdx === index) vg.parentColIdx = targetIndex;
            else if (vg.parentColIdx === targetIndex) vg.parentColIdx = index;
            
            // Update Variants
            vg.variants = vg.variants.map(v => {
                if (v === index) return targetIndex;
                if (v === targetIndex) return index;
                return v;
            });
        });
    }

    sheet.outputMerges = [];
    sheet.systemMerges = [];

    this.notify({ reason: 'columns' }, { clone: false });
  }

  setColVisibility(index, isVisible) {
    const sheet = this.activeSheet;
    if (!sheet.columns[index]) return;

    this.pushHistory();
    sheet.columns[index].isVisible = !!isVisible;

    this.notify({ reason: 'columns' }, { clone: false });
  }

  saveStickyDetails() {
    this.pushHistory();
    this.notify({ reason: 'details' }, { clone: false });
  }

  toggleParallel(colIdx) {
    this.pushHistory();
    const col = this.activeSheet?.columns?.[colIdx];
    if (!col) return;

    col.isParallel = !col.isParallel;

    this.notify({ reason: 'columns' }, { clone: false });
  }

  toggleVariant(colIdx) {
    this.pushHistory();
    const col = this.activeSheet?.columns?.[colIdx];
    if (!col) return;

    col.isVariant = !col.isVariant;

    this.notify({ reason: 'columns' }, { clone: false });
  }

  // === NIEUWE FUNCTIES VOOR VARIANTS (ROUTES) ===
  getVariantGroupForCol(colIdx) {
    const sheet = this.activeSheet;
    if (!sheet || !Array.isArray(sheet.variantGroups)) return null;
    
    // Check of deze kolom een 'Main' (Parent) is
    const asParent = sheet.variantGroups.find(vg => vg.parentColIdx === colIdx);
    if (asParent) return { role: 'parent', group: asParent };

    // Check of deze kolom een 'Variant' (Child) is
    const asChild = sheet.variantGroups.find(vg => vg.variants.includes(colIdx));
    if (asChild) return { role: 'child', group: asChild };

    return null;
  }

  setVariantGroup(data) {
    this.pushHistory();
    const sheet = this.activeSheet;
    if (!sheet) return;

    if (!Array.isArray(sheet.variantGroups)) sheet.variantGroups = [];

    const parentIdx = Number(data.parentColIdx);
    const variantIndices = Array.isArray(data.variants) ? data.variants.map(Number) : [];

    const allInvolved = [parentIdx, ...variantIndices];
    
    sheet.variantGroups = sheet.variantGroups.filter(vg => {
        if (allInvolved.includes(vg.parentColIdx)) return false;
        vg.variants = vg.variants.filter(v => !allInvolved.includes(v));
        return vg.variants.length > 0;
    });

    if (variantIndices.length > 0) {
        sheet.variantGroups.push({
            id: this._makeId('var'),
            parentColIdx: parentIdx,
            variants: variantIndices
        });
        
        const allCols = sheet.columns;
        if(allCols[parentIdx]) allCols[parentIdx].isVariant = true; 
        variantIndices.forEach(idx => {
            if(allCols[idx]) allCols[idx].isVariant = true;
        });
    } else {
        const col = sheet.columns[parentIdx];
        if(col) col.isVariant = false;
    }

    this.notify({ reason: 'columns' }, { clone: false });
  }

  removeVariantGroup(groupId) {
      this.pushHistory();
      const sheet = this.activeSheet;
      if (!sheet || !Array.isArray(sheet.variantGroups)) return;

      const group = sheet.variantGroups.find(g => g.id === groupId);
      if(group) {
          const cols = sheet.columns;
          if(cols[group.parentColIdx]) cols[group.parentColIdx].isVariant = false;
          group.variants.forEach(idx => {
              if(cols[idx]) cols[idx].isVariant = false;
          });
      }

      sheet.variantGroups = sheet.variantGroups.filter(g => g.id !== groupId);
      this.notify({ reason: 'columns' }, { clone: false });
  }
  // ===============================================

  toggleQuestion(colIdx) {
    this.pushHistory();
    const col = this.activeSheet?.columns?.[colIdx];
    if (!col) return;

    col.isQuestion = !col.isQuestion;

    this.notify({ reason: 'columns' }, { clone: false });
  }

  toggleConditional(colIdx) {
    this.pushHistory();
    const col = this.activeSheet?.columns?.[colIdx];
    if (!col) return;

    col.isConditional = !col.isConditional;

    this.notify({ reason: 'columns' }, { clone: false });
  }

  setColumnLogic(colIdx, logicData) {
    this.pushHistory();
    const col = this.activeSheet?.columns?.[colIdx];
    if (!col) return;

    col.logic = {
      condition: String(logicData.condition || ''),
      ifTrue: logicData.ifTrue !== '' && logicData.ifTrue !== null ? Number(logicData.ifTrue) : null,
      ifFalse: logicData.ifFalse !== '' && logicData.ifFalse !== null ? Number(logicData.ifFalse) : null
    };

    col.isConditional = true;

    this.notify({ reason: 'columns' }, { clone: false });
  }

  toggleGroup(colIdx) {
    this.pushHistory();
    const col = this.activeSheet?.columns?.[colIdx];
    if (!col) return;

    col.isGroup = !col.isGroup;

    this.notify({ reason: 'columns' }, { clone: false });
  }

  getGroupForCol(colIdx) {
    const sheet = this.activeSheet;
    if (!sheet || !Array.isArray(sheet.groups)) return null;
    return sheet.groups.find(g => Array.isArray(g.cols) && g.cols.includes(colIdx)) || null;
  }

  setColumnGroup(groupData) {
    this.pushHistory();
    const sheet = this.activeSheet;
    if (!sheet) return;

    if (!Array.isArray(sheet.groups)) sheet.groups = [];

    const newCols = Array.isArray(groupData.cols) ? groupData.cols : [];
    
    sheet.groups.forEach(g => {
        g.cols = g.cols.filter(c => !newCols.includes(c));
    });
    sheet.groups = sheet.groups.filter(g => g.cols.length > 0);

    if (newCols.length > 0) {
      const existingIdx = groupData.id ? sheet.groups.findIndex(g => g.id === groupData.id) : -1;
      
      if (existingIdx !== -1) {
          sheet.groups[existingIdx].cols = newCols;
          sheet.groups[existingIdx].title = String(groupData.title || '').trim();
      } else {
          sheet.groups.push({
            id: this._makeId('grp'),
            cols: newCols,
            title: String(groupData.title || '').trim()
          });
      }
    }

    this.notify({ reason: 'groups' }, { clone: false });
  }

  removeGroup(groupId) {
    this.pushHistory();
    const sheet = this.activeSheet;
    if (!sheet || !Array.isArray(sheet.groups)) return;

    sheet.groups = sheet.groups.filter(g => g.id !== groupId);
    this.notify({ reason: 'groups' }, { clone: false });
  }

  getGlobalCountersBeforeActive() {
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

  getAllOutputsDetailed() {
    const list = [];
    let counter = 0;

    (this.project.sheets || []).forEach((sheet) => {
      (sheet.columns || []).forEach((col) => {
        if (col?.isVisible === false) return;
        const outSlot = col?.slots?.[4];
        if (!outSlot?.text?.trim()) return;

        if (!outSlot.outputUid || String(outSlot.outputUid).trim() === '') outSlot.outputUid = this._makeId('out');

        counter += 1;
        list.push({
          outId: `OUT${counter}`,
          uid: outSlot.outputUid,
          text: outSlot.text
        });
      });
    });

    return list;
  }

  resolveOutUidsToOutIds(uids) {
    const wanted = Array.isArray(uids) ? uids.map((x) => String(x || '').trim()).filter(Boolean) : [];
    if (wanted.length === 0) return [];

    const det = this.getAllOutputsDetailed();
    const byUid = new Map(det.map((x) => [String(x.uid), x.outId]));

    return wanted.map((u) => byUid.get(u) || '').filter(Boolean);
  }

  getOutputBundles() {
    return Array.isArray(this.project.outputBundles) ? this.project.outputBundles : [];
  }

  getOutputBundleById(bundleId) {
    const id = String(bundleId || '').trim();
    if (!id) return null;
    return this.getOutputBundles().find((b) => String(b.id) === id) || null;
  }

  addOutputBundle(name, outputUids = []) {
    this.pushHistory();
    this._ensureOutputBundles(this.project);

    const bundle = this._sanitizeOutputBundle({
      id: uid(),
      name: String(name || '').trim(),
      outputUids: Array.isArray(outputUids) ? outputUids : []
    });

    this.project.outputBundles.push(bundle);
    this.notify({ reason: 'details' }, { clone: false });
    return bundle;
  }

  updateOutputBundle(bundleId, { name, outputUids } = {}) {
    const b = this.getOutputBundleById(bundleId);
    if (!b) return false;

    this.pushHistory();

    if (name !== undefined) b.name = String(name || '').trim();
    if (outputUids !== undefined) {
      b.outputUids = [...new Set((Array.isArray(outputUids) ? outputUids : []).map((v) => String(v || '').trim()).filter(Boolean))];
    }

    this.notify({ reason: 'details' }, { clone: false });
    return true;
  }

  deleteOutputBundle(bundleId) {
    const id = String(bundleId || '').trim();
    if (!id) return false;

    this.pushHistory();
    this._ensureOutputBundles(this.project);

    const beforeLen = this.project.outputBundles.length;
    this.project.outputBundles = this.project.outputBundles.filter((b) => String(b.id) !== id);

    (this.project.sheets || []).forEach((sheet) => {
      (sheet.columns || []).forEach((col) => {
        const inSlot = col?.slots?.[2];
        if (!inSlot) return;
        if (!Array.isArray(inSlot.linkedBundleIds)) inSlot.linkedBundleIds = [];
        inSlot.linkedBundleIds = inSlot.linkedBundleIds.map((x) => String(x || '').trim()).filter(Boolean).filter((x) => x !== id);
        inSlot.linkedBundleId = String(inSlot.linkedBundleIds[0] || '').trim() || null;
      });
    });

    const changed = this.project.outputBundles.length !== beforeLen;
    if (changed) this.notify({ reason: 'details' }, { clone: false });
    return changed;
  }

  resolveBundleIdsToOutUids(bundleIds) {
    const ids = Array.isArray(bundleIds) ? bundleIds.map((x) => String(x || '').trim()).filter(Boolean) : [];
    if (!ids.length) return [];

    const all = [];
    ids.forEach((id) => {
      const b = this.getOutputBundleById(id);
      if (!b) return;
      (b.outputUids || []).forEach((u) => {
        const uu = String(u || '').trim();
        if (uu) all.push(uu);
      });
    });

    return [...new Set(all)];
  }

  getOutputBundleLabel(bundleId) {
    const b = this.getOutputBundleById(bundleId);
    return b ? String(b.name || '').trim() : '';
  }
}

export const state = new StateManager();