import { state } from './state.js';
import { IO_CRITERIA, PROCESS_STATUSES } from './config.js';
import { openEditModal, saveModalDetails, openLogicModal, openGroupModal, openVariantModal } from './modals.js'; // NIEUW: openVariantModal toegevoegd

const $ = (id) => document.getElementById(id);

let _openModalFn = null;
let _delegatedBound = false;
let _syncRaf = 0;
let _lastPointerDownTs = 0;

const MERGE_LS_PREFIX = 'ssipoc.mergeGroups.v2';

let _mergeGroups = [];

/** Builds a stable localStorage key scoped to project and sheet. */
function _mergeKey() {
  const pid = state?.project?.id || state?.project?.name || 'project';
  const sid = state?.activeSheet?.id || state?.activeSheet?.name || 'sheet';
  return `${MERGE_LS_PREFIX}:${pid}:${sid}`;
}

/** Loads persisted merge groups from localStorage for the active key. */
function loadMergeGroups() {
  try {
    const raw = localStorage.getItem(_mergeKey());
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

/** Persists merge groups to localStorage for the active key. */
function saveMergeGroups(groups) {
  try {
    localStorage.setItem(_mergeKey(), JSON.stringify(groups || []));
  } catch {}
}

/** Ensures merge groups are loaded for the current active sheet key. */
function ensureMergeGroupsLoaded() {
  const key = _mergeKey();
  if (_mergeGroups.__key !== key) {
    const g = loadMergeGroups();
    _mergeGroups = Array.isArray(g) ? g : [];
    _mergeGroups.__key = key;
  }
}

/** Returns true when the given zero-based column indices are a contiguous range. */
function isContiguousZeroBased(cols) {
  if (!cols || cols.length < 2) return false;
  const s = [...new Set(cols)].sort((a, b) => a - b);
  const min = s[0];
  const max = s[s.length - 1];
  return s.length === max - min + 1;
}

/** Normalizes a gate object to a safe persistence schema. */
function _sanitizeGate(gate) {
  if (!gate || typeof gate !== 'object') return null;
  const enabled = !!gate.enabled;
  const failTargetColIdx = Number.isFinite(Number(gate.failTargetColIdx))
    ? Number(gate.failTargetColIdx)
    : null;
  return { enabled, failTargetColIdx };
}

/** Normalizes systems meta to a safe schema and infers multi when 2+ systems exist. */
function _sanitizeSystemsMeta(meta) {
  if (!meta || typeof meta !== 'object') return null;

  const inferredMulti = Array.isArray(meta.systems) && meta.systems.length > 1;
  const multi = !!meta.multi || inferredMulti;

  const arr = Array.isArray(meta.systems) ? meta.systems : [];
  const systems = arr
    .map((s) => {
      if (!s || typeof s !== 'object') return null;

      const name = String(s.name ?? '').trim();
      const legacy = !!s.legacy;
      const future = String(s.future ?? '').trim();
      const qa = s.qa && typeof s.qa === 'object' ? { ...s.qa } : {};
      const score = s.score == null || !Number.isFinite(Number(s.score)) ? null : Number(s.score);

      return { name, legacy, future, qa, score };
    })
    .filter(Boolean);

  if (systems.length === 0) systems.push({ name: '', legacy: false, future: '', qa: {}, score: null });

  let activeSystemIdx = Number(meta.activeSystemIdx);
  if (!Number.isFinite(activeSystemIdx)) activeSystemIdx = 0;
  if (activeSystemIdx < 0) activeSystemIdx = 0;
  if (activeSystemIdx >= systems.length) activeSystemIdx = 0;

  return { multi, systems, activeSystemIdx };
}

/** Sanitizes one merge group to the active sheet bounds and schema. */
function sanitizeGroupForActiveSheet(g) {
  const sh = state.activeSheet;
  if (!sh) return null;

  const n = sh.columns?.length ?? 0;
  if (!n) return null;

  const slotIdx = Number(g?.slotIdx);
  if (!Number.isFinite(slotIdx)) return null;

  const cols = Array.isArray(g?.cols) ? g.cols.map((x) => Number(x)).filter(Number.isFinite) : [];
  const uniq = [...new Set(cols)].filter((c) => c >= 0 && c < n);
  if (uniq.length < 2) return null;
  if (!isContiguousZeroBased(uniq)) return null;

  let master = Number(g?.master);
  if (!Number.isFinite(master)) master = uniq[0];
  if (!uniq.includes(master)) master = uniq[0];

  const gate = slotIdx === 4 ? _sanitizeGate(g?.gate) : null;
  const systemsMeta = slotIdx === 1 ? _sanitizeSystemsMeta(g?.systemsMeta) : null;

  return {
    slotIdx,
    cols: uniq.sort((a, b) => a - b),
    master,
    label: String(g?.label || ''),
    gate,
    systemsMeta
  };
}

/** Returns all active merge groups sanitized for the current sheet. */
function getAllMergeGroupsSanitized() {
  ensureMergeGroupsLoaded();
  const out = [];
  for (const g of _mergeGroups) {
    const s = sanitizeGroupForActiveSheet(g);
    if (s) out.push(s);
  }
  return out;
}

/** Sets or clears a merge group for a slot index and persists it. */
function setMergeGroupForSlot(slotIdx, groupOrNull) {
  ensureMergeGroupsLoaded();
  const kept = _mergeGroups.filter((g) => Number(g?.slotIdx) !== Number(slotIdx));
  if (groupOrNull) kept.push(groupOrNull);
  _mergeGroups = kept;
  _mergeGroups.__key = _mergeKey();
  saveMergeGroups(_mergeGroups);
}

/** Returns the merge group that contains the given column for a slot if present. */
function getMergeGroup(colIdx, slotIdx) {
  const groups = getAllMergeGroupsSanitized();
  return groups.find((g) => g.slotIdx === slotIdx && g.cols.includes(colIdx)) || null;
}

/** Returns true when the cell is a non-master member of a merge group. */
function isMergedSlave(colIdx, slotIdx) {
  const g = getMergeGroup(colIdx, slotIdx);
  return !!g && colIdx !== g.master;
}

/** Returns a stable unique id string. */
function makeId(prefix = 'id') {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return `${prefix}_${crypto.randomUUID()}`;
  }
  return `${prefix}_${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

/** Builds a stable merge storage key for a given project and sheet. */
function mergeKeyFor(project, sheet) {
  const pid = project?.id || project?.name || project?.projectTitle || 'project';
  const sid = sheet?.id || sheet?.name || 'sheet';
  return `${MERGE_LS_PREFIX}:${pid}:${sid}`;
}

/** Loads raw merge groups for a given project and sheet key. */
function loadMergeGroupsRawFor(project, sheet) {
  try {
    const raw = localStorage.getItem(mergeKeyFor(project, sheet));
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

/** Sanitizes one merge group to the provided sheet bounds and schema. */
function sanitizeGroupForSheet(sheet, g) {
  const n = sheet?.columns?.length ?? 0;
  if (!n) return null;

  const slotIdx = Number(g?.slotIdx);
  if (![1, 4].includes(slotIdx)) return null;

  const cols = Array.isArray(g?.cols) ? g.cols.map((x) => Number(x)).filter(Number.isFinite) : [];
  const uniq = [...new Set(cols)].filter((c) => c >= 0 && c < n);
  if (uniq.length < 2) return null;
  if (!isContiguousZeroBased(uniq)) return null;

  let master = Number(g?.master);
  if (!Number.isFinite(master) || !uniq.includes(master)) master = uniq[0];

  const gate = slotIdx === 4 ? _sanitizeGate(g?.gate) : null;
  const systemsMeta = slotIdx === 1 ? _sanitizeSystemsMeta(g?.systemsMeta) : null;

  return { slotIdx, cols: uniq.sort((a, b) => a - b), master, gate, systemsMeta };
}

/** Returns sanitized merge groups for a provided project and sheet. */
function getMergeGroupsSanitizedForSheet(project, sheet) {
  const raw = loadMergeGroupsRawFor(project, sheet);
  return raw.map((g) => sanitizeGroupForSheet(sheet, g)).filter(Boolean);
}

/** Returns the merge group that contains a given cell in a provided sheet. */
function getMergeGroupForCellInSheet(groups, colIdx, slotIdx) {
  return (
    (groups || []).find(
      (x) => x.slotIdx === slotIdx && Array.isArray(x.cols) && x.cols.includes(colIdx)
    ) || null
  );
}

/** Returns true when the cell is a non-master member of a merge group in a provided sheet. */
function isMergedSlaveInSheet(groups, colIdx, slotIdx) {
  const g = getMergeGroupForCellInSheet(groups, colIdx, slotIdx);
  return !!g && colIdx !== g.master;
}

/** Returns true when a value looks like an OUT id (e.g., OUT12). */
function _looksLikeOutId(v) {
  const s = String(v || '').trim();
  return !!s && /^OUT\d+$/.test(s);
}

/** Normalizes multi-link input structure: returns array of linked sources (uids and/or OUT ids). */
function getLinkedSourcesFromInputSlot(slot) {
  if (!slot || typeof slot !== 'object') return [];

  const tokens = [];

  const arrUids = Array.isArray(slot.linkedSourceUids) ? slot.linkedSourceUids : [];
  const arrIds = Array.isArray(slot.linkedSourceIds) ? slot.linkedSourceIds : [];

  arrUids.forEach((x) => {
    const s = String(x || '').trim();
    if (s) tokens.push(s);
  });

  arrIds.forEach((x) => {
    const s = String(x || '').trim();
    if (s) tokens.push(s);
  });

  const singleUid = String(slot.linkedSourceUid || '').trim();
  const singleId = String(slot.linkedSourceId || '').trim();

  if (singleUid) tokens.push(singleUid);
  if (singleId) tokens.push(singleId);

  // de-dupe while preserving order
  const seen = new Set();
  const out = [];
  for (const t of tokens) {
    const key = String(t);
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(key);
  }

  return out;
}

/** Normalizes bundle-link input structure: returns array of linked bundle ids (strings). */
function getLinkedBundleIdsFromInputSlot(slot) {
  if (!slot || typeof slot !== 'object') return [];

  const tokens = [];

  const arr = Array.isArray(slot.linkedBundleIds) ? slot.linkedBundleIds : [];
  arr.forEach((x) => {
    const s = String(x || '').trim();
    if (s) tokens.push(s);
  });

  const single = String(slot.linkedBundleId || '').trim();
  if (single) tokens.push(single);

  const seen = new Set();
  const out = [];
  for (const t of tokens) {
    const k = String(t);
    if (seen.has(k)) continue;
    seen.add(k);
    out.push(k);
  }
  return out;
}

function _getAllOutputBundles(project) {
  const p = project || state?.project || state?.data;
  const arr = Array.isArray(p?.outputBundles) ? p.outputBundles : [];
  return arr.filter((b) => b && typeof b === 'object');
}

function _findBundleById(project, bundleId) {
  const id = String(bundleId || '').trim();
  if (!id) return null;
  return _getAllOutputBundles(project).find((b) => String(b.id || '').trim() === id) || null;
}

function _getBundleLabel(project, bundleId) {
  const b = _findBundleById(project, bundleId);
  const nm = String(b?.name || '').trim();
  return nm || String(bundleId || '').trim();
}

function _resolveBundleOutUids(project, bundleId) {
  const b = _findBundleById(project, bundleId);
  const arr = Array.isArray(b?.outputUids) ? b.outputUids : [];
  const seen = new Set();
  const out = [];
  arr.forEach((x) => {
    const s = String(x || '').trim();
    if (!s) return;
    if (seen.has(s)) return;
    seen.add(s);
    out.push(s);
  });
  return out;
}

function _joinSemiText(arr) {
  const a = Array.isArray(arr) ? arr : [];
  const cleaned = a.map((x) => String(x ?? '').trim()).filter((x) => x !== '');
  return cleaned.join('; ');
}

/** Resolves linked sources into aligned OUT-ids + display texts (all '; ' separated later). */
function resolveLinkedSourcesToOutAndText(tokens, outIdByUid, outTextByUid, outTextByOutId) {
  const ids = [];
  const texts = [];

  (Array.isArray(tokens) ? tokens : []).forEach((t) => {
    const s = String(t || '').trim();
    if (!s) return;

    // If the stored token is already an OUT id, use it directly
    if (_looksLikeOutId(s)) {
      const outId = s;
      const txt = String(outTextByOutId?.[outId] ?? '').trim() || outId;
      ids.push(outId);
      texts.push(txt);
      return;
    }

    // Otherwise treat it as outputUid
    const outId = String(outIdByUid?.[s] ?? '').trim() || s;
    const txt =
      String(outTextByUid?.[s] ?? '').trim() ||
      String(outTextByOutId?.[outId] ?? '').trim() ||
      outId;

    ids.push(outId);
    texts.push(txt);
  });

  return { ids, texts };
}

/** Builds global output id and text maps using current sheet order. */
function buildGlobalOutputMaps(project) {
  const outIdByUid = {};
  const outTextByUid = {};
  const outTextByOutId = {};

  let outCounter = 0;

  (project?.sheets || []).forEach((sheet) => {
    const groups = getMergeGroupsSanitizedForSheet(project, sheet);

    (sheet?.columns || []).forEach((col, colIdx) => {
      if (col?.isVisible === false) return;

      const outSlot = col?.slots?.[4];
      if (!outSlot?.text?.trim()) return;

      if (isMergedSlaveInSheet(groups, colIdx, 4)) return;

      if (!outSlot.outputUid || String(outSlot.outputUid).trim() === '') {
        outSlot.outputUid = makeId('out');
      }

      outCounter += 1;
      const outId = `OUT${outCounter}`;

      outIdByUid[outSlot.outputUid] = outId;
      outTextByUid[outSlot.outputUid] = outSlot.text;
      outTextByOutId[outId] = outSlot.text;
    });
  });

  return { outIdByUid, outTextByUid, outTextByOutId };
}

/** Computes IN/OUT counters that occur before the active sheet in current sheet order. */
function computeCountersBeforeActiveSheet(project, activeSheetId, outIdByUid) {
  let inCount = 0;
  let outCount = 0;

  for (const sheet of project?.sheets || []) {
    if (sheet.id === activeSheetId) break;

    const groups = getMergeGroupsSanitizedForSheet(project, sheet);

    (sheet.columns || []).forEach((col, colIdx) => {
      if (col?.isVisible === false) return;

      const inSlot = col?.slots?.[2];
      const outSlot = col?.slots?.[4];

      const tokens = getLinkedSourcesFromInputSlot(inSlot);
      const isLinked = tokens.some((t) =>
        _looksLikeOutId(t) ? true : !!(t && outIdByUid && outIdByUid[t])
      );

      if (!isLinked && inSlot?.text?.trim()) inCount += 1;

      if (outSlot?.text?.trim() && !isMergedSlaveInSheet(groups, colIdx, 4)) outCount += 1;
    });
  }

  return { inStart: inCount, outStart: outCount };
}

/** Reads and sanitizes systems meta stored in the System slot of a column. */
function _getSystemMetaFromSlot(colIdx) {
  const sh = state.activeSheet;
  const slot = sh?.columns?.[colIdx]?.slots?.[1];
  const meta = slot?.systemData?.systemsMeta;
  return _sanitizeSystemsMeta(meta) || null;
}

/** Applies sanitized systems meta to the System slot for all specified columns. */
function _applySystemMetaToColumns(cols, meta) {
  const sh = state.activeSheet;
  if (!sh) return;

  const clean = _sanitizeSystemsMeta(meta);
  if (!clean) return;

  cols.forEach((cIdx) => {
    const slot = sh.columns?.[cIdx]?.slots?.[1];
    if (!slot) return;

    if (typeof state.updateSystemMeta === 'function') {
      state.updateSystemMeta(cIdx, clean);
      return;
    }

    const sd = slot.systemData && typeof slot.systemData === 'object' ? slot.systemData : {};
    slot.systemData = { ...sd, systemsMeta: clean };
  });
}

/** Returns the process label for a column, falling back to a default label. */
function getProcessLabel(colIdx) {
  const sh = state.activeSheet;
  if (!sh) return `Kolom ${colIdx + 1}`;
  const t = sh.columns?.[colIdx]?.slots?.[3]?.text;
  const s = String(t ?? '').trim();
  return s || `Kolom ${colIdx + 1}`;
}

/** Returns the next visible column index after a given index or null. */
function getNextVisibleColIdx(fromIdx) {
  const sh = state.activeSheet;
  if (!sh) return null;
  const n = sh.columns?.length ?? 0;
  for (let i = fromIdx + 1; i < n; i++) {
    if (sh.columns[i]?.isVisible !== false) return i;
  }
  return null;
}

/** Returns the pass-route label for a merge group based on the next visible column. */
function getPassLabelForGroup(group) {
  const maxCol = Math.max(...(group?.cols ?? [group?.master ?? 0]));
  const nextIdx = getNextVisibleColIdx(maxCol);
  if (nextIdx == null) return 'Einde proces';
  return getProcessLabel(nextIdx);
}

/** Returns all process options used for routing dropdowns. */
function getAllProcessOptions() {
  const sh = state.activeSheet;
  if (!sh) return [];
  const n = sh.columns?.length ?? 0;
  const opts = [];
  for (let i = 0; i < n; i++) {
    const label = getProcessLabel(i);
    opts.push({ colIdx: i, label });
  }
  return opts;
}

/** Escapes HTML special characters for safe injection into HTML content. */
function escapeHTML(v) {
  return String(v ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** Escapes HTML attribute characters for safe injection into HTML attributes. */
function escapeAttr(v) {
  return String(v ?? '')
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

/** Builds the gate footer HTML block for Output validation display. */
function buildGateBlockHTML(gate, passLabel, failLabel) {
  if (!gate?.enabled) return '';
  return `
    <div class="auth-module" data-auth-gate="1" style="margin-top:auto; padding-top:6px;">
      <div class="auth-footer" style="display:flex; justify-content:space-between; font-size:11px; color:#546e7a;">
        <div class="logic-line fail" style="display:flex; align-items:center; gap:4px;">
            <span>❎</span>
            <span>Rework:</span>
            <span>${escapeHTML(failLabel || '—')}</span>
        </div>
        <div class="logic-line pass" style="display:flex; align-items:center; gap:4px;">
            <span style="font-size:12px;">✅</span>
            <span style="color:#388e3c; font-weight:700;">Pass:</span>
            <span>${escapeHTML(passLabel || '—')}</span>
        </div>
      </div>
    </div>
  `;
}

/** Applies or removes gate markup on a sticky based on the gate config. */
function applyGateToSticky(stickyEl, gate, passLabel, failLabel) {
  if (!stickyEl) return;

  stickyEl.classList.remove('is-auth-gate');
  const old = stickyEl.querySelector?.('[data-auth-gate="1"]');
  if (old) old.remove();

  if (!gate?.enabled) return;

  stickyEl.classList.add('is-auth-gate');
  const content = stickyEl.querySelector?.('.sticky-content');
  if (!content) return;

  content.insertAdjacentHTML('beforeend', buildGateBlockHTML(gate, passLabel, failLabel));
}

/** Removes the merge modal overlay from the DOM if present. */
function removeMergeModal() {
  const old = document.getElementById('mergeModalOverlay');
  if (old) old.remove();
}

/** Removes any legacy system-merge UI block that may still be present in the modal. */
function removeLegacySystemMergeUI() {
  const TARGET = 'SYSTEEM SAMENVOEGEN (AANEENGESLOTEN KOLOMMEN)';
  const all = Array.from(document.querySelectorAll('*'));
  const headerEl = all.find((el) => {
    const t = (el.textContent || '').trim();
    return t && t.toUpperCase() === TARGET;
  });
  if (!headerEl) return;

  const container =
    headerEl.closest('section') ||
    headerEl.closest('.card') ||
    headerEl.closest('.panel') ||
    headerEl.closest('.modal-section') ||
    headerEl.closest('.modal') ||
    headerEl.parentElement;

  if (container) container.remove();
  else headerEl.remove();
}

/** Opens the unified merge modal for System (1) or Output (4) with explicit merge enablement. */
function openMergeModal(clickedColIdx, slotIdx, openModalFn) {
  const sh = state.activeSheet;
  if (!sh) return;

  const n = sh.columns?.length ?? 0;
  if (!n) return;

  if (![1, 4].includes(slotIdx)) return;

  const slotName = slotIdx === 1 ? 'Systeem' : 'Output';

  const cur = getAllMergeGroupsSanitized().find((g) => g.slotIdx === slotIdx) || null;
  const curCols = cur?.cols?.length ? cur.cols : [];

  let selected = curCols.length ? [...curCols] : [clickedColIdx];
  let master = cur?.master ?? clickedColIdx;

  let mergeEnabled = curCols.length >= 2;
  if (!mergeEnabled) {
    selected = [clickedColIdx];
    master = clickedColIdx;
  }

  const masterCol = cur?.master ?? clickedColIdx;
  const curText = sh.columns?.[masterCol]?.slots?.[slotIdx]?.text ?? '';

  let gate = slotIdx === 4 ? _sanitizeGate(cur?.gate) || { enabled: false, failTargetColIdx: null } : null;

  let systemsMeta = null;

  if (slotIdx === 1) {
    const existing = _sanitizeSystemsMeta(cur?.systemsMeta) || _getSystemMetaFromSlot(masterCol);
    if (existing) {
      systemsMeta = existing;
    } else {
      // FIX: Gebruik de tekst van de post-it als die er is
      const initialName = String(curText || '').trim();
      systemsMeta = {
        multi: false,
        systems: [{ name: initialName, legacy: false, future: '', qa: {}, score: null }],
        activeSystemIdx: 0
      };
    }
  }

  const processOptions = getAllProcessOptions();

  removeMergeModal();

  const overlay = document.createElement('div');
  overlay.id = 'mergeModalOverlay';
  overlay.className = 'modal-overlay';
  overlay.style.display = 'grid';

  const modal = document.createElement('div');
  modal.className = 'modal';

  const systemExtraHTML =
    slotIdx === 1
      ? `
    <div style="border-top:1px solid #eee; margin: 16px 0 12px 0;"></div>

    <label style="display:flex; gap:12px; align-items:center; cursor:pointer; font-size:14px; margin: 6px 0 10px 0;">
      <input id="multiSystemsInStep" type="checkbox" />
      <span style="font-weight:600;">Ik werk in meerdere systemen binnen deze processtap</span>
    </label>

    <div id="systemsWrap" style="display:flex; flex-direction:column; gap:14px;"></div>

    <button id="addSystemBtn" class="std-btn primary" type="button" style="width:fit-content; padding:10px 14px;">+ Systeem toevoegen</button>

    <div style="margin-top:10px; font-size:14px; color:#cfd8dc;">
      Overall score (kolom): <strong id="overallSystemScore">—</strong>
    </div>
  `
      : '';

  const gateBlockHTML =
    slotIdx === 4
      ? `
    <div style="border-top:1px solid #eee; margin: 16px 0 12px 0;"></div>

    <div style="display:flex; justify-content:space-between; align-items:center;">
        <label class="modal-label" style="margin:0;">Procesvalidatie</label>
        <label style="display:flex; gap:8px; align-items:center; cursor:pointer; font-size:13px;">
            <input id="gateEnabled" type="checkbox" />
            <span style="font-weight:600;">Validatiestap toevoegen</span>
        </label>
    </div>

    <div id="gateDetails" style="margin-top:12px; background:#f9fcfd; padding:12px; border-radius:6px; border:1px solid #e0e6ed;">
      <label class="modal-label" style="font-size:11px; margin-top:0;">Routing bij 'Rework'</label>
      <select id="gateFailTarget" class="modal-input"></select>

      <div style="margin-top:8px; font-size:12px; color:#546e7a;">
        <strong>Routing bij 'Pass':</strong> <span id="gatePassLabel" style="color:#2e7d32;">—</span>
      </div>
    </div>
  `
      : '';

  modal.innerHTML = `
    <h3>Consolidatie ${slotName}</h3>

    <div class="sub-text">
      Selecteer de te combineren proceskolommen.
    </div>

    <label style="display:flex; gap:12px; align-items:center; cursor:pointer; font-size:14px; margin: 6px 0 10px 0;">
      <input id="mergeEnabled" type="checkbox" />
      <span style="font-weight:700;">Merge inschakelen</span>
    </label>

    <div id="mergeRangeWrap">
      <label class="modal-label">Bereik</label>
      <div id="mergeColsGrid" class="radio-group-container" style="gap:10px;"></div>
    </div>

    <div id="mergeStatusLine" style="margin-top:6px; font-size:12px; opacity:0.8;">Niet gemerged</div>

    ${
      slotIdx === 4
        ? `
    <label class="modal-label">${slotName} Definitie</label>
    <textarea id="mergeText" class="modal-input" rows="3" placeholder="Beschrijving van het geconsolideerde resultaat..."></textarea>
        `
        : ''
    }

    ${gateBlockHTML}
    ${systemExtraHTML}

    <div class="modal-btns">
      <button id="mergeOffBtn" class="std-btn danger-text" type="button">Opheffen</button>
      <button id="mergeCancelBtn" class="std-btn" type="button">Annuleren</button>
      <button id="mergeSaveBtn" class="std-btn primary" type="button">Toepassen</button>
    </div>
  `;

  overlay.appendChild(modal);
  document.body.appendChild(overlay);
  removeLegacySystemMergeUI();

  const grid = modal.querySelector('#mergeColsGrid');
  const txt = modal.querySelector('#mergeText');
  if (txt) txt.value = String(curText || '');

  const mergeEnabledEl = modal.querySelector('#mergeEnabled');
  const mergeRangeWrapEl = modal.querySelector('#mergeRangeWrap');
  const mergeStatusLineEl = modal.querySelector('#mergeStatusLine');

  const gateEnabledEl = modal.querySelector('#gateEnabled');
  const gateDetailsEl = modal.querySelector('#gateDetails');
  const gateFailTarget = modal.querySelector('#gateFailTarget');
  const gatePassLabelEl = modal.querySelector('#gatePassLabel');

  const multiSystemsInStepEl = modal.querySelector('#multiSystemsInStep');
  const systemsWrapEl = modal.querySelector('#systemsWrap');
  const addSystemBtnEl = modal.querySelector('#addSystemBtn');
  const overallSystemScoreEl = modal.querySelector('#overallSystemScore');

  function computePassLabel() {
    const fakeGroup = { cols: selected, master };
    return getPassLabelForGroup(fakeGroup);
  }

  function syncMergeUI() {
    if (mergeEnabledEl) mergeEnabledEl.checked = !!mergeEnabled;
    if (mergeRangeWrapEl) mergeRangeWrapEl.style.display = mergeEnabled ? 'block' : 'none';

    if (!mergeEnabled) {
      selected = [clickedColIdx];
      master = clickedColIdx;
      if (mergeStatusLineEl) mergeStatusLineEl.textContent = 'Niet gemerged';
      return;
    }

    const left = Math.min(...selected) + 1;
    const right = Math.max(...selected) + 1;
    if (mergeStatusLineEl) mergeStatusLineEl.textContent = `Gemerged: Kolom ${left} t/m ${right}`;
  }

  function syncGateEnabledUI() {
    if (!gate) return;
    if (gateEnabledEl) gateEnabledEl.checked = !!gate.enabled;
    if (gateDetailsEl) gateDetailsEl.style.display = gate.enabled ? 'block' : 'none';
  }

  function renderFailTargetOptions() {
    if (!gateFailTarget) return;

    gateFailTarget.innerHTML = '';
    const ph = document.createElement('option');
    ph.value = '';
    ph.textContent = 'Selecteer retourproces...';
    gateFailTarget.appendChild(ph);

    processOptions.forEach((o) => {
      const opt = document.createElement('option');
      opt.value = String(o.colIdx);
      opt.textContent = `${o.label}`;
      gateFailTarget.appendChild(opt);
    });

    if (gate?.failTargetColIdx != null && Number.isFinite(gate.failTargetColIdx)) {
      gateFailTarget.value = String(gate.failTargetColIdx);
    } else {
      gateFailTarget.value = '';
    }
  }

  function syncPassLabel() {
    if (gatePassLabelEl) gatePassLabelEl.textContent = computePassLabel();
  }

  if (slotIdx === 4) {
    gate.enabled = !!gate.enabled;
    syncGateEnabledUI();
    renderFailTargetOptions();
    syncPassLabel();

    gateEnabledEl?.addEventListener('change', () => {
      gate.enabled = !!gateEnabledEl.checked;
      syncGateEnabledUI();
    });

    gateFailTarget?.addEventListener('change', () => {
      const v = String(gateFailTarget.value || '').trim();
      gate.failTargetColIdx = v ? Number(v) : null;
    });
  }

  const SYSFIT_Q = [
    { id: 'q1', title: 'Hoe vaak dwingt het systeem je tot workarounds?', type: 'freq' },
    { id: 'q2', title: 'Hoe vaak remt het systeem je af?', type: 'freq' },
    { id: 'q3', title: 'Hoe vaak moet je gegevens dubbel registreren?', type: 'freq' },
    { id: 'q4', title: 'Hoe vaak laat het systeem ruimte voor fouten?', type: 'freq' },
    { id: 'q5', title: 'Wat is de impact bij systeemuitval?', type: 'impact' }
  ];

  const SYSFIT_OPTS = {
    freq: [
      { key: 'NEVER', label: '(Bijna) nooit', score: 1 },
      { key: 'SOMETIMES', label: 'Soms', score: 0.66 },
      { key: 'OFTEN', label: 'Vaak', score: 0.33 },
      { key: 'ALWAYS', label: '(Bijna) altijd', score: 0 }
    ],
    impact: [
      { key: 'SAFE', label: 'Veilig (Fallback)', score: 1 },
      { key: 'DELAY', label: 'Vertraging', score: 0.66 },
      { key: 'RISK', label: 'Groot Risico', score: 0.33 },
      { key: 'STOP', label: 'Volledige Stilstand', score: 0 }
    ]
  };

  function _computeSystemScore(sys) {
    const qa = sys?.qa && typeof sys.qa === 'object' ? sys.qa : {};
    let sum = 0;
    let nAns = 0;

    for (const q of SYSFIT_Q) {
      const ansKey = qa[q.id];
      if (!ansKey) continue;
      const opt = (SYSFIT_OPTS[q.type] || []).find((o) => o.key === ansKey);
      if (!opt) continue;
      sum += opt.score;
      nAns += 1;
    }

    if (nAns === 0) return null;
    return Math.round((sum / nAns) * 100);
  }

  function _computeOverallScore(meta) {
    const clean = _sanitizeSystemsMeta(meta);
    if (!clean) return null;

    const scores = (clean.systems || [])
      .map((s) => _computeSystemScore(s))
      .filter((x) => Number.isFinite(Number(x)));
    if (!scores.length) return null;

    const avg = scores.reduce((a, b) => a + b, 0) / scores.length;
    return Math.round(avg);
  }

  function _renderSystemCards() {
    if (slotIdx !== 1) return;
    if (!systemsWrapEl || !systemsMeta) return;

    systemsMeta =
      _sanitizeSystemsMeta(systemsMeta) || {
        multi: false,
        systems: [{ name: '', legacy: false, future: '', qa: {}, score: null }],
        activeSystemIdx: 0
      };

    if (multiSystemsInStepEl) multiSystemsInStepEl.checked = !!systemsMeta.multi;

    systemsWrapEl.innerHTML = '';
    const systems = systemsMeta.systems || [];

    systems.forEach((sys, idx) => {
      const card = document.createElement('div');
      card.style.background = 'rgba(0,0,0,0.08)';
      card.style.border = '1px solid rgba(255,255,255,0.08)';
      card.style.borderRadius = '12px';
      card.style.padding = '14px';

      const showDelete = !!systemsMeta.multi && systems.length > 1;
      const legacyChecked = !!sys.legacy;

      card.innerHTML = `
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
          <div style="font-weight:800; letter-spacing:0.04em; opacity:0.9;">SYSTEEM</div>
          ${
            showDelete
              ? `<button data-sys-del="${idx}" class="std-btn danger-text" type="button" style="padding:8px 12px; border:1px solid rgba(244,67,54,0.35);">Verwijderen</button>`
              : `<span></span>`
          }
        </div>

        <div style="font-size:12px; letter-spacing:0.06em; font-weight:800; opacity:0.75; margin-bottom:6px;">SYSTEEMNAAM</div>
        <input data-sys-name="${idx}" class="modal-input" style="margin:0; height:40px;" placeholder="Bijv. ARIA / EPIC / Radiotherapieweb / Monaco..." value="${escapeAttr(sys.name || '')}" />

        <div style="display:grid; grid-template-columns: 1fr 1fr; gap:12px; margin-top:12px; align-items:end;">
          <label style="display:flex; gap:10px; align-items:center; cursor:pointer; font-size:14px;">
            <input data-sys-legacy="${idx}" type="checkbox" ${legacyChecked ? 'checked' : ''} />
            <span>Legacy systeem</span>
          </label>

          <div>
            <div style="font-size:12px; letter-spacing:0.06em; font-weight:800; opacity:0.75; margin-bottom:6px;">TOEKOMSTIG SYSTEEM (VERWACHTING)</div>
            <input data-sys-future="${idx}" class="modal-input" style="margin:0; height:40px; ${
              legacyChecked ? '' : 'opacity:0.45; pointer-events:none;'
            }" placeholder="Bijv. ARIA / EPIC / nieuw portaal..." value="${escapeAttr(sys.future || '')}" />
          </div>
        </div>

        <div style="margin-top:16px; border-top:1px solid rgba(255,255,255,0.08); padding-top:14px;">
          <div style="font-weight:900; letter-spacing:0.06em; opacity:0.85; margin-bottom:6px;">SYSTEM FIT VRAGEN</div>
          <div style="font-size:13px; opacity:0.75; margin-bottom:12px;">Beantwoord per vraag hoe goed dit systeem jouw taak ondersteunt.</div>

          <div data-sys-qs="${idx}" style="display:flex; flex-direction:column; gap:14px;"></div>

          <div style="margin-top:12px; font-size:14px; opacity:0.85;">
            Score (dit systeem): <strong data-sys-score="${idx}">—</strong>
          </div>
        </div>
      `;

      systemsWrapEl.appendChild(card);

      const qWrap = card.querySelector(`[data-sys-qs="${idx}"]`);
      const qa = sys.qa && typeof sys.qa === 'object' ? sys.qa : {};

      SYSFIT_Q.forEach((q, qi) => {
        const row = document.createElement('div');
        row.innerHTML = `
          <div style="font-size:15px; font-weight:800; margin-bottom:10px;">
            ${qi + 1}. ${escapeHTML(q.title)}
          </div>
          <div style="display:grid; grid-template-columns: repeat(4, 1fr); gap:12px;">
            ${(SYSFIT_OPTS[q.type] || [])
              .map((o) => {
                const selectedBtn = qa[q.id] === o.key;
                return `
                  <button
                    type="button"
                    data-qid="${q.id}"
                    data-opt="${o.key}"
                    style="
                      height:54px;
                      border-radius:14px;
                      border:1px solid rgba(255,255,255,0.10);
                      background:${selectedBtn ? 'rgba(25,118,210,0.25)' : 'rgba(255,255,255,0.04)'};
                      color:inherit;
                      font-size:15px;
                      cursor:pointer;
                    "
                  >
                    ${escapeHTML(o.label)}
                  </button>
                `;
              })
              .join('')}
          </div>
        `;
        qWrap.appendChild(row);

        row.querySelectorAll('button[data-qid]').forEach((btn) => {
          btn.addEventListener('click', () => {
            const qid = btn.getAttribute('data-qid');
            const opt = btn.getAttribute('data-opt');
            if (!qid || !opt) return;

            systemsMeta.systems[idx].qa =
              systemsMeta.systems[idx].qa && typeof systemsMeta.systems[idx].qa === 'object'
                ? systemsMeta.systems[idx].qa
                : {};
            systemsMeta.systems[idx].qa[qid] = opt;

            const sScore = _computeSystemScore(systemsMeta.systems[idx]);
            systemsMeta.systems[idx].score = Number.isFinite(Number(sScore)) ? Number(sScore) : null;

            _renderSystemCards();
          });
        });
      });

      card.querySelectorAll('input[data-sys-name]').forEach((inp) => {
        inp.addEventListener('input', () => {
          systemsMeta.systems[idx].name = String(inp.value ?? '');
        });
      });

      card.querySelectorAll('input[data-sys-legacy]').forEach((chk) => {
        chk.addEventListener('change', () => {
          systemsMeta.systems[idx].legacy = !!chk.checked;
          if (!systemsMeta.systems[idx].legacy) systemsMeta.systems[idx].future = '';
          _renderSystemCards();
        });
      });

      card.querySelectorAll('input[data-sys-future]').forEach((inp) => {
        inp.addEventListener('input', () => {
          systemsMeta.systems[idx].future = String(inp.value ?? '');
        });
      });

      card.querySelectorAll('button[data-sys-del]').forEach((btn) => {
        btn.addEventListener('click', () => {
          const i = Number(btn.getAttribute('data-sys-del'));
          if (!Number.isFinite(i)) return;
          systemsMeta.systems.splice(i, 1);
          if (systemsMeta.systems.length === 0) {
            systemsMeta.systems.push({ name: '', legacy: false, future: '', qa: {}, score: null });
          }
          _renderSystemCards();
        });
      });

      const scoreEl = card.querySelector(`[data-sys-score="${idx}"]`);
      const sScore = _computeSystemScore(sys);
      if (scoreEl) scoreEl.textContent = Number.isFinite(Number(sScore)) ? `${Number(sScore)}%` : '—';
    });

    const overall = _computeOverallScore(systemsMeta);
    if (overallSystemScoreEl) {
      overallSystemScoreEl.textContent = Number.isFinite(Number(overall)) ? `${Number(overall)}%` : '—';
    }
  }

  if (slotIdx === 1) {
    systemsMeta =
      _sanitizeSystemsMeta(systemsMeta) || {
        multi: false,
        systems: [{ name: '', legacy: false, future: '', qa: {}, score: null }],
        activeSystemIdx: 0
      };

    _renderSystemCards();

    multiSystemsInStepEl?.addEventListener('change', () => {
      const multi = !!multiSystemsInStepEl.checked;
      systemsMeta.multi = multi;

      if (!multi) {
        systemsMeta.systems = [
          systemsMeta.systems?.[0] || { name: '', legacy: false, future: '', qa: {}, score: null }
        ];
      } else {
        if ((systemsMeta.systems?.length ?? 0) < 2) {
          systemsMeta.systems = [
            systemsMeta.systems?.[0] || { name: '', legacy: false, future: '', qa: {}, score: null },
            { name: '', legacy: false, future: '', qa: {}, score: null }
          ];
        }
      }

      _renderSystemCards();
    });

    addSystemBtnEl?.addEventListener('click', () => {
      systemsMeta.multi = true;
      if (multiSystemsInStepEl) multiSystemsInStepEl.checked = true;
      systemsMeta.systems = Array.isArray(systemsMeta.systems) ? systemsMeta.systems : [];
      systemsMeta.systems.push({ name: '', legacy: false, future: '', qa: {}, score: null });
      _renderSystemCards();
    });
  }

  for (let i = 0; i < n; i++) {
    const btn = document.createElement('div');
    btn.className = 'sys-opt';
    btn.textContent = `${i + 1}`;
    btn.title = `Kolom ${i + 1}`;
    btn.dataset.col = String(i);
    btn.style.minWidth = '32px';
    btn.style.textAlign = 'center';

    btn.classList.toggle('selected', selected.includes(i));

    btn.addEventListener('click', () => {
      if (!mergeEnabled) return;

      if (selected.includes(i)) selected = selected.filter((x) => x !== i);
      else selected = [...selected, i];

      selected = [...new Set(selected)].sort((a, b) => a - b);

      if (!selected.includes(master)) master = selected[0] ?? clickedColIdx;

      Array.from(grid.children).forEach((child) => {
        const c = Number(child.dataset.col);
        child.classList.toggle('selected', selected.includes(c));
      });

      if (slotIdx === 4) syncPassLabel();
      syncMergeUI();
    });

    grid.appendChild(btn);
  }

  mergeEnabledEl?.addEventListener('change', () => {
    mergeEnabled = !!mergeEnabledEl.checked;

    if (mergeEnabled && selected.length < 2) {
      const next = Math.min(clickedColIdx + 1, n - 1);
      selected = next !== clickedColIdx ? [clickedColIdx, next] : [clickedColIdx];
      selected = [...new Set(selected)].sort((a, b) => a - b);
      master = cur?.master ?? clickedColIdx;
    }

    syncMergeUI();

    Array.from(grid?.children || []).forEach((child) => {
      const c = Number(child.dataset.col);
      child.classList.toggle('selected', selected.includes(c));
    });

    if (slotIdx === 4) syncPassLabel();
  });

  syncMergeUI();

  function close() {
    removeMergeModal();
  }

  overlay.addEventListener('click', (e) => {
    if (e.target === overlay) close();
  });

  modal.querySelector('#mergeCancelBtn')?.addEventListener('click', close);

  modal.querySelector('#mergeOffBtn')?.addEventListener('click', () => {
    setMergeGroupForSlot(slotIdx, null);
    close();
    renderColumnsOnly(_openModalFn);
    scheduleSyncRowHeights();
  });

  modal.querySelector('#mergeSaveBtn')?.addEventListener('click', () => {
    const vEl = modal.querySelector('#mergeText');
    const v = vEl ? String(vEl.value ?? '') : String(curText || '');

    if (!mergeEnabled) {
      setMergeGroupForSlot(slotIdx, null);
      state.updateStickyText(clickedColIdx, slotIdx, v);

      if (slotIdx === 1) {
        const cleanMeta = _sanitizeSystemsMeta(systemsMeta);
        if (cleanMeta) {
          const overall = _computeOverallScore(cleanMeta);
          _applySystemMetaToColumns([clickedColIdx], cleanMeta);

          const slot = state.activeSheet?.columns?.[clickedColIdx]?.slots?.[1];
          if (slot) {
            const sd = slot.systemData && typeof slot.systemData === 'object' ? slot.systemData : {};
            slot.systemData = {
              ...sd,
              systemsMeta: cleanMeta,
              calculatedScore: Number.isFinite(Number(overall)) ? Number(overall) : sd.calculatedScore
            };
          }
        }
      }

      close();
      renderColumnsOnly(_openModalFn);
      scheduleSyncRowHeights();
      return;
    }

    if (selected.length < 2) {
      alert('Selecteer minimaal 2 aaneengesloten kolommen.');
      return;
    }
    if (!isContiguousZeroBased(selected)) {
      alert('De selectie bevat onderbrekingen. Selecteer een aaneengesloten reeks.');
      return;
    }

    const finalMaster = selected.includes(clickedColIdx) ? clickedColIdx : selected[0];

    const payload = {
      slotIdx,
      cols: selected,
      master: finalMaster,
      label: slotIdx === 1 ? 'Merged System' : 'Merged Output'
    };

    if (slotIdx === 4) payload.gate = _sanitizeGate(gate);
    if (slotIdx === 1) payload.systemsMeta = _sanitizeSystemsMeta(systemsMeta);

    setMergeGroupForSlot(slotIdx, payload);

    selected.forEach((cIdx) => {
      state.updateStickyText(cIdx, slotIdx, v);
    });

    if (slotIdx === 1) {
      const cleanMeta = _sanitizeSystemsMeta(systemsMeta);
      if (cleanMeta) {
        const overall = _computeOverallScore(cleanMeta);
        _applySystemMetaToColumns(selected, cleanMeta);

        selected.forEach((cIdx) => {
          const slot = state.activeSheet?.columns?.[cIdx]?.slots?.[1];
          if (!slot) return;
          const sd = slot.systemData && typeof slot.systemData === 'object' ? slot.systemData : {};
          slot.systemData = {
            ...sd,
            systemsMeta: cleanMeta,
            calculatedScore: Number.isFinite(Number(overall)) ? Number(overall) : sd.calculatedScore
          };
        });
      }
    }

    close();
    renderColumnsOnly(_openModalFn);
    scheduleSyncRowHeights();
  });
}

const TTF_FREQ_SCORES = { NEVER: 1, SOMETIMES: 0.66, OFTEN: 0.33, ALWAYS: 0 };
const TTF_IMPACT_SCORES = { SAFE: 1, DELAY: 0.66, RISK: 0.33, STOP: 0 };

/** Computes a 0-100 System Fit score for a single system entry based on stored answers. */
function computeTTFSystemScore(sys) {
  const qa = sys?.qa && typeof sys.qa === 'object' ? sys.qa : {};
  const keys = [
    { id: 'q1', map: TTF_FREQ_SCORES },
    { id: 'q2', map: TTF_FREQ_SCORES },
    { id: 'q3', map: TTF_FREQ_SCORES },
    { id: 'q4', map: TTF_FREQ_SCORES },
    { id: 'q5', map: TTF_IMPACT_SCORES }
  ];

  let sum = 0;
  let n = 0;

  for (const k of keys) {
    const ans = qa[k.id];
    if (!ans) continue;
    const v = k.map[ans];
    if (typeof v !== 'number' || !Number.isFinite(v)) continue;
    sum += v;
    n += 1;
  }

  if (n === 0) return null;
  return Math.round((sum / n) * 100);
}

/** Returns the ordered per-system TTF scores from systems meta, computing when missing. */
function computeTTFScoreListFromMeta(meta) {
  const clean = _sanitizeSystemsMeta(meta);
  if (!clean?.systems?.length) return [];

  return clean.systems.map((s) => {
    const stored = s?.score;
    if (Number.isFinite(Number(stored))) return Number(stored);
    const computed = computeTTFSystemScore(s);
    return Number.isFinite(Number(computed)) ? Number(computed) : null;
  });
}

/** Computes the weighted Input Quality score (0-100) from stored QA results. */
function calculateLSSScore(qa) {
  if (!qa) return null;

  let totalW = 0;
  let earnedW = 0;

  IO_CRITERIA.forEach((c) => {
    const val = qa[c.key]?.result;
    const isScored = ['GOOD', 'POOR', 'MODERATE', 'MINOR', 'FAIL', 'OK', 'NOT_OK'].includes(val);
    if (!isScored) return;

    totalW += c.weight;

    if (val === 'GOOD' || val === 'OK') earnedW += c.weight;
    else if (val === 'MINOR') earnedW += c.weight * 0.75;
    else if (val === 'MODERATE') earnedW += c.weight * 0.5;
    else earnedW += 0;
  });

  return totalW === 0 ? null : Math.round((earnedW / totalW) * 100);
}

/** Returns the emoji corresponding to a process status value. */
function getProcessEmoji(status) {
  if (!status) return '';
  const s = PROCESS_STATUSES?.find?.((x) => x.value === status);
  return s?.emoji || '';
}

/** Maps a zero-based index to a route letter for variant flows. */
function _toLetter(i0) {
  const n = Number(i0);
  if (!Number.isFinite(n) || n < 0) return 'A';
  const base = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  return base[n] || `R${n + 1}`;
}

/** Computes the per-column route letter map for variant columns. */
function computeVariantLetterMap(activeSheet) {
  // NIEUWE LOGICA: Haal letters op uit de Variant Groups als die er zijn
  // Fallback naar de oude logica voor backward compat of losse varianten
  
  const map = {};
  if (!activeSheet?.columns?.length) return map;

  // 1. Check groups
  if (Array.isArray(activeSheet.variantGroups)) {
      activeSheet.variantGroups.forEach(vg => {
          vg.variants.forEach((vIdx, i) => {
              map[vIdx] = _toLetter(i);
          });
      });
  }

  // 2. Vul aan voor varianten die NIET in een groep zitten (oude stijl)
  let inRun = false;
  let runIdx = 0;

  for (let i = 0; i < activeSheet.columns.length; i++) {
    // Als we al een letter hebben uit de groups, resetten we de 'run' teller niet, 
    // want die kolommen tellen als 'behandeld'.
    if (map[i]) {
        inRun = false; 
        runIdx = 0;
        continue;
    }

    const col = activeSheet.columns[i];
    if (col?.isVisible === false) continue;

    const isVar = !!col?.isVariant;

    if (isVar) {
      if (!inRun) {
        inRun = true;
        runIdx = 0;
      }
      map[i] = _toLetter(runIdx);
      runIdx += 1;
    } else {
      inRun = false;
      runIdx = 0;
    }
  }

  return map;
}

/** Returns UI metadata for a given work experience value. */
function getWorkExpMeta(workExp) {
  const v = String(workExp || '').toUpperCase();
  if (v === 'OBSTACLE') return { icon: '🛠️', short: 'Obstakel', context: 'Kost energie' };
  if (v === 'ROUTINE') return { icon: '🤖', short: 'Routine', context: 'Saai & Repeterend' };
  if (v === 'FLOW') return { icon: '🚀', short: 'Flow', context: 'Geeft energie' };
  return null;
}

/** Returns the icon corresponding to the Lean value classification. */
function getLeanIcon(val) {
  if (val === 'VA') return '💚';
  if (val === 'BNVA') return '⚖️';
  if (val === 'NVA') return '🗑️';
  return '';
}

/** Builds HTML for score badges including IQF (Input) and TTF (System). */
function buildScoreBadges({ slotIdx, slot }) {
  let html = '';

  const qaScore = calculateLSSScore(slot.qa);
  if (qaScore !== null && slotIdx === 2) {
    const badgeClass = qaScore >= 80 ? 'score-high' : qaScore >= 60 ? 'score-med' : 'score-low';
    html += `<div class="qa-score-badge ${badgeClass}">IQF: ${qaScore}%</div>`;
  }

  if (slotIdx === 1) {
    const meta = slot.systemData?.systemsMeta;
    const scoreList = computeTTFScoreListFromMeta(meta);

    if (scoreList.length) {
      const scores = scoreList.map((v) => (Number.isFinite(Number(v)) ? `${Number(v)}%` : '—'));
      const label = `TTF: ${scores.join('; ')}`;

      const overallStored = slot.systemData?.calculatedScore;
      const overallDerived = (() => {
        if (Number.isFinite(Number(overallStored))) return Number(overallStored);
        const valid = scoreList.filter((x) => Number.isFinite(Number(x)));
        if (!valid.length) return null;
        return Math.round(valid.reduce((a, b) => a + b, 0) / valid.length);
      })();

      const badgeClass =
        Number.isFinite(Number(overallDerived))
          ? Number(overallDerived) >= 80
            ? 'score-high'
            : Number(overallDerived) >= 60
              ? 'score-med'
              : 'score-low'
          : 'score-med';

      html += `<div class="qa-score-badge ${badgeClass}">${escapeHTML(label)}</div>`;
      return html;
    }

    if (slot.systemData?.calculatedScore != null) {
      const sysScore = slot.systemData.calculatedScore;
      const badgeClass = sysScore >= 80 ? 'score-high' : sysScore >= 60 ? 'score-med' : 'score-low';
      html += `<div class="qa-score-badge ${badgeClass}">TTF: ${sysScore}%</div>`;
    }
  }

  return html;
}

/** Builds the HTML for one sticky slot including badges and the editable content area. */
function buildSlotHTML({
  colIdx,
  slotIdx,
  slot,
  statusClass,
  typeIcon,
  myInputId,
  myOutputId,
  isLinked,
  scoreBadgeHTML,
  extraStickyClass = '',
  extraStickyStyle = ''
}) {
  const procEmoji = slotIdx === 3 && slot.processStatus ? getProcessEmoji(slot.processStatus) : '';
  const leanIcon = slotIdx === 3 && slot.processValue ? getLeanIcon(slot.processValue) : '';
  const workExpIcon = slotIdx === 3 ? (getWorkExpMeta(slot?.workExp)?.icon || '') : '';
  const linkIcon = isLinked ? '🔗' : '';

  const b1 = slotIdx === 3 && slot.type ? typeIcon : '';
  const b2 = slotIdx === 3 ? leanIcon : '';
  const b3 = slotIdx === 3 ? workExpIcon : (slotIdx === 2 && isLinked ? linkIcon : '');
  const b4 = slotIdx === 3 ? procEmoji : '';

  const editableAttr = isLinked ? 'contenteditable="false" data-linked="true"' : 'contenteditable="true"';

  return `
    <div class="sticky ${statusClass} ${extraStickyClass}" style="${escapeAttr(
    extraStickyStyle
  )}" data-col="${colIdx}" data-slot="${slotIdx}">
      <div class="sticky-grip"></div>

      <div class="badges-row">
        <div class="sticky-badge">${escapeHTML(b1)}</div>
        <div class="sticky-badge">${escapeHTML(b2)}</div>
        <div class="sticky-badge">${escapeHTML(b3)}</div>
        <div class="sticky-badge emoji-only">${escapeHTML(b4)}</div>
      </div>

      ${slotIdx === 2 && myInputId ? `<div class="id-tag">${myInputId}</div>` : ''}
      ${slotIdx === 4 && myOutputId ? `<div class="id-tag">${myOutputId}</div>` : ''}

      ${scoreBadgeHTML}

      <div class="sticky-content">
        <div class="text" ${editableAttr} spellcheck="false"></div>
      </div>
    </div>
  `;
}

/** Computes an element offset relative to an ancestor element. */
function getOffsetWithin(el, ancestor) {
  let x = 0;
  let y = 0;
  let cur = el;
  while (cur && cur !== ancestor) {
    x += cur.offsetLeft || 0;
    y += cur.offsetTop || 0;
    cur = cur.offsetParent;
  }
  return { x, y };
}

/** Schedules a single row-height sync and merged overlay render on the next frame. */
function scheduleSyncRowHeights() {
  if (_syncRaf) cancelAnimationFrame(_syncRaf);
  _syncRaf = requestAnimationFrame(() => {
    _syncRaf = 0;
    syncRowHeightsNow();
    renderMergedOverlays(_openModalFn);
    renderGroupOverlays(); // FIX: Nu ook groepen renderen
  });
}

/** Synchronizes row heights across columns using the tallest sticky per row. */
function syncRowHeightsNow() {
  const rowHeadersEl = $('row-headers');
  const rowHeaders = rowHeadersEl?.children;
  if (!rowHeaders || !rowHeaders.length) return;

  const colsContainer = $('cols');
  const cols = colsContainer?.querySelectorAll?.('.col');
  if (!cols || !cols.length) return;

  // 1) Reset heights so content can naturally expand before measuring.
  for (let r = 0; r < 6; r++) {
    if (rowHeaders[r]) rowHeaders[r].style.height = 'auto';
    cols.forEach((col) => {
      const slotNodes = col.querySelectorAll('.slots .slot');
      if (!slotNodes[r]) return;
      slotNodes[r].style.height = 'auto';
      const sticky = slotNodes[r].firstElementChild;
      if (sticky) sticky.style.height = 'auto';
    });
  }

  // 2) Measure tallest sticky per row after reset.
  const MIN_ROW_HEIGHT = 160;
  const heights = Array(6).fill(MIN_ROW_HEIGHT);

  cols.forEach((col) => {
    const slotNodes = col.querySelectorAll('.slots .slot');
    for (let r = 0; r < 6; r++) {
      const slot = slotNodes[r];
      if (!slot) continue;
      const sticky = slot.firstElementChild;
      if (!sticky) continue;

      // Use scrollHeight as source-of-truth for content growth.
      // FIX: Increase buffer from +2 to +32 to prevent any text cutoff or overlap.
      const h = Math.ceil(
        Math.max(sticky.scrollHeight || 0, sticky.getBoundingClientRect?.().height || 0)
      ) + 32;
      if (h > heights[r]) heights[r] = h;
    }
  });

  // 2b) Make ALL rows the same height for consistent layout across the entire board.
  //     If one row grows, every row grows to match the largest row.
  const globalMax = Math.max(...heights);
  for (let r = 0; r < 6; r++) heights[r] = globalMax;

  // 3) Apply unified heights across all columns per row.
  for (let r = 0; r < 6; r++) {
    const hStr = `${heights[r]}px`;
    if (rowHeaders[r]) rowHeaders[r].style.height = hStr;

    cols.forEach((col) => {
      const slotNodes = col.querySelectorAll('.slots .slot');
      if (slotNodes[r]) slotNodes[r].style.height = hStr;
    });
  }

  // Keep connector vertical alignment correct.
  const gapSize = 20;
  const processOffset = heights[0] + heights[1] + heights[2] + 3 * gapSize;

  colsContainer.querySelectorAll('.col-connector').forEach((c) => {
    if (!c.classList.contains('parallel-connector') && !c.classList.contains('combo-connector')) {
      c.style.paddingTop = `${processOffset}px`;
    }
  });
}

/** Ensures the SSIPOC row headers are present on the left side of the board. */
function ensureRowHeaders() {
  const rowHeaderContainer = $('row-headers');
  if (!rowHeaderContainer || rowHeaderContainer.children.length > 0) return;

  ['Leverancier', 'Systeem', 'Input', 'Proces', 'Output', 'Klant'].forEach((label) => {
    const div = document.createElement('div');
    div.className = 'row-header';
    div.innerHTML = `<span>${label}</span>`;
    rowHeaderContainer.appendChild(div);
  });
}

/** Renders the sheet selector based on the active project sheets. */
function renderSheetSelect() {
  const select = $('sheetSelect');
  if (!select) return;

  const activeId = state.project.activeSheetId;
  select.innerHTML = '';

  state.project.sheets.forEach((s) => {
    const opt = document.createElement('option');
    opt.value = s.id;
    opt.textContent = s.name;
    opt.selected = s.id === activeId;
    select.appendChild(opt);
  });
}

/** Renders the current sheet title in the board header. */
function renderHeader(activeSheet) {
  const headDisp = $('board-header-display');
  if (headDisp) headDisp.textContent = activeSheet.name;
}

/** Attaches click and double-click interactions to a sticky and its text element. */
function attachStickyInteractions({ stickyEl, textEl, colIdx, slotIdx, openModalFn }) {
  const onDblClick = (e) => {
    if (![1, 2, 3, 4].includes(slotIdx)) return;
    e.preventDefault();
    e.stopPropagation();

    const sel = window.getSelection?.();
    if (sel) sel.removeAllRanges();

    if (slotIdx === 4 || slotIdx === 1) openMergeModal(colIdx, slotIdx, openModalFn);
    else {
      openModalFn?.(colIdx, slotIdx);
      requestAnimationFrame(() => removeLegacySystemMergeUI());
    }
  };

  const focusText = (e) => {
    if (e.detail && e.detail > 1) return;

    if (
      e.target.closest(
        '.sticky-grip, .qa-score-badge, .id-tag, .badges-row, .workexp-badge, .btn-col-action, .col-actions'
      )
    ) {
      return;
    }

    requestAnimationFrame(() => {
      textEl.focus();

      const range = document.createRange();
      const sel = window.getSelection();
      range.selectNodeContents(textEl);
      range.collapse(false);
      sel.removeAllRanges();
      sel.addRange(range);
    });
  };

  stickyEl.addEventListener('dblclick', onDblClick);
  textEl.addEventListener('dblclick', onDblClick);
  stickyEl.addEventListener('click', focusText);
}

/** Renders a connector between columns for parallel or variant flow visualization. */
function renderConnector({ frag, activeSheet, colIdx, variantLetterMap }) {
  if (colIdx >= activeSheet.columns.length - 1) return;

  let nextVisibleIdx = null;
  for (let i = colIdx + 1; i < activeSheet.columns.length; i++) {
    if (activeSheet.columns[i].isVisible !== false) {
      nextVisibleIdx = i;
      break;
    }
  }
  if (nextVisibleIdx == null) return;

  const nextCol = activeSheet.columns[nextVisibleIdx];

  const hasParallel = !!nextCol.isParallel;
  const hasVariant = !!nextCol.isVariant;
  const hasQuestion = !!nextCol.isQuestion;
  const hasConditional = !!nextCol.isConditional;
  const hasGroup = !!nextCol.isGroup;

  // NIEUWE LOGICA: Haal opgeslagen conditie-data op voor tooltip
  const conditionData = nextCol.logic || null;
  const conditionTooltip = conditionData && conditionData.condition
    ? `Logica: ${escapeAttr(conditionData.condition)}`
    : 'Conditionele stap';

  let badgesHTML = '';

  if (hasVariant) {
    const letter = variantLetterMap?.[nextVisibleIdx] || 'A';
    badgesHTML += `<div class="variant-badge">🔀${letter}</div>`;
  }
  if (hasParallel) {
    badgesHTML += `<div class="parallel-badge">||</div>`;
  }
  
  // LOGIC CHANGE: Conditional (Lightning) met tooltip
  if (hasConditional) {
    badgesHTML += `<div class="conditional-badge" title="${conditionTooltip}">⚡</div>`;
  }
  
  if (hasQuestion) {
    badgesHTML += `<div class="question-badge">❓</div>`;
  }

  const count = (hasParallel ? 1 : 0) + (hasVariant ? 1 : 0) + (hasQuestion ? 1 : 0) + (hasConditional ? 1 : 0) + (hasGroup ? 1 : 0);

  const connEl = document.createElement('div');

  // If one or more special types are active, use the combo/stack logic
  if (count > 0) {
    connEl.className = 'col-connector combo-connector';
    connEl.innerHTML = `
      <div class="combo-badge-stack">
        ${badgesHTML}
      </div>
    `;
  } else {
    // Standard arrow
    connEl.className = 'col-connector';
    connEl.innerHTML = `<div class="connector-active"></div>`;
  }

  frag.appendChild(connEl);
}

/** Renders the process status counters in the UI header. */
function renderStats(stats) {
  const happyEl = $('countHappy');
  const neutralEl = $('countNeutral');
  const sadEl = $('countSad');

  if (happyEl) happyEl.textContent = stats.happy;
  if (neutralEl) neutralEl.textContent = stats.neutral;
  if (sadEl) sadEl.textContent = stats.sad;
}

/** Removes any previously rendered merged overlays from the columns container. */
function clearMergedOverlays(colsContainer) {
  colsContainer?.querySelectorAll?.('.merged-overlay')?.forEach((n) => n.remove());
}

/** Builds system summary HTML using inline Legacy pills for merged System overlays. */
function _formatSystemsSummaryFromMeta(meta) {
  const clean = _sanitizeSystemsMeta(meta);
  if (!clean) return '';

  const systems = clean.systems || [];

  const lineHTML = (s) => {
    const nm = String(s?.name || '').trim() || '—';
    const legacy = !!s?.legacy;
    return `<div class="sys-line"><span class="sys-name">${escapeHTML(
      nm
    )}</span>${legacy ? `<span class="legacy-tag" aria-label="Legacy">Legacy</span>` : ''}</div>`;
  };

  if (!clean.multi) {
    const s = systems[0] || { name: '', legacy: false };
    const nm = String(s.name || '').trim();
    if (!nm) return '';
    return `<div class="sys-summary">${lineHTML(s)}</div>`;
  }

  return `<div class="sys-summary">${systems.map(lineHTML).join('')}</div>`;
}

/** Renders merged overlays for System and Output across active merge group ranges. */
function renderMergedOverlays(openModalFn) {
  const colsContainer = $('cols');
  if (!colsContainer) return;

  if (getComputedStyle(colsContainer).position === 'static') colsContainer.style.position = 'relative';

  // REMOVED: clearMergedOverlays(colsContainer);

  const activeSheet = state.activeSheet;
  if (!activeSheet) return;

  const groups = getAllMergeGroupsSanitized();
  
  // FIX: Tracken welke overlays we deze frame verwerken, om te voorkomen dat we gefocuste elementen weggooien
  const processedKeys = new Set();
  
  groups.forEach((g) => {
    const visibleCols = g.cols.filter((cIdx) => activeSheet.columns[cIdx]?.isVisible !== false);
    if (visibleCols.length < 2) return;

    const firstCol = visibleCols[0];
    const lastCol = visibleCols[visibleCols.length - 1];
    const masterCol = g.master;

    const firstColEl = colsContainer.querySelector(`.col[data-idx="${firstCol}"]`);
    const lastColEl = colsContainer.querySelector(`.col[data-idx="${lastCol}"]`);
    const masterColEl = colsContainer.querySelector(`.col[data-idx="${masterCol}"]`);
    if (!firstColEl || !lastColEl || !masterColEl) return;

    const firstSlot = firstColEl.querySelectorAll('.slots .slot')[g.slotIdx];
    const lastSlot = lastColEl.querySelectorAll('.slots .slot')[g.slotIdx];
    const masterSlot = masterColEl.querySelectorAll('.slots .slot')[g.slotIdx];
    if (!firstSlot || !lastSlot || !masterSlot) return;

    const masterSticky = masterSlot.querySelector(`.sticky[data-col="${masterCol}"][data-slot="${g.slotIdx}"]`);
    if (!masterSticky) return;

    const p1 = getOffsetWithin(firstSlot, colsContainer);
    const p2 = getOffsetWithin(lastSlot, colsContainer);

    const left = p1.x;
    const top = p1.y;
    const width = p2.x + lastSlot.offsetWidth - p1.x;
    const height = firstSlot.offsetHeight;

    // FIX: Unieke key per merge group om element te hergebruiken
    const mergeKey = `g-${g.slotIdx}-${g.master}`;
    processedKeys.add(mergeKey);

    let overlay = colsContainer.querySelector(`.merged-overlay[data-merge-key="${mergeKey}"]`);
    const isNew = !overlay;

    if (isNew) {
        overlay = document.createElement('div');
        overlay.className = 'merged-overlay';
        overlay.dataset.mergeKey = mergeKey;
        overlay.style.position = 'absolute';
        overlay.style.zIndex = '500';
        overlay.style.pointerEvents = 'auto';

        const cloned = masterSticky.cloneNode(true);
        if (g.slotIdx === 1) cloned.classList.add('has-sys-summary');

        cloned.classList.remove('merged-source');
        cloned.style.visibility = 'visible';
        cloned.style.pointerEvents = 'auto';
        cloned.style.width = '100%';
        cloned.style.height = '100%';
        cloned.classList.add('merged-sticky');
        
        const txt = cloned.querySelector('.text');
        if (txt) {
            txt.removeAttribute('data-linked');
            txt.addEventListener(
                'input',
                () => {
                  if (g.slotIdx === 1 && g.systemsMeta) return;
                  state.updateStickyText(masterCol, g.slotIdx, txt.textContent);
                  scheduleSyncRowHeights();
                },
                { passive: true }
            );
        }

        const stickyEl = cloned;
        const textEl = cloned.querySelector('.text');
        attachStickyInteractions({ stickyEl, textEl, colIdx: masterCol, slotIdx: g.slotIdx, openModalFn });
        
        overlay.appendChild(cloned);
        colsContainer.appendChild(overlay);
    }

    // UPDATE GEOMETRY (Always)
    overlay.style.left = `${Math.round(left)}px`;
    overlay.style.top = `${Math.round(top)}px`;
    overlay.style.width = `${Math.round(width)}px`;
    overlay.style.height = `${Math.round(height)}px`;

    // UPDATE CONTENT (Only if not focused)
    const textEl = overlay.querySelector('.text');
    const stickyEl = overlay.querySelector('.sticky');
    
    // Check focus
    const activeEl = document.activeElement;
    const isFocused = activeEl && (activeEl === textEl || overlay.contains(activeEl));

    if (!isFocused && textEl) {
        const masterData = activeSheet.columns[masterCol]?.slots?.[g.slotIdx];
        const baseText = masterData?.text ?? '';

        if (g.slotIdx === 1 && g.systemsMeta) {
            const summaryHTML = _formatSystemsSummaryFromMeta(g.systemsMeta);
            if (textEl.innerHTML !== summaryHTML) {
                textEl.innerHTML = summaryHTML;
                textEl.setAttribute('contenteditable', 'false');
            }
        } else {
            if (textEl.textContent !== baseText) {
                textEl.textContent = baseText;
                textEl.setAttribute('contenteditable', 'true');
            }
        }
    }
    
    // ALWAYS update Gate badges (in case enabled/disabled via modal)
    if (g.slotIdx === 4 && stickyEl) {
      const gate = _sanitizeGate(g?.gate);
      const passLabel = getPassLabelForGroup(g);
      let failLabel = '—';
      if (gate?.enabled && gate.failTargetColIdx != null) {
        const idx = gate.failTargetColIdx;
        if (Number.isFinite(idx)) failLabel = getProcessLabel(idx);
      }
      applyGateToSticky(stickyEl, gate, passLabel, failLabel);
    }
  });

  // FIX: Cleanup old overlays
  const allOverlays = Array.from(colsContainer.querySelectorAll('.merged-overlay'));
  allOverlays.forEach(el => {
      if (!processedKeys.has(el.dataset.mergeKey)) el.remove();
  });
}

/**
 * Renders overlay boxes for column groups (ranges).
 * Uses the updated logic from state.js.
 */
function renderGroupOverlays() {
  const colsContainer = $('cols');
  if (!colsContainer) return;
  const sheet = state.activeSheet;
  if (!sheet || !Array.isArray(sheet.groups)) return;

  // Verwijder oude overlays
  colsContainer.querySelectorAll('.group-header-overlay').forEach(el => el.remove());

  sheet.groups.forEach(g => {
    // Vind de kolommen in de DOM op basis van de lijst
    const colElements = g.cols.map(cIdx => colsContainer.querySelector(`.col[data-idx="${cIdx}"]`)).filter(Boolean);

    if (colElements.length === 0) return;

    // Bereken positie: minimale linker en maximale rechter grens
    let minLeft = Infinity;
    let maxRight = -Infinity;

    colElements.forEach(el => {
        // Gebruik getOffsetWithin voor relatieve positie binnen colsContainer
        const rect = getOffsetWithin(el, colsContainer);
        if (rect.x < minLeft) minLeft = rect.x;
        
        const right = rect.x + el.offsetWidth;
        if (right > maxRight) maxRight = right;
    });
    
    const width = maxRight - minLeft;

    const overlay = document.createElement('div');
    overlay.className = 'group-header-overlay';
    overlay.style.left = `${minLeft}px`;
    overlay.style.width = `${width}px`;
    
    // HTML: Titel label + witte lijn
    overlay.innerHTML = `
        <div class="group-header-label">${escapeHTML(g.title)}</div>
        <div class="group-header-line"></div>
    `;

    colsContainer.appendChild(overlay);
  });
}

/** Renders only the columns grid including stickies, badges, and connectors. */
function renderColumnsOnly(openModalFn) {
  const activeSheet = state.activeSheet;
  if (!activeSheet) return;

  ensureMergeGroupsLoaded();

  const colsContainer = $('cols');
  if (!colsContainer) return;

  const variantLetterMap = computeVariantLetterMap(activeSheet);

  const project = state.project || state.data;
  const { outIdByUid, outTextByUid, outTextByOutId } = buildGlobalOutputMaps(project);

  // Fallback: UI-selecties gebruiken vaak OUTx labels (incl. merged-slaves).
  // buildGlobalOutputMaps() slaat merged-slaves over, waardoor OUT2 soms geen tekst heeft.
  // Daarom vullen we ontbrekende OUTx->tekst aan via state.getAllOutputs().
  try {
    const all = typeof state.getAllOutputs === 'function' ? state.getAllOutputs() : {};
    Object.keys(all || {}).forEach((k) => {
      if (!outTextByOutId[k] && all[k]) outTextByOutId[k] = all[k];
    });
  } catch {}

  const offsets = computeCountersBeforeActiveSheet(project, project.activeSheetId, outIdByUid);

  let localInCounter = 0;
  let localOutCounter = 0;

  const stats = { happy: 0, neutral: 0, sad: 0 };

  const frag = document.createDocumentFragment();

  activeSheet.columns.forEach((col, colIdx) => {
    if (col.isVisible === false) return;

    let myInputId = '';
    let myOutputId = '';

    const inputSlot = col.slots?.[2];
    const outputSlot = col.slots?.[4];

    const bundleIdsForInput = getLinkedBundleIdsFromInputSlot(inputSlot);
    const bundleLabelsForInput = bundleIdsForInput.map((bid) => _getBundleLabel(project, bid));

    // Direct links (outputs) + bundle links
    const tokens = getLinkedSourcesFromInputSlot(inputSlot);
    const resolved = resolveLinkedSourcesToOutAndText(tokens, outIdByUid, outTextByUid, outTextByOutId);

    if (bundleLabelsForInput.length) {
      // Bundels zijn bedoeld om de input compact te houden: toon alleen bundelnaam/nam(en) in de tag.
      myInputId = _joinSemiText(bundleLabelsForInput);
    } else if (resolved.ids.length) {
      myInputId = _joinSemiText(resolved.ids);
    } else if (inputSlot?.text?.trim()) {
      localInCounter += 1;
      myInputId = `IN${offsets.inStart + localInCounter}`;
    }

    if (outputSlot?.text?.trim() && !isMergedSlave(colIdx, 4)) {
      localOutCounter += 1;
      myOutputId = `OUT${offsets.outStart + localOutCounter}`;
    }

    const colEl = document.createElement('div');
    // NIEUW: is-group class toevoegen aan de kolom
    colEl.className = `col ${col.isParallel ? 'is-parallel' : ''} ${col.isVariant ? 'is-variant' : ''} ${col.isGroup ? 'is-group' : ''}`;
    colEl.dataset.idx = colIdx;

    if (col.isVariant) colEl.dataset.route = variantLetterMap[colIdx] || 'A';
    else colEl.dataset.route = '';

    const actionsEl = document.createElement('div');
    actionsEl.className = 'col-actions';
    actionsEl.innerHTML = `
      <button class="btn-col-action btn-arrow" data-action="move" data-dir="-1" type="button">←</button>
      <button class="btn-col-action btn-arrow" data-action="move" data-dir="1" type="button">→</button>
      ${
        colIdx > 0
          ? `<button class="btn-col-action btn-parallel ${col.isParallel ? 'active' : ''}" data-action="parallel" type="button">∥</button>`
          : ''
      }
      ${
        colIdx > 0
          ? `<button class="btn-col-action btn-variant ${col.isVariant ? 'active' : ''}" data-action="variant" type="button">🔀</button>`
          : ''
      }
      
      <button class="btn-col-action btn-group ${col.isGroup ? 'active' : ''}" data-action="group" title="Markeer als onderdeel van groep" type="button">🧩</button>
      
      <button class="btn-col-action btn-conditional ${col.isConditional ? 'active' : ''}" data-action="conditional" title="Voorwaardelijke stap (optioneel)" type="button">⚡</button>
      
      <button class="btn-col-action btn-question ${col.isQuestion ? 'active' : ''}" data-action="question" title="Markeer als vraag" type="button">❓</button>
      
      <button class="btn-col-action btn-hide-col" data-action="hide" type="button">👁️</button>
      <button class="btn-col-action btn-add-col-here" data-action="add" type="button">+</button>
      <button class="btn-col-action btn-delete-col" data-action="delete" type="button">×</button>
    `;
    colEl.appendChild(actionsEl);

    const slotsEl = document.createElement('div');
    slotsEl.className = 'slots';

    col.slots.forEach((slot, slotIdx) => {
      if (slotIdx === 3) {
        if (slot.processStatus === 'HAPPY') stats.happy += 1;
        else if (slot.processStatus === 'NEUTRAL') stats.neutral += 1;
        else if (slot.processStatus === 'SAD') stats.sad += 1;
      }

      let displayText = slot.text;
      let isLinked = false;

      if (slotIdx === 2) {
        const bundleIds = getLinkedBundleIdsFromInputSlot(slot);
        const bundleLabels = bundleIds.map((bid) => _getBundleLabel(project, bid));

        const tokens = getLinkedSourcesFromInputSlot(slot);
        const resolved = resolveLinkedSourcesToOutAndText(tokens, outIdByUid, outTextByUid, outTextByOutId);

        const parts = [];
        if (bundleLabels.length) {
          parts.push(...bundleLabels);
        } else if (resolved.texts.length) {
          parts.push(...resolved.texts);
        }

        if (parts.length) {
          displayText = _joinSemiText(parts);
          isLinked = true;
        }
      }

      const scoreBadgeHTML = buildScoreBadges({ slotIdx, slot });

      let statusClass = '';
      if (slotIdx === 3 && slot.processStatus) statusClass = `status-${slot.processStatus.toLowerCase()}`;

      let typeIcon = '📝';
      if (slot.type === 'Afspraak') typeIcon = '📅';

      let extraStickyClass = '';
      let extraStickyStyle = '';

      if (getMergeGroup(colIdx, slotIdx)) {
        extraStickyClass = 'merged-source';
        extraStickyStyle = 'visibility:hidden; pointer-events:none;';
      }

      const slotDiv = document.createElement('div');
      slotDiv.className = 'slot';
      slotDiv.innerHTML = buildSlotHTML({
        colIdx,
        slotIdx,
        slot,
        statusClass,
        typeIcon,
        myInputId,
        myOutputId,
        isLinked,
        scoreBadgeHTML,
        extraStickyClass,
        extraStickyStyle
      });

      const textEl = slotDiv.querySelector('.text');
      const stickyEl = slotDiv.querySelector('.sticky');
      if (textEl) textEl.textContent = displayText;

      const isMergedSource = !!getMergeGroup(colIdx, slotIdx);

      if (!isMergedSource) attachStickyInteractions({ stickyEl, textEl, colIdx, slotIdx, openModalFn });

      if (!isLinked && textEl && !isMergedSource) {
        textEl.addEventListener(
          'input',
          () => {
            state.updateStickyText(colIdx, slotIdx, textEl.textContent);
            scheduleSyncRowHeights();
          },
          { passive: true }
        );
        textEl.addEventListener(
          'blur',
          () => {
            state.updateStickyText(colIdx, slotIdx, textEl.textContent);
            scheduleSyncRowHeights();
          },
          { passive: true }
        );
      }

      slotsEl.appendChild(slotDiv);
    });

    colEl.appendChild(slotsEl);
    frag.appendChild(colEl);

    renderConnector({ frag, activeSheet, colIdx, variantLetterMap });
  });

  colsContainer.replaceChildren(frag);
  renderStats(stats);
  scheduleSyncRowHeights();
  
  // RENDER GROUP OVERLAYS NA ELKE RENDER
  requestAnimationFrame(() => renderGroupOverlays()); 
}

/** Updates one rendered text cell when state signals a text-only change. */
function updateSingleText(colIdx, slotIdx) {
  const colsContainer = $('cols');
  const colEl = colsContainer?.querySelector?.(`.col[data-idx="${colIdx}"]`);
  if (!colEl) return false;

  const slot = state.activeSheet.columns[colIdx]?.slots?.[slotIdx];
  if (!slot) return false;

  const g = getMergeGroup(colIdx, slotIdx);
  if (g && slotIdx === g.slotIdx) {
    const active = document.activeElement;
    if (active && active.closest('.merged-overlay')) return true;
    scheduleSyncRowHeights();
    return true;
  }

  const slotEl = colEl.querySelector(`.sticky[data-col="${colIdx}"][data-slot="${slotIdx}"] .text`);
  if (!slotEl) return false;

  if (slotEl && slotEl.isContentEditable && document.activeElement === slotEl) return true;

  if (slotIdx === 2) {
    const project = state.project || state.data;
    const { outIdByUid, outTextByUid, outTextByOutId } = buildGlobalOutputMaps(project);

    // Zelfde fallback als in renderColumnsOnly(): OUTx labels uit UI moeten altijd een tekst kunnen tonen.
    try {
      const all = typeof state.getAllOutputs === 'function' ? state.getAllOutputs() : {};
      Object.keys(all || {}).forEach((k) => {
        if (!outTextByOutId[k] && all[k]) outTextByOutId[k] = all[k];
      });
    } catch {}

    const bundleIds = getLinkedBundleIdsFromInputSlot(slot);
    const bundleLabels = bundleIds.map((bid) => _getBundleLabel(project, bid));

    const tokens = getLinkedSourcesFromInputSlot(slot);
    const resolved = resolveLinkedSourcesToOutAndText(tokens, outIdByUid, outTextByUid, outTextByOutId);

    const parts = [];
    if (bundleLabels.length) {
      parts.push(...bundleLabels);
    } else if (resolved.texts.length) {
      parts.push(...resolved.texts);
    }

    if (parts.length) {
      slotEl.textContent = _joinSemiText(parts);
      return true;
    }
  }

  slotEl.textContent = slot.text ?? '';
  return true;
}

/** Renders the full board for the currently active sheet. */
export function renderBoard(openModalFn) {
  _openModalFn = openModalFn || _openModalFn;

  const activeSheet = state.activeSheet;
  if (!activeSheet) return;

  ensureMergeGroupsLoaded();

  renderSheetSelect();
  renderHeader(activeSheet);
  ensureRowHeaders();
  renderColumnsOnly(_openModalFn);
}

/** Applies a state update reason to the UI with minimal re-render where possible. */
export function applyStateUpdate(meta, openModalFn) {
  _openModalFn = openModalFn || _openModalFn;

  const reason = meta?.reason || 'full';

  if (reason === 'text' && Number.isFinite(meta?.colIdx) && Number.isFinite(meta?.slotIdx)) {
    const ok = updateSingleText(meta.colIdx, meta.slotIdx);
    if (ok) return;
  }

  if (reason === 'title') return;

  if (reason === 'sheet' || reason === 'sheets') {
    const activeSheet = state.activeSheet;
    if (activeSheet) {
      ensureMergeGroupsLoaded();
      renderSheetSelect();
      renderHeader(activeSheet);
    }
    renderColumnsOnly(_openModalFn);
    return;
  }

  if (reason === 'columns' || reason === 'transition' || reason === 'details' || reason === 'groups') {
    renderColumnsOnly(_openModalFn);
    return;
  }

  renderBoard(_openModalFn);
}

/** Installs delegated handlers for column action buttons and prevents duplicate binding. */
export function setupDelegatedEvents() {
  if (_delegatedBound) return;
  _delegatedBound = true;

  const act = (e) => {
    const btn = e.target.closest('.btn-col-action');
    if (!btn) return;

    const action = btn.dataset.action;
    if (!action) return;

    if (e.type === 'mousedown' && performance.now() - _lastPointerDownTs < 250) return;
    if (e.type === 'pointerdown') _lastPointerDownTs = performance.now();

    e.preventDefault();
    e.stopPropagation();

    const colEl = btn.closest('.col');
    if (!colEl) return;

    const idx = parseInt(colEl.dataset.idx, 10);
    if (!Number.isFinite(idx)) return;

    switch (action) {
      case 'move':
        state.moveColumn(idx, parseInt(btn.dataset.dir, 10));
        break;
      case 'delete':
        if (confirm('Kolom verwijderen?')) state.deleteColumn(idx);
        break;
      case 'add':
        state.addColumn(idx);
        break;
      case 'hide':
        state.setColVisibility(idx, false);
        break;
      case 'parallel':
        state.toggleParallel?.(idx);
        break;
      case 'variant':
        // NIEUWE LOGICA: Open de modal ipv direct toggle
        openVariantModal(idx);
        break;
      case 'conditional':
        openLogicModal(idx);
        break;
      case 'group':
        openGroupModal(idx);
        break;
      case 'question':
        state.toggleQuestion?.(idx);
        break;
    }
  };

  document.addEventListener('pointerdown', act, true);
  document.addEventListener('mousedown', act, true);
  document.addEventListener('touchstart', act, { capture: true, passive: false });
}