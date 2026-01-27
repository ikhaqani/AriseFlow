// dom.js
import { state } from './state.js';
import { IO_CRITERIA, PROCESS_STATUSES } from './config.js';
import { openEditModal, saveModalDetails, openLogicModal, openGroupModal, openVariantModal } from './modals.js';

const $ = (id) => document.getElementById(id);

let _openModalFn = null;
let _delegatedBound = false;
let _syncRaf = 0;
let _lastPointerDownTs = 0;

const MERGE_LS_PREFIX = 'ssipoc.mergeGroups.v2';

let _mergeGroups = [];

// =========================================================
// HULPFUNCTIE: Diepte berekenen (Alleen voor data-attribuut)
// =========================================================
function getDependencyDepth(colIdx) {
  const sheet = state.activeSheet;
  if (!sheet || !Array.isArray(sheet.variantGroups)) return 0;

  let depth = 0;
  let currentIdx = colIdx;
  let safety = 0;

  while (true) {
    const parentGroup = sheet.variantGroups.find((g) => g.variants.includes(currentIdx));
    if (parentGroup) {
      depth++;
      const p =
        parentGroup.parents && parentGroup.parents.length > 0 ? parentGroup.parents[0] : parentGroup.parentColIdx;

      if (p !== undefined) {
        if (typeof p === 'string' && p.includes('::')) break;
        currentIdx = p;
      } else {
        break;
      }
    } else {
      break;
    }
    if (++safety > 10) break;
  }
  return depth;
}

/* =========================================================
   ROUTE PREFIX (OPTIE 2): sheetPrefix + routeLetter
   - scoped route labels zoals: RF1-A, RF1-A.1, RF2-B.2, etc.
   ========================================================= */
function getSheetRoutePrefix() {
  const project = state.project || state.data;
  const sh = state.activeSheet;
  if (!project || !sh) return 'RF?';

  const idx = (project.sheets || []).findIndex((s) => s.id === sh.id);
  const n = idx >= 0 ? idx + 1 : 1;

  return `RF${n}`;
}

function getScopedRouteLabel(label) {
  const p = getSheetRoutePrefix();
  const s = String(label || '').trim();
  return s ? `${p}-${s}` : `${p}-?`;
}

/** Follow-up route label (A.1/A.2/...) op basis van routeLabel + volgorde binnen sheet. */
function getFollowupRouteLabel(colIdx) {
  const sh = state.activeSheet;
  if (!sh) return null;

  const col = sh.columns?.[colIdx];
  if (!col) return null;

  const manualRoute = String(col.routeLabel || '').trim(); // "A" / "B" / ...
  const isSplitStart = !!col.isVariant;
  if (!manualRoute || isSplitStart) return null;

  let count = 1;
  for (let i = 0; i < colIdx; i++) {
    const c = sh.columns?.[i];
    if (!c) continue;
    if (c.isVariant) continue;
    if (String(c.routeLabel || '').trim() === manualRoute) count++;
  }

  return `${manualRoute}.${count}`; // A.1 / A.2 / ...
}

/* =========================================================
   ROUTE INDENT + KLEUR + BADGE (jouw gewenste gedrag)
   - hoofdproces: geen indent (level 0), default geel
   - route A/B: indent level 1 + kleur
   - A.1/A.2: indent level 2 + dezelfde kleur als A
   - badge niet in connector, maar boven de Proces post-it (slotIdx 3)
   ========================================================= */
function getRouteLabelForColumn(colIdx, variantLetterMap) {
  const sh = state.activeSheet;
  const col = sh?.columns?.[colIdx];
  if (!col) return null;

  // Split-start (variant): label komt uit variantLetterMap (A / B / A.1 / etc)
  if (col.isVariant) {
    const v = String(variantLetterMap?.[colIdx] || '').trim();
    return v || null;
  }

  // Follow-up subprocess: A.1 / A.2 / ...
  const follow = getFollowupRouteLabel(colIdx);
  if (follow) return follow;

  // hoofdproces: geen route
  return null;
}

function getIndentLevelFromRouteLabel(routeLabel) {
  const s = String(routeLabel || '').trim();
  if (!s) return 0;
  return s.split('.').length; // A => 1, A.1 => 2, etc
}

function getRouteBaseLetter(routeLabel) {
  const s = String(routeLabel || '').trim();
  if (!s) return null;
  return s.split('.')[0].toUpperCase();
}

function getRouteColorByLetter(letter) {
  const L = String(letter || '').toUpperCase();

  // Geen groen/oranje/rood. Palet t/m G: paars, magenta, blauw, teal, indigo, amber, blauwgrijs
  if (L === 'A') return { bg: 'rgb(49, 74, 12)', text: '#FFFFFF' };
  if (L === 'B') return { bg: '#D81B60', text: '#FFFFFF' };
  if (L === 'C') return { bg: '#2979FF', text: '#FFFFFF' };
  if (L === 'D') return { bg: '#00ACC1', text: '#FFFFFF' };
  if (L === 'E') return { bg: '#3949AB', text: '#FFFFFF' };
  if (L === 'F') return { bg: 'rgb(121, 121, 121)', text: '#111111' };
  if (L === 'G') return { bg: '#3c7690', text: '#FFFFFF' };
  return null;
}

// =========================================================
// MERGE GRADIENT HELPERS (diagonale multi-kleur merged stickies)
// =========================================================
function _parseColorToRgb(input) {
  const s = String(input || '').trim();
  if (!s) return null;

  // #RRGGBB
  if (s[0] === '#' && s.length === 7) {
    const r = parseInt(s.slice(1, 3), 16);
    const g = parseInt(s.slice(3, 5), 16);
    const b = parseInt(s.slice(5, 7), 16);
    if ([r, g, b].every((v) => Number.isFinite(v))) return { r, g, b };
    return null;
  }

  // rgb(...) / rgba(...)
  const m = s.match(/rgba?\(([^)]+)\)/i);
  if (m) {
    const parts = m[1].split(',').map((x) => parseFloat(x.trim()));
    if (parts.length >= 3 && parts.slice(0, 3).every((v) => Number.isFinite(v))) {
      return { r: parts[0], g: parts[1], b: parts[2] };
    }
  }

  return null;
}

function _relLuminance({ r, g, b }) {
  // sRGB -> linear
  const toLin = (v) => {
    const x = v / 255;
    return x <= 0.03928 ? x / 12.92 : Math.pow((x + 0.055) / 1.055, 2.4);
  };
  const R = toLin(r);
  const G = toLin(g);
  const B = toLin(b);
  return 0.2126 * R + 0.7152 * G + 0.0722 * B;
}

function _pickTextColorFromColors(colors) {
  const rgbs = (colors || []).map(_parseColorToRgb).filter(Boolean);
  if (!rgbs.length) return null;
  const avgLum = rgbs.map(_relLuminance).reduce((a, b) => a + b, 0) / rgbs.length;
  return avgLum > 0.55 ? '#111111' : '#FFFFFF';
}

function buildDiagonalMergedGradient(colors, angleDeg = 135) {
  const list = (colors || []).map((c) => String(c || '').trim()).filter(Boolean);
  if (!list.length) return null;
  if (list.length === 1) return list[0];

  const n = list.length;
  const stops = [];
  for (let i = 0; i < n; i++) {
    const a = (i / n) * 100;
    const b = ((i + 1) / n) * 100;
    stops.push(`${list[i]} ${a.toFixed(2)}% ${b.toFixed(2)}%`);
  }
  return `linear-gradient(${angleDeg}deg, ${stops.join(', ')})`;
}

function buildRoutePillHTML(routeLabel) {
  const s = String(routeLabel || '').trim();
  if (!s) return '';
  return `<div class="route-pill">Route ${escapeHTML(getScopedRouteLabel(s))}</div>`;
}

function buildConditionalPillHTML(isConditional) {
  if (!isConditional) return '';
  return `<div class="conditional-pill" title="Conditionele stap">⚡</div>`;
}

function buildSupplierTopBadgesHTML({ routeLabel, isConditional }) {
  const a = buildRoutePillHTML(routeLabel);
  const b = buildConditionalPillHTML(isConditional);
  if (!a && !b) return '';
  return `<div class="supplier-top-badges">${a}${b}</div>`;
}

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
  const failTargetColIdx = Number.isFinite(Number(gate.failTargetColIdx)) ? Number(gate.failTargetColIdx) : null;
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
    (groups || []).find((x) => x.slotIdx === slotIdx && Array.isArray(x.cols) && x.cols.includes(colIdx)) || null
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
      const isLinked = tokens.some((t) => (_looksLikeOutId(t) ? true : !!(t && outIdByUid && outIdByUid[t])));

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

  const _slot4 = state.activeSheet?.columns?.[clickedColIdx]?.slots?.[4];
let gate = slotIdx === 4
  ? (_sanitizeGate(cur?.gate) || _sanitizeGate(_slot4?.outputData?.gate) || _sanitizeGate(_slot4?.gate) || { enabled: false, failTargetColIdx: null })
  : null;

  let systemsMeta = null;

  if (slotIdx === 1) {
    const existing = _sanitizeSystemsMeta(cur?.systemsMeta) || _getSystemMetaFromSlot(masterCol);
    if (existing) {
      systemsMeta = existing;
    } else {
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

    const scores = (clean.systems || []).map((s) => _computeSystemScore(s)).filter((x) => Number.isFinite(Number(x)));
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
        <input data-sys-name="${idx}" class="modal-input" style="margin:0; height:40px;" placeholder="Bijv. ARIA / EPIC / Radiotherapieweb / Monaco..." value="${escapeAttr(
          sys.name || ''
        )}" />

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
          <label class="sysfit-nvt-toggle" style="display:flex;align-items:center;gap:10px;margin:8px 0 12px 0;font-size:14px;cursor:pointer;opacity:.95;">
            <input type="checkbox" class="sysfitNvtCheck" ${(sys.qa && typeof sys.qa === "object" && sys.qa.__nvt===true) ? "checked" : ""} />
            <strong>NVT</strong> <span style="opacity:.75;font-weight:500;">(niet beoordeelbaar, bv. extern systeem)</span>
          </label>
          <div style="font-size:13px; opacity:0.75; margin-bottom:12px;">Beantwoord per vraag hoe goed dit systeem jouw taak ondersteunt.</div>

          <div data-sys-qs="${idx}" style="display:flex; flex-direction:column; gap:14px;"></div>

          <div style="margin-top:12px; font-size:14px; opacity:0.85;">
            Score (dit systeem): <strong data-sys-score="${idx}">—</strong>
          </div>
        </div>
      `;

      systemsWrapEl.appendChild(card);

      const qWrap = card.querySelector(`[data-sys-qs="${idx}"]`);

      window.__sysfitDomSystemsByIdx = window.__sysfitDomSystemsByIdx || {};
      window.__sysfitDomSystemsByIdx[String(idx)] = sys;
      /* __SYSFIT_DOM_SYSID_MAP_V1__ */
window.__sysfitDomSystemsById = window.__sysfitDomSystemsById || {};
// sysId: gebruik sys.id als aanwezig, anders idx
try { window.__sysfitDomSystemsById[String(sys.id ?? idx)] = sys; } catch {}
const qa = sys.qa && typeof sys.qa === "object" ? sys.qa : {};

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
          <div style="margin-top:10px;">
            <input
              type="text"
              class="modal-input"
              data-qid-note="${q.id}"
              placeholder="Opmerking (optioneel)..."
              value="${escapeAttr(qa[q.id + '_note'] || '')}"
              style="margin:0; font-size:13px; opacity:0.8; height:36px;"
            />
          </div>
        `;
        qWrap.appendChild(row);

        // Listener voor knoppen (update score)
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

        // Listener voor opmerkingen (TOEGEVOEGD)
        row.querySelectorAll('input[data-qid-note]').forEach((inp) => {
          inp.addEventListener('input', () => {
            const qid = inp.getAttribute('data-qid-note');
            if (!qid) return;
            systemsMeta.systems[idx].qa = systemsMeta.systems[idx].qa || {};
            systemsMeta.systems[idx].qa[qid + '_note'] = inp.value;
          });
        });
      });

      // Overige listeners
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

      const scoreEl = card.querySelector(`[data-sys-score=""]`);
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
        systemsMeta.systems = [systemsMeta.systems?.[0] || { name: '', legacy: false, future: '', qa: {}, score: null }];
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

    // ✅ HARD REQUIREMENT: als Procesvalidatie aan staat, moet Rework routing gekozen zijn
    if (slotIdx === 4 && gate?.enabled) {
      if (gate.failTargetColIdx == null || !Number.isFinite(Number(gate.failTargetColIdx))) {
        alert("Selecteer een waarde bij 'Routing bij Rework' of zet 'Validatiestap toevoegen' uit.");
        return;
      }
    }

    if (!mergeEnabled) {
      setMergeGroupForSlot(slotIdx, null);
      state.updateStickyText(clickedColIdx, slotIdx, v);


      // SINGLE-COL: persist Procesvalidatie (gate) also when merge is OFF
      if (slotIdx === 4) {
        const cleanGate = _sanitizeGate(gate);
        const slot = state.activeSheet?.columns?.[clickedColIdx]?.slots?.[4];
        if (slot) {
          const od = slot.outputData && typeof slot.outputData === 'object' ? slot.outputData : {};
          slot.outputData = { ...od, gate: cleanGate };
          slot.gate = cleanGate; // legacy compatibility
        }
        state.notify({ reason: 'columns' }, { clone: false });
      }
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
  const map = {};
  if (!activeSheet?.columns?.length) return map;

  const colGroups = {}; // key: parentColIdx, val: array of child indices

  if (Array.isArray(activeSheet.variantGroups)) {
    activeSheet.variantGroups.forEach((vg) => {
      const primaryParent = vg.parents && vg.parents.length > 0 ? vg.parents[0] : vg.parentColIdx;

      if (primaryParent !== undefined) {
        if (!colGroups[primaryParent]) colGroups[primaryParent] = [];
        vg.variants.forEach((vIdx) => colGroups[primaryParent].push(vIdx));
      }
    });
  }

  function assignLabels(parentIdx, prefix) {
    const children = colGroups[parentIdx];
    if (!children) return;

    children.sort((a, b) => a - b);

    children.forEach((childIdx, i) => {
      let myLabel = '';
      if (!prefix) myLabel = _toLetter(i);
      else myLabel = `${prefix}.${i + 1}`;

      map[childIdx] = myLabel;
      assignLabels(childIdx, myLabel);
    });
  }

  const allChildren = new Set();
  Object.values(colGroups).forEach((list) => list.forEach((c) => allChildren.add(c)));

  const rootParents = Object.keys(colGroups)
    .map(Number)
    .filter((p) => !allChildren.has(p));
  rootParents.forEach((root) => assignLabels(root, ''));

  let legacyCounter = 0;
  activeSheet.columns.forEach((col, i) => {
    if (col.isVariant && !map[i]) map[i] = _toLetter(legacyCounter++);
  });

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
  routeBadgeHTML = '',
  extraStickyClass = '',
  extraStickyStyle = ''
}) {
  const procEmoji = slotIdx === 3 && slot.processStatus ? getProcessEmoji(slot.processStatus) : '';
  const leanIcon = slotIdx === 3 && slot.processValue ? getLeanIcon(slot.processValue) : '';
  const workExpIcon = slotIdx === 3 ? getWorkExpMeta(slot?.workExp)?.icon || '' : '';
  const linkIcon = isLinked ? '🔗' : '';

  const b1 = slotIdx === 3 && slot.type ? typeIcon : '';
  const b2 = slotIdx === 3 ? leanIcon : '';
  const b3 = slotIdx === 3 ? workExpIcon : slotIdx === 2 && isLinked ? linkIcon : '';
  const b4 = slotIdx === 3 ? procEmoji : '';

  const editableAttr = isLinked ? 'contenteditable="false" data-linked="true"' : 'contenteditable="true"';

  return `
    <div class="sticky ${statusClass} ${extraStickyClass}" style="${escapeAttr(
    extraStickyStyle
  )}" data-col="${colIdx}" data-slot="${slotIdx}">
      ${routeBadgeHTML}
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
    renderGroupOverlays();
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

  const MIN_ROW_HEIGHT = 160;
  const heights = Array(6).fill(MIN_ROW_HEIGHT);

  cols.forEach((col) => {
    const slotNodes = col.querySelectorAll('.slots .slot');
    for (let r = 0; r < 6; r++) {
      const slot = slotNodes[r];
      if (!slot) continue;
      const sticky = slot.firstElementChild;
      if (!sticky) continue;

      // ✅ FIX: neem margins mee (o.a. supplier-indent via margin-top)
      const cs = getComputedStyle(sticky);
      const mt = parseFloat(cs.marginTop || '0') || 0;
      const mb = parseFloat(cs.marginBottom || '0') || 0;

      const rectH = sticky.getBoundingClientRect?.().height || 0;
      const scrollH = sticky.scrollHeight || 0;

      const h = Math.ceil(Math.max(scrollH, rectH) + mt + mb) + 32;

      if (h > heights[r]) heights[r] = h;
    }
  });

  const globalMax = Math.max(...heights);
  for (let r = 0; r < 6; r++) heights[r] = globalMax;

  for (let r = 0; r < 6; r++) {
    const hStr = `${heights[r]}px`;
    if (rowHeaders[r]) rowHeaders[r].style.height = hStr;

    cols.forEach((col) => {
      const slotNodes = col.querySelectorAll('.slots .slot');
      if (slotNodes[r]) slotNodes[r].style.height = hStr;
    });
  }

  const gapSize = 20;
  const processOffset = heights[0] + heights[1] + heights[2] + 3 * gapSize;

  colsContainer.querySelectorAll('.col-connector').forEach((c) => {
    if (!c.classList.contains('parallel-connector') && !c.classList.contains('combo-connector')) {
      const colEl = c.closest('.col-connector').previousElementSibling;
      const indentLevel = colEl ? parseInt(colEl.dataset.indentLevel || 0, 10) : 0;

      const extraPadding = indentLevel * 30;

      c.style.paddingTop = `${processOffset + extraPadding}px`;
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

    // ✅ toegevoegd: help/vraagteken buttons mogen focus niet kapen
    if (
      e.target.closest(
        '.sticky-grip, .qa-score-badge, .id-tag, .badges-row, .workexp-badge, .btn-col-action, .col-actions, .qmark-btn, [data-action="help"]'
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

  const connEl = document.createElement('div');

  if (hasVariant && hasParallel) {
    connEl.className = 'col-connector combo-connector variant-connector';
    connEl.innerHTML = `<div class="parallel-badge">||</div>`;
    frag.appendChild(connEl);
    return;
  }

  if (hasVariant) {
    connEl.className = 'col-connector variant-connector';
    connEl.innerHTML = '';
    frag.appendChild(connEl);
    return;
  }

  if (hasParallel) {
    connEl.className = 'col-connector combo-connector';
    connEl.innerHTML = `<div class="parallel-badge">||</div>`;
    frag.appendChild(connEl);
    return;
  }

  connEl.className = 'col-connector';
  connEl.innerHTML = `<div class="connector-active"></div>`;
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

/** Builds system summary HTML using inline Legacy pills for merged System overlays. */
function _formatSystemsSummaryFromMeta(meta) {
  const clean = _sanitizeSystemsMeta(meta);
  if (!clean) return '';

  const systems = clean.systems || [];

  const lineHTML = (s) => {
    const nm = String(s?.name || '').trim() || '—';
    const legacy = !!s?.legacy;
    return `<div class="sys-line"><span class="sys-name">${escapeHTML(nm)}</span>${
      legacy ? `<span class="legacy-tag" aria-label="Legacy">Legacy</span>` : ''
    }</div>`;
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

  const activeSheet = state.activeSheet;
  if (!activeSheet) return;

  const groups = getAllMergeGroupsSanitized();

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

    overlay.style.left = `${Math.round(left)}px`;
    overlay.style.top = `${Math.round(top)}px`;
    overlay.style.width = `${Math.round(width)}px`;
    overlay.style.height = `${Math.round(height)}px`;

    const textEl = overlay.querySelector('.text');
    const stickyEl = overlay.querySelector('.sticky');

    if (stickyEl) {
      const variantLetterMap = computeVariantLetterMap(activeSheet);

      const rootDefault = getComputedStyle(document.documentElement).getPropertyValue('--sticky-bg').trim() || '#ffd600';

      const colors = visibleCols.map((cIdx) => {
        const rl = getRouteLabelForColumn(cIdx, variantLetterMap);
        const base = getRouteBaseLetter(rl);
        const clr = base ? getRouteColorByLetter(base) : null;
        return String(clr?.bg || rootDefault).trim();
      });

      const gradient = buildDiagonalMergedGradient(colors, 135);
      if (gradient) stickyEl.style.setProperty('--merged-bg', gradient);

      const mergedText = _pickTextColorFromColors(colors);
      if (mergedText) stickyEl.style.setProperty('--sticky-text', mergedText);
    }

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

  const allOverlays = Array.from(colsContainer.querySelectorAll('.merged-overlay'));
  allOverlays.forEach((el) => {
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

  colsContainer.querySelectorAll('.group-header-overlay').forEach((el) => el.remove());

  sheet.groups.forEach((g) => {
    const colElements = g.cols
      .map((cIdx) => colsContainer.querySelector(`.col[data-idx="${cIdx}"]`))
      .filter(Boolean);

    if (colElements.length === 0) return;

    let minLeft = Infinity;
    let maxRight = -Infinity;

    colElements.forEach((el) => {
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

    const tokens = getLinkedSourcesFromInputSlot(inputSlot);
    const resolved = resolveLinkedSourcesToOutAndText(tokens, outIdByUid, outTextByUid, outTextByOutId);

    if (bundleLabelsForInput.length) {
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
    colEl.className = `col ${col.isParallel ? 'is-parallel' : ''} ${col.isVariant ? 'is-variant' : ''} ${
      col.isGroup ? 'is-group' : ''
    }`;
    colEl.dataset.idx = colIdx;

    const depth = getDependencyDepth(colIdx);
    if (depth > 0) colEl.dataset.depth = depth;
    if (depth > 0) colEl.style.transformOrigin = 'top center';

    // ====== ROUTE LABEL + INDENT LEVEL + KLEUR ======
    const routeLabel = getRouteLabelForColumn(colIdx, variantLetterMap);
    const indentLevel = getIndentLevelFromRouteLabel(routeLabel);
    const baseLetter = getRouteBaseLetter(routeLabel);

    colEl.dataset.routeLabel = routeLabel || '';
    colEl.dataset.indentLevel = String(indentLevel);
    colEl.dataset.route = routeLabel || '';

    if (indentLevel > 0) {
      colEl.classList.add('is-route-col');
      colEl.style.setProperty('--indent-level', String(indentLevel));

      const clr = getRouteColorByLetter(baseLetter);
      if (clr) {
        colEl.style.setProperty('--sticky-bg', clr.bg);
        colEl.style.setProperty('--sticky-text', clr.text);
      }
    }

    const inner = document.createElement('div');
    inner.className = 'col-inner';
    Object.assign(inner.style, { position: 'relative', zIndex: '2' });

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
    inner.appendChild(actionsEl);

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

        const tokens2 = getLinkedSourcesFromInputSlot(slot);
        const resolved2 = resolveLinkedSourcesToOutAndText(tokens2, outIdByUid, outTextByUid, outTextByOutId);

        const parts = [];
        if (bundleLabels.length) parts.push(...bundleLabels);
        else if (resolved2.texts.length) parts.push(...resolved2.texts);

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

      const routeBadgeHTML =
        slotIdx === 0 ? buildSupplierTopBadgesHTML({ routeLabel, isConditional: !!col.isConditional }) : '';

      if (slotIdx === 0 && indentLevel > 0) {
        const indentPx = indentLevel * 30;
        extraStickyStyle += `margin-top: ${indentPx}px;`;
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
        routeBadgeHTML,
        extraStickyClass,
        extraStickyStyle
      });

      const textEl = slotDiv.querySelector('.text');
      const stickyEl = slotDiv.querySelector('.sticky');
      // SINGLE-COL: render gate on normal sticky (no merge required)
      if (stickyEl && Number(stickyEl?.dataset?.slot) === 4) {
        const _colIdx = Number(stickyEl?.dataset?.col);
        const _slot = state.activeSheet?.columns?.[_colIdx]?.slots?.[4];
        const gate = _sanitizeGate(_slot?.outputData?.gate) || _sanitizeGate(_slot?.gate);

        const passLabel = (typeof getPassLabelForGroup === 'function')
          ? getPassLabelForGroup({ slotIdx: 4, cols: [_colIdx], master: _colIdx })
          : ((state.activeSheet?.columns?.[_colIdx + 1]) ? getProcessLabel(_colIdx + 1) : '—');

        let failLabel = '—';
        if (gate?.enabled && gate.failTargetColIdx != null) {
          const idx = Number(gate.failTargetColIdx);
          if (Number.isFinite(idx)) failLabel = getProcessLabel(idx);
        }

        applyGateToSticky(stickyEl, gate, passLabel, failLabel);
      }
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

    inner.appendChild(slotsEl);
    colEl.appendChild(inner);
    frag.appendChild(colEl);

    renderConnector({ frag, activeSheet, colIdx, variantLetterMap });
  });

  colsContainer.replaceChildren(frag);
  renderStats(stats);
  scheduleSyncRowHeights();

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
    const now = performance.now();

    // ---- HELP / vraagteken buttons (niet de kolom ❓) ----
    const helpBtn = e.target.closest?.('[data-action="help"], .qmark-btn');
    if (helpBtn) {
      if ((e.type === 'mousedown' || e.type === 'click') && now - _lastPointerDownTs < 250) return;
      if (e.type === 'pointerdown' || e.type === 'touchstart') _lastPointerDownTs = now;

      e.preventDefault();
      e.stopPropagation();

      const helpKey =
        String(helpBtn.dataset.helpKey || helpBtn.dataset.criterionKey || helpBtn.dataset.key || '').trim() || null;

      // Dispatch zodat je eigen help-module dit kan oppakken
      document.dispatchEvent(
        new CustomEvent('ssipoc:help', {
          detail: { helpKey, source: 'dom', target: helpBtn }
        })
      );
      return;
    }

    const btn = e.target.closest('.btn-col-action');
    if (!btn) return;

    const action = btn.dataset.action;
    if (!action) return;

    // Voorkom dubbele triggers (pointerdown -> mousedown/click)
    if ((e.type === 'mousedown' || e.type === 'click') && now - _lastPointerDownTs < 250) return;

    // Pointer/touch markeren
    if (e.type === 'pointerdown' || e.type === 'touchstart') _lastPointerDownTs = now;

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
        if (typeof state.toggleParallel === 'function') state.toggleParallel(idx);
        else if (typeof state.toggleParallelColumn === 'function') state.toggleParallelColumn(idx);
        else console.warn('Parallel toggle missing on state');
        break;
      case 'variant':
        openVariantModal(idx);
        break;
      case 'conditional':
        openLogicModal(idx);
        break;
      case 'group':
        openGroupModal(idx);
        break;

      // ✅ FIX: vraagteken (kolomactie) met fallbacks, zodat hij altijd werkt
      case 'question': {
        if (typeof state.toggleQuestion === 'function') {
          state.toggleQuestion(idx);
          break;
        }
        if (typeof state.toggleQuestionColumn === 'function') {
          state.toggleQuestionColumn(idx);
          break;
        }

        // absolute fallback: toggle flag in state + re-render
        const sh = state.activeSheet;
        const col = sh?.columns?.[idx];
        if (col) {
          col.isQuestion = !col.isQuestion;
          if (typeof state.notify === 'function') state.notify({ reason: 'columns' });
          else renderColumnsOnly(_openModalFn);
        } else {
          console.warn('Question toggle missing on state');
        }
        break;
      }
    }
  };

  document.addEventListener('pointerdown', act, true);
  document.addEventListener('mousedown', act, true);
  document.addEventListener('click', act, true);
  document.addEventListener('touchstart', act, { capture: true, passive: false });
}
/* __SYSFIT_DOM_NVT_PERSIST__ */
(() => {
  function _getSysByIdx(idx) {
    const m = window.__sysfitDomSystemsByIdx && typeof window.__sysfitDomSystemsByIdx === "object"
      ? window.__sysfitDomSystemsByIdx
      : null;
    return m ? m[String(idx)] : null;
  }

  function _setNvt(sys, isNvt) {
    if (!sys || typeof sys !== "object") return;
    if (!sys.qa || typeof sys.qa !== "object") sys.qa = {};
    sys.qa.__nvt = !!isNvt;

    if (sys.qa.__nvt) {
      ["q1","q2","q3","q4","q5"].forEach((k) => { delete sys.qa[k]; });
      ["q1_note","q2_note","q3_note","q4_note","q5_note"].forEach((k) => { delete sys.qa[k]; });
      delete sys.score;
      delete sys.calculatedScore;
    }
  }

  function _syncCardUi(card, isNvt) {
    if (!card) return;
    try {
      const scoreEl = card.querySelector('[data-sys-score]');
      if (scoreEl) scoreEl.textContent = isNvt ? "" : (scoreEl.textContent || "—");

      const qWrap = card.querySelector("[data-sys-qs]");
      if (qWrap) {
        qWrap.style.opacity = isNvt ? "0.35" : "1";
        qWrap.style.pointerEvents = isNvt ? "none" : "auto";
      }
    } catch {}
  }

  document.addEventListener("change", (ev) => {
    const cb = ev.target && ev.target.closest ? ev.target.closest(".sysfitNvtCheck") : null;
    if (!cb) return;

    const idx = cb.getAttribute("data-sys-idx");
    const sys = _getSysByIdx(idx);
    if (!sys) return;

    const next = !!cb.checked;
    _setNvt(sys, next);

    const card = cb.closest(".system-card");
    _syncCardUi(card, next);

    try { if (typeof state !== "undefined" && typeof state.notify === "function") state.notify(); } catch {}
  });
})();



/* __SYSFIT_DOM_NVT_APPLY_HOOK__ */
(() => {
  function _syncAllCardsToModel() {
    const m = (window.__sysfitDomSystemsByIdx && typeof window.__sysfitDomSystemsByIdx === "object")
      ? window.__sysfitDomSystemsByIdx
      : null;
    if (!m) return;

    document.querySelectorAll(".system-card").forEach((card) => {
      // idx vinden via score element (data-sys-score="0", "1", ...)
      const scoreEl = card.querySelector("[data-sys-score]");
      const idx = scoreEl ? String(scoreEl.getAttribute("data-sys-score") || "").trim() : "";
      if (!idx) return;

      const sys = m[idx];
      if (!sys || typeof sys !== "object") return;

      const cb = card.querySelector(".sysfitNvtCheck");
      const isNvt = !!(cb && cb.checked);

      if (!sys.qa || typeof sys.qa !== "object") sys.qa = {};
      sys.qa.__nvt = isNvt;

      // Als NVT: maak antwoorden leeg zodat apply/scores niet terugkomen
      if (isNvt) {
        ["q1","q2","q3","q4","q5"].forEach((k) => { delete sys.qa[k]; });
        ["q1_note","q2_note","q3_note","q4_note","q5_note"].forEach((k) => { delete sys.qa[k]; });
        delete sys.score;
        delete sys.calculatedScore;
      }
    });
  }

  // Capture-phase: vóór andere click handlers (zoals "Toepassen")
  document.addEventListener("click", (ev) => {
    const el = ev.target;
    if (!el) return;

    // Button/element dat “Toepassen” triggert (best-effort)
    const btn = el.closest ? el.closest("button, [data-action], .std-btn") : null;
    if (!btn) return;

    const txt = String(btn.textContent || "").trim().toLowerCase();
    const act = String(btn.getAttribute("data-action") || "").trim().toLowerCase();

    if (txt.includes("toepassen") || act.includes("apply") || act.includes("save")) {
      _syncAllCardsToModel();
    }
  }, true);
})();




/* __SYSFIT_NVT_COMMIT_ON_APPLY__ */
(() => {
  function _getActiveSheet(project) {
    if (!project || !Array.isArray(project.sheets)) return null;
    const id = project.activeSheetId;
    return project.sheets.find(s => s && s.id === id) || project.sheets[0] || null;
  }

  function _findActiveColIdx(sheet) {
    // Best-effort: probeer bekende velden
    const cand = [
      sheet?.activeColIdx, sheet?.activeColIndex,
      (window && window.__activeColIdx),
      (window && window.__editColIdx),
      (window && window.__activeColIndex),
      (typeof state !== "undefined" ? state.activeColIdx : undefined),
      (typeof state !== "undefined" ? state.activeColIndex : undefined),
      (typeof state !== "undefined" && state.ui ? state.ui.activeColIdx : undefined),
      (typeof state !== "undefined" && state.ui ? state.ui.colIdx : undefined),
    ].find(v => Number.isFinite(Number(v)));
    return Number.isFinite(Number(cand)) ? Number(cand) : null;
  }

  function _commitNvtToProjectFromModal() {
    try {
      if (typeof state === "undefined") return;
      const project = state.project || state.data;
      const sheet = _getActiveSheet(project);
      if (!sheet || !Array.isArray(sheet.columns)) return;

      const colIdx = _findActiveColIdx(sheet);
      if (colIdx == null || !sheet.columns[colIdx]) {
        // fallback: als dom-mapping bestaat, commit alleen daarin (minimaal)
        return;
      }

      const col = sheet.columns[colIdx];
      const sysSlot = col.slots && col.slots[1] ? col.slots[1] : null;
      window.__sysfitActiveSysSlot = sysSlot; // injected


      // Locatie waar dom.js meestal in schrijft:
      const sd = sysSlot && sysSlot.systemData && typeof sysSlot.systemData === "object"
        ? sysSlot.systemData
        : (sysSlot ? (sysSlot.systemData = {}) : null);

      if (!sd) return;

      // Zorg dat we altijd een array "systems" hebben; als er systemsMeta is, ook daarin zetten.
      const systemsArr = Array.isArray(sd.systems) ? sd.systems : (sd.systems = []);
      const meta = sd.systemsMeta && typeof sd.systemsMeta === "object" ? sd.systemsMeta : null;
      const metaSystems = meta && Array.isArray(meta.systems) ? meta.systems : null;

      // Lees UI cards in volgorde (zelfde volgorde als systems meestal)
      const cards = Array.from(document.querySelectorAll(".system-card"));
      cards.forEach((card, i) => {
        const cb = card.querySelector(".sysfitNvtCheck");
        if (!cb) return;
        const isNvt = !!cb.checked;

        const sysObj = systemsArr[i];
        if (sysObj && typeof sysObj === "object") {
          if (!sysObj.qa || typeof sysObj.qa !== "object") sysObj.qa = {};
          sysObj.qa.__nvt = isNvt;
          if (isNvt) {
            ["q1","q2","q3","q4","q5"].forEach(k => { delete sysObj.qa[k]; });
            ["q1_note","q2_note","q3_note","q4_note","q5_note"].forEach(k => { delete sysObj.qa[k]; });
            delete sysObj.score;
            delete sysObj.calculatedScore;
          }
        }

        const metaSys = metaSystems && metaSystems[i];
        if (metaSys && typeof metaSys === "object") {
          if (!metaSys.qa || typeof metaSys.qa !== "object") metaSys.qa = {};
          metaSys.qa.__nvt = isNvt;
          if (isNvt) {
            ["q1","q2","q3","q4","q5"].forEach(k => { delete metaSys.qa[k]; });
            ["q1_note","q2_note","q3_note","q4_note","q5_note"].forEach(k => { delete metaSys.qa[k]; });
            delete metaSys.score;
          }
        }
      });

      // Trigger rerender
      try { if (typeof state.notify === "function") state.notify(); } catch {}
    } catch (e) {
      console.warn("SYSFIT NVT commit failed:", e);
    }
  }

  // Capture-phase: vóór de bestaande 'Toepassen' handler(s)
  document.addEventListener("click", (ev) => {
    const btn = ev.target && ev.target.closest ? ev.target.closest("button") : null;
    if (!btn) return;
    const txt = String(btn.textContent || "").trim().toLowerCase();
    if (txt === "toepassen" || txt.includes("toepassen")) {
      _commitNvtToProjectFromModal();
    }
  }, true);
})();




/* __SYSFIT_NVT_COMMIT_ON_APPLY_V2__ */
(() => {
  function _getActiveSysSlot() {
    return (window.__sysfitActiveSysSlot && typeof window.__sysfitActiveSysSlot === "object")
      ? window.__sysfitActiveSysSlot
      : null;
  }

  function _ensureSystemArrays(sysSlot) {
    if (!sysSlot) return { systems: null, metaSystems: null };
    const sd = (sysSlot.systemData && typeof sysSlot.systemData === "object")
      ? sysSlot.systemData
      : (sysSlot.systemData = {});

    const systems = Array.isArray(sd.systems) ? sd.systems : (sd.systems = []);

    const meta = (sd.systemsMeta && typeof sd.systemsMeta === "object") ? sd.systemsMeta : null;
    const metaSystems = (meta && Array.isArray(meta.systems)) ? meta.systems : null;

    return { systems, metaSystems };
  }

  function _applyNvtToSys(sysObj, isNvt) {
    if (!sysObj || typeof sysObj !== "object") return;
    if (!sysObj.qa || typeof sysObj.qa !== "object") sysObj.qa = {};
    sysObj.qa.__nvt = !!isNvt;

    if (sysObj.qa.__nvt) {
      ["q1","q2","q3","q4","q5"].forEach(k => { delete sysObj.qa[k]; });
      ["q1_note","q2_note","q3_note","q4_note","q5_note"].forEach(k => { delete sysObj.qa[k]; });
      delete sysObj.score;
      delete sysObj.calculatedScore;
    }
  }

  function _commitFromUi() {
    const sysSlot = _getActiveSysSlot();
    if (!sysSlot) return;

    const { systems, metaSystems } = _ensureSystemArrays(sysSlot);
    const cards = Array.from(document.querySelectorAll(".system-card"));
    if (!cards.length) return;

    cards.forEach((card, i) => {
      const cb = card.querySelector(".sysfitNvtCheck");
      if (!cb) return;
      const isNvt = !!cb.checked;

      // primary storage
      if (systems && systems[i]) _applyNvtToSys(systems[i], isNvt);

      // optional meta mirror
      if (metaSystems && metaSystems[i]) _applyNvtToSys(metaSystems[i], isNvt);
    });
  }

  // Capture-phase: vóór bestaande Apply handlers
  document.addEventListener("click", (ev) => {
    const btn = ev.target && ev.target.closest ? ev.target.closest("button") : null;
    if (!btn) return;
    const t = String(btn.textContent || "").trim().toLowerCase();
    if (t.includes("toepassen")) {
      _commitFromUi();
      try { if (typeof state !== "undefined" && typeof state.notify === "function") state.notify(); } catch {}
    }
  }, true);
})();
/* __END_SYSFIT_NVT_COMMIT_ON_APPLY_V2__ */





/* __SYSFIT_DOM_NVT_PERSIST_V4__ */
(() => {
  function _getSysById(sysId) {
    const m = (window.__sysfitDomSystemsById && typeof window.__sysfitDomSystemsById === "object")
      ? window.__sysfitDomSystemsById
      : null;
    if (m && sysId != null && m[String(sysId)]) return m[String(sysId)];
    return null;
  }

  function _ensureQa(sys) {
    if (!sys || typeof sys !== "object") return null;
    if (!sys.qa || typeof sys.qa !== "object") sys.qa = {};
    return sys.qa;
  }

  function _setNvt(sys, isNvt) {
    const qa = _ensureQa(sys);
    if (!qa) return;
    qa.__nvt = !!isNvt;

    if (qa.__nvt) {
      ["q1","q2","q3","q4","q5"].forEach(k => { delete qa[k]; });
      ["q1_note","q2_note","q3_note","q4_note","q5_note"].forEach(k => { delete qa[k]; });
      delete sys.score;
      delete sys.calculatedScore;
    }
  }

  function _syncCardUi(card, isNvt) {
    if (!card) return;

    // vragenblok (in dom.js is dit meestal de div met data-sys-qs)
    const qWrap = card.querySelector("[data-sys-qs]");
    if (qWrap) {
      qWrap.style.opacity = isNvt ? "0.35" : "1";
      qWrap.style.pointerEvents = isNvt ? "none" : "auto";

      const controls = qWrap.querySelectorAll("input,button,select,textarea");
      controls.forEach((el) => {
        if (!el) return;

        if (isNvt) {
          // markeer alleen als wij hem uitzetten (zodat we later niet per ongeluk andere disables overschrijven)
          if (!el.disabled) el.dataset.nvtDisabled = "1";
          el.disabled = true;

          // leeg maken zodat je niet NVT + antwoorden tegelijk hebt
          if (el.tagName === "INPUT") {
            const t = (el.getAttribute("type") || "").toLowerCase();
            if (t === "radio" || t === "checkbox") el.checked = false;
          }
          if (el.tagName === "TEXTAREA" || el.tagName === "INPUT") {
            const t = (el.getAttribute("type") || "").toLowerCase();
            if (t === "text") el.value = "";
          }
        } else {
          // alleen terugzetten als wij hem eerder disabled hebben
          if (el.dataset && el.dataset.nvtDisabled === "1") {
            el.disabled = false;
            delete el.dataset.nvtDisabled;
          }
        }
      });
    }

    // score UI (leeglaten is het makkelijkst/duidelijkst)
    const scoreEl = card.querySelector('[data-sys-score], .sys-score, strong[data-sys-score]');
    if (scoreEl) scoreEl.textContent = isNvt ? "" : (scoreEl.textContent || "—");
  }

  function _syncAll() {
    document.querySelectorAll(".system-card").forEach((card) => {
      const cb = card.querySelector(".sysfitNvtCheck");
      const isNvt = !!(cb && cb.checked);
      _syncCardUi(card, isNvt);
    });
  }

  // CHANGE handler: NVT opslaan + UI sync
  document.addEventListener("change", (ev) => {
    const cb = ev.target && ev.target.closest ? ev.target.closest(".sysfitNvtCheck") : null;
    if (!cb) return;

    const card = cb.closest(".system-card");
    const sysId = (card && (card.getAttribute("data-sys-id") || (card.dataset ? card.dataset.sysId : ""))) || null;
    const sys = _getSysById(sysId);

    if (!sys) {
      console.warn("NVT: geen sys object gevonden voor sysId=", sysId);
      _syncAll();
      return;
    }

    const next = !!cb.checked;
    _setNvt(sys, next);

    // ook mirrors (systemData / systemsMeta) best-effort, zodat export/score nooit terugvalt naar 0%
    try {
      const sysSlot = (window.__sysfitActiveSysSlot && typeof window.__sysfitActiveSysSlot === "object")
        ? window.__sysfitActiveSysSlot
        : null;
      const sd = sysSlot && sysSlot.systemData && typeof sysSlot.systemData === "object" ? sysSlot.systemData : null;

      const sid = String(sys.id ?? sysId ?? "");
      if (sd && Array.isArray(sd.systems)) {
        const tgt = sd.systems.find(x => x && String(x.id ?? "") === sid);
        if (tgt) _setNvt(tgt, next);
      }
      if (sd && sd.systemsMeta && Array.isArray(sd.systemsMeta.systems)) {
        const tgt2 = sd.systemsMeta.systems.find(x => x && String(x.id ?? "") === sid);
        if (tgt2) _setNvt(tgt2, next);
      }
    } catch {}

    _syncCardUi(card, next);

    try { if (typeof state !== "undefined" && typeof state.notify === "function") state.notify(); } catch {}
  }, true);

  // Init sync + keep in sync bij rerenders
  const _init = () => {
    _syncAll();
    const mo = new MutationObserver(() => _syncAll());
    mo.observe(document.body, { childList: true, subtree: true });
  };

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", _init);
  else _init();
})();
/* __END_SYSFIT_DOM_NVT_PERSIST_V4__ */


