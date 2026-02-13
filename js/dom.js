// dom.js
console.info("[AriseFlow dom.js] build=20260129_FLOW_ALL_VIEW_FIXED_UNMERGE_PATCH");
import { state } from './state.js';
import { IO_CRITERIA, PROCESS_STATUSES } from './config.js';
import { openEditModal, saveModalDetails, openLogicModal, openGroupModal, openVariantModal } from './modals.js';

const $ = (id) => document.getElementById(id);

// === CSS Injection voor "Flow All" view ===
function injectFlowAllStyles() {
  if (document.getElementById('flow-all-styles')) return;
  const style = document.createElement('style');
  style.id = 'flow-all-styles';
  style.textContent = `
    .tab-btn.special-tab {
        border-bottom: 2px solid #ff9f43;
        font-weight: bold;
    }
    .flow-separator {
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        min-width: 60px;
        margin: 0 10px;
        height: 100%;
        position: relative;
        z-index: 10;
        opacity: 0.8;
    }
    .flow-separator-line {
        position: absolute;
        top: 180px; /* Onder de headers */
        bottom: 20px;
        width: 2px;
        background: repeating-linear-gradient(to bottom, #444, #444 10px, transparent 10px, transparent 20px);
    }
    .flow-separator-title {
        background-color: #222;
        color: #fff;
        padding: 10px 15px;
        border-radius: 20px;
        border: 1px solid #555;
        font-size: 13px;
        white-space: nowrap;
        writing-mode: vertical-rl;
        text-orientation: mixed;
        transform: rotate(180deg);
        z-index: 11;
        box-shadow: 0 4px 10px rgba(0,0,0,0.5);
    }
    body.view-mode-all .col-actions { display: none !important; } /* Geen acties in master view */
  `;
  document.head.appendChild(style);
}
injectFlowAllStyles();

// === Emoji rendering (board post-its) =======================================
function _escapeHtml(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

let _EMOJI_SEQ_RE;
try {
  _EMOJI_SEQ_RE = new RegExp(
    "(\\p{Extended_Pictographic}(?:\\uFE0F)?(?:\\u200D\\p{Extended_Pictographic}(?:\\uFE0F)?)*)",
    "gu"
  );
} catch (e) {
  _EMOJI_SEQ_RE = /([\uD83C-\uDBFF][\uDC00-\uDFFF])/g;
}

function renderTextWithEmojiSpan(rawText) {
  const s = String(rawText ?? "");
  if (!s) return "";
  const parts = s.split(_EMOJI_SEQ_RE);
  let out = "";
  for (let i = 0; i < parts.length; i++) {
    const part = parts[i] ?? "";
    if (!part) continue;
    if (i % 2 === 1) out += `<span class="emoji-inline">${part}</span>`;
    else out += _escapeHtml(part);
  }
  return out;
}
// ============================================================================

let _openModalFn = null;
let _delegatedBound = false;
let _syncRaf = 0;
let _lastPointerDownTs = 0;

const MERGE_LS_PREFIX = 'ssipoc.mergeGroups.v2';
let _mergeGroups = [];

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
      const p = parentGroup.parents && parentGroup.parents.length > 0 ? parentGroup.parents[0] : parentGroup.parentColIdx;
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

function getFollowupRouteLabel(colIdx) {
  const sh = state.activeSheet;
  if (!sh) return null;
  const col = sh.columns?.[colIdx];
  if (!col) return null;
  const manualRoute = String(col.routeLabel || '').trim();
  const isSplitStart = !!col.isVariant;
  if (!manualRoute || isSplitStart) return null;
  let count = 1;
  for (let i = 0; i < colIdx; i++) {
    const c = sh.columns?.[i];
    if (!c) continue;
    if (c.isVariant) continue;
    if (String(c.routeLabel || '').trim() === manualRoute) count++;
  }
  return `${manualRoute}.${count}`;
}

function getRouteLabelForColumn(colIdx, variantLetterMap) {
  const sh = state.activeSheet;
  const col = sh?.columns?.[colIdx];
  if (!col) return null;
  if (col.isVariant) {
    const v = String(variantLetterMap?.[colIdx] || '').trim();
    return v || null;
  }
  const follow = getFollowupRouteLabel(colIdx);
  if (follow) return follow;
  return null;
}

function getIndentLevelFromRouteLabel(routeLabel) {
  const s = String(routeLabel || '').trim();
  if (!s) return 0;
  return s.split('.').length;
}

function getRouteBaseLetter(routeLabel) {
  const s = String(routeLabel || '').trim();
  if (!s) return null;
  return s.split('.')[0].toUpperCase();
}

function getRouteColorByLetter(letter) {
  const L = String(letter || '').toUpperCase();
  if (L === 'A') return { bg: 'rgb(49, 74, 12)', text: '#FFFFFF' };
  if (L === 'B') return { bg: '#D81B60', text: '#FFFFFF' };
  if (L === 'C') return { bg: '#2979FF', text: '#FFFFFF' };
  if (L === 'D') return { bg: '#00ACC1', text: '#FFFFFF' };
  if (L === 'E') return { bg: '#3949AB', text: '#FFFFFF' };
  if (L === 'F') return { bg: 'rgb(121, 121, 121)', text: '#111111' };
  if (L === 'G') return { bg: '#3c7690', text: '#FFFFFF' };
  return null;
}

function _parseColorToRgb(input) {
  const s = String(input || '').trim();
  if (!s) return null;
  if (s[0] === '#' && s.length === 7) {
    const r = parseInt(s.slice(1, 3), 16);
    const g = parseInt(s.slice(3, 5), 16);
    const b = parseInt(s.slice(5, 7), 16);
    if ([r, g, b].every((v) => Number.isFinite(v))) return { r, g, b };
    return null;
  }
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

function _mergeKey() {
  const pid = state?.project?.id || state?.project?.name || 'project';
  const sid = state?.activeSheet?.id || state?.activeSheet?.name || 'sheet';
  return `${MERGE_LS_PREFIX}:${pid}:${sid}`;
}

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

function saveMergeGroups(groups) {
  try {
    localStorage.setItem(_mergeKey(), JSON.stringify(groups || []));
  } catch {}
}

function ensureMergeGroupsLoaded() {
  const key = _mergeKey();
  if (_mergeGroups.__key !== key) {
    const g = loadMergeGroups();
    _mergeGroups = Array.isArray(g) ? g : [];
    _mergeGroups.__key = key;
  }
}

function getMergedMasterColIdx(colIdx, slotIdx) {
  try { ensureMergeGroupsLoaded(); } catch (_) {}
  const c = Number(colIdx);
  const s = Number(slotIdx);
  const groups = Array.isArray(_mergeGroups) ? _mergeGroups : [];
  const g = groups.find((g) =>
    Number(g?.slotIdx) === s &&
    Array.isArray(g?.cols) &&
    g.cols.map(Number).includes(c)
  );
  if (!g) return null;
  const candidates = [
    g.masterCol, g.masterIdx, g.master, g.masterColIdx,
    Array.isArray(g.cols) && g.cols.length ? Math.min(...g.cols.map(Number)) : null
  ];
  for (const v of candidates) {
    const n = Number(v);
    if (Number.isFinite(n)) return n;
  }
  return null;
}

function getEffectiveInputSlot(colIdx, inputSlot) {
  try {
    if (!inputSlot) return inputSlot;
    if (typeof isMergedSlave !== "function") return inputSlot;
    if (typeof getMergedMasterColIdx !== "function") return inputSlot;
    if (!isMergedSlave(colIdx, 2)) return inputSlot;
    const m = getMergedMasterColIdx(colIdx, 2);
    const mi = Number(m);
    if (!Number.isFinite(mi)) return inputSlot;
    return (state?.activeSheet?.columns?.[mi]?.slots?.[2]) || inputSlot;
  } catch (_) {
    return inputSlot;
  }
}

function isContiguousZeroBased(cols) {
  if (!cols || cols.length < 2) return false;
  const s = [...new Set(cols)].sort((a, b) => a - b);
  const min = s[0];
  const max = s[s.length - 1];
  return s.length === max - min + 1;
}

function _sanitizeGate(gate) {
  if (!gate || typeof gate !== 'object') return null;
  const enabled = !!gate.enabled;
  const failTargetColIdx = Number.isFinite(Number(gate.failTargetColIdx)) ? Number(gate.failTargetColIdx) : null;
  return { enabled, failTargetColIdx };
}

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
      if (qa && qa.__nvt === true) {
        return { name, legacy, future, qa: { __nvt: true }, score: null };
      }
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
  
  // Gate op index 5
  const gate = slotIdx === 5 ? _sanitizeGate(g?.gate) : null;
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

function getAllMergeGroupsSanitized() {
  ensureMergeGroupsLoaded();
  const out = [];
  for (const g of _mergeGroups) {
    const s = sanitizeGroupForActiveSheet(g);
    if (s) out.push(s);
  }
  return out;
}

function updateMergeStorage(slotIdx, groupToAdd, colsToClear) {
  ensureMergeGroupsLoaded();
  const s = Number(slotIdx);
  const next = [];
  const excludeCols = new Set();
  if (groupToAdd && Array.isArray(groupToAdd.cols)) {
    groupToAdd.cols.forEach(c => excludeCols.add(Number(c)));
  }
  if (Array.isArray(colsToClear)) {
    colsToClear.forEach(c => excludeCols.add(Number(c)));
  }
  for (const g of _mergeGroups) {
    if (Number(g.slotIdx) !== s) {
      next.push(g);
      continue;
    }
    const gCols = Array.isArray(g.cols) ? g.cols.map(Number) : [];
    const hasOverlap = gCols.some(c => excludeCols.has(c));
    if (!hasOverlap) {
      next.push(g);
    }
  }
  if (groupToAdd) {
    next.push(groupToAdd);
  }
  _mergeGroups = next;
  _mergeGroups.__key = _mergeKey();
  saveMergeGroups(_mergeGroups);
}

function setMergeGroupForSlot(slotIdx, groupOrNull) {
  updateMergeStorage(slotIdx, groupOrNull, groupOrNull?.cols || []);
}

function getMergeGroup(colIdx, slotIdx) {
  const groups = getAllMergeGroupsSanitized();
  return groups.find((g) => g.slotIdx === slotIdx && g.cols.includes(colIdx)) || null;
}

function isMergedSlave(colIdx, slotIdx) {
  const g = getMergeGroup(colIdx, slotIdx);
  return !!g && colIdx !== g.master;
}

function makeId(prefix = 'id') {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return `${prefix}_${crypto.randomUUID()}`;
  }
  return `${prefix}_${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

function mergeKeyFor(project, sheet) {
  const pid = project?.id || project?.name || project?.projectTitle || 'project';
  const sid = sheet?.id || sheet?.name || 'sheet';
  return `${MERGE_LS_PREFIX}:${pid}:${sid}`;
}

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

function sanitizeGroupForSheet(sheet, g) {
  const n = sheet?.columns?.length ?? 0;
  if (!n) return null;
  const slotIdx = Number(g?.slotIdx);
  
  // UPDATE: Geldige merge slots zijn nu 1, 2, 5
  if (![1, 2, 5].includes(slotIdx)) return null;

  const cols = Array.isArray(g?.cols) ? g.cols.map((x) => Number(x)).filter(Number.isFinite) : [];
  const uniq = [...new Set(cols)].filter((c) => c >= 0 && c < n);
  if (uniq.length < 2) return null;
  if (!isContiguousZeroBased(uniq)) return null;
  let master = Number(g?.master);
  if (!Number.isFinite(master) || !uniq.includes(master)) master = uniq[0];
  
  // Gate zit nu op index 5
  const gate = slotIdx === 5 ? _sanitizeGate(g?.gate) : null;
  const systemsMeta = slotIdx === 1 ? _sanitizeSystemsMeta(g?.systemsMeta) : null;
  return { slotIdx, cols: uniq.sort((a, b) => a - b), master, gate, systemsMeta };
}

function getMergeGroupsSanitizedForSheet(project, sheet) {
  const raw = loadMergeGroupsRawFor(project, sheet);
  return raw.map((g) => sanitizeGroupForSheet(sheet, g)).filter(Boolean);
}

function getMergeGroupForCellInSheet(groups, colIdx, slotIdx) {
  return (
    (groups || []).find((x) => x.slotIdx === slotIdx && Array.isArray(x.cols) && x.cols.includes(colIdx)) || null
  );
}

function isMergedSlaveInSheet(groups, colIdx, slotIdx) {
  const g = getMergeGroupForCellInSheet(groups, colIdx, slotIdx);
  return !!g && colIdx !== g.master;
}

function _looksLikeOutId(v) {
  const s = String(v || '').trim();
  return !!s && /^OUT\d+$/.test(s);
}

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

function getLinkedBundleIdsFromInputSlot(slot) {
  if (!slot || typeof slot !== 'object') return [];
  const tokens = [];
  const modernKeyExists = 'linkedBundleIds' in slot;
  const hasArray = Array.isArray(slot.linkedBundleIds);
  if (hasArray) {
    slot.linkedBundleIds.forEach((x) => {
      const s = String(x || '').trim();
      if (s) tokens.push(s);
    });
  }
  if (!modernKeyExists && !hasArray) {
    const single = String(slot.linkedBundleId || '').trim();
    if (single) tokens.push(single);
  }
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
  if (!b) return null;
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

function resolveLinkedSourcesToOutAndText(tokens, outIdByUid, outTextByUid, outTextByOutId) {
  const ids = [];
  const texts = [];
  (Array.isArray(tokens) ? tokens : []).forEach((t) => {
    const s = String(t || '').trim();
    if (!s) return;
    if (_looksLikeOutId(s)) {
      const outId = s;
      const txt = String(outTextByOutId?.[outId] ?? '').trim() || outId;
      ids.push(outId);
      texts.push(txt);
      return;
    }
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

function buildGlobalOutputMaps(project) {
  const outIdByUid = {};
  const outTextByUid = {};
  const outTextByOutId = {};
  let outCounter = 0;
  (project?.sheets || []).forEach((sheet) => {
    const groups = getMergeGroupsSanitizedForSheet(project, sheet);
    (sheet?.columns || []).forEach((col, colIdx) => {
      if (col?.isVisible === false) return;
      
      // Output zit nu op slot 5
      const outSlot = col?.slots?.[5];
      if (!outSlot?.text?.trim()) return;
      
      // Merge check op slot 5
      if (isMergedSlaveInSheet(groups, colIdx, 5)) return;

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

function computeCountersBeforeActiveSheet(project, activeSheetId, outIdByUid) {
  let inCount = 0;
  let outCount = 0;
  for (const sheet of project?.sheets || []) {
    if (sheet.id === activeSheetId) break;
    const groups = getMergeGroupsSanitizedForSheet(project, sheet);
    (sheet.columns || []).forEach((col, colIdx) => {
      if (col?.isVisible === false) return;
      const inSlot = col?.slots?.[2];
      // Output zit nu op slot 5
      const outSlot = col?.slots?.[5];
      const tokens = getLinkedSourcesFromInputSlot(inSlot);
      const isLinked = tokens.some((t) => (_looksLikeOutId(t) ? true : !!(t && outIdByUid && outIdByUid[t])));
      if (!isLinked && inSlot?.text?.trim()) inCount += 1;
      
      // Merge check op slot 5
      if (outSlot?.text?.trim() && !isMergedSlaveInSheet(groups, colIdx, 5)) outCount += 1;
    });
  }
  return { inStart: inCount, outStart: outCount };
}

function _getSystemMetaFromSlot(colIdx) {
  const sh = state.activeSheet;
  const slot = sh?.columns?.[colIdx]?.slots?.[1];
  const meta = slot?.systemData?.systemsMeta;
  return _sanitizeSystemsMeta(meta) || null;
}

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

function getProcessLabel(colIdx) {
  const sh = state.activeSheet;
  if (!sh) return `Kolom ${colIdx + 1}`;
  
  // Proces tekst op index 4
  const t = sh.columns?.[colIdx]?.slots?.[4]?.text;
  const s = String(t ?? '').trim();
  return s || `Kolom ${colIdx + 1}`;
}

function getNextVisibleColIdx(fromIdx) {
  const sh = state.activeSheet;
  if (!sh) return null;
  const n = sh.columns?.length ?? 0;
  for (let i = fromIdx + 1; i < n; i++) {
    if (sh.columns[i]?.isVisible !== false) return i;
  }
  return null;
}

function getPassLabelForGroup(group) {
  const maxCol = Math.max(...(group?.cols ?? [group?.master ?? 0]));
  const nextIdx = getNextVisibleColIdx(maxCol);
  if (nextIdx == null) return 'Einde proces';
  return getProcessLabel(nextIdx);
}

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

function escapeHTML(v) {
  return String(v ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function escapeAttr(v) {
  return String(v ?? '')
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

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

function removeMergeModal() {
  const old = document.getElementById('mergeModalOverlay');
  if (old) old.remove();
}

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

function openMergeModal(clickedColIdx, slotIdx, openModalFn) {
  const sh = state.activeSheet;
  if (!sh) return;
  const n = sh.columns?.length ?? 0;
  if (!n) return;

  // Geldige merge slots zijn nu 1, 2, 5
  if (![1, 2, 5].includes(slotIdx)) return;

  const slotName = slotIdx === 1 ? 'Systeem' : slotIdx === 2 ? 'Input' : 'Output';
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
  
  // Gate op index 5
  const _slotOut = state.activeSheet?.columns?.[clickedColIdx]?.slots?.[5];
  let gate = slotIdx === 5
    ? (_sanitizeGate(cur?.gate) || _sanitizeGate(_slotOut?.outputData?.gate) || _sanitizeGate(_slotOut?.gate) || { enabled: false, failTargetColIdx: null })
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

  // Gate block alleen tonen bij slot 5
  const gateBlockHTML =
    slotIdx === 5
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
      slotIdx === 5
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
  if (txt) txt.value = String(curText || "");
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

  if (slotIdx === 5) {
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
    if (qa && qa.__nvt === true) return null;
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
    function _isSysfitNvt(sysObj) {
      const qa = sysObj?.qa && typeof sysObj.qa === 'object' ? sysObj.qa : null;
      return !!(qa && qa.__nvt === true);
    }
    function _setSysfitNvt(sysIdx, enabled) {
      if (!systemsMeta?.systems?.[sysIdx]) return;
      const cur = systemsMeta.systems[sysIdx];
      const qa = cur.qa && typeof cur.qa === 'object' ? cur.qa : {};
      if (enabled) {
        cur.qa = { __nvt: true };
        cur.score = null;
      } else {
        const next = { ...qa };
        delete next.__nvt;
        cur.qa = Object.keys(next).length ? next : {};
        cur.score = Number.isFinite(Number(cur.score)) ? cur.score : null;
      }
    }
    systems.forEach((sys, idx) => {
      const card = document.createElement('div');
      card.className = 'system-card';
      Object.assign(card.style, { background: 'rgba(0,0,0,0.08)', border: '1px solid rgba(255,255,255,0.08)', borderRadius: '12px', padding: '14px' });
      const showDelete = !!systemsMeta.multi && systems.length > 1;
      const legacyChecked = !!sys.legacy;
      const isNvt = _isSysfitNvt(sys);
      card.innerHTML = `
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
          <div style="font-weight:800; letter-spacing:0.04em; opacity:0.9;">SYSTEEM</div>
          ${showDelete ? `<button data-sys-del="${idx}" class="std-btn danger-text" type="button" style="padding:8px 12px; border:1px solid rgba(244,67,54,0.35);">Verwijderen</button>` : `<span></span>`}
        </div>
        <div style="font-size:12px; letter-spacing:0.06em; font-weight:800; opacity:0.75; margin-bottom:6px;">SYSTEEMNAAM</div>
        <input data-sys-name="${idx}" class="modal-input" style="margin:0; height:40px;" placeholder="Bijv. ARIA / EPIC..." value="${escapeAttr(sys.name || '')}" />
        <div style="display:grid; grid-template-columns: 1fr 1fr; gap:12px; margin-top:12px; align-items:end;">
          <label style="display:flex; gap:10px; align-items:center; cursor:pointer; font-size:14px;">
            <input data-sys-legacy="${idx}" type="checkbox" ${legacyChecked ? 'checked' : ''} />
            <span>Legacy systeem</span>
          </label>
          <div>
            <div style="font-size:12px; letter-spacing:0.06em; font-weight:800; opacity:0.75; margin-bottom:6px;">TOEKOMSTIG SYSTEEM (VERWACHTING)</div>
            <input data-sys-future="${idx}" class="modal-input" style="margin:0; height:40px; ${legacyChecked ? '' : 'opacity:0.45; pointer-events:none;'}" placeholder="Bijv. ARIA / EPIC..." value="${escapeAttr(sys.future || '')}" />
          </div>
        </div>
        <div style="margin-top:16px; border-top:1px solid rgba(255,255,255,0.08); padding-top:14px;">
          <div style="font-weight:900; letter-spacing:0.06em; opacity:0.85; margin-bottom:6px;">SYSTEM FIT VRAGEN</div>
          <label class="sysfit-nvt-toggle" style="display:flex;align-items:center;gap:10px;margin:8px 0 12px 0;font-size:14px;cursor:pointer;opacity:.95;">
            <input type="checkbox" class="sysfitNvtCheck" data-sys-idx="${idx}" ${isNvt ? "checked" : ""} />
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
                  <button type="button" data-qid="${q.id}" data-opt="${o.key}" ${isNvt ? 'disabled="disabled"' : ''}
                    style="height:54px; border-radius:14px; border:1px solid rgba(255,255,255,0.10); background:${selectedBtn ? 'rgba(25,118,210,0.25)' : 'rgba(255,255,255,0.04)'}; color:inherit; font-size:15px; cursor:${isNvt ? 'not-allowed' : 'pointer'}; opacity:${isNvt ? '0.35' : '1'}; pointer-events:${isNvt ? 'none' : 'auto'};">
                    ${escapeHTML(o.label)}
                  </button>
                `;
              })
              .join('')}
          </div>
          <div style="margin-top:10px;">
            <input type="text" class="modal-input" data-qid-note="${q.id}" placeholder="Opmerking (optioneel)..." value="${escapeAttr(qa[q.id + '_note'] || '')}" ${isNvt ? 'disabled="disabled"' : ''} style="margin:0; font-size:13px; opacity:${isNvt ? '0.35' : '0.8'}; height:36px;" />
          </div>
        `;
        qWrap.appendChild(row);
        row.querySelectorAll('button[data-qid]').forEach((btn) => {
          btn.addEventListener('click', () => {
            const qid = btn.getAttribute('data-qid');
            const opt = btn.getAttribute('data-opt');
            if (!qid || !opt) return;
            systemsMeta.systems[idx].qa = systemsMeta.systems[idx].qa || {};
            systemsMeta.systems[idx].qa[qid] = opt;
            const sScore = _computeSystemScore(systemsMeta.systems[idx]);
            systemsMeta.systems[idx].score = Number.isFinite(Number(sScore)) ? Number(sScore) : null;
            _renderSystemCards();
          });
        });
        row.querySelectorAll('input[data-qid-note]').forEach((inp) => {
          inp.addEventListener('input', () => {
            const qid = inp.getAttribute('data-qid-note');
            if (!qid) return;
            systemsMeta.systems[idx].qa = systemsMeta.systems[idx].qa || {};
            systemsMeta.systems[idx].qa[qid + '_note'] = inp.value;
          });
        });
      });
      card.querySelectorAll('input[data-sys-name]').forEach((inp) => {
        inp.addEventListener('input', () => { systemsMeta.systems[idx].name = String(inp.value ?? ''); });
      });
      card.querySelectorAll('input[data-sys-legacy]').forEach((chk) => {
        chk.addEventListener('change', () => {
          systemsMeta.systems[idx].legacy = !!chk.checked;
          if (!systemsMeta.systems[idx].legacy) systemsMeta.systems[idx].future = '';
          _renderSystemCards();
        });
      });
      card.querySelectorAll('input[data-sys-future]').forEach((inp) => {
        inp.addEventListener('input', () => { systemsMeta.systems[idx].future = String(inp.value ?? ''); });
      });
      card.querySelectorAll('button[data-sys-del]').forEach((btn) => {
        btn.addEventListener('click', () => {
          const i = Number(btn.getAttribute('data-sys-del'));
          if (!Number.isFinite(i)) return;
          systemsMeta.systems.splice(i, 1);
          if (systemsMeta.systems.length === 0) systemsMeta.systems.push({ name: '', legacy: false, future: '', qa: {}, score: null });
          _renderSystemCards();
        });
      });
      const nvtChk = card.querySelector(`.sysfitNvtCheck[data-sys-idx="${idx}"]`);
      if (nvtChk) {
        nvtChk.addEventListener('change', () => {
          const enabled = !!nvtChk.checked;
          _setSysfitNvt(idx, enabled);
          _renderSystemCards();
        });
      }
      const scoreEl = card.querySelector(`[data-sys-score="${idx}"]`);
      const sScore = _isSysfitNvt(sys) ? null : _computeSystemScore(sys);
      if (scoreEl) scoreEl.textContent = Number.isFinite(Number(sScore)) ? `${Number(sScore)}%` : '—';
    });
    const overall = _computeOverallScore(systemsMeta);
    if (overallSystemScoreEl) {
      overallSystemScoreEl.textContent = Number.isFinite(Number(overall)) ? `${Number(overall)}%` : '—';
    }
  }

  if (slotIdx === 1) {
    systemsMeta = _sanitizeSystemsMeta(systemsMeta) || { multi: false, systems: [{ name: '', legacy: false, future: '', qa: {}, score: null }], activeSystemIdx: 0 };
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
      if (slotIdx === 5) syncPassLabel();
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
    if (slotIdx === 5) syncPassLabel();
  });

  syncMergeUI();
  function close() { removeMergeModal(); }
  overlay.addEventListener('click', (e) => { if (e.target === overlay) close(); });
  modal.querySelector('#mergeCancelBtn')?.addEventListener('click', close);
  
  // === MODIFIED: Opheffen Button met Metadata Reset Fix ===
  modal.querySelector('#mergeOffBtn')?.addEventListener('click', () => {
    // 1. Remove group
    updateMergeStorage(slotIdx, null, selected);

    // 2. Clean up metadata on the clicked column to prevent "Ghost" complex states
    // Dit zorgt ervoor dat de modal niet meer in "Multi" modus opent na unmerge
    if (slotIdx === 1 && systemsMeta) {
        systemsMeta.multi = false;
        // Keep only the first system to ensure single-mode validity
        if (Array.isArray(systemsMeta.systems) && systemsMeta.systems.length > 0) {
            systemsMeta.systems = [systemsMeta.systems[0]];
        }
        const clean = _sanitizeSystemsMeta(systemsMeta);
        if (clean) {
            _applySystemMetaToColumns([clickedColIdx], clean);
            const overall = _computeOverallScore(clean);
            const slot = state.activeSheet?.columns?.[clickedColIdx]?.slots?.[1];
            if (slot) {
                const sd = slot.systemData || {};
                slot.systemData = { ...sd, systemsMeta: clean, calculatedScore: overall };
            }
        }
    }

    if (slotIdx === 5 && gate) {
        gate.enabled = false;
        const clean = _sanitizeGate(gate);
        const slot = state.activeSheet?.columns?.[clickedColIdx]?.slots?.[5];
        if (slot) {
            const od = slot.outputData || {};
            slot.outputData = { ...od, gate: clean };
            slot.gate = clean;
        }
    }

    // Force state save/notify
    if (typeof state.notify === 'function') {
        state.notify({ reason: 'columns' }, { clone: false });
    }

    close();
    renderColumnsOnly(_openModalFn);
    scheduleSyncRowHeights();
  });
  // ========================================================

  modal.querySelector('#mergeSaveBtn')?.addEventListener('click', () => {
    const vEl = modal.querySelector('#mergeText');
    const v = vEl ? String(vEl.value ?? '') : String(curText || '');
    if (slotIdx === 5 && gate?.enabled) {
      if (gate.failTargetColIdx == null || !Number.isFinite(Number(gate.failTargetColIdx))) {
        alert("Selecteer een waarde bij 'Routing bij Rework' of zet 'Validatiestap toevoegen' uit.");
        return;
      }
    }
    if (!mergeEnabled) {
      updateMergeStorage(slotIdx, null, selected);
      if (slotIdx !== 2) state.updateStickyText(clickedColIdx, slotIdx, v);
      // Single-col Gate op slot 5
      if (slotIdx === 5) {
        const cleanGate = _sanitizeGate(gate);
        const slot = state.activeSheet?.columns?.[clickedColIdx]?.slots?.[5];
        if (slot) {
          const od = slot.outputData && typeof slot.outputData === 'object' ? slot.outputData : {};
          slot.outputData = { ...od, gate: cleanGate };
          slot.gate = cleanGate; 
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
            slot.systemData = { ...sd, systemsMeta: cleanMeta, calculatedScore: Number.isFinite(Number(overall)) ? Number(overall) : null };
          }
        }
      }
      close();
      renderColumnsOnly(_openModalFn);
      scheduleSyncRowHeights();
      return;
    }
    if (selected.length < 2) { alert('Selecteer minimaal 2 aaneengesloten kolommen.'); return; }
    if (!isContiguousZeroBased(selected)) { alert('De selectie bevat onderbrekingen. Selecteer een aaneengesloten reeks.'); return; }
    const finalMaster = selected.includes(clickedColIdx) ? clickedColIdx : selected[0];
    const payload = { slotIdx, cols: selected, master: finalMaster, label: slotIdx === 1 ? 'Merged System' : 'Merged Output' };
    if (slotIdx === 5) payload.gate = _sanitizeGate(gate);
    if (slotIdx === 1) payload.systemsMeta = _sanitizeSystemsMeta(systemsMeta);
    updateMergeStorage(slotIdx, payload, null);
    selected.forEach((cIdx) => { if (slotIdx !== 2) state.updateStickyText(cIdx, slotIdx, v); });
    if (slotIdx === 1) {
      const cleanMeta = _sanitizeSystemsMeta(systemsMeta);
      if (cleanMeta) {
        const overall = _computeOverallScore(cleanMeta);
        _applySystemMetaToColumns(selected, cleanMeta);
        selected.forEach((cIdx) => {
          const slot = state.activeSheet?.columns?.[cIdx]?.slots?.[1];
          if (!slot) return;
          const sd = slot.systemData && typeof slot.systemData === 'object' ? slot.systemData : {};
          slot.systemData = { ...sd, systemsMeta: cleanMeta, calculatedScore: Number.isFinite(Number(overall)) ? Number(overall) : null };
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

function computeTTFSystemScore(sys) {
  const qa = sys?.qa && typeof sys.qa === 'object' ? sys.qa : {};
  const keys = [ { id: 'q1', map: TTF_FREQ_SCORES }, { id: 'q2', map: TTF_FREQ_SCORES }, { id: 'q3', map: TTF_FREQ_SCORES }, { id: 'q4', map: TTF_FREQ_SCORES }, { id: 'q5', map: TTF_IMPACT_SCORES } ];
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

function getProcessEmoji(status) {
  if (!status) return '';
  const s = PROCESS_STATUSES?.find?.((x) => x.value === status);
  return s?.emoji || '';
}

function _toLetter(i0) {
  const n = Number(i0);
  if (!Number.isFinite(n) || n < 0) return 'A';
  const base = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  return base[n] || `R${n + 1}`;
}

function computeVariantLetterMap(activeSheet) {
  const map = {};
  if (!activeSheet?.columns?.length) return map;
  const colGroups = {}; 
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
  const rootParents = Object.keys(colGroups).map(Number).filter((p) => !allChildren.has(p));
  rootParents.forEach((root) => assignLabels(root, ''));
  let legacyCounter = 0;
  activeSheet.columns.forEach((col, i) => {
    if (col.isVariant && !map[i]) map[i] = _toLetter(legacyCounter++);
  });
  return map;
}

function getWorkExpMeta(workExp) {
  const v = String(workExp || '').toUpperCase();
  if (v === 'OBSTACLE') return { icon: '🛠️', short: 'Obstakel', context: 'Kost energie' };
  if (v === 'ROUTINE') return { icon: '🤖', short: 'Routine', context: 'Saai & Repeterend' };
  if (v === 'FLOW') return { icon: '🚀', short: 'Flow', context: 'Geeft energie' };
  return null;
}

function getLeanIcon(val) {
  if (val === 'VA') return '💚';
  if (val === 'BNVA') return '⚖️';
  if (val === 'NVA') return '🗑️';
  return '';
}

function buildScoreBadges({ slotIdx, slot }) {
  let html = '';
  const qaScore = calculateLSSScore(slot.qa);
  if (qaScore !== null && slotIdx === 2) {
    const badgeClass = qaScore >= 80 ? 'score-high' : qaScore >= 60 ? 'score-med' : 'score-low';
    html += `<div class="qa-score-badge ${badgeClass}">IQF: ${qaScore}%</div>`;
  }
  if (slotIdx === 1) {
    const meta = slot.systemData?.systemsMeta;
    const _sysArr = (meta && Array.isArray(meta.systems)) ? meta.systems : [];
    const _allNvt = (_sysArr.length > 0) && _sysArr.every((x) => {
      const qa = x && x.qa && typeof x.qa === "object" ? x.qa : null;
      return !!(qa && qa.__nvt === true);
    });
    if (_allNvt) {
      html += `<div class="qa-score-badge score-med">TTF: NVT</div>`;
    } else {
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
        const badgeClass = Number.isFinite(Number(overallDerived))
          ? Number(overallDerived) >= 80 ? 'score-high' : Number(overallDerived) >= 60 ? 'score-med' : 'score-low'
          : 'score-med';
        html += `<div class="qa-score-badge ${badgeClass}">${escapeHTML(label)}</div>`;
        return html;
      }
    }
    if (!_allNvt && slot.systemData?.calculatedScore != null) {
      const sysScore = slot.systemData.calculatedScore;
      const badgeClass = sysScore >= 80 ? 'score-high' : sysScore >= 60 ? 'score-med' : 'score-low';
      html += `<div class="qa-score-badge ${badgeClass}">TTF: ${sysScore}%</div>`;
    }
  }
  return html;
}

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
  // Status checks nu op index 4 (Proces)
  const procEmoji = slotIdx === 4 && slot.processStatus ? getProcessEmoji(slot.processStatus) : '';
  const leanIcon = slotIdx === 4 && slot.processValue ? getLeanIcon(slot.processValue) : '';
  const workExpIcon = slotIdx === 4 ? getWorkExpMeta(slot?.workExp)?.icon || '' : '';
  const linkIcon = isLinked ? '🔗' : '';

  const b1 = slotIdx === 4 && slot.type ? typeIcon : '';
  const b2 = slotIdx === 4 ? leanIcon : '';
  const b3 = slotIdx === 4 ? workExpIcon : slotIdx === 2 && isLinked ? linkIcon : '';
  const b4 = slotIdx === 4 ? procEmoji : '';

  const editableAttr = isLinked ? 'contenteditable="false" data-linked="true"' : 'contenteditable="true"';

  return `
    <div class="sticky ${statusClass} ${extraStickyClass}" style="${escapeAttr(extraStickyStyle)}" data-col="${colIdx}" data-slot="${slotIdx}">
      ${routeBadgeHTML}
      <div class="sticky-grip"></div>
      <div class="badges-row">
        <div class="sticky-badge">${escapeHTML(b1)}</div>
        <div class="sticky-badge">${escapeHTML(b2)}</div>
        <div class="sticky-badge">${escapeHTML(b3)}</div>
        <div class="sticky-badge emoji-only">${escapeHTML(b4)}</div>
      </div>
      ${slotIdx === 2 && myInputId ? `<div class="id-tag">${myInputId}</div>` : ''}
      ${slotIdx === 5 && myOutputId ? `<div class="id-tag">${myOutputId}</div>` : ''}
      ${scoreBadgeHTML}
      <div class="sticky-content">
        <div class="text" ${editableAttr} spellcheck="false"></div>
      </div>
    </div>
  `;
}

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

function scheduleSyncRowHeights() {
  if (_syncRaf) cancelAnimationFrame(_syncRaf);
  _syncRaf = requestAnimationFrame(() => {
    _syncRaf = 0;
    syncRowHeightsNow();
    renderMergedOverlays(_openModalFn);
    renderGroupOverlays();
  });
}

function syncRowHeightsNow() {
  const rowHeadersEl = $('row-headers');
  const rowHeaders = rowHeadersEl?.children;
  if (!rowHeaders || !rowHeaders.length) return;
  const colsContainer = $('cols');
  const cols = colsContainer?.querySelectorAll?.('.col');
  if (!cols || !cols.length) return;
  
  // Nu 7 rijen
  for (let r = 0; r < 7; r++) {
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
  const heights = Array(7).fill(MIN_ROW_HEIGHT);
  cols.forEach((col) => {
    const slotNodes = col.querySelectorAll('.slots .slot');
    for (let r = 0; r < 7; r++) {
      const slot = slotNodes[r];
      if (!slot) continue;
      const sticky = slot.firstElementChild;
      if (!sticky) continue;
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
  for (let r = 0; r < 7; r++) heights[r] = globalMax;
  for (let r = 0; r < 7; r++) {
    const hStr = `${heights[r]}px`;
    if (rowHeaders[r]) rowHeaders[r].style.height = hStr;
    cols.forEach((col) => {
      const slotNodes = col.querySelectorAll('.slots .slot');
      if (slotNodes[r]) slotNodes[r].style.height = hStr;
    });
  }
  const gapSize = 20;
  // Offset o.b.v. 4 rijen (Bron, Sys, In, Actor)
  const processOffset = heights[0] + heights[1] + heights[2] + heights[3] + 4 * gapSize;
  colsContainer.querySelectorAll('.col-connector').forEach((c) => {
    if (!c.classList.contains('parallel-connector') && !c.classList.contains('combo-connector')) {
      const colEl = c.closest('.col-connector').previousElementSibling;
      const indentLevel = colEl ? parseInt(colEl.dataset.indentLevel || 0, 10) : 0;
      const extraPadding = indentLevel * 30;
      c.style.paddingTop = `${processOffset + extraPadding}px`;
    }
  });
}

function ensureRowHeaders() {
  const rowHeaderContainer = $('row-headers');
  if (!rowHeaderContainer || rowHeaderContainer.children.length > 0) return;
  ['Bron', 'Systeem', 'Input', 'Actor','Proces', 'Output', 'Klant'].forEach((label) => {
    const div = document.createElement('div');
    div.className = 'row-header';
    div.innerHTML = `<span class="lane-label-text">${escapeHTML(label)}</span>`;
    rowHeaderContainer.appendChild(div);
  });
}

export function renderSheetTabs() {
  const container = document.getElementById('sheet-tabs');
  if (!container) return;
  container.innerHTML = '';

  const project = state.project || state.data;
  if (!project || !Array.isArray(project.sheets)) return;

  // 1. "Alle Flows" knop
  const allBtn = document.createElement('button');
  allBtn.className = 'tab-btn special-tab';
  allBtn.innerHTML = '<span>∞</span> Alle Flows';
  if (state.project.activeSheetId === 'ALL') {
    allBtn.classList.add('active');
  }
  allBtn.onclick = () => {
    state.setActiveSheet('ALL');
  };
  container.appendChild(allBtn);

  // 2. Normale sheets
  project.sheets.forEach((sheet) => {
    const btn = document.createElement('button');
    btn.className = 'tab-btn';
    btn.textContent = sheet.title || sheet.name || 'Naamloos';
    if (state.project.activeSheetId === sheet.id && state.project.activeSheetId !== 'ALL') {
      btn.classList.add('active');
    }
    btn.onclick = () => {
      state.setActiveSheet(sheet.id);
    };
    container.appendChild(btn);
  });
}

function renderSheetSelect() {
  const select = $('sheetSelect');
  if (!select) return;
  
  const activeId = state.project.activeSheetId || state.activeSheetId;
  select.innerHTML = '';

  // 1. "Alle Flows" optie
  const allOpt = document.createElement('option');
  allOpt.value = 'ALL';
  allOpt.textContent = '∞ Alle Flows (Totaaloverzicht)';
  allOpt.selected = activeId === 'ALL';
  select.appendChild(allOpt);

  // 2. Divider
  const sep = document.createElement('option');
  sep.disabled = true;
  sep.textContent = '──────────';
  select.appendChild(sep);

  // 3. De sheets
  const project = state.project || state.data;
  (project.sheets || []).forEach((s) => {
    const opt = document.createElement('option');
    opt.value = s.id;
    opt.textContent = s.name;
    opt.selected = s.id === activeId;
    select.appendChild(opt);
  });
}

function renderHeader(activeSheet) {
  const headDisp = $('board-header-display');
  if (!headDisp) return;
  if (state.project.activeSheetId === 'ALL') {
    headDisp.textContent = "Alle Procesflows (Totaaloverzicht)";
  } else {
    headDisp.textContent = activeSheet?.name || activeSheet?.title || '';
  }
}

function attachStickyInteractions({ stickyEl, textEl, colIdx, slotIdx, openModalFn }) {
  const onDblClick = (e) => {
    if (![1, 2, 3, 4, 5].includes(slotIdx)) return;
    e.preventDefault();
    e.stopPropagation();
    const sel = window.getSelection?.();
    if (sel) sel.removeAllRanges();
    if (slotIdx === 5 || slotIdx === 1) openMergeModal(colIdx, slotIdx, openModalFn);
    else {
      openModalFn?.(colIdx, slotIdx);
      requestAnimationFrame(() => removeLegacySystemMergeUI());
    }
  };
  const focusText = (e) => {
    if (e.detail && e.detail > 1) return;
    if (e.target.closest('.sticky-grip, .qa-score-badge, .id-tag, .badges-row, .workexp-badge, .btn-col-action, .col-actions, .qmark-btn, [data-action="help"]')) {
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

function renderStats(stats) {
  const happyEl = $('countHappy');
  const neutralEl = $('countNeutral');
  const sadEl = $('countSad');
  if (happyEl) happyEl.textContent = stats.happy;
  if (neutralEl) neutralEl.textContent = stats.neutral;
  if (sadEl) sadEl.textContent = stats.sad;
}

function _formatSystemsSummaryFromMeta(meta) {
  const clean = _sanitizeSystemsMeta(meta);
  if (!clean) return '';
  const systems = clean.systems || [];
  const lineHTML = (s) => {
    const nm = String(s?.name || '').trim() || '—';
    const legacy = !!s?.legacy;
    return `<div class="sys-line"><span class="sys-name">${escapeHTML(nm)}</span>${legacy ? `<span class="legacy-tag" aria-label="Legacy">Legacy</span>` : ''}</div>`;
  };
  if (!clean.multi) {
    const s = systems[0] || { name: '', legacy: false };
    const nm = String(s.name || '').trim();
    if (!nm) return '';
    return `<div class="sys-summary">${lineHTML(s)}</div>`;
  }
  return `<div class="sys-summary">${systems.map(lineHTML).join('')}</div>`;
}

function renderMergedOverlays(openModalFn) {
  const colsContainer = $('cols');
  if (!colsContainer) return;
  if (getComputedStyle(colsContainer).position === 'static') colsContainer.style.position = 'relative';
  const activeSheet = state.activeSheet;
  if (!activeSheet) return;
  const groups = getAllMergeGroupsSanitized();
  const project = state.project || state.data;
  const _outMaps = buildGlobalOutputMaps(project || {});
  const outIdByUid = _outMaps?.outIdByUid || {};
  const outTextByUid = _outMaps?.outTextByUid || {};
  const outTextByOutId = _outMaps?.outTextByOutId || {};

  function _renderInputDisplayText(slot) {
    const eff = getEffectiveInputSlot(-1, slot);
    const bundleIds = getLinkedBundleIdsFromInputSlot(eff);
    const tokens = getLinkedSourcesFromInputSlot(eff);
    const hasTokens = tokens.some((t) => {
      const s = String(t || '').trim();
      if (!s) return false;
      if (_looksLikeOutId(s)) return true;
      return !!outIdByUid[s];
    });
    const hasBundles = bundleIds.length > 0;
    if (!hasTokens && !hasBundles) {
      return { isLinked: false, text: String(eff?.text || '') };
    }
    const parts = [];
    if (hasBundles) {
      bundleIds.forEach(bid => {
        const lbl = _getBundleLabel(project, bid);
        if (lbl) parts.push(lbl);
      });
    }
    if (hasTokens) {
      const resolved = resolveLinkedSourcesToOutAndText(tokens, outIdByUid, outTextByUid, outTextByOutId);
      const txt = _joinSemiText(resolved?.texts) || _joinSemiText(resolved?.ids) || '';
      if (txt) parts.push(txt);
    }
    return { isLinked: true, text: _joinSemiText(parts) };
  }

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
            if (g.slotIdx !== 2) state.updateStickyText(masterCol, g.slotIdx, txt.textContent);
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
        return;
      }
      if (g.slotIdx === 2) {
        const disp = _renderInputDisplayText(masterData);
        const nextText = String(disp?.text || '');
        if (textEl.textContent !== nextText) {
          textEl.textContent = nextText;
        }
        if (disp?.isLinked) {
          textEl.setAttribute('contenteditable', 'false');
          textEl.setAttribute('data-linked', 'true');
        } else {
          textEl.setAttribute('contenteditable', 'true');
          textEl.removeAttribute('data-linked');
        }
        return;
      }
      if (textEl.textContent !== baseText) {
        textEl.textContent = baseText;
      }
      textEl.setAttribute('contenteditable', 'true');
      textEl.removeAttribute('data-linked');
    }
    
    // Gate op index 5
    if (g.slotIdx === 5 && stickyEl) {
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

function renderColumnsOnly(openModalFn, targetContainer = $('cols'), append = false, overrideSheet = null) {
  // Use override if provided, else fallback to state.activeSheet
  const activeSheet = overrideSheet || state.activeSheet;
  if (!activeSheet) return;
  
  if (!append && targetContainer) targetContainer.innerHTML = '';
  if (!targetContainer) return;

  ensureMergeGroupsLoaded();
  
  const variantLetterMap = computeVariantLetterMap(activeSheet);
  const project = state.project || state.data;
  const { outIdByUid, outTextByUid, outTextByOutId } = buildGlobalOutputMaps(project);
  try {
    const all = typeof state.getAllOutputs === 'function' ? state.getAllOutputs() : {};
    Object.keys(all || {}).forEach((k) => {
      if (!outTextByOutId[k] && all[k]) outTextByOutId[k] = all[k];
    });
  } catch {}
  
  const offsets = computeCountersBeforeActiveSheet(project, activeSheet.id, outIdByUid);
  let localInCounter = 0;
  let localOutCounter = 0;
  const stats = { happy: 0, neutral: 0, sad: 0 };
  const frag = document.createDocumentFragment();
  
  activeSheet.columns.forEach((col, colIdx) => {
    if (col.isVisible === false) return;
    let myInputId = '';
    let myOutputId = '';
    const inputSlot = col.slots?.[2];
    const outputSlot = col.slots?.[5];
    
    const effectiveInputSlot = getEffectiveInputSlot(colIdx, inputSlot);
    let bundleIdsForInput = getLinkedBundleIdsFromInputSlot(effectiveInputSlot);
    let tokens = getLinkedSourcesFromInputSlot(effectiveInputSlot);
    try {
      const masterCol = isMergedSlave(colIdx, 2) ? getMergedMasterColIdx(colIdx, 2) : colIdx;
      const g = getMergeGroup(masterCol, 2) || getMergeGroup(colIdx, 2);
      if (g && Array.isArray(g.cols) && g.cols.length) {
        const seenB = new Set((bundleIdsForInput || []).map((x) => String(x)));
        const allB = [...(bundleIdsForInput || [])];
        const seenT = new Set();
        const allT = [...(tokens || [])];
        for (const t of allT) {
          const k = (typeof t === "string") ? t : JSON.stringify(t);
          if (k) seenT.add(k);
        }
        for (const ci of g.cols) {
          const s = activeSheet.columns?.[ci]?.slots?.[2];
          for (const bid of (getLinkedBundleIdsFromInputSlot(s) || [])) {
            const k = String(bid);
            if (k && !seenB.has(k)) { seenB.add(k); allB.push(bid); }
          }
          for (const t of (getLinkedSourcesFromInputSlot(s) || [])) {
            const k = (typeof t === "string") ? t : JSON.stringify(t);
            if (k && !seenT.has(k)) { seenT.add(k); allT.push(t); }
          }
        }
        bundleIdsForInput = allB;
        tokens = allT;
      }
    } catch {}
    const bundleLabelsForInput = (bundleIdsForInput || []).map((bid) => _getBundleLabel(project, bid));
    const resolved = resolveLinkedSourcesToOutAndText(tokens, outIdByUid, outTextByUid, outTextByOutId);
    if (bundleLabelsForInput.length) {
      myInputId = _joinSemiText(bundleLabelsForInput);
    } else if (resolved.ids.length) {
      myInputId = _joinSemiText(resolved.ids);
    } else if (inputSlot?.text?.trim() && !isMergedSlave(colIdx, 2)) {
      localInCounter += 1;
      myInputId = `IN${offsets.inStart + localInCounter}`;
    }
    
    if (outputSlot?.text?.trim() && !isMergedSlave(colIdx, 5)) {
      localOutCounter += 1;
      myOutputId = `OUT${offsets.outStart + localOutCounter}`;
    }
    
    const colEl = document.createElement('div');
    colEl.className = `col ${col.isParallel ? 'is-parallel' : ''} ${col.isVariant ? 'is-variant' : ''} ${col.isGroup ? 'is-group' : ''}`;
    colEl.dataset.idx = colIdx;
    const depth = getDependencyDepth(colIdx);
    if (depth > 0) colEl.dataset.depth = depth;
    if (depth > 0) colEl.style.transformOrigin = 'top center';
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
      ${colIdx > 0 ? `<button class="btn-col-action btn-parallel ${col.isParallel ? 'active' : ''}" data-action="parallel" type="button">∥</button>` : ''}
      ${colIdx > 0 ? `<button class="btn-col-action btn-variant ${col.isVariant ? 'active' : ''}" data-action="variant" type="button">🔀</button>` : ''}
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
      if (slotIdx === 4) {
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
      if (slotIdx === 4 && slot.processStatus) statusClass = `status-${slot.processStatus.toLowerCase()}`;
      
      let typeIcon = '📝';
      if (slot.type === 'Afspraak') typeIcon = '📅';
      
      let extraStickyClass = '';
      let extraStickyStyle = '';
      if (getMergeGroup(colIdx, slotIdx)) {
        extraStickyClass = 'merged-source';
        extraStickyStyle = 'visibility:hidden; pointer-events:none;';
      }
      const routeBadgeHTML = slotIdx === 0 ? buildSupplierTopBadgesHTML({ routeLabel, isConditional: !!col.isConditional }) : '';
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
      
      if (stickyEl && Number(stickyEl?.dataset?.slot) === 5) {
        const _colIdx = Number(stickyEl?.dataset?.col);
        const _slot = activeSheet.columns?.[_colIdx]?.slots?.[5];
        const gate = _sanitizeGate(_slot?.outputData?.gate) || _sanitizeGate(_slot?.gate);
        const passLabel = (typeof getPassLabelForGroup === 'function')
          ? getPassLabelForGroup({ slotIdx: 5, cols: [_colIdx], master: _colIdx })
          : ((activeSheet.columns?.[_colIdx + 1]) ? getProcessLabel(_colIdx + 1) : '—');
        let failLabel = '—';
        if (gate?.enabled && gate.failTargetColIdx != null) {
          const idx = Number(gate.failTargetColIdx);
          if (Number.isFinite(idx)) failLabel = getProcessLabel(idx);
        }
        applyGateToSticky(stickyEl, gate, passLabel, failLabel);
      }
      if (textEl) textEl.textContent = displayText;
      const isMergedSource = (slotIdx === 2) ? isMergedSlave(colIdx, 2) : !!getMergeGroup(colIdx, slotIdx);
      if (!isMergedSource) attachStickyInteractions({ stickyEl, textEl, colIdx, slotIdx, openModalFn });
      if (!isLinked && textEl && !isMergedSource) {
        textEl.addEventListener(
          'input',
          () => {
            if (slotIdx === 2 && isMergedSlave(colIdx, 2)) {
              const masterCol = getMergedMasterColIdx(colIdx, 2);
              if (masterCol != null) state.updateStickyText(masterCol, slotIdx, textEl.textContent);
            } else {
              state.updateStickyText(colIdx, slotIdx, textEl.textContent);
            } scheduleSyncRowHeights();
          },
          { passive: true }
        );
        textEl.addEventListener(
          'blur',
          () => {
            if (slotIdx === 2 && isMergedSlave(colIdx, 2)) {
              const masterCol = getMergedMasterColIdx(colIdx, 2);
              if (masterCol != null) state.updateStickyText(masterCol, slotIdx, textEl.textContent);
            } else {
              state.updateStickyText(colIdx, slotIdx, textEl.textContent);
            } scheduleSyncRowHeights();
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
  
  targetContainer.appendChild(frag); 
  renderStats(stats);
  scheduleSyncRowHeights();
  requestAnimationFrame(() => renderGroupOverlays());
}

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
      slotEl.innerHTML = renderTextWithEmojiSpan(_joinSemiText(parts));
      return true;
    }
  }
  slotEl.innerHTML = renderTextWithEmojiSpan(slot.text ?? "");
  return true;
}

window.addEventListener("ssipoc:openMergeModal", (ev) => {
  const d = (ev && ev.detail) ? ev.detail : {};
  const colIdx = Number(d.colIdx);
  const slotIdx = Number(d.slotIdx);
  if (!Number.isFinite(colIdx) || !Number.isFinite(slotIdx)) return;
  openMergeModal(colIdx, slotIdx, _openModalFn);
});

export function renderBoard(openModalFn) {
  _openModalFn = openModalFn || _openModalFn;
  
  ensureMergeGroupsLoaded();
  renderSheetTabs(); 
  renderSheetSelect();
  
  const colsContainer = $('cols');
  if (!colsContainer) return;
  colsContainer.innerHTML = ''; 

  const project = state.project || state.data;

  // Use project.activeSheetId as source of truth for ALL mode
  if (state.project && state.project.activeSheetId === 'ALL') {
    document.body.classList.add('view-mode-all');
    renderHeader(null); 
    ensureRowHeaders();

    const originalId = state.project.activeSheetId;

    // Loop through all sheets and temporarily swap active ID for helpers to work
    project.sheets.forEach((sheet, index) => {
      const separator = document.createElement('div');
      separator.className = 'flow-separator';
      separator.innerHTML = `
        <div class="flow-separator-line"></div>
        <div class="flow-separator-title">${escapeHTML(sheet.title || sheet.name || 'Proces')}</div>
      `;
      colsContainer.appendChild(separator);

      // SWAP CONTEXT so helpers use correct sheet
      state.project.activeSheetId = sheet.id; 
      
      // Render with explicit override
      renderColumnsOnly(_openModalFn, colsContainer, true, sheet);
    });

    // RESTORE CONTEXT
    state.project.activeSheetId = originalId;

  } else {
    document.body.classList.remove('view-mode-all');
    if (state.activeSheet) {
        renderHeader(state.activeSheet);
        ensureRowHeaders();
        renderColumnsOnly(_openModalFn, colsContainer, false);
    }
  }
}

export function applyStateUpdate(meta, openModalFn) {
  _openModalFn = openModalFn || _openModalFn;
  const reason = meta?.reason || 'full';
  
  if (state.project && state.project.activeSheetId !== 'ALL' && reason === 'text' && Number.isFinite(meta?.colIdx) && Number.isFinite(meta?.slotIdx)) {
    const ok = updateSingleText(meta.colIdx, meta.slotIdx);
    if (ok) return;
  }
  
  if (reason === 'title') return;
  
  if (reason === 'sheet' || reason === 'sheets') {
    renderBoard(_openModalFn);
    return;
  }
  
  renderBoard(_openModalFn);
}

export function setupDelegatedEvents() {
  if (_delegatedBound) return;
  _delegatedBound = true;
  const act = (e) => {
    const now = performance.now();
    const helpBtn = e.target.closest?.('[data-action="help"], .qmark-btn');
    if (helpBtn) {
      if ((e.type === 'mousedown' || e.type === 'click') && now - _lastPointerDownTs < 250) return;
      if (e.type === 'pointerdown' || e.type === 'touchstart') _lastPointerDownTs = now;
      e.preventDefault();
      e.stopPropagation();
      const helpKey = String(helpBtn.dataset.helpKey || helpBtn.dataset.criterionKey || helpBtn.dataset.key || '').trim() || null;
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
    if ((e.type === 'mousedown' || e.type === 'click') && now - _lastPointerDownTs < 250) return;
    if (e.type === 'pointerdown' || e.type === 'touchstart') _lastPointerDownTs = now;
    e.preventDefault();
    e.stopPropagation();
    const colEl = btn.closest('.col');
    if (!colEl) return;
    const idx = parseInt(colEl.dataset.idx, 10);
    if (!Number.isFinite(idx)) return;
    
    // Check project.activeSheetId
    if (state.project && state.project.activeSheetId === 'ALL') {
        alert("Ga naar een specifieke procesflow om de structuur aan te passen.");
        return;
    }

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
      case 'question': {
        if (typeof state.toggleQuestion === 'function') {
          state.toggleQuestion(idx);
          break;
        }
        if (typeof state.toggleQuestionColumn === 'function') {
          state.toggleQuestionColumn(idx);
          break;
        }
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

(function(){
  function _esc(s){
    return String(s)
      .replace(/&/g,"&amp;")
      .replace(/</g,"&lt;")
      .replace(/>/g,"&gt;")
      .replace(/"/g,"&quot;")
      .replace(/'/g,"&#39;");
  }
  function _isEmojiChar(ch){
    try { 
      return new RegExp("\\p{Extended_Pictographic}", "u").test(ch);
    } catch (e) {
      return /[\uD83C-\uDBFF][\uDC00-\uDFFF]/.test(ch);
    }
  }
  function _wrapInlineEmoji(el){
    if (!el) return;
    if (el === document.activeElement) return;
    const t = el.textContent ?? "";
    if (!t) return;
    if (el.dataset && el.dataset.emojiSrc === t) return;
    let out = "";
    for (const ch of Array.from(t)){
      if (_isEmojiChar(ch)){
        out += `<span class="emoji-inline">${_esc(ch)}</span>`;
      } else {
        out += _esc(ch);
      }
    }
    el.innerHTML = out;
    if (el.dataset) el.dataset.emojiSrc = t;
  }
  function _applyAll(){
    document.querySelectorAll(".sticky .text").forEach(_wrapInlineEmoji);
  }
  window.addEventListener("DOMContentLoaded", () => {
    _applyAll();
    const target = document.getElementById("board") || document.body;
    const obs = new MutationObserver(() => { _applyAll(); });
    obs.observe(target, { childList: true, subtree: true, characterData: true });
  });
})();