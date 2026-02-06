// io.js (FINAL VERSION: Fix Slot Indices 4->5 for Output & ID Resolution)
// + FIX: Output data wordt nu correct uit Slot 5 gelezen (was 4)
// + FIX: Proces data wordt nu correct uit Slot 4 gelezen (was 3)
// + FIX: Actor data wordt nu correct uit Slot 3 gelezen
// + FIX: Mapping van Input-links naar Output-tekst werkt weer (geen ruwe ID's meer)
// + INCLUDES: Eerdere fixes (Headers, PNG, PDF, Github SHA)

import { state } from './state.js';
import { Toast } from './toast.js';
import { IO_CRITERIA, LEAN_VALUES, PROCESS_STATUSES } from './config.js';

const MERGE_LS_PREFIX = 'ssipoc.mergeGroups.v2';
let currentLoadedSha = null;

/* ==========================================================================
   EXPORT TYPOGRAPHY (exact match with your CSS)
   ========================================================================== */
const EXPORT_ROW_FONT_FAMILY = '"JetBrains Mono", monospace';
const EXPORT_ROW_FONT_SIZE_PX = 24;
const EXPORT_ROW_FONT_WEIGHT = 800;
const EXPORT_ROW_LETTER_SPACING = '0.08em';
const EXPORT_ROW_LINE_HEIGHT = 1;
const EXPORT_ROW_FILL = 'rgba(255,255,255,0.95)';
const EXPORT_ROW_SHADOW = 'drop-shadow(0px 2px 10px rgba(0,0,0,0.85))';
const EXPORT_ROW_UPPERCASE = true;
const EXPORT_ROW_RIGHT_PAD_PX = 6;

/* ==========================================================================
   HELPER: Replace row headers with SVG using exact typography
   ========================================================================== */
function replaceRowHeadersWithExactSVG(doc, headerWidthPx) {
  const rh = doc.getElementById('row-headers');
  if (!rh) return;

  const rows = Array.from(rh.querySelectorAll('.row-header'));
  const labels = rows.map((r) => {
    const node = r.querySelector('.lane-label-text') || r.querySelector('span') || r;
    const raw = String(node?.textContent || '').trim();
    return EXPORT_ROW_UPPERCASE ? raw.toUpperCase() : raw;
  });
  const heights = rows.map((r) =>
    r.getBoundingClientRect().height > 0 ? r.getBoundingClientRect().height : 160
  );

  rh.innerHTML = '';
  rh.style.width = `${headerWidthPx}px`;
  rh.style.minWidth = `${headerWidthPx}px`;
  rh.style.marginRight = '0px';
  rh.style.overflow = 'visible';
  rh.style.position = 'relative';
  rh.style.zIndex = '5';

  const totalH = Math.max(1, Math.round(heights.reduce((a, b) => a + b, 0)));
  const svgNS = 'http://www.w3.org/2000/svg';
  const svg = doc.createElementNS(svgNS, 'svg');
  svg.setAttribute('width', String(headerWidthPx));
  svg.setAttribute('height', String(totalH));
  svg.setAttribute('viewBox', `0 0 ${headerWidthPx} ${totalH}`);
  svg.style.overflow = 'visible';
  svg.style.display = 'block';

  const anchorX = Math.max(10, headerWidthPx - EXPORT_ROW_RIGHT_PAD_PX);

  let y = 0;
  for (let i = 0; i < labels.length; i++) {
    const h = Math.max(40, Math.round(heights[i] || 160));
    const cy = y + h / 2;

    const g = doc.createElementNS(svgNS, 'g');
    g.setAttribute('transform', `rotate(-90 ${anchorX} ${cy})`);

    const t = doc.createElementNS(svgNS, 'text');
    t.setAttribute('x', String(anchorX));
    t.setAttribute('y', String(cy));
    t.setAttribute('text-anchor', 'middle');
    t.setAttribute('dominant-baseline', 'middle');

    t.setAttribute('fill', EXPORT_ROW_FILL);
    t.setAttribute('font-family', EXPORT_ROW_FONT_FAMILY);
    t.setAttribute('font-size', String(EXPORT_ROW_FONT_SIZE_PX));
    t.setAttribute('font-weight', String(EXPORT_ROW_FONT_WEIGHT));
    t.style.letterSpacing = EXPORT_ROW_LETTER_SPACING;
    t.setAttribute('letter-spacing', EXPORT_ROW_LETTER_SPACING);
    t.style.filter = EXPORT_ROW_SHADOW;
    t.style.lineHeight = String(EXPORT_ROW_LINE_HEIGHT);

    t.textContent = labels[i];

    g.appendChild(t);
    svg.appendChild(g);

    y += h;
  }

  rh.appendChild(svg);
}

/* ==========================================================================
   HELPER: Load jsPDF dynamically
   ========================================================================== */
async function loadJsPDF() {
  if (window.jspdf) return window.jspdf;
  return new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js";
    script.onload = () => resolve(window.jspdf);
    script.onerror = () => reject(new Error("Kon jsPDF niet laden"));
    document.head.appendChild(script);
  });
}

/* ==========================================================================
   CSV helpers & File names
   ========================================================================== */
function toCsvField(text) {
  if (text === null || text === undefined) return '""';
  const str = String(text);
  return `"${str.replace(/"/g, '""').replace(/\r?\n/g, ' ')}"`;
}

function joinSemi(arr) {
  const a = Array.isArray(arr) ? arr : [];
  const cleaned = a.map((x) => String(x ?? '').trim()).filter((x) => x !== '');
  return cleaned.join('; ');
}

function getFileName(ext) {
  const title = (state.data || state.project)?.projectTitle || 'sipoc_project';
  const safeTitle = String(title).replace(/[^a-z0-9]/gi, '_').toLowerCase();
  const now = new Date();
  const dateStr = now.toISOString().slice(0, 10);
  const timeStr = now.toTimeString().slice(0, 8).replace(/:/g, '');
  return `${safeTitle}_${dateStr}_${timeStr}.${ext}`;
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function downloadCanvas(canvas) {
  const a = document.createElement('a');
  a.href = canvas.toDataURL('image/png');
  a.download = getFileName('png');
  a.click();
}

/* ==========================================================================
   Merge groups logic & Sanitization
   ========================================================================== */
function mergeKeyForSheet(project, sheet) {
  const pid = project?.id || project?.name || project?.projectTitle || 'project';
  const sid = sheet?.id || sheet?.name || 'sheet';
  return `${MERGE_LS_PREFIX}:${pid}:${sid}`;
}

function loadMergeGroupsRaw(project, sheet) {
  try {
    const raw = localStorage.getItem(mergeKeyForSheet(project, sheet));
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch { return []; }
}

function snapshotMergeGroupsIntoProject(project) {
  const p = project || state.data;
  if (!p || !Array.isArray(p.sheets)) return;
  (p.sheets || []).forEach((sheet) => {
    const raw = loadMergeGroupsRaw(p, sheet);
    if (Array.isArray(raw) && raw.length) sheet.mergeGroupsV2 = raw;
    else if ('mergeGroupsV2' in sheet) delete sheet.mergeGroupsV2;
  });
}

function restoreMergeGroupsToLocalStorage(project) {
  const p = project || state.data;
  if (!p || !Array.isArray(p.sheets)) return;
  (p.sheets || []).forEach((sheet) => {
    const raw = Array.isArray(sheet?.mergeGroupsV2) ? sheet.mergeGroupsV2 : null;
    if (!raw) return;
    try {
      localStorage.setItem(mergeKeyForSheet(p, sheet), JSON.stringify(raw));
    } catch {}
  });
}

function isContiguousZeroBased(cols) {
  if (!Array.isArray(cols) || cols.length < 2) return false;
  const s = [...new Set(cols)].sort((a, b) => a - b);
  return s.length === s[s.length - 1] - s[0] + 1;
}
function sanitizeGate(gate) {
  if (!gate || typeof gate !== 'object') return null;
  return {
    enabled: !!gate.enabled,
    failTargetColIdx: Number.isFinite(Number(gate.failTargetColIdx)) ? Number(gate.failTargetColIdx) : null
  };
}
function finalizeGate(gate) {
  const g = sanitizeGate(gate);
  if (!g?.enabled || g.failTargetColIdx == null || !Number.isFinite(Number(g.failTargetColIdx))) return null;
  return { enabled: true, failTargetColIdx: Number(g.failTargetColIdx) };
}

// UPDATE: Aangepast naar slot indices [1, 2, 5] en Gate op 5
function sanitizeMergeGroupForSheet(sheet, g) {
  const n = sheet?.columns?.length ?? 0;
  if (!n) return null;
  const slotIdx = Number(g?.slotIdx);
  // Allow 1 (Sys), 2 (Input), 5 (Output)
  if (![1, 2, 5].includes(slotIdx)) return null;

  const cols = Array.isArray(g?.cols) ? g.cols.map((x) => Number(x)).filter(Number.isFinite) : [];
  const uniq = [...new Set(cols)].filter((c) => c >= 0 && c < n);
  if (uniq.length < 2) return null;
  if (!isContiguousZeroBased(uniq)) return null;
  let master = Number(g?.master);
  if (!Number.isFinite(master) || !uniq.includes(master)) master = uniq[0];
  
  // Gate zit nu op slot 5
  const gate = slotIdx === 5 ? finalizeGate(g?.gate) : null;
  const systemsMeta = slotIdx === 1 && g?.systemsMeta && typeof g.systemsMeta === 'object' ? g.systemsMeta : null;
  return { slotIdx, cols: uniq.sort((a, b) => a - b), master, gate, systemsMeta };
}

function getMergeGroupsSanitized(project, sheet) {
  return loadMergeGroupsRaw(project, sheet)
    .map((g) => sanitizeMergeGroupForSheet(sheet, g))
    .filter(Boolean);
}
function getMergeGroupForCell(groups, colIdx, slotIdx) {
  return ((groups || []).find((x) => x.slotIdx === slotIdx && Array.isArray(x.cols) && x.cols.includes(colIdx)) || null);
}
function isMergedSlaveInSheet(groups, colIdx, slotIdx) {
  const g = getMergeGroupForCell(groups, colIdx, slotIdx);
  return !!g && colIdx !== g.master;
}
function getNextVisibleColIdx(sheet, fromIdx) {
  const n = sheet?.columns?.length ?? 0;
  for (let i = fromIdx + 1; i < n; i++) { if (sheet.columns[i]?.isVisible !== false) return i; }
  return null;
}
// Proces label zit op slot 4 (was 3)
function getProcessLabel(sheet, colIdx) {
  const t = sheet?.columns?.[colIdx]?.slots?.[4]?.text;
  return String(t ?? '').trim() || `Kolom ${Number(colIdx) + 1}`;
}
function getPassTargetFromGroup(sheet, group) {
  if (!group?.cols?.length) return '';
  const maxCol = Math.max(...group.cols);
  const nextIdx = getNextVisibleColIdx(sheet, maxCol);
  return nextIdx == null ? 'Einde proces' : getProcessLabel(sheet, nextIdx);
}
function getFailTargetFromGate(sheet, gate) {
  const g = finalizeGate(gate);
  return g ? getProcessLabel(sheet, g.failTargetColIdx) : '';
}
function getSheetRoutePrefix(project, sheet) {
  const p = project || state.data;
  const sheets = Array.isArray(p?.sheets) ? p.sheets : [];
  const idx = sheets.findIndex((s) => s?.id === sheet?.id);
  return `RF${idx >= 0 ? idx + 1 : 1}`;
}
function getScopedRouteLabel(project, sheet, label) {
  const base = String(label || '').trim();
  return base ? `${getSheetRoutePrefix(project, sheet)}-${base}` : '';
}
function toLetter(i0) {
  const n = Number(i0);
  if (!Number.isFinite(n) || n < 0) return 'A';
  const base = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  return base[n] || `R${n + 1}`;
}
function computeVariantLetterMap(sheet) {
  const map = {};
  if (!sheet?.columns?.length) return map;
  const colGroups = {};
  if (Array.isArray(sheet.variantGroups)) {
    sheet.variantGroups.forEach((vg) => {
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
      const myLabel = prefix ? `${prefix}.${i + 1}` : toLetter(i);
      map[childIdx] = myLabel;
      assignLabels(childIdx, myLabel);
    });
  }
  const allChildren = new Set();
  Object.values(colGroups).forEach((list) => list.forEach((c) => allChildren.add(c)));
  const rootParents = Object.keys(colGroups).map(Number).filter((p) => !allChildren.has(p));
  rootParents.forEach((root) => assignLabels(root, ''));
  let legacyCounter = 0;
  sheet.columns.forEach((col, i) => {
    if (col.isVariant && !map[i]) map[i] = toLetter(legacyCounter++);
  });
  return map;
}
function getFollowupRouteLabel(sheet, colIdx) {
  const col = sheet.columns?.[colIdx];
  if (!col || (!col.routeLabel && !col.isVariant)) return null;
  const manualRoute = String(col.routeLabel || '').trim();
  if (!manualRoute || !!col.isVariant) return null;
  let count = 1;
  for (let i = 0; i < colIdx; i++) {
    const c = sheet.columns?.[i];
    if (c && !c.isVariant && String(c.routeLabel || '').trim() === manualRoute) count++;
  }
  return `${manualRoute}.${count}`;
}

// ... [Output IDs / Bundles / Systems] ...
function makeId(prefix = 'id') {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') return `${prefix}_${crypto.randomUUID()}`;
  return `${prefix}_${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

// UPDATE: Output zit op slot 5
function ensureOutputUids(project) {
  const p = project || state.data;
  if (!p?.sheets?.length) return;
  (p.sheets || []).forEach((sheet) => {
    const groups = getMergeGroupsSanitized(p, sheet);
    (sheet?.columns || []).forEach((col, colIdx) => {
      if (col?.isVisible === false) return;
      const outSlot = col?.slots?.[5]; // was 4
      if (!String(outSlot?.text ?? '').trim() || isMergedSlaveInSheet(groups, colIdx, 5)) return; // was 4
      if (!outSlot.outputUid || !String(outSlot.outputUid).trim()) outSlot.outputUid = makeId('out');
    });
  });
}

// UPDATE: Output zit op slot 5
function buildGlobalOutputMaps(project) {
  const outIdByUid = {}, outTextByUid = {}, outTextByOutId = {};
  let outCounter = 0;
  (project?.sheets || []).forEach((sheet) => {
    const groups = getMergeGroupsSanitized(project, sheet);
    (sheet?.columns || []).forEach((col, colIdx) => {
      if (col?.isVisible === false) return;
      const outSlot = col?.slots?.[5]; // was 4
      const outText = String(outSlot?.text ?? '').trim();
      if (!outText || isMergedSlaveInSheet(groups, colIdx, 5)) return; // was 4
      if (!outSlot.outputUid) outSlot.outputUid = makeId('out');
      outCounter += 1;
      const outId = `OUT${outCounter}`;
      outIdByUid[outSlot.outputUid] = outId;
      outTextByUid[outSlot.outputUid] = outText;
      outTextByOutId[outId] = outText;
    });
  });
  return { outIdByUid, outTextByUid, outTextByOutId };
}
function _looksLikeOutId(v) { return !!String(v || '').trim() && /^OUT\d+$/.test(v); }
function normalizeLinkedSources(inputSlot) {
  const out = [];
  const uidsArr = Array.isArray(inputSlot?.linkedSourceUids) ? inputSlot.linkedSourceUids : [];
  uidsArr.forEach(u => String(u).trim() && out.push({ kind: 'uid', value: String(u).trim() }));
  const idsArr = Array.isArray(inputSlot?.linkedSourceIds) ? inputSlot.linkedSourceIds : [];
  idsArr.forEach(id => String(id).trim() && out.push({ kind: 'id', value: String(id).trim() }));
  if (inputSlot?.linkedSourceUid) out.push({ kind: 'uid', value: String(inputSlot.linkedSourceUid).trim() });
  if (inputSlot?.linkedSourceId) out.push({ kind: 'id', value: String(inputSlot.linkedSourceId).trim() });
  const seen = new Set(); const uniq = [];
  for (const x of out) { if (!seen.has(`${x.kind}:${x.value}`)) { seen.add(`${x.kind}:${x.value}`); uniq.push(x); } }
  return uniq;
}
function resolveLinkedSourcesToOutPairs(sources, outIdByUid, outTextByUid, outTextByOutId) {
  const ids = [], texts = [];
  (sources || []).forEach((src) => {
    if (!src?.value) return;
    if (src.kind === 'id') {
      const raw = String(src.value).trim();
      if (!_looksLikeOutId(raw)) return;
      ids.push(raw); texts.push(String(outTextByOutId?.[raw] ?? '').trim() || raw);
    } else if (src.kind === 'uid') {
      const uid = String(src.value).trim();
      const outId = String(outIdByUid?.[uid] ?? '').trim() || uid;
      ids.push(outId); texts.push(String(outTextByUid?.[uid] ?? '').trim() || outId);
    }
  });
  return { ids, texts };
}
function normalizeLinkedBundles(inputSlot) {
  const out = [];
  const idsArr = Array.isArray(inputSlot?.linkedBundleIds) ? inputSlot.linkedBundleIds : [];
  idsArr.forEach(id => String(id).trim() && out.push(String(id).trim()));
  if (inputSlot?.linkedBundleId) out.push(String(inputSlot.linkedBundleId).trim());
  return [...new Set(out)];
}
function buildBundleMaps(project, outIdByUid, outTextByUid, outTextByOutId) {
  buildGlobalOutputMaps(project);
  const nameById = {}, outIdsById = {}, outTextsById = {};
  (Array.isArray(project?.outputBundles) ? project.outputBundles : []).forEach((b) => {
    const id = String(b?.id ?? '').trim();
    if (!id) return;
    nameById[id] = String(b.name ?? '').trim() || id;
    const ids = [], texts = [];
    (Array.isArray(b.outIds) ? b.outIds : []).forEach((oid) => {
      const outId = String(oid ?? '').trim();
      if (outId) { ids.push(outId); texts.push(String(outTextByOutId?.[outId] ?? '').trim() || outId); }
    });
    (Array.isArray(b.outputUids) ? b.outputUids : []).forEach((u) => {
      const uid = String(u ?? '').trim();
      if (uid) {
        const oid = String(outIdByUid?.[uid] ?? '').trim() || uid;
        ids.push(oid); texts.push(String(outTextByUid?.[uid] ?? '').trim() || oid);
      }
    });
    outIdsById[id] = [...new Set(ids)]; outTextsById[id] = texts;
  });
  return { nameById, outIdsById, outTextsById };
}
function resolveBundleIdsToLists(bundleIds, bundleMaps) {
  const bundleNames = [], memberOutIds = [], memberOutTexts = [];
  (Array.isArray(bundleIds) ? bundleIds : []).forEach((bid) => {
    const id = String(bid ?? '').trim();
    if (!id) return;
    bundleNames.push(String(bundleMaps?.nameById?.[id] ?? '').trim() || id);
    (bundleMaps?.outIdsById?.[id] || []).forEach(x => memberOutIds.push(String(x).trim()));
    (bundleMaps?.outTextsById?.[id] || []).forEach(x => memberOutTexts.push(String(x).trim()));
  });
  return {
    bundleNames: bundleNames.filter(Boolean),
    memberOutIds: memberOutIds.filter(Boolean),
    memberOutTexts: memberOutTexts.filter(Boolean)
  };
}

// ... [System Fit / IQF / etc.] ...
const SYSFIT_Q = [
  { id: 'q1', title: 'Workarounds?', type: 'freq' }, { id: 'q2', title: 'Remmend?', type: 'freq' },
  { id: 'q3', title: 'Dubbel?', type: 'freq' }, { id: 'q4', title: 'Fouten?', type: 'freq' }, { id: 'q5', title: 'Uitval?', type: 'impact' }
];
const SYSFIT_OPTS = {
  freq: [{ key: 'NEVER', label: '(Bijna) nooit', score: 1 }, { key: 'SOMETIMES', label: 'Soms', score: 0.66 }, { key: 'OFTEN', label: 'Vaak', score: 0.33 }, { key: 'ALWAYS', label: '(Bijna) altijd', score: 0 }],
  impact: [{ key: 'SAFE', label: 'Veilig', score: 1 }, { key: 'DELAY', label: 'Vertraging', score: 0.66 }, { key: 'RISK', label: 'Groot risico', score: 0.33 }, { key: 'STOP', label: 'Stilstand', score: 0 }]
};
function hasValidSystems(meta) { return meta && Array.isArray(meta.systems) && meta.systems.some(s => s && String(s.name || '').trim() !== ''); }
function buildSystemsMetaFallbackFromSlot(sysSlot) {
  const slot = sysSlot || {};
  if (slot.systemData?.systemsMeta && hasValidSystems(slot.systemData.systemsMeta)) return slot.systemData.systemsMeta;
  const name = String(slot.systemData?.systemName || slot.text || '').trim();
  if (name) return { multi: false, systems: [{ name, legacy: false, future: '', qa: {}, score: null }] };
  return null;
}
function sanitizeSystemsMeta(meta) {
  if (!meta || typeof meta !== 'object') return null;
  const systems = (Array.isArray(meta.systems) ? meta.systems : []).map(s => {
    if (!s) return null;
    return {
      name: String(s.name ?? '').trim(),
      legacy: !!s.legacy,
      future: String(s.future ?? '').trim(),
      qa: { ...s.qa },
      score: Number.isFinite(Number(s.score)) ? Number(s.score) : null
    };
  }).filter(Boolean);
  if (systems.length === 0) systems.push({ name: '', legacy: false, future: '', qa: {}, score: null });
  return { multi: !!meta.multi || systems.length > 1, systems };
}
function computeTTFSystemScore(sys) {
  const qa = sys?.qa || {};
  if (qa.__nvt === true) return null;
  let sum = 0, nAns = 0;
  SYSFIT_Q.forEach(q => {
    const k = qa[q.id];
    const o = (SYSFIT_OPTS[q.type] || []).find(opt => opt.key === k);
    if (o) { sum += o.score; nAns++; }
  });
  return nAns === 0 ? null : Math.round((sum / nAns) * 100);
}
function getSysAnswerLabel(sys, qid) {
  if (sys?.qa?.__nvt) return '';
  const k = sys?.qa?.[qid], q = SYSFIT_Q.find(x => x.id === qid);
  return (SYSFIT_OPTS[q?.type] || []).find(o => o.key === k)?.label || '';
}
function getSysNote(sys, qid) { return sys?.qa?.__nvt ? '' : String(sys?.qa?.[qid + '_note'] || '').trim(); }
function systemsToLists(meta) {
  const clean = sanitizeSystemsMeta(meta);
  const systems = (clean?.systems || []).filter(s => s.name);
  if (!systems.length && clean?.systems?.length) return { systemNames: '', systemsCount: 1, ttfScores: '' }; // empty fallback
  const mapVal = (fn) => joinSemi(systems.map(fn));
  return {
    systemNames: mapVal(s => s.name),
    legacySystems: mapVal(s => s.legacy ? s.name : ''),
    targetSystems: mapVal(s => s.legacy ? s.future : ''),
    ttfScores: mapVal(s => s.qa?.__nvt ? 'NVT' : (s.score != null ? `${s.score}%` : (computeTTFSystemScore(s) != null ? `${computeTTFSystemScore(s)}%` : ''))),
    systemWorkarounds: mapVal(s => getSysAnswerLabel(s, 'q1')),
    systemWorkaroundsNotes: mapVal(s => getSysNote(s, 'q1')),
    belemmering: mapVal(s => getSysAnswerLabel(s, 'q2')),
    belemmeringNotes: mapVal(s => getSysNote(s, 'q2')),
    dubbelRegistreren: mapVal(s => getSysAnswerLabel(s, 'q3')),
    dubbelRegistrerenNotes: mapVal(s => getSysNote(s, 'q3')),
    foutgevoeligheid: mapVal(s => getSysAnswerLabel(s, 'q4')),
    foutgevoeligheidNotes: mapVal(s => getSysNote(s, 'q4')),
    gevolgUitval: mapVal(s => getSysAnswerLabel(s, 'q5')),
    gevolgUitvalNotes: mapVal(s => getSysNote(s, 'q5')),
    systemsCount: Math.max(1, systems.length)
  };
}

function calculateIQFScore(qa) {
  if (!qa) return null;
  let t = 0, e = 0;
  IO_CRITERIA.forEach(c => {
    const v = qa?.[c.key]?.result;
    if (['GOOD', 'POOR', 'MODERATE', 'MINOR', 'FAIL', 'OK', 'NOT_OK'].includes(v)) {
      t += c.weight;
      e += (['GOOD', 'OK'].includes(v) ? c.weight : (['MINOR'].includes(v) ? c.weight * 0.75 : (['MODERATE'].includes(v) ? c.weight * 0.5 : 0)));
    }
  });
  return t === 0 ? null : Math.round((e / t) * 100);
}
function getIOTripleForLabel(slotQa, label, systemsCount) {
  const key = (IO_CRITERIA.find(x => x.label.toLowerCase() === String(label).toLowerCase()) || {})?.key;
  if (!key) return { result: '', impact: '', note: '' };
  const q = slotQa?.[key];
  const n = Math.max(1, Number(systemsCount) || 1);
  const getArr = (f) => {
    if (q?.bySystem) return q.bySystem.slice(0, n).map(x => String(x?.[f] || '').trim());
    return Array(n).fill(String(q?.[f] || '').trim());
  };
  const normRes = (v) => {
    const U = v.toUpperCase();
    if (['OK', 'GOOD', 'PASS', 'VOLDOET'].includes(U)) return 'Voldoet';
    if (['MINOR'].includes(U)) return 'Grotendeels';
    if (['MODERATE'].includes(U)) return 'Matig';
    if (['NOT_OK', 'FAIL', 'POOR', 'NOK', 'VOLDOET_NIET'].includes(U)) return 'Voldoet niet';
    return v;
  };
  const normImp = (v) => {
    const U = String(v).toUpperCase();
    return U === 'A' ? 'A. Blokkerend' : (U === 'B' ? 'B. Extra werk' : (U === 'C' ? 'C. Kleine frictie' : v));
  };
  return { result: joinSemi(getArr('result').map(normRes)), impact: joinSemi(getArr('impact').map(normImp)), note: joinSemi(getArr('note')) };
}
function splitDefs(defs) {
  const a = Array.isArray(defs) ? defs : [];
  return { items: joinSemi(a.map(d => d.item)), types: joinSemi(a.map(d => d.type)), specs: joinSemi(a.map(d => d.specifications)) };
}
function splitDisruptions(dis) {
  const a = Array.isArray(dis) ? dis : [];
  const normF = (v) => {
    const U = String(v || '').toUpperCase();
    return { NEVER: '(Bijna) nooit', SOMETIMES: 'Soms', OFTEN: 'Vaak', ALWAYS: '(Bijna) altijd' }[U] || v;
  };
  return { scenarios: joinSemi(a.map(d => d.scenario)), frequencies: joinSemi(a.map(d => normF(d.frequency))), workarounds: joinSemi(a.map(d => d.workaround)) };
}
function formatWorkExp(w) { const U = String(w || '').toUpperCase(); return { OBSTACLE: 'Obstakel', ROUTINE: 'Routine', FLOW: 'Flow' }[U] || w; }

/* ==========================================================================
   Persist helpers
   ========================================================================== */
function prepareProjectForPersist(project) {
  const p = project || state.data;
  if (!p) return p;
  snapshotMergeGroupsIntoProject(p);
  ensureOutputUids(p);
  return p;
}

export async function saveToFile() {
  const p = prepareProjectForPersist(state.data || state.project);
  const dataStr = JSON.stringify(p, null, 2);
  const fileName = getFileName('json');
  try {
    if ('showSaveFilePicker' in window) {
      const handle = await window.showSaveFilePicker({
        suggestedName: fileName,
        types: [{ description: 'SIPOC Project File', accept: { 'application/json': ['.json'] } }]
      });
      const writable = await handle.createWritable(); await writable.write(dataStr); await writable.close(); return;
    }
  } catch (err) { if (err?.name === 'AbortError') return; }
  downloadBlob(new Blob([dataStr], { type: 'application/json' }), fileName);
}

export function loadFromFile(file, onSuccess) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (ev) => {
    try {
      const parsed = JSON.parse(ev.target.result);
      if (!parsed || !Array.isArray(parsed.sheets)) throw new Error('Ongeldig formaat');
      restoreMergeGroupsToLocalStorage(parsed);

      if (typeof state.loadProjectFromObject === 'function') {
        state.loadProjectFromObject(parsed);
      } else {
        try { state.project = parsed; } catch {}
        try { state.data = parsed; } catch {}
        try { if (typeof state.notify === 'function') state.notify({ reason: 'load' }, { clone: false }); }
        catch { try { state.notify(); } catch {} }
      }

      currentLoadedSha = null;
      if (onSuccess) onSuccess();
    } catch (err) { Toast.show(`Fout bij laden: ${err.message}`, 'error'); }
  };
  reader.readAsText(file);
}

/* ==========================================================================
   CSV EXPORT: Helpers & Main Function
   ========================================================================== */

// Helper: Vertaal Lean-waarde (VA, BNVA, NVA) naar label
function getLeanValueLabel(val) {
  if (!val) return '';
  const item = LEAN_VALUES.find(x => x.value === val);
  return item ? item.label : val;
}

// Helper: Vertaal Proces-status (HAPPY, NEUTRAL, SAD) naar label
function getProcessStatusLabel(val) {
  if (!val) return '';
  const item = PROCESS_STATUSES.find(x => x.value === val);
  return item ? item.label : val;
}

export function exportToCSV() {
  doFullCSVExport();
}

function doFullCSVExport() {
  try {
    const headers = [
      'Kolomnummer', 'Fase', 'Parallel?', 'Parallel met?', 'Split?', 'Route', 'Conditioneel?', 'Logica', 'Groep?', 'Groepsnaam', 'Leverancier',
      'Systemen', 'Legacy systemen', 'Target systemen', 'Systeem workarounds', 'Systeem workarounds opmerking',
      'Belemmering', 'Belemmering opmerking', 'Dubbel registreren', 'Dubbel registreren opmerking',
      'Foutgevoeligheid', 'Foutgevoeligheid opmerking', 'Gevolg bij uitval', 'Gevolg bij uitval opmerking', 'TTF Scores',
      'Input ID', 'Input', 'Input bundel(s)', 'Bundel Output IDs', 'Bundel Output teksten',
      'Compleetheid', 'Compleetheid taakimpact', 'Compleetheid taakimpact opmerking',
      'Datakwaliteit', 'Datakwaliteit taakimpact', 'Datakwaliteit taakimpact opmerking',
      'Eenduidigheid', 'Eenduidigheid taakimpact', 'Eenduidigheid taakimpact opmerking',
      'Tijdigheid', 'Tijdigheid taakimpact', 'Tijdigheid taakimpact opmerking',
      'Standaardisatie', 'Standaardisatie taakimpact', 'Standaardisatie taakimpact opmerking',
      'Overdracht', 'Overdracht taakimpact', 'Overdracht taakimpact opmerking',
      'IQF score', 'Items', 'Type', 'Specificaties', 
      'Actor', // <--- NIEUW: Actor
      'Proces', 'Type activiteit', 'Werkbeleving', 'Toelichting', 'Leanwaarde', 'Status proces',
      'Oorzaken', 'Maatregelen', 'Verstoringen', 'Frequentie', 'Proces workarounds', 'Output ID', 'Output',
      'Procesvalidatie', 'Routing bij rework', 'Routing bij pass', 'Klant'
    ];
    const lines = [headers.map(toCsvField).join(';')];
    const project = state.data || state.project;
    if (!project || !Array.isArray(project.sheets)) throw new Error('Geen data');
    ensureOutputUids(project);
    const { outIdByUid, outTextByUid, outTextByOutId } = buildGlobalOutputMaps(project);
    const bundleMaps = buildBundleMaps(project, outIdByUid, outTextByUid, outTextByOutId);
    let globalColNr = 0, globalInCounter = 0;

    (project?.sheets || []).forEach((sheet) => {
      const mergeGroups = getMergeGroupsSanitized(project, sheet);
      const variantMap = computeVariantLetterMap(sheet);
      const visibleColIdxs = (sheet.columns || []).map((c, i) => ({ c, i })).filter(x => x.c?.isVisible !== false).map(x => x.i);
      const prevMap = {}; let p = null; visibleColIdxs.forEach(i => { prevMap[i] = p; p = i; });

      (sheet.columns || []).forEach((col, colIdx) => {
        if (col?.isVisible === false) return;
        globalColNr++;
        const sysG = getMergeGroupForCell(mergeGroups, colIdx, 1);
        const sysMeta = (sysG?.systemsMeta && hasValidSystems(sysG.systemsMeta)) ? sysG.systemsMeta : buildSystemsMetaFallbackFromSlot(col?.slots?.[1]);
        const sl = systemsToLists(sysMeta);

        const isP = !!col.isParallel, pWith = isP && prevMap[colIdx] != null ? getProcessLabel(sheet, prevMap[colIdx]) : '-';
        const isS = !!col.isVariant, route = (variantMap[colIdx] || getFollowupRouteLabel(sheet, colIdx)) ? `Route ${getScopedRouteLabel(project, sheet, variantMap[colIdx] || getFollowupRouteLabel(sheet, colIdx))}` : '-';
        const isC = !!col.isConditional, logic = col.logic || {};
        let logExp = '';
        if (isC && logic.condition) logExp = `VRAAG: ${logic.condition}; JA: ${logic.ifTrue == 'SKIP' ? 'SKIP' : (logic.ifTrue ? `Ga naar ${getProcessLabel(sheet, logic.ifTrue)}` : 'Doe')}; NEE: ${logic.ifFalse == 'SKIP' ? 'SKIP' : (logic.ifFalse ? `Ga naar ${getProcessLabel(sheet, logic.ifFalse)}` : 'Doe')}`;

        const gFor = (sheet.groups || []).find(g => g.cols?.includes(colIdx));
        
        // UPDATE: Nieuwe slot indices
        // 0=Bron, 1=Sys, 2=In, 3=Actor, 4=Proces, 5=Output, 6=Klant
        const inS = col?.slots?.[2];
        const actorS = col?.slots?.[3];
        const procS = col?.slots?.[4] || {};
        const outS = col?.slots?.[5];

        let inId = '', inTxt = String(inS?.text || '').trim();
        const bRes = resolveBundleIdsToLists(normalizeLinkedBundles(inS), bundleMaps);
        const srcRes = resolveLinkedSourcesToOutPairs(normalizeLinkedSources(inS), outIdByUid, outTextByUid, outTextByOutId);
        if (bRes.bundleNames.length || srcRes.ids.length) { inId = joinSemi([...bRes.bundleNames, ...srcRes.ids]); inTxt = joinSemi([...bRes.bundleNames, ...srcRes.texts]); }
        else if (inTxt) { globalInCounter++; inId = `IN${globalInCounter}`; }

        const qa = inS?.qa || {};
        const c_ = (l) => getIOTripleForLabel(qa, l, sl.systemsCount);
        const def = splitDefs(inS?.inputDefinitions), dis = splitDisruptions(procS?.disruptions);

        // UPDATE: Output merge check op 5
        const outG = getMergeGroupForCell(mergeGroups, colIdx, 5), validGate = finalizeGate(outG?.gate);
        let outId = '';
        if (String(outS?.text || '').trim() && !isMergedSlaveInSheet(mergeGroups, colIdx, 5)) {
          const u = String(outS?.outputUid || '').trim(); outId = u && outIdByUid[u] ? outIdByUid[u] : '';
        }

        // Nieuw: Actor ophalen (uit slot 3)
        const actor = String(actorS?.text || '').trim();

        lines.push([
          globalColNr, `Procesflow ${globalColNr}`, isP ? 'Ja' : 'Nee', pWith, isS ? 'Ja' : 'Nee', route, isC ? 'Ja' : 'Nee', logExp,
          col.isGroup ? 'Ja' : 'Nee', gFor?.title || '', String(col?.slots?.[0]?.text || '').trim(),
          sl.systemNames, sl.legacySystems, sl.targetSystems, sl.systemWorkarounds, sl.systemWorkaroundsNotes,
          sl.belemmering, sl.belemmeringNotes, sl.dubbelRegistreren, sl.dubbelRegistrerenNotes,
          sl.foutgevoeligheid, sl.foutgevoeligheidNotes, sl.gevolgUitval, sl.gevolgUitvalNotes, sl.ttfScores,
          inId, inTxt, joinSemi(bRes.bundleNames), joinSemi(bRes.memberOutIds), joinSemi(bRes.memberOutTexts),
          c_('Compleetheid').result, c_('Compleetheid').impact, c_('Compleetheid').note,
          c_('Datakwaliteit').result, c_('Datakwaliteit').impact, c_('Datakwaliteit').note,
          c_('Eenduidigheid').result, c_('Eenduidigheid').impact, c_('Eenduidigheid').note,
          c_('Tijdigheid').result, c_('Tijdigheid').impact, c_('Tijdigheid').note,
          c_('Standaardisatie').result, c_('Standaardisatie').impact, c_('Standaardisatie').note,
          c_('Overdracht').result, c_('Overdracht').impact, c_('Overdracht').note,
          calculateIQFScore(qa) || '', def.items, def.types, def.specs,
          actor, // <--- Actor veld
          String(procS.text || '').trim(), String(procS.type || '').trim(), formatWorkExp(procS.workExp || procS.workJoy),
          String(procS.note || procS.toelichting || '').trim(), getLeanValueLabel(procS.processValue), getProcessStatusLabel(procS.processStatus),
          joinSemi(procS.causes || []), joinSemi(procS.improvements || []),
          dis.scenarios, dis.frequencies, dis.workarounds || joinSemi(procS.workarounds || []),
          outId, String(outS?.text || '').trim(), validGate ? 'Ja' : '', validGate ? getFailTargetFromGate(sheet, validGate) : '', validGate && outG ? getPassTargetFromGroup(sheet, outG) : '',
          String(col?.slots?.[6]?.text || '').trim() // Klant zit op 6
        ].map(toCsvField).join(';'));
      });
    });
    downloadBlob(new Blob(['\uFEFF' + lines.join('\n')], { type: 'text/csv;charset=utf-8;' }), getFileName('csv'));
  } catch (e) { Toast.show('CSV Fout: ' + e.message, 'error'); }
}

/* ==========================================================================
   EXPORT: HD image (PNG)
   ========================================================================== */
export async function exportHD(copyToClipboard = false) {
  if (!copyToClipboard) {
    await exportToPDF();
    return;
  }

  if (typeof html2canvas === 'undefined') { Toast.show('Export module niet geladen', 'error'); return; }
  const board = document.getElementById('board'); if (!board) return;
  const viewport = document.getElementById('viewport') || board;

  Toast.show('Afbeelding genereren...', 'info', 2000);

  // ensure fonts are loaded (JetBrains Mono) before cloning/rendering
  try { if (document.fonts?.ready) await document.fonts.ready; } catch {}

  try {
    const canvas = await html2canvas(viewport, {
      backgroundColor: null, // TRANSPARANT
      scale: 2.5, logging: false, useCORS: true,
      ignoreElements: (el) => el.classList.contains('col-actions'),
      onclone: (doc) => {
        doc.body.classList.add('exporting');

        // Force transparent backgrounds
        doc.documentElement.style.background = 'transparent';
        doc.body.style.background = 'transparent';
        const v = doc.getElementById('viewport'); if (v) v.style.background = 'transparent';
        const b = doc.getElementById('board'); if (b) b.style.background = 'transparent';

        // DYNAMIC WIDTH MEASUREMENT (Same as PDF)
        let dynWidth = 58;
        const vEl = doc.getElementById('viewport');
        const fc = doc.querySelector('.col');
        if (vEl && fc) {
          const vr = vEl.getBoundingClientRect();
          const cr = fc.getBoundingClientRect();
          const calc = cr.left - vr.left;
          if (calc > 10) dynWidth = calc;
        }

        // === SVG Header Replacement (EXACT TYPO) ===
        replaceRowHeadersWithExactSVG(doc, dynWidth);

        // Styles injection
        const s = doc.createElement('style');
        s.textContent = `
          .group-header-overlay, .group-header-label, .group-header-line { display: block !important; visibility: visible !important; opacity: 1 !important; }
          .group-header-overlay { position: absolute !important; top: 0px !important; z-index: 9999 !important; pointer-events: none !important; }
          .group-header-label { position: absolute !important; top: 0px !important; left: 0px !important; z-index: 10000 !important; pointer-events: none !important; }
          .group-header-line { position: absolute !important; top: 22px !important; left: 0px !important; right: 0px !important; z-index: 9999 !important; pointer-events: none !important; }
        `;
        doc.head.appendChild(s);
        doc.querySelectorAll('.group-header-overlay').forEach(el => { el.style.position = 'absolute'; el.style.top = '0px'; el.style.zIndex = '9999'; });
        doc.querySelectorAll('.group-header-label').forEach(el => { el.style.position = 'absolute'; el.style.top = '0px'; el.style.left = '0px'; el.style.zIndex = '10000'; });
        doc.querySelectorAll('.group-header-line').forEach(el => { el.style.position = 'absolute'; el.style.top = '22px'; el.style.right = '0px'; el.style.zIndex = '9999'; });
      }
    });

    if (copyToClipboard) {
      canvas.toBlob((blob) => {
        try { navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]); Toast.show('Gekopieerd!', 'success'); }
        catch { downloadCanvas(canvas); Toast.show('Klembord mislukt, gedownload', 'info'); }
      });
    } else { downloadCanvas(canvas); }
  } catch (err) { Toast.show('Screenshot mislukt', 'error'); }
}

/* ==========================================================================
   EXPORT: PDF (A3 Tiled + Smart Split + Sticky Headers + Dark Mode)
   ========================================================================== */
export async function exportToPDF() {
  if (typeof html2canvas === 'undefined') { Toast.show('Export module niet geladen', 'error'); return; }

  let jsPDF;
  try { jsPDF = await loadJsPDF(); }
  catch (e) { Toast.show('Kan PDF module niet laden', 'error'); return; }

  const board = document.getElementById('board');
  if (!board) return;
  const viewport = document.getElementById('viewport') || board;

  Toast.show('PDF Genereren (Slimme pagina-indeling)...', 'info', 4000);

  // ensure fonts are loaded (JetBrains Mono) before cloning/rendering
  try { if (document.fonts?.ready) await document.fonts.ready; } catch {}

  // Variabelen om layout op te vangen UIT de clone
  let colRegions = [];
  let headerWidthFromClone = 58; // fallback
  const scale = 2; // Hoge kwaliteit voor PDF

  try {
    const canvas = await html2canvas(viewport, {
      backgroundColor: null, // TRANSPARANT
      scale: scale,
      logging: false,
      useCORS: true,
      ignoreElements: (el) => el.classList.contains('col-actions'),
      onclone: (doc) => {
        doc.body.classList.add('exporting');

        // Force transparent backgrounds
        doc.documentElement.style.background = 'transparent';
        doc.body.style.background = 'transparent';
        const v = doc.getElementById('viewport');
        if (v) {
          v.style.background = 'transparent';
          v.style.display = 'block';
          v.style.width = 'fit-content';
          v.style.height = 'auto';
          v.style.overflow = 'visible';
        }
        const b = doc.getElementById('board');
        if (b) { b.style.background = 'transparent'; }

        // --- 1. DYNAMISCH METEN HEADER BREEDTE (PRECIES VANAF START PAGINA TOT START 1E KOLOM) ---
        const vEl = doc.getElementById('viewport');
        const firstCol = doc.querySelector('.col');
        if (vEl && firstCol) {
          const vRect = vEl.getBoundingClientRect();
          const cRect = firstCol.getBoundingClientRect();
          const calc = cRect.left - vRect.left;
          if (calc > 10) headerWidthFromClone = calc;
        }

        // === SVG Header Replacement (EXACT TYPO) ===
        replaceRowHeadersWithExactSVG(doc, headerWidthFromClone);

        // Styles injection
        const s = doc.createElement('style');
        s.textContent = `
          .group-header-overlay, .group-header-label, .group-header-line { display: block !important; visibility: visible !important; opacity: 1 !important; }
          .group-header-overlay { position: absolute !important; top: 0px !important; z-index: 9999 !important; pointer-events: none !important; }
          .group-header-label { position: absolute !important; top: 0px !important; left: 0px !important; z-index: 10000 !important; pointer-events: none !important; }
          .group-header-line { position: absolute !important; top: 22px !important; left: 0px !important; right: 0px !important; z-index: 9999 !important; pointer-events: none !important; }
        `;
        doc.head.appendChild(s);
        doc.querySelectorAll('.group-header-overlay').forEach(el => { el.style.position = 'absolute'; el.style.top = '0px'; el.style.zIndex = '9999'; });
        doc.querySelectorAll('.group-header-label').forEach(el => { el.style.position = 'absolute'; el.style.top = '0px'; el.style.left = '0px'; el.style.zIndex = '10000'; });
        doc.querySelectorAll('.group-header-line').forEach(el => { el.style.position = 'absolute'; el.style.top = '22px'; el.style.right = '0px'; el.style.zIndex = '9999'; });

        // --- MEASURE COLUMNS IN CLONE ---
        const vRect = doc.getElementById('viewport').getBoundingClientRect();
        const cols = Array.from(doc.querySelectorAll('.col'));
        colRegions = cols.map(c => {
          const cRect = c.getBoundingClientRect();
          return {
            start: (cRect.left - vRect.left) * scale,
            end: (cRect.right - vRect.left) * scale
          };
        });
        colRegions.sort((a, b) => a.start - b.start);
      }
    });

    const pdf = new jsPDF.jsPDF({ orientation: 'l', unit: 'mm', format: 'a3' });
    const pdfW = 420;
    const pdfH = 297;
    const imgHeight = canvas.height;
    const imgWidth = canvas.width;

    // Header Width (Dynamic from clone measurement)
    const headerWidthPx = headerWidthFromClone * scale;

    // Determine ratio (fit height on A3)
    const ratio = imgHeight / pdfH;

    // === CANONICAL CUT: gebruik start van de 1e kolom (voorkomt "te vroeg knippen") ===
    const firstColStartPx = (Array.isArray(colRegions) && colRegions.length) ? colRegions[0].start : null;
    const headerCutPx = Math.max(1, Math.round(firstColStartPx != null ? firstColStartPx : headerWidthPx));

    // Calculate content width available on paper (gebaseerd op canonical cut)
    const headerWidthOnPaper = headerCutPx / ratio;
    const contentWidthOnPaper = pdfW - headerWidthOnPaper;

    // Fixed viewport content width in pixels (strict alignment)
    const fixedContentWidthPx = Math.max(1, Math.round(contentWidthOnPaper * ratio));

    // Start content slicing exact op canonical cut
    let currentX = headerCutPx;
    const pageBreaks = [];

    // Smart Split Logic (sequential, maar elke slice wordt gepadded naar fixedContentWidthPx)
    while (currentX < imgWidth) {
      const idealEnd = currentX + fixedContentWidthPx;

      if (idealEnd >= imgWidth) {
        pageBreaks.push({ start: currentX, end: imgWidth });
        break;
      }

      // Find last column that fits completely
      const candidates = colRegions.filter(c => c.start >= currentX && c.end <= idealEnd);

      let cutPoint = idealEnd;
      if (candidates.length > 0) {
        cutPoint = candidates[candidates.length - 1].end;
      }

      // Force advance if stuck
      if (cutPoint <= currentX) cutPoint = idealEnd;

      pageBreaks.push({ start: currentX, end: cutPoint });
      currentX = cutPoint;
    }

    // PDF Page Generation
    for (let i = 0; i < pageBreaks.length; i++) {
      if (i > 0) pdf.addPage();

      // 1. Dark Background
      pdf.setFillColor(18, 22, 25); // #121619
      pdf.rect(0, 0, pdfW, pdfH, 'F');

      // 2. Draw Sticky Header (Left) — crop 0..headerCutPx (canonical)
      const hCanvas = document.createElement('canvas');
      hCanvas.width = headerCutPx;
      hCanvas.height = imgHeight;
      hCanvas.getContext('2d').drawImage(
        canvas,
        0, 0, hCanvas.width, imgHeight,
        0, 0, hCanvas.width, imgHeight
      );

      const hData = hCanvas.toDataURL('image/png');
      pdf.addImage(hData, 'PNG', 0, 0, headerWidthOnPaper, pdfH, undefined, 'FAST');

      // 3. Draw Content Slice (Right) — ALWAYS render in fixed viewport width
      const pb = pageBreaks[i];
      const startX = Math.max(0, Math.round(pb.start));
      const endX = Math.max(startX, Math.round(pb.end));
      const sliceWidth = Math.max(1, Math.min(endX - startX, imgWidth - startX));

      const cCanvas = document.createElement('canvas');
      cCanvas.width = fixedContentWidthPx;       // fixed width => strict alignment across pages
      cCanvas.height = imgHeight;
      const ctx = cCanvas.getContext('2d');

      // Transparent padding stays transparent; dark bg is drawn in PDF
      ctx.clearRect(0, 0, cCanvas.width, cCanvas.height);

      // Draw the available slice left-aligned into the fixed-width canvas
      ctx.drawImage(
        canvas,
        startX, 0, sliceWidth, imgHeight,
        0, 0, sliceWidth, imgHeight
      );

      const cData = cCanvas.toDataURL('image/png');

      // IMPORTANT: always same on-paper width for content area
      pdf.addImage(cData, 'PNG', headerWidthOnPaper, 0, contentWidthOnPaper, pdfH, undefined, 'FAST');
    }

    pdf.save(getFileName('pdf'));
    Toast.show('PDF gedownload!', 'success');

  } catch (err) {
    console.error(err);
    Toast.show('PDF Fout: ' + err.message, 'error');
  }
}

/* ==========================================================================
   GITHUB CLOUD OPSLAG
   ========================================================================== */

function utf8_to_b64(str) { return window.btoa(unescape(encodeURIComponent(str))); }
function b64_to_utf8(str) { return decodeURIComponent(escape(window.atob(str))); }
function getGitHubConfig() {
  return {
    token: localStorage.getItem('gh_token'), owner: localStorage.getItem('gh_owner'),
    repo: localStorage.getItem('gh_repo'), path: localStorage.getItem('gh_path') || 'ariseflow_data.json'
  };
}

export async function loadFromGitHub() {
  const { token, owner, repo, path } = getGitHubConfig();
  if (!token || !owner || !repo) throw new Error('GitHub settings missen');
  const url = `https://api.github.com/repos/${owner}/${repo}/contents/${path}`;

  // GEBRUIK 'RAW' HEADER OM GROTE BESTANDEN TE LADEN (>1MB)
  const response = await fetch(url, { headers: { Authorization: `token ${token}`, Accept: 'application/vnd.github.v3.raw' } });

  if (!response.ok) {
    if (response.status === 404) throw new Error('Bestand niet gevonden op GitHub.');
    throw new Error(`Fout bij laden: ${response.statusText}`);
  }

  const content = await response.text();
  const parsed = JSON.parse(content);
  restoreMergeGroupsToLocalStorage(parsed);

  if (typeof state.loadProjectFromObject === 'function') {
    state.loadProjectFromObject(parsed);
  } else {
    try { state.project = parsed; } catch {}
    try { state.data = parsed; } catch {}
    try { if (typeof state.notify === 'function') state.notify({ reason: 'load' }, { clone: false }); }
    catch { try { state.notify(); } catch {} }
  }

  currentLoadedSha = null;
  return true;
}

export async function saveToGitHub() {
  const { token, owner, repo, path } = getGitHubConfig();
  if (!token || !owner || !repo) { alert("Vul eerst GitHub settings in"); return; }
  const url = `https://api.github.com/repos/${owner}/${repo}/contents/${path}`;

  // 1. SHA ophalen (als bestand bestaat)
  if (!currentLoadedSha) {
    try {
      const getResp = await fetch(url, { 
        method: 'GET', 
        headers: { 
          Authorization: `token ${token}`, 
          Accept: 'application/vnd.github.v3+json' 
        } 
      });
      if (getResp.ok) { 
        const d = await getResp.json(); 
        currentLoadedSha = d.sha; 
      }
    } catch {}
  }

  const p = prepareProjectForPersist(state.data || state.project);
  
  // 2. Body opbouwen MET SHA (indien aanwezig)
  const body = {
    message: `Update via AriseFlow: ${new Date().toLocaleString()}`,
    content: utf8_to_b64(JSON.stringify(p, null, 2))
  };
  
  if (currentLoadedSha) {
    body.sha = currentLoadedSha;
  }

  // 3. Opslaan
  const putResp = await fetch(url, {
    method: 'PUT',
    headers: { Authorization: `token ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  });

  if (putResp.status === 409) {
    alert("LET OP: Iemand anders heeft dit bestand tussentijds gewijzigd!\n\nJe kunt niet opslaan om overschrijven te voorkomen.\n\nDownload je werk lokaal (Save JSON) en ververs de pagina.");
    throw new Error('Versieconflict');
  }

  if (putResp.status === 422) {
     // Soms is de SHA toch verouderd, probeer een retry als je zeker weet dat je wilt overschrijven?
     // Voor nu: error gooien
     throw new Error("GitHub weigert update (SHA mismatch of validatiefout). Ververs de pagina en probeer opnieuw.");
  }

  if (!putResp.ok) {
    const errData = await putResp.json();
    throw new Error(`Opslaan mislukt: ${errData.message}`);
  }

  const respData = await putResp.json();
  currentLoadedSha = respData.content.sha;
  Toast.show('Succesvol opgeslagen op GitHub!', 'success');
}