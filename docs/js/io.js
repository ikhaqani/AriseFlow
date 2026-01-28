// io.js (AANGEPAST: System Fit Notes + 4-punts IQF schaal + Impact A/B/C)
// + FIX: Groepsnamen zichtbaar in exportHD
// + FIX: Procesvalidatie export checkt gate completeness
// + FIX: Merge-groups (incl. gate + systemsMeta) worden nu ook persistente data bij Save/Load (JSON + GitHub)
// + FIX: NVT (System Fit) wordt meegenomen in CSV export: TTF Scores = "NVT", overige sysfit velden leeg

import { state } from './state.js';
import { Toast } from './toast.js';
import { IO_CRITERIA } from './config.js';

const MERGE_LS_PREFIX = 'ssipoc.mergeGroups.v2';

/* ==========================================================================
   CSV helpers
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
   Merge groups (System + Output) — gelezen uit localStorage
   + Persist/restore naar project JSON (sheet.mergeGroupsV2)
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
  } catch {
    return [];
  }
}

// ✅ Persist merge-groups (raw) mee in het project-bestand
function snapshotMergeGroupsIntoProject(project) {
  const p = project || state.data;
  if (!p || !Array.isArray(p.sheets)) return;

  (p.sheets || []).forEach((sheet) => {
    const raw = loadMergeGroupsRaw(p, sheet);
    if (Array.isArray(raw) && raw.length) sheet.mergeGroupsV2 = raw;
    else if ('mergeGroupsV2' in sheet) delete sheet.mergeGroupsV2;
  });
}

// ✅ Restore merge-groups (raw) terug naar localStorage na load
function restoreMergeGroupsToLocalStorage(project) {
  const p = project || state.data;
  if (!p || !Array.isArray(p.sheets)) return;

  (p.sheets || []).forEach((sheet) => {
    const raw = Array.isArray(sheet?.mergeGroupsV2) ? sheet.mergeGroupsV2 : null;
    if (!raw) return;

    try {
      localStorage.setItem(mergeKeyForSheet(p, sheet), JSON.stringify(raw));
    } catch {
      // ignore (storage full / blocked)
    }
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

/**
 * ✅ Alleen "geldige" procesvalidatie gate:
 * - enabled === true
 * - failTargetColIdx is finite number
 * Anders: null (export blijft leeg, geen bypass route in Excel)
 */
function finalizeGate(gate) {
  const g = sanitizeGate(gate);
  if (!g?.enabled) return null;
  if (g.failTargetColIdx == null) return null;
  if (!Number.isFinite(Number(g.failTargetColIdx))) return null;
  return { enabled: true, failTargetColIdx: Number(g.failTargetColIdx) };
}

function sanitizeMergeGroupForSheet(sheet, g) {
  const n = sheet?.columns?.length ?? 0;
  if (!n) return null;

  const slotIdx = Number(g?.slotIdx);
  if (![1, 2, 4].includes(slotIdx)) return null;

  const cols = Array.isArray(g?.cols) ? g.cols.map((x) => Number(x)).filter(Number.isFinite) : [];
  const uniq = [...new Set(cols)].filter((c) => c >= 0 && c < n);
  if (uniq.length < 2) return null;
  if (!isContiguousZeroBased(uniq)) return null;

  let master = Number(g?.master);
  if (!Number.isFinite(master) || !uniq.includes(master)) master = uniq[0];

  // ✅ Gate wordt alleen meegenomen als hij compleet is
  const gate = slotIdx === 4 ? finalizeGate(g?.gate) : null;

  const systemsMeta =
    slotIdx === 1 && g?.systemsMeta && typeof g.systemsMeta === 'object' ? g.systemsMeta : null;

  return { slotIdx, cols: uniq.sort((a, b) => a - b), master, gate, systemsMeta };
}

function getMergeGroupsSanitized(project, sheet) {
  return loadMergeGroupsRaw(project, sheet)
    .map((g) => sanitizeMergeGroupForSheet(sheet, g))
    .filter(Boolean);
}

function getMergeGroupForCell(groups, colIdx, slotIdx) {
  return (
    (groups || []).find(
      (x) => x.slotIdx === slotIdx && Array.isArray(x.cols) && x.cols.includes(colIdx)
    ) || null
  );
}

function isMergedSlaveInSheet(groups, colIdx, slotIdx) {
  const g = getMergeGroupForCell(groups, colIdx, slotIdx);
  return !!g && colIdx !== g.master;
}

function getNextVisibleColIdx(sheet, fromIdx) {
  const n = sheet?.columns?.length ?? 0;
  for (let i = fromIdx + 1; i < n; i++) {
    if (sheet.columns[i]?.isVisible !== false) return i;
  }
  return null;
}

function getProcessLabel(sheet, colIdx) {
  const t = sheet?.columns?.[colIdx]?.slots?.[3]?.text;
  const s = String(t ?? '').trim();
  return s || `Kolom ${Number(colIdx) + 1}`;
}

function getPassTargetFromGroup(sheet, group) {
  if (!group?.cols?.length) return '';
  const maxCol = Math.max(...group.cols);
  const nextIdx = getNextVisibleColIdx(sheet, maxCol);
  if (nextIdx == null) return 'Einde proces';
  return getProcessLabel(sheet, nextIdx);
}

function getFailTargetFromGate(sheet, gate) {
  const g = finalizeGate(gate);
  if (!g) return '';
  return getProcessLabel(sheet, g.failTargetColIdx);
}

/* ==========================================================================
   ROUTE PREFIX (SCOPED): RF{sheetIndex}-{routeLabel}
   ========================================================================== */

function getSheetRoutePrefix(project, sheet) {
  const p = project || state.data;
  const sheets = Array.isArray(p?.sheets) ? p.sheets : [];
  const idx = sheets.findIndex((s) => s?.id === sheet?.id);
  const n = idx >= 0 ? idx + 1 : 1;
  return `RF${n}`;
}

function getScopedRouteLabel(project, sheet, label) {
  const base = String(label || '').trim();
  if (!base) return '';
  return `${getSheetRoutePrefix(project, sheet)}-${base}`;
}

/* ==========================================================================
   Variant route letters & Follow-up logic
   ========================================================================== */

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
  const rootParents = Object.keys(colGroups)
    .map(Number)
    .filter((p) => !allChildren.has(p));

  rootParents.forEach((root) => assignLabels(root, ''));

  let legacyCounter = 0;
  sheet.columns.forEach((col, i) => {
    if (col.isVariant && !map[i]) {
      map[i] = toLetter(legacyCounter++);
    }
  });

  return map;
}

function getFollowupRouteLabel(sheet, colIdx) {
  const col = sheet.columns?.[colIdx];
  if (!col) return null;

  const manualRoute = String(col.routeLabel || '').trim(); // "A", "B", etc.
  const isSplitStart = !!col.isVariant;

  if (!manualRoute || isSplitStart) return null;

  let count = 1;
  for (let i = 0; i < colIdx; i++) {
    const c = sheet.columns?.[i];
    if (!c) continue;
    if (c.isVariant) continue;
    if (String(c.routeLabel || '').trim() === manualRoute) count++;
  }

  return `${manualRoute}.${count}`; // A.1, A.2, ...
}

/* ==========================================================================
   Output ID stabilisatie (outputUid -> OUTx) + Output tekst maps
   ========================================================================== */

function makeId(prefix = 'id') {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return `${prefix}_${crypto.randomUUID()}`;
  }
  return `${prefix}_${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

// ✅ Zorg dat outputs altijd outputUid krijgen vóór persist (zodat linking stabiel is)
function ensureOutputUids(project) {
  const p = project || state.data;
  if (!p?.sheets?.length) return;

  (p.sheets || []).forEach((sheet) => {
    const groups = getMergeGroupsSanitized(p, sheet);

    (sheet?.columns || []).forEach((col, colIdx) => {
      if (col?.isVisible === false) return;

      const outSlot = col?.slots?.[4];
      const outText = String(outSlot?.text ?? '').trim();
      if (!outText) return;

      if (isMergedSlaveInSheet(groups, colIdx, 4)) return;

      if (!outSlot.outputUid || String(outSlot.outputUid).trim() === '') {
        outSlot.outputUid = makeId('out');
      }
    });
  });
}

function buildGlobalOutputMaps(project) {
  const outIdByUid = {};
  const outTextByUid = {};
  const outTextByOutId = {};

  let outCounter = 0;

  (project?.sheets || []).forEach((sheet) => {
    const groups = getMergeGroupsSanitized(project, sheet);

    (sheet?.columns || []).forEach((col, colIdx) => {
      if (col?.isVisible === false) return;

      const outSlot = col?.slots?.[4];
      const outText = String(outSlot?.text ?? '').trim();
      if (!outText) return;

      if (isMergedSlaveInSheet(groups, colIdx, 4)) return;

      if (!outSlot.outputUid || String(outSlot.outputUid).trim() === '') {
        outSlot.outputUid = makeId('out');
      }

      outCounter += 1;
      const outId = `OUT${outCounter}`;

      outIdByUid[outSlot.outputUid] = outId;
      outTextByUid[outSlot.outputUid] = outText;
      outTextByOutId[outId] = outText;
    });
  });

  return { outIdByUid, outTextByUid, outTextByOutId };
}

/* ==========================================================================
   INPUT linking: single + multiple links (OUTx / outputUid)
   ========================================================================== */

function _looksLikeOutId(v) {
  const s = String(v || '').trim();
  return !!s && /^OUT\d+$/.test(s);
}

function normalizeLinkedSources(inputSlot) {
  const out = [];

  const idsArr = Array.isArray(inputSlot?.linkedSourceIds) ? inputSlot.linkedSourceIds : null;
  const uidsArr = Array.isArray(inputSlot?.linkedSourceUids) ? inputSlot.linkedSourceUids : null;

  if (uidsArr && uidsArr.length) {
    uidsArr.forEach((u) => {
      const s = String(u ?? '').trim();
      if (s) out.push({ kind: 'uid', value: s });
    });
  }

  if (idsArr && idsArr.length) {
    idsArr.forEach((id) => {
      const s = String(id ?? '').trim();
      if (s) out.push({ kind: 'id', value: s });
    });
  }

  const singleUid = String(inputSlot?.linkedSourceUid || '').trim();
  const singleId = String(inputSlot?.linkedSourceId || '').trim();

  if (singleUid) out.push({ kind: 'uid', value: singleUid });
  if (singleId) out.push({ kind: 'id', value: singleId });

  const seen = new Set();
  const uniq = [];
  for (const x of out) {
    const key = `${x.kind}:${x.value}`;
    if (seen.has(key)) continue;
    seen.add(key);
    uniq.push(x);
  }

  return uniq;
}

function resolveLinkedSourcesToOutPairs(sources, outIdByUid, outTextByUid, outTextByOutId) {
  const ids = [];
  const texts = [];

  (sources || []).forEach((src) => {
    if (!src?.value) return;

    if (src.kind === 'id') {
      const raw = String(src.value).trim();
      if (!_looksLikeOutId(raw)) return;

      const outId = raw;
      const txt = String(outTextByOutId?.[outId] ?? '').trim() || outId;

      ids.push(outId);
      texts.push(txt);
      return;
    }

    if (src.kind === 'uid') {
      const uid = String(src.value).trim();
      if (!uid) return;

      const outId = String(outIdByUid?.[uid] ?? '').trim() || uid;
      const txt = String(outTextByUid?.[uid] ?? '').trim() || outId;

      ids.push(outId);
      texts.push(txt);
    }
  });

  return { ids, texts };
}

/* ==========================================================================
   OUTPUT bundels (bundelnaam -> OUTx; OUTy) voor compacte input
   ========================================================================== */

function normalizeLinkedBundles(inputSlot) {
  const out = [];

  const idsArr = Array.isArray(inputSlot?.linkedBundleIds) ? inputSlot.linkedBundleIds : null;
  if (idsArr && idsArr.length) {
    idsArr.forEach((id) => {
      const s = String(id ?? '').trim();
      if (s) out.push(s);
    });
  }

  const single = String(inputSlot?.linkedBundleId || '').trim();
  if (single) out.push(single);

  const seen = new Set();
  const uniq = [];
  for (const x of out) {
    if (seen.has(x)) continue;
    seen.add(x);
    uniq.push(x);
  }

  return uniq;
}

function buildBundleMaps(project, outIdByUid, outTextByUid, outTextByOutId) {
  buildGlobalOutputMaps(project); // no-op safety
  const nameById = {};
  const outIdsById = {};
  const outTextsById = {};

  const bundles = Array.isArray(project?.outputBundles) ? project.outputBundles : [];

  bundles.forEach((b) => {
    if (!b || typeof b !== 'object') return;

    const id = String(b.id ?? '').trim();
    if (!id) return;

    const name = String(b.name ?? '').trim();
    nameById[id] = name || id;

    const rawOutIds = Array.isArray(b.outIds) ? b.outIds : [];
    const rawUids = Array.isArray(b.outputUids) ? b.outputUids : [];

    const ids = [];
    const texts = [];

    rawOutIds.forEach((oid) => {
      const outId = String(oid ?? '').trim();
      if (!outId) return;
      ids.push(outId);
      texts.push(String(outTextByOutId?.[outId] ?? '').trim() || outId);
    });

    rawUids.forEach((u) => {
      const uid = String(u ?? '').trim();
      if (!uid) return;
      const outId = String(outIdByUid?.[uid] ?? '').trim() || uid;
      ids.push(outId);
      texts.push(String(outTextByUid?.[uid] ?? '').trim() || outId);
    });

    outIdsById[id] = [...new Set(ids.filter(Boolean))];
    outTextsById[id] = texts.filter(Boolean);
  });

  return { nameById, outIdsById, outTextsById };
}

function resolveBundleIdsToLists(bundleIds, bundleMaps) {
  const ids = Array.isArray(bundleIds) ? bundleIds : [];

  const bundleNames = [];
  const memberOutIds = [];
  const memberOutTexts = [];

  ids.forEach((bid) => {
    const id = String(bid ?? '').trim();
    if (!id) return;

    const name = String(bundleMaps?.nameById?.[id] ?? '').trim() || id;
    bundleNames.push(name);

    const outs = Array.isArray(bundleMaps?.outIdsById?.[id]) ? bundleMaps.outIdsById[id] : [];
    const txts = Array.isArray(bundleMaps?.outTextsById?.[id]) ? bundleMaps.outTextsById[id] : [];

    outs.forEach((x) => memberOutIds.push(String(x ?? '').trim()));
    txts.forEach((x) => memberOutTexts.push(String(x ?? '').trim()));
  });

  return {
    bundleNames: bundleNames.filter(Boolean),
    memberOutIds: memberOutIds.filter(Boolean),
    memberOutTexts: memberOutTexts.filter(Boolean)
  };
}

/* ==========================================================================
   Systems meta + TTF (per systeem)
   ========================================================================== */

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
    { key: 'RISK', label: 'Groot risico', score: 0.33 },
    { key: 'STOP', label: 'Volledige stilstand', score: 0 }
  ]
};

function hasValidSystems(meta) {
  if (!meta || typeof meta !== 'object') return false;
  if (!Array.isArray(meta.systems) || meta.systems.length === 0) return false;
  return meta.systems.some((s) => s && String(s.name || '').trim() !== '');
}

function buildSystemsMetaFallbackFromSlot(sysSlot) {
  const slot = sysSlot && typeof sysSlot === 'object' ? sysSlot : {};
  const sd = slot.systemData && typeof slot.systemData === 'object' ? slot.systemData : null;
  const postItText = String(slot.text || '').trim();

  if (sd?.systemsMeta && hasValidSystems(sd.systemsMeta)) {
    return sd.systemsMeta;
  }

  if (Array.isArray(sd?.systems) && sd.systems.length > 0) {
    const validLegacy = sd.systems.filter((s) => s && String(s.name || '').trim() !== '');
    if (validLegacy.length > 0) {
      return { multi: validLegacy.length > 1, systems: validLegacy };
    }
  }

  const nameFromSd = String(sd?.systemName ?? '').trim();
  if (nameFromSd) {
    return {
      multi: false,
      systems: [{ name: nameFromSd, legacy: false, future: '', qa: {}, score: null }]
    };
  }

  if (postItText) {
    return {
      multi: false,
      systems: [{ name: postItText, legacy: false, future: '', qa: {}, score: null }]
    };
  }

  return null;
}

function sanitizeSystemsMeta(meta) {
  if (!meta || typeof meta !== 'object') return null;

  const arr = Array.isArray(meta.systems) ? meta.systems : [];
  const systems = arr
    .map((s) => {
      if (!s || typeof s !== 'object') return null;
      return {
        name: String(s.name ?? '').trim(),
        legacy: !!s.legacy,
        future: String(s.future ?? '').trim(),
        qa: s.qa && typeof s.qa === 'object' ? { ...s.qa } : {},
        score: Number.isFinite(Number(s.score)) ? Number(s.score) : null
      };
    })
    .filter(Boolean);

  const inferredMulti = systems.length > 1;
  const multi = !!meta.multi || inferredMulti;

  if (systems.length === 0) systems.push({ name: '', legacy: false, future: '', qa: {}, score: null });
  return { multi, systems };
}

function computeTTFSystemScore(sys) {
  const qa = sys?.qa && typeof sys.qa === 'object' ? sys.qa : {};

  // NVT: niet meenemen in scoreberekening
  if (qa.__nvt === true) return null;

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

function getSysAnswerLabel(sys, qid) {
  const qa = sys?.qa && typeof sys.qa === 'object' ? sys.qa : {};
  if (qa.__nvt === true) return '';

  const q = SYSFIT_Q.find((x) => x.id === qid);
  if (!q) return '';
  const key = qa[qid];
  if (!key) return '';
  const opt = (SYSFIT_OPTS[q.type] || []).find((o) => o.key === key);
  return opt?.label || '';
}

function getSysNote(sys, qid) {
  const qa = sys?.qa && typeof sys.qa === 'object' ? sys.qa : {};
  if (qa.__nvt === true) return '';
  return String(qa[qid + '_note'] || '').trim();
}

function computeTTFScoreListFromMeta(meta) {
  const clean = sanitizeSystemsMeta(meta);
  if (!clean?.systems?.length) return [];

  return clean.systems.map((s) => {
    const stored = s?.score;
    if (Number.isFinite(Number(stored))) return Number(stored);
    const computed = computeTTFSystemScore(s);
    return Number.isFinite(Number(computed)) ? Number(computed) : null;
  });
}

function systemsToLists(meta) {
  const clean = sanitizeSystemsMeta(meta);
  const systems = (clean?.systems || []).filter((s) => s && String(s.name || '').trim() !== '');

  if (systems.length === 0 && clean?.systems?.length > 0) {
    return {
      systemNames: '',
      legacySystems: '',
      targetSystems: '',
      systemWorkarounds: '',
      systemWorkaroundsNotes: '',
      belemmering: '',
      belemmeringNotes: '',
      dubbelRegistreren: '',
      dubbelRegistrerenNotes: '',
      foutgevoeligheid: '',
      foutgevoeligheidNotes: '',
      gevolgUitval: '',
      gevolgUitvalNotes: '',
      ttfScores: '',
      systemsCount: 1
    };
  }

  const names = systems.map((s) => String(s?.name || '').trim()).filter(Boolean);
  const legacyNames = systems
    .filter((s) => !!s.legacy)
    .map((s) => String(s?.name || '').trim())
    .filter(Boolean);
  const targetNames = systems
    .filter((s) => !!s.legacy)
    .map((s) => String(s?.future || '').trim())
    .filter(Boolean);

  // ✅ NVT export: TTF Scores = "NVT"
  const ttfScores = systems.map((sys) => {
    const qa = sys?.qa && typeof sys.qa === 'object' ? sys.qa : {};
    if (qa.__nvt === true) return 'NVT';

    const stored = sys?.score;
    if (Number.isFinite(Number(stored))) return `${Number(stored)}%`;

    const computed = computeTTFSystemScore(sys);
    return computed == null ? '' : `${Number(computed)}%`;
  });

  const sysWorkarounds = systems.map((s) => getSysAnswerLabel(s, 'q1'));
  const sysWorkaroundsNotes = systems.map((s) => getSysNote(s, 'q1'));

  const belemmering = systems.map((s) => getSysAnswerLabel(s, 'q2'));
  const belemmeringNotes = systems.map((s) => getSysNote(s, 'q2'));

  const dubbelReg = systems.map((s) => getSysAnswerLabel(s, 'q3'));
  const dubbelRegNotes = systems.map((s) => getSysNote(s, 'q3'));

  const foutgevoelig = systems.map((s) => getSysAnswerLabel(s, 'q4'));
  const foutgevoeligNotes = systems.map((s) => getSysNote(s, 'q4'));

  const uitvalImpact = systems.map((s) => getSysAnswerLabel(s, 'q5'));
  const uitvalImpactNotes = systems.map((s) => getSysNote(s, 'q5'));

  return {
    systemNames: joinSemi(names),
    legacySystems: joinSemi(legacyNames),
    targetSystems: joinSemi(targetNames),
    systemWorkarounds: joinSemi(sysWorkarounds),
    systemWorkaroundsNotes: joinSemi(sysWorkaroundsNotes),
    belemmering: joinSemi(belemmering),
    belemmeringNotes: joinSemi(belemmeringNotes),
    dubbelRegistreren: joinSemi(dubbelReg),
    dubbelRegistrerenNotes: joinSemi(dubbelRegNotes),
    foutgevoeligheid: joinSemi(foutgevoelig),
    foutgevoeligheidNotes: joinSemi(foutgevoeligNotes),
    gevolgUitval: joinSemi(uitvalImpact),
    gevolgUitvalNotes: joinSemi(uitvalImpactNotes),
    ttfScores: joinSemi(ttfScores),
    systemsCount: Math.max(1, systems.length)
  };
}

/* ==========================================================================
   IQF (Input quality) score
   ========================================================================== */

function calculateIQFScore(qa) {
  if (!qa) return null;

  let totalW = 0;
  let earnedW = 0;

  IO_CRITERIA.forEach((c) => {
    const val = qa?.[c.key]?.result;
    const isScored = ['GOOD', 'POOR', 'MODERATE', 'MINOR', 'FAIL', 'OK', 'NOT_OK'].includes(val);
    if (!isScored) return;

    totalW += c.weight;

    if (val === 'GOOD' || val === 'OK') earnedW += c.weight;
    else if (val === 'MINOR') earnedW += c.weight * 0.75;
    else if (val === 'MODERATE') earnedW += c.weight * 0.5;
  });

  return totalW === 0 ? null : Math.round((earnedW / totalW) * 100);
}

/* ==========================================================================
   IO criteria per systeem (result/impact/opmerking) met ";" alignment
   ========================================================================== */

function normalizeIOResultLabel(v) {
  const s = String(v ?? '').trim();
  if (!s) return '';
  const U = s.toUpperCase();
  if (U === 'OK' || U === 'GOOD' || U === 'PASS' || U === 'VOLDOET') return 'Voldoet';
  if (U === 'MINOR') return 'Grotendeels';
  if (U === 'MODERATE') return 'Matig';
  if (U === 'NOT_OK' || U === 'FAIL' || U === 'POOR' || U === 'NOK' || U === 'VOLDOET_NIET') return 'Voldoet niet';
  if (U === 'NVT' || U === 'NA') return '';
  return s;
}

function normalizeImpactLabel(v) {
  const s = String(v ?? '').trim().toUpperCase();
  if (s === 'A') return 'A. Blokkerend';
  if (s === 'B') return 'B. Extra werk';
  if (s === 'C') return 'C. Kleine frictie';
  return s;
}

function getCriterionKeyByLabel(label) {
  const L = String(label || '').trim().toLowerCase();
  const c = (IO_CRITERIA || []).find((x) => String(x?.label || '').trim().toLowerCase() === L);
  return c?.key || null;
}

function extractPerSystem(q, systemsCount, field) {
  const n = Math.max(1, Number(systemsCount) || 1);
  const empty = Array(n).fill('');

  if (q == null) return empty;

  if (typeof q === 'string' || typeof q === 'number' || typeof q === 'boolean') {
    return Array(n).fill(String(q));
  }

  if (q && typeof q === 'object') {
    if (Array.isArray(q.bySystem)) {
      return Array(n)
        .fill(null)
        .map((_, i) => {
          const it = q.bySystem[i];
          if (!it || typeof it !== 'object') return '';
          return String(it?.[field] ?? '').trim();
        });
    }
    return Array(n).fill(String(q?.[field] ?? '').trim());
  }

  return empty;
}

function getIOTripleForLabel(slotQa, label, systemsCount) {
  const key = getCriterionKeyByLabel(label);
  if (!key) return { result: '', impact: '', note: '' };

  const q = slotQa?.[key];

  const resArr = extractPerSystem(q, systemsCount, 'result').map(normalizeIOResultLabel);
  const impactArr = extractPerSystem(q, systemsCount, 'impact').map(normalizeImpactLabel);
  const noteArr = extractPerSystem(q, systemsCount, 'note');

  return {
    result: joinSemi(resArr),
    impact: joinSemi(impactArr),
    note: joinSemi(noteArr)
  };
}

/* ==========================================================================
   Input definitions (Items / Type / Specificaties) aligned met ";"
   ========================================================================== */

function splitDefs(defs) {
  const arr = Array.isArray(defs) ? defs : [];
  const items = [];
  const types = [];
  const specs = [];

  arr.forEach((d) => {
    if (!d || typeof d !== 'object') return;
    items.push(String(d.item ?? '').trim());
    types.push(String(d.type ?? '').trim());
    specs.push(String(d.specifications ?? '').trim());
  });

  return {
    items: joinSemi(items),
    types: joinSemi(types),
    specs: joinSemi(specs)
  };
}

/* ==========================================================================
   Analyse: oorzaken/maatregelen + verstoringen/frequentie/workarounds (aligned)
   ========================================================================== */

function normalizeFreqLabel(v) {
  const s = String(v ?? '').trim();
  if (!s) return '';
  const U = s.toUpperCase();
  if (U === 'NEVER') return '(Bijna) nooit';
  if (U === 'SOMETIMES') return 'Soms';
  if (U === 'OFTEN') return 'Vaak';
  if (U === 'ALWAYS') return '(Bijna) altijd';
  if (U === 'ZELDEN') return 'Zelden';
  if (U === 'SOMS') return 'Soms';
  if (U === 'VAAK') return 'Vaak';
  if (U === 'ALTIJD') return 'Altijd';
  return s;
}

function splitDisruptions(disruptions) {
  const arr = Array.isArray(disruptions) ? disruptions : [];
  const scenarios = [];
  const freqs = [];
  const workarounds = [];

  arr.forEach((d) => {
    if (!d || typeof d !== 'object') return;
    scenarios.push(String(d.scenario ?? '').trim());
    freqs.push(normalizeFreqLabel(d.frequency));
    workarounds.push(String(d.workaround ?? '').trim());
  });

  return {
    scenarios: joinSemi(scenarios),
    frequencies: joinSemi(freqs),
    workarounds: joinSemi(workarounds)
  };
}

/* ==========================================================================
   Werkbeleving / Lean / Status
   ========================================================================== */

function formatWorkExp(workExp) {
  const v = String(workExp ?? '').trim().toUpperCase();
  if (v === 'OBSTACLE') return 'Obstakel';
  if (v === 'ROUTINE') return 'Routine';
  if (v === 'FLOW') return 'Flow';
  return String(workExp ?? '').trim();
}

function getLeanValueLabel(v) {
  const s = String(v ?? '').trim();
  return s || '';
}

function getProcessStatusLabel(v) {
  const s = String(v ?? '').trim();
  return s || '';
}

/* ==========================================================================
   Persist helpers (voor Save/Load/GitHub)
   ========================================================================== */

function prepareProjectForPersist(project) {
  const p = project || state.data;
  if (!p) return p;

  // 1) Merge-groups mee in JSON (incl gate + systemsMeta)
  snapshotMergeGroupsIntoProject(p);

  // 2) OutputUids afdwingen vóór save (zodat links stabiel blijven)
  ensureOutputUids(p);

  return p;
}

/* ==========================================================================
   JSON save/load
   ========================================================================== */

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
      const writable = await handle.createWritable();
      await writable.write(dataStr);
      await writable.close();
      return;
    }
  } catch (err) {
    if (err?.name === 'AbortError') return;
    console.warn('FS API failed, falling back to legacy download.', err);
  }

  downloadBlob(new Blob([dataStr], { type: 'application/json' }), fileName);
}

export function loadFromFile(file, onSuccess) {
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (ev) => {
    try {
      const parsed = JSON.parse(ev.target.result);
      if (!parsed || !Array.isArray(parsed.sheets)) throw new Error('Ongeldig formaat: Geen sheets gevonden.');

      // ✅ Restore merge groups naar localStorage vóór render
      restoreMergeGroupsToLocalStorage(parsed);

  // ✅ Update state (robust: state.data kan readonly zijn)
  if (typeof state.loadProjectFromObject === 'function') {
    state.loadProjectFromObject(parsed);
  } else {
    try { state.project = parsed; } catch {}
    try { state.data = parsed; } catch {}
    try { if (typeof state.notify === 'function') state.notify({ reason: 'load' }, { clone: false }); }
    catch { try { state.notify(); } catch {} }
  }

      if (onSuccess) onSuccess();
    } catch (err) {
      console.error('Load Error:', err);
      Toast.show(`Fout bij laden: ${err.message}`, 'error');
    }
  };
  reader.readAsText(file);
}

/* ==========================================================================
   EXPORT: As-Is CSV (1 rij per kolom)
   ========================================================================== */

export function exportToCSV() {
  try {
    const headers = [
      'Kolomnummer',
      'Fase',
      'Parallel?',
      'Parallel met?',
      'Split?',
      'Route',
      'Conditioneel?',
      'Logica',
      'Groep?',
      'Groepsnaam',
      'Leverancier',
      'Systemen',
      'Legacy systemen',
      'Target systemen',
      'Systeem workarounds',
      'Systeem workarounds opmerking',
      'Belemmering',
      'Belemmering opmerking',
      'Dubbel registreren',
      'Dubbel registreren opmerking',
      'Foutgevoeligheid',
      'Foutgevoeligheid opmerking',
      'Gevolg bij uitval',
      'Gevolg bij uitval opmerking',
      'TTF Scores',
      'Input ID',
      'Input',
      'Input bundel(s)',
      'Bundel Output IDs',
      'Bundel Output teksten',
      'Compleetheid',
      'Compleetheid taakimpact',
      'Compleetheid taakimpact opmerking',
      'Datakwaliteit',
      'Datakwaliteit taakimpact',
      'Datakwaliteit taakimpact opmerking',
      'Eenduidigheid',
      'Eenduidigheid taakimpact',
      'Eenduidigheid taakimpact opmerking',
      'Tijdigheid',
      'Tijdigheid taakimpact',
      'Tijdigheid taakimpact opmerking',
      'Standaardisatie',
      'Standaardisatie taakimpact',
      'Standaardisatie taakimpact opmerking',
      'Overdracht',
      'Overdracht taakimpact',
      'Overdracht taakimpact opmerking',
      'IQF score',
      'Items',
      'Type',
      'Specificaties',
      'Proces',
      'Type activiteit',
      'Werkbeleving',
      'Toelichting',
      'Leanwaarde',
      'Status proces',
      'Oorzaken',
      'Maatregelen',
      'Verstoringen',
      'Frequentie',
      'Proces workarounds',
      'Output ID',
      'Output',
      'Procesvalidatie',
      'Routing bij rework',
      'Routing bij pass',
      'Klant'
    ];

    const lines = [headers.map(toCsvField).join(';')];

    const project = state.data || state.project;

    if (!project || !Array.isArray(project.sheets)) {
      throw new Error('Geen project data gevonden (state.data ontbreekt of is ongeldig).');
    }

    // Output UIDs afdwingen voor stabiele mapping (ook zonder eerdere exports)
    ensureOutputUids(project);

    const { outIdByUid, outTextByUid, outTextByOutId } = buildGlobalOutputMaps(project);
    const bundleMaps = buildBundleMaps(project, outIdByUid, outTextByUid, outTextByOutId);

    let globalColNr = 0;
    let globalInCounter = 0;

    (project?.sheets || []).forEach((sheet) => {
      const mergeGroups = getMergeGroupsSanitized(project, sheet);
      const variantMap = computeVariantLetterMap(sheet);

      const visibleColIdxs = (sheet.columns || [])
        .map((c, i) => ({ c, i }))
        .filter(({ c }) => c?.isVisible !== false)
        .map(({ i }) => i);

      const prevVisibleByIdx = {};
      let prev = null;
      visibleColIdxs.forEach((idx) => {
        prevVisibleByIdx[idx] = prev;
        prev = idx;
      });

      (sheet.columns || []).forEach((col, colIdx) => {
        if (col?.isVisible === false) return;

        globalColNr += 1;

        const leverancier = String(col?.slots?.[0]?.text ?? '').trim();

        const sysSlot = col?.slots?.[1] || {};
        const sysGroup = getMergeGroupForCell(mergeGroups, colIdx, 1);

        let sysMeta = null;

        if (sysGroup?.systemsMeta && hasValidSystems(sysGroup.systemsMeta)) {
          sysMeta = sysGroup.systemsMeta;
        }

        if (!sysMeta) {
          sysMeta = buildSystemsMetaFallbackFromSlot(sysSlot);
        }

        const sysLists = systemsToLists(sysMeta);
        const systemsCount = sysLists.systemsCount;

        const isParallel = !!col.isParallel;
        const prevIdx = prevVisibleByIdx[colIdx];
        const parallelWith = isParallel && prevIdx != null ? getProcessLabel(sheet, prevIdx) : '-';

        const isSplit = !!col.isVariant;

        let calculatedRoute = variantMap[colIdx] || null;
        if (!calculatedRoute) {
          calculatedRoute = getFollowupRouteLabel(sheet, colIdx);
        }

        const scoped = calculatedRoute ? getScopedRouteLabel(project, sheet, calculatedRoute) : '';
        const route = scoped ? `Route ${scoped}` : '-';

        const isConditional = !!col.isConditional;

        const logic = col.logic || {};
        let logicExport = '';

        if (isConditional && logic.condition) {
          const getLabel = (val) => {
            if (val === 'SKIP') return 'SKIP (Overslaan)';
            if (val !== null) return `Ga naar ${getProcessLabel(sheet, val)}`;
            return 'Voer stap uit';
          };

          const trueAction = getLabel(logic.ifTrue);
          const falseAction = getLabel(logic.ifFalse);

          logicExport = `VRAAG: ${logic.condition}`;
          logicExport += `; INDIEN JA: ${trueAction}`;
          logicExport += `; INDIEN NEE: ${falseAction}`;
        }

        const isGroup = !!col.isGroup;
        const groupForThisCol = (sheet.groups || []).find((g) => g.cols && g.cols.includes(colIdx));
        const groupName = groupForThisCol ? groupForThisCol.title || '' : '';

        const fase = `Procesflow ${globalColNr}`;

        const inputSlot = col?.slots?.[2];

        const inGroup = getMergeGroupForCell(mergeGroups, colIdx, 2);
        const isInSlave = !!inGroup && isMergedSlaveInSheet(mergeGroups, colIdx, 2);

        const outputSlot = col?.slots?.[4];

        let inputId = '';
        let inputText = String(inputSlot?.text ?? '').trim();

        const linkedBundleIds = normalizeLinkedBundles(inputSlot);
        const bundleResolved = resolveBundleIdsToLists(linkedBundleIds, bundleMaps);

        let inputBundlesStr = joinSemi(bundleResolved.bundleNames);
        let bundleOutIdsStr = joinSemi(bundleResolved.memberOutIds);
        let bundleOutTextsStr = joinSemi(bundleResolved.memberOutTexts);
        
        if (isInSlave) {
          inputBundlesStr = '';
          bundleOutIdsStr = '';
          bundleOutTextsStr = '';
        }

        const sources = normalizeLinkedSources(inputSlot);
        const resolved = resolveLinkedSourcesToOutPairs(sources, outIdByUid, outTextByUid, outTextByOutId);

        const partsId = [];
        const partsText = [];

        if (bundleResolved.bundleNames.length) {
          partsId.push(...bundleResolved.bundleNames);
          partsText.push(...bundleResolved.bundleNames);
        }

        if (resolved.ids.length) {
          partsId.push(...resolved.ids);
          partsText.push(...resolved.texts);
        }

        if (!resolved.ids.length && !bundleResolved.bundleNames.length) {
          if (inputText) {
            globalInCounter += 1;
            inputId = `IN${globalInCounter}`;
          } else {
            inputId = '';
            inputText = '';
          }
        } else {
          inputId = joinSemi(partsId);
          inputText = joinSemi(partsText);
        }
        if (isInSlave) {
          inputId = '';
          inputText = '';
        }

        const slotQa = isInSlave ? {} : (inputSlot?.qa || {});

        const comp = getIOTripleForLabel(slotQa, 'Compleetheid', systemsCount);
        const dq = getIOTripleForLabel(slotQa, 'Datakwaliteit', systemsCount);
        const ed = getIOTripleForLabel(slotQa, 'Eenduidigheid', systemsCount);
        const tj = getIOTripleForLabel(slotQa, 'Tijdigheid', systemsCount);
        const st = getIOTripleForLabel(slotQa, 'Standaardisatie', systemsCount);
        const ov = getIOTripleForLabel(slotQa, 'Overdracht', systemsCount);

        const iqfScore = calculateIQFScore(slotQa);
        const iqfScoreStr = iqfScore == null ? '' : String(iqfScore);

        const def = isInSlave ? splitDefs(null) : splitDefs(inputSlot?.inputDefinitions);

        const procSlot = col?.slots?.[3] || {};
        const proces = String(procSlot?.text ?? '').trim();
        const typeActiviteit = String(procSlot?.type ?? '').trim();
        const werkbeleving = formatWorkExp(procSlot?.workExp ?? procSlot?.workjoy ?? procSlot?.workJoy ?? '');
        const toelichting = String(procSlot?.note ?? procSlot?.toelichting ?? procSlot?.context ?? '').trim();

        const leanwaarde = getLeanValueLabel(procSlot?.processValue ?? '');
        const statusProces = getProcessStatusLabel(procSlot?.processStatus ?? '');

        const oorzaken = Array.isArray(procSlot?.causes)
          ? joinSemi(procSlot.causes)
          : String(procSlot?.causes ?? '').trim();

        const maatregelen = Array.isArray(procSlot?.improvements)
          ? joinSemi(procSlot.improvements)
          : String(procSlot?.improvements ?? '').trim();

        const dis = splitDisruptions(procSlot?.disruptions);
        const procesWorkarounds =
          dis.workarounds ||
          (Array.isArray(procSlot?.workarounds)
            ? joinSemi(procSlot.workarounds)
            : String(procSlot?.processWorkarounds ?? '').trim());

        const outGroup = getMergeGroupForCell(mergeGroups, colIdx, 4);
        const outText = String(outputSlot?.text ?? '').trim();

        let outputId = '';
        if (outText && !isMergedSlaveInSheet(mergeGroups, colIdx, 4)) {
          const outUid = String(outputSlot?.outputUid || '').trim();
          outputId = outUid && outIdByUid[outUid] ? outIdByUid[outUid] : '';
        }

        // ✅ FIX: Procesvalidatie + routing alleen exporteren als gate compleet is
        const validGate = finalizeGate(outGroup?.gate);

        const procesValidatie = validGate ? 'Ja' : '';
        const routingRework = validGate ? getFailTargetFromGate(sheet, validGate) : '';
        const routingPass = validGate && outGroup ? getPassTargetFromGroup(sheet, outGroup) : '';

        const klant = String(col?.slots?.[5]?.text ?? '').trim();

        const row = [
          globalColNr,
          fase,
          isParallel ? 'Ja' : 'Nee',
          parallelWith,
          isSplit ? 'Ja' : 'Nee',
          route,
          isConditional ? 'Ja' : 'Nee',
          logicExport,
          isGroup ? 'Ja' : 'Nee',
          groupName,
          leverancier,
          sysLists.systemNames,
          sysLists.legacySystems,
          sysLists.targetSystems,
          sysLists.systemWorkarounds,
          sysLists.systemWorkaroundsNotes,
          sysLists.belemmering,
          sysLists.belemmeringNotes,
          sysLists.dubbelRegistreren,
          sysLists.dubbelRegistrerenNotes,
          sysLists.foutgevoeligheid,
          sysLists.foutgevoeligheidNotes,
          sysLists.gevolgUitval,
          sysLists.gevolgUitvalNotes,
          sysLists.ttfScores,
          inputId,
          inputText,
          inputBundlesStr,
          bundleOutIdsStr,
          bundleOutTextsStr,
          comp.result,
          comp.impact,
          comp.note,
          dq.result,
          dq.impact,
          dq.note,
          ed.result,
          ed.impact,
          ed.note,
          tj.result,
          tj.impact,
          tj.note,
          st.result,
          st.impact,
          st.note,
          ov.result,
          ov.impact,
          ov.note,
          iqfScoreStr,
          def.items,
          def.types,
          def.specs,
          proces,
          typeActiviteit,
          werkbeleving,
          toelichting,
          leanwaarde,
          statusProces,
          oorzaken,
          maatregelen,
          dis.scenarios,
          dis.frequencies,
          procesWorkarounds,
          outputId,
          outText,
          procesValidatie,
          routingRework,
          routingPass,
          klant
        ];

        lines.push(row.map(toCsvField).join(';'));
      });
    });

    const blob = new Blob(['\uFEFF' + lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
    downloadBlob(blob, getFileName('csv'));
  } catch (e) {
    console.error('CSV export error:', e);
    const msg = e && (e.message || String(e)) ? (e.message || String(e)) : 'onbekend';
    Toast.show('Fout bij genereren CSV: ' + msg, 'error', 6000);
  }
}

/* ==========================================================================
   EXPORT: HD image
   ========================================================================== */

export async function exportHD(copyToClipboard = false) {
  if (typeof html2canvas === 'undefined') {
    Toast.show('Export module niet geladen', 'error');
    return;
  }

  const board = document.getElementById('board');
  if (!board) return;

  Toast.show('Afbeelding genereren...', 'info', 2000);

  try {
    const canvas = await html2canvas(board, {
      backgroundColor: '#121619',
      scale: 2.5,
      logging: false,
      ignoreElements: (el) => el.classList.contains('col-actions'),
      onclone: (doc) => {
        doc.body.classList.add('exporting');
        // EXPORT FIX: left-align viewport for export
        const v_exp = doc.getElementById('viewport');
        if (v_exp) {
          v_exp.style.display = 'block';
          v_exp.style.width = 'fit-content';
          v_exp.style.height = 'auto';
          v_exp.style.overflow = 'visible';
          v_exp.style.justifyContent = 'flex-start';
          v_exp.style.alignItems = 'flex-start';
        }
        const bcw_exp = doc.getElementById('board-content-wrapper');
        if (bcw_exp) {
          bcw_exp.style.justifyContent = 'flex-start';
          bcw_exp.style.alignItems = 'flex-start';
        }


        // =========================================================
        // EXPORT FIX (ROBUST): rebuild lane labels as SVG (no CSS rotate)
// (html2canvas heeft bugs met CSS rotate/writing-mode => glyphs/rare letters)
(() => {
  const rh = doc.getElementById('row-headers');
  if (!rh) return;

  const win = doc.defaultView || window;

  // labels + heights uit bestaande nodes
  const rows = Array.from(rh.querySelectorAll('.row-header'));
  const labels = rows.map((r) => {
    const el = r.querySelector('.lane-label-text') || r.querySelector('span') || r;
    return String(el?.textContent || '').trim();
  });

  const heights = rows.map((r) => {
    const cssH = parseFloat((win.getComputedStyle(r).height || '').replace('px', ''));
    if (Number.isFinite(cssH) && cssH > 0) return cssH;
    const rectH = (r.getBoundingClientRect && r.getBoundingClientRect().height) ? r.getBoundingClientRect().height : 0;
    return rectH > 0 ? rectH : 160;
  });

  // container strak (minder padding/gap)
  rh.innerHTML = '';
  rh.style.width = '58px';
rh.style.minWidth = '58px';
rh.style.marginRight = '0px';
rh.style.overflow = 'visible';
  rh.style.position = 'relative';
  rh.style.zIndex = '5';

  const totalH = Math.max(1, Math.round(heights.reduce((a, b) => a + b, 0)));

  // één SVG met per rij een rotated <text> (geen CSS rotate)
  const svgNS = 'http://www.w3.org/2000/svg';
  const svg = doc.createElementNS(svgNS, 'svg');
  svg.setAttribute('width', '58');
  svg.setAttribute('height', String(totalH));
  svg.setAttribute('viewBox', `0 0 58 ${totalH}`);
  svg.style.overflow = 'visible';
  svg.style.display = 'block';

  let y = 0;
  for (let i = 0; i < labels.length; i++) {
    const h = Math.max(40, Math.round(heights[i] || 160));
    const text = labels[i] || '';

    const cx = 29;
    const cy = y + h / 2;

    const g = doc.createElementNS(svgNS, 'g');
    g.setAttribute('transform', `rotate(-90 ${cx} ${cy})`);

    const t = doc.createElementNS(svgNS, 'text');
    t.setAttribute('x', String(cx));
    t.setAttribute('y', String(cy));
    t.setAttribute('text-anchor', 'middle');
    t.setAttribute('dominant-baseline', 'middle');
    t.setAttribute('fill', 'rgba(255,255,255,0.95)');
    t.setAttribute('font-size', '24');
    t.setAttribute('font-weight', '800');
    t.setAttribute('font-family', 'ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Arial');
    t.setAttribute('letter-spacing', '0');
    t.textContent = text;

    g.appendChild(t);
    svg.appendChild(g);

    y += h;
  }

  rh.appendChild(svg);
})();

        const forceStyle = doc.createElement('style');
        forceStyle.textContent = `
          /* --- EXPORT: lane labels closer to columns --- */
          /* --- EXPORT: force left alignment (remove empty left gutter) --- */
          #viewport { justify-content: flex-start !important; align-items: flex-start !important; }
          #board-content-wrapper { justify-content: flex-start !important; align-items: flex-start !important; }
          #cols { margin-left: 0px !important; }

          #row-headers { margin-right: 0px !important; }
          .exporting-row-header{
            justify-content: flex-end !important;   /* label naar rechts in de header-cel */
            padding-right: 0px !important;          /* 0–2px, zet op 0 voor 'heel dicht' */
          }
          .exporting-row-header svg{
            overflow: visible !important;
          }

          .group-header-overlay,
          .group-header-label,
          .group-header-line {
            display: block !important;
            visibility: visible !important;
            opacity: 1 !important;
          }
          .group-header-overlay {
            position: absolute !important;
            top: 0px !important;
            z-index: 9999 !important;
            pointer-events: none !important;
          }
          .group-header-label {
            position: absolute !important;
            top: 0px !important;
            left: 0px !important;
            z-index: 10000 !important;
            pointer-events: none !important;
          }
          .group-header-line {
            position: absolute !important;
            top: 22px !important;
            left: 0px !important;
            right: 0px !important;
            z-index: 9999 !important;
            pointer-events: none !important;
          }
        `;
        doc.head.appendChild(forceStyle);

        doc.querySelectorAll('.group-header-overlay').forEach((el) => {
          el.style.display = 'block';
          el.style.visibility = 'visible';
          el.style.opacity = '1';
          el.style.position = 'absolute';
          el.style.top = '0px';
          el.style.zIndex = '9999';
          el.style.pointerEvents = 'none';
        });

        doc.querySelectorAll('.group-header-label').forEach((el) => {
          el.style.display = 'block';
          el.style.visibility = 'visible';
          el.style.opacity = '1';
          el.style.position = 'absolute';
          el.style.top = '0px';
          el.style.left = '0px';
          el.style.zIndex = '10000';
          el.style.pointerEvents = 'none';
        });

        doc.querySelectorAll('.group-header-line').forEach((el) => {
          el.style.display = 'block';
          el.style.visibility = 'visible';
          el.style.opacity = '1';
          el.style.position = 'absolute';
          el.style.top = '22px';
          el.style.left = '0px';
          el.style.right = '0px';
          el.style.zIndex = '9999';
          el.style.pointerEvents = 'none';
        });
      }
    });

    if (copyToClipboard) {
      canvas.toBlob((blob) => {
        try {
          const item = new ClipboardItem({ 'image/png': blob });
          navigator.clipboard.write([item]);
          Toast.show('Afbeelding gekopieerd naar klembord!', 'success');
        } catch {
          downloadCanvas(canvas);
          Toast.show('Klembord mislukt, afbeelding gedownload', 'info');
        }
      });
    } else {
      downloadCanvas(canvas);
    }
  } catch (err) {
    console.error('Export failed:', err);
    Toast.show('Screenshot mislukt', 'error');
  }
}

/* ==========================================================================
   GITHUB CLOUD OPSLAG
   ========================================================================== */

function utf8_to_b64(str) {
  return window.btoa(unescape(encodeURIComponent(str)));
}

function b64_to_utf8(str) {
  return decodeURIComponent(escape(window.atob(str)));
}

function getGitHubConfig() {
  return {
    token: localStorage.getItem('gh_token'),
    owner: localStorage.getItem('gh_owner'),
    repo: localStorage.getItem('gh_repo'),
    path: localStorage.getItem('gh_path') || 'ariseflow_data.json'
  };
}

// 1. LADEN VAN GITHUB
export async function loadFromGitHub() {
  const { token, owner, repo, path } = getGitHubConfig();
  if (!token || !owner || !repo) throw new Error('GitHub instellingen ontbreken.');

  const url = `https://api.github.com/repos/${owner}/${repo}/contents/${path}`;

  const response = await fetch(url, {
    headers: {
      Authorization: `token ${token}`,
      Accept: 'application/vnd.github.v3+json'
    }
  });

  if (!response.ok) throw new Error(`Fout bij laden: ${response.statusText}`);

  const data = await response.json();
  const content = b64_to_utf8(data.content);

  const parsed = JSON.parse(content);

  // ✅ Restore merge-groups naar localStorage vóór render
  restoreMergeGroupsToLocalStorage(parsed);

  // ✅ Update state (robust: state.data kan readonly zijn)
  if (typeof state.loadProjectFromObject === 'function') {
    state.loadProjectFromObject(parsed);
  } else {
    try { state.project = parsed; } catch {}
    try { state.data = parsed; } catch {}
    try { if (typeof state.notify === 'function') state.notify({ reason: 'load' }, { clone: false }); }
    catch { try { state.notify(); } catch {} }
  }

  return true;
}

// 2. OPSLAAN NAAR GITHUB (OVERWRITE)
export async function saveToGitHub() {
  const { token, owner, repo, path } = getGitHubConfig();

  if (!token || !owner || !repo) {
    alert("Vul eerst je GitHub gegevens in via de knop 'Setup'.");
    return;
  }

  const url = `https://api.github.com/repos/${owner}/${repo}/contents/${path}`;

  let sha = null;
  try {
    const getResp = await fetch(url, {
      method: 'GET',
      headers: {
        Authorization: `token ${token}`,
        Accept: 'application/vnd.github.v3+json'
      }
    });
    if (getResp.ok) {
      const getData = await getResp.json();
      sha = getData.sha;
    }
  } catch {
    console.warn('Bestand bestaat nog niet, er wordt een nieuwe gemaakt.');
  }

  // ✅ Project voorbereiden: merge-groups + outputUids mee saven
  const p = prepareProjectForPersist(state.data || state.project);

  const contentStr = JSON.stringify(p, null, 2);
  const body = {
    message: `Update via AriseFlow: ${new Date().toLocaleString()}`,
    content: utf8_to_b64(contentStr),
    sha: sha
  };

  const putResp = await fetch(url, {
    method: 'PUT',
    headers: {
      Authorization: `token ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(body)
  });

  if (!putResp.ok) {
    const errData = await putResp.json();
    throw new Error(`Opslaan mislukt: ${errData.message}`);
  }
}