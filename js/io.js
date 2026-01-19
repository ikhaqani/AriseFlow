// io.js (AANGEPAST: FIX VOOR AFGEKAPTE GROEP TITELS BIJ EXPORT)

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
  const title = state.data?.projectTitle || 'sipoc_project';
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

function isContiguousZeroBased(cols) {
  if (!Array.isArray(cols) || cols.length < 2) return false;
  const s = [...new Set(cols)].sort((a, b) => a - b);
  return s.length === s[s.length - 1] - s[0] + 1;
}

function sanitizeMergeGroupForSheet(sheet, g) {
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

  const gate =
    slotIdx === 4 && g?.gate && typeof g.gate === 'object'
      ? {
          enabled: !!g.gate.enabled,
          failTargetColIdx: Number.isFinite(Number(g.gate.failTargetColIdx))
            ? Number(g.gate.failTargetColIdx)
            : null
        }
      : null;

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
  if (!gate?.enabled) return '';
  const idx = gate.failTargetColIdx;
  if (idx == null || !Number.isFinite(Number(idx))) return '';
  return getProcessLabel(sheet, idx);
}

/* ==========================================================================
   Variant route letters
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

  if (Array.isArray(sheet.variantGroups)) {
      sheet.variantGroups.forEach(vg => {
          vg.variants.forEach((vIdx, i) => {
              map[vIdx] = toLetter(i);
          });
      });
  }

  let inRun = false;
  let runIdx = 0;

  for (let i = 0; i < sheet.columns.length; i++) {
    if (map[i]) {
        inRun = false;
        runIdx = 0;
        continue;
    }

    const col = sheet.columns[i];
    if (col?.isVisible === false) continue;

    const isVar = !!col?.isVariant;
    if (isVar) {
      if (!inRun) {
        inRun = true;
        runIdx = 0;
      }
      map[i] = toLetter(runIdx);
      runIdx += 1;
    } else {
      inRun = false;
      runIdx = 0;
    }
  }

  return map;
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

  // dedupe while preserving order
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
   Verwacht in project: project.outputBundles = [{ id, name, outIds: [] }]
   Verwacht in inputSlot: linkedBundleId (string) of linkedBundleIds (array)
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

  // dedupe while preserving order
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
  buildGlobalOutputMaps(project); // no-op safety (ensures uids exist)
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

    // NEW schema: outIds (OUT1..), legacy: outputUids
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

/**
 * Checks if a meta object actually contains valid system data.
 * Returns true if at least one system has a name.
 */
function hasValidSystems(meta) {
  if (!meta || typeof meta !== 'object') return false;
  if (!Array.isArray(meta.systems) || meta.systems.length === 0) return false;
  
  // Check if at least one system has a non-empty name
  return meta.systems.some(s => s && String(s.name || '').trim() !== '');
}

/**
 * ROBUST FALLBACK: Builds system meta from slot data, ensuring we don't
 * return empty structures that override the post-it text.
 */
function buildSystemsMetaFallbackFromSlot(sysSlot) {
  const slot = sysSlot && typeof sysSlot === 'object' ? sysSlot : {};
  const sd = slot.systemData && typeof slot.systemData === 'object' ? slot.systemData : null;
  const postItText = String(slot.text || '').trim();

  // 1) Try explicit new metadata (sd.systemsMeta)
  if (sd?.systemsMeta && hasValidSystems(sd.systemsMeta)) {
    return sd.systemsMeta;
  }

  // 2) Try legacy metadata (sd.systems)
  if (Array.isArray(sd?.systems) && sd.systems.length > 0) {
    const validLegacy = sd.systems.filter(s => s && String(s.name || '').trim() !== '');
    if (validLegacy.length > 0) {
      return { 
        multi: validLegacy.length > 1, 
        systems: validLegacy 
      };
    }
  }

  // 3) Try simple legacy name (sd.systemName)
  const nameFromSd = String(sd?.systemName ?? '').trim();
  if (nameFromSd) {
    return {
      multi: false,
      systems: [{ name: nameFromSd, legacy: false, future: '', qa: {}, score: null }]
    };
  }

  // 4) Ultimate Fallback: The Post-it Text itself
  if (postItText) {
    return {
      multi: false,
      systems: [{ name: postItText, legacy: false, future: '', qa: {}, score: null }]
    };
  }

  // No data found at all
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
    // Note: We don't filter out empty names here yet, because the UI might be 
    // mid-edit. But for export, we prefer valid names.

  const inferredMulti = systems.length > 1;
  const multi = !!meta.multi || inferredMulti;

  if (systems.length === 0) systems.push({ name: '', legacy: false, future: '', qa: {}, score: null });
  return { multi, systems };
}

function computeTTFSystemScore(sys) {
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

function getSysAnswerLabel(sys, qid) {
  const qa = sys?.qa && typeof sys.qa === 'object' ? sys.qa : {};
  const q = SYSFIT_Q.find((x) => x.id === qid);
  if (!q) return '';
  const key = qa[qid];
  if (!key) return '';
  const opt = (SYSFIT_OPTS[q.type] || []).find((o) => o.key === key);
  return opt?.label || '';
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
  // Ensure we process the clean result
  const systems = (clean?.systems || []).filter(s => s && String(s.name || '').trim() !== '');

  // If after cleaning we have no systems, allow one empty to prevent crashes,
  // but logically this should have been caught by the export loop logic.
  if (systems.length === 0 && clean?.systems?.length > 0) {
     // If truly empty, return empty strings
     return {
         systemNames: '',
         legacySystems: '',
         targetSystems: '',
         systemWorkarounds: '',
         belemmering: '',
         dubbelRegistreren: '',
         foutgevoeligheid: '',
         gevolgUitval: '',
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

  const ttfScores = computeTTFScoreListFromMeta({ ...clean, systems }).map((v) =>
    Number.isFinite(Number(v)) ? `${Number(v)}%` : '—'
  );

  const sysWorkarounds = systems.map((s) => getSysAnswerLabel(s, 'q1'));
  const belemmering = systems.map((s) => getSysAnswerLabel(s, 'q2'));
  const dubbelReg = systems.map((s) => getSysAnswerLabel(s, 'q3'));
  const foutgevoelig = systems.map((s) => getSysAnswerLabel(s, 'q4'));
  const uitvalImpact = systems.map((s) => getSysAnswerLabel(s, 'q5'));

  return {
    systemNames: joinSemi(names),
    legacySystems: joinSemi(legacyNames),
    targetSystems: joinSemi(targetNames),
    systemWorkarounds: joinSemi(sysWorkarounds),
    belemmering: joinSemi(belemmering),
    dubbelRegistreren: joinSemi(dubbelReg),
    foutgevoeligheid: joinSemi(foutgevoelig),
    gevolgUitval: joinSemi(uitvalImpact),
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
  if (U === 'NOT_OK' || U === 'FAIL' || U === 'NOK' || U === 'VOLDOET_NIET') return 'Voldoet niet';
  if (U === 'NVT' || U === 'NA') return 'NVT';
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
  const impactArr = extractPerSystem(q, systemsCount, 'impact');
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
   JSON save/load
   ========================================================================== */

export async function saveToFile() {
  const dataStr = JSON.stringify(state.data, null, 2);
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

      state.project = parsed;
      if (typeof state.notify === 'function') state.notify();
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
      'Logica', // <--- NIEUW: Logica kolom
      'Groep?',
      'Groepsnaam', // <--- NIEUW: Groepsnaam kolom
      'Leverancier',
      'Systemen',
      'Legacy systemen',
      'Target systemen',
      'Systeem workarounds',
      'Belemmering',
      'Dubbel registreren',
      'Foutgevoeligheid',
      'Gevolg bij uitval',
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

    const project = state.data;

    // FIX: bouw eerst 1x project-brede OUT mapping + output teksten
    const { outIdByUid, outTextByUid, outTextByOutId } = buildGlobalOutputMaps(project);
    const bundleMaps = buildBundleMaps(project, outIdByUid, outTextByUid, outTextByOutId);

    let globalColNr = 0;
    let globalInCounter = 0;

    (project?.sheets || []).forEach((sheet) => {
      const mergeGroups = getMergeGroupsSanitized(project, sheet);
      
      // PRE-CALCULATE GLOBAL COLUMN NUMBERS FOR THIS SHEET
      // globalColNr is currently the count of *previous* sheets' columns.
      // We need to simulate the loop to map local indices to future global numbers.
      const localToGlobalMap = {};
      let tempGlobalCounter = globalColNr; // Start where previous sheet left off

      (sheet.columns || []).forEach((c, idx) => {
          if (c.isVisible !== false) {
              tempGlobalCounter++;
              localToGlobalMap[idx] = tempGlobalCounter;
          }
      });

      // === NIEUWE LOGICA VOOR ROUTE LABELS (GLOBAL NUMBERS) ===
      const routeLookup = {};
      if (Array.isArray(sheet.variantGroups)) {
          sheet.variantGroups.forEach(vg => {
              // Naam van parent kolom (gebruik global nummer als backup)
              const parentCol = sheet.columns[vg.parentColIdx];
              const parentName = parentCol?.slots?.[3]?.text || `Kolom ${localToGlobalMap[vg.parentColIdx] || '?'}`;
              
              // === AANGEPASTE REGEL: Map local variant indices naar GLOBALE kolomnummers ===
              const routeNums = vg.variants
                  .map(v => localToGlobalMap[v]) // Get global number
                  .filter(n => n) // Filter out undefined (hidden cols)
                  .join(', ');
                  
              routeLookup[vg.parentColIdx] = `Main (Routes: ${routeNums})`;

              vg.variants.forEach((vIdx, i) => {
                  const letter = String.fromCharCode(65 + i); 
                  routeLookup[vIdx] = `Route ${letter} (van: ${parentName})`;
              });
          });
      }
      // ========================================

      // Fallback voor oude varianten (losse split knop)
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

        // ============================
        // SYSTEMS (ROBUUSTE FIX)
        // ============================
        const sysSlot = col?.slots?.[1] || {};
        const sysGroup = getMergeGroupForCell(mergeGroups, colIdx, 1);
        
        let sysMeta = null;

        // 1. Check Merge Group Meta: ONLY use if it has actual names
        if (sysGroup?.systemsMeta && hasValidSystems(sysGroup.systemsMeta)) {
             sysMeta = sysGroup.systemsMeta;
        }

        // 2. Fallback to Slot Logic (Explicit meta -> Legacy meta -> Post-it Text)
        if (!sysMeta) {
             sysMeta = buildSystemsMetaFallbackFromSlot(sysSlot);
        }

        const sysLists = systemsToLists(sysMeta);
        const systemsCount = sysLists.systemsCount;

        // Parallel / Split / Route
        const isParallel = !!col.isParallel;
        const prevIdx = prevVisibleByIdx[colIdx];
        const parallelWith = isParallel && prevIdx != null ? getProcessLabel(sheet, prevIdx) : '-';

        const isSplit = !!col.isVariant;
        
        // AANGEPAST: Haal route info uit lookup of fallback map
        const route = routeLookup[colIdx] || (isSplit ? (variantMap[colIdx] || 'Onbekende route') : '-');

        // NIEUW: Conditioneel (Trigger) & Logica
        const isConditional = !!col.isConditional;
        
        // --- LOGICA EXTRACTIE VOOR CSV ---
        const logic = col.logic || {};
        let logicExport = '';

        if (isConditional && logic.condition) {
            const trueAction = logic.ifTrue !== null ? `Ga naar ${getProcessLabel(sheet, logic.ifTrue)}` : 'Voer stap uit';
            const falseAction = logic.ifFalse !== null ? `Ga naar ${getProcessLabel(sheet, logic.ifFalse)}` : 'Voer stap uit';

            logicExport = `VRAAG: ${logic.condition}`;
            logicExport += `; INDIEN JA: ${trueAction}`;
            logicExport += `; INDIEN NEE: ${falseAction}`;
        }
        // ---------------------------------
        
        // NIEUW: Group
        const isGroup = !!col.isGroup;
        const groupName = state.getGroupForCol(colIdx)?.title || '';

        // Fase
        const fase = `Procesflow ${globalColNr}`;

        // Input + InputID (linked of nieuw) — ondersteunt single + multiple links + bundels
        const inputSlot = col?.slots?.[2];
        const outputSlot = col?.slots?.[4];

        let inputId = '';
        let inputText = String(inputSlot?.text ?? '').trim();

        // 1) bundels (compact)
        const linkedBundleIds = normalizeLinkedBundles(inputSlot);
        const bundleResolved = resolveBundleIdsToLists(linkedBundleIds, bundleMaps);

        const inputBundlesStr = joinSemi(bundleResolved.bundleNames);
        const bundleOutIdsStr = joinSemi(bundleResolved.memberOutIds);
        const bundleOutTextsStr = joinSemi(bundleResolved.memberOutTexts);

        // 2) losse OUT-links (single + multiple)
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

        // IO QA (Input slot qa) per systeem aligned
        const slotQa = inputSlot?.qa || {};

        const comp = getIOTripleForLabel(slotQa, 'Compleetheid', systemsCount);
        const dq = getIOTripleForLabel(slotQa, 'Datakwaliteit', systemsCount);
        const ed = getIOTripleForLabel(slotQa, 'Eenduidigheid', systemsCount);
        const tj = getIOTripleForLabel(slotQa, 'Tijdigheid', systemsCount);
        const st = getIOTripleForLabel(slotQa, 'Standaardisatie', systemsCount);
        const ov = getIOTripleForLabel(slotQa, 'Overdracht', systemsCount);

        const iqfScore = calculateIQFScore(slotQa);
        const iqfScoreStr = iqfScore == null ? '' : String(iqfScore);

        // Input definitions -> items/types/specs
        const def = splitDefs(inputSlot?.inputDefinitions);

        // Proces
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

        // Output + OutputID (stabiel via outputUid->OUTx) + Gate/routing
        const outGroup = getMergeGroupForCell(mergeGroups, colIdx, 4);
        const outText = String(outputSlot?.text ?? '').trim();

        let outputId = '';
        if (outText && !isMergedSlaveInSheet(mergeGroups, colIdx, 4)) {
          const outUid = String(outputSlot?.outputUid || '').trim();
          outputId = outUid && outIdByUid[outUid] ? outIdByUid[outUid] : '';
        }

        const procesValidatie = outGroup?.gate?.enabled ? 'Ja' : 'Nee';
        const routingRework = outGroup?.gate?.enabled ? getFailTargetFromGate(sheet, outGroup.gate) || '-' : '-';
        const routingPass = outGroup ? getPassTargetFromGroup(sheet, outGroup) || '-' : '-';

        // Klant
        const klant = String(col?.slots?.[5]?.text ?? '').trim();

        const row = [
          globalColNr,
          fase,
          isParallel ? 'Ja' : 'Nee',
          parallelWith,
          isSplit ? 'Ja' : 'Nee',
          route,
          isConditional ? 'Ja' : 'Nee',
          logicExport, // <--- Logica data
          isGroup ? 'Ja' : 'Nee',
          groupName, // <--- Groepsnaam

          leverancier,

          sysLists.systemNames,
          sysLists.legacySystems,
          sysLists.targetSystems,
          sysLists.systemWorkarounds,
          sysLists.belemmering,
          sysLists.dubbelRegistreren,
          sysLists.foutgevoeligheid,
          sysLists.gevolgUitval,
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
    console.error(e);
    Toast.show('Fout bij genereren CSV', 'error');
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
        const v = doc.getElementById('viewport');
        const b = doc.getElementById('board');

        if (v) {
          v.style.overflow = 'visible';
          v.style.width = 'fit-content';
          v.style.height = 'auto';
          v.style.position = 'static';
        }

        if (b) {
            b.style.transform = 'none'; 
            b.style.marginTop = '80px'; 
            b.style.padding = '20px';   
        }
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

// Helpers voor tekst-codering (nodig voor GitHub API)
function utf8_to_b64(str) {
  return window.btoa(unescape(encodeURIComponent(str)));
}

function b64_to_utf8(str) {
  return decodeURIComponent(escape(window.atob(str)));
}

// Haal instellingen op
function getGitHubConfig() {
  return {
    token: localStorage.getItem('gh_token'),
    owner: localStorage.getItem('gh_owner'),
    repo: localStorage.getItem('gh_repo'),
    path: localStorage.getItem('gh_path') || 'ariseflow_data.json' // Naam van bestand in je repo
  };
}

// 1. LADEN VAN GITHUB
export async function loadFromGitHub() {
  const { token, owner, repo, path } = getGitHubConfig();
  if (!token || !owner || !repo) throw new Error('GitHub instellingen ontbreken.');

  // URL naar het bestand via de API
  const url = `https://api.github.com/repos/${owner}/${repo}/contents/${path}`;
  
  const response = await fetch(url, {
    headers: {
      'Authorization': `token ${token}`,
      'Accept': 'application/vnd.github.v3+json'
    }
  });

  if (!response.ok) throw new Error(`Fout bij laden: ${response.statusText}`);

  const data = await response.json();
  const content = b64_to_utf8(data.content); // Decodeer de inhoud
  
  // Update de applicatie state
  const parsed = JSON.parse(content);
  state.project = parsed;
  if (typeof state.notify === 'function') state.notify();
  
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
  
  // STAP A: Haal eerst de huidige SHA (versie-code) op van het bestand online
  let sha = null;
  try {
    const getResp = await fetch(url, {
      method: 'GET',
      headers: { 
        'Authorization': `token ${token}`,
        'Accept': 'application/vnd.github.v3+json'
      }
    });
    if (getResp.ok) {
      const getData = await getResp.json();
      sha = getData.sha; // Dit is de 'sleutel' om te mogen overschrijven
    }
  } catch (e) {
    console.warn("Bestand bestaat nog niet, er wordt een nieuwe gemaakt.");
  }

  // STAP B: Bereid de nieuwe data voor
  const contentStr = JSON.stringify(state.data, null, 2);
  const body = {
    message: `Update via AriseFlow: ${new Date().toLocaleString()}`,
    content: utf8_to_b64(contentStr),
    sha: sha // Als we deze SHA meesturen, weet GitHub dat we deze specifieke versie overschrijven
  };

  // STAP C: Stuur de update (PUT request)
  const putResp = await fetch(url, {
    method: 'PUT',
    headers: {
      'Authorization': `token ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(body)
  });

  if (!putResp.ok) {
    const errData = await putResp.json();
    throw new Error(`Opslaan mislukt: ${errData.message}`);
  }
}