// io.js (VOLLEDIG - EXPORT AS-IS CSV + JSON + HD)
// - 1 rij per proceskolom (geen 6 rijen per SSIPOC-slot)
// - Alle multi-waarden in 1 cel gescheiden door ";" (aligned per systeem waar van toepassing)
// - Disruptions/oorzaken/maatregelen/toelichtingen ook ";"-gescheiden
// - FIX: Input kan gelinkt zijn aan OUTx via linkedSourceId óf linkedSourceUid óf MEERDERE links (arrays)
//        -> alles wordt als "OUTx; OUTy" geëxporteerd (en input-tekst als "tekstOutx; tekstOuty")
// - FIX: Input kan ook een bundel zijn (bundelnaam -> OUTx; OUTy) voor compacte input in post-it/export
// - FIX: Input-tekst bij link gebruikt output-tekst (outTextByUid / outTextByOutId) indien beschikbaar
// - FIX: Output IDs zijn project-breed stabiel (op basis van outputUid + sheet/kolom volgorde + merge-slaves overslaan)
// - FIX (NIEUW): Systeem export werkt ook als je systeemnaam direct in de post-it hebt getypt (slot.text),
//                dus zonder “multi-systeem” toggle / systemsMeta.

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

  let inRun = false;
  let runIdx = 0;

  for (let i = 0; i < sheet.columns.length; i++) {
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

// NEW: bouw een bruikbare meta-structuur als de gebruiker alleen slot.text heeft ingevuld.
function buildSystemsMetaFallbackFromSlot(sysSlot) {
  const slot = sysSlot && typeof sysSlot === 'object' ? sysSlot : {};
  const sd = slot.systemData && typeof slot.systemData === 'object' ? slot.systemData : null;

  // 1) Nieuwe vorm: sd.systemsMeta.systems
  const meta = sd?.systemsMeta && typeof sd.systemsMeta === 'object' ? sd.systemsMeta : null;
  if (Array.isArray(meta?.systems)) {
    return {
      multi: !!meta.multi || meta.systems.length > 1,
      systems: meta.systems.map((s) => ({
        name: String(s?.name ?? '').trim(),
        legacy: !!(s?.legacy ?? s?.isLegacy),
        future: String(s?.future ?? s?.futureSystem ?? '').trim(),
        qa: s?.qa && typeof s.qa === 'object' ? { ...s.qa } : {},
        score: Number.isFinite(Number(s?.score ?? s?.calculatedScore)) ? Number(s?.score ?? s?.calculatedScore) : null
      }))
    };
  }

  // 2) Legacy vorm: sd.systems (uit oudere tabs)
  if (Array.isArray(sd?.systems)) {
    const arr = sd.systems
      .map((s) => ({
        name: String(s?.name ?? '').trim(),
        legacy: !!(s?.legacy ?? s?.isLegacy),
        future: String(s?.future ?? s?.futureSystem ?? '').trim(),
        qa: s?.qa && typeof s.qa === 'object' ? { ...s.qa } : {},
        score: Number.isFinite(Number(s?.score ?? s?.calculatedScore)) ? Number(s?.score ?? s?.calculatedScore) : null
      }))
      .filter((x) => x.name || x.future || x.legacy || Object.keys(x.qa || {}).length);

    if (arr.length) return { multi: arr.length > 1, systems: arr };
  }

  // 3) Enkel systeem uit sd.systemName
  const nameFromSd = String(sd?.systemName ?? '').trim();
  if (nameFromSd) {
    return {
      multi: false,
      systems: [{ name: nameFromSd, legacy: false, future: '', qa: {}, score: null }]
    };
  }

  // 4) Cruciale fallback: direct getypte post-it tekst (slot.text)
  const nameFromText = String(slot?.text ?? '').trim();
  if (nameFromText) {
    return {
      multi: false,
      systems: [{ name: nameFromText, legacy: false, future: '', qa: {}, score: null }]
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
  const systems = clean?.systems || [];

  const names = systems.map((s) => String(s?.name || '').trim()).filter(Boolean);
  const legacyNames = systems
    .filter((s) => !!s.legacy)
    .map((s) => String(s?.name || '').trim())
    .filter(Boolean);
  const targetNames = systems
    .filter((s) => !!s.legacy)
    .map((s) => String(s?.future || '').trim())
    .filter(Boolean);

  const ttfScores = computeTTFScoreListFromMeta(clean).map((v) =>
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
        // SYSTEMS (FIX: fallback slot.text)
        // ============================
        const sysSlot = col?.slots?.[1] || {};
        const sysGroup = getMergeGroupForCell(mergeGroups, colIdx, 1);

        // 1) bij system-merge: group.systemsMeta heeft prioriteit
        // 2) anders: uit slot.systemData (systemsMeta/systems/systemName)
        // 3) anders: slot.text (direct getypte post-it)
        const sysMeta =
          sysGroup?.systemsMeta ||
          sysSlot?.systemData?.systemsMeta ||
          buildSystemsMetaFallbackFromSlot(sysSlot);

        const sysLists = systemsToLists(sysMeta);
        const systemsCount = sysLists.systemsCount;

        // Parallel / Split / Route
        const isParallel = !!col.isParallel;
        const prevIdx = prevVisibleByIdx[colIdx];
        const parallelWith = isParallel && prevIdx != null ? getProcessLabel(sheet, prevIdx) : '-';

        const isSplit = !!col.isVariant;
        const route = isSplit ? String(variantMap[colIdx] || '') : '-';

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
        if (v) {
          v.style.overflow = 'visible';
          v.style.width = 'fit-content';
          v.style.height = 'auto';
          v.style.padding = '40px';
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