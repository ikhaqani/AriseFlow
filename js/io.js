// io.js (VOLLEDIG - AANGEPAST)
// -----------------------------------------------------------
import { state } from './state.js';
import { Toast } from './toast.js';
import { IO_CRITERIA } from './config.js';

const MERGE_LS_PREFIX = 'ssipoc.mergeGroups.v2';

/* =========================================================
   CSV helpers
   ========================================================= */

function toCsvField(text) {
  if (text === null || text === undefined) return '""';
  const str = String(text);
  return `"${str.replace(/"/g, '""').replace(/\r?\n/g, ' ')}"`;
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

/* =========================================================
   Scoring / formatting
   ========================================================= */

function calculateLSSScore(qa) {
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

function yesNo(v) {
  return v ? 'Ja' : 'Nee';
}

function safeText(v) {
  return String(v ?? '').trim();
}

function formatInputSpecs(defs) {
  if (!Array.isArray(defs) || defs.length === 0) return '';
  return defs
    .filter((d) => d && (d.item || d.specifications || d.type))
    .map((d) => {
      const item = safeText(d.item);
      const specs = safeText(d.specifications);
      const type = safeText(d.type);
      const main = [item, specs].filter(Boolean).join(': ');
      return type ? `${main} (${type})` : main;
    })
    .filter(Boolean)
    .join(' | ');
}

function formatDisruptions(disruptions) {
  if (!Array.isArray(disruptions) || disruptions.length === 0) return '';
  return disruptions
    .filter((d) => d && (d.scenario || d.frequency || d.workaround))
    .map((d) => {
      const scenario = safeText(d.scenario);
      const freq = safeText(d.frequency);
      const workaround = safeText(d.workaround);
      const left = [scenario, freq].filter(Boolean).join(' - ');
      return workaround ? `${left}: ${workaround}` : left;
    })
    .filter(Boolean)
    .join(' | ');
}

function formatWorkjoy(slot) {
  const v = slot?.workExp ?? slot?.workjoy ?? slot?.workJoy ?? null;
  if (!v) return { value: '', label: '', icon: '', context: '' };

  const V = String(v);
  if (V === 'OBSTACLE') return { value: V, label: 'Obstakel', icon: '🛠️', context: 'Kost energie & frustreert.' };
  if (V === 'ROUTINE') return { value: V, label: 'Routine', icon: '🤖', context: 'Saai & repeterend.' };
  if (V === 'FLOW') return { value: V, label: 'Flow', icon: '🚀', context: 'Geeft energie & voldoening.' };
  return { value: V, label: '', icon: '', context: '' };
}

/* =========================================================
   Merge helpers (localStorage)
   ========================================================= */

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

function isContiguous(cols) {
  if (!Array.isArray(cols) || cols.length < 2) return false;
  const s = [...new Set(cols)].sort((a, b) => a - b);
  return s.length === s[s.length - 1] - s[0] + 1;
}

function sanitizeMergeGroup(g, colCount) {
  const slotIdx = Number(g?.slotIdx);
  if (![1, 4].includes(slotIdx)) return null;

  const cols = Array.isArray(g?.cols) ? g.cols.map((x) => Number(x)).filter(Number.isFinite) : [];
  const uniq = [...new Set(cols)].filter((c) => c >= 0 && c < colCount);
  if (uniq.length < 2) return null;
  if (!isContiguous(uniq)) return null;

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

  const systemsMeta = slotIdx === 1 && g?.systemsMeta && typeof g.systemsMeta === 'object' ? g.systemsMeta : null;

  return { slotIdx, cols: uniq.sort((a, b) => a - b), master, gate, systemsMeta };
}

function getMergeGroupsSanitized(project, sheet) {
  const n = sheet?.columns?.length ?? 0;
  if (!n) return [];
  return loadMergeGroupsRaw(project, sheet)
    .map((g) => sanitizeMergeGroup(g, n))
    .filter(Boolean);
}

function getMergeGroupForCell(groups, colIdx, slotIdx) {
  return (groups || []).find((x) => x.slotIdx === slotIdx && Array.isArray(x.cols) && x.cols.includes(colIdx)) || null;
}

function getMergeRangeString(group) {
  if (!group?.cols?.length) return '';
  const min = Math.min(...group.cols) + 1;
  const max = Math.max(...group.cols) + 1;
  return min === max ? String(min) : `${min}-${max}`;
}

/* =========================================================
   Routing helpers
   ========================================================= */

function getProcessLabel(sheet, colIdx) {
  const t = sheet?.columns?.[colIdx]?.slots?.[3]?.text;
  const s = safeText(t);
  return s || `Kolom ${Number(colIdx) + 1}`;
}

function getNextVisibleColIdx(sheet, fromIdx) {
  const n = sheet?.columns?.length ?? 0;
  for (let i = fromIdx + 1; i < n; i++) {
    if (sheet.columns[i]?.isVisible !== false) return i;
  }
  return null;
}

function getPassTargetFromGroup(sheet, group) {
  if (!group?.cols?.length) return { label: '', col: '' };
  const maxCol = Math.max(...group.cols);
  const nextIdx = getNextVisibleColIdx(sheet, maxCol);
  if (nextIdx == null) return { label: 'Einde proces', col: '' };
  return { label: getProcessLabel(sheet, nextIdx), col: String(nextIdx + 1) };
}

function getFailTargetFromGate(sheet, gate) {
  if (!gate?.enabled) return { label: '', col: '' };
  const idx = gate.failTargetColIdx;
  if (idx == null || !Number.isFinite(Number(idx))) return { label: '', col: '' };
  return { label: getProcessLabel(sheet, idx), col: String(Number(idx) + 1) };
}

/* =========================================================
   Variant helpers
   ========================================================= */

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

/* =========================================================
   Systems meta formatting (multi-system)
   ========================================================= */

function sanitizeSystemsMeta(meta) {
  if (!meta || typeof meta !== 'object') return null;

  const arr = Array.isArray(meta.systems) ? meta.systems : [];
  const systems = arr
    .map((s) => {
      if (!s || typeof s !== 'object') return null;
      return {
        name: safeText(s.name),
        legacy: !!s.legacy,
        future: safeText(s.future),
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

function formatSystemsList(meta) {
  const clean = sanitizeSystemsMeta(meta);
  if (!clean) return { names: '', scores: '', legacyFuture: '', answersJson: '' };

  const names = clean.systems
    .map((s) => {
      const nm = safeText(s.name);
      if (!nm) return '';
      return s.legacy ? `${nm} (Legacy)` : nm;
    })
    .filter(Boolean)
    .join(' | ');

  const scores = clean.systems.map((s) => (Number.isFinite(Number(s.score)) ? `${Number(s.score)}%` : '—')).join('; ');

  const legacyFuture = clean.systems
    .filter((s) => s.legacy && (safeText(s.name) || safeText(s.future)))
    .map((s) => `${safeText(s.name) || '—'} -> ${safeText(s.future) || '—'}`)
    .join(' | ');

  const answersJson = (() => {
    const payload = clean.systems.map((s) => ({
      name: s.name || '',
      legacy: !!s.legacy,
      future: s.future || '',
      score: Number.isFinite(Number(s.score)) ? Number(s.score) : null,
      qa: s.qa || {}
    }));
    try {
      return JSON.stringify(payload);
    } catch {
      return '';
    }
  })();

  return { names, scores, legacyFuture, answersJson };
}

/* =========================================================
   Output ID map (OUT1..)
   ========================================================= */

function buildOutputMaps(sheet, mergeGroups) {
  const outIdBySlotId = {};
  const outTextByOutId = {};
  let outCounter = 0;

  sheet.columns.forEach((col, colIdx) => {
    const sOut = col?.slots?.[4];
    if (!sOut?.text?.trim()) return;

    const g = getMergeGroupForCell(mergeGroups, colIdx, 4);
    if (g && colIdx !== g.master) return;

    outCounter += 1;
    const outId = `OUT${outCounter}`;
    if (sOut.id) outIdBySlotId[sOut.id] = outId;
    outTextByOutId[outId] = sOut.text;
  });

  return { outIdBySlotId, outTextByOutId };
}

/* =========================================================
   File save/load
   ========================================================= */

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

/* =========================================================
   EXPORT 1: AsIs_Overview (1 rij per processtap)
   ========================================================= */

function getSlotText(col, slotIdx, outTextByOutId) {
  const slot = col?.slots?.[slotIdx];
  if (!slot) return '';

  // Input linked -> use linked output text
  if (slotIdx === 2 && slot?.linkedSourceId) {
    const linkedId = slot.linkedSourceId;
    if (outTextByOutId && outTextByOutId[linkedId]) return safeText(outTextByOutId[linkedId]);
  }

  return safeText(slot?.text);
}

function computeIoScoreForCol(col, slotIdx) {
  const slot = col?.slots?.[slotIdx];
  return slot ? calculateLSSScore(slot.qa) : null;
}

function getSystemOverallForCol(col) {
  const sd = col?.slots?.[1]?.systemData;
  if (!sd) return null;
  const v = sd.calculatedScore ?? null;
  return Number.isFinite(Number(v)) ? Number(v) : null;
}

function getSystemsMetaForCol(col) {
  const sd = col?.slots?.[1]?.systemData;
  if (!sd) return null;
  return sd.systemsMeta ?? null;
}

function buildAsIsOverviewRows(project) {
  const headers = [
    'Project',
    'Sheet',
    'Step Nr (Global)',
    'Step Nr (Sheet)',
    'Column Id',
    'Step Label',

    'Supplier',
    'System Text',
    'Input',
    'Process',
    'Output',
    'Customer',

    'Activity Type',
    'Process Status',
    'Lean Value',
    'WorkExp',
    'WorkExp Note',

    'TTF Overall (%)',
    'IQF Input (%)',
    'IQF Output (%)',

    'Split (Variant)',
    'Variant Route',
    'Parallel',

    'System Merged',
    'System Merge Range',

    'Output Merged',
    'Output Merge Range',
    'Gate Enabled',
    'Gate Fail To (Label)',
    'Gate Fail To (Col)',
    'Gate Pass To (Label)',
    'Gate Pass To (Col)',

    'Systemen (met Legacy)',
    'TTF Scores (per systeem)',
    'Legacy -> Toekomst',
    'TTF Antwoorden (JSON)',

    'Succesfactoren',
    'Oorzaken (RC)',
    'Maatregelen (CM)',
    'Verstoringen',
    'Input Specs',
    'Output Specs'
  ];

  const lines = [headers.map(toCsvField).join(';')];

  let globalStepNr = 0;

  (project?.sheets || []).forEach((sheet) => {
    const mergeGroups = getMergeGroupsSanitized(project, sheet);
    const variantMap = computeVariantLetterMap(sheet);
    const { outTextByOutId } = buildOutputMaps(sheet, mergeGroups);

    let sheetStepNr = 0;

    (sheet.columns || []).forEach((col, colIdx) => {
      if (col?.isVisible === false) return;

      sheetStepNr += 1;
      globalStepNr += 1;

      const sysGroup = getMergeGroupForCell(mergeGroups, colIdx, 1);
      const outGroup = getMergeGroupForCell(mergeGroups, colIdx, 4);

      const sysMerged = !!sysGroup;
      const sysRange = sysGroup ? getMergeRangeString(sysGroup) : '';

      const outMerged = !!outGroup;
      const outRange = outGroup ? getMergeRangeString(outGroup) : '';

      const gateEnabled = outGroup?.gate?.enabled ? true : false;
      const failTarget = outGroup ? getFailTargetFromGate(sheet, outGroup.gate) : { label: '', col: '' };
      const passTarget = outGroup ? getPassTargetFromGroup(sheet, outGroup) : { label: '', col: '' };

      const supplierText = getSlotText(col, 0, outTextByOutId);
      const systemText = getSlotText(col, 1, outTextByOutId);
      const inputText = getSlotText(col, 2, outTextByOutId);
      const processText = getSlotText(col, 3, outTextByOutId);
      const outputText = getSlotText(col, 4, outTextByOutId);
      const customerText = getSlotText(col, 5, outTextByOutId);

      const procSlot = col?.slots?.[3] || {};
      const activityType = safeText(procSlot.type);
      const processStatus = safeText(procSlot.processStatus);
      const leanValue = safeText(procSlot.processValue);

      const wj = formatWorkjoy(procSlot);
      const workExpNote = safeText(procSlot.workExpNote);

      const ttfOverall = getSystemOverallForCol(col);
      const sysFmt = formatSystemsList(getSystemsMetaForCol(col));

      const iqfInput = computeIoScoreForCol(col, 2);
      const iqfOutput = computeIoScoreForCol(col, 4);

      const splitVariant = yesNo(!!col.isVariant);
      const variantRoute = col.isVariant ? String(variantMap[colIdx] || '') : '';
      const parallel = yesNo(!!col.isParallel);

      const succesfactoren = safeText(procSlot.successFactors);
      const oorzaken = Array.isArray(procSlot.causes) ? procSlot.causes.map(safeText).filter(Boolean).join(' | ') : '';
      const maatregelen = Array.isArray(procSlot.improvements)
        ? procSlot.improvements.map(safeText).filter(Boolean).join(' | ')
        : '';
      const verstoringen = formatDisruptions(procSlot.disruptions);

      const inputSpecs = formatInputSpecs(col?.slots?.[2]?.inputDefinitions);
      const outputSpecs = formatInputSpecs(col?.slots?.[4]?.inputDefinitions);

      const row = [
        safeText(project?.projectTitle || ''),
        safeText(sheet?.name || ''),
        globalStepNr,
        sheetStepNr,
        safeText(col?.id || ''),
        getProcessLabel(sheet, colIdx),

        supplierText,
        systemText,
        inputText,
        processText,
        outputText,
        customerText,

        activityType,
        processStatus,
        leanValue,
        wj.value,
        workExpNote,

        ttfOverall == null ? '' : String(ttfOverall),
        iqfInput == null ? '' : String(iqfInput),
        iqfOutput == null ? '' : String(iqfOutput),

        splitVariant,
        variantRoute,
        parallel,

        yesNo(sysMerged),
        sysRange,

        yesNo(outMerged),
        outRange,
        yesNo(gateEnabled),
        failTarget.label,
        failTarget.col,
        passTarget.label,
        passTarget.col,

        sysFmt.names,
        sysFmt.scores,
        sysFmt.legacyFuture,
        sysFmt.answersJson,

        succesfactoren,
        oorzaken,
        maatregelen,
        verstoringen,
        inputSpecs,
        outputSpecs
      ];

      lines.push(row.map(toCsvField).join(';'));
    });
  });

  return lines.join('\n');
}

/* =========================================================
   EXPORT 2: IQF_Detail (1 rij per criterium per Input/Output)
   ========================================================= */

function buildIqfDetailRows(project) {
  const headers = [
    'Project',
    'Sheet',
    'Step Nr (Global)',
    'Step Nr (Sheet)',
    'Column Id',
    'Step Label',
    'RowType', // Input | Output
    'Criterion Key',
    'Criterion Label',
    'Weight',
    'Result',
    'Note',
    'Details(JSON)'
  ];

  const lines = [headers.map(toCsvField).join(';')];

  let globalStepNr = 0;

  (project?.sheets || []).forEach((sheet) => {
    let sheetStepNr = 0;

    (sheet.columns || []).forEach((col, colIdx) => {
      if (col?.isVisible === false) return;

      sheetStepNr += 1;
      globalStepNr += 1;

      const stepLabel = getProcessLabel(sheet, colIdx);
      const base = [
        safeText(project?.projectTitle || ''),
        safeText(sheet?.name || ''),
        globalStepNr,
        sheetStepNr,
        safeText(col?.id || ''),
        stepLabel
      ];

      const addRowType = (rowType, slotIdx) => {
        const slot = col?.slots?.[slotIdx] || {};
        IO_CRITERIA.forEach((c) => {
          const q = slot?.qa?.[c.key] || {};
          const result = q?.result ?? '';
          const note = q?.note ?? '';
          let details = '';
          try {
            details = JSON.stringify(q || {});
          } catch {
            details = '';
          }

          const row = [
            ...base,
            rowType,
            c.key,
            c.label,
            c.weight,
            result,
            note,
            details
          ];
          lines.push(row.map(toCsvField).join(';'));
        });
      };

      addRowType('Input', 2);
      addRowType('Output', 4);
    });
  });

  return lines.join('\n');
}

/* =========================================================
   EXPORT 3: System_Detail (1 rij per systeem binnen stap)
   ========================================================= */

function buildSystemDetailRows(project) {
  const headers = [
    'Project',
    'Sheet',
    'Step Nr (Global)',
    'Step Nr (Sheet)',
    'Column Id',
    'Step Label',
    'SystemIdx',
    'SystemName',
    'Legacy',
    'Future',
    'SystemScore(%)',
    'Answers(JSON)'
  ];

  const lines = [headers.map(toCsvField).join(';')];

  let globalStepNr = 0;

  (project?.sheets || []).forEach((sheet) => {
    let sheetStepNr = 0;

    (sheet.columns || []).forEach((col, colIdx) => {
      if (col?.isVisible === false) return;

      sheetStepNr += 1;
      globalStepNr += 1;

      const stepLabel = getProcessLabel(sheet, colIdx);

      const meta = getSystemsMetaForCol(col);
      const clean = sanitizeSystemsMeta(meta);
      if (!clean) return;

      clean.systems.forEach((s, idx) => {
        let answers = '';
        try {
          answers = JSON.stringify(s.qa || {});
        } catch {
          answers = '';
        }

        const row = [
          safeText(project?.projectTitle || ''),
          safeText(sheet?.name || ''),
          globalStepNr,
          sheetStepNr,
          safeText(col?.id || ''),
          stepLabel,
          idx + 1,
          safeText(s.name),
          yesNo(!!s.legacy),
          safeText(s.future),
          Number.isFinite(Number(s.score)) ? String(Number(s.score)) : '',
          answers
        ];

        lines.push(row.map(toCsvField).join(';'));
      });
    });
  });

  return lines.join('\n');
}

/* =========================================================
   Public export API
   ========================================================= */

export function exportToCSV() {
  try {
    const project = state.data;

    // 1) Overview (1 row per process step)
    const overview = buildAsIsOverviewRows(project);
    downloadBlob(new Blob(['\uFEFF' + overview], { type: 'text/csv;charset=utf-8;' }), getFileName('AsIs_Overview.csv'));

    // 2) IQF detail (per criterium)
    const iqfDetail = buildIqfDetailRows(project);
    downloadBlob(new Blob(['\uFEFF' + iqfDetail], { type: 'text/csv;charset=utf-8;' }), getFileName('IQF_Detail.csv'));

    // 3) System detail (per systeem)
    const sysDetail = buildSystemDetailRows(project);
    downloadBlob(new Blob(['\uFEFF' + sysDetail], { type: 'text/csv;charset=utf-8;' }), getFileName('System_Detail.csv'));

    Toast.show('CSV exports gemaakt: AsIs_Overview + IQF_Detail + System_Detail', 'success');
  } catch (e) {
    console.error(e);
    Toast.show('Fout bij genereren CSV', 'error');
  }
}

/* =========================================================
   HD Export (ongewijzigd)
   ========================================================= */

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