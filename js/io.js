import { state } from './state.js';
import { Toast } from './toast.js';
import { IO_CRITERIA } from './config.js';

const MERGE_LS_PREFIX = 'ssipoc.mergeGroups.v2';

function toCsvField(text) {
  /** Returns a semicolon-safe CSV field with quotes and escaped quotes. */
  if (text === null || text === undefined) return '""';
  const str = String(text);
  return `"${str.replace(/"/g, '""').replace(/\r?\n/g, ' ')}"`;
}

function getFileName(ext) {
  /** Returns a timestamped filename based on the project title. */
  const title = state.data?.projectTitle || 'sipoc_project';
  const safeTitle = String(title).replace(/[^a-z0-9]/gi, '_').toLowerCase();
  const now = new Date();
  const dateStr = now.toISOString().slice(0, 10);
  const timeStr = now.toTimeString().slice(0, 8).replace(/:/g, '');
  return `${safeTitle}_${dateStr}_${timeStr}.${ext}`;
}

function downloadBlob(blob, filename) {
  /** Triggers a browser download for a given blob and filename. */
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
  /** Triggers a PNG download for a given canvas. */
  const a = document.createElement('a');
  a.href = canvas.toDataURL('image/png');
  a.download = getFileName('png');
  a.click();
}

function calculateLSSScore(qa) {
  /** Computes the weighted IQF/LSS score (0-100) from stored QA results. */
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
  /** Maps a boolean-ish value to Ja/Nee. */
  return v ? 'Ja' : 'Nee';
}

function formatInputSpecs(defs) {
  /** Formats input definitions into a single human-readable string. */
  if (!Array.isArray(defs) || defs.length === 0) return '';
  return defs
    .filter((d) => d && (d.item || d.specifications || d.type))
    .map((d) => {
      const item = String(d.item || '').trim();
      const specs = String(d.specifications || '').trim();
      const type = String(d.type || '').trim();
      const main = [item, specs].filter(Boolean).join(': ');
      return type ? `${main} (${type})` : main;
    })
    .filter(Boolean)
    .join(' | ');
}

function formatDisruptions(disruptions) {
  /** Formats disruptions into a single human-readable string. */
  if (!Array.isArray(disruptions) || disruptions.length === 0) return '';
  return disruptions
    .filter((d) => d && (d.scenario || d.frequency || d.workaround))
    .map((d) => {
      const scenario = String(d.scenario || '').trim();
      const freq = String(d.frequency || '').trim();
      const workaround = String(d.workaround || '').trim();
      const left = [scenario, freq].filter(Boolean).join(' - ');
      return workaround ? `${left}: ${workaround}` : left;
    })
    .filter(Boolean)
    .join(' | ');
}

function formatWorkjoy(slot) {
  /** Returns a normalized workjoy object for export. */
  const v = slot?.workjoy ?? slot?.workJoy ?? null;
  if (!v) return { value: '', label: '', icon: '', context: '' };

  const V = String(v);
  if (V === 'OBSTACLE') return { value: V, label: 'Obstakel', icon: '🛠️', context: 'Kost energie & frustreert.' };
  if (V === 'ROUTINE') return { value: V, label: 'Routine', icon: '🤖', context: 'Saai & repeterend.' };
  if (V === 'FLOW') return { value: V, label: 'Flow', icon: '🚀', context: 'Geeft energie & voldoening.' };
  return { value: V, label: '', icon: '', context: '' };
}

function mergeKeyForSheet(project, sheet) {
  /** Returns the localStorage key for merge groups scoped to project and sheet. */
  const pid = project?.id || project?.name || project?.projectTitle || 'project';
  const sid = sheet?.id || sheet?.name || 'sheet';
  return `${MERGE_LS_PREFIX}:${pid}:${sid}`;
}

function loadMergeGroupsRaw(project, sheet) {
  /** Loads raw merge group definitions from localStorage for the given sheet key. */
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
  /** Returns true when the provided indices form a contiguous range. */
  if (!Array.isArray(cols) || cols.length < 2) return false;
  const s = [...new Set(cols)].sort((a, b) => a - b);
  return s.length === s[s.length - 1] - s[0] + 1;
}

function sanitizeMergeGroup(g, colCount) {
  /** Normalizes one merge group to a safe schema within current sheet bounds. */
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
          failTargetColIdx: Number.isFinite(Number(g.gate.failTargetColIdx)) ? Number(g.gate.failTargetColIdx) : null
        }
      : null;

  const systemsMeta = slotIdx === 1 && g?.systemsMeta && typeof g.systemsMeta === 'object' ? g.systemsMeta : null;

  return { slotIdx, cols: uniq.sort((a, b) => a - b), master, gate, systemsMeta };
}

function getMergeGroupsSanitized(project, sheet) {
  /** Returns sanitized merge groups for a sheet using localStorage data. */
  const n = sheet?.columns?.length ?? 0;
  if (!n) return [];
  return loadMergeGroupsRaw(project, sheet)
    .map((g) => sanitizeMergeGroup(g, n))
    .filter(Boolean);
}

function getMergeGroupForCell(groups, colIdx, slotIdx) {
  /** Returns the merge group containing a given cell or null. */
  return (groups || []).find((x) => x.slotIdx === slotIdx && Array.isArray(x.cols) && x.cols.includes(colIdx)) || null;
}

function getMergeRangeString(group) {
  /** Returns a human-readable 1-based merge range like "2-4" or an empty string. */
  if (!group?.cols?.length) return '';
  const min = Math.min(...group.cols) + 1;
  const max = Math.max(...group.cols) + 1;
  return min === max ? String(min) : `${min}-${max}`;
}

function getProcessLabel(sheet, colIdx) {
  /** Returns the process label for a column or a fallback label. */
  const t = sheet?.columns?.[colIdx]?.slots?.[3]?.text;
  const s = String(t ?? '').trim();
  return s || `Kolom ${Number(colIdx) + 1}`;
}

function getNextVisibleColIdx(sheet, fromIdx) {
  /** Returns the next visible column index after fromIdx or null. */
  const n = sheet?.columns?.length ?? 0;
  for (let i = fromIdx + 1; i < n; i++) {
    if (sheet.columns[i]?.isVisible !== false) return i;
  }
  return null;
}

function getPassTargetFromGroup(sheet, group) {
  /** Returns the pass target label and column number derived from the next visible column. */
  if (!group?.cols?.length) return { label: '', col: '' };
  const maxCol = Math.max(...group.cols);
  const nextIdx = getNextVisibleColIdx(sheet, maxCol);
  if (nextIdx == null) return { label: 'Einde proces', col: '' };
  return { label: getProcessLabel(sheet, nextIdx), col: String(nextIdx + 1) };
}

function getFailTargetFromGate(sheet, gate) {
  /** Returns the fail target label and column number from a gate config. */
  if (!gate?.enabled) return { label: '', col: '' };
  const idx = gate.failTargetColIdx;
  if (idx == null || !Number.isFinite(Number(idx))) return { label: '', col: '' };
  return { label: getProcessLabel(sheet, idx), col: String(Number(idx) + 1) };
}

function toLetter(i0) {
  /** Maps a zero-based index to an alphabetic letter for variant routes. */
  const n = Number(i0);
  if (!Number.isFinite(n) || n < 0) return 'A';
  const base = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  return base[n] || `R${n + 1}`;
}

function computeVariantLetterMap(sheet) {
  /** Computes the per-column variant route letter map for contiguous variant runs. */
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

function sanitizeSystemsMeta(meta) {
  /** Normalizes systems meta for export while preserving stored answers. */
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

function formatSystemsList(meta) {
  /** Formats systems meta into export-friendly strings including an answers JSON blob. */
  const clean = sanitizeSystemsMeta(meta);
  if (!clean) return { names: '', scores: '', legacyFuture: '', answersJson: '' };

  const names = clean.systems
    .map((s) => {
      const nm = String(s.name || '').trim();
      if (!nm) return '';
      return s.legacy ? `${nm} (Legacy)` : nm;
    })
    .filter(Boolean)
    .join(' | ');

  const scores = clean.systems.map((s) => (Number.isFinite(Number(s.score)) ? `${Number(s.score)}%` : '—')).join('; ');

  const legacyFuture = clean.systems
    .filter((s) => s.legacy && (String(s.name || '').trim() || String(s.future || '').trim()))
    .map((s) => `${String(s.name || '').trim() || '—'} -> ${String(s.future || '').trim() || '—'}`)
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

function buildOutputMaps(sheet, mergeGroups) {
  /** Builds stable per-sheet OUT ids while skipping merged output slaves. */
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

export async function saveToFile() {
  /** Saves the project JSON to disk via File System Access API or legacy download. */
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
  /** Loads a project JSON file into state and triggers UI updates. */
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

export function exportToCSV() {
  /** Exports the current project to a detailed CSV including merge, gate, and QA details. */
  try {
    const headers = [
      'Sheet',
      'Kolom Nr',
      'SIPOC Categorie',
      'Input ID',
      'Output ID',
      'Inhoud (Sticky)',
      'Type',
      'Proces Status',
      'LSS Waarde',
      'IQF Score (%)',
      'TTF Overall (%)',
      'TTF Scores (per systeem)',
      'Systemen (met Legacy)',
      'Legacy -> Toekomst',
      'TTF Antwoorden (JSON)',

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

      'Succesfactoren',
      'Oorzaken (RC)',
      'Maatregelen (CM)',
      'Verstoringen',
      'Input/Output Specificaties',

      'Werkplezier (waarde)',
      'Werkplezier (label)',
      'Werkplezier (icoon)',
      'Werkplezier (context)',

      ...IO_CRITERIA.flatMap((c) => [
        `IQF:${c.label} - Result`,
        `IQF:${c.label} - Note`,
        `IQF:${c.label} - Details(JSON)`
      ])
    ];

    const lines = [headers.map(toCsvField).join(';')];

    const project = state.data;
    let globalVisibleColNr = 0;

    (project?.sheets || []).forEach((sheet) => {
      const mergeGroups = getMergeGroupsSanitized(project, sheet);
      const variantMap = computeVariantLetterMap(sheet);
      const { outIdBySlotId, outTextByOutId } = buildOutputMaps(sheet, mergeGroups);

      let sheetVisibleColNr = 0;
      let globalIn = 0;

      (sheet.columns || []).forEach((col, colIdx) => {
        if (col?.isVisible === false) return;
        sheetVisibleColNr += 1;
        globalVisibleColNr += 1;

        const slotInput = col.slots?.[2];
        const slotOutput = col.slots?.[4];

        const sysGroup = getMergeGroupForCell(mergeGroups, colIdx, 1);
        const outGroup = getMergeGroupForCell(mergeGroups, colIdx, 4);

        const sysMerged = !!sysGroup;
        const sysRange = sysGroup ? getMergeRangeString(sysGroup) : '';

        const outMerged = !!outGroup;
        const outRange = outGroup ? getMergeRangeString(outGroup) : '';

        const gateEnabled = outGroup?.gate?.enabled ? true : false;
        const failTarget = outGroup ? getFailTargetFromGate(sheet, outGroup.gate) : { label: '', col: '' };
        const passTarget = outGroup ? getPassTargetFromGroup(sheet, outGroup) : { label: '', col: '' };

        const outputId = (() => {
          if (outGroup && colIdx !== outGroup.master) return '';
          return slotOutput?.id ? outIdBySlotId[slotOutput.id] || '' : '';
        })();

        const inputId = (() => {
          const hasLinked = !!slotInput?.linkedSourceId;
          const hasInputText = !!slotInput?.text?.trim();
          if (hasLinked) return slotInput.linkedSourceId;
          if (hasInputText) {
            globalIn += 1;
            return `IN${globalIn}`;
          }
          return '';
        })();

        (col.slots || []).forEach((slot, slotIdx) => {
          const category = ['Leverancier', 'Systeem', 'Input', 'Proces', 'Output', 'Klant'][slotIdx] || '';

          const isProcess = slotIdx === 3;
          const isSystem = slotIdx === 1;
          const isIoRow = slotIdx === 2 || slotIdx === 4;

          let inhoud = slot?.text ?? '';
          if (slotIdx === 2 && slotInput?.linkedSourceId) {
            const linkedId = slotInput.linkedSourceId;
            if (outTextByOutId[linkedId]) inhoud = outTextByOutId[linkedId];
          }

          const type = isProcess ? (slot?.type || '') : '';
          const procStatus = isProcess ? (slot?.processStatus || '') : '';
          const lssWaarde = isProcess ? (slot?.processValue || '') : '';

          const iqfScore = isIoRow ? calculateLSSScore(slot?.qa) : null;

          const sysOverall = isSystem ? (slot?.systemData?.calculatedScore ?? null) : null;
          const sysMeta = isSystem ? slot?.systemData?.systemsMeta : null;
          const sysFmt = isSystem
            ? formatSystemsList(sysMeta)
            : { names: '', scores: '', legacyFuture: '', answersJson: '' };

          const splitVariant = yesNo(!!col.isVariant);
          const variantRoute = col.isVariant ? String(variantMap[colIdx] || '') : '';
          const parallel = yesNo(!!col.isParallel);

          const sysMergedCell = slotIdx === 1 ? yesNo(sysMerged) : '';
          const sysRangeCell = slotIdx === 1 ? sysRange : '';

          const outMergedCell = slotIdx === 4 ? yesNo(outMerged) : '';
          const outRangeCell = slotIdx === 4 ? outRange : '';
          const gateEnabledCell = slotIdx === 4 ? yesNo(gateEnabled) : '';
          const gateFailLabelCell = slotIdx === 4 ? failTarget.label : '';
          const gateFailColCell = slotIdx === 4 ? failTarget.col : '';
          const gatePassLabelCell = slotIdx === 4 ? passTarget.label : '';
          const gatePassColCell = slotIdx === 4 ? passTarget.col : '';

          const succesfactoren = isProcess ? (slot?.successFactors || '') : '';
          const oorzaken = isProcess && Array.isArray(slot?.causes) ? slot.causes.join(' | ') : '';
          const maatregelen = isProcess && Array.isArray(slot?.improvements) ? slot.improvements.join(' | ') : '';
          const verstoringen = isProcess ? formatDisruptions(slot?.disruptions) : '';

          const specs = isIoRow ? formatInputSpecs(slot?.inputDefinitions) : '';

          const wj = formatWorkjoy(slot);

          const qaDetailCols = IO_CRITERIA.flatMap((c) => {
            const q = slot?.qa?.[c.key] || {};
            const result = q?.result ?? '';
            const note = q?.note ?? '';
            let details = '';
            try {
              details = JSON.stringify(q || {});
            } catch {
              details = '';
            }
            return [result, note, details];
          });

          const row = [
            sheet.name || '',
            globalVisibleColNr,
            category,
            slotIdx === 2 ? inputId : '',
            slotIdx === 4 ? outputId : '',
            inhoud,
            type,
            procStatus,
            lssWaarde,
            iqfScore === null ? '' : String(iqfScore),
            isSystem && sysOverall != null ? String(sysOverall) : '',
            isSystem ? sysFmt.scores : '',
            isSystem ? sysFmt.names : '',
            isSystem ? sysFmt.legacyFuture : '',
            isSystem ? sysFmt.answersJson : '',

            splitVariant,
            variantRoute,
            parallel,

            sysMergedCell,
            sysRangeCell,

            outMergedCell,
            outRangeCell,
            gateEnabledCell,
            gateFailLabelCell,
            gateFailColCell,
            gatePassLabelCell,
            gatePassColCell,

            succesfactoren,
            oorzaken,
            maatregelen,
            verstoringen,
            specs,

            wj.value,
            wj.label,
            wj.icon,
            wj.context,

            ...qaDetailCols
          ];

          lines.push(row.map(toCsvField).join(';'));
        });
      });
    });

    const blob = new Blob(['\uFEFF' + lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
    downloadBlob(blob, getFileName('csv'));
  } catch (e) {
    console.error(e);
    Toast.show('Fout bij genereren CSV', 'error');
  }
}

export async function exportHD(copyToClipboard = false) {
  /** Exports the current board as a high-resolution image optionally to clipboard. */
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