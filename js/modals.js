// modals.js
import { state } from './state.js';
import {
  IO_CRITERIA,
  SYSTEM_QUESTIONS,
  ACTIVITY_TYPES,
  LEAN_VALUES,
  PROCESS_STATUSES,
  DISRUPTION_FREQUENCIES,
  DEFINITION_TYPES
} from './config.js';

let editingSticky = null;
let areListenersAttached = false;

const $ = (id) => document.getElementById(id);

const deepClone = (obj) => {
  try {
    if (typeof structuredClone === 'function') return structuredClone(obj);
  } catch (_) {}
  return JSON.parse(JSON.stringify(obj));
};

// --- ROUTE LETTERS HELPERS (Gebruikt in Variant modal) ---
function _toLetter(i) {
  const n = Number(i);
  if (!Number.isFinite(n) || n < 0) return 'A';
  const base = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  return base[n] || `R${n + 1}`;
}
// ---------------------------------------

const PROCESS_STATUS_DEFS = {
  HAPPY: {
    title: 'Onder controle',
    body:
      'Het proces verloopt voorspelbaar en stabiel. Input/werkstappen zijn duidelijk, afwijkingen zijn zeldzaam en impact is laag. Er is geen herstelwerk nodig om door te kunnen.'
  },
  NEUTRAL: {
    title: 'Aandachtspunt',
    body:
      'Het proces werkt meestal, maar is niet altijd voorspelbaar. Er zijn terugkerende haperingen of variatie waardoor soms extra afstemming/herstelwerk nodig is. Risico op verstoring is aanwezig.'
  },
  SAD: {
    title: 'Niet onder controle',
    body:
      'Het proces is instabiel of faalt regelmatig. Variatie en verstoringen zijn hoog, er is vaak herstelwerk nodig, en doorlooptijd/kwaliteit wordt structureel geraakt.'
  }
};

const LEAN_VALUE_DEFS = {
  VA: {
    title: '💚 VA - Patiëntwaarde (Value Added)',
    body: 'Activiteiten die direct bijdragen aan de genezing, zorg of het welzijn van de patiënt. Dit is waar de patiënt voor komt.'
  },
  BNVA: {
    title: '⚖️ BNVA - Business Noodzaak (Business Non-Value Added)',
    body: 'Voegt geen directe waarde toe voor de patiënt, maar is noodzakelijk vanuit wet- en regelgeving, veiligheid of administratie (Muda Type 1).'
  },
  NVA: {
    title: '🗑️ NVA - Verspilling (Non-Value Added)',
    body: 'Voegt geen waarde toe voor de patiënt en is niet noodzakelijk. Kost alleen tijd en middelen (bv. wachten, zoeken, herstelwerk). Moet geëlimineerd worden (Muda Type 2).'
  }
};

const ACTIVITY_TYPE_DEFS = {
  Taak: {
    title: '📝 Taak (Zonder Patiënt)',
    body: 'Activiteiten die worden uitgevoerd zonder dat de patiënt erbij aanwezig is (bv. voorbereidingen, administratie, labwerk, uitwerken).'
  },
  Afspraak: {
    title: '📅 Afspraak (Met Patiënt)',
    body: 'Activiteit waarbij de patiënt fysiek of digitaal aanwezig is (bv. consult, behandeling, bloedprikken, intake).'
  },
  TimeOut: {
    title: '🛑 Time Out (Pauze/Check)',
    body: 'Een korte pauze in het proces om te controleren of alles compleet is en er niets vergeten is voordat men verder gaat.'
  },
  Beoordeling: {
    title: '🔎 Beoordeling (Collega Check)',
    body: 'Een moment waarop het werk wordt beoordeeld of gecontroleerd door een andere collega (bv. arts, supervisor of 2e beoordelaar).'
  }
};

const WORK_EXP_OPTIONS = [
  {
    value: 'OBSTACLE',
    icon: '🛠️',
    label: 'Obstakel',
    title: 'Obstakel',
    body: 'Kost energie & frustreert. Het proces werkt tegen me. (Actie: Verbeteren)',
    cls: 'selected-sad'
  },
  {
    value: 'ROUTINE',
    icon: '🤖',
    label: 'Routine',
    title: 'Routine',
    body: 'Saai & repeterend. Ik voeg hier geen unieke waarde toe. (Actie: Automatiseren)',
    cls: 'selected-neu'
  },
  {
    value: 'FLOW',
    icon: '🚀',
    label: 'Flow',
    title: 'Flow',
    body: 'Geeft energie & voldoening. Hier maak ik het verschil. (Actie: Koesteren)',
    cls: 'selected-hap'
  }
];

const IO_CRITERIA_DEFS = {
  compleet: {
    gating: 'Alle benodigde data/materialen zijn aanwezig om mijn taak uit te voeren.',
    impactA: 'A. Blokkerend: Ik kan niet starten of niet afronden zonder deze input.',
    impactB:
      'B. Extra werk: Ik kan doorgaan, maar moet actief ontbrekende input ophalen (zoeken/bellen/mailen/aanvullen).',
    impactC: 'C. Kleine frictie: Ik kan doorgaan met minimale handelingen; geen relevant risico.'
  },
  kwaliteit: {
    gating: 'De ontvangen data is correct en bruikbaar (formaat/resolutie/inhoud).',
    impactA: 'A. Blokkerend: Onbruikbaar: taak kan niet veilig/correct uitgevoerd worden.',
    impactB: 'B. Extra werk: Bruikbaar na correctie/omweg (converteren, herexport, opnieuw opvragen).',
    impactC: 'C. Kleine schoonheidsfout; nauwelijks effect op uitvoering.'
  },
  duidelijkheid: {
    gating: 'De input is eenduidig; ik hoef niets te interpreteren of na te vragen om te starten.',
    impactA:
      'A. Blokkerend: Onvoldoende eenduidig: ik kan niet verantwoord starten zonder verduidelijking.',
    impactB: 'B. Extra werk: Ik kan starten, maar moet 1+ verduidelijkingsacties doen (vragen/afstemmen).',
    impactC: 'C. Kleine frictie: Kleine onduidelijkheid; geen effect op uitkomst.'
  },
  tijdigheid: {
    gating: 'De input is beschikbaar op het moment dat ik deze nodig heb.',
    impactA: 'A. Blokkerend: Niet op tijd: planning/uitvoering ligt stil.',
    impactB: 'B. Extra werk: Niet op tijd: ik kan deels door, maar moet omplannen/extra afstemmen.',
    impactC: 'C. Kleine frictie: Licht te laat; geen echte consequenties.'
  },
  standaard: {
    gating: 'De input volgt de afgesproken naamgeving/templates/protocollen.',
    impactA: 'A. Blokkerend: Niet conform: ik kan niet veilig/correct verder (risico/foutkans te groot).',
    impactB: 'B. Extra werk: Niet conform: ik kan verder na herlabelen/structureren/opschonen.',
    impactC: 'C. Kleine frictie: Afwijking is cosmetisch; nauwelijks effect.'
  },
  overdracht: {
    gating: 'Status/registratie is correct bijgewerkt in de bronsystemen.',
    impactA: 'A. Blokkerend: Onjuiste status: ik kan niet verantwoord handelen (risico/foutkans te groot).',
    impactB: 'B. Extra werk: Status klopt niet: ik kan handelen, maar moet actief verifiëren/corrigeren.',
    impactC: 'C. Kleine frictie: Kleine administratieve afwijking; geen effect op uitvoering.'
  }
};

function escapeAttr(v) {
  return String(v ?? '')
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

const createRadioGroup = (name, options, selectedValue, isHorizontal = false) => `
  <div class="radio-group-container ${isHorizontal ? 'horizontal' : 'vertical'}">
    ${options
      .map((opt) => {
        const val = opt.value ?? opt;
        const label = opt.label ?? opt;
        const ttTitle = opt.ttTitle || '';
        const ttBody = opt.ttBody || '';
        const isSelected =
          selectedValue !== null &&
          selectedValue !== undefined &&
          String(val) === String(selectedValue);

        return `
          <div class="sys-opt ${isSelected ? 'selected' : ''}" 
               data-value="${escapeAttr(val)}"
               ${ttBody ? `data-tt-title="${escapeAttr(ttTitle)}" data-tt-body="${escapeAttr(ttBody)}"` : ''}>
            ${escapeAttr(label)}
          </div>
        `;
      })
      .join('')}
    <input type="hidden" name="${escapeAttr(name)}" value="${
      selectedValue !== null && selectedValue !== undefined ? escapeAttr(selectedValue) : ''
    }">
  </div>
`;

const createCombinedAnalysisRows = (causes, improvements) => {
  const c = causes || [];
  const i = improvements || [];
  const count = Math.max(c.length, i.length, 1);
  let html = '';

  for (let idx = 0; idx < count; idx++) {
    const causeVal = c[idx] || '';
    const imprVal = i[idx] || '';
    if (count > 1 && !causeVal && !imprVal) continue;

    html += `
      <tr>
        <td><input type="text" class="def-input input-cause" value="${escapeAttr(
          causeVal
        )}" placeholder="Oorzaak..."></td>
        <td><input type="text" class="def-input input-measure" value="${escapeAttr(
          imprVal
        )}" placeholder="Maatregel..."></td>
        <td><button class="btn-row-del-tiny" data-action="remove-row" type="button">×</button></td>
      </tr>
    `;
  }

  if (!html) {
    html = `
      <tr>
        <td><input type="text" class="def-input input-cause" placeholder="Oorzaak..."></td>
        <td><input type="text" class="def-input input-measure" placeholder="Maatregel..."></td>
        <td><button class="btn-row-del-tiny" data-action="remove-row" type="button">×</button></td>
      </tr>
    `;
  }
  return html;
};

const getStickyData = () => {
  if (!editingSticky) return null;
  const sheet = state.activeSheet;
  return sheet.columns[editingSticky.colIdx].slots[editingSticky.slotIdx];
};

const getVisibleColIdxs = () => {
  const sheet = state.activeSheet;
  if (!sheet) return [];
  const vis = [];
  sheet.columns.forEach((c, idx) => {
    if (c?.isVisible !== false) vis.push(idx);
  });
  return vis;
};

/* =========================================================
   Input-linking (MULTI) helpers (UUIDs nu!)
   ========================================================= */

function normalizeLinkedSourceIds(data) {
  if (!data || typeof data !== 'object') return [];

  if (Array.isArray(data.linkedSourceUids) && data.linkedSourceUids.length > 0) {
    return data.linkedSourceUids;
  }

  const arr = Array.isArray(data.linkedSourceIds)
    ? data.linkedSourceIds
    : data.linkedSourceId
      ? [data.linkedSourceId]
      : [];

  return [...new Set(arr.map((x) => String(x ?? '').trim()).filter(Boolean))];
}

function setLinkedSourceIds(data, uids) {
  if (!data || typeof data !== 'object') return;
  const clean = (Array.isArray(uids) ? uids : [])
    .map((x) => String(x ?? '').trim())
    .filter((x) => x !== '');
  const uniq = [...new Set(clean)];

  data.linkedSourceUids = uniq.length ? uniq : [];
  data.linkedSourceUid = uniq.length ? uniq[0] : null;

  data.linkedSourceIds = [];
  data.linkedSourceId = null;
}

function getOutputUidMap() {
  const details = state.getAllOutputsDetailed();
  const map = {};
  details.forEach((d) => {
    map[d.uid] = d;
  });
  return map;
}

function _formatLinkedSources(ids, uidMap) {
  const map = uidMap || getOutputUidMap();
  const arr = Array.isArray(ids) ? ids : [];

  const parts = arr
    .map((uid) => {
      const key = String(uid ?? '').trim();
      if (!key) return '';

      const entry = map[key];
      if (!entry) return '???';

      const raw = String(entry.text ?? '').trim();
      const label = entry.outId;

      if (!raw) return label;
      const short = raw.length > 42 ? `${raw.slice(0, 42)}...` : raw;
      return `${short} (${label})`;
    })
    .filter((x) => x !== '');

  return parts.join('; ');
}

function updateLinkedInfoUI(ids) {
  const info = $('linkedInfoText');
  const summary = $('linkedSourcesSummary');
  const n = Array.isArray(ids) ? ids.length : 0;

  const uidMap = getOutputUidMap();
  const formatted = n ? _formatLinkedSources(ids, uidMap) : '';

  if (info) {
    info.style.display = n ? 'block' : 'none';
    info.textContent = n ? `🔗 Gekoppeld (${n}): ${formatted}` : '';
  }
  if (summary) {
    summary.textContent = n ? formatted : '—';
  }
}

/* =========================================================
   Output-bundles (Pakketten) helpers
   ========================================================= */

function ensureOutputBundlesArray() {
  const project = state.data;
  if (!project || typeof project !== 'object') return [];
  if (!Array.isArray(project.outputBundles)) project.outputBundles = [];
  return project.outputBundles;
}

function makeLocalId(prefix = 'id') {
  return `${prefix}_${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

function getActiveBundleId(data) {
  if (!data || typeof data !== 'object') return '';
  const a = Array.isArray(data.linkedBundleIds) ? data.linkedBundleIds : [];
  const first = a.length ? String(a[0] ?? '').trim() : '';
  const single = String(data.linkedBundleId ?? '').trim();
  return single || first;
}

function setActiveBundleId(data, bundleId) {
  if (!data || typeof data !== 'object') return;
  const id = String(bundleId ?? '').trim();
  data.linkedBundleId = id || null;
  data.linkedBundleIds = id ? [id] : null;
}

function findBundleById(bundleId) {
  const id = String(bundleId ?? '').trim();
  if (!id) return null;
  const arr = ensureOutputBundlesArray();
  return arr.find((b) => String(b?.id ?? '').trim() === id) || null;
}

function formatBundleOutputs(bundle, uidMap) {
  const map = uidMap || getOutputUidMap();
  const ids = Array.isArray(bundle?.outputUids) ? bundle.outputUids : [];

  const parts = ids
    .map((uid) => {
      const key = String(uid ?? '').trim();
      if (!key) return '';

      const entry = map[key];
      if (!entry) return '???';

      const raw = String(entry.text ?? '').trim();
      const label = entry.outId;

      if (!raw) return label;
      const short = raw.length > 42 ? `${raw.slice(0, 42)}...` : raw;
      return `${short} (${label})`;
    })
    .filter((x) => x !== '');

  return parts.join('; ');
}

function syncOutputCheckboxesFromLinkedSources(data) {
  const content = $('modalContent');
  if (!content) return;

  const ids = new Set(normalizeLinkedSourceIds(data));
  content.querySelectorAll('.input-source-cb').forEach((cb) => {
    cb.disabled = false;
    cb.checked = ids.has(String(cb.value || '').trim());
  });
}

function syncOutputCheckboxesFromBundle(bundle) {
  const content = $('modalContent');
  if (!content) return;

  const ids = new Set(
    (Array.isArray(bundle?.outputUids) ? bundle.outputUids : []).map((x) => String(x ?? '').trim())
  );

  content.querySelectorAll('.input-source-cb').forEach((cb) => {
    const key = String(cb.value || '').trim();
    cb.checked = ids.has(key);
    cb.disabled = true;
  });
}

function updateBundleInfoUI(data) {
  const content = $('modalContent');
  if (!content) return;

  const bundleInfo = content.querySelector('#bundleActiveInfo');
  const bundleHint = content.querySelector('#bundleHint');
  const bundlePick = content.querySelector('#bundlePick');
  const bundleName = content.querySelector('#bundleName');
  const bundleDeleteBtn = content.querySelector('#bundleDeleteDefBtn');

  const uidMap = getOutputUidMap();
  const activeId = getActiveBundleId(data);
  const bundle = activeId ? findBundleById(activeId) : null;

  if (bundlePick) bundlePick.value = bundle ? String(bundle.id) : '';
  if (bundleName) bundleName.value = bundle ? String(bundle.name || '').trim() : '';

  if (bundleInfo) {
    if (!bundle) {
      bundleInfo.style.display = 'none';
      bundleInfo.textContent = '';
    } else {
      const nm = String(bundle.name || '').trim() || 'Pakket';
      const detail = formatBundleOutputs(bundle, uidMap);
      bundleInfo.style.display = 'block';
      bundleInfo.textContent = detail ? `📦 ${nm}: ${detail}` : `📦 ${nm}`;
    }
  }

  if (bundleHint) {
    if (!bundle) bundleHint.textContent = '';
    else bundleHint.textContent = `Bevat: ${(Array.isArray(bundle.outputUids) ? bundle.outputUids.length : 0)} outputs`;
  }

  if (bundleDeleteBtn) {
    bundleDeleteBtn.style.display = bundle ? 'inline-block' : 'none';
  }

  if (bundle) {
    data.text = '';
    setLinkedSourceIds(data, []);

    const ext = $('inputExternalToggle');
    if (ext) ext.checked = false;

    syncOutputCheckboxesFromBundle(bundle);

    const summary = $('linkedSourcesSummary');
    if (summary) summary.textContent = `📦 ${String(bundle.name || 'Pakket').trim()}`;

    const info = $('linkedInfoText');
    if (info) {
      info.style.display = 'none';
      info.textContent = '';
    }

    const list = content.querySelector('#inputSourcesList');
    if (list) {
      list.style.opacity = '0.55';
      list.style.pointerEvents = 'none';
      list.style.filter = 'grayscale(0.2)';
    }
  } else {
    syncOutputCheckboxesFromLinkedSources(data);

    const ids = normalizeLinkedSourceIds(data);
    const list = content.querySelector('#inputSourcesList');
    if (list) {
      const hasAny = ids.length > 0;
      list.style.opacity = hasAny ? '1' : '0.55';
      list.style.pointerEvents = hasAny ? 'auto' : 'none';
      list.style.filter = hasAny ? 'none' : 'grayscale(0.2)';
    }
  }
}

/* =========================================================
   Merge helpers in modal
   ========================================================= */

const getCurrentOutputMergeRange = () => {
  const sheet = state.activeSheet;
  if (!sheet || !editingSticky) return null;
  if (editingSticky.slotIdx !== 4) return null;
  return state.getOutputMergeForCol(editingSticky.colIdx);
};

const getCurrentSystemMergeRange = () => {
  const sheet = state.activeSheet;
  if (!sheet || !editingSticky) return null;
  if (editingSticky.slotIdx !== 1) return null;
  return state.getSystemMergeForCol(editingSticky.colIdx);
};

const renderOutputMergeControls = () => {
  if (!editingSticky || editingSticky.slotIdx !== 4) return '';

  const sheet = state.activeSheet;
  if (!sheet) return '';

  const colIdx = editingSticky.colIdx;
  const vis = getVisibleColIdxs();
  const pos = vis.indexOf(colIdx);
  if (pos === -1) return '';

  const currentRange = getCurrentOutputMergeRange();
  const hasMerge = !!currentRange;

  const startDefault = hasMerge ? currentRange.startCol : colIdx;
  const endDefault = hasMerge ? currentRange.endCol : colIdx;

  const startOpts = vis
    .slice(0, pos + 1)
    .map((idx) => {
      const lbl = `Kolom ${idx + 1}`;
      return `<option value="${idx}" ${idx === startDefault ? 'selected' : ''}>${escapeAttr(lbl)}</option>`;
    })
    .join('');

  const endOpts = vis
    .slice(pos)
    .map((idx) => {
      const lbl = `Kolom ${idx + 1}`;
      return `<option value="${idx}" ${idx === endDefault ? 'selected' : ''}>${escapeAttr(lbl)}</option>`;
    })
    .join('');

  return `
    <div style="margin-bottom:16px; padding:12px; background: rgba(0,0,0,0.22); border-radius: 10px; border: 1px solid rgba(255,255,255,0.10);">
      <div class="modal-label" style="margin-top:0;">Output samenvoegen (aaneengesloten kolommen)</div>
      <div class="io-helper" style="margin-top:6px; margin-bottom:10px;">
        Kies zelf welke <strong>adjacente</strong> outputs je wil mergen. De range moet aaneengesloten zijn en jouw huidige kolom bevatten.
      </div>

      <label style="display:flex; align-items:center; gap:10px; font-size:13px; margin-bottom:10px; cursor:pointer;">
        <input id="mergeEnable" type="checkbox" ${hasMerge ? 'checked' : ''} />
        Merge inschakelen
      </label>

      <div style="display:flex; gap:10px; flex-wrap:wrap;">
        <div style="flex:1; min-width:180px;">
          <div style="font-size:12px; opacity:.85; margin-bottom:6px;">Van (linker grens)</div>
          <select id="mergeStartSelect" class="modal-input" ${hasMerge ? '' : 'disabled'}>${startOpts}</select>
        </div>

        <div style="flex:1; min-width:180px;">
          <div style="font-size:12px; opacity:.85; margin-bottom:6px;">Tot (rechter grens)</div>
          <select id="mergeEndSelect" class="modal-input" ${hasMerge ? '' : 'disabled'}>${endOpts}</select>
        </div>
      </div>

      <div id="mergeHint" style="margin-top:10px; font-size:12px; opacity:.85;">
        ${hasMerge ? `Actief: kolom ${startDefault + 1} t/m ${endDefault + 1}` : 'Niet gemerged'}
      </div>
    </div>
  `;
};

const renderSystemMergeControls = () => {
  if (!editingSticky || editingSticky.slotIdx !== 1) return '';

  const sheet = state.activeSheet;
  if (!sheet) return '';

  const colIdx = editingSticky.colIdx;
  const vis = getVisibleColIdxs();
  const pos = vis.indexOf(colIdx);
  if (pos === -1) return '';

  const currentRange = getCurrentSystemMergeRange();
  const hasMerge = !!currentRange;

  const startDefault = hasMerge ? currentRange.startCol : colIdx;
  const endDefault = hasMerge ? currentRange.endCol : colIdx;

  const startOpts = vis
    .slice(0, pos + 1)
    .map((idx) => {
      const lbl = `Kolom ${idx + 1}`;
      return `<option value="${idx}" ${idx === startDefault ? 'selected' : ''}>${escapeAttr(lbl)}</option>`;
    })
    .join('');

  const endOpts = vis
    .slice(pos)
    .map((idx) => {
      const lbl = `Kolom ${idx + 1}`;
      return `<option value="${idx}" ${idx === endDefault ? 'selected' : ''}>${escapeAttr(lbl)}</option>`;
    })
    .join('');

  return `
    <div style="margin-bottom:16px; padding:12px; background: rgba(0,0,0,0.22); border-radius: 10px; border: 1px solid rgba(255,255,255,0.10);">
      <div class="modal-label" style="margin-top:0;">Systeem samenvoegen (aaneengesloten kolommen)</div>
      <div class="io-helper" style="margin-top:6px; margin-bottom:10px;">
        Handig als meerdere processtappen in hetzelfde systeem werken. Je krijgt dan één strak systeem-post-it (zoals bij Output merge).
      </div>

      <label style="display:flex; align-items:center; gap:10px; font-size:13px; margin-bottom:10px; cursor:pointer;">
        <input id="sysMergeEnable" type="checkbox" ${hasMerge ? 'checked' : ''} />
        Merge inschakelen
      </label>

      <div style="display:flex; gap:10px; flex-wrap:wrap;">
        <div style="flex:1; min-width:180px;">
          <div style="font-size:12px; opacity:.85; margin-bottom:6px;">Van (linker grens)</div>
          <select id="sysMergeStartSelect" class="modal-input" ${hasMerge ? '' : 'disabled'}>${startOpts}</select>
        </div>

        <div style="flex:1; min-width:180px;">
          <div style="font-size:12px; opacity:.85; margin-bottom:6px;">Tot (rechter grens)</div>
          <select id="sysMergeEndSelect" class="modal-input" ${hasMerge ? '' : 'disabled'}>${endOpts}</select>
        </div>
      </div>

      <div id="sysMergeHint" style="margin-top:10px; font-size:12px; opacity:.85;">
        ${hasMerge ? `Actief: kolom ${startDefault + 1} t/m ${endDefault + 1}` : 'Niet gemerged'}
      </div>
    </div>
  `;
};

/* =========================================================
   SYSTEM TAB (multi-system + per-system score)
   ========================================================= */

function ensureSystemDataShape(data) {
  data.systemData = data.systemData || {};
  const sd = data.systemData;

  if (typeof sd.isMulti !== 'boolean') sd.isMulti = false;
  if (typeof sd.noSystem !== 'boolean') sd.noSystem = false;

  if (!Array.isArray(sd.systems)) {
    const legacyName = typeof sd.systemName === 'string' ? sd.systemName.trim() : '';
    sd.systems = legacyName
      ? [
          {
            id: `${Date.now()}_${Math.random().toString(16).slice(2)}`,
            name: legacyName,
            isLegacy: false,
            futureSystem: '',
            qa: {},
            calculatedScore: null
          }
        ]
      : [
          {
            id: `${Date.now()}_${Math.random().toString(16).slice(2)}`,
            name: '',
            isLegacy: false,
            futureSystem: '',
            qa: {},
            calculatedScore: null
          }
        ];
  }

  sd.systems = sd.systems.map((s) => ({
    id: String(s?.id || `${Date.now()}_${Math.random().toString(16).slice(2)}`),
    name: String(s?.name || ''),
    isLegacy: !!s?.isLegacy,
    futureSystem: String(s?.futureSystem || ''),
    qa: s?.qa && typeof s.qa === 'object' ? { ...s.qa } : {},
    calculatedScore: Number.isFinite(Number(s?.calculatedScore)) ? Number(s.calculatedScore) : null
  }));

  if (sd.systems.length === 0) {
    sd.systems = [
      {
        id: `${Date.now()}_${Math.random().toString(16).slice(2)}`,
        name: '',
        isLegacy: false,
        futureSystem: '',
        qa: {},
        calculatedScore: null
      }
    ];
  }

  if (typeof sd.systemName !== 'string') sd.systemName = sd.systems[0]?.name || '';
  if (!Number.isFinite(Number(sd.calculatedScore))) sd.calculatedScore = null;
}

function computeSystemScoreFromAnswers(answerObj) {
  let total = 0;
  let answeredCount = 0;

  SYSTEM_QUESTIONS.forEach((q) => {
    const v = answerObj?.[q.id];
    if (!Number.isFinite(Number(v))) return;
    total += Number(v);
    answeredCount++;
  });

  if (answeredCount === 0) return null;

  const maxPoints = SYSTEM_QUESTIONS.length * 3;
  const safeTotal = Math.min(total, maxPoints);
  return Math.round(100 * (1 - safeTotal / maxPoints));
}

function persistSystemTabFromDOM(contentEl, data) {
  if (!contentEl || !data) return;

  ensureSystemDataShape(data);
  const sd = data.systemData;

  const cards = contentEl.querySelectorAll('.system-card');
  const nextSystems = [];

  cards.forEach((card) => {
    const sysId = String(card.dataset.sysId || '');
    if (!sysId) return;

    const name = card.querySelector('.sys-name')?.value || '';
    const isLegacy = !!card.querySelector('.sys-legacy')?.checked;
    const futureSystem = card.querySelector('.sys-future')?.value || '';

    const qa = {};
    SYSTEM_QUESTIONS.forEach((q) => {
      const sel = `input[name="sys_${CSS.escape(sysId)}_${CSS.escape(q.id)}"]`;
      const vStr = contentEl.querySelector(sel)?.value ?? '';
      if (vStr === '') {
        qa[q.id] = null;
        return;
      }
      const n = parseInt(vStr, 10);
      qa[q.id] = Number.isFinite(n) ? n : null;
    });

    const score = computeSystemScoreFromAnswers(
      Object.fromEntries(Object.entries(qa).filter(([, v]) => Number.isFinite(Number(v))))
    );

    nextSystems.push({
      id: sysId,
      name: String(name),
      isLegacy,
      futureSystem: String(futureSystem),
      qa,
      calculatedScore: score
    });
  });

  if (nextSystems.length) sd.systems = nextSystems;

  sd.systemName = sd.systems?.[0]?.name || '';

  const scores = (sd.systems || []).map((s) => s.calculatedScore).filter((x) => x != null);
  sd.calculatedScore =
    scores.length === 0 ? null : Math.round(scores.reduce((a, b) => a + b, 0) / scores.length);
}

const renderSystemTab = (data) => {
  ensureSystemDataShape(data);

  const sd = data.systemData || {};
  const isMulti = !!sd.isMulti;
  const systems = Array.isArray(sd.systems) ? sd.systems : [];
  const noSystem = !!sd.noSystem;

  let html = `
    <div id="systemWrapper">
      ${renderSystemMergeControls()}

      <div class="io-helper">
        Geef aan hoe goed het systeem het proces ondersteunt. Werk je in meerdere systemen? Voeg ze toe en beantwoord System Fit per systeem.
      </div>

      <label style="display:flex; align-items:center; gap:10px; font-size:13px; margin: 12px 0 10px 0; cursor:pointer;">
        <input id="sysNoSystem" type="checkbox" ${noSystem ? 'checked' : ''} />
        Voor dit proces werk ik niet in een systeem
      </label>

      <label style="display:flex; align-items:center; gap:10px; font-size:13px; margin: 6px 0 14px 0; cursor:pointer; ${noSystem ? 'opacity:0.55; pointer-events:none;' : ''}">
        <input id="sysMultiEnable" type="checkbox" ${isMulti ? 'checked' : ''} ${noSystem ? 'disabled' : ''} />
        Ik werk in meerdere systemen binnen deze processtap
      </label>

      <div id="sysList" style="display:grid; gap:14px; ${noSystem ? 'opacity:0.55; pointer-events:none; filter:grayscale(0.2);' : ''}">
  `;

  systems.forEach((sys, idx) => {
    const sysId = sys.id || `${idx}`;
    const name = sys.name || '';
    const isLegacy = !!sys.isLegacy;
    const future = sys.futureSystem || '';
    const qa = sys.qa || {};

    html += `
      <div class="system-card" data-sys-id="${escapeAttr(sysId)}"
           style="background: rgba(0,0,0,0.18); border: 1px solid rgba(255,255,255,0.10); border-radius: 12px; padding: 12px;">
        <div style="display:flex; justify-content:space-between; align-items:center; gap:10px; margin-bottom:10px;">
          <div style="font-weight:900; font-size:12px; letter-spacing:.6px; text-transform:uppercase; opacity:.9;">
            ${isMulti ? `Systeem ${idx + 1}` : 'Systeem'}
          </div>

          <button class="std-btn danger-text" type="button"
                  data-action="remove-system"
                  ${(noSystem || systems.length <= 1) ? 'disabled' : ''}
                  style="padding:6px 10px; font-size:12px;">
            Verwijderen
          </button>
        </div>

        <div style="display:grid; gap:10px;">
          <div style="display:grid; grid-template-columns: 1fr; gap:6px;">
            <div class="modal-label" style="margin:0; font-size:11px;">Systeemnaam</div>
            <input class="modal-input sys-name" type="text" value="${escapeAttr(
              name
            )}" placeholder="Bijv. ARIA / EPIC / Radiotherapieweb / Monaco..." ${noSystem ? 'disabled' : ''} />
          </div>

          <div style="display:grid; grid-template-columns: 1fr 1fr; gap:10px;">
            <label style="display:flex; align-items:center; gap:10px; font-size:13px; cursor:pointer; margin:0;">
              <input class="sys-legacy" type="checkbox" ${isLegacy ? 'checked' : ''} ${noSystem ? 'disabled' : ''} />
              Legacy systeem
            </label>

            <div style="display:grid; grid-template-columns: 1fr; gap:6px;">
              <div class="modal-label" style="margin:0; font-size:11px;">Toekomstig systeem (verwachting)</div>
              <input class="modal-input sys-future" type="text" value="${escapeAttr(
                future
              )}" placeholder="Bijv. ARIA / EPIC / nieuw portaal..." ${noSystem ? 'disabled' : ''} />
            </div>
          </div>
        </div>

        <div style="margin-top:12px; padding-top:12px; border-top:1px solid rgba(255,255,255,0.08);">
          <div class="modal-label" style="margin-top:0;">System Fit vragen</div>
          <div class="io-helper" style="margin-top:6px; margin-bottom:10px;">
            Beantwoord per vraag hoe goed dit systeem jouw taak ondersteunt.
          </div>

          <div style="${noSystem ? 'opacity:0.6; pointer-events:none; filter:grayscale(0.2);' : ''}">
            ${SYSTEM_QUESTIONS
              .map((q) => {
                const currentVal = qa[q.id] !== undefined ? qa[q.id] : null;
                const optionsMapped = q.options.map((optText, optIdx) => ({
                  value: optIdx,
                  label: optText
                }));

                return `
                <div class="system-question" style="margin-bottom:12px;">
                  <div class="sys-q-title">${escapeAttr(q.label)}</div>
                  ${createRadioGroup(
                    `sys_${escapeAttr(sysId)}_${escapeAttr(q.id)}`,
                    optionsMapped,
                    currentVal,
                    true
                  )}
                </div>
              `;
              })
              .join('')}
          </div>

          <div style="margin-top:10px; font-size:12px; opacity:.9;">
            Score (dit systeem): <span class="sys-score" data-sys-score="${escapeAttr(sysId)}">${
              Number.isFinite(Number(sys.calculatedScore)) ? `${sys.calculatedScore}%` : '—'
            }</span>
          </div>
        </div>
      </div>
    `;
  });

  html += `
      </div>

      <button class="std-btn primary" id="btnAddSystem" data-action="add-system" type="button" ${noSystem ? 'disabled' : ''} style="margin-top:14px; display:${
        isMulti ? 'inline-flex' : 'none'
      }; ${noSystem ? 'opacity:0.55; pointer-events:none;' : ''}">
        + Systeem toevoegen
      </button>

      <div style="margin-top:14px; font-size:12px; opacity:.9;">
        Overall score (kolom): <strong id="sysOverallScore">${
          Number.isFinite(Number(sd.calculatedScore)) ? `${sd.calculatedScore}%` : '—'
        }</strong>
      </div>
    </div>
  `;

  return html;
};

/* =========================================================
   PROCESS TAB
   ========================================================= */

const renderProcessTab = (data) => {
  const status = data.processStatus;
  const isHappy = status === 'HAPPY';

  const showSuccessFactors = isHappy;
  const showRootCauses = !isHappy;

  const statusHtml = PROCESS_STATUSES
    .map((s) => {
      const def = PROCESS_STATUS_DEFS[s.value] || {};
      const tTitle = def.title || s.label || '';
      const tBody = def.body || '';
      return `
      <div class="status-option ${status === s.value ? s.class : ''}"
           data-action="set-status"
           data-val="${escapeAttr(s.value)}"
           data-tt-title="${escapeAttr(tTitle)}"
           data-tt-body="${escapeAttr(tBody)}"
           tabindex="0"
           role="button"
           aria-label="${escapeAttr(tTitle)}">
        <span class="status-emoji">${escapeAttr(s.emoji)}</span>
        <span class="status-text">${escapeAttr(s.label)}</span>
      </div>
    `;
    })
    .join('');

  const workExp = data.workExp || null;
  const workExpHtml = WORK_EXP_OPTIONS
    .map(
      (o) => `
      <div class="status-option ${workExp === o.value ? o.cls : ''}"
           data-action="set-workexp"
           data-val="${escapeAttr(o.value)}"
           data-tt-title="${escapeAttr(o.title)}"
           data-tt-body="${escapeAttr(o.body)}"
           tabindex="0"
           role="button"
           aria-label="${escapeAttr(o.title)}">
        <span class="status-emoji">${escapeAttr(o.icon)}</span>
        <span class="status-text">${escapeAttr(o.label)}</span>
      </div>
  `
    )
    .join('');

  const leanValue = data.processValue || null;
  const leanValueHtml = LEAN_VALUES
    .map((opt) => {
      const val = opt.value;
      const label = opt.label;

      let displayLabel = label;
      if (val === 'VA' && !label.includes('💚')) displayLabel = '💚 ' + label;
      if (val === 'BNVA' && !label.includes('⚖️')) displayLabel = '⚖️ ' + label;
      if (val === 'NVA' && !label.includes('🗑️')) displayLabel = '🗑️ ' + label;

      const def = LEAN_VALUE_DEFS[val] || {};
      const isSelected = leanValue === val;

      return `
      <div class="sys-opt ${isSelected ? 'selected' : ''}" 
           data-value="${escapeAttr(val)}"
           data-tt-title="${escapeAttr(def.title || label)}"
           data-tt-body="${escapeAttr(def.body || '')}">
        ${escapeAttr(displayLabel)}
      </div>
    `;
    })
    .join('');

  const activityType = data.type || null;
  const activityTypeHtml = ACTIVITY_TYPES
    .map((opt) => {
      const val = opt.value;
      const label = opt.label;
      const def = ACTIVITY_TYPE_DEFS[val] || {};
      const isSelected = activityType === val;

      let displayLabel = label;
      if (val === 'Taak' && !label.includes('📝')) displayLabel = '📝 ' + label;
      if (val === 'Afspraak' && !label.includes('📅')) displayLabel = '📅 ' + label;
      if (val === 'TimeOut' && !label.includes('🛑')) displayLabel = '🛑 ' + label;
      if (val === 'Beoordeling' && !label.includes('🔎')) displayLabel = '🔎 ' + label;

      return `
      <div class="sys-opt ${isSelected ? 'selected' : ''}"
           data-value="${escapeAttr(val)}"
           data-tt-title="${escapeAttr(def.title || label)}"
           data-tt-body="${escapeAttr(def.body || '')}">
        ${escapeAttr(displayLabel)}
      </div>
    `;
    })
    .join('');

  const disruptions =
    data.disruptions && data.disruptions.length > 0
      ? data.disruptions
      : [{ scenario: '', frequency: null, workaround: '' }];

  const disruptRows = disruptions
    .map(
      (dis, i) => `
        <tr>
          <td><input class="def-input" value="${escapeAttr(dis.scenario || '')}" placeholder="Scenario..."></td>
          <td>${createRadioGroup(`dis_freq_${i}`, DISRUPTION_FREQUENCIES, dis.frequency, false)}</td>
          <td><input class="def-input" value="${escapeAttr(dis.workaround || '')}" placeholder="Workaround..."></td>
          <td><button class="btn-row-del-tiny" data-action="remove-row" type="button">×</button></td>
        </tr>
      `
    )
    .join('');

  const analysisRows = createCombinedAnalysisRows(data.causes, data.improvements);

  return `
    <div class="modal-label">Type Activiteit ${!data.type ? '<span style="color:#ff5252">*</span>' : ''}</div>
    <div class="radio-group-container horizontal">
        ${activityTypeHtml}
        <input type="hidden" name="metaType" value="${escapeAttr(activityType || '')}">
    </div>

    <div class="modal-label" style="margin-top:16px;">Werkbeleving (Werkplezier)</div>
    <div class="io-helper" style="margin-top:0; margin-bottom:12px; font-size:13px;">
      Kies wat dit met je doet (en de bijbehorende actie-richting).
    </div>
    <div class="status-selector">${workExpHtml}</div>
    <input type="hidden" id="workExp" value="${escapeAttr(workExp || '')}">
    <textarea id="workExpNote" class="modal-input" placeholder="Korte context (optioneel): wat maakt dit een obstakel/routine/flow?">${escapeAttr(
      data.workExpNote || ''
    )}</textarea>

    <div class="modal-label" style="margin-top:24px;">Lean Waarde ${
      !data.processValue ? '<span style="color:#ff5252">*</span>' : ''
    }</div>
    <div class="radio-group-container horizontal">
        ${leanValueHtml}
        <input type="hidden" name="metaValue" value="${escapeAttr(leanValue || '')}">
    </div>

    <div class="modal-label" style="margin-top:24px;">Proces Status ${
      !status ? '<span style="color:#ff5252">*</span>' : ''
    }</div>
    <div class="status-selector">${statusHtml}</div>
    <input type="hidden" id="processStatus" value="${escapeAttr(status || '')}">

    <div id="sectionAnalyse" style="margin-top:24px; padding-top:20px; border-top:1px solid rgba(255,255,255,0.1);">
        
        <div id="wrapperSuccess" style="display: ${showSuccessFactors ? 'block' : 'none'};">
           <div class="modal-label" style="color:var(--ui-success)">Waarom werkt dit goed? (Succesfactoren)</div>
           <textarea id="successFactors" class="modal-input" placeholder="Bv. Standaard protocol gevolgd...">${escapeAttr(
             data.successFactors || ''
           )}</textarea>
        </div>

        <div id="wrapperIssues" style="display: ${showRootCauses ? 'block' : 'none'};">
           <div class="modal-label">Analyse (Oorzaken & Maatregelen)</div>
           <div class="io-helper">Vul per regel de oorzaak en de bijbehorende maatregel in.</div>
           
           <table class="io-table proc-table">
             <thead>
               <tr>
                 <th style="width:45%">Oorzaak (Root Cause)</th>
                 <th style="width:45%">Maatregel (Countermeasure)</th>
                 <th></th>
               </tr>
             </thead>
             <tbody id="analysisTbody">
               ${analysisRows}
             </tbody>
           </table>
           <button class="btn-row-add" data-action="add-analysis-row" type="button">+ Analyse regel toevoegen</button>
        </div>
    </div>

    <div id="sectionDisrupt" style="margin-top:24px; padding-top:20px; border-top:1px solid rgba(255,255,255,0.1);">
        <div class="modal-label">Verstoringen & Workarounds</div>
        <div class="io-helper">Welke verstoringen treden op en wat is de workaround? (Ook invullen als proces in control is)</div>
        <table class="io-table proc-table">
          <thead><tr><th style="width:30%">Scenario</th><th style="width:25%">Frequentie</th><th style="width:40%">Workaround</th><th></th></tr></thead>
          <tbody id="disruptTbody">${disruptRows}</tbody>
        </table>
        <button class="btn-row-add" data-action="add-disrupt-row" type="button">+ Verstoring regel toevoegen</button>
    </div>
  `;
};

/* =========================================================
   IO TAB (Input / Output)
   ========================================================= */

const renderIoTab = (data, isInputRow) => {
  let linkHtml = '';

  if (isInputRow) {
    const allDetails = state.getAllOutputsDetailed();
    const uidMap = getOutputUidMap();

    const selectedUids = normalizeLinkedSourceIds(data);
    const selectedSet = new Set(selectedUids);

    const activeBundleId = getActiveBundleId(data);
    const hasBundle = !!activeBundleId;
    const bundle = hasBundle ? findBundleById(activeBundleId) : null;
    const bundleName = String(bundle?.name || '').trim();

    const listItems = allDetails
      .map((item) => {
        const id = item.uid;
        const outLabel = item.outId;
        const text = (item.text || '').substring(0, 70);

        const checked = selectedSet.has(id);

        return `
          <label style="display:flex; align-items:flex-start; gap:10px; padding:8px 10px; border-radius:8px; cursor:pointer; background: rgba(255,255,255,0.04); border: 1px solid rgba(255,255,255,0.08);">
            <input class="input-source-cb" type="checkbox" value="${escapeAttr(id)}" ${checked ? 'checked' : ''} style="margin-top:2px;" />
            <div style="display:flex; flex-direction:column; gap:3px;">
              <div style="font-weight:900; font-size:12px; letter-spacing:.3px;">${escapeAttr(outLabel)}</div>
              <div style="font-size:11px; opacity:.85; line-height:1.35;">${escapeAttr(text)}${
                (item.text || '').length > 70 ? '...' : ''
              }</div>
            </div>
          </label>
        `;
      })
      .join('');

    const extChecked = !hasBundle && selectedUids.length === 0;

    const summaryText = hasBundle
      ? `📦 ${bundleName || 'Pakket'}`
      : selectedUids.length
        ? escapeAttr(_formatLinkedSources(selectedUids, uidMap))
        : '—';

    linkHtml = `
      <div style="margin-bottom: 20px; padding: 12px; background: rgba(0,0,0,0.2); border-radius: 8px; border: 1px solid rgba(255,255,255,0.1);">
        <div class="modal-label" style="margin-top:0;">Input Bron (koppel aan 1+ Outputs)</div>

        <div class="io-helper" style="margin-top:6px; margin-bottom:12px;">
          Selecteer meerdere outputs als deze processtap meerdere inputs nodig heeft.
        </div>

        <div style="display:flex; align-items:center; gap:10px; margin-bottom:12px;">
          <label style="display:flex; align-items:center; gap:10px; font-size:13px; cursor:pointer; margin:0;">
            <input id="inputExternalToggle" type="checkbox" ${extChecked ? 'checked' : ''} />
            Geen koppeling (externe input)
          </label>

          <div style="margin-left:auto; font-size:12px; opacity:.85;">
            Geselecteerd: <span id="linkedSourcesSummary" style="font-weight:900; color: var(--ui-accent);">${summaryText}</span>
          </div>
        </div>

        <div id="inputSourcesList" style="display:grid; gap:8px; max-height: 260px; overflow:auto; padding-right:4px; ${
          (hasBundle || selectedUids.length === 0)
            ? 'opacity:0.55; pointer-events:none; filter: grayscale(0.2);'
            : ''
        }">
          ${listItems || '<div style="opacity:.7; font-size:12px;">Geen outputs beschikbaar.</div>'}
        </div>

        <div id="linkedInfoText" style="display:${selectedUids.length ? 'block' : 'none'}; color:var(--ui-accent); font-size:11px; margin-top:10px; font-weight:700;">
          🔗 Gekoppeld (${selectedUids.length}): ${escapeAttr(_formatLinkedSources(selectedUids, uidMap))}
        </div>

        <div id="bundleActiveInfo" style="display:none; color:var(--ui-accent); font-size:11px; margin-top:10px; font-weight:800;"></div>

        <div style="margin-top:14px; padding-top:12px; border-top:1px dashed rgba(255,255,255,0.14);"></div>

        <div class="modal-label" style="margin-top:0;">Output-pakket (bundel)</div>
        <div class="io-helper" style="margin-top:6px; margin-bottom:12px;">
          Geef een naam aan een set outputs (bijv. <strong>Verwijspakket</strong>) en gebruik die als 1 input.
        </div>

        <label class="modal-label" style="font-size:11px; margin-top:0;">Bestaand pakket (optioneel)</label>
        <select id="bundlePick" class="modal-input">
          <option value="">— Nieuw pakket —</option>
          ${(() => {
            const project = state.data;
            const bundles = Array.isArray(project?.outputBundles) ? project.outputBundles : [];
            return bundles
              .map((b) => {
                const id = String(b?.id || '').trim();
                const nm = String(b?.name || '').trim() || id;
                return `<option value="${escapeAttr(id)}">${escapeAttr(nm)}</option>`;
              })
              .join('');
          })()}
        </select>

        <label class="modal-label" style="font-size:11px; margin-top:10px;">Pakketnaam</label>
        <input id="bundleName" class="modal-input" placeholder="Bijv. Verwijspakket" />

        <div style="display:flex; gap:10px; margin-top:10px; flex-wrap:wrap;">
          <button id="bundleApplyBtn" class="std-btn primary" type="button">Pakket opslaan/gebruiken</button>
          <button id="bundleClearBtn" class="std-btn" type="button">Loskoppelen</button>
          
          <button id="bundleDeleteDefBtn" class="std-btn danger-text" type="button" style="margin-left:auto; border:1px solid #ff5252; color:#ff5252; display:none;">
              🗑️ Pakket definitief verwijderen
          </button>
        </div>

        <div id="bundleHint" style="margin-top:8px; font-size:12px; opacity:0.8;"></div>
      </div>
    `;
  }

  const mergeControls = !isInputRow ? renderOutputMergeControls() : '';

  const qaRows = IO_CRITERIA.map((c) => {
    const qa = data.qa?.[c.key] || {};
    const defs = IO_CRITERIA_DEFS[c.key] || {};
    const currentRes = qa.result; // GOOD, MINOR, MODERATE, FAIL

    // We tonen het opmerkingenveld en de impact opties als het NIET 'GOOD' is
    const showImpact = ['MINOR', 'MODERATE', 'FAIL', 'POOR'].includes(currentRes);

    // Bepaal de impact value obv resultaat
    let impactValue = null;
    if (currentRes === 'FAIL' || currentRes === 'POOR') impactValue = 'A';
    if (currentRes === 'MODERATE') impactValue = 'B';
    if (currentRes === 'MINOR') impactValue = 'C';

    return `
      <div class="qa-item-wrapper" style="margin-bottom: 24px; padding-bottom: 24px; border-bottom: 1px solid rgba(255,255,255,0.05);">
        <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:12px;">
            <div style="max-width: 50%;">
                <div style="font-weight:bold; color:#fff; font-size:14px; margin-bottom:4px;">${escapeAttr(
                  c.label
                )}</div>
                <div style="font-size:12px; opacity:0.8; line-height:1.4;">${escapeAttr(
                  defs.gating || c.meet
                )}</div>
            </div>
            
            <div style="flex-shrink:0;">
                 ${createRadioGroup(
                   `qa_gate_${c.key}`,
                   [
                     { value: 'GOOD', label: 'Voldoet' },
                     { value: 'MINOR', label: 'Grotendeels' },
                     { value: 'MODERATE', label: 'Matig' },
                     { value: 'FAIL', label: 'Voldoet niet' }
                   ],
                   currentRes,
                   false
                 )}
            </div>
        </div>

        <div id="impact_wrapper_${escapeAttr(c.key)}" style="display:${
      showImpact ? 'block' : 'none'
    }; background: rgba(255,82,82, 0.1); border-left: 3px solid #ff5252; padding: 12px; margin-top:8px; border-radius: 0 4px 4px 0;">
            <div style="font-size:11px; font-weight:bold; color:#ff5252; text-transform:uppercase; margin-bottom:8px;">Wat is de impact op jouw taak?</div>
            
            <div class="radio-group-container vertical">
                <div class="sys-opt impact-opt ${impactValue === 'A' ? 'selected' : ''}" 
                     data-value="A" 
                     data-key="${escapeAttr(c.key)}"
                     style="text-align:left; height:auto; padding:8px 12px; margin-bottom:6px;">
                    <div style="font-weight:bold; font-size:12px;">🔴 A. Blokkerend</div>
                    <div style="font-size:11px; opacity:0.8; font-weight:normal;">${escapeAttr(
                      defs.impactA
                    )}</div>
                </div>
                
                <div class="sys-opt impact-opt ${impactValue === 'B' ? 'selected' : ''}" 
                     data-value="B" 
                     data-key="${escapeAttr(c.key)}"
                     style="text-align:left; height:auto; padding:8px 12px; margin-bottom:6px;">
                    <div style="font-weight:bold; font-size:12px;">🟠 B. Extra werk</div>
                    <div style="font-size:11px; opacity:0.8; font-weight:normal;">${escapeAttr(
                      defs.impactB
                    )}</div>
                </div>

                <div class="sys-opt impact-opt ${impactValue === 'C' ? 'selected' : ''}" 
                     data-value="C" 
                     data-key="${escapeAttr(c.key)}"
                     style="text-align:left; height:auto; padding:8px 12px;">
                    <div style="font-weight:bold; font-size:12px;">🟡 C. Kleine frictie</div>
                    <div style="font-size:11px; opacity:0.8; font-weight:normal;">${escapeAttr(
                      defs.impactC
                    )}</div>
                </div>
            </div>
            <input type="hidden" name="qa_impact_${escapeAttr(c.key)}" value="${escapeAttr(impactValue || '')}">
        </div>

        <div id="note_wrapper_${escapeAttr(c.key)}" style="display:${showImpact ? 'block' : 'none'};">
            <textarea id="note_${escapeAttr(c.key)}" class="io-note" style="margin-top:12px; width:100%;" placeholder="Opmerking (optioneel)...">${escapeAttr(
              qa.note || ''
            )}</textarea>
        </div>
      </div>
    `;
  }).join('');

  const definitions =
    data.inputDefinitions && data.inputDefinitions.length > 0
      ? data.inputDefinitions
      : [{ item: '', specifications: '', type: null }];

  const defRows = definitions
    .map(
      (def, i) => `
        <tr>
          <td><input class="def-input" value="${escapeAttr(def.item || '')}" placeholder="Naam item..."></td>
          <td><textarea class="def-sub-input" placeholder="Specificaties...">${escapeAttr(
            def.specifications || ''
          )}</textarea></td>
          <td>${createRadioGroup(`def_type_${i}`, DEFINITION_TYPES, def.type, true)}</td>
          <td><button class="btn-row-del-tiny" data-action="remove-row" type="button">×</button></td>
        </tr>
      `
    )
    .join('');

  return `
    ${mergeControls}
    ${linkHtml}

    <div class="modal-label">1. Kwaliteits Criteria (System Fit)</div>
    <div class="io-helper" style="margin-bottom:20px;">
       Beoordeel of de ${isInputRow ? 'input' : 'output'} voldoende is om de taak zonder problemen uit te voeren.
    </div>
    
    <div id="ioTabQual">
        ${qaRows}
    </div>

    <div style="margin-top:32px; padding-top:24px; border-top:1px solid rgba(255,255,255,0.1);">
      <div class="modal-label">2. Definitie & Specificaties</div>
      <table class="io-table def-table">
        <thead><tr><th style="width:25%">Item</th><th style="width:40%">Specificaties</th><th style="width:30%">Type</th><th></th></tr></thead>
        <tbody id="defTbody">${defRows}</tbody>
      </table>
      <button class="btn-row-add" data-action="add-def-row" type="button">+ Specificatie regel toevoegen</button>
    </div>
  `;
};

/* =========================================================
   RENDER MODAL CONTENT
   ========================================================= */

const renderContent = () => {
  const data = getStickyData();
  if (!data) return;

  const slotIdx = editingSticky.slotIdx;
  const content = $('modalContent');
  const title = $('modalTitle');
  if (!content || !title) return;

  if (slotIdx === 3) {
    title.textContent = 'Proces Stap Analyse';
    content.innerHTML = renderProcessTab(data);
    return;
  }

  if (slotIdx === 1) {
    title.textContent = 'Systeem Fit Analyse';
    content.innerHTML = renderSystemTab(data);
    return;
  }

  if (slotIdx === 2 || slotIdx === 4) {
    if (slotIdx === 2) title.textContent = 'Input Specificaties';
    if (slotIdx === 4) title.textContent = 'Output Specificaties';
    content.innerHTML = renderIoTab(data, slotIdx === 2);
  }
};

/* =========================================================
   TOOLTIP
   ========================================================= */

let _ttEl = null;
let _ttVisible = false;

function ensureTooltipEl() {
  if (_ttEl) return _ttEl;

  const el = document.createElement('div');
  el.id = 'customTooltip';
  el.style.position = 'fixed';
  el.style.left = '0px';
  el.style.top = '0px';
  el.style.transform = 'translate(-9999px, -9999px)';
  el.style.zIndex = '10000';
  el.style.pointerEvents = 'none';
  el.style.opacity = '0';
  el.style.transition = 'opacity 120ms ease, transform 120ms ease';
  el.style.maxWidth = '360px';
  el.style.padding = '10px 12px';
  el.style.borderRadius = '10px';
  el.style.background = 'rgba(20, 24, 28, 0.95)';
  el.style.border = '1px solid rgba(255,255,255,0.12)';
  el.style.boxShadow = '0 10px 30px rgba(0,0,0,0.45)';
  el.style.backdropFilter = 'blur(10px)';
  el.style.color = '#fff';
  el.style.fontFamily = '"Inter", sans-serif';
  el.style.fontSize = '12px';
  el.style.lineHeight = '1.4';

  el.innerHTML = `
    <div style="font-weight:900; font-size:11px; letter-spacing:.6px; text-transform:uppercase; opacity:.9; margin-bottom:6px;" data-tt="title"></div>
    <div style="opacity:.85" data-tt="body"></div>
  `;

  document.body.appendChild(el);
  _ttEl = el;
  return el;
}

function showTooltip(target, x, y) {
  const el = ensureTooltipEl();

  const title = target?.dataset?.ttTitle || '';
  const body = target?.dataset?.ttBody || '';
  if (!title && !body) return;

  const tEl = el.querySelector('[data-tt="title"]');
  const bEl = el.querySelector('[data-tt="body"]');
  if (tEl) tEl.textContent = title;
  if (bEl) bEl.textContent = body;

  const pad = 12;
  const vw = window.innerWidth;
  const vh = window.innerHeight;

  el.style.opacity = '1';
  _ttVisible = true;

  const rect = el.getBoundingClientRect();
  let left = x + pad;
  let top = y + pad;

  if (left + rect.width + 8 > vw) left = x - rect.width - pad;
  if (top + rect.height + 8 > vh) top = y - rect.height - pad;

  left = Math.max(8, Math.min(vw - rect.width - 8, left));
  top = Math.max(8, Math.min(vh - rect.height - 8, top));

  el.style.transform = `translate(${Math.round(left)}px, ${Math.round(top)}px)`;
}

function hideTooltip() {
  if (!_ttEl || !_ttVisible) return;
  _ttEl.style.opacity = '0';
  _ttEl.style.transform = 'translate(-9999px, -9999px)';
  _ttVisible = false;
}

/* =========================================================
   LISTENERS
   ========================================================= */

const setupPermanentListeners = () => {
  const modal = $('editModal');
  const content = $('modalContent');
  const saveBtn = $('modalSaveBtn');
  const cancelBtn = $('modalCancelBtn');

  if (saveBtn) saveBtn.onclick = () => saveModalDetails(true);
  if (cancelBtn && modal) cancelBtn.onclick = () => (modal.style.display = 'none');
  if (!content) return;

  const getTooltipTarget = (e) => e.target.closest('.status-option, .sys-opt[data-tt-title]');

  content.addEventListener(
    'pointerenter',
    (e) => {
      const opt = getTooltipTarget(e);
      if (!opt) return;
      showTooltip(opt, e.clientX, e.clientY);
    },
    true
  );

  content.addEventListener(
    'pointermove',
    (e) => {
      const opt = getTooltipTarget(e);
      if (!opt) {
        hideTooltip();
        return;
      }
      showTooltip(opt, e.clientX, e.clientY);
    },
    true
  );

  content.addEventListener(
    'pointerleave',
    (e) => {
      const opt = getTooltipTarget(e);
      if (!opt) return;
      hideTooltip();
    },
    true
  );

  content.addEventListener('scroll', () => hideTooltip(), { passive: true });
  if (modal) modal.addEventListener('scroll', () => hideTooltip(), { passive: true });

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') hideTooltip();
  });

  // Output merge enable toggle
  content.addEventListener('change', (e) => {
    if (e.target?.id !== 'mergeEnable') return;
    const enabled = !!e.target.checked;
    const s = $('mergeStartSelect');
    const en = $('mergeEndSelect');
    if (s) s.disabled = !enabled;
    if (en) en.disabled = !enabled;

    const hint = $('mergeHint');
    if (hint) hint.textContent = enabled ? hint.textContent || 'Actief' : 'Niet gemerged';
  });

  // Output merge select change -> update hint live
  content.addEventListener('change', (e) => {
    if (e.target?.id !== 'mergeStartSelect' && e.target?.id !== 'mergeEndSelect') return;
    const hint = $('mergeHint');
    if (!hint || !editingSticky) return;

    const colIdx = editingSticky.colIdx;
    const startCol = parseInt($('mergeStartSelect')?.value ?? `${colIdx}`, 10);
    const endCol = parseInt($('mergeEndSelect')?.value ?? `${colIdx}`, 10);
    const s = Math.min(startCol, endCol);
    const ee = Math.max(startCol, endCol);

    const ok = s <= colIdx && ee >= colIdx && s !== ee;
    hint.textContent = ok ? `Actief: kolom ${s + 1} t/m ${ee + 1}` : `⚠ Ongeldig: range moet huidige kolom bevatten`;
  });

  // System merge enable toggle
  content.addEventListener('change', (e) => {
    if (e.target?.id !== 'sysMergeEnable') return;
    const enabled = !!e.target.checked;
    const s = $('sysMergeStartSelect');
    const en = $('sysMergeEndSelect');
    if (s) s.disabled = !enabled;
    if (en) en.disabled = !enabled;

    const hint = $('sysMergeHint');
    if (hint) hint.textContent = enabled ? hint.textContent || 'Actief' : 'Niet gemerged';
  });

  // System merge select change -> update hint live
  content.addEventListener('change', (e) => {
    if (e.target?.id !== 'sysMergeStartSelect' && e.target?.id !== 'sysMergeEndSelect') return;
    const hint = $('sysMergeHint');
    if (!hint || !editingSticky) return;

    const colIdx = editingSticky.colIdx;
    const startCol = parseInt($('sysMergeStartSelect')?.value ?? `${colIdx}`, 10);
    const endCol = parseInt($('sysMergeEndSelect')?.value ?? `${colIdx}`, 10);
    const s = Math.min(startCol, endCol);
    const ee = Math.max(startCol, endCol);

    const ok = s <= colIdx && ee >= colIdx && s !== ee;
    hint.textContent = ok ? `Actief: kolom ${s + 1} t/m ${ee + 1}` : `⚠ Ongeldig: range moet huidige kolom bevatten`;
  });

  // Multi-system toggle + add/remove system rows
  content.addEventListener('change', (e) => {
    if (e.target?.id !== 'sysMultiEnable') return;

    const data = getStickyData();
    if (!data) return;

    ensureSystemDataShape(data);
    if (data.systemData.noSystem) {
      e.target.checked = false;
      data.systemData.isMulti = false;
      state.saveStickyDetails();
      renderContent();
      return;
    }

    persistSystemTabFromDOM(content, data);

    data.systemData.isMulti = !!e.target.checked;

    state.saveStickyDetails();
    renderContent();
  });

  // No-system toggle
  content.addEventListener('change', (e) => {
    if (e.target?.id !== 'sysNoSystem') return;

    const data = getStickyData();
    if (!data) return;

    ensureSystemDataShape(data);

    const checked = !!e.target.checked;
    data.systemData.noSystem = checked;

    if (checked) {
      data.systemData.isMulti = false;
      data.systemData.systems = [
        {
          id: `${Date.now()}_${Math.random().toString(16).slice(2)}`,
          name: '',
          isLegacy: false,
          futureSystem: '',
          qa: {},
          calculatedScore: null
        }
      ];
      data.systemData.systemName = '';
      data.systemData.calculatedScore = null;
    }

    state.saveStickyDetails();
    renderContent();
  });

  content.addEventListener('click', (e) => {
    const btnAdd = e.target?.closest?.('#btnAddSystem, [data-action="add-system"]');
    if (btnAdd) {
      const data = getStickyData();
      if (!data) return;

      ensureSystemDataShape(data);
      data.systemData.isMulti = true;

      persistSystemTabFromDOM(content, data);

      data.systemData.systems.push({
        id: `${Date.now()}_${Math.random().toString(16).slice(2)}`,
        name: '',
        isLegacy: false,
        futureSystem: '',
        qa: {},
        calculatedScore: null
      });

      state.saveStickyDetails();
      renderContent();
      return;
    }

    const btnRemove = e.target?.closest?.('[data-action="remove-system"]');
    if (btnRemove) {
      const card = btnRemove.closest('.system-card');
      const sysId = card?.dataset?.sysId;
      const data = getStickyData();
      if (!data || !sysId) return;

      ensureSystemDataShape(data);
      if (data.systemData.systems.length <= 1) return;

      persistSystemTabFromDOM(content, data);

      data.systemData.systems = data.systemData.systems.filter((s) => String(s.id) !== String(sysId));
      if (data.systemData.systems.length === 0) {
        data.systemData.systems = [
          {
            id: `${Date.now()}_${Math.random().toString(16).slice(2)}`,
            name: '',
            isLegacy: false,
            futureSystem: '',
            qa: {},
            calculatedScore: null
          }
        ];
      }

      state.saveStickyDetails();
      renderContent();
      return;
    }
  });

  // Sys-opt radio buttons
  content.addEventListener('click', (e) => {
    const opt = e.target.closest('.sys-opt');
    if (!opt) return;

    if (opt.classList.contains('impact-opt')) {
      handleImpactClick(opt);
      return;
    }

    const container = opt.closest('.radio-group-container');
    const input = container?.querySelector('input[type="hidden"]');
    const wasSelected = opt.classList.contains('selected');

    container?.querySelectorAll('.sys-opt').forEach((el) => el.classList.remove('selected'));

    if (wasSelected) {
      if (input) input.value = '';
      if (input && input.name.startsWith('qa_gate_')) {
        const key = input.name.replace('qa_gate_', '');
        const impactWrapper = $(`impact_wrapper_${key}`);
        const noteWrapper = $(`note_wrapper_${key}`);
        if (impactWrapper) impactWrapper.style.display = 'none';
        if (noteWrapper) noteWrapper.style.display = 'none';
        
        const impactInput = content.querySelector(`input[name="qa_impact_${key}"]`);
        if (impactInput) impactInput.value = '';
        content.querySelectorAll(`.impact-opt[data-key="${CSS.escape(key)}"]`).forEach(el => el.classList.remove('selected'));
        
        const noteInput = content.querySelector(`#note_${key}`);
        if (noteInput) noteInput.value = '';
      }
      return;
    }

    opt.classList.add('selected');
    if (input) input.value = opt.dataset.value;

    if (input && input.name.startsWith('qa_gate_')) {
      const key = input.name.replace('qa_gate_', '');
      const impactWrapper = $(`impact_wrapper_${key}`);
      const noteWrapper = $(`note_wrapper_${key}`);
      if (!impactWrapper || !noteWrapper) return;

      if (['MINOR', 'MODERATE', 'FAIL'].includes(opt.dataset.value)) {
        impactWrapper.style.display = 'block';
        noteWrapper.style.display = 'block';
      } else {
        impactWrapper.style.display = 'none';
        noteWrapper.style.display = 'none';
        
        const impactInput = content.querySelector(`input[name="qa_impact_${key}"]`);
        if (impactInput) impactInput.value = '';
        content.querySelectorAll(`.impact-opt[data-key="${CSS.escape(key)}"]`).forEach(el => el.classList.remove('selected'));
        
        const noteInput = content.querySelector(`#note_${key}`);
        if (noteInput) noteInput.value = '';
      }
    }

    if (input && input.name.startsWith('sys_')) {
      updateLiveSystemScoresInUI();
    }
  });

  function handleImpactClick(opt) {
    const key = String(opt?.dataset?.key || '').trim();
    const val = String(opt?.dataset?.value || '').trim();
    if (!key) return;

    const group = opt.closest('.radio-group-container') || opt.parentElement;
    const hidden = content.querySelector(`input[name="qa_impact_${CSS.escape(key)}"]`);

    const wasSelected = opt.classList.contains('selected');

    if (group) {
      group
        .querySelectorAll(`.impact-opt[data-key="${CSS.escape(key)}"]`)
        .forEach((el) => el.classList.remove('selected'));
    }

    if (wasSelected) {
      if (hidden) hidden.value = '';
      return;
    }

    opt.classList.add('selected');
    if (hidden) hidden.value = val;
  }

  function updateLiveSystemScoresInUI() {
    const wrapper = $('systemWrapper');
    if (!wrapper) return;

    const cards = wrapper.querySelectorAll('.system-card');
    let overallScores = [];

    cards.forEach((card) => {
      const sysId = card.dataset.sysId;
      const answers = {};

      SYSTEM_QUESTIONS.forEach((q) => {
        const v = wrapper.querySelector(
          `input[name="sys_${CSS.escape(sysId)}_${CSS.escape(q.id)}"]`
        )?.value;
        if (v === '' || v == null) return;
        const n = parseInt(v, 10);
        if (Number.isFinite(n)) answers[q.id] = n;
      });

      const score = computeSystemScoreFromAnswers(answers);
      const scoreEl = wrapper.querySelector(`[data-sys-score="${CSS.escape(sysId)}"]`);
      if (scoreEl) scoreEl.textContent = score == null ? '—' : `${score}%`;

      if (score != null) overallScores.push(score);
    });

    const overallEl = $('sysOverallScore');
    if (overallEl) {
      if (overallScores.length === 0) overallEl.textContent = '—';
      else {
        overallEl.textContent = `${Math.round(
          overallScores.reduce((a, b) => a + b, 0) / overallScores.length
        )}%`;
      }
    }
  }

  // Proces status select
  content.addEventListener('click', (e) => {
    const statusOpt = e.target.closest('.status-option[data-action="set-status"]');
    if (!statusOpt) return;

    const val = statusOpt.dataset.val;
    const configStatus = PROCESS_STATUSES.find((s) => s.value === val);
    const input = $('processStatus');
    if (!configStatus || !input) return;

    const wasActive = statusOpt.classList.contains(configStatus.class);

    PROCESS_STATUSES.forEach((s) => {
      content
        .querySelectorAll(`.status-option.${s.class}[data-action="set-status"]`)
        .forEach((el) => el.classList.remove(s.class));
    });

    if (wasActive) input.value = '';
    else {
      input.value = val;
      statusOpt.classList.add(configStatus.class);
    }

    saveModalDetails(false);
    renderContent();
    hideTooltip();
  });

  // Werkbeleving select
  content.addEventListener('click', (e) => {
    const opt = e.target.closest('.status-option[data-action="set-workexp"]');
    if (!opt) return;

    const input = $('workExp');
    if (!input) return;

    const val = opt.dataset.val;
    const wasActive = (input.value || '') === val;

    content
      .querySelectorAll('.status-option[data-action="set-workexp"]')
      .forEach((el) => el.classList.remove('selected-hap', 'selected-neu', 'selected-sad'));

    if (wasActive) {
      input.value = '';
      hideTooltip();
      return;
    }

    input.value = val;

    const found = WORK_EXP_OPTIONS.find((o) => o.value === val);
    if (found?.cls) opt.classList.add(found.cls);
  });

  // Dynamic lists / tables
  content.addEventListener('click', (e) => {
    const target = e.target;

    if (target?.dataset?.action === 'remove-row') {
      target.closest('.dynamic-row, tr')?.remove();
      return;
    }

    if (target?.dataset?.action === 'add-analysis-row') {
      const tbody = $('analysisTbody');
      if (!tbody) return;

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><input type="text" class="def-input input-cause" placeholder="Oorzaak..."></td>
        <td><input type="text" class="def-input input-measure" placeholder="Maatregel..."></td>
        <td><button class="btn-row-del-tiny" data-action="remove-row" type="button">×</button></td>
      `;
      tbody.appendChild(tr);
      tr.querySelector('input')?.focus();
      return;
    }

    if (target?.dataset?.action === 'add-def-row') {
      const tbody = $('defTbody');
      if (!tbody) return;

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><input class="def-input" placeholder="Naam item..."></td>
        <td><textarea class="def-sub-input" placeholder="Specificaties..."></textarea></td>
        <td>${createRadioGroup(`def_type_new_${Date.now()}`, DEFINITION_TYPES, null, true)}</td>
        <td><button class="btn-row-del-tiny" data-action="remove-row" type="button">×</button></td>
      `;
      tbody.appendChild(tr);
      tr.querySelector('input')?.focus();
      return;
    }

    if (target?.dataset?.action === 'add-disrupt-row') {
      const tbody = $('disruptTbody');
      if (!tbody) return;

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><input class="def-input" placeholder="Scenario..."></td>
        <td>${createRadioGroup(`dis_freq_new_${Date.now()}`, DISRUPTION_FREQUENCIES, null, false)}</td>
        <td><input class="def-input" placeholder="Workaround..."></td>
        <td><button class="btn-row-del-tiny" data-action="remove-row" type="button">×</button></td>
      `;
      tbody.appendChild(tr);
      tr.querySelector('input')?.focus();
    }
  });

  // Input bron (MULTI) — checkboxes + externe toggle
  content.addEventListener('change', (e) => {
    const data = getStickyData();
    if (!data) return;

    if (e.target?.id === 'inputExternalToggle') {
      const checked = !!e.target.checked;

      const list = $('inputSourcesList');
      if (list) {
        list.style.opacity = checked ? '0.55' : '1';
        list.style.pointerEvents = checked ? 'none' : 'auto';
        list.style.filter = checked ? 'grayscale(0.2)' : 'none';
      }

      if (checked) {
        content.querySelectorAll('.input-source-cb').forEach((cb) => {
          cb.checked = false;
          cb.disabled = false;
        });
        setLinkedSourceIds(data, []);
        setActiveBundleId(data, '');
        updateLinkedInfoUI([]);
        updateBundleInfoUI(data);
        state.saveStickyDetails();
      }
      return;
    }

    if (e.target?.classList?.contains('input-source-cb')) {
      if (getActiveBundleId(data)) return;

      const ext = $('inputExternalToggle');
      if (ext) ext.checked = false;

      const ids = Array.from(content.querySelectorAll('.input-source-cb'))
        .filter((cb) => cb.checked)
        .map((cb) => cb.value);

      setActiveBundleId(data, '');
      setLinkedSourceIds(data, ids);

      updateLinkedInfoUI(normalizeLinkedSourceIds(data));
      updateBundleInfoUI(data);

      const list = $('inputSourcesList');
      if (list) {
        const hasAny = ids.length > 0;
        list.style.opacity = hasAny ? '1' : '0.55';
        list.style.pointerEvents = hasAny ? 'auto' : 'none';
        list.style.filter = hasAny ? 'none' : 'grayscale(0.2)';
      }

      state.saveStickyDetails();
    }
  });

  // Output-pakket (bundel) — pick existing
  content.addEventListener('change', (e) => {
    if (e.target?.id !== 'bundlePick') return;
    const data = getStickyData();
    if (!data) return;

    const id = String(e.target.value || '').trim();
    setActiveBundleId(data, id);

    if (id) {
      data.text = '';
      setLinkedSourceIds(data, []);
      updateLinkedInfoUI([]);
    }

    updateBundleInfoUI(data);
    state.saveStickyDetails();
  });

  // Output-pakket (bundel) — create/update/use + clear + DELETE
  content.addEventListener('click', (e) => {
    const data = getStickyData();
    if (!data) return;

    const applyBtn = e.target?.closest?.('#bundleApplyBtn');
    const clearBtn = e.target?.closest?.('#bundleClearBtn');
    const deleteDefBtn = e.target?.closest?.('#bundleDeleteDefBtn');

    if (deleteDefBtn) {
      const pick = content.querySelector('#bundlePick');
      const id = String(pick?.value || '').trim();

      if (!id) return;
      if (!confirm('Weet je zeker dat je dit pakket definitief wilt verwijderen? Het is dan nergens meer beschikbaar.'))
        return;

      const bundles = ensureOutputBundlesArray();
      const project = state.data;
      project.outputBundles = bundles.filter((b) => String(b.id) !== id);

      setActiveBundleId(data, '');
      if (pick) {
        pick.value = '';
        const opt = pick.querySelector(`option[value="${escapeAttr(id)}"]`);
        if (opt) opt.remove();
      }

      const nameEl = content.querySelector('#bundleName');
      if (nameEl) nameEl.value = '';

      updateBundleInfoUI(data);

      const ids = normalizeLinkedSourceIds(data);
      const list = $('inputSourcesList');
      if (list) {
        list.style.opacity = '1';
        list.style.pointerEvents = 'auto';
        list.style.filter = 'none';
      }
      syncOutputCheckboxesFromLinkedSources(data);
      updateLinkedInfoUI(ids);

      state.saveStickyDetails();
      return;
    }

    if (clearBtn) {
      setActiveBundleId(data, '');
      updateBundleInfoUI(data);

      const ids = normalizeLinkedSourceIds(data);
      const list = $('inputSourcesList');
      if (list) {
        const hasAny = ids.length > 0;
        list.style.opacity = hasAny ? '1' : '0.55';
        list.style.pointerEvents = hasAny ? 'auto' : 'none';
        list.style.filter = hasAny ? 'none' : 'grayscale(0.2)';
      }

      syncOutputCheckboxesFromLinkedSources(data);
      updateLinkedInfoUI(ids);

      state.saveStickyDetails();
      return;
    }

    if (!applyBtn) return;

    const pick = content.querySelector('#bundlePick');
    const nameEl = content.querySelector('#bundleName');
    const name = String(nameEl?.value || '').trim();

    if (!name) {
      alert('Vul een pakketnaam in.');
      return;
    }

    const ids = Array.from(content.querySelectorAll('.input-source-cb'))
      .filter((cb) => cb.checked)
      .map((cb) => cb.value);

    if (!ids.length) {
      alert('Selecteer minimaal 1 output om in het pakket te zetten.');
      return;
    }

    const bundles = ensureOutputBundlesArray();
    const pickedId = String(pick?.value || '').trim();

    let bundle = pickedId ? findBundleById(pickedId) : null;

    if (!bundle) {
      bundle = { id: makeLocalId('bundle'), name, outputUids: [] };
      bundles.push(bundle);

      if (pick) {
        const opt = document.createElement('option');
        opt.value = bundle.id;
        opt.textContent = bundle.name;
        pick.appendChild(opt);
        pick.value = bundle.id;
      }
    } else {
      bundle.name = name;
      if (pick) {
        const opt = pick.querySelector(`option[value="${bundle.id}"]`);
        if (opt) opt.textContent = name;
      }
    }

    bundle.outputUids = [...new Set(ids.map((x) => String(x ?? '').trim()).filter((x) => x))];

    setActiveBundleId(data, bundle.id);

    setLinkedSourceIds(data, []);
    data.text = '';

    const ext = $('inputExternalToggle');
    if (ext) ext.checked = false;

    updateLinkedInfoUI([]);
    updateBundleInfoUI(data);
    syncOutputCheckboxesFromBundle(bundle);

    state.saveStickyDetails();
  });
};

/* =========================================================
   OPEN / SAVE
   ========================================================= */

export function openEditModal(colIdx, slotIdx) {
  editingSticky = { colIdx, slotIdx };

  if (!areListenersAttached) {
    setupPermanentListeners();
    areListenersAttached = true;
  }

  renderContent();

  const modal = $('editModal');
  if (modal) modal.style.display = 'grid';

  if (slotIdx === 2) {
    const data = getStickyData();
    const ids = normalizeLinkedSourceIds(data);
    updateLinkedInfoUI(ids);
    updateBundleInfoUI(data);
  }
}

export function saveModalDetails(closeModal = true) {
  const data = getStickyData();
  if (!data) return;

  const slotIdx = editingSticky.slotIdx;
  const content = $('modalContent');
  if (!content) return;

  if (slotIdx === 3) {
    const statusVal = $('processStatus')?.value ?? '';
    data.processStatus = statusVal === '' ? null : statusVal;

    const expVal = $('workExp')?.value ?? '';
    data.workExp = expVal === '' ? null : expVal;

    const expNote = $('workExpNote');
    data.workExpNote = expNote ? expNote.value : '';

    const typeVal = content.querySelector('input[name="metaType"]')?.value ?? '';
    data.type = typeVal === '' ? null : typeVal;

    const procVal = content.querySelector('input[name="metaValue"]')?.value ?? '';
    data.processValue = procVal === '' ? null : procVal;

    const success = $('successFactors');
    data.successFactors = success ? success.value : '';

    data.causes = Array.from(content.querySelectorAll('.input-cause'))
      .map((i) => i.value)
      .filter((v) => v.trim());

    data.improvements = Array.from(content.querySelectorAll('.input-measure'))
      .map((i) => i.value)
      .filter((v) => v.trim());

    const disruptTbody = $('disruptTbody');
    if (disruptTbody) {
      const rows = disruptTbody.querySelectorAll('tr');
      data.disruptions = Array.from(rows)
        .map((tr) => ({
          scenario: tr.querySelector('td:nth-child(1) input')?.value || '',
          frequency: tr.querySelector('td:nth-child(2) input[type="hidden"]')?.value || null,
          workaround: tr.querySelector('td:nth-child(3) input')?.value || ''
        }))
        .filter((d) => d.scenario.trim());
    }

  } else if (slotIdx === 1) {
    ensureSystemDataShape(data);
    const sd = data.systemData;

    sd.noSystem = !!content.querySelector('#sysNoSystem')?.checked;

    if (sd.noSystem) {
      sd.isMulti = false;
      sd.systems = [
        {
          id: `${Date.now()}_${Math.random().toString(16).slice(2)}`,
          name: '',
          isLegacy: false,
          futureSystem: '',
          qa: {},
          calculatedScore: null
        }
      ];
      sd.systemName = '';
      sd.calculatedScore = null;
    }

    sd.isMulti = sd.noSystem ? false : !!content.querySelector('#sysMultiEnable')?.checked;

    if (!sd.noSystem) {
      const cards = content.querySelectorAll('.system-card');
      const nextSystems = [];

      cards.forEach((card) => {
        const sysId = card.dataset.sysId;
        const name = card.querySelector('.sys-name')?.value || '';
        const isLegacy = !!card.querySelector('.sys-legacy')?.checked;
        const futureSystem = card.querySelector('.sys-future')?.value || '';

        const qa = {};
        SYSTEM_QUESTIONS.forEach((q) => {
          const sel = `input[name="sys_${CSS.escape(sysId)}_${CSS.escape(q.id)}"]`;
          const vStr = contentEl.querySelector(sel)?.value ?? '';
          if (vStr === '') {
            qa[q.id] = null;
            return;
          }
          const n = parseInt(vStr, 10);
          qa[q.id] = Number.isFinite(n) ? n : null;
        });

        const score = computeSystemScoreFromAnswers(
          Object.fromEntries(Object.entries(qa).filter(([, v]) => Number.isFinite(Number(v))))
        );

        nextSystems.push({
          id: String(sysId),
          name: String(name),
          isLegacy,
          futureSystem: String(futureSystem),
          qa,
          calculatedScore: score
        });
      });

      sd.systems = nextSystems.length ? nextSystems : sd.systems;

      sd.systemName = sd.systems?.[0]?.name || '';

      const scores = sd.systems.map((s) => s.calculatedScore).filter((x) => x != null);
      sd.calculatedScore =
        scores.length === 0 ? null : Math.round(scores.reduce((a, b) => a + b, 0) / scores.length);
    }

    // ===== SYSTEM MERGE APPLY =====
    const enable = $('sysMergeEnable')?.checked ?? false;
    const startSel = $('sysMergeStartSelect');
    const endSel = $('sysMergeEndSelect');

    const colIdx = editingSticky.colIdx;

    if (!enable) {
      state.setSystemMergeRangeForCol(colIdx, colIdx, colIdx);
    } else {
      const startCol = startSel ? parseInt(startSel.value, 10) : colIdx;
      const endCol = endSel ? parseInt(endSel.value, 10) : colIdx;

      const s = Math.min(startCol, endCol);
      const e = Math.max(startCol, endCol);

      if (s > colIdx || e < colIdx || s === e) {
        state.setSystemMergeRangeForCol(colIdx, colIdx, colIdx);
      } else {
        state.setSystemMergeRangeForCol(colIdx, s, e);

        const sheet = state.activeSheet;
        const source = deepClone(sheet.columns[colIdx].slots[1]);

        state.beginBatch({ reason: 'batch' });
        for (let c = s; c <= e; c++) {
          if (sheet.columns[c]?.isVisible === false) continue;
          sheet.columns[c].slots[1] = deepClone(source);
        }
        state.endBatch({ reason: 'columns' });
      }
    }
  } else if (slotIdx === 2 || slotIdx === 4) {
    if (slotIdx === 2) {
      const activeBundleId = getActiveBundleId(data);

      if (!activeBundleId) {
        const ext = $('inputExternalToggle');
        if (ext && ext.checked) {
          setLinkedSourceIds(data, []);
        } else {
          const ids = Array.from(content.querySelectorAll('.input-source-cb'))
            .filter((cb) => cb.checked)
            .map((cb) => cb.value);
          setLinkedSourceIds(data, ids);
        }
      } else {
        const b = findBundleById(activeBundleId);
        const nm = String(content.querySelector('#bundleName')?.value || '').trim();
        if (b && nm) b.name = nm;

        data.text = '';
        setLinkedSourceIds(data, []);
      }
    }

    const defTbody = $('defTbody');
    if (defTbody) {
      const rows = defTbody.querySelectorAll('tr');
      data.inputDefinitions = Array.from(rows)
        .map((tr) => ({
          item: tr.querySelector('td:nth-child(1) input')?.value || '',
          specifications: tr.querySelector('td:nth-child(2) textarea')?.value || '',
          type: tr.querySelector('td:nth-child(3) input[type="hidden"]')?.value || null
        }))
        .filter(
          (d) =>
            d.item.trim() ||
            d.specifications.trim() ||
            (d.type != null && String(d.type).trim() !== '')
        );
    }

    const ioQual = $('ioTabQual');
    if (ioQual) {
      data.qa = data.qa || {};
      IO_CRITERIA.forEach((c) => {
        const gateInput = content.querySelector(`input[name="qa_gate_${CSS.escape(c.key)}"]`);
        const impactInput = content.querySelector(`input[name="qa_impact_${CSS.escape(c.key)}"]`);
        const noteInput = $(`note_${c.key}`);

        let finalResult = null;
        let finalImpact = null;

        if (gateInput) {
          finalResult = gateInput.value || null;
        }
        if (impactInput) {
          finalImpact = impactInput.value || null;
        }

        // If 'GOOD' is selected, reset impact
        if (finalResult === 'GOOD') finalImpact = null;

        data.qa[c.key] = {
          result: finalResult,
          impact: finalImpact,
          note: noteInput ? noteInput.value : ''
        };
      });
    }

    // ===== OUTPUT MERGE APPLY =====
    if (slotIdx === 4) {
      const enable = $('mergeEnable')?.checked ?? false;
      const startSel = $('mergeStartSelect');
      const endSel = $('mergeEndSelect');

      const colIdx = editingSticky.colIdx;

      if (!enable) {
        state.setOutputMergeRangeForCol(colIdx, colIdx, colIdx);
      } else {
        const startCol = startSel ? parseInt(startSel.value, 10) : colIdx;
        const endCol = endSel ? parseInt(endSel.value, 10) : colIdx;

        const s = Math.min(startCol, endCol);
        const e = Math.max(startCol, endCol);

        if (s > colIdx || e < colIdx || s === e) {
          state.setOutputMergeRangeForCol(colIdx, colIdx, colIdx);
        } else {
          state.setOutputMergeRangeForCol(colIdx, s, e);

          const sheet = state.activeSheet;
          const source = deepClone(sheet.columns[colIdx].slots[4]);

          state.beginBatch({ reason: 'batch' });
          for (let c = s; c <= e; c++) {
            if (sheet.columns[c]?.isVisible === false) continue;
            sheet.columns[c].slots[4] = deepClone(source);
          }
          state.endBatch({ reason: 'columns' });
        }
      }
    }
  }

  state.saveStickyDetails();

  if (closeModal) {
    const modal = $('editModal');
    if (modal) modal.style.display = 'none';
    hideTooltip();
  }
}

// === LOGICA MODAL (BLIKSEM) - MET SKIP OPTIE ===
export function openLogicModal(colIdx) {
  const sheet = state.activeSheet;
  if (!sheet) return;

  const col = sheet.columns[colIdx];
  const logic = col.logic || { condition: '', ifTrue: null, ifFalse: null };

  const createStepOptions = (selectedVal) => {
    let html = `<option value="SKIP" ${selectedVal === 'SKIP' ? 'selected' : ''}>⏭️ Sla deze stap over (SKIP)</option>`;

    html += sheet.columns
      .map((c, i) => {
        if (c.isVisible === false) return '';

        const label = c.slots?.[3]?.text || `Stap ${i + 1}`;
        const isSelected =
          selectedVal !== null && String(selectedVal) === String(i) ? 'selected' : '';

        if (i === colIdx) return '';

        return `<option value="${i}" ${isSelected}>Ga naar: ${i + 1}. ${escapeAttr(label)}</option>`;
      })
      .join('');

    return html;
  };

  const html = `
    <h3>⚡ Conditionele Stap</h3>
    <div class="sub-text">
        Bepaal wanneer deze stap uitgevoerd moet worden of overgeslagen.<br>
        <em>Huidige stap: <strong>${escapeAttr(col.slots?.[3]?.text || 'Naamloos')}</strong></em>
    </div>

    <div style="margin-top:16px;">
      <label class="modal-label">De Conditie (Vraag)</label>
      <textarea id="logicCondition" class="modal-input" rows="2" placeholder="Bijv: Komt patiënt voor in systeem?">${escapeAttr(
        logic.condition
      )}</textarea>
    </div>

    <div style="margin-top:20px; display:grid; grid-template-columns: 1fr 1fr; gap:20px;">
       <div style="background:rgba(0,255,0,0.05); padding:10px; border-radius:8px; border:1px solid rgba(0,255,0,0.1);">
         <label class="modal-label" style="color:#69f0ae; margin-top:0;">Indien JA (Waar)</label>
         <div class="io-helper" style="margin-bottom:8px; font-size:12px;">Als het antwoord JA is...</div>
         <select id="logicIfTrue" class="modal-input">
            <option value="">⬇️ Voer deze stap uit (Standaard)</option>
            ${createStepOptions(logic.ifTrue)}
         </select>
       </div>

       <div style="background:rgba(255,0,0,0.05); padding:10px; border-radius:8px; border:1px solid rgba(255,0,0,0.1);">
         <label class="modal-label" style="color:#ff8a80; margin-top:0;">Indien NEE (Niet Waar)</label>
         <div class="io-helper" style="margin-bottom:8px; font-size:12px;">Als het antwoord NEE is...</div>
         <select id="logicIfFalse" class="modal-input">
            <option value="">⬇️ Voer deze stap uit (Standaard)</option>
            ${createStepOptions(logic.ifFalse)}
         </select>
       </div>
    </div>
    
    <div style="margin-top:12px; font-size:12px; opacity:0.6; font-style:italic;">
        Tip: Gebruik <strong>SKIP</strong> om de stap over te slaan en direct door te gaan naar de volgende, zonder harde verwijzing.
    </div>

    <div class="modal-btns">
      <button id="logicRemoveBtn" class="std-btn danger-text" type="button">Logica wissen</button>
      <button id="logicCancelBtn" class="std-btn" type="button">Annuleren</button>
      <button id="logicSaveBtn" class="std-btn primary" type="button">Opslaan</button>
    </div>
  `;

  let overlay = document.getElementById('logicModalOverlay');
  if (overlay) overlay.remove();

  overlay = document.createElement('div');
  overlay.id = 'logicModalOverlay';
  overlay.className = 'modal-overlay';
  overlay.style.display = 'grid';

  const modal = document.createElement('div');
  modal.className = 'modal';
  modal.innerHTML = html;

  overlay.appendChild(modal);
  document.body.appendChild(overlay);

  const close = () => overlay.remove();

  overlay.addEventListener('click', (e) => {
    if (e.target === overlay) close();
  });
  modal.querySelector('#logicCancelBtn').onclick = close;

  modal.querySelector('#logicRemoveBtn').onclick = () => {
    state.toggleConditional(colIdx);
    close();
  };

  modal.querySelector('#logicSaveBtn').onclick = () => {
    const condition = document.getElementById('logicCondition').value;
    const ifTrue = document.getElementById('logicIfTrue').value;
    const ifFalse = document.getElementById('logicIfFalse').value;

    state.setColumnLogic(colIdx, { condition, ifTrue, ifFalse });
    close();
  };
}

// === GROEP MODAL (PUZZELSTUK) - MET VERWIJDER KNOP ===
export function openGroupModal(colIdx) {
  const sheet = state.activeSheet;
  if (!sheet) return;

  const existingGroup = state.getGroupForCol(colIdx);
  const titleVal = existingGroup ? existingGroup.title : '';

  let currentCols =
    existingGroup && Array.isArray(existingGroup.cols) ? [...existingGroup.cols] : [colIdx];

  const renderColsList = () => {
    const listEl = document.getElementById('groupColList');
    if (!listEl) return;

    listEl.innerHTML = '';

    if (currentCols.length === 0) {
      listEl.innerHTML =
        '<div style="font-size:12px; opacity:0.6; padding:8px;">Geen kolommen geselecteerd.</div>';
      return;
    }

    currentCols.sort((a, b) => a - b);

    currentCols.forEach((cIdx) => {
      const col = sheet.columns[cIdx];
      if (!col || col.isVisible === false) return;

      const label = col.slots?.[3]?.text || `Kolom ${cIdx + 1}`;

      const item = document.createElement('div');
      item.className = 'col-manager-item';
      item.innerHTML = `
             <span style="font-size:13px;"><strong>${cIdx + 1}.</strong> ${escapeAttr(label)}</span>
             <button class="btn-icon danger" type="button" style="width:24px; height:24px; min-width:24px;">×</button>
          `;

      item.querySelector('button').onclick = () => {
        currentCols = currentCols.filter((x) => x !== cIdx);
        renderColsList();
        renderAddOptions();
      };

      listEl.appendChild(item);
    });
  };

  const renderAddOptions = () => {
    const select = document.getElementById('groupAddSelect');
    if (!select) return;

    select.innerHTML = '<option value="">-- Kies een kolom --</option>';

    sheet.columns.forEach((c, i) => {
      if (c.isVisible === false) return;
      if (currentCols.includes(i)) return;

      const label = c.slots?.[3]?.text || `Kolom ${i + 1}`;
      const opt = document.createElement('option');
      opt.value = i;
      opt.textContent = `${i + 1}. ${label}`;
      select.appendChild(opt);
    });
  };

  const html = `
    <h3>🧩 Groep Definitie</h3>
    <div class="sub-text">Beheer de kolommen in deze groep.</div>

    <div style="margin-top:16px;">
      <label class="modal-label">Naam van de groep</label>
      <input id="groupTitle" class="modal-input" type="text" placeholder="Bijv. Triage Fase" value="${escapeAttr(
        titleVal
      )}">
    </div>

    <div style="margin-top:24px;">
       <label class="modal-label">Kolommen in deze groep</label>
       <div id="groupColList" style="border:1px solid rgba(255,255,255,0.1); border-radius:8px; max-height:200px; overflow-y:auto; margin-bottom:12px;"></div>
       
       <div style="display:flex; gap:8px;">
          <select id="groupAddSelect" class="modal-input" style="flex:1;"></select>
          <button id="groupAddBtn" class="std-btn" type="button">Toevoegen</button>
       </div>
    </div>

    <div class="modal-btns">
      ${
        existingGroup
          ? `<button id="groupRemoveBtn" class="std-btn danger-text" type="button">Groep opheffen</button>`
          : '<div></div>'
      }
      <button id="groupCancelBtn" class="std-btn" type="button">Annuleren</button>
      <button id="groupSaveBtn" class="std-btn primary" type="button">Opslaan</button>
    </div>
  `;

  let overlay = document.getElementById('groupModalOverlay');
  if (overlay) overlay.remove();

  overlay = document.createElement('div');
  overlay.id = 'groupModalOverlay';
  overlay.className = 'modal-overlay';
  overlay.style.display = 'grid';

  const modal = document.createElement('div');
  modal.className = 'modal';
  modal.innerHTML = html;

  overlay.appendChild(modal);
  document.body.appendChild(overlay);

  renderColsList();
  renderAddOptions();

  const addBtn = document.getElementById('groupAddBtn');
  const addSelect = document.getElementById('groupAddSelect');

  addBtn.onclick = () => {
    const val = addSelect.value;
    if (!val) return;
    const idx = parseInt(val, 10);
    if (!currentCols.includes(idx)) {
      currentCols.push(idx);
      renderColsList();
      renderAddOptions();
    }
  };

  const close = () => overlay.remove();

  overlay.addEventListener('click', (e) => {
    if (e.target === overlay) close();
  });
  modal.querySelector('#groupCancelBtn').onclick = close;

  const removeBtn = modal.querySelector('#groupRemoveBtn');
  if (removeBtn) {
    removeBtn.onclick = () => {
      if (confirm('Weet je zeker dat je deze groep wilt opheffen? De kolommen blijven gewoon bestaan.')) {
        state.removeGroup(existingGroup.id);
        close();
      }
    };
  }

  modal.querySelector('#groupSaveBtn').onclick = () => {
    const title = document.getElementById('groupTitle').value;

    if (currentCols.length === 0) {
      if (existingGroup) state.removeGroup(existingGroup.id);
    } else {
      state.setColumnGroup({
        id: existingGroup ? existingGroup.id : null,
        cols: currentCols,
        title
      });
    }
    close();
  };
}

// === VARIANT MODAL (ROUTES) - VERBETERDE VERSIE (MULTI PARENT) ===
export function openVariantModal(colIdx) {
  const sheet = state.activeSheet;
  if (!sheet) return;

  const col = sheet.columns[colIdx];

  const info = state.getVariantGroupForCol(colIdx);
  const existingGroup = info ? info.group : null;
  const isVariant = col.isVariant || !!existingGroup;

  let currentParents = [];
  if (existingGroup) {
    if (Array.isArray(existingGroup.parents)) currentParents = [...existingGroup.parents];
    else if (existingGroup.parentColIdx !== undefined) currentParents = [existingGroup.parentColIdx];
  } else {
    currentParents = [colIdx];
  }

  let currentVariants = existingGroup ? [...existingGroup.variants] : [];

  const getLetter = (i) => _toLetter(i);

  const renderParentList = () => {
    const container = document.getElementById('variantParentList');
    if (!container) return;
    container.innerHTML = '';

    (state.project.sheets || []).forEach((s) => {
      const isCurrentSheet = s.id === sheet.id;

      const header = document.createElement('div');
      header.style.fontSize = '11px';
      header.style.fontWeight = 'bold';
      header.style.marginTop = '12px';
      header.style.marginBottom = '4px';
      header.style.color = isCurrentSheet ? 'var(--ui-accent)' : '#aaa';
      header.style.textTransform = 'uppercase';
      header.textContent = s.name;
      container.appendChild(header);

      (s.columns || []).forEach((c, i) => {
        if (c.isVisible === false) return;

        if (isCurrentSheet && currentVariants.includes(i)) return;

        const val = isCurrentSheet ? i : `${s.id}::${i}`;
        const valStr = String(val);

        const isSelected = currentParents.some((p) => String(p) === valStr);
        const isCurrentCol = isCurrentSheet && i === colIdx;

        const div = document.createElement('label');
        div.className = 'sys-opt';
        div.style.display = 'flex';
        div.style.alignItems = 'center';
        div.style.gap = '10px';
        div.style.cursor = 'pointer';
        div.style.padding = '6px 10px';
        if (isSelected) div.classList.add('selected');

        const labelText = c.slots?.[3]?.text || `Kolom ${i + 1}`;

        div.innerHTML = `
                 <input type="checkbox" value="${escapeAttr(val)}" ${isSelected ? 'checked' : ''} style="accent-color:var(--ui-accent);">
                 <span style="flex:1; font-size:13px; color:${isCurrentSheet ? '#fff' : '#ccc'};">
                   ${i + 1}. ${escapeAttr(labelText)} ${isCurrentCol ? '<strong>(Dit)</strong>' : ''}
                 </span>
              `;

        div.querySelector('input').onchange = (e) => {
          if (e.target.checked) {
            const storeVal = isCurrentSheet ? i : val;
            currentParents.push(storeVal);
            div.classList.add('selected');
          } else {
            currentParents = currentParents.filter((p) => String(p) !== valStr);
            div.classList.remove('selected');
          }
          renderAddOptions();
        };

        container.appendChild(div);
      });
    });
  };

  const renderVariantList = () => {
    const listEl = document.getElementById('variantList');
    if (!listEl) return;
    listEl.innerHTML = '';

    if (currentVariants.length === 0) {
      listEl.innerHTML =
        '<div style="font-size:12px; opacity:0.6; padding:8px;">Nog geen routes toegevoegd.</div>';
      return;
    }

    currentVariants.forEach((vIdx, i) => {
      const c = sheet.columns[vIdx];
      if (!c || c.isVisible === false) return;

      const label = c.slots?.[3]?.text || `Kolom ${vIdx + 1}`;
      const letter = getLetter(i);

      const item = document.createElement('div');
      item.className = 'col-manager-item';
      item.style.display = 'grid';
      item.style.gridTemplateColumns = '30px 1fr 30px';
      item.style.alignItems = 'center';

      item.innerHTML = `
           <div style="font-weight:bold; color:var(--ui-accent);">${escapeAttr(letter)}</div>
           <div style="font-size:13px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
             ${vIdx + 1}. ${escapeAttr(label)}
           </div>
           <button class="btn-icon danger" type="button" style="width:24px; height:24px;">×</button>
        `;

      item.querySelector('button').onclick = () => {
        currentVariants = currentVariants.filter((x) => x !== vIdx);
        renderVariantList();
        renderAddOptions();
        renderParentList();
      };

      listEl.appendChild(item);
    });
  };

  const renderAddOptions = () => {
    const select = document.getElementById('variantAddSelect');
    if (!select) return;
    select.innerHTML = '<option value="">-- Kies een route stap --</option>';

    sheet.columns.forEach((c, i) => {
      if (c.isVisible === false) return;
      if (currentParents.includes(i)) return;
      if (currentVariants.includes(i)) return;

      const label = c.slots?.[3]?.text || `Kolom ${i + 1}`;
      const opt = document.createElement('option');
      opt.value = i;
      opt.textContent = `${i + 1}. ${label}`;
      select.appendChild(opt);
    });
  };

  const html = `
    <h3>🔀 Route / Variant Definitie</h3>
    <div class="sub-text">Selecteer hoofdprocessen (ook uit andere sheets) en lokale sub-routes.</div>

    <div style="margin-top:16px;">
      <label class="modal-label">1. Hoofdprocessen</label>
      <div id="variantParentList" style="border:1px solid rgba(255,255,255,0.1); border-radius:8px; max-height:200px; overflow-y:auto; padding:4px 8px;"></div>
    </div>

    <div style="margin-top:20px;">
       <label class="modal-label">2. Lokale Routes</label>
       <div id="variantList" style="border:1px solid rgba(255,255,255,0.1); border-radius:8px; max-height:150px; overflow-y:auto; margin-bottom:12px;"></div>
       <div style="display:flex; gap:8px;">
          <select id="variantAddSelect" class="modal-input" style="flex:1;"></select>
          <button id="variantAddBtn" class="std-btn" type="button">Toevoegen</button>
       </div>
    </div>

    <div class="modal-btns">
      ${isVariant ? `<button id="variantRemoveBtn" class="std-btn danger-text" type="button">Split opheffen</button>` : '<div></div>'}
      <button id="variantCancelBtn" class="std-btn" type="button">Annuleren</button>
      <button id="variantSaveBtn" class="std-btn primary" type="button">Opslaan</button>
    </div>
  `;

  let overlay = document.getElementById('variantModalOverlay');
  if (overlay) overlay.remove();

  overlay = document.createElement('div');
  overlay.id = 'variantModalOverlay';
  overlay.className = 'modal-overlay';
  overlay.style.display = 'grid';

  const modal = document.createElement('div');
  modal.className = 'modal';
  modal.innerHTML = html;

  overlay.appendChild(modal);
  document.body.appendChild(overlay);

  renderParentList();
  renderVariantList();
  renderAddOptions();

  document.getElementById('variantAddBtn').onclick = () => {
    const val = document.getElementById('variantAddSelect').value;
    if (!val) return;
    currentVariants.push(parseInt(val, 10));
    renderVariantList();
    renderAddOptions();
    renderParentList();
  };

  const close = () => overlay.remove();
  overlay.addEventListener('click', (e) => {
    if (e.target === overlay) close();
  });
  modal.querySelector('#variantCancelBtn').onclick = close;

  const rmBtn = modal.querySelector('#variantRemoveBtn');
  if (rmBtn) {
    rmBtn.onclick = () => {
      if (confirm('Weet je zeker dat je de split wilt opheffen?')) {
        if (existingGroup) {
          state.removeVariantGroup(existingGroup.id);
        } else {
          state.toggleVariant(colIdx);
        }
        close();
      }
    };
  }

  modal.querySelector('#variantSaveBtn').onclick = () => {
    if (currentParents.length === 0) {
      alert('Selecteer minimaal één hoofdproces.');
      return;
    }
    if (currentVariants.length === 0) {
      if (existingGroup) state.removeVariantGroup(existingGroup.id);
    } else {
      state.setVariantGroup({
        parents: currentParents,
        variants: currentVariants
      });
    }
    close();
  };
}