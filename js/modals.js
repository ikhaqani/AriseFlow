// modals.js  (VOLLEDIG)
// -----------------------------------------------------------
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
    impactB:
      'B. Extra werk: Bruikbaar na correctie/omweg (converteren, herexport, opnieuw opvragen).',
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

const escapeAttr = (v) =>
  String(v ?? '')
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');

const createRadioGroup = (name, options, selectedValue, isHorizontal = false) => `
  <div class="radio-group-container ${isHorizontal ? 'horizontal' : 'vertical'}">
    ${options
      .map((opt) => {
        const val = opt.value ?? opt;
        const label = opt.label ?? opt;
        const isSelected =
          selectedValue !== null &&
          selectedValue !== undefined &&
          String(val) === String(selectedValue);

        return `
          <div class="sys-opt ${isSelected ? 'selected' : ''}" data-value="${val}">
            ${label}
          </div>
        `;
      })
      .join('')}
    <input type="hidden" name="${name}" value="${
      selectedValue !== null && selectedValue !== undefined ? selectedValue : ''
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
      return `<option value="${idx}" ${idx === startDefault ? 'selected' : ''}>${escapeAttr(
        lbl
      )}</option>`;
    })
    .join('');

  const endOpts = vis
    .slice(pos)
    .map((idx) => {
      const lbl = `Kolom ${idx + 1}`;
      return `<option value="${idx}" ${idx === endDefault ? 'selected' : ''}>${escapeAttr(
        lbl
      )}</option>`;
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
      return `<option value="${idx}" ${idx === startDefault ? 'selected' : ''}>${escapeAttr(
        lbl
      )}</option>`;
    })
    .join('');

  const endOpts = vis
    .slice(pos)
    .map((idx) => {
      const lbl = `Kolom ${idx + 1}`;
      return `<option value="${idx}" ${idx === endDefault ? 'selected' : ''}>${escapeAttr(
        lbl
      )}</option>`;
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
          <select id="sysMergeStartSelect" class="modal-input" ${
            hasMerge ? '' : 'disabled'
          }>${startOpts}</select>
        </div>

        <div style="flex:1; min-width:180px;">
          <div style="font-size:12px; opacity:.85; margin-bottom:6px;">Tot (rechter grens)</div>
          <select id="sysMergeEndSelect" class="modal-input" ${
            hasMerge ? '' : 'disabled'
          }>${endOpts}</select>
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

  if (!Array.isArray(sd.systems)) {
    const legacyName = typeof sd.systemName === 'string' ? sd.systemName.trim() : '';
    sd.systems = legacyName
      ? [{ id: `${Date.now()}_${Math.random().toString(16).slice(2)}`, name: legacyName, isLegacy: false, futureSystem: '', qa: {}, calculatedScore: null }]
      : [{ id: `${Date.now()}_${Math.random().toString(16).slice(2)}`, name: '', isLegacy: false, futureSystem: '', qa: {}, calculatedScore: null }];
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
    sd.systems = [{ id: `${Date.now()}_${Math.random().toString(16).slice(2)}`, name: '', isLegacy: false, futureSystem: '', qa: {}, calculatedScore: null }];
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

  // assumes options indices 0..3 (max 3 each); keeps your previous scoring logic
  const maxPoints = SYSTEM_QUESTIONS.length * 3;
  const safeTotal = Math.min(total, maxPoints);
  return Math.round(100 * (1 - safeTotal / maxPoints));
}

// Helper: persist systems from DOM into data.systemData (from currently rendered cards)
function persistSystemTabFromDOM(contentEl, data) {
  if (!contentEl || !data) return;

  ensureSystemDataShape(data);
  const sd = data.systemData;

  // Read systems from the currently rendered cards
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

  // Backward compat fields
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

  let html = `
    <div id="systemWrapper">
      ${renderSystemMergeControls()}

      <div class="io-helper">
        Geef aan hoe goed het systeem het proces ondersteunt. Werk je in meerdere systemen? Voeg ze toe en beantwoord System Fit per systeem.
      </div>

      <label style="display:flex; align-items:center; gap:10px; font-size:13px; margin: 12px 0 14px 0; cursor:pointer;">
        <input id="sysMultiEnable" type="checkbox" ${isMulti ? 'checked' : ''} />
        Ik werk in meerdere systemen binnen deze processtap
      </label>

      <div id="sysList" style="display:grid; gap:14px;">
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
                  ${systems.length <= 1 ? 'disabled' : ''}
                  style="padding:6px 10px; font-size:12px;">
            Verwijderen
          </button>
        </div>

        <div style="display:grid; gap:10px;">
          <div style="display:grid; grid-template-columns: 1fr; gap:6px;">
            <div class="modal-label" style="margin:0; font-size:11px;">Systeemnaam</div>
            <input class="modal-input sys-name" type="text" value="${escapeAttr(name)}" placeholder="Bijv. ARIA / EPIC / Radiotherapieweb / Monaco..." />
          </div>

          <div style="display:grid; grid-template-columns: 1fr 1fr; gap:10px;">
            <label style="display:flex; align-items:center; gap:10px; font-size:13px; cursor:pointer; margin:0;">
              <input class="sys-legacy" type="checkbox" ${isLegacy ? 'checked' : ''} />
              Legacy systeem
            </label>

            <div style="display:grid; grid-template-columns: 1fr; gap:6px;">
              <div class="modal-label" style="margin:0; font-size:11px;">Toekomstig systeem (verwachting)</div>
              <input class="modal-input sys-future" type="text" value="${escapeAttr(future)}" placeholder="Bijv. ARIA / EPIC / nieuw portaal..." />
            </div>
          </div>
        </div>

        <div style="margin-top:12px; padding-top:12px; border-top:1px solid rgba(255,255,255,0.08);">
          <div class="modal-label" style="margin-top:0;">System Fit vragen</div>
          <div class="io-helper" style="margin-top:6px; margin-bottom:10px;">
            Beantwoord per vraag hoe goed dit systeem jouw taak ondersteunt.
          </div>

          ${SYSTEM_QUESTIONS.map((q) => {
            const currentVal = qa[q.id] !== undefined ? qa[q.id] : null;
            const optionsMapped = q.options.map((optText, optIdx) => ({ value: optIdx, label: optText }));

            return `
              <div class="system-question" style="margin-bottom:12px;">
                <div class="sys-q-title">${escapeAttr(q.label)}</div>
                ${createRadioGroup(`sys_${escapeAttr(sysId)}_${escapeAttr(q.id)}`, optionsMapped, currentVal, true)}
              </div>
            `;
          }).join('')}

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

      <button class="std-btn primary" id="btnAddSystem" data-action="add-system" type="button" style="margin-top:14px; display:${isMulti ? 'inline-flex' : 'none'};">
        + Systeem toevoegen
      </button>

      <div style="margin-top:14px; font-size:12px; opacity:.9;">
        Overall score (kolom): <strong id="sysOverallScore">${Number.isFinite(Number(sd.calculatedScore)) ? `${sd.calculatedScore}%` : '—'}</strong>
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
    const allOutputs = state.getAllOutputs();
    const options = Object.entries(allOutputs)
      .map(([id, text]) => {
        const t = (text || '').substring(0, 40);
        return `<option value="${escapeAttr(id)}" ${
          data.linkedSourceId === id ? 'selected' : ''
        }>${escapeAttr(id)}: ${escapeAttr(t)}${(text || '').length > 40 ? '...' : ''}</option>`;
      })
      .join('');

    linkHtml = `
      <div style="margin-bottom: 20px; padding: 12px; background: rgba(0,0,0,0.2); border-radius: 8px; border: 1px solid rgba(255,255,255,0.1);">
        <div class="modal-label" style="margin-top:0;">Input Bron (Koppel aan Output)</div>
        <select id="inputSourceSelect" class="modal-input">
          <option value="">-- Geen / Externe Input --</option>
          ${options}
        </select>
        <div id="linkedInfoText" style="display:${
          data.linkedSourceId ? 'block' : 'none'
        }; color:var(--ui-accent); font-size:11px; margin-top:8px; font-weight:600;">
          🔗 Gekoppeld. Tekst wordt automatisch bijgewerkt.
        </div>
      </div>
    `;
  }

  const mergeControls = !isInputRow ? renderOutputMergeControls() : '';

  const qaRows = IO_CRITERIA.map((c) => {
    const qa = data.qa?.[c.key] || {};
    const defs = IO_CRITERIA_DEFS[c.key] || {};
    const currentRes = qa.result;

    const showImpact =
      currentRes === 'FAIL' ||
      currentRes === 'MODERATE' ||
      currentRes === 'POOR' ||
      currentRes === 'MINOR';

    let initialValue = null;
    if (currentRes === 'GOOD') initialValue = 'GOOD';
    else if (currentRes === 'NA') initialValue = 'NA';
    else if (['FAIL', 'MODERATE', 'POOR', 'MINOR'].includes(currentRes)) initialValue = 'FAIL';

    let impactValue = null;
    if (currentRes === 'POOR') impactValue = 'A';
    if (currentRes === 'MODERATE') impactValue = 'B';
    if (currentRes === 'MINOR') impactValue = 'C';

    return `
      <div class="qa-item-wrapper" style="margin-bottom: 24px; padding-bottom: 24px; border-bottom: 1px solid rgba(255,255,255,0.05);">
        <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:12px;">
            <div style="max-width: 60%;">
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
                     { value: 'FAIL', label: 'Voldoet niet' },
                     { value: 'NA', label: 'N.V.T.' }
                   ],
                   initialValue,
                   true
                 )}
            </div>
        </div>

        <div id="impact_wrapper_${c.key}" style="display:${
      showImpact ? 'block' : 'none'
    }; background: rgba(255,82,82, 0.1); border-left: 3px solid #ff5252; padding: 12px; margin-top:8px; border-radius: 0 4px 4px 0;">
            <div style="font-size:11px; font-weight:bold; color:#ff5252; text-transform:uppercase; margin-bottom:8px;">Wat is de impact op jouw taak?</div>
            
            <div class="radio-group-container vertical">
                <div class="sys-opt impact-opt ${impactValue === 'A' ? 'selected' : ''}" 
                     data-value="A" 
                     data-key="${c.key}"
                     style="text-align:left; height:auto; padding:8px 12px; margin-bottom:6px;">
                    <div style="font-weight:bold; font-size:12px;">🔴 A. Blokkerend</div>
                    <div style="font-size:11px; opacity:0.8; font-weight:normal;">${escapeAttr(
                      defs.impactA
                    )}</div>
                </div>
                
                <div class="sys-opt impact-opt ${impactValue === 'B' ? 'selected' : ''}" 
                     data-value="B" 
                     data-key="${c.key}"
                     style="text-align:left; height:auto; padding:8px 12px; margin-bottom:6px;">
                    <div style="font-weight:bold; font-size:12px;">🟠 B. Extra werk</div>
                    <div style="font-size:11px; opacity:0.8; font-weight:normal;">${escapeAttr(
                      defs.impactB
                    )}</div>
                </div>

                <div class="sys-opt impact-opt ${impactValue === 'C' ? 'selected' : ''}" 
                     data-value="C" 
                     data-key="${c.key}"
                     style="text-align:left; height:auto; padding:8px 12px;">
                    <div style="font-weight:bold; font-size:12px;">🟡 C. Kleine frictie</div>
                    <div style="font-size:11px; opacity:0.8; font-weight:normal;">${escapeAttr(
                      defs.impactC
                    )}</div>
                </div>
            </div>
            <input type="hidden" name="qa_impact_${c.key}" value="${escapeAttr(impactValue || '')}">
        </div>

        <textarea id="note_${c.key}" class="io-note" style="margin-top:12px;" placeholder="Opmerking (optioneel)...">${escapeAttr(
          qa.note || ''
        )}</textarea>
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

  // merge enable toggle (Output)
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

  // merge enable toggle (System)
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

  // Multi-system toggle + add/remove system rows
  content.addEventListener('change', (e) => {
    if (e.target?.id !== 'sysMultiEnable') return;

    const data = getStickyData();
    if (!data) return;

    // Eerst huidige UI → data (anders raak je wijzigingen kwijt)
    persistSystemTabFromDOM(content, data);

    ensureSystemDataShape(data);
    data.systemData.isMulti = !!e.target.checked;

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

      // Bewaar eerst wat er nu in de UI staat (anders wordt je push direct overschreven)
      persistSystemTabFromDOM(content, data);

      // Voeg nieuw systeem toe
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

      // Bewaar eerst wat er nu in de UI staat
      persistSystemTabFromDOM(content, data);

      // Verwijder systeem
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
        if (impactWrapper) impactWrapper.style.display = 'none';
        const impactInput = content.querySelector(`input[name="qa_impact_${key}"]`);
        if (impactInput) impactInput.value = '';
        content
          .querySelectorAll(`.impact-opt[data-key="${key}"]`)
          .forEach((el) => el.classList.remove('selected'));
      }
      return;
    }

    opt.classList.add('selected');
    if (input) input.value = opt.dataset.value;

    if (input && input.name.startsWith('qa_gate_')) {
      const key = input.name.replace('qa_gate_', '');
      const impactWrapper = $(`impact_wrapper_${key}`);
      if (!impactWrapper) return;

      if (opt.dataset.value === 'FAIL') {
        impactWrapper.style.display = 'block';
      } else {
        impactWrapper.style.display = 'none';
        const impactInput = content.querySelector(`input[name="qa_impact_${key}"]`);
        if (impactInput) impactInput.value = '';
        content
          .querySelectorAll(`.impact-opt[data-key="${key}"]`)
          .forEach((el) => el.classList.remove('selected'));
      }
    }

    // If this is a system question radio (sys_<sysId>_<qId>), update live scores in UI
    if (input && input.name.startsWith('sys_')) {
      // Light live update: compute per-card score from current hidden inputs
      updateLiveSystemScoresInUI();
    }
  });

  function handleImpactClick(opt) {
    const key = String(opt?.dataset?.key || '').trim();
    const val = String(opt?.dataset?.value || '').trim();
    if (!key) return;

    const group = opt.closest('.radio-group-container') || opt.parentElement;
    const hidden = content.querySelector(`input[name="qa_impact_${key}"]`);

    const wasSelected = opt.classList.contains('selected');

    if (group) {
      group
        .querySelectorAll(`.impact-opt[data-key="${key}"]`)
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
    // Only works when system tab is rendered
    const wrapper = $('systemWrapper');
    if (!wrapper) return;

    const cards = wrapper.querySelectorAll('.system-card');
    let overallScores = [];

    cards.forEach((card) => {
      const sysId = card.dataset.sysId;
      const answers = {};

      SYSTEM_QUESTIONS.forEach((q) => {
        const v = wrapper.querySelector(`input[name="sys_${CSS.escape(sysId)}_${CSS.escape(q.id)}"]`)?.value;
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
      else overallEl.textContent = `${Math.round(overallScores.reduce((a, b) => a + b, 0) / overallScores.length)}%`;
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

  // Input linked source select
  content.addEventListener('change', (e) => {
    if (e.target?.id !== 'inputSourceSelect') return;

    const data = getStickyData();
    if (!data) return;

    data.linkedSourceId = e.target.value || null;

    const info = $('linkedInfoText');
    if (info) info.style.display = e.target.value ? 'block' : 'none';

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
    // ===== SYSTEM TAB SAVE (multi-system) =====
    ensureSystemDataShape(data);
    const sd = data.systemData;

    sd.isMulti = !!content.querySelector('#sysMultiEnable')?.checked;

    // Read all system cards from DOM
    const cards = content.querySelectorAll('.system-card');
    const nextSystems = [];

    cards.forEach((card) => {
      const sysId = card.dataset.sysId;
      const name = card.querySelector('.sys-name')?.value || '';
      const isLegacy = !!card.querySelector('.sys-legacy')?.checked;
      const futureSystem = card.querySelector('.sys-future')?.value || '';

      const qa = {};
      SYSTEM_QUESTIONS.forEach((q) => {
        const vStr = content.querySelector(`input[name="sys_${CSS.escape(sysId)}_${CSS.escape(q.id)}"]`)?.value ?? '';
        if (vStr === '') {
          qa[q.id] = null;
          return;
        }
        const n = parseInt(vStr, 10);
        qa[q.id] = Number.isFinite(n) ? n : null;
      });

      const score = computeSystemScoreFromAnswers(
        Object.fromEntries(
          Object.entries(qa).filter(([, v]) => Number.isFinite(Number(v)))
        )
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

    // Backward compat
    sd.systemName = sd.systems?.[0]?.name || '';

    // Overall score: average of non-null per-system scores
    const scores = sd.systems.map((s) => s.calculatedScore).filter((x) => x != null);
    sd.calculatedScore =
      scores.length === 0 ? null : Math.round(scores.reduce((a, b) => a + b, 0) / scores.length);

    // ===== SYSTEM MERGE APPLY (slotIdx === 1) =====
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

        // Propagate system slot (master) to all cols in range
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
      const select = $('inputSourceSelect');
      if (select) data.linkedSourceId = select.value || null;
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
        const gateInput = content.querySelector(`input[name="qa_gate_${c.key}"]`);
        const impactInput = content.querySelector(`input[name="qa_impact_${c.key}"]`);
        const noteInput = $(`note_${c.key}`);

        let finalResult = null;
        if (gateInput) {
          const gateVal = gateInput.value;
          if (gateVal === 'GOOD') finalResult = 'GOOD';
          else if (gateVal === 'NA') finalResult = 'NA';
          else if (gateVal === 'FAIL') {
            const impactVal = impactInput ? impactInput.value : '';
            if (impactVal === 'A') finalResult = 'POOR';
            else if (impactVal === 'B') finalResult = 'MODERATE';
            else if (impactVal === 'C') finalResult = 'MINOR';
            else finalResult = 'FAIL';
          }
        }

        data.qa[c.key] = {
          result: finalResult,
          note: noteInput ? noteInput.value : ''
        };
      });
    }

    // Output merge selection handling (slotIdx === 4)
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