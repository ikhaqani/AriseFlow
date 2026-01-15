export const STORAGE_KEY = 'pro_lss_sipoc_v2_ultimate';

const deepFreeze = (obj) => {
  Object.keys(obj).forEach((prop) => {
    const val = obj[prop];
    if (typeof val === 'object' && val !== null) deepFreeze(val);
  });
  return Object.freeze(obj);
};

export const APP_CONFIG = deepFreeze({
  SLOT_COUNT: 6,
  MAX_SCORE_SYSTEM: 100,
  VERSION: '2.7',
  LOCALE: 'nl-NL'
});

export const DEFAULTS = deepFreeze({
  PROJECT_TITLE: 'Nieuw Proces Project',
  SHEET_NAME: 'Proces Flow 1',
  AUTHOR: 'Anoniem'
});

export const SYSTEM_QUESTIONS = deepFreeze([
  {
    id: 'workarounds',
    label: '1. Hoe vaak dwingt het systeem je tot workarounds?',
    options: ['(Bijna) nooit', 'Soms', 'Vaak', '(Bijna) altijd']
  },
  {
    id: 'performance',
    label: '2. Hoe vaak remt het systeem je af?',
    options: ['(Bijna) nooit', 'Soms', 'Vaak', '(Bijna) altijd']
  },
  {
    id: 'double',
    label: '3. Hoe vaak moet je gegevens dubbel registreren?',
    options: ['(Bijna) nooit', 'Soms', 'Vaak', '(Bijna) altijd']
  },
  {
    id: 'error',
    label: '4. Hoe vaak laat het systeem ruimte voor fouten?',
    options: ['(Bijna) nooit', 'Soms', 'Vaak', '(Bijna) altijd']
  },
  {
    id: 'depend',
    label: '5. Wat is de impact bij systeemuitval?',
    options: ['Veilig (Fallback)', 'Vertraging', 'Groot Risico', 'Volledige Stilstand']
  }
]);

export const IO_CRITERIA = deepFreeze([
  { key: 'compleet', label: 'Compleetheid', weight: 5, meet: 'Alle benodigde data/materialen zijn aanwezig.' },
  { key: 'kwaliteit', label: 'Datakwaliteit', weight: 5, meet: 'Formaat, resolutie en inhoud zijn correct.' },
  { key: 'duidelijkheid', label: 'Eenduidigheid', weight: 3, meet: 'Geen interpretatie of vragen nodig om te starten.' },
  { key: 'tijdigheid', label: 'Tijdigheid', weight: 3, meet: 'Beschikbaar op het geplande moment.' },
  { key: 'standaard', label: 'Standaardisatie', weight: 1, meet: 'Conform naamgeving en protocollen.' },
  { key: 'overdracht', label: 'Overdracht', weight: 1, meet: 'Status correct bijgewerkt in bronsystemen.' }
]);

export const ACTIVITY_TYPES = deepFreeze([
  { value: 'Taak', label: '📝 Taak' },
  { value: 'Afspraak', label: '📅 Afspraak' },
  { value: 'TimeOut', label: '🛑 Time Out' },
  { value: 'Beoordeling', label: '🔎 Beoordeling' }
]);

export const LEAN_VALUES = deepFreeze([
  { value: 'VA', label: 'VA - Patiëntwaarde' },
  { value: 'BNVA', label: 'BNVA - Business Noodzaak' },
  { value: 'NVA', label: 'NVA - Verspilling' }
]);

export const PROCESS_STATUSES = deepFreeze([
  { value: 'SAD', label: 'Niet in control', emoji: '☹️', class: 'selected-sad' },
  { value: 'NEUTRAL', label: 'Aandachtspunt', emoji: '😐', class: 'selected-neu' },
  { value: 'HAPPY', label: 'In control', emoji: '🙂', class: 'selected-hap' }
]);

export const DISRUPTION_FREQUENCIES = deepFreeze(['Zelden', 'Soms', 'Vaak', 'Altijd']);

export const DEFINITION_TYPES = deepFreeze([
  { value: 'HARD', label: 'Hard' },
  { value: 'SOFT', label: 'Soft' }
]);

export const uid = () => {
  if (typeof crypto !== 'undefined' && crypto.randomUUID) return crypto.randomUUID();
  return `${Date.now().toString(36)}${Math.random().toString(36).slice(2)}`;
};

const createInitialQA = () => {
  const qa = {};
  IO_CRITERIA.forEach((c) => {
    qa[c.key] = { result: '', note: '' };
  });
  return qa;
};

const createEmptySystemQa = () => {
  const qa = {};
  SYSTEM_QUESTIONS.forEach((q) => {
    qa[q.id] = null;
  });
  return qa;
};

const createInitialSystemData = () => ({
  isMulti: false,
  systemName: '',
  calculatedScore: null,
  systems: [
    {
      id: uid(),
      name: '',
      isLegacy: false,
      futureSystem: '',
      qa: createEmptySystemQa(),
      calculatedScore: null
    }
  ]
});

export const createSticky = () => ({
  id: uid(),
  created: Date.now(),
  text: '',
  linkedSourceId: null,
  type: null,
  processValue: null,
  processStatus: null,
  successFactors: '',
  causes: [],
  improvements: [],
  qa: createInitialQA(),
  systemData: createInitialSystemData(),
  inputDefinitions: [],
  disruptions: [],
  workExp: null,
  workExpNote: ''
});

export const createColumn = (order = 1) => ({
  id: uid(),
  order: Number.isFinite(Number(order)) ? Number(order) : 1,
  isVisible: true,
  isParallel: false,
  isVariant: false,
  hasTransition: false,
  transitionNext: '',
  outputId: null,
  slots: Array.from({ length: APP_CONFIG.SLOT_COUNT }, () => createSticky())
});

export const createSheet = (name = DEFAULTS.SHEET_NAME, order = 1) => ({
  id: uid(),
  name,
  order: Number.isFinite(Number(order)) ? Number(order) : 1,
  columns: [createColumn(1)],
  outputMerges: [],
  systemMerges: []
});

export const createProjectState = () => {
  const firstSheet = createSheet(DEFAULTS.SHEET_NAME, 1);
  return {
    id: uid(),
    projectTitle: DEFAULTS.PROJECT_TITLE,
    author: DEFAULTS.AUTHOR,
    created: new Date().toISOString(),
    version: APP_CONFIG.VERSION,
    activeSheetId: firstSheet.id,
    sheets: [firstSheet]
  };
};