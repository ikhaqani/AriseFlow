// config.js  (VOLLEDIG)
// -----------------------------------------------------------
export const STORAGE_KEY = "pro_lss_sipoc_v2_ultimate";

const deepFreeze = (obj) => {
  Object.keys(obj).forEach((prop) => {
    const val = obj[prop];
    if (typeof val === "object" && val !== null) deepFreeze(val);
  });
  return Object.freeze(obj);
};

export const APP_CONFIG = deepFreeze({
  SLOT_COUNT: 6,
  MAX_SCORE_SYSTEM: 100,
  VERSION: "2.6", // Update: Time Out & Beoordeling gesplitst
  LOCALE: "nl-NL"
});

export const DEFAULTS = deepFreeze({
  PROJECT_TITLE: "Nieuw Proces Project",
  SHEET_NAME: "Proces Flow 1",
  STICKY_TYPE: null,
  PROCESS_VALUE: null,
  PROCESS_STATUS: null,
  AUTHOR: "Anoniem"
});

export const SYSTEM_QUESTIONS = deepFreeze([
  { id: "workarounds", label: "1. Hoe vaak dwingt het systeem je tot workarounds?", options: ["(Bijna) nooit", "Soms", "Vaak", "(Bijna) altijd"] },
  { id: "performance", label: "2. Hoe vaak remt het systeem je af?", options: ["(Bijna) nooit", "Soms", "Vaak", "(Bijna) altijd"] },
  { id: "double", label: "3. Hoe vaak moet je gegevens dubbel registreren?", options: ["(Bijna) nooit", "Soms", "Vaak", "(Bijna) altijd"] },
  { id: "error", label: "4. Hoe vaak laat het systeem ruimte voor fouten?", options: ["(Bijna) nooit", "Soms", "Vaak", "(Bijna) altijd"] },
  { id: "depend", label: "5. Wat is de impact bij systeemuitval?", options: ["Veilig (Fallback)", "Vertraging", "Groot Risico", "Volledige Stilstand"] }
]);

export const IO_CRITERIA = deepFreeze([
  { key: "compleet", label: "Compleetheid", weight: 5, meet: "Alle benodigde data/materialen zijn aanwezig." },
  { key: "kwaliteit", label: "Datakwaliteit", weight: 5, meet: "Formaat, resolutie en inhoud zijn correct." },
  { key: "duidelijkheid", label: "Eenduidigheid", weight: 3, meet: "Geen interpretatie of vragen nodig om te starten." },
  { key: "tijdigheid", label: "Tijdigheid", weight: 3, meet: "Beschikbaar op het geplande moment." },
  { key: "standaard", label: "Standaardisatie", weight: 1, meet: "Conform naamgeving en protocollen." },
  { key: "overdracht", label: "Overdracht", weight: 1, meet: "Status correct bijgewerkt in bronsystemen." }
]);

export const ACTIVITY_TYPES = deepFreeze([
  { value: "Taak", label: "📝 Taak" },
  { value: "Afspraak", label: "📅 Afspraak" },
  { value: "TimeOut", label: "🛑 Time Out" },
  { value: "Beoordeling", label: "🔎 Beoordeling" }
]);

export const LEAN_VALUES = deepFreeze([
  { value: "VA", label: "VA - Patiëntwaarde" },
  { value: "BNVA", label: "BNVA - Business Noodzaak" },
  { value: "NVA", label: "NVA - Verspilling" }
]);

export const PROCESS_STATUSES = deepFreeze([
  { value: "SAD", label: "Niet in control", emoji: "☹️", class: "selected-sad" },
  { value: "NEUTRAL", label: "Aandachtspunt", emoji: "😐", class: "selected-neu" },
  { value: "HAPPY", label: "In control", emoji: "🙂", class: "selected-hap" }
]);

export const DISRUPTION_FREQUENCIES = deepFreeze(["Zelden", "Soms", "Vaak", "Altijd"]);
export const DEFINITION_TYPES = deepFreeze([
  { value: "HARD", label: "Hard" },
  { value: "SOFT", label: "Soft" }
]);

export const uid = () => {
  if (typeof crypto !== "undefined" && crypto.randomUUID) return crypto.randomUUID();
  return Date.now().toString(36) + Math.random().toString(36).substring(2);
};

const createInitialQA = () => {
  const qa = {};
  IO_CRITERIA.forEach((c) => (qa[c.key] = { result: "", note: "" }));
  return qa;
};

const createInitialSystemData = () => {
  const sys = { calculatedScore: null };
  SYSTEM_QUESTIONS.forEach((q) => (sys[q.id] = 0));
  return sys;
};

export const createSticky = () => ({
  id: uid(),
  created: Date.now(),
  text: "",
  linkedSourceId: null,
  type: null,
  processValue: null,
  processStatus: null,
  successFactors: "",
  causes: [],
  improvements: [],
  qa: createInitialQA(),
  systemData: createInitialSystemData(),
  inputDefinitions: [],
  disruptions: []
});

export const createColumn = () => ({
  id: uid(),
  isVisible: true,
  isParallel: false,
  hasTransition: false,
  transitionNext: "",
  outputId: null,
  slots: Array.from({ length: APP_CONFIG.SLOT_COUNT }, () => createSticky())
});

/**
 * outputMerges: array of ranges for OUTPUT row only.
 * Each: { id, slotIdx: 4, startCol, endCol }
 */
export const createSheet = (name = DEFAULTS.SHEET_NAME) => ({
  id: uid(),
  name,
  columns: [createColumn()],
  outputMerges: [] // NEW
});

export const createProjectState = () => {
  const firstSheet = createSheet();
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