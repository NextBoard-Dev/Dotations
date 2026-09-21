const DEFAULT_FILTERS = {
  search: "",
  site: "",
  typePersonnel: "",
  typeContrat: "",
  statutDossier: "",
  statutObjet: "",
  typeEffet: "",
};
const DEFAULT_TABLE_SORTS = {
  sheetEffects: { key: "typeEffet", dir: "asc" },
  arrivalEffects: { key: "typeEffet", dir: "asc" },
  exitEffects: { key: "typeEffet", dir: "asc" },
  overviewPersons: { key: "nom", dir: "asc" },
  referenceEffects: { key: "site", dir: "asc" },
  documentsArchives: { key: "nom", dir: "asc" },
};

const state = {
  data: null,
  supabaseRevision: null,
  latestDataEtag: "",
  currentSheetPersonId: "",
  editingEffectId: "",
  editingReferenceId: "",
  editingReplacementCostKey: "",
  editingSimpleReference: null,
  editingRepresentativeId: "",
  isDirty: false,
  shortcutsBound: false,
  undoStack: [],
  statusTimerId: 0,
  pdfProgressTimerId: 0,
  pdfGenerationActive: false,
  mobileSignaturePollTimerId: 0,
  mobileSignaturePollInFlight: false,
  mobileSignaturePollBackoffUntil: 0,
  mobileSignaturePollErrorCount: 0,
  mobileSignaturePollResumeTimerId: 0,
  mobileSignaturePollRenderSyncTimerId: 0,
  mobileSignaturePollRenderSyncAt: 0,
  mobileSignaturePollHasPendingRequest: true,
  mobileSignaturePollIntervalMs: 0,
  mobileSignaturePollLastSyncAt: 0,
  mobileSignaturePollRequestCache: new Map(),
  mobileSignaturePollStateSignature: "",
  mobileSignaturePollSyncModeSignature: "",
  mobileSignaturePollStatusSignature: "",
  mobileSignatureRecoveryModalOpen: false,
  mobileSignatureVisibilityBound: false,
  browserStorageQuotaLevel: 0,
  mobileSignatureNetworkInfo: null,
  autoSaveNavigationBound: false,
  searchClearBrowserEventsBound: false,
  dirtyFallbackBound: false,
  saveButtonLatchedDirty: false,
  autoSaveInFlightPromise: null,
  autoSaveTimerId: 0,
  actionAutoSaveTimerId: 0,
  tableSorts: {
    sheetEffects: { ...DEFAULT_TABLE_SORTS.sheetEffects },
    arrivalEffects: { ...DEFAULT_TABLE_SORTS.arrivalEffects },
    exitEffects: { ...DEFAULT_TABLE_SORTS.exitEffects },
    overviewPersons: { ...DEFAULT_TABLE_SORTS.overviewPersons },
    documentsArchives: { ...DEFAULT_TABLE_SORTS.documentsArchives },
  },
  filters: { ...DEFAULT_FILTERS },
  referenceRenderContext: null,
  urgentMode: false,
  lastSaveInfo: null,
  lastPersistedDataSignature: null,
  effectRowFlash: null,
  effectTableFlash: null,
  autoPdfGenerationInFlight: false,
  autoPdfGeneratedKeys: new Set(),
  signedDocumentsPopupSeenKeys: new Set(),
  previousSignatureValidationMap: new Map(),
  stockTableFilters: { site: "", typeEffet: "", referenceEffetId: "" },
  stockHighlightKey: "",
  isReferencePageResetting: false,
  currentUserRoleLabel: "",
  adminCreateUserInFlight: false,
  saveInFlight: false,
  lastPdfOpenKey: "",
  lastPdfOpenAtMs: 0,
  latestDataFetchPromise: null,
  latestDataFetchAt: 0,
  latestDataSnapshotCache: null,
  hostedSyncState: "unknown",
  hostedSyncInFlight: false,
  hostedSyncDetails: "",
  documentRenderCache: { arrival: "", exit: "" },
  documentCostRenderCache: { arrival: "", exit: "", mobile: "" },
  documentViewRenderCache: { arrival: "", exit: "" },
  networkDebug: {
    samples: [],
    routeStats: {},
    requestCount: 0,
    totalBytes: 0,
    startedAt: Date.now(),
  },
  listRenderCache: {
    overview: "",
    global: "",
    overviewAlerts: "",
    mobileSignature: "",
    documentsArchives: "",
    referenceBases: "",
    stockMovements: "",
    stockSummary: "",
    stockKpis: "",
    stockInstantKpi: "",
    referenceCounts: "",
  },
  pageRenderRafId: 0,
  filterInputDebounceTimerId: 0,
  referenceFilterDebounceTimerId: 0,
    localMutationTick: 0,
    mobileSignatureRequestContextCache: new Map(),
    personPickerRenderCache: {
    "person-sheet": "",
    "arrival-document": "",
    "exit-document": "",
  },
  filteredPersonsCache: {
    key: "",
    persons: [],
  },
  dirtyStateRenderSignature: "",
  pageRenderSignature: "",
};
const MOBILE_SIGNATURE_POLL_INTERVAL_MS = 60000;
const MOBILE_SIGNATURE_POLL_ERROR_BACKOFF_MS = 30000;
const MOBILE_SIGNATURE_POLL_RESUME_DELAY_MS = 10000;
const MOBILE_SIGNATURE_POLL_SYNC_MIN_GAP_MS = 10000;
const MOBILE_SIGNATURE_POLL_RENDER_SYNC_DEBOUNCE_MS = 60000;
const MOBILE_SIGNATURE_POLL_IDLE_INTERVAL_MS = 180000;
const BROWSER_STORAGE_WARNING_RATIO = 0.8;
const BROWSER_STORAGE_ALERT_RATIO = 0.9;
const DATA_FETCH_DEBOUNCE_MS = 5000;
const FILTER_INPUT_DEBOUNCE_MS = 140;
const NETWORK_DEBUG_SAMPLE_LIMIT = 80;
const NETWORK_DEBUG_STORAGE_KEY = "dotations-network-debug-v1";
const NETWORK_DEBUG_WINDOW_MS = 10 * 60 * 1000;

const WORKING_DATA_KEY = "dashboard-working-data";
const LEGACY_CONTRACT_TYPES = ["CDI", "CDD", "INTERIMAIRE"];
const MAX_UNDO_STACK = 30;
const ALL_SITES_VALUE = "TOUS SITES";
const ALL_TYPES_VALUE = "TOUS TYPES";
const ALL_DESIGNATIONS_VALUE = "__ALL_DESIGNATIONS__";
const STOCK_SYNTHETIC_REFERENCE_PREFIX = "__STOCK_SYNTHETIC__:";
const STOCK_EMPTY_DESIGNATION_LABEL = "SANS DESIGNATION";
const EFFECT_TYPES_WITHOUT_REFERENCE_DESIGNATION = ["RADIATEUR APPOINT"];
const EFFECT_STATUS_CAUSES = ["HS", "PERTE", "VOL", "NON RENDU", "DETRUIT"];
const BILLABLE_EFFECT_CAUSES = ["PERTE", "VOL", "NON RENDU", "DETRUIT"];
const NON_RENDU_REFERENCE_COSTS = {
  "BADGE INTRUSION": 15,
  "CARTE TURBOSELF": 10,
  CLE: 5,
  "CLE CES": 50,
  "CLE DE SECURITE": 45,
  "RADIATEUR APPOINT": 45,
  "TELECOMMANDE URMET": 40,
  VENTILATEUR: 30,
};
const MOBILE_SIGNATURE_REQUEST_TTL_MS = 10 * 60 * 1000;
const AUTO_GENERATE_SIGNED_PDFS = false;
const SUPABASE_PROJECT_URL = "https://dphrvdhqhgycmllietuk.supabase.co";
const SUPABASE_PUBLISHABLE_KEY = "sb_publishable_2wYXnIDj4-c8daQZW8D5hA_2Py6k7z6";
const SUPABASE_EDGE_API_URL = "https://dphrvdhqhgycmllietuk.supabase.co/functions/v1/dotations-api";
const SUPABASE_APP_STATE_TABLE = "app_state";
const SUPABASE_APP_STATE_ID = "main";
const DEFAULT_SUPABASE_PDF_BUCKET = "pdf";
const DEFAULT_SUPABASE_SIGNATURES_BUCKET = "signatures";
const NAVIGATION_CONTEXT_KEY = "dotations-navigation-context";
const SIGNED_POPUP_SEEN_STORAGE_KEY = "dotations-signed-popup-seen";
const PENDING_PDF_REMINDER_SNOOZE_KEY = "dotations-pending-pdf-reminder-snooze";
const PENDING_PDF_TASK_STORAGE_KEY = "dotations-pending-pdf-task";
const LOGIN_EVENT_MARKER_KEY = "dotations-login-event-sent";
const ADMIN_CONTACT_EMAIL = "sebastien.duc@outlook.fr";
const PASSWORD_RESET_COOLDOWN_KEY = "dotations-reset-password-last-sent-at";
const PASSWORD_RESET_COOLDOWN_MS = 70 * 1000;
const ADMIN_INVITE_COOLDOWN_KEY = "dotations-admin-invite-last-sent-at";
const ADMIN_INVITE_COOLDOWN_MS = 70 * 1000;
const ADMIN_INTERACTIVE_AUTH_MARKER_KEY = "dotations-admin-interactive-auth-ok";
const LOGIN_EMAIL_SUGGESTIONS_KEY = "dotations-login-email-suggestions-v1";
const LAST_LOGIN_EMAIL_KEY = "dotations-last-login-email-v1";
const PDF_LAYOUT_VERSION = "2026-03-14-exit-layout-fix-3";
const PDF_FORMAT_LOCK = "v1";
let pdfModalCleanupBound = false;
let reminderSnoozeMap = {};
const signatureCanvases = new WeakMap();
const EFFECT_EDIT_CONTEXT_STORAGE_KEY = "dotations-pending-effect-edit-context-v1";
const SAVE_AUDIT_LOG_KEY = "dotations-save-audit-log-v1";
const MAX_SAVE_AUDIT_ENTRIES = 250;

try {
  reminderSnoozeMap = JSON.parse(localStorage.getItem(PENDING_PDF_REMINDER_SNOOZE_KEY) || "{}") || {};
} catch (error) {
  reminderSnoozeMap = {};
}

function getPendingPdfTaskFromStorage() {
  try {
    const raw = JSON.parse(localStorage.getItem(PENDING_PDF_TASK_STORAGE_KEY) || "null");
    if (!raw || typeof raw !== "object") {
      return null;
    }
    const personId = String(raw.personId || "");
    const docType = String(raw.docType || "");
    const validatedAt = String(raw.validatedAt || "");
    if (!personId || !docType || !validatedAt) {
      return null;
    }
    return { personId, docType, validatedAt };
  } catch (error) {
    return null;
  }
}

function setPendingPdfTaskToStorage(task) {
  try {
    if (!task) {
      localStorage.removeItem(PENDING_PDF_TASK_STORAGE_KEY);
      return;
    }
    localStorage.setItem(PENDING_PDF_TASK_STORAGE_KEY, JSON.stringify(task));
  } catch (error) {
    // ignore storage failures
  }
}

function clearPendingPdfTaskFor(personId, docType) {
  const current = getPendingPdfTaskFromStorage();
  if (!current) {
    return;
  }
  if (String(current.personId) === String(personId) && String(current.docType) === String(docType)) {
    setPendingPdfTaskToStorage(null);
  }
}

function clearPendingPdfTaskIfArchived() {
  const pendingTask = getPendingPdfTaskFromStorage();
  if (!pendingTask || !state.data) {
    return false;
  }
  const personId = String(pendingTask.personId || "");
  const docType = String(pendingTask.docType || "");
  if (!personId || !docType) {
    setPendingPdfTaskToStorage(null);
    return true;
  }
  const person = (state.data.personnes || []).find((candidate) => String(candidate.id || "") === personId);
  if (!person) {
    return false;
  }
  const hasArchive = Boolean(findReusableArchivedDocument(person, docType));
  if (!hasArchive) {
    return false;
  }
  clearPendingPdfTaskFor(personId, docType);
  const snoozeKey = `${personId}:${docType}`;
  delete reminderSnoozeMap[snoozeKey];
  try {
    localStorage.setItem(PENDING_PDF_REMINDER_SNOOZE_KEY, JSON.stringify(reminderSnoozeMap));
  } catch (error) {
    // ignore storage failures
  }
  return true;
}

function getCachedLoginEmailSuggestions() {
  try {
    const raw = JSON.parse(localStorage.getItem(LOGIN_EMAIL_SUGGESTIONS_KEY) || "[]");
    if (!Array.isArray(raw)) return [];
    return Array.from(
      new Set(
        raw
          .map((email) => String(email || "").trim().toLowerCase())
          .filter(Boolean),
      ),
    ).slice(0, 50);
  } catch (error) {
    return [];
  }
}

function setCachedLoginEmailSuggestions(emails = []) {
  try {
    const normalized = Array.from(
      new Set(
        (emails || [])
          .map((email) => String(email || "").trim().toLowerCase())
          .filter(Boolean),
      ),
    ).slice(0, 50);
    localStorage.setItem(LOGIN_EMAIL_SUGGESTIONS_KEY, JSON.stringify(normalized));
  } catch (error) {
    // ignore storage failures
  }
}

function mergeLoginEmailSuggestions(emails = []) {
  const merged = Array.from(new Set([...getCachedLoginEmailSuggestions(), ...(emails || [])]));
  setCachedLoginEmailSuggestions(merged);
  return merged;
}

async function fetchLoginEmailSuggestionsFromBackend(accessToken) {
  const token = String(accessToken || "").trim();
  const baseUrl = normalizeHttpUrl(SUPABASE_EDGE_API_URL);
  const key = String(SUPABASE_PUBLISHABLE_KEY || "").trim();
  if (!token || !baseUrl || !key) return [];
  try {
    const response = await fetch(`${baseUrl}/admin/users`, {
      method: "GET",
      headers: {
        apikey: key,
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      cache: "no-store",
    });
    if (!response.ok) return [];
    const payload = await response.json().catch(() => ({}));
    const users = normalizeAdminUsersList(payload);
    const authBackedUsers = users.filter((entry) => isAdminUserResettable(entry));
    const source = authBackedUsers.length ? authBackedUsers : users;
    return source
      .map((entry) => String(entry?.email || "").trim().toLowerCase())
      .filter(Boolean);
  } catch (error) {
    return [];
  }
}

function setKpiCountAnimated(node, nextValue) {
  if (!node) {
    return;
  }
  const target = Number(nextValue);
  if (!Number.isFinite(target)) {
    node.textContent = String(nextValue);
    return;
  }
  const current = Number.parseInt(String(node.dataset.kpiValue || node.textContent || "0"), 10);
  if (!Number.isFinite(current) || current === target) {
    node.textContent = String(target);
    node.dataset.kpiValue = String(target);
    return;
  }

  const start = current;
  const delta = target - start;
  const duration = 260;
  const startAt = performance.now();
  node.classList.remove("kpi-value--changed");
  void node.offsetWidth;
  node.classList.add("kpi-value--changed");

  const tick = (now) => {
    const progress = Math.min(1, (now - startAt) / duration);
    const eased = 1 - Math.pow(1 - progress, 3);
    const value = Math.round(start + delta * eased);
    node.textContent = String(value);
    if (progress < 1) {
      window.requestAnimationFrame(tick);
      return;
    }
    node.dataset.kpiValue = String(target);
    window.setTimeout(() => {
      node.classList.remove("kpi-value--changed");
    }, 280);
  };

  window.requestAnimationFrame(tick);
}

function pulseSaveButtons() {
  document.querySelectorAll(".js-save-data").forEach((button) => {
    if (!(button instanceof HTMLElement)) {
      return;
    }
    button.classList.remove("button--save-ok");
    void button.offsetWidth;
    button.classList.add("button--save-ok");
    window.setTimeout(() => {
      button.classList.remove("button--save-ok");
    }, 520);
  });
}

function bindGlobalButtonClickFeedback() {
  const selector = "button, .button, .btn, .sidebar__link, .tab";
  document.addEventListener(
    "click",
    (event) => {
      const target = event.target instanceof Element ? event.target : null;
      const button = target ? target.closest(selector) : null;
      if (!button) {
        return;
      }
      if (button.hasAttribute("disabled") || button.classList.contains("is-disabled")) {
        return;
      }
      button.classList.remove("btn-click-ack");
      void button.offsetWidth;
      button.classList.add("btn-click-ack");
      window.setTimeout(() => {
        button.classList.remove("btn-click-ack");
      }, 190);
    },
    true
  );
}

redirectToLocalServerIfNeeded();
applyPdfModeFromQuery();
bindGlobalButtonClickFeedback();

function normalizeText(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .trim();
}

function normalizeFunctionLabel(value) {
  const normalized = normalizeText(value);
  const corrections = {
    ENSEIGNENT: "ENSEIGNANT",
    "LABORENTIN(E)": "LABORANTIN(E)",
    RESPONSBLE: "RESPONSABLE",
    "RESPONSBLE INFORMATIQUE": "RESPONSABLE INFORMATIQUE",
  };
  return corrections[normalized] || normalized;
}

function normalizeAmount(value) {
  const normalized = String(value ?? "")
    .replace(/\s/g, "")
    .replace(",", ".")
    .replace(/[^0-9.-]/g, "");
  const parsed = Number.parseFloat(normalized);
  return Number.isFinite(parsed) ? Number(parsed.toFixed(2)) : 0;
}

function normalizeEffectCause(value) {
  const normalized = normalizeText(value);
  if (normalized === "CASSE") return "DETRUIT";
  if (normalized === "PERDU") return "PERTE";
  if (["DETRUIT", "PERTE", "VOL", "HS", "NON RENDU"].includes(normalized)) return normalized;
  return "";
}

function normalizePricingKey(value, { cause = false } = {}) {
  let normalized = normalizeText(value)
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!cause) {
    return normalized;
  }
  if (normalized === "NON RENDU") return "NON RENDU";
  if (normalized === "PERDU") return "PERTE";
  if (normalized === "CASSE") return "DETRUIT";
  return normalized;
}

function getFallbackNonRenduCost(typeEffet, designation = "") {
  const normalizedType = normalizePricingKey(typeEffet);
  const normalizedDesignation = normalizePricingKey(designation);
  if (!normalizedType) return 0;
  if (normalizedType === "CLE") {
    return isCesKeyDesignation(designation) ? 50 : 5;
  }
  if (normalizedType === "CLE DE SECURITE") {
    return 45;
  }
  if (normalizedType === "VENTILATEUR") {
    return normalizedDesignation === "VENTILATEUR SUR PIED" ? 35 : 30;
  }
  return NON_RENDU_REFERENCE_COSTS[normalizedType] || 0;
}

function getCauseFromManualStatus(manualStatus) {
  const normalized = normalizeText(manualStatus);
  if (normalized === "PERDU") return "PERTE";
  if (normalized === "DETRUIT") return "DETRUIT";
  if (normalized === "VOL") return "VOL";
  if (normalized === "HS") return "HS";
  if (normalized === "CASSE") return "DETRUIT";
  return "";
}

function normalizeHttpUrl(value) {
  const raw = String(value || "").trim();
  if (!raw) {
    return "";
  }
  try {
    const parsed = new URL(raw);
    if (!["http:", "https:"].includes(parsed.protocol)) {
      return "";
    }
    return parsed.href.replace(/\/$/, "");
  } catch (error) {
    return "";
  }
}

function normalizeMobileSignatureBaseUrl(value) {
  const normalized = normalizeHttpUrl(value);
  if (!normalized) {
    return "";
  }
  try {
    const parsed = new URL(normalized);
    const host = String(parsed.hostname || "").toLowerCase();
    if (host === "mililumatt.github.io") {
      parsed.hostname = "nextboard-dev.github.io";
      return parsed.href.replace(/\/$/, "");
    }
    return normalized;
  } catch (error) {
    return normalized;
  }
}

function normalizeBucketName(value, fallback = "") {
  const normalized = String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9_-]/g, "");
  return normalized || String(fallback || "").trim().toLowerCase();
}

function isLocalRuntime() {
  const host = String(window.location.hostname || "").toLowerCase();
  const isPrivateIpv4 =
    /^10\.\d{1,3}\.\d{1,3}\.\d{1,3}$/.test(host) ||
    /^192\.168\.\d{1,3}\.\d{1,3}$/.test(host) ||
    /^172\.(1[6-9]|2\d|3[0-1])\.\d{1,3}\.\d{1,3}$/.test(host);
  return (
    host === "localhost" ||
    host === "127.0.0.1" ||
    isPrivateIpv4
  );
}

function isSupabaseConfigured() {
  const url = normalizeHttpUrl(SUPABASE_PROJECT_URL);
  const key = String(SUPABASE_PUBLISHABLE_KEY || "").trim();
  return Boolean(url && key && key.startsWith("sb_"));
}

function getDataBackendMode() {
  if (isLocalRuntime()) {
    return "LOCAL_API";
  }
  if (isSupabaseConfigured()) {
    return "SUPABASE";
  }
  return "HOSTED_NO_BACKEND";
}

function getSupabaseRestEndpoint() {
  const baseUrl = normalizeHttpUrl(SUPABASE_PROJECT_URL);
  return `${baseUrl}/rest/v1/${SUPABASE_APP_STATE_TABLE}`;
}

function openAdminContactMailto(contextLabel = "") {
  const email = String(ADMIN_CONTACT_EMAIL || "").trim();
  if (!email) {
    return false;
  }
  const subject = encodeURIComponent("SUIVI DES DOTATIONS - Demande identifiant oublié");
  const context = String(contextLabel || "").trim();
  const pageUrl = String(window.location.href || "");
  const requestedAt = new Date().toLocaleString("fr-FR");
  const body = encodeURIComponent(
    [
      "SUIVI DES DOTATIONS - ENTREE / SORTIE",
      "--------------------------------------",
      "",
      "DEMANDE : IDENTIFIANT OUBLIE",
      "Une demande utilisateur a ete declenchee.",
      "",
      "CONTEXTE",
      `- Application : SUIVI DES DOTATIONS ENTREE / SORTIE`,
      `- Canal       : ${context || "N/A"}`,
      `- Date        : ${requestedAt}`,
      `- Page        : ${pageUrl || "N/A"}`,
      "",
      "ACTION ADMIN DEMANDEE",
      "- Verifier le compte utilisateur.",
      "- Communiquer l'identifiant de connexion a l'utilisateur.",
      "",
      "Message genere automatiquement depuis la page de connexion.",
    ].join("\n")
  );
  window.location.href = `mailto:${encodeURIComponent(email)}?subject=${subject}&body=${body}`;
  return true;
}

function getSupabaseProjectRef() {
  const baseUrl = normalizeHttpUrl(SUPABASE_PROJECT_URL);
  if (!baseUrl) return "";
  try {
    return String(new URL(baseUrl).hostname || "").split(".")[0] || "";
  } catch (error) {
    return "";
  }
}

function getSupabaseAuthStorageKey() {
  const projectRef = getSupabaseProjectRef();
  return projectRef ? `sb-${projectRef}-auth-token` : "";
}

function mapRoleToFrenchLabel(role) {
  const normalized = String(role || "").trim().toLowerCase();
  if (normalized === "admin") return "ADMIN";
  if (normalized === "editor") return "EDITEUR";
  return "LECTURE";
}

function extractAccessTokenFromStoredSession(raw) {
  if (!raw) return "";
  if (typeof raw === "string") {
    try {
      return extractAccessTokenFromStoredSession(JSON.parse(raw));
    } catch (error) {
      return "";
    }
  }
  if (Array.isArray(raw)) {
    for (const item of raw) {
      const token = extractAccessTokenFromStoredSession(item);
      if (token) return token;
    }
    return "";
  }
  if (typeof raw !== "object") return "";
  const direct = String(raw.access_token || raw.accessToken || "").trim();
  if (direct) return direct;
  return (
    extractAccessTokenFromStoredSession(raw.currentSession) ||
    extractAccessTokenFromStoredSession(raw.session) ||
    ""
  );
}

function extractSessionFromStoredSession(raw) {
  if (!raw) return null;
  if (typeof raw === "string") {
    try {
      return extractSessionFromStoredSession(JSON.parse(raw));
    } catch (error) {
      return null;
    }
  }
  if (Array.isArray(raw)) {
    for (const item of raw) {
      const session = extractSessionFromStoredSession(item);
      if (session) return session;
    }
    return null;
  }
  if (typeof raw !== "object") return null;
  if (raw.currentSession && typeof raw.currentSession === "object") {
    return raw.currentSession;
  }
  if (raw.session && typeof raw.session === "object") {
    return raw.session;
  }
  if (raw.access_token) {
    return raw;
  }
  return null;
}

function getStoredSupabaseAccessToken() {
  try {
    const storageKey = getSupabaseAuthStorageKey();
    if (!storageKey) return "";
    const sessionRaw = sessionStorage.getItem(storageKey);
    const localRaw = localStorage.getItem(storageKey);
    return extractAccessTokenFromStoredSession(sessionRaw) || extractAccessTokenFromStoredSession(localRaw);
  } catch (error) {
    return "";
  }
}

function getStoredSupabaseSession() {
  try {
    const storageKey = getSupabaseAuthStorageKey();
    if (!storageKey) return null;
    const sessionRaw = sessionStorage.getItem(storageKey);
    const localRaw = localStorage.getItem(storageKey);
    return extractSessionFromStoredSession(sessionRaw) || extractSessionFromStoredSession(localRaw);
  } catch (error) {
    return null;
  }
}

function getStoredSupabaseUserId() {
  const session = getStoredSupabaseSession();
  return String(session?.user?.id || "").trim();
}

function storeSupabaseSession(session) {
  try {
    const storageKey = getSupabaseAuthStorageKey();
    if (!storageKey || !session || typeof session !== "object") return;
    const payload = JSON.stringify({
      currentSession: session,
      expiresAt: session?.expires_at || 0,
    });
    sessionStorage.setItem(storageKey, payload);
    localStorage.setItem(storageKey, payload);
  } catch (error) {
    // ignore storage failures
  }
}

function importSupabaseSessionFromUrlIfPresent() {
  try {
    const url = new URL(window.location.href);
    const hashParams = new URLSearchParams(url.hash ? url.hash.replace(/^\#/, "") : "");

    const accessToken = String(url.searchParams.get("sbat") || hashParams.get("sbat") || "").trim();
    const refreshToken = String(url.searchParams.get("sbrt") || hashParams.get("sbrt") || "").trim();
    const expiresAtRaw = String(url.searchParams.get("sbea") || hashParams.get("sbea") || "").trim();
    const hasSessionBridge = Boolean(accessToken || refreshToken || expiresAtRaw);
    if (!hasSessionBridge) {
      return;
    }
    const expiresAt = Number.parseInt(expiresAtRaw, 10);
    storeSupabaseSession({
      access_token: accessToken,
      refresh_token: refreshToken,
      expires_at: Number.isFinite(expiresAt) ? expiresAt : 0,
      token_type: "bearer",
    });
    url.searchParams.delete("sbat");
    url.searchParams.delete("sbrt");
    url.searchParams.delete("sbea");
    window.history.replaceState({}, "", url);
  } catch (error) {
    // ignore malformed URL params
  }
}

function appendSupabaseSessionBridgeParams(url, options = {}) {
  try {
    const session = getStoredSupabaseSession();
    const accessToken = String(session?.access_token || "").trim();
    const refreshToken = String(session?.refresh_token || "").trim();
    const expiresAt = Number.parseInt(String(session?.expires_at || "0"), 10);
    const includeAccessToken = options?.includeAccessToken !== false;
    const includeRefreshToken = options?.includeRefreshToken !== false;
    const includeExpiresAt = options?.includeExpiresAt !== false;
    if (!accessToken && !refreshToken) {
      return url;
    }
    const nextUrl = new URL(url, window.location.href);
    if (includeAccessToken && accessToken) {
      nextUrl.searchParams.set("sbat", accessToken);
    } else {
      nextUrl.searchParams.delete("sbat");
    }
    if (includeRefreshToken && refreshToken) {
      nextUrl.searchParams.set("sbrt", refreshToken);
    } else {
      nextUrl.searchParams.delete("sbrt");
    }
    if (includeExpiresAt && Number.isFinite(expiresAt) && expiresAt > 0) {
      nextUrl.searchParams.set("sbea", String(expiresAt));
    } else {
      nextUrl.searchParams.delete("sbea");
    }
    return nextUrl.toString();
  } catch (error) {
    return url;
  }
}

function clearStoredSupabaseSession() {
  try {
    const storageKey = getSupabaseAuthStorageKey();
    if (storageKey) {
      sessionStorage.removeItem(storageKey);
      localStorage.removeItem(storageKey);
    }
    sessionStorage.removeItem(ADMIN_INTERACTIVE_AUTH_MARKER_KEY);
  } catch (error) {
    // ignore storage failures
  }
}

async function promptSupabaseLoginAndStoreSession() {
  const credentials = await promptSupabaseCredentialsForm();
  const email = String(credentials?.email || "").trim();
  const password = String(credentials?.password || "");
  if (!email || !password) throw new Error("BACKEND_AUTH_REQUIRED");
  if (getDataBackendMode() === "LOCAL_API") {
    const response = await fetch("/api/local-login", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ username: email, password }),
      cache: "no-store",
    });
    if (!response.ok) {
      throw new Error("BACKEND_AUTH_INVALID");
    }
    const payload = await response.json().catch(() => ({}));
    state.currentUserRoleLabel = String(payload?.role_label || "LOCAL_SECOURS");
    try {
      sessionStorage.setItem(ADMIN_INTERACTIVE_AUTH_MARKER_KEY, "1");
    } catch (error) {
      // ignore marker failures
    }
    return "LOCAL_MODE_TOKEN";
  }
  const baseUrl = normalizeHttpUrl(SUPABASE_PROJECT_URL);
  const key = String(SUPABASE_PUBLISHABLE_KEY || "").trim();
  const response = await fetch(`${baseUrl}/auth/v1/token?grant_type=password`, {
    method: "POST",
    headers: {
      apikey: key,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ email, password }),
    cache: "no-store",
  });
  if (!response.ok) {
    throw new Error("BACKEND_AUTH_INVALID");
  }
  const data = await response.json().catch(() => null);
  const accessToken = String(data?.access_token || "").trim();
  if (!accessToken) {
    throw new Error("BACKEND_AUTH_INVALID");
  }
  storeSupabaseSession(data);
  try {
    sessionStorage.setItem(ADMIN_INTERACTIVE_AUTH_MARKER_KEY, "1");
  } catch (error) {
    // ignore marker failures
  }
  await logAdminLoginEvent(accessToken, data);
  return accessToken;
}

async function requestSupabasePasswordReset(email, options = {}) {
  const safeEmail = String(email || "").trim();
  if (!safeEmail) {
    throw new Error("EMAIL_OBLIGATOIRE");
  }
  const bypassCooldown = options?.bypassCooldown === true;
  const now = Date.now();
  if (!bypassCooldown) {
    const lastSentAt = Number.parseInt(String(localStorage.getItem(PASSWORD_RESET_COOLDOWN_KEY) || ""), 10);
    if (Number.isFinite(lastSentAt)) {
      const remainingMs = PASSWORD_RESET_COOLDOWN_MS - (now - lastSentAt);
      if (remainingMs > 0) {
        const remainingSec = Math.ceil(remainingMs / 1000);
        throw new Error(`RESET_MDP_COOLDOWN:${remainingSec}`);
      }
    }
  }
  const resetRedirectTo = "https://nextboard-dev.github.io/Dotations/index.html?view=desktop";
  const baseUrl = normalizeHttpUrl(SUPABASE_PROJECT_URL);
  const key = String(SUPABASE_PUBLISHABLE_KEY || "").trim();
  const response = await fetch(`${baseUrl}/auth/v1/recover`, {
    method: "POST",
    headers: {
      apikey: key,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      email: safeEmail,
      redirect_to: resetRedirectTo,
    }),
    cache: "no-store",
  });
  if (!response.ok) {
    const text = await response.text().catch(() => "");
    throw new Error(`RESET_MDP_ECHEC:${response.status}:${text}`);
  }
  if (!bypassCooldown) {
    localStorage.setItem(PASSWORD_RESET_COOLDOWN_KEY, String(now));
  }
}

async function requestSupabaseMagicLink(email) {
  const safeEmail = String(email || "").trim();
  if (!safeEmail) {
    throw new Error("EMAIL_OBLIGATOIRE");
  }
  const redirectTo = "https://nextboard-dev.github.io/Dotations/";
  const baseUrl = normalizeHttpUrl(SUPABASE_PROJECT_URL);
  const key = String(SUPABASE_PUBLISHABLE_KEY || "").trim();
  const response = await fetch(`${baseUrl}/auth/v1/otp`, {
    method: "POST",
    headers: {
      apikey: key,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      email: safeEmail,
      create_user: false,
      redirect_to: redirectTo,
      email_redirect_to: redirectTo,
    }),
    cache: "no-store",
  });
  if (!response.ok) {
    const text = await response.text().catch(() => "");
    throw new Error(`MAGIC_LINK_ECHEC:${response.status}:${text}`);
  }
}

async function requestAdminUserInvitation(email, role) {
  const safeEmail = String(email || "").trim();
  if (!safeEmail) {
    throw new Error("EMAIL_OBLIGATOIRE");
  }
  try {
    await callEdgeApi("admin/users/invite", {
      method: "POST",
      body: JSON.stringify({ email: safeEmail, role: normalizeRequestedUserRole(role) }),
    });
  } catch (error) {
    const message = String(error?.message || "");
    if (/email_exists/i.test(message)) {
      throw new Error("INVITE_USER_EXISTS");
    }
    if (message.includes("EDGE_API_FAILED:404")) {
      throw new Error("INVITE_ENDPOINT_ABSENT");
    }
    throw error;
  }
  return true;
}

async function fetchUserRoleFromProfile(accessToken, userId) {
  const safeUserId = String(userId || "").trim();
  if (!accessToken || !safeUserId) return "viewer";
  try {
    const endpoint = `${getSupabaseRestEndpoint().replace(/\/app_state$/i, "/profiles")}?id=eq.${encodeURIComponent(
      safeUserId
    )}&select=role&limit=1`;
    const response = await fetch(endpoint, {
      method: "GET",
      headers: {
        apikey: String(SUPABASE_PUBLISHABLE_KEY || "").trim(),
        Authorization: `Bearer ${accessToken}`,
      },
      cache: "no-store",
    });
    if (!response.ok) return "viewer";
    const rows = await response.json().catch(() => []);
    const role = Array.isArray(rows) ? rows[0]?.role : "";
    return normalizeRequestedUserRole(role);
  } catch (error) {
    return "viewer";
  }
}

async function logAdminLoginEvent(accessToken, sessionPayload) {
  try {
    const userId = String(
      sessionPayload?.user?.id ||
      sessionPayload?.currentSession?.user?.id ||
      ""
    ).trim();
    const email = String(
      sessionPayload?.user?.email ||
      sessionPayload?.currentSession?.user?.email ||
      ""
    ).trim();
    if (!userId) return;
    const role = await fetchUserRoleFromProfile(accessToken, userId);
    const baseUrl = normalizeHttpUrl(SUPABASE_EDGE_API_URL);
    const key = String(SUPABASE_PUBLISHABLE_KEY || "").trim();
    if (!baseUrl || !key || !accessToken) return;
    await fetch(`${baseUrl}/admin/login-events`, {
      method: "POST",
      headers: {
        apikey: key,
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        user_id: userId,
        email,
        role,
      }),
      cache: "no-store",
    }).catch(() => null);
  } catch (error) {
    // login log is non-blocking
  }
}

async function ensureLoginEventLogged(accessToken) {
  try {
    const marker = sessionStorage.getItem(LOGIN_EVENT_MARKER_KEY);
    if (marker === "1") return;
    const session = getStoredSupabaseSession();
    if (!session?.user?.id) return;
    logAdminLoginEvent(accessToken, session).catch(() => null);
    sessionStorage.setItem(LOGIN_EVENT_MARKER_KEY, "1");
  } catch (error) {
    // ignore
  }
}

function promptSupabaseCredentialsForm() {
  return new Promise((resolve, reject) => {
    const suggestedEmails = getCachedLoginEmailSuggestions();
    const rememberedEmail = String(localStorage.getItem(LAST_LOGIN_EMAIL_KEY) || "").trim().toLowerCase();
    const primaryEmail = rememberedEmail || suggestedEmails[0] || "";
    const backdrop = document.createElement("div");
    backdrop.style.position = "fixed";
    backdrop.style.inset = "0";
    backdrop.style.background = "rgba(0,0,0,0.35)";
    backdrop.style.display = "flex";
    backdrop.style.alignItems = "center";
    backdrop.style.justifyContent = "center";
    backdrop.style.zIndex = "99999";

    const box = document.createElement("div");
    box.style.width = "min(94vw, 560px)";
    box.style.background = "#ffffff";
    box.style.borderRadius = "12px";
    box.style.padding = "16px";
    box.style.boxShadow = "0 10px 28px rgba(0,0,0,0.25)";
    box.innerHTML = `
      <div style="font-weight:700;font-size:16px;margin-bottom:10px;color:#132833;">Connexion requise</div>
      <form id="supabase-login-form" autocomplete="on">
        <label style="display:block;font-size:12px;color:#334c58;margin-bottom:6px;">Email</label>
        <input id="supabase-login-email-input" name="username" type="email" autocomplete="username email" required
          list="supabase-login-email-options" placeholder="Selectionner un compte..."
          style="width:100%;padding:10px;border:1px solid #9bb2be;border-radius:8px;margin-bottom:10px;background:#fff;" />
        <datalist id="supabase-login-email-options">
          ${suggestedEmails.map((email) => `<option value="${escapeHtml(email)}"></option>`).join("")}
        </datalist>
        <label style="display:block;font-size:12px;color:#334c58;margin-bottom:6px;">Mot de passe</label>
        <div style="display:flex;align-items:center;gap:6px;margin-bottom:12px;">
          <input id="supabase-login-password" name="password" type="password" autocomplete="current-password" required
            style="flex:1 1 auto;width:100%;padding:10px;border:1px solid #9bb2be;border-radius:8px;" />
          <button type="button" id="supabase-login-toggle-password" aria-label="Afficher le mot de passe"
            style="position:relative;border:1px solid #9bb2be;background:#fff;border-radius:8px;width:40px;min-width:40px;height:40px;min-height:40px;padding:0;cursor:pointer;overflow:hidden;background-size:18px 18px;background-repeat:no-repeat;background-position:center;background-image:url(&quot;data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%230f172a' stroke-width='1.8' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M2 3l19.5 19.5'/%3E%3Cpath d='M4.71 4.71C3.2 6.03 2 7.88 2 7.88s3.5 6 10 6c1.25 0 2.44-0.28 3.53-0.77'/%3E%3Cpath d='M8.53 8.53a3 3 0 0 0 4.24 4.24'/%3E%3Cpath d='M14.48 14.48c-0.88 0.37-1.84 0.56-2.86 0.56-6.5 0-10-6-10-6a17.5 17.5 0 0 1 5.37-5.1'/%3E%3Cpath d='M22 12s-3.5 6-10 6c-0.9 0-1.75-0.07-2.56-0.2M15.5 15.5L9.95 9.95'/%3E%3C/svg%3E');font-size:0;color:transparent;text-indent:-9999px;-webkit-appearance:none;appearance:none;"></button>
        </div>
        <div id="supabase-login-status" style="font-size:12px;color:#4a6170;margin:-6px 0 10px;"></div>
        <div style="display:flex;justify-content:flex-end;gap:8px;flex-wrap:nowrap;">
          <button type="button" id="supabase-login-forgot-id"
            style="padding:8px 12px;border:1px solid #9bb2be;background:#fff;border-radius:8px;cursor:pointer;white-space:nowrap;">Identifiant oublié</button>
          <button type="button" id="supabase-login-forgot"
            style="padding:8px 12px;border:1px solid #9bb2be;background:#fff;border-radius:8px;cursor:pointer;white-space:nowrap;">Mot de passe oublié</button>
          <button type="button" id="supabase-login-cancel"
            style="padding:8px 12px;border:1px solid #9bb2be;background:#fff;border-radius:8px;cursor:pointer;white-space:nowrap;">Annuler</button>
          <button type="submit"
            style="padding:8px 14px;border:0;background:#2f5f76;color:#fff;border-radius:8px;cursor:pointer;white-space:nowrap;">Se connecter</button>
        </div>
      </form>
    `;
    backdrop.appendChild(box);
    document.body.appendChild(backdrop);

    const form = box.querySelector("#supabase-login-form");
    const emailInput = box.querySelector("#supabase-login-email-input");
    const emailOptions = box.querySelector("#supabase-login-email-options");
    const passwordInput = box.querySelector("#supabase-login-password");
    const togglePasswordButton = box.querySelector("#supabase-login-toggle-password");
    const statusNode = box.querySelector("#supabase-login-status");
    const forgotIdButton = box.querySelector("#supabase-login-forgot-id");
    const forgotButton = box.querySelector("#supabase-login-forgot");
    const cancelButton = box.querySelector("#supabase-login-cancel");

    const renderEmailOptions = (emails = [], preferredEmail = "") => {
      if (!emailOptions) return;
      const normalized = Array.from(
        new Set((emails || []).map((entry) => String(entry || "").trim().toLowerCase()).filter(Boolean)),
      );
      emailOptions.innerHTML = normalized.map((email) => `<option value="${escapeHtml(email)}"></option>`).join("");
      const preferred = String(preferredEmail || "").trim().toLowerCase();
      if (emailInput) {
        if (preferred && normalized.includes(preferred)) {
          emailInput.value = preferred;
        } else if (normalized[0]) {
          emailInput.value = normalized[0];
        } else {
          emailInput.value = "";
        }
      }
    };

    togglePasswordButton?.addEventListener("click", () => {
      if (!passwordInput) return;
      const show = passwordInput.type === "password";
      passwordInput.type = show ? "text" : "password";
      togglePasswordButton.style.backgroundImage = show
        ? "url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%230f172a' stroke-width='1.8' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M2 3l19 19'/%3E%3Cpath d='M6.7 6.7a7 7 0 0 0-4.7 5.3s3.5 6 10 6c1.2 0 2.33-0.2 3.37-0.57'/%3E%3Cpath d='M10.7 10.7A3 3 0 0 0 12 15.5a3 3 0 0 0 2.12-5.12'/%3E%3Cpath d='M14.1 14.1A3 3 0 0 0 12 9a3 3 0 0 0-1.38 5.75'/%3E%3Cpath d='M17.35 17.35C20.08 15.83 22 12 22 12s-3.5-6-10-6c-0.7 0-1.35 0.07-1.98 0.19'/%3E%3Cpath d='M12 9.5v5'/%3E%3C/svg%3E\")"
        : "url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%230f172a' stroke-width='1.8' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M2 12s3.5-6 10-6 10 6 10 6-3.5 6-10 6S2 12 2 12Z'/%3E%3Ccircle cx='12' cy='12' r='3.5'/%3E%3C/svg%3E\")";
      togglePasswordButton.setAttribute("aria-label", show ? "Masquer le mot de passe" : "Afficher le mot de passe");
    });

    if (passwordInput && togglePasswordButton) {
      togglePasswordButton.style.backgroundImage = "url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%230f172a' stroke-width='1.8' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M2 3l19.5 19.5'/%3E%3Cpath d='M4.71 4.71C3.2 6.03 2 7.88 2 7.88s3.5 6 10 6c1.25 0 2.44-0.28 3.53-0.77'/%3E%3Cpath d='M8.53 8.53a3 3 0 0 0 4.24 4.24'/%3E%3Cpath d='M14.48 14.48c-0.88 0.37-1.84 0.56-2.86 0.56-6.5 0-10-6-10-6a17.5 17.5 0 0 1 5.37-5.1'/%3E%3Cpath d='M22 12s-3.5 6-10 6c-0.9 0-1.75-0.07-2.56-0.2M15.5 15.5L9.95 9.95'/%3E%3C/svg%3E\")";
      togglePasswordButton.setAttribute("aria-label", "Afficher le mot de passe");
    }

    const cleanup = () => {
      backdrop.remove();
    };

    cancelButton.addEventListener("click", () => {
      cleanup();
      reject(new Error("BACKEND_AUTH_REQUIRED"));
    });

    forgotIdButton.addEventListener("click", () => {
      const opened = openAdminContactMailto("PC");
      statusNode.textContent = opened
        ? "Email pre-rempli ouvert vers l'administrateur."
        : "Contact administrateur indisponible.";
      statusNode.style.color = opened ? "#2f5e43" : "#8e2c2c";
    });

    forgotButton.addEventListener("click", async () => {
      const email = String(emailInput?.value || "").trim();
      if (!email) {
        statusNode.textContent = "Saisissez d'abord votre email.";
        statusNode.style.color = "#8e2c2c";
        return;
      }
      statusNode.textContent = "Envoi email de réinitialisation...";
      statusNode.style.color = "#355464";
      try {
        await requestSupabasePasswordReset(email);
        statusNode.textContent = "Email envoyé. Vérifiez votre boîte mail.";
        statusNode.style.color = "#2f5e43";
      } catch (error) {
        const message = String(error?.message || "");
        if (message.startsWith("RESET_MDP_COOLDOWN:")) {
          const seconds = Number.parseInt(message.split(":")[1] || "60", 10) || 60;
          statusNode.textContent = `Attends ${seconds}s avant de recommencer.`;
        } else if (message.includes("429")) {
          statusNode.textContent = "Trop de demandes. Réessayez dans 1 heure.";
        } else {
          statusNode.textContent = "Échec envoi email de réinitialisation.";
        }
        statusNode.style.color = "#8e2c2c";
      }
    });

    form.addEventListener("submit", (event) => {
      event.preventDefault();
      const email = String(emailInput?.value || "").trim();
      const password = String(passwordInput.value || "");
      if (email) {
        mergeLoginEmailSuggestions([email]);
        localStorage.setItem(LAST_LOGIN_EMAIL_KEY, email.toLowerCase());
      }
      cleanup();
      if (!email || !password) {
        reject(new Error("BACKEND_AUTH_REQUIRED"));
        return;
      }
      resolve({ email, password });
    });

    window.setTimeout(() => {
      renderEmailOptions(suggestedEmails, primaryEmail);
      emailInput?.focus();
    }, 0);

    const session = getStoredSupabaseSession();
    const sessionToken = String(session?.access_token || "").trim();
    if (sessionToken) {
      fetchLoginEmailSuggestionsFromBackend(sessionToken)
        .then((emails) => {
          if (!emails.length) return;
          const merged = mergeLoginEmailSuggestions(emails);
          renderEmailOptions(merged, primaryEmail);
        })
        .catch(() => null);
    }
  });
}

function isSessionTokenFresh(session) {
  const accessToken = String(session?.access_token || "").trim();
  if (!accessToken) return false;
  const expiresAt = Number(session?.expires_at || 0);
  if (!Number.isFinite(expiresAt) || expiresAt <= 0) return false;
  const nowSec = Math.floor(Date.now() / 1000);
  return expiresAt - nowSec > 45;
}

async function refreshSupabaseSession(refreshToken) {
  const baseUrl = normalizeHttpUrl(SUPABASE_PROJECT_URL);
  const key = String(SUPABASE_PUBLISHABLE_KEY || "").trim();
  const response = await fetch(`${baseUrl}/auth/v1/token?grant_type=refresh_token`, {
    method: "POST",
    headers: {
      apikey: key,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ refresh_token: String(refreshToken || "") }),
    cache: "no-store",
  });
  if (!response.ok) {
    throw new Error("BACKEND_AUTH_REQUIRED");
  }
  const data = await response.json().catch(() => null);
  const accessToken = String(data?.access_token || "").trim();
  if (!accessToken) {
    throw new Error("BACKEND_AUTH_REQUIRED");
  }
  storeSupabaseSession(data);
  return data;
}

async function validateSupabaseAccessToken(accessToken) {
  const token = String(accessToken || "").trim();
  if (!token) return false;
  const baseUrl = normalizeHttpUrl(SUPABASE_PROJECT_URL);
  const key = String(SUPABASE_PUBLISHABLE_KEY || "").trim();
  if (!baseUrl || !key) return false;
  try {
    const response = await fetch(`${baseUrl}/auth/v1/user`, {
      method: "GET",
      headers: {
        apikey: key,
        Authorization: `Bearer ${token}`,
      },
      cache: "no-store",
    });
    return response.ok;
  } catch (error) {
    return false;
  }
}

function isMobileSignaturePageContext() {
  return (
    String(document?.body?.dataset?.page || "") === "mobile-signature" ||
    /signature-mobile\.html$/i.test(String(window.location.pathname || ""))
  );
}

function getUrlSearchParamsSafe() {
  try {
    return new URLSearchParams(String(window.location.search || ""));
  } catch (error) {
    return new URLSearchParams("");
  }
}

function hasMobileSignatureTokenInUrl() {
  const params = getUrlSearchParamsSafe();
  return Boolean(String(params.get("token") || "").trim());
}

function hasSessionBridgeParamsInUrl() {
  const params = getUrlSearchParamsSafe();
  return (
    Boolean(String(params.get("sbat") || "").trim()) ||
    Boolean(String(params.get("sbrt") || "").trim()) ||
    Boolean(String(params.get("sbea") || "").trim())
  );
}

async function getSupabaseUserAccessToken(options = {}) {
  if (getDataBackendMode() === "LOCAL_API") {
    const current = await fetch("/api/local-session", { cache: "no-store" }).catch(() => null);
    if (current?.ok) {
      return "LOCAL_MODE_TOKEN";
    }
    if (options?.requireInteractiveLogin === true || !isMobileSignaturePageContext()) {
      return promptSupabaseLoginAndStoreSession();
    }
    throw new Error("BACKEND_AUTH_REQUIRED");
  }
  const isMobileSignaturePage = isMobileSignaturePageContext();
  const requireInteractiveLogin = options?.requireInteractiveLogin === true;
  if (requireInteractiveLogin) {
    const alreadyValidated = (() => {
      try {
        return sessionStorage.getItem(ADMIN_INTERACTIVE_AUTH_MARKER_KEY) === "1";
      } catch (error) {
        return false;
      }
    })();
    if (!alreadyValidated) {
      if (isMobileSignaturePage) {
        throw new Error("BACKEND_AUTH_REQUIRED");
      }
      return promptSupabaseLoginAndStoreSession();
    }
  }
  const session = getStoredSupabaseSession();
  if (session && isSessionTokenFresh(session)) {
    const token = String(session.access_token || "").trim();
    const isValidToken = await validateSupabaseAccessToken(token);
    if (isValidToken) {
      ensureLoginEventLogged(token).catch(() => null);
      return token;
    }
    clearStoredSupabaseSession();
  }
  const refreshToken = String(session?.refresh_token || "").trim();
  if (refreshToken) {
    try {
      const refreshed = await refreshSupabaseSession(refreshToken);
      const token = String(refreshed?.access_token || "").trim();
      ensureLoginEventLogged(token).catch(() => null);
      return token;
    } catch (error) {
      clearStoredSupabaseSession();
    }
  }
  if (isMobileSignaturePage) {
    throw new Error("BACKEND_AUTH_REQUIRED");
  }
  return promptSupabaseLoginAndStoreSession();
}

async function enforceUiLoginOnEachOpen() {
  if (isPdfRenderMode()) {
    return;
  }
  if (getDataBackendMode() === "LOCAL_API") {
    while (true) {
      try {
        const current = await fetch("/api/local-session", { cache: "no-store" });
        if (current.ok) {
          const payload = await current.json().catch(() => ({}));
          state.currentUserRoleLabel = String(payload?.role_label || "LOCAL_SECOURS");
          return;
        }
        await promptSupabaseLoginAndStoreSession();
        return;
      } catch (error) {
        const message = String(error?.message || "");
        if (message === "BACKEND_AUTH_INVALID") {
          window.alert("IDENTIFIANT OU MOT DE PASSE INCORRECT.");
          continue;
        }
        throw error;
      }
    }
  }
  if (isMobileSignaturePageContext()) {
    return;
  }
  try {
    const alreadyAuthenticatedInThisTab =
      sessionStorage.getItem(ADMIN_INTERACTIVE_AUTH_MARKER_KEY) === "1";
    if (!alreadyAuthenticatedInThisTab) {
      // Wrong credentials should reopen login; cancel should stop opening.
      // Keep looping until success or explicit cancel.
      while (true) {
        try {
          await promptSupabaseLoginAndStoreSession();
          break;
        } catch (error) {
          const message = String(error?.message || "");
          if (message === "BACKEND_AUTH_INVALID") {
            window.alert("IDENTIFIANT OU MOT DE PASSE INCORRECT.");
            continue;
          }
          throw error;
        }
      }
      return;
    }
    const token = await getSupabaseUserAccessToken();
    if (!token) {
      throw new Error("BACKEND_AUTH_REQUIRED");
    }
  } catch (error) {
    try {
      sessionStorage.removeItem(ADMIN_INTERACTIVE_AUTH_MARKER_KEY);
    } catch (markerError) {
      // ignore marker failures
    }
    window.alert("CONNEXION OBLIGATOIRE POUR OUVRIR L'INTERFACE.");
    throw new Error("BACKEND_AUTH_REQUIRED");
  }
}

function renderRoleBadge() {
  let badge = document.getElementById("dotations-role-badge");
  if (!badge) {
    badge = document.createElement("div");
    badge.id = "dotations-role-badge";
    badge.className = "page-header__role-badge";
    const headerActions = document.querySelector(".page-header__actions");
    if (headerActions) {
      headerActions.insertBefore(badge, headerActions.firstChild || null);
    } else {
      document.body.appendChild(badge);
    }
  }
  const label = state.currentUserRoleLabel || "LECTURE";
  badge.textContent = `DROIT : ${label}`;
  const canManageUsers = label === "ADMIN";
  badge.style.cursor = canManageUsers ? "pointer" : "default";
  badge.title = canManageUsers
    ? "Admin: cliquer pour gerer les utilisateurs"
    : "Droit utilisateur";
  badge.onclick = canManageUsers ? openAdminUsersModal : null;
  renderSwitchUserButton();
}

async function handleSwitchUserClick() {
  if (getDataBackendMode() === "LOCAL_API") {
    try {
      await fetch("/api/local-logout", { method: "POST", cache: "no-store" }).catch(() => null);
      while (true) {
        try {
          await promptSupabaseLoginAndStoreSession();
          break;
        } catch (error) {
          const message = String(error?.message || "");
          if (message === "BACKEND_AUTH_INVALID") {
            window.alert("IDENTIFIANT OU MOT DE PASSE INCORRECT.");
            continue;
          }
          throw error;
        }
      }
      showDataStatus("CONNEXION LOCALE MISE A JOUR");
      await reloadData("RECHARGEMENT APRES CHANGEMENT UTILISATEUR...");
    } catch (error) {
      showDataStatus("CHANGEMENT UTILISATEUR IMPOSSIBLE");
    }
    return;
  }
  try {
    clearStoredSupabaseSession();
    while (true) {
      try {
        await promptSupabaseLoginAndStoreSession();
        break;
      } catch (error) {
        const message = String(error?.message || "");
        if (message === "BACKEND_AUTH_INVALID") {
          window.alert("IDENTIFIANT OU MOT DE PASSE INCORRECT.");
          continue;
        }
        throw error;
      }
    }
    showDataStatus("CONNEXION UTILISATEUR MISE A JOUR");
    await reloadData("RECHARGEMENT APRES CHANGEMENT UTILISATEUR...");
  } catch (error) {
    const message = String(error?.message || "");
    if (message !== "BACKEND_AUTH_REQUIRED") {
      window.alert("CHANGEMENT UTILISATEUR IMPOSSIBLE.");
    }
  }
}

function renderSwitchUserButton() {
  if (isMobileSignaturePageContext()) return;
  let button = document.getElementById("dotations-switch-user-button");
  if (!button) {
    button = document.createElement("button");
    button.id = "dotations-switch-user-button";
    button.type = "button";
    button.textContent = "CHANGER D'UTILISATEUR";
    button.className = "button button--secondary page-header__switch-user";
    const headerActions = document.querySelector(".page-header__actions");
    if (headerActions) {
      headerActions.insertBefore(button, headerActions.firstChild || null);
    } else {
      document.body.appendChild(button);
    }
  }
  button.onclick = handleSwitchUserClick;
}

function normalizeRequestedUserRole(role) {
  const value = String(role || "").trim().toLowerCase();
  if (value === "admin") return "admin";
  if (value === "editor" || value === "edition") return "editor";
  return "viewer";
}

function normalizeAdminUserRole(role) {
  return normalizeRequestedUserRole(role);
}

function toOptionalBoolean(value) {
  if (value === true || value === false) return value;
  if (value === 1 || value === 0) return Boolean(value);
  const normalized = String(value || "").trim().toLowerCase();
  if (!normalized) return null;
  if (["true", "1", "yes", "y", "ok", "auth", "valid"].includes(normalized)) return true;
  if (["false", "0", "no", "n", "invalid", "profile_only", "orphan"].includes(normalized)) return false;
  return null;
}

function isAdminUserResettable(entry) {
  if (!entry) return false;
  if (entry.canResetPassword === false) return false;
  if (entry.isAuthUser === false) return false;
  return Boolean(String(entry.email || "").trim());
}

function normalizeAdminUsersList(payload) {
  const candidates = [
    payload?.users,
    payload?.data?.users,
    payload?.items,
    payload?.data?.items,
    payload?.data,
    payload,
  ];
  const list = candidates.find((entry) => Array.isArray(entry));
  if (!Array.isArray(list)) return [];
  return list.map((entry) => {
    const role = normalizeAdminUserRole(entry?.role || entry?.profile?.role || "viewer");
    const isAuthUser = toOptionalBoolean(
      entry?.is_auth_user ?? entry?.isAuthUser ?? entry?.auth_user_exists ?? entry?.authUserExists
    );
    const canResetPassword = toOptionalBoolean(
      entry?.can_reset_password ?? entry?.canResetPassword ?? entry?.resettable ?? entry?.password_reset_enabled
    );
    return {
      id: String(entry?.id || entry?.user_id || "").trim(),
      email: String(entry?.email || "").trim(),
      role,
      createdAt: String(entry?.created_at || entry?.createdAt || "").trim(),
      isAuthUser,
      canResetPassword,
    };
  });
}

function formatAdminUserDate(value) {
  const raw = String(value || "").trim();
  if (!raw) return "-";
  const date = new Date(raw);
  if (!Number.isFinite(date.getTime())) return raw;
  return date.toLocaleString("fr-FR");
}

function toLocalDayKey(value) {
  const date = value instanceof Date ? value : new Date(value);
  if (!Number.isFinite(date.getTime())) return "";
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

const ADMIN_ARCHIVED_USERS_KEY = "dotations_admin_archived_users";
function getAdminArchivedUserIds() {
  try {
    const raw = window.localStorage.getItem(ADMIN_ARCHIVED_USERS_KEY);
    const list = JSON.parse(raw || "[]");
    if (!Array.isArray(list)) return [];
    return list.map((id) => String(id || "").trim()).filter(Boolean);
  } catch {
    return [];
  }
}
function setAdminArchivedUserIds(ids = []) {
  try {
    const next = Array.from(new Set((ids || []).map((id) => String(id || "").trim()).filter(Boolean)));
    window.localStorage.setItem(ADMIN_ARCHIVED_USERS_KEY, JSON.stringify(next));
  } catch {}
}

async function openAdminUsersModal() {
  if (state.currentUserRoleLabel !== "ADMIN") {
    window.alert("ACCES REFUSE : COMPTE ADMIN REQUIS.");
    return;
  }
  if (document.getElementById("admin-users-modal")) {
    return;
  }

  const backdrop = document.createElement("div");
  backdrop.id = "admin-users-modal";
  backdrop.style.position = "fixed";
  backdrop.style.inset = "0";
  backdrop.style.background = "rgba(12,25,33,0.45)";
  backdrop.style.zIndex = "99999";
  backdrop.style.display = "flex";
  backdrop.style.alignItems = "center";
  backdrop.style.justifyContent = "center";
  backdrop.innerHTML = `
    <div id="admin-users-dialog" style="position:fixed;left:50%;top:50%;transform:translate(-50%,-50%);width:min(94vw,920px);max-height:88vh;overflow:auto;background:#fff;border-radius:12px;border:1px solid #9bb2be;box-shadow:0 18px 36px rgba(0,0,0,0.28);padding:14px;">
      <div id="admin-users-drag-handle" style="display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:10px;cursor:move;user-select:none;">
        <div>
          <div style="font-size:16px;font-weight:700;color:#122430;">PILOTAGE UTILISATEURS</div>
          <div style="font-size:11px;color:#4a6170;">ACCES ADMIN UNIQUEMENT</div>
        </div>
        <button type="button" id="admin-users-close" style="height:30px;padding:0 10px;border:1px solid #9bb2be;background:#fff;border-radius:8px;cursor:pointer;">Fermer</button>
      </div>
      <div id="admin-users-status" style="font-size:12px;color:#355464;margin-bottom:8px;"></div>
      <div id="admin-users-list-wrap" style="border:1px solid #d4e0e6;border-radius:10px;overflow:hidden;margin-bottom:10px;"></div>
      <form id="admin-users-create-form" style="display:grid;grid-template-columns:2fr 1.4fr 1fr auto auto;gap:8px;align-items:end;margin-bottom:10px;">
        <label style="display:block;font-size:11px;color:#3c5561;">Email
          <input name="email" type="email" required autocomplete="username email" style="width:100%;height:34px;padding:0 10px;border:1px solid #9bb2be;border-radius:8px;" />
        </label>
        <label style="display:block;font-size:11px;color:#3c5561;">Mot de passe temporaire (creation manuelle)
          <input name="password" type="text" autocomplete="new-password" style="width:100%;height:34px;padding:0 10px;border:1px solid #9bb2be;border-radius:8px;" />
        </label>
        <label style="display:block;font-size:11px;color:#3c5561;">Role
          <select name="role" style="width:100%;height:34px;padding:0 10px;border:1px solid #9bb2be;border-radius:8px;">
            <option value="viewer">LECTURE</option>
            <option value="editor">EDITEUR</option>
            <option value="admin">ADMIN</option>
          </select>
        </label>
        <button type="submit" data-submit-action="create" style="height:34px;padding:0 12px;border:0;background:#2f5f76;color:#fff;border-radius:8px;cursor:pointer;font-weight:700;">Creer</button>
        <button type="submit" data-submit-action="invite" style="height:34px;padding:0 12px;border:1px solid #9bb2be;background:#fff;color:#1d3440;border-radius:8px;cursor:pointer;font-weight:700;">Envoyer invitation</button>
      </form>
      <div id="admin-restore-wrap" style="border:1px solid #d4e0e6;border-radius:10px;overflow:hidden;margin-bottom:10px;"></div>
      <div style="margin:8px 0 6px;font-size:12px;font-weight:700;color:#1d3440;">JOURNAL DES CONNEXIONS</div>
      <div id="admin-logins-status" style="font-size:12px;color:#355464;margin-bottom:8px;"></div>
      <div id="admin-logins-heatmap" style="border:1px solid #d4e0e6;border-radius:10px;overflow:auto;margin-bottom:8px;"></div>
      <div id="admin-logins-summary" style="border:1px solid #d4e0e6;border-radius:10px;overflow:hidden;margin-bottom:8px;"></div>
      <div id="admin-logins-list-wrap" style="border:1px solid #d4e0e6;border-radius:10px;overflow:hidden;margin-bottom:10px;max-height:210px;overflow-y:auto;"></div>
    </div>
  `;
  document.body.appendChild(backdrop);

  const closeButton = backdrop.querySelector("#admin-users-close");
  const dialog = backdrop.querySelector("#admin-users-dialog");
  const dragHandle = backdrop.querySelector("#admin-users-drag-handle");
  const statusNode = backdrop.querySelector("#admin-users-status");
  const loginsStatusNode = backdrop.querySelector("#admin-logins-status");
  const loginsHeatmapNode = backdrop.querySelector("#admin-logins-heatmap");
  const loginsSummaryNode = backdrop.querySelector("#admin-logins-summary");
  const loginsListWrap = backdrop.querySelector("#admin-logins-list-wrap");
  const listWrap = backdrop.querySelector("#admin-users-list-wrap");
  const restoreWrap = backdrop.querySelector("#admin-restore-wrap");
  const createForm = backdrop.querySelector("#admin-users-create-form");

  const setStatus = (message, variant = "info") => {
    if (!statusNode) return;
    statusNode.textContent = String(message || "");
    statusNode.style.color =
      variant === "error" ? "#8e2c2c" : variant === "ok" ? "#2f5e43" : "#355464";
  };
  const setLoginsStatus = (message, variant = "info") => {
    if (!loginsStatusNode) return;
    loginsStatusNode.textContent = String(message || "");
    loginsStatusNode.style.color =
      variant === "error" ? "#8e2c2c" : variant === "ok" ? "#2f5e43" : "#355464";
  };

  const closeModal = () => {
    backdrop.remove();
  };

  closeButton?.addEventListener("click", closeModal);
  backdrop.addEventListener("click", (event) => {
    if (event.target === backdrop) closeModal();
  });

  let dragState = null;
  const onDragMove = (event) => {
    if (!dragState || !dialog) return;
    const nextLeft = event.clientX - dragState.offsetX;
    const nextTop = event.clientY - dragState.offsetY;
    dialog.style.left = `${Math.max(8, nextLeft)}px`;
    dialog.style.top = `${Math.max(8, nextTop)}px`;
    dialog.style.transform = "none";
  };
  const onDragEnd = () => {
    if (!dragState) return;
    dragState = null;
    window.removeEventListener("mousemove", onDragMove);
    window.removeEventListener("mouseup", onDragEnd);
  };
  dragHandle?.addEventListener("mousedown", (event) => {
    if (!dialog) return;
    const target = event.target;
    if (target instanceof HTMLElement && target.closest("button,input,select,textarea,a")) {
      return;
    }
    const rect = dialog.getBoundingClientRect();
    dragState = {
      offsetX: event.clientX - rect.left,
      offsetY: event.clientY - rect.top,
    };
    window.addEventListener("mousemove", onDragMove);
    window.addEventListener("mouseup", onDragEnd);
  });

  let users = [];
  let loginEvents = [];
  let appStateVersions = [];
  const isLocalAdminMode = getDataBackendMode() === "LOCAL_API";
  let archivedUserIds = new Set(getAdminArchivedUserIds());
  const getArchivedPeople = () => (state.data?.personnes || []).filter((person) => isSoftDeletedEntity(person));
  const getArchivedEffects = () =>
    (state.data?.personnes || []).flatMap((person) =>
      (person?.effetsConfies || []).filter((effect) => isSoftDeletedEntity(effect))
    );
  const renderRestorePanel = () => {
    if (!restoreWrap) return;
    const archivedPeople = getArchivedPeople();
    const archivedEffects = getArchivedEffects();
    const archivedUsers = users.filter((user) => archivedUserIds.has(String(user.id || "")));
    const appStateOptions = appStateVersions
      .map((entry, index) => {
        const appStateId = String(entry?.app_state_id || "main");
        const sourceRevision = Number(entry?.source_revision || 0);
        const createdAt = formatAdminUserDate(entry?.created_at || "");
        const reason = String(entry?.reason || "snapshot_before_update");
        const value = `${appStateId}|${sourceRevision}`;
        const label = `${createdAt} | rev ${sourceRevision} | ${reason}`;
        return `<option value="${escapeHtml(value)}"${index === 0 ? " selected" : ""}>${escapeHtml(label)}</option>`;
      })
      .join("");
    restoreWrap.innerHTML = `
      <div style="padding:8px;background:#f3f7f9;border-bottom:1px solid #e2ebef;font-size:11px;color:#3d5865;font-weight:700;">ARCHIVES ADMIN (REINJECTION)</div>
      <div style="padding:8px;display:grid;grid-template-columns:1fr auto;gap:8px;align-items:center;border-bottom:1px solid #e2ebef;">
        <div style="font-size:12px;color:#1c3440;">Personnes archivees: <b>${escapeHtml(String(archivedPeople.length))}</b></div>
        <button type="button" data-restore-action="people" style="height:30px;padding:0 10px;border:1px solid #9bb2be;background:#fff;border-radius:7px;cursor:pointer;">Restaurer personnes</button>
      </div>
      <div style="padding:8px;display:grid;grid-template-columns:1fr auto;gap:8px;align-items:center;border-bottom:1px solid #e2ebef;">
        <div style="font-size:12px;color:#1c3440;">Effets archives: <b>${escapeHtml(String(archivedEffects.length))}</b></div>
        <button type="button" data-restore-action="effects" style="height:30px;padding:0 10px;border:1px solid #9bb2be;background:#fff;border-radius:7px;cursor:pointer;">Restaurer effets</button>
      </div>
      <div style="padding:8px;display:grid;grid-template-columns:1fr auto;gap:8px;align-items:center;border-bottom:1px solid #e2ebef;">
        <div style="font-size:12px;color:#1c3440;">Utilisateurs archives: <b>${escapeHtml(String(archivedUsers.length))}</b></div>
        <button type="button" data-restore-action="users" style="height:30px;padding:0 10px;border:1px solid #9bb2be;background:#fff;border-radius:7px;cursor:pointer;">Restaurer utilisateurs</button>
      </div>
      <div style="padding:8px;display:flex;justify-content:flex-end;">
        <button type="button" data-restore-action="all" style="height:32px;padding:0 12px;border:0;background:#2f5f76;color:#fff;border-radius:8px;cursor:pointer;font-weight:700;">Restaurer tout</button>
      </div>
      <div style="padding:8px;background:#f3f7f9;border-top:1px solid #e2ebef;border-bottom:1px solid #e2ebef;font-size:11px;color:#3d5865;font-weight:700;">COFFRE-FORT DONNEES (APP_STATE)</div>
      <div style="padding:8px;display:grid;grid-template-columns:1fr auto;gap:8px;align-items:center;">
        <select id="admin-app-state-version-select" style="height:32px;padding:0 8px;border:1px solid #9bb2be;border-radius:7px;">
          ${appStateOptions || '<option value="">AUCUNE VERSION DISPONIBLE</option>'}
        </select>
        <button type="button" data-restore-action="app-state" style="height:32px;padding:0 12px;border:0;background:#2f5f76;color:#fff;border-radius:8px;cursor:pointer;font-weight:700;" ${appStateVersions.length ? "" : "disabled"}>Restaurer version</button>
      </div>
    `;
  };
  const renderUsers = () => {
    if (!listWrap) return;
    const activeUsers = users.filter((user) => !archivedUserIds.has(String(user.id || "")));
    const rows = activeUsers
      .map(
        (user) => `
          <tr data-user-id="${escapeHtml(user.id)}">
            <td style="padding:8px;border-bottom:1px solid #e2ebef;font-size:12px;color:#1c3440;">${escapeHtml(user.email || "-")}</td>
            <td style="padding:8px;border-bottom:1px solid #e2ebef;font-size:12px;color:#1c3440;">${escapeHtml(formatAdminUserDate(user.createdAt))}</td>
            <td style="padding:8px;border-bottom:1px solid #e2ebef;font-size:11px;color:#1c3440;">
              ${
                user.canResetPassword === false || user.isAuthUser === false
                  ? "NON RESETTABLE"
                  : user.canResetPassword === true || user.isAuthUser === true
                    ? "AUTH OK"
                    : "INCONNU"
              }
            </td>
            <td style="padding:8px;border-bottom:1px solid #e2ebef;">
              <select data-action="role" style="height:30px;padding:0 8px;border:1px solid #9bb2be;border-radius:7px;">
                <option value="viewer"${user.role === "viewer" ? " selected" : ""}>LECTURE</option>
                <option value="editor"${user.role === "editor" ? " selected" : ""}>EDITEUR</option>
                <option value="admin"${user.role === "admin" ? " selected" : ""}>ADMIN</option>
              </select>
            </td>
            <td style="padding:8px;border-bottom:1px solid #e2ebef;white-space:nowrap;">
              <button type="button" data-action="save" style="height:30px;padding:0 10px;border:1px solid #9bb2be;background:#fff;border-radius:7px;cursor:pointer;">Modifier</button>
              <button type="button" data-action="delete" style="height:30px;padding:0 10px;border:1px solid #d7a6a6;background:#fff5f5;color:#8e2c2c;border-radius:7px;cursor:pointer;margin-left:6px;">Supprimer</button>
            </td>
          </tr>`
      )
      .join("");
    listWrap.innerHTML = `
      <table style="width:100%;border-collapse:collapse;">
        <thead>
          <tr style="background:#f3f7f9;">
            <th style="text-align:left;padding:8px;font-size:11px;color:#3d5865;">EMAIL</th>
            <th style="text-align:left;padding:8px;font-size:11px;color:#3d5865;">CREATION</th>
            <th style="text-align:left;padding:8px;font-size:11px;color:#3d5865;">AUTH</th>
            <th style="text-align:left;padding:8px;font-size:11px;color:#3d5865;">ROLE</th>
            <th style="text-align:left;padding:8px;font-size:11px;color:#3d5865;">ACTION</th>
          </tr>
        </thead>
        <tbody>${rows || `<tr><td colspan="5" style="padding:12px;text-align:center;color:#4a6170;font-size:12px;">AUCUN UTILISATEUR</td></tr>`}</tbody>
      </table>
    `;
  };

  const fetchUsers = async () => {
    setStatus("Chargement utilisateurs...");
    try {
      if (isLocalAdminMode) {
        const response = await fetch("/api/local-users", {
          method: "GET",
          credentials: "same-origin",
          cache: "no-store",
        });
        if (!response.ok) {
          const details = await response.text().catch(() => "");
          throw new Error(`LOCAL_USERS_LOAD_FAILED:${response.status}:${details}`);
        }
        const payload = await response.json().catch(() => ({}));
        users = normalizeAdminUsersList(payload?.users || []);
      } else {
        const response = await callEdgeApi("admin/users", { method: "GET" });
        const payload = await response.json().catch(() => ({}));
        users = normalizeAdminUsersList(payload);
      }
      mergeLoginEmailSuggestions(users.map((entry) => entry?.email || ""));
      renderUsers();
      const visibleCount = users.filter((user) => !archivedUserIds.has(String(user.id || ""))).length;
      setStatus(`${visibleCount} utilisateur(s) charge(s).`, "ok");
      renderRestorePanel();
    } catch (error) {
      const message = String(error?.message || "");
      if (message.includes("EDGE_API_FAILED:404")) {
        setStatus("API ADMIN ABSENTE: ajouter /admin/users dans dotations-api.", "error");
      } else {
        setStatus(`Chargement impossible: ${message || "erreur"}`, "error");
      }
      users = [];
      renderUsers();
      renderRestorePanel();
    }
  };
  const fetchAppStateVersions = async () => {
    if (isLocalAdminMode) {
      appStateVersions = [];
      setStatus("Mode local : historique Supabase indisponible", "info");
      renderRestorePanel();
      return;
    }
    const loadVersions = async (accessToken) => {
      const endpoint = `${getSupabaseRestEndpoint().replace(/\/app_state$/i, "/app_state_versions")}?select=app_state_id,source_revision,created_at,reason&order=created_at.desc&limit=20`;
      return fetch(endpoint, {
        method: "GET",
        headers: getSupabaseHeaders(
          {
            Authorization: `Bearer ${accessToken}`,
          },
          { includeAuthorization: false },
        ),
        cache: "no-store",
      });
    };
    try {
      let accessToken = await getSupabaseUserAccessToken({ requireInteractiveLogin: true });
      if (!accessToken) {
        throw new Error("BACKEND_AUTH_REQUIRED");
      }
      let response = await loadVersions(accessToken);
      if (response.status === 401 || response.status === 403) {
        clearStoredSupabaseSession();
        accessToken = await getSupabaseUserAccessToken({ requireInteractiveLogin: true });
        if (!accessToken) {
          throw new Error("BACKEND_AUTH_REQUIRED");
        }
        response = await loadVersions(accessToken);
      }
      if (!response.ok) {
        const details = await response.text().catch(() => "");
        throw new Error(`VERSIONS_LOAD_FAILED:${response.status}:${details}`);
      }
      const rows = await response.json().catch(() => []);
      appStateVersions = Array.isArray(rows) ? rows : [];
      if (!appStateVersions.length) {
        setStatus("Historique versions vide (0 ligne retournee).", "error");
      }
    } catch (error) {
      appStateVersions = [];
      const message = String(error?.message || "erreur");
      setStatus(`Historique versions indisponible: ${message}`, "error");
    } finally {
      renderRestorePanel();
    }
  };
  const normalizeLoginsList = (payload) => {
    const candidates = [
      payload?.events,
      payload?.data?.events,
      payload?.items,
      payload?.data?.items,
      payload?.data,
      payload,
    ];
    const list = candidates.find((entry) => Array.isArray(entry));
    if (!Array.isArray(list)) return [];
    return list.map((entry) => ({
      at: String(entry?.at || entry?.created_at || entry?.date || "").trim(),
      name: String(entry?.name || entry?.nom || "").trim(),
      email: String(entry?.email || "").trim(),
      role: normalizeAdminUserRole(entry?.role || "viewer"),
      count: Number(entry?.count || entry?.nb || 1) || 1,
    }));
  };
  const renderLoginEvents = () => {
    if (!loginsSummaryNode || !loginsListWrap || !loginsHeatmapNode) return;
    if (!Array.isArray(loginEvents) || !loginEvents.length) {
      loginsHeatmapNode.innerHTML =
        `<div style="padding:10px;font-size:12px;color:#4a6170;text-align:center;">AUCUN HISTORIQUE</div>`;
      loginsSummaryNode.innerHTML =
        `<div style="padding:10px;font-size:12px;color:#4a6170;text-align:center;">AUCUN EVENEMENT</div>`;
      loginsListWrap.innerHTML =
        `<div style="padding:10px;font-size:12px;color:#4a6170;text-align:center;">AUCUN DETAIL</div>`;
      return;
    }
    const recentDays = 30;
    const dayKeys = [];
    const dayLabels = [];
    for (let offset = recentDays - 1; offset >= 0; offset -= 1) {
      const d = new Date();
      d.setHours(0, 0, 0, 0);
      d.setDate(d.getDate() - offset);
      dayKeys.push(toLocalDayKey(d));
      dayLabels.push(`${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}`);
    }
    const byEmail = new Map();
    loginEvents.forEach((event) => {
      const key = String(event.email || "-").toUpperCase();
      const prev = byEmail.get(key) || { email: key, total: 0, role: event.role };
      prev.total += Number(event.count || 1);
      prev.role = event.role || prev.role;
      const day = toLocalDayKey(event.at) || String(event.at || "").slice(0, 10);
      if (!prev.byDay) prev.byDay = {};
      prev.byDay[day] = (Number(prev.byDay[day] || 0) + Number(event.count || 1));
      byEmail.set(key, prev);
    });
    const heatRows = Array.from(byEmail.values())
      .sort((a, b) => b.total - a.total || String(a.email).localeCompare(String(b.email), "fr"))
      .map((row) => {
        const cells = dayKeys
          .map((day) => {
            const value = Number(row.byDay?.[day] || 0);
            let bg = "#d7e1ea";
            let fg = "#34515f";
            if (value >= 8) {
              bg = "#0b7a47";
              fg = "#ffffff";
            } else if (value >= 5) {
              bg = "#1f9b5e";
              fg = "#ffffff";
            } else if (value >= 3) {
              bg = "#46b979";
              fg = "#ffffff";
            } else if (value >= 2) {
              bg = "#7fd09f";
              fg = "#113a24";
            } else if (value >= 1) {
              bg = "#b8e5c9";
              fg = "#1d5a39";
            }
            return `<span title="${escapeHtml(day)} : ${value}" style="display:inline-flex;align-items:center;justify-content:center;width:14px;height:14px;border-radius:3px;background:${bg};border:1px solid rgba(55,85,101,0.2);font-size:9px;font-weight:700;color:${fg};line-height:1;">${value > 0 ? escapeHtml(String(value)) : ""}</span>`;
          })
          .join("");
        return `<div style="display:grid;grid-template-columns:220px 1fr auto;gap:8px;align-items:center;padding:6px 8px;border-bottom:1px solid #e2ebef;">
          <div style="font-size:11px;color:#1d3440;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${escapeHtml(row.email)}</div>
          <div style="display:grid;grid-template-columns:repeat(${recentDays}, 14px);gap:3px;min-width:max-content;">${cells}</div>
          <div style="font-size:11px;color:#3d5865;font-weight:700;">${escapeHtml(String(row.total))}</div>
        </div>`;
      })
      .join("");
    loginsHeatmapNode.innerHTML = `
      <div style="padding:6px 8px;background:#f3f7f9;border-bottom:1px solid #e2ebef;font-size:11px;color:#3d5865;font-weight:700;">
        ACTIVITE CONNEXIONS - 30 JOURS
      </div>
      <div style="display:grid;grid-template-columns:220px 1fr auto;gap:8px;align-items:end;padding:4px 8px;border-bottom:1px solid #e2ebef;">
        <div></div>
        <div style="display:grid;grid-template-columns:repeat(${recentDays}, 14px);gap:3px;min-width:max-content;">
          ${dayLabels
            .map(
              (label) => `<span style="width:14px;height:16px;display:flex;align-items:flex-end;justify-content:center;overflow:visible;" title="${escapeHtml(label)}">
            <span style="font-size:8px;color:#5a7381;transform:rotate(-65deg);transform-origin:center bottom;display:block;line-height:1;white-space:nowrap;">${escapeHtml(label.slice(0, 2))}</span>
          </span>`
            )
            .join("")}
        </div>
        <div></div>
      </div>
      <div>${heatRows}</div>
      <div style="display:flex;justify-content:flex-end;align-items:center;gap:6px;padding:6px 8px;border-top:1px solid #e2ebef;">
        <span style="font-size:10px;color:#4a6170;">MOINS</span>
        <span style="display:inline-block;width:14px;height:14px;border-radius:3px;background:#d7e1ea;border:1px solid rgba(55,85,101,0.2);"></span>
        <span style="display:inline-block;width:14px;height:14px;border-radius:3px;background:#b8e5c9;border:1px solid rgba(55,85,101,0.2);"></span>
        <span style="display:inline-block;width:14px;height:14px;border-radius:3px;background:#7fd09f;border:1px solid rgba(55,85,101,0.2);"></span>
        <span style="display:inline-block;width:14px;height:14px;border-radius:3px;background:#46b979;border:1px solid rgba(55,85,101,0.2);"></span>
        <span style="display:inline-block;width:14px;height:14px;border-radius:3px;background:#1f9b5e;border:1px solid rgba(55,85,101,0.2);"></span>
        <span style="display:inline-block;width:14px;height:14px;border-radius:3px;background:#0b7a47;border:1px solid rgba(55,85,101,0.2);"></span>
        <span style="font-size:10px;color:#4a6170;">PLUS</span>
      </div>
    `;
    const summaryRows = Array.from(byEmail.values())
      .sort((a, b) => b.total - a.total || String(a.email).localeCompare(String(b.email), "fr"))
      .map((row) => {
        return `<tr>
          <td style="padding:6px 8px;border-bottom:1px solid #e2ebef;font-size:12px;">${escapeHtml(row.email)}</td>
          <td style="padding:6px 8px;border-bottom:1px solid #e2ebef;font-size:12px;">${escapeHtml(mapRoleToFrenchLabel(row.role))}</td>
          <td style="padding:6px 8px;border-bottom:1px solid #e2ebef;font-size:12px;text-align:right;">${escapeHtml(String(row.total))}</td>
        </tr>`;
      })
      .join("");
    loginsSummaryNode.innerHTML = `
      <table style="width:100%;border-collapse:collapse;">
        <thead><tr style="background:#f3f7f9;">
          <th style="text-align:left;padding:6px 8px;font-size:11px;color:#3d5865;">EMAIL</th>
          <th style="text-align:left;padding:6px 8px;font-size:11px;color:#3d5865;">ROLE</th>
          <th style="text-align:right;padding:6px 8px;font-size:11px;color:#3d5865;">NB</th>
        </tr></thead>
        <tbody>${summaryRows}</tbody>
      </table>
    `;
    const detailRows = loginEvents
      .slice()
      .sort((a, b) => String(b.at).localeCompare(String(a.at)))
      .slice(0, 100)
      .map((row) => {
        return `<tr>
          <td style="padding:6px 8px;border-bottom:1px solid #e2ebef;font-size:12px;">${escapeHtml(formatAdminUserDate(row.at))}</td>
          <td style="padding:6px 8px;border-bottom:1px solid #e2ebef;font-size:12px;">${escapeHtml(row.name || "-")}</td>
          <td style="padding:6px 8px;border-bottom:1px solid #e2ebef;font-size:12px;">${escapeHtml(row.email || "-")}</td>
          <td style="padding:6px 8px;border-bottom:1px solid #e2ebef;font-size:12px;">${escapeHtml(mapRoleToFrenchLabel(row.role))}</td>
          <td style="padding:6px 8px;border-bottom:1px solid #e2ebef;font-size:12px;text-align:right;">${escapeHtml(String(row.count || 1))}</td>
        </tr>`;
      })
      .join("");
    loginsListWrap.innerHTML = `
      <table style="width:100%;border-collapse:collapse;">
        <thead><tr style="background:#f3f7f9;position:sticky;top:0;">
          <th style="text-align:left;padding:6px 8px;font-size:11px;color:#3d5865;">DATE</th>
          <th style="text-align:left;padding:6px 8px;font-size:11px;color:#3d5865;">NOM</th>
          <th style="text-align:left;padding:6px 8px;font-size:11px;color:#3d5865;">EMAIL</th>
          <th style="text-align:left;padding:6px 8px;font-size:11px;color:#3d5865;">ROLE</th>
          <th style="text-align:right;padding:6px 8px;font-size:11px;color:#3d5865;">NB</th>
        </tr></thead>
        <tbody>${detailRows}</tbody>
      </table>
    `;
  };
  const fetchLoginEvents = async () => {
    if (isLocalAdminMode) {
      loginEvents = [];
      renderLoginEvents();
      setLoginsStatus("Mode local : historique Supabase indisponible", "info");
      return;
    }
    setLoginsStatus("Chargement journal connexions...");
    try {
      const response = await callEdgeApi("admin/login-events", { method: "GET" });
      const payload = await response.json().catch(() => ({}));
      loginEvents = normalizeLoginsList(payload);
      renderLoginEvents();
      setLoginsStatus(`${loginEvents.length} evenement(s) charge(s).`, "ok");
    } catch (error) {
      const message = String(error?.message || "");
      if (message.includes("EDGE_API_FAILED:404")) {
        setLoginsStatus("API JOURNAL ABSENTE: ajouter /admin/login-events dans dotations-api.", "error");
      } else {
        setLoginsStatus(`Chargement journal impossible: ${message || "erreur"}`, "error");
      }
      loginEvents = [];
      renderLoginEvents();
    }
  };

  listWrap?.addEventListener("click", async (event) => {
    const button = event.target?.closest?.("button[data-action]");
    if (!button) return;
    const row = button.closest("tr[data-user-id]");
    const userId = String(row?.getAttribute("data-user-id") || "").trim();
    if (!userId) return;
    const roleSelect = row.querySelector('select[data-action="role"]');
    const role = normalizeRequestedUserRole(roleSelect?.value || "viewer");
    const action = String(button.getAttribute("data-action") || "");
    try {
      if (action === "save") {
        setStatus("Mise a jour utilisateur...");
        if (isLocalAdminMode) {
          const response = await fetch("/api/local-users/update", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            credentials: "same-origin",
            body: JSON.stringify({ userId, role }),
          });
          if (!response.ok) {
            const details = await response.text().catch(() => "");
            throw new Error(`LOCAL_USERS_UPDATE_FAILED:${response.status}:${details}`);
          }
        } else {
          await callEdgeApi(`admin/users/${encodeURIComponent(userId)}`, {
            method: "PATCH",
            body: JSON.stringify({ role }),
          });
        }
        setStatus("Role utilisateur mis a jour.", "ok");
        await fetchUsers();
        return;
      }
      if (action === "delete") {
        if (!window.confirm(isLocalAdminMode ? "Desactiver cet utilisateur local ?" : "Archiver cet utilisateur de la vue admin ?")) return;
        if (isLocalAdminMode) {
          const response = await fetch("/api/local-users/deactivate", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            credentials: "same-origin",
            body: JSON.stringify({ userId }),
          });
          if (!response.ok) {
            const details = await response.text().catch(() => "");
            throw new Error(`LOCAL_USERS_DEACTIVATE_FAILED:${response.status}:${details}`);
          }
          setStatus("Utilisateur local desactive.", "ok");
          await fetchUsers();
        } else {
          archivedUserIds.add(userId);
          setAdminArchivedUserIds(Array.from(archivedUserIds));
          setStatus("Utilisateur archive (masque de la liste).", "ok");
          renderUsers();
          renderRestorePanel();
        }
      }
    } catch (error) {
      setStatus(`Action impossible: ${String(error?.message || "erreur")}`, "error");
    }
  });

  restoreWrap?.addEventListener("click", async (event) => {
    const button = event.target?.closest?.("button[data-restore-action]");
    if (!button) return;
    const action = String(button.getAttribute("data-restore-action") || "");
    try {
      if (action === "users") {
        archivedUserIds = new Set();
        setAdminArchivedUserIds([]);
        renderUsers();
        renderRestorePanel();
        setStatus("Utilisateurs reinjectes dans la liste admin.", "ok");
        return;
      }
      if (action === "app-state") {
        const selectNode = restoreWrap.querySelector("#admin-app-state-version-select");
        const rawValue = String(selectNode?.value || "").trim();
        if (!rawValue || !rawValue.includes("|")) {
          setStatus("Aucune version selectionnee.", "error");
          return;
        }
        const [appStateId, sourceRevisionRaw] = rawValue.split("|");
        const sourceRevision = Number.parseInt(String(sourceRevisionRaw || ""), 10);
        if (!appStateId || !Number.isFinite(sourceRevision)) {
          setStatus("Version invalide.", "error");
          return;
        }
        if (!window.confirm(`Restaurer app_state ${appStateId} revision ${sourceRevision} ?`)) {
          return;
        }
        const accessToken = await getSupabaseUserAccessToken({ requireInteractiveLogin: true });
        if (!accessToken) {
          throw new Error("BACKEND_AUTH_REQUIRED");
        }
        setStatus("Restauration version en cours...");
        const rpcEndpoint = getSupabaseRestEndpoint().replace(/\/app_state$/i, "/rpc/restore_app_state_revision");
        const rpcResponse = await fetch(rpcEndpoint, {
          method: "POST",
          headers: getSupabaseHeaders(
            {
              "Content-Type": "application/json",
              Authorization: `Bearer ${accessToken}`,
            },
            { includeAuthorization: false },
          ),
          body: JSON.stringify({
            _app_state_id: appStateId,
            _source_revision: sourceRevision,
            _reason: "restore_ui_admin",
          }),
          cache: "no-store",
        });
        if (!rpcResponse.ok) {
          const details = await rpcResponse.text().catch(() => "");
          throw new Error(`RESTORE_VERSION_FAILED:${rpcResponse.status}:${details}`);
        }
        await reloadData("RECHARGEMENT APRES RESTAURATION...");
        await fetchAppStateVersions();
        setStatus(`Restauration appliquee (revision source ${sourceRevision}).`, "ok");
        return;
      }
      const people = state.data?.personnes || [];
      let changed = false;
      if (action === "people" || action === "all") {
        people.forEach((person) => {
          if (isSoftDeletedEntity(person)) {
            person.is_deleted = false;
            person.isDeleted = false;
            delete person.deleted_at;
            delete person.deletedAt;
            changed = true;
          }
        });
      }
      if (action === "effects" || action === "all") {
        people.forEach((person) => {
          (person?.effetsConfies || []).forEach((effect) => {
            if (isSoftDeletedEntity(effect)) {
              effect.is_deleted = false;
              effect.isDeleted = false;
              delete effect.deleted_at;
              delete effect.deletedAt;
              changed = true;
            }
          });
        });
      }
      if (action === "all") {
        archivedUserIds = new Set();
        setAdminArchivedUserIds([]);
      }
      if (!changed && action !== "all") {
        setStatus("Aucun element a restaurer.", "info");
        renderRestorePanel();
        return;
      }
      if (changed) {
      markDirty();
      schedulePageRender();
        await saveDataToFile({
          silent: true,
          reloadAfter: true,
          successText: "RESTAURATION ARCHIVES - SAUVEGARDE OK",
        });
      }
      renderUsers();
      renderRestorePanel();
      setStatus("Restauration terminee.", "ok");
    } catch (error) {
      setStatus(`Restauration impossible: ${String(error?.message || "erreur")}`, "error");
    }
  });

  createForm?.addEventListener("submit", async (event) => {
    event.preventDefault();
    const formData = new FormData(createForm);
    const email = String(formData.get("email") || "").trim();
    const password = String(formData.get("password") || "");
    const role = normalizeRequestedUserRole(formData.get("role") || "viewer");
    const submitter = event.submitter instanceof HTMLElement ? event.submitter : null;
    const submitAction = String(submitter?.getAttribute("data-submit-action") || "create");
    if (!email) {
      setStatus("Email obligatoire.", "error");
      return;
    }
    if (submitAction === "create" && !password) {
      setStatus("Mot de passe obligatoire pour creation manuelle.", "error");
      return;
    }
    const generatedPassword = `Tmp#${Math.random().toString(36).slice(2, 10)}A1`;
    const passwordToUse = password || generatedPassword;
    const isInviteFlow = submitAction === "invite";
    if (isLocalAdminMode && isInviteFlow) {
      setStatus("Mode local : invitation Supabase indisponible.", "info");
      return;
    }
    if (isInviteFlow) {
      const now = Date.now();
      const lastInviteAt = Number.parseInt(String(localStorage.getItem(ADMIN_INVITE_COOLDOWN_KEY) || ""), 10);
      if (Number.isFinite(lastInviteAt)) {
        const remainingMs = ADMIN_INVITE_COOLDOWN_MS - (now - lastInviteAt);
        if (remainingMs > 0) {
          const remainingSec = Math.ceil(remainingMs / 1000);
          setStatus(`Invitation recente: attends ${remainingSec}s avant un nouvel envoi.`, "error");
          return;
        }
      }
    }
    setStatus(isInviteFlow ? "Envoi invitation..." : "Creation utilisateur...");
    try {
      let userAlreadyExists = false;
      if (isInviteFlow) {
        try {
          await requestAdminUserInvitation(email, role);
        } catch (error) {
          const message = String(error?.message || "");
          if (message === "INVITE_USER_EXISTS") {
            userAlreadyExists = true;
          } else {
            throw error;
          }
        }
      } else {
        if (isLocalAdminMode) {
          const response = await fetch("/api/local-users/create", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            credentials: "same-origin",
            body: JSON.stringify({ email, password: passwordToUse, role }),
          });
          if (!response.ok) {
            const details = await response.text().catch(() => "");
            throw new Error(`LOCAL_USERS_CREATE_FAILED:${response.status}:${details}`);
          }
        } else {
          await callEdgeApi("admin/users", {
            method: "POST",
            body: JSON.stringify({ email, password: passwordToUse, role }),
          });
        }
      }
      if (isInviteFlow) {
        localStorage.setItem(ADMIN_INVITE_COOLDOWN_KEY, String(Date.now()));
      }
      if (isInviteFlow && userAlreadyExists) {
        setStatus("Compte deja existant: invitation renvoyee.", "ok");
      } else {
        setStatus(isInviteFlow ? "Invitation envoyee." : "Utilisateur cree.", "ok");
      }
      createForm.reset();
      await fetchUsers();
    } catch (error) {
      const message = String(error?.message || "erreur");
      if (message === "INVITE_ENDPOINT_ABSENT") {
        setStatus("API INVITATION ABSENTE: ajouter /admin/users/invite dans dotations-api.", "error");
        return;
      }
      if (
        message.includes("MAGIC_LINK_ECHEC:429") ||
        message.includes("over_email_send_rate_limit")
      ) {
        setStatus("Quota Email atteint. Reessayez dans environ 1 heure.", "error");
      } else {
        setStatus(`Creation impossible: ${message}`, "error");
      }
    }
  });

  await fetchUsers();
  await fetchAppStateVersions();
  await fetchLoginEvents();
}

async function openAdminCreateUserFlow() {
  // Compat: keep old entry name if referenced elsewhere.
  await openAdminUsersModal();
}

async function refreshCurrentUserRoleLabel() {
  if (getDataBackendMode() === "LOCAL_API") {
    try {
      const response = await fetch("/api/local-session", { cache: "no-store" });
      if (!response.ok) {
        state.currentUserRoleLabel = "";
      } else {
        const payload = await response.json().catch(() => ({}));
        state.currentUserRoleLabel = String(payload?.role_label || "LOCAL_SECOURS");
      }
    } catch (error) {
      state.currentUserRoleLabel = "";
    }
    renderRoleBadge();
    return;
  }
  if (document.body?.dataset?.page === "mobile-signature") {
    state.currentUserRoleLabel = "";
    renderRoleBadge();
    return;
  }

  try {
    const userId = getStoredSupabaseUserId();
    const accessToken = getStoredSupabaseAccessToken();
    if (!userId || !accessToken) {
      state.currentUserRoleLabel = "";
      renderRoleBadge();
      return;
    }
    const endpoint = `${getSupabaseRestEndpoint().replace(/\/app_state$/i, "/profiles")}?id=eq.${encodeURIComponent(userId)}&select=role&limit=1`;
    const response = await fetch(endpoint, {
      method: "GET",
      headers: {
        apikey: String(SUPABASE_PUBLISHABLE_KEY || "").trim(),
        Authorization: `Bearer ${accessToken}`,
      },
      cache: "no-store",
    });
    if (!response.ok) {
      state.currentUserRoleLabel = "";
      renderRoleBadge();
      return;
    }
    const rows = await response.json().catch(() => []);
    const role = Array.isArray(rows) ? rows[0]?.role : "";
    state.currentUserRoleLabel = mapRoleToFrenchLabel(role);
    renderRoleBadge();
  } catch (error) {
    state.currentUserRoleLabel = "";
    renderRoleBadge();
  }
}

async function callEdgeApi(pathname, options = {}, retryOnAuthFailure = true) {
  if (getDataBackendMode() === "LOCAL_API") {
    const normalizedPath = String(pathname || "").trim().replace(/^\/+/, "");
    if (normalizedPath === "save") {
      const body = options?.body || "{}";
      return fetch(`/api/state/save?ts=${Date.now()}`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          ...(options.headers || {}),
        },
        body,
        cache: "no-store",
      });
    }
    if (normalizedPath === "data") {
      return fetch(appendPdfTokenToUrl(`/api/state?ts=${Date.now()}`), {
        method: "GET",
        headers: options.headers || {},
        cache: "no-store",
      });
    }
  }
  const baseUrl = normalizeHttpUrl(SUPABASE_EDGE_API_URL);
  const key = String(SUPABASE_PUBLISHABLE_KEY || "").trim();
  if (!baseUrl || !key) {
    throw new Error("EDGE_API_NOT_CONFIGURED");
  }
  const userToken = await getSupabaseUserAccessToken();
  if (!userToken) {
    throw new Error("BACKEND_AUTH_REQUIRED");
  }
  const normalizedPath = String(pathname || "").trim().replace(/^\/+/, "");
  const endpoint = `${baseUrl}/${normalizedPath}`;
  const method = String(options.method || "GET").toUpperCase();
  const start = performance.now();
  let response;
  try {
    response = await fetch(endpoint, {
      method,
      headers: {
        apikey: key,
        Authorization: `Bearer ${userToken}`,
        "Content-Type": "application/json",
        ...(options.headers || {}),
      },
      body: options.body || undefined,
      cache: "no-store",
    });
  } catch (error) {
    const durationMs = Math.round(performance.now() - start);
    if (options.trackDebug !== false) {
      recordNetworkDebugSample({
        route: normalizedPath,
        method,
        status: 0,
        durationMs,
        responseBytes: 0,
        source: "edge",
        note: "edge_api_network_error",
      });
    }
    throw error;
  }

  const durationMs = Math.round(performance.now() - start);
  const acceptedStatuses = Array.isArray(options?.acceptedStatuses)
    ? options.acceptedStatuses.map((code) => Number(code)).filter((code) => Number.isFinite(code))
    : [];
  const isExplicitlyAccepted = acceptedStatuses.includes(Number(response.status));
  const responseBytes = response.ok || isExplicitlyAccepted
    ? Number.parseInt(String(response.headers.get("content-length") || ""), 10) || 0
    : 0;
  if (options.trackDebug !== false) {
    recordNetworkDebugSample({
      route: normalizedPath,
      method,
      status: response.status,
      durationMs,
      responseBytes,
      source: "edge",
      note: response.ok ? (options.trackDebugNote || "edge_api_ok") : "edge_api_error",
    });
  }
  if (!response.ok && !isExplicitlyAccepted) {
    const detail = await response.text().catch(() => "");
    const detailUpper = String(detail || "").toUpperCase();
    if ((response.status === 401 || response.status === 403) && retryOnAuthFailure) {
      clearStoredSupabaseSession();
      return callEdgeApi(pathname, options, false);
    }
    if (
      response.status === 409 ||
      response.status === 412 ||
      detailUpper.includes("APP_STATE_CONFLICT") ||
      detailUpper.includes("CONFLIT")
    ) {
      throw buildSaveConflictError();
    }
    throw new Error(`EDGE_API_FAILED:${response.status}:${detail}`);
  }
  return response;
}

function getSupabaseHeaders(extra = {}, options = {}) {
  const key = String(SUPABASE_PUBLISHABLE_KEY || "").trim();
  const userAccessToken = String(getStoredSupabaseAccessToken() || "").trim();
  const includeAuthorization = options?.includeAuthorization ?? "auto";
  const headers = {
    apikey: key,
    ...extra,
  };
  const keyLooksLikeJwt = /^[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+$/.test(key);
  const keyLooksLikeSupabasePublishable = key.startsWith("sb_");
  const shouldAddAuthorization =
    includeAuthorization === true ||
    (includeAuthorization === "auto" && (keyLooksLikeJwt || keyLooksLikeSupabasePublishable));
  if (shouldAddAuthorization) {
    headers.Authorization = `Bearer ${userAccessToken || key}`;
  }
  return headers;
}

function encodeStorageObjectPath(path) {
  return String(path || "")
    .split("/")
    .map((part) => encodeURIComponent(String(part || "").trim()))
    .filter(Boolean)
    .join("/");
}

function getSupabaseStoragePublicUrl(bucket, objectPath) {
  const baseUrl = normalizeHttpUrl(SUPABASE_PROJECT_URL);
  const normalizedBucket = normalizeBucketName(bucket);
  const encodedPath = encodeStorageObjectPath(objectPath);
  if (!baseUrl || !normalizedBucket || !encodedPath) {
    return "";
  }
  return `${baseUrl}/storage/v1/object/public/${encodeURIComponent(normalizedBucket)}/${encodedPath}`;
}

function parseStorageSchemePath(value) {
  const raw = String(value || "").trim();
  if (!raw) {
    return null;
  }
  const match = raw.match(/^storage:\/\/([^\/]+)\/(.+)$/i);
  if (!match) {
    return null;
  }
  return {
    bucket: normalizeBucketName(match[1]),
    objectPath: String(match[2] || "").trim().replace(/^\/+/, ""),
  };
}

function getSupabaseProjectHost() {
  const baseUrl = normalizeHttpUrl(SUPABASE_PROJECT_URL);
  if (!baseUrl) {
    return "";
  }
  try {
    return new URL(baseUrl).hostname.toLowerCase();
  } catch (error) {
    return "";
  }
}

function isSafeArchiveHttpUrl(value) {
  const host = getSupabaseProjectHost();
  let parsed;
  try {
    parsed = new URL(value);
  } catch (error) {
    return false;
  }
  if (!["http:", "https:"].includes(parsed.protocol)) {
    return false;
  }
  const parsedHost = String(parsed.hostname || "").toLowerCase();
  const path = String(parsed.pathname || "");
  const isStoragePath = path.startsWith("/storage/v1/object/public/") || path.startsWith("/storage/v1/object/");
  if (!isStoragePath) {
    return false;
  }
  if (host) {
    return parsedHost === host;
  }
  return parsedHost.endsWith(".supabase.co");
}

function isSafeArchiveRelativePath(value) {
  const normalized = String(value || "").trim().replace(/^\/+/, "");
  if (!normalized) {
    return false;
  }
  if (normalized.length > 2048) {
    return false;
  }
  if (normalized.includes("\\")) {
    return false;
  }
  if (/(^|\/)\.\.(\/|$)/.test(normalized)) {
    return false;
  }
  // Autorise les URLs relatives de type document-*.html?param=... tout en restant strict.
  return /^[A-Za-z0-9._%+\-\/?&=]+$/.test(normalized);
}

function getStoragePdfBucketName() {
  return normalizeBucketName(
    state.data?.meta?.storagePdfBucket,
    DEFAULT_SUPABASE_PDF_BUCKET
  );
}

function getStorageSignaturesBucketName() {
  return normalizeBucketName(
    state.data?.meta?.storageSignaturesBucket,
    DEFAULT_SUPABASE_SIGNATURES_BUCKET
  );
}

function getSupabaseStorageUploadEndpoint(bucket, objectPath) {
  const baseUrl = normalizeHttpUrl(SUPABASE_PROJECT_URL);
  const normalizedBucket = normalizeBucketName(bucket);
  const encodedPath = encodeStorageObjectPath(objectPath);
  if (!baseUrl || !normalizedBucket || !encodedPath) {
    return "";
  }
  return `${baseUrl}/storage/v1/object/${encodeURIComponent(normalizedBucket)}/${encodedPath}`;
}

async function blobToBase64(blob) {
  const buffer = await blob.arrayBuffer();
  let binary = "";
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  for (let index = 0; index < bytes.length; index += chunkSize) {
    const chunk = bytes.subarray(index, index + chunkSize);
    binary += String.fromCharCode(...chunk);
  }
  return btoa(binary);
}

async function extractSupabaseErrorText(response, limit = 220) {
  let errorText = "";
  try {
    errorText = (await response.text()) || "";
  } catch (error) {
    errorText = "";
  }
  return errorText.replace(/\s+/g, " ").trim().slice(0, limit);
}

async function uploadBlobToSupabaseStorage(bucket, objectPath, blob, contentType) {
  if (getDataBackendMode() === "LOCAL_API") {
    const response = await fetch("/api/storage-upload", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        projectUrl: normalizeHttpUrl(SUPABASE_PROJECT_URL),
        publishableKey: String(SUPABASE_PUBLISHABLE_KEY || "").trim(),
        bucket: normalizeBucketName(bucket),
        objectPath: String(objectPath || "").trim().replace(/^\/+/, ""),
        contentType: String(contentType || "application/octet-stream"),
        payloadBase64: await blobToBase64(blob),
      }),
    });
    const payload = await response.json().catch(() => null);
    if (!response.ok || !payload?.ok) {
      const status = Number(payload?.status || response.status || 0);
      const details = String(payload?.error || "").trim().slice(0, 220);
      throw new Error(`SUPABASE STORAGE UPLOAD FAILED [${status}]${details ? ` ${details}` : ""}`);
    }
    return;
  }

  const endpoint = getSupabaseStorageUploadEndpoint(bucket, objectPath);
  if (!endpoint) {
    throw new Error("SUPABASE STORAGE ENDPOINT INVALIDE");
  }
  const executeUpload = async (includeAuthorization) => {
    const buildRequestOptions = (method) => ({
      method,
      headers: getSupabaseHeaders(
        {
          "Content-Type": contentType,
          "x-upsert": "true",
        },
        { includeAuthorization }
      ),
      body: blob,
    });
    let response = await fetch(endpoint, buildRequestOptions("POST"));
    if (!response.ok) {
      console.warn("[SUPABASE][STORAGE] POST failed", {
        status: response.status,
        bucket,
        objectPath,
        includeAuthorization: Boolean(includeAuthorization),
      });
      response = await fetch(endpoint, buildRequestOptions("PUT"));
    }
    return response;
  };

  let response = await executeUpload(false);
  if (!response.ok && (response.status === 401 || response.status === 403)) {
    console.warn("[SUPABASE][STORAGE] retry with Authorization header", {
      status: response.status,
      bucket,
      objectPath,
    });
    response = await executeUpload(true);
  }
  if (!response.ok) {
    const compactError = await extractSupabaseErrorText(response);
    throw new Error(
      `SUPABASE STORAGE UPLOAD FAILED [${response.status}]${compactError ? ` ${compactError}` : ""}`
    );
  }
}

function dataUrlToBlob(dataUrl) {
  const raw = String(dataUrl || "").trim();
  const match = raw.match(/^data:([^;,]+)?(?:;base64)?,(.*)$/i);
  if (!match) {
    return null;
  }
  const mimeType = match[1] || "application/octet-stream";
  const payload = match[2] || "";
  try {
    const binary = atob(payload);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) {
      bytes[index] = binary.charCodeAt(index);
    }
    return new Blob([bytes], { type: mimeType });
  } catch (error) {
    return null;
  }
}

async function uploadPdfBlobToSupabaseStorage(docType, person, blob, archiveMode = "STANDARD") {
  if (!isSupabaseConfigured() || !(blob instanceof Blob) || !person) {
    return null;
  }

  const bucket = getStoragePdfBucketName();
  if (!bucket) {
    return null;
  }

  const folder = normalizeText(docType) === "EXIT" ? "sortie" : "arrivee";
  const objectPath = `${folder}/${sanitizeFilePart(String(person.id || "P0000"))}/COURANT.pdf`;
  console.info("[SUPABASE][PDF] upload start", { bucket, objectPath });
  await uploadBlobToSupabaseStorage(bucket, objectPath, blob, "application/pdf");
  const storageRef = `storage://${bucket}/${objectPath}`;
  console.info("[SUPABASE][PDF] upload success", { storageRef });

  return {
    storageRef,
    publicUrl: getSupabaseStoragePublicUrl(bucket, objectPath),
  };
}

async function uploadSignatureImageToSupabaseStorage(docType, person, signer, signatureDataUrl) {
  if (!isSupabaseConfigured() || !person || !signatureDataUrl) {
    return null;
  }
  const bucket = getStorageSignaturesBucketName();
  if (!bucket) {
    return null;
  }
  const signatureBlob = dataUrlToBlob(signatureDataUrl);
  if (!(signatureBlob instanceof Blob) || signatureBlob.size <= 0) {
    throw new Error("SIGNATURE INVALIDE (BLOB VIDE)");
  }
  const folder = normalizeText(docType) === "EXIT" ? "sortie" : "arrivee";
  const signerLabel = normalizeText(signer) === "REPRESENTANT" ? "representant" : "personnel";
  const objectPath = `${folder}/${sanitizeFilePart(String(person.id || "P0000"))}/COURANT_${signerLabel}.png`;
  console.info("[SUPABASE][SIGNATURE] upload start", { bucket, objectPath });
  await uploadBlobToSupabaseStorage(bucket, objectPath, signatureBlob, "image/png");
  const storageRef = `storage://${bucket}/${objectPath}`;
  console.info("[SUPABASE][SIGNATURE] upload success", { storageRef });
  return {
    storageRef,
    publicUrl: getSupabaseStoragePublicUrl(bucket, objectPath),
  };
}

async function saveMobileSignatureRecordToSupabase({
  token = "",
  personId = "",
  docType = "",
  signer = "",
  person = null,
  signatureValue = "",
  validatedAt = "",
  storageRef = "",
  storagePublicUrl = "",
} = {}) {
  if (!isSupabaseConfigured() || !token || !personId || !signatureValue) {
    return false;
  }
  const normalizedDocType = normalizeText(docType) === "EXIT" ? "exit" : "arrival";
  const normalizedSigner = normalizeMobileSignatureSigner(signer || "");
  if (!normalizedDocType || !normalizedSigner) {
    return false;
  }
  const representative = normalizedSigner === "representant" ? getRepresentativeInfo(person, normalizedDocType) : null;
  const nowIso = new Date().toISOString();
  const row = {
    token: String(token || ""),
    person_id: String(personId || ""),
    doc_type: normalizedDocType,
    signer: normalizedSigner,
    status: "SIGNEE",
    signature_data: String(signatureValue || ""),
    storage_ref: String(storageRef || ""),
    storage_public_url: String(storagePublicUrl || ""),
    validated_at_text: String(validatedAt || ""),
    signed_at: nowIso,
    person_nom: normalizeText(person?.nom || ""),
    person_prenom: normalizeText(person?.prenom || ""),
    signer_name: normalizeText(representative?.nom || ""),
    signer_function: normalizeText(representative?.fonction || ""),
    updated_at: nowIso,
  };
  const camelRow = {
    token: row.token,
    personId: row.person_id,
    docType: row.doc_type,
    signer: row.signer,
    status: row.status,
    signatureData: row.signature_data,
    storageRef: row.storage_ref,
    storagePublicUrl: row.storage_public_url,
    validatedAt: row.validated_at_text,
    signedAt: row.signed_at,
    personNom: row.person_nom,
    personPrenom: row.person_prenom,
    signerName: row.signer_name,
    signerFunction: row.signer_function,
    updatedAt: row.updated_at,
  };
  const endpoint = `${normalizeHttpUrl(SUPABASE_PROJECT_URL)}/rest/v1/signatures?on_conflict=token`;
  const requestOptions = (payload) => ({
    method: "POST",
    headers: getSupabaseHeaders({
      "Content-Type": "application/json",
      Prefer: "resolution=merge-duplicates,return=minimal",
    }),
    body: JSON.stringify(payload),
    cache: "no-store",
  });
  let response = await fetch(endpoint, requestOptions(row));
  if (!response.ok) {
    const detail = await response.text().catch(() => "");
    if (isSupabaseSignatureSchemaMismatch(detail)) {
      response = await fetch(endpoint, requestOptions(camelRow));
      if (response.ok) {
        return true;
      }
      const camelDetail = await response.text().catch(() => "");
      throw new Error(`SUPABASE_SIGNATURE_RECORD_FAILED:${response.status}:${camelDetail.slice(0, 220)}`);
    }
    throw new Error(`SUPABASE_SIGNATURE_RECORD_FAILED:${response.status}:${detail.slice(0, 220)}`);
  }
  return true;
}

function isSupabaseSignatureSchemaMismatch(detail) {
  const text = String(detail || "");
  return /42703|column .* does not exist|person_id|doc_type|signature_data|storage_ref|signed_at|updated_at/i.test(text);
}

async function fetchSupabaseMobileSignatureRows(personId, docType) {
  if (!isSupabaseConfigured()) return [];
  const normalizedPersonId = String(personId || "").trim();
  const normalizedDocType = normalizeText(docType) === "EXIT" ? "exit" : "arrival";
  if (!normalizedPersonId || !normalizedDocType) return [];
  const endpoint = `${normalizeHttpUrl(SUPABASE_PROJECT_URL)}/rest/v1/signatures`;
  const buildUrl = (schema = "snake") => {
    if (schema === "camel") {
      return `${endpoint}?person_id=eq.${encodeURIComponent(normalizedPersonId)}&doc_type=eq.${encodeURIComponent(normalizedDocType)}&select=*&order=updated_at.desc,signed_at.desc&limit=30`;
    }
    return `${endpoint}?person_id=eq.${encodeURIComponent(normalizedPersonId)}&doc_type=eq.${encodeURIComponent(normalizedDocType)}&select=*&order=updated_at.desc,signed_at.desc&limit=30`;
  };
  const requestOptions = {
    method: "GET",
    headers: getSupabaseHeaders(),
    cache: "no-store",
  };
  let response = await fetch(buildUrl("snake"), requestOptions);
  if (!response.ok) {
    const detail = await response.text().catch(() => "");
    if (isSupabaseSignatureSchemaMismatch(detail)) {
      response = await fetch(buildUrl("camel"), requestOptions);
      if (response.ok) {
        const rows = await response.json().catch(() => []);
        return Array.isArray(rows) ? rows : [];
      }
      const camelDetail = await response.text().catch(() => "");
      throw new Error(`SUPABASE_SIGNATURE_ROWS_READ_FAILED:${response.status}:${camelDetail.slice(0, 180)}`);
    }
    throw new Error(`SUPABASE_SIGNATURE_ROWS_READ_FAILED:${response.status}:${detail.slice(0, 180)}`);
  }
  const rows = await response.json().catch(() => []);
  return Array.isArray(rows) ? rows : [];
}

function getSupabaseSignatureRowField(row, ...names) {
  for (const name of names) {
    if (row && Object.prototype.hasOwnProperty.call(row, name)) {
      const value = row[name];
      if (value !== null && value !== undefined && String(value).trim()) {
        return String(value).trim();
      }
    }
  }
  return "";
}

function mergeSupabaseMobileSignatureRows(data, rows, personId, docType) {
  const normalizedPersonId = String(personId || "").trim();
  const normalizedDocType = normalizeText(docType) === "EXIT" ? "exit" : "arrival";
  const person = Array.isArray(data?.personnes)
    ? data.personnes.find((entry) => String(entry?.id || "") === normalizedPersonId) || null
    : null;
  if (!person || !Array.isArray(rows) || !rows.length) return false;

  const latestBySigner = new Map();
  for (const row of rows) {
    const rowPersonId = getSupabaseSignatureRowField(row, "person_id", "personId");
    const rowDocType = normalizeText(getSupabaseSignatureRowField(row, "doc_type", "docType")) === "EXIT" ? "exit" : "arrival";
    const signer = normalizeMobileSignatureSigner(getSupabaseSignatureRowField(row, "signer"));
    if (rowPersonId && rowPersonId !== normalizedPersonId) continue;
    if (rowDocType !== normalizedDocType) continue;
    if (signer !== "personnel" && signer !== "representant") continue;
    const image = getSupabaseSignatureRowField(row, "signature_data", "signatureData", "image");
    const storageRef = getSupabaseSignatureRowField(row, "storage_ref", "storageRef");
    const storagePublicUrl = getSupabaseSignatureRowField(row, "storage_public_url", "storagePublicUrl");
    const validatedAt = getSupabaseSignatureRowField(row, "validated_at_text", "validatedAt", "signed_at", "signedAt", "updated_at", "updatedAt");
    const token = getSupabaseSignatureRowField(row, "token");
    if (!(image || storageRef || storagePublicUrl) || !validatedAt) continue;
    const rank = Date.parse(validatedAt) || Date.parse(getSupabaseSignatureRowField(row, "updated_at", "updatedAt")) || 0;
    const previous = latestBySigner.get(signer);
    if (!previous || rank >= previous.rank) {
      latestBySigner.set(signer, { image, storageRef, storagePublicUrl, validatedAt, token, rank });
    }
  }

  let changed = false;
  for (const signer of ["personnel", "representant"]) {
    const entry = latestBySigner.get(signer);
    if (!entry) continue;
    if (!person.signatures || typeof person.signatures !== "object") person.signatures = {};
    if (!person.signatures[normalizedDocType] || typeof person.signatures[normalizedDocType] !== "object") {
      person.signatures[normalizedDocType] = {};
    }
    const current = person.signatures[normalizedDocType][signer] || {};
    const currentMs = Date.parse(current.validatedAt || "") || 0;
    const nextMs = Date.parse(entry.validatedAt || "") || 0;
        const currentImage = String(current.image || "").trim();
    const currentStorageRef = String(current.storageRef || "").trim();
    const currentStoragePublicUrl = String(current.storagePublicUrl || "").trim();
    const nextImage = String(entry.image || "").trim();
    const nextStorageRef = String(entry.storageRef || "").trim();
    const nextStoragePublicUrl = String(entry.storagePublicUrl || "").trim();
    const currentHasPayload = Boolean(currentImage || currentStorageRef || currentStoragePublicUrl);
    const nextHasPayload = Boolean(nextImage || nextStorageRef || nextStoragePublicUrl);
    const nextSignatureImage = nextImage || nextStoragePublicUrl || (nextStorageRef ? `storage://${DEFAULT_SUPABASE_SIGNATURES_BUCKET}/${nextStorageRef}` : "");
    const payloadChanged =
      currentImage !== nextSignatureImage ||
      currentStorageRef !== nextStorageRef ||
      currentStoragePublicUrl !== nextStoragePublicUrl ||
      String(current.validatedAt || "") !== String(entry.validatedAt || "");
    if (nextHasPayload && (!currentHasPayload || nextMs >= currentMs || payloadChanged)) {
      person.signatures[normalizedDocType][signer] = {
        image: nextSignatureImage,
        validatedAt: entry.validatedAt,
        storageRef: nextStorageRef,
        storagePublicUrl: nextStoragePublicUrl,
      };
      changed = true;
    }
    const requests = Array.isArray(data?.demandesSignatureMobile) ? data.demandesSignatureMobile : [];
    for (const request of requests) {
      if (!request || typeof request !== "object") continue;
      const sameToken = entry.token && String(request.token || "") === String(entry.token || "");
      const sameContext =
        String(request.personId || "") === normalizedPersonId &&
        normalizeText(request.docType || "") === normalizeText(normalizedDocType) &&
        normalizeMobileSignatureSigner(request.signer || "") === signer;
      if (!sameToken && !(sameContext && !entry.token)) continue;
      if (request.status !== "SIGNEE" || String(request.validatedAt || "") !== String(entry.validatedAt || "")) {
        request.status = "SIGNEE";
        request.validatedAt = entry.validatedAt;
        changed = true;
      }
    }
  }
  return changed;
}
async function fetchSupabaseStateData() {
  if (getDataBackendMode() === "LOCAL_API") {
    const response = await fetch(appendPdfTokenToUrl(`/api/state?ts=${Date.now()}`), { cache: "no-store" });
    if (!response.ok) {
      throw new Error("LOCAL LOAD FAILED");
    }
    const payload = await response.json();
    const data = payload && typeof payload === "object" && Object.prototype.hasOwnProperty.call(payload, "data")
      ? payload.data
      : payload;
    state.supabaseRevision = Number.isFinite(Number(payload?.revision)) ? Number(payload.revision) : 0;
    return data;
  }
  const start = performance.now();
  const endpoint = `${getSupabaseRestEndpoint()}?id=eq.${encodeURIComponent(
    SUPABASE_APP_STATE_ID
  )}&select=payload,revision&limit=1`;
  const response = await fetch(endpoint, {
    headers: getSupabaseHeaders(),
    cache: "no-store",
  });
  if (!response.ok) {
    const durationMs = Math.round(performance.now() - start);
    const responseBodyForDebug = await response.text().catch(() => "");
    recordNetworkDebugSample({
      route: endpoint,
      method: "GET",
      status: response.status,
      durationMs,
      responseBytes: estimateByteLength(responseBodyForDebug),
      source: "supabase-rest-state",
      note: "app_state fetch failed",
    });
    throw new Error("SUPABASE LOAD FAILED");
  }
  const rows = await response.json();
  if (!Array.isArray(rows) || !rows.length || !rows[0]?.payload) {
    const durationMs = Math.round(performance.now() - start);
    recordNetworkDebugSample({
      route: endpoint,
      method: "GET",
      status: 204,
      durationMs,
      source: "supabase-rest-state",
      note: "app_state payload empty",
    });
    throw new Error("SUPABASE EMPTY PAYLOAD");
  }
  const revision = Number(rows[0]?.revision);
  state.supabaseRevision = Number.isFinite(revision) ? revision : 0;
  const durationMs = Math.round(performance.now() - start);
  recordNetworkDebugSample({
    route: endpoint,
    method: "GET",
    status: response.status,
    durationMs,
    responseBytes: estimateByteLength(rows[0]?.payload),
    source: "supabase-rest-state",
  });
  return rows[0].payload;
}

function buildSaveConflictError() {
  const error = new Error("Conflit de sauvegarde : les donnees ont ete modifiees ailleurs. Recharge puis reessaie.");
  error.code = "APP_STATE_CONFLICT";
  return error;
}

function isSaveConflictError(error) {
  return String(error?.code || "") === "APP_STATE_CONFLICT";
}

function estimateByteLength(value) {
  if (!value) {
    return 0;
  }
  if (typeof value === "string") {
    return new TextEncoder().encode(value).length;
  }
  if (value instanceof Blob) {
    return Number.isFinite(value.size) ? value.size : 0;
  }
  if (value instanceof ArrayBuffer) {
    return Number.isFinite(value.byteLength) ? value.byteLength : 0;
  }
  try {
    return new TextEncoder().encode(JSON.stringify(value)).length;
  } catch (error) {
    return 0;
  }
}

function getNetworkDebugRoute(pathname) {
  try {
    if (!pathname) return "unknown";
    const url = new URL(String(pathname), "https://example.local");
    return String(url.pathname || "unknown").replace(/\/+/g, "/");
  } catch (error) {
    return String(pathname || "unknown");
  }
}

function recordNetworkDebugSample(sample = {}) {
  const route = getNetworkDebugRoute(sample.route);
  const method = String(sample.method || "GET").toUpperCase();
  const status = Number.isFinite(Number(sample.status)) ? Number(sample.status) : 0;
  const durationMs = Number.isFinite(Number(sample.durationMs)) ? Math.max(0, Number(sample.durationMs)) : 0;
  const responseBytes = Math.max(0, Number(sample.responseBytes) || 0);
  const requestBytes = Math.max(0, Number(sample.requestBytes) || 0);
  const cached = Boolean(sample.cached);

  if (!state.networkDebug) {
    state.networkDebug = {
      samples: [],
      routeStats: {},
      requestCount: 0,
      totalBytes: 0,
      startedAt: Date.now(),
    };
  }

  const key = `${method} ${route}`;
  const stat = state.networkDebug.routeStats[key] || {
    route,
    method,
    count: 0,
    bytes: 0,
    avgMs: 0,
    lastStatus: 0,
  };
  const nextCount = stat.count + 1;
  stat.count = nextCount;
  stat.bytes += requestBytes + responseBytes;
  stat.avgMs = ((stat.avgMs * (nextCount - 1)) + durationMs) / nextCount;
  stat.lastStatus = status;
  state.networkDebug.routeStats[key] = stat;

  state.networkDebug.samples.push({
    at: Date.now(),
    route,
    method,
    status,
    durationMs,
    requestBytes,
    responseBytes,
    cached,
    source: String(sample.source || "network"),
    note: String(sample.note || ""),
  });
  if (state.networkDebug.samples.length > NETWORK_DEBUG_SAMPLE_LIMIT) {
    state.networkDebug.samples.shift();
  }
  state.networkDebug.requestCount += 1;
  state.networkDebug.totalBytes += requestBytes + responseBytes;

  try {
    window.localStorage.setItem(
      NETWORK_DEBUG_STORAGE_KEY,
      JSON.stringify({
        at: Date.now(),
        route,
        method,
        status,
        durationMs,
        requestBytes,
        responseBytes,
        cached,
      }),
    );
  } catch (error) {
    // ignore storage write failures
  }
}

function getNetworkDebugReport() {
  const now = Date.now();
  const samples = Array.isArray(state.networkDebug?.samples) ? state.networkDebug.samples : [];
  const windowSamples = samples.filter((entry) => now - (entry?.at || 0) <= NETWORK_DEBUG_WINDOW_MS);
  const windowBytes = windowSamples.reduce(
    (sum, entry) => sum + Number(entry.requestBytes || 0) + Number(entry.responseBytes || 0),
    0,
  );
  return {
    requestCount: state.networkDebug?.requestCount || 0,
    totalBytes: state.networkDebug?.totalBytes || 0,
    routeStats: { ...(state.networkDebug?.routeStats || {}) },
    samples,
    windowMinutes: NETWORK_DEBUG_WINDOW_MS / (60 * 1000),
    windowRequestCount: windowSamples.length,
    windowBytes,
    startedAt: state.networkDebug?.startedAt || now,
  };
}

function isSoftDeletedEntity(entry) {
  if (!entry || typeof entry !== "object") return false;
  return entry.is_deleted === true || entry.isDeleted === true;
}

function markSoftDeletedEntity(entry = {}) {
  const nowIso = new Date().toISOString();
  return {
    ...(entry || {}),
    is_deleted: true,
    isDeleted: true,
    deleted_at: nowIso,
    deletedAt: nowIso,
    updated_at: nowIso,
    updatedAt: nowIso,
  };
}

async function saveSupabaseStateData(payload) {
  if (getDataBackendMode() === "LOCAL_API") {
    const response = await fetch(`/api/state/save?ts=${Date.now()}`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ data: payload }),
      cache: "no-store",
    });
    if (!response.ok) {
      throw new Error("LOCAL SAVE FAILED");
    }
    return;
  }
  if (!Number.isFinite(Number(state.supabaseRevision))) {
    await fetchSupabaseStateData();
  }
  const expectedRevision = Number(state.supabaseRevision);
  if (!Number.isFinite(expectedRevision)) {
    throw new Error("SUPABASE REVISION INDISPONIBLE");
  }

  const endpoint = `${getSupabaseRestEndpoint()}?id=eq.${encodeURIComponent(
    SUPABASE_APP_STATE_ID
  )}&revision=eq.${encodeURIComponent(String(expectedRevision))}&select=id,revision`;
  const response = await fetch(endpoint, {
    method: "PATCH",
    headers: getSupabaseHeaders({
      "Content-Type": "application/json",
      Prefer: "return=representation",
    }),
    body: JSON.stringify({
      payload,
      revision: expectedRevision + 1,
    }),
  });
  if (!response.ok) {
    throw new Error("SUPABASE SAVE FAILED");
  }

  const rows = await response.json().catch(() => null);
  const updated = Array.isArray(rows) ? rows[0] : null;
  if (!updated?.id) {
    throw buildSaveConflictError();
  }
  const nextRevision = Number(updated.revision);
  if (Number.isFinite(nextRevision)) {
    state.supabaseRevision = nextRevision;
  }
}

async function saveSupabasePayloadWithRetry(payload, maxAttempts = 3) {
  if (getDataBackendMode() === "LOCAL_API") {
    await saveSupabaseStateData(payload);
    return;
  }
  void maxAttempts;
  for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
    if (!Number.isFinite(Number(state.supabaseRevision))) {
      await fetchSupabaseStateData();
    }
    try {
      await callEdgeApi("save", {
        method: "POST",
        body: JSON.stringify({
          payload,
          expectedRevision: Number(state.supabaseRevision),
        }),
      });
      return;
    } catch (error) {
      if (!isSaveConflictError(error)) {
        throw error;
      }
      // Do not retry the same stale payload with a fresh revision: it can overwrite
      // a mobile signature that was saved by another device between the two writes.
      await fetchSupabaseStateData();
      throw error;
    }
  }
  throw buildSaveConflictError();
}

async function forceSaveSupabaseStateData(payload) {
  const currentRevision = Number(state.supabaseRevision);
  const nextRevision = Number.isFinite(currentRevision) ? currentRevision + 1 : 1;
  const endpoint = `${getSupabaseRestEndpoint()}?id=eq.${encodeURIComponent(
    SUPABASE_APP_STATE_ID
  )}&select=id,revision`;
  const response = await fetch(endpoint, {
    method: "PATCH",
    headers: getSupabaseHeaders({
      "Content-Type": "application/json",
      Prefer: "return=representation",
    }),
    body: JSON.stringify({
      payload,
      revision: nextRevision,
    }),
  });
  if (!response.ok) {
    throw new Error("SUPABASE FORCE SAVE FAILED");
  }
  const rows = await response.json().catch(() => null);
  const updated = Array.isArray(rows) ? rows[0] : null;
  if (!updated?.id) {
    throw new Error("SUPABASE FORCE SAVE EMPTY");
  }
  const persistedRevision = Number(updated.revision);
  state.supabaseRevision = Number.isFinite(persistedRevision) ? persistedRevision : nextRevision;
}

async function saveSupabaseSignatureWithRebase({
  personId,
  docType,
  signer,
  signatureValue,
  validatedAt,
  storageRef = "",
  storagePublicUrl = "",
  mobileRequestToken = "",
}) {
  const normalizedPersonId = String(personId || "");
  const normalizedDocType = normalizeText(docType) === "EXIT" ? "exit" : "arrival";
  const normalizedSigner = normalizeMobileSignatureSigner(signer || "");
  if (!normalizedPersonId || !normalizedDocType || !normalizedSigner) {
    throw new Error("SIGNATURE CONTEXT INVALIDE");
  }
  let lastConflictError = null;
  for (let attempt = 0; attempt < 8; attempt += 1) {
    const latestPayload = await fetchSupabaseStateData();
    const persons = Array.isArray(latestPayload?.personnes) ? latestPayload.personnes : [];
    const requests = Array.isArray(latestPayload?.demandesSignatureMobile) ? latestPayload.demandesSignatureMobile : [];
    const request = requests.find((entry) => String(entry?.token || "") === mobileRequestToken);
    let person = persons.find((entry) => String(entry?.id || "") === normalizedPersonId);
    if (!person && request) {
      person = buildMobileSignaturePersonFromRequest(request, normalizedPersonId);
      if (person) {
        latestPayload.personnes = persons;
        persons.push(person);
      }
    }
    if (!person) throw new Error("PERSONNE INTROUVABLE POUR SAUVEGARDE SIGNATURE");

    setSignatureValue(
      person,
      normalizedDocType,
      normalizedSigner,
      String(signatureValue || ""),
      String(validatedAt || ""),
      String(storageRef || ""),
      String(storagePublicUrl || "")
    );
    if (normalizedDocType === "exit" && String(signatureValue || "")) {
      applySignedExitCompletion(person);
    }
    if (request) markMobileSignatureRequestSigned(request);

    try {
      await saveSupabaseStateData(latestPayload);
      state.data = latestPayload;
      return latestPayload;
    } catch (error) {
      if (isSaveConflictError(error)) {
        lastConflictError = error;
        continue;
      }
      throw error;
    }
  }
  try {
    const latestPayload = await fetchSupabaseStateData();
    const persons = Array.isArray(latestPayload?.personnes) ? latestPayload.personnes : [];
    const requests = Array.isArray(latestPayload?.demandesSignatureMobile) ? latestPayload.demandesSignatureMobile : [];
    const request = requests.find((entry) => String(entry?.token || "") === mobileRequestToken);
    let person = persons.find((entry) => String(entry?.id || "") === normalizedPersonId);
    if (!person && request) {
      person = buildMobileSignaturePersonFromRequest(request, normalizedPersonId);
      if (person) {
        latestPayload.personnes = persons;
        persons.push(person);
      }
    }
    if (!person) throw new Error("PERSONNE INTROUVABLE POUR SAUVEGARDE SIGNATURE");
    setSignatureValue(
      person,
      normalizedDocType,
      normalizedSigner,
      String(signatureValue || ""),
      String(validatedAt || ""),
      String(storageRef || ""),
      String(storagePublicUrl || "")
    );
    if (normalizedDocType === "exit" && String(signatureValue || "")) {
      applySignedExitCompletion(person);
    }
    if (request) markMobileSignatureRequestSigned(request);
    await forceSaveSupabaseStateData(latestPayload);
    state.data = latestPayload;
    return latestPayload;
  } catch (error) {
    throw lastConflictError || error || buildSaveConflictError();
  }
}

async function fetchLatestDataSnapshot({ forceFresh = false } = {}) {
  const now = Date.now();
  const hasFreshCachedSnapshot = state.latestDataSnapshotCache && now - (state.latestDataFetchAt || 0) <= DATA_FETCH_DEBOUNCE_MS;
  if (state.latestDataFetchPromise) {
    return state.latestDataFetchPromise;
  }
  if (!forceFresh && hasFreshCachedSnapshot && !state.isDirty) {
    return state.latestDataSnapshotCache;
  }

  const fetchPromise = (async () => {
    const snapshotLoadStart = performance.now();
    const mode = getDataBackendMode();
    if (mode === "SUPABASE") {
      const etagKey = "__dotations_data_etag";
      const cacheKey = "__dotations_data_snapshot";
      let knownEtag = String(state.latestDataEtag || "");
      let cachedSnapshot = null;
      try {
        const storedEtag = String(sessionStorage.getItem(etagKey) || "");
        if (storedEtag) {
          knownEtag = storedEtag;
        }
        const raw = sessionStorage.getItem(cacheKey);
        cachedSnapshot = raw ? JSON.parse(raw) : null;
      } catch (error) {
        cachedSnapshot = null;
      }
      const headers = knownEtag ? { "If-None-Match": knownEtag } : {};
      const response = await callEdgeApi(
        "data",
        {
          method: "GET",
          headers,
          acceptedStatuses: [304],
          trackDebug: false,
        }
      );
      if (response.status === 304) {
        const durationMs = Math.round(performance.now() - snapshotLoadStart);
        if (cachedSnapshot && typeof cachedSnapshot === "object") {
          recordNetworkDebugSample({
            route: "api/data",
            method: "GET",
            status: 304,
            durationMs,
            responseBytes: 0,
            cached: true,
            source: "edge",
            note: "data_snapshot_cached",
          });
          return cachedSnapshot;
        }
        const fallbackResponse = await callEdgeApi("data", { method: "GET", trackDebug: false });
        const fallbackSnapshot = await fallbackResponse.json();
        const fallbackEtag = String(fallbackResponse.headers.get("etag") || "").trim();
        if (fallbackEtag) {
          state.latestDataEtag = fallbackEtag;
        }
        const fallbackDurationMs = Math.round(performance.now() - snapshotLoadStart);
        recordNetworkDebugSample({
          route: "api/data",
          method: "GET",
          status: fallbackResponse.status,
          durationMs: fallbackDurationMs,
          responseBytes: estimateByteLength(fallbackSnapshot),
          source: "edge",
          note: "data_snapshot_fallback",
        });
        return fallbackSnapshot;
      }
      const nextSnapshot = await response.json();
      const durationMs = Math.round(performance.now() - snapshotLoadStart);
      const responseEtag = String(response.headers.get("etag") || "").trim();
      if (responseEtag) {
        state.latestDataEtag = responseEtag;
      }
      if (responseEtag && nextSnapshot && typeof nextSnapshot === "object") {
        try {
          sessionStorage.setItem(etagKey, responseEtag);
        } catch (error) {
          // ignore etag cache write failures
        }
        try {
          sessionStorage.setItem(cacheKey, JSON.stringify(nextSnapshot));
        } catch (error) {
          // ignore storage cache failures
        }
      }
      recordNetworkDebugSample({
        route: "api/data",
        method: "GET",
        status: response.status,
        durationMs,
        responseBytes: estimateByteLength(nextSnapshot),
        source: "edge",
        note: "data_snapshot_loaded",
      });
      return nextSnapshot;
    }
    const start = performance.now();
    const response = await fetch(appendPdfTokenToUrl(`/api/state?ts=${Date.now()}`), { cache: "no-store" });
    if (!response.ok) {
      const durationMs = Math.round(performance.now() - start);
      const responseText = await response.text().catch(() => "");
      recordNetworkDebugSample({
        route: "/api/data",
        method: "GET",
        status: response.status,
        durationMs,
        responseBytes: estimateByteLength(responseText),
        source: "local",
        note: "api_data_failed",
      });
      throw new Error("LOCAL LOAD FAILED");
    }
    const responsePayload = await response.json();
    const responseData =
      responsePayload &&
      typeof responsePayload === "object" &&
      Object.prototype.hasOwnProperty.call(responsePayload, "data")
        ? responsePayload.data
        : responsePayload;
    const durationMs = Math.round(performance.now() - start);
    recordNetworkDebugSample({
      route: "/api/data",
      method: "GET",
      status: response.status,
      durationMs,
      responseBytes: estimateByteLength(responseData),
      source: "local",
      note: "api_data_loaded",
    });
    return responseData;
  })();

  state.latestDataFetchPromise = fetchPromise;
  try {
    const snapshot = await fetchPromise;
    state.latestDataSnapshotCache = snapshot;
    state.latestDataFetchAt = Date.now();
    return snapshot;
  } finally {
    if (state.latestDataFetchPromise === fetchPromise) {
      state.latestDataFetchPromise = null;
    }
  }
}

async function fetchHostedSupabaseStateSnapshot() {
  if (!isSupabaseConfigured()) {
    throw new Error("SUPABASE NON CONFIGURE");
  }
  const endpoint = `${getSupabaseRestEndpoint()}?id=eq.${encodeURIComponent(
    SUPABASE_APP_STATE_ID
  )}&select=payload,revision&limit=1`;
  const response = await fetch(endpoint, {
    headers: getSupabaseHeaders(),
    cache: "no-store",
  });
  if (!response.ok) {
    throw new Error(`SUPABASE LOAD FAILED:${response.status}`);
  }
  const rows = await response.json();
  if (!Array.isArray(rows) || !rows.length || !rows[0]?.payload) {
    throw new Error("SUPABASE EMPTY PAYLOAD");
  }
  const revision = Number(rows[0]?.revision);
  if (Number.isFinite(revision)) {
    state.supabaseRevision = revision;
  }
  return rows[0].payload;
}

function mergeHostedMobileSignaturesIntoLocal(localPayload, hostedPayload, personId, docType) {
  if (!localPayload || !hostedPayload) {
    return false;
  }
  const normalizedPersonId = String(personId || "");
  const normalizedDocType = normalizeText(docType) === "EXIT" ? "exit" : "arrival";
  const localPerson = Array.isArray(localPayload.personnes)
    ? localPayload.personnes.find((entry) => String(entry?.id || "") === normalizedPersonId)
    : null;
  const hostedPerson = Array.isArray(hostedPayload.personnes)
    ? hostedPayload.personnes.find((entry) => String(entry?.id || "") === normalizedPersonId)
    : null;
  if (!localPerson || !hostedPerson) {
    return false;
  }

  let changed = false;
  ["personnel", "representant"].forEach((signer) => {
    const hostedEntry = hostedPerson.signatures?.[normalizedDocType]?.[signer] || {};
    const hostedImage = String(hostedEntry.image || "").trim();
    const hostedValidatedAt = String(hostedEntry.validatedAt || "").trim();
    const hostedStorageRef = String(hostedEntry.storageRef || "").trim();
    const hostedStoragePublicUrl = String(hostedEntry.storagePublicUrl || "").trim();
    const hostedHasSignature = Boolean(hostedImage || hostedStorageRef || hostedStoragePublicUrl);
    if (!hostedHasSignature || !hostedValidatedAt) {
      return;
    }
    const localEntry = localPerson.signatures?.[normalizedDocType]?.[signer] || {};
    const localImage = String(localEntry.image || "").trim();
    const localValidatedAt = String(localEntry.validatedAt || "").trim();
    const hostedMs = Date.parse(hostedValidatedAt) || 0;
    const localMs = Date.parse(localValidatedAt) || 0;
    if (localImage === hostedImage && localValidatedAt === hostedValidatedAt) {
      return;
    }
    if (localImage && localMs > hostedMs) {
      return;
    }
    setSignatureValue(
      localPerson,
      normalizedDocType,
      signer,
      hostedImage,
      hostedValidatedAt,
      hostedStorageRef,
      hostedStoragePublicUrl
    );
    changed = true;
  });

  const localRequests = Array.isArray(localPayload.demandesSignatureMobile)
    ? localPayload.demandesSignatureMobile
    : [];
  const hostedRequests = Array.isArray(hostedPayload.demandesSignatureMobile)
    ? hostedPayload.demandesSignatureMobile
    : [];
  const hostedByToken = new Map(
    hostedRequests
      .map((request) => [String(request?.token || ""), request])
      .filter(([token]) => Boolean(token))
  );
  localRequests.forEach((localRequest) => {
    const token = String(localRequest?.token || "");
    const hostedRequest = hostedByToken.get(token);
    if (!hostedRequest) {
      return;
    }
    if (
      String(localRequest.personId || "") !== normalizedPersonId ||
      normalizeText(localRequest.docType || "") !== normalizeText(normalizedDocType)
    ) {
      return;
    }
    ["status", "validatedAt"].forEach((key) => {
      const nextValue = String(hostedRequest?.[key] || "");
      if (nextValue && String(localRequest?.[key] || "") !== nextValue) {
        localRequest[key] = nextValue;
        changed = true;
      }
    });
  });

  return changed;
}

function isLikelyLocalUrl(value) {
  const normalized = normalizeHttpUrl(value);
  if (!normalized) {
    return false;
  }
  try {
    const host = String(new URL(normalized).hostname || "").toLowerCase();
    if (host === "localhost" || host === "127.0.0.1") {
      return true;
    }
    if (host.startsWith("10.") || host.startsWith("192.168.") || host.startsWith("169.254.")) {
      return true;
    }
    if (host.startsWith("172.")) {
      const second = Number.parseInt(host.split(".")[1] || "", 10);
      return Number.isFinite(second) && second >= 16 && second <= 31;
    }
    return false;
  } catch (error) {
    return false;
  }
}

function getConfiguredMobileSignatureBaseUrl() {
  return normalizeMobileSignatureBaseUrl(state.data?.meta?.signatureMobileBaseUrl || "");
}

function getMobileSignatureReachabilityHint(url) {
  if (!url) {
    return "";
  }
  if (isLikelyLocalUrl(url)) {
    return "TELEPHONE ET ORDINATEUR DOIVENT ETRE SUR LE MEME RESEAU WIFI.";
  }
  return "LIEN OUVRABLE EN WIFI OU EN 4G/5G.";
}

function areSameHost(leftUrl, rightUrl) {
  try {
    const leftHost = String(new URL(leftUrl).hostname || "").toLowerCase().replace(/^www\./, "");
    const rightHost = String(new URL(rightUrl).hostname || "").toLowerCase().replace(/^www\./, "");
    return Boolean(leftHost && rightHost && leftHost === rightHost);
  } catch (error) {
    return false;
  }
}

function compareTextValues(left, right) {
  return normalizeText(left).localeCompare(normalizeText(right), "fr");
}

function formatAmount(value) {
  const amount = normalizeAmount(value);
  return amount.toLocaleString("fr-FR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function formatAmountWithEuro(value) {
  return `${formatAmount(value)} €`;
}

function getCurrentSignatureTimestamp() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  const hours = String(now.getHours()).padStart(2, "0");
  const minutes = String(now.getMinutes()).padStart(2, "0");
  const seconds = String(now.getSeconds()).padStart(2, "0");
  return `${year}-${month}-${day}T${hours}:${minutes}:${seconds}`;
}

function formatSignatureTimestamp(value) {
  if (!value) {
    return "";
  }
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return "";
  }
  return date.toLocaleString("fr-FR", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function formatCurrentUiTimestamp() {
  return formatSignatureTimestamp(getCurrentSignatureTimestamp());
}

function typeUsesReferenceCatalog(typeEffet) {
  return ["CLE", "CLE CES", "CLE DE SECURITE", "VENTILATEUR"].includes(normalizeText(typeEffet));
}

function getReferenceCatalogType(typeEffet) {
  return normalizeText(typeEffet);
}

function getReferenceEffectiveType(reference) {
  const baseType = normalizeText(reference?.typeEffet || "");
  if (baseType === "CLE" && isCesKeyDesignation(reference?.designation || "")) {
    return "CLE CES";
  }
  return baseType;
}

function getStockReferenceDesignation(reference) {
  return normalizeText(reference?.designation || "") || STOCK_EMPTY_DESIGNATION_LABEL;
}

function getStockGroupingDesignation(typeEffet, designation) {
  const normalizedType = normalizeText(typeEffet);
  if (["BADGE INTRUSION", "CARTE TURBOSELF", "RADIATEUR APPOINT", "TELECOMMANDE URMET"].includes(normalizedType)) {
    return STOCK_EMPTY_DESIGNATION_LABEL;
  }
  return normalizeText(designation || "") || STOCK_EMPTY_DESIGNATION_LABEL;
}

function typeIgnoresStockDesignation(typeEffet) {
  return ["BADGE INTRUSION", "CARTE TURBOSELF", "RADIATEUR APPOINT", "TELECOMMANDE URMET"].includes(normalizeText(typeEffet));
}

function getStockSyntheticReferenceValue(site, typeEffet, designation) {
  const parts = [
    normalizeText(site) || "SANS SITE",
    normalizeText(typeEffet),
    getStockGroupingDesignation(typeEffet, designation),
  ];
  return `${STOCK_SYNTHETIC_REFERENCE_PREFIX}${parts.map((part) => encodeURIComponent(part)).join("|")}`;
}

function parseStockSyntheticReferenceValue(value) {
  const rawValue = String(value || "");
  if (!rawValue.startsWith(STOCK_SYNTHETIC_REFERENCE_PREFIX)) {
    return null;
  }
  try {
    const [site = "", typeEffet = "", designation = ""] = rawValue
      .slice(STOCK_SYNTHETIC_REFERENCE_PREFIX.length)
      .split("|")
      .map((part) => normalizeText(decodeURIComponent(part)));
    return {
      site: site || "SANS SITE",
      typeEffet,
      designation: designation || STOCK_EMPTY_DESIGNATION_LABEL,
    };
  } catch (error) {
    console.error(error);
    return null;
  }
}

function referenceMatchesType(reference, selectedType) {
  const normalizedType = normalizeText(selectedType || "");
  if (!normalizedType) {
    return true;
  }
  return getReferenceEffectiveType(reference) === normalizedType;
}

function typeUsesSiteField(typeEffet) {
  return Boolean(normalizeText(typeEffet));
}

function normalizeSites(values) {
  const normalizedValues = Array.from(new Set((values || []).map(normalizeText).filter(Boolean)));
  if (normalizedValues.includes(ALL_SITES_VALUE)) {
    return [ALL_SITES_VALUE];
  }
  return normalizedValues;
}

function getPersonSites(person) {
  const baseValues = Array.isArray(person?.sitesAffectation)
    ? person.sitesAffectation
    : Array.isArray(person?.sites)
      ? person.sites
      : person?.site
        ? String(person.site).split("/").map((value) => value.trim())
        : [];
  return normalizeSites(baseValues);
}

function personUsesAllSites(person) {
  return getPersonSites(person).includes(ALL_SITES_VALUE);
}

function getPersonSiteLabel(person) {
  const sites = getPersonSites(person);
  return sites.length ? sites.join(" / ") : "";
}

function getReferenceSites(reference) {
  const baseValues = Array.isArray(reference?.sitesAffectation)
    ? reference.sitesAffectation
    : reference?.site
      ? String(reference.site).split("/").map((value) => value.trim())
      : [];
  return normalizeSites(baseValues);
}

function getReferenceSiteLabel(reference) {
  const sites = getReferenceSites(reference);
  return sites.length ? sites.join(" / ") : "";
}

function isReferenceEffectActive(reference) {
  return reference?.active !== false;
}

function referenceHasSite(reference, site) {
  const normalizedSite = normalizeText(site);
  if (!normalizedSite) {
    return true;
  }
  const sites = getReferenceSites(reference);
  return sites.includes(ALL_SITES_VALUE) || sites.includes(normalizedSite);
}

function resolveStockMovementSite(site, reference) {
  const normalizedSite = normalizeText(site);
  if (normalizedSite !== ALL_SITES_VALUE || !reference) {
    return normalizedSite;
  }
  const referenceSites = getReferenceSites(reference).filter((entry) => normalizeText(entry) !== ALL_SITES_VALUE);
  return referenceSites.length === 1 ? referenceSites[0] : normalizedSite;
}

function getComparableSites(value) {
  return normalizeSites(value).slice().sort((left, right) => left.localeCompare(right, "fr"));
}

function haveSameSites(leftSites, rightSites) {
  const left = getComparableSites(leftSites);
  const right = getComparableSites(rightSites);
  if (left.length !== right.length) {
    return false;
  }
  return left.every((value, index) => value === right[index]);
}

function getPersonSiteMarkup(person) {
  const sites = getPersonSites(person);
  if (!sites.length) {
    return "";
  }

  return sites
    .map((site) => {
      const classes =
        normalizeText(site) === ALL_SITES_VALUE ? "site-pill site-pill--all" : "site-pill";
      return `<span class="${classes}">${escapeHtml(site)}</span>`;
    })
    .join(" ");
}

function personHasSite(person, site) {
  const normalizedSite = normalizeText(site);
  if (!normalizedSite) {
    return true;
  }
  const sites = getPersonSites(person);
  return sites.includes(ALL_SITES_VALUE) || sites.includes(normalizedSite);
}

function readSelectedSites(form, prefix) {
  const values = Array.from(form.querySelectorAll(`input[name="${prefix}Sites"]:checked`)).map(
    (input) => normalizeText(input.value)
  );
  return normalizeSites(values);
}

function redirectToLocalServerIfNeeded() {
  return;
}

function getStoredNavigationContext() {
  try {
    const raw = window.sessionStorage.getItem(NAVIGATION_CONTEXT_KEY);
    if (!raw) {
      return null;
    }
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : null;
  } catch (error) {
    return null;
  }
}

function saveNavigationContext(partial = {}) {
  const current = getStoredNavigationContext() || {};
  const merged = {
    personId: String(partial.personId ?? current.personId ?? ""),
    urgentMode: Boolean(partial.urgentMode ?? current.urgentMode ?? false),
    filters: {
      ...DEFAULT_FILTERS,
      ...(current.filters || {}),
      ...(partial.filters || {}),
    },
  };
  try {
    window.sessionStorage.setItem(NAVIGATION_CONTEXT_KEY, JSON.stringify(merged));
  } catch (error) {
    // Ignore storage errors
  }
}

function restoreNavigationContext() {
  const context = getStoredNavigationContext();
  const params = new URLSearchParams(window.location.search);
  const personIdInQuery = params.get("personId") || params.get("personld") || "";
  const page = String(document.body?.dataset?.page || "");

  state.urgentMode = Boolean(context?.urgentMode);
  state.filters = {
    ...DEFAULT_FILTERS,
    ...((context && context.filters) || {}),
  };
  state.filters.search = "";

  if (personIdInQuery) {
    if (!params.get("personId") && params.get("personld")) {
      const nextUrl = new URL(window.location.href);
      nextUrl.searchParams.set("personId", personIdInQuery);
      nextUrl.searchParams.delete("personld");
      window.history.replaceState({}, "", nextUrl);
    }
    saveNavigationContext({ personId: personIdInQuery, filters: state.filters });
    return;
  }

  // Regle UX: si aucune personne n'est explicitement selectionnee sur la vue d'ensemble,
  // on ne doit pas pre-remplir automatiquement la fiche personne avec un ancien contexte.
  if (page === "overview") {
    saveNavigationContext({ personId: "" });
    return;
  }

  const storedPersonId = String((context && context.personId) || "");
  if (storedPersonId) {
    setCurrentPersonId(storedPersonId, "replace");
  }
}

function getCurrentPersonId() {
  const params = new URLSearchParams(window.location.search);
  const personIdInQuery = params.get("personId") || params.get("personld") || "";
  if (personIdInQuery) {
    if (!params.get("personId") && params.get("personld")) {
      const nextUrl = new URL(window.location.href);
      nextUrl.searchParams.set("personId", personIdInQuery);
      nextUrl.searchParams.delete("personld");
      window.history.replaceState({}, "", nextUrl);
    }
    return personIdInQuery;
  }
  const context = getStoredNavigationContext();
  return String(context?.personId || "");
}

function getSheetTargetPersonId() {
  return state.currentSheetPersonId || getCurrentPersonId();
}

function setCurrentPersonId(personId, mode = "replace") {
  const nextUrl = new URL(window.location.href);
  if (personId) {
    nextUrl.searchParams.set("personId", personId);
  } else {
    nextUrl.searchParams.delete("personId");
  }
  if (mode === "push") {
    window.history.pushState({}, "", nextUrl);
  } else {
    window.history.replaceState({}, "", nextUrl);
  }
  saveNavigationContext({ personId: personId || "" });
}

function performPageNavigation(url, mode = "href") {
  if (!url) {
    return;
  }
  if (mode === "replace") {
    window.location.replace(url);
    return;
  }
  window.location.href = url;
}

function capturePersonSheetDraftToState() {
  const form = document.getElementById("person-sheet-form");
  const person = getCurrentPerson();
  if (!(form instanceof HTMLFormElement) || !person) {
    return false;
  }

  const formData = new FormData(form);
  const draft = {
    nom: normalizeText(formData.get("sheetNom")),
    prenom: normalizeText(formData.get("sheetPrenom")),
    fonction: normalizeText(formData.get("sheetFonction")),
    sitesAffectation: readSelectedSites(form, "sheet"),
    typePersonnel: normalizeText(formData.get("sheetTypePersonnel")),
    typeContrat: normalizeText(formData.get("sheetTypeContrat")),
    dateEntree: normalizeDateString(formData.get("sheetDateEntree")),
    email: String(formData.get("sheetEmail") || "").trim(),
    phoneMobile: String(formData.get("sheetPhoneMobile") || "").trim(),
    dateSortiePrevue: normalizeDateString(formData.get("sheetDateSortiePrevue")),
    dateSortieReelle: normalizeDateString(formData.get("sheetDateSortieReelle")),
  };

  const changed =
    person.nom !== draft.nom ||
    person.prenom !== draft.prenom ||
    normalizeText(person.fonction) !== draft.fonction ||
    !haveSameSites(getPersonSites(person), draft.sitesAffectation) ||
    normalizeText(person.typePersonnel) !== draft.typePersonnel ||
    normalizeText(person.typeContrat) !== draft.typeContrat ||
    String(person.dateEntree || "") !== draft.dateEntree ||
    String(person.email || "") !== draft.email ||
    String(person.phoneMobile || "") !== draft.phoneMobile ||
    String(person.dateSortiePrevue || "") !== draft.dateSortiePrevue ||
    String(person.dateSortieReelle || "") !== draft.dateSortieReelle;

  if (!changed) {
    return false;
  }

  person.nom = draft.nom;
  person.prenom = draft.prenom;
  person.fonction = draft.fonction;
  person.sitesAffectation = draft.sitesAffectation;
  person.site = getPersonSiteLabel(person);
  person.typePersonnel = draft.typePersonnel;
  person.typeContrat = draft.typeContrat;
  person.dateEntree = draft.dateEntree;
  person.email = draft.email;
  person.phoneMobile = draft.phoneMobile;
  person.dateSortiePrevue = draft.dateSortiePrevue;
  person.dateSortieReelle = draft.dateSortieReelle;
  markDirty();
  return true;
}

function captureMobileSignatureSettingsDraftToState() {
  const form = document.getElementById("mobile-signature-settings-form");
  if (!(form instanceof HTMLFormElement) || !state.data?.meta) {
    return false;
  }
  const rawValue = String(form.elements.mobileSignatureBaseUrl?.value || "").trim();
  const normalized = normalizeMobileSignatureBaseUrl(rawValue);
  if (rawValue && !normalized) {
    return false;
  }
  const currentValue = String(state.data.meta.signatureMobileBaseUrl || "");
  if (currentValue === normalized) {
    return false;
  }
  state.data.meta.signatureMobileBaseUrl = normalized;
  state.mobileSignatureNetworkInfo = null;
  markDirty();
  return true;
}

function capturePendingEditsBeforeNavigation() {
  return capturePersonSheetDraftToState() || captureMobileSignatureSettingsDraftToState();
}

function runAutoSaveBeforeNavigation() {
  if (!state.isDirty || !state.data) {
    return Promise.resolve(false);
  }
  if (state.autoSaveInFlightPromise) {
    return state.autoSaveInFlightPromise;
  }
  state.autoSaveInFlightPromise = saveDataToFile({
    silent: true,
    reloadAfter: false,
    promptDownload: false,
    successText: "DONNEES SAUVEGARDEES",
  })
    .catch(() => undefined)
    .finally(() => {
      state.autoSaveInFlightPromise = null;
    });
  return state.autoSaveInFlightPromise.then(() => !state.isDirty);
}

function scheduleBackgroundAutoSave() {
  if (!state.isDirty || !state.data) {
    return;
  }
  if (state.autoSaveTimerId) {
    window.clearTimeout(state.autoSaveTimerId);
    state.autoSaveTimerId = 0;
  }
  state.autoSaveTimerId = window.setTimeout(() => {
    state.autoSaveTimerId = 0;
    runAutoSaveBeforeNavigation();
  }, 280);
}

function navigateWithAutoSave(url, mode = "href") {
  if (!url) {
    return;
  }
  capturePendingEditsBeforeNavigation();
  if (!state.isDirty) {
    performPageNavigation(url, mode);
    return;
  }
  let navigationDone = false;
  const navigateNow = () => {
    if (navigationDone) {
      return;
    }
    navigationDone = true;
    performPageNavigation(url, mode);
  };
  runAutoSaveBeforeNavigation()
    .then((saved) => {
      if (saved || !state.isDirty) {
        navigateNow();
        return;
      }
      showDataStatus("NAVIGATION BLOQUEE : SAUVEGARDE AUTOMATIQUE IMPOSSIBLE");
    })
    .catch((error) => {
      console.error(error);
      showDataStatus("NAVIGATION BLOQUEE : SAUVEGARDE AUTOMATIQUE IMPOSSIBLE");
    });
}

function openPersonSheet(personId) {
  const normalizedId = String(personId || "");
  if (!normalizedId) {
    return;
  }
  setCurrentPersonId(normalizedId, "replace");
  navigateWithAutoSave(`fiche-personne.html?personId=${normalizedId}`);
}

function openPersonSheetEffectEditor(personId, effectId) {
  const normalizedPersonId = String(personId || "");
  const normalizedEffectId = String(effectId || "");
  if (!normalizedPersonId || !normalizedEffectId) {
    return;
  }
  try {
    sessionStorage.setItem(
      EFFECT_EDIT_CONTEXT_STORAGE_KEY,
      JSON.stringify({
        personId: normalizedPersonId,
        effectId: normalizedEffectId,
      })
    );
  } catch (error) {
    // Ignore storage errors.
  }
  setCurrentPersonId(normalizedPersonId, "replace");
  navigateWithAutoSave(
    `fiche-personne.html?personId=${encodeURIComponent(normalizedPersonId)}&editEffectId=${encodeURIComponent(normalizedEffectId)}`
  );
}

function getRequestedEditEffectContext() {
  const nextUrl = new URL(window.location.href);
  const urlPersonId = String(nextUrl.searchParams.get("personId") || nextUrl.searchParams.get("personld") || "");
  const urlEffectId = String(nextUrl.searchParams.get("editEffectId") || "");
  if (urlPersonId && urlEffectId) {
    nextUrl.searchParams.delete("editEffectId");
    window.history.replaceState({}, "", nextUrl);
    return { personId: urlPersonId, effectId: urlEffectId };
  }

  const fallback = getRequestedEditEffectContextFromStorage();
  return fallback;
}

function getRequestedEditEffectContextFromStorage() {
  try {
    const raw = sessionStorage.getItem(EFFECT_EDIT_CONTEXT_STORAGE_KEY);
    if (!raw) {
      return { personId: "", effectId: "" };
    }
    const parsed = JSON.parse(raw);
    const personId = String(parsed?.personId || "");
    const effectId = String(parsed?.effectId || "");
    return { personId, effectId };
  } catch (error) {
    return { personId: "", effectId: "" };
  } finally {
    try {
      sessionStorage.removeItem(EFFECT_EDIT_CONTEXT_STORAGE_KEY);
    } catch (error) {
      // Ignore storage errors.
    }
  }
}

function consumeRequestedEditEffectId() {
  const context = getRequestedEditEffectContext();
  return context.effectId;
}

function getRequestedEditPersonId() {
  const context = getRequestedEditEffectContext();
  return context.personId;
}

function getCurrentMobileSignatureToken() {
  return new URLSearchParams(window.location.search).get("token") || "";
}

function getCurrentMobileSignatureDocType() {
  return new URLSearchParams(window.location.search).get("docType") || "";
}

function normalizeMobileSignatureSigner(value) {
  return normalizeText(value) === "REPRESENTANT" ? "representant" : "personnel";
}

function getCurrentMobileSignatureSigner() {
  return normalizeMobileSignatureSigner(new URLSearchParams(window.location.search).get("signer") || "");
}

function generateMobileSignatureToken() {
  return `SIG-${Date.now()}-${Math.random().toString(36).slice(2, 12).toUpperCase()}`;
}

function findMobileSignatureRequestByToken(token) {
  return (state.data?.demandesSignatureMobile || []).find((entry) => entry.token === token) || null;
}

function getMobileSignatureIdentityFromUrl() {
  const params = new URLSearchParams(window.location.search);
  return {
    personNom: normalizeText(params.get("personNom") || ""),
    personPrenom: normalizeText(params.get("personPrenom") || ""),
    personSite: normalizeText(params.get("personSite") || ""),
    personTypePersonnel: normalizeText(params.get("personTypePersonnel") || ""),
    personTypeContrat: normalizeText(params.get("personTypeContrat") || ""),
    representativeNom: normalizeText(params.get("representativeNom") || ""),
    representativeFonction: normalizeText(params.get("representativeFonction") || ""),
  };
}

function applyMobileSignatureIdentityFromUrl(request) {
  if (!request) {
    return request;
  }
  const identity = getMobileSignatureIdentityFromUrl();
  Object.entries(identity).forEach(([key, value]) => {
    if (value && !String(request[key] || "").trim()) {
      request[key] = value;
    }
  });
  return request;
}

function buildFallbackMobileSignatureRequestFromUrl(token) {
  const normalizedToken = String(token || "").trim();
  if (!normalizedToken) {
    return null;
  }
  const params = new URLSearchParams(window.location.search);
  const personId = String(params.get("personId") || params.get("personld") || "").trim();
  const docType = normalizeText(params.get("docType") || "");
  const signer = normalizeMobileSignatureSigner(params.get("signer") || "");
  if (!personId || !docType) {
    return null;
  }
  const createdAt = new Date();
  const expiresAt = new Date(createdAt.getTime() + MOBILE_SIGNATURE_REQUEST_TTL_MS);
  const request = {
    id: "DSM-FALLBACK",
    token: normalizedToken,
    personId,
    docType,
    signer: signer.toUpperCase(),
    createdAt: createdAt.toISOString(),
    expiresAt: expiresAt.toISOString(),
    status: "EN ATTENTE",
    validatedAt: "",
    fallbackFromUrl: true,
  };
  return applyMobileSignatureIdentityFromUrl(request);
}

function getActiveMobileSignatureRequestContext(personId, docType) {
  const normalizedPersonId = String(personId || "");
  const normalizedDocType = normalizeText(docType);
  const cacheKey = `${String(state.supabaseRevision || "r0")}|${String(state.localMutationTick || 0)}|${normalizedPersonId}|${normalizedDocType}`;
  const cacheState = state.mobileSignatureRequestContextCache || (state.mobileSignatureRequestContextCache = new Map());
  const now = Date.now();
  const cachedEntry = cacheState.get(cacheKey);
  if (cachedEntry && cachedEntry.expiresAt > now && cachedEntry.context) {
    return cachedEntry.context;
  }

  cleanupExpiredMobileSignatureRequests();
  const requestContext = {
    personnel: null,
    representant: null,
    hasAny: false,
  };
  let nearestExpiration = now + 1000;
  for (const entry of state.data?.demandesSignatureMobile || []) {
    const signer = normalizeMobileSignatureSigner(entry?.signer || "");
    const entryExpiresAt = Date.parse(entry?.expiresAt || "");
    if (
      String(entry?.personId || "") !== normalizedPersonId ||
      normalizeText(entry?.docType || "") !== normalizedDocType ||
      !signer ||
      entry?.status !== "EN ATTENTE" ||
      !Number.isFinite(entryExpiresAt) ||
      entryExpiresAt <= now
    ) {
      continue;
    }
    if (entryExpiresAt > now) {
      nearestExpiration = Math.min(nearestExpiration, entryExpiresAt);
    }
    requestContext.hasAny = true;
    if (signer === "personnel" && !requestContext.personnel) {
      requestContext.personnel = entry;
      continue;
    }
    if (signer === "representant" && !requestContext.representant) {
      requestContext.representant = entry;
    }
  }

  cacheState.set(cacheKey, {
    context: requestContext,
    expiresAt: Math.min(now + 5000, nearestExpiration),
    createdAt: now,
  });

  if (cacheState.size > 40) {
    const oldestKeys = Array.from(cacheState.entries())
      .sort((left, right) => left[1].createdAt - right[1].createdAt)
      .slice(0, Math.max(0, cacheState.size - 40))
      .map((entry) => entry[0]);
    oldestKeys.forEach((key) => cacheState.delete(key));
  }

  return requestContext;
}

function getActiveMobileSignatureRequest(personId, docType, signer = "personnel") {
  const requestContext = getActiveMobileSignatureRequestContext(personId, docType);
  const normalizedSigner = normalizeMobileSignatureSigner(signer);
  if (normalizedSigner === "representant") {
    return requestContext.representant;
  }
  return requestContext.personnel;
}

function hasActiveMobileSignatureRequest(personId, docType) {
  return getActiveMobileSignatureRequestContext(personId, docType).hasAny;
}

function createMobileSignatureRequest(personId, docType, signer = "personnel") {
  const createdAt = new Date();
  const expiresAt = new Date(createdAt.getTime() + MOBILE_SIGNATURE_REQUEST_TTL_MS);
  const normalizedSigner = normalizeMobileSignatureSigner(signer);
  const person = (state.data?.personnes || []).find((entry) => String(entry?.id || "") === String(personId || "")) || null;
  const normalizedDocType = normalizeText(docType) === "EXIT" ? "exit" : "arrival";
  const representative = person ? getRepresentativeInfo(person, normalizedDocType) : null;
  const request = {
    id: getNextId("DSM", state.data?.demandesSignatureMobile || []),
    token: generateMobileSignatureToken(),
    personId: String(personId || ""),
    docType: normalizeText(docType),
    signer: normalizedSigner.toUpperCase(),
    personNom: normalizeText(person?.nom || ""),
    personPrenom: normalizeText(person?.prenom || ""),
    personSite: normalizeText(person?.site || ""),
    personTypePersonnel: normalizeText(person?.typePersonnel || ""),
    personTypeContrat: normalizeText(person?.typeContrat || ""),
    representativeNom: normalizeText(representative?.nom || ""),
    representativeFonction: normalizeText(representative?.fonction || ""),
    createdAt: createdAt.toISOString(),
    expiresAt: expiresAt.toISOString(),
    status: "EN ATTENTE",
    validatedAt: "",
  };
  state.data.demandesSignatureMobile.push(request);
  return request;
}

function enrichMobileSignatureRequestIdentity(request) {
  if (!request || !state.data) {
    return false;
  }
  const person = (state.data.personnes || []).find((entry) => String(entry?.id || "") === String(request.personId || "")) || null;
  if (!person) {
    return false;
  }
  const normalizedDocType = normalizeText(request.docType) === "EXIT" ? "exit" : "arrival";
  const representative = getRepresentativeInfo(person, normalizedDocType);
  const nextValues = {
    personNom: normalizeText(person.nom || ""),
    personPrenom: normalizeText(person.prenom || ""),
    personSite: normalizeText(person.site || ""),
    personTypePersonnel: normalizeText(person.typePersonnel || ""),
    personTypeContrat: normalizeText(person.typeContrat || ""),
    representativeNom: normalizeText(representative?.nom || ""),
    representativeFonction: normalizeText(representative?.fonction || ""),
  };
  let changed = false;
  Object.entries(nextValues).forEach(([key, value]) => {
    if (value && String(request[key] || "") !== value) {
      request[key] = value;
      changed = true;
    }
  });
  return changed;
}

function getMobileSignaturePageUrl(request, options = {}) {
  const includeSessionBridge = options?.includeSessionBridge !== false;
  const docType = normalizeText(request?.docType) === "EXIT" ? "exit" : "arrival";
  const signer = normalizeMobileSignatureSigner(request?.signer || "");
  const params = new URLSearchParams({
    personId: String(request?.personId || ""),
    docType,
    token: String(request?.token || ""),
    signer,
  });
  [
    "personNom",
    "personPrenom",
    "personSite",
    "personTypePersonnel",
    "personTypeContrat",
    "representativeNom",
    "representativeFonction",
  ].forEach((key) => {
    const value = String(request?.[key] || "").trim();
    if (value) {
      params.set(key, value);
    }
  });
  const relativeUrl = `signature-mobile.html?${params.toString()}`;
  if (includeSessionBridge) {
    return appendSupabaseSessionBridgeParams(relativeUrl, options?.sessionBridgeOptions || {});
  }
  return relativeUrl;
}

async function getMobileSignatureBaseUrl() {
  const configuredBaseUrl = getConfiguredMobileSignatureBaseUrl();

  let currentRuntimeBaseUrl = "";
  try {
    const runtimeUrl = new URL(window.location.href || "", window.location.origin);
    if (/\.[a-z0-9]+$/i.test(runtimeUrl.pathname)) {
      runtimeUrl.pathname = runtimeUrl.pathname.replace(/[^/]+$/, "");
    }
    if (!runtimeUrl.pathname.endsWith("/")) {
      runtimeUrl.pathname = `${runtimeUrl.pathname}/`;
    }
    runtimeUrl.search = "";
    runtimeUrl.hash = "";
    currentRuntimeBaseUrl = normalizeHttpUrl(runtimeUrl.href);
  } catch (error) {
    currentRuntimeBaseUrl = normalizeHttpUrl(window.location.origin || "");
  }

  if (configuredBaseUrl) {
    if (isLikelyLocalUrl(currentRuntimeBaseUrl)) {
      try {
        const modeResponse = await fetch(`/api/network-info?ts=${Date.now()}`, { cache: "no-store" });
        if (modeResponse.ok) {
          const modeInfo = await modeResponse.json().catch(() => ({}));
          state.mobileSignatureNetworkInfo = modeInfo;
          if (String(modeInfo?.mode || "").toLowerCase() === "local" && modeInfo?.preferredUrl) {
            return String(modeInfo.preferredUrl);
          }
          if (String(modeInfo?.mode || "").toLowerCase() === "supabase" && modeInfo?.publicUrl) {
            return String(modeInfo.publicUrl);
          }
        }
      } catch (error) {
        // ignore network info failures and keep configured fallback
      }
    }
    if (!currentRuntimeBaseUrl || isLikelyLocalUrl(currentRuntimeBaseUrl)) {
      return configuredBaseUrl;
    }
    if (areSameHost(configuredBaseUrl, currentRuntimeBaseUrl)) {
      return configuredBaseUrl;
    }
  }

  if (currentRuntimeBaseUrl && !isLikelyLocalUrl(currentRuntimeBaseUrl)) {
    return currentRuntimeBaseUrl;
  }

  if (state.mobileSignatureNetworkInfo?.preferredUrl) {
    return state.mobileSignatureNetworkInfo.preferredUrl;
  }

  try {
    const response = await fetch(`/api/network-info?ts=${Date.now()}`, { cache: "no-store" });
    if (!response.ok) {
      throw new Error("NETWORK INFO UNAVAILABLE");
    }
    const info = await response.json();
    state.mobileSignatureNetworkInfo = info;
    return String(info?.preferredUrl || window.location.origin);
  } catch (error) {
    state.mobileSignatureNetworkInfo = {
      preferredUrl: window.location.origin,
      lanUrls: [],
    };
    return window.location.origin;
  }
}

function getQrProviderUrls(absoluteUrl) {
  if (!absoluteUrl) {
    return [];
  }
  const encoded = encodeURIComponent(String(absoluteUrl || ""));
  const qrSize = "420x420";
  const providers = [
    `/api/qr?size=420&margin=2&text=${encoded}`,
    `https://quickchart.io/qr?size=420&margin=2&ecLevel=L&dark=000000&light=ffffff&text=${encoded}`,
    `https://api.qrserver.com/v1/create-qr-code/?size=${qrSize}&margin=2&ecc=L&color=000000&bgcolor=ffffff&data=${encoded}`,
    `https://chart.googleapis.com/chart?cht=qr&chs=${qrSize}&chld=L|2&choe=UTF-8&chl=${encoded}`,
  ];
  return providers;
}

function getQrFriendlySignatureUrl(absoluteUrl) {
  const raw = String(absoluteUrl || "").trim();
  if (!raw) return "";
  try {
    const parsed = new URL(raw);
    // Keep mobile auth bridge required for signature save.
    // Do not strip `sbat/sbrt/sbea`: they are needed on phone to authorize save.
    return parsed.toString();
  } catch (error) {
    return raw;
  }
}

async function getAbsoluteMobileSignatureUrl(request, options = {}) {
  if (!request) {
    return "";
  }
  const relativeUrl = getMobileSignaturePageUrl(request, options);
  const baseUrl = await getMobileSignatureBaseUrl();
  let resolvedBaseUrl = String(baseUrl || window.location.origin || "").trim();
  try {
    const parsed = new URL(resolvedBaseUrl, window.location.origin);
    if (/\.[a-z0-9]+$/i.test(parsed.pathname)) {
      parsed.pathname = parsed.pathname.replace(/[^/]+$/, "");
    }
    if (!parsed.pathname.endsWith("/")) {
      parsed.pathname = `${parsed.pathname}/`;
    }
    parsed.search = "";
    parsed.hash = "";
    resolvedBaseUrl = parsed.href;
  } catch (error) {
    resolvedBaseUrl = `${resolvedBaseUrl.replace(/\/$/, "").replace(/\/[^/]+\.[a-z0-9]+$/i, "")}/`;
  }
  return new URL(relativeUrl, resolvedBaseUrl).href;
}

async function fillMobileSignatureShareLink(request) {
  const wrapper = document.getElementById("mobile-signature-share");
  const input = document.getElementById("mobile-signature-share-url");
  const copyButton = document.getElementById("mobile-signature-copy-link");
  const qrWrapper = document.getElementById("mobile-signature-share-qr");
  const qrImage = document.getElementById("mobile-signature-share-qr-image");
  const reachabilityHintNode = document.getElementById("mobile-signature-share-network-hint");

  if (!wrapper || !input || !copyButton) {
    return;
  }

  if (!request) {
    wrapper.hidden = true;
    input.value = "";
    copyButton.disabled = true;
    if (qrWrapper) {
      qrWrapper.hidden = true;
    }
    if (qrImage) {
      qrImage.removeAttribute("src");
    }
    if (reachabilityHintNode) {
      reachabilityHintNode.textContent = "";
    }
    return;
  }

  // Keep QR links camera-friendly while preserving silent auth: refresh token only.
  const absoluteUrl = await getAbsoluteMobileSignatureUrl(request, {
    includeSessionBridge: true,
    sessionBridgeOptions: {
      // Include both tokens to avoid refresh-only failures on mobile devices.
      includeAccessToken: true,
      includeRefreshToken: true,
      includeExpiresAt: true,
    },
  });
  const qrCompactUrl = await getAbsoluteMobileSignatureUrl(request, {
    includeSessionBridge: true,
    sessionBridgeOptions: {
      // QR readability: keep refresh flow only for shorter payload.
      includeAccessToken: false,
      includeRefreshToken: true,
      includeExpiresAt: true,
    },
  });

  wrapper.hidden = false;
  input.value = absoluteUrl;
  copyButton.disabled = false;
  if (reachabilityHintNode) {
    reachabilityHintNode.textContent = getMobileSignatureReachabilityHint(absoluteUrl);
  }
  if (qrWrapper && qrImage) {
    const qrTargetUrl = getQrFriendlySignatureUrl(qrCompactUrl || absoluteUrl);
    const providerUrls = getQrProviderUrls(qrTargetUrl);
    if (!providerUrls || !providerUrls.length) {
      qrWrapper.hidden = true;
      if (reachabilityHintNode) {
        const hint = getMobileSignatureReachabilityHint(absoluteUrl);
        reachabilityHintNode.textContent = `${hint} QR INDISPONIBLE: UTILISER LE LIEN DIRECT.`;
      }
      qrImage.removeAttribute("src");
      return;
    }
    let providerIndex = 0;
    qrImage.referrerPolicy = "no-referrer";
    qrImage.decoding = "async";
    qrImage.alt = "QR CODE DE SIGNATURE MOBILE";
    qrWrapper.hidden = false;
    qrImage.src = providerUrls[providerIndex];
    qrImage.onerror = () => {
      providerIndex += 1;
      if (providerIndex < providerUrls.length) {
        qrImage.src = providerUrls[providerIndex];
        return;
      }
      qrWrapper.hidden = true;
      if (reachabilityHintNode) {
        const hint = getMobileSignatureReachabilityHint(absoluteUrl);
        reachabilityHintNode.textContent = `${hint} Le QR n’a pas pu être généré. Utilise le lien direct ci-dessous.`;
      }
    };
    qrImage.onload = () => {
      qrWrapper.hidden = false;
    };
  }
  copyButton.onclick = async () => {
    try {
      if (navigator.clipboard?.writeText) {
        await navigator.clipboard.writeText(absoluteUrl);
      } else {
        input.focus();
        input.select();
        document.execCommand("copy");
      }
      showDataStatus("LIEN DE SIGNATURE COPIE");
    } catch (error) {
      input.focus();
      input.select();
      showDataStatus("COPIE MANUELLE DU LIEN");
    }
  };
}

function getMobileSignatureLinkNodeId(docType, signer = "personnel") {
  return `${docType}-mobile-signature-link-${normalizeMobileSignatureSigner(signer)}`;
}

function renderMobileSignatureLink(docType, signer, absoluteUrl) {
  const normalizedSigner = normalizeMobileSignatureSigner(signer);
  const legacyNodeId = `${docType}-mobile-signature-link`;
  const linkNode =
    document.getElementById(getMobileSignatureLinkNodeId(docType, normalizedSigner)) ||
    (normalizedSigner === "personnel" ? document.getElementById(legacyNodeId) : null);
  if (!linkNode) {
    return;
  }
  if (isPdfRenderMode()) {
    linkNode.hidden = true;
    linkNode.innerHTML = "";
    return;
  }
  if (!absoluteUrl) {
    linkNode.hidden = true;
    linkNode.innerHTML = "";
    return;
  }
  const hint = getMobileSignatureReachabilityHint(absoluteUrl);
  let linkLabel = absoluteUrl;
  try {
    const parsed = new URL(absoluteUrl);
    const shortToken = String(parsed.searchParams.get("token") || "").slice(0, 10);
    linkLabel = `${parsed.origin}${parsed.pathname}${shortToken ? `?token=${shortToken}...` : ""}`;
  } catch (error) {
    linkLabel = absoluteUrl.length > 96 ? `${absoluteUrl.slice(0, 96)}...` : absoluteUrl;
  }
  linkNode.hidden = false;
  linkNode.innerHTML = `
    <span>LIEN TELEPHONE :</span>
    <a href="${escapeHtml(absoluteUrl)}" target="_blank" rel="noopener">${escapeHtml(linkLabel)}</a>${hint ? `<small>${escapeHtml(hint)}</small>` : ""}
  `;
}

async function syncDocumentMobileSignatureLink(docType, personId, signer = "personnel") {
  const normalizedSigner = normalizeMobileSignatureSigner(signer);
  if (!personId || !state.data) {
    renderMobileSignatureLink(docType, normalizedSigner, "");
    return;
  }
  const request = getActiveMobileSignatureRequest(personId, docType, normalizedSigner);
  if (!request) {
    renderMobileSignatureLink(docType, normalizedSigner, "");
    return;
  }
  const absoluteUrl = await getAbsoluteMobileSignatureUrl(request);
  renderMobileSignatureLink(docType, normalizedSigner, absoluteUrl);
}

async function syncDocumentMobileSignatureLinks(docType, personId) {
  if (!personId || !state.data) {
    renderMobileSignatureLink(docType, "personnel", "");
    renderMobileSignatureLink(docType, "representant", "");
    return;
  }
  const { personnel, representant } = getActiveMobileSignatureRequestContext(personId, docType);
  const personnelUrl = personnel ? await getAbsoluteMobileSignatureUrl(personnel) : "";
  const representantUrl = representant ? await getAbsoluteMobileSignatureUrl(representant) : "";
  renderMobileSignatureLink(docType, "personnel", personnelUrl);
  renderMobileSignatureLink(docType, "representant", representantUrl);
}

async function pushMobileSignatureRequestToHosted(request) {
  if (getDataBackendMode() !== "LOCAL_API") {
    return true;
  }
  const response = await fetch("/api/sync/send-mobile-signature-request", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      personId: String(request?.personId || ""),
      docType: normalizeText(request?.docType || "") === "EXIT" ? "exit" : "arrival",
      signer: normalizeMobileSignatureSigner(request?.signer || ""),
      token: String(request?.token || ""),
    }),
    cache: "no-store",
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok || !payload?.ok) {
    const detail = String(payload?.detail || payload?.error || `HTTP_${response.status}`);
    throw new Error(detail || "MOBILE_SIGNATURE_REQUEST_PUSH_FAILED");
  }
  return true;
}

async function openMobileSignatureRequest(docType, personId, signer = "personnel") {
  if (!state.data) {
    showDataStatus("DONNEES NON CHARGEES");
    return;
  }
  const normalizedSigner = normalizeMobileSignatureSigner(signer);
  const person = (state.data?.personnes || []).find((entry) => String(entry?.id || "") === String(personId || "")) || null;
  const representativeDraftChanged =
    normalizedSigner === "representant" ? captureRepresentativeIdentityDraftToState(docType, person) : false;
  if (normalizedSigner === "representant" && !hasRepresentativeIdentityForDocument(docType)) {
    showDataStatus("IDENTITE DU REPRESENTANT OBLIGATOIRE AVANT VALIDATION");
    window.alert("VOUS DEVEZ IDENTIFIER L'IDENTITE DU REPRESENTANT DE L'ETABLISSEMENT POUR VALIDATION.");
    updateRepresentativeSignatureActionState(docType);
    return;
  }

  const signatureWindow = window.open("", "_blank");
  if (!signatureWindow) {
    showDataStatus("AUTORISER L'OUVERTURE DE LA PAGE DE SIGNATURE DANS LE NAVIGATEUR");
    window.alert("LE NAVIGATEUR A BLOQUE L'OUVERTURE DE LA PAGE DE SIGNATURE. AUTORISEZ LES POPUPS POUR DOTATIONS.");
    return;
  }
  try {
    signatureWindow.opener = null;
    signatureWindow.document.title = "SIGNATURE MOBILE";
    signatureWindow.document.body.style.fontFamily = "Arial, sans-serif";
    signatureWindow.document.body.style.padding = "24px";
    signatureWindow.document.body.textContent = "PREPARATION DE LA PAGE DE SIGNATURE...";
  } catch (error) {
    // noop: some browsers restrict access even to the just-opened blank window.
  }
  showMobileSignatureRecoveryModal("PAGE DE SIGNATURE EN PREPARATION. EN ATTENTE DU RETOUR DE SIGNATURE.");

  let request = getActiveMobileSignatureRequest(personId, docType, normalizedSigner);
  if (!request) {
    request = createMobileSignatureRequest(personId, docType, normalizedSigner);
    markDirty();
    await saveDataToFile({ silent: true, reloadAfter: false, autoPushHosted: false });
  } else if (representativeDraftChanged || enrichMobileSignatureRequestIdentity(request)) {
    markDirty();
    enrichMobileSignatureRequestIdentity(request);
    await saveDataToFile({ silent: true, reloadAfter: false, autoPushHosted: false });
  }
  if (getDataBackendMode() === "LOCAL_API") {
    if (state.isDirty) {
      await saveDataToFile({ silent: true, reloadAfter: false, autoPushHosted: false });
    }
    if (!state.isDirty && !isLocalHostedSyncDisabled()) {
      try {
        await pushMobileSignatureRequestToHosted(request);
      } catch (error) {
        const message = String(error?.message || error || "MOBILE_SIGNATURE_REQUEST_PUSH_FAILED");
        console.warn("Demande signature mobile non poussee dans app_state heberge", message);
        showDataStatus("SIGNATURE MOBILE : OUVERTURE PAR LIEN DIRECT, SANS PUSH APP_STATE");
      }
    }
  }

  const absoluteUrl = await getAbsoluteMobileSignatureUrl(request);
  try {
    signatureWindow.location.replace(absoluteUrl);
    signatureWindow.focus();
  } catch (error) {
    const fallbackWindow = window.open(absoluteUrl, "_blank", "noopener");
    if (!fallbackWindow) {
      showDataStatus("PAGE DE SIGNATURE PRETE - OUVERTURE BLOQUEE PAR LE NAVIGATEUR");
      renderMobileSignatureLink(docType, normalizedSigner, absoluteUrl);
      return;
    }
  }
  renderMobileSignatureLink(docType, normalizedSigner, absoluteUrl);

  showDataStatus("PAGE DE SIGNATURE MOBILE OUVERTE");
  showMobileSignatureRecoveryModal("PAGE DE SIGNATURE OUVERTE. EN ATTENTE DU RETOUR DE SIGNATURE.");
  syncMobileSignaturePolling();
}



async function refreshHostedSyncStatusOnLocalOpen() {
  if (getDataBackendMode() !== "LOCAL_API") return;
  if (state.isDirty) {
    state.hostedSyncState = "pending";
    renderHostedSyncUi();
    return;
  }
  try {
    const response = await fetch("/api/local/check-before-send", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
    });
    const payload = await response.json().catch(() => ({}));
    if (response.ok && payload?.ok) {
      state.hostedSyncState = "up_to_date";
      state.hostedSyncDetails = "";
    } else {
      state.hostedSyncState = payload?.supabaseReachable === false ? "inaccessible" : "blocked";
      state.hostedSyncDetails = JSON.stringify(payload?.technicalDetails || payload || {}, null, 2);
    }
  } catch (error) {
    state.hostedSyncState = "inaccessible";
    state.hostedSyncDetails = String(error?.message || error || "");
  }
  renderHostedSyncUi();
}
function isRescueModalRequest() {
  try {
    const url = new URL(window.location.href);
    const path = String(url.pathname || "").replace(/\/+$/, "") || "/";
    return path === "/rescue" || url.searchParams.has("rescue");
  } catch {
    return false;
  }
}

async function openRescueModalFromCurrentCheck() {
  if (typeof openRescueActionModal !== "function") return;
  let hostedState = "blocked";
  let detailsRaw = "";
  try {
    const response = await fetch("/api/local/check-before-send", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
    });
    const payload = await response.json().catch(() => ({}));
    hostedState = payload?.ok ? "pending" : (payload?.supabaseReachable === false ? "inaccessible" : "blocked");
    detailsRaw = JSON.stringify(payload?.technicalDetails || payload || {}, null, 2);
    state.hostedSyncState = hostedState;
    state.hostedSyncDetails = detailsRaw;
    if (response.ok && payload?.ok) {
      state.hostedSyncState = "up_to_date";
      state.hostedSyncDetails = "";
      renderHostedSyncUi();
      window.location.replace(`${window.location.origin}/`);
      return;
    }
  } catch (error) {
    hostedState = "blocked";
    detailsRaw = JSON.stringify({ errors: [`LOCAL_CHECK_FAILED:${String(error?.message || error || "")}`] }, null, 2);
  }
  openRescueActionModal({ detailsRaw, hostedState });
}

function scheduleRescueModalFromUrl() {
  if (!isRescueModalRequest()) return;
  if (getDataBackendMode() !== "LOCAL_API") return;
  window.setTimeout(() => {
    void openRescueModalFromCurrentCheck();
  }, 350);
}
async function loadData() {
  bindBeforeUnloadGuard();
  bindPdfModalCleanup();
  reorderOverviewSearchBlock();
  importSupabaseSessionFromUrlIfPresent();
  restoreNavigationContext();
  await enforceUiLoginOnEachOpen();
  clearSearchInputsOnInitialLoad();
  bindSearchClearOnBrowserEvents();
  applyActiveNav();
  bindHistoryNavigation();
  bindAutoSaveOnNavigation();
  bindGlobalResetSelectionClear();
  bindFilterResetHighlights();
  bindGlobalShortcuts();
  bindArchiveFilterDelegation();
  bindArchiveFilterValueWatcher();
  bindLoadButton();
  bindSaveButtons();
  bindDirtyFallbackTracking();
  bindPdfButtons();
  bindMobileSignatureButtons();
  bindOverviewUrgencyActions();
  bindOverviewControlExport();
  bindFilterForms();
  bindAddPersonForm();
  bindPersonSheetForm();
  bindEffectForm();
  bindReferenceListForms();
  bindReferenceEffectForm();
  bindReplacementCostForm();
  bindStockAdjustmentForm();
  bindRepresentativeSignatoryForm();
  bindMobileSignatureSettingsForm();
  bindMobileSignatureVisibilityPolling();
  bindRegisterButtonsAutoSave();
  bindReferenceFilters();
  bindOverviewAlertActions();
  bindArchiveFilterForm();
  bindSignatureCanvases();
  bindRepresentativeFields();
  if (getDataBackendMode() === "LOCAL_API") {
    console.info("MODE LOCAL DATA ACTIF");
  }
  renderRoleBadge();
  refreshCurrentUserRoleLabel();
  if (getDataBackendMode() === "LOCAL_API") {
    state.hostedSyncState = state.isDirty ? "pending" : "unknown";
    ensureHostedSyncUi();
    renderHostedSyncUi();
    maybeOpenRescueModalPreview();
  }

  const workingData = loadWorkingData();
  if (workingData) {
    state.data = workingData;
    migrateDataModel({ suppressDirty: true });
    state.isDirty = true;
    state.lastPersistedDataSignature = null;
    clearUndoStack();
    applyMeta();
    hydrateStaticLists();
    schedulePageRender();
    clearSearchInputsOnInitialLoad();
    showDataStatus("DONNEES EN COURS REPRISES - SAUVEGARDER POUR LES RENDRE DEFINITIVES");
    void updateBrowserStorageAlert();
    scheduleBackgroundAutoSave();
    scheduleRescueModalFromUrl();
    void refreshHostedSyncStatusOnLocalOpen();
    return;
  }

  await reloadData("OUVERTURE DES DONNEES...");
  scheduleRescueModalFromUrl();
  void refreshHostedSyncStatusOnLocalOpen();
}

function reorderOverviewSearchBlock() {
  if (document.body?.dataset?.page !== "overview") {
    return;
  }
  const container = document.querySelector(".overview-top-fixed");
  if (!(container instanceof HTMLElement)) {
    return;
  }
  const sections = Array.from(container.querySelectorAll(":scope > section.section"));
  const getHeading = (section) =>
    normalizeText(section.querySelector(".section__heading h3")?.textContent || "");
  const searchSection = sections.find((section) => getHeading(section) === "RECHERCHE ET FILTRES");
  const overviewSection = sections.find((section) => getHeading(section) === "VUE D'ENSEMBLE");
  if (!(searchSection instanceof HTMLElement) || !(overviewSection instanceof HTMLElement)) {
    return;
  }
  if (searchSection.compareDocumentPosition(overviewSection) & Node.DOCUMENT_POSITION_PRECEDING) {
    container.insertBefore(searchSection, overviewSection);
  }
}

function applyPdfModeFromQuery() {
  const params = new URLSearchParams(window.location.search);
  if (params.get("pdf") === "1") {
    document.body.dataset.pdfMode = "true";
    document.body.dataset.pdfLayoutLock = PDF_FORMAT_LOCK;
  }
}

function isPdfRenderMode() {
  if (document.body?.dataset?.pdfMode === "true") {
    return true;
  }
  return new URLSearchParams(window.location.search).get("pdf") === "1";
}

async function reloadData(statusText = "RECHARGEMENT DES DONNEES...") {
  showDataStatus(statusText);

  try {
    const previousSignatureValidationMap = new Map(state.previousSignatureValidationMap || []);
    const json = await fetchLatestDataSnapshot();
    state.data = json;
    migrateDataModel();
    clearWorkingData();
    state.isDirty = false;
    state.lastPersistedDataSignature = computeDataPersistenceSignature(state.data);
    clearUndoStack();
    applyMeta();
    hydrateStaticLists();
    schedulePageRender();
    refreshCurrentUserRoleLabel();
    clearSearchInputsOnInitialLoad();
    showDataStatus(
      getDataBackendMode() === "SUPABASE" ? "DONNEES SUPABASE CHARGEES" : "DONNEES LOCALES CHARGEES"
    );
    void updateBrowserStorageAlert();
    state.previousSignatureValidationMap = buildSignatureValidationMap(state.data);
    window.setTimeout(() => {
      notifyFullySignedDocumentsOnReload(previousSignatureValidationMap);
    }, 0);
    queueAutoGenerateSignedDocumentsPdfIfMissing();
  } catch (error) {
    console.error(error);
    state.supabaseRevision = null;
    state.data = null;
    resetUiWithoutData();
    if (getDataBackendMode() === "HOSTED_NO_BACKEND") {
      showDataStatus("CONFIGURATION SUPABASE INCOMPLETE");
    } else {
      showDataStatus("OUVRIR L'APPLICATION VIA LE SERVEUR LOCAL");
    }
  }
}

function applyMeta() {
  if (!state.data?.meta) {
    return;
  }

  document.title = state.data.meta.appTitle || document.title;

  document.querySelectorAll(".sidebar__title").forEach((node) => {
    node.textContent = state.data.meta.appTitle || node.textContent;
  });

  document.querySelectorAll(".sidebar__subtitle").forEach((node) => {
    node.textContent = state.data.meta.appSubtitle || node.textContent;
  });
}

function normalizeArchiveQualityForLegacyEntry(entry) {
  const quality = normalizePdfQualityStatus(entry?.pdfQualityStatus || "");
  const hasLegacyPath = Boolean(String(entry?.pdfPath || "").trim());
  const hasLocalPath = Boolean(String(entry?.localPath || "").trim());
  const hasOpenLocalUrl = Boolean(String(entry?.openLocalUrl || "").trim());
  const hasRemoteUrl = Boolean(String(entry?.openRemoteUrl || entry?.publicUrl || "").trim());
  if (quality === "UNKNOWN" && hasLegacyPath && !hasLocalPath && !hasOpenLocalUrl && !hasRemoteUrl) {
    return "LEGACY_ONLY";
  }
  return quality;
}

function migrateDataModel(options = {}) {
  const suppressDirty = options?.suppressDirty === true;
  if (!state.data) {
    return;
  }

  if (!state.data.listes) {
    state.data.listes = {};
  }
  if (!state.data.meta || typeof state.data.meta !== "object") {
    state.data.meta = {};
  }
  state.data.meta.signatureMobileBaseUrl = normalizeMobileSignatureBaseUrl(state.data.meta.signatureMobileBaseUrl || "");
  state.data.meta.storagePdfBucket = normalizeBucketName(
    state.data.meta.storagePdfBucket,
    DEFAULT_SUPABASE_PDF_BUCKET
  );
  state.data.meta.storageSignaturesBucket = normalizeBucketName(
    state.data.meta.storageSignaturesBucket,
    DEFAULT_SUPABASE_SIGNATURES_BUCKET
  );

  delete state.data.listes.services;

  if (!Array.isArray(state.data.listes.typesPersonnel)) {
    state.data.listes.typesPersonnel = [];
  }
  if (!Array.isArray(state.data.listes.sites)) {
    state.data.listes.sites = [];
  }
  if (!Array.isArray(state.data.listes.typesEffets)) {
    state.data.listes.typesEffets = [];
  }
  if (!Array.isArray(state.data.listes.typesContrats)) {
    state.data.listes.typesContrats = [];
  }
  if (!Array.isArray(state.data.listes.fonctions)) {
    state.data.listes.fonctions = [];
  }
  if (!Array.isArray(state.data.listes.causesRemplacement)) {
    state.data.listes.causesRemplacement = [];
  }
  if (!Array.isArray(state.data.listes.statutsObjetManuels)) {
    state.data.listes.statutsObjetManuels = [];
  }
  if (!Array.isArray(state.data.listes.coutsRemplacement)) {
    state.data.listes.coutsRemplacement = [];
  }
  if (!Array.isArray(state.data.listes.representantsSignataires)) {
    state.data.listes.representantsSignataires = [];
  }
  if (!Array.isArray(state.data.documentsArchives)) {
    state.data.documentsArchives = [];
  }
  if (!Array.isArray(state.data.stocksEffetsManuels)) {
    state.data.stocksEffetsManuels = [];
  }
  if (!Array.isArray(state.data.demandesSignatureMobile)) {
    state.data.demandesSignatureMobile = [];
  }

  state.data.listes.typesContrats = Array.from(new Set(state.data.listes.typesContrats.map(normalizeText))).filter(
    Boolean
  );
  state.data.listes.fonctions = state.data.listes.fonctions.map(normalizeFunctionLabel).filter(Boolean);

  state.data.listes.typesPersonnel = state.data.listes.typesPersonnel.map(normalizeText).filter(Boolean);
  state.data.listes.sites = Array.from(new Set(state.data.listes.sites.map(normalizeText))).filter(Boolean);
  state.data.listes.typesEffets = Array.from(new Set(state.data.listes.typesEffets.map(normalizeText))).filter(
    Boolean
  );
  ["RADIATEUR APPOINT", "VENTILATEUR"].forEach((typeEffet) => {
    if (!state.data.listes.typesEffets.includes(typeEffet)) {
      state.data.listes.typesEffets.push(typeEffet);
    }
  });
  state.data.listes.statutsObjetManuels = Array.from(
    new Set(state.data.listes.statutsObjetManuels.map(normalizeText).map((value) => (value === "CASSE" ? "DETRUIT" : value)))
  ).filter(Boolean);
  state.data.listes.causesRemplacement = Array.from(
    new Set(state.data.listes.causesRemplacement.map(normalizeText).map((value) => (value === "CASSE" ? "DETRUIT" : value)))
  ).filter(Boolean);
  if (!state.data.listes.causesRemplacement.length) {
    state.data.listes.causesRemplacement = [...EFFECT_STATUS_CAUSES];
  }
  state.data.listes.coutsRemplacement = state.data.listes.coutsRemplacement
    .map((entry) => ({
      typeEffet: normalizeText(entry.typeEffet),
      designation: normalizeText(entry.designation),
      cause: normalizeText(entry.cause) === "CASSE" ? "DETRUIT" : normalizeText(entry.cause),
      montant: normalizeAmount(entry.montant),
    }))
    .filter((entry) => entry.typeEffet && entry.cause);
  state.data.listes.representantsSignataires = state.data.listes.representantsSignataires
    .map((entry, index) => ({
      id: String(entry.id || `REP${String(index + 1).padStart(4, "0")}`),
      nom: normalizeText(entry.nom),
      fonction: normalizeText(entry.fonction),
    }))
    .filter((entry) => entry.nom || entry.fonction);

  state.data.documentsArchives = state.data.documentsArchives
    .map((entry, index) => {
      const nextEntry = {
        id: String(entry.id || `DOCARCH${String(index + 1).padStart(4, "0")}`),
        personId: String(entry.personId || ""),
        nom: normalizeText(entry.nom),
        prenom: normalizeText(entry.prenom),
        typeDocument: normalizeArchiveTypeLabel(entry.typeDocument),
        dateDocument: String(entry.dateDocument || ""),
        sites: normalizeText(entry.sites),
        typePersonnel: normalizeText(entry.typePersonnel),
        typeContrat: normalizeText(entry.typeContrat),
        statutSignature: normalizeText(entry.statutSignature) || "EN ATTENTE",
        totalEffets: Number(entry.totalEffets || 0),
        totalFacturable: normalizeAmount(entry.totalFacturable),
        pdfPath: String(entry.pdfPath || ""),
        filename: String(entry.filename || ""),
        localPath: String(entry.localPath || ""),
        openLocalUrl: String(entry.openLocalUrl || ""),
        storagePath: String(entry.storagePath || ""),
        publicUrl: String(entry.publicUrl || ""),
        openRemoteUrl: String(entry.openRemoteUrl || ""),
        storageStatus: String(entry.storageStatus || ""),
        pdfQualityStatus: normalizePdfQualityStatus(entry.pdfQualityStatus || ""),
        metadataPath: String(entry.metadataPath || ""),
        dateArchivage: String(entry.dateArchivage || ""),
        fingerprint: isLegacyArrivalArchiveFingerprint(entry.fingerprint) ? "" : String(entry.fingerprint || ""),
      };
      nextEntry.pdfQualityStatus = normalizeArchiveQualityForLegacyEntry(nextEntry);
      return nextEntry;
    })
    .filter((entry) => entry.personId && entry.typeDocument && entry.pdfPath);

  state.data.stocksEffetsManuels = state.data.stocksEffetsManuels
    .map((entry, index) => ({
      id: String(entry.id || `STKM${String(index + 1).padStart(4, "0")}`),
      typeEffet: normalizeText(entry.typeEffet),
      site: normalizeText(entry.typeEffet) === "CARTE TURBOSELF" ? ALL_SITES_VALUE : normalizeText(entry.site),
      referenceEffetId: String(entry.referenceEffetId || ""),
      designation: normalizeText(entry.designation),
      action: normalizeText(entry.action),
      quantite: Math.max(1, Number.parseInt(String(entry.quantite || 1), 10) || 1),
      motif: normalizeText(entry.motif),
      commentaire: normalizeText(entry.commentaire),
      source: normalizeText(entry.source) === "AUTO" ? "AUTO" : "MANUEL",
      date: String(entry.date || getTodayIsoDate()),
    }))
    .filter((entry) => entry.typeEffet && entry.designation && entry.action);

  state.data.demandesSignatureMobile = state.data.demandesSignatureMobile
    .map((entry, index) => ({
      id: String(entry.id || `DSM${String(index + 1).padStart(4, "0")}`),
      token: String(entry.token || ""),
      personId: String(entry.personId || ""),
      docType: normalizeText(entry.docType),
      signer: normalizeText(entry.signer) === "REPRESENTANT" ? "REPRESENTANT" : "PERSONNEL",
      personNom: normalizeText(entry.personNom || ""),
      personPrenom: normalizeText(entry.personPrenom || ""),
      personSite: normalizeText(entry.personSite || ""),
      personTypePersonnel: normalizeText(entry.personTypePersonnel || ""),
      personTypeContrat: normalizeText(entry.personTypeContrat || ""),
      representativeNom: normalizeText(entry.representativeNom || ""),
      representativeFonction: normalizeText(entry.representativeFonction || ""),
      createdAt: String(entry.createdAt || ""),
      expiresAt: String(entry.expiresAt || ""),
      status: normalizeText(entry.status) || "EN ATTENTE",
      validatedAt: String(entry.validatedAt || ""),
    }))
    .filter(
      (entry) =>
        entry.token &&
        entry.personId &&
        ["ARRIVAL", "EXIT"].includes(entry.docType) &&
        ["PERSONNEL", "REPRESENTANT"].includes(entry.signer)
    );

  cleanupExpiredMobileSignatureRequests();

  state.data.personnes = (state.data.personnes || []).filter((person) => !isSoftDeletedEntity(person));

  (state.data.personnes || []).forEach((person) => {
    if (!person.representants || typeof person.representants !== "object") {
      person.representants = {};
    }
    if (!person.representants.arrival || typeof person.representants.arrival !== "object") {
      person.representants.arrival = {};
    }
    if (!person.representants.exit || typeof person.representants.exit !== "object") {
      person.representants.exit = {};
    }
    person.representants.arrival.nom = normalizeText(person.representants.arrival.nom);
    person.representants.arrival.fonction = normalizeText(person.representants.arrival.fonction);
    person.representants.arrival.id = String(person.representants.arrival.id || "");
    person.representants.exit.nom = normalizeText(person.representants.exit.nom);
    person.representants.exit.fonction = normalizeText(person.representants.exit.fonction);
    person.representants.exit.id = String(person.representants.exit.id || "");

    ["arrival", "exit"].forEach((docType) => {
      const rep = person.representants[docType];
      if (!rep.id && (rep.nom || rep.fonction)) {
        const existing = findRepresentativeByValues(rep.nom, rep.fonction);
        if (existing) {
          rep.id = existing.id;
        } else {
          const created = {
            id: getNextId("REP", state.data.listes.representantsSignataires),
            nom: rep.nom || "",
            fonction: rep.fonction || "",
          };
          state.data.listes.representantsSignataires.push(created);
          rep.id = created.id;
        }
      }
      if (rep.id) {
        const linked = state.data.listes.representantsSignataires.find((entry) => entry.id === rep.id);
        if (linked) {
          rep.nom = linked.nom;
          rep.fonction = linked.fonction;
        } else {
          rep.id = "";
        }
      }
    });

    if (!person.signatures || typeof person.signatures !== "object") {
      person.signatures = {};
    }
    if (!person.signatures.arrival || typeof person.signatures.arrival !== "object") {
      person.signatures.arrival = {};
    }
    if (!person.signatures.exit || typeof person.signatures.exit !== "object") {
      person.signatures.exit = {};
    }
    ["arrival", "exit"].forEach((docType) => {
      ["personnel", "representant"].forEach((signer) => {
        const currentEntry = person.signatures[docType][signer];
        if (currentEntry && typeof currentEntry === "object") {
          person.signatures[docType][signer] = {
            image: String(currentEntry.image || ""),
            validatedAt: String(currentEntry.validatedAt || ""),
          };
          return;
        }
        person.signatures[docType][signer] = {
          image: String(currentEntry || ""),
          validatedAt: "",
        };
      });
    });

    delete person.service;
    delete person.historiqueEffets;
    person.typePersonnel = normalizeText(person.typePersonnel);
    person.typeContrat = normalizeText(person.typeContrat);
    person.fonction = normalizeFunctionLabel(person.fonction);
    person.email = String(person.email || "").trim();
    person.phoneMobile = String(person.phoneMobile || "").trim();
    person.sitesAffectation = getPersonSites(person);
    person.site = getPersonSiteLabel(person);

    if (!person.typeContrat && LEGACY_CONTRACT_TYPES.includes(person.typePersonnel)) {
      person.typeContrat = person.typePersonnel;
      person.typePersonnel = "";
    }

    if (!Array.isArray(person.effetsConfies)) {
      person.effetsConfies = [];
    }
    person.effetsConfies = person.effetsConfies.filter((effect) => !isSoftDeletedEntity(effect));

    person.effetsConfies.forEach((effect) => {
      effect.typeEffet = normalizeText(effect.typeEffet);
      effect.siteReference = normalizeText(effect.siteReference);
      effect.referenceEffetId = String(effect.referenceEffetId || "");
      effect.designation = normalizeText(effect.designation);
      effect.numeroIdentification = normalizeText(effect.numeroIdentification);
      effect.vehiculeImmatriculation = normalizeText(effect.vehiculeImmatriculation);
      effect.dateRetour = String(effect.dateRetour || "");
      effect.statutManuel = getStoredManualStatusForEffect(effect.statutManuel, effect.dateRetour);
      const legacyCause = normalizeText(effect.causeRemplacement);
      if (!normalizeEffectCause(effect.cause) && legacyCause) {
        effect.cause = legacyCause;
      }
      effect.dateRemplacement = String(effect.dateRemplacement || "");
      effect.coutRemplacement = normalizeAmount(effect.coutRemplacement);
      effect.commentaire = normalizeText(effect.commentaire);
      effect.etatFacturation = normalizeText(effect.etatFacturation || "");

      if (!typeUsesReferenceCatalog(effect.typeEffet)) {
        effect.referenceEffetId = "";
        effect.designation = "";
      }

        if (!typeUsesSiteField(effect.typeEffet)) {
          effect.siteReference = "";
        } else if (normalizeText(effect.typeEffet) === "CARTE TURBOSELF") {
          effect.siteReference = ALL_SITES_VALUE;
        } else if (!effect.siteReference) {
          effect.siteReference = getDefaultEffectSiteReference(person, effect);
        } else {
          effect.siteReference = normalizeText(effect.siteReference);
        }

      effect.coutRemplacement = getEffectReplacementCost(person, effect);
    });
  });

  let mobileSignatureStatusesReconciled = false;
  const personsByIdForSignatures = new Map(
    (state.data.personnes || []).map((person) => [String(person?.id || ""), person]).filter(([id]) => Boolean(id))
  );
  (state.data.demandesSignatureMobile || []).forEach((request) => {
    const person = personsByIdForSignatures.get(String(request.personId || ""));
    const docType = normalizeText(request.docType || "") === "EXIT" ? "exit" : "arrival";
    const signer = normalizeMobileSignatureSigner(request.signer || "");
    const signature = person?.signatures?.[docType]?.[signer];
    if (!signature || !hasStoredSignaturePayload(signature) || !String(signature.validatedAt || "").trim()) {
      return;
    }
    if (request.status !== "SIGNEE" || String(request.validatedAt || "") !== String(signature.validatedAt || "")) {
      request.status = "SIGNEE";
      request.validatedAt = String(signature.validatedAt || "");
      mobileSignatureStatusesReconciled = true;
    }
  });
  if (mobileSignatureStatusesReconciled && !suppressDirty) {
    markDirty();
  }
  if (!Array.isArray(state.data.listes.referencesEffets)) {
    state.data.listes.referencesEffets = [];
  }

  const legacyRefs = Array.isArray(state.data.referencesEffets)
    ? state.data.referencesEffets
    : Array.isArray(state.data.references)
      ? state.data.references
      : [];
  if (state.data.listes.referencesEffets.length === 0 && legacyRefs.length > 0) {
    const existingIds = new Set(state.data.listes.referencesEffets.map((reference) => String(reference?.id || "")));
    legacyRefs.forEach((reference) => {
      if (!reference) {
        return;
      }
      const normalizedId = String(reference.id || "").trim();
      if (!normalizedId) {
        return;
      }
      if (existingIds.has(normalizedId)) {
        return;
      }
      state.data.listes.referencesEffets.push({
        id: normalizedId,
        site: String(reference.site || ""),
        sitesAffectation: Array.isArray(reference.sitesAffectation)
          ? reference.sitesAffectation
          : reference.site
            ? [reference.site]
            : [],
        typeEffet: normalizeText(reference.typeEffet),
        designation: normalizeText(reference.designation),
        active: reference.active !== false,
      });
      existingIds.add(normalizedId);
    });
  }

  state.data.listes.referencesEffets = state.data.listes.referencesEffets
    .map((reference) => ({
      ...reference,
      site: normalizeText(reference.site),
      sitesAffectation: getReferenceSites(reference),
      typeEffet: normalizeText(reference.typeEffet),
      designation: normalizeText(reference.designation),
      active: reference?.active !== false,
    }))
      .map((reference) => {
        const nextSites =
          (normalizeText(reference.designation) === "CES-PG" ||
            normalizeText(reference.typeEffet) === "CARTE TURBOSELF")
            ? [ALL_SITES_VALUE]
            : getReferenceSites(reference);
        return {
        ...reference,
        sitesAffectation: nextSites,
        site: nextSites.join(" / "),
      };
    });

  const referencesWereReconciled = ensureCatalogReferencesFromAssignedEffects();
  if (referencesWereReconciled && !suppressDirty) {
    markDirty();
  }

  sortListValues(state.data.listes.typesPersonnel);
  sortListValues(state.data.listes.sites);
  sortListValues(state.data.listes.typesEffets);
  sortListValues(state.data.listes.typesContrats);
  sortListValues(state.data.listes.fonctions);
  sortListValues(state.data.listes.causesRemplacement);
  sortListValues(state.data.listes.statutsObjetManuels);
  sortRepresentatives();
  sortReferenceEffects();
  sortDocumentsArchives();
}

function cleanupExpiredMobileSignatureRequests() {
  if (!Array.isArray(state.data?.demandesSignatureMobile)) {
    return;
  }
  const now = Date.now();
  state.data.demandesSignatureMobile = state.data.demandesSignatureMobile.filter((entry) => {
    if (normalizeText(entry?.status) === "SIGNEE") {
      return true;
    }
    const expiresAt = Date.parse(entry?.expiresAt || "");
    return !Number.isFinite(expiresAt) || expiresAt >= now;
  });
}

function isLegacyArrivalArchiveFingerprint(fingerprint) {
  if (!fingerprint) {
    return false;
  }

  try {
    const payload = JSON.parse(fingerprint);
    if (normalizeText(payload?.docType) !== "ARRIVAL") {
      return false;
    }
    if (String(payload?.dateSortieReelle || "")) {
      return true;
    }
    return Array.isArray(payload?.effects) && payload.effects.some((effect) =>
      String(effect?.dateRetour || "") ||
      normalizeText(effect?.statut) ||
      normalizeText(effect?.cause) ||
      String(effect?.dateRemplacement || "")
    );
  } catch (error) {
    return false;
  }
}

function stopMobileSignaturePolling() {
  hideMobileSignatureRecoveryModal();
  if (state.mobileSignaturePollTimerId) {
    window.clearInterval(state.mobileSignaturePollTimerId);
    state.mobileSignaturePollTimerId = 0;
  }
  if (state.mobileSignaturePollRenderSyncTimerId) {
    window.clearTimeout(state.mobileSignaturePollRenderSyncTimerId);
    state.mobileSignaturePollRenderSyncTimerId = 0;
  }
  if (state.mobileSignaturePollResumeTimerId) {
    window.clearTimeout(state.mobileSignaturePollResumeTimerId);
    state.mobileSignaturePollResumeTimerId = 0;
  }
  state.mobileSignaturePollInFlight = false;
  state.mobileSignaturePollBackoffUntil = 0;
  state.mobileSignaturePollErrorCount = 0;
  state.mobileSignaturePollIntervalMs = 0;
  if (document.body.dataset.page === "arrival-document" || document.body.dataset.page === "exit-document") {
    setMobileSignaturePollStatus("Vérification en pause (onglet masqué ou navigation ailleurs)", "warning");
  } else {
    setMobileSignaturePollStatus("");
  }
}

function updateBrowserStorageAlert() {
  if (!navigator?.storage?.estimate) {
    return Promise.resolve(null);
  }

  return (async () => {
    try {
      const estimate = await navigator.storage.estimate();
      const usage = Number(estimate?.usage || 0);
      const quota = Number(estimate?.quota || 0);
      if (!quota || !Number.isFinite(usage) || !Number.isFinite(quota)) {
        return null;
      }

      const ratio = usage / quota;
      const usageMo = Math.round(usage / (1024 * 1024));
      const quotaMo = Math.max(1, Math.round(quota / (1024 * 1024)));
      const ratioPercent = Math.round(ratio * 100);
      const level = ratio >= BROWSER_STORAGE_ALERT_RATIO ? 2 : ratio >= BROWSER_STORAGE_WARNING_RATIO ? 1 : 0;

      const headerActions = document.querySelector(".page-header__actions");
      if (!headerActions) {
        return null;
      }

      let node = document.getElementById("dotations-browser-storage-badge");
      if (!node) {
        node = document.createElement("span");
        node.id = "dotations-browser-storage-badge";
        node.className = "page-header__storage-alert";
        node.setAttribute("role", "status");
        headerActions.appendChild(node);
      }

      node.classList.remove(
        "page-header__storage-alert--warning",
        "page-header__storage-alert--critical"
      );

      if (level === 0) {
        node.hidden = true;
        state.browserStorageQuotaLevel = 0;
        return level;
      }

      node.hidden = false;
      node.textContent = `STOCKAGE NAVIGATEUR ${ratioPercent}% (${usageMo}/${quotaMo} Mo)`;
      if (level === 2) {
        node.classList.add("page-header__storage-alert--critical");
      } else {
        node.classList.add("page-header__storage-alert--warning");
      }
      state.browserStorageQuotaLevel = level;
      return level;
    } catch (error) {
      return null;
    }
  })();
}

function getActiveDocumentMobileSignatureRequest() {
  const page = document.body.dataset.page || "";
  if (page !== "arrival-document" && page !== "exit-document") {
    return null;
  }
  const personId = getCurrentPersonId();
  if (!personId || !state.data) {
    return null;
  }
  const docType = page === "exit-document" ? "exit" : "arrival";
  const { personnel, representant } = getActiveMobileSignatureRequestContext(personId, docType);
  return personnel || representant;
}

function appendPdfTokenToUrl(rawUrl) {
  try {
    if (!isPdfRenderMode()) return String(rawUrl || "");
    const token = String(new URLSearchParams(window.location.search).get("pdfToken") || "").trim();
    if (!token) return String(rawUrl || "");
    const base = new URL(String(rawUrl || ""), window.location.origin);
    if (!base.searchParams.get("pdfToken")) {
      base.searchParams.set("pdfToken", token);
    }
    return `${base.pathname}${base.search}${base.hash}`;
  } catch {
    return String(rawUrl || "");
  }
}

function setMobileSignaturePollStatus(message, tone = "normal") {
  const isPollPage = document.body.dataset.page === "arrival-document" || document.body.dataset.page === "exit-document";
  const existingNode = document.getElementById("dotations-mobile-signature-poll-status");
  if (!isPollPage) {
    if (existingNode) {
      existingNode.hidden = true;
    }
    state.mobileSignaturePollSyncModeSignature = "";
    state.mobileSignaturePollStatusSignature = "";
    return;
  }

  const headerActions = document.querySelector(".page-header__actions");
  if (!headerActions) {
    return;
  }

  let node = existingNode;
  if (!node) {
    node = document.createElement("span");
    node.id = "dotations-mobile-signature-poll-status";
    node.className = "page-header__storage-alert";
    node.setAttribute("role", "status");
    headerActions.appendChild(node);
  }

  if (!message) {
    state.mobileSignaturePollSyncModeSignature = "";
    state.mobileSignaturePollStatusSignature = "";
    node.hidden = true;
    return;
  }

  node.hidden = false;
  node.textContent = message;
  node.classList.remove(
    "page-header__storage-alert--warning",
    "page-header__storage-alert--critical"
  );
  if (tone === "warning") {
    node.classList.add("page-header__storage-alert--warning");
  } else if (tone === "critical") {
    node.classList.add("page-header__storage-alert--critical");
  }
}

function ensureMobileSignatureRecoveryModal() {
  let modal = document.getElementById("mobile-signature-recovery-modal");
  if (modal) {
    return modal;
  }
  modal = document.createElement("div");
  modal.id = "mobile-signature-recovery-modal";
  modal.className = "mobile-signature-recovery-modal";
  modal.hidden = true;
  modal.setAttribute("role", "status");
  modal.setAttribute("aria-live", "polite");
  modal.innerHTML = [
    '<div class="mobile-signature-recovery-modal__card">',
    '<p class="mobile-signature-recovery-modal__eyebrow">SIGNATURE MOBILE</p>',
    '<p class="mobile-signature-recovery-modal__title">RECUPERATION EN COURS</p>',
    '<p class="mobile-signature-recovery-modal__text" id="mobile-signature-recovery-modal-text">L\'UI VERIFIE SI LA SIGNATURE EST DISPONIBLE.</p>',
    '<div class="mobile-signature-recovery-modal__track" aria-hidden="true"><span class="mobile-signature-recovery-modal__bar"></span></div>',
    '</div>',
  ].join("");
  document.body.appendChild(modal);
  return modal;
}

function showMobileSignatureRecoveryModal(message = "") {
  const page = document.body.dataset.page || "";
  if (page !== "arrival-document" && page !== "exit-document") {
    hideMobileSignatureRecoveryModal();
    return;
  }
  const modal = ensureMobileSignatureRecoveryModal();
  const textNode = document.getElementById("mobile-signature-recovery-modal-text");
  if (textNode) {
    textNode.textContent = message || "L'UI VERIFIE SI LA SIGNATURE EST DISPONIBLE.";
  }
  modal.hidden = false;
  state.mobileSignatureRecoveryModalOpen = true;
}

function hideMobileSignatureRecoveryModal() {
  const modal = document.getElementById("mobile-signature-recovery-modal");
  if (modal) {
    modal.hidden = true;
  }
  state.mobileSignatureRecoveryModalOpen = false;
}

function hideMobileSignatureRecoveryModalAfterRender() {
  window.requestAnimationFrame(() => {
    window.setTimeout(() => hideMobileSignatureRecoveryModal(), 180);
  });
}

async function pollMobileSignatureRequest() {
  if (document.visibilityState === "hidden") {
    return;
  }
  if (state.mobileSignaturePollBackoffUntil && Date.now() < state.mobileSignaturePollBackoffUntil) {
    return;
  }
  if (state.mobileSignaturePollInFlight) {
    return;
  }
  state.mobileSignaturePollInFlight = true;
  const page = document.body.dataset.page || "";
  if (page !== "arrival-document" && page !== "exit-document") {
    hideMobileSignatureRecoveryModal();
    stopMobileSignaturePolling();
    return;
  }
  const personId = getCurrentPersonId();
  if (!personId) {
    hideMobileSignatureRecoveryModal();
    stopMobileSignaturePolling();
    return;
  }
  const docType = page === "exit-document" ? "exit" : "arrival";
  try {
    let json = await fetchLatestDataSnapshot({ forceFresh: true });
    if (getDataBackendMode() === "LOCAL_API" && isSupabaseConfigured()) {
      try {
        const pullResponse = await fetch("/api/sync/pull-mobile-signatures", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ personId, docType }),
          cache: "no-store",
        });
        const pullPayload = await pullResponse.json().catch(() => ({}));
        if (pullResponse.ok && pullPayload?.data) {
          json = pullPayload.data;
          if (pullPayload.changed) {
            state.data = pullPayload.data;
            migrateDataModel({ suppressDirty: true });
            state.isDirty = false;
            state.lastPersistedDataSignature = computeDataPersistenceSignature(state.data);
            clearUndoStack();
            showDataStatus("SIGNATURE MOBILE HEBERGEE REPRISE EN LOCAL");
          }
        }
      } catch (hostedError) {
        console.warn("Lecture signature mobile hebergee indisponible", hostedError);
      }
    }
    if (isSupabaseConfigured()) {
      try {
        const signatureRows = await fetchSupabaseMobileSignatureRows(personId, docType);
        if (mergeSupabaseMobileSignatureRows(json, signatureRows, personId, docType)) {
          state.data = json;
          migrateDataModel({ suppressDirty: true });
          state.isDirty = true;
          await saveDataToFile({
            silent: true,
            reloadAfter: false,
            promptDownload: false,
            autoPushHosted: false,
            successText: "SIGNATURE MOBILE HEBERGEE REPRISE",
          });
          state.isDirty = false;
          state.lastPersistedDataSignature = computeDataPersistenceSignature(state.data);
          clearUndoStack();
          showDataStatus("SIGNATURE MOBILE HEBERGEE REPRISE");
        }
      } catch (signatureRowsError) {
        console.warn("Lecture table signatures indisponible", signatureRowsError);
      }
    }
    const normalizedDocType = normalizeText(docType);
    const pollNow = Date.now();
    const requests = Array.isArray(json?.demandesSignatureMobile) ? json.demandesSignatureMobile : [];
    const mobilePollRequestCache =
      state.mobileSignaturePollRequestCache || (state.mobileSignaturePollRequestCache = new Map());
    const relevantRequestsSignature = requests
      .filter((entry) => {
        return (
          String(entry?.personId || "") === String(personId || "") &&
          normalizeText(entry?.docType || "") === normalizedDocType
        );
      })
      .map((entry) =>
        [
          String(entry?.token || ""),
          normalizeMobileSignatureSigner(entry?.signer || ""),
          String(entry?.status || ""),
          String(entry?.validatedAt || ""),
          String(entry?.expiresAt || ""),
        ].join(":")
      )
      .sort()
      .join(";");
    const pollContextCacheKey = `${String(state.supabaseRevision || "")}|${String(state.localMutationTick || 0)}|${String(state.latestDataEtag || "")}|${String(personId || "")}|${normalizedDocType}|${relevantRequestsSignature}`;

    const cachedPollContext = mobilePollRequestCache.get(pollContextCacheKey);
    let trackedPersonnelRequest = cachedPollContext?.trackedPersonnelRequest || null;
    let trackedRepresentativeRequest = cachedPollContext?.trackedRepresentativeRequest || null;
    let activeServerRequests = cachedPollContext?.activeServerRequests || [];
    let matchingServerRequests = cachedPollContext?.matchingServerRequests || [];
    let trackedRequestsLength = cachedPollContext?.trackedRequestsLength || 0;
    let nextRequestsByToken = cachedPollContext?.nextRequestsByToken || new Map();
    let requestStateSignature = cachedPollContext?.requestStateSignature || "";
    let anyPending = cachedPollContext?.anyPending || false;

    if (!cachedPollContext) {
      activeServerRequests = [];
      matchingServerRequests = [];
      for (const entry of requests) {
        const isSamePerson = String(entry?.personId || "") === String(personId || "");
        const isSameDocType = normalizeText(entry?.docType || "") === normalizedDocType;
        const signer = normalizeMobileSignatureSigner(entry?.signer || "");
        if (!isSamePerson || !isSameDocType || !signer || !entry) {
          continue;
        }
        matchingServerRequests.push(entry);
        if (entry?.status === "EN ATTENTE" && Date.parse(entry?.expiresAt || "") > pollNow) {
          activeServerRequests.push(entry);
          if (signer === "personnel" && !trackedPersonnelRequest) {
            trackedPersonnelRequest = entry;
          } else if (signer === "representant" && !trackedRepresentativeRequest) {
            trackedRepresentativeRequest = entry;
          }
        }
      }

      const trackedRequests = [trackedPersonnelRequest, trackedRepresentativeRequest].filter(Boolean);
      const allTrackedRequests = new Map();
      trackedRequests.forEach((request) => {
        const token = String(request?.token || "").trim();
        if (token) {
          allTrackedRequests.set(token, request);
        }
      });
      activeServerRequests.forEach((request) => {
        const token = String(request?.token || "").trim();
        if (token) {
          allTrackedRequests.set(token, request);
        }
      });
      matchingServerRequests.forEach((request) => {
        const token = String(request?.token || "").trim();
        if (token) {
          allTrackedRequests.set(token, request);
        }
      });

      nextRequestsByToken = new Map(
        Array.from(allTrackedRequests.values()).map((request) => [String(request?.token || ""), request])
      );
      anyPending = Array.from(nextRequestsByToken.values()).some((request) => request?.status === "EN ATTENTE" && Date.parse(request?.expiresAt || "") > pollNow);
      requestStateSignature = Array.from(nextRequestsByToken.entries())
        .map(([token, request]) => `${token}:${String(request?.status || "")}`)
        .sort()
        .join(";");
      trackedRequestsLength = trackedRequests.length;

      if (mobilePollRequestCache.size > 40) {
        const oldestKeys = Array.from(mobilePollRequestCache.entries())
          .sort((left, right) => (left[1].createdAt || 0) - (right[1].createdAt || 0))
          .slice(0, Math.max(1, mobilePollRequestCache.size - 40));
        oldestKeys.forEach(([cacheKeyToDelete]) => {
          mobilePollRequestCache.delete(cacheKeyToDelete);
        });
      }

      mobilePollRequestCache.set(pollContextCacheKey, {
        trackedPersonnelRequest,
        trackedRepresentativeRequest,
        activeServerRequests,
        matchingServerRequests,
        nextRequestsByToken,
        requestStateSignature,
        anyPending,
        trackedRequestsLength,
        createdAt: pollNow,
      });
    }

    state.mobileSignaturePollErrorCount = 0;
    state.mobileSignaturePollBackoffUntil = 0;
    const person = Array.isArray(json?.personnes)
      ? json.personnes.find((entry) => String(entry.id || "") === String(personId || "")) || null
      : null;

    if (!person) {
      stopMobileSignaturePolling();
      return;
    }
    state.mobileSignaturePollHasPendingRequest = Boolean(anyPending);
    ensureMobileSignaturePollingInterval();
    if (anyPending && state.mobileSignatureRecoveryModalOpen) {
      showMobileSignatureRecoveryModal("SIGNATURE EN ATTENTE. VERIFICATION DE L'HEBERGE EN COURS.");
    }

    const nextSignaturePersonnel = getSignatureValue(person, docType, "personnel");
    const nextValidatedAtPersonnel = getSignatureValidationDate(person, docType, "personnel");
    const nextSignatureRepresentant = getSignatureValue(person, docType, "representant");
    const nextValidatedAtRepresentant = getSignatureValidationDate(person, docType, "representant");
    const pollStateSignature = [
      personId,
      docType,
      nextSignaturePersonnel,
      nextValidatedAtPersonnel,
      nextSignatureRepresentant,
      nextValidatedAtRepresentant,
      requestStateSignature,
    ].join("|");

    if (state.mobileSignaturePollStateSignature === pollStateSignature) {
      if (!anyPending && trackedRequestsLength === 0 && activeServerRequests.length === 0 && nextRequestsByToken.size === 0) {
        hideMobileSignatureRecoveryModal();
        setMobileSignaturePollStatus("");
        return;
      }
      if (!anyPending) {
        renderMobileSignatureLink(docType, "personnel", "");
        renderMobileSignatureLink(docType, "representant", "");
        hideMobileSignatureRecoveryModalAfterRender();
      }
      return;
    }
    state.mobileSignaturePollStateSignature = pollStateSignature;
    state.data = json;
    migrateDataModel({ suppressDirty: true });

    if (!anyPending && trackedRequestsLength === 0 && activeServerRequests.length === 0 && nextRequestsByToken.size === 0) {
      hideMobileSignatureRecoveryModal();
      setMobileSignaturePollStatus("");
      return;
    }

    if (!anyPending) {
      renderMobileSignatureLink(docType, "personnel", "");
      renderMobileSignatureLink(docType, "representant", "");
      setMobileSignaturePollStatus("");
    }

    schedulePageRender();
    queueAutoGenerateSignedDocumentsPdfIfMissing();
    const signedRepresentative = Array.from(nextRequestsByToken.values()).some(
      (request) => normalizeMobileSignatureSigner(request.signer || "") === "representant" && request.status === "SIGNEE"
    );
    const signedPersonnel = Array.from(nextRequestsByToken.values()).some(
      (request) => normalizeMobileSignatureSigner(request.signer || "") === "personnel" && request.status === "SIGNEE"
    );
    if (signedRepresentative || signedPersonnel) {
      hideMobileSignatureRecoveryModalAfterRender();
    }
    if (signedRepresentative) showDataStatus("SIGNATURE MOBILE DU REPRESENTANT ENREGISTREE");
    else if (signedPersonnel) showDataStatus("SIGNATURE MOBILE DU PERSONNEL ENREGISTREE");
  } catch (error) {
    const previousErrorCount = Math.min((state.mobileSignaturePollErrorCount || 0) + 1, 4);
    state.mobileSignaturePollErrorCount = previousErrorCount;
    state.mobileSignaturePollBackoffUntil = Date.now() + MOBILE_SIGNATURE_POLL_ERROR_BACKOFF_MS * previousErrorCount;
  } finally {
    state.mobileSignaturePollInFlight = false;
  }
}

function ensureMobileSignaturePollingInterval() {
  if (!state.mobileSignaturePollTimerId) {
    state.mobileSignaturePollIntervalMs = 0;
    return;
  }
  const desiredInterval = state.mobileSignaturePollHasPendingRequest
    ? MOBILE_SIGNATURE_POLL_INTERVAL_MS
    : MOBILE_SIGNATURE_POLL_IDLE_INTERVAL_MS;
  if (state.mobileSignaturePollIntervalMs === desiredInterval) {
    return;
  }
  window.clearInterval(state.mobileSignaturePollTimerId);
  state.mobileSignaturePollTimerId = window.setInterval(() => {
    pollMobileSignatureRequest();
  }, desiredInterval);
  state.mobileSignaturePollIntervalMs = desiredInterval;
}

function syncMobileSignaturePolling() {
  const page = document.body.dataset.page || "";
  if (page !== "arrival-document" && page !== "exit-document") {
    return;
  }
  const personId = getCurrentPersonId();
  if (!personId) {
    return;
  }
  const docType = page === "exit-document" ? "exit" : "arrival";
  const hasActiveRequest = hasActiveMobileSignatureRequest(personId, docType);
  if (!hasActiveRequest) {
    state.mobileSignaturePollHasPendingRequest = false;
    setMobileSignaturePollStatus("Verification de fond des signatures mobiles", "normal");
  }
  const now = Date.now();
  if (state.mobileSignaturePollLastSyncAt && now - state.mobileSignaturePollLastSyncAt < MOBILE_SIGNATURE_POLL_SYNC_MIN_GAP_MS) {
    return;
  }
  state.mobileSignaturePollLastSyncAt = now;
  const desiredInterval = state.mobileSignaturePollHasPendingRequest
    ? MOBILE_SIGNATURE_POLL_INTERVAL_MS
    : MOBILE_SIGNATURE_POLL_IDLE_INTERVAL_MS;
  const intervalSeconds = Math.max(1, Math.round(desiredInterval / 1000));
  const isWaiting = !state.mobileSignaturePollHasPendingRequest;
  const modeText = isWaiting
    ? "Aucune signature en attente: vérification de fond"
    : "Signature en attente: vérification active";
  const syncModeSignature = `${docType}|${state.mobileSignaturePollHasPendingRequest ? "1" : "0"}|${intervalSeconds}`;
  if (state.mobileSignaturePollSyncModeSignature !== syncModeSignature) {
    state.mobileSignaturePollSyncModeSignature = syncModeSignature;
    setMobileSignaturePollStatus(
      `${modeText} · prochaine vérification toutes les ${intervalSeconds}s`,
      state.mobileSignaturePollHasPendingRequest ? "normal" : "warning"
    );
  }
  if (state.mobileSignaturePollTimerId) {
    if (state.mobileSignaturePollIntervalMs !== desiredInterval) {
      window.clearInterval(state.mobileSignaturePollTimerId);
      state.mobileSignaturePollTimerId = 0;
      state.mobileSignaturePollIntervalMs = 0;
    } else {
      pollMobileSignatureRequest();
      return;
    }
  }
  if (!state.mobileSignaturePollIntervalMs) {
    pollMobileSignatureRequest();
  }
  if (!state.mobileSignaturePollTimerId) {
    stopMobileSignaturePolling();
    pollMobileSignatureRequest();
    state.mobileSignaturePollIntervalMs = desiredInterval;
    state.mobileSignaturePollTimerId = window.setInterval(() => {
      pollMobileSignatureRequest();
    }, desiredInterval);
  }
}

function scheduleMobileSignatureRenderSync() {
  const page = document.body.dataset.page || "";
  if (page !== "arrival-document" && page !== "exit-document") {
    return;
  }
  const now = Date.now();
  const lastRequested = state.mobileSignaturePollRenderSyncAt || 0;
  if (!lastRequested || now - lastRequested >= MOBILE_SIGNATURE_POLL_RENDER_SYNC_DEBOUNCE_MS) {
    state.mobileSignaturePollRenderSyncAt = now;
    syncMobileSignaturePolling();
    return;
  }
  const delay = MOBILE_SIGNATURE_POLL_RENDER_SYNC_DEBOUNCE_MS - (now - lastRequested);
  setMobileSignaturePollStatus(
    `Nouvelle vérification prévue dans ${Math.ceil(delay / 1000)} seconde(s)`,
    "critical"
  );
  if (state.mobileSignaturePollRenderSyncTimerId) {
    window.clearTimeout(state.mobileSignaturePollRenderSyncTimerId);
  }
  state.mobileSignaturePollRenderSyncTimerId = window.setTimeout(() => {
    state.mobileSignaturePollRenderSyncTimerId = 0;
    state.mobileSignaturePollRenderSyncAt = Date.now();
    syncMobileSignaturePolling();
  }, delay);
}

function bindMobileSignatureVisibilityPolling() {
  if (state.mobileSignatureVisibilityBound) {
    return;
  }
  document.addEventListener("visibilitychange", () => {
    const page = document.body.dataset.page || "";
    if (page !== "arrival-document" && page !== "exit-document") {
      return;
    }
    if (document.visibilityState === "hidden") {
      stopMobileSignaturePolling();
      return;
    }
    if (state.mobileSignaturePollResumeTimerId) {
      window.clearTimeout(state.mobileSignaturePollResumeTimerId);
      state.mobileSignaturePollResumeTimerId = 0;
    }
    state.mobileSignaturePollResumeTimerId = window.setTimeout(() => {
      state.mobileSignaturePollResumeTimerId = 0;
      syncMobileSignaturePolling();
    }, MOBILE_SIGNATURE_POLL_RESUME_DELAY_MS);
  });
  state.mobileSignatureVisibilityBound = true;
}

function applyActiveNav() {
  const page = document.body.dataset.page;
  document.querySelectorAll("[data-nav]").forEach((link) => {
    link.classList.toggle("is-active", link.dataset.nav === page);
  });
}

function bindLoadButton() {
  document.querySelectorAll(".js-load-data").forEach((button) => {
    button.onclick = () => reloadData("RECHARGEMENT DES DONNEES...");
  });
}

function bindSaveButtons() {
  document.querySelectorAll(".js-save-data").forEach((button) => {
    button.onclick = () => {
      state.saveButtonLatchedDirty = false;
      renderDirtyState();
      return saveDataToFile();
    };
  });
}

function bindDirtyFallbackTracking() {
  if (state.dirtyFallbackBound) {
    return;
  }

  const trackedSelectors = [
    "#add-person-form",
    "#person-sheet-form",
    "#effect-form",
    "#mobile-signature-settings-form",
    "#representative-signatory-form",
    "#reference-effect-form",
    "#replacement-cost-form",
    ".js-reference-list-form",
  ];

  const shouldTrack = (form) => trackedSelectors.some((selector) => form.matches(selector));

  const handlePotentialChange = (event) => {
    const target = event?.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const form = target.closest("form");
    if (!(form instanceof HTMLFormElement) || !shouldTrack(form)) {
      return;
    }
    if (target instanceof HTMLButtonElement) {
      return;
    }
    if (!state.isDirty) {
      markDirty();
    }
  };

  document.addEventListener("input", handlePotentialChange, true);
  document.addEventListener("change", handlePotentialChange, true);
  state.dirtyFallbackBound = true;
}

function bindPdfButtons() {
  document.querySelectorAll(".js-open-pdf").forEach((button) => {
    button.onclick = () => {
      const docType = String(button.getAttribute("data-doc-type") || "");
      const person = getCurrentPerson();
      if (!person) {
        showDataStatus("AUCUNE PERSONNE SELECTIONNEE");
        return;
      }
      openPdfDocument(docType, getCurrentPersonId());
    };
  });
}

function updateDocumentPdfButtonsState() {
  const person = getCurrentPerson();
  document.querySelectorAll(".js-open-pdf").forEach((button) => {
    const docType = String(button.getAttribute("data-doc-type") || "");
    const canOpen = Boolean(person);
    const fullySigned = Boolean(person && isDocumentFullySigned(person, docType));
    button.classList.toggle("is-disabled", !canOpen);
    button.setAttribute("aria-disabled", canOpen ? "false" : "true");
    button.setAttribute(
      "title",
      canOpen
        ? (fullySigned ? "GENERER LE PDF SIGNE ET ARCHIVABLE" : "GENERER UN PDF DE TRAVAIL NON SIGNE")
        : "INDISPONIBLE : AUCUNE PERSONNE SELECTIONNEE"
    );
  });
}
function bindMobileSignatureButtons() {
  document.querySelectorAll(".js-open-mobile-signature").forEach((button) => {
    button.onclick = async () => {
      const docType = String(button.getAttribute("data-doc-type") || "");
      const signer = normalizeMobileSignatureSigner(button.getAttribute("data-signer") || "");
      const personId = getCurrentPersonId();
      if (!docType || !personId) {
        showDataStatus("AUCUNE PERSONNE SELECTIONNEE");
        return;
      }
      try {
        await openMobileSignatureRequest(docType, personId, signer);
      } catch (error) {
        const message = String(error?.message || error || "ERREUR INCONNUE").slice(0, 180);
        console.error("Ouverture signature mobile impossible", error);
        showDataStatus(`SIGNATURE MOBILE IMPOSSIBLE : ${message}`);
        window.alert(`SIGNATURE MOBILE IMPOSSIBLE : ${message}`);
      }
    };
  });
}

function updateUrgencyModeUi() {
  document.querySelectorAll(".js-toggle-urgency").forEach((button) => {
    if (!(button instanceof HTMLButtonElement)) {
      return;
    }
    button.classList.toggle("is-active", state.urgentMode);
    button.textContent = state.urgentMode ? "URGENCES ACTIVES" : "MODE URGENCES";
  });
}

function bindOverviewUrgencyActions() {
  document.querySelectorAll(".js-toggle-urgency").forEach((button) => {
    if (!(button instanceof HTMLButtonElement)) {
      return;
    }
    button.onclick = () => {
      state.urgentMode = !state.urgentMode;
      saveNavigationContext({ filters: state.filters, urgentMode: state.urgentMode });
      updateUrgencyModeUi();
      schedulePageRender();
      showActionStatus(
        state.urgentMode ? "warning" : "update",
        state.urgentMode ? "MODE URGENCES ACTIVE" : "MODE URGENCES DESACTIVE"
      );
    };
  });
  updateUrgencyModeUi();
}

function buildControlReportHtml(persons) {
  const alerts = persons.filter((person) => hasOverdueExit(person));
  const critical = persons.filter((person) => hasUrgencyCondition(person));
  let nonRendus = 0;
  let totalFacturable = 0;

  persons.forEach((person) => {
    (person.effetsConfies || []).forEach((effect) => {
      if (normalizeText(getEffectStatus(person, effect)) === "NON RENDU") {
        nonRendus += 1;
      }
      if (isEffectChargeable(person, effect)) {
        totalFacturable += getEffectReplacementCost(person, effect);
      }
    });
  });

  const criticalRows = critical.length
    ? critical
        .map(
          (person) => `<tr>
      <td>${escapeHtml(person.nom || "")}</td>
      <td>${escapeHtml(person.prenom || "")}</td>
      <td>${escapeHtml(getPersonSiteLabel(person) || "-")}</td>
      <td class="alert-cell">${escapeHtml(getOverdueExitMessage(person) || "-")}</td>
      <td>${(person.effetsConfies || []).filter((effect) => normalizeText(getEffectStatus(person, effect)) === "NON RENDU").length}</td>
    </tr>`
        )
        .join("")
    : `<tr><td colspan="5">AUCUN DOSSIER CRITIQUE</td></tr>`;

  return `<!doctype html>
<html lang="fr">
<head>
  <meta charset="utf-8" />
  <title>ETAT DE CONTROLE</title>
  <style>
    :root{
      --bg:#f2f6f9;
      --card:#ffffff;
      --line:#d3dee6;
      --line-strong:#b8cad6;
      --title:#193243;
      --text:#294757;
      --muted:#5a7585;
      --accent:#3f6170;
      --accent-soft:#e9f1f5;
      --warn:#8d4d2f;
      --warn-bg:#f8eee8;
    }
    *{box-sizing:border-box}
    body{
      margin:0;
      padding:20px;
      font-family:"Segoe UI",Arial,sans-serif;
      background:var(--bg);
      color:var(--text);
    }
    .wrap{
      max-width:1200px;
      margin:0 auto;
      background:var(--card);
      border:1px solid var(--line);
      border-radius:14px;
      padding:18px;
      box-shadow:0 8px 24px rgba(33,60,75,.08);
    }
    .head{
      display:flex;
      justify-content:space-between;
      align-items:flex-end;
      gap:16px;
      padding-bottom:10px;
      border-bottom:1px solid var(--line);
      margin-bottom:14px;
    }
    h1{
      margin:0;
      font-size:24px;
      line-height:1.1;
      color:var(--title);
      letter-spacing:.01em;
    }
    .meta{
      font-size:12px;
      color:var(--muted);
      text-transform:uppercase;
      letter-spacing:.06em;
      white-space:nowrap;
    }
    .kpis{
      display:grid;
      grid-template-columns:repeat(4,minmax(150px,1fr));
      gap:10px;
      margin-bottom:14px;
    }
    .kpi{
      border:1px solid var(--line);
      border-radius:10px;
      padding:9px 11px;
      background:var(--accent-soft);
      font-size:11px;
      color:var(--muted);
      letter-spacing:.05em;
      text-transform:uppercase;
    }
    .kpi b{
      display:block;
      margin-top:4px;
      font-size:23px;
      color:var(--title);
      letter-spacing:0;
      text-transform:none;
    }
    .kpi--warn{
      background:var(--warn-bg);
      border-color:#ecc2ae;
      color:#8a5539;
    }
    .kpi--warn b{color:var(--warn)}
    .table-wrap{
      border:1px solid var(--line);
      border-radius:10px;
      overflow:hidden;
    }
    table{
      width:100%;
      border-collapse:collapse;
      font-size:13px;
    }
    th,td{
      border-bottom:1px solid var(--line);
      padding:8px 9px;
      text-align:left;
      vertical-align:top;
    }
    th{
      background:#edf4f8;
      color:#35596b;
      font-size:11px;
      letter-spacing:.06em;
      text-transform:uppercase;
      border-bottom:1px solid var(--line-strong);
    }
    tbody tr:nth-child(even) td{background:#fbfdff}
    tbody tr:last-child td{border-bottom:none}
    .alert-cell{
      color:#8a4e30;
      font-weight:600;
    }
  </style>
</head>
<body>
  <div class="wrap">
  <div class="head">
    <h1>ETAT DE CONTROLE</h1>
    <div class="meta">EDITE LE ${escapeHtml(formatCurrentUiTimestamp())}</div>
  </div>
  <div class="kpis">
    <div class="kpi">DOSSIERS FILTRES<b>${persons.length}</b></div>
    <div class="kpi kpi--warn">ALERTES SORTIE<b>${alerts.length}</b></div>
    <div class="kpi">EFFETS NON RENDUS<b>${nonRendus}</b></div>
    <div class="kpi kpi--warn">TOTAL FACTURABLE<b>${escapeHtml(formatAmountWithEuro(totalFacturable))}</b></div>
  </div>
  <div class="table-wrap">
  <table>
    <thead>
      <tr><th>NOM</th><th>PRENOM</th><th>SITE(S)</th><th>ALERTE</th><th>NON RENDUS</th></tr>
    </thead>
    <tbody>${criticalRows}</tbody>
  </table>
  </div>
  </div>
</body>
</html>`;
}

function bindOverviewControlExport() {
  document.querySelectorAll(".js-export-control").forEach((button) => {
    if (!(button instanceof HTMLButtonElement)) {
      return;
    }
    button.onclick = () => {
      try {
        const persons = getFilteredPersons();
        const popup = window.open("about:blank", "_blank");
        if (!popup) {
          showDataStatus("AUTORISER L'OUVERTURE DE FENETRE POUR L'EXPORT");
          return;
        }
        const html = buildControlReportHtml(persons);
        popup.document.open();
        popup.document.write(html);
        popup.document.close();
        showActionStatus("update", "ETAT DE CONTROLE OUVERT");
      } catch (error) {
        console.error("EXPORT ETAT DE CONTROLE IMPOSSIBLE", error);
        showDataStatus("EXPORT ETAT DE CONTROLE IMPOSSIBLE");
      }
    };
  });
}

function getTableColumnCount(target, fallback = 1) {
  const body = typeof target === "string" ? document.getElementById(target) : target;
  const table = body?.closest("table");
  const headerRow = table?.querySelector("thead tr:last-child");
  const count = headerRow?.children?.length || table?.querySelectorAll("thead th")?.length || fallback;
  return count || fallback;
}

function buildEmptyTableRow(target, text, fallback = 1) {
  return `<tr><td colspan="${getTableColumnCount(target, fallback)}" class="table-empty">${text}</td></tr>`;
}

function getEffectTableSort(tableName) {
  const current = state.tableSorts?.[tableName];
  return current && current.key && current.dir
    ? current
    : (DEFAULT_TABLE_SORTS[tableName] || { key: "nom", dir: "asc" });
}

function isTableSortModified(tableName) {
  const defaults = DEFAULT_TABLE_SORTS[tableName];
  if (!defaults) {
    return false;
  }
  const current = getEffectTableSort(tableName);
  return current.key !== defaults.key || current.dir !== defaults.dir;
}

function getSortTablesForCurrentPage() {
  switch (String(document.body?.dataset?.page || "")) {
    case "overview":
      return ["overviewPersons"];
    case "person-sheet":
      return ["sheetEffects"];
    case "arrival-document":
      return ["arrivalEffects"];
    case "exit-document":
      return ["exitEffects"];
    case "documents-archives":
      return ["documentsArchives"];
    case "reference-bases":
      return ["referenceEffects"];
    default:
      return [];
  }
}

function resetTableSortsForCurrentPage() {
  getSortTablesForCurrentPage().forEach((tableName) => {
    const defaults = DEFAULT_TABLE_SORTS[tableName];
    if (defaults) {
      state.tableSorts[tableName] = { ...defaults };
    }
  });
}

function hasActiveTableSortForCurrentPage() {
  return getSortTablesForCurrentPage().some((tableName) => isTableSortModified(tableName));
}

function setEffectTableSort(tableName, key) {
  if (!state.tableSorts) {
    state.tableSorts = {};
  }
  const current = getEffectTableSort(tableName);
  state.tableSorts[tableName] = {
    key,
    dir: current.key === key && current.dir === "asc" ? "desc" : "asc",
  };
}

function getEffectSortValue(person, effect, key) {
  switch (key) {
    case "typeEffet":
      return effect?.typeEffet || "";
    case "designation":
      return getEffectDisplayDesignation(effect) || "";
    case "siteReference":
      return getEffectDisplaySite(effect) || "";
    case "numeroIdentification":
      return effect?.numeroIdentification || "";
    case "dateRemise":
      return effect?.dateRemise || "";
    case "dateRetour":
      return effect?.dateRetour || "";
    case "statut":
      return getEffectStatus(person, effect) || "";
    case "cause":
      return getEffectReplacementCause(person, effect) || "";
    case "dateRemplacement":
      return effect?.dateRemplacement || "";
    case "cout":
      return key === "cout" ? getEffectUnitValue(effect) : 0;
    case "coutFacturable":
      return getEffectReplacementCost(person, effect);
    case "commentaire":
      return effect?.commentaire || "";
    default:
      return "";
  }
}

function compareEffectValues(left, right, isNumeric = false) {
  if (isNumeric) {
    return normalizeAmount(left) - normalizeAmount(right);
  }
  return compareTextValues(left, right);
}

function getOverviewSortValue(person, key) {
  const currentEffects = getCurrentAssignedEffectsForActiveFilters(person);
  const movementMap = getArrivalComplementMovementMap(person, getEffectsForActiveFilters(person));
  const movementCount = movementMap.size;
  switch (key) {
    case "nom":
      return person?.nom || "";
    case "prenom":
      return person?.prenom || "";
    case "site":
      return getPersonSiteLabel(person) || "";
    case "typePersonnel":
      return person?.typePersonnel || "";
    case "typeContrat":
      return person?.typeContrat || "";
    case "dateEntree":
      return person?.dateEntree || "";
    case "dateSortiePrevue":
      return person?.dateSortiePrevue || "";
    case "dateSortieReelle":
      return person?.dateSortieReelle || "";
    case "statutDossier":
      return getDossierStatus(person) || "";
    case "nbEffets":
      return currentEffects.length;
    case "nonRendus":
      return currentEffects.filter((effect) => getEffectStatus(person, effect) === "NON RENDU").length;
    case "mouvements":
      return movementCount;
    default:
      return "";
  }
}

function sortPersonsForOverview(persons) {
  const sort = getEffectTableSort("overviewPersons");
  const numericKeys = new Set(["nbEffets", "nonRendus", "mouvements"]);
  return [...persons].sort((left, right) => {
    const primary = compareEffectValues(
      getOverviewSortValue(left, sort.key),
      getOverviewSortValue(right, sort.key),
      numericKeys.has(sort.key)
    );
    if (primary !== 0) {
      return sort.dir === "asc" ? primary : -primary;
    }

    const nomCompare = compareTextValues(left?.nom || "", right?.nom || "");
    if (nomCompare !== 0) {
      return nomCompare;
    }

    const prenomCompare = compareTextValues(left?.prenom || "", right?.prenom || "");
    if (prenomCompare !== 0) {
      return prenomCompare;
    }

    return compareTextValues(left?.id || "", right?.id || "");
  });
}

function sortEffectsForTable(person, effects, tableName) {
  const sort = getEffectTableSort(tableName);
  const numericKeys = new Set(["cout", "coutFacturable"]);
  const sorted = [...effects].sort((left, right) => {
    const primary = compareEffectValues(
      getEffectSortValue(person, left, sort.key),
      getEffectSortValue(person, right, sort.key),
      numericKeys.has(sort.key)
    );
    if (primary !== 0) {
      return sort.dir === "asc" ? primary : -primary;
    }

    const typeCompare = compareTextValues(left?.typeEffet || "", right?.typeEffet || "");
    if (typeCompare !== 0) {
      return typeCompare;
    }

    const designationCompare = compareTextValues(
      getEffectDisplayDesignation(left) || "",
      getEffectDisplayDesignation(right) || ""
    );
    if (designationCompare !== 0) {
      return designationCompare;
    }

    return compareTextValues(left?.id || "", right?.id || "");
  });
  return sorted;
}

function getReferenceSortValue(reference, key, renderContext = null) {
  switch (key) {
    case "site":
      return getReferenceSiteLabel(reference) || "";
    case "typeEffet":
      return reference?.typeEffet || "";
    case "designation":
      return reference?.designation || "";
    case "usage":
      return renderContext?.referenceEffectUsage?.get(String(reference?.id || "")) || 0;
    default:
      return "";
  }
}

function sortReferencesForTable(references, tableName, renderContext = null) {
  const sort = getEffectTableSort(tableName);
  const numericKeys = new Set(["usage"]);
  return [...references].sort((left, right) => {
    const primary = compareEffectValues(
      getReferenceSortValue(left, sort.key, renderContext),
      getReferenceSortValue(right, sort.key, renderContext),
      numericKeys.has(sort.key)
    );
    if (primary !== 0) {
      return sort.dir === "asc" ? primary : -primary;
    }

    const siteCompare = compareTextValues(
      getReferenceSiteLabel(left) || "",
      getReferenceSiteLabel(right) || ""
    );
    if (siteCompare !== 0) {
      return sort.dir === "asc" ? siteCompare : -siteCompare;
    }

    const typeCompare = compareTextValues(
      String(left?.typeEffet || ""),
      String(right?.typeEffet || "")
    );
    if (typeCompare !== 0) {
      return sort.dir === "asc" ? typeCompare : -typeCompare;
    }

    const designationCompare = compareTextValues(
      String(left?.designation || ""),
      String(right?.designation || "")
    );
    if (designationCompare !== 0) {
      return sort.dir === "asc" ? designationCompare : -designationCompare;
    }

    return compareTextValues(String(left?.id || ""), String(right?.id || ""));
  });
}

function updateSortableHeaders(tableName) {
  document.querySelectorAll(`[data-sort-table="${tableName}"]`).forEach((header) => {
    const key = String(header.getAttribute("data-sort-key") || "");
    const sort = getEffectTableSort(tableName);
    const active = sort.key === key;
    header.classList.toggle("is-sorted", active);
    header.dataset.sortDirection = active ? sort.dir : "";
    header.setAttribute(
      "aria-sort",
      active ? (sort.dir === "asc" ? "ascending" : "descending") : "none"
    );
  });
}

function bindEffectTableSorting(tableName = "") {
  const selector = tableName
    ? `[data-sort-table="${String(tableName)}"][data-sort-key]`
    : "[data-sort-table][data-sort-key]";

  const headers = document.querySelectorAll(selector);
  if (!headers.length) {
    return;
  }

  headers.forEach((header) => {
    if (header.dataset.sortBound === "1") {
      return;
    }
    header.dataset.sortBound = "1";
    header.tabIndex = 0;
    header.role = "button";
    const activate = () => {
      const tableName = String(header.getAttribute("data-sort-table") || "");
      const key = String(header.getAttribute("data-sort-key") || "");
      if (!tableName || !key) {
        return;
      }
      setEffectTableSort(tableName, key);
      updateAllFilterResetHighlights();
      schedulePageRender();
    };
    header.onclick = activate;
    header.onkeydown = (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        activate();
      }
    };
  });
}

function ensurePdfProgressModal() {
  return null;
}

function setPdfProgress(percent, text) {
  return;
}

function showPdfProgressModal(docType) {
  state.pdfGenerationActive = true;
}

function stopPdfProgressTimer() {
  if (state.pdfProgressTimerId) {
    window.clearInterval(state.pdfProgressTimerId);
    state.pdfProgressTimerId = 0;
  }
}

function hidePdfProgressModal() {
  stopPdfProgressTimer();
  state.pdfGenerationActive = false;
}

function bindPdfModalCleanup() {
  if (pdfModalCleanupBound) {
    return;
  }

  pdfModalCleanupBound = true;
}

function waitForNextPaint() {
  return new Promise((resolve) => {
    window.requestAnimationFrame(() => {
      window.setTimeout(resolve, 40);
    });
  });
}

function wait(milliseconds) {
  return new Promise((resolve) => {
    window.setTimeout(resolve, milliseconds);
  });
}

function getDocumentPagePath(docType) {
  return normalizeText(docType) === "EXIT" ? "document-sortie.html" : "document-arrivee.html";
}

function getHostedPdfDocumentPath(docType, personId, mode = "STANDARD") {
  const pagePath = getDocumentPagePath(docType);
  return `${pagePath}?personId=${encodeURIComponent(personId)}&pdf=1&mode=${encodeURIComponent(normalizeText(mode || "STANDARD"))}`;
}

function getCurrentHostedAppBasePath() {
  const pathname = String(window.location.pathname || "/");
  const segments = pathname.split("/").filter(Boolean);
  if (!segments.length) {
    return "/";
  }
  const firstSegment = segments[0];
  const host = String(window.location.hostname || "").toLowerCase();
  if (host.endsWith(".github.io")) {
    return `/${firstSegment}/`;
  }
  return pathname.includes("/Dotations/") || firstSegment === "Dotations" ? "/Dotations/" : "/";
}

function resolveHostedRelativeUrl(raw) {
  const value = String(raw || "").trim();
  if (!value) {
    return null;
  }
  const base = new URL(window.location.href || "./", window.location.origin);
  if (/^\/(document-(arrivee|sortie)\.html(?:\?|$))/i.test(value)) {
    return new URL(value.replace(/^\/+/, getCurrentHostedAppBasePath()), window.location.origin);
  }
  const parsed = new URL(value, base);
  const host = String(parsed.hostname || "").toLowerCase();
  const isHostedDocument = /\/document-(arrivee|sortie)\.html$/i.test(parsed.pathname);
  if (isHostedDocument && host.endsWith(".github.io") && !parsed.pathname.startsWith(getCurrentHostedAppBasePath())) {
    parsed.pathname = `${getCurrentHostedAppBasePath()}${parsed.pathname.replace(/^\/+/, "")}`;
  }
  return parsed;
}

function normalizeDirectPdfOpenUrl(value) {
  const raw = String(value || "").trim();
  if (!raw) {
    return "";
  }
  if (/^data\/pdf\//i.test(raw)) {
    const apiPath = raw.replace(/^data\/pdf\//i, "");
    return new URL(`/api/pdf-file?path=${encodeURIComponent(apiPath)}`, window.location.origin).href;
  }
  let parsed = null;
  try {
    parsed = resolveHostedRelativeUrl(raw);
  } catch {
    return "";
  }
  const pathname = parsed.pathname.toLowerCase();
  const isPdfApi = pathname.endsWith("/api/pdf-file");
  const isPdfFile = pathname.endsWith(".pdf");
  const isHostedPdfRender =
    (pathname.endsWith("/document-arrivee.html") || pathname.endsWith("/document-sortie.html")) &&
    parsed.searchParams.get("pdf") === "1";
  return isPdfApi || isPdfFile || isHostedPdfRender ? parsed.href : "";
}

function openPdfUrlInBrowserWindow(popup, value) {
  const pdfUrl = normalizeDirectPdfOpenUrl(value);
  if (!pdfUrl) {
    return false;
  }
  try {
    popup.location.replace(pdfUrl);
    popup.focus();
    return true;
  } catch (popupNavigationError) {
    console.error(popupNavigationError);
    const fallbackPopup = window.open(pdfUrl, "_blank", "noopener");
    return Boolean(fallbackPopup);
  }
}

async function openPdfDocument(docType, personId) {
  if (state.isDirty) {
    showDataStatus("SAUVEGARDER AVANT OUVERTURE DU PDF");
    return;
  }

  if (!personId) {
    showDataStatus("AUCUNE PERSONNE SELECTIONNEE");
    return;
  }

  const person = state.data?.personnes?.find((entry) => entry.id === personId) || null;
  const pdfOpenKey = `${normalizeText(docType)}|${String(personId || "")}`;
  const nowMs = Date.now();
  const isRapidDuplicateOpen =
    state.lastPdfOpenKey === pdfOpenKey && nowMs - Number(state.lastPdfOpenAtMs || 0) < 2000;
  if (isRapidDuplicateOpen) {
    showDataStatus("OUVERTURE PDF DEJA EN COURS");
    return;
  }
  state.lastPdfOpenKey = pdfOpenKey;
  state.lastPdfOpenAtMs = nowMs;
  const shouldArchive = isDocumentFullySigned(person, docType);
  const archiveMode = getDocumentArchiveMode(person, docType);
  const reusableArchive = shouldArchive ? findReusableArchivedDocument(person, docType) : null;
  const reusableArchiveStorageRef = reusableArchive ? parseStorageSchemePath(reusableArchive.pdfPath) : null;
  const canPromoteReusableArchiveToStorage =
    Boolean(reusableArchive) &&
    !reusableArchiveStorageRef &&
    shouldArchive &&
    getDataBackendMode() === "LOCAL_API" &&
    isSupabaseConfigured();

  const popup = window.open("", "_blank");
  if (!popup) {
    showDataStatus("AUTORISER L'OUVERTURE DU PDF DANS LE NAVIGATEUR");
    return;
  }

  try {
    if (reusableArchive && !canPromoteReusableArchiveToStorage) {
      const reusableOpenUrl = getArchivePreferredOpenPath(reusableArchive) || getDocumentArchiveOpenPath(reusableArchive);
      if (openPdfUrlInBrowserWindow(popup, reusableOpenUrl)) {
        showActionStatus("update", "PDF ARCHIVE REUTILISE");
        return;
      }
      showDataStatus("PDF ARCHIVE EXISTANT NON DIRECT - REGENERATION POUR VERIFICATION");
    }
    if (canPromoteReusableArchiveToStorage) {
      showDataStatus("PDF ARCHIVE LOCAL DETECTE - MIGRATION VERS SUPABASE EN COURS");
    }

    if (getDataBackendMode() !== "LOCAL_API") {
      const hostedPath = getHostedPdfDocumentPath(docType, personId, archiveMode);
      popup.location.href = hostedPath;
      if (person && shouldArchive) {
        await registerArchivedDocument(person, docType, hostedPath, "", archiveMode);
        clearPendingPdfTaskFor(person.id, docType);
      }
      showDataStatus("DOCUMENT OUVERT - UTILISER IMPRIMER POUR GENERER LE PDF");
      return;
    }

    showPdfProgressModal(docType);

    const url = `/api/pdf?type=${encodeURIComponent(docType)}&personId=${encodeURIComponent(personId)}&archive=${shouldArchive ? "1" : "0"}&mode=${encodeURIComponent(archiveMode)}&ts=${Date.now()}`;
    const response = await fetch(url, { cache: "no-store" });
    const payload = await response.json().catch(() => null);
    if (!response.ok || !payload?.ok) {
      throw new Error(String(payload?.error || "PDF impossible"));
    }
    const archiveSaved = payload.archiveSaved === true;
    const archivePdfPath = String(payload.archivePdfPath || "");
    const archiveMetadataPath = String(payload.archiveMetadataPath || "");
    const archiveLocalPath = String(payload.archiveLocalPath || "");
    const archiveStoragePath = String(payload.archiveStoragePath || "");
    const archivePublicUrl = String(payload.archivePublicUrl || "");
    const archiveOpenLocalUrl = String(payload.archiveOpenLocalUrl || "");
    const archiveOpenRemoteUrl = String(payload.archiveOpenRemoteUrl || "");
    const archiveFilename = String(payload.archiveFilename || "");
    const archiveStorageStatus = String(payload.archiveStorageStatus || "");

    hidePdfProgressModal();
    const pdfOpenUrl = archiveOpenLocalUrl || archiveOpenRemoteUrl || archivePublicUrl || "";
    if (!openPdfUrlInBrowserWindow(popup, pdfOpenUrl)) {
      popup.location.href = "about:blank";
      showDataStatus(pdfOpenUrl ? "PDF GENERE - URL NON PDF BLOQUEE" : "PDF GENERE - URL D'OUVERTURE INTROUVABLE");
    }
    const finalArchivePath = archiveStoragePath || archivePdfPath || archiveLocalPath;
    const archiveDetails = {
      filename: archiveFilename,
      localPath: archiveLocalPath,
      openLocalUrl: archiveOpenLocalUrl,
      storagePath: archiveStoragePath,
      publicUrl: archivePublicUrl,
      openRemoteUrl: archiveOpenRemoteUrl || archivePublicUrl,
      storageStatus: archiveStorageStatus || (archiveStoragePath ? "synced" : "failed"),
    };
    if (person && finalArchivePath) {
      clearPendingPdfTaskFor(person.id, docType);
      await registerArchivedDocument(person, docType, finalArchivePath, archiveMetadataPath, archiveMode, archiveDetails);
      const archiveCheck = await verifyActiveArchiveOpenable(person.id, docType);
      if (archiveCheck.ok) {
        showDataStatus("LE DOCUMENT A ETE REGENERE ET VERIFIE.");
      } else {
        showDataStatus("LE DOCUMENT A ETE REGENERE, MAIS IL NE PEUT TOUJOURS PAS ETRE OUVERT. UTILISEZ LE MODE SECOURS.");
      }
    } else if (shouldArchive && !finalArchivePath) {
      showDataStatus("PDF OUVERT - ARCHIVAGE NON REALISE");
    } else if (!shouldArchive) {
      showDataStatus("PDF OUVERT - DOCUMENT NON SIGNE, UPLOAD SUPABASE NON LANCE");
    }
  } catch (error) {
    console.error(error);
    hidePdfProgressModal();
    try {
      popup.document.body.innerHTML =
        "<div style=\"font-family:Segoe UI,Arial,sans-serif;padding:24px;color:#8f4a32;\">GENERATION PDF IMPOSSIBLE</div>";
    } catch (popupError) {
      console.error(popupError);
    }
    showDataStatus("GENERATION PDF IMPOSSIBLE");
  }
}

function applyRequestedPdfFocus() {
  const params = new URLSearchParams(window.location.search);
  const requestedDocType = normalizeText(params.get("focusPdf") || "");
  if (!requestedDocType) {
    return;
  }

  const targetDocType = requestedDocType === "EXIT" ? "exit" : "arrival";
  const targetPage = targetDocType === "exit" ? "exit-document" : "arrival-document";
  if ((document.body.dataset.page || "") !== targetPage) {
    return;
  }

  const button = document.querySelector(`.js-open-pdf[data-doc-type="${targetDocType}"]`);
  if (!(button instanceof HTMLButtonElement)) {
    return;
  }

  const nextUrl = new URL(window.location.href);
  nextUrl.searchParams.delete("focusPdf");
  window.history.replaceState({}, "", nextUrl);

  button.focus();
  button.scrollIntoView({ behavior: "smooth", block: "center" });
  button.classList.remove("button--pdf-attention");
  void button.offsetWidth;
  button.classList.add("button--pdf-attention");
  window.setTimeout(() => {
    button.classList.remove("button--pdf-attention");
  }, 6200);
}

async function generatePdfArchiveSilently(person, docType) {
  if (!person || !isDocumentFullySigned(person, docType)) {
    return false;
  }
  if (getDataBackendMode() !== "LOCAL_API") {
    return false;
  }
  if (findReusableArchivedDocument(person, docType)) {
    return false;
  }

  const archiveMode = getDocumentArchiveMode(person, docType);
  const url = `/api/pdf?type=${encodeURIComponent(docType)}&personId=${encodeURIComponent(person.id)}&archive=1&mode=${encodeURIComponent(archiveMode)}&ts=${Date.now()}`;
  const response = await fetch(url, { cache: "no-store" });
  const payload = await response.json().catch(() => null);
  if (!response.ok || !payload?.ok) {
    throw new Error(`AUTO PDF IMPOSSIBLE (${docType}/${person.id})`);
  }
  const archivePdfPath = String(payload.archivePdfPath || "");
  const archiveMetadataPath = String(payload.archiveMetadataPath || "");
  const archiveLocalPath = String(payload.archiveLocalPath || "");
  const archiveStoragePath = String(payload.archiveStoragePath || "");
  const archivePublicUrl = String(payload.archivePublicUrl || "");
  const archiveOpenLocalUrl = String(payload.archiveOpenLocalUrl || "");
  const archiveOpenRemoteUrl = String(payload.archiveOpenRemoteUrl || "");
  const archiveFilename = String(payload.archiveFilename || "");
  const archiveStorageStatus = String(payload.archiveStorageStatus || "");

  const finalArchivePath = archiveStoragePath || archivePdfPath || archiveLocalPath;
  if (!finalArchivePath) {
    return false;
  }
  await registerArchivedDocument(person, docType, finalArchivePath, archiveMetadataPath, archiveMode, {
    filename: archiveFilename,
    localPath: archiveLocalPath,
    openLocalUrl: archiveOpenLocalUrl,
    storagePath: archiveStoragePath,
    publicUrl: archivePublicUrl,
    openRemoteUrl: archiveOpenRemoteUrl || archivePublicUrl,
    storageStatus: archiveStorageStatus || (archiveStoragePath ? "synced" : "failed"),
  });
  return true;
}

function queueAutoGenerateSignedDocumentsPdfIfMissing() {
  if (!AUTO_GENERATE_SIGNED_PDFS) {
    updateDocumentPdfButtonsState();
    return;
  }
  if (state.autoPdfGenerationInFlight) {
    return;
  }
  window.requestAnimationFrame(() => {
    window.setTimeout(() => {
      autoGenerateSignedDocumentsPdfIfMissing().catch((error) => {
        console.error(error);
      });
    }, 16);
  });
}

async function autoGenerateSignedDocumentsPdfIfMissing() {
  if (state.autoPdfGenerationInFlight || !state.data || getDataBackendMode() !== "LOCAL_API") {
    return;
  }

  const candidates = [];
  (state.data.personnes || []).forEach((person) => {
    ["arrival", "exit"].forEach((docType) => {
      if (!isDocumentFullySigned(person, docType)) {
        return;
      }
      if (findReusableArchivedDocument(person, docType)) {
        return;
      }
      const key = `${person.id}:${docType}:${getDocumentFingerprint(person, docType)}`;
      if (state.autoPdfGeneratedKeys.has(key)) {
        return;
      }
      candidates.push({ person, docType, key });
    });
  });

  if (!candidates.length) {
    return;
  }

  const generatedLabels = [];
  const BATCH_SIZE = 4;
  state.autoPdfGenerationInFlight = true;
  try {
    for (let index = 0; index < candidates.length; index += 1) {
      const candidate = candidates[index];
      try {
        const generated = await generatePdfArchiveSilently(candidate.person, candidate.docType);
        if (generated) {
          state.autoPdfGeneratedKeys.add(candidate.key);
          generatedLabels.push(
            `${getDocumentTypeLabel(candidate.docType)} - ${candidate.person.nom || ""} ${candidate.person.prenom || ""}`.trim()
          );
          showDataStatus(
            `UN DOCUMENT ${getDocumentTypeLabel(candidate.docType)} A ETE CREE - ${candidate.person.nom || ""} ${candidate.person.prenom || ""}`.trim()
          );
        }
      } catch (error) {
        console.error(error);
      }

      if ((index + 1) % BATCH_SIZE === 0 && index + 1 < candidates.length) {
        await new Promise((resolve) => window.setTimeout(resolve, 0));
      }
    }
    if (generatedLabels.length) {
      window.alert(`UN DOCUMENT A ETE CREE :\n${generatedLabels.join("\n")}`);
    }
    updateDocumentPdfButtonsState();
  } finally {
    state.autoPdfGenerationInFlight = false;
  }
}

function buildSignatureValidationMap(data) {
  const map = new Map();
  (data?.personnes || []).forEach((person) => {
    ["arrival", "exit"].forEach((docType) => {
      ["personnel", "representant"].forEach((signer) => {
        const validatedAt = getSignatureValidationDate(person, docType, signer);
        if (!validatedAt) {
          return;
        }
        map.set(`${person.id}:${docType}:${signer}`, String(validatedAt));
      });
    });
  });
  return map;
}

function notifyFullySignedDocumentsOnReload(previousSignatureValidationMap = new Map()) {
  if (!state.data) {
    return;
  }
  clearPendingPdfTaskIfArchived();
  const currentSignatureValidationMap = buildSignatureValidationMap(state.data);
  const newEvents = [];
  currentSignatureValidationMap.forEach((validatedAt, key) => {
    if (previousSignatureValidationMap.get(key) === validatedAt) {
      return;
    }
    const [personId, docType, signer] = String(key).split(":");
    const validatedAtMs = Date.parse(String(validatedAt || ""));
    if (!personId || !docType || !signer || !Number.isFinite(validatedAtMs)) {
      return;
    }
    newEvents.push({ personId, docType, signer, validatedAt, validatedAtMs });
  });
  const latestRequestFromNewEvent = newEvents.reduce((latest, entry) => {
    if (!latest || entry.validatedAtMs > latest.validatedAtMs) {
      return entry;
    }
    return latest;
  }, null);

  const pendingTask = getPendingPdfTaskFromStorage();
  const latestPendingPdf =
    pendingTask && pendingTask.personId && pendingTask.docType && pendingTask.validatedAt
      ? {
          source: "pending",
          personId: String(pendingTask.personId),
          docType: String(pendingTask.docType),
          signer: "representant",
          validatedAt: String(pendingTask.validatedAt),
          validatedAtMs: Date.parse(String(pendingTask.validatedAt)),
        }
      : null;

  const latestRequest = latestRequestFromNewEvent || latestPendingPdf;
  if (!latestRequest) {
    return;
  }
  const snoozeKey = `${latestRequest.personId}:${latestRequest.docType}`;
  const snoozedUntil = Number(reminderSnoozeMap[snoozeKey] || 0);
  if (Number.isFinite(snoozedUntil) && snoozedUntil > Date.now()) {
    return;
  }

  const labels = [];
  const person = (state.data.personnes || []).find(
    (candidate) => String(candidate.id || "") === latestRequest.personId
  );
  if (person && isDocumentFullySigned(person, latestRequest.docType)) {
    const signatureDate = getSignatureValidationDate(person, latestRequest.docType, latestRequest.signer);
    const isPendingReminder = String(latestRequest.source || "") === "pending";
    if (isPendingReminder || String(signatureDate || "") === String(latestRequest.validatedAt || "")) {
      const hasArchive = Boolean(findReusableArchivedDocument(person, latestRequest.docType));
      const key = `SIG:${person.id}:${latestRequest.docType}:${latestRequest.signer}:${latestRequest.validatedAt}`;
      if (hasArchive) {
        clearPendingPdfTaskFor(person.id, latestRequest.docType);
        state.signedDocumentsPopupSeenKeys.add(key);
        return;
      }
      if (state.signedDocumentsPopupSeenKeys.has(key)) {
        return;
      }
      labels.push(
        `${getDocumentTypeLabel(latestRequest.docType)} - ${person.nom || ""} ${person.prenom || ""}`.trim()
      );
    }
  }

  if (!labels.length) {
    return;
  }
  // Mode silencieux: ne jamais afficher de popup de focus document.
  reminderSnoozeMap[snoozeKey] = Date.now() + 180 * 1000;
  setPendingPdfTaskToStorage({
    personId: latestRequest.personId,
    docType: latestRequest.docType,
    validatedAt: latestRequest.validatedAt,
  });
  try {
    localStorage.setItem(PENDING_PDF_REMINDER_SNOOZE_KEY, JSON.stringify(reminderSnoozeMap));
  } catch (error) {
    // ignore storage failures
  }
  if (person && latestRequest?.docType) {
    const seenKey = `SIG:${person.id}:${latestRequest.docType}:${latestRequest.signer}:${latestRequest.validatedAt}`;
    state.signedDocumentsPopupSeenKeys.add(seenKey);
  }
  showDataStatus("NOUVEAU DOCUMENT DETECTE (MODE SILENCIEUX)");
  return;
}

function getArchiveSortValue(entry, key, resolveArchiveDisplayData) {
  const display = resolveArchiveDisplayData(entry);
  switch (key) {
    case "nom":
      return display.nom || "";
    case "prenom":
      return display.prenom || "";
    case "typeDocument":
      return entry?.typeDocument || "";
    case "dateDocument":
      return Date.parse(String(entry?.dateDocument || "")) || 0;
    case "heureDocument":
      return Date.parse(String(entry?.dateArchivage || "")) || 0;
    case "sites":
      return display.sites || "";
    case "statutSignature":
      return String(
        entry?.__workflowStatus ||
          normalizeText(entry?.statutSignature || "") ||
          "EN ATTENTE DE SIGNATURE"
      ).trim();
      case "totalEffets":
        return Number(getArchiveDisplayedTotalEffets(entry) || 0);
    case "totalFacturable":
      return normalizeAmount(entry?.totalFacturable || 0);
    case "version":
      return getDocumentArchiveVersionLabel(entry) || "";
    default:
      return "";
  }
}

function sortArchivesForTable(entries, resolveArchiveDisplayData) {
  const sort = getEffectTableSort("documentsArchives");
  const numericKeys = new Set(["dateDocument", "heureDocument", "totalEffets", "totalFacturable"]);
  return [...entries].sort((left, right) => {
    const primary = compareEffectValues(
      getArchiveSortValue(left, sort.key, resolveArchiveDisplayData),
      getArchiveSortValue(right, sort.key, resolveArchiveDisplayData),
      numericKeys.has(sort.key)
    );
    if (primary !== 0) {
      return sort.dir === "asc" ? primary : -primary;
    }

    const nomCompare = compareTextValues(
      getArchiveSortValue(left, "nom", resolveArchiveDisplayData),
      getArchiveSortValue(right, "nom", resolveArchiveDisplayData)
    );
    if (nomCompare !== 0) {
      return nomCompare;
    }

    const prenomCompare = compareTextValues(
      getArchiveSortValue(left, "prenom", resolveArchiveDisplayData),
      getArchiveSortValue(right, "prenom", resolveArchiveDisplayData)
    );
    if (prenomCompare !== 0) {
      return prenomCompare;
    }

    return compareTextValues(String(left?.id || ""), String(right?.id || ""));
  });
}

function applyFiltersToForm(form) {
  if (!form) {
    return;
  }
  const filters = state.filters || DEFAULT_FILTERS;
  const assign = (name, value) => {
    const field = form.elements[name];
    if (!field) {
      return;
    }
    field.value = value || "";
  };
  assign("search", filters.search);
  assign("person-picker-search", filters.search);
  assign("site", filters.site);
  assign("typePersonnel", filters.typePersonnel);
  assign("typeContrat", filters.typeContrat);
  assign("statutDossier", filters.statutDossier);
  assign("statutObjet", filters.statutObjet);
  assign("typeEffet", filters.typeEffet);
}

function clearFormSearchFields(form) {
  if (!(form instanceof HTMLFormElement)) {
    return;
  }
  form
    .querySelectorAll("input[type=\"search\"], input[name*=\"search\" i], input[id*=\"search\" i]")
    .forEach((field) => {
      if (field instanceof HTMLInputElement) {
        field.value = "";
      }
    });
}

function clearSearchInputsOnInitialLoad() {
  document
    .querySelectorAll("input[type=\"search\"], input[name*=\"search\" i], input[id*=\"search\" i]")
    .forEach((field) => {
      if (!(field instanceof HTMLInputElement)) {
        return;
      }
      if (field.id === "person-picker-search" && document.body?.dataset?.page === "overview") {
        const selectedPerson = getCurrentPerson();
        if (selectedPerson) {
          const label = getPersonPickerLabel(selectedPerson);
          field.value = label;
          field.defaultValue = label;
          field.setAttribute("autocomplete", "off");
          return;
        }
      }
      if (field.name === "archiveSearch" && document.body?.dataset?.page === "documents-archives") {
        const selectedPerson = getCurrentPerson();
        if (selectedPerson) {
          const label = getPersonPickerLabel(selectedPerson);
          field.value = label;
          field.defaultValue = label;
          field.setAttribute("autocomplete", "off");
          return;
        }
      }
      field.value = "";
      field.defaultValue = "";
      field.setAttribute("autocomplete", "off");
    });
}

function bindSearchClearOnBrowserEvents() {
  if (state.searchClearBrowserEventsBound) {
    return;
  }
  const applyClear = () => {
    window.setTimeout(() => {
      clearSearchInputsOnInitialLoad();
    }, 0);
    window.setTimeout(() => {
      clearSearchInputsOnInitialLoad();
    }, 180);
    window.setTimeout(() => {
      clearSearchInputsOnInitialLoad();
    }, 700);
  };
  window.addEventListener("load", applyClear);
  window.addEventListener("pageshow", applyClear);
  state.searchClearBrowserEventsBound = true;
}
function bindFilterForms() {
  document.querySelectorAll(".js-filter-form").forEach((form) => {
    applyFiltersToForm(form);

    const scheduleFilterUpdate = () => {
      if (state.filterInputDebounceTimerId) {
        window.clearTimeout(state.filterInputDebounceTimerId);
      }
      state.filterInputDebounceTimerId = window.setTimeout(() => {
        state.filterInputDebounceTimerId = 0;
        state.filters = {
          ...DEFAULT_FILTERS,
          ...readFilters(form),
        };
        saveNavigationContext({ filters: state.filters, urgentMode: state.urgentMode });
        schedulePageRender();
      }, FILTER_INPUT_DEBOUNCE_MS);
    };

    const applyFullReset = () => {
      if (state.filterInputDebounceTimerId) {
        window.clearTimeout(state.filterInputDebounceTimerId);
        state.filterInputDebounceTimerId = 0;
      }
      state.filters = { ...DEFAULT_FILTERS };
      state.urgentMode = false;
      resetTableSortsForCurrentPage();
      saveNavigationContext({ filters: state.filters, personId: "", urgentMode: false });
      setCurrentPersonId("", "replace");
      clearFormSearchFields(form);
      applyFiltersToForm(form);
      updateUrgencyModeUi();
      schedulePageRender();
    };

    form.oninput = (event) => {
      const target = event?.target instanceof HTMLElement ? event.target : null;
      const searchField = form.elements.search || form.elements["person-picker-search"];
      const searchClearedByField =
        target &&
        searchField &&
        target === searchField &&
        !String(searchField.value || "").trim();
      if (searchClearedByField) {
        applyFullReset();
        return;
      }
      scheduleFilterUpdate();
    };

    form.onreset = () => {
      clearFormSearchFields(form);
      window.setTimeout(() => {
        applyFullReset();
      }, 0);
    };

    const searchField = form.elements.search || form.elements["person-picker-search"];
    if (searchField) {
      searchField.addEventListener("search", () => {
        if (!String(searchField.value || "").trim()) {
          applyFullReset();
        }
      });
    }
  });
}

function syncFilterFormsFromState() {
  document.querySelectorAll(".js-filter-form").forEach((form) => applyFiltersToForm(form));
}

function schedulePageRender() {
  if (state.pageRenderRafId) {
    return;
  }

  const runScheduledRender = () => {
    if (!state.pageRenderRafId) {
      return;
    }
    if (state.pageRenderTimeoutId) {
      window.clearTimeout(state.pageRenderTimeoutId);
      state.pageRenderTimeoutId = 0;
    }
    state.pageRenderRafId = 0;
    renderPage();
    if (typeof updateAllFilterResetHighlights === "function") {
      updateAllFilterResetHighlights();
    }
  };

  state.pageRenderRafId = window.requestAnimationFrame(() => {
    runScheduledRender();
  });
  state.pageRenderTimeoutId = window.setTimeout(runScheduledRender, 120);
}

function bindArchiveFilterForm() {
  const form = document.getElementById("documents-archives-filter-form");
  if (!form) {
    return;
  }
  form.dataset.archiveFiltersBound = "1";
  const getArchiveFilterField = (fieldName) => form.querySelector(`[name="${fieldName}"]`);

  const resetArchiveFilters = () => {
    clearFormSearchFields(form);
    const searchField = getArchiveFilterField("archiveSearch");
    if (searchField instanceof HTMLInputElement) {
      searchField.value = "";
      searchField.defaultValue = "";
    }
    ["archiveTypeDocument", "archiveSite", "archiveStatutSignature"].forEach((fieldName) => {
      const field = getArchiveFilterField(fieldName);
      if (field instanceof HTMLSelectElement) {
        field.value = "";
      }
    });
  };

  const applyArchiveReset = () => {
    setCurrentPersonId("", "replace");
    resetArchiveFilters();
    state.tableSorts.documentsArchives = { key: "nom", dir: "asc" };
    state.listRenderCache.documentsArchives = "";
    schedulePageRender();
  };

  const applyArchiveFilters = () => {
    state.listRenderCache.documentsArchives = "";
    schedulePageRender();
  };
  form.oninput = applyArchiveFilters;
  form.onchange = applyArchiveFilters;
  ["archiveSearch", "archiveTypeDocument", "archiveSite", "archiveStatutSignature"].forEach((fieldName) => {
    const field = getArchiveFilterField(fieldName);
    if (!(field instanceof HTMLElement)) {
      return;
    }
    field.addEventListener("input", applyArchiveFilters);
    field.addEventListener("change", applyArchiveFilters);
  });

  form.onreset = (event) => {
    event.preventDefault();
    setCurrentPersonId("", "replace");
    resetArchiveFilters();
    state.tableSorts.documentsArchives = { key: "nom", dir: "asc" };
    state.listRenderCache.documentsArchives = "";
    schedulePageRender();
  };

  const searchField = getArchiveFilterField("archiveSearch");
  if (searchField) {
    searchField.value = "";
    searchField.addEventListener("search", () => {
      if (!String(searchField.value || "").trim()) {
        applyArchiveReset();
      }
    });
  }
}

function bindAddPersonForm() {
  const form = document.getElementById("add-person-form");
  if (!form) {
    return;
  }

  form.onsubmit = (event) => {
    event.preventDefault();

    if (!state.data) {
      showDataStatus("DONNEES NON CHARGEES");
      return;
    }

    const formData = new FormData(form);
    const setAddFieldMissingState = (fieldName, isMissing) => {
      const field = form.elements[fieldName];
      if (!(field instanceof HTMLElement)) {
        return;
      }
      const node = field.closest(".field");
      if (node) {
        node.classList.toggle("field--missing", Boolean(isMissing));
      }
    };
    const selectedSites = readSelectedSites(form, "add");
    const addSiteField = form.querySelector("#add-site-selector")?.closest(".field");
    if (addSiteField) {
      addSiteField.classList.toggle("field--missing", selectedSites.length === 0);
    }
    if (!selectedSites.length) {
      const message = "AU MOINS UN SITE EST OBLIGATOIRE";
      showDataStatus(message);
      window.alert(message);
      const firstSiteInput = form.querySelector('#add-site-selector input[name="addSites"]');
      if (firstSiteInput instanceof HTMLElement) {
        firstSiteInput.focus();
      }
      return;
    }
    const person = {
      id: getNextId("P", state.data.personnes || []),
      nom: normalizeText(formData.get("nom")),
      prenom: normalizeText(formData.get("prenom")),
      sitesAffectation: selectedSites,
      site: "",
      typePersonnel: normalizeText(formData.get("typePersonnel")),
      typeContrat: normalizeText(formData.get("typeContrat")),
      dateEntree: normalizeDateString(formData.get("dateEntree")),
      dateSortiePrevue: normalizeDateString(formData.get("dateSortiePrevue")),
      dateSortieReelle: normalizeDateString(formData.get("dateSortieReelle")),
      effetsConfies: [],
    };

    const addDateChecks = [
      { value: person.dateEntree, label: "DATE D'ENTREE", field: form.elements.dateEntree },
      { value: person.dateSortiePrevue, label: "DATE DE SORTIE PREVUE", field: form.elements.dateSortiePrevue },
      { value: person.dateSortieReelle, label: "DATE DE SORTIE REELLE", field: form.elements.dateSortieReelle },
    ];
    for (const check of addDateChecks) {
      const validation = validateDateFieldFormat(check.value, check.label);
      if (check.field instanceof HTMLElement) {
        const fieldName = String(check.field.getAttribute("name") || "");
        if (fieldName) {
          setAddFieldMissingState(fieldName, !validation.ok);
        }
      }
      if (!validation.ok) {
        showDataStatus(validation.message);
        if (check.field instanceof HTMLElement) {
          check.field.focus();
        }
        return;
      }
    }

    if (!person.nom && !person.prenom) {
      person.nom = "PERSONNE";
      person.prenom = person.id;
    }
    person.site = getPersonSiteLabel(person);

    const duplicate = (state.data.personnes || []).some(
      (entry) =>
        entry.nom === person.nom &&
        entry.prenom === person.prenom &&
        haveSameSites(getPersonSites(entry), person.sitesAffectation)
    );

    if (duplicate) {
      showDataStatus("CETTE PERSONNE EXISTE DEJA SUR CE SITE");
      return;
    }

    pushUndoSnapshot("AJOUT PERSONNE");
    state.data.personnes.push(person);
    markDirty();
    form.reset();
    renderSiteSelector("add-site-selector", "add", []);
    schedulePageRender();
    showActionStatus("create", `PERSONNE AJOUTEE : ${person.nom} ${person.prenom}`);
    setCurrentPersonId(person.id);
    openPersonSheet(person.id);
  };
}

function bindPersonSheetForm() {
  const form = document.getElementById("person-sheet-form");
  if (!form) {
    return;
  }

  const addButton = document.getElementById("sheet-add-person");
  const arrivalDocumentButton = document.getElementById("sheet-open-arrival-document");
  const exitDocumentButton = document.getElementById("sheet-open-exit-document");
  const arrivalPdfButton = document.getElementById("sheet-open-arrival-pdf");
  const exitPdfButton = document.getElementById("sheet-open-exit-pdf");
  const deletePersonButton = document.getElementById("sheet-delete-person");
  const typeContratField = form.elements.sheetTypeContrat;

  const setSheetFieldMissingState = (fieldName, isMissing) => {
    const field = form.elements[fieldName];
    if (!(field instanceof HTMLElement)) {
      return;
    }
    const node = field.closest(".field");
    if (node) {
      node.classList.toggle("field--missing", Boolean(isMissing));
    }
  };

  const updateSheetRequiredHighlights = () => {
    const formData = new FormData(form);
    const nom = normalizeText(formData.get("sheetNom"));
    const prenom = normalizeText(formData.get("sheetPrenom"));
    const fonction = normalizeText(formData.get("sheetFonction"));
    const selectedSites = readSelectedSites(form, "sheet");
    const typePersonnel = normalizeText(formData.get("sheetTypePersonnel"));
    const typeContrat = normalizeText(formData.get("sheetTypeContrat"));
    const dateEntree = String(formData.get("sheetDateEntree") || "").trim();
    const needsExpectedExitDate = ["CDD", "INTERIMAIRE"].includes(typeContrat);
    const dateSortiePrevue = String(formData.get("sheetDateSortiePrevue") || "").trim();
    const dateSortieReelle = String(formData.get("sheetDateSortieReelle") || "").trim();
    const dateEntreeValidation = validateDateFieldFormat(dateEntree, "DATE D'ENTREE");
    const dateSortiePrevueValidation = validateDateFieldFormat(dateSortiePrevue, "DATE DE SORTIE PREVUE");
    const dateSortieReelleValidation = validateDateFieldFormat(dateSortieReelle, "DATE DE SORTIE REELLE");

    setSheetFieldMissingState("sheetNom", !nom);
    setSheetFieldMissingState("sheetPrenom", !prenom);
    setSheetFieldMissingState("sheetFonction", !fonction);
    setSheetFieldMissingState("sheetTypePersonnel", !typePersonnel);
    setSheetFieldMissingState("sheetTypeContrat", !typeContrat);
    setSheetFieldMissingState("sheetDateEntree", !dateEntree || !dateEntreeValidation.ok);
    setSheetFieldMissingState(
      "sheetDateSortiePrevue",
      (needsExpectedExitDate && !dateSortiePrevue) || !dateSortiePrevueValidation.ok
    );
    setSheetFieldMissingState("sheetDateSortieReelle", !dateSortieReelleValidation.ok);

    const siteField = form.querySelector("#sheet-site-selector")?.closest(".field");
    if (siteField) {
      siteField.classList.toggle("field--missing", selectedSites.length === 0);
    }
  };

  const updateSheetContractDateRequirement = () => {
    const normalizedTypeContrat = normalizeText(form.elements.sheetTypeContrat?.value || "");
    const needsExpectedExitDate = ["CDD", "INTERIMAIRE"].includes(normalizedTypeContrat);
    const dateSortiePrevueField = form.elements.sheetDateSortiePrevue;
    const dateSortiePrevueNode = dateSortiePrevueField instanceof HTMLElement
      ? dateSortiePrevueField.closest(".field")
      : null;
    if (dateSortiePrevueField instanceof HTMLElement) {
      dateSortiePrevueField.required = needsExpectedExitDate;
    }
    if (dateSortiePrevueNode) {
      dateSortiePrevueNode.classList.toggle("field--key", needsExpectedExitDate);
    }
    updateSheetRequiredHighlights();
  };

  const validateSheetRequiredFields = (formData) => {
    const nom = normalizeText(formData.get("sheetNom"));
    if (!nom) {
      showDataStatus("LE NOM EST OBLIGATOIRE");
      form.elements.sheetNom?.focus();
      return false;
    }

    const prenom = normalizeText(formData.get("sheetPrenom"));
    if (!prenom) {
      showDataStatus("LE PRENOM EST OBLIGATOIRE");
      form.elements.sheetPrenom?.focus();
      return false;
    }

    const fonction = normalizeText(formData.get("sheetFonction"));
    if (!fonction) {
      showDataStatus("LA FONCTION EST OBLIGATOIRE");
      form.elements.sheetFonction?.focus();
      return false;
    }

    const selectedSites = readSelectedSites(form, "sheet");
    if (!selectedSites.length) {
      showDataStatus("AU MOINS UN SITE EST OBLIGATOIRE");
      const firstSiteInput = form.querySelector('#sheet-site-selector input[name="sheetSites"]');
      if (firstSiteInput instanceof HTMLElement) {
        firstSiteInput.focus();
      }
      return false;
    }

    const typePersonnel = normalizeText(formData.get("sheetTypePersonnel"));
    if (!typePersonnel) {
      showDataStatus("LE TYPE DE PERSONNEL EST OBLIGATOIRE");
      form.elements.sheetTypePersonnel?.focus();
      return false;
    }

    const typeContrat = normalizeText(formData.get("sheetTypeContrat"));
    if (!typeContrat) {
      showDataStatus("LE TYPE DE CONTRAT EST OBLIGATOIRE");
      form.elements.sheetTypeContrat?.focus();
      return false;
    }

    const needsExpectedExitDate = ["CDD", "INTERIMAIRE"].includes(typeContrat);
    const dateSortiePrevue = String(formData.get("sheetDateSortiePrevue") || "").trim();
    if (needsExpectedExitDate && !dateSortiePrevue) {
      showDataStatus("LA DATE DE SORTIE PREVUE EST OBLIGATOIRE POUR CDD / INTERIMAIRE");
      form.elements.sheetDateSortiePrevue?.focus();
      updateSheetRequiredHighlights();
      return false;
    }

    const dateEntree = String(formData.get("sheetDateEntree") || "").trim();
    if (!dateEntree) {
      showDataStatus("LA DATE D'ENTREE EST OBLIGATOIRE");
      form.elements.sheetDateEntree?.focus();
      updateSheetRequiredHighlights();
      return false;
    }

    const sheetDateChecks = [
      { value: dateEntree, label: "DATE D'ENTREE", field: form.elements.sheetDateEntree },
      { value: dateSortiePrevue, label: "DATE DE SORTIE PREVUE", field: form.elements.sheetDateSortiePrevue },
      { value: String(formData.get("sheetDateSortieReelle") || "").trim(), label: "DATE DE SORTIE REELLE", field: form.elements.sheetDateSortieReelle },
    ];
    for (const check of sheetDateChecks) {
      const validation = validateDateFieldFormat(check.value, check.label);
      if (!validation.ok) {
        showDataStatus(validation.message);
        if (check.field instanceof HTMLElement) {
          check.field.focus();
        }
        updateSheetRequiredHighlights();
        return false;
      }
    }

    updateSheetRequiredHighlights();
    return true;
  };

  if (typeContratField instanceof HTMLElement) {
    typeContratField.addEventListener("change", () => {
      updateSheetContractDateRequirement();
    });
  }
  [
    "sheetNom",
    "sheetPrenom",
    "sheetFonction",
    "sheetTypePersonnel",
    "sheetTypeContrat",
    "sheetDateEntree",
    "sheetDateSortiePrevue",
    "sheetDateSortieReelle",
  ].forEach((fieldName) => {
    const field = form.elements[fieldName];
    if (!(field instanceof HTMLElement)) {
      return;
    }
    field.addEventListener("input", updateSheetRequiredHighlights);
    field.addEventListener("change", updateSheetRequiredHighlights);
  });
  form.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) {
      return;
    }
    if (target.name === "sheetSites") {
      updateSheetRequiredHighlights();
    }
  });
  updateSheetContractDateRequirement();
  updateSheetRequiredHighlights();

  const buildPersonFromSheetForm = () => {
    const formData = new FormData(form);
    if (!validateSheetRequiredFields(formData)) {
      return null;
    }
    const person = {
      id: getNextId("P", state.data?.personnes || []),
      nom: normalizeText(formData.get("sheetNom")),
      prenom: normalizeText(formData.get("sheetPrenom")),
      fonction: normalizeText(formData.get("sheetFonction")),
      sitesAffectation: readSelectedSites(form, "sheet"),
      site: "",
      typePersonnel: normalizeText(formData.get("sheetTypePersonnel")),
      typeContrat: normalizeText(formData.get("sheetTypeContrat")),
      dateEntree: normalizeDateString(formData.get("sheetDateEntree")),
      email: String(formData.get("sheetEmail") || "").trim(),
      phoneMobile: String(formData.get("sheetPhoneMobile") || "").trim(),
      dateSortiePrevue: normalizeDateString(formData.get("sheetDateSortiePrevue")),
      dateSortieReelle: normalizeDateString(formData.get("sheetDateSortieReelle")),
      effetsConfies: [],
    };

    person.site = getPersonSiteLabel(person);
    return person;
  };

  const autoSaveAfterPersonChange = async (successLabel) => {
    await saveDataToFile({
      silent: true,
      reloadAfter: true,
      successText: successLabel || "SAUVEGARDE AUTOMATIQUE",
    });
  };

  form.onsubmit = async (event) => {
    event.preventDefault();
    const person = getCurrentPerson();
    if (!person) {
      showDataStatus("AUCUNE PERSONNE SELECTIONNEE");
      return;
    }

    const formData = new FormData(form);
    if (!validateSheetRequiredFields(formData)) {
      return;
    }
    pushUndoSnapshot("MODIFICATION FICHE PERSONNE");
    person.nom = normalizeText(formData.get("sheetNom"));
    person.prenom = normalizeText(formData.get("sheetPrenom"));
    person.fonction = normalizeText(formData.get("sheetFonction"));
    person.sitesAffectation = readSelectedSites(form, "sheet");
    person.site = getPersonSiteLabel(person);
    person.typePersonnel = normalizeText(formData.get("sheetTypePersonnel"));
    person.typeContrat = normalizeText(formData.get("sheetTypeContrat"));
    person.dateEntree = normalizeDateString(formData.get("sheetDateEntree"));
    person.email = String(formData.get("sheetEmail") || "").trim();
    person.phoneMobile = String(formData.get("sheetPhoneMobile") || "").trim();
    person.dateSortiePrevue = normalizeDateString(formData.get("sheetDateSortiePrevue"));
    person.dateSortieReelle = normalizeDateString(formData.get("sheetDateSortieReelle"));

    markDirty();
    schedulePageRender();
    showActionStatus("update", `FICHE MISE A JOUR : ${person.nom} ${person.prenom}`);
    await autoSaveAfterPersonChange("FICHE MISE A JOUR - SAUVEGARDE AUTOMATIQUE");
  };

  if (addButton) {
    addButton.onclick = async () => {
      if (!state.data?.personnes) {
        showDataStatus("DONNEES NON CHARGEES");
        return;
      }

      const person = buildPersonFromSheetForm();
      if (!person) {
        return;
      }
      const duplicate = (state.data.personnes || []).some(
        (entry) =>
          entry.nom === person.nom &&
          entry.prenom === person.prenom &&
          haveSameSites(getPersonSites(entry), person.sitesAffectation)
      );

      if (duplicate) {
        showDataStatus("CETTE PERSONNE EXISTE DEJA SUR CE SITE");
        return;
      }

      pushUndoSnapshot("AJOUT PERSONNE");
      state.data.personnes.push(person);
      setCurrentPersonId(person.id);
      markDirty();
      schedulePageRender();
      showActionStatus("create", `PERSONNE AJOUTEE : ${person.nom} ${person.prenom}`);
      await autoSaveAfterPersonChange("PERSONNE AJOUTEE - SAUVEGARDE AUTOMATIQUE");
    };
  }

  if (arrivalDocumentButton) {
    arrivalDocumentButton.onclick = () => {
      const personId = getSheetTargetPersonId();
      if (!personId) {
        showDataStatus("AUCUNE PERSONNE SELECTIONNEE");
        return;
      }
      navigateWithAutoSave(`document-arrivee.html?personId=${personId}`);
    };
  }

  if (exitDocumentButton) {
    exitDocumentButton.onclick = () => {
      const personId = getSheetTargetPersonId();
      if (!personId) {
        showDataStatus("AUCUNE PERSONNE SELECTIONNEE");
        return;
      }
      navigateWithAutoSave(`document-sortie.html?personId=${personId}`);
    };
  }

  if (arrivalPdfButton) {
    arrivalPdfButton.onclick = () => openPdfDocument("arrival", getSheetTargetPersonId());
  }

  if (exitPdfButton) {
    exitPdfButton.onclick = () => openPdfDocument("exit", getSheetTargetPersonId());
  }

  if (deletePersonButton) {
    deletePersonButton.onclick = () => {
      const personId = getSheetTargetPersonId();
      if (!personId) {
        showDataStatus("AUCUNE PERSONNE SELECTIONNEE");
        return;
      }
      deletePerson(personId);
    };
  }
}

function updateSheetDocumentButtons(person) {
  const arrivalDocumentButton = document.getElementById("sheet-open-arrival-document");
  const exitDocumentButton = document.getElementById("sheet-open-exit-document");
  const arrivalPdfButton = document.getElementById("sheet-open-arrival-pdf");
  const exitPdfButton = document.getElementById("sheet-open-exit-pdf");
  const addPersonButton = document.getElementById("sheet-add-person");
  const savePersonButton =
    document.getElementById("sheet-save-person") ||
    document.querySelector('#person-sheet-form button[type="submit"]');
  const deletePersonButton = document.getElementById("sheet-delete-person");
  const isDisabled = !person;
  const isEditingPerson = Boolean(person);

  [
    arrivalDocumentButton,
    exitDocumentButton,
    arrivalPdfButton,
    exitPdfButton,
    savePersonButton,
    deletePersonButton,
  ].forEach((button) => {
    if (!button) {
      return;
    }
    button.disabled = isDisabled;
  });

  if (addPersonButton) {
    addPersonButton.disabled = false;
    addPersonButton.classList.toggle("button--primary", !isEditingPerson);
    addPersonButton.classList.toggle("button--secondary", isEditingPerson);
  }

  if (savePersonButton) {
    savePersonButton.classList.toggle("button--primary", isEditingPerson);
    savePersonButton.classList.toggle("button--secondary", !isEditingPerson);
  }
}

function bindEffectForm() {
  const form = document.getElementById("effect-form");
  if (!form) {
    return;
  }

  const addButton = document.getElementById("effect-add-button");
  const updateButton = document.getElementById("effect-update-button");
  const deleteButton = document.getElementById("effect-delete-button");
  const cancelButton = document.getElementById("effect-cancel-button");
  const resetFieldsButton = document.getElementById("effect-reset-fields-button");
  const typeField = form.elements.typeEffet;
  const referenceSiteField = form.elements.referenceSite;
  const replacementDateField = form.elements.dateRemplacement;
  if (typeField) {
    typeField.onchange = () => {
      const person = getCurrentPerson();
      hydrateEffectReferenceSiteSelect(person, "", typeField.value);
      hydrateReferenceSelect(person || "", typeField.value, "", getSelectedEffectReferenceSite());
      updateEffectFormMode(typeField.value);
      syncReplacementCostField();
      focusNextEffectKeyField(form, "typeEffet");
      updateEffectRequiredHighlights(form);
    };
  }
  if (referenceSiteField) {
    referenceSiteField.onchange = () => {
      const person = getCurrentPerson();
      hydrateReferenceSelect(person || "", form.elements.typeEffet.value, "", getSelectedEffectReferenceSite());
      syncReplacementCostField();
      focusNextEffectKeyField(form, "referenceSite");
      updateEffectRequiredHighlights(form);
    };
  }
  if (form.elements.statutManuel) {
    form.elements.statutManuel.onchange = () => {
      syncReplacementCostField();
      updateEffectRequiredHighlights(form);
      updateManualStatusCriticalState(form);
    };
  }
  if (replacementDateField) {
    replacementDateField.onchange = () => {
      syncReplacementCostField();
      updateEffectRequiredHighlights(form);
    };
  }
  if (form.elements.referenceEffet) {
    form.elements.referenceEffet.onchange = () => {
      syncReplacementCostField();
      focusNextEffectKeyField(form, "referenceEffet");
      updateEffectRequiredHighlights(form);
    };
  }
  if (form.elements.designationLibre) {
    let designationLibreInputDebounceId = 0;
    const scheduleDesignationLibreUpdate = () => {
      if (designationLibreInputDebounceId) {
        window.clearTimeout(designationLibreInputDebounceId);
      }
      designationLibreInputDebounceId = window.setTimeout(() => {
        designationLibreInputDebounceId = 0;
        syncReplacementCostField();
        updateEffectRequiredHighlights(form);
      }, FILTER_INPUT_DEBOUNCE_MS);
    };

    form.elements.designationLibre.oninput = () => {
      scheduleDesignationLibreUpdate();
    };
  }

  ["numeroIdentification", "vehiculeImmatriculation", "dateRemise"].forEach((fieldName) => {
    const field = form.elements[fieldName];
    if (!(field instanceof HTMLElement)) {
      return;
    }
    field.addEventListener("keydown", (event) => {
      if (event.key !== "Enter") {
        return;
      }
      event.preventDefault();
      focusNextEffectKeyField(form, fieldName);
    });
  });

  const submitEffect = async (mode) => {
    const person = getCurrentPerson();
    if (!person) {
      showDataStatus("SELECTIONNER UNE PERSONNE AVANT D'AJOUTER UN EFFET");
      return;
    }

    if (mode === "edit" && !state.editingEffectId) {
      showDataStatus("SELECTIONNER D'ABORD UN EFFET A MODIFIER");
      return;
    }

    const validation = validateEffectFormContext(form, { markMissing: true });
    updateEffectActionButtons();
    if (!validation.ok) {
      showDataStatus(validation.message);
      window.alert(validation.message);
      focusEffectField(form, validation.field);
      return;
    }

    const formData = new FormData(form);
    const typeEffet = normalizeText(formData.get("typeEffet"));
    const isTurboSelf = typeEffet === "CARTE TURBOSELF";
    const referenceSite = normalizeText(formData.get("referenceSite"));
    const usesReferenceCatalog = typeUsesReferenceCatalog(typeEffet);
    const usesSiteField = typeUsesSiteField(typeEffet);
    const referenceEffetId = usesReferenceCatalog ? String(formData.get("referenceEffet") || "") : "";
    const reference = findReferenceById(referenceEffetId);
    const designationLibre = usesReferenceCatalog ? normalizeText(formData.get("designationLibre")) : "";
    const availableReferenceSites = getAvailableReferenceSites(person);
    const resolvedReferenceSite = isTurboSelf
      ? ALL_SITES_VALUE
      : usesReferenceCatalog
      ? normalizeText(
          reference?.site || referenceSite || (availableReferenceSites.length === 1 ? availableReferenceSites[0] : "")
        )
      : usesSiteField
        ? normalizeText(referenceSite || (availableReferenceSites.length === 1 ? availableReferenceSites[0] : ""))
        : "";
    const dateRemplacement = String(formData.get("dateRemplacement") || "");
    const dateRetour = String(formData.get("dateRetour") || "");
    const coutRemplacement = normalizeAmount(formData.get("coutRemplacement"));
    const manualStatus = normalizeText(formData.get("statutManuel"));
    const storedManualStatus = getStoredManualStatusForEffect(manualStatus, dateRetour);

    if (!typeEffet) {
      showDataStatus("SELECTIONNER UN TYPE D'EFFET");
      form.elements.typeEffet?.focus();
      return;
    }

    if (!resolvedReferenceSite) {
      showDataStatus("SELECTIONNER LE SITE DE L'EFFET");
      form.elements.referenceSite?.focus();
      return;
    }

    if (!manualStatus && !dateRetour) {
      showDataStatus("SELECTIONNER LE STATUT MANUEL");
      form.elements.statutManuel?.focus();
      return;
    }

    if (usesReferenceCatalog && !referenceEffetId) {
      showDataStatus(typeEffet === "VENTILATEUR" ? "CHOISIR UN TYPE DE VENTILATEUR DANS LA LISTE" : "CHOISIR UNE CLE EXISTANTE DANS LA LISTE");
      return;
    }

    const effectId = mode === "edit" ? state.editingEffectId : getNextId("E", person.effetsConfies || []);
    const vehiculeImmatriculation =
      typeEffet === "TELECOMMANDE URMET" ? normalizeText(formData.get("vehiculeImmatriculation")) : "";

    const effect = {
        id: effectId,
        typeEffet,
        siteReference: resolvedReferenceSite,
        referenceEffetId,
        designation: usesReferenceCatalog ? reference?.designation || "" : "",
        numeroIdentification: normalizeText(formData.get("numeroIdentification")),
        vehiculeImmatriculation,
      dateRemise: String(formData.get("dateRemise") || ""),
      dateRetour,
      statutManuel: storedManualStatus,
      cause: "",
      dateRemplacement,
      coutRemplacement,
      commentaire: normalizeText(formData.get("commentaire")),
    };
    const normalizedEffectId = String(effectId || "");
    const existingEffect =
      mode === "edit"
        ? (person.effetsConfies || []).find((entry) => String(entry?.id || "") === normalizedEffectId)
        : null;
    const previousStatus = normalizeText(getEffectStatus(person, existingEffect || {}));
    const nextStatus = normalizeText(getEffectStatus(person, effect));
    const nextCause = getCauseFromManualStatus(manualStatus);
    const preservedCause = normalizeEffectCause(existingEffect?.cause || existingEffect?.causeRemplacement);
    effect.cause = preservedCause || nextCause;
    if (usesReferenceCatalog && !effect.designation) {
      effect.designation = `EFFET ${effect.id}`;
    }

    if (!Array.isArray(person.effetsConfies)) {
      person.effetsConfies = [];
    }
    pushUndoSnapshot(mode === "edit" ? "MODIFICATION EFFET" : "AJOUT EFFET");
    const existingIndex =
      mode === "edit"
        ? person.effetsConfies.findIndex((entry) => String(entry?.id || "") === String(effect.id || ""))
        : -1;
    if (mode === "edit" && existingIndex >= 0) {
      person.effetsConfies[existingIndex] = effect;
    } else {
      person.effetsConfies.push(effect);
      addAutoStockMovement(person, effect, "SORTIE", "AFFECTATION");
    }
    if (mode === "edit" && ["HS", "DETRUIT", "PERDU", "VOL"].includes(nextStatus) && nextStatus !== previousStatus) {
      addAutoStockMovement(person, effect, "INFO", nextStatus);
    }
    markEffectRowFlash(mode === "edit" ? "update" : "create", person.id, effect.id);

    markDirty();
    form.reset();
    state.editingEffectId = "";
    hydrateEffectReferenceSiteSelect(person, "", "");
    hydrateReferenceSelect(person, "", "", "");
    updateEffectFormMode("");
    schedulePageRender();
    const effectLabel = effect.designation || effect.numeroIdentification || effect.id;
    showActionStatus(
      mode === "edit" ? "update" : "create",
      mode === "edit"
        ? `EFFET MODIFIE : ${effectLabel}`
        : `EFFET AJOUTE : ${effectLabel}`
    );

    await saveAfterEffectChangeWithAvenantAlert(person.id);
  };

  form.onsubmit = async (event) => {
    event.preventDefault();
    await submitEffect(state.editingEffectId ? "edit" : "add");
  };

  if (addButton) {
    addButton.onclick = async () => {
      await submitEffect("add");
    };
  }
  if (updateButton) {
    updateButton.onclick = async () => {
      await submitEffect("edit");
    };
  }
  if (deleteButton) {
    deleteButton.onclick = async () => {
      const person = getCurrentPerson();
      if (!person || !state.editingEffectId) {
        showDataStatus("SELECTIONNER D'ABORD UN EFFET A SUPPRIMER");
        return;
      }
      await deleteEffect(person.id, state.editingEffectId);
    };
  }
  if (cancelButton) {
    cancelButton.onclick = () => {
      state.editingEffectId = "";
      resetEffectForm();
      showDataStatus("MODIFICATION DE L'EFFET ANNULEE");
    };
  }
  if (resetFieldsButton) {
    resetFieldsButton.onclick = () => {
      resetEffectForm();
      showDataStatus("FORMULAIRE EFFET REINITIALISE");
    };
  }

  form.addEventListener("input", () => {
    updateEffectResetButtonState(form);
    updateEffectRequiredHighlights(form);
  });
  form.addEventListener("change", () => {
    updateEffectResetButtonState(form);
    updateEffectRequiredHighlights(form);
  });

  updateEffectActionButtons();
  updateEffectRequiredHighlights(form);
  updateEffectResetButtonState(form);
  updateManualStatusCriticalState(form);
}

async function deleteEffect(personId, effectId) {
  const person = state.data?.personnes?.find((entry) => entry.id === personId);
  if (!person || !Array.isArray(person.effetsConfies)) {
    return;
  }

  const effect = person.effetsConfies.find((entry) => entry.id === effectId);
  const confirmDelete = window.confirm(
    `ARCHIVER CET EFFET${effect?.designation ? ` : ${effect.designation}` : effect?.numeroIdentification ? ` : ${effect.numeroIdentification}` : ""} ?`
  );
  if (!confirmDelete) {
    return;
  }

  pushUndoSnapshot("SUPPRESSION EFFET");
  if (effect) {
    addAutoStockMovement(person, effect, "ENTREE", "SUPPRESSION_DOTATION");
  }
  const effectIndex = person.effetsConfies.findIndex((entry) => entry.id === effectId);
  if (effectIndex >= 0) {
    person.effetsConfies[effectIndex] = markSoftDeletedEntity(person.effetsConfies[effectIndex]);
  }
  markEffectTableFlash("delete", personId);
  if (state.editingEffectId === effectId) {
    state.editingEffectId = "";
    resetEffectForm();
  }
  markDirty();
  schedulePageRender();
  showActionStatus("delete", `EFFET ARCHIVE : ${effectId}`);
  await saveAfterEffectChangeWithAvenantAlert(person.id);
}

function hasStoredSignaturePayload(entry) {
  return Boolean(
    String(entry?.image || "").trim() ||
    String(entry?.validatedAt || "").trim() ||
    String(entry?.storageRef || "").trim() ||
    String(entry?.storagePublicUrl || "").trim()
  );
}

function cloneSignatureEntry(entry) {
  return {
    image: String(entry?.image || ""),
    validatedAt: String(entry?.validatedAt || ""),
    storageRef: String(entry?.storageRef || ""),
    storagePublicUrl: String(entry?.storagePublicUrl || ""),
  };
}

async function saveAfterEffectChangeWithAvenantAlert(personId = "") {
  const targetPersonId = String(personId || "");
  const beforePerson = targetPersonId
    ? (state.data?.personnes || []).find((entry) => String(entry?.id || "") === targetPersonId) || null
    : null;
  const beforeArrivalPersonnel = cloneSignatureEntry(beforePerson?.signatures?.arrival?.personnel);
  const beforeArrivalRepresentant = cloneSignatureEntry(beforePerson?.signatures?.arrival?.representant);

  await saveDataToFile({
    silent: true,
    reloadAfter: true,
  });
  if (state.isDirty) {
    showDataStatus("SAUVEGARDE IMPOSSIBLE - ALERTE ANNULEE");
    return;
  }

  // Safety net: effect updates must never erase already stored ARRIVEE signatures.
  if (targetPersonId && (hasStoredSignaturePayload(beforeArrivalPersonnel) || hasStoredSignaturePayload(beforeArrivalRepresentant))) {
    const currentPerson =
      (state.data?.personnes || []).find((entry) => String(entry?.id || "") === targetPersonId) || null;
    if (currentPerson) {
      const currentArrivalPersonnel = currentPerson?.signatures?.arrival?.personnel;
      const currentArrivalRepresentant = currentPerson?.signatures?.arrival?.representant;
      const mustRestorePersonnel =
        hasStoredSignaturePayload(beforeArrivalPersonnel) && !hasStoredSignaturePayload(currentArrivalPersonnel);
      const mustRestoreRepresentant =
        hasStoredSignaturePayload(beforeArrivalRepresentant) && !hasStoredSignaturePayload(currentArrivalRepresentant);
      if (mustRestorePersonnel || mustRestoreRepresentant) {
        if (!currentPerson.signatures || typeof currentPerson.signatures !== "object") {
          currentPerson.signatures = {};
        }
        if (!currentPerson.signatures.arrival || typeof currentPerson.signatures.arrival !== "object") {
          currentPerson.signatures.arrival = {};
        }
        if (mustRestorePersonnel) {
          currentPerson.signatures.arrival.personnel = { ...beforeArrivalPersonnel };
        }
        if (mustRestoreRepresentant) {
          currentPerson.signatures.arrival.representant = { ...beforeArrivalRepresentant };
        }
        markDirty();
        await saveDataToFile({
          silent: true,
          reloadAfter: false,
          successText: "SIGNATURES ARRIVEE RESTAUREES",
        });
      }
    }
  }

  window.alert(
    "DES MODIFICATIONS D'EFFETS ONT ETE EFFECTUEES. VOUS DEVEZ DONC PROCEDER A UNE NOUVELLE SIGNATURE DE L'AVENANT."
  );
}

async function deletePerson(personId) {
  if (!state.data?.personnes) {
    return;
  }

  const person = state.data.personnes.find((entry) => entry.id === personId);
  if (!person) {
    return;
  }

  const confirmDelete = window.confirm(
    `ARCHIVER ${person.nom} ${person.prenom} ?`
  );
  if (!confirmDelete) {
    return;
  }

  pushUndoSnapshot("SUPPRESSION PERSONNE");
  const personIndex = state.data.personnes.findIndex((entry) => entry.id === personId);
  if (personIndex >= 0) {
    state.data.personnes[personIndex] = markSoftDeletedEntity(state.data.personnes[personIndex]);
  }
  if (getCurrentPersonId() === personId) {
    setCurrentPersonId("");
  }
  state.editingEffectId = "";
  markDirty();
  schedulePageRender();
  showActionStatus("delete", `PERSONNE ARCHIVEE : ${person.nom} ${person.prenom}`);
  await saveDataToFile({
    silent: true,
    reloadAfter: true,
    successText: "PERSONNE ARCHIVEE - SAUVEGARDE AUTOMATIQUE",
  });

  if (document.body.dataset.page === "person-sheet") {
    navigateWithAutoSave("fiche-personne.html");
  }
}

function startEditEffect(personId, effectId) {
  const normalizedPersonId = String(personId || "");
  const normalizedEffectId = String(effectId || "");
  const person = state.data?.personnes?.find(
    (entry) => String(entry.id || "") === normalizedPersonId
  );
  const effect = person?.effetsConfies?.find(
    (entry) => String(entry.id || "") === normalizedEffectId
  );
  const form = document.getElementById("effect-form");
  if (!person || !effect || !form) {
    return;
  }

  state.editingEffectId = normalizedEffectId;
  hydrateEffectReferenceSiteSelect(person, effect.siteReference || referenceSiteFromEffect(effect), effect.typeEffet);
  hydrateReferenceSelect(
    person,
    effect.typeEffet,
    effect.referenceEffetId,
    effect.siteReference || referenceSiteFromEffect(effect)
  );
  const usesReferenceCatalog = typeUsesReferenceCatalog(effect.typeEffet);
  const reference = findReferenceById(effect.referenceEffetId);
  const editDesignation = usesReferenceCatalog
    ? effect.designation || reference?.designation || ""
    : "";
  form.elements.typeEffet.value = effect.typeEffet || "";
  form.elements.referenceSite.value = effect.siteReference || referenceSiteFromEffect(effect) || "";
  form.elements.referenceEffet.value = effect.referenceEffetId || "";
  form.elements.designationLibre.value = editDesignation;
  form.elements.numeroIdentification.value = effect.numeroIdentification || "";
  form.elements.vehiculeImmatriculation.value = effect.vehiculeImmatriculation || "";
  form.elements.dateRemise.value = effect.dateRemise || "";
  form.elements.dateRetour.value = effect.dateRetour || "";
  form.elements.statutManuel.value = effect.statutManuel || "";
  form.elements.dateRemplacement.value = effect.dateRemplacement || "";
  form.elements.coutRemplacement.value = formatAmountWithEuro(effect.coutRemplacement);
  form.elements.commentaire.value = effect.commentaire || "";
  updateEffectFormMode(effect.typeEffet || "");
  const selectedEffectReferenceSite = effect.siteReference || referenceSiteFromEffect(effect) || getDefaultEffectSiteReference(person, effect) || "";
  hydrateEffectReferenceSiteSelect(person, selectedEffectReferenceSite, effect.typeEffet);
  if (form.elements.referenceSite) {
    form.elements.referenceSite.value = selectedEffectReferenceSite;
  }
  hydrateReferenceSelect(person, effect.typeEffet, effect.referenceEffetId, selectedEffectReferenceSite);
  if (form.elements.referenceEffet) {
    form.elements.referenceEffet.value = effect.referenceEffetId || "";
  }
  syncReplacementCostField();
  updateEffectActionButtons();
  updateEffectResetButtonState(form);
  updateEffectRequiredHighlights(form);
  updateManualStatusCriticalState(form);
  form.scrollIntoView({ behavior: "smooth", block: "center" });
  if (usesReferenceCatalog) {
    form.elements.referenceEffet.focus();
  } else {
    form.elements.numeroIdentification.focus();
    form.elements.numeroIdentification.select();
  }

  showDataStatus(
    `EFFET EN COURS DE MODIFICATION : ${editDesignation || effect.numeroIdentification || effect.id}`
  );
}

function resetEffectForm() {
  const form = document.getElementById("effect-form");
  if (!form) {
    return;
  }
  state.editingEffectId = "";
  form.reset();
  hydrateEffectReferenceSiteSelect(getCurrentPerson(), "", "");
  hydrateReferenceSelect(getCurrentPerson() || "", "", "", "");
  updateEffectFormMode("");
  updateEffectActionButtons();
  updateEffectRequiredHighlights(form);
  updateEffectResetButtonState(form);
  updateManualStatusCriticalState(form);
}

function bindArchiveFilterDelegation() {
  if (window.__documentsArchiveFilterDelegationBound) {
    return;
  }
  const handleArchiveFilterChange = (event) => {
    if (document.body?.dataset?.page !== "documents-archives") {
      return;
    }
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const form = target.closest("#documents-archives-filter-form");
    if (!(form instanceof HTMLFormElement)) {
      return;
    }
    state.listRenderCache.documentsArchives = "";
    schedulePageRender();
  };
  document.addEventListener("input", handleArchiveFilterChange, true);
  document.addEventListener("change", handleArchiveFilterChange, true);
  window.__documentsArchiveFilterDelegationBound = true;
}

function getArchiveFilterSignatureFromDom() {
  if (document.body?.dataset?.page !== "documents-archives") {
    return "";
  }
  const form = document.getElementById("documents-archives-filter-form");
  if (!(form instanceof HTMLFormElement)) {
    return "";
  }
  return ["archiveSearch", "archiveTypeDocument", "archiveSite", "archiveStatutSignature"]
    .map((fieldName) => {
      const field = form.querySelector(`[name="${fieldName}"]`);
      return field instanceof HTMLInputElement || field instanceof HTMLSelectElement
        ? String(field.value || "")
        : "";
    })
    .join("|");
}

function bindArchiveFilterValueWatcher() {
  if (window.__documentsArchiveFilterWatcherBound) {
    return;
  }
  const checkArchiveFilterValues = () => {
    if (document.body?.dataset?.page !== "documents-archives") {
      return;
    }
    const signature = getArchiveFilterSignatureFromDom();
    if (!signature) {
      return;
    }
    if (state.documentsArchiveFilterSignature === undefined) {
      state.documentsArchiveFilterSignature = signature;
      return;
    }
    if (state.documentsArchiveFilterSignature === signature) {
      return;
    }
    state.documentsArchiveFilterSignature = signature;
    state.listRenderCache.documentsArchives = "";
    schedulePageRender();
  };
  document.addEventListener("input", checkArchiveFilterValues, true);
  document.addEventListener("change", checkArchiveFilterValues, true);
  window.__documentsArchiveFilterWatcherBound = true;
}

function resetEffectFormFieldsExceptCost() {
  const form = document.getElementById("effect-form");
  if (!form) {
    return;
  }
  const preservedCost = String(form.elements.coutRemplacement?.value || "");
  const person = getCurrentPerson();
  state.editingEffectId = "";
  form.reset();
  hydrateEffectReferenceSiteSelect(person, "", "");
  hydrateReferenceSelect(person || "", "", "", "");
  updateEffectFormMode("");
  updateEffectActionButtons();
  if (form.elements.coutRemplacement) {
    form.elements.coutRemplacement.value = preservedCost;
  }
  updateEffectRequiredHighlights(form);
  updateEffectResetButtonState(form);
  updateManualStatusCriticalState(form);
  showDataStatus("CHAMPS REINITIALISES (COUT CONSERVE)");
}

function isCriticalManualStatus(value) {
  return ["PERDU", "VOL", "HS", "DETRUIT", "CASSE", "NON RENDU"].includes(normalizeText(value));
}

function isReturnDateRequiredForManualStatus(value) {
  return ["RENDU", "RESTITUE"].includes(normalizeText(value));
}

function getStoredManualStatusForEffect(manualStatus, dateRetour) {
  if (String(dateRetour || "").trim()) {
    return "RESTITUE";
  }
  const normalizedStatus = normalizeText(manualStatus);
  if (normalizedStatus === "CASSE") {
    return "DETRUIT";
  }
  if (["RENDU", "RESTITUE"].includes(normalizedStatus)) {
    return "ACTIF";
  }
  return normalizedStatus;
}

function isReplacementDateRequiredForManualStatus(value) {
  return ["PERDU", "VOL", "HS", "DETRUIT", "CASSE", "NON RENDU"].includes(normalizeText(value));
}

function updateManualStatusCriticalState(form) {
  const fieldNode = getEffectFormFieldNode(form, "statutManuel");
  const statusValue = form?.elements?.statutManuel?.value || "";
  if (!fieldNode) {
    return;
  }
  fieldNode.classList.toggle("field--status-critical", isCriticalManualStatus(statusValue));
}

function markEffectRowFlash(kind, personId, effectId) {
  state.effectRowFlash = {
    kind: String(kind || ""),
    personId: String(personId || ""),
    effectId: String(effectId || ""),
  };
}

function markEffectTableFlash(kind, personId) {
  state.effectTableFlash = {
    kind: String(kind || ""),
    personId: String(personId || ""),
  };
}

function hasEffectFormUserContent(form) {
  if (!form) {
    return false;
  }
  const trackedFields = [
    "typeEffet",
    "referenceSite",
    "referenceEffet",
    "designationLibre",
    "numeroIdentification",
    "vehiculeImmatriculation",
    "dateRemise",
    "dateRetour",
    "statutManuel",
    "dateRemplacement",
    "commentaire",
  ];
  return trackedFields.some((fieldName) => {
    const value = form.elements[fieldName]?.value;
    return Boolean(String(value || "").trim());
  });
}

function updateEffectResetButtonState(form) {
  const resetButton = document.getElementById("effect-reset-fields-button");
  if (!resetButton) {
    return;
  }
  resetButton.classList.toggle("is-ready", hasEffectFormUserContent(form));
}

function updateEffectActionButtons() {
  const form = document.getElementById("effect-form");
  const addButton = document.getElementById("effect-add-button");
  const updateButton = document.getElementById("effect-update-button");
  const deleteButton = document.getElementById("effect-delete-button");
  const cancelButton = document.getElementById("effect-cancel-button");
  const isEditing = Boolean(state.editingEffectId);
  const validation = validateEffectFormContext(form, { markMissing: false });
  const canSubmit = Boolean(validation.ok);

  if (addButton) {
    addButton.disabled = isEditing;
    addButton.classList.toggle("is-disabled", !isEditing && !canSubmit);
    addButton.setAttribute("aria-disabled", !isEditing && !canSubmit ? "true" : "false");
    addButton.classList.toggle("button--primary", !isEditing);
    addButton.classList.toggle("button--secondary", isEditing);
    addButton.title = canSubmit ? "" : validation.message;
  }
  if (updateButton) {
    updateButton.disabled = !isEditing;
    updateButton.classList.toggle("is-disabled", isEditing && !canSubmit);
    updateButton.setAttribute("aria-disabled", isEditing && !canSubmit ? "true" : "false");
    updateButton.classList.toggle("button--primary", isEditing);
    updateButton.classList.toggle("button--secondary", !isEditing);
    updateButton.title = canSubmit ? "" : validation.message;
  }
  if (deleteButton) {
    deleteButton.disabled = !isEditing;
  }
  if (cancelButton) {
    cancelButton.disabled = !isEditing;
  }
}

function getEffectKeyFieldSequence(typeEffet) {
  const normalizedType = normalizeText(typeEffet);
  if (normalizedType === "TELECOMMANDE URMET") {
    return ["typeEffet", "referenceSite", "numeroIdentification", "vehiculeImmatriculation", "dateRemise", "statutManuel"];
  }
  if (normalizedType === "BADGE INTRUSION" || normalizedType === "CARTE TURBOSELF") {
    return ["typeEffet", "referenceSite", "numeroIdentification", "dateRemise", "statutManuel"];
  }
  if (["CLE", "CLE CES", "CLE DE SECURITE"].includes(normalizedType)) {
    return ["typeEffet", "referenceSite", "referenceEffet", "numeroIdentification", "dateRemise", "statutManuel"];
  }
  return ["typeEffet", "numeroIdentification", "dateRemise", "statutManuel"];
}

function getEffectFormFieldNode(form, name) {
  const node = form?.elements?.[name];
  if (!(node instanceof HTMLElement)) {
    return null;
  }
  return node.closest(".field");
}

function isEffectFieldAvailable(form, name) {
  const node = form?.elements?.[name];
  if (!(node instanceof HTMLElement) || node.disabled) {
    return false;
  }
  return true;
}

function focusEffectField(form, name) {
  const node = form?.elements?.[name];
  if (!(node instanceof HTMLElement) || !isEffectFieldAvailable(form, name)) {
    return false;
  }
  node.focus();
  if (node instanceof HTMLInputElement && node.type === "text") {
    node.select();
  }
  return true;
}

function focusNextEffectKeyField(form, currentFieldName) {
  if (!form) {
    return;
  }
  const sequence = getEffectKeyFieldSequence(form.elements.typeEffet?.value || "");
  const currentIndex = sequence.indexOf(currentFieldName);
  if (currentIndex < 0) {
    return;
  }
  for (let i = currentIndex + 1; i < sequence.length; i += 1) {
    if (focusEffectField(form, sequence[i])) {
      return;
    }
  }
}

function setEffectFieldVisualState(form, name, enabled, isKey) {
  const fieldNode = getEffectFormFieldNode(form, name);
  const control = form?.elements?.[name];
  if (fieldNode) {
    fieldNode.classList.toggle("field--inactive", !enabled);
    fieldNode.classList.toggle("field--key", Boolean(isKey));
  }
  if (control instanceof HTMLElement) {
    control.disabled = !enabled;
  }
}

function setEffectFieldMissingState(form, name, isMissing) {
  const fieldNode = getEffectFormFieldNode(form, name);
  if (fieldNode) {
    fieldNode.classList.toggle("field--missing", Boolean(isMissing));
  }
}

function getEffectFormContext(form) {
  if (!form) {
    return null;
  }
  const typeEffet = normalizeText(form.elements.typeEffet?.value || "");
  const isTurboSelf = typeEffet === "CARTE TURBOSELF";
  const person = getCurrentPerson();
  const availableReferenceSites = getAvailableReferenceSites(person);
  const referenceSite = normalizeText(form.elements.referenceSite?.value || "");
  const usesReferenceCatalog = typeUsesReferenceCatalog(typeEffet);
  const usesSiteField = typeUsesSiteField(typeEffet);
  const referenceEffetId = usesReferenceCatalog ? String(form.elements.referenceEffet?.value || "") : "";
  const reference = findReferenceById(referenceEffetId);
  const resolvedReferenceSite = isTurboSelf
    ? ALL_SITES_VALUE
    : usesReferenceCatalog
      ? normalizeText(
          reference?.site || referenceSite || (availableReferenceSites.length === 1 ? availableReferenceSites[0] : "")
        )
      : usesSiteField
        ? normalizeText(referenceSite || (availableReferenceSites.length === 1 ? availableReferenceSites[0] : ""))
        : "";
  const statutManuel = normalizeText(form.elements.statutManuel?.value || "");
  const dateRemise = String(form.elements.dateRemise?.value || "").trim();
  const dateRetour = String(form.elements.dateRetour?.value || "").trim();
  const dateRemplacement = String(form.elements.dateRemplacement?.value || "").trim();
  const returnDateRequired = isReturnDateRequiredForManualStatus(statutManuel);
  const replacementDateRequired = isReplacementDateRequiredForManualStatus(statutManuel);
  return {
    typeEffet,
    referenceSite,
    usesReferenceCatalog,
    usesSiteField,
    referenceEffetId,
    resolvedReferenceSite,
    statutManuel,
    dateRemise,
    dateRetour,
    dateRemplacement,
    returnDateRequired,
    replacementDateRequired,
    siteRequired: Boolean(form.elements.referenceSite?.required),
  };
}

function getEffectReferenceMissingMessage(typeEffet) {
  return normalizeText(typeEffet) === "VENTILATEUR"
    ? "VOUS DEVEZ REMPLIR LE CHAMP TYPE DE VENTILATEUR"
    : "VOUS DEVEZ REMPLIR LE CHAMP DESIGNATION EXISTANTE";
}

function validateEffectFormContext(form, options = {}) {
  const markMissing = options.markMissing !== false;
  const context = getEffectFormContext(form);
  if (!context) {
    return { ok: false, field: "", message: "FORMULAIRE EFFET INDISPONIBLE" };
  }
  const missing = {
    typeEffet: !context.typeEffet,
    referenceSite: Boolean(context.siteRequired && !context.referenceSite),
    referenceEffet: Boolean(context.usesReferenceCatalog && !context.referenceEffetId),
    dateRemise: !context.dateRemise,
    dateRetour: Boolean(context.returnDateRequired && !context.dateRetour),
    statutManuel: !context.statutManuel && !context.dateRetour,
    dateRemplacement: Boolean(context.replacementDateRequired && !context.dateRemplacement),
  };
  const dateValidations = {
    dateRemise: validateDateFieldFormat(context.dateRemise, "DATE DE REMISE"),
    dateRetour: validateDateFieldFormat(context.dateRetour, "DATE DE RETOUR"),
    dateRemplacement: validateDateFieldFormat(context.dateRemplacement, "DATE DE REMPLACEMENT"),
  };

  if (form.elements.dateRemise instanceof HTMLElement) {
    form.elements.dateRemise.required = true;
  }
  if (form.elements.dateRetour instanceof HTMLElement) {
    form.elements.dateRetour.required = context.returnDateRequired;
  }
  if (form.elements.dateRemplacement instanceof HTMLElement) {
    form.elements.dateRemplacement.required = context.replacementDateRequired;
  }

  if (markMissing) {
    const returnDateField = getEffectFormFieldNode(form, "dateRetour");
    const replacementDateField = getEffectFormFieldNode(form, "dateRemplacement");
    if (returnDateField) {
      returnDateField.classList.toggle("field--key", Boolean(context.returnDateRequired));
    }
    if (replacementDateField) {
      replacementDateField.classList.toggle("field--key", Boolean(context.replacementDateRequired));
    }
    setEffectFieldMissingState(form, "typeEffet", missing.typeEffet);
    setEffectFieldMissingState(form, "referenceSite", missing.referenceSite);
    setEffectFieldMissingState(form, "referenceEffet", missing.referenceEffet);
    setEffectFieldMissingState(form, "dateRemise", missing.dateRemise || !dateValidations.dateRemise.ok);
    setEffectFieldMissingState(form, "dateRetour", missing.dateRetour || !dateValidations.dateRetour.ok);
    setEffectFieldMissingState(form, "statutManuel", missing.statutManuel);
    setEffectFieldMissingState(form, "dateRemplacement", missing.dateRemplacement || !dateValidations.dateRemplacement.ok);
  }

  if (missing.typeEffet) {
    return { ok: false, field: "typeEffet", message: "VOUS DEVEZ REMPLIR LE CHAMP TYPE D'EFFET" };
  }
  if (missing.referenceSite) {
    return { ok: false, field: "referenceSite", message: "VOUS DEVEZ REMPLIR LE CHAMP SITE DE L'EFFET" };
  }
  if (missing.referenceEffet) {
    return { ok: false, field: "referenceEffet", message: getEffectReferenceMissingMessage(context.typeEffet) };
  }
  if (missing.dateRemise) {
    return { ok: false, field: "dateRemise", message: "VOUS DEVEZ REMPLIR LE CHAMP DATE DE REMISE" };
  }
  if (!dateValidations.dateRemise.ok) {
    return { ok: false, field: "dateRemise", message: dateValidations.dateRemise.message };
  }
  if (missing.dateRetour) {
    return { ok: false, field: "dateRetour", message: "VOUS DEVEZ REMPLIR LE CHAMP DATE DE RETOUR" };
  }
  if (!dateValidations.dateRetour.ok) {
    return { ok: false, field: "dateRetour", message: dateValidations.dateRetour.message };
  }
  if (missing.statutManuel) {
    return { ok: false, field: "statutManuel", message: "VOUS DEVEZ REMPLIR LE CHAMP STATUT MANUEL" };
  }
  if (missing.dateRemplacement) {
    return {
      ok: false,
      field: "dateRemplacement",
      message: "VOUS DEVEZ REMPLIR LE CHAMP DATE DE REMPLACEMENT",
    };
  }
  if (!dateValidations.dateRemplacement.ok) {
    return { ok: false, field: "dateRemplacement", message: dateValidations.dateRemplacement.message };
  }
  return { ok: true, field: "", message: "" };
}

function updateEffectRequiredHighlights(form) {
  validateEffectFormContext(form, { markMissing: true });
  updateEffectActionButtons();
}

function updateEffectFormMode(typeEffet) {
  const normalizedType = normalizeText(typeEffet);
  const person = getCurrentPerson();
  const form = document.getElementById("effect-form");
  const availableReferenceSites = getAvailableReferenceSites(person);
  const referenceSiteField = document.getElementById("effect-reference-site-field");
  const referenceSiteLabel = document.getElementById("effect-reference-site-label");
  const referenceField = document.getElementById("effect-reference-field");
  const referenceLabel = document.getElementById("effect-reference-label");
  const designationField = document.getElementById("effect-designation-field");
  const designationLabel = document.getElementById("effect-designation-label");
  const numberLabel = document.getElementById("effect-number-label");
  const vehicleField = document.getElementById("effect-vehicle-field");
  const vehicleLabel = document.getElementById("effect-vehicle-label");
  const helpNode = document.getElementById("effect-form-help");

  if (
    !form ||
    !referenceSiteField ||
    !referenceSiteLabel ||
    !referenceField ||
    !referenceLabel ||
    !designationField ||
    !designationLabel ||
    !numberLabel ||
    !vehicleField ||
    !vehicleLabel ||
    !helpNode
  ) {
    return;
  }

  const vehicleInput = vehicleField.querySelector("input");
  if (vehicleInput && normalizedType !== "TELECOMMANDE URMET") {
    vehicleInput.value = "";
  }
  numberLabel.textContent = "N° D'IDENTIFICATION";

  let showReferenceSite = Boolean(normalizedType);
  let showReference = true;
  let showDesignation = true;
  let showVehicle = false;
  let keyFields = normalizedType ? getEffectKeyFieldSequence(normalizedType) : ["typeEffet"];

  if (["CLE", "CLE CES", "CLE DE SECURITE"].includes(normalizedType)) {
    showReferenceSite = true;
    referenceSiteLabel.textContent = "SITE DE LA CLE";
    referenceLabel.textContent = "NOM EXISTANT DE LA CLE";
    designationLabel.textContent = "NOUVEAU NOM / MODIFICATION";
    showDesignation = false;
    numberLabel.textContent = "N° DE LA CLE";
    if (normalizedType === "CLE CES") {
      helpNode.textContent =
        availableReferenceSites.length > 1
          ? "POUR UNE CLE CES : CHOISIR D'ABORD LE SITE, PUIS UNE CLE COMMENCANT PAR CES-"
          : "POUR UNE CLE CES : CHOISIR UNE CLE COMMENCANT PAR CES-";
    } else if (normalizedType === "CLE DE SECURITE") {
      helpNode.textContent =
        availableReferenceSites.length > 1
          ? "POUR UNE CLE DE SECURITE : CHOISIR D'ABORD LE SITE, PUIS LE NOM DE LA CLE"
          : "POUR UNE CLE DE SECURITE : CHOISIR UN NOM DE CLE DU SITE";
    } else {
      helpNode.textContent =
        availableReferenceSites.length > 1
          ? "POUR UNE CLE : CHOISIR D'ABORD LE SITE, PUIS LE NOM DE LA CLE"
          : "POUR UNE CLE : CHOISIR UN NOM DE CLE DU SITE";
    }
  } else if (normalizedType === "VENTILATEUR") {
    showReferenceSite = true;
    referenceSiteLabel.textContent = "SITE DU VENTILATEUR";
    referenceLabel.textContent = "TYPE DE VENTILATEUR";
    designationLabel.textContent = "DESIGNATION";
    showDesignation = false;
    numberLabel.textContent = "N° VENTILATEUR";
    helpNode.textContent =
      availableReferenceSites.length > 1
        ? "POUR UN VENTILATEUR : CHOISIR D'ABORD LE SITE, PUIS LE TYPE"
        : "POUR UN VENTILATEUR : CHOISIR LE TYPE";
  } else if (["BADGE INTRUSION", "RADIATEUR APPOINT", "TELECOMMANDE URMET", "CARTE TURBOSELF"].includes(normalizedType)) {
    showReferenceSite = true;
    referenceSiteLabel.textContent =
      normalizedType === "BADGE INTRUSION"
        ? "SITE DU BADGE"
        : normalizedType === "RADIATEUR APPOINT"
          ? "SITE DU RADIATEUR"
        : normalizedType === "CARTE TURBOSELF"
          ? "SITE DE LA CARTE"
          : "SITE DE LA TELECOMMANDE";
    showReference = false;
    showDesignation = false;
    showVehicle = normalizedType === "TELECOMMANDE URMET";
    vehicleLabel.textContent = "VEHICULE / IMMATRICULATION";
    referenceLabel.textContent = "REFERENCE EXISTANTE";
    designationLabel.textContent = "DESIGNATION";
    if (normalizedType === "BADGE INTRUSION") {
      numberLabel.textContent = "N° BADGE";
      helpNode.textContent = "POUR UN BADGE INTRUSION : RENSEIGNER UNIQUEMENT LE N°";
    } else if (normalizedType === "RADIATEUR APPOINT") {
      numberLabel.textContent = "N° RADIATEUR";
      helpNode.textContent = "POUR UN RADIATEUR APPOINT : CHOISIR LE SITE, PAS DE DESIGNATION";
    } else if (normalizedType === "TELECOMMANDE URMET") {
      numberLabel.textContent = "N° TELECOMMANDE";
      helpNode.textContent = "POUR UNE TELECOMMANDE URMET : RENSEIGNER UNIQUEMENT LE N°";
    } else if (normalizedType === "CARTE TURBOSELF") {
      numberLabel.textContent = "N° CARTE";
      helpNode.textContent = "POUR UNE CARTE TURBOSELF : RENSEIGNER UNIQUEMENT LE N°";
    }
  } else {
    showReferenceSite = Boolean(normalizedType);
    referenceSiteLabel.textContent = "SITE DE L'EFFET";
    referenceLabel.textContent = "DESIGNATION EXISTANTE";
    designationLabel.textContent = "NOUVELLE DESIGNATION / MODIFICATION";
    helpNode.textContent = normalizedType
      ? "SI BESOIN : CHOISIR UNE REFERENCE EXISTANTE OU SAISIR UNE DESIGNATION"
      : "CHOISIR UN TYPE D'EFFET POUR ADAPTER LA SAISIE";
  }

  setEffectFieldVisualState(form, "typeEffet", true, true);
  setEffectFieldVisualState(form, "referenceSite", showReferenceSite, keyFields.includes("referenceSite"));
  setEffectFieldVisualState(form, "referenceEffet", showReference, keyFields.includes("referenceEffet"));
  setEffectFieldVisualState(form, "designationLibre", showDesignation, keyFields.includes("designationLibre"));
  setEffectFieldVisualState(form, "numeroIdentification", true, keyFields.includes("numeroIdentification"));
  setEffectFieldVisualState(
    form,
    "vehiculeImmatriculation",
    showVehicle,
    keyFields.includes("vehiculeImmatriculation")
  );
  setEffectFieldVisualState(form, "dateRemise", true, keyFields.includes("dateRemise"));
  setEffectFieldVisualState(form, "statutManuel", true, true);
  form.elements.typeEffet.required = true;
  form.elements.referenceSite.required = showReferenceSite;
  form.elements.dateRemise.required = true;
  form.elements.dateRetour.required = isReturnDateRequiredForManualStatus(form.elements.statutManuel?.value || "");
  form.elements.statutManuel.required = true;
  form.elements.dateRemplacement.required = isReplacementDateRequiredForManualStatus(
    form.elements.statutManuel?.value || ""
  );
  updateEffectRequiredHighlights(form);
}

function isCesKeyDesignation(designation) {
  const normalized = normalizeText(designation);
  return normalized === "CES" || normalized.startsWith("CES-") || normalized.startsWith("CES ");
}

function bindRegisterButtonsAutoSave() {
  if (document.body?.dataset?.registerAutosaveBound === "true") {
    return;
  }

  document.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const button = target.closest("button");
    if (!(button instanceof HTMLButtonElement)) {
      return;
    }
    const label = normalizeText(button.textContent || "");
    if (!label.startsWith("ENREGISTRER")) {
      return;
    }
    if (button.id === "stock-save-movement") {
      return;
    }

    window.setTimeout(async () => {
      if (!state.isDirty) {
        return;
      }
      await saveDataToFile({
        silent: true,
        reloadAfter: false,
        promptDownload: false,
      });
    }, 220);
  });

  document.body.dataset.registerAutosaveBound = "true";
}

function showStockAdjustmentStatus(message, variant = "info") {
  const node = document.getElementById("stock-adjustment-status");
  if (node) {
    node.textContent = message;
    node.dataset.status = variant;
  }
  showDataStatus(message);
}

function setStockResetButtonPending(isPending) {
  const resetButton = document.getElementById("stock-reset-filters");
  if (!(resetButton instanceof HTMLButtonElement)) {
    return;
  }
  resetButton.classList.toggle("button--primary", Boolean(isPending));
  resetButton.classList.toggle("button--secondary", !isPending);
  resetButton.classList.toggle("stock-reset-pending", Boolean(isPending));
}

function getReplacementCostValue(typeEffet, causeRemplacement, designation = "") {
  const normalizedType = normalizePricingKey(typeEffet);
  const normalizedCause = normalizePricingKey(causeRemplacement, { cause: true });
  if (!normalizedType) {
    return 0;
  }

  if (normalizedType === "CLE CES") {
    return BILLABLE_EFFECT_CAUSES.includes(normalizedCause) ? 50 : 0;
  }

  const matchingEntry = (state.data?.listes?.coutsRemplacement || []).find(
    (entry) =>
      normalizePricingKey(entry?.typeEffet) === normalizedType &&
      normalizePricingKey(entry?.cause, { cause: true }) === normalizedCause &&
      normalizePricingKey(entry?.designation || "") === normalizePricingKey(designation || "")
  ) || (state.data?.listes?.coutsRemplacement || []).find(
    (entry) =>
      normalizePricingKey(entry?.typeEffet) === normalizedType &&
      normalizePricingKey(entry?.cause, { cause: true }) === normalizedCause &&
      !normalizePricingKey(entry?.designation || "")
  );

  if (!matchingEntry) {
    if (normalizedCause === "NON RENDU") {
      return getFallbackNonRenduCost(normalizedType, designation);
    }
    return 0;
  }

  if (!BILLABLE_EFFECT_CAUSES.includes(normalizedCause)) {
    return 0;
  }

  if (normalizedType === "CLE") {
    return isCesKeyDesignation(designation) ? 50 : 5;
  }

  if (normalizedType === "VENTILATEUR") {
    return normalizePricingKey(designation) === "VENTILATEUR SUR PIED" ? 35 : 30;
  }

  const rawAmount = String(matchingEntry.montant ?? "").trim();
  const normalizedRawAmount = rawAmount.replace(/\s/g, "").replace(",", ".");
  const hasValidNumericAmount = /^-?\d+(\.\d+)?$/.test(normalizedRawAmount);
  const parsedAmount = normalizeAmount(matchingEntry.montant);
  if (!rawAmount || !hasValidNumericAmount) {
    console.warn(
      "Tarif invalide detecte dans coutsRemplacement",
      normalizedType,
      normalizedCause,
      matchingEntry.montant
    );
    if (normalizedCause === "NON RENDU") {
      return getFallbackNonRenduCost(normalizedType, designation);
    }
  }
  return parsedAmount;
}

function getEffectUnitValue(effect) {
  const normalizedType = normalizeText(effect?.typeEffet);
  if (!normalizedType) {
    return 0;
  }

  if (normalizedType === "CLE CES") {
    return 50;
  }

  if (normalizedType === "CLE") {
    return isCesKeyDesignation(effect?.designation || "") ? 50 : 5;
  }

  if (normalizedType === "CLE DE SECURITE") {
    return 45;
  }

  if (normalizedType === "BADGE INTRUSION") {
    return 15;
  }

  if (normalizedType === "TELECOMMANDE URMET") {
    return 40;
  }

  if (normalizedType === "CARTE TURBOSELF") {
    return 10;
  }

  if (normalizedType === "VENTILATEUR") {
    return normalizePricingKey(effect?.designation || "") === "VENTILATEUR SUR PIED" ? 35 : 30;
  }

  if (normalizedType === "RADIATEUR APPOINT") {
    return 45;
  }

  return 0;
}

function getEffectReplacementCause(person, effect) {
  if (typeof deriveEffectState === "function") {
    return deriveEffectState(person, effect).cause;
  }
  const persistedCause = normalizeEffectCause(effect?.cause || effect?.causeRemplacement);
  if (persistedCause) {
    return persistedCause;
  }
  if (!String(effect?.dateRetour || "").trim() && isExitDue(person)) {
    return "NON RENDU";
  }
  return "";
}

function getEffectReplacementCost(person, effect) {
  const cause = normalizeText(getEffectReplacementCause(person, effect));
  if (!cause) {
    return 0;
  }

  if (cause === "HS") {
    return 0;
  }

  return getReplacementCostValue(effect?.typeEffet, cause, effect?.designation || "");
}

function isEffectChargeable(person, effect) {
  if (typeof deriveEffectState === "function") {
    return deriveEffectState(person, effect).chargeable;
  }
  return getEffectReplacementCost(person, effect) > 0;
}

function syncReplacementCostField() {
  const form = document.getElementById("effect-form");
  if (!form) {
    return;
  }

  const typeEffet = form.elements.typeEffet?.value || "";
  const manualStatus = normalizeText(form.elements.statutManuel?.value || "");
  const billingCause = manualStatus === "PERDU"
    ? "PERTE"
    : manualStatus === "DETRUIT"
      ? "DETRUIT"
      : manualStatus === "VOL"
        ? "VOL"
        : "";
  const referenceId = form.elements.referenceEffet?.value || "";
  const reference = findReferenceById(referenceId);
  const designation =
    form.elements.designationLibre?.value || reference?.designation || "";
  const coutField = form.elements.coutRemplacement;
  if (!coutField) {
    return;
  }

  if (!normalizeText(typeEffet)) {
    coutField.value = "0,00 €";
    return;
  }

  if (manualStatus === "HS") {
    coutField.value = "0,00 €";
    return;
  }

  // En mode ACTIF (ou sans cause), on affiche le cout previsionnel standard de remplacement.
  const effectiveCause = billingCause || "PERTE";
  let previewCost = getReplacementCostValue(typeEffet, effectiveCause, designation);
  if (previewCost <= 0) {
    previewCost = getEffectUnitValue({ typeEffet, designation });
  }
  coutField.value = formatAmountWithEuro(previewCost);
}

function bindReferenceListForms() {
  document.querySelectorAll(".js-reference-list-form").forEach((form) => {
    form.onsubmit = (event) => {
      event.preventDefault();

      if (!state.data?.listes) {
        showDataStatus("DONNEES NON CHARGEES");
        return;
      }

      const listName = form.dataset.listName || "";
      const input = form.querySelector('input[name="value"]');
      const normalizeForList = (entry) =>
        listName === "causesRemplacement" ? normalizeReferenceCauseLabel(entry) : normalizeText(entry);
      const rawValue = normalizeText(input?.value);
      const value = normalizeForList(rawValue);
      if (!listName || !Array.isArray(state.data.listes[listName])) {
        return;
      }
      if (!value) {
        showDataStatus("VALEUR VIDE");
        return;
      }

      const currentEdit = state.editingSimpleReference;
      const list = state.data.listes[listName];

      if (currentEdit && currentEdit.listName === listName) {
        const oldValue = currentEdit.originalValue;
        const duplicate = list.some(
          (entry) => normalizeForList(entry) === value && normalizeForList(entry) !== normalizeForList(oldValue)
        );
        if (duplicate) {
          showDataStatus("VALEUR DEJA PRESENTE");
          return;
        }

        const index = list.findIndex((entry) => normalizeForList(entry) === normalizeForList(oldValue));
        if (index >= 0) {
          pushUndoSnapshot("MODIFICATION BASE");
          list[index] = value;
          cascadeSimpleReferenceRename(listName, oldValue, value);
          sortListValues(list);
          state.editingSimpleReference = null;
          markDirty();
          hydrateStaticLists();
          schedulePageRender();
          form.reset();
          const submitButton = form.querySelector('button[type="submit"]');
          if (submitButton) {
            submitButton.textContent = "ENREGISTRER";
          }
          showActionStatus("update", `BASE MISE A JOUR : ${value}`);
          return;
        }
      }

      if (list.some((entry) => normalizeForList(entry) === value)) {
        showDataStatus("VALEUR DEJA PRESENTE");
        return;
      }

      pushUndoSnapshot("AJOUT BASE");
      list.push(value);
      sortListValues(list);
      state.editingSimpleReference = null;
      markDirty();
      hydrateStaticLists();
      schedulePageRender();
      form.reset();
      const submitButton = form.querySelector('button[type="submit"]');
      if (submitButton) {
        submitButton.textContent = "ENREGISTRER";
      }
      showActionStatus("create", `BASE AJOUTEE : ${value}`);
    };
  });
}

function bindRepresentativeSignatoryForm() {
  const form = document.getElementById("representative-signatory-form");
  if (!form) {
    return;
  }

  form.onsubmit = (event) => {
    event.preventDefault();

    if (!state.data?.listes?.representantsSignataires) {
      showDataStatus("DONNEES NON CHARGEES");
      return;
    }

    const representativeName = normalizeText(form.elements.representativeName?.value);
    const representativeFunction = normalizeText(form.elements.representativeFunction?.value);
    if (!representativeName && !representativeFunction) {
      showDataStatus("NOM ET FONCTION VIDES");
      return;
    }

    const existingMatch = findRepresentativeByValues(representativeName, representativeFunction);
    const editingRepresentative = state.editingRepresentativeId
      ? findRepresentativeById(state.editingRepresentativeId)
      : null;

    if (editingRepresentative) {
      if (existingMatch && existingMatch.id !== editingRepresentative.id) {
        pushUndoSnapshot("FUSION REPRESENTANT");
        updateRepresentativeLinks(editingRepresentative.id, existingMatch);
        state.data.listes.representantsSignataires = state.data.listes.representantsSignataires.filter(
          (entry) => entry.id !== editingRepresentative.id
        );
        state.editingRepresentativeId = "";
        markDirty();
        schedulePageRender();
        form.reset();
        showActionStatus("update", "REPRESENTANT FUSIONNE");
        return;
      }

      pushUndoSnapshot("MODIFICATION REPRESENTANT");
      editingRepresentative.nom = representativeName;
      editingRepresentative.fonction = representativeFunction;
      updateRepresentativeLinks(editingRepresentative.id, editingRepresentative);
      sortRepresentatives();
      state.editingRepresentativeId = "";
      markDirty();
      schedulePageRender();
      form.reset();
      showActionStatus("update", "REPRESENTANT MIS A JOUR");
      return;
    }

    if (existingMatch) {
      showDataStatus("REPRESENTANT DEJA PRESENT");
      return;
    }

    pushUndoSnapshot("AJOUT REPRESENTANT");
    state.data.listes.representantsSignataires.push({
      id: getNextId("REP", state.data.listes.representantsSignataires),
      nom: representativeName,
      fonction: representativeFunction,
    });
    sortRepresentatives();
    markDirty();
    schedulePageRender();
    form.reset();
    showActionStatus("create", "REPRESENTANT AJOUTE");
  };
}

function bindReferenceEffectForm() {
  const form = document.getElementById("reference-effect-form");
  if (!form) {
    return;
  }

  const typeField = form.elements.referenceTypeEffet;
  const siteField = form.elements.referenceSite;
  const designationField = form.elements.referenceDesignation;
  if (typeField) {
    typeField.onchange = () => {
      updateReferenceEffectFormMode(typeField.value);
      syncReferenceSitesSelector();
    };
  }
  if (siteField) {
    siteField.onchange = () => {};
  }
  if (designationField) {
    let referenceDesignationInputDebounceId = 0;
    const scheduleReferenceDesignationUpdate = () => {
      if (referenceDesignationInputDebounceId) {
        window.clearTimeout(referenceDesignationInputDebounceId);
      }
      referenceDesignationInputDebounceId = window.setTimeout(() => {
        referenceDesignationInputDebounceId = 0;
        if (normalizeText(typeField?.value || "") === "CLE CES") {
          designationField.value = normalizeCesDesignationLabel(designationField.value);
        }
        syncReferenceSitesSelector();
      }, FILTER_INPUT_DEBOUNCE_MS);
    };

    designationField.oninput = () => {
      scheduleReferenceDesignationUpdate();
    };
  }

  form.onsubmit = (event) => {
    event.preventDefault();

    if (!state.data?.listes?.referencesEffets) {
      showDataStatus("DONNEES NON CHARGEES");
      return;
    }

    const formData = new FormData(form);
    const normalizedTypeEffet = normalizeText(formData.get("referenceTypeEffet"));
    const referenceId = state.editingReferenceId || getNextId("REF", state.data.listes.referencesEffets);
    const selectedSite = normalizedTypeEffet === "CARTE TURBOSELF"
      ? ALL_SITES_VALUE
      : (normalizeText(formData.get("referenceSite")) || "SANS SITE");
    const normalizedDesignation = normalizeReferenceDesignationByType(
      normalizedTypeEffet,
      formData.get("referenceDesignation")
    );
    const reference = {
      id: referenceId,
      site: selectedSite,
      sitesAffectation: normalizeSites([selectedSite]),
      typeEffet: normalizedTypeEffet || "EFFET",
      designation: normalizedDesignation || `REFERENCE ${referenceId}`,
    };
    reference.site = getReferenceSiteLabel(reference) || reference.site;

    if (!typeUsesReferenceCatalog(reference.typeEffet)) {
      showDataStatus("POUR CE TYPE : PAS DE DESIGNATION EN BASE");
      window.alert("POUR CE TYPE : UTILISER SEULEMENT LE N° D'IDENTIFICATION");
      return;
    }

    const currentIndex = state.data.listes.referencesEffets.findIndex((entry) => entry.id === referenceId);
    const duplicate = state.data.listes.referencesEffets.some(
      (entry) =>
        entry.id !== referenceId &&
        haveSameSites(getReferenceSites(entry), reference.sitesAffectation) &&
        normalizeText(entry.typeEffet) === reference.typeEffet &&
        normalizeText(entry.designation) === reference.designation
    );
    if (duplicate) {
      showDataStatus("REFERENCE DEJA PRESENTE");
      return;
    }

    pushUndoSnapshot(currentIndex >= 0 ? "MODIFICATION REFERENCE" : "AJOUT REFERENCE");
    if (currentIndex >= 0) {
      const previous = state.data.listes.referencesEffets[currentIndex];
      state.data.listes.referencesEffets[currentIndex] = reference;
      cascadeReferenceEffectUpdate(previous, reference);
    } else {
      state.data.listes.referencesEffets.push(reference);
    }

    sortReferenceEffects();
    state.editingReferenceId = "";
    markDirty();
    hydrateStaticLists();
    schedulePageRender();
    form.reset();
    renderReferenceSitesSelector([]);
    updateReferenceEffectFormMode("");
    const submitButton = form.querySelector('button[type="submit"]');
    if (submitButton) {
      submitButton.textContent = "ENREGISTRER LA REFERENCE";
    }
    showActionStatus(
      currentIndex >= 0 ? "update" : "create",
      currentIndex >= 0
        ? `REFERENCE MISE A JOUR : ${reference.designation}`
        : `REFERENCE AJOUTEE : ${reference.designation}`
    );
  };
}

function updateReferenceEffectFormMode(typeEffet) {
  const field = document.getElementById("reference-sites-field");
  if (!field) {
    return;
  }
  field.classList.add("is-hidden");
}

function normalizeCesDesignationLabel(value) {
  const raw = String(value || "").trim().toUpperCase();
  if (!raw) {
    return "";
  }
  const compact = raw.replace(/_/g, "-").replace(/\s+/g, " ");
  const withoutPrefix = compact.replace(/^(?:CES[\s-]*)+/i, "").trim();
  if (!withoutPrefix) {
    return "CES";
  }
  const suffix = withoutPrefix.replace(/\s*-\s*/g, "-").replace(/\s+/g, "-");
  return `CES-${suffix}`;
}

function normalizeReferenceDesignationByType(typeEffet, designation) {
  const normalizedType = normalizeText(typeEffet);
  if (normalizedType === "CLE CES") {
    return normalizeCesDesignationLabel(designation);
  }
  return normalizeText(designation);
}

function canUseAllSitesForReference() {
  const form = document.getElementById("reference-effect-form");
  if (!form) {
    return false;
  }
  const typeEffet = normalizeText(form.elements.referenceTypeEffet?.value);
  const designation = normalizeText(form.elements.referenceDesignation?.value);
  return typeEffet === "CLE CES" && designation === "CES-PG";
}

function syncReferenceSitesSelector() {
  const container = document.getElementById("reference-sites-selector");
  if (!container) {
    return;
  }

  const allowAllSites = canUseAllSitesForReference();
  const allSitesItem = container.querySelector('input[name="referenceSites"][value="TOUS SITES"]')?.closest(".site-selector__item");
  const allSitesCheckbox = container.querySelector('input[name="referenceSites"][value="TOUS SITES"]');

  if (allSitesItem) {
    allSitesItem.classList.toggle("is-hidden", !allowAllSites);
  }

  if (!allowAllSites && allSitesCheckbox) {
    allSitesCheckbox.checked = false;
  }
}

function getReplacementCostKey(typeEffet, cause, designation = "") {
  return `${normalizeText(typeEffet)}__${normalizeText(cause)}__${normalizeText(designation)}`;
}

function bindReplacementCostForm() {
  const form = document.getElementById("replacement-cost-form");
  if (!form) {
    return;
  }

  form.onsubmit = (event) => {
    event.preventDefault();

    if (!state.data?.listes?.coutsRemplacement) {
      showDataStatus("DONNEES NON CHARGEES");
      return;
    }

    const formData = new FormData(form);
    const typeEffet = normalizeText(formData.get("costTypeEffet"));
    const cause = normalizeText(formData.get("costCauseRemplacement"));
    const montant = normalizeAmount(formData.get("costMontant"));

    if (!typeEffet || !cause) {
      showDataStatus("TYPE D'EFFET ET CAUSE OBLIGATOIRES");
      return;
    }

    const nextCostKey = getReplacementCostKey(typeEffet, cause);
    const lookupKey = state.editingReplacementCostKey || nextCostKey;
    const currentIndex = state.data.listes.coutsRemplacement.findIndex(
      (entry) => getReplacementCostKey(entry.typeEffet, entry.cause, entry.designation) === lookupKey
    );

    pushUndoSnapshot(currentIndex >= 0 ? "MODIFICATION COUT" : "AJOUT COUT");
    if (currentIndex >= 0) {
      state.data.listes.coutsRemplacement[currentIndex] = {
        ...state.data.listes.coutsRemplacement[currentIndex],
        typeEffet,
        cause,
        montant,
      };
    } else {
      state.data.listes.coutsRemplacement.push({ typeEffet, cause, montant });
    }

  state.editingReplacementCostKey = "";
  markDirty();
  schedulePageRender();
  form.reset();
    const submitButton = form.querySelector('button[type="submit"]');
    if (submitButton) {
      submitButton.textContent = "ENREGISTRER LE COUT";
    }
    showActionStatus("update", `COUT MIS A JOUR : ${typeEffet} / ${cause}`);
  };
}

function bindStockAdjustmentForm() {
  const form = document.getElementById("stock-adjustment-form");
  if (!(form instanceof HTMLFormElement)) {
    return;
  }
  const saveButton = document.getElementById("stock-save-movement");
  let stockSaveInFlight = false;

  const saveStockMovement = async () => {
    if (stockSaveInFlight) {
      return;
    }
    stockSaveInFlight = true;
    if (saveButton instanceof HTMLButtonElement) {
      saveButton.disabled = true;
    }
    if (!state.data) {
      showStockAdjustmentStatus("DONNEES NON CHARGEES", "error");
      stockSaveInFlight = false;
      if (saveButton instanceof HTMLButtonElement) {
        saveButton.disabled = false;
      }
      return;
    }
    if (!Array.isArray(state.data.stocksEffetsManuels)) {
      state.data.stocksEffetsManuels = [];
    }

    try {
      const formData = new FormData(form);
      const typeEffet = normalizeText(formData.get("stockTypeEffet"));
      const site = normalizeText(formData.get("stockSite"));
      const referenceEffetId = String(formData.get("stockReferenceId") || "");
      let syntheticReference = parseStockSyntheticReferenceValue(referenceEffetId);
      const reference = referenceEffetId === ALL_DESIGNATIONS_VALUE || syntheticReference ? null : findReferenceById(referenceEffetId);
      const resolvedSite = reference ? resolveStockMovementSite(site, reference) : normalizeText(syntheticReference?.site || site);
      const movementSite = resolvedSite || site;
      const effectiveTypeEffet = reference ? getReferenceEffectiveType(reference) : typeEffet;
      let designation = getStockGroupingDesignation(
        effectiveTypeEffet,
        reference ? getStockReferenceDesignation(reference) : syntheticReference?.designation || ""
      );
      const action = normalizeText(formData.get("stockAction"));
      const quantite = Math.max(1, Number.parseInt(String(formData.get("stockQuantity") || "1"), 10) || 1);
      const motif = normalizeText(formData.get("stockReason"));
      const commentaire = normalizeText(formData.get("stockComment"));

      if (!reference && !syntheticReference && typeEffet && site) {
        const matchingUndesignatedRow = getStockSummaryRows().find(
          (row) =>
            normalizeText(row.site) === site &&
            normalizeText(row.typeEffet) === typeEffet &&
            normalizeText(row.designation) === getStockGroupingDesignation(typeEffet, row.designation)
        );
        if (matchingUndesignatedRow || typeIgnoresStockDesignation(typeEffet)) {
          syntheticReference = {
            site,
            typeEffet,
            designation: getStockGroupingDesignation(typeEffet, ""),
          };
          designation = getStockGroupingDesignation(typeEffet, "");
        }
      }

      if (!typeEffet || !movementSite || (!reference && !syntheticReference) || !action) {
        showStockAdjustmentStatus("TYPE, SITE ET MOUVEMENT OBLIGATOIRES", "error");
        return;
      }
      const isExactSyntheticStockRow = Boolean(syntheticReference);
      if (
        typeEffet === ALL_TYPES_VALUE ||
        referenceEffetId === ALL_DESIGNATIONS_VALUE ||
        (site === ALL_SITES_VALUE &&
          movementSite === ALL_SITES_VALUE &&
          !isExactSyntheticStockRow &&
          !referenceHasSite(reference, ALL_SITES_VALUE))
      ) {
        showStockAdjustmentStatus("POUR ENREGISTRER : CHOISIR SITE/TYPE/DESIGNATION PRECIS", "error");
        return;
      }
      if (reference && !isReferenceEffectActive(reference)) {
        showStockAdjustmentStatus("REFERENCE EFFET DESACTIVEE - MOUVEMENT BLOQUE", "error");
        return;
      }
      if (reference && !referenceMatchesType(reference, typeEffet)) {
        showStockAdjustmentStatus("TYPE D'EFFET ET DESIGNATION INCOHERENTS", "error");
        return;
      }
      if (reference && !referenceHasSite(reference, movementSite)) {
        showStockAdjustmentStatus("SITE ET DESIGNATION INCOHERENTS", "error");
        return;
      }
      if (!reference) {
        if (syntheticReference.site !== movementSite || syntheticReference.typeEffet !== typeEffet) {
          showStockAdjustmentStatus("TYPE, SITE ET DESIGNATION STOCK INCOHERENTS", "error");
          return;
        }
        const stockRowExists = getStockSummaryRows().some(
          (row) =>
            normalizeText(row.site) === movementSite &&
            normalizeText(row.typeEffet) === typeEffet &&
            normalizeText(row.designation) === getStockGroupingDesignation(typeEffet, designation)
        );
        if (!stockRowExists && !typeIgnoresStockDesignation(typeEffet)) {
          showStockAdjustmentStatus("MOUVEMENT STOCK BLOQUE : LIGNE ABSENTE DE LA SYNTHESE STOCK", "error");
          return;
        }
      }

      pushUndoSnapshot("MOUVEMENT STOCK MANUEL");
      const movement = {
        id: getNextId("STKM", state.data.stocksEffetsManuels),
        typeEffet: effectiveTypeEffet,
        site: movementSite,
        referenceEffetId: reference ? String(reference.id || "") : "",
        designation,
        action,
        quantite,
        motif,
        commentaire,
        source: "MANUEL",
        date: getTodayIsoDate(),
      };
      state.data.stocksEffetsManuels.push(movement);
      state.stockHighlightKey = `${effectiveTypeEffet}__${movementSite}__${designation}`;
      markDirty();
      state.stockTableFilters = {
        site: movementSite,
        typeEffet: effectiveTypeEffet,
        referenceEffetId: reference ? String(reference.id || "") : referenceEffetId,
      };
      renderStockMovementsTable();
      renderStockSummaryTable();
      renderStockInstantKpi();
      renderReferenceCounts();
      showActionStatus("create", `MOUVEMENT STOCK ENREGISTRE : ${effectiveTypeEffet} / ${designation}`);
      showStockAdjustmentStatus(`MOUVEMENT STOCK AJOUTE : ${effectiveTypeEffet} / ${designation} - SAUVEGARDE EN COURS`, "success");
      setStockResetButtonPending(true);
      saveDataToFile({
        silent: true,
        reloadAfter: false,
        promptDownload: false,
        reloadOnConflict: false,
      }).then(() => {
        if (state.isDirty) {
          showStockAdjustmentStatus("MOUVEMENT STOCK AJOUTE - SAUVEGARDE A RELANCER", "warning");
          return;
        }
        showStockAdjustmentStatus(`MOUVEMENT STOCK SAUVEGARDE : ${effectiveTypeEffet} / ${designation}`, "success");
      });
    } finally {
      stockSaveInFlight = false;
      if (saveButton instanceof HTMLButtonElement) {
        saveButton.disabled = false;
      }
    }
  };

  form.onsubmit = async (event) => {
    event.preventDefault();
    await saveStockMovement();
  };

  if (saveButton instanceof HTMLButtonElement) {
    saveButton.onclick = async (event) => {
      event.preventDefault();
      await saveStockMovement();
    };
  }

  const typeSelect = form.elements.stockTypeEffet;
  const siteSelect = form.elements.stockSite;
  const referenceSelect = form.elements.stockReferenceId;
  if (typeSelect instanceof HTMLSelectElement) {
    typeSelect.onchange = () => {
      updateStockDesignationOptions();
      refreshStockTableFiltersFromForm();
      renderStockMovementsTable();
      renderStockSummaryTable();
      renderStockTypeKpis();
      renderStockInstantKpi();
    };
  }
  if (siteSelect instanceof HTMLSelectElement) {
    siteSelect.onchange = () => {
      updateStockDesignationOptions();
      refreshStockTableFiltersFromForm();
      renderStockMovementsTable();
      renderStockSummaryTable();
      renderStockTypeKpis();
      renderStockInstantKpi();
    };
  }
  if (referenceSelect instanceof HTMLSelectElement) {
    referenceSelect.onchange = () => {
      refreshStockTableFiltersFromForm();
      renderStockMovementsTable();
      renderStockSummaryTable();
      renderStockTypeKpis();
      renderStockInstantKpi();
    };
  }
  const resetButton = document.getElementById("stock-reset-filters");
  if (resetButton instanceof HTMLButtonElement) {
    resetButton.onclick = () => {
      resetReferenceBasesPageFields();
      setStockResetButtonPending(false);
    };
  }
}

function bindStockMovementActions() {
  const body = document.getElementById("stock-movements-table-body");
  if (!body || body.dataset.bound === "true") {
    return;
  }
  body.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const deleteButton = target.closest(".js-delete-stock-movement");
    if (deleteButton instanceof HTMLElement) {
      event.preventDefault();
      const movementId = String(deleteButton.dataset.stockMovementId || "");
      const movement = (state.data?.stocksEffetsManuels || []).find((entry) => String(entry.id || "") === movementId);
      if (!movement) {
        return;
      }
      const linkedReference = movement.referenceEffetId ? findReferenceById(movement.referenceEffetId) : null;
      if (movement.referenceEffetId && !linkedReference) {
        showDataStatus("MOUVEMENT STOCK VERROUILLE : REFERENCE ABSENTE EN BASE");
        return;
      }
      if (!window.confirm(`SUPPRIMER LE MOUVEMENT STOCK : ${movement.typeEffet} / ${movement.designation} ?`)) {
        return;
      }
      pushUndoSnapshot("SUPPRESSION MOUVEMENT STOCK");
      state.data.stocksEffetsManuels = state.data.stocksEffetsManuels.filter((entry) => String(entry.id || "") !== movementId);
      markDirty();
      renderReferenceBases();
      showActionStatus("delete", "MOUVEMENT STOCK SUPPRIME");
      return;
    }

    const row = target.closest(".js-stock-movement-row");
    if (!(row instanceof HTMLElement)) {
      return;
    }

    hydrateStockAdjustmentFromSelection({
      site: String(row.dataset.site || ""),
      typeEffet: String(row.dataset.typeEffet || ""),
      referenceEffetId: String(row.dataset.referenceEffetId || ""),
      designation: String(row.dataset.designation || ""),
    });
  });
  body.dataset.bound = "true";
}

function bindStockSummaryActions() {
  const body = document.getElementById("stock-summary-table-body");
  if (!body || body.dataset.bound === "true") {
    return;
  }
  body.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const row = target.closest(".js-stock-summary-row");
    if (!(row instanceof HTMLElement)) {
      return;
    }
    hydrateStockAdjustmentFromSelection({
      site: String(row.dataset.site || ""),
      typeEffet: String(row.dataset.typeEffet || ""),
      referenceEffetId: "",
      designation: String(row.dataset.designation || ""),
    });
  });
  body.dataset.bound = "true";
}

function hydrateStockAdjustmentFromSelection(selection) {
  const form = document.getElementById("stock-adjustment-form");
  if (!(form instanceof HTMLFormElement)) {
    return;
  }

  const site = normalizeText(selection?.site || "");
  const typeEffet = normalizeText(selection?.typeEffet || "");
  const referenceEffetId = String(selection?.referenceEffetId || "");
  const designation = normalizeText(selection?.designation || "");

  const siteSelect = form.elements.stockSite;
  const typeSelect = form.elements.stockTypeEffet;
  const referenceSelect = form.elements.stockReferenceId;
  if (!(siteSelect instanceof HTMLSelectElement) || !(typeSelect instanceof HTMLSelectElement) || !(referenceSelect instanceof HTMLSelectElement)) {
    return;
  }

  if (site && Array.from(siteSelect.options).some((opt) => normalizeText(opt.value) === site)) {
    siteSelect.value = site;
  }
  if (typeEffet && Array.from(typeSelect.options).some((opt) => normalizeText(opt.value) === typeEffet)) {
    typeSelect.value = typeEffet;
  }

  updateStockDesignationOptions();

  let resolvedReferenceId = "";
  if (referenceEffetId && Array.from(referenceSelect.options).some((opt) => String(opt.value || "") === referenceEffetId)) {
    resolvedReferenceId = referenceEffetId;
  } else {
    const match = (state.data?.listes?.referencesEffets || []).find((reference) => {
      if (!isReferenceEffectActive(reference)) {
        return false;
      }
      if (getStockReferenceDesignation(reference) !== designation) {
        return false;
      }
      if (!referenceMatchesType(reference, typeEffet)) {
        return false;
      }
      return referenceHasSite(reference, site);
    });
    if (match && Array.from(referenceSelect.options).some((opt) => String(opt.value || "") === String(match.id || ""))) {
      resolvedReferenceId = String(match.id || "");
    }
  }
  if (!resolvedReferenceId && designation) {
    const syntheticValue = getStockSyntheticReferenceValue(site, typeEffet, designation);
    if (Array.from(referenceSelect.options).some((opt) => String(opt.value || "") === syntheticValue)) {
      resolvedReferenceId = syntheticValue;
    }
  }

  if (resolvedReferenceId) {
    referenceSelect.value = resolvedReferenceId;
  }

  refreshStockTableFiltersFromForm();
  renderStockMovementsTable();
  renderStockSummaryTable();
  renderStockTypeKpis();
  renderStockInstantKpi();
}

function getReferenceCauseOptions() {
  const baseCauses = Array.isArray(state.data?.listes?.causesRemplacement)
    ? state.data.listes.causesRemplacement
    : [];
  const normalizedFromBase = baseCauses.map(normalizeReferenceCauseLabel).filter(Boolean);
  const normalizedFromCosts = (state.data?.listes?.coutsRemplacement || [])
    .map((entry) => normalizeReferenceCauseLabel(entry?.cause))
    .filter(Boolean);
  const causes = Array.from(new Set([...normalizedFromBase, ...normalizedFromCosts]));
  return causes.length ? causes : [...EFFECT_STATUS_CAUSES];
}

function normalizeReferenceCauseLabel(value) {
  const normalized = normalizeText(value);
  if (normalized === "CASSE") return "DETRUIT";
  if (normalized === "PERDU") return "PERTE";
  return normalized;
}

function resetReferenceBasesPageFields() {
  if (document.body?.dataset?.page !== "reference-bases") {
    return;
  }
  state.isReferencePageResetting = true;

  const formIds = [
    "mobile-signature-settings-form",
    "representative-signatory-form",
    "reference-filter-form",
    "reference-effect-form",
    "stock-adjustment-form",
    "replacement-cost-form",
  ];

  formIds.forEach((formId) => {
    const form = document.getElementById(formId);
    if (form instanceof HTMLFormElement) {
      clearFormSearchFields(form);
      form.reset();
    }
  });

  state.editingSimpleReference = null;
  state.editingReferenceId = "";
  state.editingRepresentativeId = "";
  state.editingReplacementCostKey = "";
  state.stockTableFilters = { site: "", typeEffet: "", referenceEffetId: "" };

  const referenceSubmit = document.querySelector('#reference-effect-form button[type="submit"]');
  if (referenceSubmit instanceof HTMLButtonElement) {
    referenceSubmit.textContent = "ENREGISTRER LA REFERENCE";
  }

  const representativeSubmit = document.querySelector('#representative-signatory-form button[type="submit"]');
  if (representativeSubmit instanceof HTMLButtonElement) {
    representativeSubmit.textContent = "ENREGISTRER LE REPRESENTANT";
  }

  const replacementCostSubmit = document.querySelector('#replacement-cost-form button[type="submit"]');
  if (replacementCostSubmit instanceof HTMLButtonElement) {
    replacementCostSubmit.textContent = "ENREGISTRER LE COUT";
  }

  updateReferenceEffectFormMode("");
  updateStockDesignationOptions();
  refreshStockTableFiltersFromForm();
  renderReferenceBases();

  const filterForm = document.getElementById("reference-filter-form");
  if (filterForm instanceof HTMLFormElement) {
    const searchField = filterForm.elements.filterReferenceSearch;
    const siteField = filterForm.elements.filterReferenceSite;
    const typeField = filterForm.elements.filterReferenceTypeEffet;
    if (searchField instanceof HTMLInputElement) {
      searchField.value = "";
      searchField.defaultValue = "";
    }
    if (siteField instanceof HTMLSelectElement) {
      siteField.value = "";
    }
    if (typeField instanceof HTMLSelectElement) {
      typeField.value = "";
    }
  }
  const editorForm = document.getElementById("reference-effect-form");
  if (editorForm instanceof HTMLFormElement) {
    if (editorForm.elements.referenceSite instanceof HTMLSelectElement) {
      editorForm.elements.referenceSite.value = "";
    }
    if (editorForm.elements.referenceTypeEffet instanceof HTMLSelectElement) {
      editorForm.elements.referenceTypeEffet.value = "";
    }
  }
  renderReferenceEffectsTable(state.referenceRenderContext || buildReferenceRenderContext());
  state.isReferencePageResetting = false;
}

function bindReferenceBaseResetButtons() {
  if (document.body?.dataset?.page !== "reference-bases") {
    return;
  }
  document.querySelectorAll("button").forEach((button) => {
    if (!(button instanceof HTMLButtonElement)) {
      return;
    }
    if (normalizeText(button.textContent || "") !== "REINITIALISER") {
      return;
    }
    if (button.dataset.boundGlobalReset === "1") {
      return;
    }
    button.dataset.boundGlobalReset = "1";
    button.addEventListener("click", (event) => {
      event.preventDefault();
      resetReferenceBasesPageFields();
    });
  });
}

function bindReferenceFilters() {
  const form = document.getElementById("reference-filter-form");
  if (!form) {
    return;
  }

  const scheduleReferenceFilterUpdate = () => {
    if (state.referenceFilterDebounceTimerId) {
      window.clearTimeout(state.referenceFilterDebounceTimerId);
    }
    state.referenceFilterDebounceTimerId = window.setTimeout(() => {
      state.referenceFilterDebounceTimerId = 0;
      renderReferenceEffectsTable();
    }, FILTER_INPUT_DEBOUNCE_MS);
  };

  const applyReferenceReset = () => {
    if (state.referenceFilterDebounceTimerId) {
      window.clearTimeout(state.referenceFilterDebounceTimerId);
      state.referenceFilterDebounceTimerId = 0;
    }
    resetReferenceBasesPageFields();
  };

  form.oninput = () => {
    scheduleReferenceFilterUpdate();
  };

  form.onreset = (event) => {
    event.preventDefault();
    applyReferenceReset();
  };

  const searchField = form.elements.filterReferenceSearch;
  if (searchField) {
    searchField.addEventListener("search", () => {
      if (!String(searchField.value || "").trim()) {
        applyReferenceReset();
      }
    });
  }
}

function bindMobileSignatureSettingsForm() {
  const form = document.getElementById("mobile-signature-settings-form");
  if (!(form instanceof HTMLFormElement)) {
    return;
  }

  form.onsubmit = (event) => {
    event.preventDefault();
    if (!state.data?.meta) {
      return;
    }

    const rawValue = String(form.elements.mobileSignatureBaseUrl?.value || "").trim();
    const normalized = normalizeMobileSignatureBaseUrl(rawValue);
    if (rawValue && !normalized) {
      showDataStatus("URL INVALIDE (UTILISER HTTP OU HTTPS)");
      form.elements.mobileSignatureBaseUrl?.focus();
      return;
    }

    pushUndoSnapshot("CONFIGURATION URL SIGNATURE MOBILE");
    state.data.meta.signatureMobileBaseUrl = normalized;
    state.mobileSignatureNetworkInfo = null;
    markDirty();
    renderMobileSignatureSettings();
    showActionStatus("update", normalized ? "URL PUBLIQUE SIGNATURE MOBILE ENREGISTREE" : "URL PUBLIQUE SIGNATURE MOBILE VIDEE");
  };
}

function readFilters(form) {
  const formData = new FormData(form);
  return {
    search: normalizeText(formData.get("search") || formData.get("person-picker-search")),
    site: normalizeText(formData.get("site")),
    typePersonnel: normalizeText(formData.get("typePersonnel")),
    typeContrat: normalizeText(formData.get("typeContrat")),
    statutDossier: normalizeText(formData.get("statutDossier")),
    statutObjet: normalizeText(formData.get("statutObjet")),
    typeEffet: normalizeText(formData.get("typeEffet")),
  };
}

function hydrateStaticLists() {
  if (!state.data?.listes) {
    return;
  }

  const {
    sites = [],
    typesPersonnel = [],
    typesContrats = [],
    fonctions = [],
    typesEffets = [],
    statutsObjetManuels = [],
  } = state.data.listes;

  populateSelect('select[name="typePersonnel"]', typesPersonnel);
  populateSelect('select[name="typeContrat"]', typesContrats);
  populateSelect('select[name="site"]', sites);
  populateSelect('select[name="sheetTypePersonnel"]', typesPersonnel);
  populateSelect('select[name="sheetTypeContrat"]', typesContrats);
  populateSelect('select[name="sheetFonction"]', fonctions);
  populateSelect('select[name="typeEffet"]', typesEffets);
  populateSelect('select[name="referenceSite"]', sites);
  populateSelect('select[name="referenceTypeEffet"]', typesEffets);
  populateSelect('select[name="filterReferenceSite"]', sites);
  populateSelect('select[name="filterReferenceTypeEffet"]', typesEffets);
  populateSelect('select[name="statutManuel"]', statutsObjetManuels);
  populateSelect('select[name="costTypeEffet"]', typesEffets);
  populateSelect('select[name="costCauseRemplacement"]', getReferenceCauseOptions());
  renderSiteSelector("add-site-selector", "add", []);
  renderSiteSelector("sheet-site-selector", "sheet", getPersonSites(getCurrentPerson()));
  renderReferenceSitesSelector([]);
}

function renderSiteSelector(containerId, prefix, selectedSites = []) {
  const container = document.getElementById(containerId);
  if (!container || !state.data?.listes?.sites) {
    return;
  }

  const normalizedSelectedSites = normalizeSites(selectedSites);
  const selectedAllSites = normalizedSelectedSites.includes(ALL_SITES_VALUE);
  const items = Array.from(new Set([ALL_SITES_VALUE, ...(state.data.listes.sites || [])]))
    .map((site) => {
      const normalizedSite = normalizeText(site);
      const checked = selectedAllSites
        ? normalizedSite === ALL_SITES_VALUE
        : normalizedSelectedSites.includes(normalizedSite);
      return `<label class="site-selector__item">
        <input type="checkbox" name="${prefix}Sites" value="${escapeHtml(site)}" ${checked ? "checked" : ""} />
        <span>${escapeHtml(site)}</span>
      </label>`;
    })
    .join("");

  container.innerHTML = items;

  const checkboxes = Array.from(container.querySelectorAll(`input[name="${prefix}Sites"]`));
  checkboxes.forEach((checkbox) => {
    checkbox.onchange = () => {
      const normalizedValue = normalizeText(checkbox.value);
      if (normalizedValue === ALL_SITES_VALUE && checkbox.checked) {
        checkboxes.forEach((entry) => {
          if (normalizeText(entry.value) !== ALL_SITES_VALUE) {
            entry.checked = false;
          }
        });
      }

      if (normalizedValue !== ALL_SITES_VALUE && checkbox.checked) {
        const allSitesCheckbox = checkboxes.find((entry) => normalizeText(entry.value) === ALL_SITES_VALUE);
        if (allSitesCheckbox) {
          allSitesCheckbox.checked = false;
        }
      }
    };
  });
}

function renderReferenceSitesSelector(selectedSites = []) {
  const container = document.getElementById("reference-sites-selector");
  if (!container || !state.data?.listes?.sites) {
    return;
  }

  const normalizedSelectedSites = normalizeSites(selectedSites);
  const selectedAllSites = normalizedSelectedSites.includes(ALL_SITES_VALUE);
  const items = Array.from(new Set([ALL_SITES_VALUE, ...(state.data.listes.sites || [])]))
    .map((site) => {
      const normalizedSite = normalizeText(site);
      const checked = selectedAllSites
        ? normalizedSite === ALL_SITES_VALUE
        : normalizedSelectedSites.includes(normalizedSite);
      return `<label class="site-selector__item">
        <input type="checkbox" name="referenceSites" value="${escapeHtml(site)}" ${checked ? "checked" : ""} />
        <span>${escapeHtml(site)}</span>
      </label>`;
    })
    .join("");

  container.innerHTML = items;

  const checkboxes = Array.from(container.querySelectorAll('input[name="referenceSites"]'));
  checkboxes.forEach((checkbox) => {
    checkbox.onchange = () => {
      const normalizedValue = normalizeText(checkbox.value);
      if (normalizedValue === ALL_SITES_VALUE && checkbox.checked) {
        checkboxes.forEach((entry) => {
          if (normalizeText(entry.value) !== ALL_SITES_VALUE) {
            entry.checked = false;
          }
        });
      }

      if (normalizedValue !== ALL_SITES_VALUE && checkbox.checked) {
        const allSitesCheckbox = checkboxes.find((entry) => normalizeText(entry.value) === ALL_SITES_VALUE);
        if (allSitesCheckbox) {
          allSitesCheckbox.checked = false;
        }
      }
    };
  });

  syncReferenceSitesSelector();
}

function populateSelect(selector, values) {
  const elements = document.querySelectorAll(selector);
  elements.forEach((element) => {
    const firstOption = element.querySelector("option");
    const baseValue = firstOption ? firstOption.outerHTML : "";
    const currentValue = normalizeText(element.value);
    const options = values
      .map((value) => `<option value="${value}">${normalizeText(value)}</option>`)
      .join("");
    element.innerHTML = `${baseValue}${options}`;
    if (currentValue) {
      element.value = currentValue;
    }
  });
}

function getDocumentViewStateSignature(docType, personId = "") {
  const isPdfMode = isPdfRenderMode() ? "pdf" : "ui";
  const normalizedType = normalizeText(docType) === "EXIT" ? "exit" : "arrival";
  const activePerson = (state.data?.personnes || []).find(
    (entry) => String(entry?.id || "") === String(personId || "")
  );
  const sortConfig = state.tableSorts?.[`${normalizedType}Effects`] || {};
  const noPersonKey = [
    normalizedType,
    "no-person",
    isPdfMode,
    String(sortConfig.key || ""),
    String(sortConfig.dir || ""),
    String(new URLSearchParams(window.location.search).get("mode") || ""),
  ].join("|");

  if (!activePerson) {
    return noPersonKey;
  }

  if (normalizedType === "arrival") {
    const explicitMode = normalizeText(new URLSearchParams(window.location.search).get("mode") || "");
    const computedMode = getDocumentArchiveMode(activePerson, "arrival");
    const mode = explicitMode || computedMode;
    const isComplement = mode === "COMPLEMENTAIRE";
    const representative = getRepresentativeInfo(activePerson, "arrival");
    const allEffects = Array.isArray(activePerson.effetsConfies) ? activePerson.effetsConfies : [];
    const fallbackMovements = isComplement ? getArrivalComplementMovementMap(activePerson, allEffects) : null;
    const activeEffects = isComplement
      ? allEffects
      : allEffects.filter((effect) => Boolean(effect.dateRemise));
    const deletedEffects = isComplement ? getArrivalDeletedEffects(activePerson, allEffects) : [];
    const effectsForSignature = [...activeEffects, ...deletedEffects];
    const effectsSignature = effectsForSignature
      .map((effect) => {
        const movement = getEffectMovementLabel(activePerson, effect, isComplement ? fallbackMovements : null);
        const replacementValue = getEffectUnitValue(effect);
        return [
          String(effect.id || ""),
          String(effect.typeEffet || ""),
          String(getEffectDisplayDesignation(effect) || ""),
          String(effect.numeroIdentification || ""),
          String(effect.dateRemise || ""),
          String(replacementValue || ""),
          String(movement || ""),
        ].join("|");
      })
      .join("||");

    return [
      "arrival",
      String(activePerson.id || ""),
      String(mode),
      String(isComplement),
      String(activePerson.nom || ""),
      String(activePerson.prenom || ""),
      String(activePerson.fonction || ""),
      String(activePerson.typePersonnel || ""),
      String(activePerson.typeContrat || ""),
      String(activePerson.dateEntree || ""),
      String(activePerson.dateSortiePrevue || ""),
      String(representative?.id || ""),
      String(representative?.nom || ""),
      String(representative?.fonction || ""),
      String(getSignatureValue(activePerson, "arrival", "personnel") || ""),
      String(getSignatureValue(activePerson, "arrival", "representant") || ""),
      String(getSignatureValidationDate(activePerson, "arrival", "personnel") || ""),
      String(getSignatureValidationDate(activePerson, "arrival", "representant") || ""),
      String(activeEffects.length),
      String(deletedEffects.length),
      String(effectsForSignature.length),
      String(getPersonSiteLabel(activePerson) || ""),
      String(sortConfig.key || ""),
      String(sortConfig.dir || ""),
      String(isPdfMode),
      String(effectsSignature),
    ].join("|");
  }

  const representative = getRepresentativeInfo(activePerson, "exit");
  const effects = (activePerson.effetsConfies || []).filter((effect) => {
    const hasType = Boolean(normalizeText(effect?.typeEffet));
    const hasDesignation = Boolean(normalizeText(getEffectDisplayDesignation(effect)));
    const hasId = Boolean(normalizeText(effect?.numeroIdentification));
    const hasDateRemise = Boolean(String(effect?.dateRemise || "").trim());
    const hasDateRetour = Boolean(String(effect?.dateRetour || "").trim());
    const hasStatus = Boolean(normalizeText(getEffectStatus(activePerson, effect)));
    const hasAmount = getEffectReplacementCost(activePerson, effect) > 0;
    return hasType || hasDesignation || hasId || hasDateRemise || hasDateRetour || hasStatus || hasAmount;
  });
  const effectsSignature = effects
    .map((effect) => {
      const movement = getEffectMovementLabel(activePerson, effect);
      const replacementCost = getEffectReplacementCost(activePerson, effect);
      const billingStatus = getEffectBillingStatus(effect, replacementCost > 0);
      const status = String(getEffectStatus(activePerson, effect));
      return [
        String(effect.id || ""),
        String(effect.typeEffet || ""),
        String(getEffectDisplayDesignation(effect) || ""),
        String(effect.numeroIdentification || ""),
        String(effect.dateRemise || ""),
        String(effect.dateRetour || ""),
        String(status || ""),
        String(replacementCost || 0),
        String(billingStatus || ""),
        String(movement || ""),
      ].join("|");
    })
    .join("||");

  const chargeableEffects = effects.filter((effect) => isEffectChargeable(activePerson, effect));
  const chargeableValue = chargeableEffects.reduce((sum, effect) => sum + getEffectReplacementCost(activePerson, effect), 0);

  return [
    "exit",
    String(activePerson.id || ""),
    String(activePerson.nom || ""),
    String(activePerson.prenom || ""),
    String(activePerson.fonction || ""),
    String(activePerson.typePersonnel || ""),
    String(activePerson.typeContrat || ""),
    String(activePerson.dateEntree || ""),
    String(activePerson.dateSortiePrevue || ""),
    String(activePerson.dateSortieReelle || ""),
    String(representative?.id || ""),
    String(representative?.nom || ""),
    String(representative?.fonction || ""),
    String(getSignatureValue(activePerson, "exit", "personnel") || ""),
    String(getSignatureValue(activePerson, "exit", "representant") || ""),
    String(getSignatureValidationDate(activePerson, "exit", "personnel") || ""),
    String(getSignatureValidationDate(activePerson, "exit", "representant") || ""),
    String(effects.length),
    String(effects.filter((effect) => getEffectStatus(activePerson, effect) === "RESTITUE").length),
    String(chargeableEffects.length),
    String(getPersonSiteLabel(activePerson) || ""),
    String(sortConfig.key || ""),
    String(sortConfig.dir || ""),
    String(isPdfMode),
    String(effectsSignature),
    String(chargeableValue || 0),
  ].join("|");
}

function renderPage() {
  const page = document.body.dataset.page || "";
  const filters = state.filters || DEFAULT_FILTERS;
  let currentPersonId = getCurrentPersonId();
  const pageRenderSignatureParts = [
    "page-render",
    String(page),
    String(state.localMutationTick || 0),
    String(state.supabaseRevision || ""),
    String(state.isDirty ? "1" : "0"),
    String(state.saveButtonLatchedDirty ? "1" : "0"),
    String(state.currentUserRoleLabel || ""),
    String(state.currentSheetPersonId || ""),
    String(state.editingEffectId || ""),
    String(state.editingReferenceId || ""),
    String(state.editingReplacementCostKey || ""),
    String(state.editingRepresentativeId || ""),
    String(state.editingSimpleReference ? `${state.editingSimpleReference.listName || ""}/${state.editingSimpleReference.originalValue || ""}` : ""),
    String(currentPersonId || ""),
  ];

  if (page === "overview" || page === "global") {
    pageRenderSignatureParts.push(
      `filters|${String(filters.search || "")}|${String(filters.site || "")}|${String(filters.typePersonnel || "")}|${String(filters.typeContrat || "")}|${String(filters.statutDossier || "")}|${String(filters.statutObjet || "")}|${String(filters.typeEffet || "")}|${String(state.urgentMode ? "1" : "0")}`,
      `sort|${String(state.tableSorts?.overviewPersons?.key || "")}|${String(state.tableSorts?.overviewPersons?.dir || "")}|${String(state.tableSorts?.global?.key || "")}|${String(state.tableSorts?.global?.dir || "")}`
    );
  } else if (page === "person-sheet") {
    pageRenderSignatureParts.push(
      `person-sheet|${String(state.effectRowFlash?.personId || "")}|${String(state.effectRowFlash?.effectId || "")}|${String(state.effectRowFlash?.kind || "")}|${String(state.effectTableFlash?.personId || "")}|${String(state.effectTableFlash?.kind || "")}`
    );
  } else if (page === "arrival-document" || page === "exit-document") {
    const docType = page === "arrival-document" ? "arrival" : "exit";
    pageRenderSignatureParts.push(getDocumentViewStateSignature(docType, currentPersonId));
  } else if (page === "mobile-signature") {
    const signer = getCurrentMobileSignatureSigner();
    const docType = getCurrentMobileSignatureDocType();
    const request = getCurrentMobileSignatureRequest();
    pageRenderSignatureParts.push(`mobile|${signer}|${docType}|${String(request?.token || "")}|${String(request?.status || "")}|${String(request?.signer || "")}`);
  } else if (page === "documents-archives") {
    pageRenderSignatureParts.push(
      `archive|${String(state.tableSorts?.documentsArchives?.key || "")}|${String(state.tableSorts?.documentsArchives?.dir || "")}|${getArchiveFilterSignatureFromDom()}`
    );
  } else if (page === "reference-bases") {
    pageRenderSignatureParts.push(
      `references|${String(state.editingSimpleReference ? state.editingSimpleReference.listName || "" : "")}|${String(state.editingReferenceId || "")}|${String(state.editingReplacementCostKey || "")}`
    );
  }

  const nextPageRenderSignature = pageRenderSignatureParts.join("|");
  if (state.pageRenderSignature === nextPageRenderSignature) {
    return;
  }
  state.pageRenderSignature = nextPageRenderSignature;

  if (page !== "arrival-document" && page !== "exit-document") {
    stopMobileSignaturePolling();
  }
  const hasFilteredPersons = page === "overview" || page === "global";
  let persons = [];
  if (hasFilteredPersons) {
    const filteredPersonsSignature = [
      "filtered-persons",
      String(state.supabaseRevision || ""),
      String(state.localMutationTick || 0),
      String(state.latestDataEtag || ""),
      String(Array.isArray(state.data?.personnes) ? state.data.personnes.length : 0),
      String(state.urgentMode ? "1" : "0"),
      String(filters.search || ""),
      String(filters.site || ""),
      String(filters.typePersonnel || ""),
      String(filters.typeContrat || ""),
      String(filters.statutDossier || ""),
      String(filters.statutObjet || ""),
      String(filters.typeEffet || ""),
    ].join("|");
    if (state.filteredPersonsCache.key !== filteredPersonsSignature) {
      state.filteredPersonsCache = {
        key: filteredPersonsSignature,
        persons: getFilteredPersons(),
      };
    }
    persons = state.filteredPersonsCache.persons || [];
  }
  const personExists = (state.data?.personnes || []).some(
    (entry) => String(entry?.id || "") === String(currentPersonId || "")
  );

  if (page === "person-sheet" && currentPersonId && !personExists) {
    setCurrentPersonId("", "replace");
    currentPersonId = "";
  }

  if (page === "overview") {
    const didOverviewChange = renderOverview(persons);
    if (didOverviewChange) {
      updateSortableHeaders("overviewPersons");
    }
  }

  if (page === "global") {
    const didGlobalChange = renderGlobalTable(persons);
    if (didGlobalChange) {
      renderGlobalEffectsChart(persons);
    }
  }

  if (page === "documents-archives") {
    const didArchiveTableChange = renderDocumentsArchivePage();
    if (didArchiveTableChange) {
      bindEffectTableSorting("documentsArchives");
      updateSortableHeaders("documentsArchives");
    }
  }

  if (page === "person-sheet" || page === "arrival-document" || page === "exit-document") {
    try {
      renderPersonPicker();
    } catch (error) {
      console.error("Erreur affichage recherche personne", error);
      showDataStatus("FICHE CHARGEE - RECHERCHE PERSONNE INDISPONIBLE");
    }
  }

  if (page === "mobile-signature") {
    const didMobileSignatureChange = renderMobileSignaturePage();
    if (didMobileSignatureChange) {
      refreshDocumentSignatureCanvases(getCurrentMobileSignatureDocType(), getMobileSignatureTargetPerson());
    }
  }

  if (page === "person-sheet") {
    try {
      const didPersonSheetChange = renderPersonSheet(currentPersonId);
      if (didPersonSheetChange) {
        bindEffectTableSorting("sheetEffects");
        updateSortableHeaders("sheetEffects");
      }
    } catch (error) {
      console.error("Erreur affichage fiche personne", error);
      showDataStatus("ERREUR AFFICHAGE FICHE PERSONNE - VOIR CONSOLE");
    }
  }

  if (page === "arrival-document") {
    const nextArrivalViewSignature = getDocumentViewStateSignature("arrival", currentPersonId);
    if (state.documentViewRenderCache.arrival !== nextArrivalViewSignature) {
      state.documentViewRenderCache.arrival = nextArrivalViewSignature;
      const didRenderArrivalDocument = renderArrivalDocument(currentPersonId);
      if (didRenderArrivalDocument) {
        scheduleMobileSignatureRenderSync();
      }
    }
    refreshDocumentSignatureCanvases("arrival");
  }

  if (page === "exit-document") {
    const nextExitViewSignature = getDocumentViewStateSignature("exit", currentPersonId);
    if (state.documentViewRenderCache.exit !== nextExitViewSignature) {
      state.documentViewRenderCache.exit = nextExitViewSignature;
      const didRenderExitDocument = renderExitDocument(currentPersonId);
      if (didRenderExitDocument) {
        scheduleMobileSignatureRenderSync();
      }
    }
    refreshDocumentSignatureCanvases("exit");
  }

  if (page === "reference-bases") {
    const didRenderReferenceBases = renderReferenceBases();
    if (didRenderReferenceBases) {
      bindEffectTableSorting("referenceEffects");
      updateSortableHeaders("referenceEffects");
    }
  }

  updateDocumentPdfButtonsState();
  renderDirtyState();
}

function renderEffectsChart(nodeId, persons) {
  const node = document.getElementById(nodeId);
  if (!node) {
    return;
  }

  const filters = state.filters || DEFAULT_FILTERS;
  const effects = getAllEffects(persons)
    .filter(({ person, effect }) => {
      if (!effectMatchesSiteFilter(person, effect, filters.site)) {
        return false;
      }
      if (filters.typeEffet && normalizeText(effect?.typeEffet) !== filters.typeEffet) {
        return false;
      }
      if (filters.statutObjet && getEffectStatus(person, effect) !== filters.statutObjet) {
        return false;
      }
      return true;
    })
    .map(({ person, effect }) => ({ person, effect }));
  if (!effects.length) {
    node.innerHTML = '<div class="effects-chart__empty">AUCUNE DONNEE A AFFICHER</div>';
    return;
  }

  const counts = new Map();
  const totals = {
    actif: 0,
    nonRendu: 0,
    restitue: 0,
    perdu: 0,
    vole: 0,
    hs: 0,
  };
  const totalsCost = {
    actif: 0,
    nonRendu: 0,
    restitue: 0,
    perdu: 0,
    vole: 0,
    hs: 0,
  };
  let totalEntrustedCost = 0;
  let totalFacturable = 0;
  effects.forEach(({ person, effect }) => {
    const type = normalizeText(effect?.typeEffet) || "EFFET";
    if (!counts.has(type)) {
      counts.set(type, {
        total: 0,
        segments: {
          actif: 0,
          nonRendu: 0,
          restitue: 0,
          perdu: 0,
          vole: 0,
          hs: 0,
        },
      });
    }

    const row = counts.get(type);
    row.total += 1;
    const category = getEffectChartCategory(person, effect);
    row.segments[category] += 1;
    totals[category] += 1;
    totalEntrustedCost += getReplacementCostValue(effect?.typeEffet, "NON RENDU", effect?.designation || "");
    const replacementCost = getEffectReplacementCost(person, effect);
    totalsCost[category] += replacementCost;
    totalFacturable += replacementCost;
  });

  const configuredTypes = Array.from(
    new Set(
      [
        ...(state.data?.listes?.typesEffets || []),
        ...(state.data?.listes?.referencesEffets || []).map((reference) => reference?.typeEffet),
        ...(state.data?.listes?.coutsRemplacement || []).map((cost) => cost?.typeEffet),
      ]
        .map((typeEffet) => normalizeText(typeEffet))
        .filter(Boolean)
    )
  );
  configuredTypes.forEach((type) => {
    if (filters.typeEffet && type !== filters.typeEffet) {
      return;
    }
    if (!counts.has(type)) {
      counts.set(type, {
        total: 0,
        segments: {
          actif: 0,
          nonRendu: 0,
          restitue: 0,
          perdu: 0,
          vole: 0,
          hs: 0,
        },
      });
    }
  });

  const rows = Array.from(counts.entries())
    .sort((left, right) => {
      const typeOrder = left[0].localeCompare(right[0], "fr");
      if (typeOrder !== 0) {
        return typeOrder;
      }
      return left[1].total - right[1].total;
    });

  const maxValue = Math.max(...rows.map(([, row]) => row.total), 1);
  const summaryMarkup = `
    <div class="effects-chart__summary">
      <span class="effects-chart__summary-item">TOTAL EFFETS CONFIES <strong>${effects.length}</strong><span class="effects-chart__summary-sub">COUT <strong>${formatAmountWithEuro(totalEntrustedCost)}</strong></span></span>
      <span class="effects-chart__summary-item">TOTAL FACTURABLE <strong>${formatAmountWithEuro(totalFacturable)}</strong></span>
    </div>`;
  const legendMarkup = `
    <div class="effects-chart__legend">
      <span class="effects-chart__legend-item"><span class="effects-chart__legend-dot effects-chart__legend-dot--actif"></span>ACTIF <strong>${totals.actif}</strong><span class="effects-chart__legend-cost">${formatAmountWithEuro(totalsCost.actif)}</span></span>
      <span class="effects-chart__legend-item"><span class="effects-chart__legend-dot effects-chart__legend-dot--nonRendu"></span>NON RENDU <strong>${totals.nonRendu}</strong><span class="effects-chart__legend-cost">${formatAmountWithEuro(totalsCost.nonRendu)}</span></span>
      <span class="effects-chart__legend-item"><span class="effects-chart__legend-dot effects-chart__legend-dot--restitue"></span>RENDU <strong>${totals.restitue}</strong><span class="effects-chart__legend-cost">${formatAmountWithEuro(totalsCost.restitue)}</span></span>
      <span class="effects-chart__legend-item"><span class="effects-chart__legend-dot effects-chart__legend-dot--perdu"></span>PERDU <strong>${totals.perdu}</strong><span class="effects-chart__legend-cost">${formatAmountWithEuro(totalsCost.perdu)}</span></span>
      <span class="effects-chart__legend-item"><span class="effects-chart__legend-dot effects-chart__legend-dot--vole"></span>VOLE <strong>${totals.vole}</strong><span class="effects-chart__legend-cost">${formatAmountWithEuro(totalsCost.vole)}</span></span>
      <span class="effects-chart__legend-item"><span class="effects-chart__legend-dot effects-chart__legend-dot--hs"></span>HS <strong>${totals.hs}</strong><span class="effects-chart__legend-cost">${formatAmountWithEuro(totalsCost.hs)}</span></span>
    </div>`;
  const rowsMarkup = rows
    .map(([type, row]) => {
      const width = row.total > 0 ? Math.max(8, Math.round((row.total / maxValue) * 100)) : 0;
      const segmentMarkup = [
        ["actif", "ACTIF"],
        ["nonRendu", "NON RENDU"],
        ["restitue", "RENDU"],
        ["perdu", "PERDU"],
        ["vole", "VOLE"],
        ["hs", "HS"],
      ]
        .filter(([key]) => row.segments[key] > 0)
        .map(
          ([key, label]) =>
            `<span class="effects-chart__segment effects-chart__segment--${key}" style="width:${(row.segments[key] / row.total) * 100}%" title="${label} : ${row.segments[key]}"></span>`
        )
        .join("");
      return `<div class="effects-chart__row">
        <span class="effects-chart__label">${escapeHtml(type)}</span>
        <span class="effects-chart__track" aria-hidden="true">
          <span class="effects-chart__bar-group" style="width:${width}%">${segmentMarkup}</span>
        </span>
        <strong class="effects-chart__value">${row.total}</strong>
      </div>`;
    })
    .join("");
  node.innerHTML = `${summaryMarkup}${legendMarkup}${rowsMarkup}`;
}

function renderGlobalEffectsChart(persons) {
  renderEffectsChart("global-effects-chart", persons);
}

function effectMatchesSiteFilter(person, effect, site) {
  const normalizedSite = normalizeText(site);
  if (!normalizedSite) {
    return true;
  }

  const effectSite = normalizeText(effect?.siteReference || referenceSiteFromEffect(effect));
  if (effectSite) {
    return effectSite === ALL_SITES_VALUE || effectSite === normalizedSite;
  }

  return personHasSite(person, normalizedSite);
}

function getCurrentPerson() {
  const personId = getCurrentPersonId();
  return state.data?.personnes?.find((entry) => entry.id === personId) || null;
}

function sortDocumentsArchives() {
  if (!Array.isArray(state.data?.documentsArchives)) {
    return;
  }
  state.data.documentsArchives.sort((left, right) => {
    const leftLabel = `${String(left.dateDocument || "")} ${normalizeText(left.nom)} ${normalizeText(left.prenom)} ${normalizeText(left.typeDocument)}`;
    const rightLabel = `${String(right.dateDocument || "")} ${normalizeText(right.nom)} ${normalizeText(right.prenom)} ${normalizeText(right.typeDocument)}`;
    return rightLabel.localeCompare(leftLabel, "fr");
  });
}

function sanitizeFilePart(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Za-z0-9_-]+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "") || "document";
}

function getDocumentArchiveStoragePath(entry) {
  const folder = getArchiveDocTypeKey(entry?.typeDocument) === "SORTIE" ? "archives/pdf/sortie" : "archives/pdf/arrivee";
  const name = `${sanitizeFilePart(entry?.dateDocument || getTodayIsoDate())}_${sanitizeFilePart(entry?.nom)}_${sanitizeFilePart(entry?.prenom)}.pdf`;
  return `${folder}/${name}`;
}

function getDocumentArchiveSignatureStatus(entry) {
  const personId = String(entry?.personId || "").trim();
  const rawStatus = normalizeText(entry?.statutSignature) || "EN ATTENTE";
  const normalizedType = getArchiveDocTypeKey(entry?.typeDocument);
  const normalizedStatus = rawStatus === "SIGNEE" ? "SIGNE" : rawStatus;
  const person = (state.data?.personnes || []).find((currentPerson) =>
    String(currentPerson?.id || "").trim() === personId
  );
  if (!personId || !person || !normalizedType) {
    return normalizedStatus || "EN ATTENTE";
  }

  const signatureType = normalizedType === "SORTIE" ? "exit" : "arrival";
  if (normalizedStatus === "SIGNE") {
    return isDocumentFullySigned(person, signatureType) ? "SIGNE" : "EN ATTENTE DE SIGNATURE";
  }
  if (
    normalizedStatus === "ATTENTE DE GENERATION" ||
    normalizedStatus === "EN ATTENTE DE SIGNATURE"
  ) {
    return normalizedStatus;
  }
  const isSigned = isDocumentFullySigned(person, signatureType);
  return isSigned ? "SIGNE" : "EN ATTENTE DE SIGNATURE";
}

function getDocumentArchiveWorkflowStatus(entry, person) {
  const status = getDocumentArchiveSignatureStatus(entry);
  if (status === "SIGNE") {
    return "SIGNE";
  }
  if (status === "ATTENTE DE GENERATION" || status === "EN ATTENTE DE SIGNATURE") {
    return status;
  }
  const docType = getArchiveDocTypeKey(entry?.typeDocument) === "SORTIE" ? "exit" : "arrival";
  return isDocumentFullySigned(person, docType) ? "ATTENTE DE GENERATION" : "EN ATTENTE DE SIGNATURE";
}

function isHostedPdfDocumentPath(pathValue) {
  const value = String(pathValue || "").trim();
  if (!value) return false;
  const isLocalDocumentPdf =
    /(^|\/)document-(arrivee|sortie)\.html/i.test(value) ||
    /(^|\/)document-(arrivee|sortie)\.html\?(.*)?/i.test(value);
  if (!isLocalDocumentPdf) {
    return false;
  }
  if (!/^(?:https?:\/\/)/i.test(value)) {
    return true;
  }
  try {
    const parsed = new URL(value);
    if (parsed.origin !== window.location.origin) {
      return false;
    }
    return isLocalDocumentPdf;
  } catch (error) {
    return false;
  }
}

function getDocumentArchiveStatusCellMarkup(status) {
  const normalized = normalizeText(status);
  let iconClass = "status-icon-inline status-icon-inline--pending";

  if (normalized === "SIGNE") {
    iconClass = "status-icon-inline status-icon-inline--active";
  } else if (normalized === "ATTENTE DE GENERATION") {
    iconClass = "status-icon-inline status-icon-inline--pending";
  } else if (normalized === "EN ATTENTE DE SIGNATURE") {
    iconClass = "status-icon-inline status-icon-inline--warning";
  } else {
    iconClass = "status-icon-inline status-icon-inline--warning";
  }

  return `<span class="status-cell"><span class="${iconClass}" aria-hidden="true"></span><span>${escapeHtml(status || "")}</span></span>`;
}

function getLatestArchivePerPersonAndType(archives) {
  const map = new Map();
  (archives || []).forEach((entry) => {
    const personId = String(entry?.personId || "").trim();
    const normalizedType = getArchiveDocTypeKey(entry?.typeDocument);
    if (!personId || !normalizedType) {
      return;
    }
    const key = `${personId}|${normalizedType}`;
    const current = map.get(key);
    const currentMs = Date.parse(String(current?.dateArchivage || "")) || 0;
    const nextMs = Date.parse(String(entry?.dateArchivage || "")) || 0;
    if (!current || nextMs >= currentMs) {
      map.set(key, entry);
    }
  });
  return map;
}

function getDocumentTypeLabel(docType) {
  return getArchiveDocTypeKey(docType) === "SORTIE" ? "SORTIE" : "ARRIVEE";
}

function normalizeArchiveTypeLabel(value) {
  return getArchiveDocTypeKey(value) || normalizeText(value);
}

function getArchiveDocTypeKey(value) {
  const normalized = normalizeText(value);
  if (!normalized) {
    return "";
  }
  if (["ARRIVEE", "ARRIVAL", "ENTREE"].includes(normalized)) {
    return "ARRIVEE";
  }
  if (["SORTIE", "EXIT"].includes(normalized)) {
    return "SORTIE";
  }
  return "";
}

function isDocumentFullySigned(person, docType) {
  const hasPersonnelSignature =
    Boolean(getSignatureValue(person, docType, "personnel")) &&
    Boolean(getSignatureValidationDate(person, docType, "personnel"));
  const hasRepresentativeSignature =
    Boolean(getSignatureValue(person, docType, "representant")) &&
    Boolean(getSignatureValidationDate(person, docType, "representant"));
  return hasPersonnelSignature && hasRepresentativeSignature;
}

function getDocumentLatestSignatureTimestampMs(person, docType) {
  const personnelDate = String(getSignatureValidationDate(person, docType, "personnel") || "");
  const representantDate = String(getSignatureValidationDate(person, docType, "representant") || "");
  const personnelMs = Date.parse(personnelDate) || 0;
  const representantMs = Date.parse(representantDate) || 0;
  return Math.max(personnelMs, representantMs);
}

function getDocumentSignatureProgress(person, docType) {
  const hasPersonnelSignature =
    Boolean(getSignatureValue(person, docType, "personnel")) &&
    Boolean(getSignatureValidationDate(person, docType, "personnel"));
  const hasRepresentativeSignature =
    Boolean(getSignatureValue(person, docType, "representant")) &&
    Boolean(getSignatureValidationDate(person, docType, "representant"));
  const count = (hasPersonnelSignature ? 1 : 0) + (hasRepresentativeSignature ? 1 : 0);
  return {
    hasPersonnelSignature,
    hasRepresentativeSignature,
    count,
    isFullySigned: hasPersonnelSignature && hasRepresentativeSignature,
  };
}

function getSignedPdfStateForDocumentType(person, docType) {
  const normalizedDocType = normalizeText(docType);
  if (!state.data || !person) {
    return "MISSING";
  }
  const normalizedFingerprintType = normalizedDocType === "SORTIE" || normalizedDocType === "EXIT" ? "exit" : "arrival";
  const typeLabel = getDocumentTypeLabel(normalizedFingerprintType);
  const fingerprint = getDocumentFingerprint(person, normalizedFingerprintType);
  if (!typeLabel) {
    return "MISSING";
  }
  const archives = (state.data.documentsArchives || []).filter((entry) => {
    if (String(entry?.personId || "") !== String(person.id || "")) {
      return false;
    }
    if (normalizeText(entry?.typeDocument) !== typeLabel) {
      return false;
    }
    if (!String(entry?.pdfPath || "").trim()) {
      return false;
    }
    return true;
  });

  if (!archives.length) {
    return "MISSING";
  }

  const hasFingerprintMatch = archives.some(
    (entry) => String(entry?.fingerprint || "").trim() && String(entry.fingerprint || "").trim() === fingerprint
  );
  if (hasFingerprintMatch) {
    return "UP_TO_DATE";
  }

  const latestSignatureMs = getDocumentLatestSignatureTimestampMs(person, normalizedFingerprintType);
  const hasSignedArchiveAfterSignature = archives.some((entry) => {
    if (normalizeText(entry?.statutSignature) !== "SIGNE") {
      return false;
    }
    if (!latestSignatureMs) {
      return true;
    }
    const archivedMs = Date.parse(String(entry?.dateArchivage || "")) || 0;
    return archivedMs >= latestSignatureMs;
  });
  if (hasSignedArchiveAfterSignature) {
    return "UP_TO_DATE";
  }

  const hasAnySignedArchive = archives.some((entry) => normalizeText(entry?.statutSignature) === "SIGNE");
  return hasAnySignedArchive ? "OUTDATED" : "MISSING";
}

function hasCurrentSignedPdfForDocumentType(person, docType) {
  return getSignedPdfStateForDocumentType(person, docType) === "UP_TO_DATE";
}

function getSignaturePdfPendingAlerts(person) {
  const alerts = [];
  const currentEffectsCount = getCurrentAssignedEffects(person).length;
  const documentContexts = [
    { label: "d'entrée", key: "arrival", active: currentEffectsCount > 0 },
    { label: "de sortie", key: "exit", active: getDossierStatus(person) === "SORTI" },
  ];

  documentContexts.forEach((context) => {
    if (!context.active) return;
    const signatureProgress = getDocumentSignatureProgress(person, context.key);
    if (signatureProgress.count === 0) {
      alerts.push({
        docType: context.key === "exit" ? "SORTIE" : "ARRIVEE",
        message: `Le document ${context.label} n'est pas encore signe.`,
        type: "signaturePdf",
      });
      return;
    }
    if (!signatureProgress.isFullySigned) {
      alerts.push({
        docType: context.key === "exit" ? "SORTIE" : "ARRIVEE",
        message: `Le document ${context.label} doit encore etre signe par l'autre partie.`,
        type: "signaturePdf",
      });
      return;
    }
    const pdfState = getSignedPdfStateForDocumentType(person, context.key);
    if (pdfState === "MISSING") {
      alerts.push({
        docType: context.key === "exit" ? "SORTIE" : "ARRIVEE",
        message: `Les deux signatures sont presentes. Le PDF ${context.label} doit etre genere.`,
        type: "signaturePdf",
      });
      return;
    }
    if (pdfState === "OUTDATED") {
      alerts.push({
        docType: context.key === "exit" ? "SORTIE" : "ARRIVEE",
        message: `Les deux signatures sont presentes. Le PDF ${context.label} doit etre mis a jour.`,
        type: "signaturePdf",
      });
    }
  });

  return alerts.map((item) => ({
    id: person.id,
    nom: person.nom,
    prenom: person.prenom,
    message: item.message,
    type: item.type || "dateSortiePrevue",
    docType: item.docType,
  }));
}

function getPersonAlerts(person) {
  const alerts = [];
  const exitAlertMeta = getOverdueExitAlertMeta(person);
  if (exitAlertMeta?.message) {
    alerts.push({
      id: person?.id,
      nom: person?.nom,
      prenom: person?.prenom,
      message: exitAlertMeta.message,
      type: exitAlertMeta.type || "dateSortiePrevue",
    });
  }
  return alerts.concat(getSignaturePdfPendingAlerts(person));
}

function validateFinalSignatureBeforeSave(person, docType) {
  if (!person) {
    return { ok: false, message: "AUCUNE PERSONNE SELECTIONNEE" };
  }
  if (normalizeText(docType) !== "EXIT") {
    return { ok: true, message: "" };
  }

  const effects = person.effetsConfies || [];
  const hasInvalidRestitution = effects.some(
    (effect) =>
      normalizeText(getEffectStatus(person, effect)) === "RESTITUE" &&
      !normalizeDateString(effect?.dateRetour || "")
  );
  if (hasInvalidRestitution) {
    return {
      ok: false,
      message: "VERROU SIGNATURE: AU MOINS UN EFFET RESTITUE N'A PAS DE DATE DE RETOUR",
    };
  }

  const nonRendus = effects.filter((effect) => normalizeText(getEffectStatus(person, effect)) === "NON RENDU");
  const totalFacturable = effects
    .filter((effect) => isEffectChargeable(person, effect))
    .reduce((sum, effect) => sum + getEffectReplacementCost(person, effect), 0);

  if (nonRendus.length > 0 && totalFacturable <= 0) {
    return {
      ok: false,
      message: "VERROU SIGNATURE: DES EFFETS NON RENDUS SONT PRESENTS MAIS LE TOTAL FACTURABLE EST A ZERO",
    };
  }

  return { ok: true, message: "" };
}

function getDocumentArchiveDate(person, docType) {
  return normalizeText(docType) === "EXIT"
    ? person?.dateSortieReelle || person?.dateSortiePrevue || getTodayIsoDate()
    : person?.dateEntree || getTodayIsoDate();
}

function getDocumentArchiveEntryMode(entry) {
  return normalizeText(entry?.documentMode || "STANDARD");
}

function getDocumentArchiveVersionLabel(entry) {
  if (getDocumentArchiveEntryMode(entry) === "COMPLEMENTAIRE") {
    return "AVENANT";
  }

  // Strict rule: any change on ARRIVEE document makes it an AVENANT.
  if (normalizeText(entry?.typeDocument) === "ARRIVEE") {
    const personId = String(entry?.personId || "");
    const person = (state.data?.personnes || []).find((currentPerson) => String(currentPerson?.id || "") === personId);
    if (!person) {
      return "INITIAL";
    }
    const currentFingerprint = getDocumentFingerprint(person, "arrival");
    const archivedFingerprint = String(entry?.fingerprint || "");
    if (currentFingerprint && archivedFingerprint && currentFingerprint !== archivedFingerprint) {
      return "AVENANT";
    }
  }

  return "INITIAL";
}

function getArchiveDisplayedTotalEffets(entry) {
  const archivedCount = Number(entry?.totalEffets || 0);
  const personId = String(entry?.personId || "");
  const person = (state.data?.personnes || []).find((currentPerson) => String(currentPerson?.id || "") === personId);
  const currentCount = Array.isArray(person?.effetsConfies) ? person.effetsConfies.length : archivedCount;
  return currentCount;
}

function getDocumentArchiveMode(person, docType) {
  if (!state.data || !person || normalizeText(docType) !== "ARRIVAL") {
    return "STANDARD";
  }
  const signedArrivalArchives = (state.data.documentsArchives || []).filter(
    (entry) =>
      String(entry.personId || "") === String(person.id || "") &&
      normalizeText(entry.typeDocument) === "ARRIVEE" &&
      getDocumentArchiveSignatureStatus(entry) === "SIGNE"
  );
  if (!signedArrivalArchives.length) {
    return "STANDARD";
  }
  const fingerprint = getDocumentFingerprint(person, docType);
  const matchingArchive = signedArrivalArchives.find((entry) => String(entry.fingerprint || "") === fingerprint);
  if (matchingArchive) {
    return getDocumentArchiveEntryMode(matchingArchive);
  }
  return "COMPLEMENTAIRE";
}

function getEffectMovementKey(effect) {
  const explicitId = String(effect?.id || "").trim();
  if (explicitId) {
    return `ID:${explicitId}`;
  }
  return [
    normalizeText(effect?.typeEffet),
    normalizeText(effect?.siteReference || referenceSiteFromEffect(effect)),
    normalizeText(getEffectDisplayDesignation(effect)),
    normalizeText(effect?.numeroIdentification),
    String(effect?.dateRemise || ""),
  ].join("|");
}

function getEffectStableKey(effect) {
  const explicitId = String(effect?.id || "").trim();
  if (explicitId) {
    return `ID:${explicitId}`;
  }
  return [
    normalizeText(effect?.typeEffet),
    normalizeText(effect?.siteReference || referenceSiteFromEffect(effect)),
    normalizeText(getEffectDisplayDesignation(effect)),
    normalizeText(effect?.numeroIdentification),
  ].join("|");
}

function getEffectComparableSignature(person, effect) {
  return JSON.stringify({
    typeEffet: normalizeText(effect?.typeEffet),
    site: normalizeText(effect?.siteReference || referenceSiteFromEffect(effect)),
    designation: normalizeText(getEffectDisplayDesignation(effect)),
    numeroIdentification: normalizeText(effect?.numeroIdentification),
    vehiculeImmatriculation: normalizeText(effect?.vehiculeImmatriculation),
    dateRemise: String(effect?.dateRemise || ""),
    dateRetour: String(effect?.dateRetour || ""),
    statut: normalizeText(getEffectStatus(person, effect)),
    cause: normalizeText(getEffectReplacementCause(person, effect)),
    dateRemplacement: String(effect?.dateRemplacement || ""),
    commentaire: normalizeText(effect?.commentaire),
    cout: normalizeAmount(getEffectReplacementCost(person, effect)),
  });
}

function getMovementBadgeVariant(movement) {
  const normalized = normalizeText(movement);
  if (normalized === "RENDU") return "retour";
  if (normalized === "PERDU") return "perdu";
  if (normalized === "DETRUIT") return "detruit";
  if (normalized === "VOLE") return "vole";
  if (normalized === "HS") return "hs";
  if (normalized === "NON RENDU") return "non-rendu";
  if (normalized === "SUPPRIME") return "supprime";
  if (normalized === "MODIFIE") return "modifie";
  return "ajout";
}

function getMovementRowVariant(movement) {
  const normalized = normalizeText(movement);
  if (normalized === "RENDU") return "returned";
  if (normalized === "PERDU") return "lost";
  if (normalized === "DETRUIT") return "detruit";
  if (normalized === "VOLE") return "stolen";
  if (normalized === "HS") return "hs";
  if (normalized === "SUPPRIME") return "removed";
  if (normalized === "MODIFIE") return "updated";
  return "added";
}

function getArrivalDeletedEffects(person, currentEffects) {
  const latestSignedArrival = getLatestSignedArrivalArchiveForPerson(person?.id);
  if (!latestSignedArrival?.fingerprint) {
    return [];
  }

  let baselineEffects = [];
  try {
    const payload = JSON.parse(String(latestSignedArrival.fingerprint || ""));
    baselineEffects = Array.isArray(payload?.effects) ? payload.effects : [];
  } catch (error) {
    baselineEffects = [];
  }
  if (!baselineEffects.length) {
    return [];
  }

  const currentById = new Set(
    (currentEffects || [])
      .map((effect) => String(effect?.id || "").trim())
      .filter(Boolean)
  );
  const currentByStable = new Set(
    (currentEffects || [])
      .map((effect) => getEffectStableKey(effect))
      .filter(Boolean)
  );

  return baselineEffects
    .filter((baselineEffect) => {
      const baselineId = String(baselineEffect?.id || "").trim();
      if (baselineId && currentById.has(baselineId)) {
        return false;
      }
      const baselineStable = getEffectStableKey(baselineEffect);
      if (baselineStable && currentByStable.has(baselineStable)) {
        return false;
      }
      return true;
    })
    .map((baselineEffect, index) => ({
      ...baselineEffect,
      id: `DEL-${person?.id || "P"}-${index}-${String(baselineEffect?.id || "").trim() || getEffectStableKey(baselineEffect)}`,
      __movementOverride: "SUPPRIME",
      __archivedDeleted: true,
    }));
}

function normalizeDateString(value) {
  const raw = String(value || "").trim();
  if (!raw) {
    return "";
  }
  const isoMatch = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) {
    return `${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}`;
  }
  const frMatch = raw.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (frMatch) {
    return `${frMatch[3]}-${frMatch[2]}-${frMatch[1]}`;
  }
  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) {
    return "";
  }
  const year = String(parsed.getFullYear());
  const month = String(parsed.getMonth() + 1).padStart(2, "0");
  const day = String(parsed.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function isStrictIsoCalendarDate(value) {
  const raw = String(value || "").trim();
  if (!raw) {
    return true;
  }
  const match = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) {
    return false;
  }
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) {
    return false;
  }
  const date = new Date(Date.UTC(year, month - 1, day));
  return (
    date.getUTCFullYear() === year &&
    date.getUTCMonth() + 1 === month &&
    date.getUTCDate() === day
  );
}

function validateDateFieldFormat(value, label) {
  const raw = String(value || "").trim();
  if (!raw) {
    return { ok: true, message: "" };
  }
  if (!isStrictIsoCalendarDate(raw)) {
    return {
      ok: false,
      message: `${label} INVALIDE (FORMAT ATTENDU : JJ/MM/AAAA)`,
    };
  }
  const year = Number(raw.slice(0, 4));
  const currentYear = new Date().getFullYear();
  if (!Number.isFinite(year) || year < 1900 || year > currentYear + 20) {
    return {
      ok: false,
      message: `${label} INVALIDE (ANNEE INCOHERENTE)`,
    };
  }
  return { ok: true, message: "" };
}

function getLatestSignedArrivalArchiveForPerson(personId) {
  const entries = (state.data?.documentsArchives || [])
    .filter(
      (entry) =>
        String(entry?.personId || "") === String(personId || "") &&
        normalizeText(entry?.typeDocument) === "ARRIVEE" &&
        getDocumentArchiveSignatureStatus(entry) === "SIGNE"
    )
    .slice()
    .sort((left, right) => String(right?.dateArchivage || "").localeCompare(String(left?.dateArchivage || "")));
  return entries[0] || null;
}

function getArrivalComplementMovementMap(person, effects) {
  const movements = new Map();
  if (!person || !Array.isArray(effects) || !effects.length) {
    return movements;
  }

  const latestSignedArrival = getLatestSignedArrivalArchiveForPerson(person.id);
  let baselineKeys = new Set();
  let baselineByStableKey = new Map();
  const baselineArchivedAt = normalizeDateString(latestSignedArrival?.dateArchivage || "");
  if (latestSignedArrival?.fingerprint) {
    try {
      const payload = JSON.parse(String(latestSignedArrival.fingerprint || ""));
      const baselineEffects = Array.isArray(payload?.effects) ? payload.effects : [];
      baselineKeys = new Set(
        baselineEffects
          .map((effect) => getEffectMovementKey(effect))
          .filter(Boolean)
      );
      baselineByStableKey = new Map(
        baselineEffects
          .map((effect) => [getEffectStableKey(effect), effect])
          .filter(([key]) => Boolean(key))
      );
    } catch (error) {
      baselineKeys = new Set();
      baselineByStableKey = new Map();
    }
  }

  effects.forEach((effect) => {
    const key = getEffectMovementKey(effect);
    if (!key) {
      return;
    }
    if (String(effect?.dateRetour || "")) {
      movements.set(key, "RENDU");
      return;
    }
    const effectStatus = normalizeText(getEffectStatus(person, effect));
    const effectCause = normalizeText(getEffectReplacementCause(person, effect));
    if (effectStatus === "HS") {
      movements.set(key, "HS");
      return;
    }
    if (effectCause === "VOL") {
      movements.set(key, "VOLE");
      return;
    }
    if (effectStatus === "PERDU" || effectCause === "PERTE") {
      movements.set(key, "PERDU");
      return;
    }
    if (baselineKeys.size && !baselineKeys.has(key)) {
      movements.set(key, "AJOUTE");
      return;
    }
    const stableKey = getEffectStableKey(effect);
    const baselineEffect = stableKey ? baselineByStableKey.get(stableKey) : null;
    if (baselineEffect) {
      const baselineSignature = getEffectComparableSignature(
        person,
        {
          ...effect,
          ...baselineEffect,
          siteReference: baselineEffect?.site || baselineEffect?.siteReference || effect?.siteReference || "",
        }
      );
      const currentSignature = getEffectComparableSignature(person, effect);
      if (baselineSignature !== currentSignature) {
        movements.set(key, "MODIFIE");
        return;
      }
    }
    if (!baselineKeys.size && baselineArchivedAt) {
      const remiseDate = normalizeDateString(effect?.dateRemise || "");
      if (remiseDate && remiseDate >= baselineArchivedAt) {
        movements.set(key, "AJOUTE");
      }
      return;
    }
    if (!baselineKeys.size && !baselineArchivedAt) {
      movements.set(key, "AJOUTE");
    }
  });

  return movements;
}

function getEffectMovementLabel(person, effect, movementMap = null) {
  const forcedMovement = normalizeText(effect?.__movementOverride || "");
  if (forcedMovement) {
    return forcedMovement;
  }
  const key = getEffectMovementKey(effect);
  if (movementMap instanceof Map && key) {
    const fromMap = String(movementMap.get(key) || "").trim();
    if (fromMap) {
      return fromMap;
    }
  }

  if (typeof deriveEffectState === "function") {
    return deriveEffectState(person, effect).movement;
  }
  if (String(effect?.dateRetour || "").trim()) return "RENDU";
  const effectStatus = normalizeText(getEffectStatus(person, effect));
  const effectCause = normalizeText(getEffectReplacementCause(person, effect));
  if (effectStatus === "DETRUIT") return "DETRUIT";
  if (effectStatus === "VOL") return "VOLE";
  if (effectStatus === "HS") return "HS";
  if (effectCause === "VOL") return "VOLE";
  if (effectStatus === "PERDU" || effectCause === "PERTE") return "PERDU";
  if (effectStatus === "NON RENDU") return "NON RENDU";
  return "";
}

function getDocumentFingerprint(person, docType) {
  if (!person) {
    return "";
  }
  const normalizedDocType = normalizeText(docType);
  const bucket = normalizedDocType === "EXIT" ? "exit" : "arrival";
  const effects = (person.effetsConfies || []).map((effect) => {
    const baseEffect = {
      id: String(effect.id || ""),
      typeEffet: normalizeText(effect.typeEffet),
      site: normalizeText(effect.siteReference || referenceSiteFromEffect(effect)),
      designation: normalizeText(getEffectDisplayDesignation(effect)),
      numeroIdentification: normalizeText(effect.numeroIdentification),
      vehiculeImmatriculation: normalizeText(effect.vehiculeImmatriculation),
      dateRemise: String(effect.dateRemise || ""),
      commentaire: normalizeText(effect.commentaire),
    };

    if (normalizedDocType === "ARRIVAL") {
      return {
        ...baseEffect,
        dateRetour: String(effect.dateRetour || ""),
        statut: normalizeText(getEffectStatus(person, effect)),
        mouvement: normalizeText(getEffectMovementLabel(person, effect)),
        cause: normalizeText(getEffectReplacementCause(person, effect)),
        dateRemplacement: String(effect.dateRemplacement || ""),
        cout: normalizeAmount(getEffectUnitValue(effect)),
      };
    }

    return {
      ...baseEffect,
      dateRetour: String(effect.dateRetour || ""),
      statut: normalizeText(getEffectStatus(person, effect)),
      cause: normalizeText(getEffectReplacementCause(person, effect)),
      dateRemplacement: String(effect.dateRemplacement || ""),
      cout: normalizeAmount(getEffectReplacementCost(person, effect)),
    };
  });

  return JSON.stringify({
    layoutVersion: PDF_LAYOUT_VERSION,
    docType: normalizedDocType,
    personId: String(person.id || ""),
    nom: normalizeText(person.nom),
    prenom: normalizeText(person.prenom),
    fonction: normalizeFunctionLabel(person.fonction),
    typePersonnel: normalizeText(person.typePersonnel),
    typeContrat: normalizeText(person.typeContrat),
    sites: getPersonSites(person),
    dateEntree: String(person.dateEntree || ""),
    dateSortiePrevue: String(person.dateSortiePrevue || ""),
    dateSortieReelle: normalizedDocType === "EXIT" ? String(person.dateSortieReelle || "") : "",
    representant: getRepresentativeInfo(person, bucket),
    signatures: {
      personnel: Boolean(getSignatureValue(person, bucket, "personnel")),
      representant: Boolean(getSignatureValue(person, bucket, "representant")),
      personnelDate: String(getSignatureValidationDate(person, bucket, "personnel") || ""),
      representantDate: String(getSignatureValidationDate(person, bucket, "representant") || ""),
    },
    effects,
  });
}

function getDocumentArchiveOpenPath(entry) {
  const raw = String(entry?.pdfPath || "").trim();
  if (!raw) {
    return "";
  }
  if (raw.startsWith("//")) {
    return "";
  }
  if (/^[a-z][a-z0-9+.-]*:/i.test(raw) && !/^https?:\/\//i.test(raw)) {
    return "";
  }
  if (/^https?:\/\//i.test(raw)) {
    if (isHostedPdfDocumentPath(raw)) {
      try {
        return resolveHostedRelativeUrl(raw).href;
      } catch {
        return "";
      }
    }
    return isSafeArchiveHttpUrl(raw) ? normalizeHttpUrl(raw) : "";
  }
  const storageRef = parseStorageSchemePath(raw);
  if (storageRef) {
    return getSupabaseStoragePublicUrl(storageRef.bucket, storageRef.objectPath) || "";
  }
  if (!isSafeArchiveRelativePath(raw)) {
    return "";
  }
  return raw.replace(/^\/+/, "");
}

function getCanonicalArchivePdfRemoteUrl(entry) {
  if (getDataBackendMode() === "LOCAL_API") {
    return "";
  }
  if (normalizePdfQualityStatus(entry?.pdfQualityStatus) === "LEGACY_ONLY") {
    return "";
  }
  const personId = sanitizeFilePart(String(entry?.personId || ""));
  const docType = getArchiveDocTypeKey(entry?.typeDocument);
  if (!personId || !docType) {
    return "";
  }
  const bucket = getStoragePdfBucketName();
  if (!bucket) {
    return "";
  }
  const folder = docType === "SORTIE" ? "sortie" : "arrivee";
  return getSupabaseStoragePublicUrl(bucket, `${folder}/${personId}/COURANT.pdf`) || "";
}

function getArchiveFileNameFromPath(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  const cleaned = raw.split("?")[0].split("#")[0];
  const parts = cleaned.split(/[\\/]/).filter(Boolean);
  const last = parts.length ? parts[parts.length - 1] : "";
  return /\.pdf$/i.test(last) ? last : "";
}

function resolveArchivePdfLocations(pdfPath, existing = {}) {
  const raw = String(pdfPath || "").trim();
  const existingLocalPath = String(existing.localPath || "").trim();
  const existingOpenLocalUrl = String(existing.openLocalUrl || "").trim();
  const existingStoragePath = String(existing.storagePath || "").trim();
  const existingPublicUrl = String(existing.publicUrl || "").trim();
  const existingOpenRemoteUrl = String(existing.openRemoteUrl || "").trim();

  let localPath = existingLocalPath;
  let openLocalUrl = existingOpenLocalUrl;
  let storagePath = existingStoragePath;
  let publicUrl = existingPublicUrl;
  let openRemoteUrl = existingOpenRemoteUrl;

  const storageRef = parseStorageSchemePath(raw);
  if (storageRef) {
    storagePath = `storage://${storageRef.bucket}/${storageRef.objectPath}`;
    const resolvedPublicUrl = getSupabaseStoragePublicUrl(storageRef.bucket, storageRef.objectPath) || "";
    if (resolvedPublicUrl) {
      publicUrl = resolvedPublicUrl;
      openRemoteUrl = resolvedPublicUrl;
    }
  } else if (/^https?:\/\//i.test(raw) && (isSafeArchiveHttpUrl(raw) || isHostedPdfDocumentPath(raw))) {
    const safeUrl = isHostedPdfDocumentPath(raw)
      ? (() => {
          try {
            return resolveHostedRelativeUrl(raw).href;
          } catch {
            return "";
          }
        })()
      : normalizeHttpUrl(raw) || "";
    if (safeUrl) {
      publicUrl = isHostedPdfDocumentPath(raw) ? publicUrl : safeUrl;
      openRemoteUrl = safeUrl;
    }
  } else if (isHostedPdfDocumentPath(raw)) {
    openRemoteUrl = (() => {
      try {
        return resolveHostedRelativeUrl(raw).href;
      } catch {
        return "";
      }
    })();
  } else if (/\.pdf(?:$|\?)/i.test(raw) && isSafeArchiveRelativePath(raw)) {
    localPath = raw.replace(/^\/+/, "").split("?")[0];
    if (localPath.toLowerCase().startsWith("data/pdf/")) {
      const apiPath = localPath.slice("data/pdf/".length);
      openLocalUrl = `/api/pdf-file?path=${encodeURIComponent(apiPath)}`;
    }
  }

  const filename =
    getArchiveFileNameFromPath(localPath) ||
    getArchiveFileNameFromPath(storagePath) ||
    getArchiveFileNameFromPath(publicUrl) ||
    getArchiveFileNameFromPath(raw) ||
    String(existing.filename || "").trim();

  return {
    filename,
    localPath,
    openLocalUrl,
    storagePath,
    publicUrl,
    openRemoteUrl,
  };
}

function getArchivePreferredOpenPath(entry) {
  return resolveArchiveOpenTarget(entry).url;
}

function normalizePdfQualityStatus(value) {
  const normalized = normalizeText(value);
  if (normalized === "VALID") return "VALID";
  if (normalized === "LEGACY_ONLY") return "LEGACY_ONLY";
  if (normalized === "MISSING_ACTIVE_FILE") return "INVALID_MISSING_FILE";
  if (normalized.startsWith("INVALID_LOGIN")) return "INVALID_LOGIN";
  if (normalized.startsWith("INVALID_TOO_SMALL")) return "INVALID_TOO_SMALL";
  if (normalized.startsWith("INVALID_BLANK")) return "INVALID_BLANK";
  if (normalized.startsWith("INVALID")) return "INVALID_BLANK";
  return "UNKNOWN";
}

function resolveArchiveOpenTarget(entry) {
  const localUrl = String(entry?.openLocalUrl || "").trim();
  const canonicalRemoteUrl = getCanonicalArchivePdfRemoteUrl(entry);
  const remoteUrl = String(entry?.openRemoteUrl || entry?.publicUrl || "").trim();
  const legacyUrl = getDocumentArchiveOpenPath(entry);
  const pdfQualityStatus = normalizePdfQualityStatus(entry?.pdfQualityStatus);
  const backendMode = getDataBackendMode();
  if (pdfQualityStatus.startsWith("INVALID")) {
    return { url: legacyUrl || "", source: legacyUrl ? "legacy" : "", qualityIssue: pdfQualityStatus };
  }
  if (pdfQualityStatus === "UNKNOWN" && legacyUrl) {
    return { url: legacyUrl, source: "legacy", qualityIssue: "UNKNOWN" };
  }
  if (backendMode === "LOCAL_API") {
    if (localUrl) return { url: localUrl, source: "local", qualityIssue: pdfQualityStatus };
    if (remoteUrl) return { url: remoteUrl, source: "remote", qualityIssue: pdfQualityStatus };
    return { url: legacyUrl || "", source: legacyUrl ? "legacy" : "", qualityIssue: pdfQualityStatus };
  }
  if (canonicalRemoteUrl) return { url: canonicalRemoteUrl, source: "remote-current", qualityIssue: pdfQualityStatus };
  if (remoteUrl) return { url: remoteUrl, source: "remote", qualityIssue: pdfQualityStatus };
  if (legacyUrl) return { url: legacyUrl, source: "legacy", qualityIssue: pdfQualityStatus };
  if (localUrl) return { url: localUrl, source: "local", qualityIssue: pdfQualityStatus };
  return { url: "", source: "", qualityIssue: pdfQualityStatus };
}

function handleArchiveOpenClick(archive, fallbackEntry = null) {
  const resolved = resolveArchiveOpenTarget(archive || fallbackEntry || {});
  const pdfUrl =
    normalizeDirectPdfOpenUrl(resolved.url) ||
    normalizeDirectPdfOpenUrl(fallbackEntry?.directOpenUrl || "");
  if (!pdfUrl) {
    return false;
  }
  window.open(pdfUrl, "_blank", "noopener");
  return true;
}

function getArchiveEntrySites(entry) {
  return normalizeSites(String(entry?.sites || "").split("/").map((value) => value.trim()));
}

function getLatestSignedArchivesPerPersonAndType(entries) {
  const latestByKey = new Map();
  (entries || []).forEach((entry) => {
    if (getDocumentArchiveSignatureStatus(entry) !== "SIGNE") {
      return;
    }
    const personKey = String(entry?.personId || "").trim() || `${normalizeText(entry?.nom || "")}|${normalizeText(entry?.prenom || "")}`;
    const typeKey = normalizeText(entry?.typeDocument || "");
    const key = `${personKey}|${typeKey}`;
    const currentMs = Date.parse(String(entry?.dateArchivage || "")) || 0;
    const existing = latestByKey.get(key);
    if (!existing) {
      latestByKey.set(key, entry);
      return;
    }
    const existingMs = Date.parse(String(existing?.dateArchivage || "")) || 0;
    if (currentMs >= existingMs) {
      latestByKey.set(key, entry);
    }
  });
  return Array.from(latestByKey.values());
}

function archiveEntryMatchesSite(entry, site) {
  const normalizedSite = normalizeText(site);
  if (!normalizedSite) {
    return true;
  }
  const sites = getArchiveEntrySites(entry);
  return sites.includes(ALL_SITES_VALUE) || sites.includes(normalizedSite);
}

function buildDocumentArchiveEntry(person, docType, pdfPath, metadataPath, archiveMode = "STANDARD", existingEntry = null, archiveDetails = null) {
  const effects = person?.effetsConfies || [];
  const documentMode = normalizeText(archiveMode || "STANDARD");
  const baseId = `DOC-${getDocumentTypeLabel(docType)}-${person?.id || ""}`;
  const locations = resolveArchivePdfLocations(pdfPath, existingEntry || {});
  const nextPdfQualityStatus =
    archiveDetails && !Object.prototype.hasOwnProperty.call(archiveDetails, "pdfQualityStatus")
      ? "VALID"
      : archiveDetails?.pdfQualityStatus || existingEntry?.pdfQualityStatus || "VALID";
  return {
    id: baseId,
    personId: String(person?.id || ""),
    nom: String(person?.nom || ""),
    prenom: String(person?.prenom || ""),
    typeDocument: getDocumentTypeLabel(docType),
    documentMode,
    dateDocument: getDocumentArchiveDate(person, docType),
    sites: getPersonSiteLabel(person),
    typePersonnel: String(person?.typePersonnel || ""),
    typeContrat: String(person?.typeContrat || ""),
    statutSignature: isDocumentFullySigned(person, docType) ? "SIGNE" : "EN ATTENTE",
    totalEffets: effects.length,
    totalFacturable:
      normalizeText(docType) === "EXIT"
        ? effects.reduce((sum, effect) => sum + getEffectReplacementCost(person, effect), 0)
        : 0,
    fingerprint: getDocumentFingerprint(person, docType),
    pdfPath: String(pdfPath || ""),
    filename: String(archiveDetails?.filename || locations.filename || ""),
    localPath: String(archiveDetails?.localPath || locations.localPath || ""),
    openLocalUrl: String(archiveDetails?.openLocalUrl || locations.openLocalUrl || ""),
    storagePath: String(archiveDetails?.storagePath || locations.storagePath || ""),
    publicUrl: String(archiveDetails?.publicUrl || locations.publicUrl || ""),
    openRemoteUrl: String(archiveDetails?.openRemoteUrl || locations.openRemoteUrl || ""),
    storageStatus: String(archiveDetails?.storageStatus || existingEntry?.storageStatus || ""),
    pdfQualityStatus: normalizePdfQualityStatus(nextPdfQualityStatus),
    metadataPath: String(metadataPath || ""),
    dateArchivage: getCurrentSignatureTimestamp(),
  };
}

function findReusableArchivedDocument(person, docType) {
  if (!state.data || !person) {
    return null;
  }
  const typeLabel = getDocumentTypeLabel(docType);
  const archives = (state.data.documentsArchives || []).filter(
    (entry) =>
      String(entry.personId || "") === String(person.id || "") &&
      normalizeText(entry.typeDocument) === typeLabel &&
      Boolean(entry.pdfPath)
  );
  const fingerprint = getDocumentFingerprint(person, docType);
  if (typeLabel === "ARRIVEE") {
    return (
      archives.find(
        (entry) =>
          getDocumentArchiveSignatureStatus(entry) === "SIGNE" &&
          String(entry.fingerprint || "") === fingerprint
      ) || null
    );
  }
  return archives.find((entry) => String(entry.fingerprint || "") === fingerprint) || null;
}

function findActiveArchiveEntry(personId, docType) {
  if (!state.data || !Array.isArray(state.data.documentsArchives)) {
    return null;
  }
  const typeLabel = getDocumentTypeLabel(docType);
  const expectedId = `DOC-${typeLabel}-${personId}`;
  const byId = state.data.documentsArchives.find((entry) => String(entry?.id || "") === expectedId);
  if (byId) {
    return byId;
  }
  return (
    state.data.documentsArchives.find(
      (entry) =>
        String(entry?.personId || "") === String(personId || "") &&
        normalizeText(entry?.typeDocument) === typeLabel
    ) || null
  );
}

async function verifyActiveArchiveOpenable(personId, docType) {
  const entry = findActiveArchiveEntry(personId, docType);
  if (!entry) {
    return { ok: false, reason: "ARCHIVE_ENTRY_NOT_FOUND" };
  }
  const resolved = resolveArchiveOpenTarget(entry);
  const targetUrl = String(resolved?.url || "").trim();
  if (!targetUrl) {
    return { ok: false, reason: "ARCHIVE_OPEN_URL_EMPTY" };
  }
  if (normalizePdfQualityStatus(entry?.pdfQualityStatus) !== "VALID") {
    return { ok: false, reason: "ARCHIVE_QUALITY_NOT_VALID" };
  }
  if (/^\/api\/pdf-file\?/i.test(targetUrl)) {
    try {
      const response = await fetch(targetUrl, { method: "HEAD", cache: "no-store" });
      if (!response.ok) {
        return { ok: false, reason: `ARCHIVE_FILE_HEAD_${response.status}` };
      }
    } catch (error) {
      return { ok: false, reason: "ARCHIVE_FILE_HEAD_FAILED" };
    }
  }
  return { ok: true, reason: "OK" };
}

function upsertDocumentArchiveEntry(entry) {
  if (!state.data) {
    return;
  }
  if (!Array.isArray(state.data.documentsArchives)) {
    state.data.documentsArchives = [];
  }
  const index = state.data.documentsArchives.findIndex((currentEntry) => String(currentEntry.id) === String(entry.id));
  if (index >= 0) {
    state.data.documentsArchives[index] = {
      ...state.data.documentsArchives[index],
      ...entry,
    };
  } else {
    state.data.documentsArchives.push(entry);
  }
  state.data.documentsArchives = state.data.documentsArchives.filter((currentEntry) => {
    if (String(currentEntry.id || "") === String(entry.id || "")) {
      return true;
    }
    return !(
      String(currentEntry.personId || "") === String(entry.personId || "") &&
      normalizeText(currentEntry.typeDocument) === normalizeText(entry.typeDocument)
    );
  });
  sortDocumentsArchives();
}

async function registerArchivedDocument(person, docType, pdfPath, metadataPath, archiveMode = "STANDARD", archiveDetails = null) {
  if (!state.data || !person || !pdfPath) {
    return;
  }
  const existingTypeArchive = findActiveArchiveEntry(person.id, docType);
  if (normalizeText(docType) === "EXIT" && getDossierStatus(person) !== "SORTI" && !existingTypeArchive) {
    return;
  }
  const existingEntry = findReusableArchivedDocument(person, docType) || existingTypeArchive;
  upsertDocumentArchiveEntry(buildDocumentArchiveEntry(person, docType, pdfPath, metadataPath, archiveMode, existingEntry, archiveDetails));
  markDirty();
  await saveDataToFile({
    silent: true,
    reloadAfter: false,
    successText: "ARCHIVE PDF MISE A JOUR",
  });
  renderDocumentsArchivePage();
  showActionStatus("update", "PDF SIGNE ARCHIVE");
}

function renderDocumentsArchivePage() {
  const body = document.getElementById("documents-archives-body");
  if (!body) {
    return false;
  }

  const params = new URLSearchParams(window.location.search);
  const personIdFromQuery = String(params.get("personId") || params.get("personld") || "");
  if (!params.get("personId") && params.get("personld")) {
    const nextUrl = new URL(window.location.href);
    nextUrl.searchParams.set("personId", personIdFromQuery);
    nextUrl.searchParams.delete("personld");
    window.history.replaceState({}, "", nextUrl.toString());
  }
  const filterForm = document.getElementById("documents-archives-filter-form");
  const lockedPersonId = personIdFromQuery || String(getCurrentPersonId() || "");
  const lockedPerson = lockedPersonId
    ? (state.data?.personnes || []).find((person) => String(person?.id || "") === lockedPersonId) || null
    : null;
  const getArchiveFilterField = (fieldName) =>
    filterForm ? filterForm.querySelector(`[name="${fieldName}"]`) : null;
  const archiveSearchField = getArchiveFilterField("archiveSearch");
  if (archiveSearchField instanceof HTMLInputElement && lockedPerson) {
    const label = `${lockedPerson.nom || ""} ${lockedPerson.prenom || ""}`.trim();
    const currentSearch = normalizeText(archiveSearchField.value || "");
    const expectedTokens = normalizeText(label).split(/\s+/).filter(Boolean);
    const matchesLockedPersonLabel = expectedTokens.length
      ? expectedTokens.every((token) => currentSearch.includes(token))
      : true;
    if (!currentSearch || !matchesLockedPersonLabel) {
      archiveSearchField.value = label;
      archiveSearchField.defaultValue = label;
    }
  }
  const search = normalizeText(archiveSearchField?.value);
  const searchTokens = search
    ? search
        .split(/[\s\-–—/]+/)
        .map((token) => normalizeText(token))
        .filter(Boolean)
    : [];
  const typeDocumentField = getArchiveFilterField("archiveTypeDocument");
  const archiveSiteSelect = getArchiveFilterField("archiveSite");
  const statutSignatureField = getArchiveFilterField("archiveStatutSignature");
  const typeDocument = normalizeText(typeDocumentField?.value);
  const site = normalizeText(archiveSiteSelect?.value);
  const statutSignature = normalizeText(statutSignatureField?.value);
  syncSelectOptions(archiveSiteSelect, state.data?.listes?.sites || [], "TOUS");

  let totalArchives = 0;
  let totalArrivalArchives = 0;
  let totalExitArchives = 0;
  const personsById = new Map(
    (state.data?.personnes || []).map((person) => [String(person?.id || ""), person])
  );
  const allPersons = state.data?.personnes || [];
  const allArchives = state.data?.documentsArchives || [];
  const latestArchivePerPersonAndType = getLatestArchivePerPersonAndType(allArchives);
  const trackedDocumentTypes = [
    { label: "ARRIVEE", signatureType: "arrival" },
    { label: "SORTIE", signatureType: "exit" },
  ];
  const personsToProcess = lockedPersonId
    ? allPersons.filter((person) => String(person?.id || "") === String(lockedPersonId).trim())
    : allPersons;
  const signatureStateCache = new Map();
  const getCachedSignatureState = (person, docType) => {
    const personId = String(person?.id || "").trim();
    const key = `${personId}|${docType}`;
    if (!signatureStateCache.has(key)) {
      signatureStateCache.set(key, isDocumentFullySigned(person, docType));
    }
    return signatureStateCache.get(key);
  };
  const resolveWorkflowStatus = (entry, person) => {
    const baseStatus = getDocumentArchiveSignatureStatus(entry);
    if (
      baseStatus === "SIGNE" ||
      baseStatus === "ATTENTE DE GENERATION" ||
      baseStatus === "EN ATTENTE DE SIGNATURE"
    ) {
      return baseStatus;
    }
    if (!person) {
      return baseStatus || "EN ATTENTE DE SIGNATURE";
    }
    const docType = getArchiveDocTypeKey(entry?.typeDocument) === "SORTIE" ? "exit" : "arrival";
    return getCachedSignatureState(person, docType) ? "ATTENTE DE GENERATION" : "EN ATTENTE DE SIGNATURE";
  };

  const resolveArchiveDisplayData = (entry) => {
    const person = personsById.get(String(entry?.personId || ""));
    if (!person) {
      return {
        nom: entry?.nom || "-",
        prenom: entry?.prenom || "-",
        sites: entry?.sites || "-",
      };
    }
    return {
      nom: person.nom || entry?.nom || "-",
      prenom: person.prenom || entry?.prenom || "-",
      sites: getPersonSiteLabel(person) || entry?.sites || "-",
    };
  };
  const findLatestArchiveForPersonAndType = (person, typeLabel) => {
    const personId = String(person?.id || "").trim();
    if (!personId || !typeLabel) {
      return null;
    }
    const key = `${personId}|${typeLabel}`;
    const direct = latestArchivePerPersonAndType.get(key);
    if (direct) {
      return direct;
    }
    // Legacy fallback: some old archives may miss personId but keep nom/prenom.
    const personNom = normalizeText(person?.nom || "");
    const personPrenom = normalizeText(person?.prenom || "");
    const candidates = allArchives.filter((entry) => {
      if (normalizeArchiveTypeLabel(entry?.typeDocument) !== typeLabel) {
        return false;
      }
      const entryPersonId = String(entry?.personId || "").trim();
      if (entryPersonId && entryPersonId === personId) {
        return true;
      }
      // Legacy / inconsistent data fallback:
      // if personId is missing OR incorrect, keep same nom+prenom match as source of truth.
      return (
        normalizeText(entry?.nom || "") === personNom &&
        normalizeText(entry?.prenom || "") === personPrenom
      );
    });
    if (!candidates.length) {
      return null;
    }
    return candidates
      .slice()
      .sort((left, right) => (Date.parse(String(right?.dateArchivage || "")) || 0) - (Date.parse(String(left?.dateArchivage || "")) || 0))[0];
  };
  const archiveWorkflowRows = [];
  const latestSignedArchives = [];
  personsToProcess.forEach((person) => {
    trackedDocumentTypes.forEach(({ label, signatureType }) => {
      const personId = String(person?.id || "");
      if (!personId) {
        return;
      }
      if (label === "SORTIE" && getDossierStatus(person) !== "SORTI") {
        return;
      }

      const latestEntry = findLatestArchiveForPersonAndType(person, label);
      const effects = Array.isArray(person?.effetsConfies) ? person.effetsConfies : [];
      const totalFacturable = signatureType === "exit"
        ? effects.reduce((sum, effect) => sum + getEffectReplacementCost(person, effect), 0)
        : 0;
      const workflowStatus = latestEntry
        ? resolveWorkflowStatus(latestEntry, person)
        : getCachedSignatureState(person, signatureType)
          ? "ATTENTE DE GENERATION"
          : "EN ATTENTE DE SIGNATURE";

      if (latestEntry) {
        archiveWorkflowRows.push({
          ...latestEntry,
          __workflowStatus: workflowStatus,
          __isWorkflowSynthetic: false,
        });
        return;
      }
      // No synthetic row: archives must show only real generated documents.
      return;
    });
  });
  const archiveMatchesLockedPerson = (entry) => {
    if (!lockedPersonId) {
      return true;
    }
    const entryPersonId = String(entry?.personId || "").trim();
    if (entryPersonId && entryPersonId === String(lockedPersonId).trim()) {
      return true;
    }
    if (!lockedPerson) {
      return false;
    }
    return (
      normalizeText(entry?.nom || "") === normalizeText(lockedPerson.nom || "") &&
      normalizeText(entry?.prenom || "") === normalizeText(lockedPerson.prenom || "")
    );
  };
  const archives = archiveWorkflowRows.filter((entry) => {
    totalArchives += 1;
    if (getArchiveDocTypeKey(entry.typeDocument) === "ARRIVEE") {
      totalArrivalArchives += 1;
    }
    if (getArchiveDocTypeKey(entry.typeDocument) === "SORTIE") {
      totalExitArchives += 1;
    }
    if (!archiveMatchesLockedPerson(entry)) {
      return false;
    }
    if (typeDocument && getArchiveDocTypeKey(entry.typeDocument) !== getArchiveDocTypeKey(typeDocument)) {
      return false;
    }
    if (!archiveEntryMatchesSite(entry, site)) {
      return false;
    }
    const workflowStatus = String(
      entry?.__workflowStatus || getDocumentArchiveWorkflowStatus(entry, personsById.get(String(entry?.personId || "")))
    );
    if (statutSignature && workflowStatus !== statutSignature) {
      return false;
    }
    if (searchTokens.length) {
      const display = resolveArchiveDisplayData(entry);
      const haystack = [
        display.nom,
        display.prenom,
        entry.typeDocument,
        display.sites,
        entry.typePersonnel,
        entry.typeContrat,
        entry.pdfPath,
      ]
        .map(normalizeText)
        .join(" ");
      const matchesAllTokens = searchTokens.every((token) => haystack.includes(token));
      if (!matchesAllTokens) {
        return false;
      }
    }
    return true;
  });
  const groupedArchives = sortArchivesForTable(archives, resolveArchiveDisplayData);
  latestArchivePerPersonAndType.forEach((entry) => {
    const person = personsById.get(String(entry?.personId || ""));
    const status = resolveWorkflowStatus(entry, person);
    if (status === "SIGNE") {
      latestSignedArchives.push({ ...entry, __workflowStatus: status });
    }
  });

  const totalNode = document.getElementById("archive-count-total");
  const arrivalNode = document.getElementById("archive-count-arrival");
  const exitNode = document.getElementById("archive-count-exit");
  if (totalNode) {
    totalNode.textContent = String(totalArchives);
  }
  if (arrivalNode) {
    arrivalNode.textContent = String(totalArrivalArchives);
  }
  if (exitNode) {
    exitNode.textContent = String(totalExitArchives);
  }

  const storageArrivalNode = document.getElementById("archive-storage-arrival");
  const storageExitNode = document.getElementById("archive-storage-exit");
  const storageLastUpdateNode = document.getElementById("archive-storage-last-update");
  const arrivalStorageCount = latestSignedArchives.filter(
    (entry) => normalizeText(entry?.typeDocument) === "ARRIVEE"
  ).length;
  const exitStorageCount = latestSignedArchives.filter(
    (entry) => normalizeText(entry?.typeDocument) === "SORTIE"
  ).length;
  const latestArchiveMs = latestSignedArchives.reduce((latest, entry) => {
    const ms = Date.parse(String(entry?.dateArchivage || ""));
    if (!Number.isFinite(ms)) {
      return latest;
    }
    return ms > latest ? ms : latest;
  }, 0);
  if (storageArrivalNode) {
    storageArrivalNode.textContent = String(arrivalStorageCount);
  }
  if (storageExitNode) {
    storageExitNode.textContent = String(exitStorageCount);
  }
  if (storageLastUpdateNode) {
    storageLastUpdateNode.textContent = latestArchiveMs > 0 ? formatSignatureTimestamp(new Date(latestArchiveMs).toISOString()) : "AUCUNE";
  }

  if (!groupedArchives.length) {
    const documentsArchivesRenderSignature = [
      "documentsArchives",
      "empty",
      String(state.supabaseRevision || ""),
      String(state.latestDataEtag || ""),
      String(state.urgentMode ? "1" : "0"),
      String(personIdFromQuery || ""),
      String(search || ""),
      String(typeDocument || ""),
      String(site || ""),
      String(statutSignature || ""),
      String(totalArchives || 0),
      String(totalArrivalArchives || 0),
      String(totalExitArchives || 0),
      String(lockedPersonId || ""),
      "0",
    ].join("|");
    if (state.listRenderCache.documentsArchives === documentsArchivesRenderSignature) {
      return false;
    }
    const peopleCount = Array.isArray(state.data?.personnes) ? state.data.personnes.length : 0;
    const emptyMessage = peopleCount === 0
      ? "AUCUNE DONNEE CHARGEE (VERIFIER CONNEXION/SESSION)"
      : "AUCUN DOCUMENT ARCHIVE";
    state.listRenderCache.documentsArchives = documentsArchivesRenderSignature;
    body.innerHTML = buildEmptyTableRow(body, emptyMessage, 11);
    return true;
  }

  const rowsHtml = groupedArchives
    .map(
      (entry) => {
        const display = resolveArchiveDisplayData(entry);
        const person = personsById.get(String(entry?.personId || ""));
        const workflowStatus = String(
          entry?.__workflowStatus || resolveWorkflowStatus(entry, person)
        );
        const openResolution = resolveArchiveOpenTarget(entry);
        const openPath = String(openResolution.url || "");
        const legacyOpenPath = getDocumentArchiveOpenPath(entry);
        const openHref = normalizeDirectPdfOpenUrl(openPath) || normalizeDirectPdfOpenUrl(legacyOpenPath);
        const hasPdf = Boolean(openHref);
        const typeLabel = normalizeText(entry.typeDocument || "");
        const typeIcon = typeLabel === "SORTIE" ? "🔴" : typeLabel === "ARRIVEE" ? "🟢" : "⚪";
        const typeTitle = typeLabel || "TYPE INCONNU";
        const openLocalUrl = String(entry?.openLocalUrl || "").trim();
        const openRemoteUrl = String(entry?.openRemoteUrl || "").trim();
        const publicUrl = String(entry?.publicUrl || "").trim();
        const filename = String(entry?.filename || "").trim();
        const localPathInfo = String(entry?.localPath || "").trim();
        const targetAttributes = 'target="_blank" rel="noopener"';
        const openButton = hasPdf
          ? `<a class="archive-pdf-button js-open-archive-pdf" href="${escapeHtml(openHref)}" ${targetAttributes} aria-label="OUVRIR PDF" data-archive-id="${escapeHtml(String(entry?.id || ""))}" data-open-url="${escapeHtml(openHref)}" data-local-url="${escapeHtml(openLocalUrl)}" data-remote-url="${escapeHtml(openRemoteUrl)}" data-public-url="${escapeHtml(publicUrl)}" data-legacy-url="${escapeHtml(legacyOpenPath)}" data-filename="${escapeHtml(filename)}"><span class="archive-pdf-button__icon" aria-hidden="true"><img src="assets/ui/icone-pdf.png" alt="" class="archive-pdf-button__image" /></span></a>`
          : "-";
        const deleteButton = hasPdf && !entry?.__isWorkflowSynthetic
          ? `<button type="button" class="table-link js-delete-archive-row" data-archive-id="${escapeHtml(String(entry.id || ""))}">SUPPRIMER</button>`
          : "";
        return `<tr>
        <td>${escapeHtml(display.nom)}</td>
        <td>${escapeHtml(display.prenom)}</td>
        <td title="${escapeHtml(typeTitle)}" aria-label="${escapeHtml(typeTitle)}">${typeIcon}</td>
        <td>${escapeHtml(formatDate(entry.dateDocument) || "-")}</td>
        <td>${escapeHtml(formatTime(entry.dateArchivage) || "-")}</td>
        <td>${escapeHtml(display.sites)}</td>
         <td>${getDocumentArchiveStatusCellMarkup(workflowStatus)}</td>
        <td>${escapeHtml(String(getArchiveDisplayedTotalEffets(entry)))}</td>
        <td>${formatAmountWithEuro(entry.totalFacturable || 0)}</td>
        <td title="${escapeHtml(localPathInfo)}">${escapeHtml(getDocumentArchiveVersionLabel(entry))}</td>
        <td class="archive-actions-cell">${openButton} ${deleteButton}</td>
      </tr>`;
      }
    )
    .map((rowHtml) => rowHtml.trim());

  const documentsArchivesRenderSignature = [
    "documentsArchives",
    String(state.supabaseRevision || ""),
    String(state.latestDataEtag || ""),
    String(state.urgentMode ? "1" : "0"),
    String(personIdFromQuery || ""),
    String(search || ""),
    String(typeDocument || ""),
    String(site || ""),
    String(statutSignature || ""),
    String(totalArchives || 0),
    String(totalArrivalArchives || 0),
    String(totalExitArchives || 0),
    String(lockedPersonId || ""),
    String(groupedArchives.length || 0),
    groupedArchives
      .map((entry) => {
        const display = resolveArchiveDisplayData(entry);
        return [
          String(entry.id || ""),
          String(entry.personId || ""),
          String(entry.typeDocument || ""),
          String(formatDate(entry.dateDocument) || "-"),
          String(formatTime(entry.dateArchivage) || "-"),
          String(display.nom || ""),
          String(display.prenom || ""),
          String(display.sites || ""),
          String(entry.totalFacturable || 0),
          String(getDocumentArchiveWorkflowStatus(entry, personsById.get(String(entry?.personId || ""))) || ""),
        ].join("|");
      })
      .join("||"),
  ].join("|");

  if (state.listRenderCache.documentsArchives === documentsArchivesRenderSignature) {
    return false;
  }
  state.listRenderCache.documentsArchives = documentsArchivesRenderSignature;

  renderTableRowsProgressively(body, rowsHtml, buildEmptyTableRow(body, "AUCUN DOCUMENT ARCHIVE", 11), 24);
  bindArchiveRowActions();
  return true;
}

function deleteDocumentArchiveEntry(archiveId) {
  if (!archiveId || !Array.isArray(state.data?.documentsArchives)) {
    return;
  }
  const archive = state.data.documentsArchives.find((entry) => String(entry?.id || "") === String(archiveId));
  if (!archive) {
    return;
  }
  const displayName = `${archive.nom || ""} ${archive.prenom || ""}`.trim();
  const confirmDelete = window.confirm(
    `SUPPRIMER CETTE LIGNE D'ARCHIVE${displayName ? ` : ${displayName}` : ""} ?`
  );
  if (!confirmDelete) {
    return;
  }

  pushUndoSnapshot("SUPPRESSION ARCHIVE");
  state.data.documentsArchives = state.data.documentsArchives.filter(
    (entry) => String(entry?.id || "") !== String(archiveId)
  );
  markDirty();
  renderDocumentsArchivePage();
  showActionStatus("delete", `ARCHIVE SUPPRIMEE : ${archive.typeDocument || "DOCUMENT"} ${displayName}`.trim());
}

function bindArchiveRowActions() {
  const body = document.getElementById("documents-archives-body");
  if (!body || body.dataset.boundArchiveActions === "true") {
    return;
  }

  body.addEventListener("click", async (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const openPdfLink = target.closest(".js-open-archive-pdf");
    if (openPdfLink instanceof HTMLAnchorElement) {
      event.preventDefault();
      const archiveId = String(openPdfLink.dataset.archiveId || "").trim();
      const archiveEntry = archiveId && Array.isArray(state.data?.documentsArchives)
        ? state.data.documentsArchives.find((entry) => String(entry?.id || "") === archiveId)
        : null;
      const fallbackEntry = {
        directOpenUrl: String(openPdfLink.dataset.openUrl || openPdfLink.href || "").trim(),
        openLocalUrl: String(openPdfLink.dataset.localUrl || "").trim(),
        openRemoteUrl: String(openPdfLink.dataset.remoteUrl || "").trim(),
        publicUrl: String(openPdfLink.dataset.publicUrl || "").trim(),
        pdfPath: String(openPdfLink.dataset.legacyUrl || "").trim(),
        pdfQualityStatus: "UNKNOWN",
      };
      if (!handleArchiveOpenClick(archiveEntry, fallbackEntry)) {
        return;
      }
      return;
    }
    const deleteButton = target.closest(".js-delete-archive-row");
    if (!(deleteButton instanceof HTMLElement)) {
      return;
    }
    const archiveId = String(deleteButton.dataset.archiveId || "");
    if (!archiveId) {
      return;
    }
    deleteDocumentArchiveEntry(archiveId);
  });

  body.dataset.boundArchiveActions = "true";
}

function getSignatureValue(person, docType, signer) {
  if (!person?.signatures?.[docType]) {
    return "";
  }
  const rawValue = String(person.signatures[docType][signer]?.image || "");
  const publicUrl = String(person.signatures[docType][signer]?.storagePublicUrl || "");
  if (publicUrl && isSafeArchiveHttpUrl(publicUrl)) {
    return normalizeHttpUrl(publicUrl) || "";
  }
  const storedRef = String(person.signatures[docType][signer]?.storageRef || "");
  const storedRefPath = parseStorageSchemePath(storedRef);
  if (storedRefPath) {
    return getSupabaseStoragePublicUrl(storedRefPath.bucket, storedRefPath.objectPath) || "";
  }
  const storageRef = parseStorageSchemePath(rawValue);
  if (storageRef) {
    return getSupabaseStoragePublicUrl(storageRef.bucket, storageRef.objectPath) || "";
  }
  if (/^data:image\/[a-zA-Z0-9.+-]+;base64,/.test(rawValue)) {
    return rawValue;
  }
  if (/^https?:\/\//i.test(rawValue)) {
    return isSafeArchiveHttpUrl(rawValue) ? normalizeHttpUrl(rawValue) : "";
  }
  return isSafeArchiveRelativePath(rawValue) ? rawValue.replace(/^\/+/, "") : "";
}

function getRepresentativeInfo(person, docType) {
  if (!person?.representants?.[docType]) {
    return { id: "", nom: "", fonction: "" };
  }
  const representativeId = String(person.representants[docType].id || "");
  const linkedRepresentative = representativeId ? findRepresentativeById(representativeId) : null;
  if (linkedRepresentative) {
    return {
      id: String(linkedRepresentative.id || ""),
      nom: String(linkedRepresentative.nom || person.representants[docType].nom || ""),
      fonction: String(linkedRepresentative.fonction || person.representants[docType].fonction || ""),
    };
  }
  return {
    id: representativeId,
    nom: String(person.representants[docType].nom || ""),
    fonction: String(person.representants[docType].fonction || ""),
  };
}

function setRepresentativeInfo(person, docType, values) {
  if (!person) {
    return;
  }
  if (!person.representants || typeof person.representants !== "object") {
    person.representants = {};
  }
  if (!person.representants[docType] || typeof person.representants[docType] !== "object") {
    person.representants[docType] = {};
  }
  const normalizedNom = normalizeText(values.nom);
  const normalizedFonction = normalizeText(values.fonction);
  const representative = ensureRepresentativeReference(normalizedNom, normalizedFonction);
  person.representants[docType].id = representative?.id || "";
  person.representants[docType].nom = representative?.nom || normalizedNom;
  person.representants[docType].fonction = representative?.fonction || normalizedFonction;
}

function findRepresentativeByValues(nom, fonction) {
  const normalizedNom = normalizeText(nom);
  const normalizedFonction = normalizeText(fonction);
  return (state.data?.listes?.representantsSignataires || []).find(
    (entry) => normalizeText(entry.nom) === normalizedNom && normalizeText(entry.fonction) === normalizedFonction
  ) || null;
}

function findRepresentativeById(representativeId) {
  return (state.data?.listes?.representantsSignataires || []).find((entry) => entry.id === representativeId) || null;
}

function ensureRepresentativeReference(nom, fonction) {
  const normalizedNom = normalizeText(nom);
  const normalizedFonction = normalizeText(fonction);
  if (!normalizedNom && !normalizedFonction) {
    return null;
  }
  const existing = findRepresentativeByValues(normalizedNom, normalizedFonction);
  if (existing) {
    return existing;
  }
  const created = {
    id: getNextId("REP", state.data?.listes?.representantsSignataires || []),
    nom: normalizedNom,
    fonction: normalizedFonction,
  };
  state.data.listes.representantsSignataires.push(created);
  sortRepresentatives();
  return created;
}

function updateRepresentativeLinks(previousRepresentativeId, nextRepresentative) {
  (state.data?.personnes || []).forEach((person) => {
    ["arrival", "exit"].forEach((docType) => {
      const currentRepresentative = person.representants?.[docType];
      if (!currentRepresentative) {
        return;
      }
      if (String(currentRepresentative.id || "") !== String(previousRepresentativeId || "")) {
        return;
      }
      currentRepresentative.id = String(nextRepresentative?.id || "");
      currentRepresentative.nom = String(nextRepresentative?.nom || "");
      currentRepresentative.fonction = String(nextRepresentative?.fonction || "");
    });
  });
}

function sortRepresentatives() {
  if (!Array.isArray(state.data?.listes?.representantsSignataires)) {
    return;
  }
  state.data.listes.representantsSignataires.sort((left, right) => {
    const leftLabel = `${normalizeText(left.nom)} ${normalizeText(left.fonction)}`;
    const rightLabel = `${normalizeText(right.nom)} ${normalizeText(right.fonction)}`;
    return leftLabel.localeCompare(rightLabel, "fr");
  });
}

function getRepresentativeUsage(representativeId) {
  return (state.data?.personnes || []).reduce((count, person) => {
    const arrivalMatch = String(person.representants?.arrival?.id || "") === representativeId ? 1 : 0;
    const exitMatch = String(person.representants?.exit?.id || "") === representativeId ? 1 : 0;
    return count + arrivalMatch + exitMatch;
  }, 0);
}

function resolveRepresentativeSelectedId(person, docType) {
  if (!person) {
    return "";
  }
  const info = getRepresentativeInfo(person, docType);
  if (String(info.id || "").trim()) {
    return String(info.id || "");
  }
  const linkedByValues = findRepresentativeByValues(info.nom, info.fonction);
  return String(linkedByValues?.id || "");
}

function populateRepresentativeSelect(select, selectedId = "") {
  if (!(select instanceof HTMLSelectElement)) {
    return;
  }
  const currentValue = String(selectedId || "");
  const representatives = (state.data?.listes?.representantsSignataires || []).slice();
  const hasCurrentValue = !currentValue || representatives.some((entry) => String(entry.id || "") === currentValue);
  const fallbackOption =
    currentValue && !hasCurrentValue
      ? [`<option value="${escapeHtml(currentValue)}">${escapeHtml(currentValue)}</option>`]
      : [];
  select.innerHTML = ['<option value="">SELECTIONNER</option>']
    .concat(fallbackOption)
    .concat(
      representatives.map(
        (entry) =>
          `<option value="${escapeHtml(entry.id)}">${escapeHtml(entry.nom || entry.fonction || entry.id)}</option>`
      )
    )
    .join("");
  select.value = currentValue;
}

function hasRepresentativeIdentityForDocument(docType) {
  const nameInput = document.getElementById(`${docType}-signature-representant-name-input`);
  const functionInput = document.getElementById(`${docType}-signature-representant-function-input`);
  const person = getCurrentPerson();
  const representative = getRepresentativeInfo(person, docType);
  const storedIdentityReady = Boolean(normalizeText(representative.nom) && normalizeText(representative.fonction));
  if (storedIdentityReady) {
    return true;
  }
  return Boolean(
    nameInput instanceof HTMLSelectElement &&
      functionInput instanceof HTMLInputElement &&
      normalizeText(nameInput.value) &&
      normalizeText(functionInput.value)
  );
}

function captureRepresentativeIdentityDraftToState(docType, person = getCurrentPerson()) {
  if (!person) {
    return false;
  }
  const nameInput = document.getElementById(`${docType}-signature-representant-name-input`);
  const functionInput = document.getElementById(`${docType}-signature-representant-function-input`);
  if (!(nameInput instanceof HTMLSelectElement) || !(functionInput instanceof HTMLInputElement)) {
    return false;
  }
  const selectedRepresentative = findRepresentativeById(nameInput.value);
  const nextNom = normalizeText(selectedRepresentative?.nom || nameInput.options[nameInput.selectedIndex]?.text || nameInput.value || "");
  const nextFonction = normalizeText(selectedRepresentative?.fonction || functionInput.value || "");
  if (!nextNom || !nextFonction) {
    return false;
  }
  const current = getRepresentativeInfo(person, docType);
  if (normalizeText(current.nom) === nextNom && normalizeText(current.fonction) === nextFonction) {
    return false;
  }
  setRepresentativeInfo(person, docType, { nom: nextNom, fonction: nextFonction });
  markDirty();
  return true;
}

function updateRepresentativeSignatureActionState(docType) {
  document
    .querySelectorAll(`.js-signature-save[data-doc-type="${docType}"][data-signer="representant"]`)
    .forEach((button) => {
      const enabled = hasRepresentativeIdentityForDocument(docType);
      button.classList.toggle("is-disabled", !enabled);
      button.setAttribute("aria-disabled", enabled ? "false" : "true");
      button.title = enabled
        ? ""
        : "VOUS DEVEZ IDENTIFIER L'IDENTITE DU REPRESENTANT DE L'ETABLISSEMENT";
    });
}

function getSignatureValidationDate(person, docType, signer) {
  if (!person?.signatures?.[docType]) {
    return "";
  }
  const entry = person.signatures[docType][signer] || {};
  const hasSignatureImage = Boolean(
    String(entry.image || "").trim() ||
      String(entry.storageRef || "").trim() ||
      String(entry.storagePublicUrl || "").trim()
  );
  if (!hasSignatureImage) {
    return "";
  }
  return String(entry.validatedAt || "");
}

function getSignatureStorageRef(person, docType, signer) {
  if (!person?.signatures?.[docType]) {
    return "";
  }
  return String(person.signatures[docType][signer]?.storageRef || "");
}

function getSignatureStoragePublicUrl(person, docType, signer) {
  if (!person?.signatures?.[docType]) {
    return "";
  }
  return String(person.signatures[docType][signer]?.storagePublicUrl || "");
}

function setSignatureValue(person, docType, signer, value, validatedAt = "", storageRef = "", storagePublicUrl = "") {
  if (!person) {
    return;
  }
  if (!person.signatures || typeof person.signatures !== "object") {
    person.signatures = {};
  }
  if (!person.signatures[docType] || typeof person.signatures[docType] !== "object") {
    person.signatures[docType] = {};
  }
  person.signatures[docType][signer] = {
    image: String(value || ""),
    validatedAt: String(validatedAt || ""),
    storageRef: String(storageRef || ""),
    storagePublicUrl: String(storagePublicUrl || ""),
  };
}

function applySignedExitCompletion(person) {
  if (!person || !isDocumentFullySigned(person, "exit")) {
    return;
  }
  if (!person.dateSortieReelle) {
    person.dateSortieReelle = getTodayIsoDate();
  }
}

function resizeSignatureCanvas(canvas) {
  const ratio = window.devicePixelRatio || 1;
  const width = Math.max(1, Math.round(canvas.clientWidth));
  const height = Math.max(1, Math.round(canvas.clientHeight));
  const targetWidth = Math.max(1, Math.round(width * ratio));
  const targetHeight = Math.max(1, Math.round(height * ratio));
  const needsResize = canvas.width !== targetWidth || canvas.height !== targetHeight;
  if (needsResize) {
    canvas.width = targetWidth;
    canvas.height = targetHeight;
  }
  const context = canvas.getContext("2d");
  if (!context) {
    return null;
  }
  context.setTransform(ratio, 0, 0, ratio, 0, 0);
  context.lineCap = "round";
  context.lineJoin = "round";
  context.lineWidth = 2;
  context.strokeStyle = "#233f4d";
  return context;
}

function clearSignatureCanvas(canvas) {
  const context = resizeSignatureCanvas(canvas);
  if (!context) {
    return;
  }
  context.clearRect(0, 0, canvas.clientWidth, canvas.clientHeight);
}

function drawSignatureFromDataUrl(canvas, dataUrl) {
  const context = resizeSignatureCanvas(canvas);
  if (!context) {
    return;
  }
  context.clearRect(0, 0, canvas.clientWidth, canvas.clientHeight);
  if (!dataUrl) {
    return;
  }

  const image = new Image();
  image.onload = () => {
    const width = canvas.clientWidth;
    const height = canvas.clientHeight;
    context.clearRect(0, 0, width, height);
    const sourceWidth = Math.max(1, image.naturalWidth || image.width || 1);
    const sourceHeight = Math.max(1, image.naturalHeight || image.height || 1);
    const scale = Math.min(width / sourceWidth, height / sourceHeight);
    const drawWidth = sourceWidth * scale;
    const drawHeight = sourceHeight * scale;
    const offsetX = (width - drawWidth) / 2;
    const offsetY = (height - drawHeight) / 2;
    context.drawImage(image, offsetX, offsetY, drawWidth, drawHeight);
  };
  image.src = dataUrl;
}

function isCanvasSignatureBlank(canvas) {
  if (!(canvas instanceof HTMLCanvasElement)) {
    return true;
  }
  const context = canvas.getContext("2d");
  if (!context) {
    return true;
  }
  try {
    const imageData = context.getImageData(0, 0, canvas.width, canvas.height);
    const values = imageData.data;
    for (let i = 3; i < values.length; i += 4) {
      if (values[i] > 0) {
        return false;
      }
    }
    return true;
  } catch (error) {
    return true;
  }
}

function getCanvasSignaturePayload(canvas, stateRef) {
  const storedValue = String(stateRef?.pendingDataUrl || "").trim();
  if (storedValue) {
    return storedValue;
  }
  if (!(canvas instanceof HTMLCanvasElement) || isCanvasSignatureBlank(canvas)) {
    return "";
  }
  try {
    return String(canvas.toDataURL("image/png"));
  } catch (error) {
    return "";
  }
}

function refreshDocumentSignatureCanvases(docType, forcedPerson = null) {
  document.querySelectorAll(`.js-signature-canvas[data-doc-type="${docType}"]`).forEach((canvas) => {
    const isMobileSignaturePage = document.body.dataset.page === "mobile-signature";
    const person = forcedPerson || getSignatureContextPerson(isMobileSignaturePage);
    const signer = String(canvas.getAttribute("data-signer") || "");
    const dataUrl = getSignatureValue(person, docType, signer);
    const signatureValidatedAt = formatSignatureTimestamp(getSignatureValidationDate(person, docType, signer));
    const canvasState = signatureCanvases.get(canvas);
    const hasSignature = Boolean(dataUrl);
    const canvasStateSignature = [
      String(person?.id || ""),
      String(docType || ""),
      signer,
      String(dataUrl || ""),
      String(signatureValidatedAt || ""),
    ].join("|||");

    if (canvasState?.lastRenderedSignature === canvasStateSignature) {
      const statusNode = canvas
        .closest(".signature-box")
        ?.querySelector(".signature-box__status");
      if (statusNode) {
        const statusText = hasSignature
          ? signatureValidatedAt
            ? `SIGNATURE ENREGISTREE LE ${signatureValidatedAt}`
            : "SIGNATURE ENREGISTREE"
          : "AUCUNE SIGNATURE";
        statusNode.textContent = statusText;
        statusNode.classList.toggle("is-signed", hasSignature);
      }
      return;
    }

    if (canvasState) {
      canvasState.pendingDataUrl = dataUrl;
      canvasState.lastRenderedSignature = canvasStateSignature;
    }
    drawSignatureFromDataUrl(canvas, dataUrl);

    const statusNode = canvas
      .closest(".signature-box")
      ?.querySelector(".signature-box__status");
    if (statusNode) {
      const statusText = hasSignature
        ? signatureValidatedAt
          ? `SIGNATURE ENREGISTREE LE ${signatureValidatedAt}`
          : "SIGNATURE ENREGISTREE"
        : "AUCUNE SIGNATURE";
      statusNode.textContent = statusText;
      statusNode.classList.toggle("is-signed", hasSignature);
    }
    return;
  });
}

function bindRepresentativeFields() {
  [
    { docType: "arrival", nameInputId: "arrival-signature-representant-name-input", functionInputId: "arrival-signature-representant-function-input" },
    { docType: "exit", nameInputId: "exit-signature-representant-name-input", functionInputId: "exit-signature-representant-function-input" },
  ].forEach(({ docType, nameInputId, functionInputId }) => {
    const nameInput = document.getElementById(nameInputId);
    const functionInput = document.getElementById(functionInputId);
    if (!(nameInput instanceof HTMLSelectElement) || !(functionInput instanceof HTMLInputElement)) {
      return;
    }

    const syncRepresentativeOptions = () => {
      const person = getCurrentPerson();
      populateRepresentativeSelect(nameInput, person ? resolveRepresentativeSelectedId(person, docType) : "");
    };

    const applyRepresentativeSelection = () => {
      const person = getCurrentPerson();
      if (!person) {
        functionInput.value = "";
        return;
      }
      const representative = findRepresentativeById(nameInput.value);
      functionInput.value = representative?.fonction || "";
    };
    let representativeNameDebounceId = 0;
    const scheduleRepresentativeSelection = () => {
      if (representativeNameDebounceId) {
        window.clearTimeout(representativeNameDebounceId);
      }
      representativeNameDebounceId = window.setTimeout(() => {
        representativeNameDebounceId = 0;
        applyRepresentativeSelection();
      }, FILTER_INPUT_DEBOUNCE_MS);
    };

    const syncRepresentative = () => {
      const person = getCurrentPerson();
      syncRepresentativeOptions();
      if (!person) {
        functionInput.value = "";
        updateRepresentativeSignatureActionState(docType);
        return;
      }
      const currentRepresentative = getRepresentativeInfo(person, docType);
      const linkedRepresentative =
        findRepresentativeById(currentRepresentative.id) ||
        findRepresentativeByValues(currentRepresentative.nom, currentRepresentative.fonction);
      functionInput.value = linkedRepresentative?.fonction || currentRepresentative.fonction || "";
      updateRepresentativeSignatureActionState(docType);
    };

    const saveRepresentative = () => {
      const person = getCurrentPerson();
      if (!person) {
        functionInput.value = "";
        updateRepresentativeSignatureActionState(docType);
        return;
      }
      const representative = findRepresentativeById(nameInput.value);
      functionInput.value = representative?.fonction || "";
      setRepresentativeInfo(person, docType, {
        nom: representative?.nom || "",
        fonction: representative?.fonction || "",
      });
      const representantNameNode = document.getElementById(`${docType}-signature-representant-name`);
      const representantFunctionNode = document.getElementById(`${docType}-signature-representant-function`);
      const signatureRepresentantDateNode = document.getElementById(`${docType}-signature-representant-date`);
      if (representantNameNode) {
        representantNameNode.textContent = representative?.nom || "-";
      }
      if (representantFunctionNode) {
        representantFunctionNode.textContent = representative?.fonction || "-";
      }
      if (signatureRepresentantDateNode) {
        signatureRepresentantDateNode.textContent =
          formatSignatureTimestamp(getSignatureValidationDate(person, docType, "representant")) || "-";
      }
      markDirty();
      updateRepresentativeSignatureActionState(docType);
      syncDocumentMobileSignatureLink(docType, person.id, "representant");
      showActionStatus("update", "REPRESENTANT MIS A JOUR");
    };

    syncRepresentative();
    nameInput.onfocus = syncRepresentativeOptions;
    nameInput.onmousedown = syncRepresentativeOptions;
    nameInput.onclick = syncRepresentativeOptions;
    nameInput.oninput = () => {
      scheduleRepresentativeSelection();
    };
    nameInput.onchange = () => {
      applyRepresentativeSelection();
      updateRepresentativeSignatureActionState(docType);
      saveRepresentative();
    };
    updateRepresentativeSignatureActionState(docType);
  });
}

function bindSignatureCanvases() {
  document.querySelectorAll(".js-signature-canvas").forEach((canvas) => {
    if (signatureCanvases.has(canvas)) {
      return;
    }

    const stateRef = {
      drawing: false,
      moved: false,
      pointerId: null,
      context: null,
      pendingDataUrl: "",
    };
    signatureCanvases.set(canvas, stateRef);

    const getContext = () => {
      stateRef.context = resizeSignatureCanvas(canvas);
      return stateRef.context;
    };

    const getPoint = (event) => {
      const rect = canvas.getBoundingClientRect();
      return {
        x: event.clientX - rect.left,
        y: event.clientY - rect.top,
      };
    };

    const storePendingSignature = () => {
      stateRef.pendingDataUrl = canvas.toDataURL("image/png");
    };

    const saveSignature = async () => {
      const isMobileSignaturePage = document.body.dataset.page === "mobile-signature";
      const person = getSignatureContextPerson(isMobileSignaturePage);
      const docType = String(canvas.getAttribute("data-doc-type") || "");
      const signer = String(canvas.getAttribute("data-signer") || "");
      if (!person || !docType || !signer) {
        return;
      }
      const wasFullySigned = isDocumentFullySigned(person, docType);
      const nextValue = getCanvasSignaturePayload(canvas, stateRef);
      const validatedAt = nextValue ? getCurrentSignatureTimestamp() : "";
      let signatureStorageRef = "";
      let signatureStoragePublicUrl = "";
      const currentMobileRequest =
        document.body.dataset.page === "mobile-signature"
          ? getCurrentMobileSignatureRequest()
          : null;
      const requestMatchesCanvas =
        currentMobileRequest &&
        normalizeText(currentMobileRequest.docType) === normalizeText(docType) &&
        normalizeMobileSignatureSigner(currentMobileRequest.signer || "") === signer;

      stateRef.pendingDataUrl = nextValue;
      if (nextValue && isSupabaseConfigured() && !isMobileSignaturePage) {
        try {
          const signatureUpload = await uploadSignatureImageToSupabaseStorage(docType, person, signer, nextValue);
          signatureStorageRef = String(signatureUpload?.storageRef || "");
          signatureStoragePublicUrl = String(signatureUpload?.publicUrl || "");
          if (signatureStorageRef) {
            console.info("[SUPABASE][SIGNATURE] final storage path", signatureStorageRef);
          }
        } catch (signatureUploadError) {
          console.error("[SUPABASE][SIGNATURE] upload fail", signatureUploadError);
          const message = String(signatureUploadError?.message || "ERREUR INCONNUE").slice(0, 160);
          showDataStatus(`UPLOAD SIGNATURE SUPABASE IMPOSSIBLE (${message})`);
        }
      }
      setSignatureValue(
        person,
        docType,
        signer,
        nextValue,
        validatedAt,
        signatureStorageRef,
        signatureStoragePublicUrl
      );
      if (isMobileSignaturePage && nextValue) {
        rememberMobileSignatureRuntimeSignature(
          currentMobileRequest,
          person,
          docType,
          signer,
          nextValue,
          validatedAt,
          signatureStorageRef,
          signatureStoragePublicUrl
        );
      }
      if (!isMobileSignaturePage) {
        if (docType === "arrival") {
          renderArrivalDocument(person.id);
        } else if (docType === "exit") {
          renderExitDocument(person.id);
        }
        refreshDocumentSignatureCanvases(docType);
      }
      if (nextValue && docType === "exit") {
        applySignedExitCompletion(person);
      }
      const isNowFullySigned = isDocumentFullySigned(person, docType);
      markDirty();
      const saveText =
        isNowFullySigned && !wasFullySigned
          ? "DOCUMENT SIGNE - SAUVEGARDE AUTOMATIQUE"
          : nextValue
            ? "SIGNATURE VALIDEE"
            : "SIGNATURE SUPPRIMEE";
      const signatureBox = canvas.closest(".signature-box");
      if (nextValue && signatureBox instanceof HTMLElement) {
        signatureBox.classList.remove("signature-box--validated-once");
        void signatureBox.offsetWidth;
        signatureBox.classList.add("signature-box--validated-once");
        window.setTimeout(() => {
          signatureBox.classList.remove("signature-box--validated-once");
        }, 700);
      }
      showActionStatus(nextValue ? "update" : "delete", saveText);
      // Keep mobile signature page open to show explicit confirmation and refreshed UI.
      const mustAlertAndClose = false;
      if (requestMatchesCanvas && nextValue) {
        markMobileSignatureRequestSigned(currentMobileRequest);
      }
      if (isMobileSignaturePage && getDataBackendMode() === "SUPABASE" && nextValue) {
        const mobileRequestToken = String(currentMobileRequest?.token || getCurrentMobileSignatureToken() || "");
        const latestPayload = await saveSupabaseSignatureWithRebase({
          personId: person.id,
          docType,
          signer,
          signatureValue: nextValue,
          validatedAt,
          storageRef: signatureStorageRef,
          storagePublicUrl: signatureStoragePublicUrl,
          mobileRequestToken,
        });
        const latestPerson = Array.isArray(latestPayload?.personnes)
          ? latestPayload.personnes.find((entry) => String(entry?.id || "") === String(person.id || "")) || person
          : person;
        setTimeout(() => {
          saveMobileSignatureRecordToSupabase({
            token: mobileRequestToken,
            personId: person.id,
            docType,
            signer,
            person: latestPerson,
            signatureValue: nextValue,
            validatedAt,
            storageRef: signatureStorageRef,
            storagePublicUrl: signatureStoragePublicUrl,
          }).catch((signatureRecordError) => {
            console.warn("[SUPABASE][SIGNATURE] table signatures non bloquante", signatureRecordError);
          });
        }, 0);
        clearWorkingData();
        state.isDirty = false;
        clearUndoStack();
        renderDirtyState();
        showDataStatus("SIGNATURE ENREGISTREE");
      } else {
        await saveDataToFile(
          isMobileSignaturePage
            ? {
                // Mobile signature: keep UI stable, avoid immediate stale reload wiping the just-signed value.
                silent: false,
                reloadAfter: false,
                successText: saveText,
                alertText: "SIGNATURE ENREGISTREE",
                closeAfterAlert: false,
                throwOnConflict: true,
              }
            : {
                silent: !mustAlertAndClose,
                reloadAfter: !mustAlertAndClose,
                successText: saveText,
                alertText: "DONNEES SUPABASE MISES A JOUR",
                closeAfterAlert: mustAlertAndClose,
              }
        );
      }
      if (document.body.dataset.page === "mobile-signature") {
        renderMobileSignaturePage();
        refreshDocumentSignatureCanvases(docType, person);
        showDataStatus("SIGNATURE ENREGISTREE - VOUS POUVEZ FERMER CETTE PAGE");
      }
    };

    canvas.addEventListener("pointerdown", (event) => {
      const isMobileSignaturePage = document.body.dataset.page === "mobile-signature";
      const targetPerson = getSignatureContextPerson(isMobileSignaturePage);
      if (document.body.dataset.pdfMode === "true" || !targetPerson) {
        return;
      }
      const context = getContext();
      if (!context) {
        return;
      }
      stateRef.drawing = true;
      stateRef.moved = false;
      stateRef.pointerId = event.pointerId;
      canvas.setPointerCapture(event.pointerId);
      const point = getPoint(event);
      context.beginPath();
      context.arc(point.x, point.y, 0.7, 0, Math.PI * 2);
      context.fillStyle = "#233f4d";
      context.fill();
      context.beginPath();
      context.moveTo(point.x, point.y);
      stateRef.moved = true;
      storePendingSignature();
      event.preventDefault();
    });

    canvas.addEventListener("pointermove", (event) => {
      if (!stateRef.drawing || stateRef.pointerId !== event.pointerId) {
        return;
      }
      const context = stateRef.context || getContext();
      if (!context) {
        return;
      }
      const point = getPoint(event);
      context.lineTo(point.x, point.y);
      context.stroke();
      stateRef.moved = true;
      if (!isCanvasSignatureBlank(canvas)) {
        storePendingSignature();
      }
      event.preventDefault();
    });

    const finishDrawing = (event) => {
      if (!stateRef.drawing || stateRef.pointerId !== event.pointerId) {
        return;
      }
      if (canvas.hasPointerCapture(event.pointerId)) {
        canvas.releasePointerCapture(event.pointerId);
      }
      stateRef.drawing = false;
      stateRef.pointerId = null;
      if (stateRef.moved) {
        storePendingSignature();
        showDataStatus("SIGNATURE DESSINEE - CLIQUER SUR VALIDER LA SIGNATURE");
      }
      event.preventDefault();
    };

    canvas.addEventListener("pointerup", finishDrawing);
    canvas.addEventListener("pointercancel", finishDrawing);

    const signatureBox = canvas.closest(".signature-box");
    const clearButton = signatureBox?.querySelector(".js-signature-clear");
    const saveButton = signatureBox?.querySelector(".js-signature-save");

    if (saveButton instanceof HTMLButtonElement) {
      saveButton.onclick = async () => {
        const isMobileSignaturePage = document.body.dataset.page === "mobile-signature";
        const person = getSignatureContextPerson(isMobileSignaturePage);
        if (!person) {
          return;
        }
        const docType = String(canvas.getAttribute("data-doc-type") || "");
        const signer = String(canvas.getAttribute("data-signer") || "");
        const currentSignerHasSignature = Boolean(getSignatureValue(person, docType, signer));
        const otherSigner = signer === "personnel" ? "representant" : "personnel";
        const otherSignerHasSignature = Boolean(getSignatureValue(person, docType, otherSigner));
        const pendingValue = getCanvasSignaturePayload(canvas, stateRef);
        stateRef.pendingDataUrl = pendingValue;
        if (!pendingValue) {
          const message = "AUCUNE SIGNATURE DETECTEE. SIGNEZ DANS LA ZONE AVANT DE VALIDER.";
          showDataStatus(message);
          window.alert(message);
          return;
        }
        const willFinalizeDocument = Boolean(pendingValue) && (otherSignerHasSignature || currentSignerHasSignature);
        if (
          signer === "representant" &&
          document.body.dataset.page !== "mobile-signature" &&
          !hasRepresentativeIdentityForDocument(docType)
        ) {
          showDataStatus("IDENTITE DU REPRESENTANT OBLIGATOIRE AVANT VALIDATION");
          window.alert("VOUS DEVEZ IDENTIFIER L'IDENTITE DU REPRESENTANT DE L'ETABLISSEMENT POUR VALIDATION.");
          updateRepresentativeSignatureActionState(docType);
          return;
        }
        if (willFinalizeDocument) {
          const check = validateFinalSignatureBeforeSave(person, docType);
          if (!check.ok) {
            showDataStatus(check.message);
            window.alert(check.message);
            return;
          }
        }
        await saveSignature();
      };
    }

    if (clearButton instanceof HTMLButtonElement) {
      clearButton.onclick = async () => {
        const isMobileSignaturePage = document.body.dataset.page === "mobile-signature";
        const person = getSignatureContextPerson(isMobileSignaturePage);
        const docType = String(canvas.getAttribute("data-doc-type") || "");
        const signer = String(canvas.getAttribute("data-signer") || "");
        if (!person || !docType || !signer) {
          return;
        }
        stateRef.pendingDataUrl = "";
        clearSignatureCanvas(canvas);
        setSignatureValue(person, docType, signer, "", "", "", "");
        refreshDocumentSignatureCanvases(docType, person);
        if (document.body.dataset.page === "mobile-signature") {
          const request = getCurrentMobileSignatureRequest();
          if (
            request &&
            normalizeText(request.docType) === normalizeText(docType) &&
            normalizeMobileSignatureSigner(request.signer || "") === signer
          ) {
            request.status = "EN ATTENTE";
            request.validatedAt = "";
          }
        }
        markDirty();
        showActionStatus("delete", "SIGNATURE EFFACEE");
        await saveDataToFile();
        if (document.body.dataset.page === "mobile-signature") {
          renderMobileSignaturePage();
        }
      };
    }
  });
}

function getCurrentMobileSignatureRequest() {
  const token = getCurrentMobileSignatureToken();
  if (!token) {
    return null;
  }
  cleanupExpiredMobileSignatureRequests();
  return applyMobileSignatureIdentityFromUrl(findMobileSignatureRequestByToken(token)) || buildFallbackMobileSignatureRequestFromUrl(token);
}

function getMobileSignatureTargetPerson() {
  const request = getCurrentMobileSignatureRequest();
  const signatureParams = new URLSearchParams(window.location.search);
  const personIdFromUrl = String(signatureParams.get("personId") || signatureParams.get("personld") || "").trim();
  const requestPersonId = String(request?.personId || "").trim();
  const targetPersonId = requestPersonId || personIdFromUrl;

  if (!targetPersonId || !Array.isArray(state.data?.personnes)) {
    return getCurrentPerson();
  }

  return (
    state.data.personnes.find((entry) => String(entry?.id || "") === targetPersonId) ||
    buildMobileSignaturePersonFromRequest(request, targetPersonId) ||
    getCurrentPerson()
  );
}

function buildMobileSignaturePersonFromRequest(request, personId = "") {
  if (!request) {
    return null;
  }
  const normalizedPersonId = String(personId || request.personId || "").trim();
  const nom = normalizeText(request.personNom || "");
  const prenom = normalizeText(request.personPrenom || "");
  if (!normalizedPersonId || (!nom && !prenom)) {
    return null;
  }
  return {
    id: normalizedPersonId,
    nom,
    prenom,
    site: normalizeText(request.personSite || ""),
    typePersonnel: normalizeText(request.personTypePersonnel || ""),
    typeContrat: normalizeText(request.personTypeContrat || ""),
    effetsConfies: [],
    signatures: {
      arrival: {
        personnel: { image: "", validatedAt: "", storageRef: "", storagePublicUrl: "" },
        representant: { image: "", validatedAt: "", storageRef: "", storagePublicUrl: "" },
      },
      exit: {
        personnel: { image: "", validatedAt: "", storageRef: "", storagePublicUrl: "" },
        representant: { image: "", validatedAt: "", storageRef: "", storagePublicUrl: "" },
      },
    },
    representants: {
      arrival: {
        nom: normalizeText(request.representativeNom || ""),
        fonction: normalizeText(request.representativeFonction || ""),
      },
      exit: {
        nom: normalizeText(request.representativeNom || ""),
        fonction: normalizeText(request.representativeFonction || ""),
      },
    },
    __mobileSignatureFallback: true,
  };
}

function getMobileSignatureRepresentativeInfo(person, docType, request = null) {
  const stored = person ? getRepresentativeInfo(person, docType) : { id: "", nom: "", fonction: "" };
  const requestNom = normalizeText(request?.representativeNom || "");
  const requestFonction = normalizeText(request?.representativeFonction || "");
  return {
    id: String(stored.id || ""),
    nom: normalizeText(stored.nom || requestNom),
    fonction: normalizeText(stored.fonction || requestFonction),
  };
}

function isMobileSignatureRequestValid(request) {
  if (!request) {
    return false;
  }
  if (request.status !== "EN ATTENTE") {
    return false;
  }
  const expiresAt = Date.parse(request.expiresAt || "");
  return Number.isFinite(expiresAt) && expiresAt > Date.now();
}

function markMobileSignatureRequestSigned(request) {
  if (!request) {
    return;
  }
  request.status = "SIGNEE";
  request.validatedAt = getCurrentSignatureTimestamp();
}

function getMobileSignatureRuntimeStore() {
  if (!(state.mobileSignatureRuntimeCache instanceof Map)) {
    state.mobileSignatureRuntimeCache = new Map();
  }
  return state.mobileSignatureRuntimeCache;
}

function getMobileSignatureRuntimeKey(request, person, docType, signer) {
  const token = String(request?.token || getCurrentMobileSignatureToken() || "").trim();
  const signatureParams = new URLSearchParams(window.location.search);
  const personId = String(person?.id || request?.personId || signatureParams.get("personId") || signatureParams.get("personld") || "").trim();
  const normalizedDocType = normalizeText(docType) === "EXIT" ? "exit" : "arrival";
  const normalizedSigner = normalizeMobileSignatureSigner(signer || request?.signer || "");
  if (!token || !personId || !normalizedDocType || !normalizedSigner) {
    return "";
  }
  return [token, personId, normalizedDocType, normalizedSigner].join("|||");
}

function rememberMobileSignatureRuntimeSignature(request, person, docType, signer, image, validatedAt, storageRef = "", storagePublicUrl = "") {
  const key = getMobileSignatureRuntimeKey(request, person, docType, signer);
  if (!key || !image) {
    return;
  }
  getMobileSignatureRuntimeStore().set(key, {
    image: String(image || ""),
    validatedAt: String(validatedAt || getCurrentSignatureTimestamp()),
    storageRef: String(storageRef || ""),
    storagePublicUrl: String(storagePublicUrl || ""),
  });
}

function getMobileSignatureRuntimeSignature(request, person, docType, signer) {
  const key = getMobileSignatureRuntimeKey(request, person, docType, signer);
  if (!key) {
    return null;
  }
  return getMobileSignatureRuntimeStore().get(key) || null;
}
function renderMobileSignaturePage() {
  const request = getCurrentMobileSignatureRequest();
  const person = getMobileSignatureTargetPerson();
  const docType = getCurrentMobileSignatureDocType();
  const signerFromUrl = getCurrentMobileSignatureSigner();
  const signerFromRequest = normalizeMobileSignatureSigner(request?.signer || "");
  const signer = request ? signerFromRequest : signerFromUrl;
  const normalizedDocType = normalizeText(docType) === "EXIT" ? "exit" : "arrival";
  const docLabel = normalizedDocType === "exit" ? "DOCUMENT DE SORTIE" : "DOCUMENT D'ARRIVEE";
  const representative = getMobileSignatureRepresentativeInfo(person, normalizedDocType, request);
  const representativeReady =
    signer !== "representant" || Boolean(normalizeText(representative?.nom) && normalizeText(representative?.fonction));
  const runtimeSignature = getMobileSignatureRuntimeSignature(request, person, normalizedDocType, signer);
  const signatureValidationDate = String(runtimeSignature?.validatedAt || getSignatureValidationDate(person, normalizedDocType, signer) || "");
  const signatureValue = String(runtimeSignature?.image || person?.signatures?.[normalizedDocType]?.[signer]?.image || "");
  const signatureValidatedAt = String(runtimeSignature?.validatedAt || person?.signatures?.[normalizedDocType]?.[signer]?.validatedAt || "");
  const signatureStorageRef = String(runtimeSignature?.storageRef || person?.signatures?.[normalizedDocType]?.[signer]?.storageRef || "");
  const signatureStoragePublicUrl = String(runtimeSignature?.storagePublicUrl || person?.signatures?.[normalizedDocType]?.[signer]?.storagePublicUrl || "");
  const isAlreadySigned = Boolean((request && normalizeText(request.status) === "SIGNEE") || runtimeSignature?.image);
  const isRequestUsable = Boolean(
    request &&
      isMobileSignatureRequestValid(request) &&
      normalizeText(request.docType) === normalizeText(docType) &&
      signerFromRequest === signer &&
      person
  );
  const requestStatus = String(request?.status || "");
  const requestSigner = normalizeMobileSignatureSigner(request?.signer || "");
  const requestExpiresAt = String(request?.expiresAt || "");
  const requestDocType = String(request?.docType || "");
  const requestToken = String(request?.token || "");
  const requestExpiresStamp = formatSignatureTimestamp(request?.expiresAt || "");
  const personKey = `${String(person?.id || "")}|${String(person?.nom || "")}|${String(person?.prenom || "")}`;
  const representativeKey = `${normalizeText(representative?.nom || "")}|${normalizeText(representative?.fonction || "")}`;
  const cacheSignature = [
    "mobileSignaturePage",
    requestToken,
    String(requestStatus),
    String(requestDocType),
    requestSigner,
    requestExpiresAt,
    requestExpiresStamp,
    String(isAlreadySigned),
    String(isRequestUsable),
    String(representativeReady),
    String(!request),
    personKey,
    representativeKey,
    signatureValidationDate,
    signatureValidatedAt,
    signatureValue,
    signatureStorageRef,
    signatureStoragePublicUrl,
    normalizedDocType,
    signer,
    docLabel,
  ].join("|||");

  if (state.listRenderCache.mobileSignature === cacheSignature) {
    return false;
  }
  state.listRenderCache.mobileSignature = cacheSignature;

  const titleNode = document.getElementById("mobile-signature-title");
  const subtitleNode = document.getElementById("mobile-signature-subtitle");
  const identityLabelNode = document.getElementById("mobile-signature-identity-label");
  const personNode = document.getElementById("mobile-signature-person");
  const dateNode = document.getElementById("mobile-signature-date");
  const statusNode = document.getElementById("mobile-signature-request-status");
  const panelNode = document.getElementById("mobile-signature-panel");
  const mobileCostsHead = document.getElementById("mobile-costs-head");
  const mobileCostsBody = document.getElementById("mobile-costs-body");
  const saveButton = document.querySelector(".js-signature-save");
  const clearButton = document.querySelector(".js-signature-clear");
  const canvas = document.querySelector(".js-signature-canvas");

  if (canvas) {
    canvas.setAttribute("data-doc-type", normalizedDocType);
    canvas.setAttribute("data-signer", signer);
  }
  if (saveButton) {
    saveButton.setAttribute("data-doc-type", normalizedDocType);
    saveButton.setAttribute("data-signer", signer);
  }
  if (clearButton) {
    clearButton.setAttribute("data-doc-type", normalizedDocType);
    clearButton.setAttribute("data-signer", signer);
  }

  if (titleNode) {
    titleNode.textContent =
      signer === "representant"
        ? "SIGNATURE DU REPRESENTANT DE L'ETABLISSEMENT"
        : "SIGNATURE DU PERSONNEL";
  }
  if (subtitleNode) {
    subtitleNode.textContent = docLabel;
  }
  if (identityLabelNode) {
    identityLabelNode.textContent = signer === "representant" ? "REPRESENTANT" : "PERSONNEL";
  }
  if (personNode) {
    personNode.textContent =
      signer === "representant"
        ? representative?.nom || "-"
        : person
        ? `${person.nom || ""} ${person.prenom || ""}`.trim() || "-"
        : "-";
  }
  if (dateNode) {
    const signatureDate = formatSignatureTimestamp(signatureValidationDate) || formatCurrentUiTimestamp();
    dateNode.textContent = signatureDate;
  }

  if (panelNode) {
    panelNode.hidden = !request;
  }
  if (saveButton instanceof HTMLButtonElement) {
    if (!saveButton.dataset.defaultLabel) {
      saveButton.dataset.defaultLabel = saveButton.textContent || "VALIDER LA SIGNATURE";
    }
    if (isAlreadySigned) {
      saveButton.disabled = true;
      saveButton.classList.add("is-disabled");
      saveButton.classList.add("button--validated");
      saveButton.textContent = "VALIDE";
    } else {
      saveButton.disabled = !isRequestUsable;
      saveButton.classList.toggle("is-disabled", !isRequestUsable);
      saveButton.classList.remove("button--validated");
      saveButton.textContent = saveButton.dataset.defaultLabel;
    }
  }
  if (clearButton instanceof HTMLButtonElement) {
    clearButton.disabled = !isRequestUsable;
    clearButton.classList.toggle("is-disabled", !isRequestUsable);
  }

  if (statusNode) {
    if (!request) {
      statusNode.textContent = "DEMANDE DE SIGNATURE INTROUVABLE";
    } else if (!person) {
      statusNode.textContent = "PERSONNEL INTROUVABLE";
    } else if (!isMobileSignatureRequestValid(request)) {
      statusNode.textContent =
        request.status === "SIGNEE"
          ? "SIGNATURE ENREGISTREE - VOUS POUVEZ FERMER CETTE PAGE"
          : "DEMANDE EXPIREE";
    } else if (!representativeReady) {
      statusNode.textContent = "IDENTITE DU REPRESENTANT INCOMPLETE: RENSEIGNER NOM ET FONCTION SUR LE DOCUMENT";
    } else {
      const expires = formatSignatureTimestamp(request.expiresAt);
      statusNode.textContent = expires ? `DEMANDE ACTIVE JUSQU'A ${expires}` : "DEMANDE ACTIVE";
    }
  }

  if (mobileCostsHead && mobileCostsBody) {
    renderDocumentCostsTable(normalizedDocType, mobileCostsHead, mobileCostsBody);
  }

  fillMobileSignatureShareLink(request && isRequestUsable ? request : null);
  return true;
}

function renderPersonPicker() {
  const picker = document.getElementById("person-picker-search");
  const pickerList = document.getElementById("person-picker-list");
  const suggestionBox = document.getElementById("person-picker-suggestions");
  if (!picker || !state.data?.personnes) {
    return;
  }

  const currentPersonId = getCurrentPersonId();
  const page = document.body.dataset.page || "";
  const selectedPerson = currentPersonId
    ? state.data.personnes.find((person) => String(person?.id || "") === String(currentPersonId || "")) || null
    : null;
  const personPickerSignature = [
    page,
    String(currentPersonId || ""),
    String(selectedPerson ? getPersonPickerLabel(selectedPerson) : ""),
    String(state.supabaseRevision || ""),
    String(Array.isArray(state.data.personnes) ? state.data.personnes.length : 0),
  ].join("|");
  const lastPersonPickerSignature = state.personPickerRenderCache?.[page] || "";

  if (lastPersonPickerSignature === personPickerSignature) {
    const pickerIsFocused = document.activeElement === picker;
    if (!pickerIsFocused) {
      picker.value = selectedPerson ? getPersonPickerLabel(selectedPerson) : "";
    }
    return;
  }

  state.personPickerRenderCache = state.personPickerRenderCache || {
    "person-sheet": "",
    "arrival-document": "",
    "exit-document": "",
  };
  state.personPickerRenderCache[page] = personPickerSignature;

  picker.setAttribute("autocomplete", "off");
  picker.setAttribute("autocorrect", "off");
  picker.setAttribute("autocapitalize", "off");
  picker.setAttribute("spellcheck", "false");
  picker.removeAttribute("list");

  const pickerIsFocused = document.activeElement === picker;
  if (!pickerIsFocused) {
    picker.value = selectedPerson ? getPersonPickerLabel(selectedPerson) : "";
  }

  const useDirectNavigation = page === "arrival-document" || page === "exit-document";
  const useSuggestionBox = Boolean(suggestionBox);
  const options = state.data.personnes
    .map((person) => {
      const label = getPersonPickerLabel(person);
      return `<option value="${escapeHtml(label)}"></option>`;
    })
    .join("");

  if (useSuggestionBox) {
    picker.removeAttribute("list");
    if (pickerList) {
      pickerList.innerHTML = "";
    }
  } else {
    picker.setAttribute("list", "person-picker-list");
    pickerList.innerHTML = options;
  }

  const renderSuggestions = (rawQuery = "") => {
    if (!useSuggestionBox) {
      return;
    }

    const query = normalizeText(rawQuery);
    if (!query) {
      suggestionBox.innerHTML = "";
      suggestionBox.hidden = true;
      return;
    }
    const matches = state.data.personnes
      .filter((person) => {
        return normalizeText(getPersonPickerLabel(person)).includes(query);
      })
      .slice(0, 8);

    if (!matches.length) {
      suggestionBox.innerHTML = "";
      suggestionBox.hidden = true;
      return;
    }

    suggestionBox.innerHTML = matches
      .map((person) => {
        const label = getPersonPickerLabel(person);
        return `<button type="button" class="picker-suggestions__item" data-person-id="${escapeHtml(person.id)}">${escapeHtml(label)}</button>`;
      })
      .join("");
    suggestionBox.hidden = false;
  };

  const hideSuggestions = () => {
    if (!useSuggestionBox) {
      return;
    }
    suggestionBox.hidden = true;
  };

  const applyDocumentNavigation = (personId) => {
    setCurrentPersonId(personId, "replace");
    renderPersonPicker();
    if (page === "arrival-document") {
      state.mobileSignaturePollLastSyncAt = 0;
      const arrivalRendered = renderArrivalDocument(personId);
      if (arrivalRendered) {
        refreshDocumentSignatureCanvases("arrival");
        scheduleMobileSignatureRenderSync();
      }
    } else if (page === "exit-document") {
      state.mobileSignaturePollLastSyncAt = 0;
      const exitRendered = renderExitDocument(personId);
      if (exitRendered) {
        refreshDocumentSignatureCanvases("exit");
        scheduleMobileSignatureRenderSync();
      }
    } else {
      schedulePageRender();
    }
    renderDirtyState();
    picker.blur();
  };

  const applyFullResetFromPicker = () => {
    state.filters = { ...DEFAULT_FILTERS };
    resetTableSortsForCurrentPage();
    saveNavigationContext({ filters: state.filters, personId: "" });
    if (useDirectNavigation) {
      hideSuggestions();
      applyDocumentNavigation("");
      return true;
    }
    setCurrentPersonId("", "replace");
    schedulePageRender();
    return true;
  };

  const applyPickerSelection = (mode = "push") => {
    const rawValue = String(picker.value || "");
    if (!rawValue.trim()) {
      return applyFullResetFromPicker();
    }

    const normalizedSearch = normalizeText(rawValue);
    const exactMatch = state.data.personnes.find(
      (person) => normalizeText(getPersonPickerLabel(person)) === normalizedSearch
    );
    const partialMatches = state.data.personnes.filter((person) =>
      normalizeText(getPersonPickerLabel(person)).includes(normalizedSearch)
    );
    const matchedPerson =
      exactMatch || (partialMatches.length === 1 ? partialMatches[0] : null);

    if (!matchedPerson) {
      showDataStatus("PERSONNE NON TROUVEE");
      return false;
    }

    picker.value = getPersonPickerLabel(matchedPerson);
    if (useDirectNavigation) {
      hideSuggestions();
      applyDocumentNavigation(matchedPerson.id);
      return true;
    }
    setCurrentPersonId(matchedPerson.id, mode);
    schedulePageRender();
    picker.blur();
    return true;
  };

  let pickerInputDebounceId = 0;
  const schedulePickerInputUpdate = () => {
    if (pickerInputDebounceId) {
      window.clearTimeout(pickerInputDebounceId);
    }
    pickerInputDebounceId = window.setTimeout(() => {
      pickerInputDebounceId = 0;
      const rawValue = String(picker.value || "");
      if (!rawValue.trim()) {
        if (useDirectNavigation) {
          hideSuggestions();
        }
        applyFullResetFromPicker();
        return;
      }
      if (useSuggestionBox) {
        renderSuggestions(rawValue);
        return;
      }
      const normalizedSearch = normalizeText(rawValue);
      const exactMatch = state.data.personnes.find(
        (person) => normalizeText(getPersonPickerLabel(person)) === normalizedSearch
      );
      const partialMatches = state.data.personnes.filter((person) =>
        normalizeText(getPersonPickerLabel(person)).includes(normalizedSearch)
      );
      if (exactMatch || partialMatches.length === 1) {
        applyPickerSelection("push");
      }
    }, FILTER_INPUT_DEBOUNCE_MS);
  };

  picker.oninput = () => {
    schedulePickerInputUpdate();
  };
  picker.onfocus = () => {
    if (useSuggestionBox) {
      hideSuggestions();
    }
  };
  picker.onchange = () => applyPickerSelection("push");
  picker.onsearch = () => {
    if (!String(picker.value || "").trim()) {
      if (useSuggestionBox) {
        hideSuggestions();
      }
      applyFullResetFromPicker();
      return;
    }
    applyPickerSelection("push");
  };
  picker.onblur = () => {
    if (document.activeElement === picker) {
      return;
    }
    if (useSuggestionBox) {
      window.setTimeout(() => {
        hideSuggestions();
      }, 120);
    }
    if (useDirectNavigation && !String(picker.value || "").trim()) {
      applyFullResetFromPicker();
      return;
    }
    applyPickerSelection("replace");
  };
  picker.onkeydown = (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      applyPickerSelection("push");
    }
  };

  if (useSuggestionBox && suggestionBox.dataset.bound !== "true") {
    suggestionBox.addEventListener("pointerdown", (event) => {
      const button = event.target.closest(".picker-suggestions__item");
      if (!(button instanceof HTMLElement)) {
        return;
      }
      event.preventDefault();
      const personId = button.dataset.personId || "";
      const person = state.data.personnes.find((entry) => entry.id === personId);
      if (!person) {
        return;
      }
      picker.value = getPersonPickerLabel(person);
      hideSuggestions();
      applyDocumentNavigation(person.id);
    });
    suggestionBox.dataset.bound = "true";
  }
}

function getPersonPickerLabel(person) {
  return [person.nom, person.prenom, getPersonSiteLabel(person)].filter(Boolean).join(" - ");
}

function getFilteredPersons() {
  if (!state.data?.personnes) {
    return [];
  }
  return state.data.personnes.filter((person) => matchesFilters(person, state.filters));
}

function hasUrgencyCondition(person) {
  if (!person) {
    return false;
  }
  const hasOverdueExitAlert = hasOverdueExit(person);
  const hasNonRendu = hasCurrentNonRenduEffects(person);
  return hasOverdueExitAlert && hasNonRendu;
}

function hasCurrentNonRenduEffects(person) {
  if (!person) {
    return false;
  }
  const currentEffects = getCurrentAssignedEffects(person);
  if (!currentEffects.length) {
    return false;
  }
  return (
    currentEffects.some((effect) => normalizeText(getEffectStatus(person, effect)) === "NON RENDU") ||
    isExitDue(person)
  );
}

function matchesFilters(person, filters) {
  const dossierStatus = getDossierStatus(person);
  const effects = (person.effetsConfies || []).map((effect) => ({
    ...effect,
    statutAffiche: getEffectStatus(person, effect),
  }));

  if (filters.site && !personHasSite(person, filters.site)) return false;
  if (filters.typePersonnel && normalizeText(person.typePersonnel) !== filters.typePersonnel) return false;
  if (filters.typeContrat && normalizeText(person.typeContrat) !== filters.typeContrat) return false;
  if (filters.statutDossier && dossierStatus !== filters.statutDossier) return false;
  if (
    (filters.typeEffet || filters.statutObjet) &&
    !effects.some((effect) => effectMatchesActiveObjectFilters(person, effect, filters))
  ) {
    return false;
  }
  if (state.urgentMode && !hasUrgencyCondition(person)) return false;
  if (!filters.search) return true;

  const personText = [person.nom, person.prenom, getPersonSiteLabel(person), person.typePersonnel, person.typeContrat]
    .map(normalizeText)
    .join(" ");
  const effectsText = effects
    .flatMap((effect) => [
        effect.typeEffet,
        effect.designation,
        effect.numeroIdentification,
        effect.vehiculeImmatriculation,
        effect.commentaire,
        effect.statutAffiche,
    ])
    .map(normalizeText)
    .join(" ");

  return `${personText} ${effectsText}`.includes(filters.search);
}

function renderOverview(persons) {
  const inPostNode = document.getElementById("kpi-personnes-en-poste");
  const totalEffectsNode = document.getElementById("kpi-effets-confies");
  const missingEffectsNode = document.getElementById("kpi-effets-non-rendus");
  const body = document.getElementById("overview-table-body");
  const alertsSection = document.getElementById("overview-alerts-section");
  const alertsList = document.getElementById("overview-alerts-list");

  let hasOverviewRowsChanged = false;
  if (!body) {
    return false;
  } else {

    let inPostCount = 0;
    let totalEffectsCount = 0;
    let missingEffectsCount = 0;
    const sortedPersons = sortPersonsForOverview(persons);
    const overviewSortConfig = state.tableSorts?.overviewPersons || {};
    const overviewRowsSignature = `overview|${String(overviewSortConfig.key || "")}|${String(overviewSortConfig.dir || "")}|${sortedPersons
      .map((person) => {
        const allEffects = getEffectsForActiveFilters(person);
        const assignedEffects = getCurrentAssignedEffectsForActiveFilters(person);
        const totalCosts = assignedEffects.reduce(
          (sum, effect) => sum + getEffectReplacementCost(person, effect),
          0
        );
        const nonRendus = assignedEffects.filter((effect) => getEffectStatus(person, effect) === "NON RENDU").length;
        const movementMap = getArrivalComplementMovementMap(person, allEffects);
        const movementCounts = {
          AJOUTE: 0,
          MODIFIE: 0,
          RENDU: 0,
          PERDU: 0,
          VOLE: 0,
          HS: 0,
        };
        movementMap.forEach((movement) => {
          const normalized = normalizeText(movement);
          if (Object.prototype.hasOwnProperty.call(movementCounts, normalized)) {
            movementCounts[normalized] += 1;
          }
        });
        if (getDossierStatus(person) === "EN POSTE") {
          inPostCount += 1;
        }
        allEffects.forEach((effect) => {
          totalEffectsCount += 1;
          if (getEffectStatus(person, effect) === "NON RENDU") {
            missingEffectsCount += 1;
          }
        });
        return [
          person.id || "",
          String(person.nom || ""),
          String(person.prenom || ""),
          String(person.typePersonnel || ""),
          String(person.typeContrat || ""),
          String(person.dateEntree || ""),
          String(person.dateSortiePrevue || ""),
          String(person.dateSortieReelle || ""),
          String(getDossierStatus(person)),
          String(allEffects.length || 0),
          String(assignedEffects.length || 0),
          String(nonRendus || 0),
          String(totalCosts),
          String(movementCounts.AJOUTE),
          String(movementCounts.MODIFIE),
          String(movementCounts.RENDU),
          String(movementCounts.PERDU),
          String(movementCounts.VOLE),
          String(movementCounts.HS),
        ].join("|");
      })
      .join("||")}`;
    hasOverviewRowsChanged = state.listRenderCache.overview !== overviewRowsSignature;

    if (hasOverviewRowsChanged) {
      state.listRenderCache.overview = overviewRowsSignature;
      if (inPostNode) {
        setKpiCountAnimated(inPostNode, inPostCount);
      }
      if (totalEffectsNode) {
        setKpiCountAnimated(totalEffectsNode, totalEffectsCount);
      }
      if (missingEffectsNode) {
        setKpiCountAnimated(missingEffectsNode, missingEffectsCount);
      }
      renderEffectsChart("overview-effects-chart", persons);
      const rowsHtml = buildOverviewRows(sortedPersons);
      renderTableRowsProgressively(body, [rowsHtml], buildEmptyTableRow("overview-table-body", "AUCUNE DONNEE A AFFICHER", 14), 1);
      bindPersonRowActions();
    }
  }

  if (alertsSection && alertsList) {
    const sourcePersons = Array.isArray(state.data?.personnes) ? state.data.personnes : persons;
    const alerts = sourcePersons.flatMap((person) => getPersonAlerts(person));
    const alertsSignature = `overview-alerts|${alerts
      .map((alert) =>
        `${String(alert.id || "")}|${normalizeText(alert.type || "")}|${String(alert.nom || "")}|${String(alert.prenom || "")}|${String(alert.message || "")}`
      )
      .join("||")}`;
    const hasAlertsChanged = state.listRenderCache.overviewAlerts !== alertsSignature;
    alertsSection.hidden = alerts.length === 0;
    if (hasAlertsChanged) {
      state.listRenderCache.overviewAlerts = alertsSignature;
      alertsList.innerHTML = alerts
        .map(
          (alert) => `<button type="button" class="overview-alert-item overview-alert-item--${alert.type || "dateSortiePrevue"} js-open-person-alert" data-person-id="${alert.id}">
            <span class="overview-alert-item__icon overview-alert-item__icon--${alert.type || "dateSortiePrevue"}" aria-hidden="true">${alert.type === "signaturePdf" ? "✎" : alert.type === "dateSortieReelle" ? "✕" : "!"}</span>
            <span class="overview-alert-item__content">
              <strong>${escapeHtml(`${alert.nom} ${alert.prenom}`.trim().toUpperCase())}</strong>
              <span>${escapeHtml(String(alert.message || "").toUpperCase())}</span>
            </span>
          </button>`
        )
        .join("");
    }
  }
  return hasOverviewRowsChanged;
}

function buildOverviewRows(persons) {
  if (!persons.length) {
    return buildEmptyTableRow("overview-table-body", "AUCUNE DONNEE A AFFICHER", 14);
  }
  const totals = {
    persons: persons.length,
    inPost: 0,
    effects: 0,
    nonRendus: 0,
    costs: 0,
    movements: {
      AJOUTE: 0,
      MODIFIE: 0,
      RENDU: 0,
      PERDU: 0,
      VOLE: 0,
      HS: 0,
    },
  };
  const rows = persons
    .map((person) => {
      const currentEffects = getCurrentAssignedEffectsForActiveFilters(person);
      const totalEffects = currentEffects.length;
      const totalCosts = currentEffects.reduce(
        (sum, effect) => sum + getEffectReplacementCost(person, effect),
        0
      );
      const nonRendus = currentEffects.filter(
        (effect) => getEffectStatus(person, effect) === "NON RENDU"
      ).length;
      const movementMap = getArrivalComplementMovementMap(person, getEffectsForActiveFilters(person));
      const movementCounts = {
        AJOUTE: 0,
        MODIFIE: 0,
        RENDU: 0,
        PERDU: 0,
        VOLE: 0,
        HS: 0,
      };
      movementMap.forEach((movement) => {
        const normalized = normalizeText(movement);
        if (Object.prototype.hasOwnProperty.call(movementCounts, normalized)) {
          movementCounts[normalized] += 1;
          totals.movements[normalized] += 1;
        }
      });
      if (getDossierStatus(person) === "EN POSTE") {
        totals.inPost += 1;
      }
      totals.effects += totalEffects;
      totals.nonRendus += nonRendus;
      totals.costs += totalCosts;
      const movementMarkup = Object.entries(movementCounts)
        .filter(([, count]) => count > 0)
        .map(
          ([movement, count]) =>
            `<span class="movement-badge movement-badge--${getMovementBadgeVariant(movement)}">${movement} ${count}</span>`
        )
        .join(" ");
      const alertType = getOverdueExitAlertMeta(person).type;
      const alertClass = alertType ? ` is-alert-row is-alert-row--${alertType}` : "";
      return `<tr class="js-person-row${alertClass}" data-person-id="${person.id}">
        <td>${person.nom}</td>
        <td>${person.prenom}</td>
        <td>${getPersonSiteMarkup(person)}</td>
        <td>${person.typePersonnel || ""}</td>
        <td>${person.typeContrat || ""}</td>
        <td>${formatDate(person.dateEntree)}</td>
        <td>${formatDate(person.dateSortiePrevue)}</td>
        <td>${formatDate(person.dateSortieReelle)}</td>
        <td>${getDossierStatusCellMarkup(getDossierStatus(person))}</td>
        <td>${totalEffects}</td>
        <td>${nonRendus > 0 ? '<span class="row-alert-dot" aria-hidden="true"></span>' : ""}${nonRendus}</td>
        <td>${totalCosts > 0 ? formatAmountWithEuro(totalCosts) : "-"}</td>
        <td>${movementMarkup || "-"}</td>
        <td>
          <a class="table-link js-open-person-link" data-person-id="${person.id}" href="fiche-personne.html?personId=${person.id}">VOIR</a>
          <button type="button" class="table-link js-delete-person" data-person-id="${person.id}">SUPPRIMER</button>
        </td>
      </tr>`;
    })
    .join("");
  const totalMovementMarkup = Object.entries(totals.movements)
    .filter(([, count]) => count > 0)
    .map(
      ([movement, count]) =>
        `<span class="movement-badge movement-badge--${getMovementBadgeVariant(movement)}">${movement} ${count}</span>`
    )
    .join(" ");
  const totalRow = `<tr class="table-total-row table-total-row--overview">
    <td>TOTAL FILTRE</td>
    <td>${totals.persons} PERSONNE${totals.persons > 1 ? "S" : ""}</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td>EN POSTE ${totals.inPost}</td>
    <td>${totals.effects}</td>
    <td>${totals.nonRendus}</td>
    <td>${totals.costs > 0 ? formatAmountWithEuro(totals.costs) : "-"}</td>
    <td>${totalMovementMarkup || "-"}</td>
    <td></td>
  </tr>`;
  return `${rows}${totalRow}`;
}

function renderGlobalTable(persons) {
  const body = document.getElementById("global-table-body");
  if (!body) {
    return false;
  }

  if (!persons.length) {
    const hasEmptyRowsChanged = state.listRenderCache.global !== "";
    state.listRenderCache.global = "";
    if (hasEmptyRowsChanged) {
      body.innerHTML = buildEmptyTableRow(body, "AUCUNE DONNEE A AFFICHER", 14);
    }
    return hasEmptyRowsChanged;
  }

  const globalRowsSignature = `global|${persons
    .map((person) => {
      const currentEffects = getCurrentAssignedEffectsForActiveFilters(person);
      const totalEffects = currentEffects.length;
      const totalCosts = currentEffects.reduce(
        (sum, effect) => sum + getEffectReplacementCost(person, effect),
        0
      );
      const nonRendus = currentEffects.filter(
        (effect) => getEffectStatus(person, effect) === "NON RENDU"
      ).length;
      const movementMap = getArrivalComplementMovementMap(person, getEffectsForActiveFilters(person));
      const movementCounts = {
        AJOUTE: 0,
        MODIFIE: 0,
        RENDU: 0,
        PERDU: 0,
        VOLE: 0,
        HS: 0,
      };
      movementMap.forEach((movement) => {
        const normalized = normalizeText(movement);
        if (Object.prototype.hasOwnProperty.call(movementCounts, normalized)) {
          movementCounts[normalized] += 1;
        }
      });
      return [
        person.id || "",
        String(person.nom || ""),
        String(person.prenom || ""),
        String(getPersonSiteMarkup(person)),
        String(person.typePersonnel || ""),
        String(person.typeContrat || ""),
        String(person.dateEntree || ""),
        String(person.dateSortiePrevue || ""),
        String(person.dateSortieReelle || ""),
        String(getDossierStatus(person)),
        String(totalEffects || 0),
        String(nonRendus || 0),
        String(totalCosts || 0),
        String(movementCounts.AJOUTE),
        String(movementCounts.MODIFIE),
        String(movementCounts.RENDU),
        String(movementCounts.PERDU),
        String(movementCounts.VOLE),
        String(movementCounts.HS),
      ].join("|");
    })
    .join("||")}`;
  const hasGlobalRowsChanged = state.listRenderCache.global !== globalRowsSignature;
  if (!hasGlobalRowsChanged) {
    return false;
  }
  state.listRenderCache.global = globalRowsSignature;

  const rowsHtml = persons
    .map((person) => {
      const currentEffects = getCurrentAssignedEffectsForActiveFilters(person);
      const totalEffects = currentEffects.length;
      const totalCosts = currentEffects.reduce(
        (sum, effect) => sum + getEffectReplacementCost(person, effect),
        0
      );
      const nonRendus = currentEffects.filter(
        (effect) => getEffectStatus(person, effect) === "NON RENDU"
      ).length;
      const movementMap = getArrivalComplementMovementMap(person, getEffectsForActiveFilters(person));
      const movementCounts = {
        AJOUTE: 0,
        MODIFIE: 0,
        RENDU: 0,
        PERDU: 0,
        VOLE: 0,
        HS: 0,
      };
      movementMap.forEach((movement) => {
        const normalized = normalizeText(movement);
        if (Object.prototype.hasOwnProperty.call(movementCounts, normalized)) {
          movementCounts[normalized] += 1;
        }
      });
      const movementMarkup = Object.entries(movementCounts)
        .filter(([, count]) => count > 0)
        .map(
          ([movement, count]) =>
            `<span class="movement-badge movement-badge--${getMovementBadgeVariant(movement)}">${movement} ${count}</span>`
        )
        .join(" ");
      const alertType = getOverdueExitAlertMeta(person).type;
      const alertClass = alertType ? ` is-alert-row is-alert-row--${alertType}` : "";
      return `<tr class="js-person-row${alertClass}" data-person-id="${person.id}">
        <td>${person.nom}</td>
        <td>${person.prenom}</td>
        <td>${getPersonSiteMarkup(person)}</td>
        <td>${person.typePersonnel}</td>
        <td>${person.typeContrat || ""}</td>
        <td>${formatDate(person.dateEntree)}</td>
        <td>${formatDate(person.dateSortiePrevue)}</td>
        <td>${formatDate(person.dateSortieReelle)}</td>
        <td>${getDossierStatusCellMarkup(getDossierStatus(person))}</td>
        <td>${totalEffects}</td>
        <td>${nonRendus > 0 ? '<span class="row-alert-dot" aria-hidden="true"></span>' : ""}${nonRendus}</td>
        <td>${totalCosts > 0 ? formatAmountWithEuro(totalCosts) : "-"}</td>
        <td>${movementMarkup || "-"}</td>
        <td>
          <a class="table-link js-open-person-link" data-person-id="${person.id}" href="fiche-personne.html?personId=${person.id}">VOIR</a>
          <button type="button" class="table-link js-delete-person" data-person-id="${person.id}">SUPPRIMER</button>
        </td>
      </tr>`;
    })
    ;

  if (hasGlobalRowsChanged) {
    renderTableRowsProgressively(body, rowsHtml, buildEmptyTableRow(body, "AUCUNE DONNEE A AFFICHER", 14), 24);
    bindPersonRowActions();
  }
  return true;
}

function getDossierStatusCellMarkup(status) {
  const normalizedStatus = normalizeText(status);
  let iconClass = "status-icon-inline status-icon-inline--pending";

  if (normalizedStatus === "EN POSTE") {
    iconClass = "status-icon-inline status-icon-inline--active";
  } else if (normalizedStatus === "SORTIE PREVUE") {
    iconClass = "status-icon-inline status-icon-inline--warning";
  } else if (normalizedStatus === "SORTI") {
    iconClass = "status-icon-inline status-icon-inline--exit";
  }

  return `<span class="status-cell"><span class="${iconClass}" aria-hidden="true"></span><span>${escapeHtml(status || "")}</span></span>`;
}

function bindPersonRowActions() {
  ["overview-table-body", "global-table-body"].forEach((bodyId) => {
    const body = document.getElementById(bodyId);
    if (!body || body.dataset.bound === "true") {
      return;
    }

    body.addEventListener("click", (event) => {
      const target = event.target;
      if (!(target instanceof HTMLElement)) {
        return;
      }
      const link = target.closest("a.js-open-person-link") || target.closest("button.js-delete-person");
      if (link instanceof HTMLElement) {
        const personId = link.getAttribute("data-person-id") || getCurrentPersonId();
        if (link.classList.contains("js-open-person-link") && personId) {
          setCurrentPersonId(personId, "replace");
        }
        return;
      }

      if (target.closest("a, button")) {
        return;
      }

      const row = target.closest(".js-person-row");
      if (!(row instanceof HTMLElement)) {
        return;
      }

      const personId = row.dataset.personId || "";
      if (!personId) {
        return;
      }

      openPersonSheet(personId);
    });
    body.addEventListener("click", (event) => {
      const target = event.target;
      if (!(target instanceof HTMLElement)) {
        return;
      }
      const deleteButton = target.closest("button.js-delete-person");
      if (!(deleteButton instanceof HTMLElement)) {
        return;
      }

      const personId = deleteButton.getAttribute("data-person-id") || getCurrentPersonId();
      if (!personId) {
        showDataStatus("AUCUNE PERSONNE SELECTIONNEE");
        return;
      }

      deletePerson(personId);
    });

    body.dataset.bound = "true";
  });
}

function bindOverviewAlertActions() {
  const alertsList = document.getElementById("overview-alerts-list");
  if (!alertsList || alertsList.dataset.bound === "true") {
    return;
  }

  alertsList.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const button = target.closest(".js-open-person-alert");
    if (!(button instanceof HTMLElement)) {
      return;
    }
    const personId = button.getAttribute("data-person-id") || "";
    if (personId) {
      openPersonSheet(personId);
    }
  });

  alertsList.dataset.bound = "true";
}

function getSheetEffectTypeIconVariant(typeEffet) {
  const normalizedType = normalizeText(typeEffet);
  if (normalizedType === "CLE CES") {
    return "cle-ces";
  }
  if (normalizedType === "BADGE INTRUSION") {
    return "badge";
  }
  if (normalizedType === "TELECOMMANDE URMET") {
    return "telecommande";
  }
  if (normalizedType === "CARTE TURBOSELF") {
    return "carte";
  }
  if (normalizedType === "RADIATEUR APPOINT") {
    return "radiateur";
  }
  if (normalizedType === "VENTILATEUR") {
    return "ventilateur";
  }
  if (["CLE", "CLE CES", "CLE DE SECURITE"].includes(normalizedType)) {
    return "cle";
  }
  return "total";
}

function getSheetEffectTypeIconSvg(typeEffet) {
  const variant = getSheetEffectTypeIconVariant(typeEffet);
  if (variant === "cle-ces") {
    return '<img src="assets/sidebar/icone-cle-ces.png?v=20260706-assets" alt="" loading="lazy">';
  }
  if (variant === "badge") {
    return '<img src="assets/sidebar/icone-badge.png?v=20260706-assets" alt="" loading="lazy">';
  }
  if (variant === "telecommande") {
    return '<img src="assets/sidebar/icone-telecommande.png?v=20260706-assets" alt="" loading="lazy">';
  }
  if (variant === "carte") {
    return '<img src="assets/sidebar/icone-carte.png?v=20260706-assets" alt="" loading="lazy">';
  }
  if (variant === "radiateur") {
    return '<img src="assets/effects/icone-radiateur-appoint.png?v=20260706-assets" alt="" loading="lazy">';
  }
  if (variant === "ventilateur") {
    return '<img src="assets/effects/icone-ventilateur.png?v=20260706-assets" alt="" loading="lazy">';
  }
  if (variant === "cle") {
    return '<img src="assets/sidebar/icone-cle.png?v=20260706-assets" alt="" loading="lazy">';
  }
  return `<svg viewBox="0 0 24 24" focusable="false">
    <rect x="5" y="5" width="14" height="14" rx="3" fill="currentColor"></rect>
    <path d="M9 12h6M12 9v6" stroke="#FBFAF7" stroke-width="2" stroke-linecap="round"></path>
  </svg>`;
}

function renderSheetEffectTypeKpis(effects) {
  const container = document.getElementById("sheet-effect-type-kpis");
  if (!container) {
    return;
  }

  const baseTypes = Array.from(new Set((state.data?.listes?.typesEffets || []).filter(Boolean)));
  container.innerHTML = baseTypes
    .map((typeEffet) => {
      const normalizedType = normalizeText(typeEffet);
      const matchingEffects = effects.filter(
        (effect) => normalizeText(effect.typeEffet) === normalizedType
      );
      const amount = matchingEffects.reduce((sum, effect) => sum + getEffectUnitValue(effect), 0);
      const variant = getSheetEffectTypeIconVariant(typeEffet);

      return `<div class="effect-type-kpi">
        <span class="effect-type-kpi__icon effect-type-kpi__icon--${variant}" aria-hidden="true">
          ${getSheetEffectTypeIconSvg(typeEffet)}
        </span>
        <div class="effect-type-kpi__content">
          <span class="effect-type-kpi__label">${escapeHtml(typeEffet)}</span>
          <strong class="effect-type-kpi__value">${matchingEffects.length}</strong>
          <span class="effect-type-kpi__amount">${formatAmountWithEuro(amount)}</span>
        </div>
      </div>`;
    })
    .join("");
}

function renderPersonSheet(personId) {
  const nameNode = document.getElementById("sheet-person-name");
  const metaNode = document.getElementById("sheet-person-meta");
  const alertNode = document.getElementById("sheet-date-alert");
  const statusNode = document.getElementById("sheet-person-status");
  const body = document.getElementById("sheet-effects-body");
  const totalNode = document.getElementById("sheet-summary-total");
  const returnedNode = document.getElementById("sheet-summary-returned");
  const missingNode = document.getElementById("sheet-summary-missing");
  const costNode = document.getElementById("sheet-summary-cost");
  const totalTypesNode = document.getElementById("sheet-kpi-total-types");
  const totalTypesAmountNode = document.getElementById("sheet-kpi-total-amount");
  if (!nameNode || !metaNode || !alertNode || !statusNode || !body) {
    return;
  }

  const requestedPersonId = String(personId || getCurrentPersonId() || state.currentSheetPersonId || "");
  const person = (state.data?.personnes || []).find(
    (entry) => String(entry?.id || "") === requestedPersonId
  );
  const requestedEditContext = getRequestedEditEffectContext();
  const requestedEditEffectId = String(requestedEditContext.effectId || "");
  const requestedEditPersonId = String(requestedEditContext.personId || "");
  const requestedEditPersonMatches = !requestedEditPersonId || requestedEditPersonId === requestedPersonId;


  if (!person) {
    const emptySheetSignature = "sheet|none";
    if (state.listRenderCache.sheet !== emptySheetSignature) {
      state.listRenderCache.sheet = emptySheetSignature;
      state.currentSheetPersonId = "";
      nameNode.textContent = "AUCUNE PERSONNE SELECTIONNEE";
      metaNode.textContent = "SELECTIONNER UNE PERSONNE POUR AFFICHER LA FICHE";
      alertNode.hidden = true;
      alertNode.textContent = "";
      applySheetPersonStatus(statusNode, "EN ATTENTE");
      body.innerHTML = buildEmptyTableRow(body, "AUCUN EFFET A AFFICHER", 11);
      fillSheetForm(null);
      if (totalNode) totalNode.textContent = "0";
      if (returnedNode) returnedNode.textContent = "0";
      if (missingNode) missingNode.textContent = "0";
      if (costNode) costNode.textContent = "0,00 €";
      renderSheetEffectTypeKpis([]);
      if (totalTypesNode) totalTypesNode.textContent = "0";
      if (totalTypesAmountNode) totalTypesAmountNode.textContent = "0,00 €";
      updateSheetDocumentButtons(null);
      hydrateEffectReferenceSiteSelect(null, "", "");
      hydrateReferenceSelect("", "", "");
      updateEffectFormMode("");
      updateManualStatusCriticalState(document.getElementById("effect-form"));
      return true;
    }
    if (requestedEditEffectId) {
      return false;
    }
    return false;
  }

  const effects = person.effetsConfies || [];
  const currentEffects = getCurrentAssignedEffects(person);
  const displayedEffects = effects;
  const sortedEffects = sortEffectsForTable(person, displayedEffects, "sheetEffects");
  const movementMap = getArrivalComplementMovementMap(person, displayedEffects);
  const returned = effects.filter((effect) => getEffectStatus(person, effect) === "RESTITUE").length;
  const missing = currentEffects.filter((effect) => getEffectStatus(person, effect) === "NON RENDU").length;
  const totalCost = currentEffects.reduce((sum, effect) => sum + getEffectReplacementCost(person, effect), 0);
  const totalEffectsUnitValue = currentEffects.reduce((sum, effect) => sum + getEffectUnitValue(effect), 0);
  const overdueMessage = getOverdueExitMessage(person);
  const rowFlash =
    state.effectRowFlash && String(state.effectRowFlash.personId || "") === String(person.id || "")
      ? state.effectRowFlash
      : null;
  const effectTableFlash =
    state.effectTableFlash &&
    String(state.effectTableFlash.personId || "") === String(person.id || "") &&
    String(state.effectTableFlash.kind || "") === "delete"
      ? state.effectTableFlash
      : null;
  const rowFlashSignature = rowFlash ? `${String(rowFlash.kind || "")}:${String(rowFlash.effectId || "")}` : "";
  const effectTableFlashSignature = effectTableFlash ? `${String(effectTableFlash.kind || "")}:${String(effectTableFlash.personId || "")}` : "";
  const sheetRowsSignature = `sheet|${String(person.id || "")}|${normalizeText(person.nom || "")}|${normalizeText(person.prenom || "")}|${normalizeText(person.typePersonnel || "")}|${normalizeText(person.typeContrat || "")}|${String(person.dateEntree || "")}|${String(person.dateSortiePrevue || "")}|${String(person.dateSortieReelle || "")}|${String(getDossierStatus(person))}|${String(overdueMessage || "")}|${String(currentEffects.length)}|${String(displayedEffects.length)}|${String(returned)}|${String(missing)}|${String(totalCost)}|${String(totalEffectsUnitValue)}|${rowFlashSignature}|${effectTableFlashSignature}|${sortedEffects
    .map((effect) => {
      const effectStatus = getEffectStatus(person, effect);
      const effectDesignation = getEffectDisplayDesignation(effect);
      const effectSite = getEffectDisplaySite(effect);
      const movement =
        movementMap.get(getEffectMovementKey(effect)) ||
        movementMap.get(getEffectStableKey(effect)) ||
        getEffectMovementLabel(person, effect);
      return [
        String(effect.id || ""),
        String(effect.typeEffet || ""),
        String(effectDesignation),
        String(effectSite),
        String(effect.numeroIdentification || ""),
        String(effect.dateRemise || ""),
        String(effect.dateRetour || ""),
        String(effect.dateRemplacement || ""),
        String(effectStatus),
        String(getEffectUnitValue(effect)),
        String(effect.commentaire || ""),
        String(movement || ""),
      ].join("|");
    })
    .join("||")}`;

  if (state.listRenderCache.sheet === sheetRowsSignature) {
    if (requestedEditPersonMatches && requestedEditEffectId && effects.some((effect) => String(effect.id || "") === requestedEditEffectId)) {
      startEditEffect(person.id, requestedEditEffectId);
    }
    return false;
  }

  state.listRenderCache.sheet = sheetRowsSignature;
  state.currentSheetPersonId = String(person.id || "");
  nameNode.textContent = `${person.nom} ${person.prenom}`;
  metaNode.innerHTML = [
    getPersonSiteMarkup(person),
    escapeHtml(person.typePersonnel || ""),
    escapeHtml(person.typeContrat || ""),
  ]
    .filter(Boolean)
    .join(' <span class="meta-separator">|</span> ');
  alertNode.hidden = !overdueMessage;
  alertNode.textContent = overdueMessage;
  applySheetPersonStatus(statusNode, getDossierStatus(person));
  updateSheetDocumentButtons(person);
  fillSheetForm(person);

  if (totalNode) totalNode.textContent = String(currentEffects.length);
  if (returnedNode) returnedNode.textContent = String(returned);
  if (missingNode) missingNode.textContent = String(missing);
  if (costNode) costNode.textContent = formatAmountWithEuro(totalCost);
  renderSheetEffectTypeKpis(currentEffects);
  if (totalTypesNode) totalTypesNode.textContent = String(currentEffects.length);
  if (totalTypesAmountNode) totalTypesAmountNode.textContent = formatAmountWithEuro(totalEffectsUnitValue);

  body.innerHTML = sortedEffects.length
    ? `${sortedEffects
        .map((effect) => {
          const effectStatus = getEffectStatus(person, effect);
          const effectDesignation = getEffectDisplayDesignation(effect);
          const effectSite = getEffectDisplaySite(effect);
          const effectUnitValue = getEffectUnitValue(effect);
          const movement =
            movementMap.get(getEffectMovementKey(effect)) ||
            movementMap.get(getEffectStableKey(effect)) ||
            getEffectMovementLabel(person, effect);
          const movementBadge = movement
            ? `<span class="movement-badge movement-badge--${getMovementBadgeVariant(movement)}">${movement}</span>`
            : "";
          const statusWithDot =
            effectStatus === "NON RENDU"
              ? `<span>${effectStatus}</span><span class="row-alert-dot row-alert-dot--inside" aria-hidden="true"></span>`
              : `<span>${effectStatus}</span>`;
          const rowFlashClass =
            rowFlash &&
            String(rowFlash.effectId || "") === String(effect.id || "") &&
            ["create", "update"].includes(String(rowFlash.kind || ""))
              ? ` row-flash row-flash--${rowFlash.kind}`
              : "";
          return `<tr class="js-effect-row${rowFlashClass}" data-person-id="${person.id}" data-effect-id="${effect.id}">
            <td>${effect.typeEffet || ""}</td>
            <td>${effectDesignation}</td>
            <td>${effectSite}</td>
            <td>${effect.numeroIdentification || ""}</td>
            <td>${formatDate(effect.dateRemise)}</td>
            <td>${formatDate(effect.dateRetour)}</td>
            <td><span class="status-text-inline">${statusWithDot}</span></td>
            <td class="movement-cell">${movementBadge}</td>
            <td>${formatDate(effect.dateRemplacement)}</td>
            <td>${formatAmountWithEuro(effectUnitValue)}</td>
            <td>${effect.commentaire || ""}</td>
          </tr>`;
        })
        .join("")}
        <tr class="table-total-row">
          <td colspan="9">TOTAL DES EFFETS CONFIES</td>
          <td>${formatAmountWithEuro(totalEffectsUnitValue)}</td>
          <td></td>
        </tr>`
    : buildEmptyTableRow(body, "AUCUN EFFET A AFFICHER", 11);

  if (rowFlash) {
    state.effectRowFlash = null;
  }
  if (effectTableFlash) {
    body.classList.remove("table-flash--delete");
    void body.offsetWidth;
    body.classList.add("table-flash--delete");
    window.setTimeout(() => {
      body.classList.remove("table-flash--delete");
    }, 420);
    state.effectTableFlash = null;
  }

  const currentTypeEffet = document.querySelector('#effect-form [name="typeEffet"]')?.value || "";
  const currentReferenceSite = document.querySelector('#effect-form [name="referenceSite"]')?.value || "";
  const currentReferenceId = document.querySelector('#effect-form [name="referenceEffet"]')?.value || "";
  hydrateEffectReferenceSiteSelect(person, currentReferenceSite, currentTypeEffet);
  hydrateReferenceSelect(person, currentTypeEffet, currentReferenceId, currentReferenceSite);
  updateEffectFormMode(currentTypeEffet);
  updateManualStatusCriticalState(document.getElementById("effect-form"));
  bindEffectRowActions();
  updateSortableHeaders("sheetEffects");

  if (requestedEditPersonMatches && requestedEditEffectId && effects.some((effect) => String(effect.id || "") === requestedEditEffectId)) {
    startEditEffect(person.id, requestedEditEffectId);
  }
  return true;
}

function getSheetPersonStatusClass(status) {
  const normalizedStatus = normalizeText(status);
  if (normalizedStatus === "EN POSTE") {
    return "status-pill status-pill--sheet status-pill--sheet-active";
  }
  if (normalizedStatus === "SORTIE PREVUE") {
    return "status-pill status-pill--sheet status-pill--sheet-warning";
  }
  if (normalizedStatus === "SORTI") {
    return "status-pill status-pill--sheet status-pill--sheet-exit";
  }
  return "status-pill status-pill--sheet status-pill--sheet-pending";
}

function applySheetPersonStatus(node, status) {
  if (!node) {
    return;
  }
  node.className = getSheetPersonStatusClass(status);
  node.textContent = status || "EN ATTENTE";
  node.setAttribute("title", status || "EN ATTENTE");
  node.setAttribute("aria-label", status || "EN ATTENTE");
}

function bindEffectRowActions() {
  const body = document.getElementById("sheet-effects-body");
  if (!body || body.dataset.bound === "true") {
    return;
  }

  body.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    if (target.closest("button, a")) {
      return;
    }

    const row = target.closest(".js-effect-row");
    if (!(row instanceof HTMLElement)) {
      return;
    }

    const personId = row.dataset.personId || "";
    const effectId = row.dataset.effectId || "";
    if (!personId || !effectId) {
      return;
    }
    startEditEffect(personId, effectId);
  });

  body.dataset.bound = "true";
}

function fillSheetForm(person) {
  const form = document.getElementById("person-sheet-form");
  const mapping = {
    sheetNom: person?.nom || "",
    sheetPrenom: person?.prenom || "",
    sheetFonction: person?.fonction || "",
    sheetTypePersonnel: person?.typePersonnel || "",
    sheetTypeContrat: person?.typeContrat || "",
    sheetDateEntree: person?.dateEntree || "",
    sheetEmail: person?.email || "",
    sheetPhoneMobile: person?.phoneMobile || "",
    sheetDateSortiePrevue: person?.dateSortiePrevue || "",
    sheetDateSortieReelle: person?.dateSortieReelle || "",
    sheetStatutDossier: person ? getDossierStatus(person) : "",
  };

  Object.entries(mapping).forEach(([name, value]) => {
    const field = document.querySelector(`[name="${name}"]`);
    if (field) {
      field.value = value;
    }
  });
  renderSiteSelector("sheet-site-selector", "sheet", getPersonSites(person));
  if (form instanceof HTMLFormElement) {
    const normalizedTypeContrat = normalizeText(form.elements.sheetTypeContrat?.value || "");
    const needsExpectedExitDate = ["CDD", "INTERIMAIRE"].includes(normalizedTypeContrat);
    const nom = normalizeText(form.elements.sheetNom?.value || "");
    const prenom = normalizeText(form.elements.sheetPrenom?.value || "");
    const fonction = normalizeText(form.elements.sheetFonction?.value || "");
    const typePersonnel = normalizeText(form.elements.sheetTypePersonnel?.value || "");
    const typeContrat = normalizeText(form.elements.sheetTypeContrat?.value || "");
    const dateEntree = String(form.elements.sheetDateEntree?.value || "").trim();
    const dateSortiePrevueValue = String(form.elements.sheetDateSortiePrevue?.value || "").trim();
    const dateSortieReelleValue = String(form.elements.sheetDateSortieReelle?.value || "").trim();
    const dateEntreeValidation = validateDateFieldFormat(dateEntree, "DATE D'ENTREE");
    const dateSortiePrevueValidation = validateDateFieldFormat(dateSortiePrevueValue, "DATE DE SORTIE PREVUE");
    const dateSortieReelleValidation = validateDateFieldFormat(dateSortieReelleValue, "DATE DE SORTIE REELLE");
    const dateSortiePrevueField = form.elements.sheetDateSortiePrevue;
    const dateSortiePrevueNode = dateSortiePrevueField instanceof HTMLElement
      ? dateSortiePrevueField.closest(".field")
      : null;
    if (dateSortiePrevueField instanceof HTMLElement) {
      dateSortiePrevueField.required = needsExpectedExitDate;
    }
    if (dateSortiePrevueNode) {
      dateSortiePrevueNode.classList.toggle("field--key", needsExpectedExitDate);
      dateSortiePrevueNode.classList.toggle(
        "field--missing",
        (needsExpectedExitDate && !dateSortiePrevueValue) || !dateSortiePrevueValidation.ok
      );
    }
    form.elements.sheetNom?.closest(".field")?.classList.toggle("field--missing", !nom);
    form.elements.sheetPrenom?.closest(".field")?.classList.toggle("field--missing", !prenom);
    form.elements.sheetFonction?.closest(".field")?.classList.toggle("field--missing", !fonction);
    form.elements.sheetTypePersonnel?.closest(".field")?.classList.toggle("field--missing", !typePersonnel);
    form.elements.sheetTypeContrat?.closest(".field")?.classList.toggle("field--missing", !typeContrat);
    form.elements.sheetDateEntree?.closest(".field")?.classList.toggle(
      "field--missing",
      !dateEntree || !dateEntreeValidation.ok
    );
    form.elements.sheetDateSortieReelle?.closest(".field")?.classList.toggle(
      "field--missing",
      !dateSortieReelleValidation.ok
    );
    const siteField = form.querySelector("#sheet-site-selector")?.closest(".field");
    if (siteField) {
      siteField.classList.toggle("field--missing", readSelectedSites(form, "sheet").length === 0);
    }
  }
}

function renderArrivalDocument(personId) {
  const person = state.data?.personnes?.find((entry) => entry.id === personId) || null;
  const explicitMode = normalizeText(new URLSearchParams(window.location.search).get("mode") || "");
  const computedMode = person ? getDocumentArchiveMode(person, "arrival") : "STANDARD";
  let mode = explicitMode || computedMode;
  let isComplement = mode === "COMPLEMENTAIRE";
  const dateNode = document.getElementById("arrival-doc-date");
  const referenceNode = document.getElementById("arrival-doc-reference");
  const titleNode = document.getElementById("arrival-doc-title");
  const subtitleNode = document.getElementById("arrival-doc-subtitle");
  const nomNode = document.getElementById("arrival-person-nom");
  const prenomNode = document.getElementById("arrival-person-prenom");
  const fonctionNode = document.getElementById("arrival-person-fonction");
  const typePersonnelNode = document.getElementById("arrival-person-type-personnel");
  const typeContratNode = document.getElementById("arrival-person-type-contrat");
  const sitesNode = document.getElementById("arrival-person-sites");
  const dateEntreeNode = document.getElementById("arrival-person-date-entree");
  const dateSortiePrevueNode = document.getElementById("arrival-person-date-sortie-prevue");
  const body = document.getElementById("arrival-effects-body");
  const totalEffectsNode = document.getElementById("arrival-total-effects");
  const totalValueNode = document.getElementById("arrival-total-value");
  const signatureNameNode = document.getElementById("arrival-signature-person-name");
  const signaturePersonDateNode = document.getElementById("arrival-signature-person-date");
  const signatureRepresentantDateNode = document.getElementById("arrival-signature-representant-date");
  const representantNameInput = document.getElementById("arrival-signature-representant-name-input");
  const representantFunctionInput = document.getElementById("arrival-signature-representant-function-input");
  const representantNameNode = document.getElementById("arrival-signature-representant-name");
  const representantFunctionNode = document.getElementById("arrival-signature-representant-function");
  const costsHead = document.getElementById("arrival-costs-head");
  const costsBody = document.getElementById("arrival-costs-body");

  if (
    !dateNode ||
    !referenceNode ||
    !titleNode ||
    !subtitleNode ||
    !nomNode ||
    !prenomNode ||
    !fonctionNode ||
    !typePersonnelNode ||
    !typeContratNode ||
    !sitesNode ||
    !dateEntreeNode ||
    !dateSortiePrevueNode ||
    !body ||
    !totalEffectsNode ||
    !totalValueNode ||
    !signatureNameNode ||
    !signaturePersonDateNode ||
    !signatureRepresentantDateNode ||
    !representantNameInput ||
    !representantFunctionInput ||
    !representantNameNode ||
    !representantFunctionNode ||
    !costsHead ||
    !costsBody
  ) {
    return;
  }
  const isPdfMode = isPdfRenderMode();
  const sortConfig = state.tableSorts?.arrivalEffects || {};
  const noPersonRenderKey = [
    "arrival",
    "no-person",
    String(explicitMode || ""),
    isPdfMode ? "pdf" : "ui",
    String(sortConfig.key || ""),
    String(sortConfig.dir || ""),
  ].join("|");

  if (!person) {
    if (state.documentRenderCache.arrival === noPersonRenderKey) {
      return false;
    }
    state.documentRenderCache.arrival = noPersonRenderKey;
    titleNode.textContent = isComplement
      ? "AVENANT DE REMISE DES EFFETS CONFIES"
      : "ATTESTATION DE REMISE DES EFFETS CONFIES A L'ARRIVEE";
    subtitleNode.textContent = isComplement
      ? "COMPLEMENT DE DOTATION APRES DOCUMENT D'ARRIVEE SIGNE"
      : "DOCUMENT DE REMISE DES EFFETS CONFIES ET ACCEPTATION DES CONDITIONS DE RESTITUTION";
    dateNode.textContent = formatDateTimeForDocument("");
    referenceNode.textContent = "-";
    nomNode.textContent = "-";
    prenomNode.textContent = "-";
    fonctionNode.textContent = "-";
    typePersonnelNode.textContent = "-";
    typeContratNode.textContent = "-";
    sitesNode.textContent = "-";
    dateEntreeNode.textContent = "-";
    dateSortiePrevueNode.textContent = "-";
    body.innerHTML = buildEmptyTableRow(body, "AUCUN EFFET A AFFICHER", 7);
    renderArrivalCostsTable(costsHead, costsBody);
    totalEffectsNode.textContent = "0";
    totalValueNode.textContent = "0,00 €";
    signatureNameNode.textContent = "-";
    signaturePersonDateNode.textContent = "-";
    signatureRepresentantDateNode.textContent = "-";
    populateRepresentativeSelect(representantNameInput, "");
    representantFunctionInput.value = "";
    representantNameNode.textContent = "-";
    representantFunctionNode.textContent = "-";
    updateRepresentativeSignatureActionState("arrival");
      syncDocumentMobileSignatureLinks("arrival", "");
    return;
  }

  const allEffects = Array.isArray(person.effetsConfies) ? person.effetsConfies : [];
  const fallbackMovements = getArrivalComplementMovementMap(person, allEffects);

  if (!explicitMode && !isComplement && fallbackMovements.size && isPdfMode) {
    mode = "COMPLEMENTAIRE";
    isComplement = true;
  }

  const activeEffects = isComplement
    ? allEffects
    : allEffects.filter((effect) => Boolean(effect.dateRemise));
  const deletedEffects = isComplement ? getArrivalDeletedEffects(person, allEffects) : [];
  const effectsForDisplay = [...activeEffects, ...deletedEffects];
  const complementMovements = isComplement ? fallbackMovements : new Map();
  const arrivalRowsSignature = effectsForDisplay
    .map((effect) => {
      const movement = getEffectMovementLabel(person, effect, isComplement ? complementMovements : null);
      return `${String(effect.id || "")}|${String(effect.typeEffet || "")}|${String(getEffectDisplayDesignation(effect) || "")}|${String(
        effect.numeroIdentification || ""
      )}|${String(effect.dateRemise || "")}|${getEffectUnitValue(effect)}|${movement || ""}`;
    })
    .join("||");
  const totalValue = activeEffects.reduce((sum, effect) => sum + getEffectUnitValue(effect), 0);
  const arrivalPersonnelSignature = String(getSignatureValidationDate(person, "arrival", "personnel") || "");
  const arrivalRepresentantSignature = String(getSignatureValidationDate(person, "arrival", "representant") || "");
  const arrivalRepresentative = getRepresentativeInfo(person, "arrival");
  const personSummaryFingerprint = [
    "arrival",
    String(person.id || ""),
    String(mode || ""),
    isPdfMode ? "pdf" : "ui",
    String(sortConfig.key || ""),
    String(sortConfig.dir || ""),
    String(person.nom || ""),
    String(person.prenom || ""),
    String(person.fonction || ""),
    String(person.typePersonnel || ""),
    String(person.typeContrat || ""),
    String(person.dateEntree || ""),
    String(person.dateSortiePrevue || ""),
    String(arrivalRepresentative.id || ""),
    String(arrivalRepresentative.nom || ""),
    String(arrivalRepresentative.fonction || ""),
    String(arrivalPersonnelSignature),
    String(arrivalRepresentantSignature),
    String(effectsForDisplay.length),
    String(activeEffects.length),
    String(totalValue),
    String(totalValueNode?.textContent || ""),
    String(arrivalRowsSignature || ""),
  ].join("|");

  if (state.documentRenderCache.arrival === personSummaryFingerprint) {
    return false;
  }
  state.documentRenderCache.arrival = personSummaryFingerprint;
  const sortedEffects = sortEffectsForTable(person, effectsForDisplay, "arrivalEffects");
  const arrivalRowsMarkup = sortedEffects
    .map((effect) => {
      const movement = getEffectMovementLabel(person, effect, isComplement ? complementMovements : null);
      const movementBadge = movement
        ? `<span class="movement-badge movement-badge--${getMovementBadgeVariant(movement)}">${movement}</span>`
        : "";
      const rowClass = movement
        ? ` class="arrival-effect-row arrival-effect-row--${getMovementRowVariant(movement)}"`
        : "";
      const designation = getEffectDisplayDesignation(effect);
      const unitValue = getEffectUnitValue(effect);
      const actionCell = !isPdfMode
        ? `<span class="document-effect-actions">
                  <button type="button" class="table-link js-doc-edit-effect" data-person-id="${escapeHtml(person.id || "")}" data-effect-id="${escapeHtml(effect.id || "")}">MODIFIER</button>
                  <button type="button" class="table-link js-doc-delete-effect" data-person-id="${escapeHtml(person.id || "")}" data-effect-id="${escapeHtml(effect.id || "")}">SUPPRIMER</button>
                </span>`
        : "-";
      return {
        id: String(effect.id || ""),
        typeEffet: String(effect.typeEffet || ""),
        designation: String(designation || "-"),
        numeroIdentification: String(effect.numeroIdentification || ""),
        dateRemise: String(effect.dateRemise || ""),
        unitValue,
        movementBadge,
        rowClass,
        actionCell,
      };
    })
    .map(
      ({ id, typeEffet, designation, numeroIdentification, dateRemise, unitValue, movementBadge, rowClass, actionCell }) => `
        <tr${rowClass}>
          <td>${typeEffet || ""}</td>
          <td>${designation || "-"}</td>
          <td class="movement-cell">${movementBadge}</td>
          <td>${numeroIdentification || "-"}</td>
          <td>${formatDate(dateRemise) || "-"}</td>
          <td>${formatAmountWithEuro(unitValue)}</td>
          <td class="document-effects-action-col">${actionCell}</td>
        </tr>
      `
    )
    .join("");

  titleNode.textContent = isComplement
    ? "AVENANT DE REMISE DES EFFETS CONFIES"
    : "ATTESTATION DE REMISE DES EFFETS CONFIES A L'ARRIVEE";
  subtitleNode.textContent = isComplement
    ? "COMPLEMENT DE DOTATION APRES DOCUMENT D'ARRIVEE SIGNE"
    : "DOCUMENT DE REMISE DES EFFETS CONFIES ET ACCEPTATION DES CONDITIONS DE RESTITUTION";
  dateNode.textContent = formatDateTimeForDocument(new Date().toISOString());
  referenceNode.textContent = `${isComplement ? "AVD" : "ARR"}-${person.id || "-"}`;
  nomNode.textContent = person.nom || "-";
  prenomNode.textContent = person.prenom || "-";
  fonctionNode.textContent = person.fonction || "-";
  typePersonnelNode.textContent = person.typePersonnel || "-";
  typeContratNode.textContent = person.typeContrat || "-";
  sitesNode.textContent = getPersonSiteLabel(person) || "-";
  dateEntreeNode.textContent = formatDate(person.dateEntree) || "-";
  dateSortiePrevueNode.textContent = formatDate(person.dateSortiePrevue) || "-";
  signatureNameNode.textContent = `${person.nom || ""} ${person.prenom || ""}`.trim() || "-";
  populateRepresentativeSelect(representantNameInput, resolveRepresentativeSelectedId(person, "arrival"));
  representantFunctionInput.value = arrivalRepresentative.fonction;
  representantNameNode.textContent = arrivalRepresentative.nom || "-";
  representantFunctionNode.textContent = arrivalRepresentative.fonction || "-";
  updateRepresentativeSignatureActionState("arrival");
  const arrivalPersonnelValidation = arrivalPersonnelSignature ? formatSignatureTimestamp(arrivalPersonnelSignature) : "";
  const arrivalRepresentantValidation = arrivalRepresentantSignature
    ? formatSignatureTimestamp(arrivalRepresentantSignature)
    : "";
  signaturePersonDateNode.textContent =
    arrivalPersonnelValidation || formatCurrentUiTimestamp();
  signatureRepresentantDateNode.textContent =
    arrivalRepresentantValidation || formatCurrentUiTimestamp();

  body.innerHTML = sortedEffects.length
    ? `${arrivalRowsMarkup}
        <tr class="table-total-row">
          <td colspan="${isPdfMode ? "5" : "6"}">TOTAL DES EFFETS REMIS</td>
          <td>${formatAmountWithEuro(totalValue)}</td>
        </tr>`
    : buildEmptyTableRow(body, "AUCUN EFFET A AFFICHER", 7);

  renderArrivalCostsTable(costsHead, costsBody);
  totalEffectsNode.textContent = String(activeEffects.length);
  totalValueNode.textContent = formatAmountWithEuro(totalValue);
  bindDocumentEffectActions();
  updateSortableHeaders("arrivalEffects");
  syncDocumentMobileSignatureLinks("arrival", person.id);
  applyRequestedPdfFocus();
  return true;
}

function renderArrivalCostsTable(headNode, bodyNode) {
  renderDocumentCostsTable("arrival", headNode, bodyNode);
}

function getArrivalCostTypes() {
  const types = Array.isArray(state.data?.listes?.typesEffets) ? state.data.listes.typesEffets : [];
  return types
    .map(normalizeText)
    .filter(Boolean)
    .filter((value, index, array) => array.indexOf(value) === index);
}

function getArrivalCostCauses() {
  return getReferenceCauseOptions();
}

function getArrivalCostDesignation(typeEffet) {
  return normalizeText(typeEffet) === "CLE CES" ? "CES-PG" : "";
}

function getDocumentCostRows() {
  const effectTypes = getArrivalCostTypes();
  const rows = [];
  effectTypes.forEach((typeEffet) => {
    const normalizedType = normalizeText(typeEffet);
    if (normalizedType === "VENTILATEUR") {
      rows.push({
        typeEffet: normalizedType,
        designation: "VENTILATEUR BUREAU",
        label: "VENTILATEUR BUREAU",
      });
      rows.push({
        typeEffet: normalizedType,
        designation: "VENTILATEUR SUR PIED",
        label: "VENTILATEUR SUR PIED",
      });
      return;
    }
    rows.push({
      typeEffet: normalizedType,
      designation: getArrivalCostDesignation(normalizedType),
      label: normalizedType,
    });
  });
  return rows;
}

function renderDocumentCostsTable(docType, headNode, bodyNode) {
  const normalizedDocType = normalizeText(docType) === "EXIT" ? "exit" : "arrival";
  const causes = getArrivalCostCauses();
  const costRowDefinitions = getDocumentCostRows();
  const costRows = costRowDefinitions
    .map((row) => {
      const values = causes.map((cause) => getReplacementCostValue(row.typeEffet, cause, row.designation));
      return `${String(row.label)}|||${String(row.typeEffet)}|||${String(row.designation)}|||${values.join("||")}`;
    })
    .join(";;");
  const nextCostSignature = [
    normalizedDocType,
    String(causes.join("||")),
    String(costRows),
  ].join("##");

  if (state.documentCostRenderCache?.[normalizedDocType] === nextCostSignature) {
    return;
  }
  state.documentCostRenderCache[normalizedDocType] = nextCostSignature;

  headNode.innerHTML = `<tr>
    <th>TYPE D'EFFET</th>
    ${causes.map((cause) => `<th>${escapeHtml(cause)}</th>`).join("")}
  </tr>`;

  if (!costRowDefinitions.length) {
    bodyNode.innerHTML = buildEmptyTableRow(bodyNode, "AUCUN COUT A AFFICHER", causes.length + 1);
    return;
  }

  bodyNode.innerHTML = costRowDefinitions
    .map(
      (row) => `<tr>
        <td>${escapeHtml(row.label)}</td>
        ${causes
          .map((cause) => `<td>${formatAmountWithEuro(getReplacementCostValue(row.typeEffet, cause, row.designation))}</td>`)
          .join("")}
      </tr>`
    )
    .join("");
}

function renderExitDocument(personId) {
  const person = state.data?.personnes?.find((entry) => entry.id === personId) || null;
  const dateNode = document.getElementById("exit-doc-date");
  const referenceNode = document.getElementById("exit-doc-reference");
  const nomNode = document.getElementById("exit-person-nom");
  const prenomNode = document.getElementById("exit-person-prenom");
  const fonctionNode = document.getElementById("exit-person-fonction");
  const typePersonnelNode = document.getElementById("exit-person-type-personnel");
  const typeContratNode = document.getElementById("exit-person-type-contrat");
  const sitesNode = document.getElementById("exit-person-sites");
  const dateEntreeNode = document.getElementById("exit-person-date-entree");
  const dateSortiePrevueNode = document.getElementById("exit-person-date-sortie-prevue");
  const dateSortieReelleNode = document.getElementById("exit-person-date-sortie-reelle");
  const body = document.getElementById("exit-effects-body");
  const totalEffectsNode = document.getElementById("exit-total-effects");
  const totalReturnedNode = document.getElementById("exit-total-returned");
  const totalChargeableNode = document.getElementById("exit-total-chargeable");
  const totalValueNode = document.getElementById("exit-total-value");
  const totalBillingPendingNode = document.getElementById("exit-total-billing-pending");
  const totalBillingBilledNode = document.getElementById("exit-total-billing-billed");
  const totalBillingClosedNode = document.getElementById("exit-total-billing-closed");
  const totalBilledValueNode = document.getElementById("exit-total-billed-value");
  const totalRemainingValueNode = document.getElementById("exit-total-remaining-value");
  const signatureNameNode = document.getElementById("exit-signature-person-name");
  const signaturePersonDateNode = document.getElementById("exit-signature-person-date");
  const signatureRepresentantDateNode = document.getElementById("exit-signature-representant-date");
  const representantNameInput = document.getElementById("exit-signature-representant-name-input");
  const representantFunctionInput = document.getElementById("exit-signature-representant-function-input");
  const representantNameNode = document.getElementById("exit-signature-representant-name");
  const representantFunctionNode = document.getElementById("exit-signature-representant-function");
  const costsHead = document.getElementById("exit-costs-head");
  const costsBody = document.getElementById("exit-costs-body");

  if (
    !dateNode ||
    !referenceNode ||
    !nomNode ||
    !prenomNode ||
    !fonctionNode ||
    !typePersonnelNode ||
    !typeContratNode ||
    !sitesNode ||
    !dateEntreeNode ||
    !dateSortiePrevueNode ||
    !dateSortieReelleNode ||
    !body ||
    !totalEffectsNode ||
    !totalReturnedNode ||
    !totalChargeableNode ||
    !totalValueNode ||
    !totalBillingPendingNode ||
    !totalBillingBilledNode ||
    !totalBillingClosedNode ||
    !totalBilledValueNode ||
    !totalRemainingValueNode ||
    !signatureNameNode ||
    !signaturePersonDateNode ||
    !signatureRepresentantDateNode ||
    !representantNameInput ||
    !representantFunctionInput ||
    !representantNameNode ||
    !representantFunctionNode ||
    !costsHead ||
    !costsBody
  ) {
    return;
  }
  const isPdfMode = isPdfRenderMode();
  const sortConfig = state.tableSorts?.exitEffects || {};
  const noPersonRenderKey = [
    "exit",
    "no-person",
    isPdfMode ? "pdf" : "ui",
    String(sortConfig.key || ""),
    String(sortConfig.dir || ""),
  ].join("|");

  if (!person) {
    if (state.documentRenderCache.exit === noPersonRenderKey) {
      return false;
    }
    state.documentRenderCache.exit = noPersonRenderKey;
    dateNode.textContent = formatDateTimeForDocument("");
    referenceNode.textContent = "-";
    nomNode.textContent = "-";
    prenomNode.textContent = "-";
    fonctionNode.textContent = "-";
    typePersonnelNode.textContent = "-";
    typeContratNode.textContent = "-";
    sitesNode.textContent = "-";
    dateEntreeNode.textContent = "-";
    dateSortiePrevueNode.textContent = "-";
    dateSortieReelleNode.textContent = "-";
    body.innerHTML = buildEmptyTableRow(body, "AUCUN EFFET A AFFICHER", 11);
    totalEffectsNode.textContent = "0";
    totalReturnedNode.textContent = "0";
    totalChargeableNode.textContent = "0";
    totalValueNode.textContent = "0,00 €";
    totalBillingPendingNode.textContent = "0";
    totalBillingBilledNode.textContent = "0";
    totalBillingClosedNode.textContent = "0";
    totalBilledValueNode.textContent = "0,00 €";
    totalRemainingValueNode.textContent = "0,00 €";
    signatureNameNode.textContent = "-";
    signaturePersonDateNode.textContent = "-";
    signatureRepresentantDateNode.textContent = "-";
    populateRepresentativeSelect(representantNameInput, "");
    representantFunctionInput.value = "";
    representantNameNode.textContent = "-";
    representantFunctionNode.textContent = "-";
    renderDocumentCostsTable("arrival", costsHead, costsBody);
    updateRepresentativeSignatureActionState("exit");
    syncDocumentMobileSignatureLinks("exit", "");
    return;
  }

  const effects = (person.effetsConfies || []).filter((effect) => {
    const hasType = Boolean(normalizeText(effect?.typeEffet));
    const hasDesignation = Boolean(normalizeText(getEffectDisplayDesignation(effect)));
    const hasId = Boolean(normalizeText(effect?.numeroIdentification));
    const hasDateRemise = Boolean(String(effect?.dateRemise || "").trim());
    const hasDateRetour = Boolean(String(effect?.dateRetour || "").trim());
    const hasStatus = Boolean(normalizeText(getEffectStatus(person, effect)));
    const hasAmount = getEffectReplacementCost(person, effect) > 0;
    return hasType || hasDesignation || hasId || hasDateRemise || hasDateRetour || hasStatus || hasAmount;
  });
  const exitRowsSignature = effects
    .map((effect) => {
      const movement = getEffectMovementLabel(person, effect);
      const replacementCost = getEffectReplacementCost(person, effect);
      const rawStatus = getEffectStatus(person, effect);
      const currentStatus = normalizeText(rawStatus);
      const billingStatus = getEffectBillingStatus(effect, replacementCost > 0);
      return `${String(effect.id || "")}|${String(effect.typeEffet || "")}|${String(getEffectDisplayDesignation(effect) || "")}|${String(effect.numeroIdentification || "")}|${String(effect.dateRemise || "")}|${String(effect.dateRetour || "")}|${currentStatus}|${replacementCost}|${billingStatus}|${movement || ""}`;
    })
    .join("||");
  const totalReturned = effects.filter((effect) => getEffectStatus(person, effect) === "RESTITUE").length;
  const chargeableEffects = effects.filter((effect) => isEffectChargeable(person, effect));
  const totalValue = chargeableEffects.reduce((sum, effect) => sum + getEffectReplacementCost(person, effect), 0);
  const billingPendingCount = chargeableEffects.filter(
    (effect) => getEffectBillingStatus(effect, true) === "A FACTURER"
  ).length;
  const billingBilledCount = chargeableEffects.filter(
    (effect) => getEffectBillingStatus(effect, true) === "FACTURE"
  ).length;
  const billingClosedCount = chargeableEffects.filter(
    (effect) => getEffectBillingStatus(effect, true) === "CLOTURE"
  ).length;
  const totalPendingValue = chargeableEffects
    .filter((effect) => getEffectBillingStatus(effect, true) === "A FACTURER")
    .reduce((sum, effect) => sum + getEffectReplacementCost(person, effect), 0);
  const totalClosedValue = chargeableEffects
    .filter((effect) => getEffectBillingStatus(effect, true) === "CLOTURE")
    .reduce((sum, effect) => sum + getEffectReplacementCost(person, effect), 0);
  const totalBilledValue = chargeableEffects
    .filter((effect) => getEffectBillingStatus(effect, true) === "FACTURE")
    .reduce((sum, effect) => sum + getEffectReplacementCost(person, effect), 0);
  const totalRemainingValue = chargeableEffects
    .filter((effect) => getEffectBillingStatus(effect, true) === "A FACTURER")
    .reduce((sum, effect) => sum + getEffectReplacementCost(person, effect), 0);
  const todayIso = getTodayIsoDate();
  const exitPersonnelSignature = String(getSignatureValidationDate(person, "exit", "personnel") || "");
  const exitRepresentantSignature = String(getSignatureValidationDate(person, "exit", "representant") || "");
  const exitRepresentative = getRepresentativeInfo(person, "exit");
  const personSummaryFingerprint = [
    "exit",
    String(person.id || ""),
    isPdfMode ? "pdf" : "ui",
    String(sortConfig.key || ""),
    String(sortConfig.dir || ""),
    String(person.nom || ""),
    String(person.prenom || ""),
    String(person.fonction || ""),
    String(person.typePersonnel || ""),
    String(person.typeContrat || ""),
    String(person.dateEntree || ""),
    String(person.dateSortiePrevue || ""),
    String(person.dateSortieReelle || ""),
    String(exitRepresentative.id || ""),
    String(exitRepresentative.nom || ""),
    String(exitRepresentative.fonction || ""),
    String(exitPersonnelSignature),
    String(exitRepresentantSignature),
    String(todayIso),
    String(effects.length),
    String(totalReturned),
    String(totalChargeableNode ? chargeableEffects.length : 0),
    String(totalValue),
    String(totalPendingValue),
    String(totalBilledValue),
    String(totalClosedValue),
    String(totalRemainingValue),
    String(exitRowsSignature || ""),
  ].join("|");
  if (state.documentRenderCache.exit === personSummaryFingerprint) {
    return false;
  }
  state.documentRenderCache.exit = personSummaryFingerprint;
  const sortedEffects = sortEffectsForTable(person, effects, "exitEffects");
  const exitRowsMarkup = sortedEffects
    .map((effect) => {
      const movement = getEffectMovementLabel(person, effect);
      const movementBadge = movement
        ? `<span class="movement-badge movement-badge--${getMovementBadgeVariant(movement)}">${movement}</span>`
        : "";
      const rawStatus = getEffectStatus(person, effect);
      const currentStatus = normalizeText(rawStatus);
      const statusLabel = currentStatus === "RESTITUE" ? "RENDU" : rawStatus;
      const replacementCost = getEffectReplacementCost(person, effect);
      const billingStatus = getEffectBillingStatus(effect, replacementCost > 0);
      const billingControls = !isPdfMode && replacementCost > 0
        ? `<label class="return-today-toggle"><input type="checkbox" class="js-exit-billed" data-effect-id="${escapeHtml(effect.id || "")}" ${billingStatus === "FACTURE" ? "checked" : ""} /><span>FACTURE</span></label>
                 <label class="return-today-toggle"><input type="checkbox" class="js-exit-closed" data-effect-id="${escapeHtml(effect.id || "")}" ${billingStatus === "CLOTURE" ? "checked" : ""} /><span>CLOTURE</span></label>`
        : "";
      const billingCell =
        billingStatus === "-"
          ? "-"
          : `<span class="${getStatusClass(billingStatus)}">${billingStatus}</span>${billingControls ? `<div class="exit-billing-toggles">${billingControls}</div>` : ""}`;
      const retourDateIso = normalizeDateString(effect.dateRetour || "");
      const isLockedStatusForReturn = ["PERDU", "HS", "VOL", "DETRUIT", "NON RENDU"].includes(currentStatus);
      const canToggleReturnToday =
        !isLockedStatusForReturn &&
        (!retourDateIso || retourDateIso === todayIso);
      const actionCell = !isPdfMode
        ? `<span class="document-effect-actions">
                  <button type="button" class="table-link js-doc-edit-effect" data-person-id="${escapeHtml(person.id || "")}" data-effect-id="${escapeHtml(effect.id || "")}">MODIFIER</button>
                  <button type="button" class="table-link js-doc-delete-effect" data-person-id="${escapeHtml(person.id || "")}" data-effect-id="${escapeHtml(effect.id || "")}">SUPPRIMER</button>
                </span>`
        : "-";
      return {
        id: String(effect.id || ""),
        typeEffet: String(effect.typeEffet || ""),
        designation: String(getEffectDisplayDesignation(effect) || ""),
        numeroIdentification: String(effect.numeroIdentification || ""),
        dateRemise: String(effect.dateRemise || ""),
        dateRetour: String(effect.dateRetour || ""),
        statusLabel: String(statusLabel || ""),
        movement,
        movementBadge,
        billingCell,
        canToggleReturnToday,
        retourDateIso,
        replacementCost,
        actionCell,
      };
    })
    .map(
      ({
        id,
        typeEffet,
        designation,
        numeroIdentification,
        dateRemise,
        dateRetour,
        statusLabel,
        movementBadge,
        billingCell,
        canToggleReturnToday,
        retourDateIso,
        replacementCost,
        actionCell,
      }) => `
        <tr>
          <td>${typeEffet || ""}</td>
          <td>${designation}</td>
          <td>${numeroIdentification}</td>
          <td>${formatDate(dateRemise)}</td>
          <td>${formatDate(dateRetour)}</td>
          <td>${statusLabel}</td>
          <td class="movement-cell">${movementBadge}</td>
          <td class="movement-cell">${
            !isPdfMode && canToggleReturnToday
              ? `<label class="return-today-toggle"><input type="checkbox" class="js-exit-return-today" data-effect-id="${escapeHtml(id || "")}" ${String(retourDateIso) === todayIso ? "checked" : ""} /><span>RENDU</span></label>`
              : "-"
          }</td>
          <td>${formatAmountWithEuro(replacementCost)}</td>
          <td>${billingCell}</td>
          <td class="document-effects-action-col">${actionCell}</td>
        </tr>
      `
    )
    .join("");
  dateNode.textContent = formatDateTimeForDocument(new Date().toISOString());
  referenceNode.textContent = `SOR-${person.id || "-"}`;
  nomNode.textContent = person.nom || "-";
  prenomNode.textContent = person.prenom || "-";
  fonctionNode.textContent = person.fonction || "-";
  typePersonnelNode.textContent = person.typePersonnel || "-";
  typeContratNode.textContent = person.typeContrat || "-";
  sitesNode.textContent = getPersonSiteLabel(person) || "-";
  dateEntreeNode.textContent = formatDate(person.dateEntree) || "-";
  dateSortiePrevueNode.textContent = formatDate(person.dateSortiePrevue) || "-";
  dateSortieReelleNode.textContent = formatDate(person.dateSortieReelle) || "-";
  signatureNameNode.textContent = `${person.nom || ""} ${person.prenom || ""}`.trim() || "-";
  populateRepresentativeSelect(representantNameInput, resolveRepresentativeSelectedId(person, "exit"));
  representantFunctionInput.value = exitRepresentative.fonction;
  representantNameNode.textContent = exitRepresentative.nom || "-";
  representantFunctionNode.textContent = exitRepresentative.fonction || "-";
  updateRepresentativeSignatureActionState("exit");
  signaturePersonDateNode.textContent =
    formatSignatureTimestamp(exitPersonnelSignature) || formatCurrentUiTimestamp();
  signatureRepresentantDateNode.textContent =
    formatSignatureTimestamp(exitRepresentantSignature) || formatCurrentUiTimestamp();

  body.innerHTML = sortedEffects.length
    ? `${exitRowsMarkup}
        <tr class="table-total-row">
          <td colspan="${isPdfMode ? "7" : "10"}">TOTAL FACTURABLE DES EFFETS</td>
          <td>${formatAmountWithEuro(totalValue)}</td>
        </tr>`
    : buildEmptyTableRow(body, "AUCUN EFFET A AFFICHER", 11);

  totalEffectsNode.textContent = String(effects.length);
  totalReturnedNode.textContent = String(totalReturned);
  totalChargeableNode.textContent = String(chargeableEffects.length);
  totalValueNode.textContent = formatAmountWithEuro(totalValue);
  totalBillingPendingNode.textContent = formatAmountWithEuro(totalPendingValue);
  totalBillingBilledNode.textContent = formatAmountWithEuro(totalBilledValue);
  totalBillingClosedNode.textContent = formatAmountWithEuro(totalClosedValue);
  totalBilledValueNode.textContent = formatAmountWithEuro(totalBilledValue);
  totalRemainingValueNode.textContent = formatAmountWithEuro(totalRemainingValue);
  renderDocumentCostsTable("exit", costsHead, costsBody);
  bindExitReturnTodayToggles();
  bindExitBillingToggles();
  bindDocumentEffectActions();
  updateSortableHeaders("exitEffects");
  syncDocumentMobileSignatureLinks("exit", person.id);
  applyRequestedPdfFocus();
  return true;
}

function bindDocumentEffectActions() {
  const page = document.body.dataset.page || "";
  const bodyId =
    page === "arrival-document"
      ? "arrival-effects-body"
      : page === "exit-document"
        ? "exit-effects-body"
        : "";
  if (!bodyId) {
    return;
  }
  const body = document.getElementById(bodyId);
  if (!body || body.dataset.effectActionsBound === "true") {
    return;
  }

  body.addEventListener("click", async (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    const editButton = target.closest(".js-doc-edit-effect");
    if (editButton instanceof HTMLElement) {
      const personId = String(editButton.dataset.personId || "");
      const effectId = String(editButton.dataset.effectId || "");
      openPersonSheetEffectEditor(personId, effectId);
      return;
    }

    const deleteButton = target.closest(".js-doc-delete-effect");
    if (deleteButton instanceof HTMLElement) {
      const personId = String(deleteButton.dataset.personId || "");
      const effectId = String(deleteButton.dataset.effectId || "");
      if (!personId || !effectId) {
        return;
      }
      await deleteEffect(personId, effectId);
    }
  });

  body.dataset.effectActionsBound = "true";
}

function bindExitReturnTodayToggles() {
  const body = document.getElementById("exit-effects-body");
  if (!body) {
    return;
  }
  if (body.dataset.returnBound === "true") {
    return;
  }
  body.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement) || !target.classList.contains("js-exit-return-today")) {
      return;
    }
    const person = getCurrentPerson();
    if (!person) {
      return;
    }
    const effectId = String(target.dataset.effectId || "");
    const effect = (person.effetsConfies || []).find((entry) => String(entry.id || "") === effectId);
    if (!effect) {
      return;
    }
    const effectStatus = normalizeText(getEffectStatus(person, effect));
    if (["PERDU", "HS", "VOL", "DETRUIT", "NON RENDU"].includes(effectStatus)) {
      target.checked = false;
      showDataStatus("RETOUR DU JOUR IMPOSSIBLE POUR CE STATUT");
      return;
    }
    const wasReturned = Boolean(String(effect.dateRetour || "").trim());
    effect.dateRetour = target.checked ? getTodayIsoDate() : "";
    effect.statutManuel = getStoredManualStatusForEffect(effect.statutManuel, effect.dateRetour);
    const isReturned = Boolean(String(effect.dateRetour || "").trim());
    if (isReturned !== wasReturned) {
    addAutoStockMovement(person, effect, isReturned ? "ENTREE" : "SORTIE", isReturned ? "RETOUR" : "ANNULATION_RETOUR");
    }
    markDirty();
    schedulePageRender();
    showActionStatus("update", target.checked ? "EFFET MARQUE RENDU CE JOUR" : "RETOUR DU JOUR ANNULE");
  });
  body.dataset.returnBound = "true";
}

function bindExitBillingToggles() {
  const body = document.getElementById("exit-effects-body");
  if (!body) {
    return;
  }
  if (body.dataset.billingBound === "true") {
    return;
  }
  body.addEventListener("change", (event) => {
    const target = event.target;
    if (
      !(target instanceof HTMLInputElement) ||
      (!target.classList.contains("js-exit-billed") && !target.classList.contains("js-exit-closed"))
    ) {
      return;
    }
    const person = getCurrentPerson();
    if (!person) {
      return;
    }
    const effectId = String(target.dataset.effectId || "");
    const effect = (person.effetsConfies || []).find((entry) => String(entry.id || "") === effectId);
    if (!effect) {
      return;
    }
    const isBilledToggle = target.classList.contains("js-exit-billed");
    const row = target.closest("tr");
    const oppositeSelector = isBilledToggle ? ".js-exit-closed" : ".js-exit-billed";
    const oppositeToggle = row instanceof HTMLElement ? row.querySelector(oppositeSelector) : null;
    if (!target.checked) {
      // If user selected the opposite checkbox, ignore this uncheck event to avoid clearing the new state.
      if (oppositeToggle instanceof HTMLInputElement && oppositeToggle.checked) {
        return;
      }
      effect.etatFacturation = "";
      markDirty();
      renderExitDocument(person.id);
      showActionStatus("update", "ETAT FACTURATION REMIS A A FACTURER");
      return;
    }
    effect.etatFacturation = isBilledToggle ? "FACTURE" : "CLOTURE";
    markDirty();
    renderExitDocument(person.id);
    showActionStatus("update", isBilledToggle ? "ETAT FACTURATION: FACTURE" : "ETAT FACTURATION: CLOTURE");
  });
  body.dataset.billingBound = "true";
}

function getAvailableReferenceSites(person) {
  if (!person || !state.data?.listes?.sites) {
    return [];
  }

  if (personUsesAllSites(person)) {
    const baseSites = (state.data.listes.sites || []).map(normalizeText).filter(Boolean);
    return Array.from(new Set([ALL_SITES_VALUE, ...baseSites.filter((site) => site !== ALL_SITES_VALUE)]));
  }

  return getPersonSites(person).filter((site) => normalizeText(site) !== ALL_SITES_VALUE);
}

function getSelectedEffectReferenceSite() {
  const field = document.querySelector('#effect-form [name="referenceSite"]');
  return normalizeText(field?.value);
}

function referenceSiteFromEffect(effect) {
  if (effect?.siteReference) {
    return normalizeText(effect.siteReference);
  }
  if (effect?.referenceEffetId) {
    const reference = findReferenceById(effect.referenceEffetId);
    return normalizeText(reference?.site);
  }
  return "";
}

function getDefaultEffectSiteReference(person, effect) {
  if (!typeUsesSiteField(effect?.typeEffet)) {
    return "";
  }
  const referenceSite = referenceSiteFromEffect(effect);
  if (referenceSite) {
    return referenceSite;
  }
  const personSites = getPersonSites(person);
  if (personSites.includes(ALL_SITES_VALUE)) {
    return ALL_SITES_VALUE;
  }
  if (personSites.length === 1) {
    return personSites[0];
  }
  return "";
}

function getReferenceSitesForType(typeEffet) {
  const normalizedTypeEffet = normalizeText(typeEffet);
  if (!normalizedTypeEffet || !Array.isArray(state.data?.listes?.referencesEffets)) {
    return [];
  }

  const sites = state.data.listes.referencesEffets
    .filter((reference) => {
      if (!isReferenceEffectActive(reference)) {
        return false;
      }
      if (!referenceMatchesType(reference, normalizedTypeEffet)) {
        return false;
      }
      return true;
    })
    .flatMap((reference) => getReferenceSites(reference))
    .map(normalizeText)
    .filter(Boolean);

  return Array.from(new Set(sites)).filter((site) => site !== ALL_SITES_VALUE);
}

function hydrateEffectReferenceSiteSelect(person, selectedSite = "", typeEffet = "") {
  const select = document.querySelector('#effect-form [name="referenceSite"]');
  if (!select) {
    return;
  }

  const normalizedType = normalizeText(
    typeEffet || document.querySelector('#effect-form [name="typeEffet"]')?.value
  );
  const baseOption = '<option value="">SELECTIONNER</option>';
  if (!typeUsesSiteField(normalizedType)) {
    select.innerHTML = baseOption;
    select.value = "";
    return;
  }

  let sites = getAvailableReferenceSites(person);
  if (typeUsesReferenceCatalog(normalizedType)) {
    sites = getReferenceSitesForType(normalizedType);
  }
  if (normalizedType === "VENTILATEUR") {
    sites = getAvailableReferenceSites(person);
  }
  if (normalizedType === "CARTE TURBOSELF") {
    sites = Array.from(new Set([ALL_SITES_VALUE, ...sites.filter((site) => site !== ALL_SITES_VALUE)]));
  }
  const options = sites.map((site) => `<option value="${site}">${site}</option>`).join("");
  select.innerHTML = `${baseOption}${options}`;
  const nextValue = normalizeText(selectedSite || (sites.length === 1 ? sites[0] : ""));
  select.value = nextValue;
}

function hydrateReferenceSelect(siteSource, typeEffet = "", selectedId = "", referenceSite = "") {
  const select = document.getElementById("effect-reference-select");
  if (!select || !state.data?.listes?.referencesEffets) {
    return;
  }

  void siteSource;
  const normalizedTypeEffet = normalizeText(typeEffet);
  const baseOption = '<option value="">SELECTIONNER</option>';
  if (!typeUsesReferenceCatalog(normalizedTypeEffet)) {
    select.innerHTML = baseOption;
    select.value = "";
    return;
  }
  const normalizedReferenceSite = normalizeText(referenceSite);
  const visibleSiteCount = getReferenceSitesForType(normalizedTypeEffet).length;
  const options = state.data.listes.referencesEffets
    .filter((reference) => {
      if (!isReferenceEffectActive(reference)) {
        return false;
      }
      if (normalizedTypeEffet !== "VENTILATEUR" && normalizedReferenceSite && !referenceHasSite(reference, normalizedReferenceSite)) {
        return false;
      }
      if (normalizedTypeEffet && !referenceMatchesType(reference, normalizedTypeEffet)) {
        return false;
      }
      return true;
    })
    .sort((left, right) => {
      const leftLabel = `${normalizeText(left.designation)} ${normalizeText(getReferenceSiteLabel(left))}`;
      const rightLabel = `${normalizeText(right.designation)} ${normalizeText(getReferenceSiteLabel(right))}`;
      return leftLabel.localeCompare(rightLabel, "fr");
    })
    .map((reference) => {
      const label =
        !normalizedReferenceSite && visibleSiteCount > 1
          ? `${reference.designation} - ${getReferenceSiteLabel(reference)}`
          : reference.designation;
      return `<option value="${reference.id}">${label}</option>`;
    })
    .join("");
  const fallbackOption = `<option value="">${
    normalizedReferenceSite ? "AUCUNE CLE POUR CE SITE" : "AUCUNE CLE DISPONIBLE"
  }</option>`;
  select.innerHTML = options ? `${baseOption}${options}` : fallbackOption;
  if (selectedId && options.includes(`value="${selectedId}"`)) {
    select.value = selectedId;
  } else {
    select.value = "";
  }
}

function ensureReferenceExists(site, typeEffet, designation, existingId) {
  if (
    existingId ||
    !site ||
    !state.data?.listes?.referencesEffets ||
    !typeUsesReferenceCatalog(typeEffet)
  ) {
    return;
  }

  const exists = state.data.listes.referencesEffets.some(
    (reference) =>
      reference.site === site &&
      referenceMatchesType(reference, typeEffet) &&
      reference.designation === designation
  );

  if (!exists) {
    state.data.listes.referencesEffets.push({
      id: getNextId("REF", state.data.listes.referencesEffets),
      site,
      typeEffet: normalizeText(typeEffet),
      designation,
    });
    sortReferenceEffects();
  }
}

function getCatalogEffectReferenceSite(person, effect) {
  const normalizedType = normalizeText(effect?.typeEffet || "");
  if (normalizedType === "CARTE TURBOSELF") {
    return ALL_SITES_VALUE;
  }
  return normalizeText(
    effect?.siteReference ||
      referenceSiteFromEffect(effect) ||
      getPersonSiteLabel(person) ||
      ALL_SITES_VALUE
  );
}

function ensureCatalogReferencesFromAssignedEffects() {
  if (!state.data?.listes?.referencesEffets || !Array.isArray(state.data?.personnes)) {
    return false;
  }
  let hasChanged = false;

  for (const person of state.data.personnes) {
    const effects = Array.isArray(person?.effetsConfies) ? person.effetsConfies : [];
    for (const effect of effects) {
      if (isSoftDeletedEntity(effect)) {
        continue;
      }
      const normalizedEffectType = normalizeText(effect?.typeEffet || "");
      const supportsAutoReference =
        typeUsesReferenceCatalog(normalizedEffectType) ||
        normalizedEffectType === "CARTE TURBOSELF" ||
        normalizedEffectType === "BADGE INTRUSION";
      if (!supportsAutoReference) {
        continue;
      }

      const typeEffet = normalizedEffectType;
      const site = getCatalogEffectReferenceSite(person, effect);
      const designation = normalizeReferenceDesignationByType(typeEffet, effect?.designation || "");
      const previousReferenceId = String(effect.referenceEffetId || "");
      const previousSiteReference = String(effect.siteReference || "");
      const previousDesignation = String(effect.designation || "");

      let reference =
        (state.data.listes.referencesEffets || []).find(
          (entry) =>
            referenceMatchesType(entry, typeEffet) &&
            referenceHasSite(entry, site) &&
            normalizeText(entry.designation) === normalizeText(designation)
        ) || null;

      if (!reference) {
        reference = {
          id: getNextId("REF", state.data.listes.referencesEffets),
          site,
          sitesAffectation: [site],
          typeEffet,
          designation,
          active: true,
        };
        state.data.listes.referencesEffets.push(reference);
      }

      const nextReferenceId = String(reference.id || "");
      if (previousReferenceId !== nextReferenceId) {
        hasChanged = true;
        effect.referenceEffetId = nextReferenceId;
      }
      if (previousSiteReference !== site) {
        hasChanged = true;
        effect.siteReference = site;
      }
      if (!previousDesignation && designation) {
        hasChanged = true;
        effect.designation = designation;
      }
    }
  }
  return hasChanged;
}

function findReferenceById(referenceId) {
  const normalizedId = String(referenceId || "");
  return state.data?.listes?.referencesEffets?.find((reference) => String(reference.id || "") === normalizedId) || null;
}

function getCurrentPerson() {
  const currentPersonId = String(getCurrentPersonId() || "");
  return (state.data?.personnes || []).find(
    (person) => String(person?.id || "") === currentPersonId
  ) || null;
}

function getTodayIsoDate() {
  const currentDate = new Date();
  const year = currentDate.getFullYear();
  const month = String(currentDate.getMonth() + 1).padStart(2, "0");
  const day = String(currentDate.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function getDossierStatus(person) {
  if (person.dateSortieReelle) return "SORTI";
  if (person.dateSortiePrevue) return "SORTIE PREVUE";
  return "EN POSTE";
}

function isPastDate(value) {
  if (!value) {
    return false;
  }
  const today = new Date();
  const todayOnly = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const target = new Date(`${value}T00:00:00`);
  return Number.isFinite(target.getTime()) && target < todayOnly;
}

function getOverdueExitAlertMeta(person) {
  const today = getTodayIsoDate();
  const plannedExit = normalizeDateString(person?.dateSortiePrevue || "");
  const realExit = normalizeDateString(person?.dateSortieReelle || "");
  const plannedDate = plannedExit ? new Date(`${plannedExit}T00:00:00`) : null;
  const todayDate = new Date(`${today}T00:00:00`);
  const daysUntilPlannedExit =
    plannedDate && Number.isFinite(plannedDate.getTime())
      ? Math.round((plannedDate.getTime() - todayDate.getTime()) / (24 * 60 * 60 * 1000))
      : null;

  if (
    plannedExit &&
    !realExit &&
    Number.isFinite(daysUntilPlannedExit) &&
    daysUntilPlannedExit >= 1 &&
    daysUntilPlannedExit <= 2
  ) {
    return {
      type: "dateSortiePrevue",
      message: `ALERTE : SORTIE PREVUE DANS ${daysUntilPlannedExit} JOUR${daysUntilPlannedExit > 1 ? "S" : ""} (${formatDate(plannedExit)})`,
    };
  }

  if (plannedExit && !realExit && plannedExit <= today) {
    return {
      type: "dateSortiePrevue",
      message:
        plannedExit === today
          ? `ALERTE : SORTIE PREVUE AUJOURD'HUI (${formatDate(plannedExit)})`
          : `ALERTE : DATE DE SORTIE PREVUE DEPASSEE (${formatDate(plannedExit)})`,
    };
  }
  const hasNonRenduEffects = hasCurrentNonRenduEffects(person);
  if (realExit && realExit <= today && hasNonRenduEffects) {
    return {
      type: "dateSortieReelle",
      message:
        realExit === today
          ? `ALERTE : SORTIE REELLE AUJOURD'HUI AVEC EFFETS NON RENDUS (${formatDate(realExit)})`
          : `ALERTE : DATE DE SORTIE REELLE DEPASSEE (${formatDate(realExit)})`,
    };
  }
  return { type: "", message: "" };
}

function getOverdueExitMessage(person) {
  return getOverdueExitAlertMeta(person).message;
}

function hasOverdueExit(person) {
  return Boolean(getOverdueExitAlertMeta(person).message);
}

function isExitDue(person) {
  const today = getTodayIsoDate();
  const realExit = normalizeDateString(person?.dateSortieReelle || "");
  const plannedExit = normalizeDateString(person?.dateSortiePrevue || "");
  if (realExit && realExit <= today) {
    return true;
  }
  if (plannedExit && plannedExit <= today) {
    return true;
  }
  return false;
}

function deriveEffectState(person, effect) {
  const hasReturnDate = Boolean(String(effect?.dateRetour || "").trim());
  const manualStatus = normalizeText(effect?.statutManuel);
  const persistedCause = normalizeEffectCause(effect?.cause || effect?.causeRemplacement);
  const fallbackCause = !hasReturnDate && isExitDue(person) ? "NON RENDU" : "";

  let status = "ACTIF";
  if (hasReturnDate) status = "RESTITUE";
  else if (manualStatus === "CASSE") status = "DETRUIT";
  else if (["PERDU", "HS", "VOL"].includes(manualStatus)) status = manualStatus;
  else if ((!manualStatus || manualStatus === "ACTIF") && isExitDue(person)) status = "NON RENDU";
  else if (manualStatus) status = manualStatus;

  const cause = persistedCause || fallbackCause;
  const movement =
    status === "RESTITUE"
      ? "RENDU"
      : status === "DETRUIT"
        ? "DETRUIT"
        : status === "VOL" || cause === "VOL"
          ? "VOLE"
          : status === "HS"
            ? "HS"
            : status === "PERDU" || cause === "PERTE"
              ? "PERDU"
              : status === "NON RENDU"
                ? "NON RENDU"
                : "";
  const replacementCost = cause ? getReplacementCostValue(effect?.typeEffet, cause, effect?.designation || "") : 0;
  const chargeable = replacementCost > 0;
  const storedBilling = normalizeText(effect?.etatFacturation || "");
  const billingStatus =
    storedBilling === "FACTURE" ? "FACTURE" : storedBilling === "CLOTURE" ? "CLOTURE" : chargeable ? "A FACTURER" : "-";
  return { status, cause, movement, chargeable, billingStatus };
}

function getEffectStatus(person, effect) {
  if (typeof deriveEffectState === "function") {
    return deriveEffectState(person, effect).status;
  }
  if (effect?.dateRetour) return "RESTITUE";
  const manualStatus = normalizeText(effect?.statutManuel);
  if (manualStatus === "CASSE") return "DETRUIT";
  if (["PERDU", "HS", "VOL"].includes(manualStatus)) return manualStatus;
  if ((!manualStatus || manualStatus === "ACTIF") && isExitDue(person)) return "NON RENDU";
  return manualStatus || "ACTIF";
}

function getEffectBillingStatus(effect, isChargeable) {
  const stored = normalizeText(effect?.etatFacturation || "");
  if (stored === "FACTURE") return "FACTURE";
  if (stored === "CLOTURE") return "CLOTURE";
  return isChargeable ? "A FACTURER" : "-";
}

function getEffectDisplayDesignation(effect) {
  return typeUsesReferenceCatalog(effect?.typeEffet) ? effect?.designation || "" : "";
}

function getEffectDisplaySite(effect) {
  if (effect?.siteReference) {
    return effect.siteReference;
  }
  if (effect?.referenceEffetId) {
    const reference = findReferenceById(effect.referenceEffetId);
    return getReferenceSiteLabel(reference) || normalizeText(reference?.site);
  }
  return "";
}

function getStatusClass(status) {
  const normalizedStatus = normalizeText(status).replace(/\s+/g, "-").toLowerCase();
  return `status-pill status-pill--${normalizedStatus}`;
}

function getAllEffects(persons) {
  return persons.flatMap((person) =>
    (person.effetsConfies || []).map((effect) => ({ person, effect }))
  );
}

function effectMatchesActiveObjectFilters(person, effect, filters = state.filters || DEFAULT_FILTERS) {
  if (filters.site && !effectMatchesSiteFilter(person, effect, filters.site)) {
    return false;
  }
  if (filters.typeEffet && normalizeText(effect?.typeEffet) !== filters.typeEffet) {
    return false;
  }
  if (filters.statutObjet && getEffectStatus(person, effect) !== filters.statutObjet) {
    return false;
  }
  return true;
}

function getEffectsForActiveFilters(person, filters = state.filters || DEFAULT_FILTERS) {
  return (person?.effetsConfies || []).filter((effect) =>
    effectMatchesActiveObjectFilters(person, effect, filters)
  );
}

function isCurrentAssignedEffect(person, effect) {
  const status = normalizeText(getEffectStatus(person, effect));
  return !["RESTITUE", "PERDU", "HS", "DETRUIT", "VOL"].includes(status);
}

function getCurrentAssignedEffects(person) {
  return (person?.effetsConfies || []).filter((effect) => isCurrentAssignedEffect(person, effect));
}

function getCurrentAssignedEffectsForActiveFilters(person, filters = state.filters || DEFAULT_FILTERS) {
  return getCurrentAssignedEffects(person).filter((effect) =>
    effectMatchesActiveObjectFilters(person, effect, filters)
  );
}

function getEffectChartCategory(person, effect) {
  const status = normalizeText(getEffectStatus(person, effect));
  if (status === "NON RENDU") {
    return "nonRendu";
  }
  if (status === "RESTITUE") {
    return "restitue";
  }
  if (status === "PERDU") {
    return "perdu";
  }
  if (status === "VOL") {
    return "vole";
  }
  if (status === "HS") {
    return "hs";
  }
  return "actif";
}

function getNextId(prefix, items) {
  const max = items.reduce((highest, item) => {
    const digits = Number.parseInt(String(item.id || "").replace(prefix, ""), 10);
    return Number.isNaN(digits) ? highest : Math.max(highest, digits);
  }, 0);

  return `${prefix}${String(max + 1).padStart(4, "0")}`;
}

function formatDate(value) {
  if (!value) return "";
  const [year, month, day] = value.split("-");
  if (!year || !month || !day) return value;
  return `${day}/${month}/${year}`;
}

function formatTime(value) {
  if (!value) {
    return "";
  }
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return "";
  }
  return date.toLocaleTimeString("fr-FR", {
    hour: "2-digit",
    minute: "2-digit",
  });
}

function formatDateForDocument(value) {
  return formatDate(value) || formatDate(getTodayIsoDate());
}

function formatDateTimeForDocument(value) {
  const date = value ? new Date(value) : new Date();
  if (Number.isNaN(date.getTime())) {
    return formatDateForDocument("");
  }
  return date.toLocaleString("fr-FR", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function renderReferenceBases() {
  const listSignature = (values) =>
    Array.isArray(values)
      ? values
          .map((value) => {
            const source = value || {};
            const raw = `${String(source.id || "")}|${normalizeText(source.typePersonnel || source.nom || source.designation || source.cause || source.site || source.typeEffet || source.typeContrat || source.fonction || "")}|${String(
              source.actif || source.montant || source.type || source.cause || ""
            )}`;
            return normalizeText(raw);
          })
          .sort()
          .join("||")
      : "";

  const listMapSignature = (map) =>
    map instanceof Map
      ? Array.from(map.entries())
          .map(([key, value]) => `${String(key || "")}:${String(value)}`)
          .sort()
          .join("||")
      : "";

  const filterForm = document.getElementById("reference-filter-form");
  const filterSearch = normalizeText(filterForm?.elements?.filterReferenceSearch?.value);
  const filterSite = normalizeText(filterForm?.elements?.filterReferenceSite?.value);
  const filterTypeEffet = normalizeText(filterForm?.elements?.filterReferenceTypeEffet?.value);
  const stockFilters = state.stockTableFilters || { site: "", typeEffet: "", referenceEffetId: "" };
  const stockMovementsSignature = (state.data?.stocksEffetsManuels || [])
    .map((movement) => {
      const entry = movement || {};
      return `${String(entry.id || "")}|${normalizeText(entry.typeEffet)}|${normalizeText(entry.site)}|${normalizeText(
        entry.designation
      )}|${normalizeText(entry.referenceEffetId)}|${String(entry.quantite || 0)}|${normalizeText(entry.action)}|${String(
        entry.date || ""
      )}`;
    })
    .sort()
    .join("||");

  const referenceRenderContext = buildReferenceRenderContext();
  const referenceSort = getEffectTableSort("referenceEffects");
  const nextReferenceBasesRenderSignature = JSON.stringify({
    revision: String(state.supabaseRevision || ""),
    lists: {
      sites: listSignature(state.data?.listes?.sites),
      typesPersonnel: listSignature(state.data?.listes?.typesPersonnel),
      typesContrats: listSignature(state.data?.listes?.typesContrats),
      fonctions: listSignature(state.data?.listes?.fonctions),
      causesRemplacement: listSignature(state.data?.listes?.causesRemplacement),
      typesEffets: listSignature(state.data?.listes?.typesEffets),
      references: listSignature(state.data?.listes?.referencesEffets),
      cots: listSignature(state.data?.listes?.coutsRemplacement),
      representants: listSignature(state.data?.listes?.representantsSignataires),
      stockMovements: stockMovementsSignature,
      listesFingerprint:
        state.data?.listes?.referencesEffets?.length || 0,
    },
    context: {
      typesPersonnel: listMapSignature(referenceRenderContext.simpleUsage.typesPersonnel),
      typesContrats: listMapSignature(referenceRenderContext.simpleUsage.typesContrats),
      fonctions: listMapSignature(referenceRenderContext.simpleUsage.fonctions),
      typesEffets: listMapSignature(referenceRenderContext.simpleUsage.typesEffets),
      representativeUsage: listMapSignature(referenceRenderContext.representativeUsage),
      referenceEffectUsage: listMapSignature(referenceRenderContext.referenceEffectUsage),
    },
    filters: {
      search: filterSearch,
      site: filterSite,
      typeEffet: filterTypeEffet,
      sortKey: String(referenceSort.key || ""),
      sortDir: String(referenceSort.dir || ""),
      stock: `${normalizeText(stockFilters.site || "")}|${normalizeText(stockFilters.typeEffet || "")}|${String(
        stockFilters.referenceEffetId || ""
      )}`,
    },
  });

  if (state.listRenderCache.referenceBases === nextReferenceBasesRenderSignature) {
    return false;
  }

  state.listRenderCache.referenceBases = nextReferenceBasesRenderSignature;
  state.referenceRenderContext = referenceRenderContext;
  renderSimpleReferenceList("sites", state.referenceRenderContext);
  renderSimpleReferenceList("typesPersonnel", state.referenceRenderContext);
  renderSimpleReferenceList("typesContrats", state.referenceRenderContext);
  renderSimpleReferenceList("fonctions", state.referenceRenderContext);
  renderSimpleReferenceList("typesEffets", state.referenceRenderContext);
  renderSimpleReferenceList("causesRemplacement", state.referenceRenderContext);
  renderRepresentativesTable(state.referenceRenderContext);
  renderReferenceEffectsTable(state.referenceRenderContext);
  renderReplacementCostsTable();
  renderStockFormOptions();
  renderStockMovementsTable();
  renderStockSummaryTable();
  renderStockTypeKpis();
  renderStockInstantKpi();
  renderReferenceCounts();
  renderMobileSignatureSettings();
  bindReferenceBaseResetButtons();
  return true;
}

function renderStockTypeKpis() {
  const summaryRows = getFilteredStockSummaryRows();
  const nextStockKpisSignature = (() => {
    const knownTypes = Array.from(
      new Set(((state.data?.listes?.typesEffets || []).map(normalizeText).filter(Boolean)))
    ).sort((a, b) => a.localeCompare(b, "fr"));
    const rowTypes = Array.from(new Set(summaryRows.map((row) => normalizeText(row.typeEffet)).filter(Boolean))).sort(
      (a, b) => a.localeCompare(b, "fr")
    );
    const mergedTypes = Array.from(new Set([...knownTypes, ...rowTypes])).slice(0, 5);
    const valuesSignature = mergedTypes
      .map((typeEffet) => {
        const totalStock = summaryRows
          .filter((row) => normalizeText(row.typeEffet) === typeEffet)
          .reduce((sum, row) => sum + Number(row.stockCourant || 0), 0);
        return `${typeEffet || "EMPTY"}:${String(totalStock)}`;
      })
      .join("|");
    return `stock-kpis|${knownTypes.join("|")}|${mergedTypes.join("|")}|${valuesSignature}`;
  })();
  if (state.listRenderCache.stockKpis === nextStockKpisSignature) {
    return false;
  }
  state.listRenderCache.stockKpis = nextStockKpisSignature;

  const knownTypes = Array.from(
    new Set(((state.data?.listes?.typesEffets || []).map(normalizeText).filter(Boolean)))
  ).sort((a, b) => a.localeCompare(b, "fr"));
  const rowTypes = Array.from(
    new Set(summaryRows.map((row) => normalizeText(row.typeEffet)).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b, "fr"));
  const mergedTypes = Array.from(new Set([...knownTypes, ...rowTypes])).slice(0, 5);

  for (let index = 1; index <= 5; index += 1) {
    const typeEffet = mergedTypes[index - 1] || "";
    const labelNode = document.getElementById(`stock-type-kpi-label-${index}`);
    const valueNode = document.getElementById(`stock-type-kpi-value-${index}`);
    if (labelNode) {
      labelNode.textContent = typeEffet || `TYPE ${index}`;
    }
    if (!valueNode) {
      continue;
    }
    if (!typeEffet) {
      setKpiCountAnimated(valueNode, 0);
      continue;
    }
    const totalStock = summaryRows
      .filter((row) => normalizeText(row.typeEffet) === typeEffet)
      .reduce((sum, row) => sum + Number(row.stockCourant || 0), 0);
    setKpiCountAnimated(valueNode, totalStock);
  }
  return true;
}

function getFilteredStockSummaryRows() {
  const filters = state.stockTableFilters || { site: "", typeEffet: "", referenceEffetId: "" };
  return getStockSummaryRows().filter((row) => {
    if (filters.site && filters.site !== ALL_SITES_VALUE && normalizeText(row.site) !== filters.site) {
      return false;
    }
    if (filters.typeEffet && filters.typeEffet !== ALL_TYPES_VALUE && normalizeText(row.typeEffet) !== filters.typeEffet) {
      return false;
    }
    if (filters.referenceEffetId && filters.referenceEffetId !== ALL_DESIGNATIONS_VALUE) {
      const syntheticReference = parseStockSyntheticReferenceValue(filters.referenceEffetId);
      const reference = syntheticReference ? null : findReferenceById(filters.referenceEffetId);
      if (syntheticReference) {
        return (
          normalizeText(row.site) === syntheticReference.site &&
          normalizeText(row.typeEffet) === syntheticReference.typeEffet &&
          normalizeText(row.designation) === syntheticReference.designation
        );
      }
      const referenceDesignation = reference
        ? getStockGroupingDesignation(getReferenceEffectiveType(reference), getStockReferenceDesignation(reference))
        : "";
      if (referenceDesignation && normalizeText(row.designation) !== referenceDesignation) {
        return false;
      }
    }
    return true;
  });
}

function getStockInstantSelectionValue(filters = state.stockTableFilters || {}) {
  const previousFilters = state.stockTableFilters;
  state.stockTableFilters = filters;
  const rows = getFilteredStockSummaryRows();
  state.stockTableFilters = previousFilters;
  const stockCourant = rows.reduce((sum, row) => sum + Number(row.stockCourant || 0), 0);
  return {
    isPrecise: rows.length > 0,
    stockCourant,
    label: `${rows.length} ligne(s) de stock concernee(s)`,
  };
}

function renderStockInstantKpi() {
  const valueNode = document.getElementById("stock-instant-value");
  if (!valueNode) {
    return false;
  }
  const selection = getStockInstantSelectionValue(state.stockTableFilters || {});
  const nextStockInstantSignature = `instant|${String(state.stockTableFilters?.site || "")}|${String(
    state.stockTableFilters?.typeEffet || ""
  )}|${String(state.stockTableFilters?.referenceEffetId || "")}|${String(selection.stockCourant)}|${selection.label}`;
  if (state.listRenderCache.stockInstantKpi === nextStockInstantSignature) {
    return false;
  }
  state.listRenderCache.stockInstantKpi = nextStockInstantSignature;
  valueNode.textContent = String(selection.stockCourant);
  valueNode.title = selection.label || "Aucune ligne de stock concernee.";
  return true;
}

function renderStockFormOptions() {
  const form = document.getElementById("stock-adjustment-form");
  if (!(form instanceof HTMLFormElement)) {
    return;
  }
  const typeSelect = form.elements.stockTypeEffet;
  const siteSelect = form.elements.stockSite;
  const reasonSelect = form.elements.stockReason;
  if (
    !(typeSelect instanceof HTMLSelectElement) ||
    !(siteSelect instanceof HTMLSelectElement) ||
    !(reasonSelect instanceof HTMLSelectElement)
  ) {
    return;
  }
  const values = Array.from(
    new Set(
      []
        .concat((state.data?.listes?.typesEffets || []).map(normalizeText))
        .concat(
          (state.data?.listes?.referencesEffets || [])
            .filter((reference) => isReferenceEffectActive(reference))
            .map((reference) => normalizeText(reference.typeEffet))
        )
        .concat(
          getAllEffects(state.data?.personnes || []).map(({ effect }) => normalizeText(effect?.typeEffet || ""))
        )
        .filter(Boolean)
    )
  ).sort((a, b) => normalizeText(a).localeCompare(normalizeText(b), "fr"));
  const currentType = String(typeSelect.value || "");
  typeSelect.innerHTML = [`<option value="">SELECTIONNER</option>`, `<option value="${ALL_TYPES_VALUE}">${ALL_TYPES_VALUE}</option>`]
    .concat(values.map((entry) => `<option value="${escapeHtml(entry)}">${escapeHtml(entry)}</option>`))
    .join("");
  if (currentType && Array.from(typeSelect.options).some((opt) => opt.value === currentType)) {
    typeSelect.value = currentType;
  }
  const sites = (state.data?.listes?.sites || [])
    .slice()
    .sort((a, b) => normalizeText(a).localeCompare(normalizeText(b), "fr"));
  const currentSite = String(siteSelect.value || "");
  siteSelect.innerHTML = [`<option value="">SELECTIONNER</option>`, `<option value="${ALL_SITES_VALUE}">${ALL_SITES_VALUE}</option>`]
    .concat(sites.map((entry) => `<option value="${escapeHtml(entry)}">${escapeHtml(entry)}</option>`))
    .join("");
  if (currentSite && Array.from(siteSelect.options).some((opt) => opt.value === currentSite)) {
    siteSelect.value = currentSite;
  }
  syncSelectOptions(reasonSelect, getStockReasonOptions(), "SELECTIONNER");
  updateStockDesignationOptions();
  refreshStockTableFiltersFromForm();
}

function getStockReasonOptions() {
  const baseCauses = Array.isArray(state.data?.listes?.causesRemplacement)
    ? state.data.listes.causesRemplacement.map(normalizeText).filter(Boolean)
    : [];
  const defaults = [
    "ACHAT",
    "DOUBLE",
    "CORRECTION INVENTAIRE",
    "REBUT",
    "DON",
    "TRANSFERT INTERNE",
    "REMPLACEMENT",
  ];
  return Array.from(new Set([...baseCauses, ...defaults])).sort((a, b) =>
    normalizeText(a).localeCompare(normalizeText(b), "fr")
  );
}

function updateStockDesignationOptions() {
  const form = document.getElementById("stock-adjustment-form");
  if (!(form instanceof HTMLFormElement)) {
    return;
  }
  const typeEffet = normalizeText(form.elements.stockTypeEffet?.value || "");
  const site = normalizeText(form.elements.stockSite?.value || "");
  const designationSelect = form.elements.stockReferenceId;
  if (!(designationSelect instanceof HTMLSelectElement)) {
    return;
  }
  if (!typeEffet) {
    designationSelect.innerHTML = `<option value="">SELECTIONNER</option>`;
    designationSelect.disabled = true;
    return;
  }
  const references = (state.data?.listes?.referencesEffets || [])
    .filter((reference) => isReferenceEffectActive(reference))
    .filter((reference) => {
      if (!typeEffet || typeEffet === ALL_TYPES_VALUE) {
        return true;
      }
      if (!referenceMatchesType(reference, typeEffet)) {
        return false;
      }
      return true;
    })
    .filter((reference) => !site || site === ALL_SITES_VALUE || referenceHasSite(reference, site))
    .sort((a, b) => getStockReferenceDesignation(a).localeCompare(getStockReferenceDesignation(b), "fr"));
  const referenceOptionValues = new Set(references.map((reference) => String(reference.id || "")));
  const syntheticRows = getStockSummaryRows()
    .filter((row) => {
      if (!site || site === ALL_SITES_VALUE || normalizeText(row.site) === site) {
        return !typeEffet || typeEffet === ALL_TYPES_VALUE || normalizeText(row.typeEffet) === typeEffet;
      }
      return false;
    })
    .filter((row) => normalizeText(row.designation) === STOCK_EMPTY_DESIGNATION_LABEL)
    .filter((row) => {
      return !references.some(
        (reference) =>
          referenceMatchesType(reference, row.typeEffet) &&
          referenceHasSite(reference, row.site) &&
          getStockReferenceDesignation(reference) === STOCK_EMPTY_DESIGNATION_LABEL
      );
    })
    .sort((a, b) => {
      const left = `${normalizeText(a.site)} ${normalizeText(a.typeEffet)} ${normalizeText(a.designation)}`;
      const right = `${normalizeText(b.site)} ${normalizeText(b.typeEffet)} ${normalizeText(b.designation)}`;
      return left.localeCompare(right, "fr");
    });

  const currentValue = String(designationSelect.value || "");
  const canUseAllDesignations = !typeEffet || typeEffet === ALL_TYPES_VALUE || !site || site === ALL_SITES_VALUE;
  const optionHtml = [`<option value="">SELECTIONNER</option>`];
  if (canUseAllDesignations) {
    optionHtml.push(`<option value="${ALL_DESIGNATIONS_VALUE}">TOUTES DESIGNATIONS</option>`);
  }
  designationSelect.innerHTML = optionHtml
    .concat(
      references.map(
        (reference) =>
          `<option value="${escapeHtml(String(reference.id || ""))}">${escapeHtml(getStockReferenceDesignation(reference))}</option>`
      )
    )
    .concat(
      syntheticRows.map((row) => {
        const value = getStockSyntheticReferenceValue(row.site, row.typeEffet, row.designation);
        referenceOptionValues.add(value);
        return `<option value="${escapeHtml(value)}">${escapeHtml(row.designation)}</option>`;
      })
    )
    .join("");
  if (currentValue && references.some((entry) => String(entry.id || "") === currentValue)) {
    designationSelect.value = currentValue;
  } else if (currentValue && referenceOptionValues.has(currentValue)) {
    designationSelect.value = currentValue;
  } else if (currentValue === ALL_DESIGNATIONS_VALUE && canUseAllDesignations) {
    designationSelect.value = ALL_DESIGNATIONS_VALUE;
  } else {
    designationSelect.value = "";
  }
  designationSelect.disabled = references.length === 0 && syntheticRows.length === 0;
}

function refreshStockTableFiltersFromForm() {
  const form = document.getElementById("stock-adjustment-form");
  if (!(form instanceof HTMLFormElement)) {
    return;
  }
  state.stockTableFilters = {
    site: normalizeText(form.elements.stockSite?.value || ""),
    typeEffet: normalizeText(form.elements.stockTypeEffet?.value || ""),
    referenceEffetId: String(form.elements.stockReferenceId?.value || ""),
  };
}

function resetStockTableFiltersFromForm() {
  const form = document.getElementById("stock-adjustment-form");
  if (!(form instanceof HTMLFormElement)) {
    return;
  }
  if (form.elements.stockSite) form.elements.stockSite.value = "";
  if (form.elements.stockTypeEffet) form.elements.stockTypeEffet.value = "";
  updateStockDesignationOptions();
  if (form.elements.stockReferenceId) form.elements.stockReferenceId.value = "";
  refreshStockTableFiltersFromForm();
  renderStockMovementsTable();
  renderStockSummaryTable();
  renderStockTypeKpis();
  renderStockInstantKpi();
}

function renderMobileSignatureSettings() {
  const input = document.querySelector('#mobile-signature-settings-form [name="mobileSignatureBaseUrl"]');
  const statusNode = document.getElementById("mobile-signature-settings-status");
  if (!(input instanceof HTMLInputElement)) {
    return;
  }

  const withTrailingSlash = (url) => {
    const raw = String(url || "").trim();
    if (!raw) return "";
    return raw.endsWith("/") ? raw : `${raw}/`;
  };

  const configured = getConfiguredMobileSignatureBaseUrl();
  input.value = withTrailingSlash(configured);
  if (!statusNode) {
    return;
  }

  const setStatusWithUrl = (prefix, url) => {
    const safeUrl = escapeHtml(withTrailingSlash(url));
    statusNode.innerHTML = `${prefix} : <a href="${safeUrl}" target="_blank" rel="noopener">${safeUrl}</a>`;
  };

  if (configured) {
    setStatusWithUrl("URL PUBLIQUE ACTIVE", configured);
    return;
  }

  const currentOrigin = normalizeHttpUrl(window.location.origin || "");
  if (currentOrigin && !isLikelyLocalUrl(currentOrigin)) {
    setStatusWithUrl("MODE AUTO HEBERGE", currentOrigin);
    return;
  }

  const autoBase = normalizeHttpUrl(state.mobileSignatureNetworkInfo?.preferredUrl || "");
  if (autoBase) {
    setStatusWithUrl("MODE AUTO RESEAU LOCAL", autoBase);
    return;
  }

  statusNode.textContent = "MODE AUTO RESEAU LOCAL (URL PUBLIQUE NON DEFINIE)";
}

function buildReferenceRenderContext() {
  const persons = state.data?.personnes || [];
  const references = state.data?.listes?.referencesEffets || [];
  const effects = getAllEffects(persons);
  const context = {
    simpleUsage: {
      typesPersonnel: new Map(),
      typesContrats: new Map(),
      fonctions: new Map(),
      typesEffets: new Map(),
    },
    representativeUsage: new Map(),
    referenceEffectUsage: new Map(),
  };

  const increment = (map, key, amount = 1) => {
    const normalizedKey = normalizeText(key);
    if (!normalizedKey) {
      return;
    }
    map.set(normalizedKey, (map.get(normalizedKey) || 0) + amount);
  };

  persons.forEach((person) => {
    increment(context.simpleUsage.typesPersonnel, person.typePersonnel);
    increment(context.simpleUsage.typesContrats, person.typeContrat);
    increment(context.simpleUsage.fonctions, person.fonction);

    ["arrival", "exit"].forEach((docType) => {
      const representativeId = String(person?.representants?.[docType]?.id || "");
      if (representativeId) {
        context.representativeUsage.set(
          representativeId,
          (context.representativeUsage.get(representativeId) || 0) + 1
        );
      }
    });
  });

  references.forEach((reference) => {
    increment(context.simpleUsage.typesEffets, reference.typeEffet);
  });

  effects.forEach(({ effect }) => {
    increment(context.simpleUsage.typesEffets, effect.typeEffet);
    const referenceId = String(effect.referenceEffetId || "");
    if (referenceId) {
      context.referenceEffectUsage.set(
        referenceId,
        (context.referenceEffectUsage.get(referenceId) || 0) + 1
      );
    }
  });

  return context;
}

function renderRepresentativesTable(renderContext = null) {
  const body = document.getElementById("reference-representantsSignataires-body");
  const form = document.getElementById("representative-signatory-form");
  if (!body || !form || !state.data?.listes?.representantsSignataires) {
    return;
  }

  const submitButton = form.querySelector('button[type="submit"]');
  if (submitButton) {
    submitButton.textContent = state.editingRepresentativeId
      ? "ENREGISTRER LA MODIFICATION"
      : "ENREGISTRER LE REPRESENTANT";
  }

  const representatives = state.data.listes.representantsSignataires.slice();
  if (!representatives.length) {
    body.innerHTML = buildEmptyTableRow(body, "AUCUN REPRESENTANT", 4);
    return;
  }

  const rowsHtml = representatives
    .map((representative) => {
      const usage = renderContext?.representativeUsage?.get(representative.id) || 0;
      return `<tr class="js-representative-row" data-representative-id="${representative.id}">
        <td>${escapeHtml(representative.nom || "-")}</td>
        <td>${escapeHtml(representative.fonction || "-")}</td>
        <td>${usage}</td>
        <td>
          <button type="button" class="table-link js-edit-representative" data-representative-id="${representative.id}">MODIFIER</button>
          <button type="button" class="table-link js-delete-representative" data-representative-id="${representative.id}">SUPPRIMER</button>
        </td>
      </tr>`;
    });

  renderTableRowsProgressively(body, rowsHtml, buildEmptyTableRow(body, "AUCUN REPRESENTANT", 4), 24);

  bindRepresentativeActions();
}

function renderSimpleReferenceList(listName, renderContext = null) {
  const body = document.getElementById(`reference-${listName}-body`);
  if (!body || !state.data?.listes?.[listName]) {
    return;
  }

  const values = state.data.listes[listName];
  if (!values.length) {
    body.innerHTML = buildEmptyTableRow(body, "AUCUNE VALEUR", 2);
    return;
  }

  const rowsHtml = values
    .map((value) => {
      const usage =
        listName === "sites" || listName === "causesRemplacement"
          ? getSimpleReferenceUsage(listName, value)
          : renderContext?.simpleUsage?.[listName]?.get(normalizeText(value)) || 0;
      return `<tr class="js-reference-item-row" data-list-name="${listName}" data-value="${escapeHtml(value)}">
        <td>${escapeHtml(value)}</td>
        <td>
          <button type="button" class="table-link js-edit-reference-item" data-list-name="${listName}" data-value="${escapeHtml(value)}">MODIFIER</button>
          <button type="button" class="table-link js-delete-reference-item" data-list-name="${listName}" data-value="${escapeHtml(value)}">SUPPRIMER</button>
          <span class="usage-pill">${usage}</span>
        </td>
      </tr>`;
    });

  renderTableRowsProgressively(body, rowsHtml, buildEmptyTableRow(body, "AUCUNE VALEUR", 2), 30);

  bindReferenceListActions();
}

function renderReferenceEffectsTable(renderContext = null) {
  const body = document.getElementById("reference-effects-table-body");
  if (!body || !state.data?.listes?.referencesEffets) {
    return;
  }

  const filterForm = document.getElementById("reference-filter-form");
  const filterSearch = normalizeText(filterForm?.elements?.filterReferenceSearch?.value);
  const filterSite = normalizeText(filterForm?.elements?.filterReferenceSite?.value);
  const filterTypeEffet = normalizeText(filterForm?.elements?.filterReferenceTypeEffet?.value);
  const references = state.data.listes.referencesEffets.filter((reference) => {
    if (filterSite && !referenceHasSite(reference, filterSite)) {
      return false;
    }
    if (filterTypeEffet) {
      if (!referenceMatchesType(reference, filterTypeEffet)) {
        return false;
      }
    }
    if (filterSearch) {
      const text = [getReferenceSiteLabel(reference), reference.typeEffet, reference.designation]
        .map(normalizeText)
        .join(" ");
      if (!text.includes(filterSearch)) {
        return false;
      }
    }
    return true;
  });
  const referenceLessTypes = (state.data?.listes?.typesEffets || [])
    .map(normalizeText)
    .filter((typeEffet) => EFFECT_TYPES_WITHOUT_REFERENCE_DESIGNATION.includes(typeEffet))
    .filter((typeEffet) => !filterTypeEffet || typeEffet === filterTypeEffet)
    .filter((typeEffet) => {
      if (!filterSearch) {
        return true;
      }
      return `${typeEffet} SANS DESIGNATION PAS DE REFERENCE`.includes(filterSearch);
    })
    .filter((typeEffet) => !state.data.listes.referencesEffets.some((reference) => normalizeText(reference?.typeEffet) === typeEffet));
  const sortedReferences = sortReferencesForTable(references, "referenceEffects", renderContext);
  if (!references.length && !referenceLessTypes.length) {
    body.innerHTML = buildEmptyTableRow(body, "AUCUNE REFERENCE", 5);
    return;
  }

  const rowsHtml = sortedReferences
    .map((reference) => {
      const usage = renderContext?.referenceEffectUsage?.get(String(reference.id || "")) || 0;
      return `<tr class="js-reference-effect-row" data-reference-id="${reference.id}">
        <td>${escapeHtml(getReferenceSiteLabel(reference))}</td>
        <td>${escapeHtml(reference.typeEffet)}</td>
        <td>${escapeHtml(reference.designation)}${isReferenceEffectActive(reference) ? "" : ' <span class="table-muted">(DESACTIVEE)</span>'}</td>
        <td>${usage}</td>
        <td>
          <button type="button" class="table-link js-edit-reference-effect" data-reference-id="${reference.id}">MODIFIER</button>
          <button type="button" class="table-link js-delete-reference-effect" data-reference-id="${reference.id}">${isReferenceEffectActive(reference) ? "DESACTIVER" : "REACTIVER"}</button>
          <button type="button" class="table-link js-hard-delete-reference-effect" data-reference-id="${reference.id}">SUPPRIMER</button>
        </td>
      </tr>`;
    })
    .concat(
      referenceLessTypes.map((typeEffet) => {
        const usage = renderContext?.simpleUsage?.typesEffets?.get(typeEffet) || 0;
        return `<tr class="js-reference-effect-row">
          <td>${escapeHtml(ALL_SITES_VALUE)}</td>
          <td>${escapeHtml(typeEffet)}</td>
          <td>${escapeHtml(STOCK_EMPTY_DESIGNATION_LABEL)} <span class="table-muted">(PAS DE REFERENCE)</span></td>
          <td>${usage}</td>
          <td><span class="table-muted">TYPE SANS DESIGNATION</span></td>
        </tr>`;
      })
    );

  renderTableRowsProgressively(body, rowsHtml, buildEmptyTableRow(body, "AUCUNE REFERENCE", 5), 24);
  updateSortableHeaders("referenceEffects");

  bindReferenceEffectActions();
}

function renderReplacementCostsTable() {
  const body = document.getElementById("replacement-costs-body");
  if (!body || !state.data?.listes?.coutsRemplacement) {
    return;
  }

  const entries = state.data.listes.coutsRemplacement
    .slice()
    .sort((left, right) => {
      const leftLabel = `${normalizeText(left.typeEffet)} ${normalizeText(left.designation)} ${normalizeText(left.cause)}`;
      const rightLabel = `${normalizeText(right.typeEffet)} ${normalizeText(right.designation)} ${normalizeText(right.cause)}`;
      return leftLabel.localeCompare(rightLabel, "fr");
    });

  if (!entries.length) {
    body.innerHTML = buildEmptyTableRow(body, "AUCUN COUT", 4);
    return;
  }

  const rowsHtml = entries
    .map((entry) => {
      const key = getReplacementCostKey(entry.typeEffet, entry.cause, entry.designation);
      const typeLabel = entry.designation ? `${entry.typeEffet} / ${entry.designation}` : entry.typeEffet;
      return `<tr class="js-replacement-cost-row" data-cost-key="${key}">
        <td>${escapeHtml(typeLabel)}</td>
        <td>${escapeHtml(entry.cause)}</td>
        <td>${formatAmountWithEuro(entry.montant)}</td>
        <td>
          <button type="button" class="table-link js-edit-replacement-cost" data-cost-key="${key}">MODIFIER</button>
          <button type="button" class="table-link js-delete-replacement-cost" data-cost-key="${key}">SUPPRIMER</button>
        </td>
      </tr>`;
    });

  renderTableRowsProgressively(body, rowsHtml, buildEmptyTableRow(body, "AUCUN COUT", 4), 24);

  bindReplacementCostActions();
}

function getStockMovementSignedQuantity(entry) {
  const qty = Math.max(1, Number.parseInt(String(entry?.quantite || 1), 10) || 1);
  const action = normalizeText(entry?.action);
  if (action === "ENTREE" || action === "AJUSTEMENT_PLUS") return qty;
  if (action === "SORTIE" || action === "AJUSTEMENT_MOINS") return -qty;
  return 0;
}

function getSignatureContextPerson(isMobileSignatureContext = false) {
  if (!isMobileSignatureContext) {
    return getCurrentPerson();
  }
  return getMobileSignatureTargetPerson();
}

function addAutoStockMovement(person, effect, action, motif, commentaire = "") {
  if (!state.data || !effect) {
    return;
  }
  if (!Array.isArray(state.data.stocksEffetsManuels)) {
    state.data.stocksEffetsManuels = [];
  }
  const typeEffet = normalizeText(effect?.typeEffet);
  if (!typeEffet) {
    return;
  }
  const site = typeEffet === "CARTE TURBOSELF"
    ? ALL_SITES_VALUE
    : normalizeText(effect?.siteReference || referenceSiteFromEffect(effect) || getPersonSiteLabel(person) || "SANS SITE");
  const designation = getStockGroupingDesignation(typeEffet, getEffectDisplayDesignation(effect));
  if (!designation) {
    return;
  }
  state.data.stocksEffetsManuels.push({
    id: getNextId("STKM", state.data.stocksEffetsManuels),
    typeEffet,
    site,
    referenceEffetId: String(effect?.referenceEffetId || ""),
    designation,
    action: normalizeText(action),
    quantite: 1,
    motif: normalizeText(motif),
    commentaire: normalizeText(commentaire),
    source: "AUTO",
    date: getTodayIsoDate(),
  });
}

function getStockSummaryRows() {
  const rowsByKey = new Map();
  const ensureRow = (typeEffet, site, designation) => {
    const key = `${normalizeText(typeEffet)}__${normalizeText(site)}__${normalizeText(designation)}`;
    if (!rowsByKey.has(key)) {
      rowsByKey.set(key, {
        key,
        typeEffet: normalizeText(typeEffet),
        site: normalizeText(site) || "SANS SITE",
        designation: normalizeText(designation),
        dotes: 0,
        rendus: 0,
        nonRendus: 0,
        perdus: 0,
        voles: 0,
        hs: 0,
        detruits: 0,
        manuelDelta: 0,
      });
    }
    return rowsByKey.get(key);
  };

  (state.data?.listes?.referencesEffets || [])
    .filter((reference) => isReferenceEffectActive(reference))
    .forEach((reference) => {
      const typeEffet = getReferenceEffectiveType(reference);
      const designation = getStockGroupingDesignation(typeEffet, getStockReferenceDesignation(reference));
      const sites = getReferenceSites(reference);
      if (!typeEffet) {
        return;
      }
      if (!sites.length) {
        ensureRow(typeEffet, "SANS SITE", designation);
        return;
      }
      sites.forEach((site) => {
        ensureRow(typeEffet, normalizeText(site) || "SANS SITE", designation);
      });
    });

  getAllEffects(state.data?.personnes || []).forEach(({ person, effect }) => {
    const linkedReference = effect?.referenceEffetId ? findReferenceById(effect.referenceEffetId) : null;
    const typeEffet = linkedReference ? getReferenceEffectiveType(linkedReference) : normalizeText(effect?.typeEffet || "");
    const site = normalizeText(getEffectDisplaySite(effect) || getPersonSiteLabel(person) || "SANS SITE");
    const designation = getStockGroupingDesignation(
      typeEffet,
      getEffectDisplayDesignation(effect) || effect?.designation || STOCK_EMPTY_DESIGNATION_LABEL
    );
    if (!typeEffet || !designation) {
      return;
    }
    const row = ensureRow(typeEffet, site, designation);
    const status = normalizeText(getEffectStatus(person, effect));
    row.dotes += 1;
    if (status === "RESTITUE") row.rendus += 1;
    else if (status === "NON RENDU") row.nonRendus += 1;
    else if (status === "PERDU") row.perdus += 1;
    else if (status === "VOL") row.voles += 1;
    else if (status === "HS") row.hs += 1;
    else if (status === "DETRUIT") row.detruits += 1;
  });

  (state.data?.stocksEffetsManuels || []).forEach((entry) => {
    if (!isStockMovementAllowedInSummary(entry)) {
      return;
    }
    const typeEffet = normalizeText(entry?.typeEffet || "");
    const site = normalizeText(entry?.site || "SANS SITE");
    const designation = getStockGroupingDesignation(typeEffet, entry?.designation || "");
    if (!typeEffet || !designation) {
      return;
    }
    const row = ensureRow(typeEffet, site, designation);
    row.manuelDelta += getStockMovementSignedQuantity(entry);
  });

  return Array.from(rowsByKey.values())
    .map((row) => {
      // Stock available = manual stock delta - entrusted effects + returned effects.
      // Lost/HS/vol/detruit remain consumed because they are included in entrusted effects and not returned.
      return {
        ...row,
        stockCourant: row.manuelDelta - row.dotes + row.rendus,
      };
    })
    .sort((left, right) => {
      const a = `${normalizeText(left.typeEffet)} ${normalizeText(left.site)} ${normalizeText(left.designation)}`;
      const b = `${normalizeText(right.typeEffet)} ${normalizeText(right.site)} ${normalizeText(right.designation)}`;
      return a.localeCompare(b, "fr");
    });
}

function renderStockMovementsTable() {
  const body = document.getElementById("stock-movements-table-body");
  if (!body) {
    return false;
  }
  const filters = state.stockTableFilters || { site: "", typeEffet: "", referenceEffetId: "" };
  const entries = (state.data?.stocksEffetsManuels || [])
    .filter((entry) => {
      if (!isStockMovementAllowedInSummary(entry)) {
        return false;
      }
      if (filters.site && filters.site !== ALL_SITES_VALUE && normalizeText(entry.site) !== filters.site) {
        return false;
      }
      if (filters.typeEffet && filters.typeEffet !== ALL_TYPES_VALUE && normalizeText(entry.typeEffet) !== filters.typeEffet) {
        return false;
      }
      if (
        filters.referenceEffetId &&
        filters.referenceEffetId !== ALL_DESIGNATIONS_VALUE
      ) {
        const syntheticReference = parseStockSyntheticReferenceValue(filters.referenceEffetId);
        if (syntheticReference) {
          return (
            normalizeText(entry.site) === syntheticReference.site &&
            normalizeText(entry.typeEffet) === syntheticReference.typeEffet &&
            normalizeText(entry.designation) === syntheticReference.designation
          );
        }
        if (String(entry.referenceEffetId || "") !== String(filters.referenceEffetId || "")) {
          return false;
        }
      }
      return true;
    })
    .slice()
    .sort((a, b) => String(b.date || "").localeCompare(String(a.date || ""), "fr"));
  const entriesSignature = entries
    .map((entry) => {
      const source = entry || {};
      return `${String(source.id || "")}|${normalizeText(source.site || "")}|${normalizeText(source.typeEffet || "")}|${normalizeText(
        source.designation || ""
      )}|${normalizeText(source.action || "")}|${String(source.quantite || 0)}|${normalizeText(source.referenceEffetId || "")}|${normalizeText(source.motif || "")}|${normalizeText(source.source || "")}|${String(source.date || "")}`;
    })
    .join("||");
  const nextStockMovementsSignature = `movements|${normalizeText(filters.site || "")}|${normalizeText(filters.typeEffet || "")}|${String(
    filters.referenceEffetId || ""
  )}|${String(entries.length)}|${entriesSignature}`;
  if (state.listRenderCache.stockMovements === nextStockMovementsSignature) {
    return false;
  }
  state.listRenderCache.stockMovements = nextStockMovementsSignature;

  if (!entries.length) {
    body.innerHTML = buildEmptyTableRow(body, "AUCUN MOUVEMENT", 9);
    return true;
  }
  const rowsHtml = entries.map((entry) => {
    const signedQty = getStockMovementSignedQuantity(entry);
    const qtyLabel = signedQty > 0 ? `+${signedQty}` : `${signedQty}`;
    const source = normalizeText(entry.source) === "AUTO" ? "AUTO" : "MANUEL";
    const canDelete = source !== "AUTO";
    return `<tr class="js-stock-movement-row" data-stock-movement-id="${escapeHtml(entry.id)}" data-site="${escapeHtml(entry.site || "")}" data-type-effet="${escapeHtml(entry.typeEffet || "")}" data-designation="${escapeHtml(entry.designation || "")}" data-reference-effet-id="${escapeHtml(String(entry.referenceEffetId || ""))}">
      <td>${escapeHtml(formatDate(entry.date) || entry.date || "-")}</td>
      <td>${escapeHtml(entry.site || "-")}</td>
      <td>${escapeHtml(entry.typeEffet || "-")}</td>
      <td>${escapeHtml(entry.designation || "-")}</td>
      <td>${escapeHtml(entry.action || "-")}</td>
      <td>${escapeHtml(qtyLabel)}</td>
      <td>${escapeHtml(entry.motif || "-")}</td>
      <td>${escapeHtml(source)}</td>
      <td>${canDelete ? `<button type="button" class="table-link js-delete-stock-movement" data-stock-movement-id="${escapeHtml(entry.id)}">SUPPRIMER</button>` : "-"}</td>
    </tr>`;
  });
  renderTableRowsProgressively(body, rowsHtml, buildEmptyTableRow(body, "AUCUN MOUVEMENT", 9), 24);
  bindStockMovementActions();
  return true;
}

function renderStockSummaryTable() {
  const body = document.getElementById("stock-summary-table-body");
  if (!body) {
    return false;
  }
  const rows = getFilteredStockSummaryRows();
  const rowsSignature = rows
    .map((row) =>
      [
        normalizeText(row.site || ""),
        normalizeText(row.typeEffet || ""),
        normalizeText(row.designation || ""),
        String(Number(row.dotes || 0)),
        String(Number(row.rendus || 0)),
        String(Number(row.nonRendus || 0)),
        String(Number(row.perdus || 0)),
        String(Number(row.voles || 0)),
        String(Number(row.hs || 0)),
        String(Number(row.detruits || 0)),
        String(Number(row.manuelDelta || 0)),
        String(Number(row.stockCourant || 0)),
      ].join("|")
    )
    .sort()
    .join("||");
  const totals = rows.reduce(
    (accumulator, row) => {
      accumulator.dotes += Number(row.dotes || 0);
      accumulator.rendus += Number(row.rendus || 0);
      accumulator.nonRendus += Number(row.nonRendus || 0);
      accumulator.perdus += Number(row.perdus || 0);
      accumulator.voles += Number(row.voles || 0);
      accumulator.hs += Number(row.hs || 0);
      accumulator.detruits += Number(row.detruits || 0);
      accumulator.manuelDelta += Number(row.manuelDelta || 0);
      accumulator.stockCourant += Number(row.stockCourant || 0);
      return accumulator;
    },
    { dotes: 0, rendus: 0, nonRendus: 0, perdus: 0, voles: 0, hs: 0, detruits: 0, manuelDelta: 0, stockCourant: 0 }
  );
  const nextStockSummarySignature = `summary|${normalizeText(state.stockTableFilters?.site || "")}|${normalizeText(
    state.stockTableFilters?.typeEffet || ""
  )}|${String(state.stockTableFilters?.referenceEffetId || "")}|${rows.length}|${rowsSignature}|${[
    totals.dotes,
    totals.rendus,
    totals.nonRendus,
    totals.perdus,
    totals.voles,
    totals.hs,
    totals.detruits,
    totals.manuelDelta,
    totals.stockCourant,
  ].join("|")}|${String(state.stockHighlightKey || "")}`;
  if (state.listRenderCache.stockSummary === nextStockSummarySignature) {
    return false;
  }
  state.listRenderCache.stockSummary = nextStockSummarySignature;

  if (!rows.length) {
    body.innerHTML = buildEmptyTableRow(body, "AUCUN STOCK CALCULE", 12);
    return true;
  }
  const formatSignedStockValue = (value) => {
    const numericValue = Number(value || 0);
    const variant = numericValue > 0 ? "positive" : numericValue < 0 ? "negative" : "zero";
    const icon = numericValue > 0 ? "🟢" : numericValue < 0 ? "🔴" : "🟠";
    return `<span class="stock-current stock-current--${variant}" aria-label="Stock courant ${numericValue}">
      <span class="stock-current__icon" aria-hidden="true">${icon}</span>
      <strong>${numericValue}</strong>
    </span>`;
  };
  const rowsHtml = rows.map((row) => {
    const manualDeltaLabel = row.manuelDelta > 0 ? `+${row.manuelDelta}` : String(row.manuelDelta);
    const rowKey = `${normalizeText(row.typeEffet)}__${normalizeText(row.site)}__${normalizeText(row.designation)}`;
    const focusClass = rowKey === String(state.stockHighlightKey || "") ? " stock-row-focus" : "";
    return `<tr class="js-stock-summary-row${focusClass}" data-site="${escapeHtml(row.site || "")}" data-type-effet="${escapeHtml(row.typeEffet || "")}" data-designation="${escapeHtml(row.designation || "")}">
      <td>${escapeHtml(row.site)}</td>
      <td>${escapeHtml(row.typeEffet)}</td>
      <td>${escapeHtml(row.designation)}</td>
      <td>${row.dotes}</td>
      <td>${row.rendus}</td>
      <td>${row.nonRendus}</td>
      <td>${row.perdus}</td>
      <td>${row.voles}</td>
      <td>${row.hs}</td>
      <td>${row.detruits}</td>
      <td>${escapeHtml(manualDeltaLabel)}</td>
      <td>${formatSignedStockValue(row.stockCourant)}</td>
    </tr>`;
  });
  const totalManualDeltaLabel = totals.manuelDelta > 0 ? `+${totals.manuelDelta}` : String(totals.manuelDelta);
  rowsHtml.push(`<tr class="table-total-row table-total-row--stock-summary">
    <td colspan="3">TOTAL</td>
    <td>${totals.dotes}</td>
    <td>${totals.rendus}</td>
    <td>${totals.nonRendus}</td>
    <td>${totals.perdus}</td>
    <td>${totals.voles}</td>
    <td>${totals.hs}</td>
    <td>${totals.detruits}</td>
    <td>${escapeHtml(totalManualDeltaLabel)}</td>
    <td>${formatSignedStockValue(totals.stockCourant)}</td>
  </tr>`);
  renderTableRowsProgressively(body, rowsHtml, buildEmptyTableRow(body, "AUCUN STOCK CALCULE", 12), 24);
  bindStockSummaryActions();
  if (state.stockHighlightKey) {
    window.setTimeout(() => {
      state.stockHighlightKey = "";
    }, 1200);
  }
  return true;
}

function renderTableRowsProgressively(body, rowsHtml, emptyMarkup = "", batchSize = 40) {
  if (!body) {
    return;
  }

  const token = `${Date.now()}-${Math.random()}`;
  body.dataset.renderToken = token;

  if (!rowsHtml.length) {
    body.innerHTML = emptyMarkup;
    return;
  }

  body.innerHTML = "";

  if (rowsHtml.length <= batchSize) {
    body.innerHTML = rowsHtml.join("");
    return;
  }

  let index = 0;

  const appendBatch = () => {
    if (body.dataset.renderToken !== token) {
      return;
    }

    body.insertAdjacentHTML("beforeend", rowsHtml.slice(index, index + batchSize).join(""));
    index += batchSize;

    if (index < rowsHtml.length) {
      window.requestAnimationFrame(appendBatch);
    }
  };

  window.requestAnimationFrame(appendBatch);
}

function syncSelectOptions(select, values, emptyLabel = "TOUS") {
  if (!(select instanceof HTMLSelectElement)) {
    return;
  }

  const signature = JSON.stringify([emptyLabel, ...values]);
  if (select.dataset.optionsSignature === signature) {
    return;
  }

  const currentValue = select.value;
  select.innerHTML = [`<option value="">${escapeHtml(emptyLabel)}</option>`]
    .concat(values.map((entry) => `<option value="${escapeHtml(entry)}">${escapeHtml(entry)}</option>`))
    .join("");
  select.value = currentValue;
  select.dataset.optionsSignature = signature;
}

function renderReferenceCounts() {
  const realSitesCount = (state.data?.listes?.sites || []).filter(
    (site) => normalizeText(site) !== ALL_SITES_VALUE
  ).length;
  const mapping = {
    "reference-count-sites": realSitesCount,
    "reference-count-typesPersonnel": state.data?.listes?.typesPersonnel?.length || 0,
    "reference-count-typesContrats": state.data?.listes?.typesContrats?.length || 0,
    "reference-count-fonctions": state.data?.listes?.fonctions?.length || 0,
    "reference-count-typesEffets": state.data?.listes?.typesEffets?.length || 0,
    "reference-count-causesRemplacement": state.data?.listes?.causesRemplacement?.length || 0,
    "reference-count-referencesEffets": state.data?.listes?.referencesEffets?.length || 0,
    "reference-count-coutsRemplacement": state.data?.listes?.coutsRemplacement?.length || 0,
    "reference-count-representantsSignataires": state.data?.listes?.representantsSignataires?.length || 0,
  };
  const nextReferenceCountsSignature = `reference-counts|${JSON.stringify(mapping)}`;
  if (state.listRenderCache.referenceCounts === nextReferenceCountsSignature) {
    return false;
  }
  state.listRenderCache.referenceCounts = nextReferenceCountsSignature;

  Object.entries(mapping).forEach(([id, value]) => {
    const node = document.getElementById(id);
    if (node) {
      setKpiCountAnimated(node, Number(value) || 0);
    }
  });
  return true;
}

function bindRepresentativeActions() {
  const body = document.getElementById("reference-representantsSignataires-body");
  if (!body || body.dataset.bound === "true") {
    return;
  }

  body.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    const deleteButton = target.closest(".js-delete-representative");
    if (deleteButton instanceof HTMLElement) {
      event.preventDefault();
      event.stopPropagation();
      deleteRepresentativeSignatory(deleteButton.dataset.representativeId || "");
      return;
    }

    const editButton = target.closest(".js-edit-representative");
    if (editButton instanceof HTMLElement) {
      event.preventDefault();
      event.stopPropagation();
      startEditRepresentativeSignatory(editButton.dataset.representativeId || "");
      return;
    }

    const row = target.closest(".js-representative-row");
    if (row instanceof HTMLElement) {
      startEditRepresentativeSignatory(row.dataset.representativeId || "");
    }
  });

  body.dataset.bound = "true";
}

function startEditRepresentativeSignatory(representativeId) {
  const representative = findRepresentativeById(representativeId);
  const form = document.getElementById("representative-signatory-form");
  if (!representative || !form) {
    return;
  }

  state.editingRepresentativeId = representativeId;
  form.elements.representativeName.value = representative.nom || "";
  form.elements.representativeFunction.value = representative.fonction || "";
  const submitButton = form.querySelector('button[type="submit"]');
  if (submitButton) {
    submitButton.textContent = "ENREGISTRER LA MODIFICATION";
  }
  form.scrollIntoView({ behavior: "smooth", block: "center" });
  showDataStatus(`REPRESENTANT EN COURS DE MODIFICATION : ${representative.nom || representative.fonction}`);
}

function deleteRepresentativeSignatory(representativeId) {
  if (!state.data?.listes?.representantsSignataires) {
    return;
  }

  const representative = findRepresentativeById(representativeId);
  if (!representative) {
    return;
  }

  const usage = getRepresentativeUsage(representativeId);
  if (usage > 0) {
    showDataStatus("SUPPRESSION BLOQUEE - REPRESENTANT DEJA UTILISE");
    window.alert("SUPPRESSION IMPOSSIBLE : REPRESENTANT DEJA UTILISE");
    return;
  }

  if (!window.confirm(`SUPPRIMER DEFINITIVEMENT : ${representative.nom || representative.fonction} ?`)) {
    return;
  }

  pushUndoSnapshot("SUPPRESSION REPRESENTANT");
  state.data.listes.representantsSignataires = state.data.listes.representantsSignataires.filter(
    (entry) => entry.id !== representativeId
  );
  if (state.editingRepresentativeId === representativeId) {
    state.editingRepresentativeId = "";
    const form = document.getElementById("representative-signatory-form");
    if (form) {
      form.reset();
    }
  }
  markDirty();
  schedulePageRender();
  showActionStatus("delete", `REPRESENTANT SUPPRIME : ${representative.nom || representative.fonction}`);
}

function bindReferenceListActions() {
  const bodyIds = [
    "reference-sites-body",
    "reference-typesPersonnel-body",
    "reference-typesContrats-body",
    "reference-fonctions-body",
    "reference-typesEffets-body",
    "reference-causesRemplacement-body",
  ];

  bodyIds.forEach((bodyId) => {
    const body = document.getElementById(bodyId);
    if (!body || body.dataset.bound === "true") {
      return;
    }

    body.addEventListener("click", (event) => {
      const target = event.target;
      if (!(target instanceof HTMLElement)) {
        return;
      }

      const startEdit = (listName, value) => {
        const form = document.querySelector(`.js-reference-list-form[data-list-name="${listName}"]`);
        const input = form?.querySelector('input[name="value"]');
        if (!form || !input) {
          return;
        }
        state.editingSimpleReference = { listName, originalValue: value };
        input.value = value;
        const submitButton = form.querySelector('button[type="submit"]');
        if (submitButton) {
          submitButton.textContent = "ENREGISTRER LA MODIFICATION";
        }
        input.focus();
        input.select();
        showDataStatus(`BASE EN COURS DE MODIFICATION : ${value}`);
      };

      const deleteButton = target.closest(".js-delete-reference-item");
      if (deleteButton instanceof HTMLElement) {
        event.preventDefault();
        event.stopPropagation();
        deleteSimpleReference(
          deleteButton.dataset.listName || "",
          normalizeText(deleteButton.dataset.value)
        );
        return;
      }

      const editButton = target.closest(".js-edit-reference-item");
      if (editButton instanceof HTMLElement) {
        event.preventDefault();
        event.stopPropagation();
        startEdit(editButton.dataset.listName || "", normalizeText(editButton.dataset.value));
        return;
      }

      const row = target.closest(".js-reference-item-row");
      if (row instanceof HTMLElement) {
        startEdit(row.dataset.listName || "", normalizeText(row.dataset.value));
      }
    });

    body.dataset.bound = "true";
  });
}

function bindReferenceEffectActions() {
  const body = document.getElementById("reference-effects-table-body");
  if (!body || body.dataset.bound === "true") {
    return;
  }

  body.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    const hardDeleteButton = target.closest(".js-hard-delete-reference-effect");
    if (hardDeleteButton instanceof HTMLElement) {
      event.preventDefault();
      event.stopPropagation();
      hardDeleteReferenceEffect(hardDeleteButton.dataset.referenceId || "");
      return;
    }

    const deleteButton = target.closest(".js-delete-reference-effect");
    if (deleteButton instanceof HTMLElement) {
      event.preventDefault();
      event.stopPropagation();
      deleteReferenceEffect(deleteButton.dataset.referenceId || "");
      return;
    }

    const editButton = target.closest(".js-edit-reference-effect");
    if (editButton instanceof HTMLElement) {
      event.preventDefault();
      event.stopPropagation();
      startEditReferenceEffect(editButton.dataset.referenceId || "");
      return;
    }

    const row = target.closest(".js-reference-effect-row");
    if (row instanceof HTMLElement) {
      startEditReferenceEffect(row.dataset.referenceId || "");
    }
  });

  body.dataset.bound = "true";
}

function startEditReplacementCost(costKey) {
  const form = document.getElementById("replacement-cost-form");
  if (!form || !state.data?.listes?.coutsRemplacement) {
    return;
  }

  const entry = state.data.listes.coutsRemplacement.find(
    (item) => getReplacementCostKey(item.typeEffet, item.cause, item.designation) === costKey
  );
  if (!entry) {
    return;
  }

  state.editingReplacementCostKey = costKey;
  form.elements.costTypeEffet.value = entry.typeEffet || "";
  form.elements.costCauseRemplacement.value = entry.cause || "";
  form.elements.costMontant.value = formatAmountWithEuro(entry.montant);
  const submitButton = form.querySelector('button[type="submit"]');
  if (submitButton) {
    submitButton.textContent = "ENREGISTRER LA MODIFICATION";
  }
  form.scrollIntoView({ behavior: "smooth", block: "center" });
  showDataStatus(`COUT EN COURS DE MODIFICATION : ${entry.typeEffet} / ${entry.cause}`);
}

function deleteReplacementCost(costKey) {
  if (!state.data?.listes?.coutsRemplacement) {
    return;
  }
  const entry = state.data.listes.coutsRemplacement.find(
    (item) => getReplacementCostKey(item.typeEffet, item.cause, item.designation) === costKey
  );
  if (!entry) {
    return;
  }
  if (!window.confirm(`SUPPRIMER DEFINITIVEMENT LE COUT : ${entry.typeEffet} / ${entry.cause} ?`)) {
    return;
  }
  pushUndoSnapshot("SUPPRESSION COUT");
  state.data.listes.coutsRemplacement = state.data.listes.coutsRemplacement.filter(
    (item) => getReplacementCostKey(item.typeEffet, item.cause, item.designation) !== costKey
  );
  if (state.editingReplacementCostKey === costKey) {
    state.editingReplacementCostKey = "";
    const form = document.getElementById("replacement-cost-form");
    if (form) {
      form.reset();
      const submitButton = form.querySelector('button[type="submit"]');
      if (submitButton) {
        submitButton.textContent = "ENREGISTRER LE COUT";
      }
    }
  }
  markDirty();
  schedulePageRender();
  showActionStatus("delete", `COUT SUPPRIME : ${entry.typeEffet} / ${entry.cause}`);
}

function bindReplacementCostActions() {
  const body = document.getElementById("replacement-costs-body");
  if (!body || body.dataset.bound === "true") {
    return;
  }

  body.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    const deleteButton = target.closest(".js-delete-replacement-cost");
    if (deleteButton instanceof HTMLElement) {
      event.preventDefault();
      event.stopPropagation();
      deleteReplacementCost(deleteButton.dataset.costKey || "");
      return;
    }

    const editButton = target.closest(".js-edit-replacement-cost");
    if (editButton instanceof HTMLElement) {
      event.preventDefault();
      event.stopPropagation();
      startEditReplacementCost(editButton.dataset.costKey || "");
      return;
    }

    const row = target.closest(".js-replacement-cost-row");
    if (row instanceof HTMLElement) {
      startEditReplacementCost(row.dataset.costKey || "");
    }
  });

  body.dataset.bound = "true";
}

function deleteSimpleReference(listName, value) {
  const list = state.data?.listes?.[listName];
  if (!Array.isArray(list)) {
    return;
  }
  const normalizeForList = (entry) =>
    listName === "causesRemplacement" ? normalizeReferenceCauseLabel(entry) : normalizeText(entry);

  const usage = getSimpleReferenceUsage(listName, value);
  if (usage > 0) {
    showDataStatus("SUPPRESSION BLOQUEE - VALEUR DEJA UTILISEE");
    window.alert("SUPPRESSION IMPOSSIBLE : VALEUR DEJA UTILISEE");
    return;
  }

  const confirmDelete = window.confirm(`SUPPRIMER DEFINITIVEMENT : ${value} ?`);
  if (!confirmDelete) {
    return;
  }

  pushUndoSnapshot("SUPPRESSION BASE");
  state.data.listes[listName] = list.filter((entry) => normalizeForList(entry) !== normalizeForList(value));
  if (
    state.editingSimpleReference &&
    state.editingSimpleReference.listName === listName &&
    normalizeForList(state.editingSimpleReference.originalValue) === normalizeForList(value)
  ) {
    state.editingSimpleReference = null;
  }
  markDirty();
  hydrateStaticLists();
  schedulePageRender();
  showActionStatus("delete", `BASE SUPPRIMEE : ${value}`);
}

function deleteReferenceEffect(referenceId) {
  if (!state.data?.listes?.referencesEffets) {
    return;
  }

  const reference = findReferenceById(referenceId);
  if (!reference) {
    return;
  }

  const nextActive = !isReferenceEffectActive(reference);
  const actionLabel = nextActive ? "REACTIVER" : "DESACTIVER";
  const confirmDelete = window.confirm(`${actionLabel} LA REFERENCE : ${reference.designation || referenceId} ?`);
  if (!confirmDelete) {
    return;
  }

  if (!nextActive) {
    const usage = getReferenceEffectCurrentUsage(referenceId);
    if (usage > 0) {
      showDataStatus("DESACTIVATION BLOQUEE - REFERENCE ENCORE EN DOTATION");
      window.alert("DESACTIVATION IMPOSSIBLE : CETTE REFERENCE EST ENCORE EN DOTATION.");
      return;
    }
  }

  pushUndoSnapshot("SUPPRESSION REFERENCE");
  reference.active = nextActive;
  if (!nextActive && Array.isArray(state.data.stocksEffetsManuels)) {
    state.data.stocksEffetsManuels = state.data.stocksEffetsManuels.filter(
      (entry) => !isStockMovementLinkedToReference(entry, reference)
    );
  }
  if (state.editingReferenceId === referenceId) {
    state.editingReferenceId = "";
    resetReferenceEffectForm();
  }
  markDirty();
  schedulePageRender();
  showActionStatus("update", `REFERENCE ${nextActive ? "REACTIVEE" : "DESACTIVEE"} : ${reference.designation || referenceId}`);
}

function hardDeleteReferenceEffect(referenceId) {
  if (!state.data?.listes?.referencesEffets) {
    return;
  }
  const reference = findReferenceById(referenceId);
  if (!reference) {
    return;
  }
  const usage = getReferenceEffectCurrentUsage(referenceId);
  if (usage > 0) {
    showDataStatus("SUPPRESSION DEFINITIVE BLOQUEE - REFERENCE ENCORE EN DOTATION");
    window.alert("SUPPRESSION DEFINITIVE IMPOSSIBLE : CETTE REFERENCE EST ENCORE EN DOTATION.");
    return;
  }
  const confirmDelete = window.confirm(
    `SUPPRIMER DEFINITIVEMENT : ${reference.designation || referenceId} ?`
  );
  if (!confirmDelete) {
    return;
  }
  pushUndoSnapshot("SUPPRESSION DEFINITIVE REFERENCE");
  state.data.listes.referencesEffets = state.data.listes.referencesEffets.filter(
    (entry) => entry.id !== referenceId
  );
  if (Array.isArray(state.data.stocksEffetsManuels)) {
    state.data.stocksEffetsManuels = state.data.stocksEffetsManuels.filter(
      (entry) => !isStockMovementLinkedToReference(entry, reference)
    );
  }
  if (state.editingReferenceId === referenceId) {
    state.editingReferenceId = "";
    resetReferenceEffectForm();
  }
  markDirty();
  schedulePageRender();
  showActionStatus("delete", `REFERENCE SUPPRIMEE DEFINITIVEMENT : ${reference.designation || referenceId}`);
}

function startEditReferenceEffect(referenceId) {
  const reference = findReferenceById(referenceId);
  const form = document.getElementById("reference-effect-form");
  if (!reference || !form) {
    return;
  }

  state.editingReferenceId = referenceId;
  form.elements.referenceSite.value = getReferenceSites(reference)[0] || "";
  form.elements.referenceTypeEffet.value = reference.typeEffet || "";
  form.elements.referenceDesignation.value = reference.designation || "";
  renderReferenceSitesSelector(getReferenceSites(reference));
  updateReferenceEffectFormMode(reference.typeEffet || "");
  const submitButton = form.querySelector('button[type="submit"]');
  if (submitButton) {
    submitButton.textContent = "ENREGISTRER LA MODIFICATION";
  }
  form.scrollIntoView({ behavior: "smooth", block: "center" });
  showDataStatus(`REFERENCE EN COURS DE MODIFICATION : ${reference.designation}`);
}

function resetReferenceEffectForm() {
  const form = document.getElementById("reference-effect-form");
  if (!form) {
    return;
  }
  form.reset();
  renderReferenceSitesSelector([]);
  updateReferenceEffectFormMode("");
}

function getSimpleReferenceUsage(listName, value) {
  const normalizedValue =
    listName === "causesRemplacement" ? normalizeReferenceCauseLabel(value) : normalizeText(value);
  const persons = state.data?.personnes || [];
  const references = state.data?.listes?.referencesEffets || [];

  if (listName === "sites") {
    return (
      persons.filter((person) => personHasSite(person, normalizedValue)).length +
      references.filter((reference) => referenceHasSite(reference, normalizedValue)).length
    );
  }
  if (listName === "typesPersonnel") {
    return persons.filter((person) => normalizeText(person.typePersonnel) === normalizedValue).length;
  }
  if (listName === "typesContrats") {
    return persons.filter((person) => normalizeText(person.typeContrat) === normalizedValue).length;
  }
  if (listName === "fonctions") {
    return persons.filter((person) => normalizeText(person.fonction) === normalizedValue).length;
  }
  if (listName === "typesEffets") {
    return (
      references.filter((reference) => normalizeText(reference.typeEffet) === normalizedValue).length +
      getAllEffects(persons).filter(({ effect }) => normalizeText(effect.typeEffet) === normalizedValue).length
    );
  }
  if (listName === "causesRemplacement") {
    return (state.data?.listes?.coutsRemplacement || []).filter(
      (entry) => normalizeReferenceCauseLabel(entry?.cause) === normalizedValue
    ).length;
  }
  return 0;
}

function getReferenceEffectUsage(referenceId) {
  return getAllEffects(state.data?.personnes || []).filter(
    ({ effect }) => String(effect.referenceEffetId || "") === String(referenceId)
  ).length;
}

function getReferenceEffectCurrentUsage(referenceId) {
  return getAllEffects(state.data?.personnes || []).filter(({ person, effect }) => {
    if (String(effect.referenceEffetId || "") !== String(referenceId)) {
      return false;
    }
    return normalizeText(getEffectStatus(person, effect)) !== "RESTITUE";
  }).length;
}

function referenceMatchesStockIdentity(reference, typeEffet, site, designation, requireActive = true) {
  if (!reference || (requireActive && !isReferenceEffectActive(reference))) {
    return false;
  }
  const normalizedType = normalizeText(typeEffet);
  const normalizedSite = normalizeText(site);
  const normalizedDesignation = getStockGroupingDesignation(normalizedType, designation);
  return (
    referenceMatchesType(reference, normalizedType) &&
    referenceHasSite(reference, normalizedSite) &&
    getStockGroupingDesignation(normalizedType, getStockReferenceDesignation(reference)) === normalizedDesignation
  );
}

function hasActiveStockReference(typeEffet, site, designation) {
  return (state.data?.listes?.referencesEffets || []).some((reference) =>
    referenceMatchesStockIdentity(reference, typeEffet, site, designation)
  );
}

function isStockMovementAllowedInSummary(entry) {
  const typeEffet = normalizeText(entry?.typeEffet || "");
  const site = normalizeText(entry?.site || "SANS SITE");
  const designation = getStockGroupingDesignation(typeEffet, entry?.designation || "");
  if (!typeEffet || !designation) {
    return false;
  }
  if (entry?.referenceEffetId) {
    const reference = findReferenceById(entry.referenceEffetId);
    return referenceMatchesStockIdentity(reference, typeEffet, site, designation);
  }
  if (typeUsesReferenceCatalog(typeEffet) && designation !== STOCK_EMPTY_DESIGNATION_LABEL) {
    return hasActiveStockReference(typeEffet, site, designation);
  }
  return true;
}

function isStockMovementLinkedToReference(entry, reference) {
  if (!reference) {
    return false;
  }
  if (String(entry?.referenceEffetId || "") === String(reference.id || "")) {
    return true;
  }
  return referenceMatchesStockIdentity(
    reference,
    entry?.typeEffet || "",
    entry?.site || "",
    entry?.designation || "",
    false
  );
}

function cascadeSimpleReferenceRename(listName, oldValue, newValue) {
  const oldNormalized =
    listName === "causesRemplacement" ? normalizeReferenceCauseLabel(oldValue) : normalizeText(oldValue);
  const nextValue =
    listName === "causesRemplacement" ? normalizeReferenceCauseLabel(newValue) : normalizeText(newValue);

  if (listName === "sites") {
    (state.data.personnes || []).forEach((person) => {
      person.sitesAffectation = normalizeSites(
        getPersonSites(person).map((site) =>
          normalizeText(site) === oldNormalized ? nextValue : normalizeText(site)
        )
      );
      person.site = getPersonSiteLabel(person);
    });
    (state.data.listes.referencesEffets || []).forEach((reference) => {
      reference.sitesAffectation = normalizeSites(
        getReferenceSites(reference).map((site) =>
          normalizeText(site) === oldNormalized ? nextValue : normalizeText(site)
        )
      );
      reference.site = getReferenceSiteLabel(reference);
    });
  }

  if (listName === "typesPersonnel") {
    (state.data.personnes || []).forEach((person) => {
      if (normalizeText(person.typePersonnel) === oldNormalized) {
        person.typePersonnel = nextValue;
      }
    });
  }

  if (listName === "typesContrats") {
    (state.data.personnes || []).forEach((person) => {
      if (normalizeText(person.typeContrat) === oldNormalized) {
        person.typeContrat = nextValue;
      }
    });
  }

  if (listName === "fonctions") {
    (state.data.personnes || []).forEach((person) => {
      if (normalizeText(person.fonction) === oldNormalized) {
        person.fonction = nextValue;
      }
    });
  }

  if (listName === "typesEffets") {
    (state.data.listes.referencesEffets || []).forEach((reference) => {
      if (normalizeText(reference.typeEffet) === oldNormalized) {
        reference.typeEffet = nextValue;
      }
    });
    (state.data.personnes || []).forEach((person) => {
      (person.effetsConfies || []).forEach((effect) => {
        if (normalizeText(effect.typeEffet) === oldNormalized) {
          effect.typeEffet = nextValue;
        }
      });
    });
  }

  if (listName === "causesRemplacement") {
    (state.data.listes.coutsRemplacement || []).forEach((entry) => {
      if (normalizeReferenceCauseLabel(entry.cause) === oldNormalized) {
        entry.cause = nextValue;
      }
    });
  }
}

function cascadeReferenceEffectUpdate(previous, nextReference) {
  (state.data.personnes || []).forEach((person) => {
    (person.effetsConfies || []).forEach((effect) => {
      if (String(effect.referenceEffetId || "") === String(previous.id)) {
        effect.typeEffet = nextReference.typeEffet;
        effect.designation = nextReference.designation;
      }
    });
  });
}

function sortListValues(values) {
  values.sort((a, b) => normalizeText(a).localeCompare(normalizeText(b), "fr"));
}

function sortReferenceEffects() {
  if (!state.data?.listes?.referencesEffets) {
    return;
  }
  state.data.listes.referencesEffets.sort((a, b) => {
    const left = `${normalizeText(getReferenceSiteLabel(a))} ${normalizeText(a.typeEffet)} ${normalizeText(a.designation)}`;
    const right = `${normalizeText(getReferenceSiteLabel(b))} ${normalizeText(b.typeEffet)} ${normalizeText(b.designation)}`;
    return left.localeCompare(right, "fr");
  });
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function tryCloseCurrentWindow() {
  try {
    window.close();
  } catch (error) {
    // ignore close errors
  }
  if (window.closed) {
    return;
  }
  try {
    window.open("", "_self");
    window.close();
  } catch (error) {
    // ignore close errors
  }
  if (window.closed) {
    return;
  }
  if (window.history.length > 1) {
    try {
      window.history.back();
    } catch (error) {
      // ignore history errors
    }
  }
  if (window.closed) {
    return;
  }
  try {
    if (document.body.dataset.page === "mobile-signature") {
      window.location.replace("about:blank");
      return;
    }
    const currentUrl = new URL(window.location.href);
    const fallbackPath = currentUrl.pathname.replace(/[^/]*$/, "index.html");
    const fallbackUrl = `${currentUrl.origin}${fallbackPath}`;
    if (window.location.href !== fallbackUrl) {
      window.location.replace(fallbackUrl);
    }
  } catch (error) {
    // ignore redirect errors
  }
}

function scheduleCloseAttempts() {
  tryCloseCurrentWindow();
  [250, 800, 1600].forEach((delay) => {
    window.setTimeout(() => {
      if (!window.closed) {
        tryCloseCurrentWindow();
      }
    }, delay);
  });
}

function normalizeArchiveTypeLabel(typeDocument) {
  const normalized = normalizeText(typeDocument);
  if (normalized === "ARRIVEE" || normalized === "ARRIVAL" || normalized === "ENTREE") {
    return "ARRIVEE";
  }
  if (normalized === "SORTIE" || normalized === "EXIT") {
    return "SORTIE";
  }
  return normalized ? String(typeDocument || "").trim() : "";
}

function buildPeopleNameIndex(data) {
  const index = new Map();
  const people = Array.isArray(data?.personnes) ? data.personnes : [];
  people.forEach((person) => {
    const key = `${normalizeText(person?.nom)}|${normalizeText(person?.prenom)}`;
    if (key === "|") {
      return;
    }
    if (!index.has(key)) {
      index.set(key, String(person?.id || ""));
    }
  });
  return index;
}

function protectAndNormalizeArchivesInPlace(data) {
  if (!data || !Array.isArray(data.documentsArchives)) {
    return;
  }
  const peopleById = new Map(
    (Array.isArray(data.personnes) ? data.personnes : [])
      .map((person) => [String(person?.id || ""), person])
      .filter(([id]) => Boolean(id))
  );
  const peopleByName = buildPeopleNameIndex(data);
  data.documentsArchives.forEach((entry) => {
    if (!entry || typeof entry !== "object") {
      return;
    }
    entry.typeDocument = normalizeArchiveTypeLabel(entry.typeDocument || "");
    const currentPersonId = String(entry.personId || "");
    const currentPerson = currentPersonId ? peopleById.get(currentPersonId) : null;
    if (!currentPerson) {
      const fallbackKey = `${normalizeText(entry.nom)}|${normalizeText(entry.prenom)}`;
      const fallbackId = peopleByName.get(fallbackKey) || "";
      if (fallbackId) {
        entry.personId = fallbackId;
      }
      return;
    }
    if (!String(entry.nom || "").trim()) {
      entry.nom = currentPerson.nom || "";
    }
    if (!String(entry.prenom || "").trim()) {
      entry.prenom = currentPerson.prenom || "";
    }
  });
}

function buildSignatureProtectionSnapshot(data) {
  const snapshot = new Map();
  const personnes = Array.isArray(data?.personnes) ? data.personnes : [];
  personnes.forEach((person) => {
    const personId = String(person?.id || "");
    if (!personId) {
      return;
    }
    ["arrival", "exit"].forEach((docType) => {
      ["personnel", "representant"].forEach((signer) => {
        const value = String(getSignatureValue(person, docType, signer) || "");
        const validatedAt = String(getSignatureValidationDate(person, docType, signer) || "");
        if (!value || !validatedAt) {
          return;
        }
        const key = `${personId}|${docType}|${signer}`;
        snapshot.set(key, {
          value,
          validatedAt,
          storageRef: String(getSignatureStorageRef(person, docType, signer) || ""),
          storagePublicUrl: String(getSignatureStoragePublicUrl(person, docType, signer) || ""),
        });
      });
    });
  });
  return snapshot;
}

function enforceProtectedSignaturesInPlace(data, signatureSnapshot) {
  if (!data || !(signatureSnapshot instanceof Map) || signatureSnapshot.size === 0) {
    return;
  }
  const personnes = Array.isArray(data.personnes) ? data.personnes : [];
  const byId = new Map(
    personnes
      .map((person) => [String(person?.id || ""), person])
      .filter(([id]) => Boolean(id))
  );
  signatureSnapshot.forEach((savedSignature, key) => {
    const [personId, docType, signer] = key.split("|");
    const person = byId.get(String(personId || ""));
    if (!person) {
      return;
    }
    const existingValue = String(getSignatureValue(person, docType, signer) || "");
    const existingValidatedAt = String(getSignatureValidationDate(person, docType, signer) || "");
    if (existingValue && existingValidatedAt) {
      return;
    }
    setSignatureValue(
      person,
      docType,
      signer,
      savedSignature.value,
      savedSignature.validatedAt,
      savedSignature.storageRef,
      savedSignature.storagePublicUrl
    );
  });
}

function buildArchiveProtectionSnapshot(data) {
  const snapshot = new Map();
  const archives = Array.isArray(data?.documentsArchives) ? data.documentsArchives : [];
  archives.forEach((entry) => {
    const typeKey = getArchiveDocTypeKey(entry?.typeDocument);
    const personId = String(entry?.personId || "").trim();
    const pdfPath = String(entry?.pdfPath || "").trim();
    const isSigned = normalizeText(entry?.statutSignature) === "SIGNE";
    if (!typeKey || !personId || !pdfPath || !isSigned) {
      return;
    }
    const key = `${personId}|${typeKey}`;
    const current = snapshot.get(key);
    const currentMs = Date.parse(String(current?.dateArchivage || "")) || 0;
    const nextMs = Date.parse(String(entry?.dateArchivage || "")) || 0;
    if (!current || nextMs >= currentMs) {
      snapshot.set(key, { ...entry, typeDocument: typeKey });
    }
  });
  return snapshot;
}

function enforceProtectedArchivesInPlace(data, archiveSnapshot) {
  if (!data || !(archiveSnapshot instanceof Map) || archiveSnapshot.size === 0) {
    return;
  }
  if (!Array.isArray(data.documentsArchives)) {
    data.documentsArchives = [];
  }
  archiveSnapshot.forEach((savedEntry, key) => {
    const [personId, typeKey] = String(key || "").split("|");
    if (!personId || !typeKey) {
      return;
    }
    const hasProtectedArchive = data.documentsArchives.some((entry) => {
      if (String(entry?.personId || "").trim() !== personId) {
        return false;
      }
      if (getArchiveDocTypeKey(entry?.typeDocument) !== typeKey) {
        return false;
      }
      if (normalizeText(entry?.statutSignature) !== "SIGNE") {
        return false;
      }
      return Boolean(String(entry?.pdfPath || "").trim());
    });
    if (!hasProtectedArchive) {
      data.documentsArchives.push({ ...savedEntry, typeDocument: typeKey });
    }
  });
}

function appendSaveAuditEntry(entry) {
  try {
    const raw = localStorage.getItem(SAVE_AUDIT_LOG_KEY) || "[]";
    const parsed = JSON.parse(raw);
    const entries = Array.isArray(parsed) ? parsed : [];
    entries.push({
      at: getCurrentSignatureTimestamp(),
      ...entry,
    });
    const compact = entries.slice(-MAX_SAVE_AUDIT_ENTRIES);
    localStorage.setItem(SAVE_AUDIT_LOG_KEY, JSON.stringify(compact));
  } catch (error) {
    // Silent: audit must never block save.
  }
}

function computeDataPersistenceSignature(data) {
  try {
    return JSON.stringify(data ?? null);
  } catch (error) {
    return "";
  }
}

async function runHostedLocalControlAfterSave() {
  if (getDataBackendMode() !== "LOCAL_API") {
    return false;
  }
  try {
    state.hostedSyncState = "checking";
    state.hostedSyncDetails = "";
    renderHostedSyncUi();
    const response = await fetch("/api/local/check-before-send", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
    });
    const payload = await response.json().catch(() => ({}));
    if (response.ok && payload?.ok) {
      state.hostedSyncState = "pending";
      state.hostedSyncDetails = "";
      showDataStatus("SAUVEGARDE LOCALE OK - CONTROLE LOCAL OK");
      return true;
    }
    const localOk = Boolean(payload?.localOk);
    const pdfOk = Boolean(payload?.pdfIntegrityOk);
    const supabaseOk = payload?.supabaseReachable !== false;
    state.hostedSyncState = supabaseOk && (!localOk || !pdfOk) ? "blocked" : "inaccessible";
    state.hostedSyncDetails = JSON.stringify(payload?.technicalDetails || payload || {}, null, 2);
    showDataStatus("SAUVEGARDE LOCALE OK - CONTROLE LOCAL A CORRIGER");
    return false;
  } catch (error) {
    state.hostedSyncState = "inaccessible";
    state.hostedSyncDetails = String(error?.message || error || "");
    showDataStatus("SAUVEGARDE LOCALE OK - CONTROLE LOCAL INDISPONIBLE");
    return false;
  } finally {
    renderHostedSyncUi();
  }
}

async function saveDataToFile(options = {}) {
  if (!state.data) {
    showDataStatus("AUCUNE DONNEE A SAUVEGARDER");
    return;
  }

  const isEventCall =
    options &&
    typeof options === "object" &&
    typeof options.preventDefault === "function";
  const resolvedOptions = isEventCall ? {} : options;

  const {
    silent = false,
    reloadAfter = null,
    successText = "data.json MIS A JOUR",
    alertText = "",
    closeAfterAlert = false,
    promptDownload = !silent,
    reloadOnConflict = true,
    throwOnConflict = false,
    autoPushHosted = getDataBackendMode() === "LOCAL_API",
  } = resolvedOptions;
  const shouldReloadAfter =
    typeof reloadAfter === "boolean" ? reloadAfter : getDataBackendMode() !== "SUPABASE";

  if (state.saveInFlight) {
    showDataStatus("SAUVEGARDE EN COURS...");
    return;
  }

  const currentDataSignature = computeDataPersistenceSignature(state.data);
  const persistedSignatureKnown =
    typeof state.lastPersistedDataSignature === "string" && state.lastPersistedDataSignature.length > 0;
  const hasEffectiveChange = persistedSignatureKnown
    ? currentDataSignature !== state.lastPersistedDataSignature
    : Boolean(state.isDirty);
  if (!hasEffectiveChange) {
    state.isDirty = false;
    state.saveButtonLatchedDirty = false;
    renderDirtyState();
    showDataStatus("AUCUNE MODIFICATION DETECTEE. RIEN A SAUVEGARDER.");
    if (getDataBackendMode() === "LOCAL_API") renderHostedSyncUi();
    appendSaveAuditEntry({
      outcome: "noop",
      source: "NO_CHANGE",
      mode: getDataBackendMode(),
    });
    if (getDataBackendMode() === "LOCAL_API" && autoPushHosted) {
      await runHostedSyncPushFromUi({ confirmBeforeSend: false, silent: true });
    }
    return;
  }

  const downloadDataJson = () => {
    try {
      const blob = new Blob([JSON.stringify(state.data, null, 2)], {
        type: "application/json;charset=utf-8",
      });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = "data.json";
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 800);
    } catch (error) {
      console.error(error);
      window.alert("TELECHARGEMENT DE data.json IMPOSSIBLE");
    }
  };

  try {
    state.saveInFlight = true;
    const signatureSnapshot = buildSignatureProtectionSnapshot(state.data);
    const archiveSnapshot = buildArchiveProtectionSnapshot(state.data);
    protectAndNormalizeArchivesInPlace(state.data);
    enforceProtectedSignaturesInPlace(state.data, signatureSnapshot);
    enforceProtectedArchivesInPlace(state.data, archiveSnapshot);
    const mode = getDataBackendMode();
      let saveStatusText = successText;
      let saveAlertText = alertText || "data.json A ETE MIS A JOUR";
      let saveSource = "LOCAL";
      if (mode === "SUPABASE") {
        await saveSupabasePayloadWithRetry(state.data, 3);
        saveStatusText = "DONNEES SAUVEGARDEES VIA BACKEND";
        saveAlertText = alertText || "DONNEES MISES A JOUR VIA BACKEND";
        saveSource = "BACKEND";
      } else if (mode === "LOCAL_API") {
        const response = await fetch("/api/state/save", {
          method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ data: state.data }),
      });

        if (!response.ok) {
          throw new Error("Sauvegarde locale impossible");
        }
        saveStatusText = "DONNEES SAUVEGARDEES EN LOCAL UNIQUEMENT";
        saveAlertText = alertText || "DONNEES SAUVEGARDEES EN LOCAL UNIQUEMENT";
        saveSource = "BACKEND LOCAL";
        state.hostedSyncState = "pending";
      } else {
      throw new Error("SUPABASE NON CONFIGURE");
    }

    clearWorkingData();
    state.isDirty = false;
    clearUndoStack();
    renderDirtyState();
    state.lastSaveInfo = {
      at: getCurrentSignatureTimestamp(),
      source: saveSource,
    };
    state.lastPersistedDataSignature = computeDataPersistenceSignature(state.data);
    const saveConfirmation = `SAUVEGARDEE LE ${formatCurrentUiTimestamp()} - SOURCE: ${saveSource}`;
    showDataStatus(saveConfirmation);
    if (mode === "LOCAL_API") {
      showDataStatus("DONNEES SAUVEGARDEES EN LOCAL - CONTROLE LOCAL EN COURS...");
      renderHostedSyncUi();
      const localControlOk = await runHostedLocalControlAfterSave();
      if (localControlOk && autoPushHosted) {
        await runHostedSyncPushFromUi({ confirmBeforeSend: false, silent: true });
      }
    }
    appendSaveAuditEntry({
      outcome: "ok",
      source: saveSource,
      mode,
      personnes: Array.isArray(state.data?.personnes) ? state.data.personnes.length : 0,
      archives: Array.isArray(state.data?.documentsArchives) ? state.data.documentsArchives.length : 0,
    });
    if (!silent) {
      window.alert(saveAlertText);
      pulseSaveButtons();
      if (promptDownload && saveSource === "LOCAL" && document.body.dataset.page !== "mobile-signature") {
        downloadDataJson();
      }
      if (closeAfterAlert) {
        scheduleCloseAttempts();
      }
    } else {
      pulseSaveButtons();
    }
    if (shouldReloadAfter) {
      try {
        await reloadData(mode === "SUPABASE" ? "RELECTURE DES DONNEES SUPABASE..." : "RELECTURE DES DONNEES LOCALES...");
      } catch (reloadError) {
        console.error(reloadError);
        showDataStatus("SAUVEGARDE OK - RELECTURE IMPOSSIBLE");
      }
    }
    } catch (error) {
      console.error(error);
      if (
        error?.code === "BACKEND_SAVE_REQUIRED" ||
        error?.code === "BACKEND_AUTH_REQUIRED" ||
        String(error?.message || "") === "BACKEND_SAVE_REQUIRED" ||
        String(error?.message || "") === "BACKEND_AUTH_REQUIRED" ||
        String(error?.message || "").includes("EDGE_API_NOT_CONFIGURED")
      ) {
        showDataStatus("SAUVEGARDE BLOQUEE: CONNEXION REQUISE");
        if (!silent) {
          window.alert("SAUVEGARDE BLOQUEE : CONNEXION SUPABASE REQUISE.");
        }
        return;
      }
      if (isSaveConflictError(error)) {
      if (throwOnConflict) {
        throw error;
      }
      const currentPage = String(document?.body?.dataset?.page || "");
      const isDocumentSignaturePage =
        currentPage === "arrival-document" || currentPage === "exit-document";
      if (reloadOnConflict && (getDataBackendMode() === "SUPABASE" || isSupabaseConfigured())) {
        try {
          await reloadData("CONFLIT DETECTE - RECHARGEMENT DES DONNEES DISTANTES...");
        } catch (refreshError) {
          console.error(refreshError);
        }
      }
      showDataStatus(
        isDocumentSignaturePage
          ? "CONFLIT DETECTE - DONNEES RECHARGEES"
          : "CONFLIT DE SAUVEGARDE - RECHARGER PUIS REESSAYER"
      );
      if (!silent && !isDocumentSignaturePage) {
        window.alert(error.message);
      }
      return;
    }
    showDataStatus("SAUVEGARDE IMPOSSIBLE");
    const compactErrorMessage = String(error?.message || "").replace(/\s+/g, " ").trim().slice(0, 180);
    const isMobileSignaturePage = String(document?.body?.dataset?.page || "") === "mobile-signature";
    if (isMobileSignaturePage && compactErrorMessage) {
      showDataStatus(`SAUVEGARDE IMPOSSIBLE (${compactErrorMessage})`);
    }
    appendSaveAuditEntry({
      outcome: "error",
      message: String(error?.message || "UNKNOWN_ERROR").slice(0, 200),
      mode: getDataBackendMode(),
    });
    if (!silent) {
      if (isMobileSignaturePage && compactErrorMessage) {
        window.alert(`SAUVEGARDE IMPOSSIBLE\n${compactErrorMessage}`);
      } else {
        window.alert("SAUVEGARDE IMPOSSIBLE");
      }
    }
  } finally {
    state.saveInFlight = false;
  }

}

function resetUiWithoutData() {
  const targets = [
    { id: "overview-table-body", colspan: 14 },
    { id: "global-table-body", colspan: 14 },
    { id: "sheet-effects-body", colspan: 11 },
    { id: "reference-sites-body", colspan: 2 },
    { id: "reference-typesPersonnel-body", colspan: 2 },
    { id: "reference-typesContrats-body", colspan: 2 },
    { id: "reference-fonctions-body", colspan: 2 },
    { id: "reference-typesEffets-body", colspan: 2 },
    { id: "reference-causesRemplacement-body", colspan: 2 },
    { id: "reference-effects-table-body", colspan: 5 },
    { id: "replacement-costs-body", colspan: 4 },
    { id: "stock-movements-table-body", colspan: 9 },
    { id: "stock-summary-table-body", colspan: 12 },
  ];

  targets.forEach(({ id, colspan }) => {
    const node = document.getElementById(id);
    if (node) {
      node.innerHTML = buildEmptyTableRow(node, "DONNEES NON DISPONIBLES", colspan);
    }
  });

  [
    "kpi-personnes-en-poste",
    "kpi-effets-confies",
    "kpi-effets-non-rendus",
    "reference-count-sites",
    "reference-count-typesPersonnel",
    "reference-count-typesContrats",
    "reference-count-fonctions",
    "reference-count-typesEffets",
    "reference-count-causesRemplacement",
    "reference-count-referencesEffets",
    "reference-count-coutsRemplacement",
    "reference-count-representantsSignataires",
    "stock-type-kpi-value-1",
    "stock-type-kpi-value-2",
    "stock-type-kpi-value-3",
    "stock-type-kpi-value-4",
    "stock-type-kpi-value-5",
  ].forEach((id) => {
    const node = document.getElementById(id);
    if (node) node.textContent = "0";
  });

  renderPersonSheet("");
}

function getHostedSyncSimpleText() {
  if (state.hostedSyncInFlight) {
    return "Synchronisation en cours...";
  }
  switch (String(state.hostedSyncState || "")) {
    case "up_to_date":
      return "Heberge a jour";
    case "pending":
      return "Envoi utile seulement si modification locale volontaire";
    case "inaccessible":
      return "Heberge inaccessible";
    case "blocked":
      return "Controle local a corriger avant tout envoi";
    default:
      return "Statut heberge inconnu";
  }
}

function ensureHostedSyncUi() {
  if (getDataBackendMode() !== "LOCAL_API") return;
  if (isPdfRenderMode()) return;
  let node = document.getElementById("dotations-sync-banner");
  if (!node) {
    node = document.createElement("div");
    node.id = "dotations-sync-banner";
    node.style.cssText = "margin:10px 0;padding:10px 12px;border:1px solid #c7d2fe;background:#eef2ff;border-radius:10px;display:flex;gap:10px;align-items:center;flex-wrap:wrap;font-size:13px;";
    node.innerHTML = [
      '<strong style="margin-right:8px;">Etat Dotations</strong>',
      '<span id="dotations-sync-local">Local : -</span>',
      '<span id="dotations-sync-remote" class="dotations-sync-remote">Heberge : -</span>',
      '<button id="dotations-sync-refresh-btn" class="button button--secondary" type="button" style="height:28px;min-height:28px;padding:0 10px;border-radius:999px;font-size:11px;display:inline-flex;align-items:center;line-height:1;">Reverifier heberge</button>',
      '<button id="dotations-sync-push-btn" class="button button--secondary" type="button" style="height:28px;min-height:28px;padding:0 10px;border-radius:999px;font-size:11px;display:inline-flex;align-items:center;line-height:1;">Envoyer vers l\'heberge</button>',
      '<button id="dotations-sync-help-btn" class="button button--secondary" type="button" style="display:none;height:28px;min-height:28px;padding:0 10px;border-radius:999px;font-size:11px;align-items:center;line-height:1;">Voir ce qu\'il faut faire</button>',
      '<button id="dotations-close-btn" class="button button--secondary" type="button" style="height:28px;min-height:28px;padding:0 10px;border-radius:999px;font-size:11px;display:inline-flex;align-items:center;line-height:1;">Fermer Dotations</button>',
      '<details id="dotations-sync-details" style="width:100%;margin-top:6px;display:none;"><summary>Details techniques</summary><pre id="dotations-sync-details-text" style="white-space:pre-wrap;margin:8px 0 0 0;"></pre></details>'
    ].join("");
    const header = document.querySelector(".page-header");
    if (header && header.parentElement) {
      header.parentElement.insertBefore(node, header.nextSibling);
    } else {
      document.body.insertBefore(node, document.body.firstChild);
    }
  }
  const refreshBtn = document.getElementById("dotations-sync-refresh-btn");
  if (refreshBtn && !refreshBtn.dataset.bound) {
    refreshBtn.dataset.bound = "1";
    refreshBtn.addEventListener("click", () => {
      void refreshHostedSyncStatusOnLocalOpen();
    });
  }
  const pushBtn = document.getElementById("dotations-sync-push-btn");
  if (pushBtn && !pushBtn.dataset.bound) {
    pushBtn.dataset.bound = "1";
    pushBtn.addEventListener("click", () => runHostedSyncPushFromUi());
  }
  const helpBtn = document.getElementById("dotations-sync-help-btn");
  if (helpBtn && !helpBtn.dataset.bound) {
    helpBtn.dataset.bound = "1";
    helpBtn.addEventListener("click", () => openRescueActionModal());
  }
  const closeBtn = document.getElementById("dotations-close-btn");
  if (closeBtn && !closeBtn.dataset.bound) {
    closeBtn.dataset.bound = "1";
    closeBtn.addEventListener("click", () => openCloseDotationsModal());
  }
}

function ensureRescueModalUi() {
  let modal = document.getElementById("dotations-rescue-modal");
  if (modal) return modal;
  modal = document.createElement("div");
  modal.id = "dotations-rescue-modal";
  modal.style.cssText = "position:fixed;inset:0;background:rgba(15,23,42,.45);display:none;align-items:center;justify-content:center;z-index:9999;padding:16px;";
  modal.innerHTML = [
    '<div role="dialog" aria-modal="true" aria-labelledby="dotations-rescue-title" style="width:min(760px,100%);max-height:90vh;overflow:auto;background:#f8fbff;border:1px solid #dbe6f5;border-radius:14px;box-shadow:0 12px 36px rgba(15,23,42,.18);padding:18px 18px 14px 18px;">',
    '<div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px;">',
    '<h3 id="dotations-rescue-title" style="margin:0;font-size:20px;line-height:1.2;color:#0f172a;">Une action est necessaire</h3>',
    '<button id="dotations-rescue-close" type="button" class="button button--secondary" style="height:30px;padding:0 12px;border-radius:10px;font-size:12px;">Fermer</button>',
    "</div>",
    '<p id="dotations-rescue-summary" style="margin:8px 0 12px 0;color:#334155;"></p>',
    '<div style="background:#ffffff;border:1px solid #e2e8f0;border-radius:12px;padding:12px;">',
    '<strong style="display:block;color:#0f172a;">Ce que vous devez faire</strong>',
    '<ol id="dotations-rescue-steps" style="margin:8px 0 0 18px;padding:0;color:#334155;"></ol>',
    "</div>",
    '<div style="margin-top:12px;">',
    '<strong style="display:block;color:#0f172a;margin-bottom:8px;">Action recommandee</strong>',
    '<button id="dotations-rescue-primary" type="button" class="button button--primary" style="height:34px;padding:0 14px;border-radius:10px;font-weight:600;">Reessayer</button>',
    '<span id="dotations-rescue-primary-status" style="margin-left:10px;color:#334155;font-size:12px;"></span>',
    "</div>",
    '<div style="margin-top:12px;">',
    '<strong style="display:block;color:#0f172a;margin-bottom:8px;">Autres actions</strong>',
    '<div id="dotations-rescue-secondary-actions" style="display:flex;gap:8px;flex-wrap:wrap;"></div>',
    "</div>",
    '<p style="margin:12px 0 0 0;color:#64748b;font-size:12px;">Ne pas envoyer vers l\'heberge tant que le controle local n\'est pas repasse OK.</p>',
    '<details id="dotations-rescue-details" style="margin-top:10px;">',
    "<summary>Details techniques</summary>",
    '<pre id="dotations-rescue-details-text" style="margin-top:8px;white-space:pre-wrap;background:#fff;border:1px solid #e2e8f0;border-radius:8px;padding:8px;font-size:12px;"></pre>',
    "</details>",
    '<div id="dotations-rescue-run-result" style="margin-top:10px;font-size:12px;color:#334155;"></div>',
    '<div style="margin-top:12px;display:flex;gap:8px;align-items:center;">',
    '<button id="dotations-rescue-reload" type="button" class="button button--primary" style="display:none;height:32px;padding:0 12px;border-radius:10px;">Recharger Dotations</button>',
    '<button id="dotations-rescue-close-secondary" type="button" class="button button--secondary" style="height:32px;padding:0 12px;border-radius:10px;">Fermer</button>',
    "</div>",
    "</div>",
  ].join("");
  document.body.appendChild(modal);
  const closeBtn = document.getElementById("dotations-rescue-close");
  if (closeBtn) {
    closeBtn.addEventListener("click", () => {
      modal.style.display = "none";
    });
  }
  const closeSecondaryBtn = document.getElementById("dotations-rescue-close-secondary");
  if (closeSecondaryBtn) {
    closeSecondaryBtn.addEventListener("click", () => {
      modal.style.display = "none";
    });
  }
  const reloadBtn = document.getElementById("dotations-rescue-reload");
  if (reloadBtn) {
    reloadBtn.addEventListener("click", () => {
      window.location.assign(`${window.location.origin}/`);
    });
  }
  modal.addEventListener("click", (event) => {
    if (event.target === modal) {
      modal.style.display = "none";
    }
  });
  return modal;
}

function parseHostedSyncDetails(rawOverride = null) {
  const raw = rawOverride == null ? String(state.hostedSyncDetails || "").trim() : String(rawOverride || "").trim();
  if (!raw) return {};
  try {
    return JSON.parse(raw);
  } catch (error) {
    return { raw };
  }
}

function buildRescuePlanFromState(rawDetailsOverride = null, stateOverride = null) {
  const details = parseHostedSyncDetails(rawDetailsOverride);
  const hostedState = String(stateOverride || state.hostedSyncState || "");
  const hasMissingPdf =
    Number(details?.missingActiveFiles || 0) > 0 ||
    /ACTIVE_FILE_MISSING|MISSING_ACTIVE_FILE/i.test(String(details?.errors || ""));
  const hasOrphans = Number(details?.orphansOutsideQuarantine || 0) > 0;
  const supabaseDown =
    hostedState === "inaccessible" ||
    /heberge.*inaccessible|supabase.*inaccessible|fetch failed/i.test(String(details?.raw || state.hostedSyncDetails || ""));
  const hasOpenableIssue = /ARCHIVE_NOT_OPENABLE|VALID_WITHOUT_LOCALPATH|VALID_WITHOUT_ENRICHED_FIELDS/i.test(
    String(details?.errors || state.hostedSyncDetails || "")
  );

  if (hasMissingPdf) {
    return {
      title: "Des documents doivent etre verifies",
      summary: "Certains documents existent dans les donnees, mais leurs PDF ne sont pas disponibles en local.",
      steps: [
        "Restaurer les PDF depuis les sauvegardes locales.",
        "Relancer le controle local.",
        "Envoyer vers l'heberge uniquement si une vraie modification locale doit etre publiee.",
      ],
      primary: { label: "Restaurer les PDF", kind: "repair", action: "repair_missing_pdf" },
    };
  }
  if (hasOrphans) {
    return {
      title: "Des fichiers doivent etre ranges",
      summary: "Des PDF non rattaches ont ete detectes hors quarantaine.",
      steps: [
        "Mettre les orphelins en quarantaine.",
        "Relancer le controle local.",
        "Verifier que l'envoi n'est plus bloque.",
      ],
      primary: { label: "Mettre en quarantaine", kind: "repair", action: "repair_orphans_quarantine" },
    };
  }
  if (hasOpenableIssue) {
    return {
      title: "Le local doit etre corrige avant l'envoi",
      summary: "Certains documents doivent etre repares avant d'envoyer vers l'heberge.",
      steps: [
        "Lancer la reparation proposee.",
        "Relancer le controle local.",
        "Ne pas envoyer vers l'heberge si la reparation ne concerne que le local.",
      ],
      primary: { label: "Reparer les archives", kind: "repair", action: "repair_archive_not_openable" },
    };
  }
  if (supabaseDown) {
    return {
      title: "L'heberge n'est pas accessible",
      summary: "Le local est sauvegarde, mais la connexion a l'heberge est temporairement indisponible.",
      steps: [
        "Reessayer la connexion.",
        "Continuer en local sans envoyer.",
        "Envoyer plus tard quand l'heberge repond.",
      ],
      primary: { label: "Reessayer la connexion", kind: "repair", action: "repair_retry_supabase_check" },
    };
  }
  return {
    title: "Une action est necessaire",
    summary: "Le local doit etre verifie avant l'envoi vers l'heberge.",
    steps: [
      "Relancer le controle local.",
      "Corriger les points signales.",
      "Envoyer vers l'heberge uniquement si une vraie modification locale doit etre publiee.",
    ],
    primary: { label: "Reessayer", kind: "preflight" },
  };
}

async function runRescueRepairAction(action, buttonNode, statusNode, resultNode) {
  if (!action) return;
  const previous = buttonNode.textContent;
  buttonNode.disabled = true;
  buttonNode.textContent = "Action en cours...";
  if (statusNode) statusNode.textContent = "Action en cours...";
  try {
    const response = await fetch("/api/rescue/run-repair", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ action }),
    });
    const payload = await response.json().catch(() => ({}));
    const ok = Boolean(response.ok && payload?.ok && Number(payload?.exitCode) === 0);
    if (statusNode) statusNode.textContent = ok ? "Action terminee." : "L'action n'a pas pu etre terminee.";
    if (resultNode) {
      resultNode.innerHTML = ok
        ? "<strong>Action terminee.</strong> Relancez le controle."
        : `<strong>Action en erreur.</strong> ${escapeHtml(String(payload?.error || "Erreur inconnue"))}`;
    }
    if (ok) {
      await runRescuePreflight(statusNode, resultNode, { afterRepairSuccess: true });
    }
  } catch (error) {
    if (statusNode) statusNode.textContent = "L'action n'a pas pu etre terminee.";
    if (resultNode) {
      resultNode.textContent = `Impossible d'appeler l'API locale: ${String(error?.message || error || "")}`;
    }
  } finally {
    buttonNode.disabled = false;
    buttonNode.textContent = previous;
  }
}

function setRescueReloadAvailability(isAvailable) {
  const reloadBtn = document.getElementById("dotations-rescue-reload");
  if (!reloadBtn) return;
  reloadBtn.style.display = isAvailable ? "" : "none";
}

async function runRescuePreflight(statusNode, resultNode, { afterRepairSuccess = false } = {}) {
  if (statusNode) statusNode.textContent = "Action en cours...";
  setRescueReloadAvailability(false);
  try {
    const response = await fetch("/api/local/check-before-send", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
    });
    const payload = await response.json().catch(() => ({}));
    if (payload?.ok) {
      state.hostedSyncState = afterRepairSuccess ? "unknown" : "pending";
      state.hostedSyncDetails = afterRepairSuccess
        ? "Controle local OK apres reparation locale. Aucun envoi vers l'heberge n'est necessaire sauf modification metier locale volontaire."
        : "";
      if (statusNode) statusNode.textContent = afterRepairSuccess ? "Correction terminee. Le local est maintenant correct." : "Le controle est revenu OK.";
      if (resultNode) {
        resultNode.innerHTML = afterRepairSuccess
          ? "<strong>Correction terminee. Dotations peut etre recharge.</strong> Aucun envoi vers l'heberge n'est necessaire sauf modification metier locale volontaire."
          : "<strong>Le controle est revenu OK.</strong>";
      }
      const detailsText = document.getElementById("dotations-rescue-details-text");
      if (detailsText) {
        detailsText.textContent = state.hostedSyncDetails || "Controle local OK.";
      }
      setRescueReloadAvailability(true);
    } else {
      const localOk = Boolean(payload?.localOk);
      const pdfOk = Boolean(payload?.pdfIntegrityOk);
      const supabaseOk = Boolean(payload?.supabaseReachable);
      if (!supabaseOk) {
        state.hostedSyncState = "inaccessible";
      } else if (!localOk || !pdfOk) {
        state.hostedSyncState = "blocked";
      }
      state.hostedSyncDetails = JSON.stringify(payload?.technicalDetails || {}, null, 2);
      if (statusNode) statusNode.textContent = "Le controle est toujours bloque.";
      if (resultNode) resultNode.innerHTML = "<strong>Le controle est toujours bloque.</strong>";
      setRescueReloadAvailability(false);
    }
  } catch (error) {
    if (statusNode) statusNode.textContent = "L'action n'a pas pu etre terminee.";
    if (resultNode) resultNode.textContent = `Erreur reseau: ${String(error?.message || error || "")}`;
    setRescueReloadAvailability(false);
  } finally {
    renderHostedSyncUi();
  }
}

function openRescueActionModal(options = {}) {
  const rawDetailsOverride = options && Object.prototype.hasOwnProperty.call(options, "detailsRaw")
    ? String(options.detailsRaw || "")
    : null;
  const stateOverride = options && Object.prototype.hasOwnProperty.call(options, "hostedState")
    ? String(options.hostedState || "")
    : null;
  const modal = ensureRescueModalUi();
  const plan = buildRescuePlanFromState(rawDetailsOverride, stateOverride);
  const titleNode = document.getElementById("dotations-rescue-title");
  const summaryNode = document.getElementById("dotations-rescue-summary");
  const stepsNode = document.getElementById("dotations-rescue-steps");
  const primaryButton = document.getElementById("dotations-rescue-primary");
  const primaryStatusNode = document.getElementById("dotations-rescue-primary-status");
  const secondaryWrap = document.getElementById("dotations-rescue-secondary-actions");
  const detailsText = document.getElementById("dotations-rescue-details-text");
  const resultNode = document.getElementById("dotations-rescue-run-result");

  if (titleNode) titleNode.textContent = plan.title;
  if (summaryNode) summaryNode.textContent = plan.summary;
  if (stepsNode) {
    stepsNode.innerHTML = (plan.steps || []).map((step) => `<li>${escapeHtml(String(step || ""))}</li>`).join("");
  }
  if (detailsText) {
    detailsText.textContent = String(rawDetailsOverride == null ? (state.hostedSyncDetails || "Aucun detail technique.") : rawDetailsOverride || "Aucun detail technique.");
  }
  if (resultNode) resultNode.textContent = "";
  if (primaryStatusNode) primaryStatusNode.textContent = "";
  setRescueReloadAvailability(false);

  if (primaryButton) {
    primaryButton.textContent = plan.primary?.label || "Reessayer";
    primaryButton.onclick = async () => {
      if (plan.primary?.kind === "repair") {
        await runRescueRepairAction(plan.primary.action, primaryButton, primaryStatusNode, resultNode);
      } else {
        await runRescuePreflight(primaryStatusNode, resultNode);
      }
    };
  }

  if (secondaryWrap) {
    const secondaryActions = [
      { label: "Reessayer le controle", kind: "preflight" },
      { label: "Reessayer la connexion", kind: "repair", action: "repair_retry_supabase_check" },
      { label: "Mettre en quarantaine", kind: "repair", action: "repair_orphans_quarantine" },
      { label: "Restaurer les PDF", kind: "repair", action: "repair_missing_pdf" },
    ];
    secondaryWrap.innerHTML = "";
    secondaryActions.forEach((item) => {
      const btn = document.createElement("button");
      btn.textContent = item.label;
      btn.className = "button button--secondary";
      btn.style.cssText = "height:30px;padding:0 12px;border-radius:10px;font-size:12px;opacity:.95;";
      btn.type = "button";
      btn.addEventListener("click", async () => {
        if (item.kind === "repair") {
          await runRescueRepairAction(item.action, btn, primaryStatusNode, resultNode);
        } else {
          await runRescuePreflight(primaryStatusNode, resultNode);
        }
      });
      secondaryWrap.appendChild(btn);
    });
  }
  modal.style.display = "flex";
}

function renderHostedSyncUi() {
  if (getDataBackendMode() !== "LOCAL_API") return;
  ensureHostedSyncUi();
  const localNode = document.getElementById("dotations-sync-local");
  const remoteNode = document.getElementById("dotations-sync-remote");
  const pushBtn = document.getElementById("dotations-sync-push-btn");
  const helpBtn = document.getElementById("dotations-sync-help-btn");
  const detailsNode = document.getElementById("dotations-sync-details");
  const detailsTextNode = document.getElementById("dotations-sync-details-text");
  if (localNode) {
    localNode.textContent = state.isDirty ? "Local : non sauvegarde" : "Local : sauvegarde";
  }
  if (remoteNode) {
    const isChecking = Boolean(state.hostedSyncInFlight);
    const isUpToDate = !isChecking && String(state.hostedSyncState || "") === "up_to_date";
    remoteNode.textContent = `Heberge : ${getHostedSyncSimpleText()}`;
    remoteNode.classList.toggle("is-checking", isChecking);
    remoteNode.classList.toggle("is-ok", isUpToDate);
    remoteNode.classList.toggle("is-alert", !isChecking && !isUpToDate);
    remoteNode.setAttribute("aria-busy", isChecking ? "true" : "false");
  }
  if (pushBtn) {
    pushBtn.disabled = Boolean(state.hostedSyncInFlight);
  }
  if (helpBtn) {
    helpBtn.style.display = state.hostedSyncState === "blocked" ? "" : "none";
  }
  if (detailsNode && detailsTextNode) {
    const details = String(state.hostedSyncDetails || "").trim();
    detailsNode.style.display = details ? "" : "none";
    detailsTextNode.textContent = details;
  }
}

function maybeOpenRescueModalPreview() {
  try {
    const params = new URLSearchParams(window.location.search || "");
    if (params.get("demoRescueModal") !== "1") return;
    const demoDetails = JSON.stringify(
      {
        missingActiveFiles: 2,
        errors: ["ACTIVE_FILE_MISSING[DOC-EXEMPLE-1]", "ACTIVE_FILE_MISSING[DOC-EXEMPLE-2]"],
        supabaseReachable: true,
      },
      null,
      2
    );
    ensureHostedSyncUi();
    renderHostedSyncUi();
    openRescueActionModal({ detailsRaw: demoDetails, hostedState: "blocked" });
  } catch (error) {
    // ignore preview errors
  }
}

function ensureCloseDotationsModalUi() {
  let modal = document.getElementById("dotations-close-modal");
  if (modal) return modal;
  modal = document.createElement("div");
  modal.id = "dotations-close-modal";
  modal.style.cssText = "position:fixed;inset:0;background:rgba(15,23,42,.45);display:none;align-items:center;justify-content:center;z-index:10000;padding:16px;";
  modal.innerHTML = [
    '<div role="dialog" aria-modal="true" aria-labelledby="dotations-close-title" style="width:min(640px,100%);background:#f8fbff;border:1px solid #dbe6f5;border-radius:14px;box-shadow:0 12px 36px rgba(15,23,42,.18);padding:18px;">',
    '<h3 id="dotations-close-title" style="margin:0 0 8px 0;color:#0f172a;">Confirmation de fermeture</h3>',
    '<p id="dotations-close-summary" style="margin:0 0 10px 0;color:#334155;"></p>',
    '<div id="dotations-close-actions" style="display:flex;gap:8px;flex-wrap:wrap;"></div>',
    '<details style="margin-top:10px;"><summary>Details techniques</summary><pre id="dotations-close-details" style="white-space:pre-wrap;background:#fff;border:1px solid #e2e8f0;border-radius:8px;padding:8px;"></pre></details>',
    '<p id="dotations-close-result" style="margin:10px 0 0 0;color:#334155;font-size:12px;"></p>',
    "</div>",
  ].join("");
  modal.addEventListener("click", (event) => {
    if (event.target === modal) modal.style.display = "none";
  });
  document.body.appendChild(modal);
  return modal;
}

function isHostedNotSentState() {
  const s = String(state.hostedSyncState || "");
  return s === "pending" || s === "blocked" || s === "inaccessible";
}

function isLocalHostedSyncDisabled() {
  if (getDataBackendMode() !== "LOCAL_API") {
    return true;
  }
  const hostedState = String(state.hostedSyncState || "");
  return Boolean(state.hostedSyncInFlight || hostedState === "blocked" || hostedState === "inaccessible");
}

function getCloseModalScenario() {
  const hostedState = String(state.hostedSyncState || "unknown");
  if (state.isDirty) {
    return {
      key: "UNSAVED",
      summary: "Des modifications ne sont pas encore sauvegardees.",
      actions: [
        { label: "Retourner sauvegarder", kind: "closeModal", primary: true },
        { label: "Fermer sans sauvegarder", kind: "shutdown" },
      ],
    };
  }
  if (String(state.hostedSyncState || "") === "pending") {
    return {
      key: "LOCAL_ONLY",
      summary: "Les donnees sont sauvegardees en local, mais l'heberge n'est pas encore mis a jour.",
      actions: [
        { label: "Rester", kind: "closeModal", primary: true },
        { label: "Envoyer vers l'heberge", kind: "sendHosted" },
        { label: "Fermer quand meme", kind: "shutdown" },
      ],
    };
  }
  if (String(state.hostedSyncState || "") === "inaccessible") {
    return {
      key: "HOSTED_DOWN",
      summary: "L'heberge n'est pas accessible pour l'instant. Vos donnees restent sauvegardees en local.",
      actions: [
        { label: "Rester", kind: "closeModal", primary: true },
        { label: "Fermer Dotations", kind: "shutdown" },
      ],
    };
  }
  if (hostedState === "unknown") {
    return {
      key: "SAFE_UNKNOWN",
      summary: "Dotations peut etre ferme. Aucune modification locale non sauvegardee n'a ete detectee. L'etat heberge n'a pas ete verifie.",
      actions: [
        { label: "Non, rester", kind: "closeModal", primary: true },
        { label: "Oui, fermer Dotations", kind: "shutdown" },
      ],
    };
  }
  return {
    key: "SAFE",
    summary: "Dotations peut etre ferme.",
    actions: [
      { label: "Non, rester", kind: "closeModal", primary: true },
      { label: "Oui, fermer Dotations", kind: "shutdown" },
    ],
  };
}

async function shutdownDotationsFromUi(resultNode, triggerBtn) {
  if (triggerBtn) triggerBtn.disabled = true;
  if (resultNode) resultNode.textContent = "Fermeture en cours...";
  try {
    const response = await fetch("/api/local/shutdown", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
    });
    const payload = await response.json().catch(() => ({}));
    if (response.ok && payload?.ok) {
      if (resultNode) resultNode.textContent = "Dotations est ferme. Vous pouvez fermer cet onglet.";
      try { window.close(); } catch {}
      window.setTimeout(() => {
        try {
          document.body.innerHTML = "<main style='font-family:Arial,sans-serif;padding:24px'><h2>Dotations est ferme.</h2><p>Vous pouvez fermer cet onglet.</p></main>";
        } catch {}
      }, 300);
      return;
    }
    if (resultNode) resultNode.textContent = "La fermeture automatique a echoue. Vous pouvez fermer cet onglet manuellement.";
  } catch (error) {
    if (resultNode) resultNode.textContent = "La fermeture automatique a echoue. Vous pouvez fermer cet onglet manuellement.";
  } finally {
    if (triggerBtn) triggerBtn.disabled = false;
  }
}

function openCloseDotationsModal() {
  const modal = ensureCloseDotationsModalUi();
  const summaryNode = document.getElementById("dotations-close-summary");
  const actionsNode = document.getElementById("dotations-close-actions");
  const detailsNode = document.getElementById("dotations-close-details");
  const detailsWrapNode = detailsNode ? detailsNode.closest("details") : null;
  const resultNode = document.getElementById("dotations-close-result");
  const scenario = getCloseModalScenario();
  if (summaryNode) summaryNode.textContent = scenario.summary;
  if (detailsNode) {
    detailsNode.textContent = JSON.stringify({
      isDirty: Boolean(state.isDirty),
      hostedSyncState: String(state.hostedSyncState || "unknown"),
      hostedSyncDetails: String(state.hostedSyncDetails || ""),
    }, null, 2);
  }
  if (detailsWrapNode) {
    detailsWrapNode.open = false;
  }
  if (resultNode) resultNode.textContent = "";
  if (actionsNode) {
    actionsNode.innerHTML = "";
    scenario.actions.forEach((action) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = action.primary ? "button button--primary" : "button button--secondary";
      btn.style.cssText = "height:32px;padding:0 12px;border-radius:10px;";
      btn.textContent = action.label;
      btn.addEventListener("click", async () => {
        if (action.kind === "closeModal") {
          modal.style.display = "none";
          return;
        }
        if (action.kind === "sendHosted") {
          modal.style.display = "none";
          await runHostedSyncPushFromUi();
          return;
        }
        if (action.kind === "shutdown") {
          await shutdownDotationsFromUi(resultNode, btn);
        }
      });
      actionsNode.appendChild(btn);
    });
  }
  modal.style.display = "flex";
}

function bindBeforeUnloadGuard() {
  if (window.__dotationsBeforeUnloadBound) return;
  window.__dotationsBeforeUnloadBound = true;
  window.addEventListener("beforeunload", (event) => {
    if (!state.isDirty) return;
    event.preventDefault();
    event.returnValue = "";
  });
}

async function runHostedSyncPushFromUi(options = {}) {
  const {
    confirmBeforeSend = true,
    silent = false,
  } = options || {};
  if (state.hostedSyncInFlight) return false;
  state.hostedSyncInFlight = true;
  renderHostedSyncUi();
  try {
    const preflightResponse = await fetch("/api/local/check-before-send", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
    });
    const preflight = await preflightResponse.json().catch(() => ({}));
    if (!preflightResponse.ok || !preflight?.ok) {
      const localOk = Boolean(preflight?.localOk);
      const pdfOk = Boolean(preflight?.pdfIntegrityOk);
      const supabaseOk = Boolean(preflight?.supabaseReachable);
      if (!supabaseOk) {
        state.hostedSyncState = "inaccessible";
        if (!silent) showDataStatus("L'heberge n'est pas accessible pour l'instant. Vos modifications restent sauvegardees en local. Reessayez plus tard.");
      } else if (!localOk || !pdfOk) {
        state.hostedSyncState = "blocked";
        if (!silent) showDataStatus("L'envoi est bloque. Le local doit d'abord etre corrige.");
      } else {
        state.hostedSyncState = "blocked";
        if (!silent) showDataStatus("Controle local necessaire avant envoi.");
      }
      state.hostedSyncDetails = JSON.stringify(preflight?.technicalDetails || {}, null, 2);
      renderHostedSyncUi();
      return false;
    }
    if (confirmBeforeSend) {
      const confirmed = window.confirm("Cette action va envoyer les donnees locales vers l'heberge.\nA utiliser uniquement si les donnees locales sont correctes.\nContinuer ?");
      if (!confirmed) {
        renderHostedSyncUi();
        return false;
      }
    }
    const pushResponse = await fetch("/api/sync/send-to-hosted", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
    });
    const payload = await pushResponse.json().catch(() => ({}));
    if (pushResponse.ok && payload?.ok) {
      state.hostedSyncState = "up_to_date";
      showDataStatus(silent ? "SAUVEGARDE LOCALE ET HEBERGEE OK" : "Heberge mis a jour.");
      state.hostedSyncDetails = "";
      renderHostedSyncUi();
      return true;
    }
    if (payload?.blocked) {
      const localMsg = String(payload?.userMessage || "");
      if (/heberge/i.test(localMsg)) {
        state.hostedSyncState = "inaccessible";
      } else {
        state.hostedSyncState = "blocked";
      }
      if (!silent) showDataStatus(localMsg || "L'envoi est bloque. Le local doit d'abord etre corrige.");
      state.hostedSyncDetails = JSON.stringify(payload?.technicalDetails || {}, null, 2);
      renderHostedSyncUi();
      return false;
    }
    if (payload?.error === "PUSH_BLOCKED_PDF_INTEGRITY_FAILED") {
      state.hostedSyncState = "blocked";
      if (!silent) showDataStatus("L'envoi est bloque. Le local doit d'abord etre corrige.");
      state.hostedSyncDetails = JSON.stringify(payload || {}, null, 2);
      renderHostedSyncUi();
      return false;
    }
    state.hostedSyncState = "inaccessible";
    if (!silent) showDataStatus("L'envoi vers l'heberge a echoue. Vos donnees restent sauvegardees en local.");
    state.hostedSyncDetails = JSON.stringify(payload || {}, null, 2);
    renderHostedSyncUi();
    return false;
  } catch (error) {
    state.hostedSyncState = "inaccessible";
    if (!silent) showDataStatus("L'heberge n'est pas accessible pour l'instant. Reessayez plus tard.");
    state.hostedSyncDetails = String(error?.message || error || "");
    renderHostedSyncUi();
    return false;
  } finally {
    state.hostedSyncInFlight = false;
    renderHostedSyncUi();
  }
}

function showDataStatus(text) {
  const node = document.getElementById("data-status");
  if (node) {
    node.textContent = text;
    node.classList.remove(
      "data-status--create",
      "data-status--update",
      "data-status--delete",
      "data-status--warning",
      "data-status--neutral"
    );
    node.classList.add("data-status--neutral");
  }
}

function showActionStatus(type, text) {
  const node = document.getElementById("data-status");
  if (!node) {
    return;
  }

  if (state.statusTimerId) {
    window.clearTimeout(state.statusTimerId);
    state.statusTimerId = 0;
  }

  node.textContent = text;
  node.classList.remove(
    "data-status--create",
    "data-status--update",
    "data-status--delete",
    "data-status--warning",
    "data-status--neutral"
  );
  node.classList.add(`data-status--${type}`);

  state.statusTimerId = window.setTimeout(() => {
    showDataStatus(text);
    state.statusTimerId = 0;
  }, 3200);

  const isMobileSignaturePage = isMobileSignaturePageContext();
  const shouldAutoSaveAction =
    (type === "create" || type === "update" || type === "delete") && !isMobileSignaturePage;
  if (shouldAutoSaveAction && state.isDirty && state.data) {
    if (state.actionAutoSaveTimerId) {
      window.clearTimeout(state.actionAutoSaveTimerId);
      state.actionAutoSaveTimerId = 0;
    }
    state.actionAutoSaveTimerId = window.setTimeout(() => {
      state.actionAutoSaveTimerId = 0;
      saveDataToFile({
        silent: true,
        reloadAfter: true,
        promptDownload: false,
        successText: "SAUVEGARDE AUTOMATIQUE",
      }).catch((error) => {
        console.error(error);
      });
    }, 180);
  }
}

function loadWorkingData() {
  try {
    const raw = window.sessionStorage.getItem(WORKING_DATA_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch (error) {
    console.error(error);
    return null;
  }
}

function cloneData(value) {
  return JSON.parse(JSON.stringify(value));
}

function pushUndoSnapshot(label) {
  if (!state.data) {
    return;
  }
  state.undoStack.push({
    label,
    data: cloneData(state.data),
  });
  if (state.undoStack.length > MAX_UNDO_STACK) {
    state.undoStack.shift();
  }
}

function clearUndoStack() {
  state.undoStack = [];
}

function undoLastChange() {
  const lastSnapshot = state.undoStack.pop();
  if (!lastSnapshot) {
    showDataStatus("AUCUNE ANNULATION DISPONIBLE");
    return;
  }

  state.data = cloneData(lastSnapshot.data);
  state.editingEffectId = "";
  state.editingReferenceId = "";
  state.editingReplacementCostKey = "";
  state.editingSimpleReference = null;
  migrateDataModel();
  state.isDirty = true;
  saveWorkingData();
  schedulePageRender();
  showDataStatus(`ANNULATION : ${lastSnapshot.label}`);
}

function saveWorkingData() {
  if (!state.data) {
    return;
  }
  try {
    window.sessionStorage.setItem(WORKING_DATA_KEY, JSON.stringify(state.data));
  } catch (error) {
    console.error(error);
  }
}

function clearWorkingData() {
  try {
    window.sessionStorage.removeItem(WORKING_DATA_KEY);
  } catch (error) {
    console.error(error);
  }
}

function bindAutoSaveOnNavigation() {
  if (state.autoSaveNavigationBound) {
    return;
  }

  const preservePersonForInternalNav = (nextUrl) => {
    if (!(nextUrl instanceof URL)) {
      return;
    }

    const targetFile = nextUrl.pathname.split("/").pop() || "";
    const needsPerson = ["fiche-personne.html", "document-arrivee.html", "document-sortie.html"].includes(targetFile);
    if (!needsPerson || nextUrl.searchParams.get("personId")) {
      return;
    }

    const currentPersonId = getCurrentPersonId();
    if (currentPersonId) {
      nextUrl.searchParams.set("personId", currentPersonId);
    }
  };

  document.addEventListener(
    "click",
    (event) => {
      const target = event.target;
      if (!(target instanceof Element)) {
        return;
      }
      const anchor = target.closest("a[href]");
      if (!(anchor instanceof HTMLAnchorElement)) {
        return;
      }
      if (anchor.target === "_blank" || anchor.hasAttribute("download")) {
        return;
      }
      const rawHref = String(anchor.getAttribute("href") || "").trim();
      if (!rawHref || rawHref.startsWith("#") || rawHref.toLowerCase().startsWith("javascript:")) {
        return;
      }
      const nextUrl = new URL(anchor.href, window.location.href);
      const currentUrl = new URL(window.location.href);
      const samePage =
        nextUrl.origin === currentUrl.origin &&
        nextUrl.pathname === currentUrl.pathname &&
        nextUrl.search === currentUrl.search;
      if (samePage) {
        return;
      }
      preservePersonForInternalNav(nextUrl);
      capturePendingEditsBeforeNavigation();
      if (!state.isDirty) {
        return;
      }
      event.preventDefault();
      navigateWithAutoSave(nextUrl.href);
    },
    true
  );

  state.autoSaveNavigationBound = true;
}

function bindGlobalShortcuts() {
  if (state.shortcutsBound) {
    return;
  }

  document.addEventListener("keydown", (event) => {
    if (!(event.ctrlKey || event.metaKey) || event.key.toLowerCase() !== "z") {
      return;
    }

    const target = event.target;
    if (
      target instanceof HTMLElement &&
      (target.tagName === "TEXTAREA" ||
        target.isContentEditable ||
        (target.tagName === "INPUT" &&
          !["checkbox", "radio", "button", "submit"].includes(
            String(target.getAttribute("type") || "").toLowerCase()
          )))
    ) {
      return;
    }

    event.preventDefault();
    undoLastChange();
  });

  state.shortcutsBound = true;
}

function bindHistoryNavigation() {
  if (window.__dashboardHistoryBound) {
    return;
  }

  window.addEventListener("popstate", () => {
    applyActiveNav();
    if (!state.data) {
      return;
    }
    schedulePageRender();
  });

  window.__dashboardHistoryBound = true;
}

function isPersonFilterResetButton(button) {
  if (!(button instanceof HTMLButtonElement)) {
    return false;
  }
  const form = button.closest("form");
  return form instanceof HTMLFormElement && form.matches(".js-filter-form, #documents-archives-filter-form");
}

function bindGlobalResetSelectionClear() {
  if (window.__dashboardResetSelectionBound) {
    return;
  }
  document.addEventListener(
    "click",
    (event) => {
      const target = event.target;
      if (!(target instanceof Element)) return;
      const button = target.closest("button");
      if (!(button instanceof HTMLButtonElement)) return;
      const label = normalizeText(button.textContent || "");
      if (!label.startsWith("REINITIALISER")) return;
      if (!isPersonFilterResetButton(button)) return;
      setCurrentPersonId("", "replace");
    },
    true
  );
  window.__dashboardResetSelectionBound = true;
}

function isNeutralFilterValue(value) {
  const normalized = normalizeText(value);
  return !normalized || normalized === "TOUS" || normalized === "TOUTES" || normalized === "ALL";
}

function isFilterControlActive(control) {
  if (!(control instanceof HTMLElement) || control.disabled) {
    return false;
  }
  if (control instanceof HTMLInputElement) {
    const type = String(control.type || "text").toLowerCase();
    if (["button", "submit", "reset", "hidden"].includes(type)) {
      return false;
    }
    if (type === "checkbox" || type === "radio") {
      return control.checked !== control.defaultChecked;
    }
    return !isNeutralFilterValue(control.value);
  }
  if (control instanceof HTMLSelectElement || control instanceof HTMLTextAreaElement) {
    return !isNeutralFilterValue(control.value);
  }
  return false;
}

function hasActiveFilterControls(form) {
  if (!(form instanceof HTMLFormElement)) {
    return false;
  }
  return Array.from(form.elements).some((control) => isFilterControlActive(control));
}

function getResetButtonsForForm(form) {
  if (!(form instanceof HTMLFormElement)) {
    return [];
  }
  return Array.from(form.querySelectorAll("button, input[type='reset']"))
    .filter((button) => {
      const type = String(button.getAttribute("type") || "").toLowerCase();
      if (type === "reset") return true;
      return normalizeText(button.textContent || button.value || "").startsWith("REINITIALISER");
    });
}

function updateFilterResetHighlight(form) {
  if (!(form instanceof HTMLFormElement)) {
    return;
  }
  const active = hasActiveFilterControls(form) || hasActiveTableSortForCurrentPage();
  getResetButtonsForForm(form).forEach((button) => {
    button.classList.toggle("filter-reset--active", active);
    button.classList.toggle("is-active", active);
    button.setAttribute("aria-pressed", active ? "true" : "false");
  });
}

function updateAllFilterResetHighlights() {
  document.querySelectorAll("form").forEach((form) => updateFilterResetHighlight(form));
}

function bindFilterResetHighlights() {
  if (window.__dashboardFilterResetHighlightsBound) {
    return;
  }
  const scheduleUpdate = () => window.setTimeout(updateAllFilterResetHighlights, 0);
  document.addEventListener("input", scheduleUpdate, true);
  document.addEventListener("change", scheduleUpdate, true);
  document.addEventListener("reset", scheduleUpdate, true);
  window.__dashboardFilterResetHighlightsBound = true;
  updateAllFilterResetHighlights();
}


function markDirty() {
  state.isDirty = true;
  state.saveButtonLatchedDirty = true;
  state.localMutationTick = Number(state.localMutationTick || 0) + 1;
  saveWorkingData();
  renderDirtyState();
  scheduleBackgroundAutoSave();
}

function renderDirtyState() {
  const node = document.getElementById("dirty-status");
  const saveButtons = document.querySelectorAll(".js-save-data");
  const saveButtonSignature = Array.from(saveButtons)
    .map((button) => `${button.className}`)
    .join("|");
  const nextDirtyStateSignature = [
    String(state.isDirty ? "1" : "0"),
    String(state.saveButtonLatchedDirty ? "1" : "0"),
    String(saveButtons.length),
    String(saveButtonSignature),
  ].join("|");
  if (state.dirtyStateRenderSignature === nextDirtyStateSignature) {
    return;
  }
  state.dirtyStateRenderSignature = nextDirtyStateSignature;

  if (node) {
    node.hidden = false;
    node.textContent = state.isDirty ? "MODIFICATIONS NON SAUVEGARDEES" : "DONNEES SAUVEGARDEES";
    node.classList.toggle("is-saved", !state.isDirty);
  }

  const saveButtonActive = Boolean(state.isDirty || state.saveButtonLatchedDirty);
  document.querySelectorAll(".js-save-data").forEach((button) => {
    if (!(button instanceof HTMLElement)) {
      return;
    }
    button.classList.toggle("button--primary", saveButtonActive);
    button.classList.toggle("button--secondary", !saveButtonActive);
  });
  renderHostedSyncUi();
}

window.dotationsGetNetworkDebug = getNetworkDebugReport;
window.dotationsResetNetworkDebug = () => {
  state.networkDebug = {
    samples: [],
    routeStats: {},
    requestCount: 0,
    totalBytes: 0,
    startedAt: Date.now(),
  };
};

window.getNetworkDebug = () => getNetworkDebugReport();
window.resetNetworkDebug = () => {
  state.networkDebug = {
    samples: [],
    routeStats: {},
    requestCount: 0,
    totalBytes: 0,
    startedAt: Date.now(),
  };
};

loadData();




