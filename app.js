const DEFAULT_FILTERS = {
  search: "",
  site: "",
  typePersonnel: "",
  typeContrat: "",
  statutDossier: "",
  statutObjet: "",
  typeEffet: "",
};

const state = {
  data: null,
  supabaseRevision: null,
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
  mobileSignatureNetworkInfo: null,
  autoSaveNavigationBound: false,
  searchClearBrowserEventsBound: false,
  dirtyFallbackBound: false,
  saveButtonLatchedDirty: false,
  autoSaveInFlightPromise: null,
  autoSaveTimerId: 0,
  tableSorts: {
    sheetEffects: { key: "typeEffet", dir: "asc" },
    arrivalEffects: { key: "typeEffet", dir: "asc" },
    exitEffects: { key: "typeEffet", dir: "asc" },
    overviewPersons: { key: "nom", dir: "asc" },
  },
  filters: { ...DEFAULT_FILTERS },
  referenceRenderContext: null,
  urgentMode: false,
  lastSaveInfo: null,
  effectRowFlash: null,
  effectTableFlash: null,
  autoPdfGenerationInFlight: false,
  autoPdfGeneratedKeys: new Set(),
  signedDocumentsPopupSeenKeys: new Set(),
  previousSignatureValidationMap: new Map(),
};

const WORKING_DATA_KEY = "dashboard-working-data";
const LEGACY_CONTRACT_TYPES = ["CDI", "CDD", "INTERIMAIRE"];
const MAX_UNDO_STACK = 30;
const ALL_SITES_VALUE = "TOUS SITES";
const EFFECT_STATUS_CAUSES = ["HS", "PERTE", "VOL", "NON RENDU", "DETRUIT"];
const BILLABLE_EFFECT_CAUSES = ["PERTE", "VOL", "NON RENDU", "DETRUIT"];
const NON_RENDU_REFERENCE_COSTS = {
  "BADGE INTRUSION": 15,
  "CARTE TURBOSELF": 10,
  CLE: 5,
  "CLE CES": 50,
  "TELECOMMANDE URMET": 40,
};
const MOBILE_SIGNATURE_REQUEST_TTL_MS = 10 * 60 * 1000;
const SUPABASE_PROJECT_URL = "https://dphrvdhqhgycmllietuk.supabase.co";
const SUPABASE_PUBLISHABLE_KEY = "sb_publishable_2wYXnIDj4-c8daQZW8D5hA_2Py6k7z6";
const SUPABASE_APP_STATE_TABLE = "app_state";
const SUPABASE_APP_STATE_ID = "main";
const DEFAULT_SUPABASE_PDF_BUCKET = "pdf";
const DEFAULT_SUPABASE_SIGNATURES_BUCKET = "signatures";
const NAVIGATION_CONTEXT_KEY = "dotations-navigation-context";
const SIGNED_POPUP_SEEN_STORAGE_KEY = "dotations-signed-popup-seen";
const PENDING_PDF_REMINDER_SNOOZE_KEY = "dotations-pending-pdf-reminder-snooze";
const PENDING_PDF_TASK_STORAGE_KEY = "dotations-pending-pdf-task";
const PDF_LAYOUT_VERSION = "2026-03-14-exit-layout-fix-3";
const PDF_FORMAT_LOCK = "v1";
let pdfModalCleanupBound = false;
let reminderSnoozeMap = {};
const signatureCanvases = new WeakMap();

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
  if (normalized === "CASSE") return "HS";
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
  if (normalized === "CASSE") return "HS";
  return normalized;
}

function getFallbackNonRenduCost(typeEffet, designation = "") {
  const normalizedType = normalizePricingKey(typeEffet);
  if (!normalizedType) return 0;
  if (normalizedType === "CLE") {
    return isCesKeyDesignation(designation) ? 50 : 5;
  }
  return NON_RENDU_REFERENCE_COSTS[normalizedType] || 0;
}

function getCauseFromManualStatus(manualStatus) {
  const normalized = normalizeText(manualStatus);
  if (normalized === "PERDU") return "PERTE";
  if (normalized === "DETRUIT") return "DETRUIT";
  if (normalized === "VOL") return "VOL";
  if (normalized === "HS" || normalized === "CASSE") return "HS";
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
  return (
    window.location.protocol === "file:" ||
    host === "localhost" ||
    host === "127.0.0.1"
  );
}

function isSupabaseConfigured() {
  const url = normalizeHttpUrl(SUPABASE_PROJECT_URL);
  const key = String(SUPABASE_PUBLISHABLE_KEY || "").trim();
  return Boolean(url && key && key.startsWith("sb_"));
}

function getDataBackendMode() {
  if (isSupabaseConfigured()) {
    return "SUPABASE";
  }
  if (isLocalRuntime()) {
    return "LOCAL_API";
  }
  return "HOSTED_NO_BACKEND";
}

function getSupabaseRestEndpoint() {
  const baseUrl = normalizeHttpUrl(SUPABASE_PROJECT_URL);
  return `${baseUrl}/rest/v1/${SUPABASE_APP_STATE_TABLE}`;
}

function getSupabaseHeaders(extra = {}, options = {}) {
  const key = String(SUPABASE_PUBLISHABLE_KEY || "").trim();
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
    headers.Authorization = `Bearer ${key}`;
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

async function fetchSupabaseStateData() {
  const endpoint = `${getSupabaseRestEndpoint()}?id=eq.${encodeURIComponent(
    SUPABASE_APP_STATE_ID
  )}&select=payload,revision&limit=1`;
  const response = await fetch(endpoint, {
    headers: getSupabaseHeaders(),
    cache: "no-store",
  });
  if (!response.ok) {
    throw new Error("SUPABASE LOAD FAILED");
  }
  const rows = await response.json();
  if (!Array.isArray(rows) || !rows.length || !rows[0]?.payload) {
    throw new Error("SUPABASE EMPTY PAYLOAD");
  }
  const revision = Number(rows[0]?.revision);
  state.supabaseRevision = Number.isFinite(revision) ? revision : 0;
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

async function saveSupabaseStateData(payload) {
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

async function fetchLatestDataSnapshot() {
  const mode = getDataBackendMode();
  if (mode === "SUPABASE") {
    return fetchSupabaseStateData();
  }
  const response = await fetch(`/api/data?ts=${Date.now()}`, { cache: "no-store" });
  if (!response.ok) {
    throw new Error("LOCAL LOAD FAILED");
  }
  return response.json();
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
  return ["CLE", "CLE CES"].includes(normalizeText(typeEffet));
}

function getReferenceCatalogType(typeEffet) {
  return typeUsesReferenceCatalog(typeEffet) ? "CLE" : normalizeText(typeEffet);
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
    dateEntree: String(formData.get("sheetDateEntree") || ""),
    dateSortiePrevue: String(formData.get("sheetDateSortiePrevue") || ""),
    dateSortieReelle: String(formData.get("sheetDateSortieReelle") || ""),
  };

  const changed =
    person.nom !== draft.nom ||
    person.prenom !== draft.prenom ||
    normalizeText(person.fonction) !== draft.fonction ||
    !haveSameSites(getPersonSites(person), draft.sitesAffectation) ||
    normalizeText(person.typePersonnel) !== draft.typePersonnel ||
    normalizeText(person.typeContrat) !== draft.typeContrat ||
    String(person.dateEntree || "") !== draft.dateEntree ||
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
  const maxWaitMs = 220;
  let navigationDone = false;
  const navigateNow = () => {
    if (navigationDone) {
      return;
    }
    navigationDone = true;
    performPageNavigation(url, mode);
  };
  const fallbackTimer = window.setTimeout(() => {
    navigateNow();
  }, maxWaitMs);
  runAutoSaveBeforeNavigation().finally(() => {
    if (navigationDone) {
      return;
    }
    window.clearTimeout(fallbackTimer);
    navigateNow();
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
  setCurrentPersonId(normalizedPersonId, "replace");
  navigateWithAutoSave(
    `fiche-personne.html?personId=${encodeURIComponent(normalizedPersonId)}&editEffectId=${encodeURIComponent(normalizedEffectId)}`
  );
}

function consumeRequestedEditEffectId() {
  const nextUrl = new URL(window.location.href);
  const requestedId = String(nextUrl.searchParams.get("editEffectId") || "");
  if (!requestedId) {
    return "";
  }
  nextUrl.searchParams.delete("editEffectId");
  window.history.replaceState({}, "", nextUrl);
  return requestedId;
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

function getActiveMobileSignatureRequest(personId, docType, signer = "personnel") {
  cleanupExpiredMobileSignatureRequests();
  const normalizedSigner = normalizeMobileSignatureSigner(signer);
  return (state.data?.demandesSignatureMobile || []).find(
    (entry) =>
      entry.personId === personId &&
      entry.docType === normalizeText(docType) &&
      normalizeMobileSignatureSigner(entry.signer) === normalizedSigner &&
      entry.status === "EN ATTENTE" &&
      Date.parse(entry.expiresAt || "") > Date.now()
  ) || null;
}

function createMobileSignatureRequest(personId, docType, signer = "personnel") {
  const createdAt = new Date();
  const expiresAt = new Date(createdAt.getTime() + MOBILE_SIGNATURE_REQUEST_TTL_MS);
  const normalizedSigner = normalizeMobileSignatureSigner(signer);
  const request = {
    id: getNextId("DSM", state.data?.demandesSignatureMobile || []),
    token: generateMobileSignatureToken(),
    personId: String(personId || ""),
    docType: normalizeText(docType),
    signer: normalizedSigner.toUpperCase(),
    createdAt: createdAt.toISOString(),
    expiresAt: expiresAt.toISOString(),
    status: "EN ATTENTE",
    validatedAt: "",
  };
  state.data.demandesSignatureMobile.push(request);
  return request;
}

function getMobileSignaturePageUrl(request) {
  const docType = normalizeText(request?.docType) === "EXIT" ? "exit" : "arrival";
  const signer = normalizeMobileSignatureSigner(request?.signer || "");
  return `signature-mobile.html?personId=${encodeURIComponent(request?.personId || "")}&docType=${encodeURIComponent(docType)}&token=${encodeURIComponent(request?.token || "")}&signer=${encodeURIComponent(signer)}`;
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
  const encoded = encodeURIComponent(absoluteUrl);
  return [
    `https://api.qrserver.com/v1/create-qr-code/?size=180x180&margin=0&data=${encoded}`,
    `https://quickchart.io/qr?size=180&margin=0&text=${encoded}`,
    `https://chart.googleapis.com/chart?chs=180x180&cht=qr&chl=${encoded}`,
  ];
}

async function getAbsoluteMobileSignatureUrl(request) {
  if (!request) {
    return "";
  }
  const relativeUrl = getMobileSignaturePageUrl(request);
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

  const absoluteUrl = await getAbsoluteMobileSignatureUrl(request);

  wrapper.hidden = false;
  input.value = absoluteUrl;
  copyButton.disabled = false;
  if (reachabilityHintNode) {
    reachabilityHintNode.textContent = getMobileSignatureReachabilityHint(absoluteUrl);
  }
  if (qrWrapper && qrImage) {
    const providerUrls = getQrProviderUrls(absoluteUrl);
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
        reachabilityHintNode.textContent = `${hint} QR INDISPONIBLE: UTILISER LE LIEN DIRECT.`;
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
  if (isPdfMode()) {
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
  linkNode.hidden = false;
  linkNode.innerHTML = `
    <span>LIEN TELEPHONE :</span>
    <a href="${escapeHtml(absoluteUrl)}" target="_blank" rel="noopener">${escapeHtml(absoluteUrl)}</a>${hint ? `<small>${escapeHtml(hint)}</small>` : ""}
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

async function openMobileSignatureRequest(docType, personId, signer = "personnel") {
  if (!state.data) {
    showDataStatus("DONNEES NON CHARGEES");
    return;
  }
  const normalizedSigner = normalizeMobileSignatureSigner(signer);
  if (normalizedSigner === "representant" && !hasRepresentativeIdentityForDocument(docType)) {
    showDataStatus("IDENTITE DU REPRESENTANT OBLIGATOIRE AVANT VALIDATION");
    window.alert("VOUS DEVEZ IDENTIFIER L'IDENTITE DU REPRESENTANT DE L'ETABLISSEMENT POUR VALIDATION.");
    updateRepresentativeSignatureActionState(docType);
    return;
  }

  let request = getActiveMobileSignatureRequest(personId, docType, normalizedSigner);
  if (!request) {
    request = createMobileSignatureRequest(personId, docType, normalizedSigner);
    markDirty();
    await saveDataToFile({ silent: true });
  }

  const absoluteUrl = await getAbsoluteMobileSignatureUrl(request);
  window.open(absoluteUrl, "_blank", "noopener");
  renderMobileSignatureLink(docType, normalizedSigner, absoluteUrl);

  showDataStatus("PAGE DE SIGNATURE MOBILE OUVERTE");
  syncMobileSignaturePolling();
}

async function loadData() {
  bindPdfModalCleanup();
  reorderOverviewSearchBlock();
  restoreNavigationContext();
  clearSearchInputsOnInitialLoad();
  bindSearchClearOnBrowserEvents();
  applyActiveNav();
  bindHistoryNavigation();
  bindAutoSaveOnNavigation();
  bindGlobalShortcuts();
  bindLoadButton();
  bindSaveButtons();
  bindDirtyFallbackTracking();
  bindPdfButtons();
  bindMobileSignatureButtons();
  bindOverviewUrgencyActions();
  bindOverviewControlExport();
  bindDeletePersonButtons();
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
  bindReferenceFilters();
  bindArchiveFilterForm();
  bindSignatureCanvases();
  bindRepresentativeFields();

  const workingData = loadWorkingData();
  if (workingData) {
    state.data = workingData;
    migrateDataModel();
    state.isDirty = true;
    clearUndoStack();
    applyMeta();
    hydrateStaticLists();
    renderPage();
    clearSearchInputsOnInitialLoad();
    showDataStatus("DONNEES EN COURS REPRISES - SAUVEGARDER POUR LES RENDRE DEFINITIVES");
    scheduleBackgroundAutoSave();
    return;
  }

  await reloadData("OUVERTURE DES DONNEES...");
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
    clearUndoStack();
    applyMeta();
    hydrateStaticLists();
    renderPage();
    clearSearchInputsOnInitialLoad();
    showDataStatus(
      getDataBackendMode() === "SUPABASE" ? "DONNEES SUPABASE CHARGEES" : "DONNEES LOCALES CHARGEES"
    );
    state.previousSignatureValidationMap = buildSignatureValidationMap(state.data);
    window.setTimeout(() => {
      notifyFullySignedDocumentsOnReload(previousSignatureValidationMap);
    }, 0);
    window.setTimeout(() => {
      autoGenerateSignedDocumentsPdfIfMissing().catch((error) => {
        console.error(error);
      });
    }, 0);
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

function migrateDataModel() {
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
  state.data.listes.statutsObjetManuels = Array.from(
    new Set(state.data.listes.statutsObjetManuels.map(normalizeText).map((value) => (value === "CASSE" ? "HS" : value)))
  ).filter(Boolean);
  state.data.listes.causesRemplacement = Array.from(
    new Set(state.data.listes.causesRemplacement.map(normalizeText).map((value) => (value === "CASSE" ? "HS" : value)))
  ).filter(Boolean);
  if (!state.data.listes.causesRemplacement.length) {
    state.data.listes.causesRemplacement = [...EFFECT_STATUS_CAUSES];
  }
  state.data.listes.coutsRemplacement = state.data.listes.coutsRemplacement
    .map((entry) => ({
      typeEffet: normalizeText(entry.typeEffet),
      cause: normalizeText(entry.cause) === "CASSE" ? "HS" : normalizeText(entry.cause),
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
    .map((entry, index) => ({
      id: String(entry.id || `DOCARCH${String(index + 1).padStart(4, "0")}`),
      personId: String(entry.personId || ""),
      nom: normalizeText(entry.nom),
      prenom: normalizeText(entry.prenom),
      typeDocument: normalizeText(entry.typeDocument),
      dateDocument: String(entry.dateDocument || ""),
      sites: normalizeText(entry.sites),
      typePersonnel: normalizeText(entry.typePersonnel),
      typeContrat: normalizeText(entry.typeContrat),
      statutSignature: normalizeText(entry.statutSignature) || "EN ATTENTE",
      totalEffets: Number(entry.totalEffets || 0),
      totalFacturable: normalizeAmount(entry.totalFacturable),
      pdfPath: String(entry.pdfPath || ""),
      metadataPath: String(entry.metadataPath || ""),
      dateArchivage: String(entry.dateArchivage || ""),
      fingerprint: isLegacyArrivalArchiveFingerprint(entry.fingerprint) ? "" : String(entry.fingerprint || ""),
    }))
    .filter((entry) => entry.personId && entry.typeDocument && entry.pdfPath);

  state.data.stocksEffetsManuels = state.data.stocksEffetsManuels
    .map((entry, index) => ({
      id: String(entry.id || `STKM${String(index + 1).padStart(4, "0")}`),
      typeEffet: normalizeText(entry.typeEffet),
      site: normalizeText(entry.site),
      referenceEffetId: String(entry.referenceEffetId || ""),
      designation: normalizeText(entry.designation),
      action: normalizeText(entry.action),
      quantite: Math.max(1, Number.parseInt(String(entry.quantite || 1), 10) || 1),
      motif: normalizeText(entry.motif),
      commentaire: normalizeText(entry.commentaire),
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
    person.sitesAffectation = getPersonSites(person);
    person.site = getPersonSiteLabel(person);

    if (!person.typeContrat && LEGACY_CONTRACT_TYPES.includes(person.typePersonnel)) {
      person.typeContrat = person.typePersonnel;
      person.typePersonnel = "";
    }

    if (!Array.isArray(person.effetsConfies)) {
      person.effetsConfies = [];
    }

    person.effetsConfies.forEach((effect) => {
      effect.typeEffet = normalizeText(effect.typeEffet);
      effect.siteReference = normalizeText(effect.siteReference);
      effect.referenceEffetId = String(effect.referenceEffetId || "");
      effect.designation = normalizeText(effect.designation);
      effect.numeroIdentification = normalizeText(effect.numeroIdentification);
      effect.vehiculeImmatriculation = normalizeText(effect.vehiculeImmatriculation);
      effect.statutManuel = normalizeText(effect.statutManuel);
      const legacyCause = normalizeText(effect.causeRemplacement);
      if (!normalizeEffectCause(effect.cause) && legacyCause) {
        effect.cause = legacyCause;
      }
      effect.dateRemplacement = String(effect.dateRemplacement || "");
      effect.coutRemplacement = normalizeAmount(effect.coutRemplacement);
      effect.commentaire = normalizeText(effect.commentaire);

      if (!typeUsesReferenceCatalog(effect.typeEffet)) {
        effect.referenceEffetId = "";
        effect.designation = "";
      }

      if (!typeUsesSiteField(effect.typeEffet)) {
        effect.siteReference = "";
      } else if (!effect.siteReference) {
        effect.siteReference = getDefaultEffectSiteReference(person, effect);
      } else {
        effect.siteReference = normalizeText(effect.siteReference);
      }

      effect.coutRemplacement = getEffectReplacementCost(person, effect);
    });
  });

  if (!Array.isArray(state.data.listes.referencesEffets)) {
    state.data.listes.referencesEffets = [];
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
        normalizeText(reference.designation) === "CES-PG"
          ? [ALL_SITES_VALUE]
          : getReferenceSites(reference);
      return {
        ...reference,
        sitesAffectation: nextSites,
        site: nextSites.join(" / "),
      };
    });

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
  if (state.mobileSignaturePollTimerId) {
    window.clearInterval(state.mobileSignaturePollTimerId);
    state.mobileSignaturePollTimerId = 0;
  }
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
  return (
    getActiveMobileSignatureRequest(personId, docType, "personnel") ||
    getActiveMobileSignatureRequest(personId, docType, "representant")
  );
}

async function pollMobileSignatureRequest() {
  const activeRequest = getActiveDocumentMobileSignatureRequest();
  if (!activeRequest) {
    stopMobileSignaturePolling();
    return;
  }

  try {
    const json = await fetchLatestDataSnapshot();
    const requests = Array.isArray(json?.demandesSignatureMobile) ? json.demandesSignatureMobile : [];
    const nextRequest = requests.find((entry) => entry.token === activeRequest.token) || null;
    const person = Array.isArray(json?.personnes)
      ? json.personnes.find((entry) => String(entry.id || "") === String(activeRequest.personId || "")) || null
      : null;

    if (!nextRequest || !person) {
      stopMobileSignaturePolling();
      return;
    }

    const currentPage = document.body.dataset.page || "";
    const docType = currentPage === "exit-document" ? "exit" : "arrival";
    const signer = normalizeMobileSignatureSigner(activeRequest.signer || "");
    const previousPerson = getCurrentPerson();
    const previousSignature = getSignatureValue(previousPerson, docType, signer);
    const previousValidatedAt = getSignatureValidationDate(previousPerson, docType, signer);

    state.data = json;
    migrateDataModel();

    if (nextRequest.status !== "EN ATTENTE") {
      renderMobileSignatureLink(docType, signer, "");
      stopMobileSignaturePolling();
    }

    const updatedPerson = getCurrentPerson();
    const nextSignature = getSignatureValue(updatedPerson, docType, signer);
    const nextValidatedAt = getSignatureValidationDate(updatedPerson, docType, signer);

    if (previousSignature !== nextSignature || previousValidatedAt !== nextValidatedAt || nextRequest.status !== "EN ATTENTE") {
      renderPage();
      if (nextRequest.status === "SIGNEE") {
        showDataStatus(
          signer === "representant"
            ? "SIGNATURE MOBILE DU REPRESENTANT ENREGISTREE"
            : "SIGNATURE MOBILE DU PERSONNEL ENREGISTREE"
        );
      }
    }
  } catch (error) {
    // ignore polling errors
  }
}

function syncMobileSignaturePolling() {
  stopMobileSignaturePolling();
  const request = getActiveDocumentMobileSignatureRequest();
  if (!request) {
    return;
  }
  state.mobileSignaturePollTimerId = window.setInterval(() => {
    pollMobileSignatureRequest();
  }, 2500);
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
      if (!isDocumentFullySigned(person, docType)) {
        window.alert("GENERATION PDF IMPOSSIBLE : LE DOCUMENT DOIT ETRE SIGNE PAR LE PERSONNEL ET LE REPRESENTANT.");
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
    const canOpen = Boolean(person && isDocumentFullySigned(person, docType));
    button.classList.toggle("is-disabled", !canOpen);
    button.setAttribute("aria-disabled", canOpen ? "false" : "true");
    button.setAttribute(
      "title",
      canOpen
        ? "GENERER LE PDF"
        : "INDISPONIBLE : SIGNATURES PERSONNEL ET REPRESENTANT OBLIGATOIRES"
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
      await openMobileSignatureRequest(docType, personId, signer);
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
      renderPage();
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
  const defaults = {
    sheetEffects: { key: "typeEffet", dir: "asc" },
    arrivalEffects: { key: "typeEffet", dir: "asc" },
    exitEffects: { key: "typeEffet", dir: "asc" },
    overviewPersons: { key: "nom", dir: "asc" },
    referenceEffects: { key: "site", dir: "asc" },
  };
  const current = state.tableSorts?.[tableName];
  return current && current.key && current.dir ? current : (defaults[tableName] || { key: "nom", dir: "asc" });
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
  const currentEffects = getCurrentAssignedEffects(person);
  const movementMap = getArrivalComplementMovementMap(person, person?.effetsConfies || []);
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
    if (sort.key === "site") {
      const designationCompare = compareTextValues(
        String(left?.designation || ""),
        String(right?.designation || "")
      );
      if (designationCompare !== 0) {
        return sort.dir === "asc" ? designationCompare : -designationCompare;
      }
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

function bindEffectTableSorting() {
  document.querySelectorAll("[data-sort-table][data-sort-key]").forEach((header) => {
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
      renderPage();
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

function buildPdfWaitingHtml() {
  return `<!DOCTYPE html>
<html lang="fr">
  <head>
    <meta charset="utf-8">
    <title>GENERATION PDF</title>
    <style>
      body{margin:0;font-family:Segoe UI,Arial,sans-serif;background:#f4f1ea;color:#213b48;display:grid;place-items:center;min-height:100vh}
      .wrap{width:min(420px,calc(100vw - 32px));padding:24px;border:1px solid #adbec7;border-radius:14px;background:#fffdfa;box-shadow:0 14px 30px rgba(15,30,38,.16)}
      .eyebrow{margin:0 0 6px;font-size:11px;letter-spacing:.14em;color:#4a6170}
      h1{margin:0 0 10px;font-size:22px}
      p{margin:0 0 14px;color:#3f5662}
      .status{display:flex;align-items:center;gap:10px;margin-top:6px}
      .dot{width:14px;height:14px;border-radius:999px;background:#4c7787;box-shadow:0 0 0 0 rgba(76,119,135,.35);animation:pulse 1.2s infinite ease-out}
      .value{font-size:12px;letter-spacing:.08em;color:#213b48}
      @keyframes pulse{
        0%{transform:scale(.9);box-shadow:0 0 0 0 rgba(76,119,135,.35)}
        70%{transform:scale(1);box-shadow:0 0 0 10px rgba(76,119,135,0)}
        100%{transform:scale(.95);box-shadow:0 0 0 0 rgba(76,119,135,0)}
      }
    </style>
  </head>
  <body>
    <div class="wrap">
      <p class="eyebrow">EXPORT PDF</p>
      <h1>GENERATION DU PDF</h1>
      <p>PREPARATION DU DOCUMENT EN COURS...</p>
      <div class="status"><div class="dot"></div><div class="value">PATIENTEZ...</div></div>
    </div>
  </body>
</html>`;
}

function getDocumentPagePath(docType) {
  return normalizeText(docType) === "EXIT" ? "document-sortie.html" : "document-arrivee.html";
}

function getHostedPdfDocumentPath(docType, personId, mode = "STANDARD") {
  const pagePath = getDocumentPagePath(docType);
  return `${pagePath}?personId=${encodeURIComponent(personId)}&pdf=1&mode=${encodeURIComponent(normalizeText(mode || "STANDARD"))}`;
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
    popup.document.write(buildPdfWaitingHtml());
    popup.document.close();
  } catch (popupError) {
    console.error(popupError);
  }

  try {
    if (reusableArchive && !canPromoteReusableArchiveToStorage) {
      popup.location.href = getDocumentArchiveOpenPath(reusableArchive);
      showActionStatus("update", "PDF ARCHIVE REUTILISE");
      return;
    }
    if (canPromoteReusableArchiveToStorage) {
      showDataStatus("PDF ARCHIVE LOCAL DETECTE - MIGRATION VERS SUPABASE EN COURS");
    }

    if (getDataBackendMode() !== "LOCAL_API") {
      const hostedPath = getHostedPdfDocumentPath(docType, personId, archiveMode);
      const hostedUrl = `${hostedPath}&ts=${Date.now()}`;
      popup.location.href = hostedUrl;
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
    if (!response.ok) {
      throw new Error("PDF impossible");
    }

    const archiveSaved = response.headers.get("X-Archive-Saved") === "1";
    const archivePdfPath = response.headers.get("X-Archive-Pdf-Path") || "";
    const archiveMetadataPath = response.headers.get("X-Archive-Metadata-Path") || "";
    const blob = await response.blob();

    let hostedStoragePdfPath = "";
    if (shouldArchive && getDataBackendMode() === "LOCAL_API" && person) {
      try {
        const uploadResult = await uploadPdfBlobToSupabaseStorage(docType, person, blob, archiveMode);
        hostedStoragePdfPath = String(uploadResult?.storageRef || "");
        if (hostedStoragePdfPath) {
          console.info("[SUPABASE][PDF] final storage path", hostedStoragePdfPath);
          showDataStatus("PDF ARCHIVE ENVOYE VERS SUPABASE STORAGE");
        } else {
          console.warn("[SUPABASE][PDF] upload returned empty storageRef");
        }
      } catch (uploadError) {
        console.error("[SUPABASE][PDF] upload fail", uploadError);
        const message = String(uploadError?.message || "ERREUR INCONNUE").slice(0, 220);
        showDataStatus(`UPLOAD STORAGE PDF IMPOSSIBLE - ARCHIVAGE LOCAL CONSERVE (${message})`);
      }
    }

    const objectUrl = window.URL.createObjectURL(blob);
    hidePdfProgressModal();
    popup.location.href = objectUrl;
    const finalArchivePath = hostedStoragePdfPath || archivePdfPath;
    if (person && finalArchivePath) {
      clearPendingPdfTaskFor(person.id, docType);
      registerArchivedDocument(person, docType, finalArchivePath, archiveMetadataPath, archiveMode).catch((error) => {
        console.error(error);
        showDataStatus("ARCHIVAGE PDF IMPOSSIBLE");
      });
    } else if (shouldArchive && !finalArchivePath) {
      showDataStatus("PDF OUVERT - ARCHIVAGE NON REALISE");
    } else if (!shouldArchive) {
      showDataStatus("PDF OUVERT - DOCUMENT NON SIGNE, UPLOAD SUPABASE NON LANCE");
    }
    window.setTimeout(() => {
      try {
        window.URL.revokeObjectURL(objectUrl);
      } catch (error) {
        console.error(error);
      }
    }, 60000);
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
  if (!response.ok) {
    throw new Error(`AUTO PDF IMPOSSIBLE (${docType}/${person.id})`);
  }

  const archivePdfPath = response.headers.get("X-Archive-Pdf-Path") || "";
  const archiveMetadataPath = response.headers.get("X-Archive-Metadata-Path") || "";
  const blob = await response.blob();

  let hostedStoragePdfPath = "";
  if (isSupabaseConfigured()) {
    try {
      const uploadResult = await uploadPdfBlobToSupabaseStorage(docType, person, blob, archiveMode);
      hostedStoragePdfPath = String(uploadResult?.storageRef || "");
    } catch (error) {
      console.error("[SUPABASE][PDF][AUTO] upload fail", error);
    }
  }

  const finalArchivePath = hostedStoragePdfPath || archivePdfPath;
  if (!finalArchivePath) {
    return false;
  }
  await registerArchivedDocument(person, docType, finalArchivePath, archiveMetadataPath, archiveMode);
  return true;
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
  state.autoPdfGenerationInFlight = true;
  try {
    for (const candidate of candidates) {
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
  const personLabel = person ? `${person.nom || ""} ${person.prenom || ""}`.trim() : "";
  const docLabel = latestRequest?.docType === "exit" ? "SORTIE" : "ENTREE";
  const messageLines = [
    "DOCUMENT SIGNE (2 SIGNATURES VALIDEES) :",
    ...labels,
    "",
    `VOUS AVEZ UN PDF A CREER POUR ${personLabel || "CE PERSONNEL"} - DOCUMENT DE ${docLabel}.`,
    "OK = OUVRIR LE DOCUMENT",
  ];

  const shouldOpenDocument = window.confirm(messageLines.join("\n"));
  if (!shouldOpenDocument) {
    reminderSnoozeMap[snoozeKey] = Date.now() + 120 * 1000;
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
    window.alert("VOTRE BASE N'EST PAS A JOUR");
    return;
  }

  if (person && latestRequest?.docType) {
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
    setCurrentPersonId(person.id, "replace");
    const pagePath = getDocumentPagePath(latestRequest.docType);
    navigateWithAutoSave(`${pagePath}?personId=${encodeURIComponent(person.id)}&focusPdf=${encodeURIComponent(latestRequest.docType)}`);
    const seenKey = `SIG:${person.id}:${latestRequest.docType}:${latestRequest.signer}:${latestRequest.validatedAt}`;
    state.signedDocumentsPopupSeenKeys.add(seenKey);
    window.alert("DOCUMENT OUVERT - GENERER LE PDF");
  }
}

function bindDeletePersonButtons() {
  document.querySelectorAll(".js-delete-person").forEach((button) => {
    button.onclick = () => {
      const personId = button.getAttribute("data-person-id") || getCurrentPersonId();
      if (!personId) {
        showDataStatus("AUCUNE PERSONNE SELECTIONNEE");
        return;
      }
      deletePerson(personId);
    };
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

    const applyFullReset = () => {
      state.filters = { ...DEFAULT_FILTERS };
      state.urgentMode = false;
      saveNavigationContext({ filters: state.filters, personId: "", urgentMode: false });
      setCurrentPersonId("", "replace");
      clearFormSearchFields(form);
      applyFiltersToForm(form);
      updateUrgencyModeUi();
      renderPage();
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
      state.filters = {
        ...DEFAULT_FILTERS,
        ...readFilters(form),
      };
      saveNavigationContext({ filters: state.filters, urgentMode: state.urgentMode });
      renderPage();
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
function bindArchiveFilterForm() {
  const form = document.getElementById("documents-archives-filter-form");
  if (!form) {
    return;
  }

  const resetArchiveFilters = () => {
    clearFormSearchFields(form);
    const searchField = form.elements.archiveSearch;
    if (searchField instanceof HTMLInputElement) {
      searchField.value = "";
      searchField.defaultValue = "";
    }
    ["archiveTypeDocument", "archiveSite", "archiveStatutSignature"].forEach((fieldName) => {
      const field = form.elements[fieldName];
      if (field instanceof HTMLSelectElement) {
        field.value = "";
      }
    });
  };

  const applyArchiveReset = () => {
    setCurrentPersonId("", "replace");
    resetArchiveFilters();
    renderDocumentsArchivePage();
  };

  form.oninput = () => {
    renderDocumentsArchivePage();
  };

  form.onreset = (event) => {
    event.preventDefault();
    setCurrentPersonId("", "replace");
    resetArchiveFilters();
    renderDocumentsArchivePage();
  };

  const searchField = form.elements.archiveSearch;
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
    const selectedSites = readSelectedSites(form, "add");
    const person = {
      id: getNextId("P", state.data.personnes || []),
      nom: normalizeText(formData.get("nom")),
      prenom: normalizeText(formData.get("prenom")),
      sitesAffectation: selectedSites,
      site: "",
      typePersonnel: normalizeText(formData.get("typePersonnel")),
      typeContrat: normalizeText(formData.get("typeContrat")),
      dateEntree: String(formData.get("dateEntree") || ""),
      dateSortiePrevue: String(formData.get("dateSortiePrevue") || ""),
      dateSortieReelle: String(formData.get("dateSortieReelle") || ""),
      effetsConfies: [],
    };

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
    renderPage();
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

    setSheetFieldMissingState("sheetNom", !nom);
    setSheetFieldMissingState("sheetPrenom", !prenom);
    setSheetFieldMissingState("sheetFonction", !fonction);
    setSheetFieldMissingState("sheetTypePersonnel", !typePersonnel);
    setSheetFieldMissingState("sheetTypeContrat", !typeContrat);
    setSheetFieldMissingState("sheetDateEntree", !dateEntree);
    setSheetFieldMissingState("sheetDateSortiePrevue", needsExpectedExitDate && !dateSortiePrevue);

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
      dateEntree: String(formData.get("sheetDateEntree") || ""),
      dateSortiePrevue: String(formData.get("sheetDateSortiePrevue") || ""),
      dateSortieReelle: String(formData.get("sheetDateSortieReelle") || ""),
      effetsConfies: [],
    };

    person.site = getPersonSiteLabel(person);
    return person;
  };

  form.onsubmit = (event) => {
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
    person.dateEntree = String(formData.get("sheetDateEntree") || "");
    person.dateSortiePrevue = String(formData.get("sheetDateSortiePrevue") || "");
    person.dateSortieReelle = String(formData.get("sheetDateSortieReelle") || "");

    markDirty();
    renderPage();
    renderPersonSheet(person.id);
    showActionStatus("update", `FICHE MISE A JOUR : ${person.nom} ${person.prenom}`);
  };

  if (addButton) {
    addButton.onclick = () => {
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
      renderPage();
      renderPersonSheet(person.id);
      showActionStatus("create", `PERSONNE AJOUTEE : ${person.nom} ${person.prenom}`);
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
    form.elements.designationLibre.oninput = () => {
      syncReplacementCostField();
      updateEffectRequiredHighlights(form);
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

    const formData = new FormData(form);
    const typeEffet = normalizeText(formData.get("typeEffet"));
    const referenceSite = normalizeText(formData.get("referenceSite"));
    const usesReferenceCatalog = typeUsesReferenceCatalog(typeEffet);
    const usesSiteField = typeUsesSiteField(typeEffet);
    const referenceEffetId = usesReferenceCatalog ? String(formData.get("referenceEffet") || "") : "";
    const reference = findReferenceById(referenceEffetId);
    const designationLibre = usesReferenceCatalog ? normalizeText(formData.get("designationLibre")) : "";
    const availableReferenceSites = getAvailableReferenceSites(person);
    const resolvedReferenceSite = usesReferenceCatalog
      ? normalizeText(
          reference?.site || referenceSite || (availableReferenceSites.length === 1 ? availableReferenceSites[0] : "")
        )
      : usesSiteField
        ? normalizeText(referenceSite || (availableReferenceSites.length === 1 ? availableReferenceSites[0] : ""))
        : "";
    const dateRemplacement = String(formData.get("dateRemplacement") || "");
    const coutRemplacement = normalizeAmount(formData.get("coutRemplacement"));
    const manualStatus = normalizeText(formData.get("statutManuel"));

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

    if (!manualStatus) {
      showDataStatus("SELECTIONNER LE STATUT MANUEL");
      form.elements.statutManuel?.focus();
      return;
    }

    if (usesReferenceCatalog && !referenceEffetId) {
      showDataStatus("CHOISIR UNE CLE EXISTANTE DANS LA LISTE");
      return;
    }

    const effectId = mode === "edit" ? state.editingEffectId : getNextId("E", person.effetsConfies || []);
    const vehiculeImmatriculation =
      typeEffet === "TELECOMMANDE URMET" ? normalizeText(formData.get("vehiculeImmatriculation")) : "";

    const existingEffect =
      mode === "edit"
        ? (person.effetsConfies || []).find((entry) => String(entry?.id || "") === String(effectId))
        : null;
    const nextCause = getCauseFromManualStatus(manualStatus);
    const preservedCause = normalizeEffectCause(existingEffect?.cause || existingEffect?.causeRemplacement);
    const effect = {
        id: effectId,
        typeEffet,
        siteReference: resolvedReferenceSite,
        referenceEffetId,
        designation: usesReferenceCatalog ? reference?.designation || "" : "",
        numeroIdentification: normalizeText(formData.get("numeroIdentification")),
        vehiculeImmatriculation,
      dateRemise: String(formData.get("dateRemise") || ""),
      dateRetour: String(formData.get("dateRetour") || ""),
      statutManuel: manualStatus === "CASSE" ? "HS" : manualStatus,
      cause: preservedCause || nextCause,
      dateRemplacement,
      coutRemplacement,
      commentaire: normalizeText(formData.get("commentaire")),
    };
    if (usesReferenceCatalog && !effect.designation) {
      effect.designation = `EFFET ${effect.id}`;
    }

    if (!Array.isArray(person.effetsConfies)) {
      person.effetsConfies = [];
    }
    pushUndoSnapshot(mode === "edit" ? "MODIFICATION EFFET" : "AJOUT EFFET");
    const existingIndex =
      mode === "edit" ? person.effetsConfies.findIndex((entry) => entry.id === effect.id) : -1;
    if (mode === "edit" && existingIndex >= 0) {
      person.effetsConfies[existingIndex] = effect;
    } else {
      person.effetsConfies.push(effect);
    }
    markEffectRowFlash(mode === "edit" ? "update" : "create", person.id, effect.id);

    markDirty();
    form.reset();
    state.editingEffectId = "";
    hydrateEffectReferenceSiteSelect(person, "", "");
    hydrateReferenceSelect(person, "", "", "");
    updateEffectFormMode("");
    renderPage();
    renderPersonSheet(person.id);
    const effectLabel = effect.designation || effect.numeroIdentification || effect.id;
    showActionStatus(
      mode === "edit" ? "update" : "create",
      mode === "edit"
        ? `EFFET MODIFIE : ${effectLabel}`
        : `EFFET AJOUTE : ${effectLabel}`
    );

    await saveAfterEffectChangeWithAvenantAlert();
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
      resetEffectFormFieldsExceptCost();
    };
  }

  form.addEventListener("input", () => {
    updateEffectResetButtonState(form);
  });
  form.addEventListener("change", () => {
    updateEffectResetButtonState(form);
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
    `SUPPRIMER DEFINITIVEMENT CET EFFET${effect?.designation ? ` : ${effect.designation}` : effect?.numeroIdentification ? ` : ${effect.numeroIdentification}` : ""} ?`
  );
  if (!confirmDelete) {
    return;
  }

  pushUndoSnapshot("SUPPRESSION EFFET");
  person.effetsConfies = person.effetsConfies.filter((effect) => effect.id !== effectId);
  markEffectTableFlash("delete", personId);
  if (state.editingEffectId === effectId) {
    state.editingEffectId = "";
    resetEffectForm();
  }
  markDirty();
  renderPage();
  renderPersonSheet(personId);
  showActionStatus("delete", `EFFET SUPPRIME : ${effectId}`);
  await saveAfterEffectChangeWithAvenantAlert();
}

async function saveAfterEffectChangeWithAvenantAlert() {
  await saveDataToFile({
    silent: true,
    reloadAfter: false,
  });
  if (state.isDirty) {
    showDataStatus("SAUVEGARDE IMPOSSIBLE - ALERTE ANNULEE");
    return;
  }
  window.alert(
    "DES MODIFICATIONS D'EFFETS ONT ETE EFFECTUEES. VOUS DEVEZ DONC PROCEDER A UNE NOUVELLE SIGNATURE DE L'AVENANT."
  );
}

function deletePerson(personId) {
  if (!state.data?.personnes) {
    return;
  }

  const person = state.data.personnes.find((entry) => entry.id === personId);
  if (!person) {
    return;
  }

  const confirmDelete = window.confirm(
    `SUPPRIMER DEFINITIVEMENT ${person.nom} ${person.prenom} ?`
  );
  if (!confirmDelete) {
    return;
  }

  pushUndoSnapshot("SUPPRESSION PERSONNE");
  state.data.personnes = state.data.personnes.filter((entry) => entry.id !== personId);
  if (getCurrentPersonId() === personId) {
    setCurrentPersonId("");
  }
  state.editingEffectId = "";
  markDirty();
  renderPage();
  showActionStatus("delete", `PERSONNE SUPPRIMEE : ${person.nom} ${person.prenom}`);

  if (document.body.dataset.page === "person-sheet") {
    navigateWithAutoSave("fiche-personne.html");
  }
}

function startEditEffect(personId, effectId) {
  const person = state.data?.personnes?.find((entry) => entry.id === personId);
  const effect = person?.effetsConfies?.find((entry) => entry.id === effectId);
  const form = document.getElementById("effect-form");
  if (!person || !effect || !form) {
    return;
  }

  state.editingEffectId = effectId;
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
  updateEffectActionButtons();
  updateEffectResetButtonState(form);
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
  form.reset();
  hydrateEffectReferenceSiteSelect(getCurrentPerson(), "", "");
  hydrateReferenceSelect(getCurrentPerson() || "", "", "", "");
  updateEffectFormMode("");
  updateEffectActionButtons();
  updateEffectResetButtonState(form);
  updateManualStatusCriticalState(form);
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
  return ["PERDU", "VOL", "HS"].includes(normalizeText(value));
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
  const addButton = document.getElementById("effect-add-button");
  const updateButton = document.getElementById("effect-update-button");
  const deleteButton = document.getElementById("effect-delete-button");
  const cancelButton = document.getElementById("effect-cancel-button");
  const isEditing = Boolean(state.editingEffectId);

  if (addButton) {
    addButton.disabled = false;
    addButton.classList.toggle("button--primary", !isEditing);
    addButton.classList.toggle("button--secondary", isEditing);
  }
  if (updateButton) {
    updateButton.disabled = !isEditing;
    updateButton.classList.toggle("button--primary", isEditing);
    updateButton.classList.toggle("button--secondary", !isEditing);
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
  if (normalizedType === "CLE" || normalizedType === "CLE CES") {
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

function updateEffectRequiredHighlights(form) {
  if (!form) {
    return;
  }
  const typeEffet = normalizeText(form.elements.typeEffet?.value || "");
  const site = normalizeText(form.elements.referenceSite?.value || "");
  const statutManuel = normalizeText(form.elements.statutManuel?.value || "");
  const siteRequired = Boolean(form.elements.referenceSite?.required);

  setEffectFieldMissingState(form, "typeEffet", !typeEffet);
  setEffectFieldMissingState(form, "referenceSite", siteRequired && !site);
  setEffectFieldMissingState(form, "statutManuel", !statutManuel);
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

  if (["CLE", "CLE CES"].includes(normalizedType)) {
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
    } else {
      helpNode.textContent =
        availableReferenceSites.length > 1
          ? "POUR UNE CLE : CHOISIR D'ABORD LE SITE, PUIS LE NOM DE LA CLE"
          : "POUR UNE CLE : CHOISIR UN NOM DE CLE DU SITE";
    }
  } else if (["BADGE INTRUSION", "TELECOMMANDE URMET", "CARTE TURBOSELF"].includes(normalizedType)) {
    showReferenceSite = true;
    referenceSiteLabel.textContent =
      normalizedType === "BADGE INTRUSION"
        ? "SITE DU BADGE"
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
  form.elements.statutManuel.required = true;
  updateEffectRequiredHighlights(form);
}

function isCesKeyDesignation(designation) {
  return normalizeText(designation).startsWith("CES-");
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
      normalizePricingKey(entry?.cause, { cause: true }) === normalizedCause
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

  if (normalizedType === "BADGE INTRUSION") {
    return 15;
  }

  if (normalizedType === "TELECOMMANDE URMET") {
    return 40;
  }

  if (normalizedType === "CARTE TURBOSELF") {
    return 10;
  }

  return 0;
}

function getEffectReplacementCause(person, effect) {
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
          renderPage();
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
      renderPage();
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
        renderPage();
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
      renderPage();
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
    renderPage();
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
  const syncReferenceTableFiltersFromEditor = () => {
    const filterForm = document.getElementById("reference-filter-form");
    if (!filterForm) {
      return;
    }
    const selectedSite = normalizeText(form.elements.referenceSite?.value || "");
    const selectedTypeEffet = normalizeText(form.elements.referenceTypeEffet?.value || "");
    if (filterForm.elements.filterReferenceSite) {
      filterForm.elements.filterReferenceSite.value = selectedSite;
    }
    if (filterForm.elements.filterReferenceTypeEffet) {
      filterForm.elements.filterReferenceTypeEffet.value = selectedTypeEffet;
    }
    renderReferenceEffectsTable(state.referenceRenderContext || buildReferenceRenderContext());
  };

  if (typeField) {
    typeField.onchange = () => {
      updateReferenceEffectFormMode(typeField.value);
      syncReferenceSitesSelector();
      syncReferenceTableFiltersFromEditor();
    };
  }
  if (siteField) {
    siteField.onchange = () => {
      syncReferenceTableFiltersFromEditor();
    };
  }
  if (designationField) {
    designationField.oninput = () => {
      syncReferenceSitesSelector();
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
    const selectedReferenceSites = Array.from(
      form.querySelectorAll('input[name="referenceSites"]:checked')
    ).map((input) => normalizeText(input.value));
    const referenceId = state.editingReferenceId || getNextId("REF", state.data.listes.referencesEffets);
    const reference = {
      id: referenceId,
      site: normalizeText(formData.get("referenceSite")) || "SANS SITE",
      sitesAffectation:
        normalizedTypeEffet === "CLE CES"
          ? normalizeSites(selectedReferenceSites)
          : normalizeSites([normalizeText(formData.get("referenceSite")) || "SANS SITE"]),
      typeEffet: normalizedTypeEffet || "EFFET",
      designation: normalizeText(formData.get("referenceDesignation")) || `REFERENCE ${referenceId}`,
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
    renderPage();
    form.reset();
    renderReferenceSitesSelector([]);
    updateReferenceEffectFormMode("");
    syncReferenceTableFiltersFromEditor();
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
  const normalizedType = normalizeText(typeEffet);
  if (!field) {
    return;
  }
  field.classList.toggle("is-hidden", normalizedType !== "CLE CES");
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

function getReplacementCostKey(typeEffet, cause) {
  return `${normalizeText(typeEffet)}__${normalizeText(cause)}`;
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
      (entry) => getReplacementCostKey(entry.typeEffet, entry.cause) === lookupKey
    );

    pushUndoSnapshot(currentIndex >= 0 ? "MODIFICATION COUT" : "AJOUT COUT");
    if (currentIndex >= 0) {
      state.data.listes.coutsRemplacement[currentIndex] = { typeEffet, cause, montant };
    } else {
      state.data.listes.coutsRemplacement.push({ typeEffet, cause, montant });
    }

    state.editingReplacementCostKey = "";
    markDirty();
    renderPage();
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

  form.onsubmit = (event) => {
    event.preventDefault();
    if (!Array.isArray(state.data?.stocksEffetsManuels)) {
      showDataStatus("DONNEES NON CHARGEES");
      return;
    }
    const formData = new FormData(form);
    const typeEffet = normalizeText(formData.get("stockTypeEffet"));
    const site = normalizeText(formData.get("stockSite"));
    const referenceEffetId = String(formData.get("stockReferenceId") || "");
    const reference = findReferenceById(referenceEffetId);
    const designation = normalizeText(reference?.designation || "");
    const action = normalizeText(formData.get("stockAction"));
    const quantite = Math.max(1, Number.parseInt(String(formData.get("stockQuantity") || "1"), 10) || 1);
    const motif = normalizeText(formData.get("stockReason"));
    const commentaire = normalizeText(formData.get("stockComment"));

    if (!typeEffet || !site || !reference || !designation || !action) {
      showDataStatus("TYPE, SITE, DESIGNATION BASE ET MOUVEMENT OBLIGATOIRES");
      return;
    }
    if (!isReferenceEffectActive(reference)) {
      showDataStatus("REFERENCE EFFET DESACTIVEE - MOUVEMENT BLOQUE");
      return;
    }

    pushUndoSnapshot("MOUVEMENT STOCK MANUEL");
    state.data.stocksEffetsManuels.push({
      id: getNextId("STKM", state.data.stocksEffetsManuels),
      typeEffet,
      site,
      referenceEffetId: String(reference.id || ""),
      designation,
      action,
      quantite,
      motif,
      commentaire,
      date: getTodayIsoDate(),
    });
    markDirty();
    renderReferenceBases();
    form.reset();
    if (form.elements.stockAction) {
      form.elements.stockAction.value = "ENTREE";
    }
    if (form.elements.stockQuantity) {
      form.elements.stockQuantity.value = "1";
    }
    showActionStatus("create", `MOUVEMENT STOCK ENREGISTRE : ${typeEffet} / ${designation}`);
  };

  const typeSelect = form.elements.stockTypeEffet;
  const siteSelect = form.elements.stockSite;
  if (typeSelect instanceof HTMLSelectElement) {
    typeSelect.onchange = () => updateStockDesignationOptions();
  }
  if (siteSelect instanceof HTMLSelectElement) {
    siteSelect.onchange = () => updateStockDesignationOptions();
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
    if (!(deleteButton instanceof HTMLElement)) {
      return;
    }
    event.preventDefault();
    const movementId = String(deleteButton.dataset.stockMovementId || "");
    const movement = (state.data?.stocksEffetsManuels || []).find((entry) => String(entry.id || "") === movementId);
    if (!movement) {
      return;
    }
    const linkedReference = movement.referenceEffetId ? findReferenceById(movement.referenceEffetId) : null;
    if (!linkedReference) {
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
  });
  body.dataset.bound = "true";
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
  if (normalized === "CASSE") return "HS";
  if (normalized === "PERDU") return "PERTE";
  return normalized;
}

function bindReferenceFilters() {
  const form = document.getElementById("reference-filter-form");
  if (!form) {
    return;
  }
  const applyReferenceReset = () => {
    clearFormSearchFields(form);
    form.reset();
    window.setTimeout(() => {
      renderReferenceEffectsTable();
    }, 0);
  };

  form.oninput = () => {
    renderReferenceEffectsTable();
  };

  form.onreset = () => {
    clearFormSearchFields(form);
    window.setTimeout(() => {
      renderReferenceEffectsTable();
    }, 0);
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

function renderPage() {
  const page = document.body.dataset.page || "";
  const persons = getFilteredPersons();
  let currentPersonId = getCurrentPersonId();
  const personExists = (state.data?.personnes || []).some(
    (entry) => String(entry?.id || "") === String(currentPersonId || "")
  );

  if (page === "person-sheet" && (!currentPersonId || !personExists)) {
    const fallbackPerson =
      (state.data?.personnes || [])[0] ||
      persons[0] ||
      null;
    if (fallbackPerson?.id) {
      setCurrentPersonId(String(fallbackPerson.id || ""), "replace");
    }
    currentPersonId = getCurrentPersonId();
  }

  if (page === "overview") {
    renderOverview(persons);
  }

  if (page === "global") {
    renderGlobalEffectsChart(persons);
    renderGlobalTable(persons);
  }

  if (page === "documents-archives") {
    renderDocumentsArchivePage();
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
    renderMobileSignaturePage();
    refreshDocumentSignatureCanvases(getCurrentMobileSignatureDocType());
  }

  if (page === "person-sheet") {
    try {
      renderPersonSheet(currentPersonId);
      bindEffectTableSorting();
      updateSortableHeaders("sheetEffects");
    } catch (error) {
      console.error("Erreur affichage fiche personne", error);
      showDataStatus("ERREUR AFFICHAGE FICHE PERSONNE - VOIR CONSOLE");
    }
  }

  if (page === "arrival-document") {
    renderArrivalDocument(currentPersonId);
    refreshDocumentSignatureCanvases("arrival");
    updateSortableHeaders("arrivalEffects");
    syncMobileSignaturePolling();
  }

  if (page === "exit-document") {
    renderExitDocument(currentPersonId);
    refreshDocumentSignatureCanvases("exit");
    updateSortableHeaders("exitEffects");
    syncMobileSignaturePolling();
  }

  if (page === "reference-bases") {
    renderReferenceBases();
    bindEffectTableSorting();
    updateSortableHeaders("referenceEffects");
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
      const width = Math.max(8, Math.round((row.total / maxValue) * 100));
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
  const folder = normalizeText(entry?.typeDocument) === "SORTIE" ? "archives/pdf/sortie" : "archives/pdf/arrivee";
  const name = `${sanitizeFilePart(entry?.dateDocument || getTodayIsoDate())}_${sanitizeFilePart(entry?.nom)}_${sanitizeFilePart(entry?.prenom)}.pdf`;
  return `${folder}/${name}`;
}

function getDocumentArchiveSignatureStatus(entry) {
  return normalizeText(entry?.statutSignature) || "EN ATTENTE";
}

function getDocumentTypeLabel(docType) {
  return normalizeText(docType) === "EXIT" ? "SORTIE" : "ARRIVEE";
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
  return getDocumentArchiveEntryMode(entry) === "COMPLEMENTAIRE" ? "AVENANT" : "INITIAL";
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

function getLatestSignedArrivalArchiveForPerson(personId) {
  const entries = (state.data?.documentsArchives || [])
    .filter(
      (entry) =>
        String(entry?.personId || "") === String(personId || "") &&
        normalizeText(entry?.typeDocument) === "ARRIVEE" &&
        normalizeText(entry?.statutSignature) === "SIGNE"
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

  if (String(effect?.dateRetour || "").trim()) {
    return "RENDU";
  }

  const effectStatus = normalizeText(getEffectStatus(person, effect));
  const effectCause = normalizeText(getEffectReplacementCause(person, effect));

  if (effectStatus === "DETRUIT") {
    return "DETRUIT";
  }
  if (effectStatus === "HS") {
    return "HS";
  }
  if (effectCause === "VOL") {
    return "VOLE";
  }
  if (effectStatus === "PERDU" || effectCause === "PERTE") {
    return "PERDU";
  }
  if (effectStatus === "NON RENDU") {
    return "NON RENDU";
  }
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
  if (/^https?:\/\//i.test(raw)) {
    return raw;
  }
  const storageRef = parseStorageSchemePath(raw);
  if (storageRef) {
    return getSupabaseStoragePublicUrl(storageRef.bucket, storageRef.objectPath) || "";
  }
  return raw.replace(/^\/+/, "");
}

function getArchiveEntrySites(entry) {
  return normalizeSites(String(entry?.sites || "").split("/").map((value) => value.trim()));
}

function archiveEntryMatchesSite(entry, site) {
  const normalizedSite = normalizeText(site);
  if (!normalizedSite) {
    return true;
  }
  const sites = getArchiveEntrySites(entry);
  return sites.includes(ALL_SITES_VALUE) || sites.includes(normalizedSite);
}

function buildDocumentArchiveEntry(person, docType, pdfPath, metadataPath, archiveMode = "STANDARD") {
  const effects = person?.effetsConfies || [];
  const documentMode = normalizeText(archiveMode || "STANDARD");
  const baseId = `DOC-${getDocumentTypeLabel(docType)}-${person?.id || ""}`;
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

async function registerArchivedDocument(person, docType, pdfPath, metadataPath, archiveMode = "STANDARD") {
  if (!state.data || !person || !pdfPath) {
    return;
  }
  if (!isDocumentFullySigned(person, docType)) {
    return;
  }
  upsertDocumentArchiveEntry(buildDocumentArchiveEntry(person, docType, pdfPath, metadataPath, archiveMode));
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
    return;
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
  const archiveSearchField = filterForm?.elements?.archiveSearch;
  if (archiveSearchField instanceof HTMLInputElement && lockedPerson && !String(archiveSearchField.value || "").trim()) {
    const label = getPersonPickerLabel(lockedPerson);
    archiveSearchField.value = label;
    archiveSearchField.defaultValue = label;
  }
  const search = normalizeText(archiveSearchField?.value);
  const searchTokens = search
    ? search
        .split(/[\s\-–—/]+/)
        .map((token) => normalizeText(token))
        .filter(Boolean)
    : [];
  const typeDocument = normalizeText(filterForm?.elements?.archiveTypeDocument?.value);
  const site = normalizeText(filterForm?.elements?.archiveSite?.value);
  const statutSignature = normalizeText(filterForm?.elements?.archiveStatutSignature?.value);
  const archiveSiteSelect = filterForm?.elements?.archiveSite;
  syncSelectOptions(archiveSiteSelect, state.data?.listes?.sites || [], "TOUS");

  let totalArchives = 0;
  let totalArrivalArchives = 0;
  let totalExitArchives = 0;
  const personsById = new Map(
    (state.data?.personnes || []).map((person) => [String(person?.id || ""), person])
  );
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
  const signedArchives = (state.data?.documentsArchives || []).filter(
    (entry) => getDocumentArchiveSignatureStatus(entry) === "SIGNE"
  );
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
  const archives = signedArchives.filter((entry) => {
    totalArchives += 1;
    if (normalizeText(entry.typeDocument) === "ARRIVEE") {
      totalArrivalArchives += 1;
    }
    if (normalizeText(entry.typeDocument) === "SORTIE") {
      totalExitArchives += 1;
    }
    if (!archiveMatchesLockedPerson(entry)) {
      return false;
    }
    if (typeDocument && normalizeText(entry.typeDocument) !== typeDocument) {
      return false;
    }
    if (!archiveEntryMatchesSite(entry, site)) {
      return false;
    }
    if (statutSignature && getDocumentArchiveSignatureStatus(entry) !== statutSignature) {
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
  const groupedArchives = archives.slice().sort((left, right) => {
    const leftPersonKey = String(left.personId || "")
      || `${normalizeText(left.nom)}|${normalizeText(left.prenom)}`;
    const rightPersonKey = String(right.personId || "")
      || `${normalizeText(right.nom)}|${normalizeText(right.prenom)}`;
    const personCompare = leftPersonKey.localeCompare(rightPersonKey, "fr");
    if (personCompare !== 0) {
      return personCompare;
    }

    const typeRank = (entry) => {
      const type = normalizeText(entry?.typeDocument || "");
      if (type === "ARRIVEE") return 0;
      if (type === "SORTIE") return 1;
      return 2;
    };
    const rankCompare = typeRank(left) - typeRank(right);
    if (rankCompare !== 0) {
      return rankCompare;
    }

    const leftDate = Date.parse(String(left?.dateDocument || "")) || 0;
    const rightDate = Date.parse(String(right?.dateDocument || "")) || 0;
    if (leftDate !== rightDate) {
      return leftDate - rightDate;
    }
    return String(left?.id || "").localeCompare(String(right?.id || ""), "fr");
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
  const allSignedArchives = signedArchives;
  const arrivalStorageCount = allSignedArchives.filter(
    (entry) => normalizeText(entry?.typeDocument) === "ARRIVEE"
  ).length;
  const exitStorageCount = allSignedArchives.filter(
    (entry) => normalizeText(entry?.typeDocument) === "SORTIE"
  ).length;
  const latestArchiveMs = allSignedArchives.reduce((latest, entry) => {
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
    body.innerHTML = buildEmptyTableRow(body, "AUCUN DOCUMENT ARCHIVE", 11);
    return;
  }

  const rowsHtml = groupedArchives
    .map(
      (entry) => {
        const display = resolveArchiveDisplayData(entry);
        const openPath = getDocumentArchiveOpenPath(entry);
        const typeLabel = normalizeText(entry.typeDocument || "");
        const typeIcon = typeLabel === "SORTIE" ? "🔴" : typeLabel === "ARRIVEE" ? "🟢" : "⚪";
        const typeTitle = typeLabel || "TYPE INCONNU";
        return `<tr>
        <td>${escapeHtml(display.nom)}</td>
        <td>${escapeHtml(display.prenom)}</td>
        <td title="${escapeHtml(typeTitle)}" aria-label="${escapeHtml(typeTitle)}">${typeIcon}</td>
        <td>${escapeHtml(formatDate(entry.dateDocument) || "-")}</td>
        <td>${escapeHtml(formatTime(entry.dateArchivage) || "-")}</td>
        <td>${escapeHtml(display.sites)}</td>
        <td>${escapeHtml(getDocumentArchiveSignatureStatus(entry))}</td>
        <td>${escapeHtml(String(entry.totalEffets ?? "-"))}</td>
        <td>${formatAmountWithEuro(entry.totalFacturable || 0)}</td>
        <td>${escapeHtml(getDocumentArchiveVersionLabel(entry))}</td>
        <td class="archive-actions-cell">${openPath ? `<a class="archive-pdf-button" href="${escapeHtml(openPath)}" target="_blank" rel="noopener" aria-label="OUVRIR PDF"><span class="archive-pdf-button__icon" aria-hidden="true"><img src="https://dphrvdhqhgycmllietuk.supabase.co/storage/v1/object/public/ui-assets/ui/icone-pdf.png" alt="" class="archive-pdf-button__image" /></span></a>` : "-"} <button type="button" class="table-link js-delete-archive-row" data-archive-id="${escapeHtml(String(entry.id || ""))}">SUPPRIMER</button></td>
      </tr>`;
      }
    );

  renderTableRowsProgressively(body, rowsHtml, buildEmptyTableRow(body, "AUCUN DOCUMENT ARCHIVE", 11), 24);
  bindArchiveRowActions();
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

  body.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
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
  const storageRef = parseStorageSchemePath(rawValue);
  if (storageRef) {
    return getSupabaseStoragePublicUrl(storageRef.bucket, storageRef.objectPath) || "";
  }
  return rawValue;
}

function getRepresentativeInfo(person, docType) {
  if (!person?.representants?.[docType]) {
    return { id: "", nom: "", fonction: "" };
  }
  return {
    id: String(person.representants[docType].id || ""),
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

function populateRepresentativeSelect(select, selectedId = "") {
  if (!(select instanceof HTMLSelectElement)) {
    return;
  }
  const currentValue = String(selectedId || "");
  const representatives = (state.data?.listes?.representantsSignataires || []).slice();
  select.innerHTML = ['<option value="">SELECTIONNER</option>']
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
  return Boolean(
    nameInput instanceof HTMLSelectElement &&
      functionInput instanceof HTMLInputElement &&
      normalizeText(nameInput.value) &&
      normalizeText(functionInput.value)
  );
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
  return String(person.signatures[docType][signer]?.validatedAt || "");
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

function refreshDocumentSignatureCanvases(docType) {
  document.querySelectorAll(`.js-signature-canvas[data-doc-type="${docType}"]`).forEach((canvas) => {
    const person = getCurrentPerson();
    const signer = String(canvas.getAttribute("data-signer") || "");
    const dataUrl = getSignatureValue(person, docType, signer);
    const canvasState = signatureCanvases.get(canvas);
    if (canvasState) {
      canvasState.pendingDataUrl = dataUrl;
    }
    drawSignatureFromDataUrl(canvas, dataUrl);

    const statusNode = canvas
      .closest(".signature-box")
      ?.querySelector(".signature-box__status");
    if (statusNode) {
      const hasSignature = Boolean(dataUrl);
      const validatedAt = formatSignatureTimestamp(getSignatureValidationDate(person, docType, signer));
      statusNode.textContent = hasSignature
        ? validatedAt
          ? `SIGNATURE ENREGISTREE LE ${validatedAt}`
          : "SIGNATURE ENREGISTREE"
        : "AUCUNE SIGNATURE";
      statusNode.classList.toggle("is-signed", hasSignature);
    }
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
      populateRepresentativeSelect(nameInput, person ? getRepresentativeInfo(person, docType).id : "");
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
      if (representative) {
        const currentSignature = person.signatures?.[docType]?.representant || {};
        if (!String(currentSignature.validatedAt || "")) {
          setSignatureValue(
            person,
            docType,
            "representant",
            String(currentSignature.image || ""),
            getCurrentSignatureTimestamp(),
            String(currentSignature.storageRef || ""),
            String(currentSignature.storagePublicUrl || "")
          );
        }
      }
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
    nameInput.oninput = applyRepresentativeSelection;
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
      const person = getCurrentPerson();
      const docType = String(canvas.getAttribute("data-doc-type") || "");
      const signer = String(canvas.getAttribute("data-signer") || "");
      if (!person || !docType || !signer) {
        return;
      }
      const wasFullySigned = isDocumentFullySigned(person, docType);
      const nextValue = stateRef.pendingDataUrl || "";
      const validatedAt = nextValue ? getCurrentSignatureTimestamp() : "";
      let signatureStorageRef = "";
      let signatureStoragePublicUrl = "";
      if (nextValue && isSupabaseConfigured()) {
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
      if (docType === "arrival") {
        renderArrivalDocument(person.id);
      } else if (docType === "exit") {
        renderExitDocument(person.id);
      }
      refreshDocumentSignatureCanvases(docType);
      if (document.body.dataset.page === "mobile-signature" && nextValue) {
        const request = getCurrentMobileSignatureRequest();
        if (
          request &&
          normalizeText(request.docType) === normalizeText(docType) &&
          normalizeMobileSignatureSigner(request.signer || "") === signer
        ) {
          markMobileSignatureRequestSigned(request);
        }
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
      const isMobileSignaturePage = document.body.dataset.page === "mobile-signature";
      const mustAlertAndClose = Boolean(nextValue) && isMobileSignaturePage;
      await saveDataToFile({
        silent: !mustAlertAndClose,
        reloadAfter: !mustAlertAndClose,
        successText: saveText,
        alertText: "DONNEES SUPABASE MISES A JOUR",
        closeAfterAlert: mustAlertAndClose,
      });
      if (document.body.dataset.page === "mobile-signature") {
        renderMobileSignaturePage();
      }
    };

    canvas.addEventListener("pointerdown", (event) => {
      if (document.body.dataset.pdfMode === "true" || !getCurrentPerson()) {
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
      context.moveTo(point.x, point.y);
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
        const person = getCurrentPerson();
        if (!person) {
          return;
        }
        const docType = String(canvas.getAttribute("data-doc-type") || "");
        const signer = String(canvas.getAttribute("data-signer") || "");
        const currentSignerHasSignature = Boolean(getSignatureValue(person, docType, signer));
        const otherSigner = signer === "personnel" ? "representant" : "personnel";
        const otherSignerHasSignature = Boolean(getSignatureValue(person, docType, otherSigner));
        const pendingValue = String(stateRef.pendingDataUrl || "");
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
        const person = getCurrentPerson();
        const docType = String(canvas.getAttribute("data-doc-type") || "");
        const signer = String(canvas.getAttribute("data-signer") || "");
        if (!person || !docType || !signer) {
          return;
        }
        stateRef.pendingDataUrl = "";
        clearSignatureCanvas(canvas);
        setSignatureValue(person, docType, signer, "", "", "", "");
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
  return findMobileSignatureRequestByToken(token);
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

function renderMobileSignaturePage() {
  const request = getCurrentMobileSignatureRequest();
  const person = getCurrentPerson();
  const docType = getCurrentMobileSignatureDocType();
  const signerFromUrl = getCurrentMobileSignatureSigner();
  const signerFromRequest = normalizeMobileSignatureSigner(request?.signer || "");
  const signer = request ? signerFromRequest : signerFromUrl;
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
  const docLabel = normalizeText(docType) === "EXIT" ? "DOCUMENT DE SORTIE" : "DOCUMENT D'ARRIVEE";

  if (canvas) {
    canvas.setAttribute("data-doc-type", normalizeText(docType) === "EXIT" ? "exit" : "arrival");
    canvas.setAttribute("data-signer", signer);
  }
  if (saveButton) {
    saveButton.setAttribute("data-doc-type", normalizeText(docType) === "EXIT" ? "exit" : "arrival");
    saveButton.setAttribute("data-signer", signer);
  }
  if (clearButton) {
    clearButton.setAttribute("data-doc-type", normalizeText(docType) === "EXIT" ? "exit" : "arrival");
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
    const representative = person ? getRepresentativeInfo(person, normalizeText(docType) === "EXIT" ? "exit" : "arrival") : null;
    personNode.textContent =
      signer === "representant"
        ? representative?.nom || "-"
        : person
        ? `${person.nom || ""} ${person.prenom || ""}`.trim() || "-"
        : "-";
  }
  if (dateNode) {
    dateNode.textContent = docType === "exit" ? formatDate(person?.dateSortieReelle || person?.dateSortiePrevue) || "-" : formatDate(person?.dateEntree) || "-";
  }

  const representative = person ? getRepresentativeInfo(person, normalizeText(docType) === "EXIT" ? "exit" : "arrival") : null;
  const representativeReady =
    signer !== "representant" || Boolean(normalizeText(representative?.nom) && normalizeText(representative?.fonction));
  const valid = Boolean(
    person &&
      request &&
      isMobileSignatureRequestValid(request) &&
      normalizeText(request.docType) === normalizeText(docType) &&
      signerFromRequest === signer &&
      representativeReady
  );
  if (panelNode) {
    panelNode.hidden = !valid;
  }
  if (saveButton instanceof HTMLButtonElement) {
    saveButton.disabled = !valid;
    saveButton.classList.toggle("is-disabled", !valid);
  }
  if (clearButton instanceof HTMLButtonElement) {
    clearButton.disabled = !valid;
    clearButton.classList.toggle("is-disabled", !valid);
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
    renderDocumentCostsTable(mobileCostsHead, mobileCostsBody);
  }

  fillMobileSignatureShareLink(valid ? request : null);
}

function renderPersonPicker() {
  const picker = document.getElementById("person-picker-search");
  const pickerList = document.getElementById("person-picker-list");
  const suggestionBox = document.getElementById("person-picker-suggestions");
  if (!picker || !state.data?.personnes) {
    return;
  }

  const currentPersonId = getCurrentPersonId();
  picker.setAttribute("autocomplete", "off");
  picker.setAttribute("autocorrect", "off");
  picker.setAttribute("autocapitalize", "off");
  picker.setAttribute("spellcheck", "false");
  picker.removeAttribute("list");
  const selectedPerson = currentPersonId
    ? state.data.personnes.find((person) => String(person?.id || "") === String(currentPersonId || "")) || null
    : null;

  const pickerIsFocused = document.activeElement === picker;
  if (!pickerIsFocused) {
    picker.value = selectedPerson ? getPersonPickerLabel(selectedPerson) : "";
  }

  const page = document.body.dataset.page || "";
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
      renderArrivalDocument(personId);
      refreshDocumentSignatureCanvases("arrival");
      updateSortableHeaders("arrivalEffects");
      syncMobileSignaturePolling();
    } else if (page === "exit-document") {
      renderExitDocument(personId);
      refreshDocumentSignatureCanvases("exit");
      updateSortableHeaders("exitEffects");
      syncMobileSignaturePolling();
    } else {
      renderPage();
    }
    renderDirtyState();
    picker.blur();
  };

  const applyFullResetFromPicker = () => {
    state.filters = { ...DEFAULT_FILTERS };
    saveNavigationContext({ filters: state.filters, personId: "" });
    if (useDirectNavigation) {
      hideSuggestions();
      applyDocumentNavigation("");
      return true;
    }
    setCurrentPersonId("", "replace");
    renderPage();
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
    renderPage();
    picker.blur();
    return true;
  };

  picker.oninput = () => {
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
  const hasNonRendu = (person.effetsConfies || []).some(
    (effect) => normalizeText(getEffectStatus(person, effect)) === "NON RENDU"
  );
  return hasOverdueExitAlert && hasNonRendu;
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
  if (filters.typeEffet && !effects.some((effect) => normalizeText(effect.typeEffet) === filters.typeEffet)) return false;
  if (filters.statutObjet && !effects.some((effect) => effect.statutAffiche === filters.statutObjet)) return false;
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

  let inPostCount = 0;
  let totalEffectsCount = 0;
  let missingEffectsCount = 0;

  persons.forEach((person) => {
    const allEffects = person.effetsConfies || [];
    if (getDossierStatus(person) === "EN POSTE") {
      inPostCount += 1;
    }
    allEffects.forEach((effect) => {
      totalEffectsCount += 1;
      if (getEffectStatus(person, effect) === "NON RENDU") {
        missingEffectsCount += 1;
      }
    });
  });

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

  if (body) {
    const sortedPersons = sortPersonsForOverview(persons);
    const rowsHtml = buildOverviewRows(sortedPersons);
    renderTableRowsProgressively(body, [rowsHtml], buildEmptyTableRow("overview-table-body", "AUCUNE DONNEE A AFFICHER", 13), 1);
    bindPersonRowActions();
    bindDeletePersonButtons();
    updateSortableHeaders("overviewPersons");
  }

  if (alertsSection && alertsList) {
    const alerts = persons
      .filter((person) => hasOverdueExit(person))
      .map((person) => {
        const alertMeta = getOverdueExitAlertMeta(person);
        return {
          id: person.id,
          nom: person.nom,
          prenom: person.prenom,
          message: alertMeta.message,
          type: alertMeta.type,
        };
      });

    alertsSection.hidden = alerts.length === 0;
    alertsList.innerHTML = alerts
      .map(
        (alert) => `<button type="button" class="overview-alert-item overview-alert-item--${alert.type || "dateSortiePrevue"} js-open-person-alert" data-person-id="${alert.id}">
          <span class="overview-alert-item__icon overview-alert-item__icon--${alert.type || "dateSortiePrevue"}" aria-hidden="true">${alert.type === "dateSortieReelle" ? "✕" : "!"}</span>
          <span class="overview-alert-item__content">
            <strong>${escapeHtml(`${alert.nom} ${alert.prenom}`.trim())}</strong>
            <span>${escapeHtml(alert.message)}</span>
          </span>
        </button>`
      )
      .join("");

    alertsList.querySelectorAll(".js-open-person-alert").forEach((button) => {
      button.onclick = () => {
        const personId = button.getAttribute("data-person-id") || "";
        if (personId) {
          openPersonSheet(personId);
        }
      };
    });
  }
}

function buildOverviewRows(persons) {
  if (!persons.length) {
    return buildEmptyTableRow("overview-table-body", "AUCUNE DONNEE A AFFICHER", 13);
  }
  return persons
    .map((person) => {
      const currentEffects = getCurrentAssignedEffects(person);
      const totalEffects = currentEffects.length;
      const nonRendus = currentEffects.filter(
        (effect) => getEffectStatus(person, effect) === "NON RENDU"
      ).length;
      const movementMap = getArrivalComplementMovementMap(person, person.effetsConfies || []);
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
        <td>${person.typePersonnel || ""}</td>
        <td>${person.typeContrat || ""}</td>
        <td>${formatDate(person.dateEntree)}</td>
        <td>${formatDate(person.dateSortiePrevue)}</td>
        <td>${formatDate(person.dateSortieReelle)}</td>
        <td>${getDossierStatusCellMarkup(getDossierStatus(person))}</td>
        <td>${totalEffects}</td>
        <td>${nonRendus > 0 ? '<span class="row-alert-dot" aria-hidden="true"></span>' : ""}${nonRendus}</td>
        <td>${movementMarkup || "-"}</td>
        <td>
          <a class="table-link js-open-person-link" data-person-id="${person.id}" href="fiche-personne.html?personId=${person.id}">VOIR</a>
          <button type="button" class="table-link js-delete-person" data-person-id="${person.id}">SUPPRIMER</button>
        </td>
      </tr>`;
    })
    .join("");
}

function renderGlobalTable(persons) {
  const body = document.getElementById("global-table-body");
  if (!body) {
    return;
  }

  if (!persons.length) {
    body.innerHTML = buildEmptyTableRow(body, "AUCUNE DONNEE A AFFICHER", 13);
    return;
  }

  const rowsHtml = persons
    .map((person) => {
      const currentEffects = getCurrentAssignedEffects(person);
      const totalEffects = currentEffects.length;
      const nonRendus = currentEffects.filter(
        (effect) => getEffectStatus(person, effect) === "NON RENDU"
      ).length;
      const movementMap = getArrivalComplementMovementMap(person, person.effetsConfies || []);
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
        <td>${movementMarkup || "-"}</td>
        <td>
          <a class="table-link js-open-person-link" data-person-id="${person.id}" href="fiche-personne.html?personId=${person.id}">VOIR</a>
          <button type="button" class="table-link js-delete-person" data-person-id="${person.id}">SUPPRIMER</button>
        </td>
      </tr>`;
    })
    ;

  renderTableRowsProgressively(body, rowsHtml, buildEmptyTableRow(body, "AUCUNE DONNEE A AFFICHER", 13), 24);

  bindPersonRowActions();
  bindDeletePersonButtons();
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

    body.dataset.bound = "true";
  });
}

function bindOpenPersonLinks() {
  document.querySelectorAll(".js-open-person-link").forEach((link) => {
    if (link.dataset.bound === "true") {
      return;
    }
    link.addEventListener("click", () => {
      const personId = link.getAttribute("data-person-id") || "";
      if (personId) {
        setCurrentPersonId(personId, "replace");
      }
    });
    link.dataset.bound = "true";
  });
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
  if (normalizedType === "CLE" || normalizedType === "CLE CES") {
    return "cle";
  }
  return "total";
}

function getSheetEffectTypeIconSvg(typeEffet) {
  const variant = getSheetEffectTypeIconVariant(typeEffet);
  if (variant === "cle-ces") {
    return '<img src="https://dphrvdhqhgycmllietuk.supabase.co/storage/v1/object/public/ui-assets/sidebar/icone-cle-ces.png" alt="" loading="lazy">';
  }
  if (variant === "badge") {
    return '<img src="https://dphrvdhqhgycmllietuk.supabase.co/storage/v1/object/public/ui-assets/sidebar/icone-badge.png" alt="" loading="lazy">';
  }
  if (variant === "telecommande") {
    return '<img src="https://dphrvdhqhgycmllietuk.supabase.co/storage/v1/object/public/ui-assets/sidebar/icone-telecommande.png" alt="" loading="lazy">';
  }
  if (variant === "carte") {
    return '<img src="https://dphrvdhqhgycmllietuk.supabase.co/storage/v1/object/public/ui-assets/sidebar/icone-carte.png" alt="" loading="lazy">';
  }
  if (variant === "cle") {
    return '<img src="https://dphrvdhqhgycmllietuk.supabase.co/storage/v1/object/public/ui-assets/sidebar/icone-cle.png" alt="" loading="lazy">';
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
  if (!person) {
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
    return;
  }

  state.currentSheetPersonId = String(person.id || "");
  nameNode.textContent = `${person.nom} ${person.prenom}`;
  metaNode.innerHTML = [
    getPersonSiteMarkup(person),
    escapeHtml(person.typePersonnel || ""),
    escapeHtml(person.typeContrat || ""),
  ]
    .filter(Boolean)
    .join(' <span class="meta-separator">|</span> ');
  const overdueMessage = getOverdueExitMessage(person);
  alertNode.hidden = !overdueMessage;
  alertNode.textContent = overdueMessage;
  applySheetPersonStatus(statusNode, getDossierStatus(person));
  updateSheetDocumentButtons(person);
  fillSheetForm(person);

  const effects = person.effetsConfies || [];
  const currentEffects = getCurrentAssignedEffects(person);
  const sortedEffects = sortEffectsForTable(person, currentEffects, "sheetEffects");
  const rowFlash =
    state.effectRowFlash && String(state.effectRowFlash.personId || "") === String(person.id || "")
      ? state.effectRowFlash
      : null;
  const movementMap = getArrivalComplementMovementMap(person, currentEffects);
  const returned = effects.filter((effect) => getEffectStatus(person, effect) === "RESTITUE").length;
  const missing = currentEffects.filter((effect) => getEffectStatus(person, effect) === "NON RENDU").length;
  const totalCost = currentEffects.reduce((sum, effect) => sum + getEffectReplacementCost(person, effect), 0);
  const totalEffectsUnitValue = currentEffects.reduce((sum, effect) => sum + getEffectUnitValue(effect), 0);

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
  if (
    state.effectTableFlash &&
    String(state.effectTableFlash.personId || "") === String(person.id || "") &&
    String(state.effectTableFlash.kind || "") === "delete"
  ) {
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

  const requestedEditEffectId = consumeRequestedEditEffectId();
  if (requestedEditEffectId && currentEffects.some((effect) => String(effect.id || "") === requestedEditEffectId)) {
    startEditEffect(person.id, requestedEditEffectId);
  }
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
        needsExpectedExitDate && !dateSortiePrevueValue
      );
    }
    form.elements.sheetNom?.closest(".field")?.classList.toggle("field--missing", !nom);
    form.elements.sheetPrenom?.closest(".field")?.classList.toggle("field--missing", !prenom);
    form.elements.sheetFonction?.closest(".field")?.classList.toggle("field--missing", !fonction);
    form.elements.sheetTypePersonnel?.closest(".field")?.classList.toggle("field--missing", !typePersonnel);
    form.elements.sheetTypeContrat?.closest(".field")?.classList.toggle("field--missing", !typeContrat);
    form.elements.sheetDateEntree?.closest(".field")?.classList.toggle("field--missing", !dateEntree);
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

  if (!person) {
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
    syncDocumentMobileSignatureLink("arrival", "", "personnel");
    syncDocumentMobileSignatureLink("arrival", "", "representant");
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
  const sortedEffects = sortEffectsForTable(person, effectsForDisplay, "arrivalEffects");
  const totalValue = activeEffects.reduce((sum, effect) => sum + getEffectUnitValue(effect), 0);

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
  const arrivalRepresentative = getRepresentativeInfo(person, "arrival");
  populateRepresentativeSelect(representantNameInput, arrivalRepresentative.id);
  representantFunctionInput.value = arrivalRepresentative.fonction;
  representantNameNode.textContent = arrivalRepresentative.nom || "-";
  representantFunctionNode.textContent = arrivalRepresentative.fonction || "-";
  updateRepresentativeSignatureActionState("arrival");
  signaturePersonDateNode.textContent =
    formatSignatureTimestamp(getSignatureValidationDate(person, "arrival", "personnel")) || "-";
  signatureRepresentantDateNode.textContent =
    formatSignatureTimestamp(getSignatureValidationDate(person, "arrival", "representant")) || "-";

  const complementMovements = isComplement ? fallbackMovements : new Map();

  body.innerHTML = sortedEffects.length
    ? `${sortedEffects
        .map(
          (effect) => {
            const movement = getEffectMovementLabel(
              person,
              effect,
              isComplement ? complementMovements : null
            );
            const movementBadge = movement
              ? `<span class="movement-badge movement-badge--${getMovementBadgeVariant(movement)}">${movement}</span>`
              : "";
            const rowClass = movement
              ? ` class="arrival-effect-row arrival-effect-row--${getMovementRowVariant(movement)}"`
              : "";
            const actionCell = !isPdfMode
              ? `<span class="document-effect-actions">
                  <button type="button" class="table-link js-doc-edit-effect" data-person-id="${escapeHtml(person.id || "")}" data-effect-id="${escapeHtml(effect.id || "")}">MODIFIER</button>
                  <button type="button" class="table-link js-doc-delete-effect" data-person-id="${escapeHtml(person.id || "")}" data-effect-id="${escapeHtml(effect.id || "")}">SUPPRIMER</button>
                </span>`
              : "-";
            return `<tr${rowClass}>
            <td>${effect.typeEffet || ""}</td>
            <td>${getEffectDisplayDesignation(effect) || "-"}</td>
            <td class="movement-cell">${movementBadge}</td>
            <td>${effect.numeroIdentification || "-"}</td>
            <td>${formatDate(effect.dateRemise) || "-"}</td>
            <td>${formatAmountWithEuro(getEffectUnitValue(effect))}</td>
            <td class="document-effects-action-col">${actionCell}</td>
          </tr>`;
          }
        )
        .join("")}
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
  syncDocumentMobileSignatureLink("arrival", person.id, "personnel");
  syncDocumentMobileSignatureLink("arrival", person.id, "representant");
  applyRequestedPdfFocus();
}

function renderArrivalCostsTable(headNode, bodyNode) {
  renderDocumentCostsTable(headNode, bodyNode);
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

function renderDocumentCostsTable(headNode, bodyNode) {
  const causes = getArrivalCostCauses();
  const effectTypes = getArrivalCostTypes();

  headNode.innerHTML = `<tr>
    <th>TYPE D'EFFET</th>
    ${causes.map((cause) => `<th>${escapeHtml(cause)}</th>`).join("")}
  </tr>`;

  if (!effectTypes.length) {
    bodyNode.innerHTML = buildEmptyTableRow(bodyNode, "AUCUN COUT A AFFICHER", causes.length + 1);
    return;
  }

  bodyNode.innerHTML = effectTypes
    .map(
      (typeEffet) => `<tr>
        <td>${escapeHtml(typeEffet)}</td>
        ${causes
          .map((cause) => `<td>${formatAmountWithEuro(getReplacementCostValue(typeEffet, cause, getArrivalCostDesignation(typeEffet)))}</td>`)
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

  if (!person) {
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
    body.innerHTML = buildEmptyTableRow(body, "AUCUN EFFET A AFFICHER", 10);
    totalEffectsNode.textContent = "0";
    totalReturnedNode.textContent = "0";
    totalChargeableNode.textContent = "0";
    totalValueNode.textContent = "0,00 €";
    signatureNameNode.textContent = "-";
    signaturePersonDateNode.textContent = "-";
    signatureRepresentantDateNode.textContent = "-";
    populateRepresentativeSelect(representantNameInput, "");
    representantFunctionInput.value = "";
    representantNameNode.textContent = "-";
    representantFunctionNode.textContent = "-";
    renderDocumentCostsTable(costsHead, costsBody);
    updateRepresentativeSignatureActionState("exit");
    syncDocumentMobileSignatureLink("exit", "", "personnel");
    syncDocumentMobileSignatureLink("exit", "", "representant");
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
  const isPdfMode = isPdfRenderMode();
  const sortedEffects = sortEffectsForTable(person, effects, "exitEffects");
  const totalReturned = effects.filter((effect) => getEffectStatus(person, effect) === "RESTITUE").length;
  const chargeableEffects = effects.filter((effect) => isEffectChargeable(person, effect));
  const totalValue = chargeableEffects.reduce((sum, effect) => sum + getEffectReplacementCost(person, effect), 0);
  const todayIso = getTodayIsoDate();

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
  const exitRepresentative = getRepresentativeInfo(person, "exit");
  populateRepresentativeSelect(representantNameInput, exitRepresentative.id);
  representantFunctionInput.value = exitRepresentative.fonction;
  representantNameNode.textContent = exitRepresentative.nom || "-";
  representantFunctionNode.textContent = exitRepresentative.fonction || "-";
  updateRepresentativeSignatureActionState("exit");
  signaturePersonDateNode.textContent =
    formatSignatureTimestamp(getSignatureValidationDate(person, "exit", "personnel")) || "-";
  signatureRepresentantDateNode.textContent =
    formatSignatureTimestamp(getSignatureValidationDate(person, "exit", "representant")) || "-";

  body.innerHTML = sortedEffects.length
    ? `${sortedEffects
        .map(
          (effect) => {
            const movement = getEffectMovementLabel(person, effect);
            const movementBadge = movement
              ? `<span class="movement-badge movement-badge--${getMovementBadgeVariant(movement)}">${movement}</span>`
              : "";
            const rawStatus = getEffectStatus(person, effect);
            const currentStatus = normalizeText(rawStatus);
            const statusLabel = currentStatus === "RESTITUE" ? "RENDU" : rawStatus;
            const retourDateIso = normalizeDateString(effect.dateRetour || "");
            const canToggleReturnToday =
              !["PERDU", "HS", "VOL"].includes(currentStatus) &&
              (!retourDateIso || retourDateIso === todayIso);
            const actionCell = !isPdfMode
              ? `<span class="document-effect-actions">
                  <button type="button" class="table-link js-doc-edit-effect" data-person-id="${escapeHtml(person.id || "")}" data-effect-id="${escapeHtml(effect.id || "")}">MODIFIER</button>
                  <button type="button" class="table-link js-doc-delete-effect" data-person-id="${escapeHtml(person.id || "")}" data-effect-id="${escapeHtml(effect.id || "")}">SUPPRIMER</button>
                </span>`
              : "-";
            return `<tr>
            <td>${effect.typeEffet || ""}</td>
            <td>${getEffectDisplayDesignation(effect)}</td>
            <td>${effect.numeroIdentification || ""}</td>
            <td>${formatDate(effect.dateRemise)}</td>
            <td>${formatDate(effect.dateRetour)}</td>
            <td>${statusLabel}</td>
            <td class="movement-cell">${movementBadge}</td>
            <td class="movement-cell">${
              !isPdfMode && canToggleReturnToday
                ? `<label class="return-today-toggle"><input type="checkbox" class="js-exit-return-today" data-effect-id="${escapeHtml(effect.id || "")}" ${retourDateIso === todayIso ? "checked" : ""} /><span>RENDU</span></label>`
                : "-"
            }</td>
            <td>${formatAmountWithEuro(getEffectReplacementCost(person, effect))}</td>
            <td class="document-effects-action-col">${actionCell}</td>
          </tr>`;
          }
        )
        .join("")}
        <tr class="table-total-row">
          <td colspan="${isPdfMode ? "7" : "9"}">TOTAL FACTURABLE DES EFFETS</td>
          <td>${formatAmountWithEuro(totalValue)}</td>
        </tr>`
    : buildEmptyTableRow(body, "AUCUN EFFET A AFFICHER", 10);

  totalEffectsNode.textContent = String(effects.length);
  totalReturnedNode.textContent = String(totalReturned);
  totalChargeableNode.textContent = String(chargeableEffects.length);
  totalValueNode.textContent = formatAmountWithEuro(totalValue);
  renderDocumentCostsTable(costsHead, costsBody);
  bindExitReturnTodayToggles();
  bindDocumentEffectActions();
  updateSortableHeaders("exitEffects");
  syncDocumentMobileSignatureLink("exit", person.id, "personnel");
  syncDocumentMobileSignatureLink("exit", person.id, "representant");
  applyRequestedPdfFocus();
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
    effect.dateRetour = target.checked ? getTodayIsoDate() : "";
    markDirty();
    renderPage();
    renderExitDocument(person.id);
    renderPersonSheet(person.id);
    showActionStatus("update", target.checked ? "EFFET MARQUE RENDU CE JOUR" : "RETOUR DU JOUR ANNULE");
  });
  body.dataset.returnBound = "true";
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
      if (normalizeText(reference.typeEffet) !== getReferenceCatalogType(normalizedTypeEffet)) {
        return false;
      }
      if (normalizedTypeEffet === "CLE CES" && !isCesKeyDesignation(reference.designation)) {
        return false;
      }
      if (normalizedTypeEffet === "CLE" && isCesKeyDesignation(reference.designation)) {
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
      if (normalizedReferenceSite && !referenceHasSite(reference, normalizedReferenceSite)) {
        return false;
      }
      if (
        normalizedTypeEffet &&
        normalizeText(reference.typeEffet) !== getReferenceCatalogType(normalizedTypeEffet)
      ) {
        return false;
      }
      if (normalizedTypeEffet === "CLE CES" && !isCesKeyDesignation(reference.designation)) {
        return false;
      }
      if (normalizedTypeEffet === "CLE" && isCesKeyDesignation(reference.designation)) {
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
      reference.typeEffet === getReferenceCatalogType(typeEffet) &&
      reference.designation === designation
  );

  if (!exists) {
    state.data.listes.referencesEffets.push({
      id: getNextId("REF", state.data.listes.referencesEffets),
      site,
      typeEffet: getReferenceCatalogType(typeEffet),
      designation,
    });
    sortReferenceEffects();
  }
}

function findReferenceById(referenceId) {
  return state.data?.listes?.referencesEffets?.find((reference) => reference.id === referenceId) || null;
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
  if (person?.dateSortiePrevue && !person?.dateSortieReelle && isPastDate(person.dateSortiePrevue)) {
    return {
      type: "dateSortiePrevue",
      message: `ALERTE : DATE DE SORTIE PREVUE DEPASSEE (${formatDate(person.dateSortiePrevue)})`,
    };
  }
  const hasNonRenduEffects = (person?.effetsConfies || []).some(
    (effect) => normalizeText(getEffectStatus(person, effect)) === "NON RENDU"
  );
  if (person?.dateSortieReelle && isPastDate(person.dateSortieReelle) && hasNonRenduEffects) {
    return {
      type: "dateSortieReelle",
      message: `ALERTE : DATE DE SORTIE REELLE DEPASSEE (${formatDate(person.dateSortieReelle)})`,
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
  if (person?.dateSortieReelle && String(person.dateSortieReelle) <= today) {
    return true;
  }
  if (person?.dateSortiePrevue && isPastDate(person.dateSortiePrevue)) {
    return true;
  }
  return false;
}

function getEffectStatus(person, effect) {
  if (effect.dateRetour) return "RESTITUE";

  const manualStatus = normalizeText(effect.statutManuel);
  if (manualStatus === "CASSE") return "HS";
  if (["PERDU", "HS", "VOL"].includes(manualStatus)) return manualStatus;
  if (isExitDue(person)) return "NON RENDU";
  return manualStatus || "ACTIF";
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

function isCurrentAssignedEffect(person, effect) {
  const status = normalizeText(getEffectStatus(person, effect));
  return !["RESTITUE", "PERDU", "HS", "DETRUIT", "VOL"].includes(status);
}

function getCurrentAssignedEffects(person) {
  return (person?.effetsConfies || []).filter((effect) => isCurrentAssignedEffect(person, effect));
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
  state.referenceRenderContext = buildReferenceRenderContext();
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
  renderReferenceCounts();
  renderMobileSignatureSettings();
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
  const values = (state.data?.listes?.typesEffets || []).slice().sort((a, b) => normalizeText(a).localeCompare(normalizeText(b), "fr"));
  syncSelectOptions(typeSelect, values, "SELECTIONNER");
  const sites = (state.data?.listes?.sites || [])
    .filter((site) => normalizeText(site) !== ALL_SITES_VALUE)
    .slice()
    .sort((a, b) => normalizeText(a).localeCompare(normalizeText(b), "fr"));
  syncSelectOptions(siteSelect, sites, "SELECTIONNER");
  syncSelectOptions(reasonSelect, getStockReasonOptions(), "SELECTIONNER");
  updateStockDesignationOptions();
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
  const references = (state.data?.listes?.referencesEffets || [])
    .filter((reference) => isReferenceEffectActive(reference))
    .filter((reference) => !typeEffet || normalizeText(reference.typeEffet) === typeEffet)
    .filter((reference) => !site || referenceHasSite(reference, site))
    .sort((a, b) => normalizeText(a.designation).localeCompare(normalizeText(b.designation), "fr"));

  const currentValue = designationSelect.value;
  designationSelect.innerHTML = [`<option value="">SELECTIONNER</option>`]
    .concat(
      references.map(
        (reference) =>
          `<option value="${escapeHtml(String(reference.id || ""))}">${escapeHtml(reference.designation || "-")}</option>`
      )
    )
    .join("");
  if (currentValue && references.some((entry) => String(entry.id || "") === currentValue)) {
    designationSelect.value = currentValue;
  }
  designationSelect.disabled = references.length === 0;
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
    if (filterTypeEffet && normalizeText(reference.typeEffet) !== filterTypeEffet) {
      return false;
    }
    if (filterSearch) {
      const text = [reference.site, reference.typeEffet, reference.designation]
        .map(normalizeText)
        .join(" ");
      if (!text.includes(filterSearch)) {
        return false;
      }
    }
    return true;
  });
  const sortedReferences = sortReferencesForTable(references, "referenceEffects", renderContext);
  if (!references.length) {
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
    });

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
      const leftLabel = `${normalizeText(left.typeEffet)} ${normalizeText(left.cause)}`;
      const rightLabel = `${normalizeText(right.typeEffet)} ${normalizeText(right.cause)}`;
      return leftLabel.localeCompare(rightLabel, "fr");
    });

  if (!entries.length) {
    body.innerHTML = buildEmptyTableRow(body, "AUCUN COUT", 4);
    return;
  }

  const rowsHtml = entries
    .map((entry) => {
      const key = getReplacementCostKey(entry.typeEffet, entry.cause);
      return `<tr class="js-replacement-cost-row" data-cost-key="${key}">
        <td>${escapeHtml(entry.typeEffet)}</td>
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

  getAllEffects(state.data?.personnes || []).forEach(({ person, effect }) => {
    const typeEffet = normalizeText(effect?.typeEffet || "");
    const site = normalizeText(getEffectDisplaySite(effect) || getPersonSiteLabel(person) || "SANS SITE");
    const designation = normalizeText(getEffectDisplayDesignation(effect) || effect?.designation || typeEffet || "SANS DESIGNATION");
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
    const typeEffet = normalizeText(entry?.typeEffet || "");
    const site = normalizeText(entry?.site || "SANS SITE");
    const designation = normalizeText(entry?.designation || "");
    if (!typeEffet || !designation) {
      return;
    }
    const row = ensureRow(typeEffet, site, designation);
    row.manuelDelta += getStockMovementSignedQuantity(entry);
  });

  return Array.from(rowsByKey.values())
    .map((row) => {
      const sortiesSansRetour = row.nonRendus + row.perdus + row.voles + row.hs + row.detruits;
      return {
        ...row,
        stockCourant: row.manuelDelta + row.rendus - sortiesSansRetour,
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
    return;
  }
  const entries = (state.data?.stocksEffetsManuels || [])
    .slice()
    .sort((a, b) => String(b.date || "").localeCompare(String(a.date || ""), "fr"));
  if (!entries.length) {
    body.innerHTML = buildEmptyTableRow(body, "AUCUN MOUVEMENT MANUEL", 8);
    return;
  }
  const rowsHtml = entries.map((entry) => {
    const signedQty = getStockMovementSignedQuantity(entry);
    const qtyLabel = signedQty > 0 ? `+${signedQty}` : `${signedQty}`;
    return `<tr class="js-stock-movement-row" data-stock-movement-id="${escapeHtml(entry.id)}">
      <td>${escapeHtml(formatDate(entry.date) || entry.date || "-")}</td>
      <td>${escapeHtml(entry.site || "-")}</td>
      <td>${escapeHtml(entry.typeEffet || "-")}</td>
      <td>${escapeHtml(entry.designation || "-")}</td>
      <td>${escapeHtml(entry.action || "-")}</td>
      <td>${escapeHtml(qtyLabel)}</td>
      <td>${escapeHtml(entry.motif || "-")}</td>
      <td><button type="button" class="table-link js-delete-stock-movement" data-stock-movement-id="${escapeHtml(entry.id)}">SUPPRIMER</button></td>
    </tr>`;
  });
  renderTableRowsProgressively(body, rowsHtml, buildEmptyTableRow(body, "AUCUN MOUVEMENT MANUEL", 8), 24);
  bindStockMovementActions();
}

function renderStockSummaryTable() {
  const body = document.getElementById("stock-summary-table-body");
  if (!body) {
    return;
  }
  const rows = getStockSummaryRows();
  if (!rows.length) {
    body.innerHTML = buildEmptyTableRow(body, "AUCUN STOCK CALCULE", 12);
    return;
  }
  const rowsHtml = rows.map((row) => {
    const manualDeltaLabel = row.manuelDelta > 0 ? `+${row.manuelDelta}` : String(row.manuelDelta);
    return `<tr>
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
      <td><strong>${row.stockCourant}</strong></td>
    </tr>`;
  });
  renderTableRowsProgressively(body, rowsHtml, buildEmptyTableRow(body, "AUCUN STOCK CALCULE", 12), 24);
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

  Object.entries(mapping).forEach(([id, value]) => {
    const node = document.getElementById(id);
    if (node) {
      setKpiCountAnimated(node, Number(value) || 0);
    }
  });
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
  renderPage();
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
    (item) => getReplacementCostKey(item.typeEffet, item.cause) === costKey
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
    (item) => getReplacementCostKey(item.typeEffet, item.cause) === costKey
  );
  if (!entry) {
    return;
  }
  if (!window.confirm(`SUPPRIMER DEFINITIVEMENT LE COUT : ${entry.typeEffet} / ${entry.cause} ?`)) {
    return;
  }
  pushUndoSnapshot("SUPPRESSION COUT");
  state.data.listes.coutsRemplacement = state.data.listes.coutsRemplacement.filter(
    (item) => getReplacementCostKey(item.typeEffet, item.cause) !== costKey
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
  renderPage();
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
  renderPage();
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
    const usage = getReferenceEffectUsage(referenceId);
    if (usage > 0) {
      const forceDisable = window.confirm(
        "REFERENCE DEJA UTILISEE DANS DES DOSSIERS. DESACTIVER QUAND MEME ?"
      );
      if (!forceDisable) {
        return;
      }
    }
  }

  pushUndoSnapshot("SUPPRESSION REFERENCE");
  reference.active = nextActive;
  if (state.editingReferenceId === referenceId) {
    state.editingReferenceId = "";
    resetReferenceEffectForm();
  }
  markDirty();
  renderPage();
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
  const usage = getReferenceEffectUsage(referenceId);
  if (usage > 0) {
    showDataStatus("SUPPRESSION DEFINITIVE BLOQUEE - REFERENCE DEJA UTILISEE");
    window.alert("SUPPRESSION DEFINITIVE IMPOSSIBLE : REFERENCE DEJA UTILISEE");
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
  if (state.editingReferenceId === referenceId) {
    state.editingReferenceId = "";
    resetReferenceEffectForm();
  }
  markDirty();
  renderPage();
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
    reloadAfter = true,
    successText = "data.json MIS A JOUR",
    alertText = "",
    closeAfterAlert = false,
    promptDownload = !silent,
  } = resolvedOptions;

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
    const mode = getDataBackendMode();
    const shouldMirrorToSupabase = mode === "LOCAL_API" && isSupabaseConfigured();
    let saveStatusText = successText;
    let saveAlertText = alertText || "data.json A ETE MIS A JOUR";
    let saveSource = "LOCAL";
    if (mode === "SUPABASE") {
      await saveSupabaseStateData(state.data);
      saveStatusText = "DONNEES SUPABASE SAUVEGARDEES";
      saveAlertText = alertText || "DONNEES SUPABASE MISES A JOUR";
      saveSource = "SUPABASE";
    } else if (mode === "LOCAL_API") {
      const response = await fetch("/api/save", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(state.data),
      });

      if (!response.ok) {
        throw new Error("Sauvegarde locale impossible");
      }
      if (shouldMirrorToSupabase) {
        await saveSupabaseStateData(state.data);
        saveStatusText = "DONNEES LOCALES ET SUPABASE SAUVEGARDEES";
        saveAlertText = alertText || "DONNEES LOCALES ET SUPABASE MISES A JOUR";
        saveSource = "LOCAL + SUPABASE";
      }
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
    const saveConfirmation = `SAUVEGARDEE LE ${formatCurrentUiTimestamp()} - SOURCE: ${saveSource}`;
    showDataStatus(saveConfirmation);
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
    if (reloadAfter) {
      await reloadData(mode === "SUPABASE" ? "RELECTURE DES DONNEES SUPABASE..." : "RELECTURE DE data.json...");
    }
  } catch (error) {
    console.error(error);
    if (isSaveConflictError(error)) {
      if (getDataBackendMode() === "SUPABASE" || isSupabaseConfigured()) {
        try {
          await reloadData("CONFLIT DETECTE - RECHARGEMENT DES DONNEES DISTANTES...");
        } catch (refreshError) {
          console.error(refreshError);
        }
      }
      showDataStatus("CONFLIT DE SAUVEGARDE - RECHARGER PUIS REESSAYER");
      if (!silent) {
        window.alert(error.message);
      }
      return;
    }
    showDataStatus("SAUVEGARDE IMPOSSIBLE");
    if (!silent) {
      window.alert("SAUVEGARDE IMPOSSIBLE");
    }
  }

}

function resetUiWithoutData() {
  const targets = [
    { id: "overview-table-body", colspan: 13 },
    { id: "global-table-body", colspan: 13 },
    { id: "sheet-effects-body", colspan: 11 },
    { id: "reference-sites-body", colspan: 2 },
    { id: "reference-typesPersonnel-body", colspan: 2 },
    { id: "reference-typesContrats-body", colspan: 2 },
    { id: "reference-fonctions-body", colspan: 2 },
    { id: "reference-typesEffets-body", colspan: 2 },
    { id: "reference-causesRemplacement-body", colspan: 2 },
    { id: "reference-effects-table-body", colspan: 5 },
    { id: "replacement-costs-body", colspan: 4 },
    { id: "stock-movements-table-body", colspan: 8 },
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
  ].forEach((id) => {
    const node = document.getElementById(id);
    if (node) node.textContent = "0";
  });

  renderPersonSheet("");
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
  renderPage();
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
    renderPage();
  });

  window.__dashboardHistoryBound = true;
}

function markDirty() {
  state.isDirty = true;
  state.saveButtonLatchedDirty = true;
  saveWorkingData();
  renderDirtyState();
  scheduleBackgroundAutoSave();
}

function renderDirtyState() {
  const node = document.getElementById("dirty-status");
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
}

loadData();
