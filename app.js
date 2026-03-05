const state = {
  data: null,
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
  tableSorts: {
    sheetEffects: { key: "typeEffet", dir: "asc" },
    arrivalEffects: { key: "typeEffet", dir: "asc" },
    exitEffects: { key: "typeEffet", dir: "asc" },
    overviewPersons: { key: "nom", dir: "asc" },
  },
  filters: {
    search: "",
    site: "",
    typePersonnel: "",
    typeContrat: "",
    statutDossier: "",
    statutObjet: "",
    typeEffet: "",
  },
  referenceRenderContext: null,
};

const WORKING_DATA_KEY = "dashboard-working-data";
const LEGACY_CONTRACT_TYPES = ["CDI", "CDD", "INTERIMAIRE"];
const MAX_UNDO_STACK = 30;
const ALL_SITES_VALUE = "TOUS SITES";
const CHARGED_REPLACEMENT_CAUSES = ["PERTE", "CASSE", "NON RENDU", "VOL"];
const MOBILE_SIGNATURE_REQUEST_TTL_MS = 10 * 60 * 1000;
const SUPABASE_PROJECT_URL = "https://dphrvdhqhgycmllietuk.supabase.co";
const SUPABASE_PUBLISHABLE_KEY = "sb_publishable_2wYXnIDj4-c8daQZW8D5hA_2Py6k7z6";
const SUPABASE_APP_STATE_TABLE = "app_state";
const SUPABASE_APP_STATE_ID = "main";
let pdfModalCleanupBound = false;
const signatureCanvases = new WeakMap();

redirectToLocalServerIfNeeded();
applyPdfModeFromQuery();

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

function getSupabaseHeaders(extra = {}) {
  const key = String(SUPABASE_PUBLISHABLE_KEY || "").trim();
  return {
    apikey: key,
    Authorization: `Bearer ${key}`,
    ...extra,
  };
}

async function fetchSupabaseStateData() {
  const endpoint = `${getSupabaseRestEndpoint()}?id=eq.${encodeURIComponent(
    SUPABASE_APP_STATE_ID
  )}&select=payload&limit=1`;
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
  return rows[0].payload;
}

async function saveSupabaseStateData(payload) {
  const response = await fetch(getSupabaseRestEndpoint(), {
    method: "POST",
    headers: getSupabaseHeaders({
      "Content-Type": "application/json",
      Prefer: "resolution=merge-duplicates,return=representation",
    }),
    body: JSON.stringify([
      {
        id: SUPABASE_APP_STATE_ID,
        payload,
      },
    ]),
  });
  if (!response.ok) {
    throw new Error("SUPABASE SAVE FAILED");
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
  return normalizeHttpUrl(state.data?.meta?.signatureMobileBaseUrl || "");
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

function typeUsesReferenceCatalog(typeEffet) {
  return ["CLE", "CLE CES"].includes(normalizeText(typeEffet));
}

function getReferenceCatalogType(typeEffet) {
  return typeUsesReferenceCatalog(typeEffet) ? "CLE" : normalizeText(typeEffet);
}

function typeUsesSiteField(typeEffet) {
  return ["CLE", "CLE CES", "BADGE INTRUSION", "CARTE TURBOSELF"].includes(normalizeText(typeEffet));
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
  if (window.location.protocol !== "file:") {
    return;
  }

  const fileName = window.location.pathname.split("/").pop() || "index.html";
  window.location.replace(`http://127.0.0.1:8123/${fileName}${window.location.search || ""}`);
}

function getCurrentPersonId() {
  const params = new URLSearchParams(window.location.search);
  return params.get("personId") || "";
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
    return;
  }
  window.history.replaceState({}, "", nextUrl);
}

function getCurrentMobileSignatureToken() {
  return new URLSearchParams(window.location.search).get("token") || "";
}

function getCurrentMobileSignatureDocType() {
  return new URLSearchParams(window.location.search).get("docType") || "";
}

function generateMobileSignatureToken() {
  return `SIG-${Date.now()}-${Math.random().toString(36).slice(2, 12).toUpperCase()}`;
}

function findMobileSignatureRequestByToken(token) {
  return (state.data?.demandesSignatureMobile || []).find((entry) => entry.token === token) || null;
}

function getActiveMobileSignatureRequest(personId, docType) {
  cleanupExpiredMobileSignatureRequests();
  return (state.data?.demandesSignatureMobile || []).find(
    (entry) =>
      entry.personId === personId &&
      entry.docType === normalizeText(docType) &&
      entry.status === "EN ATTENTE" &&
      Date.parse(entry.expiresAt || "") > Date.now()
  ) || null;
}

function createMobileSignatureRequest(personId, docType) {
  const createdAt = new Date();
  const expiresAt = new Date(createdAt.getTime() + MOBILE_SIGNATURE_REQUEST_TTL_MS);
  const request = {
    id: getNextId("DSM", state.data?.demandesSignatureMobile || []),
    token: generateMobileSignatureToken(),
    personId: String(personId || ""),
    docType: normalizeText(docType),
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
  return `signature-mobile.html?personId=${encodeURIComponent(request?.personId || "")}&docType=${encodeURIComponent(docType)}&token=${encodeURIComponent(request?.token || "")}`;
}

async function getMobileSignatureBaseUrl() {
  const configuredBaseUrl = getConfiguredMobileSignatureBaseUrl();
  if (configuredBaseUrl) {
    return configuredBaseUrl;
  }

  const currentOrigin = normalizeHttpUrl(window.location.origin || "");
  if (currentOrigin && !isLikelyLocalUrl(currentOrigin)) {
    return currentOrigin;
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
  return new URL(relativeUrl, `${String(baseUrl || window.location.origin).replace(/\/$/, "")}/`).href;
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

function renderMobileSignatureLink(docType, absoluteUrl) {
  const linkNode = document.getElementById(`${docType}-mobile-signature-link`);
  if (!linkNode) {
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

async function syncDocumentMobileSignatureLink(docType, personId) {
  if (!personId || !state.data) {
    renderMobileSignatureLink(docType, "");
    return;
  }
  const request = getActiveMobileSignatureRequest(personId, docType);
  if (!request) {
    renderMobileSignatureLink(docType, "");
    return;
  }
  const absoluteUrl = await getAbsoluteMobileSignatureUrl(request);
  renderMobileSignatureLink(docType, absoluteUrl);
}

async function openMobileSignatureRequest(docType, personId) {
  if (!state.data) {
    showDataStatus("DONNEES NON CHARGEES");
    return;
  }

  let request = getActiveMobileSignatureRequest(personId, docType);
  if (!request) {
    request = createMobileSignatureRequest(personId, docType);
    markDirty();
    await saveDataToFile({ silent: true });
  }

  const absoluteUrl = await getAbsoluteMobileSignatureUrl(request);
  window.open(absoluteUrl, "_blank", "noopener");
  renderMobileSignatureLink(docType, absoluteUrl);

  showDataStatus("PAGE DE SIGNATURE MOBILE OUVERTE");
  syncMobileSignaturePolling();
}

async function loadData() {
  bindPdfModalCleanup();
  applyActiveNav();
  bindHistoryNavigation();
  bindGlobalShortcuts();
  bindLoadButton();
  bindSaveButtons();
  bindPdfButtons();
  bindMobileSignatureButtons();
  bindDeletePersonButtons();
  bindFilterForms();
  bindAddPersonForm();
  bindPersonSheetForm();
  bindEffectForm();
  bindReferenceListForms();
  bindReferenceEffectForm();
  bindReplacementCostForm();
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
    showDataStatus("DONNEES EN COURS REPRISES - SAUVEGARDER POUR LES RENDRE DEFINITIVES");
    return;
  }

  await reloadData("OUVERTURE DES DONNEES...");
}

function applyPdfModeFromQuery() {
  const params = new URLSearchParams(window.location.search);
  if (params.get("pdf") === "1") {
    document.body.dataset.pdfMode = "true";
  }
}

async function reloadData(statusText = "RECHARGEMENT DES DONNEES...") {
  showDataStatus(statusText);

  try {
    const json = await fetchLatestDataSnapshot();
    state.data = json;
    migrateDataModel();
    clearWorkingData();
    state.isDirty = false;
    clearUndoStack();
    applyMeta();
    hydrateStaticLists();
    renderPage();
    showDataStatus(
      getDataBackendMode() === "SUPABASE" ? "DONNEES SUPABASE CHARGEES" : "DONNEES LOCALES CHARGEES"
    );
  } catch (error) {
    console.error(error);
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
  state.data.meta.signatureMobileBaseUrl = normalizeHttpUrl(state.data.meta.signatureMobileBaseUrl || "");

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
  if (!Array.isArray(state.data.demandesSignatureMobile)) {
    state.data.demandesSignatureMobile = [];
  }

  state.data.listes.typesContrats = Array.from(
    new Set([...state.data.listes.typesContrats.map(normalizeText), ...LEGACY_CONTRACT_TYPES])
  ).filter(Boolean);
  state.data.listes.fonctions = state.data.listes.fonctions.map(normalizeFunctionLabel).filter(Boolean);

  state.data.listes.typesPersonnel = state.data.listes.typesPersonnel
    .map(normalizeText)
    .filter((value) => value && !LEGACY_CONTRACT_TYPES.includes(value));
  state.data.listes.sites = Array.from(
    new Set([ALL_SITES_VALUE, ...state.data.listes.sites.map(normalizeText)])
  ).filter(Boolean);
  state.data.listes.typesEffets = Array.from(
    new Set([
      ...state.data.listes.typesEffets.map(normalizeText),
      "CLE",
      "CLE CES",
      "BADGE INTRUSION",
      "TELECOMMANDE URMET",
      "CARTE TURBOSELF",
    ])
  ).filter(Boolean);
  state.data.listes.statutsObjetManuels = Array.from(
    new Set([
      ...state.data.listes.statutsObjetManuels.map(normalizeText),
      "ACTIF",
      "PERDU",
      "DETRUIT",
      "HS",
      "REMPLACE",
    ])
  ).filter(Boolean);
  state.data.listes.causesRemplacement = Array.from(
    new Set([...state.data.listes.causesRemplacement.map(normalizeText), "PERTE", "CASSE", "NON RENDU", "VOL", "HS"])
  ).filter(Boolean);
  state.data.listes.coutsRemplacement = state.data.listes.coutsRemplacement
    .map((entry) => ({
      typeEffet: normalizeText(entry.typeEffet),
      cause: normalizeText(entry.cause),
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

  state.data.demandesSignatureMobile = state.data.demandesSignatureMobile
    .map((entry, index) => ({
      id: String(entry.id || `DSM${String(index + 1).padStart(4, "0")}`),
      token: String(entry.token || ""),
      personId: String(entry.personId || ""),
      docType: normalizeText(entry.docType),
      createdAt: String(entry.createdAt || ""),
      expiresAt: String(entry.expiresAt || ""),
      status: normalizeText(entry.status) || "EN ATTENTE",
      validatedAt: String(entry.validatedAt || ""),
    }))
    .filter((entry) => entry.token && entry.personId && ["ARRIVAL", "EXIT"].includes(entry.docType));

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
      effect.causeRemplacement = normalizeText(effect.causeRemplacement);
      effect.dateRemplacement = String(effect.dateRemplacement || "");
      effect.coutRemplacement = normalizeAmount(effect.coutRemplacement);
      effect.commentaire = normalizeText(effect.commentaire);
      delete effect.prixPaye;

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

      if (effect.dateRemplacement && !effect.statutManuel) {
        effect.statutManuel = "REMPLACE";
      }

      if (effect.statutManuel === "NON RENDU") {
        effect.statutManuel = "ACTIF";
      }

      if (effect.causeRemplacement === "NON RENDU" && getEffectStatus(person, effect) !== "NON RENDU") {
        effect.causeRemplacement = "";
      }

      if (effect.causeRemplacement && effect.causeRemplacement !== "NON RENDU") {
        effect.coutRemplacement = getEffectReplacementCost(person, effect);
      } else if (getEffectStatus(person, effect) === "NON RENDU") {
        effect.coutRemplacement = getEffectReplacementCost(person, {
          ...effect,
          causeRemplacement: "NON RENDU",
        });
      } else {
        effect.coutRemplacement = 0;
      }
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
    }))
    .filter((reference) => typeUsesReferenceCatalog(reference.typeEffet))
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
  sortListValues(state.data.listes.statutsObjetManuels);
  sortListValues(state.data.listes.causesRemplacement);
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
  return getActiveMobileSignatureRequest(personId, docType);
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
    const previousPerson = getCurrentPerson();
    const previousSignature = getSignatureValue(previousPerson, docType, "personnel");
    const previousValidatedAt = getSignatureValidationDate(previousPerson, docType, "personnel");

    state.data = json;
    migrateDataModel();

    if (nextRequest.status !== "EN ATTENTE") {
      renderMobileSignatureLink(docType, "");
      stopMobileSignaturePolling();
    }

    const updatedPerson = getCurrentPerson();
    const nextSignature = getSignatureValue(updatedPerson, docType, "personnel");
    const nextValidatedAt = getSignatureValidationDate(updatedPerson, docType, "personnel");

    if (previousSignature !== nextSignature || previousValidatedAt !== nextValidatedAt || nextRequest.status !== "EN ATTENTE") {
      renderPage();
      if (nextRequest.status === "SIGNEE") {
        showDataStatus("SIGNATURE MOBILE DU PERSONNEL ENREGISTREE");
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
    button.onclick = saveDataToFile;
  });
}

function bindPdfButtons() {
  document.querySelectorAll(".js-open-pdf").forEach((button) => {
    button.onclick = () => openPdfDocument(button.getAttribute("data-doc-type") || "", getCurrentPersonId());
  });
}

function bindMobileSignatureButtons() {
  document.querySelectorAll(".js-open-mobile-signature").forEach((button) => {
    button.onclick = async () => {
      const docType = String(button.getAttribute("data-doc-type") || "");
      const personId = getCurrentPersonId();
      if (!docType || !personId) {
        showDataStatus("AUCUNE PERSONNE SELECTIONNEE");
        return;
      }
      await openMobileSignatureRequest(docType, personId);
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
  switch (key) {
    case "nom":
      return person?.nom || "";
    case "prenom":
      return person?.prenom || "";
    case "fonction":
      return person?.fonction || "";
    case "site":
      return getPersonSiteLabel(person) || "";
    case "typeContrat":
      return person?.typeContrat || "";
    case "statutDossier":
      return getDossierStatus(person) || "";
    case "nbEffets":
      return (person?.effetsConfies || []).length;
    case "nonRendus":
      return (person?.effetsConfies || []).filter((effect) => getEffectStatus(person, effect) === "NON RENDU").length;
    case "couts":
      return (person?.effetsConfies || []).reduce((sum, effect) => sum + getEffectUnitValue(effect), 0);
    default:
      return "";
  }
}

function sortPersonsForOverview(persons) {
  const sort = getEffectTableSort("overviewPersons");
  const numericKeys = new Set(["nbEffets", "nonRendus", "couts"]);
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

function getHostedPdfDocumentPath(docType, personId) {
  const pagePath = getDocumentPagePath(docType);
  return `${pagePath}?personId=${encodeURIComponent(personId)}&pdf=1`;
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
    if (reusableArchive) {
      popup.location.href = getDocumentArchiveOpenPath(reusableArchive);
      showActionStatus("update", "PDF ARCHIVE REUTILISE");
      return;
    }

    if (getDataBackendMode() !== "LOCAL_API") {
      const hostedPath = getHostedPdfDocumentPath(docType, personId);
      const hostedUrl = `${hostedPath}&ts=${Date.now()}`;
      popup.location.href = hostedUrl;
      if (person) {
        await registerArchivedDocument(person, docType, hostedPath, "", archiveMode);
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
    const objectUrl = window.URL.createObjectURL(blob);
    hidePdfProgressModal();
    popup.location.href = objectUrl;
    if (archiveSaved && person) {
      registerArchivedDocument(person, docType, archivePdfPath, archiveMetadataPath, archiveMode).catch((error) => {
        console.error(error);
        showDataStatus("ARCHIVAGE PDF IMPOSSIBLE");
      });
    } else if (shouldArchive) {
      showDataStatus("PDF OUVERT - ARCHIVAGE NON REALISE");
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

function bindFilterForms() {
  document.querySelectorAll(".js-filter-form").forEach((form) => {
    form.oninput = () => {
      state.filters = readFilters(form);
      renderPage();
    };
    form.onreset = () => {
      window.setTimeout(() => {
        state.filters = readFilters(form);
        renderPage();
      }, 0);
    };
  });
}

function bindArchiveFilterForm() {
  const form = document.getElementById("documents-archives-filter-form");
  if (!form) {
    return;
  }
  form.oninput = () => {
    renderDocumentsArchivePage();
  };
  form.onreset = () => {
    window.setTimeout(() => {
      renderDocumentsArchivePage();
    }, 0);
  };
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
    window.location.href = `fiche-personne.html?personId=${person.id}`;
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

  const buildPersonFromSheetForm = () => {
    const formData = new FormData(form);
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

    if (!person.nom && !person.prenom) {
      person.nom = "PERSONNE";
      person.prenom = person.id;
    }

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
      window.location.href = `document-arrivee.html?personId=${personId}`;
    };
  }

  if (exitDocumentButton) {
    exitDocumentButton.onclick = () => {
      const personId = getSheetTargetPersonId();
      if (!personId) {
        showDataStatus("AUCUNE PERSONNE SELECTIONNEE");
        return;
      }
      window.location.href = `document-sortie.html?personId=${personId}`;
    };
  }

  if (arrivalPdfButton) {
    arrivalPdfButton.onclick = () => openPdfDocument("arrival", getSheetTargetPersonId());
  }

  if (exitPdfButton) {
    exitPdfButton.onclick = () => openPdfDocument("exit", getSheetTargetPersonId());
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
  const typeField = form.elements.typeEffet;
  const referenceSiteField = form.elements.referenceSite;
  const causeField = form.elements.causeRemplacement;
  const replacementDateField = form.elements.dateRemplacement;
  if (typeField) {
    typeField.onchange = () => {
      const person = getCurrentPerson();
      hydrateEffectReferenceSiteSelect(person, "", typeField.value);
      hydrateReferenceSelect(person || "", typeField.value, "", getSelectedEffectReferenceSite());
      updateEffectFormMode(typeField.value);
      syncReplacementCostField();
    };
  }
  if (referenceSiteField) {
    referenceSiteField.onchange = () => {
      const person = getCurrentPerson();
      hydrateReferenceSelect(person || "", form.elements.typeEffet.value, "", getSelectedEffectReferenceSite());
      syncReplacementCostField();
    };
  }
  if (causeField) {
    causeField.onchange = () => {
      syncReplacementCostField();
    };
  }
  if (replacementDateField) {
    replacementDateField.onchange = () => {
      if (replacementDateField.value) {
        form.elements.statutManuel.value = "REMPLACE";
      }
    };
  }
  if (form.elements.referenceEffet) {
    form.elements.referenceEffet.onchange = () => {
      syncReplacementCostField();
    };
  }
  if (form.elements.designationLibre) {
    form.elements.designationLibre.oninput = () => {
      syncReplacementCostField();
    };
  }

  const submitEffect = (mode) => {
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
    const causeRemplacement = normalizeText(formData.get("causeRemplacement"));
    const dateRemplacement = String(formData.get("dateRemplacement") || "");
    const coutRemplacement = normalizeAmount(formData.get("coutRemplacement"));

    if (usesReferenceCatalog && !referenceEffetId) {
      showDataStatus("CHOISIR UNE CLE EXISTANTE DANS LA LISTE");
      return;
    }

    if (usesReferenceCatalog && designationLibre && !resolvedReferenceSite) {
      showDataStatus("CHOISIR LE SITE DE LA CLE AVANT D'ENREGISTRER");
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
      dateRetour: String(formData.get("dateRetour") || ""),
      statutManuel: dateRemplacement
        ? "REMPLACE"
        : normalizeText(formData.get("statutManuel")) || "ACTIF",
      causeRemplacement,
      dateRemplacement,
      coutRemplacement,
      commentaire: normalizeText(formData.get("commentaire")),
    };

    if (!effect.typeEffet) {
      effect.typeEffet = "EFFET";
    }
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
  };

  form.onsubmit = (event) => {
    event.preventDefault();
    submitEffect(state.editingEffectId ? "edit" : "add");
  };

  if (addButton) {
    addButton.onclick = () => submitEffect("add");
  }
  if (updateButton) {
    updateButton.onclick = () => submitEffect("edit");
  }
  if (deleteButton) {
    deleteButton.onclick = () => {
      const person = getCurrentPerson();
      if (!person || !state.editingEffectId) {
        showDataStatus("SELECTIONNER D'ABORD UN EFFET A SUPPRIMER");
        return;
      }
      deleteEffect(person.id, state.editingEffectId);
    };
  }
  if (cancelButton) {
    cancelButton.onclick = () => {
      state.editingEffectId = "";
      resetEffectForm();
      showDataStatus("MODIFICATION DE L'EFFET ANNULEE");
    };
  }

  updateEffectActionButtons();
}

function deleteEffect(personId, effectId) {
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
  if (state.editingEffectId === effectId) {
    state.editingEffectId = "";
    resetEffectForm();
  }
  markDirty();
  renderPage();
  renderPersonSheet(personId);
  showActionStatus("delete", `EFFET SUPPRIME : ${effectId}`);
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
    window.location.href = "suivi-global.html";
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
  form.elements.causeRemplacement.value = effect.causeRemplacement || "";
  form.elements.dateRemplacement.value = effect.dateRemplacement || "";
  form.elements.coutRemplacement.value = formatAmountWithEuro(effect.coutRemplacement);
  form.elements.commentaire.value = effect.commentaire || "";
  updateEffectFormMode(effect.typeEffet || "");
  updateEffectActionButtons();
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

function updateEffectFormMode(typeEffet) {
  const normalizedType = normalizeText(typeEffet);
  const person = getCurrentPerson();
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

    referenceSiteField.classList.add("is-hidden");
    referenceField.classList.remove("is-hidden");
    designationField.classList.remove("is-hidden");
    vehicleField.classList.add("is-hidden");
    const vehicleInput = vehicleField.querySelector("input");
    if (vehicleInput && normalizedType !== "TELECOMMANDE URMET") {
      vehicleInput.value = "";
    }
    numberLabel.textContent = "N° D'IDENTIFICATION";

  if (["CLE", "CLE CES"].includes(normalizedType)) {
    referenceSiteField.classList.remove("is-hidden");
    referenceSiteLabel.textContent = "SITE DE LA CLE";
    referenceLabel.textContent = "NOM EXISTANT DE LA CLE";
    designationLabel.textContent = "NOUVEAU NOM / MODIFICATION";
    designationField.classList.add("is-hidden");
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
    return;
  }

    if (["BADGE INTRUSION", "TELECOMMANDE URMET", "CARTE TURBOSELF"].includes(normalizedType)) {
      if (["BADGE INTRUSION", "CARTE TURBOSELF"].includes(normalizedType)) {
        referenceSiteField.classList.remove("is-hidden");
        referenceSiteLabel.textContent = normalizedType === "BADGE INTRUSION" ? "SITE DU BADGE" : "SITE DE LA CARTE";
      }
      vehicleField.classList.toggle("is-hidden", normalizedType !== "TELECOMMANDE URMET");
      vehicleLabel.textContent = "VEHICULE / IMMATRICULATION";
      referenceField.classList.add("is-hidden");
    designationField.classList.add("is-hidden");
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
    return;
  }

  referenceLabel.textContent = "DESIGNATION EXISTANTE";
  designationLabel.textContent = "NOUVELLE DESIGNATION / MODIFICATION";
  helpNode.textContent = normalizedType
    ? "SI BESOIN : CHOISIR UNE REFERENCE EXISTANTE OU SAISIR UNE DESIGNATION"
    : "CHOISIR UN TYPE D'EFFET POUR ADAPTER LA SAISIE";
}

function isCesKeyDesignation(designation) {
  return normalizeText(designation).startsWith("CES-");
}

function getReplacementCostValue(typeEffet, causeRemplacement, designation = "") {
  const normalizedType = normalizeText(typeEffet);
  const normalizedCause = normalizeText(causeRemplacement) || "PERTE";
  if (!normalizedType) {
    return 0;
  }

  if (normalizedType === "CLE CES") {
    return CHARGED_REPLACEMENT_CAUSES.includes(normalizedCause) ? 50 : 0;
  }

  const matchingEntry = (state.data?.listes?.coutsRemplacement || []).find(
    (entry) =>
      normalizeText(entry.typeEffet) === normalizedType && normalizeText(entry.cause) === normalizedCause
  );

  if (!matchingEntry) {
    return 0;
  }

  if (!CHARGED_REPLACEMENT_CAUSES.includes(normalizedCause)) {
    return 0;
  }

  if (normalizedType === "CLE") {
    return isCesKeyDesignation(designation) ? 50 : 5;
  }

  return normalizeAmount(matchingEntry.montant);
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
  const explicitCause = normalizeText(effect?.causeRemplacement);
  if (explicitCause && explicitCause !== "NON RENDU") {
    return explicitCause;
  }

  return "";
}

function getEffectReplacementCost(person, effect) {
  const explicitCause = getEffectReplacementCause(person, effect);
  const effectiveCause = explicitCause || (getEffectStatus(person, effect) === "NON RENDU" ? "NON RENDU" : "");
  const cause = normalizeText(effectiveCause);
  if (!cause) {
    return 0;
  }

  if (cause === "HS") {
    return 0;
  }

  const storedCost = normalizeAmount(effect?.coutRemplacement);
  if (storedCost > 0) {
    return storedCost;
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
  const causeRemplacement = form.elements.causeRemplacement?.value || "";
  const referenceId = form.elements.referenceEffet?.value || "";
  const reference = findReferenceById(referenceId);
  const designation =
    form.elements.designationLibre?.value || reference?.designation || "";
  const coutField = form.elements.coutRemplacement;
  if (!coutField) {
    return;
  }

  coutField.value = formatAmountWithEuro(getReplacementCostValue(typeEffet, causeRemplacement, designation));
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
      const value = normalizeText(input?.value);
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
          (entry) => normalizeText(entry) === value && normalizeText(entry) !== normalizeText(oldValue)
        );
        if (duplicate) {
          showDataStatus("VALEUR DEJA PRESENTE");
          return;
        }

        const index = list.findIndex((entry) => normalizeText(entry) === normalizeText(oldValue));
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

      if (list.some((entry) => normalizeText(entry) === value)) {
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
  const designationField = form.elements.referenceDesignation;
  if (typeField) {
    typeField.onchange = () => {
      updateReferenceEffectFormMode(typeField.value);
      syncReferenceSitesSelector();
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

function bindReferenceFilters() {
  const form = document.getElementById("reference-filter-form");
  if (!form) {
    return;
  }

  form.oninput = () => {
    renderReferenceEffectsTable();
  };

  form.onreset = () => {
    window.setTimeout(() => {
      renderReferenceEffectsTable();
    }, 0);
  };
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
    const normalized = normalizeHttpUrl(rawValue);
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
    search: normalizeText(formData.get("search")),
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
    causesRemplacement = [],
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
  populateSelect('select[name="causeRemplacement"]', causesRemplacement);
  populateSelect('select[name="costTypeEffet"]', typesEffets);
  populateSelect('select[name="costCauseRemplacement"]', causesRemplacement);
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
  const currentPersonId = getCurrentPersonId();

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
    renderPersonPicker();
  }

  if (page === "mobile-signature") {
    renderMobileSignaturePage();
    refreshDocumentSignatureCanvases(getCurrentMobileSignatureDocType());
  }

  if (page === "person-sheet") {
    renderPersonSheet(currentPersonId);
    bindEffectTableSorting();
    updateSortableHeaders("sheetEffects");
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
  }

  renderDirtyState();
}

function renderEffectsChart(nodeId, persons) {
  const node = document.getElementById(nodeId);
  if (!node) {
    return;
  }

  const effects = getAllEffects(persons)
    .filter(({ person, effect }) => effectMatchesSiteFilter(person, effect, state.filters?.site))
    .map(({ effect }) => effect);
  if (!effects.length) {
    node.innerHTML = '<div class="effects-chart__empty">AUCUNE DONNEE A AFFICHER</div>';
    return;
  }

  const counts = new Map();
  effects.forEach((effect) => {
    const type = normalizeText(effect?.typeEffet) || "EFFET";
    counts.set(type, (counts.get(type) || 0) + 1);
  });

  const rows = Array.from(counts.entries())
    .sort((left, right) => {
      const typeOrder = left[0].localeCompare(right[0], "fr");
      if (typeOrder !== 0) {
        return typeOrder;
      }
      return left[1] - right[1];
    });

  const maxValue = Math.max(...rows.map(([, count]) => count), 1);
  node.innerHTML = rows
    .map(([type, count]) => {
      const width = Math.max(8, Math.round((count / maxValue) * 100));
      return `<div class="effects-chart__row">
        <span class="effects-chart__label">${escapeHtml(type)}</span>
        <span class="effects-chart__track" aria-hidden="true">
          <span class="effects-chart__bar" style="width:${width}%"></span>
        </span>
        <strong class="effects-chart__value">${count}</strong>
      </div>`;
    })
    .join("");
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
  return Boolean(getSignatureValue(person, docType, "personnel")) && Boolean(getSignatureValue(person, docType, "representant"));
}

function getDocumentArchiveDate(person, docType) {
  return normalizeText(docType) === "EXIT"
    ? person?.dateSortieReelle || person?.dateSortiePrevue || getTodayIsoDate()
    : person?.dateEntree || getTodayIsoDate();
}

function getDocumentArchiveEntryMode(entry) {
  return normalizeText(entry?.documentMode || "STANDARD");
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
  return `/${raw.replace(/^\/+/, "")}`;
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
  const baseId = `DOC-${getDocumentTypeLabel(docType)}-${person?.id || ""}-${documentMode}`;
  return {
    id: documentMode === "COMPLEMENTAIRE" ? `${baseId}-${Date.now()}` : baseId,
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
  sortDocumentsArchives();
}

async function registerArchivedDocument(person, docType, pdfPath, metadataPath, archiveMode = "STANDARD") {
  if (!state.data || !person || !pdfPath) {
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

  const filterForm = document.getElementById("documents-archives-filter-form");
  const search = normalizeText(filterForm?.elements?.archiveSearch?.value);
  const typeDocument = normalizeText(filterForm?.elements?.archiveTypeDocument?.value);
  const site = normalizeText(filterForm?.elements?.archiveSite?.value);
  const statutSignature = normalizeText(filterForm?.elements?.archiveStatutSignature?.value);
  const archiveSiteSelect = filterForm?.elements?.archiveSite;
  syncSelectOptions(archiveSiteSelect, state.data?.listes?.sites || [], "TOUS");

  let totalArchives = 0;
  let totalArrivalArchives = 0;
  let totalExitArchives = 0;
  const archives = (state.data?.documentsArchives || []).filter((entry) => {
    totalArchives += 1;
    if (normalizeText(entry.typeDocument) === "ARRIVEE") {
      totalArrivalArchives += 1;
    }
    if (normalizeText(entry.typeDocument) === "SORTIE") {
      totalExitArchives += 1;
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
    if (search) {
      const haystack = [
        entry.nom,
        entry.prenom,
        entry.typeDocument,
        entry.sites,
        entry.typePersonnel,
        entry.typeContrat,
        entry.pdfPath,
      ]
        .map(normalizeText)
        .join(" ");
      if (!haystack.includes(search)) {
        return false;
      }
    }
    return true;
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

  if (!archives.length) {
    body.innerHTML = buildEmptyTableRow(body, "AUCUN DOCUMENT ARCHIVE", 11);
    return;
  }

  const rowsHtml = archives
    .map(
      (entry) => {
        const openPath = getDocumentArchiveOpenPath(entry);
        return `<tr>
        <td>${escapeHtml(entry.nom || "-")}</td>
        <td>${escapeHtml(entry.prenom || "-")}</td>
        <td>${escapeHtml(entry.typeDocument || "-")}</td>
        <td>${escapeHtml(formatDate(entry.dateDocument) || "-")}</td>
        <td>${escapeHtml(formatTime(entry.dateArchivage) || "-")}</td>
        <td>${escapeHtml(entry.sites || "-")}</td>
        <td>${escapeHtml(getDocumentArchiveSignatureStatus(entry))}</td>
        <td>${escapeHtml(String(entry.totalEffets ?? "-"))}</td>
        <td>${formatAmountWithEuro(entry.totalFacturable || 0)}</td>
        <td>${escapeHtml(entry.pdfPath || getDocumentArchiveStoragePath(entry))}</td>
        <td>${openPath ? `<a class="archive-pdf-button" href="${escapeHtml(openPath)}" target="_blank" rel="noopener" aria-label="OUVRIR PDF"><span class="archive-pdf-button__icon" aria-hidden="true"><img src="assets/images/ui/icone-pdf.png" alt="" class="archive-pdf-button__image" /></span></a>` : "-"}</td>
      </tr>`;
      }
    );

  renderTableRowsProgressively(body, rowsHtml, buildEmptyTableRow(body, "AUCUN DOCUMENT ARCHIVE", 11), 24);
}

function getSignatureValue(person, docType, signer) {
  if (!person?.signatures?.[docType]) {
    return "";
  }
  return String(person.signatures[docType][signer]?.image || "");
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

function setSignatureValue(person, docType, signer, value, validatedAt = "") {
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
  canvas.width = Math.max(1, Math.round(width * ratio));
  canvas.height = Math.max(1, Math.round(height * ratio));
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
      markDirty();
      renderPage();
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
      setSignatureValue(person, docType, signer, nextValue, nextValue ? getCurrentSignatureTimestamp() : "");
      if (document.body.dataset.page === "mobile-signature" && signer === "personnel" && nextValue) {
        const request = getCurrentMobileSignatureRequest();
        if (request && normalizeText(request.docType) === normalizeText(docType)) {
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
      showActionStatus(nextValue ? "update" : "delete", saveText);
      await saveDataToFile({
        silent: true,
        successText: saveText,
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
        if (signer === "representant" && !hasRepresentativeIdentityForDocument(docType)) {
          showDataStatus("IDENTITE DU REPRESENTANT OBLIGATOIRE AVANT VALIDATION");
          window.alert("VOUS DEVEZ IDENTIFIER L'IDENTITE DU REPRESENTANT DE L'ETABLISSEMENT POUR VALIDATION.");
          updateRepresentativeSignatureActionState(docType);
          return;
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
        setSignatureValue(person, docType, signer, "", "");
        if (document.body.dataset.page === "mobile-signature" && signer === "personnel") {
          const request = getCurrentMobileSignatureRequest();
          if (request) {
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
  const titleNode = document.getElementById("mobile-signature-title");
  const subtitleNode = document.getElementById("mobile-signature-subtitle");
  const personNode = document.getElementById("mobile-signature-person");
  const dateNode = document.getElementById("mobile-signature-date");
  const statusNode = document.getElementById("mobile-signature-request-status");
  const panelNode = document.getElementById("mobile-signature-panel");
  const saveButton = document.querySelector('.js-signature-save[data-signer="personnel"]');
  const clearButton = document.querySelector('.js-signature-clear[data-signer="personnel"]');
  const canvas = document.querySelector('.js-signature-canvas[data-signer="personnel"]');
  const docLabel = normalizeText(docType) === "EXIT" ? "DOCUMENT DE SORTIE" : "DOCUMENT D'ARRIVEE";

  if (canvas) {
    canvas.setAttribute("data-doc-type", normalizeText(docType) === "EXIT" ? "exit" : "arrival");
  }
  if (saveButton) {
    saveButton.setAttribute("data-doc-type", normalizeText(docType) === "EXIT" ? "exit" : "arrival");
  }
  if (clearButton) {
    clearButton.setAttribute("data-doc-type", normalizeText(docType) === "EXIT" ? "exit" : "arrival");
  }

  if (titleNode) {
    titleNode.textContent = "SIGNATURE DU PERSONNEL";
  }
  if (subtitleNode) {
    subtitleNode.textContent = docLabel;
  }
  if (personNode) {
    personNode.textContent = person ? `${person.nom || ""} ${person.prenom || ""}`.trim() || "-" : "-";
  }
  if (dateNode) {
    dateNode.textContent = docType === "exit" ? formatDate(person?.dateSortieReelle || person?.dateSortiePrevue) || "-" : formatDate(person?.dateEntree) || "-";
  }

  const valid = Boolean(person && request && isMobileSignatureRequestValid(request) && normalizeText(request.docType) === normalizeText(docType));
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
    } else {
      const expires = formatSignatureTimestamp(request.expiresAt);
      statusNode.textContent = expires ? `DEMANDE ACTIVE JUSQU'A ${expires}` : "DEMANDE ACTIVE";
    }
  }

  fillMobileSignatureShareLink(valid ? request : null);
}

function renderPersonPicker() {
  const picker = document.getElementById("person-picker-search");
  const pickerList = document.getElementById("person-picker-list");
  const suggestionBox = document.getElementById("person-picker-suggestions");
  if (!picker || !pickerList || !state.data?.personnes) {
    return;
  }

  const currentPersonId = getCurrentPersonId();
  const selectedPerson = currentPersonId
    ? state.data.personnes.find((person) => person.id === currentPersonId) || null
    : null;

  const pickerIsFocused = document.activeElement === picker;
  if (!pickerIsFocused) {
    picker.value = selectedPerson ? getPersonPickerLabel(selectedPerson) : "";
  }

  const page = document.body.dataset.page || "";
  const useDirectNavigation = page === "arrival-document" || page === "exit-document";
  const useSuggestionBox = useDirectNavigation && suggestionBox;
  const options = state.data.personnes
    .map((person) => {
      const label = getPersonPickerLabel(person);
      return `<option value="${escapeHtml(label)}"></option>`;
    })
    .join("");

  if (useSuggestionBox) {
    picker.removeAttribute("list");
    pickerList.innerHTML = "";
  } else {
    picker.setAttribute("list", "person-picker-list");
    pickerList.innerHTML = options;
  }

  const renderSuggestions = (rawQuery = "") => {
    if (!useSuggestionBox) {
      return;
    }

    const query = normalizeText(rawQuery);
    const matches = state.data.personnes
      .filter((person) => {
        if (!query) {
          return true;
        }
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

  const applyPickerSelection = (mode = "push") => {
    const rawValue = String(picker.value || "");
    if (!rawValue.trim()) {
      if (useDirectNavigation) {
        hideSuggestions();
        applyDocumentNavigation("");
        return true;
      }
      setCurrentPersonId("", mode);
      renderPage();
      return false;
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
        renderSuggestions("");
        return;
      }
      setCurrentPersonId("", "replace");
      renderPage();
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
      renderSuggestions(picker.value);
    }
  };
  picker.onchange = () => applyPickerSelection("push");
  picker.onsearch = () => applyPickerSelection("push");
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
      applyDocumentNavigation("");
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
    if (getDossierStatus(person) === "EN POSTE") {
      inPostCount += 1;
    }
    (person.effetsConfies || []).forEach((effect) => {
      totalEffectsCount += 1;
      if (getEffectStatus(person, effect) === "NON RENDU") {
        missingEffectsCount += 1;
      }
    });
  });

  if (inPostNode) {
    inPostNode.textContent = String(inPostCount);
  }
  if (totalEffectsNode) {
    totalEffectsNode.textContent = String(totalEffectsCount);
  }
  if (missingEffectsNode) {
    missingEffectsNode.textContent = String(missingEffectsCount);
  }

  renderEffectsChart("overview-effects-chart", persons);

  if (body) {
    const sortedPersons = sortPersonsForOverview(persons);
    const rowsHtml = buildOverviewRows(sortedPersons.slice(0, 8));
    renderTableRowsProgressively(body, [rowsHtml], buildEmptyTableRow("overview-table-body", "AUCUNE DONNEE A AFFICHER", 9), 1);
    bindPersonRowActions();
    updateSortableHeaders("overviewPersons");
  }

  if (alertsSection && alertsList) {
    const alerts = persons
      .filter((person) => hasOverdueExit(person))
      .map((person) => ({
        id: person.id,
        nom: person.nom,
        prenom: person.prenom,
        message: getOverdueExitMessage(person),
      }));

    alertsSection.hidden = alerts.length === 0;
    alertsList.innerHTML = alerts
      .map(
        (alert) => `<button type="button" class="overview-alert-item js-open-person-alert" data-person-id="${alert.id}">
          <strong>${escapeHtml(`${alert.nom} ${alert.prenom}`.trim())}</strong>
          <span>${escapeHtml(alert.message)}</span>
        </button>`
      )
      .join("");

    alertsList.querySelectorAll(".js-open-person-alert").forEach((button) => {
      button.onclick = () => {
        const personId = button.getAttribute("data-person-id") || "";
        if (personId) {
          window.location.href = `fiche-personne.html?personId=${personId}`;
        }
      };
    });
  }
}

function buildOverviewRows(persons) {
  if (!persons.length) {
    return buildEmptyTableRow("overview-table-body", "AUCUNE DONNEE A AFFICHER", 9);
  }
  const rows = persons
    .map((person) => {
      const nonRendus = (person.effetsConfies || []).filter(
        (effect) => getEffectStatus(person, effect) === "NON RENDU"
      ).length;
      const totalCost = (person.effetsConfies || []).reduce(
        (sum, effect) => sum + getEffectUnitValue(effect),
        0
      );
      const alertClass = hasOverdueExit(person) ? " is-alert-row" : "";
      return `<tr class="js-person-row${alertClass}" data-person-id="${person.id}">
        <td>${person.nom}</td>
        <td>${person.prenom}</td>
        <td>${person.fonction || ""}</td>
        <td>${getPersonSiteMarkup(person)}</td>
        <td>${person.typeContrat || ""}</td>
        <td>${getDossierStatusCellMarkup(getDossierStatus(person))}</td>
        <td>${(person.effetsConfies || []).length}</td>
        <td>${nonRendus > 0 ? '<span class="row-alert-dot" aria-hidden="true"></span>' : ""}${nonRendus}</td>
        <td>${formatAmountWithEuro(totalCost)}</td>
      </tr>`;
    })
    .join("");
  const totalEffects = persons.reduce((sum, person) => sum + (person.effetsConfies || []).length, 0);
  const totalCost = persons.reduce(
    (sum, person) =>
      sum + (person.effetsConfies || []).reduce((innerSum, effect) => innerSum + getEffectUnitValue(effect), 0),
    0
  );
  const totalNonRendus = persons.reduce(
    (sum, person) =>
      sum +
      (person.effetsConfies || []).filter((effect) => getEffectStatus(person, effect) === "NON RENDU").length,
    0
  );

  return `${rows}
    <tr class="table-total-row">
      <td colspan="6">TOTAL</td>
      <td>${totalEffects}</td>
      <td>${totalNonRendus}</td>
      <td>${formatAmountWithEuro(totalCost)}</td>
    </tr>`;
}

function renderGlobalTable(persons) {
  const body = document.getElementById("global-table-body");
  if (!body) {
    return;
  }

  if (!persons.length) {
    body.innerHTML = buildEmptyTableRow(body, "AUCUNE DONNEE A AFFICHER", 12);
    return;
  }

  const rowsHtml = persons
    .map((person) => {
      const totalEffects = (person.effetsConfies || []).length;
      const nonRendus = (person.effetsConfies || []).filter(
        (effect) => getEffectStatus(person, effect) === "NON RENDU"
      ).length;
      const alertClass = hasOverdueExit(person) ? " is-alert-row" : "";
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
        <td>
          <a class="table-link" href="fiche-personne.html?personId=${person.id}">VOIR</a>
          <button type="button" class="table-link js-delete-person" data-person-id="${person.id}">SUPPRIMER</button>
        </td>
      </tr>`;
    })
    ;

  renderTableRowsProgressively(body, rowsHtml, buildEmptyTableRow(body, "AUCUNE DONNEE A AFFICHER", 12), 24);

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

      window.location.href = `fiche-personne.html?personId=${personId}`;
    });

    body.dataset.bound = "true";
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
    return '<img src="assets/images/sidebar/icone-cle-ces.png" alt="" loading="lazy">';
  }
  if (variant === "badge") {
    return '<img src="assets/images/sidebar/icone-badge.png" alt="" loading="lazy">';
  }
  if (variant === "telecommande") {
    return '<img src="assets/images/sidebar/icone-telecommande.png" alt="" loading="lazy">';
  }
  if (variant === "carte") {
    return '<img src="assets/images/sidebar/icone-carte.png" alt="" loading="lazy">';
  }
  if (variant === "cle") {
    return '<img src="assets/images/sidebar/icone-cle.png" alt="" loading="lazy">';
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

  const person = state.data?.personnes?.find((entry) => entry.id === personId);
  if (!person) {
    state.currentSheetPersonId = "";
    nameNode.textContent = "AUCUNE PERSONNE SELECTIONNEE";
    metaNode.textContent = "SELECTIONNER UNE PERSONNE POUR AFFICHER LA FICHE";
    alertNode.hidden = true;
    alertNode.textContent = "";
    applySheetPersonStatus(statusNode, "EN ATTENTE");
    body.innerHTML = buildEmptyTableRow(body, "AUCUN EFFET A AFFICHER", 10);
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

  state.currentSheetPersonId = person.id;
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
  const sortedEffects = sortEffectsForTable(person, effects, "sheetEffects");
  const returned = effects.filter((effect) => getEffectStatus(person, effect) === "RESTITUE").length;
  const missing = effects.filter((effect) => getEffectStatus(person, effect) === "NON RENDU").length;
  const totalCost = effects.reduce((sum, effect) => sum + getEffectReplacementCost(person, effect), 0);
  const totalEffectsUnitValue = effects.reduce((sum, effect) => sum + getEffectUnitValue(effect), 0);

  if (totalNode) totalNode.textContent = String(effects.length);
  if (returnedNode) returnedNode.textContent = String(returned);
  if (missingNode) missingNode.textContent = String(missing);
  if (costNode) costNode.textContent = formatAmountWithEuro(totalCost);
  renderSheetEffectTypeKpis(effects);
  if (totalTypesNode) totalTypesNode.textContent = String(effects.length);
  if (totalTypesAmountNode) totalTypesAmountNode.textContent = formatAmountWithEuro(totalEffectsUnitValue);

  body.innerHTML = sortedEffects.length
    ? `${sortedEffects
        .map((effect) => {
          const effectStatus = getEffectStatus(person, effect);
          const effectCause = getEffectReplacementCause(person, effect);
          const effectDesignation = getEffectDisplayDesignation(effect);
          const effectSite = getEffectDisplaySite(effect);
          const effectUnitValue = getEffectUnitValue(effect);
          return `<tr class="js-effect-row" data-person-id="${person.id}" data-effect-id="${effect.id}">
            <td>${effect.typeEffet || ""}</td>
            <td>${effectDesignation}</td>
            <td>${effectSite}</td>
            <td>${effect.numeroIdentification || ""}</td>
            <td>${formatDate(effect.dateRemise)}</td>
            <td>${formatDate(effect.dateRetour)}</td>
            <td><span class="${getStatusClass(effectStatus)}"><span>${effectStatus}</span>${effectStatus === "NON RENDU" ? '<span class="row-alert-dot row-alert-dot--inside" aria-hidden="true"></span>' : ""}</span></td>
            <td>${effectCause}</td>
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

  const currentTypeEffet = document.querySelector('#effect-form [name="typeEffet"]')?.value || "";
  const currentReferenceSite = document.querySelector('#effect-form [name="referenceSite"]')?.value || "";
  const currentReferenceId = document.querySelector('#effect-form [name="referenceEffet"]')?.value || "";
  hydrateEffectReferenceSiteSelect(person, currentReferenceSite, currentTypeEffet);
  hydrateReferenceSelect(person, currentTypeEffet, currentReferenceId, currentReferenceSite);
  updateEffectFormMode(currentTypeEffet);
  bindEffectRowActions();
  updateSortableHeaders("sheetEffects");
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
}

function renderArrivalDocument(personId) {
  const person = state.data?.personnes?.find((entry) => entry.id === personId) || null;
  const mode = normalizeText(new URLSearchParams(window.location.search).get("mode") || "STANDARD");
  const isComplement = mode === "COMPLEMENTAIRE";
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
    body.innerHTML = buildEmptyTableRow(body, "AUCUN EFFET A AFFICHER", 5);
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
    syncDocumentMobileSignatureLink("arrival", "");
    return;
  }

  const effects = (person.effetsConfies || []).filter((effect) => Boolean(effect.dateRemise));
  const sortedEffects = sortEffectsForTable(person, effects, "arrivalEffects");
  const totalValue = effects.reduce((sum, effect) => sum + getEffectUnitValue(effect), 0);

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

  body.innerHTML = sortedEffects.length
    ? `${sortedEffects
        .map(
          (effect) => `<tr>
            <td>${effect.typeEffet || ""}</td>
            <td>${getEffectDisplayDesignation(effect) || "-"}</td>
            <td>${effect.numeroIdentification || "-"}</td>
            <td>${formatDate(effect.dateRemise) || "-"}</td>
            <td>${formatAmountWithEuro(getEffectUnitValue(effect))}</td>
          </tr>`
        )
        .join("")}
        <tr class="table-total-row">
          <td colspan="4">TOTAL DES EFFETS REMIS</td>
          <td>${formatAmountWithEuro(totalValue)}</td>
        </tr>`
    : buildEmptyTableRow(body, "AUCUN EFFET A AFFICHER", 5);

  renderArrivalCostsTable(costsHead, costsBody);
  totalEffectsNode.textContent = String(effects.length);
  totalValueNode.textContent = formatAmountWithEuro(totalValue);
  updateSortableHeaders("arrivalEffects");
  syncDocumentMobileSignatureLink("arrival", person.id);
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
  const baseOrder = ["HS", "PERTE", "CASSE", "VOL", "NON RENDU"];
  const causes = Array.isArray(state.data?.listes?.causesRemplacement)
    ? state.data.listes.causesRemplacement.map(normalizeText).filter(Boolean)
    : [];
  const merged = Array.from(new Set([...baseOrder, ...causes]));
  return merged.filter((value) => value !== "REMPLACE");
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
    body.innerHTML = buildEmptyTableRow(body, "AUCUN EFFET A AFFICHER", 8);
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
    syncDocumentMobileSignatureLink("exit", "");
    return;
  }

  const effects = person.effetsConfies || [];
  const sortedEffects = sortEffectsForTable(person, effects, "exitEffects");
  const totalReturned = effects.filter((effect) => getEffectStatus(person, effect) === "RESTITUE").length;
  const chargeableEffects = effects.filter((effect) => isEffectChargeable(person, effect));
  const totalValue = chargeableEffects.reduce((sum, effect) => sum + getEffectReplacementCost(person, effect), 0);

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
          (effect) => `<tr>
            <td>${effect.typeEffet || ""}</td>
            <td>${getEffectDisplayDesignation(effect)}</td>
            <td>${effect.numeroIdentification || ""}</td>
            <td>${formatDate(effect.dateRemise)}</td>
            <td>${formatDate(effect.dateRetour)}</td>
            <td>${getEffectStatus(person, effect)}</td>
            <td>${getEffectReplacementCause(person, effect)}</td>
            <td>${formatAmountWithEuro(getEffectReplacementCost(person, effect))}</td>
          </tr>`
        )
        .join("")}
        <tr class="table-total-row">
          <td colspan="7">TOTAL FACTURABLE DES EFFETS</td>
          <td>${formatAmountWithEuro(totalValue)}</td>
        </tr>`
    : buildEmptyTableRow(body, "AUCUN EFFET A AFFICHER", 8);

  totalEffectsNode.textContent = String(effects.length);
  totalReturnedNode.textContent = String(totalReturned);
  totalChargeableNode.textContent = String(chargeableEffects.length);
  totalValueNode.textContent = formatAmountWithEuro(totalValue);
  renderDocumentCostsTable(costsHead, costsBody);
  updateSortableHeaders("exitEffects");
  syncDocumentMobileSignatureLink("exit", person.id);
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

  const selectedSites = Array.isArray(siteSource)
    ? normalizeSites(siteSource)
    : siteSource && typeof siteSource === "object"
      ? getPersonSites(siteSource)
      : normalizeSites(siteSource ? [siteSource] : []);
  const normalizedTypeEffet = normalizeText(typeEffet);
  const baseOption =
    select.querySelector("option")?.outerHTML || '<option value="">SELECTIONNER</option>';
  if (!typeUsesReferenceCatalog(normalizedTypeEffet)) {
    select.innerHTML = baseOption;
    select.value = "";
    return;
  }
  const normalizedReferenceSite = normalizeText(referenceSite);
  const visibleSiteCount = selectedSites.filter((site) => site !== ALL_SITES_VALUE).length;
  const options = state.data.listes.referencesEffets
    .filter((reference) => {
      if (normalizedReferenceSite && !referenceHasSite(reference, normalizedReferenceSite)) {
        return false;
      }
      if (
        selectedSites.length &&
        !selectedSites.includes(ALL_SITES_VALUE) &&
        !selectedSites.some((site) => referenceHasSite(reference, site))
      ) {
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
    .map((reference) => {
      const label =
        !normalizedReferenceSite && visibleSiteCount > 1
          ? `${reference.designation} - ${getReferenceSiteLabel(reference)}`
          : reference.designation;
      return `<option value="${reference.id}">${label}</option>`;
    })
    .join("");
  select.innerHTML = `${baseOption}${options}`;
  if (selectedId) {
    select.value = selectedId;
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
  return state.data?.personnes?.find((person) => person.id === getCurrentPersonId()) || null;
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

function getOverdueExitMessage(person) {
  if (person?.dateSortiePrevue && !person?.dateSortieReelle && isPastDate(person.dateSortiePrevue)) {
    return `ALERTE : DATE DE SORTIE PREVUE DEPASSEE (${formatDate(person.dateSortiePrevue)})`;
  }
  if (person?.dateSortieReelle && isPastDate(person.dateSortieReelle)) {
    return `ALERTE : DATE DE SORTIE REELLE DEPASSEE (${formatDate(person.dateSortieReelle)})`;
  }
  return "";
}

function hasOverdueExit(person) {
  return Boolean(getOverdueExitMessage(person));
}

function isExitDue(person) {
  if (person?.dateSortieReelle) {
    return true;
  }
  if (person?.dateSortiePrevue && isPastDate(person.dateSortiePrevue)) {
    return true;
  }
  return false;
}

function getEffectStatus(person, effect) {
  if (effect.dateRemplacement || normalizeText(effect.statutManuel) === "REMPLACE") return "REMPLACE";
  if (effect.dateRetour) return "RESTITUE";

  const manualStatus = normalizeText(effect.statutManuel);
  if (["PERDU", "DETRUIT", "HS"].includes(manualStatus)) return manualStatus;
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
  renderRepresentativesTable(state.referenceRenderContext);
  renderReferenceEffectsTable(state.referenceRenderContext);
  renderReplacementCostsTable();
  renderReferenceCounts();
  renderMobileSignatureSettings();
}

function renderMobileSignatureSettings() {
  const input = document.querySelector('#mobile-signature-settings-form [name="mobileSignatureBaseUrl"]');
  const statusNode = document.getElementById("mobile-signature-settings-status");
  if (!(input instanceof HTMLInputElement)) {
    return;
  }

  const configured = getConfiguredMobileSignatureBaseUrl();
  input.value = configured;
  if (!statusNode) {
    return;
  }

  if (configured) {
    statusNode.textContent = `URL PUBLIQUE ACTIVE : ${configured}`;
    return;
  }

  const currentOrigin = normalizeHttpUrl(window.location.origin || "");
  if (currentOrigin && !isLikelyLocalUrl(currentOrigin)) {
    statusNode.textContent = `MODE AUTO HEBERGE : ${currentOrigin}`;
    return;
  }

  const autoBase = normalizeHttpUrl(state.mobileSignatureNetworkInfo?.preferredUrl || "");
  if (autoBase) {
    statusNode.textContent = `MODE AUTO RESEAU LOCAL : ${autoBase}`;
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
        listName === "sites"
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
  if (!references.length) {
    body.innerHTML = buildEmptyTableRow(body, "AUCUNE REFERENCE", 5);
    return;
  }

  const rowsHtml = references
    .map((reference) => {
      const usage = renderContext?.referenceEffectUsage?.get(String(reference.id || "")) || 0;
      return `<tr class="js-reference-effect-row" data-reference-id="${reference.id}">
        <td>${escapeHtml(getReferenceSiteLabel(reference))}</td>
        <td>${escapeHtml(reference.typeEffet)}</td>
        <td>${escapeHtml(reference.designation)}</td>
        <td>${usage}</td>
        <td>
          <button type="button" class="table-link js-edit-reference-effect" data-reference-id="${reference.id}">MODIFIER</button>
          <button type="button" class="table-link js-delete-reference-effect" data-reference-id="${reference.id}">SUPPRIMER</button>
        </td>
      </tr>`;
    });

  renderTableRowsProgressively(body, rowsHtml, buildEmptyTableRow(body, "AUCUNE REFERENCE", 5), 24);

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
        </td>
      </tr>`;
    });

  renderTableRowsProgressively(body, rowsHtml, buildEmptyTableRow(body, "AUCUN COUT", 4), 24);

  bindReplacementCostActions();
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
  const mapping = {
    "reference-count-sites": state.data?.listes?.sites?.length || 0,
    "reference-count-typesPersonnel": state.data?.listes?.typesPersonnel?.length || 0,
    "reference-count-typesContrats": state.data?.listes?.typesContrats?.length || 0,
    "reference-count-fonctions": state.data?.listes?.fonctions?.length || 0,
    "reference-count-typesEffets": state.data?.listes?.typesEffets?.length || 0,
    "reference-count-referencesEffets": state.data?.listes?.referencesEffets?.length || 0,
    "reference-count-coutsRemplacement": state.data?.listes?.coutsRemplacement?.length || 0,
    "reference-count-representantsSignataires": state.data?.listes?.representantsSignataires?.length || 0,
  };

  Object.entries(mapping).forEach(([id, value]) => {
    const node = document.getElementById(id);
    if (node) {
      node.textContent = String(value);
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
  state.data.listes[listName] = list.filter((entry) => normalizeText(entry) !== value);
  if (
    state.editingSimpleReference &&
    state.editingSimpleReference.listName === listName &&
    normalizeText(state.editingSimpleReference.originalValue) === value
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

  const usage = getReferenceEffectUsage(referenceId);
  if (usage > 0) {
    showDataStatus("SUPPRESSION BLOQUEE - REFERENCE DEJA UTILISEE");
    window.alert("SUPPRESSION IMPOSSIBLE : REFERENCE DEJA UTILISEE");
    return;
  }

  const reference = findReferenceById(referenceId);
  const confirmDelete = window.confirm(
    `SUPPRIMER DEFINITIVEMENT : ${reference?.designation || referenceId} ?`
  );
  if (!confirmDelete) {
    return;
  }
  pushUndoSnapshot("SUPPRESSION REFERENCE");
  state.data.listes.referencesEffets = state.data.listes.referencesEffets.filter(
    (entry) => entry.id !== referenceId
  );
  if (state.editingReferenceId === referenceId) {
    state.editingReferenceId = "";
    resetReferenceEffectForm();
  }
  markDirty();
  renderPage();
  showActionStatus("delete", `REFERENCE SUPPRIMEE : ${reference?.designation || referenceId}`);
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
  const normalizedValue = normalizeText(value);
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
  return 0;
}

function getReferenceEffectUsage(referenceId) {
  return getAllEffects(state.data?.personnes || []).filter(
    ({ effect }) => String(effect.referenceEffetId || "") === String(referenceId)
  ).length;
}

function cascadeSimpleReferenceRename(listName, oldValue, newValue) {
  const oldNormalized = normalizeText(oldValue);
  const nextValue = normalizeText(newValue);

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

async function saveDataToFile(options = {}) {
  if (!state.data) {
    showDataStatus("AUCUNE DONNEE A SAUVEGARDER");
    return;
  }

  const {
    silent = false,
    reloadAfter = true,
    successText = "data.json MIS A JOUR",
  } = options;

  try {
    const mode = getDataBackendMode();
    if (mode === "SUPABASE") {
      await saveSupabaseStateData(state.data);
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
    } else {
      throw new Error("SUPABASE NON CONFIGURE");
    }

    clearWorkingData();
    state.isDirty = false;
    clearUndoStack();
    renderDirtyState();
    showDataStatus(mode === "SUPABASE" ? "DONNEES SUPABASE SAUVEGARDEES" : successText);
    if (!silent) {
      window.alert(mode === "SUPABASE" ? "DONNEES SUPABASE MISES A JOUR" : "data.json A ETE MIS A JOUR");
    }
    if (reloadAfter) {
      await reloadData(mode === "SUPABASE" ? "RELECTURE DES DONNEES SUPABASE..." : "RELECTURE DE data.json...");
    }
  } catch (error) {
    console.error(error);
    showDataStatus("SAUVEGARDE IMPOSSIBLE");
    if (!silent) {
      window.alert("SAUVEGARDE IMPOSSIBLE");
    }
  }
}

function resetUiWithoutData() {
  const targets = [
    { id: "overview-table-body", colspan: 9 },
    { id: "global-table-body", colspan: 12 },
    { id: "sheet-effects-body", colspan: 11 },
    { id: "reference-sites-body", colspan: 2 },
    { id: "reference-typesPersonnel-body", colspan: 2 },
    { id: "reference-typesContrats-body", colspan: 2 },
    { id: "reference-fonctions-body", colspan: 2 },
    { id: "reference-typesEffets-body", colspan: 2 },
    { id: "reference-effects-table-body", colspan: 5 },
    { id: "replacement-costs-body", colspan: 4 },
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
  saveWorkingData();
  renderDirtyState();
}

function renderDirtyState() {
  const node = document.getElementById("dirty-status");
  if (!node) {
    return;
  }
  node.hidden = false;
  node.textContent = state.isDirty ? "MODIFICATIONS NON SAUVEGARDEES" : "DONNEES SAUVEGARDEES";
  node.classList.toggle("is-saved", !state.isDirty);
}

loadData();
  const getSheetTargetPersonId = () => state.currentSheetPersonId || getCurrentPersonId();
