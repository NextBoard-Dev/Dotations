// @ts-nocheck
import { supabase } from "@/lib/supabaseClient";
import { normalizeManualStatus, normalizeText } from "@/lib/businessRules";

const WRITE_ROLES = new Set(["editor", "admin"]);
const DELETE_ROLES = new Set(["admin"]);
const DOC_TYPES = new Set(["arrival", "exit"]);
const SIGNERS = new Set(["personnel", "representant"]);
const MAX_SIGNATURE_DATA_LENGTH = 2_000_000;
const EDGE_API_URL = "https://dphrvdhqhgycmllietuk.supabase.co/functions/v1/dotations-api";
const SUPABASE_PUBLISHABLE_KEY = "sb_publishable_2wYXnIDj4-c8daQZW8D5hA_2Py6k7z6";
const FILTERABLE_COLUMNS = {
  personnes: new Set(["id", "nom", "prenom", "fonction", "typePersonnel", "typeContrat", "dateEntree", "dateSortiePrevue", "dateSortieReelle", "statutDossier"]),
  effetsConfies: new Set(["id", "personId", "typeEffet", "designation", "siteReference", "numeroIdentification", "vehiculeImmatriculation", "dateRemise", "dateRetour", "statut", "cause", "dateRemplacement"]),
  signatures: new Set(["id", "personId", "person_id", "docType", "doc_type", "signer", "signedAt", "signed_at"]),
};
let authContextCache = null;
let authContextCacheTs = 0;
const AUTH_CONTEXT_TTL_MS = 5000;
let appStateSaveQueue = Promise.resolve();
let appStateRowCache = null;
let appStateRowCacheTs = 0;
let appStateRowFetchPromise = null;
const APP_STATE_READ_CACHE_TTL_MS = 8000;

async function fetchAuthContext({ force = false } = {}) {
  const now = Date.now();
  if (!force && authContextCache && now - authContextCacheTs < AUTH_CONTEXT_TTL_MS) {
    return authContextCache;
  }

  const { data, error } = await supabase.auth.getSession();
  if (error) {
    throw new Error(`Session Supabase indisponible: ${error.message || "Erreur inconnue"}`);
  }
  const session = data?.session || null;
  const userId = session?.user?.id || "";
  if (!userId) {
    throw new Error("Connexion requise pour acceder aux donnees.");
  }

  let role = "viewer";
  const { data: profile, error: profileError } = await supabase
    .from("profiles")
    .select("role")
    .eq("id", userId)
    .limit(1)
    .maybeSingle();

  if (profileError) {
    throw new Error(`Lecture role utilisateur impossible: ${profileError.message || "Erreur inconnue"}`);
  }

  if (profile?.role) {
    role = String(profile.role).trim().toLowerCase() || "viewer";
  }

  authContextCache = { userId, role, session };
  authContextCacheTs = now;
  return authContextCache;
}

async function requireAuthenticated(actionLabel = "operation") {
  try {
    return await fetchAuthContext();
  } catch (error) {
    throw new Error(`${actionLabel}: ${error.message || "Connexion requise."}`);
  }
}

async function requireRole(actionLabel, allowedRoles) {
  const ctx = await requireAuthenticated(actionLabel);
  if (allowedRoles.has(ctx.role)) {
    return ctx;
  }
  throw new Error(`${actionLabel}: role '${ctx.role}' non autorise.`);
}

function parseReadonlyFlag(value) {
  const raw = String(value ?? "").trim().toLowerCase();
  if (!raw) return null;
  if (raw === "false" || raw === "0" || raw === "off" || raw === "no") return false;
  if (raw === "true" || raw === "1" || raw === "on" || raw === "yes") return true;
  return null;
}

function resolveReadonlyMode() {
  const envFlag = parseReadonlyFlag(import.meta?.env?.VITE_SMARTPHONE_READONLY);
  if (envFlag !== null) return envFlag;

  try {
    const queryFlag = parseReadonlyFlag(new URLSearchParams(window.location.search).get("readonly"));
    if (queryFlag !== null) return queryFlag;
  } catch {}

  return false;
}

const READ_ONLY_MODE = resolveReadonlyMode();

function ensureWritable(actionLabel = "operation") {
  if (!READ_ONLY_MODE) return;
  throw new Error(`Sauvegarde Supabase temporairement bloquee: ${actionLabel}.`);
}

function normalizeOrder(order) {
  const raw = String(order || "-created_at");
  const ascending = !raw.startsWith("-");
  const column = raw.replace(/^-/, "") || "created_at";
  return { column, ascending };
}

function normalizeDates(table, payload = {}) {
  const data = { ...(payload || {}) };
  const fieldsByTable = {
    personnes: ["dateEntree", "dateSortiePrevue", "dateSortieReelle"],
    effetsConfies: ["dateRemise", "dateRetour", "dateRemplacement"],
    signatures: ["signedAt"],
  };

  (fieldsByTable[table] || []).forEach((field) => {
    if (Object.prototype.hasOwnProperty.call(data, field) && data[field] === "") {
      data[field] = null;
    }
  });

  if (table === "personnes" && Object.prototype.hasOwnProperty.call(data, "sites")) {
    data.sites = Array.isArray(data.sites) ? data.sites : [];
  }

  return data;
}

async function runQuery(queryPromise, contextText) {
  const { data, error } = await queryPromise;
  if (error) {
    const details = [error.message, error.details, error.hint].filter(Boolean).join(" | ");
    throw new Error(`${contextText}: ${details || "Erreur inconnue"}`);
  }
  return data;
}

function normalizeStringArray(value) {
  if (!Array.isArray(value)) return [];
  return value.map((entry) => String(entry || "").trim()).filter(Boolean);
}

function safeJsonParse(raw) {
  if (raw == null) return {};
  if (typeof raw === "string") {
    try {
      return JSON.parse(raw);
    } catch {
      return {};
    }
  }
  return typeof raw === "object" ? raw : {};
}

function cloneJson(value) {
  if (!value || typeof value !== "object") return value;
  try {
    return JSON.parse(JSON.stringify(value));
  } catch {
    return value;
  }
}

function cloneAppStateRow(row) {
  if (!row || typeof row !== "object") return row || null;
  return {
    id: row.id,
    payload: cloneJson(row.payload || {}),
    revision: Number(row.revision) || 0,
  };
}

function extractPayloadFromEdgeData(raw) {
  if (!raw || typeof raw !== "object") return null;
  if (raw.payload && typeof raw.payload === "object") return raw.payload;
  if (raw.data && typeof raw.data === "object") {
    if (raw.data.payload && typeof raw.data.payload === "object") return raw.data.payload;
    if (Array.isArray(raw.data.personnes) || raw.data.listes) return raw.data;
  }
  if (Array.isArray(raw.personnes) || raw.listes) return raw;
  return null;
}

function extractRevisionFromEdgeData(raw) {
  const direct = Number(raw?.revision);
  if (Number.isFinite(direct)) return direct;
  const nested = Number(raw?.data?.revision);
  if (Number.isFinite(nested)) return nested;
  return 0;
}

async function getAppStateRowFromEdge(session) {
  const token = String(session?.access_token || "").trim();
  if (!token) return null;
  const response = await fetch(`${EDGE_API_URL}/data`, {
    method: "GET",
    headers: {
      apikey: SUPABASE_PUBLISHABLE_KEY,
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    cache: "no-store",
  });
  if (!response.ok) return null;
  const raw = await response.json().catch(() => null);
  const payload = extractPayloadFromEdgeData(raw);
  if (!payload) return null;
  return {
    id: "main",
    payload: safeJsonParse(payload),
    revision: extractRevisionFromEdgeData(raw),
  };
}

function toString(value) {
  return value == null ? "" : String(value);
}

function isSoftDeleted(entry) {
  if (!entry || typeof entry !== "object") return false;
  const rawIsDeleted = entry.is_deleted ?? entry.isDeleted;
  const deletedFlag =
    rawIsDeleted === true ||
    String(rawIsDeleted || "").trim().toLowerCase() === "true" ||
    String(rawIsDeleted || "").trim() === "1";
  const deletedAt = String(entry.deleted_at || entry.deletedAt || "").trim();
  return deletedFlag || Boolean(deletedAt);
}

function markSoftDeleted(entry = {}, deletedBy = "") {
  const nowIso = new Date().toISOString();
  return {
    ...(entry || {}),
    is_deleted: true,
    isDeleted: true,
    deleted_at: nowIso,
    deletedAt: nowIso,
    updated_at: nowIso,
    updatedAt: nowIso,
    updated_by: deletedBy || null,
    updatedBy: deletedBy || null,
  };
}

function ensureValidDocType(value) {
  const docType = toString(value).trim();
  if (!DOC_TYPES.has(docType)) {
    throw new Error(`docType invalide: '${docType || "vide"}'`);
  }
  return docType;
}

function ensureValidSigner(value) {
  const signer = toString(value).trim();
  if (!SIGNERS.has(signer)) {
    throw new Error(`signer invalide: '${signer || "vide"}'`);
  }
  return signer;
}

function normalizeSignatureData(value) {
  const signatureData = toString(value).trim();
  if (!signatureData) return "";
  if (signatureData.length > MAX_SIGNATURE_DATA_LENGTH) {
    throw new Error("Signature trop volumineuse.");
  }
  if (!/^data:image\/(png|jpeg|jpg);base64,[a-z0-9+/=]+$/i.test(signatureData)) {
    throw new Error("Format de signature invalide.");
  }
  return signatureData;
}

function sanitizeFilterObject(table, filters = {}) {
  const allowed = FILTERABLE_COLUMNS[table];
  if (!allowed) return {};
  const out = {};
  Object.entries(filters || {}).forEach(([key, value]) => {
    if (allowed.has(key)) out[key] = value;
  });
  return out;
}

function cleanDate(value) {
  const v = toString(value).trim();
  return v || "";
}

function getField(row = {}, ...names) {
  for (const name of names) {
    if (Object.prototype.hasOwnProperty.call(row, name) && row[name] != null) {
      return row[name];
    }
  }
  return "";
}

function normalizeSignatureRecord(row = {}) {
  const personId = toString(getField(row, "personId", "person_id")).trim();
  const docType = toString(getField(row, "docType", "doc_type")).trim();
  const signer = toString(getField(row, "signer")).trim();
  const signedAt = cleanDate(getField(row, "signedAt", "signed_at", "validatedAt", "validated_at_text", "updatedAt", "updated_at"));
  return {
    ...row,
    id: toString(getField(row, "id")).trim() || `${personId}__${docType}__${signer}`,
    personId,
    docType,
    signer,
    signatureData: toString(getField(row, "signatureData", "signature_data", "image")),
    signedAt,
    signataireId: toString(getField(row, "signataireId", "signataire_id", "signer_id")).trim(),
    signataireName: toString(getField(row, "signataireName", "signataire_name", "signer_name")).trim(),
    signataireFunction: toString(getField(row, "signataireFunction", "signataire_function", "signer_function")).trim(),
    storageRef: toString(getField(row, "storageRef", "storage_ref")),
    storagePublicUrl: toString(getField(row, "storagePublicUrl", "storage_public_url")),
    token: toString(getField(row, "token")).trim(),
  };
}

function signatureRank(signature = {}) {
  const direct = Date.parse(signature.signedAt || "");
  if (Number.isFinite(direct)) return direct;
  const updated = Date.parse(signature.updatedAt || signature.updated_at || "");
  if (Number.isFinite(updated)) return updated;
  const created = Date.parse(signature.createdAt || signature.created_at || "");
  return Number.isFinite(created) ? created : 0;
}

function mergeSignatureRecords(...groups) {
  const byKey = new Map();
  groups.flat().map(normalizeSignatureRecord).forEach((signature) => {
    if (!signature.personId || !signature.docType || !signature.signer) return;
    if (!signature.signatureData && !signature.signedAt && !signature.signataireName && !signature.signataireFunction) return;
    const key = `${signature.personId}__${signature.docType}__${signature.signer}`;
    const current = byKey.get(key);
    if (!current || signatureRank(signature) >= signatureRank(current)) {
      byKey.set(key, signature);
    }
  });
  return Array.from(byKey.values());
}

async function fetchSqlSignatureRecords(filters = {}) {
  await requireAuthenticated("Lecture signatures");
  let query = supabase.from("signatures").select("*").eq("is_deleted", false);
  const personId = toString(filters.personId ?? filters.person_id).trim();
  const docType = toString(filters.docType ?? filters.doc_type).trim();
  const signer = toString(filters.signer).trim();
  if (personId) query = query.eq("person_id", personId);
  if (docType) query = query.eq("doc_type", docType);
  if (signer) query = query.eq("signer", signer);
  const rows = await runQuery(query.order("updated_at", { ascending: false }).limit(1000), "Lecture signatures impossible");
  return ensureArray(rows).map(normalizeSignatureRecord);
}

function normalizeCause(value) {
  const normalized = toString(value).trim().toUpperCase();
  if (normalized === "CASSE") return "DETRUIT";
  if (normalized === "PERDU") return "PERTE";
  if (["DETRUIT", "PERTE", "VOL", "HS"].includes(normalized)) return normalized;
  return "";
}

function normalizePricingKey(value, { cause = false } = {}) {
  let normalized = toString(value)
    .trim()
    .toUpperCase()
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!cause) return normalized;
  if (normalized === "NON RENDU") return "NON RENDU";
  if (normalized === "PERDU") return "PERTE";
  if (normalized === "CASSE") return "DETRUIT";
  return normalized;
}

function inferCauseFromStatus(value) {
  const status = normalizeManualStatus(value);
  if (status === "PERDU") return "PERTE";
  if (status === "DETRUIT") return "DETRUIT";
  if (status === "VOL") return "VOL";
  if (status === "HS") return "HS";
  return "";
}

function getTodayIsoDate() {
  return new Date().toISOString().slice(0, 10);
}

function ensureArray(value) {
  return Array.isArray(value) ? value : [];
}

function computeStatutDossier(person) {
  const dateSortieReelle = cleanDate(person?.dateSortieReelle);
  if (dateSortieReelle) return "SORTI";
  const dateSortiePrevue = cleanDate(person?.dateSortiePrevue);
  if (dateSortiePrevue) return "SORTIE PREVUE";
  return "EN POSTE";
}

function normalizeLegacyPerson(person = {}) {
  const sites = normalizeStringArray(
    ensureArray(person.sitesAffectation).length ? person.sitesAffectation : [person.site]
  );
  return {
    id: toString(person.id),
    nom: toString(person.nom).trim(),
    prenom: toString(person.prenom).trim(),
    fonction: toString(person.fonction).trim(),
    sites,
    typePersonnel: toString(person.typePersonnel).trim(),
    typeContrat: toString(person.typeContrat).trim(),
    dateEntree: cleanDate(person.dateEntree),
    dateSortiePrevue: cleanDate(person.dateSortiePrevue),
    dateSortieReelle: cleanDate(person.dateSortieReelle),
    statutDossier: computeStatutDossier(person),
  };
}

function normalizeLegacyEffet(effet = {}, personId = "") {
  const normalizedStatus = normalizeManualStatus(effet.statutManuel || effet.statut || "ACTIF") || "ACTIF";
  return {
    id: toString(effet.id),
    personId: toString(personId),
    typeEffet: toString(effet.typeEffet).trim(),
    designation: toString(effet.designation).trim(),
    siteReference: toString(effet.siteReference).trim(),
    numeroIdentification: toString(effet.numeroIdentification).trim(),
    vehiculeImmatriculation: toString(effet.vehiculeImmatriculation).trim(),
    dateRemise: cleanDate(effet.dateRemise),
    dateRetour: cleanDate(effet.dateRetour),
    statut: normalizedStatus,
    cause: normalizeCause(effet.cause || effet.causeRemplacement),
    dateRemplacement: cleanDate(effet.dateRemplacement),
    coutRemplacement: Number(effet.coutRemplacement) || 0,
    commentaire: toString(effet.commentaire).trim(),
    mouvement: toString(effet.mouvement).trim(),
  };
}

function listLegacyPersonsFromPayload(payload = {}) {
  return ensureArray(payload.personnes)
    .filter((person) => !isSoftDeleted(person))
    .map(normalizeLegacyPerson);
}

function listLegacyEffetsFromPayload(payload = {}) {
  const out = [];
  ensureArray(payload.personnes).forEach((person) => {
    if (isSoftDeleted(person)) return;
    ensureArray(person?.effetsConfies).forEach((effet) => {
      if (isSoftDeleted(effet)) return;
      out.push(normalizeLegacyEffet(effet, person?.id));
    });
  });
  return out;
}

function sortAndLimit(items = [], order = "-created_at", limit = 200) {
  const { column, ascending } = normalizeOrder(order);
  const safeLimit = Number.isFinite(Number(limit)) ? Math.max(1, Math.min(Number(limit), 10000)) : 200;
  const sorted = [...items].sort((a, b) => {
    const av = a?.[column];
    const bv = b?.[column];
    const an = av == null ? "" : String(av);
    const bn = bv == null ? "" : String(bv);
    const cmp = an.localeCompare(bn, "fr", { numeric: true, sensitivity: "base" });
    return ascending ? cmp : -cmp;
  });
  return sorted.slice(0, safeLimit);
}

function nextNumericId(existingIds = [], prefix = "P", pad = 4) {
  let maxNum = 0;
  existingIds.forEach((id) => {
    const txt = toString(id).trim();
    const m = txt.match(new RegExp(`^${prefix}(\\d+)$`, "i"));
    if (!m) return;
    const n = Number(m[1]);
    if (Number.isFinite(n) && n > maxNum) maxNum = n;
  });
  return `${prefix}${String(maxNum + 1).padStart(pad, "0")}`;
}

function defaultLegacySignatures() {
  return {
    arrival: {
      personnel: { image: "", validatedAt: "", signataireName: "", signataireFunction: "" },
      representant: { image: "", validatedAt: "", signataireName: "", signataireFunction: "" },
    },
    exit: {
      personnel: { image: "", validatedAt: "", signataireName: "", signataireFunction: "" },
      representant: { image: "", validatedAt: "", signataireName: "", signataireFunction: "" },
    },
  };
}

function isLegacyDocFullySigned(person, docType) {
  const personnelDate = cleanDate(person?.signatures?.[docType]?.personnel?.validatedAt);
  const representantDate = cleanDate(person?.signatures?.[docType]?.representant?.validatedAt);
  return Boolean(personnelDate && representantDate);
}

function applySignedDocumentCompletion(person, docType) {
  if (!person || !isLegacyDocFullySigned(person, docType)) return false;
  if (docType === "arrival" && !cleanDate(person.dateEntree)) {
    person.dateEntree = getTodayIsoDate();
    return true;
  }
  if (docType === "exit" && !cleanDate(person.dateSortieReelle)) {
    person.dateSortieReelle = getTodayIsoDate();
    return true;
  }
  return false;
}

async function fetchAppStateRowFresh(ctx) {
  try {
    const edgeRow = await getAppStateRowFromEdge(ctx?.session);
    if (edgeRow?.payload) return edgeRow;
  } catch {}
  const rows = await runQuery(
    supabase.from("app_state").select("id,payload,revision").eq("id", "main").limit(1),
    "Lecture app_state impossible"
  );
  const row = rows?.[0];
  if (!row) return null;
  const revision = Number(row.revision);
  return {
    id: row.id,
    payload: safeJsonParse(row.payload),
    revision: Number.isFinite(revision) ? revision : 0,
  };
}

async function getAppStateRow({ forceFresh = false } = {}) {
  const ctx = await requireAuthenticated("Lecture app_state");
  const now = Date.now();
  if (!forceFresh && appStateRowCache && now - appStateRowCacheTs <= APP_STATE_READ_CACHE_TTL_MS) {
    return cloneAppStateRow(appStateRowCache);
  }
  if (!forceFresh && appStateRowFetchPromise) {
    return cloneAppStateRow(await appStateRowFetchPromise);
  }

  appStateRowFetchPromise = fetchAppStateRowFresh(ctx)
    .then((row) => {
      appStateRowCache = cloneAppStateRow(row);
      appStateRowCacheTs = Date.now();
      return row;
    })
    .finally(() => {
      appStateRowFetchPromise = null;
    });

  return cloneAppStateRow(await appStateRowFetchPromise);
}

async function getAppStateRowDirect() {
  await requireAuthenticated("Lecture app_state");
  const rows = await runQuery(
    supabase.from("app_state").select("id,payload,revision").eq("id", "main").limit(1),
    "Lecture app_state impossible"
  );
  const row = rows?.[0];
  if (!row) return null;
  const revision = Number(row.revision);
  return {
    id: row.id,
    payload: safeJsonParse(row.payload),
    revision: Number.isFinite(revision) ? revision : 0,
  };
}

function buildAppStateConflictError() {
  const error = new Error("Conflit de sauvegarde : les donnees ont ete modifiees ailleurs. Recharge puis reessaie.");
  error.code = "APP_STATE_CONFLICT";
  return error;
}

async function saveAppStatePayloadOnce(payload, expectedRevision, attempt = 0) {
  const normalizedPayload = payload && typeof payload === "object" ? payload : {};
  const revision = Number(expectedRevision);
  if (!Number.isFinite(revision)) {
    throw new Error("Mise a jour app_state impossible: revision indisponible");
  }
  const { data, error } = await supabase
    .from("app_state")
    .update({ payload: normalizedPayload, revision: revision + 1 })
    .eq("id", "main")
    .eq("revision", revision)
    .select("id,revision");
  if (error) {
    const details = [error.message, error.details, error.hint].filter(Boolean).join(" | ");
    throw new Error(`Mise a jour app_state impossible: ${details || "Erreur inconnue"}`);
  }
  const updated = Array.isArray(data) ? data[0] : null;
  if (!updated?.id) {
    if (attempt < 4) {
      try {
        const latestRow = await getAppStateRowDirect();
        const latestRevision = Number(latestRow?.revision);
        if (Number.isFinite(latestRevision) && latestRevision !== revision) {
          await new Promise((resolve) => setTimeout(resolve, 120));
          return saveAppStatePayloadOnce(normalizedPayload, latestRevision, attempt + 1);
        }
      } catch {}
    }
    throw buildAppStateConflictError();
  }
  const nextRevision = Number(updated.revision);
  appStateRowCache = cloneAppStateRow({ id: "main", payload: normalizedPayload, revision: nextRevision });
  appStateRowCacheTs = Date.now();
  return Number.isFinite(nextRevision) ? nextRevision : revision + 1;
}

async function saveAppStatePayload(payload, expectedRevision) {
  // Serialize app_state writes on mobile to avoid local concurrent revision conflicts.
  appStateSaveQueue = appStateSaveQueue
    .catch(() => {})
    .then(() => saveAppStatePayloadOnce(payload, expectedRevision, 0));
  return appStateSaveQueue;
}

async function forceSaveAppStatePayload(payload) {
  const normalizedPayload = payload && typeof payload === "object" ? payload : {};
  const latestRow = await getAppStateRowDirect();
  const latestRevision = Number(latestRow?.revision);
  const nextRevision = Number.isFinite(latestRevision) ? latestRevision + 1 : 1;
  const { data, error } = await supabase
    .from("app_state")
    .update({ payload: normalizedPayload, revision: nextRevision })
    .eq("id", "main")
    .select("id,revision");
  if (error) {
    const details = [error.message, error.details, error.hint].filter(Boolean).join(" | ");
    throw new Error(`Mise a jour forcee app_state impossible: ${details || "Erreur inconnue"}`);
  }
  const updated = Array.isArray(data) ? data[0] : null;
  if (!updated?.id) {
    throw new Error("Mise a jour forcee app_state impossible: ligne introuvable");
  }
  const persistedRevision = Number(updated.revision) || nextRevision;
  appStateRowCache = cloneAppStateRow({ id: "main", payload: normalizedPayload, revision: persistedRevision });
  appStateRowCacheTs = Date.now();
  return persistedRevision;
}

async function getAppStatePayload() {
  const row = await getAppStateRow();
  return row?.payload || {};
}

function listRepresentantsSignatairesFromPayload(payload = {}) {
  const reps = payload?.listes?.representantsSignataires;
  if (!Array.isArray(reps)) return [];
  return reps
    .map((entry, index) => ({
      id: String(entry?.id || `REP${String(index + 1).padStart(4, "0")}`),
      nom: String(entry?.nom || "").trim(),
      fonction: String(entry?.fonction || "").trim(),
    }))
    .filter((entry) => entry.nom || entry.fonction)
    .sort((a, b) => `${a.nom} ${a.fonction}`.localeCompare(`${b.nom} ${b.fonction}`, "fr"));
}

async function listRepresentantsSignataires() {
  try {
    const payload = await getAppStatePayload();
    return listRepresentantsSignatairesFromPayload(payload);
  } catch {
    return [];
  }
}

async function getReferenceBases() {
  try {
    const payload = await getAppStatePayload();
    const listes = payload?.listes || {};
    const coutsRemplacement = ensureArray(listes.coutsRemplacement)
      .map((entry) => ({
        typeEffet: normalizePricingKey(entry?.typeEffet),
        cause: normalizePricingKey(entry?.cause, { cause: true }),
        montant: Number(entry?.montant) || 0,
      }))
      .filter((entry) => entry.typeEffet && entry.cause);
    return {
      sites: normalizeStringArray(listes.sites),
      fonctions: normalizeStringArray(listes.fonctions),
      typesPersonnel: normalizeStringArray(listes.typesPersonnel),
      typesContrats: normalizeStringArray(listes.typesContrats),
      typesEffets: normalizeStringArray(listes.typesEffets),
      statutsObjetManuels: normalizeStringArray(listes.statutsObjetManuels),
      coutsRemplacement,
      referencesEffets: ensureArray(listes.referencesEffets)
        .map((entry) => ({
          id: toString(entry?.id),
          site: toString(entry?.site).trim(),
          typeEffet: toString(entry?.typeEffet).trim(),
          designation: toString(entry?.designation).trim(),
        }))
        .filter((entry) => entry.designation),
      representantsSignataires: listRepresentantsSignatairesFromPayload(payload),
    };
  } catch {
    return {
      sites: [],
      fonctions: [],
      typesPersonnel: [],
      typesContrats: [],
      typesEffets: [],
      statutsObjetManuels: [],
      coutsRemplacement: [],
      referencesEffets: [],
      representantsSignataires: [],
    };
  }
}

async function getLegacyOperationalData() {
  const payload = await getAppStatePayload();
  return {
    payload,
    persons: listLegacyPersonsFromPayload(payload),
    effets: listLegacyEffetsFromPayload(payload),
    bases: await getReferenceBases(),
  };
}

function matchesFilters(item = {}, filters = {}) {
  return Object.entries(filters || {}).every(([key, value]) => item?.[key] === value);
}

function isLegacySignatureId(id = "") {
  return String(id).includes("__");
}

function parseLegacySignatureId(id = "") {
  const [personId, docType, signer] = String(id).split("__");
  return { personId, docType, signer };
}

function buildLegacySignatureId(personId, docType, signer) {
  return `${personId}__${docType}__${signer}`;
}

function normalizeLegacySignature(personId, docType, signer, raw = {}) {
  const signatureData = toString(raw?.image);
  const signedAt = cleanDate(raw?.validatedAt);
  const signataireId = toString(raw?.signataireId);
  const signataireName = toString(raw?.signataireName).trim();
  const signataireFunction = toString(raw?.signataireFunction).trim();
  return {
    id: buildLegacySignatureId(personId, docType, signer),
    personId,
    docType,
    signer,
    signatureData,
    signedAt,
    signataireId,
    signataireName,
    signataireFunction,
  };
}

function upsertLegacySignatureOnPerson(person, { docType, signer, data = {} }) {
  const sigRoot = { ...defaultLegacySignatures(), ...(person?.signatures || {}) };
  const prev = sigRoot?.[docType]?.[signer] || {};
  sigRoot[docType] = sigRoot[docType] || {};
  sigRoot[docType][signer] = {
    ...prev,
    image: Object.prototype.hasOwnProperty.call(data || {}, "signatureData")
      ? normalizeSignatureData(data.signatureData)
      : toString(prev.image),
    validatedAt: Object.prototype.hasOwnProperty.call(data || {}, "signedAt")
      ? cleanDate(data.signedAt) || new Date().toISOString()
      : cleanDate(prev.validatedAt),
    signataireId: Object.prototype.hasOwnProperty.call(data || {}, "signataireId")
      ? toString(data.signataireId)
      : toString(prev.signataireId),
    signataireName: Object.prototype.hasOwnProperty.call(data || {}, "signataireName")
      ? toString(data.signataireName).trim()
      : toString(prev.signataireName).trim(),
    signataireFunction: Object.prototype.hasOwnProperty.call(data || {}, "signataireFunction")
      ? toString(data.signataireFunction).trim()
      : toString(prev.signataireFunction).trim(),
  };
  if (signer === "representant") {
    if (!person.representants || typeof person.representants !== "object") person.representants = {};
    if (!person.representants[docType] || typeof person.representants[docType] !== "object") person.representants[docType] = {};
    person.representants[docType].id = toString(sigRoot[docType][signer].signataireId);
    person.representants[docType].nom = toString(sigRoot[docType][signer].signataireName).trim();
    person.representants[docType].fonction = toString(sigRoot[docType][signer].signataireFunction).trim();
  }
  person.signatures = sigRoot;
  applySignedDocumentCompletion(person, docType);
  return sigRoot[docType][signer];
}

async function saveLegacySignatureWithConflictRetry({ personId, docType, signer, data = {} }) {
  const normalizedPersonId = toString(personId);
  ensureValidDocType(docType);
  ensureValidSigner(signer);
  let lastConflict = null;
  for (let attempt = 0; attempt < 6; attempt += 1) {
    const row = await getAppStateRowDirect();
    const payload = row?.payload;
    if (!payload || !Array.isArray(payload.personnes)) {
      throw new Error("Sauvegarde signature impossible: source legacy indisponible");
    }
    const person = payload.personnes.find((p) => toString(p?.id) === normalizedPersonId);
    if (!person) throw new Error("Sauvegarde signature impossible: personne introuvable");
    const savedRaw = upsertLegacySignatureOnPerson(person, { docType, signer, data });
    try {
      await saveAppStatePayloadOnce(payload, row?.revision, 0);
      return normalizeLegacySignature(normalizedPersonId, docType, signer, savedRaw);
    } catch (error) {
      if (error?.code === "APP_STATE_CONFLICT") {
        lastConflict = error;
        await new Promise((resolve) => setTimeout(resolve, 80));
        continue;
      }
      throw error;
    }
  }
  try {
    const row = await getAppStateRowDirect();
    const payload = row?.payload;
    if (!payload || !Array.isArray(payload.personnes)) {
      throw lastConflict || buildAppStateConflictError();
    }
    const person = payload.personnes.find((p) => toString(p?.id) === normalizedPersonId);
    if (!person) throw new Error("Sauvegarde signature impossible: personne introuvable");
    const savedRaw = upsertLegacySignatureOnPerson(person, { docType, signer, data });
    await forceSaveAppStatePayload(payload);
    return normalizeLegacySignature(normalizedPersonId, docType, signer, savedRaw);
  } catch {
    throw lastConflict || buildAppStateConflictError();
  }
}

function toSqlSignatureCreatePayload(data = {}) {
  const docType = ensureValidDocType(data?.docType);
  const signer = ensureValidSigner(data?.signer);
  return {
    person_id: toString(data?.personId),
    doc_type: docType,
    signer,
    signature_data: normalizeSignatureData(data?.signatureData),
    signed_at: cleanDate(data?.signedAt) || new Date().toISOString(),
    signer_name: toString(data?.signataireName).trim(),
    signer_function: toString(data?.signataireFunction).trim(),
  };
}

function toSqlSignatureUpdatePayload(data = {}) {
  const out = {};
  if (Object.prototype.hasOwnProperty.call(data, "signatureData")) out.signature_data = normalizeSignatureData(data.signatureData);
  if (Object.prototype.hasOwnProperty.call(data, "signedAt")) out.signed_at = cleanDate(data.signedAt);
  if (Object.prototype.hasOwnProperty.call(data, "signataireName")) out.signer_name = toString(data.signataireName).trim();
  if (Object.prototype.hasOwnProperty.call(data, "signataireFunction")) out.signer_function = toString(data.signataireFunction).trim();
  return out;
}

function extractLegacySignatures(payload = {}, filters = {}) {
  const personIdFilter = filters?.personId;
  const docTypeFilter = filters?.docType;
  const signerFilter = filters?.signer;

  const out = [];
  ensureArray(payload.personnes).forEach((person) => {
    const personId = toString(person?.id);
    if (personIdFilter && personId !== personIdFilter) return;

    const sigRoot = { ...defaultLegacySignatures(), ...(person?.signatures || {}) };
    ["arrival", "exit"].forEach((docType) => {
      if (docTypeFilter && docType !== docTypeFilter) return;
      ["personnel", "representant"].forEach((signer) => {
        if (signerFilter && signer !== signerFilter) return;
        const raw = sigRoot?.[docType]?.[signer] || {};
        const normalized = normalizeLegacySignature(personId, docType, signer, raw);
        if (normalized.signatureData || normalized.signedAt || normalized.signataireName || normalized.signataireFunction) {
          out.push(normalized);
        }
      });
    });
  });
  return out;
}

function applyLegacyPersonToRaw(rawPerson = {}, data = {}) {
  const out = { ...rawPerson };
  if (Object.prototype.hasOwnProperty.call(data, "nom")) out.nom = normalizeText(data.nom);
  if (Object.prototype.hasOwnProperty.call(data, "prenom")) out.prenom = normalizeText(data.prenom);
  if (Object.prototype.hasOwnProperty.call(data, "fonction")) out.fonction = normalizeText(data.fonction);
  if (Object.prototype.hasOwnProperty.call(data, "typePersonnel")) out.typePersonnel = toString(data.typePersonnel).trim();
  if (Object.prototype.hasOwnProperty.call(data, "typeContrat")) out.typeContrat = toString(data.typeContrat).trim();
  if (Object.prototype.hasOwnProperty.call(data, "dateEntree")) out.dateEntree = cleanDate(data.dateEntree);
  if (Object.prototype.hasOwnProperty.call(data, "dateSortiePrevue")) out.dateSortiePrevue = cleanDate(data.dateSortiePrevue);
  if (Object.prototype.hasOwnProperty.call(data, "dateSortieReelle")) out.dateSortieReelle = cleanDate(data.dateSortieReelle);
  if (Object.prototype.hasOwnProperty.call(data, "sites")) {
    const sites = normalizeStringArray(data.sites);
    out.sitesAffectation = sites;
    out.site = sites[0] || "";
  }
  return out;
}

function applyLegacyEffetToRaw(rawEffet = {}, data = {}) {
  const out = { ...rawEffet };
  const existingCause = normalizeCause(out.cause || out.causeRemplacement);
  if (existingCause) {
    out.cause = existingCause;
  }
  if (Object.prototype.hasOwnProperty.call(data, "typeEffet")) out.typeEffet = toString(data.typeEffet).trim();
  if (Object.prototype.hasOwnProperty.call(data, "designation")) out.designation = toString(data.designation).trim();
  if (Object.prototype.hasOwnProperty.call(data, "siteReference")) out.siteReference = toString(data.siteReference).trim();
  if (Object.prototype.hasOwnProperty.call(data, "numeroIdentification")) out.numeroIdentification = toString(data.numeroIdentification).trim();
  if (Object.prototype.hasOwnProperty.call(data, "vehiculeImmatriculation")) out.vehiculeImmatriculation = toString(data.vehiculeImmatriculation).trim();
  if (Object.prototype.hasOwnProperty.call(data, "dateRemise")) out.dateRemise = cleanDate(data.dateRemise);
  if (Object.prototype.hasOwnProperty.call(data, "dateRetour")) out.dateRetour = cleanDate(data.dateRetour);
  if (Object.prototype.hasOwnProperty.call(data, "dateRemplacement")) out.dateRemplacement = cleanDate(data.dateRemplacement);
  if (Object.prototype.hasOwnProperty.call(data, "commentaire")) out.commentaire = toString(data.commentaire).trim();
  if (Object.prototype.hasOwnProperty.call(data, "mouvement")) out.mouvement = toString(data.mouvement).trim();
  if (Object.prototype.hasOwnProperty.call(data, "coutRemplacement")) {
    const c = Number(data.coutRemplacement);
    out.coutRemplacement = Number.isFinite(c) ? c : 0;
  }
  if (Object.prototype.hasOwnProperty.call(data, "statut")) {
    out.statutManuel = normalizeManualStatus(data.statut) || "ACTIF";
    if (!Object.prototype.hasOwnProperty.call(data, "cause") && !normalizeCause(out.cause)) {
      const inferredCause = inferCauseFromStatus(data.statut);
      if (inferredCause) out.cause = inferredCause;
    }
  }
  if (Object.prototype.hasOwnProperty.call(data, "cause")) {
    const normalizedCause = normalizeCause(data.cause);
    if (normalizedCause) {
      out.cause = normalizedCause;
    }
  }
  return out;
}

async function tryLegacyData() {
  try {
    const row = await getAppStateRow();
    const payload = row?.payload || {};
    const persons = listLegacyPersonsFromPayload(payload);
    if (!persons.length) return null;
    return { payload, persons, effets: listLegacyEffetsFromPayload(payload) };
  } catch {
    return null;
  }
}

async function applySqlPersonCompletionFromSignatures(personId, docType) {
  const normalizedDocType = toString(docType).trim();
  if (!personId || (normalizedDocType !== "arrival" && normalizedDocType !== "exit")) return;

  const signatures = await fetchSqlSignatureRecords({ personId: toString(personId), docType: normalizedDocType });
  const isSigned = (signer) => signatures.some((s) => s.signer === signer && cleanDate(s.signedAt));
  if (!isSigned("personnel") || !isSigned("representant")) return;

  const persons = await sqlPerson.filter({ id: toString(personId) });
  const person = Array.isArray(persons) ? persons[0] : null;
  if (!person) return;

  if (normalizedDocType === "arrival" && !cleanDate(person.dateEntree)) {
    await sqlPerson.update(personId, { dateEntree: getTodayIsoDate() });
  } else if (normalizedDocType === "exit" && !cleanDate(person.dateSortieReelle)) {
    await sqlPerson.update(personId, { dateSortieReelle: getTodayIsoDate() });
  }
}

const makeEntity = (table) => ({
  async list(order = "-created_at", limit = 200) {
    await requireAuthenticated(`Lecture ${table}`);
    const { column, ascending } = normalizeOrder(order);
    const safeLimit = Number.isFinite(Number(limit)) ? Math.max(1, Math.min(Number(limit), 1000)) : 200;
    return runQuery(
      supabase.from(table).select("*").eq("is_deleted", false).order(column, { ascending }).limit(safeLimit),
      `Lecture ${table} impossible`
    );
  },
  async filter(filters = {}) {
    await requireAuthenticated(`Filtre ${table}`);
    const safeFilters = sanitizeFilterObject(table, filters);
    let query = supabase.from(table).select("*").eq("is_deleted", false);
    Object.entries(safeFilters).forEach(([key, value]) => {
      query = query.eq(key, value);
    });
    return runQuery(query, `Filtre ${table} impossible`);
  },
  async create(data) {
    await requireRole(`Creation ${table}`, WRITE_ROLES);
    const normalized = normalizeDates(table, data);
    return runQuery(
      supabase.from(table).insert(normalized).select().single(),
      `Creation ${table} impossible`
    );
  },
  async update(id, data) {
    await requireRole(`Mise a jour ${table}`, WRITE_ROLES);
    const normalized = normalizeDates(table, data);
    return runQuery(
      supabase.from(table).update(normalized).eq("id", id).select().single(),
      `Mise a jour ${table} impossible`
    );
  },
  async delete(id) {
    await requireRole(`Suppression ${table}`, DELETE_ROLES);
    const { error } = await supabase
      .from(table)
      .update({ is_deleted: true, updated_at: new Date().toISOString() })
      .eq("id", id)
      .eq("is_deleted", false);
    if (error) {
      const details = [error.message, error.details, error.hint].filter(Boolean).join(" | ");
      throw new Error(`Soft delete ${table} impossible: ${details || "Erreur inconnue"}`);
    }
    return { success: true };
  },
});

const sqlPerson = makeEntity("personnes");
const sqlEffet = makeEntity("effetsConfies");
const sqlSignature = makeEntity("signatures");

export const db = {
  Person: {
    async list(order = "-created_at", limit = 200) {
      const legacy = await tryLegacyData();
      if (legacy) return sortAndLimit(legacy.persons.filter((p) => !isSoftDeleted(p)), order, limit);
      return sqlPerson.list(order, limit);
    },
    async filter(filters = {}) {
      const legacy = await tryLegacyData();
      if (legacy) return legacy.persons.filter((p) => !isSoftDeleted(p) && matchesFilters(p, filters));
      return sqlPerson.filter(filters);
    },
    async create(data) {
      await requireRole("Creation personne", WRITE_ROLES);
      ensureWritable('creation personne');
      const row = await getAppStateRow({ forceFresh: true });
      const payload = row?.payload;
      if (payload && Array.isArray(payload.personnes)) {
        const raw = payload.personnes;
        const newId = nextNumericId(raw.map((p) => p?.id), "P", 4);
        const sites = normalizeStringArray(data?.sites);
        const entry = {
          id: newId,
          nom: normalizeText(data?.nom),
          prenom: normalizeText(data?.prenom),
          fonction: normalizeText(data?.fonction),
          site: sites[0] || "",
          sitesAffectation: sites,
          typePersonnel: toString(data?.typePersonnel).trim(),
          typeContrat: toString(data?.typeContrat).trim(),
          dateEntree: cleanDate(data?.dateEntree),
          dateSortiePrevue: cleanDate(data?.dateSortiePrevue),
          dateSortieReelle: cleanDate(data?.dateSortieReelle),
          signatures: defaultLegacySignatures(),
          effetsConfies: [],
          representants: { arrival: {}, exit: {} },
        };
        payload.personnes = [...raw, entry];
        await saveAppStatePayload(payload, row?.revision);
        return normalizeLegacyPerson(entry);
      }
      const normalizedData = {
        ...data,
        nom: normalizeText(data?.nom),
        prenom: normalizeText(data?.prenom),
        fonction: normalizeText(data?.fonction),
      };
      return sqlPerson.create(normalizedData);
    },
    async update(id, data) {
      await requireRole("Mise a jour personne", WRITE_ROLES);
      ensureWritable('mise a jour personne');
      const row = await getAppStateRow({ forceFresh: true });
      const payload = row?.payload;
      if (payload && Array.isArray(payload.personnes)) {
        const idx = payload.personnes.findIndex((p) => toString(p?.id) === toString(id));
        if (idx >= 0) {
          const updated = applyLegacyPersonToRaw(payload.personnes[idx], data || {});
          payload.personnes[idx] = updated;
          await saveAppStatePayload(payload, row?.revision);
          return normalizeLegacyPerson(updated);
        }
      }
      const normalizedData = { ...data };
      if (Object.prototype.hasOwnProperty.call(data || {}, "nom")) normalizedData.nom = normalizeText(data.nom);
      if (Object.prototype.hasOwnProperty.call(data || {}, "prenom")) normalizedData.prenom = normalizeText(data.prenom);
      if (Object.prototype.hasOwnProperty.call(data || {}, "fonction")) normalizedData.fonction = normalizeText(data.fonction);
      return sqlPerson.update(id, normalizedData);
    },
    async delete(id) {
      const ctx = await requireRole("Suppression personne", DELETE_ROLES);
      ensureWritable('suppression personne');
      const row = await getAppStateRow({ forceFresh: true });
      const payload = row?.payload;
      if (payload && Array.isArray(payload.personnes)) {
        const index = payload.personnes.findIndex((p) => toString(p?.id) === toString(id));
        if (index >= 0) {
          payload.personnes[index] = markSoftDeleted(payload.personnes[index], ctx?.userId || "");
          await saveAppStatePayload(payload, row?.revision);
          return { success: true };
        }
      }
      return sqlPerson.delete(id);
    },
  },

  Effet: {
    async list(order = "-created_at", limit = 200) {
      const legacy = await tryLegacyData();
      if (legacy) return sortAndLimit(legacy.effets.filter((e) => !isSoftDeleted(e)), order, limit);
      return sqlEffet.list(order, limit);
    },
    async filter(filters = {}) {
      const legacy = await tryLegacyData();
      if (legacy) return legacy.effets.filter((e) => !isSoftDeleted(e) && matchesFilters(e, filters));
      return sqlEffet.filter(filters);
    },
    async create(data) {
      await requireRole("Creation effet", WRITE_ROLES);
      ensureWritable('creation effet');
      const row = await getAppStateRow({ forceFresh: true });
      const payload = row?.payload;
      if (payload && Array.isArray(payload.personnes)) {
        const person = payload.personnes.find((p) => toString(p?.id) === toString(data?.personId));
        if (!person) throw new Error("Creation effet impossible: personne introuvable");
        const allEffetIds = payload.personnes.flatMap((p) => ensureArray(p?.effetsConfies).map((e) => e?.id));
        const newId = nextNumericId(allEffetIds, "E", 6);
        const entry = applyLegacyEffetToRaw({ id: newId, referenceEffetId: "", legacyDamageFacturable: false }, data || {});
        person.effetsConfies = [...ensureArray(person.effetsConfies), entry];
        await saveAppStatePayload(payload, row?.revision);
        return normalizeLegacyEffet(entry, person.id);
      }
      return sqlEffet.create(data);
    },
    async update(id, data) {
      await requireRole("Mise a jour effet", WRITE_ROLES);
      ensureWritable('mise a jour effet');
      const row = await getAppStateRow({ forceFresh: true });
      const payload = row?.payload;
      if (payload && Array.isArray(payload.personnes)) {
        for (const person of payload.personnes) {
          const list = ensureArray(person?.effetsConfies);
          const idx = list.findIndex((e) => toString(e?.id) === toString(id));
          if (idx >= 0) {
            const updated = applyLegacyEffetToRaw(list[idx], data || {});
            list[idx] = updated;
            person.effetsConfies = list;
            await saveAppStatePayload(payload, row?.revision);
            return normalizeLegacyEffet(updated, person.id);
          }
        }
      }
      return sqlEffet.update(id, data);
    },
    async delete(id) {
      const ctx = await requireRole("Suppression effet", DELETE_ROLES);
      ensureWritable('suppression effet');
      const row = await getAppStateRow({ forceFresh: true });
      const payload = row?.payload;
      if (payload && Array.isArray(payload.personnes)) {
        for (const person of payload.personnes) {
          const list = ensureArray(person?.effetsConfies);
          const index = list.findIndex((e) => toString(e?.id) === toString(id));
          if (index >= 0) {
            list[index] = markSoftDeleted(list[index], ctx?.userId || "");
            person.effetsConfies = list;
            await saveAppStatePayload(payload, row?.revision);
            return { success: true };
          }
        }
      }
      return sqlEffet.delete(id);
    },
  },

  Signature: {
    async list(order = "-created_at", limit = 200) {
      const legacy = await tryLegacyData();
      if (legacy) {
        const legacySigs = extractLegacySignatures(legacy.payload, {});
        let sqlSigs = [];
        try { sqlSigs = await fetchSqlSignatureRecords({}); } catch {}
        return sortAndLimit(mergeSignatureRecords(legacySigs, sqlSigs), order, limit);
      }
      const sqlSigs = await fetchSqlSignatureRecords({});
      return sortAndLimit(sqlSigs, order, limit);
    },
    async filter(filters = {}) {
      const legacy = await tryLegacyData();
      if (legacy) {
        const legacySigs = extractLegacySignatures(legacy.payload, filters);
        let sqlSigs = [];
        try { sqlSigs = await fetchSqlSignatureRecords(filters); } catch {}
        return mergeSignatureRecords(legacySigs, sqlSigs);
      }
      return fetchSqlSignatureRecords(filters);
    },
    async create(data) {
      await requireRole("Creation signature", WRITE_ROLES);
      ensureWritable('creation signature');
      const personId = toString(data?.personId);
      const docType = ensureValidDocType(data?.docType);
      const signer = ensureValidSigner(data?.signer);
      const signatureData = normalizeSignatureData(data?.signatureData);
      const row = await getAppStateRow({ forceFresh: true });
      const payload = row?.payload;
      if (payload && Array.isArray(payload.personnes) && personId) {
        return saveLegacySignatureWithConflictRetry({
          personId,
          docType,
          signer,
          data: {
            ...data,
            signatureData,
            signedAt: cleanDate(data?.signedAt) || new Date().toISOString(),
          },
        });
      }
      const created = await sqlSignature.create(toSqlSignatureCreatePayload(data));
      await applySqlPersonCompletionFromSignatures(data?.personId, data?.docType);
      return created;
    },
    async update(id, data) {
      await requireRole("Mise a jour signature", WRITE_ROLES);
      ensureWritable('mise a jour signature');
      const row = await getAppStateRow({ forceFresh: true });
      const payload = row?.payload;
      if (payload && Array.isArray(payload.personnes) && isLegacySignatureId(id)) {
        const { personId, docType, signer } = parseLegacySignatureId(id);
        return saveLegacySignatureWithConflictRetry({ personId, docType, signer, data });
      }
      const updated = await sqlSignature.update(id, toSqlSignatureUpdatePayload(data));
      const normalizedUpdated = normalizeSignatureRecord(updated || {});
      await applySqlPersonCompletionFromSignatures(normalizedUpdated.personId, normalizedUpdated.docType);
      return updated;
    },
    async delete(id) {
      await requireRole("Suppression signature", DELETE_ROLES);
      ensureWritable('suppression signature');
      const row = await getAppStateRow({ forceFresh: true });
      const payload = row?.payload;
      if (payload && Array.isArray(payload.personnes) && isLegacySignatureId(id)) {
        const { personId, docType, signer } = parseLegacySignatureId(id);
        ensureValidDocType(docType);
        ensureValidSigner(signer);
        const person = payload.personnes.find((p) => toString(p?.id) === personId);
        if (person) {
          const sigRoot = { ...defaultLegacySignatures(), ...(person.signatures || {}) };
          sigRoot[docType] = sigRoot[docType] || {};
          sigRoot[docType][signer] = { image: "", validatedAt: "", signataireName: "", signataireFunction: "" };
          person.signatures = sigRoot;
          await saveAppStatePayload(payload, row?.revision);
          return { success: true };
        }
      }
      return sqlSignature.delete(id);
    },
  },

  AppState: {
    getRepresentantsSignataires: listRepresentantsSignataires,
    getReferenceBases,
    getOperationalData: getLegacyOperationalData,
  },
};

