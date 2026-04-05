import { supabase } from "@/lib/supabaseClient";
import { normalizeManualStatus } from "@/lib/businessRules";

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

  try {
    const localFlag = parseReadonlyFlag(window.localStorage.getItem("smartphone_readonly"));
    if (localFlag !== null) return localFlag;
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

function toString(value) {
  return value == null ? "" : String(value);
}

function cleanDate(value) {
  const v = toString(value).trim();
  return v || "";
}

function normalizeCause(value) {
  const normalized = toString(value).trim().toUpperCase();
  if (normalized === "CASSE") return "HS";
  if (normalized === "PERDU") return "PERTE";
  if (["DETRUIT", "PERTE", "VOL", "HS"].includes(normalized)) return normalized;
  return "";
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
  return ensureArray(payload.personnes).map(normalizeLegacyPerson);
}

function listLegacyEffetsFromPayload(payload = {}) {
  const out = [];
  ensureArray(payload.personnes).forEach((person) => {
    ensureArray(person?.effetsConfies).forEach((effet) => {
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

async function getAppStateRow() {
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

async function saveAppStatePayload(payload, expectedRevision) {
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
    try {
      await getAppStateRow();
    } catch {}
    throw buildAppStateConflictError();
  }
  const nextRevision = Number(updated.revision);
  return Number.isFinite(nextRevision) ? nextRevision : revision + 1;
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
        typeEffet: toString(entry?.typeEffet).trim(),
        cause: toString(entry?.cause).trim().toUpperCase(),
        montant: Number(entry?.montant) || 0,
      }))
      .filter((entry) => entry.typeEffet);
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

function toSqlSignatureCreatePayload(data = {}) {
  return {
    personId: toString(data?.personId),
    docType: toString(data?.docType),
    signer: toString(data?.signer),
    signatureData: toString(data?.signatureData),
    signedAt: cleanDate(data?.signedAt),
    signataireName: toString(data?.signataireName).trim(),
    signataireFunction: toString(data?.signataireFunction).trim(),
  };
}

function toSqlSignatureUpdatePayload(data = {}) {
  const out = {};
  if (Object.prototype.hasOwnProperty.call(data, "signatureData")) out.signatureData = toString(data.signatureData);
  if (Object.prototype.hasOwnProperty.call(data, "signedAt")) out.signedAt = cleanDate(data.signedAt);
  if (Object.prototype.hasOwnProperty.call(data, "signataireName")) out.signataireName = toString(data.signataireName).trim();
  if (Object.prototype.hasOwnProperty.call(data, "signataireFunction")) out.signataireFunction = toString(data.signataireFunction).trim();
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
  if (Object.prototype.hasOwnProperty.call(data, "nom")) out.nom = toString(data.nom).trim();
  if (Object.prototype.hasOwnProperty.call(data, "prenom")) out.prenom = toString(data.prenom).trim();
  if (Object.prototype.hasOwnProperty.call(data, "fonction")) out.fonction = toString(data.fonction).trim();
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

  const signatures = await sqlSignature.filter({ personId: toString(personId), docType: normalizedDocType });
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
    const { column, ascending } = normalizeOrder(order);
    const safeLimit = Number.isFinite(Number(limit)) ? Math.max(1, Math.min(Number(limit), 1000)) : 200;
    return runQuery(
      supabase.from(table).select("*").order(column, { ascending }).limit(safeLimit),
      `Lecture ${table} impossible`
    );
  },
  async filter(filters = {}) {
    let query = supabase.from(table).select("*");
    Object.entries(filters || {}).forEach(([key, value]) => {
      query = query.eq(key, value);
    });
    return runQuery(query, `Filtre ${table} impossible`);
  },
  async create(data) {
    const normalized = normalizeDates(table, data);
    return runQuery(
      supabase.from(table).insert(normalized).select().single(),
      `Creation ${table} impossible`
    );
  },
  async update(id, data) {
    const normalized = normalizeDates(table, data);
    return runQuery(
      supabase.from(table).update(normalized).eq("id", id).select().single(),
      `Mise a jour ${table} impossible`
    );
  },
  async delete(id) {
    const { error } = await supabase.from(table).delete().eq("id", id);
    if (error) {
      const details = [error.message, error.details, error.hint].filter(Boolean).join(" | ");
      throw new Error(`Suppression ${table} impossible: ${details || "Erreur inconnue"}`);
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
      if (legacy) return sortAndLimit(legacy.persons, order, limit);
      return sqlPerson.list(order, limit);
    },
    async filter(filters = {}) {
      const legacy = await tryLegacyData();
      if (legacy) return legacy.persons.filter((p) => matchesFilters(p, filters));
      return sqlPerson.filter(filters);
    },
    async create(data) {
      ensureWritable('creation personne');
      const row = await getAppStateRow();
      const payload = row?.payload;
      if (payload && Array.isArray(payload.personnes)) {
        const raw = payload.personnes;
        const newId = nextNumericId(raw.map((p) => p?.id), "P", 4);
        const sites = normalizeStringArray(data?.sites);
        const entry = {
          id: newId,
          nom: toString(data?.nom).trim(),
          prenom: toString(data?.prenom).trim(),
          fonction: toString(data?.fonction).trim(),
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
      return sqlPerson.create(data);
    },
    async update(id, data) {
      ensureWritable('mise a jour personne');
      const row = await getAppStateRow();
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
      return sqlPerson.update(id, data);
    },
    async delete(id) {
      ensureWritable('suppression personne');
      const row = await getAppStateRow();
      const payload = row?.payload;
      if (payload && Array.isArray(payload.personnes)) {
        const before = payload.personnes.length;
        payload.personnes = payload.personnes.filter((p) => toString(p?.id) !== toString(id));
        if (payload.personnes.length !== before) {
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
      if (legacy) return sortAndLimit(legacy.effets, order, limit);
      return sqlEffet.list(order, limit);
    },
    async filter(filters = {}) {
      const legacy = await tryLegacyData();
      if (legacy) return legacy.effets.filter((e) => matchesFilters(e, filters));
      return sqlEffet.filter(filters);
    },
    async create(data) {
      ensureWritable('creation effet');
      const row = await getAppStateRow();
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
      ensureWritable('mise a jour effet');
      const row = await getAppStateRow();
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
      ensureWritable('suppression effet');
      const row = await getAppStateRow();
      const payload = row?.payload;
      if (payload && Array.isArray(payload.personnes)) {
        for (const person of payload.personnes) {
          const before = ensureArray(person?.effetsConfies).length;
          person.effetsConfies = ensureArray(person?.effetsConfies).filter((e) => toString(e?.id) !== toString(id));
          if (person.effetsConfies.length !== before) {
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
        const sigs = extractLegacySignatures(legacy.payload, {});
        return sortAndLimit(sigs, order, limit);
      }
      return sqlSignature.list(order, limit);
    },
    async filter(filters = {}) {
      const legacy = await tryLegacyData();
      if (legacy) return extractLegacySignatures(legacy.payload, filters);
      return sqlSignature.filter(filters);
    },
    async create(data) {
      ensureWritable('creation signature');
      const row = await getAppStateRow();
      const payload = row?.payload;
      const personId = toString(data?.personId);
      const docType = toString(data?.docType);
      const signer = toString(data?.signer);
      if (payload && Array.isArray(payload.personnes) && personId) {
        const person = payload.personnes.find((p) => toString(p?.id) === personId);
        if (!person) throw new Error("Creation signature impossible: personne introuvable");
        const sigRoot = { ...defaultLegacySignatures(), ...(person.signatures || {}) };
        const prev = sigRoot?.[docType]?.[signer] || {};
        sigRoot[docType] = sigRoot[docType] || {};
        sigRoot[docType][signer] = {
          ...prev,
          image: toString(data?.signatureData),
          validatedAt: cleanDate(data?.signedAt) || new Date().toISOString(),
          signataireId: toString(data?.signataireId),
          signataireName: toString(data?.signataireName).trim(),
          signataireFunction: toString(data?.signataireFunction).trim(),
        };
        if (signer === "representant") {
          if (!person.representants || typeof person.representants !== "object") person.representants = {};
          if (!person.representants[docType] || typeof person.representants[docType] !== "object") person.representants[docType] = {};
          person.representants[docType].id = toString(data?.signataireId);
          person.representants[docType].nom = toString(data?.signataireName).trim();
          person.representants[docType].fonction = toString(data?.signataireFunction).trim();
        }
        person.signatures = sigRoot;
        applySignedDocumentCompletion(person, docType);
        await saveAppStatePayload(payload, row?.revision);
        return normalizeLegacySignature(personId, docType, signer, sigRoot[docType][signer]);
      }
      const created = await sqlSignature.create(toSqlSignatureCreatePayload(data));
      await applySqlPersonCompletionFromSignatures(data?.personId, data?.docType);
      return created;
    },
    async update(id, data) {
      ensureWritable('mise a jour signature');
      const row = await getAppStateRow();
      const payload = row?.payload;
      if (payload && Array.isArray(payload.personnes) && isLegacySignatureId(id)) {
        const { personId, docType, signer } = parseLegacySignatureId(id);
        const person = payload.personnes.find((p) => toString(p?.id) === personId);
        if (!person) throw new Error("Mise a jour signature impossible: personne introuvable");
        const sigRoot = { ...defaultLegacySignatures(), ...(person.signatures || {}) };
        const prev = sigRoot?.[docType]?.[signer] || {};
        sigRoot[docType] = sigRoot[docType] || {};
        sigRoot[docType][signer] = {
          ...prev,
          image: Object.prototype.hasOwnProperty.call(data || {}, "signatureData") ? toString(data.signatureData) : toString(prev.image),
          validatedAt: Object.prototype.hasOwnProperty.call(data || {}, "signedAt") ? cleanDate(data.signedAt) : cleanDate(prev.validatedAt),
          signataireId: Object.prototype.hasOwnProperty.call(data || {}, "signataireId") ? toString(data.signataireId) : toString(prev.signataireId),
          signataireName: Object.prototype.hasOwnProperty.call(data || {}, "signataireName") ? toString(data.signataireName).trim() : toString(prev.signataireName).trim(),
          signataireFunction: Object.prototype.hasOwnProperty.call(data || {}, "signataireFunction") ? toString(data.signataireFunction).trim() : toString(prev.signataireFunction).trim(),
        };
        if (signer === "representant") {
          if (!person.representants || typeof person.representants !== "object") person.representants = {};
          if (!person.representants[docType] || typeof person.representants[docType] !== "object") person.representants[docType] = {};
          person.representants[docType].id = Object.prototype.hasOwnProperty.call(data || {}, "signataireId")
            ? toString(data.signataireId)
            : toString(person.representants[docType].id);
          person.representants[docType].nom = Object.prototype.hasOwnProperty.call(data || {}, "signataireName")
            ? toString(data.signataireName).trim()
            : toString(person.representants[docType].nom).trim();
          person.representants[docType].fonction = Object.prototype.hasOwnProperty.call(data || {}, "signataireFunction")
            ? toString(data.signataireFunction).trim()
            : toString(person.representants[docType].fonction).trim();
        }
        person.signatures = sigRoot;
        applySignedDocumentCompletion(person, docType);
        await saveAppStatePayload(payload, row?.revision);
        return normalizeLegacySignature(personId, docType, signer, sigRoot[docType][signer]);
      }
      const updated = await sqlSignature.update(id, toSqlSignatureUpdatePayload(data));
      await applySqlPersonCompletionFromSignatures(updated?.personId, updated?.docType);
      return updated;
    },
    async delete(id) {
      ensureWritable('suppression signature');
      const row = await getAppStateRow();
      const payload = row?.payload;
      if (payload && Array.isArray(payload.personnes) && isLegacySignatureId(id)) {
        const { personId, docType, signer } = parseLegacySignatureId(id);
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
