function normalizeText(value) {
  return String(value || "")
    .trim()
    .toUpperCase();
}

const ALLOWED_MANUAL_EFFECT_STATUSES = new Set(["ACTIF", "PERDU", "VOL", "HS", "DETRUIT"]);
const BILLABLE_EFFECT_CAUSES = new Set(["PERTE", "VOL", "NON RENDU", "DETRUIT"]);

function normalizeCause(rawCause) {
  const cause = normalizeText(rawCause);
  if (cause === "CASSE") return "HS";
  if (cause === "PERDU") return "PERTE";
  if (["DETRUIT", "PERTE", "VOL", "HS", "NON RENDU"].includes(cause)) return cause;
  return "";
}

function isCesKeyDesignation(designation) {
  return normalizeText(designation).startsWith("CES-");
}

export function getDossierStatus(person) {
  if (String(person?.dateSortieReelle || "").trim()) return "SORTI";
  if (String(person?.dateSortiePrevue || "").trim()) return "SORTIE PREVUE";
  return "EN POSTE";
}

function isPastDate(value) {
  if (!value) return false;
  const today = new Date();
  const todayOnly = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const target = new Date(`${value}T00:00:00`);
  return Number.isFinite(target.getTime()) && target < todayOnly;
}

function isExitDue(person) {
  const todayIso = new Date().toISOString().slice(0, 10);
  const sortieReelle = String(person?.dateSortieReelle || "");
  if (sortieReelle && sortieReelle <= todayIso) return true;
  const sortiePrevue = String(person?.dateSortiePrevue || "");
  if (sortiePrevue && isPastDate(sortiePrevue)) return true;
  return false;
}

export function normalizeManualStatus(rawStatus) {
  const status = normalizeText(rawStatus);
  if (status === "VOLE") return "VOL";
  if (status === "CASSE") return "HS";
  if (ALLOWED_MANUAL_EFFECT_STATUSES.has(status)) return status;
  return "";
}

export function getEffectStatus(person, effect) {
  if (String(effect?.dateRetour || "").trim()) return "RESTITUE";
  const manualStatus = normalizeManualStatus(effect?.statutManuel || effect?.statut);
  if (["PERDU", "HS", "VOL"].includes(manualStatus)) return manualStatus;
  if (isExitDue(person)) return "NON RENDU";
  return manualStatus || "ACTIF";
}

export function getEffectBillingCause(person, effect) {
  const persistedCause = normalizeCause(effect?.cause || effect?.causeRemplacement);
  return persistedCause;
}

export function getReplacementCostValue(pricingRules = [], typeEffet, cause, designation = "") {
  const wantedType = normalizeText(typeEffet);
  const wantedCause = normalizeText(cause);
  if (!wantedType || !wantedCause || wantedCause === "HS") return 0;
  if (wantedType === "CLE CES") {
    return BILLABLE_EFFECT_CAUSES.has(wantedCause) ? 50 : 0;
  }
  const row = (pricingRules || []).find((entry) => {
    const ruleType = normalizeText(entry?.typeEffet);
    let ruleCause = normalizeText(entry?.cause);
    if (ruleCause === "PERDU") ruleCause = "PERTE";
    return ruleType === wantedType && ruleCause === wantedCause;
  });
  if (!row) return 0;
  if (!BILLABLE_EFFECT_CAUSES.has(wantedCause)) return 0;
  if (wantedType === "CLE") {
    return isCesKeyDesignation(designation) ? 50 : 5;
  }
  const amount = Number(row?.montant);
  return Number.isFinite(amount) ? amount : 0;
}
