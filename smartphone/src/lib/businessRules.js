function normalizeText(value) {
  return String(value || "")
    .trim()
    .toUpperCase();
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
  if (status === "CASSE" || status === "DETRUIT") return "HS";
  return status;
}

export function getEffectStatus(person, effect) {
  if (String(effect?.dateRetour || "").trim()) return "RESTITUE";
  const manualStatus = normalizeManualStatus(effect?.statutManuel || effect?.statut);
  if (["PERDU", "HS", "VOL"].includes(manualStatus)) return manualStatus;
  if (isExitDue(person)) return "NON RENDU";
  return manualStatus || "ACTIF";
}

export function getEffectBillingCause(person, effect) {
  const status = getEffectStatus(person, effect);
  if (status === "PERDU") return "PERTE";
  if (status === "VOL") return "VOL";
  if (status === "NON RENDU") return "NON RENDU";
  if (status === "HS") return "HS";
  return "";
}

export function getReplacementCostValue(pricingRules = [], typeEffet, cause) {
  const wantedType = normalizeText(typeEffet);
  const wantedCause = normalizeText(cause);
  if (!wantedType || !wantedCause || wantedCause === "HS") return 0;
  const row = (pricingRules || []).find((entry) => {
    const ruleType = normalizeText(entry?.typeEffet);
    let ruleCause = normalizeText(entry?.cause);
    if (ruleCause === "PERDU") ruleCause = "PERTE";
    return ruleType === wantedType && ruleCause === wantedCause;
  });
  return Number(row?.montant) || 0;
}

