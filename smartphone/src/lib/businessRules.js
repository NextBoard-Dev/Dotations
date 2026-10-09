export function normalizeText(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toUpperCase();
}

const ALLOWED_MANUAL_EFFECT_STATUSES = new Set(["ACTIF", "PERDU", "VOL", "HS", "DETRUIT"]);
const BILLABLE_EFFECT_CAUSES = new Set(["PERTE", "VOL", "NON RENDU", "DETRUIT"]);
const NON_RENDU_REFERENCE_COSTS = {
  "BADGE INTRUSION": 15,
  "CARTE TURBOSELF": 10,
  CLE: 5,
  "CLE CES": 50,
  "TELECOMMANDE URMET": 40,
};

function normalizeCause(rawCause) {
  const cause = normalizeText(rawCause);
  if (cause === "CASSE") return "DETRUIT";
  if (cause === "PERDU") return "PERTE";
  if (["DETRUIT", "PERTE", "VOL", "HS", "NON RENDU"].includes(cause)) return cause;
  return "";
}

export function normalizeEffectCause(rawCause) {
  return normalizeCause(rawCause);
}

function normalizePricingKey(value, { cause = false } = {}) {
  let normalized = normalizeText(value)
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!cause) return normalized;
  if (normalized === "NON RENDU") return "NON RENDU";
  if (normalized === "PERDU") return "PERTE";
  if (normalized === "CASSE") return "DETRUIT";
  return normalized;
}

function getFallbackNonRenduCost(typeEffet, designation = "") {
  const normalizedType = normalizePricingKey(typeEffet);
  if (!normalizedType) return 0;
  if (normalizedType === "CLE") return isCesKeyDesignation(designation) ? 50 : 5;
  return NON_RENDU_REFERENCE_COSTS[normalizedType] || 0;
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
  if (status === "CASSE") return "DETRUIT";
  if (ALLOWED_MANUAL_EFFECT_STATUSES.has(status)) return status;
  return "";
}

export function getEffectStatus(person, effect) {
  if (String(effect?.dateRetour || "").trim()) return "RESTITUE";
  if (normalizeText(effect?.etatFacturation) === "CLOTURE") return "CLOTURE";
  const manualStatus = normalizeManualStatus(effect?.statutManuel || effect?.statut);
  if (["PERDU", "HS", "VOL", "DETRUIT"].includes(manualStatus)) return manualStatus;
  if (isExitDue(person)) return "NON RENDU";
  return manualStatus || "ACTIF";
}

export function getEffectBillingCause(person, effect) {
  const persistedCause = normalizeCause(effect?.cause || effect?.causeRemplacement);
  if (persistedCause) return persistedCause;
  if (!String(effect?.dateRetour || "").trim() && isExitDue(person)) return "NON RENDU";
  return "";
}

export function getEffectMovement(person, effect) {
  const status = normalizeText(getEffectStatus(person, effect));
  const cause = normalizeText(getEffectBillingCause(person, effect));
  if (status === "RESTITUE") return "RENDU";
  if (status === "CLOTURE") return "CLOTURE";
  if (status === "DETRUIT") return "DETRUIT";
  if (status === "VOL" || cause === "VOL") return "VOLE";
  if (status === "HS") return "HS";
  if (status === "PERDU" || cause === "PERTE") return "PERDU";
  if (status === "NON RENDU") return "NON RENDU";
  return "";
}

export function getEffectBillingStatus(effect, isChargeable) {
  const stored = normalizeText(effect?.etatFacturation || "");
  if (stored === "FACTURE") return "FACTURE";
  if (stored === "CLOTURE") return "CLOTURE";
  return isChargeable ? "A FACTURER" : "-";
}

export function getEffectChargeableAmount(person, effect, pricingRules = []) {
  const cause = normalizeText(getEffectBillingCause(person, effect));
  if (!cause) return 0;
  return getReplacementCostValue(pricingRules, effect?.typeEffet, cause, effect?.designation || "");
}

export function getTotalChargeableAmount(person, effects = [], pricingRules = []) {
  const list = Array.isArray(effects) ? effects : [];
  return list.reduce((sum, effect) => sum + getEffectChargeableAmount(person, effect, pricingRules), 0);
}

export function getReplacementCostValue(pricingRules = [], typeEffet, cause, designation = "") {
  const wantedType = normalizePricingKey(typeEffet);
  const wantedCause = normalizePricingKey(cause, { cause: true });
  if (!wantedType || !wantedCause || wantedCause === "HS") return 0;
  if (wantedType === "CLE CES") {
    return BILLABLE_EFFECT_CAUSES.has(wantedCause) ? 50 : 0;
  }
  const row = (pricingRules || []).find((entry) => {
    const ruleType = normalizePricingKey(entry?.typeEffet);
    const ruleCause = normalizePricingKey(entry?.cause, { cause: true });
    return ruleType === wantedType && ruleCause === wantedCause;
  });
  if (!row) {
    if (wantedCause === "NON RENDU") return getFallbackNonRenduCost(wantedType, designation);
    return 0;
  }
  if (!BILLABLE_EFFECT_CAUSES.has(wantedCause)) return 0;
  if (wantedType === "CLE") {
    return isCesKeyDesignation(designation) ? 50 : 5;
  }
  const rawAmount = String(row?.montant ?? "").trim();
  const normalizedRawAmount = rawAmount.replace(/\s/g, "").replace(",", ".");
  const hasValidNumericAmount = /^-?\d+(\.\d+)?$/.test(normalizedRawAmount);
  const amount = Number(row?.montant);
  if (!rawAmount || !hasValidNumericAmount) {
    console.warn("Tarif invalide detecte dans coutsRemplacement", wantedType, wantedCause, row?.montant);
    if (wantedCause === "NON RENDU") return getFallbackNonRenduCost(wantedType, designation);
  }
  return Number.isFinite(amount) ? amount : 0;
}

export function formatDateFr(value) {
  const raw = String(value || "").trim();
  const match = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) return raw;
  return `${match[3]}/${match[2]}/${match[1]}`;
}

export function isCurrentAssignedEffect(person, effect) {
  const status = normalizeText(getEffectStatus(person, effect));
  return !["RESTITUE", "PERDU", "HS", "DETRUIT", "VOL", "CLOTURE"].includes(status);
}

function isDocumentFullySigned(person, docType) {
  const personnelSignature = String(
    person?.signatures?.[docType]?.personnel?.value ||
      person?.signatures?.[docType]?.personnel?.image ||
      ""
  ).trim();
  const representantSignature = String(
    person?.signatures?.[docType]?.representant?.value ||
      person?.signatures?.[docType]?.representant?.image ||
      ""
  ).trim();
  const personnelDate = String(person?.signatures?.[docType]?.personnel?.validatedAt || "").trim();
  const representantDate = String(person?.signatures?.[docType]?.representant?.validatedAt || "").trim();
  return Boolean(personnelSignature && representantSignature && personnelDate && representantDate);
}

function getLatestSignatureMs(person, docType) {
  const p = Date.parse(String(person?.signatures?.[docType]?.personnel?.validatedAt || "")) || 0;
  const r = Date.parse(String(person?.signatures?.[docType]?.representant?.validatedAt || "")) || 0;
  return Math.max(p, r);
}

function getExpectedDocumentFingerprint(person, docType) {
  const direct = String(person?.documentFingerprints?.[docType] || "").trim();
  if (direct) return direct;
  const signatureScoped = String(person?.signatures?.[docType]?.fingerprint || "").trim();
  if (signatureScoped) return signatureScoped;
  return "";
}

function hasSignedArchiveFor(person, typeLabel, documentsArchives = [], latestSignatureMs = 0, expectedFingerprint = "") {
  const archives = (documentsArchives || []).filter((entry) => {
    if (String(entry?.personId || "") !== String(person?.id || "")) return false;
    if (normalizeText(entry?.typeDocument) !== normalizeText(typeLabel)) return false;
    if (!String(entry?.pdfPath || "").trim()) return false;
    return true;
  });
  if (!archives.length) return false;
  if (expectedFingerprint) {
    const fingerprintMatch = archives.some((entry) => String(entry?.fingerprint || "").trim() === expectedFingerprint);
    if (fingerprintMatch) return true;
  }
  const signedArchives = archives.filter((entry) => normalizeText(entry?.statutSignature) === "SIGNE");
  if (!signedArchives.length) return false;
  if (!latestSignatureMs) return true;
  return signedArchives.some((entry) => (Date.parse(String(entry?.dateArchivage || "")) || 0) >= latestSignatureMs);
}

export function buildUiOverviewAlerts(persons = [], documentsArchives = [], todayIso = new Date().toISOString().slice(0, 10)) {
  const people = Array.isArray(persons)
    ? persons.filter(
        (person) =>
          person &&
          String(person?.id || "").trim() &&
          person?.is_deleted !== true &&
          person?.isDeleted !== true
      )
    : [];
  const docs = Array.isArray(documentsArchives)
    ? documentsArchives.filter(
        (entry) =>
          entry &&
          String(entry?.personId || "").trim() &&
          String(entry?.typeDocument || "").trim() &&
          String(entry?.pdfPath || "").trim()
      )
    : [];
  const alerts = [];
  const todayDate = new Date(`${todayIso}T00:00:00`);

  people.forEach((person) => {
    const plannedExit = String(person?.dateSortiePrevue || "");
    const realExit = String(person?.dateSortieReelle || "");
    const plannedDate = plannedExit ? new Date(`${plannedExit}T00:00:00`) : null;
    const daysUntilPlannedExit =
      plannedDate && Number.isFinite(plannedDate.getTime())
        ? Math.round((plannedDate.getTime() - todayDate.getTime()) / (24 * 60 * 60 * 1000))
        : null;

    if (plannedExit && !realExit && Number.isFinite(daysUntilPlannedExit) && daysUntilPlannedExit >= 1 && daysUntilPlannedExit <= 2) {
      alerts.push({
        personId: person.id,
        type: "dateSortiePrevue",
        text: `ALERTE : SORTIE PREVUE DANS ${daysUntilPlannedExit} JOUR${daysUntilPlannedExit > 1 ? "S" : ""} (${formatDateFr(plannedExit)})`,
      });
    } else if (plannedExit && !realExit && plannedExit <= todayIso) {
      alerts.push({
        personId: person.id,
        type: "dateSortiePrevue",
        text:
          plannedExit === todayIso
            ? `ALERTE : SORTIE PREVUE AUJOURD'HUI (${formatDateFr(plannedExit)})`
            : `ALERTE : DATE DE SORTIE PREVUE DEPASSEE (${formatDateFr(plannedExit)})`,
      });
    } else {
      const effects = Array.isArray(person?.effetsConfies) ? person.effetsConfies : [];
      const hasNonRendu = effects.some((effect) => getEffectStatus(person, effect) === "NON RENDU");
      if (realExit && realExit <= todayIso && hasNonRendu) {
        alerts.push({
          personId: person.id,
          type: "dateSortieReelle",
          text:
            realExit === todayIso
              ? `ALERTE : SORTIE REELLE AUJOURD'HUI AVEC EFFETS NON RENDUS (${formatDateFr(realExit)})`
              : `ALERTE : DATE DE SORTIE REELLE DEPASSEE (${formatDateFr(realExit)})`,
        });
      }
    }

    const currentEffectsCount = (Array.isArray(person?.effetsConfies) ? person.effetsConfies : []).filter((effect) =>
      isCurrentAssignedEffect(person, effect)
    ).length;
    const arrivalSigned = isDocumentFullySigned(person, "arrival");
    const exitSigned = isDocumentFullySigned(person, "exit");
    const hasArrivalPdf = hasSignedArchiveFor(
      person,
      "ARRIVEE",
      docs,
      getLatestSignatureMs(person, "arrival"),
      getExpectedDocumentFingerprint(person, "arrival")
    );
    const hasExitPdf = hasSignedArchiveFor(
      person,
      "SORTIE",
      docs,
      getLatestSignatureMs(person, "exit"),
      getExpectedDocumentFingerprint(person, "exit")
    );

    if (currentEffectsCount > 0 && !arrivalSigned) {
      alerts.push({
        personId: person.id,
        type: "signaturePdf",
        text: "ALERTE : EFFET(S) ATTRIBUE(S) MAIS DOCUMENT D'ARRIVEE NON SIGNE (PERSONNEL + REPRESENTANT OBLIGATOIRES).",
      });
    } else if (arrivalSigned && !hasArrivalPdf) {
      alerts.push({
        personId: person.id,
        type: "signaturePdf",
        text: 'ALERTE : ARRIVEE SIGNEE (2 SIGNATURES), MAIS PDF ABSENT OU NON MIS A JOUR POUR CETTE VERSION. RE-SIGNEZ LES DEUX PARTIES PUIS CLIQUEZ SUR "GENERER LE PDF".',
      });
    }
    if (exitSigned && !hasExitPdf) {
      alerts.push({
        personId: person.id,
        type: "signaturePdf",
        text: 'ALERTE : SORTIE SIGNEE (2 SIGNATURES), MAIS PDF ABSENT OU NON MIS A JOUR POUR CETTE VERSION. RE-SIGNEZ LES DEUX PARTIES PUIS CLIQUEZ SUR "GENERER LE PDF".',
      });
    }
  });

  return alerts.map((entry, index) => ({ key: `ui-${entry.personId}-${entry.type}-${index}`, ...entry }));
}
