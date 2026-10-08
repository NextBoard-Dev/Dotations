import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";

const data = JSON.parse(fs.readFileSync("data.json", "utf8"));

function text(value) {
  return String(value || "").trim();
}

function normalize(value) {
  return text(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/\s+/g, " ");
}

function isValidSignatureImage(value) {
  const raw = text(value);
  return !raw || /^data:image\/(?:png|jpe?g);base64,/i.test(raw) || /^storage:\/\/[^/]+\/.+/i.test(raw) || /^https?:\/\//i.test(raw);
}

function isIsoDateTimeOrEmpty(value) {
  const raw = text(value);
  return !raw || /^\d{4}-\d{2}-\d{2}(?:[T ][0-2]\d:[0-5]\d(?::[0-5]\d(?:\.\d+)?)?(?:Z|[+-][0-2]\d:[0-5]\d)?)?$/.test(raw);
}

function isIsoDateOrEmpty(value) {
  const raw = text(value);
  return !raw || /^\d{4}-\d{2}-\d{2}$/.test(raw);
}

test("donnees metier: les identifiants personnes et effets sont globalement uniques", () => {
  const personIds = new Set();
  const effectIds = new Set();
  const errors = [];

  for (const person of data.personnes || []) {
    const personId = text(person?.id);
    if (!personId || personIds.has(personId)) errors.push(`personne en double: ${personId || "(vide)"}`);
    personIds.add(personId);

    for (const effect of Array.isArray(person?.effetsConfies) ? person.effetsConfies : []) {
      const effectId = text(effect?.id);
      if (!effectId || effectIds.has(effectId)) errors.push(`effet en double: ${effectId || "(vide)"}`);
      effectIds.add(effectId);
    }
  }

  assert.deepEqual(errors, []);
});

test("donnees metier: chaque effet reste rattache a son personnel parent", () => {
  const errors = [];

  for (const person of data.personnes || []) {
    const personId = text(person?.id);
    for (const effect of Array.isArray(person?.effetsConfies) ? person.effetsConfies : []) {
      const effectPersonId = text(effect?.personId || personId);
      if (effectPersonId !== personId) {
        errors.push(`${text(effect?.id) || "(effet sans id)"} rattache a ${effectPersonId}, parent ${personId}`);
      }
      if (!text(effect?.typeEffet)) {
        errors.push(`${text(effect?.id) || "(effet sans id)"} sans typeEffet`);
      }
      if (Number.isNaN(Number(effect?.quantite ?? 1))) {
        errors.push(`${text(effect?.id) || "(effet sans id)"} quantite invalide`);
      }
    }
  }

  assert.deepEqual(errors, []);
});

test("donnees metier: signatures et dates de validation restent exploitables", () => {
  const errors = [];

  for (const person of data.personnes || []) {
    const personId = text(person?.id);
    for (const docType of ["arrival", "exit"]) {
      for (const signer of ["personnel", "representant"]) {
        const signature = person?.signatures?.[docType]?.[signer] || {};
        if (!isValidSignatureImage(signature.image)) {
          errors.push(`${personId}.${docType}.${signer}.image invalide`);
        }
        if (!isIsoDateTimeOrEmpty(signature.validatedAt)) {
          errors.push(`${personId}.${docType}.${signer}.validatedAt invalide`);
        }
        if (!isValidSignatureImage(signature.storagePublicUrl)) {
          errors.push(`${personId}.${docType}.${signer}.storagePublicUrl invalide`);
        }
        const storageRef = text(signature.storageRef);
        if (storageRef && !/^storage:\/\/[^/]+\/.+/i.test(storageRef)) {
          errors.push(`${personId}.${docType}.${signer}.storageRef invalide`);
        }
      }
    }
  }

  assert.deepEqual(errors, []);
});

test("donnees metier: les dates restent au format exploitable et chronologique", () => {
  const errors = [];

  for (const person of data.personnes || []) {
    const personId = text(person?.id) || "(personne sans id)";
    for (const field of ["dateEntree", "dateSortiePrevue", "dateSortieReelle"]) {
      if (!isIsoDateOrEmpty(person?.[field])) {
        errors.push(`${personId}.${field} format invalide: ${text(person?.[field])}`);
      }
    }

    const dateEntree = text(person?.dateEntree);
    const dateSortieReelle = text(person?.dateSortieReelle);
    if (dateEntree && dateSortieReelle && dateSortieReelle < dateEntree) {
      errors.push(`${personId}.dateSortieReelle avant dateEntree`);
    }

    for (const effect of Array.isArray(person?.effetsConfies) ? person.effetsConfies : []) {
      const effectId = text(effect?.id) || "(effet sans id)";
      for (const field of ["dateRemise", "dateRetour", "dateRemplacement"]) {
        if (!isIsoDateOrEmpty(effect?.[field])) {
          errors.push(`${effectId}.${field} format invalide: ${text(effect?.[field])}`);
        }
      }

      const dateRemise = text(effect?.dateRemise);
      const dateRetour = text(effect?.dateRetour);
      if (dateRemise && dateRetour && dateRetour < dateRemise) {
        errors.push(`${effectId}.dateRetour avant dateRemise`);
      }
    }
  }

  assert.deepEqual(errors, []);
});

test("referentiels: references effets, representants et tarifs n'ont pas de doublons metier", () => {
  const errors = [];
  const referenceKeys = new Set();
  const representativeKeys = new Set();
  const costKeys = new Set();

  for (const reference of data.listes?.referencesEffets || []) {
    const key = [normalize(reference?.site), normalize(reference?.typeEffet), normalize(reference?.designation)].join("|");
    if (!text(reference?.id)) errors.push(`reference sans id: ${key}`);
    if (referenceKeys.has(key)) errors.push(`reference en double: ${key}`);
    referenceKeys.add(key);
  }

  for (const representative of data.listes?.representantsSignataires || []) {
    const key = [normalize(representative?.nom), normalize(representative?.fonction)].join("|");
    if (!text(representative?.id)) errors.push(`representant sans id: ${key}`);
    if (representativeKeys.has(key)) errors.push(`representant en double: ${key}`);
    representativeKeys.add(key);
  }

  for (const cost of data.listes?.coutsRemplacement || []) {
    const key = [normalize(cost?.typeEffet), normalize(cost?.cause)].join("|");
    if (!normalize(cost?.typeEffet) || !normalize(cost?.cause)) errors.push(`tarif incomplet: ${key}`);
    if (costKeys.has(key)) errors.push(`tarif en double: ${key}`);
    costKeys.add(key);
    if (!Number.isFinite(Number(cost?.montant)) || Number(cost?.montant) < 0) {
      errors.push(`tarif invalide: ${key}`);
    }
  }

  assert.deepEqual(errors, []);
});

test("archives: les entrees conservees pointent vers des personnes et documents valides", () => {
  const personIds = new Set((data.personnes || []).map((person) => text(person?.id)).filter(Boolean));
  const archiveIds = new Set();
  const errors = [];

  for (const archive of data.documentsArchives || []) {
    const archiveId = text(archive?.id);
    if (!archiveId || archiveIds.has(archiveId)) errors.push(`archive en double: ${archiveId || "(vide)"}`);
    archiveIds.add(archiveId);

    if (!personIds.has(text(archive?.personId))) errors.push(`${archiveId || "(archive sans id)"} personId inconnu`);
    if (!["ARRIVEE", "SORTIE"].includes(normalize(archive?.typeDocument))) errors.push(`${archiveId || "(archive sans id)"} typeDocument invalide`);
    if (!text(archive?.pdfPath) || !/\.pdf(?:$|\?)/i.test(text(archive?.pdfPath))) errors.push(`${archiveId || "(archive sans id)"} pdfPath invalide`);
    if (!isIsoDateTimeOrEmpty(archive?.dateArchivage)) errors.push(`${archiveId || "(archive sans id)"} dateArchivage invalide`);
  }

  assert.deepEqual(errors, []);
});
