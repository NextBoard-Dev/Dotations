import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";

import {
  buildUiOverviewAlerts,
  getDossierStatus,
  getEffectBillingCause,
  getEffectBillingStatus,
  getEffectChargeableAmount,
  getEffectStatus,
  getTotalChargeableAmount,
} from "../smartphone/src/lib/businessRules.js";

const pricingRules = [
  { typeEffet: "BADGE INTRUSION", cause: "NON RENDU", montant: 15 },
  { typeEffet: "BADGE INTRUSION", cause: "DETRUIT", montant: 15 },
  { typeEffet: "CARTE TURBOSELF", cause: "NON RENDU", montant: 10 },
  { typeEffet: "CLE", cause: "NON RENDU", montant: 5 },
  { typeEffet: "CLE CES", cause: "NON RENDU", montant: 50 },
  { typeEffet: "TELECOMMANDE URMET", cause: "NON RENDU", montant: 40 },
];

function signedSignature(date = "2026-10-07T10:00:00.000Z") {
  return { image: "data:image/png;base64,SIGNATURE", validatedAt: date };
}

function emptySignature() {
  return { image: "", validatedAt: "" };
}

function basePerson(overrides = {}) {
  return {
    id: "P-TEST",
    nom: "TEST",
    prenom: "REFERENCE",
    dateEntree: "2026-01-01",
    dateSortiePrevue: "",
    dateSortieReelle: "",
    signatures: {
      arrival: {
        personnel: emptySignature(),
        representant: emptySignature(),
      },
      exit: {
        personnel: emptySignature(),
        representant: emptySignature(),
      },
    },
    effetsConfies: [
      {
        id: "E-1",
        typeEffet: "BADGE INTRUSION",
        designation: "",
        dateRemise: "2026-01-01",
        dateRetour: "",
        statutManuel: "ACTIF",
        cause: "",
        etatFacturation: "",
      },
    ],
    ...overrides,
  };
}

test("workflow arrivée: effet confié sans double signature déclenche une alerte", () => {
  const person = basePerson();

  const alerts = buildUiOverviewAlerts([person], [], "2026-10-07");

  assert.equal(alerts.some((alert) => alert.type === "signaturePdf"), true);
  assert.equal(alerts.some((alert) => /DOCUMENT D'ARRIVEE NON SIGNE/.test(alert.text)), true);
});

test("workflow arrivée: double signature sans PDF signé déclenche une alerte PDF", () => {
  const person = basePerson({
    signatures: {
      arrival: {
        personnel: signedSignature("2026-10-07T10:00:00.000Z"),
        representant: signedSignature("2026-10-07T10:01:00.000Z"),
      },
      exit: {
        personnel: emptySignature(),
        representant: emptySignature(),
      },
    },
  });

  const alerts = buildUiOverviewAlerts([person], [], "2026-10-07");

  assert.equal(alerts.some((alert) => alert.type === "signaturePdf"), true);
  assert.equal(alerts.some((alert) => /PDF ABSENT OU NON MIS A JOUR/.test(alert.text)), true);
});

test("workflow arrivée: double signature + archive signée courante ne déclenche plus d'alerte PDF", () => {
  const fingerprint = "fingerprint-arrival-v1";
  const person = basePerson({
    documentFingerprints: { arrival: fingerprint },
    signatures: {
      arrival: {
        personnel: signedSignature("2026-10-07T10:00:00.000Z"),
        representant: signedSignature("2026-10-07T10:01:00.000Z"),
      },
      exit: {
        personnel: emptySignature(),
        representant: emptySignature(),
      },
    },
  });
  const archives = [
    {
      personId: person.id,
      typeDocument: "ARRIVEE",
      statutSignature: "SIGNE",
      pdfPath: "data/pdf/arrivee/P-TEST/document-arrivee-P-TEST.pdf",
      fingerprint,
      dateArchivage: "2026-10-07T10:02:00.000Z",
    },
  ];

  const alerts = buildUiOverviewAlerts([person], archives, "2026-10-07");

  assert.equal(alerts.some((alert) => /ARRIVEE SIGNEE/.test(alert.text)), false);
});

test("workflow sortie: personne sortie + double signature sans PDF signé déclenche une alerte PDF sortie", () => {
  const person = basePerson({
    dateSortieReelle: "2026-10-07",
    effetsConfies: [
      {
        id: "E-1",
        typeEffet: "CLE",
        designation: "DE",
        dateRemise: "2026-01-01",
        dateRetour: "2026-10-07",
        statutManuel: "ACTIF",
        cause: "",
        etatFacturation: "",
      },
    ],
    signatures: {
      arrival: {
        personnel: signedSignature(),
        representant: signedSignature(),
      },
      exit: {
        personnel: signedSignature("2026-10-07T11:00:00.000Z"),
        representant: signedSignature("2026-10-07T11:01:00.000Z"),
      },
    },
  });

  const alerts = buildUiOverviewAlerts([person], [], "2026-10-07");

  assert.equal(getDossierStatus(person), "SORTI");
  assert.equal(alerts.some((alert) => alert.type === "signaturePdf" && /SORTIE SIGNEE/.test(alert.text)), true);
});

test("workflow sortie: effet rendu avec date de retour n'est pas facturable", () => {
  const person = basePerson({ dateSortieReelle: "2026-10-07" });
  const effect = {
    typeEffet: "CLE",
    designation: "DE",
    dateRemise: "2026-01-01",
    dateRetour: "2026-10-07",
    statutManuel: "ACTIF",
    cause: "",
    etatFacturation: "",
  };

  assert.equal(getEffectStatus(person, effect), "RESTITUE");
  assert.equal(getEffectBillingCause(person, effect), "");
  assert.equal(getEffectChargeableAmount(person, effect, pricingRules), 0);
  assert.equal(getEffectBillingStatus(effect, false), "-");
});

test("workflow sortie: effet non rendu après sortie reste facturable", () => {
  const person = basePerson({ dateSortieReelle: "2026-10-07" });
  const effect = {
    typeEffet: "BADGE INTRUSION",
    designation: "",
    dateRemise: "2026-01-01",
    dateRetour: "",
    statutManuel: "ACTIF",
    cause: "",
    etatFacturation: "",
  };

  assert.equal(getEffectStatus(person, effect), "NON RENDU");
  assert.equal(getEffectBillingCause(person, effect), "NON RENDU");
  assert.equal(getEffectChargeableAmount(person, effect, pricingRules), 15);
  assert.equal(getEffectBillingStatus(effect, true), "A FACTURER");
  assert.equal(getTotalChargeableAmount(person, [effect], pricingRules), 15);
});

test("workflow UI: les pages documents conservent les commandes critiques", () => {
  const arrival = fs.readFileSync("document-arrivee.html", "utf8");
  const exit = fs.readFileSync("document-sortie.html", "utf8");
  const archives = fs.readFileSync("documents-archives.html", "utf8");

  assert.match(arrival, /data-doc-type="arrival"/);
  assert.match(exit, /data-doc-type="exit"/);
  assert.match(arrival, /SIGNER SUR TELEPHONE/);
  assert.match(exit, /SIGNER SUR TELEPHONE/);
  assert.match(archives, /STATUT SIGNATURE/);
  assert.match(archives, /TOTAL FACTURABLE/);
});
