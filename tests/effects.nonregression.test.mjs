import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import vm from "node:vm";

import {
  getEffectBillingCause,
  getReplacementCostValue,
  normalizeManualStatus,
} from "../smartphone/src/lib/businessRules.js";

function extractFunctionSource(source, name) {
  const start = source.indexOf(`function ${name}(`);
  if (start < 0) throw new Error(`Function not found: ${name}`);
  const sig = source.indexOf("(", start);
  if (sig < 0) throw new Error(`Function signature not found: ${name}`);
  let parenDepth = 0;
  let bodyStart = -1;
  for (let j = sig; j < source.length; j += 1) {
    const ch = source[j];
    if (ch === "(") parenDepth += 1;
    if (ch === ")") {
      parenDepth -= 1;
      if (parenDepth === 0) {
        bodyStart = source.indexOf("{", j);
        break;
      }
    }
  }
  if (bodyStart < 0) throw new Error(`Function body not found: ${name}`);
  let i = bodyStart;
  let depth = 0;
  for (; i < source.length; i += 1) {
    const ch = source[i];
    if (ch === "{") depth += 1;
    if (ch === "}") {
      depth -= 1;
      if (depth === 0) {
        return source.slice(start, i + 1);
      }
    }
  }
  throw new Error(`Function extraction failed: ${name}`);
}

function loadPcBillingFns() {
  const source = fs.readFileSync("app.js", "utf8");
  const fnNames = [
    "normalizeText",
    "normalizeAmount",
    "normalizeEffectCause",
    "isCesKeyDesignation",
    "getReplacementCostValue",
    "getTodayIsoDate",
    "isPastDate",
    "isExitDue",
    "getEffectReplacementCause",
    "getEffectReplacementCost",
  ];

  const context = {
    state: { data: { listes: { coutsRemplacement: [] } } },
    BILLABLE_EFFECT_CAUSES: ["PERTE", "VOL", "NON RENDU", "DETRUIT"],
    Date,
  };
  vm.createContext(context);
  for (const name of fnNames) {
    vm.runInContext(extractFunctionSource(source, name), context);
  }
  return {
    getEffectReplacementCause: context.getEffectReplacementCause,
    getEffectReplacementCost: context.getEffectReplacementCost,
    state: context.state,
  };
}

function loadLegacyDbFns() {
  const source = fs.readFileSync("smartphone/src/lib/db.js", "utf8");
  const fnNames = [
    "toString",
    "cleanDate",
    "normalizeCause",
    "inferCauseFromStatus",
    "applyLegacyEffetToRaw",
    "normalizeLegacyEffet",
  ];

  const context = {
    normalizeManualStatus,
  };
  vm.createContext(context);
  for (const name of fnNames) {
    vm.runInContext(extractFunctionSource(source, name), context);
  }
  return {
    applyLegacyEffetToRaw: context.applyLegacyEffetToRaw,
    normalizeLegacyEffet: context.normalizeLegacyEffet,
  };
}

const pricingRules = [
  { typeEffet: "BADGE INTRUSION", cause: "DETRUIT", montant: 15 },
  { typeEffet: "BADGE INTRUSION", cause: "VOL", montant: 20 },
  { typeEffet: "BADGE INTRUSION", cause: "NON RENDU", montant: 30 },
  { typeEffet: "BADGE INTRUSION", cause: "HS", montant: 99 },
];

function toSmartphoneCost(person, effect) {
  const cause = getEffectBillingCause(person, effect);
  return getReplacementCostValue(pricingRules, effect.typeEffet, cause, effect.designation || "");
}

function primePcCosts(pc) {
  pc.state.data.listes.coutsRemplacement = pricingRules.map((row) => ({ ...row }));
}

test("DETRUIT sans retour => cause DETRUIT et tarif DETRUIT (PC/smartphone)", () => {
  const pc = loadPcBillingFns();
  primePcCosts(pc);

  const person = { dateSortieReelle: "2000-01-01" };
  const effect = { typeEffet: "BADGE INTRUSION", cause: "DETRUIT", dateRetour: "" };

  assert.equal(pc.getEffectReplacementCause(person, effect), "DETRUIT");
  assert.equal(pc.getEffectReplacementCost(person, effect), 15);
  assert.equal(getEffectBillingCause(person, effect), "DETRUIT");
  assert.equal(toSmartphoneCost(person, effect), 15);
});

test("DETRUIT avec retour => cause conservée et tarif conservé (PC/smartphone)", () => {
  const pc = loadPcBillingFns();
  primePcCosts(pc);

  const person = { dateSortieReelle: "2000-01-01" };
  const effect = { typeEffet: "BADGE INTRUSION", cause: "DETRUIT", dateRetour: "2026-04-05" };

  assert.equal(pc.getEffectReplacementCause(person, effect), "DETRUIT");
  assert.equal(pc.getEffectReplacementCost(person, effect), 15);
  assert.equal(getEffectBillingCause(person, effect), "DETRUIT");
  assert.equal(toSmartphoneCost(person, effect), 15);
});

test("HS avec retour => coût 0 partout", () => {
  const pc = loadPcBillingFns();
  primePcCosts(pc);

  const person = { dateSortieReelle: "2000-01-01" };
  const effect = { typeEffet: "BADGE INTRUSION", cause: "HS", dateRetour: "2026-04-05" };

  assert.equal(pc.getEffectReplacementCause(person, effect), "HS");
  assert.equal(pc.getEffectReplacementCost(person, effect), 0);
  assert.equal(getEffectBillingCause(person, effect), "HS");
  assert.equal(toSmartphoneCost(person, effect), 0);
});

test("VOL avec retour => cause VOL et tarif VOL (PC/smartphone)", () => {
  const pc = loadPcBillingFns();
  primePcCosts(pc);

  const person = { dateSortieReelle: "2000-01-01" };
  const effect = { typeEffet: "BADGE INTRUSION", cause: "VOL", dateRetour: "2026-04-05" };

  assert.equal(pc.getEffectReplacementCause(person, effect), "VOL");
  assert.equal(pc.getEffectReplacementCost(person, effect), 20);
  assert.equal(getEffectBillingCause(person, effect), "VOL");
  assert.equal(toSmartphoneCost(person, effect), 20);
});

test("Aucune cause + sortie due + sans retour => fallback NON RENDU facturable (PC/smartphone)", () => {
  const pc = loadPcBillingFns();
  primePcCosts(pc);

  const person = { dateSortieReelle: "2000-01-01" };
  const effect = { typeEffet: "BADGE INTRUSION", cause: "", dateRetour: "" };

  assert.equal(pc.getEffectReplacementCause(person, effect), "NON RENDU");
  assert.equal(pc.getEffectReplacementCost(person, effect), 30);
  assert.equal(getEffectBillingCause(person, effect), "NON RENDU");
  assert.equal(toSmartphoneCost(person, effect), 30);
});

test("Cause existante non écrasée au flux save/reload legacy mobile", () => {
  const { applyLegacyEffetToRaw, normalizeLegacyEffet } = loadLegacyDbFns();

  const raw = {
    id: "E000001",
    cause: "DETRUIT",
    statutManuel: "ACTIF",
    dateRetour: "",
    typeEffet: "BADGE INTRUSION",
  };

  const updated = applyLegacyEffetToRaw(raw, { statut: "ACTIF", dateRetour: "2026-04-05" });
  assert.equal(updated.cause, "DETRUIT");

  const reloaded = normalizeLegacyEffet(updated, "P0001");
  assert.equal(reloaded.cause, "DETRUIT");
});

test("Aucune normalisation destructive explicite restante dans app.js", () => {
  const source = fs.readFileSync("app.js", "utf8");
  assert.equal(source.includes("delete effect.causeRemplacement"), false);
  assert.equal(source.includes('normalizedStatus === "RESTITUE"'), false);
});

test("Coût aligné PC/smartphone pour mêmes typeEffet + cause", () => {
  const pc = loadPcBillingFns();
  primePcCosts(pc);
  const person = { dateSortieReelle: "2000-01-01" };

  const cases = [
    { cause: "DETRUIT", dateRetour: "", expected: 15 },
    { cause: "DETRUIT", dateRetour: "2026-04-05", expected: 15 },
    { cause: "VOL", dateRetour: "2026-04-05", expected: 20 },
    { cause: "HS", dateRetour: "2026-04-05", expected: 0 },
    { cause: "", dateRetour: "", expected: 30 },
  ];

  for (const sample of cases) {
    const effect = {
      typeEffet: "BADGE INTRUSION",
      cause: sample.cause,
      dateRetour: sample.dateRetour,
    };
    const pcCost = pc.getEffectReplacementCost(person, effect);
    const smCost = toSmartphoneCost(person, effect);
    assert.equal(pcCost, sample.expected);
    assert.equal(smCost, sample.expected);
  }
});

