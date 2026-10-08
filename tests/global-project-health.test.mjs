import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";

const ROOT = process.cwd();
const CORE_HTML_FILES = [
  "index.html",
  "fiche-personne.html",
  "document-arrivee.html",
  "document-sortie.html",
  "documents-archives.html",
  "bases-reference.html",
  "signature-mobile.html",
  "suivi-global.html",
];
const REDIRECT_HTML_FILES = new Set(["suivi-global.html"]);
const CSS_FILES = [
  "style.css",
  "mobile/assets/index-0baX2u83.css",
  "smartphone/src/index.css",
];

function read(file) {
  return fs.readFileSync(path.join(ROOT, file), "utf8");
}

function listTrackedFiles() {
  return execFileSync("git", ["ls-files"], { cwd: ROOT, encoding: "utf8" })
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
}

function stripQuery(value) {
  return String(value || "").split("#")[0].split("?")[0];
}

function isLocalReference(value) {
  const raw = String(value || "").trim();
  return Boolean(
    raw &&
      !raw.startsWith("#") &&
      !raw.startsWith("mailto:") &&
      !raw.startsWith("tel:") &&
      !raw.startsWith("http://") &&
      !raw.startsWith("https://") &&
      !raw.startsWith("data:") &&
      !raw.startsWith("javascript:")
  );
}

function getHtmlLocalReferences(html) {
  const refs = [];
  const attrPattern = /\b(?:href|src)=["']([^"']+)["']/g;
  let match;
  while ((match = attrPattern.exec(html))) {
    const ref = stripQuery(match[1]);
    if (isLocalReference(ref)) refs.push(ref);
  }
  return refs;
}

function getCssLocalReferences(css) {
  const refs = [];
  const urlPattern = /url\(\s*(?:"([^"]+)"|'([^']+)'|([^)'"\s]+))\s*\)/g;
  let match;
  while ((match = urlPattern.exec(css))) {
    const ref = stripQuery(match[1] || match[2] || match[3] || "");
    if (isLocalReference(ref)) refs.push(ref);
  }
  return refs;
}

function isIsoDateOrEmpty(value) {
  const raw = String(value || "").trim();
  return !raw || /^\d{4}-\d{2}-\d{2}$/.test(raw);
}

test("socle projet: les pages principales existent et chargent style/app", () => {
  for (const file of CORE_HTML_FILES) {
    const fullPath = path.join(ROOT, file);
    assert.equal(fs.existsSync(fullPath), true, `${file} doit exister`);
    const html = read(file);
    assert.match(html, /<meta charset="UTF-8"/i, `${file} doit declarer UTF-8`);
    if (REDIRECT_HTML_FILES.has(file)) {
      assert.match(html, /http-equiv="refresh"|url=index\.html|href="index\.html"/i, `${file} doit etre une redirection lisible`);
      continue;
    }
    assert.match(html, /style\.css/i, `${file} doit charger style.css`);
    if (file !== "signature-mobile.html") {
      assert.match(html, /app\.js/i, `${file} doit charger app.js`);
    }
  }
});

test("socle projet: les liens locaux des pages principales pointent vers des fichiers existants", () => {
  const missing = [];
  for (const file of CORE_HTML_FILES) {
    const baseDir = path.dirname(path.join(ROOT, file));
    const refs = getHtmlLocalReferences(read(file));
    for (const ref of refs) {
      const normalized = ref.replace(/\//g, path.sep);
      const target = path.resolve(baseDir, normalized);
      if (!fs.existsSync(target)) {
        missing.push(`${file} -> ${ref}`);
      }
    }
  }
  assert.deepEqual(missing, []);
});

test("socle projet: les ressources locales referencees par les CSS existent", () => {
  const missing = [];
  for (const file of CSS_FILES) {
    const baseDir = path.dirname(path.join(ROOT, file));
    const refs = getCssLocalReferences(read(file));
    for (const ref of refs) {
      const normalized = ref.replace(/\//g, path.sep);
      const target = path.resolve(baseDir, normalized);
      if (!fs.existsSync(target)) {
        missing.push(`${file} -> ${ref}`);
      }
    }
  }

  assert.deepEqual(missing, []);
});

test("socle projet: les CSS ne referencent pas les anciens favicons ou anciens logos", () => {
  const forbidden = [];
  for (const file of CSS_FILES) {
    const refs = getCssLocalReferences(read(file));
    for (const ref of refs) {
      if (/ancien|old|favicon-old|nextboard-old|ancienne-icone|old-logo/i.test(ref)) {
        forbidden.push(`${file} -> ${ref}`);
      }
    }
  }

  assert.deepEqual(forbidden, []);
});

test("socle donnees: data.json est lisible et respecte les champs critiques", () => {
  const data = JSON.parse(read("data.json"));
  assert.equal(Array.isArray(data.personnes), true, "personnes doit etre une liste");
  assert.equal(Array.isArray(data.listes?.coutsRemplacement), true, "listes.coutsRemplacement doit etre une liste");

  const personIds = new Set();
  const duplicatePersonIds = [];
  const invalidDates = [];
  const missingSignatureBlocks = [];

  for (const person of data.personnes) {
    const id = String(person?.id || "").trim();
    if (!id) duplicatePersonIds.push("(id vide)");
    if (personIds.has(id)) duplicatePersonIds.push(id);
    personIds.add(id);

    assert.ok(String(person?.nom || "").trim(), `${id} doit avoir un nom`);
    assert.ok(String(person?.prenom || "").trim(), `${id} doit avoir un prenom`);

    for (const field of ["dateEntree", "dateSortiePrevue", "dateSortieReelle"]) {
      if (!isIsoDateOrEmpty(person?.[field])) invalidDates.push(`${id}.${field}=${person?.[field]}`);
    }

    for (const docType of ["arrival", "exit"]) {
      for (const signer of ["personnel", "representant"]) {
        const signature = person?.signatures?.[docType]?.[signer];
        if (!signature || typeof signature !== "object") {
          missingSignatureBlocks.push(`${id}.${docType}.${signer}`);
        }
      }
    }

    const effectIds = new Set();
    const effects = Array.isArray(person?.effetsConfies) ? person.effetsConfies : [];
    for (const effect of effects) {
      const effectId = String(effect?.id || "").trim();
      assert.ok(effectId, `${id} contient un effet sans id`);
      assert.equal(effectIds.has(effectId), false, `${id} contient un effet en double: ${effectId}`);
      effectIds.add(effectId);
      for (const field of ["dateRemise", "dateRetour", "dateRemplacement"]) {
        if (!isIsoDateOrEmpty(effect?.[field])) invalidDates.push(`${id}.${effectId}.${field}=${effect?.[field]}`);
      }
    }
  }

  assert.deepEqual(duplicatePersonIds, []);
  assert.deepEqual(invalidDates, []);
  assert.deepEqual(missingSignatureBlocks, []);
});

test("socle donnees: les tarifs critiques existent", () => {
  const data = JSON.parse(read("data.json"));
  const rows = data.listes?.coutsRemplacement || [];
  const key = (row) => `${String(row?.typeEffet || "").toUpperCase()}|${String(row?.cause || "").toUpperCase()}`;
  const available = new Set(rows.map(key));
  const required = [
    "BADGE INTRUSION|NON RENDU",
    "CARTE TURBOSELF|NON RENDU",
    "CLE|NON RENDU",
    "CLE CES|NON RENDU",
    "TELECOMMANDE URMET|NON RENDU",
  ];
  assert.deepEqual(required.filter((entry) => !available.has(entry)), []);
});

test("socle fichiers: aucun artefact local connu ne doit polluer le depot", () => {
  const forbiddenPatterns = [
    /\.bak/i,
    /~$/,
    /\.tmp$/i,
    /__pycache__/i,
    /node_modules/i,
    /_quarantine/i,
    /desktop\.ini$/i,
  ];
  const allowed = new Set(["desktop.ini"]);
  const offenders = listTrackedFiles()
    .filter((file) => !allowed.has(file))
    .filter((file) => forbiddenPatterns.some((pattern) => pattern.test(file)));

  assert.deepEqual(offenders, []);
});
