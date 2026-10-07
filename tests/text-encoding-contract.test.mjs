import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";

const ROOT = process.cwd();
const TEXT_EXTENSIONS = new Set([".bat", ".css", ".html", ".js", ".json", ".md", ".mjs", ".sql", ".url"]);
const MOJIBAKE_PATTERN = /�|Ã.|Â.|â€™|â€œ|â€|ACC�/;
const HTML_ENTRYPOINTS = [
  "index.html",
  "fiche-personne.html",
  "document-arrivee.html",
  "document-sortie.html",
  "documents-archives.html",
  "bases-reference.html",
  "signature-mobile.html",
  "suivi-global.html",
  "mobile/index.html",
  "smartphone/index.html",
];

function read(file) {
  return fs.readFileSync(path.join(ROOT, file), "utf8");
}

function trackedTextFiles() {
  return execFileSync("git", ["ls-files"], { cwd: ROOT, encoding: "utf8" })
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .filter((file) => TEXT_EXTENSIONS.has(path.extname(file).toLowerCase()));
}

test("texte: aucun fichier texte versionne ne contient d'encodage casse visible", () => {
  const offenders = trackedTextFiles().filter((file) => MOJIBAKE_PATTERN.test(read(file)));
  assert.deepEqual(offenders, []);
});

test("texte: toutes les pages publiees declarent la langue francaise", () => {
  const offenders = HTML_ENTRYPOINTS.filter((file) => !/<html lang="fr">/.test(read(file)));
  assert.deepEqual(offenders, []);
});

test("texte: les libelles critiques restent lisibles dans les pages principales", () => {
  const expectedLabels = [
    ["index.html", "SUIVI DES DOTATIONS ENTREE / SORTIE"],
    ["fiche-personne.html", "FICHE PERSONNE"],
    ["document-arrivee.html", "DOCUMENT D'ARRIVEE"],
    ["document-sortie.html", "DOCUMENT DE SORTIE"],
    ["documents-archives.html", "DOCUMENTS / ARCHIVES"],
    ["bases-reference.html", "BASES DE REFERENCE"],
    ["signature-mobile.html", "SIGNATURE MOBILE"],
    ["mobile/index.html", "SUIVI DES DOTATIONS ENTREE / SORTIE"],
    ["smartphone/index.html", "SUIVI DES DOTATIONS ENTREE / SORTIE"],
  ];

  for (const [file, label] of expectedLabels) {
    assert.match(read(file), new RegExp(label.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")), `${file} doit garder le libelle ${label}`);
  }
});
