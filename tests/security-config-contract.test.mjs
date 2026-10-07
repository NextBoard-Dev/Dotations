import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";

const ROOT = process.cwd();
const TEXT_EXTENSIONS = new Set([".bat", ".css", ".html", ".js", ".json", ".md", ".mjs", ".sql", ".url"]);
const ALLOWED_LOCAL_PATH_FILES = new Set([
  "Ouvrir-Dotations-PC-Local.url",
  "Ouvrir-Dotations-Telephone-Local.bat",
  "REGLES_PROJET.md",
  "scripts/planifier_backup_quotidien.bat",
  "tests/launchers-contract.test.mjs",
  "tests/security-config-contract.test.mjs",
]);
const ALLOWED_LOCAL_RUNTIME_FILES = new Set([
  "app.js",
  "index.html",
  "fiche-personne.html",
  "document-arrivee.html",
  "document-sortie.html",
  "documents-archives.html",
  "bases-reference.html",
  "signature-mobile.html",
  "suivi-global.html",
]);
const PRIVATE_SECRET_PATTERN = /service[_-]?role|SUPABASE_SERVICE|sk-[A-Za-z0-9_-]{20,}|ghp_[A-Za-z0-9_]{20,}|github_pat_[A-Za-z0-9_]{20,}/i;
const OLD_DASHBOARD_PATH_PATTERN = /GESTION(?:%20| )DES(?:%20| )ACC|EFFETS(?:%20| )SENSIBLES|DOTATIONS - MODE LOCAL - VERSION ACTIVE/i;
const ABSOLUTE_WINDOWS_PATH_PATTERN = /[A-Z]:\\Users\\sebastien\.duc\\/i;
const LOCAL_RUNTIME_PATTERN = /127\.0\.0\.1|localhost|192\.168\.|IP_DU_PC/i;

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

test("securite: aucun secret prive evident ne doit etre versionne", () => {
  const offenders = trackedTextFiles().filter((file) => PRIVATE_SECRET_PATTERN.test(read(file)));
  assert.deepEqual(offenders, []);
});

test("configuration: aucun ancien chemin dashboard ne doit rester dans les fichiers versionnes", () => {
  const offenders = trackedTextFiles().filter((file) => OLD_DASHBOARD_PATH_PATTERN.test(read(file)));
  assert.deepEqual(offenders, []);
});

test("configuration: les chemins absolus Windows restent limites aux lanceurs et docs explicites", () => {
  const offenders = trackedTextFiles()
    .filter((file) => !ALLOWED_LOCAL_PATH_FILES.has(file))
    .filter((file) => ABSOLUTE_WINDOWS_PATH_PATTERN.test(read(file)));

  assert.deepEqual(offenders, []);
});

test("configuration: les references locales runtime restent limitees aux points d'entree connus", () => {
  const offenders = trackedTextFiles()
    .filter((file) => !/^mobile\/assets\//.test(file))
    .filter((file) => !ALLOWED_LOCAL_PATH_FILES.has(file))
    .filter((file) => !ALLOWED_LOCAL_RUNTIME_FILES.has(file))
    .filter((file) => LOCAL_RUNTIME_PATTERN.test(read(file)));

  assert.deepEqual(offenders, []);
});

test("configuration: le manifeste des icones sidebar reste portable", () => {
  const manifest = JSON.parse(read("assets/_icones_sources/MANIFEST_ICONES_SIDEBAR.json"));

  for (const entry of manifest) {
    assert.match(String(entry.path || ""), /^assets\/sidebar\//, `${entry.name} doit utiliser un chemin relatif`);
    assert.doesNotMatch(String(entry.path || ""), ABSOLUTE_WINDOWS_PATH_PATTERN);
    assert.equal(path.basename(String(entry.path || "")), entry.name);
  }
});
