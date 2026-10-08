import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";

const ROOT = process.cwd();
const TEXT_EXTENSIONS = new Set([".bat", ".css", ".html", ".js", ".json", ".jsx", ".md", ".mjs", ".sql", ".url"]);
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
const PRIVATE_SECRET_PATTERNS = [
  ["service", "[_-]?", "role"].join(""),
  ["SUPABASE", "_", "SERVICE"].join(""),
  ["sk", "-", "[A-Za-z0-9_-]{20,}"].join(""),
  ["ghp", "_", "[A-Za-z0-9_]{20,}"].join(""),
  ["github", "_", "pat", "_", "[A-Za-z0-9_]{20,}"].join(""),
].map((source) => new RegExp(source, "i"));
const OLD_DASHBOARD_PATH_PATTERNS = [
  ["GESTION", "(?:%20| )", "DES", "(?:%20| )", "ACC"].join(""),
  ["EFFETS", "(?:%20| )", "SENSIBLES"].join(""),
  ["DOTATIONS", " - ", "MODE LOCAL", " - ", "VERSION ACTIVE"].join(""),
].map((source) => new RegExp(source, "i"));
const ABSOLUTE_WINDOWS_PATH_PATTERN = /[A-Z]:\\Users\\sebastien\.duc\\/i;
const LOCAL_RUNTIME_PATTERN = /127\.0\.0\.1|localhost|192\.168\.|IP_DU_PC/i;
const STORAGE_SET_ITEM_PATTERN = /(?:window\.)?(?:localStorage|sessionStorage)\.setItem\(([^,\n]+),\s*([^)]+)\)/g;
const STORAGE_KEY_DECLARATION_PATTERN = /const\s+([A-Z0-9_]*(?:KEY|STORAGE)[A-Z0-9_]*)\s*=\s*"([^"]+)"/g;
const ALLOWED_STORAGE_KEYS = new Set([
  "__dotations_data_etag",
  "__dotations_data_snapshot",
  "dashboard-working-data",
]);

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
  const offenders = trackedTextFiles().filter((file) => {
    const source = read(file);
    return PRIVATE_SECRET_PATTERNS.some((pattern) => pattern.test(source));
  });
  assert.deepEqual(offenders, []);
});

test("configuration: aucun ancien chemin dashboard ne doit rester dans les fichiers versionnes", () => {
  const offenders = trackedTextFiles().filter((file) => {
    const source = read(file);
    return OLD_DASHBOARD_PATH_PATTERNS.some((pattern) => pattern.test(source));
  });
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

test("securite navigateur: aucun mot de passe n'est persiste en stockage local ou session", () => {
  const offenders = [];
  for (const file of trackedTextFiles()) {
    const source = read(file);
    let match;
    while ((match = STORAGE_SET_ITEM_PATTERN.exec(source))) {
      const storedValue = String(match[2] || "");
      if (/(?:password|motdepasse|pass|loginPassword)/i.test(storedValue)) {
        offenders.push(`${file}: ${match[0]}`);
      }
    }
  }

  assert.deepEqual(offenders, []);
});

test("securite navigateur: les cles de stockage restent bornees au namespace Dotations", () => {
  const offenders = [];
  for (const file of trackedTextFiles()) {
    const source = read(file);
    let match;
    while ((match = STORAGE_KEY_DECLARATION_PATTERN.exec(source))) {
      const keyName = String(match[1] || "");
      const keyValue = String(match[2] || "");
      if (/PUBLISHABLE_KEY|SERVICE_KEY|API_KEY/.test(keyName)) {
        continue;
      }
      const isDotationsKey = /^dotations[-_]/.test(keyValue) || /^__dotations_/.test(keyValue);
      if (!isDotationsKey && !ALLOWED_STORAGE_KEYS.has(keyValue)) {
        offenders.push(`${file}: ${keyName}=${keyValue}`);
      }
    }
  }

  assert.deepEqual(offenders, []);
});
