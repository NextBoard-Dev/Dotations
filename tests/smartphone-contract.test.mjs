import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";

const ROOT = process.cwd();

function read(file) {
  return fs.readFileSync(path.join(ROOT, file), "utf8");
}

function count(source, pattern) {
  return (source.match(pattern) || []).length;
}

const LEGACY_VERSION = "v3";
const OLD_MOBILE_BRANDING_PATTERN = new RegExp(
  [
    ["nextboard-", "(?:favicon|app|brand|logo|mobile-header)", "-", LEGACY_VERSION].join(""),
    ["nextboard-mobile-header", "\\.png"].join(""),
  ].join("|"),
  "i"
);

test("smartphone package: la chaine de build reste autonome", () => {
  const pkg = JSON.parse(read("smartphone/package.json"));

  assert.equal(pkg.type, "module");
  assert.equal(pkg.scripts.build, "node ./node_modules/vite/bin/vite.js build");
  assert.equal(pkg.scripts.lint, "node ./node_modules/eslint/bin/eslint.js . --quiet");
  assert.equal(pkg.scripts.typecheck, "tsc -p ./jsconfig.json");
  assert.ok(pkg.dependencies["@supabase/supabase-js"], "Supabase doit rester une dependance directe du smartphone");
  assert.equal(pkg.dependencies.react, "18.2.0");
  assert.equal(pkg.dependencies["react-dom"], "18.2.0");
});

test("smartphone UI: navigation, marque et donnees critiques restent branchees", () => {
  const source = read("smartphone/src/pages/Mobile.jsx");

  assert.match(source, /MOBILE_BRAND_LOGO_URL = import\.meta\.env\.BASE_URL \+ "branding\/nextboard-mobile-header-v4\.png\?v=20260928b"/);
  assert.doesNotMatch(source, OLD_MOBILE_BRANDING_PATTERN);

  for (const component of [
    "MobileOverview",
    "MobileFichePerson",
    "MobileDocumentArrivee",
    "MobileDocumentSortie",
  ]) {
    assert.match(source, new RegExp(`import ${component} from`), `${component} doit rester importe`);
  }

  for (const [id, label] of [
    ["overview", "VUE D'ENSEMBLE"],
    ["fiche", "FICHE"],
    ["arrivee", "ENTREE"],
    ["sortie", "SORTIE"],
  ]) {
    assert.match(source, new RegExp(`\\{ id: "${id}", label: "${label}"`), `onglet ${label} manquant`);
  }

  assert.match(source, /await Promise\.allSettled\(\[/, "le chargement mobile doit tolerer les lectures partielles");
  assert.match(source, /db\.Person\.list\("-created_at", 5000\)/);
  assert.match(source, /db\.Effet\.list\("-created_at", 5000\)/);
  assert.match(source, /db\.AppState\.getReferenceBases\(\)/);
  assert.match(source, /db\.AppState\.getOperationalData\(\)/);
  assert.match(source, /buildUiOverviewAlerts\(/, "les alertes UI doivent venir des regles metier partagees");
});

test("smartphone donnees: droits et mode lecture seule protegent les ecritures", () => {
  const source = read("smartphone/src/lib/db.js");

  assert.match(source, /const WRITE_ROLES = new Set\(\["editor", "admin"\]\);/);
  assert.match(source, /const DELETE_ROLES = new Set\(\["admin"\]\);/);
  assert.match(source, /const READ_ONLY_MODE = resolveReadonlyMode\(\);/);
  assert.match(source, /VITE_SMARTPHONE_READONLY/);
  assert.match(source, /new URLSearchParams\(window\.location\.search\)\.get\("readonly"\)/);
  assert.match(source, /function ensureWritable\(actionLabel = "operation"\)/);
  assert.ok(count(source, /ensureWritable\('/g) >= 9, "creation, modification et suppression doivent bloquer en lecture seule");

  for (const label of [
    "creation personne",
    "mise a jour personne",
    "suppression personne",
    "creation effet",
    "mise a jour effet",
    "suppression effet",
    "creation signature",
    "mise a jour signature",
    "suppression signature",
  ]) {
    assert.match(source, new RegExp(`ensureWritable\\('${label}'\\)`), `${label} doit etre protegee`);
  }
});

test("smartphone signatures: validation, fusion et retour document restent contractuels", () => {
  const source = read("smartphone/src/lib/db.js");

  assert.match(source, /const DOC_TYPES = new Set\(\["arrival", "exit"\]\);/);
  assert.match(source, /const SIGNERS = new Set\(\["personnel", "representant"\]\);/);
  assert.match(source, /const MAX_SIGNATURE_DATA_LENGTH = 2_000_000;/);
  assert.match(source, /function normalizeSignatureData\(value\)/);
  assert.match(source, /data:image\\\/\(png\|jpeg\|jpg\);base64/);
  assert.match(source, /Format de signature invalide/);
  assert.match(source, /Signature trop volumineuse/);

  assert.match(source, /function mergeSignatureRecords\(\.\.\.groups\)/);
  assert.match(source, /extractLegacySignatures\(legacy\.payload, filters\)/);
  assert.match(source, /fetchSqlSignatureRecords\(filters\)/);
  assert.match(source, /mergeSignatureRecords\(legacySigs, sqlSigs\)/);
  assert.match(source, /saveLegacySignatureWithConflictRetry/);

  assert.match(source, /async function applySqlPersonCompletionFromSignatures\(personId, docType\)/);
  assert.match(source, /if \(!isSigned\("personnel"\) \|\| !isSigned\("representant"\)\) return;/);
  assert.match(source, /await sqlPerson\.update\(personId, \{ dateEntree: getTodayIsoDate\(\) \}\)/);
  assert.match(source, /await sqlPerson\.update\(personId, \{ dateSortieReelle: getTodayIsoDate\(\) \}\)/);
});

test("smartphone signatures: les rafraichissements nettoient timers et ecouteurs", () => {
  for (const [file, docType] of [
    ["smartphone/src/components/mobile/MobileDocumentArrivee.jsx", "arrival"],
    ["smartphone/src/components/mobile/MobileDocumentSortie.jsx", "exit"],
  ]) {
    const source = read(file);

    assert.match(source, new RegExp(`db\\.Signature\\.filter\\(\\{ personId: selectedPerson\\.id, docType: "${docType}" \\}\\)`));
    assert.match(source, /const timer = window\.setInterval\(refreshSignatures, 3000\);/);
    assert.match(source, /window\.addEventListener\("focus", refreshSignatures\);/);
    assert.match(source, /document\.addEventListener\("visibilitychange", refreshSignatures\);/);
    assert.match(source, /return \(\) => \{/);
    assert.match(source, /stopped = true;/);
    assert.match(source, /window\.clearInterval\(timer\);/);
    assert.match(source, /window\.removeEventListener\("focus", refreshSignatures\);/);
    assert.match(source, /document\.removeEventListener\("visibilitychange", refreshSignatures\);/);
  }
});

test("smartphone suppressions: les donnees metier sont neutralisees sans suppression physique directe", () => {
  const source = read("smartphone/src/lib/db.js");

  assert.match(source, /function markSoftDeleted\(entry = \{\}, deletedBy = ""\)/);
  assert.match(source, /is_deleted: true/);
  assert.match(source, /deleted_at: nowIso/);
  assert.match(source, /deletedAt: nowIso/);
  assert.match(source, /\.update\(\{ is_deleted: true, updated_at: new Date\(\)\.toISOString\(\) \}\)/);
  assert.doesNotMatch(source, /supabase\.from\([^)]+\)\.delete\(\)/, "les suppressions Supabase doivent rester logiques");
});
