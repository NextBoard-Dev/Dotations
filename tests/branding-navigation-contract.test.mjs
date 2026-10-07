import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";

const ROOT = process.cwd();
const PC_HTML_FILES = [
  "index.html",
  "fiche-personne.html",
  "document-arrivee.html",
  "document-sortie.html",
  "documents-archives.html",
  "bases-reference.html",
  "signature-mobile.html",
];
const EXPECTED_PC_FAVICON = "favicon.ico?v=20260928b";
const EXPECTED_MOBILE_ICON_VERSION = "v4";
const EXPECTED_MOBILE_CACHE_VERSION = "20260928b";

function read(file) {
  return fs.readFileSync(path.join(ROOT, file), "utf8");
}

function trackedTextFiles() {
  const textExtensions = new Set([".html", ".js", ".jsx", ".css", ".json", ".md", ".bat", ".sql", ".mjs"]);
  return execFileSync("git", ["ls-files"], { cwd: ROOT, encoding: "utf8" })
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .filter((file) => textExtensions.has(path.extname(file).toLowerCase()));
}

function localRefsFromHtml(html) {
  const refs = [];
  const pattern = /\b(?:href|src)=["']([^"']+)["']/g;
  let match;
  while ((match = pattern.exec(html))) {
    const raw = match[1].replace("%BASE_URL%", "").split("#")[0].split("?")[0];
    if (
      raw &&
      !raw.startsWith("/") &&
      !raw.startsWith("http://") &&
      !raw.startsWith("https://") &&
      !raw.startsWith("data:") &&
      !raw.startsWith("mailto:")
    ) {
      refs.push(raw);
    }
  }
  return refs;
}

test("branding PC: toutes les pages gardent le favicon NextBoard courant", () => {
  for (const file of PC_HTML_FILES) {
    const html = read(file);
    assert.match(html, new RegExp(`href="${EXPECTED_PC_FAVICON.replace(/[.?]/g, "\\$&")}"`), `${file} doit pointer vers le favicon courant`);
    assert.doesNotMatch(html, /favicon-v3|nextboard-favicon-v3|nextboard-logo-v3|nextboard-brand-v3/i, `${file} ne doit pas referencer un ancien logo`);
  }
});

test("branding mobile: index et manifest utilisent uniquement les icones v4", () => {
  const mobileIndex = read("mobile/index.html");
  const smartphoneIndex = read("smartphone/index.html");
  const mobileManifest = JSON.parse(read("mobile/manifest.json"));
  const smartphoneManifest = JSON.parse(read("smartphone/public/manifest.json"));

  for (const source of [mobileIndex, smartphoneIndex, JSON.stringify(mobileManifest), JSON.stringify(smartphoneManifest)]) {
    assert.match(source, new RegExp(`nextboard-(?:favicon|app|mobile-header)-${EXPECTED_MOBILE_ICON_VERSION}`));
    assert.match(source, new RegExp(`v=${EXPECTED_MOBILE_CACHE_VERSION}`));
    assert.doesNotMatch(source, /nextboard-(?:favicon|app|brand|logo|mobile-header)-v3/i);
    assert.doesNotMatch(source, /nextboard-mobile-header\.png/i);
  }
});

test("branding mobile: les ressources référencees existent", () => {
  const checks = [
    ["mobile/index.html", "mobile"],
    ["smartphone/index.html", "smartphone/public"],
  ];

  const missing = [];
  for (const [file, baseDir] of checks) {
    for (const ref of localRefsFromHtml(read(file))) {
      if (ref.startsWith("assets/") && file.startsWith("smartphone/")) {
        continue;
      }
      const target = path.join(ROOT, baseDir, ref.replace(/^\.\//, ""));
      if (!fs.existsSync(target)) {
        missing.push(`${file} -> ${ref}`);
      }
    }
  }

  assert.deepEqual(missing, []);
});

test("navigation PC: les pages principales gardent le même menu", () => {
  const expectedNav = [
    ["index.html", "overview", "VUE D'ENSEMBLE"],
    ["fiche-personne.html", "person-sheet", "FICHE PERSONNE"],
    ["document-arrivee.html", "arrival-document", "ENTREE / SITUATION"],
    ["document-sortie.html", "exit-document", "SORTIE / SITUATION"],
    ["documents-archives.html", "documents-archives", "DOCUMENTS / ARCHIVES"],
    ["bases-reference.html", "reference-bases", "BASES DE REFERENCE"],
  ];

  for (const file of PC_HTML_FILES.filter((entry) => entry !== "signature-mobile.html")) {
    const html = read(file);
    for (const [href, navKey, label] of expectedNav) {
      const pattern = new RegExp(`<a href="${href}" class="sidebar__link" data-nav="${navKey}">${label}</a>`);
      assert.match(html, pattern, `${file} doit conserver le lien ${label}`);
    }
  }
});

test("branding: aucun fichier texte versionné ne référence les anciennes icônes mobiles", () => {
  const offenders = [];
  for (const file of trackedTextFiles()) {
    const source = read(file);
    if (/nextboard-(?:favicon|app|brand|logo|mobile-header)-v3|nextboard-mobile-header\.png/i.test(source)) {
      offenders.push(file);
    }
  }
  assert.deepEqual(offenders, []);
});
