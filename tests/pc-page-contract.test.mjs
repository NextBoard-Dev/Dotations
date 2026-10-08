import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";

const ROOT = process.cwd();
const DESKTOP_PAGES = [
  { file: "index.html", page: "overview", title: "Suivi des dotations Entree / Sortie" },
  { file: "fiche-personne.html", page: "person-sheet", title: "Fiche personne" },
  { file: "document-arrivee.html", page: "arrival-document", title: "Document d'arrivee" },
  { file: "document-sortie.html", page: "exit-document", title: "Document de sortie" },
  { file: "documents-archives.html", page: "documents-archives", title: "Documents / Archives" },
  { file: "bases-reference.html", page: "reference-bases", title: "Bases de reference" },
  { file: "signature-mobile.html", page: "mobile-signature", title: "Signature mobile" },
];

const SHELL_PAGES = DESKTOP_PAGES.filter((entry) => entry.page !== "mobile-signature");
const APP_SCRIPT_PATTERN = /<script src="app\.js\?v=20261007-mobile-signature-chain-guard"><\/script>/;

function read(file) {
  return fs.readFileSync(path.join(ROOT, file), "utf8");
}

function localRefsFromHtml(html) {
  const refs = [];
  const pattern = /\b(?:href|src)=["']([^"']+)["']/g;
  let match;
  while ((match = pattern.exec(html))) {
    const raw = match[1].trim();
    const clean = raw.split("#")[0].split("?")[0];
    if (
      clean &&
      !clean.startsWith("/") &&
      !clean.startsWith("http://") &&
      !clean.startsWith("https://") &&
      !clean.startsWith("data:") &&
      !clean.startsWith("mailto:")
    ) {
      refs.push(clean.replace(/^\.\//, ""));
    }
  }
  return refs;
}

function idsFromHtml(html) {
  return Array.from(html.matchAll(/\bid=["']([^"']+)["']/g), (match) => match[1].trim()).filter(Boolean);
}

function labelTargetsFromHtml(html) {
  return Array.from(html.matchAll(/<label\b[^>]*\bfor=["']([^"']+)["'][^>]*>/g), (match) => match[1].trim()).filter(Boolean);
}

test("pages PC: chaque page garde son identite et charge la meme application", () => {
  for (const { file, page, title } of DESKTOP_PAGES) {
    const html = read(file);

    assert.match(html, new RegExp(`<title>${title.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}</title>`), `${file} doit garder son titre`);
    assert.match(html, new RegExp(`<body data-page="${page}">`), `${file} doit garder data-page=${page}`);
    assert.match(html, APP_SCRIPT_PATTERN, `${file} doit charger la version courante de app.js`);
  }
});

test("pages PC: les liens et ressources locales referencees existent", () => {
  const missing = [];

  for (const { file } of DESKTOP_PAGES) {
    for (const ref of localRefsFromHtml(read(file))) {
      if (ref === "about:blank") {
        continue;
      }
      const target = path.join(ROOT, ref);
      if (!fs.existsSync(target)) {
        missing.push(`${file} -> ${ref}`);
      }
    }
  }

  assert.deepEqual(missing, []);
});

test("pages PC: les identifiants HTML restent uniques par page", () => {
  const duplicates = [];

  for (const { file } of DESKTOP_PAGES) {
    const seen = new Set();
    for (const id of idsFromHtml(read(file))) {
      if (seen.has(id)) {
        duplicates.push(`${file} -> ${id}`);
      }
      seen.add(id);
    }
  }

  assert.deepEqual(duplicates, []);
});

test("pages PC: les labels de formulaire pointent vers des champs existants", () => {
  const missing = [];

  for (const { file } of DESKTOP_PAGES) {
    const html = read(file);
    const ids = new Set(idsFromHtml(html));
    for (const target of labelTargetsFromHtml(html)) {
      if (!ids.has(target)) {
        missing.push(`${file} -> ${target}`);
      }
    }
  }

  assert.deepEqual(missing, []);
});

test("pages PC: les pages principales conservent la structure commune", () => {
  for (const { file } of SHELL_PAGES) {
    const html = read(file);

    assert.match(html, /<div class="app-shell">/, `${file} doit garder le shell principal`);
    assert.match(html, /<aside class="sidebar">/, `${file} doit garder la sidebar`);
    assert.match(html, /<main class="main-content">/, `${file} doit garder le contenu principal`);
    assert.match(html, /id="data-status"/, `${file} doit afficher l'etat des donnees`);
    assert.match(html, /id="dirty-status"/, `${file} doit afficher l'etat de sauvegarde`);
    assert.match(html, /class="button[^"]*js-load-data/, `${file} doit garder le bouton recharger`);
    assert.match(html, /class="button[^"]*js-save-data/, `${file} doit garder le bouton sauvegarder`);
  }
});

test("parcours PC: les commandes metier critiques restent presentes", () => {
  const overview = read("index.html");
  const sheet = read("fiche-personne.html");
  const arrival = read("document-arrivee.html");
  const exit = read("document-sortie.html");
  const archives = read("documents-archives.html");
  const references = read("bases-reference.html");
  const mobileSignature = read("signature-mobile.html");

  assert.match(overview, /id="overview-table-body"/);
  assert.match(overview, /id="overview-alerts-list"/);

  assert.match(sheet, /id="sheet-open-arrival-document"/);
  assert.match(sheet, /id="sheet-open-exit-document"/);
  assert.match(sheet, /id="effect-form"/);

  assert.match(arrival, /data-doc-type="arrival"/);
  assert.match(arrival, /data-signer="personnel"/);
  assert.match(arrival, /data-signer="representant"/);
  assert.match(exit, /data-doc-type="exit"/);
  assert.match(exit, /data-signer="personnel"/);
  assert.match(exit, /data-signer="representant"/);

  assert.match(archives, /id="documents-archives-body"/);
  assert.match(archives, /name="archiveStatutSignature"/);

  assert.match(references, /id="reference-representantsSignataires-body"/);
  assert.match(references, /id="replacement-costs-body"/);
  assert.match(references, /id="stock-summary-table-body"/);

  assert.match(mobileSignature, /id="mobile-signature-request-status"/);
  assert.match(mobileSignature, /class="signature-box__canvas js-signature-canvas"/);
  assert.match(mobileSignature, /class="button button--primary js-signature-save"/);
});

test("archives PC: la page affiche un wording de production", () => {
  const archives = read("documents-archives.html");

  assert.match(archives, /SUIVI DU STOCKAGE STRUCTURE DES PDF SIGNES/);
  assert.match(archives, /RECHERCHE DES PDF ARCHIVES PAR PERSONNE/);
  assert.match(archives, /RETROUVER ET OUVRIR LES PDF STOCKES/);
  assert.doesNotMatch(archives, /FUTURS PDF ARCHIVES/);
  assert.doesNotMatch(archives, /SERVIRA A TERME/);
  assert.doesNotMatch(archives, /PREPARATION DU STOCKAGE/);
});
