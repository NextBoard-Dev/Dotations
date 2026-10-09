import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";

const ROOT = process.cwd();
const DESKTOP_ENTRYPOINTS = [
  "index.html",
  "fiche-personne.html",
  "document-arrivee.html",
  "document-sortie.html",
  "documents-archives.html",
  "bases-reference.html",
  "signature-mobile.html",
];

function read(file) {
  return fs.readFileSync(path.join(ROOT, file), "utf8");
}

function trackedFiles() {
  return execFileSync("git", ["ls-files"], { cwd: ROOT, encoding: "utf8" })
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
}

function refsFromHtml(html) {
  const refs = [];
  const pattern = /\b(?:href|src)=["']([^"']+)["']/g;
  let match;
  while ((match = pattern.exec(html))) {
    const value = match[1].trim();
    const clean = value.split("#")[0].split("?")[0].replace(/^\.\//, "");
    if (
      clean &&
      !clean.startsWith("/") &&
      !clean.startsWith("http://") &&
      !clean.startsWith("https://") &&
      !clean.startsWith("data:") &&
      !clean.startsWith("mailto:")
    ) {
      refs.push(clean);
    }
  }
  return refs;
}

function mobileBundleRefs() {
  const html = read("mobile/index.html");
  const js = Array.from(html.matchAll(/src="\.\/assets\/([^"]+\.js)"/g), (match) => match[1]);
  const css = Array.from(html.matchAll(/href="\.\/assets\/([^"]+\.css)"/g), (match) => match[1]);
  return { js, css };
}

test("publication: les points d'entree publics chargent les fichiers publies attendus", () => {
  for (const file of DESKTOP_ENTRYPOINTS) {
    const html = read(file);
    assert.match(html, /<link rel="icon" href="favicon\.ico\?v=20260928b"/, `${file} doit exposer le favicon courant`);
    assert.match(html, /<link rel="stylesheet" href="style\.css\?v=20260624-pdf-sync-hidden"/, `${file} doit charger le CSS publie`);
    assert.match(html, /<script src="app\.js\?v=20261009-mobile-signature-refresh"><\/script>/, `${file} doit charger app.js publie`);
  }
});

test("publication mobile: l'index publie référence exactement le bundle disponible", () => {
  const { js, css } = mobileBundleRefs();

  assert.equal(js.length, 1, "mobile/index.html doit charger un seul bundle JS");
  assert.equal(css.length, 1, "mobile/index.html doit charger un seul bundle CSS");
  assert.equal(fs.existsSync(path.join(ROOT, "mobile", "assets", js[0])), true, `bundle JS absent: ${js[0]}`);
  assert.equal(fs.existsSync(path.join(ROOT, "mobile", "assets", css[0])), true, `bundle CSS absent: ${css[0]}`);
});

test("publication mobile: aucun ancien bundle asset ne reste publie", () => {
  const { js, css } = mobileBundleRefs();
  const expectedAssets = new Set([...js, ...css]);
  const actualAssets = fs
    .readdirSync(path.join(ROOT, "mobile", "assets"))
    .filter((file) => /\.(?:js|css)$/i.test(file))
    .sort();

  assert.deepEqual(actualAssets, Array.from(expectedAssets).sort());
});

test("publication: toutes les references locales des entrees publiees existent", () => {
  const missing = [];
  const htmlFiles = [...DESKTOP_ENTRYPOINTS, "mobile/index.html"];

  for (const file of htmlFiles) {
    const baseDir = path.dirname(path.join(ROOT, file));
    for (const ref of refsFromHtml(read(file))) {
      const target = path.resolve(baseDir, ref.replace(/\//g, path.sep));
      if (!fs.existsSync(target)) {
        missing.push(`${file} -> ${ref}`);
      }
    }
  }

  assert.deepEqual(missing, []);
});

test("publication: aucun etat local Supabase CLI ne doit etre versionne", () => {
  const offenders = trackedFiles().filter((file) => /^supabase\/\.temp\//i.test(file));
  assert.deepEqual(offenders, []);
  assert.match(read(".gitignore"), /supabase\/\.temp\//);
});
