import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";

const ROOT = process.cwd();
const SMARTPHONE_SRC = path.join(ROOT, "smartphone", "src");

function read(file) {
  return fs.readFileSync(path.join(ROOT, file), "utf8");
}

function existsAny(basePath) {
  const candidates = [
    basePath,
    `${basePath}.js`,
    `${basePath}.jsx`,
    `${basePath}.ts`,
    `${basePath}.tsx`,
    `${basePath}.css`,
    path.join(basePath, "index.js"),
    path.join(basePath, "index.jsx"),
    path.join(basePath, "index.ts"),
    path.join(basePath, "index.tsx"),
  ];
  return candidates.some((candidate) => fs.existsSync(candidate));
}

function trackedFiles(...extensions) {
  const allowed = new Set(extensions.map((ext) => ext.toLowerCase()));
  return execFileSync("git", ["ls-files"], { cwd: ROOT, encoding: "utf8" })
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .filter((file) => allowed.has(path.extname(file).toLowerCase()));
}

function importSpecifiers(source) {
  const specs = [];
  const patterns = [
    /\bimport\s+(?:[^'"]+?\s+from\s+)?["']([^"']+)["']/g,
    /\bimport\s*\(\s*["']([^"']+)["']\s*\)/g,
    /\bexport\s+[^'"]+?\s+from\s+["']([^"']+)["']/g,
  ];
  for (const pattern of patterns) {
    let match;
    while ((match = pattern.exec(source))) {
      specs.push(match[1]);
    }
  }
  return specs;
}

function resolveSmartphoneSpecifier(importer, specifier) {
  if (specifier.startsWith("@/")) {
    return path.join(SMARTPHONE_SRC, specifier.slice(2));
  }
  if (specifier.startsWith("./") || specifier.startsWith("../")) {
    return path.resolve(path.dirname(path.join(ROOT, importer)), specifier);
  }
  return null;
}

function htmlLocalRefs(source) {
  const refs = [];
  const pattern = /\b(?:href|src)=["']([^"']+)["']/g;
  let match;
  while ((match = pattern.exec(source))) {
    const raw = match[1].trim();
    if (
      !raw ||
      raw.startsWith("#") ||
      raw.startsWith("data:") ||
      raw.startsWith("mailto:") ||
      /^https?:\/\//i.test(raw)
    ) {
      continue;
    }
    refs.push(raw.split("#")[0].split("?")[0]);
  }
  return refs;
}

function resolveHtmlRef(htmlFile, ref) {
  const normalized = ref.replace("%BASE_URL%", "").replace(/^\.\//, "");
  if (htmlFile === "smartphone/index.html" && normalized.startsWith("/src/")) {
    return path.join(ROOT, "smartphone", normalized.slice(1));
  }
  if (htmlFile.startsWith("smartphone/")) {
    return path.join(ROOT, "smartphone", "public", normalized.replace(/^\//, ""));
  }
  if (ref.startsWith("/")) {
    return path.join(ROOT, ref.slice(1));
  }
  return path.join(ROOT, path.dirname(htmlFile), normalized);
}

test("modules smartphone: tous les imports locaux resolvent vers un fichier existant", () => {
  const errors = [];
  const files = trackedFiles(".js", ".jsx", ".ts", ".tsx").filter((file) =>
    file.replace(/\\/g, "/").startsWith("smartphone/src/")
  );

  for (const file of files) {
    for (const specifier of importSpecifiers(read(file))) {
      const target = resolveSmartphoneSpecifier(file, specifier);
      if (target && !existsAny(target)) {
        errors.push(`${file} -> ${specifier}`);
      }
    }
  }

  assert.deepEqual(errors, []);
});

test("configuration smartphone: l'alias source et la sortie mobile restent coherents", () => {
  const vite = read("smartphone/vite.config.js");
  const jsconfig = JSON.parse(read("smartphone/jsconfig.json"));

  assert.match(vite, /base: "\.\/"/);
  assert.match(vite, /outDir: "\.\.\/mobile"/);
  assert.match(vite, /find: "@", replacement: fileURLToPath\(new URL\("\.\/src", import\.meta\.url\)\)/);
  assert.deepEqual(jsconfig.compilerOptions?.paths?.["@/*"], ["./src/*"]);
});

test("HTML: toutes les references locales versionnees pointent vers des fichiers presents", () => {
  const htmlFiles = trackedFiles(".html");
  const errors = [];

  for (const file of htmlFiles) {
    for (const ref of htmlLocalRefs(read(file))) {
      const target = resolveHtmlRef(file, ref);
      if (!fs.existsSync(target)) {
        errors.push(`${file} -> ${ref}`);
      }
    }
  }

  assert.deepEqual(errors, []);
});

test("mobile publie: index pointe vers un bundle JS et CSS existants", () => {
  const html = read("mobile/index.html");
  const js = html.match(/src="\.\/assets\/([^"]+\.js)"/)?.[1] || "";
  const css = html.match(/href="\.\/assets\/([^"]+\.css)"/)?.[1] || "";

  assert.ok(js, "bundle JS mobile absent de mobile/index.html");
  assert.ok(css, "bundle CSS mobile absent de mobile/index.html");
  assert.ok(fs.existsSync(path.join(ROOT, "mobile", "assets", js)), `bundle JS introuvable: ${js}`);
  assert.ok(fs.existsSync(path.join(ROOT, "mobile", "assets", css)), `bundle CSS introuvable: ${css}`);
});
