import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import { execFileSync } from "node:child_process";

const appSource = fs.readFileSync("app.js", "utf8");
const styleSource = fs.readFileSync("style.css", "utf8");
const signatureMobileSource = fs.readFileSync("signature-mobile.html", "utf8");

test("garde-fou signature mobile: enregistrement puis reprise par document", () => {
  assert.match(appSource, /rpc\/submit_mobile_signature/);
  assert.match(appSource, /rpc\/fetch_mobile_signature_rows_for_document/);
  assert.match(appSource, /function mergeSupabaseMobileSignatureRows\(/);
  assert.match(appSource, /void syncHostedMobileSignaturesForDocument\("arrival", person\.id\);/);
  assert.match(appSource, /void syncHostedMobileSignaturesForDocument\("exit", person\.id\);/);
  assert.match(signatureMobileSource, /VALIDER LA SIGNATURE|signatureData|canvas|toDataURL/i);
});

test("garde-fou PDF sortie: changement de mise en page invalide les anciens PDF", () => {
  const versionMatch = appSource.match(/const PDF_LAYOUT_VERSION = "([^"]+)";/);
  assert.ok(versionMatch, "PDF_LAYOUT_VERSION doit exister");
  assert.equal(versionMatch[1], "2026-10-07-exit-table-compact-v1");
  assert.match(appSource, /layoutVersion: PDF_LAYOUT_VERSION/);
  assert.match(appSource, /function getDocumentFingerprint\(person, docType\)/);
});

test("garde-fou PDF sortie: tableau compact avec une seule ligne", () => {
  assert.match(styleSource, /body\[data-pdf-mode="true"\]\[data-page="exit-document"\] #exit-effects-body tr/);
  assert.match(styleSource, /body\[data-pdf-mode="true"\]\[data-page="exit-document"\] #exit-effects-body td/);
  assert.match(styleSource, /grid-template-columns: 15% 17% 13% 14% 14% 10% 9% 8%;/);
  assert.match(styleSource, /\.table-total-row \{\s*grid-template-columns: 1fr auto;/);
});

test("garde-fou depollution: aucun fichier de sauvegarde local ne doit etre versionne", () => {
  const forbiddenPatterns = [
    /\.bak/i,
    /~$/,
    /\.tmp$/i,
    /__pycache__/i,
    /node_modules/i,
    /_quarantine/i,
  ];
  const trackedLikeFiles = execFileSync("git", ["ls-files"], { encoding: "utf8" })
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  const offenders = trackedLikeFiles.filter((file) =>
    forbiddenPatterns.some((pattern) => pattern.test(file))
  );

  assert.deepEqual(offenders, []);
});
