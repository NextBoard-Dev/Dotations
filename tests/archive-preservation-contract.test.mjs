import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";

const appSource = fs.readFileSync("app.js", "utf8");
const data = JSON.parse(fs.readFileSync("data.json", "utf8"));

function normalizeText(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ");
}

function isSafeArchivePath(value) {
  const raw = String(value || "").trim();
  return Boolean(raw && !/^\s*(?:https?:)?\/\//i.test(raw) && !raw.includes("..") && /\.pdf(?:$|\?)/i.test(raw));
}

test("archives donnees: les documents archives restent rattaches et ouvrables", () => {
  const peopleById = new Set((data.personnes || []).map((person) => String(person?.id || "")).filter(Boolean));
  const archives = Array.isArray(data.documentsArchives) ? data.documentsArchives : [];
  const errors = [];

  for (const entry of archives) {
    const id = String(entry?.id || "");
    const personId = String(entry?.personId || "");
    const typeDocument = normalizeText(entry?.typeDocument);
    const pdfPath = String(entry?.pdfPath || "");
    if (!id) errors.push("archive sans id");
    if (!peopleById.has(personId)) errors.push(`${id || "(sans id)"} personId inconnu: ${personId}`);
    if (!["ARRIVEE", "SORTIE"].includes(typeDocument)) errors.push(`${id || "(sans id)"} typeDocument invalide: ${entry?.typeDocument}`);
    if (!isSafeArchivePath(pdfPath)) errors.push(`${id || "(sans id)"} pdfPath invalide: ${pdfPath}`);
  }

  assert.deepEqual(errors, []);
});

test("archives sauvegarde: les signatures et PDF signes sont proteges avant persistance", () => {
  const saveBlockStart = appSource.indexOf("const signatureSnapshot = buildSignatureProtectionSnapshot(state.data);");
  const saveBlock = appSource.slice(saveBlockStart, saveBlockStart + 3000);
  const signatureSnapshotIndex = saveBlock.indexOf("const signatureSnapshot = buildSignatureProtectionSnapshot(state.data);");
  const archiveSnapshotIndex = saveBlock.indexOf("const archiveSnapshot = buildArchiveProtectionSnapshot(state.data);");
  const normalizeArchivesIndex = saveBlock.indexOf("protectAndNormalizeArchivesInPlace(state.data);");
  const enforceSignaturesIndex = saveBlock.indexOf("enforceProtectedSignaturesInPlace(state.data, signatureSnapshot);");
  const enforceArchivesIndex = saveBlock.indexOf("enforceProtectedArchivesInPlace(state.data, archiveSnapshot);");
  const supabasePersistIndex = saveBlock.indexOf("await saveSupabasePayloadWithRetry(state.data, 3);");
  const localPersistIndex = saveBlock.indexOf('await fetch("/api/state/save"');

  assert.ok(saveBlockStart > 0, "bloc de sauvegarde protege absent");
  assert.ok(signatureSnapshotIndex >= 0, "snapshot signatures absent");
  assert.ok(archiveSnapshotIndex > signatureSnapshotIndex, "snapshot archives doit suivre le snapshot signatures");
  assert.ok(normalizeArchivesIndex > archiveSnapshotIndex, "normalisation archives absente avant sauvegarde");
  assert.ok(enforceSignaturesIndex > normalizeArchivesIndex, "reprotection signatures absente avant sauvegarde");
  assert.ok(enforceArchivesIndex > enforceSignaturesIndex, "reprotection archives absente avant sauvegarde");
  assert.ok(supabasePersistIndex > enforceArchivesIndex, "persistance Supabase avant protection archives");
  assert.ok(localPersistIndex > enforceArchivesIndex, "persistance locale avant protection archives");
});

test("archives sauvegarde: une archive signee ne peut pas disparaitre silencieusement", () => {
  assert.match(appSource, /function buildArchiveProtectionSnapshot\(data\)/);
  assert.match(appSource, /normalizeText\(entry\?\.statutSignature\) === "SIGNE"/);
  assert.match(appSource, /const key = `\$\{personId\}\|\$\{typeKey\}`/);
  assert.match(appSource, /snapshot\.set\(key, \{ \.\.\.entry, typeDocument: typeKey \}\)/);
  assert.match(appSource, /function enforceProtectedArchivesInPlace\(data, archiveSnapshot\)/);
  assert.match(appSource, /data\.documentsArchives\.push\(\{ \.\.\.savedEntry, typeDocument: typeKey \}\)/);
});

test("archives UI: ouvrir un PDF archive utilise une URL resolue et ne supprime aucun fichier", () => {
  assert.match(appSource, /function resolveArchiveOpenTarget\(entry\)/);
  assert.match(appSource, /function handleArchiveOpenClick\(archive, fallbackEntry = null\)/);
  assert.match(appSource, /window\.open\(pdfUrl, "_blank", "noopener"\)/);
  assert.match(appSource, /function deleteDocumentArchiveEntry\(archiveId\)/);
  assert.match(appSource, /window\.confirm\(/);
  assert.match(appSource, /state\.data\.documentsArchives = state\.data\.documentsArchives\.filter/);
  assert.doesNotMatch(appSource, /deleteDocumentArchiveEntry[\s\S]{0,1200}fetch\([^)]*DELETE/i);
});

test("archives PDF: les chemins d'ouverture restent bornes et verificables", () => {
  assert.match(appSource, /function resolveArchivePdfLocations\(pdfPath, existing = \{\}\)/);
  assert.match(appSource, /parseStorageSchemePath\(raw\)/);
  assert.match(appSource, /getSupabaseStoragePublicUrl\(storageRef\.bucket, storageRef\.objectPath\)/);
  assert.match(appSource, /isSafeArchiveHttpUrl\(raw\) \|\| isHostedPdfDocumentPath\(raw\)/);
  assert.match(appSource, /isSafeArchiveRelativePath\(raw\)/);
  assert.match(appSource, /localPath\.toLowerCase\(\)\.startsWith\("data\/pdf\/"\)/);
  assert.match(appSource, /openLocalUrl = `\/api\/pdf-file\?path=\$\{encodeURIComponent\(apiPath\)\}`/);

  assert.match(appSource, /function normalizeDirectPdfOpenUrl\(value\)/);
  assert.match(appSource, /const isPdfApi = pathname\.endsWith\("\/api\/pdf-file"\)/);
  assert.match(appSource, /const isPdfFile = pathname\.endsWith\("\.pdf"\)/);
  assert.match(appSource, /parsed\.searchParams\.get\("pdf"\) === "1"/);
  assert.match(appSource, /return isPdfApi \|\| isPdfFile \|\| isHostedPdfRender \? parsed\.href : ""/);
  assert.match(appSource, /function verifyActiveArchiveOpenable\(personId, docType\)/);
  assert.match(appSource, /fetch\(targetUrl, \{ method: "HEAD", cache: "no-store" \}\)/);
  assert.match(appSource, /ARCHIVE_FILE_HEAD_FAILED/);
});
