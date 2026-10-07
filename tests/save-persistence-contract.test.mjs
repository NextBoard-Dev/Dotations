import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";

const appSource = fs.readFileSync("app.js", "utf8");

function functionBlock(name) {
  const start =
    appSource.indexOf(`async function ${name}(`) >= 0
      ? appSource.indexOf(`async function ${name}(`)
      : appSource.indexOf(`function ${name}(`);
  assert.ok(start >= 0, `fonction introuvable: ${name}`);

  const signatureStart = appSource.indexOf("(", start);
  assert.ok(signatureStart > start, `signature introuvable: ${name}`);
  let parenDepth = 0;
  let signatureEnd = -1;
  for (let index = signatureStart; index < appSource.length; index += 1) {
    const char = appSource[index];
    if (char === "(") parenDepth += 1;
    if (char === ")") {
      parenDepth -= 1;
      if (parenDepth === 0) {
        signatureEnd = index;
        break;
      }
    }
  }
  assert.ok(signatureEnd > signatureStart, `fin de signature introuvable: ${name}`);

  const bodyStart = appSource.indexOf("{", signatureEnd);
  assert.ok(bodyStart > start, `corps introuvable: ${name}`);

  let depth = 0;
  for (let index = bodyStart; index < appSource.length; index += 1) {
    const char = appSource[index];
    if (char === "{") depth += 1;
    if (char === "}") {
      depth -= 1;
      if (depth === 0) return appSource.slice(start, index + 1);
    }
  }
  throw new Error(`extraction impossible: ${name}`);
}

test("sauvegarde: une seule persistance a la fois et aucun ecrasement sans changement", () => {
  const block = functionBlock("saveDataToFile");

  const inFlightIndex = block.indexOf("if (state.saveInFlight)");
  const signatureIndex = block.indexOf("const currentDataSignature = computeDataPersistenceSignature(state.data);");
  const noChangeIndex = block.indexOf("if (!hasEffectiveChange)");
  const persistIndex = block.indexOf("await saveSupabasePayloadWithRetry(state.data, 3);");

  assert.ok(inFlightIndex >= 0, "verrou sauvegarde en cours absent");
  assert.ok(signatureIndex > inFlightIndex, "signature de donnees absente apres verrou");
  assert.ok(noChangeIndex > signatureIndex, "detection aucun changement absente");
  assert.ok(persistIndex > noChangeIndex, "persistance appelee avant controle des changements");
  assert.match(block, /showDataStatus\("SAUVEGARDE EN COURS\.\.\."\);/);
  assert.match(block, /showDataStatus\("AUCUNE MODIFICATION DETECTEE\. RIEN A SAUVEGARDER\."\);/);
  assert.match(block, /outcome: "noop"/);
  assert.match(block, /state\.lastPersistedDataSignature = computeDataPersistenceSignature\(state\.data\);/);
});

test("sauvegarde: signatures et archives protegees avant toute ecriture backend", () => {
  const block = functionBlock("saveDataToFile");

  const signatureSnapshotIndex = block.indexOf("const signatureSnapshot = buildSignatureProtectionSnapshot(state.data);");
  const archiveSnapshotIndex = block.indexOf("const archiveSnapshot = buildArchiveProtectionSnapshot(state.data);");
  const normalizeArchivesIndex = block.indexOf("protectAndNormalizeArchivesInPlace(state.data);");
  const enforceSignaturesIndex = block.indexOf("enforceProtectedSignaturesInPlace(state.data, signatureSnapshot);");
  const enforceArchivesIndex = block.indexOf("enforceProtectedArchivesInPlace(state.data, archiveSnapshot);");
  const supabasePersistIndex = block.indexOf("await saveSupabasePayloadWithRetry(state.data, 3);");
  const localPersistIndex = block.indexOf('await fetch("/api/state/save"');

  assert.ok(signatureSnapshotIndex >= 0, "snapshot signatures absent");
  assert.ok(archiveSnapshotIndex > signatureSnapshotIndex, "snapshot archives doit suivre signatures");
  assert.ok(normalizeArchivesIndex > archiveSnapshotIndex, "normalisation archives absente");
  assert.ok(enforceSignaturesIndex > normalizeArchivesIndex, "reprotection signatures absente");
  assert.ok(enforceArchivesIndex > enforceSignaturesIndex, "reprotection archives absente");
  assert.ok(supabasePersistIndex > enforceArchivesIndex, "Supabase avant protections");
  assert.ok(localPersistIndex > enforceArchivesIndex, "sauvegarde locale avant protections");
});

test("sauvegarde Supabase: un conflit ne relance pas le meme payload stale", () => {
  const retryBlock = functionBlock("saveSupabasePayloadWithRetry");
  const saveBlock = functionBlock("saveDataToFile");

  assert.match(retryBlock, /for \(let attempt = 0; attempt < maxAttempts; attempt \+= 1\)/);
  assert.match(retryBlock, /expectedRevision: Number\(state\.supabaseRevision\)/);
  assert.match(retryBlock, /Do not retry the same stale payload with a fresh revision/);
  assert.match(retryBlock, /await fetchSupabaseStateData\(\);\s*throw error;/);
  assert.match(retryBlock, /throw buildSaveConflictError\(\);/);

  assert.match(saveBlock, /if \(isSaveConflictError\(error\)\) \{/);
  assert.match(saveBlock, /if \(throwOnConflict\) \{\s*throw error;\s*\}/);
  assert.match(saveBlock, /await reloadData\("CONFLIT DETECTE - RECHARGEMENT DES DONNEES DISTANTES\.\.\."\);/);
  assert.match(saveBlock, /"CONFLIT DE SAUVEGARDE - RECHARGER PUIS REESSAYER"/);
});

test("sauvegarde locale: l'envoi heberge reste bloque tant que le controle local n'est pas OK", () => {
  const saveBlock = functionBlock("saveDataToFile");
  const controlBlock = functionBlock("runHostedLocalControlAfterSave");

  assert.match(controlBlock, /fetch\("\/api\/local\/check-before-send"/);
  assert.match(controlBlock, /const localOk = Boolean\(payload\?\.localOk\);/);
  assert.match(controlBlock, /const pdfOk = Boolean\(payload\?\.pdfIntegrityOk\);/);
  assert.match(controlBlock, /state\.hostedSyncState = supabaseOk && \(!localOk \|\| !pdfOk\) \? "blocked" : "inaccessible";/);
  assert.match(controlBlock, /showDataStatus\("SAUVEGARDE LOCALE OK - CONTROLE LOCAL A CORRIGER"\);/);
  assert.match(controlBlock, /finally \{\s*renderHostedSyncUi\(\);\s*\}/);

  assert.match(saveBlock, /const localControlOk = await runHostedLocalControlAfterSave\(\);/);
  assert.match(saveBlock, /if \(localControlOk && autoPushHosted\) \{\s*await runHostedSyncPushFromUi\(\{ confirmBeforeSend: false, silent: true \}\);/);
});

test("journal de sauvegarde: l'audit est borne et ne peut pas bloquer la sauvegarde", () => {
  const block = functionBlock("appendSaveAuditEntry");

  assert.match(appSource, /const SAVE_AUDIT_LOG_KEY = "dotations-save-audit-log-v1"/);
  assert.match(appSource, /const MAX_SAVE_AUDIT_ENTRIES = 250/);
  assert.match(block, /entries\.push\(\{\s*at: getCurrentSignatureTimestamp\(\),/);
  assert.match(block, /const compact = entries\.slice\(-MAX_SAVE_AUDIT_ENTRIES\);/);
  assert.match(block, /localStorage\.setItem\(SAVE_AUDIT_LOG_KEY, JSON\.stringify\(compact\)\);/);
  assert.match(block, /catch \(error\) \{\s*\/\/ Silent: audit must never block save\./);
});
