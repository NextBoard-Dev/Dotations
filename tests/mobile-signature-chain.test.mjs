import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import vm from "node:vm";

function extractFunctionSource(source, name) {
  let start = source.indexOf(`async function ${name}(`);
  if (start < 0) start = source.indexOf(`function ${name}(`);
  if (start < 0) throw new Error(`Function not found: ${name}`);
  const sig = source.indexOf("(", start);
  if (sig < 0) throw new Error(`Function signature not found: ${name}`);
  let parenDepth = 0;
  let bodyStart = -1;
  for (let j = sig; j < source.length; j += 1) {
    const ch = source[j];
    if (ch === "(") parenDepth += 1;
    if (ch === ")") {
      parenDepth -= 1;
      if (parenDepth === 0) {
        bodyStart = source.indexOf("{", j);
        break;
      }
    }
  }
  if (bodyStart < 0) throw new Error(`Function body not found: ${name}`);
  let i = bodyStart;
  let depth = 0;
  for (; i < source.length; i += 1) {
    const ch = source[i];
    if (ch === "{") depth += 1;
    if (ch === "}") {
      depth -= 1;
      if (depth === 0) return source.slice(start, i + 1);
    }
  }
  throw new Error(`Function extraction failed: ${name}`);
}

function createSignatureContext(fetchImpl = async () => ({ ok: true, json: async () => [] })) {
  const source = fs.readFileSync("app.js", "utf8");
  const fnNames = [
    "normalizeText",
    "normalizeHttpUrl",
    "isSupabaseConfigured",
    "getSupabaseHeaders",
    "normalizeMobileSignatureSigner",
    "isSupabaseSignatureSchemaMismatch",
    "getMobileSignatureRequestTokensForContext",
    "fetchSupabaseMobileSignatureRows",
    "getSupabaseSignatureRowField",
    "mergeSupabaseMobileSignatureRows",
  ];
  const context = {
    SUPABASE_PROJECT_URL: "https://example.supabase.co",
    SUPABASE_PUBLISHABLE_KEY: "sb_publishable_test_key",
    SUPABASE_ANON_KEY: "anon-key",
    SUPABASE_ACCESS_TOKEN: "",
    DEFAULT_SUPABASE_SIGNATURES_BUCKET: "signatures",
    fetch: fetchImpl,
    encodeURIComponent,
    URL,
    Date,
    String,
    Array,
    Set,
    Map,
    Object,
    Boolean,
    RegExp,
    console,
    getStoredSupabaseAccessToken: () => "",
  };
  vm.createContext(context);
  for (const name of fnNames) {
    vm.runInContext(extractFunctionSource(source, name), context);
  }
  return context;
}

test("signature mobile: la lecture par token bascule vers la lecture par document si le token ne retourne rien", async () => {
  const calls = [];
  const documentRows = [{ token: "SIG-1", person_id: "P1", doc_type: "arrival", signer: "personnel", signature_data: "data:image/png;base64,AAA", validated_at_text: "2026-10-07T10:00:00" }];
  const ctx = createSignatureContext(async (url, options) => {
    calls.push({ url, body: JSON.parse(options.body || "{}") });
    if (String(url).includes("fetch_mobile_signature_rows_for_document")) {
      return { ok: true, json: async () => documentRows };
    }
    return { ok: true, json: async () => [] };
  });

  const rows = await ctx.fetchSupabaseMobileSignatureRows("P1", "arrival", ["SIG-UNKNOWN"]);

  assert.equal(rows, documentRows);
  assert.equal(calls.length, 2);
  assert.match(calls[0].url, /fetch_mobile_signature_rows$/);
  assert.deepEqual(calls[0].body, { p_tokens: ["SIG-UNKNOWN"] });
  assert.match(calls[1].url, /fetch_mobile_signature_rows_for_document$/);
  assert.deepEqual(calls[1].body, { p_person_id: "P1", p_doc_type: "arrival" });
});

test("signature mobile: la lecture par document fonctionne sans demande active ni token connu", async () => {
  const calls = [];
  const documentRows = [{ token: "SIG-2", person_id: "P2", doc_type: "exit", signer: "representant", signature_data: "data:image/png;base64,BBB", validated_at_text: "2026-10-07T11:00:00" }];
  const ctx = createSignatureContext(async (url, options) => {
    calls.push({ url, body: JSON.parse(options.body || "{}") });
    return { ok: true, json: async () => documentRows };
  });

  const rows = await ctx.fetchSupabaseMobileSignatureRows("P2", "EXIT", []);

  assert.equal(rows, documentRows);
  assert.equal(calls.length, 1);
  assert.match(calls[0].url, /fetch_mobile_signature_rows_for_document$/);
  assert.deepEqual(calls[0].body, { p_person_id: "P2", p_doc_type: "exit" });
});

test("signature mobile: le secours REST ancien format utilise les colonnes camelCase", async () => {
  const calls = [];
  const ctx = createSignatureContext(async (url) => {
    calls.push(String(url));
    if (String(url).includes("fetch_mobile_signature_rows_for_document")) {
      return { ok: false, text: async () => "rpc missing", json: async () => [] };
    }
    if (String(url).includes("person_id=eq.")) {
      return { ok: false, text: async () => "42703 column person_id does not exist" };
    }
    return { ok: true, json: async () => [{ personId: "P3", docType: "arrival", signer: "personnel", signatureData: "data:image/png;base64,CCC", validatedAt: "2026-10-07T12:00:00" }] };
  });

  const rows = await ctx.fetchSupabaseMobileSignatureRows("P3", "arrival", []);

  assert.equal(rows.length, 1);
  assert.equal(calls.some((url) => url.includes("person_id=eq.P3")), true);
  assert.equal(calls.some((url) => url.includes("personId=eq.P3") && url.includes("docType=eq.arrival")), true);
});

test("signature mobile: la fusion accepte les lignes snake_case et camelCase et conserve la plus recente", () => {
  const ctx = createSignatureContext();
  const data = {
    personnes: [
      {
        id: "P4",
        signatures: {
          arrival: {
            personnel: { image: "old", validatedAt: "2026-01-01T08:00:00" },
          },
        },
      },
    ],
    demandesSignatureMobile: [
      { token: "SIG-PERS", personId: "P4", docType: "ARRIVAL", signer: "PERSONNEL", status: "EN ATTENTE" },
      { token: "SIG-REP", personId: "P4", docType: "ARRIVAL", signer: "REPRESENTANT", status: "EN ATTENTE" },
    ],
  };
  const rows = [
    { token: "SIG-PERS", person_id: "P4", doc_type: "arrival", signer: "personnel", signature_data: "old-row", validated_at_text: "2026-01-01T09:00:00" },
    { token: "SIG-PERS", person_id: "P4", doc_type: "arrival", signer: "personnel", signature_data: "new-personnel", validated_at_text: "2026-10-07T10:00:00" },
    { token: "SIG-REP", personId: "P4", docType: "arrival", signer: "representant", signatureData: "new-representant", validatedAt: "2026-10-07T10:01:00" },
  ];

  const changed = ctx.mergeSupabaseMobileSignatureRows(data, rows, "P4", "arrival");

  assert.equal(changed, true);
  assert.equal(data.personnes[0].signatures.arrival.personnel.image, "new-personnel");
  assert.equal(data.personnes[0].signatures.arrival.representant.image, "new-representant");
  assert.equal(data.demandesSignatureMobile[0].status, "SIGNEE");
  assert.equal(data.demandesSignatureMobile[1].status, "SIGNEE");
});

test("signature mobile: les documents entree et sortie declenchent la reprise directe Supabase", () => {
  const source = fs.readFileSync("app.js", "utf8");
  assert.match(source, /syncDocumentMobileSignatureLinks\("arrival", person\.id\);\s*void syncHostedMobileSignaturesForDocument\("arrival", person\.id\);/);
  assert.match(source, /syncDocumentMobileSignatureLinks\("exit", person\.id\);\s*void syncHostedMobileSignaturesForDocument\("exit", person\.id\);/);
  assert.match(source, /documentMobileSignatureSyncInFlight: new Set\(\)/);
});

test("signature mobile: chaque reprise distante force le rafraichissement du document courant", () => {
  const source = fs.readFileSync("app.js", "utf8");
  const refreshSource = extractFunctionSource(source, "refreshDocumentAfterMobileSignatureMerge");
  const pollSource = extractFunctionSource(source, "pollMobileSignatureRequest");
  const syncSource = extractFunctionSource(source, "syncHostedMobileSignaturesForDocument");

  assert.match(refreshSource, /state\.documentViewRenderCache\[normalizedDocType\] = "";/);
  assert.match(refreshSource, /state\.pageRenderSignature = "";/);
  assert.match(refreshSource, /renderArrivalDocument\(normalizedPersonId\);/);
  assert.match(refreshSource, /renderExitDocument\(normalizedPersonId\);/);
  assert.match(refreshSource, /refreshDocumentSignatureCanvases\(normalizedDocType, getPersonById\(normalizedPersonId\) \|\| person\);/);
  assert.match(source, /mobileSignaturePollUiRefreshSignature: ""/);
  assert.match(pollSource, /pullPayload\.changed[\s\S]*?refreshDocumentAfterMobileSignatureMerge\(docType, personId, null, \{ force: true \}\);/);
  assert.match(pollSource, /mergeSupabaseMobileSignatureRows\(json, signatureRows, personId, docType\)[\s\S]*?refreshDocumentAfterMobileSignatureMerge\(docType, personId, null, \{ force: true \}\);/);
  assert.match(pollSource, /state\.mobileSignaturePollStateSignature === pollStateSignature[\s\S]*?refreshDocumentAfterMobileSignatureMerge\(docType, personId, person\);/);
  assert.match(syncSource, /refreshDocumentAfterMobileSignatureMerge\(normalizedDocType, normalizedPersonId, null, \{ force: true \}\);/);
});

test("signature mobile: la page autonome conserve les points d'ancrage critiques", () => {
  const html = fs.readFileSync("signature-mobile.html", "utf8");

  assert.match(html, /<body data-page="mobile-signature">/);
  assert.match(html, /id="mobile-signature-request-status">CHARGEMENT\.\.\.<\/p>/);
  assert.match(html, /id="mobile-signature-panel"/);
  assert.match(html, /class="signature-box__canvas js-signature-canvas" data-doc-type="arrival" data-signer="personnel"/);
  assert.match(html, /class="button button--primary js-signature-save" data-doc-type="arrival" data-signer="personnel"/);
  assert.match(html, /id="mobile-signature-personnel-status">EN ATTENTE DE SIGNATURE<\/p>/);
  assert.match(html, /id="mobile-costs-head"/);
  assert.match(html, /id="mobile-costs-body"/);
  assert.match(html, /href="favicon\.ico\?v=20260928b"/);
  assert.match(html, /src="app\.js\?v=20261009-mobile-signature-refresh"/);
});

test("signature mobile: le bouton reste verrouille apres validation locale", () => {
  const source = fs.readFileSync("app.js", "utf8");

  assert.match(source, /const showValidatedSignatureState = \(validatedAt = ""\) => \{/);
  assert.match(source, /statusNode\.textContent = validatedLabel\s*\?\s*`SIGNATURE ENREGISTREE LE \$\{validatedLabel\}`\s*:\s*"SIGNATURE ENREGISTREE";/);
  assert.match(source, /saveButton\.disabled = true;\s*saveButton\.classList\.add\("is-disabled"\);\s*saveButton\.classList\.add\("button--validated"\);\s*saveButton\.textContent = "VALIDE";/);
  assert.match(source, /clearButton\.disabled = true;\s*clearButton\.classList\.add\("is-disabled"\);/);
  assert.match(source, /showValidatedSignatureState\(validatedAt\);/);
  assert.match(source, /showDataStatus\("SIGNATURE ENREGISTREE - VOUS POUVEZ FERMER CETTE PAGE"\);/);
});

test("signature mobile: le relais Supabase est sauvegarde avant la reprise app_state", () => {
  const source = fs.readFileSync("app.js", "utf8");
  const relaySaveIndex = source.indexOf("signatureRecordSaved = await saveMobileSignatureRecordToSupabase");
  const rebaseSaveIndex = source.indexOf("latestPayload = await saveSupabaseSignatureWithRebase", relaySaveIndex);
  const relayFallbackIndex = source.indexOf("if (!signatureRecordSaved) {", rebaseSaveIndex);

  assert.ok(relaySaveIndex > 0, "la sauvegarde relais signatures doit etre appelee");
  assert.ok(rebaseSaveIndex > relaySaveIndex, "la reprise app_state doit arriver apres le relais");
  assert.ok(relayFallbackIndex > rebaseSaveIndex, "un second essai relais non bloquant doit rester present");
  assert.match(source, /if \(!signatureRecordSaved\) \{\s*throw stateSaveError;\s*\}/);
  assert.match(source, /showDataStatus\("SIGNATURE ENREGISTREE - SYNCHRONISATION DOCUMENT EN COURS"\);/);
  assert.match(source, /saveMobileSignatureRecordToSupabase\(\{\s*token: mobileRequestToken,/);
  assert.match(source, /saveSupabaseSignatureWithRebase\(\{\s*personId: person\.id,/);
});

test("signature mobile: le rendu reutilise la signature runtime pour eviter un retour visuel vide", () => {
  const source = fs.readFileSync("app.js", "utf8");

  assert.match(source, /function getMobileSignatureRuntimeKey\(request, person, docType, signer\)/);
  assert.match(source, /return \[token, personId, normalizedDocType, normalizedSigner\]\.join\("\|\|\|"\);/);
  assert.match(source, /function rememberMobileSignatureRuntimeSignature\(request, person, docType, signer, image, validatedAt, storageRef = "", storagePublicUrl = ""\)/);
  assert.match(source, /function getMobileSignatureRuntimeSignature\(request, person, docType, signer\)/);
  assert.match(source, /const runtimeSignature = getMobileSignatureRuntimeSignature\(request, person, normalizedDocType, signer\);/);
  assert.match(source, /const signatureValue = String\(runtimeSignature\?\.image \|\| person\?\.signatures\?\.\[normalizedDocType\]\?\.\[signer\]\?\.image \|\| ""\);/);
  assert.match(source, /const isAlreadySigned = Boolean\(\(request && normalizeText\(request\.status\) === "SIGNEE"\) \|\| runtimeSignature\?\.image\);/);
  assert.match(source, /if \(isAlreadySigned\) \{\s*saveButton\.disabled = true;/);
});
