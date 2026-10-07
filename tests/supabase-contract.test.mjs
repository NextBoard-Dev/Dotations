import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";

const appSource = fs.readFileSync("app.js", "utf8");
const sqlSource = fs.readFileSync("sql/correctif_immediat_signature_mobile_supabase.sql", "utf8");

test("contrat Supabase signatures: les RPC appelees par l'UI sont versionnees", () => {
  const requiredFunctions = [
    "submit_mobile_signature",
    "fetch_mobile_signature_rows",
    "fetch_mobile_signature_rows_for_document",
  ];

  for (const name of requiredFunctions) {
    assert.match(appSource, new RegExp(`/rpc/${name}`), `app.js doit appeler ${name}`);
    assert.match(sqlSource, new RegExp(`function public\\.${name}\\b`, "i"), `le SQL doit creer ${name}`);
    assert.match(sqlSource, new RegExp(`grant execute on function public\\.${name}`, "i"), `le SQL doit accorder l'execution de ${name}`);
  }
});

test("contrat Supabase signatures: les colonnes snake_case et le secours camelCase restent compatibles", () => {
  const requiredColumns = [
    "token",
    "person_id",
    "doc_type",
    "signer",
    "status",
    "signature_data",
    "validated_at_text",
    "updated_at",
  ];

  for (const column of requiredColumns) {
    assert.match(sqlSource, new RegExp(`add column if not exists ${column}\\b`, "i"), `colonne SQL manquante: ${column}`);
  }

  assert.match(appSource, /getSupabaseSignatureRowField\(row, "person_id", "personId"\)/);
  assert.match(appSource, /getSupabaseSignatureRowField\(row, "doc_type", "docType"\)/);
  assert.match(appSource, /getSupabaseSignatureRowField\(row, "signature_data", "signatureData", "image"\)/);
});

test("contrat Supabase signatures: RLS et politiques minimales restent presentes", () => {
  assert.match(sqlSource, /alter table public\.signatures enable row level security/i);
  assert.match(sqlSource, /create policy "mobile_signature_select_by_token"/i);
  assert.match(sqlSource, /create policy "mobile_signature_insert_by_token"/i);
  assert.match(sqlSource, /create policy "mobile_signature_update_by_token"/i);
  assert.doesNotMatch(sqlSource, /security definer/i);
});
