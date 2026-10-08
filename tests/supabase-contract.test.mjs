import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";

const appSource = fs.readFileSync("app.js", "utf8");
const sqlSource = fs.readFileSync("sql/correctif_immediat_signature_mobile_supabase.sql", "utf8");
const sqlFiles = fs
  .readdirSync("sql")
  .filter((file) => file.endsWith(".sql"))
  .sort();

function readSql(file) {
  return fs.readFileSync(path.join("sql", file), "utf8");
}

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

test("contrat Supabase SQL: les scripts restent bornes et sans secret serveur", () => {
  const errors = [];
  const serverRolePattern = new RegExp(["\\bservice", "_", "role\\b"].join(""), "i");

  for (const file of sqlFiles) {
    const source = readSql(file);
    if (serverRolePattern.test(source)) {
      errors.push(`${file}: reference cle serveur interdite dans les scripts versionnes`);
    }
    if (/\bauth\.role\s*\(/i.test(source)) {
      errors.push(`${file}: auth.role() interdit, utiliser les policies TO explicites`);
    }
    if (/\braw_user_meta_data\b/i.test(source)) {
      errors.push(`${file}: raw_user_meta_data interdit pour les decisions de securite`);
    }
    if (/\bdrop\s+table\b|\bdrop\s+column\b|\btruncate\b/i.test(source)) {
      errors.push(`${file}: operation destructive interdite sans procedure dediee`);
    }
  }

  assert.deepEqual(errors, []);
});

test("contrat Supabase SQL: les fonctions security definer cadrent le search_path", () => {
  const errors = [];
  const functionPattern =
    /create\s+or\s+replace\s+function\s+public\.([a-zA-Z0-9_]+)\s*\([\s\S]*?\$\$;/gi;

  for (const file of sqlFiles) {
    const source = readSql(file);
    let match;
    while ((match = functionPattern.exec(source))) {
      const block = match[0];
      const functionName = match[1];
      if (/security\s+definer/i.test(block) && !/set\s+search_path\s*=\s*public/i.test(block)) {
        errors.push(`${file}: public.${functionName} est security definer sans search_path public`);
      }
    }
  }

  assert.deepEqual(errors, []);
});
