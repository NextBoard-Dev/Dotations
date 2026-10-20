import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
const security=fs.readFileSync('sql/securisation_supabase_dotations.sql','utf8');
const vault=fs.readFileSync('sql/coffre_fort_supabase_app_state.sql','utf8');
test('Data API: journaux accessibles en lecture seulement avec politiques admin',()=>{
  for(const [source,table] of [[security,'audit_log'],[vault,'app_state_versions'],[vault,'app_state_save_journal']]){
    assert.ok(source.includes(`grant select on table public.${table} to authenticated;`));
    assert.ok(source.includes(`alter table public.${table} enable row level security;`));
    assert.match(source,/read_admin_only/);
    assert.doesNotMatch(source,new RegExp(`grant [^;]*on table public\\.${table} to anon`, 'i'));
  }
});
test('Data API: droits metier explicites sans DELETE ni ouverture anonyme',()=>{
  assert.ok(security.includes("grant select, insert, update on table public.%I to authenticated"));
  assert.ok(security.includes("array['personnes','effetsConfies','signatures','app_state']"));
  assert.doesNotMatch(security,/grant[^;]+to anon/i);
  assert.doesNotMatch(vault,/grant[^;]+to anon/i);
});
