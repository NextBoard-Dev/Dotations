-- DOTATIONS - Correctif immediat signature mobile / Supabase
-- A executer seul dans Supabase SQL Editor.
-- Objectif:
-- 1. supprimer la dependance a digest()/pgcrypto;
-- 2. permettre aux pages mobiles d'ecrire une signature dediee sans reecrire tout app_state;
-- 3. permettre a l'UI de relire uniquement les signatures mobiles necessaires;
-- 4. conserver les anciennes colonnes existantes sans destruction.

create or replace function public.jsonb_sha256(input jsonb)
returns text
language sql
immutable
as $$
  select md5(coalesce(input, '{}'::jsonb)::text)
$$;

create table if not exists public.signatures (
  id uuid primary key default gen_random_uuid()
);

alter table public.signatures
  add column if not exists token text,
  add column if not exists person_id text,
  add column if not exists doc_type text,
  add column if not exists signer text,
  add column if not exists status text,
  add column if not exists signature_data text,
  add column if not exists storage_ref text,
  add column if not exists storage_public_url text,
  add column if not exists validated_at_text text,
  add column if not exists signed_at timestamptz default now(),
  add column if not exists person_nom text,
  add column if not exists person_prenom text,
  add column if not exists signer_name text,
  add column if not exists signer_function text,
  add column if not exists updated_at timestamptz default now();

create unique index if not exists ux_signatures_mobile_token
  on public.signatures(token)
  where token is not null;

create index if not exists ix_signatures_mobile_lookup
  on public.signatures(person_id, doc_type, signer, updated_at desc)
  where person_id is not null and doc_type is not null and signer is not null;

alter table public.signatures enable row level security;

drop policy if exists "mobile_signature_select_by_token" on public.signatures;
create policy "mobile_signature_select_by_token"
on public.signatures
for select
to anon, authenticated
using (token is not null and person_id is not null and doc_type in ('arrival', 'exit'));

drop policy if exists "mobile_signature_insert_by_token" on public.signatures;
create policy "mobile_signature_insert_by_token"
on public.signatures
for insert
to anon, authenticated
with check (
  token is not null
  and person_id is not null
  and doc_type in ('arrival', 'exit')
  and signer in ('personnel', 'representant')
);

drop policy if exists "mobile_signature_update_by_token" on public.signatures;
create policy "mobile_signature_update_by_token"
on public.signatures
for update
to anon, authenticated
using (token is not null)
with check (
  token is not null
  and person_id is not null
  and doc_type in ('arrival', 'exit')
  and signer in ('personnel', 'representant')
);

create or replace function public.submit_mobile_signature(
  p_token text,
  p_person_id text,
  p_doc_type text,
  p_signer text,
  p_signature_data text,
  p_validated_at_text text default null,
  p_person_nom text default null,
  p_person_prenom text default null,
  p_signer_name text default null,
  p_signer_function text default null
)
returns public.signatures
language plpgsql
as $$
declare
  saved public.signatures;
begin
  insert into public.signatures (
    token,
    person_id,
    doc_type,
    signer,
    status,
    signature_data,
    validated_at_text,
    signed_at,
    person_nom,
    person_prenom,
    signer_name,
    signer_function,
    updated_at
  )
  values (
    p_token,
    p_person_id,
    lower(p_doc_type),
    lower(p_signer),
    'SIGNEE',
    p_signature_data,
    coalesce(p_validated_at_text, now()::text),
    now(),
    p_person_nom,
    p_person_prenom,
    p_signer_name,
    p_signer_function,
    now()
  )
  on conflict (token) where token is not null
  do update set
    person_id = excluded.person_id,
    doc_type = excluded.doc_type,
    signer = excluded.signer,
    status = excluded.status,
    signature_data = excluded.signature_data,
    validated_at_text = excluded.validated_at_text,
    signed_at = excluded.signed_at,
    person_nom = excluded.person_nom,
    person_prenom = excluded.person_prenom,
    signer_name = excluded.signer_name,
    signer_function = excluded.signer_function,
    updated_at = excluded.updated_at
  returning * into saved;

  return saved;
end;
$$;

create or replace function public.fetch_mobile_signature_rows(p_tokens text[])
returns setof public.signatures
language sql
stable
as $$
  select *
  from public.signatures
  where token = any(coalesce(p_tokens, array[]::text[]))
  order by updated_at desc, signed_at desc
$$;

create or replace function public.fetch_mobile_signature_rows_for_document(
  p_person_id text,
  p_doc_type text
)
returns setof public.signatures
language sql
stable
as $$
  select *
  from public.signatures
  where person_id = p_person_id
    and doc_type = lower(p_doc_type)
    and signer in ('personnel', 'representant')
    and signature_data is not null
  order by updated_at desc, signed_at desc
$$;

grant execute on function public.submit_mobile_signature(text, text, text, text, text, text, text, text, text, text) to anon, authenticated;
grant execute on function public.fetch_mobile_signature_rows(text[]) to anon, authenticated;
grant execute on function public.fetch_mobile_signature_rows_for_document(text, text) to anon, authenticated;

-- Test rapide : doit retourner une valeur md5, sans erreur digest().
select public.jsonb_sha256('{"test":"ok"}'::jsonb) as checksum_test;
