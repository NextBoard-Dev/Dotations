-- Signature mobile : droits explicites et acces par jeton valide uniquement.
-- Transaction atomique, aucune suppression de table ou de donnees.
-- Prerequis : public.app_state(id,payload), public.profiles(id,role).
begin;
create table if not exists public.signatures (id uuid primary key default gen_random_uuid());
alter table public.signatures add column if not exists "token" text;
alter table public.signatures add column if not exists "person_id" text;
alter table public.signatures add column if not exists "doc_type" text;
alter table public.signatures add column if not exists "signer" text;
alter table public.signatures add column if not exists "status" text;
alter table public.signatures add column if not exists "signature_data" text;
alter table public.signatures add column if not exists "storage_ref" text;
alter table public.signatures add column if not exists "storage_public_url" text;
alter table public.signatures add column if not exists "validated_at_text" text;
alter table public.signatures add column if not exists "signed_at" timestamptz default now();
alter table public.signatures add column if not exists "person_nom" text;
alter table public.signatures add column if not exists "person_prenom" text;
alter table public.signatures add column if not exists "signer_name" text;
alter table public.signatures add column if not exists "signer_function" text;
alter table public.signatures add column if not exists "updated_at" timestamptz default now();
alter table public.signatures add column if not exists "created_at" timestamptz default now();
alter table public.signatures add column if not exists "is_deleted" boolean default false;
alter table public.signatures add column if not exists "signatureData" text;
alter table public.signatures add column if not exists "signedAt" timestamptz;
alter table public.signatures add column if not exists "signataireName" text;
alter table public.signatures add column if not exists "signataireFunction" text;
-- Les doublons bloquent l'installation au lieu de supprimer des signatures.
create unique index if not exists ux_signatures_mobile_token_complete on public.signatures(token);
alter table public.signatures enable row level security;
revoke all on table public.signatures from PUBLIC, anon;
grant select, insert, update on table public.signatures to authenticated;
drop policy if exists "mobile_signature_select_by_token" on public.signatures;
drop policy if exists "mobile_signature_insert_by_token" on public.signatures;
drop policy if exists "mobile_signature_update_by_token" on public.signatures;
-- Garde restrictive : aucune policy historique ne peut ouvrir les ecritures.
drop policy if exists signatures_registered_read on public.signatures;
create policy signatures_registered_read on public.signatures as restrictive for select to authenticated
using (exists(select 1 from public.profiles p where p.id=auth.uid()));
drop policy if exists signatures_registered_insert on public.signatures;
create policy signatures_registered_insert on public.signatures as restrictive for insert to authenticated
with check (exists(select 1 from public.profiles p where p.id=auth.uid() and p.role in ('admin','editor')));
drop policy if exists signatures_registered_update on public.signatures;
create policy signatures_registered_update on public.signatures as restrictive for update to authenticated
using (exists(select 1 from public.profiles p where p.id=auth.uid() and p.role in ('admin','editor')))
with check (exists(select 1 from public.profiles p where p.id=auth.uid() and p.role in ('admin','editor')));
-- Pas de CASCADE : une dependance inattendue fait annuler la transaction.
drop function if exists public.submit_mobile_signature(text,text,text,text,text,text,text,text,text,text);
drop function if exists public.fetch_mobile_signature_rows(text[]);
drop function if exists public.fetch_mobile_signature_rows_for_document(text,text);
CREATE OR REPLACE FUNCTION public.fetch_mobile_signature_rows(p_tokens text[])
 RETURNS jsonb
 LANGUAGE plpgsql
 SECURITY DEFINER
 SET search_path TO 'public', 'pg_temp'
AS $function$
declare
  v_tokens text[];
  v_rows jsonb;
begin
  select array_agg(distinct trim(token))
  into v_tokens
  from unnest(coalesce(p_tokens, array[]::text[])) as token
  where trim(token) <> '';

  if v_tokens is null or array_length(v_tokens, 1) is null then
    return '[]'::jsonb;
  end if;

  select coalesce(jsonb_agg(to_jsonb(s) order by coalesce(s.signed_at, s.updated_at, s.created_at) desc), '[]'::jsonb)
  into v_rows
  from public.signatures s
  where s.token = any(v_tokens)
    and coalesce(s.is_deleted, false) = false
    and exists (
      select 1
      from public.app_state app
      cross join lateral jsonb_array_elements(coalesce(app.payload->'demandesSignatureMobile', '[]'::jsonb)) req
      where app.id = 'main'
        and req->>'token' = s.token
    );

  return coalesce(v_rows, '[]'::jsonb);
end;
$function$;

CREATE OR REPLACE FUNCTION public.fetch_mobile_signature_rows_for_document(p_person_id text, p_doc_type text)
 RETURNS jsonb
 LANGUAGE plpgsql
 SECURITY DEFINER
 SET search_path TO 'public', 'pg_temp'
AS $function$
declare
  v_person_id text := trim(coalesce(p_person_id, ''));
  v_doc_type text := lower(trim(coalesce(p_doc_type, '')));
  v_rows jsonb;
begin
  if auth.uid() is null or not exists(select 1 from public.profiles p where p.id=auth.uid()) then
    raise exception 'SIGNATURE_READ_FORBIDDEN' using errcode='42501';
  end if;
  if v_person_id = '' or v_doc_type not in ('arrival', 'exit') then
    return '[]'::jsonb;
  end if;

  select coalesce(jsonb_agg(to_jsonb(s) order by coalesce(s.signed_at, s.updated_at, s.created_at) desc), '[]'::jsonb)
  into v_rows
  from public.signatures s
  where s.person_id = v_person_id
    and s.doc_type = v_doc_type
    and coalesce(s.is_deleted, false) = false
    and exists (
      select 1
      from public.app_state app
      cross join lateral jsonb_array_elements(coalesce(app.payload->'demandesSignatureMobile', '[]'::jsonb)) req
      where app.id = 'main'
        and req->>'token' = s.token
        and req->>'personId' = s.person_id
        and lower(coalesce(req->>'docType', '')) = s.doc_type
        and lower(coalesce(req->>'signer', '')) = lower(coalesce(s.signer, ''))
    );

  return coalesce(v_rows, '[]'::jsonb);
end;
$function$;

CREATE OR REPLACE FUNCTION public.submit_mobile_signature(p_token text, p_person_id text, p_doc_type text, p_signer text, p_signature_data text, p_validated_at_text text DEFAULT ''::text, p_storage_ref text DEFAULT ''::text, p_storage_public_url text DEFAULT ''::text, p_person_nom text DEFAULT ''::text, p_person_prenom text DEFAULT ''::text, p_signer_name text DEFAULT ''::text, p_signer_function text DEFAULT ''::text)
 RETURNS jsonb
 LANGUAGE plpgsql
 SECURITY DEFINER
 SET search_path TO 'public', 'pg_temp'
AS $function$
declare
  v_doc_type text := lower(trim(coalesce(p_doc_type, '')));
  v_signer text := lower(trim(coalesce(p_signer, '')));
  v_now timestamptz := now();
  v_is_valid boolean := false;
begin
  if coalesce(trim(p_token), '') = ''
     or coalesce(trim(p_person_id), '') = ''
     or v_doc_type not in ('arrival', 'exit')
     or v_signer not in ('personnel', 'representant')
     or coalesce(trim(p_signature_data), '') = '' then
    raise exception 'SIGNATURE_PAYLOAD_INVALID';
  end if;

  select exists (
    select 1
    from public.app_state s
    cross join lateral jsonb_array_elements(coalesce(s.payload->'demandesSignatureMobile', '[]'::jsonb)) req
    where s.id = 'main'
      and req->>'token' = p_token
      and req->>'personId' = p_person_id
      and lower(coalesce(req->>'docType', '')) = v_doc_type
      and lower(coalesce(req->>'signer', '')) = v_signer
  ) into v_is_valid;

  if not v_is_valid then
    raise exception 'SIGNATURE_TOKEN_INVALID';
  end if;

  insert into public.signatures (
    token,
    person_id,
    doc_type,
    signer,
    status,
    signature_data,
    "signatureData",
    storage_ref,
    storage_public_url,
    validated_at_text,
    signed_at,
    "signedAt",
    person_nom,
    person_prenom,
    signer_name,
    signer_function,
    "signataireName",
    "signataireFunction",
    updated_at,
    is_deleted
  ) values (
    p_token,
    p_person_id,
    v_doc_type,
    v_signer,
    'SIGNEE',
    p_signature_data,
    p_signature_data,
    coalesce(p_storage_ref, ''),
    coalesce(p_storage_public_url, ''),
    coalesce(p_validated_at_text, ''),
    v_now,
    v_now,
    coalesce(p_person_nom, ''),
    coalesce(p_person_prenom, ''),
    coalesce(p_signer_name, ''),
    coalesce(p_signer_function, ''),
    coalesce(p_signer_name, ''),
    coalesce(p_signer_function, ''),
    v_now,
    false
  )
  on conflict (token) do update set
    person_id = excluded.person_id,
    doc_type = excluded.doc_type,
    signer = excluded.signer,
    status = excluded.status,
    signature_data = excluded.signature_data,
    "signatureData" = excluded."signatureData",
    storage_ref = excluded.storage_ref,
    storage_public_url = excluded.storage_public_url,
    validated_at_text = excluded.validated_at_text,
    signed_at = excluded.signed_at,
    "signedAt" = excluded."signedAt",
    person_nom = excluded.person_nom,
    person_prenom = excluded.person_prenom,
    signer_name = excluded.signer_name,
    signer_function = excluded.signer_function,
    "signataireName" = excluded."signataireName",
    "signataireFunction" = excluded."signataireFunction",
    updated_at = excluded.updated_at,
    is_deleted = false;

  return jsonb_build_object('ok', true, 'token', p_token, 'personId', p_person_id, 'docType', v_doc_type, 'signer', v_signer);
end;
$function$;

revoke all on function public.submit_mobile_signature(text,text,text,text,text,text,text,text,text,text,text,text) from PUBLIC;
revoke all on function public.fetch_mobile_signature_rows(text[]) from PUBLIC;
revoke all on function public.fetch_mobile_signature_rows_for_document(text,text) from PUBLIC, anon;
grant execute on function public.submit_mobile_signature(text,text,text,text,text,text,text,text,text,text,text,text) to anon, authenticated;
grant execute on function public.fetch_mobile_signature_rows(text[]) to anon, authenticated;
grant execute on function public.fetch_mobile_signature_rows_for_document(text,text) to authenticated;
commit;
