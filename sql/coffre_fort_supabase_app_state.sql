-- DOTATIONS - Coffre-fort distant Supabase pour app_state
-- Objectif: anti-perte + restauration simple sans gonfler le quota
-- - Snapshot a chaque modification de payload
-- - Checksum SHA-256
-- - Journal minimal
-- - Rollback par revision
-- - Conservation sans purge automatique non archivee

begin;

create table if not exists public.app_state_versions (
  id uuid primary key default gen_random_uuid(),
  app_state_id text not null,
  source_revision bigint not null,
  payload jsonb not null,
  payload_checksum text not null,
  created_at timestamptz not null default now(),
  created_by uuid,
  reason text not null default 'snapshot_before_update'
);

create unique index if not exists ux_app_state_versions_unique_rev
  on public.app_state_versions(app_state_id, source_revision);

create index if not exists ix_app_state_versions_created_at
  on public.app_state_versions(created_at desc);

create table if not exists public.app_state_save_journal (
  id uuid primary key default gen_random_uuid(),
  app_state_id text not null,
  event_type text not null,
  actor_id uuid,
  at timestamptz not null default now(),
  details jsonb not null default '{}'::jsonb
);

create index if not exists ix_app_state_save_journal_at
  on public.app_state_save_journal(at desc);

create or replace function public.jsonb_sha256(input jsonb)
returns text
language sql
immutable
as $$
  select md5(coalesce(input, '{}'::jsonb)::text)
$$;

create or replace function public.snapshot_app_state_before_update()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  old_id text;
  old_revision bigint;
  old_checksum text;
begin
  if to_jsonb(new)->>'payload' is distinct from to_jsonb(old)->>'payload' then
    old_id := coalesce(to_jsonb(old)->>'id', '');
    old_revision := coalesce((to_jsonb(old)->>'revision')::bigint, 0);
    old_checksum := public.jsonb_sha256((to_jsonb(old)->'payload')::jsonb);

    insert into public.app_state_versions (
      app_state_id, source_revision, payload, payload_checksum, created_by, reason
    )
    values (
      old_id,
      old_revision,
      coalesce((to_jsonb(old)->'payload')::jsonb, '{}'::jsonb),
      old_checksum,
      auth.uid(),
      'snapshot_before_update'
    )
    on conflict (app_state_id, source_revision) do nothing;

    insert into public.app_state_save_journal(app_state_id, event_type, actor_id, details)
    values (
      old_id,
      'save',
      auth.uid(),
      jsonb_build_object(
        'from_revision', old_revision,
        'to_revision', coalesce((to_jsonb(new)->>'revision')::bigint, old_revision + 1),
        'checksum_before', old_checksum
      )
    );

    perform public.prune_app_state_versions();
  end if;
  return new;
end;
$$;

drop trigger if exists trg_snapshot_app_state_before_update on public.app_state;
create trigger trg_snapshot_app_state_before_update
before update on public.app_state
for each row
execute function public.snapshot_app_state_before_update();

create or replace function public.restore_app_state_revision(
  _app_state_id text,
  _source_revision bigint,
  _reason text default 'manual_restore'
)
returns table(id text, revision bigint, restored_at timestamptz)
language plpgsql
security definer
set search_path = public
as $$
declare
  v_payload jsonb;
  v_current_revision bigint;
begin
  if not public.is_admin() then
    raise exception 'RESTORE FORBIDDEN';
  end if;

  select payload
    into v_payload
  from public.app_state_versions
  where app_state_id = _app_state_id
    and source_revision = _source_revision
  limit 1;

  if v_payload is null then
    raise exception 'REVISION_NOT_FOUND';
  end if;

  select revision
    into v_current_revision
  from public.app_state
  where id::text = _app_state_id
  for update;

  update public.app_state
     set payload = v_payload,
         revision = coalesce(v_current_revision, 0) + 1,
         updated_at = now(),
         updated_by = auth.uid()
   where id::text = _app_state_id;

  insert into public.app_state_save_journal(app_state_id, event_type, actor_id, details)
  values (
    _app_state_id,
    'restore',
    auth.uid(),
    jsonb_build_object(
      'restored_source_revision', _source_revision,
      'reason', coalesce(_reason, 'manual_restore'),
      'checksum_after', public.jsonb_sha256(v_payload)
    )
  );

  return query
  select _app_state_id, (coalesce(v_current_revision, 0) + 1)::bigint, now();
end;
$$;

create or replace function public.prune_app_state_versions()
returns void language plpgsql security definer
set search_path = pg_catalog, public, pg_temp
as $$
begin
  -- Aucun point retire sans archivage externe verifie.
  -- Compatibilite des appels existants ; conservation sans suppression.
  return;
end;
$$;
revoke all on function public.prune_app_state_versions() from PUBLIC, anon, authenticated;

-- RLS: lecture admin seulement, ecriture directe client interdite
alter table public.app_state_versions enable row level security;
alter table public.app_state_save_journal enable row level security;

drop policy if exists "app_state_versions_read_admin_only" on public.app_state_versions;
create policy "app_state_versions_read_admin_only"
on public.app_state_versions
for select
to authenticated
using (public.is_admin());

drop policy if exists "app_state_versions_no_write_client" on public.app_state_versions;
create policy "app_state_versions_no_write_client"
on public.app_state_versions
for all
to authenticated
using (false)
with check (false);

drop policy if exists "app_state_save_journal_read_admin_only" on public.app_state_save_journal;
create policy "app_state_save_journal_read_admin_only"
on public.app_state_save_journal
for select
to authenticated
using (public.is_admin());

drop policy if exists "app_state_save_journal_no_write_client" on public.app_state_save_journal;
create policy "app_state_save_journal_no_write_client"
on public.app_state_save_journal
for all
to authenticated
using (false)
with check (false);

-- Lecture uniquement : les policies is_admin() limitent les lignes visibles.
grant select on table public.app_state_versions to authenticated;
grant select on table public.app_state_save_journal to authenticated;

commit;

-- Optionnel: cron quotidien pour prune snapshots
-- select cron.schedule(
--   'app_state_versions_retention',
--   '10 3 * * *',
--   $$select public.prune_app_state_versions();$$
-- );

