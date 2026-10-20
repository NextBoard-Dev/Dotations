-- DOTATIONS - Script unique de securisation (idempotent)
-- A executer dans Supabase SQL Editor

begin;

create table if not exists public.audit_log (
  id uuid primary key default gen_random_uuid(),
  table_name text not null,
  action text not null,
  row_id text,
  actor_id uuid,
  old_data jsonb,
  new_data jsonb,
  created_at timestamptz not null default now()
);

alter table public.audit_log enable row level security;

-- Acces explicite apres changement des privileges par defaut Supabase.
-- La policy audit_log_read_admin_only reste obligatoire et limite aux admins.
grant select on table public.audit_log to authenticated;

create index if not exists ix_audit_log_created_at
  on public.audit_log(created_at desc);

create or replace function public.prune_audit_log()
returns void
language plpgsql
security definer
set search_path = public
as $$
begin
  -- Retention temps: 7 jours
  delete from public.audit_log
  where created_at < now() - interval '7 days';

  -- Retention volume: max 5000 lignes recentes
  with ranked as (
    select id,
           row_number() over (order by created_at desc, id desc) as rn
    from public.audit_log
  )
  delete from public.audit_log a
  using ranked r
  where a.id = r.id
    and r.rn > 5000;
end;
$$;

drop policy if exists "audit_log_read_admin_only" on public.audit_log;
create policy "audit_log_read_admin_only"
on public.audit_log
for select
to authenticated
using (
  exists (
    select 1
    from public.profiles p
    where p.id = auth.uid() and lower(coalesce(p.role, '')) = 'admin'
  )
);

drop policy if exists "audit_log_no_write_client" on public.audit_log;
create policy "audit_log_no_write_client"
on public.audit_log
for all
to authenticated
using (false)
with check (false);

create or replace function public.current_user_role()
returns text
language sql
stable
as $$
  select lower(coalesce(p.role, 'viewer'))
  from public.profiles p
  where p.id = auth.uid()
$$;

create or replace function public.is_admin()
returns boolean
language sql
stable
as $$
  select coalesce(public.current_user_role() = 'admin', false)
$$;

create or replace function public.is_editor_or_admin()
returns boolean
language sql
stable
as $$
  select coalesce(public.current_user_role() in ('editor','admin'), false)
$$;

create or replace function public.audit_row_change()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  v_row_id text;
  v_old_data jsonb;
  v_new_data jsonb;
  v_old_payload jsonb;
  v_new_payload jsonb;
begin
  if tg_op = 'INSERT' then
    v_row_id := coalesce((to_jsonb(new)->>'id'), '');
    v_new_data := to_jsonb(new);
    if tg_table_name = 'app_state' then
      v_new_payload := coalesce(v_new_data->'payload', '{}'::jsonb);
      v_new_data := jsonb_build_object(
        'id', v_new_data->>'id',
        'revision', v_new_data->>'revision',
        'updated_at', v_new_data->>'updated_at',
        'payload_bytes', length(v_new_payload::text),
        'payload_checksum', public.jsonb_sha256(v_new_payload)
      );
    end if;
    insert into public.audit_log(table_name, action, row_id, actor_id, old_data, new_data)
    values (tg_table_name, tg_op, v_row_id, auth.uid(), null, v_new_data);
    return new;
  elsif tg_op = 'UPDATE' then
    v_row_id := coalesce((to_jsonb(new)->>'id'), (to_jsonb(old)->>'id'), '');
    v_old_data := to_jsonb(old);
    v_new_data := to_jsonb(new);
    if tg_table_name = 'app_state' then
      v_old_payload := coalesce(v_old_data->'payload', '{}'::jsonb);
      v_new_payload := coalesce(v_new_data->'payload', '{}'::jsonb);
      v_old_data := jsonb_build_object(
        'id', v_old_data->>'id',
        'revision', v_old_data->>'revision',
        'updated_at', v_old_data->>'updated_at',
        'payload_bytes', length(v_old_payload::text),
        'payload_checksum', public.jsonb_sha256(v_old_payload)
      );
      v_new_data := jsonb_build_object(
        'id', v_new_data->>'id',
        'revision', v_new_data->>'revision',
        'updated_at', v_new_data->>'updated_at',
        'payload_bytes', length(v_new_payload::text),
        'payload_checksum', public.jsonb_sha256(v_new_payload)
      );
    end if;
    insert into public.audit_log(table_name, action, row_id, actor_id, old_data, new_data)
    values (tg_table_name, tg_op, v_row_id, auth.uid(), v_old_data, v_new_data);
    return new;
  elsif tg_op = 'DELETE' then
    v_row_id := coalesce((to_jsonb(old)->>'id'), '');
    v_old_data := to_jsonb(old);
    if tg_table_name = 'app_state' then
      v_old_payload := coalesce(v_old_data->'payload', '{}'::jsonb);
      v_old_data := jsonb_build_object(
        'id', v_old_data->>'id',
        'revision', v_old_data->>'revision',
        'updated_at', v_old_data->>'updated_at',
        'payload_bytes', length(v_old_payload::text),
        'payload_checksum', public.jsonb_sha256(v_old_payload)
      );
    end if;
    insert into public.audit_log(table_name, action, row_id, actor_id, old_data, new_data)
    values (tg_table_name, tg_op, v_row_id, auth.uid(), v_old_data, null);
    return old;
  end if;
  return null;
end;
$$;

create or replace function public.protect_sensitive_columns()
returns trigger
language plpgsql
as $$
begin
  if auth.uid() is distinct from null and not public.is_admin() then
    if to_jsonb(new)->>'id' is distinct from to_jsonb(old)->>'id' then
      raise exception 'Modification de id interdite';
    end if;
    if (to_jsonb(old) ? 'created_at') and (to_jsonb(new)->>'created_at' is distinct from to_jsonb(old)->>'created_at') then
      raise exception 'Modification de created_at interdite';
    end if;
    if (to_jsonb(old) ? 'owner_id') and (to_jsonb(new)->>'owner_id' is distinct from to_jsonb(old)->>'owner_id') then
      raise exception 'Modification de owner_id interdite';
    end if;
  end if;

  if (to_jsonb(new) ? 'updated_at') then
    new.updated_at := now();
  end if;
  if (to_jsonb(new) ? 'updated_by') then
    new.updated_by := auth.uid();
  end if;
  return new;
end;
$$;

do $$
declare
  t text;
begin
  foreach t in array array['personnes','effetsConfies','signatures','app_state'] loop
    if to_regclass(format('public.%I', t)) is not null then
      execute format('alter table public.%I enable row level security', t);
      -- Droits explicites, identiques au perimetre actuel ; RLS reste applicable.
      -- Aucun droit anonyme et aucune suppression physique accordee.
      execute format('grant select, insert, update on table public.%I to authenticated', t);
    end if;
  end loop;
end$$;

do $$
begin
  if to_regclass('public.personnes') is not null then
    execute 'alter table public.personnes add column if not exists is_deleted boolean not null default false';
    execute 'alter table public.personnes add column if not exists updated_at timestamptz';
    execute 'alter table public.personnes add column if not exists updated_by uuid';
  end if;
  if to_regclass('public."effetsConfies"') is not null then
    execute 'alter table public."effetsConfies" add column if not exists is_deleted boolean not null default false';
    execute 'alter table public."effetsConfies" add column if not exists updated_at timestamptz';
    execute 'alter table public."effetsConfies" add column if not exists updated_by uuid';
  end if;
  if to_regclass('public.signatures') is not null then
    execute 'alter table public.signatures add column if not exists is_deleted boolean not null default false';
    execute 'alter table public.signatures add column if not exists updated_at timestamptz';
    execute 'alter table public.signatures add column if not exists updated_by uuid';
  end if;
  if to_regclass('public.app_state') is not null then
    execute 'alter table public.app_state add column if not exists updated_at timestamptz';
    execute 'alter table public.app_state add column if not exists updated_by uuid';
  end if;
end$$;

do $$
declare
  t text;
begin
  foreach t in array array['personnes','effetsConfies','signatures','app_state'] loop
    if to_regclass(format('public.%I', t)) is not null then
      execute format('drop policy if exists "read_%s" on public.%I', t, t);
      execute format('drop policy if exists "insert_%s" on public.%I', t, t);
      execute format('drop policy if exists "update_%s" on public.%I', t, t);
      execute format('drop policy if exists "delete_%s_admin_only" on public.%I', t, t);

      execute format(
        'create policy "read_%s" on public.%I for select to authenticated using (coalesce((to_jsonb(%I)->>''is_deleted'')::boolean, false) = false)',
        t, t, t
      );
      execute format(
        'create policy "insert_%s" on public.%I for insert to authenticated with check (public.is_editor_or_admin())',
        t, t
      );
      execute format(
        'create policy "update_%s" on public.%I for update to authenticated using (public.is_editor_or_admin()) with check (public.is_editor_or_admin())',
        t, t
      );
      execute format(
        'create policy "delete_%s_admin_only" on public.%I for delete to authenticated using (public.is_admin())',
        t, t
      );
    end if;
  end loop;
end$$;

do $$
begin
  if to_regclass('public.personnes') is not null then
    execute 'drop trigger if exists trg_protect_sensitive_personnes on public.personnes';
    execute 'create trigger trg_protect_sensitive_personnes before update on public.personnes for each row execute function public.protect_sensitive_columns()';
    execute 'drop trigger if exists trg_audit_personnes on public.personnes';
    execute 'create trigger trg_audit_personnes after insert or update or delete on public.personnes for each row execute function public.audit_row_change()';
  end if;

  if to_regclass('public."effetsConfies"') is not null then
    execute 'drop trigger if exists trg_protect_sensitive_effets on public."effetsConfies"';
    execute 'create trigger trg_protect_sensitive_effets before update on public."effetsConfies" for each row execute function public.protect_sensitive_columns()';
    execute 'drop trigger if exists trg_audit_effets on public."effetsConfies"';
    execute 'create trigger trg_audit_effets after insert or update or delete on public."effetsConfies" for each row execute function public.audit_row_change()';
  end if;

  if to_regclass('public.signatures') is not null then
    execute 'drop trigger if exists trg_protect_sensitive_signatures on public.signatures';
    execute 'create trigger trg_protect_sensitive_signatures before update on public.signatures for each row execute function public.protect_sensitive_columns()';
    execute 'drop trigger if exists trg_audit_signatures on public.signatures';
    execute 'create trigger trg_audit_signatures after insert or update or delete on public.signatures for each row execute function public.audit_row_change()';
  end if;

  if to_regclass('public.app_state') is not null then
    execute 'drop trigger if exists trg_protect_sensitive_app_state on public.app_state';
    execute 'create trigger trg_protect_sensitive_app_state before update on public.app_state for each row execute function public.protect_sensitive_columns()';
    execute 'drop trigger if exists trg_audit_app_state on public.app_state';
    execute 'create trigger trg_audit_app_state after insert or update or delete on public.app_state for each row execute function public.audit_row_change()';
  end if;
end$$;

commit;

-- Verifications rapides (lecture)
select schemaname, tablename, rowsecurity
from pg_tables
where schemaname = 'public'
  and tablename in ('personnes','effetsConfies','signatures','app_state','audit_log')
order by tablename;

select schemaname, tablename, policyname, permissive, roles, cmd
from pg_policies
where schemaname = 'public'
  and tablename in ('personnes','effetsConfies','signatures','app_state','audit_log')
order by tablename, policyname;

