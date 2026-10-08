-- DOTATIONS - Verrou total "zero delete physique"
-- Objectif:
-- 1) interdire tout DELETE physique sur tables metier
-- 2) convertir DELETE en soft delete (is_deleted=true) quand possible
-- 3) renforcer RLS pour ne jamais autoriser de suppression physique

begin;

-- 1) Colonnes soft-delete minimales
alter table if exists public.personnes
  add column if not exists is_deleted boolean not null default false,
  add column if not exists deleted_at timestamptz,
  add column if not exists deleted_by uuid;

alter table if exists public."effetsConfies"
  add column if not exists is_deleted boolean not null default false,
  add column if not exists deleted_at timestamptz,
  add column if not exists deleted_by uuid;

alter table if exists public.signatures
  add column if not exists is_deleted boolean not null default false,
  add column if not exists deleted_at timestamptz,
  add column if not exists deleted_by uuid;

-- 2) Trigger: transformer DELETE en UPDATE soft-delete
create or replace function public.soft_delete_instead_of_delete()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  execute format(
    'update %I.%I set is_deleted = true, deleted_at = now(), deleted_by = auth.uid() where id = $1 and coalesce(is_deleted,false) = false',
    tg_table_schema, tg_table_name
  ) using old.id;
  return null; -- annule le DELETE physique
end;
$$;

drop trigger if exists trg_soft_delete_personnes on public.personnes;
create trigger trg_soft_delete_personnes
before delete on public.personnes
for each row
execute function public.soft_delete_instead_of_delete();

drop trigger if exists trg_soft_delete_effets on public."effetsConfies";
create trigger trg_soft_delete_effets
before delete on public."effetsConfies"
for each row
execute function public.soft_delete_instead_of_delete();

drop trigger if exists trg_soft_delete_signatures on public.signatures;
create trigger trg_soft_delete_signatures
before delete on public.signatures
for each row
execute function public.soft_delete_instead_of_delete();

-- 3) Hard block DELETE sur app_state + audit_log
create or replace function public.block_physical_delete()
returns trigger
language plpgsql
as $$
begin
  raise exception 'DELETE physique interdit sur %.%', tg_table_schema, tg_table_name;
end;
$$;

drop trigger if exists trg_block_delete_app_state on public.app_state;
create trigger trg_block_delete_app_state
before delete on public.app_state
for each row
execute function public.block_physical_delete();

drop trigger if exists trg_block_delete_audit_log on public.audit_log;
create trigger trg_block_delete_audit_log
before delete on public.audit_log
for each row
execute function public.block_physical_delete();

-- 4) RLS: aucune policy DELETE active
drop policy if exists "delete_personnes_admin_only" on public.personnes;
drop policy if exists "delete_effetsConfies_admin_only" on public."effetsConfies";
drop policy if exists "delete_signatures_admin_only" on public.signatures;
drop policy if exists "delete_app_state_admin_only" on public.app_state;
drop policy if exists "delete_audit_log_admin_only" on public.audit_log;

-- 5) Lecture active uniquement (soft-delete masqué)
drop policy if exists "read_personnes" on public.personnes;
create policy "read_personnes"
on public.personnes
for select
to authenticated
using (coalesce(is_deleted, false) = false);

drop policy if exists "read_effetsConfies" on public."effetsConfies";
create policy "read_effetsConfies"
on public."effetsConfies"
for select
to authenticated
using (coalesce(is_deleted, false) = false);

drop policy if exists "read_signatures" on public.signatures;
create policy "read_signatures"
on public.signatures
for select
to authenticated
using (coalesce(is_deleted, false) = false);

commit;

-- Verif rapide
select schemaname, tablename, rowsecurity
from pg_tables
where schemaname = 'public'
  and tablename in ('personnes','effetsConfies','signatures','app_state','audit_log')
order by tablename;

select schemaname, tablename, policyname, cmd
from pg_policies
where schemaname = 'public'
  and tablename in ('personnes','effetsConfies','signatures','app_state','audit_log')
order by tablename, policyname;
