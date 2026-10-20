-- DOTATIONS - Maintenance quota Supabase
-- Objectif: empecher app_state_versions et audit_log de regonfler.
-- A lancer dans Supabase SQL Editor apres sauvegarde / verification.

begin;

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

create or replace function public.prune_audit_log()
returns void
language plpgsql
security definer
set search_path = public
as $$
begin
  -- Ne jamais conserver les copies completes de public.app_state.
  delete from public.audit_log
  where table_name = 'app_state';

  delete from public.audit_log
  where created_at < now() - interval '48 hours';

  with ranked as (
    select id,
           row_number() over (order by created_at desc, id desc) as rn
    from public.audit_log
  )
  delete from public.audit_log a
  using ranked r
  where a.id = r.id
    and r.rn > 1000;
end;
$$;

select public.prune_app_state_versions();
select public.prune_audit_log();

commit;

-- Optionnel si pg_cron est disponible:
-- select cron.schedule(
--   'dotations_quota_maintenance',
--   '20 3 * * *',
--   $$select public.prune_app_state_versions(); select public.prune_audit_log();$$
-- );
