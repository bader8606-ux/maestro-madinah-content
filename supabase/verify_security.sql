-- Read-only post-deployment check. Run in the Supabase SQL Editor as administrator.
-- Returns metadata booleans only: no workbook payloads, member IDs, emails or keys.
-- Every column should be true. This does not test live JWT verification or API
-- routing. Also confirm in Data API settings that public is exposed and
-- maestro_private is NOT exposed; SQL Editor settings may omit API configuration.
with expected_rpc(signature) as (
  values ('public.maestro_get_state(uuid)'),
         ('public.maestro_commit(uuid,bigint,jsonb,text,jsonb)'),
         ('public.maestro_undo(uuid,bigint)')
), rpc as (
  select e.signature,p.*
  from expected_rpc e
  left join pg_proc p on p.oid=to_regprocedure(e.signature)
), private_tables as (
  select c.* from pg_class c join pg_namespace n on n.oid=c.relnamespace
  where n.nspname='maestro_private' and c.relkind='r'
    and c.relname in ('workspaces','workspace_members','workspace_changes')
), api_roles(role_name) as (values ('anon'),('authenticated'))
select
  (select count(oid)=3 from rpc)
    and (select count(*)=3 from pg_proc p join pg_namespace n on n.oid=p.pronamespace
         where n.nspname='public' and p.proname like 'maestro\_%' escape '\')
    as exact_three_rpc_signatures,
  (select bool_and(coalesce(has_function_privilege('authenticated',oid,'EXECUTE'),false))
     from rpc)
    and has_schema_privilege('authenticated','public','USAGE')
    as authenticated_rpc_access,
  (select bool_and(not coalesce(has_function_privilege('anon',oid,'EXECUTE'),true))
     from rpc) as anonymous_rpc_denied,
  (select bool_and(coalesce(prosecdef,false)
       and coalesce('search_path=""'=any(proconfig),false)) from rpc)
    as definer_search_paths_pinned,
  not has_schema_privilege('anon','maestro_private','USAGE,CREATE')
    and not has_schema_privilege('authenticated','maestro_private','USAGE,CREATE')
    as private_schema_denied,
  (select count(*)=3 and bool_and(relrowsecurity) from private_tables)
    as three_private_tables_rls,
  (select bool_and(not has_table_privilege(r.role_name,t.oid,
      'SELECT,INSERT,UPDATE,DELETE,TRUNCATE,REFERENCES,TRIGGER'))
     from private_tables t cross join api_roles r)
    as private_table_privileges_denied,
  not exists (select 1 from pg_policy p
    join private_tables t on t.oid=p.polrelid) as no_direct_table_policies,
  not exists (select 1 from pg_proc p join pg_namespace n on n.oid=p.pronamespace
    cross join api_roles r
    where n.nspname='maestro_private'
      and has_function_privilege(r.role_name,p.oid,'EXECUTE'))
    as private_helper_execution_denied,
  exists (select 1 from pg_trigger t
    where t.tgrelid=to_regclass('maestro_private.workspace_changes')
      and t.tgname='workspace_changes_immutable'
      and not t.tgisinternal and t.tgenabled in ('O','A')
      and (t.tgtype::int & 27)=27)
    as immutable_audit_trigger_enabled,
  to_regprocedure('auth.uid()') is not null as auth_uid_available;

-- SQL grants establish reachability; they do not prove sign-in works. Final API
-- checks: unauthenticated RPC is denied; authenticated nonmember is denied;
-- allowlisted viewer can read and cannot commit; owner/editor can save and undo.
-- auth.uid() is usually NULL in SQL Editor because no user JWT is attached.
