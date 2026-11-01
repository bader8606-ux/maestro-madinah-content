-- Run as database administrator only, after creating the owner's real Auth user.
-- Replace OWNER_EMAIL before running. No password or service key belongs here.
-- The transaction creates a new workspace, grants its creator the owner role,
-- and returns the generated workspace id for dist/config.js.
do $$
declare
  v_user uuid;
  v_workspace uuid;
begin
  select id into v_user from auth.users where lower(email)=lower('OWNER_EMAIL');
  if v_user is null then raise exception 'Create the Auth user first and replace OWNER_EMAIL'; end if;
  insert into maestro_private.workspaces(name) values('Madinah Content') returning id into v_workspace;
  insert into maestro_private.workspace_members(workspace_id,user_id,role) values(v_workspace,v_user,'owner');
  raise notice 'workspaceId: %', v_workspace;
end;
$$;

-- To add an existing Auth user as a viewer/editor, use the generated workspace UUID:
-- insert into maestro_private.workspace_members(workspace_id,user_id,role)
-- select 'WORKSPACE_UUID'::uuid,id,'viewer' from auth.users where lower(email)=lower('MEMBER_EMAIL');
