-- Apply as the Supabase database administrator. No workbook data or users are seeded.
-- The private schema must NOT be added to the project's exposed API schemas.
begin;

create schema if not exists maestro_private;
revoke all on schema maestro_private from public, anon, authenticated;

create table if not exists maestro_private.workspaces (
  id uuid primary key default gen_random_uuid(),
  name text not null check (length(name) between 1 and 200),
  revision bigint not null default 0 check (revision >= 0),
  payload jsonb,
  saved_at timestamptz not null default now(),
  created_at timestamptz not null default now(),
  check (payload is null or jsonb_typeof(payload) = 'object')
);
create table if not exists maestro_private.workspace_members (
  workspace_id uuid not null references maestro_private.workspaces(id) on delete cascade,
  user_id uuid not null references auth.users(id) on delete cascade,
  role text not null check (role in ('owner','editor','viewer')),
  created_at timestamptz not null default now(),
  primary key (workspace_id, user_id)
);
create table if not exists maestro_private.workspace_changes (
  id bigint generated always as identity primary key,
  workspace_id uuid not null references maestro_private.workspaces(id),
  revision bigint not null check (revision > 0),
  actor_id uuid not null,
  action text not null check (action in ('add','edit','delete','import','restore','save','undo')),
  detail jsonb not null default '{}'::jsonb,
  at timestamptz not null default now(),
  before_payload jsonb,
  after_payload jsonb,
  undo_target bigint references maestro_private.workspace_changes(id),
  unique (workspace_id, revision)
);
create index if not exists workspace_changes_recent
  on maestro_private.workspace_changes (workspace_id, revision desc);
alter table maestro_private.workspaces enable row level security;
alter table maestro_private.workspace_members enable row level security;
alter table maestro_private.workspace_changes enable row level security;
-- Deliberately no direct-access policies or grants: RLS defaults to deny.
-- Only guarded SECURITY DEFINER RPCs expose workspace data to authenticated members.
revoke all on all tables in schema maestro_private from public, anon, authenticated;
revoke all on all sequences in schema maestro_private from public, anon, authenticated;

create or replace function maestro_private.member_role(p_workspace_id uuid)
returns text language plpgsql stable security definer set search_path = '' as $$
declare v_role text;
begin
  if auth.uid() is null then
    raise exception using errcode='42501', message='not_authorized: sign in is required';
  end if;
  select m.role into v_role from maestro_private.workspace_members m
    where m.workspace_id=p_workspace_id and m.user_id=auth.uid();
  if v_role is null then
    raise exception using errcode='42501', message='not_authorized: workspace membership required';
  end if;
  return v_role;
end;
$$;

create or replace function maestro_private.envelope(p_workspace_id uuid, p_role text)
returns jsonb language sql volatile security definer set search_path = '' as $$
  select jsonb_build_object(
    'revision', w.revision, 'payload', w.payload, 'savedAt', w.saved_at,
    'role', p_role,
    'history', coalesce((select jsonb_agg(jsonb_build_object(
      'id', h.id, 'action', h.action, 'detail', h.detail, 'at', h.at)
      order by h.revision desc)
      from (select c.id,c.action,c.detail,c.at,c.revision
        from maestro_private.workspace_changes c where c.workspace_id=w.id
        order by c.revision desc limit 200) h), '[]'::jsonb),
    'canUndo', p_role in ('owner','editor') and exists(
      select 1 from maestro_private.workspace_changes c
      where c.workspace_id=w.id and c.revision=w.revision and c.action<>'undo')
  ) from maestro_private.workspaces w where w.id=p_workspace_id;
$$;

create or replace function maestro_private.prevent_audit_mutation()
returns trigger language plpgsql set search_path = '' as $$
begin
  raise exception using errcode='42501', message='Audit history is append-only';
end;
$$;
drop trigger if exists workspace_changes_immutable on maestro_private.workspace_changes;
create trigger workspace_changes_immutable before update or delete
  on maestro_private.workspace_changes for each row
  execute function maestro_private.prevent_audit_mutation();

create or replace function public.maestro_get_state(p_workspace_id uuid)
returns jsonb language plpgsql stable security definer set search_path = '' as $$
declare v_role text;
begin
  v_role := maestro_private.member_role(p_workspace_id);
  return maestro_private.envelope(p_workspace_id, v_role);
end;
$$;

create or replace function public.maestro_commit(
  p_workspace_id uuid, p_expected_revision bigint, p_payload jsonb,
  p_action text, p_detail jsonb
)
returns jsonb language plpgsql security definer set search_path = '' as $$
declare
  v_role text;
  v_workspace maestro_private.workspaces%rowtype;
  v_source bytea;
  v_detail jsonb := coalesce(p_detail,'{}'::jsonb);
begin
  v_role := maestro_private.member_role(p_workspace_id);
  if v_role not in ('owner','editor') then
    raise exception using errcode='42501', message='not_authorized: editor permission required';
  end if;
  select * into v_workspace from maestro_private.workspaces
    where id=p_workspace_id for update;
  if p_expected_revision is distinct from v_workspace.revision then
    raise exception using errcode='40001', message='revision_conflict: a newer version was saved';
  end if;
  if p_action is null or p_action not in ('add','edit','delete','import','restore','save') then
    raise exception using errcode='22023', message='Unsupported commit action';
  end if;
  if jsonb_typeof(v_detail)<>'object' or octet_length(v_detail::text)>8192 then
    raise exception using errcode='22023', message='Invalid audit detail';
  end if;
  if p_payload is null or jsonb_typeof(p_payload)<>'object'
    or octet_length(p_payload::text)>26214400
    or jsonb_typeof(p_payload->'sourceName') is distinct from 'string'
    or length(p_payload->>'sourceName') not between 1 and 240
    or jsonb_typeof(p_payload->'sourceHash') is distinct from 'string'
    or (p_payload->>'sourceHash') !~ '^[0-9a-f]{64}$'
    or jsonb_typeof(p_payload->'sourceBase64') is distinct from 'string'
    or jsonb_typeof(p_payload->'model') is distinct from 'object'
    or jsonb_typeof(p_payload->'model'->'sheets') is distinct from 'array'
    or p_payload->'model'->>'version' is distinct from '1'
    or p_payload->'model'->>'sourceHash' is distinct from p_payload->>'sourceHash' then
    raise exception using errcode='22023', message='Invalid workbook payload';
  end if;
  begin
    v_source := decode(p_payload->>'sourceBase64','base64');
  exception when others then
    raise exception using errcode='22023', message='Invalid source encoding';
  end;
  if octet_length(v_source) not between 1 and 10485760
    or encode(sha256(v_source),'hex') is distinct from p_payload->>'sourceHash'
    or substring(v_source from 1 for 4) <> decode('504b0304','hex') then
    raise exception using errcode='22023', message='Source checksum or XLSX signature is invalid';
  end if;
  -- App inputs and Excel formulas are validated by the workbook engine before this
  -- atomic commit. Editors may import a different valid workbook; viewers may not.
  insert into maestro_private.workspace_changes
    (workspace_id,revision,actor_id,action,detail,before_payload,after_payload)
    values(p_workspace_id,v_workspace.revision+1,auth.uid(),p_action,v_detail,
      v_workspace.payload,p_payload);
  update maestro_private.workspaces set payload=p_payload,
    revision=v_workspace.revision+1,saved_at=clock_timestamp() where id=p_workspace_id;
  return maestro_private.envelope(p_workspace_id,v_role);
end;
$$;

create or replace function public.maestro_undo(
  p_workspace_id uuid, p_expected_revision bigint
)
returns jsonb language plpgsql security definer set search_path = '' as $$
declare
  v_role text;
  v_workspace maestro_private.workspaces%rowtype;
  v_change maestro_private.workspace_changes%rowtype;
begin
  v_role := maestro_private.member_role(p_workspace_id);
  if v_role not in ('owner','editor') then
    raise exception using errcode='42501', message='not_authorized: editor permission required';
  end if;
  select * into v_workspace from maestro_private.workspaces
    where id=p_workspace_id for update;
  if p_expected_revision is distinct from v_workspace.revision then
    raise exception using errcode='40001', message='revision_conflict: a newer version was saved';
  end if;
  select * into v_change from maestro_private.workspace_changes
    where workspace_id=p_workspace_id and revision=v_workspace.revision;
  -- One global undo of the latest commit. A new commit enables undo again.
  -- No audit row is rewritten, and the revision always increases, including undo.
  if not found or v_change.action='undo' then
    raise exception using errcode='22023', message='There is no latest commit to undo';
  end if;
  insert into maestro_private.workspace_changes
    (workspace_id,revision,actor_id,action,detail,before_payload,after_payload,undo_target)
    values(p_workspace_id,v_workspace.revision+1,auth.uid(),'undo',
      jsonb_build_object('undoneRevision',v_change.revision,'undoneAction',v_change.action),
      v_workspace.payload,v_change.before_payload,v_change.id);
  update maestro_private.workspaces set payload=v_change.before_payload,
    revision=v_workspace.revision+1,saved_at=clock_timestamp() where id=p_workspace_id;
  return maestro_private.envelope(p_workspace_id,v_role);
end;
$$;

revoke all on all functions in schema maestro_private from public, anon, authenticated;
revoke all on function public.maestro_get_state(uuid) from public, anon, authenticated;
revoke all on function public.maestro_commit(uuid,bigint,jsonb,text,jsonb) from public, anon, authenticated;
revoke all on function public.maestro_undo(uuid,bigint) from public, anon, authenticated;
grant execute on function public.maestro_get_state(uuid) to authenticated;
grant execute on function public.maestro_commit(uuid,bigint,jsonb,text,jsonb) to authenticated;
grant execute on function public.maestro_undo(uuid,bigint) to authenticated;

comment on table maestro_private.workspace_members is
  'Invite-only allowlist; provision/revoke users through trusted administrator SQL only.';
comment on function public.maestro_commit(uuid,bigint,jsonb,text,jsonb) is
  'Atomic CAS commit with immutable snapshots. Requires an authenticated owner/editor.';
comment on function public.maestro_undo(uuid,bigint) is
  'One global latest-commit undo. Revision increases; audit history is never edited.';

-- ADMIN BOOTSTRAP (examples only; replace both UUID placeholders in the SQL editor):
-- insert into maestro_private.workspaces(name) values ('Madinah Content') returning id;
-- insert into maestro_private.workspace_members(workspace_id,user_id,role)
-- values ('<workspace-uuid>'::uuid,'<invited-auth-user-uuid>'::uuid,'owner');
-- Invite accounts through Supabase Auth, disable public sign-ups, and add their
-- auth.users UUIDs here with role owner/editor/viewer. The client has no member-
-- management or workspace-creation RPC. Never place service-role/secret keys in GitHub.
commit;
