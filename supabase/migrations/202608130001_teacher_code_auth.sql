create schema if not exists private;

create table if not exists private.schedule_auth_config (
  key text primary key,
  value text not null,
  updated_at timestamptz not null default now()
);

create table if not exists private.schedule_login_attempts (
  attempt_key text primary key,
  attempts integer not null default 0,
  window_started_at timestamptz not null default now(),
  locked_until timestamptz
);

revoke all on private.schedule_auth_config from public, anon, authenticated;
revoke all on private.schedule_login_attempts from public, anon, authenticated;

create or replace function public.schedule_auth_settings()
returns jsonb
language sql
stable
security definer
set search_path = private, pg_temp
as $$
  select jsonb_build_object(
    'shared_secret', max(value) filter (where key = 'teacher_auth_shared_secret'),
    'api_url', max(value) filter (where key = 'teacher_auth_api_url')
  )
  from private.schedule_auth_config;
$$;

revoke all on function public.schedule_auth_settings() from public, anon, authenticated;
grant execute on function public.schedule_auth_settings() to service_role;

create or replace function public.schedule_login_rate_allowed(p_key text)
returns boolean
language sql
stable
security definer
set search_path = private, pg_temp
as $$
  select not exists (
    select 1 from private.schedule_login_attempts
    where attempt_key = p_key and locked_until > now()
  );
$$;

revoke all on function public.schedule_login_rate_allowed(text) from public, anon, authenticated;
grant execute on function public.schedule_login_rate_allowed(text) to service_role;

create or replace function public.schedule_login_rate_record(p_key text, p_limit integer, p_success boolean)
returns void
language plpgsql
security definer
set search_path = private, pg_temp
as $$
begin
  if p_success then
    delete from private.schedule_login_attempts where attempt_key = p_key;
    return;
  end if;

  insert into private.schedule_login_attempts(attempt_key, attempts, window_started_at, locked_until)
  values (p_key, 1, now(), null)
  on conflict (attempt_key) do update set
    attempts = case
      when schedule_login_attempts.window_started_at < now() - interval '15 minutes' then 1
      else schedule_login_attempts.attempts + 1
    end,
    window_started_at = case
      when schedule_login_attempts.window_started_at < now() - interval '15 minutes' then now()
      else schedule_login_attempts.window_started_at
    end,
    locked_until = case
      when (case when schedule_login_attempts.window_started_at < now() - interval '15 minutes' then 1 else schedule_login_attempts.attempts + 1 end) >= greatest(p_limit, 1)
      then now() + interval '15 minutes'
      else schedule_login_attempts.locked_until
    end;
end;
$$;

revoke all on function public.schedule_login_rate_record(text, integer, boolean) from public, anon, authenticated;
grant execute on function public.schedule_login_rate_record(text, integer, boolean) to service_role;

create or replace function private.schedule_teacher_code()
returns text
language sql
stable
security invoker
set search_path = pg_catalog
as $$
  select case
    when coalesce((auth.jwt() -> 'app_metadata' ->> 'schedule_teacher')::boolean, false)
    then coalesce(auth.jwt() -> 'app_metadata' ->> 'schedule_teacher_code', '')
    else ''
  end;
$$;

revoke all on function private.schedule_teacher_code() from public, anon;
grant execute on function private.schedule_teacher_code() to authenticated;

create or replace function private.schedule_is_teacher()
returns boolean
language sql
stable
security definer
set search_path = public, auth, pg_temp
as $$
  select exists (
    select 1 from public.teachers
    where active is true
      and code = private.schedule_teacher_code()
  );
$$;

create or replace function private.schedule_teacher_owns(p_teacher_code text)
returns boolean
language sql
stable
security definer
set search_path = public, auth, pg_temp
as $$
  select exists (
    select 1 from public.teachers
    where active is true
      and code = p_teacher_code
      and code = private.schedule_teacher_code()
  );
$$;

drop policy if exists schedule_teacher_own_record on public.teachers;
create policy schedule_teacher_own_record
on public.teachers for select
to authenticated
using (active is true and code = private.schedule_teacher_code());

comment on function private.schedule_teacher_code() is 'Returns the immutable schedule teacher code from auth app_metadata.';
