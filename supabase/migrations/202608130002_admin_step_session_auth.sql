create or replace function private.schedule_is_admin()
returns boolean
language sql
stable
security definer
set search_path = public, auth, pg_temp
as $$
  select
    lower(coalesce(auth.jwt() ->> 'email', '')) in (
      'stepkobetsu@gmail.com',
      'stepkobetsustaff@gmail.com'
    )
    or coalesce((auth.jwt() -> 'app_metadata' ->> 'schedule_admin')::boolean, false);
$$;

revoke all on function private.schedule_is_admin() from public, anon;
grant execute on function private.schedule_is_admin() to authenticated;

comment on function private.schedule_is_admin() is
  'Allows legacy approved admin emails or immutable app_metadata.schedule_admin issued after STEP staff session verification.';
