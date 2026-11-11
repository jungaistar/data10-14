-- 로컬 임시 PostgreSQL 전용 — Supabase 환경 흉내(auth.uid · 역할 · 실시간 발행). 운영 DB 에서는 스스로 멈춥니다.
do $$ begin
  if exists (select 1 from pg_roles where rolname in ('supabase_admin', 'authenticator')) or exists (select 1 from pg_namespace where nspname = 'graphql') then
    raise exception '로컬 검증용 파일입니다. Supabase 프로젝트에서 실행하지 마세요.';
  end if;
end $$;
create schema if not exists auth;
create or replace function auth.uid() returns uuid language sql stable as $$ select nullif(current_setting('request.jwt.claim.sub', true), '')::uuid $$;
do $$ begin
  if not exists (select 1 from pg_roles where rolname = 'anon') then create role anon; end if;
  if not exists (select 1 from pg_roles where rolname = 'authenticated') then create role authenticated; end if;
end $$;
grant usage on schema public, auth to anon, authenticated;
grant execute on function auth.uid() to anon, authenticated;
alter default privileges in schema public grant all on tables to anon, authenticated;
alter default privileges in schema public grant all on sequences to anon, authenticated;
do $$ begin if not exists (select 1 from pg_publication where pubname = 'supabase_realtime') then create publication supabase_realtime; end if; end $$;
