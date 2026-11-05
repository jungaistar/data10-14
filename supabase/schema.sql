-- 재고실사 Start-End — 2단계 「실시간 취합」용 DB 초안 (data10-14)
--
-- ※ 수강생 본인(또는 회사가 허용한) Supabase 프로젝트 전용입니다. 1단계 앱은 이 DB 에 연결하지 않습니다.
-- ※ 다른 서비스와 함께 쓰는 공용 프로젝트에서는 실행하지 마세요. 아래 §0 가드가 그런 프로젝트를 알아보면 멈춥니다.
-- ※ 재실행해도 안전하게 썼습니다(IF NOT EXISTS · DROP POLICY IF EXISTS).
--
-- 표 구성(1단계 브라우저 저장 구조를 그대로 옮김)
--   stocktake_sessions    실사 회차(기준 일자·시각, 지시서 버전)            ← 인벤토리 데이터 01
--   stocktake_members     회차별 참여자(manager = 실사 총괄, counter = 실사자)
--   stocktake_warehouses  창고 배정표(저장위치 · 저장위치명 · 창고위치 · 담당자 · 실사자 · 제외 여부)  ← 2차: 수강생 시트 2
--   stocktake_items       ERP 에서 올린 재고 품목(시스템 수량·단가, 원래 열 그대로 ext)
--   stocktake_assignments 창고(#나눈 몫)·구역별 실사자 배정
--   stocktake_counts      현장 기록 — 덧붙이기만(고치면 새 줄, 이전 줄은 남음)  ← 데이터 02·03·04
--   stocktake_picks       충돌 때 채택한 기록
--   stocktake_reports     보고서 발행 이력(버전·기준 시각·기록 수)           ← 데이터 06
-- 현장 기록은 수정·삭제 정책을 두지 않습니다(회계감사 근거). 고친 것은 새 기록으로 남깁니다.

-- §0 실행 위치 가드 — 여러 사이트가 함께 쓰는 공용 프로젝트(공용 회원 표가 있는 곳)면 멈춥니다
do $$
begin
  if to_regclass('public.www_profiles') is not null or to_regclass('public.user_profiles') is not null then
    raise exception '이 파일은 재고실사 전용(본인) Supabase 프로젝트용입니다. 공용 회원 표(www_profiles / user_profiles)가 있는 프로젝트에서는 실행하지 마세요.';
  end if;
end $$;

-- §1 표
create table if not exists public.stocktake_sessions (
  id          uuid primary key default gen_random_uuid(),
  code        text not null unique,                 -- 지시서 ID (예: ST-202610311800-AB12)
  title       text not null,
  base_at     timestamptz not null,                 -- 기준 일자·시각(ERP 재고를 내려받은 시각)
  source_file text,
  unit_mode   text not null default 'wh' check (unit_mode in ('wh', 'zone')),
  order_ver   int  not null default 1,
  tol_pct     numeric not null default 2,           -- 허용 범위(tol_mode = 'range' 일 때만)
  tol_amt     numeric not null default 100000,
  created_by  uuid not null default auth.uid(),
  created_at  timestamptz not null default now()
);

create table if not exists public.stocktake_members (
  session_id uuid not null references public.stocktake_sessions(id) on delete cascade,
  user_id    uuid not null,
  name       text not null,
  role       text not null default 'counter' check (role in ('manager', 'counter')),
  primary key (session_id, user_id)
);

create table if not exists public.stocktake_items (
  session_id uuid not null references public.stocktake_sessions(id) on delete cascade,
  no         text not null,                         -- 재고번호(없으면 자재코드@창고/위치)
  code text, name text, spec text, unit text,
  wh   text not null, loc text, mtype text, mkind text,
  sys_qty numeric not null default 0,
  price   numeric not null default 0,
  primary key (session_id, no)
);

-- 2차 반영(재실행 안전): 허용 기준 기본 = 허용 0(차이가 있으면 모두 사유 필수), 사유 목록 설정, 배정표 파일
alter table public.stocktake_sessions add column if not exists tol_mode text not null default 'zero';
alter table public.stocktake_sessions add column if not exists reasons text[];
alter table public.stocktake_sessions add column if not exists roster_file text;
alter table public.stocktake_sessions drop constraint if exists stocktake_sessions_tol_mode_check;
alter table public.stocktake_sessions add constraint stocktake_sessions_tol_mode_check check (tol_mode in ('zero', 'range'));
alter table public.stocktake_items add column if not exists wh_name text;   -- 저장위치명
alter table public.stocktake_items add column if not exists site text;      -- 창고위치(사내/사외)
alter table public.stocktake_items add column if not exists part int;       -- 실사자 여럿인 창고를 나눈 몫(1, 2 …)
alter table public.stocktake_items add column if not exists ext jsonb;      -- ERP 에 추가된 열(필드가 추가될 수 있음)

create table if not exists public.stocktake_warehouses (
  session_id   uuid not null references public.stocktake_sessions(id) on delete cascade,
  wh           text not null,                       -- 저장위치(창고 코드)
  wh_name      text,                                -- 저장위치명
  site         text,                                -- 창고위치(사내/사외)
  team         text,
  manager_name text,                                -- 담당자(창고 책임 · 서명)
  counters     text[] not null default '{}',        -- 실사자(여럿이면 보관 자리 순서로 나눔)
  excluded     boolean not null default false,      -- 실사인원 0 → 이번 실사 제외
  note         text,
  primary key (session_id, wh)
);

create table if not exists public.stocktake_assignments (
  session_id uuid not null references public.stocktake_sessions(id) on delete cascade,
  unit_key   text not null,                         -- 창고, 「창고 #1」(나눈 몫) 또는 「창고 / 구역」
  assignee   uuid,
  assignee_name text not null,
  primary key (session_id, unit_key)
);

create table if not exists public.stocktake_counts (
  id          bigint generated always as identity primary key,
  rid         text not null unique,                 -- 기기에서 만든 기록 ID(같은 기록을 두 번 올려도 한 번만)
  session_id  uuid not null references public.stocktake_sessions(id) on delete cascade,
  no          text not null,
  qty         numeric not null check (qty >= 0),
  reason      text,                                 -- 사유(목록은 회차 설정 reasons — 아래 길이 검사만)
  note        text,
  extra       jsonb,                                -- 목록에 없는 자재(자재코드·자재명·창고·위치·단위)
  photo_path  text,                                 -- Storage 비공개 버킷 경로(사진 자체는 표에 넣지 않음)
  counted_by  uuid not null default auth.uid(),     -- 작성자(서버가 찍음)
  by_name     text not null,
  counted_at  timestamptz not null,                 -- 확인 시각(기기 시계)
  received_at timestamptz not null default now(),   -- 서버가 받은 시각
  device      text
);
-- 2차: 사유 목록을 설정에서 바꿀 수 있어 고정 목록 검사 대신 길이만 봄(1단계 표에서 옮겨 와도 안전)
alter table public.stocktake_counts drop constraint if exists stocktake_counts_reason_check;
alter table public.stocktake_counts add constraint stocktake_counts_reason_check check (reason is null or char_length(reason) between 1 and 40);
create index if not exists stocktake_counts_session_no on public.stocktake_counts (session_id, no, counted_at);

create table if not exists public.stocktake_picks (
  session_id uuid not null references public.stocktake_sessions(id) on delete cascade,
  no         text not null,
  rid        text not null references public.stocktake_counts(rid),
  picked_by  uuid not null default auth.uid(),
  picked_at  timestamptz not null default now(),
  primary key (session_id, no)
);

create table if not exists public.stocktake_reports (
  id         bigint generated always as identity primary key,
  session_id uuid not null references public.stocktake_sessions(id) on delete cascade,
  ver        int not null,
  issued_at  timestamptz not null default now(),
  issued_by  uuid not null default auth.uid(),
  records    int not null,
  counted    int not null,
  net_amt    numeric,
  unique (session_id, ver)
);

-- §2 판정 함수 (정책 식에서 씀 — 이 프로젝트 전용 이름)
create or replace function public.stocktake_role(p_session uuid)
returns text language sql stable security definer set search_path = public as $$
  select role from public.stocktake_members where session_id = p_session and user_id = auth.uid()
$$;
revoke all on function public.stocktake_role(uuid) from public, anon;
grant execute on function public.stocktake_role(uuid) to authenticated;

-- §3 RLS — 로그인한 참여자만, 현장 기록은 덧붙이기만
alter table public.stocktake_sessions    enable row level security;
alter table public.stocktake_members     enable row level security;
alter table public.stocktake_items       enable row level security;
alter table public.stocktake_warehouses  enable row level security;
alter table public.stocktake_assignments enable row level security;
alter table public.stocktake_counts      enable row level security;
alter table public.stocktake_picks       enable row level security;
alter table public.stocktake_reports     enable row level security;

drop policy if exists st_sessions_read   on public.stocktake_sessions;
drop policy if exists st_sessions_insert on public.stocktake_sessions;
drop policy if exists st_sessions_update on public.stocktake_sessions;
create policy st_sessions_read   on public.stocktake_sessions for select to authenticated using (public.stocktake_role(id) is not null or created_by = auth.uid());
create policy st_sessions_insert on public.stocktake_sessions for insert to authenticated with check (created_by = auth.uid());
create policy st_sessions_update on public.stocktake_sessions for update to authenticated using (public.stocktake_role(id) = 'manager');

drop policy if exists st_members_read  on public.stocktake_members;
drop policy if exists st_members_write on public.stocktake_members;
create policy st_members_read  on public.stocktake_members for select to authenticated using (public.stocktake_role(session_id) is not null);
-- 회차를 만든 사람이 자기를 manager 로 넣는 첫 줄, 그 뒤로는 manager 가 관리
create policy st_members_write on public.stocktake_members for all to authenticated
  using (public.stocktake_role(session_id) = 'manager')
  with check (public.stocktake_role(session_id) = 'manager'
    or exists (select 1 from public.stocktake_sessions s where s.id = session_id and s.created_by = auth.uid()));

drop policy if exists st_items_read  on public.stocktake_items;
drop policy if exists st_items_write on public.stocktake_items;
create policy st_items_read  on public.stocktake_items for select to authenticated using (public.stocktake_role(session_id) is not null);
create policy st_items_write on public.stocktake_items for all to authenticated using (public.stocktake_role(session_id) = 'manager') with check (public.stocktake_role(session_id) = 'manager');

drop policy if exists st_wh_read  on public.stocktake_warehouses;
drop policy if exists st_wh_write on public.stocktake_warehouses;
create policy st_wh_read  on public.stocktake_warehouses for select to authenticated using (public.stocktake_role(session_id) is not null);
create policy st_wh_write on public.stocktake_warehouses for all to authenticated using (public.stocktake_role(session_id) = 'manager') with check (public.stocktake_role(session_id) = 'manager');

drop policy if exists st_assign_read  on public.stocktake_assignments;
drop policy if exists st_assign_write on public.stocktake_assignments;
create policy st_assign_read  on public.stocktake_assignments for select to authenticated using (public.stocktake_role(session_id) is not null);
create policy st_assign_write on public.stocktake_assignments for all to authenticated using (public.stocktake_role(session_id) = 'manager') with check (public.stocktake_role(session_id) = 'manager');

drop policy if exists st_counts_read   on public.stocktake_counts;
drop policy if exists st_counts_insert on public.stocktake_counts;
create policy st_counts_read   on public.stocktake_counts for select to authenticated using (public.stocktake_role(session_id) is not null);
create policy st_counts_insert on public.stocktake_counts for insert to authenticated with check (public.stocktake_role(session_id) is not null and counted_by = auth.uid());
-- UPDATE · DELETE 정책 없음(기록성 표)

drop policy if exists st_picks_read  on public.stocktake_picks;
drop policy if exists st_picks_write on public.stocktake_picks;
create policy st_picks_read  on public.stocktake_picks for select to authenticated using (public.stocktake_role(session_id) is not null);
create policy st_picks_write on public.stocktake_picks for all to authenticated using (public.stocktake_role(session_id) = 'manager') with check (public.stocktake_role(session_id) = 'manager' and picked_by = auth.uid());

drop policy if exists st_reports_read   on public.stocktake_reports;
drop policy if exists st_reports_insert on public.stocktake_reports;
create policy st_reports_read   on public.stocktake_reports for select to authenticated using (public.stocktake_role(session_id) is not null);
create policy st_reports_insert on public.stocktake_reports for insert to authenticated with check (public.stocktake_role(session_id) = 'manager' and issued_by = auth.uid());

-- §4 실시간 — 현장 기록이 들어오면 사무실 화면에 바로 보이게(Supabase Realtime)
do $$
begin
  if exists (select 1 from pg_publication where pubname = 'supabase_realtime')
     and not exists (select 1 from pg_publication_tables where pubname = 'supabase_realtime' and schemaname = 'public' and tablename = 'stocktake_counts') then
    alter publication supabase_realtime add table public.stocktake_counts;
  end if;
end $$;
