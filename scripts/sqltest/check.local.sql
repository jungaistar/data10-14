-- 로컬 검증: 참여자만 읽기, 현장 기록은 덧붙이기만, 남의 이름으로 못 씀, 실시간 발행 등록
do $$ begin
  if exists (select 1 from pg_roles where rolname in ('supabase_admin', 'authenticator')) then raise exception '로컬 검증용 파일입니다.'; end if;
end $$;
\set ON_ERROR_STOP on
-- 관리자(M) 가 회차를 만들고 자기를 manager 로, 담당자(C) 를 counter 로 넣음
set role authenticated;
do $$ begin perform set_config('request.jwt.claim.sub', '00000000-0000-0000-0000-00000000000a', false); end $$;
insert into public.stocktake_sessions (id, code, title, base_at) values ('11111111-1111-1111-1111-111111111111', 'ST-T', '시험', '2026-10-31 18:00+09');
insert into public.stocktake_members values ('11111111-1111-1111-1111-111111111111', '00000000-0000-0000-0000-00000000000a', '관리자', 'manager');
insert into public.stocktake_members values ('11111111-1111-1111-1111-111111111111', '00000000-0000-0000-0000-00000000000c', '담당자C', 'counter');
insert into public.stocktake_items (session_id, no, wh, sys_qty, price) values ('11111111-1111-1111-1111-111111111111', 'S-001', '부품창고', 100, 1000);
insert into public.stocktake_warehouses (session_id, wh, wh_name, site, manager_name, counters) values ('11111111-1111-1111-1111-111111111111', 'W10', '가상창고', '사내', '관리자', '{담당자C,담당자D}');
do $$ begin assert (select tol_mode from public.stocktake_sessions) = 'zero', '허용 기준 기본이 허용 0 이 아님'; end $$;
-- 담당자 C: 품목을 읽고 기록을 덧붙임
do $$ begin perform set_config('request.jwt.claim.sub', '00000000-0000-0000-0000-00000000000c', false); end $$;
do $$ begin assert (select count(*) from public.stocktake_items) = 1, '담당자가 품목을 못 읽음'; end $$;
insert into public.stocktake_counts (rid, session_id, no, qty, reason, by_name, counted_at) values ('dC-1', '11111111-1111-1111-1111-111111111111', 'S-001', 97, '입출고 누락', '담당자C', now());
-- 설정에서 바꾼 사유도 들어감(고정 목록 검사 없음), 너무 긴 사유는 거절
insert into public.stocktake_counts (rid, session_id, no, qty, reason, by_name, counted_at) values ('dC-3', '11111111-1111-1111-1111-111111111111', 'S-001', 98, '보관 위치 다름', '담당자C', now());
do $$ begin
  begin
    insert into public.stocktake_counts (rid, session_id, no, qty, reason, by_name, counted_at) values ('dC-4', '11111111-1111-1111-1111-111111111111', 'S-001', 1, repeat('가', 41), '담당자C', now());
    raise exception '41자 사유가 들어감';
  exception when check_violation then null; end;
end $$;
-- 실사자는 창고 배정표를 읽기만(고치기 0 줄)
do $$ declare n int; begin
  assert (select count(*) from public.stocktake_warehouses) = 1, '실사자가 배정표를 못 읽음';
  update public.stocktake_warehouses set excluded = true; get diagnostics n = row_count; assert n = 0, '실사자가 배정표를 고침';
end $$;
-- 고치기·지우기는 정책이 없어 0 줄
do $$ declare n int; begin
  update public.stocktake_counts set qty = 100 where rid = 'dC-1'; get diagnostics n = row_count; assert n = 0, '기록을 고칠 수 있음';
  delete from public.stocktake_counts where rid = 'dC-1'; get diagnostics n = row_count; assert n = 0, '기록을 지울 수 있음';
end $$;
-- 남의 이름(counted_by)으로 쓰면 거절
do $$ begin
  begin
    insert into public.stocktake_counts (rid, session_id, no, qty, by_name, counted_at, counted_by) values ('dC-2', '11111111-1111-1111-1111-111111111111', 'S-001', 1, '가짜', now(), '00000000-0000-0000-0000-00000000000a');
    raise exception '남의 이름으로 기록됨';
  exception when insufficient_privilege then null; end;
end $$;
-- 담당자는 품목을 못 고침
do $$ declare n int; begin update public.stocktake_items set sys_qty = 0; get diagnostics n = row_count; assert n = 0, '담당자가 품목을 고침'; end $$;
-- 참여자가 아닌 사람(X)은 아무것도 못 봄
do $$ begin perform set_config('request.jwt.claim.sub', '00000000-0000-0000-0000-0000000000ff', false); end $$;
do $$ begin
  assert (select count(*) from public.stocktake_sessions) = 0, '남이 회차를 봄';
  assert (select count(*) from public.stocktake_counts) = 0, '남이 기록을 봄';
end $$;
-- 비로그인(anon)은 판정 함수도 못 부름
reset role; set role anon;
do $$ begin
  begin perform public.stocktake_role('11111111-1111-1111-1111-111111111111'); raise exception 'anon 이 판정 함수를 부름';
  exception when insufficient_privilege then null; end;
end $$;
reset role;
do $$ begin assert exists (select 1 from pg_publication_tables where pubname = 'supabase_realtime' and tablename = 'stocktake_counts'), '실시간 발행 없음'; end $$;
select '검사 통과' as 결과;
