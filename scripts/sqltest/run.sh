#!/bin/bash
# supabase/schema.sql 을 로컬 임시 PostgreSQL 에 실제로 적용해 검사합니다(운영 DB 를 건드리지 않음).
# 필요: PostgreSQL(initdb·pg_ctl·psql). 실행: bash scripts/sqltest/run.sh
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/../.." && pwd)"
TMP="$(mktemp -d)"; PORT=$((20000 + RANDOM % 20000))
trap 'pg_ctl -D "$TMP/db" -m immediate stop >/dev/null 2>&1 || true; rm -rf "$TMP"' EXIT
initdb -D "$TMP/db" -U postgres >/dev/null 2>&1
pg_ctl -D "$TMP/db" -o "-p $PORT -k $TMP" -l "$TMP/log" start >/dev/null
sleep 1
P="env PGOPTIONS=-cclient_min_messages=error psql -h $TMP -p $PORT -U postgres -v ON_ERROR_STOP=1 -q"
$P -c 'create database t' >/dev/null
echo '① Supabase 흉내 + schema.sql 두 번(재실행 안전)'
$P -d t -f "$ROOT/scripts/sqltest/stubs.local.sql"
$P -d t -f "$ROOT/supabase/schema.sql"
$P -d t -f "$ROOT/supabase/schema.sql"
echo '② 권한 검사'
$P -d t -f "$ROOT/scripts/sqltest/check.local.sql"
echo '③ 공용 프로젝트(user_profiles 있음)에서는 가드가 멈추는지'
$P -c 'create database g' >/dev/null
$P -d g -f "$ROOT/scripts/sqltest/stubs.local.sql"
$P -d g -c 'create table public.user_profiles (id uuid)'
if $P -d g -f "$ROOT/supabase/schema.sql" 2>"$TMP/err"; then echo '실패: 가드가 멈추지 않음'; exit 1; fi
grep -q '재고실사 전용' "$TMP/err" && echo '가드 멈춤 확인'
[ "$($P -d g -tAc "select count(*) from pg_tables where tablename like 'stocktake_%'")" = 0 ] && echo '표가 하나도 안 생김 확인'
echo '모두 통과'
