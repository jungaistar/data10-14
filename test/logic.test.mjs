// 실행: node test/logic.test.mjs   (의존성 없음 — vendor 의 SheetJS 만 씀)
// 2차 반영: 수강생이 보낸 ERP 재고 엑셀 견본 · 창고 배정표의 「모양」(열 이름 · 줄 배치)을 그대로 본뜬 가상 표로 검사합니다.
// 실제 값(자재번호 · 창고명 · 사람 이름)은 리포에 넣지 않습니다. 가상 샘플(248품목)은 js/sample-data.js.
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const L = require('../js/logic.js');
const Sample = require('../js/sample-data.js');
const XLSX = require('../vendor/xlsx.full.min.js');
const E = require('../js/ai-endpoint.js');

let passed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log('  ok  ' + name); }
  catch (e) { console.error('  FAIL ' + name + '\n       ' + e.message); process.exitCode = 1; }
}
const near = (a, b, eps = 1e-6) => assert.ok(Math.abs(a - b) < eps, a + ' ≠ ' + b);

// ERP 에서 내려받은 모양: 제목 · 조회 조건 · 빈 줄 · 머리 · 자료 · 합계
const ERP = [
  ['재고 현황 조회'], ['조회 기준 2026-10-31 18:00'], [],
  ['재고번호', '창고', '저장위치', '자재코드', '자재내역', '단위', '자재유형', '자재타입', '재고수량(EA)', '단가', '재고금액'],
  ['S-001', '부품창고', 'A-01-1', 'M001', '유압 밸브', 'EA', '유압부품', '부자재', '100', '1,000', 100000],
  ['S-002', '부품창고', 'A-02-1', 'M002', 'O-링', 'EA', '실·패킹', '부자재', 50, null, 25000],          // 단가 빈칸 → 금액 ÷ 수량
  ['S-003', '부품창고', 'B-01-1', 'M003', '볼트', 'EA', '볼트·너트', '부자재', null, 10, null],           // 수량 빈칸 → 0
  ['S-004', '제품창고', 'A-01-1', 'M004', '유압 실린더', 'EA', '완제품', '제품', 10, 500000, 5000000],
  ['S-005', '제품창고', 'A-01-2', 'M005', '반제품', 'EA', '반제품', '반제품', '열개', 1000, 0],           // 수량 못 읽음 → 건너뜀
  [null, '제품창고', 'A-01-3', null, '번호 없음', 'EA', '완제품', '제품', 1, 1, 1],                       // 열쇠 없음 → 건너뜀
  [],
  ['합계', null, null, null, null, null, null, null, 160, null, 5125000]
];
const parsed = () => L.parseErp(ERP);
const session = (mode) => L.makeSession({ id: 'ST-TEST', title: '테스트 실사', baseAt: '2026-10-31T18:00', sourceFile: 'erp.xlsx', unitMode: mode || 'wh', now: '2026-10-31T18:10:00' });
function dbWith(records, extra) {
  const p = parsed();
  return Object.assign(L.restoreDb({ session: session(), items: p.items, assign: { 부품창고: '김', 제품창고: '이' }, records: records || [] }), extra || {});
}
const rec = (no, qty, by, at, reason, rid) => ({ rid: rid || (no + by + at), no, qty, by, at, reason: reason || '', note: '', src: '테스트' });
const toBook = sheets => {
  const wb = XLSX.utils.book_new();
  sheets.forEach(x => { const ws = XLSX.utils.aoa_to_sheet(x.aoa); (x.formula || []).forEach(f => { ws[f.a] = f.v == null ? { t: 's', v: '', f: f.f } : { t: 'n', v: f.v, f: f.f }; }); XLSX.utils.book_append_sheet(wb, ws, x.name); });
  const back = XLSX.read(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }), { type: 'buffer', cellDates: false });
  return back.SheetNames.map(n => ({ name: n, aoa: XLSX.utils.sheet_to_json(back.Sheets[n], { header: 1, raw: true, defval: null }) }));
};
const col = (aoa, h) => aoa[0].indexOf(h);

// 수강생 견본과 같은 모양(값은 가짜): 첫 줄 안내 · 빈 줄 2개 · 머리행 9칸
const REAL = [
  ['샘플임. 필드가 추가될 수 있음.'], [], [],
  ['자재코드', '자재내역', '저장위치', '저장위치명', '창고위치', '단가', '금액', 'Location', '수량'],
  ['X01', '가상 판재', 'W10', '가상창고1', '사내', 200, 600, '5번 랙', 3],
  ['X02', '가상 봉재', 'W10', '가상창고1', '사내', 50, 500, '6번 랙', 10],
  ['X03', '가상 볼트', 'W10', '가상창고1', '사내', 10, 1000, '7번 랙', 100],
  ['X04', '가상 커버', 'W10', '가상창고1', '사내', 1000, 4000, '8번 랙', 4],
  ['X05', '가상 호스', 'W20', '가상창고2', '사외', 300, 3000, 'A-1', 10],
  ['X06', '가상 시편', 'W30', '가상창고3', '사내', 100, 100, 'A-1', 1]
];
// 창고 배정표 모양: 안내 줄 · 제목 · 빈 줄 · 날짜와 제목 · 빈 줄 · 머리행
const ROSTER = [
  ['작업배정 시 활용하는 자료'], ['창고 목록과 창고별 담당 인원'], [],
  [46316, '재고실사 배정표'], [],
  ['저장위치', '저장위치명', '창고위치', '실사인원', '담당팀', '담당자', '실사자', '비고'],
  ['W30', '가상창고3', '사내', 0, '가팀', '관리갑', '-', '이번 실사 제외, 확인서로 대체'],
  ['W10', '가상창고1', '사내', 2, '나팀', '관리을', '실사갑, 실사을', null],
  ['W20', '가상창고2', '사외', 1, '나팀', '관리을', '실사병', null],
  ['W40', '가상창고4', '사내', 1, '나팀', '관리병', '실사정', null]
];

console.log('열 찾기 · ERP 읽기');
test('제목·조회 조건 아래 머리행을 찾고 12개 항목 중 있는 열을 모두 연결(단위 괄호 무시)', () => {
  const d = L.detectColumns(ERP);
  assert.equal(d.headerRow, 3);
  assert.deepEqual(d.missing, []);
  const want = { no: 0, wh: 1, loc: 2, code: 3, name: 4, unit: 5, mtype: 6, mkind: 7, sys: 8, price: 9, amt: 10 };
  Object.keys(want).forEach(k => assert.equal(d.map[k], want[k], k));
});
test('열 이름이 달라도 찾음 — 자재/자재설명/기본단위/현재고/이동평균단가/로케이션/창고코드', () => {
  assert.equal(L.matchField('자재'), 'code');
  assert.equal(L.matchField('자재 설명'), 'name');
  assert.equal(L.matchField('현재고'), 'sys');
  assert.equal(L.matchField('장부 수량(EA)'), 'sys');
  assert.equal(L.matchField('이동평균단가'), 'price');
  assert.equal(L.matchField('로케이션'), 'loc');
  assert.equal(L.matchField('창고 코드'), 'wh');
  assert.equal(L.matchField('실사 수량'), 'cnt');
  assert.equal(L.matchField('확인 일시'), 'at');
});
test('필수 열(재고번호·자재코드 / 창고 / 수량)이 없으면 알리고, 열 연결(덮어쓰기)로 고칠 수 있음', () => {
  const aoa = [['번호', '보관처', '재고수량'], ['A1', '창고1', 3]];
  const d = L.detectColumns(aoa);
  assert.ok(d.missing.length > 0);
  const d2 = L.detectColumns(aoa, { no: 0, wh: 1 });
  assert.deepEqual(d2.missing, []);
  const p = L.parseErp(aoa, { overrides: { no: 0, wh: 1 } });
  assert.equal(p.items.length, 1);
  assert.equal(p.items[0].wh, '창고1');
});
test('합계 줄·빈 줄 빼고, 수량 빈칸 0, 못 읽은 수량·열쇠 없는 줄은 건너뜀(이유 남김)', () => {
  const p = parsed();
  assert.equal(p.items.length, 4);
  assert.equal(p.summary, 1);
  assert.equal(p.blankQty, 1);
  assert.equal(p.skipped.length, 2);
  assert.ok(p.skipped.some(x => /숫자/.test(x.why)) && p.skipped.some(x => /빈칸/.test(x.why)));
});
test('단가가 없으면 금액 ÷ 수량, 금액이 없으면 수량 × 단가, 「1,000」 같은 쉼표 숫자', () => {
  const p = parsed();
  const by = Object.fromEntries(p.items.map(x => [x.no, x]));
  assert.equal(by['S-001'].price, 1000);
  assert.equal(by['S-002'].price, 500);
  assert.equal(by['S-003'].sys, 0);
  assert.equal(by['S-003'].amt, 0);
});
test('재고번호 열이 없으면 「자재코드@창고/위치」를 열쇠로 — 같은 열쇠의 줄은 합치지 않고 「#2」를 붙여 따로', () => {
  const aoa = [['자재코드', '창고', '저장위치', '재고수량', '단가'], ['M1', 'W1', 'A', 10, 100], ['M1', 'W1', 'A', 5, 100], ['M1', 'W1', 'B', 1, 100]];
  const p = L.parseErp(aoa);
  assert.ok(p.generatedKey);
  assert.equal(p.items.length, 3);
  assert.deepEqual(p.items.map(x => x.no), ['M1@W1/A', 'M1@W1/A#2', 'M1@W1/B']);
  assert.deepEqual(p.items.map(x => x.sys), [10, 5, 1]);
  assert.equal(p.dupKeys, 1);
  assert.ok(p.warnings.some(w => /합치지 않고/.test(w)), p.warnings.join(' / '));
});

/* 10-28 수강생 데이터 신뢰도 지적(③)의 모양 — 값은 모두 가짜.
   같은 자재 · 같은 Location 인데 저장 위치(창고)가 다른 두 줄(R1 175 / R2 20), 같은 자재가 같은 창고의 다른 Location 에 더,
   완전히 같은 열쇠(자재 · 창고 · Location)도 한 줄 더. 「저장 위치 내역」 머리가 있고 「Location」 은 뒤쪽에 있음.
   예전: 「저장 위치 내역」이 보관 자리로 읽혀 열쇠가 「자재@창고/창고이름」 → 같은 창고의 줄이 모두 한 줄로 합쳐져 기말재고가 그 합으로 찍힘 */
const RAW = [
  ['자재', '자재내역', '저장 위치', '저장 위치 내역', '단가', '금액', '내역', 'Location', '기말재고', '최종 입고일', '자재 유형', '자재내역명'],
  ['Z900-1', '가상 판재', 'R1', '가상원재료창고', 10, 1750, '원자재', 'L-1-01', 175, 45930, 'ROH', '3. 원자재'],
  ['Z900-1', '가상 판재', 'R2', '가상라인창고', 10, 200, '원자재', 'L-1-01', 20, 45931, 'ROH', '3. 원자재'],
  ['Z900-1', '가상 판재', 'R1', '가상원재료창고', 10, 3000, '원자재', 'L-2-05', 300, null, 'ROH', '3. 원자재'],
  ['Z900-1', '가상 판재', 'R1', '가상원재료창고', 10, 710, '원자재', 'L-3-02', 71, null, 'ROH', '3. 원자재'],
  ['Z900-1', '가상 판재', 'R1', '가상원재료창고', 10, 40, '원자재', 'L-1-01', 4, null, 'ROH', '3. 원자재'],
  ['Z777-2', '가상 봉재', 'R2', '가상라인창고', 1000, 12000000, '원자재', 'L-1-01', 12000, null, 'ROH', '3. 원자재'],
  ['Z900-1', '가상 판재', 'R3', '가상외주창고', 10, 90, '원자재', 'L-1-01', 9, null, 'ROH', '3. 원자재']
];
const rawSum = wh => RAW.slice(1).filter(r => r[2] === wh).reduce((s, r) => s + r[8], 0);
test('ERP 줄마다 한 줄 — 같은 자재 · 같은 Location · 다른 저장 위치는 따로, 기말재고는 그 줄 값, 창고별 합계 = ERP 합계 (10-28 ③)', () => {
  const d = L.detectColumns(RAW);
  assert.equal(d.map.loc, 7, '보관 자리 = 「Location」 열(「저장 위치 내역」이 아님)');
  assert.equal(d.map.whName, 3, '「저장 위치 내역」 = 창고명');
  assert.equal(d.map.wh, 2); assert.equal(d.map.sys, 8);
  const p = L.parseErp(RAW);
  assert.equal(p.items.length, RAW.length - 1, 'ERP 7줄 → 7품목');
  assert.equal(new Set(p.items.map(x => x.no)).size, p.items.length, '재고번호가 줄마다 하나');
  p.items.forEach((it, i) => { assert.equal(it.sys, RAW[i + 1][8], it.no); assert.equal(it.wh, RAW[i + 1][2]); assert.equal(it.loc, RAW[i + 1][7]); });
  const r1 = p.items.find(x => x.wh === 'R1' && x.loc === 'L-1-01'), r2 = p.items.find(x => x.wh === 'R2' && x.code === 'Z900-1');
  assert.equal(r1.no, 'Z900-1@R1/L-1-01'); assert.equal(r1.sys, 175);
  assert.equal(r2.no, 'Z900-1@R2/L-1-01'); assert.equal(r2.sys, 20);
  assert.equal(p.items[4].no, 'Z900-1@R1/L-1-01#2'); assert.equal(p.items[4].sys, 4);
  ['R1', 'R2', 'R3'].forEach(wh => assert.equal(p.items.filter(x => x.wh === wh).reduce((s, x) => s + x.sys, 0), rawSum(wh), wh + ' 합계'));
  assert.equal(r1.ext['최종 입고일'], L.toDateTime(45930).slice(0, 10), '날짜 열은 yyyy-mm-dd');
  assert.match(r1.ext['최종 입고일'], /^\d{4}-\d\d-\d\d$/);
  // 내려받는 기록서도 올린 그대로 — 줄 수 · 줄마다 기말재고
  const s = L.makeSession({ id: 'ST-RAW', baseAt: '2026-10-28T09:00', cols: p.cols, now: '2026-10-28T09:01:00' });
  const sh = L.orderSheets(s, p.items, { R1: '실사갑', R2: '실사을', R3: '실사병' }, {});
  const main = sh.find(x => x.name === L.MAIN_SHEET).aoa;
  assert.equal(main.length - 1, RAW.length - 1, '기록서 줄 수 = ERP 줄 수');
  const ci = main[0].indexOf('기말재고'), ni = main[0].indexOf('재고번호');
  assert.equal(main.find(r => r[ni] === 'Z900-1@R1/L-1-01')[ci], 175);
  assert.equal(main.find(r => r[ni] === 'Z900-1@R2/L-1-01')[ci], 20);
  assert.equal(main.slice(1).reduce((t, r) => t + r[ci], 0), RAW.slice(1).reduce((t, r) => t + r[8], 0));
});
test('같은 열쇠가 둘인 기록서를 채워 다시 넣어도(합치기 · 두 번 넣기) 줄마다 제 값 (10-28 ③)', () => {
  const p = L.parseErp(RAW);
  const s = L.makeSession({ id: 'ST-RAW', baseAt: '2026-10-28T09:00', cols: p.cols, now: '2026-10-28T09:01:00' });
  const assign = { R1: '실사갑', R2: '실사을', R3: '실사병' };
  const sh = L.orderSheets(s, p.items, assign, {});
  const main = sh.find(x => x.name === L.MAIN_SHEET).aoa, ni = main[0].indexOf('재고번호'), ci = main[0].indexOf('실물수량');
  const set = (aoa, no, v) => { const r = aoa.find(x => x[ni] === no); assert.ok(r, no); r[ci] = v; };
  set(main, 'Z900-1@R1/L-1-01', 170); set(main, 'Z900-1@R1/L-1-01#2', 4); set(main, 'Z900-1@R2/L-1-01', 20);
  set(sh.find(x => x.name === '실사병').aoa, 'Z900-1@R3/L-1-01', 8);
  const book = toBook(sh);
  const w = L.parseOrderWorkbook(book, { fileName: 'sp.xlsx', now: '2026-10-28T12:00:00' });
  assert.equal(w.error, undefined);
  assert.equal(w.items.length, RAW.length - 1, '다시 읽어도 7품목');
  assert.deepEqual(w.items.map(x => x.no).sort(), p.items.map(x => x.no).sort());
  w.items.forEach(it => assert.equal(it.sys, p.items.find(x => x.no === it.no).sys, it.no));
  const q = Object.fromEntries(w.records.map(r => [r.no, r.qty]));
  assert.deepEqual(q, { 'Z900-1@R1/L-1-01': 170, 'Z900-1@R1/L-1-01#2': 4, 'Z900-1@R2/L-1-01': 20, 'Z900-1@R3/L-1-01': 8 });
  // 같은 파일을 두 번 합쳐도 한 번만
  const again = L.mergeRecords(w.records, L.parseOrderWorkbook(book, { fileName: 'sp.xlsx', now: '2026-10-28T13:00:00' }).records);
  assert.equal(again.added, 0); assert.equal(again.list.length, 4);
  const db = L.restoreDb({ session: s, items: p.items, assign, records: again.list });
  const vr = L.varianceRows(db);
  assert.equal(vr.length, RAW.length - 1);
  const v = Object.fromEntries(vr.map(r => [r.no, r]));
  assert.equal(v['Z900-1@R1/L-1-01'].diff, -5); assert.equal(v['Z900-1@R1/L-1-01#2'].diff, 0); assert.equal(v['Z900-1@R2/L-1-01'].diff, 0); assert.equal(v['Z900-1@R3/L-1-01'].diff, -1);
  const an = L.analyze(db);
  ['R1', 'R2', 'R3'].forEach(wh => { const x = an.dims.wh.find(d => d.key.split(' ')[0] === wh); assert.ok(x, wh); assert.equal(x.items, RAW.slice(1).filter(r => r[2] === wh).length, wh + ' 품목 수'); });
});
test('엑셀 서식 — 수량·금액 숫자 칸은 숫자 그대로 #,##0, 날짜는 yyyy-mm-dd 날짜 칸 (10-28 ①②)', () => {
  const ws = XLSX.utils.aoa_to_sheet([['기말재고', '단가', '최종 입고일', '메모'], [12000, 1234.5, '2025-09-30', '2025-09-30 메모']]);
  ws.E2 = { t: 's', v: '', f: 'IF(A2="","",A2-A2)' };
  L.formatCells(ws);
  assert.equal(ws.A2.t, 'n'); assert.equal(ws.A2.v, 12000); assert.equal(ws.A2.z, '#,##0');
  assert.equal(ws.B2.z, '#,##0.###');
  assert.equal(ws.C2.t, 'n'); assert.equal(ws.C2.z, 'yyyy-mm-dd'); assert.equal(L.toDateTime(ws.C2.v).slice(0, 10), '2025-09-30');
  assert.equal(ws.D2.t, 's'); assert.equal(ws.A1.t, 's');
  assert.equal(ws.E2.z, '#,##0');
  const back = XLSX.read(XLSX.write({ SheetNames: ['a'], Sheets: { a: ws } }, { type: 'buffer', bookType: 'xlsx' }), { type: 'buffer', cellNF: true });
  const b = back.Sheets.a;
  assert.equal(b.A2.v, 12000); assert.equal(b.A2.z, '#,##0'); assert.equal(b.A2.w, '12,000');
  assert.equal(b.C2.z, 'yyyy-mm-dd'); assert.equal(b.C2.w, '2025-09-30');
  assert.equal(L.fmtNum(12000), '12,000');
  assert.ok(L.isDateHead('최종 입고일') && L.isDateHead('생산일자') && !L.isDateHead('단위') && !L.isDateHead('자재'));
});

test('수강생 ERP 견본 모양 — 안내 줄 아래 머리행, 저장위치 = 창고 코드, 저장위치명 = 창고명, 창고위치 = 사내/사외, Location = 보관 자리, 수량 = 시스템 수량', () => {
  const d = L.detectColumns(REAL);
  assert.equal(d.headerRow, 3);
  assert.deepEqual(d.missing, []);
  const want = { code: 0, name: 1, wh: 2, whName: 3, site: 4, price: 5, amt: 6, loc: 7, sys: 8 };
  Object.keys(want).forEach(k => assert.equal(d.map[k], want[k], k));
  assert.match(d.info, /자재유형·자재타입 열이 없어/);
  const p = L.parseErp(REAL);
  assert.equal(p.items.length, 6);
  const x = p.items[0];
  assert.equal(x.no, 'X01@W10/5번 랙'); assert.equal(x.wh, 'W10'); assert.equal(x.whName, '가상창고1'); assert.equal(x.site, '사내'); assert.equal(x.loc, '5번 랙');
  assert.equal(x.sys, 3); assert.equal(x.price, 200);
  assert.equal(L.whLabel(x), 'W10 가상창고1');
});
test('1단계 모양(창고 + 저장위치 = 보관 자리)도 그대로 — 「창고」가 따로 있고 보관 자리 열이 없으면 저장위치를 보관 자리로', () => {
  const d = L.detectColumns(ERP);
  assert.equal(d.map.wh, 1); assert.equal(d.map.loc, 2);
  assert.equal(d.warnings.length, 0, d.warnings.join(' / '));
});
test('「필드가 추가될 수 있음」 — 모르는 열도 버리지 않고 품목에 담아 기록서에 원래 순서대로 실음', () => {
  const aoa = REAL.map((r, i) => i < 3 ? r : i === 3 ? r.concat(['최종입고일', '자재유형']) : r.concat(['2026-09-01', '판재']));
  const p = L.parseErp(aoa);
  assert.equal(p.items[0].ext['최종입고일'], '2026-09-01');
  assert.equal(p.items[0].mtype, '판재', '자재유형은 알아서 연결');
  const s = L.makeSession({ id: 'ST-X', baseAt: '2026-10-21T18:00', cols: p.cols });
  const sh = L.orderSheets(s, p.items, {});
  assert.deepEqual(sh[1].aoa[0], ['재고번호', '자재코드', '자재내역', '저장위치', '저장위치명', '창고위치', '단가', '금액', 'Location', '수량', '최종입고일', '자재유형', ...L.COUNT_HEAD]);
  assert.equal(sh[1].aoa[1][col(sh[1].aoa, '최종입고일')], '2026-09-01');
});
test('여러 시트 파일에서 ERP 시트 고르기 — 실사 수량 칸이 있는 기록서 시트보다 원본 ERP 시트', () => {
  const book = REAL.slice(0, 3).concat([REAL[3].concat(['실물수량', '차이수량', '사유', '특이사항'])], REAL.slice(4).map(r => r.concat([null, null, null, null])));
  const pick = L.pickSheet([{ name: '기록서', aoa: book }, { name: 'ERP', aoa: REAL }, { name: '배정표', aoa: ROSTER }]);
  assert.equal(pick.name, 'ERP');
});
test('엑셀 날짜 일련번호 → 그날 0시(52초 어긋남 없이)', () => {
  assert.equal(L.toDateTime(46316), '2026-10-21T00:00:00');
});

console.log('창고 배정표 (시트 2 모양)');
test('배정표 — 제목 줄 아래 머리행, 배정표 날짜, 실사인원 0 은 제외, 「갑, 을」은 두 사람', () => {
  const r = L.parseRoster(ROSTER);
  assert.equal(r.error, undefined);
  assert.equal(r.headerRow, 5); assert.equal(r.date, '2026-10-21');
  assert.equal(r.rows.length, 4);
  const by = Object.fromEntries(r.rows.map(x => [x.wh, x]));
  assert.equal(by.W30.excluded, true); assert.equal(by.W30.n, 0);
  assert.deepEqual(by.W10.counters, ['실사갑', '실사을']); assert.equal(by.W10.manager, '관리을'); assert.equal(by.W10.site, '사내');
  assert.deepEqual(by.W30.counters, [], '「-」는 이름이 아님');
  assert.match(L.parseRoster([['아무 표'], ['a', 'b']]).error, /머리행/);
});
test('배정표 적용 — 제외 창고 품목은 빼서 따로 남기고, 실사자 2명 창고는 Location 순서로 반씩, 담당자는 창고마다', () => {
  const p = L.parseErp(REAL), ap = L.applyRoster(p.items, L.parseRoster(ROSTER));
  assert.equal(ap.items.length, 5);
  assert.deepEqual(ap.excluded.map(x => x.wh + ':' + x.items + ':' + x.note), ['W30:1:이번 실사 제외, 확인서로 대체']);
  assert.equal(ap.excludedItems.length, 1);
  assert.deepEqual(ap.assign, { 'W10 #1': '실사갑', 'W10 #2': '실사을', W20: '실사병' });
  const w10 = L.sortForWalk(ap.items.filter(x => x.wh === 'W10'));
  assert.deepEqual(w10.map(x => x.part), [1, 1, 2, 2], '보관 자리 순서로 이어서 반씩');
  const s = L.makeSession({ baseAt: '2026-10-21T18:00' }); s.managers = ap.managers;
  assert.deepEqual(w10.map(x => L.assigneeOf(x, s, ap.assign)), ['실사갑', '실사갑', '실사을', '실사을']);
  assert.equal(L.managerOf(w10[0], s), '관리을');
  assert.deepEqual(ap.missing, ['W40 가상창고4'], '배정표에만 있는 창고');
  const ap2 = L.applyRoster(p.items.concat([{ no: 'Z', wh: 'W99', loc: '', sys: 1, price: 1, amt: 1 }]), L.parseRoster(ROSTER));
  assert.deepEqual(ap2.unrostered, [{ wh: 'W99', items: 1 }], 'ERP 에만 있는 창고는 배정 안 됨으로 알림');
});
test('배정표 머리 「실사인원」 — 끝 글자 「원」을 단위로 떼어 못 읽는 일이 없어야 함', () => {
  const r = L.parseRoster(ROSTER);
  assert.ok(r.rows.every(x => x.n != null), r.rows.map(x => x.n).join(','));
});

console.log('작업지시 · 배정');
test('기준 일자·시각이 없으면 작업지시를 만들지 않음(인벤토리 01 의 빠진 항목)', () => {
  assert.throws(() => L.makeSession({ baseAt: '' }), /기준 일자·시각/);
  assert.throws(() => L.makeSession({ baseAt: '2026-10-31' }), /기준 일자·시각/);
  const s = L.makeSession({ baseAt: '2026-10-31T18:00' });
  assert.equal(s.baseAt, '2026-10-31T18:00:00');
  assert.match(s.id, /^ST-202610311800-/);
  assert.equal(s.title, '2026-10 정기 재고실사');
});
test('배정 단위 — 창고별 / 창고·구역별(저장위치 앞자리)', () => {
  const p = parsed();
  assert.deepEqual(L.unitsOf(p.items, 'wh').map(u => u.key + ':' + u.items), ['부품창고:3', '제품창고:1']);
  assert.deepEqual(L.unitsOf(p.items, 'zone').map(u => u.key + ':' + u.items), ['부품창고 / A:2', '부품창고 / B:1', '제품창고 / A:1']);
  assert.equal(L.zoneOf('2F_R12'), '2F');
  assert.equal(L.zoneOf(''), '(위치 없음)');
});
test('자동 배정 — 모든 단위에 담당자, 품목 수가 고르게(가장 큰 단위 하나보다 더 벌어지지 않음)', () => {
  const b = Sample.build(), items = L.parseErp(b.erpAoa).items;
  const units = L.unitsOf(items, 'zone');
  const a = L.autoAssign(units, ['갑', '을', '병']);
  assert.equal(Object.keys(a).length, units.length);
  const load = {}; units.forEach(u => { load[a[u.key]] = (load[a[u.key]] || 0) + u.items; });
  const v = Object.values(load), maxUnit = Math.max(...units.map(u => u.items));
  assert.equal(v.length, 3);
  assert.ok(Math.max(...v) - Math.min(...v) <= maxUnit, JSON.stringify(load));
});
test('기록서 엑셀(지시정보 + 전체 + 실사자별)을 실제 xlsx 로 쓰고 다시 읽으면 회차·기준 시각·품목·배정이 같음', () => {
  const p = parsed(), s = session('zone');
  const assign = { '부품창고 / A': '김', '부품창고 / B': '박', '제품창고 / A': '이' };
  const sheets = L.orderSheets(s, p.items, assign);
  assert.deepEqual(sheets.map(x => x.name), ['지시정보', L.MAIN_SHEET, '김', '박', '이']);
  const sh = toBook(sheets);
  const w = L.parseOrderWorkbook(sh, { fileName: '기록서.xlsx' });
  assert.equal(w.error, undefined);
  assert.equal(w.session.id, 'ST-TEST');
  assert.equal(w.session.baseAt, '2026-10-31T18:00:00');
  assert.equal(w.session.unitMode, 'zone');
  assert.deepEqual(w.items.map(x => x.no), ['S-001', 'S-002', 'S-003', 'S-004']);
  assert.deepEqual(w.assign, assign);
  assert.equal(w.records.length, 0, '아직 실사수량이 빈칸');
});
test('지시정보 시트가 없으면 작업지시서로 받지 않음', () => {
  assert.match(L.parseOrderWorkbook([{ name: 'Sheet1', aoa: ERP }]).error, /지시정보/);
});
test('기록서 양식 — ERP 원래 열 뒤에 담당자 · 실사자 · 실물수량 · 차이수량(수식) · 사유 · 특이사항, 배정표로 나눈 창고·담당자가 다시 읽힘', () => {
  const p = L.parseErp(REAL), ap = L.applyRoster(p.items, L.parseRoster(ROSTER));
  const s = L.makeSession({ id: 'ST-R', baseAt: '2026-10-21T18:00', cols: p.cols });
  s.managers = ap.managers; s.excluded = ap.excluded;
  const sheets = L.orderSheets(s, ap.items, ap.assign);
  assert.deepEqual(sheets.map(x => x.name), ['지시정보', L.MAIN_SHEET, '실사갑', '실사병', '실사을']);
  const main = sheets[1].aoa;
  assert.deepEqual(main[0], ['재고번호', '자재코드', '자재내역', '저장위치', '저장위치명', '창고위치', '단가', '금액', 'Location', '수량', '담당자', '실사자', '실물수량', '차이수량', '사유', '특이사항', '확인자', '확인시각']);
  assert.equal(sheets[1].formula[0].f, 'IF(M2="","",M2-J2)', '차이수량 = 실물수량 − 수량(3차 반영: 모자라면 −)');
  assert.ok(sheets[0].aoa.some(r => r[0] === '이번 실사 제외 창고'));
  const w = L.parseOrderWorkbook(toBook(sheets));
  assert.equal(w.error, undefined);
  assert.deepEqual(w.assign, ap.assign, '한 창고를 두 사람이 나눈 것도 그대로');
  assert.deepEqual(w.session.managers, { W10: '관리을', W20: '관리을' });
  assert.deepEqual(w.session.cols.map(c => c.h).slice(0, 3), ['재고번호', '자재코드', '자재내역']);
  assert.equal(w.items.length, 5);
});

console.log('현장 입력');
test('수량이 같으면 사유 없이 저장, 다르면 사유가 꼭 필요, 「기타」는 특이사항 필요', () => {
  const it = parsed().items[0]; // 시스템 100
  assert.ok(L.validateCount({ qty: '100', by: '김' }, it).ok);
  const v = L.validateCount({ qty: 97, by: '김' }, it);
  assert.ok(!v.ok); assert.match(v.errors.reason, /사유/);
  assert.ok(!L.validateCount({ qty: 97, by: '김', reason: '기타' }, it).ok);
  const ok = L.validateCount({ qty: 97, by: '김', reason: '입출고 누락', note: '불출 미처리', dev: 'd1' }, it, { sessionId: 'ST-TEST' });
  assert.ok(ok.ok); assert.equal(ok.diff, -3);
  assert.equal(ok.record.by, '김'); assert.equal(ok.record.sessionId, 'ST-TEST');
  assert.match(ok.record.at, /^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d$/, '확인 시각 자동');
  assert.match(ok.record.rid, /^d1-/);
});
test('수량 빈칸·음수·글자, 작성자 없음은 막음', () => {
  const it = parsed().items[0];
  assert.ok(L.validateCount({ qty: '', by: '김' }, it).errors.qty);
  assert.ok(L.validateCount({ qty: -1, by: '김' }, it).errors.qty);
  assert.ok(L.validateCount({ qty: '열', by: '김' }, it).errors.qty);
  assert.ok(L.validateCount({ qty: 100 }, it).errors.by);
});
test('목록에 없는 자재 — 창고·이름이 있어야 하고, 재고번호는 NEW-…, 시스템 0 으로 차이에 들어감', () => {
  assert.ok(!L.validateCount({ qty: 3, by: '김', reason: '기타', note: '잔재', extra: { name: '잔재' } }, null).ok);
  const v = L.validateCount({ qty: 3, by: '김', reason: '기타', note: '잔재', extra: { name: '절단 잔재', wh: '부품창고', loc: 'C-1' } }, null);
  assert.ok(v.ok); assert.match(v.record.no, /^NEW-/);
  const rows = L.varianceRows(dbWith([v.record]));
  const r = rows.find(x => x.extra);
  assert.equal(r.sys, 0); assert.equal(r.diff, 3); assert.ok(r.over, '목록에 없던 자재는 늘 확인 대상');
});
test('사유 글자 맞추기 — 별칭·부분 일치, 모르는 글은 「기타」', () => {
  assert.equal(L.matchReason('입고 누락'), '입출고 누락');
  assert.equal(L.matchReason('위치이동'), '위치 이동·분산');
  assert.equal(L.matchReason('라벨 오류'), '라벨·식별 오류');
  assert.equal(L.matchReason('파손'), '파손·폐기 미반영');
  assert.equal(L.matchReason('재실사'), '조사 오류');
  assert.equal(L.matchReason('모름'), '기타');
  assert.equal(L.matchReason(''), '');
});
test('사유 목록을 설정에서 바꿈 — 그 목록 안에서 맞추고, 모르는 글은 「기타」(없으면 마지막), 바꾸기 전 사유도 따로 셈', () => {
  const list = L.reasonList({ reasons: ['입출고 누락', '보관 위치 다름', '불량 격리', '기타', '기타'] });
  assert.deepEqual(list, ['입출고 누락', '보관 위치 다름', '불량 격리', '기타'], '겹친 것은 한 번');
  assert.equal(L.matchReason('입고 누락', list), '입출고 누락');
  assert.equal(L.matchReason('불량격리', list), '불량 격리');
  assert.equal(L.matchReason('모름', list), '기타');
  assert.equal(L.matchReason('모름', ['가', '나']), '나');
  assert.deepEqual(L.reasonList({}), L.REASONS);
  const it = parsed().items[0];
  assert.equal(L.validateCount({ qty: 90, by: '김', reason: '보관 위치 다름' }, it, { reasons: list }).record.reason, '보관 위치 다름');
  const db = dbWith([rec('S-001', 90, '김', 't1', '위치 이동·분산')]); db.settings = L.cleanSettings({ reasons: list });
  const rs = L.reasonStats(L.varianceRows(db), list);
  assert.equal(rs.find(x => x.key === '위치 이동·분산').items, 1);
  assert.ok(rs.some(x => x.key === '불량 격리'));
});
test('기록서 엑셀에 적은 것 — 실물수량 빈 줄은 안 센 것, 확인자 없으면 실사자, 사유 빈 차이는 경고', () => {
  const p = parsed(), s = session();
  const sheets = L.orderSheets(s, p.items, { 부품창고: '김', 제품창고: '이' });
  const main = sheets[1].aoa, c = h => col(main, h); // [머리, S-001, S-002, S-003, S-004] (창고 → 위치 순)
  main[1][c('실물수량')] = 100;                       // S-001 일치
  main[2][c('실물수량')] = 48; main[2][c('사유')] = '위치이동'; main[2][c('확인자')] = '박'; main[2][c('확인시각')] = '2026-11-01 09:30';
  main[4][c('실물수량')] = 9;                         // S-004 차이인데 사유 없음
  const r = L.parseCountSheet(main, { sessionId: s.id, fileName: '종이.xlsx', items: p.items });
  assert.equal(r.records.length, 3); assert.equal(r.blank, 1);
  const by = Object.fromEntries(r.records.map(x => [x.no, x]));
  assert.equal(by['S-001'].by, '김', '확인자 빈칸 → 배정 실사자');
  assert.equal(by['S-002'].reason, '위치 이동·분산');
  assert.equal(by['S-002'].at, '2026-11-01T09:30:00');
  assert.ok(r.warnings.some(w => /사유가 빈/.test(w)));
  const rows = L.varianceRows(dbWith(r.records));
  assert.equal(rows.find(x => x.no === 'S-004').reason, L.NO_REASON);
});

console.log('결과 파일 · 합치기 · 충돌');
test('같은 기록(rid)은 두 번 넣어도 한 번 — 사진 빠진 사본이 사진 있는 기록을 덮지 않음', () => {
  const a = [rec('S-001', 100, '김', '2026-11-01T09:00:00', '', 'r1')];
  const withPhoto = Object.assign({}, a[0], { photo: 'data:image/jpeg;base64,xx' });
  let m = L.mergeRecords([withPhoto], [Object.assign({}, a[0], { photo: 'omitted' })]);
  assert.equal(m.list.length, 1); assert.equal(m.dup, 1); assert.match(m.list[0].photo, /^data:/);
  m = L.mergeRecords(m.list, [rec('S-002', 50, '김', '2026-11-01T09:01:00', '', 'r2')]);
  assert.equal(m.added, 1); assert.equal(m.list.length, 2);
});
test('결과 파일 — 다른 실사 회차 파일은 거절, 지시서에 없는 재고번호는 건너뜀', () => {
  const s = session(), items = parsed().items;
  const f = L.resultFile(s, [rec('S-001', 100, '김', 't1'), rec('X-999', 1, '김', 't2')], { by: '김', dev: 'd1' });
  const c = L.checkResultFile(JSON.parse(JSON.stringify(f)), s, items);
  assert.equal(c.records.length, 1); assert.deepEqual(c.unknown, ['X-999']);
  assert.equal(L.checkResultFile(JSON.parse(JSON.stringify(f)), s, items, '실사결과_김.json').records[0].file, '실사결과_김.json', '받은 파일 이름을 기록에 붙임');
  const other = Object.assign({}, s, { id: 'ST-OTHER' });
  assert.match(L.checkResultFile(f, other, items).error, /다른 실사 회차/);
  assert.match(L.checkResultFile({ foo: 1 }, s, items).error, /결과 파일이 아닙니다/);
});
test('결과 파일 — 사진은 고를 때만 넣음(기본은 「omitted」 표시만)', () => {
  const r = Object.assign(rec('S-001', 100, '김', 't1'), { photo: 'data:image/jpeg;base64,xx' });
  assert.equal(L.resultFile(session(), [r]).records[0].photo, 'omitted');
  assert.match(L.resultFile(session(), [r], { withPhoto: true }).records[0].photo, /^data:/);
});
test('충돌 — 다른 두 사람의 수량이 다르면 충돌, 고르면 해결, 같은 사람이 다시 센 것은 정정', () => {
  const recs = [
    rec('S-001', 95, '김', '2026-11-01T09:00:00', '조사 오류', 'a'),
    rec('S-001', 100, '박', '2026-11-01T09:05:00', '', 'b'),
    rec('S-002', 40, '김', '2026-11-01T09:00:00', '위치 이동·분산', 'c'),
    rec('S-002', 50, '김', '2026-11-01T09:10:00', '', 'd'),
    rec('S-004', 10, '이', '2026-11-01T09:00:00', '', 'e'),
    rec('S-004', 10, '박', '2026-11-01T09:20:00', '', 'f')
  ];
  let r = L.resolveCounts(recs, {});
  assert.ok(r['S-001'].conflict); assert.equal(r['S-001'].current.rid, 'b', '고르기 전에는 늦은 기록');
  assert.ok(!r['S-002'].conflict); assert.ok(r['S-002'].corrected); assert.equal(r['S-002'].current.qty, 50);
  assert.ok(!r['S-004'].conflict, '두 사람이 같은 수량이면 충돌 아님');
  assert.equal(L.conflictList(r).length, 1);
  r = L.resolveCounts(recs, { 'S-001': 'a' });
  assert.ok(!r['S-001'].conflict); assert.ok(r['S-001'].resolved); assert.equal(r['S-001'].current.qty, 95);
});

console.log('SharePoint 취합 파일 (답 ⑤)');
function sharePointBook() {
  const p = L.parseErp(REAL), ap = L.applyRoster(p.items, L.parseRoster(ROSTER));
  const s = L.makeSession({ id: 'ST-SP', baseAt: '2026-10-21T18:00', cols: p.cols }); s.managers = ap.managers;
  return { s, items: ap.items, assign: ap.assign, sheets: L.orderSheets(s, ap.items, ap.assign) };
}
function put(sheet, code, vals) {
  const a = sheet.aoa, i = a.findIndex((r, k) => k && r[col(a, '자재코드')] === code);
  assert.ok(i > 0, code + ' 줄 없음');
  Object.entries(vals).forEach(([h, v]) => { a[i][col(a, h)] = v; });
}
test('여러 사람이 한 파일에 — 자기 시트에 적은 줄 · 전체 시트에 적은 줄을 모두 읽고, 실사자가 확인자가 됨', () => {
  const b = sharePointBook(), sh = Object.fromEntries(b.sheets.map(x => [x.name, x]));
  put(sh['실사갑'], 'X01', { 실물수량: 2, 사유: '조사 오류', 특이사항: '1개 못 찾음' });
  put(sh['실사갑'], 'X02', { 실물수량: 10 });
  put(sh[L.MAIN_SHEET], 'X05', { 실물수량: 10 });                 // 실사병은 전체 시트의 자기 줄에
  put(sh['실사을'], 'X03', { 실물수량: 95, 사유: '입고 누락' });
  const w = L.parseOrderWorkbook(toBook(b.sheets), { fileName: 'sp.xlsx', now: '2026-10-22T10:00:00' });
  assert.equal(w.error, undefined);
  const by = Object.fromEntries(w.records.map(r => [r.no.split('@')[0], r]));
  assert.deepEqual(Object.keys(by).sort(), ['X01', 'X02', 'X03', 'X05']);
  assert.equal(by.X01.by, '실사갑'); assert.equal(by.X05.by, '실사병'); assert.equal(by.X03.reason, '입출고 누락');
  assert.match(by.X05.src, /재고실사기록서 시트/); assert.match(by.X01.src, /실사갑 시트/);
});
test('같은 사람이 전체 시트와 자기 시트에 다르게 적으면 자기 시트 값을 쓰고 알림', () => {
  const b = sharePointBook(), sh = Object.fromEntries(b.sheets.map(x => [x.name, x]));
  put(sh[L.MAIN_SHEET], 'X01', { 실물수량: 1, 사유: '기타', 특이사항: '?' });
  put(sh['실사갑'], 'X01', { 실물수량: 3 });
  const w = L.parseOrderWorkbook(toBook(b.sheets));
  assert.equal(w.records.length, 1); assert.equal(w.records[0].qty, 3);
  assert.ok(w.warnings.some(x => /자기 시트 값을 씁니다/.test(x)), w.warnings.join(' / '));
});
test('같은 SharePoint 파일을 다시 넣으면 0건, 한 줄을 고쳐 다시 넣으면 그 줄만 정정으로 들어감(충돌 아님)', () => {
  const b = sharePointBook(), sh = Object.fromEntries(b.sheets.map(x => [x.name, x]));
  put(sh['실사갑'], 'X01', { 실물수량: 2, 사유: '조사 오류', 특이사항: '1개 못 찾음' });
  put(sh['실사갑'], 'X02', { 실물수량: 10 });
  const first = L.parseOrderWorkbook(toBook(b.sheets), { now: '2026-10-22T10:00:00' });
  let recs = L.mergeRecords([], first.records).list;
  const again = L.mergeRecords(recs, L.parseOrderWorkbook(toBook(b.sheets), { now: '2026-10-22T11:00:00' }).records);
  assert.equal(again.added, 0);
  put(sh['실사갑'], 'X01', { 실물수량: 3, 사유: null, 특이사항: '다시 셈 — 맞음' });
  const fixed = L.mergeRecords(recs, L.parseOrderWorkbook(toBook(b.sheets), { now: '2026-10-22T12:00:00' }).records);
  assert.equal(fixed.added, 1);
  const res = L.resolveCounts(fixed.list, {});
  const x01 = Object.values(res).find(r => r.no.startsWith('X01'));
  assert.ok(!x01.conflict); assert.ok(x01.corrected); assert.equal(x01.current.qty, 3, '나중에 넣은 정정이 채택');
});
test('휴대폰 기록을 기록서 엑셀로 내보냈다가 다시 넣어도 두 번 들어가지 않음(내용이 같으면 같은 기록)', () => {
  const b = sharePointBook();
  const x01 = b.items.find(x => x.code === 'X01');
  const v = L.validateCount({ qty: 2, reason: '조사 오류', note: '못 찾음', by: '실사갑', dev: 'd1', at: '2026-10-21T19:05:31' }, x01, { sessionId: b.s.id });
  const out = L.orderSheets(b.s, b.items, b.assign, { records: [v.record], only: '실사갑' });
  assert.deepEqual(out.map(x => x.name), ['지시정보', L.MAIN_SHEET], '내 줄만 담은 결과 엑셀');
  assert.equal(out[1].aoa.length - 1, 2, '실사갑 배정 2줄');
  const back = L.parseOrderWorkbook(toBook(out));
  assert.equal(back.records.length, 1);
  assert.equal(L.mergeRecords([v.record], back.records).added, 0);
  const row = out[1].aoa.find((r, i) => i && r[col(out[1].aoa, '자재코드')] === 'X01');
  assert.equal(row[col(out[1].aoa, '차이수량')], -1, '실물 2 − 수량 3 = −1 (모자라면 −)');
});

console.log('차이대사');
test('허용 기준 기본값 = 허용 0 — 차이가 있으면 1개라도 확인 대상(답 ③)', () => {
  assert.equal(L.cleanSettings({}).tolMode, 'zero');
  const rows = Object.fromEntries(L.varianceRows(dbWith([rec('S-001', 99, '김', 't1', '조사 오류'), rec('S-002', 50, '김', 't2')])).map(r => [r.no, r]));
  assert.ok(rows['S-001'].over); assert.equal(rows['S-001'].status, '차이(사유 필수)');
  assert.equal(rows['S-002'].status, '일치'); assert.ok(!rows['S-002'].over);
  assert.match(L.tolText({}), /허용 0/);
  assert.equal(L.restoreDb({ settings: { tolPct: 2, tolAmt: 100000 } }).settings.tolMode, 'zero', '1단계 백업도 허용 0 으로');
});
test('허용 범위를 고르면(선택) 1단계처럼 차이율 또는 금액으로 판정', () => {
  const db = dbWith([
    rec('S-001', 99, '김', 't1', '조사 오류'),          // -1, -1,000원, -1%  → 허용 안
    rec('S-002', 40, '김', 't2', '위치 이동·분산'),     // -10, -5,000원, -20% → 초과(차이율)
    rec('S-004', 11, '이', 't3', '입출고 누락')         // +1, +500,000원, +10% → 초과
  ]);
  db.settings = L.cleanSettings({ tolMode: 'range' });
  const rows = Object.fromEntries(L.varianceRows(db).map(r => [r.no, r]));
  assert.equal(rows['S-001'].diff, -1); assert.equal(rows['S-001'].diffAmt, -1000); near(rows['S-001'].rate, -0.01);
  assert.equal(rows['S-001'].status, '허용 안 차이');
  assert.equal(rows['S-002'].status, '허용 초과');
  assert.equal(rows['S-004'].diffAmt, 500000); assert.ok(rows['S-004'].over);
  assert.equal(rows['S-003'].status, '미실사');
  db.settings = L.cleanSettings({ tolMode: 'range', tolPct: 25, tolAmt: 1000000 });
  const r2 = Object.fromEntries(L.varianceRows(db).map(r => [r.no, r]));
  assert.ok(!r2['S-002'].over && !r2['S-004'].over, '기준을 넓히면 허용 안');
});
test('시스템 0 인 품목에서 나온 수량은 차이율 대신 「신규」, 확인 대상', () => {
  const db = dbWith([rec('S-003', 5, '김', 't', '입출고 누락')]); db.settings = L.cleanSettings({ tolMode: 'range' });
  const rows = Object.fromEntries(L.varianceRows(db).map(r => [r.no, r]));
  assert.equal(rows['S-003'].rate, null); assert.ok(rows['S-003'].over);
});

console.log('집계 (BI)');
test('진행률·정확도·증가/감소/순차이/절대 차이 — 전체와 창고별·담당자별', () => {
  const db = dbWith([rec('S-001', 100, '김', 't1'), rec('S-002', 40, '김', 't2', '위치 이동·분산'), rec('S-004', 11, '이', 't3', '입출고 누락')]);
  const an = L.analyze(db), a = an.all;
  assert.equal(a.items, 4); assert.equal(a.counted, 3); near(a.progress, 0.75);
  assert.equal(a.matched, 1); near(a.accuracy, 1 / 3);
  assert.equal(a.plusAmt, 500000); assert.equal(a.minusAmt, -5000); assert.equal(a.netAmt, 495000); assert.equal(a.absAmt, 505000);
  const wh = Object.fromEntries(an.dims.wh.map(x => [x.key, x]));
  near(wh['부품창고'].progress, 2 / 3); assert.equal(wh['제품창고'].diffItems, 1);
  const as = Object.fromEntries(an.dims.assignee.map(x => [x.key, x]));
  assert.equal(as['김'].items, 3); assert.equal(as['김'].counted, 2);
});
test('사유별 — 차이 품목만, 비중 합 100%, 사유 없는 차이는 「사유 미입력」', () => {
  const db = dbWith([rec('S-001', 90, '김', 't1', '입출고 누락'), rec('S-002', 40, '김', 't2', ''), rec('S-004', 10, '이', 't3', '')]);
  const rs = Object.fromEntries(L.reasonStats(L.varianceRows(db)).map(x => [x.key, x]));
  assert.equal(rs['입출고 누락'].items, 1); assert.equal(rs[L.NO_REASON].items, 1);
  near(Object.values(rs).reduce((s, x) => s + x.share, 0), 1);
  near(Object.values(rs).reduce((s, x) => s + x.amtShare, 0), 1);
  assert.equal(Object.keys(rs).length, 7, '여섯 사유는 0 이어도 늘 보임 + 사유 미입력');
});
test('차이 금액 상위 — 절대 금액 큰 순', () => {
  const db = dbWith([rec('S-001', 90, '김', 't1', '입출고 누락'), rec('S-002', 40, '김', 't2', '위치 이동·분산'), rec('S-004', 11, '이', 't3', '입출고 누락')]);
  assert.deepEqual(L.topDiffs(L.varianceRows(db), 3).map(r => r.no), ['S-004', 'S-001', 'S-002']);
});

console.log('보고서');
test('발행할 때마다 버전이 오르고 기준 시각·기록 수가 남음', () => {
  const db = dbWith([rec('S-001', 100, '김', 't1')]);
  const m1 = L.nextReportMeta(db, { now: '2026-11-02T10:00:00', by: '이미란' });
  assert.equal(m1.ver, 1); assert.equal(m1.baseAt, '2026-10-31T18:00:00'); assert.equal(m1.records, 1);
  db.reports.push(m1);
  assert.equal(L.nextReportMeta(db).ver, 2);
  assert.equal(L.verText(m1), 'v1 (2026-11-02 10:00 발행)');
});
test('결과서 엑셀 시트 — 버전·기준 시각, 원천 기록 연결(기록 ID → 기록 이력), 이전 기록도 남김', () => {
  const recs = [rec('S-002', 40, '김', '2026-11-01T09:00:00', '위치 이동·분산', 'r-old'), rec('S-002', 45, '김', '2026-11-01T09:10:00', '위치 이동·분산', 'r-new'), rec('S-001', 100, '김', '2026-11-01T09:20:00', '', 'r-1')];
  const db = dbWith(recs);
  const meta = L.nextReportMeta(db, { now: '2026-11-02T10:00:00' });
  const sh = Object.fromEntries(L.reportSheets(db, null, meta).map(s => [s.name, s.aoa]));
  assert.deepEqual(Object.keys(sh), ['보고서 정보', '창고별', '자재유형별', '자재타입별', '실사자별', '사유별', '차이 상세', '전체 대사표', '기록 이력'], '파일에 없는 창고위치는 시트를 만들지 않음');
  assert.ok(sh['보고서 정보'].some(r => r[0] === '허용 기준' && /허용 0/.test(r[1])));
  assert.equal(sh['보고서 정보'][1][1], 'v1 (2026-11-02 10:00 발행)');
  assert.ok(sh['보고서 정보'].some(r => r[0] === '기준 일자·시각(ERP 재고)' && r[1] === '2026-10-31 18:00'));
  const head = sh['차이 상세'][0], row = sh['차이 상세'][1];
  assert.equal(row[head.indexOf('기록 ID')], 'r-new');
  assert.equal(row[head.indexOf('충돌·정정')], '같은 사람이 다시 셈');
  const hist = sh['기록 이력'];
  assert.equal(hist.length, 4);
  assert.deepEqual(hist.slice(1).map(r => r[0] + ':' + r[10]), ['r-old:이전 기록', 'r-new:채택', 'r-1:채택']);
});
test('AI 프롬프트 — 실사자 이름은 별칭(되돌리기 가능), 자재명·자재코드는 기본으로 넣지 않음', () => {
  const b = Sample.build();
  const db = L.restoreDb({ session: b.session, items: b.items, assign: b.assign, records: b.records });
  const an = L.analyze(db);
  const info = L.buildAiPromptInfo(an);
  b.names.slice(0, 4).forEach(n => assert.ok(!info.prompt.includes(n), n + ' 가 프롬프트에 있음'));
  assert.deepEqual(info.leaks, []);
  an.top.slice(0, 10).forEach(r => { assert.ok(!info.prompt.includes(r.name), r.name); assert.ok(!info.prompt.includes(r.code), r.code); });
  assert.ok(info.prompt.includes('담당1'));
  const back = L.unmaskText('담당1 과 담당2 확인', info.map);
  assert.ok(!back.includes('담당1') && b.names.some(n => back.includes(n)));
  const open = L.buildAiPromptInfo(an, { mask: false, withItem: true });
  assert.ok(open.prompt.includes(an.top[0].name));
});

console.log('가상 샘플 (실제 양식과 같은 모양 · 심어 둔 것을 찾는지)');
const S = Sample.build();
const sdb = L.restoreDb({ session: S.session, items: S.items, assign: S.assign, records: S.records });
test('ERP 샘플은 견본과 같은 머리행(+ 추가 열 2개) 248품목, 배정표 적용 → 시험실창고 8품목 제외 · 240품목 · 부품창고는 실사자 2명이 반씩', () => {
  assert.deepEqual(S.erpAoa[3].slice(0, 9), ['자재코드', '자재내역', '저장위치', '저장위치명', '창고위치', '단가', '금액', 'Location', '수량']);
  assert.deepEqual(S.rosterAoa[5], ['저장위치', '저장위치명', '창고위치', '실사인원', '담당팀', '담당자', '실사자', '비고']);
  const p = L.parseErp(S.erpAoa);
  assert.equal(p.items.length, 248); assert.deepEqual(p.det.missing, []);
  assert.equal(S.items.length, 240);
  assert.deepEqual(S.session.excluded.map(x => x.whName + ':' + x.items), ['시험실창고:8']);
  const units = L.unitsOf(S.items, 'wh');
  units.forEach(u => assert.ok(S.assign[u.key], u.key));
  const part = units.filter(u => u.wh === 'S220').map(u => u.key + ':' + u.items);
  assert.deepEqual(part, ['S220 #1:35', 'S220 #2:35']);
});
test('충돌 2건(부품창고 C 구역), 외주창고 진행률 가장 낮음, 실사자E 는 합치기 전 0%, 창고위치(사내/사외)별 집계', () => {
  const an = L.analyze(sdb);
  assert.equal(an.conflicts.length, 2);
  an.conflicts.forEach(r => assert.equal(r.zone, 'S220 부품창고 / C'));
  const wh = an.dims.wh.filter(x => !/소모품/.test(x.key)).sort((a, b) => a.progress - b.progress);
  assert.equal(wh[0].key, 'S310 외주창고'); assert.ok(wh[0].progress < 0.6);
  assert.equal(an.dims.assignee.find(x => x.key === '실사자E').counted, 0);
  assert.deepEqual(an.dims.site.map(x => x.key).sort(), ['사내', '사외']);
  assert.ok(an.hasDim.site && an.hasDim.mtype && an.hasDim.mkind);
});
test('SharePoint 취합 파일 합치기 → 소모품창고 다 셈, 「입출고 누락」이 소모품창고·저장품에 몰림, 사유 미입력 1건, 외주창고 +2', () => {
  const sheets = Sample.fillPaper(L.orderSheets(sdb.session, sdb.items, sdb.assign), S.heldOut, S.mainRows);
  const w = L.parseOrderWorkbook(toBook(sheets), { fileName: S.eFile });
  assert.equal(w.error, undefined); assert.equal(w.records.length, 42);
  const db2 = L.restoreDb(JSON.parse(JSON.stringify(sdb)));
  db2.records = L.mergeRecords(db2.records, w.records).list;
  const again = L.mergeRecords(db2.records, w.records);
  assert.equal(again.added, 0, '같은 파일을 또 넣어도 늘지 않음');
  const an = L.analyze(db2);
  assert.equal(an.dims.wh.find(x => /소모품/.test(x.key)).progress, 1);
  assert.equal(an.dims.wh.find(x => /외주/.test(x.key)).counted, 21);
  const io = an.rows.filter(r => r.reason === '입출고 누락');
  const sp2 = io.filter(r => /소모품/.test(r.wh)).length;
  assert.ok(sp2 / io.length > 0.6, '입출고 누락 ' + sp2 + '/' + io.length);
  const mk = an.dims.mkind.find(x => x.key === '저장품');
  assert.ok(mk.diffItems >= 10 && mk.netAmt < 0);
  assert.equal(an.reasons.find(x => x.key === L.NO_REASON).items, 1);
});
test('부품창고 B 구역에 「위치 이동·분산」이 몰림, 제품창고 라벨 짝은 수량이 서로 맞물림', () => {
  const rows = L.varianceRows(sdb);
  const mv = rows.filter(r => r.reason === '위치 이동·분산');
  assert.ok(mv.length >= 6 && mv.every(r => r.zone === 'S220 부품창고 / B'), mv.map(r => r.zone).join(','));
  const lb = rows.filter(r => r.reason === '라벨·식별 오류');
  assert.equal(lb.length, 6); assert.equal(lb.reduce((s, r) => s + r.diff, 0), 0);
});
test('같은 사람이 다시 센 정정 2건은 충돌이 아니고 마지막 기록(일치)을 씀, 잔재 2건은 목록에 없는 자재로', () => {
  const rows = L.varianceRows(sdb);
  const cor = rows.filter(r => r.corrected && !r.conflict);
  assert.equal(cor.length, 2); cor.forEach(r => assert.equal(r.diff, 0));
  const ex = rows.filter(r => r.extra);
  assert.equal(ex.length, 2); ex.forEach(r => assert.equal(r.wh, 'S210 소재창고'));
});

console.log('저장 · AI 연결');
test('백업 되살리기 — 모르는 값은 버리고 빈 db 는 기본값', () => {
  const d = L.restoreDb(JSON.parse(JSON.stringify(sdb)));
  assert.equal(d.items.length, 240); assert.equal(d.records.length, S.records.length);
  assert.equal(L.restoreDb({ excludedItems: [{ no: 'a' }, {}] }).excludedItems.length, 1);
  const e = L.restoreDb({ records: [{ foo: 1 }], settings: { tolPct: -3 } });
  assert.equal(e.records.length, 0); assert.equal(e.settings.tolPct, 2);
});
test('AI 연결 설정 — data10-14 저장 칸, 키가 없으면 인증 머리글 없음', () => {
  assert.equal(E.STORE_KEY, 'data10-14.ai');
  const r = E.buildChatRequest({ baseUrl: 'https://llm.example.com/v1/', model: 'm' }, E.promptMessages('x'));
  assert.equal(r.url, 'https://llm.example.com/v1/chat/completions');
  assert.equal(r.init.headers.Authorization, undefined);
});

const { versions } = await import('../scripts/bump-version.mjs');
test('판 맞추기 — index.html ?v= · logic.js · app.js 가 같은 판 (scripts/bump-version.mjs)', () => {
  const v = versions();
  assert.equal(v.logic, L.VERSION, 'logic.js VERSION'); assert.equal(v.app, L.VERSION, 'app.js APP_VERSION');
  assert.ok(v.html.length >= 11 && v.html.every(x => x === L.VERSION), 'index.html 의 모든 js·css 에 ?v=' + L.VERSION + ' — ' + JSON.stringify(v.html));
});

console.log('\n' + passed + '개 통과' + (process.exitCode ? ' — 실패가 있습니다' : ''));
