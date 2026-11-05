/* 가상 샘플 — 창고·자재·사람·수량·단가 모두 지어낸 값입니다.
   2차 반영(2026-10-21): 수강생이 보낸 실제 양식과 같은 모양으로 다시 만들었습니다(값은 전부 가짜).
     · ERP 재고 엑셀 — 첫 줄 안내 · 빈 줄 2개 · 머리행: 자재코드 · 자재내역 · 저장위치 · 저장위치명 · 창고위치 · 단가 · 금액 · Location · 수량
       (견본에 「필드가 추가될 수 있음」이라 되어 있어, 추가 열 예시로 자재유형 · 자재타입을 뒤에 붙였습니다)
     · 창고 배정표 — 제목 줄 · 배정표 날짜 · 머리행: 저장위치 · 저장위치명 · 창고위치 · 실사인원 · 담당팀 · 담당자 · 실사자 · 비고
   브라우저에서는 window.STSample(먼저 js/logic.js), node 에서는 require('./sample-data.js').

   ERP 재고 248품목(창고 6곳, 그중 1곳은 이번 실사 제외) · 기준 2026-09-30 18:00 · 실사 2026-10-01. 일부러 심어 둔 것:
   ① 소모품창고 저장품은 불출하고 출고 처리를 안 한 「입출고 누락」 감소가 몰림(실사자E 가 SharePoint 취합 파일에 적음 — ③ 취합에서 「합치기」)
   ② 부품창고 B 구역은 「위치 이동·분산」(다른 칸에서 나옴 · 없음)이 몰림. 부품창고는 실사자 2명이라 보관 자리 순서로 반씩 나눔
   ③ 제품창고는 라벨이 바뀐 짝 3쌍(+n / -n, 「라벨·식별 오류」) — 수량은 서로 맞물리지만 단가가 달라 금액이 크게 흔들림
   ④ 소재창고는 「파손·폐기 미반영」 감소 3건 + 목록에 없는 잔재 2건 발견
   ⑤ 부품창고 C 구역 2품목은 두 사람이 따로 세어 수량이 다름(충돌) · 2품목은 같은 사람이 다시 세어 정정
   ⑥ 외주창고(사외)는 절반 조금 넘게만 셈(진행률 낮음) — SharePoint 파일 전체 시트에 2줄을 더 적음
   ⑦ 시험실창고는 배정표에서 실사인원 0 → 이번 실사 제외(담당자 확인서로 대신) */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./logic.js'));
  else root.STSample = factory(root.STLogic);
})(typeof self !== 'undefined' ? self : this, function (L) {
  'use strict';

  var SESSION_ID = 'ST-202609301800-SMPL';
  var BASE_AT = '2026-09-30T18:00';
  var ERP_FILE = 'ERP_재고현황_다운로드_샘플_2026-09.xlsx';
  var ROSTER_FILE = '창고배정표_샘플_2026-09.xlsx';
  var E_FILE = '재고실사기록서_SharePoint취합_샘플.xlsx';
  var ERP_HEAD = ['자재코드', '자재내역', '저장위치', '저장위치명', '창고위치', '단가', '금액', 'Location', '수량', '자재유형', '자재타입'];
  var ROSTER_HEAD = ['저장위치', '저장위치명', '창고위치', '실사인원', '담당팀', '담당자', '실사자', '비고'];

  function rng(seed) { return function () { seed |= 0; seed = seed + 0x6D2B79F5 | 0; var t = Math.imul(seed ^ seed >>> 15, 1 | seed); t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0) / 4294967296; }; }

  var TYPES = {
    '강재': { kind: '원자재', q: [300, 3000], p: [1200, 2600], names: ['강판 SS400', '환봉 S45C', '각파이프 STKR', '후판 SM490'] },
    '주물': { kind: '원자재', q: [20, 300], p: [15000, 80000], names: ['하우징 주물', '커버 주물', '블록 주물'] },
    '봉재': { kind: '원자재', q: [50, 600], p: [3000, 25000], names: ['크롬도금봉', '피스톤봉 소재'] },
    '유압부품': { kind: '부자재', q: [10, 200], p: [20000, 350000], names: ['유압 밸브', '체크 밸브', '릴리프 밸브', '유압 호스', '피팅'] },
    '실·패킹': { kind: '부자재', q: [100, 800], p: [800, 15000], names: ['O-링', '더스트 실', '피스톤 실', '백업 링'] },
    '볼트·너트': { kind: '부자재', q: [200, 800], p: [50, 900], names: ['육각 볼트', '플랜지 너트', '스터드 볼트', '와셔'] },
    '베어링': { kind: '부자재', q: [10, 150], p: [8000, 60000], names: ['볼 베어링', '니들 베어링', '부싱'] },
    '완제품': { kind: '제품', q: [5, 60], p: [800000, 4000000], names: ['유압 실린더', '유압 펌프', '주행 모터'] },
    '반제품': { kind: '반제품', q: [10, 120], p: [120000, 900000], names: ['실린더 튜브 가공품', '로드 가공품', '헤드 조립품'] },
    '공구': { kind: '저장품', q: [20, 200], p: [5000, 90000], names: ['엔드밀', '드릴', '탭', '인서트 팁'] },
    '소모품': { kind: '저장품', q: [50, 500], p: [500, 20000], names: ['장갑', '연마지', '절삭유', '방청제', '용접봉'] },
    '포장재': { kind: '포장재', q: [20, 300], p: [1500, 8000], names: ['포장 박스', '완충재', '팔레트 랩'] },
    '외주가공품': { kind: '반제품', q: [10, 300], p: [30000, 300000], names: ['열처리품', '도금품', '정밀 가공품'] },
    '시험품': { kind: '저장품', q: [1, 20], p: [10000, 200000], names: ['시험용 시편', '시제품 부품'] }
  };
  /* 창고 — 저장위치(코드) · 저장위치명 · 창고위치(사내/사외) */
  var WH = [
    { code: 'S210', name: '소재창고', site: '사내', zones: ['A', 'B'], n: 50, types: ['강재', '주물', '봉재'] },
    { code: 'S220', name: '부품창고', site: '사내', zones: ['A', 'B', 'C'], n: 70, types: ['유압부품', '실·패킹', '볼트·너트', '베어링'] },
    { code: 'S230', name: '제품창고', site: '사내', zones: ['A', 'B'], n: 45, types: ['완제품', '반제품'] },
    { code: 'S240', name: '소모품창고', site: '사내', zones: ['A'], n: 40, types: ['공구', '소모품', '포장재'] },
    { code: 'S250', name: '시험실창고', site: '사내', zones: ['A'], n: 8, types: ['시험품'] },
    { code: 'S310', name: '외주창고', site: '사외', zones: ['A'], n: 35, types: ['외주가공품', '주물'] }
  ];
  /* 창고 배정표 — 수강생 시트 2 와 같은 모양(사람 이름은 가상) */
  var ROSTER = [
    ['S250', '시험실창고', '사내', 0, '개발팀', '관리자4', '-', '이번 실사 제외 — 담당자 확인서로 대신함'],
    ['S310', '외주창고', '사외', 1, '구매팀', '관리자3', '실사자C', null],
    ['S210', '소재창고', '사내', 1, '물류팀', '관리자1', '실사자A', null],
    ['S220', '부품창고', '사내', 2, '물류팀', '관리자1', '실사자B, 실사자C', null],
    ['S230', '제품창고', '사내', 1, '물류팀', '관리자2', '실사자D', null],
    ['S240', '소모품창고', '사내', 1, '물류팀', '관리자2', '실사자E', '종이에 적고 SharePoint 파일로 옮김']
  ];
  var NAMES = ['실사자A', '실사자B', '실사자C', '실사자D', '실사자E'];

  function between(r, a, b) { return a + (b - a) * r(); }
  function int(r, a, b) { return Math.round(between(r, a, b)); }
  function pad(n, w) { var s = String(n); while (s.length < w) s = '0' + s; return s; }

  function makeItems() {
    var r = rng(1410), items = [];
    WH.forEach(function (w, wi) {
      for (var i = 0; i < w.n; i++) {
        var mtype = w.types[i % w.types.length], T = TYPES[mtype];
        var zone = w.zones[i % w.zones.length];
        var nm = T.names[int(r, 0, T.names.length - 1)];
        var size = mtype === '강재' ? int(r, 3, 30) + 'T' : mtype === '볼트·너트' ? 'M' + [8, 10, 12, 16, 20][int(r, 0, 4)] : int(r, 20, 160) + 'A';
        var price = mtype === '볼트·너트' || mtype === '실·패킹' || mtype === '소모품' ? int(r, T.p[0], T.p[1]) : Math.round(between(r, T.p[0], T.p[1]) / 100) * 100;
        var qty = int(r, T.q[0], T.q[1]);
        if (((wi * 100 + i) % 53) === 0) qty = 0; // 재고 0 인 품목도 내려받음
        items.push({
          code: 'M' + pad(100000 + wi * 20000 + i * 7 + int(r, 0, 6), 6), name: nm + ' ' + size, wh: w.code, whName: w.name, site: w.site,
          loc: zone + '-' + pad(int(r, 1, 12), 2) + '-' + int(r, 1, 4), mtype: mtype, mkind: T.kind, sys: qty, price: price
        });
      }
    });
    return items;
  }
  function erpAoa(items) {
    var aoa = [['가상 샘플 — 재고 현황 조회(ERP 내려받기 모양). 필드가 추가될 수 있음(이 샘플은 자재유형 · 자재타입을 덧붙인 예).'], [], [], ERP_HEAD.slice()];
    items.forEach(function (it) { aoa.push([it.code, it.name, it.wh, it.whName, it.site, it.price, it.sys * it.price, it.loc, it.sys, it.mtype, it.mkind]); });
    return aoa;
  }
  function rosterAoa() {
    return [['작업배정 시 활용하는 자료(가상 샘플)'], ['창고 목록과 창고별 담당 인원'], [], ['2026-09-30', '재고실사 배정표'], [], ROSTER_HEAD.slice()].concat(ROSTER.map(function (r) { return r.slice(); }));
  }

  /* 실사 회차 · 배정 · 현장 기록 만들기 */
  function build() {
    var raw = makeItems(), r = rng(2026);
    var eAoa = erpAoa(raw), rAoa = rosterAoa();
    var p = L.parseErp(eAoa), ro = L.parseRoster(rAoa), ap = L.applyRoster(p.items, ro);
    var session = L.makeSession({ id: SESSION_ID, baseAt: BASE_AT, title: '2026년 9월 정기 재고실사 (가상 샘플)', sourceFile: ERP_FILE, unitMode: 'wh', cols: p.cols, now: '2026-09-30T18:20:00' });
    session.managers = ap.managers; session.excluded = ap.excluded; session.rosterFile = ROSTER_FILE;
    var items = ap.items;
    function who(it) { return L.assigneeOf(it, session, ap.assign); }
    var clock = {};
    NAMES.forEach(function (n, i) { clock[n] = Date.UTC(2026, 9, 1, 8, 30 + i * 2, 0); });
    function tick(n) { clock[n] += int(r, 60, 230) * 1000; var d = new Date(clock[n]); return d.toISOString().slice(0, 19); }
    var seqs = {};
    function rid(n) { seqs[n] = (seqs[n] || 0) + 1; return 'smp' + n.slice(-1) + '-' + pad(seqs[n], 4); }
    function rec(n, it, qty, reason, note, at) {
      return { rid: rid(n), no: it ? it.no : null, qty: qty, reason: reason || '', note: note || '', by: n, at: at || tick(n), dev: 'dev' + n.slice(-1), src: '현장 입력(' + n + ' 휴대폰)', file: '실사결과_' + n + '.json', photo: null, sessionId: SESSION_ID };
    }
    var records = [], heldOut = [], mainRows = [];
    var byWh = {}; items.forEach(function (it) { (byWh[it.whName] = byWh[it.whName] || []).push(it); });

    // 소재창고 — 실사자A: 파손·폐기 3건, 다시 세어 정정 1건, 잔재 2건 발견
    byWh['소재창고'].forEach(function (it, i) {
      if (i === 4 || i === 17 || i === 29) { records.push(rec('실사자A', it, Math.round(it.sys * 0.82), '파손·폐기 미반영', '녹 발생분 폐기 대기 구역에 따로 둠')); return; }
      if (i === 11) {
        records.push(rec('실사자A', it, it.sys - 40, '위치 이동·분산', ''));
        records.push(rec('실사자A', it, it.sys, '', '옆 랙에서 나머지 확인 — 다시 셈'));
        return;
      }
      records.push(rec('실사자A', it, it.sys, '', ''));
    });
    var a1 = rec('실사자A', null, 120, '기타', '미등록 잔재(강판 절단 잔재)'); a1.no = 'NEW-잔재SS40012T@S210';
    a1.extra = { code: '', name: '강판 SS400 12T 절단 잔재', wh: 'S210', loc: 'B-12-1', unit: 'KG' };
    var a2 = rec('실사자A', null, 6, '기타', '미등록 반납품(현장 반납 후 입고 처리 안 됨)'); a2.no = 'NEW-하우징주물반납품@S210';
    a2.extra = { code: '', name: '하우징 주물 반납품', wh: 'S210', loc: 'A-03-2', unit: 'EA' };
    records.push(a1, a2);

    // 부품창고 — 실사자B·C 가 보관 자리 순서로 반씩(배정표 실사인원 2). C 구역 충돌 2건은 실사자B 가 도와 따로 셈
    var conflictsDone = 0, fixDone = 0;
    L.sortForWalk(byWh['부품창고']).forEach(function (it, i) {
      var zone = it.loc.charAt(0), n = who(it);
      if (zone === 'C' && conflictsDone < 2 && i % 3 === 2 && i > 50) {
        conflictsDone++;
        records.push(rec(n, it, it.sys - 5, '조사 오류', '박스 단위로 셈'));
        records.push(rec('실사자B', it, it.sys, '', '낱개로 다시 셈'));
        return;
      }
      if (zone === 'B' && r() < 0.42) {
        var d = Math.max(1, Math.round(it.sys * between(r, 0.05, 0.25)));
        var plus = r() < 0.35;
        records.push(rec(n, it, plus ? it.sys + d : it.sys - d, '위치 이동·분산', plus ? '다른 칸에 있던 물량 발견' : '일부가 다른 칸으로 옮겨진 것으로 보임'));
        return;
      }
      if (zone === 'A' && i % 12 === 3 && it.sys >= 100) { records.push(rec(n, it, it.sys - 1, '입출고 누락', '소량 불출 출고 처리 전')); return; }
      if (zone === 'A' && !fixDone && i > 4) {
        fixDone++;
        records.push(rec(n, it, it.sys + 12, '조사 오류', ''));
        records.push(rec(n, it, it.sys, '', '다시 셈'));
        return;
      }
      records.push(rec(n, it, it.sys, '', ''));
    });

    // 제품창고 — 실사자D: 라벨이 바뀐 짝 3쌍
    var fg = byWh['제품창고'], pairs = [[2, 5], [12, 21], [30, 33]], pairOf = {};
    pairs.forEach(function (pp, k) { pairOf[pp[0]] = { other: pp[1], sign: 1, k: k }; pairOf[pp[1]] = { other: pp[0], sign: -1, k: k }; });
    fg.forEach(function (it, i) {
      var pr = pairOf[i];
      if (pr) {
        var n = [3, 2, 4][pr.k];
        records.push(rec('실사자D', it, Math.max(0, it.sys + pr.sign * n), '라벨·식별 오류', '자재코드 ' + fg[pr.other].code + ' 와 라벨이 바뀐 것으로 보임'));
        return;
      }
      if (i === 40) { records.push(rec('실사자D', it, it.sys - 1, '입출고 누락', '9/30 야간 출하분 출고 처리 전')); return; }
      records.push(rec('실사자D', it, it.sys, '', ''));
    });

    // 외주창고(사외) — 실사자C: 휴대폰으로 19품목, SharePoint 파일 전체 시트에 2품목 더(합치기 전에는 안 보임)
    byWh['외주창고'].forEach(function (it, i) {
      if (i === 19 || i === 20) { mainRows.push({ no: it.no, qty: it.sys, reason: '', note: '', by: '실사자C', at: tick('실사자C') }); return; }
      if (i >= 19) return;
      if (i === 6) { records.push(rec('실사자C', it, it.sys - 8, '입출고 누락', '외주 입고 전표 미처리')); return; }
      records.push(rec('실사자C', it, it.sys, '', ''));
    });

    // 소모품창고 — 실사자E 는 종이에 적고 SharePoint 취합 파일의 자기 시트로 옮김(「합치기」로 들어옴)
    byWh['소모품창고'].forEach(function (it, i) {
      var at = tick('실사자E');
      var o = { no: it.no, qty: it.sys, reason: '', note: '', by: '실사자E', at: at };
      if (it.mkind === '저장품' && r() < 0.48 && it.sys > 0) {
        o.qty = it.sys - Math.max(1, Math.round(it.sys * between(r, 0.04, 0.2)));
        o.reason = '입출고 누락'; o.note = '현장 불출 후 출고 처리 안 됨';
      } else if (it.mtype === '포장재' && i % 9 === 2) {
        o.qty = Math.max(0, it.sys - int(r, 2, 6)); o.reason = '파손·폐기 미반영'; o.note = '젖은 박스 폐기';
      }
      if (i === 13 && o.qty === it.sys) { o.qty = it.sys - 3; o.reason = ''; o.note = ''; } // 사유를 안 적은 줄
      heldOut.push(o);
    });

    return {
      sessionId: SESSION_ID, baseAt: BASE_AT, erpFile: ERP_FILE, rosterFile: ROSTER_FILE, eFile: E_FILE, title: session.title,
      raw: raw, erpAoa: eAoa, rosterAoa: rAoa, session: session, items: items, assign: ap.assign, roster: ro, applied: ap,
      names: NAMES, records: records, heldOut: heldOut, mainRows: mainRows,
      files: ['실사자A', '실사자B', '실사자C', '실사자D'].map(function (n, k) { return { name: '실사결과_' + n + '.json', records: records.filter(function (x) { return x.by === n; }).length, at: '2026-10-01T16:' + pad(10 + k * 5, 2) + ':00', kind: '결과 파일' }; })
    };
  }

  /* 기록서 시트(L.orderSheets 결과)에 SharePoint 에서 채운 값을 적습니다 — samples/ 파일과 「합치기 체험」 버튼이 같이 씀.
     · 실사자E 는 자기 이름 시트에 40줄 · 실사자C 는 「재고실사기록서」(전체) 시트의 자기 줄에 2줄 */
  function fillPaper(sheets, heldOut, mainRows) {
    function fill(s, list) {
      var by = {}; list.forEach(function (o) { by[o.no] = o; });
      var hd = s.aoa[0], c = function (h) { return hd.indexOf(h); };
      s.aoa.forEach(function (row, i) {
        var o = i ? by[row[c('재고번호')]] : null; if (!o) return;
        var sys = row[c('수량')];
        row[c('실물수량')] = o.qty; row[c('차이수량')] = o.qty - sys; row[c('사유')] = o.reason || null; row[c('특이사항')] = o.note || null;
        row[c('확인자')] = o.by; row[c('확인시각')] = o.at.slice(0, 16).replace('T', ' ');
        (s.formula || []).forEach(function (f) { if (f.a === L.colName(c('차이수량')) + (i + 1)) f.v = o.qty - sys; });
      });
    }
    sheets.forEach(function (s) {
      if (s.name === '실사자E') fill(s, heldOut);
      if (s.name === L.MAIN_SHEET) fill(s, mainRows || []);
    });
    return sheets;
  }

  return { build: build, fillPaper: fillPaper, SESSION_ID: SESSION_ID, BASE_AT: BASE_AT, ERP_FILE: ERP_FILE, ROSTER_FILE: ROSTER_FILE, E_FILE: E_FILE, ERP_HEAD: ERP_HEAD, ROSTER_HEAD: ROSTER_HEAD, NAMES: NAMES };
});
