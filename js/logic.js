/* 재고실사 Start-End — 순수 로직 (화면·저장소와 무관)
   브라우저에서는 window.STLogic, node 에서는 require('./logic.js') 로 씁니다.
   기준 문서: docs/01_프로젝트_기획서.md (3장 데이터, 5장 기능, 8장 1단계)

   흐름(인벤토리 데이터 01~06):
     ERP 재고 엑셀 → detectColumns·parseErp(열 찾기 → 재고 품목)                       [01 작업지시서]
     → makeSession(기준 일자·시각) → parseRoster·applyRoster(창고 배정표: 제외 창고 · 실사자 수만큼 나누기) 또는 unitsOf·autoAssign
     → orderSheets(재고실사 기록서 엑셀 = 작업지시서 = SharePoint 취합 파일)
     → validateCount(현장 입력: 실사 수량·사유·특이사항·확인 시각·작성자)              [02 재고수량확인 · 04 차이기록]
     → parseOrderWorkbook(여러 사람이 채운 SharePoint 취합 파일) · resultFile·checkResultFile(휴대폰 결과 파일) → mergeRecords(합치기, 충돌 찾기) → resolveCounts
     → varianceRows(시스템 수량 vs 실사 수량, 수량·금액 차이, 차이율, 허용 기준)       [03 차이대사]
     → summarize·reasonStats·topDiffs(창고별·창고위치별·자재유형별·자재타입별·실사자별·사유별) [05 작업결과]
     → nextReportMeta·reportSheets·buildReportText·buildAiPromptInfo(버전·원천 기록 연결)  [06 보고서] */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.STLogic = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';
  /* 판(版) — index.html 의 ?v= · app.js 의 APP_VERSION 과 같아야 합니다(scripts/bump-version.mjs 로 함께 올림).
     브라우저가 옛 logic.js 와 새 app.js 를 섞어 쓰면 app.js 가 이 값을 보고 계산을 멈춥니다. */
  var VERSION = '2026-11-10.1';

  var SCHEMA_VERSION = 1;
  var RESULT_KIND = 'data10-14.result';

  /* ── 1. 열 이름 ────────────────────────────────────────────────
     2차 반영(2026-10-21): 수강생이 보낸 ERP 재고 엑셀 견본(값 가린 사본)의 열 —
       자재코드 · 자재내역 · 저장위치 · 저장위치명 · 창고위치 · 단가 · 금액 · Location · 수량
     이 견본에서 「저장위치」는 창고 코드(SAP 의 Storage Location), 「저장위치명」은 창고 이름,
     「창고위치」는 사내/사외, 「Location」은 창고 안 보관 자리(팔레트·랙)입니다. 그래서 저장위치 → 창고로 읽습니다.
     (같은 파일에 「창고」 열이 따로 있고 보관 자리 열이 없으면 저장위치를 보관 자리로 돌려 읽습니다.)
     견본 첫 줄에 「필드가 추가될 수 있음」이라 되어 있어, 모르는 열도 버리지 않고 작업지시서·기록서에 그대로 싣습니다.
     머리행 글자는 공백·괄호·가운데점·빗금을 빼고 소문자로 맞춘 뒤, 끝의 단위(개·EA·원 …)를 떼고 비교합니다. */
  var ERP_FIELDS = [
    { id: 'no', label: '재고번호', aliases: ['재고번호', '재고no', '재고id', '재고관리번호', '관리번호', 'stockno', 'stockid', 'inventoryno', '실사번호', '태그번호', 'tagno'] },
    { id: 'code', label: '자재코드', aliases: ['자재코드', '자재번호', '자재', '품번', '품목코드', '품목번호', '부품번호', 'material', 'materialno', 'materialcode', 'itemcode', 'itemno', 'partno'] },
    { id: 'name', label: '자재명', aliases: ['자재명', '자재내역', '자재설명', '품명', '품목명', '부품명', 'description', 'materialdescription', 'itemname'] },
    { id: 'spec', label: '규격', aliases: ['규격', '사양', '규격사양', 'spec', 'specification'] },
    { id: 'unit', label: '단위', aliases: ['단위', '기본단위', '재고단위', 'uom', 'unit', 'bun'] },
    { id: 'wh', label: '창고(저장위치)', aliases: ['저장위치', '저장위치코드', '창고', '창고코드', '저장소', '저장장소', 'warehouse', 'wh', 'sloc', 'storagelocation', 'lgort'] },
    { id: 'whName', label: '창고명(저장위치명)', aliases: ['저장위치명', '저장위치내역', '저장위치설명', '창고명', '창고내역', '창고이름', '저장소명', 'slocname', 'warehousename', 'storagelocationname'] },
    { id: 'site', label: '창고위치(사내/사외)', aliases: ['창고위치', '사내외', '사내외구분', '창고구분', '보관지역', 'site'] },
    { id: 'loc', label: '보관 자리(Location)', aliases: ['location', 'loc', 'bin', 'storagebin', '빈', '보관위치', '보관자리', '적치위치', '로케이션', '위치', '랙', '랙번호', '선반', '팔레트', '파레트'] },
    { id: 'batch', label: '배치(로트)', aliases: ['배치', '배치번호', '로트', '로트번호', 'batch', 'batchno', 'lot', 'lotno', 'charg'] },
    { id: 'mtype', label: '자재유형', aliases: ['자재유형', '자재구분', '자재분류', '자재그룹', '자재군', '품목유형', '품목군', '품목그룹', '분류', 'materialgroup', 'category'] },
    { id: 'mkind', label: '자재타입', aliases: ['자재타입', '자재유형코드', '품목타입', '자재계정', '계정구분', '계정', '타입', 'materialtype', 'mtart', 'type'] },
    { id: 'sys', label: '시스템 수량', aliases: ['시스템수량', '전산수량', '장부수량', '재고수량', '현재고', '현재고수량', '재고', '재고량', '기말재고', '가용재고', '수량', 'qty', 'quantity', 'stock', 'stockqty', 'unrestricted'] },
    { id: 'price', label: '단가', aliases: ['단가', '재고단가', '이동평균단가', '평균단가', '표준단가', '원가', 'unitprice', 'price', 'map'] },
    { id: 'amt', label: '금액', aliases: ['재고금액', '장부금액', '금액', '재고가액', '가액', 'amount', 'value', 'stockvalue'] }
  ];
  /* 현장 기록(작업지시서 · 재고실사 기록서 · SharePoint 취합 엑셀)에서 더 읽는 열.
     수강생 기록서 견본의 말을 따릅니다: 담당자 = 창고 책임자(서명), 실사자 = 실제로 세는 사람(배정), 실물수량 = 센 수량 */
  var COUNT_FIELDS = [
    { id: 'cnt', label: '실물수량', aliases: ['실물수량', '실사수량', '현장수량', '실제수량', '확인수량', '실사', 'countqty', 'counted', 'physicalqty'] },
    { id: 'diffq', label: '차이수량', aliases: ['차이수량', '차이', '수량차이', 'diff', 'difference'] },
    { id: 'reason', label: '사유', aliases: ['사유', '차이사유', '사유구분', '사유카테고리', 'reason'] },
    { id: 'note', label: '특이사항', aliases: ['특이사항', '비고', '메모', 'note', 'remark', 'remarks'] },
    { id: 'by', label: '확인자', aliases: ['확인자', '작성자', '확인자명', 'countedby'] },
    { id: 'at', label: '확인 시각', aliases: ['확인시각', '확인일시', '확인시간', '실사시각', '실사일시', 'countedat'] },
    { id: 'assignee', label: '실사자', aliases: ['실사자', '실사담당', '실사자명', '배정실사자', '배정담당자', '배정', 'assignee'] },
    { id: 'manager', label: '담당자', aliases: ['담당자', '창고담당자', '책임자', '관리담당자', 'manager'] }
  ];
  var ALL_FIELDS = ERP_FIELDS.concat(COUNT_FIELDS);
  var FIELD_BY_ID = {}; ALL_FIELDS.forEach(function (f) { FIELD_BY_ID[f.id] = f; });
  /* 도구가 직접 채우는 열 — 기록서를 다시 읽을 때 「ERP 에서 온 열」로 보지 않습니다 */
  var TOOL_FIELDS = ['no', 'cnt', 'diffq', 'reason', 'note', 'by', 'at', 'assignee', 'manager'];

  /* 차이 사유 — 수강생 인벤토리 [데이터 04] 의 예시 6개가 기본값입니다.
     기록서 견본에 「설정에서 사유 리스트를 몇 개 작성할 수 있으면 함」이 있어 설정에서 바꿀 수 있습니다(settings.reasons). */
  var REASONS = ['입출고 누락', '위치 이동·분산', '라벨·식별 오류', '파손·폐기 미반영', '조사 오류', '기타'];
  var REASON_ALIASES = {
    '입출고 누락': ['입출고누락', '입고누락', '출고누락', '입출고미처리', '미출고', '미입고', '전표누락', '입출고'],
    '위치 이동·분산': ['위치이동분산', '위치이동', '분산', '분산보관', '이동', '위치'],
    '라벨·식별 오류': ['라벨식별오류', '라벨오류', '식별오류', '라벨', '품번혼동', '오식별'],
    '파손·폐기 미반영': ['파손폐기미반영', '파손', '폐기', '불량', '파손미반영', '폐기미반영'],
    '조사 오류': ['조사오류', '오기', '카운트오류', '실사오류', '재실사', '못찾음'],
    '기타': ['기타', 'etc']
  };
  var NO_REASON = '사유 미입력';

  var DIMS = {
    wh: { label: '창고' }, zone: { label: '창고·구역' }, site: { label: '창고위치' }, mtype: { label: '자재유형' }, mkind: { label: '자재타입' }, assignee: { label: '실사자' }
  };
  /* 허용 기준 — 2차 반영: 수강생 답 「차이가 있으면 무조건 사유 필수」 → 기본은 허용 0(tolMode 'zero', 모든 차이가 확인 대상).
     1단계의 「차이율 2% 또는 10만 원」은 설정에서 고를 때만(tolMode 'range') 씁니다. 사유는 두 방식 모두 차이가 있으면 필수입니다. */
  var DEFAULT_SETTINGS = { tolMode: 'zero', tolPct: 2, tolAmt: 100000, topN: 15 };

  /* ── 2. 작은 도구 ──────────────────────────────────────────── */
  var UNIT_RE = /(ea|pcs|개|원|krw|kg|m)$/;
  function str(v) { return v == null ? '' : String(v).trim(); }
  function normRaw(s) { return String(s == null ? '' : s).replace(/[\s()（）\[\]·.:_\/\-~]/g, '').toLowerCase(); }
  function norm(s) {
    var n = String(s == null ? '' : s).replace(/[\s()（）\[\]·.:_\/\-~]/g, '').toLowerCase();
    var m = n.replace(UNIT_RE, '');
    return m || n;
  }
  function toNum(v) {
    if (v == null) return null;
    if (typeof v === 'number') return isFinite(v) ? v : NaN;
    if (typeof v === 'boolean') return NaN;
    var s = String(v).trim();
    if (!s || s === '-') return null;
    if (s.charAt(0) === '=') return null;
    s = s.replace(/,/g, '').replace(/\s*(ea|pcs|개|원|kg|m|box|set)$/i, '');
    if (/^\(\d+(\.\d+)?\)$/.test(s)) s = '-' + s.slice(1, -1); // 회계식 음수 (1,200)
    if (!/^-?\d+(\.\d+)?$/.test(s)) return NaN;
    return Number(s);
  }
  function round(n, d) { var p = Math.pow(10, d == null ? 2 : d); return Math.round(n * p) / p; }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function colName(i) {
    var s = ''; i = Number(i);
    do { s = String.fromCharCode(65 + (i % 26)) + s; i = Math.floor(i / 26) - 1; } while (i >= 0);
    return s;
  }
  /* 지역 시각 'YYYY-MM-DDTHH:MM:SS' (확인 시각은 현장 기기의 시계 그대로) */
  function localIso(d) {
    d = d || new Date();
    return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) + 'T' + pad(d.getHours()) + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds());
  }
  /* 시각 읽기 — Date, 엑셀 일련번호, '2026-10-01 09:30', '2026.10.01 9:30' … → 'YYYY-MM-DDTHH:MM:SS' 또는 '' */
  function toDateTime(v) {
    if (v == null || v === '') return '';
    if (v instanceof Date && !isNaN(v)) {
      // SheetJS 가 날짜 칸을 UTC 자정 기준으로 줄 때가 있어 초 단위 반올림
      var t = new Date(Math.round(v.getTime() / 1000) * 1000);
      return localIso(t);
    }
    if (typeof v === 'number' && v > 20000 && v < 80000) {
      var ms = Math.round((v - 25569) * 86400) * 1000, u = new Date(ms);
      return u.getUTCFullYear() + '-' + pad(u.getUTCMonth() + 1) + '-' + pad(u.getUTCDate()) + 'T' + pad(u.getUTCHours()) + ':' + pad(u.getUTCMinutes()) + ':' + pad(u.getUTCSeconds());
    }
    var m = /^(\d{4})[-.\/](\d{1,2})[-.\/](\d{1,2})\.?(?:[ T]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(str(v));
    if (!m) return '';
    return m[1] + '-' + pad(+m[2]) + '-' + pad(+m[3]) + 'T' + pad(+(m[4] || 0)) + ':' + (m[5] || '00') + ':' + (m[6] || '00');
  }
  function fmtTime(iso) { return iso ? iso.slice(0, 16).replace('T', ' ') : '-'; }
  function fmtNum(n) { return n == null || !isFinite(n) ? '-' : Number(round(n, 2)).toLocaleString('ko-KR'); }
  function fmtWon(n) { return n == null || !isFinite(n) ? '-' : Math.round(n).toLocaleString('ko-KR') + '원'; }
  function fmtSigned(n, won) { if (n == null || !isFinite(n)) return '-'; var s = won ? fmtWon(Math.abs(n)) : fmtNum(Math.abs(n)); return (n > 0 ? '+' : n < 0 ? '-' : '') + s; }
  function pctText(v, d) { return v == null || !isFinite(v) ? '-' : round(v * 100, d == null ? 1 : d) + '%'; }
  function cmp(a, b) { a = a == null ? '' : String(a); b = b == null ? '' : String(b); return a < b ? -1 : a > b ? 1 : 0; }
  function searchKey(s) { return String(s == null ? '' : s).replace(/[\s\-_.\/]/g, '').toLowerCase(); }

  /* ── 3. 머리행 · 열 찾기 ───────────────────────────────────── */
  function matchField(text, fields) {
    var n = norm(text);
    if (!n) return null;
    fields = fields || ALL_FIELDS;
    for (var i = 0; i < fields.length; i++) if (fields[i].aliases.indexOf(n) >= 0) return fields[i].id;
    // 비슷한 머리(예: 「재고 수량(EA)」「창고 코드」「자재 내역」)
    var ids = fields.map(function (f) { return f.id; });
    function has(id) { return ids.indexOf(id) >= 0; }
    if (has('cnt') && /실사|실물|현장/.test(n) && /수량|qty/.test(n)) return 'cnt';
    if (has('sys') && /(장부|전산|시스템|재고)/.test(n) && /수량|qty/.test(n)) return 'sys';
    if (has('amt') && /금액|가액/.test(n)) return 'amt';
    if (has('price') && /단가/.test(n)) return 'price';
    // 「저장 위치 내역」(SAP 저장위치 설명)은 창고명입니다 — 「위치」가 들어 있다고 보관 자리로 읽으면 안 됩니다(10-28 수강생 ERP RAW)
    if (has('whName') && /(창고|저장위치)/.test(n) && /(명|이름|name|내역|설명)/.test(n)) return 'whName';
    if (has('site') && /(사내|사외)/.test(n)) return 'site';
    if (has('wh') && /창고/.test(n)) return 'wh';
    if (has('loc') && /(위치|로케이션|location|bin)/.test(n)) return 'loc';
    if (has('name') && /(자재명|품명|내역)/.test(n)) return 'name';
    if (has('at') && /(확인|실사)/.test(n) && /(시각|일시|시간)/.test(n)) return 'at';
    return null;
  }
  /* aoa: 시트를 행 배열로 읽은 것(SheetJS sheet_to_json header:1).
     반환: { headerRow, map:{field:col}, headers:[글자], missing:[], warnings:[] }
     꼭 필요한 것: (재고번호 또는 자재코드) + 창고 + 시스템 수량. mode:'count' 이면 (재고번호 또는 자재코드) + 실사 수량. */
  /* 머리 글자 → 항목 후보. 별칭과 똑같은 머리(예: 「Location」)가 비슷하기만 한 머리보다 먼저 —
     10-28: 「저장 위치 내역」이 비슷한 말로 보관 자리 후보가 되어, 뒤에 있는 진짜 「Location」 열을 밀어냈습니다 */
  function headerCands(headers) {
    var cands = {}, exact = {};
    headers.forEach(function (t, c) {
      var fid = matchField(t);
      if (!fid) return;
      (cands[fid] = cands[fid] || []).push(c);
      if (FIELD_BY_ID[fid].aliases.indexOf(norm(t)) >= 0) exact[c] = 1;
    });
    Object.keys(cands).forEach(function (fid) { cands[fid].sort(function (a, b) { return (exact[b] || 0) - (exact[a] || 0) || a - b; }); });
    return cands;
  }
  function detectColumns(aoa, overrides, mode) {
    aoa = aoa || [];
    var limit = Math.min(aoa.length, 30);
    var best = -1, bestScore = 0, bestMap = null, bestHeaders = null, bestDup = null;
    for (var r = 0; r < limit; r++) {
      var row = aoa[r] || [], headers = [];
      for (var c = 0; c < row.length; c++) headers.push(str(row[c]));
      var cands = headerCands(headers);
      var res = resolveCands(cands, headers), map = res.map, dup = res.dup;
      var key = map.no != null || map.code != null;
      var ok = mode === 'count' ? key && map.cnt != null : key && map.sys != null && map.wh != null;
      var score = Object.keys(map).length;
      if (ok && score > bestScore) { best = r; bestScore = score; bestMap = map; bestHeaders = headers; bestDup = dup; }
    }
    var out = { headerRow: best, map: {}, headers: [], missing: [], warnings: [] };
    if (best < 0) {
      out.warnings.push(mode === 'count'
        ? '머리행(재고번호·실사 수량 …)을 찾지 못했습니다. 작업지시서에 실사 수량을 적은 파일인지 확인해 주세요.'
        : '머리행(재고번호 또는 자재코드 · 창고 · 재고수량 …)을 찾지 못했습니다. ERP 재고 시트인지 확인하고, 열 연결에서 골라 주세요.');
      // 열 연결을 직접 하도록 가장 글자가 많은 줄을 머리행 후보로 둡니다
      var most = -1, cnt = 0;
      for (var k = 0; k < limit; k++) { var n = (aoa[k] || []).filter(function (x) { return str(x); }).length; if (n > cnt) { cnt = n; most = k; } }
      out.headerRow = most; out.headers = most >= 0 ? (aoa[most] || []).map(str) : [];
      out.map = resolveCands(headerCands(out.headers), out.headers).map; // 찾은 것만이라도 미리 연결
    } else {
      out.map = bestMap; out.headers = bestHeaders;
      bestDup.forEach(function (d) { out.warnings.push('「' + d.label + '」 열이 ' + FIELD_BY_ID[d.fid].label + '(으)로 한 번 더 보입니다. 앞의 것(' + colName(out.map[d.fid]) + ')을 씁니다.'); });
    }
    if (overrides) Object.keys(overrides).forEach(function (fid) {
      if (!FIELD_BY_ID[fid]) return;
      var v = overrides[fid];
      if (v === -1 || v === '' || v == null) delete out.map[fid]; else out.map[fid] = Number(v);
    });
    var m = out.map;
    if (m.no == null && m.code == null) out.missing.push('no');
    if (mode === 'count') { if (m.cnt == null) out.missing.push('cnt'); }
    else { if (m.wh == null) out.missing.push('wh'); if (m.sys == null) out.missing.push('sys'); }
    if (out.missing.length && best >= 0) out.warnings.push('꼭 필요한 열이 없습니다: ' + out.missing.map(function (id) { return id === 'no' ? '재고번호 또는 자재코드' : FIELD_BY_ID[id].label; }).join(', ') + ' — 열 연결에서 골라 주세요.');
    if (!out.missing.length && mode !== 'count') {
      var soft = ['price', 'loc'].filter(function (id) { return m[id] == null; });
      if (m.price == null && m.amt != null) soft = soft.filter(function (x) { return x !== 'price'; });
      if (soft.length) out.warnings.push('이 열이 없어 해당 값이 비어 보입니다: ' + soft.map(function (id) { return FIELD_BY_ID[id].label; }).join(', ') + '. 있는데 못 찾았다면 열 연결에서 골라 주세요.');
      var dims = ['mtype', 'mkind'].filter(function (id) { return m[id] == null; });
      if (dims.length) out.info = dims.map(function (id) { return FIELD_BY_ID[id].label; }).join('·') + ' 열이 없어 그 분석은 숨깁니다(열이 추가되면 자동으로 보이고, 이름이 다르면 열 연결에서 고름).';
    }
    return out;
  }
  /* 한 머리행에서 같은 항목으로 읽힌 열이 여럿일 때 정리합니다.
     · 「창고」와 「저장위치」가 함께 있고 보관 자리 열이 없으면 → 창고 = 창고, 저장위치 = 보관 자리(1단계 모양)
     · 창고 코드 열이 없고 창고명만 있으면 → 창고명을 창고로 씀 */
  function resolveCands(cands, headers) {
    var wh = cands.wh || [];
    if (wh.length > 1 && !(cands.loc && cands.loc.length)) {
      var sl = wh.filter(function (c) { return /^저장위치/.test(norm(headers[c])); });
      if (sl.length && sl.length < wh.length) { cands.loc = [sl[0]]; cands.wh = wh.filter(function (c) { return c !== sl[0]; }); }
    }
    var map = {}, dup = [];
    Object.keys(cands).forEach(function (fid) {
      map[fid] = cands[fid][0];
      cands[fid].slice(1).forEach(function (c) { dup.push({ label: headers[c], fid: fid }); });
    });
    if (map.wh == null && map.whName != null) map.wh = map.whName;
    return { map: map, dup: dup };
  }
  function isSummaryRow(row) {
    for (var c = 0; c < Math.min(row.length, 8); c++) {
      if (/^(합계|소계|총계|계|누계|total|sum|subtotal|grandtotal)$/i.test(str(row[c]).replace(/\s/g, ''))) return true;
    }
    return false;
  }
  function isBlankRow(row) { return !row || !row.some(function (x) { return str(x) !== ''; }); }
  /* 재고번호가 없는 ERP 파일은 자재코드 @ 창고 / 위치 (/ 배치) 로 열쇠를 만듭니다 */
  function makeKey(code, wh, loc, batch) { return code + '@' + wh + (loc ? '/' + loc : '') + (batch ? '/' + batch : ''); }
  /* 줄마다 하나뿐인 재고번호 — 10-28 수강생 데이터 신뢰도 지적: 「작업지시는 올린 그대로, 한 줄도 합치지 않는다」.
     같은 열쇠가 또 나오면(같은 자재·창고·위치가 여러 줄) 합치지 않고 나온 순서대로 「#2」「#3」을 붙입니다.
     ERP 파일과 그 파일로 만든 기록서를 같은 순서로 읽으므로 같은 줄은 같은 번호가 됩니다. */
  function rowKeyer() {
    var seen = {};
    return function (base) { var n = seen[base] = (seen[base] || 0) + 1; return n > 1 ? base + '#' + n : base; };
  }
  /* 날짜 열(최종 입고일 · 생산일자 …) — 엑셀 일련번호(45930)로 오면 'YYYY-MM-DD' 로 둡니다 */
  var DATE_HEAD_RE = /(일자|날짜|일시|date|일)$/;
  function isDateHead(h) { var n = norm(h); return !!n && DATE_HEAD_RE.test(n) && !/(요일|단위|일련)/.test(n); }
  function extValue(h, v) {
    if (v instanceof Date) return localIso(v).slice(0, 10);
    if (isDateHead(h)) { var d = toDateTime(v); if (d) return d.slice(11) === '00:00:00' ? d.slice(0, 10) : d.slice(0, 16).replace('T', ' '); }
    return v;
  }

  /* 머리행의 열 목록 — [{ h: 머리 글자, f: 항목 id 또는 null(모르는 열 — 그대로 싣기), c: 열 번호 }] */
  function sourceCols(det) {
    var used = {}, cols = [], seen = {};
    Object.keys(det.map).forEach(function (f) { if (used[det.map[f]] == null) used[det.map[f]] = f; });
    (det.headers || []).forEach(function (t, c) {
      if (!t) return;
      var f = used[c] != null ? used[c] : null, h = t;
      if (f == null) { while (seen[h]) h = h + '_'; seen[h] = 1; }
      cols.push({ h: h, f: f, c: c });
    });
    return cols;
  }
  /* ERP 재고 시트 → 재고 품목. opts: { overrides, fileName } */
  function parseErp(aoa, opts) {
    opts = opts || {};
    var det = detectColumns(aoa, opts.overrides);
    var out = { det: det, items: [], cols: [], skipped: [], warnings: [], summary: 0, blankQty: 0, noPrice: 0, merged: 0, dupKeys: 0, generatedKey: false };
    if (det.missing.length) return out;
    out.cols = sourceCols(det).filter(function (x) { return TOOL_FIELDS.indexOf(x.f) < 0 || x.f === 'no'; });
    var extCols = out.cols.filter(function (x) { return x.f == null; });
    var m = det.map, keyOf = rowKeyer();
    function cell(row, f) { return m[f] == null ? null : row[m[f]]; }
    out.generatedKey = m.no == null;
    for (var r = det.headerRow + 1; r < aoa.length; r++) {
      var row = aoa[r] || [];
      if (isBlankRow(row)) continue;
      if (isSummaryRow(row)) { out.summary++; continue; }
      var code = str(cell(row, 'code')), wh = str(cell(row, 'wh')), loc = str(cell(row, 'loc'));
      var no = m.no != null ? str(cell(row, 'no')) : (code ? makeKey(code, wh, loc, str(cell(row, 'batch'))) : '');
      if (!no) { out.skipped.push({ row: r + 1, why: '재고번호·자재코드 빈칸' }); continue; }
      var key = keyOf(no);
      if (key !== no) out.dupKeys++;
      no = key;
      if (!wh) { out.skipped.push({ row: r + 1, why: '창고 빈칸 (' + no + ')' }); continue; }
      var sys = toNum(cell(row, 'sys'));
      if (sys === null) { sys = 0; out.blankQty++; }
      if (isNaN(sys)) { out.skipped.push({ row: r + 1, why: '수량을 숫자로 못 읽음 (' + no + ': ' + str(cell(row, 'sys')) + ')' }); continue; }
      var price = toNum(cell(row, 'price')), amt = toNum(cell(row, 'amt'));
      if (isNaN(price)) price = null;
      if (isNaN(amt)) amt = null;
      if (price == null && amt != null && sys) price = amt / sys;
      if (price == null) { price = 0; out.noPrice++; }
      if (amt == null) amt = sys * price;
      var it = {
        no: no, code: code, name: str(cell(row, 'name')), spec: str(cell(row, 'spec')), unit: str(cell(row, 'unit')),
        wh: wh, whName: m.whName != null && m.whName !== m.wh ? str(cell(row, 'whName')) : '', site: str(cell(row, 'site')),
        loc: loc, mtype: str(cell(row, 'mtype')), mkind: str(cell(row, 'mkind')),
        sys: sys, price: price, amt: amt, row: r + 1
      };
      if (m.batch != null) it.batch = str(cell(row, 'batch'));
      if (extCols.length) { it.ext = {}; extCols.forEach(function (x) { var v = row[x.c]; if (v != null && v !== '') it.ext[x.h] = extValue(x.h, v); }); }
      // 10-28: 같은 열쇠의 줄을 합치지 않습니다(예전에는 수량·금액을 더해 한 줄로 — 수강생 기록서에 엉뚱한 기말재고가 찍힌 원인)
      out.items.push(it);
    }
    if (out.dupKeys) out.warnings.push('같은 ' + (m.no != null ? '재고번호' : '자재·창고·위치') + '인 줄이 ' + out.dupKeys + '개 더 있습니다. 합치지 않고 ERP 줄마다 따로 실사하도록 재고번호 뒤에 「#2」「#3」을 붙였습니다.');
    if (out.blankQty) out.warnings.push('수량 빈칸 ' + out.blankQty + '행은 0 으로 읽었습니다.');
    if (out.noPrice) out.warnings.push('단가·금액이 없는 ' + out.noPrice + '행은 금액 차이를 0 으로 셉니다.');
    if (out.generatedKey) out.warnings.push('재고번호 열이 없어 「자재코드@창고/위치」를 재고번호로 씁니다.');
    return out;
  }
  /* ERP 재고 시트 고르기 — ERP 항목을 많이 찾은 시트, 같으면 실사 수량 열이 없는 쪽(이미 실사한 기록서가 아닌 원본) */
  function pickSheet(sheets) {
    var best = null, bestScore = -1;
    (sheets || []).forEach(function (s) {
      var d = detectColumns(s.aoa);
      var erp = Object.keys(d.map).filter(function (f) { return ERP_FIELDS.some(function (x) { return x.id === f; }); }).length;
      var sc = d.missing.length ? -1 : erp * 1000 - (d.map.cnt != null ? 500 : 0) + Math.min(400, s.aoa.length - d.headerRow);
      if (sc > bestScore) { bestScore = sc; best = s; }
    });
    return bestScore >= 0 ? best : null;
  }
  function headerSignature(det) { return det && det.headers ? det.headers.map(norm).join('|') : ''; }

  /* ── 4. 실사 회차 · 작업 배정 ──────────────────────────────── */
  function validBaseAt(v) { return /^\d{4}-\d\d-\d\dT\d\d:\d\d(:\d\d)?$/.test(str(v)); }
  /* opts: { title, baseAt('YYYY-MM-DDTHH:MM'), sourceFile, unitMode('wh'|'zone'), cols(ERP 원래 열), id?, now? } */
  function makeSession(opts) {
    opts = opts || {};
    if (!validBaseAt(opts.baseAt)) throw new Error('기준 일자·시각을 넣어 주세요(ERP 에서 재고를 내려받은 시각).');
    var baseAt = opts.baseAt.length === 16 ? opts.baseAt + ':00' : opts.baseAt;
    var now = opts.now || localIso();
    var id = opts.id || ('ST-' + baseAt.slice(0, 16).replace(/[-:T]/g, '') + '-' + Math.random().toString(36).slice(2, 6).toUpperCase());
    return {
      id: id, title: str(opts.title) || (baseAt.slice(0, 7) + ' 정기 재고실사'), baseAt: baseAt,
      createdAt: now, sourceFile: str(opts.sourceFile), unitMode: opts.unitMode === 'zone' ? 'zone' : 'wh', orderVer: 1,
      cols: (opts.cols || []).map(function (x) { return { h: x.h, f: x.f }; }), managers: {}, excluded: [], rosterFile: ''
    };
  }
  /* 보관 자리 앞자리(「A-01-03」→「A」, 「2F_R12」→「2F」, 「7번 랙」→「7번」)를 구역으로 봅니다 */
  function zoneOf(loc) { var s = str(loc); if (!s) return '(위치 없음)'; return s.split(/[-_\s\/.]/)[0] || s; }
  /* 배정 단위 — 창고별(배정표로 한 창고를 여러 실사자가 나누면 「창고 #1」「창고 #2」) 또는 창고·구역별 */
  function unitKey(item, mode) { return mode === 'zone' ? item.wh + ' / ' + zoneOf(item.loc) : item.wh + (item.part ? ' #' + item.part : ''); }
  function whLabel(it) { return it.whName && it.whName !== it.wh ? it.wh + ' ' + it.whName : it.wh; }
  function unitsOf(items, mode) {
    var by = {}, list = [];
    (items || []).forEach(function (it) {
      var k = unitKey(it, mode);
      if (!by[k]) { by[k] = { key: k, wh: it.wh, whName: it.whName || '', site: it.site || '', part: mode === 'zone' ? 0 : it.part || 0, zone: mode === 'zone' ? zoneOf(it.loc) : '', items: 0, sysAmt: 0 }; list.push(by[k]); }
      by[k].items++; by[k].sysAmt += it.amt || 0;
    });
    return list.sort(function (a, b) { return cmp(a.wh, b.wh) || a.part - b.part || cmp(a.key, b.key); });
  }
  /* 실사자 이름 목록으로 배정 단위를 나눕니다 — 품목 수가 많은 단위부터 가장 덜 맡은 사람에게(창고는 쪼개지 않음) */
  function autoAssign(units, names) {
    names = (names || []).map(str).filter(Boolean);
    var out = {};
    if (!names.length) return out;
    var load = names.map(function () { return 0; });
    units.slice().sort(function (a, b) { return b.items - a.items || cmp(a.key, b.key); }).forEach(function (u) {
      var k = 0; for (var i = 1; i < names.length; i++) if (load[i] < load[k]) k = i;
      out[u.key] = names[k]; load[k] += u.items;
    });
    return out;
  }
  function assigneeOf(item, session, assign) { return (assign || {})[unitKey(item, session ? session.unitMode : 'wh')] || ''; }
  function managerOf(item, session) { return session && session.managers ? session.managers[item.wh] || '' : ''; }
  function assignedNames(assign) {
    var s = {}; Object.keys(assign || {}).forEach(function (k) { if (assign[k]) s[assign[k]] = 1; });
    return Object.keys(s).sort(cmp);
  }

  /* ── 4-1. 창고 배정표 (수강생 시트 2 「창고 목록과 창고별 담당 인원」) ──
     머리: 저장위치 · 저장위치명 · 창고위치 · 실사인원 · 담당팀 · 담당자 · 실사자 · 비고
     · 실사인원 0 → 이번 실사에서 뺌(비고 예: 담당자 확인서로 대신함)
     · 실사자 「갑, 을」 → 그 창고 품목을 보관 자리 순서로 사람 수만큼 나눔
     · 표 위 칸의 날짜(배정표 날짜)는 기준 일자 칸의 처음 값으로 씁니다 */
  var ROSTER_FIELDS = [
    { id: 'wh', aliases: ['저장위치', '저장위치코드', '창고', '창고코드', 'sloc', 'storagelocation'] },
    { id: 'whName', aliases: ['저장위치명', '창고명', '창고이름'] },
    { id: 'site', aliases: ['창고위치', '사내외', '사내외구분'] },
    { id: 'n', aliases: ['실사인원', '인원', '인원수', '실사인원수', '투입인원'] },
    { id: 'team', aliases: ['담당팀', '팀', '부서', '담당부서'] },
    { id: 'manager', aliases: ['담당자', '창고담당자', '책임자'] },
    { id: 'counters', aliases: ['실사자', '실사담당', '실사자명', '조사자', '작업자'] },
    { id: 'note', aliases: ['비고', '메모', '특이사항'] }
  ];
  function splitNames(v) {
    return str(v).split(/[,，、;\/\n]+|\s{2,}/).map(function (x) { return x.trim(); })
      .filter(function (x) { return x && !/^[-–—_.]+$/.test(x) && x !== '없음'; });
  }
  function parseRoster(aoa) {
    aoa = aoa || [];
    var limit = Math.min(aoa.length, 30), hr = -1, map = null;
    for (var r = 0; r < limit && hr < 0; r++) {
      var m = {};
      (aoa[r] || []).forEach(function (x, c) {
        var n = norm(x), raw = normRaw(x); if (!n) return; // 「실사인원」은 끝 글자 「원」이 단위로 떼이지 않게 원래 글자로도 견줌
        ROSTER_FIELDS.forEach(function (f) { if (m[f.id] == null && (f.aliases.indexOf(raw) >= 0 || f.aliases.indexOf(n) >= 0)) m[f.id] = c; });
      });
      if ((m.wh != null || m.whName != null) && (m.counters != null || m.n != null || m.manager != null)) { hr = r; map = m; }
    }
    if (hr < 0) return { error: '배정표 머리행(저장위치 · 실사인원 · 담당자 · 실사자 …)을 찾지 못했습니다.' };
    var date = '';
    for (var k = 0; k < hr && !date; k++) (aoa[k] || []).forEach(function (x) { if (!date) { var d = toDateTime(x); if (d && (x instanceof Date || /^\d{4}[-.\/]/.test(str(x)) || typeof x === 'number')) date = d.slice(0, 10); } });
    var rows = [], warnings = [];
    function cell(row, f) { return map[f] == null ? null : row[map[f]]; }
    for (var i = hr + 1; i < aoa.length; i++) {
      var row = aoa[i] || [];
      if (isBlankRow(row) || isSummaryRow(row)) continue;
      var wh = str(cell(row, 'wh')), whName = str(cell(row, 'whName'));
      if (!wh && !whName) continue;
      var n = toNum(cell(row, 'n')); if (n != null && isNaN(n)) n = null;
      var counters = splitNames(cell(row, 'counters')), manager = str(cell(row, 'manager'));
      var o = { wh: wh || whName, whName: wh ? whName : '', site: str(cell(row, 'site')), n: n, team: str(cell(row, 'team')), manager: manager, counters: counters, note: str(cell(row, 'note')), excluded: n === 0, row: i + 1 };
      var label = whLabel(o);
      if (!o.excluded) {
        if (!counters.length) warnings.push(label + ': 실사자가 비어 있어 ' + (manager ? '담당자(' + manager + ')에게 맡깁니다.' : '배정하지 않습니다.'));
        else if (n != null && n !== counters.length) warnings.push(label + ': 실사인원 ' + n + '명인데 실사자 이름은 ' + counters.length + '명 — 이름 수대로 나눕니다.');
      }
      rows.push(o);
    }
    if (!rows.length) return { error: '배정표에 창고 줄이 없습니다.' };
    return { headerRow: hr, date: date, rows: rows, warnings: warnings };
  }
  /* 배정표를 재고 품목에 적용 → { items(제외 창고 뺀 것, 나눈 창고는 part), assign, managers, excluded, unrostered, missing } */
  function applyRoster(items, roster) {
    var byKey = {};
    roster.rows.forEach(function (r) { byKey['c:' + norm(r.wh)] = r; if (r.whName) byKey['n:' + norm(r.whName)] = r; });
    function find(it) { return byKey['c:' + norm(it.wh)] || (it.whName && byKey['n:' + norm(it.whName)]) || byKey['n:' + norm(it.wh)] || null; }
    var kept = [], ex = {}, exList = [], exItems = [], groups = {}, gList = [], unr = {}, managers = {}, assign = {}, used = [];
    (items || []).forEach(function (it0) {
      var it = {}; Object.keys(it0).forEach(function (k) { if (k !== 'part') it[k] = it0[k]; });
      var r = find(it);
      if (!r) { unr[it.wh] = (unr[it.wh] || 0) + 1; kept.push(it); return; }
      if (used.indexOf(r) < 0) used.push(r);
      if (!it.whName && r.whName) it.whName = r.whName;
      if (!it.site && r.site) it.site = r.site;
      if (r.excluded) {
        if (!ex[it.wh]) { ex[it.wh] = { wh: it.wh, whName: it.whName || '', items: 0, sysAmt: 0, manager: r.manager, team: r.team, note: r.note }; exList.push(ex[it.wh]); }
        ex[it.wh].items++; ex[it.wh].sysAmt += it.amt || 0; exItems.push(it);
        return;
      }
      if (r.manager) managers[it.wh] = r.manager;
      if (!groups[it.wh]) { groups[it.wh] = { r: r, list: [] }; gList.push(it.wh); }
      groups[it.wh].list.push(it); kept.push(it);
    });
    gList.forEach(function (wh) {
      var g = groups[wh], names = g.r.counters.length ? g.r.counters : (g.r.manager ? [g.r.manager] : []);
      if (!names.length) return;
      if (names.length === 1) { assign[wh] = names[0]; return; }
      var list = sortForWalk(g.list), n = list.length, k = names.length;
      list.forEach(function (it, i) { it.part = Math.min(k - 1, Math.floor(i * k / n)) + 1; });
      names.forEach(function (nm, i) { assign[wh + ' #' + (i + 1)] = nm; });
    });
    return {
      items: kept, assign: assign, managers: managers, excluded: exList, excludedItems: exItems,
      unrostered: Object.keys(unr).sort(cmp).map(function (w) { return { wh: w, items: unr[w] }; }),
      missing: roster.rows.filter(function (r) { return used.indexOf(r) < 0; }).map(function (r) { return whLabel(r); })
    };
  }

  /* ── 4-2. 재고실사 기록서 엑셀 = 작업지시서 = SharePoint 취합 파일 ──
     수강생 기록서 견본(시트 3)처럼 ERP 원래 열을 그대로 두고 뒤에 담당자 · 실사자 · 실물수량 · 차이수량 · 사유 · 특이사항을 붙입니다
     (확인자 · 확인시각은 인벤토리 02 의 빈칸이라 끝에 둠). 「지시정보」 시트에 기준 일자·시각 · 지시서 ID · 버전.
     지금은 실사 뒤 SharePoint 에 취합 파일 하나를 올려 각자 채운다고 해서(답 ⑤), 이 파일 하나를 그대로 올리고
     각자 자기 이름 시트(또는 전체 시트의 자기 줄)에 적은 뒤 ③ 취합에 넣으면 됩니다. 현장 기록이 있으면 채워서 냅니다. */
  var COUNT_HEAD = ['담당자', '실사자', '실물수량', '차이수량', '사유', '특이사항', '확인자', '확인시각'];
  var MAIN_SHEET = '재고실사기록서';
  var DEFAULT_COLS = [
    { h: '재고번호', f: 'no' }, { h: '자재코드', f: 'code' }, { h: '자재명', f: 'name' }, { h: '규격', f: 'spec' }, { h: '단위', f: 'unit' },
    { h: '창고', f: 'wh' }, { h: '창고명', f: 'whName' }, { h: '창고위치', f: 'site' }, { h: '보관 자리', f: 'loc' }, { h: '자재유형', f: 'mtype' }, { h: '자재타입', f: 'mkind' },
    { h: '시스템수량', f: 'sys' }, { h: '단가', f: 'price' }, { h: '금액', f: 'amt' }
  ];
  function bookCols(session, items) {
    var cols = session && session.cols && session.cols.length ? session.cols.slice() : DEFAULT_COLS.filter(function (c) {
      return ['no', 'wh', 'sys'].indexOf(c.f) >= 0 || (items || []).some(function (it) { return it[c.f] != null && it[c.f] !== '' && it[c.f] !== 0; });
    });
    if (!cols.some(function (c) { return c.f === 'no'; })) cols.unshift({ h: '재고번호', f: 'no' });
    return cols;
  }
  function colValue(it, c) {
    if (c.f == null) return it.ext && it.ext[c.h] != null ? it.ext[c.h] : null;
    if (c.f === 'amt') return round(it.sys * it.price, 2);
    if (c.f === 'price') return round(it.price, 2);
    var v = it[c.f]; return v === '' ? null : v == null ? null : v;
  }
  function reasonList(settings) {
    var r = settings && Array.isArray(settings.reasons) ? settings.reasons.map(str).filter(Boolean) : [];
    var seen = {}; r = r.filter(function (x) { if (seen[x]) return false; seen[x] = 1; return true; }).slice(0, 12);
    return r.length ? r : REASONS.slice();
  }
  function tolText(S) {
    S = cleanSettings(S);
    return S.tolMode === 'range' ? '차이율 ' + S.tolPct + '% 또는 차이 금액 ' + fmtWon(S.tolAmt) + ' 넘으면 「허용 초과」' : '허용 0 — 차이가 있으면 모두 사유 필수(확인 대상)';
  }
  function orderInfoAoa(session, items, assign, opts) {
    var names = assignedNames(assign), per = {};
    items.forEach(function (it) { var a = assigneeOf(it, session, assign) || '(미배정)'; per[a] = (per[a] || 0) + 1; });
    var aoa = [
      ['재고실사 기록서 (작업지시서 · SharePoint 취합 파일)'],
      ['실사명', session.title],
      ['지시서 ID', session.id],
      ['지시서 버전', session.orderVer || 1],
      ['기준 일자·시각', fmtTime(session.baseAt)],
      ['작성 시각', fmtTime(session.createdAt)],
      ['ERP 파일', session.sourceFile || ''],
      ['배정표 파일', session.rosterFile || ''],
      ['배정 단위', session.unitMode === 'zone' ? '창고·구역' : '창고'],
      ['품목 수', items.length],
      [],
      ['실사자', '품목 수']
    ];
    names.concat(per['(미배정)'] ? ['(미배정)'] : []).forEach(function (n) { aoa.push([n, per[n] || 0]); });
    if ((session.excluded || []).length) {
      aoa.push([]); aoa.push(['이번 실사 제외 창고', '품목 수', '담당자', '비고']);
      session.excluded.forEach(function (x) { aoa.push([whLabel(x), x.items, x.manager || '', x.note || '']); });
    }
    aoa.push([]);
    aoa.push(['사유 구분', reasonList(opts.settings).join(' / ')]);
    aoa.push(['허용 기준', tolText(opts.settings)]);
    aoa.push(['차이수량', '실물수량 − 수량(시스템). 모자라면 −, 남으면 +. 실물수량을 적으면 저절로 계산됩니다.']);
    aoa.push(['작성 방법', '이 파일 하나를 SharePoint 에 올리고, 각자 자기 이름 시트(또는 「' + MAIN_SHEET + '」 시트의 자기 줄)에 실물수량을 적습니다. 차이가 있으면 사유를 위 구분에서 골라 적고, 필요하면 특이사항·확인자·확인시각도 적습니다. 다 적으면 이 파일을 그대로 ③ 취합에 넣습니다.']);
    return aoa;
  }
  /* opts: { records, picks, only(이 실사자 줄만), settings } */
  function orderSheets(session, items, assign, opts) {
    opts = opts || {};
    var cols = bookCols(session, items), head = cols.map(function (c) { return c.h; }).concat(COUNT_HEAD);
    var sysIdx = cols.map(function (c) { return c.f; }).indexOf('sys'), cntIdx = cols.length + 2, dIdx = cols.length + 3;
    var res = opts.records ? resolveCounts(opts.records, opts.picks) : {};
    var sorted = sortForWalk(items);
    if (opts.only) sorted = sorted.filter(function (it) { var c = res[it.no] && res[it.no].current; return assigneeOf(it, session, assign) === opts.only || (c && c.by === opts.only); });
    function row(it) {
      var c = res[it.no] ? res[it.no].current : null;
      return cols.map(function (x) { return colValue(it, x); }).concat([
        managerOf(it, session) || null, assigneeOf(it, session, assign) || null,
        c ? c.qty : null, c ? round(c.qty - it.sys, 6) : null, c ? c.reason || null : null, c ? c.note || null : null,
        c ? c.by : null, c && c.at ? fmtTime(c.at) : null]);
    }
    function sheet(name, list) {
      var aoa = [head].concat(list.map(row));
      // 차이수량 칸은 엑셀 수식(실물수량을 적으면 저절로) — 쓰는 쪽(app · make-samples)에서 셀에 붙입니다
      // 3차 반영(10-27): 방향은 실물 − 수량(모자라면 −). 수강생 답 「모자라면 - 로 표현」. 화면 · 보고서의 「차이」와 같은 방향
      var formula = sysIdx < 0 ? [] : list.map(function (it, i) {
        var r = i + 2, sa = colName(sysIdx) + r, ca = colName(cntIdx) + r, v = aoa[i + 1][dIdx];
        return { a: colName(dIdx) + r, f: 'IF(' + ca + '="","",' + ca + '-' + sa + ')', v: v };
      });
      return { name: name, aoa: aoa, formula: formula };
    }
    var sheets = [{ name: '지시정보', aoa: orderInfoAoa(session, items, assign, opts) }, sheet(MAIN_SHEET, sorted)];
    if (!opts.only) assignedNames(assign).forEach(function (n) {
      sheets.push(sheet(safeSheetName(n, sheets), sorted.filter(function (it) { return assigneeOf(it, session, assign) === n; })));
    });
    return sheets;
  }
  /* 엑셀 칸 서식 — 10-28 수강생 요청: 수량·금액은 천 단위 쉼표(12,000), 날짜는 yyyy-mm-dd.
     ws: SheetJS 시트 객체(aoa_to_sheet 결과). 숫자 칸은 숫자 그대로 두고 표시 서식(z)만 붙입니다(엑셀에서 계산 · 정렬 그대로).
     'YYYY-MM-DD' 글자는 엑셀 날짜 칸(일련번호 + yyyy-mm-dd 서식)으로 바꿉니다. 수식 칸(차이수량)도 쉼표 서식. */
  var NUM_FMT = '#,##0', NUM_FMT_DEC = '#,##0.###', DATE_FMT = 'yyyy-mm-dd';
  function formatCells(ws) {
    Object.keys(ws || {}).forEach(function (a) {
      if (a.charAt(0) === '!') return;
      var c = ws[a];
      if (!c || c.z) return;
      if (c.f) { c.z = typeof c.v === 'number' && !Number.isInteger(c.v) ? NUM_FMT_DEC : NUM_FMT; return; }
      if (c.t === 'n' && isFinite(c.v)) { c.z = Number.isInteger(c.v) ? NUM_FMT : NUM_FMT_DEC; return; }
      if (c.t === 's' && /^\d{4}-\d\d-\d\d$/.test(c.v)) {
        var t = Date.UTC(+c.v.slice(0, 4), +c.v.slice(5, 7) - 1, +c.v.slice(8, 10));
        if (!isNaN(t)) { c.t = 'n'; c.v = t / 86400000 + 25569; c.z = DATE_FMT; delete c.w; }
      }
    });
    return ws;
  }
  function sortForWalk(items) {
    // 현장에서 걷는 순서: 창고 → 보관 자리 → 재고번호
    return items.slice().sort(function (a, b) { return cmp(a.wh, b.wh) || cmp(a.loc, b.loc) || cmp(a.no, b.no); });
  }
  function safeSheetName(n, sheets) {
    var base = String(n).replace(/[\\\/?*\[\]:]/g, '').slice(0, 28) || '시트', name = base, i = 2;
    while (sheets.some(function (s) { return s.name === name; })) name = base.slice(0, 25) + '_' + (i++);
    return name;
  }
  /* 기록서(작업지시서 · SharePoint 취합 파일) 다시 읽기 — 휴대폰에 지시서를 넣을 때, 여러 사람이 채운 취합 파일을 합칠 때.
     sheets: [{name, aoa}] → { session, items, assign, records, reasons, warnings } 또는 { error }
     · 「지시정보」 외 모든 시트를 읽고, 같은 사람이 전체 시트와 자기 시트에 같은 품목을 적었으면 자기 시트를 씁니다 */
  function parseOrderWorkbook(sheets, opts) {
    opts = opts || {};
    var info = (sheets || []).filter(function (s) { return s.name === '지시정보'; })[0];
    if (!info) return { error: '「지시정보」 시트가 없습니다. 이 도구에서 내려받은 기록서(작업지시서) 엑셀인지 확인해 주세요.' };
    var kv = {};
    info.aoa.forEach(function (r) { if (r && str(r[0]) && kv[str(r[0])] === undefined) kv[str(r[0])] = r[1]; });
    var baseAt = toDateTime(kv['기준 일자·시각']);
    if (!str(kv['지시서 ID']) || !baseAt) return { error: '지시정보에 지시서 ID·기준 일자·시각이 없습니다.' };
    var session = {
      id: str(kv['지시서 ID']), title: str(kv['실사명']), baseAt: baseAt, createdAt: toDateTime(kv['작성 시각']) || '',
      sourceFile: str(kv['ERP 파일']), rosterFile: str(kv['배정표 파일']), unitMode: str(kv['배정 단위']) === '창고·구역' ? 'zone' : 'wh', orderVer: Number(kv['지시서 버전']) || 1,
      cols: [], managers: {}, excluded: []
    };
    var reasons = str(kv['사유 구분']).split(/\s*\/\s*/).filter(Boolean);
    var main = sheets.filter(function (s) { return s.name === MAIN_SHEET || s.name === '작업지시서'; })[0];
    if (!main) return { error: '「' + MAIN_SHEET + '」 시트가 없습니다.' };
    var p = parseErp(main.aoa);
    if (p.det.missing.length) return { error: MAIN_SHEET + ' 시트의 열을 읽지 못했습니다.' };
    session.cols = p.cols.map(function (x) { return { h: x.h, f: x.f }; });
    var aCol = p.det.map.assignee, mCol = p.det.map.manager, names = {};
    p.items.forEach(function (it) {
      var row = main.aoa[it.row - 1] || [];
      var a = aCol == null ? '' : str(row[aCol]), mg = mCol == null ? '' : str(row[mCol]);
      if (mg && !session.managers[it.wh]) session.managers[it.wh] = mg;
      it.asg = a;
      if (a) { var l = names[it.wh] = names[it.wh] || []; if (l.indexOf(a) < 0) l.push(a); }
    });
    var assign = {};
    p.items.forEach(function (it) {
      var l = names[it.wh] || [];
      if (session.unitMode !== 'zone' && l.length > 1 && it.asg) it.part = l.indexOf(it.asg) + 1;
      if (it.asg) assign[unitKey(it, session.unitMode)] = it.asg;
      delete it.asg;
    });
    var now = opts.now || localIso(), warnings = p.warnings.filter(function (w) { return !/단가·금액|재고번호 열이 없어/.test(w); });
    var cOpts = { sessionId: session.id, fileName: opts.fileName, items: p.items, reasons: reasons, now: now };
    var mainRec = parseCountSheet(main.aoa, Object.assign({ sheet: main.name }, cOpts));
    warnings = warnings.concat(mainRec.warnings);
    var own = [];
    sheets.forEach(function (s) {
      if (s === main || s.name === '지시정보') return;
      var r2 = parseCountSheet(s.aoa, Object.assign({ sheet: s.name }, cOpts));
      if (r2.det.missing.length) return; // 기록서 모양이 아닌 시트(메모 등)는 건너뜀
      own = mergeRecords(own, r2.records).list;
      warnings = warnings.concat(r2.warnings);
    });
    var ownBy = {}; own.forEach(function (r) { ownBy[r.no + '|' + r.by] = r; });
    var differ = [];
    var kept = mainRec.records.filter(function (r) {
      var o = ownBy[r.no + '|' + r.by]; if (!o) return true;
      if (o.qty !== r.qty) differ.push(r.no + '(' + r.by + ': 전체 ' + r.qty + ' / 자기 시트 ' + o.qty + ')');
      return false;
    });
    if (differ.length) warnings.push('같은 사람이 전체 시트와 자기 시트에 다르게 적은 품목 ' + differ.length + '개 — 자기 시트 값을 씁니다: ' + differ.slice(0, 3).join(', ') + (differ.length > 3 ? ' …' : ''));
    var records = mergeRecords(kept, own).list;
    return {
      session: session, reasons: reasons,
      items: p.items.map(function (it) { var o = {}; Object.keys(it).forEach(function (k) { o[k] = it[k]; }); o.amt = it.sys * it.price; return o; }),
      assign: assign, records: records, warnings: warnings
    };
  }

  /* ── 5. 현장 기록 ───────────────────────────────────────────── */
  /* 사유 글자 맞추기 — list(설정의 사유 목록, 없으면 기본 6개) 안에서. 기본 사유의 별칭도 씁니다.
     못 맞추면 목록에 「기타」가 있으면 「기타」, 없으면 목록의 마지막 사유 */
  function matchReason(text, list) {
    var t = str(text);
    if (!t) return '';
    list = list && list.length ? list : REASONS;
    if (list.indexOf(t) >= 0) return t;
    var n = norm(t);
    function al(r) { return REASON_ALIASES[r] || []; }
    for (var i = 0; i < list.length; i++) if (norm(list[i]) === n || al(list[i]).indexOf(n) >= 0) return list[i];
    for (var j = 0; j < list.length; j++) {
      var rr = list[j], nr = norm(rr);
      if (al(rr).some(function (a) { return a.length >= 2 && n.indexOf(a) >= 0; }) || (nr.length >= 2 && n.indexOf(nr) >= 0)) return rr;
    }
    return list.indexOf('기타') >= 0 ? '기타' : list[list.length - 1];
  }
  var seq = 0;
  function newRid(dev) { seq = (seq + 1) % 1679616; return (dev || 'dev') + '-' + Date.now().toString(36) + '-' + seq.toString(36) + Math.random().toString(36).slice(2, 5); }
  /* 현장 입력 하나 점검. input: { no, qty, reason, note, by, at?, dev?, photo?, extra? }, item: 재고 품목(없으면 목록에 없는 자재)
     반환: { ok, errors:{칸:문장}, warnings:[], record } */
  function validateCount(input, item, opts) {
    opts = opts || {};
    var errors = {}, warnings = [];
    var qty = toNum(input.qty);
    if (qty === null) errors.qty = '실사 수량을 넣어 주세요(없으면 0).';
    else if (isNaN(qty)) errors.qty = '숫자로 넣어 주세요.';
    else if (qty < 0) errors.qty = '0 이상으로 넣어 주세요.';
    var by = str(input.by);
    if (!by) errors.by = '작성자 이름을 넣어 주세요(처음 한 번만).';
    var extra = null;
    if (!item) {
      var x = input.extra || {};
      if (!str(x.code) && !str(x.name)) errors.extra = '목록에 없는 자재는 자재코드나 자재명을 적어 주세요.';
      if (!str(x.wh)) errors.wh = '창고를 골라 주세요.';
      extra = { code: str(x.code), name: str(x.name), wh: str(x.wh), loc: str(x.loc), unit: str(x.unit) };
    }
    var sys = item ? item.sys : 0;
    var reason = input.reason ? matchReason(input.reason, opts.reasons) : '';
    var diff = qty != null && !isNaN(qty) ? round(qty - sys, 6) : null;
    if (diff != null && diff !== 0 && !reason) errors.reason = '시스템 수량과 ' + fmtSigned(diff) + ' 차이가 있습니다. 사유를 골라 주세요.';
    if (diff === 0 && reason) warnings.push('수량이 같아 사유는 참고로만 남깁니다.');
    if (reason === '기타' && !str(input.note)) errors.note = '「기타」는 특이사항에 내용을 적어 주세요.';
    if (item && qty != null && !isNaN(qty) && sys > 0 && Math.abs(diff) / sys >= 0.5 && Math.abs(diff) >= 10) warnings.push('차이가 시스템 수량의 ' + pctText(Math.abs(diff) / sys, 0) + ' 입니다. 단위(EA·BOX)를 한 번 더 확인해 주세요.');
    var ok = !Object.keys(errors).length;
    var no = item ? item.no : 'NEW-' + searchKey(extra && (extra.code || extra.name)).toUpperCase().slice(0, 20) + '@' + (extra ? extra.wh : '');
    var rec = ok ? {
      rid: input.rid || newRid(input.dev), no: no, qty: qty, reason: diff === 0 ? (reason || '') : reason, note: str(input.note),
      by: by, at: input.at || localIso(), dev: str(input.dev), src: str(input.src) || '현장 입력', photo: input.photo || null, sessionId: str(opts.sessionId)
    } : null;
    if (rec && extra) rec.extra = extra;
    return { ok: ok, errors: errors, warnings: warnings, record: rec, diff: diff };
  }
  /* 종이 기록을 옮긴 엑셀(작업지시서 + 실사 수량 칸) → 기록. 실사 수량이 빈 줄은 아직 안 센 것으로 둡니다. */
  function parseCountSheet(aoa, opts) {
    opts = opts || {};
    var det = detectColumns(aoa, opts.overrides, 'count');
    var out = { det: det, records: [], warnings: [], blank: 0, unknown: [] };
    if (det.missing.length) return out;
    var m = det.map, byNo = {}, keyOf = rowKeyer();
    (opts.items || []).forEach(function (it) { byNo[it.no] = it; });
    function cell(row, f) { return m[f] == null ? null : row[m[f]]; }
    var noReason = 0, noBy = 0;
    for (var r = det.headerRow + 1; r < aoa.length; r++) {
      var row = aoa[r] || [];
      if (isBlankRow(row) || isSummaryRow(row)) continue;
      var no = m.no != null ? str(cell(row, 'no')) : '';
      if (!no && m.code != null) { var code = str(cell(row, 'code')); if (code) no = makeKey(code, str(cell(row, 'wh')), str(cell(row, 'loc')), str(cell(row, 'batch'))); }
      if (!no) continue;
      no = keyOf(no);
      var cnt = toNum(cell(row, 'cnt'));
      if (cnt === null) { out.blank++; continue; }
      if (isNaN(cnt) || cnt < 0) { out.warnings.push((r + 1) + '행 ' + no + ': 실사 수량을 읽지 못했습니다(' + str(cell(row, 'cnt')) + ').'); continue; }
      if (opts.items && !byNo[no]) { out.unknown.push(no); continue; }
      var by = str(cell(row, 'by')) || str(cell(row, 'assignee')) || '';
      if (!by) { noBy++; by = '(확인자 미기재)'; }
      var at = toDateTime(cell(row, 'at')) || '';
      var reason = matchReason(cell(row, 'reason'), opts.reasons);
      var it = byNo[no];
      if (it && round(cnt - it.sys, 6) !== 0 && !reason) { noReason++; }
      var src = '엑셀: ' + (opts.fileName || '종이 기록') + (opts.sheet ? ' · ' + opts.sheet + ' 시트' : '');
      // 기록 ID 는 내용으로 만듭니다 — 같은 파일을 여러 번 넣거나(SharePoint 에서 하루에 몇 번 내려받아도) 한 번만,
      // 수량·사유를 고친 줄은 새 기록(정정)이 됩니다. 확인시각이 빈 줄은 처음 넣은 시각(imp)으로 순서를 정합니다.
      out.records.push({
        rid: 'xl:' + (opts.sessionId || '') + '|' + no + '|' + by + '|' + at + '|' + cnt + '|' + reason, no: no, qty: cnt, reason: reason, note: str(cell(row, 'note')),
        by: by, at: at, imp: opts.now || '', dev: 'excel', src: src, photo: null, sessionId: str(opts.sessionId)
      });
    }
    if (noReason) out.warnings.push('차이가 있는데 사유가 빈 줄 ' + noReason + '개 — 「' + NO_REASON + '」으로 집계됩니다.');
    if (noBy) out.warnings.push('확인자·실사자가 빈 줄 ' + noBy + '개 — 「(확인자 미기재)」로 둡니다.');
    if (out.unknown.length) out.warnings.push('작업지시서에 없는 재고번호 ' + out.unknown.length + '개는 건너뜁니다: ' + out.unknown.slice(0, 5).join(', ') + (out.unknown.length > 5 ? ' …' : ''));
    return out;
  }

  /* ── 6. 결과 파일 · 합치기 · 충돌 ───────────────────────────── */
  /* 담당자 휴대폰에서 내보내는 결과 파일(JSON). 사진은 넣을지 고를 수 있습니다(파일이 커짐). */
  function resultFile(session, records, meta) {
    meta = meta || {};
    return {
      kind: RESULT_KIND, schema: SCHEMA_VERSION, sessionId: session.id, sessionTitle: session.title, baseAt: session.baseAt,
      by: str(meta.by), dev: str(meta.dev), exportedAt: meta.now || localIso(),
      records: (records || []).map(function (r) { var o = {}; Object.keys(r).forEach(function (k) { o[k] = r[k]; }); if (!meta.withPhoto) o.photo = r.photo ? 'omitted' : null; return o; })
    };
  }
  /* 가져온 결과 파일 점검 — 다른 실사 회차의 파일이면 합치지 않습니다 */
  function checkResultFile(obj, session, items, fileName) {
    if (!obj || obj.kind !== RESULT_KIND || !Array.isArray(obj.records)) return { error: '이 도구의 결과 파일이 아닙니다.' };
    if (!session) return { error: '먼저 작업지시(실사 회차)를 만들거나 작업지시서를 넣어 주세요.' };
    if (obj.sessionId !== session.id) return { error: '다른 실사 회차의 결과 파일입니다(파일: ' + (obj.sessionTitle || obj.sessionId) + ' / 지금: ' + session.title + ').' };
    var byNo = {}; (items || []).forEach(function (it) { byNo[it.no] = 1; });
    var ok = [], unknown = [];
    obj.records.forEach(function (r) {
      if (!r || !r.rid || !r.no || typeof r.qty !== 'number') return;
      if (!byNo[r.no] && !r.extra) { unknown.push(r.no); return; }
      var o = {}; Object.keys(r).forEach(function (k) { o[k] = r[k]; });
      if (fileName && !o.file) o.file = fileName; // 원천 기록 연결: 어느 결과 파일로 들어왔는지
      ok.push(o);
    });
    return { records: ok, unknown: unknown, by: obj.by, exportedAt: obj.exportedAt };
  }
  /* 기록 합치기 — 같은 기록(rid)은 한 번만. 사진이 빠진 사본('omitted')이 사진 있는 기록을 덮지 않게 합니다. */
  /* 같은 내용(재고번호·작성자·수량·사유·확인 시각 분 단위)도 같은 기록으로 봅니다 — 휴대폰 기록을 기록서 엑셀로 내보냈다가
     SharePoint 취합 파일로 다시 받아도 두 번 들어가지 않게. 확인시각이 빈 엑셀 줄은 시각을 빼고 견줍니다. */
  function contentKey(r, withAt) { return r.no + '|' + r.by + '|' + r.qty + '|' + (r.reason || '') + (withAt ? '|' + str(r.at).slice(0, 16) : ''); }
  function mergeRecords(existing, incoming) {
    var list = (existing || []).slice(), at = {}, ck = {}, loose = {}, added = 0, dup = 0;
    function index(r, i) { at[r.rid] = i; ck[contentKey(r, true)] = i; if (loose[contentKey(r, false)] == null) loose[contentKey(r, false)] = i; }
    list.forEach(index);
    (incoming || []).forEach(function (r) {
      var i = at[r.rid];
      if (i == null) i = ck[contentKey(r, true)];
      if (i == null && !r.at) i = loose[contentKey(r, false)];
      if (i != null) {
        dup++;
        var cur = list[i];
        if ((!cur.photo || cur.photo === 'omitted') && r.photo && r.photo !== 'omitted') { list[i] = r; index(r, i); }
        return;
      }
      index(r, list.length); list.push(r); added++;
    });
    return { list: list, added: added, dup: dup };
  }
  function byAt(a, b) { return cmp(a.at || a.imp, b.at || b.imp) || cmp(a.rid, b.rid); }
  /* 품목마다 지금 채택한 기록을 정합니다.
     · 같은 사람이 다시 센 것은 정정 → 그 사람의 마지막 기록
     · 서로 다른 사람의 마지막 기록끼리 수량이 다르면 「충돌」 — picks[재고번호] = rid 로 고르기 전까지 표시
     · 충돌이 없으면 가장 늦은 기록 */
  function resolveCounts(records, picks) {
    var by = {};
    (records || []).forEach(function (r) { (by[r.no] = by[r.no] || []).push(r); });
    var out = {};
    Object.keys(by).forEach(function (no) {
      var recs = by[no].sort(byAt), last = {};
      recs.forEach(function (r) { last[r.by] = r; });
      var cands = Object.keys(last).map(function (k) { return last[k]; }).sort(byAt);
      var qtys = {}; cands.forEach(function (r) { qtys[String(r.qty)] = 1; });
      var differ = Object.keys(qtys).length > 1;
      var pick = picks && picks[no] ? recs.filter(function (r) { return r.rid === picks[no]; })[0] : null;
      out[no] = {
        no: no, recs: recs, candidates: cands, current: pick || recs[recs.length - 1],
        conflict: differ && !pick, resolved: differ && !!pick, corrected: recs.length > cands.length
      };
    });
    return out;
  }
  function conflictList(resolved) { return Object.keys(resolved).map(function (k) { return resolved[k]; }).filter(function (x) { return x.conflict; }); }

  /* ── 7. 차이 대사 ──────────────────────────────────────────── */
  function cleanSettings(p) {
    var s = {}; p = p || {};
    ['tolPct', 'tolAmt', 'topN'].forEach(function (k) { var v = Number(p[k]); s[k] = p[k] != null && p[k] !== '' && isFinite(v) && v >= 0 ? v : DEFAULT_SETTINGS[k]; });
    s.tolMode = p.tolMode === 'range' ? 'range' : 'zero';
    if (Array.isArray(p.reasons) && p.reasons.some(function (x) { return str(x); })) s.reasons = reasonList(p);
    return s;
  }
  /* 목록에 없는 자재(현장 발견)를 품목으로 만듭니다 — 시스템 수량 0, 단가 모름 */
  function extraItems(records, items) {
    var have = {}; (items || []).forEach(function (it) { have[it.no] = 1; });
    var out = {}, list = [];
    (records || []).forEach(function (r) {
      if (!r.extra || have[r.no] || out[r.no]) return;
      var same = (items || []).filter(function (it) { return it.wh === r.extra.wh; })[0] || {};
      out[r.no] = { no: r.no, code: r.extra.code, name: r.extra.name, spec: '', unit: r.extra.unit, wh: r.extra.wh, whName: same.whName || '', site: same.site || '', loc: r.extra.loc, mtype: '(목록에 없음)', mkind: '(목록에 없음)', sys: 0, price: 0, amt: 0, extra: true };
      list.push(out[r.no]);
    });
    return list;
  }
  /* 품목마다 한 줄: 시스템 vs 실사, 수량·금액 차이, 차이율, 허용 기준 판정, 원천 기록 */
  function varianceRows(db) {
    var S = cleanSettings(db.settings), session = db.session;
    var items = (db.items || []).concat(extraItems(db.records, db.items));
    var res = resolveCounts(db.records, db.picks);
    return items.map(function (it) {
      var rs = res[it.no], c = rs ? rs.current : null;
      var row = {
        no: it.no, code: it.code, name: it.name, spec: it.spec, unit: it.unit, wh: whLabel(it), whCode: it.wh, whName: it.whName || it.wh, site: it.site || '(빈칸)', loc: it.loc, zone: whLabel(it) + ' / ' + zoneOf(it.loc),
        manager: managerOf(it, session),
        mtype: it.mtype || '(빈칸)', mkind: it.mkind || '(빈칸)', sys: it.sys, price: it.price, sysAmt: it.sys * it.price,
        assignee: it.extra ? (c ? c.by : '') : (assigneeOf(it, session, db.assign) || '(미배정)'), extra: !!it.extra,
        counted: !!c, cnt: null, diff: null, diffAmt: null, rate: null, status: '미실사', over: false, reason: '', note: '', by: '', at: '', rid: '', src: '', photo: false, photoSrc: '', photoOld: false, photoOmitted: false,
        recCount: rs ? rs.recs.length : 0, conflict: !!(rs && rs.conflict), resolved: !!(rs && rs.resolved), corrected: !!(rs && rs.corrected)
      };
      if (!c) return row;
      row.cnt = c.qty; row.diff = round(c.qty - it.sys, 6); row.diffAmt = row.diff * it.price;
      row.rate = it.sys ? row.diff / it.sys : (row.diff ? null : 0);
      row.reason = row.diff !== 0 ? (c.reason || NO_REASON) : (c.reason || '');
      row.note = c.note || ''; row.by = c.by; row.at = c.at; row.rid = c.rid; row.src = (c.src || '') + (c.file ? ' · ' + c.file : ''); row.photo = !!(c.photo);
      // 11-08 수강생 TEST: 「사진」 표시만 있고 볼 수 없었음 → 볼 수 있는 사진(data:image)을 함께 넘김.
      // 채택한 기록에 사진이 없으면 같은 품목의 이전 기록 사진을, 결과 파일에 사진이 빠진 것('omitted')은 따로 알림.
      var ph = isPhoto(c.photo) ? c : rs.recs.slice().reverse().filter(function (x) { return isPhoto(x.photo); })[0];
      if (ph) { row.photo = true; row.photoSrc = ph.photo; row.photoOld = ph !== c; }
      else if (c.photo === 'omitted') row.photoOmitted = true;
      if (row.diff === 0) row.status = '일치';
      else if (S.tolMode !== 'range') { row.over = true; row.status = '차이(사유 필수)'; }
      else {
        row.over = it.extra || (row.rate == null) || Math.abs(row.rate) * 100 > S.tolPct || Math.abs(row.diffAmt) > S.tolAmt;
        row.status = row.over ? '허용 초과' : '허용 안 차이';
      }
      return row;
    });
  }

  function isPhoto(p) { return typeof p === 'string' && /^data:image\//.test(p); }

  /* ── 8. 집계 (BI) ───────────────────────────────────────────── */
  function emptyAgg(key) { return { key: key, items: 0, counted: 0, matched: 0, diffItems: 0, overItems: 0, sysAmt: 0, cntAmt: 0, plusAmt: 0, minusAmt: 0, netAmt: 0, absAmt: 0, plusQty: 0, minusQty: 0, conflicts: 0, noReason: 0 }; }
  function addRow(a, r) {
    a.items++; a.sysAmt += r.sysAmt; if (r.conflict) a.conflicts++;
    if (!r.counted) return;
    a.counted++; a.cntAmt += r.cnt * r.price;
    if (r.diff === 0) { a.matched++; return; }
    a.diffItems++; if (r.over) a.overItems++; if (r.reason === NO_REASON) a.noReason++;
    if (r.diffAmt > 0) a.plusAmt += r.diffAmt; else a.minusAmt += r.diffAmt;
    if (r.diff > 0) a.plusQty += r.diff; else a.minusQty += r.diff;
    a.netAmt += r.diffAmt; a.absAmt += Math.abs(r.diffAmt);
  }
  function finish(a) {
    a.progress = a.items ? a.counted / a.items : 0;
    a.accuracy = a.counted ? a.matched / a.counted : null;
    a.amtAccuracy = a.sysAmt ? 1 - a.absAmt / a.sysAmt : null;
    return a;
  }
  function overall(rows) { var a = emptyAgg('전체'); rows.forEach(function (r) { addRow(a, r); }); return finish(a); }
  /* dim: wh · zone · mtype · mkind · assignee */
  function summarize(rows, dim) {
    var by = {}, list = [];
    rows.forEach(function (r) { var k = r[dim] || '(빈칸)'; if (!by[k]) { by[k] = emptyAgg(k); by[k].label = dim === 'wh' ? r.whName || k : k; list.push(by[k]); } addRow(by[k], r); });
    list.forEach(finish);
    uniqueLabels(list);
    return list.sort(function (a, b) { return b.absAmt - a.absAmt || b.items - a.items || cmp(a.key, b.key); });
  }
  /* 그래프 눈금 이름이 겹치면(창고 이름만 같은 두 창고) 코드까지 붙인 원래 이름으로 */
  function uniqueLabels(list) {
    var n = {}; list.forEach(function (a) { n[a.label] = (n[a.label] || 0) + 1; });
    list.forEach(function (a) { if (n[a.label] > 1) a.label = a.key; });
    return list;
  }
  /* 실사자별 — 11-08 수강생 TEST: 결과 파일의 By(작성자)가 다른 사람인데 배정받은 사람 쪽에 셈해졌음.
     배정 품목 · 배정 진행률은 배정표 기준, 직접 센 품목 · 일치 · 차이 · 금액은 채택한 기록의 작성자(By) 기준으로 나눕니다.
     items = 배정 품목, done = 배정 중 확인된 품목(누가 세었든), progress = done ÷ items,
     counted = 이 사람이 직접 센 품목(배정과 무관), byOthers = 배정 중 다른 사람이 센 품목, forOthers = 남의 배정을 대신 센 품목 */
  function personStats(rows) {
    var by = {}, list = [];
    function get(k) { k = k || '(빈칸)'; if (!by[k]) { by[k] = emptyAgg(k); by[k].label = k; by[k].done = 0; by[k].byOthers = 0; by[k].forOthers = 0; list.push(by[k]); } return by[k]; }
    rows.forEach(function (r) {
      var own = r.assignee || '(미배정)', who = r.counted ? (r.by || own) : '';
      var a = get(own);
      a.items++; a.sysAmt += r.sysAmt; if (r.conflict) a.conflicts++;
      if (r.counted) { a.done++; if (who !== own) a.byOthers++; }
      if (!r.counted) return;
      var b = get(who);
      if (who !== own) b.forOthers++;
      b.counted++; b.cntAmt += r.cnt * r.price;
      if (r.diff === 0) { b.matched++; return; }
      b.diffItems++; if (r.over) b.overItems++; if (r.reason === NO_REASON) b.noReason++;
      if (r.diffAmt > 0) b.plusAmt += r.diffAmt; else b.minusAmt += r.diffAmt;
      if (r.diff > 0) b.plusQty += r.diff; else b.minusQty += r.diff;
      b.netAmt += r.diffAmt; b.absAmt += Math.abs(r.diffAmt);
    });
    list.forEach(function (a) {
      a.progress = a.items ? a.done / a.items : null;
      a.accuracy = a.counted ? a.matched / a.counted : null;
      a.amtAccuracy = a.sysAmt ? 1 - a.absAmt / a.sysAmt : null;
    });
    return list.sort(function (a, b) { return b.absAmt - a.absAmt || b.items - a.items || cmp(a.key, b.key); });
  }
  /* 분석(BI) 그래프에 처음 그릴 행 — 절대 차이 금액이 큰 n 개(차이가 없으면 품목 수 순) */
  function defaultPicks(list, n) {
    n = n == null ? 8 : n;
    return (list || []).slice().sort(function (a, b) { return b.absAmt - a.absAmt || b.items - a.items || cmp(a.key, b.key); }).slice(0, n).map(function (a) { return a.key; });
  }
  /* 사유별 — 차이가 있는 품목만. 사유가 빈 차이는 「사유 미입력」 */
  function reasonStats(rows, reasons) {
    reasons = reasons && reasons.length ? reasons : REASONS;
    function mk(k) { by[k] = { key: k, items: 0, plusAmt: 0, minusAmt: 0, netAmt: 0, absAmt: 0 }; return by[k]; }
    var by = {}, list = reasons.concat([NO_REASON]).map(mk);
    var total = 0;
    rows.forEach(function (r) {
      if (!r.counted || r.diff === 0) return;
      var a = by[r.reason];
      if (!a) { a = mk(r.reason); list.splice(list.length - 1, 0, a); } // 목록을 바꾸기 전에 적힌 사유도 따로 셈
      a.items++; total++;
      if (r.diffAmt > 0) a.plusAmt += r.diffAmt; else a.minusAmt += r.diffAmt;
      a.netAmt += r.diffAmt; a.absAmt += Math.abs(r.diffAmt);
    });
    var absAll = list.reduce(function (s, a) { return s + a.absAmt; }, 0);
    list.forEach(function (a) { a.share = total ? a.items / total : 0; a.amtShare = absAll ? a.absAmt / absAll : 0; });
    return list.filter(function (a) { return a.items || reasons.indexOf(a.key) >= 0; });
  }
  function topDiffs(rows, n) {
    return rows.filter(function (r) { return r.counted && r.diff !== 0; })
      .sort(function (a, b) { return Math.abs(b.diffAmt) - Math.abs(a.diffAmt) || Math.abs(b.diff) - Math.abs(a.diff) || cmp(a.no, b.no); })
      .slice(0, n == null ? 15 : n);
  }
  /* 시간대별 진행(확인 시각 기준, 시 단위) — 실사 당일 진척 보기 */
  function progressByHour(rows) {
    var by = {};
    rows.forEach(function (r) { if (r.counted && r.at) { var k = r.at.slice(0, 13); by[k] = (by[k] || 0) + 1; } });
    var keys = Object.keys(by).sort(), cum = 0;
    return keys.map(function (k) { cum += by[k]; return { key: k, label: k.slice(0, 10) + ' ' + k.slice(11, 13) + '시', n: by[k], cum: cum }; });
  }
  function hasValues(rows, k) { return rows.some(function (r) { return r[k] && r[k] !== '(빈칸)' && r[k] !== '(목록에 없음)'; }); }
  function analyze(db) {
    var rows = varianceRows(db), S = cleanSettings(db.settings), session = db.session;
    return {
      rows: rows, S: S, session: db.session, all: overall(rows),
      dims: { wh: summarize(rows, 'wh'), site: summarize(rows, 'site'), mtype: summarize(rows, 'mtype'), mkind: summarize(rows, 'mkind'), assignee: personStats(rows) },
      // 파일에 그 열이 없어 모두 「(빈칸)」인 기준은 화면·보고서에서 숨깁니다(예: 견본 ERP 에는 자재유형·자재타입이 없음)
      hasDim: { wh: true, assignee: true, site: hasValues(rows, 'site'), mtype: hasValues(rows, 'mtype'), mkind: hasValues(rows, 'mkind') },
      reasonList: reasonList(S), excluded: (session && session.excluded) || [],
      reasons: reasonStats(rows, reasonList(S)), top: topDiffs(rows, S.topN), conflicts: rows.filter(function (r) { return r.conflict; }), hours: progressByHour(rows),
      records: (db.records || []).length, files: (db.files || []).slice()
    };
  }

  /* ── 9. 보고서 (버전 · 원천 기록 연결) ───────────────────────── */
  /* 보고서를 낼 때마다 버전을 하나 올리고, 그때의 기준 시각·기록 수·결과 파일을 함께 남깁니다 */
  function nextReportMeta(db, opts) {
    opts = opts || {};
    var prev = db.reports || [];
    var an = opts.an || analyze(db);
    return {
      ver: prev.length ? prev[prev.length - 1].ver + 1 : 1, issuedAt: opts.now || localIso(), by: str(opts.by),
      sessionId: db.session ? db.session.id : '', baseAt: db.session ? db.session.baseAt : '',
      records: an.records, counted: an.all.counted, items: an.all.items, netAmt: round(an.all.netAmt, 0), files: (db.files || []).map(function (f) { return f.name; })
    };
  }
  function verText(meta) { return meta ? 'v' + meta.ver + ' (' + fmtTime(meta.issuedAt) + ' 발행)' : '초안(아직 발행 안 함)'; }
  var DIM_SHEETS = [['wh', '창고별'], ['site', '창고위치별'], ['mtype', '자재유형별'], ['mkind', '자재타입별'], ['assignee', '실사자별']];
  function aggHead(label, person) { return person ? [label, '배정 품목', '직접 센 품목', '배정 진행률(%)', '일치', '정확도(%)', '차이 품목', '허용 초과', '배정 장부금액', '증가 금액', '감소 금액', '순차이 금액', '절대 차이 금액'] : [label, '품목 수', '실사 완료', '진행률(%)', '일치', '정확도(%)', '차이 품목', '허용 초과', '장부금액', '증가 금액', '감소 금액', '순차이 금액', '절대 차이 금액']; }
  function aggRow(a) {
    return [a.key, a.items, a.counted, a.progress == null ? '' : round(a.progress * 100, 1), a.matched, a.accuracy == null ? '' : round(a.accuracy * 100, 1), a.diffItems, a.overItems,
      Math.round(a.sysAmt), Math.round(a.plusAmt), Math.round(a.minusAmt), Math.round(a.netAmt), Math.round(a.absAmt)];
  }
  var DETAIL_HEAD = ['재고번호', '창고', '창고위치', '보관 자리', '자재코드', '자재명', '규격', '단위', '자재유형', '자재타입', '담당자', '실사자', '시스템수량', '실사수량', '수량 차이', '단가', '차이 금액', '차이율(%)', '판정', '사유', '특이사항',
    '확인자', '확인 시각', '기록 ID', '원천(입력 경로)', '사진', '같은 품목 기록 수', '충돌·정정'];
  function detailRow(r) {
    return [r.no, r.wh, r.site === '(빈칸)' ? '' : r.site, r.loc, r.code, r.name, r.spec, r.unit, r.mtype, r.mkind, r.manager || '', r.assignee, r.sys, r.cnt == null ? '' : r.cnt, r.diff == null ? '' : r.diff, round(r.price, 2),
      r.diffAmt == null ? '' : Math.round(r.diffAmt), r.rate == null ? (r.counted && r.diff ? '신규' : '') : round(r.rate * 100, 2), r.status, r.reason, r.note,
      r.by, r.at ? fmtTime(r.at) : '', r.rid, r.src, r.photo ? '있음' : '', r.recCount || '',
      r.conflict ? '충돌(미해결)' : r.resolved ? '충돌 → 선택함' : r.corrected ? '같은 사람이 다시 셈' : ''];
  }
  /* 엑셀 결과서 시트들 — [{name, aoa}] */
  function reportSheets(db, an, meta) {
    an = an || analyze(db);
    var s = db.session || {}, a = an.all;
    var info = [
      ['재고실사 작업결과서'],
      ['보고서 버전', verText(meta)],
      ['실사명', s.title || ''], ['지시서 ID', s.id || ''], ['기준 일자·시각(ERP 재고)', fmtTime(s.baseAt)], ['ERP 파일', s.sourceFile || ''],
      ['품목 수', a.items], ['실사 완료', a.counted], ['진행률(%)', round(a.progress * 100, 1)], ['일치 품목', a.matched], ['정확도(%)', a.accuracy == null ? '' : round(a.accuracy * 100, 1)],
      ['차이 품목', a.diffItems], ['허용 초과', a.overItems], ['장부금액', Math.round(a.sysAmt)], ['순차이 금액', Math.round(a.netAmt)], ['절대 차이 금액', Math.round(a.absAmt)],
      ['허용 기준', tolText(an.S)], ['사유 구분', an.reasonList.join(' / ')],
      ['미해결 충돌', an.conflicts.length], ['현장 기록 수', an.records],
      [],
      ['원천 기록 연결', '「차이 상세」·「전체 대사표」의 기록 ID 로 「기록 이력」 시트의 원래 입력(누가·언제·어느 경로로)을 찾을 수 있습니다.'],
      ['합친 결과 파일', (db.files || []).map(function (f) { return f.name + '(' + f.records + '건, ' + fmtTime(f.at) + ')'; }).join(' / ') || '없음(이 기기에서 입력)']
    ];
    if (an.excluded.length) { info.push([]); info.push(['이번 실사 제외 창고', '품목 수', '장부금액', '담당자', '비고']); an.excluded.forEach(function (x) { info.push([whLabel(x), x.items, Math.round(x.sysAmt), x.manager || '', x.note || '']); }); }
    if ((db.reports || []).length) { info.push([]); info.push(['발행 이력', '발행 시각', '기록 수', '실사 완료', '순차이 금액']); db.reports.forEach(function (m) { info.push(['v' + m.ver, fmtTime(m.issuedAt), m.records, m.counted, m.netAmt]); }); }
    var sheets = [{ name: '보고서 정보', aoa: info }];
    DIM_SHEETS.filter(function (d) { return an.hasDim[d[0]]; }).forEach(function (d) {
      sheets.push({ name: d[1], aoa: [aggHead(DIMS[d[0]].label, d[0] === 'assignee')].concat(an.dims[d[0]].map(aggRow)).concat([aggRow(a)]) });
    });
    sheets.push({ name: '사유별', aoa: [['사유', '차이 품목', '비중(%)', '증가 금액', '감소 금액', '순차이 금액', '절대 차이 금액', '금액 비중(%)']].concat(an.reasons.map(function (x) {
      return [x.key, x.items, round(x.share * 100, 1), Math.round(x.plusAmt), Math.round(x.minusAmt), Math.round(x.netAmt), Math.round(x.absAmt), round(x.amtShare * 100, 1)];
    })) });
    sheets.push({ name: '차이 상세', aoa: [DETAIL_HEAD].concat(an.rows.filter(function (r) { return r.counted && r.diff !== 0; }).sort(function (x, y) { return Math.abs(y.diffAmt) - Math.abs(x.diffAmt); }).map(detailRow)) });
    sheets.push({ name: '전체 대사표', aoa: [DETAIL_HEAD].concat(sortForWalk(an.rows).map(detailRow)) });
    var res = resolveCounts(db.records, db.picks);
    sheets.push({ name: '기록 이력', aoa: [['기록 ID', '재고번호', '실사수량', '사유', '특이사항', '확인자', '확인 시각', '원천(입력 경로)', '받은 파일', '사진', '채택']].concat((db.records || []).slice().sort(byAt).map(function (r) {
      var cur = res[r.no] && res[r.no].current;
      return [r.rid, r.no, r.qty, r.reason || '', r.note || '', r.by, fmtTime(r.at), r.src || '', r.file || '', r.photo ? '있음' : '', cur && cur.rid === r.rid ? '채택' : '이전 기록'];
    })) });
    return sheets;
  }
  /* 글자 보고서(복사·AI 요약 원문) */
  function buildReportText(db, an, meta, aiText) {
    an = an || analyze(db);
    var s = db.session || {}, a = an.all, L = [];
    L.push('재고실사 결과 보고 — ' + (s.title || ''));
    L.push('보고서 버전: ' + verText(meta) + ' · 기준 일자·시각(ERP 재고): ' + fmtTime(s.baseAt) + ' · 지시서 ID: ' + (s.id || ''));
    L.push('허용 기준: ' + tolText(an.S));
    if (an.excluded.length) L.push('이번 실사 제외: ' + an.excluded.map(function (x) { return whLabel(x) + '(' + fmtNum(x.items) + '품목' + (x.note ? ', ' + x.note : '') + ')'; }).join(' · '));
    L.push('');
    L.push('1. 총괄');
    L.push('- 대상 ' + fmtNum(a.items) + '품목 중 ' + fmtNum(a.counted) + '품목 실사(진행률 ' + pctText(a.progress) + '), 일치 ' + fmtNum(a.matched) + '품목(정확도 ' + pctText(a.accuracy) + ')');
    L.push('- 차이 ' + fmtNum(a.diffItems) + '품목(' + (an.S.tolMode === 'range' ? '허용 초과 ' + fmtNum(a.overItems) : '사유 미입력 ' + fmtNum(a.noReason)) + '), 증가 ' + fmtWon(a.plusAmt) + ' · 감소 ' + fmtWon(Math.abs(a.minusAmt)) + ' · 순차이 ' + fmtSigned(a.netAmt, true) + ' (장부금액 ' + fmtWon(a.sysAmt) + ')');
    if (an.conflicts.length) L.push('- 미해결 충돌 ' + an.conflicts.length + '품목 — 취합 화면에서 채택할 기록을 골라야 확정됩니다.');
    L.push('');
    L.push('2. 창고별');
    an.dims.wh.forEach(function (x) { L.push('- ' + x.key + ': ' + fmtNum(x.counted) + '/' + fmtNum(x.items) + ' 실사(' + pctText(x.progress) + '), 차이 ' + fmtNum(x.diffItems) + '품목, 순차이 ' + fmtSigned(x.netAmt, true)); });
    L.push('');
    L.push('3. 사유별 (차이 품목)');
    an.reasons.filter(function (x) { return x.items; }).sort(function (x, y) { return y.absAmt - x.absAmt; }).forEach(function (x) { L.push('- ' + x.key + ': ' + fmtNum(x.items) + '품목(' + pctText(x.share) + '), 절대 차이 ' + fmtWon(x.absAmt)); });
    L.push('');
    L.push('4. 차이 금액 상위');
    an.top.slice(0, 10).forEach(function (r, i) { L.push((i + 1) + '. ' + r.no + ' ' + (r.name || r.code) + ' [' + r.wh + '] 시스템 ' + fmtNum(r.sys) + ' → 실사 ' + fmtNum(r.cnt) + ' (' + fmtSigned(r.diff) + ', ' + fmtSigned(r.diffAmt, true) + ') 사유: ' + r.reason + ' · 기록 ' + r.rid); });
    if (aiText) { L.push(''); L.push('5. 요약 의견 (AI 초안, 검토 후 사용)'); L.push(aiText); }
    return L.join('\n');
  }

  /* ── 10. AI 프롬프트 (반자동) ───────────────────────────────── */
  /* 담당자·확인자 이름은 「담당1」 같은 별칭으로 가립니다. 자재코드·자재명은 기본으로 넣지 않습니다(창고·유형·타입·숫자만). */
  function buildAiPromptInfo(an, opts) {
    opts = opts || {};
    var mask = opts.mask !== false, map = {}, n = 0;
    function who(name) { if (!mask || !name) return name; if (!map[name]) map[name] = '담당' + (++n); return map[name]; }
    an.dims.assignee.forEach(function (x) { who(x.key); });
    var a = an.all, s = an.session || {}, L = [];
    L.push('너는 제조업 자재관리·회계감사 보조야. 아래는 재고실사 결과를 도구가 계산한 사실이야. 주어진 숫자 밖의 값은 만들지 말고, 추정은 「추정」이라고 밝혀줘.');
    L.push('요청: ① 경영진 보고용 요약 5줄 ② 차이가 몰린 곳(창고·자재유형·자재타입·사유)과 그 근거 숫자 ③ 다음 실사 전에 할 개선 조치 3가지(입출고 처리·위치 관리·라벨 등) ④ 현장에 다시 확인할 항목. 한국어 존댓말로 써줘.');
    L.push('');
    L.push('[실사 개요] ' + (s.title || '') + ' · 기준 일자·시각 ' + fmtTime(s.baseAt) + ' · 허용 기준: ' + tolText(an.S));
    L.push('[총괄] 대상 ' + fmtNum(a.items) + '품목, 실사 ' + fmtNum(a.counted) + '품목(진행률 ' + pctText(a.progress) + '), 일치 ' + fmtNum(a.matched) + '(정확도 ' + pctText(a.accuracy) + '), 차이 ' + fmtNum(a.diffItems) + '품목(허용 초과 ' + fmtNum(a.overItems) + '), 장부금액 ' + fmtWon(a.sysAmt) +
      ', 증가 ' + fmtWon(a.plusAmt) + ', 감소 ' + fmtWon(Math.abs(a.minusAmt)) + ', 순차이 ' + fmtSigned(a.netAmt, true) + (an.conflicts.length ? ', 미해결 충돌 ' + an.conflicts.length + '품목' : ''));
    function block(title, list) {
      L.push('[' + title + ']');
      list.forEach(function (x) { L.push('- ' + (title === '실사자별' ? who(x.key) : x.key) + ': 품목 ' + fmtNum(x.items) + ', 실사 ' + fmtNum(x.counted) + '(' + pctText(x.progress) + '), 정확도 ' + pctText(x.accuracy) + ', 차이 ' + fmtNum(x.diffItems) + '(허용 초과 ' + fmtNum(x.overItems) + '), 순차이 ' + fmtSigned(x.netAmt, true) + ', 절대 차이 ' + fmtWon(x.absAmt)); });
    }
    DIM_SHEETS.filter(function (d) { return an.hasDim[d[0]]; }).forEach(function (d) { block(d[1], an.dims[d[0]]); });
    L.push('[사유별]');
    an.reasons.filter(function (x) { return x.items; }).forEach(function (x) { L.push('- ' + x.key + ': ' + fmtNum(x.items) + '품목(' + pctText(x.share) + '), 증가 ' + fmtWon(x.plusAmt) + ', 감소 ' + fmtWon(Math.abs(x.minusAmt)) + ', 절대 차이 ' + fmtWon(x.absAmt)); });
    L.push('[차이 금액 상위 ' + Math.min(10, an.top.length) + '품목]');
    an.top.slice(0, 10).forEach(function (r, i) {
      L.push((i + 1) + '. ' + (opts.withItem ? (r.code || r.no) + ' ' + r.name + ' · ' : '') + r.wh + (an.hasDim.mtype ? ' · ' + r.mtype : '') + (an.hasDim.mkind ? ' · ' + r.mkind : '') + ' · 시스템 ' + fmtNum(r.sys) + ' → 실사 ' + fmtNum(r.cnt) + ' (' + fmtSigned(r.diff) + ', ' + fmtSigned(r.diffAmt, true) + ') · 사유 ' + r.reason);
    });
    var prompt = L.join('\n');
    var leaks = mask ? Object.keys(map).filter(function (name) { return name.length >= 2 && prompt.indexOf(name) >= 0; }) : [];
    return { prompt: prompt, map: map, leaks: leaks };
  }
  function unmaskText(text, map) {
    var out = String(text || '');
    Object.keys(map || {}).sort(function (a, b) { return map[b].length - map[a].length; }).forEach(function (name) {
      out = out.split(map[name]).join(name);
    });
    return out;
  }

  /* ── 11. 찾기 (현장 입력) ───────────────────────────────────── */
  function searchItems(items, q) {
    var k = searchKey(q);
    if (!k) return items.slice();
    return items.filter(function (it) { return [it.no, it.code, it.name, it.loc, it.spec].some(function (v) { return searchKey(v).indexOf(k) >= 0; }); });
  }

  /* ── 12. 저장 ───────────────────────────────────────────────── */
  function emptyDb() { return { schema: SCHEMA_VERSION, session: null, items: [], excludedItems: [], assign: {}, records: [], picks: {}, files: [], reports: [], settings: cleanSettings({}), aiNotes: '', mask: true, sample: false, pptOpts: {} }; }
  function restoreDb(o) {
    var db = emptyDb();
    if (!o || typeof o !== 'object') return db;
    if (o.session && o.session.id) db.session = o.session;
    if (Array.isArray(o.items)) db.items = o.items.filter(function (x) { return x && x.no; });
    if (o.assign && typeof o.assign === 'object') db.assign = o.assign;
    if (Array.isArray(o.records)) db.records = o.records.filter(function (r) { return r && r.rid && r.no; });
    if (o.picks && typeof o.picks === 'object') db.picks = o.picks;
    if (Array.isArray(o.excludedItems)) db.excludedItems = o.excludedItems.filter(function (x) { return x && x.no; });
    if (Array.isArray(o.files)) db.files = o.files;
    if (Array.isArray(o.reports)) db.reports = o.reports;
    db.settings = cleanSettings(o.settings);
    if (typeof o.aiNotes === 'string') db.aiNotes = o.aiNotes;
    db.mask = o.mask !== false;
    db.sample = !!o.sample;
    if (o.pptOpts && typeof o.pptOpts === 'object') db.pptOpts = {
      statusBy: ['auto', 'wh', 'site', 'mkind', 'mtype'].indexOf(o.pptOpts.statusBy) >= 0 ? o.pptOpts.statusBy : 'auto',
      diffBy: o.pptOpts.diffBy === 'site' ? 'site' : 'wh',
      rowsPerPage: Math.max(0, Math.min(40, +o.pptOpts.rowsPerPage || 0)),
      removeMemos: o.pptOpts.removeMemos !== false,
      issuesText: typeof o.pptOpts.issuesText === 'string' ? o.pptOpts.issuesText.slice(0, 4000) : ''
    };
    return db;
  }

  return {
    VERSION: VERSION,
    SCHEMA_VERSION: SCHEMA_VERSION, RESULT_KIND: RESULT_KIND, ERP_FIELDS: ERP_FIELDS, COUNT_FIELDS: COUNT_FIELDS, ROSTER_FIELDS: ROSTER_FIELDS, FIELD_BY_ID: FIELD_BY_ID, REASONS: REASONS, NO_REASON: NO_REASON, DIMS: DIMS, DEFAULT_SETTINGS: DEFAULT_SETTINGS, DETAIL_HEAD: DETAIL_HEAD,
    str: str, norm: norm, toNum: toNum, round: round, colName: colName, localIso: localIso, toDateTime: toDateTime, fmtTime: fmtTime, fmtNum: fmtNum, fmtWon: fmtWon, fmtSigned: fmtSigned, pctText: pctText,
    matchField: matchField, detectColumns: detectColumns, parseErp: parseErp, pickSheet: pickSheet, headerSignature: headerSignature, isSummaryRow: isSummaryRow,
    validBaseAt: validBaseAt, makeSession: makeSession, zoneOf: zoneOf, unitKey: unitKey, whLabel: whLabel, unitsOf: unitsOf, autoAssign: autoAssign, assigneeOf: assigneeOf, managerOf: managerOf, assignedNames: assignedNames,
    splitNames: splitNames, parseRoster: parseRoster, applyRoster: applyRoster,
    orderSheets: orderSheets, parseOrderWorkbook: parseOrderWorkbook, sortForWalk: sortForWalk, bookCols: bookCols, reasonList: reasonList, tolText: tolText, MAIN_SHEET: MAIN_SHEET, COUNT_HEAD: COUNT_HEAD, DIM_SHEETS: DIM_SHEETS,
    matchReason: matchReason, newRid: newRid, validateCount: validateCount, parseCountSheet: parseCountSheet,
    resultFile: resultFile, checkResultFile: checkResultFile, mergeRecords: mergeRecords, resolveCounts: resolveCounts, conflictList: conflictList,
    cleanSettings: cleanSettings, extraItems: extraItems, varianceRows: varianceRows, overall: overall, summarize: summarize, personStats: personStats, defaultPicks: defaultPicks, isPhoto: isPhoto, reasonStats: reasonStats, topDiffs: topDiffs, progressByHour: progressByHour, analyze: analyze,
    nextReportMeta: nextReportMeta, verText: verText, reportSheets: reportSheets, buildReportText: buildReportText, buildAiPromptInfo: buildAiPromptInfo, unmaskText: unmaskText,
    formatCells: formatCells, isDateHead: isDateHead, searchItems: searchItems, emptyDb: emptyDb, restoreDb: restoreDb
  };
});
