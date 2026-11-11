/* 재고실사 Start-End — 화면
   계산은 js/logic.js(STLogic), 그래프는 js/charts.js, 저장은 js/store.js, AI 연결은 js/ai-endpoint.js·ai-panel.js.
   뼈대(가져오기·열 연결·AI 패널·보고서 내보내기)는 data10-11, 휴대폰 입력 화면은 data10-05 의 방식을 따랐습니다. */
(function () {
  'use strict';
  var L = window.STLogic, S = window.STStore, C = window.STCharts, AIP = window.AIPanel, SMP = window.STSample;
  /* 판 맞추기 — 브라우저가 옛 logic.js 와 새 app.js 를 섞어 쓰면(Pages 캐시) 버튼과 결과가 어긋날 수 있어, 섞였으면 계산하지 않고 새로고침을 안내합니다.
     판은 scripts/bump-version.mjs 로 index.html(?v=) · logic.js 와 함께 올립니다. */
  var APP_VERSION = '2026-11-11.1';
  if (!L || L.VERSION !== APP_VERSION) {
    var m0 = document.getElementById('main') || document.body;
    m0.innerHTML = '<section class="card" id="versionMismatch"><h1>새 판이 나왔습니다 — 새로고침(Ctrl+F5)해 주세요</h1><p class="alert warn">이 브라우저가 옛 파일과 새 파일을 섞어 불러왔습니다(js/logic.js ' +
      (L && L.VERSION ? L.VERSION : '옛 판') + ' · 화면 ' + APP_VERSION + '). 이대로 계산하면 버튼과 결과가 어긋날 수 있어 멈췄습니다.</p>' +
      '<p>Windows 는 <b>Ctrl+F5</b>, Mac 은 <b>Cmd+Shift+R</b> 을 한 번 눌러 주세요. 쌓아 둔 자료는 지워지지 않습니다.</p></section>';
    return;
  }
  (function () { var ft = document.querySelector('.site-footer .inner') || document.querySelector('.site-footer'); if (!ft) return; var s = document.createElement('span'); s.className = 'ver'; s.textContent = ' · 판 ' + APP_VERSION; ft.appendChild(s); })();
  var db = S.loadDb();
  var me = S.loadMe();
  var BI_TOP = 8; // 분석(BI) 그래프에 처음 그릴 행 수(절대 차이 상위)
  var ui = { biPick: {}, pending: null, errors: [], fieldScope: '', fieldFilter: 'todo', q: '', open: '', varWh: '', varState: 'diff', varQ: '', dim: 'wh', photo: null, withPhoto: false };
  var main = document.getElementById('main');

  /* ── 작은 도구 ───────────────────────────── */
  function h(tag, attrs) {
    var e = document.createElement(tag);
    if (attrs) Object.keys(attrs).forEach(function (k) {
      var v = attrs[k];
      if (v == null || v === false) return;
      if (k === 'class') e.className = v;
      else if (k.slice(0, 2) === 'on' && typeof v === 'function') e.addEventListener(k.slice(2), v);
      else if (v === true) e.setAttribute(k, '');
      else e.setAttribute(k, v);
    });
    for (var i = 2; i < arguments.length; i++) add(e, arguments[i]);
    return e;
  }
  function add(e, c) {
    if (c == null || c === false) return;
    if (Array.isArray(c)) { c.forEach(function (x) { add(e, x); }); return; }
    e.appendChild(typeof c === 'object' ? c : document.createTextNode(String(c)));
  }
  function fld(label, control, hint, err) {
    return h('label', { class: 'field' + (err ? ' invalid' : '') }, h('span', null, label), control, err ? h('small', { class: 'err' }, err) : null, hint ? h('small', { class: 'hint' }, hint) : null);
  }
  var toastTimer = null;
  function toast(msg, err) {
    var t = document.getElementById('toast');
    t.textContent = msg; t.className = 'toast' + (err ? ' error' : ''); t.hidden = false;
    clearTimeout(toastTimer); toastTimer = setTimeout(function () { t.hidden = true; }, 3500);
  }
  function save() {
    var ok = S.saveDb(db);
    var kb = S.sizeKb();
    document.getElementById('storeNote').textContent = ok ? '이 기기에 저장됨 · ' + kb + 'KB' + (kb > 3500 ? ' (용량 거의 참 — 사진을 줄이거나 결과 파일을 내보내 주세요)' : '') : '저장 안 됨(저장소가 막혔거나 가득 참) — 결과 파일·JSON 백업을 받아 두세요';
    return ok;
  }
  var N = L.fmtNum, W = L.fmtWon, pct = function (v) { return L.pctText(v, 1); };
  function download(name, blob) {
    var a = h('a', { href: URL.createObjectURL(blob), download: name });
    document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  }
  function copyText(t) {
    function fallback() { var ta = h('textarea'); ta.value = t; document.body.appendChild(ta); ta.select(); try { document.execCommand('copy'); toast('복사했습니다.'); } catch (e) { toast('복사하지 못했습니다. 직접 선택해 복사해 주세요.', true); } ta.remove(); }
    if (navigator.clipboard && window.isSecureContext) navigator.clipboard.writeText(t).then(function () { toast('복사했습니다.'); }, fallback); else fallback();
  }
  function pageHead(stage, title, lead, right) {
    return h('div', { class: 'page-head' }, h('div', { class: 'titles' }, h('div', { class: 'stage' }, stage), h('h1', null, title), lead ? h('p', { class: 'lead' }, lead) : null), right || null);
  }
  function tile(k, v, sub, cls) { return h('div', { class: 'tile' }, h('div', { class: 'k' }, k), h('div', { class: 'v' + (cls ? ' ' + cls : '') }, v), sub ? h('div', { class: 'note' }, sub) : null); }
  function table(head, rows, opts) {
    opts = opts || {};
    var t = h('table', { class: 'list' + (opts.cls ? ' ' + opts.cls : '') });
    var tr = h('tr');
    head.forEach(function (x, i) { tr.appendChild(h('th', { class: opts.num && opts.num[i] ? 'num' : null, scope: 'col' }, x)); });
    t.appendChild(h('thead', null, tr));
    var tb = h('tbody');
    rows.forEach(function (r) {
      var row = h('tr', { class: r.cls || null });
      (r.cells || r).forEach(function (x, i) { row.appendChild(h('td', { class: ((opts.num && opts.num[i] ? 'num' : '') + (r.tdCls && r.tdCls[i] ? ' ' + r.tdCls[i] : '')) || null }, x)); });
      tb.appendChild(row);
    });
    t.appendChild(tb);
    return h('div', { class: 'table-wrap' }, t);
  }
  function writeSheets(fileName, sheets) {
    var wb = XLSX.utils.book_new();
    sheets.forEach(function (s) {
      var ws = XLSX.utils.aoa_to_sheet(s.aoa);
      (s.formula || []).forEach(function (f) { ws[f.a] = f.v == null ? { t: 's', v: '', f: f.f } : { t: 'n', v: f.v, f: f.f }; });
      L.formatCells(ws); // 수량·금액 천 단위 쉼표 · 날짜 yyyy-mm-dd (10-28)
      if (s.cols) ws['!cols'] = s.cols;
      XLSX.utils.book_append_sheet(wb, ws, s.name);
    });
    XLSX.writeFile(wb, fileName);
  }
  function hasSession() { return !!(db.session && db.items.length); }
  function emptyNotice() {
    return h('section', { class: 'card empty' },
      h('p', null, '아직 이 기기에 실사 회차가 없습니다.'),
      h('p', null, '사무실에서는 ERP 재고 엑셀로 작업지시를 만들고, 현장(휴대폰)에서는 사무실에서 받은 기록서 엑셀을 넣으면 같은 작업지시가 이 기기에 생깁니다. 가상 샘플로 먼저 둘러봐도 됩니다.'),
      h('div', { class: 'btn-row', style: 'justify-content:center' },
        orderBookButton('btn btn-primary', '받은 기록서 엑셀 넣기(현장)'),
        h('a', { class: 'btn', href: '#/order' }, '① 작업지시 만들기(사무실)'),
        h('button', { type: 'button', class: 'btn', onclick: loadSample }, '가상 샘플로 시작')),
      teamSteps());
  }
  /* 현장: 사무실에서 받은 「재고실사 기록서 엑셀」(지시정보 시트가 있는 파일)만 받습니다 */
  function orderBookButton(cls, text) {
    var f = h('input', { type: 'file', accept: '.xlsx,.xls,.xlsm' });
    f.addEventListener('change', function () {
      var file = f.files[0]; f.value = ''; if (!file) return;
      readBook(file).then(function (sheets) {
        if (sheets.some(function (x) { return x.name === '지시정보'; })) loadOrderBook(sheets, file.name);
        else toast(file.name + ': 작업지시(기록서) 엑셀이 아닙니다 — 사무실에서 ① 작업지시 「기록서 엑셀」로 받은 파일을 넣어 주세요. ERP 원본 엑셀은 ① 작업지시에서 올립니다.', true);
      }, function (e) { toast(e.message, true); });
    });
    return h('label', { class: cls }, f, text);
  }
  /* 여러 사람이 함께 쓰는 순서 — 1단계는 서버(DB)가 없어 기기끼리 파일로 주고받습니다 (11-08 수강생 질문) */
  function teamSteps() {
    return h('details', { class: 'team-steps' }, h('summary', null, '여러 사람이 각자 휴대폰으로 입력하는 순서'),
      h('ol', { class: 'steps' },
        h('li', null, '사무실(1명): ① 작업지시에서 ERP 엑셀 → 배정표 → 「기록서 엑셀」을 받습니다.'),
        h('li', null, '그 파일을 실사자들에게 보냅니다(SharePoint · 메일 · 메신저 등).'),
        h('li', null, '실사자(N명): 휴대폰에서 이 주소를 열고 「받은 기록서 엑셀 넣기」 → 작성자 이름 → 「내 배정만」 보고 입력합니다. 같은 파일을 넣은 휴대폰은 모두 같은 실사 회차가 됩니다.'),
        h('li', null, '다 적으면 ② 현장 입력 맨 아래 「결과 파일 받기」(사진이 있으면 「사진도 넣기」)로 받아 사무실에 보냅니다.'),
        h('li', null, '사무실: ③ 취합에 결과 파일을 모두 넣으면 합쳐집니다. 여러 번 보내도 같은 기록은 한 번만 들어갑니다.')),
      h('p', { class: 'note' }, '기록은 각 기기의 브라우저에만 저장됩니다(서버로 보내지 않음). 그래서 주소만 열면 빈 화면이고, 기록서 엑셀을 넣어야 작업지시가 보입니다. 넣지 않고 바로 실시간으로 쌓이게 하는 것은 2단계(회사가 허용한 공유 DB 또는 SharePoint 목록)입니다.'));
  }
  function today() { return L.localIso().slice(0, 10); }

  /* ── 샘플 ───────────────────────────────── */
  function loadSample() {
    if ((db.records.length || db.items.length) && !db.sample && !confirm('지금 자료를 지우고 가상 샘플을 넣을까요? (먼저 JSON 백업을 받아 두면 되돌릴 수 있습니다)')) return;
    var b = SMP.build();
    db = L.emptyDb();
    db.session = b.session; db.items = b.items; db.assign = b.assign; db.records = b.records; db.files = b.files; db.sample = true;
    db.excludedItems = b.applied.excludedItems || [];
    save();
    toast('가상 샘플을 넣었습니다 — 240품목(창고 1곳 제외), 실사자 5명. 실사자E 가 SharePoint 파일에 적은 것은 ③ 취합에서 합쳐 보세요.');
    location.hash = '#/home'; render();
  }
  function mergeSampleE() {
    var b = SMP.build();
    var sheets = SMP.fillPaper(L.orderSheets(db.session, db.items, db.assign, { settings: db.settings }), b.heldOut, b.mainRows);
    importOrderBook(sheets, b.eFile);
  }

  /* ── 0. 진행 현황 ────────────────────────── */
  function pageHome() {
    var head = pageHead('진행 현황', '재고실사 Start-End', '작업 배정 → 현장 확인 → 차이 확인 → 사유 확인 → 취합 → 보고서를 한 흐름으로 잇습니다.');
    if (!hasSession()) return [head, flowCard(), emptyNotice()];
    var an = L.analyze(db), a = an.all, s = db.session;
    var asg = an.dims.assignee.slice().sort(function (x, y) { return x.key < y.key ? -1 : 1; });
    return [head,
      h('section', { class: 'card' }, h('h2', null, s.title),
        h('p', { class: 'note' }, '지시서 ID ' + s.id + ' · 기준 일자·시각(ERP 재고) ' + L.fmtTime(s.baseAt) + ' · 배정 단위 ' + (s.unitMode === 'zone' ? '창고·구역' : '창고') + (s.sourceFile ? ' · ' + s.sourceFile : '') + (s.rosterFile ? ' · 배정표 ' + s.rosterFile : '') + ' · 허용 기준: ' + L.tolText(db.settings)),
        an.excluded.length ? h('p', { class: 'note' }, '이번 실사 제외: ' + an.excluded.map(function (x) { return L.whLabel(x) + ' ' + N(x.items) + '품목' + (x.note ? '(' + x.note + ')' : ''); }).join(' · ')) : null),
      h('div', { class: 'tiles' },
        tile('진행률', pct(a.progress), N(a.counted) + ' / ' + N(a.items) + '품목'),
        tile('정확도(수량 일치)', pct(a.accuracy), '일치 ' + N(a.matched) + '품목'),
        tile('차이 품목', N(a.diffItems), an.S.tolMode === 'range' ? '허용 초과 ' + N(a.overItems) : '사유 미입력 ' + N(a.noReason), a.overItems ? 'bad' : ''),
        tile('순차이 금액', L.fmtSigned(a.netAmt, true), '절대 차이 ' + W(a.absAmt))),
      an.conflicts.length ? h('p', { class: 'alert warn' }, '두 사람이 따로 센 수량이 다른 품목 ' + an.conflicts.length + '개 — ', h('a', { href: '#/merge' }, '③ 취합'), '에서 채택할 기록을 골라 주세요.') : null,
      h('section', { class: 'card' }, h('h2', null, '실사자별 진행'),
        C.figure('실사자별 배정 진행률(%)', { type: 'hbar', unit: '%', labels: asg.filter(function (x) { return x.items; }).map(function (x) { return x.key; }), series: [{ name: '진행률', values: asg.filter(function (x) { return x.items; }).map(function (x) { return x.progress; }), color: 'var(--series-1)', colors: asg.filter(function (x) { return x.items; }).map(function (x) { return x.progress < 0.8 ? 'var(--series-2)' : 'var(--series-1)'; }) }] }),
        personTable(asg, false)),
      flowCard()];
  }
  /* 실사자별 표 — 배정 · 배정 진행률은 배정표 기준, 직접 센 품목 · 일치 · 차이는 기록의 작성자(By) 기준 (11-08) */
  function personTable(asg, detail) {
    var head = ['실사자', '배정 품목', '직접 센 품목', '배정 진행률'].concat(detail ? ['일치', '차이', '사유 미입력'] : ['차이', '충돌']).concat(['다른 사람이 셈 / 대신 셈']);
    return [table(head, asg.map(function (x) {
      var cross = (x.byOthers ? '배정 중 ' + N(x.byOthers) + '개를 다른 사람이' : '') + (x.byOthers && x.forOthers ? ' · ' : '') + (x.forOthers ? '남의 배정 ' + N(x.forOthers) + '개를 대신' : '');
      return [x.key, x.items ? N(x.items) : '-', N(x.counted), x.items ? pct(x.progress) : '-'].concat(detail ? [N(x.matched), N(x.diffItems), x.noReason ? N(x.noReason) : ''] : [N(x.diffItems), x.conflicts ? N(x.conflicts) : '']).concat([cross]);
    }), { num: detail ? [0, 1, 1, 1, 1, 1, 1, 0] : [0, 1, 1, 1, 1, 1, 0] }),
      h('p', { class: 'note' }, '「직접 센 품목 · 일치 · 차이」는 실제로 센 사람(기록의 작성자 By) 기준이고, 「배정 품목 · 배정 진행률」은 배정표 기준입니다(배정된 품목은 누가 세었든 확인되면 진행으로 셈). 그래서 남의 배정을 대신 센 사람도 직접 센 품목에 나옵니다.')];
  }
  function flowCard() {
    var steps = [
      ['① 작업지시', '#/order', 'ERP 재고 엑셀 올리기 → 기준 일자·시각 → 창고 배정표(창고 · 담당자 · 실사자)로 배정 → 재고실사 기록서 엑셀(SharePoint 취합 파일)·인쇄'],
      ['② 현장 입력', '#/field', '휴대폰으로 자재를 찾아 실물수량·사유·특이사항·사진(필요할 때). 확인 시각·작성자는 자동'],
      ['③ 취합', '#/merge', '여러 사람이 채운 SharePoint 취합 엑셀(또는 휴대폰 결과 파일)을 합치기 — 같은 기록은 한 번만, 두 사람 수량이 다르면 충돌로 표시'],
      ['④ 차이대사', '#/variance', '시스템 수량 vs 실사 수량, 수량·금액 차이, 차이율, 허용 기준 판정'],
      ['⑤ 분석(BI)', '#/bi', '창고별·창고위치(사내/사외)별·실사자별·사유별(자재유형·자재타입 열이 있으면 그것도)'],
      ['⑥ 보고서', '#/report', '버전·기준 시각·원천 기록 연결이 들어간 결과서(엑셀)와 인쇄용 보고서, AI 요약 프롬프트']
    ];
    return h('section', { class: 'card' }, h('h2', null, '흐름'),
      h('ol', { class: 'flow' }, steps.map(function (s) { return h('li', null, h('a', { href: s[1] }, s[0]), h('span', { class: 'note' }, ' — ' + s[2])); })),
      h('p', { class: 'note' }, '1단계는 기기마다 따로 저장합니다(서버 없음). 지금 쓰는 방식(SharePoint 에 취합 엑셀 하나를 올려 각자 채움)을 그대로 살려, 이 도구가 낸 기록서 엑셀을 SharePoint 에 올리고 다 채운 파일을 ③ 취합에 넣습니다. 여러 휴대폰이 바로 한곳에 쌓이는 실시간 공유는 2단계(SharePoint 목록 + Power Automate, 또는 회사가 허용한 공유 DB)입니다.'));
  }

  /* ── ① 작업지시 ──────────────────────────── */
  function readBook(file) {
    return new Promise(function (resolve, reject) {
      var fr = new FileReader();
      fr.onload = function () {
        try {
          // 날짜 칸은 일련번호로 받아 직접 바꿉니다(cellDates 를 켜면 한국 시간에서 52초 어긋나 자정이 전날 23:59 가 됨)
          var wb = XLSX.read(new Uint8Array(fr.result), { type: 'array', cellDates: false });
          resolve(wb.SheetNames.map(function (n) { return { name: n, aoa: XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null }) }; }));
        } catch (e) { reject(new Error(file.name + ': 엑셀로 읽지 못했습니다(' + e.message + ')')); }
      };
      fr.onerror = function () { reject(new Error(file.name + ': 파일을 읽지 못했습니다.')); };
      fr.readAsArrayBuffer(file);
    });
  }
  function handleOrderFile(file) {
    if (!file) return;
    readBook(file).then(function (sheets) {
      if (sheets.some(function (s) { return s.name === '지시정보'; })) { loadOrderBook(sheets, file.name); return; }
      var pick = L.pickSheet(sheets) || sheets[0];
      // 같은 파일에 창고 배정표 시트가 있으면(수강생 첨부처럼) 작업지시를 만들 때 함께 적용
      var rs = sheets.filter(function (s) { return s !== pick && !L.parseRoster(s.aoa).error; })[0];
      ui.pending = { fileName: file.name, sheets: sheets, sheet: pick ? pick.name : '', overrides: {}, roster: rs ? { sheet: rs.name, ro: L.parseRoster(rs.aoa) } : null };
      reparse(); render();
    }, function (e) { toast(e.message, true); });
  }
  function reparse() {
    var p = ui.pending; if (!p) return;
    var sh = p.sheets.filter(function (s) { return s.name === p.sheet; })[0];
    p.aoa = sh ? sh.aoa : [];
    p.result = L.parseErp(p.aoa, { overrides: p.overrides, fileName: p.fileName });
  }
  /* 휴대폰: 사무실에서 받은 작업지시서 엑셀을 넣으면 같은 실사 회차가 이 기기에 생깁니다 */
  function loadOrderBook(sheets, fileName) {
    var w = L.parseOrderWorkbook(sheets, { fileName: fileName });
    if (w.error) { toast(w.error, true); return; }
    if (db.session && db.session.id === w.session.id) { importOrderBook(sheets, fileName); return; }
    if (db.records.length && !db.sample && !confirm('이 기기에 다른 실사 회차(' + db.session.title + ')의 기록 ' + db.records.length + '건이 있습니다. 결과 파일을 내보냈는지 확인하셨나요? 새 작업지시서로 바꿉니다.')) return;
    db = L.emptyDb();
    db.session = w.session; db.items = w.items; db.assign = w.assign; db.records = w.records; db.sample = false;
    if (w.reasons.length && w.reasons.join('/') !== L.REASONS.join('/')) db.settings.reasons = w.reasons; // 사무실 설정의 사유 목록을 따름
    db.files = w.records.length ? [{ name: fileName, records: w.records.length, dup: 0, at: L.localIso() }] : [];
    save();
    toast('작업지시서를 넣었습니다 — ' + N(w.items.length) + '품목. ② 현장 입력에서 담당을 골라 시작해 주세요.');
    location.hash = '#/field'; render();
  }
  function pageOrder() {
    var head = pageHead('데이터 01 · 작업 배정', '작업지시', 'ERP 에서 내려받은 재고 엑셀로 시작합니다. 기준 일자·시각을 박고, 창고 배정표(창고 목록 · 담당자 · 실사자)로 배정해 재고실사 기록서(작업지시서 · SharePoint 취합 파일)를 냅니다.');
    var file = h('input', { type: 'file', accept: '.xlsx,.xls,.xlsm,.csv' });
    file.addEventListener('change', function () { handleOrderFile(file.files[0]); file.value = ''; });
    var up = h('section', { class: 'card' }, h('h2', null, '1. 엑셀 올리기'),
      h('div', { class: 'drop' }, h('label', { class: 'btn btn-primary' }, file, 'ERP 재고 엑셀 고르기'),
        h('span', { class: 'note' }, '사무실: ERP 재고 내려받기 파일 · 휴대폰: 받은 「작업지시서」 엑셀(그대로 넣으면 같은 실사 회차가 이 기기에 생김)')),
      h('div', { class: 'btn-row' }, h('button', { type: 'button', class: 'btn', onclick: loadSample }, '가상 샘플로 시작'),
        h('span', { class: 'note' }, '샘플 파일: samples/ERP_재고현황_다운로드_샘플_2026-09.xlsx')),
      h('p', { class: 'note' }, '파일은 서버로 보내지 않고 이 브라우저 안에서만 읽습니다. ERP 견본의 열(자재코드 · 자재내역 · 저장위치 · 저장위치명 · 창고위치 · 단가 · 금액 · Location · 수량)을 그대로 읽고, 열이 더 붙어 있으면 기록서에 그대로 싣습니다. 못 찾은 열은 아래에서 골라 연결합니다. 같은 파일에 창고 배정표 시트가 있으면 함께 적용합니다.'));
    var out = [head, up];
    if (ui.pending) out.push(pendingCard());
    if (hasSession()) out.push(assignCard(), exportCard());
    else if (!ui.pending) out.push(emptyNotice());
    return out;
  }
  function pendingCard() {
    var p = ui.pending, r = p.result, det = r.det;
    var sheetSel = h('select', { 'aria-label': '시트' }, p.sheets.map(function (s) { var o = h('option', { value: s.name }, s.name + ' (' + s.aoa.length + '행)'); if (s.name === p.sheet) o.selected = true; return o; }));
    sheetSel.addEventListener('change', function () { p.sheet = sheetSel.value; p.overrides = {}; reparse(); render(); });
    var headers = det.headers || [];
    var maps = L.ERP_FIELDS.map(function (f) {
      var sel = h('select', { 'aria-label': f.label + ' 열' });
      sel.appendChild(h('option', { value: '-1' }, '(없음)'));
      headers.forEach(function (t, i) { var o = h('option', { value: String(i) }, L.colName(i) + ' · ' + (t || '(빈 머리)')); if (det.map[f.id] === i) o.selected = true; sel.appendChild(o); });
      sel.addEventListener('change', function () { p.overrides[f.id] = Number(sel.value); reparse(); render(); });
      var col = det.map[f.id];
      var ex = col == null ? '' : (p.aoa.slice(det.headerRow + 1, det.headerRow + 4).map(function (row) { return row ? row[col] : ''; }).filter(function (v) { return v != null && v !== ''; }).join(' · '));
      var need = (f.id === 'no' && det.missing.indexOf('no') >= 0) || det.missing.indexOf(f.id) >= 0;
      return h('div', { class: 'map-row' + (need ? ' missing' : '') }, h('span', { class: 'col' }, f.label, need ? h('span', { class: 'req' }, ' 꼭 필요') : null), sel, h('span', { class: 'ex' }, ex ? '예: ' + ex : ''));
    });
    var title = h('input', { value: db.session && db.session.sourceFile === p.fileName ? db.session.title : '', placeholder: '예: 2026년 10월 정기 재고실사' });
    var base = h('input', { type: 'datetime-local', step: 60 });
    var unitSel = h('select', null, h('option', { value: 'wh' }, '창고별'), h('option', { value: 'zone' }, '창고·구역별(저장위치 앞자리)'));
    var baseErr = ui.baseErr;
    function create() {
      if (det.missing.length) { toast('꼭 필요한 열을 먼저 연결해 주세요.', true); return; }
      if (!L.validBaseAt(base.value)) { ui.baseErr = 'ERP 에서 재고를 내려받은 일자·시각을 넣어 주세요. 이 시각이 작업지시서·보고서에 함께 찍힙니다.'; render(); return; }
      if (db.records.length && !db.sample && !confirm('지금 실사 회차의 현장 기록 ' + db.records.length + '건이 지워집니다. 새 작업지시를 만들까요?')) return;
      var s;
      try { s = L.makeSession({ title: title.value, baseAt: base.value, sourceFile: p.fileName, unitMode: unitSel.value }); } catch (e) { toast(e.message, true); return; }
      var keepNames = db.sample ? [] : L.assignedNames(db.assign), keepSettings = db.settings;
      s.cols = r.cols.map(function (x) { return { h: x.h, f: x.f }; });
      db = L.emptyDb(); db.session = s; db.items = r.items; db.settings = keepSettings;
      if (p.roster) applyRosterToDb(p.roster.ro, p.fileName + ' · ' + p.roster.sheet);
      else if (keepNames.length) db.assign = L.autoAssign(L.unitsOf(db.items, s.unitMode), keepNames);
      ui.pending = null; ui.baseErr = '';
      if (location.hash !== '#/order') history.replaceState(null, '', '#/order'); // 처음 연 주소(해시 없음)에서도 배정 화면에 머묾
      save(); toast('작업지시를 만들었습니다 — ' + N(db.items.length) + '품목' + (p.roster ? '(배정표 적용)' : '. 아래에서 창고 배정표를 올리거나 실사자를 배정해 주세요.')); render();
    }
    return h('section', { class: 'card' }, h('h2', null, '2. 열 확인 · 기준 일자·시각 — ' + p.fileName),
      p.sheets.length > 1 ? fld('시트', sheetSel) : null,
      h('div', { class: 'map-grid' }, maps),
      det.warnings.concat(r.warnings).map(function (w) { return h('p', { class: 'alert info' }, w); }),
      det.info ? h('p', { class: 'note' }, det.info) : null,
      p.roster ? h('p', { class: 'alert info' }, '같은 파일의 「' + p.roster.sheet + '」 시트를 창고 배정표로 읽었습니다 — 창고 ' + p.roster.ro.rows.length + '곳' + (p.roster.ro.rows.some(function (x) { return x.excluded; }) ? '(실사인원 0 인 ' + p.roster.ro.rows.filter(function (x) { return x.excluded; }).length + '곳은 이번 실사에서 뺌)' : '') + (p.roster.ro.date ? ' · 배정표 날짜 ' + p.roster.ro.date : '') + '. 작업지시를 만들 때 함께 적용합니다.') : null,
      det.missing.length ? null : h('p', { class: 'note' }, '읽은 품목 ' + N(r.items.length) + '개 · 합계 줄 ' + r.summary + '개 제외' + (r.skipped.length ? ' · 건너뛴 줄 ' + r.skipped.length + '개' : '') + ' · 장부금액 ' + W(r.items.reduce(function (s, it) { return s + it.amt; }, 0))),
      r.skipped.length ? h('details', { class: 'warn-list' }, h('summary', null, '건너뛴 줄 ' + r.skipped.length + '개'), h('ul', null, r.skipped.slice(0, 100).map(function (x) { return h('li', null, x.row + '행: ' + x.why); }))) : null,
      h('div', { class: 'form-grid', style: 'margin-top:12px' },
        fld('실사명', title, '비우면 「YYYY-MM 정기 재고실사」'),
        fld('기준 일자·시각 (필수)', base, 'ERP 에서 재고를 내려받은 시각 — 인벤토리에서 빠진 항목' + (p.roster && p.roster.ro.date ? ' · 배정표 날짜 ' + p.roster.ro.date : ''), baseErr),
        fld('배정 단위', unitSel, '창고가 크면 구역별로 나눠 여러 사람에게')),
      h('div', { class: 'btn-row', style: 'margin-top:12px' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: create, disabled: !!det.missing.length }, '작업지시 만들기'),
        h('button', { type: 'button', class: 'btn', onclick: function () { ui.pending = null; render(); } }, '취소')));
  }
  /* 창고 배정표(창고 목록 · 담당자 · 실사자) 적용 — 실사인원 0 창고는 빼고, 실사자가 여럿이면 보관 자리 순서로 나눔 */
  function applyRosterToDb(ro, fileName) {
    var s = db.session;
    var all = db.items.concat(db.excludedItems || []);
    var ap = L.applyRoster(all, ro);
    db.items = ap.items; db.excludedItems = ap.excludedItems; db.assign = ap.assign;
    s.unitMode = 'wh'; s.managers = ap.managers; s.excluded = ap.excluded; s.rosterFile = fileName || ''; s.orderVer = (s.orderVer || 1) + 1;
    ui.rosterMsg = ro.warnings.concat(
      ap.unrostered.length ? ['배정표에 없는 창고(배정 안 됨): ' + ap.unrostered.map(function (x) { return x.wh + ' ' + N(x.items) + '품목'; }).join(', ')] : [],
      ap.missing.length ? ['배정표에 있지만 ERP 파일에 품목이 없는 창고: ' + ap.missing.join(', ')] : []);
    return ap;
  }
  function handleRosterFile(file) {
    if (!file) return;
    readBook(file).then(function (sheets) {
      var hit = null;
      sheets.some(function (sh) { var ro = L.parseRoster(sh.aoa); if (!ro.error) { hit = { sheet: sh.name, ro: ro }; return true; } return false; });
      if (!hit) { toast(file.name + ': 창고 배정표(저장위치 · 실사인원 · 담당자 · 실사자 …)를 찾지 못했습니다.', true); return; }
      if (db.records.length && !db.sample && !confirm('배정을 배정표대로 다시 나눕니다. 이미 받은 현장 기록은 그대로 둡니다. 계속할까요?')) return;
      var ap = applyRosterToDb(hit.ro, file.name);
      save(); render();
      toast('배정표를 적용했습니다 — 실사자 ' + L.assignedNames(db.assign).length + '명' + (ap.excluded.length ? ', 제외 창고 ' + ap.excluded.length + '곳' : ''));
    }, function (e) { toast(e.message, true); });
  }
  function assignCard() {
    var s = db.session, units = L.unitsOf(db.items, s.unitMode);
    var names = L.assignedNames(db.assign);
    var dl = h('datalist', { id: 'nameList' }, names.map(function (n) { return h('option', { value: n }); }));
    var auto = h('input', { value: names.join(', '), placeholder: '예: 실사자A, 실사자B, 실사자C' });
    var unitSel = h('select', null, h('option', { value: 'wh' }, '창고별(배정표)'), h('option', { value: 'zone' }, '창고·구역별'));
    unitSel.value = s.unitMode;
    unitSel.addEventListener('change', function () {
      if (db.records.length && !confirm('배정 단위를 바꾸면 실사자를 다시 나눕니다. 계속할까요?')) { unitSel.value = s.unitMode; return; }
      s.unitMode = unitSel.value; s.orderVer = (s.orderVer || 1) + 1;
      db.assign = L.autoAssign(L.unitsOf(db.items, s.unitMode), names); save(); render();
    });
    var rf = h('input', { type: 'file', accept: '.xlsx,.xls,.xlsm,.csv' });
    rf.addEventListener('change', function () { handleRosterFile(rf.files[0]); rf.value = ''; });
    var rows = units.map(function (u) {
      var inp = h('input', { list: 'nameList', value: db.assign[u.key] || '', placeholder: '실사자', 'aria-label': u.key + ' 실사자' });
      inp.addEventListener('change', function () {
        var v = inp.value.trim();
        if (v) db.assign[u.key] = v; else delete db.assign[u.key];
        s.orderVer = (s.orderVer || 1) + 1; save(); render();
      });
      var label = L.whLabel(u) + (u.part ? ' #' + u.part : '') + (u.zone ? ' / ' + u.zone : '');
      return { cls: db.assign[u.key] ? null : 'differs', cells: [label, u.site || '', (s.managers || {})[u.wh] || '', N(u.items), W(u.sysAmt), inp] };
    });
    var per = {};
    db.items.forEach(function (it) { var a = L.assigneeOf(it, s, db.assign) || '(미배정)'; per[a] = (per[a] || 0) + 1; });
    return h('section', { class: 'card' }, h('h2', null, '3. 배정 — 창고 배정표 · 실사자'),
      h('div', { class: 'drop' }, h('label', { class: 'btn btn-primary' }, rf, '창고 배정표 엑셀 올리기'),
        h('span', { class: 'note' }, '머리: 저장위치 · 저장위치명 · 창고위치 · 실사인원 · 담당팀 · 담당자 · 실사자 · 비고 — 실사인원 0 은 이번 실사에서 빼고, 실사자가 「갑, 을」이면 그 창고를 보관 자리 순서로 사람 수만큼 나눕니다.')),
      s.rosterFile ? h('p', { class: 'note' }, '적용한 배정표: ' + s.rosterFile) : null,
      (ui.rosterMsg || []).map(function (w) { return h('p', { class: 'alert info' }, w); }),
      (s.excluded || []).length ? h('p', { class: 'alert warn' }, '이번 실사 제외: ' + s.excluded.map(function (x) { return L.whLabel(x) + ' ' + N(x.items) + '품목 ' + W(x.sysAmt) + (x.note ? ' — ' + x.note : ''); }).join(' · ')) : null,
      h('details', { class: 'warn-list' }, h('summary', null, '배정표 없이 나누기(이름 목록 · 구역별)'),
        h('div', { class: 'form-grid' },
          fld('배정 단위', unitSel),
          h('div', { class: 'span-2' }, fld('실사자 이름(쉼표로) → 자동 배정', auto, '품목 수가 고르게 되도록 나눕니다. 창고(구역)는 쪼개지 않습니다.'))),
        h('div', { class: 'btn-row', style: 'margin:8px 0 14px' }, h('button', { type: 'button', class: 'btn', onclick: function () {
          var list = auto.value.split(/[,，\n]/).map(function (x) { return x.trim(); }).filter(Boolean);
          if (!list.length) { toast('실사자 이름을 넣어 주세요.', true); return; }
          db.items.forEach(function (it) { delete it.part; });
          db.assign = L.autoAssign(L.unitsOf(db.items, s.unitMode), list); s.orderVer = (s.orderVer || 1) + 1; save(); render(); toast(list.length + '명에게 나눴습니다.');
        } }, '자동 배정'))),
      dl,
      table([s.unitMode === 'zone' ? '창고 / 구역' : '창고', '창고위치', '담당자', '품목', '장부금액', '실사자'], rows, { num: [0, 0, 0, 1, 1, 0], cls: 'items' }),
      h('p', { class: 'note' }, '실사자별: ' + Object.keys(per).sort().map(function (k) { return k + ' ' + per[k] + '품목'; }).join(' · ') + ' · 지시서 버전 v' + (s.orderVer || 1) + '(배정을 바꿀 때마다 올라감)'));
  }
  function bookCols(sheets) {
    sheets.forEach(function (x, i) {
      x.cols = i === 0 ? [{ wch: 20 }, { wch: 70 }] : x.aoa[0].map(function (hd) { return { wch: /자재내역|자재명|특이사항/.test(hd) ? 22 : /재고번호/.test(hd) ? 18 : /시각/.test(hd) ? 16 : Math.max(8, String(hd).length * 2 + 2) }; });
    });
    return sheets;
  }
  function exportCard() {
    var s = db.session;
    var unassigned = db.items.filter(function (it) { return !L.assigneeOf(it, s, db.assign); }).length;
    function xlsx() {
      var sheets = bookCols(L.orderSheets(s, db.items, db.assign, { records: db.records, picks: db.picks, settings: db.settings }));
      writeSheets('재고실사기록서_' + s.baseAt.slice(0, 10) + '_v' + (s.orderVer || 1) + '.xlsx', sheets);
    }
    return h('section', { class: 'card' }, h('h2', null, '4. 재고실사 기록서 내보내기 (SharePoint 취합 파일)'),
      unassigned ? h('p', { class: 'alert warn' }, '실사자가 없는 품목 ' + unassigned + '개가 있습니다.') : null,
      h('p', null, '엑셀 하나에 「지시정보」(지시서 ID · 버전 · 기준 일자·시각 · 사유 구분) · 「' + L.MAIN_SHEET + '」(전체) · 실사자별 시트가 들어갑니다. 열은 ERP 원래 열 뒤에 담당자 · 실사자 · 실물수량 · 차이수량 · 사유 · 특이사항 · 확인자 · 확인시각입니다. 차이수량은 실물수량을 적으면 저절로 계산됩니다(실물수량 − 수량, 모자라면 −).'),
      h('ol', { class: 'steps' },
        h('li', null, '이 파일을 SharePoint 에 올립니다(지금처럼 취합 파일 1개).'),
        h('li', null, '각자 자기 이름 시트에(또는 전체 시트의 자기 줄에) 실물수량을 적고, 차이가 있으면 사유를 적습니다. 휴대폰으로 이 도구에서 적은 사람은 「결과 엑셀 받기」 파일을 보내도 됩니다.'),
        h('li', null, '다 채운 SharePoint 파일을 내려받아 ③ 취합에 그대로 넣습니다. 여러 번 넣어도 같은 기록은 한 번만 들어갑니다.')),
      h('p', { class: 'alert info' }, '휴대폰으로 입력할 사람에게도 이 파일을 보내 주세요. 휴대폰에서 이 주소를 열고 「받은 기록서 엑셀 넣기」를 누르면 같은 작업지시가 생기고, 다 적은 뒤 「결과 파일」을 보내면 ③ 취합에서 합쳐집니다(주소만 열면 빈 화면입니다 — 기록이 기기마다 따로 저장되기 때문).'),
      teamSteps(),
      h('div', { class: 'btn-row' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: xlsx }, '기록서 엑셀 (SharePoint 취합용)'),
        h('button', { type: 'button', class: 'btn', onclick: printOrder }, '인쇄용 기록서')),
      h('p', { class: 'note' }, '현장 기록이 이미 있으면 채워서 냅니다(지금까지 합친 취합본). 인쇄본은 실사자마다 새 쪽에서 시작하고, 위에 기준 일자·시각 · 지시서 ID · 버전과 담당자 · 실사자 서명 칸이 있습니다.'));
  }
  function printOrder() {
    var s = db.session, area = document.getElementById('printArea');
    area.innerHTML = '';
    var items = L.sortForWalk(db.items);
    var names = L.assignedNames(db.assign);
    var reasons = L.reasonList(db.settings);
    if (items.some(function (it) { return !L.assigneeOf(it, s, db.assign); })) names.push('');
    names.forEach(function (n) {
      var mine = items.filter(function (it) { return (L.assigneeOf(it, s, db.assign) || '') === n; });
      var mgr = unique(mine.map(function (it) { return L.managerOf(it, s); })).join(', ');
      area.appendChild(h('section', { class: 'print-sheet' },
        h('h1', null, '재고실사 기록서 — ' + (n || '(미배정)')),
        h('p', null, s.title + ' · 지시서 ID ' + s.id + ' · v' + (s.orderVer || 1) + ' · 기준 일자·시각 ' + L.fmtTime(s.baseAt) + ' · ' + N(mine.length) + '품목'),
        h('p', { class: 'sign' }, '담당자 ' + (mgr || '________') + ' (인)      실사자 ' + (n || '________') + ' (인)'),
        h('table', { class: 'list' }, h('thead', null, h('tr', null, ['자재코드', '자재내역', '저장위치', 'Location', '수량', '실물수량', '차이수량', '사유', '특이사항'].map(function (x) { return h('th', null, x); }))),
          h('tbody', null, mine.map(function (it) { return h('tr', null, h('td', null, it.code || it.no), h('td', null, it.name), h('td', null, L.whLabel(it)), h('td', null, it.loc), h('td', { class: 'num' }, N(it.sys)), h('td', null, ''), h('td', null, ''), h('td', null, ''), h('td', null, '')); }))),
        h('p', null, '사유: ' + reasons.join(' / ') + ' · 차이수량 = 실물수량 − 수량(모자라면 −)'),
        h('p', { class: 'sign' }, '확인 시각 ____ 년 __ 월 __ 일 __ 시 __ 분')));
    });
    document.body.classList.add('printing-order');
    setTimeout(function () { window.print(); document.body.classList.remove('printing-order'); }, 50);
  }

  /* ── ② 현장 입력 (휴대폰) ───────────────── */
  function shrinkImage(file) {
    return new Promise(function (resolve, reject) {
      var fr = new FileReader();
      fr.onload = function () {
        var img = new Image();
        img.onload = function () {
          var max = 800, w = img.width, hh = img.height, k = Math.min(1, max / Math.max(w, hh));
          var cv = document.createElement('canvas'); cv.width = Math.round(w * k); cv.height = Math.round(hh * k);
          cv.getContext('2d').drawImage(img, 0, 0, cv.width, cv.height);
          resolve(cv.toDataURL('image/jpeg', 0.6));
        };
        img.onerror = function () { reject(new Error('사진을 읽지 못했습니다.')); };
        img.src = fr.result;
      };
      fr.onerror = function () { reject(new Error('사진을 읽지 못했습니다.')); };
      fr.readAsDataURL(file);
    });
  }
  function pageField() {
    var head = pageHead('데이터 02 · 04 · 현장 확인', '현장 입력', '휴대폰으로 자재를 찾아 실물수량을 적습니다. 차이가 있으면 사유를 꼭 고르고(사진은 필요할 때만), 확인 시각·작성자는 자동으로 남습니다.');
    if (!hasSession()) return [head, emptyNotice()];
    var s = db.session;
    var names = L.assignedNames(db.assign);
    var meIn = h('input', { list: 'meList', value: me.name || '', placeholder: '작성자 이름', autocomplete: 'off' });
    meIn.addEventListener('change', function () { me.name = meIn.value.trim(); if (!ui.fieldScope && names.indexOf(me.name) >= 0) ui.fieldScope = 'me'; S.saveMe(me); render(); });
    var scope = ui.fieldScope || (names.indexOf(me.name) >= 0 ? 'me' : 'all');
    var scopeSel = h('select', { 'aria-label': '보는 범위' },
      h('option', { value: 'me' }, '내 배정만' + (me.name ? ' (' + me.name + ')' : '')),
      h('option', { value: 'all' }, '전체'),
      whOptions().map(function (w) { return h('option', { value: 'wh:' + w[0] }, w[1]); }));
    scopeSel.value = scope;
    scopeSel.addEventListener('change', function () { ui.fieldScope = scopeSel.value; render(); });
    var res = L.resolveCounts(db.records, db.picks);
    var base = db.items.filter(function (it) {
      if (scope === 'me') return L.assigneeOf(it, s, db.assign) === me.name;
      if (scope.indexOf('wh:') === 0) return it.wh === scope.slice(3);
      return true;
    });
    var doneN = base.filter(function (it) { return res[it.no]; }).length;
    function byFilter(it) {
      var r = res[it.no];
      if (ui.fieldFilter === 'todo') return !r;
      if (ui.fieldFilter === 'done') return !!r;
      if (ui.fieldFilter === 'diff') return r && r.current.qty !== it.sys;
      return true;
    }
    function currentList() {
      // 10-28 수강생 요청: 찾는 말(자재코드·내역 포함)과 보기(확인함 등)를 함께 겁니다 — 「확인함 중에서 ○○」
      var list = L.sortForWalk(base).filter(byFilter);
      return ui.q ? L.searchItems(list, ui.q) : list;
    }
    var q = h('input', { type: 'search', value: ui.q, placeholder: '자재코드·자재명·Location', inputmode: 'search', autocomplete: 'off', 'aria-label': '찾기' });
    q.addEventListener('input', function () { ui.q = q.value; renderList(); });
    var filters = h('div', { class: 'seg', role: 'radiogroup', 'aria-label': '보기' }, [['todo', '미확인'], ['done', '확인함'], ['diff', '차이'], ['all', '전체']].map(function (f) {
      var r = h('input', { type: 'radio', name: 'ff', value: f[0] }); if (ui.fieldFilter === f[0]) r.checked = true;
      r.addEventListener('change', function () { ui.fieldFilter = f[0]; ui.open = ''; render(); });
      return h('label', { class: 'pill' }, r, h('span', null, f[1]));
    }));
    var listBox = h('div', { class: 'item-list' });
    function renderList() {
      listBox.innerHTML = '';
      var items = currentList();
      var other = ui.q && !items.length && ui.fieldFilter !== 'all' ? L.searchItems(L.sortForWalk(base), ui.q).length : 0;
      if (!items.length) { listBox.appendChild(h('p', { class: 'note' }, other ? '이 보기에는 찾는 품목이 없습니다. 「전체」 보기에서는 ' + N(other) + '개가 맞습니다.' : ui.q ? '찾는 품목이 없습니다. 목록에 없는 자재면 아래 「목록에 없는 자재 적기」를 써 주세요.' : (ui.fieldFilter === 'todo' ? '미확인 품목이 없습니다. 아래에서 결과 파일을 내보내 주세요.' : '해당 품목이 없습니다.'))); return; }
      items.slice(0, 80).forEach(function (it) { listBox.appendChild(itemCard(it, res[it.no], items)); });
      if (items.length > 80) listBox.appendChild(h('p', { class: 'note' }, N(items.length) + '개 중 처음 80개만 보입니다. 재고번호·위치로 찾아 주세요.'));
    }
    renderList();
    var mine = db.records.filter(function (r) { return r.dev === me.dev; });
    return [head,
      h('section', { class: 'card field-top' },
        h('datalist', { id: 'meList' }, names.map(function (n) { return h('option', { value: n }); })),
        h('div', { class: 'form-grid two' }, fld('작성자', meIn, me.name ? null : '처음 한 번만 — 기록마다 자동으로 붙습니다'), fld('보는 범위', scopeSel)),
        h('div', { class: 'progress-line' }, h('div', { class: 'bar' }, h('i', { style: 'width:' + (base.length ? Math.round(doneN / base.length * 100) : 0) + '%' })),
          h('span', { class: 'note' }, '확인 ' + N(doneN) + ' / ' + N(base.length) + ' · 기준 ' + L.fmtTime(s.baseAt))),
        q, filters),
      listBox,
      extraCard(),
      h('section', { class: 'card' }, h('h2', null, '결과 파일 내보내기 (사무실로 보내기)'),
        h('p', null, '이 기기에서 입력한 기록 ' + mine.length + '건' + (mine.filter(function (r) { return r.photo; }).length ? ' (사진 ' + mine.filter(function (r) { return r.photo; }).length + '장)' : '') + '을 파일 하나로 받아 보내 주세요. 「결과 엑셀」은 SharePoint 취합 파일과 같은 기록서 양식(내 줄만)이라 그대로 올리거나 줄을 옮겨 붙여도 됩니다. 사진까지 보내려면 「결과 파일(JSON)」을 씁니다. 여러 번 보내도 같은 기록은 한 번만 들어갑니다.'),
        h('label', { class: 'opt' }, (function () { var c = h('input', { type: 'checkbox' }); c.checked = ui.withPhoto; c.addEventListener('change', function () { ui.withPhoto = c.checked; }); return c; })(),
          h('span', null, h('span', { class: 't' }, '사진도 넣기'), h('span', { class: 's' }, '파일이 커집니다(사진 1장에 약 50~100KB).'))),
        h('div', { class: 'btn-row', style: 'margin-top:8px' },
          h('button', { type: 'button', class: 'btn btn-primary', disabled: !mine.length, onclick: function () {
            var f = L.resultFile(s, mine, { by: me.name, dev: me.dev, withPhoto: ui.withPhoto });
            download('실사결과_' + (me.name || me.dev) + '_' + L.localIso().slice(0, 16).replace(/[-:T]/g, '') + '.json', new Blob([JSON.stringify(f)], { type: 'application/json' }));
          } }, '결과 파일(JSON) 받기'),
          h('button', { type: 'button', class: 'btn', disabled: !mine.length, onclick: function () {
            var sheets = bookCols(L.orderSheets(s, db.items, db.assign, { records: mine, only: me.name, settings: db.settings }));
            writeSheets('재고실사기록서_' + (me.name || me.dev) + '_' + L.localIso().slice(0, 16).replace(/[-:T]/g, '') + '.xlsx', sheets);
          } }, '결과 엑셀 받기(기록서 양식)'))),
      recentCard(mine)];
  }
  function unique(a) { var o = []; a.forEach(function (x) { if (x && o.indexOf(x) < 0) o.push(x); }); return o.sort(); }
  /* [창고 코드, 「코드 이름」] — 화면에는 이름을 붙여 보여 줌 */
  function whOptions() { var seen = {}, o = []; L.sortForWalk(db.items).forEach(function (it) { if (!seen[it.wh]) { seen[it.wh] = 1; o.push([it.wh, L.whLabel(it)]); } }); return o; }
  function itemCard(it, rs, list) {
    var cur = rs ? rs.current : null;
    var open = ui.open === it.no;
    var stateTxt = !cur ? '미확인' : cur.qty === it.sys ? '일치' : '차이 ' + L.fmtSigned(L.round(cur.qty - it.sys, 6));
    var headBtn = h('button', { type: 'button', class: 'item-head', 'aria-expanded': open ? 'true' : 'false', onclick: function () { ui.open = open ? '' : it.no; ui.photo = null; ui.err = null; ui.draft = null; render(); if (!open) setTimeout(function () { var x = document.getElementById('qty-' + it.no); if (x) { x.focus(); x.scrollIntoView({ block: 'center' }); } }, 30); } },
      h('span', { class: 'ih-main' }, h('b', null, it.code || it.no), ' ', h('span', null, it.name)),
      h('span', { class: 'ih-sub' }, L.whLabel(it) + ' · ' + it.loc + ' · 수량 ' + N(it.sys) + (it.unit ? ' ' + it.unit : '')),
      h('span', { class: 'state ' + (!cur ? 'todo' : cur.qty === it.sys ? 'ok' : 'diff') + (rs && rs.conflict ? ' conflict' : '') }, rs && rs.conflict ? '충돌' : stateTxt));
    var box = h('div', { class: 'item' + (open ? ' open' : '') }, headBtn);
    if (open) box.appendChild(countForm(it, cur, list));
    return box;
  }
  function countForm(it, cur, list) {
    var E = ui.err || {};
    var dr = ui.draft && ui.draft.no === it.no ? ui.draft : null; // 점검에 걸려 다시 그릴 때 적던 값을 살림
    var qty = h('input', { id: 'qty-' + it.no, inputmode: 'decimal', value: dr ? dr.qty : cur ? String(cur.qty) : '', placeholder: '실사 수량', autocomplete: 'off', class: 'qty' });
    var diffOut = h('div', { class: 'diff-out', 'aria-live': 'polite' });
    var reasonBox = h('div', { class: 'seg reasons', role: 'radiogroup', 'aria-label': '사유' }, L.reasonList(db.settings).map(function (r) {
      var x = h('input', { type: 'radio', name: 'rs-' + it.no, value: r }); if (dr ? dr.reason === r : cur && cur.reason === r) x.checked = true;
      return h('label', { class: 'pill' }, x, h('span', null, r));
    }));
    var reasonWrap = h('div', { class: 'field' + (E.reason ? ' invalid' : '') }, h('span', null, '사유 (차이가 있으면 필수)'), reasonBox, E.reason ? h('small', { class: 'err' }, E.reason) : null);
    var note = h('input', { value: dr ? dr.note : cur ? cur.note || '' : '', placeholder: '예: 일부가 B-03 칸에 있음, 포장 훼손', autocomplete: 'off' });
    function upd() {
      var v = L.toNum(qty.value);
      if (v == null || isNaN(v)) { diffOut.textContent = '시스템 수량 ' + N(it.sys) + ' ' + (it.unit || ''); diffOut.className = 'diff-out'; reasonWrap.hidden = true; return; }
      var d = L.round(v - it.sys, 6);
      diffOut.textContent = d === 0 ? '시스템 수량과 같습니다' : '시스템 ' + N(it.sys) + ' → 실물 ' + N(v) + ' · 차이수량 ' + L.fmtSigned(d) + '(실물 − 수량) · 사유를 골라 주세요';
      diffOut.className = 'diff-out ' + (d === 0 ? 'ok' : 'bad');
      reasonWrap.hidden = d === 0;
    }
    qty.addEventListener('input', upd);
    function step(n) { var v = L.toNum(qty.value); v = (v == null || isNaN(v) ? it.sys : v) + n; qty.value = String(Math.max(0, L.round(v, 6))); upd(); }
    var photoIn = h('input', { type: 'file', accept: 'image/*', capture: 'environment' });
    var photoNote = h('span', { class: 'note' }, ui.photo ? '사진 1장 준비됨(' + Math.round(ui.photo.length * 0.75 / 1024) + 'KB)' : cur && cur.photo && cur.photo !== 'omitted' ? '이전 사진 있음' : '');
    photoIn.addEventListener('change', function () {
      var f = photoIn.files[0]; if (!f) return;
      shrinkImage(f).then(function (d) { ui.photo = d; photoNote.textContent = '사진 1장 준비됨(' + Math.round(d.length * 0.75 / 1024) + 'KB)'; }, function (e) { toast(e.message, true); });
    });
    function submit(ev) {
      ev.preventDefault();
      if (!me.name) { toast('위 「작성자」에 이름을 먼저 넣어 주세요.', true); return; }
      var chk = reasonBox.querySelector('input:checked');
      var v = L.validateCount({ no: it.no, qty: qty.value, reason: chk ? chk.value : '', note: note.value, by: me.name, dev: me.dev, photo: ui.photo || (cur && cur.by === me.name && cur.photo !== 'omitted' ? cur.photo : null) }, it, { sessionId: db.session.id, reasons: L.reasonList(db.settings) });
      if (!v.ok) { ui.err = v.errors; ui.draft = { no: it.no, qty: qty.value, reason: chk ? chk.value : '', note: note.value }; render(); toast(v.errors.qty || v.errors.reason || v.errors.note || '확인할 칸이 있습니다.', true); return; }
      db.records.push(v.record);
      if (!save() && S.lastError() === 'quota') toast('저장 공간이 가득 찼습니다. 결과 파일을 내보내고 사진을 줄여 주세요.', true);
      else toast('저장 — ' + (it.code || it.no) + ' ' + N(v.record.qty) + (it.unit || '') + ' · ' + L.fmtTime(v.record.at) + (v.warnings.length ? ' (' + v.warnings[0] + ')' : ''));
      ui.err = null; ui.photo = null; ui.draft = null;
      // 다음 미확인 품목으로
      var res = L.resolveCounts(db.records, db.picks);
      var idx = list.indexOf(it), next = null;
      for (var i = idx + 1; i < list.length; i++) if (!res[list[i].no]) { next = list[i]; break; }
      ui.open = next ? next.no : '';
      render();
      if (next) setTimeout(function () { var x = document.getElementById('qty-' + next.no); if (x) { x.focus(); x.scrollIntoView({ block: 'center' }); } }, 30);
    }
    var form = h('form', { class: 'count-form', novalidate: true, onsubmit: submit },
      h('div', { class: 'meta-line note' }, [L.whLabel(it), it.site, it.loc, it.spec, it.mtype, it.mkind].filter(Boolean).join(' · ') + (L.managerOf(it, db.session) ? ' · 담당자 ' + L.managerOf(it, db.session) : '')),
      h('div', { class: 'field' + (E.qty ? ' invalid' : '') }, h('span', null, '실물수량' + (it.unit ? ' (' + it.unit + ')' : '')),
        h('div', { class: 'stepper' }, h('button', { type: 'button', class: 'btn', 'aria-label': '1 빼기', onclick: function () { step(-1); } }, '−'), qty, h('button', { type: 'button', class: 'btn', 'aria-label': '1 더하기', onclick: function () { step(1); } }, '+')),
        E.qty ? h('small', { class: 'err' }, E.qty) : null),
      h('div', { class: 'btn-row' }, h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { qty.value = String(it.sys); upd(); } }, '시스템 수량과 같음')),
      diffOut, reasonWrap,
      fld('특이사항', note, null, E.note),
      h('div', { class: 'photo-row' }, h('label', { class: 'btn btn-sm' }, photoIn, '현장 사진(필요할 때만)'), photoNote),
      h('p', { class: 'note' }, '확인 시각 ' + L.fmtTime(L.localIso()) + '(저장할 때 시각으로) · 작성자 ' + (me.name || '(이름 없음)') + (cur ? ' · 이전 기록: ' + N(cur.qty) + ' (' + cur.by + ', ' + L.fmtTime(cur.at) + ')' : '')),
      h('button', { type: 'submit', class: 'btn btn-primary btn-big btn-block' }, cur ? '다시 저장(정정)' : '저장하고 다음'));
    upd();
    return form;
  }
  function extraCard() {
    var whs = whOptions();
    var code = h('input', { placeholder: '라벨에 적힌 자재코드(있으면)', autocomplete: 'off' }), name = h('input', { placeholder: '예: 강판 절단 잔재', autocomplete: 'off' });
    var wh = h('select', null, whs.map(function (w) { return h('option', { value: w[0] }, w[1]); })), loc = h('input', { placeholder: '예: B-12-1, 3번 팔레트', autocomplete: 'off' });
    var qty = h('input', { inputmode: 'decimal', placeholder: '수량' }), unit = h('input', { placeholder: 'EA·KG …' });
    var rl = L.reasonList(db.settings);
    var reason = h('select', null, rl.map(function (r) { return h('option', { value: r }, r); })); reason.value = rl.indexOf('기타') >= 0 ? '기타' : rl[rl.length - 1];
    var note = h('input', { placeholder: '예: 미등록 잔재, 반납품' });
    function addExtra() {
      if (!me.name) { toast('위 「작성자」에 이름을 먼저 넣어 주세요.', true); return; }
      var v = L.validateCount({ qty: qty.value, reason: reason.value, note: note.value, by: me.name, dev: me.dev, extra: { code: code.value, name: name.value, wh: wh.value, loc: loc.value, unit: unit.value } }, null, { sessionId: db.session.id, reasons: rl });
      if (!v.ok) { toast(Object.keys(v.errors).map(function (k) { return v.errors[k]; })[0], true); return; }
      db.records.push(v.record); save(); toast('목록에 없는 자재를 적었습니다 — ' + (name.value || code.value)); render();
    }
    return h('details', { class: 'card' }, h('summary', null, '목록에 없는 자재 적기 (현장에서 발견)'),
      h('p', { class: 'note' }, 'ERP 에 없는데 현장에 있는 자재(잔재·반납품·라벨 없는 자재)를 적습니다. 시스템 수량 0 으로 차이에 들어갑니다.'),
      h('div', { class: 'form-grid two' }, fld('자재코드', code), fld('자재명', name), fld('창고(저장위치)', wh), fld('Location', loc), fld('수량', qty), fld('단위', unit), fld('사유', reason), fld('특이사항', note)),
      h('button', { type: 'button', class: 'btn btn-primary btn-block', style: 'margin-top:10px', onclick: addExtra }, '적기'));
  }
  function recentCard(mine) {
    if (!mine.length) return null;
    var list = mine.slice().sort(function (a, b) { return a.at < b.at ? 1 : -1; }).slice(0, 30);
    var byNo = {}; db.items.forEach(function (it) { byNo[it.no] = it; });
    return h('details', { class: 'card' }, h('summary', null, '이 기기 최근 기록 ' + Math.min(30, mine.length) + '건'),
      h('div', { class: 'entry-list' }, list.map(function (r) {
        var it = byNo[r.no];
        return h('div', { class: 'entry-row' }, h('div', null, h('b', null, r.no), ' ', it ? it.name : (r.extra ? r.extra.name : ''), h('div', { class: 'note' }, N(r.qty) + (it ? ' (시스템 ' + N(it.sys) + ')' : '') + (r.reason ? ' · ' + r.reason : '') + (r.note ? ' · ' + r.note : '') + ' · ' + L.fmtTime(r.at) + (r.photo ? ' · 사진' : ''))),
          it ? h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { ui.fieldFilter = 'all'; ui.q = r.no; ui.open = r.no; render(); window.scrollTo(0, 0); } }, '고치기') : null);
      })),
      h('p', { class: 'note' }, '고치면 새 기록이 하나 더 남고(정정), 보고서에는 마지막 기록이 들어갑니다. 이전 기록은 「기록 이력」에 남습니다.'));
  }

  /* ── ③ 취합 ─────────────────────────────── */
  function importOrderBook(sheets, fileName) {
    var w = L.parseOrderWorkbook(sheets, { fileName: fileName });
    if (w.error) { toast(fileName + ': ' + w.error, true); return; }
    if (w.session.id !== db.session.id) { toast(fileName + ': 다른 실사 회차의 작업지시서입니다(' + w.session.title + ').', true); return; }
    if (!w.records.length) { toast(fileName + ': 실물수량을 적은 줄이 없습니다.', true); return; }
    var m = L.mergeRecords(db.records, w.records);
    db.records = m.list;
    db.files.push({ name: fileName, records: m.added, dup: m.dup, at: L.localIso(), kind: '기록서 엑셀(SharePoint 취합)' });
    ui.mergeWarn = w.warnings.map(function (x) { return fileName + ': ' + x; });
    save(); render();
    toast(fileName + ' — ' + m.added + '건 합침' + (m.dup ? ', 이미 있던 ' + m.dup + '건은 건너뜀' : '') + (w.warnings.length ? ' · ' + w.warnings[0] : ''));
  }
  function handleMergeFiles(files) {
    Array.prototype.slice.call(files || []).forEach(function (f) {
      if (/\.json$/i.test(f.name)) {
        f.text().then(function (t) {
          var o; try { o = JSON.parse(t); } catch (e) { toast(f.name + ': JSON 을 읽지 못했습니다.', true); return; }
          var c = L.checkResultFile(o, db.session, db.items, f.name);
          if (c.error) { toast(f.name + ': ' + c.error, true); return; }
          var m = L.mergeRecords(db.records, c.records);
          db.records = m.list;
          db.files.push({ name: f.name, records: m.added, dup: m.dup, at: L.localIso(), kind: '결과 파일', by: c.by });
          save(); render();
          toast(f.name + ' — ' + m.added + '건 합침' + (m.dup ? ', 이미 있던 ' + m.dup + '건은 건너뜀' : '') + (c.unknown.length ? ', 지시서에 없는 재고번호 ' + c.unknown.length + '건 건너뜀' : ''));
        });
      } else if (/\.(xlsx|xls|xlsm)$/i.test(f.name)) {
        readBook(f).then(function (sheets) { importOrderBook(sheets, f.name); }, function (e) { toast(e.message, true); });
      } else toast(f.name + ': 기록서 엑셀(.xlsx)이나 결과 파일(.json)을 골라 주세요.', true);
    });
  }
  function pageMerge() {
    var head = pageHead('데이터 05 · 취합', '취합', '여러 사람이 채운 SharePoint 취합 엑셀(기록서)과 휴대폰 결과 파일을 합칩니다. 같은 기록은 한 번만 들어가고, 두 사람이 센 수량이 다르면 충돌로 보여 줍니다.');
    if (!hasSession()) return [head, emptyNotice()];
    var file = h('input', { type: 'file', accept: '.json,.xlsx,.xls,.xlsm', multiple: true });
    file.addEventListener('change', function () { handleMergeFiles(file.files); file.value = ''; });
    var an = L.analyze(db);
    var res = L.resolveCounts(db.records, db.picks);
    var conf = Object.keys(res).map(function (k) { return res[k]; }).filter(function (x) { return x.conflict || x.resolved; });
    var byNo = {}; db.items.forEach(function (it) { byNo[it.no] = it; });
    var asg = an.dims.assignee.slice().sort(function (x, y) { return x.key < y.key ? -1 : 1; });
    return [head,
      h('section', { class: 'card' }, h('h2', null, '파일 합치기'),
        h('div', { class: 'drop' }, h('label', { class: 'btn btn-primary' }, file, '기록서 엑셀 · 결과 파일 고르기(여러 개)'),
          h('span', { class: 'note' }, 'SharePoint 에서 내려받은 취합 엑셀(각자 자기 시트나 자기 줄에 적은 것) · 휴대폰의 결과 엑셀 · 실사결과_….json')),
        db.sample && !db.files.some(function (f) { return f.name === SMP.E_FILE; }) ? h('div', { class: 'btn-row' },
          h('button', { type: 'button', class: 'btn', onclick: mergeSampleE }, '샘플: SharePoint 취합 파일 합치기 체험'),
          h('span', { class: 'note' }, '실사자E 는 자기 시트에 40줄, 실사자C 는 전체 시트에 2줄 · 같은 파일: samples/' + SMP.E_FILE)) : null,
        (ui.mergeWarn || []).map(function (w) { return h('p', { class: 'alert info' }, w); }),
        db.files.length ? table(['파일', '종류', '합친 기록', '건너뜀(중복)', '시각'], db.files.map(function (f) { return [f.name, f.kind || '결과 파일', N(f.records), N(f.dup || 0), L.fmtTime(f.at)]; }), { num: [0, 0, 1, 1, 0] }) : h('p', { class: 'note' }, '아직 합친 파일이 없습니다.'),
        h('p', { class: 'note' }, '다른 실사 회차(지시서 ID 가 다른) 파일은 합치지 않습니다. 같은 SharePoint 파일을 하루에 여러 번 넣어도 새로 적힌 줄만 들어가고, 고친 줄은 정정으로 남습니다. 같은 사람이 전체 시트와 자기 시트에 다르게 적었으면 자기 시트 값을 씁니다. 실시간으로 바로 쌓이게 하는 것은 2단계입니다.')),
      h('section', { class: 'card' }, h('h2', null, '충돌 — 채택할 기록 고르기 (' + conf.filter(function (x) { return x.conflict; }).length + '건 남음)'),
        conf.length ? h('div', { class: 'conf-list' }, conf.map(function (x) {
          var it = byNo[x.no] || {};
          return h('div', { class: 'conf' + (x.conflict ? ' open' : '') }, h('div', null, h('b', null, x.no), ' ', it.name || '', h('span', { class: 'note' }, ' · ' + (it.wh || '') + ' ' + (it.loc || '') + ' · 시스템 ' + N(it.sys))),
            x.candidates.map(function (r) {
              var rd = h('input', { type: 'radio', name: 'pick-' + x.no, value: r.rid }); if (x.current.rid === r.rid && x.resolved) rd.checked = true;
              rd.addEventListener('change', function () { db.picks[x.no] = r.rid; save(); render(); toast(x.no + ' — ' + r.by + ' 의 ' + N(r.qty) + ' 을 채택했습니다.'); });
              return h('label', { class: 'opt' }, rd, h('span', null, h('span', { class: 't' }, N(r.qty) + ' — ' + r.by + (r.reason ? ' · ' + r.reason : '')), h('span', { class: 's' }, L.fmtTime(r.at) + ' · ' + (r.src || '') + (r.note ? ' · ' + r.note : '') + ' · 기록 ' + r.rid)));
            }),
            x.resolved ? h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { delete db.picks[x.no]; save(); render(); } }, '고른 것 취소') : null);
        })) : h('p', { class: 'note' }, '충돌이 없습니다.'),
        h('p', { class: 'note' }, '같은 사람이 다시 센 것은 충돌이 아니라 정정으로 보고 마지막 기록을 씁니다.')),
      h('section', { class: 'card' }, h('h2', null, '실사자별 진행'), personTable(asg, true))];
  }

  /* ── ④ 차이대사 ─────────────────────────── */
  function pageVariance() {
    var head = pageHead('데이터 03 · 04 · 차이 · 사유 확인', '차이대사', '시스템 수량과 실물수량을 맞대어 수량·금액 차이와 차이율을 봅니다. 기본은 허용 0 — 차이가 있으면 모두 사유를 확인합니다(설정에서 차이율·금액 허용으로 바꿀 수 있음).');
    if (!hasSession()) return [head, emptyNotice()];
    var an = L.analyze(db), S0 = an.S;
    var whSel = h('select', { 'aria-label': '창고' }, h('option', { value: '' }, '모든 창고'), unique(an.rows.map(function (r) { return r.wh; })).map(function (w) { return h('option', { value: w }, w); }));
    whSel.value = ui.varWh; whSel.addEventListener('change', function () { ui.varWh = whSel.value; render(); });
    var stSel = h('select', { 'aria-label': '보기' }, [['diff', '차이 있는 품목']].concat(S0.tolMode === 'range' ? [['over', '허용 초과만']] : []).concat([ ['noreason', '사유 미입력'], ['todo', '미실사'], ['conflict', '충돌'], ['all', '전체']]).map(function (x) { return h('option', { value: x[0] }, x[1]); }));
    stSel.value = ui.varState; stSel.addEventListener('change', function () { ui.varState = stSel.value; render(); });
    var q = h('input', { type: 'search', value: ui.varQ, placeholder: '자재코드·자재명·Location', 'aria-label': '찾기' });
    q.addEventListener('change', function () { ui.varQ = q.value; render(); });
    var rows = an.rows.filter(function (r) {
      if (ui.varWh && r.wh !== ui.varWh) return false;
      if (ui.varState === 'diff' && !(r.counted && r.diff !== 0)) return false;
      if (ui.varState === 'over' && !r.over) return false;
      if (ui.varState === 'noreason' && r.reason !== L.NO_REASON) return false;
      if (ui.varState === 'todo' && r.counted) return false;
      if (ui.varState === 'conflict' && !r.conflict) return false;
      return true;
    });
    if (ui.varQ) { var hit = L.searchItems(rows, ui.varQ); rows = rows.filter(function (r) { return hit.indexOf(r) >= 0; }); }
    rows.sort(function (a, b) { return Math.abs(b.diffAmt || 0) - Math.abs(a.diffAmt || 0) || (a.no < b.no ? -1 : 1); });
    var a = an.all;
    return [head,
      h('div', { class: 'tiles' },
        tile('실사 완료', N(a.counted) + ' / ' + N(a.items), pct(a.progress)),
        tile('차이 품목', N(a.diffItems), S0.tolMode === 'range' ? '허용 초과 ' + N(a.overItems) : '사유 미입력 ' + N(a.noReason)),
        tile('증가 / 감소', W(a.plusAmt), '감소 ' + W(Math.abs(a.minusAmt))),
        tile('금액 정확도', pct(a.amtAccuracy), '1 − 절대 차이 ÷ 장부금액')),
      h('section', { class: 'card filters no-print' }, fld('창고', whSel), fld('보기', stSel), fld('찾기', q),
        h('p', { class: 'note span-all' }, '허용 기준: ' + L.tolText(S0) + '(설정에서 바꿈). 「차이」 = 실물 − 시스템(+ 남음, − 모자람). 기록서 엑셀의 「차이수량」 · 보고서 PPT 도 같은 방향입니다.')),
      table(['자재코드', '창고·Location', '자재명', '시스템', '실물', '차이', '차이 금액', '차이율', '판정', '사유', '특이사항', '확인자·시각', '기록'], rows.slice(0, 300).map(function (r) {
        return { cls: r.conflict ? 'differs' : r.counted ? null : 'small', tdCls: [null, null, null, null, null, r.diff ? 'over' : null, null, null, r.over ? 'over' : null],
          cells: [r.code || r.no, r.wh + ' · ' + r.loc, r.name || r.code, N(r.sys), r.counted ? N(r.cnt) : '-', r.counted ? L.fmtSigned(r.diff) : '-', r.counted ? L.fmtSigned(r.diffAmt, true) : '-',
            r.rate == null ? (r.counted && r.diff ? '신규' : '-') : L.pctText(r.rate, 1), r.status + (r.conflict ? ' · 충돌' : ''), r.reason, r.note,
            r.counted ? r.by + ' ' + L.fmtTime(r.at) : '', r.counted ? [(r.recCount > 1 ? r.recCount + '건' : '1건'), photoCell(r)] : ''] };
      }), { num: [0, 0, 0, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0] }),
      h('p', { class: 'note' }, rows.length + '품목' + (rows.length > 300 ? ' 중 처음 300개' : '') + ' · 차이 금액 큰 순 · 모두 보려면 ⑥ 보고서의 엑셀 결과서(「전체 대사표」)')];
  }

  /* 차이대사 「사진」 — 누르면 크게 봄 (11-08 수강생 TEST: 사진 있음 표시만 있고 볼 수 없었음) */
  function photoCell(r) {
    if (r.photoSrc) return [' · ', h('button', { type: 'button', class: 'btn btn-sm photo-btn', onclick: function () { showPhoto(r); } },
      h('img', { src: r.photoSrc, alt: '', class: 'photo-thumb' }), r.photoOld ? '이전 사진' : '사진 보기')];
    if (r.photoOmitted) return h('span', { class: 'note', title: '결과 파일(JSON)을 「사진도 넣기」 없이 내보내 사진이 빠졌습니다. 찍은 휴대폰에서 「사진도 넣기」를 켜고 결과 파일을 다시 받아 ③ 취합에 넣으면 볼 수 있습니다.' }, ' · 사진(파일에 안 담김)');
    return null;
  }
  function showPhoto(r) {
    var dlg = h('dialog', { class: 'dialog photo-dialog', 'aria-label': '현장 사진' },
      h('div', { class: 'dialog-body' },
        h('h2', null, (r.code || r.no) + ' ' + (r.name || '')),
        h('p', { class: 'note' }, r.wh + ' · ' + r.loc + ' · 시스템 ' + N(r.sys) + ' / 실물 ' + N(r.cnt) + (r.reason ? ' · ' + r.reason : '') + ' · ' + r.by + ' ' + L.fmtTime(r.at) + (r.photoOld ? ' · 같은 품목의 이전 기록 사진' : '')),
        h('img', { src: r.photoSrc, alt: (r.name || r.no) + ' 현장 사진', class: 'photo-full' }),
        h('div', { class: 'dialog-actions' },
          h('a', { class: 'btn', href: r.photoSrc, download: '현장사진_' + String(r.code || r.no).replace(/[\\/:*?"<>|]/g, '_') + '.jpg' }, '사진 받기'),
          h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { dlg.close(); } }, '닫기'))));
    dlg.addEventListener('close', function () { dlg.remove(); });
    dlg.addEventListener('click', function (e) { if (e.target === dlg) dlg.close(); });
    document.body.appendChild(dlg);
    if (dlg.showModal) dlg.showModal(); else dlg.setAttribute('open', '');
  }

  /* ── ⑤ 분석(BI) ─────────────────────────── */
  function pageBi() {
    var head = pageHead('데이터 05 · 작업결과 분석', '분석(BI)', '창고별·창고위치(사내/사외)별·실사자별(파일에 있으면 자재유형·자재타입별도)로 진행률·정확도·차이 금액을, 사유별로 차이가 어디서 왔는지 봅니다.');
    if (!hasSession()) return [head, emptyNotice()];
    var an = L.analyze(db);
    var dims = L.DIM_SHEETS.filter(function (d) { return an.hasDim[d[0]]; });
    if (!an.hasDim[ui.dim]) ui.dim = 'wh';
    var tog = h('div', { class: 'view-toggle', role: 'group', 'aria-label': '기준' }, dims.map(function (d) {
      return h('button', { type: 'button', 'aria-pressed': ui.dim === d[0] ? 'true' : 'false', onclick: function () { ui.dim = d[0]; render(); } }, d[1]);
    }));
    var g = an.dims[ui.dim], label = L.DIMS[ui.dim].label, person = ui.dim === 'assignee';
    var man = function (v) { return L.round(v / 10000, 1); };
    var rs = an.reasons.filter(function (x) { return x.items; }).sort(function (a, b) { return b.absAmt - a.absAmt; });
    // 11-08 수강생: 창고가 많으면 그래프 이름이 겹쳐 안 보임 → 아래 표에서 고른 행만 그래프에 그림(처음엔 절대 차이 상위 8개)
    if (ui.biPickSid !== db.session.id) { ui.biPick = {}; ui.biPickSid = db.session.id; } // 새 실사 회차면 처음 값으로
    var picks = ui.biPick[ui.dim];
    if (!picks) picks = ui.biPick[ui.dim] = L.defaultPicks(g, BI_TOP);
    var shown = g.filter(function (x) { return picks.indexOf(x.key) >= 0; });
    function setPicks(list) { ui.biPick[ui.dim] = list; render(); }
    var lbl = function (x) { return x.label || x.key; };
    var chartNote = shown.length ? null : h('p', { class: 'note' }, '아래 표에서 그래프에 그릴 행을 골라 주세요.');
    return [head,
      h('section', { class: 'card no-print' }, tog),
      chartNote,
      h('div', { class: 'charts' },
        C.figure(label + ' 진행률·정확도(%) — ' + shown.length + '/' + g.length + '개', { type: 'bar', unit: '%', labels: shown.map(lbl), series: [
          { name: person ? '배정 진행률' : '진행률', values: shown.map(function (x) { return x.progress; }), color: 'var(--series-1)' },
          { name: '정확도', values: shown.map(function (x) { return x.accuracy; }), color: 'var(--series-3)' }] }),
        C.figure(label + ' 차이 금액(만원) — ' + shown.length + '/' + g.length + '개', { type: 'bar', unit: '', labels: shown.map(lbl), series: [
          { name: '증가', values: shown.map(function (x) { return man(x.plusAmt); }), color: 'var(--series-1)' },
          { name: '감소(절댓값)', values: shown.map(function (x) { return man(Math.abs(x.minusAmt)); }), color: 'var(--series-2)' }] })),
      h('div', { class: 'btn-row no-print bi-pick' },
        h('span', { class: 'note' }, '그래프에 그릴 행: 표 왼쪽 칸을 고르세요(' + shown.length + '개 고름).'),
        h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { setPicks(L.defaultPicks(g, BI_TOP)); } }, '절대 차이 상위 ' + BI_TOP + '개'),
        h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { setPicks(g.map(function (x) { return x.key; })); } }, '모두'),
        h('button', { type: 'button', class: 'btn btn-sm', onclick: function () { setPicks([]); } }, '모두 해제'),
        shown.length > 12 ? h('span', { class: 'note' }, '12개가 넘으면 이름이 잘립니다(막대에 마우스를 올리면 전체 이름).') : null),
      table(['그래프', label, person ? '배정 품목' : '품목', person ? '직접 센 품목' : '완료', person ? '배정 진행률' : '진행률', '정확도', '차이', '허용 초과', '장부금액', '증가', '감소', '순차이', '절대 차이'], g.map(function (x) {
        var cb = h('input', { type: 'checkbox', 'aria-label': lbl(x) + ' 그래프에 그리기' }); cb.checked = picks.indexOf(x.key) >= 0;
        cb.addEventListener('change', function () { var list = picks.filter(function (k) { return k !== x.key; }); if (cb.checked) list.push(x.key); setPicks(list); });
        return [cb, h('span', { title: x.key }, lbl(x)), x.items || !person ? N(x.items) : '-', N(x.counted), x.progress == null ? '-' : pct(x.progress), pct(x.accuracy), N(x.diffItems), N(x.overItems), W(x.sysAmt), W(x.plusAmt), W(Math.abs(x.minusAmt)), L.fmtSigned(x.netAmt, true), W(x.absAmt)];
      }).concat([{ cls: 'sum', cells: ['', '전체', N(an.all.items), N(an.all.counted), pct(an.all.progress), pct(an.all.accuracy), N(an.all.diffItems), N(an.all.overItems), W(an.all.sysAmt), W(an.all.plusAmt), W(Math.abs(an.all.minusAmt)), L.fmtSigned(an.all.netAmt, true), W(an.all.absAmt)] }]), { num: [0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1], cls: 'bi' }),
      person ? h('p', { class: 'note' }, '실사자별: 배정 품목 · 배정 진행률은 배정표 기준, 직접 센 품목 · 정확도 · 차이 · 금액은 실제로 센 사람(기록의 작성자 By) 기준입니다.') : null,
      h('h2', { class: 'sec', style: 'margin-top:20px' }, '사유별'),
      h('div', { class: 'charts' },
        C.figure('사유별 차이 품목 수', { type: 'hbar', unit: '개', labels: rs.map(function (x) { return x.key; }), series: [{ name: '품목', values: rs.map(function (x) { return x.items; }), color: 'var(--series-1)', colors: rs.map(function (x) { return x.key === L.NO_REASON ? 'var(--series-plan)' : 'var(--series-1)'; }) }] }),
        C.figure('사유별 절대 차이 금액(만원)', { type: 'hbar', unit: '', labels: rs.map(function (x) { return x.key; }), series: [{ name: '금액', values: rs.map(function (x) { return man(x.absAmt); }), color: 'var(--series-2)' }] })),
      table(['사유', '차이 품목', '비중', '증가', '감소', '순차이', '절대 차이', '금액 비중'], an.reasons.map(function (x) { return [x.key, N(x.items), pct(x.share), W(x.plusAmt), W(Math.abs(x.minusAmt)), L.fmtSigned(x.netAmt, true), W(x.absAmt), pct(x.amtShare)]; }), { num: [0, 1, 1, 1, 1, 1, 1, 1] }),
      h('p', { class: 'note' }, '순차이는 늘고 준 것이 서로 지워진 값이라, 라벨이 바뀐 짝처럼 +와 −가 함께 생기면 작아 보입니다. 관리 대상은 절대 차이로 봅니다.'),
      an.hours.length ? h('section', { class: 'card', style: 'margin-top:16px' }, C.figure('시간대별 확인 품목(누적)', { type: 'line', unit: '', labels: an.hours.map(function (x) { return x.label; }), series: [{ name: '누적', values: an.hours.map(function (x) { return x.cum; }), color: 'var(--series-1)' }] })) : null];
  }

  /* ── ⑥ 보고서 · AI ──────────────────────── */
  function pageReport() {
    var head0 = pageHead('데이터 06 · 보고서', '보고서 · AI 요약');
    if (!hasSession()) return [head0, emptyNotice()];
    var an = L.analyze(db);
    var last = db.reports.length ? db.reports[db.reports.length - 1] : null;
    var meta = last && last.sessionId === db.session.id ? last : null;
    var changed = meta && (meta.records !== an.records || meta.counted !== an.all.counted);
    var paper = reportPaper(an, meta);
    function issue() {
      var by = me.name || prompt('발행하는 사람 이름', '') || '';
      var m = L.nextReportMeta(db, { an: an, by: by });
      db.reports.push(m); save(); render(); toast('보고서 v' + m.ver + ' 을 발행했습니다. 엑셀·인쇄본에 버전이 찍힙니다.');
    }
    return [
      pageHead('데이터 06 · 보고서', '보고서 · AI 요약', '보고서를 「발행」하면 버전이 올라가고 그때의 기준 시각·기록 수가 남습니다. 결과서(엑셀)의 기록 ID 로 원래 현장 기록을 찾을 수 있습니다.',
        h('div', { class: 'btn-row no-print' },
          h('button', { type: 'button', class: 'btn btn-primary', onclick: issue }, meta ? 'v' + (meta.ver + 1) + ' 발행' : 'v1 발행'),
          h('button', { type: 'button', class: 'btn', onclick: function () { window.print(); } }, '인쇄 / PDF'),
          h('button', { type: 'button', class: 'btn', onclick: function () { exportReportXlsx(an, meta); } }, '결과서 엑셀'),
          h('button', { type: 'button', class: 'btn', onclick: function () { exportWord(paper, meta); } }, 'Word 로'),
          h('button', { type: 'button', class: 'btn', onclick: function () { copyText(L.buildReportText(db, an, meta, db.aiNotes)); } }, '글자로 복사'))),
      !meta ? h('p', { class: 'alert info no-print' }, '아직 발행하지 않은 초안입니다. 내보내기 전에 「v1 발행」을 눌러 버전을 박아 주세요.') :
        changed ? h('p', { class: 'alert warn no-print' }, 'v' + meta.ver + ' 발행 뒤 기록이 바뀌었습니다(발행 때 ' + meta.records + '건 → 지금 ' + an.records + '건). 새 버전을 발행해 주세요.') : null,
      an.conflicts.length ? h('p', { class: 'alert warn no-print' }, '미해결 충돌 ' + an.conflicts.length + '건이 있습니다(③ 취합).') : null,
      h('section', { class: 'card paper' }, paper),
      pptCard(an, meta),
      aiCard(an)];
  }
  function reportPaper(an, meta) {
    var s = db.session, a = an.all;
    var man = function (v) { return L.round(v / 10000, 1); };
    var rs = an.reasons.filter(function (x) { return x.items; }).sort(function (x, y) { return y.absAmt - x.absAmt; });
    function dimTable(d, label) {
      return [h('h3', null, label), table([L.DIMS[d].label, '품목', '진행률', '정확도', '차이', '허용 초과', '순차이', '절대 차이'], an.dims[d].map(function (x) {
        return [x.key, N(x.items), pct(x.progress), pct(x.accuracy), N(x.diffItems), N(x.overItems), L.fmtSigned(x.netAmt, true), W(x.absAmt)];
      }), { num: [0, 1, 1, 1, 1, 1, 1, 1] })];
    }
    return h('div', { class: 'rp' },
      h('h1', null, '재고실사 결과 보고 — ' + s.title),
      table(['보고서 버전', '기준 일자·시각(ERP 재고)', '지시서 ID', '현장 기록', '작성'], [[L.verText(meta), L.fmtTime(s.baseAt), s.id, N(an.records) + '건', '경영관리팀 · 자동 생성 초안(검토 후 사용)']]),
      h('h2', null, '1. 총괄'),
      table(['품목', '실사 완료', '진행률', '일치(정확도)', '차이 품목', '허용 초과', '장부금액', '증가', '감소', '순차이'],
        [[N(a.items), N(a.counted), pct(a.progress), N(a.matched) + ' (' + pct(a.accuracy) + ')', N(a.diffItems), N(a.overItems), W(a.sysAmt), W(a.plusAmt), W(Math.abs(a.minusAmt)), L.fmtSigned(a.netAmt, true)]], { num: [1, 1, 1, 1, 1, 1, 1, 1, 1, 1] }),
      h('p', { class: 'note' }, '허용 기준: ' + L.tolText(an.S) + '. 미해결 충돌 ' + an.conflicts.length + '건.'),
      an.excluded.length ? h('p', { class: 'note' }, '이번 실사 제외: ' + an.excluded.map(function (x) { return L.whLabel(x) + ' ' + N(x.items) + '품목 ' + W(x.sysAmt) + (x.note ? '(' + x.note + ')' : ''); }).join(' · ')) : null,
      h('h2', null, '2. 창고 · 창고위치' + (an.hasDim.mtype ? ' · 자재유형' : '') + (an.hasDim.mkind ? ' · 자재타입' : '') + '별'),
      h('div', { class: 'charts' },
        (function () {
          var top = L.defaultPicks(an.dims.wh, BI_TOP), wl = an.dims.wh.filter(function (x) { return top.indexOf(x.key) >= 0; });
          return C.figure('창고별 차이 금액(만원)' + (an.dims.wh.length > wl.length ? ' — 절대 차이 상위 ' + wl.length + '곳' : ''), { type: 'bar', labels: wl.map(function (x) { return x.label || x.key; }), series: [
            { name: '증가', values: wl.map(function (x) { return man(x.plusAmt); }), color: 'var(--series-1)' },
            { name: '감소(절댓값)', values: wl.map(function (x) { return man(Math.abs(x.minusAmt)); }), color: 'var(--series-2)' }] });
        })(),
        C.figure('사유별 절대 차이 금액(만원)', { type: 'hbar', labels: rs.map(function (x) { return x.key; }), series: [{ name: '금액', values: rs.map(function (x) { return man(x.absAmt); }), color: 'var(--series-2)' }] })),
      L.DIM_SHEETS.filter(function (d) { return d[0] !== 'assignee' && an.hasDim[d[0]]; }).map(function (d) { return dimTable(d[0], d[1]); }),
      h('h2', null, '3. 사유별'),
      table(['사유', '차이 품목', '비중', '절대 차이', '금액 비중'], rs.map(function (x) { return [x.key, N(x.items), pct(x.share), W(x.absAmt), pct(x.amtShare)]; }), { num: [0, 1, 1, 1, 1] }),
      h('h2', null, '4. 차이 금액 상위 ' + an.top.length + '품목 (원천 기록 연결)'),
      table(['자재코드', '자재명', '창고·Location', '시스템', '실물', '차이', '차이 금액', '사유', '확인자·시각', '기록 ID'], an.top.map(function (r) {
        return [r.code || r.no, r.name || r.code, r.wh + ' · ' + r.loc, N(r.sys), N(r.cnt), L.fmtSigned(r.diff), L.fmtSigned(r.diffAmt, true), r.reason, r.by + ' ' + L.fmtTime(r.at), r.rid];
      }), { num: [0, 0, 0, 1, 1, 1, 1, 0, 0, 0] }),
      h('p', { class: 'note' }, '기록 ID 는 결과서 엑셀 「기록 이력」 시트의 원래 입력(누가·언제·어느 결과 파일로)과 이어집니다.'),
      db.aiNotes ? [h('h2', null, '5. 요약 의견 (AI 초안, 검토 후 사용)'), h('div', { class: 'pre' }, db.aiNotes)] : null);
  }
  function exportReportXlsx(an, meta) {
    var sheets = L.reportSheets(db, an, meta);
    sheets.forEach(function (x, i) { if (i === 0) x.cols = [{ wch: 24 }, { wch: 70 }]; });
    writeSheets('재고실사_결과서_' + db.session.baseAt.slice(0, 10) + '_' + (meta ? 'v' + meta.ver : '초안') + '.xlsx', sheets);
  }
  function exportWord(paper, meta) {
    var clone = paper.cloneNode(true);
    clone.querySelectorAll('.charts').forEach(function (x) { x.remove(); });
    var html = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word"><head><meta charset="utf-8"><title>재고실사 결과 보고</title>' +
      '<style>body{font-family:"맑은 고딕",sans-serif;font-size:10.5pt} table{border-collapse:collapse;margin:6px 0} th,td{border:1px solid #999;padding:3px 6px;font-size:9pt} th{background:#eef2f7} h1{font-size:16pt} h2{font-size:12.5pt;margin-top:14pt} h3{font-size:11pt} .pre{white-space:pre-wrap}</style></head><body>' +
      clone.innerHTML + '</body></html>';
    download('재고실사_결과보고_' + db.session.baseAt.slice(0, 10) + '_' + (meta ? 'v' + meta.ver : '초안') + '.doc', new Blob(['﻿' + html], { type: 'application/msword' }));
  }
  /* 보고서 PPT — 직전 보고서(pptx)를 올리면 디자인은 그대로 두고 내용만 바꿈(js/pptx-report.js). 없으면 기본 양식 */
  function pptCard(an, meta) {
    var P = window.STPptx, Z = window.JSZip;
    if (!P || !Z) return null;
    var o = db.pptOpts || (db.pptOpts = {});
    var cur = ui.ppt; // { name, buf, info } — 올린 파일은 이 화면에만 들고 있음(새로 고침하면 다시 올림)
    if (!cur && !ui.pptDef) { ui.pptDef = 'loading'; P.inspect(Z, null).then(function (i) { ui.pptDef = i; render(); }, function () { ui.pptDef = null; }); }
    var info = cur ? cur.info : ui.pptDef && ui.pptDef !== 'loading' ? ui.pptDef : null;
    var file = h('input', { type: 'file', accept: '.pptx', 'aria-label': '직전 보고서 pptx' });
    file.addEventListener('change', function () {
      var f = file.files[0]; file.value = ''; if (!f) return;
      f.arrayBuffer().then(function (buf) { return P.inspect(Z, buf).then(function (i) { ui.ppt = { name: f.name, buf: buf, info: i }; ui.pptOv = {}; ui.pptLog = null; render(); toast('「' + f.name + '」 구조를 읽었습니다. 아래에서 확인한 뒤 「PPT 만들기」를 눌러 주세요.'); }); })
        .catch(function (e) { toast('PPT 를 읽지 못했습니다: ' + e.message, true); });
    });
    var data = P.reportData(an, db);
    var ov = ui.pptOv || (ui.pptOv = {});
    var draft = P.issueTextOf(P.issueDraft(an, info && info.issue ? info.issue.rows : 2));
    var issues = h('textarea', { rows: 5, class: 'mono', 'aria-label': '3장 이슈 · 개선' }); issues.value = o.issuesText || draft;
    issues.addEventListener('change', function () { o.issuesText = issues.value; save(); });
    function sel(label, key, opts, hint) {
      var s = h('select', null, opts.map(function (x) { var e = h('option', { value: x[0] }, x[1]); if ((o[key] || opts[0][0]) === x[0]) e.selected = true; return e; }));
      s.addEventListener('change', function () { o[key] = s.value; save(); });
      return fld(label, s, hint);
    }
    var dimOpts = [['auto', '표의 행 이름에 맞춤(자동)'], ['wh', '창고(저장위치명)'], ['site', '창고위치(사내/사외)']]
      .concat(an.hasDim.mkind ? [['mkind', '자재타입']] : []).concat(an.hasDim.mtype ? [['mtype', '자재유형']] : []);
    var rpp = h('input', { type: 'number', min: 0, max: 40, step: 1, value: o.rowsPerPage || '', placeholder: info && info.diff ? '자동(' + info.diff.cap + '줄)' : '자동', inputmode: 'numeric' });
    rpp.addEventListener('change', function () { o.rowsPerPage = Math.max(0, Math.min(40, +rpp.value || 0)); save(); });
    var memo = h('input', { type: 'checkbox' }); memo.checked = o.removeMemos !== false;
    memo.addEventListener('change', function () { o.removeMemos = memo.checked; save(); });
    var ovRows = info && info.overview ? info.overview.rows.map(function (x) {
      var k = x.r + ',' + x.c;
      var inp = h('input', { value: ov[k] != null ? ov[k] : x.auto && data.overview[x.auto] ? data.overview[x.auto] : x.text, autocomplete: 'off' });
      inp.addEventListener('change', function () { ov[k] = inp.value; });
      return fld(x.label || '(이름 없음)', inp, x.auto ? '도구가 채움(' + (x.auto === 'baseAt' ? '기준 일자·시각' : '첫 ~ 마지막 확인 시각') + ') · 고칠 수 있음' : '직전 보고서 내용 그대로 · 고칠 수 있음');
    }) : [];
    function makePpt() {
      var opts = { statusBy: o.statusBy || 'auto', diffBy: o.diffBy || 'wh', rowsPerPage: o.rowsPerPage || 0, removeMemos: o.removeMemos !== false, overviewEdits: ov, issues: P.parseIssueText(issues.value), type: 'blob' };
      P.make(Z, cur ? cur.buf : null, an, db, opts).then(function (r) {
        ui.pptLog = r.log;
        download('재고실사_결과보고_' + db.session.baseAt.slice(0, 10) + '_' + (meta ? 'v' + meta.ver : '초안') + '.pptx', r.out);
        render(); toast('보고서 PPT 를 만들었습니다 — 2장 ' + r.log.diffPages + '장(창고 ' + r.log.groups + '곳).');
      }).catch(function (e) { toast('PPT 를 만들지 못했습니다: ' + e.message, true); });
    }
    function found(x, yes, no) { return h('li', { class: x ? null : 'miss' }, x ? yes : [h('strong', null, '못 찾음 — '), no]); }
    var lg = ui.pptLog;
    return h('section', { class: 'card no-print' }, h('h2', null, '보고서 PPT (직전 보고서 내용만 바꾸기)'),
      h('p', null, '지난번 보고서(pptx)를 올리면 디자인 · 마스터 · 글꼴은 그대로 두고 숫자와 표만 이번 실사로 바꿉니다. 올리지 않으면 기본 양식으로 만듭니다. 파일은 이 브라우저 안에서만 읽습니다.'),
      h('div', { class: 'drop' }, h('label', { class: 'btn btn-primary' }, file, '직전 보고서 pptx 고르기'),
        cur ? h('button', { type: 'button', class: 'btn', onclick: function () { ui.ppt = null; ui.pptOv = {}; ui.pptLog = null; render(); } }, '기본 양식으로') : null),
      h('p', { class: 'note' }, cur ? '양식: 「' + cur.name + '」 (슬라이드 ' + info.slides + '장)' : '양식: 기본 양식(도구에 들어 있음)'),
      info ? h('ul', { class: 'note' },
        found(info.overview, '1장 개요 표 — 실사 시행일 · 집계 기준 시점을 채움, 실사 방법 칸은 직전 내용 그대로(아래에서 고침)', '1장 개요 표(실사 시행일 · 집계 기준 시점)를 찾지 못함'),
        found(info.status, info.status ? '1장 실사 현황 표 — 행 이름: ' + (info.status.labels.filter(Boolean).join(', ') || '(비어 있음)') + (info.status.unit ? ' · ' + info.status.unit : '') : '', '1장 실사 현황 표(전체재고 · 실사완료 · 미실시)를 찾지 못함'),
        found(info.diff, info.diff ? '2장 차이 내용 — 차이 슬라이드 ' + info.diffSlides.length + '장 중 첫 장을 본떠 창고마다 새로 만듦(한 장 ' + info.diff.cap + '줄, 넘치면 이어지는 장) · 열: ' + info.diff.head.join(' / ') : '', '2장 차이 내용 표(전산수량 · 실물수량 · 차이수량)를 찾지 못해 2장은 그대로 둠'),
        found(info.issue, info.issue ? '3장 이슈사항 및 개선 — 칸 ' + info.issue.rows + '줄에 아래 초안을 넣음(PPT 에서 고칠 수 있음)' : '', '3장 이슈 슬라이드(「이슈 사항」 머리 상자)를 찾지 못함'),
        info.memos ? h('li', null, '작성 메모 상자 ' + info.memos + '개(「(계산로직)」 · 「AI가 초안 …」) — 아래 선택에 따라 지움') : null) : h('p', { class: 'note' }, '양식을 읽는 중…'),
      h('div', { class: 'form-grid' },
        sel('실사 현황 표의 행', 'statusBy', dimOpts, '자동: 직전 표의 행 이름(예: 원자재 · 반제품)이 자재타입이면 자재타입, 창고명이면 창고. 맞는 것이 없으면 창고.'),
        sel('2장 슬라이드 나누기', 'diffBy', [['wh', '창고(저장위치명)마다 한 장'], ['site', '창고위치(사내/사외)마다 한 장']]),
        fld('2장 한 장에 넣을 줄 수', rpp, '비우면 직전 보고서의 줄 높이로 자동. 넘치는 품목은 「(2/3)」 처럼 이어지는 장으로.'),
        h('label', { class: 'opt' }, memo, h('span', null, h('span', { class: 't' }, '작성 메모 상자 지우기'), h('span', { class: 's' }, '「(계산로직)」 처럼 만드는 사람에게 남긴 빨간 메모'))) ),
      ovRows.length ? [h('h3', null, '1장 개요 칸'), h('div', { class: 'form-grid' }, ovRows)] : null,
      h('h3', null, '3장 이슈사항 및 개선 (초안 — 「주제 | 이슈 | 개선」 한 줄씩, 한 칸 안 줄바꿈은 「 ; 」)'), issues,
      h('div', { class: 'btn-row', style: 'margin:10px 0' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: makePpt }, 'PPT 만들기'),
        h('button', { type: 'button', class: 'btn', onclick: function () { o.issuesText = ''; save(); render(); } }, '3장 초안 다시 채우기')),
      h('p', { class: 'note' }, '차이수량 = 실물수량 − 전산수량(모자라면 −), 차이금액 = 차이수량 × 단가, 0 이 아니면 빨간 글자, 합계 줄은 차이수량 · 차이금액만. 실사 현황의 「실사완료」는 현장 입력이 된 품목의 실물 수량과 실물 수량 × ERP 단가, 「미실시」는 아직 안 센 품목의 전산 수량 · 금액, 행 기준은 자재타입이 있으면 자재타입, 실사인원 0 창고는 「-」.'),
      lg ? h('p', { class: 'alert info' }, '만든 PPT: 슬라이드 ' + lg.slides + '장 · 2장 ' + lg.diffPages + '장(창고 ' + lg.groups + '곳, 한 장 ' + (lg.cap || '-') + '줄) · 실사 현황 행 기준 ' + ({ wh: '창고', site: '창고위치', mkind: '자재타입', mtype: '자재유형' }[lg.statusDim] || '-') + (lg.removedMemos ? ' · 메모 상자 ' + lg.removedMemos + '개 지움' : '')) : null,
      lg && lg.warnings.length ? lg.warnings.map(function (w) { return h('p', { class: 'alert warn' }, w); }) : null);
  }
  function aiCard(an) {
    var mask = db.mask !== false;
    var withItem = !!ui.aiWithItem;
    var info = L.buildAiPromptInfo(an, { mask: mask, withItem: withItem });
    var box = h('textarea', { rows: 12, readonly: true, class: 'mono' }); box.value = info.prompt;
    var answer = h('textarea', { rows: 8, placeholder: 'AI 답을 여기에 붙여 넣어 주세요. 「담당1」 같은 별칭은 원래 이름으로 되돌려 보고서 「요약 의견」에 넣습니다.' });
    answer.value = db.aiNotes || '';
    function back(t) { return mask ? L.unmaskText(t, info.map) : t; }
    answer.addEventListener('change', function () { db.aiNotes = back(answer.value); answer.value = db.aiNotes; save(); render(); });
    var mk = h('input', { type: 'checkbox' }); mk.checked = mask; mk.addEventListener('change', function () { db.mask = mk.checked; save(); render(); });
    var wi = h('input', { type: 'checkbox' }); wi.checked = withItem; wi.addEventListener('change', function () { ui.aiWithItem = wi.checked; render(); });
    return h('section', { class: 'card no-print' }, h('h2', null, 'AI 요약 (반자동)'),
      h('ol', { class: 'steps' }, h('li', null, '프롬프트를 복사해 회사가 허용한 AI 에 붙여 넣어 주세요. 도구가 계산한 숫자만 들어 있습니다.'), h('li', null, '받은 답을 아래 칸에 붙여 넣으면 보고서 「요약 의견」에 들어갑니다. 꼭 검토해 주세요.')),
      h('div', { class: 'mask-box' },
        h('label', { class: 'opt' }, mk, h('span', null, h('span', { class: 't' }, '실사자 이름 가리기'), h('span', { class: 's' }, '「담당1」 같은 별칭으로 보냅니다. 대응표는 이 기기에만 있습니다.'))),
        h('label', { class: 'opt' }, wi, h('span', null, h('span', { class: 't' }, '상위 품목의 자재코드·자재명 넣기'), h('span', { class: 's' }, '기본은 넣지 않고 창고·숫자(파일에 있으면 유형·타입)만 보냅니다.'))),
        info.leaks.length ? h('p', { class: 'alert warn' }, '가리지 못한 이름: ' + info.leaks.join(', ')) : null),
      box,
      h('div', { class: 'btn-row', style: 'margin:10px 0' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { copyText(info.prompt); } }, '프롬프트 복사'),
        AIP.sendButton('설정한 AI 서버로 보내기(선택)', function () { return info.prompt; }, function (t) { db.aiNotes = back(t); save(); render(); toast('AI 답을 받았습니다. 검토 후 쓰세요.'); },
          { toast: toast, settingsHref: '#/settings', system: '너는 자재관리·회계감사 보조야. 주어진 사실 밖의 숫자를 만들지 마.' })),
      fld('AI 답', answer));
  }

  /* ── 설정 ───────────────────────────────── */
  function pageSettings() {
    var s = L.cleanSettings(db.settings);
    var tp = h('input', { type: 'number', min: 0, step: 0.1, value: s.tolPct, inputmode: 'decimal' });
    var ta = h('input', { type: 'number', min: 0, step: 1000, value: s.tolAmt, inputmode: 'numeric' });
    var tn = h('input', { type: 'number', min: 5, step: 1, value: s.topN, inputmode: 'numeric' });
    var modeSel = h('select', null, h('option', { value: 'zero' }, '허용 0 — 차이가 있으면 모두 사유 필수(기본)'), h('option', { value: 'range' }, '차이율 · 금액 허용 범위 쓰기(선택)'));
    modeSel.value = s.tolMode;
    function syncMode() { tp.disabled = ta.disabled = modeSel.value !== 'range'; }
    modeSel.addEventListener('change', syncMode); syncMode();
    var rsBox = h('textarea', { rows: 7, 'aria-label': '사유 목록' }); rsBox.value = L.reasonList(s).join('\n');
    var nm = h('input', { value: me.name || '', placeholder: '작성자 이름' });
    var backup = h('input', { type: 'file', accept: '.json' });
    backup.addEventListener('change', function () {
      var f = backup.files[0]; if (!f) return;
      f.text().then(function (t) {
        var o = JSON.parse(t); if (!o || !Array.isArray(o.records) || !Array.isArray(o.items)) throw new Error('이 도구의 백업 파일이 아닙니다(결과 파일은 ③ 취합에서 합칩니다).');
        if (!confirm('지금 자료를 백업 파일 내용으로 바꿀까요?')) return;
        db = L.restoreDb(o); save(); toast('백업을 불러왔습니다.'); render();
      }).catch(function (e) { toast('불러오지 못했습니다: ' + e.message, true); });
      backup.value = '';
    });
    return [pageHead('설정', '허용 기준 · 사유 · 이 기기 · AI 연결 · 백업'),
      h('section', { class: 'card' }, h('h2', null, '차이 허용 기준 · 사유 목록'),
        h('div', { class: 'form-grid' }, h('div', { class: 'span-2' }, fld('허용 기준', modeSel, '수강생 답: 「차이가 있으면 무조건 사유 필수」 — 그래서 기본은 허용 0 입니다.')), fld('보고서 상위 품목 수', tn)),
        h('div', { class: 'form-grid' }, fld('차이율 (%)', tp, '허용 범위를 쓸 때만 — 이보다 크면 「허용 초과」'), fld('차이 금액 (원)', ta, '허용 범위를 쓸 때만 — 이보다 크면 「허용 초과」')),
        h('p', { class: 'note' }, '어느 쪽이든 현장 입력에서 차이가 있으면 사유를 꼭 골라야 저장됩니다. 허용 범위는 「허용 초과」 표시와 보고서 집계만 바꿉니다. 현장 사진은 필요할 때만 붙입니다.'),
        fld('사유 목록 (한 줄에 하나, 12개까지)', rsBox, '기록서 견본 메모 「설정에서 사유 리스트를 작성」 — 고친 목록은 기록서 엑셀 「지시정보」에 실려 휴대폰에도 같이 갑니다. 「기타」를 두면 고를 때 특이사항이 필수입니다.'),
        h('div', { class: 'btn-row', style: 'margin-top:10px' },
          h('button', { type: 'button', class: 'btn btn-primary', onclick: function () {
            var list = rsBox.value.split(/\n/).map(function (x) { return x.trim(); }).filter(Boolean);
            var o = { tolMode: modeSel.value, tolPct: tp.value, tolAmt: ta.value, topN: tn.value };
            if (list.length && list.join('/') !== L.REASONS.join('/')) o.reasons = list;
            db.settings = L.cleanSettings(o); save(); toast('저장했습니다.'); render();
          } }, '저장'),
          h('button', { type: 'button', class: 'btn', onclick: function () { rsBox.value = L.REASONS.join('\n'); } }, '사유 기본 6개로'))),
      h('section', { class: 'card' }, h('h2', null, '이 기기'),
        h('div', { class: 'form-grid' }, fld('작성자 이름', nm, '현장 기록에 자동으로 붙습니다'), fld('기기 ID', h('input', { value: me.dev, readonly: true }), '결과 파일에서 어느 기기 기록인지 가립니다')),
        h('div', { class: 'btn-row', style: 'margin-top:10px' }, h('button', { type: 'button', class: 'btn', onclick: function () { me.name = nm.value.trim(); S.saveMe(me); toast('저장했습니다.'); render(); } }, '저장'))),
      AIP.settingsCard({ toast: toast, onChange: render }),
      h('section', { class: 'card' }, h('h2', null, '백업 · 지우기'),
        h('p', { class: 'note' }, '자료는 이 기기 브라우저에만 있습니다. 브라우저 기록을 지우기 전에 JSON 백업을 받아 두세요(사진도 들어 있어 클 수 있습니다).'),
        h('div', { class: 'btn-row' },
          h('button', { type: 'button', class: 'btn', onclick: function () { download('재고실사_백업_' + today() + '.json', new Blob([JSON.stringify(db)], { type: 'application/json' })); } }, 'JSON 백업 받기'),
          h('label', { class: 'btn' }, backup, 'JSON 백업 불러오기'),
          h('button', { type: 'button', class: 'btn btn-danger', onclick: function () { if (!confirm('이 기기의 실사 자료를 모두 지울까요? (작성자 이름·AI 설정은 남습니다)')) return; db = L.emptyDb(); save(); render(); toast('지웠습니다.'); } }, '모든 자료 지우기')))];
  }

  /* ── 라우터 ─────────────────────────────── */
  var PAGES = [
    { id: 'home', no: '0', t: '진행 현황', s: '실사 회차 · 실사자별', fn: pageHome },
    { id: 'order', no: '①', t: '작업지시', s: 'ERP 엑셀 · 기준 시각 · 배정', fn: pageOrder },
    { id: 'field', no: '②', t: '현장 입력', s: '휴대폰 · 수량 · 사유', fn: pageField },
    { id: 'merge', no: '③', t: '취합', s: 'SharePoint 파일 합치기 · 충돌', fn: pageMerge },
    { id: 'variance', no: '④', t: '차이대사', s: '수량 · 금액 · 차이율', fn: pageVariance },
    { id: 'bi', no: '⑤', t: '분석(BI)', s: '창고 · 사내/사외 · 사유', fn: pageBi },
    { id: 'report', no: '⑥', t: '보고서', s: '버전 · 원천 연결 · AI', fn: pageReport },
    { id: 'settings', no: '설', t: '설정', s: '허용 기준 · 사유 · 백업', fn: pageSettings }
  ];
  function route() { var id = (location.hash || '').replace(/^#\//, '').split('?')[0]; return PAGES.filter(function (p) { return p.id === id; })[0] || (hasSession() ? PAGES[0] : PAGES[1]); }
  function render() {
    var p = route();
    var nav = document.getElementById('nav'); nav.innerHTML = '';
    PAGES.forEach(function (x) {
      nav.appendChild(h('a', { href: '#/' + x.id, 'aria-current': x === p ? 'page' : null }, h('span', { class: 'no' }, x.no), h('span', { class: 't' }, x.t), h('span', { class: 's' }, x.s)));
    });
    var chip = document.getElementById('dataChip');
    chip.textContent = hasSession() ? db.session.title + ' · 기준 ' + L.fmtTime(db.session.baseAt) + ' · ' + N(db.items.length) + '품목 · 기록 ' + N(db.records.length) + '건' : '실사 회차 없음';
    var b = document.getElementById('sampleBanner');
    b.hidden = !db.sample;
    b.textContent = '지금 보시는 것은 가상 샘플입니다(실제 양식과 같은 모양, 창고·자재·사람·수량·단가는 모두 지어낸 값). 소모품창고 입출고 누락, 부품창고 B 구역 위치 이동·분산, 제품창고 라벨이 바뀐 짝, 충돌 2건, 외주창고 진행 지연, 시험실창고 실사 제외를 일부러 심었습니다. 실제 파일을 쓰기 전에 설정에서 「모든 자료 지우기」를 해 주세요.';
    main.innerHTML = '';
    try { add(main, p.fn()); }
    catch (e) { console.error(e); main.appendChild(h('div', { class: 'alert warn' }, '화면을 그리지 못했습니다: ' + e.message)); }
    document.title = p.t + ' — 재고실사 Start-End';
    save();
  }
  window.addEventListener('hashchange', function () { ui.open = ''; render(); main.focus({ preventScroll: true }); window.scrollTo(0, 0); });
  render();
})();
