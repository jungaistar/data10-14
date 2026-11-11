// 헤드리스 브라우저로 주요 흐름 점검 (선택 — playwright-core 가 있는 PC 에서만)
//   PLAYWRIGHT_CORE=/경로/node_modules/playwright-core node scripts/check-browser.mjs
// file:// 로 index.html 을 열어:
//   ① 사무실 — samples/ 의 ERP 엑셀 올리기 → 기준 시각 없이는 안 만들어짐 → 만들기 → 창고 배정표 올리기(제외 창고 · 반씩 나누기)
//      → 기록서 엑셀(SharePoint 취합용) 내려받아 다시 읽기(시트 · 열 · 차이수량 수식)
//   ② 한 파일에 ERP 시트 + 배정표 시트가 함께 있는 경우(수강생 첨부 모양)도 한 번에 적용되는지
//   ③ 휴대폰(360px) — 받은 기록서 엑셀 넣기 → 사유 없이 저장 막힘 → 사유 고르면 저장 → 「결과 엑셀 받기」
//   ④ 사무실 — 그 결과 엑셀 합치기 → 같은 파일 다시 넣으면 0건
//   ⑤ 가상 샘플 → SharePoint 취합 파일 합치기 체험 → 충돌 고르기 → 설정(허용 0 기본 · 사유 목록) → 보고서 발행 · 결과서 엑셀
//   ⑥ 8개 화면 × 360 · 1280px 가로 넘침. JS 오류가 하나라도 있으면 실패.
// BREAK=1 이면 일부러 넓은 상자를 넣어 넘침 점검기가 잡는지 봅니다.
import { createRequire } from 'node:module';
import path from 'node:path';
import fs from 'node:fs';
import os from 'node:os';
import { fileURLToPath } from 'node:url';
const require = createRequire(import.meta.url);
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const { chromium } = require(process.env.PLAYWRIGHT_CORE || 'playwright-core');
const XLSX = require(path.join(root, 'vendor/xlsx.full.min.js'));
const Sample = require(path.join(root, 'js/sample-data.js'));
const url = 'file://' + path.join(root, 'index.html');
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'd1014-'));

const errors = [], problems = [];
const browser = await chromium.launch();
function ok(cond, msg) { if (!cond) { problems.push(msg); console.log('  FAIL ' + msg); } else console.log('  ok  ' + msg); }
function watch(page) {
  page.on('pageerror', e => errors.push('pageerror: ' + e.message));
  page.on('console', m => { if (m.type() === 'error') errors.push('console: ' + m.text()); });
  page.on('dialog', d => d.accept());
}
async function open(width) {
  const ctx = await browser.newContext({ acceptDownloads: true, viewport: { width, height: 900 } });
  const page = await ctx.newPage(); watch(page);
  await page.goto(url); await page.waitForTimeout(400);
  return page;
}
const go = async (page, hash) => { await page.evaluate(h => { location.hash = h; }, hash); await page.waitForTimeout(250); };
const click = async (page, sel, text) => { await page.locator(sel, text ? { hasText: text } : undefined).first().click(); await page.waitForTimeout(300); };
const db = page => page.evaluate(() => JSON.parse(localStorage.getItem('data10-14.db')));
async function dl(page, fn) {
  const [d] = await Promise.all([page.waitForEvent('download'), fn()]);
  const f = path.join(tmp, d.suggestedFilename());
  await d.saveAs(f);
  return f;
}
function book(file) {
  const wb = XLSX.read(fs.readFileSync(file), { type: 'buffer', cellDates: false });
  return { wb, sheets: wb.SheetNames.map(n => ({ name: n, aoa: XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null }) })) };
}
async function overflow(page) {
  return page.evaluate(() => {
    const W = document.documentElement.clientWidth, bad = [];
    for (const el of document.querySelectorAll('body *')) {
      const r = el.getBoundingClientRect();
      if (!r.width || !r.height) continue;
      let p = el.parentElement, inner = false;
      while (p) { const ox = getComputedStyle(p).overflowX; if (ox === 'auto' || ox === 'scroll') { inner = true; break; } p = p.parentElement; }
      if (!inner && r.right > W + 1) bad.push(el.tagName.toLowerCase() + '.' + el.className + ' ' + Math.round(r.right) + '>' + W);
    }
    return bad;
  });
}

try {
  // ① 사무실 — ERP 엑셀 → 기준 시각 → 배정표
  console.log('① 사무실: ERP 엑셀 · 기준 시각 · 창고 배정표');
  const office = await open(1280);
  await office.locator('input[type=file]').first().setInputFiles(path.join(root, 'samples/ERP_재고현황_다운로드_샘플_2026-09.xlsx'));
  await office.waitForTimeout(500);
  const mapText = await office.locator('.map-grid').innerText();
  ok(/저장위치/.test(mapText) && /Location/.test(mapText) && /저장위치명/.test(mapText), '열 연결 화면에 견본 열(저장위치 · 저장위치명 · Location)이 연결됨');
  await click(office, 'button', '작업지시 만들기');
  ok(await office.locator('.field.invalid').count() >= 1, '기준 일자·시각 없이는 만들어지지 않음');
  await office.locator('input[type=datetime-local]').fill('2026-09-30T18:00');
  await click(office, 'button', '작업지시 만들기');
  let d = await db(office);
  ok(d.items.length === 248 && d.session.cols.length === 11, 'ERP 248품목 · 원래 열 11개를 기억 — ' + d.items.length + ' · ' + d.session.cols.length);
  await office.locator('.card', { hasText: '창고 배정표 엑셀 올리기' }).locator('input[type=file]').setInputFiles(path.join(root, 'samples/창고배정표_샘플_2026-09.xlsx'));
  await office.waitForTimeout(600);
  d = await db(office);
  ok(d.items.length === 240 && d.excludedItems.length === 8 && d.session.excluded.length === 1, '배정표 적용 → 시험실창고 8품목 제외, 240품목');
  ok(d.assign['S220 #1'] === '실사자B' && d.assign['S220 #2'] === '실사자C', '부품창고는 실사자 2명이 반씩(S220 #1 · #2)');
  const assignTxt = await office.locator('table.items').innerText();
  ok(/S220 부품창고 #2/.test(assignTxt) && /관리자1/.test(assignTxt) && /사외/.test(assignTxt), '배정 표에 창고명 · 나눈 몫 · 담당자 · 창고위치가 보임');
  const orderFile = await dl(office, () => office.locator('button', { hasText: '기록서 엑셀 (SharePoint 취합용)' }).click());
  const ob = book(orderFile);
  ok(JSON.stringify(ob.wb.SheetNames) === JSON.stringify(['지시정보', '재고실사기록서', '실사자A', '실사자B', '실사자C', '실사자D', '실사자E']), '기록서 엑셀 시트 — ' + ob.wb.SheetNames.join(','));
  const head = ob.sheets[1].aoa[0];
  ok(head.slice(1, 10).join('|') === '자재코드|자재내역|저장위치|저장위치명|창고위치|단가|금액|Location|수량' && head.includes('차이수량') && head.includes('실사자'), '열 = ERP 원래 열 + 담당자 · 실사자 · 실물수량 · 차이수량 …');
  const dc = XLSX.utils.encode_col(head.indexOf('차이수량'));
  ok(/^IF\(/.test((ob.wb.Sheets['재고실사기록서'][dc + '2'] || {}).f || ''), '차이수량 칸은 엑셀 수식');
  ok(ob.sheets[0].aoa.some(r => r[0] === '이번 실사 제외 창고'), '지시정보에 제외 창고');
  // 10-28 ①: 수량·금액 칸은 숫자 그대로 + 천 단위 쉼표 서식
  const nf = XLSX.read(fs.readFileSync(orderFile), { type: 'buffer', cellNF: true }).Sheets['재고실사기록서'];
  const amtCell = nf[XLSX.utils.encode_col(head.indexOf('금액')) + '2'] || {};
  ok(amtCell.t === 'n' && amtCell.z === '#,##0' && (amtCell.v < 1000 || /,/.test(amtCell.w)), '기록서 금액 칸 = 숫자 · #,##0 — ' + amtCell.v + ' → ' + amtCell.w);
  ok((nf[dc + '2'] || {}).z === '#,##0', '차이수량 수식 칸도 #,##0');

  // ② 한 파일에 ERP + 배정표 시트
  console.log('② 한 파일에 ERP 시트 + 배정표 시트');
  const b = Sample.build();
  const combo = path.join(tmp, '재고_배정표_한파일.xlsx');
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(b.erpAoa), '① ERP 재고엑셀');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(b.rosterAoa), '② 창고 목록과 담당 인원');
  fs.writeFileSync(combo, XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' }));
  const p2 = await open(1280);
  await p2.locator('input[type=file]').first().setInputFiles(combo);
  await p2.waitForTimeout(500);
  ok(await p2.locator('.alert', { hasText: '창고 배정표로 읽었습니다' }).count() === 1, '배정표 시트를 알아봄');
  await p2.locator('input[type=datetime-local]').fill('2026-09-30T18:00');
  await click(p2, 'button', '작업지시 만들기');
  d = await db(p2);
  ok(d.items.length === 240 && Object.keys(d.assign).length === 6 && d.session.rosterFile.includes('②'), '만들 때 배정표까지 적용 — ' + d.items.length + '품목 · 배정 ' + Object.keys(d.assign).length);
  await p2.context().close();

  // ③ 휴대폰
  console.log('③ 휴대폰(360px)');
  const phone = await open(360);
  await phone.locator('input[type=file]').first().setInputFiles(orderFile);
  await phone.waitForTimeout(500);
  ok(phone.url().endsWith('#/field'), '기록서 엑셀을 넣으면 현장 입력으로');
  await phone.locator('input[list=meList]').fill('실사자E'); await phone.locator('input[list=meList]').dispatchEvent('change'); await phone.waitForTimeout(300);
  const cards = phone.locator('.item');
  ok(await cards.count() === 40, '실사자E 배정 40품목 — ' + await cards.count());
  await cards.first().locator('.item-head').click(); await phone.waitForTimeout(300);
  const sys = Number((await phone.locator('.item.open .ih-sub').innerText()).match(/수량 ([\d,]+)/)[1].replace(/,/g, ''));
  await phone.locator('.item.open input.qty').fill(String(sys - 1));
  await click(phone, '.item.open button[type=submit]');
  ok(await phone.locator('.field.invalid .err').count() >= 1, '차이가 있는데 사유 없이 저장하면 막힘');
  ok(await phone.locator('.item.open input.qty').inputValue() === String(sys - 1), '막힌 뒤에도 적던 수량이 남음');
  await phone.locator('.item.open .reasons label', { hasText: '입출고 누락' }).click();
  await click(phone, '.item.open button[type=submit]');
  d = await db(phone);
  ok(d.records.length === 1 && d.records[0].reason === '입출고 누락' && d.records[0].by === '실사자E', '사유 고르면 저장(작성자 · 시각 자동)');
  ok(await phone.locator('.item.open').count() === 1, '다음 미확인 품목이 열림');
  // 10-28 ④: 찾는 말과 보기(확인함)를 함께
  const doneIt = d.items.find(x => x.no === d.records[0].no), todoIt = d.items.find(x => x.no !== d.records[0].no && d.assign && x.wh === doneIt.wh);
  await phone.locator('.seg[aria-label="보기"] label', { hasText: '확인함' }).click(); await phone.waitForTimeout(300);
  await phone.locator('input[type=search]').fill(doneIt.code); await phone.waitForTimeout(300);
  ok(await phone.locator('.item').count() === 1, '확인함 + 「' + doneIt.code + '」 → 1품목 — ' + await phone.locator('.item').count());
  await phone.locator('input[type=search]').fill(todoIt.code); await phone.waitForTimeout(300);
  ok(await phone.locator('.item').count() === 0 && /「전체」 보기에서는/.test(await phone.locator('.item-list').innerText()), '확인함 + 미확인 품목 코드 → 0품목 · 「전체」 안내');
  await phone.locator('input[type=search]').fill(''); await phone.waitForTimeout(200);
  const phoneXlsx = await dl(phone, () => phone.locator('button', { hasText: '결과 엑셀 받기' }).click());
  const pb = book(phoneXlsx);
  ok(pb.wb.SheetNames.join(',') === '지시정보,재고실사기록서' && pb.sheets[1].aoa.length === 41, '결과 엑셀 = 기록서 양식, 내 줄 40개');

  // ④ 사무실에서 합치기
  console.log('④ 사무실: 결과 엑셀 합치기');
  await go(office, '#/merge');
  await office.locator('input[type=file][multiple]').setInputFiles(phoneXlsx); await office.waitForTimeout(600);
  d = await db(office);
  ok(d.records.length === 1, '휴대폰 결과 엑셀 1건 합침');
  await office.locator('input[type=file][multiple]').setInputFiles(phoneXlsx); await office.waitForTimeout(600);
  d = await db(office);
  ok(d.records.length === 1 && d.files[d.files.length - 1].records === 0, '같은 파일을 또 넣으면 0건');

  // ⑤ 가상 샘플 흐름
  console.log('⑤ 가상 샘플 · SharePoint 합치기 · 설정 · 보고서');
  await click(office, 'a', '설정');
  await go(office, '#/order');
  await click(office, 'button', '가상 샘플로 시작');
  d = await db(office);
  ok(d.sample && d.items.length === 240 && d.session.excluded.length === 1, '가상 샘플 240품목 · 제외 창고 1');
  await go(office, '#/merge');
  await click(office, 'button', 'SharePoint 취합 파일 합치기 체험');
  d = await db(office);
  ok(d.files.some(f => /SharePoint/.test(f.name) && f.records === 42), 'SharePoint 취합 파일 42건 합침');
  const confBefore = await office.locator('.conf.open').count();
  await office.locator('.conf.open input[type=radio]').first().click(); await office.waitForTimeout(300);
  ok(confBefore === 2 && await office.locator('.conf.open').count() === 1, '충돌 2건 → 하나 고르면 1건');
  await go(office, '#/settings');
  ok(await office.locator('select').first().inputValue() === 'zero', '허용 기준 기본 = 허용 0');
  ok(await office.locator('label.field', { hasText: '차이율 (%)' }).locator('input').isDisabled(), '허용 0 일 때 차이율 칸은 잠김');
  ok((await office.locator('textarea[aria-label="사유 목록"]').inputValue()).split('\n').length === 6, '사유 목록 기본 6개');
  await go(office, '#/variance');
  ok(/허용 0/.test(await office.locator('.filters').innerText()), '차이대사에 허용 0 기준 표시');
  await go(office, '#/bi');
  const bt = await office.locator('.view-toggle').innerText();
  ok(/창고위치별/.test(bt) && /자재유형별/.test(bt), 'BI 기준에 창고위치(사내/사외)별 — ' + bt.replace(/\s+/g, ' '));
  await go(office, '#/report');
  await click(office, 'button', 'v1 발행');
  const rep = book(await dl(office, () => office.locator('button', { hasText: '결과서 엑셀' }).click()));
  ok(rep.wb.SheetNames.includes('창고위치별') && rep.wb.SheetNames.includes('실사자별'), '결과서 엑셀 — ' + rep.wb.SheetNames.join(','));
  ok(rep.sheets[0].aoa.some(r => r[0] === '이번 실사 제외 창고'), '결과서에 제외 창고');
  // 3차 반영 — 보고서 PPT: 기본 양식으로 만들기 → 직전 보고서(가상 양식) 올려서 만들기
  const JSZip = require(path.join(root, 'vendor/jszip.min.js'));
  const P = require(path.join(root, 'js/pptx-report.js'));
  const pptCard = office.locator('section.card', { hasText: '보고서 PPT' });
  ok(/기본 양식/.test(await pptCard.innerText()), 'PPT 카드 — 기본 양식 안내');
  const d0 = await P.dump(JSZip, fs.readFileSync(await dl(office, () => pptCard.locator('button', { hasText: 'PPT 만들기' }).click())));
  ok(d0.length >= 3 && d0.some(s => s.tables.some(t => t.kind === 'diff')), '기본 양식 PPT — 슬라이드 ' + d0.length + '장');
  await pptCard.locator('input[type=file]').setInputFiles(path.join(root, 'samples/재고실사_결과보고_직전양식_샘플.pptx'));
  await office.waitForTimeout(600);
  const ct = await pptCard.innerText();
  ok(/차이 슬라이드 2장/.test(ct) && /원자재/.test(ct) && /메모 상자 4개/.test(ct), '직전 보고서 구조 읽음(차이 슬라이드 2장 · 현황 행 · 메모 4개)');
  await pptCard.locator('label.field', { hasText: '협력사재고' }).first().locator('input').fill('협력사 확인서 2곳');
  await pptCard.locator('label.field', { hasText: '협력사재고' }).first().locator('input').dispatchEvent('change');
  const pf = await dl(office, () => pptCard.locator('button', { hasText: 'PPT 만들기' }).click());
  const d1 = await P.dump(JSZip, fs.readFileSync(pf));
  const ov = d1[0].tables.find(t => t.kind === 'overview').rows;
  ok(ov[2][1] === '26.09.30  18:00' && ov[5][1] === '협력사 확인서 2곳', '1장 개요 — 기준 시점 자동 · 고친 칸 반영');
  ok(!d1.some(s => s.shapes.some(t => /계산로직/.test(t))), '메모 상자 지움');
  const dr = d1.flatMap(s => s.tables.filter(t => t.kind === 'diff').flatMap(t => t.rows.slice(1).filter(r => r[0] !== '합계')));
  ok(dr.length > 20 && dr.every(r => Number(r[5].replace(/,/g, '')) === Number(r[4].replace(/,/g, '')) - Number(r[3].replace(/,/g, ''))), '2장 차이수량 = 실물 − 전산 (' + dr.length + '품목)');
  ok(/만든 PPT: 슬라이드/.test(await pptCard.innerText()), '만든 결과 안내');

  // ⑦ 10-28 ③: 같은 자재 · 같은 Location · 다른 저장 위치(가짜 값) — ERP 줄마다 한 줄
  console.log('⑦ 같은 자재가 여러 줄인 ERP(저장 위치 내역 · Location) — 올린 그대로');
  const RAW = [
    ['자재', '자재내역', '저장 위치', '저장 위치 내역', '단가', '금액', '내역', 'Location', '기말재고', '최종 입고일'],
    ['Z900-1', '가상 판재', 'R1', '가상원재료창고', 10, 1750, '원자재', 'L-1-01', 175, 45930],
    ['Z900-1', '가상 판재', 'R2', '가상라인창고', 10, 200, '원자재', 'L-1-01', 20, 45931],
    ['Z900-1', '가상 판재', 'R1', '가상원재료창고', 10, 3000, '원자재', 'L-2-05', 300, null],
    ['Z900-1', '가상 판재', 'R1', '가상원재료창고', 10, 710, '원자재', 'L-3-02', 71, null],
    ['Z777-2', '가상 봉재', 'R2', '가상라인창고', 1000, 12000000, '원자재', 'L-1-01', 12000, null]
  ];
  const rawFile = path.join(tmp, 'raw.xlsx'), rwb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(rwb, XLSX.utils.aoa_to_sheet(RAW), 'RAW');
  fs.writeFileSync(rawFile, XLSX.write(rwb, { bookType: 'xlsx', type: 'buffer' }));
  const p7 = await open(1280);
  await p7.locator('input[type=file]').first().setInputFiles(rawFile); await p7.waitForTimeout(500);
  await p7.locator('input[type=datetime-local]').fill('2026-10-28T09:00');
  await click(p7, 'button', '작업지시 만들기');
  d = await db(p7);
  ok(d.items.length === 5 && d.items.map(x => x.sys).join(',') === '175,20,300,71,12000', 'ERP 5줄 → 5품목, 수량 그대로 — ' + d.items.map(x => x.no + '=' + x.sys).join(' '));
  const rbFile = await dl(p7, () => p7.locator('button', { hasText: '기록서 엑셀 (SharePoint 취합용)' }).click()), rb = book(rbFile);
  const rm = rb.sheets.find(x => x.name === '재고실사기록서').aoa, ki = rm[0].indexOf('기말재고');
  ok(rm.length === 6 && rm.slice(1).map(r => r[ki]).sort((a, b) => a - b).join(',') === '20,71,175,300,12000', '기록서 5줄 · 기말재고 줄마다 그대로');
  const rnf = XLSX.read(fs.readFileSync(rbFile), { type: 'buffer', cellNF: true }).Sheets['재고실사기록서'];
  const dateCol = XLSX.utils.encode_col(rm[0].indexOf('최종 입고일'));
  const dcell = Object.keys(rnf).filter(a => a.startsWith(dateCol) && a !== dateCol + '1').map(a => rnf[a]).find(c => c && c.v != null) || {};
  ok(dcell.z === 'yyyy-mm-dd' && /^\d{4}-\d\d-\d\d$/.test(dcell.w || ''), '최종 입고일 = 날짜 칸 yyyy-mm-dd — ' + dcell.w);
  const tt = await p7.locator('table.items').innerText();
  ok(/12,000,200원/.test(tt), '화면 배정 표 금액에 천 단위 쉼표(R2 12,000,200원)');
  await p7.context().close();

  // ⑧ 11-08 수강생 요청 — 분석(BI) 그래프 행 고르기 · 차이대사 사진 보기 · 실사자별 By 기준 · 빈 기기 현장 안내
  console.log('⑧ 11-08 — BI 행 고르기 · 사진 보기 · 실사자별 By · 빈 기기 안내');
  await office.evaluate(() => {
    const d = JSON.parse(localStorage.getItem('data10-14.db'));
    const px = document.createElement('canvas'); px.width = 40; px.height = 30; px.getContext('2d').fillRect(0, 0, 40, 30);
    const counted = new Set(d.records.map(r => r.no));
    const it = d.items.find(i => i.sys > 0 && !counted.has(i.no));
    d.records.push({ rid: 'chk-photo', no: it.no, qty: it.sys + 1, reason: '입출고 누락', note: '', by: '대리실사', at: '2026-10-01T09:00', dev: 'chk', src: '점검', photo: px.toDataURL('image/jpeg'), sessionId: d.session.id });
    localStorage.setItem('data10-14.db', JSON.stringify(d));
  });
  await office.reload(); await office.waitForTimeout(400);
  await go(office, '#/bi');
  const nWh = await office.locator('table.bi tbody tr').count() - 1;
  const checked = await office.locator('table.bi tbody input[type=checkbox]:checked').count();
  ok(checked === Math.min(8, nWh), 'BI 창고별 — 처음 그래프 행 = 절대 차이 상위 ' + checked + '/' + nWh);
  const names = await office.locator('figure.fig').first().evaluate(f => [...f.querySelectorAll('svg text.tick')].filter(t => t.querySelector('tspan')).map(t => [...t.querySelectorAll('tspan')].map(x => x.textContent).join('')));
  ok(names.length === checked && names.every(t => !/^[A-Z]\d{2,}/.test(t)), 'BI 그래프 이름에 저장위치 코드 없음 — ' + names.join(' / '));
  const bars0 = await office.locator('figure.fig').first().locator('path.bar').count();
  await office.locator('table.bi tbody input[type=checkbox]').first().click(); await office.waitForTimeout(300);
  const bars1 = await office.locator('figure.fig').first().locator('path.bar').count();
  ok(bars1 < bars0, '체크를 풀면 그래프 막대가 줄어듦 ' + bars0 + ' → ' + bars1);
  await click(office, 'button', '모두 해제');
  ok(await office.locator('figure.fig').first().locator('path.bar').count() === 0, '모두 해제 → 막대 0');
  await go(office, '#/merge');
  const mt = await office.locator('.card', { hasText: '실사자별 진행' }).locator('table').innerText();
  ok(/대리실사\t-\t1\t-/.test(mt), '실사자별 — 배정 없는 「대리실사」가 직접 센 품목 1로 보임');
  await go(office, '#/variance');
  await office.locator('select[aria-label="보기"]').selectOption('all');
  await office.locator('input[aria-label="찾기"]').fill(''); await office.waitForTimeout(200);
  const phb = office.locator('button.photo-btn').first();
  ok(await phb.count() === 1, '차이대사에 「사진 보기」 단추');
  await phb.click(); await office.waitForTimeout(300);
  const dimg = await office.evaluate(() => { const d = document.querySelector('dialog.photo-dialog'); const i = d && d.querySelector('img.photo-full'); return d && d.open && i && i.naturalWidth; });
  ok(dimg > 0, '사진 보기 → 큰 사진 창(' + dimg + 'px)');
  await click(office, 'dialog button', '닫기');
  ok(!(await office.$('dialog.photo-dialog')), '닫기 → 창 사라짐');
  const fresh = await open(360); await go(fresh, '#/field');
  ok(await fresh.locator('label', { hasText: '받은 기록서 엑셀 넣기' }).count() === 1 && await fresh.locator('summary', { hasText: '여러 사람이' }).count() === 1, '빈 기기 ② 현장 입력 — 「받은 기록서 엑셀 넣기」 · 여러 사람 순서');
  await fresh.context().close();

  // ⑥ 폭별 가로 넘침
  console.log('⑥ 화면 × 폭 가로 넘침');
  const pages = ['home', 'order', 'field', 'merge', 'variance', 'bi', 'report', 'settings'];
  for (const w of [360, 1280]) {
    await office.setViewportSize({ width: w, height: 900 });
    for (const pg of pages) {
      await go(office, '#/' + pg);
      if (process.env.BREAK && !(await office.$('#brk'))) await office.evaluate(() => { const x = document.createElement('div'); x.style.width = '2000px'; x.style.height = '5px'; x.id = 'brk'; document.body.appendChild(x); });
      const bad = await overflow(office);
      ok(!bad.length, w + 'px ' + pg + ' 넘침 ' + bad.length + (bad.length ? ' — ' + bad.slice(0, 2).join(' / ') : ''));
    }
  }
} finally {
  await browser.close();
  fs.rmSync(tmp, { recursive: true, force: true });
}
ok(!errors.length, 'JS 오류 ' + errors.length + (errors.length ? ' — ' + errors.slice(0, 3).join(' / ') : ''));
console.log(problems.length ? '\n실패 ' + problems.length + '개' : '\n모두 통과');
process.exitCode = problems.length ? 1 : 0;
