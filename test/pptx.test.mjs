// 실행: node test/pptx.test.mjs   (의존성 없음 — vendor 의 JSZip 만 씀)
// 3차 반영(2026-10-27): 「직전 보고서(pptx)를 올리면 내용만 바꿔 새 보고서」 — js/pptx-report.js 검사.
// 양식은 리포의 가상 양식(samples/재고실사_결과보고_직전양식_샘플.pptx, 회사명 · 숫자 가상)과 도구의 기본 양식만 씁니다.
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
const require = createRequire(import.meta.url);
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const L = require('../js/logic.js');
const P = require('../js/pptx-report.js');
const Sample = require('../js/sample-data.js');
const JSZip = require('../vendor/jszip.min.js');
const XLSX = require('../vendor/xlsx.full.min.js');

let passed = 0;
async function test(name, fn) {
  try { await fn(); passed++; console.log('  ok  ' + name); }
  catch (e) { console.error('  FAIL ' + name + '\n       ' + (e && e.message)); process.exitCode = 1; }
}

/* 가상 샘플 + SharePoint 취합 파일(실사자E · C)까지 합친 상태 */
const S = Sample.build();
const db = L.restoreDb({ session: S.session, items: S.items, assign: S.assign, records: S.records });
{
  const sheets = Sample.fillPaper(L.orderSheets(db.session, db.items, db.assign), S.heldOut, S.mainRows);
  const wb = XLSX.utils.book_new();
  sheets.forEach(s => XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(s.aoa), s.name));
  const back = XLSX.read(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }), { type: 'buffer' });
  const w = L.parseOrderWorkbook(back.SheetNames.map(n => ({ name: n, aoa: XLSX.utils.sheet_to_json(back.Sheets[n], { header: 1, raw: true, defval: null }) })), { fileName: S.eFile });
  db.records = L.mergeRecords(db.records, w.records).list;
}
const an = L.analyze(db);
const TPL = fs.readFileSync(path.join(root, 'samples/재고실사_결과보고_직전양식_샘플.pptx'));
const diffRows = an.rows.filter(r => r.counted && r.diff !== 0);

/* ── 꾸러미 검사(파워포인트가 「복구」를 묻는 원인들) ── */
function wellFormed(xml, name) {
  const st = [], re = /<(\/?)([A-Za-z_][\w:.-]*)((?:\s+[\w:.-]+\s*=\s*"[^"]*")*)\s*(\/?)>|<\?[^>]*\?>|<!--[\s\S]*?-->/g;
  let m, pos = 0;
  while ((m = re.exec(xml))) {
    const gap = xml.slice(pos, m.index);
    if (/[<>]/.test(gap)) throw new Error(name + ': 태그 밖 < > (' + gap.slice(0, 40) + ')');
    if (/&(?!(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/.test(gap)) throw new Error(name + ': 잘못된 & (' + gap.slice(0, 40) + ')');
    pos = re.lastIndex;
    if (!m[2]) continue;
    if (m[1]) { const top = st.pop(); if (top !== m[2]) throw new Error(name + ': </' + m[2] + '> 짝 없음(열린 것 ' + top + ')'); }
    else if (!m[4]) st.push(m[2]);
  }
  if (/[<>]/.test(xml.slice(pos).trim())) throw new Error(name + ': 끝에 남은 < >');
  if (st.length) throw new Error(name + ': 닫히지 않음 ' + st.slice(-3).join(','));
}
function resolve(base, t) {
  if (t.startsWith('/')) return t.slice(1);
  const parts = base.split('/'); parts.pop();
  t.split('/').forEach(p => { if (p === '..') parts.pop(); else if (p !== '.') parts.push(p); });
  return parts.join('/');
}
async function checkPkg(buf) {
  const z = await JSZip.loadAsync(buf), names = Object.keys(z.files).filter(n => !z.files[n].dir);
  const text = {};
  for (const n of names) if (/\.(xml|rels)$/.test(n)) { text[n] = await z.file(n).async('string'); wellFormed(text[n], n); }
  for (const n of names.filter(n => n.endsWith('.rels'))) {
    const owner = n.replace(/_rels\//, '').replace(/\.rels$/, '');
    for (const m of text[n].matchAll(/<Relationship\s[^>]*>/g)) {
      if (/TargetMode="External"/.test(m[0])) continue;
      const t = /Target="([^"]+)"/.exec(m[0])[1], p = resolve(owner, t);
      assert.ok(names.includes(p), n + ' → 없는 파트 ' + p);
    }
  }
  const ct = text['[Content_Types].xml'];
  for (const m of ct.matchAll(/PartName="\/([^"]+)"/g)) assert.ok(names.includes(m[1]), '[Content_Types] 에 남은 지운 파트 ' + m[1]);
  names.filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n)).forEach(n => assert.ok(ct.includes('PartName="/' + n + '"'), '[Content_Types] 에 없는 슬라이드 ' + n));
  const pres = text['ppt/presentation.xml'], prels = text['ppt/_rels/presentation.xml.rels'];
  const ids = [...pres.matchAll(/<p:sldId id="(\d+)" r:id="(\w+)"\/>/g)];
  assert.equal(new Set(ids.map(x => x[1])).size, ids.length, '슬라이드 id 겹침');
  ids.forEach(x => { assert.ok(+x[1] >= 256 && +x[1] < 2147483648, 'id 범위 ' + x[1]); assert.ok(new RegExp('Id="' + x[2] + '"[^>]*/slide"').test(prels) || new RegExp('Id="' + x[2] + '"').test(prels), 'rels 에 없는 ' + x[2]); });
  const slideParts = names.filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n));
  assert.equal(slideParts.length, ids.length, '슬라이드 파트 수 = 목록 수(남은 파트 없음)');
  const notesOwner = {};
  names.filter(n => /^ppt\/slides\/_rels\//.test(n)).forEach(n => { for (const m of text[n].matchAll(/Target="\.\.\/notesSlides\/([^"]+)"/g)) { assert.ok(!notesOwner[m[1]], '메모 슬라이드 공유 ' + m[1]); notesOwner[m[1]] = n; } });
  // 11-08: PowerPoint 가 「복구」하며 슬라이드를 비우는 고유값 겹침 — 표 안 줄 번호(a16:rowId), 슬라이드 번호(p14:creationId), 개체 번호(a16:creationId)
  const slideXml = slideParts.map(n => text[n]);
  slideXml.forEach((x, i) => { for (const t of x.matchAll(/<a:tbl>[\s\S]*?<\/a:tbl>/g)) { const ids = [...t[0].matchAll(/<a16:rowId\b[^>]*\sval="(\d+)"/g)].map(m => m[1]); assert.equal(new Set(ids).size, ids.length, slideParts[i] + ' 표 안 a16:rowId 겹침'); } });
  const cids = slideXml.flatMap(x => [...x.matchAll(/<p14:creationId\b[^>]*\sval="(\d+)"/g)].map(m => m[1]));
  assert.equal(new Set(cids).size, cids.length, '슬라이드 p14:creationId 겹침');
  const gids = slideXml.flatMap(x => [...x.matchAll(/<a16:creationId\b[^>]*\sid="(\{[^"]+\})"/g)].map(m => m[1]));
  assert.equal(new Set(gids).size, gids.length, '개체 a16:creationId 겹침');
  const app = text['docProps/app.xml'];
  if (app && /<Slides>/.test(app)) assert.equal(+/<Slides>(\d+)<\/Slides>/.exec(app)[1], ids.length, 'app.xml 슬라이드 수');
  return { z, names, text, ids };
}
const num = s => s === '-' || s === '' ? null : Number(String(s).replace(/,/g, ''));
const dump = buf => P.dump(JSZip, buf);
const make = (tpl, opts) => P.make(JSZip, tpl, an, db, Object.assign({ type: 'nodebuffer' }, opts || {}));

console.log('양식 읽기');
await test('기본 양식 — 1장 개요 · 실사 현황, 2장 차이(9열), 3장 이슈를 찾음', async () => {
  const i = await P.inspect(JSZip, null);
  assert.equal(i.slides, 3);
  assert.deepEqual(i.overview.rows.map(x => x.auto), ['countAt', 'baseAt', '', '', '']);
  assert.ok(i.status.ok && i.status.cols === 9);
  assert.deepEqual(i.diff.fields, ['code', 'name', 'whCode', 'sys', 'cnt', 'diff', 'price', 'diffAmt', 'reasonNote']);
  assert.equal(i.issue.rows, 2); assert.equal(i.memos, 0);
});
await test('직전 양식 샘플 — 차이 슬라이드 2장 · 메모 상자 4개 · 현황 행 이름 7개 · 단위 「천 EA, 백만원」', async () => {
  const i = await P.inspect(JSZip, TPL);
  assert.deepEqual(i.diffSlides, [1, 2]); assert.equal(i.memos, 4);
  assert.deepEqual(i.status.labels, ['원자재', '부자재', '제품', '반제품', '상품', '저장품', '포장재']);
  assert.match(i.status.unit, /천 EA/);
});
await test('열 이름 · 단위 읽기 — 「자재」 만 있어도 자재코드, 「차이 사유 및 조치사항」 은 사유+특이사항', () => {
  assert.equal(P.diffFieldOf('자재'), 'code'); assert.equal(P.diffFieldOf('자재 코드'), 'code'); assert.equal(P.diffFieldOf('자재내역'), 'name');
  assert.equal(P.diffFieldOf('저장위치'), 'whCode'); assert.equal(P.diffFieldOf('저장위치명'), 'whName'); assert.equal(P.diffFieldOf('차이 사유 및 조치사항'), 'reasonNote');
  assert.equal(P.diffFieldOf('차이 사유'), 'reason'); assert.equal(P.diffFieldOf('조치사항'), 'note'); assert.equal(P.diffFieldOf('전산수량'), 'sys'); assert.equal(P.diffFieldOf('차이금액'), 'diffAmt');
  assert.deepEqual([P.parseUnit('(단위 : 천 EA, 백만원)').q, P.parseUnit('(단위 : 천 EA, 백만원)').a], [1000, 1e6]);
  assert.deepEqual([P.parseUnit('(단위 : EA, 원)').q, P.parseUnit('(단위 : EA, 원)').a], [1, 1]);
  assert.equal(P.fmtScaled(1234567, 1e6), '1.2'); assert.equal(P.fmtScaled(-1500, 1), '-1,500');
});
await test('글자색 — 빨강을 칠하고, 걷을 때는 도구가 칠한 빨강만 걷음(양식의 흰 글자는 그대로)', () => {
  const white = '<a:rPr lang="ko-KR" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="X"/></a:rPr>';
  const red = P.setColor(white, P.RED);
  assert.match(red, /val="FF0000"/); assert.doesNotMatch(red, /FFFFFF/); assert.match(red, /<a:latin/);
  assert.doesNotMatch(P.setColor(red, null), /FF0000/);
  assert.match(P.setColor(white, null), /FFFFFF/);
  assert.match(P.setColor('<a:rPr lang="ko-KR"/>', P.RED), /^<a:rPr lang="ko-KR"><a:solidFill><a:srgbClr val="FF0000"\/><\/a:solidFill><\/a:rPr>$/);
  const withLn = '<a:rPr><a:ln w="1"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:rPr>';
  assert.match(P.setColor(withLn, P.RED), /<\/a:ln><a:solidFill><a:srgbClr val="FF0000"\/>/, '선 색은 두고 글자색만');
});

console.log('만들기 — 직전 양식 샘플');
const R = await make(TPL);
const D = await dump(R.out);
await test('꾸러미가 깨지지 않음 — XML 짝 · rels 대상 · [Content_Types] · 슬라이드 id · 메모 슬라이드 · app.xml', async () => {
  const k = await checkPkg(R.out);
  assert.equal(k.ids.length, 1 + R.log.diffPages + 1);
});
await test('작성 메모 상자(「(계산로직)」 · 「AI가 초안」)를 지움 — 끄면 남음', async () => {
  assert.ok(R.log.removedMemos >= 3);
  assert.ok(!D.some(s => s.shapes.some(t => /계산로직|AI가 초안/.test(t))));
  const keep = await dump((await make(TPL, { removeMemos: false })).out);
  assert.ok(keep.some(s => s.shapes.some(t => /계산로직/.test(t))));
});
await test('1장 개요 — 집계 기준 시점 = ERP 기준 일자·시각, 실사 시행일 = 첫 ~ 마지막 확인 시각, 실사 방법 칸은 그대로', () => {
  const ov = D[0].tables.find(t => t.kind === 'overview').rows;
  assert.equal(ov[2][1], '26.09.30  18:00');
  assert.match(ov[1][1], /^26\.10\.01 {2}\d\d:\d\d ~ /);
  assert.equal(ov[4][0], '가나산업 사내 및 사외창고'); assert.equal(ov[5][1], '(직접 입력함)');
});
await test('1장 실사 현황 — 행 이름이 자재타입이라 자재타입으로, 없는 「상품」 은 「-」, 합계 = 장부 합(백만원) · 실사완료 = 실물 수량(학생 답 10-27)', () => {
  assert.equal(R.log.statusDim, 'mkind');
  const t = D[0].tables.find(x => x.kind === 'status').rows;
  const body = t.slice(2, -1), tot = t[t.length - 1];
  assert.deepEqual(body.map(r => r[0]), ['원자재', '부자재', '제품', '반제품', '상품', '저장품', '포장재']);
  assert.deepEqual(body.find(r => r[0] === '상품').slice(1), Array(8).fill('-'));
  const sysAmt = an.rows.filter(r => r.mkind && !/^\(/.test(r.mkind)).reduce((s, r) => s + r.sysAmt, 0);
  assert.equal(tot[0], '합계'); assert.equal(num(tot[2]), Math.round(sysAmt / 1e6 * 10) / 10);
  const raw = an.rows.filter(r => r.mkind === '원자재');
  const done = raw.filter(r => r.counted).reduce((s, r) => s + r.cnt, 0), all = raw.reduce((s, r) => s + r.sys, 0);
  const rest = raw.filter(r => !r.counted).reduce((s, r) => s + r.sys, 0);
  assert.equal(body[0][4], (Math.round(done / all * 1000) / 10).toFixed(1), '실사완료 % = 실물 수량 ÷ 전산 전체');
  assert.equal(num(body[0][7]), Math.round(rest / 1000 * 10) / 10, '미실시 = 아직 안 센 품목의 전산 수량');
  assert.ok(raw.some(r => r.counted && r.cnt !== r.sys), '샘플에 차이 품목이 있어야 실물/전산이 갈린다');
});
await test('1장 실사 현황 — 창고로 고르면 실사인원 0 창고(시험실창고)는 모든 칸 「-」 이고 합계에서 빠짐', async () => {
  const d = await dump((await make(TPL, { statusBy: 'wh' })).out);
  const t = d[0].tables.find(x => x.kind === 'status').rows, body = t.slice(2, -1);
  const ex = body.find(r => r[0] === '시험실창고');
  assert.ok(ex, '제외 창고 행'); assert.deepEqual(ex.slice(1), Array(8).fill('-'));
  const target = an.rows.reduce((s, r) => s + r.sysAmt, 0);
  assert.equal(num(t[t.length - 1][2]), Math.round(target / 1e6 * 10) / 10);
  assert.ok(!body.some(r => ['원자재', '상품'].includes(r[0])), '행 이름이 창고가 아니면 직전 행 이름은 버림');
});
const diffSlides = D.filter(s => s.tables.some(t => t.kind === 'diff'));
await test('2장 — 창고마다 한 장, 넘치면 (1/n) 이어지는 장, 모든 차이 품목이 한 번씩', () => {
  assert.equal(diffSlides.length, R.log.diffPages);
  const groups = new Set(diffRows.map(r => r.whCode));
  assert.equal(R.log.groups, groups.size);
  let n = 0;
  diffSlides.forEach(s => {
    const t = s.tables.find(x => x.kind === 'diff').rows;
    const body = t.slice(1).filter(r => r[0] !== '합계');
    assert.ok(body.length <= R.log.cap, '한 장 ' + body.length + ' > ' + R.log.cap);
    n += body.length;
    assert.ok(s.shapes.some(x => /^차이 내용 : \S/.test(x)), '제목');
  });
  assert.equal(n, diffRows.length);
  const multi = diffSlides.filter(s => s.shapes.some(x => /\(\d+\/\d+\)$/.test(x)));
  assert.ok(multi.length >= 2, '이어지는 장');
});
await test('2장 — 차이수량 = 실물 − 전산(모자라면 −), 차이금액 = 차이수량 × 단가, 0 이 아니면 빨간 글자', () => {
  let short = 0;
  diffSlides.forEach(s => {
    const T = s.tables.find(x => x.kind === 'diff');
    T.rows.forEach((r, i) => {
      if (i === 0 || r[0] === '합계') return;
      const sys = num(r[3]), cnt = num(r[4]), d = num(r[5]), price = num(r[6]), amt = num(r[7]);
      assert.equal(d, cnt - sys, r[0] + ' 차이수량');
      assert.ok(Math.abs(amt - Math.round(d * price)) <= 1, r[0] + ' 차이금액');
      assert.ok(T.red[i][5], r[0] + ' 차이수량 빨간 글자'); assert.equal(T.red[i][7], amt !== 0, r[0] + ' 차이금액은 0 이 아닐 때만 빨강(목록에 없던 자재는 단가 0)');
      assert.ok(!T.red[i][3] && !T.red[i][4], '전산 · 실물은 빨갛지 않음');
      if (d < 0) short++;
    });
  });
  assert.ok(short > 10, '모자란 품목은 − 로');
});
await test('2장 — 합계 줄은 창고의 마지막 장에만, 차이수량 · 차이금액만 더함', () => {
  const byWh = {};
  diffRows.forEach(r => { const k = r.whCode; byWh[k] = byWh[k] || { q: 0, a: 0 }; byWh[k].q += r.diff; byWh[k].a += r.diffAmt; });
  let totals = 0;
  diffSlides.forEach(s => {
    const t = s.tables.find(x => x.kind === 'diff').rows, last = t[t.length - 1];
    const title = s.shapes.find(x => /^차이 내용/.test(x));
    const pg = /\((\d+)\/(\d+)\)$/.exec(title);
    if (last[0] !== '합계') { assert.ok(pg && pg[1] !== pg[2], '마지막 장이 아닌데 합계가 없음은 정상 — 마지막 장이면 실패'); return; }
    totals++;
    assert.ok(!pg || pg[1] === pg[2], '합계는 마지막 장에만');
    const code = t[1][2], g = byWh[code];
    assert.equal(num(last[5]), L.round(g.q, 6)); assert.ok(Math.abs(num(last[7]) - Math.round(g.a)) <= 1);
    assert.equal(last[3], ''); assert.equal(last[6], '', '전산 · 단가는 비움');
  });
  assert.equal(totals, Object.keys(byWh).length);
});
await test('3장 — 사유별 초안(주제 · 이슈 · 개선)이 칸에 들어가고 「…」 자리표시가 남지 않음, 넘치는 줄은 마지막 칸에 「주제: 」', () => {
  const s = D[D.length - 1];
  const top = an.reasons.filter(x => x.items).sort((a, b) => b.absAmt - a.absAmt)[0];
  assert.ok(s.shapes.includes(top.key), '첫 주제 = 절대 차이 금액 1위 사유');
  assert.ok(!s.shapes.some(t => t === '…'));
  assert.ok(s.shapes.some(t => /^\d+품목 · 순차이/.test(t)));
  const draft = P.issueDraft(an, 2);
  assert.ok(draft.length > 2);
  assert.ok(s.shapes.some(t => t.includes(draft[2].label + ': ')), '세 번째 줄은 마지막 칸에 주제를 붙여');
});
await test('3장 — 화면에서 고친 「주제 | 이슈 | 개선」 을 그대로, 「 ; 」 는 줄바꿈', async () => {
  const issues = P.parseIssueText('재고 정확도 | 입출고 마감이 늦음 ; 야간 출하 | 마감 시각 합의\n위치 | B 구역 분산 | 위치 재정리');
  assert.deepEqual(issues[0], { label: '재고 정확도', issue: '입출고 마감이 늦음\n야간 출하', act: '마감 시각 합의' });
  const s = (await dump((await make(TPL, { issues })).out)).pop();
  assert.ok(s.shapes.includes('입출고 마감이 늦음\n야간 출하')); assert.ok(s.shapes.includes('위치 재정리'));
  assert.equal(P.issueTextOf(issues).split('\n')[0], '재고 정확도 | 입출고 마감이 늦음 ; 야간 출하 | 마감 시각 합의');
});
await test('개요 칸을 화면에서 고치면 그 값으로(협력사재고 「직접 입력」)', async () => {
  const d = await dump((await make(TPL, { overviewEdits: { '5,1': '협력사 3곳 확인서 수령', '1,1': '26.10.01  08:30' } })).out);
  const ov = d[0].tables.find(t => t.kind === 'overview').rows;
  assert.equal(ov[5][1], '협력사 3곳 확인서 수령'); assert.equal(ov[1][1], '26.10.01  08:30');
});
await test('한 장 줄 수 5 → 쪽 수 = 창고마다 올림(품목 ÷ 5), 창고위치(사내/사외)로 나누기', async () => {
  const r5 = await make(TPL, { rowsPerPage: 5 });
  const by = {}; diffRows.forEach(r => { by[r.whCode] = (by[r.whCode] || 0) + 1; });
  assert.equal(r5.log.diffPages, Object.values(by).reduce((s, n) => s + Math.ceil(n / 5), 0));
  const rs = await make(TPL, { diffBy: 'site' });
  const titles = (await dump(rs.out)).flatMap(s => s.shapes.filter(t => /^차이 내용/.test(t)));
  assert.ok(titles.every(t => /사내|사외/.test(t))); assert.equal(rs.log.groups, 2);
  await checkPkg(rs.out);
});
await test('긴 사유 · 자재내역은 칸 너비에 맞춰 줄임(…) — 줄이 늘어나 표가 슬라이드 밖으로 나가지 않게', () => {
  const t = P.fitText('가'.repeat(200), 2900000, 383672, 1100, 91440);
  assert.ok(t.endsWith('…') && t.length < 60);
  assert.equal(P.fitText('파손·폐기 미반영', 2900000, 383672, 1100, 91440), '파손·폐기 미반영');
});
await test('다시 넣기 — 만든 보고서를 다음 달 「직전 보고서」로 올려도 같은 모양(차이 슬라이드를 모두 지우고 다시 만듦)', async () => {
  const R2 = await make(R.out);
  await checkPkg(R2.out);
  const D2 = await dump(R2.out);
  assert.equal(D2.length, D.length);
  assert.deepEqual(D2.map(s => s.tables.map(t => t.rows)), D.map(s => s.tables.map(t => t.rows)));
  assert.equal(R2.log.cap, R.log.cap);
});

await test('메모 슬라이드(발표자 노트)가 달린 양식 — 지운 차이 슬라이드의 노트도 지우고, 복제한 슬라이드는 노트를 함께 쓰지 않음', async () => {
  const z = await JSZip.loadAsync(TPL);
  const RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
  for (const k of [2, 3]) {
    z.file('ppt/notesSlides/notesSlide' + k + '.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:notes>');
    z.file('ppt/notesSlides/_rels/notesSlide' + k + '.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + RT + 'slide" Target="../slides/slide' + k + '.xml"/></Relationships>');
    const rp = 'ppt/slides/_rels/slide' + k + '.xml.rels';
    z.file(rp, (await z.file(rp).async('string')).replace('</Relationships>', '<Relationship Id="rId9" Type="' + RT + 'notesSlide" Target="../notesSlides/notesSlide' + k + '.xml"/></Relationships>'));
    z.file('[Content_Types].xml', (await z.file('[Content_Types].xml').async('string')).replace('</Types>', '<Override PartName="/ppt/notesSlides/notesSlide' + k + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/></Types>'));
  }
  const withNotes = await z.generateAsync({ type: 'nodebuffer' });
  const r = await make(withNotes);
  const k = await checkPkg(r.out);
  assert.ok(!k.names.some(n => /^ppt\/notesSlides\//.test(n)), '지운 차이 슬라이드의 노트는 함께 지움');
});

await test('PowerPoint 로 저장한 직전 보고서(줄 번호 a16:rowId · 슬라이드 번호 p14:creationId 가 붙은 양식) — 본뜬 장마다 번호가 겹치지 않음(11-08 TEST: 첫 차이 장만 보이던 결함)', async () => {
  // 리포의 가상 양식에 PowerPoint 가 저장할 때 붙이는 고유값을 넣어 「PowerPoint 로 한 번 저장한」 모양을 만듦
  const z = await JSZip.loadAsync(TPL);
  const A16 = 'xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main"';
  let seq = 0;
  for (const n of Object.keys(z.files).filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n))) {
    let x = await z.file(n).async('string');
    x = x.replace(/<\/a:tr>/g, () => '<a:extLst><a:ext uri="{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}"><a16:rowId ' + A16 + ' val="' + (3000000000 + (++seq)) + '"/></a:ext></a:extLst></a:tr>');
    x = x.replace(/<p:cNvPr id="(\d+)" name="([^"]*)"\/>/g, (m, id, nm) => '<p:cNvPr id="' + id + '" name="' + nm + '"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId ' + A16 + ' id="{6A1B2C3D-0000-4000-8000-' + String(++seq).padStart(12, '0') + '}"/></a:ext></a:extLst></p:cNvPr>');
    x = x.replace('</p:clrMapOvr>', '</p:clrMapOvr><p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"><p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="' + (2754183745 + (++seq)) + '"/></p:ext></p:extLst>');
    z.file(n, x);
  }
  const saved = await z.generateAsync({ type: 'nodebuffer' });
  await checkPkg(saved); // 넣은 양식 자체는 정상
  const r = await make(saved);
  assert.ok(r.log.diffPages >= 3, '차이 장이 여러 장이어야 검사가 뜻이 있음 ' + r.log.diffPages);
  const k = await checkPkg(r.out);
  assert.equal(k.ids.length, 1 + r.log.diffPages + 1);
  const d = await dump(r.out);
  const diffs = d.filter(s => s.tables.some(t => t.kind === 'diff'));
  assert.equal(diffs.length, r.log.diffPages, '차이 장마다 표가 남아 있음');
  assert.ok(diffs.every(s => s.tables.find(t => t.kind === 'diff').rows.length >= 2), '차이 장마다 머리 + 줄');
});

console.log('만들기 — 기본 양식');
await test('기본 양식(직전 보고서 없음)으로도 같은 짜임 — 1 + 2장 n장 + 1, 꾸러미 정상, 행은 자재타입(있으면)', async () => {
  const r = await make(null);
  const k = await checkPkg(r.out);
  assert.equal(k.ids.length, 2 + r.log.diffPages);
  assert.equal(r.log.statusDim, 'mkind');
  const d = await dump(r.out);
  assert.ok(d[0].tables.find(t => t.kind === 'status').rows.some(r => r[0] === '원자재'));
});
await test('차이가 하나도 없으면 2장을 넣지 않고 알림', async () => {
  const db0 = L.restoreDb({ session: S.session, items: S.items, assign: S.assign, records: [] });
  const r = await P.make(JSZip, TPL, L.analyze(db0), db0, { type: 'nodebuffer' });
  assert.equal(r.log.diffPages, 0); assert.ok(r.log.warnings.some(w => /차이가 있는 품목이 없어/.test(w)));
  const k = await checkPkg(r.out);
  assert.equal(k.ids.length, 2);
});
await test('실사완료 금액 = 실물 수량 × ERP 단가 — 전산 10개 · 200원 · 실물 9개 → 9개 · 1,800원', () => {
  const rows = [{ whCode: 'W1', site: '사내', mkind: '원자재', sys: 10, price: 200, sysAmt: 2000, counted: true, cnt: 9 },
    { whCode: 'W1', site: '사내', mkind: '원자재', sys: 5, price: 100, sysAmt: 500, counted: false, cnt: null }];
  const st = P.statusGroups({ rows, whNames: {} }, ['원자재'], 'mkind');
  const g = st.rows.find(x => x.label === '원자재');
  assert.equal(g.totQty, 15); assert.equal(g.totAmt, 2500);
  assert.equal(g.doneQty, 9, '실사완료 수량 = 실물'); assert.equal(g.doneAmt, 1800, '실사완료 금액 = 9 × 200');
  assert.equal(g.restQty, 5, '미실시 = 안 센 품목 전산'); assert.equal(g.restAmt, 500);
});
await test('pptx 가 아니면 알아듣게 거절', async () => {
  const z = new JSZip(); z.file('a.txt', 'x');
  await assert.rejects(P.inspect(JSZip, await z.generateAsync({ type: 'nodebuffer' })), /PowerPoint\(pptx\) 파일이 아닙니다/);
});

console.log('\n' + passed + '개 통과');
