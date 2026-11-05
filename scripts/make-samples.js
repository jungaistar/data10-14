// 가상 샘플 파일 만들기: node scripts/make-samples.js
// js/sample-data.js 의 값(전부 지어낸 값, 모양은 수강생이 보낸 실제 양식과 같음)으로 세 파일을 씁니다.
//   samples/ERP_재고현황_다운로드_샘플_2026-09.xlsx       — ERP 재고 엑셀(첫 줄 안내 · 빈 줄 2개 · 머리행 · 추가 열 예시)
//   samples/창고배정표_샘플_2026-09.xlsx                  — 창고 목록과 창고별 담당 인원(실사인원 0 = 이번 실사 제외)
//   samples/재고실사기록서_SharePoint취합_샘플.xlsx        — 이 도구가 낸 기록서를 SharePoint 에 올려 실사자E(자기 시트)·실사자C(전체 시트)가 채운 것
//   samples/재고실사_결과보고_직전양식_샘플.pptx           — 3차 반영: 「직전 보고서」 모양(1장 개요 · 실사 현황 / 2장 차이 내용 2장 / 3장 이슈, 작성 메모 상자). 회사명 · 숫자는 가상
// 쓴 파일을 앱과 같은 방식으로 다시 읽어 품목 수·배정·기록 수가 맞는지 확인합니다.
const fs = require('fs');
const path = require('path');
const XLSX = require('../vendor/xlsx.full.min.js');
const L = require('../js/logic.js');
const Sample = require('../js/sample-data.js');
const P = require('../js/pptx-report.js');
const JSZip = require('../vendor/jszip.min.js');

const out = path.join(__dirname, '..', 'samples');
fs.mkdirSync(out, { recursive: true });
for (const f of fs.readdirSync(out)) if (f.endsWith('.xlsx')) fs.unlinkSync(path.join(out, f)); // 1단계 샘플 이름이 남지 않게
const b = Sample.build();
function readBack(file) {
  const wb = XLSX.read(fs.readFileSync(file), { type: 'buffer', cellDates: false });
  return wb.SheetNames.map(n => ({ name: n, aoa: XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null }) }));
}
function write(file, sheets) {
  const wb = XLSX.utils.book_new();
  sheets.forEach(s => {
    const ws = XLSX.utils.aoa_to_sheet(s.aoa);
    (s.formula || []).forEach(f => { ws[f.a] = f.v == null ? { t: 's', v: '', f: f.f } : { t: 'n', v: f.v, f: f.f }; });
    if (s.cols) ws['!cols'] = s.cols;
    XLSX.utils.book_append_sheet(wb, ws, s.name);
  });
  fs.writeFileSync(file, XLSX.write(wb, { bookType: 'xlsx', type: 'buffer', compression: true }));
}

// 1) ERP 재고 엑셀
const f1 = path.join(out, b.erpFile);
write(f1, [{ name: '재고현황', aoa: b.erpAoa, cols: [{ wch: 10 }, { wch: 22 }, { wch: 8 }, { wch: 12 }, { wch: 8 }, { wch: 10 }, { wch: 13 }, { wch: 10 }, { wch: 8 }, { wch: 10 }, { wch: 8 }] }]);
const p = L.parseErp(L.pickSheet(readBack(f1)).aoa);
if (p.det.missing.length || p.items.length !== b.raw.length || p.det.headerRow !== 3) throw new Error('ERP 샘플 다시 읽기 실패: ' + p.items.length);
console.log(b.erpFile, '품목', p.items.length);

// 2) 창고 배정표
const f2 = path.join(out, b.rosterFile);
write(f2, [{ name: '창고배정표', aoa: b.rosterAoa, cols: [{ wch: 10 }, { wch: 12 }, { wch: 8 }, { wch: 8 }, { wch: 10 }, { wch: 10 }, { wch: 18 }, { wch: 34 }] }]);
const ro = L.parseRoster(readBack(f2)[0].aoa);
const ap = L.applyRoster(p.items, ro);
if (ro.error || ap.items.length !== b.items.length || JSON.stringify(ap.assign) !== JSON.stringify(b.assign) || ro.date !== '2026-09-30') throw new Error('배정표 샘플 다시 읽기 실패: ' + (ro.error || ap.items.length));
console.log(b.rosterFile, '창고', ro.rows.length, '· 제외', ap.excluded.map(x => x.whName).join(','), '· 실사자', L.assignedNames(ap.assign).join(','));

// 3) SharePoint 취합 파일(실사자E 자기 시트 40줄 + 실사자C 전체 시트 2줄)
const sheets = Sample.fillPaper(L.orderSheets(b.session, b.items, b.assign), b.heldOut, b.mainRows);
sheets.forEach((s, i) => { s.cols = i === 0 ? [{ wch: 20 }, { wch: 70 }] : s.aoa[0].map(h => ({ wch: /자재내역|특이사항/.test(h) ? 22 : /재고번호/.test(h) ? 18 : 10 })); });
const f3 = path.join(out, b.eFile);
write(f3, sheets);
const w = L.parseOrderWorkbook(readBack(f3), { fileName: b.eFile });
if (w.error || w.records.length !== b.heldOut.length + b.mainRows.length || w.session.id !== b.sessionId) throw new Error('SharePoint 샘플 다시 읽기 실패: ' + (w.error || w.records.length));
console.log(b.eFile, '기록', w.records.length, '·', w.warnings.join(' / '));

// 4) 직전 보고서 양식(pptx) — 수강생이 보낸 보고서와 같은 짜임(내용 칸은 비어 있고 「(계산로직)」 메모가 붙은 상태). 회사명은 가상
const PPT = '재고실사_결과보고_직전양식_샘플.pptx';
for (const f of fs.readdirSync(out)) if (f.endsWith('.pptx')) fs.unlinkSync(path.join(out, f));
const tpl = P.defaultTemplate(JSZip, { company: '가나산업', memos: true, extraDiffSlide: true, placeholder: true, statusLabels: ['원자재', '부자재', '제품', '반제품', '상품', '저장품', '포장재'] });
tpl.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }).then(buf => {
  fs.writeFileSync(path.join(out, PPT), buf);
  return P.inspect(JSZip, buf);
}).then(info => {
  if (!info.overview || !info.status || !info.diff || info.diffSlides.length !== 2 || !info.issue || info.memos !== 4) throw new Error('보고서 양식 샘플 다시 읽기 실패: ' + JSON.stringify(info).slice(0, 300));
  console.log(PPT, '슬라이드', info.slides, '· 차이 슬라이드', info.diffSlides.length, '· 메모', info.memos, '· 현황 행', info.status.labels.join(','));
}).catch(e => { console.error(e); process.exitCode = 1; });
