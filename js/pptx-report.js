/* 재고실사 Start-End — 보고서 PPT (직전 보고서를 올리면 디자인은 그대로 두고 내용만 바꿈)
   3차 반영(2026-10-27): 수강생 답 ① 「직전 보고서를 업로드하면 거기에 내용만 변경하여 업데이트된 아웃풋」
     1장 재고실사 개요 및 실사 현황 — 개요 표(실사 시행일 · 집계 기준 시점 · 실사 방법=직접 입력), 실사 현황 표
     2장 차이 내용 — 창고별로 한 장에 표 하나(넘치면 이어지는 장), 열: 자재코드 · 자재내역 · 저장위치 · 전산수량 · 실물수량 · 차이수량 · 단가 · 차이금액 · 차이 사유(및 조치사항)
     3장 이슈사항 및 개선 — 고칠 수 있는 글상자(도구가 초안을 채움)
   pptx 는 zip 안의 XML 묶음입니다. JSZip 으로 열어 표 칸(a:tc)의 글자만 바꾸고, 2장은 직전 보고서의 첫 차이 슬라이드를
   본떠 창고 수 · 쪽 수만큼 복제합니다(마스터 · 레이아웃 · 글꼴 · 색 · 테두리는 손대지 않음).
   직전 보고서가 없으면 기본 양식(defaultTemplate)을 만들어 같은 길로 채웁니다.
   브라우저에서는 window.STPptx, node 에서는 require('./pptx-report.js') — 화면과 무관한 순수 함수입니다. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./logic.js'));
  else root.STPptx = factory(root.STLogic);
})(typeof self !== 'undefined' ? self : this, function (L) {
  'use strict';

  var CT_SLIDE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
  var REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
  var NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
  var XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  var RED = 'FF0000'; // 도구가 칠하는 빨간 글자(차이수량 · 차이금액이 0 이 아닐 때). 다시 쓸 때 이 색만 걷어냅니다
  var PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';

  /* ── 1. XML 문자열 도구 (DOM 없이 — node 검사와 브라우저가 같은 코드) ───────── */
  function esc(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
  function unesc(s) {
    return String(s).replace(/&#x([0-9a-f]+);/gi, function (_, h) { return String.fromCodePoint(parseInt(h, 16)); })
      .replace(/&#(\d+);/g, function (_, d) { return String.fromCodePoint(+d); })
      .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
  }
  /* tag 요소의 [start, end) — 같은 이름이 안에 겹쳐 있어도(그룹 안 그룹) 바깥 것만 */
  function elems(xml, tag, start, end) {
    start = start || 0; end = end == null ? xml.length : end;
    var out = [], re = new RegExp('<' + tag + '(?=[\\s/>])|</' + tag + '>', 'g'), depth = 0, s = -1, m;
    re.lastIndex = start;
    while ((m = re.exec(xml)) && m.index < end) {
      if (m[0].charAt(1) === '/') { depth--; if (depth === 0) out.push({ s: s, e: m.index + m[0].length }); continue; }
      var gt = xml.indexOf('>', m.index), self = xml.charAt(gt - 1) === '/';
      if (depth === 0) s = m.index;
      if (self) { if (depth === 0) out.push({ s: s, e: gt + 1 }); } else depth++;
    }
    return out;
  }
  function first(xml, tag) { var x = elems(xml, tag)[0]; return x ? xml.slice(x.s, x.e) : ''; }
  function openTag(frag) { return frag.slice(0, frag.indexOf('>') + 1); }
  function getAttr(tag, name) { var m = new RegExp('\\s' + name + '="([^"]*)"').exec(tag); return m ? m[1] : null; }
  function setAttr(tag, name, val) {
    var re = new RegExp('(\\s' + name + '=")[^"]*(")');
    if (re.test(tag)) return tag.replace(re, '$1' + val + '$2');
    return tag.replace(/(\/?>)$/, ' ' + name + '="' + val + '"$1');
  }
  function delAttr(tag, name) { return tag.replace(new RegExp('\\s' + name + '="[^"]*"', 'g'), ''); }
  /* 글자 — 문단은 \n 으로 */
  function textOf(frag) {
    var ps = elems(frag, 'a:p');
    return ps.map(function (p) {
      var x = frag.slice(p.s, p.e), t = '', re = /<a:t>([^<]*)<\/a:t>|<a:t\/>|<a:br\/>/g, m;
      while ((m = re.exec(x))) t += m[0] === '<a:br/>' ? '\n' : m[1] ? unesc(m[1]) : '';
      return t;
    }).join('\n');
  }
  function nn(s) { return String(s == null ? '' : s).replace(/[\s\u00a0]/g, '').toLowerCase(); }
  function xfrmOf(frag) {
    var m = /<a:off x="(-?\d+)" y="(-?\d+)"\/>\s*<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(frag);
    return m ? { x: +m[1], y: +m[2], cx: +m[3], cy: +m[4] } : null;
  }

  /* ── 2. 글자 바꾸기 — 첫 문단 · 첫 글자의 모양(크기 · 굵기 · 색 · 글꼴)을 그대로 씀 ─── */
  function runProps(body) {
    var r = elems(body, 'a:rPr')[0];
    if (r) return body.slice(r.s, r.e);
    var e = elems(body, 'a:endParaRPr')[0];
    if (e) return body.slice(e.s, e.e).replace(/^<a:endParaRPr/, '<a:rPr').replace(/<\/a:endParaRPr>$/, '</a:rPr>');
    return '<a:rPr lang="ko-KR" altLang="en-US"/>';
  }
  /* 글자색: hex 를 주면 그 색으로, null 이면 도구가 칠한 빨강(RED)만 걷어 냄(원래 양식의 색은 둠) */
  var FILLS = ['a:solidFill', 'a:noFill', 'a:gradFill', 'a:pattFill', 'a:blipFill', 'a:grpFill'];
  function setColor(rPr, hex) {
    var ot = openTag(rPr), selfClose = /\/>$/.test(ot);
    ot = delAttr(delAttr(ot, 'err'), 'dirty');
    var inner = selfClose ? '' : rPr.slice(rPr.indexOf('>') + 1, rPr.length - '</a:rPr>'.length);
    var ln = '';
    if (/^<a:ln[\s>/]/.test(inner)) { var l = elems(inner, 'a:ln')[0]; ln = inner.slice(0, l.e); inner = inner.slice(l.e); }
    var fill = '';
    FILLS.some(function (f) { if (inner.indexOf('<' + f) === 0) { var x = elems(inner, f)[0]; fill = inner.slice(0, x.e); inner = inner.slice(x.e); return true; } return false; });
    if (hex) fill = '<a:solidFill><a:srgbClr val="' + hex + '"/></a:solidFill>';
    else if (fill === '<a:solidFill><a:srgbClr val="' + RED + '"/></a:solidFill>') fill = '';
    ot = ot.replace(/\/>$/, '>');
    return ot + ln + fill + inner + '</a:rPr>';
  }
  /* frag(표 칸 a:tc 또는 도형 p:sp)의 글자를 lines 로 바꿈. opts.red: 빨간 글자 */
  function setText(frag, text, opts) {
    opts = opts || {};
    var tag = frag.indexOf('<p:txBody') >= 0 ? 'p:txBody' : 'a:txBody';
    var tb = elems(frag, tag)[0];
    if (!tb) {
      if (tag === 'p:txBody') return frag;
      var ins = frag.indexOf('>') + 1;
      frag = frag.slice(0, ins) + '<a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="ko-KR"/></a:p></a:txBody>' + frag.slice(ins);
      tb = elems(frag, tag)[0];
    }
    var body = frag.slice(tb.s, tb.e), ps = elems(body, 'a:p');
    var head = ps.length ? body.slice(0, ps[0].s) : body.slice(0, body.lastIndexOf('</'));
    var p0 = ps.length ? body.slice(ps[0].s, ps[0].e) : '';
    var pPr = first(p0, 'a:pPr');
    var rPr = setColor(runProps(body), opts.red ? RED : null);
    var endP = rPr.replace(/^<a:rPr/, '<a:endParaRPr').replace(/<\/a:rPr>$/, '</a:endParaRPr>');
    var lines = String(text == null ? '' : text).split('\n');
    var paras = lines.map(function (ln) {
      return '<a:p>' + pPr + (ln !== '' ? '<a:r>' + rPr + '<a:t>' + esc(ln) + '</a:t></a:r>' : '') + endP + '</a:p>';
    }).join('');
    return frag.slice(0, tb.s) + head + paras + '</' + tag + '>' + frag.slice(tb.e);
  }

  /* ── 3. 표 ───────────────────────────────────────────── */
  function parseTable(gf) {
    var t = elems(gf, 'a:tbl')[0];
    if (!t) return null;
    var tbl = gf.slice(t.s, t.e);
    var rows = elems(tbl, 'a:tr').map(function (r) {
      var x = tbl.slice(r.s, r.e);
      var cells = elems(x, 'a:tc').map(function (c) {
        var cx = x.slice(c.s, c.e), ot = openTag(cx);
        return { xml: cx, text: textOf(cx), rowSpan: +getAttr(ot, 'rowSpan') || 1, gridSpan: +getAttr(ot, 'gridSpan') || 1, hMerge: getAttr(ot, 'hMerge') === '1', vMerge: getAttr(ot, 'vMerge') === '1' };
      });
      return { s: r.s, e: r.e, xml: x, h: +getAttr(openTag(x), 'h') || 0, cells: cells };
    });
    var cols = (elems(tbl, 'a:gridCol')).map(function (g) { return +getAttr(tbl.slice(g.s, g.e), 'w') || 0; });
    return { s: t.s, e: t.e, tbl: tbl, rows: rows, cols: cols, xf: xfrmOf(first(gf, 'p:xfrm')) };
  }
  function headText(T, n) { return T.rows.slice(0, n || 2).map(function (r) { return r.cells.map(function (c) { return nn(c.text); }).join('|'); }).join('|'); }
  /* 행 하나 — 칸마다 글자를 바꾸고 높이를 정함. 합치기 표시(rowSpan · vMerge · gridSpan · hMerge)는 걷어 냄 */
  function makeRow(rowXml, values, h) {
    var cells = elems(rowXml, 'a:tc'), out = '', pos = 0;
    cells.forEach(function (c, i) {
      var cx = rowXml.slice(c.s, c.e), ot = openTag(cx), ot2 = ot;
      ['rowSpan', 'gridSpan', 'hMerge', 'vMerge'].forEach(function (a) { ot2 = delAttr(ot2, a); });
      cx = ot2 + cx.slice(ot.length);
      var v = values[i];
      if (v !== undefined) cx = setText(cx, v && typeof v === 'object' ? v.t : v, { red: v && typeof v === 'object' && v.red });
      out += rowXml.slice(pos, c.s) + cx; pos = c.e;
    });
    out += rowXml.slice(pos);
    var ot = openTag(out);
    return setAttr(ot, 'h', String(Math.round(h || +getAttr(ot, 'h') || 300000))) + out.slice(ot.length);
  }
  function rebuildTable(gf, T, rowXmls) {
    var first0 = T.rows[0].s, lastE = T.rows[T.rows.length - 1].e;
    var tbl = T.tbl.slice(0, first0) + rowXmls.join('') + T.tbl.slice(lastE);
    gf = gf.slice(0, T.s) + tbl + gf.slice(T.e);
    var cy = rowXmls.reduce(function (s, r) { return s + (+getAttr(openTag(r), 'h') || 0); }, 0);
    return gf.replace(/(<p:xfrm>\s*<a:off [^>]*\/>\s*<a:ext cx="\d+" cy=")(\d+)(")/, '$1' + cy + '$3');
  }
  function isTotalLabel(t) { return /^(합계|총계|계|total|소계)$/.test(nn(t)); }
  function fontSz(xml) { var m = /<a:(?:rPr|endParaRPr|defRPr)[^>]*\ssz="(\d+)"/.exec(xml); return m ? +m[1] : 1000; }

  /* ── 4. 단위 「(단위 : 천 EA, 백만원)」 ─────────────────────── */
  function parseUnit(t) {
    var s = nn(t), q = 1, a = 1;
    if (/천(ea|개|pcs)/.test(s)) q = 1000; else if (/만(ea|개|pcs)/.test(s)) q = 10000;
    if (/백만원/.test(s)) a = 1e6; else if (/억원/.test(s)) a = 1e8; else if (/천만원/.test(s)) a = 1e7; else if (/만원/.test(s)) a = 1e4; else if (/천원/.test(s)) a = 1e3;
    return { q: q, a: a, text: t };
  }
  function fmtScaled(v, scale) {
    if (v == null || !isFinite(v)) return '-';
    if (scale === 1) return L.fmtNum(Math.round(v * 100) / 100);
    var x = Math.round(v / scale * 10) / 10;
    return x.toLocaleString('ko-KR', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  }
  function fmtPct(a, b) { return b ? (Math.round(a / b * 1000) / 10).toFixed(1) : '-'; }

  /* ── 5. 슬라이드 살펴보기 ────────────────────────────────── */
  var MEMO_RE = /^\((계산로직|작성안내|작성방법|메모)\)|ai가초안|초안작성해주면/;
  function shapesOf(xml) {
    var groups = elems(xml, 'p:grpSp');
    function inGroup(x) { return groups.some(function (g) { return x.s > g.s && x.e <= g.e; }); }
    return elems(xml, 'p:sp').map(function (x) {
      var frag = xml.slice(x.s, x.e), t = textOf(frag);
      return { s: x.s, e: x.e, xml: frag, text: t, n: nn(t), xf: xfrmOf(first(frag, 'p:spPr')), top: !inGroup(x) };
    });
  }
  function isMemo(sh) { return MEMO_RE.test(sh.n); }
  function unitNear(shapes, xf) {
    var best = null;
    shapes.forEach(function (s) {
      if (!/^\(?단위/.test(s.n) || !s.xf || !xf) return;
      var d = xf.y - s.xf.y;
      if (d >= -200000 && (!best || d < best.d)) best = { d: d, s: s };
    });
    return best ? parseUnit(best.s.text) : { q: 1, a: 1, text: '' };
  }
  function tablesOf(xml) {
    return elems(xml, 'p:graphicFrame').map(function (g) {
      var gf = xml.slice(g.s, g.e), T = parseTable(gf);
      return T ? { s: g.s, e: g.e, gf: gf, T: T, name: getAttr(openTag(first(gf, 'p:cNvPr')), 'name') || '' } : null;
    }).filter(Boolean);
  }
  function tableKind(T) {
    var h = headText(T, 2);
    if (/차이수량/.test(h) || (/전산수량|시스템수량|장부수량/.test(h) && /실물수량|실사수량/.test(h))) return 'diff';
    if (/전체재고|실사완료|미실시/.test(h)) return 'status';
    var firstCol = T.rows.map(function (r) { return r.cells[0] ? nn(r.cells[0].text) : ''; }).join('|');
    if (/시행일|기준시점|기준일자|기준시각/.test(firstCol)) return 'overview';
    return '';
  }
  function issueHeads(shapes) {
    var top = shapes.filter(function (s) { return s.top && s.xf; });
    var h1 = top.filter(function (s) { return /^(이슈사항|이슈|문제점|이슈사항및원인)$/.test(s.n); })[0];
    var h2 = top.filter(function (s) { return s !== h1 && /^(개선|조치|대책|개선및|개선사항)/.test(s.n); })[0];
    return h1 ? { h1: h1, h2: h2 || null } : null;
  }
  function issueCells(shapes, heads) {
    var h1 = heads.h1, h2 = heads.h2, y0 = h1.xf.y + h1.xf.cy * 0.5;
    var body = shapes.filter(function (s) { return s.top && s.xf && s !== h1 && s !== h2 && !isMemo(s) && s.xf.y >= y0; });
    function within(s, h) { var c = s.xf.x + s.xf.cx / 2; return h && c >= h.xf.x && c <= h.xf.x + h.xf.cx; }
    var by = function (a, b) { return a.xf.y - b.xf.y; };
    return {
      issue: body.filter(function (s) { return within(s, h1); }).sort(by),
      act: h2 ? body.filter(function (s) { return within(s, h2); }).sort(by) : [],
      label: body.filter(function (s) { return s.xf.x + s.xf.cx <= h1.xf.x + 60000; }).sort(by)
    };
  }

  /* 패키지 읽기 — 슬라이드 순서 · 관계 */
  function relsOf(xml) {
    var out = [], re = /<Relationship\s[^>]*\/?>/g, m;
    while ((m = re.exec(xml))) out.push({ xml: m[0], id: getAttr(m[0], 'Id'), type: getAttr(m[0], 'Type') || '', target: getAttr(m[0], 'Target') || '', mode: getAttr(m[0], 'TargetMode') });
    return out;
  }
  function resolvePath(base, target) {
    if (target.charAt(0) === '/') return target.slice(1);
    var parts = base.split('/'); parts.pop();
    target.split('/').forEach(function (p) { if (p === '..') parts.pop(); else if (p !== '.') parts.push(p); });
    return parts.join('/');
  }
  function relsPath(part) { var i = part.lastIndexOf('/'); return part.slice(0, i) + '/_rels/' + part.slice(i + 1) + '.rels'; }
  function readPkg(zip) {
    function txt(p) { var f = zip.file(p); return f ? f.async('string') : Promise.resolve(null); }
    return Promise.all([txt('ppt/presentation.xml'), txt('ppt/_rels/presentation.xml.rels'), txt('[Content_Types].xml')]).then(function (a) {
      if (!a[0] || !a[1] || !a[2]) throw new Error('PowerPoint(pptx) 파일이 아닙니다 — ppt/presentation.xml 이 없습니다.');
      var pres = a[0], prels = relsOf(a[1]);
      var sz = /<p:sldSz cx="(\d+)" cy="(\d+)"/.exec(pres);
      var lst = first(pres, 'p:sldIdLst');
      var ids = [], re = /<p:sldId\s[^>]*\/>/g, m;
      while ((m = re.exec(lst))) ids.push({ id: +getAttr(m[0], 'id'), rid: getAttr(m[0], 'r:id') });
      var slides = ids.map(function (x) { var r = prels.filter(function (q) { return q.id === x.rid; })[0]; return { id: x.id, rid: x.rid, path: r ? resolvePath('ppt/presentation.xml', r.target) : '' }; });
      return Promise.all(slides.map(function (s) { return Promise.all([txt(s.path), txt(relsPath(s.path))]); })).then(function (xs) {
        slides.forEach(function (s, i) { s.xml = xs[i][0] || ''; s.rels = xs[i][1] || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'; });
        return { pres: pres, presRels: a[1], ct: a[2], slides: slides, w: sz ? +sz[1] : 12192000, h: sz ? +sz[2] : 6858000 };
      });
    });
  }

  /* 화면에 보여 줄 구조 요약(무엇을 찾았는지) */
  function inspectPkg(pkg) {
    var info = { slides: pkg.slides.length, overview: null, status: null, diff: null, diffSlides: [], issue: null, memos: 0, otherSlides: [] };
    pkg.slides.forEach(function (s, i) {
      var shapes = shapesOf(s.xml), tabs = tablesOf(s.xml), kinds = [];
      info.memos += shapes.filter(isMemo).length;
      tabs.forEach(function (t) {
        var k = tableKind(t.T);
        if (k === 'overview' && !info.overview) { info.overview = { slide: i, rows: overviewRows(t.T) }; kinds.push(k); }
        if (k === 'status' && !info.status) { var st = statusLayout(t.T); info.status = { slide: i, labels: st.labels, cols: st.cols.length, unit: unitNear(shapes, t.T.xf).text, ok: st.ok }; kinds.push(k); }
        if (k === 'diff') { if (!info.diff) info.diff = { slide: i, head: t.T.rows[0].cells.map(function (c) { return c.text.replace(/\n/g, ' '); }), fields: diffFields(t.T).map(function (f) { return f.f; }), cap: diffCapacity(t.T, pkg.h, {}).cap, unit: unitNear(shapes, t.T.xf).text }; info.diffSlides.push(i); kinds.push(k); }
      });
      var heads = issueHeads(shapes);
      if (heads && !info.issue) { var c = issueCells(shapes, heads); info.issue = { slide: i, rows: Math.max(c.issue.length, c.label.length, 1), hasAct: !!heads.h2 }; kinds.push('issue'); }
      if (!kinds.length) info.otherSlides.push(i);
    });
    return info;
  }

  /* ── 6. 1장 — 개요 표 ────────────────────────────────── */
  /* 고칠 수 있는 칸: [{r, c, label, text, auto}] — 첫 칸이 이름표, 나머지가 내용 */
  function overviewRows(T) {
    var out = [], lastLabel = '';
    T.rows.forEach(function (row, r) {
      if (r === 0) return;
      var c0 = row.cells[0];
      var label = c0 && !c0.vMerge ? c0.text.trim() : lastLabel;
      if (c0 && !c0.vMerge) lastLabel = label;
      row.cells.forEach(function (c, ci) {
        if (ci === 0 || c.hMerge || c.vMerge) return;
        var n = nn(label), auto = /시행일|실사일/.test(n) ? 'countAt' : /기준/.test(n) ? 'baseAt' : '';
        out.push({ r: r, c: ci, label: label + (row.cells.filter(function (x, k) { return k > 0 && !x.hMerge && !x.vMerge; }).length > 1 ? ' (' + ci + ')' : ''), text: c.text, auto: auto });
      });
    });
    return out;
  }
  function fillOverview(gf, T, data, edits) {
    var rows = overviewRows(T), byKey = {};
    rows.forEach(function (x) { byKey[x.r + ',' + x.c] = x; });
    var rowXmls = T.rows.map(function (row, r) {
      var vals = {};
      row.cells.forEach(function (c, ci) {
        var x = byKey[r + ',' + ci]; if (!x) return;
        var k = r + ',' + ci;
        if (edits && edits[k] != null) vals[ci] = edits[k];
        else if (x.auto && data.overview[x.auto]) vals[ci] = data.overview[x.auto];
      });
      if (!Object.keys(vals).length) return row.xml;
      // 합치기 표시는 그대로 두고 글자만 바꿈
      var cells = elems(row.xml, 'a:tc'), out = '', pos = 0;
      cells.forEach(function (c, ci) {
        var cx = row.xml.slice(c.s, c.e);
        if (vals[ci] !== undefined) cx = setText(cx, vals[ci]);
        out += row.xml.slice(pos, c.s) + cx; pos = c.e;
      });
      return out + row.xml.slice(pos);
    });
    return rebuildTable(gf, T, rowXmls);
  }

  /* ── 7. 1장 — 실사 현황 표 ──────────────────────────────── */
  function statusLayout(T) {
    var r0 = T.rows[0], hc = 1;
    r0.cells.forEach(function (c) { if (c.rowSpan > hc) hc = c.rowSpan; });
    if (hc === 1 && T.rows[1] && T.rows[1].cells[0] && T.rows[1].cells[0].vMerge) hc = 2;
    hc = Math.min(hc, T.rows.length - 1);
    var ncol = T.cols.length || r0.cells.length, cols = [];
    var lastGroup = '', lastMeasure = { tot: 'qty', done: 'qty', rest: 'qty' };
    for (var c = 0; c < ncol; c++) {
      var k = c; while (k > 0 && r0.cells[k] && r0.cells[k].hMerge) k--;
      var top = r0.cells[k] ? nn(r0.cells[k].text) : '', sub = hc > 1 && T.rows[hc - 1].cells[c] ? nn(T.rows[hc - 1].cells[c].text) : '';
      var all = top + '|' + sub, f = '';
      if (c === 0) f = 'label';
      else {
        var g = /미실시|미실사|잔여|미완료/.test(all) ? 'rest' : /실사완료|완료|실사수량|실사금액/.test(all) ? 'done' : /전체|장부|재고|전산/.test(all) ? 'tot' : lastGroup;
        var m = /%|율|비율|진행/.test(sub || top) ? 'pct' : /금액|원/.test(sub || top) ? 'amt' : /수량|ea|개/.test(sub || top) ? 'qty' : '';
        if (g && m === 'pct') f = g + 'Pct' + (lastMeasure[g] === 'amt' ? 'Amt' : 'Qty');
        else if (g && m) { f = g + (m === 'amt' ? 'Amt' : 'Qty'); lastMeasure[g] = m; }
        if (g) lastGroup = g;
      }
      cols.push(f);
    }
    var data = T.rows.slice(hc), totalIdx = -1;
    data.forEach(function (r, i) { if (r.cells[0] && isTotalLabel(r.cells[0].text)) totalIdx = hc + i; });
    var labels = T.rows.slice(hc).filter(function (r, i) { return hc + i !== totalIdx; }).map(function (r) { return r.cells[0] ? r.cells[0].text.trim() : ''; });
    return { hc: hc, cols: cols, labels: labels, totalIdx: totalIdx, ok: cols.filter(Boolean).length > 2 };
  }
  function statusCell(f, g, U) {
    if (f === 'label') return g.label;
    if (g.excluded) return '-';
    switch (f) {
      case 'totQty': return fmtScaled(g.totQty, U.q);
      case 'totAmt': return fmtScaled(g.totAmt, U.a);
      case 'doneQty': return fmtScaled(g.doneQty, U.q);
      case 'doneAmt': return fmtScaled(g.doneAmt, U.a);
      case 'donePctQty': return fmtPct(g.doneQty, g.totQty);
      case 'donePctAmt': return fmtPct(g.doneAmt, g.totAmt);
      case 'restQty': return fmtScaled(g.restQty, U.q);
      case 'restAmt': return fmtScaled(g.restAmt, U.a);
      case 'restPctQty': return fmtPct(g.restQty, g.totQty);
      case 'restPctAmt': return fmtPct(g.restAmt, g.totAmt);
    }
    return undefined;
  }
  function fillStatus(gf, T, data, U, opts) {
    var L0 = statusLayout(T);
    var st = statusGroups(data, L0.labels, opts.statusBy);
    var head = T.rows.slice(0, L0.hc).map(function (r) { return r.xml; });
    var tplRow = (T.rows[L0.hc] && L0.hc !== L0.totalIdx ? T.rows[L0.hc] : T.rows[L0.hc - 1]);
    var totRow = L0.totalIdx >= 0 ? T.rows[L0.totalIdx] : tplRow;
    var h = tplRow.h > 0 ? tplRow.h : 226000;
    function vals(g) { var v = {}; L0.cols.forEach(function (f, i) { if (f) v[i] = statusCell(f, g, U); }); return v; }
    var body = st.rows.map(function (g) { return makeRow(tplRow.xml, vals(g), h); });
    var tot = makeRow(totRow.xml, vals(st.total), totRow.h > 0 ? totRow.h : h);
    return { gf: rebuildTable(gf, T, head.concat(body, [tot])), dim: st.dim, rows: st.rows.length };
  }

  /* ── 8. 2장 — 차이 내용 표 ──────────────────────────────── */
  function diffFieldOf(h) {
    var s = nn(h);
    if (/사유/.test(s) && /조치|대책/.test(s)) return 'reasonNote';
    if (/사유/.test(s)) return 'reason';
    if (/조치|특이|비고|대책/.test(s)) return 'note';
    if (/차이금액|차액/.test(s)) return 'diffAmt';
    if (/차이수량|수량차이|^차이$/.test(s)) return 'diff';
    if (/전산|시스템|장부/.test(s)) return 'sys';
    if (/실물|실사수량/.test(s)) return 'cnt';
    if (/단가/.test(s)) return 'price';
    if (/저장위치명|창고명/.test(s)) return 'whName';
    if (/창고위치|사내|사외/.test(s)) return 'site';
    if (/저장위치|^창고$|창고코드/.test(s)) return 'whCode';
    if (/location|로케이션|보관|랙/.test(s)) return 'loc';
    if (/자재내역|자재명|품명|품목명|내역/.test(s)) return 'name';
    if (/자재코드|^자재$|품번|코드|재고번호/.test(s)) return 'code';
    if (/규격/.test(s)) return 'spec';
    if (/단위/.test(s)) return 'unit';
    return '';
  }
  function diffFields(T) { return T.rows[0].cells.map(function (c, i) { return { i: i, f: c.hMerge ? '' : diffFieldOf(c.text) }; }); }
  function marginsOf(tcXml) {
    var tp = openTag(first(tcXml, 'a:tcPr') || '<a:tcPr/>');
    return (+(getAttr(tp, 'marT') || 45720)) + (+(getAttr(tp, 'marB') || 45720));
  }
  /* 한 장에 넣을 줄 수 — 표 위쪽부터 슬라이드 아래 여백(6.5%)까지, 직전 보고서의 가장 낮은 줄 높이로 */
  function diffCapacity(T, slideH, opts) {
    var top = T.xf ? T.xf.y : 1800000, limit = Math.round(slideH * 0.935);
    var head = T.rows[0], tot = T.rows.filter(function (r, i) { return i > 0 && r.cells[0] && isTotalLabel(r.cells[0].text); })[0];
    var dataRows = T.rows.filter(function (r, i) { return i > 0 && r !== tot; });
    var tpl = dataRows[0] || head;
    var sz = fontSz(tpl.xml), minH = Math.round(sz / 100 * 12700 * 1.25 + marginsOf(elems(tpl.xml, 'a:tc')[0] ? tpl.xml.slice(elems(tpl.xml, 'a:tc')[0].s) : ''));
    var hs = dataRows.map(function (r) { return r.h; }).filter(function (x) { return x > 0; });
    var rowH = Math.max(minH, hs.length ? Math.min.apply(null, hs) : minH);
    var headH = head.h || 400000, totH = tot ? tot.h || rowH : rowH;
    var avail = limit - top - headH - totH;
    var cap = Math.max(1, Math.floor(avail / rowH)), warn = '';
    var want = +opts.rowsPerPage || 0;
    if (want > 0) {
      var fit = Math.floor(avail / minH);
      if (want > fit) { warn = '한 장 ' + want + '줄은 글자 크기(' + sz / 100 + 'pt)로 들어가지 않아 ' + fit + '줄로 나눴습니다.'; want = fit; }
      cap = Math.max(1, want); rowH = Math.max(minH, Math.floor(avail / cap));
      if (rowH > 600000) rowH = Math.max(minH, Math.min(rowH, hs.length ? Math.max.apply(null, hs) : rowH));
    }
    return { cap: cap, rowH: rowH, headH: headH, totH: totH, tpl: tpl, tot: tot, sz: sz, minH: minH, warn: warn };
  }
  /* 칸 너비 · 줄 높이에 들어가게 글자를 줄임(넘치면 PowerPoint 가 줄을 늘려 표가 슬라이드 밖으로 나감) */
  function fitText(t, colW, rowH, sz, marH) {
    t = String(t == null ? '' : t).replace(/\s+/g, ' ').trim();
    if (!t) return t;
    var lineH = sz / 100 * 12700 * 1.2, lines = Math.max(1, Math.floor((rowH - marH) / lineH));
    var perLine = Math.max(4, Math.floor((colW - 20000) / (sz / 100 * 12700 * 0.95)));
    var max = lines * perLine;
    var w = 0, out = '';
    for (var i = 0; i < t.length; i++) { var ch = t.charAt(i); w += /[\u0000-\u00ff]/.test(ch) ? 0.55 : 1; if (w > max) return out.replace(/\s+$/, '') + '…'; out += ch; }
    return out;
  }
  function diffCell(f, r, U, cut) {
    switch (f) {
      case 'code': return r.extra ? (r.code || '(목록에 없음)') : r.code || r.no;
      case 'name': return cut(r.name || '');
      case 'whCode': return r.whCode;
      case 'whName': return r.whName;
      case 'site': return r.site;
      case 'loc': return r.loc || '';
      case 'spec': return r.spec || '';
      case 'unit': return r.unit || '';
      case 'sys': return fmtScaled(r.sys, U.q);
      case 'cnt': return fmtScaled(r.cnt, U.q);
      case 'diff': return { t: fmtScaled(r.diff, U.q), red: r.diff !== 0 };
      case 'price': return L.fmtNum(L.round(r.price, 2));
      case 'diffAmt': return { t: fmtScaled(r.diffAmt, U.a), red: Math.round(r.diffAmt) !== 0 };
      case 'reason': return cut(r.reason);
      case 'note': return cut(r.note);
      case 'reasonNote': return cut(r.reason + (r.note ? ' — ' + r.note : ''));
    }
    return '';
  }
  function diffTotalCell(f, g, U) {
    if (f === 'diff') return { t: fmtScaled(g.sumDiff, U.q), red: g.sumDiff !== 0 };
    if (f === 'diffAmt') return { t: fmtScaled(g.sumAmt, U.a), red: Math.round(g.sumAmt) !== 0 };
    return '';
  }
  /* 한 창고의 한 쪽 표 */
  function diffTable(gf, T, g, pageRows, isLast, C, U) {
    var fields = diffFields(T);
    var cutFor = {};
    fields.forEach(function (x) { cutFor[x.i] = function (t) { return fitText(t, T.cols[x.i] || 900000, C.rowH, C.sz, 91440); }; });
    var head = [T.rows[0].xml];
    var body = pageRows.map(function (r) {
      var v = {}; fields.forEach(function (x) { v[x.i] = x.f ? diffCell(x.f, r, U, cutFor[x.i]) : ''; });
      return makeRow(C.tpl.xml, v, C.rowH);
    });
    var rows = head.concat(body);
    if (isLast) {
      var totXml = C.tot ? C.tot.xml : C.tpl.xml;
      var v = {}; fields.forEach(function (x, i) { v[x.i] = i === 0 ? (C.tot ? C.tot.cells[0].text.trim() || '합계' : '합계') : x.f ? diffTotalCell(x.f, g, U) : ''; });
      rows.push(makeRow(totXml, v, C.totH));
    }
    return rebuildTable(gf, T, rows);
  }
  function setTitle(xml, sh, text) { var nx = setText(sh.xml, text); return xml.slice(0, sh.s) + nx + xml.slice(sh.e); }
  function removeMemos(xml) {
    var memos = shapesOf(xml).filter(isMemo).sort(function (a, b) { return b.s - a.s; });
    memos.forEach(function (m) { xml = xml.slice(0, m.s) + xml.slice(m.e); });
    return { xml: xml, n: memos.length };
  }
  function replaceAt(xml, at, len, s) { return xml.slice(0, at) + s + xml.slice(at + len); }

  /* ── 9. 3장 — 이슈사항 및 개선 ──────────────────────────── */
  function fillIssue(xml, lines) {
    var shapes = shapesOf(xml), heads = issueHeads(shapes);
    if (!heads) return { xml: xml, ok: false };
    var c = issueCells(shapes, heads), n = Math.max(c.issue.length, c.label.length);
    var edits = [];
    function put(list, i, text) { if (list[i]) edits.push({ sh: list[i], t: text }); }
    for (var i = 0; i < n; i++) {
      // 줄이 칸보다 많으면 남은 줄은 마지막 칸에 「주제: 내용」으로 이어 붙임
      var rows = i === n - 1 ? lines.slice(i) : lines.slice(i, i + 1);
      if (!rows.length) rows = [{ label: '', issue: '', act: '' }];
      var tag = function (x, k, t) { return k > 0 && x.label && t ? x.label + ': ' + t : t; };
      put(c.label, i, rows[0].label || '');
      put(c.issue, i, rows.map(function (x, k) { return tag(x, k, x.issue); }).filter(Boolean).join('\n'));
      put(c.act, i, rows.map(function (x, k) { return tag(x, k, x.act); }).filter(Boolean).join('\n'));
    }
    edits.sort(function (a, b) { return b.sh.s - a.sh.s; }).forEach(function (e) { xml = xml.slice(0, e.sh.s) + setText(e.sh.xml, e.t) + xml.slice(e.sh.e); });
    return { xml: xml, ok: true, rows: n };
  }
  /* 「라벨 | 이슈 | 개선」 한 줄씩. 한 칸 안에서 「 ; 」 는 줄바꿈 */
  function parseIssueText(t) {
    return String(t || '').split('\n').map(function (ln) { return ln.trim(); }).filter(Boolean).map(function (ln) {
      var p = ln.split('|').map(function (x) { return x.trim().replace(/\s*;\s*/g, '\n'); });
      return p.length === 1 ? { label: '', issue: p[0], act: '' } : { label: p[0] || '', issue: p[1] || '', act: p[2] || '' };
    });
  }
  function issueTextOf(lines) { return lines.map(function (x) { return [x.label, x.issue, x.act].map(function (s) { return String(s || '').replace(/\n/g, ' ; '); }).join(' | '); }).join('\n'); }

  /* ── 10. 숫자 모으기 (분석 결과 → 보고서 PPT 에 들어갈 값) ───────────── */
  function yymmdd(iso) { return iso ? iso.slice(2, 4) + '.' + iso.slice(5, 7) + '.' + iso.slice(8, 10) + '  ' + iso.slice(11, 16) : ''; }
  function whNames(db) { var m = {}; (db.items || []).concat(db.excludedItems || [], (db.session && db.session.excluded) || []).forEach(function (it) { if (it.wh && it.whName && !m[it.wh]) m[it.wh] = it.whName; }); return m; }
  var ACTIONS = {
    '입출고 누락': '실사 기준 시각 전후 입출고 전표 마감 확인 · 미처리 전표 정리 후 재대사',
    '위치 이동·분산': 'Location 재정리와 ERP 위치 정보 갱신 · 분산 보관 품목 표시',
    '라벨·식별 오류': '라벨 재부착 · 유사 품목 분리 보관 · 품번 확인 절차',
    '파손·폐기 미반영': '파손 · 폐기 처리 절차 정리와 ERP 반영 시점 합의',
    '조사 오류': '재실사(2인 확인) · 계수 단위(박스/낱개) 표준화',
    '기타': '특이사항 확인 후 담당 부서와 조치 협의'
  };
  ACTIONS[L.NO_REASON] = '사유 미입력 품목의 사유 확인(현장 담당자)';
  function issueDraft(an, n) {
    var rs = an.reasons.filter(function (x) { return x.items; }).sort(function (a, b) { return b.absAmt - a.absAmt; });
    var lines = rs.slice(0, Math.max(1, n || 3)).map(function (x) {
      var whs = {};
      an.rows.forEach(function (r) { if (r.counted && r.diff !== 0 && r.reason === x.key) whs[r.wh] = (whs[r.wh] || 0) + 1; });
      var top = Object.keys(whs).sort(function (a, b) { return whs[b] - whs[a]; }).slice(0, 2).map(function (k) { return k + ' ' + whs[k] + '품목'; });
      return {
        label: x.key,
        issue: x.items + '품목 · 순차이 ' + L.fmtSigned(x.netAmt, true) + ' (절대 ' + L.fmtWon(x.absAmt) + ')' + (top.length ? ' ; 주로 ' + top.join(', ') : ''),
        act: ACTIONS[x.key] || '원인 확인 후 조치 협의'
      };
    });
    if (!lines.length) lines.push({ label: '차이 없음', issue: '실사 품목 모두 전산수량과 일치', act: '-' });
    var extra = [];
    if (an.all.items > an.all.counted) extra.push('미실사 ' + (an.all.items - an.all.counted) + '품목');
    if (an.conflicts.length) extra.push('미해결 충돌 ' + an.conflicts.length + '품목');
    if (an.all.noReason) extra.push('사유 미입력 ' + an.all.noReason + '품목');
    if (extra.length) lines.push({ label: '진행', issue: extra.join(' · '), act: '보고 전 확인 · 재실사 여부 결정' });
    return lines.map(function (x) { return { label: x.label, issue: x.issue.replace(/ ; /g, '\n'), act: x.act }; });
  }
  /* 실사 현황 표의 행 — dim: wh(창고) · site(창고위치 사내/사외) · mtype(자재유형) · mkind(자재타입) · auto(표의 행 이름으로 고름) */
  var STATUS_DIMS = ['wh', 'site', 'mkind', 'mtype'];
  function statusGroups(data, labels, want) {
    var rows = data.rows, names = data.whNames || {};
    function keyOf(r, dim) { return dim === 'wh' ? r.whCode : dim === 'site' ? r.site : r[dim]; }
    function labelOf(k, dim) { return dim === 'wh' ? names[k] || k : k; }
    function aliases(k, dim) { return dim === 'wh' ? [k, names[k] || '', k + (names[k] || ''), (names[k] || '') + k] : [k]; }
    var groupsBy = {};
    STATUS_DIMS.forEach(function (dim) {
      var m = {};
      rows.forEach(function (r) { var k = keyOf(r, dim); if (k == null || k === '' || /^\(.*\)$/.test(k)) return; m[k] = 1; });
      if (dim === 'wh') (data.excluded || []).forEach(function (x) { m[x.wh] = 1; });
      groupsBy[dim] = Object.keys(m);
    });
    function matches(dim) {
      var hit = {};
      labels.forEach(function (lb, i) { var n = nn(lb); if (!n) return; groupsBy[dim].forEach(function (k) { if (aliases(k, dim).some(function (a) { return nn(a) === n; })) hit[i] = k; }); });
      return hit;
    }
    var dim = want && want !== 'auto' ? want : null, hit = {};
    if (!dim) {
      // 2026-10-27 학생 답: 실사 현황 표는 「자재타입」별 — ERP 에 자재타입이 있으면 행 이름이 안 맞을 때의 기본값도 자재타입
      var best = { dim: groupsBy.mkind && groupsBy.mkind.length ? 'mkind' : 'wh', n: 0, hit: {} };
      STATUS_DIMS.forEach(function (d) { if (!groupsBy[d].length) return; var h = matches(d), n = Object.keys(h).length; if (n > best.n) best = { dim: d, n: n, hit: h }; });
      dim = best.dim; hit = best.hit;
    } else hit = matches(dim);
    var keepLabels = Object.keys(hit).length > 0;
    var order = [];
    if (keepLabels) labels.forEach(function (lb, i) { order.push(hit[i] != null ? { k: hit[i], label: lb } : { k: null, label: lb }); });
    var have = {}; order.forEach(function (o) { if (o.k != null) have[o.k] = 1; });
    groupsBy[dim].slice().sort(function (a, b) { return a < b ? -1 : a > b ? 1 : 0; }).forEach(function (k) { if (!have[k]) order.push({ k: k, label: labelOf(k, dim) }); });
    var exMap = {}; if (dim === 'wh') (data.excluded || []).forEach(function (x) { exMap[x.wh] = x; });
    var total = { label: '합계', totQty: 0, totAmt: 0, doneQty: 0, doneAmt: 0, restQty: 0, restAmt: 0 };
    var out = order.map(function (o) {
      var g = { label: o.label, key: o.k, totQty: 0, totAmt: 0, doneQty: 0, doneAmt: 0, restQty: 0, restAmt: 0, items: 0, excluded: o.k == null || !!exMap[o.k] && !rows.some(function (r) { return keyOf(r, dim) === o.k; }) };
      if (o.k != null) rows.forEach(function (r) { if (keyOf(r, dim) !== o.k) return; g.items++; g.totQty += r.sys; g.totAmt += r.sysAmt; if (r.counted) { var cq = Number(r.cnt) || 0; g.doneQty += cq; g.doneAmt += cq * (Number(r.price) || 0); } else { g.restQty += r.sys; g.restAmt += r.sysAmt; } });
      // 실사완료 = 실사를 마친 품목의 실물 수량 · 실물 수량 × ERP 단가(학생 답 10-27: 전산 10개·200원 → 실물 9개 → 1,800원), 미실시 = 아직 안 센 품목의 전산 수량 · 금액
      if (!g.excluded) ['totQty', 'totAmt', 'doneQty', 'doneAmt', 'restQty', 'restAmt'].forEach(function (k) { total[k] += g[k]; });
      return g;
    });
    return { dim: dim, rows: out, total: total, keptLabels: keepLabels };
  }
  /* 2장 — 차이 품목을 창고(또는 창고위치)별로 */
  function diffGroups(data, by) {
    var names = data.whNames || {}, m = {}, list = [];
    data.rows.forEach(function (r) {
      if (!r.counted || r.diff === 0) return;
      var k = by === 'site' ? r.site : r.whCode;
      if (!m[k]) { m[k] = { key: k, label: by === 'site' ? k : names[k] || k, rows: [], sumDiff: 0, sumAmt: 0 }; list.push(m[k]); }
      m[k].rows.push(r); m[k].sumDiff += r.diff; m[k].sumAmt += r.diffAmt;
    });
    list.sort(function (a, b) { return a.key < b.key ? -1 : a.key > b.key ? 1 : 0; });
    list.forEach(function (g) {
      g.rows.sort(function (a, b) { return Math.abs(b.diffAmt) - Math.abs(a.diffAmt) || (a.code < b.code ? -1 : 1); });
      g.sumDiff = L.round(g.sumDiff, 6);
    });
    return list;
  }
  function reportData(an, db) {
    var names = whNames(db);
    var rows = an.rows.map(function (r) { var o = {}; for (var k in r) o[k] = r[k]; o.whName = names[r.whCode] || ''; return o; });
    var ats = rows.filter(function (r) { return r.counted && r.at; }).map(function (r) { return r.at; }).sort();
    var s0 = ats[0], s1 = ats[ats.length - 1];
    var countAt = s0 ? yymmdd(s0) + (s1 && s1.slice(0, 16) !== s0.slice(0, 16) ? ' ~ ' + (s1.slice(0, 10) === s0.slice(0, 10) ? s1.slice(11, 16) : yymmdd(s1)) : '') : '';
    return {
      rows: rows, whNames: names, excluded: an.excluded || [],
      overview: { countAt: countAt, baseAt: db.session ? yymmdd(db.session.baseAt) : '' }
    };
  }

  /* ── 11. 만들기 ──────────────────────────────────────── */
  function maxNum(list, re) { var m = 0; list.forEach(function (s) { var x = re.exec(s); if (x && +x[1] > m) m = +x[1]; }); return m; }
  /* opts: { statusBy, diffBy('wh'|'site'), rowsPerPage, removeMemos(기본 true), overviewEdits{"r,c":text}, issues[{label,issue,act}], type('uint8array'|'blob') } */
  function build(JSZip, zip, an, db, opts) {
    opts = opts || {};
    var data = reportData(an, db), log = { warnings: [], diffPages: 0, groups: 0, removedMemos: 0, statusDim: '', slides: 0 };
    return readPkg(zip).then(function (pkg) {
      var info = inspectPkg(pkg);
      if (!info.diff) log.warnings.push('2장 차이 내용 표(전산수량 · 실물수량 · 차이수량 열)를 찾지 못해 2장은 그대로 두었습니다.');
      if (!info.overview) log.warnings.push('1장 개요 표(실사 시행일 · 집계 기준 시점)를 찾지 못했습니다.');
      if (!info.status) log.warnings.push('1장 실사 현황 표(전체재고 · 실사완료 · 미실시)를 찾지 못했습니다.');
      if (!info.issue) log.warnings.push('3장 이슈사항 슬라이드(「이슈 사항」 머리 상자)를 찾지 못했습니다.');
      var rm = opts.removeMemos !== false;
      // ① 1장 · 3장 · 그 밖의 슬라이드 — 제자리에서 글자만
      pkg.slides.forEach(function (s, i) {
        if (info.diffSlides.indexOf(i) >= 0) return;
        var xml = s.xml, changed = false;
        if (rm) { var r = removeMemos(xml); if (r.n) { xml = r.xml; changed = true; log.removedMemos += r.n; } }
        var shapes = shapesOf(xml);
        tablesOf(xml).sort(function (a, b) { return b.s - a.s; }).forEach(function (t) {
          var k = tableKind(t.T), gf = null;
          if (k === 'overview' && info.overview && info.overview.slide === i) gf = fillOverview(t.gf, t.T, data, opts.overviewEdits);
          if (k === 'status' && info.status && info.status.slide === i) { var fs = fillStatus(t.gf, t.T, data, unitNear(shapes, t.T.xf), opts); gf = fs.gf; log.statusDim = fs.dim; log.statusRows = fs.rows; }
          if (gf) { xml = replaceAt(xml, t.s, t.e - t.s, gf); changed = true; }
        });
        if (info.issue && info.issue.slide === i) {
          var fi = fillIssue(xml, opts.issues && opts.issues.length ? opts.issues : issueDraft(an, info.issue.rows));
          if (fi.ok) { xml = fi.xml; changed = true; }
        }
        if (changed) { s.xml = xml; zip.file(s.path, xml); }
      });
      // ② 2장 — 첫 차이 슬라이드를 본떠 창고 · 쪽마다 새로 만들고, 원래 차이 슬라이드는 지움
      if (info.diff) {
        var tplSlide = pkg.slides[info.diff.slide];
        var txml = tplSlide.xml;
        if (rm) { var rr = removeMemos(txml); txml = rr.xml; log.removedMemos += rr.n; }
        var t0 = tablesOf(txml).filter(function (t) { return tableKind(t.T) === 'diff'; })[0];
        var C = diffCapacity(t0.T, pkg.h, opts);
        if (C.warn) log.warnings.push(C.warn);
        var U = unitNear(shapesOf(txml), t0.T.xf);
        var titleSh = shapesOf(txml).filter(function (s) { return /^차이내용/.test(s.n); })[0];
        var prefix = titleSh ? titleSh.text.split(/[:：]/)[0].trim() : '차이 내용';
        var groups = diffGroups(data, opts.diffBy === 'site' ? 'site' : 'wh');
        var pages = [];
        groups.forEach(function (g) {
          var n = Math.max(1, Math.ceil(g.rows.length / C.cap));
          for (var p = 0; p < n; p++) pages.push({ g: g, rows: g.rows.slice(p * C.cap, (p + 1) * C.cap), p: p + 1, n: n });
        });
        if (!groups.length) log.warnings.push('차이가 있는 품목이 없어 2장 슬라이드를 넣지 않았습니다.');
        log.groups = groups.length; log.diffPages = pages.length; log.cap = C.cap;
        var newXml = pages.map(function (pg) {
          var x = txml;
          var tt = tablesOf(x).filter(function (t) { return tableKind(t.T) === 'diff'; })[0];
          x = replaceAt(x, tt.s, tt.e - tt.s, diffTable(tt.gf, tt.T, pg.g, pg.rows, pg.p === pg.n, C, U));
          var sh = shapesOf(x).filter(function (s) { return /^차이내용/.test(s.n); })[0];
          if (sh) x = setTitle(x, sh, prefix + ' : ' + pg.g.label + (pg.n > 1 ? ' (' + pg.p + '/' + pg.n + ')' : ''));
          return x;
        });
        return replaceDiffSlides(zip, pkg, info.diffSlides, tplSlide, newXml).then(function (cnt) { log.slides = cnt; return finish(); });
      }
      log.slides = pkg.slides.length;
      return finish();
      function finish() {
        return zip.generateAsync({ type: opts.type || 'uint8array', mimeType: PPTX_MIME, compression: 'DEFLATE', compressionOptions: { level: 6 } }).then(function (out) { return { out: out, log: log, info: info }; });
      }
    });
  }
  /* 슬라이드 갈아 끼우기 — [Content_Types].xml · presentation.xml(sldIdLst · 구역) · presentation.xml.rels · 슬라이드 rels · 메모 슬라이드 */
  function replaceDiffSlides(zip, pkg, removeIdx, tplSlide, newXmls) {
    var pres = pkg.pres, prels = pkg.presRels, ct = pkg.ct;
    var removed = removeIdx.map(function (i) { return pkg.slides[i]; });
    var anchorId = removed[0].id;
    var tplRels = relsOf(tplSlide.rels).filter(function (r) { return !/\/notesSlide$/.test(r.type); });
    var relsXml = XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + tplRels.map(function (r) { return r.xml; }).join('') + '</Relationships>';
    // 새 이름 · 번호
    var names = Object.keys(zip.files);
    var nextPart = maxNum(names, /^ppt\/slides\/slide(\d+)\.xml$/) + 1;
    var nextRid = maxNum(relsOf(prels).map(function (r) { return r.id; }), /^rId(\d+)$/) + 1;
    var nextId = Math.max(255, maxNum(pkg.slides.map(function (s) { return 'x' + s.id; }), /^x(\d+)$/)) + 1;
    var added = newXmls.map(function (x) { var o = { path: 'ppt/slides/slide' + nextPart + '.xml', rid: 'rId' + nextRid, id: nextId, xml: x }; nextPart++; nextRid++; nextId++; return o; });
    // 지우기 — 슬라이드 · 그 rels · 메모 슬라이드
    var jobs = removed.map(function (s) {
      relsOf(s.rels).forEach(function (r) {
        if (!/\/notesSlide$/.test(r.type)) return;
        var np = resolvePath(s.path, r.target);
        zip.remove(np); zip.remove(relsPath(np));
        ct = ct.replace(new RegExp('<Override PartName="/' + np.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '"[^>]*/>'), '');
      });
      zip.remove(s.path); zip.remove(relsPath(s.path));
      ct = ct.replace(new RegExp('<Override PartName="/' + s.path.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '"[^>]*/>'), '');
      prels = prels.replace(new RegExp('<Relationship\\s[^>]*Id="' + s.rid + '"[^>]*/>'), '');
      pres = pres.replace(new RegExp('<p:sld r:id="' + s.rid + '"\\s*/>', 'g'), '');
    });
    void jobs;
    // 넣기
    added.forEach(function (a) {
      zip.file(a.path, a.xml);
      zip.file(relsPath(a.path), relsXml);
      ct = ct.replace('</Types>', '<Override PartName="/' + a.path + '" ContentType="' + CT_SLIDE + '"/></Types>');
      prels = prels.replace('</Relationships>', '<Relationship Id="' + a.rid + '" Type="' + REL + 'slide" Target="slides/' + a.path.split('/').pop() + '"/></Relationships>');
    });
    // 순서 — 첫 차이 슬라이드 자리에 새 슬라이드들, 나머지 차이 슬라이드는 뺌
    var newIds = added.map(function (a) { return '<p:sldId id="' + a.id + '" r:id="' + a.rid + '"/>'; }).join('');
    removed.forEach(function (s) {
      var re = new RegExp('<p:sldId id="' + s.id + '" r:id="' + s.rid + '"\\s*/>');
      pres = pres.replace(re, s.id === anchorId ? newIds : '');
      var secRe = new RegExp('<p14:sldId id="' + s.id + '"\\s*/>');
      pres = pres.replace(secRe, s.id === anchorId ? added.map(function (a) { return '<p14:sldId id="' + a.id + '"/>'; }).join('') : '');
    });
    zip.file('ppt/presentation.xml', pres);
    zip.file('ppt/_rels/presentation.xml.rels', prels);
    zip.file('[Content_Types].xml', ct);
    var count = pkg.slides.length - removed.length + added.length;
    var app = zip.file('docProps/app.xml');
    return (app ? app.async('string') : Promise.resolve(null)).then(function (ax) {
      if (ax) zip.file('docProps/app.xml', ax.replace(/<Slides>\d+<\/Slides>/, '<Slides>' + count + '</Slides>'));
      return count;
    });
  }
  function load(JSZip, buf) { return buf ? JSZip.loadAsync(buf) : Promise.resolve(defaultTemplate(JSZip)); }
  function inspect(JSZip, buf) { return load(JSZip, buf).then(readPkg).then(inspectPkg); }
  function make(JSZip, buf, an, db, opts) { return load(JSZip, buf).then(function (zip) { return build(JSZip, zip, an, db, opts); }); }

  /* ── 12. 기본 양식 (직전 보고서가 없을 때 · 검사용 샘플) ─────────────────── */
  var NAVY = '1F3864', GRID = '8C8C8C', GRAY = 'D9D9D9', LIGHT = 'EEF2F7';
  var FONT = '<a:latin typeface="Malgun Gothic" panose="020B0503020000020004" pitchFamily="34" charset="-127"/><a:ea typeface="Malgun Gothic" panose="020B0503020000020004" pitchFamily="34" charset="-127"/><a:cs typeface="Malgun Gothic"/>';
  function rPrX(o) { return '<a:rPr lang="ko-KR" altLang="en-US" sz="' + (o.sz || 1100) + '"' + (o.b ? ' b="1"' : '') + '><a:solidFill><a:srgbClr val="' + (o.color || '1A1A1A') + '"/></a:solidFill>' + FONT + '</a:rPr>'; }
  function pX(text, o) {
    o = o || {};
    var pPr = '<a:pPr algn="' + (o.algn || 'ctr') + '"' + (o.bullet ? ' marL="228600" indent="-228600"><a:buFont typeface="Malgun Gothic"/><a:buChar char="' + o.bullet + '"/></a:pPr>' : '><a:buNone/></a:pPr>');
    var r = rPrX(o);
    return String(text).split('\n').map(function (ln) { return '<a:p>' + pPr + (ln ? '<a:r>' + r + '<a:t>' + esc(ln) + '</a:t></a:r>' : '') + r.replace('<a:rPr', '<a:endParaRPr').replace('</a:rPr>', '</a:endParaRPr>') + '</a:p>'; }).join('');
  }
  function lnX(t, w, c) { return '<a:' + t + ' w="' + w + '"><a:solidFill><a:srgbClr val="' + c + '"/></a:solidFill></a:' + t + '>'; }
  function tcX(text, o) {
    o = o || {};
    return '<a:tc' + (o.attrs || '') + '><a:txBody><a:bodyPr/><a:lstStyle/>' + pX(text, o) + '</a:txBody><a:tcPr marL="36000" marR="36000" marT="18000" marB="18000" anchor="ctr">' +
      lnX('lnL', 6350, GRID) + lnX('lnR', 6350, GRID) + lnX('lnT', 6350, GRID) + lnX('lnB', 6350, GRID) +
      (o.fill ? '<a:solidFill><a:srgbClr val="' + o.fill + '"/></a:solidFill>' : '<a:noFill/>') + '</a:tcPr></a:tc>';
  }
  function tblX(id, name, x, y, widths, rows) {
    var cy = rows.reduce(function (s, r) { return s + r.h; }, 0), cx = widths.reduce(function (s, w) { return s + w; }, 0);
    return '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="' + id + '" name="' + name + '"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>' +
      '<p:xfrm><a:off x="' + x + '" y="' + y + '"/><a:ext cx="' + cx + '" cy="' + cy + '"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1"/><a:tblGrid>' +
      widths.map(function (w) { return '<a:gridCol w="' + w + '"/>'; }).join('') + '</a:tblGrid>' +
      rows.map(function (r) { return '<a:tr h="' + r.h + '">' + r.cells.join('') + '</a:tr>'; }).join('') + '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>';
  }
  function spX(id, name, x, y, cx, cy, paras, o) {
    o = o || {};
    return '<p:sp><p:nvSpPr><p:cNvPr id="' + id + '" name="' + name + '"/><p:cNvSpPr' + (o.tx ? ' txBox="1"' : '') + '/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="' + x + '" y="' + y + '"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
      (o.fill ? '<a:solidFill><a:srgbClr val="' + o.fill + '"/></a:solidFill>' : '<a:noFill/>') + (o.line ? '<a:ln w="12700"><a:solidFill><a:srgbClr val="' + o.line + '"/></a:solidFill></a:ln>' : '<a:ln><a:noFill/></a:ln>') +
      '</p:spPr><p:txBody><a:bodyPr wrap="square" lIns="72000" tIns="36000" rIns="72000" bIns="36000" anchor="' + (o.anchor || 'ctr') + '"><a:normAutofit/></a:bodyPr><a:lstStyle/>' + paras + '</p:txBody></p:sp>';
  }
  function lineX(id, x, y, cx, o) {
    o = o || {};
    return '<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="' + id + '" name="연결선 ' + id + '"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="' + x + '" y="' + y + '"/><a:ext cx="' + cx + '" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="' + (o.w || 12700) + '"><a:solidFill><a:srgbClr val="' + (o.color || NAVY) + '"/></a:solidFill>' + (o.dash ? '<a:prstDash val="dash"/>' : '') + '</a:ln></p:spPr></p:cxnSp>';
  }
  function slideX(shapes) {
    return XML_HEAD + '<p:sld ' + NS + '><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
      shapes.join('') + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
  }
  function headerX(o) {
    return [spX(2, '제목 1', 380000, 180000, 11400000, 420000, pX('재고실사 결과 보고', { sz: 2000, b: true, color: NAVY, algn: 'l' }), { tx: true, anchor: 'b' }),
      lineX(3, 300000, 660000, 11600000, { w: 19050 }),
      spX(4, '회사', 9000000, 6450000, 2900000, 260000, pX(o.company || '', { sz: 900, color: '7F7F7F', algn: 'r' }), { tx: true })];
  }
  function secBox(id, text, x, y, cx, cy) { return spX(id, '구역 ' + id, x, y, cx, cy, pX(text, { sz: 1200, b: true }), { line: '1A1A1A' }); }
  function memo(id, text, x, y, cx, cy) { return spX(id, '메모 ' + id, x, y, cx, cy, pX(text, { sz: 1000, b: true, color: 'C00000', algn: 'l' }), { tx: true, anchor: 't' }); }
  /* o: { company, statusLabels(실사 현황 표의 행 이름), memos(작성 안내 상자), extraDiffSlide(차이 슬라이드 2장), placeholder(개요 칸 자리표시) } */
  function defaultTemplate(JSZip, o) {
    o = o || {};
    var company = o.company == null ? '' : o.company;
    var H = function (t, a) { return tcX(t, { b: true, color: 'FFFFFF', fill: NAVY, sz: 1100, attrs: a }); };
    var C = function (t, a, f) { return tcX(t, { sz: 1100, attrs: a, fill: f, algn: 'ctr' }); };
    // 1장 — 개요 · 실사 현황
    var ov = tblX(10, '개요 표', 1670000, 900000, [2200000, 3160000], [
      { h: 260000, cells: [H('구분'), H('내용')] },
      { h: 260000, cells: [C('실사 시행일'), C(o.placeholder ? 'YY.MM.DD  HH:MM' : '')] },
      { h: 260000, cells: [C('집계 기준 시점'), C(o.placeholder ? 'YY.MM.DD  HH:MM' : '')] },
      { h: 260000, cells: [C('실사방법', ' gridSpan="2"', LIGHT), C('', ' hMerge="1"', LIGHT)] },
      { h: 260000, cells: [C((company ? company + ' ' : '') + '사내 및 사외창고'), C('전수 재고 실사')] },
      { h: 260000, cells: [C('협력사재고', ' rowSpan="2"'), C('(직접 입력함)')] },
      { h: 260000, cells: [C('', ' vMerge="1"'), C('(직접 입력함)')] }
    ]);
    var labels = o.statusLabels && o.statusLabels.length ? o.statusLabels : [''];
    var num = function () { return C(''); };
    var st = tblX(20, '실사 현황 표', 1670000, 3250000, [1150000, 850000, 900000, 850000, 650000, 900000, 650000, 850000, 900000], [
      { h: 250000, cells: [H('창고위치', ' rowSpan="2"'), H('전체재고', ' gridSpan="2"'), H('', ' hMerge="1"'), H('실사완료', ' gridSpan="4"'), H('', ' hMerge="1"'), H('', ' hMerge="1"'), H('', ' hMerge="1"'), H('미실시', ' gridSpan="2"'), H('', ' hMerge="1"')] },
      { h: 250000, cells: [H('', ' vMerge="1"'), H('수량'), H('금액'), H('수량'), H('%'), H('금액'), H('%'), H('수량'), H('금액')] }
    ].concat(labels.map(function (lb) { return { h: 240000, cells: [C(lb)].concat([1, 2, 3, 4, 5, 6, 7, 8].map(num)) }; }))
      .concat([{ h: 240000, cells: [C('합계', '', GRAY)].concat([1, 2, 3, 4, 5, 6, 7, 8].map(function () { return C('', '', GRAY); })) }]));
    var s1 = headerX({ company: company }).concat([
      secBox(5, '재고 실사\n개요', 300000, 860000, 1000000, 600000), ov,
      lineX(6, 300000, 2950000, 11600000, { w: 6350, color: '7F7F7F' }),
      secBox(7, '재고 실사\n현황', 300000, 3060000, 1000000, 600000),
      spX(8, '단위', 7900000, 2990000, 1650000, 240000, pX('(단위 : 천 EA, 백만원)', { sz: 900, algn: 'r' }), { tx: true }), st
    ]).concat(o.memos ? [memo(9, '(계산로직)\n- 전체재고 = 작업지시에 올린 ERP 수량 · 금액의 합\n- 실사인원 0 인 창고는 0 이 아니라 - 로 표시', 1670000, 5750000, 7700000, 650000)] : []);
    // 2장 — 차이 내용
    function diffSlide(title) {
      var dt = tblX(10, '차이 표', 300000, 1840000, [1016000, 2832000, 725000, 725000, 725000, 725000, 729000, 944000, 2909000], [
        { h: 400000, cells: ['자재코드', '자재내역', '저장위치', '전산수량', '실물수량', '차이수량', '단가', '차이금액', '차이 사유 및 조치사항'].map(function (t) { return H(t); }) },
        { h: 365000, cells: [0, 1, 2, 3, 4, 5, 6, 7, 8].map(function () { return C(''); }) },
        { h: 330000, cells: [C('합계', '', GRAY)].concat([1, 2, 3, 4, 5, 6, 7, 8].map(function () { return C('', '', GRAY); })) }
      ]);
      return headerX({ company: company }).concat([
        spX(5, '차이 제목', 300000, 840000, 3600000, 540000, pX(title, { sz: 1300, b: true }), { line: '1A1A1A' }),
        spX(6, '단위', 10000000, 1600000, 1630000, 220000, pX('(단위 : EA, 원)', { sz: 900, algn: 'r' }), { tx: true }), dt
      ]).concat(o.memos ? [memo(7, '(계산로직) 차이수량 · 차이금액이 0 이 아니면 빨간 글자, 합계는 차이수량 · 차이금액만', 1490000, 6330000, 7700000, 300000)] : []);
    }
    // 3장 — 이슈사항 및 개선
    var s3 = headerX({ company: company }).concat([
      secBox(5, '이슈사항 및 개선', 300000, 840000, 1830000, 600000),
      spX(6, '이슈 머리', 2200000, 1830000, 4305000, 427000, pX('이슈 사항', { sz: 1200, b: true, color: 'FFFFFF' }), { fill: NAVY }),
      spX(7, '개선 머리', 7083000, 1830000, 4305000, 427000, pX('개선 및 협의사항', { sz: 1200, b: true, color: 'FFFFFF' }), { fill: NAVY })
    ]);
    [0, 1].forEach(function (i) {
      var y = 2380000 + i * 1160000;
      s3.push(spX(10 + i, '주제 ' + (i + 1), 800000, y, 1324000, 903000, pX('…', { sz: 1100 }), { fill: GRAY }));
      s3.push(spX(20 + i, '이슈 ' + (i + 1), 2200000, y + 60000, 4305000, 840000, pX('…', { sz: 1100, algn: 'l', bullet: '•' }), { tx: true, anchor: 't' }));
      s3.push(spX(30 + i, '개선 ' + (i + 1), 7083000, y + 60000, 4305000, 840000, pX('…', { sz: 1100, algn: 'l', bullet: '✓' }), { tx: true, anchor: 't' }));
      s3.push(lineX(40 + i, 650000, y + 1000000, 10740000, { w: 9525, color: '404040', dash: true }));
    });
    if (o.memos) s3.push(memo(50, 'AI가 초안 작성해주면, 담당자가 수정하여 활용', 1870000, 4900000, 7700000, 260000));
    var slides = [slideX(s1), slideX(diffSlide('차이 내용 : 창고명'))];
    if (o.extraDiffSlide) slides.push(slideX(diffSlide('차이 내용 : 창고명2')));
    slides.push(slideX(s3));
    return pkgOf(JSZip, slides, company);
  }
  function pkgOf(JSZip, slides, company) {
    var z = new JSZip();
    var n = slides.length;
    z.file('[Content_Types].xml', XML_HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' +
      '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>' +
      '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>' +
      '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
      '<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>' +
      '<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>' +
      '<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>' +
      slides.map(function (_, i) { return '<Override PartName="/ppt/slides/slide' + (i + 1) + '.xml" ContentType="' + CT_SLIDE + '"/>'; }).join('') +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>');
    z.file('_rels/.rels', XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + REL + 'officeDocument" Target="ppt/presentation.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="' + REL + 'extended-properties" Target="docProps/app.xml"/></Relationships>');
    z.file('docProps/core.xml', XML_HEAD + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>재고실사 결과 보고</dc:title><dc:creator>' + esc(company || '재고실사 Start-End') + '</dc:creator></cp:coreProperties>');
    z.file('docProps/app.xml', XML_HEAD + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>재고실사 Start-End (data10-14)</Application><Slides>' + n + '</Slides></Properties>');
    z.file('ppt/presentation.xml', XML_HEAD + '<p:presentation ' + NS + ' saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>' +
      slides.map(function (_, i) { return '<p:sldId id="' + (256 + i) + '" r:id="rId' + (10 + i) + '"/>'; }).join('') +
      '</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle><a:lvl1pPr><a:defRPr lang="ko-KR" sz="1800"><a:latin typeface="Malgun Gothic"/><a:ea typeface="Malgun Gothic"/></a:defRPr></a:lvl1pPr></p:defaultTextStyle></p:presentation>');
    z.file('ppt/_rels/presentation.xml.rels', XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + REL + 'slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="' + REL + 'theme" Target="theme/theme1.xml"/><Relationship Id="rId3" Type="' + REL + 'presProps" Target="presProps.xml"/><Relationship Id="rId4" Type="' + REL + 'viewProps" Target="viewProps.xml"/><Relationship Id="rId5" Type="' + REL + 'tableStyles" Target="tableStyles.xml"/>' +
      slides.map(function (_, i) { return '<Relationship Id="rId' + (10 + i) + '" Type="' + REL + 'slide" Target="slides/slide' + (i + 1) + '.xml"/>'; }).join('') + '</Relationships>');
    z.file('ppt/presProps.xml', XML_HEAD + '<p:presentationPr ' + NS + '/>');
    z.file('ppt/viewProps.xml', XML_HEAD + '<p:viewPr ' + NS + '><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:gridSpacing cx="72008" cy="72008"/></p:viewPr>');
    z.file('ppt/tableStyles.xml', XML_HEAD + '<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>');
    var tree = '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>';
    z.file('ppt/slideMasters/slideMaster1.xml', XML_HEAD + '<p:sldMaster ' + NS + '><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>' + tree.slice('<p:cSld>'.length) +
      '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>' +
      '<p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="2800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/></a:defRPr></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr sz="1800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:defRPr></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:defRPr></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>');
    z.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + REL + 'slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="' + REL + 'theme" Target="../theme/theme1.xml"/></Relationships>');
    z.file('ppt/slideLayouts/slideLayout1.xml', XML_HEAD + '<p:sldLayout ' + NS + ' type="blank" preserve="1">' + tree.replace('<p:cSld>', '<p:cSld name="빈 화면">') + '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>');
    z.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + REL + 'slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>');
    z.file('ppt/theme/theme1.xml', themeX());
    slides.forEach(function (x, i) {
      z.file('ppt/slides/slide' + (i + 1) + '.xml', x);
      z.file('ppt/slides/_rels/slide' + (i + 1) + '.xml.rels', XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + REL + 'slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>');
    });
    return z;
  }
  function themeX() {
    var clr = function (n, v) { return '<a:' + n + '><a:srgbClr val="' + v + '"/></a:' + n + '>'; };
    var fill3 = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
    var ln = function (w) { return '<a:ln w="' + w + '"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'; };
    return XML_HEAD + '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="재고실사 기본"><a:themeElements><a:clrScheme name="재고실사">' +
      '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' + clr('dk2', NAVY) + clr('lt2', 'E7E6E6') +
      clr('accent1', '2F5597') + clr('accent2', 'C55A11') + clr('accent3', '7F7F7F') + clr('accent4', 'BF9000') + clr('accent5', '2E75B6') + clr('accent6', '548235') + clr('hlink', '0563C1') + clr('folHlink', '954F72') +
      '</a:clrScheme><a:fontScheme name="재고실사"><a:majorFont><a:latin typeface="Malgun Gothic"/><a:ea typeface="Malgun Gothic"/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Malgun Gothic"/><a:ea typeface="Malgun Gothic"/><a:cs typeface=""/></a:minorFont></a:fontScheme>' +
      '<a:fmtScheme name="재고실사"><a:fillStyleLst>' + fill3 + '</a:fillStyleLst><a:lnStyleLst>' + ln(6350) + ln(12700) + ln(19050) + '</a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst>' + fill3 + '</a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>';
  }

  /* 검사 · 화면용: 슬라이드 순서대로 표 · 글자 꺼내기 */
  function dump(JSZip, buf) {
    return JSZip.loadAsync(buf).then(readPkg).then(function (pkg) {
      return pkg.slides.map(function (s) {
        return { path: s.path, id: s.id, rels: s.rels, xml: s.xml, shapes: shapesOf(s.xml).map(function (x) { return x.text; }), tables: tablesOf(s.xml).map(function (t) { return { kind: tableKind(t.T), rows: t.T.rows.map(function (r) { return r.cells.map(function (c) { return c.text; }); }), red: t.T.rows.map(function (r) { return r.cells.map(function (c) { return c.xml.indexOf('<a:srgbClr val="' + RED + '"/>') >= 0 && /<a:t>/.test(c.xml); }); }), h: t.T.rows.map(function (r) { return r.h; }), xf: t.T.xf }; }) };
      });
    });
  }

  return {
    RED: RED, PPTX_MIME: PPTX_MIME, STATUS_DIMS: STATUS_DIMS,
    esc: esc, elems: elems, textOf: textOf, setText: setText, setColor: setColor, parseTable: parseTable, parseUnit: parseUnit, fmtScaled: fmtScaled, fitText: fitText,
    diffFieldOf: diffFieldOf, statusLayout: statusLayout, statusGroups: statusGroups, diffGroups: diffGroups, reportData: reportData, issueDraft: issueDraft, parseIssueText: parseIssueText, issueTextOf: issueTextOf,
    readPkg: readPkg, inspectPkg: inspectPkg, inspect: inspect, make: make, build: build, defaultTemplate: defaultTemplate, dump: dump
  };
});
