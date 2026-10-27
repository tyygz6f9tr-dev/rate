/*
 * schema-check.js
 * 올린 엑셀의 머리글(컬럼명)을 저장소의 schema.json(전달 기준)과 비교해
 * 추가·삭제·변경된 컬럼을 파일/시트/열 위치별로 알려준다.
 * 매달 자연스럽게 바뀌는 부분은 변경으로 보지 않는다:
 *   - '비교(2025-09)' → '비교(2025-10)'  (YYYY-MM)
 *   - 수금 파일의 월 블록 추가('10월'), '9월 누계' → '10월 누계', '1~9' → '1~10'
 */
(function (global) {
  'use strict';

  const str = v => (v === null || v === undefined) ? '' : String(v).replace(/\s+/g, ' ').trim();
  const colLetter = c => { let s = ''; c += 1; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; };
  const colIndex = l => { let n = 0; for (const ch of l) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; };

  // 매달 바뀌는 숫자 표기를 일반화
  function norm(label) {
    return str(label)
      .replace(/\d{4}-\d{2}~\d{4}-\d{2}/g, 'YYYY-MM~YYYY-MM')
      .replace(/\d{4}-\d{2}/g, 'YYYY-MM')
      .replace(/^\d{1,2}월 ?누계$/, 'N월 누계')
      .replace(/^1~\d{1,2}$/, '1~N')
      .replace(/^\d{1,2}월$/, 'N월');
  }
  function normSheet(name) {
    return str(name).replace(/\d{4}/g, 'YYYY').replace(/1~\d{1,2}/g, '1~N');
  }

  function textCount(sh, r) { return sh.row(r).filter(v => typeof v === 'string' && str(v)).length; }
  function findHeaderRow(sh, key, maxRow = 12) {
    for (let r = 0; r < Math.min(maxRow, sh.rowCount); r++) if (sh.row(r).some(v => str(v) === key)) return r;
    let best = -1, bestN = 0;
    for (let r = 0; r < Math.min(maxRow, sh.rowCount); r++) { const n = textCount(sh, r); if (n > bestN) { bestN = n; best = r; } }
    return best;
  }
  function headersOf(sh, r, from = 0, to = Infinity) {
    const out = [];
    const row = sh.row(r);
    for (let c = from; c < Math.min(row.length, to); c++) {
      const t = str(row[c]);
      if (t) out.push({ col: colLetter(c), label: norm(t) });
    }
    return out;
  }

  // ---------- 머리글 추출 ----------
  async function extractSales(wb) {
    const months = wb.sheetNames.filter(n => /^\d{1,2}월$/.test(n.trim()));
    const sheets = {};
    for (const n of months) {
      const sh = await wb.readSheet(n);
      const r = findHeaderRow(sh, '매장코드');
      sheets[n] = { headerRow: r + 1, headers: r >= 0 ? headersOf(sh, r) : [] };
    }
    const first = months[0];
    return {
      sheetNames: wb.sheetNames.map(normSheet),
      monthSheets: months,
      headerRow: first ? sheets[first].headerRow : null,
      headers: first ? sheets[first].headers : [],
      perSheet: sheets
    };
  }

  async function extractCollection(wb, sheetName) {
    const sh = await wb.readSheet(sheetName);
    let r3 = findHeaderRow(sh, '유통수수료율', 10);
    if (r3 < 1) r3 = 1;
    const r2 = r3 - 1;
    const top = sh.row(r2), sub = sh.row(r3);
    const width = Math.max(top.length, sub.length);
    const monthStarts = [];
    for (let c = 0; c < width; c++) if (/^\d{1,2}월$/.test(str(top[c]))) monthStarts.push({ c, label: str(top[c]) });
    let cumStart = width;
    for (let c = 0; c < width; c++) {
      if (/^1~\d+$/.test(str(top[c])) || /누계/.test(str(sub[c]))) { cumStart = c; break; }
    }
    const firstBlock = monthStarts.length ? monthStarts[0].c : cumStart;
    const label = c => str(sub[c]) || str(top[c]);
    const id = [];
    for (let c = 0; c < firstBlock; c++) { const t = label(c); if (t) id.push({ col: colLetter(c), label: norm(t) }); }
    const blocks = {};
    monthStarts.forEach((m, i) => {
      const end = Math.min(monthStarts[i + 1] ? monthStarts[i + 1].c : Infinity, cumStart > m.c ? cumStart : Infinity, width);
      const list = [];
      for (let c = m.c; c < end; c++) { const t = str(sub[c]); if (t) list.push({ col: colLetter(c), label: norm(t) }); }
      blocks[m.label] = list;
    });
    const cum = [];
    for (let c = cumStart; c < width; c++) {
      const t = label(c);
      if (t) cum.push({ col: colLetter(c), label: norm(t), top: norm(str(top[c])) || undefined });
    }
    const blockList = Object.values(blocks);
    return {
      sheet: normSheet(sheetName),
      sheetName,
      headerRows: [r2 + 1, r3 + 1],
      id,
      block: blockList.length ? blockList[0].map(h => h.label) : [],
      blocks,
      cum
    };
  }

  async function extractPrev(wb) {
    const name = wb.sheetNames.find(n => n.trim() === '종합') || wb.sheetNames[0];
    const sh = await wb.readSheet(name);
    const r = findHeaderRow(sh, '매장코드');
    return { sheet: name, headerRow: r + 1, headers: r >= 0 ? headersOf(sh, r) : [] };
  }

  // ---------- 비교 (LCS 정렬) ----------
  function diffLists(oldL, newL) {
    const a = oldL.map(h => h.label), b = newL.map(h => h.label);
    const n = a.length, m = b.length;
    const dp = Array.from({ length: n + 1 }, () => new Array(m + 1).fill(0));
    for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--)
      dp[i][j] = a[i] === b[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
    // LCS 역추적 → '=', '-', '+' 순서
    const seq = [];
    let i = 0, j = 0;
    while (i < n && j < m) {
      if (a[i] === b[j]) { seq.push({ t: '=' }); i++; j++; }
      else if (dp[i + 1][j] >= dp[i][j + 1]) { seq.push({ t: '-', h: oldL[i] }); i++; }
      else { seq.push({ t: '+', h: newL[j] }); j++; }
    }
    while (i < n) { seq.push({ t: '-', h: oldL[i] }); i++; }
    while (j < m) { seq.push({ t: '+', h: newL[j] }); j++; }
    // '=' 사이 구간마다 삭제·추가를 짝지어 '변경'으로
    const ops = [];
    let rem = [], add = [];
    const flush = () => {
      const k = Math.min(rem.length, add.length);
      for (let x = 0; x < k; x++) ops.push({ type: '변경', oldCol: rem[x].col, newCol: add[x].col, old: rem[x].label, now: add[x].label });
      for (let x = k; x < rem.length; x++) ops.push({ type: '삭제', oldCol: rem[x].col, newCol: '', old: rem[x].label, now: '' });
      for (let x = k; x < add.length; x++) ops.push({ type: '추가', oldCol: '', newCol: add[x].col, old: '', now: add[x].label });
      rem = []; add = [];
    };
    for (const s of seq) {
      if (s.t === '=') flush();
      else if (s.t === '-') rem.push(s.h);
      else add.push(s.h);
    }
    flush();
    return ops;
  }
  const asList = arr => arr.map((label, k) => ({ col: '', label, k }));

  // 계산에 꼭 필요한 컬럼
  const REQUIRED = {
    sales: ['매장명', '매장코드', '실판금액', '비교(YYYY-MM)'],
    collection: ['코드', '매장명', '계열', '매출', '수금', '유통수수료율', 'N월 누계'],
    prev: ['매장코드', '매출', '%']
  };

  function compare(base, cur) {
    const rows = [];
    const push = (file, sheet, area, ops, reqKey) => ops.forEach(o => rows.push({
      file, sheet, area, ...o,
      required: REQUIRED[reqKey].includes(o.old) || REQUIRED[reqKey].includes(o.now)
    }));
    const notes = [];
    if (cur.sales) {
      if (!base || !base.sales) notes.push('판매현황: 비교할 전달 기준이 없어요.');
      else {
        push('월별 판매현황', '월 시트(1월~)', '머리글', diffLists(base.sales.headers, cur.sales.headers), 'sales');
        if (base.sales.headerRow !== cur.sales.headerRow) notes.push(`판매현황: 머리글 행 위치가 ${base.sales.headerRow}행 → ${cur.sales.headerRow}행으로 바뀌었어요.`);
        // 월 시트끼리 머리글이 다른지
        const ref = cur.sales.headers.map(h => h.label).join('|');
        for (const [sn, s] of Object.entries(cur.sales.perSheet)) {
          if (s.headers.map(h => h.label).join('|') !== ref)
            push('월별 판매현황', sn, '머리글 (다른 월 시트와 다름)', diffLists(cur.sales.headers, s.headers), 'sales');
        }
      }
    }
    if (cur.collection) {
      const b = base && base.collection;
      if (!b) notes.push('수금 파일: 비교할 전달 기준이 없어요.');
      else {
        const sn = cur.collection.sheetName;
        if (b.sheet !== cur.collection.sheet) notes.push(`수금 파일: 시트 이름 형식이 '${b.sheet}' → '${cur.collection.sheet}'로 달라요.`);
        push('수금 파일', sn, '매장 정보 열', diffLists(b.id, cur.collection.id), 'collection');
        for (const [m, list] of Object.entries(cur.collection.blocks)) {
          const ops = diffLists(asList(b.block), list);
          ops.forEach(o => { if (!o.oldCol) o.oldCol = ''; });
          push('수금 파일', sn, `${m} 블록`, ops, 'collection');
        }
        push('수금 파일', sn, '누계 열', diffLists(b.cum, cur.collection.cum), 'collection');
      }
    }
    if (cur.prev) {
      const b = base && base.prev;
      if (!b) notes.push('전년실적 파일: 비교할 전달 기준이 없어요.');
      else {
        if (b.sheet !== cur.prev.sheet) notes.push(`전년실적 파일: 시트 이름이 '${b.sheet}' → '${cur.prev.sheet}'로 달라요.`);
        push('전년실적', cur.prev.sheet, '머리글', diffLists(b.headers, cur.prev.headers), 'prev');
      }
    }
    return { rows, notes };
  }

  function makeSchema(cur, files) {
    const out = { version: 1, saved: new Date().toISOString().slice(0, 10), files: files || {} };
    if (cur.sales) out.sales = { headerRow: cur.sales.headerRow, headers: cur.sales.headers };
    if (cur.collection) out.collection = { sheet: cur.collection.sheet, headerRows: cur.collection.headerRows, id: cur.collection.id, block: cur.collection.block, cum: cur.collection.cum };
    if (cur.prev) out.prev = { sheet: cur.prev.sheet, headerRow: cur.prev.headerRow, headers: cur.prev.headers };
    return out;
  }

  global.SchemaCheck = { extractSales, extractCollection, extractPrev, compare, makeSchema };
})(window);
