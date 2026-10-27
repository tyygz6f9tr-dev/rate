/*
 * 임대매장 유통수수료율 월별 현황 생성기
 * - 매출·신장률: 월별 판매현황 파일 (시트 '1월' ~ 'N월', 실판금액 / 비교(전년) 열)
 * - 수수료율: 수금 파일의 '1~N누계' 시트 (월별 유통수수료율, N월 누계 수수료율)
 * - 전년실적(선택): 매출·수수료 RAW 파일 '종합' 시트 (매장코드 / 매출 / %)
 * 규칙: 매장명 앞에 (행)이 붙은 매장은 매장코드 첫 글자(S/T/A)가 같은 브랜드 안에서
 *       (행)을 뺀 이름이 같은 본 매장에 합산하고, 매장코드는 본 매장 코드를 쓴다.
 */
(function () {
  'use strict';

  const BRANDS = { S: '쉬즈미스', T: '리스트', A: '시스티나' };
  const $ = id => document.getElementById(id);

  const str = v => (v === null || v === undefined) ? '' : String(v).trim();
  const num = v => (typeof v === 'number' && isFinite(v)) ? v : null;
  const isEvent = name => /^\(행\)/.test(str(name));
  const normName = name => str(name).replace(/^\(행\)/, '').replace(/\s+/g, '').replace(/점$/, '');
  const isCode = v => /^[A-Z]\d{4,}[A-Z0-9]*$/.test(str(v));

  function findRow(sheet, test, maxRow = 12) {
    for (let r = 0; r < Math.min(maxRow, sheet.rowCount); r++) {
      if (sheet.row(r).some(v => test(str(v)))) return r;
    }
    return -1;
  }
  function findCol(sheet, r, test, from = 0, to = 400) {
    const row = sheet.row(r);
    for (let c = from; c < Math.min(to, Math.max(row.length, from)); c++) if (test(str(row[c]))) return c;
    return -1;
  }

  function isGreen(rgb) {
    if (!rgb) return false;
    const [r, g, b] = rgb.map(v => v / 255);
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    if (max - min < 0.03) return false;
    let h;
    if (max === r) h = 60 * (((g - b) / (max - min)) % 6);
    else if (max === g) h = 60 * ((b - r) / (max - min) + 2);
    else h = 60 * ((r - g) / (max - min) + 4);
    if (h < 0) h += 360;
    return h >= 75 && h <= 170;
  }

  // ---------- 1. 월별 판매현황 ----------
  async function readSales(wb, log) {
    const months = {};
    let year = null;
    const missingCols = new Set();
    let monthSheets = 0;
    for (const name of wb.sheetNames) {
      const m = /^(\d{1,2})월$/.exec(name.trim());
      if (!m) continue;
      const month = parseInt(m[1], 10);
      monthSheets++;
      const sh = await wb.readSheet(name);
      if (!year) {
        for (let r = 0; r < 4; r++) {
          const y = /(\d{4})-\d{2}/.exec(sh.row(r).map(str).join(' '));
          if (y) { year = parseInt(y[1], 10); break; }
        }
      }
      const hr = findRow(sh, t => t === '매장코드');
      if (hr < 0) { log.warn(`판매현황 '${name}' 시트에서 '매장코드' 머리글을 찾지 못했어요.`); continue; }
      const cCode = findCol(sh, hr, t => t === '매장코드');
      const cName = findCol(sh, hr, t => t === '매장명');
      const cCur = findCol(sh, hr, t => t === '실판금액');
      const cPrev = findCol(sh, hr, t => t.startsWith('비교'));
      const missing = [['매장코드', cCode], ['매장명', cName], ['실판금액', cCur], ['비교(…)', cPrev]].filter(x => x[1] < 0).map(x => x[0]);
      if (missing.length) { missingCols.add(missing.join(', ')); log.warn(`판매현황 '${name}' 시트에서 ${missing.join(', ')} 열을 찾지 못했어요.`); continue; }
      const base = {}, events = [];
      for (let r = hr + 1; r < sh.rowCount; r++) {
        const code = str(sh.get(r, cCode));
        if (!isCode(code)) continue;
        const rec = { code, name: str(sh.get(r, cName)), f: num(sh.get(r, cCur)) || 0, g: num(sh.get(r, cPrev)) || 0 };
        if (isEvent(rec.name)) events.push(rec); else base[code] = rec;
      }
      months[month] = { base, events };
    }
    const list = Object.keys(months).map(Number).sort((a, b) => a - b);
    if (!list.length) {
      if (monthSheets) throw new Error(`판매현황 파일의 월 시트에서 필요한 열(${Array.from(missingCols).join(' / ')})을 찾지 못했어요. 컬럼명이 바뀌었는지 확인해주세요.`);
      throw new Error("판매현황 파일에서 '1월', '2월' 같은 월별 시트를 찾지 못했어요.");
    }
    return { months, monthList: list, year };
  }

  // 같은 브랜드(코드 첫 글자) + 같은 이름의 (행) 매장을 본 매장에 합산
  function mergedMonth(monthData, code) {
    const b = monthData.base[code];
    if (!b) return { f: 0, g: 0, events: [], name: null };
    const key = normName(b.name);
    const ev = monthData.events.filter(e => e.code[0] === code[0] && normName(e.name) === key);
    return {
      f: b.f + ev.reduce((s, e) => s + e.f, 0),
      g: b.g + ev.reduce((s, e) => s + e.g, 0),
      events: ev,
      name: b.name
    };
  }

  // ---------- 2. 수금 파일 ----------
  function pickCollectionSheet(names) {
    const cands = names.map(n => ({ n, m: /1~(\d+)\s*누계/.exec(n) })).filter(x => x.m);
    if (!cands.length) return names[0];
    cands.sort((a, b) =>
      (parseInt(b.m[1], 10) - parseInt(a.m[1], 10)) ||
      ((b.n.includes('최종') ? 1 : 0) - (a.n.includes('최종') ? 1 : 0)) ||
      (a.n.length - b.n.length));
    return cands[0].n;
  }

  async function readCollection(wb, sheetName, log) {
    const sh = await wb.readSheet(sheetName);
    const r3 = findRow(sh, t => t === '유통수수료율');
    if (r3 < 1) throw new Error(`수금 파일 '${sheetName}' 시트에서 '유통수수료율' 머리글을 찾지 못했어요.`);
    const r2 = r3 - 1;
    const colIn = (r, t) => findCol(sh, r, x => x === t);
    const cCode = colIn(r3, '코드') >= 0 ? colIn(r3, '코드') : colIn(r2, '코드');
    const cName = colIn(r3, '매장명') >= 0 ? colIn(r3, '매장명') : colIn(r2, '매장명');
    const cGroup = colIn(r3, '계열');
    const cMain = colIn(r3, '주요매장');
    if (cCode < 0) throw new Error(`수금 파일 '${sheetName}' 시트에서 '코드' 열을 찾지 못했어요.`);

    // 월 블록 (2번째 머리글 행의 '1월', '2월' ...)
    const starts = [];
    sh.row(r2).forEach((v, c) => { const m = /^(\d{1,2})월$/.exec(str(v)); if (m) starts.push({ month: parseInt(m[1], 10), c }); });
    const cCumLabel = findCol(sh, r2, t => /^1~\d+$/.test(t));
    const cCumRate = findCol(sh, r3, t => /누계/.test(t));
    const blockEnd = (i) => {
      const next = starts[i + 1] ? starts[i + 1].c : Infinity;
      const caps = [next, cCumRate >= 0 ? cCumRate : Infinity, cCumLabel >= 0 ? cCumLabel : Infinity].filter(x => x > starts[i].c);
      return Math.min(...caps);
    };
    const blocks = {};
    starts.sort((a, b) => a.c - b.c).forEach((s, i) => {
      const end = blockEnd(i);
      blocks[s.month] = {
        sales: findCol(sh, r3, t => t === '매출', s.c, end),
        coll: findCol(sh, r3, t => t === '수금', s.c, end),
        rate: findCol(sh, r3, t => t === '유통수수료율', s.c, end)
      };
    });
    const cumFrom = cCumLabel >= 0 ? cCumLabel : (cCumRate >= 0 ? cCumRate : 0);
    const cCumSales = findCol(sh, r3, t => t === '매출', cumFrom);
    const cCumColl = findCol(sh, r3, t => t === '수금', cumFrom);

    const stores = [];
    for (let r = r3 + 1; r < sh.rowCount; r++) {
      const code = str(sh.get(r, cCode));
      if (!isCode(code)) continue;
      const rec = {
        code, row: r,
        name: cName >= 0 ? str(sh.get(r, cName)).replace(/\s*\([^)]*\)\s*$/, '') : code,
        group: cGroup >= 0 ? str(sh.get(r, cGroup)) : '',
        mainMark: cMain >= 0 ? str(sh.get(r, cMain)) : '',
        green: isGreen(sh.fill(r, cCode)) || (cName >= 0 && isGreen(sh.fill(r, cName))),
        months: {},
        cumRate: cCumRate >= 0 ? num(sh.get(r, cCumRate)) : null,
        cumSales: cCumSales >= 0 ? num(sh.get(r, cCumSales)) : null,
        cumColl: cCumColl >= 0 ? num(sh.get(r, cCumColl)) : null
      };
      for (const [m, b] of Object.entries(blocks)) {
        rec.months[m] = {
          sales: b.sales >= 0 ? num(sh.get(r, b.sales)) : null,
          coll: b.coll >= 0 ? num(sh.get(r, b.coll)) : null,
          rate: b.rate >= 0 ? num(sh.get(r, b.rate)) : null
        };
      }
      stores.push(rec);
    }
    if (cCumRate < 0) log.warn('수금 파일에서 누계 수수료율 열(예: "9월 누계")을 찾지 못해, 월별 매출·수금 합계로 계산했어요.');

    // 주요매장: '주요매장' 열 표시 → 없으면 초록색 칠
    let main = [], how = '';
    const byMark = stores.filter(s => s.mainMark);
    const byGreen = stores.filter(s => s.green);
    if (byMark.length) { main = byMark; how = "'주요매장' 열에 표시된 매장"; }
    else if (byGreen.length) { main = byGreen; how = '코드 칸이 초록색으로 칠해진 매장'; }
    else {
      main = stores.filter(s => /^[STA]4/.test(s.code));
      how = '주요매장 표시를 찾지 못해 쇼핑몰(코드 두 번째 자리 4) 매장 전체';
      log.warn('수금 파일에서 주요매장 표시(초록색 칠 또는 주요매장 열)를 찾지 못했어요. 쇼핑몰 매장 전체로 만들었어요.');
    }
    // 중복 코드 제거
    const seen = new Set();
    main = main.filter(s => !seen.has(s.code) && seen.add(s.code));
    return { stores, main, how, months: Object.keys(blocks).map(Number).sort((a, b) => a - b) };
  }

  // ---------- 3. 전년실적 (선택) ----------
  async function readPrev(wb, log) {
    if (!wb) return null;
    const sheetName = wb.sheetNames.find(n => n.trim() === '종합') || wb.sheetNames[0];
    const sh = await wb.readSheet(sheetName);
    const hr = findRow(sh, t => t === '매장코드');
    if (hr < 0) { log.warn(`전년실적 파일 '${sheetName}' 시트에서 '매장코드' 머리글을 찾지 못해 전년실적을 비워둘게요.`); return null; }
    const cCode = findCol(sh, hr, t => t === '매장코드');
    const cSales = findCol(sh, hr, t => t === '매출');
    const cRate = findCol(sh, hr, t => t === '%' || t === '수수료율');
    const out = {};
    for (let r = hr + 1; r < sh.rowCount; r++) {
      const code = str(sh.get(r, cCode));
      if (!isCode(code) || out[code]) continue;
      out[code] = { sales: cSales >= 0 ? num(sh.get(r, cSales)) : null, rate: cRate >= 0 ? num(sh.get(r, cRate)) : null };
    }
    return out;
  }

  // ---------- 계산 ----------
  function build(sales, coll, prev, log) {
    const N = sales.monthList[sales.monthList.length - 1];
    const months = Array.from({ length: N }, (_, i) => i + 1);
    const missingCollMonths = months.filter(m => !coll.months.includes(m));
    if (missingCollMonths.length) log.warn(`수금 파일에 ${missingCollMonths.join(', ')}월 수수료율 열이 없어요. 해당 월 수수료율은 '-'로 표시돼요.`);
    const missingSalesMonths = months.filter(m => !sales.months[m]);
    if (missingSalesMonths.length) log.warn(`판매현황 파일에 ${missingSalesMonths.join(', ')}월 시트가 없어요.`);

    const merges = [];
    const data = {};
    const groupOrder = [];
    for (const s of coll.main) {
      const letter = s.code[0];
      if (!BRANDS[letter]) { log.warn(`${s.code} ${s.name}: S·T·A로 시작하지 않아 제외했어요.`); continue; }
      if (s.group && !groupOrder.includes(s.group)) groupOrder.push(s.group);
      const st = {
        code: s.code, name: s.name, open_date: null, close_date: null, contract_margin: null,
        sales: [], growth: [], rates: [], total_sales: 0, total_growth: null,
        total_rate: s.cumRate, total_o: s.cumSales, total_p: s.cumColl,
        group: s.group, prev_sales: null, prev_rate: null
      };
      let F = 0, G = 0, found = false;
      const evCodes = new Set();
      for (const m of months) {
        const md = sales.months[m];
        const mm = md ? mergedMonth(md, s.code) : { f: 0, g: 0, events: [], name: null };
        if (mm.name) { found = true; if (!st.name) st.name = mm.name; }
        mm.events.forEach(e => evCodes.add(`${e.code} ${e.name}`));
        st.sales.push(mm.f ? mm.f / 1e6 : null);
        st.growth.push(mm.f && mm.g ? mm.f / mm.g - 1 : null);
        F += mm.f; G += mm.g;
        const cm = s.months[m];
        st.rates.push(cm && cm.sales ? cm.rate : null);
      }
      if (!found) log.warn(`${s.code} ${s.name}: 판매현황 파일에 없는 매장이에요(매출이 비어 보여요).`);
      if (evCodes.size) merges.push({ code: s.code, name: st.name, events: Array.from(evCodes) });
      st.total_sales = F / 1e6;
      st.total_growth = G ? F / G - 1 : null;
      if (st.total_o === null || st.total_p === null) {
        let so = 0, sp = 0;
        for (const m of months) { const cm = s.months[m]; if (cm && cm.sales) { so += cm.sales; sp += cm.coll || 0; } }
        st.total_o = so; st.total_p = sp;
        if (st.total_rate === null && so) st.total_rate = 1 - sp / so;
      }
      if (prev && prev[s.code]) {
        st.prev_sales = prev[s.code].sales !== null ? prev[s.code].sales / 1e6 : null;
        st.prev_rate = prev[s.code].rate;
      }
      (data[BRANDS[letter]] = data[BRANDS[letter]] || []).push(st);
    }
    const ordered = {};
    for (const letter of ['S', 'T', 'A']) {
      const b = BRANDS[letter];
      if (!data[b]) continue;
      data[b].sort((x, y) => (groupOrder.indexOf(x.group) - groupOrder.indexOf(y.group)) || (y.total_sales - x.total_sales));
      ordered[b] = data[b];
    }
    return { data: ordered, months, merges };
  }

  function renderPage(template, data, months, year) {
    const n = months.length;
    const json = JSON.stringify(data).replace(/</g, '\\u003c');
    return template
      .replace('__DATA__', () => json)
      .replace('__MONTHS__', () => JSON.stringify(months.map(m => m + '월')))
      .replace('__PERIOD__', () => `${year}년 1월 ~ ${months[n - 1]}월`)
      .replace('__MW__', () => (63.2 / n).toFixed(2))
      .replace('__TITLE__', () => '임대매장 유통수수료율 월별 현황 | INDONG FN');
  }

  // ---------- UI ----------
  function makeLog() {
    const warns = [];
    return { warn: m => warns.push(m), warns };
  }

  let collWb = null, lastHtml = null, lastName = 'index.html';

  async function onCollectionFile() {
    const f = $('fileColl').files[0];
    const sel = $('collSheet');
    sel.innerHTML = ''; collWb = null;
    if (!f) { sel.disabled = true; return; }
    try {
      setStatus('수금 파일 시트 목록을 읽는 중…');
      collWb = await XLSXLite.openWorkbook(f);
      const pick = pickCollectionSheet(collWb.sheetNames);
      for (const n of collWb.sheetNames) {
        const o = document.createElement('option'); o.value = o.textContent = n; if (n === pick) o.selected = true; sel.appendChild(o);
      }
      sel.disabled = false;
      setStatus('');
    } catch (e) { setStatus(e.message, true); }
  }

  function setStatus(msg, err) {
    const el = $('status'); el.textContent = msg; el.className = 'status' + (err ? ' err' : '');
  }

  async function run() {
    const fs = $('fileSales').files[0], fc = $('fileColl').files[0], fp = $('filePrev').files[0];
    if (!fs || !fc) { setStatus('판매현황 파일과 수금 파일을 모두 선택해주세요.', true); return; }
    $('btnRun').disabled = true;
    $('result').hidden = true;
    const log = makeLog();
    $('schemaCard').hidden = true;
    try {
      setStatus('파일을 여는 중…');
      const salesWb = await XLSXLite.openWorkbook(fs);
      if (!collWb) collWb = await XLSXLite.openWorkbook(fc);
      const prevWb = fp ? await XLSXLite.openWorkbook(fp) : null;
      const collSheet = $('collSheet').value || pickCollectionSheet(collWb.sheetNames);

      setStatus('전달 기준과 컬럼명을 비교하는 중…');
      await runSchemaCheck(salesWb, collWb, collSheet, prevWb, { sales: fs.name, collection: fc.name, prev: fp ? fp.name : null });

      setStatus('판매현황 파일을 읽는 중…');
      const sales = await readSales(salesWb, log);
      setStatus('수금 파일을 읽는 중…');
      const coll = await readCollection(collWb, collSheet, log);
      setStatus('전년실적 파일을 읽는 중…');
      const prev = await readPrev(prevWb, log);
      const { data, months, merges } = build(sales, coll, prev, log);
      const template = await (await fetch('template.html', { cache: 'no-store' })).text();
      const year = sales.year || new Date().getFullYear();
      lastHtml = renderPage(template, data, months, year);
      lastName = `임대매장_유통수수료율_${year}_1-${months[months.length - 1]}월.html`;
      window.__lastResult = { data, months, merges, warns: log.warns, how: coll.how };
      showSummary(data, months, merges, log.warns, coll.how, !!prev);
      $('preview').srcdoc = lastHtml;
      $('result').hidden = false;
      setStatus('완료했어요. 아래에서 결과를 확인하고 HTML로 저장하세요.');
    } catch (e) {
      console.error(e);
      setStatus('오류: ' + e.message + (lastSchemaRows ? ' — 아래 「컬럼 변경 확인」에서 바뀐 컬럼을 확인하세요.' : ''), true);
    } finally {
      $('btnRun').disabled = false;
    }
  }

  function showSummary(data, months, merges, warns, how, hasPrev) {
    const all = Object.values(data).flat();
    const tot = all.reduce((s, x) => s + x.total_sales, 0);
    const O = all.reduce((s, x) => s + (x.total_o || 0), 0), P = all.reduce((s, x) => s + (x.total_p || 0), 0);
    const fmt = v => Math.round(v).toLocaleString('ko-KR');
    const brandLine = Object.entries(data).map(([b, v]) => `${b} ${v.length}개`).join(' · ');
    let h = `<ul class="facts">
      <li><b>기간</b> 1월 ~ ${months[months.length - 1]}월</li>
      <li><b>매장</b> ${all.length}개 (${brandLine}) — ${how}</li>
      <li><b>누계 합계 매출</b> ${fmt(tot)}백만원 · <b>전체 수수료율</b> ${O ? ((1 - P / O) * 100).toFixed(1) : '-'}%</li>
      <li><b>전년실적</b> ${hasPrev ? '전년실적 파일 반영' : '파일이 없어 비워둠'}</li>
    </ul>`;
    h += `<h3>행사 매장 합산 (${merges.length}곳)</h3>`;
    h += merges.length ? `<ul class="list">${merges.map(m => `<li><b>${m.code} ${m.name}</b> ← ${m.events.join(', ')}</li>`).join('')}</ul>` : '<p class="muted">합산한 행사 매장이 없어요.</p>';
    if (warns.length) h += `<h3>확인이 필요한 내용</h3><ul class="list warn">${warns.map(w => `<li>${w}</li>`).join('')}</ul>`;
    $('summary').innerHTML = h;
  }

  // ---------- 컬럼 변경 확인 ----------
  let lastSchema = null, lastSchemaRows = 0;
  const esc = t => String(t == null ? '' : t).replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]));

  async function loadBaseSchema() {
    try {
      const res = await fetch('schema.json', { cache: 'no-store' });
      if (!res.ok) return null;
      return await res.json();
    } catch (e) { return null; }
  }

  async function runSchemaCheck(salesWb, collWb, collSheet, prevWb, names) {
    const cur = {
      sales: await SchemaCheck.extractSales(salesWb),
      collection: await SchemaCheck.extractCollection(collWb, collSheet),
      prev: prevWb ? await SchemaCheck.extractPrev(prevWb) : null
    };
    const base = await loadBaseSchema();
    const { rows, notes } = SchemaCheck.compare(base, cur);
    lastSchema = SchemaCheck.makeSchema(cur, names);
    if (base && base.prev && !cur.prev) lastSchema.prev = base.prev;   // 전년실적을 안 올린 달은 기존 기준 유지
    lastSchemaRows = rows.length;
    const req = rows.filter(r => r.required).length;
    let h = '';
    if (!base) {
      h += `<p class="chip warn">전달 기준(schema.json)이 없어요</p><p class="muted">아래 버튼으로 이번 파일의 컬럼을 기준으로 저장한 뒤, 저장소에 schema.json으로 올려주세요. 다음 달부터 비교할 수 있어요.</p>`;
    } else if (!rows.length) {
      h += `<p class="chip ok">변경 없음</p><p class="muted">전달 기준(${esc(base.saved || '')} 저장)과 컬럼명이 모두 같아요. 매달 바뀌는 월 표기(예: 비교(YYYY-MM), N월 누계, 새 월 블록)는 비교에서 뺐어요.</p>`;
    } else {
      h += `<p class="chip ${req ? 'bad' : 'warn'}">변경 ${rows.length}건${req ? ` · 계산에 필요한 컬럼 ${req}건` : ''}</p>`;
      h += `<p class="muted">전달 기준(${esc(base.saved || '')} 저장)과 비교한 결과예요. <b>필수</b> 표시가 있는 컬럼은 이름이 바뀌면 계산이 안 되거나 값이 비어요.</p>`;
      h += `<div class="tbl"><table class="diff"><thead><tr><th>파일</th><th>시트</th><th>구역</th><th>구분</th><th>전달</th><th>이번 달</th><th></th></tr></thead><tbody>`;
      for (const r of rows) {
        const prevCell = r.old ? `${esc(r.old)}${r.oldCol ? ` <span class="col">${r.oldCol}열</span>` : ''}` : '—';
        const nowCell = r.now ? `${esc(r.now)}${r.newCol ? ` <span class="col">${r.newCol}열</span>` : ''}` : '—';
        h += `<tr class="${r.required ? 'req' : ''}"><td>${esc(r.file)}</td><td>${esc(r.sheet)}</td><td>${esc(r.area)}</td><td><span class="t t-${r.type}">${r.type}</span></td><td>${prevCell}</td><td>${nowCell}</td><td>${r.required ? '<b class="need">필수</b>' : ''}</td></tr>`;
      }
      h += `</tbody></table></div>`;
    }
    if (notes.length) h += `<ul class="list warn">${notes.map(n => `<li>${esc(n)}</li>`).join('')}</ul>`;
    $('schemaBox').innerHTML = h;
    $('schemaCard').hidden = false;
    window.__lastSchemaCheck = { rows, notes, hasBase: !!base };
  }

  function saveSchema() {
    if (!lastSchema) return;
    const blob = new Blob([JSON.stringify(lastSchema, null, 2)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = 'schema.json';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  }

  function download() {
    if (!lastHtml) return;
    const blob = new Blob([lastHtml], { type: 'text/html;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = lastName;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  }
  function openNew() {
    if (!lastHtml) return;
    const url = URL.createObjectURL(new Blob([lastHtml], { type: 'text/html;charset=utf-8' }));
    window.open(url, '_blank');
  }

  window.addEventListener('DOMContentLoaded', () => {
    $('fileColl').addEventListener('change', onCollectionFile);
    $('btnRun').addEventListener('click', run);
    $('btnSave').addEventListener('click', download);
    $('btnOpen').addEventListener('click', openNew);
    $('btnSchema').addEventListener('click', saveSchema);
    for (const id of ['fileSales', 'fileColl', 'filePrev']) {
      $(id).addEventListener('change', e => {
        const f = e.target.files[0];
        $(id + 'Name').textContent = f ? f.name : '선택된 파일 없음';
      });
    }
  });
})();
