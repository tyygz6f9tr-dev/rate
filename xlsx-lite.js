/*
 * xlsx-lite.js
 * 외부 라이브러리 없이 브라우저 기본 기능(DecompressionStream + DOMParser)만으로
 * .xlsx 파일의 값과 셀 채우기 색을 읽는 최소 파서.
 * 지원: 숫자/문자/공유문자열/수식 결과값(캐시값), 셀 배경색(테마색+tint, RGB)
 */
(function (global) {
  'use strict';

  // ---------- ZIP ----------
  function u16(dv, o) { return dv.getUint16(o, true); }
  function u32(dv, o) { return dv.getUint32(o, true); }

  async function inflateRaw(bytes) {
    if (typeof DecompressionStream === 'undefined') {
      throw new Error('이 브라우저는 압축 해제를 지원하지 않습니다. 최신 크롬이나 엣지에서 열어주세요.');
    }
    const ds = new DecompressionStream('deflate-raw');
    const stream = new Blob([bytes]).stream().pipeThrough(ds);
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  function readZip(buf) {
    const dv = new DataView(buf);
    let eocd = -1;
    for (let i = buf.byteLength - 22; i >= Math.max(0, buf.byteLength - 70000); i--) {
      if (u32(dv, i) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('엑셀(.xlsx) 파일이 아닙니다.');
    const count = u16(dv, eocd + 10);
    let p = u32(dv, eocd + 16);
    const dec = new TextDecoder('utf-8');
    const entries = {};
    for (let i = 0; i < count; i++) {
      if (u32(dv, p) !== 0x02014b50) break;
      const method = u16(dv, p + 10);
      const csize = u32(dv, p + 20);
      const nlen = u16(dv, p + 28), xlen = u16(dv, p + 30), clen = u16(dv, p + 32);
      const lho = u32(dv, p + 42);
      const name = dec.decode(new Uint8Array(buf, p + 46, nlen));
      entries[name] = { method, csize, lho };
      p += 46 + nlen + xlen + clen;
    }
    return {
      has(name) { return !!entries[name]; },
      async text(name) {
        const e = entries[name];
        if (!e) return null;
        const ln = u16(dv, e.lho + 26), lx = u16(dv, e.lho + 28);
        const start = e.lho + 30 + ln + lx;
        const raw = new Uint8Array(buf, start, e.csize);
        const data = e.method === 0 ? raw : await inflateRaw(raw);
        return dec.decode(data);
      }
    };
  }

  // ---------- XML helpers ----------
  const parser = new DOMParser();
  function xml(str) { return parser.parseFromString(str, 'application/xml'); }
  function byTag(node, tag) { return node.getElementsByTagNameNS('*', tag); }
  function textOf(node) { return node ? node.textContent : ''; }

  function colToIdx(col) {
    let n = 0;
    for (let i = 0; i < col.length; i++) n = n * 26 + (col.charCodeAt(i) - 64);
    return n - 1;
  }
  function splitRef(ref) {
    const m = /^([A-Z]+)(\d+)$/.exec(ref);
    return m ? { c: colToIdx(m[1]), r: parseInt(m[2], 10) - 1 } : null;
  }

  // ---------- 테마색 / tint ----------
  function hexToRgb(h) { return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)]; }
  function applyTint(rgb, tint) {
    if (!tint) return rgb;
    return rgb.map(v => tint < 0 ? Math.round(v * (1 + tint)) : Math.round(v + (255 - v) * tint));
  }

  async function readTheme(zip) {
    const t = await zip.text('xl/theme/theme1.xml');
    // Excel 테마 인덱스: 0 lt1, 1 dk1, 2 lt2, 3 dk2, 4~9 accent1~6
    const fallback = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];
    if (!t) return fallback;
    const doc = xml(t);
    const scheme = byTag(doc, 'clrScheme')[0];
    if (!scheme) return fallback;
    const pick = name => {
      const el = byTag(scheme, name)[0];
      if (!el) return null;
      const srgb = byTag(el, 'srgbClr')[0];
      if (srgb) return srgb.getAttribute('val');
      const sys = byTag(el, 'sysClr')[0];
      return sys ? (sys.getAttribute('lastClr') || null) : null;
    };
    const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
    return order.map((n, i) => pick(n) || fallback[i]);
  }

  async function readStyles(zip, theme) {
    const t = await zip.text('xl/styles.xml');
    if (!t) return [];
    const doc = xml(t);
    const fills = [];
    const fillsNode = byTag(doc, 'fills')[0];
    if (fillsNode) {
      for (const f of Array.from(fillsNode.children)) {
        const pf = byTag(f, 'patternFill')[0];
        let rgb = null;
        if (pf && pf.getAttribute('patternType') && pf.getAttribute('patternType') !== 'none') {
          const fg = byTag(pf, 'fgColor')[0];
          if (fg) {
            if (fg.hasAttribute('rgb')) rgb = hexToRgb(fg.getAttribute('rgb').slice(-6));
            else if (fg.hasAttribute('theme')) {
              const ti = parseInt(fg.getAttribute('theme'), 10);
              if (theme[ti]) rgb = applyTint(hexToRgb(theme[ti]), parseFloat(fg.getAttribute('tint') || '0'));
            }
          }
        }
        fills.push(rgb);
      }
    }
    const xfs = [];
    const cellXfs = byTag(doc, 'cellXfs')[0];
    if (cellXfs) {
      for (const xf of Array.from(cellXfs.children)) {
        const fid = parseInt(xf.getAttribute('fillId') || '0', 10);
        xfs.push(fills[fid] || null);
      }
    }
    return xfs; // 스타일 인덱스 → 배경 RGB(없으면 null)
  }

  async function openWorkbook(file) {
    const buf = await file.arrayBuffer();
    const zip = readZip(buf);
    const wbDoc = xml(await zip.text('xl/workbook.xml'));
    const relsDoc = xml(await zip.text('xl/_rels/workbook.xml.rels'));
    const rels = {};
    for (const r of Array.from(byTag(relsDoc, 'Relationship'))) {
      let target = r.getAttribute('Target');
      target = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
      rels[r.getAttribute('Id')] = target;
    }
    const sheets = [];
    for (const s of Array.from(byTag(wbDoc, 'sheet'))) {
      const rid = s.getAttribute('r:id') || s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
      sheets.push({ name: s.getAttribute('name'), path: rels[rid] });
    }
    let shared = [];
    const ssText = await zip.text('xl/sharedStrings.xml');
    if (ssText) {
      const ss = xml(ssText);
      shared = Array.from(byTag(ss, 'si')).map(si => Array.from(byTag(si, 't')).map(textOf).join(''));
    }
    const theme = await readTheme(zip);
    const fills = await readStyles(zip, theme);

    async function readSheet(name) {
      const info = sheets.find(s => s.name === name);
      if (!info) throw new Error(`'${name}' 시트를 찾을 수 없습니다.`);
      const doc = xml(await zip.text(info.path));
      const rows = [];
      const fillRows = [];
      for (const c of Array.from(byTag(doc, 'c'))) {
        const pos = splitRef(c.getAttribute('r'));
        if (!pos) continue;
        const t = c.getAttribute('t');
        const s = parseInt(c.getAttribute('s') || '0', 10);
        const vNode = byTag(c, 'v')[0];
        let v = null;
        if (t === 's') v = vNode ? shared[parseInt(textOf(vNode), 10)] : null;
        else if (t === 'inlineStr') v = Array.from(byTag(c, 't')).map(textOf).join('');
        else if (t === 'str' || t === 'e') v = vNode ? textOf(vNode) : null;
        else if (t === 'b') v = vNode ? textOf(vNode) === '1' : null;
        else if (vNode) { const n = parseFloat(textOf(vNode)); v = isNaN(n) ? textOf(vNode) : n; }
        (rows[pos.r] = rows[pos.r] || [])[pos.c] = v;
        if (fills[s]) (fillRows[pos.r] = fillRows[pos.r] || [])[pos.c] = fills[s];
      }
      return {
        name,
        rowCount: rows.length,
        get(r, c) { return rows[r] ? (rows[r][c] === undefined ? null : rows[r][c]) : null; },
        fill(r, c) { return fillRows[r] ? (fillRows[r][c] || null) : null; },
        row(r) { return rows[r] || []; }
      };
    }
    return { sheetNames: sheets.map(s => s.name), readSheet };
  }

  global.XLSXLite = { openWorkbook };
})(window);
