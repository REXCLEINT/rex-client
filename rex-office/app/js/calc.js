/* REX Tabelle – Tabellenkalkulation */
(function () {
  'use strict';
  const { $, $$, esc } = Rex.util;
  const RF = RexFormula, FMT = RexFormat;
  const DEF_W = RexXlsx.DEF_W, DEF_H = RexXlsx.DEF_H;
  const MAX_ROWS = 1048576, MAX_COLS = 16384;
  const BASE_FONT_PT = 11;

  const wrap = $('#gridWrap'), scroller = $('#scroller'), sizer = $('#sizer');
  const cellsEl = $('#cells'), gridEl = $('#gridlines'), input = $('#cellInput'), fx = $('#fxInput'), nameBox = $('#nameBox');

  // =====================================================================
  //  Modell
  // =====================================================================
  function newSheet(name) {
    return { name, cells: {}, colW: {}, rowH: {}, hidden: {}, merges: [], charts: [], grid: true, freeze: { r: 0, c: 0 }, filter: null };
  }
  let wb = { sheets: [newSheet('Tabelle1')], active: 0 };
  let S = wb.sheets[0];
  const key = (r, c) => r + ',' + c;
  const getCell = (r, c, sh = S) => sh.cells[key(r, c)];
  let zoom = 1;
  let showFormulas = false;

  // ---------- Werte ----------
  let cache = new Map();
  const computing = new Set();
  function sheetIndex(name) { return wb.sheets.findIndex(s => s.name.toLowerCase() === String(name).toLowerCase()); }
  function maxRow(si) {
    const sh = wb.sheets[si];
    if (sh._maxRow != null) return sh._maxRow;
    let m = 0;
    for (const k in sh.cells) { const r = +k.slice(0, k.indexOf(',')); if (r > m) m = r; }
    sh._maxRow = m;
    return m;
  }
  function getValue(si, r, c) {
    const k = si + '!' + r + ',' + c;
    if (cache.has(k)) return cache.get(k);
    const cell = wb.sheets[si] && wb.sheets[si].cells[r + ',' + c];
    if (!cell) return null;
    if (cell.f == null) return cell.v ?? null;
    if (computing.has(k)) return RF.ERR.CIRC;
    computing.add(k);
    let v;
    try {
      v = RF.compute(cell.f, { sheet: si, cur: { r, c }, value: getValue, sheetIndex, maxRow });
    } finally { computing.delete(k); }
    if (v === RF.ERR.NAME && cell.c != null) v = cell.c; // unbekannte Funktion: gespeicherten Wert zeigen
    cache.set(k, v);
    return v;
  }
  function invalidate() { cache = new Map(); for (const sh of wb.sheets) sh._maxRow = null; }

  // ---------- Eingaben interpretieren ----------
  function parseInput(text, oldStyle) {
    const out = { s: oldStyle ? { ...oldStyle } : undefined };
    if (text == null || text === '') return out;
    if (text[0] === "'") { out.v = text.slice(1); return out; }
    if (text[0] === '=' && text.length > 1) { out.f = text.slice(1); return out; }
    const t = text.trim();
    const setFmt = (f) => { if (!out.s || !out.s.fmt) out.s = { ...(out.s || {}), fmt: f }; };
    let m;
    if (/^(wahr|true)$/i.test(t)) { out.v = true; return out; }
    if (/^(falsch|false)$/i.test(t)) { out.v = false; return out; }
    if ((m = /^(\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})?$/.exec(t))) {
      let y = m[3] ? +m[3] : new Date().getFullYear();
      if (y < 100) y += y < 50 ? 2000 : 1900;
      const d = new Date(y, +m[2] - 1, +m[1]);
      if (d.getDate() === +m[1]) { out.v = Math.floor(RF.dateToSerial(d)); setFmt('DD.MM.YYYY'); return out; }
    }
    if ((m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(t))) { out.v = Math.floor(RF.dateToSerial(new Date(+m[1], +m[2] - 1, +m[3]))); setFmt('DD.MM.YYYY'); return out; }
    if ((m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(t))) { out.v = (+m[1] * 3600 + +m[2] * 60 + +(m[3] || 0)) / 86400; setFmt(m[3] ? 'hh:mm:ss' : 'hh:mm'); return out; }
    const n = RF.parseNumber(t.replace(/\s*(€|EUR)$/i, '').replace(/^(€|EUR)\s*/i, ''));
    if (n != null) {
      out.v = n;
      if (/%$/.test(t)) setFmt(/[.,]\d/.test(t) ? '0.00%' : '0%');
      else if (/€|EUR/i.test(t)) setFmt('#,##0.00 "€"');
      return out;
    }
    out.v = text;
    return out;
  }
  function inputText(cell) {
    if (!cell) return '';
    if (cell.f != null) return '=' + cell.f;
    const v = cell.v;
    if (v == null) return '';
    if (typeof v === 'boolean') return v ? 'WAHR' : 'FALSCH';
    if (typeof v === 'number') {
      const f = cell.s && cell.s.fmt;
      if (f && FMT.isDateFormat(f.split(';')[0])) {
        const d = RF.serialToDate(v);
        const date = `${String(d.getDate()).padStart(2, '0')}.${String(d.getMonth() + 1).padStart(2, '0')}.${d.getFullYear()}`;
        const hasTime = /h|s/i.test(f.replace(/"[^"]*"/g, ''));
        if (v < 1 && hasTime) return `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
        return hasTime && v % 1 ? `${date} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}` : date;
      }
      if (f && /%/.test(f)) return RF.fmtGeneral(Math.round(v * 1e10) / 1e8) + '%';
      return RF.fmtGeneral(v);
    }
    if (RF.isErr(v)) return v.code;
    if (typeof v === 'string' && v[0] === '=') return "'" + v;
    return String(v);
  }
  function display(si, r, c) {
    const sh = wb.sheets[si];
    const cell = sh.cells[key(r, c)];
    if (!cell) return null;
    if (showFormulas && cell.f != null) return { text: '=' + cell.f, align: 'left' };
    const v = getValue(si, r, c);
    if (v == null || v === '') return cell.f != null ? { text: '', align: 'left' } : null;
    const fmt = cell.s && cell.s.fmt;
    const res = FMT.formatFull(v, fmt);
    return {
      text: res.text, color: res.color, err: RF.isErr(v),
      align: typeof v === 'number' ? 'right' : (typeof v === 'boolean' || RF.isErr(v) ? 'center' : 'left'), num: typeof v === 'number'
    };
  }

  // =====================================================================
  //  Verlauf
  // =====================================================================
  const hist = { stack: [], index: -1 };
  function snapshot() { return JSON.stringify(wb, (k, v) => k === '_maxRow' ? undefined : v); }
  function resetHistory() { hist.stack = [snapshot()]; hist.index = 0; }
  function commit() {
    const s = snapshot();
    if (hist.stack[hist.index] === s) return;
    hist.stack = hist.stack.slice(0, hist.index + 1);
    hist.stack.push(s);
    if (hist.stack.length > 100) hist.stack.shift();
    hist.index = hist.stack.length - 1;
    invalidate();
    Rex.setDirty(true);
    render();
  }
  function restore(s) {
    const active = wb.active;
    wb = JSON.parse(s);
    wb.active = Math.min(active, wb.sheets.length - 1);
    S = wb.sheets[wb.active];
    invalidate();
    clampSel();
    renderTabs();
    render();
    Rex.setDirty(true);
  }
  function undo() { if (editing) cancelEdit(); if (hist.index > 0) restore(hist.stack[--hist.index]); }
  function redo() { if (editing) cancelEdit(); if (hist.index < hist.stack.length - 1) restore(hist.stack[++hist.index]); }

  // =====================================================================
  //  Geometrie
  // =====================================================================
  let geo = null;
  function buildGeo() {
    const cw = Object.entries(S.colW).map(([c, w]) => [+c, w]).sort((a, b) => a[0] - b[0]);
    const rhKeys = new Set([...Object.keys(S.rowH), ...Object.keys(S.hidden).filter(k => S.hidden[k])]);
    const rh = [...rhKeys].map(r => [+r, S.hidden[r] ? 0 : (S.rowH[r] ?? DEF_H)]).sort((a, b) => a[0] - b[0]);
    geo = { cw, rh };
  }
  const colW = (c) => (S.colW[c] ?? DEF_W) * zoom;
  const rowH = (r) => (S.hidden[r] ? 0 : (S.rowH[r] ?? DEF_H)) * zoom;
  function colX(c) {
    let x = c * DEF_W;
    for (const [cc, w] of geo.cw) { if (cc >= c) break; x += w - DEF_W; }
    return x * zoom;
  }
  function rowY(r) {
    let y = r * DEF_H;
    for (const [rr, h] of geo.rh) { if (rr >= r) break; y += h - DEF_H; }
    return y * zoom;
  }
  function colAt(x) {
    let lo = 0, hi = MAX_COLS - 1;
    while (lo < hi) { const mid = (lo + hi + 1) >> 1; if (colX(mid) <= x) lo = mid; else hi = mid - 1; }
    return lo;
  }
  function rowAt(y) {
    let lo = 0, hi = MAX_ROWS - 1;
    while (lo < hi) { const mid = (lo + hi + 1) >> 1; if (rowY(mid) <= y) lo = mid; else hi = mid - 1; }
    // versteckte Zeilen ueberspringen
    while (lo < MAX_ROWS - 1 && rowH(lo) === 0) lo++;
    return lo;
  }
  let extentRows = 200, extentCols = 40;
  function usedExtent() {
    let mr = 0, mc = 0;
    for (const k in S.cells) { const [r, c] = k.split(',').map(Number); if (r > mr) mr = r; if (c > mc) mc = c; }
    return { mr, mc };
  }
  function mergeAt(r, c, sh = S) {
    for (const m of sh.merges) if (r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) return m;
    return null;
  }

  // =====================================================================
  //  Auswahl
  // =====================================================================
  let sel = { r: 0, c: 0, r1: 0, c1: 0, r2: 0, c2: 0 };
  function normSel() {
    let { r1, c1, r2, c2 } = sel;
    if (r1 > r2) [r1, r2] = [r2, r1];
    if (c1 > c2) [c1, c2] = [c2, c1];
    // Verbundene Zellen einschliessen
    let changed = true;
    while (changed) {
      changed = false;
      for (const m of S.merges) {
        if (m.r2 >= r1 && m.r1 <= r2 && m.c2 >= c1 && m.c1 <= c2) {
          if (m.r1 < r1) { r1 = m.r1; changed = true; }
          if (m.c1 < c1) { c1 = m.c1; changed = true; }
          if (m.r2 > r2) { r2 = m.r2; changed = true; }
          if (m.c2 > c2) { c2 = m.c2; changed = true; }
        }
      }
    }
    return { r1, c1, r2, c2 };
  }
  let anchor = { r: 0, c: 0 };
  function select(r, c, extend = false) {
    r = Math.max(0, Math.min(MAX_ROWS - 1, r));
    c = Math.max(0, Math.min(MAX_COLS - 1, c));
    if (extend) {
      sel.r1 = anchor.r; sel.c1 = anchor.c; sel.r2 = r; sel.c2 = c;
      const n = normSel();
      sel = { ...sel, ...n, r1: n.r1, c1: n.c1, r2: n.r2, c2: n.c2 };
    } else {
      const m = mergeAt(r, c);
      if (m) { r = m.r1; c = m.c1; }
      anchor = { r, c };
      sel = { r, c, r1: r, c1: c, r2: m ? m.r2 : r, c2: m ? m.c2 : c };
    }
    if (r >= extentRows - 20) extentRows = Math.min(MAX_ROWS, r + 100);
    if (c >= extentCols - 5) extentCols = Math.min(MAX_COLS, c + 20);
    scrollIntoView(extend ? r : sel.r, extend ? c : sel.c);
    render();
    updateFormulaBar();
    updateToolbar();
  }
  function clampSel() {
    sel.r = Math.min(sel.r, MAX_ROWS - 1);
    sel.c = Math.min(sel.c, MAX_COLS - 1);
  }
  function selRect() { return { r1: sel.r1, c1: sel.c1, r2: sel.r2, c2: sel.c2 }; }
  function eachSel(fn) {
    const { r1, c1, r2, c2 } = selRect();
    const rr2 = Math.min(r2, Math.max(r1, usedExtent().mr + 1)), cc2 = Math.min(c2, Math.max(c1, usedExtent().mc + 1));
    for (let r = r1; r <= (r2 > 5000 ? rr2 : r2); r++) for (let c = c1; c <= (c2 > 500 ? cc2 : c2); c++) fn(r, c);
  }
  function selAddr() {
    const { r1, c1, r2, c2 } = selRect();
    if (r1 === r2 && c1 === c2) return RF.addr(r1, c1);
    const m = mergeAt(r1, c1);
    if (m && m.r1 === r1 && m.c1 === c1 && m.r2 === r2 && m.c2 === c2) return RF.addr(r1, c1);
    if (r1 === 0 && r2 >= MAX_ROWS - 1) return RF.colName(c1) + ':' + RF.colName(c2);
    if (c1 === 0 && c2 >= MAX_COLS - 1) return (r1 + 1) + ':' + (r2 + 1);
    return RF.addr(r1, c1) + ':' + RF.addr(r2, c2);
  }
  function scrollIntoView(r, c) {
    const fr = S.freeze || { r: 0, c: 0 };
    const frH = rowY(fr.r), frW = colX(fr.c);
    const x = colX(c), y = rowY(r), w = colW(c), h = rowH(r);
    const vw = scroller.clientWidth, vh = scroller.clientHeight;
    if (c >= fr.c) {
      if (x - frW < scroller.scrollLeft) scroller.scrollLeft = x - frW;
      else if (x + w > scroller.scrollLeft + vw) scroller.scrollLeft = x + w - vw;
    }
    if (r >= fr.r) {
      if (y - frH < scroller.scrollTop) scroller.scrollTop = y - frH;
      else if (y + h > scroller.scrollTop + vh) scroller.scrollTop = y + h - vh;
    }
  }

  // =====================================================================
  //  Darstellung
  // =====================================================================
  const measureCtx = document.createElement('canvas').getContext('2d');
  function fontCss(s) {
    const size = ((s && s.size) || BASE_FONT_PT) * 4 / 3 * zoom;
    return `${s && s.i ? 'italic ' : ''}${s && s.b ? 'bold ' : ''}${size}px ${s && s.font ? `"${s.font}", ` : ''}Calibri, Carlito, "Segoe UI", Arial, sans-serif`;
  }
  function textWidth(text, s) { measureCtx.font = fontCss(s); return measureCtx.measureText(text).width; }

  function cellStyleCss(s) {
    if (!s) return '';
    let css = '';
    if (s.b) css += 'font-weight:bold;';
    if (s.i) css += 'font-style:italic;';
    if (s.u || s.s) css += `text-decoration:${s.u ? 'underline ' : ''}${s.s ? 'line-through' : ''};`;
    if (s.font) css += `font-family:"${s.font.replace(/"/g, '')}",Calibri,Carlito,sans-serif;`;
    if (s.size) css += `font-size:${s.size * 4 / 3 * zoom}px;`;
    if (s.color) css += `color:${s.color};`;
    if (s.bg) css += `background:${s.bg};`;
    if (s.bd) {
      const w = { thin: 1, medium: 2, thick: 3 };
      for (const [k, prop] of [['t', 'top'], ['r', 'right'], ['b', 'bottom'], ['l', 'left']]) {
        if (s.bd[k]) css += `border-${prop}:${w[s.bd[k]]}px solid ${s.bd[k + 'c'] || '#000'};`;
      }
    }
    return css;
  }

  function cellHTML(r, c, x, y, w, h) {
    const cell = getCell(r, c);
    if (!cell && !mergeAt(r, c)) return '';
    const s = cell && cell.s;
    const d = cell ? display(wb.active, r, c) : null;
    if (!d && !(s && (s.bg || s.bd))) return '';
    let cls = 'c';
    let align = (s && s.al) || (d && d.align) || 'left';
    let text = d ? d.text : '';
    let width = w;
    if (s && s.wrap) cls += ' wrap';
    if (s && s.va === 'top') cls += ' va-top';
    else if (s && s.va === 'middle') cls += ' va-middle';
    if (align === 'right') cls += ' al-right';
    else if (align === 'center') cls += ' al-center';
    if (d && d.err) cls += ' err';
    let style = cellStyleCss(s);
    if (d && d.color) style += `color:${d.color};`;
    if (d && text && !(s && s.wrap)) {
      const tw = textWidth(text, s) + 6;
      if (d.num && tw > w && !showFormulas) {
        text = '#'.repeat(Math.max(1, Math.floor(w / (7 * zoom))));
      } else if (tw > w && align === 'left') {
        // Text in leere Nachbarzellen ueberlaufen lassen
        let cc = c + 1, ww = w;
        while (ww < tw && cc < c + 30 && !getCell(r, cc) && !mergeAt(r, cc)) { ww += colW(cc); cc++; }
        if (ww > w) { width = ww; cls += ' ov'; if (!(s && s.bg)) style += 'background:transparent;'; }
      } else if (tw > w && align === 'center') {
        cls += ' ov';
      }
    }
    const inner = esc(text);
    return `<div class="${cls}" style="left:${x}px;top:${y}px;width:${width}px;height:${h}px;${style}${width > w ? 'z-index:1;' : ''}"><span>${inner}</span></div>`;
  }

  function regionHTML(r0, r1, c0, c1) {
    let html = '';
    const done = new Set();
    // Verbundene Zellen zuerst
    for (const m of S.merges) {
      if (m.r2 < r0 || m.r1 > r1 || m.c2 < c0 || m.c1 > c1) continue;
      const x = colX(m.c1), y = rowY(m.r1);
      const w = colX(m.c2 + 1) - x, h = rowY(m.r2 + 1) - y;
      const cell = getCell(m.r1, m.c1);
      const bg = cell && cell.s && cell.s.bg ? '' : 'background:var(--cell-bg);';
      const inner = cellHTML(m.r1, m.c1, x, y, w, h);
      html += inner ? inner.replace('style="', `style="${bg}`) : `<div class="c" style="left:${x}px;top:${y}px;width:${w}px;height:${h}px;${bg}"></div>`;
      for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) done.add(key(r, c));
    }
    for (let r = r0; r <= r1; r++) {
      const h = rowH(r);
      if (!h) continue;
      const y = rowY(r);
      for (let c = c0; c <= c1; c++) {
        if (done.has(key(r, c))) continue;
        const w = colW(c);
        if (!w) continue;
        html += cellHTML(r, c, colX(c), y, w, h);
      }
    }
    return html;
  }
  function linesHTML(r0, r1, c0, c1) {
    const top = rowY(r0), bottom = rowY(r1 + 1), left = colX(c0), right = colX(c1 + 1);
    let html = '';
    for (let c = c0; c <= c1 + 1; c++) html += `<div class="gl-v" style="left:${colX(c) - 1}px;top:${top}px;height:${bottom - top}px"></div>`;
    for (let r = r0; r <= r1 + 1; r++) { if (r > r0 && !rowH(r - 1)) continue; html += `<div class="gl-h" style="top:${rowY(r) - 1}px;left:${left}px;width:${right - left}px"></div>`; }
    return html;
  }
  function visibleRange() {
    const fr = S.freeze || { r: 0, c: 0 };
    const sl = scroller.scrollLeft, st = scroller.scrollTop;
    const vw = scroller.clientWidth, vh = scroller.clientHeight;
    const c0 = Math.max(fr.c, colAt(sl + colX(fr.c))), c1 = Math.min(MAX_COLS - 1, colAt(sl + vw) + 1);
    const r0 = Math.max(fr.r, rowAt(st + rowY(fr.r)) - 1), r1 = Math.min(MAX_ROWS - 1, rowAt(st + vh) + 1);
    return { r0: Math.max(0, r0), r1, c0: Math.max(0, c0), c1 };
  }

  let renderPending = false;
  function render() {
    if (renderPending) return;
    renderPending = true;
    requestAnimationFrame(() => { renderPending = false; renderNow(); });
  }
  function renderNow() {
    buildGeo();
    const { mr, mc } = usedExtent();
    extentRows = Math.max(extentRows, mr + 60, sel.r2 < MAX_ROWS - 1 ? sel.r2 + 30 : 0);
    extentCols = Math.max(extentCols, mc + 10, sel.c2 < MAX_COLS - 1 ? sel.c2 + 5 : 0);
    for (const ch of S.charts) {
      extentRows = Math.max(extentRows, rowAt(ch.y * zoom + ch.h * zoom) + 10);
      extentCols = Math.max(extentCols, colAt(ch.x * zoom + ch.w * zoom) + 3);
    }
    sizer.style.width = colX(extentCols) + 'px';
    sizer.style.height = rowY(extentRows) + 'px';
    wrap.style.setProperty('--hdr-w', Math.max(46, String(extentRows).length * 8 + 14) * Math.max(1, zoom * 0.9) + 'px');
    wrap.style.setProperty('--hdr-h', Math.round(22 * Math.max(1, zoom * 0.9)) + 'px');
    wrap.classList.toggle('no-grid', S.grid === false);

    const { r0, r1, c0, c1 } = visibleRange();
    gridEl.innerHTML = linesHTML(r0, r1, c0, c1);
    cellsEl.innerHTML = regionHTML(r0, r1, c0, c1) + filterButtonsHTML(r0, r1, c0, c1);
    renderHeaders(r0, r1, c0, c1);
    renderSelection();
    renderCharts();
    renderFrozen(r1, c1);
    positionInput();
    updateStatus();
  }

  function renderHeaders(r0, r1, c0, c1) {
    const fr = S.freeze || { r: 0, c: 0 };
    const { r1: sr1, c1: sc1, r2: sr2, c2: sc2 } = selRect();
    let h = '';
    const colList = [];
    for (let c = 0; c < fr.c; c++) colList.push(c);
    for (let c = c0; c <= c1; c++) colList.push(c);
    for (const c of colList) {
      const w = colW(c);
      if (!w) continue;
      const frozen = c < fr.c;
      const x = frozen ? colX(c) : colX(c) - scroller.scrollLeft;
      h += `<div class="hc${c >= sc1 && c <= sc2 ? ' sel' : ''}" data-c="${c}" style="left:${x}px;width:${w}px;${frozen ? 'z-index:1;background:var(--hdr-bg);' : ''}">${RF.colName(c)}<div class="rz" data-rc="${c}"></div></div>`;
    }
    $('#colHeadInner').innerHTML = h;
    h = '';
    const rowList = [];
    for (let r = 0; r < fr.r; r++) rowList.push(r);
    for (let r = r0; r <= r1; r++) rowList.push(r);
    for (const r of rowList) {
      const hh = rowH(r);
      if (!hh) continue;
      const frozen = r < fr.r;
      const y = frozen ? rowY(r) : rowY(r) - scroller.scrollTop;
      h += `<div class="hr${r >= sr1 && r <= sr2 ? ' sel' : ''}" data-r="${r}" style="top:${y}px;height:${hh}px;${frozen ? 'z-index:1;background:var(--hdr-bg);' : ''}">${r + 1}<div class="rz" data-rr="${r}"></div></div>`;
    }
    $('#rowHeadInner').innerHTML = h;
  }

  function selBoxes() {
    const { r1, c1, r2, c2 } = selRect();
    const x = colX(c1), y = rowY(r1);
    const w = colX(Math.min(c2, MAX_COLS - 1) + 1) - x, h = rowY(Math.min(r2, MAX_ROWS - 1) + 1) - y;
    const m = mergeAt(sel.r, sel.c);
    const ax = colX(sel.c), ay = rowY(sel.r);
    const aw = m ? colX(m.c2 + 1) - ax : colW(sel.c), ah = m ? rowY(m.r2 + 1) - ay : rowH(sel.r);
    return { x, y, w, h, ax, ay, aw, ah, multi: !(r1 === r2 && c1 === c2) && !(m && m.r1 === r1 && m.c1 === c1 && m.r2 === r2 && m.c2 === c2) };
  }
  function renderSelection() {
    const b = selBoxes();
    const rg = $('#selRange'), cur = $('#selCursor'), fh = $('#fillHandle');
    rg.style.display = b.multi ? 'block' : 'none';
    Object.assign(rg.style, { left: b.x + 'px', top: b.y + 'px', width: b.w + 'px', height: b.h + 'px' });
    Object.assign(cur.style, { left: b.ax - 1 + 'px', top: b.ay - 1 + 'px', width: b.aw + 2 + 'px', height: b.ah + 2 + 'px' });
    Object.assign(fh.style, { left: b.x + b.w - 5 + 'px', top: b.y + b.h - 5 + 'px' });
    fh.style.display = editing ? 'none' : 'block';
    const ants = $('#copyAnts');
    if (clip && clip.si === wb.active && !clip.done) {
      const x = colX(clip.c1), y = rowY(clip.r1);
      Object.assign(ants.style, { left: x - 1 + 'px', top: y - 1 + 'px', width: colX(clip.c2 + 1) - x + 2 + 'px', height: rowY(clip.r2 + 1) - y + 2 + 'px' });
      ants.hidden = false;
    } else ants.hidden = true;
    // Formelbezuege farbig markieren
    $$('.ref-hl', sizer).forEach(e => e.remove());
    if (editing) {
      const txt = input.value;
      if (txt.startsWith('=')) {
        const colors = ['#4472c4', '#c00000', '#7030a0', '#00b050', '#ed7d31', '#0070c0'];
        RF.refsIn(txt.slice(1)).forEach((t, i) => {
          if (t.sheet && sheetIndex(t.sheet) !== wb.active) return;
          const rg2 = RF.refRange(t.v);
          if (rg2.r2 === Infinity) rg2.r2 = Math.min(extentRows, 2000);
          const d = document.createElement('div');
          d.className = 'ref-hl';
          const x = colX(rg2.c1), y = rowY(rg2.r1);
          Object.assign(d.style, { left: x + 'px', top: y + 'px', width: colX(rg2.c2 + 1) - x + 'px', height: rowY(rg2.r2 + 1) - y + 'px', borderColor: colors[i % colors.length] });
          sizer.appendChild(d);
        });
      }
    }
  }

  // Fixierte Bereiche als Ueberlagerung
  function renderFrozen(r1, c1) {
    $$('.frozen-pane', wrap).forEach(e => e.remove());
    const fr = S.freeze || { r: 0, c: 0 };
    if (!fr.r && !fr.c) return;
    const hdrW = parseFloat(getComputedStyle(wrap).getPropertyValue('--hdr-w')), hdrH = parseFloat(getComputedStyle(wrap).getPropertyValue('--hdr-h'));
    const fh = rowY(fr.r), fw = colX(fr.c);
    const { c0, r0 } = visibleRange();
    const mk = (left, top, w, h, dx, dy, html) => {
      const p = document.createElement('div');
      p.className = 'frozen-pane';
      p.style.cssText = `position:absolute;left:${left}px;top:${top}px;width:${w}px;height:${h}px;overflow:hidden;z-index:1;background:var(--cell-bg);pointer-events:none;`;
      p.innerHTML = `<div style="position:absolute;left:${-dx}px;top:${-dy}px">${html}</div>`;
      wrap.insertBefore(p, $('#acList'));
      return p;
    };
    const sw = scroller.clientWidth, sh = scroller.clientHeight;
    const selHTML = () => {
      const b = selBoxes();
      return `<div class="sel-range" style="display:${b.multi ? 'block' : 'none'};left:${b.x}px;top:${b.y}px;width:${b.w}px;height:${b.h}px"></div>
        <div class="sel-cursor" style="left:${b.ax - 1}px;top:${b.ay - 1}px;width:${b.aw + 2}px;height:${b.ah + 2}px"></div>`;
    };
    const lines = (a, b, c, d) => S.grid === false ? '' : linesHTML(a, b, c, d).replace(/class="gl-/g, 'style="position:absolute;background:var(--grid-line);" class="gl-').replace(/" style="/g, ';');
    if (fr.r) {
      const html = `<div class="gridlines" style="position:absolute">${lines(0, fr.r - 1, c0, c1)}</div>${regionHTML(0, fr.r - 1, c0, c1)}${selHTML()}`;
      mk(hdrW + fw, hdrH, Math.max(0, sw - fw), fh, scroller.scrollLeft + fw, 0, html).style.borderBottom = '1px solid #9a9a9a';
    }
    if (fr.c) {
      const html = `<div class="gridlines" style="position:absolute">${lines(r0, r1, 0, fr.c - 1)}</div>${regionHTML(r0, r1, 0, fr.c - 1)}${selHTML()}`;
      mk(hdrW, hdrH + fh, fw, Math.max(0, sh - fh), 0, scroller.scrollTop + fh, html).style.borderRight = '1px solid #9a9a9a';
    }
    if (fr.r && fr.c) {
      const html = `<div class="gridlines" style="position:absolute">${lines(0, fr.r - 1, 0, fr.c - 1)}</div>${regionHTML(0, fr.r - 1, 0, fr.c - 1)}${selHTML()}`;
      mk(hdrW, hdrH, fw, fh, 0, 0, html);
    }
  }

  // =====================================================================
  //  Diagramme
  // =====================================================================
  let selChart = null;
  function resolveRef(ref, si = wb.active) {
    if (!ref) return [];
    if (ref.startsWith('"')) return [ref.slice(1, -1)];
    const toks = RF.refsIn(ref);
    if (!toks.length) return [];
    const t = toks[0];
    const s2 = t.sheet != null ? sheetIndex(t.sheet) : si;
    if (s2 < 0) return [];
    const rg = RF.refRange(t.v);
    const r2 = Math.min(rg.r2, maxRow(s2));
    const out = [];
    for (let r = rg.r1; r <= r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
      const v = getValue(s2, r, c);
      out.push(RF.isErr(v) ? null : v);
    }
    return out;
  }
  function chartData(ch) {
    const series = ch.series.map((s, i) => ({
      name: s.name ? String(resolveRef(s.name)[0] ?? '') : `Reihe ${i + 1}`,
      values: resolveRef(s.val).map(v => typeof v === 'number' ? v : (RF.parseNumber(String(v ?? '')) ?? 0))
    }));
    const catRef = ch.series[0] && ch.series[0].cat;
    const categories = catRef ? resolveRef(catRef).map((v, i) => v == null ? String(i + 1) : (typeof v === 'number' ? RF.fmtGeneral(v) : String(v))) : null;
    return { type: ch.type, title: ch.title, series, categories };
  }
  function renderCharts() {
    const layer = $('#charts');
    const dark = document.documentElement.dataset.theme === 'dark';
    const existing = new Map($$('.chart', layer).map(e => [e.dataset.id, e]));
    for (const ch of S.charts) {
      let el = existing.get(ch.id);
      existing.delete(ch.id);
      if (!el) {
        el = document.createElement('div');
        el.className = 'chart';
        el.dataset.id = ch.id;
        el.innerHTML = '<div class="svg"></div><div class="ch-rz"></div><div class="ch-tools"><button data-ct="edit" title="Diagramm bearbeiten"><i data-lucide="pencil"></i></button><button data-ct="del" title="Diagramm löschen"><i data-lucide="trash-2"></i></button></div>';
        layer.appendChild(el);
        Rex.icons();
      }
      const w = ch.w * zoom, h = ch.h * zoom;
      Object.assign(el.style, { left: ch.x * zoom + 'px', top: ch.y * zoom + 'px', width: w + 'px', height: h + 'px' });
      el.classList.toggle('sel', selChart === ch.id);
      const svg = RexCharts.render(chartData(ch), ch.w, ch.h, { dark });
      const box = el.querySelector('.svg');
      if (box._svg !== svg) { box.innerHTML = svg; box._svg = svg; box.style.cssText = 'width:100%;height:100%'; }
    }
    for (const el of existing.values()) el.remove();
  }
  function currentRegion(r, c) {
    let r1 = r, r2 = r, c1 = c, c2 = c, grew = true;
    const filled = (rr, cc) => rr >= 0 && cc >= 0 && getCell(rr, cc) && (getCell(rr, cc).v != null || getCell(rr, cc).f != null);
    while (grew) {
      grew = false;
      for (let cc = c1 - 1; cc <= c2 + 1; cc++) {
        if (r1 > 0 && filled(r1 - 1, cc)) { r1--; grew = true; break; }
      }
      for (let cc = c1 - 1; cc <= c2 + 1; cc++) {
        if (filled(r2 + 1, cc)) { r2++; grew = true; break; }
      }
      for (let rr = r1; rr <= r2; rr++) {
        if (c1 > 0 && filled(rr, c1 - 1)) { c1--; grew = true; break; }
      }
      for (let rr = r1; rr <= r2; rr++) {
        if (filled(rr, c2 + 1)) { c2++; grew = true; break; }
      }
      if (r2 - r1 > 100000 || c2 - c1 > 1000) break;
    }
    return { r1, c1, r2, c2 };
  }
  function insertChart(type) {
    let rg = selRect();
    if (rg.r1 === rg.r2 && rg.c1 === rg.c2) rg = currentRegion(sel.r, sel.c);
    if (rg.r2 > 100000) rg.r2 = Math.min(rg.r2, maxRow(wb.active));
    const v = (r, c) => getValue(wb.active, r, c);
    const isNum = (x) => typeof x === 'number';
    let headerRow = false;
    for (let c = rg.c1; c <= rg.c2; c++) { const x = v(rg.r1, c); if (x != null && !isNum(x)) headerRow = true; }
    if (rg.r1 === rg.r2) headerRow = false;
    let catCol = false;
    for (let r = rg.r1 + (headerRow ? 1 : 0); r <= rg.r2; r++) { const x = v(r, rg.c1); if (x != null && !isNum(x)) catCol = true; }
    if (headerRow && v(rg.r1, rg.c1) == null && rg.c2 > rg.c1) catCol = true;
    if (rg.c1 === rg.c2) catCol = false;
    const dr1 = rg.r1 + (headerRow ? 1 : 0);
    const abs = (r, c) => '$' + RF.colName(c) + '$' + (r + 1);
    const series = [];
    for (let c = rg.c1 + (catCol ? 1 : 0); c <= rg.c2; c++) {
      series.push({
        name: headerRow ? abs(rg.r1, c) : null,
        cat: catCol ? abs(dr1, rg.c1) + ':' + abs(rg.r2, rg.c1) : null,
        val: abs(dr1, c) + ':' + abs(rg.r2, c)
      });
    }
    if (!series.length || dr1 > rg.r2) { Rex.toast('Bitte zuerst Zahlen auswählen (z. B. A1:B6).'); return; }
    const titles = { column: 'Säulendiagramm', bar: 'Balkendiagramm', line: 'Liniendiagramm', pie: 'Kreisdiagramm' };
    const title = series.length === 1 && headerRow ? String(v(rg.r1, rg.c1 + (catCol ? 1 : 0)) ?? '') : titles[type];
    const ch = { id: 'ch' + Date.now().toString(36), type, title, series, x: colX(rg.c2 + 2) / zoom, y: rowY(rg.r1) / zoom, w: 480, h: 290 };
    S.charts.push(ch);
    selChart = ch.id;
    commit();
  }
  async function editChart(ch) {
    const v = await Rex.dialog({ title: 'Diagramm bearbeiten', fields: [
      { name: 'title', label: 'Titel', value: ch.title || '' },
      { name: 'type', label: 'Diagrammtyp', type: 'select', value: ch.type, options: [{ value: 'column', label: 'Säulen' }, { value: 'bar', label: 'Balken' }, { value: 'line', label: 'Linie' }, { value: 'pie', label: 'Kreis' }] }
    ] });
    if (!v) return;
    ch.title = v.title;
    ch.type = v.type;
    commit();
  }
  $('#charts').addEventListener('mousedown', e => {
    const el = e.target.closest('.chart');
    if (!el) return;
    e.stopPropagation();
    e.preventDefault();
    const ch = S.charts.find(x => x.id === el.dataset.id);
    if (!ch) return;
    const tool = e.target.closest('[data-ct]');
    if (tool) {
      if (tool.dataset.ct === 'del') { S.charts = S.charts.filter(x => x !== ch); selChart = null; commit(); }
      else editChart(ch);
      return;
    }
    selChart = ch.id;
    renderCharts();
    const resize = e.target.classList.contains('ch-rz');
    const sx = e.clientX, sy = e.clientY, ox = ch.x, oy = ch.y, ow = ch.w, oh = ch.h;
    let moved = false;
    const move = ev => {
      const dx = (ev.clientX - sx) / zoom, dy = (ev.clientY - sy) / zoom;
      if (Math.abs(dx) + Math.abs(dy) > 2) moved = true;
      if (resize) { ch.w = Math.max(160, ow + dx); ch.h = Math.max(110, oh + dy); }
      else { ch.x = Math.max(0, ox + dx); ch.y = Math.max(0, oy + dy); }
      renderCharts();
    };
    const up = () => { document.removeEventListener('mousemove', move); document.removeEventListener('mouseup', up); if (moved) commit(); focusGrid(); };
    document.addEventListener('mousemove', move);
    document.addEventListener('mouseup', up);
  });
  $('#charts').addEventListener('dblclick', e => {
    const el = e.target.closest('.chart');
    const ch = el && S.charts.find(x => x.id === el.dataset.id);
    if (ch) editChart(ch);
  });

  // =====================================================================
  //  Filter
  // =====================================================================
  function filterButtonsHTML(r0, r1, c0, c1) {
    const f = S.filter;
    if (!f || f.r1 < r0 - 1 || f.r1 > r1) return '';
    let html = '';
    for (let c = Math.max(c0, f.c1); c <= Math.min(c1, f.c2); c++) {
      const x = colX(c + 1) - 18 * Math.max(1, zoom), y = rowY(f.r1) + rowH(f.r1) - 18;
      const active = f.crit && f.crit[c];
      html += `<div class="filter-btn${active ? ' active' : ''}" data-fc="${c}" style="left:${x}px;top:${y}px">${active ? '▼' : '▾'}</div>`;
    }
    return html;
  }
  function applyFilter() {
    const f = S.filter;
    if (!f) return;
    for (let r = f.r1 + 1; r <= f.r2; r++) {
      let show = true;
      for (const [c, allowed] of Object.entries(f.crit || {})) {
        const d = display(wb.active, r, +c);
        if (!allowed.includes(d ? d.text : '')) { show = false; break; }
      }
      if (show) delete S.hidden[r]; else S.hidden[r] = true;
    }
  }
  function toggleFilter() {
    if (S.filter) {
      for (let r = S.filter.r1 + 1; r <= S.filter.r2; r++) delete S.hidden[r];
      S.filter = null;
      commit();
      return;
    }
    let rg = selRect();
    if (rg.r1 === rg.r2) rg = currentRegion(sel.r, sel.c);
    if (rg.r1 === rg.r2) { Rex.toast('Bitte einen Bereich mit Überschriften auswählen.'); return; }
    S.filter = { ...rg, crit: {} };
    commit();
  }
  function filterMenu(c, anchorEl) {
    const f = S.filter;
    const values = new Map();
    for (let r = f.r1 + 1; r <= f.r2; r++) {
      const d = display(wb.active, r, c);
      const t = d ? d.text : '';
      values.set(t, (values.get(t) || 0) + 1);
    }
    const allowed = f.crit[c] ? new Set(f.crit[c]) : null;
    const sorted = [...values.keys()].sort((a, b) => a.localeCompare(b, 'de', { numeric: true }));
    Rex.menu(anchorEl, m => {
      const mkBtn = (label, icon, fn) => { const b = document.createElement('button'); b.innerHTML = `<i data-lucide="${icon}"></i> ${label}`; b.onclick = () => { Rex.closeMenu(); fn(); }; m.appendChild(b); };
      mkBtn('Von A bis Z sortieren', 'arrow-up-a-z', () => sortRange({ r1: f.r1 + 1, c1: f.c1, r2: f.r2, c2: f.c2 }, c, true));
      mkBtn('Von Z bis A sortieren', 'arrow-down-z-a', () => sortRange({ r1: f.r1 + 1, c1: f.c1, r2: f.r2, c2: f.c2 }, c, false));
      m.appendChild(document.createElement('hr'));
      const list = document.createElement('div');
      list.style.cssText = 'max-height:240px;overflow:auto;padding:4px 8px;display:flex;flex-direction:column;gap:3px';
      const all = document.createElement('label');
      all.innerHTML = `<input type="checkbox" ${!allowed ? 'checked' : ''}> <b>(Alles auswählen)</b>`;
      list.appendChild(all);
      const boxes = sorted.map(v => {
        const l = document.createElement('label');
        l.innerHTML = `<input type="checkbox" ${!allowed || allowed.has(v) ? 'checked' : ''}> ${esc(v || '(Leer)')}`;
        l.querySelector('input').dataset.v = v;
        list.appendChild(l);
        return l.querySelector('input');
      });
      all.querySelector('input').onchange = (ev) => boxes.forEach(b => { b.checked = ev.target.checked; });
      m.appendChild(list);
      const ok = document.createElement('button');
      ok.innerHTML = '<i data-lucide="check"></i> OK';
      ok.onclick = () => {
        Rex.closeMenu();
        const chosen = boxes.filter(b => b.checked).map(b => b.dataset.v);
        if (chosen.length === boxes.length) delete f.crit[c]; else f.crit[c] = chosen;
        applyFilter();
        commit();
      };
      m.appendChild(ok);
    });
  }

  // =====================================================================
  //  Bearbeitung
  // =====================================================================
  let editing = null; // { r, c, mode: 'enter'|'edit', fromBar }
  function positionInput() {
    const m = mergeAt(sel.r, sel.c);
    const x = colX(sel.c), y = rowY(sel.r);
    const w = m ? colX(m.c2 + 1) - x : colW(sel.c), h = m ? rowY(m.r2 + 1) - y : rowH(sel.r);
    input.style.left = x - 1 + 'px';
    input.style.top = y - 1 + 'px';
    if (editing) {
      const cell = getCell(sel.r, sel.c);
      const s = cell && cell.s;
      input.style.font = fontCss(s);
      input.style.fontWeight = s && s.b ? 'bold' : '';
      input.style.minWidth = w + 2 + 'px';
      input.style.minHeight = h + 2 + 'px';
      const tw = Math.max(...input.value.split('\n').map(l => textWidth(l, s))) + 16;
      input.style.width = Math.max(w + 2, Math.min(tw, scroller.clientWidth - 20)) + 'px';
      input.style.height = Math.max(h + 2, input.scrollHeight) + 'px';
      input.style.textAlign = s && s.al === 'right' ? 'right' : s && s.al === 'center' ? 'center' : 'left';
    }
  }
  function beginEdit(mode, initial) {
    if (editing) return;
    editing = { r: sel.r, c: sel.c, mode };
    input.classList.remove('hidden');
    if (initial !== undefined) input.value = initial;
    else if (mode === 'edit') input.value = inputText(getCell(sel.r, sel.c));
    fx.value = input.value;
    $('#stMode').textContent = mode === 'enter' ? 'Eingeben' : 'Bearbeiten';
    positionInput();
    renderSelection();
    if (document.activeElement !== input) input.focus();
    if (mode === 'edit') { const l = input.value.length; input.setSelectionRange(l, l); }
  }
  function cancelEdit() {
    if (!editing) return;
    editing = null;
    input.value = '';
    input.classList.add('hidden');
    hideAC();
    $('#stMode').textContent = 'Bereit';
    updateFormulaBar();
    render();
  }
  function commitEdit() {
    if (!editing) return true;
    const text = input.value;
    const { r, c } = editing;
    if (text.startsWith('=') && text.length > 1) {
      let f = text.slice(1);
      // fehlende Klammern ergaenzen (wie Excel)
      const open = (f.match(/\(/g) || []).length - (f.match(/\)/g) || []).length;
      if (open > 0 && open < 5) f += ')'.repeat(open);
      const err = RF.check(f);
      if (err) {
        Rex.alert('Fehler in der Formel', `Die Formel „=${f}“ enthält einen Fehler: ${err}`).then(() => input.focus());
        return false;
      }
      setInput(r, c, '=' + RF.normalize(f));
    } else setInput(r, c, text);
    editing = null;
    input.value = '';
    input.classList.add('hidden');
    hideAC();
    $('#stMode').textContent = 'Bereit';
    commit();
    updateFormulaBar();
    return true;
  }
  function setInput(r, c, text, sh = S) {
    const old = sh.cells[key(r, c)];
    const p = parseInput(text, old && old.s);
    const cell = {};
    if (p.v !== undefined) cell.v = p.v;
    if (p.f !== undefined) cell.f = p.f;
    if (p.s && Object.keys(p.s).length) cell.s = p.s;
    if (old && old.link && p.v !== undefined) cell.link = old.link;
    if (Object.keys(cell).length) sh.cells[key(r, c)] = cell; else delete sh.cells[key(r, c)];
  }

  // Formel-Autovervollstaendigung
  let acItems = [], acIndex = 0, acStart = -1;
  function updateAC() {
    const v = input.value, pos = input.selectionStart;
    if (!v.startsWith('=')) return hideAC();
    const m = /([A-Za-zÄÖÜäöü][A-Za-zÄÖÜäöü0-9.]*)$/.exec(v.slice(0, pos));
    if (!m || (m.index > 0 && /[A-Za-z0-9$!]/.test(v[m.index - 1]))) return hideAC();
    const q = m[1].toUpperCase();
    const seen = new Set();
    acItems = [];
    for (const f of RF.FUNCTION_LIST) {
      for (const name of [f.de, f.en]) {
        if (name.startsWith(q) && !seen.has(name)) { seen.add(name); acItems.push({ name, other: name === f.de ? f.en : f.de }); }
      }
    }
    acItems.sort((a, b) => a.name.length - b.name.length || a.name.localeCompare(b.name));
    acItems = acItems.slice(0, 12);
    if (!acItems.length) return hideAC();
    acStart = m.index;
    acIndex = 0;
    const list = $('#acList');
    list.innerHTML = acItems.map((it, i) => `<div data-i="${i}" class="${i === 0 ? 'on' : ''}"><span>${it.name}</span><small>${it.other !== it.name ? it.other : ''}</small></div>`).join('');
    const ir = input.getBoundingClientRect(), wr = wrap.getBoundingClientRect();
    list.style.left = ir.left - wr.left + 'px';
    list.style.top = ir.bottom - wr.top + 2 + 'px';
    list.hidden = false;
  }
  function hideAC() { $('#acList').hidden = true; acItems = []; }
  function acceptAC() {
    const it = acItems[acIndex];
    if (!it) return;
    const v = input.value, pos = input.selectionStart;
    input.value = v.slice(0, acStart) + it.name + '(' + v.slice(pos);
    const np = acStart + it.name.length + 1;
    input.setSelectionRange(np, np);
    hideAC();
    fx.value = input.value;
    positionInput();
  }
  $('#acList').addEventListener('mousedown', e => {
    const d = e.target.closest('[data-i]');
    if (!d) return;
    e.preventDefault();
    acIndex = +d.dataset.i;
    acceptAC();
  });

  // Bezug beim Formel-Bearbeiten per Mausklick einfuegen
  let refInsert = null;
  function canInsertRef() {
    if (!editing) return false;
    const v = editing.fromBar ? fx.value : input.value;
    if (!v.startsWith('=')) return false;
    const el = editing.fromBar ? fx : input;
    const before = v.slice(0, el.selectionStart).trimEnd();
    if (refInsert && el.selectionStart === refInsert.end) return true;
    return /[=(;,+\-*/^&<>:]$/.test(before);
  }
  function insertRef(r1, c1, r2, c2) {
    const el = editing.fromBar ? fx : input;
    const v = el.value;
    const ref = r1 === r2 && c1 === c2 ? RF.addr(r1, c1) : RF.addr(Math.min(r1, r2), Math.min(c1, c2)) + ':' + RF.addr(Math.max(r1, r2), Math.max(c1, c2));
    let start = el.selectionStart;
    if (refInsert && el.selectionStart === refInsert.end) start = refInsert.start;
    el.value = v.slice(0, start) + ref + v.slice(el.selectionStart);
    refInsert = { start, end: start + ref.length };
    el.setSelectionRange(refInsert.end, refInsert.end);
    if (editing.fromBar) input.value = fx.value; else fx.value = input.value;
    positionInput();
    renderSelection();
  }

  // =====================================================================
  //  Formelleiste
  // =====================================================================
  function updateFormulaBar() {
    nameBox.value = selAddr();
    if (!editing) fx.value = inputText(getCell(sel.r, sel.c));
  }
  fx.addEventListener('focus', () => {
    if (!editing) { beginEdit('edit'); editing.fromBar = true; fx.focus(); }
    else editing.fromBar = true;
  });
  fx.addEventListener('input', () => { input.value = fx.value; positionInput(); renderSelection(); refInsert = null; });
  fx.addEventListener('keydown', e => {
    if (e.key === 'Enter') { e.preventDefault(); if (commitEdit()) { select(sel.r + 1, sel.c); focusGrid(); } }
    else if (e.key === 'Escape') { e.preventDefault(); cancelEdit(); focusGrid(); }
    else if (e.key === 'Tab') { e.preventDefault(); if (commitEdit()) { select(sel.r, sel.c + 1); focusGrid(); } }
  });
  nameBox.addEventListener('keydown', e => {
    if (e.key === 'Enter') {
      e.preventDefault();
      const v = nameBox.value.trim().toUpperCase();
      try {
        const rg = RF.refRange(v);
        if (!isFinite(rg.r1) || isNaN(rg.c1)) throw new Error();
        select(rg.r1, rg.c1);
        if (v.includes(':')) { anchor = { r: rg.r1, c: rg.c1 }; select(Math.min(rg.r2, MAX_ROWS - 1), rg.c2, true); }
      } catch { Rex.toast('Ungültiger Bezug'); }
      focusGrid();
    }
    if (e.key === 'Escape') { updateFormulaBar(); focusGrid(); }
  });
  function focusGrid() { if (!editing) input.focus({ preventScroll: true }); }

  // =====================================================================
  //  Tastatur
  // =====================================================================
  function dataEdge(r, c, dr, dc) {
    const filled = (rr, cc) => { const x = getCell(rr, cc); return !!(x && (x.v != null || x.f != null)); };
    const lim = (rr, cc) => rr >= 0 && cc >= 0 && rr < MAX_ROWS && cc < MAX_COLS;
    const { mr, mc } = usedExtent();
    if (!lim(r + dr, c + dc)) return { r, c };
    if (filled(r, c) && filled(r + dr, c + dc)) {
      while (lim(r + dr, c + dc) && filled(r + dr, c + dc)) { r += dr; c += dc; }
      return { r, c };
    }
    r += dr; c += dc;
    while (lim(r, c) && !filled(r, c)) {
      if ((dr > 0 && r > mr) || (dc > 0 && c > mc)) return { r: dr > 0 ? Math.min(MAX_ROWS - 1, Math.max(r, mr)) : r, c: dc > 0 ? Math.max(c, mc) : c };
      if (!lim(r + dr, c + dc)) break;
      r += dr; c += dc;
    }
    return { r, c };
  }
  function move(dr, dc, extend, jump) {
    let r = extend ? (sel.r1 === anchor.r ? sel.r2 : sel.r1) : sel.r;
    let c = extend ? (sel.c1 === anchor.c ? sel.c2 : sel.c1) : sel.c;
    if (!extend) {
      const m = mergeAt(sel.r, sel.c);
      if (m) { if (dr > 0) r = m.r2; if (dc > 0) c = m.c2; }
    }
    if (jump) ({ r, c } = dataEdge(r, c, dr, dc));
    else { r += dr; c += dc; while (dr && r > 0 && r < MAX_ROWS - 1 && rowH(r) === 0) r += dr; }
    select(r, c, extend);
  }
  // Nach Tab-Eingaben springt Enter zurueck in die Startspalte (wie Excel)
  let tabOrigin = null;
  function enterMove(up) {
    const c = tabOrigin != null && !up ? tabOrigin : sel.c;
    tabOrigin = null;
    const m = mergeAt(sel.r, sel.c);
    select((up ? sel.r - 1 : (m ? m.r2 : sel.r) + 1), c);
  }
  input.addEventListener('keydown', e => {
    const ctrl = e.ctrlKey || e.metaKey;
    if (!['Tab', 'Enter', 'Shift'].includes(e.key) && !editing) { if (/^Arrow|Home|End|Page/.test(e.key)) tabOrigin = null; }
    const k = e.key;
    if (editing) {
      if (acItems.length && !$('#acList').hidden) {
        if (k === 'ArrowDown' || k === 'ArrowUp') {
          e.preventDefault();
          acIndex = (acIndex + (k === 'ArrowDown' ? 1 : -1) + acItems.length) % acItems.length;
          $$('#acList div').forEach((d, i) => d.classList.toggle('on', i === acIndex));
          return;
        }
        if (k === 'Tab' || (k === 'Enter' && !e.altKey)) { e.preventDefault(); acceptAC(); return; }
        if (k === 'Escape') { e.preventDefault(); hideAC(); return; }
      }
      if (k === 'Enter' && e.altKey) {
        e.preventDefault();
        const p = input.selectionStart;
        input.value = input.value.slice(0, p) + '\n' + input.value.slice(input.selectionEnd);
        input.setSelectionRange(p + 1, p + 1);
        const cell = getCell(editing.r, editing.c);
        if (!cell || !cell.s || !cell.s.wrap) { S.cells[key(editing.r, editing.c)] = { ...(cell || {}), s: { ...((cell && cell.s) || {}), wrap: true } }; }
        positionInput();
        return;
      }
      if (k === 'Enter') { e.preventDefault(); if (commitEdit()) enterMove(e.shiftKey); return; }
      if (k === 'Tab') { e.preventDefault(); if (tabOrigin == null) tabOrigin = sel.c; if (commitEdit()) { move(0, e.shiftKey ? -1 : 1); } return; }
      if (k === 'Escape') { e.preventDefault(); cancelEdit(); return; }
      if (editing.mode === 'enter' && /^Arrow/.test(k) && !input.value.startsWith('=')) {
        e.preventDefault();
        if (commitEdit()) move(k === 'ArrowDown' ? 1 : k === 'ArrowUp' ? -1 : 0, k === 'ArrowRight' ? 1 : k === 'ArrowLeft' ? -1 : 0);
        return;
      }
      if (k === 'F4' && input.value.startsWith('=')) { e.preventDefault(); toggleAbsolute(); return; }
      return;
    }
    // Nicht im Bearbeitungsmodus
    if (selChart && (k === 'Delete' || k === 'Backspace')) {
      e.preventDefault();
      S.charts = S.charts.filter(x => x.id !== selChart);
      selChart = null;
      commit();
      return;
    }
    const arrows = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (arrows[k]) { e.preventDefault(); move(...arrows[k], e.shiftKey, ctrl); return; }
    if (k === 'Enter') { e.preventDefault(); enterMove(e.shiftKey); return; }
    if (k === 'Tab') { e.preventDefault(); if (tabOrigin == null) tabOrigin = sel.c; move(0, e.shiftKey ? -1 : 1); return; }
    if (k === 'PageDown' || k === 'PageUp') {
      e.preventDefault();
      const rows = Math.max(1, Math.floor(scroller.clientHeight / (DEF_H * zoom)) - 1);
      scroller.scrollTop += (k === 'PageDown' ? 1 : -1) * rows * DEF_H * zoom;
      move(k === 'PageDown' ? rows : -rows, 0, e.shiftKey);
      return;
    }
    if (k === 'Home') { e.preventDefault(); if (ctrl) select(0, 0, e.shiftKey); else select(sel.r, 0, e.shiftKey); return; }
    if (k === 'End' && ctrl) { e.preventDefault(); const u = usedExtent(); select(u.mr, u.mc, e.shiftKey); return; }
    if (k === 'Delete') { e.preventDefault(); clearRange('content'); return; }
    if (k === 'Backspace') { e.preventDefault(); beginEdit('enter', ''); return; }
    if (k === 'F2') { e.preventDefault(); beginEdit('edit'); return; }
    if (k === 'Escape') { if (clip) { clip = null; renderSelection(); } selChart = null; renderCharts(); return; }
    if (ctrl) {
      const kl = k.toLowerCase();
      if (kl === 'a') { e.preventDefault(); anchor = { r: 0, c: 0 }; sel = { r: sel.r, c: sel.c, r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }; render(); updateFormulaBar(); return; }
      if (kl === 'b') { e.preventDefault(); toggleStyle('b'); return; }
      if (kl === 'i') { e.preventDefault(); toggleStyle('i'); return; }
      if (kl === 'u') { e.preventDefault(); toggleStyle('u'); return; }
      if (kl === 'z') { e.preventDefault(); undo(); return; }
      if (kl === 'y') { e.preventDefault(); redo(); return; }
      if (kl === 'f') { e.preventDefault(); findDialog(); return; }
      if (kl === 'd') { e.preventDefault(); fillDirection('down'); return; }
      if (kl === 'r') { e.preventDefault(); fillDirection('right'); return; }
      if (k === ';' || k === ',') { e.preventDefault(); setInput(sel.r, sel.c, new Date().toLocaleDateString('de-DE')); commit(); return; }
      if (k === '+' ) { e.preventDefault(); insertRowsCols('row'); return; }
      if (k === '-') { e.preventDefault(); deleteRowsCols('row'); return; }
    }
    if (e.altKey && (k === '=' || k === '0')) { e.preventDefault(); autoSum('SUMME'); return; }
  });
  input.addEventListener('input', () => {
    if (!editing) {
      // erstes Zeichen beginnt die Eingabe (auch fuer Umlaute/Akzente)
      const v = input.value;
      editing = null;
      beginEdit('enter', v);
    }
    fx.value = input.value;
    refInsert = null;
    positionInput();
    renderSelection();
    updateAC();
  });
  function toggleAbsolute() {
    const v = input.value, pos = input.selectionStart;
    const re = /(\$?)([A-Za-z]{1,3})(\$?)(\d+)/g;
    let m;
    while ((m = re.exec(v))) {
      if (pos >= m.index && pos <= m.index + m[0].length) {
        // A1 -> $A$1 -> A$1 -> $A1 -> A1
        const state = (m[1] ? 2 : 0) + (m[3] ? 1 : 0);
        const ns = { 0: 3, 3: 1, 1: 2, 2: 0 }[state];
        const rep = (ns & 2 ? '$' : '') + m[2] + (ns & 1 ? '$' : '') + m[4];
        input.value = v.slice(0, m.index) + rep + v.slice(m.index + m[0].length);
        const np = m.index + rep.length;
        input.setSelectionRange(np, np);
        fx.value = input.value;
        return;
      }
    }
  }

  // =====================================================================
  //  Maus
  // =====================================================================
  function cellFromEvent(e) {
    const fr = S.freeze || { r: 0, c: 0 };
    const sr = scroller.getBoundingClientRect();
    let x = e.clientX - sr.left, y = e.clientY - sr.top;
    x = x < colX(fr.c) ? x : x + scroller.scrollLeft;
    y = y < rowY(fr.r) ? y : y + scroller.scrollTop;
    return { r: rowAt(Math.max(0, y)), c: colAt(Math.max(0, x)) };
  }
  let dragMode = null;
  scroller.addEventListener('mousedown', e => {
    if (e.button !== 0) return;
    if (e.target.closest('.chart')) return;
    if (e.target === input && editing) return;
    const fb = e.target.closest('.filter-btn');
    if (fb) { e.preventDefault(); filterMenu(+fb.dataset.fc, fb); return; }
    if (e.offsetX > scroller.clientWidth || e.offsetY > scroller.clientHeight) return; // Scrollbalken
    if (e.target.id === 'fillHandle') { e.preventDefault(); startFill(e); return; }
    e.preventDefault();
    selChart = null;
    const p = cellFromEvent(e);
    if (canInsertRef()) {
      const start = p;
      insertRef(p.r, p.c, p.r, p.c);
      dragMode = { type: 'ref', start };
    } else {
      if (editing && !commitEdit()) return;
      if (painter) { applyPainter(p); return; }
      tabOrigin = null;
      select(p.r, p.c, e.shiftKey);
      dragMode = { type: 'sel' };
      focusGrid();
    }
    const mv = ev => {
      const q = cellFromEvent(ev);
      if (dragMode && dragMode.type === 'sel') { if (q.r !== sel.r2 || q.c !== sel.c2) select(q.r, q.c, true); }
      else if (dragMode && dragMode.type === 'ref') insertRef(dragMode.start.r, dragMode.start.c, q.r, q.c);
    };
    const up = () => { document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up); dragMode = null; };
    document.addEventListener('mousemove', mv);
    document.addEventListener('mouseup', up);
  });
  scroller.addEventListener('dblclick', e => {
    if (e.target.closest('.chart')) return;
    const p = cellFromEvent(e);
    if (p.r !== sel.r || p.c !== sel.c) select(p.r, p.c);
    beginEdit('edit');
  });
  scroller.addEventListener('scroll', () => {
    const { r1, c1 } = visibleRange();
    if (r1 > extentRows - 30 && extentRows < MAX_ROWS) extentRows = Math.min(MAX_ROWS, extentRows + 200);
    if (c1 > extentCols - 5 && extentCols < MAX_COLS) extentCols = Math.min(MAX_COLS, extentCols + 20);
    render();
  });
  new ResizeObserver(() => render()).observe(scroller);
  wrap.addEventListener('wheel', e => {
    if (e.ctrlKey) { e.preventDefault(); setZoom(zoom + (e.deltaY < 0 ? 0.1 : -0.1)); return; }
    if (e.target.closest('.frozen-pane') || !scroller.contains(e.target)) {
      scroller.scrollTop += e.deltaY;
      scroller.scrollLeft += e.deltaX;
    }
  }, { passive: false });

  // Kopfzeilen: Auswahl und Groesse
  function headerDrag(e, axis) {
    const rz = e.target.closest('.rz');
    if (rz) {
      e.preventDefault();
      const idx = axis === 'col' ? +rz.dataset.rc : +rz.dataset.rr;
      const start = axis === 'col' ? e.clientX : e.clientY;
      const orig = axis === 'col' ? colW(idx) / zoom : rowH(idx) / zoom;
      const targets = axis === 'col' && sel.r1 === 0 && sel.r2 >= MAX_ROWS - 1 && idx >= sel.c1 && idx <= sel.c2 ? range(sel.c1, sel.c2)
        : axis === 'row' && sel.c1 === 0 && sel.c2 >= MAX_COLS - 1 && idx >= sel.r1 && idx <= sel.r2 ? range(sel.r1, sel.r2) : [idx];
      const mv = ev => {
        const d = ((axis === 'col' ? ev.clientX : ev.clientY) - start) / zoom;
        const v = Math.max(axis === 'col' ? 8 : 6, Math.round(orig + d));
        for (const t of targets) { if (axis === 'col') S.colW[t] = v; else S.rowH[t] = v; }
        render();
      };
      const up = () => { document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up); commit(); };
      document.addEventListener('mousemove', mv);
      document.addEventListener('mouseup', up);
      return;
    }
    const h = e.target.closest(axis === 'col' ? '.hc' : '.hr');
    if (!h) return;
    e.preventDefault();
    if (editing && !commitEdit()) return;
    const idx = axis === 'col' ? +h.dataset.c : +h.dataset.r;
    const selectTo = (i, extend) => {
      if (axis === 'col') {
        if (!extend) anchor = { r: 0, c: i };
        const c1 = Math.min(anchor.c, i), c2 = Math.max(anchor.c, i);
        sel = { r: 0, c: anchor.c, r1: 0, c1, r2: MAX_ROWS - 1, c2 };
      } else {
        if (!extend) anchor = { r: i, c: 0 };
        const r1 = Math.min(anchor.r, i), r2 = Math.max(anchor.r, i);
        sel = { r: anchor.r, c: 0, r1, c1: 0, r2, c2: MAX_COLS - 1 };
      }
      render(); updateFormulaBar(); updateToolbar();
    };
    selectTo(idx, e.shiftKey);
    focusGrid();
    const mv = ev => {
      const el = document.elementFromPoint(ev.clientX, ev.clientY);
      const hh = el && el.closest(axis === 'col' ? '.hc' : '.hr');
      if (hh) selectTo(axis === 'col' ? +hh.dataset.c : +hh.dataset.r, true);
    };
    const up = () => { document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up); };
    document.addEventListener('mousemove', mv);
    document.addEventListener('mouseup', up);
  }
  const range = (a, b) => { const o = []; for (let i = a; i <= b; i++) o.push(i); return o; };
  $('#colHead').addEventListener('mousedown', e => headerDrag(e, 'col'));
  $('#rowHead').addEventListener('mousedown', e => headerDrag(e, 'row'));
  $('#colHead').addEventListener('dblclick', e => { const rz = e.target.closest('.rz'); if (rz) autoFitCol(+rz.dataset.rc); });
  $('#rowHead').addEventListener('dblclick', e => { const rz = e.target.closest('.rz'); if (rz) { delete S.rowH[+rz.dataset.rr]; commit(); } });
  $('#corner').addEventListener('mousedown', e => { e.preventDefault(); anchor = { r: 0, c: 0 }; sel = { r: 0, c: 0, r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }; render(); updateFormulaBar(); focusGrid(); });

  function autoFitCol(c) {
    const cols = sel.r1 === 0 && sel.r2 >= MAX_ROWS - 1 && c >= sel.c1 && c <= sel.c2 ? range(sel.c1, sel.c2) : [c];
    const { mr } = usedExtent();
    for (const cc of cols) {
      let w = 20;
      for (let r = 0; r <= mr; r++) {
        const cell = getCell(r, cc);
        if (!cell || mergeAt(r, cc)) continue;
        const d = display(wb.active, r, cc);
        if (d && d.text) w = Math.max(w, (textWidth(d.text, cell.s) / zoom) + 12);
      }
      S.colW[cc] = Math.min(600, Math.ceil(w));
    }
    commit();
  }

  // Kontextmenue
  wrap.addEventListener('contextmenu', e => {
    e.preventDefault();
    if (editing) return;
    const onHeader = e.target.closest('.hc,.hr');
    if (!onHeader && scroller.contains(e.target) && !e.target.closest('.chart')) {
      const p = cellFromEvent(e);
      const { r1, c1, r2, c2 } = selRect();
      if (!(p.r >= r1 && p.r <= r2 && p.c >= c1 && p.c <= c2)) select(p.r, p.c);
    }
    const fake = { getBoundingClientRect: () => ({ left: e.clientX, right: e.clientX, top: e.clientY, bottom: e.clientY }) };
    const items = [
      { label: 'Ausschneiden', icon: 'scissors', action: () => doCopy(true) },
      { label: 'Kopieren', icon: 'copy', action: () => doCopy(false) },
      { label: 'Einfügen', icon: 'clipboard-paste', action: () => commands.paste() },
      '-',
      { label: 'Zeilen oberhalb einfügen', icon: 'between-horizontal-end', action: () => insertRowsCols('row') },
      { label: 'Spalten links einfügen', icon: 'between-vertical-end', action: () => insertRowsCols('col') },
      { label: 'Zeilen löschen', icon: 'rows-3', action: () => deleteRowsCols('row') },
      { label: 'Spalten löschen', icon: 'columns-3', action: () => deleteRowsCols('col') },
      '-',
      { label: 'Inhalte löschen', icon: 'eraser', action: () => clearRange('content') },
      { label: 'Zeilen ausblenden', icon: 'eye-off', action: () => { for (let r = sel.r1; r <= Math.min(sel.r2, sel.r1 + 10000); r++) S.hidden[r] = true; commit(); } },
      { label: 'Zeilen einblenden', icon: 'eye', action: () => { const a = Math.max(0, sel.r1 - 1), b = Math.min(sel.r2 + 1, sel.r1 + 10000); for (let r = a; r <= b; r++) delete S.hidden[r]; commit(); } },
      '-',
      { label: 'Diagramm aus Auswahl', icon: 'chart-column', action: () => insertChart('column') }
    ];
    Rex.menu(fake, items);
  });

  // =====================================================================
  //  Ausfuellen (Ausfuellkaestchen)
  // =====================================================================
  const SERIES = [
    ['Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag', 'Sonntag'],
    ['Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So'],
    ['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'],
    ['Jan', 'Feb', 'Mär', 'Apr', 'Mai', 'Jun', 'Jul', 'Aug', 'Sep', 'Okt', 'Nov', 'Dez'],
    ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'],
    ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
  ];
  function startFill(e) {
    const src = selRect();
    const prev = $('#fillPreview');
    let target = null;
    const mv = ev => {
      const q = cellFromEvent(ev);
      let t = { ...src };
      const dDown = q.r - src.r2, dUp = src.r1 - q.r, dRight = q.c - src.c2, dLeft = src.c1 - q.c;
      const best = Math.max(dDown, dUp, dRight, dLeft);
      if (best <= 0) { prev.hidden = true; target = null; return; }
      if (best === dDown) t.r2 = q.r; else if (best === dUp) t.r1 = q.r; else if (best === dRight) t.c2 = q.c; else t.c1 = q.c;
      target = t;
      const x = colX(t.c1), y = rowY(t.r1);
      Object.assign(prev.style, { left: x + 'px', top: y + 'px', width: colX(t.c2 + 1) - x + 'px', height: rowY(t.r2 + 1) - y + 'px' });
      prev.hidden = false;
    };
    const up = () => {
      document.removeEventListener('mousemove', mv);
      document.removeEventListener('mouseup', up);
      prev.hidden = true;
      if (target) {
        doFill(src, target);
        sel = { ...sel, r1: target.r1, c1: target.c1, r2: target.r2, c2: target.c2 };
        commit();
        updateFormulaBar();
      }
    };
    document.addEventListener('mousemove', mv);
    document.addEventListener('mouseup', up);
  }
  function doFill(src, t) {
    const vertical = t.r1 !== src.r1 || t.r2 !== src.r2;
    const forward = vertical ? t.r2 > src.r2 : t.c2 > src.c2;
    const lines = vertical ? range(src.c1, src.c2) : range(src.r1, src.r2);
    for (const line of lines) {
      const srcCells = (vertical ? range(src.r1, src.r2) : range(src.c1, src.c2)).map(i => {
        const r = vertical ? i : line, c = vertical ? line : i;
        return { r, c, cell: getCell(r, c) };
      });
      const n = srcCells.length;
      const vals = srcCells.map(x => x.cell && x.cell.f == null ? x.cell.v : undefined);
      const allNum = vals.every(v => typeof v === 'number');
      let step = 0;
      if (allNum && n >= 2) step = (vals[n - 1] - vals[0]) / (n - 1);
      const targetIdx = vertical ? (forward ? range(src.r2 + 1, t.r2) : range(t.r1, src.r1 - 1).reverse())
        : (forward ? range(src.c2 + 1, t.c2) : range(t.c1, src.c1 - 1).reverse());
      targetIdx.forEach((ti, k) => {
        const srcI = forward ? k % n : (n - 1 - (k % n));
        const s = srcCells[srcI];
        const r = vertical ? ti : line, c = vertical ? line : ti;
        if (!s.cell) { delete S.cells[key(r, c)]; return; }
        const nc = JSON.parse(JSON.stringify(s.cell));
        const off = (k + 1) * (forward ? 1 : -1);
        if (nc.f != null) {
          nc.f = RF.shiftFormula(s.cell.f, r - s.r, c - s.c);
        } else if (allNum && n >= 2) {
          nc.v = (forward ? vals[n - 1] : vals[0]) + step * off;
          nc.v = Math.round(nc.v * 1e10) / 1e10;
        } else if (allNum && n === 1 && nc.s && nc.s.fmt && FMT.isDateFormat(nc.s.fmt)) {
          nc.v = vals[0] + off;
        } else if (typeof nc.v === 'string') {
          const base = forward ? srcCells[n - 1].cell : srcCells[0].cell;
          const bv = base && typeof base.v === 'string' ? base.v : nc.v;
          let done = false;
          for (const list of SERIES) {
            const i = list.findIndex(x => x.toLowerCase() === bv.toLowerCase());
            if (i >= 0) {
              const steps = Math.floor(k / 1) + 1;
              nc.v = list[((i + (forward ? steps : -steps)) % list.length + list.length) % list.length];
              done = true; break;
            }
          }
          if (!done && n === 1) {
            const m = /^(.*?)(\d+)(\D*)$/.exec(bv);
            if (m) nc.v = m[1] + Math.max(0, +m[2] + off) + m[3];
          }
        }
        S.cells[key(r, c)] = nc;
      });
    }
  }
  function fillDirection(dir) {
    const src = selRect();
    if (dir === 'down') {
      if (src.r1 === src.r2) { if (src.r1 === 0) return; doFill({ ...src, r1: src.r1 - 1, r2: src.r1 - 1 }, src); }
      else doFill({ ...src, r2: src.r1 }, src);
    } else {
      if (src.c1 === src.c2) { if (src.c1 === 0) return; doFill({ ...src, c1: src.c1 - 1, c2: src.c1 - 1 }, src); }
      else doFill({ ...src, c2: src.c1 }, src);
    }
    commit();
  }

  // =====================================================================
  //  Zwischenablage
  // =====================================================================
  let clip = null;
  function clipRange() {
    const { r1, c1, r2, c2 } = selRect();
    const u = usedExtent();
    return { r1, c1, r2: Math.min(r2, Math.max(r1, u.mr)), c2: Math.min(c2, Math.max(c1, u.mc)) };
  }
  function buildCopy(cut) {
    const rg = clipRange();
    const cells = [], rowsText = [];
    let html = '<table>';
    for (let r = rg.r1; r <= rg.r2; r++) {
      const row = [], txt = [];
      html += '<tr>';
      for (let c = rg.c1; c <= rg.c2; c++) {
        const cell = getCell(r, c);
        row.push(cell ? JSON.parse(JSON.stringify(cell)) : null);
        const d = display(wb.active, r, c);
        txt.push(d ? d.text : '');
        html += `<td style="${cellStyleCss(cell && cell.s).replace(/"/g, "'")}">${esc(d ? d.text : '')}</td>`;
      }
      html += '</tr>';
      cells.push(row);
      rowsText.push(txt.join('\t'));
    }
    html += '</table>';
    const text = rowsText.join('\n');
    clip = { ...rg, cells, text, cut, si: wb.active, merges: S.merges.filter(m => m.r1 >= rg.r1 && m.r2 <= rg.r2 && m.c1 >= rg.c1 && m.c2 <= rg.c2) };
    renderSelection();
    return { text, html };
  }
  function doCopy(cut) {
    const { text, html } = buildCopy(cut);
    try {
      navigator.clipboard.write([new ClipboardItem({ 'text/plain': new Blob([text], { type: 'text/plain' }), 'text/html': new Blob([html], { type: 'text/html' }) })]);
    } catch { try { navigator.clipboard.writeText(text); } catch { /* */ } }
  }
  document.addEventListener('copy', e => {
    if (document.activeElement !== input || editing) return;
    e.preventDefault();
    const { text, html } = buildCopy(false);
    e.clipboardData.setData('text/plain', text);
    e.clipboardData.setData('text/html', html);
  });
  document.addEventListener('cut', e => {
    if (document.activeElement !== input || editing) return;
    e.preventDefault();
    const { text, html } = buildCopy(true);
    e.clipboardData.setData('text/plain', text);
    e.clipboardData.setData('text/html', html);
  });
  document.addEventListener('paste', e => {
    if (document.activeElement !== input || editing) return;
    e.preventDefault();
    pasteText(e.clipboardData.getData('text/plain'));
  });
  function pasteText(text) {
    const t0 = sel.r1, c0 = sel.c1;
    if (clip && text != null && text.replace(/\r/g, '') === clip.text) {
      const srcSheet = wb.sheets[clip.si];
      const rows = clip.cells.length, cols = clip.cells[0].length;
      // Bei groesserer Auswahl kacheln
      const selRows = sel.r2 - sel.r1 + 1, selCols = sel.c2 - sel.c1 + 1;
      const repR = selRows % rows === 0 && selRows < 5000 ? selRows / rows : 1;
      const repC = selCols % cols === 0 && selCols < 500 ? selCols / cols : 1;
      if (clip.cut) {
        for (let r = clip.r1; r <= clip.r2; r++) for (let c = clip.c1; c <= clip.c2; c++) delete srcSheet.cells[key(r, c)];
        srcSheet.merges = srcSheet.merges.filter(m => !clip.merges.includes(m));
      }
      for (let rr = 0; rr < repR; rr++) for (let cc = 0; cc < repC; cc++) {
        for (let i = 0; i < rows; i++) for (let j = 0; j < cols; j++) {
          const src = clip.cells[i][j];
          const r = t0 + rr * rows + i, c = c0 + cc * cols + j;
          if (!src) { delete S.cells[key(r, c)]; continue; }
          const nc = JSON.parse(JSON.stringify(src));
          if (nc.f != null && !clip.cut) nc.f = RF.shiftFormula(nc.f, r - (clip.r1 + i), c - (clip.c1 + j));
          S.cells[key(r, c)] = nc;
        }
        for (const m of clip.merges) {
          const dr = t0 + rr * rows - clip.r1, dc = c0 + cc * cols - clip.c1;
          S.merges.push({ r1: m.r1 + dr, c1: m.c1 + dc, r2: m.r2 + dr, c2: m.c2 + dc });
        }
      }
      sel = { ...sel, r1: t0, c1: c0, r2: t0 + rows * repR - 1, c2: c0 + cols * repC - 1 };
      if (clip.cut) clip = null;
      else clip.done = true;
    } else if (text) {
      const rows = text.replace(/\r/g, '').replace(/\n$/, '').split('\n').map(l => l.split('\t'));
      rows.forEach((row, i) => row.forEach((v, j) => setInput(t0 + i, c0 + j, v)));
      sel = { ...sel, r1: t0, c1: c0, r2: t0 + rows.length - 1, c2: c0 + Math.max(...rows.map(r => r.length)) - 1 };
    }
    commit();
    updateFormulaBar();
  }

  // =====================================================================
  //  Format-Befehle
  // =====================================================================
  function styleAt(r, c) { const x = getCell(r, c); return (x && x.s) || {}; }
  function setStyle(fn) {
    const { r1, c1, r2, c2 } = selRect();
    const full = r2 - r1 > 5000 || c2 - c1 > 500;
    const u = usedExtent();
    const rr2 = full ? Math.min(r2, u.mr) : r2, cc2 = full ? Math.min(c2, u.mc) : c2;
    for (let r = r1; r <= rr2; r++) for (let c = c1; c <= cc2; c++) {
      const k = key(r, c);
      const cell = S.cells[k] || {};
      const s = { ...(cell.s || {}) };
      fn(s, r, c);
      for (const kk of Object.keys(s)) if (s[kk] == null || s[kk] === false || s[kk] === '') delete s[kk];
      if (Object.keys(s).length) cell.s = s; else delete cell.s;
      if (Object.keys(cell).length) S.cells[k] = cell; else delete S.cells[k];
    }
    commit();
    updateToolbar();
  }
  function toggleStyle(p) {
    const on = !styleAt(sel.r, sel.c)[p];
    setStyle(s => { s[p] = on; });
  }
  function setBorders(kind) {
    const { r1, c1, r2, c2 } = selRect();
    const w = kind === 'thickOuter' ? 'medium' : 'thin';
    setStyle((s, r, c) => {
      const bd = { ...(s.bd || {}) };
      if (kind === 'none') { s.bd = null; return; }
      if (kind === 'all') { bd.t = bd.b = bd.l = bd.r = 'thin'; }
      if (kind === 'outer' || kind === 'thickOuter') {
        if (r === r1) bd.t = w; if (r === r2) bd.b = w; if (c === c1) bd.l = w; if (c === c2) bd.r = w;
      }
      if (kind === 'bottom' && r === r2) bd.b = 'thin';
      if (kind === 'top' && r === r1) bd.t = 'thin';
      if (kind === 'left' && c === c1) bd.l = 'thin';
      if (kind === 'right' && c === c2) bd.r = 'thin';
      if (kind === 'doubleBottom' && r === r2) bd.b = 'thick';
      s.bd = Object.keys(bd).length ? bd : null;
    });
  }
  function changeDecimals(delta) {
    const cur = styleAt(sel.r, sel.c).fmt;
    const v = getValue(wb.active, sel.r, sel.c);
    setStyle((s) => {
      let f = s.fmt || cur;
      if (!f || f === 'General') {
        const dec = typeof v === 'number' ? Math.min(10, ((RF.fmtGeneral(v).split(',')[1]) || '').length) : 0;
        const nd = Math.max(0, dec + delta);
        s.fmt = nd ? '0.' + '0'.repeat(nd) : '0';
        return;
      }
      s.fmt = f.split(';').map(sec => {
        if (/\.0+/.test(sec)) return delta > 0 ? sec.replace(/\.(0+)/, '.$10') : sec.replace(/\.(0+)/, (m, z) => z.length > 1 ? '.' + z.slice(1) : '');
        if (delta > 0) return sec.replace(/(0)(?![\d.])/, '$1.0');
        return sec;
      }).join(';');
    });
  }
  function toggleMerge() {
    const { r1, c1, r2, c2 } = selRect();
    const existing = S.merges.filter(m => m.r1 >= r1 && m.r2 <= r2 && m.c1 >= c1 && m.c2 <= c2);
    if (existing.length) {
      S.merges = S.merges.filter(m => !existing.includes(m));
      commit();
      return;
    }
    if (r1 === r2 && c1 === c2) return;
    if (r2 - r1 > 1000 || c2 - c1 > 100) { Rex.toast('Bereich zu groß zum Verbinden'); return; }
    let others = false;
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) if ((r !== r1 || c !== c1) && getCell(r, c) && (getCell(r, c).v != null || getCell(r, c).f != null)) others = true;
    const go = () => {
      for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) if (r !== r1 || c !== c1) delete S.cells[key(r, c)];
      S.merges = S.merges.filter(m => !(m.r2 >= r1 && m.r1 <= r2 && m.c2 >= c1 && m.c1 <= c2));
      S.merges.push({ r1, c1, r2, c2 });
      const cell = S.cells[key(r1, c1)] || {};
      cell.s = { ...(cell.s || {}), al: 'center' };
      S.cells[key(r1, c1)] = cell;
      sel = { r: r1, c: c1, r1, c1, r2, c2 };
      commit();
    };
    if (others) Rex.confirm('Zellen verbinden', 'Beim Verbinden bleibt nur der Wert oben links erhalten. Fortfahren?').then(ok => ok && go());
    else go();
  }
  function clearRange(what) {
    const { r1, c1, r2, c2 } = selRect();
    const u = usedExtent();
    for (let r = r1; r <= Math.min(r2, u.mr); r++) for (let c = c1; c <= Math.min(c2, u.mc); c++) {
      const k = key(r, c);
      const cell = S.cells[k];
      if (!cell) continue;
      if (what === 'all') delete S.cells[k];
      else if (what === 'format') { delete cell.s; if (!Object.keys(cell).length) delete S.cells[k]; }
      else { delete cell.v; delete cell.f; delete cell.c; delete cell.link; if (!Object.keys(cell).length) delete S.cells[k]; }
    }
    if (what === 'all') S.merges = S.merges.filter(m => !(m.r1 >= r1 && m.r2 <= r2 && m.c1 >= c1 && m.c2 <= c2));
    commit();
    updateFormulaBar();
  }

  // Format uebertragen
  let painter = null;
  function applyPainter(p) {
    const src = painter;
    painter = null;
    $('#painterBtn').classList.remove('on');
    select(p.r, p.c);
    const rows = src.length, cols = src[0].length;
    for (let i = 0; i < rows; i++) for (let j = 0; j < cols; j++) {
      const k = key(p.r + i, p.c + j);
      const cell = S.cells[k] || {};
      if (src[i][j]) cell.s = { ...src[i][j] }; else delete cell.s;
      if (Object.keys(cell).length) S.cells[k] = cell; else delete S.cells[k];
    }
    sel = { ...sel, r2: p.r + rows - 1, c2: p.c + cols - 1 };
    commit();
  }

  // =====================================================================
  //  Zeilen/Spalten einfuegen und loeschen
  // =====================================================================
  function shiftStructure(axis, at, count) {
    const sheetName = S.name;
    // Zellen verschieben
    const newCells = {};
    for (const [k, cell] of Object.entries(S.cells)) {
      let [r, c] = k.split(',').map(Number);
      const v = axis === 'row' ? r : c;
      if (count < 0 && v >= at && v < at - count) continue;
      if (v >= at) { if (axis === 'row') r += count; else c += count; }
      newCells[key(r, c)] = cell;
    }
    S.cells = newCells;
    const shiftMap = (obj) => {
      const out = {};
      for (const [k, val] of Object.entries(obj)) {
        const i = +k;
        if (count < 0 && i >= at && i < at - count) continue;
        out[i >= at ? i + count : i] = val;
      }
      return out;
    };
    if (axis === 'row') { S.rowH = shiftMap(S.rowH); S.hidden = shiftMap(S.hidden); }
    else S.colW = shiftMap(S.colW);
    S.merges = S.merges.map(m => {
      const a = axis === 'row' ? 'r' : 'c';
      const n = { ...m };
      for (const p of [a + '1', a + '2']) {
        if (count > 0 && n[p] >= at) n[p] += count;
        if (count < 0) { if (n[p] >= at - count) n[p] += count; else if (n[p] >= at) n[p] = p.endsWith('1') ? at : at - 1; }
      }
      return n[a + '2'] < n[a + '1'] ? null : n;
    }).filter(Boolean);
    // Formeln in allen Tabellen anpassen
    for (const sh of wb.sheets) {
      for (const cell of Object.values(sh.cells)) {
        if (cell.f != null) cell.f = RF.adjustForInsert(cell.f, axis, at, count, sheetName, sh.name);
      }
      for (const ch of sh.charts) for (const s of ch.series) for (const p of ['name', 'cat', 'val']) {
        if (s[p] && !s[p].startsWith('"')) s[p] = RF.adjustForInsert(s[p], axis, at, count, sheetName, sh.name);
      }
    }
    if (S.filter) {
      const a = axis === 'row' ? 'r' : 'c';
      for (const p of [a + '1', a + '2']) if (S.filter[p] >= at) S.filter[p] = Math.max(at, S.filter[p] + count);
    }
  }
  function insertRowsCols(axis) {
    const { r1, c1, r2, c2 } = selRect();
    const n = axis === 'row' ? Math.min(r2 - r1 + 1, 1000) : Math.min(c2 - c1 + 1, 100);
    shiftStructure(axis, axis === 'row' ? r1 : c1, n);
    commit();
  }
  function deleteRowsCols(axis) {
    const { r1, c1, r2, c2 } = selRect();
    const n = axis === 'row' ? Math.min(r2, Math.max(r1, usedExtent().mr)) - r1 + 1 : Math.min(c2, Math.max(c1, usedExtent().mc)) - c1 + 1;
    shiftStructure(axis, axis === 'row' ? r1 : c1, -n);
    select(sel.r1, sel.c1);
    commit();
  }

  // =====================================================================
  //  Sortieren, AutoSumme, Duplikate, Suchen
  // =====================================================================
  function sortRange(rg, byCol, asc) {
    const rows = [];
    for (let r = rg.r1; r <= rg.r2; r++) {
      if (S.hidden[r]) continue;
      const row = [];
      for (let c = rg.c1; c <= rg.c2; c++) row.push(getCell(r, c));
      rows.push({ row, v: getValue(wb.active, r, byCol), r });
    }
    rows.sort((a, b) => {
      const x = a.v, y = b.v;
      const ex = x == null || x === '', ey = y == null || y === '';
      if (ex && ey) return 0;
      if (ex) return 1;
      if (ey) return -1;
      let cmp;
      if (typeof x === 'number' && typeof y === 'number') cmp = x - y;
      else if (typeof x === 'number') cmp = -1;
      else if (typeof y === 'number') cmp = 1;
      else cmp = String(x).localeCompare(String(y), 'de', { numeric: true, sensitivity: 'base' });
      return asc ? cmp : -cmp;
    });
    const targets = range(rg.r1, rg.r2).filter(r => !S.hidden[r]);
    targets.forEach((r, i) => {
      rows[i].row.forEach((cell, j) => {
        const k = key(r, rg.c1 + j);
        if (cell) {
          const nc = JSON.parse(JSON.stringify(cell));
          if (nc.f != null) nc.f = RF.shiftFormula(nc.f, r - rows[i].r, 0);
          S.cells[k] = nc;
        } else delete S.cells[k];
      });
    });
    commit();
  }
  function sortSel(asc) {
    let rg = selRect();
    if (rg.r1 === rg.r2 && rg.c1 === rg.c2) {
      rg = currentRegion(sel.r, sel.c);
      // Ueberschriftenzeile erkennen
      const first = getValue(wb.active, rg.r1, sel.c), second = getValue(wb.active, rg.r1 + 1, sel.c);
      const allText = range(rg.c1, rg.c2).every(c => { const v = getValue(wb.active, rg.r1, c); return v == null || typeof v === 'string'; });
      if (allText && rg.r2 > rg.r1 && (typeof second === 'number' || (typeof first === 'string' && styleAt(rg.r1, sel.c).b))) rg.r1++;
    } else if (rg.r2 > 100000) rg.r2 = usedExtent().mr;
    if (S.filter && rg.r1 <= S.filter.r1 && rg.r2 >= S.filter.r1) rg.r1 = S.filter.r1 + 1;
    sortRange(rg, sel.c, asc);
  }
  function autoSum(fn) {
    const { r1, c1, r2, c2 } = selRect();
    const isNum = (r, c) => typeof getValue(wb.active, r, c) === 'number';
    if (r1 === r2 && c1 === c2) {
      let r = r1 - 1;
      while (r >= 0 && isNum(r, c1)) r--;
      if (r < r1 - 1) { setInput(r1, c1, `=${fn}(${RF.addr(r + 1, c1)}:${RF.addr(r1 - 1, c1)})`); commit(); updateFormulaBar(); return; }
      let c = c1 - 1;
      while (c >= 0 && isNum(r1, c)) c--;
      if (c < c1 - 1) { setInput(r1, c1, `=${fn}(${RF.addr(r1, c + 1)}:${RF.addr(r1, c1 - 1)})`); commit(); updateFormulaBar(); return; }
      beginEdit('enter', `=${fn}()`);
      input.setSelectionRange(fn.length + 2, fn.length + 2);
      return;
    }
    const rr2 = Math.min(r2, usedExtent().mr);
    if (rr2 > r1 || c1 === c2) {
      for (let c = c1; c <= c2; c++) setInput(rr2 + 1, c, `=${fn}(${RF.addr(r1, c)}:${RF.addr(rr2, c)})`);
    } else {
      setInput(r1, c2 + 1, `=${fn}(${RF.addr(r1, c1)}:${RF.addr(r1, c2)})`);
    }
    commit();
  }
  function dedupe() {
    let rg = selRect();
    if (rg.r1 === rg.r2) rg = currentRegion(sel.r, sel.c);
    if (rg.r2 > 100000) rg.r2 = usedExtent().mr;
    const seen = new Set();
    const keep = [];
    let removed = 0;
    for (let r = rg.r1; r <= rg.r2; r++) {
      const sig = range(rg.c1, rg.c2).map(c => JSON.stringify(getValue(wb.active, r, c))).join('|');
      if (seen.has(sig)) { removed++; continue; }
      seen.add(sig);
      keep.push(range(rg.c1, rg.c2).map(c => getCell(r, c)));
    }
    for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) delete S.cells[key(r, c)];
    keep.forEach((row, i) => row.forEach((cell, j) => { if (cell) S.cells[key(rg.r1 + i, rg.c1 + j)] = cell; }));
    commit();
    Rex.toast(`${removed} doppelte Zeile(n) entfernt, ${keep.length} eindeutige verbleiben.`);
  }
  async function textToCols() {
    const v = await Rex.dialog({ title: 'Text in Spalten', fields: [{ name: 'd', label: 'Trennzeichen', type: 'select', value: ';', options: [
      { value: ';', label: 'Semikolon (;)' }, { value: ',', label: 'Komma (,)' }, { value: ' ', label: 'Leerzeichen' }, { value: '\t', label: 'Tabulator' }, { value: '-', label: 'Bindestrich (-)' }] }] });
    if (!v) return;
    const { r1, c1, r2 } = selRect();
    for (let r = r1; r <= Math.min(r2, usedExtent().mr); r++) {
      const x = getValue(wb.active, r, c1);
      if (typeof x !== 'string') continue;
      x.split(v.d).forEach((part, j) => setInput(r, c1 + j, part.trim()));
    }
    commit();
  }
  async function findDialog() {
    const v = await Rex.dialog({ title: 'Suchen und Ersetzen', ok: 'Weitersuchen', fields: [
      { name: 'q', label: 'Suchen nach', value: lastFind },
      { name: 'rep', label: 'Ersetzen durch (leer lassen zum nur Suchen)', value: '' },
      { name: 'all', label: 'Alle ersetzen', type: 'checkbox', value: false }] });
    if (!v || !v.q) return;
    lastFind = v.q;
    const q = v.q.toLowerCase();
    const hits = Object.keys(S.cells).map(k => k.split(',').map(Number)).filter(([r, c]) => {
      const cell = getCell(r, c);
      const d = display(wb.active, r, c);
      return (d && d.text.toLowerCase().includes(q)) || (cell.f != null && cell.f.toLowerCase().includes(q));
    }).sort((a, b) => a[0] - b[0] || a[1] - b[1]);
    if (!hits.length) { Rex.toast('Nichts gefunden'); return; }
    if (v.all && v.rep !== undefined && v.rep !== '') {
      let n = 0;
      for (const [r, c] of hits) {
        const cell = getCell(r, c);
        if (cell.f == null && typeof cell.v === 'string') {
          cell.v = cell.v.replace(new RegExp(v.q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi'), v.rep);
          n++;
        }
      }
      commit();
      Rex.toast(`${n} Zelle(n) ersetzt`);
      return;
    }
    const next = hits.find(([r, c]) => r > sel.r || (r === sel.r && c > sel.c)) || hits[0];
    select(next[0], next[1]);
    Rex.toast(`${hits.length} Treffer`);
  }
  let lastFind = '';

  // =====================================================================
  //  Tabellenblaetter
  // =====================================================================
  function renderTabs() {
    $('#sheetTabs').innerHTML = wb.sheets.map((sh, i) => `<button class="${i === wb.active ? 'active' : ''}" data-si="${i}">${esc(sh.name)}</button>`).join('');
  }
  function switchSheet(i) {
    if (editing && !commitEdit()) return;
    wb.active = i;
    S = wb.sheets[i];
    selChart = null;
    sel = { r: 0, c: 0, r1: 0, c1: 0, r2: 0, c2: 0 };
    anchor = { r: 0, c: 0 };
    scroller.scrollTop = 0; scroller.scrollLeft = 0;
    extentRows = 200; extentCols = 40;
    renderTabs();
    render();
    updateFormulaBar();
    updateToolbar();
    focusGrid();
  }
  function renameSheetRefs(oldName, newName) {
    const q = (n) => /^[A-Za-z_À-ɏ][\w.À-ɏ]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`;
    const fix = (src) => {
      let toks;
      try { toks = RF.tokenize(src); } catch { return src; }
      let out = '', pos = 0;
      for (const t of toks) {
        const at = src.indexOf(t.raw, pos);
        out += src.slice(pos, at);
        out += t.t === 'ref' && t.sheet != null && t.sheet.toLowerCase() === oldName.toLowerCase() ? q(newName) + '!' + t.v : t.raw;
        pos = at + t.raw.length;
      }
      return out + src.slice(pos);
    };
    for (const sh of wb.sheets) {
      for (const cell of Object.values(sh.cells)) if (cell.f != null) cell.f = fix(cell.f);
      for (const ch of sh.charts) for (const s of ch.series) for (const p of ['name', 'cat', 'val']) if (s[p] && !s[p].startsWith('"')) s[p] = fix(s[p]);
    }
  }
  async function renameSheet(i) {
    const v = await Rex.dialog({ title: 'Tabellenblatt umbenennen', fields: [{ name: 'n', label: 'Name', value: wb.sheets[i].name }] });
    if (!v || !v.n.trim()) return;
    const name = v.n.trim().replace(/[\\/?*[\]:]/g, '').slice(0, 31);
    if (wb.sheets.some((s, j) => j !== i && s.name.toLowerCase() === name.toLowerCase())) { Rex.toast('Dieser Name ist schon vergeben.'); return; }
    renameSheetRefs(wb.sheets[i].name, name);
    wb.sheets[i].name = name;
    renderTabs();
    commit();
  }
  function uniqueSheetName(base) {
    let n = 1, name = base;
    while (wb.sheets.some(s => s.name.toLowerCase() === name.toLowerCase())) name = base.replace(/\d*$/, '') + (++n);
    return name;
  }
  $('#sheetTabs').addEventListener('click', e => { const b = e.target.closest('[data-si]'); if (b) switchSheet(+b.dataset.si); });
  $('#sheetTabs').addEventListener('dblclick', e => { const b = e.target.closest('[data-si]'); if (b) renameSheet(+b.dataset.si); });
  $('#sheetTabs').addEventListener('contextmenu', e => {
    const b = e.target.closest('[data-si]');
    if (!b) return;
    e.preventDefault();
    e.stopPropagation();
    const i = +b.dataset.si;
    Rex.menu(b, [
      { label: 'Umbenennen', icon: 'pencil', action: () => renameSheet(i) },
      { label: 'Duplizieren', icon: 'copy', action: () => {
        const copy = JSON.parse(JSON.stringify(wb.sheets[i]));
        copy.name = uniqueSheetName(wb.sheets[i].name + ' (2)');
        copy.charts.forEach(ch => { ch.id = 'ch' + Math.random().toString(36).slice(2, 8); });
        wb.sheets.splice(i + 1, 0, copy);
        commit(); switchSheet(i + 1);
      } },
      { label: 'Nach links', icon: 'arrow-left', action: () => { if (i > 0) { const [s] = wb.sheets.splice(i, 1); wb.sheets.splice(i - 1, 0, s); commit(); switchSheet(i - 1); } } },
      { label: 'Nach rechts', icon: 'arrow-right', action: () => { if (i < wb.sheets.length - 1) { const [s] = wb.sheets.splice(i, 1); wb.sheets.splice(i + 1, 0, s); commit(); switchSheet(i + 1); } } },
      '-',
      { label: 'Löschen', icon: 'trash-2', action: async () => {
        if (wb.sheets.length === 1) { Rex.toast('Die letzte Tabelle kann nicht gelöscht werden.'); return; }
        if (!(await Rex.confirm('Tabellenblatt löschen', `„${wb.sheets[i].name}“ wirklich löschen?`))) return;
        wb.sheets.splice(i, 1);
        commit(); switchSheet(Math.max(0, i - 1));
      } }
    ]);
  });

  // =====================================================================
  //  Statusleiste und Symbolleiste
  // =====================================================================
  function updateStatus() {
    const { r1, c1, r2, c2 } = selRect();
    if (r1 === r2 && c1 === c2) { $('#stAgg').textContent = ''; return; }
    let sum = 0, cnt = 0, num = 0;
    const u = usedExtent();
    const rr2 = Math.min(r2, u.mr), cc2 = Math.min(c2, u.mc);
    if ((rr2 - r1 + 1) * (cc2 - c1 + 1) > 200000) { $('#stAgg').textContent = ''; return; }
    for (let r = r1; r <= rr2; r++) for (let c = c1; c <= cc2; c++) {
      if (S.hidden[r]) continue;
      const v = getValue(wb.active, r, c);
      if (v == null || v === '') continue;
      cnt++;
      if (typeof v === 'number') { sum += v; num++; }
    }
    const f = (n) => n.toLocaleString('de-DE', { maximumFractionDigits: 10 });
    $('#stAgg').textContent = num ? `Mittelwert: ${f(sum / num)}     Anzahl: ${cnt}     Summe: ${f(sum)}` : (cnt ? `Anzahl: ${cnt}` : '');
  }
  const FONTS = ['Calibri', 'Aptos', 'Arial', 'Cambria', 'Consolas', 'Courier New', 'Georgia', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana'];
  const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 28, 36, 48, 72];
  const NUMFMTS = [
    { v: '', l: 'Standard' }, { v: '0.00', l: 'Zahl' }, { v: '#,##0.00', l: 'Zahl mit Trennzeichen' }, { v: '#,##0.00 "€"', l: 'Währung (€)' },
    { v: '_-* #,##0.00 "€"_-;-* #,##0.00 "€"_-;_-* "-"?? "€"_-;_-@_-', l: 'Buchhaltung' }, { v: 'DD.MM.YYYY', l: 'Datum, kurz' },
    { v: 'DDDD, D. MMMM YYYY', l: 'Datum, lang' }, { v: 'hh:mm:ss', l: 'Uhrzeit' }, { v: '0.00%', l: 'Prozent' }, { v: '0.00E+00', l: 'Wissenschaftlich' }, { v: '@', l: 'Text' }
  ];
  function updateToolbar() {
    const s = styleAt(sel.r, sel.c);
    $$('[data-style]').forEach(b => b.classList.toggle('on', !!s[b.dataset.style]));
    $$('[data-align]').forEach(b => b.classList.toggle('on', s.al === b.dataset.align));
    $$('[data-valign]').forEach(b => b.classList.toggle('on', (s.va || 'bottom') === b.dataset.valign));
    $('#wrapBtn').classList.toggle('on', !!s.wrap);
    $('#mergeBtn').classList.toggle('on', !!mergeAt(sel.r, sel.c));
    const ff = $('#fontFamily'), fs = $('#fontSize'), nf = $('#numFmt');
    const fam = s.font || 'Calibri';
    if (![...ff.options].some(o => o.value === fam)) ff.add(new Option(fam, fam));
    ff.value = fam;
    const sz = String(s.size || 11);
    if (![...fs.options].some(o => o.value === sz)) fs.add(new Option(sz, sz));
    fs.value = sz;
    const fmt = s.fmt || '';
    if (![...nf.options].some(o => o.value === fmt)) nf.add(new Option('Benutzerdefiniert', fmt));
    nf.value = fmt;
    $('#gridBtn').classList.toggle('on', S.grid !== false);
    $('#freezeBtn').classList.toggle('on', !!(S.freeze && (S.freeze.r || S.freeze.c)));
    $('#filterBtn').classList.toggle('on', !!S.filter);
    $('#showFormulasBtn').classList.toggle('on', showFormulas);
  }

  function setZoom(z) {
    const old = zoom;
    zoom = Math.min(2, Math.max(0.5, Math.round(z * 10) / 10));
    if (zoom === old) return;
    $('#zoomRange').value = Math.round(zoom * 100);
    $('#zoomLabel').textContent = Math.round(zoom * 100) + ' %';
    $$('.chart .svg').forEach(b => { b._svg = null; });
    render();
  }

  // =====================================================================
  //  Befehle / Menueband
  // =====================================================================
  const commands = {
    paste: async () => {
      try { pasteText(await navigator.clipboard.readText()); }
      catch { if (clip) pasteText(clip.text); else Rex.toast('Bitte Strg+V zum Einfügen benutzen'); }
      focusGrid();
    },
    cut: () => { doCopy(true); focusGrid(); },
    copy: () => { doCopy(false); focusGrid(); },
    formatPainter: () => {
      const { r1, c1, r2, c2 } = clipRange();
      painter = range(r1, r2).map(r => range(c1, c2).map(c => getCell(r, c) && getCell(r, c).s ? { ...getCell(r, c).s } : null));
      $('#painterBtn').classList.add('on');
      Rex.toast('Klicke auf die Zielzelle');
      focusGrid();
    },
    wrap: () => { const on = !styleAt(sel.r, sel.c).wrap; setStyle(s => { s.wrap = on; }); },
    merge: () => toggleMerge(),
    decInc: () => changeDecimals(1),
    decDec: () => changeDecimals(-1),
    autosum: () => autoSum('SUMME'),
    insertDate: () => { setInput(sel.r, sel.c, new Date().toLocaleDateString('de-DE')); commit(); updateFormulaBar(); },
    insertFunction: () => functionDialog(),
    showFormulas: () => { showFormulas = !showFormulas; render(); updateToolbar(); },
    recalc: () => { invalidate(); render(); Rex.toast('Neu berechnet'); },
    sortAsc: () => sortSel(true),
    sortDesc: () => sortSel(false),
    filter: () => toggleFilter(),
    dedupe: () => dedupe(),
    textToCols: () => textToCols(),
    gridlines: () => { S.grid = S.grid === false; commit(); updateToolbar(); },
    freeze: () => {
      if (S.freeze && (S.freeze.r || S.freeze.c)) { S.freeze = { r: 0, c: 0 }; commit(); updateToolbar(); return; }
      Rex.menu($('#freezeBtn'), [
        { label: 'Oberste Zeile fixieren', icon: 'panel-top', action: () => { S.freeze = { r: 1, c: 0 }; commit(); updateToolbar(); } },
        { label: 'Erste Spalte fixieren', icon: 'panel-left', action: () => { S.freeze = { r: 0, c: 1 }; commit(); updateToolbar(); } },
        { label: `Fenster fixieren (oberhalb/links von ${RF.addr(sel.r, sel.c)})`, icon: 'grid-2x2', action: () => { S.freeze = { r: sel.r, c: sel.c }; scroller.scrollTop = 0; scroller.scrollLeft = 0; commit(); updateToolbar(); } }
      ]);
    },
    zoomIn: () => setZoom(zoom + 0.1),
    zoomOut: () => setZoom(zoom - 0.1),
    zoom100: () => setZoom(1),
    addSheet: () => { wb.sheets.push(newSheet(uniqueSheetName('Tabelle' + (wb.sheets.length + 1)))); commit(); switchSheet(wb.sheets.length - 1); }
  };
  document.addEventListener('mousedown', e => {
    const b = e.target.closest('.ribbon .btn, .formula-bar .btn, .statusbar .btn, .sheet-bar .btn');
    if (b) e.preventDefault();
  });
  document.addEventListener('click', e => {
    const b = e.target.closest('[data-cmd]');
    if (b && commands[b.dataset.cmd]) { if (editing && b.dataset.cmd !== 'insertFunction' && !commitEdit()) return; commands[b.dataset.cmd](); return; }
    const st = e.target.closest('[data-style]');
    if (st) { if (editing && !commitEdit()) return; toggleStyle(st.dataset.style); focusGrid(); return; }
    const al = e.target.closest('[data-align]');
    if (al) { if (editing && !commitEdit()) return; const v = al.dataset.align; const on = styleAt(sel.r, sel.c).al !== v; setStyle(s => { s.al = on ? v : null; }); focusGrid(); return; }
    const va = e.target.closest('[data-valign]');
    if (va) { if (editing && !commitEdit()) return; const v = va.dataset.valign; setStyle(s => { s.va = v === 'bottom' ? null : v; }); focusGrid(); return; }
    const nf = e.target.closest('[data-fmt]');
    if (nf) { if (editing && !commitEdit()) return; setStyle(s => { s.fmt = nf.dataset.fmt; }); focusGrid(); return; }
    const chb = e.target.closest('[data-chart]');
    if (chb) { if (editing && !commitEdit()) return; insertChart(chb.dataset.chart); }
  });

  const ffSel = $('#fontFamily');
  FONTS.forEach(f => ffSel.add(new Option(f, f)));
  ffSel.addEventListener('change', () => { const v = ffSel.value; setStyle(s => { s.font = v === 'Calibri' ? null : v; }); focusGrid(); });
  const fsSel = $('#fontSize');
  SIZES.forEach(s => fsSel.add(new Option(s, s)));
  fsSel.addEventListener('change', () => {
    const v = +fsSel.value;
    setStyle(s => { s.size = v === 11 ? null : v; });
    // Zeilenhoehe automatisch anpassen
    const need = Math.round(v * 4 / 3 * 1.35 + 4);
    for (let r = sel.r1; r <= Math.min(sel.r2, sel.r1 + 1000); r++) if ((S.rowH[r] ?? DEF_H) < need) S.rowH[r] = need;
    commit();
    focusGrid();
  });
  const nfSel = $('#numFmt');
  NUMFMTS.forEach(f => nfSel.add(new Option(f.l, f.v)));
  nfSel.addEventListener('change', () => { const v = nfSel.value; setStyle(s => { s.fmt = v || null; }); focusGrid(); });

  $('#fillBtn').addEventListener('click', e => Rex.colorMenu(e.currentTarget, c => { if (c) $('#fillBar').style.background = c; setStyle(s => { s.bg = c; }); focusGrid(); }, { noneLabel: 'Keine Füllung' }));
  $('#colorBtn').addEventListener('click', e => Rex.colorMenu(e.currentTarget, c => { if (c) $('#colorBar').style.background = c; setStyle(s => { s.color = c; }); focusGrid(); }, { noneLabel: 'Automatisch' }));
  $('#borderBtn').addEventListener('click', e => Rex.menu(e.currentTarget, [
    { label: 'Alle Rahmenlinien', icon: 'grid-2x2', action: () => setBorders('all') },
    { label: 'Rahmenlinien außen', icon: 'square', action: () => setBorders('outer') },
    { label: 'Dicke Rahmenlinie außen', icon: 'square', action: () => setBorders('thickOuter') },
    { label: 'Rahmenlinie unten', icon: 'panel-bottom', action: () => setBorders('bottom') },
    { label: 'Rahmenlinie oben', icon: 'panel-top', action: () => setBorders('top') },
    { label: 'Rahmenlinie links', icon: 'panel-left', action: () => setBorders('left') },
    { label: 'Rahmenlinie rechts', icon: 'panel-right', action: () => setBorders('right') },
    { label: 'Doppelte Rahmenlinie unten', icon: 'panel-bottom', action: () => setBorders('doubleBottom') },
    '-',
    { label: 'Kein Rahmen', icon: 'square-dashed', action: () => setBorders('none') }
  ]));
  $('#insertCellsBtn').addEventListener('click', e => Rex.menu(e.currentTarget, [
    { label: 'Zeilen einfügen', icon: 'between-horizontal-end', action: () => insertRowsCols('row') },
    { label: 'Spalten einfügen', icon: 'between-vertical-end', action: () => insertRowsCols('col') },
    { label: 'Tabellenblatt einfügen', icon: 'sheet', action: () => commands.addSheet() }
  ]));
  $('#deleteCellsBtn').addEventListener('click', e => Rex.menu(e.currentTarget, [
    { label: 'Zeilen löschen', icon: 'rows-3', action: () => deleteRowsCols('row') },
    { label: 'Spalten löschen', icon: 'columns-3', action: () => deleteRowsCols('col') }
  ]));
  $('#formatCellsBtn').addEventListener('click', e => Rex.menu(e.currentTarget, [
    { label: 'Zeilenhöhe…', icon: 'move-vertical', action: async () => {
      const v = await Rex.dialog({ title: 'Zeilenhöhe', fields: [{ name: 'h', label: 'Höhe in Punkt', type: 'number', value: Math.round((S.rowH[sel.r] ?? DEF_H) * 0.75 * 10) / 10 }] });
      if (v && v.h > 0) { for (let r = sel.r1; r <= Math.min(sel.r2, sel.r1 + 5000); r++) S.rowH[r] = Math.round(v.h / 0.75); commit(); }
    } },
    { label: 'Spaltenbreite…', icon: 'move-horizontal', action: async () => {
      const v = await Rex.dialog({ title: 'Spaltenbreite', fields: [{ name: 'w', label: 'Breite in Zeichen', type: 'number', step: 0.5, value: RexXlsx.pxToChars(S.colW[sel.c] ?? DEF_W) }] });
      if (v && v.w > 0) { for (let c = sel.c1; c <= Math.min(sel.c2, sel.c1 + 500); c++) S.colW[c] = RexXlsx.charsToPx(v.w); commit(); }
    } },
    { label: 'Spaltenbreite automatisch anpassen', icon: 'unfold-horizontal', action: () => { for (let c = sel.c1; c <= Math.min(sel.c2, sel.c1 + 100); c++) autoFitCol(c); } },
    '-',
    { label: 'Eigenes Zahlenformat…', icon: 'hash', action: async () => {
      const v = await Rex.dialog({ title: 'Zahlenformat', html: '<p style="color:var(--text-2);margin:0 0 10px">Beispiele: <code>0.00</code>, <code>#,##0 "Stück"</code>, <code>DD.MM.YYYY</code>, <code>0%</code></p>',
        fields: [{ name: 'f', label: 'Formatcode (Excel)', value: styleAt(sel.r, sel.c).fmt || 'General' }] });
      if (v) setStyle(s => { s.fmt = v.f === 'General' ? null : v.f; });
    } }
  ]));
  $('#autosumDrop').addEventListener('click', e => Rex.menu(e.currentTarget, [
    { label: 'Summe', icon: 'sigma', action: () => autoSum('SUMME') },
    { label: 'Mittelwert', icon: 'divide', action: () => autoSum('MITTELWERT') },
    { label: 'Anzahl', icon: 'hash', action: () => autoSum('ANZAHL') },
    { label: 'Maximum', icon: 'arrow-up-to-line', action: () => autoSum('MAX') },
    { label: 'Minimum', icon: 'arrow-down-to-line', action: () => autoSum('MIN') }
  ]));
  $('#sortBtn').addEventListener('click', e => Rex.menu(e.currentTarget, [
    { label: 'Von A bis Z sortieren', icon: 'arrow-up-a-z', action: () => sortSel(true) },
    { label: 'Von Z bis A sortieren', icon: 'arrow-down-z-a', action: () => sortSel(false) },
    '-',
    { label: S.filter ? 'Filter entfernen' : 'Filtern', icon: 'filter', action: () => toggleFilter() }
  ]));
  $('#clearBtn').addEventListener('click', e => Rex.menu(e.currentTarget, [
    { label: 'Alles löschen', icon: 'eraser', action: () => clearRange('all') },
    { label: 'Formate löschen', icon: 'remove-formatting', action: () => clearRange('format') },
    { label: 'Inhalte löschen', icon: 'delete', action: () => clearRange('content') }
  ]));
  $('#zoomRange').addEventListener('input', e => setZoom(+e.target.value / 100));

  // Funktionsbibliothek
  const FN_INFO = {
    SUM: ['Mathematik', 'Addiert alle Zahlen', 'SUMME(A1:A10)'], AVERAGE: ['Statistik', 'Mittelwert der Zahlen', 'MITTELWERT(A1:A10)'],
    MIN: ['Statistik', 'Kleinster Wert', 'MIN(A1:A10)'], MAX: ['Statistik', 'Größter Wert', 'MAX(A1:A10)'], COUNT: ['Statistik', 'Zählt Zellen mit Zahlen', 'ANZAHL(A1:A10)'],
    COUNTA: ['Statistik', 'Zählt nicht-leere Zellen', 'ANZAHL2(A1:A10)'], IF: ['Logik', 'Wenn Bedingung, dann … sonst …', 'WENN(A1>10;"groß";"klein")'],
    IFERROR: ['Logik', 'Ersatzwert bei Fehler', 'WENNFEHLER(A1/B1;0)'], AND: ['Logik', 'Alle Bedingungen wahr?', 'UND(A1>0;B1>0)'], OR: ['Logik', 'Eine Bedingung wahr?', 'ODER(A1>0;B1>0)'],
    ROUND: ['Mathematik', 'Rundet auf Stellen', 'RUNDEN(A1;2)'], SUMIF: ['Mathematik', 'Summe mit Bedingung', 'SUMMEWENN(A1:A10;">5")'],
    COUNTIF: ['Statistik', 'Zählt mit Bedingung', 'ZÄHLENWENN(A1:A10;"Ja")'], AVERAGEIF: ['Statistik', 'Mittelwert mit Bedingung', 'MITTELWERTWENN(A1:A10;">0")'],
    VLOOKUP: ['Nachschlagen', 'Sucht senkrecht in einer Tabelle', 'SVERWEIS("Apfel";A1:C10;2;FALSCH)'], XLOOKUP: ['Nachschlagen', 'Moderne Suche', 'XVERWEIS("Apfel";A1:A10;B1:B10)'],
    INDEX: ['Nachschlagen', 'Wert an Position', 'INDEX(A1:C10;2;3)'], MATCH: ['Nachschlagen', 'Position eines Wertes', 'VERGLEICH("Apfel";A1:A10;0)'],
    CONCATENATE: ['Text', 'Verbindet Texte', 'VERKETTEN(A1;" ";B1)'], LEFT: ['Text', 'Zeichen von links', 'LINKS(A1;3)'], RIGHT: ['Text', 'Zeichen von rechts', 'RECHTS(A1;3)'],
    LEN: ['Text', 'Länge eines Textes', 'LÄNGE(A1)'], UPPER: ['Text', 'In Großbuchstaben', 'GROSS(A1)'], TODAY: ['Datum', 'Heutiges Datum', 'HEUTE()'],
    NOW: ['Datum', 'Jetzt (Datum + Uhrzeit)', 'JETZT()'], DATE: ['Datum', 'Datum aus Jahr/Monat/Tag', 'DATUM(2025;12;24)'], YEAR: ['Datum', 'Jahr eines Datums', 'JAHR(A1)'],
    MEDIAN: ['Statistik', 'Median', 'MEDIAN(A1:A10)'], RAND: ['Mathematik', 'Zufallszahl 0–1', 'ZUFALLSZAHL()'], RANDBETWEEN: ['Mathematik', 'Zufallszahl im Bereich', 'ZUFALLSBEREICH(1;6)']
  };
  async function functionDialog(cat) {
    const list = RF.FUNCTION_LIST.filter(f => !cat || (FN_INFO[f.en] && FN_INFO[f.en][0] === cat))
      .sort((a, b) => (FN_INFO[b.en] ? 1 : 0) - (FN_INFO[a.en] ? 1 : 0) || a.de.localeCompare(b.de));
    const v = await Rex.dialog({ title: 'Funktion einfügen', html: '<p style="color:var(--text-2);margin:0 0 10px">Deutsche und englische Namen funktionieren beide (z. B. SUMME = SUM).</p>', fields: [
      { name: 'f', label: 'Funktion', type: 'select', value: 'SUM', options: list.map(f => ({ value: f.en, label: `${f.de}${f.de !== f.en ? ' / ' + f.en : ''}${FN_INFO[f.en] ? ' – ' + FN_INFO[f.en][1] : ''}` })) }] });
    if (!v) return;
    const f = RF.FUNCTION_LIST.find(x => x.en === v.f);
    const name = f ? f.de : v.f;
    if (!editing) beginEdit('enter', '');
    const el = editing.fromBar ? fx : input;
    const p = el.selectionStart;
    let txt = el.value;
    if (!txt.startsWith('=')) { txt = '=' + txt; }
    const pos = el.value.startsWith('=') ? p : p + 1;
    el.value = txt.slice(0, pos) + name + '(' + txt.slice(pos);
    el.setSelectionRange(pos + name.length + 1, pos + name.length + 1);
    if (el === fx) input.value = fx.value; else fx.value = input.value;
    if (FN_INFO[v.f]) Rex.toast('Beispiel: =' + FN_INFO[v.f][2], 4000);
    positionInput();
    input.focus();
  }
  $('#fnCatBtn').addEventListener('click', e => Rex.menu(e.currentTarget,
    ['Mathematik', 'Statistik', 'Logik', 'Text', 'Datum', 'Nachschlagen'].map(c => ({ label: c, action: () => functionDialog(c) }))));

  // =====================================================================
  //  Laden / Speichern / Drucken
  // =====================================================================
  function setWorkbook(model) {
    wb = model;
    for (const sh of wb.sheets) {
      sh.cells = sh.cells || {}; sh.colW = sh.colW || {}; sh.rowH = sh.rowH || {}; sh.hidden = sh.hidden || {};
      sh.merges = sh.merges || []; sh.charts = sh.charts || []; sh.freeze = sh.freeze || { r: 0, c: 0 };
      if (sh.filter && !sh.filter.crit) sh.filter.crit = {};
    }
    wb.active = Math.min(wb.active || 0, wb.sheets.length - 1);
    S = wb.sheets[wb.active];
    invalidate();
    switchSheet(wb.active);
    resetHistory();
  }
  async function load(bytes, ext) {
    if (ext === 'xlsx') {
      setWorkbook(await RexXlsx.importXlsx(bytes));
    } else {
      const rows = RexXlsx.parseCSV(new TextDecoder().decode(bytes));
      const sh = newSheet('Tabelle1');
      const prev = S;
      S = sh;
      rows.forEach((row, r) => row.forEach((v, c) => { if (v !== '') setInput(r, c, v, sh); }));
      S = prev;
      setWorkbook({ sheets: [sh], active: 0 });
    }
  }
  async function save(fmt) {
    if (editing) commitEdit();
    if (fmt === 'xlsx') return RexXlsx.exportXlsx(wb, getValue);
    if (fmt === 'csv') {
      const { mr, mc } = usedExtent();
      const rows = [];
      for (let r = 0; r <= mr; r++) {
        const row = [];
        for (let c = 0; c <= mc; c++) { const d = display(wb.active, r, c); row.push(d ? d.text : ''); }
        rows.push(row);
      }
      return new TextEncoder().encode(RexXlsx.toCSV(rows));
    }
    throw new Error('Unbekanntes Format');
  }
  function buildPrint() {
    const area = $('#printArea');
    let html = '';
    wb.sheets.forEach((sh, si) => {
      let mr = -1, mc = -1;
      for (const k in sh.cells) { const [r, c] = k.split(',').map(Number); const cell = sh.cells[k]; if (cell.v != null || cell.f != null || (cell.s && (cell.s.bg || cell.s.bd))) { mr = Math.max(mr, r); mc = Math.max(mc, c); } }
      if (mr < 0 && !sh.charts.length) return;
      html += `<div class="sheet-print"><h3>${esc(sh.name)}</h3><table>`;
      const covered = new Set();
      for (let r = 0; r <= mr; r++) {
        if (sh.hidden[r]) continue;
        html += '<tr>';
        for (let c = 0; c <= mc; c++) {
          if (covered.has(key(r, c))) continue;
          const cell = sh.cells[key(r, c)];
          const m = sh.merges.find(x => x.r1 === r && x.c1 === c);
          if (m) for (let rr = m.r1; rr <= m.r2; rr++) for (let cc = m.c1; cc <= m.c2; cc++) covered.add(key(rr, cc));
          const prevS = S; S = sh;
          const d = cell ? display(si, r, c) : null;
          let css = cellStyleCss(cell && cell.s).replace(/font-size:[^;]+;/, '');
          S = prevS;
          const al = (cell && cell.s && cell.s.al) || (d && d.align) || 'left';
          css += `text-align:${al};min-width:${((sh.colW[c] ?? DEF_W) * 0.8)}px;`;
          html += `<td${m ? ` colspan="${m.c2 - m.c1 + 1}" rowspan="${m.r2 - m.r1 + 1}"` : ''} style="${css}">${d ? esc(d.text) : ''}</td>`;
        }
        html += '</tr>';
      }
      html += '</table>';
      const prevS = S;
      S = sh;
      const prevActive = wb.active;
      wb.active = si;
      for (const ch of sh.charts) html += `<div class="chart-print">${RexCharts.render(chartData(ch), ch.w, ch.h, {})}</div>`;
      wb.active = prevActive;
      S = prevS;
      html += '</div>';
    });
    area.innerHTML = html || '<p>Leere Tabelle</p>';
  }

  // Vorlagen
  const TEMPLATES = [
    { name: 'Leere Arbeitsmappe', icon: 'sheet', docName: 'Mappe', build: () => ({ sheets: [newSheet('Tabelle1')], active: 0 }) },
    { name: 'Haushaltsbuch', icon: 'wallet', docName: 'Haushaltsbuch', build: () => {
      const sh = newSheet('Budget');
      const prev = S; S = sh;
      const rows = [['Kategorie', 'Geplant', 'Tatsächlich', 'Differenz'], ['Miete', '650', '650'], ['Lebensmittel', '300', '342,50'], ['Handy', '20', '20'],
        ['Freizeit', '100', '135'], ['Sparen', '150', '120']];
      rows.forEach((row, r) => row.forEach((v, c) => setInput(r, c, v, sh)));
      for (let r = 1; r < rows.length; r++) setInput(r, 3, `=B${r + 1}-C${r + 1}`, sh);
      setInput(rows.length, 0, 'Summe', sh);
      for (const c of [1, 2, 3]) setInput(rows.length, c, `=SUMME(${RF.colName(c)}2:${RF.colName(c)}${rows.length})`, sh);
      for (let c = 0; c < 4; c++) { sh.cells[key(0, c)].s = { b: true, bg: '#217346', color: '#ffffff' }; }
      for (let r = 1; r <= rows.length; r++) for (let c = 1; c < 4; c++) { const k = key(r, c); sh.cells[k] = sh.cells[k] || {}; sh.cells[k].s = { ...(sh.cells[k].s || {}), fmt: '#,##0.00 "€"' }; }
      for (let c = 0; c < 4; c++) { const k = key(rows.length, c); sh.cells[k].s = { ...(sh.cells[k].s || {}), b: true, bd: { t: 'thin' } }; }
      sh.colW[0] = 130; sh.colW[1] = 100; sh.colW[2] = 100; sh.colW[3] = 100;
      sh.charts.push({ id: 'ch1', type: 'column', title: 'Geplant vs. Tatsächlich', series: [
        { name: '$B$1', cat: '$A$2:$A$6', val: '$B$2:$B$6' }, { name: '$C$1', cat: '$A$2:$A$6', val: '$C$2:$C$6' }], x: 470, y: 10, w: 460, h: 280 });
      sh.freeze = { r: 1, c: 0 };
      S = prev;
      return { sheets: [sh], active: 0 };
    } },
    { name: 'Notenrechner', icon: 'graduation-cap', docName: 'Noten', build: () => {
      const sh = newSheet('Noten');
      const rows = [['Fach', 'Note 1', 'Note 2', 'Note 3', 'Schnitt'], ['Mathe', '2', '3', '1'], ['Deutsch', '2', '2', '3'], ['Englisch', '1', '2', '2'], ['Physik', '3', '2', '2']];
      rows.forEach((row, r) => row.forEach((v, c) => setInput(r, c, v, sh)));
      for (let r = 1; r < rows.length; r++) setInput(r, 4, `=RUNDEN(MITTELWERT(B${r + 1}:D${r + 1});2)`, sh);
      setInput(rows.length + 1, 0, 'Gesamt', sh);
      setInput(rows.length + 1, 4, `=RUNDEN(MITTELWERT(E2:E${rows.length});2)`, sh);
      for (let c = 0; c < 5; c++) sh.cells[key(0, c)].s = { b: true, bg: '#d9e2f3', bd: { b: 'thin' } };
      sh.cells[key(rows.length + 1, 0)].s = { b: true };
      sh.cells[key(rows.length + 1, 4)].s = { b: true, bg: '#fff2cc' };
      sh.colW[0] = 110;
      return { sheets: [sh], active: 0 };
    } }
  ];

  // =====================================================================
  //  Start
  // =====================================================================
  Rex.init({
    app: 'calc',
    letter: 'X',
    icon: 'file-spreadsheet',
    defaultName: 'Mappe',
    openFormats: [{ name: 'Excel-Arbeitsmappe', extensions: ['xlsx'] }, { name: 'CSV-Datei', extensions: ['csv'] }],
    saveFormats: [{ name: 'Excel-Arbeitsmappe', extensions: ['xlsx'], hint: 'Kompatibel mit Microsoft Excel, LibreOffice und Google Tabellen' }, { name: 'CSV-Datei', extensions: ['csv'], hint: 'Nur aktuelle Tabelle, ohne Formatierung' }],
    templates: TEMPLATES.map(t => ({ ...t, create: () => setWorkbook(t.build()) })),
    load, save, undo, redo,
    shortcutsHelp: [['F2', 'Zelle bearbeiten'], ['Alt+Enter', 'Zeilenumbruch in Zelle'], ['F4', '$-Bezug umschalten'], ['Strg+Pfeil', 'Zum Datenrand springen'],
      ['Alt+=', 'AutoSumme'], ['Strg+D / Strg+R', 'Nach unten / rechts ausfüllen'], ['Strg+;', 'Heutiges Datum']],
    pdf: { pageSize: 'A4', landscape: true, margins: () => ({ top: 0.5, bottom: 0.5, left: 0.5, right: 0.5 }), before: buildPrint, after: () => { $('#printArea').innerHTML = ''; } },
    onTheme: () => { $$('.chart .svg').forEach(b => { b._svg = null; }); render(); }
  });

  input.classList.add('hidden');
  renderTabs();
  resetHistory();
  render();
  updateFormulaBar();
  updateToolbar();
  focusGrid();
  window.addEventListener('focus', () => setTimeout(() => { if (!editing && document.activeElement === document.body) focusGrid(); }, 0));
  document.addEventListener('mouseup', () => setTimeout(() => {
    if (!editing && (document.activeElement === document.body || document.activeElement === wrap)) focusGrid();
  }, 0));

  window.RexCalc = { get wb() { return wb; }, getValue, setInput: (r, c, t) => { setInput(r, c, t); commit(); }, select, save, load, insertChart, display: (r, c) => display(wb.active, r, c) };
})();
