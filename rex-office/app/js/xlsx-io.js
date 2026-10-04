/* REX Tabelle – Lesen und Schreiben von Excel-Dateien (.xlsx) und CSV */
(function () {
  'use strict';
  const DEF_W = 72, DEF_H = 21;
  const RF = () => window.RexFormula;
  const EMU = 9525; // EMU pro Pixel

  const pxToChars = (px) => Math.max(0, Math.round(((px - 5) / 7) * 100) / 100);
  const charsToPx = (w) => Math.round(w * 7 + 5);
  const THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];
  const INDEXED = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
    '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
    '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'];

  function colorIn(c) {
    if (!c) return null;
    let hex = null;
    if (c.argb) hex = c.argb.slice(-6);
    else if (c.theme != null) hex = THEME[c.theme] || null;
    else if (c.indexed != null && c.indexed < 64) hex = INDEXED[c.indexed];
    if (!hex) return null;
    if (c.tint) {
      const t = c.tint;
      hex = [0, 2, 4].map(i => {
        let v = parseInt(hex.slice(i, i + 2), 16);
        v = t > 0 ? v + (255 - v) * t : v * (1 + t);
        return Math.round(Math.max(0, Math.min(255, v))).toString(16).padStart(2, '0');
      }).join('');
    }
    return '#' + hex.toLowerCase();
  }
  const argb = (hex) => ({ argb: 'FF' + hex.replace('#', '').toUpperCase().padStart(6, '0') });

  // ---------- Geometrie (fuer Diagramm-Anker) ----------
  function colX(sh, c) { let x = 0; for (let i = 0; i < c; i++) x += sh.colW[i] ?? DEF_W; return x; }
  function rowY(sh, r) { let y = 0; for (let i = 0; i < r; i++) y += sh.hidden && sh.hidden[i] ? 0 : (sh.rowH[i] ?? DEF_H); return y; }
  function colAt(sh, x) { let c = 0, acc = 0; while (c < 16383) { const w = sh.colW[c] ?? DEF_W; if (acc + w > x) return { c, off: x - acc }; acc += w; c++; } return { c, off: 0 }; }
  function rowAt(sh, y) { let r = 0, acc = 0; while (r < 1048575) { const h = sh.hidden && sh.hidden[r] ? 0 : (sh.rowH[r] ?? DEF_H); if (acc + h > y) return { r, off: y - acc }; acc += h; r++; } return { r, off: 0 }; }

  // =====================================================================
  //  IMPORT
  // =====================================================================
  function styleIn(cs) {
    const s = {};
    if (!cs) return s;
    const f = cs.font;
    if (f) {
      if (f.bold) s.b = true;
      if (f.italic) s.i = true;
      if (f.underline && f.underline !== 'none') s.u = true;
      if (f.strike) s.s = true;
      if (f.name && f.name !== 'Calibri') s.font = f.name;
      if (f.size && f.size !== 11) s.size = f.size;
      const col = colorIn(f.color);
      if (col && col !== '#000000') s.color = col;
    }
    const fill = cs.fill;
    if (fill && fill.type === 'pattern' && fill.pattern && fill.pattern !== 'none') {
      const col = colorIn(fill.fgColor) || colorIn(fill.bgColor);
      if (col) s.bg = col;
    } else if (fill && fill.type === 'gradient' && fill.stops && fill.stops[0]) {
      const col = colorIn(fill.stops[0].color);
      if (col) s.bg = col;
    }
    const a = cs.alignment;
    if (a) {
      if (['left', 'center', 'right'].includes(a.horizontal)) s.al = a.horizontal;
      else if (a.horizontal === 'centerContinuous') s.al = 'center';
      if (a.vertical === 'top') s.va = 'top';
      if (a.vertical === 'middle') s.va = 'middle';
      if (a.wrapText) s.wrap = true;
    }
    if (cs.numFmt && cs.numFmt !== 'General') s.fmt = cs.numFmt;
    const b = cs.border;
    if (b) {
      const bd = {};
      for (const [k, kk] of [['top', 't'], ['right', 'r'], ['bottom', 'b'], ['left', 'l']]) {
        if (b[k] && b[k].style) {
          bd[kk] = /thick/.test(b[k].style) ? 'thick' : /medium/.test(b[k].style) ? 'medium' : 'thin';
          const col = colorIn(b[k].color);
          if (col && col !== '#000000') bd[kk + 'c'] = col;
        }
      }
      if (Object.keys(bd).length) s.bd = bd;
    }
    return s;
  }

  async function importXlsx(bytes) {
    const ExcelJS = window.ExcelJS;
    const wbx = new ExcelJS.Workbook();
    const ab = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
    await wbx.xlsx.load(ab);
    const sheets = [];
    const worksheets = wbx.worksheets.slice().sort((a, b) => (a.orderNo ?? 0) - (b.orderNo ?? 0));
    for (const ws of worksheets) {
      const sh = { name: ws.name, cells: {}, colW: {}, rowH: {}, hidden: {}, merges: [], charts: [], grid: true, freeze: { r: 0, c: 0 }, filter: null };
      (ws.columns || []).forEach((col, i) => {
        if (col && col.hidden) sh.colW[i] = 0;
        else if (col && col.width && Math.abs(charsToPx(col.width) - DEF_W) > 1) sh.colW[i] = charsToPx(col.width);
      });
      ws.eachRow({ includeEmpty: true }, (row, rn) => {
        const r = rn - 1;
        if (row.hidden) sh.hidden[r] = true;
        if (row.height && Math.abs(row.height / 0.75 - DEF_H) > 1) sh.rowH[r] = Math.round(row.height / 0.75);
        row.eachCell({ includeEmpty: true }, (cell, cn) => {
          const c = cn - 1;
          const T = ExcelJS.ValueType;
          const out = {};
          const st = styleIn(cell.style);
          if (cell.type === T.Merge) { if (Object.keys(st).length) sh.cells[r + ',' + c] = { s: st }; return; }
          const v = cell.value;
          switch (cell.type) {
            case T.Null: break;
            case T.Number: out.v = v; break;
            case T.String: case T.SharedString: out.v = String(v); break;
            case T.Boolean: out.v = !!v; break;
            case T.Date: out.v = RF().dateToSerial(v); if (!st.fmt || st.fmt === 'General') st.fmt = 'DD.MM.YYYY'; break;
            case T.RichText: out.v = (v.richText || []).map(t => t.text).join(''); break;
            case T.Hyperlink: out.v = typeof v.text === 'string' ? v.text : (v.text && v.text.richText ? v.text.richText.map(t => t.text).join('') : String(v.hyperlink || '')); out.link = v.hyperlink; break;
            case T.Error: out.v = RF().EXCEL_ERR[v.error] ? RF().EXCEL_ERR[v.error].code : v.error; break;
            case T.Formula: {
              let f = cell.formula;
              if (!f && v && v.sharedFormula) f = null;
              if (f) out.f = RF().toGerman(f.replace(/^=/, '').replace(/_xlfn\./gi, ''));
              let res = v && v.result;
              if (res instanceof Date) { res = RF().dateToSerial(res); if (!st.fmt) st.fmt = 'DD.MM.YYYY'; }
              if (res && typeof res === 'object' && res.error) res = null;
              if (res != null) out.c = res;
              if (!out.f) out.v = res;
              break;
            }
            default: if (v != null) out.v = typeof v === 'object' ? (cell.text || '') : v;
          }
          if (Object.keys(st).length) out.s = st;
          if (out.v !== undefined || out.f !== undefined || out.s) sh.cells[r + ',' + c] = out;
        });
      });
      // Verbundene Zellen
      const merges = (ws.model && ws.model.merges) || [];
      for (const m of merges) {
        const rg = RF().refRange(m);
        sh.merges.push({ r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2 });
      }
      // Ansicht
      const view = (ws.views || [])[0];
      if (view) {
        if (view.state === 'frozen') sh.freeze = { r: view.ySplit || 0, c: view.xSplit || 0 };
        if (view.showGridLines === false) sh.grid = false;
      }
      if (ws.autoFilter) {
        const af = typeof ws.autoFilter === 'string' ? ws.autoFilter : (ws.autoFilter.from && ws.autoFilter.to ?
          (typeof ws.autoFilter.from === 'string' ? ws.autoFilter.from + ':' + ws.autoFilter.to : RF().addr(ws.autoFilter.from.row - 1, ws.autoFilter.from.column - 1) + ':' + RF().addr(ws.autoFilter.to.row - 1, ws.autoFilter.to.column - 1)) : null);
        if (af) { const rg = RF().refRange(af); sh.filter = { r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2, crit: {} }; }
      }
      sheets.push(sh);
    }
    try { await importCharts(bytes, sheets); } catch (e) { console.warn('Diagramme konnten nicht gelesen werden', e); }
    if (!sheets.length) sheets.push({ name: 'Tabelle1', cells: {}, colW: {}, rowH: {}, hidden: {}, merges: [], charts: [], grid: true, freeze: { r: 0, c: 0 }, filter: null });
    return { sheets, active: Math.max(0, Math.min(sheets.length - 1, (wbx.views && wbx.views[0] && wbx.views[0].activeTab) || 0)) };
  }

  // Diagramme aus der Datei lesen (ExcelJS kennt keine Diagramme)
  async function importCharts(bytes, sheets) {
    const zip = await JSZip.loadAsync(bytes);
    const parse = async (p) => { const f = zip.file(p); return f ? new DOMParser().parseFromString(await f.async('string'), 'application/xml') : null; };
    const relsOf = async (p) => {
      const dir = p.slice(0, p.lastIndexOf('/'));
      const name = p.slice(p.lastIndexOf('/') + 1);
      const x = await parse(`${dir}/_rels/${name}.rels`);
      const map = {};
      if (x) for (const r of x.getElementsByTagName('Relationship')) map[r.getAttribute('Id')] = resolvePath(dir, r.getAttribute('Target'));
      return map;
    };
    const wbXml = await parse('xl/workbook.xml');
    if (!wbXml) return;
    const wbRels = await relsOf('xl/workbook.xml');
    for (const sEl of wbXml.getElementsByTagNameNS('*', 'sheet')) {
      const name = sEl.getAttribute('name');
      const rid = sEl.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || sEl.getAttribute('r:id');
      const sheetPath = wbRels[rid];
      const sh = sheets.find(s => s.name === name);
      if (!sh || !sheetPath) continue;
      const sx = await parse(sheetPath);
      if (!sx) continue;
      const dEl = sx.getElementsByTagNameNS('*', 'drawing')[0];
      if (!dEl) continue;
      const sRels = await relsOf(sheetPath);
      const drawingPath = sRels[dEl.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || dEl.getAttribute('r:id')];
      const dx = drawingPath && await parse(drawingPath);
      if (!dx) continue;
      const dRels = await relsOf(drawingPath);
      const anchors = [...dx.getElementsByTagNameNS('*', 'twoCellAnchor'), ...dx.getElementsByTagNameNS('*', 'oneCellAnchor')];
      for (const an of anchors) {
        const chartRef = an.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/chart', 'chart')[0];
        if (!chartRef) continue;
        const cPath = dRels[chartRef.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || chartRef.getAttribute('r:id')];
        const cx = cPath && await parse(cPath);
        if (!cx) continue;
        const spec = parseChartXml(cx);
        if (!spec) continue;
        const pos = (el) => el ? {
          c: +el.getElementsByTagNameNS('*', 'col')[0].textContent, r: +el.getElementsByTagNameNS('*', 'row')[0].textContent,
          dx: +(el.getElementsByTagNameNS('*', 'colOff')[0]?.textContent || 0) / EMU, dy: +(el.getElementsByTagNameNS('*', 'rowOff')[0]?.textContent || 0) / EMU
        } : null;
        const from = pos(an.getElementsByTagNameNS('*', 'from')[0]);
        const to = pos(an.getElementsByTagNameNS('*', 'to')[0]);
        const x = colX(sh, from.c) + from.dx, y = rowY(sh, from.r) + from.dy;
        let w = 480, h = 290;
        if (to) { w = colX(sh, to.c) + to.dx - x; h = rowY(sh, to.r) + to.dy - y; }
        else {
          const ext = an.getElementsByTagNameNS('*', 'ext')[0];
          if (ext) { w = +ext.getAttribute('cx') / EMU; h = +ext.getAttribute('cy') / EMU; }
        }
        sh.charts.push({ ...spec, id: 'ch' + Math.random().toString(36).slice(2, 8), x, y, w: Math.max(120, w), h: Math.max(90, h) });
      }
    }
  }
  function resolvePath(dir, target) {
    if (target.startsWith('/')) return target.slice(1);
    const parts = (dir + '/' + target).split('/');
    const out = [];
    for (const p of parts) { if (p === '..') out.pop(); else if (p !== '.') out.push(p); }
    return out.join('/');
  }
  function parseChartXml(x) {
    const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
    const kinds = ['barChart', 'bar3DChart', 'lineChart', 'line3DChart', 'areaChart', 'area3DChart', 'pieChart', 'pie3DChart', 'doughnutChart', 'scatterChart', 'radarChart'];
    let plot = null, kind = null;
    for (const k of kinds) { plot = x.getElementsByTagNameNS(C, k)[0]; if (plot) { kind = k; break; } }
    if (!plot) return null;
    let type = 'column';
    if (/bar/.test(kind)) { const dir = plot.getElementsByTagNameNS(C, 'barDir')[0]; type = dir && dir.getAttribute('val') === 'bar' ? 'bar' : 'column'; }
    else if (/line|area|scatter|radar/.test(kind)) type = 'line';
    else if (/pie|doughnut/.test(kind)) type = 'pie';
    const titleEl = x.getElementsByTagNameNS(C, 'title')[0];
    let title = '';
    if (titleEl) title = Array.from(titleEl.getElementsByTagNameNS('*', 't')).map(t => t.textContent).join('');
    const series = [];
    for (const ser of plot.getElementsByTagNameNS(C, 'ser')) {
      const f = (tag) => {
        const el = ser.getElementsByTagNameNS(C, tag)[0];
        if (!el) return null;
        const fe = el.getElementsByTagNameNS(C, 'f')[0];
        return fe ? fe.textContent : null;
      };
      const txEl = ser.getElementsByTagNameNS(C, 'tx')[0];
      let name = f('tx');
      if (!name && txEl) { const v = txEl.getElementsByTagNameNS(C, 'v')[0]; if (v) name = '"' + v.textContent + '"'; }
      series.push({ name, cat: f('cat') || f('xVal'), val: f('val') || f('yVal') });
    }
    if (!series.length) return null;
    return { type, title, series };
  }

  // =====================================================================
  //  EXPORT
  // =====================================================================
  function styleOut(xc, s) {
    if (!s) return;
    if (s.b || s.i || s.u || s.s || s.font || s.size || s.color) {
      xc.font = { name: s.font || 'Calibri', size: s.size || 11, bold: !!s.b, italic: !!s.i, underline: !!s.u, strike: !!s.s, ...(s.color ? { color: argb(s.color) } : {}) };
    }
    if (s.bg) xc.fill = { type: 'pattern', pattern: 'solid', fgColor: argb(s.bg), bgColor: argb(s.bg) };
    if (s.al || s.va || s.wrap) xc.alignment = { ...(s.al ? { horizontal: s.al } : {}), ...(s.va ? { vertical: s.va } : {}), ...(s.wrap ? { wrapText: true } : {}) };
    if (s.fmt) xc.numFmt = s.fmt;
    if (s.bd) {
      const b = {};
      for (const [k, kk] of [['top', 't'], ['right', 'r'], ['bottom', 'b'], ['left', 'l']]) {
        if (s.bd[kk]) b[k] = { style: s.bd[kk], color: argb(s.bd[kk + 'c'] || '#000000') };
      }
      xc.border = b;
    }
  }

  async function exportXlsx(model, valueOf) {
    const ExcelJS = window.ExcelJS;
    const R = RF();
    const wbx = new ExcelJS.Workbook();
    wbx.creator = 'REX Office';
    wbx.created = new Date();
    wbx.views = [{ activeTab: model.active || 0, firstSheet: 0, visibility: 'visible' }];
    const used = new Set();
    const chartJobs = [];
    model.sheets.forEach((sh, si) => {
      let name = sh.name.replace(/[\\/?*[\]:]/g, '_').slice(0, 31) || 'Tabelle' + (si + 1);
      while (used.has(name.toLowerCase())) name = name.slice(0, 28) + '_' + si;
      used.add(name.toLowerCase());
      const ws = wbx.addWorksheet(name, { properties: { defaultColWidth: pxToChars(DEF_W), defaultRowHeight: 15.75 } });
      const view = { showGridLines: sh.grid !== false };
      if (sh.freeze && (sh.freeze.r || sh.freeze.c)) Object.assign(view, { state: 'frozen', xSplit: sh.freeze.c || 0, ySplit: sh.freeze.r || 0 });
      ws.views = [view];
      for (const [c, w] of Object.entries(sh.colW)) {
        const col = ws.getColumn(+c + 1);
        if (w === 0) col.hidden = true; else col.width = pxToChars(w);
      }
      for (const [r, h] of Object.entries(sh.rowH)) ws.getRow(+r + 1).height = Math.round(h * 0.75 * 100) / 100;
      for (const r of Object.keys(sh.hidden || {})) if (sh.hidden[r]) ws.getRow(+r + 1).hidden = true;
      for (const [k, cell] of Object.entries(sh.cells)) {
        const [r, c] = k.split(',').map(Number);
        const xc = ws.getCell(r + 1, c + 1);
        if (cell.f != null && cell.f !== '') {
          let res = valueOf(si, r, c);
          if (R.isErr(res)) res = { error: R.ERR_EXCEL[res.code] || '#VALUE!' };
          xc.value = { formula: R.toExcel(cell.f), result: res == null ? undefined : res };
        } else if (cell.v != null && cell.v !== '') {
          if (cell.link) xc.value = { text: String(cell.v), hyperlink: cell.link };
          else xc.value = cell.v;
        }
        styleOut(xc, cell.s);
      }
      for (const m of sh.merges || []) {
        try { ws.mergeCells(m.r1 + 1, m.c1 + 1, m.r2 + 1, m.c2 + 1); } catch (e) { console.warn('Verbinden fehlgeschlagen', e); }
      }
      if (sh.filter) ws.autoFilter = R.addr(sh.filter.r1, sh.filter.c1) + ':' + R.addr(sh.filter.r2, sh.filter.c2);
      if (sh.charts && sh.charts.length) chartJobs.push({ index: si + 1, sheet: sh, name, charts: sh.charts, si });
    });
    let buf = new Uint8Array(await wbx.xlsx.writeBuffer());
    if (chartJobs.length) buf = await injectCharts(buf, chartJobs, model, valueOf);
    return buf;
  }

  // ---------- Diagramme in die Datei schreiben ----------
  const xmlEsc = (s) => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  function qualify(ref, sheetName) {
    if (!ref) return null;
    if (ref.startsWith('"')) return ref;
    if (ref.includes('!')) return ref;
    const abs = ref.split(':').map(p => p.replace(/^\$?([A-Za-z]+)\$?(\d+)$/, '$$$1$$$2')).join(':');
    return `'${sheetName.replace(/'/g, "''")}'!${abs}`;
  }
  function cacheXml(values, numeric) {
    const pts = values.map((v, i) => `<c:pt idx="${i}"><c:v>${xmlEsc(numeric ? (typeof v === 'number' ? v : (+v || 0)) : (v ?? ''))}</c:v></c:pt>`).join('');
    return numeric ? `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${pts}</c:numCache>`
      : `<c:strCache><c:ptCount val="${values.length}"/>${pts}</c:strCache>`;
  }
  function chartXml(ch, sheetName, resolve) {
    const colors = window.RexCharts ? RexCharts.COLORS : ['#4472C4'];
    const ser = ch.series.map((s, i) => {
      const nameRef = qualify(s.name, sheetName);
      const catRef = qualify(s.cat, sheetName);
      const valRef = qualify(s.val, sheetName);
      const vals = resolve(s.val);
      const cats = s.cat ? resolve(s.cat) : null;
      const nameVal = s.name ? (s.name.startsWith('"') ? s.name.slice(1, -1) : String(resolve(s.name)[0] ?? '')) : `Reihe ${i + 1}`;
      const color = colors[i % colors.length].slice(1);
      const fill = ch.type === 'line'
        ? `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:round/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>`
        : ch.type === 'pie' ? '' : `<c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr><c:invertIfNegative val="0"/>`;
      const pieColors = ch.type === 'pie' ? vals.map((_, k) => `<c:dPt><c:idx val="${k}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${colors[k % colors.length].slice(1)}"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>`).join('') : '';
      const tx = nameRef && !nameRef.startsWith('"') ? `<c:tx><c:strRef><c:f>${xmlEsc(nameRef)}</c:f>${cacheXml([nameVal], false)}</c:strRef></c:tx>` : `<c:tx><c:v>${xmlEsc(nameVal)}</c:v></c:tx>`;
      const cat = catRef ? `<c:cat><c:strRef><c:f>${xmlEsc(catRef)}</c:f>${cacheXml(cats.map(v => v == null ? '' : String(v)), false)}</c:strRef></c:cat>` : '';
      return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${fill}${pieColors}${cat}<c:val><c:numRef><c:f>${xmlEsc(valRef)}</c:f>${cacheXml(vals, true)}</c:numRef></c:val>${ch.type === 'line' ? '<c:smooth val="0"/>' : ''}</c:ser>`;
    }).join('');
    const axes = `<c:axId val="111111"/><c:axId val="222222"/>`;
    let plot;
    if (ch.type === 'pie') plot = `<c:pieChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/></c:pieChart>`;
    else if (ch.type === 'line') plot = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${ser}<c:marker val="1"/>${axes}</c:lineChart>`;
    else plot = `<c:barChart><c:barDir val="${ch.type === 'bar' ? 'bar' : 'col'}"/><c:grouping val="clustered"/><c:varyColors val="0"/>${ser}<c:gapWidth val="219"/><c:overlap val="-27"/>${axes}</c:barChart>`;
    const axXml = ch.type === 'pie' ? '' : `
      <c:catAx><c:axId val="111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${ch.type === 'bar' ? 'l' : 'b'}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="222222"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>
      <c:valAx><c:axId val="222222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${ch.type === 'bar' ? 'b' : 'l'}"/><c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr><c:crossAx val="111111"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
    const title = ch.title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0"/></a:pPr><a:r><a:rPr lang="de-DE" sz="1400" b="0"/><a:t>${xmlEsc(ch.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>';
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:lang val="de-DE"/><c:roundedCorners val="0"/><c:chart>${title}<c:plotArea><c:layout/>${plot}${axXml}</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:chartSpace>`;
  }

  async function injectCharts(buf, jobs, model, valueOf) {
    const zip = await JSZip.loadAsync(buf);
    let ct = await zip.file('[Content_Types].xml').async('string');
    let chartNo = 0, drawingNo = 0;
    while (zip.file(`xl/drawings/drawing${drawingNo + 1}.xml`)) drawingNo++;
    while (zip.file(`xl/charts/chart${chartNo + 1}.xml`)) chartNo++;
    const R = RF();
    for (const job of jobs) {
      const sheetPath = `xl/worksheets/sheet${job.index}.xml`;
      let sx = await zip.file(sheetPath)?.async('string');
      if (!sx || /<drawing\s/.test(sx)) continue;
      drawingNo++;
      const resolve = (ref) => {
        if (!ref) return [];
        if (ref.startsWith('"')) return [ref.slice(1, -1)];
        const toks = R.refsIn(ref);
        if (!toks.length) return [];
        const t = toks[0];
        const si = t.sheet != null ? model.sheets.findIndex(s => s.name.toLowerCase() === t.sheet.toLowerCase()) : job.si;
        if (si < 0) return [];
        const rg = R.refRange(t.v);
        const out = [];
        for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 5000); r++) for (let c = rg.c1; c <= rg.c2; c++) {
          const v = valueOf(si, r, c);
          out.push(R.isErr(v) ? null : v);
        }
        return out;
      };
      let anchors = '', drawRels = '';
      job.charts.forEach((ch, i) => {
        chartNo++;
        const from = colAt(job.sheet, ch.x), fromR = rowAt(job.sheet, ch.y);
        const to = colAt(job.sheet, ch.x + ch.w), toR = rowAt(job.sheet, ch.y + ch.h);
        anchors += `<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>${from.c}</xdr:col><xdr:colOff>${Math.round(from.off * EMU)}</xdr:colOff><xdr:row>${fromR.r}</xdr:row><xdr:rowOff>${Math.round(fromR.off * EMU)}</xdr:rowOff></xdr:from><xdr:to><xdr:col>${to.c}</xdr:col><xdr:colOff>${Math.round(to.off * EMU)}</xdr:colOff><xdr:row>${toR.r}</xdr:row><xdr:rowOff>${Math.round(toR.off * EMU)}</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${i + 2}" name="Diagramm ${i + 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId${i + 1}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
        drawRels += `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${chartNo}.xml"/>`;
        zip.file(`xl/charts/chart${chartNo}.xml`, chartXml(ch, job.name, resolve));
        ct = ct.replace('</Types>', `<Override PartName="/xl/charts/chart${chartNo}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>`);
      });
      zip.file(`xl/drawings/drawing${drawingNo}.xml`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${anchors}</xdr:wsDr>`);
      zip.file(`xl/drawings/_rels/drawing${drawingNo}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${drawRels}</Relationships>`);
      ct = ct.replace('</Types>', `<Override PartName="/xl/drawings/drawing${drawingNo}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>`);

      // Beziehung Tabelle -> Zeichnung
      const relPath = `xl/worksheets/_rels/sheet${job.index}.xml.rels`;
      let rels = zip.file(relPath) ? await zip.file(relPath).async('string')
        : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
      let rid = 'rIdRexDrawing1';
      rels = rels.replace('</Relationships>', `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawingNo}.xml"/></Relationships>`);
      zip.file(relPath, rels);
      if (!/xmlns:r=/.test(sx.slice(0, 600))) sx = sx.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ');
      const tag = `<drawing r:id="${rid}"/>`;
      const m = /<(legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)[\s>/]/.exec(sx);
      sx = m ? sx.slice(0, m.index) + tag + sx.slice(m.index) : sx.replace('</worksheet>', tag + '</worksheet>');
      zip.file(sheetPath, sx);
    }
    zip.file('[Content_Types].xml', ct);
    return new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
  }

  // =====================================================================
  //  CSV
  // =====================================================================
  function parseCSV(text) {
    text = text.replace(/^﻿/, '');
    const first = text.split(/\r?\n/)[0] || '';
    const count = (ch) => first.split(ch).length - 1;
    const delim = [';', ',', '\t'].sort((a, b) => count(b) - count(a))[0];
    const rows = [];
    let row = [], field = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) {
        if (ch === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else q = false; }
        else field += ch;
      } else if (ch === '"') q = true;
      else if (ch === delim) { row.push(field); field = ''; }
      else if (ch === '\n' || ch === '\r') {
        if (ch === '\r' && text[i + 1] === '\n') i++;
        row.push(field); rows.push(row); row = []; field = '';
      } else field += ch;
    }
    if (field || row.length) { row.push(field); rows.push(row); }
    return rows;
  }
  function toCSV(rows) {
    return '﻿' + rows.map(r => r.map(v => {
      const s = String(v ?? '');
      return /[";\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    }).join(';')).join('\r\n');
  }

  window.RexXlsx = { importXlsx, exportXlsx, parseCSV, toCSV, DEF_W, DEF_H, pxToChars, charsToPx };
})();
