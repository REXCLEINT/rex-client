/* REX Praesentation – Lesen und Schreiben von PowerPoint-Dateien (.pptx) */
(function () {
  'use strict';
  const EMU = 9525;          // EMU pro Pixel (96 dpi)
  const PX_PER_IN = 96;
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
  const TRANSITIONS = { fade: '<p:fade/>', push: '<p:push dir="l"/>', wipe: '<p:wipe dir="r"/>', zoom: '<p:zoom/>', cover: '<p:cover dir="u"/>' };
  const SHAPES = ['rect', 'roundRect', 'ellipse', 'triangle', 'rightArrow', 'leftArrow', 'line', 'star5', 'hexagon', 'chevron', 'diamond', 'pentagon', 'heart', 'parallelogram', 'cloud'];

  const hex = (c) => (c || '#000000').replace('#', '').toUpperCase().slice(0, 6).padStart(6, '0');
  const inch = (px) => Math.round(px / PX_PER_IN * 1000) / 1000;

  // =====================================================================
  //  EXPORT
  // =====================================================================
  async function exportPptx(pres, helpers) {
    const pptx = new PptxGenJS();
    pptx.author = 'REX Office';
    pptx.company = 'REX';
    pptx.title = helpers.title || 'Präsentation';
    pptx.defineLayout({ name: 'REX', width: pres.w / PX_PER_IN, height: pres.h / PX_PER_IN });
    pptx.layout = 'REX';

    for (let si = 0; si < pres.slides.length; si++) {
      const sl = pres.slides[si];
      const slide = pptx.addSlide();
      if (sl.bgImage && pres.media[sl.bgImage]) slide.background = { data: pres.media[sl.bgImage].replace(/^data:/, '') };
      else slide.background = { color: hex(sl.bg || '#ffffff') };
      if (sl.notes) slide.addNotes(sl.notes);
      if (pres.showNumbers) slide.slideNumber = { x: inch(pres.w - 90), y: inch(pres.h - 34), w: 0.8, h: 0.3, fontSize: 11, color: '8C8C8C', align: 'right' };
      const paraMap = await helpers.paragraphs(si); // elId -> Absaetze (aus dem DOM berechnet)

      for (const el of sl.els) {
        const box = { x: inch(el.x), y: inch(el.y), w: inch(Math.max(1, el.w)), h: inch(Math.max(1, el.h)) };
        if (el.rot) box.rotate = el.rot;
        if (el.type === 'image') {
          const src = pres.media[el.media];
          if (!src) continue;
          const norm = await Rex.util.normalizeImage(src);
          slide.addImage({ data: norm.replace(/^data:/, ''), ...box, ...(el.flipH ? { flipH: true } : {}), ...(el.flipV ? { flipV: true } : {}) });
          continue;
        }
        if (el.type === 'table') {
          const rows = el.cells.map((row, ri) => row.map(cell => {
            const st = helpers.cellStyle(el, ri, cell);
            return { text: helpers.plain(cell.html), options: { bold: st.bold, color: hex(st.color), fill: { color: hex(st.bg) }, fontSize: el.size || 18, fontFace: el.font || 'Calibri', valign: 'middle', align: st.align || 'left' } };
          }));
          const total = el.colW ? el.colW.reduce((a, b) => a + b, 0) : 0;
          slide.addTable(rows, { ...box, colW: el.colW ? el.colW.map(f => box.w * f / total) : undefined, border: { type: 'solid', pt: 1, color: 'FFFFFF' }, margin: [3.6, 7.2, 3.6, 7.2] });
          continue;
        }
        if (el.type === 'chart') {
          const ch = el.chart;
          const type = ch.type === 'line' ? pptx.ChartType.line : ch.type === 'pie' ? pptx.ChartType.pie : pptx.ChartType.bar;
          const data = (ch.type === 'pie' ? ch.series.slice(0, 1) : ch.series).map(s => ({ name: s.name, labels: ch.categories, values: s.values.map(v => +v || 0) }));
          slide.addChart(type, data, {
            ...box, barDir: ch.type === 'bar' ? 'bar' : 'col', showLegend: true, legendPos: 'b', showTitle: !!ch.title, title: ch.title || undefined,
            chartColors: (window.RexCharts ? RexCharts.COLORS : ['4472C4']).map(hex), titleFontSize: 16, dataLabelFontSize: 10,
            showPercent: ch.type === 'pie', lineSize: 2, valGridLine: { color: 'D9D9D9', size: 0.75 }, catGridLine: { style: 'none' }
          });
          continue;
        }
        // Text und Formen
        const paras = paraMap.get(el.id) || [];
        const runs = [];
        paras.forEach((p, pi) => {
          const rs = p.runs.length ? p.runs : [{ text: '' }];
          rs.forEach((r, ri) => {
            const o = {
              fontFace: r.font || el.font || 'Calibri', fontSize: r.size || el.size || 18, color: hex(r.color || el.color || '#000000'),
              bold: !!r.bold, italic: !!r.italic, align: p.align || el.align || 'left',
              breakLine: ri === rs.length - 1 && pi < paras.length - 1
            };
            if (r.underline) o.underline = { style: 'sng' };
            if (r.strike) o.strike = 'sngStrike';
            if (r.br) o.softBreakBefore = true;
            if (r.highlight) o.highlight = hex(r.highlight);
            if (p.bullet) { o.bullet = p.bullet === 'number' ? { type: 'number' } : true; o.indentLevel = p.level || 0; }
            else if (p.level) o.indentLevel = p.level;
            runs.push({ text: r.text, options: o });
          });
        });
        const opts = {
          ...box, valign: el.va || (el.type === 'shape' ? 'middle' : 'top'), margin: [7.2, 7.2, 3.6, 3.6], fit: 'none',
          fontFace: el.font || 'Calibri', fontSize: el.size || 18, color: hex(el.color || '#000000')
        };
        if (el.fill) opts.fill = { color: hex(el.fill) };
        if (el.line && el.line.width > 0) opts.line = { color: hex(el.line.color), width: el.line.width };
        if (el.shadow) opts.shadow = { type: 'outer', blur: 4, offset: 3, angle: 45, color: '000000', opacity: 0.35 };
        if (el.type === 'shape') {
          const shapeName = SHAPES.includes(el.shape) ? el.shape : 'rect';
          if (shapeName === 'line') {
            slide.addShape('line', { ...box, line: { color: hex((el.line && el.line.color) || el.fill || '#000000'), width: (el.line && el.line.width) || 2 }, ...(el.flipV ? { flipV: true } : {}), ...(el.flipH ? { flipH: true } : {}) });
            continue;
          }
          opts.shape = shapeName;
          if (shapeName === 'roundRect') opts.rectRadius = Math.min(el.w, el.h) * 0.1667 / PX_PER_IN;
          if (!opts.fill) opts.fill = { type: 'none' };
          if (!opts.line) opts.line = { type: 'none' };
          if (!runs.length || runs.every(r => !r.text)) { slide.addShape(shapeName, opts); continue; }
        }
        if (!runs.length) runs.push({ text: '', options: {} });
        slide.addText(runs, opts);
      }
    }
    let out = await pptx.write({ outputType: 'uint8array' });
    out = await postProcess(out, pres);
    return out;
  }

  // Folien-XML nachbearbeiten:
  //  - PptxGenJS schreibt bei mehreren Textlaeufen pro Absatz mehrere <a:pPr> (ungueltig) -> nur das erste behalten
  //  - Uebergaenge eintragen (kann PptxGenJS nicht)
  async function postProcess(bytes, pres) {
    const zip = await JSZip.loadAsync(bytes);
    const PPR = /<a:pPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:pPr>)/g;
    for (let i = 0; i < pres.slides.length; i++) {
      const p = `ppt/slides/slide${i + 1}.xml`;
      const f = zip.file(p);
      if (!f) continue;
      let x = await f.async('string');
      x = x.replace(/<a:p>([\s\S]*?)<\/a:p>/g, (m, body) => {
        let first = true;
        const fixed = body.replace(PPR, (pp, offset) => {
          if (first && offset === 0) { first = false; return pp; }
          first = false;
          return '';
        });
        return '<a:p>' + fixed + '</a:p>';
      });
      const t = pres.slides[i].transition;
      if (t && t !== 'none' && TRANSITIONS[t] && !/<p:transition/.test(x)) {
        const tr = `<p:transition spd="med">${TRANSITIONS[t]}</p:transition>`;
        if (x.includes('</p:clrMapOvr>')) x = x.replace('</p:clrMapOvr>', '</p:clrMapOvr>' + tr);
        else x = x.replace('</p:cSld>', '</p:cSld>' + tr);
      }
      zip.file(p, x);
    }
    return new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
  }

  // =====================================================================
  //  IMPORT
  // =====================================================================
  const kids = (el, name) => el ? Array.from(el.children).filter(n => n.localName === name) : [];
  const kid = (el, name) => kids(el, name)[0] || null;
  const desc = (el, name) => el ? el.getElementsByTagNameNS('*', name)[0] || null : null;
  const rid = (el, attr = 'embed') => el ? (el.getAttributeNS(NS_R, attr) || el.getAttribute('r:' + attr)) : null;
  const esc = (s) => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

  // Farbmodifikationen (lumMod, lumOff, tint, shade)
  function rgbToHsl(r, g, b) {
    r /= 255; g /= 255; b /= 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    let h = 0, s = 0;
    const l = (max + min) / 2;
    if (max !== min) {
      const d = max - min;
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
      h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
      h /= 6;
    }
    return [h, s, l];
  }
  function hslToRgb(h, s, l) {
    if (s === 0) return [l, l, l].map(v => Math.round(v * 255));
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
    const f = (t) => { if (t < 0) t += 1; if (t > 1) t -= 1; if (t < 1 / 6) return p + (q - p) * 6 * t; if (t < 1 / 2) return q; if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6; return p; };
    return [f(h + 1 / 3), f(h), f(h - 1 / 3)].map(v => Math.round(v * 255));
  }
  function applyMods(hexStr, clrEl) {
    let r = parseInt(hexStr.slice(0, 2), 16), g = parseInt(hexStr.slice(2, 4), 16), b = parseInt(hexStr.slice(4, 6), 16);
    for (const m of clrEl.children) {
      const v = +m.getAttribute('val') / 100000;
      if (m.localName === 'lumMod' || m.localName === 'lumOff') {
        const [h, s, l] = rgbToHsl(r, g, b);
        const nl = m.localName === 'lumMod' ? l * v : Math.min(1, l + v);
        [r, g, b] = hslToRgb(h, s, Math.max(0, Math.min(1, nl)));
      } else if (m.localName === 'tint') {
        [r, g, b] = [r, g, b].map(c => Math.round(c * v + 255 * (1 - v)));
      } else if (m.localName === 'shade') {
        [r, g, b] = [r, g, b].map(c => Math.round(c * v));
      }
    }
    return '#' + [r, g, b].map(c => Math.max(0, Math.min(255, c)).toString(16).padStart(2, '0')).join('');
  }

  async function importPptx(bytes) {
    const zip = await JSZip.loadAsync(bytes);
    const xmlCache = new Map();
    const parse = async (p) => {
      if (xmlCache.has(p)) return xmlCache.get(p);
      const f = zip.file(p);
      const x = f ? new DOMParser().parseFromString(await f.async('string'), 'application/xml') : null;
      xmlCache.set(p, x);
      return x;
    };
    const resolvePath = (dir, target) => {
      if (target.startsWith('/')) return target.slice(1);
      const out = [];
      for (const part of (dir + '/' + target).split('/')) { if (part === '..') out.pop(); else if (part !== '.' && part !== '') out.push(part); }
      return out.join('/');
    };
    const relsOf = async (p) => {
      const dir = p.slice(0, p.lastIndexOf('/'));
      const x = await parse(`${dir}/_rels/${p.slice(p.lastIndexOf('/') + 1)}.rels`);
      const map = {};
      if (x) for (const r of x.getElementsByTagName('Relationship')) {
        map[r.getAttribute('Id')] = { path: r.getAttribute('TargetMode') === 'External' ? null : resolvePath(dir, r.getAttribute('Target')), type: r.getAttribute('Type') || '' };
      }
      return map;
    };
    const byType = (rels, t) => Object.values(rels).find(r => r.type.endsWith('/' + t));

    const presXml = await parse('ppt/presentation.xml');
    if (!presXml) throw new Error('Keine gültige PowerPoint-Datei.');
    const presRels = await relsOf('ppt/presentation.xml');
    const sz = desc(presXml, 'sldSz');
    const W = sz ? Math.round(+sz.getAttribute('cx') / EMU) : 960;
    const H = sz ? Math.round(+sz.getAttribute('cy') / EMU) : 540;
    const defaultTextStyle = desc(presXml, 'defaultTextStyle');

    const media = {};
    let mediaN = 0;
    const mediaByPath = new Map();
    async function loadMedia(path) {
      if (!path) return null;
      if (mediaByPath.has(path)) return mediaByPath.get(path);
      const f = zip.file(path);
      if (!f) return null;
      const ext = path.split('.').pop().toLowerCase();
      const mime = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp', tif: 'image/tiff', tiff: 'image/tiff' }[ext];
      if (!mime) return null;
      const id = 'm' + (++mediaN) + '_' + Math.random().toString(36).slice(2, 6);
      media[id] = `data:${mime};base64,${await f.async('base64')}`;
      mediaByPath.set(path, id);
      return id;
    }

    // Master/Layout/Theme-Infos zwischenspeichern
    const ctxCache = new Map();
    async function masterCtx(masterPath) {
      if (ctxCache.has(masterPath)) return ctxCache.get(masterPath);
      const mx = await parse(masterPath);
      const mrels = await relsOf(masterPath);
      const themeRel = byType(mrels, 'theme');
      const tx = themeRel ? await parse(themeRel.path) : null;
      const theme = { colors: {}, major: 'Calibri Light', minor: 'Calibri' };
      if (tx) {
        const scheme = desc(tx, 'clrScheme');
        if (scheme) for (const c of scheme.children) {
          const v = c.firstElementChild;
          if (v) theme.colors[c.localName] = ((v.localName === 'sysClr' ? v.getAttribute('lastClr') : v.getAttribute('val')) || (v.getAttribute('val') === 'window' ? 'FFFFFF' : '000000')).toUpperCase();
        }
        const maj = desc(tx, 'majorFont'), min = desc(tx, 'minorFont');
        if (maj && desc(maj, 'latin')) theme.major = desc(maj, 'latin').getAttribute('typeface') || theme.major;
        if (min && desc(min, 'latin')) theme.minor = desc(min, 'latin').getAttribute('typeface') || theme.minor;
      }
      const clrMap = {};
      const cm = mx && desc(mx, 'clrMap');
      if (cm) for (const a of cm.attributes) clrMap[a.name] = a.value;
      const ctx = { xml: mx, rels: mrels, theme, clrMap, path: masterPath, txStyles: mx && desc(mx, 'txStyles') };
      ctxCache.set(masterPath, ctx);
      return ctx;
    }

    function colorOf(el, ctx, phClr) {
      if (!el) return null;
      const c = el.localName === 'solidFill' ? el.firstElementChild : el;
      if (!c) return null;
      let base = null;
      switch (c.localName) {
        case 'srgbClr': base = c.getAttribute('val'); break;
        case 'sysClr': base = c.getAttribute('lastClr') || (c.getAttribute('val') === 'window' ? 'FFFFFF' : '000000'); break;
        case 'prstClr': base = { black: '000000', white: 'FFFFFF', red: 'FF0000', blue: '0000FF', green: '008000', yellow: 'FFFF00', gray: '808080' }[c.getAttribute('val')] || '000000'; break;
        case 'schemeClr': {
          let v = c.getAttribute('val');
          if (v === 'phClr') { base = phClr ? phClr.replace('#', '') : null; break; }
          v = ctx.clrMap[v] || ({ tx1: 'dk1', tx2: 'dk2', bg1: 'lt1', bg2: 'lt2' }[v]) || v;
          base = ctx.theme.colors[v] || null;
          break;
        }
        case 'scrgbClr': base = [c.getAttribute('r'), c.getAttribute('g'), c.getAttribute('b')].map(x => Math.round(+x / 100000 * 255).toString(16).padStart(2, '0')).join(''); break;
      }
      if (!base) return null;
      return applyMods(base.toUpperCase(), c);
    }
    function fillOf(spPr, ctx, styleEl) {
      if (spPr) {
        if (kid(spPr, 'noFill')) return null;
        const sf = kid(spPr, 'solidFill');
        if (sf) return colorOf(sf, ctx);
        const gf = kid(spPr, 'gradFill');
        if (gf) { const gs = desc(gf, 'gs'); if (gs) return colorOf(gs.firstElementChild, ctx); }
      }
      if (styleEl) {
        const fr = kid(styleEl, 'fillRef');
        if (fr && fr.getAttribute('idx') !== '0' && fr.firstElementChild) return colorOf(fr.firstElementChild, ctx);
      }
      return null;
    }
    function lineOf(spPr, ctx, styleEl) {
      const ln = spPr && kid(spPr, 'ln');
      if (ln && kid(ln, 'noFill')) return null;
      let color = ln && kid(ln, 'solidFill') ? colorOf(kid(ln, 'solidFill'), ctx) : null;
      let width = ln && ln.getAttribute('w') ? +ln.getAttribute('w') / 12700 : null;
      if (!color && styleEl) {
        const lr = kid(styleEl, 'lnRef');
        if (lr && lr.getAttribute('idx') !== '0' && lr.firstElementChild) color = colorOf(lr.firstElementChild, ctx);
      }
      if (!color) return null;
      return { color, width: width || 1 };
    }
    function xfrmOf(spPr) {
      const x = spPr && kid(spPr, 'xfrm');
      if (!x) return null;
      const off = kid(x, 'off'), ext = kid(x, 'ext');
      if (!off || !ext) return null;
      return {
        x: +off.getAttribute('x') / EMU, y: +off.getAttribute('y') / EMU, w: +ext.getAttribute('cx') / EMU, h: +ext.getAttribute('cy') / EMU,
        rot: x.getAttribute('rot') ? +x.getAttribute('rot') / 60000 : 0, flipH: x.getAttribute('flipH') === '1', flipV: x.getAttribute('flipV') === '1'
      };
    }
    const phInfo = (sp) => {
      const ph = desc(kid(sp, 'nvSpPr') || kid(sp, 'nvPicPr') || kid(sp, 'nvGraphicFramePr'), 'ph');
      if (!ph) return null;
      return { type: ph.getAttribute('type') || 'body', idx: ph.getAttribute('idx') };
    };
    function findPh(tree, ph) {
      if (!tree || !ph) return null;
      const sps = Array.from(tree.getElementsByTagNameNS('*', 'sp'));
      const norm = (t) => t === 'ctrTitle' ? 'title' : t;
      if (ph.idx != null) {
        const m = sps.find(sp => { const p = phInfo(sp); return p && p.idx === ph.idx; });
        if (m) return m;
      }
      return sps.find(sp => { const p = phInfo(sp); return p && norm(p.type) === norm(ph.type); }) ||
        (ph.type === 'subTitle' || ph.type === 'obj' ? sps.find(sp => { const p = phInfo(sp); return p && p.type === 'body'; }) : null);
    }

    // Absatz-/Zeicheneigenschaften aus einer lvlNpPr-Kette
    function applyPPr(t, pPr, ctx) {
      if (!pPr) return;
      if (pPr.getAttribute('algn')) t.align = { l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify' }[pPr.getAttribute('algn')] || 'left';
      if (kid(pPr, 'buNone')) t.bullet = null;
      if (kid(pPr, 'buChar')) t.bullet = 'bullet';
      if (kid(pPr, 'buAutoNum')) t.bullet = 'number';
      const d = kid(pPr, 'defRPr');
      if (d) applyRPr(t, d, ctx);
    }
    function applyRPr(t, rPr, ctx) {
      if (!rPr) return;
      if (rPr.getAttribute('sz')) t.size = +rPr.getAttribute('sz') / 100;
      if (rPr.getAttribute('b') != null) t.bold = rPr.getAttribute('b') === '1' || rPr.getAttribute('b') === 'true';
      if (rPr.getAttribute('i') != null) t.italic = rPr.getAttribute('i') === '1' || rPr.getAttribute('i') === 'true';
      if (rPr.getAttribute('u') != null) t.underline = rPr.getAttribute('u') !== 'none';
      if (rPr.getAttribute('strike') != null) t.strike = rPr.getAttribute('strike') !== 'noStrike';
      const sf = kid(rPr, 'solidFill');
      if (sf) t.color = colorOf(sf, ctx, t.phClr);
      const hl = kid(rPr, 'highlight');
      if (hl) t.highlight = colorOf(hl.firstElementChild, ctx);
      const lat = kid(rPr, 'latin');
      if (lat && lat.getAttribute('typeface')) {
        const tf = lat.getAttribute('typeface');
        t.font = tf === '+mj-lt' ? ctx.theme.major : tf === '+mn-lt' ? ctx.theme.minor : tf;
      }
    }
    function lvlEl(listStyle, lvl) { return listStyle ? kid(listStyle, `lvl${lvl + 1}pPr`) : null; }

    async function textFromTxBody(txBody, chain, ctx, fontScale, phClr) {
      // chain: Liste von lstStyle-Elementen (niedrige Prioritaet zuerst)
      const paras = [];
      for (const p of kids(txBody, 'p')) {
        const pPr = kid(p, 'pPr');
        const lvl = pPr && pPr.getAttribute('lvl') ? +pPr.getAttribute('lvl') : 0;
        const t = { size: 18, color: '#000000', font: ctx.theme.minor, align: 'left', bullet: null, phClr };
        for (const ls of chain) applyPPr(t, lvlEl(ls, lvl), ctx);
        if (phClr) t.color = phClr; // Schriftfarbe aus der Formenvorlage (z. B. weiss auf Akzentfarbe)
        applyPPr(t, pPr, ctx);
        const runs = [];
        for (const r of p.children) {
          if (r.localName === 'r' || r.localName === 'fld') {
            const rt = { ...t };
            applyRPr(rt, kid(r, 'rPr'), ctx);
            const text = (kid(r, 't') || { textContent: '' }).textContent;
            runs.push({ text, ...rt, size: rt.size * fontScale });
          } else if (r.localName === 'br') runs.push({ br: true });
        }
        const endRPr = kid(p, 'endParaRPr');
        const pt = { ...t };
        if (!runs.length && endRPr) applyRPr(pt, endRPr, ctx);
        paras.push({ lvl, align: t.align, bullet: t.bullet, runs, base: { ...pt, size: pt.size * fontScale } });
      }
      return paras;
    }

    function parasToHtml(paras, base) {
      let html = '';
      const stack = []; // offene Listen
      const close = (depth) => { while (stack.length > depth) html += `</li></${stack.pop()}>`; };
      for (const p of paras) {
        let inner = '';
        for (const r of p.runs) {
          if (r.br) { inner += '<br>'; continue; }
          let t = esc(r.text);
          if (!t) continue;
          const css = [];
          if (r.font && r.font !== base.font) css.push(`font-family:'${r.font.replace(/'/g, '')}'`);
          if (r.size && Math.abs(r.size - base.size) > 0.1) css.push(`font-size:${Math.round(r.size * 10) / 10}pt`);
          if (r.color && r.color.toLowerCase() !== (base.color || '').toLowerCase()) css.push(`color:${r.color}`);
          if (r.highlight) css.push(`background-color:${r.highlight}`);
          if (r.bold && !base.bold) t = `<b>${t}</b>`;
          if (!r.bold && base.bold) css.push('font-weight:normal');
          if (r.italic) t = `<i>${t}</i>`;
          if (r.underline) t = `<u>${t}</u>`;
          if (r.strike) t = `<s>${t}</s>`;
          inner += css.length ? `<span style="${css.join(';')}">${t}</span>` : t;
        }
        if (!inner) {
          // Leerer Absatz: Groesse erhalten
          const s = p.base && p.base.size && Math.abs(p.base.size - base.size) > 0.1 ? ` style="font-size:${p.base.size}pt"` : '';
          inner = `<br${s ? '' : ''}>`;
          if (s) inner = `<span${s}><br></span>`;
        }
        const st = p.align && p.align !== 'left' ? ` style="text-align:${p.align}"` : '';
        if (p.bullet) {
          const tag = p.bullet === 'number' ? 'ol' : 'ul';
          const depth = p.lvl + 1;
          if (stack.length > depth) close(depth);
          if (stack.length === depth && stack[depth - 1] !== tag) close(depth - 1);
          if (stack.length === depth) html += '</li>';
          while (stack.length < depth) { html += `<${stack.length < depth - 1 ? 'ul' : tag}${st}>`; stack.push(stack.length < depth - 1 ? 'ul' : tag); if (stack.length < depth) html += '<li style="list-style:none">'; }
          html += `<li${st}>${inner}`;
        } else {
          close(0);
          html += `<p${st}>${inner}</p>`;
        }
      }
      close(0);
      return html;
    }

    let elN = 0;
    const newId = () => 'e' + (++elN) + Math.random().toString(36).slice(2, 6);

    async function shapesFrom(tree, ctx, rels, layoutTree, masterTree, transform, opts) {
      const out = [];
      for (const node of tree.children) {
        const name = node.localName;
        if (name === 'grpSp') {
          const gx = kid(kid(node, 'grpSpPr'), 'xfrm');
          let tf = transform;
          if (gx) {
            const off = kid(gx, 'off'), ext = kid(gx, 'ext'), chOff = kid(gx, 'chOff'), chExt = kid(gx, 'chExt');
            if (off && ext && chOff && chExt) {
              const sx = (+ext.getAttribute('cx') || 1) / (+chExt.getAttribute('cx') || 1), sy = (+ext.getAttribute('cy') || 1) / (+chExt.getAttribute('cy') || 1);
              const ox = +off.getAttribute('x') / EMU, oy = +off.getAttribute('y') / EMU, cx = +chOff.getAttribute('x') / EMU, cy = +chOff.getAttribute('y') / EMU;
              tf = (b) => transform({ ...b, x: ox + (b.x - cx) * sx, y: oy + (b.y - cy) * sy, w: b.w * sx, h: b.h * sy });
            }
          }
          out.push(...await shapesFrom(node, ctx, rels, layoutTree, masterTree, tf, opts));
          continue;
        }
        if (name === 'AlternateContent') {
          const ch = kid(node, 'Choice') || kid(node, 'Fallback');
          if (ch) out.push(...await shapesFrom(ch, ctx, rels, layoutTree, masterTree, transform, opts));
          continue;
        }
        if (name === 'sp' || name === 'cxnSp') {
          const ph = phInfo(node);
          if (opts.skipPlaceholders && ph) continue;
          if (ph && ['dt', 'ftr', 'sldNum'].includes(ph.type) && opts.isSlide) {
            // Fusszeilen nur uebernehmen, wenn sie Text enthalten
            const txt = desc(node, 'txBody');
            if (!txt || !txt.textContent.trim() || ph.type === 'sldNum') continue;
          }
          const spPr = kid(node, 'spPr');
          const layoutPh = ph ? findPh(layoutTree, ph) : null;
          const masterPh = ph ? findPh(masterTree, ph) : null;
          let box = xfrmOf(spPr) || (layoutPh && xfrmOf(kid(layoutPh, 'spPr'))) || (masterPh && xfrmOf(kid(masterPh, 'spPr')));
          if (!box) continue;
          box = transform(box);
          const styleEl = kid(node, 'style');
          const geom = spPr && desc(spPr, 'prstGeom');
          let prst = geom ? geom.getAttribute('prst') : (name === 'cxnSp' ? 'line' : 'rect');
          if (/Connector|line/i.test(prst)) prst = 'line';
          const fill = fillOf(spPr, ctx, styleEl) || (layoutPh ? fillOf(kid(layoutPh, 'spPr'), ctx) : null);
          const line = lineOf(spPr, ctx, styleEl);
          const txBody = kid(node, 'txBody');
          // Textformat-Kette
          const chain = [];
          const phType = ph ? (ph.type === 'ctrTitle' ? 'title' : ph.type) : null;
          if (ph) {
            const ts = ctx.txStyles;
            chain.push(phType === 'title' ? desc(ts, 'titleStyle') : desc(ts, 'bodyStyle'));
            if (masterPh) chain.push(desc(masterPh, 'lstStyle'));
            if (layoutPh) chain.push(desc(layoutPh, 'lstStyle'));
          } else {
            chain.push(defaultTextStyle, ctx.txStyles ? desc(ctx.txStyles, 'otherStyle') : null);
          }
          if (txBody) chain.push(kid(txBody, 'lstStyle'));
          let phClr = null;
          if (styleEl) { const fr = kid(styleEl, 'fontRef'); if (fr && fr.firstElementChild) phClr = colorOf(fr.firstElementChild, ctx); }
          const bodyPr = txBody && kid(txBody, 'bodyPr');
          const norm = bodyPr && kid(bodyPr, 'normAutofit');
          const fontScale = norm && norm.getAttribute('fontScale') ? +norm.getAttribute('fontScale') / 100000 : 1;
          let anchor = bodyPr && bodyPr.getAttribute('anchor');
          for (const src of [layoutPh, masterPh]) {
            if (!anchor && src) { const bp = desc(src, 'bodyPr'); if (bp && bp.getAttribute('anchor')) anchor = bp.getAttribute('anchor'); }
          }
          const paras = txBody ? await textFromTxBody(txBody, chain.filter(Boolean), ctx, fontScale, phClr) : [];
          const hasText = paras.some(p => p.runs.some(r => r.text && r.text.trim()));
          // Grundformat = erster Textlauf (sonst Absatzende-Format)
          const firstRun = paras.flatMap(p => p.runs).find(r => r.text && r.text.trim());
          const base = firstRun ? { font: firstRun.font, size: firstRun.size, color: firstRun.color, bold: firstRun.bold }
            : (paras[0] ? { ...paras[0].base } : { size: 18, color: '#000000', font: ctx.theme.minor });
          if (phClr && !paras.some(p => p.runs.some(r => r.color))) base.color = base.color || phClr;
          const el = {
            id: newId(), x: box.x, y: box.y, w: box.w, h: box.h, rot: box.rot || 0,
            font: base.font, size: Math.round(base.size * 10) / 10, color: base.color, bold: !!base.bold,
            va: anchor === 'ctr' ? 'middle' : anchor === 'b' ? 'bottom' : 'top', html: hasText ? parasToHtml(paras, base) : ''
          };
          if (box.flipH) el.flipH = true;
          if (box.flipV) el.flipV = true;
          if (ph) {
            el.ph = phType === 'title' ? 'title' : phType === 'subTitle' ? 'subtitle' : 'body';
            if (!hasText && opts.isSlide) el.html = '';
          }
          const isPlainBox = (prst === 'rect' && !fill && !line) || (ph && prst === 'rect');
          if (isPlainBox) {
            el.type = 'text';
            if (fill) el.fill = fill;
            if (line) el.line = line;
            if (!hasText && !ph) continue;
          } else {
            el.type = 'shape';
            el.shape = SHAPES.includes(prst) ? prst : (/star/.test(prst) ? 'star5' : /arrow/i.test(prst) ? 'rightArrow' : /round/i.test(prst) ? 'roundRect' : /ellipse|oval|circle/i.test(prst) ? 'ellipse' : 'rect');
            el.fill = fill;
            el.line = line;
            if (!anchor) el.va = 'middle';
            if (!paras.some(p => p.align && p.align !== 'left')) el.align = 'center';
          }
          out.push(el);
          continue;
        }
        if (name === 'pic') {
          if (opts.skipPlaceholders && phInfo(node)) continue;
          const spPr = kid(node, 'spPr');
          let box = xfrmOf(spPr);
          if (!box) { const ph = phInfo(node); const lp = ph ? findPh(layoutTree, ph) : null; box = lp ? xfrmOf(kid(lp, 'spPr')) : null; }
          if (!box) continue;
          box = transform(box);
          const blip = desc(node, 'blip');
          const rel = rels[rid(blip)];
          const id = rel ? await loadMedia(rel.path) : null;
          if (!id) continue;
          const el = { id: newId(), type: 'image', media: id, x: box.x, y: box.y, w: box.w, h: box.h, rot: box.rot || 0 };
          if (box.flipH) el.flipH = true;
          if (box.flipV) el.flipV = true;
          out.push(el);
          continue;
        }
        if (name === 'graphicFrame') {
          const xf = kid(node, 'xfrm');
          if (!xf) continue;
          const off = kid(xf, 'off'), ext = kid(xf, 'ext');
          const box = transform({ x: +off.getAttribute('x') / EMU, y: +off.getAttribute('y') / EMU, w: +ext.getAttribute('cx') / EMU, h: +ext.getAttribute('cy') / EMU });
          const tbl = desc(node, 'tbl');
          if (tbl) {
            const grid = kids(kid(tbl, 'tblGrid'), 'gridCol').map(g => +g.getAttribute('w') || 1);
            const cells = [];
            let fontSize = 18;
            for (const tr of kids(tbl, 'tr')) {
              const row = [];
              for (const tc of kids(tr, 'tc')) {
                if (tc.getAttribute('hMerge') === '1' || tc.getAttribute('vMerge') === '1') { row.push({ html: '' }); continue; }
                const tb = kid(tc, 'txBody');
                const paras = tb ? await textFromTxBody(tb, [defaultTextStyle].filter(Boolean), ctx, 1) : [];
                if (paras[0] && paras[0].runs[0] && paras[0].runs[0].size) fontSize = paras[0].runs[0].size;
                const base = paras[0] ? paras[0].base : { size: fontSize, color: '#000000', font: ctx.theme.minor };
                const cell = { html: parasToHtml(paras, base).replace(/^<p>(.*)<\/p>$/s, '$1') };
                const tcPr = kid(tc, 'tcPr');
                const bg = tcPr ? fillOf(tcPr, ctx) : null;
                if (bg) cell.bg = bg;
                const r0 = paras[0] && paras[0].runs.find(r => r.text);
                if (r0 && r0.color && r0.color !== '#000000') cell.color = r0.color;
                if (r0 && r0.bold) cell.bold = true;
                row.push(cell);
              }
              cells.push(row);
            }
            const total = grid.reduce((a, b) => a + b, 0);
            const tp = kid(tbl, 'tblPr');
            out.push({ id: newId(), type: 'table', x: box.x, y: box.y, w: box.w, h: box.h, cells, colW: grid.map(g => g / total),
              header: !tp || tp.getAttribute('firstRow') === '1', band: !tp || tp.getAttribute('bandRow') === '1', size: Math.round(fontSize), font: ctx.theme.minor, accent: '#' + (ctx.theme.colors.accent1 || '4472C4') });
            continue;
          }
          const chartRef = node.getElementsByTagNameNS(NS_C, 'chart')[0];
          if (chartRef) {
            const rel = rels[rid(chartRef, 'id')];
            const cx = rel && await parse(rel.path);
            const chart = cx && parseChart(cx);
            if (chart) out.push({ id: newId(), type: 'chart', x: box.x, y: box.y, w: box.w, h: box.h, chart });
          }
        }
      }
      return out;
    }

    function parseChart(x) {
      const kinds = ['barChart', 'bar3DChart', 'lineChart', 'line3DChart', 'areaChart', 'pieChart', 'pie3DChart', 'doughnutChart', 'scatterChart'];
      let plot = null, kind = null;
      for (const k of kinds) { plot = x.getElementsByTagNameNS(NS_C, k)[0]; if (plot) { kind = k; break; } }
      if (!plot) return null;
      let type = 'column';
      if (/bar/.test(kind)) { const d = plot.getElementsByTagNameNS(NS_C, 'barDir')[0]; type = d && d.getAttribute('val') === 'bar' ? 'bar' : 'column'; }
      else if (/line|area|scatter/.test(kind)) type = 'line';
      else type = 'pie';
      const titleEl = x.getElementsByTagNameNS(NS_C, 'title')[0];
      const title = titleEl ? Array.from(titleEl.getElementsByTagNameNS('*', 't')).map(t => t.textContent).join('') : '';
      const cache = (el) => el ? Array.from(el.getElementsByTagNameNS(NS_C, 'pt')).sort((a, b) => +a.getAttribute('idx') - +b.getAttribute('idx')).map(p => (p.getElementsByTagNameNS(NS_C, 'v')[0] || {}).textContent) : [];
      const series = [];
      let categories = [];
      for (const ser of plot.getElementsByTagNameNS(NS_C, 'ser')) {
        const tx = ser.getElementsByTagNameNS(NS_C, 'tx')[0];
        const name = tx ? (cache(tx)[0] || (tx.getElementsByTagNameNS(NS_C, 'v')[0] || {}).textContent || '') : '';
        const cat = ser.getElementsByTagNameNS(NS_C, 'cat')[0] || ser.getElementsByTagNameNS(NS_C, 'xVal')[0];
        const val = ser.getElementsByTagNameNS(NS_C, 'val')[0] || ser.getElementsByTagNameNS(NS_C, 'yVal')[0];
        if (cat && !categories.length) categories = cache(cat);
        series.push({ name: name || `Reihe ${series.length + 1}`, values: cache(val).map(v => +v || 0) });
      }
      if (!series.length) return null;
      if (!categories.length) categories = series[0].values.map((_, i) => String(i + 1));
      return { type, title, categories, series };
    }

    function bgOf(cSld, ctx) {
      const bg = cSld && kid(cSld, 'bg');
      if (!bg) return null;
      const bgPr = kid(bg, 'bgPr');
      if (bgPr) {
        const blip = desc(bgPr, 'blip');
        if (blip) return { blip: rid(blip) };
        const c = fillOf(bgPr, ctx);
        if (c) return { color: c };
      }
      const ref = kid(bg, 'bgRef');
      if (ref && ref.firstElementChild) return { color: colorOf(ref.firstElementChild, ctx) };
      return null;
    }

    const slides = [];
    const sldIds = presXml.getElementsByTagNameNS('*', 'sldId');
    for (const sid of sldIds) {
      const rel = presRels[rid(sid, 'id')];
      if (!rel || !rel.path) continue;
      const sx = await parse(rel.path);
      if (!sx) continue;
      if (sx.documentElement.getAttribute('show') === '0') continue;
      const srels = await relsOf(rel.path);
      const layoutRel = byType(srels, 'slideLayout');
      const lx = layoutRel ? await parse(layoutRel.path) : null;
      const lrels = layoutRel ? await relsOf(layoutRel.path) : {};
      const masterRel = byType(lrels, 'slideMaster');
      const ctx = masterRel ? await masterCtx(masterRel.path) : { theme: { colors: {}, major: 'Calibri Light', minor: 'Calibri' }, clrMap: {}, txStyles: null, rels: {} };
      const masterTree = ctx.xml ? desc(ctx.xml, 'spTree') : null;
      const layoutTree = lx ? desc(lx, 'spTree') : null;
      const cSld = desc(sx, 'cSld');

      // Hintergrund: Folie > Layout > Master
      let bg = bgOf(cSld, ctx), bgRels = srels;
      if (!bg && lx) { bg = bgOf(desc(lx, 'cSld'), ctx); bgRels = lrels; }
      if (!bg && ctx.xml) { bg = bgOf(desc(ctx.xml, 'cSld'), ctx); bgRels = ctx.rels; }
      const slide = { id: 's' + Math.random().toString(36).slice(2, 9), bg: '#ffffff', bgImage: null, notes: '', transition: 'none', els: [] };
      if (bg && bg.color) slide.bg = bg.color;
      if (bg && bg.blip && bgRels[bg.blip]) slide.bgImage = await loadMedia(bgRels[bg.blip].path);

      // Design-Elemente aus Master und Layout (ohne Platzhalter)
      const showMaster = sx.documentElement.getAttribute('showMasterSp') !== '0';
      const ident = (b) => b;
      if (showMaster && lx && lx.documentElement.getAttribute('showMasterSp') !== '0' && masterTree) {
        slide.els.push(...await shapesFrom(masterTree, ctx, ctx.rels, null, null, ident, { skipPlaceholders: true }));
      }
      if (showMaster && layoutTree) slide.els.push(...await shapesFrom(layoutTree, ctx, lrels, null, masterTree, ident, { skipPlaceholders: true }));
      const tree = desc(cSld, 'spTree');
      if (tree) slide.els.push(...await shapesFrom(tree, ctx, srels, layoutTree, masterTree, ident, { isSlide: true }));

      // Notizen
      const notesRel = byType(srels, 'notesSlide');
      if (notesRel) {
        const nx = await parse(notesRel.path);
        if (nx) {
          for (const sp of nx.getElementsByTagNameNS('*', 'sp')) {
            const ph = phInfo(sp);
            if (ph && ph.type === 'body') {
              slide.notes = kids(desc(sp, 'txBody'), 'p').map(p => Array.from(p.getElementsByTagNameNS('*', 't')).map(t => t.textContent).join('')).join('\n').trim();
            }
          }
        }
      }
      // Uebergang
      const tr = sx.getElementsByTagNameNS('*', 'transition')[0];
      if (tr) {
        const eff = Array.from(tr.children).map(c => c.localName).find(n => n !== 'sndAc' && n !== 'extLst');
        slide.transition = { fade: 'fade', push: 'push', wipe: 'wipe', zoom: 'zoom', cover: 'cover', pull: 'cover', split: 'wipe', dissolve: 'fade', cut: 'none' }[eff] || (eff ? 'fade' : 'none');
      }
      slides.push(slide);
    }
    if (!slides.length) throw new Error('Die Präsentation enthält keine Folien.');
    return { w: W, h: H, slides, media, showNumbers: false, themeId: null };
  }

  window.RexPptx = { exportPptx, importPptx, SHAPES };
})();
