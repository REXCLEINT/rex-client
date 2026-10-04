/* REX Praesentation */
(function () {
  'use strict';
  const { $, $$, esc } = Rex.util;
  const canvas = $('#canvas'), stage = $('#stage'), stageInner = $('#stageInner'), selLayer = $('#selLayer');

  // =====================================================================
  //  Designs und Layouts
  // =====================================================================
  const THEMES = [
    { id: 'office', name: 'Office', bg: '#ffffff', title: '#1f1f1f', text: '#262626', accent: '#4472c4', fontT: 'Calibri Light', fontB: 'Calibri' },
    { id: 'slate', name: 'Schiefer', bg: '#f4f6f8', title: '#1d3557', text: '#334155', accent: '#e63946', fontT: 'Segoe UI Semibold', fontB: 'Segoe UI' },
    { id: 'night', name: 'Nacht', bg: '#1b1d2a', title: '#ffffff', text: '#cfd3ea', accent: '#7aa2f7', fontT: 'Segoe UI Light', fontB: 'Segoe UI' },
    { id: 'ocean', name: 'Ozean', bg: '#0f4c75', title: '#ffffff', text: '#dcefff', accent: '#3fc1c9', fontT: 'Segoe UI Semibold', fontB: 'Segoe UI' },
    { id: 'sunset', name: 'Sonne', bg: '#fff4e6', title: '#c2410c', text: '#431407', accent: '#f97316', fontT: 'Georgia', fontB: 'Calibri' },
    { id: 'forest', name: 'Wald', bg: '#eef6ee', title: '#1b5e20', text: '#2e3b2e', accent: '#43a047', fontT: 'Cambria', fontB: 'Calibri' },
    { id: 'berry', name: 'Beere', bg: '#2d1b3d', title: '#ffd6f5', text: '#f3e8ff', accent: '#e879f9', fontT: 'Trebuchet MS', fontB: 'Trebuchet MS' },
    { id: 'gaming', name: 'Gaming', bg: '#0b0b0f', title: '#39ff14', text: '#e6e6e6', accent: '#ff2e63', fontT: 'Impact', fontB: 'Segoe UI' }
  ];
  const LAYOUTS = {
    title: { name: 'Titelfolie', els: [
      { ph: 'title', x: 120, y: 140, w: 720, h: 140, size: 48, align: 'center', va: 'bottom' },
      { ph: 'subtitle', x: 120, y: 292, w: 720, h: 90, size: 24, align: 'center' }] },
    titleContent: { name: 'Titel und Inhalt', els: [
      { ph: 'title', x: 66, y: 29, w: 828, h: 105, size: 40, va: 'middle' },
      { ph: 'body', x: 66, y: 144, w: 828, h: 343, size: 24 }] },
    section: { name: 'Abschnittsüberschrift', els: [
      { ph: 'title', x: 66, y: 190, w: 828, h: 120, size: 48, va: 'bottom' },
      { ph: 'subtitle', x: 66, y: 318, w: 828, h: 64, size: 22 }] },
    twoContent: { name: 'Zwei Inhalte', els: [
      { ph: 'title', x: 66, y: 29, w: 828, h: 105, size: 40, va: 'middle' },
      { ph: 'body', x: 66, y: 144, w: 408, h: 343, size: 22 },
      { ph: 'body', x: 486, y: 144, w: 408, h: 343, size: 22 }] },
    titleOnly: { name: 'Nur Titel', els: [{ ph: 'title', x: 66, y: 29, w: 828, h: 105, size: 40, va: 'middle' }] },
    blank: { name: 'Leer', els: [] }
  };
  const PH_TEXT = { title: 'Titel hinzufügen', subtitle: 'Untertitel hinzufügen', body: 'Text hinzufügen' };
  const FONTS = ['Calibri', 'Calibri Light', 'Aptos', 'Arial', 'Arial Black', 'Cambria', 'Century Gothic', 'Comic Sans MS', 'Consolas', 'Georgia', 'Impact',
    'Segoe UI', 'Segoe UI Light', 'Segoe UI Semibold', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana'];
  const SIZES = [8, 10, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 96];
  const SHAPE_LIST = [
    ['rect', 'Rechteck', 'square'], ['roundRect', 'Abgerundet', 'square'], ['ellipse', 'Ellipse', 'circle'], ['triangle', 'Dreieck', 'triangle'],
    ['rightArrow', 'Pfeil', 'arrow-right'], ['line', 'Linie', 'minus'], ['star5', 'Stern', 'star'], ['hexagon', 'Sechseck', 'hexagon'],
    ['diamond', 'Raute', 'diamond'], ['chevron', 'Chevron', 'chevron-right'], ['heart', 'Herz', 'heart']
  ];
  const TRANSITIONS = [['none', 'Ohne', 'circle-off'], ['fade', 'Verblassen', 'blend'], ['push', 'Schieben', 'arrow-left-to-line'],
    ['wipe', 'Wischen', 'arrow-right-from-line'], ['zoom', 'Zoom', 'zoom-in'], ['cover', 'Aufdecken', 'arrow-up-from-line']];

  // =====================================================================
  //  Modell
  // =====================================================================
  let pres;
  let cur = 0;               // aktive Folie
  let selected = [];         // ausgewaehlte Element-IDs
  let editingId = null;      // Element im Textbearbeitungsmodus
  let zoom = 1, autoFit = true;
  const uid = (p = 'e') => p + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  const theme = () => THEMES.find(t => t.id === pres.themeId) || THEMES[0];
  const slide = () => pres.slides[cur];
  const elById = (id, s = slide()) => s.els.find(e => e.id === id);

  function makeSlide(layoutId = 'titleContent') {
    const th = theme();
    const L = LAYOUTS[layoutId] || LAYOUTS.blank;
    const sx = pres.w / 960, sy = pres.h / 540;
    return {
      id: uid('s'), layout: layoutId, bg: th.bg, bgImage: null, notes: '', transition: 'none',
      els: L.els.map(d => ({
        id: uid(), type: 'text', ph: d.ph, x: d.x * sx, y: d.y * sy, w: d.w * sx, h: d.h * sy, rot: 0, html: '',
        font: d.ph === 'title' ? th.fontT : th.fontB, size: d.size, color: d.ph === 'title' ? th.title : th.text,
        align: d.align || 'left', va: d.va || 'top'
      }))
    };
  }
  function newPresentation() {
    pres = { w: 960, h: 540, themeId: 'office', showNumbers: false, slides: [], media: {} };
    pres.slides.push(makeSlide('title'));
    cur = 0;
  }

  // =====================================================================
  //  Verlauf (Bilder werden getrennt gespeichert)
  // =====================================================================
  const hist = { stack: [], index: -1 };
  const snapshot = () => JSON.stringify({ w: pres.w, h: pres.h, themeId: pres.themeId, showNumbers: pres.showNumbers, slides: pres.slides, cur });
  function resetHistory() { hist.stack = [snapshot()]; hist.index = 0; }
  function commit(opts = {}) {
    const s = snapshot();
    if (hist.stack[hist.index] !== s) {
      hist.stack = hist.stack.slice(0, hist.index + 1);
      hist.stack.push(s);
      if (hist.stack.length > 120) hist.stack.shift();
      hist.index = hist.stack.length - 1;
      Rex.setDirty(true);
    }
    if (!opts.noRender) renderAll();
  }
  function restore(s) {
    const o = JSON.parse(s);
    Object.assign(pres, { w: o.w, h: o.h, themeId: o.themeId, showNumbers: o.showNumbers, slides: o.slides });
    cur = Math.min(o.cur ?? cur, pres.slides.length - 1);
    selected = selected.filter(id => elById(id));
    renderAll();
    Rex.setDirty(true);
  }
  function undo() { if (editingId) { stopEditing(); } if (hist.index > 0) restore(hist.stack[--hist.index]); }
  function redo() { if (editingId) stopEditing(); if (hist.index < hist.stack.length - 1) restore(hist.stack[++hist.index]); }

  // =====================================================================
  //  Darstellung von Folien
  // =====================================================================
  function shapeSVG(el) {
    const w = Math.max(1, el.w), h = Math.max(1, el.h);
    const lw = el.line && el.line.width ? el.line.width * 1.333 : 0;
    const stroke = el.line && el.line.width ? `stroke="${el.line.color}" stroke-width="${lw}"` : 'stroke="none"';
    const fill = el.fill ? `fill="${el.fill}"` : 'fill="none"';
    const i = lw / 2;
    let body;
    switch (el.shape) {
      case 'ellipse': body = `<ellipse cx="${w / 2}" cy="${h / 2}" rx="${Math.max(0, w / 2 - i)}" ry="${Math.max(0, h / 2 - i)}" ${fill} ${stroke}/>`; break;
      case 'roundRect': { const r = Math.min(w, h) * 0.1667; body = `<rect x="${i}" y="${i}" width="${w - lw}" height="${h - lw}" rx="${r}" ${fill} ${stroke}/>`; break; }
      case 'triangle': body = `<polygon points="${w / 2},${i} ${w - i},${h - i} ${i},${h - i}" ${fill} ${stroke} stroke-linejoin="round"/>`; break;
      case 'diamond': body = `<polygon points="${w / 2},${i} ${w - i},${h / 2} ${w / 2},${h - i} ${i},${h / 2}" ${fill} ${stroke}/>`; break;
      case 'rightArrow': case 'leftArrow': {
        const hl = Math.min(w, h * 0.5 * 2) * 0.5, t = h * 0.25;
        let pts = [[0, t], [w - hl, t], [w - hl, 0], [w, h / 2], [w - hl, h], [w - hl, h - t], [0, h - t]];
        if (el.shape === 'leftArrow') pts = pts.map(([x, y]) => [w - x, y]);
        body = `<polygon points="${pts.map(p => p.join(',')).join(' ')}" ${fill} ${stroke} stroke-linejoin="round"/>`;
        break;
      }
      case 'line': {
        const c = (el.line && el.line.color) || el.fill || '#000';
        const sw = (el.line && el.line.width ? el.line.width : 2) * 1.333;
        const [x1, y1, x2, y2] = el.flipV ? [0, h, w, 0] : [0, 0, w, h];
        body = `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${c}" stroke-width="${sw}" stroke-linecap="round"/>`;
        break;
      }
      case 'star5': {
        const pts = [];
        for (let k = 0; k < 10; k++) {
          const a = -Math.PI / 2 + k * Math.PI / 5, r = k % 2 ? 0.2 : 0.5;
          pts.push(`${w / 2 + Math.cos(a) * w * r},${h / 2 + 0.05 * h + Math.sin(a) * h * r * 1.05}`);
        }
        body = `<polygon points="${pts.join(' ')}" ${fill} ${stroke} stroke-linejoin="round"/>`;
        break;
      }
      case 'hexagon': { const d = Math.min(w / 2, h * 0.29); body = `<polygon points="${d},${i} ${w - d},${i} ${w - i},${h / 2} ${w - d},${h - i} ${d},${h - i} ${i},${h / 2}" ${fill} ${stroke}/>`; break; }
      case 'pentagon': body = `<polygon points="${w / 2},${i} ${w - i},${h * 0.38} ${w * 0.81},${h - i} ${w * 0.19},${h - i} ${i},${h * 0.38}" ${fill} ${stroke}/>`; break;
      case 'chevron': { const d = Math.min(w, h) * 0.5; body = `<polygon points="${i},${i} ${w - d},${i} ${w - i},${h / 2} ${w - d},${h - i} ${i},${h - i} ${d},${h / 2}" ${fill} ${stroke}/>`; break; }
      case 'parallelogram': { const d = Math.min(w, h) * 0.25; body = `<polygon points="${d},${i} ${w - i},${i} ${w - d},${h - i} ${i},${h - i}" ${fill} ${stroke}/>`; break; }
      case 'heart': body = `<path d="M${w / 2},${h * 0.25} C${w / 2},${h * 0.05} ${w * 0.05},${-h * 0.02} ${w * 0.05},${h * 0.3} C${w * 0.05},${h * 0.55} ${w / 2},${h * 0.75} ${w / 2},${h - i} C${w / 2},${h * 0.75} ${w * 0.95},${h * 0.55} ${w * 0.95},${h * 0.3} C${w * 0.95},${-h * 0.02} ${w / 2},${h * 0.05} ${w / 2},${h * 0.25} Z" ${fill} ${stroke}/>`; break;
      case 'cloud': body = `<ellipse cx="${w / 2}" cy="${h / 2}" rx="${w / 2 - i}" ry="${h / 2 - i}" ${fill} ${stroke}/>`; break;
      default: body = `<rect x="${i}" y="${i}" width="${Math.max(0, w - lw)}" height="${Math.max(0, h - lw)}" ${fill} ${stroke}/>`;
    }
    const flip = el.shape !== 'line' && (el.flipH || el.flipV) ? ` style="transform:scale(${el.flipH ? -1 : 1},${el.flipV ? -1 : 1})"` : '';
    return `<svg viewBox="0 0 ${w} ${h}" preserveAspectRatio="none"${flip}>${body}</svg>`;
  }
  function tableCellStyle(el, ri, cell) {
    const accent = el.accent || theme().accent;
    const header = el.header !== false && ri === 0;
    const band = el.band !== false && ri % 2 === 1;
    const bg = cell.bg || (header ? accent : band ? mix(accent, '#ffffff', 0.78) : mix(accent, '#ffffff', 0.9));
    return { bg, color: cell.color || (header ? '#ffffff' : '#1f1f1f'), bold: cell.bold || header, align: cell.align };
  }
  function mix(a, b, t) {
    const p = (h) => [1, 3, 5].map(i => parseInt(h.replace('#', '').padEnd(6, '0').slice(i - 1, i + 1), 16));
    const [x, y] = [p(a), p(b)];
    return '#' + x.map((v, i) => Math.round(v * (1 - t) + y[i] * t).toString(16).padStart(2, '0')).join('');
  }
  function textBlock(el, mode) {
    const ph = el.ph && !stripTags(el.html).trim() && !/<img/.test(el.html || '');
    if (ph && mode !== 'edit') return '';
    const style = [`font-family:'${(el.font || 'Calibri').replace(/'/g, '')}',Calibri,Carlito,sans-serif`, `font-size:${el.size || 18}pt`, `color:${el.color || '#000'}`,
      `text-align:${el.align || 'left'}`];
    if (el.bold) style.push('font-weight:bold');
    if (el.shadow) style.push('text-shadow:2px 3px 4px rgba(0,0,0,.35)');
    const va = el.va || (el.type === 'shape' ? 'middle' : 'top');
    return `<div class="txt va-${va}" style="${style.join(';')}"><div class="txt-inner" data-ph="${ph ? PH_TEXT[el.ph] || '' : ''}">${el.html || ''}</div></div>`;
  }
  const stripTags = (h) => String(h || '').replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>');

  function elHTML(el, mode) {
    const tf = [];
    if (el.rot) tf.push(`rotate(${el.rot}deg)`);
    const base = `left:${el.x}px;top:${el.y}px;width:${el.w}px;height:${el.h}px;${tf.length ? `transform:${tf.join(' ')};` : ''}`;
    const emptyPh = el.ph && !stripTags(el.html).trim();
    const cls = `el el-${el.type}${emptyPh && mode === 'edit' ? ' placeholder-empty' : ''}`;
    let inner = '';
    let extra = '';
    if (el.type === 'text') {
      if (el.fill) extra += `background:${el.fill};`;
      if (el.line && el.line.width) extra += `box-shadow:inset 0 0 0 ${el.line.width * 1.333}px ${el.line.color};`;
      inner = textBlock(el, mode);
      if (!inner && mode !== 'edit' && !el.fill) return '';
    } else if (el.type === 'shape') {
      inner = shapeSVG(el) + (el.html ? textBlock(el, mode) : (mode === 'edit' ? textBlock(el, mode) : ''));
    } else if (el.type === 'image') {
      const src = pres.media[el.media] || '';
      const flip = el.flipH || el.flipV ? ` style="transform:scale(${el.flipH ? -1 : 1},${el.flipV ? -1 : 1})"` : '';
      inner = `<img src="${src}" alt=""${flip}>`;
      if (el.line && el.line.width) extra += `outline:${el.line.width * 1.333}px solid ${el.line.color};outline-offset:-${el.line.width * 1.333}px;`;
    } else if (el.type === 'table') {
      const total = el.colW ? el.colW.reduce((a, b) => a + b, 0) : 0;
      const cols = el.cells[0] ? el.cells[0].length : 0;
      inner = `<table class="ptable" style="font-family:'${el.font || 'Calibri'}',Calibri,sans-serif;font-size:${el.size || 18}pt">` +
        (total ? `<colgroup>${el.colW.map(f => `<col style="width:${f / total * 100}%">`).join('')}</colgroup>` : `<colgroup>${'<col>'.repeat(cols)}</colgroup>`) +
        el.cells.map((row, ri) => `<tr>${row.map((c, ci) => {
          const st = tableCellStyle(el, ri, c);
          return `<td data-r="${ri}" data-c="${ci}" style="background:${st.bg};color:${st.color};${st.bold ? 'font-weight:bold;' : ''}${st.align ? `text-align:${st.align};` : ''}">${c.html || ''}</td>`;
        }).join('')}</tr>`).join('') + '</table>';
    } else if (el.type === 'chart') {
      const dark = isDark(slide0(el));
      inner = `<div class="chart-el">${RexCharts.render(el.chart, el.w, el.h, { transparent: true, dark })}</div>`;
    }
    return `<div class="${cls}" data-id="${el.id}" style="${base}${extra}">${inner}</div>`;
  }
  // Hilfsfunktion: zu welcher Folie gehoert ein Element (fuer dunkle Diagramme)
  function slide0(el) { return pres.slides.find(s => s.els.includes(el)) || slide(); }
  function isDark(sl) {
    const c = (sl.bg || '#ffffff').replace('#', '');
    const [r, g, b] = [0, 2, 4].map(i => parseInt(c.slice(i, i + 2), 16));
    return (r * 299 + g * 587 + b * 114) / 1000 < 110;
  }
  function slideHTML(sl, mode, index) {
    const bg = sl.bgImage && pres.media[sl.bgImage] ? `background:${sl.bg || '#fff'} url('${pres.media[sl.bgImage]}') center/cover no-repeat;` : `background:${sl.bg || '#fff'};`;
    let html = `<div class="slide" style="width:${pres.w}px;height:${pres.h}px;${bg}">`;
    for (const el of sl.els) html += elHTML(el, mode);
    if (pres.showNumbers && index != null) {
      html += `<div style="position:absolute;right:30px;bottom:12px;font:11pt Calibri,Carlito,sans-serif;color:${isDark(sl) ? '#bbb' : '#8c8c8c'}">${index + 1}</div>`;
    }
    return html + '</div>';
  }

  // =====================================================================
  //  Editor: Leinwand, Auswahl, Miniaturen
  // =====================================================================
  function fitZoom() {
    const aw = stage.clientWidth - 60, ah = stage.clientHeight - 60;
    return Math.max(0.1, Math.min(aw / pres.w, ah / pres.h));
  }
  function applyZoom() {
    if (autoFit) zoom = fitZoom();
    canvas.style.width = pres.w + 'px';
    canvas.style.height = pres.h + 'px';
    canvas.style.transform = `scale(${zoom})`;
    stageInner.style.width = pres.w * zoom + 60 + 'px';
    stageInner.style.height = pres.h * zoom + 60 + 'px';
    $('#zoomRange').value = Math.round(zoom * 100);
    $('#zoomLabel').textContent = Math.round(zoom * 100) + ' %';
  }
  function renderCanvas() {
    if (editingId) return;
    canvas.innerHTML = slideHTML(slide(), 'edit', cur);
    applyZoom();
    renderSelection();
  }
  function renderAll() {
    if (cur >= pres.slides.length) cur = pres.slides.length - 1;
    renderCanvas();
    renderThumbs();
    updateStatus();
    updateRibbon();
    $('#notes').value = slide().notes || '';
  }
  function updateStatus() {
    $('#stSlide').textContent = `Folie ${cur + 1} von ${pres.slides.length}`;
  }

  const thumbCache = new Map();
  function renderThumbs() {
    const list = $('#slideList');
    const tw = list.clientWidth - 40 || 160;
    const scale = tw / pres.w;
    const rows = $$('.thumb-row', list);
    while (rows.length > pres.slides.length) rows.pop().remove();
    pres.slides.forEach((sl, i) => {
      let row = list.children[i];
      if (!row) {
        row = document.createElement('div');
        row.className = 'thumb-row';
        row.innerHTML = '<div class="num"></div><div class="thumb" draggable="true"><div class="thumb-inner"></div><span class="trans-mark"></span></div>';
        list.appendChild(row);
      }
      row.dataset.i = i;
      row.querySelector('.num').textContent = i + 1;
      const th = row.querySelector('.thumb');
      th.classList.toggle('active', i === cur);
      th.style.height = pres.h * scale + 'px';
      const key = JSON.stringify(sl) + scale + pres.showNumbers + i;
      if (thumbCache.get(row) !== key) {
        thumbCache.set(row, key);
        const inner = row.querySelector('.thumb-inner');
        inner.innerHTML = slideHTML(sl, 'thumb', i);
        inner.style.transform = `scale(${scale})`;
        row.querySelector('.trans-mark').textContent = sl.transition && sl.transition !== 'none' ? '★' : '';
      }
    });
    const act = list.children[cur];
    if (act) {
      const r = act.getBoundingClientRect(), lr = list.getBoundingClientRect();
      if (r.top < lr.top || r.bottom > lr.bottom) act.scrollIntoView({ block: 'nearest' });
    }
  }

  function renderSelection() {
    let html = '';
    for (const id of selected) {
      const el = elById(id);
      if (!el) continue;
      const one = selected.length === 1;
      const x = el.x * zoom, y = el.y * zoom, w = el.w * zoom, h = el.h * zoom;
      const rot = el.rot ? `transform:rotate(${el.rot}deg);` : '';
      html += `<div class="sel-box${editingId === id ? ' editing' : ''}" data-id="${id}" style="left:${x}px;top:${y}px;width:${w}px;height:${h}px;${rot}">`;
      if (one && editingId !== id) {
        for (const hnd of el.type === 'shape' && el.shape === 'line' ? ['nw', 'se'] : ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']) html += `<div class="h ${hnd}" data-h="${hnd}"></div>`;
        if (el.type !== 'table') html += '<div class="h rot" data-h="rot" title="Drehen"></div>';
      }
      html += '</div>';
    }
    selLayer.innerHTML = html;
    selLayer.style.width = pres.w * zoom + 'px';
    selLayer.style.height = pres.h * zoom + 'px';
  }
  function select(ids) {
    if (editingId && !ids.includes(editingId)) stopEditing();
    selected = ids;
    renderSelection();
    updateRibbon();
  }

  // =====================================================================
  //  Textbearbeitung
  // =====================================================================
  function startEditing(id, opts = {}) {
    const el = elById(id);
    if (!el || !['text', 'shape', 'table'].includes(el.type)) return;
    if (el.type === 'shape' && el.shape === 'line') return;
    if (editingId === id) return;
    if (editingId) stopEditing();
    selected = [id];
    const node = canvas.querySelector(`.el[data-id="${id}"]`);
    if (!node) return;
    editingId = id;
    node.classList.add('editing');
    if (el.type === 'table') {
      $$('td', node).forEach(td => { td.contentEditable = 'true'; });
      const td = opts.point ? document.elementFromPoint(opts.point.x, opts.point.y)?.closest('td') : null;
      (td || node.querySelector('td')).focus();
      renderSelection();
      return;
    }
    let inner = node.querySelector('.txt-inner');
    if (!inner) {
      node.insertAdjacentHTML('beforeend', textBlock({ ...el, html: '' }, 'edit'));
      inner = node.querySelector('.txt-inner');
    }
    node.classList.remove('placeholder-empty');
    if (el.ph && !stripTags(el.html).trim()) {
      inner.dataset.ph = '';
      inner.innerHTML = el.ph === 'body' ? '<ul><li><br></li></ul>' : '<p><br></p>';
    }
    inner.contentEditable = 'true';
    inner.focus();
    const sel = getSelection();
    if (opts.point && document.caretRangeFromPoint) {
      const r = document.caretRangeFromPoint(opts.point.x, opts.point.y);
      if (r && inner.contains(r.startContainer)) { sel.removeAllRanges(); sel.addRange(r); }
      else placeCaretEnd(inner);
    } else if (opts.selectAll) {
      const r = document.createRange(); r.selectNodeContents(inner); sel.removeAllRanges(); sel.addRange(r);
    } else placeCaretEnd(inner);
    renderSelection();
    updateRibbon();
  }
  function placeCaretEnd(node) {
    const r = document.createRange();
    r.selectNodeContents(node);
    r.collapse(false);
    const sel = getSelection();
    sel.removeAllRanges();
    sel.addRange(r);
  }
  function stopEditing() {
    if (!editingId) return;
    const id = editingId;
    const el = elById(id);
    const node = canvas.querySelector(`.el[data-id="${id}"]`);
    editingId = null;
    if (el && node) {
      if (el.type === 'table') {
        $$('td', node).forEach(td => { el.cells[+td.dataset.r][+td.dataset.c].html = cleanHTML(td.innerHTML); });
      } else {
        const inner = node.querySelector('.txt-inner');
        if (inner) {
          let html = cleanHTML(inner.innerHTML);
          if (!stripTags(html).trim() && !/<img/.test(html)) html = '';
          el.html = html;
        }
      }
    }
    getSelection().removeAllRanges();
    commit();
  }
  function cleanHTML(html) {
    return html.replace(/<span[^>]*>\s*<\/span>/g, '').replace(/ contenteditable="[^"]*"/g, '').replace(/^(<br>)+$/, '');
  }
  function autoGrow() {
    if (!editingId) return;
    const el = elById(editingId);
    const node = canvas.querySelector(`.el[data-id="${editingId}"]`);
    if (!el || !node || el.type !== 'text' || el.ph) return;
    const inner = node.querySelector('.txt-inner');
    const need = inner.scrollHeight + 10;
    if (need > el.h) { el.h = need; node.style.height = need + 'px'; renderSelection(); }
  }
  canvas.addEventListener('input', () => { autoGrow(); Rex.setDirty(true); });
  canvas.addEventListener('paste', e => {
    if (!editingId) return;
    e.preventDefault();
    const text = e.clipboardData.getData('text/plain');
    document.execCommand('insertText', false, text);
  });

  // Formatierung auf Text anwenden (im Bearbeitungsmodus auf die Auswahl, sonst auf das ganze Element)
  function withText(fn) {
    const ids = editingId ? [editingId] : selected.filter(id => { const e = elById(id); return e && ['text', 'shape', 'table'].includes(e.type); });
    if (!ids.length) { Rex.toast('Bitte zuerst ein Textfeld auswählen.'); return; }
    if (editingId) { fn(); autoGrow(); return; }
    for (const id of ids) {
      const el = elById(id);
      if (el.type === 'table') {
        startEditing(id);
        const node = canvas.querySelector(`.el[data-id="${id}"] table`);
        const r = document.createRange(); r.selectNodeContents(node);
        getSelection().removeAllRanges(); getSelection().addRange(r);
        fn();
        stopEditing();
        continue;
      }
      if (el.ph && !stripTags(el.html).trim()) continue;
      startEditing(id, { selectAll: true });
      fn();
      stopEditing();
    }
    select(ids);
  }
  function tcmd(cmd, val) {
    document.execCommand('styleWithCSS', false, true);
    withText(() => document.execCommand(cmd, false, val ?? null));
  }
  function setFontSize(pt) {
    const ids = editingId ? [] : selected;
    if (!editingId) {
      // ganze Elemente: Grundgroesse setzen und eingebettete Groessen entfernen
      for (const id of ids) {
        const el = elById(id);
        if (!el || !['text', 'shape', 'table'].includes(el.type)) continue;
        el.size = pt;
        if (el.html) el.html = el.html.replace(/font-size:\s*[^;"]+;?/g, '');
      }
      commit();
      return;
    }
    document.execCommand('styleWithCSS', false, true);
    document.execCommand('fontSize', false, '7');
    const node = canvas.querySelector(`.el[data-id="${editingId}"]`);
    $$('font[size="7"]', node).forEach(f => { const s = document.createElement('span'); s.style.fontSize = pt + 'pt'; while (f.firstChild) s.appendChild(f.firstChild); f.replaceWith(s); });
    $$('span[style*="xxx-large"]', node).forEach(s => { s.style.fontSize = pt + 'pt'; });
    autoGrow();
  }
  function stepFont(dir) {
    const el = elById(editingId || selected[0]);
    if (!el) return;
    let curPt = el.size || 18;
    if (editingId) {
      const s = getSelection();
      const n = s.anchorNode && (s.anchorNode.nodeType === 3 ? s.anchorNode.parentElement : s.anchorNode);
      if (n) curPt = Math.round(parseFloat(getComputedStyle(n).fontSize) * 0.75);
    }
    const next = dir > 0 ? SIZES.find(x => x > curPt) || curPt + 4 : [...SIZES].reverse().find(x => x < curPt) || Math.max(6, curPt - 2);
    setFontSize(next);
  }

  // =====================================================================
  //  Maus auf der Leinwand
  // =====================================================================
  const toSlide = (e) => {
    const r = canvas.getBoundingClientRect();
    return { x: (e.clientX - r.left) / zoom, y: (e.clientY - r.top) / zoom };
  };
  let guidesOn = true;
  function snap(box, movingIds) {
    if (!guidesOn) return { dx: 0, dy: 0, lines: [] };
    const th = 6 / zoom;
    const xs = [0, pres.w / 2, pres.w], ys = [0, pres.h / 2, pres.h];
    for (const e of slide().els) {
      if (movingIds.includes(e.id) || e.rot) continue;
      xs.push(e.x, e.x + e.w / 2, e.x + e.w);
      ys.push(e.y, e.y + e.h / 2, e.y + e.h);
    }
    let best = { dx: 0, dy: 0, lx: null, ly: null };
    let mx = th, my = th;
    for (const v of [box.x, box.x + box.w / 2, box.x + box.w]) {
      for (const g of xs) if (Math.abs(g - v) < mx) { mx = Math.abs(g - v); best.dx = g - v; best.lx = g; }
    }
    for (const v of [box.y, box.y + box.h / 2, box.y + box.h]) {
      for (const g of ys) if (Math.abs(g - v) < my) { my = Math.abs(g - v); best.dy = g - v; best.ly = g; }
    }
    const lines = [];
    if (best.lx != null) lines.push({ v: true, at: best.lx });
    if (best.ly != null) lines.push({ v: false, at: best.ly });
    return { dx: best.dx, dy: best.dy, lines };
  }
  function showGuides(lines) {
    $$('.guide', selLayer).forEach(g => g.remove());
    for (const l of lines) {
      const g = document.createElement('div');
      g.className = 'guide ' + (l.v ? 'v' : 'h');
      if (l.v) Object.assign(g.style, { left: l.at * zoom + 'px', top: 0, height: pres.h * zoom + 'px' });
      else Object.assign(g.style, { top: l.at * zoom + 'px', left: 0, width: pres.w * zoom + 'px' });
      selLayer.appendChild(g);
    }
  }

  stage.addEventListener('mousedown', e => {
    if (e.button !== 0) return;
    const handle = e.target.closest('.h');
    if (handle) { e.preventDefault(); startResize(e, handle.dataset.h); return; }
    const node = e.target.closest('.slide-canvas .el');
    if (editingId && node && node.dataset.id === editingId) return; // Text bearbeiten
    if (editingId) stopEditing();
    if (!node) {
      // Auswahlrahmen aufziehen
      if (!e.target.closest('.stage')) return;
      e.preventDefault();
      if (!e.shiftKey) select([]);
      const start = toSlide(e);
      const mq = document.createElement('div');
      mq.className = 'marquee';
      selLayer.appendChild(mq);
      const mv = ev => {
        const p = toSlide(ev);
        const x = Math.min(start.x, p.x), y = Math.min(start.y, p.y), w = Math.abs(p.x - start.x), h = Math.abs(p.y - start.y);
        Object.assign(mq.style, { left: x * zoom + 'px', top: y * zoom + 'px', width: w * zoom + 'px', height: h * zoom + 'px' });
        mq._r = { x, y, w, h };
      };
      const up = () => {
        document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up);
        const r = mq._r;
        mq.remove();
        if (r && r.w > 3 && r.h > 3) {
          const ids = slide().els.filter(el => el.x < r.x + r.w && el.x + el.w > r.x && el.y < r.y + r.h && el.y + el.h > r.y).map(el => el.id);
          select(e.shiftKey ? [...new Set([...selected, ...ids])] : ids);
        }
      };
      document.addEventListener('mousemove', mv); document.addEventListener('mouseup', up);
      return;
    }
    e.preventDefault();
    const id = node.dataset.id;
    const wasOnlySelected = selected.length === 1 && selected[0] === id;
    if (e.shiftKey || e.ctrlKey) select(selected.includes(id) ? selected.filter(x => x !== id) : [...selected, id]);
    else if (!selected.includes(id)) select([id]);
    const start = toSlide(e);
    const origin = selected.map(sid => { const el = elById(sid); return { el, x: el.x, y: el.y }; });
    let moved = false;
    const mv = ev => {
      const p = toSlide(ev);
      let dx = p.x - start.x, dy = p.y - start.y;
      if (!moved && Math.abs(dx) + Math.abs(dy) < 3 / zoom) return;
      moved = true;
      if (ev.shiftKey) { if (Math.abs(dx) > Math.abs(dy)) dy = 0; else dx = 0; }
      // Ausrichtungshilfe an der Gesamtbox
      const bx = Math.min(...origin.map(o => o.x)) + dx, by = Math.min(...origin.map(o => o.y)) + dy;
      const bw = Math.max(...origin.map(o => o.x + o.el.w)) - Math.min(...origin.map(o => o.x));
      const bh = Math.max(...origin.map(o => o.y + o.el.h)) - Math.min(...origin.map(o => o.y));
      const s = ev.altKey ? { dx: 0, dy: 0, lines: [] } : snap({ x: bx, y: by, w: bw, h: bh }, selected);
      for (const o of origin) {
        o.el.x = Math.round((o.x + dx + s.dx) * 10) / 10;
        o.el.y = Math.round((o.y + dy + s.dy) * 10) / 10;
        const n = canvas.querySelector(`.el[data-id="${o.el.id}"]`);
        if (n) { n.style.left = o.el.x + 'px'; n.style.top = o.el.y + 'px'; }
      }
      renderSelection();
      showGuides(s.lines);
    };
    const up = ev => {
      document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up);
      showGuides([]);
      if (moved) commit();
      else if (wasOnlySelected && !e.shiftKey) {
        const el = elById(id);
        if (el && ['text', 'shape', 'table'].includes(el.type)) startEditing(id, { point: { x: ev.clientX, y: ev.clientY } });
      }
    };
    document.addEventListener('mousemove', mv); document.addEventListener('mouseup', up);
  });
  stage.addEventListener('dblclick', e => {
    const node = e.target.closest('.slide-canvas .el');
    if (!node) return;
    const el = elById(node.dataset.id);
    if (!el) return;
    if (el.type === 'chart') { editChart(el); return; }
    if (el.type === 'image') { replaceImage(el); return; }
    if (editingId !== el.id) startEditing(el.id, { point: { x: e.clientX, y: e.clientY } });
  });

  function startResize(e, h) {
    const el = elById(selected[0]);
    if (!el) return;
    const start = toSlide(e);
    const o = { x: el.x, y: el.y, w: el.w, h: el.h, rot: el.rot || 0 };
    const keepRatio = el.type === 'image' && h.length === 2;
    const ratio = o.w / Math.max(1, o.h);
    const nodeOf = () => canvas.querySelector(`.el[data-id="${el.id}"]`);
    const mv = ev => {
      const p = toSlide(ev);
      const node = nodeOf();
      if (!node) return;
      if (h === 'rot') {
        const cx = o.x + o.w / 2, cy = o.y + o.h / 2;
        let a = Math.atan2(p.y - cy, p.x - cx) * 180 / Math.PI + 90;
        if (ev.shiftKey) a = Math.round(a / 15) * 15;
        else if (Math.abs(((a % 90) + 90) % 90) < 4 || Math.abs(((a % 90) + 90) % 90) > 86) a = Math.round(a / 90) * 90;
        el.rot = Math.round(((a % 360) + 360) % 360);
        if (el.rot === 360) el.rot = 0;
        node.style.transform = el.rot ? `rotate(${el.rot}deg)` : '';
        renderSelection();
        return;
      }
      let dx = p.x - start.x, dy = p.y - start.y;
      let { x, y, w, h: hh } = o;
      if (h.includes('e')) w = o.w + dx;
      if (h.includes('s')) hh = o.h + dy;
      if (h.includes('w')) { w = o.w - dx; x = o.x + dx; }
      if (h.includes('n')) { hh = o.h - dy; y = o.y + dy; }
      if ((keepRatio && !ev.shiftKey) || (!keepRatio && ev.shiftKey && h.length === 2)) {
        if (Math.abs(w / ratio) > Math.abs(hh)) hh = w / ratio; else w = hh * ratio;
        if (h.includes('w')) x = o.x + o.w - w;
        if (h.includes('n')) y = o.y + o.h - hh;
      }
      const isLine = el.type === 'shape' && el.shape === 'line';
      const min = isLine ? 0 : 8;
      if (w < min) { if (h.includes('w')) x -= min - w; w = min; }
      if (hh < min) { if (h.includes('n')) y -= min - hh; hh = min; }
      Object.assign(el, { x: Math.round(x * 10) / 10, y: Math.round(y * 10) / 10, w: Math.round(w * 10) / 10, h: Math.round(hh * 10) / 10 });
      node.outerHTML = elHTML(el, 'edit');
      renderSelection();
    };
    const up = () => { document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up); commit(); };
    document.addEventListener('mousemove', mv); document.addEventListener('mouseup', up);
  }

  // Kontextmenue
  stage.addEventListener('contextmenu', e => {
    const node = e.target.closest('.slide-canvas .el');
    if (editingId && node && node.dataset.id === editingId) {
      const el = elById(editingId);
      if (el.type !== 'table') return; // Rechtschreibmenue
      e.preventDefault();
      const td = e.target.closest('td');
      const ri = td ? +td.dataset.r : el.cells.length - 1, ci = td ? +td.dataset.c : 0;
      const fake = { getBoundingClientRect: () => ({ left: e.clientX, right: e.clientX, top: e.clientY, bottom: e.clientY }) };
      const done = () => { stopEditing(); select([el.id]); };
      stopEditing();
      Rex.menu(fake, [
        { label: 'Zeile darunter einfügen', icon: 'between-horizontal-start', action: () => { el.cells.splice(ri + 1, 0, el.cells[0].map(() => ({ html: '' }))); el.h += 40; done(); } },
        { label: 'Spalte rechts einfügen', icon: 'between-vertical-start', action: () => { el.cells.forEach(r => r.splice(ci + 1, 0, { html: '' })); el.colW = null; done(); } },
        { label: 'Zeile löschen', icon: 'rows-3', action: () => { if (el.cells.length > 1) { el.cells.splice(ri, 1); el.h = Math.max(30, el.h - 40); } done(); } },
        { label: 'Spalte löschen', icon: 'columns-3', action: () => { if (el.cells[0].length > 1) { el.cells.forEach(r => r.splice(ci, 1)); el.colW = null; } done(); } },
        '-',
        { label: el.header !== false ? 'Kopfzeile aus' : 'Kopfzeile an', icon: 'panel-top', action: () => { el.header = el.header === false; done(); } }
      ]);
      return;
    }
    e.preventDefault();
    if (node && !selected.includes(node.dataset.id)) select([node.dataset.id]);
    const fake = { getBoundingClientRect: () => ({ left: e.clientX, right: e.clientX, top: e.clientY, bottom: e.clientY }) };
    const items = selected.length ? [
      { label: 'Ausschneiden', icon: 'scissors', action: () => { copyEls(); deleteSelected(); } },
      { label: 'Kopieren', icon: 'copy', action: () => copyEls() },
      { label: 'Einfügen', icon: 'clipboard-paste', action: () => pasteEls() },
      { label: 'Duplizieren', icon: 'copy-plus', action: () => duplicateEls() },
      '-',
      { label: 'In den Vordergrund', icon: 'bring-to-front', action: () => arrange('front') },
      { label: 'In den Hintergrund', icon: 'send-to-back', action: () => arrange('back') },
      '-',
      { label: 'Löschen', icon: 'trash-2', action: () => deleteSelected() }
    ] : [
      { label: 'Einfügen', icon: 'clipboard-paste', action: () => pasteEls() },
      { label: 'Neue Folie', icon: 'file-plus-2', action: () => commands.newSlide() },
      { label: 'Hintergrund formatieren', icon: 'paint-bucket', action: () => $('#bgBtn').click() }
    ];
    Rex.menu(fake, items);
  });

  // =====================================================================
  //  Elemente: einfuegen, kopieren, anordnen
  // =====================================================================
  function addEl(el) {
    el.id = uid();
    el.rot = el.rot || 0;
    slide().els.push(el);
    commit();
    select([el.id]);
    return el;
  }
  function insertText() {
    const th = theme();
    const el = addEl({ type: 'text', x: pres.w / 2 - 160, y: pres.h / 2 - 25, w: 320, h: 50, html: '', font: th.fontB, size: 18, color: th.text, align: 'left', va: 'top' });
    startEditing(el.id);
  }
  function insertShape(shape) {
    const th = theme();
    const isLine = shape === 'line';
    addEl({ type: 'shape', shape, x: pres.w / 2 - 90, y: pres.h / 2 - (isLine ? 0 : 60), w: 180, h: isLine ? 0 : 120,
      fill: isLine ? null : th.accent, line: isLine ? { color: th.accent, width: 2.5 } : { color: mix(th.accent, '#000000', 0.3), width: 1 },
      html: '', font: th.fontB, size: 18, color: '#ffffff', align: 'center', va: 'middle' });
  }
  async function addImageFromSrc(src, opts = {}) {
    if (!src) return;
    const { w, h } = await Rex.util.imageSize(src);
    const id = uid('m');
    pres.media[id] = src;
    const max = Math.min(pres.w * 0.7, pres.h * 0.7 * (w / h));
    const ww = Math.min(w, max), hh = ww * h / w;
    addEl({ type: 'image', media: id, x: opts.x ?? (pres.w - ww) / 2, y: opts.y ?? (pres.h - hh) / 2, w: ww, h: hh });
  }
  async function replaceImage(el) {
    const src = await Rex.util.readImageFile();
    if (!src) return;
    const id = uid('m');
    pres.media[id] = src;
    el.media = id;
    commit();
  }
  async function insertTable() {
    const v = await Rex.dialog({ title: 'Tabelle einfügen', fields: [
      { name: 'cols', label: 'Spalten', type: 'number', value: 3, min: 1, max: 12 },
      { name: 'rows', label: 'Zeilen', type: 'number', value: 4, min: 1, max: 30 }] });
    if (!v) return;
    const rows = Math.max(1, Math.min(30, v.rows)), cols = Math.max(1, Math.min(12, v.cols));
    const w = Math.min(pres.w - 120, cols * 180), h = Math.min(pres.h - 160, rows * 42);
    addEl({ type: 'table', x: (pres.w - w) / 2, y: Math.max(120, (pres.h - h) / 2), w, h, cells: Array.from({ length: rows }, () => Array.from({ length: cols }, () => ({ html: '' }))),
      header: true, band: true, size: 18, font: theme().fontB, accent: theme().accent });
  }
  function chartDataText(ch) {
    return ['Kategorie;' + ch.series.map(s => s.name).join(';'), ...ch.categories.map((c, i) => c + ';' + ch.series.map(s => String(s.values[i] ?? 0).replace('.', ',')).join(';'))].join('\n');
  }
  function parseChartText(text, type, title) {
    const lines = text.trim().split(/\r?\n/).map(l => l.split(/[;\t]/));
    const head = lines.shift() || [];
    const series = head.slice(1).map((n, i) => ({ name: n.trim() || `Reihe ${i + 1}`, values: lines.map(l => parseFloat(String(l[i + 1] || '0').replace(/\./g, '').replace(',', '.')) || 0) }));
    return { type, title, categories: lines.map(l => (l[0] || '').trim()), series };
  }
  async function chartDialog(ch) {
    const v = await Rex.dialog({ title: ch ? 'Diagramm bearbeiten' : 'Diagramm einfügen', html: '<p style="color:var(--text-2);margin:0 0 8px">Daten: erste Zeile = Reihennamen, danach je Zeile eine Kategorie. Trennzeichen: Semikolon.</p>', fields: [
      { name: 'type', label: 'Typ', type: 'select', value: ch ? ch.type : 'column', options: [{ value: 'column', label: 'Säulen' }, { value: 'bar', label: 'Balken' }, { value: 'line', label: 'Linie' }, { value: 'pie', label: 'Kreis' }] },
      { name: 'title', label: 'Titel', value: ch ? ch.title : 'Umsatz' },
      { name: 'data', label: 'Daten', type: 'textarea', rows: 7, value: ch ? chartDataText(ch) : 'Kategorie;2024;2025\nQ1;12;15\nQ2;18;21\nQ3;14;24\nQ4;22;28' }] });
    if (!v) return null;
    return parseChartText(v.data, v.type, v.title);
  }
  async function insertChart(type) {
    const base = { type, title: 'Umsatz', categories: ['Q1', 'Q2', 'Q3', 'Q4'], series: [{ name: '2024', values: [12, 18, 14, 22] }, { name: '2025', values: [15, 21, 24, 28] }] };
    const ch = await chartDialog(base);
    if (!ch) return;
    addEl({ type: 'chart', x: pres.w / 2 - 280, y: pres.h / 2 - 170, w: 560, h: 340, chart: ch });
  }
  async function editChart(el) {
    const ch = await chartDialog(el.chart);
    if (ch) { el.chart = ch; commit(); }
  }
  let elClip = null;
  function copyEls() {
    const els = selected.map(id => elById(id)).filter(Boolean);
    if (!els.length) return;
    elClip = JSON.parse(JSON.stringify(els));
    try { navigator.clipboard.writeText(els.map(e => stripTags(e.html || '')).join('\n')); } catch { /* */ }
  }
  function pasteEls() {
    if (!elClip) return;
    const ids = [];
    for (const e of elClip) {
      const c = JSON.parse(JSON.stringify(e));
      c.id = uid();
      c.x += 16; c.y += 16;
      e.x += 16; e.y += 16;
      slide().els.push(c);
      ids.push(c.id);
    }
    commit();
    select(ids);
  }
  function duplicateEls() { copyEls(); pasteEls(); }
  function deleteSelected() {
    if (!selected.length) return;
    slide().els = slide().els.filter(e => !selected.includes(e.id));
    selected = [];
    commit();
  }
  function arrange(where) {
    const els = slide().els;
    const sel = els.filter(e => selected.includes(e.id));
    const rest = els.filter(e => !selected.includes(e.id));
    if (where === 'front') slide().els = [...rest, ...sel];
    else if (where === 'back') slide().els = [...sel, ...rest];
    else {
      for (const id of (where === 'forward' ? [...selected].reverse() : selected)) {
        const i = els.findIndex(e => e.id === id);
        const j = where === 'forward' ? i + 1 : i - 1;
        if (j >= 0 && j < els.length) [els[i], els[j]] = [els[j], els[i]];
      }
    }
    commit();
  }
  function align(how) {
    const els = selected.map(id => elById(id)).filter(Boolean);
    if (!els.length) return;
    const ref = els.length > 1 ? {
      x: Math.min(...els.map(e => e.x)), y: Math.min(...els.map(e => e.y)),
      r: Math.max(...els.map(e => e.x + e.w)), b: Math.max(...els.map(e => e.y + e.h))
    } : { x: 0, y: 0, r: pres.w, b: pres.h };
    for (const e of els) {
      if (how === 'left') e.x = ref.x;
      if (how === 'center') e.x = (ref.x + ref.r) / 2 - e.w / 2;
      if (how === 'right') e.x = ref.r - e.w;
      if (how === 'top') e.y = ref.y;
      if (how === 'middle') e.y = (ref.y + ref.b) / 2 - e.h / 2;
      if (how === 'bottom') e.y = ref.b - e.h;
    }
    commit();
  }

  // =====================================================================
  //  Folien
  // =====================================================================
  function goTo(i) {
    if (editingId) stopEditing();
    cur = Math.max(0, Math.min(pres.slides.length - 1, i));
    selected = [];
    renderAll();
  }
  function newSlide(layout) {
    if (editingId) stopEditing();
    const l = layout || (slide() && slide().layout === 'title' ? 'titleContent' : (slide() ? slide().layout || 'titleContent' : 'titleContent'));
    const s = makeSlide(l);
    if (slide() && slide().bgImage) s.bgImage = slide().bgImage;
    if (slide()) s.bg = slide().bg;
    pres.slides.splice(cur + 1, 0, s);
    cur++;
    selected = [];
    commit();
  }
  function duplicateSlide() {
    const c = JSON.parse(JSON.stringify(slide()));
    c.id = uid('s');
    c.els.forEach(e => { e.id = uid(); });
    pres.slides.splice(cur + 1, 0, c);
    cur++;
    commit();
  }
  function deleteSlide() {
    if (pres.slides.length === 1) { Rex.toast('Die letzte Folie kann nicht gelöscht werden.'); return; }
    pres.slides.splice(cur, 1);
    cur = Math.min(cur, pres.slides.length - 1);
    selected = [];
    commit();
  }
  function applyLayout(layoutId) {
    const s = slide();
    const L = LAYOUTS[layoutId];
    const th = theme();
    const sx = pres.w / 960, sy = pres.h / 540;
    const old = s.els.filter(e => e.ph);
    const others = s.els.filter(e => !e.ph);
    const used = new Set();
    const els = L.els.map(d => {
      const prev = old.find(o => o.ph === d.ph && !used.has(o)) || (d.ph !== 'title' ? old.find(o => o.ph !== 'title' && !used.has(o)) : null);
      if (prev) used.add(prev);
      return {
        id: prev ? prev.id : uid(), type: 'text', ph: d.ph, x: d.x * sx, y: d.y * sy, w: d.w * sx, h: d.h * sy, rot: 0,
        html: prev ? prev.html : '', font: prev ? prev.font : (d.ph === 'title' ? th.fontT : th.fontB), size: d.size,
        color: prev ? prev.color : (d.ph === 'title' ? th.title : th.text), align: d.align || 'left', va: d.va || 'top'
      };
    });
    // Nicht verwendete Platzhalter mit Inhalt behalten
    for (const o of old) if (!used.has(o) && stripTags(o.html).trim()) { delete o.ph; others.push(o); }
    s.els = [...els, ...others];
    s.layout = layoutId;
    commit();
  }
  function applyTheme(id) {
    const old = theme();
    pres.themeId = id;
    const th = theme();
    for (const s of pres.slides) {
      s.bg = th.bg;
      for (const e of s.els) {
        if (e.ph) {
          e.color = e.ph === 'title' ? th.title : th.text;
          e.font = e.ph === 'title' ? th.fontT : th.fontB;
        } else if (e.type === 'text' && (e.color === old.text || e.color === old.title)) {
          e.color = e.color === old.title ? th.title : th.text;
          if (e.font === old.fontB) e.font = th.fontB;
          if (e.font === old.fontT) e.font = th.fontT;
        } else if (e.type === 'shape' && e.fill === old.accent) {
          e.fill = th.accent;
          if (e.line) e.line.color = mix(th.accent, '#000000', 0.3);
        } else if (e.type === 'shape' && e.shape === 'line' && e.line && e.line.color === old.accent) {
          e.line.color = th.accent;
        } else if (e.type === 'table') e.accent = th.accent;
      }
    }
    commit();
  }

  // Miniaturen: Klick, Ziehen, Kontextmenue
  const list = $('#slideList');
  list.addEventListener('mousedown', e => {
    const row = e.target.closest('.thumb-row');
    if (row && e.button === 0) goTo(+row.dataset.i);
  });
  let dragFrom = null;
  list.addEventListener('dragstart', e => { const row = e.target.closest('.thumb-row'); if (row) { dragFrom = +row.dataset.i; e.dataTransfer.effectAllowed = 'move'; } });
  list.addEventListener('dragover', e => {
    if (dragFrom == null) return;
    e.preventDefault();
    $$('.thumb-row', list).forEach(r => r.classList.remove('drop-before'));
    const row = e.target.closest('.thumb-row');
    if (row) row.classList.add('drop-before');
  });
  list.addEventListener('drop', e => {
    e.preventDefault();
    $$('.thumb-row', list).forEach(r => r.classList.remove('drop-before'));
    const row = e.target.closest('.thumb-row');
    if (dragFrom == null) return;
    let to = row ? +row.dataset.i : pres.slides.length;
    const [s] = pres.slides.splice(dragFrom, 1);
    if (to > dragFrom) to--;
    pres.slides.splice(to, 0, s);
    cur = to;
    dragFrom = null;
    thumbCache.clear();
    commit();
  });
  list.addEventListener('dragend', () => { dragFrom = null; $$('.thumb-row', list).forEach(r => r.classList.remove('drop-before')); });
  list.addEventListener('contextmenu', e => {
    e.preventDefault();
    const row = e.target.closest('.thumb-row');
    if (row) goTo(+row.dataset.i);
    const fake = { getBoundingClientRect: () => ({ left: e.clientX, right: e.clientX, top: e.clientY, bottom: e.clientY }) };
    Rex.menu(fake, [
      { label: 'Neue Folie', icon: 'file-plus-2', action: () => newSlide() },
      { label: 'Folie duplizieren', icon: 'copy-plus', action: () => duplicateSlide() },
      { label: 'Folie löschen', icon: 'trash-2', action: () => deleteSlide() },
      '-',
      { label: 'Nach oben', icon: 'arrow-up', action: () => { if (cur > 0) { const [s] = pres.slides.splice(cur, 1); pres.slides.splice(cur - 1, 0, s); cur--; thumbCache.clear(); commit(); } } },
      { label: 'Nach unten', icon: 'arrow-down', action: () => { if (cur < pres.slides.length - 1) { const [s] = pres.slides.splice(cur, 1); pres.slides.splice(cur + 1, 0, s); cur++; thumbCache.clear(); commit(); } } }
    ]);
  });

  $('#notes').addEventListener('input', e => { slide().notes = e.target.value; Rex.setDirty(true); });
  $('#notes').addEventListener('change', () => commit({ noRender: true }));

  // =====================================================================
  //  Tastatur
  // =====================================================================
  document.addEventListener('keydown', e => {
    if (!$('#show').hidden) return;
    const t = e.target;
    const inField = t.closest && (t.closest('input,select,textarea,.modal') || t.isContentEditable);
    const ctrl = e.ctrlKey || e.metaKey;
    const k = e.key;
    if (k === 'F5') { e.preventDefault(); startShow(e.shiftKey ? cur : 0); return; }
    if (editingId) {
      if (k === 'Escape') { e.preventDefault(); const id = editingId; stopEditing(); select([id]); }
      if (ctrl && k.toLowerCase() === 'z') { e.preventDefault(); stopEditing(); undo(); }
      if (ctrl && k.toLowerCase() === 'y') { e.preventDefault(); stopEditing(); redo(); }
      if (k === 'Tab' && elById(editingId) && elById(editingId).type !== 'table') { e.preventDefault(); document.execCommand(e.shiftKey ? 'outdent' : 'indent'); }
      return;
    }
    if (inField) return;
    if (ctrl) {
      const kl = k.toLowerCase();
      if (kl === 'z') { e.preventDefault(); undo(); return; }
      if (kl === 'y') { e.preventDefault(); redo(); return; }
      if (kl === 'm') { e.preventDefault(); newSlide(); return; }
      if (kl === 'd') { e.preventDefault(); duplicateEls(); return; }
      if (kl === 'a') { e.preventDefault(); select(slide().els.map(x => x.id)); return; }
      if (kl === 'c' && selected.length) { e.preventDefault(); copyEls(); return; }
      if (kl === 'x' && selected.length) { e.preventDefault(); copyEls(); deleteSelected(); return; }
      if (kl === 'b' && selected.length) { e.preventDefault(); tcmd('bold'); return; }
      if (kl === 'i' && selected.length) { e.preventDefault(); tcmd('italic'); return; }
      if (kl === 'u' && selected.length) { e.preventDefault(); tcmd('underline'); return; }
      return;
    }
    if (selected.length) {
      if (k === 'Delete' || k === 'Backspace') { e.preventDefault(); deleteSelected(); return; }
      const arrows = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] };
      if (arrows[k]) {
        e.preventDefault();
        const step = e.shiftKey ? 10 : 1;
        for (const id of selected) { const el = elById(id); el.x += arrows[k][0] * step; el.y += arrows[k][1] * step; }
        commit();
        return;
      }
      if (k === 'Escape') { select([]); return; }
      if (k === 'Enter' || k === 'F2') { e.preventDefault(); startEditing(selected[0]); return; }
      // Tippen beginnt die Bearbeitung des ausgewaehlten Textes
      if (k.length === 1 && selected.length === 1) {
        const el = elById(selected[0]);
        if (el && (el.type === 'text' || el.type === 'shape')) startEditing(el.id, { selectAll: true });
      }
      return;
    }
    if (k === 'PageDown' || k === 'ArrowDown' || k === 'ArrowRight') { e.preventDefault(); goTo(cur + 1); }
    if (k === 'PageUp' || k === 'ArrowUp' || k === 'ArrowLeft') { e.preventDefault(); goTo(cur - 1); }
    if (k === 'Home') goTo(0);
    if (k === 'End') goTo(pres.slides.length - 1);
  });
  document.addEventListener('paste', async e => {
    if (editingId || !$('#show').hidden) return;
    if (e.target.closest && e.target.closest('input,textarea,.modal')) return;
    const file = Array.from(e.clipboardData.files || []).find(f => f.type.startsWith('image/'));
    if (file) {
      e.preventDefault();
      const r = new FileReader();
      r.onload = () => addImageFromSrc(r.result);
      r.readAsDataURL(file);
      return;
    }
    const text = e.clipboardData.getData('text/plain');
    if (elClip && text === elClip.map(x => stripTags(x.html || '')).join('\n')) { e.preventDefault(); pasteEls(); return; }
    if (text) {
      e.preventDefault();
      const th = theme();
      addEl({ type: 'text', x: 80, y: 80, w: pres.w - 160, h: 60, html: text.split(/\r?\n/).map(l => `<p>${esc(l) || '<br>'}</p>`).join(''), font: th.fontB, size: 18, color: th.text, align: 'left', va: 'top' });
    }
  });
  stage.addEventListener('dragover', e => e.preventDefault());
  stage.addEventListener('drop', async e => {
    const f = e.dataTransfer.files[0];
    if (!f) return;
    e.preventDefault();
    if (f.type.startsWith('image/')) {
      const r = new FileReader();
      r.onload = () => addImageFromSrc(r.result);
      r.readAsDataURL(f);
    } else {
      Rex.loadFile({ name: f.name, path: Rex.util.filePath(f), data: new Uint8Array(await f.arrayBuffer()) });
    }
  });

  // =====================================================================
  //  Menueband
  // =====================================================================
  function updateRibbon() {
    const el = elById(editingId || selected[0]);
    const ff = $('#fontFamily'), fs = $('#fontSize');
    if (el && el.font) { if (![...ff.options].some(o => o.value === el.font)) ff.add(new Option(el.font, el.font)); ff.value = el.font; }
    if (el && el.size) { const v = String(el.size); if (![...fs.options].some(o => o.value === v)) fs.add(new Option(v, v)); fs.value = v; }
    $$('#themeGallery button').forEach(b => b.classList.toggle('on', b.dataset.theme === pres.themeId));
    $$('#transGallery .btn').forEach(b => b.classList.toggle('on', b.dataset.trans === (slide().transition || 'none')));
    $('#notesBtn').classList.toggle('on', !$('#notesWrap').classList.contains('hidden'));
    $('#guidesBtn').classList.toggle('on', guidesOn);
  }
  const commands = {
    newSlide: () => newSlide(),
    duplicateSlide, deleteSlide,
    insertText, insertTable,
    insertImage: async () => addImageFromSrc(await Rex.util.readImageFile()),
    insertWordArt: () => {
      const th = theme();
      const el = addEl({ type: 'text', x: pres.w / 2 - 300, y: pres.h / 2 - 60, w: 600, h: 120, html: '<p>Dein Text hier</p>', font: 'Impact', size: 60, color: th.accent, align: 'center', va: 'middle', bold: false, shadow: true });
      startEditing(el.id, { selectAll: true });
    },
    slideNumber: () => { pres.showNumbers = !pres.showNumbers; thumbCache.clear(); commit(); Rex.toast(pres.showNumbers ? 'Foliennummern an' : 'Foliennummern aus'); },
    growFont: () => stepFont(1),
    shrinkFont: () => stepFont(-1),
    bgImage: async () => {
      const src = await Rex.util.readImageFile();
      if (!src) return;
      const id = uid('m');
      pres.media[id] = src;
      const all = await Rex.confirm('Hintergrundbild', 'Auf alle Folien anwenden?', 'Alle Folien', 'Nur diese Folie');
      for (const s of all ? pres.slides : [slide()]) s.bgImage = id;
      commit();
    },
    transAll: () => { const t = slide().transition || 'none'; pres.slides.forEach(s => { s.transition = t; }); commit(); Rex.toast('Übergang auf alle Folien angewendet'); },
    showStart: () => startShow(0),
    showCurrent: () => startShow(cur),
    zoomIn: () => { autoFit = false; zoom = Math.min(3, zoom + 0.1); applyZoom(); renderSelection(); },
    zoomOut: () => { autoFit = false; zoom = Math.max(0.15, zoom - 0.1); applyZoom(); renderSelection(); },
    zoomFit: () => { autoFit = true; applyZoom(); renderSelection(); },
    toggleNotes: () => { $('#notesWrap').classList.toggle('hidden'); applyZoom(); renderSelection(); updateRibbon(); },
    toggleGuides: () => { guidesOn = !guidesOn; updateRibbon(); }
  };
  document.addEventListener('mousedown', e => {
    if (e.target.closest('.ribbon .btn, .titlebar button, .statusbar .btn')) e.preventDefault();
  });
  document.addEventListener('click', e => {
    const c = e.target.closest('[data-cmd]');
    if (c && commands[c.dataset.cmd]) { commands[c.dataset.cmd](); return; }
    const t = e.target.closest('[data-tcmd]');
    if (t) tcmd(t.dataset.tcmd);
  });

  const ff = $('#fontFamily');
  FONTS.forEach(f => { const o = new Option(f, f); o.style.fontFamily = f; ff.add(o); });
  let savedRange = null;
  const saveRange = () => { const s = getSelection(); if (editingId && s.rangeCount) savedRange = s.getRangeAt(0).cloneRange(); };
  const restoreRange = () => {
    if (!editingId || !savedRange) return;
    const node = canvas.querySelector(`.el[data-id="${editingId}"] .txt-inner, .el[data-id="${editingId}"] table`);
    if (node) node.focus();
    const s = getSelection(); s.removeAllRanges(); s.addRange(savedRange);
  };
  ff.addEventListener('mousedown', saveRange);
  ff.addEventListener('change', () => {
    const v = ff.value;
    if (editingId) { restoreRange(); document.execCommand('styleWithCSS', false, true); document.execCommand('fontName', false, v); }
    else { for (const id of selected) { const el = elById(id); if (el && el.type !== 'image') { el.font = v; if (el.html) el.html = el.html.replace(/font-family:[^;"]+;?/g, ''); } } commit(); }
  });
  const fs = $('#fontSize');
  SIZES.forEach(s => fs.add(new Option(s, s)));
  fs.addEventListener('mousedown', saveRange);
  fs.addEventListener('change', () => { restoreRange(); setFontSize(+fs.value); });

  $('#textColorBtn').addEventListener('click', e => Rex.colorMenu(e.currentTarget, c => {
    if (!c) c = '#000000';
    $('#textColorBar').style.background = c;
    if (!editingId) {
      for (const id of selected) { const el = elById(id); if (el && el.type !== 'image') { el.color = c; if (el.html) el.html = el.html.replace(/(?<![-\w])color:[^;"]+;?/g, ''); } }
      commit();
    } else { document.execCommand('styleWithCSS', false, true); document.execCommand('foreColor', false, c); }
  }, { noneLabel: 'Automatisch' }));
  $('#fillBtn').addEventListener('click', e => Rex.colorMenu(e.currentTarget, c => {
    if (c) $('#fillBar').style.background = c;
    for (const id of selected) { const el = elById(id); if (el && (el.type === 'shape' || el.type === 'text')) el.fill = c; if (el && el.type === 'table') el.accent = c || theme().accent; }
    commit();
  }, { noneLabel: 'Keine Füllung' }));
  $('#lineBtn').addEventListener('click', e => Rex.menu(e.currentTarget, m => {
    const mk = (label, fn) => { const b = document.createElement('button'); b.textContent = label; b.onclick = () => { Rex.closeMenu(); fn(); }; m.appendChild(b); };
    const setLine = (fn) => { for (const id of selected) { const el = elById(id); if (el) { el.line = el.line || { color: '#000000', width: 1 }; fn(el.line, el); if (!el.line.width) el.line = null; } } commit(); };
    mk('Keine Kontur', () => { for (const id of selected) { const el = elById(id); if (el) el.line = null; } commit(); });
    for (const w of [0.75, 1.5, 2.25, 3, 4.5, 6]) mk(`Stärke ${String(w).replace('.', ',')} pt`, () => setLine(l => { l.width = w; }));
    m.appendChild(document.createElement('hr'));
    const g = document.createElement('div');
    g.className = 'swatches';
    for (const c of Rex.PALETTE) {
      const b = document.createElement('button');
      b.style.background = c;
      b.onclick = () => { Rex.closeMenu(); $('#lineBar').style.background = c; setLine(l => { l.color = c; if (!l.width) l.width = 1; }); };
      g.appendChild(b);
    }
    m.appendChild(g);
  }));
  $('#arrangeBtn').addEventListener('click', e => Rex.menu(e.currentTarget, [
    { label: 'In den Vordergrund', icon: 'bring-to-front', action: () => arrange('front') },
    { label: 'Eine Ebene nach vorne', icon: 'arrow-up', action: () => arrange('forward') },
    { label: 'Eine Ebene nach hinten', icon: 'arrow-down', action: () => arrange('backward') },
    { label: 'In den Hintergrund', icon: 'send-to-back', action: () => arrange('back') },
    '-',
    { label: 'Links ausrichten', icon: 'align-start-vertical', action: () => align('left') },
    { label: 'Horizontal zentrieren', icon: 'align-center-vertical', action: () => align('center') },
    { label: 'Rechts ausrichten', icon: 'align-end-vertical', action: () => align('right') },
    { label: 'Oben ausrichten', icon: 'align-start-horizontal', action: () => align('top') },
    { label: 'Vertikal zentrieren', icon: 'align-center-horizontal', action: () => align('middle') },
    { label: 'Unten ausrichten', icon: 'align-end-horizontal', action: () => align('bottom') },
    '-',
    { label: 'Horizontal spiegeln', icon: 'flip-horizontal', action: () => { for (const id of selected) { const el = elById(id); if (el) el.flipH = !el.flipH; } commit(); } },
    { label: 'Vertikal spiegeln', icon: 'flip-vertical', action: () => { for (const id of selected) { const el = elById(id); if (el) el.flipV = !el.flipV; } commit(); } },
    { label: 'Drehung zurücksetzen', icon: 'rotate-ccw', action: () => { for (const id of selected) { const el = elById(id); if (el) el.rot = 0; } commit(); } }
  ]));
  $('#valignBtn').addEventListener('click', e => Rex.menu(e.currentTarget, [['top', 'Oben', 'align-vertical-justify-start'], ['middle', 'Mitte', 'align-vertical-justify-center'], ['bottom', 'Unten', 'align-vertical-justify-end']].map(([v, l, i]) => ({
    label: l, icon: i, action: () => { if (editingId) stopEditing(); for (const id of selected) { const el = elById(id); if (el) el.va = v; } commit(); }
  }))));
  const shapeMenu = (e) => Rex.menu(e.currentTarget, SHAPE_LIST.map(([id, label, icon]) => ({ label, icon, action: () => insertShape(id) })));
  $('#shapeBtn').addEventListener('click', shapeMenu);
  $('#shapeBtn2').addEventListener('click', shapeMenu);
  $('#chartBtn').addEventListener('click', e => Rex.menu(e.currentTarget, [
    { label: 'Säulendiagramm', icon: 'chart-column', action: () => insertChart('column') },
    { label: 'Balkendiagramm', icon: 'chart-bar', action: () => insertChart('bar') },
    { label: 'Liniendiagramm', icon: 'chart-line', action: () => insertChart('line') },
    { label: 'Kreisdiagramm', icon: 'chart-pie', action: () => insertChart('pie') }
  ]));
  $('#layoutBtn').addEventListener('click', e => Rex.menu(e.currentTarget, Object.entries(LAYOUTS).map(([id, L]) => ({ label: L.name, icon: 'layout-template', action: () => applyLayout(id) }))));
  $('#sizeBtn').addEventListener('click', e => Rex.menu(e.currentTarget, [
    { label: 'Breitbild (16:9)', icon: 'rectangle-horizontal', action: () => resizePres(960, 540) },
    { label: 'Standard (4:3)', icon: 'square', action: () => resizePres(960, 720) },
    { label: 'Hochformat (A4)', icon: 'rectangle-vertical', action: () => resizePres(720, 1018) }
  ]));
  function resizePres(w, h) {
    const sx = w / pres.w, sy = h / pres.h;
    for (const s of pres.slides) for (const e of s.els) { e.x *= sx; e.w *= sx; e.y *= sy; e.h *= sy; }
    pres.w = w; pres.h = h;
    thumbCache.clear();
    commit();
  }
  $('#bgBtn').addEventListener('click', e => Rex.colorMenu(e.currentTarget, c => {
    Rex.confirm('Hintergrund', 'Auf alle Folien anwenden?', 'Alle Folien', 'Nur diese Folie').then(all => {
      for (const s of all ? pres.slides : [slide()]) { s.bg = c || '#ffffff'; s.bgImage = null; }
      commit();
    });
  }, { noneLabel: 'Weiß / Bild entfernen' }));
  $('#symbolBtn').addEventListener('click', e => {
    saveRange();
    Rex.menu(e.currentTarget, m => {
      const g = document.createElement('div');
      g.style.cssText = 'display:grid;grid-template-columns:repeat(10,28px);gap:2px;padding:4px';
      for (const s of '€ © ® ™ § ° ± × ÷ ≠ ≈ ≤ ≥ ∞ √ π → ← ↑ ↓ ⇒ • … – — „ “ ✓ ✗ ★ ☆ ♥ ☺ ✉ ♪ ⚠ ☀ ☁ ⚡ ✈ ⚽ 🎮 🚀 💡 🔥 ✅ ❌ 👍 😀 🎉'.split(' ')) {
        const b = document.createElement('button');
        b.textContent = s;
        b.style.cssText = 'width:28px;height:28px;padding:0;justify-content:center;font-size:15px';
        b.onclick = () => {
          Rex.closeMenu();
          if (editingId) { restoreRange(); document.execCommand('insertText', false, s); }
          else { const th = theme(); addEl({ type: 'text', x: pres.w / 2 - 40, y: pres.h / 2 - 40, w: 80, h: 80, html: `<p>${s}</p>`, font: th.fontB, size: 48, color: th.text, align: 'center', va: 'middle' }); }
        };
        g.appendChild(b);
      }
      m.appendChild(g);
    });
  });
  // Designs-Galerie
  for (const th of THEMES) {
    const b = document.createElement('button');
    b.dataset.theme = th.id;
    b.title = th.name;
    b.style.background = th.bg;
    b.innerHTML = `<span style="color:${th.title};font-family:'${th.fontT}'">Aa</span><i style="background:${th.accent}"></i>`;
    b.onclick = () => applyTheme(th.id);
    $('#themeGallery').appendChild(b);
  }
  for (const [id, label, icon] of TRANSITIONS) {
    const b = document.createElement('button');
    b.className = 'btn big';
    b.dataset.trans = id;
    b.innerHTML = `<i data-lucide="${icon}"></i>${label}`;
    b.onclick = () => { slide().transition = id; commit(); if (id !== 'none') previewTransition(id); };
    $('#transGallery').appendChild(b);
  }
  function previewTransition(t) {
    const node = canvas.querySelector('.slide');
    if (!node) return;
    node.classList.add('tr-layer', t);
    setTimeout(() => node.classList.remove('tr-layer', t), 700);
  }
  $('#zoomRange').addEventListener('input', e => { autoFit = false; zoom = +e.target.value / 100; applyZoom(); renderSelection(); });
  new ResizeObserver(() => { if (autoFit) { applyZoom(); renderSelection(); } }).observe(stage);

  // =====================================================================
  //  Bildschirmpraesentation
  // =====================================================================
  let showIdx = 0, showEnded = false, cursorTimer = null;
  const showEl = $('#show'), showSlide = $('#showSlide');
  function showScale() { return Math.min(innerWidth / pres.w, innerHeight / pres.h); }
  function renderShow(dir = 1, animate = true) {
    const sc = showScale();
    showSlide.style.width = pres.w * sc + 'px';
    showSlide.style.height = pres.h * sc + 'px';
    if (showEnded) {
      showSlide.innerHTML = '<div class="tr-layer fade" style="background:#000;display:grid;place-items:center"><div class="show-end">Ende der Bildschirmpräsentation. Zum Beenden klicken.</div></div>';
      $('#showNum').textContent = '';
      return;
    }
    const sl = pres.slides[showIdx];
    const layer = document.createElement('div');
    layer.className = 'tr-layer';
    if (animate && sl.transition && sl.transition !== 'none') layer.classList.add(sl.transition, ...(dir < 0 && sl.transition === 'push' ? ['back'] : []));
    layer.innerHTML = slideHTML(sl, 'show', showIdx);
    layer.firstElementChild.style.transform = `scale(${sc})`;
    layer.firstElementChild.style.transformOrigin = '0 0';
    // alte Ebene nach dem Uebergang entfernen
    const old = Array.from(showSlide.children);
    showSlide.appendChild(layer);
    setTimeout(() => old.forEach(o => o.remove()), animate ? 700 : 0);
    $('#showNum').textContent = `${showIdx + 1} / ${pres.slides.length}`;
  }
  function startShow(i) {
    if (editingId) stopEditing();
    showIdx = i; showEnded = false;
    showEl.hidden = false;
    showSlide.innerHTML = '';
    if (document.documentElement.requestFullscreen) document.documentElement.requestFullscreen().catch(() => {});
    setTimeout(() => renderShow(1, true), 60);
  }
  function endShow() {
    showEl.hidden = true;
    showSlide.innerHTML = '';
    if (document.fullscreenElement) document.exitFullscreen().catch(() => {});
    goTo(Math.min(showIdx, pres.slides.length - 1));
  }
  function showNext() {
    if (showEnded) { endShow(); return; }
    if (showIdx < pres.slides.length - 1) { showIdx++; renderShow(1); }
    else { showEnded = true; renderShow(1); }
  }
  function showPrev() {
    if (showEnded) { showEnded = false; renderShow(-1, false); return; }
    if (showIdx > 0) { showIdx--; renderShow(-1); }
  }
  showEl.addEventListener('click', e => {
    const b = e.target.closest('[data-show]');
    if (b) { const a = b.dataset.show; if (a === 'next') showNext(); if (a === 'prev') showPrev(); if (a === 'end') endShow(); return; }
    showNext();
  });
  showEl.addEventListener('contextmenu', e => { e.preventDefault(); showPrev(); });
  showEl.addEventListener('mousemove', () => {
    showEl.classList.add('cursor');
    clearTimeout(cursorTimer);
    cursorTimer = setTimeout(() => showEl.classList.remove('cursor'), 2000);
  });
  document.addEventListener('keydown', e => {
    if (showEl.hidden) return;
    e.preventDefault();
    const k = e.key;
    if (['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Enter', 'n', 'N'].includes(k)) showNext();
    else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace', 'p', 'P'].includes(k)) showPrev();
    else if (k === 'Escape') endShow();
    else if (k === 'Home') { showIdx = 0; showEnded = false; renderShow(-1); }
    else if (k === 'End') { showIdx = pres.slides.length - 1; showEnded = false; renderShow(1); }
    else if (k === 'b' || k === 'B' || k === '.') showSlide.style.visibility = showSlide.style.visibility === 'hidden' ? '' : 'hidden';
  }, true);
  document.addEventListener('fullscreenchange', () => { if (!document.fullscreenElement && !showEl.hidden) endShow(); });
  window.addEventListener('resize', () => { if (!showEl.hidden) renderShow(1, false); });

  // =====================================================================
  //  Export-Hilfen: Absaetze aus dem DOM berechnen
  // =====================================================================
  function cssHex(c) {
    const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?/.exec(c || '');
    if (!m || (m[4] !== undefined && +m[4] === 0)) return null;
    return '#' + [m[1], m[2], m[3]].map(x => (+x).toString(16).padStart(2, '0')).join('');
  }
  function extractParas(inner) {
    const paras = [];
    let curP = null;
    const blockAlign = (el) => {
      if (!(el instanceof Element)) el = el && el.alignEl instanceof Element ? el.alignEl : inner;
      const a = getComputedStyle(el).textAlign;
      return a === 'start' ? 'left' : a === 'end' ? 'right' : a;
    };
    const runOf = (textNode) => {
      const p = textNode.parentElement;
      const cs = getComputedStyle(p);
      const r = {
        text: textNode.nodeValue.replace(/[\n\r\t]+/g, ' '),
        bold: parseInt(cs.fontWeight, 10) >= 600, italic: cs.fontStyle === 'italic',
        color: cssHex(cs.color), size: Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2,
        font: cs.fontFamily.split(',')[0].replace(/["']/g, '').trim()
      };
      for (let n = p; n && n !== inner.parentElement; n = n.parentElement) {
        const ns = getComputedStyle(n);
        if ((ns.textDecorationLine || '').includes('underline')) r.underline = true;
        if ((ns.textDecorationLine || '').includes('line-through')) r.strike = true;
        if (!r.highlight && n !== inner && !n.classList.contains('txt')) { const bg = cssHex(ns.backgroundColor); if (bg) r.highlight = bg; }
        if (n === inner) break;
      }
      return r;
    };
    const inline = (node) => {
      for (const ch of node.childNodes) {
        if (ch.nodeType === 3) { if (ch.nodeValue) { if (!curP) newP(blockAlign(node), null, 0); curP.runs.push(runOf(ch)); } }
        else if (ch.nodeType === 1) {
          if (ch.tagName === 'BR') { if (!curP) newP(blockAlign(node), null, 0); curP.runs.push({ text: '', br: true }); }
          else if (/^(P|DIV|H[1-6]|UL|OL|LI|BLOCKQUOTE)$/.test(ch.tagName)) block(ch, 0);
          else inline(ch);
        }
      }
    };
    const newP = (align, bullet, level) => { curP = { align, bullet, level, runs: [] }; paras.push(curP); };
    function block(node, level) {
      if (node.tagName === 'UL' || node.tagName === 'OL') {
        for (const li of node.children) {
          if (li.tagName !== 'LI') { block(li, level); continue; }
          const hidden = li.style.listStyle === 'none' || li.style.listStyleType === 'none';
          newP(blockAlign(li), hidden ? null : (node.tagName === 'OL' ? 'number' : 'bullet'), level);
          for (const ch of li.childNodes) {
            if (ch.nodeType === 1 && (ch.tagName === 'UL' || ch.tagName === 'OL')) block(ch, level + 1);
            else if (ch.nodeType === 1 && /^(P|DIV)$/.test(ch.tagName)) inline(ch);
            else inline({ childNodes: [ch], alignEl: li });
          }
          curP = null;
        }
        return;
      }
      newP(blockAlign(node), null, level);
      inline(node);
      curP = null;
    }
    for (const ch of inner.childNodes) {
      if (ch.nodeType === 1 && /^(P|DIV|H[1-6]|UL|OL|BLOCKQUOTE)$/.test(ch.tagName)) { curP = null; block(ch, 0); }
      else inline({ childNodes: [ch], alignEl: inner });
    }
    // Trailing <br> in Absaetzen entfernen
    for (const p of paras) while (p.runs.length > 1 && p.runs[p.runs.length - 1].br) p.runs.pop();
    for (const p of paras) if (p.runs.length === 1 && p.runs[0].br) p.runs = [];
    return paras;
  }
  async function paragraphsForSlide(si) {
    const host = document.createElement('div');
    host.style.cssText = 'position:fixed;left:-20000px;top:0;visibility:hidden';
    host.innerHTML = slideHTML(pres.slides[si], 'export', si);
    document.body.appendChild(host);
    const map = new Map();
    for (const el of pres.slides[si].els) {
      const node = host.querySelector(`.el[data-id="${el.id}"] .txt-inner`);
      if (node) map.set(el.id, extractParas(node));
    }
    host.remove();
    return map;
  }

  // =====================================================================
  //  Laden / Speichern / Drucken
  // =====================================================================
  function setPresentation(p) {
    pres = p;
    pres.media = pres.media || {};
    cur = 0;
    selected = [];
    editingId = null;
    thumbCache.clear();
    $('#slideList').innerHTML = '';
    autoFit = true;
    resetHistory();
    renderAll();
  }
  async function load(bytes, ext) {
    if (ext !== 'pptx') throw new Error('Nur PowerPoint-Dateien (.pptx) werden unterstützt.');
    setPresentation(await RexPptx.importPptx(bytes));
  }
  async function save(fmt) {
    if (editingId) stopEditing();
    if (fmt !== 'pptx') throw new Error('Unbekanntes Format');
    return RexPptx.exportPptx(pres, {
      title: Rex.doc.name,
      paragraphs: paragraphsForSlide,
      cellStyle: tableCellStyle,
      plain: (h) => stripTags(h).trim()
    });
  }
  let printStyle = null;
  function buildPrint() {
    if (editingId) stopEditing();
    $('#printArea').innerHTML = pres.slides.map((s, i) => `<div class="print-slide" style="width:${pres.w}px;height:${pres.h}px">${slideHTML(s, 'show', i)}</div>`).join('');
    if (!printStyle) { printStyle = document.createElement('style'); document.head.appendChild(printStyle); }
    printStyle.textContent = `@media print { @page { size: ${pres.w / 96}in ${pres.h / 96}in; margin: 0; } }`;
  }

  // Vorlagen
  function tpl(themeId, slidesSpec) {
    newPresentation();
    pres.themeId = themeId;
    pres.slides = [];
    for (const [layout, title, body] of slidesSpec) {
      const s = makeSlide(layout);
      for (const e of s.els) {
        if (e.ph === 'title') e.html = `<p>${esc(title)}</p>`;
        else if (e.ph === 'subtitle') e.html = `<p>${esc(body || '')}</p>`;
        else if (e.ph === 'body' && body) e.html = '<ul>' + body.split('\n').map(l => `<li>${esc(l)}</li>`).join('') + '</ul>';
      }
      pres.slides.push(s);
    }
    setPresentation(pres);
  }
  const TEMPLATES = [
    { name: 'Leere Präsentation', icon: 'presentation', docName: 'Präsentation', create: () => { newPresentation(); setPresentation(pres); } },
    { name: 'Referat', icon: 'graduation-cap', docName: 'Referat', create: () => tpl('slate', [
      ['title', 'Mein Referat', 'Name · Klasse · Datum'], ['titleContent', 'Gliederung', 'Einleitung\nHauptteil\nFazit\nQuellen'],
      ['titleContent', 'Einleitung', 'Worum geht es?\nWarum ist das Thema wichtig?'], ['section', 'Hauptteil', 'Die wichtigsten Fakten'],
      ['twoContent', 'Vorteile und Nachteile', 'Vorteil 1\nVorteil 2'], ['titleContent', 'Fazit', 'Zusammenfassung\nEigene Meinung'],
      ['titleContent', 'Quellen', 'Buch / Webseite 1\nBuch / Webseite 2'], ['title', 'Danke fürs Zuhören!', 'Fragen?']]) },
    { name: 'Gaming', icon: 'gamepad-2', docName: 'Gaming', create: () => tpl('gaming', [
      ['title', 'REX Client', 'Der beste Minecraft-Launcher'], ['titleContent', 'Features', 'Schneller Start\nMods mit einem Klick\nAutomatische Updates'],
      ['section', 'Live-Demo', 'Los geht\'s!'], ['title', 'GG!', 'Danke fürs Zuschauen']]) },
    { name: 'Geschäftlich', icon: 'briefcase', docName: 'Präsentation', create: () => tpl('ocean', [
      ['title', 'Quartalsbericht', 'Q3 · Team Vertrieb'], ['titleContent', 'Agenda', 'Ergebnisse\nZiele\nNächste Schritte'],
      ['titleOnly', 'Ergebnisse', ''], ['titleContent', 'Nächste Schritte', 'Kampagne starten\nTeam erweitern\nKunden befragen']]) }
  ];

  newPresentation();
  Rex.init({
    app: 'present',
    letter: 'P',
    icon: 'presentation',
    defaultName: 'Präsentation',
    openFormats: [{ name: 'PowerPoint-Präsentation', extensions: ['pptx'] }],
    saveFormats: [{ name: 'PowerPoint-Präsentation', extensions: ['pptx'], hint: 'Kompatibel mit Microsoft PowerPoint, LibreOffice Impress und Google Präsentationen' }],
    templates: TEMPLATES,
    load, save, undo, redo,
    shortcutsHelp: [['F5 / Umschalt+F5', 'Präsentation starten (Anfang / aktuelle Folie)'], ['Strg+M', 'Neue Folie'], ['Strg+D', 'Duplizieren'],
      ['Pfeiltasten', 'Objekt verschieben (Umschalt = 10 px)'], ['Alt beim Ziehen', 'Ohne Ausrichtungshilfe'], ['B', 'Schwarzer Bildschirm (in der Präsentation)']],
    pdf: { pageSize: () => ({ width: pres.w / 96, height: pres.h / 96 }), landscape: false, margins: () => ({ top: 0, bottom: 0, left: 0, right: 0 }), before: buildPrint, after: () => { $('#printArea').innerHTML = ''; } },
    onTheme: () => renderAll()
  });
  setPresentation(pres);
  Rex.icons();

  window.RexPresent = { get pres() { return pres; }, save, load, goTo, startShow, insertShape, insertText, addImageFromSrc, applyTheme, select: (ids) => select(ids), stopEditing, startEditing, newSlide };
})();
