/* REX Text – Textverarbeitung */
(function () {
  'use strict';
  const { $, $$, esc } = Rex.util;
  const editor = $('#editor');
  const page = $('#page');
  const workspace = $('#workspace');

  const MM = 96 / 25.4;
  const FONTS = ['Calibri', 'Calibri Light', 'Arial', 'Aptos', 'Cambria', 'Candara', 'Century Gothic', 'Comic Sans MS', 'Consolas',
    'Courier New', 'Franklin Gothic Medium', 'Garamond', 'Georgia', 'Impact', 'Lucida Sans', 'Palatino Linotype', 'Segoe UI',
    'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana'];
  const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
  const PAGE_SIZES = { A4: [210, 297], A5: [148, 210], Letter: [215.9, 279.4], Legal: [215.9, 355.6] };
  const STYLES = [
    { id: 'p', label: 'Standard', tag: 'p', css: 'font-size:12px' },
    { id: 'title', label: 'Titel', tag: 'p', cls: 'title', css: 'font-size:15px;font-family:Calibri Light' },
    { id: 'subtitle', label: 'Untertitel', tag: 'p', cls: 'subtitle', css: 'color:#5a5a5a' },
    { id: 'h1', label: 'Überschrift 1', tag: 'h1', css: 'color:#2f5496;font-size:13px' },
    { id: 'h2', label: 'Überschrift 2', tag: 'h2', css: 'color:#2f5496;font-size:12.5px' },
    { id: 'h3', label: 'Überschrift 3', tag: 'h3', css: 'color:#1f3763;font-size:12px' },
    { id: 'blockquote', label: 'Zitat', tag: 'blockquote', css: 'font-style:italic;color:#404040' },
    { id: 'pre', label: 'Code', tag: 'pre', css: 'font-family:Consolas;font-size:10pt' }
  ];

  let settings = defaultSettings();
  let zoom = 1;

  function defaultSettings() {
    return { size: 'A4', w: 210, h: 297, orientation: 'portrait', margins: { top: 25, right: 25, bottom: 25, left: 25 }, pageNumbers: false };
  }

  // ---------------------------------------------------------------
  // Verlauf (Rueckgaengig / Wiederholen)
  // ---------------------------------------------------------------
  const hist = { stack: [], index: -1, timer: null };
  function caretOffsets() {
    const sel = getSelection();
    if (!sel.rangeCount || !editor.contains(sel.anchorNode)) return null;
    const r = sel.getRangeAt(0);
    const pre = document.createRange();
    pre.selectNodeContents(editor);
    pre.setEnd(r.startContainer, r.startOffset);
    const start = pre.toString().length;
    return { start, end: start + r.toString().length };
  }
  function restoreOffsets(o) {
    if (!o) return;
    const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
    let pos = 0, n, startSet = false;
    const range = document.createRange();
    while ((n = walker.nextNode())) {
      const len = n.nodeValue.length;
      if (!startSet && o.start <= pos + len) { range.setStart(n, o.start - pos); startSet = true; }
      if (startSet && o.end <= pos + len) { range.setEnd(n, o.end - pos); break; }
      pos += len;
    }
    if (!startSet) { range.selectNodeContents(editor); range.collapse(false); }
    const sel = getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  }
  function pushHistory() {
    clearTimeout(hist.timer);
    hist.timer = null;
    const html = editor.innerHTML;
    const cur = hist.stack[hist.index];
    if (cur && cur.html === html) { cur.sel = caretOffsets(); return; }
    hist.stack = hist.stack.slice(0, hist.index + 1);
    hist.stack.push({ html, sel: caretOffsets() });
    if (hist.stack.length > 300) hist.stack.shift();
    hist.index = hist.stack.length - 1;
  }
  function scheduleHistory() { clearTimeout(hist.timer); hist.timer = setTimeout(pushHistory, 450); }
  function resetHistory() { hist.stack = []; hist.index = -1; pushHistory(); }
  function undo() {
    if (hist.timer) pushHistory();
    if (hist.index <= 0) return;
    hist.index--;
    const s = hist.stack[hist.index];
    editor.innerHTML = s.html;
    restoreOffsets(s.sel);
    afterChange(false);
  }
  function redo() {
    if (hist.index >= hist.stack.length - 1) return;
    hist.index++;
    const s = hist.stack[hist.index];
    editor.innerHTML = s.html;
    restoreOffsets(s.sel);
    afterChange(false);
  }

  function afterChange(record = true) {
    if (record) pushHistory();
    Rex.setDirty(true);
    schedulePageUpdate();
    updateToolbar();
  }

  // ---------------------------------------------------------------
  // Auswahl-Hilfen
  // ---------------------------------------------------------------
  let savedRange = null;
  function saveRange() {
    const sel = getSelection();
    if (sel.rangeCount && editor.contains(sel.anchorNode)) savedRange = sel.getRangeAt(0).cloneRange();
  }
  function restoreRange() {
    const cur = getSelection();
    if (document.activeElement === editor && cur.rangeCount && editor.contains(cur.anchorNode)) return;
    editor.focus({ preventScroll: true });
    if (savedRange) {
      const sel = getSelection();
      sel.removeAllRanges();
      sel.addRange(savedRange);
    }
  }
  function exec(cmd, val = null) {
    restoreRange();
    document.execCommand(cmd, false, val);
    saveRange();
    afterChange();
  }
  const BLOCK_SEL = 'p,h1,h2,h3,h4,h5,h6,blockquote,pre,li,div:not(.page-break):not(.toc),td,th';
  function selectedBlocks() {
    restoreRange();
    const sel = getSelection();
    if (!sel.rangeCount) return [];
    const r = sel.getRangeAt(0);
    const all = $$(BLOCK_SEL, editor).filter(b => r.intersectsNode(b) && !b.querySelector(BLOCK_SEL));
    if (!all.length) {
      let n = r.startContainer;
      if (n.nodeType === 3) n = n.parentElement;
      const b = n.closest && n.closest(BLOCK_SEL);
      if (b && editor.contains(b)) return [b];
      if (n === editor) { // Text direkt im Editor -> in Absatz packen
        document.execCommand('formatBlock', false, 'p');
        return selectedBlocks();
      }
    }
    return all;
  }
  function currentElement() {
    const sel = getSelection();
    if (!sel.rangeCount) return null;
    let n = sel.anchorNode;
    if (!n || !editor.contains(n)) return null;
    return n.nodeType === 3 ? n.parentElement : n;
  }

  // ---------------------------------------------------------------
  // Schrift
  // ---------------------------------------------------------------
  let pendingSize = null;
  function applyFontSize(pt) {
    pendingSize = pt;
    restoreRange();
    document.execCommand('styleWithCSS', false, true);
    document.execCommand('fontSize', false, '7');
    fixFontSizes();
    saveRange();
    afterChange();
  }
  function fixFontSizes() {
    if (pendingSize == null) return;
    for (const el of $$('font[size="7"]', editor)) {
      const s = document.createElement('span');
      s.style.fontSize = pendingSize + 'pt';
      while (el.firstChild) s.appendChild(el.firstChild);
      el.replaceWith(s);
    }
    for (const el of $$('span[style*="xxx-large"]', editor)) el.style.fontSize = pendingSize + 'pt';
  }
  function currentFontSizePt() {
    const el = currentElement();
    if (!el) return 11;
    return Math.round(parseFloat(getComputedStyle(el).fontSize) * 0.75 * 2) / 2;
  }
  function stepFont(dir) {
    const cur = currentFontSizePt();
    let next = dir > 0 ? SIZES.find(s => s > cur) : [...SIZES].reverse().find(s => s < cur);
    if (!next) next = dir > 0 ? cur + 2 : Math.max(1, cur - 1);
    applyFontSize(next);
  }

  // ---------------------------------------------------------------
  // Formatvorlagen
  // ---------------------------------------------------------------
  function applyStyle(st) {
    restoreRange();
    document.execCommand('formatBlock', false, st.tag);
    const blocks = selectedBlocks().filter(b => !/^(LI|TD|TH)$/.test(b.tagName));
    for (const b of blocks) {
      b.classList.remove('title', 'subtitle', 'code');
      if (st.cls) b.classList.add(st.cls);
      if (!b.classList.length) b.removeAttribute('class');
    }
    saveRange();
    afterChange();
  }
  function currentStyleId() {
    const el = currentElement();
    if (!el) return 'p';
    const b = el.closest('p,h1,h2,h3,h4,h5,h6,blockquote,pre');
    if (!b) return 'p';
    if (b.classList.contains('title')) return 'title';
    if (b.classList.contains('subtitle')) return 'subtitle';
    return b.tagName.toLowerCase();
  }

  // ---------------------------------------------------------------
  // Einzug
  // ---------------------------------------------------------------
  function indent(dir) {
    const el = currentElement();
    if (el && el.closest('li')) { exec(dir > 0 ? 'indent' : 'outdent'); return; }
    for (const b of selectedBlocks()) {
      if (/^(TD|TH)$/.test(b.tagName)) continue;
      const cur = parseFloat(b.style.marginLeft) || 0;
      const next = Math.max(0, cur + dir * 48);
      b.style.marginLeft = next ? next + 'px' : '';
      if (!b.getAttribute('style')) b.removeAttribute('style');
    }
    afterChange();
  }

  // ---------------------------------------------------------------
  // Tabellen
  // ---------------------------------------------------------------
  function currentCell() {
    const el = currentElement();
    return el ? el.closest('td,th') : null;
  }
  function insertTable(rows, cols) {
    const row = '<tr>' + '<td><br></td>'.repeat(cols) + '</tr>';
    exec('insertHTML', `<table><tbody>${row.repeat(rows)}</tbody></table><p><br></p>`);
  }
  function tableOp(op) {
    const cell = currentCell();
    if (!cell) return;
    const tr = cell.parentElement;
    const table = cell.closest('table');
    const idx = cell.cellIndex;
    const newCell = (tag = 'td') => { const c = document.createElement(tag); c.innerHTML = '<br>'; return c; };
    if (op === 'rowAbove' || op === 'rowBelow') {
      const nr = document.createElement('tr');
      for (const c of tr.cells) { const n = newCell(c.tagName.toLowerCase() === 'th' && op === 'rowAbove' ? 'th' : 'td'); if (c.colSpan > 1) n.colSpan = c.colSpan; nr.appendChild(n); }
      tr.parentElement.insertBefore(nr, op === 'rowAbove' ? tr : tr.nextSibling);
    } else if (op === 'colLeft' || op === 'colRight') {
      for (const r of table.rows) {
        const ref = r.cells[Math.min(idx, r.cells.length - 1)];
        const n = newCell(ref ? ref.tagName.toLowerCase() : 'td');
        if (!ref) r.appendChild(n);
        else r.insertBefore(n, op === 'colLeft' ? ref : ref.nextSibling);
      }
      $$('colgroup', table).forEach(c => c.remove());
    } else if (op === 'delRow') {
      if (table.rows.length <= 1) return tableOp('delTable');
      tr.remove();
    } else if (op === 'delCol') {
      if (tr.cells.length <= 1) return tableOp('delTable');
      for (const r of Array.from(table.rows)) if (r.cells[idx]) r.cells[idx].remove();
      $$('colgroup', table).forEach(c => c.remove());
    } else if (op === 'delTable') {
      const p = document.createElement('p');
      p.innerHTML = '<br>';
      table.replaceWith(p);
    } else if (op === 'headerRow') {
      const first = table.rows[0];
      const toTh = first.cells[0].tagName === 'TD';
      for (const c of Array.from(first.cells)) {
        const n = document.createElement(toTh ? 'th' : 'td');
        n.innerHTML = c.innerHTML;
        if (c.getAttribute('style')) n.setAttribute('style', c.getAttribute('style'));
        if (toTh && !n.style.backgroundColor) n.style.backgroundColor = '#d9e2f3';
        else if (!toTh && n.style.backgroundColor === 'rgb(217, 226, 243)') n.style.backgroundColor = '';
        if (c.colSpan > 1) n.colSpan = c.colSpan;
        c.replaceWith(n);
      }
    }
    afterChange();
  }

  // ---------------------------------------------------------------
  // Seite / Layout
  // ---------------------------------------------------------------
  function pageDims() {
    const land = settings.orientation === 'landscape';
    return { w: land ? settings.h : settings.w, h: land ? settings.w : settings.h };
  }
  let pageStyleEl = null;
  function applySettings() {
    const d = pageDims();
    const m = settings.margins;
    page.style.setProperty('--page-w', d.w * MM + 'px');
    page.style.setProperty('--page-h', d.h * MM + 'px');
    page.style.setProperty('--m-top', m.top * MM + 'px');
    page.style.setProperty('--m-right', m.right * MM + 'px');
    page.style.setProperty('--m-bottom', m.bottom * MM + 'px');
    page.style.setProperty('--m-left', m.left * MM + 'px');
    if (!pageStyleEl) { pageStyleEl = document.createElement('style'); document.head.appendChild(pageStyleEl); }
    pageStyleEl.textContent = `@page { size: ${d.w}mm ${d.h}mm; margin: ${m.top}mm ${m.right}mm ${m.bottom}mm ${m.left}mm; }`;
    $('#footerNum').hidden = !settings.pageNumbers;
    $('#pageNumBtn').classList.toggle('on', !!settings.pageNumbers);
    schedulePageUpdate();
  }

  let pageTimer = null;
  function schedulePageUpdate() {
    if (pageTimer) return;
    pageTimer = requestAnimationFrame(() => { pageTimer = null; updatePages(); });
  }
  function updatePages() {
    const d = pageDims();
    const contentH = (d.h - settings.margins.top - settings.margins.bottom) * MM;
    // Manuelle Seitenumbrueche bis zum Seitenende strecken
    for (const pb of $$('.page-break', editor)) {
      pb.style.marginBottom = '';
      const top = pb.offsetTop - editor.offsetTop;
      const rest = contentH - (top % contentH);
      pb.style.marginBottom = Math.max(14, rest - 14) + 'px';
    }
    const h = editor.scrollHeight;
    const pages = Math.max(1, Math.ceil((h - 2) / contentH));
    page.style.minHeight = (pages * contentH + (settings.margins.top + settings.margins.bottom) * MM) + 'px';
    const lines = $('#pageLines');
    let html = '';
    for (let i = 1; i < pages; i++) {
      html += `<div class="pl" style="top:${settings.margins.top * MM + i * contentH}px"><span>Seite ${i + 1}</span></div>`;
    }
    lines.innerHTML = html;
    // Statusleiste
    const text = editor.innerText || '';
    const words = (text.match(/[\p{L}\p{N}][\p{L}\p{N}'’\-]*/gu) || []).length;
    $('#stWords').textContent = `${words.toLocaleString('de-DE')} ${words === 1 ? 'Wort' : 'Wörter'}`;
    $('#stChars').textContent = `${text.replace(/\n/g, '').length.toLocaleString('de-DE')} Zeichen`;
    let cur = 1;
    const sel = getSelection();
    if (sel.rangeCount && editor.contains(sel.anchorNode)) {
      const rect = sel.getRangeAt(0).getBoundingClientRect();
      const er = editor.getBoundingClientRect();
      if (rect.height) cur = Math.min(pages, Math.max(1, Math.floor((rect.top - er.top) / zoom / contentH) + 1));
    }
    $('#stPages').textContent = `Seite ${cur} von ${pages}`;
    positionImgHandle();
  }

  // ---------------------------------------------------------------
  // Zoom
  // ---------------------------------------------------------------
  function setZoom(z) {
    zoom = Math.min(3, Math.max(0.3, Math.round(z * 100) / 100));
    page.style.zoom = zoom;
    $('#zoomRange').value = Math.round(zoom * 100);
    $('#zoomLabel').textContent = Math.round(zoom * 100) + ' %';
    positionImgHandle();
  }

  // ---------------------------------------------------------------
  // Bilder
  // ---------------------------------------------------------------
  let selImg = null;
  function selectImage(img) {
    if (selImg) selImg.classList.remove('sel');
    selImg = img;
    if (img) img.classList.add('sel');
    positionImgHandle();
  }
  function positionImgHandle() {
    const h = $('#imgHandle');
    if (!selImg || !editor.contains(selImg)) { h.hidden = true; selImg = null; return; }
    const r = selImg.getBoundingClientRect();
    const wr = workspace.getBoundingClientRect();
    h.hidden = false;
    h.style.left = (r.right - wr.left + workspace.scrollLeft - 6) + 'px';
    h.style.top = (r.bottom - wr.top + workspace.scrollTop - 6) + 'px';
  }
  $('#imgHandle').addEventListener('mousedown', e => {
    e.preventDefault();
    const img = selImg;
    if (!img) return;
    const startX = e.clientX;
    const startW = img.getBoundingClientRect().width / zoom;
    const maxW = editor.clientWidth;
    const move = ev => {
      const w = Math.max(20, Math.min(maxW, startW + (ev.clientX - startX) / zoom));
      img.style.width = Math.round(w) + 'px';
      img.style.height = '';
      positionImgHandle();
    };
    const up = () => {
      document.removeEventListener('mousemove', move);
      document.removeEventListener('mouseup', up);
      afterChange();
    };
    document.addEventListener('mousemove', move);
    document.addEventListener('mouseup', up);
  });
  async function insertImage(src) {
    if (!src) return;
    const { w } = await Rex.util.imageSize(src);
    const maxW = editor.clientWidth || 600;
    exec('insertHTML', `<img src="${src}" style="width:${Math.round(Math.min(w, maxW))}px">`);
  }

  // ---------------------------------------------------------------
  // Suchen & Ersetzen
  // ---------------------------------------------------------------
  let matches = [], matchIdx = -1;
  function findAll(q) {
    matches = [];
    if (!q) return;
    const ql = q.toLowerCase();
    const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
    let n;
    while ((n = walker.nextNode())) {
      const t = n.nodeValue.toLowerCase();
      let i = t.indexOf(ql);
      while (i >= 0) { matches.push({ node: n, start: i, end: i + q.length }); i = t.indexOf(ql, i + q.length); }
    }
  }
  function gotoMatch(dir) {
    const q = $('#findInput').value;
    findAll(q);
    if (!matches.length) { $('#findInfo').textContent = q ? 'Keine Treffer' : ''; return; }
    matchIdx = (matchIdx + dir + matches.length) % matches.length;
    const m = matches[matchIdx];
    const r = document.createRange();
    r.setStart(m.node, m.start);
    r.setEnd(m.node, m.end);
    const sel = getSelection();
    sel.removeAllRanges();
    sel.addRange(r);
    savedRange = r.cloneRange();
    const rect = r.getBoundingClientRect();
    const wr = workspace.getBoundingClientRect();
    if (rect.top < wr.top + 40 || rect.bottom > wr.bottom - 40) workspace.scrollTop += rect.top - wr.top - wr.height / 3;
    $('#findInfo').textContent = `Treffer ${matchIdx + 1} von ${matches.length}`;
  }
  function openFind(replace) {
    const bar = $('#findBar');
    bar.classList.add('open');
    $('#replaceRow').style.display = replace ? 'flex' : 'none';
    const s = getSelection().toString();
    if (s && s.length < 80) $('#findInput').value = s;
    $('#findInput').focus();
    $('#findInput').select();
  }
  $('#findInput').addEventListener('input', () => { matchIdx = -1; gotoMatch(1); });
  $('#findInput').addEventListener('keydown', e => {
    if (e.key === 'Enter') { e.preventDefault(); gotoMatch(e.shiftKey ? -1 : 1); }
    if (e.key === 'Escape') $('#findBar').classList.remove('open');
  });
  $('#findNext').onclick = () => gotoMatch(1);
  $('#findPrev').onclick = () => gotoMatch(-1);
  $('#findClose').onclick = () => $('#findBar').classList.remove('open');
  $('#replaceOne').onclick = () => {
    const q = $('#findInput').value;
    const sel = getSelection();
    if (q && sel.toString().toLowerCase() === q.toLowerCase()) {
      document.execCommand('insertText', false, $('#replaceInput').value);
      afterChange();
    }
    gotoMatch(1);
  };
  $('#replaceAll').onclick = () => {
    const q = $('#findInput').value;
    if (!q) return;
    findAll(q);
    const rep = $('#replaceInput').value;
    let count = 0;
    for (let i = matches.length - 1; i >= 0; i--) {
      const m = matches[i];
      m.node.nodeValue = m.node.nodeValue.slice(0, m.start) + rep + m.node.nodeValue.slice(m.end);
      count++;
    }
    afterChange();
    $('#findInfo').textContent = `${count} Ersetzung(en)`;
    Rex.toast(`${count} Stelle(n) ersetzt`);
  };

  // ---------------------------------------------------------------
  // Einfuegen aus der Zwischenablage (bereinigen)
  // ---------------------------------------------------------------
  const ALLOWED_TAGS = new Set(['P', 'BR', 'B', 'STRONG', 'I', 'EM', 'U', 'S', 'STRIKE', 'DEL', 'SUB', 'SUP', 'SPAN', 'A', 'IMG',
    'UL', 'OL', 'LI', 'TABLE', 'THEAD', 'TBODY', 'TFOOT', 'TR', 'TD', 'TH', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'BLOCKQUOTE',
    'PRE', 'CODE', 'HR', 'DIV', 'FONT', 'COLGROUP', 'COL']);
  const ALLOWED_CSS = ['font-weight', 'font-style', 'text-decoration', 'text-decoration-line', 'color', 'background-color',
    'font-size', 'font-family', 'text-align', 'vertical-align', 'margin-left', 'text-indent', 'line-height', 'width'];
  function sanitize(html) {
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const clean = (node) => {
      for (const ch of Array.from(node.childNodes)) {
        if (ch.nodeType === 8) { ch.remove(); continue; }
        if (ch.nodeType !== 1) continue;
        if (!ALLOWED_TAGS.has(ch.tagName)) {
          if (/^(SCRIPT|STYLE|META|LINK|TITLE|XML|OBJECT|IFRAME|SVG|NOSCRIPT|TEMPLATE|BUTTON|INPUT|SELECT|TEXTAREA|FORM)$/.test(ch.tagName)) { ch.remove(); continue; }
          clean(ch);
          ch.replaceWith(...ch.childNodes);
          continue;
        }
        const keep = {};
        for (const a of ['href', 'src', 'colspan', 'rowspan']) if (ch.hasAttribute(a)) keep[a] = ch.getAttribute(a);
        const style = [];
        for (const prop of ALLOWED_CSS) {
          const v = ch.style.getPropertyValue(prop);
          if (v && !(prop === 'width' && !/^(IMG|COL)$/.test(ch.tagName))) style.push(`${prop}:${v}`);
        }
        const cls = ['title', 'subtitle', 'page-break', 'code'].filter(c => ch.classList.contains(c));
        for (const a of Array.from(ch.attributes)) ch.removeAttribute(a.name);
        for (const [k, v] of Object.entries(keep)) {
          if ((k === 'href' && /^\s*javascript:/i.test(v)) || (k === 'src' && !/^(data:image\/|https?:|blob:)/i.test(v))) continue;
          ch.setAttribute(k, v);
        }
        if (style.length) ch.setAttribute('style', style.join(';'));
        if (cls.length) ch.className = cls.join(' ');
        if (ch.classList.contains('page-break')) ch.setAttribute('contenteditable', 'false');
        clean(ch);
      }
    };
    clean(doc.body);
    return doc.body.innerHTML;
  }

  editor.addEventListener('paste', async e => {
    const cd = e.clipboardData;
    if (!cd) return;
    const file = Array.from(cd.files || []).find(f => f.type.startsWith('image/'));
    if (file) {
      e.preventDefault();
      const r = new FileReader();
      r.onload = () => insertImage(r.result);
      r.readAsDataURL(file);
      return;
    }
    const html = cd.getData('text/html');
    e.preventDefault();
    saveRange();
    if (html && !plainPaste) exec('insertHTML', sanitize(html));
    else exec('insertText', cd.getData('text/plain'));
    plainPaste = false;
  });
  let plainPaste = false;

  // Dateien per Drag & Drop
  workspace.addEventListener('dragover', e => { e.preventDefault(); });
  workspace.addEventListener('drop', async e => {
    const f = e.dataTransfer.files[0];
    if (!f) return;
    e.preventDefault();
    if (f.type.startsWith('image/')) {
      const r = new FileReader();
      r.onload = () => insertImage(r.result);
      r.readAsDataURL(f);
    } else {
      Rex.loadFile({ name: f.name, path: Rex.util.filePath(f), data: new Uint8Array(await f.arrayBuffer()) });
    }
  });

  // ---------------------------------------------------------------
  // Symbolleiste aktualisieren
  // ---------------------------------------------------------------
  function updateToolbar() {
    const el = currentElement();
    if (!el) return;
    for (const cmd of ['bold', 'italic', 'underline', 'strikeThrough', 'subscript', 'superscript', 'insertUnorderedList', 'insertOrderedList',
      'justifyLeft', 'justifyCenter', 'justifyRight', 'justifyFull']) {
      let on = false;
      try { on = document.queryCommandState(cmd); } catch { /* */ }
      $$(`.btn[data-cmd="${cmd}"]`).forEach(b => b.classList.toggle('on', on));
    }
    const cs = getComputedStyle(el);
    const fam = cs.fontFamily.split(',')[0].replace(/["']/g, '').trim();
    const ff = $('#fontFamily');
    if (document.activeElement !== ff) {
      if (![...ff.options].some(o => o.value === fam)) ff.add(new Option(fam, fam));
      ff.value = fam;
    }
    const fs = $('#fontSize');
    if (document.activeElement !== fs) {
      const pt = Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2;
      if (![...fs.options].some(o => +o.value === pt)) fs.add(new Option(pt, pt));
      fs.value = String(pt);
    }
    const sid = currentStyleId();
    $$('#styleGallery button').forEach(b => b.classList.toggle('on', b.dataset.style === sid));
    const inTable = !!el.closest('td,th');
    const tt = $('#tableTab');
    if (tt.hidden === inTable) {
      tt.hidden = !inTable;
      if (!inTable && tt.classList.contains('active')) Rex.selectTab('start');
    }
    const blk = el.closest('p,h1,h2,h3,h4,h5,h6,blockquote,pre,li');
    if (blk) {
      const bcs = getComputedStyle(blk);
      if (document.activeElement !== $('#spaceBefore')) $('#spaceBefore').value = Math.round(parseFloat(bcs.marginTop) * 0.75);
      if (document.activeElement !== $('#spaceAfter')) $('#spaceAfter').value = Math.round(parseFloat(bcs.marginBottom) * 0.75);
    }
  }

  // ---------------------------------------------------------------
  // Befehle
  // ---------------------------------------------------------------
  const SYMBOLS = '€ £ $ ¥ © ® ™ § ¶ ° ± × ÷ ≠ ≈ ≤ ≥ ∞ √ ∑ π µ Ω α β γ δ ½ ¼ ¾ ² ³ → ← ↑ ↓ ↔ ⇒ • · … – — „ “ ” ‚ ‘ ’ « » ‹ › ✓ ✗ ★ ☆ ♥ ☺ ☎ ✉ ✂ ♪ ⚠'.split(' ');

  const commands = {
    paste: async () => {
      restoreRange();
      try {
        const items = await navigator.clipboard.read();
        for (const it of items) {
          const imgType = it.types.find(t => t.startsWith('image/'));
          if (imgType) {
            const b = await it.getType(imgType);
            const r = new FileReader();
            r.onload = () => insertImage(r.result);
            r.readAsDataURL(b);
            return;
          }
          if (it.types.includes('text/html')) { exec('insertHTML', sanitize(await (await it.getType('text/html')).text())); return; }
          if (it.types.includes('text/plain')) { exec('insertText', await (await it.getType('text/plain')).text()); return; }
        }
      } catch {
        if (!document.execCommand('paste')) Rex.toast('Bitte Strg+V zum Einfügen benutzen');
      }
    },
    cut: () => exec('cut'),
    copy: () => { restoreRange(); document.execCommand('copy'); },
    growFont: () => stepFont(1),
    shrinkFont: () => stepFont(-1),
    indent: () => indent(1),
    outdent: () => indent(-1),
    removeFormat: () => {
      exec('removeFormat');
      for (const b of selectedBlocks()) { if (!/^(TD|TH)$/.test(b.tagName)) b.removeAttribute('style'); }
      afterChange();
    },
    find: () => openFind(false),
    replace: () => openFind(true),
    selectAll: () => { editor.focus(); document.execCommand('selectAll'); saveRange(); },
    pageBreak: () => exec('insertHTML', '<div class="page-break" contenteditable="false"></div><p><br></p>'),
    image: async () => { saveRange(); insertImage(await Rex.util.readImageFile()); },
    hr: () => exec('insertHTML', '<hr><p><br></p>'),
    link: async () => {
      saveRange();
      const text = savedRange ? savedRange.toString() : '';
      const el = currentElement();
      const a = el && el.closest('a');
      const v = await Rex.dialog({ title: 'Link einfügen', fields: [
        { name: 'text', label: 'Anzuzeigender Text', value: a ? a.textContent : text },
        { name: 'url', label: 'Adresse (URL)', value: a ? a.getAttribute('href') : 'https://' }] });
      if (!v || !v.url) return;
      const url = /^[a-z]+:/i.test(v.url) ? v.url : 'https://' + v.url;
      if (a) { a.href = url; a.textContent = v.text || url; afterChange(); return; }
      if (text && (!v.text || v.text === text)) exec('createLink', url);
      else exec('insertHTML', `<a href="${esc(url)}">${esc(v.text || url)}</a>&nbsp;`);
    },
    toc: () => {
      const heads = $$('h1,h2,h3', editor);
      if (!heads.length) { Rex.toast('Keine Überschriften gefunden – nutze „Überschrift 1–3“ als Formatvorlage.'); return; }
      const d = pageDims();
      const contentH = (d.h - settings.margins.top - settings.margins.bottom) * MM;
      const counters = [0, 0, 0];
      const lines = heads.map(h => {
        const lvl = +h.tagName[1] - 1;
        counters[lvl]++;
        for (let i = lvl + 1; i < 3; i++) counters[i] = 0;
        const num = counters.slice(0, lvl + 1).join('.');
        const pg = Math.floor((h.offsetTop - editor.offsetTop) / contentH) + 1;
        return `<p style="margin-left:${lvl * 24}px">${num}&nbsp;&nbsp;${esc(h.textContent)} <span style="color:#7f7f7f">…… ${pg}</span></p>`;
      });
      const old = $('.toc', editor);
      const html = `<div class="toc"><p><b>Inhaltsverzeichnis</b></p>${lines.join('')}</div>`;
      if (old) { old.outerHTML = html; afterChange(); } else exec('insertHTML', html + '<p><br></p>');
    },
    pageNumbers: () => { settings.pageNumbers = !settings.pageNumbers; applySettings(); Rex.setDirty(true); },
    date: () => exec('insertText', new Date().toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' })),
    zoomIn: () => setZoom(zoom + 0.1),
    zoomOut: () => setZoom(zoom - 0.1),
    zoom100: () => setZoom(1),
    zoomWidth: () => setZoom((workspace.clientWidth - 60) / page.offsetWidth * zoom),
    focus: () => { document.body.classList.toggle('focus-mode'); if (document.body.classList.contains('focus-mode')) Rex.toast('Fokusmodus – Esc zum Beenden'); },
    wordCount: () => {
      const text = editor.innerText || '';
      const words = (text.match(/[\p{L}\p{N}][\p{L}\p{N}'’\-]*/gu) || []).length;
      const paras = $$('p,h1,h2,h3,h4,li,blockquote,pre', editor).filter(p => p.textContent.trim()).length;
      Rex.dialog({ title: 'Wörter zählen', cancel: null, html: `<table style="border-collapse:collapse;min-width:260px">
        <tr><td style="padding:4px 0">Seiten</td><td style="text-align:right">${$('#stPages').textContent.split(' von ')[1]}</td></tr>
        <tr><td style="padding:4px 0">Wörter</td><td style="text-align:right">${words.toLocaleString('de-DE')}</td></tr>
        <tr><td style="padding:4px 0">Zeichen (ohne Leerzeichen)</td><td style="text-align:right">${text.replace(/\s/g, '').length.toLocaleString('de-DE')}</td></tr>
        <tr><td style="padding:4px 0">Zeichen (mit Leerzeichen)</td><td style="text-align:right">${text.replace(/\n/g, '').length.toLocaleString('de-DE')}</td></tr>
        <tr><td style="padding:4px 0">Absätze</td><td style="text-align:right">${paras}</td></tr></table>` });
    },
    spell: () => {
      editor.spellcheck = !editor.spellcheck;
      $('#spellBtn').classList.toggle('on', editor.spellcheck);
      editor.innerHTML = editor.innerHTML; // Markierungen neu berechnen
      Rex.toast(editor.spellcheck ? 'Rechtschreibprüfung an' : 'Rechtschreibprüfung aus');
    },
    rowAbove: () => tableOp('rowAbove'), rowBelow: () => tableOp('rowBelow'),
    colLeft: () => tableOp('colLeft'), colRight: () => tableOp('colRight'),
    delRow: () => tableOp('delRow'), delCol: () => tableOp('delCol'), delTable: () => tableOp('delTable'),
    headerRow: () => tableOp('headerRow')
  };

  document.addEventListener('mousedown', e => {
    const b = e.target.closest('.btn, .styles button');
    if (b && !b.closest('.find-bar')) { saveRange(); e.preventDefault(); }
  });
  document.addEventListener('click', e => {
    const b = e.target.closest('[data-cmd]');
    if (!b) return;
    const c = b.dataset.cmd;
    if (commands[c]) commands[c]();
    else exec(c);
  });

  // Schriftart / -groesse
  const ff = $('#fontFamily');
  FONTS.forEach(f => { const o = new Option(f, f); o.style.fontFamily = f; ff.add(o); });
  ff.addEventListener('change', () => { document.execCommand('styleWithCSS', false, true); exec('fontName', ff.value); editor.focus(); });
  const fsSel = $('#fontSize');
  SIZES.forEach(s => fsSel.add(new Option(s, s)));
  fsSel.addEventListener('change', () => { applyFontSize(+fsSel.value); editor.focus(); });
  [ff, fsSel].forEach(s => s.addEventListener('mousedown', saveRange));

  // Farben
  let textColor = '#e00000', hiColor = '#ffff00', shadeColor = '#dbe6f6';
  $('#textColorBtn').addEventListener('click', e => {
    Rex.colorMenu(e.currentTarget, c => {
      textColor = c || '#000000';
      $('#textColorBar').style.background = textColor;
      document.execCommand('styleWithCSS', false, true);
      exec('foreColor', textColor);
    }, { noneLabel: 'Automatisch (Schwarz)' });
  });
  $('#highlightBtn').addEventListener('click', e => {
    Rex.colorMenu(e.currentTarget, c => {
      if (c) { hiColor = c; $('#highlightBar').style.background = c; }
      document.execCommand('styleWithCSS', false, true);
      exec('hiliteColor', c || 'transparent');
    }, { noneLabel: 'Keine Hervorhebung' });
  });
  $('#shadingBtn').addEventListener('click', e => {
    Rex.colorMenu(e.currentTarget, c => {
      if (c) { shadeColor = c; $('#shadingBar').style.background = c; }
      for (const b of selectedBlocks()) b.style.backgroundColor = c || '';
      afterChange();
    }, { noneLabel: 'Keine Schattierung' });
  });
  $('#cellShadeBtn').addEventListener('click', e => {
    const cell = currentCell();
    if (!cell) return;
    Rex.colorMenu(e.currentTarget, c => {
      const table = cell.closest('table');
      const sel = getSelection();
      const r = sel.rangeCount ? sel.getRangeAt(0) : null;
      const cells = $$('td,th', table).filter(td => r && r.intersectsNode(td));
      (cells.length ? cells : [cell]).forEach(td => { td.style.backgroundColor = c || ''; });
      afterChange();
    }, { noneLabel: 'Keine Farbe' });
  });

  // Zeilenabstand
  $('#lineSpacingBtn').addEventListener('click', e => {
    Rex.menu(e.currentTarget, [
      ...[1, 1.15, 1.5, 2, 2.5, 3].map(v => ({ label: v.toLocaleString('de-DE', { minimumFractionDigits: 1 }), action: () => {
        for (const b of selectedBlocks()) b.style.lineHeight = v === 1.15 ? '' : String(v);
        afterChange();
      } })),
      '-',
      { label: 'Abstand vor Absatz hinzufügen', action: () => { for (const b of selectedBlocks()) b.style.marginTop = '12pt'; afterChange(); } },
      { label: 'Abstand nach Absatz entfernen', action: () => { for (const b of selectedBlocks()) b.style.marginBottom = '0pt'; afterChange(); } }
    ]);
  });
  for (const id of ['spaceBefore', 'spaceAfter']) {
    $('#' + id).addEventListener('mousedown', saveRange);
    $('#' + id).addEventListener('focus', saveRange);
    $('#' + id).addEventListener('change', e => {
      const v = Math.max(0, +e.target.value || 0);
      for (const b of selectedBlocks()) b.style[id === 'spaceBefore' ? 'marginTop' : 'marginBottom'] = v + 'pt';
      afterChange();
    });
  }

  // Formatvorlagen-Galerie
  const gal = $('#styleGallery');
  for (const st of STYLES) {
    const b = document.createElement('button');
    b.dataset.style = st.id;
    b.textContent = st.label;
    b.style.cssText = st.css;
    b.onclick = () => applyStyle(st);
    gal.appendChild(b);
  }

  // Tabelle einfuegen (Raster)
  $('#tableBtn').addEventListener('click', e => {
    saveRange();
    Rex.menu(e.currentTarget, m => {
      const R = 8, C = 10;
      const label = document.createElement('div');
      label.className = 'gp-label';
      label.textContent = 'Tabelle einfügen';
      const g = document.createElement('div');
      g.className = 'grid-picker';
      g.style.gridTemplateColumns = `repeat(${C}, 16px)`;
      for (let r = 0; r < R; r++) for (let c = 0; c < C; c++) {
        const d = document.createElement('div');
        d.dataset.r = r; d.dataset.c = c;
        g.appendChild(d);
      }
      g.addEventListener('mouseover', ev => {
        const d = ev.target.closest('[data-r]');
        if (!d) return;
        const rr = +d.dataset.r, cc = +d.dataset.c;
        for (const x of g.children) x.classList.toggle('hl', +x.dataset.r <= rr && +x.dataset.c <= cc);
        label.textContent = `${cc + 1} × ${rr + 1} Tabelle`;
      });
      g.addEventListener('click', ev => {
        const d = ev.target.closest('[data-r]');
        if (!d) return;
        Rex.closeMenu();
        insertTable(+d.dataset.r + 1, +d.dataset.c + 1);
      });
      m.appendChild(label);
      m.appendChild(g);
      const more = document.createElement('button');
      more.textContent = 'Tabelle mit eigener Größe…';
      more.onclick = async () => {
        Rex.closeMenu();
        const v = await Rex.dialog({ title: 'Tabelle einfügen', fields: [
          { name: 'cols', label: 'Spalten', type: 'number', value: 3, min: 1, max: 30 },
          { name: 'rows', label: 'Zeilen', type: 'number', value: 4, min: 1, max: 200 }] });
        if (v) insertTable(Math.max(1, Math.min(200, v.rows)), Math.max(1, Math.min(30, v.cols)));
      };
      m.appendChild(more);
    });
  });

  // Sonderzeichen
  $('#symbolBtn').addEventListener('click', e => {
    saveRange();
    Rex.menu(e.currentTarget, m => {
      const g = document.createElement('div');
      g.style.cssText = 'display:grid;grid-template-columns:repeat(10,28px);gap:2px;padding:4px';
      for (const s of SYMBOLS) {
        const b = document.createElement('button');
        b.textContent = s;
        b.style.cssText = 'width:28px;height:28px;padding:0;justify-content:center;font-size:15px';
        b.onclick = () => { Rex.closeMenu(); exec('insertText', s); };
        g.appendChild(b);
      }
      m.appendChild(g);
    });
  });

  // Seite einrichten
  $('#marginsBtn').addEventListener('click', e => {
    const set = (t, r, b, l) => { settings.margins = { top: t, right: r, bottom: b, left: l }; applySettings(); Rex.setDirty(true); };
    Rex.menu(e.currentTarget, [
      { label: 'Normal (2,5 cm)', action: () => set(25, 25, 25, 25) },
      { label: 'Schmal (1,27 cm)', action: () => set(12.7, 12.7, 12.7, 12.7) },
      { label: 'Mittel (2,54 / 1,91 cm)', action: () => set(25.4, 19.1, 25.4, 19.1) },
      { label: 'Breit (2,54 / 5,08 cm)', action: () => set(25.4, 50.8, 25.4, 50.8) },
      '-',
      { label: 'Benutzerdefiniert…', action: async () => {
        const m = settings.margins;
        const v = await Rex.dialog({ title: 'Seitenränder (in cm)', fields: [
          { name: 'top', label: 'Oben', type: 'number', step: 0.1, value: m.top / 10 },
          { name: 'bottom', label: 'Unten', type: 'number', step: 0.1, value: m.bottom / 10 },
          { name: 'left', label: 'Links', type: 'number', step: 0.1, value: m.left / 10 },
          { name: 'right', label: 'Rechts', type: 'number', step: 0.1, value: m.right / 10 }] });
        if (v) set(v.top * 10, v.right * 10, v.bottom * 10, v.left * 10);
      } }
    ]);
  });
  $('#orientBtn').addEventListener('click', e => {
    Rex.menu(e.currentTarget, [
      { label: 'Hochformat', icon: 'rectangle-vertical', action: () => { settings.orientation = 'portrait'; applySettings(); Rex.setDirty(true); } },
      { label: 'Querformat', icon: 'rectangle-horizontal', action: () => { settings.orientation = 'landscape'; applySettings(); Rex.setDirty(true); } }
    ]);
  });
  $('#sizeBtn').addEventListener('click', e => {
    Rex.menu(e.currentTarget, Object.entries(PAGE_SIZES).map(([k, [w, h]]) => ({
      label: `${k} (${(w / 10).toLocaleString('de-DE')} × ${(h / 10).toLocaleString('de-DE')} cm)`,
      action: () => { settings.size = k; settings.w = w; settings.h = h; applySettings(); Rex.setDirty(true); }
    })));
  });

  $('#zoomRange').addEventListener('input', e => setZoom(+e.target.value / 100));

  // ---------------------------------------------------------------
  // Editor-Ereignisse
  // ---------------------------------------------------------------
  editor.addEventListener('input', () => {
    fixFontSizes();
    // Leeren Editor immer mit Absatz fuellen
    if (!editor.firstElementChild && !editor.textContent) editor.innerHTML = '<p><br></p>';
    scheduleHistory();
    Rex.setDirty(true);
    schedulePageUpdate();
  });
  editor.addEventListener('beforeinput', e => {
    if (e.inputType === 'historyUndo') { e.preventDefault(); undo(); }
    if (e.inputType === 'historyRedo') { e.preventDefault(); redo(); }
  });
  document.addEventListener('selectionchange', () => {
    if (editor.contains(getSelection().anchorNode)) {
      saveRange();
      updateToolbar();
      schedulePageUpdate();
    }
  });
  editor.addEventListener('click', e => {
    selectImage(e.target.tagName === 'IMG' ? e.target : null);
    if (e.target.tagName === 'A' && (e.ctrlKey || e.metaKey)) window.open(e.target.href, '_blank');
  });
  workspace.addEventListener('scroll', positionImgHandle);
  window.addEventListener('resize', positionImgHandle);

  editor.addEventListener('keydown', e => {
    const ctrl = e.ctrlKey || e.metaKey;
    const k = e.key.toLowerCase();
    if (selImg && (e.key === 'Delete' || e.key === 'Backspace')) {
      e.preventDefault();
      selImg.remove();
      selectImage(null);
      afterChange();
      return;
    }
    if (ctrl && !e.shiftKey && k === 'z') { e.preventDefault(); undo(); return; }
    if (ctrl && (k === 'y' || (e.shiftKey && k === 'z'))) { e.preventDefault(); redo(); return; }
    if (ctrl && e.shiftKey && k === 'v') { plainPaste = true; return; }
    if (ctrl && k === 'b') { e.preventDefault(); exec('bold'); return; }
    if (ctrl && k === 'i') { e.preventDefault(); exec('italic'); return; }
    if (ctrl && k === 'u') { e.preventDefault(); exec('underline'); return; }
    if (ctrl && k === 'l') { e.preventDefault(); exec('justifyLeft'); return; }
    if (ctrl && k === 'e') { e.preventDefault(); exec('justifyCenter'); return; }
    if (ctrl && k === 'r') { e.preventDefault(); exec('justifyRight'); return; }
    if (ctrl && k === 'j') { e.preventDefault(); exec('justifyFull'); return; }
    if (ctrl && k === 'k') { e.preventDefault(); commands.link(); return; }
    if (ctrl && k === 'f') { e.preventDefault(); openFind(false); return; }
    if (ctrl && k === 'h') { e.preventDefault(); openFind(true); return; }
    if (ctrl && e.key === 'Enter') { e.preventDefault(); commands.pageBreak(); return; }
    if (ctrl && (e.key === '+' || e.key === '=')) { e.preventDefault(); setZoom(zoom + 0.1); return; }
    if (ctrl && e.key === '-') { e.preventDefault(); setZoom(zoom - 0.1); return; }
    if (ctrl && e.shiftKey && (e.key === '>' || e.key === '.')) { e.preventDefault(); stepFont(1); return; }
    if (ctrl && e.shiftKey && (e.key === '<' || e.key === ',')) { e.preventDefault(); stepFont(-1); return; }
    if (e.key === 'Tab') {
      e.preventDefault();
      saveRange();
      const cell = currentCell();
      if (cell) {
        const cells = $$('td,th', cell.closest('table'));
        let i = cells.indexOf(cell) + (e.shiftKey ? -1 : 1);
        if (i >= cells.length) { tableOp('rowBelow'); return moveToCell($$('td,th', cell.closest('table'))[i]); }
        if (i >= 0) moveToCell(cells[i]);
        return;
      }
      if (currentElement() && currentElement().closest('li')) { exec(e.shiftKey ? 'outdent' : 'indent'); return; }
      exec('insertText', ' ');
      return;
    }
    if (e.key === 'Escape' && document.body.classList.contains('focus-mode')) document.body.classList.remove('focus-mode');
    // Bei Enter neuen Absatz statt <div>
    if (e.key === 'Enter' && !e.shiftKey) setTimeout(() => {
      const el = currentElement();
      if (el && el.tagName === 'DIV' && el.parentElement === editor && !el.className) {
        const p = document.createElement('p');
        if (el.getAttribute('style')) p.setAttribute('style', el.getAttribute('style'));
        while (el.firstChild) p.appendChild(el.firstChild);
        el.replaceWith(p);
        const r = document.createRange();
        r.setStart(p, 0); r.collapse(true);
        getSelection().removeAllRanges(); getSelection().addRange(r);
      }
    }, 0);
  });
  function moveToCell(td) {
    if (!td) return;
    const r = document.createRange();
    r.selectNodeContents(td);
    r.collapse(false);
    const sel = getSelection();
    sel.removeAllRanges();
    sel.addRange(r);
    saveRange();
  }
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape' && document.body.classList.contains('focus-mode')) document.body.classList.remove('focus-mode');
    if ((e.ctrlKey || e.metaKey) && !editor.contains(document.activeElement) && document.activeElement === document.body) {
      const k = e.key.toLowerCase();
      if (k === 'z') { e.preventDefault(); undo(); }
      if (k === 'y') { e.preventDefault(); redo(); }
    }
  });

  // ---------------------------------------------------------------
  // Laden / Speichern
  // ---------------------------------------------------------------
  function setContent(html, newSettings) {
    editor.innerHTML = html || '<p><br></p>';
    settings = newSettings ? { ...defaultSettings(), ...newSettings, margins: { ...defaultSettings().margins, ...(newSettings.margins || {}) } } : defaultSettings();
    selectImage(null);
    applySettings();
    resetHistory();
    workspace.scrollTop = 0;
    const r = document.createRange();
    r.setStart(editor.firstChild || editor, 0);
    r.collapse(true);
    savedRange = r;
    updatePages();
  }

  async function load(bytes, ext) {
    if (ext === 'docx') {
      const res = await RexDocx.importDocx(bytes);
      setContent(res.html, res.settings);
    } else if (ext === 'html' || ext === 'htm') {
      const text = new TextDecoder().decode(bytes);
      setContent(sanitize(text));
    } else if (ext === 'md') {
      setContent(markdownToHtml(new TextDecoder().decode(bytes)));
    } else {
      const text = new TextDecoder().decode(bytes);
      setContent(text.split(/\r?\n/).map(l => `<p>${esc(l) || '<br>'}</p>`).join(''));
    }
  }

  function markdownToHtml(md) {
    const inline = s => esc(s)
      .replace(/\*\*(.+?)\*\*/g, '<b>$1</b>')
      .replace(/\*(.+?)\*/g, '<i>$1</i>')
      .replace(/`(.+?)`/g, '<span style="font-family:Consolas">$1</span>')
      .replace(/\[(.+?)\]\((.+?)\)/g, '<a href="$2">$1</a>');
    let html = '', list = null;
    for (const line of md.split(/\r?\n/)) {
      const h = /^(#{1,4})\s+(.*)$/.exec(line);
      const li = /^\s*([-*]|\d+\.)\s+(.*)$/.exec(line);
      if (li) {
        const t = /\d/.test(li[1]) ? 'ol' : 'ul';
        if (list !== t) { if (list) html += `</${list}>`; html += `<${t}>`; list = t; }
        html += `<li>${inline(li[2])}</li>`;
        continue;
      }
      if (list) { html += `</${list}>`; list = null; }
      if (h) html += `<h${h[1].length}>${inline(h[2])}</h${h[1].length}>`;
      else if (/^>\s?/.test(line)) html += `<blockquote>${inline(line.replace(/^>\s?/, ''))}</blockquote>`;
      else if (/^(-{3,}|\*{3,})$/.test(line.trim())) html += '<hr>';
      else if (line.trim()) html += `<p>${inline(line)}</p>`;
    }
    if (list) html += `</${list}>`;
    return html;
  }

  const HTML_EXPORT_CSS = `body{font-family:Calibri,Carlito,Arial,sans-serif;font-size:11pt;line-height:1.15;max-width:800px;margin:40px auto;padding:0 20px;color:#000}
p{margin:0 0 8pt}h1,h2,h3,h4{font-family:'Calibri Light',Calibri,sans-serif;font-weight:normal}h1{font-size:16pt;color:#2f5496}h2{font-size:13pt;color:#2f5496}h3{font-size:12pt;color:#1f3763}
.title{font-size:28pt;font-family:'Calibri Light',Calibri,sans-serif}.subtitle{color:#5a5a5a}blockquote{font-style:italic;color:#404040;text-align:center}
table{border-collapse:collapse;width:100%}td,th{border:1px solid #000;padding:2pt 5pt;vertical-align:top}img{max-width:100%}
.page-break{break-after:page}pre{font-family:Consolas,monospace;background:#f4f4f4;padding:6pt}`;

  async function save(fmt) {
    if (hist.timer) pushHistory();
    selectImage(null);
    if (fmt === 'docx') return RexDocx.exportDocx(editor, settings, { title: Rex.doc.name });
    if (fmt === 'txt') return new TextEncoder().encode(editor.innerText.replace(/ /g, '\t'));
    if (fmt === 'html') {
      const body = editor.cloneNode(true);
      $$('[contenteditable]', body).forEach(e => e.removeAttribute('contenteditable'));
      return new TextEncoder().encode(`<!DOCTYPE html>\n<html lang="de"><head><meta charset="utf-8"><title>${esc(Rex.doc.name)}</title><style>${HTML_EXPORT_CSS}</style></head><body>\n${body.innerHTML}\n</body></html>`);
    }
    throw new Error('Unbekanntes Format: ' + fmt);
  }

  // ---------------------------------------------------------------
  // Vorlagen
  // ---------------------------------------------------------------
  const today = () => new Date().toLocaleDateString('de-DE', { day: 'numeric', month: 'long', year: 'numeric' });
  const TEMPLATES = [
    { name: 'Leeres Dokument', icon: 'file', docName: 'Dokument', html: () => '<p><br></p>' },
    { name: 'Brief', icon: 'mail', docName: 'Brief', html: () => `
      <p>Max Mustermann<br>Musterstraße 1<br>12345 Musterstadt</p><p><br></p>
      <p>Firma Beispiel GmbH<br>Frau Erika Beispiel<br>Beispielweg 5<br>54321 Beispielstadt</p>
      <p style="text-align:right">Musterstadt, ${today()}</p><p><br></p>
      <p><b>Betreff: Ihr Anliegen</b></p><p><br></p>
      <p>Sehr geehrte Frau Beispiel,</p>
      <p>hier steht der Text Ihres Briefes. Klicken Sie einfach hinein und beginnen Sie zu schreiben.</p>
      <p>Mit freundlichen Grüßen</p><p><br></p><p>Max Mustermann</p>` },
    { name: 'Bericht', icon: 'notebook-text', docName: 'Bericht', html: () => `
      <p class="title">Titel des Berichts</p><p class="subtitle">Untertitel · ${today()}</p>
      <h1>Einleitung</h1><p>Beschreibe hier, worum es in diesem Bericht geht.</p>
      <h1>Ergebnisse</h1><h2>Erster Punkt</h2><ul><li>Wichtige Erkenntnis</li><li>Noch eine Erkenntnis</li></ul>
      <table><tbody><tr><th style="background-color:#d9e2f3">Kategorie</th><th style="background-color:#d9e2f3">Wert</th></tr><tr><td>A</td><td>42</td></tr><tr><td>B</td><td>17</td></tr></tbody></table>
      <h1>Fazit</h1><p>Fasse die wichtigsten Punkte zusammen.</p>` },
    { name: 'Lebenslauf', icon: 'user-round', docName: 'Lebenslauf', html: () => `
      <p class="title">Max Mustermann</p><p class="subtitle">Musterstraße 1 · 12345 Musterstadt · max@example.de · 0123 456789</p>
      <h1>Ausbildung</h1><p><b>2018 – 2024</b>&emsp;Realschule Musterstadt – Mittlere Reife</p>
      <h1>Praktische Erfahrung</h1><p><b>2023</b>&emsp;Praktikum bei Beispiel GmbH</p>
      <h1>Kenntnisse</h1><ul><li>Sprachen: Deutsch (Muttersprache), Englisch (gut)</li><li>EDV: REX Office, Programmieren</li></ul>
      <h1>Hobbys</h1><p>Gaming, Fußball, Musik</p>` }
  ];

  // ---------------------------------------------------------------
  // Start
  // ---------------------------------------------------------------
  document.execCommand('defaultParagraphSeparator', false, 'p');
  document.execCommand('styleWithCSS', false, true);

  Rex.init({
    app: 'writer',
    letter: 'T',
    icon: 'file-text',
    defaultName: 'Dokument',
    openFormats: [
      { name: 'Word-Dokument', extensions: ['docx'] },
      { name: 'Textdatei', extensions: ['txt'] },
      { name: 'Webseite', extensions: ['html', 'htm'] },
      { name: 'Markdown', extensions: ['md'] }
    ],
    saveFormats: [{ name: 'Word-Dokument', extensions: ['docx'], hint: 'Kompatibel mit Microsoft Word, LibreOffice und Google Docs' }],
    exportFormats: [
      { name: 'Webseite', extensions: ['html'], hint: 'Zum Veröffentlichen im Internet' },
      { name: 'Nur Text', extensions: ['txt'], hint: 'Ohne Formatierung' }
    ],
    templates: TEMPLATES.map(t => ({ ...t, create: () => setContent(t.html()) })),
    load, save, undo, redo,
    shortcutsHelp: [['Strg+B / I / U', 'Fett / Kursiv / Unterstrichen'], ['Strg+L / E / R / J', 'Ausrichtung'],
      ['Strg+K', 'Link'], ['Strg+F / H', 'Suchen / Ersetzen'], ['Strg+Enter', 'Seitenumbruch'], ['Strg+Umschalt+V', 'Als Text einfügen']],
    pdf: {
      pageSize: () => { const d = pageDims(); return { width: d.w / 25.4, height: d.h / 25.4 }; },
      margins: () => ({ top: settings.margins.top / 25.4, bottom: settings.margins.bottom / 25.4, left: settings.margins.left / 25.4, right: settings.margins.right / 25.4 }),
      before: () => { selectImage(null); $$('.page-break', editor).forEach(pb => { pb.dataset.mb = pb.style.marginBottom; pb.style.marginBottom = '0'; }); },
      after: () => { schedulePageUpdate(); }
    }
  });

  setContent('<p><br></p>');
  editor.focus();
  new ResizeObserver(schedulePageUpdate).observe(editor);
  window.RexWriter = { editor, setContent, save, load, get settings() { return settings; } };
})();
