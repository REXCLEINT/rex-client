/* REX Office – gemeinsame Funktionen fuer alle Programme */
(function () {
  'use strict';

  const native = window.rexNative || null;
  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

  const APP_NAMES = { writer: 'REX Text', calc: 'REX Tabelle', present: 'REX Präsentation', home: 'REX Office' };
  const EXT_APP = { docx: 'writer', txt: 'writer', html: 'writer', htm: 'writer', md: 'writer',
    xlsx: 'calc', csv: 'calc', pptx: 'present' };

  // ---------- Theme ----------
  function store(key, val) {
    try {
      if (val === undefined) return localStorage.getItem('rex.' + key);
      localStorage.setItem('rex.' + key, val);
    } catch { return null; }
  }
  function applyTheme(t) {
    document.documentElement.dataset.theme = t;
    const btn = $('#themeBtn');
    if (btn) { btn.innerHTML = `<i data-lucide="${t === 'dark' ? 'sun' : 'moon'}"></i>`; icons(); }
  }
  applyTheme(store('theme') || 'light');

  function icons() { if (window.lucide) window.lucide.createIcons({ attrs: { 'stroke-width': 1.8 } }); }

  // ---------- Hilfsfunktionen ----------
  function filePath(f) { return native && native.pathForFile ? native.pathForFile(f) : (f.path || null); }
  function ext(name) { const m = /\.([^.]+)$/.exec(name || ''); return m ? m[1].toLowerCase() : ''; }
  function stripExt(name) { return (name || '').replace(/\.[^.]+$/, ''); }
  function esc(s) {
    return String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }
  function toBytes(data) {
    if (data instanceof Uint8Array) return data;
    if (data instanceof ArrayBuffer) return new Uint8Array(data);
    if (data && data.buffer instanceof ArrayBuffer) return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
    if (typeof data === 'string') return new TextEncoder().encode(data);
    return new Uint8Array(data);
  }
  async function blobToBytes(blob) { return new Uint8Array(await blob.arrayBuffer()); }
  function bytesToDataURL(bytes, mime) {
    let bin = '';
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    return `data:${mime};base64,${btoa(bin)}`;
  }
  function dataURLToBytes(url) {
    const b64 = url.split(',')[1] || '';
    const bin = atob(b64);
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }
  function mimeFromDataURL(url) { const m = /^data:([^;,]+)/.exec(url); return m ? m[1] : 'application/octet-stream'; }
  function readImageFile() {
    return new Promise(resolve => {
      const inp = document.createElement('input');
      inp.type = 'file';
      inp.accept = 'image/*';
      inp.onchange = () => {
        const f = inp.files[0];
        if (!f) return resolve(null);
        const r = new FileReader();
        r.onload = () => resolve(r.result);
        r.readAsDataURL(f);
      };
      inp.click();
    });
  }
  function imageSize(src) {
    return new Promise(resolve => {
      const img = new Image();
      img.onload = () => resolve({ w: img.naturalWidth, h: img.naturalHeight });
      img.onerror = () => resolve({ w: 300, h: 200 });
      img.src = src;
    });
  }
  // Bilder als PNG/JPEG normalisieren (z. B. WebP/SVG nicht ueberall unterstuetzt)
  async function normalizeImage(src) {
    const mime = mimeFromDataURL(src);
    if (mime === 'image/png' || mime === 'image/jpeg' || mime === 'image/gif') return src;
    const { w, h } = await imageSize(src);
    const c = document.createElement('canvas');
    c.width = w; c.height = h;
    const img = new Image();
    await new Promise(r => { img.onload = r; img.onerror = r; img.src = src; });
    c.getContext('2d').drawImage(img, 0, 0);
    return c.toDataURL('image/png');
  }

  // ---------- Toast ----------
  let toastTimer;
  function toast(msg, ms = 2200) {
    let el = $('.toast');
    if (!el) { el = document.createElement('div'); el.className = 'toast'; document.body.appendChild(el); }
    el.textContent = msg;
    el.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => el.classList.remove('show'), ms);
  }

  // ---------- Dialoge (prompt() gibt es in Electron nicht) ----------
  function dialog({ title, fields = [], ok = 'OK', cancel = 'Abbrechen', html = '' }) {
    return new Promise(resolve => {
      const bg = document.createElement('div');
      bg.className = 'modal-bg';
      const fieldHTML = fields.map((f, i) => {
        const id = `dlg_${i}`;
        if (f.type === 'select') {
          return `<div class="field"><label for="${id}">${esc(f.label)}</label><select id="${id}">${
            f.options.map(o => `<option value="${esc(o.value ?? o)}" ${String(o.value ?? o) === String(f.value) ? 'selected' : ''}>${esc(o.label ?? o)}</option>`).join('')
          }</select></div>`;
        }
        if (f.type === 'checkbox') {
          return `<div class="field inline"><input type="checkbox" id="${id}" ${f.value ? 'checked' : ''}><label for="${id}">${esc(f.label)}</label></div>`;
        }
        if (f.type === 'textarea') {
          return `<div class="field"><label for="${id}">${esc(f.label)}</label><textarea id="${id}" rows="${f.rows || 4}">${esc(f.value ?? '')}</textarea></div>`;
        }
        return `<div class="field"><label for="${id}">${esc(f.label)}</label><input id="${id}" type="${f.type || 'text'}" value="${esc(f.value ?? '')}" ${f.min != null ? `min="${f.min}"` : ''} ${f.max != null ? `max="${f.max}"` : ''} ${f.step != null ? `step="${f.step}"` : ''} placeholder="${esc(f.placeholder || '')}"></div>`;
      }).join('');
      bg.innerHTML = `<div class="modal" role="dialog"><h3>${esc(title)}</h3>${html}${fieldHTML}
        <div class="actions">${cancel ? `<button data-a="cancel">${esc(cancel)}</button>` : ''}<button class="primary" data-a="ok">${esc(ok)}</button></div></div>`;
      document.body.appendChild(bg);
      const sel = window.getSelection();
      const savedRange = sel.rangeCount ? sel.getRangeAt(0).cloneRange() : null;
      const first = bg.querySelector('input,select,textarea');
      if (first) { first.focus(); if (first.select) first.select(); } else bg.querySelector('.primary').focus();
      const finish = (okd) => {
        const vals = {};
        fields.forEach((f, i) => {
          const el = bg.querySelector(`#dlg_${i}`);
          vals[f.name || i] = f.type === 'checkbox' ? el.checked : (f.type === 'number' ? Number(el.value) : el.value);
        });
        bg.remove();
        if (savedRange) { sel.removeAllRanges(); sel.addRange(savedRange); }
        resolve(okd ? vals : null);
      };
      bg.addEventListener('click', e => {
        const a = e.target.dataset.a;
        if (a === 'ok') finish(true);
        else if (a === 'cancel' || e.target === bg) finish(false);
      });
      bg.addEventListener('keydown', e => {
        e.stopPropagation();
        if (e.key === 'Enter' && e.target.tagName !== 'TEXTAREA') { e.preventDefault(); finish(true); }
        if (e.key === 'Escape') { e.preventDefault(); finish(false); }
      });
    });
  }
  const confirmBox = (title, text, ok = 'Ja', cancel = 'Nein') =>
    dialog({ title, html: `<p style="margin:0 0 6px">${esc(text)}</p>`, ok, cancel }).then(v => !!v);
  const alertBox = (title, text) => dialog({ title, html: `<p style="margin:0 0 6px">${esc(text)}</p>`, cancel: null });

  // ---------- Dropdown-Menues ----------
  let openMenuEl = null;
  function closeMenu() { if (openMenuEl) { openMenuEl.remove(); openMenuEl = null; } }
  function menu(anchor, build) {
    closeMenu();
    const m = document.createElement('div');
    m.className = 'menu';
    if (typeof build === 'string') m.innerHTML = build;
    else if (Array.isArray(build)) {
      for (const it of build) {
        if (it === '-') { m.appendChild(document.createElement('hr')); continue; }
        const b = document.createElement('button');
        b.innerHTML = `${it.icon ? `<i data-lucide="${it.icon}"></i>` : ''}<span style="${it.style || ''}">${it.html || esc(it.label)}</span>`;
        b.onclick = () => { closeMenu(); it.action(); };
        m.appendChild(b);
      }
    } else build(m);
    document.body.appendChild(m);
    const r = anchor.getBoundingClientRect();
    m.style.left = Math.min(r.left, innerWidth - m.offsetWidth - 8) + 'px';
    m.style.top = Math.min(r.bottom + 2, innerHeight - m.offsetHeight - 8) + 'px';
    openMenuEl = m;
    icons();
    m.addEventListener('mousedown', e => { if (!e.target.closest('input,select,textarea')) e.preventDefault(); });
    return m;
  }
  document.addEventListener('mousedown', e => { if (openMenuEl && !openMenuEl.contains(e.target)) closeMenu(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape') closeMenu(); });

  const PALETTE = ['#000000', '#434343', '#666666', '#999999', '#b7b7b7', '#cccccc', '#d9d9d9', '#efefef', '#f3f3f3', '#ffffff',
    '#980000', '#ff0000', '#ff9900', '#ffff00', '#00ff00', '#00ffff', '#4a86e8', '#0000ff', '#9900ff', '#ff00ff',
    '#e6b8af', '#f4cccc', '#fce5cd', '#fff2cc', '#d9ead3', '#d0e0e3', '#c9daf8', '#cfe2f3', '#d9d2e9', '#ead1dc',
    '#cc4125', '#e06666', '#f6b26b', '#ffd966', '#93c47d', '#76a5af', '#6d9eeb', '#6fa8dc', '#8e7cc3', '#c27ba0',
    '#85200c', '#990000', '#b45f06', '#bf9000', '#38761d', '#134f5c', '#1155cc', '#0b5394', '#351c75', '#741b47'];
  function colorMenu(anchor, onPick, { noneLabel } = {}) {
    menu(anchor, m => {
      if (noneLabel) {
        const b = document.createElement('button');
        b.textContent = noneLabel;
        b.onclick = () => { closeMenu(); onPick(null); };
        m.appendChild(b);
      }
      const g = document.createElement('div');
      g.className = 'swatches';
      for (const c of PALETTE) {
        const b = document.createElement('button');
        b.style.background = c;
        b.title = c;
        b.onclick = () => { closeMenu(); onPick(c); };
        g.appendChild(b);
      }
      m.appendChild(g);
      const more = document.createElement('button');
      more.innerHTML = '<i data-lucide="palette"></i> Weitere Farben…';
      const inp = document.createElement('input');
      inp.type = 'color';
      inp.style.cssText = 'position:absolute;opacity:0;pointer-events:none';
      inp.onchange = () => { closeMenu(); onPick(inp.value); };
      more.onclick = () => inp.click();
      m.appendChild(more);
      m.appendChild(inp);
    });
  }

  // ---------- Dokument-Zustand ----------
  const doc = { name: '', path: null, format: null, dirty: false };
  let cfg = null;

  function updateTitle() {
    const inp = $('#docName');
    if (inp && document.activeElement !== inp) inp.value = doc.name;
    const d = $('#dirtyMark');
    if (d) d.textContent = doc.dirty ? '•' : '';
    document.title = `${doc.dirty ? '• ' : ''}${doc.name || 'Unbenannt'} – ${APP_NAMES[cfg ? cfg.app : 'home']}`;
  }
  function setDirty(v = true) {
    if (doc.dirty === v) return;
    doc.dirty = v;
    if (native) native.setDirty(v);
    updateTitle();
  }
  function setDoc(name, path, format) {
    doc.name = stripExt(name) || doc.name;
    doc.path = path || null;
    doc.format = format || null;
    setDirty(false);
    updateTitle();
  }

  // ---------- Dateien ----------
  function browserPickFile(accept) {
    return new Promise(resolve => {
      const inp = document.createElement('input');
      inp.type = 'file';
      inp.accept = accept;
      inp.onchange = async () => {
        const f = inp.files[0];
        if (!f) return resolve(null);
        resolve({ name: f.name, path: null, data: new Uint8Array(await f.arrayBuffer()) });
      };
      inp.click();
    });
  }
  function browserDownload(bytes, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([bytes]));
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }

  async function openFile() {
    const exts = cfg.openFormats.flatMap(f => f.extensions);
    let file;
    if (native) {
      file = await native.openFile([
        { name: 'Alle unterstützten Dateien', extensions: exts },
        ...cfg.openFormats,
        { name: 'Alle Office-Dateien', extensions: Object.keys(EXT_APP) }
      ]);
    } else {
      file = await browserPickFile(exts.map(e => '.' + e).join(','));
    }
    if (!file) return;
    await loadFile(file);
  }

  async function loadFile(file) {
    const e = ext(file.name);
    if (EXT_APP[e] && EXT_APP[e] !== cfg.app) {
      if (native && file.path) { native.openPath(file.path); return; }
      toast('Diese Datei gehört zu einem anderen Programm.');
      return;
    }
    if (doc.dirty && !(await confirmBox('Dokument ersetzen?', 'Ungespeicherte Änderungen gehen verloren. Trotzdem öffnen?'))) {
      if (native && file.path) native.openWindow(cfg.app, file.path);
      return;
    }
    try {
      document.body.style.cursor = 'progress';
      await cfg.load(toBytes(file.data), e, file.name);
      setDoc(file.name, file.path, e);
      closeBackstage();
      toast(`„${file.name}“ geöffnet`);
    } catch (err) {
      console.error(err);
      alertBox('Datei konnte nicht geöffnet werden', String(err && err.message || err));
    } finally {
      document.body.style.cursor = '';
    }
  }

  async function writeAs(format, filePath) {
    const bytes = toBytes(await cfg.save(format));
    if (native) {
      await native.writeFile(filePath, bytes);
    } else {
      browserDownload(bytes, `${doc.name || cfg.defaultName}.${format}`);
    }
  }

  let saving = false;
  async function save(asNew = false) {
    if (saving) return false;
    saving = true;
    try {
      const canOverwrite = doc.path && cfg.saveFormats.some(f => f.extensions.includes(doc.format));
      let target = canOverwrite && !asNew ? doc.path : null;
      let format = canOverwrite && !asNew ? doc.format : cfg.saveFormats[0].extensions[0];
      if (!target && native) {
        target = await native.chooseSave({ suggestedName: `${doc.name || cfg.defaultName}.${format}`, filters: cfg.saveFormats });
        if (!target) return false;
        format = ext(target) || format;
        if (!cfg.saveFormats.some(f => f.extensions.includes(format))) { target += '.' + cfg.saveFormats[0].extensions[0]; format = cfg.saveFormats[0].extensions[0]; }
      }
      if (!native && asNew) {
        const v = await dialog({ title: 'Speichern unter', fields: [
          { name: 'name', label: 'Dateiname', value: doc.name || cfg.defaultName },
          { name: 'format', label: 'Format', type: 'select', value: format,
            options: cfg.saveFormats.map(f => ({ value: f.extensions[0], label: `${f.name} (.${f.extensions[0]})` })) }] });
        if (!v) return false;
        doc.name = v.name; format = v.format;
      }
      document.body.style.cursor = 'progress';
      await writeAs(format, target);
      const nm = target ? target.split(/[\\/]/).pop() : `${doc.name || cfg.defaultName}.${format}`;
      setDoc(nm, target, format);
      closeBackstage();
      toast(`Gespeichert: ${nm}`);
      return true;
    } catch (err) {
      console.error(err);
      alertBox('Speichern fehlgeschlagen', String(err && err.message || err));
      return false;
    } finally {
      saving = false;
      document.body.style.cursor = '';
    }
  }

  async function exportFormat(format, label) {
    try {
      let target = null;
      if (native) {
        target = await native.chooseSave({ suggestedName: `${doc.name || cfg.defaultName}.${format}`, filters: [{ name: label, extensions: [format] }] });
        if (!target) return;
      }
      await writeAs(format, target);
      closeBackstage();
      toast(`Exportiert als ${label}`);
    } catch (err) {
      console.error(err);
      alertBox('Export fehlgeschlagen', String(err && err.message || err));
    }
  }

  async function exportPDF() {
    closeBackstage();
    const opts = cfg.pdf || {};
    if (opts.before) await opts.before();
    try {
      if (native) {
        const target = await native.chooseSave({ suggestedName: `${doc.name || cfg.defaultName}.pdf`, filters: [{ name: 'PDF', extensions: ['pdf'] }] });
        if (target) { await native.exportPDF({ filePath: target, pageSize: typeof opts.pageSize === 'function' ? opts.pageSize() : opts.pageSize, landscape: typeof opts.landscape === 'function' ? opts.landscape() : opts.landscape, margins: opts.margins ? opts.margins() : undefined }); toast('PDF erstellt'); }
      } else {
        window.print();
      }
    } finally {
      if (opts.after) await opts.after();
    }
  }

  async function printDoc() {
    closeBackstage();
    const opts = cfg.pdf || {};
    if (opts.before) await opts.before();
    try {
      if (native) await native.print(); else window.print();
    } finally {
      if (opts.after) setTimeout(opts.after, 500);
    }
  }

  function newDoc() {
    closeBackstage();
    if (native) native.openWindow(cfg.app);
    else window.open(location.pathname, '_blank');
  }

  // ---------- Backstage ----------
  function closeBackstage() { const b = $('#backstage'); if (b) b.classList.remove('open'); }
  async function openBackstage(section = 'home') {
    const b = $('#backstage');
    b.classList.add('open');
    const sec = $('section', b);
    if (section === 'home') {
      let recent = [];
      if (native) recent = (await native.recent()).filter(r => r.page === cfg.app);
      sec.innerHTML = `<h1>Willkommen</h1>
        <h2>Neu</h2>
        <div class="template-row">${(cfg.templates || []).map((t, i) => `
          <button class="template" data-tpl="${i}"><div class="thumb">${t.thumb || `<i data-lucide="${t.icon || 'file-plus'}" style="width:40px;height:40px"></i>`}</div>${esc(t.name)}</button>`).join('')}
        </div>
        <h2>Zuletzt verwendet</h2>
        <div class="recent-list">${recent.length ? recent.map(r => `
          <button data-path="${esc(r.path)}"><i data-lucide="${cfg.icon}"></i><div><div>${esc(r.name)}</div><div class="path">${esc(r.path)}</div></div></button>`).join('')
          : '<div style="color:var(--text-2);padding:8px 12px">Noch keine Dateien geöffnet.</div>'}</div>`;
      $$('[data-tpl]', sec).forEach(btn => btn.onclick = async () => {
        const t = cfg.templates[+btn.dataset.tpl];
        if (doc.dirty && !(await confirmBox('Neues Dokument', 'Ungespeicherte Änderungen gehen verloren. Fortfahren?'))) return;
        await t.create();
        setDoc(t.docName || cfg.defaultName, null, null);
        closeBackstage();
      });
      $$('[data-path]', sec).forEach(btn => btn.onclick = async () => {
        try { await loadFile(await native.readFile(btn.dataset.path)); } catch (e) { alertBox('Fehler', String(e.message || e)); }
      });
    } else if (section === 'export') {
      sec.innerHTML = `<h1>Exportieren</h1><div class="recent-list">
        <button data-exp="pdf"><i data-lucide="file-down"></i><div><div>Als PDF speichern</div><div class="path">Zum Teilen und Drucken – sieht überall gleich aus</div></div></button>
        ${cfg.saveFormats.concat(cfg.exportFormats || []).map(f => `
        <button data-exp="${f.extensions[0]}" data-label="${esc(f.name)}"><i data-lucide="file-output"></i><div><div>${esc(f.name)} (.${f.extensions[0]})</div><div class="path">${esc(f.hint || '')}</div></div></button>`).join('')}
      </div>`;
      $$('[data-exp]', sec).forEach(btn => btn.onclick = () =>
        btn.dataset.exp === 'pdf' ? exportPDF() : exportFormat(btn.dataset.exp, btn.dataset.label));
    } else if (section === 'info') {
      sec.innerHTML = `<h1>Info</h1>
        <p><b>${esc(doc.name || 'Unbenannt')}</b></p>
        <p style="color:var(--text-2)">${esc(doc.path || 'Noch nicht gespeichert')}</p>
        ${cfg.info ? cfg.info() : ''}
        <h2>Über REX Office</h2>
        <p style="color:var(--text-2);max-width:620px">REX Office 1.0 – Text, Tabellen und Präsentationen.
        Kompatibel mit Microsoft Word (.docx), Excel (.xlsx) und PowerPoint (.pptx).</p>
        <h2>Tastenkürzel</h2>
        <table style="border-collapse:collapse;color:var(--text-2)">
          <tr><td style="padding:3px 18px 3px 0">Strg+N</td><td>Neu</td></tr>
          <tr><td style="padding:3px 18px 3px 0">Strg+O</td><td>Öffnen</td></tr>
          <tr><td style="padding:3px 18px 3px 0">Strg+S</td><td>Speichern</td></tr>
          <tr><td style="padding:3px 18px 3px 0">Strg+Umschalt+S</td><td>Speichern unter</td></tr>
          <tr><td style="padding:3px 18px 3px 0">Strg+P</td><td>Drucken</td></tr>
          <tr><td style="padding:3px 18px 3px 0">Strg+Z / Strg+Y</td><td>Rückgängig / Wiederholen</td></tr>
          ${(cfg.shortcutsHelp || []).map(([k, v]) => `<tr><td style="padding:3px 18px 3px 0">${esc(k)}</td><td>${esc(v)}</td></tr>`).join('')}
        </table>`;
    }
    $$('nav [data-sec]', b).forEach(x => x.classList.toggle('on', x.dataset.sec === section));
    icons();
  }

  function buildChrome() {
    // Titelleiste
    const tb = $('.titlebar');
    if (tb) {
      tb.innerHTML = `
        <div class="brand"><div class="logo">${cfg.letter}</div>${APP_NAMES[cfg.app]}</div>
        <div class="qa">
          <button title="Speichern (Strg+S)" data-qa="save"><i data-lucide="save"></i></button>
          <button title="Rückgängig (Strg+Z)" data-qa="undo"><i data-lucide="undo-2"></i></button>
          <button title="Wiederholen (Strg+Y)" data-qa="redo"><i data-lucide="redo-2"></i></button>
        </div>
        <div class="doc-title"><input id="docName" spellcheck="false" title="Dokumentname"><span id="dirtyMark" class="dirty"></span></div>
        <div class="right">
          <button title="Startseite" data-qa="home"><i data-lucide="house"></i></button>
          <button title="Hell/Dunkel" id="themeBtn"><i data-lucide="moon"></i></button>
        </div>`;
      tb.addEventListener('click', e => {
        const b = e.target.closest('[data-qa]');
        if (!b) return;
        const a = b.dataset.qa;
        if (a === 'save') save();
        if (a === 'undo') cfg.undo && cfg.undo();
        if (a === 'redo') cfg.redo && cfg.redo();
        if (a === 'home') { if (native) native.openWindow('home'); else window.open('index.html', '_blank'); }
      });
      $('#themeBtn').onclick = () => {
        const t = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark';
        store('theme', t);
        applyTheme(t);
        if (cfg.onTheme) cfg.onTheme(t);
      };
      const dn = $('#docName');
      dn.addEventListener('change', () => { doc.name = dn.value.trim() || doc.name; updateTitle(); });
      dn.addEventListener('keydown', e => { if (e.key === 'Enter') dn.blur(); e.stopPropagation(); });
    }

    // Registerkarten
    const tabs = $('.tabs');
    if (tabs) {
      tabs.insertAdjacentHTML('afterbegin', '<button class="file-tab" data-tab="__file">Datei</button>');
      tabs.addEventListener('click', e => {
        const b = e.target.closest('button[data-tab]');
        if (!b) return;
        if (b.dataset.tab === '__file') { openBackstage('home'); return; }
        selectTab(b.dataset.tab);
      });
      const first = $('button[data-tab]:not(.file-tab)', tabs);
      if (first) selectTab(first.dataset.tab);
    }

    // Backstage
    const bs = document.createElement('div');
    bs.id = 'backstage';
    bs.className = 'backstage';
    bs.innerHTML = `<nav>
        <button class="back" data-act="close"><i data-lucide="arrow-left"></i> Zurück</button>
        <button data-sec="home"><i data-lucide="house"></i> Start</button>
        <button data-act="new"><i data-lucide="file-plus"></i> Neu (neues Fenster)</button>
        <button data-act="open"><i data-lucide="folder-open"></i> Öffnen</button>
        <hr>
        <button data-act="save"><i data-lucide="save"></i> Speichern</button>
        <button data-act="saveas"><i data-lucide="save-all"></i> Speichern unter</button>
        <button data-sec="export"><i data-lucide="file-output"></i> Exportieren</button>
        <button data-act="print"><i data-lucide="printer"></i> Drucken</button>
        <hr>
        <button data-sec="info"><i data-lucide="info"></i> Info</button>
      </nav><section></section>`;
    document.body.appendChild(bs);
    bs.addEventListener('click', e => {
      const b = e.target.closest('nav button');
      if (!b) return;
      if (b.dataset.sec) return openBackstage(b.dataset.sec);
      const a = b.dataset.act;
      if (a === 'close') closeBackstage();
      if (a === 'new') newDoc();
      if (a === 'open') openFile();
      if (a === 'save') save();
      if (a === 'saveas') save(true);
      if (a === 'print') printDoc();
    });
  }

  function selectTab(name) {
    $$('.tabs button[data-tab]').forEach(b => b.classList.toggle('active', b.dataset.tab === name));
    $$('.ribbon .pane').forEach(p => p.classList.toggle('active', p.dataset.pane === name));
    if (cfg && cfg.onTab) cfg.onTab(name);
  }

  // ---------- Initialisierung ----------
  async function init(c) {
    cfg = c;
    document.body.classList.add('app-' + c.app);
    doc.name = c.defaultName;
    buildChrome();
    updateTitle();
    icons();

    document.addEventListener('keydown', e => {
      if (!(e.ctrlKey || e.metaKey)) {
        if (e.key === 'Escape' && $('#backstage.open')) closeBackstage();
        return;
      }
      const k = e.key.toLowerCase();
      if (k === 's') { e.preventDefault(); save(e.shiftKey); }
      else if (k === 'o') { e.preventDefault(); openFile(); }
      else if (k === 'n') { e.preventDefault(); newDoc(); }
      else if (k === 'p') { e.preventDefault(); printDoc(); }
    }, true);

    window.addEventListener('beforeunload', e => {
      if (!native && doc.dirty) { e.preventDefault(); e.returnValue = ''; }
    });

    if (native) {
      native.onSaveAndClose(async () => { if (await save()) native.closeNow(); });
      const startup = await native.startupFile();
      if (startup) {
        try { await loadFile(await native.readFile(startup)); } catch (e) { alertBox('Fehler', String(e.message || e)); }
      }
    }
  }

  window.Rex = {
    init, native, doc, setDirty, setDoc, save, openFile, loadFile, exportPDF, printDoc, selectTab,
    dialog, confirm: confirmBox, alert: alertBox, toast, menu, closeMenu, colorMenu, icons, PALETTE,
    util: { $, $$, ext, stripExt, esc, toBytes, blobToBytes, bytesToDataURL, dataURLToBytes, mimeFromDataURL,
      readImageFile, imageSize, normalizeImage, store, filePath }
  };
})();
