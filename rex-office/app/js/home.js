/* REX Office – Startseite */
(function () {
  'use strict';
  const { $, esc } = Rex.util;
  const native = Rex.native;
  const PAGES = { writer: 'writer.html', calc: 'calc.html', present: 'present.html' };
  const COLORS = { writer: '#2b579a', calc: '#217346', present: '#c43e1c' };
  const LETTERS = { writer: 'T', calc: 'X', present: 'P' };
  const EXT_APP = { docx: 'writer', txt: 'writer', html: 'writer', htm: 'writer', md: 'writer', xlsx: 'calc', csv: 'calc', pptx: 'present' };

  const h = new Date().getHours();
  $('#greet').textContent = h < 5 ? 'Gute Nacht' : h < 11 ? 'Guten Morgen' : h < 18 ? 'Guten Tag' : 'Guten Abend';

  function openApp(app, filePath) {
    if (native) native.openWindow(app, filePath || null);
    else location.href = PAGES[app];
  }
  document.querySelectorAll('[data-app]').forEach(b => b.addEventListener('click', () => openApp(b.dataset.app)));

  $('#openBtn').addEventListener('click', async () => {
    if (!native) { Rex.toast('Dateien öffnen geht im Programm über Datei → Öffnen.'); return; }
    const f = await native.openFile([
      { name: 'Alle Office-Dateien', extensions: Object.keys(EXT_APP) },
      { name: 'Word', extensions: ['docx'] }, { name: 'Excel', extensions: ['xlsx', 'csv'] }, { name: 'PowerPoint', extensions: ['pptx'] }
    ]);
    if (f) { native.openPath(f.path); loadRecent(); }
  });

  async function loadRecent() {
    if (!native) return;
    const list = await native.recent();
    const box = $('#recent');
    if (!list.length) return;
    box.innerHTML = list.map(r => `<button data-path="${esc(r.path)}">
      <span class="badge" style="background:${COLORS[r.page] || '#777'}">${LETTERS[r.page] || '?'}</span>
      <span><div>${esc(r.name)}</div><div class="path">${esc(r.path)}</div></span></button>`).join('');
    box.querySelectorAll('[data-path]').forEach(b => b.addEventListener('click', () => native.openPath(b.dataset.path)));
  }
  loadRecent();
  window.addEventListener('focus', loadRecent);

  // Theme
  const applyTheme = (t) => {
    document.documentElement.dataset.theme = t;
    $('#themeBtn').innerHTML = `<i data-lucide="${t === 'dark' ? 'sun' : 'moon'}"></i>`;
    Rex.icons();
  };
  applyTheme(Rex.util.store('theme') || 'light');
  $('#themeBtn').addEventListener('click', () => {
    const t = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark';
    Rex.util.store('theme', t);
    applyTheme(t);
  });

  // Dateien hineinziehen
  document.addEventListener('dragover', e => { e.preventDefault(); document.body.classList.add('dragging'); });
  document.addEventListener('dragleave', e => { if (!e.relatedTarget) document.body.classList.remove('dragging'); });
  document.addEventListener('drop', e => {
    e.preventDefault();
    document.body.classList.remove('dragging');
    const f = e.dataTransfer.files[0];
    if (!f) return;
    const ext = (f.name.split('.').pop() || '').toLowerCase();
    if (!EXT_APP[ext]) { Rex.toast('Dieses Dateiformat wird nicht unterstützt.'); return; }
    const p = native && native.pathForFile ? native.pathForFile(f) : f.path;
    if (native && p) native.openPath(p);
    else openApp(EXT_APP[ext]);
  });
  Rex.icons();
})();
