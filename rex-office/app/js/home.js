/* REX Office – Startseite */
(function () {
  'use strict';
  const { $, esc } = Rex.util;
  const native = Rex.native;
  const PAGES = { writer: 'writer.html', calc: 'calc.html', present: 'present.html' };
  const COLORS = { writer: '#2b579a', calc: '#217346', present: '#c43e1c' };
  const LETTERS = { writer: 'T', calc: 'X', present: 'P' };
  const EXT_APP = { docx: 'writer', txt: 'writer', html: 'writer', htm: 'writer', md: 'writer', xlsx: 'calc', csv: 'calc', pptx: 'present' };

  const t = window.RexI18n ? RexI18n.t : (x) => x;

  // ---------- Begruessung mit Name und Profilbild ----------
  function renderProfile() {
    const st = Rex.settings.get();
    const h = new Date().getHours();
    const base = t(h < 5 ? 'Gute Nacht' : h < 11 ? 'Guten Morgen' : h < 18 ? 'Guten Tag' : 'Guten Abend');
    $('#greet').textContent = st.name ? `${base}, ${st.name}!` : base;
    $('#greetAvatar').innerHTML = st.name || st.avatar ? Rex.avatarHTML(st, 58) : '';
    $('#greetAvatar').style.display = st.name || st.avatar ? '' : 'none';
    $('#profileBtn').innerHTML = `${Rex.avatarHTML(st, 30)}<span>${esc(st.name || t('Profil & Einstellungen'))}</span>`;
  }
  renderProfile();
  $('#profileBtn').addEventListener('click', () => openSetup(true));

  // ---------- Einrichtungs-Assistent ----------
  const TXT = {
    de: {
      langTitle: 'Willkommen bei REX Office!', langSub: 'Wähle zuerst deine Sprache.',
      nameTitle: 'Wie heißt du?', nameSub: 'Dein Name erscheint in der Begrüßung und als Autor deiner Dateien.', namePh: 'Dein Name',
      picTitle: 'Dein Profilbild', picSub: 'Wähle ein Bild aus oder nimm eine Farbe für deine Initialen.', picChoose: 'Bild auswählen', picRemove: 'Bild entfernen',
      themeTitle: 'Hell oder dunkel?', themeSub: 'Du kannst das später jederzeit oben rechts ändern.', light: 'Hell', lightSub: 'Klassisch und klar', dark: 'Dunkel', darkSub: 'Schont die Augen',
      back: 'Zurück', next: 'Weiter', finish: 'Los geht\'s!', save: 'Speichern', skip: 'Überspringen'
    },
    en: {
      langTitle: 'Welcome to REX Office!', langSub: 'First, choose your language.',
      nameTitle: 'What\'s your name?', nameSub: 'Your name appears in the greeting and as the author of your files.', namePh: 'Your name',
      picTitle: 'Your profile picture', picSub: 'Choose a picture or pick a color for your initials.', picChoose: 'Choose picture', picRemove: 'Remove picture',
      themeTitle: 'Light or dark?', themeSub: 'You can change this anytime in the top right corner.', light: 'Light', lightSub: 'Classic and clear', dark: 'Dark', darkSub: 'Easy on the eyes',
      back: 'Back', next: 'Next', finish: 'Let\'s go!', save: 'Save', skip: 'Skip'
    }
  };
  const AV_COLORS = ['#5b3fd1', '#2b579a', '#217346', '#c43e1c', '#e3008c', '#0097a7', '#f59f00', '#333333'];
  let draft = null, step = 0, editMode = false;
  const STEPS = ['lang', 'name', 'pic', 'theme'];

  function openSetup(edit) {
    const st = Rex.settings.get();
    editMode = !!edit;
    draft = {
      lang: st.lang || ((navigator.language || 'de').startsWith('en') ? 'en' : 'de'),
      name: st.name || '', avatar: st.avatar || null, avatarColor: st.avatarColor || AV_COLORS[0],
      theme: Rex.util.store('theme') || 'light'
    };
    step = 0;
    $('#setup').hidden = false;
    renderStep();
  }
  function renderStep() {
    const L = TXT[draft.lang];
    const body = $('#setupBody');
    const kind = STEPS[step];
    let html = '<div class="step">';
    if (kind === 'lang') {
      html += `<h2>${L.langTitle}</h2><p class="sub">${L.langSub}</p><div class="choices">
        <button class="choice ${draft.lang === 'de' ? 'on' : ''}" data-lang="de"><span class="badge" style="background:#d00">DE</span><span>Deutsch<small>Deutschland, Österreich, Schweiz</small></span></button>
        <button class="choice ${draft.lang === 'en' ? 'on' : ''}" data-lang="en"><span class="badge" style="background:#1f4e9c">EN</span><span>English<small>International</small></span></button></div>`;
    } else if (kind === 'name') {
      html += `<h2>${L.nameTitle}</h2><p class="sub">${L.nameSub}</p><input class="name" id="setupName" maxlength="40" placeholder="${L.namePh}" value="${esc(draft.name)}" autocomplete="off">`;
    } else if (kind === 'pic') {
      html += `<h2>${L.picTitle}</h2><p class="sub">${L.picSub}</p><div class="pic">
        <div class="pic-preview">${Rex.avatarHTML(draft, 110)}</div>
        <div class="pic-actions">
          <button class="sbtn" id="picChoose">${L.picChoose}</button>
          ${draft.avatar ? `<button class="sbtn" id="picRemove">${L.picRemove}</button>` : ''}
          <div class="colors">${AV_COLORS.map(c => `<button data-color="${c}" class="${!draft.avatar && draft.avatarColor === c ? 'on' : ''}" style="background:${c}"></button>`).join('')}</div>
        </div></div>`;
    } else {
      html += `<h2>${L.themeTitle}</h2><p class="sub">${L.themeSub}</p><div class="choices">
        <button class="choice ${draft.theme === 'light' ? 'on' : ''}" data-theme="light"><span class="badge" style="background:#f3f3f3;color:#333;border:1px solid #ccc">☀</span><span>${L.light}<small>${L.lightSub}</small></span></button>
        <button class="choice ${draft.theme === 'dark' ? 'on' : ''}" data-theme="dark"><span class="badge" style="background:#1f1f1f">☾</span><span>${L.dark}<small>${L.darkSub}</small></span></button></div>`;
    }
    body.innerHTML = html + '</div>';
    $('#setupDots').innerHTML = STEPS.map((_, i) => `<i class="${i === step ? 'on' : ''}"></i>`).join('');
    $('#setupBack').textContent = step === 0 ? (editMode ? (draft.lang === 'en' ? 'Cancel' : 'Abbrechen') : '') : L.back;
    $('#setupBack').style.visibility = step === 0 && !editMode ? 'hidden' : '';
    $('#setupNext').textContent = step === STEPS.length - 1 ? (editMode ? L.save : L.finish) : L.next;
    // Ereignisse
    body.querySelectorAll('[data-lang]').forEach(b => b.onclick = () => { draft.lang = b.dataset.lang; renderStep(); });
    body.querySelectorAll('[data-theme]').forEach(b => b.onclick = () => {
      draft.theme = b.dataset.theme;
      document.documentElement.dataset.theme = draft.theme;
      renderStep();
    });
    const nameInput = $('#setupName');
    if (nameInput) {
      nameInput.focus();
      nameInput.oninput = () => { draft.name = nameInput.value; };
      nameInput.onkeydown = (e) => { if (e.key === 'Enter') next(); };
    }
    body.querySelectorAll('[data-color]').forEach(b => b.onclick = () => { draft.avatarColor = b.dataset.color; draft.avatar = null; renderStep(); });
    if ($('#picRemove')) $('#picRemove').onclick = () => { draft.avatar = null; renderStep(); };
    if ($('#picChoose')) $('#picChoose').onclick = async () => {
      const src = await Rex.util.readImageFile();
      if (!src) return;
      draft.avatar = await squareImage(src, 256);
      renderStep();
    };
  }
  // Bild quadratisch zuschneiden und verkleinern
  function squareImage(src, size) {
    return new Promise(resolve => {
      const img = new Image();
      img.onload = () => {
        const s = Math.min(img.naturalWidth, img.naturalHeight);
        const c = document.createElement('canvas');
        c.width = c.height = size;
        const ctx = c.getContext('2d');
        ctx.imageSmoothingQuality = 'high';
        ctx.drawImage(img, (img.naturalWidth - s) / 2, (img.naturalHeight - s) / 2, s, s, 0, 0, size, size);
        resolve(c.toDataURL('image/jpeg', 0.9));
      };
      img.onerror = () => resolve(null);
      img.src = src;
    });
  }
  async function next() {
    if (step < STEPS.length - 1) { step++; renderStep(); return; }
    const before = Rex.settings.get();
    await Rex.settings.set({ setupDone: true, lang: draft.lang, name: draft.name.trim(), avatar: draft.avatar, avatarColor: draft.avatarColor });
    Rex.util.store('theme', draft.theme);
    $('#setup').hidden = true;
    if ((before.lang || 'de') !== draft.lang) { location.reload(); return; }
    applyTheme(draft.theme);
    renderProfile();
  }
  $('#setupNext').addEventListener('click', next);
  $('#setupBack').addEventListener('click', () => {
    if (step > 0) { step--; renderStep(); return; }
    if (editMode) { $('#setup').hidden = true; applyTheme(Rex.util.store('theme') || 'light'); }
  });
  if (!Rex.settings.get().setupDone) openSetup(false);

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
  function applyTheme(th) {
    document.documentElement.dataset.theme = th;
    $('#themeBtn').innerHTML = `<i data-lucide="${th === 'dark' ? 'sun' : 'moon'}"></i>`;
    Rex.icons();
  }
  applyTheme(Rex.util.store('theme') || 'light');
  $('#themeBtn').addEventListener('click', () => {
    const th = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark';
    Rex.util.store('theme', th);
    applyTheme(th);
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
