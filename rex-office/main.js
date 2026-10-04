// REX Office – Electron-Hauptprozess
const { app, BrowserWindow, ipcMain, dialog, shell, Menu } = require('electron');
const path = require('path');
const fs = require('fs');

const APP_DIR = path.join(__dirname, 'app');
const PAGES = { writer: 'writer.html', calc: 'calc.html', present: 'present.html', home: 'index.html' };
const EXT_TO_PAGE = { docx: 'writer', txt: 'writer', html: 'writer', htm: 'writer', md: 'writer',
  xlsx: 'calc', csv: 'calc', pptx: 'present' };

const windowState = new Map(); // webContents.id -> { dirty, startupFile }

function recentFile() { return path.join(app.getPath('userData'), 'recent.json'); }
function readRecent() {
  try { return JSON.parse(fs.readFileSync(recentFile(), 'utf8')); } catch { return []; }
}
function addRecent(filePath) {
  if (!filePath) return;
  const list = readRecent().filter(p => p !== filePath);
  list.unshift(filePath);
  try { fs.writeFileSync(recentFile(), JSON.stringify(list.slice(0, 15))); } catch { /* ignore */ }
  if (process.platform === 'win32') app.addRecentDocument(filePath);
}

// ---------- Einstellungen (Name, Sprache, Profilbild) ----------
function settingsFile() { return path.join(app.getPath('userData'), 'settings.json'); }
function readSettings() {
  try { return JSON.parse(fs.readFileSync(settingsFile(), 'utf8')); } catch { return {}; }
}
function writeSettings(obj) {
  const merged = { ...readSettings(), ...obj };
  fs.writeFileSync(settingsFile(), JSON.stringify(merged));
  for (const w of BrowserWindow.getAllWindows()) w.webContents.send('settings:changed', merged);
  return merged;
}

function pageForFile(filePath) {
  const ext = path.extname(filePath || '').slice(1).toLowerCase();
  return EXT_TO_PAGE[ext] || null;
}

function createWindow(page = 'home', startupFile = null) {
  const win = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 900,
    minHeight: 600,
    title: 'REX Office',
    icon: path.join(__dirname, 'build', 'icon.png'),
    backgroundColor: '#f3f3f3',
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      spellcheck: true,
      plugins: true
    }
  });
  windowState.set(win.webContents.id, { dirty: false, startupFile });
  win.webContents.session.setSpellCheckerLanguages(['de-DE', 'en-US']);
  win.loadFile(path.join(APP_DIR, PAGES[page] || PAGES.home));

  // Externe Links im Standardbrowser oeffnen
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https?:/.test(url)) shell.openExternal(url);
    return { action: 'deny' };
  });

  // Rechtschreib-Vorschlaege im Kontextmenue
  win.webContents.on('context-menu', (_e, params) => {
    const items = [];
    for (const s of params.dictionarySuggestions || []) {
      items.push({ label: s, click: () => win.webContents.replaceMisspelling(s) });
    }
    if (params.misspelledWord) {
      items.push({ label: 'Zum Wörterbuch hinzufügen', click: () =>
        win.webContents.session.addWordToSpellCheckerDictionary(params.misspelledWord) });
      items.push({ type: 'separator' });
    }
    if (params.isEditable) {
      items.push({ label: 'Ausschneiden', role: 'cut' }, { label: 'Kopieren', role: 'copy' },
        { label: 'Einfügen', role: 'paste' }, { type: 'separator' }, { label: 'Alles auswählen', role: 'selectAll' });
    } else if (params.selectionText) {
      items.push({ label: 'Kopieren', role: 'copy' });
    }
    if (items.length) Menu.buildFromTemplate(items).popup({ window: win });
  });

  const id = win.webContents.id;
  win.on('close', (e) => {
    const st = windowState.get(id);
    if (st && st.dirty && !st.forceClose) {
      e.preventDefault();
      win.webContents.send('app:ask-close');
    }
  });
  win.on('closed', () => windowState.delete(id));
  return win;
}

function fileFromArgv(argv) {
  return argv.slice(1).find(a => !a.startsWith('-') && pageForFile(a) && fs.existsSync(a)) || null;
}

function openPath(filePath) {
  const page = pageForFile(filePath);
  if (!page) return;
  createWindow(page, filePath);
}

const gotLock = app.requestSingleInstanceLock();
if (!gotLock) {
  app.quit();
} else {
  app.on('second-instance', (_e, argv) => {
    const f = fileFromArgv(argv);
    if (f) openPath(f);
    else {
      const w = BrowserWindow.getAllWindows()[0];
      if (w) { if (w.isMinimized()) w.restore(); w.focus(); }
    }
  });

  app.whenReady().then(() => {
    Menu.setApplicationMenu(null);
    const f = fileFromArgv(process.argv);
    if (f) openPath(f); else createWindow('home');
    app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow('home'); });
  });
  app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
}

// ---------- IPC ----------
const senderWin = (e) => BrowserWindow.fromWebContents(e.sender);

ipcMain.handle('file:open', async (e, { filters }) => {
  const res = await dialog.showOpenDialog(senderWin(e), { properties: ['openFile'], filters });
  if (res.canceled || !res.filePaths.length) return null;
  const p = res.filePaths[0];
  addRecent(p);
  return { path: p, name: path.basename(p), data: fs.readFileSync(p) };
});

ipcMain.handle('file:read', async (_e, p) => {
  addRecent(p);
  return { path: p, name: path.basename(p), data: fs.readFileSync(p) };
});

ipcMain.handle('file:choose-save', async (e, { suggestedName, filters }) => {
  const res = await dialog.showSaveDialog(senderWin(e), {
    defaultPath: path.join(app.getPath('documents'), suggestedName || 'Dokument'),
    filters
  });
  if (res.canceled || !res.filePath) return null;
  return res.filePath;
});

ipcMain.handle('file:write', async (_e, { filePath, data }) => {
  fs.writeFileSync(filePath, Buffer.from(data));
  addRecent(filePath);
  return { path: filePath, name: path.basename(filePath) };
});

ipcMain.handle('file:pdf', async (e, { filePath, pageSize, landscape, margins }) => {
  const win = senderWin(e);
  const pdf = await win.webContents.printToPDF({
    printBackground: true,
    pageSize: pageSize || 'A4',
    landscape: !!landscape,
    margins: margins || { top: 0, bottom: 0, left: 0, right: 0 }
  });
  fs.writeFileSync(filePath, pdf);
  shell.openPath(filePath);
  return { path: filePath };
});

ipcMain.handle('app:print', (e) => { senderWin(e).webContents.print({ printBackground: true }); });
ipcMain.handle('print:printers', async (e) => {
  try { return await senderWin(e).webContents.getPrintersAsync(); } catch { return []; }
});
ipcMain.handle('print:preview', async (e, { pageSize, landscape, margins }) => {
  return senderWin(e).webContents.printToPDF({
    printBackground: true, pageSize: pageSize || 'A4', landscape: !!landscape,
    margins: margins || { top: 0, bottom: 0, left: 0, right: 0 }
  });
});
ipcMain.handle('print:run', (e, opts) => new Promise((resolve) => {
  const o = {
    silent: true, printBackground: true, deviceName: opts.deviceName || '', copies: Math.max(1, opts.copies || 1),
    collate: true, color: opts.color !== false, landscape: !!opts.landscape,
    duplexMode: opts.duplex ? 'longEdge' : 'simplex'
  };
  if (opts.pageRanges && opts.pageRanges.length) o.pageRanges = opts.pageRanges;
  if (opts.pageSize) o.pageSize = typeof opts.pageSize === 'string' ? opts.pageSize
    : { width: Math.round(opts.pageSize.width * 25400), height: Math.round(opts.pageSize.height * 25400) };
  if (opts.margins) o.margins = { marginType: 'custom', top: opts.margins.top * 72, bottom: opts.margins.bottom * 72, left: opts.margins.left * 72, right: opts.margins.right * 72 };
  senderWin(e).webContents.print(o, (ok, reason) => resolve({ ok, reason }));
}));

// ---------- Automatische Updates ueber GitHub-Releases (Tags "office-v1.2.3") ----------
const UPDATE_REPO = 'REXCLEINT/rex-client';
const isNewer = (a, b) => {
  const pa = a.split('.').map(Number), pb = b.split('.').map(Number);
  for (let i = 0; i < 3; i++) { if ((pa[i] || 0) > (pb[i] || 0)) return true; if ((pa[i] || 0) < (pb[i] || 0)) return false; }
  return false;
};
async function latestOfficeRelease() {
  const headers = { 'User-Agent': 'REX-Office', Accept: 'application/vnd.github+json' };
  try {
    const res = await fetch(`https://api.github.com/repos/${UPDATE_REPO}/releases?per_page=30`, { headers });
    if (!res.ok) throw new Error('API ' + res.status);
    const list = await res.json();
    const rel = list.filter(r => !r.draft && !r.prerelease && /^office-v\d+\.\d+\.\d+$/.test(r.tag_name))
      .sort((a, b) => (isNewer(a.tag_name.slice(8), b.tag_name.slice(8)) ? -1 : 1))[0];
    if (!rel) return null;
    const asset = (rel.assets || []).find(a => /^REX-Office-Setup-.*\.exe$/.test(a.name));
    return { version: rel.tag_name.slice(8), notes: rel.body || '', page: rel.html_url,
      url: asset ? asset.browser_download_url : null, size: asset ? asset.size : 0 };
  } catch (apiErr) {
    // Ausweichweg ohne API-Limit: Release-Feed der GitHub-Webseite
    const res = await fetch(`https://github.com/${UPDATE_REPO}/releases.atom`, { headers: { 'User-Agent': 'REX-Office' } });
    if (!res.ok) throw new Error('GitHub antwortet nicht (' + res.status + ')');
    const xml = await res.text();
    const tags = [...xml.matchAll(/releases\/tag\/office-v(\d+\.\d+\.\d+)/g)].map(m => m[1]);
    if (!tags.length) return null;
    const version = tags.sort((a, b) => (isNewer(a, b) ? -1 : 1))[0];
    return { version, notes: '', page: `https://github.com/${UPDATE_REPO}/releases/tag/office-v${version}`,
      url: `https://github.com/${UPDATE_REPO}/releases/download/office-v${version}/REX-Office-Setup-${version}.exe`, size: 0 };
  }
}
async function findUpdate() {
  const rel = await latestOfficeRelease();
  const base = { current: app.getVersion(), portable: !!process.env.PORTABLE_EXECUTABLE_FILE, packaged: app.isPackaged };
  if (!rel) return { ...base, available: false };
  return { ...base, ...rel, available: isNewer(rel.version, app.getVersion()) };
}
ipcMain.handle('update:check', async () => {
  try { return await findUpdate(); } catch (err) { return { error: String(err.message || err), current: app.getVersion() }; }
});
ipcMain.handle('update:install', async (e) => {
  const info = await findUpdate();
  if (!info.available) return { ok: false, reason: 'Kein Update verfügbar' };
  if (!info.url || info.portable || process.platform !== 'win32') { shell.openExternal(info.page); return { ok: true, opened: true }; }
  const target = path.join(app.getPath('temp'), `REX-Office-Setup-${info.version}.exe`);
  const res = await fetch(info.url, { headers: { 'User-Agent': 'REX-Office' } });
  if (!res.ok || !res.body) throw new Error('Download fehlgeschlagen (' + res.status + ')');
  const total = Number(res.headers.get('content-length')) || info.size || 0;
  const out = fs.createWriteStream(target);
  let done = 0, lastSent = 0;
  for await (const chunk of res.body) {
    out.write(chunk);
    done += chunk.length;
    if (Date.now() - lastSent > 150) { lastSent = Date.now(); e.sender.send('update:progress', { done, total }); }
  }
  await new Promise(r => out.end(r));
  e.sender.send('update:progress', { done: total, total });
  // Installer still im Hintergrund ausfuehren und REX Office danach neu starten
  const { spawn } = require('child_process');
  spawn(target, ['/S', '--updated', '--force-run'], { detached: true, stdio: 'ignore' }).unref();
  setTimeout(() => {
    for (const w of BrowserWindow.getAllWindows()) { const st = windowState.get(w.webContents.id); if (st) st.forceClose = true; }
    app.quit();
  }, 800);
  return { ok: true };
});
ipcMain.handle('app:version', () => app.getVersion());
ipcMain.handle('app:recent', () => readRecent().filter(p => fs.existsSync(p)).map(p => ({ path: p, name: path.basename(p), page: pageForFile(p) })));
ipcMain.handle('app:startup-file', (e) => {
  const st = windowState.get(e.sender.id);
  const f = st && st.startupFile;
  if (st) st.startupFile = null;
  return f;
});
ipcMain.handle('app:open-window', (_e, { page, filePath }) => { createWindow(page, filePath || null); });
ipcMain.handle('app:open-path', (_e, filePath) => {
  const page = pageForFile(filePath);
  if (page) createWindow(page, filePath);
});
ipcMain.on('settings:get-sync', (e) => { e.returnValue = readSettings(); });
ipcMain.handle('settings:set', (_e, obj) => writeSettings(obj));
ipcMain.on('app:dirty', (e, dirty) => {
  const st = windowState.get(e.sender.id);
  if (st) st.dirty = !!dirty;
});
ipcMain.on('app:close-now', (e) => {
  const st = windowState.get(e.sender.id);
  if (st) { st.dirty = false; st.forceClose = true; }
  const w = senderWin(e);
  if (w) w.close();
});
