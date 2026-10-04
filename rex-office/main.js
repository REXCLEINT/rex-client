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
      spellcheck: true
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
    if (st && st.dirty) {
      const choice = dialog.showMessageBoxSync(win, {
        type: 'question',
        buttons: ['Speichern', 'Nicht speichern', 'Abbrechen'],
        defaultId: 0,
        cancelId: 2,
        title: 'REX Office',
        message: 'Möchtest du die Änderungen speichern?',
        detail: 'Wenn du nicht speicherst, gehen deine Änderungen verloren.'
      });
      if (choice === 2) { e.preventDefault(); return; }
      if (choice === 0) {
        e.preventDefault();
        win.webContents.send('app:save-and-close');
        return;
      }
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
ipcMain.on('app:dirty', (e, dirty) => {
  const st = windowState.get(e.sender.id);
  if (st) st.dirty = !!dirty;
});
ipcMain.on('app:close-now', (e) => {
  const st = windowState.get(e.sender.id);
  if (st) st.dirty = false;
  const w = senderWin(e);
  if (w) w.close();
});
