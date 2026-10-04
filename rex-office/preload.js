// Sichere Bruecke zwischen Oberflaeche und Dateisystem
const { contextBridge, ipcRenderer, webUtils } = require('electron');

let settings = {};
try { settings = ipcRenderer.sendSync('settings:get-sync') || {}; } catch { settings = {}; }
const settingsListeners = [];
ipcRenderer.on('settings:changed', (_e, s) => { settings = s; settingsListeners.forEach(cb => cb(s)); });

contextBridge.exposeInMainWorld('rexNative', {
  getSettings: () => settings,
  setSettings: (obj) => ipcRenderer.invoke('settings:set', obj),
  onSettingsChanged: (cb) => settingsListeners.push(cb),
  openFile: (filters) => ipcRenderer.invoke('file:open', { filters }),
  readFile: (p) => ipcRenderer.invoke('file:read', p),
  chooseSave: (opts) => ipcRenderer.invoke('file:choose-save', opts),
  writeFile: (filePath, data) => ipcRenderer.invoke('file:write', { filePath, data }),
  exportPDF: (opts) => ipcRenderer.invoke('file:pdf', opts),
  print: () => ipcRenderer.invoke('app:print'),
  recent: () => ipcRenderer.invoke('app:recent'),
  startupFile: () => ipcRenderer.invoke('app:startup-file'),
  openWindow: (page, filePath) => ipcRenderer.invoke('app:open-window', { page, filePath }),
  openPath: (p) => ipcRenderer.invoke('app:open-path', p),
  setDirty: (d) => ipcRenderer.send('app:dirty', d),
  closeNow: () => ipcRenderer.send('app:close-now'),
  onSaveAndClose: (cb) => ipcRenderer.on('app:save-and-close', cb),
  pathForFile: (file) => { try { return webUtils.getPathForFile(file); } catch { return null; } }
});
