import { contextBridge, ipcRenderer, webUtils, type IpcRendererEvent } from 'electron';
import type { AfficeBridge, ContextMenuParams } from '../src/shared/types';

function listen<T extends unknown[]>(channel: string, cb: (...args: T) => void): () => void {
  const handler = (_e: IpcRendererEvent, ...args: unknown[]) => cb(...(args as T));
  ipcRenderer.on(channel, handler);
  return () => ipcRenderer.removeListener(channel, handler);
}

const bridge: AfficeBridge = {
  isElectron: true,
  platform: process.platform,
  window: {
    minimize: () => ipcRenderer.send('win:minimize'),
    toggleMaximize: () => ipcRenderer.send('win:toggle-maximize'),
    close: () => ipcRenderer.send('win:close'),
    forceClose: () => ipcRenderer.send('win:force-close'),
    isMaximized: () => ipcRenderer.invoke('win:is-maximized'),
    setFullScreen: (on) => ipcRenderer.send('win:set-fullscreen', on),
    isFullScreen: () => ipcRenderer.invoke('win:is-fullscreen'),
    setTitle: (title) => ipcRenderer.send('win:set-title', title),
    setTitleBarColors: (color, symbolColor) => ipcRenderer.send('win:titlebar-colors', color, symbolColor),
    onStateChange: (cb) => listen('win:state', cb),
    onCloseRequest: (cb) => {
      ipcRenderer.send('win:has-close-handler');
      return listen('win:close-request', cb);
    },
  },
  files: {
    openDialog: (opts) => ipcRenderer.invoke('files:open-dialog', opts),
    read: (p) => ipcRenderer.invoke('files:read', p),
    saveDialog: (opts) => ipcRenderer.invoke('files:save-dialog', opts),
    write: (p, data) => ipcRenderer.invoke('files:write', p, data),
    exists: (p) => ipcRenderer.invoke('files:exists', p),
    showInFolder: (p) => ipcRenderer.send('files:show-in-folder', p),
    takePendingOpens: () => ipcRenderer.invoke('files:take-pending'),
    onOpenRequest: (cb) => listen('files:open-request', cb),
    pickImage: () => ipcRenderer.invoke('files:pick-image'),
    pathForFile: (file) => {
      try {
        return webUtils.getPathForFile(file);
      } catch {
        return '';
      }
    },
  },
  recent: {
    list: () => ipcRenderer.invoke('recent:list'),
    add: (entry) => ipcRenderer.invoke('recent:add', entry),
    remove: (p) => ipcRenderer.invoke('recent:remove', p),
    pin: (p, pinned) => ipcRenderer.invoke('recent:pin', p, pinned),
    clear: () => ipcRenderer.invoke('recent:clear'),
  },
  settings: {
    get: () => ipcRenderer.invoke('settings:get'),
    set: (patch) => ipcRenderer.invoke('settings:set', patch),
  },
  output: {
    pdf: (job) => ipcRenderer.invoke('output:pdf', job),
    print: (job) => ipcRenderer.invoke('output:print', job),
    takeJob: () => ipcRenderer.invoke('output:take-job'),
    ready: () => ipcRenderer.send('output:ready'),
  },
  recovery: {
    save: (entry, data) => ipcRenderer.invoke('recovery:save', entry, data),
    list: () => ipcRenderer.invoke('recovery:list'),
    load: (id) => ipcRenderer.invoke('recovery:load', id),
    remove: (id) => ipcRenderer.invoke('recovery:remove', id),
  },
  convert: {
    libreOfficePath: () => ipcRenderer.invoke('convert:lo-path'),
    withLibreOffice: (data, fromExt, toExt) => ipcRenderer.invoke('convert:lo', data, fromExt, toExt),
    legacyDocText: (data) => ipcRenderer.invoke('convert:doc-text', data),
  },
  spell: {
    onContextMenu: (cb) => listen<[ContextMenuParams]>('ctx:menu', cb),
    replaceMisspelling: (word) => ipcRenderer.send('spell:replace', word),
    addToDictionary: (word) => ipcRenderer.send('spell:add', word),
    languages: () => ipcRenderer.invoke('spell:languages'),
  },
  app: {
    info: () => ipcRenderer.invoke('app:info'),
    relaunch: () => ipcRenderer.send('app:relaunch'),
    toggleDevTools: () => ipcRenderer.send('app:devtools'),
    openExternal: (url) => ipcRenderer.send('app:open-external', url),
    zoom: (factor) => ipcRenderer.send('app:zoom', factor),
    editCommand: (cmd) => ipcRenderer.send('app:edit', cmd),
  },
};

contextBridge.exposeInMainWorld('affice', bridge);
