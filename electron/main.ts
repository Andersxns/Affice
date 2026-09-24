import { app, BrowserWindow, dialog, ipcMain, nativeTheme, screen, session, shell, Menu } from 'electron';
import fs from 'node:fs';
import fsp from 'node:fs/promises';
import path from 'node:path';
import type { AppInfo, ContextMenuParams, FileFilter, OpenedFile, RecentFile, RecoveryEntry, Settings } from '../src/shared/types';
import { convertWithLibreOffice, findLibreOffice, legacyDocText } from './convert';
import { registerPrintIpc } from './print';
import {
  addRecent,
  clearRecent,
  getSettings,
  listRecent,
  listRecovery,
  loadRecovery,
  loadWindowState,
  pinRecent,
  removeRecent,
  removeRecovery,
  saveRecovery,
  saveWindowStateSync,
  updateSettings,
} from './store';

const APP_ID = 'io.github.andersxns.affice';
const DEV_URL = process.env.VITE_DEV_SERVER_URL;
const DIST = path.join(__dirname, '..', 'dist');
const PRELOAD = path.join(__dirname, 'preload.cjs');
const TITLEBAR_HEIGHT = 40;

/** Extensions Affice knows how to open — used to pick files out of argv. */
const OPENABLE = new Set([
  'docx', 'docm', 'dotx', 'doc', 'odt', 'ott', 'rtf', 'txt', 'md', 'markdown', 'html', 'htm', 'afdoc', 'wpd', 'pages',
  'xlsx', 'xlsm', 'xltx', 'xls', 'xlsb', 'ods', 'ots', 'csv', 'tsv', 'afsheet', 'numbers', 'fods',
  'pptx', 'pptm', 'potx', 'ppsx', 'ppt', 'pps', 'odp', 'otp', 'afslides', 'key',
]);

/* ---------------------------------------------------------- early startup */

app.setName('Affice');
if (process.platform === 'win32') app.setAppUserModelId(APP_ID);

const startupSettings = getSettings();
if (!startupSettings.hardwareAcceleration) app.disableHardwareAcceleration();

const gotLock = app.requestSingleInstanceLock();
if (!gotLock) {
  app.quit();
}

let mainWindow: BrowserWindow | null = null;
let pendingOpens: string[] = filesFromArgv(process.argv, process.cwd());
let allowClose = false;
let rendererHandlesClose = false;

function filesFromArgv(argv: string[], cwd: string): string[] {
  const out: string[] = [];
  // In a packaged app argv[0] is the executable; in dev argv[1] is the app path.
  for (const arg of argv.slice(app.isPackaged ? 1 : 2)) {
    if (!arg || arg.startsWith('-')) continue;
    const full = path.isAbsolute(arg) ? arg : path.resolve(cwd, arg);
    const ext = path.extname(full).slice(1).toLowerCase();
    if (!OPENABLE.has(ext)) continue;
    try {
      if (fs.statSync(full).isFile()) out.push(full);
    } catch {
      /* ignore missing */
    }
  }
  return out;
}

app.on('second-instance', (_e, argv, cwd) => {
  const files = filesFromArgv(argv, cwd);
  if (mainWindow) {
    if (mainWindow.isMinimized()) mainWindow.restore();
    mainWindow.show();
    mainWindow.focus();
    if (files.length) mainWindow.webContents.send('files:open-request', files);
  } else {
    pendingOpens.push(...files);
  }
});

/* ---------------------------------------------------------------- helpers */

function isDarkTheme(s: Settings): boolean {
  if (s.theme === 'dark') return true;
  if (s.theme === 'light') return false;
  return nativeTheme.shouldUseDarkColors;
}

function chromeColors(dark: boolean) {
  return dark ? { bg: '#15171c', symbol: '#e8eaf0' } : { bg: '#eef1f8', symbol: '#1d2433' };
}

async function loadRoute(win: BrowserWindow, route: string): Promise<void> {
  if (DEV_URL) await win.loadURL(`${DEV_URL}#${route}`);
  else await win.loadFile(path.join(DIST, 'index.html'), { hash: route });
}

function isInternalUrl(url: string): boolean {
  if (DEV_URL && url.startsWith(DEV_URL)) return true;
  return url.startsWith('file://') && url.includes('/dist/index.html');
}

function openExternalSafely(url: string): void {
  try {
    const u = new URL(url);
    if (['http:', 'https:', 'mailto:'].includes(u.protocol)) void shell.openExternal(url);
  } catch {
    /* not a URL */
  }
}

async function readOpened(p: string): Promise<OpenedFile> {
  const [data, stat] = await Promise.all([fsp.readFile(p), fsp.stat(p)]);
  return { path: p, name: path.basename(p), data: new Uint8Array(data), size: stat.size, mtime: stat.mtimeMs };
}

function visibleBounds(state: ReturnType<typeof loadWindowState>) {
  const { x, y, width, height } = state;
  if (x === undefined || y === undefined) return { width, height };
  const onScreen = screen.getAllDisplays().some((d) => {
    const a = d.workArea;
    return x + 80 > a.x && y + 40 > a.y && x < a.x + a.width - 80 && y < a.y + a.height - 40;
  });
  return onScreen ? { x, y, width, height } : { width, height };
}

/* ----------------------------------------------------------------- window */

function createMainWindow(): void {
  const settings = getSettings();
  const state = loadWindowState();
  const dark = isDarkTheme(settings);
  const colors = chromeColors(dark);
  const useOverlay = !settings.nativeTitleBar;

  const iconPath = path.join(DIST, 'icon-256.png');
  const win = new BrowserWindow({
    ...visibleBounds(state),
    minWidth: 760,
    minHeight: 520,
    show: false,
    title: 'Affice',
    backgroundColor: dark ? '#15171c' : '#f4f6fb',
    icon: fs.existsSync(iconPath) ? iconPath : undefined,
    autoHideMenuBar: true,
    ...(useOverlay
      ? {
          titleBarStyle: 'hidden' as const,
          titleBarOverlay: { color: colors.bg, symbolColor: colors.symbol, height: TITLEBAR_HEIGHT },
        }
      : {}),
    webPreferences: {
      preload: PRELOAD,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      spellcheck: settings.spellcheck,
      backgroundThrottling: true,
    },
  });
  mainWindow = win;

  if (state.maximized) win.maximize();
  win.once('ready-to-show', () => win.show());
  win.webContents.setVisualZoomLevelLimits(1, 1).catch(() => undefined);
  if (settings.uiScale && settings.uiScale !== 1) win.webContents.setZoomFactor(settings.uiScale);

  const sendState = () => {
    if (win.isDestroyed()) return;
    win.webContents.send('win:state', { maximized: win.isMaximized(), fullscreen: win.isFullScreen() });
  };
  win.on('maximize', sendState);
  win.on('unmaximize', sendState);
  win.on('enter-full-screen', sendState);
  win.on('leave-full-screen', sendState);

  win.on('close', (e) => {
    const bounds = win.getNormalBounds();
    saveWindowStateSync({ ...bounds, maximized: win.isMaximized() });
    if (!allowClose && rendererHandlesClose && !win.webContents.isCrashed()) {
      e.preventDefault();
      win.webContents.send('win:close-request');
    }
  });
  win.on('closed', () => {
    mainWindow = null;
  });

  win.webContents.setWindowOpenHandler(({ url }) => {
    openExternalSafely(url);
    return { action: 'deny' };
  });
  win.webContents.on('will-navigate', (e, url) => {
    if (!isInternalUrl(url)) {
      e.preventDefault();
      openExternalSafely(url);
    }
  });
  win.webContents.on('render-process-gone', () => {
    rendererHandlesClose = false;
  });

  win.webContents.on('context-menu', (_e, p) => {
    const params: ContextMenuParams = {
      x: p.x,
      y: p.y,
      isEditable: p.isEditable,
      selectionText: p.selectionText,
      misspelledWord: p.misspelledWord,
      dictionarySuggestions: p.dictionarySuggestions ?? [],
      editFlags: {
        canCut: p.editFlags.canCut,
        canCopy: p.editFlags.canCopy,
        canPaste: p.editFlags.canPaste,
        canSelectAll: p.editFlags.canSelectAll,
        canUndo: p.editFlags.canUndo,
        canRedo: p.editFlags.canRedo,
      },
      linkURL: p.linkURL,
    };
    win.webContents.send('ctx:menu', params);
  });

  win.webContents.on('before-input-event', (e, input) => {
    if (input.type !== 'keyDown') return;
    const ctrl = input.control || input.meta;
    if (ctrl && input.shift && input.key.toLowerCase() === 'i') {
      win.webContents.toggleDevTools();
      e.preventDefault();
    }
    if (DEV_URL && ((ctrl && input.key.toLowerCase() === 'r') || input.key === 'F5') && !input.shift) {
      win.webContents.reloadIgnoringCache();
      e.preventDefault();
    }
  });

  void loadRoute(win, 'app');
}

/* -------------------------------------------------------------------- IPC */

function senderWindow(e: Electron.IpcMainEvent | Electron.IpcMainInvokeEvent): BrowserWindow | null {
  return BrowserWindow.fromWebContents(e.sender);
}

function registerIpc(): void {
  // window
  ipcMain.on('win:minimize', (e) => senderWindow(e)?.minimize());
  ipcMain.on('win:toggle-maximize', (e) => {
    const w = senderWindow(e);
    if (!w) return;
    if (w.isMaximized()) w.unmaximize();
    else w.maximize();
  });
  ipcMain.on('win:close', (e) => senderWindow(e)?.close());
  ipcMain.on('win:force-close', (e) => {
    allowClose = true;
    senderWindow(e)?.close();
  });
  ipcMain.on('win:has-close-handler', () => {
    rendererHandlesClose = true;
  });
  ipcMain.handle('win:is-maximized', (e) => senderWindow(e)?.isMaximized() ?? false);
  ipcMain.on('win:set-fullscreen', (e, on: boolean) => senderWindow(e)?.setFullScreen(Boolean(on)));
  ipcMain.handle('win:is-fullscreen', (e) => senderWindow(e)?.isFullScreen() ?? false);
  ipcMain.on('win:set-title', (e, title: string) => senderWindow(e)?.setTitle(String(title).slice(0, 300)));
  ipcMain.on('win:titlebar-colors', (e, color: string, symbolColor: string) => {
    const w = senderWindow(e);
    if (!w || getSettings().nativeTitleBar) return;
    try {
      w.setTitleBarOverlay({ color, symbolColor, height: TITLEBAR_HEIGHT });
    } catch {
      /* overlay not enabled */
    }
  });

  // files
  ipcMain.handle('files:open-dialog', async (e, opts: { title?: string; filters?: FileFilter[]; multi?: boolean }) => {
    const w = senderWindow(e);
    const props: Array<'openFile' | 'multiSelections'> = ['openFile'];
    if (opts?.multi) props.push('multiSelections');
    const res = w
      ? await dialog.showOpenDialog(w, { title: opts?.title, filters: opts?.filters, properties: props })
      : await dialog.showOpenDialog({ title: opts?.title, filters: opts?.filters, properties: props });
    if (res.canceled) return [];
    return Promise.all(res.filePaths.map(readOpened));
  });
  ipcMain.handle('files:pick-image', async (e) => {
    const w = senderWindow(e);
    const opts: Electron.OpenDialogOptions = {
      title: 'Insert picture',
      filters: [{ name: 'Images', extensions: ['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'svg', 'avif', 'ico'] }],
      properties: ['openFile'],
    };
    const res = w ? await dialog.showOpenDialog(w, opts) : await dialog.showOpenDialog(opts);
    if (res.canceled || !res.filePaths[0]) return null;
    return readOpened(res.filePaths[0]);
  });
  ipcMain.handle('files:read', (_e, p: string) => readOpened(p));
  ipcMain.handle('files:save-dialog', async (e, opts: { title?: string; defaultPath?: string; filters?: FileFilter[] }) => {
    const w = senderWindow(e);
    const o: Electron.SaveDialogOptions = { title: opts?.title, defaultPath: opts?.defaultPath, filters: opts?.filters };
    const res = w ? await dialog.showSaveDialog(w, o) : await dialog.showSaveDialog(o);
    return res.canceled || !res.filePath ? null : res.filePath;
  });
  ipcMain.handle('files:write', async (_e, p: string, data: Uint8Array) => {
    const tmp = `${p}.affice-${process.pid}.tmp`;
    await fsp.writeFile(tmp, data);
    try {
      await fsp.rename(tmp, p);
    } catch {
      // e.g. file locked or cross-device: fall back to a direct write
      await fsp.writeFile(p, data);
      await fsp.rm(tmp, { force: true });
    }
    app.addRecentDocument(p);
  });
  ipcMain.handle('files:exists', async (_e, p: string) => {
    try {
      await fsp.access(p);
      return true;
    } catch {
      return false;
    }
  });
  ipcMain.on('files:show-in-folder', (_e, p: string) => shell.showItemInFolder(p));
  ipcMain.handle('files:take-pending', () => {
    const list = pendingOpens;
    pendingOpens = [];
    return list;
  });

  // recent
  ipcMain.handle('recent:list', () => listRecent());
  ipcMain.handle('recent:add', (_e, entry: Omit<RecentFile, 'openedAt'>) => {
    app.addRecentDocument(entry.path);
    return addRecent(entry);
  });
  ipcMain.handle('recent:remove', (_e, p: string) => removeRecent(p));
  ipcMain.handle('recent:pin', (_e, p: string, pinned: boolean) => pinRecent(p, pinned));
  ipcMain.handle('recent:clear', () => clearRecent());

  // settings
  ipcMain.handle('settings:get', () => getSettings());
  ipcMain.handle('settings:set', async (e, patch: Partial<Settings>) => {
    const next = await updateSettings(patch);
    nativeTheme.themeSource = next.theme;
    const w = senderWindow(e);
    if (w) {
      if ('uiScale' in patch) w.webContents.setZoomFactor(next.uiScale || 1);
      if ('spellcheck' in patch) w.webContents.session.setSpellCheckerEnabled(next.spellcheck);
      if ('spellcheckLanguages' in patch) applySpellLanguages(next.spellcheckLanguages);
    }
    return next;
  });

  // recovery
  ipcMain.handle('recovery:save', (_e, entry: RecoveryEntry, data: string) => saveRecovery(entry, data));
  ipcMain.handle('recovery:list', () => listRecovery());
  ipcMain.handle('recovery:load', (_e, id: string) => loadRecovery(id));
  ipcMain.handle('recovery:remove', (_e, id: string) => removeRecovery(id));

  // conversion
  ipcMain.handle('convert:lo-path', () => findLibreOffice());
  ipcMain.handle('convert:lo', (_e, data: Uint8Array, fromExt: string, toExt: string) => convertWithLibreOffice(data, fromExt, toExt));
  ipcMain.handle('convert:doc-text', (_e, data: Uint8Array) => legacyDocText(data));

  // spelling
  ipcMain.on('spell:replace', (e, word: string) => e.sender.replaceMisspelling(word));
  ipcMain.on('spell:add', (e, word: string) => e.sender.session.addWordToSpellCheckerDictionary(word));
  ipcMain.handle('spell:languages', () => session.defaultSession.availableSpellCheckerLanguages);

  // app
  ipcMain.handle('app:info', (): AppInfo => ({
    version: app.getVersion(),
    electron: process.versions.electron,
    chrome: process.versions.chrome,
    node: process.versions.node,
    platform: process.platform,
    arch: process.arch,
    userData: app.getPath('userData'),
    libreOffice: findLibreOffice(),
  }));
  ipcMain.on('app:relaunch', () => {
    app.relaunch();
    allowClose = true;
    app.quit();
  });
  ipcMain.on('app:devtools', (e) => e.sender.toggleDevTools());
  ipcMain.on('app:open-external', (_e, url: string) => openExternalSafely(url));
  ipcMain.on('app:edit', (e, cmd: string) => {
    const wc = e.sender;
    switch (cmd) {
      case 'cut': wc.cut(); break;
      case 'copy': wc.copy(); break;
      case 'paste': wc.paste(); break;
      case 'pasteAndMatchStyle': wc.pasteAndMatchStyle(); break;
      case 'selectAll': wc.selectAll(); break;
      case 'undo': wc.undo(); break;
      case 'redo': wc.redo(); break;
      case 'delete': wc.delete(); break;
    }
  });
  ipcMain.on('app:zoom', (e, factor: number) => e.sender.setZoomFactor(Math.min(2, Math.max(0.6, Number(factor) || 1))));

  registerPrintIpc(PRELOAD, loadRoute);
}

function applySpellLanguages(langs: string[]): void {
  if (process.platform === 'darwin') return;
  const available = new Set(session.defaultSession.availableSpellCheckerLanguages);
  const valid = langs.filter((l) => available.has(l));
  try {
    session.defaultSession.setSpellCheckerLanguages(valid.length ? valid : ['en-US']);
  } catch {
    /* ignore unsupported */
  }
}

/* ------------------------------------------------------------- lifecycle */

if (gotLock) {
  app.whenReady().then(() => {
    Menu.setApplicationMenu(null);
    nativeTheme.themeSource = getSettings().theme;

    const allowed = new Set(['local-fonts', 'clipboard-read', 'clipboard-sanitized-write', 'fullscreen', 'window-management']);
    session.defaultSession.setPermissionRequestHandler((_wc, permission, cb) => cb(allowed.has(permission)));
    session.defaultSession.setPermissionCheckHandler((_wc, permission) => allowed.has(permission));
    applySpellLanguages(getSettings().spellcheckLanguages);

    nativeTheme.on('updated', () => {
      mainWindow?.webContents.send('theme:system-changed', nativeTheme.shouldUseDarkColors);
    });

    registerIpc();
    createMainWindow();
  });

  app.on('window-all-closed', () => {
    app.quit();
  });

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createMainWindow();
  });
}
