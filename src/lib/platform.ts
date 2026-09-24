// Platform abstraction: uses the Electron bridge when available and falls back to
// browser APIs otherwise (so Affice can also run as a plain web app).
import {
  DEFAULT_SETTINGS,
  type AfficeBridge,
  type AppInfo,
  type ContextMenuParams,
  type FileFilter,
  type OpenedFile,
  type PdfJob,
  type RecentFile,
  type RecoveryEntry,
  type Settings,
} from '@/shared/types';

declare global {
  interface Window {
    affice?: AfficeBridge;
  }
}

export const bridge: AfficeBridge | undefined = typeof window !== 'undefined' ? window.affice : undefined;
export const isElectron = Boolean(bridge?.isElectron);
export const platformName: string = bridge?.platform ?? (navigator.userAgent.includes('Windows') ? 'win32' : 'linux');
export const isWindows = platformName === 'win32';
export const isMac = platformName === 'darwin';

/* ------------------------------------------------------------------ web fallbacks */

const LS = {
  get<T>(key: string, fallback: T): T {
    try {
      const raw = localStorage.getItem(key);
      return raw ? (JSON.parse(raw) as T) : fallback;
    } catch {
      return fallback;
    }
  },
  set(key: string, value: unknown) {
    try {
      localStorage.setItem(key, JSON.stringify(value));
    } catch {
      /* storage full or unavailable */
    }
  },
};

function pickFilesWeb(filters?: FileFilter[], multi?: boolean): Promise<OpenedFile[]> {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.multiple = Boolean(multi);
    const exts = (filters ?? []).flatMap((f) => f.extensions).filter((e) => e !== '*');
    if (exts.length) input.accept = exts.map((e) => `.${e}`).join(',');
    input.style.display = 'none';
    document.body.appendChild(input);
    let settled = false;
    input.onchange = async () => {
      settled = true;
      const files = Array.from(input.files ?? []);
      const out = await Promise.all(
        files.map(async (f) => ({
          path: '',
          name: f.name,
          data: new Uint8Array(await f.arrayBuffer()),
          size: f.size,
          mtime: f.lastModified,
        })),
      );
      input.remove();
      resolve(out);
    };
    window.addEventListener(
      'focus',
      () =>
        setTimeout(() => {
          if (!settled) {
            input.remove();
            resolve([]);
          }
        }, 800),
      { once: true },
    );
    input.click();
  });
}

function downloadBlob(name: string, data: Uint8Array | Blob, type = 'application/octet-stream') {
  const blob = data instanceof Blob ? data : new Blob([data as BlobPart], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

function printHtmlWeb(job: PdfJob): Promise<boolean> {
  return new Promise((resolve) => {
    const frame = document.createElement('iframe');
    frame.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0;';
    document.body.appendChild(frame);
    const doc = frame.contentDocument!;
    const styles = Array.from(document.querySelectorAll('style,link[rel="stylesheet"]'))
      .map((n) => n.outerHTML)
      .join('\n');
    const m = job.marginsIn ?? { top: 0, right: 0, bottom: 0, left: 0 };
    doc.open();
    doc.write(`<!doctype html><html><head><meta charset="utf-8"><title>${job.title ?? 'Affice'}</title>${styles}
      <style>@page{size:${job.pageWidthIn}in ${job.pageHeightIn}in;margin:${m.top}in ${m.right}in ${m.bottom}in ${m.left}in}
      html,body{overflow:visible!important;height:auto!important;background:#fff!important}${job.css ?? ''}</style></head>
      <body class="print-root ${job.bodyClass ?? ''}">${job.html}</body></html>`);
    doc.close();
    const go = async () => {
      try {
        await (doc as Document & { fonts?: FontFaceSet }).fonts?.ready;
      } catch {
        /* ignore */
      }
      frame.contentWindow?.focus();
      frame.contentWindow?.print();
      setTimeout(() => frame.remove(), 1000);
      resolve(true);
    };
    setTimeout(go, 300);
  });
}

const webApi: AfficeBridge = {
  isElectron: true as const,
  platform: platformName,
  window: {
    minimize: () => undefined,
    toggleMaximize: () => undefined,
    close: () => window.close(),
    forceClose: () => window.close(),
    isMaximized: async () => false,
    setFullScreen: (on) => {
      if (on) void document.documentElement.requestFullscreen?.().catch(() => undefined);
      else if (document.fullscreenElement) void document.exitFullscreen().catch(() => undefined);
    },
    isFullScreen: async () => Boolean(document.fullscreenElement),
    setTitle: (t) => {
      document.title = t;
    },
    setTitleBarColors: () => undefined,
    onStateChange: (cb) => {
      const h = () => cb({ maximized: false, fullscreen: Boolean(document.fullscreenElement) });
      document.addEventListener('fullscreenchange', h);
      return () => document.removeEventListener('fullscreenchange', h);
    },
    onCloseRequest: () => () => undefined,
  },
  files: {
    openDialog: (opts) => pickFilesWeb(opts.filters, opts.multi),
    read: async () => {
      throw new Error('Opening files by path is only available in the desktop app.');
    },
    saveDialog: async (opts) => `download:${(opts.defaultPath ?? 'Untitled').split(/[\\/]/).pop()}`,
    write: async (p, data) => downloadBlob(p.replace(/^download:/, ''), data),
    exists: async () => false,
    showInFolder: () => undefined,
    takePendingOpens: async () => [],
    onOpenRequest: () => () => undefined,
    pickImage: async () => {
      const files = await pickFilesWeb([{ name: 'Images', extensions: ['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'svg'] }]);
      return files[0] ?? null;
    },
    pathForFile: () => '',
  },
  recent: {
    list: async () => LS.get<RecentFile[]>('affice.recent', []),
    add: async (entry) => {
      const list = LS.get<RecentFile[]>('affice.recent', []).filter((r) => r.name !== entry.name);
      const next = [{ ...entry, openedAt: Date.now() }, ...list].slice(0, 40);
      LS.set('affice.recent', next);
      return next;
    },
    remove: async (p) => {
      const next = LS.get<RecentFile[]>('affice.recent', []).filter((r) => r.path !== p);
      LS.set('affice.recent', next);
      return next;
    },
    pin: async (p, pinned) => {
      const next = LS.get<RecentFile[]>('affice.recent', []).map((r) => (r.path === p ? { ...r, pinned } : r));
      LS.set('affice.recent', next);
      return next;
    },
    clear: async () => {
      LS.set('affice.recent', []);
      return [];
    },
  },
  settings: {
    get: async () => ({ ...DEFAULT_SETTINGS, ...LS.get<Partial<Settings>>('affice.settings', {}) }),
    set: async (patch) => {
      const next = { ...DEFAULT_SETTINGS, ...LS.get<Partial<Settings>>('affice.settings', {}), ...patch };
      LS.set('affice.settings', next);
      return next;
    },
  },
  output: {
    pdf: async () => {
      throw new Error('WEB_PDF_UNSUPPORTED');
    },
    print: (job) => printHtmlWeb(job),
    takeJob: async () => null,
    ready: () => undefined,
  },
  recovery: {
    save: async (entry: RecoveryEntry, data: string) => {
      const all = LS.get<Record<string, { entry: RecoveryEntry; data: string }>>('affice.recovery', {});
      all[entry.id] = { entry: { ...entry, size: data.length }, data };
      LS.set('affice.recovery', all);
    },
    list: async () => Object.values(LS.get<Record<string, { entry: RecoveryEntry }>>('affice.recovery', {})).map((v) => v.entry),
    load: async (id) => LS.get<Record<string, { data: string }>>('affice.recovery', {})[id]?.data ?? null,
    remove: async (id) => {
      const all = LS.get<Record<string, unknown>>('affice.recovery', {});
      delete all[id];
      LS.set('affice.recovery', all);
    },
  },
  convert: {
    libreOfficePath: async () => null,
    withLibreOffice: async () => {
      throw new Error('LibreOffice conversion is only available in the desktop app.');
    },
    legacyDocText: async () => {
      throw new Error('Legacy .doc files can only be opened in the desktop app.');
    },
  },
  spell: {
    onContextMenu: (_cb: (p: ContextMenuParams) => void) => () => undefined,
    replaceMisspelling: () => undefined,
    addToDictionary: () => undefined,
    languages: async () => [],
  },
  app: {
    info: async (): Promise<AppInfo> => ({
      version: __APP_VERSION__,
      electron: '—',
      chrome: navigator.userAgent.match(/Chrome\/([\d.]+)/)?.[1] ?? '—',
      node: '—',
      platform: 'web',
      arch: '—',
      userData: 'browser storage',
      libreOffice: null,
    }),
    relaunch: () => location.reload(),
    toggleDevTools: () => undefined,
    openExternal: (url) => window.open(url, '_blank', 'noopener'),
    zoom: (factor) => {
      (document.body.style as CSSStyleDeclaration & { zoom: string }).zoom = String(factor);
    },
    editCommand: (cmd) => {
      if (cmd === 'paste' || cmd === 'pasteAndMatchStyle') {
        void navigator.clipboard?.readText?.().then((t) => document.execCommand('insertText', false, t)).catch(() => undefined);
      } else document.execCommand(cmd === 'delete' ? 'delete' : cmd);
    },
  },
};

export const api: AfficeBridge = bridge ?? webApi;

/** Save bytes to a user-chosen location. Returns the final path (or a pseudo path in the browser). */
export async function saveBytesAs(defaultName: string, data: Uint8Array, filters: FileFilter[], title?: string): Promise<string | null> {
  const p = await api.files.saveDialog({ title, defaultPath: defaultName, filters });
  if (!p) return null;
  await api.files.write(p, data);
  return p;
}

export { downloadBlob };
