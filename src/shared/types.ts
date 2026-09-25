// Types shared between the Electron main process, the preload bridge and the renderer.

export type AppKind = 'doc' | 'sheet' | 'slides';

export type ThemeMode = 'system' | 'light' | 'dark';
export type MotionMode = 'full' | 'reduced' | 'off';
export type Density = 'comfortable' | 'compact';

export interface Settings {
  theme: ThemeMode;
  accent: string;
  motion: MotionMode;
  density: Density;
  /** Disables blur, heavy shadows and background effects — for older/slower machines. */
  performanceMode: boolean;
  /** Takes effect after restart. */
  hardwareAcceleration: boolean;
  /** Takes effect after restart. */
  nativeTitleBar: boolean;
  uiScale: number;
  autosaveMinutes: number;
  /** Automatically write changes back to files that already have a location on disk. */
  autoSaveToFile: boolean;
  spellcheck: boolean;
  spellcheckLanguages: string[];
  defaultDocFormat: 'docx' | 'odt' | 'afdoc';
  defaultSheetFormat: 'xlsx' | 'ods' | 'afsheet';
  defaultSlidesFormat: 'pptx' | 'afslides';
  authorName: string;
  showStartupDashboard: boolean;
  confirmOnClose: boolean;
  ribbonMode: 'full' | 'simplified';
  restoreTabs: boolean;
  useLibreOfficeForLegacy: boolean;
}

export const DEFAULT_SETTINGS: Settings = {
  theme: 'system',
  accent: '#004fff',
  motion: 'full',
  density: 'comfortable',
  performanceMode: false,
  hardwareAcceleration: true,
  nativeTitleBar: false,
  uiScale: 1,
  autosaveMinutes: 2,
  autoSaveToFile: false,
  spellcheck: true,
  spellcheckLanguages: ['en-US'],
  defaultDocFormat: 'docx',
  defaultSheetFormat: 'xlsx',
  defaultSlidesFormat: 'pptx',
  authorName: '',
  showStartupDashboard: true,
  confirmOnClose: true,
  ribbonMode: 'full',
  restoreTabs: false,
  useLibreOfficeForLegacy: true,
};

export interface OpenedFile {
  path: string;
  name: string;
  data: Uint8Array;
  size: number;
  mtime: number;
}

export interface RecentFile {
  path: string;
  name: string;
  kind: AppKind;
  openedAt: number;
  pinned?: boolean;
  size?: number;
}

export interface FileFilter {
  name: string;
  extensions: string[];
}

export interface PdfJob {
  html: string;
  /** Extra CSS appended after the app's print stylesheet. */
  css?: string;
  /** Body class used to pick the right print styles (e.g. "print-doc"). */
  bodyClass?: string;
  pageWidthIn: number;
  pageHeightIn: number;
  landscape?: boolean;
  marginsIn?: { top: number; right: number; bottom: number; left: number };
  headerHtml?: string;
  footerHtml?: string;
  title?: string;
}

export interface RecoveryEntry {
  id: string;
  kind: AppKind;
  title: string;
  path?: string;
  savedAt: number;
  size: number;
}

export interface ContextMenuParams {
  x: number;
  y: number;
  isEditable: boolean;
  selectionText: string;
  misspelledWord: string;
  dictionarySuggestions: string[];
  editFlags: {
    canCut: boolean;
    canCopy: boolean;
    canPaste: boolean;
    canSelectAll: boolean;
    canUndo: boolean;
    canRedo: boolean;
  };
  linkURL: string;
}

export interface AppInfo {
  version: string;
  electron: string;
  chrome: string;
  node: string;
  platform: string;
  arch: string;
  userData: string;
  libreOffice: string | null;
}

/** API exposed on `window.affice` by the preload script. */
export interface AfficeBridge {
  isElectron: true;
  platform: string;
  window: {
    minimize(): void;
    toggleMaximize(): void;
    close(): void;
    forceClose(): void;
    isMaximized(): Promise<boolean>;
    setFullScreen(on: boolean): void;
    isFullScreen(): Promise<boolean>;
    setTitle(title: string): void;
    setTitleBarColors(color: string, symbolColor: string): void;
    onStateChange(cb: (state: { maximized: boolean; fullscreen: boolean }) => void): () => void;
    onCloseRequest(cb: () => void): () => void;
  };
  files: {
    openDialog(opts: { title?: string; filters?: FileFilter[]; multi?: boolean }): Promise<OpenedFile[]>;
    read(path: string): Promise<OpenedFile>;
    saveDialog(opts: { title?: string; defaultPath?: string; filters?: FileFilter[] }): Promise<string | null>;
    write(path: string, data: Uint8Array): Promise<void>;
    exists(path: string): Promise<boolean>;
    showInFolder(path: string): void;
    takePendingOpens(): Promise<string[]>;
    onOpenRequest(cb: (paths: string[]) => void): () => void;
    pickImage(): Promise<OpenedFile | null>;
    /** Absolute path of a File from drag & drop (desktop only). */
    pathForFile(file: File): string;
  };
  recent: {
    list(): Promise<RecentFile[]>;
    add(entry: Omit<RecentFile, 'openedAt'>): Promise<RecentFile[]>;
    remove(path: string): Promise<RecentFile[]>;
    pin(path: string, pinned: boolean): Promise<RecentFile[]>;
    clear(): Promise<RecentFile[]>;
  };
  settings: {
    get(): Promise<Settings>;
    set(patch: Partial<Settings>): Promise<Settings>;
    /** The system switched between light and dark. */
    onSystemThemeChange(cb: (dark: boolean) => void): () => void;
  };
  output: {
    pdf(job: PdfJob): Promise<Uint8Array>;
    print(job: PdfJob): Promise<boolean>;
    /** Used by the hidden print window. */
    takeJob(): Promise<PdfJob | null>;
    ready(): void;
  };
  recovery: {
    save(entry: RecoveryEntry, data: string): Promise<void>;
    list(): Promise<RecoveryEntry[]>;
    load(id: string): Promise<string | null>;
    remove(id: string): Promise<void>;
  };
  convert: {
    libreOfficePath(): Promise<string | null>;
    withLibreOffice(data: Uint8Array, fromExt: string, toExt: string): Promise<Uint8Array>;
    legacyDocText(data: Uint8Array): Promise<string>;
  };
  spell: {
    onContextMenu(cb: (p: ContextMenuParams) => void): () => void;
    replaceMisspelling(word: string): void;
    addToDictionary(word: string): void;
    languages(): Promise<string[]>;
  };
  app: {
    info(): Promise<AppInfo>;
    relaunch(): void;
    toggleDevTools(): void;
    openExternal(url: string): void;
    zoom(factor: number): void;
    editCommand(cmd: 'cut' | 'copy' | 'paste' | 'pasteAndMatchStyle' | 'selectAll' | 'undo' | 'redo' | 'delete'): void;
  };
}
