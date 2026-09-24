import { create } from 'zustand';
import type { AppKind, OpenedFile, RecentFile } from '@/shared/types';
import { api, isElectron } from '@/lib/platform';
import { formatForExt, KIND_LABEL, openFilters } from '@/lib/formats';
import { basename, extname, stripExt } from '@/lib/utils';
import { alertDialog, saveChangesDialog } from '@/ui/dialog';
import { toast } from '@/ui/toast';
import { getController, getTab, HOME_ID, useWorkspace, type TabSource } from './workspace';

/* ------------------------------------------------------------- recent files */

interface RecentState {
  items: RecentFile[];
  refresh(): Promise<void>;
  add(entry: Omit<RecentFile, 'openedAt'>): Promise<void>;
  remove(path: string): Promise<void>;
  pin(path: string, pinned: boolean): Promise<void>;
  clear(): Promise<void>;
}

export const useRecent = create<RecentState>((set) => ({
  items: [],
  async refresh() {
    set({ items: await api.recent.list() });
  },
  async add(entry) {
    set({ items: await api.recent.add(entry) });
  },
  async remove(path) {
    set({ items: await api.recent.remove(path) });
  },
  async pin(path, pinned) {
    set({ items: await api.recent.pin(path, pinned) });
  },
  async clear() {
    set({ items: await api.recent.clear() });
  },
}));

/* ------------------------------------------------------------ new documents */

const untitledCounter: Record<AppKind, number> = { doc: 0, sheet: 0, slides: 0 };

export function newDocument(kind: AppKind, templateId?: string, title?: string): string {
  untitledCounter[kind] += 1;
  const base = title ?? `${kind === 'doc' ? 'Document' : kind === 'sheet' ? 'Book' : 'Presentation'} ${untitledCounter[kind]}`;
  const source: TabSource = templateId ? { type: 'template', id: templateId } : { type: 'blank' };
  return useWorkspace.getState().openTab({ kind, title: base, source });
}

/* ---------------------------------------------------------------- opening */

async function convertLegacy(file: OpenedFile, fromExt: string, toExt: string): Promise<OpenedFile | null> {
  const lo = await api.convert.libreOfficePath();
  if (!lo) return null;
  const t = toast.loading(`Converting ${file.name} with LibreOffice…`);
  try {
    const data = await api.convert.withLibreOffice(file.data, fromExt, toExt);
    toast.dismiss(t);
    return { ...file, data };
  } catch (e) {
    toast.update(t, 'error', `Could not convert ${file.name}`, { detail: String((e as Error).message ?? e) });
    throw e;
  }
}

export async function openFile(file: OpenedFile): Promise<string | null> {
  const ext = extname(file.name);
  const fmt = formatForExt(ext);
  if (!fmt || fmt.open === 'none') {
    await alertDialog('Unsupported file type', `Affice can't open “.${ext}” files yet.`);
    return null;
  }
  // Already open? Just switch to it.
  if (file.path) {
    const existing = useWorkspace.getState().tabs.find((t) => t.path && t.path === file.path);
    if (existing) {
      useWorkspace.getState().activate(existing.id);
      return existing.id;
    }
  }

  let data = file;
  let dataFormat = fmt.id;
  try {
    if (fmt.open === 'libreoffice' || fmt.id === 'doc') {
      const target = fmt.kind === 'doc' ? 'docx' : fmt.kind === 'sheet' ? 'xlsx' : 'pptx';
      const converted = isElectron ? await convertLegacy(file, ext, target).catch(() => null) : null;
      if (converted) {
        data = converted;
        dataFormat = target;
      } else if (fmt.id === 'doc') {
        dataFormat = 'doc-text';
      } else {
        await alertDialog(
          `${fmt.label} needs LibreOffice`,
          'Affice uses a free LibreOffice installation (if present) to read this older format. Install LibreOffice from libreoffice.org, then try again — or ask the sender for a .docx, .xlsx or .pptx copy.',
        );
        return null;
      }
    }
  } catch {
    return null;
  }

  const id = useWorkspace.getState().openTab({
    kind: fmt.kind,
    title: file.name,
    path: file.path || undefined,
    format: fmt.id,
    source: { type: 'file', file: data, formatId: dataFormat },
    loading: true,
  });
  if (file.path) void useRecent.getState().add({ path: file.path, name: file.name, kind: fmt.kind, size: file.size });
  return id;
}

export async function openPaths(paths: string[]): Promise<void> {
  for (const p of paths) {
    try {
      const f = await api.files.read(p);
      await openFile(f);
    } catch (e) {
      toast.error(`Couldn't open ${basename(p)}`, { detail: (e as Error).message });
      if (/ENOENT|no such file/i.test(String((e as Error).message))) void useRecent.getState().remove(p);
    }
  }
}

export async function showOpenDialog(kind?: AppKind): Promise<void> {
  const lo = isElectron ? await api.convert.libreOfficePath() : null;
  const files = await api.files.openDialog({ title: kind ? `Open ${KIND_LABEL[kind].toLowerCase()}` : 'Open', filters: openFilters(kind, Boolean(lo)), multi: true });
  for (const f of files) await openFile(f);
}

/** Handles files dropped onto the window. */
export async function openDroppedFiles(list: FileList | File[]): Promise<void> {
  const files = Array.from(list);
  for (const f of files) {
    const path = api.files.pathForFile(f);
    await openFile({ path, name: f.name, data: new Uint8Array(await f.arrayBuffer()), size: f.size, mtime: f.lastModified });
  }
}

/* ----------------------------------------------------------------- closing */

export async function closeTab(id: string): Promise<boolean> {
  if (id === HOME_ID) return false;
  const tab = getTab(id);
  if (!tab) return true;
  if (tab.dirty) {
    useWorkspace.getState().activate(id);
    const choice = await saveChangesDialog(tab.title);
    if (choice === 'cancel') return false;
    if (choice === 'save') {
      const ok = await getController(id)?.save();
      if (!ok) return false;
    }
  }
  void api.recovery.remove(tab.recoveryId);
  useWorkspace.getState().removeTab(id);
  return true;
}

export async function closeAllTabs(): Promise<boolean> {
  for (const t of [...useWorkspace.getState().tabs]) {
    if (t.id === HOME_ID) continue;
    const ok = await closeTab(t.id);
    if (!ok) return false;
  }
  return true;
}

export async function requestQuit(): Promise<void> {
  const ok = await closeAllTabs();
  if (ok) api.window.forceClose();
}

export function titleFromPath(p: string): string {
  return basename(p);
}

export function suggestedName(title: string, ext: string): string {
  const base = stripExt(title).replace(/[\\/:*?"<>|]+/g, '-').trim() || 'Untitled';
  return `${base}.${ext}`;
}
