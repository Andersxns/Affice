import { create } from 'zustand';
import type { AppKind, OpenedFile } from '@/shared/types';
import { uid } from '@/lib/utils';

export type TabKind = 'home' | AppKind;

export type TabSource =
  | { type: 'blank' }
  | { type: 'template'; id: string }
  | { type: 'file'; file: OpenedFile; formatId: string }
  | { type: 'recovery'; id: string; data: string }
  | { type: 'data'; data: unknown };

export interface Tab {
  id: string;
  kind: TabKind;
  title: string;
  path?: string;
  /** Format of the file at `path` (docx, xlsx, …). */
  format?: string;
  dirty: boolean;
  source?: TabSource;
  loading?: boolean;
  /** Recovery id used for autosave snapshots. */
  recoveryId: string;
}

interface WorkspaceState {
  tabs: Tab[];
  activeId: string;
  history: string[];
  openTab(tab: Omit<Tab, 'id' | 'dirty' | 'recoveryId'> & { id?: string; recoveryId?: string }): string;
  removeTab(id: string): void;
  activate(id: string): void;
  updateTab(id: string, patch: Partial<Tab>): void;
  moveTab(id: string, toIndex: number): void;
}

export const HOME_ID = 'home';

export const useWorkspace = create<WorkspaceState>((set, get) => ({
  tabs: [{ id: HOME_ID, kind: 'home', title: 'Home', dirty: false, recoveryId: '' }],
  activeId: HOME_ID,
  history: [HOME_ID],
  openTab(tab) {
    const id = tab.id ?? uid('tab');
    const full: Tab = { dirty: false, recoveryId: tab.recoveryId ?? uid('rec'), ...tab, id };
    set((s) => ({ tabs: [...s.tabs, full], activeId: id, history: [...s.history.filter((h) => h !== id), id] }));
    return id;
  },
  removeTab(id) {
    if (id === HOME_ID) return;
    const { tabs, activeId, history } = get();
    const idx = tabs.findIndex((t) => t.id === id);
    if (idx < 0) return;
    const nextTabs = tabs.filter((t) => t.id !== id);
    const nextHistory = history.filter((h) => h !== id && nextTabs.some((t) => t.id === h));
    let nextActive = activeId;
    if (activeId === id) nextActive = nextHistory[nextHistory.length - 1] ?? nextTabs[Math.max(0, idx - 1)]?.id ?? HOME_ID;
    set({ tabs: nextTabs, activeId: nextActive, history: nextHistory });
  },
  activate(id) {
    if (!get().tabs.some((t) => t.id === id)) return;
    set((s) => ({ activeId: id, history: [...s.history.filter((h) => h !== id), id] }));
  },
  updateTab(id, patch) {
    set((s) => ({ tabs: s.tabs.map((t) => (t.id === id ? { ...t, ...patch } : t)) }));
  },
  moveTab(id, toIndex) {
    set((s) => {
      const tabs = [...s.tabs];
      const from = tabs.findIndex((t) => t.id === id);
      if (from <= 0) return s;
      const [t] = tabs.splice(from, 1);
      tabs.splice(Math.max(1, Math.min(tabs.length, toIndex)), 0, t);
      return { tabs };
    });
  },
}));

export const getTab = (id: string) => useWorkspace.getState().tabs.find((t) => t.id === id);
export const activeTab = () => getTab(useWorkspace.getState().activeId);

/* --------------------------------------------------- editor controllers */

export interface Command {
  id: string;
  title: string;
  category?: string;
  keys?: string;
  icon?: React.ReactNode;
  run: () => void;
  when?: () => boolean;
}

export interface EditorController {
  /** Saves to the current path; asks for one if needed. Resolves true when saved. */
  save(): Promise<boolean>;
  saveAs(formatId?: string): Promise<boolean>;
  exportAs(targetId: string): Promise<void>;
  print(): Promise<void>;
  /** Native (lossless) serialisation used for autosave/crash recovery. */
  snapshot(): string | null;
  focus(): void;
  commands(): Command[];
  info?(): Array<{ label: string; value: string }>;
}

const controllers = new Map<string, EditorController>();

export function registerController(tabId: string, c: EditorController): () => void {
  controllers.set(tabId, c);
  return () => {
    if (controllers.get(tabId) === c) controllers.delete(tabId);
  };
}

export const getController = (tabId: string) => controllers.get(tabId);
