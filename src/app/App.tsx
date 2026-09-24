import { lazy, Suspense, useEffect, useState } from 'react';
import { Upload } from 'lucide-react';
import { api, isElectron } from '@/lib/platform';
import { Home } from '@/home/Home';
import { DialogHost } from '@/ui/dialog';
import { ContextMenuHost, openContextMenu, type MenuItem } from '@/ui/menu';
import { ToastHost, toast } from '@/ui/toast';
import { TooltipHost } from '@/ui/tooltip';
import type { ContextMenuParams } from '@/shared/types';
import { closeTab, newDocument, openDroppedFiles, openPaths, requestQuit, showOpenDialog, useRecent } from './actions';
import { CommandPalette, openCommandPalette } from './CommandPalette';
import { contextMenuExtender } from './contextMenu';
import { useSettings } from './settings';
import { openSettings } from './SettingsDialog';
import { TitleBar } from './TitleBar';
import { getController, HOME_ID, useWorkspace, type Tab } from './workspace';

const DocsEditor = lazy(() => import('@/docs/DocsEditor'));
const SheetsEditor = lazy(() => import('@/sheets/SheetsEditor'));
const SlidesEditor = lazy(() => import('@/slides/SlidesEditor'));

function EditorFallback() {
  return (
    <div className="editor-loading">
      <div className="spinner lg" />
    </div>
  );
}

function TabView({ tab, active }: { tab: Tab; active: boolean }) {
  return (
    <section className={`tab-view${active ? ' active' : ''}`} aria-hidden={!active}>
      {tab.kind === 'home' ? (
        <Home />
      ) : (
        <Suspense fallback={<EditorFallback />}>
          {tab.kind === 'doc' && <DocsEditor tab={tab} active={active} />}
          {tab.kind === 'sheet' && <SheetsEditor tab={tab} active={active} />}
          {tab.kind === 'slides' && <SlidesEditor tab={tab} active={active} />}
        </Suspense>
      )}
    </section>
  );
}

function useGlobalShortcuts() {
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.defaultPrevented) return;
      const mod = e.ctrlKey || e.metaKey;
      const key = e.key.toLowerCase();
      const ws = useWorkspace.getState();
      const active = ws.tabs.find((t) => t.id === ws.activeId);
      const ctrl = active ? getController(active.id) : undefined;
      const handled = () => {
        e.preventDefault();
        e.stopPropagation();
      };
      if (mod && !e.shiftKey && !e.altKey && key === 'n') {
        handled();
        newDocument(active && active.kind !== 'home' ? active.kind : 'doc');
      } else if (mod && !e.shiftKey && key === 't') {
        handled();
        ws.activate(HOME_ID);
      } else if (mod && !e.shiftKey && key === 'o') {
        handled();
        void showOpenDialog(active && active.kind !== 'home' ? active.kind : undefined);
      } else if (mod && key === 's') {
        handled();
        if (ctrl) void (e.shiftKey ? ctrl.saveAs() : ctrl.save());
      } else if (mod && !e.shiftKey && key === 'p') {
        handled();
        if (ctrl) void ctrl.print();
      } else if ((mod && key === 'w') || (mod && e.key === 'F4')) {
        handled();
        if (active && active.id !== HOME_ID) void closeTab(active.id);
      } else if (mod && e.key === 'Tab') {
        handled();
        const i = ws.tabs.findIndex((t) => t.id === ws.activeId);
        const n = ws.tabs.length;
        ws.activate(ws.tabs[(i + (e.shiftKey ? -1 : 1) + n) % n].id);
      } else if ((mod && e.shiftKey && key === 'p') || (e.altKey && key === 'q')) {
        handled();
        openCommandPalette();
      } else if (mod && key === ',') {
        handled();
        openSettings();
      } else if (e.key === 'F11') {
        handled();
        void api.window.isFullScreen().then((fs) => api.window.setFullScreen(!fs));
      } else if (mod && /^[1-9]$/.test(e.key) && !e.shiftKey && !e.altKey) {
        const t = ws.tabs[Number(e.key) - 1];
        if (t) {
          handled();
          ws.activate(t.id);
        }
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, []);
}

function useAutosave() {
  const minutes = useSettings((s) => s.settings.autosaveMinutes);
  const autoSaveToFile = useSettings((s) => s.settings.autoSaveToFile);
  useEffect(() => {
    if (!minutes) return;
    const timer = setInterval(() => {
      for (const t of useWorkspace.getState().tabs) {
        if (!t.dirty || t.kind === 'home') continue;
        const snap = getController(t.id)?.snapshot();
        if (snap) void api.recovery.save({ id: t.recoveryId, kind: t.kind, title: t.title, path: t.path, savedAt: Date.now(), size: snap.length }, snap);
      }
    }, minutes * 60_000);
    return () => clearInterval(timer);
  }, [minutes]);
  useEffect(() => {
    if (!autoSaveToFile || !isElectron) return;
    const timer = setInterval(() => {
      for (const t of useWorkspace.getState().tabs) {
        if (t.dirty && t.path && t.kind !== 'home') void getController(t.id)?.save();
      }
    }, 12_000);
    return () => clearInterval(timer);
  }, [autoSaveToFile]);
}

function spellingItems(p: ContextMenuParams): MenuItem[] {
  const items: MenuItem[] = [];
  if (p.misspelledWord) {
    if (p.dictionarySuggestions.length === 0) items.push({ label: 'No spelling suggestions', disabled: true });
    for (const s of p.dictionarySuggestions.slice(0, 6)) items.push({ label: s, onSelect: () => api.spell.replaceMisspelling(s) });
    items.push({ label: 'Add to dictionary', onSelect: () => api.spell.addToDictionary(p.misspelledWord) });
    items.push({ separator: true });
  }
  return items;
}

function editItems(p: ContextMenuParams): MenuItem[] {
  return [
    { label: 'Cut', shortcut: 'Ctrl+X', disabled: !p.editFlags.canCut, onSelect: () => api.app.editCommand('cut') },
    { label: 'Copy', shortcut: 'Ctrl+C', disabled: !p.editFlags.canCopy, onSelect: () => api.app.editCommand('copy') },
    { label: 'Paste', shortcut: 'Ctrl+V', disabled: !p.editFlags.canPaste, onSelect: () => api.app.editCommand('paste') },
    { label: 'Select all', shortcut: 'Ctrl+A', onSelect: () => api.app.editCommand('selectAll') },
  ];
}

export function App() {
  const tabs = useWorkspace((s) => s.tabs);
  const activeId = useWorkspace((s) => s.activeId);
  const loaded = useSettings((s) => s.loaded);
  const [dragging, setDragging] = useState(false);

  useGlobalShortcuts();
  useAutosave();

  useEffect(() => {
    void useSettings.getState().load();
    void useRecent.getState().refresh();
    void api.files.takePendingOpens().then((paths) => {
      if (paths.length) void openPaths(paths);
    });
    const offOpen = api.files.onOpenRequest((paths) => void openPaths(paths));
    const offClose = api.window.onCloseRequest(() => void requestQuit());
    const offCtx = api.spell.onContextMenu((p) => {
      const extra = contextMenuExtender.current?.(p);
      if (!p.isEditable && !extra) return;
      const items = [...spellingItems(p), ...(extra ?? editItems(p))];
      if (items.length) openContextMenu({ x: p.x, y: p.y }, items);
    });
    return () => {
      offOpen();
      offClose();
      offCtx();
    };
  }, []);

  // window title
  useEffect(() => {
    const t = tabs.find((x) => x.id === activeId);
    api.window.setTitle(t && t.kind !== 'home' ? `${t.dirty ? '• ' : ''}${t.title} — Affice` : 'Affice');
  }, [tabs, activeId]);

  // unsaved-work guard for the browser build
  useEffect(() => {
    if (isElectron) return;
    const h = (e: BeforeUnloadEvent) => {
      if (useWorkspace.getState().tabs.some((t) => t.dirty)) e.preventDefault();
    };
    window.addEventListener('beforeunload', h);
    return () => window.removeEventListener('beforeunload', h);
  }, []);

  // drag & drop files anywhere
  useEffect(() => {
    let depth = 0;
    const hasFiles = (e: DragEvent) => Array.from(e.dataTransfer?.types ?? []).includes('Files');
    const enter = (e: DragEvent) => {
      if (!hasFiles(e)) return;
      depth++;
      setDragging(true);
    };
    const leave = (e: DragEvent) => {
      if (!hasFiles(e)) return;
      depth = Math.max(0, depth - 1);
      if (!depth) setDragging(false);
    };
    const over = (e: DragEvent) => {
      if (hasFiles(e)) e.preventDefault();
    };
    const drop = (e: DragEvent) => {
      depth = 0;
      setDragging(false);
      if (!hasFiles(e)) return;
      // Let editors that accept images handle drops inside themselves
      if ((e.target as HTMLElement | null)?.closest?.('[data-accepts-drop]')) return;
      e.preventDefault();
      if (e.dataTransfer?.files?.length) void openDroppedFiles(e.dataTransfer.files).catch((err) => toast.error(String(err)));
    };
    window.addEventListener('dragenter', enter);
    window.addEventListener('dragleave', leave);
    window.addEventListener('dragover', over);
    window.addEventListener('drop', drop);
    return () => {
      window.removeEventListener('dragenter', enter);
      window.removeEventListener('dragleave', leave);
      window.removeEventListener('dragover', over);
      window.removeEventListener('drop', drop);
    };
  }, []);

  return (
    <div className={`app${loaded ? ' ready' : ''}`}>
      <TitleBar />
      <main className="workspace">
        {tabs.map((t) => (
          <TabView key={t.id} tab={t} active={t.id === activeId} />
        ))}
      </main>
      {dragging && (
        <div className="drop-overlay">
          <div className="drop-card">
            <Upload size={34} />
            <strong>Drop to open</strong>
            <span className="muted">Documents, spreadsheets and presentations</span>
          </div>
        </div>
      )}
      <CommandPalette />
      <DialogHost />
      <ContextMenuHost />
      <ToastHost />
      <TooltipHost />
    </div>
  );
}
