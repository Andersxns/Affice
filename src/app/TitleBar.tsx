import { Copy, FolderOpen, Plus, Search, X } from 'lucide-react';
import { useEffect, useRef, useState } from 'react';
import { api, isElectron } from '@/lib/platform';
import { Menu, openContextMenu } from '@/ui/menu';
import { useSettings } from './settings';
import { closeTab, newDocument, showOpenDialog } from './actions';
import { AppIcon } from './AppIcon';
import { openCommandPalette } from './CommandPalette';
import { HOME_ID, useWorkspace, type Tab } from './workspace';

export function TitleBar() {
  const tabs = useWorkspace((s) => s.tabs);
  const activeId = useWorkspace((s) => s.activeId);
  const nativeTitleBar = useSettings((s) => s.settings.nativeTitleBar);
  const [newMenu, setNewMenu] = useState(false);
  const plusRef = useRef<HTMLButtonElement>(null);
  const stripRef = useRef<HTMLDivElement>(null);

  // keep the active tab visible when many tabs are open
  useEffect(() => {
    stripRef.current?.querySelector('.tab.active')?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }, [activeId]);

  const overlay = isElectron && !nativeTitleBar && typeof window.affice !== 'undefined';

  return (
    <header className={`titlebar${overlay ? ' with-overlay' : ''}`}>
      <div className="titlebar-tabs" ref={stripRef}>
        {tabs.map((t, i) => (
          <TabButton key={t.id} tab={t} index={i} active={t.id === activeId} />
        ))}
        <button
          ref={plusRef}
          type="button"
          className="tab-new"
          data-tip="New tab"
          data-tip-key="Ctrl+T"
          onClick={() => setNewMenu((o) => !o)}
        >
          <Plus size={16} />
        </button>
        <Menu
          open={newMenu}
          anchor={plusRef.current}
          ignore={[plusRef.current]}
          onClose={() => setNewMenu(false)}
          items={[
            { label: 'Blank document', icon: <AppIcon kind="doc" size={16} />, shortcut: 'Ctrl+N', onSelect: () => newDocument('doc') },
            { label: 'Blank spreadsheet', icon: <AppIcon kind="sheet" size={16} />, onSelect: () => newDocument('sheet') },
            { label: 'Blank presentation', icon: <AppIcon kind="slides" size={16} />, onSelect: () => newDocument('slides') },
            { separator: true },
            { label: 'Open…', icon: <FolderOpen size={16} />, shortcut: 'Ctrl+O', onSelect: () => void showOpenDialog() },
            { label: 'Home', icon: <AppIcon kind="home" size={16} />, onSelect: () => useWorkspace.getState().activate(HOME_ID) },
          ]}
        />
      </div>
      <div className="titlebar-drag" />
      <button type="button" className="titlebar-search" onClick={() => openCommandPalette()} data-tip="Search commands, files and help" data-tip-key="Ctrl+Shift+P">
        <Search size={14} />
        <span>Search</span>
        <span className="kbd">
          <kbd>Ctrl</kbd>
          <kbd>⇧</kbd>
          <kbd>P</kbd>
        </span>
      </button>
      <div className="titlebar-overlay-space" />
    </header>
  );
}

function TabButton({ tab, active, index }: { tab: Tab; active: boolean; index: number }) {
  const activate = useWorkspace((s) => s.activate);
  const moveTab = useWorkspace((s) => s.moveTab);
  const isHome = tab.id === HOME_ID;
  const drag = useRef<{ x: number; started: boolean } | null>(null);

  const onPointerDown = (e: React.PointerEvent) => {
    if (e.button === 1) {
      e.preventDefault();
      return;
    }
    if (e.button !== 0) return;
    activate(tab.id);
    if (isHome) return;
    drag.current = { x: e.clientX, started: false };
    const el = e.currentTarget as HTMLElement;
    const onMove = (ev: PointerEvent) => {
      if (!drag.current) return;
      const dx = ev.clientX - drag.current.x;
      if (!drag.current.started && Math.abs(dx) > 6) {
        drag.current.started = true;
        el.classList.add('dragging');
      }
      if (drag.current.started) {
        el.style.transform = `translateX(${dx}px)`;
        const siblings = Array.from(el.parentElement!.querySelectorAll<HTMLElement>('.tab'));
        const cx = el.getBoundingClientRect().left + el.offsetWidth / 2;
        let target = index;
        siblings.forEach((s, i) => {
          if (s === el) return;
          const r = s.getBoundingClientRect();
          const mid = r.left + r.width / 2;
          if (i > index && cx > mid) target = i;
          if (i < index && cx < mid && i < target) target = i;
        });
        (el as HTMLElement & { _target?: number })._target = target;
      }
    };
    const onUp = () => {
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
      const target = (el as HTMLElement & { _target?: number })._target;
      el.style.transform = '';
      el.classList.remove('dragging');
      if (drag.current?.started && target !== undefined && target !== index) moveTab(tab.id, target);
      drag.current = null;
    };
    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp);
  };

  return (
    <div
      role="tab"
      aria-selected={active}
      className={`tab${active ? ' active' : ''}${isHome ? ' tab-home' : ''}${tab.dirty ? ' dirty' : ''} kind-${tab.kind}`}
      title={tab.path ?? tab.title}
      onPointerDown={onPointerDown}
      onAuxClick={(e) => {
        if (e.button === 1 && !isHome) void closeTab(tab.id);
      }}
      onContextMenu={(e) => {
        if (isHome) return;
        const tabs = useWorkspace.getState().tabs;
        openContextMenu(e, [
          { label: 'Close', shortcut: 'Ctrl+W', onSelect: () => void closeTab(tab.id) },
          {
            label: 'Close other tabs',
            onSelect: async () => {
              for (const t of tabs) if (t.id !== tab.id && t.id !== HOME_ID) if (!(await closeTab(t.id))) break;
            },
          },
          {
            label: 'Close tabs to the right',
            onSelect: async () => {
              const i = tabs.findIndex((t) => t.id === tab.id);
              for (const t of tabs.slice(i + 1)) if (!(await closeTab(t.id))) break;
            },
          },
          { separator: true },
          { label: 'Show in folder', icon: <FolderOpen size={15} />, disabled: !tab.path, onSelect: () => tab.path && api.files.showInFolder(tab.path) },
          { label: 'Copy file path', icon: <Copy size={15} />, disabled: !tab.path, onSelect: () => tab.path && navigator.clipboard.writeText(tab.path) },
        ]);
      }}
    >
      <AppIcon kind={tab.kind} size={16} />
      <span className="tab-title">{isHome ? 'Home' : tab.title}</span>
      {tab.loading && <span className="spinner tab-spinner" />}
      {!isHome && (
        <button
          type="button"
          className="tab-close"
          aria-label={`Close ${tab.title}`}
          onPointerDown={(e) => e.stopPropagation()}
          onClick={(e) => {
            e.stopPropagation();
            void closeTab(tab.id);
          }}
        >
          <span className="tab-dirty-dot" />
          <X size={13} className="tab-x" />
        </button>
      )}
    </div>
  );
}
