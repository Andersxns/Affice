import { CornerDownLeft, FileClock, FolderOpen, Info, Keyboard, Moon, Settings, Sun, Zap } from 'lucide-react';
import { useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import { createPortal } from 'react-dom';
import { create } from 'zustand';
import { timeAgo } from '@/lib/utils';
import { newDocument, openPaths, showOpenDialog, useRecent } from './actions';
import { AppIcon } from './AppIcon';
import { useSettings } from './settings';
import { getController, HOME_ID, useWorkspace, type Command } from './workspace';
import { openSettings } from './SettingsDialog';
import { openAbout, openShortcuts } from './AboutDialog';

const usePalette = create<{ open: boolean; query: string }>(() => ({ open: false, query: '' }));

export function openCommandPalette(query = '') {
  usePalette.setState({ open: true, query });
}

/** Fuzzy subsequence score (higher is better, -1 = no match). */
export function fuzzyScore(query: string, text: string): number {
  if (!query) return 0;
  const q = query.toLowerCase();
  const t = text.toLowerCase();
  const direct = t.indexOf(q);
  if (direct >= 0) return 1000 - direct * 2 - (t.length - q.length) * 0.1 + (direct === 0 ? 200 : 0);
  let ti = 0;
  let score = 0;
  let streak = 0;
  for (let qi = 0; qi < q.length; qi++) {
    const ch = q[qi];
    if (ch === ' ') continue;
    let found = false;
    while (ti < t.length) {
      if (t[ti] === ch) {
        const wordStart = ti === 0 || /[\s\-_/.(]/.test(t[ti - 1]);
        streak += 1;
        score += 5 + streak * 2 + (wordStart ? 12 : 0);
        ti++;
        found = true;
        break;
      }
      streak = 0;
      ti++;
    }
    if (!found) return -1;
  }
  return score;
}

export function globalCommands(): Command[] {
  const s = useSettings.getState();
  const dark = s.settings.theme === 'dark' || (s.settings.theme === 'system' && s.systemDark);
  return [
    { id: 'new-doc', title: 'New document', category: 'File', keys: 'Ctrl+N', icon: <AppIcon kind="doc" size={16} />, run: () => newDocument('doc') },
    { id: 'new-sheet', title: 'New spreadsheet', category: 'File', icon: <AppIcon kind="sheet" size={16} />, run: () => newDocument('sheet') },
    { id: 'new-slides', title: 'New presentation', category: 'File', icon: <AppIcon kind="slides" size={16} />, run: () => newDocument('slides') },
    { id: 'open', title: 'Open file…', category: 'File', keys: 'Ctrl+O', icon: <FolderOpen size={16} />, run: () => void showOpenDialog() },
    { id: 'home', title: 'Go to Home', category: 'View', icon: <AppIcon kind="home" size={16} />, run: () => useWorkspace.getState().activate(HOME_ID) },
    {
      id: 'theme',
      title: dark ? 'Switch to light theme' : 'Switch to dark theme',
      category: 'Appearance',
      icon: dark ? <Sun size={16} /> : <Moon size={16} />,
      run: () => void s.update({ theme: dark ? 'light' : 'dark' }),
    },
    {
      id: 'perf',
      title: s.settings.performanceMode ? 'Turn off performance mode' : 'Turn on performance mode (for older computers)',
      category: 'Appearance',
      icon: <Zap size={16} />,
      run: () => void s.update({ performanceMode: !s.settings.performanceMode }),
    },
    { id: 'settings', title: 'Settings', category: 'App', keys: 'Ctrl+,', icon: <Settings size={16} />, run: () => openSettings() },
    { id: 'shortcuts', title: 'Keyboard shortcuts', category: 'Help', icon: <Keyboard size={16} />, run: () => openShortcuts() },
    { id: 'about', title: 'About Affice', category: 'Help', icon: <Info size={16} />, run: () => openAbout() },
  ];
}

interface Row {
  key: string;
  title: string;
  subtitle?: string;
  category?: string;
  keys?: string;
  icon?: ReactNode;
  run: () => void;
  score: number;
}

export function CommandPalette() {
  const { open, query } = usePalette();
  const [q, setQ] = useState(query);
  const [active, setActive] = useState(0);
  const inputRef = useRef<HTMLInputElement>(null);
  const listRef = useRef<HTMLDivElement>(null);
  const recent = useRecent((s) => s.items);
  const activeId = useWorkspace((s) => s.activeId);

  useEffect(() => {
    if (open) {
      setQ(query);
      setActive(0);
      setTimeout(() => inputRef.current?.focus(), 10);
    }
  }, [open, query]);

  const rows = useMemo<Row[]>(() => {
    if (!open) return [];
    const cmds = [...(getController(activeId)?.commands() ?? []), ...globalCommands()].filter((c) => !c.when || c.when());
    const out: Row[] = [];
    for (const c of cmds) {
      const score = q ? fuzzyScore(q, `${c.title} ${c.category ?? ''}`) : 0;
      if (score < 0) continue;
      out.push({ key: c.id, title: c.title, category: c.category, keys: c.keys, icon: c.icon, run: c.run, score });
    }
    for (const r of recent.slice(0, q ? 40 : 5)) {
      const score = q ? fuzzyScore(q, r.name) : -5;
      if (q && score < 0) continue;
      out.push({
        key: `recent:${r.path}`,
        title: r.name,
        subtitle: `${r.path} · ${timeAgo(r.openedAt)}`,
        category: 'Recent file',
        icon: <FileClock size={16} />,
        run: () => void openPaths([r.path]),
        score: score - 1,
      });
    }
    if (q) out.sort((a, b) => b.score - a.score);
    return out.slice(0, 60);
  }, [open, q, recent, activeId]);

  useEffect(() => {
    listRef.current?.querySelector(`[data-idx="${active}"]`)?.scrollIntoView({ block: 'nearest' });
  }, [active]);

  if (!open) return null;
  const close = () => usePalette.setState({ open: false });
  const run = (r: Row | undefined) => {
    if (!r) return;
    close();
    setTimeout(() => r.run(), 10);
  };

  return createPortal(
    <div className="palette-backdrop" onMouseDown={(e) => e.target === e.currentTarget && close()}>
      <div className="palette" role="dialog" aria-label="Command palette">
        <div className="palette-input-row">
          <Zap size={16} className="palette-input-icon" />
          <input
            ref={inputRef}
            className="palette-input"
            placeholder="What do you want to do? Type a command, feature or file name…"
            value={q}
            onChange={(e) => {
              setQ(e.target.value);
              setActive(0);
            }}
            onKeyDown={(e) => {
              if (e.key === 'Escape') close();
              else if (e.key === 'ArrowDown') {
                e.preventDefault();
                setActive((a) => Math.min(rows.length - 1, a + 1));
              } else if (e.key === 'ArrowUp') {
                e.preventDefault();
                setActive((a) => Math.max(0, a - 1));
              } else if (e.key === 'Enter') {
                e.preventDefault();
                run(rows[active]);
              }
            }}
          />
        </div>
        <div className="palette-list thin-scroll" ref={listRef}>
          {rows.length === 0 && <div className="palette-empty">No matching commands.</div>}
          {rows.map((r, i) => (
            <div
              key={r.key}
              data-idx={i}
              className={`palette-row${i === active ? ' active' : ''}`}
              onMouseMove={() => setActive(i)}
              onClick={() => run(r)}
            >
              <span className="palette-icon">{r.icon}</span>
              <span className="palette-text">
                <span className="palette-title">{r.title}</span>
                {r.subtitle && <span className="palette-sub ellipsis">{r.subtitle}</span>}
              </span>
              {r.category && <span className="palette-cat">{r.category}</span>}
              {r.keys && <span className="palette-keys">{r.keys}</span>}
              {i === active && <CornerDownLeft size={14} className="palette-enter" />}
            </div>
          ))}
        </div>
      </div>
    </div>,
    document.body,
  );
}
