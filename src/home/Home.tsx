import {
  Clock,
  FileQuestion,
  FolderOpen,
  House,
  Info,
  Keyboard,
  LayoutTemplate,
  MoreHorizontal,
  Pin,
  PinOff,
  RotateCcw,
  Search,
  Settings,
  Sparkles,
  Trash2,
  X,
} from 'lucide-react';
import { useEffect, useMemo, useRef, useState } from 'react';
import type { AppKind, RecoveryEntry, RecentFile } from '@/shared/types';
import { api, isElectron } from '@/lib/platform';
import { KIND_APP, KIND_LABEL } from '@/lib/formats';
import { dirname, formatBytes, timeAgo } from '@/lib/utils';
import { TEMPLATES, type TemplateMeta } from '@/templates/catalog';
import { Segmented } from '@/ui/controls';
import { confirmDialog } from '@/ui/dialog';
import { Menu, openContextMenu } from '@/ui/menu';
import { newDocument, openPaths, showOpenDialog, useRecent } from '@/app/actions';
import { AppIcon, GithubIcon, logoMark } from '@/app/AppIcon';
import { openAbout, openShortcuts, REPO_URL } from '@/app/AboutDialog';
import { openSettings } from '@/app/SettingsDialog';
import { useWorkspace } from '@/app/workspace';
import { TemplatePreview } from './TemplatePreview';

function greeting(): string {
  const h = new Date().getHours();
  if (h < 5) return 'Working late';
  if (h < 12) return 'Good morning';
  if (h < 18) return 'Good afternoon';
  return 'Good evening';
}

type Filter = 'all' | AppKind | 'pinned';

export function Home() {
  const recent = useRecent((s) => s.items);
  const [filter, setFilter] = useState<Filter>('all');
  const [tplFilter, setTplFilter] = useState<'all' | AppKind>('all');
  const [query, setQuery] = useState('');
  const [recovery, setRecovery] = useState<RecoveryEntry[]>([]);
  const tabs = useWorkspace((s) => s.tabs);
  const templatesRef = useRef<HTMLDivElement>(null);
  const recentRef = useRef<HTMLDivElement>(null);
  const scrollRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    void useRecent.getState().refresh();
    void api.recovery.list().then((list) => {
      const open = new Set(useWorkspace.getState().tabs.map((t) => t.recoveryId));
      setRecovery(list.filter((r) => !open.has(r.id)));
    });
  }, []);

  const shownRecent = useMemo(() => {
    const q = query.trim().toLowerCase();
    return recent
      .filter((r) => (filter === 'all' ? true : filter === 'pinned' ? r.pinned : r.kind === filter))
      .filter((r) => !q || r.name.toLowerCase().includes(q) || r.path.toLowerCase().includes(q))
      .sort((a, b) => Number(Boolean(b.pinned)) - Number(Boolean(a.pinned)) || b.openedAt - a.openedAt);
  }, [recent, filter, query]);

  const shownTemplates = TEMPLATES.filter((t) => tplFilter === 'all' || t.kind === tplFilter);

  const restore = async (r: RecoveryEntry) => {
    const data = await api.recovery.load(r.id);
    if (!data) return;
    useWorkspace.getState().openTab({ kind: r.kind, title: r.title, path: r.path, source: { type: 'recovery', id: r.id, data }, recoveryId: r.id });
    setRecovery((list) => list.filter((x) => x.id !== r.id));
  };

  const discard = async (r: RecoveryEntry) => {
    await api.recovery.remove(r.id);
    setRecovery((list) => list.filter((x) => x.id !== r.id));
  };

  const scrollTo = (el: HTMLElement | null) => el?.scrollIntoView({ behavior: 'smooth', block: 'start' });

  return (
    <div className="home">
      <aside className="home-nav">
        <div className="home-brand">
          <img src={logoMark} alt="" width={30} height={30} draggable={false} />
          <span>Affice</span>
        </div>
        <button type="button" className="home-nav-item active" onClick={() => scrollRef.current?.scrollTo({ top: 0, behavior: 'smooth' })}>
          <House size={17} /> Home
        </button>
        <button type="button" className="home-nav-item" onClick={() => void showOpenDialog()}>
          <FolderOpen size={17} /> Open
        </button>
        <button type="button" className="home-nav-item" onClick={() => scrollTo(templatesRef.current)}>
          <LayoutTemplate size={17} /> Templates
        </button>
        <button type="button" className="home-nav-item" onClick={() => scrollTo(recentRef.current)}>
          <Clock size={17} /> Recent
        </button>
        <span className="spacer" />
        <button type="button" className="home-nav-item" onClick={() => openShortcuts()}>
          <Keyboard size={17} /> Shortcuts
        </button>
        <button type="button" className="home-nav-item" onClick={() => openSettings()}>
          <Settings size={17} /> Settings
        </button>
        <button type="button" className="home-nav-item" onClick={() => openAbout()}>
          <Info size={17} /> About
        </button>
        <button type="button" className="home-nav-item subtle" onClick={() => api.app.openExternal(REPO_URL)}>
          <GithubIcon size={16} /> Open source
        </button>
      </aside>

      <div className="home-main thin-scroll" ref={scrollRef}>
        <div className="home-inner">
          <header className="home-hero anim-rise">
            <div>
              <h1 className="home-greeting">{greeting()}</h1>
              <p className="home-sub">Create something new, pick a template, or continue where you left off.</p>
            </div>
            <label className="home-search">
              <Search size={16} />
              <input placeholder="Search recent files" value={query} onChange={(e) => setQuery(e.target.value)} />
              {query && (
                <button type="button" className="icon-btn icon-btn-sm" onClick={() => setQuery('')} aria-label="Clear search">
                  <X size={14} />
                </button>
              )}
            </label>
          </header>

          {recovery.length > 0 && (
            <section className="recovery-card anim-rise">
              <div className="recovery-head">
                <RotateCcw size={18} />
                <div>
                  <strong>We saved your unsaved work</strong>
                  <p className="muted small">Affice closed before these files were saved. Restore them to keep working.</p>
                </div>
              </div>
              {recovery.map((r) => (
                <div key={r.id} className="recovery-row">
                  <AppIcon kind={r.kind} size={22} />
                  <div className="grow">
                    <div className="ellipsis">{r.title}</div>
                    <div className="muted small">Saved {timeAgo(r.savedAt)}</div>
                  </div>
                  <button type="button" className="btn btn-sm" onClick={() => void discard(r)}>
                    Discard
                  </button>
                  <button type="button" className="btn btn-sm btn-primary" onClick={() => void restore(r)}>
                    Restore
                  </button>
                </div>
              ))}
            </section>
          )}

          <section className="create-row">
            {(['doc', 'sheet', 'slides'] as AppKind[]).map((k, i) => (
              <button key={k} type="button" className={`create-card kind-${k}`} style={{ animationDelay: `${60 + i * 60}ms` }} onClick={() => newDocument(k)}>
                <div className="create-art">
                  <CreateArt kind={k} />
                </div>
                <div className="create-text">
                  <AppIcon kind={k} size={22} />
                  <div>
                    <div className="create-title">{KIND_APP[k]}</div>
                    <div className="create-sub">Blank {KIND_LABEL[k].toLowerCase()}</div>
                  </div>
                </div>
              </button>
            ))}
            <button type="button" className="create-card create-open" style={{ animationDelay: '240ms' }} onClick={() => void showOpenDialog()}>
              <div className="create-art open-art">
                <FolderOpen size={40} strokeWidth={1.5} />
                <span className="open-formats">docx · xlsx · pptx · odt · ods · rtf · csv · md …</span>
              </div>
              <div className="create-text">
                <span className="open-icon">
                  <FolderOpen size={14} />
                </span>
                <div>
                  <div className="create-title">Open</div>
                  <div className="create-sub">From this computer</div>
                </div>
              </div>
            </button>
          </section>

          <section className="home-section" ref={templatesRef}>
            <div className="section-head">
              <h2>
                <Sparkles size={17} /> Templates
              </h2>
              <Segmented
                size="sm"
                value={tplFilter}
                onChange={setTplFilter}
                options={[
                  { value: 'all', label: 'All' },
                  { value: 'doc', label: 'Documents' },
                  { value: 'sheet', label: 'Sheets' },
                  { value: 'slides', label: 'Slides' },
                ]}
              />
            </div>
            <div className="template-grid">
              {shownTemplates.map((t, i) => (
                <TemplateCard key={t.id} t={t} delay={i * 30} />
              ))}
            </div>
          </section>

          <section className="home-section" ref={recentRef}>
            <div className="section-head">
              <h2>
                <Clock size={17} /> Recent
              </h2>
              <Segmented
                size="sm"
                value={filter}
                onChange={setFilter}
                options={[
                  { value: 'all', label: 'All' },
                  { value: 'pinned', label: 'Pinned' },
                  { value: 'doc', label: 'Documents' },
                  { value: 'sheet', label: 'Sheets' },
                  { value: 'slides', label: 'Slides' },
                ]}
              />
            </div>
            {shownRecent.length === 0 ? (
              <div className="empty-recent">
                <FileQuestion size={34} strokeWidth={1.4} />
                <div>
                  <strong>{query ? 'No files match your search' : filter === 'pinned' ? 'No pinned files yet' : 'No recent files yet'}</strong>
                  <p className="muted">
                    {isElectron ? 'Files you open or save appear here. Drag files anywhere onto this window to open them.' : 'Open a file or drag one onto this window.'}
                  </p>
                </div>
              </div>
            ) : (
              <div className="recent-list">
                <div className="recent-row recent-header">
                  <span />
                  <span>Name</span>
                  <span>Location</span>
                  <span>Opened</span>
                  <span />
                </div>
                {shownRecent.map((r) => (
                  <RecentRow key={r.path} r={r} open={tabs.some((t) => t.path === r.path)} />
                ))}
              </div>
            )}
          </section>
          <footer className="home-footer muted small">
            Affice is free and open-source software. No accounts, no ads, no tracking — your files stay on your computer.
          </footer>
        </div>
      </div>
    </div>
  );
}

function TemplateCard({ t, delay }: { t: TemplateMeta; delay: number }) {
  return (
    <button type="button" className={`template-card kind-${t.kind}`} style={{ animationDelay: `${delay}ms` }} onClick={() => newDocument(t.kind, t.id, t.name)}>
      <div className={`template-thumb thumb-${t.kind}`}>
        <TemplatePreview t={t} />
      </div>
      <div className="template-meta">
        <AppIcon kind={t.kind} size={16} />
        <div className="grow">
          <div className="template-name ellipsis">{t.name}</div>
          <div className="template-desc">{t.description}</div>
        </div>
      </div>
    </button>
  );
}

function RecentRow({ r, open }: { r: RecentFile; open: boolean }) {
  const { pin, remove } = useRecent.getState();
  const [menu, setMenu] = useState(false);
  const btn = useRef<HTMLButtonElement>(null);
  const items = [
    { label: 'Open', onSelect: () => void openPaths([r.path]) },
    { label: r.pinned ? 'Unpin from list' : 'Pin to list', icon: r.pinned ? <PinOff size={15} /> : <Pin size={15} />, onSelect: () => void pin(r.path, !r.pinned) },
    { label: 'Show in folder', icon: <FolderOpen size={15} />, disabled: !isElectron, onSelect: () => api.files.showInFolder(r.path) },
    { separator: true },
    {
      label: 'Remove from list',
      icon: <Trash2 size={15} />,
      onSelect: async () => {
        if (await confirmDialog({ title: 'Remove from recent?', message: `“${r.name}” will be removed from this list. The file itself is not deleted.`, confirmLabel: 'Remove' })) void remove(r.path);
      },
    },
  ];
  return (
    <div
      className="recent-row"
      role="button"
      tabIndex={0}
      onDoubleClick={() => void openPaths([r.path])}
      onClick={() => void openPaths([r.path])}
      onKeyDown={(e) => e.key === 'Enter' && void openPaths([r.path])}
      onContextMenu={(e) => openContextMenu(e, items)}
    >
      <AppIcon kind={r.kind} size={26} />
      <span className="recent-name">
        <span className="ellipsis">{r.name}</span>
        {open && <span className="badge">Open</span>}
        {r.pinned && <Pin size={12} className="recent-pin" />}
      </span>
      <span className="recent-path ellipsis" title={r.path}>
        {dirname(r.path) || '—'}
      </span>
      <span className="recent-when">
        {timeAgo(r.openedAt)}
        {r.size ? <span className="muted"> · {formatBytes(r.size)}</span> : null}
      </span>
      <span className="recent-actions" onClick={(e) => e.stopPropagation()}>
        <button
          type="button"
          className="icon-btn icon-btn-sm"
          data-tip={r.pinned ? 'Unpin' : 'Pin'}
          onClick={() => void pin(r.path, !r.pinned)}
        >
          {r.pinned ? <PinOff size={15} /> : <Pin size={15} />}
        </button>
        <button ref={btn} type="button" className="icon-btn icon-btn-sm" aria-label="More" onClick={() => setMenu((m) => !m)}>
          <MoreHorizontal size={16} />
        </button>
        <Menu open={menu} anchor={btn.current} ignore={[btn.current]} onClose={() => setMenu(false)} items={items} placement="bottom-end" />
      </span>
    </div>
  );
}

function CreateArt({ kind }: { kind: AppKind }) {
  if (kind === 'doc') {
    return (
      <div className="art-doc">
        <div className="art-page">
          <div className="art-h" />
          <div className="art-l w90" />
          <div className="art-l w100" />
          <div className="art-l w70" />
          <div className="art-l w95" />
          <div className="art-l w60" />
        </div>
      </div>
    );
  }
  if (kind === 'sheet') {
    return (
      <div className="art-sheet">
        {Array.from({ length: 20 }, (_, i) => (
          <span key={i} className={i < 4 ? 'hd' : i % 4 === 3 ? 'num' : ''} />
        ))}
        <div className="art-bars">
          <span style={{ height: '40%' }} />
          <span style={{ height: '65%' }} />
          <span style={{ height: '50%' }} />
          <span style={{ height: '85%' }} />
        </div>
      </div>
    );
  }
  return (
    <div className="art-slides">
      <div className="art-slide back" />
      <div className="art-slide">
        <div className="art-title" />
        <div className="art-sub" />
        <div className="art-shape" />
      </div>
    </div>
  );
}
