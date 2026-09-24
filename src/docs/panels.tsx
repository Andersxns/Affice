import { useEditorState, type Editor } from '@tiptap/react';
import { ArrowDown, ArrowUp, CaseSensitive, Check, CheckCheck, Code2, Heading1, Heading2, Heading3, Image as ImageIcon, List, ListOrdered, ListTodo, MessageSquare, Minus, Quote, Regex, SeparatorHorizontal, Sigma, Table, Trash2, Type, WholeWord, X, ListTree } from 'lucide-react';
import { useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import { createPortal } from 'react-dom';
import { timeAgo } from '@/lib/utils';
import { collectHeadings } from './editor/extensions';
import { getSearchState } from './editor/search';
import type { SlashState } from './editor/slash';
import type { DocComment } from './model';

/* ----------------------------------------------------------- find/replace */

export function FindPanel({ editor, mode, onClose }: { editor: Editor; mode: 'find' | 'replace'; onClose: () => void }) {
  const [query, setQuery] = useState(() => {
    const { from, to, empty } = editor.state.selection;
    return empty ? '' : editor.state.doc.textBetween(from, to, ' ').slice(0, 100);
  });
  const [replace, setReplace] = useState('');
  const [showReplace, setShowReplace] = useState(mode === 'replace');
  const [opts, setOpts] = useState({ caseSensitive: false, wholeWord: false, regex: false });
  const inputRef = useRef<HTMLInputElement>(null);
  const s = useEditorState({ editor, selector: ({ editor: e }) => {
    const st = getSearchState((e as Editor).state);
    return { count: st.results.length, current: st.current };
  } });

  useEffect(() => setShowReplace(mode === 'replace'), [mode]);
  useEffect(() => {
    inputRef.current?.focus();
    inputRef.current?.select();
  }, [mode]);
  useEffect(() => {
    editor.commands.setSearch(query, opts);
  }, [query, opts, editor]);
  useEffect(() => () => void editor.commands.clearSearch(), [editor]);

  const toggle = (k: keyof typeof opts) => setOpts((o) => ({ ...o, [k]: !o[k] }));
  return (
    <div className="find-panel anim-pop" onKeyDown={(e) => e.key === 'Escape' && (onClose(), editor.commands.focus())}>
      <div className="find-row">
        <input
          ref={inputRef}
          className="input input-sm find-input"
          placeholder="Find in document"
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault();
              if (e.shiftKey) editor.commands.findPrev();
              else editor.commands.findNext();
            }
          }}
        />
        <span className="find-count">{query ? (s?.count ? `${(s.current ?? 0) + 1} of ${s.count}` : 'No results') : ''}</span>
        <button className={`icon-btn icon-btn-sm${opts.caseSensitive ? ' active' : ''}`} data-tip="Match case" onClick={() => toggle('caseSensitive')}>
          <CaseSensitive size={15} />
        </button>
        <button className={`icon-btn icon-btn-sm${opts.wholeWord ? ' active' : ''}`} data-tip="Whole words only" onClick={() => toggle('wholeWord')}>
          <WholeWord size={15} />
        </button>
        <button className={`icon-btn icon-btn-sm${opts.regex ? ' active' : ''}`} data-tip="Use regular expressions" onClick={() => toggle('regex')}>
          <Regex size={15} />
        </button>
        <button className="icon-btn icon-btn-sm" data-tip="Previous (Shift+Enter)" onClick={() => editor.commands.findPrev()}>
          <ArrowUp size={15} />
        </button>
        <button className="icon-btn icon-btn-sm" data-tip="Next (Enter)" onClick={() => editor.commands.findNext()}>
          <ArrowDown size={15} />
        </button>
        <button className="icon-btn icon-btn-sm" data-tip="Close (Esc)" onClick={onClose}>
          <X size={15} />
        </button>
      </div>
      {showReplace ? (
        <div className="find-row">
          <input
            className="input input-sm find-input"
            placeholder="Replace with"
            value={replace}
            onChange={(e) => setReplace(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                editor.commands.replaceCurrent(replace);
                editor.commands.findNext();
              }
            }}
          />
          <button className="btn btn-sm" onClick={() => (editor.commands.replaceCurrent(replace), editor.commands.findNext())} disabled={!s?.count}>
            <Check size={14} /> Replace
          </button>
          <button className="btn btn-sm" onClick={() => editor.commands.replaceAll(replace)} disabled={!s?.count}>
            <CheckCheck size={14} /> Replace all
          </button>
        </div>
      ) : (
        <button className="find-more" onClick={() => setShowReplace(true)}>
          Replace…
        </button>
      )}
    </div>
  );
}

/* ------------------------------------------------------------ navigation */

export function NavPane({ editor, headingPages, onClose }: { editor: Editor; headingPages: Map<number, number>; onClose: () => void }) {
  const heads = useEditorState({ editor, selector: ({ editor: e }) => collectHeadings((e as Editor).state.doc, 6) }) ?? [];
  const cursor = useEditorState({ editor, selector: ({ editor: e }) => (e as Editor).state.selection.from }) ?? 0;
  const [q, setQ] = useState('');
  const current = useMemo(() => {
    let idx = -1;
    heads.forEach((h, i) => {
      if (h.pos <= cursor) idx = i;
    });
    return idx;
  }, [heads, cursor]);
  const shown = q ? heads.filter((h) => h.text.toLowerCase().includes(q.toLowerCase())) : heads;
  return (
    <aside className="side-panel nav-pane anim-fade">
      <div className="side-head">
        <ListTree size={16} />
        <strong>Navigation</strong>
        <span className="spacer" />
        <button className="icon-btn icon-btn-sm" onClick={onClose} aria-label="Close navigation">
          <X size={15} />
        </button>
      </div>
      <input className="input input-sm" placeholder="Filter headings" value={q} onChange={(e) => setQ(e.target.value)} style={{ margin: '0 10px 8px' }} />
      <div className="side-body thin-scroll">
        {shown.length === 0 && <p className="muted small side-empty">Headings in your document will appear here. Use the Heading styles to create them.</p>}
        {shown.map((h) => (
          <button
            key={h.pos}
            className={`nav-item level-${h.level}${heads.indexOf(h) === current ? ' active' : ''}`}
            onClick={() => {
              editor.chain().focus().setTextSelection(h.pos + 1).run();
              const dom = editor.view.nodeDOM(h.pos) as HTMLElement | null;
              dom?.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }}
          >
            <span className="ellipsis">{h.text}</span>
            {headingPages.get(h.pos) && <span className="nav-page">{headingPages.get(h.pos)}</span>}
          </button>
        ))}
      </div>
    </aside>
  );
}

/* --------------------------------------------------------------- comments */

export function CommentsPanel({
  editor,
  comments,
  onChange,
  onClose,
  author,
}: {
  editor: Editor;
  comments: DocComment[];
  onChange: (c: DocComment[]) => void;
  onClose: () => void;
  author: string;
}) {
  const activeId = useEditorState({
    editor,
    selector: ({ editor: e }) => ((e as Editor).getAttributes('comment').commentId as string | undefined) ?? null,
  });
  const [reply, setReply] = useState<Record<string, string>>({});
  const focusComment = (id: string) => {
    let found: number | null = null;
    editor.state.doc.descendants((node, pos) => {
      if (found !== null) return false;
      if (node.marks.some((m) => m.type.name === 'comment' && m.attrs.commentId === id)) found = pos;
      return true;
    });
    if (found !== null) {
      editor.chain().focus().setTextSelection(found + 0).run();
      (editor.view.domAtPos(found).node as HTMLElement)?.parentElement?.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
  };
  const remove = (id: string) => {
    editor.commands.unsetComment(id);
    onChange(comments.filter((c) => c.id !== id));
  };
  return (
    <aside className="side-panel comments-pane anim-fade">
      <div className="side-head">
        <MessageSquare size={16} />
        <strong>Comments</strong>
        <span className="badge">{comments.filter((c) => !c.resolved).length}</span>
        <span className="spacer" />
        <button className="icon-btn icon-btn-sm" onClick={onClose} aria-label="Close comments">
          <X size={15} />
        </button>
      </div>
      <div className="side-body thin-scroll">
        {comments.length === 0 && <p className="muted small side-empty">Select some text and choose Review → New comment to start a discussion.</p>}
        {comments.map((c) => (
          <div key={c.id} className={`comment-card${activeId === c.id ? ' active' : ''}${c.resolved ? ' resolved' : ''}`} onClick={() => focusComment(c.id)}>
            <div className="comment-head">
              <span className="comment-avatar">{(c.author || '?').slice(0, 1).toUpperCase()}</span>
              <div className="grow">
                <div className="comment-author">{c.author || 'Anonymous'}</div>
                <div className="comment-date">{timeAgo(new Date(c.date).getTime())}</div>
              </div>
              <button
                className="icon-btn icon-btn-sm"
                data-tip={c.resolved ? 'Reopen' : 'Resolve'}
                onClick={(e) => {
                  e.stopPropagation();
                  onChange(comments.map((x) => (x.id === c.id ? { ...x, resolved: !x.resolved } : x)));
                }}
              >
                <Check size={14} />
              </button>
              <button
                className="icon-btn icon-btn-sm"
                data-tip="Delete comment"
                onClick={(e) => {
                  e.stopPropagation();
                  remove(c.id);
                }}
              >
                <Trash2 size={14} />
              </button>
            </div>
            <div className="comment-text selectable">{c.text}</div>
            {c.replies?.map((r, i) => (
              <div key={i} className="comment-reply">
                <b>{r.author || 'Anonymous'}</b> {r.text}
              </div>
            ))}
            <input
              className="input input-sm comment-reply-input"
              placeholder="Reply…"
              value={reply[c.id] ?? ''}
              onClick={(e) => e.stopPropagation()}
              onChange={(e) => setReply({ ...reply, [c.id]: e.target.value })}
              onKeyDown={(e) => {
                if (e.key === 'Enter' && reply[c.id]?.trim()) {
                  onChange(comments.map((x) => (x.id === c.id ? { ...x, replies: [...(x.replies ?? []), { author, date: new Date().toISOString(), text: reply[c.id].trim() }] } : x)));
                  setReply({ ...reply, [c.id]: '' });
                }
              }}
            />
          </div>
        ))}
      </div>
    </aside>
  );
}

/* ----------------------------------------------------------------- slash */

export interface SlashItem {
  id: string;
  label: string;
  hint: string;
  icon: ReactNode;
  keywords: string;
  run: (editor: Editor) => void;
}

export function slashItems(actions: { insertImage: () => void; insertEquation: () => void }): SlashItem[] {
  return [
    { id: 'text', label: 'Text', hint: 'Plain paragraph', icon: <Type size={16} />, keywords: 'paragraph normal body', run: (e) => e.chain().focus().setParagraph().setStyleName(null).run() },
    { id: 'h1', label: 'Heading 1', hint: 'Large section heading', icon: <Heading1 size={16} />, keywords: 'title h1', run: (e) => e.chain().focus().setHeading({ level: 1 }).run() },
    { id: 'h2', label: 'Heading 2', hint: 'Medium heading', icon: <Heading2 size={16} />, keywords: 'subtitle h2', run: (e) => e.chain().focus().setHeading({ level: 2 }).run() },
    { id: 'h3', label: 'Heading 3', hint: 'Small heading', icon: <Heading3 size={16} />, keywords: 'h3', run: (e) => e.chain().focus().setHeading({ level: 3 }).run() },
    { id: 'bullets', label: 'Bulleted list', hint: 'Simple list', icon: <List size={16} />, keywords: 'ul unordered', run: (e) => e.chain().focus().toggleBulletList().run() },
    { id: 'numbers', label: 'Numbered list', hint: '1, 2, 3…', icon: <ListOrdered size={16} />, keywords: 'ol ordered', run: (e) => e.chain().focus().toggleOrderedList().run() },
    { id: 'todo', label: 'Checklist', hint: 'Tasks with checkboxes', icon: <ListTodo size={16} />, keywords: 'todo task check', run: (e) => e.chain().focus().toggleTaskList().run() },
    { id: 'table', label: 'Table', hint: '3 × 3 table', icon: <Table size={16} />, keywords: 'grid', run: (e) => e.chain().focus().insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run() },
    { id: 'image', label: 'Picture', hint: 'From your computer', icon: <ImageIcon size={16} />, keywords: 'image photo', run: () => actions.insertImage() },
    { id: 'quote', label: 'Quote', hint: 'Block quotation', icon: <Quote size={16} />, keywords: 'blockquote citation', run: (e) => e.chain().focus().toggleBlockquote().run() },
    { id: 'code', label: 'Code block', hint: 'Monospaced code', icon: <Code2 size={16} />, keywords: 'pre program', run: (e) => e.chain().focus().toggleCodeBlock().run() },
    { id: 'math', label: 'Equation', hint: 'LaTeX maths', icon: <Sigma size={16} />, keywords: 'math formula latex', run: () => actions.insertEquation() },
    { id: 'hr', label: 'Divider', hint: 'Horizontal line', icon: <Minus size={16} />, keywords: 'line rule separator', run: (e) => e.chain().focus().setHorizontalRule().run() },
    { id: 'pagebreak', label: 'Page break', hint: 'Continue on next page', icon: <SeparatorHorizontal size={16} />, keywords: 'new page', run: (e) => e.chain().focus().setPageBreak().run() },
    { id: 'toc', label: 'Table of contents', hint: 'From headings', icon: <ListTree size={16} />, keywords: 'toc contents index', run: (e) => e.chain().focus().insertToc().run() },
  ];
}

export function filterSlash(items: SlashItem[], q: string): SlashItem[] {
  const s = q.toLowerCase();
  if (!s) return items;
  return items.filter((i) => i.label.toLowerCase().includes(s) || i.keywords.includes(s));
}

export function SlashMenu({ editor, state, items, index, onPick }: { editor: Editor; state: SlashState; items: SlashItem[]; index: number; onPick: (i: SlashItem) => void }) {
  const listRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    listRef.current?.querySelector('.slash-item.active')?.scrollIntoView({ block: 'nearest' });
  }, [index]);
  if (!state.active || !items.length) return null;
  let coords: { left: number; bottom: number; top: number };
  try {
    coords = editor.view.coordsAtPos(state.from);
  } catch {
    return null;
  }
  const below = coords.bottom + 320 < window.innerHeight;
  return createPortal(
    <div
      ref={listRef}
      className="slash-menu popover thin-scroll"
      style={{ position: 'fixed', left: Math.min(coords.left, window.innerWidth - 280), top: below ? coords.bottom + 6 : undefined, bottom: below ? undefined : window.innerHeight - coords.top + 6, zIndex: 1000 }}
      onMouseDown={(e) => e.preventDefault()}
    >
      <div className="menu-header" style={{ paddingLeft: 10 }}>
        Insert
      </div>
      {items.map((it, i) => (
        <div key={it.id} className={`slash-item${i === index ? ' active' : ''}`} onClick={() => onPick(it)}>
          <span className="slash-icon">{it.icon}</span>
          <span className="grow">
            <div className="slash-label">{it.label}</div>
            <div className="slash-hint">{it.hint}</div>
          </span>
        </div>
      ))}
    </div>,
    document.body,
  );
}
