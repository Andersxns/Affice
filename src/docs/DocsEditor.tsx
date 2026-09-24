import 'katex/dist/katex.min.css';
import './content.css';
import './docs.css';
import { Editor } from '@tiptap/core';
import { NodeSelection } from '@tiptap/pm/state';
import type { Mark } from '@tiptap/pm/model';
import {
  AlignCenter,
  Bold,
  FileDown,
  FileText,
  Heading1,
  Image as ImageIcon,
  Italic,
  Link2,
  ListTree,
  MessageSquarePlus,
  Monitor,
  BookOpen,
  Printer,
  Replace,
  Search,
  SeparatorHorizontal,
  Sigma,
  Table as TableIcon,
  Underline,
  Volume2,
  WholeWord,
  Focus,
  SpellCheck,
} from 'lucide-react';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { ContextMenuParams } from '@/shared/types';
import { api, isElectron } from '@/lib/platform';
import { DOC_EXPORT_TARGETS, DOC_SAVE_TARGETS } from '@/lib/formats';
import { uid } from '@/lib/utils';
import { Backstage } from '@/app/Backstage';
import { contextMenuExtender } from '@/app/contextMenu';
import { exportBytes, exportPdf, markDirty, printJob, saveDocument } from '@/app/fileOps';
import { useSettings } from '@/app/settings';
import { SaveState, ZoomControl } from '@/app/StatusBits';
import { registerController, useWorkspace, type Command, type Tab } from '@/app/workspace';
import { openContextMenu, type MenuItem } from '@/ui/menu';
import { promptDialog } from '@/ui/dialog';
import { toast } from '@/ui/toast';
import { buildExtensions } from './editor';
import { useStableEditor } from './useStableEditor';
import { tocPageLookup } from './editor/extensions';
import { readImageFile } from './editor/image';
import { repaginate, setPageGeometry } from './editor/pagination';
import type { SlashState } from './editor/slash';
import { DocPages, PAGE_GAP, type DocView } from './DocPages';
import { DocsRibbon } from './DocsRibbon';
import { equationDialog, headerFooterDialog, linkDialog, pageSetupDialog, paragraphDialog, symbolDialog, wordCountDialog } from './dialogs';
import { buildPdfJob, exportDocument, loadDocument } from './formats';
import { defaultSettings, EMPTY_DOC, pageDims, type DocComment, type DocModel, type DocSettings } from './model';
import { CommentsPanel, filterSlash, FindPanel, NavPane, SlashMenu, slashItems, type SlashItem } from './panels';

export default function DocsEditor({ tab, active }: { tab: Tab; active: boolean }) {
  const appSettings = useSettings((s) => s.settings);
  const updateApp = useSettings((s) => s.update);
  const [settings, setSettings] = useState<DocSettings>(defaultSettings);
  const [comments, setComments] = useState<DocComment[]>([]);
  const [pages, setPages] = useState(1);
  const [headingPages, setHeadingPages] = useState<Map<number, number>>(new Map());
  const [curPage, setCurPage] = useState(1);
  const [zoom, setZoom] = useState(1);
  const [view, setView] = useState<DocView>('print');
  const [backstage, setBackstage] = useState(false);
  const [panel, setPanel] = useState<'none' | 'nav' | 'comments'>('none');
  const [find, setFind] = useState<null | 'find' | 'replace'>(null);
  const [focusMode, setFocusMode] = useState(false);
  const [showMarks, setShowMarks] = useState(false);
  const [darkPage, setDarkPage] = useState(false);
  const [reading, setReading] = useState(false);
  const [painter, setPainter] = useState(false);
  const [words, setWords] = useState(0);
  const [slash, setSlash] = useState<SlashState>({ active: false, from: 0, query: '' });
  const [slashIndex, setSlashIndex] = useState(0);
  const [loaded, setLoaded] = useState(false);

  const scrollRef = useRef<HTMLDivElement>(null);
  const settingsRef = useRef(settings);
  settingsRef.current = settings;
  const commentsRef = useRef(comments);
  commentsRef.current = comments;
  const headingPagesRef = useRef(headingPages);
  headingPagesRef.current = headingPages;
  const slashRef = useRef({ state: slash, index: slashIndex, items: [] as SlashItem[] });
  const painterRef = useRef<{ marks: readonly Mark[]; align: string | null } | null>(null);
  const editorRef = useRef<Editor | null>(null);

  const insertImage = useCallback(async () => {
    const f = await api.files.pickImage();
    const ed = editorRef.current;
    if (!f || !ed) return;
    const blob = new Blob([f.data as BlobPart], { type: f.name.toLowerCase().endsWith('.svg') ? 'image/svg+xml' : `image/${f.name.split('.').pop()?.toLowerCase().replace('jpg', 'jpeg')}` });
    const src = await readImageFile(new File([blob], f.name, { type: blob.type }));
    ed.chain().focus().insertImage({ src, alt: f.name.replace(/\.[^.]+$/, '') }).run();
  }, []);

  const insertEquation = useCallback(() => {
    const ed = editorRef.current;
    if (ed) equationDialog(ed);
  }, []);

  const allSlash = useMemo(() => slashItems({ insertImage: () => void insertImage(), insertEquation }), [insertImage, insertEquation]);
  const shownSlash = useMemo(() => filterSlash(allSlash, slash.query), [allSlash, slash.query]);
  slashRef.current = { state: slash, index: slashIndex, items: shownSlash };

  const pickSlash = useCallback((item: SlashItem) => {
    const ed = editorRef.current;
    if (!ed) return;
    const s = slashRef.current.state;
    const to = ed.state.selection.from;
    ed.chain().focus().deleteRange({ from: s.from, to }).run();
    item.run(ed);
  }, []);

  const editor = useStableEditor(tab.id, () =>
    new Editor({
      extensions: buildExtensions({
        onPaginate: ({ pages: n, headingPages: hp }) => {
          setPages(n);
          setHeadingPages(hp);
        },
        openLink: (href) => api.app.openExternal(href),
        slash: {
          onChange: (s) => {
            setSlash(s);
            setSlashIndex(0);
          },
          onKey: (e) => {
            const { items, index } = slashRef.current;
            if (!items.length) return false;
            if (e.key === 'ArrowDown') {
              setSlashIndex((index + 1) % items.length);
              return true;
            }
            if (e.key === 'ArrowUp') {
              setSlashIndex((index - 1 + items.length) % items.length);
              return true;
            }
            if (e.key === 'Enter' || e.key === 'Tab') {
              pickSlash(items[index]);
              return true;
            }
            return false;
          },
        },
      }),
      content: EMPTY_DOC,
      editorProps: {
        attributes: { class: 'doc-content', spellcheck: 'true', 'aria-label': 'Document' },
      },
      onUpdate: ({ editor: ed }) => {
        markDirty(tab.id);
        setWords(ed.storage.characterCount.words());
      },
    }),
  );
  editorRef.current = editor;

  // expose heading pages to the TOC node view
  useEffect(() => {
    tocPageLookup.fn = (pos) => headingPagesRef.current.get(pos) ?? null;
  }, [headingPages]);

  /* ------------------------------------------------------------ loading */
  useEffect(() => {
    if (!editor) return;
    let cancelled = false;
    (async () => {
      try {
        const model = await loadDocument(tab.source);
        if (cancelled) return;
        setSettings(model.settings);
        setComments(model.comments ?? []);
        editor.chain().setMeta('addToHistory', false).setContent(model.content, { emitUpdate: false }).run();
        setWords(editor.storage.characterCount.words());
        if (model.warning) toast.warning(model.warning, { duration: 8000 });
      } catch (e) {
        console.error(e);
        toast.error(`Couldn't open ${tab.title}`, { detail: String((e as Error)?.message ?? e) });
      } finally {
        if (!cancelled) {
          setLoaded(true);
          useWorkspace.getState().updateTab(tab.id, { loading: false, dirty: tab.source?.type === 'recovery', source: undefined });
          requestAnimationFrame(() => editor.commands.focus('start'));
        }
      }
    })();
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [editor]);

  /* ----------------------------------------------------------- geometry */
  useEffect(() => {
    if (!editor || !loaded) return;
    if (view === 'print') {
      const { h } = pageDims(settings.page);
      setPageGeometry(editor.view, { height: h * 96, gap: PAGE_GAP, marginTop: settings.page.margins.top * 96, marginBottom: settings.page.margins.bottom * 96 });
    } else {
      setPageGeometry(editor.view, null);
      setPages(1);
    }
  }, [editor, loaded, view, settings.page]);

  useEffect(() => {
    if (editor && loaded) repaginate(editor.view);
  }, [editor, loaded, settings.styleSet, settings.styles, active]);

  useEffect(() => {
    if (!editor) return;
    editor.setEditable(view !== 'read');
  }, [editor, view]);

  useEffect(() => {
    editor?.view.dom.setAttribute('spellcheck', String(appSettings.spellcheck));
  }, [editor, appSettings.spellcheck]);

  /* ------------------------------------------------------ current page */
  useEffect(() => {
    if (!editor) return;
    const update = () => {
      if (view !== 'print') return setCurPage(1);
      try {
        const c = editor.view.coordsAtPos(editor.state.selection.head);
        const top = (editor.view.dom as HTMLElement).getBoundingClientRect().top;
        const { h } = pageDims(settingsRef.current.page);
        const S = h * 96 + PAGE_GAP;
        setCurPage(Math.max(1, Math.floor((c.top - top) / zoom / S) + 1));
      } catch {
        /* ignore */
      }
    };
    editor.on('selectionUpdate', update);
    editor.on('paginate' as never, update);
    return () => {
      editor.off('selectionUpdate', update);
      editor.off('paginate' as never, update);
    };
  }, [editor, view, zoom]);

  /* ------------------------------------------------------ settings/model */
  const updateSettings = useCallback(
    (patch: Partial<DocSettings>) => {
      setSettings((s) => ({ ...s, ...patch }));
      markDirty(tab.id);
    },
    [tab.id],
  );

  const getModel = useCallback((): DocModel => {
    const s = settingsRef.current;
    return {
      settings: { ...s, modified: new Date().toISOString(), author: s.author || appSettings.authorName },
      content: editorRef.current?.getJSON() ?? EMPTY_DOC,
      comments: commentsRef.current,
    };
  }, [appSettings.authorName]);

  const serialize = useCallback(
    async (formatId: string) => {
      const ed = editorRef.current!;
      return exportDocument(getModel(), formatId, { html: ed.getHTML(), text: ed.getText({ blockSeparator: '\n\n' }), title: useWorkspace.getState().tabs.find((t) => t.id === tab.id)?.title ?? 'Document', headingPages: headingPagesRef.current });
    },
    [getModel, tab.id],
  );

  const pdfJob = useCallback(() => {
    const ed = editorRef.current!;
    const title = useWorkspace.getState().tabs.find((t) => t.id === tab.id)?.title ?? 'Document';
    return buildPdfJob(getModel(), ed.getHTML(), title, headingPagesRef.current);
  }, [getModel, tab.id]);

  /* ------------------------------------------------------------ actions */
  const addComment = useCallback(async () => {
    const ed = editorRef.current;
    if (!ed) return;
    if (ed.state.selection.empty) {
      toast.info('Select the text you want to comment on first.');
      return;
    }
    const text = await promptDialog({ title: 'New comment', label: 'Comment', multiline: true, placeholder: 'Write your comment…', confirmLabel: 'Post' });
    if (!text?.trim()) return;
    const id = uid('c');
    ed.chain().focus().setComment(id).run();
    setComments((c) => [...c, { id, author: useSettings.getState().settings.authorName || 'Me', date: new Date().toISOString(), text: text.trim() }]);
    setPanel('comments');
    markDirty(tab.id);
  }, [tab.id]);

  const readAloud = useCallback(() => {
    const synth = window.speechSynthesis;
    if (!synth) {
      toast.error('Read aloud is not available on this system.');
      return;
    }
    if (reading) {
      synth.cancel();
      setReading(false);
      return;
    }
    const ed = editorRef.current!;
    const { from, to, empty } = ed.state.selection;
    const text = empty ? ed.state.doc.textBetween(from, ed.state.doc.content.size, '\n', ' ') : ed.state.doc.textBetween(from, to, '\n', ' ');
    if (!text.trim()) return;
    const u = new SpeechSynthesisUtterance(text);
    u.rate = 1;
    u.onend = () => setReading(false);
    u.onerror = () => setReading(false);
    synth.cancel();
    synth.speak(u);
    setReading(true);
    if (!synth.getVoices().length) toast.info('If you hear nothing, your system may not have a text-to-speech voice installed.');
  }, [reading]);

  useEffect(() => () => window.speechSynthesis?.cancel(), []);

  const startPainter = useCallback(() => {
    const ed = editorRef.current!;
    if (painter) {
      setPainter(false);
      painterRef.current = null;
      return;
    }
    const $from = ed.state.selection.$from;
    painterRef.current = { marks: ed.state.storedMarks ?? $from.marks(), align: ($from.parent.attrs.textAlign as string | null) ?? null };
    setPainter(true);
    toast.info('Select text to apply the copied formatting.', { duration: 2500 });
  }, [painter]);

  useEffect(() => {
    if (!editor || !painter) return;
    const onUp = () => {
      const saved = painterRef.current;
      if (!saved) return;
      setTimeout(() => {
        const { from, to, empty } = editor.state.selection;
        if (empty) return;
        let chain = editor.chain().focus().unsetAllMarks();
        for (const m of saved.marks) chain = chain.setMark(m.type.name, m.attrs);
        if (saved.align) chain = chain.setTextAlign(saved.align);
        chain.setTextSelection({ from, to }).run();
        painterRef.current = null;
        setPainter(false);
      }, 0);
    };
    const dom = editor.view.dom;
    dom.addEventListener('mouseup', onUp);
    return () => dom.removeEventListener('mouseup', onUp);
  }, [editor, painter]);

  const toggleSpell = useCallback(() => void updateApp({ spellcheck: !useSettings.getState().settings.spellcheck }), [updateApp]);

  const fitWidth = useCallback(() => {
    const el = scrollRef.current;
    if (!el) return;
    const { w } = pageDims(settingsRef.current.page);
    setZoom(Math.max(0.25, Math.min(4, (el.clientWidth - 64) / (w * 96))));
  }, []);
  const onePage = useCallback(() => {
    const el = scrollRef.current;
    if (!el) return;
    const { h } = pageDims(settingsRef.current.page);
    setZoom(Math.max(0.25, Math.min(4, (el.clientHeight - 48) / (h * 96))));
  }, []);

  /* --------------------------------------------------------- shortcuts */
  useEffect(() => {
    if (!active || !editor) return;
    const onKey = (e: KeyboardEvent) => {
      const mod = e.ctrlKey || e.metaKey;
      const k = e.key.toLowerCase();
      const inEditor = editor.view.dom.contains(document.activeElement);
      if (mod && !e.shiftKey && k === 'f') {
        e.preventDefault();
        setFind('find');
      } else if (mod && !e.shiftKey && k === 'h') {
        e.preventDefault();
        setFind('replace');
      } else if (mod && !e.shiftKey && k === 'k' && inEditor) {
        e.preventDefault();
        linkDialog(editor);
      } else if (e.key === 'F9') {
        e.preventDefault();
        setFocusMode((f) => !f);
      } else if (e.key === 'Escape' && focusMode) {
        setFocusMode(false);
      } else if (e.key === 'Escape' && painter) {
        setPainter(false);
      } else if (inEditor && mod && !e.altKey) {
        // Word-style alignment and formatting shortcuts
        const map: Record<string, () => boolean> = {
          l: () => editor.chain().focus().setTextAlign('left').run(),
          e: () => editor.chain().focus().setTextAlign('center').run(),
          r: () => editor.chain().focus().setTextAlign('right').run(),
          j: () => editor.chain().focus().setTextAlign('justify').run(),
          ' ': () => editor.chain().focus().unsetAllMarks().run(),
          '=': () => editor.chain().focus().toggleSubscript().run(),
        };
        if (!e.shiftKey && map[k]) {
          e.preventDefault();
          map[k]();
        } else if (e.shiftKey && (e.key === '+' || e.key === '=')) {
          e.preventDefault();
          editor.chain().focus().toggleSuperscript().run();
        } else if (e.shiftKey && k === 'v') {
          e.preventDefault();
          api.app.editCommand('pasteAndMatchStyle');
        }
      }
    };
    window.addEventListener('keydown', onKey, true);
    return () => window.removeEventListener('keydown', onKey, true);
  }, [active, editor, focusMode, painter]);

  /* ------------------------------------------------------- context menu */
  const lastCtx = useRef(0);
  const buildContextItems = useCallback((): MenuItem[] => {
    const ed = editorRef.current!;
    const items: MenuItem[] = [
      { label: 'Cut', shortcut: 'Ctrl+X', disabled: ed.state.selection.empty, onSelect: () => (ed.commands.focus(), api.app.editCommand('cut')) },
      { label: 'Copy', shortcut: 'Ctrl+C', disabled: ed.state.selection.empty, onSelect: () => (ed.commands.focus(), api.app.editCommand('copy')) },
      { label: 'Paste', shortcut: 'Ctrl+V', onSelect: () => (ed.commands.focus(), api.app.editCommand('paste')) },
      { label: 'Paste as plain text', shortcut: 'Ctrl+Shift+V', onSelect: () => (ed.commands.focus(), api.app.editCommand('pasteAndMatchStyle')) },
      { separator: true },
    ];
    if (ed.isActive('link')) {
      const href = ed.getAttributes('link').href as string;
      items.push(
        { label: 'Open link', icon: <Link2 size={15} />, onSelect: () => api.app.openExternal(href) },
        { label: 'Edit link…', onSelect: () => linkDialog(ed) },
        { label: 'Remove link', onSelect: () => ed.chain().focus().extendMarkRange('link').unsetLink().run() },
      );
    } else items.push({ label: 'Link…', icon: <Link2 size={15} />, shortcut: 'Ctrl+K', onSelect: () => linkDialog(ed) });
    items.push({ label: 'New comment', icon: <MessageSquarePlus size={15} />, disabled: ed.state.selection.empty, onSelect: () => void addComment() });
    if (ed.isActive('table')) {
      items.push(
        { separator: true },
        {
          label: 'Insert',
          submenu: [
            { label: 'Row above', onSelect: () => ed.chain().focus().addRowBefore().run() },
            { label: 'Row below', onSelect: () => ed.chain().focus().addRowAfter().run() },
            { label: 'Column left', onSelect: () => ed.chain().focus().addColumnBefore().run() },
            { label: 'Column right', onSelect: () => ed.chain().focus().addColumnAfter().run() },
          ],
        },
        {
          label: 'Delete',
          submenu: [
            { label: 'Row', onSelect: () => ed.chain().focus().deleteRow().run() },
            { label: 'Column', onSelect: () => ed.chain().focus().deleteColumn().run() },
            { label: 'Table', danger: true, onSelect: () => ed.chain().focus().deleteTable().run() },
          ],
        },
        { label: 'Merge cells', onSelect: () => ed.chain().focus().mergeCells().run() },
        { label: 'Split cell', onSelect: () => ed.chain().focus().splitCell().run() },
      );
    }
    const sel = ed.state.selection;
    if (sel instanceof NodeSelection && sel.node.type.name === 'image') {
      items.push(
        { separator: true },
        {
          label: 'Wrap text',
          submenu: [
            { label: 'In line with text', onSelect: () => ed.chain().focus().setImageAttrs({ align: 'inline' }).run() },
            { label: 'Image on the left', onSelect: () => ed.chain().focus().setImageAttrs({ align: 'left' }).run() },
            { label: 'Image on the right', onSelect: () => ed.chain().focus().setImageAttrs({ align: 'right' }).run() },
            { label: 'Centred', onSelect: () => ed.chain().focus().setImageAttrs({ align: 'center' }).run() },
          ],
        },
      );
    }
    if (sel instanceof NodeSelection && (sel.node.type.name === 'mathInline' || sel.node.type.name === 'mathBlock')) {
      items.push({ separator: true }, {
        label: 'Edit equation…',
        icon: <Sigma size={15} />,
        onSelect: () =>
          equationDialog(ed, sel.node.attrs.latex, (latex) => {
            ed.chain().focus().command(({ tr }) => {
              tr.setNodeMarkup(sel.from, undefined, { latex });
              return true;
            }).run();
          }),
      });
    }
    items.push({ separator: true }, { label: 'Paragraph…', icon: <AlignCenter size={15} />, onSelect: () => paragraphDialog(ed) });
    return items;
  }, [addComment]);

  useEffect(() => {
    if (!active) return;
    contextMenuExtender.current = (_p: ContextMenuParams) => (Date.now() - lastCtx.current < 800 ? buildContextItems() : undefined);
    return () => {
      contextMenuExtender.current = null;
    };
  }, [active, buildContextItems]);

  useEffect(() => {
    if (!editor) return;
    const dom = editor.view.dom;
    const onCtx = (e: MouseEvent) => {
      lastCtx.current = Date.now();
      if (!isElectron) {
        e.preventDefault();
        openContextMenu(e, buildContextItems());
      }
    };
    // Double-click on an equation edits it
    const onDbl = (e: MouseEvent) => {
      const el = (e.target as HTMLElement).closest('[data-math-inline],[data-math-block]');
      if (!el) return;
      const pos = editor.view.posAtDOM(el, 0);
      const node = editor.state.doc.nodeAt(pos);
      if (!node) return;
      equationDialog(editor, node.attrs.latex, (latex) => {
        editor.chain().focus().command(({ tr }) => {
          tr.setNodeMarkup(pos, undefined, { latex });
          return true;
        }).run();
      });
    };
    dom.addEventListener('contextmenu', onCtx);
    dom.addEventListener('dblclick', onDbl);
    return () => {
      dom.removeEventListener('contextmenu', onCtx);
      dom.removeEventListener('dblclick', onDbl);
    };
  }, [editor, buildContextItems]);

  /* --------------------------------------------------------- controller */
  const commands = useCallback((): Command[] => {
    const ed = editorRef.current!;
    const c = () => ed.chain().focus();
    return [
      { id: 'doc-bold', title: 'Bold', category: 'Format', keys: 'Ctrl+B', icon: <Bold size={16} />, run: () => c().toggleBold().run() },
      { id: 'doc-italic', title: 'Italic', category: 'Format', keys: 'Ctrl+I', icon: <Italic size={16} />, run: () => c().toggleItalic().run() },
      { id: 'doc-underline', title: 'Underline', category: 'Format', keys: 'Ctrl+U', icon: <Underline size={16} />, run: () => c().toggleUnderline().run() },
      ...[1, 2, 3].map((l) => ({ id: `doc-h${l}`, title: `Heading ${l}`, category: 'Styles', keys: `Ctrl+Alt+${l}`, icon: <Heading1 size={16} />, run: () => c().toggleHeading({ level: l as 1 | 2 | 3 }).run() })),
      { id: 'doc-find', title: 'Find', category: 'Edit', keys: 'Ctrl+F', icon: <Search size={16} />, run: () => setFind('find') },
      { id: 'doc-replace', title: 'Replace', category: 'Edit', keys: 'Ctrl+H', icon: <Replace size={16} />, run: () => setFind('replace') },
      { id: 'doc-table', title: 'Insert table', category: 'Insert', icon: <TableIcon size={16} />, run: () => c().insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run() },
      { id: 'doc-image', title: 'Insert picture', category: 'Insert', icon: <ImageIcon size={16} />, run: () => void insertImage() },
      { id: 'doc-link', title: 'Insert link', category: 'Insert', keys: 'Ctrl+K', icon: <Link2 size={16} />, run: () => linkDialog(ed) },
      { id: 'doc-break', title: 'Insert page break', category: 'Insert', keys: 'Ctrl+Enter', icon: <SeparatorHorizontal size={16} />, run: () => c().setPageBreak().run() },
      { id: 'doc-toc', title: 'Insert table of contents', category: 'Insert', icon: <ListTree size={16} />, run: () => c().insertToc().run() },
      { id: 'doc-eq', title: 'Insert equation', category: 'Insert', icon: <Sigma size={16} />, run: insertEquation },
      { id: 'doc-symbol', title: 'Insert symbol or emoji', category: 'Insert', run: () => symbolDialog(ed) },
      { id: 'doc-comment', title: 'New comment', category: 'Review', icon: <MessageSquarePlus size={16} />, run: () => void addComment() },
      { id: 'doc-count', title: 'Word count', category: 'Review', icon: <WholeWord size={16} />, run: () => wordCountDialog(ed, pages) },
      { id: 'doc-read', title: 'Read aloud', category: 'Review', icon: <Volume2 size={16} />, run: readAloud },
      { id: 'doc-spell', title: 'Toggle spell check', category: 'Review', icon: <SpellCheck size={16} />, run: toggleSpell },
      { id: 'doc-pagesetup', title: 'Page setup (size, margins, orientation)', category: 'Layout', icon: <FileText size={16} />, run: () => pageSetupDialog(settingsRef.current, updateSettings) },
      { id: 'doc-hf', title: 'Header & footer / page numbers', category: 'Insert', run: () => headerFooterDialog(settingsRef.current, updateSettings) },
      { id: 'doc-print-layout', title: 'Print layout', category: 'View', icon: <FileText size={16} />, run: () => setView('print') },
      { id: 'doc-web-layout', title: 'Web layout', category: 'View', icon: <Monitor size={16} />, run: () => setView('web') },
      { id: 'doc-read-mode', title: 'Read mode', category: 'View', icon: <BookOpen size={16} />, run: () => setView('read') },
      { id: 'doc-focus', title: 'Focus mode', category: 'View', keys: 'F9', icon: <Focus size={16} />, run: () => setFocusMode((f) => !f) },
      { id: 'doc-nav', title: 'Navigation pane', category: 'View', icon: <ListTree size={16} />, run: () => setPanel((p) => (p === 'nav' ? 'none' : 'nav')) },
      { id: 'doc-pdf', title: 'Export as PDF', category: 'File', icon: <FileDown size={16} />, run: () => void exportPdf(useWorkspace.getState().tabs.find((t) => t.id === tab.id)?.title ?? 'Document', pdfJob()) },
      { id: 'doc-print', title: 'Print', category: 'File', keys: 'Ctrl+P', icon: <Printer size={16} />, run: () => void printJob(pdfJob()) },
    ];
  }, [addComment, insertEquation, insertImage, pages, pdfJob, readAloud, tab.id, toggleSpell, updateSettings]);

  useEffect(() => {
    if (!editor) return;
    return registerController(tab.id, {
      save: () => saveDocument({ tabId: tab.id, kind: 'doc', targets: DOC_SAVE_TARGETS, defaultFormat: useSettings.getState().settings.defaultDocFormat, serialize }),
      saveAs: (formatId) => saveDocument({ tabId: tab.id, kind: 'doc', targets: DOC_SAVE_TARGETS, defaultFormat: useSettings.getState().settings.defaultDocFormat, serialize, saveAs: true, formatId }),
      exportAs: async (id) => {
        const title = useWorkspace.getState().tabs.find((t) => t.id === tab.id)?.title ?? 'Document';
        if (id === 'pdf') return exportPdf(title, pdfJob());
        const target = DOC_EXPORT_TARGETS.find((t) => t.id === id) ?? DOC_SAVE_TARGETS.find((t) => t.id === id);
        if (target) await exportBytes(title, target, () => serialize(id));
      },
      print: () => printJob(pdfJob()),
      snapshot: () => JSON.stringify(getModel()),
      focus: () => editor.commands.focus(),
      commands,
      info: () => [
        { label: 'Pages', value: String(pages) },
        { label: 'Words', value: editor.storage.characterCount.words().toLocaleString() },
        { label: 'Characters', value: editor.storage.characterCount.characters().toLocaleString() },
      ],
    });
  }, [editor, tab.id, serialize, pdfJob, getModel, commands, pages]);

  useEffect(() => {
    if (active && editor && loaded) requestAnimationFrame(() => editor.commands.focus());
  }, [active, editor, loaded]);

  if (!editor) return null;
  const title = tab.title;

  return (
    <div className={`editor docs-editor${focusMode ? ' focus-mode' : ''}${painter ? ' painter-on' : ''}`}>
      {!focusMode && view !== 'read' && (
        <DocsRibbon
          editor={editor}
          settings={settings}
          updateSettings={updateSettings}
          onFile={() => setBackstage(true)}
          openFind={setFind}
          insertImage={() => void insertImage()}
          insertLink={() => linkDialog(editor)}
          addComment={() => void addComment()}
          pageSetup={() => pageSetupDialog(settings, updateSettings)}
          headerFooter={(w) => headerFooterDialog(settings, updateSettings, w)}
          wordCount={() => wordCountDialog(editor, pages)}
          insertSymbol={() => symbolDialog(editor)}
          insertEquation={insertEquation}
          paragraph={() => paragraphDialog(editor)}
          view={view}
          setView={setView}
          zoom={zoom}
          setZoom={setZoom}
          fitWidth={fitWidth}
          onePage={onePage}
          panel={panel}
          togglePanel={(p) => setPanel((cur) => (cur === p ? 'none' : p))}
          showMarks={showMarks}
          setShowMarks={setShowMarks}
          focusMode={focusMode}
          setFocusMode={setFocusMode}
          darkPage={darkPage}
          setDarkPage={setDarkPage}
          readAloud={readAloud}
          reading={reading}
          painter={painter}
          startPainter={startPainter}
          spellcheck={appSettings.spellcheck}
          toggleSpell={toggleSpell}
        />
      )}
      {view === 'read' && (
        <div className="read-bar">
          <BookOpen size={16} />
          <span className="ellipsis">{title}</span>
          <span className="spacer" />
          <button className="btn btn-sm" onClick={() => setView('print')}>
            Edit document
          </button>
        </div>
      )}
      <div className="editor-main">
        {panel === 'nav' && !focusMode && <NavPane editor={editor} headingPages={headingPages} onClose={() => setPanel('none')} />}
        <div
          className="editor-canvas doc-canvas"
          onWheel={(e) => {
            if (e.ctrlKey || e.metaKey) {
              e.preventDefault();
              setZoom((z) => Math.max(0.25, Math.min(4, Math.round((z - Math.sign(e.deltaY) * 0.1) * 100) / 100)));
            }
          }}
        >
          <DocPages
            ref={scrollRef}
            editor={editor}
            settings={settings}
            pages={pages}
            zoom={view === 'read' ? Math.max(zoom, 1.1) : zoom}
            view={view}
            title={title}
            darkPage={darkPage}
            showMarks={showMarks}
            focusMode={focusMode}
            onHeaderFooter={(w) => headerFooterDialog(settings, updateSettings, w)}
          />
          {find && <FindPanel editor={editor} mode={find} onClose={() => setFind(null)} />}
          {focusMode && (
            <button className="focus-exit btn btn-sm" onClick={() => setFocusMode(false)}>
              Exit focus mode (Esc)
            </button>
          )}
        </div>
        {panel === 'comments' && !focusMode && (
          <CommentsPanel
            editor={editor}
            comments={comments}
            author={appSettings.authorName || 'Me'}
            onChange={(c) => {
              setComments(c);
              markDirty(tab.id);
            }}
            onClose={() => setPanel('none')}
          />
        )}
      </div>
      {!focusMode && (
        <footer className="statusbar">
          <span className="status-item">{view === 'print' ? `Page ${Math.min(curPage, pages)} of ${pages}` : view === 'read' ? 'Read mode' : 'Web layout'}</span>
          <span className="status-sep" />
          <button className="status-item" data-tip="Word count details" onClick={() => wordCountDialog(editor, pages)}>
            {words.toLocaleString()} {words === 1 ? 'word' : 'words'}
          </button>
          <span className="status-sep" />
          <button className={`status-item${appSettings.spellcheck ? ' active' : ''}`} data-tip="Spell check" onClick={toggleSpell}>
            <SpellCheck size={13} /> {appSettings.spellcheck ? 'Spelling on' : 'Spelling off'}
          </button>
          {painter && <span className="status-item active">Format painter: select text to apply</span>}
          <span className="spacer" />
          <SaveState dirty={tab.dirty} path={tab.path} />
          <span className="status-sep" />
          <button className={`status-item${view === 'print' ? ' active' : ''}`} data-tip="Print layout" onClick={() => setView('print')}>
            <FileText size={14} />
          </button>
          <button className={`status-item${view === 'web' ? ' active' : ''}`} data-tip="Web layout" onClick={() => setView('web')}>
            <Monitor size={14} />
          </button>
          <button className={`status-item${view === 'read' ? ' active' : ''}`} data-tip="Read mode" onClick={() => setView('read')}>
            <BookOpen size={14} />
          </button>
          <span className="status-sep" />
          <ZoomControl zoom={zoom} onZoom={setZoom} extra={[{ label: 'Page width', onSelect: fitWidth }, { label: 'One page', onSelect: onePage }]} />
        </footer>
      )}
      {backstage && (
        <Backstage
          app="doc"
          tab={tab}
          onClose={() => setBackstage(false)}
          onSave={() => void saveDocument({ tabId: tab.id, kind: 'doc', targets: DOC_SAVE_TARGETS, defaultFormat: appSettings.defaultDocFormat, serialize })}
          onSaveAs={(fid) => void saveDocument({ tabId: tab.id, kind: 'doc', targets: DOC_SAVE_TARGETS, defaultFormat: appSettings.defaultDocFormat, serialize, saveAs: true, formatId: fid })}
          onExport={(id) => {
            if (id === 'pdf') void exportPdf(title, pdfJob());
            else {
              const target = DOC_EXPORT_TARGETS.find((t) => t.id === id);
              if (target) void exportBytes(title, target, () => serialize(id));
            }
          }}
          onPrint={() => void printJob(pdfJob())}
          saveTargets={DOC_SAVE_TARGETS}
          exportTargets={DOC_EXPORT_TARGETS}
          libreOffice={isElectron}
          info={[
            { label: 'Pages', value: String(pages) },
            { label: 'Words', value: words.toLocaleString() },
            { label: 'Author', value: settings.author || appSettings.authorName || '—' },
            { label: 'Created', value: new Date(settings.created).toLocaleString() },
          ]}
        />
      )}
      <SlashMenu editor={editor} state={slash} items={shownSlash} index={slashIndex} onPick={pickSlash} />
    </div>
  );
}
