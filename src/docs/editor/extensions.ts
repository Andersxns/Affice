import { Extension, Mark, Node, mergeAttributes, type CommandProps } from '@tiptap/core';
import { Table, TableCell, TableHeader, TableRow } from '@tiptap/extension-table';
import { Plugin, PluginKey, TextSelection } from '@tiptap/pm/state';
import katex from 'katex';

/* ------------------------------------------------------------ unit helpers */

/** Converts a CSS length (px, pt, in, cm, mm, em) to px. */
export function cssToPx(value: string | null | undefined, emPx = 14.667): number | null {
  if (!value) return null;
  const m = /^(-?[\d.]+)\s*(px|pt|in|cm|mm|em|rem|pc)?$/i.exec(value.trim());
  if (!m) return null;
  const n = parseFloat(m[1]);
  switch ((m[2] ?? 'px').toLowerCase()) {
    case 'pt':
      return (n * 96) / 72;
    case 'in':
      return n * 96;
    case 'cm':
      return (n * 96) / 2.54;
    case 'mm':
      return (n * 96) / 25.4;
    case 'pc':
      return n * 16;
    case 'em':
    case 'rem':
      return n * emPx;
    default:
      return n;
  }
}

export const pxToPtStr = (px: number) => `${Math.round(((px * 72) / 96) * 10) / 10}pt`;

/* -------------------------------------------------------- paragraph format */

export interface ParaAttrs {
  indent: number;
  indentRight: number;
  firstLine: number;
  lineHeight: number | null;
  spaceBefore: number | null;
  spaceAfter: number | null;
  styleName: string | null;
  dir: 'rtl' | 'ltr' | null;
  shading: string | null;
}

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    paragraphFormat: {
      indent: () => ReturnType;
      outdent: () => ReturnType;
      setParagraphFormat: (attrs: Partial<ParaAttrs>) => ReturnType;
      setStyleName: (name: string | null) => ReturnType;
      clearParagraphFormat: () => ReturnType;
    };
    pageBreak: {
      setPageBreak: () => ReturnType;
    };
    docTable: {
      setTableStyle: (style: string | null) => ReturnType;
      setCellBackground: (color: string | null) => ReturnType;
      setCellVAlign: (v: 'top' | 'middle' | 'bottom') => ReturnType;
    };
    comment: {
      setComment: (id: string) => ReturnType;
      unsetComment: (id: string) => ReturnType;
    };
    math: {
      insertMath: (latex: string, display?: boolean) => ReturnType;
    };
    toc: {
      insertToc: () => ReturnType;
    };
  }
}

const TEXT_BLOCKS = ['paragraph', 'heading'];
const INDENT_STEP = 48; // 0.5in

export const ParagraphFormat = Extension.create({
  name: 'paragraphFormat',
  addGlobalAttributes() {
    return [
      {
        types: TEXT_BLOCKS,
        attributes: {
          indent: {
            default: 0,
            parseHTML: (el) => Math.max(0, Math.round(cssToPx((el as HTMLElement).style.marginLeft) ?? 0)),
            renderHTML: (a) => (a.indent ? { style: `margin-left:${a.indent}px` } : {}),
          },
          indentRight: {
            default: 0,
            parseHTML: (el) => Math.max(0, Math.round(cssToPx((el as HTMLElement).style.marginRight) ?? 0)),
            renderHTML: (a) => (a.indentRight ? { style: `margin-right:${a.indentRight}px` } : {}),
          },
          firstLine: {
            default: 0,
            parseHTML: (el) => Math.round(cssToPx((el as HTMLElement).style.textIndent) ?? 0),
            renderHTML: (a) => (a.firstLine ? { style: `text-indent:${a.firstLine}px` } : {}),
          },
          lineHeight: {
            default: null,
            parseHTML: (el) => {
              const v = (el as HTMLElement).style.lineHeight;
              if (!v || v === 'normal') return null;
              const n = parseFloat(v);
              if (/%$/.test(v)) return n / 100;
              if (/^[\d.]+$/.test(v)) return n;
              return null;
            },
            renderHTML: (a) => (a.lineHeight ? { style: `line-height:${a.lineHeight}` } : {}),
          },
          spaceBefore: {
            default: null,
            parseHTML: (el) => {
              const px = cssToPx((el as HTMLElement).style.marginTop);
              return px === null ? null : Math.round(((px * 72) / 96) * 10) / 10;
            },
            renderHTML: (a) => (a.spaceBefore !== null && a.spaceBefore !== undefined ? { style: `margin-top:${a.spaceBefore}pt` } : {}),
          },
          spaceAfter: {
            default: null,
            parseHTML: (el) => {
              const px = cssToPx((el as HTMLElement).style.marginBottom);
              return px === null ? null : Math.round(((px * 72) / 96) * 10) / 10;
            },
            renderHTML: (a) => (a.spaceAfter !== null && a.spaceAfter !== undefined ? { style: `margin-bottom:${a.spaceAfter}pt` } : {}),
          },
          styleName: {
            default: null,
            parseHTML: (el) => (el as HTMLElement).getAttribute('data-style'),
            renderHTML: (a) => (a.styleName ? { 'data-style': a.styleName } : {}),
          },
          dir: {
            default: null,
            parseHTML: (el) => (el as HTMLElement).getAttribute('dir'),
            renderHTML: (a) => (a.dir ? { dir: a.dir } : {}),
          },
          shading: {
            default: null,
            parseHTML: (el) => (el as HTMLElement).getAttribute('data-shading'),
            renderHTML: (a) => (a.shading ? { 'data-shading': a.shading, style: `background-color:${a.shading}` } : {}),
          },
        },
      },
    ];
  },
  addCommands() {
    const eachBlock = (props: CommandProps, fn: (attrs: Record<string, unknown>, type: string) => Record<string, unknown> | null) => {
      const { tr, state, dispatch } = props;
      const { from, to } = state.selection;
      let changed = false;
      state.doc.nodesBetween(from, to, (node, pos) => {
        if (!TEXT_BLOCKS.includes(node.type.name)) return true;
        const next = fn(node.attrs, node.type.name);
        if (next) {
          tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...next });
          changed = true;
        }
        return false;
      });
      if (changed && dispatch) dispatch(tr);
      return changed;
    };
    return {
      indent:
        () =>
        (props) => {
          if (props.editor.isActive('listItem') && props.editor.can().sinkListItem('listItem')) return props.commands.sinkListItem('listItem');
          if (props.editor.isActive('taskItem') && props.editor.can().sinkListItem('taskItem')) return props.commands.sinkListItem('taskItem');
          return eachBlock(props, (a) => ({ indent: Math.min(8 * INDENT_STEP, (Number(a.indent) || 0) + INDENT_STEP) }));
        },
      outdent:
        () =>
        (props) => {
          if (props.editor.isActive('listItem') && props.editor.can().liftListItem('listItem')) return props.commands.liftListItem('listItem');
          if (props.editor.isActive('taskItem') && props.editor.can().liftListItem('taskItem')) return props.commands.liftListItem('taskItem');
          return eachBlock(props, (a) => (Number(a.indent) > 0 ? { indent: Math.max(0, Number(a.indent) - INDENT_STEP) } : null));
        },
      setParagraphFormat:
        (attrs) =>
        (props) =>
          eachBlock(props, () => attrs as Record<string, unknown>),
      setStyleName:
        (name) =>
        (props) =>
          eachBlock(props, () => ({ styleName: name })),
      clearParagraphFormat:
        () =>
        (props) =>
          eachBlock(props, () => ({ indent: 0, indentRight: 0, firstLine: 0, lineHeight: null, spaceBefore: null, spaceAfter: null, shading: null, textAlign: null })),
    };
  },
  addKeyboardShortcuts() {
    return {
      'Mod-]': () => this.editor.commands.indent(),
      'Mod-[': () => this.editor.commands.outdent(),
    };
  },
});

/* ------------------------------------------------------------ page break */

export const PageBreak = Node.create({
  name: 'pageBreak',
  group: 'block',
  atom: true,
  selectable: true,
  draggable: false,
  parseHTML() {
    return [
      { tag: 'div[data-page-break]' },
      { tag: 'div.page-break' },
      { tag: 'hr[data-page-break]' },
      {
        tag: 'br',
        getAttrs: (el) => (/page-break-before:\s*always|break-before:\s*page/i.test((el as HTMLElement).getAttribute('style') ?? '') ? {} : false),
        priority: 60,
      },
    ];
  },
  renderHTML() {
    return ['div', { class: 'page-break', 'data-page-break': 'true' }];
  },
  addCommands() {
    return {
      setPageBreak:
        () =>
        ({ chain }) =>
          chain().insertContent([{ type: 'pageBreak' }, { type: 'paragraph' }]).run(),
    };
  },
  addKeyboardShortcuts() {
    return { 'Mod-Enter': () => this.editor.commands.setPageBreak() };
  },
});

/* --------------------------------------------------------------- tab key */

/** Tab inserts a tab character (lists and tables handle Tab themselves first). */
export const TabKey = Extension.create({
  name: 'tabKey',
  priority: 50,
  addKeyboardShortcuts() {
    return {
      Tab: () => {
        const { editor } = this;
        if (editor.isActive('listItem') || editor.isActive('taskItem') || editor.isActive('table')) return false;
        return editor.commands.insertContent('\t');
      },
      'Shift-Tab': () => {
        const { editor } = this;
        if (editor.isActive('listItem') || editor.isActive('taskItem') || editor.isActive('table')) return false;
        return editor.commands.outdent();
      },
    };
  },
});

/* ----------------------------------------------------------------- tables */

export const DocTable = Table.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      styleName: {
        default: null,
        parseHTML: (el) => (el as HTMLElement).getAttribute('data-style'),
        renderHTML: (a) => (a.styleName ? { 'data-style': a.styleName } : {}),
      },
    };
  },
  addCommands() {
    return {
      ...this.parent?.(),
      setTableStyle:
        (style: string | null) =>
        ({ commands }: CommandProps) =>
          commands.updateAttributes('table', { styleName: style }),
      setCellBackground:
        (color: string | null) =>
        ({ commands }: CommandProps) =>
          commands.setCellAttribute('backgroundColor', color),
      setCellVAlign:
        (v: 'top' | 'middle' | 'bottom') =>
        ({ commands }: CommandProps) =>
          commands.setCellAttribute('valign', v),
    };
  },
});

const cellAttrs = {
  backgroundColor: {
    default: null,
    parseHTML: (el: HTMLElement) => el.getAttribute('data-bg') || el.style.backgroundColor || el.getAttribute('bgcolor') || null,
    renderHTML: (a: Record<string, unknown>) => (a.backgroundColor ? { 'data-bg': a.backgroundColor, style: `background-color:${a.backgroundColor}` } : {}),
  },
  valign: {
    default: null,
    parseHTML: (el: HTMLElement) => {
      const v = el.getAttribute('data-valign') || el.style.verticalAlign || el.getAttribute('valign');
      return v === 'center' ? 'middle' : v || null;
    },
    renderHTML: (a: Record<string, unknown>) => (a.valign && a.valign !== 'top' ? { 'data-valign': a.valign } : {}),
  },
};

export const DocTableCell = TableCell.extend({
  addAttributes() {
    return { ...this.parent?.(), ...cellAttrs };
  },
});

export const DocTableHeader = TableHeader.extend({
  addAttributes() {
    return { ...this.parent?.(), ...cellAttrs };
  },
});

export { TableRow };

/* --------------------------------------------------------------- comments */

export const CommentMark = Mark.create({
  name: 'comment',
  inclusive: false,
  excludes: '',
  addAttributes() {
    return {
      commentId: {
        default: null,
        parseHTML: (el) => (el as HTMLElement).getAttribute('data-comment'),
        renderHTML: (a) => ({ 'data-comment': a.commentId }),
      },
    };
  },
  parseHTML() {
    return [{ tag: 'span[data-comment]' }];
  },
  renderHTML({ HTMLAttributes }) {
    return ['span', mergeAttributes(HTMLAttributes), 0];
  },
  addCommands() {
    return {
      setComment:
        (id: string) =>
        ({ commands }) =>
          commands.setMark(this.name, { commentId: id }),
      unsetComment:
        (id: string) =>
        ({ tr, state, dispatch }) => {
          const type = state.schema.marks.comment;
          state.doc.descendants((node, pos) => {
            node.marks.forEach((m) => {
              if (m.type === type && m.attrs.commentId === id) tr.removeMark(pos, pos + node.nodeSize, m);
            });
          });
          dispatch?.(tr);
          return true;
        },
    };
  },
});

/* ------------------------------------------------------------------- math */

function renderKatex(el: HTMLElement, latex: string, display: boolean) {
  try {
    katex.render(latex || '\\square', el, { displayMode: display, throwOnError: false, output: 'html' });
  } catch {
    el.textContent = latex;
  }
}

export const MathInline = Node.create({
  name: 'mathInline',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: true,
  addAttributes() {
    return { latex: { default: 'x^2' } };
  },
  parseHTML() {
    return [{ tag: 'span[data-math-inline]', getAttrs: (el) => ({ latex: (el as HTMLElement).getAttribute('data-latex') ?? '' }) }];
  },
  renderHTML({ node }) {
    return ['span', { 'data-math-inline': '', 'data-latex': node.attrs.latex, class: 'math-node' }, node.attrs.latex];
  },
  addNodeView() {
    return ({ node }) => {
      const dom = document.createElement('span');
      dom.className = 'math-node';
      dom.setAttribute('data-math-inline', '');
      dom.setAttribute('data-latex', node.attrs.latex);
      renderKatex(dom, node.attrs.latex, false);
      return {
        dom,
        update: (n) => {
          if (n.type.name !== 'mathInline') return false;
          dom.setAttribute('data-latex', n.attrs.latex);
          renderKatex(dom, n.attrs.latex, false);
          return true;
        },
      };
    };
  },
  addCommands() {
    return {
      insertMath:
        (latex: string, display = false) =>
        ({ commands }) =>
          commands.insertContent({ type: display ? 'mathBlock' : 'mathInline', attrs: { latex } }),
    };
  },
});

export const MathBlock = Node.create({
  name: 'mathBlock',
  group: 'block',
  atom: true,
  selectable: true,
  addAttributes() {
    return { latex: { default: 'E = mc^2' } };
  },
  parseHTML() {
    return [{ tag: 'div[data-math-block]', getAttrs: (el) => ({ latex: (el as HTMLElement).getAttribute('data-latex') ?? '' }) }];
  },
  renderHTML({ node }) {
    return ['div', { 'data-math-block': '', 'data-latex': node.attrs.latex, class: 'math-block' }, node.attrs.latex];
  },
  addNodeView() {
    return ({ node }) => {
      const dom = document.createElement('div');
      dom.className = 'math-block';
      dom.setAttribute('data-math-block', '');
      renderKatex(dom, node.attrs.latex, true);
      return {
        dom,
        update: (n) => {
          if (n.type.name !== 'mathBlock') return false;
          renderKatex(dom, n.attrs.latex, true);
          return true;
        },
      };
    };
  },
});

export function renderMathToHtml(latex: string, display: boolean): string {
  try {
    return katex.renderToString(latex, { displayMode: display, throwOnError: false, output: 'html' });
  } catch {
    return latex;
  }
}

/* ------------------------------------------------------- table of contents */

export interface TocEntry {
  level: number;
  text: string;
  pos: number;
  page?: number;
}

export function collectHeadings(doc: import('@tiptap/pm/model').Node, maxLevel = 3): TocEntry[] {
  const out: TocEntry[] = [];
  doc.descendants((node, pos) => {
    if (node.type.name === 'heading') {
      if (node.attrs.level <= maxLevel && node.textContent.trim()) out.push({ level: node.attrs.level, text: node.textContent.trim(), pos });
      return false;
    }
    return node.type.name !== 'table';
  });
  return out;
}

/** Page lookup is injected by the pagination plugin so the TOC can show page numbers. */
export const tocPageLookup: { fn: ((pos: number) => number | null) | null } = { fn: null };

export const TableOfContents = Node.create({
  name: 'tableOfContents',
  group: 'block',
  atom: true,
  selectable: true,
  addAttributes() {
    return { title: { default: 'Contents' }, levels: { default: 3 } };
  },
  parseHTML() {
    return [{ tag: 'div[data-toc]', getAttrs: (el) => ({ title: (el as HTMLElement).getAttribute('data-title') || 'Contents' }) }];
  },
  renderHTML({ node }) {
    return ['div', { 'data-toc': '', 'data-title': node.attrs.title, class: 'toc' }];
  },
  addNodeView() {
    return ({ node, editor }) => {
      const dom = document.createElement('div');
      dom.className = 'toc';
      dom.setAttribute('data-toc', '');
      dom.contentEditable = 'false';
      let current = node;
      const render = () => {
        const heads = collectHeadings(editor.state.doc, current.attrs.levels);
        const title = `<div class="toc-title">${escape(current.attrs.title)}</div>`;
        const rows = heads.length
          ? heads
              .map((h) => {
                const page = tocPageLookup.fn?.(h.pos);
                return `<a class="toc-entry" data-level="${h.level}" data-pos="${h.pos}"><span>${escape(h.text)}</span><span class="toc-dots"></span><span>${page ?? ''}</span></a>`;
              })
              .join('')
          : '<div class="toc-empty">Add headings (Heading 1–3) and they will appear here.</div>';
        dom.innerHTML = title + rows;
      };
      render();
      const onUpdate = () => render();
      editor.on('update', onUpdate);
      editor.on('paginate' as never, onUpdate);
      dom.addEventListener('mousedown', (e) => {
        const a = (e.target as HTMLElement).closest('.toc-entry') as HTMLElement | null;
        if (!a) return;
        e.preventDefault();
        const pos = Number(a.dataset.pos);
        editor.chain().focus().setTextSelection(pos + 1).scrollIntoView().run();
      });
      return {
        dom,
        update: (n) => {
          if (n.type.name !== 'tableOfContents') return false;
          current = n;
          render();
          return true;
        },
        destroy: () => {
          editor.off('update', onUpdate);
          editor.off('paginate' as never, onUpdate);
        },
        ignoreMutation: () => true,
      };
    };
  },
  addCommands() {
    return {
      insertToc:
        () =>
        ({ commands }) =>
          commands.insertContent([{ type: 'tableOfContents' }, { type: 'paragraph' }]),
    };
  },
});

function escape(s: string) {
  return s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!);
}

/* ------------------------------------------------------------ link clicks */

/** Ctrl/Cmd+click opens links (plain click just places the cursor, like Word). */
export function linkClickPlugin(open: (href: string) => void) {
  return new Plugin({
    key: new PluginKey('linkClick'),
    props: {
      handleClick(_view, _pos, event) {
        const a = (event.target as HTMLElement).closest?.('a[href]') as HTMLAnchorElement | null;
        if (a && (event.ctrlKey || event.metaKey)) {
          open(a.href);
          return true;
        }
        return false;
      },
    },
  });
}

export const LinkClick = Extension.create<{ open: (href: string) => void }>({
  name: 'linkClick',
  addOptions() {
    return { open: () => undefined };
  },
  addProseMirrorPlugins() {
    return [linkClickPlugin(this.options.open)];
  },
});

/** Selects the word under the cursor (used before applying formatting with no selection). */
export function selectWordAtCursor(editor: import('@tiptap/core').Editor): boolean {
  const { state } = editor;
  const { $from, empty } = state.selection;
  if (!empty) return false;
  const text = $from.parent.textContent;
  const off = $from.parentOffset;
  let s = off;
  let e = off;
  while (s > 0 && /[\p{L}\p{N}_'’-]/u.test(text[s - 1])) s--;
  while (e < text.length && /[\p{L}\p{N}_'’-]/u.test(text[e])) e++;
  if (s === e) return false;
  const start = $from.start();
  editor.view.dispatch(state.tr.setSelection(TextSelection.create(state.doc, start + s, start + e)));
  return true;
}
