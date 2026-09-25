/**
 * TipTap configuration for editing slide text in place, and conversion between the slide text model
 * (paragraphs of runs) and ProseMirror documents. Paragraph styling and bullets are applied with node
 * decorations computed by the same code as the static renderer, so editing looks identical to viewing.
 */
import { Extension, Mark, type AnyExtension, type JSONContent } from '@tiptap/core';
import StarterKit from '@tiptap/starter-kit';
import TextAlign from '@tiptap/extension-text-align';
import Subscript from '@tiptap/extension-subscript';
import Superscript from '@tiptap/extension-superscript';
import { TextStyle } from '@tiptap/extension-text-style';
import type { Node as PMNode } from '@tiptap/pm/model';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import { cssFontStack } from '@/lib/fonts';
import { paragraphNumbers, PT, resolveColor, type Bullet, type Para, type Run, type TextBody, type Theme } from '../model';
import { resolveFont } from '../themes';
import { paraCss, type TextContext } from '../render/text';

/* ------------------------------------------------------------ run style */

export interface RunStyleAttrs {
  font: string | null;
  size: number | null;
  color: string | null;
  hl: string | null;
  caps: boolean | null;
  spacing: number | null;
}

export interface EditorEnv {
  theme: () => Theme;
  ctx: () => TextContext;
  body: () => TextBody;
}

/** Adds slide run attributes (theme-aware font/colour, size in pt…) to the textStyle mark. */
const RunStyle = Extension.create<{ env: EditorEnv | null }>({
  name: 'slideRunStyle',
  addOptions() {
    return { env: null };
  },
  addGlobalAttributes() {
    const opts = this.options;
    const themeNow = () => opts.env?.theme() ?? null;
    const attr = (name: keyof RunStyleAttrs, css: (v: never, t: Theme) => string | null, parse: (el: HTMLElement) => unknown) => ({
      [name]: {
        default: null,
        parseHTML: parse,
        renderHTML: (a: Record<string, unknown>) => {
          const v = a[name];
          if (v === null || v === undefined) return {};
          const t = themeNow();
          const style = t ? css(v as never, t) : null;
          return style ? { style } : {};
        },
      },
    });
    return [
      {
        types: ['textStyle'],
        attributes: {
          ...attr('font', (v: string, t) => `font-family: ${cssFontStack(resolveFont(v, t))}`, (el) => el.style.fontFamily?.split(',')[0].replace(/['"]/g, '').trim() || null),
          ...attr('size', (v: number) => `font-size: ${Math.round(v * PT * 100) / 100}px`, (el) => {
            const fs = el.style.fontSize;
            if (!fs) return null;
            const n = parseFloat(fs);
            if (fs.endsWith('pt')) return n;
            if (fs.endsWith('px')) return Math.round((n / PT) * 2) / 2;
            return null;
          }),
          ...attr('color', (v: string, t) => `color: ${resolveColor(v, t)}`, (el) => el.style.color || null),
          ...attr('hl', (v: string, t) => `background-color: ${resolveColor(v, t)}`, (el) => el.style.backgroundColor || null),
          ...attr('caps', (v: boolean) => (v ? 'text-transform: uppercase' : null), (el) => (el.style.textTransform === 'uppercase' ? true : null)),
          ...attr('spacing', (v: number) => `letter-spacing: ${v * PT}px`, () => null),
        },
      },
    ];
  },
});

/** Slide number / date fields. */
const FieldMark = Mark.create({
  name: 'field',
  addAttributes() {
    return { kind: { default: 'slidenum' } };
  },
  parseHTML() {
    return [{ tag: 'span[data-field]', getAttrs: (el) => ({ kind: (el as HTMLElement).getAttribute('data-field') }) }];
  },
  renderHTML({ HTMLAttributes }) {
    return ['span', { 'data-field': HTMLAttributes.kind, class: 'sl-field' }, 0];
  },
});

/* ------------------------------------------------------- paragraph attrs */

const jsonAttr = (name: string) => ({
  default: null,
  parseHTML: (el: HTMLElement) => {
    const v = el.getAttribute(`data-${name}`);
    if (!v) return null;
    try {
      return JSON.parse(v);
    } catch {
      return null;
    }
  },
  renderHTML: () => ({}),
});

const numAttr = () => ({ default: null, renderHTML: () => ({}) });

const ParaAttrs = Extension.create({
  name: 'slideParaAttrs',
  addGlobalAttributes() {
    return [
      {
        types: ['paragraph'],
        attributes: {
          level: { default: 0, renderHTML: () => ({}) },
          bullet: jsonAttr('bullet'),
          endRun: jsonAttr('endrun'),
          lineSpacing: numAttr(),
          spaceBefore: numAttr(),
          spaceAfter: numAttr(),
          marL: numAttr(),
          indent: numAttr(),
        },
      },
    ];
  },
  addKeyboardShortcuts() {
    const shift = (d: number) => () => {
      const { state, dispatch } = this.editor.view;
      const { from, to } = state.selection;
      let tr = state.tr;
      let changed = false;
      state.doc.nodesBetween(from, to, (node, pos) => {
        if (node.type.name !== 'paragraph') return;
        const lvl = Math.max(0, Math.min(8, (node.attrs.level ?? 0) + d));
        if (lvl !== node.attrs.level) {
          tr = tr.setNodeMarkup(pos, undefined, { ...node.attrs, level: lvl, marL: null, indent: null });
          changed = true;
        }
      });
      if (changed) dispatch(tr);
      return true;
    };
    return { Tab: shift(1), 'Shift-Tab': shift(-1) };
  },
});

/* ----------------------------------------------------- decoration plugin */

function cssText(style: Record<string, string | number | undefined>): string {
  return Object.entries(style)
    .filter(([, v]) => v !== undefined && v !== '')
    .map(([k, v]) => `${k.startsWith('--') ? k : k.replace(/[A-Z]/g, (m) => `-${m.toLowerCase()}`)}: ${v}`)
    .join('; ');
}

const decoKey = new PluginKey('slideParaDeco');

function paraDecorations(doc: PMNode, e: EditorEnv | null): DecorationSet {
  if (!e) return DecorationSet.empty;
  const body = e.body();
  const ctx = e.ctx();
  const theme = e.theme();
  const paras: { node: PMNode; pos: number; para: Para }[] = [];
  doc.forEach((node, pos) => {
    if (node.type.name === 'paragraph') paras.push({ node, pos, para: paraFromNode(node) });
  });
  const numbers = paragraphNumbers(paras.map((p) => p.para));
  const decos = paras.map(({ node, pos, para }, i) => {
    const hasText = node.textContent.length > 0;
    const pc = paraCss(para, body, ctx, theme, hasText, numbers[i]);
    const attrs: Record<string, string> = { style: cssText(pc.style), class: 'sl-p' };
    if (pc.bullet !== undefined) attrs['data-bullet'] = pc.bullet;
    return Decoration.node(pos, pos + node.nodeSize, attrs);
  });
  return DecorationSet.create(doc, decos);
}

const ParaDecorations = Extension.create<{ env: EditorEnv | null }>({
  name: 'slideParaDecorations',
  addOptions() {
    return { env: null };
  },
  addProseMirrorPlugins() {
    const e = this.options.env;
    return [
      new Plugin({
        key: decoKey,
        state: {
          init: (_, state) => paraDecorations(state.doc, e),
          apply: (tr, old, _o, state) => (tr.docChanged || tr.getMeta(decoKey) || tr.getMeta('slideRestyle') ? paraDecorations(state.doc, e) : old),
        },
        props: {
          decorations(state) {
            return decoKey.getState(state);
          },
        },
      }),
    ];
  },
});

export function refreshDecorations(view: { state: { tr: { setMeta(k: PluginKey, v: unknown): unknown } }; dispatch(tr: unknown): void }): void {
  view.dispatch(view.state.tr.setMeta(decoKey, true));
}

export function buildSlideTextExtensions(e: EditorEnv): AnyExtension[] {
  return [
    StarterKit.configure({
      blockquote: false,
      bulletList: false,
      orderedList: false,
      listItem: false,
      listKeymap: false,
      code: false,
      codeBlock: false,
      heading: false,
      horizontalRule: false,
      trailingNode: false,
      dropcursor: false,
      gapcursor: false,
      link: { openOnClick: false, autolink: true, linkOnPaste: true, defaultProtocol: 'https', HTMLAttributes: { rel: 'noopener noreferrer nofollow', target: null, class: 'sl-link' } },
      undoRedo: { depth: 200, newGroupDelay: 600 },
    }),
    TextAlign.configure({ types: ['paragraph'], alignments: ['left', 'center', 'right', 'justify'] }),
    Subscript,
    Superscript,
    TextStyle,
    RunStyle.configure({ env: e }),
    FieldMark,
    ParaAttrs,
    ParaDecorations.configure({ env: e }),
  ];
}

/* ----------------------------------------------------------- conversion */

function runMarks(r: Omit<Run, 'text'>): JSONContent['marks'] {
  const marks: NonNullable<JSONContent['marks']> = [];
  if (r.b) marks.push({ type: 'bold' });
  if (r.i) marks.push({ type: 'italic' });
  if (r.u) marks.push({ type: 'underline' });
  if (r.s) marks.push({ type: 'strike' });
  if (r.sup) marks.push({ type: 'superscript' });
  if (r.sub) marks.push({ type: 'subscript' });
  if (r.link) marks.push({ type: 'link', attrs: { href: r.link } });
  if (r.field) marks.push({ type: 'field', attrs: { kind: r.field } });
  const ts: Partial<RunStyleAttrs> = {};
  if (r.font) ts.font = r.font;
  if (r.size) ts.size = r.size;
  if (r.color) ts.color = r.color;
  if (r.hl) ts.hl = r.hl;
  if (r.caps) ts.caps = true;
  if (r.spacing) ts.spacing = r.spacing;
  if (Object.keys(ts).length) marks.push({ type: 'textStyle', attrs: ts });
  return marks.length ? marks : undefined;
}

export function bodyToDoc(body: TextBody): JSONContent {
  const paras = body.paras.length ? body.paras : [{ runs: [] }];
  return {
    type: 'doc',
    content: paras.map((p) => {
      const content: JSONContent[] = [];
      for (const r of p.runs) {
        const marks = runMarks(r);
        const parts = r.text.split('\n');
        parts.forEach((t, i) => {
          if (i > 0) content.push({ type: 'hardBreak', marks });
          if (t) content.push({ type: 'text', text: t, marks });
        });
      }
      return {
        type: 'paragraph',
        attrs: {
          textAlign: p.align ?? null,
          level: p.level ?? 0,
          bullet: p.bullet ?? null,
          endRun: p.endRun ?? null,
          lineSpacing: p.lineSpacing ?? null,
          spaceBefore: p.spaceBefore ?? null,
          spaceAfter: p.spaceAfter ?? null,
          marL: p.marL ?? null,
          indent: p.indent ?? null,
        },
        content: content.length ? content : undefined,
      };
    }),
  };
}

function runFromMarks(marks: readonly { type: { name: string }; attrs: Record<string, unknown> }[]): Omit<Run, 'text'> {
  const r: Omit<Run, 'text'> = {};
  for (const m of marks) {
    switch (m.type.name) {
      case 'bold':
        r.b = true;
        break;
      case 'italic':
        r.i = true;
        break;
      case 'underline':
        r.u = true;
        break;
      case 'strike':
        r.s = true;
        break;
      case 'superscript':
        r.sup = true;
        break;
      case 'subscript':
        r.sub = true;
        break;
      case 'link':
        if (m.attrs.href) r.link = String(m.attrs.href);
        break;
      case 'field':
        r.field = m.attrs.kind === 'date' ? 'date' : 'slidenum';
        break;
      case 'textStyle': {
        const a = m.attrs as unknown as RunStyleAttrs;
        if (a.font) r.font = a.font;
        if (a.size) r.size = a.size;
        if (a.color) r.color = a.color;
        if (a.hl) r.hl = a.hl;
        if (a.caps) r.caps = true;
        if (a.spacing) r.spacing = a.spacing;
        break;
      }
    }
  }
  return r;
}

const sameRun = (a: Omit<Run, 'text'>, b: Omit<Run, 'text'>) => JSON.stringify(a) === JSON.stringify(b);

export function paraFromNode(node: PMNode): Para {
  const a = node.attrs as Record<string, unknown>;
  const runs: Run[] = [];
  node.forEach((child) => {
    const props = runFromMarks(child.marks as never);
    const text = child.type.name === 'hardBreak' ? '\n' : (child.text ?? '');
    if (!text) return;
    const last = runs[runs.length - 1];
    if (last && sameRun({ ...last, text: undefined } as never, { ...props, text: undefined } as never)) last.text += text;
    else runs.push({ ...props, text });
  });
  const p: Para = { runs };
  if (a.textAlign) p.align = a.textAlign as Para['align'];
  if (a.level) p.level = a.level as number;
  if (a.bullet) p.bullet = a.bullet as Bullet;
  if (a.endRun) p.endRun = a.endRun as Para['endRun'];
  if (a.lineSpacing !== null && a.lineSpacing !== undefined) p.lineSpacing = a.lineSpacing as number;
  if (a.spaceBefore !== null && a.spaceBefore !== undefined) p.spaceBefore = a.spaceBefore as number;
  if (a.spaceAfter !== null && a.spaceAfter !== undefined) p.spaceAfter = a.spaceAfter as number;
  if (a.marL !== null && a.marL !== undefined) p.marL = a.marL as number;
  if (a.indent !== null && a.indent !== undefined) p.indent = a.indent as number;
  // an emptied paragraph remembers the formatting of the text it had
  if (!runs.length && !p.endRun) {
    let mk: Omit<Run, 'text'> | null = null;
    node.forEach((c) => {
      if (!mk) mk = runFromMarks(c.marks as never);
    });
    if (mk && Object.keys(mk).length) p.endRun = mk;
  }
  return p;
}

export function docToParas(doc: PMNode): Para[] {
  const paras: Para[] = [];
  doc.forEach((node) => {
    if (node.type.name === 'paragraph') paras.push(paraFromNode(node));
  });
  return paras.length ? paras : [{ runs: [] }];
}
