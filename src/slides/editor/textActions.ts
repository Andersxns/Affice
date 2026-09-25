import type { Editor } from '@tiptap/core';
import type { SlidesDoc } from '../doc';
import { findEl, type Bullet, type Para, type Run } from '../model';
import { applyParaStyle, applyRunStyle, selectionRun } from '../ops';
import { effectivePara, effectiveRun, roleOf, type TextContext } from '../render/text';
import { bodyOf } from './TextEditSession';
import type { TextState } from './SlidesRibbon';
import type { SlidesUI } from './ui';

export function activeEditor(ui: SlidesUI): Editor | null {
  const e = ui.textEditor;
  return e && !e.isDestroyed ? e : null;
}

type RunPatch = Partial<Omit<Run, 'text'>>;
const STYLE_KEYS = ['font', 'size', 'color', 'hl', 'caps', 'spacing'] as const;

/** Applies run formatting to the text selection (editing) or to all text of the selected objects. */
export function runStyle(doc: SlidesDoc, ui: SlidesUI, patch: RunPatch): void {
  const ed = activeEditor(ui);
  if (!ed) {
    applyRunStyle(doc, patch);
    return;
  }
  const attrs: Record<string, unknown> = {};
  for (const k of STYLE_KEYS) if (k in patch) attrs[k] = patch[k] ?? null;
  let chain = ed.chain().focus();
  if (Object.keys(attrs).length) chain = chain.setMark('textStyle', attrs).removeEmptyTextStyle();
  const toggles: [keyof RunPatch, string][] = [
    ['b', 'bold'],
    ['i', 'italic'],
    ['u', 'underline'],
    ['s', 'strike'],
    ['sup', 'superscript'],
    ['sub', 'subscript'],
  ];
  for (const [k, mark] of toggles) if (k in patch) chain = patch[k] ? chain.setMark(mark) : chain.unsetMark(mark);
  chain.run();
}

export function toggleMark(doc: SlidesDoc, ui: SlidesUI, mark: 'b' | 'i' | 'u' | 's' | 'sup' | 'sub', state: TextState): void {
  const ed = activeEditor(ui);
  if (ed) {
    const c = ed.chain().focus();
    ({ b: () => c.toggleBold(), i: () => c.toggleItalic(), u: () => c.toggleUnderline(), s: () => c.toggleStrike(), sup: () => c.toggleSuperscript(), sub: () => c.toggleSubscript() })[mark]().run();
    return;
  }
  const on = !state[mark];
  const patch: RunPatch = { [mark]: on || undefined };
  if (on && mark === 'sup') patch.sub = undefined;
  if (on && mark === 'sub') patch.sup = undefined;
  applyRunStyle(doc, patch);
}

export function clearFormatting(doc: SlidesDoc, ui: SlidesUI): void {
  const ed = activeEditor(ui);
  if (ed) {
    ed.chain().focus().unsetAllMarks().run();
    return;
  }
  applyRunStyle(doc, { font: undefined, size: undefined, color: undefined, hl: undefined, b: undefined, i: undefined, u: undefined, s: undefined, sup: undefined, sub: undefined, caps: undefined, spacing: undefined }, 'Clear formatting');
}

function recase(t: string, mode: 'sentence' | 'lower' | 'upper' | 'title' | 'toggle', atStart: boolean): string {
  switch (mode) {
    case 'lower':
      return t.toLowerCase();
    case 'upper':
      return t.toUpperCase();
    case 'title':
      return t.toLowerCase().replace(/(^|[\s(“"'-])(\p{L})/gu, (_m, p: string, c: string) => p + c.toUpperCase());
    case 'toggle':
      return [...t].map((c) => (c === c.toUpperCase() ? c.toLowerCase() : c.toUpperCase())).join('');
    default: {
      const low = t.toLowerCase().replace(/([.!?]\s+)(\p{L})/gu, (_m, p: string, c: string) => p + c.toUpperCase());
      return atStart ? low.replace(/^(\s*)(\p{L})/u, (_m, p: string, c: string) => p + c.toUpperCase()) : low;
    }
  }
}

export function changeCase(doc: SlidesDoc, ui: SlidesUI, mode: 'sentence' | 'lower' | 'upper' | 'title' | 'toggle'): void {
  const ed = activeEditor(ui);
  if (ed) {
    const { state } = ed;
    const { from, to } = state.selection;
    const tr = state.tr;
    state.doc.nodesBetween(from, to, (node, pos) => {
      if (!node.isText || !node.text) return;
      const start = Math.max(from, pos);
      const end = Math.min(to, pos + node.nodeSize);
      if (end <= start) return;
      const slice = node.text.slice(start - pos, end - pos);
      const atStart = start === pos && state.doc.resolve(pos).parentOffset === 0;
      const next = recase(slice, mode, atStart);
      if (next !== slice) tr.replaceWith(tr.mapping.map(start), tr.mapping.map(end), state.schema.text(next, node.marks));
    });
    if (tr.docChanged) ed.view.dispatch(tr);
    ed.commands.focus();
    return;
  }
  doc.updateEls('Change case', doc.sel.els, (e) => {
    const fix = (paras: Para[]) => paras.map((p) => ({ ...p, runs: p.runs.map((r, i) => ({ ...r, text: recase(r.text, mode, i === 0) })) }));
    if (e.type === 'shape' && e.text) return { ...e, text: { ...e.text, paras: fix(e.text.paras) } };
    if (e.type === 'table') return { ...e, rows: e.rows.map((r) => ({ ...r, cells: r.cells.map((c) => ({ ...c, text: { ...c.text, paras: fix(c.text.paras) } })) })) };
    return e;
  });
}

/** Sets paragraph attributes on the paragraphs in the text selection. */
function setParaAttrs(ed: Editor, fn: (attrs: Record<string, unknown>) => Record<string, unknown>): void {
  const { state, view } = ed;
  const { from, to } = state.selection;
  const tr = state.tr;
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (node.type.name !== 'paragraph') return;
    tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...fn(node.attrs) });
  });
  if (tr.docChanged) view.dispatch(tr);
  ed.commands.focus();
}

export function paraStyle(doc: SlidesDoc, ui: SlidesUI, patch: Partial<Omit<Para, 'runs'>> | ((p: Para) => Partial<Omit<Para, 'runs'>>)): void {
  const ed = activeEditor(ui);
  if (!ed) {
    applyParaStyle(doc, patch);
    return;
  }
  setParaAttrs(ed, (attrs) => {
    const cur: Para = { runs: [], level: attrs.level as number, bullet: attrs.bullet as Bullet | undefined, align: attrs.textAlign as Para['align'] };
    const p = typeof patch === 'function' ? patch(cur) : patch;
    const out: Record<string, unknown> = {};
    if ('align' in p) out.textAlign = p.align ?? null;
    if ('level' in p) out.level = p.level ?? 0;
    if ('bullet' in p) out.bullet = p.bullet ?? null;
    if ('lineSpacing' in p) out.lineSpacing = p.lineSpacing ?? null;
    if ('spaceBefore' in p) out.spaceBefore = p.spaceBefore ?? null;
    if ('spaceAfter' in p) out.spaceAfter = p.spaceAfter ?? null;
    if ('marL' in p) out.marL = p.marL ?? null;
    if ('indent' in p) out.indent = p.indent ?? null;
    return out;
  });
}

/** Formatting shown in the ribbon: the text selection when editing, else the first selected object's first run. */
export function readTextState(doc: SlidesDoc, ui: SlidesUI): TextState {
  const ed = activeEditor(ui);
  const target = doc.editing ? findEl(doc.slide.elements, doc.editing.el) : doc.selected[0];
  const body = target ? bodyOf(target, doc.editing?.cell ?? (target.type === 'table' ? [0, 0] : undefined)) : undefined;
  const ctx: TextContext = { pres: doc.pres, role: doc.editing?.cell || target?.type === 'table' ? 'other' : roleOf(target?.ph), ph: target?.ph, slideNumber: doc.sel.slide + 1 };
  if (ed && body) {
    const ts = ed.getAttributes('textStyle');
    const pa = ed.getAttributes('paragraph');
    const para: Para = { runs: [], level: pa.level ?? 0, bullet: pa.bullet ?? undefined };
    const eff = effectiveRun({}, para, body, ctx);
    return {
      editing: true,
      font: ts.font ?? eff.font,
      size: ts.size ?? eff.size,
      color: ts.color ?? undefined,
      hl: ts.hl ?? undefined,
      b: ed.isActive('bold'),
      i: ed.isActive('italic'),
      u: ed.isActive('underline'),
      s: ed.isActive('strike'),
      sup: ed.isActive('superscript'),
      sub: ed.isActive('subscript'),
      align: (pa.textAlign as Para['align']) ?? undefined,
      bullet: effectivePara(para, ctx).bullet,
      level: pa.level ?? 0,
      lineSpacing: pa.lineSpacing ?? undefined,
      anchor: body.anchor ?? 't',
      hasText: true,
    };
  }
  if (!target || !body) return { editing: false, b: false, i: false, u: false, s: false, sup: false, sub: false, level: 0, hasText: false };
  const r = selectionRun(doc);
  const para = r.para ?? { runs: [] };
  const eff = effectiveRun(r, para, body, ctx);
  return {
    editing: false,
    font: eff.font,
    size: eff.size,
    color: r.color,
    hl: r.hl,
    b: !!eff.b,
    i: !!eff.i,
    u: !!eff.u,
    s: !!eff.s,
    sup: !!eff.sup,
    sub: !!eff.sub,
    align: r.align,
    bullet: effectivePara(para, ctx).bullet,
    level: para.level ?? 0,
    lineSpacing: para.lineSpacing,
    anchor: body.anchor ?? 't',
    hasText: true,
  };
}
