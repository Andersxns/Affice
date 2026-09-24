import type { JSONContent } from '@tiptap/core';
import { strFromU8, unzipSync } from 'fflate';
import { bytesToDataUrl, mimeFromExt } from '@/lib/utils';
import { defaultSettings, PAPER_SIZES, type DocComment, type DocSettings, type HeaderFooter, type StyleDef, type StyleId } from '../model';
import type { LoadedDoc } from './index';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/* ------------------------------------------------------------------ xml */

function parseXml(s: string | undefined): Document | null {
  if (!s) return null;
  const d = new DOMParser().parseFromString(s, 'application/xml');
  return d.getElementsByTagName('parsererror').length ? null : d;
}

const kids = (el: Element | null | undefined, name?: string): Element[] =>
  el ? Array.from(el.children).filter((c) => !name || c.localName === name) : [];
const kid = (el: Element | null | undefined, name: string): Element | undefined => (el ? Array.from(el.children).find((c) => c.localName === name) : undefined);
const desc = (el: Element | Document | null | undefined, name: string): Element[] => (el ? Array.from(el.getElementsByTagNameNS('*', name)) : []);

function attr(el: Element | null | undefined, name: string): string | null {
  if (!el) return null;
  const v = el.getAttributeNS(W, name) ?? el.getAttribute(`w:${name}`) ?? el.getAttribute(name);
  if (v !== null) return v;
  for (const a of Array.from(el.attributes)) if (a.localName === name) return a.value;
  return null;
}
const rAttr = (el: Element | null | undefined, name: string): string | null => (el ? el.getAttributeNS(R, name) ?? el.getAttribute(`r:${name}`) : null);

/** On/off toggles: <w:b/>, <w:b w:val="true|1|on"/>, off when val is 0/false/off/none. */
function onOff(el: Element | undefined): boolean | undefined {
  if (!el) return undefined;
  const v = attr(el, 'val');
  return !(v === '0' || v === 'false' || v === 'off' || v === 'none');
}

const twipToPx = (t: number) => t / 15;
const twipToPt = (t: number) => t / 20;
const twipToIn = (t: number) => t / 1440;
const emuToPx = (e: number) => e / 9525;

/* ------------------------------------------------------------ properties */

interface RunProps {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  vert?: 'superscript' | 'subscript';
  font?: string;
  size?: number; // pt
  color?: string;
  highlight?: string;
  hidden?: boolean;
  caps?: boolean;
}

interface ParaProps {
  align?: string;
  indent?: number;
  indentRight?: number;
  firstLine?: number;
  before?: number;
  after?: number;
  lineHeight?: number;
  dir?: 'rtl';
  shading?: string;
  numId?: string;
  ilvl?: number;
  pageBreakBefore?: boolean;
  styleId?: string;
}

const HIGHLIGHT: Record<string, string> = {
  yellow: '#ffff00',
  green: '#00ff00',
  cyan: '#00ffff',
  magenta: '#ff00ff',
  blue: '#0000ff',
  red: '#ff0000',
  darkBlue: '#000080',
  darkCyan: '#008080',
  darkGreen: '#008000',
  darkMagenta: '#800080',
  darkRed: '#800000',
  darkYellow: '#808000',
  darkGray: '#808080',
  lightGray: '#c0c0c0',
  black: '#000000',
  white: '#ffffff',
};

interface Ctx {
  files: Record<string, Uint8Array>;
  rels: Map<string, { target: string; external: boolean; type: string }>;
  styles: Map<string, { type: string; name: string; basedOn?: string; rPr?: Element; pPr?: Element }>;
  defaultRun: RunProps;
  defaultPara: ParaProps;
  themeFonts: { major: string; minor: string };
  numbering: Map<string, Map<number, { fmt: string; text: string; start: number; font?: string }>>;
  comments: Map<string, DocComment>;
  footnotes: Map<string, string>;
  footnoteOrder: string[];
  warnings: Set<string>;
  headingStyles: Map<string, number>;
  imageCount: number;
}

function readRunProps(rPr: Element | undefined, ctx: Ctx): RunProps {
  const p: RunProps = {};
  if (!rPr) return p;
  const b = onOff(kid(rPr, 'b'));
  if (b !== undefined) p.bold = b;
  const i = onOff(kid(rPr, 'i'));
  if (i !== undefined) p.italic = i;
  const u = kid(rPr, 'u');
  if (u) p.underline = attr(u, 'val') !== 'none';
  const st = onOff(kid(rPr, 'strike')) || onOff(kid(rPr, 'dstrike'));
  if (st !== undefined) p.strike = st;
  const va = attr(kid(rPr, 'vertAlign'), 'val');
  if (va === 'superscript' || va === 'subscript') p.vert = va;
  const fonts = kid(rPr, 'rFonts');
  if (fonts) {
    const theme = attr(fonts, 'asciiTheme') ?? attr(fonts, 'hAnsiTheme');
    const f = attr(fonts, 'ascii') ?? attr(fonts, 'hAnsi') ?? attr(fonts, 'cs') ?? attr(fonts, 'eastAsia');
    if (f) p.font = f;
    else if (theme) p.font = theme.startsWith('major') ? ctx.themeFonts.major : ctx.themeFonts.minor;
  }
  const sz = attr(kid(rPr, 'sz'), 'val') ?? attr(kid(rPr, 'szCs'), 'val');
  if (sz && Number(sz) > 0) p.size = Number(sz) / 2;
  const color = attr(kid(rPr, 'color'), 'val');
  if (color && color !== 'auto' && /^[0-9a-f]{6}$/i.test(color)) p.color = `#${color.toLowerCase()}`;
  const hl = attr(kid(rPr, 'highlight'), 'val');
  if (hl && hl !== 'none' && HIGHLIGHT[hl]) p.highlight = HIGHLIGHT[hl];
  const shd = attr(kid(rPr, 'shd'), 'fill');
  if (!p.highlight && shd && shd !== 'auto' && /^[0-9a-f]{6}$/i.test(shd) && shd.toLowerCase() !== 'ffffff') p.highlight = `#${shd.toLowerCase()}`;
  if (onOff(kid(rPr, 'vanish'))) p.hidden = true;
  if (onOff(kid(rPr, 'caps'))) p.caps = true;
  return p;
}

function readParaProps(pPr: Element | undefined, fontSizePt = 11): ParaProps {
  const p: ParaProps = {};
  if (!pPr) return p;
  const jc = attr(kid(pPr, 'jc'), 'val');
  if (jc) p.align = jc === 'center' ? 'center' : jc === 'right' || jc === 'end' ? 'right' : jc === 'both' || jc === 'distribute' ? 'justify' : 'left';
  const ind = kid(pPr, 'ind');
  if (ind) {
    const left = attr(ind, 'left') ?? attr(ind, 'start');
    const right = attr(ind, 'right') ?? attr(ind, 'end');
    const first = attr(ind, 'firstLine');
    const hanging = attr(ind, 'hanging');
    if (left) p.indent = Math.max(0, Math.round(twipToPx(Number(left))));
    if (right) p.indentRight = Math.max(0, Math.round(twipToPx(Number(right))));
    if (first) p.firstLine = Math.round(twipToPx(Number(first)));
    if (hanging) p.firstLine = -Math.round(twipToPx(Number(hanging)));
  }
  const sp = kid(pPr, 'spacing');
  if (sp) {
    const before = attr(sp, 'before');
    const after = attr(sp, 'after');
    if (before !== null && attr(sp, 'beforeAutospacing') !== '1') p.before = twipToPt(Number(before));
    if (after !== null && attr(sp, 'afterAutospacing') !== '1') p.after = twipToPt(Number(after));
    const line = attr(sp, 'line');
    const rule = attr(sp, 'lineRule');
    if (line) {
      const n = Number(line);
      if (!rule || rule === 'auto') p.lineHeight = Math.round((n / 240) * 1.0 * 100) / 100;
      else p.lineHeight = Math.max(0.8, Math.round((twipToPt(n) / (fontSizePt * 1.17)) * 100) / 100);
    }
  }
  if (onOff(kid(pPr, 'bidi'))) p.dir = 'rtl';
  const shd = attr(kid(pPr, 'shd'), 'fill');
  if (shd && shd !== 'auto' && /^[0-9a-f]{6}$/i.test(shd) && shd.toLowerCase() !== 'ffffff') p.shading = `#${shd.toLowerCase()}`;
  const numPr = kid(pPr, 'numPr');
  if (numPr) {
    const numId = attr(kid(numPr, 'numId'), 'val');
    if (numId && numId !== '0') {
      p.numId = numId;
      p.ilvl = Number(attr(kid(numPr, 'ilvl'), 'val') ?? 0);
    }
  }
  if (onOff(kid(pPr, 'pageBreakBefore'))) p.pageBreakBefore = true;
  const ps = attr(kid(pPr, 'pStyle'), 'val');
  if (ps) p.styleId = ps;
  return p;
}

/* -------------------------------------------------------------- styles */

function styleChain(ctx: Ctx, id: string | undefined): Array<{ rPr?: Element; pPr?: Element }> {
  const out: Array<{ rPr?: Element; pPr?: Element }> = [];
  const seen = new Set<string>();
  let cur = id;
  while (cur && !seen.has(cur)) {
    seen.add(cur);
    const s = ctx.styles.get(cur);
    if (!s) break;
    out.unshift({ rPr: s.rPr, pPr: s.pPr });
    cur = s.basedOn;
  }
  return out;
}

const runCache = new WeakMap<Ctx, Map<string, RunProps>>();
const paraCache = new WeakMap<Ctx, Map<string, ParaProps>>();

function styleRun(ctx: Ctx, id: string | undefined): RunProps {
  if (!id) return {};
  let cache = runCache.get(ctx);
  if (!cache) runCache.set(ctx, (cache = new Map()));
  let v = cache.get(id);
  if (!v) {
    v = styleChain(ctx, id).reduce<RunProps>((acc, s) => ({ ...acc, ...readRunProps(s.rPr, ctx) }), {});
    cache.set(id, v);
  }
  return v;
}

function stylePara(ctx: Ctx, id: string | undefined): ParaProps {
  if (!id) return {};
  let cache = paraCache.get(ctx);
  if (!cache) paraCache.set(ctx, (cache = new Map()));
  let v = cache.get(id);
  if (!v) {
    v = styleChain(ctx, id).reduce<ParaProps>((acc, s) => ({ ...acc, ...readParaProps(s.pPr) }), {});
    cache.set(id, v);
  }
  return v;
}

/** Maps Word style ids/names to Affice's built-in styles. */
function mapStyle(ctx: Ctx, styleId: string | undefined): { kind: 'heading'; level: number } | { kind: 'para'; styleName: string | null } | { kind: 'custom' } {
  if (!styleId) return { kind: 'para', styleName: null };
  const s = ctx.styles.get(styleId);
  const name = (s?.name ?? styleId).toLowerCase().replace(/\s+/g, '');
  const h = /^heading([1-6])$/.exec(name) ?? /^heading([1-6])$/.exec(styleId.toLowerCase());
  if (h) return { kind: 'heading', level: Number(h[1]) };
  const lvl = ctx.headingStyles.get(styleId);
  if (lvl) return { kind: 'heading', level: lvl };
  if (name === 'title') return { kind: 'para', styleName: 'title' };
  if (name === 'subtitle') return { kind: 'para', styleName: 'subtitle' };
  if (name === 'quote' || name === 'intensequote') return { kind: 'para', styleName: 'quote' };
  if (name === 'nospacing') return { kind: 'para', styleName: 'nospacing' };
  if (name === 'normal' || name === 'listparagraph' || name === 'normal(web)' || name === 'bodytext' || name === 'default') return { kind: 'para', styleName: null };
  return { kind: 'custom' };
}

function toStyleDef(run: RunProps, para: ParaProps): StyleDef {
  const d: StyleDef = {};
  if (run.font) d.font = run.font;
  if (run.size) d.size = run.size;
  if (run.color) d.color = run.color;
  if (run.bold !== undefined) d.bold = run.bold;
  if (run.italic !== undefined) d.italic = run.italic;
  if (run.caps) d.caps = true;
  if (para.before !== undefined) d.before = para.before;
  if (para.after !== undefined) d.after = para.after;
  if (para.lineHeight !== undefined) d.lineHeight = para.lineHeight;
  if (para.align && para.align !== 'left') d.align = para.align as StyleDef['align'];
  return d;
}

/* -------------------------------------------------------------- content */

type Marks = NonNullable<JSONContent['marks']>;

function runMarks(p: RunProps, base: RunProps, link: string | null, commentIds: string[]): Marks {
  const marks: Marks = [];
  if (p.bold && !base.bold) marks.push({ type: 'bold' });
  if (p.italic && !base.italic) marks.push({ type: 'italic' });
  if (p.underline && !base.underline) marks.push({ type: 'underline' });
  if (p.strike) marks.push({ type: 'strike' });
  if (p.vert === 'superscript') marks.push({ type: 'superscript' });
  if (p.vert === 'subscript') marks.push({ type: 'subscript' });
  const ts: Record<string, string> = {};
  if (p.font && p.font !== base.font) ts.fontFamily = p.font;
  if (p.size && p.size !== base.size) ts.fontSize = `${p.size}pt`;
  if (p.color && p.color !== base.color && !(link && p.color === '#0563c1')) ts.color = p.color;
  if (Object.keys(ts).length) marks.push({ type: 'textStyle', attrs: ts });
  if (p.highlight) marks.push({ type: 'highlight', attrs: { color: p.highlight } });
  if (link) marks.push({ type: 'link', attrs: { href: link } });
  if (commentIds.length) marks.push({ type: 'comment', attrs: { commentId: commentIds[commentIds.length - 1] } });
  return marks;
}

interface FieldState {
  depth: number;
  instr: string;
  phase: 'instr' | 'result';
  kind: 'hyperlink' | 'toc' | 'page' | 'numpages' | 'other';
  href?: string;
}

interface Walker {
  ctx: Ctx;
  fields: FieldState[];
  comments: string[];
  inToc: boolean;
  tocEmitted: boolean;
}

function fieldKind(instr: string): FieldState['kind'] {
  const t = instr.trim().toUpperCase();
  if (t.startsWith('HYPERLINK')) return 'hyperlink';
  if (t.startsWith('TOC')) return 'toc';
  if (t.startsWith('PAGE') && !t.startsWith('PAGEREF')) return 'page';
  if (t.startsWith('NUMPAGES') || t.startsWith('SECTIONPAGES')) return 'numpages';
  return 'other';
}

function currentLink(w: Walker): string | null {
  for (let i = w.fields.length - 1; i >= 0; i--) if (w.fields[i].kind === 'hyperlink' && w.fields[i].phase === 'result') return w.fields[i].href ?? null;
  return null;
}

/** Converts a paragraph's inline content (runs, links, images, fields). Returns inline nodes; page breaks split the paragraph. */
function inlines(pEl: Element, w: Walker, base: RunProps, pStyleRun: RunProps, out: JSONContent[][], link: string | null = null) {
  const cur = () => out[out.length - 1];
  const push = (n: JSONContent) => cur().push(n);

  const handleRun = (r: Element, linkHref: string | null) => {
    const rPr = kid(r, 'rPr');
    const rStyle = attr(kid(rPr, 'rStyle'), 'val');
    const props: RunProps = { ...pStyleRun, ...styleRun(w.ctx, rStyle ?? undefined), ...readRunProps(rPr, w.ctx) };
    if (props.hidden) return;
    const effectiveLink = linkHref ?? currentLink(w);
    const marks = runMarks(props, base, effectiveLink, w.comments);
    const text = (t: string) => {
      if (!t) return;
      const txt = props.caps ? t.toUpperCase() : t;
      push(marks.length ? { type: 'text', text: txt, marks } : { type: 'text', text: txt });
    };
    for (const c of Array.from(r.children)) {
      const inResult = !w.fields.length || w.fields[w.fields.length - 1].phase === 'result';
      switch (c.localName) {
        case 't':
          if (w.inToc) break;
          if (!inResult) break;
          text(c.textContent ?? '');
          break;
        case 'tab':
          if (inResult && !w.inToc) text('\t');
          break;
        case 'br': {
          if (w.inToc) break;
          const type = attr(c, 'type');
          if (type === 'page') out.push([{ type: '__pagebreak__' } as JSONContent], []);
          else push({ type: 'hardBreak' });
          break;
        }
        case 'cr':
          push({ type: 'hardBreak' });
          break;
        case 'noBreakHyphen':
          text('‑');
          break;
        case 'softHyphen':
          text('­');
          break;
        case 'sym': {
          const ch = attr(c, 'char');
          if (ch) {
            const code = parseInt(ch, 16);
            const mapped: Record<number, string> = { 0xf0b7: '•', 0xf0a7: '▪', 0xf0fc: '✓', 0xf0d8: '➢', 0xf0e0: '→', 0xf06c: '●', 0xf0a8: '◆' };
            text(mapped[code] ?? (code >= 0xf020 && code <= 0xf0ff ? String.fromCharCode(code - 0xf000) : String.fromCharCode(code)));
          }
          break;
        }
        case 'fldChar': {
          const t = attr(c, 'fldCharType');
          if (t === 'begin') w.fields.push({ depth: w.fields.length, instr: '', phase: 'instr', kind: 'other' });
          else if (t === 'separate') {
            const f = w.fields[w.fields.length - 1];
            if (f) {
              f.phase = 'result';
              f.kind = fieldKind(f.instr);
              if (f.kind === 'hyperlink') {
                const urlM = /HYPERLINK\s+"([^"]+)"/i.exec(f.instr);
                const anchorM = /\\l\s+"([^"]+)"/i.exec(f.instr);
                f.href = urlM ? urlM[1] + (anchorM ? `#${anchorM[1]}` : '') : anchorM ? `#${anchorM[1]}` : undefined;
              }
              if (f.kind === 'toc' && !w.inToc) {
                w.inToc = true;
                if (!w.tocEmitted) {
                  w.tocEmitted = true;
                  out.push([{ type: '__toc__' } as JSONContent], []);
                }
              }
            }
          } else if (t === 'end') {
            const f = w.fields.pop();
            if (f?.kind === 'toc') w.inToc = false;
            if (f && f.phase === 'instr') {
              // field without a result: render page fields as nothing
            }
          }
          break;
        }
        case 'instrText': {
          const f = w.fields[w.fields.length - 1];
          if (f) f.instr += c.textContent ?? '';
          break;
        }
        case 'drawing':
        case 'pict':
        case 'object': {
          if (w.inToc) break;
          const img = readImage(c, w.ctx);
          if (img) push(img);
          break;
        }
        case 'footnoteReference':
        case 'endnoteReference': {
          const id = attr(c, 'id');
          if (id && w.ctx.footnotes.has(`${c.localName === 'endnoteReference' ? 'e' : 'f'}${id}`)) {
            const key = `${c.localName === 'endnoteReference' ? 'e' : 'f'}${id}`;
            if (!w.ctx.footnoteOrder.includes(key)) w.ctx.footnoteOrder.push(key);
            const n = w.ctx.footnoteOrder.indexOf(key) + 1;
            push({ type: 'text', text: String(n), marks: [{ type: 'superscript' }] });
          }
          break;
        }
        default:
          break;
      }
    }
  };

  for (const c of Array.from(pEl.children)) {
    switch (c.localName) {
      case 'r':
        handleRun(c, link);
        break;
      case 'hyperlink': {
        const id = rAttr(c, 'id');
        const anchor = attr(c, 'anchor');
        const rel = id ? w.ctx.rels.get(id) : undefined;
        const href = rel?.target ?? (anchor ? `#${anchor}` : null);
        inlines(c, w, base, pStyleRun, out, href);
        break;
      }
      case 'ins':
      case 'smartTag':
      case 'customXml':
      case 'fldSimple':
      case 'sdtContent':
      case 'dir':
      case 'bdo': {
        if (c.localName === 'fldSimple') {
          const instr = attr(c, 'instr') ?? '';
          const kind = fieldKind(instr);
          if (kind === 'hyperlink') {
            const m = /HYPERLINK\s+"([^"]+)"/i.exec(instr);
            inlines(c, w, base, pStyleRun, out, m?.[1] ?? link);
            break;
          }
        }
        inlines(c, w, base, pStyleRun, out, link);
        break;
      }
      case 'sdt':
        inlines(kid(c, 'sdtContent') ?? c, w, base, pStyleRun, out, link);
        break;
      case 'commentRangeStart': {
        const id = attr(c, 'id');
        if (id && w.ctx.comments.has(id)) w.comments.push(id);
        break;
      }
      case 'commentRangeEnd': {
        const id = attr(c, 'id');
        w.comments = w.comments.filter((x) => x !== id);
        break;
      }
      case 'oMath':
      case 'oMathPara': {
        const t = c.textContent ?? '';
        if (t) cur().push({ type: 'text', text: t, marks: [{ type: 'italic' }] });
        break;
      }
      default:
        break;
    }
  }
}

function readImage(el: Element, ctx: Ctx): JSONContent | null {
  const blip = desc(el, 'blip')[0];
  let rid = blip ? rAttr(blip, 'embed') ?? rAttr(blip, 'link') : null;
  if (!rid) {
    const vImg = desc(el, 'imagedata')[0];
    rid = vImg ? rAttr(vImg, 'id') : null;
  }
  if (!rid) return null;
  const rel = ctx.rels.get(rid);
  if (!rel || rel.external) return null;
  const path = resolvePath('word/', rel.target);
  const data = ctx.files[path];
  if (!data) return null;
  const ext = path.split('.').pop()!.toLowerCase();
  if (ext === 'emf' || ext === 'wmf' || ext === 'tif' || ext === 'tiff') {
    ctx.warnings.add('Some pictures use old Windows formats (EMF/WMF/TIFF) that can’t be displayed and were skipped.');
    return null;
  }
  const src = bytesToDataUrl(data, mimeFromExt(ext));
  const extent = desc(el, 'extent')[0];
  let width: number | null = null;
  let height: number | null = null;
  if (extent) {
    width = Math.round(emuToPx(Number(attr(extent, 'cx') ?? 0)));
    height = Math.round(emuToPx(Number(attr(extent, 'cy') ?? 0)));
  } else {
    const shape = desc(el, 'shape')[0];
    const style = shape?.getAttribute('style') ?? '';
    const w = /width:\s*([\d.]+)(pt|px|in)/.exec(style);
    const h = /height:\s*([\d.]+)(pt|px|in)/.exec(style);
    const conv = (v: RegExpExecArray | null) => (v ? (v[2] === 'pt' ? (Number(v[1]) * 96) / 72 : v[2] === 'in' ? Number(v[1]) * 96 : Number(v[1])) : null);
    width = conv(w);
    height = conv(h);
  }
  let align = 'inline';
  const anchor = desc(el, 'anchor')[0];
  if (anchor) {
    const wrapTB = desc(anchor, 'wrapTopAndBottom').length > 0;
    const hAlign = desc(kid(anchor, 'positionH'), 'align')[0]?.textContent;
    const posOff = Number(desc(kid(anchor, 'positionH'), 'posOffset')[0]?.textContent ?? 0);
    if (wrapTB || attr(anchor, 'behindDoc') === '1') align = 'center';
    else if (hAlign === 'right' || hAlign === 'outside') align = 'right';
    else if (hAlign === 'center') align = 'center';
    else if (!hAlign && emuToPx(posOff) > 300) align = 'right';
    else align = 'left';
  }
  const docPr = desc(el, 'docPr')[0];
  ctx.imageCount++;
  return { type: 'image', attrs: { src, width, height, align, alt: docPr?.getAttribute('descr') || docPr?.getAttribute('title') || null } };
}

function resolvePath(base: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = (base + target).split('/');
  const out: string[] = [];
  for (const p of parts) {
    if (p === '..') out.pop();
    else if (p && p !== '.') out.push(p);
  }
  return out.join('/');
}

/* --------------------------------------------------------------- blocks */

interface ListFrame {
  type: 'bulletList' | 'orderedList';
  level: number;
  numId: string;
  node: JSONContent;
}

function listInfo(ctx: Ctx, numId: string, ilvl: number) {
  const lvl = ctx.numbering.get(numId)?.get(ilvl);
  const fmt = lvl?.fmt ?? 'bullet';
  const type: ListFrame['type'] = fmt === 'bullet' || fmt === 'none' ? 'bulletList' : 'orderedList';
  let listStyle: string | null = null;
  if (type === 'bulletList') {
    const t = lvl?.text ?? '';
    const code = t.charCodeAt(0);
    if (t === 'o' || t === '○' || t === '◦') listStyle = 'circle';
    else if (code === 0xf0a7 || t === '▪' || t === '■' || code === 0xf06e) listStyle = 'square';
    else if (t === '–' || t === '-' || t === '—') listStyle = 'dash';
    else if (code === 0xf0fc || t === '✓' || t === '✔') listStyle = 'check';
    else if (code === 0xf0d8 || t === '➢' || t === '➤' || code === 0xf0e8) listStyle = 'arrow';
  } else {
    listStyle = { lowerLetter: 'lower-alpha', upperLetter: 'upper-alpha', lowerRoman: 'lower-roman', upperRoman: 'upper-roman' }[fmt] ?? null;
  }
  return { type, listStyle, start: lvl?.start ?? 1 };
}

function paraAttrs(pp: ParaProps, styleDefault: ParaProps): Record<string, unknown> {
  const a: Record<string, unknown> = {};
  if (pp.align && pp.align !== (styleDefault.align ?? 'left')) a.textAlign = pp.align;
  else if (pp.align && pp.align !== 'left') a.textAlign = pp.align;
  if (pp.indent) a.indent = pp.indent;
  if (pp.indentRight) a.indentRight = pp.indentRight;
  if (pp.firstLine) a.firstLine = pp.firstLine;
  if (pp.before !== undefined && pp.before !== styleDefault.before) a.spaceBefore = pp.before;
  if (pp.after !== undefined && pp.after !== styleDefault.after) a.spaceAfter = pp.after;
  if (pp.lineHeight !== undefined && pp.lineHeight !== styleDefault.lineHeight) a.lineHeight = pp.lineHeight;
  if (pp.dir) a.dir = pp.dir;
  if (pp.shading) a.shading = pp.shading;
  return a;
}

/** Creates a TOC node, absorbing a preceding "Contents" title paragraph. */
function tocNode(blocks: JSONContent[]): JSONContent {
  const prev = blocks[blocks.length - 1];
  const prevText = prev?.content?.map((c) => c.text ?? '').join('').trim() ?? '';
  if (prev && (prev.type === 'paragraph' || prev.type === 'heading') && /^(table of )?contents?$|^inhalt|^sommaire|^contenido|^índice|^indice/i.test(prevText)) {
    blocks.pop();
    return { type: 'tableOfContents', attrs: { title: prevText } };
  }
  return { type: 'tableOfContents' };
}

function convertBody(container: Element, w: Walker, mappedDefaults: Map<string, ParaProps>): JSONContent[] {
  const blocks: JSONContent[] = [];
  const lists: ListFrame[] = [];
  const closeLists = () => {
    lists.length = 0;
  };

  const addParagraph = (p: Element) => {
    const pPr = kid(p, 'pPr');
    const own = readParaProps(pPr);
    const styleId = own.styleId;
    const mapped = mapStyle(w.ctx, styleId);
    const stylePP = stylePara(w.ctx, styleId);
    const pp: ParaProps = { ...stylePP, ...own };
    // runs are compared against the paragraph style's run formatting so style-level fonts don't become marks
    const sRun = styleRun(w.ctx, styleId);
    const baseRun: RunProps = mapped.kind === 'custom' ? { ...w.ctx.defaultRun } : { ...w.ctx.defaultRun, ...sRun };
    const pStyleRun: RunProps = mapped.kind === 'custom' ? { ...sRun } : { ...sRun };
    const segments: JSONContent[][] = [[]];
    inlines(p, w, baseRun, pStyleRun, segments);

    const mappedDefault = mapped.kind === 'heading' ? mappedDefaults.get(`h${mapped.level}`) ?? {} : mapped.kind === 'para' ? mappedDefaults.get(mapped.styleName ?? 'normal') ?? {} : w.ctx.defaultPara;
    const attrs = paraAttrs(mapped.kind === 'custom' ? pp : own, mappedDefault);

    // a paragraph that contains only the TOC field result is dropped (the TOC node replaces it)
    // Headings often carry outline numbering (LibreOffice, numbered headings) — keep them as headings.
    const lvlFmt = pp.numId ? w.ctx.numbering.get(pp.numId)?.get(pp.ilvl ?? 0)?.fmt : undefined;
    const numbered = Boolean(pp.numId && pp.numId !== '0' && mapped.kind !== 'heading' && !(mapped.kind === 'para' && mapped.styleName === 'title') && lvlFmt !== 'none');
    if (pp.pageBreakBefore) blocks.push({ type: 'pageBreak' });

    segments.forEach((seg, idx) => {
      if (seg.length === 1 && (seg[0].type as string) === '__pagebreak__') {
        closeLists();
        blocks.push({ type: 'pageBreak' });
        return;
      }
      if (seg.length === 1 && (seg[0].type as string) === '__toc__') {
        closeLists();
        blocks.push(tocNode(blocks));
        return;
      }
      if (w.inToc && !seg.length) return;
      // skip the empty tail created by a trailing page break
      if (idx > 0 && !seg.length && idx === segments.length - 1) return;
      const content = seg.length ? seg : undefined;
      let node: JSONContent;
      if (mapped.kind === 'heading') node = { type: 'heading', attrs: { level: mapped.level, ...attrs }, content };
      else node = { type: 'paragraph', attrs: { ...(mapped.kind === 'para' && mapped.styleName ? { styleName: mapped.styleName } : {}), ...attrs }, content };

      if (numbered && idx === 0) {
        const info = listInfo(w.ctx, pp.numId!, pp.ilvl ?? 0);
        const level = pp.ilvl ?? 0;
        // remove list indentation (lists indent themselves)
        if (node.attrs) {
          delete node.attrs.indent;
          delete node.attrs.firstLine;
        }
        while (lists.length && lists[lists.length - 1].level > level) lists.pop();
        let top = lists[lists.length - 1];
        if (top && top.level === level && (top.type !== info.type || (level === 0 && top.numId !== pp.numId))) {
          lists.pop();
          top = lists[lists.length - 1];
        }
        if (!top || top.level < level) {
          const listNode: JSONContent = {
            type: info.type,
            attrs: { ...(info.listStyle ? { listStyle: info.listStyle } : {}), ...(info.type === 'orderedList' && info.start !== 1 ? { start: info.start } : {}) },
            content: [],
          };
          if (top) {
            const lastItem = top.node.content![top.node.content!.length - 1];
            if (lastItem) lastItem.content!.push(listNode);
            else top.node.content!.push({ type: 'listItem', content: [{ type: 'paragraph' }, listNode] });
          } else blocks.push(listNode);
          lists.push({ type: info.type, level, numId: pp.numId!, node: listNode });
          top = lists[lists.length - 1];
        }
        if (node.type === 'heading') node = { type: 'paragraph', attrs: { ...node.attrs, level: undefined }, content: node.content };
        top.node.content!.push({ type: 'listItem', content: [node] });
        return;
      }
      closeLists();
      blocks.push(node);
    });
    // section break inside a paragraph → page break
    const sect = kid(pPr, 'sectPr');
    if (sect) {
      const t = attr(kid(sect, 'type'), 'val');
      if (t !== 'continuous') {
        closeLists();
        blocks.push({ type: 'pageBreak' });
      }
    }
  };

  const walk = (parent: Element) => {
    for (const el of Array.from(parent.children)) {
      switch (el.localName) {
        case 'p':
          addParagraph(el);
          break;
        case 'tbl':
          closeLists();
          if (!w.inToc) blocks.push(convertTable(el, w, mappedDefaults));
          break;
        case 'sdt': {
          const gallery = desc(kid(el, 'sdtPr'), 'docPartGallery')[0];
          if (gallery && /table of contents/i.test(attr(gallery, 'val') ?? '')) {
            closeLists();
            const firstP = desc(kid(el, 'sdtContent'), 'p')[0];
            const brk = firstP && (onOff(kid(kid(firstP, 'pPr'), 'pageBreakBefore')) || desc(firstP, 'br').some((b) => attr(b, 'type') === 'page'));
            if (brk && blocks.length && blocks[blocks.length - 1].type !== 'pageBreak') blocks.push({ type: 'pageBreak' });
            if (!w.tocEmitted) {
              w.tocEmitted = true;
              blocks.push(tocNode(blocks));
            }
            break;
          }
          const content = kid(el, 'sdtContent');
          if (content) walk(content);
          break;
        }
        case 'customXml':
        case 'ins':
          walk(el);
          break;
        default:
          break;
      }
    }
  };
  walk(container);
  return blocks;
}

function convertTable(tbl: Element, w: Walker, mappedDefaults: Map<string, ParaProps>): JSONContent {
  const grid = kids(kid(tbl, 'tblGrid'), 'gridCol').map((g) => Math.round(twipToPx(Number(attr(g, 'w') ?? 0))));
  const tblPr = kid(tbl, 'tblPr');
  const borders = kid(tblPr, 'tblBorders');
  const tblStyle = attr(kid(tblPr, 'tblStyle'), 'val') ?? '';
  let styleName: string | null = null;
  if (borders) {
    const vals = kids(borders).map((b) => attr(b, 'val'));
    if (vals.length && vals.every((v) => v === 'nil' || v === 'none')) styleName = 'plain';
  } else if (!/grid/i.test(tblStyle) && tblStyle) {
    // many built-in table styles have minimal borders
    if (/plain|light/i.test(tblStyle)) styleName = 'minimal';
  }
  const rows = kids(tbl, 'tr');
  // vertical merge bookkeeping: grid column → the cell node that started the merge
  const mergeOrigin = new Map<number, JSONContent>();
  const outRows: JSONContent[] = [];
  rows.forEach((tr, ri) => {
    const isHeader = Boolean(kid(kid(tr, 'trPr'), 'tblHeader'));
    let col = 0;
    const gridBefore = Number(attr(kid(kid(tr, 'trPr'), 'gridBefore'), 'val') ?? 0);
    col += gridBefore;
    const cells: JSONContent[] = [];
    for (const tc of kids(tr, 'tc')) {
      const tcPr = kid(tc, 'tcPr');
      const span = Number(attr(kid(tcPr, 'gridSpan'), 'val') ?? 1) || 1;
      const vMerge = kid(tcPr, 'vMerge');
      const vm = vMerge ? attr(vMerge, 'val') ?? 'continue' : null;
      if (vm === 'continue') {
        const origin = mergeOrigin.get(col);
        if (origin) origin.attrs!.rowspan = (origin.attrs!.rowspan as number) + 1;
        col += span;
        continue;
      }
      const fill = attr(kid(tcPr, 'shd'), 'fill');
      const valign = attr(kid(tcPr, 'vAlign'), 'val');
      const content = convertBody(tc, w, mappedDefaults);
      const cell: JSONContent = {
        type: isHeader && ri === 0 ? 'tableHeader' : 'tableCell',
        attrs: {
          colspan: span,
          rowspan: 1,
          colwidth: grid.length ? grid.slice(col, col + span).map((x) => x || 100) : null,
          backgroundColor: fill && fill !== 'auto' && /^[0-9a-f]{6}$/i.test(fill) ? `#${fill.toLowerCase()}` : null,
          valign: valign === 'center' ? 'middle' : valign === 'bottom' ? 'bottom' : null,
        },
        content: content.length ? content : [{ type: 'paragraph' }],
      };
      if (vm === 'restart' || vm === '') for (let k = 0; k < span; k++) mergeOrigin.set(col + k, cell);
      else for (let k = 0; k < span; k++) mergeOrigin.delete(col + k);
      cells.push(cell);
      col += span;
    }
    if (cells.length) outRows.push({ type: 'tableRow', content: cells });
  });
  return { type: 'table', attrs: styleName ? { styleName } : {}, content: outRows.length ? outRows : [{ type: 'tableRow', content: [{ type: 'tableCell', content: [{ type: 'paragraph' }] }] }] };
}

/* -------------------------------------------------------- header/footer */

function readHeaderFooter(xml: Document | null, ctx: Ctx): HeaderFooter {
  const out: HeaderFooter = { left: '', center: '', right: '' };
  if (!xml) return out;
  const paras = desc(xml, 'p');
  for (const p of paras) {
    let text = '';
    let field: string | null = null;
    let inInstr = false;
    const pieces: string[] = [];
    for (const r of desc(p, 'r')) {
      for (const c of Array.from(r.children)) {
        if (c.localName === 'fldChar') {
          const t = attr(c, 'fldCharType');
          if (t === 'begin') {
            field = '';
            inInstr = true;
          } else if (t === 'separate') {
            inInstr = false;
            const k = fieldKind(field ?? '');
            text += k === 'page' ? '{PAGE}' : k === 'numpages' ? '{PAGES}' : '';
            field = k === 'page' || k === 'numpages' ? 'skip' : null;
          } else if (t === 'end') {
            field = null;
            inInstr = false;
          }
        } else if (c.localName === 'instrText' && inInstr) field = (field ?? '') + (c.textContent ?? '');
        else if (c.localName === 't' && field !== 'skip') text += c.textContent ?? '';
        else if (c.localName === 'tab') {
          pieces.push(text);
          text = '';
        }
      }
    }
    for (const fs of desc(p, 'fldSimple')) {
      const k = fieldKind(attr(fs, 'instr') ?? '');
      if (k === 'page') text += '{PAGE}';
      if (k === 'numpages') text += '{PAGES}';
    }
    pieces.push(text);
    const nonEmpty = pieces.filter((s) => s.trim());
    if (!nonEmpty.length) continue;
    if (pieces.length >= 3) {
      out.left ||= pieces[0].trim();
      out.center ||= pieces[1].trim();
      out.right ||= pieces.slice(2).join(' ').trim();
    } else if (pieces.length === 2) {
      out.left ||= pieces[0].trim();
      out.right ||= pieces[1].trim();
    } else {
      const jc = attr(kid(kid(p, 'pPr'), 'jc'), 'val');
      const pStyle = attr(kid(kid(p, 'pPr'), 'pStyle'), 'val');
      const styleJc = stylePara(ctx, pStyle ?? undefined).align;
      const al = jc ? (jc === 'center' ? 'center' : jc === 'right' || jc === 'end' ? 'right' : 'left') : styleJc ?? 'left';
      const key = al === 'center' ? 'center' : al === 'right' ? 'right' : 'left';
      out[key] = out[key] ? `${out[key]} ${pieces[0].trim()}` : pieces[0].trim();
    }
  }
  return out;
}

/* ---------------------------------------------------------------- entry */

export function importDocx(bytes: Uint8Array): LoadedDoc {
  let files: Record<string, Uint8Array>;
  try {
    files = unzipSync(bytes);
  } catch {
    throw new Error('This file is not a valid .docx document (it may be damaged or password-protected).');
  }
  const text = (p: string) => (files[p] ? strFromU8(files[p]) : undefined);
  const mainPath = Object.keys(files).find((k) => /^word\/document\d*\.xml$/.test(k)) ?? 'word/document.xml';
  const doc = parseXml(text(mainPath));
  if (!doc) throw new Error('The document body could not be read.');

  const rels = new Map<string, { target: string; external: boolean; type: string }>();
  const relDoc = parseXml(text(mainPath.replace('word/', 'word/_rels/') + '.rels'));
  for (const r of desc(relDoc, 'Relationship')) {
    rels.set(r.getAttribute('Id') ?? '', { target: r.getAttribute('Target') ?? '', external: r.getAttribute('TargetMode') === 'External', type: r.getAttribute('Type') ?? '' });
  }

  const theme = parseXml(text('word/theme/theme1.xml'));
  const themeFonts = {
    major: desc(desc(theme, 'majorFont')[0], 'latin')[0]?.getAttribute('typeface') || 'Calibri',
    minor: desc(desc(theme, 'minorFont')[0], 'latin')[0]?.getAttribute('typeface') || 'Calibri',
  };

  const ctx: Ctx = {
    files,
    rels,
    styles: new Map(),
    defaultRun: {},
    defaultPara: {},
    themeFonts,
    numbering: new Map(),
    comments: new Map(),
    footnotes: new Map(),
    footnoteOrder: [],
    warnings: new Set(),
    headingStyles: new Map(),
    imageCount: 0,
  };

  // styles
  const stylesDoc = parseXml(text('word/styles.xml'));
  const docDefaults = desc(stylesDoc, 'docDefaults')[0];
  ctx.defaultRun = readRunProps(desc(desc(docDefaults, 'rPrDefault')[0], 'rPr')[0], ctx);
  ctx.defaultPara = readParaProps(desc(desc(docDefaults, 'pPrDefault')[0], 'pPr')[0]);
  for (const s of desc(stylesDoc, 'style')) {
    const id = attr(s, 'styleId');
    if (!id) continue;
    ctx.styles.set(id, {
      type: attr(s, 'type') ?? 'paragraph',
      name: attr(kid(s, 'name'), 'val') ?? id,
      basedOn: attr(kid(s, 'basedOn'), 'val') ?? undefined,
      rPr: kid(s, 'rPr'),
      pPr: kid(s, 'pPr'),
    });
    const outline = attr(kid(kid(s, 'pPr'), 'outlineLvl'), 'val');
    if (outline !== null && Number(outline) < 6 && attr(s, 'type') === 'paragraph') ctx.headingStyles.set(id, Number(outline) + 1);
  }
  const defaultStyleId = desc(stylesDoc, 'style').find((s) => attr(s, 'type') === 'paragraph' && attr(s, 'default') === '1')?.getAttribute('w:styleId') ?? 'Normal';

  // numbering
  const numDoc = parseXml(text('word/numbering.xml'));
  const abstract = new Map<string, Map<number, { fmt: string; text: string; start: number }>>();
  for (const a of desc(numDoc, 'abstractNum')) {
    const levels = new Map<number, { fmt: string; text: string; start: number }>();
    for (const l of kids(a, 'lvl')) {
      levels.set(Number(attr(l, 'ilvl') ?? 0), {
        fmt: attr(kid(l, 'numFmt'), 'val') ?? 'bullet',
        text: attr(kid(l, 'lvlText'), 'val') ?? '',
        start: Number(attr(kid(l, 'start'), 'val') ?? 1),
      });
    }
    abstract.set(attr(a, 'abstractNumId') ?? '', levels);
  }
  for (const n of desc(numDoc, 'num')) {
    const aid = attr(kid(n, 'abstractNumId'), 'val') ?? '';
    const levels = new Map(abstract.get(aid) ?? []);
    for (const ov of kids(n, 'lvlOverride')) {
      const il = Number(attr(ov, 'ilvl') ?? 0);
      const startOv = attr(kid(ov, 'startOverride'), 'val');
      const base = levels.get(il);
      if (base && startOv) levels.set(il, { ...base, start: Number(startOv) });
    }
    ctx.numbering.set(attr(n, 'numId') ?? '', levels);
  }

  // comments
  const commentsDoc = parseXml(text('word/comments.xml'));
  for (const c of desc(commentsDoc, 'comment')) {
    const id = attr(c, 'id') ?? '';
    ctx.comments.set(id, {
      id: `docx-${id}`,
      author: attr(c, 'author') ?? '',
      date: attr(c, 'date') ?? new Date().toISOString(),
      text: desc(c, 'p')
        .map((p) => desc(p, 't').map((t) => t.textContent).join(''))
        .join('\n')
        .trim(),
    });
  }

  // footnotes / endnotes
  for (const [file, prefix, tag] of [
    ['word/footnotes.xml', 'f', 'footnote'],
    ['word/endnotes.xml', 'e', 'endnote'],
  ] as const) {
    const fd = parseXml(text(file));
    for (const fn of desc(fd, tag)) {
      const type = attr(fn, 'type');
      if (type === 'separator' || type === 'continuationSeparator' || type === 'continuationNotice') continue;
      const t = desc(fn, 'p')
        .map((p) => desc(p, 't').map((x) => x.textContent).join(''))
        .join(' ')
        .trim();
      ctx.footnotes.set(`${prefix}${attr(fn, 'id')}`, t);
    }
  }

  // settings from built-in styles (so imported documents keep their look)
  const settings: DocSettings = defaultSettings();
  const styleOverrides: Partial<Record<StyleId, StyleDef>> = {};
  const mappedDefaults = new Map<string, ParaProps>();
  const baseRun = ctx.defaultRun;
  const findStyle = (pred: (name: string, id: string) => boolean) => {
    for (const [id, s] of ctx.styles) if (s.type === 'paragraph' && pred(s.name.toLowerCase().replace(/\s+/g, ''), id.toLowerCase())) return id;
    return undefined;
  };
  const normalId = ctx.styles.has(defaultStyleId) ? defaultStyleId : findStyle((n) => n === 'normal');
  const normalRun = { ...baseRun, ...styleRun(ctx, normalId) };
  const normalPara = { ...ctx.defaultPara, ...stylePara(ctx, normalId) };
  styleOverrides.normal = {
    font: normalRun.font ?? themeFonts.minor,
    size: normalRun.size ?? 11,
    color: normalRun.color ?? '#000000',
    after: normalPara.after ?? 0,
    before: normalPara.before ?? 0,
    lineHeight: normalPara.lineHeight ?? 1,
  };
  mappedDefaults.set('normal', normalPara);
  const mapSpecial: Array<[StyleId, (n: string, id: string) => boolean]> = [
    ['title', (n) => n === 'title'],
    ['subtitle', (n) => n === 'subtitle'],
    ['quote', (n) => n === 'quote'],
    ['h1', (n, id) => n === 'heading1' || id === 'heading1'],
    ['h2', (n, id) => n === 'heading2' || id === 'heading2'],
    ['h3', (n, id) => n === 'heading3' || id === 'heading3'],
    ['h4', (n, id) => n === 'heading4' || id === 'heading4'],
    ['h5', (n, id) => n === 'heading5' || id === 'heading5'],
    ['h6', (n, id) => n === 'heading6' || id === 'heading6'],
  ];
  for (const [key, pred] of mapSpecial) {
    const id = findStyle(pred);
    if (!id) continue;
    const run = { ...normalRun, ...styleRun(ctx, id) };
    const para = { ...normalPara, ...stylePara(ctx, id) };
    const def = toStyleDef(run, para);
    if (key.startsWith('h') && def.bold === undefined) def.bold = false;
    styleOverrides[key] = def;
    mappedDefaults.set(key, para);
  }
  settings.styles = styleOverrides;

  // page setup from the last section
  const body = desc(doc, 'body')[0];
  const sectPrs = kids(body, 'sectPr');
  const sect = sectPrs[sectPrs.length - 1] ?? desc(body, 'sectPr').pop();
  if (sect) {
    const pgSz = kid(sect, 'pgSz');
    const wIn = twipToIn(Number(attr(pgSz, 'w') ?? 12240));
    const hIn = twipToIn(Number(attr(pgSz, 'h') ?? 15840));
    const landscape = attr(pgSz, 'orient') === 'landscape' || wIn > hIn;
    const pw = Math.min(wIn, hIn);
    const ph = Math.max(wIn, hIn);
    const known = PAPER_SIZES.find((p) => Math.abs(p.width - pw) < 0.08 && Math.abs(p.height - ph) < 0.08);
    const mar = kid(sect, 'pgMar');
    const m = (k: string, d: number) => {
      const v = attr(mar, k);
      return v === null ? d : Math.max(0, twipToIn(Math.abs(Number(v))));
    };
    settings.page = {
      size: known?.id ?? 'custom',
      width: known?.width ?? pw,
      height: known?.height ?? ph,
      orientation: landscape ? 'landscape' : 'portrait',
      margins: { top: m('top', 1), right: m('right', 1), bottom: m('bottom', 1), left: m('left', 1) },
    };
    settings.differentFirstPage = Boolean(kid(sect, 'titlePg')) && onOff(kid(sect, 'titlePg')) !== false;
    const hf = (tag: 'headerReference' | 'footerReference') => {
      const refs = kids(sect, tag);
      const ref = refs.find((r) => attr(r, 'type') === 'default') ?? refs[0];
      const rel = ref ? rels.get(rAttr(ref, 'id') ?? '') : undefined;
      return rel ? readHeaderFooter(parseXml(text(resolvePath('word/', rel.target))), ctx) : { left: '', center: '', right: '' };
    };
    settings.header = hf('headerReference');
    settings.footer = hf('footerReference');
  }

  // core properties
  const core = parseXml(text('docProps/core.xml'));
  const coreVal = (n: string) => desc(core, n)[0]?.textContent?.trim() ?? '';
  settings.title = coreVal('title');
  settings.author = coreVal('creator');
  settings.subject = coreVal('subject');
  settings.keywords = coreVal('keywords');
  settings.created = coreVal('created') || settings.created;

  const walker: Walker = { ctx, fields: [], comments: [], inToc: false, tocEmitted: false };
  const blocks = body ? convertBody(body, walker, mappedDefaults) : [];

  if (ctx.footnoteOrder.length) {
    blocks.push({ type: 'horizontalRule' });
    blocks.push({ type: 'paragraph', content: [{ type: 'text', text: 'Notes', marks: [{ type: 'bold' }] }] });
    ctx.footnoteOrder.forEach((k, i) => {
      blocks.push({ type: 'paragraph', attrs: { styleName: 'nospacing' }, content: [{ type: 'text', text: `${i + 1}. ${ctx.footnotes.get(k) ?? ''}`, marks: [{ type: 'textStyle', attrs: { fontSize: '9pt' } }] }] });
    });
  }

  // TipTap requires non-empty text nodes and at least one block
  const clean = (nodes: JSONContent[]): JSONContent[] =>
    nodes
      .map((n) => (n.content ? { ...n, content: clean(n.content) } : n))
      .filter((n) => !(n.type === 'text' && !n.text));
  const content = clean(blocks);

  const comments = Array.from(ctx.comments.values()).map((c) => c);
  // comment marks use "docx-<id>" ids
  const fixComments = (nodes: JSONContent[]) => {
    for (const n of nodes) {
      n.marks?.forEach((m) => {
        if (m.type === 'comment' && m.attrs && !String(m.attrs.commentId).startsWith('docx-')) m.attrs.commentId = `docx-${m.attrs.commentId}`;
      });
      if (n.content) fixComments(n.content);
    }
  };
  fixComments(content);

  return {
    settings,
    content: { type: 'doc', content: content.length ? content : [{ type: 'paragraph' }] },
    comments,
    warning: ctx.warnings.size ? Array.from(ctx.warnings).join(' ') : undefined,
  };
}
