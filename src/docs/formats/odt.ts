import type { JSONContent } from '@tiptap/core';
import { strFromU8, strToU8, unzipSync, zipSync, type Zippable } from 'fflate';
import { bytesToDataUrl, dataUrlToBytes, escapeXml, extFromMime, imageSize, mimeFromExt, rasterizeToPng } from '@/lib/utils';
import { defaultSettings, PAPER_SIZES, resolveStyles, type DocModel, type DocSettings, type HeaderFooter, type StyleDef } from '../model';
import { cssToPx } from '../editor/extensions';
import { collectHeadingsFromJson, renderMathPng, textOf } from './shared';
import type { LoadedDoc } from './index';

const NS = {
  office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
  style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
  text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
  table: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
  draw: 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
  fo: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
  xlink: 'http://www.w3.org/1999/xlink',
  svg: 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
  meta: 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
  dc: 'http://purl.org/dc/elements/1.1/',
};

/* =================================================================== IMPORT */

interface OdStyle {
  family: string;
  parent?: string;
  props: Record<string, string>;
  listStyle?: string;
}

const kids = (el: Element | null | undefined, local?: string) => (el ? Array.from(el.children).filter((c) => !local || c.localName === local) : []);
const at = (el: Element | null | undefined, ns: keyof typeof NS, name: string) => (el ? el.getAttributeNS(NS[ns], name) : null);

function lengthToPx(v: string | null | undefined): number | null {
  if (!v) return null;
  return cssToPx(v);
}

function collectStyles(root: Document | null, map: Map<string, OdStyle>, fontFaces: Map<string, string>) {
  if (!root) return;
  for (const ff of Array.from(root.getElementsByTagNameNS(NS.style, 'font-face'))) {
    const name = at(ff, 'style', 'name');
    const fam = at(ff, 'svg', 'font-family')?.replace(/^['"]|['"]$/g, '');
    if (name) fontFaces.set(name, fam || name);
  }
  for (const s of Array.from(root.getElementsByTagNameNS(NS.style, 'style'))) {
    const name = at(s, 'style', 'name');
    if (!name) continue;
    const props: Record<string, string> = {};
    for (const p of kids(s)) {
      if (!p.localName.endsWith('-properties')) continue;
      for (const a of Array.from(p.attributes)) props[a.localName] = a.value;
      if (p.localName === 'paragraph-properties' && p.getAttributeNS(NS.fo, 'break-before')) props['break-before'] = p.getAttributeNS(NS.fo, 'break-before')!;
    }
    map.set(name, { family: at(s, 'style', 'family') ?? 'paragraph', parent: at(s, 'style', 'parent-style-name') ?? undefined, props, listStyle: at(s, 'style', 'list-style-name') ?? undefined });
  }
}

function resolved(map: Map<string, OdStyle>, name: string | null | undefined): Record<string, string> {
  const chain: OdStyle[] = [];
  const seen = new Set<string>();
  let cur = name ?? undefined;
  while (cur && !seen.has(cur)) {
    seen.add(cur);
    const s = map.get(cur);
    if (!s) break;
    chain.unshift(s);
    cur = s.parent;
  }
  return chain.reduce<Record<string, string>>((acc, s) => ({ ...acc, ...s.props }), {});
}

function displayName(map: Map<string, OdStyle>, name: string | null | undefined): string {
  // follow automatic styles to their named parent
  let cur = name ?? '';
  const seen = new Set<string>();
  while (/^P\d+$|^T\d+$/.test(cur) && map.get(cur)?.parent && !seen.has(cur)) {
    seen.add(cur);
    cur = map.get(cur)!.parent!;
  }
  return cur.replace(/_20_/g, ' ').toLowerCase();
}

function marksFrom(p: Record<string, string>, fonts: Map<string, string>, base: Record<string, string>): JSONContent['marks'] {
  const marks: NonNullable<JSONContent['marks']> = [];
  if (p['font-weight'] && /bold|[6-9]00/.test(p['font-weight']) && !(base['font-weight'] && /bold|[6-9]00/.test(base['font-weight']))) marks.push({ type: 'bold' });
  if (p['font-style'] === 'italic' && base['font-style'] !== 'italic') marks.push({ type: 'italic' });
  if (p['text-underline-style'] && p['text-underline-style'] !== 'none') marks.push({ type: 'underline' });
  if (p['text-line-through-style'] && p['text-line-through-style'] !== 'none') marks.push({ type: 'strike' });
  const pos = p['text-position'];
  if (pos && /^super|^\d/.test(pos) && !pos.startsWith('0')) marks.push({ type: pos.startsWith('super') || parseFloat(pos) > 0 ? 'superscript' : 'subscript' });
  if (pos && /^sub|^-/.test(pos)) marks.push({ type: 'subscript' });
  const ts: Record<string, string> = {};
  const font = p['font-name'] ? fonts.get(p['font-name']) ?? p['font-name'] : p['font-family']?.replace(/['"]/g, '');
  const baseFont = base['font-name'] ? fonts.get(base['font-name']) ?? base['font-name'] : undefined;
  if (font && font !== baseFont) ts.fontFamily = font;
  if (p['font-size'] && p['font-size'] !== base['font-size'] && !p['font-size'].endsWith('%')) ts.fontSize = p['font-size'];
  if (p.color && p.color !== base.color && p.color !== '#000000') ts.color = p.color;
  if (Object.keys(ts).length) marks.push({ type: 'textStyle', attrs: ts });
  const bg = p['background-color'];
  if (bg && bg !== 'transparent' && bg !== base['background-color']) marks.push({ type: 'highlight', attrs: { color: bg } });
  return marks;
}

function paraAttrsFrom(p: Record<string, string>): Record<string, unknown> {
  const a: Record<string, unknown> = {};
  const align = p['text-align'];
  if (align) a.textAlign = align === 'center' ? 'center' : align === 'end' || align === 'right' ? 'right' : align === 'justify' ? 'justify' : undefined;
  const ml = lengthToPx(p['margin-left']);
  if (ml && ml > 0) a.indent = Math.round(ml);
  const mr = lengthToPx(p['margin-right']);
  if (mr && mr > 0) a.indentRight = Math.round(mr);
  const ti = lengthToPx(p['text-indent']);
  if (ti) a.firstLine = Math.round(ti);
  const mt = lengthToPx(p['margin-top']);
  if (mt !== null) a.spaceBefore = Math.round(((mt * 72) / 96) * 10) / 10;
  const mb = lengthToPx(p['margin-bottom']);
  if (mb !== null) a.spaceAfter = Math.round(((mb * 72) / 96) * 10) / 10;
  const lh = p['line-height'];
  if (lh && lh.endsWith('%')) a.lineHeight = parseFloat(lh) / 100;
  if (p['writing-mode'] === 'rl-tb') a.dir = 'rtl';
  const bg = p['background-color'];
  if (bg && bg !== 'transparent') a.shading = bg;
  if (!a.textAlign) delete a.textAlign;
  return a;
}

export function importOdt(bytes: Uint8Array): LoadedDoc {
  const files = unzipSync(bytes);
  const read = (p: string) => (files[p] ? new DOMParser().parseFromString(strFromU8(files[p]), 'application/xml') : null);
  const content = read('content.xml');
  if (!content) throw new Error('This file is not a valid OpenDocument text file.');
  const stylesDoc = read('styles.xml');
  const styleMap = new Map<string, OdStyle>();
  const fonts = new Map<string, string>();
  collectStyles(stylesDoc, styleMap, fonts);
  collectStyles(content, styleMap, fonts);

  // list styles: name → level → bullet/number
  const listStyles = new Map<string, Map<number, { ordered: boolean; format: string; bullet: string }>>();
  for (const root of [stylesDoc, content]) {
    if (!root) continue;
    for (const ls of Array.from(root.getElementsByTagNameNS(NS.text, 'list-style'))) {
      const levels = new Map<number, { ordered: boolean; format: string; bullet: string }>();
      for (const lvl of kids(ls)) {
        const level = Number(at(lvl, 'text', 'level') ?? 1);
        levels.set(level, { ordered: lvl.localName === 'list-level-style-number', format: at(lvl, 'style', 'num-format') ?? '1', bullet: at(lvl, 'text', 'bullet-char') ?? '•' });
      }
      const name = at(ls, 'style', 'name');
      if (name) listStyles.set(name, levels);
    }
  }

  const defaultPara = (() => {
    const d = stylesDoc?.getElementsByTagNameNS(NS.style, 'default-style');
    const out: Record<string, string> = {};
    for (const el of Array.from(d ?? [])) {
      if (at(el, 'style', 'family') !== 'paragraph') continue;
      for (const p of kids(el)) for (const a of Array.from(p.attributes)) out[a.localName] = a.value;
    }
    return out;
  })();
  const standard = { ...defaultPara, ...resolved(styleMap, 'Standard') };
  const settings: DocSettings = defaultSettings();
  const normal: StyleDef = {
    font: standard['font-name'] ? fonts.get(standard['font-name']) ?? standard['font-name'] : 'Liberation Serif',
    size: standard['font-size'] ? parseFloat(standard['font-size']) : 12,
    color: standard.color ?? '#000000',
    after: standard['margin-bottom'] ? Math.round((((lengthToPx(standard['margin-bottom']) ?? 0) * 72) / 96) * 10) / 10 : 0,
    before: 0,
    lineHeight: standard['line-height']?.endsWith('%') ? parseFloat(standard['line-height']) / 100 : 1.15,
  };
  settings.styles = { normal };
  for (let lvl = 1; lvl <= 3; lvl++) {
    const h = resolved(styleMap, `Heading_20_${lvl}`);
    if (!Object.keys(h).length) continue;
    const key = `h${lvl}` as 'h1';
    settings.styles[key] = {
      font: h['font-name'] ? fonts.get(h['font-name']) ?? h['font-name'] : undefined,
      size: h['font-size'] && !h['font-size'].endsWith('%') ? parseFloat(h['font-size']) : h['font-size']?.endsWith('%') ? ((normal.size ?? 12) * parseFloat(h['font-size'])) / 100 : undefined,
      color: h.color,
      bold: /bold|[6-9]00/.test(h['font-weight'] ?? ''),
      italic: h['font-style'] === 'italic',
    };
  }

  const baseRun = standard;
  const images: Record<string, string> = {};
  const imageFor = (href: string) => {
    if (images[href]) return images[href];
    const data = files[href];
    if (!data) return null;
    const ext = href.split('.').pop()!.toLowerCase();
    if (['svm', 'wmf', 'emf'].includes(ext)) return null;
    images[href] = bytesToDataUrl(data, mimeFromExt(ext));
    return images[href];
  };
  const notes: string[] = [];

  const inline = (el: Element, inherited: Record<string, string>, link: string | null, out: JSONContent[]) => {
    for (const n of Array.from(el.childNodes)) {
      if (n.nodeType === 3) {
        const text = (n.textContent ?? '').replace(/[\n\r]/g, '');
        if (!text) continue;
        const marks = marksFrom(inherited, fonts, baseRun) ?? [];
        if (link) marks.push({ type: 'link', attrs: { href: link } });
        out.push({ type: 'text', text, ...(marks.length ? { marks } : {}) });
        continue;
      }
      if (n.nodeType !== 1) continue;
      const c = n as Element;
      switch (c.localName) {
        case 'span': {
          const props = { ...inherited, ...resolved(styleMap, at(c, 'text', 'style-name')) };
          inline(c, props, link, out);
          break;
        }
        case 'a':
          inline(c, inherited, at(c, 'xlink', 'href'), out);
          break;
        case 's': {
          const count = Number(at(c, 'text', 'c') ?? 1);
          out.push({ type: 'text', text: ' '.repeat(count) });
          break;
        }
        case 'tab':
          out.push({ type: 'text', text: '\t' });
          break;
        case 'line-break':
          out.push({ type: 'hardBreak' });
          break;
        case 'frame': {
          const img = c.getElementsByTagNameNS(NS.draw, 'image')[0];
          const href = at(img, 'xlink', 'href');
          const src = href ? imageFor(href) : null;
          if (src) {
            const w = lengthToPx(at(c, 'svg', 'width'));
            const h = lengthToPx(at(c, 'svg', 'height'));
            const anchor = at(c, 'text', 'anchor-type');
            const wrap = resolved(styleMap, at(c, 'draw', 'style-name'));
            const align = anchor === 'as-char' ? 'inline' : wrap['horizontal-pos'] === 'right' ? 'right' : wrap['horizontal-pos'] === 'center' ? 'center' : wrap.wrap === 'none' ? 'center' : 'left';
            out.push({ type: 'image', attrs: { src, width: w ? Math.round(w) : null, height: h ? Math.round(h) : null, align, alt: c.getAttributeNS(NS.draw, 'name') } });
          }
          break;
        }
        case 'note': {
          const body = c.getElementsByTagNameNS(NS.text, 'note-body')[0];
          notes.push((body?.textContent ?? '').trim());
          out.push({ type: 'text', text: String(notes.length), marks: [{ type: 'superscript' }] });
          break;
        }
        case 'bookmark':
        case 'bookmark-start':
        case 'bookmark-end':
        case 'soft-page-break':
          break;
        default:
          inline(c, inherited, link, out);
      }
    }
  };

  const blocks = (container: Element): JSONContent[] => blocksList(kids(container));

  const blocksList = (els: Element[]): JSONContent[] => {
    const out: JSONContent[] = [];
    for (const el of els) {
      switch (el.localName) {
        case 'p':
        case 'h': {
          const sName = at(el, 'text', 'style-name');
          const props = resolved(styleMap, sName);
          if (props['break-before'] === 'page' && out.length) out.push({ type: 'pageBreak' });
          const disp = displayName(styleMap, sName);
          const inl: JSONContent[] = [];
          const isHeading = el.localName === 'h';
          const headingBase = isHeading ? resolved(styleMap, `Heading_20_${at(el, 'text', 'outline-level') ?? 1}`) : {};
          inline(el, { ...props }, null, inl);
          // don't turn heading/paragraph style formatting into marks
          const cleaned = inl.map((n) => (n.marks ? { ...n, marks: n.marks.filter((m) => !(isHeading && (m.type === 'bold' || (m.type === 'textStyle' && headingBase['font-size'])))) } : n));
          const attrs = paraAttrsFrom(props);
          if (isHeading) {
            out.push({ type: 'heading', attrs: { level: Math.min(6, Number(at(el, 'text', 'outline-level') ?? 1)), ...attrs }, content: cleaned.length ? cleaned : undefined });
          } else {
            const styleName = disp === 'title' ? 'title' : disp === 'subtitle' ? 'subtitle' : disp === 'quotations' ? 'quote' : null;
            out.push({ type: 'paragraph', attrs: { ...(styleName ? { styleName } : {}), ...attrs }, content: cleaned.length ? cleaned : undefined });
          }
          break;
        }
        case 'list': {
          out.push(list(el, at(el, 'text', 'style-name'), 1));
          break;
        }
        case 'table':
          out.push(table(el));
          break;
        case 'section':
        case 'index-body':
          out.push(...blocks(el));
          break;
        case 'table-of-content':
          out.push({ type: 'tableOfContents' });
          break;
        default:
          break;
      }
    }
    return out;
  };

  const list = (el: Element, styleName: string | null, level: number): JSONContent => {
    const ls = styleName ? listStyles.get(styleName)?.get(level) : undefined;
    const ordered = ls?.ordered ?? false;
    const fmt = ls?.format;
    const listStyle = ordered
      ? fmt === 'a'
        ? 'lower-alpha'
        : fmt === 'A'
          ? 'upper-alpha'
          : fmt === 'i'
            ? 'lower-roman'
            : fmt === 'I'
              ? 'upper-roman'
              : null
      : ls?.bullet === '–' || ls?.bullet === '-'
        ? 'dash'
        : ls?.bullet === '▪' || ls?.bullet === '■'
          ? 'square'
          : ls?.bullet === '◦' || ls?.bullet === '○'
            ? 'circle'
            : ls?.bullet === '✔' || ls?.bullet === '✓'
              ? 'check'
              : ls?.bullet === '➢' || ls?.bullet === '➔' || ls?.bullet === '➤'
                ? 'arrow'
                : null;
    const items: JSONContent[] = [];
    for (const item of kids(el)) {
      if (item.localName !== 'list-item' && item.localName !== 'list-header') continue;
      const inner: JSONContent[] = [];
      for (const c of kids(item)) {
        if (c.localName === 'list') inner.push(list(c, at(c, 'text', 'style-name') ?? styleName, level + 1));
        else inner.push(...blocksList([c]));
      }
      if (!inner.length || inner[0].type !== 'paragraph') inner.unshift({ type: 'paragraph' });
      items.push({ type: 'listItem', content: inner.map((n) => (n.type === 'heading' ? { ...n, type: 'paragraph' } : n)) });
    }
    return { type: ordered ? 'orderedList' : 'bulletList', attrs: listStyle ? { listStyle } : {}, content: items.length ? items : [{ type: 'listItem', content: [{ type: 'paragraph' }] }] };
  };

  const table = (el: Element): JSONContent => {
    const cols: number[] = [];
    for (const c of kids(el, 'table-column')) {
      const rep = Number(at(c, 'table', 'number-columns-repeated') ?? 1);
      const w = lengthToPx(resolved(styleMap, at(c, 'table', 'style-name'))['column-width']);
      for (let i = 0; i < Math.min(rep, 64); i++) cols.push(w ? Math.round(w) : 0);
    }
    const rows: JSONContent[] = [];
    const rowEls = [...kids(el, 'table-row'), ...kids(el, 'table-header-rows').flatMap((h) => kids(h, 'table-row'))];
    const headerRows = new Set(kids(el, 'table-header-rows').flatMap((h) => kids(h, 'table-row')));
    const ordered = [...rowEls].sort((a, b) => (headerRows.has(a) === headerRows.has(b) ? 0 : headerRows.has(a) ? -1 : 1));
    for (const tr of ordered) {
      const cells: JSONContent[] = [];
      let col = 0;
      for (const tc of kids(tr)) {
        if (tc.localName === 'covered-table-cell') {
          col++;
          continue;
        }
        if (tc.localName !== 'table-cell') continue;
        const span = Number(at(tc, 'table', 'number-columns-spanned') ?? 1);
        const rspan = Number(at(tc, 'table', 'number-rows-spanned') ?? 1);
        const cp = resolved(styleMap, at(tc, 'table', 'style-name'));
        const content = blocks(tc);
        cells.push({
          type: headerRows.has(tr) ? 'tableHeader' : 'tableCell',
          attrs: {
            colspan: span,
            rowspan: rspan,
            colwidth: cols.length ? cols.slice(col, col + span).map((x) => x || 100) : null,
            backgroundColor: cp['background-color'] && cp['background-color'] !== 'transparent' ? cp['background-color'] : null,
          },
          content: content.length ? content : [{ type: 'paragraph' }],
        });
        col += span;
      }
      if (cells.length) rows.push({ type: 'tableRow', content: cells });
    }
    return { type: 'table', content: rows.length ? rows : [{ type: 'tableRow', content: [{ type: 'tableCell', content: [{ type: 'paragraph' }] }] }] };
  };

  const text = content.getElementsByTagNameNS(NS.office, 'text')[0];
  const body = text ? blocks(text) : [];
  if (notes.length) {
    body.push({ type: 'horizontalRule' });
    notes.forEach((t, i) => body.push({ type: 'paragraph', attrs: { styleName: 'nospacing' }, content: t ? [{ type: 'text', text: `${i + 1}. ${t}`, marks: [{ type: 'textStyle', attrs: { fontSize: '9pt' } }] }] : undefined }));
  }

  // page layout + header/footer
  const master = stylesDoc?.getElementsByTagNameNS(NS.style, 'master-page')[0];
  const layoutName = at(master, 'style', 'page-layout-name');
  const layout = Array.from(stylesDoc?.getElementsByTagNameNS(NS.style, 'page-layout') ?? []).find((l) => at(l, 'style', 'name') === layoutName);
  const lp = layout ? kids(layout, 'page-layout-properties')[0] : undefined;
  if (lp) {
    const w = (lengthToPx(at(lp, 'fo', 'page-width')) ?? 816) / 96;
    const h = (lengthToPx(at(lp, 'fo', 'page-height')) ?? 1056) / 96;
    const land = at(lp, 'style', 'print-orientation') === 'landscape' || w > h;
    const pw = Math.min(w, h);
    const ph = Math.max(w, h);
    const known = PAPER_SIZES.find((p) => Math.abs(p.width - pw) < 0.08 && Math.abs(p.height - ph) < 0.08);
    const m = (k: string) => (lengthToPx(at(lp, 'fo', k)) ?? 96) / 96;
    settings.page = { size: known?.id ?? 'custom', width: known?.width ?? pw, height: known?.height ?? ph, orientation: land ? 'landscape' : 'portrait', margins: { top: m('margin-top'), right: m('margin-right'), bottom: m('margin-bottom'), left: m('margin-left') } };
  }
  const hfText = (el: Element | undefined): HeaderFooter => {
    const out: HeaderFooter = { left: '', center: '', right: '' };
    if (!el) return out;
    for (const p of Array.from(el.getElementsByTagNameNS(NS.text, 'p'))) {
      let s = '';
      const walk = (n: Node) => {
        if (n.nodeType === 3) s += n.textContent;
        else if (n.nodeType === 1) {
          const e = n as Element;
          if (e.localName === 'page-number') s += '{PAGE}';
          else if (e.localName === 'page-count') s += '{PAGES}';
          else if (e.localName === 'tab') s += '\t';
          else e.childNodes.forEach(walk);
        }
      };
      p.childNodes.forEach(walk);
      const parts = s.split('\t');
      if (parts.length >= 3) Object.assign(out, { left: parts[0].trim(), center: parts[1].trim(), right: parts.slice(2).join(' ').trim() });
      else if (s.trim()) {
        const al = resolved(styleMap, at(p, 'text', 'style-name'))['text-align'];
        out[al === 'center' ? 'center' : al === 'end' || al === 'right' ? 'right' : 'left'] = s.trim();
      }
    }
    return out;
  };
  settings.header = hfText(master ? kids(master, 'header')[0] : undefined);
  settings.footer = hfText(master ? kids(master, 'footer')[0] : undefined);

  const meta = read('meta.xml');
  if (meta) {
    settings.title = meta.getElementsByTagNameNS(NS.dc, 'title')[0]?.textContent ?? '';
    settings.author = meta.getElementsByTagNameNS(NS.meta, 'initial-creator')[0]?.textContent ?? meta.getElementsByTagNameNS(NS.dc, 'creator')[0]?.textContent ?? '';
    settings.subject = meta.getElementsByTagNameNS(NS.dc, 'subject')[0]?.textContent ?? '';
  }

  return { settings, content: { type: 'doc', content: body.length ? body : [{ type: 'paragraph' }] }, comments: [] };
}

/* =================================================================== EXPORT */

interface ExportCtx {
  paraStyles: Map<string, string>;
  textStyles: Map<string, string>;
  autoStyles: string[];
  fonts: Set<string>;
  images: Map<string, { path: string; w: number; h: number }>;
  pictures: Record<string, Uint8Array>;
  listCount: number;
  listStyles: Map<string, string>;
  tableCount: number;
  math: Map<string, { path: string; w: number; h: number }>;
  headingPages: number[];
  headings: Array<{ level: number; text: string }>;
  firstPageMaster: boolean;
}

const pt = (n: number) => `${Math.round(n * 100) / 100}pt`;
const inch = (n: number) => `${Math.round(n * 10000) / 10000}in`;
const px2in = (px: number) => inch(px / 96);

function textStyleFor(marks: JSONContent['marks'], ctx: ExportCtx): string | null {
  const props: string[] = [];
  for (const m of marks ?? []) {
    if (m.type === 'bold') props.push('fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"');
    if (m.type === 'italic') props.push('fo:font-style="italic" style:font-style-asian="italic" style:font-style-complex="italic"');
    if (m.type === 'underline') props.push('style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"');
    if (m.type === 'strike') props.push('style:text-line-through-style="solid"');
    if (m.type === 'superscript') props.push('style:text-position="super 58%"');
    if (m.type === 'subscript') props.push('style:text-position="sub 58%"');
    if (m.type === 'code') {
      ctx.fonts.add('Courier New');
      props.push('style:font-name="Courier New"');
    }
    if (m.type === 'highlight') props.push(`fo:background-color="${escapeXml(String(m.attrs?.color ?? '#ffff00'))}"`);
    if (m.type === 'textStyle') {
      const a = m.attrs ?? {};
      if (a.fontFamily) {
        const f = String(a.fontFamily).split(',')[0].replace(/["']/g, '').trim();
        ctx.fonts.add(f);
        props.push(`style:font-name="${escapeXml(f)}"`);
      }
      if (a.fontSize) {
        const v = String(a.fontSize);
        const n = parseFloat(v);
        const ptv = /px$/.test(v) ? (n * 72) / 96 : n;
        props.push(`fo:font-size="${pt(ptv)}" style:font-size-asian="${pt(ptv)}" style:font-size-complex="${pt(ptv)}"`);
      }
      if (a.color) props.push(`fo:color="${escapeXml(String(a.color))}"`);
      if (a.backgroundColor) props.push(`fo:background-color="${escapeXml(String(a.backgroundColor))}"`);
    }
  }
  if (!props.length) return null;
  const key = props.join(' ');
  let name = ctx.textStyles.get(key);
  if (!name) {
    name = `T${ctx.textStyles.size + 1}`;
    ctx.textStyles.set(key, name);
    ctx.autoStyles.push(`<style:style style:name="${name}" style:family="text"><style:text-properties ${key}/></style:style>`);
  }
  return name;
}

function paraStyleFor(n: JSONContent, parent: string, ctx: ExportCtx, extra: { breakBefore?: boolean; masterPage?: string } = {}): string {
  const a = n.attrs ?? {};
  const props: string[] = [];
  const al = a.textAlign as string | undefined;
  if (al) props.push(`fo:text-align="${al === 'center' ? 'center' : al === 'right' ? 'end' : al === 'justify' ? 'justify' : 'start'}"`);
  if (a.indent) props.push(`fo:margin-left="${px2in(Number(a.indent))}"`);
  if (a.indentRight) props.push(`fo:margin-right="${px2in(Number(a.indentRight))}"`);
  if (a.firstLine) props.push(`fo:text-indent="${px2in(Number(a.firstLine))}"`);
  if (a.spaceBefore !== null && a.spaceBefore !== undefined) props.push(`fo:margin-top="${pt(Number(a.spaceBefore))}"`);
  if (a.spaceAfter !== null && a.spaceAfter !== undefined) props.push(`fo:margin-bottom="${pt(Number(a.spaceAfter))}"`);
  if (a.lineHeight) props.push(`fo:line-height="${Math.round(Number(a.lineHeight) * 100)}%"`);
  if (a.dir === 'rtl') props.push('style:writing-mode="rl-tb"');
  if (a.shading) props.push(`fo:background-color="${escapeXml(String(a.shading))}"`);
  if (extra.breakBefore) props.push('fo:break-before="page"');
  if (!props.length && !extra.masterPage) return parent;
  const key = `${parent}|${props.join(' ')}|${extra.masterPage ?? ''}`;
  let name = ctx.paraStyles.get(key);
  if (!name) {
    name = `P${ctx.paraStyles.size + 1}`;
    ctx.paraStyles.set(key, name);
    ctx.autoStyles.push(
      `<style:style style:name="${name}" style:family="paragraph" style:parent-style-name="${parent}"${extra.masterPage ? ` style:master-page-name="${extra.masterPage}"` : ''}><style:paragraph-properties ${props.join(' ')}/></style:style>`,
    );
  }
  return name;
}

function textXml(t: string): string {
  // tabs, runs of spaces and newlines need special elements
  let out = '';
  const parts = t.split(/(\t|\n| {2,})/);
  for (const p of parts) {
    if (!p) continue;
    if (p === '\t') out += '<text:tab/>';
    else if (p === '\n') out += '<text:line-break/>';
    else if (/^ {2,}$/.test(p)) out += ` <text:s text:c="${p.length - 1}"/>`;
    else out += escapeXml(p);
  }
  return out;
}

function inlineXml(nodes: JSONContent[] | undefined, ctx: ExportCtx): string {
  let out = '';
  for (const n of nodes ?? []) {
    if (n.type === 'text') {
      const link = n.marks?.find((m) => m.type === 'link');
      const style = textStyleFor(n.marks?.filter((m) => m.type !== 'link' && m.type !== 'comment'), ctx);
      let x = textXml(n.text ?? '');
      if (style) x = `<text:span text:style-name="${style}">${x}</text:span>`;
      if (link) x = `<text:a xlink:type="simple" xlink:href="${escapeXml(String(link.attrs?.href ?? ''))}">${x}</text:a>`;
      out += x;
    } else if (n.type === 'hardBreak') out += '<text:line-break/>';
    else if (n.type === 'image') {
      const img = ctx.images.get(String(n.attrs?.src));
      if (!img) continue;
      const w = Number(n.attrs?.width) || img.w;
      const h = Number(n.attrs?.height) || (img.h * w) / (img.w || 1);
      const align = n.attrs?.align;
      const anchor = align === 'left' || align === 'right' ? 'paragraph' : 'as-char';
      const frameStyle = align === 'left' ? 'frLeft' : align === 'right' ? 'frRight' : 'frInline';
      out += `<draw:frame draw:style-name="${frameStyle}" draw:name="Image${ctx.images.size}${Math.random().toString(36).slice(2, 6)}" text:anchor-type="${anchor}" svg:width="${px2in(w)}" svg:height="${px2in(h)}"><draw:image xlink:href="${img.path}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>${n.attrs?.alt ? `<svg:desc>${escapeXml(String(n.attrs.alt))}</svg:desc>` : ''}</draw:frame>`;
    } else if (n.type === 'mathInline') {
      const m = ctx.math.get(String(n.attrs?.latex));
      if (m) out += `<draw:frame draw:style-name="frInline" text:anchor-type="as-char" svg:width="${px2in(m.w)}" svg:height="${px2in(m.h)}"><draw:image xlink:href="${m.path}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>`;
      else out += escapeXml(String(n.attrs?.latex ?? ''));
    }
  }
  return out;
}

const LIST_BULLETS: Record<string, string> = { disc: '•', circle: '◦', square: '▪', dash: '–', check: '✔', arrow: '➢' };

function listStyleFor(ordered: boolean, style: string | null, ctx: ExportCtx): string {
  const key = `${ordered ? 'o' : 'b'}-${style ?? ''}`;
  let name = ctx.listStyles.get(key);
  if (name) return name;
  name = `L${ctx.listStyles.size + 1}`;
  ctx.listStyles.set(key, name);
  const levels = Array.from({ length: 10 }, (_, i) => {
    const lvl = i + 1;
    const pos = `<style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="${inch(0.25 + i * 0.25 + 0.25)}" fo:text-indent="-0.25in" fo:margin-left="${inch(0.5 + i * 0.25)}"/></style:list-level-properties>`;
    if (ordered) {
      const fmt = i === 0 ? { 'lower-alpha': 'a', 'upper-alpha': 'A', 'lower-roman': 'i', 'upper-roman': 'I' }[style ?? ''] ?? '1' : ['1', 'a', 'i'][i % 3];
      return `<text:list-level-style-number text:level="${lvl}" style:num-suffix="." style:num-format="${fmt}">${pos}</text:list-level-style-number>`;
    }
    const ch = i === 0 ? LIST_BULLETS[style ?? 'disc'] ?? '•' : ['•', '◦', '▪'][i % 3];
    return `<text:list-level-style-bullet text:level="${lvl}" text:bullet-char="${ch}">${pos}</text:list-level-style-bullet>`;
  }).join('');
  ctx.autoStyles.push(`<text:list-style style:name="${name}">${levels}</text:list-style>`);
  return name;
}

function blocksXml(nodes: JSONContent[] | undefined, ctx: ExportCtx, pageBreakNext = { v: false }): string {
  let out = '';
  for (const n of nodes ?? []) {
    const brk = pageBreakNext.v;
    switch (n.type) {
      case 'paragraph': {
        const sn = n.attrs?.styleName;
        const parent = sn === 'title' ? 'Title' : sn === 'subtitle' ? 'Subtitle' : sn === 'quote' ? 'Quotations' : 'Standard';
        const master = ctx.firstPageMaster ? 'First_20_Page' : undefined;
        ctx.firstPageMaster = false;
        out += `<text:p text:style-name="${paraStyleFor(n, parent, ctx, { breakBefore: brk, masterPage: master })}">${inlineXml(n.content, ctx)}</text:p>`;
        pageBreakNext.v = false;
        break;
      }
      case 'heading': {
        const l = Math.min(6, Number(n.attrs?.level) || 1);
        const master = ctx.firstPageMaster ? 'First_20_Page' : undefined;
        ctx.firstPageMaster = false;
        out += `<text:h text:style-name="${paraStyleFor(n, `Heading_20_${l}`, ctx, { breakBefore: brk, masterPage: master })}" text:outline-level="${l}">${inlineXml(n.content, ctx)}</text:h>`;
        pageBreakNext.v = false;
        break;
      }
      case 'bulletList':
      case 'orderedList':
      case 'taskList': {
        const ordered = n.type === 'orderedList';
        const name = listStyleFor(ordered, (n.attrs?.listStyle as string | null) ?? (n.type === 'taskList' ? 'check' : null), ctx);
        out += `<text:list text:style-name="${name}">${(n.content ?? [])
          .map((item) => {
            const inner = n.type === 'taskList' ? (item.content ?? []).map((c, i) => (i === 0 && c.type === 'paragraph' ? { ...c, content: [{ type: 'text', text: item.attrs?.checked ? '☑ ' : '☐ ' }, ...(c.content ?? [])] } : c)) : item.content;
            return `<text:list-item>${blocksXml(inner, ctx)}</text:list-item>`;
          })
          .join('')}</text:list>`;
        break;
      }
      case 'blockquote':
        out += blocksXml(
          (n.content ?? []).map((c) => ({ ...c, attrs: { ...(c.attrs ?? {}), indent: 48 } })),
          ctx,
        );
        break;
      case 'codeBlock':
        out += `<text:p text:style-name="Preformatted_20_Text">${textXml(textOf(n))}</text:p>`;
        break;
      case 'horizontalRule':
        out += '<text:p text:style-name="Horizontal_20_Line"/>';
        break;
      case 'pageBreak':
        pageBreakNext.v = true;
        break;
      case 'tableOfContents': {
        const title = escapeXml(String(n.attrs?.title ?? 'Contents'));
        const headStyle = paraStyleFor({ attrs: {} }, 'Contents_20_Heading', ctx, { breakBefore: brk });
        pageBreakNext.v = false;
        const entries = ctx.headings
          .map((h, i) => ({ ...h, page: ctx.headingPages[i] }))
          .filter((h) => h.level <= 3 && h.text)
          .map((h) => `<text:p text:style-name="Contents_20_${h.level}">${escapeXml(h.text)}<text:tab/>${h.page ?? ''}</text:p>`)
          .join('');
        const tpl = [1, 2, 3]
          .map(
            (l) =>
              `<text:table-of-content-entry-template text:outline-level="${l}" text:style-name="Contents_20_${l}"><text:index-entry-link-start/><text:index-entry-chapter/><text:index-entry-text/><text:index-entry-tab-stop style:type="right" style:leader-char="."/><text:index-entry-page-number/><text:index-entry-link-end/></text:table-of-content-entry-template>`,
          )
          .join('');
        out += `<text:table-of-content text:protected="true" text:name="Table of Contents1"><text:table-of-content-source text:outline-level="3"><text:index-title-template text:style-name="Contents_20_Heading">${title}</text:index-title-template>${tpl}</text:table-of-content-source><text:index-body><text:index-title text:name="Table of Contents1_Head"><text:p text:style-name="${headStyle}">${title}</text:p></text:index-title>${entries}</text:index-body></text:table-of-content>`;
        break;
      }
      case 'mathBlock': {
        const m = ctx.math.get(String(n.attrs?.latex));
        out += `<text:p text:style-name="${paraStyleFor({ attrs: { textAlign: 'center' } }, 'Standard', ctx)}">${
          m ? `<draw:frame draw:style-name="frInline" text:anchor-type="as-char" svg:width="${px2in(m.w)}" svg:height="${px2in(m.h)}"><draw:image xlink:href="${m.path}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>` : escapeXml(String(n.attrs?.latex ?? ''))
        }</text:p>`;
        break;
      }
      case 'table':
        out += tableXml(n, ctx);
        break;
      default:
        if (n.content) out += blocksXml(n.content, ctx, pageBreakNext);
    }
  }
  return out;
}

const ODT_TABLE_THEMES: Record<string, { head?: string; headText?: string; band?: string; border?: string; noBorder?: boolean }> = {
  accent: { head: '#2f6dff', headText: '#ffffff', border: '#c9d6f3' },
  banded: { band: '#f2f5fb', border: '#d6dbe5' },
  minimal: { border: '#d6dbe5' },
  plain: { noBorder: true },
  dark: { head: '#1f2937', headText: '#ffffff', band: '#f3f4f6', border: '#c4c9d2' },
  green: { head: '#17a35a', headText: '#ffffff', border: '#bfe6cf' },
  orange: { head: '#f26a26', headText: '#ffffff', border: '#f7cdb5' },
};

function tableXml(n: JSONContent, ctx: ExportCtx): string {
  const tname = `Table${++ctx.tableCount}`;
  const rows = n.content ?? [];
  const widths: number[] = [];
  for (const c of rows[0]?.content ?? []) {
    const span = Number(c.attrs?.colspan) || 1;
    const cw = (c.attrs?.colwidth as number[] | null) ?? [];
    for (let k = 0; k < span; k++) widths.push(cw[k] || 120);
  }
  const total = widths.reduce((a, b) => a + b, 0);
  ctx.autoStyles.push(`<style:style style:name="${tname}" style:family="table"><style:table-properties style:width="${px2in(total)}" table:align="margins"/></style:style>`);
  widths.forEach((w, i) => ctx.autoStyles.push(`<style:style style:name="${tname}.C${i}" style:family="table-column"><style:table-column-properties style:column-width="${px2in(w)}"/></style:style>`));
  const theme = ODT_TABLE_THEMES[String(n.attrs?.styleName ?? '')] ?? { head: '#f1f4f9', border: '#9aa4b5' };
  const border = theme.noBorder ? 'none' : `0.5pt solid ${theme.border ?? '#9aa4b5'}`;
  let body = widths.map((_, i) => `<table:table-column table:style-name="${tname}.C${i}"/>`).join('');
  const covered = new Map<number, number>(); // col → remaining rows covered
  rows.forEach((row, ri) => {
    let xml = '';
    let col = 0;
    const isHeader = (row.content ?? []).every((c) => c.type === 'tableHeader');
    const cells = [...(row.content ?? [])];
    let ci = 0;
    while (ci < cells.length || covered.has(col)) {
      if ((covered.get(col) ?? 0) > 0) {
        xml += '<table:covered-table-cell/>';
        covered.set(col, covered.get(col)! - 1);
        if (!covered.get(col)) covered.delete(col);
        col++;
        continue;
      }
      const c = cells[ci++];
      if (!c) break;
      const span = Number(c.attrs?.colspan) || 1;
      const rspan = Number(c.attrs?.rowspan) || 1;
      const bg = c.attrs?.backgroundColor as string | null;
      const cstyle = `${tname}.R${ri}C${col}`;
      const fill = bg ?? (c.type === 'tableHeader' ? theme.head : theme.band && ri % 2 === 0 && ri > 0 ? theme.band : undefined);
      ctx.autoStyles.push(
        `<style:style style:name="${cstyle}" style:family="table-cell"><style:table-cell-properties fo:padding="0.04in" fo:border="${border}"${fill ? ` fo:background-color="${escapeXml(fill)}"` : ''}${c.attrs?.valign === 'middle' ? ' style:vertical-align="middle"' : c.attrs?.valign === 'bottom' ? ' style:vertical-align="bottom"' : ''}/></style:style>`,
      );
      const headMarks = [{ type: 'bold' }, ...(theme.headText ? [{ type: 'textStyle', attrs: { color: theme.headText } }] : [])];
      const content = c.type === 'tableHeader' ? (c.content ?? []).map((p) => ({ ...p, content: (p.content ?? []).map((t) => (t.type === 'text' ? { ...t, marks: [...(t.marks ?? []), ...headMarks] } : t)) })) : c.content;
      xml += `<table:table-cell table:style-name="${cstyle}" office:value-type="string"${span > 1 ? ` table:number-columns-spanned="${span}"` : ''}${rspan > 1 ? ` table:number-rows-spanned="${rspan}"` : ''}>${blocksXml(content, ctx) || '<text:p/>'}</table:table-cell>`;
      for (let k = 1; k < span; k++) xml += '<table:covered-table-cell/>';
      if (rspan > 1) for (let k = 0; k < span; k++) covered.set(col + k, rspan - 1);
      col += span;
    }
    const rowXml = `<table:table-row>${xml}</table:table-row>`;
    body += isHeader && ri === 0 ? `<table:table-header-rows>${rowXml}</table:table-header-rows>` : rowXml;
  });
  return `<table:table table:name="${tname}" table:style-name="${tname}">${body}</table:table>`;
}

function styleDefProps(s: StyleDef, ctx: ExportCtx): { para: string; text: string } {
  const text: string[] = [];
  const para: string[] = [];
  if (s.font) {
    ctx.fonts.add(s.font);
    text.push(`style:font-name="${escapeXml(s.font)}"`);
  }
  if (s.size) text.push(`fo:font-size="${pt(s.size)}"`);
  if (s.color) text.push(`fo:color="${s.color}"`);
  text.push(`fo:font-weight="${s.bold ? 'bold' : 'normal'}"`);
  if (s.italic) text.push('fo:font-style="italic"');
  if (s.caps) text.push('fo:text-transform="uppercase"');
  if (s.before !== undefined) para.push(`fo:margin-top="${pt(s.before)}"`);
  if (s.after !== undefined) para.push(`fo:margin-bottom="${pt(s.after)}"`);
  if (s.lineHeight) para.push(`fo:line-height="${Math.round(s.lineHeight * 100)}%"`);
  if (s.align) para.push(`fo:text-align="${s.align === 'center' ? 'center' : s.align === 'right' ? 'end' : s.align === 'justify' ? 'justify' : 'start'}"`);
  return { para: para.join(' '), text: text.join(' ') };
}

function hfXml(h: HeaderFooter, contentWidthIn: number): string {
  const f = (t: string) =>
    escapeXml(t)
      .replace(/\{PAGE\}/gi, '<text:page-number text:select-page="current">1</text:page-number>')
      .replace(/\{PAGES\}/gi, '<text:page-count>1</text:page-count>')
      .replace(/\{DATE\}/gi, `<text:date>${escapeXml(new Date().toLocaleDateString())}</text:date>`)
      .replace(/\{TITLE\}/gi, '<text:title/>');
  void contentWidthIn;
  return `<text:p text:style-name="Header">${f(h.left)}<text:tab/>${f(h.center)}<text:tab/>${f(h.right)}</text:p>`;
}

export async function exportOdt(model: DocModel, title: string, headingPages: Map<number, number> = new Map()): Promise<Uint8Array> {
  const s: DocSettings = model.settings;
  const content = model.content as JSONContent;
  const ctx: ExportCtx = {
    paraStyles: new Map(),
    textStyles: new Map(),
    autoStyles: [],
    fonts: new Set(),
    images: new Map(),
    pictures: {},
    listCount: 0,
    listStyles: new Map(),
    tableCount: 0,
    math: new Map(),
    headingPages: Array.from(headingPages.entries())
      .sort((a, b) => a[0] - b[0])
      .map((e) => e[1]),
    headings: collectHeadingsFromJson(content, 6, true),
    firstPageMaster: Boolean(model.settings.differentFirstPage),
  };

  // images & maths
  const imgs = new Set<string>();
  const maths = new Set<string>();
  const walk = (n: JSONContent) => {
    if (n.type === 'image' && n.attrs?.src) imgs.add(String(n.attrs.src));
    if ((n.type === 'mathInline' || n.type === 'mathBlock') && n.attrs?.latex) maths.add(String(n.attrs.latex));
    n.content?.forEach(walk);
  };
  walk(content);
  let i = 0;
  for (const src of imgs) {
    try {
      let parsed = dataUrlToBytes(src);
      let url = src;
      if (!parsed || !/image\/(png|jpeg|gif|svg\+xml)/.test(parsed.mime)) {
        url = (await rasterizeToPng(src)).dataUrl;
        parsed = dataUrlToBytes(url);
      }
      if (!parsed) continue;
      const size = await imageSize(url);
      const path = `Pictures/image${++i}.${extFromMime(parsed.mime)}`;
      ctx.pictures[path] = parsed.bytes;
      ctx.images.set(src, { path, w: size.width, h: size.height });
    } catch {
      /* skip */
    }
  }
  for (const latex of maths) {
    const r = await renderMathPng(latex);
    if (!r) continue;
    const path = `Pictures/math${++i}.png`;
    ctx.pictures[path] = r.data;
    ctx.math.set(latex, { path, w: r.w, h: r.h });
  }

  const bodyXml = blocksXml(content.content, ctx);
  const st = resolveStyles(s);
  const named = (name: string, display: string, def: StyleDef, extra = '', parent = 'Standard', cls = 'text') => {
    const p = styleDefProps(def, ctx);
    return `<style:style style:name="${name}" style:display-name="${display}" style:family="paragraph" style:parent-style-name="${parent}" style:class="${cls}"${extra}><style:paragraph-properties ${p.para}/><style:text-properties ${p.text}/></style:style>`;
  };
  const normal = styleDefProps(st.normal, ctx);
  const headings = [1, 2, 3, 4, 5, 6]
    .map((l) => named(`Heading_20_${l}`, `Heading ${l}`, { ...st[`h${l}` as 'h1'] }, ` style:default-outline-level="${l}" style:next-style-name="Standard"`, 'Standard', 'text').replace('<style:paragraph-properties ', '<style:paragraph-properties fo:keep-with-next="always" '))
    .join('');
  const page = s.page;
  const pw = page.orientation === 'landscape' ? page.height : page.width;
  const ph = page.orientation === 'landscape' ? page.width : page.height;
  const cw = pw - page.margins.left - page.margins.right;
  const hasH = s.header.left || s.header.center || s.header.right;
  const hasF = s.footer.left || s.footer.center || s.footer.right;
  const fontDecls = Array.from(ctx.fonts)
    .map((f) => `<style:font-face style:name="${escapeXml(f)}" svg:font-family="'${escapeXml(f)}'"/>`)
    .join('');
  const decl = `xmlns:office="${NS.office}" xmlns:style="${NS.style}" xmlns:text="${NS.text}" xmlns:table="${NS.table}" xmlns:draw="${NS.draw}" xmlns:fo="${NS.fo}" xmlns:xlink="${NS.xlink}" xmlns:dc="${NS.dc}" xmlns:meta="${NS.meta}" xmlns:svg="${NS.svg}" office:version="1.3"`;

  const stylesXml = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles ${decl}>
<office:font-face-decls>${fontDecls}</office:font-face-decls>
<office:styles>
<style:default-style style:family="paragraph"><style:paragraph-properties style:tab-stop-distance="0.5in"/><style:text-properties ${normal.text}/></style:default-style>
<style:style style:name="Standard" style:family="paragraph" style:class="text"><style:paragraph-properties ${normal.para}/><style:text-properties ${normal.text}/></style:style>
${headings}
${named('Title', 'Title', st.title, ' style:next-style-name="Subtitle"', 'Standard', 'chapter')}
${named('Subtitle', 'Subtitle', st.subtitle, '', 'Standard', 'chapter')}
${named('Quotations', 'Quotations', st.quote, '', 'Standard', 'html')}
<style:style style:name="Preformatted_20_Text" style:display-name="Preformatted Text" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:background-color="#f5f7fb" fo:padding="0.06in" fo:margin-top="0.06in" fo:margin-bottom="0.1in"/><style:text-properties style:font-name="Courier New" fo:font-size="10pt"/></style:style>
<style:style style:name="Horizontal_20_Line" style:display-name="Horizontal Line" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:border-bottom="0.75pt solid #b8c0cf" fo:margin-bottom="0.15in"/></style:style>
<style:style style:name="Contents_20_Heading" style:display-name="Contents Heading" style:family="paragraph" style:parent-style-name="Heading_20_1"/>
${[1, 2, 3]
  .map(
    (l) =>
      `<style:style style:name="Contents_20_${l}" style:display-name="Contents ${l}" style:family="paragraph" style:parent-style-name="Standard" style:class="index"><style:paragraph-properties fo:margin-left="${inch((l - 1) * 0.25)}" fo:margin-top="0in" fo:margin-bottom="0.03in"><style:tab-stops><style:tab-stop style:position="${inch(cw - (l - 1) * 0.25)}" style:type="right" style:leader-style="dotted" style:leader-text="."/></style:tab-stops></style:paragraph-properties></style:style>`,
  )
  .join('')}
<style:style style:name="Header" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties><style:tab-stops><style:tab-stop style:position="${inch(cw / 2)}" style:type="center"/><style:tab-stop style:position="${inch(cw)}" style:type="right"/></style:tab-stops></style:paragraph-properties><style:text-properties fo:font-size="9pt" fo:color="#555555"/></style:style>
<style:style style:name="frInline" style:family="graphic"><style:graphic-properties style:vertical-pos="top" style:vertical-rel="baseline"/></style:style>
<style:style style:name="frLeft" style:family="graphic"><style:graphic-properties style:wrap="parallel" style:horizontal-pos="left" style:horizontal-rel="paragraph" fo:margin-right="0.12in"/></style:style>
<style:style style:name="frRight" style:family="graphic"><style:graphic-properties style:wrap="parallel" style:horizontal-pos="right" style:horizontal-rel="paragraph" fo:margin-left="0.12in"/></style:style>
</office:styles>
<office:automatic-styles>
<style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="${inch(pw)}" fo:page-height="${inch(ph)}" style:print-orientation="${page.orientation}" fo:margin-top="${inch(hasH ? Math.max(0.2, page.margins.top - 0.4) : page.margins.top)}" fo:margin-bottom="${inch(hasF ? Math.max(0.2, page.margins.bottom - 0.4) : page.margins.bottom)}" fo:margin-left="${inch(page.margins.left)}" fo:margin-right="${inch(page.margins.right)}"${s.pageColor ? ` fo:background-color="${s.pageColor}"` : ''}/>${hasH ? '<style:header-style><style:header-footer-properties fo:min-height="0.3in" fo:margin-bottom="0.1in"/></style:header-style>' : ''}${hasF ? '<style:footer-style><style:header-footer-properties fo:min-height="0.3in" fo:margin-top="0.1in"/></style:footer-style>' : ''}</style:page-layout>
</office:automatic-styles>
<office:master-styles>${s.differentFirstPage ? '<style:master-page style:name="First_20_Page" style:display-name="First Page" style:page-layout-name="pm1" style:next-style-name="Standard"/>' : ''}<style:master-page style:name="Standard" style:page-layout-name="pm1">${hasH ? `<style:header>${hfXml(s.header, cw)}</style:header>` : ''}${hasF ? `<style:footer>${hfXml(s.footer, cw)}</style:footer>` : ''}</style:master-page></office:master-styles>
</office:document-styles>`;

  const contentXml = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content ${decl}>
<office:font-face-decls>${fontDecls}</office:font-face-decls>
<office:automatic-styles>${ctx.autoStyles.join('')}</office:automatic-styles>
<office:body><office:text>${bodyXml || '<text:p/>'}</office:text></office:body>
</office:document-content>`;

  const now = new Date().toISOString().replace(/\.\d+Z$/, '');
  const metaXml = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta ${decl}><office:meta><meta:generator>Affice</meta:generator><dc:title>${escapeXml(s.title || title.replace(/\.[^.]+$/, ''))}</dc:title>${s.author ? `<meta:initial-creator>${escapeXml(s.author)}</meta:initial-creator><dc:creator>${escapeXml(s.author)}</dc:creator>` : ''}${s.subject ? `<dc:subject>${escapeXml(s.subject)}</dc:subject>` : ''}<meta:creation-date>${escapeXml(s.created.replace(/\.\d+Z$/, ''))}</meta:creation-date><dc:date>${now}</dc:date></office:meta></office:document-meta>`;

  const manifest = `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">
<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
${Object.keys(ctx.pictures)
  .map((p) => `<manifest:file-entry manifest:full-path="${p}" manifest:media-type="${mimeFromExt(p.split('.').pop()!)}"/>`)
  .join('\n')}
</manifest:manifest>`;

  const files: Zippable = {
    mimetype: [strToU8('application/vnd.oasis.opendocument.text'), { level: 0 }],
    'content.xml': strToU8(contentXml),
    'styles.xml': strToU8(stylesXml),
    'meta.xml': strToU8(metaXml),
    'META-INF/manifest.xml': strToU8(manifest),
  };
  for (const [p, d] of Object.entries(ctx.pictures)) files[p] = [d, { level: 0 }];
  return zipSync(files, { level: 6 });
}
