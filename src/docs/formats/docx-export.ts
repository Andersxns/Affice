import type { JSONContent } from '@tiptap/core';
import {
  AlignmentType,
  Bookmark,
  BorderStyle,
  CommentRangeEnd,
  CommentRangeStart,
  CommentReference,
  Document,
  ExternalHyperlink,
  Footer,
  Header,
  HeadingLevel,
  HighlightColor,
  HorizontalPositionAlign,
  HorizontalPositionRelativeFrom,
  ImageRun,
  InternalHyperlink,
  LevelFormat,
  LineRuleType,
  Packer,
  PageBreak,
  PageNumber,
  PageOrientation,
  Paragraph,
  ShadingType,
  Tab,
  Table,
  TableCell,
  TableLayoutType,
  TableOfContents,
  TableRow,
  TabStopType,
  TextRun,
  TextWrappingSide,
  TextWrappingType,
  VerticalAlign,
  VerticalPositionRelativeFrom,
  WidthType,
  type IParagraphOptions,
  type IRunOptions,
  type ParagraphChild,
} from 'docx';
import { dataUrlToBytes, imageSize, rasterizeToPng } from '@/lib/utils';
import { collectHeadingsFromJson, slugifyHeading } from './shared';
import { resolveStyles, type DocModel, type HeaderFooter, type StyleDef } from '../model';

const PX_TO_TWIP = 15;
const hex = (c: string | null | undefined) => (c ? c.replace('#', '').toUpperCase().slice(0, 6) : undefined);

/* -------------------------------------------------------- highlight colours */

const NAMED_HL: Array<[string, string]> = [
  ['FFFF00', HighlightColor.YELLOW],
  ['00FF00', HighlightColor.GREEN],
  ['00FFFF', HighlightColor.CYAN],
  ['FF00FF', HighlightColor.MAGENTA],
  ['0000FF', HighlightColor.BLUE],
  ['FF0000', HighlightColor.RED],
  ['000080', HighlightColor.DARK_BLUE],
  ['008080', HighlightColor.DARK_CYAN],
  ['008000', HighlightColor.DARK_GREEN],
  ['800080', HighlightColor.DARK_MAGENTA],
  ['800000', HighlightColor.DARK_RED],
  ['808000', HighlightColor.DARK_YELLOW],
  ['808080', HighlightColor.DARK_GRAY],
  ['C0C0C0', HighlightColor.LIGHT_GRAY],
  ['000000', HighlightColor.BLACK],
];

/* ------------------------------------------------------------ run building */

interface Ctx {
  model: DocModel;
  commentIds: Map<string, number>;
  images: Map<string, { data: Uint8Array; type: 'png' | 'jpg' | 'gif' | 'bmp'; w: number; h: number }>;
  mathImages: Map<string, { data: Uint8Array; w: number; h: number }>;
  numberingRefs: Map<string, { format: string; start: number; bullet?: string }>;
  listInstance: number;
  headingSlugs: Map<string, number>;
  contentWidthTwip: number;
  headingPageList: number[];
}

function fontSizeHalfPts(v: string | undefined): number | undefined {
  if (!v) return undefined;
  const n = parseFloat(v);
  if (!Number.isFinite(n)) return undefined;
  const pt = /px$/.test(v) ? (n * 72) / 96 : n;
  return Math.round(pt * 2);
}

function runOptions(marks: JSONContent['marks'], extra: Partial<IRunOptions> = {}): IRunOptions {
  const o: Record<string, unknown> = { ...extra };
  for (const m of marks ?? []) {
    switch (m.type) {
      case 'bold':
        o.bold = true;
        break;
      case 'italic':
        o.italics = true;
        break;
      case 'underline':
        o.underline = { type: 'single' };
        break;
      case 'strike':
        o.strike = true;
        break;
      case 'subscript':
        o.subScript = true;
        break;
      case 'superscript':
        o.superScript = true;
        break;
      case 'code':
        o.font = 'Courier New';
        o.shading = { type: ShadingType.CLEAR, fill: 'EEF1F6', color: 'auto' };
        break;
      case 'textStyle': {
        const a = m.attrs ?? {};
        if (a.fontFamily) o.font = String(a.fontFamily).split(',')[0].replace(/["']/g, '').trim();
        const sz = fontSizeHalfPts(a.fontSize as string | undefined);
        if (sz) o.size = sz;
        if (a.color) o.color = hex(a.color as string);
        if (a.backgroundColor) o.shading = { type: ShadingType.CLEAR, fill: hex(a.backgroundColor as string), color: 'auto' };
        break;
      }
      case 'highlight': {
        const c = hex((m.attrs?.color as string) ?? '#ffff00')!;
        const named = NAMED_HL.find(([h]) => h === c);
        if (named) o.highlight = named[1];
        else o.shading = { type: ShadingType.CLEAR, fill: c, color: 'auto' };
        break;
      }
      case 'link':
        o.style = 'Hyperlink';
        break;
      default:
        break;
    }
  }
  return o as IRunOptions;
}

/** Text → TextRuns, turning tab characters into real tabs. */
function textRuns(text: string, opts: IRunOptions): TextRun[] {
  if (!text.includes('\t')) return [new TextRun({ ...opts, text })];
  const parts = text.split('\t');
  const children: Array<string | Tab> = [];
  parts.forEach((p, i) => {
    if (i > 0) children.push(new Tab());
    if (p) children.push(p);
  });
  return [new TextRun({ ...opts, children })];
}

type Inline = ParagraphChild;

function inlineChildren(nodes: JSONContent[] | undefined, ctx: Ctx, base: Partial<IRunOptions> = {}): Inline[] {
  const out: Inline[] = [];
  if (!nodes) return out;
  let i = 0;
  while (i < nodes.length) {
    const n = nodes[i];
    const link = n.marks?.find((m) => m.type === 'link');
    if (link) {
      // group consecutive nodes with the same link
      const href = String(link.attrs?.href ?? '');
      const group: JSONContent[] = [];
      while (i < nodes.length && nodes[i].marks?.some((m) => m.type === 'link' && m.attrs?.href === href)) group.push(nodes[i++]);
      const children = group.flatMap((g) => leaf(g, ctx, base)).filter((c): c is TextRun => c instanceof TextRun);
      if (href.startsWith('#')) out.push(new InternalHyperlink({ anchor: href.slice(1), children }));
      else out.push(new ExternalHyperlink({ link: href, children }));
      continue;
    }
    out.push(...leaf(n, ctx, base));
    i++;
  }
  return out;
}

function leaf(n: JSONContent, ctx: Ctx, base: Partial<IRunOptions>): Inline[] {
  switch (n.type) {
    case 'text': {
      const comment = n.marks?.find((m) => m.type === 'comment');
      const runs = textRuns(n.text ?? '', runOptions(n.marks, base));
      if (comment) {
        const id = ctx.commentIds.get(String(comment.attrs?.commentId));
        if (id !== undefined) return [new CommentRangeStart(id), ...runs, new CommentRangeEnd(id), new TextRun({ children: [new CommentReference(id)] })];
      }
      return runs;
    }
    case 'hardBreak':
      return [new TextRun({ text: '', break: 1 })];
    case 'image': {
      const img = ctx.images.get(String(n.attrs?.src));
      if (!img) return [];
      const width = Math.round(Number(n.attrs?.width) || img.w);
      const height = Math.round(Number(n.attrs?.height) || (img.h * width) / (img.w || 1));
      const align = n.attrs?.align;
      const floating =
        align === 'left' || align === 'right'
          ? {
              horizontalPosition: { relative: HorizontalPositionRelativeFrom.MARGIN, align: align === 'left' ? HorizontalPositionAlign.LEFT : HorizontalPositionAlign.RIGHT },
              verticalPosition: { relative: VerticalPositionRelativeFrom.PARAGRAPH, offset: 0 },
              wrap: { type: TextWrappingType.SQUARE, side: TextWrappingSide.BOTH_SIDES },
              margins: { left: 114300, right: 114300, top: 0, bottom: 57150 },
            }
          : undefined;
      return [
        new ImageRun({
          type: img.type,
          data: img.data,
          transformation: { width, height },
          floating,
          altText: n.attrs?.alt ? { name: String(n.attrs.alt), description: String(n.attrs.alt), title: String(n.attrs.alt) } : undefined,
        }),
      ];
    }
    case 'mathInline': {
      const m = ctx.mathImages.get(String(n.attrs?.latex));
      if (m) return [new ImageRun({ type: 'png', data: m.data, transformation: { width: m.w, height: m.h } })];
      return [new TextRun({ text: String(n.attrs?.latex ?? ''), font: 'Cambria Math', italics: true })];
    }
    default:
      return n.content ? inlineChildren(n.content, ctx, base) : [];
  }
}

/* --------------------------------------------------------- paragraphs */

function paraOptions(n: JSONContent): Partial<IParagraphOptions> {
  const a = n.attrs ?? {};
  const o: Record<string, unknown> = {};
  const align = a.textAlign as string | undefined;
  if (align) o.alignment = align === 'center' ? AlignmentType.CENTER : align === 'right' ? AlignmentType.RIGHT : align === 'justify' ? AlignmentType.JUSTIFIED : AlignmentType.LEFT;
  const indent: Record<string, number> = {};
  if (a.indent) indent.left = Math.round(Number(a.indent) * PX_TO_TWIP);
  if (a.indentRight) indent.right = Math.round(Number(a.indentRight) * PX_TO_TWIP);
  if (a.firstLine) {
    const v = Math.round(Number(a.firstLine) * PX_TO_TWIP);
    if (v > 0) indent.firstLine = v;
    else indent.hanging = -v;
  }
  if (Object.keys(indent).length) o.indent = indent;
  const spacing: Record<string, unknown> = {};
  if (a.spaceBefore !== null && a.spaceBefore !== undefined) spacing.before = Math.round(Number(a.spaceBefore) * 20);
  if (a.spaceAfter !== null && a.spaceAfter !== undefined) spacing.after = Math.round(Number(a.spaceAfter) * 20);
  if (a.lineHeight) {
    spacing.line = Math.round(Number(a.lineHeight) * 240);
    spacing.lineRule = LineRuleType.AUTO;
  }
  if (Object.keys(spacing).length) o.spacing = spacing;
  if (a.dir === 'rtl') o.bidirectional = true;
  if (a.shading) o.shading = { type: ShadingType.CLEAR, fill: hex(a.shading as string), color: 'auto' };
  const style = a.styleName as string | null;
  if (style === 'title') o.heading = HeadingLevel.TITLE;
  else if (style === 'subtitle') o.style = 'Subtitle';
  else if (style === 'quote') o.style = 'Quote';
  else if (style === 'nospacing') o.style = 'NoSpacing';
  return o as Partial<IParagraphOptions>;
}

const HEADINGS = [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3, HeadingLevel.HEADING_4, HeadingLevel.HEADING_5, HeadingLevel.HEADING_6];

type Block = Paragraph | Table | TableOfContents;

interface ListCtx {
  ref: string;
  level: number;
  instance: number;
  task?: boolean;
}

function listRef(ctx: Ctx, ordered: boolean, style: string | null, start: number): string {
  const format = ordered ? style ?? 'decimal' : 'bullet';
  const key = ordered ? `ol-${style ?? 'decimal'}-${start}` : `ul-${style ?? 'disc'}`;
  if (!ctx.numberingRefs.has(key)) ctx.numberingRefs.set(key, { format, start, bullet: ordered ? undefined : style ?? 'disc' });
  return key;
}

function blocks(nodes: JSONContent[] | undefined, ctx: Ctx, list?: ListCtx): Block[] {
  const out: Block[] = [];
  for (const n of nodes ?? []) out.push(...block(n, ctx, list));
  return out;
}

function block(n: JSONContent, ctx: Ctx, list?: ListCtx): Block[] {
  const numbering = list && !list.task ? { reference: list.ref, level: Math.min(8, list.level), instance: list.instance } : undefined;
  switch (n.type) {
    case 'paragraph': {
      const kids = inlineChildren(n.content, ctx);
      const onlyImage = n.content?.length === 1 && n.content[0].type === 'image' && n.content[0].attrs?.align === 'center';
      const opts = paraOptions(n);
      if (onlyImage) (opts as Record<string, unknown>).alignment = AlignmentType.CENTER;
      if (list?.task) {
        return [new Paragraph({ ...opts, indent: { left: 360 + list.level * 360, hanging: 360 }, children: [new TextRun({ text: `${(list as ListCtx & { checked?: boolean }).checked ? '☒' : '☐'}\t`, font: 'Segoe UI Symbol' }), ...kids] })];
      }
      return [new Paragraph({ ...opts, numbering, children: kids })];
    }
    case 'heading': {
      const level = Math.min(6, Math.max(1, Number(n.attrs?.level) || 1));
      const kids = inlineChildren(n.content, ctx);
      const text = (n.content ?? []).map((c) => c.text ?? '').join('');
      const slug = slugifyHeading(text, ctx.headingSlugs);
      return [new Paragraph({ ...paraOptions(n), heading: HEADINGS[level - 1], numbering, children: [new Bookmark({ id: slug, children: kids.filter((k): k is TextRun => k instanceof TextRun) }), ...kids.filter((k) => !(k instanceof TextRun))] })];
    }
    case 'bulletList':
    case 'orderedList': {
      const ordered = n.type === 'orderedList';
      const ref = listRef(ctx, ordered, (n.attrs?.listStyle as string | null) ?? null, Number(n.attrs?.start) || 1);
      const level = list ? list.level + 1 : 0;
      const instance = list && list.ref === ref ? list.instance : ++ctx.listInstance;
      const inner: ListCtx = { ref, level, instance };
      return (n.content ?? []).flatMap((item) => blocks(item.content, ctx, inner));
    }
    case 'taskList': {
      const level = list ? list.level + 1 : 0;
      return (n.content ?? []).flatMap((item) => blocks(item.content, ctx, { ref: '', level, instance: 0, task: true, ...({ checked: Boolean(item.attrs?.checked) } as object) }));
    }
    case 'blockquote':
      return (n.content ?? []).flatMap((c) =>
        block({ ...c, attrs: { ...(c.attrs ?? {}), indent: 48 } }, ctx, list).map((p) => p),
      );
    case 'codeBlock': {
      const lines = (n.content ?? []).map((c) => c.text ?? '').join('').split('\n');
      return [
        new Paragraph({
          shading: { type: ShadingType.CLEAR, fill: 'F5F7FB', color: 'auto' },
          spacing: { before: 120, after: 160, line: 240, lineRule: LineRuleType.AUTO },
          children: lines.map((l, i) => new TextRun({ text: l, font: 'Courier New', size: 20, break: i > 0 ? 1 : 0 })),
        }),
      ];
    }
    case 'horizontalRule':
      return [new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'B8C0CF', space: 1 } }, spacing: { after: 160 }, children: [] })];
    case 'pageBreak':
      return [new Paragraph({ children: [new PageBreak()] })];
    case 'tableOfContents': {
      const heads = collectHeadingsFromJson(ctx.model.content as JSONContent, 6, true)
        .map((h, i) => ({ ...h, page: ctx.headingPageList[i] }))
        .filter((h) => h.level <= 3 && h.text);
      return [
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(String(n.attrs?.title ?? 'Contents'))] }),
        new TableOfContents(String(n.attrs?.title ?? 'Contents'), {
          hyperlink: true,
          headingStyleRange: '1-3',
          cachedEntries: heads.map((h) => ({ title: h.text, level: h.level, page: h.page })),
        }),
      ];
    }
    case 'mathBlock': {
      const m = ctx.mathImages.get(String(n.attrs?.latex));
      return [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 120, after: 120 },
          children: m ? [new ImageRun({ type: 'png', data: m.data, transformation: { width: m.w, height: m.h } })] : [new TextRun({ text: String(n.attrs?.latex ?? ''), font: 'Cambria Math', italics: true })],
        }),
      ];
    }
    case 'table':
      return [table(n, ctx)];
    default:
      return n.content ? blocks(n.content, ctx, list) : [];
  }
}

/* ------------------------------------------------------------- tables */

const TABLE_THEMES: Record<string, { head?: string; headText?: string; band?: string; border?: string; noBorder?: boolean; minimal?: boolean }> = {
  accent: { head: '2F6DFF', headText: 'FFFFFF', border: 'C9D6F3' },
  banded: { band: 'F2F5FB', border: 'D6DBE5', minimal: true },
  minimal: { border: 'D6DBE5', minimal: true },
  plain: { noBorder: true },
  dark: { head: '1F2937', headText: 'FFFFFF', band: 'F3F4F6', border: 'C4C9D2' },
  green: { head: '17A35A', headText: 'FFFFFF', border: 'BFE6CF' },
  orange: { head: 'F26A26', headText: 'FFFFFF', border: 'F7CDB5' },
};

function table(n: JSONContent, ctx: Ctx): Table {
  const theme = TABLE_THEMES[String(n.attrs?.styleName ?? '')] ?? { border: '9AA4B5', head: 'F1F4F9' };
  const rows = n.content ?? [];
  // column widths from the first row
  const colWidths: number[] = [];
  for (const cell of rows[0]?.content ?? []) {
    const span = Number(cell.attrs?.colspan) || 1;
    const cw = (cell.attrs?.colwidth as number[] | null) ?? [];
    for (let k = 0; k < span; k++) colWidths.push(cw[k] ? cw[k] * PX_TO_TWIP : 0);
  }
  const known = colWidths.filter(Boolean).reduce((a, b) => a + b, 0);
  const unknown = colWidths.filter((w) => !w).length;
  const rest = Math.max(0, ctx.contentWidthTwip - known);
  const widths = colWidths.map((w) => w || Math.round(unknown ? rest / unknown : ctx.contentWidthTwip / colWidths.length));
  const scale = widths.reduce((a, b) => a + b, 0) > ctx.contentWidthTwip * 1.02 ? ctx.contentWidthTwip / widths.reduce((a, b) => a + b, 0) : 1;
  const finalWidths = widths.map((w) => Math.round(w * scale));

  const border = theme.noBorder
    ? { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }
    : { style: BorderStyle.SINGLE, size: 4, color: theme.border ?? '9AA4B5' };
  const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };

  const outRows = rows.map((row, ri) => {
    const isHeader = (row.content ?? []).every((c) => c.type === 'tableHeader');
    let col = 0;
    const cells = (row.content ?? []).map((cell) => {
      const span = Number(cell.attrs?.colspan) || 1;
      const width = finalWidths.slice(col, col + span).reduce((a, b) => a + b, 0);
      col += span;
      const header = cell.type === 'tableHeader';
      const fill = hex(cell.attrs?.backgroundColor as string | null) ?? (header ? theme.head : theme.band && ri % 2 === 1 ? theme.band : undefined);
      const valign = cell.attrs?.valign === 'middle' ? VerticalAlign.CENTER : cell.attrs?.valign === 'bottom' ? VerticalAlign.BOTTOM : VerticalAlign.TOP;
      const children = blocks(cell.content, ctx).filter((b): b is Paragraph | Table => !(b instanceof TableOfContents));
      const base = header ? { bold: true, color: theme.headText } : {};
      // re-create header paragraphs with bold text
      const kids = header
        ? (cell.content ?? []).map((p) => new Paragraph({ ...paraOptions(p), children: inlineChildren(p.content, ctx, base) }))
        : children;
      return new TableCell({
        children: kids.length ? kids : [new Paragraph('')],
        columnSpan: span > 1 ? span : undefined,
        rowSpan: Number(cell.attrs?.rowspan) > 1 ? Number(cell.attrs?.rowspan) : undefined,
        width: { size: width, type: WidthType.DXA },
        shading: fill ? { type: ShadingType.CLEAR, fill, color: 'auto' } : undefined,
        verticalAlign: valign,
        margins: { top: 50, bottom: 50, left: 100, right: 100 },
        borders: theme.minimal
          ? { top: noBorder, left: noBorder, right: noBorder, bottom: { style: BorderStyle.SINGLE, size: header ? 12 : 4, color: header ? '3C4454' : theme.border ?? 'D6DBE5' } }
          : { top: border, bottom: border, left: border, right: border },
      });
    });
    return new TableRow({ children: cells, tableHeader: isHeader && ri === 0 });
  });

  return new Table({
    rows: outRows,
    width: { size: finalWidths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: finalWidths,
    layout: TableLayoutType.FIXED,
  });
}

/* ------------------------------------------------------ header / footer */

function fieldRuns(text: string, title: string): TextRun[] {
  const out: TextRun[] = [];
  const re = /\{(PAGE|PAGES|DATE|TITLE)\}/gi;
  let last = 0;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text))) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), size: 18, color: '555555' }));
    const f = m[1].toUpperCase();
    if (f === 'PAGE') out.push(new TextRun({ children: [PageNumber.CURRENT], size: 18, color: '555555' }));
    else if (f === 'PAGES') out.push(new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, color: '555555' }));
    else if (f === 'DATE') out.push(new TextRun({ text: new Date().toLocaleDateString(), size: 18, color: '555555' }));
    else out.push(new TextRun({ text: title, size: 18, color: '555555' }));
    last = m.index + m[0].length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), size: 18, color: '555555' }));
  return out;
}

function hfParagraph(hf: HeaderFooter, contentWidth: number, title: string): Paragraph {
  return new Paragraph({
    tabStops: [
      { type: TabStopType.CENTER, position: Math.round(contentWidth / 2) },
      { type: TabStopType.RIGHT, position: contentWidth },
    ],
    children: [...fieldRuns(hf.left, title), new TextRun({ children: [new Tab()] }), ...fieldRuns(hf.center, title), new TextRun({ children: [new Tab()] }), ...fieldRuns(hf.right, title)],
  });
}

const hasHF = (h: HeaderFooter) => Boolean(h.left.trim() || h.center.trim() || h.right.trim());

/* ------------------------------------------------------------- styles */

function styleRun(s: StyleDef) {
  return {
    font: s.font,
    size: s.size ? Math.round(s.size * 2) : undefined,
    color: hex(s.color),
    bold: s.bold,
    italics: s.italic,
    allCaps: s.caps,
    characterSpacing: s.letterSpacing ? Math.round(s.letterSpacing * 20) : undefined,
  };
}

function styleParagraph(s: StyleDef) {
  return {
    spacing: {
      before: s.before !== undefined ? Math.round(s.before * 20) : undefined,
      after: s.after !== undefined ? Math.round(s.after * 20) : undefined,
      line: s.lineHeight ? Math.round(s.lineHeight * 240) : undefined,
      lineRule: s.lineHeight ? LineRuleType.AUTO : undefined,
    },
    alignment: s.align === 'center' ? AlignmentType.CENTER : s.align === 'right' ? AlignmentType.RIGHT : s.align === 'justify' ? AlignmentType.JUSTIFIED : undefined,
  };
}

/* -------------------------------------------------------------- assets */

async function collectAssets(content: JSONContent, ctx: Ctx) {
  const imgs = new Set<string>();
  const maths = new Set<string>();
  const walk = (n: JSONContent) => {
    if (n.type === 'image' && n.attrs?.src) imgs.add(String(n.attrs.src));
    if ((n.type === 'mathInline' || n.type === 'mathBlock') && n.attrs?.latex) maths.add(String(n.attrs.latex));
    n.content?.forEach(walk);
  };
  walk(content);
  for (const src of imgs) {
    try {
      let url = src;
      let parsed = dataUrlToBytes(url);
      const supported = parsed && /image\/(png|jpeg|gif|bmp)/.test(parsed.mime);
      if (!supported) {
        const r = await rasterizeToPng(src);
        url = r.dataUrl;
        parsed = dataUrlToBytes(url);
      }
      if (!parsed) continue;
      const size = await imageSize(url);
      const type = parsed.mime === 'image/jpeg' ? 'jpg' : parsed.mime === 'image/gif' ? 'gif' : parsed.mime === 'image/bmp' ? 'bmp' : 'png';
      ctx.images.set(src, { data: parsed.bytes, type, w: size.width, h: size.height });
    } catch {
      /* skip unreadable image */
    }
  }
  if (maths.size) {
    const { renderMathPng } = await import('./shared');
    for (const latex of maths) {
      const r = await renderMathPng(latex);
      if (r) ctx.mathImages.set(latex, r);
    }
  }
}

/* --------------------------------------------------------------- entry */

export async function exportDocx(model: DocModel, title: string, headingPages: Map<number, number> = new Map()): Promise<Uint8Array> {
  const s = model.settings;
  const styles = resolveStyles(s);
  const contentWidth = Math.round((s.page.orientation === 'landscape' ? s.page.height : s.page.width) * 1440 - (s.page.margins.left + s.page.margins.right) * 1440);
  const ctx: Ctx = {
    model,
    commentIds: new Map((model.comments ?? []).map((c, i) => [c.id, i + 1])),
    images: new Map(),
    mathImages: new Map(),
    numberingRefs: new Map(),
    listInstance: 0,
    headingSlugs: new Map(),
    contentWidthTwip: contentWidth,
    headingPageList: Array.from(headingPages.entries())
      .sort((a, b) => a[0] - b[0])
      .map((e) => e[1]),
  };
  const content = model.content as JSONContent;
  await collectAssets(content, ctx);
  const children = blocks(content.content, ctx);

  const BULLETS: Record<string, [string, string]> = {
    disc: ['●', 'Arial'],
    circle: ['○', 'Arial'],
    square: ['■', 'Arial'],
    dash: ['–', 'Arial'],
    check: ['✓', 'Segoe UI Symbol'],
    arrow: ['➤', 'Segoe UI Symbol'],
  };
  const nested = ['●', '○', '■'];
  const FORMATS: Record<string, (typeof LevelFormat)[keyof typeof LevelFormat]> = {
    decimal: LevelFormat.DECIMAL,
    'lower-alpha': LevelFormat.LOWER_LETTER,
    'upper-alpha': LevelFormat.UPPER_LETTER,
    'lower-roman': LevelFormat.LOWER_ROMAN,
    'upper-roman': LevelFormat.UPPER_ROMAN,
  };
  const nestedOrder = [LevelFormat.DECIMAL, LevelFormat.LOWER_LETTER, LevelFormat.LOWER_ROMAN];
  const numberingConfig = Array.from(ctx.numberingRefs.entries()).map(([reference, info]) => ({
    reference,
    levels: Array.from({ length: 9 }, (_, level) => {
      const indent = { left: 360 + level * 360, hanging: 360 };
      if (info.format === 'bullet') {
        const [char, font] = level === 0 ? BULLETS[info.bullet ?? 'disc'] ?? BULLETS.disc : [nested[level % 3], 'Arial'];
        return { level, format: LevelFormat.BULLET, text: char, alignment: AlignmentType.LEFT, style: { paragraph: { indent }, run: { font } } };
      }
      const fmt = level === 0 ? FORMATS[info.format] ?? LevelFormat.DECIMAL : nestedOrder[level % 3];
      return { level, format: fmt, text: `%${level + 1}.`, start: level === 0 ? info.start : 1, alignment: AlignmentType.LEFT, style: { paragraph: { indent } } };
    }),
  }));

  const hasHeader = hasHF(s.header);
  const hasFooter = hasHF(s.footer);
  const docTitle = s.title || title.replace(/\.[^.]+$/, '');

  const doc = new Document({
    creator: s.author || 'Affice',
    title: s.title || undefined,
    subject: s.subject || undefined,
    keywords: s.keywords || undefined,
    description: 'Created with Affice',
    lastModifiedBy: s.author || 'Affice',
    background: s.pageColor ? { color: hex(s.pageColor) } : undefined,
    defaultTabStop: 720,
    comments: {
      children: (model.comments ?? []).map((c, i) => ({
        id: i + 1,
        author: c.author || 'Affice',
        date: new Date(c.date),
        initials: (c.author || 'A').slice(0, 2).toUpperCase(),
        children: [new Paragraph(c.text), ...(c.replies ?? []).map((r) => new Paragraph(`${r.author}: ${r.text}`))],
      })),
    },
    styles: {
      default: {
        document: {
          run: { font: styles.normal.font, size: Math.round((styles.normal.size ?? 11) * 2), color: hex(styles.normal.color) },
          paragraph: {
            spacing: {
              after: Math.round((styles.normal.after ?? 8) * 20),
              before: Math.round((styles.normal.before ?? 0) * 20),
              line: Math.round((styles.normal.lineHeight ?? 1.15) * 240),
              lineRule: LineRuleType.AUTO,
            },
          },
        },
        title: { run: styleRun(styles.title), paragraph: styleParagraph(styles.title) },
        heading1: { run: styleRun(styles.h1), paragraph: { ...styleParagraph(styles.h1), keepNext: true } },
        heading2: { run: styleRun(styles.h2), paragraph: { ...styleParagraph(styles.h2), keepNext: true } },
        heading3: { run: styleRun(styles.h3), paragraph: { ...styleParagraph(styles.h3), keepNext: true } },
        heading4: { run: styleRun(styles.h4), paragraph: { ...styleParagraph(styles.h4), keepNext: true } },
        heading5: { run: styleRun(styles.h5), paragraph: { ...styleParagraph(styles.h5), keepNext: true } },
        heading6: { run: styleRun(styles.h6), paragraph: { ...styleParagraph(styles.h6), keepNext: true } },
        hyperlink: { run: { color: '0563C1', underline: { type: 'single' } } },
      },
      paragraphStyles: [
        { id: 'Subtitle', name: 'Subtitle', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: styleRun(styles.subtitle), paragraph: styleParagraph(styles.subtitle) },
        {
          id: 'Quote',
          name: 'Quote',
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: styleRun(styles.quote),
          paragraph: { ...styleParagraph(styles.quote), indent: { left: 864, right: 864 } },
        },
        { id: 'NoSpacing', name: 'No Spacing', basedOn: 'Normal', quickFormat: true, paragraph: { spacing: { before: 0, after: 0, line: 240, lineRule: LineRuleType.AUTO } } },
      ],
    },
    numbering: { config: numberingConfig },
    sections: [
      {
        properties: {
          titlePage: s.differentFirstPage,
          page: {
            size: {
              width: Math.round(s.page.width * 1440),
              height: Math.round(s.page.height * 1440),
              orientation: s.page.orientation === 'landscape' ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT,
            },
            margin: {
              top: Math.round(s.page.margins.top * 1440),
              right: Math.round(s.page.margins.right * 1440),
              bottom: Math.round(s.page.margins.bottom * 1440),
              left: Math.round(s.page.margins.left * 1440),
              header: Math.round(Math.min(0.5, s.page.margins.top / 2) * 1440),
              footer: Math.round(Math.min(0.5, s.page.margins.bottom / 2) * 1440),
            },
          },
        },
        headers: hasHeader
          ? { default: new Header({ children: [hfParagraph(s.header, contentWidth, docTitle)] }), first: s.differentFirstPage ? new Header({ children: [new Paragraph('')] }) : undefined }
          : undefined,
        footers: hasFooter
          ? { default: new Footer({ children: [hfParagraph(s.footer, contentWidth, docTitle)] }), first: s.differentFirstPage ? new Footer({ children: [new Paragraph('')] }) : undefined }
          : undefined,
        children: children.length ? children : [new Paragraph('')],
      },
    ],
  });
  const buf = await Packer.toArrayBuffer(doc);
  return new Uint8Array(buf);
}
