import type { JSONContent } from '@tiptap/core';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';

export interface Margins {
  top: number;
  right: number;
  bottom: number;
  left: number;
}

export interface PageSetup {
  /** Paper size id (letter, a4…) or 'custom'. */
  size: string;
  /** Portrait dimensions in inches (swapped when landscape). */
  width: number;
  height: number;
  orientation: 'portrait' | 'landscape';
  margins: Margins;
}

export interface HeaderFooter {
  left: string;
  center: string;
  right: string;
}

export interface StyleDef {
  font?: string;
  size?: number; // pt
  color?: string;
  bold?: boolean;
  italic?: boolean;
  before?: number; // pt
  after?: number; // pt
  lineHeight?: number;
  align?: 'left' | 'center' | 'right' | 'justify';
  caps?: boolean;
  letterSpacing?: number; // pt
}

export type StyleId = 'normal' | 'title' | 'subtitle' | 'h1' | 'h2' | 'h3' | 'h4' | 'h5' | 'h6' | 'quote' | 'code';

export interface DocSettings {
  page: PageSetup;
  header: HeaderFooter;
  footer: HeaderFooter;
  /** Hide header/footer on the first page (e.g. title pages). */
  differentFirstPage: boolean;
  styleSet: string;
  styles: Partial<Record<StyleId, StyleDef>>;
  pageColor: string | null;
  watermark: string;
  title: string;
  author: string;
  subject: string;
  keywords: string;
  created: string;
  modified: string;
}

export interface PaperSize {
  id: string;
  label: string;
  width: number;
  height: number;
}

export const PAPER_SIZES: PaperSize[] = [
  { id: 'letter', label: 'Letter (8.5" × 11")', width: 8.5, height: 11 },
  { id: 'legal', label: 'Legal (8.5" × 14")', width: 8.5, height: 14 },
  { id: 'a4', label: 'A4 (21 × 29.7 cm)', width: 8.2677, height: 11.6929 },
  { id: 'a5', label: 'A5 (14.8 × 21 cm)', width: 5.8268, height: 8.2677 },
  { id: 'a3', label: 'A3 (29.7 × 42 cm)', width: 11.6929, height: 16.5354 },
  { id: 'b5', label: 'B5 (17.6 × 25 cm)', width: 6.9291, height: 9.8425 },
  { id: 'executive', label: 'Executive (7.25" × 10.5")', width: 7.25, height: 10.5 },
  { id: 'tabloid', label: 'Tabloid (11" × 17")', width: 11, height: 17 },
];

export const MARGIN_PRESETS: Array<{ id: string; label: string; m: Margins }> = [
  { id: 'normal', label: 'Normal', m: { top: 1, right: 1, bottom: 1, left: 1 } },
  { id: 'narrow', label: 'Narrow', m: { top: 0.5, right: 0.5, bottom: 0.5, left: 0.5 } },
  { id: 'moderate', label: 'Moderate', m: { top: 1, right: 0.75, bottom: 1, left: 0.75 } },
  { id: 'wide', label: 'Wide', m: { top: 1, right: 2, bottom: 1, left: 2 } },
  { id: 'mirrored', label: 'Book', m: { top: 1, right: 1, bottom: 1, left: 1.25 } },
];

/** Picks a sensible default paper size for the user's locale (Letter in the Americas, A4 elsewhere). */
export function defaultPaper(): PaperSize {
  const lang = (typeof navigator !== 'undefined' ? navigator.language : 'en-US') || 'en-US';
  const region = lang.split('-')[1]?.toUpperCase();
  const letter = ['US', 'CA', 'MX', 'PH', 'CL', 'CO', 'VE', 'CR', 'GT', 'PA', 'DO', 'PR', 'SV'];
  return PAPER_SIZES.find((p) => p.id === (region && letter.includes(region) ? 'letter' : region ? 'a4' : 'letter'))!;
}

export function pageDims(p: PageSetup): { w: number; h: number } {
  return p.orientation === 'landscape' ? { w: p.height, h: p.width } : { w: p.width, h: p.height };
}

/* ----------------------------------------------------------------- styles */

export interface StyleSet {
  id: string;
  label: string;
  styles: Record<StyleId, StyleDef>;
}

const base = (body: string, head: string, headColor: string, extra: Partial<Record<StyleId, StyleDef>> = {}): Record<StyleId, StyleDef> => ({
  normal: { font: body, size: 11, color: '#000000', after: 8, lineHeight: 1.15 },
  title: { font: head, size: 28, color: '#1a1a1a', after: 4, lineHeight: 1.05, letterSpacing: -0.5 },
  subtitle: { font: body, size: 13, color: '#5a5a5a', after: 12 },
  h1: { font: head, size: 16, color: headColor, bold: false, before: 16, after: 4, lineHeight: 1.1 },
  h2: { font: head, size: 13, color: headColor, bold: false, before: 10, after: 3, lineHeight: 1.1 },
  h3: { font: head, size: 12, color: '#1f3763', bold: false, before: 8, after: 2 },
  h4: { font: head, size: 11, color: headColor, bold: true, italic: true, before: 6, after: 2 },
  h5: { font: head, size: 11, color: headColor, bold: true, before: 6, after: 2 },
  h6: { font: head, size: 11, color: '#1f3763', italic: true, before: 6, after: 2 },
  quote: { font: body, size: 11, color: '#404040', italic: true, before: 10, after: 10, align: 'center' },
  code: { font: 'Courier New', size: 10, color: '#1f2937' },
  ...extra,
});

export const STYLE_SETS: StyleSet[] = [
  { id: 'office', label: 'Office', styles: base('Calibri', 'Calibri', '#2f5496') },
  {
    id: 'classic',
    label: 'Classic',
    styles: base('Cambria', 'Cambria', '#17365d', {
      h1: { font: 'Cambria', size: 16, color: '#17365d', bold: true, before: 18, after: 6 },
      h2: { font: 'Cambria', size: 13, color: '#17365d', bold: true, before: 12, after: 4 },
    }),
  },
  {
    id: 'modern',
    label: 'Modern',
    styles: base('Arial', 'Arial', '#004fff', {
      title: { font: 'Arial', size: 30, color: '#0b1b4d', bold: true, after: 6, letterSpacing: -0.8 },
      h1: { font: 'Arial', size: 17, color: '#004fff', bold: true, before: 18, after: 6 },
      h2: { font: 'Arial', size: 13, color: '#0b1b4d', bold: true, before: 12, after: 4 },
      h3: { font: 'Arial', size: 11.5, color: '#004fff', bold: true, before: 10, after: 2, caps: true, letterSpacing: 0.5 },
    }),
  },
  {
    id: 'elegant',
    label: 'Elegant',
    styles: base('Georgia', 'Georgia', '#7a1f2b', {
      normal: { font: 'Georgia', size: 11, color: '#222222', after: 10, lineHeight: 1.35 },
      title: { font: 'Georgia', size: 30, color: '#7a1f2b', italic: true, after: 6 },
      h1: { font: 'Georgia', size: 17, color: '#7a1f2b', before: 18, after: 6 },
      h2: { font: 'Georgia', size: 13.5, color: '#7a1f2b', italic: true, before: 12, after: 4 },
    }),
  },
  {
    id: 'academic',
    label: 'Academic',
    styles: base('Times New Roman', 'Times New Roman', '#000000', {
      normal: { font: 'Times New Roman', size: 12, color: '#000000', after: 0, lineHeight: 2 },
      title: { font: 'Times New Roman', size: 12, color: '#000000', align: 'center', after: 0, lineHeight: 2 },
      subtitle: { font: 'Times New Roman', size: 12, color: '#000000', align: 'center', after: 0 },
      h1: { font: 'Times New Roman', size: 12, color: '#000000', bold: true, align: 'center', before: 0, after: 0, lineHeight: 2 },
      h2: { font: 'Times New Roman', size: 12, color: '#000000', bold: true, before: 0, after: 0, lineHeight: 2 },
      h3: { font: 'Times New Roman', size: 12, color: '#000000', bold: true, italic: true, before: 0, after: 0, lineHeight: 2 },
    }),
  },
  {
    id: 'minimal',
    label: 'Minimal',
    styles: base('Calibri', 'Calibri', '#111111', {
      title: { font: 'Calibri', size: 26, color: '#111111', after: 4, letterSpacing: -0.3 },
      h1: { font: 'Calibri', size: 15, color: '#111111', bold: true, before: 16, after: 4 },
      h2: { font: 'Calibri', size: 12.5, color: '#444444', bold: true, before: 10, after: 2 },
      h3: { font: 'Calibri', size: 11, color: '#666666', bold: true, before: 8, after: 2, caps: true, letterSpacing: 0.6 },
    }),
  },
  {
    id: 'vibrant',
    label: 'Vibrant',
    styles: base('Calibri', 'Arial', '#e24c12', {
      title: { font: 'Arial', size: 30, color: '#7c3aed', bold: true, after: 4 },
      subtitle: { font: 'Calibri', size: 13, color: '#e24c12', after: 12 },
      h1: { font: 'Arial', size: 17, color: '#7c3aed', bold: true, before: 18, after: 6 },
      h2: { font: 'Arial', size: 13, color: '#e24c12', bold: true, before: 12, after: 4 },
    }),
  },
];

export function resolveStyles(s: Pick<DocSettings, 'styleSet' | 'styles'>): Record<StyleId, StyleDef> {
  const set = STYLE_SETS.find((x) => x.id === s.styleSet) ?? STYLE_SETS[0];
  const out = { ...set.styles };
  for (const [k, v] of Object.entries(s.styles ?? {})) {
    out[k as StyleId] = { ...out[k as StyleId], ...v };
  }
  return out;
}

/** CSS custom properties that drive the document content stylesheet (content.css). */
export function styleVars(s: Pick<DocSettings, 'styleSet' | 'styles'>): Record<string, string> {
  const st = resolveStyles(s);
  const vars: Record<string, string> = {};
  for (const [id, d] of Object.entries(st)) {
    const p = `--${id}`;
    if (d.font) vars[`${p}-font`] = `${JSON.stringify(d.font)}, ${/Times|Cambria|Georgia|Garamond|Serif|Lora|Merriweather|Playfair/i.test(d.font) ? 'serif' : /Courier|Mono|Code/i.test(d.font) ? 'monospace' : 'sans-serif'}`;
    if (d.size) vars[`${p}-size`] = `${d.size}pt`;
    if (d.color) vars[`${p}-color`] = d.color;
    vars[`${p}-weight`] = d.bold ? '700' : id.startsWith('h') ? '400' : 'inherit';
    vars[`${p}-style`] = d.italic ? 'italic' : 'normal';
    if (d.before !== undefined) vars[`${p}-before`] = `${d.before}pt`;
    if (d.after !== undefined) vars[`${p}-after`] = `${d.after}pt`;
    if (d.lineHeight !== undefined) vars[`${p}-line`] = String(d.lineHeight);
    if (d.align) vars[`${p}-align`] = d.align;
    vars[`${p}-caps`] = d.caps ? 'uppercase' : 'none';
    vars[`${p}-spacing`] = d.letterSpacing ? `${d.letterSpacing}pt` : 'normal';
  }
  return vars;
}

export function styleVarsCss(s: Pick<DocSettings, 'styleSet' | 'styles'>): string {
  return Object.entries(styleVars(s))
    .map(([k, v]) => `${k}:${v};`)
    .join('');
}

/* --------------------------------------------------------------- defaults */

export function defaultSettings(): DocSettings {
  const paper = defaultPaper();
  const now = new Date().toISOString();
  return {
    page: { size: paper.id, width: paper.width, height: paper.height, orientation: 'portrait', margins: { top: 1, right: 1, bottom: 1, left: 1 } },
    header: { left: '', center: '', right: '' },
    footer: { left: '', center: '', right: '' },
    differentFirstPage: false,
    styleSet: 'office',
    styles: {},
    pageColor: null,
    watermark: '',
    title: '',
    author: '',
    subject: '',
    keywords: '',
    created: now,
    modified: now,
  };
}

export const EMPTY_DOC: JSONContent = { type: 'doc', content: [{ type: 'paragraph' }] };

export interface DocModel {
  settings: DocSettings;
  content: JSONContent;
  comments?: DocComment[];
}

export interface DocComment {
  id: string;
  author: string;
  date: string;
  text: string;
  resolved?: boolean;
  replies?: Array<{ author: string; date: string; text: string }>;
}

/* ----------------------------------------------------- native format (.afdoc) */

const AFDOC_MIME = 'application/x-affice-document';

export function serializeAfdoc(model: DocModel): Uint8Array {
  return zipSync(
    {
      mimetype: [strToU8(AFDOC_MIME), { level: 0 }],
      'document.json': strToU8(JSON.stringify({ format: 'affice-document', version: 1, ...model })),
    },
    { level: 6 },
  );
}

export function parseAfdoc(data: Uint8Array): DocModel {
  let json: string;
  if (data[0] === 0x50 && data[1] === 0x4b) {
    const files = unzipSync(data);
    const f = files['document.json'];
    if (!f) throw new Error('This .afdoc file is damaged (document.json is missing).');
    json = strFromU8(f);
  } else json = strFromU8(data);
  const raw = JSON.parse(json);
  return {
    settings: { ...defaultSettings(), ...(raw.settings ?? {}) },
    content: raw.content ?? EMPTY_DOC,
    comments: raw.comments ?? [],
  };
}

/** Replaces {PAGE}, {PAGES}, {DATE}, {TITLE} fields in header/footer text. */
export function expandFields(text: string, page: number, pages: number, title: string): string {
  return text
    .replace(/\{PAGE\}/gi, String(page))
    .replace(/\{PAGES\}/gi, String(pages))
    .replace(/\{DATE\}/gi, new Date().toLocaleDateString())
    .replace(/\{TITLE\}/gi, title);
}
