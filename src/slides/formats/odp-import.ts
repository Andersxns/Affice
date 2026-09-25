/**
 * OpenDocument presentation reader (.odp, .otp and flat .fodp).
 *
 * Master pages become slide layouts (background, decorations and placeholders); slides keep their shapes,
 * text, pictures, tables, charts, notes, transitions and animations. Formatting is resolved through
 * ODF's style inheritance (default style → parents → automatic style) and then reduced to what differs
 * from the presentation's own defaults, with LibreOffice theme colours turned back into theme
 * references, so imported slides stay easy to edit and restyle.
 */
import { strFromU8, unzipSync } from 'fflate';
import { bytesToBase64 } from '@/lib/utils';
import { DEFAULT_ADJ, isKnownPreset } from '../geometry';
import {
  colorRef,
  newId,
  parseColorRef,
  resolveHex,
  rotatedBounds,
  unionBounds,
  type Anim,
  type AnimEffect,
  type ArrowType,
  type Bullet,
  type ChartEl,
  type CustomPath,
  type Dash,
  type Dir,
  type El,
  type Fill,
  type GradientStop,
  type GroupEl,
  type ImageEl,
  type Layout,
  type LayoutType,
  type Line,
  type NumStyle,
  type Para,
  type PlaceholderType,
  type Presentation,
  type Run,
  type Shadow,
  type ShapeEl,
  type Slide,
  type TableCell,
  type TableEl,
  type TableStyleOpts,
  type TextBody,
  type Theme,
  type ThemeColorName,
  type Transition,
  type TransitionType,
} from '../model';
import { DEFAULT_INSET, effectivePara, effectiveRun, roleOf, type TextContext } from '../render/text';
import { cellLook, type CellLook } from '../tables';
import { createPresentation } from '../themes';
import type { LoadedPresentation } from './index';
import { readSlideChart } from './odp-charts';
import { evaluateEnhancedGeometry, svgPathSegments } from './odp-geometry';
import { imageSize } from './media';

const NS = {
  office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
  style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
  text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
  table: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
  draw: 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
  fo: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
  xlink: 'http://www.w3.org/1999/xlink',
  dc: 'http://purl.org/dc/elements/1.1/',
  meta: 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
  presentation: 'urn:oasis:names:tc:opendocument:xmlns:presentation:1.0',
  svg: 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
  smil: 'urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0',
  anim: 'urn:oasis:names:tc:opendocument:xmlns:animation:1.0',
  loext: 'urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0',
  drawooo: 'http://openoffice.org/2010/draw',
  xml: 'http://www.w3.org/XML/1998/namespace',
  affice: 'urn:org:affice:xmlns:slides:1.0',
  officeooo: 'http://openoffice.org/2009/office',
  number: 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0',
};

const PREFIX: Record<string, string> = Object.fromEntries(Object.entries(NS).map(([k, v]) => [v, k]));

/* ================================================================ helpers */

function kids(el: Element | Document | null | undefined): Element[] {
  const out: Element[] = [];
  if (!el) return out;
  for (let n = el.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) out.push(n as Element);
  return out;
}

const kid = (el: Element | Document | null | undefined, ns: string, name: string) => kids(el).find((k) => k.namespaceURI === ns && k.localName === name);
const kidsOf = (el: Element | Document | null | undefined, ns: string, name: string) => kids(el).filter((k) => k.namespaceURI === ns && k.localName === name);
const at = (el: Element | null | undefined, ns: string, name: string): string | undefined => (el && el.hasAttributeNS(ns, name) ? (el.getAttributeNS(ns, name) ?? undefined) : undefined);

/** A length in CSS pixels (96 dpi). */
export function lengthPx(v: string | undefined): number | undefined {
  if (!v) return undefined;
  const m = /^\s*([-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?)\s*(cm|mm|in|inch|pt|pc|px)?\s*$/i.exec(v);
  if (!m) return undefined;
  const n = Number(m[1]);
  switch ((m[2] ?? '').toLowerCase()) {
    case 'cm':
      return (n * 96) / 2.54;
    case 'mm':
      return (n * 96) / 25.4;
    case 'in':
    case 'inch':
      return n * 96;
    case 'pt':
      return (n * 96) / 72;
    case 'pc':
      return n * 16;
    default:
      return n;
  }
}

const ptOf = (v: string | undefined) => {
  const px = lengthPx(v);
  return px === undefined ? undefined : Math.round(px * 0.75 * 100) / 100;
};

const pctOf = (v: string | undefined) => {
  const m = v ? /^\s*([-+]?[\d.]+)\s*%\s*$/.exec(v) : null;
  return m ? Number(m[1]) / 100 : undefined;
};

const r2 = (v: number) => Math.round(v * 100) / 100;

/** Seconds from a SMIL clock value ("0.5s", "500ms", "00:00:01.5", "2"). */
function clockMs(v: string | undefined): number | undefined {
  if (!v || v === 'indefinite' || v === 'next') return undefined;
  const t = v.trim();
  let m = /^([\d.]+)\s*(ms|s|min|h)?$/.exec(t);
  if (m) {
    const n = Number(m[1]);
    const unit = m[2] ?? 's';
    return unit === 'ms' ? n : unit === 'min' ? n * 60000 : unit === 'h' ? n * 3600000 : n * 1000;
  }
  m = /^(?:(\d+):)?(\d+):([\d.]+)$/.exec(t);
  if (m) return ((Number(m[1] ?? 0) * 60 + Number(m[2])) * 60 + Number(m[3])) * 1000;
  return undefined;
}

/** Milliseconds from an ISO 8601 duration ("PT5S", "PT00H01M02.5S"). */
function durationMs(v: string | undefined): number | undefined {
  const m = v ? /^P(?:\d+D)?T?(?:([\d.]+)H)?(?:([\d.]+)M)?(?:([\d.]+)S)?$/.exec(v.trim()) : null;
  if (!m) return undefined;
  return ((Number(m[1] ?? 0) * 60 + Number(m[2] ?? 0)) * 60 + Number(m[3] ?? 0)) * 1000;
}

function unquoteFont(v: string | undefined): string | undefined {
  if (!v) return undefined;
  const first = v.split(',')[0].trim();
  const f = first.replace(/^['"]|['"]$/g, '').trim();
  return f || undefined;
}

const hex2 = (n: number) => Math.round(Math.max(0, Math.min(255, n))).toString(16).padStart(2, '0');

function normHex(v: string | undefined): string | undefined {
  if (!v) return undefined;
  const m = /^#?([0-9a-f]{6})$/i.exec(v.trim());
  if (m) return `#${m[1].toLowerCase()}`;
  const s = /^#?([0-9a-f]{3})$/i.exec(v.trim());
  if (s) return `#${[...s[1]].map((c) => c + c).join('').toLowerCase()}`;
  return undefined;
}

function closeHex(a: string, b: string, tol = 6): boolean {
  const p = (h: string, i: number) => parseInt(h.slice(1 + i * 2, 3 + i * 2), 16);
  return [0, 1, 2].every((i) => Math.abs(p(a, i) - p(b, i)) <= tol);
}

function withAlpha(hex: string, alpha: number): string {
  return alpha < 0.999 ? `${hex}${hex2(alpha * 255)}` : hex;
}

function mixHex(a: string, b: string, t: number): string {
  const p = (h: string, i: number) => parseInt(h.slice(1 + i * 2, 3 + i * 2), 16);
  return `#${[0, 1, 2].map((i) => hex2(p(a, i) * (1 - t) + p(b, i) * t)).join('')}`;
}

/* ================================================================= styles */

type Props = Map<string, string>;

interface Resolved {
  /** Property containers ("graphic", "paragraph", "text", "drawing-page", "table-cell", …) as prefixed attributes. */
  c: Record<string, Props>;
  /** A list style carried by the style (outline levels, list styles of shapes). */
  list?: Element;
}

const SIZE_KEYS = new Set(['fo:font-size', 'style:font-size-asian', 'style:font-size-complex']);
const FONT_KEYS = [
  ['fo:font-family', 'style:font-name'],
  ['style:font-family-asian', 'style:font-name-asian'],
  ['style:font-family-complex', 'style:font-name-complex'],
];

/** Merges property maps: later values win, percentages of font sizes scale the inherited size. */
function mergeProps(base: Props, over: Props): void {
  for (const [k, v] of over) {
    if (SIZE_KEYS.has(k) && v.trim().endsWith('%')) {
      // a relative size stays relative until it meets an absolute one
      const pct = pctOf(v) ?? 1;
      const b = base.get(k);
      if (b === undefined) base.set(k, v);
      else if (b.trim().endsWith('%')) base.set(k, `${Math.round((pctOf(b) ?? 1) * pct * 10000) / 100}%`);
      else base.set(k, `${r2((ptOf(b) ?? 18) * pct)}pt`);
      continue;
    }
    base.set(k, v);
  }
  // a font set one way replaces a font inherited the other way
  for (const [family, name] of FONT_KEYS) {
    if (over.has(name) && !over.has(family)) base.delete(family);
    if (over.has(family) && !over.has(name)) base.delete(name);
  }
  // a plain colour replaces an inherited theme colour
  if (over.has('fo:color') && !over.has('x:char-complex')) base.delete('x:char-complex');
  if (over.has('draw:fill-color') && !over.has('x:fill-complex')) base.delete('x:fill-complex');
  if (over.has('svg:stroke-color') && !over.has('x:stroke-complex')) base.delete('x:stroke-complex');
}

function cloneResolved(r: Resolved): Resolved {
  const c: Record<string, Props> = {};
  for (const [k, v] of Object.entries(r.c)) c[k] = new Map(v);
  return { c, list: r.list };
}

/** Attributes of a property element, plus theme colours, columns and list styles found inside it. */
function readContainer(el: Element): Props {
  const out: Props = new Map();
  for (let i = 0; i < el.attributes.length; i++) {
    const a = el.attributes[i];
    const prefix = a.namespaceURI ? PREFIX[a.namespaceURI] : undefined;
    if (!prefix || prefix === 'xml') continue;
    out.set(`${prefix}:${a.localName}`, a.value);
  }
  for (const k of kids(el)) {
    if (k.namespaceURI === NS.loext && k.localName.endsWith('-complex-color')) {
      const kind = k.localName.replace('-complex-color', '');
      const type = at(k, NS.loext, 'theme-type');
      if (!type || at(k, NS.loext, 'color-type') !== 'theme') continue;
      const mods = kidsOf(k, NS.loext, 'transformation').map((t) => `${at(t, NS.loext, 'type')}:${at(t, NS.loext, 'value')}`);
      out.set(`x:${kind}-complex`, [type, ...mods].join(';'));
    } else if (k.namespaceURI === NS.style && k.localName === 'columns') {
      const n = Number(at(k, NS.fo, 'column-count') ?? 1);
      if (n > 1) out.set('x:columns', String(n));
    }
  }
  return out;
}

class Styles {
  private defs = new Map<string, Element>();
  private defaults = new Map<string, Element>();
  private lists = new Map<string, Element>();
  readonly named = new Map<string, Element>();
  readonly fonts = new Map<string, string>();
  readonly templates = new Map<string, Element>();
  readonly pageLayouts = new Map<string, Element>();
  readonly masters: Element[] = [];
  readonly presLayouts = new Map<string, Element>();
  private cache = new Map<string, Resolved>();

  /** Collects styles from a styles.xml / content.xml root (or a flat document). */
  add(root: Element, scope: 'content' | 'styles') {
    for (const sec of kids(root)) {
      if (sec.namespaceURI !== NS.office) continue;
      const where = sec.localName === 'styles' ? 'common' : sec.localName === 'automatic-styles' ? scope : sec.localName === 'master-styles' ? 'master' : sec.localName === 'font-face-decls' ? 'fonts' : null;
      if (!where) continue;
      for (const s of kids(sec)) {
        const name = at(s, NS.style, 'name') ?? at(s, NS.draw, 'name') ?? at(s, NS.table, 'name');
        if (where === 'fonts') {
          if (s.localName === 'font-face' && name) this.fonts.set(name, unquoteFont(at(s, NS.svg, 'font-family')) ?? name);
          continue;
        }
        if (where === 'master') {
          // (a flat file is read for both scopes; its master pages count once)
          if (scope === 'styles' && s.namespaceURI === NS.style && s.localName === 'master-page') this.masters.push(s);
          continue;
        }
        if (s.namespaceURI === NS.style && s.localName === 'style' && name) this.defs.set(`${where}|${at(s, NS.style, 'family')}|${name}`, s);
        else if (s.namespaceURI === NS.style && s.localName === 'default-style') this.defaults.set(at(s, NS.style, 'family') ?? '', s);
        else if (s.namespaceURI === NS.text && s.localName === 'list-style' && name) this.lists.set(`${where}|${name}`, s);
        else if (s.namespaceURI === NS.style && s.localName === 'page-layout' && name) this.pageLayouts.set(name, s);
        else if (s.namespaceURI === NS.style && s.localName === 'presentation-page-layout' && name) this.presLayouts.set(name, s);
        else if (s.namespaceURI === NS.table && s.localName === 'table-template' && name) this.templates.set(name, s);
        else if (s.namespaceURI === NS.draw && name) this.named.set(`${s.localName}|${name}`, s);
      }
    }
  }

  element(name: string, family: string, scope: 'content' | 'styles'): Element | undefined {
    return this.defs.get(`${scope}|${family}|${name}`) ?? this.defs.get(`common|${family}|${name}`);
  }

  list(name: string | undefined, scope: 'content' | 'styles'): Element | undefined {
    if (!name) return undefined;
    return this.lists.get(`${scope}|${name}`) ?? this.lists.get(`common|${name}`);
  }

  private base(family: string): Resolved {
    const key = `default|${family}`;
    const hit = this.cache.get(key);
    if (hit) return hit;
    const out: Resolved = { c: {} };
    const el = this.defaults.get(family) ?? (family === 'presentation' ? this.defaults.get('graphic') : undefined);
    if (el) this.apply(out, el, 'styles');
    this.cache.set(key, out);
    return out;
  }

  private apply(out: Resolved, el: Element, scope: 'content' | 'styles') {
    for (const k of kids(el)) {
      if (!k.localName.endsWith('-properties')) continue;
      const kind = k.localName.slice(0, -'-properties'.length);
      const props = readContainer(k);
      if (!out.c[kind]) out.c[kind] = new Map();
      mergeProps(out.c[kind], props);
      const ls = kid(k, NS.text, 'list-style');
      if (ls) out.list = ls;
    }
    const ln = at(el, NS.style, 'list-style-name');
    if (ln) out.list = this.list(ln, scope) ?? out.list;
  }

  resolve(name: string | undefined, family: string, scope: 'content' | 'styles', depth = 0): Resolved {
    const key = `${scope}|${family}|${name ?? ''}`;
    const hit = this.cache.get(key);
    if (hit) return hit;
    const el = name ? this.element(name, family, scope) : undefined;
    const parent = el ? at(el, NS.style, 'parent-style-name') : undefined;
    const base = parent && parent !== name && depth < 20 ? this.resolve(parent, family, scope, depth + 1) : this.base(family);
    const out = cloneResolved(base);
    if (el) this.apply(out, el, scope);
    this.cache.set(key, out);
    return out;
  }
}

/* ============================================================ constants */

const THEME_IN: Record<string, ThemeColorName> = {
  dark1: 'tx1',
  light1: 'bg1',
  dark2: 'tx2',
  light2: 'bg2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hyperlink: 'hlink',
  'followed-hyperlink': 'folHlink',
};

const NUM_IN: Record<string, NumStyle> = { '1.': 'arabicPeriod', '1)': 'arabicParenR', 'a.': 'alphaLcPeriod', 'A.': 'alphaUcPeriod', 'a)': 'alphaLcParenR', 'i.': 'romanLcPeriod', 'I.': 'romanUcPeriod' };

const PH_IN: Record<string, PlaceholderType> = { title: 'title', subtitle: 'subTitle', outline: 'obj', text: 'body', graphic: 'pic', object: 'obj', chart: 'obj', table: 'obj', orgchart: 'obj', 'date-time': 'dt', footer: 'ftr', 'page-number': 'sldNum' };

const PH_TYPES = new Set<string>(['title', 'ctrTitle', 'subTitle', 'body', 'obj', 'pic', 'dt', 'ftr', 'sldNum']);
const LAYOUT_TYPES = new Set<string>(['title', 'obj', 'secHead', 'twoObj', 'twoTxTwoObj', 'titleOnly', 'blank', 'objTx', 'picTx', 'cust']);
const DASHES = new Set<string>(['dash', 'dot', 'dashDot', 'longDash', 'sysDash', 'sysDot']);

/** PowerPoint's layout names (kept by LibreOffice when it converts .pptx files) and their layout types. */
const LAYOUT_NAMES: [RegExp, LayoutType][] = [
  [/^title slide$/i, 'title'],
  [/^title,? and content$|^title, content$/i, 'obj'],
  [/^section header$/i, 'secHead'],
  [/^two content$/i, 'twoObj'],
  [/^comparison$/i, 'twoTxTwoObj'],
  [/^title only$/i, 'titleOnly'],
  [/^blank$/i, 'blank'],
  [/^content with caption$/i, 'objTx'],
  [/^picture with caption$/i, 'picTx'],
];

/** Well-known ODF shape types drawn by the matching preset when a file leaves out their outline. */
const ODF_PRESET: Record<string, string> = {
  rectangle: 'rect',
  rect: 'rect',
  'round-rectangle': 'roundRect',
  ellipse: 'ellipse',
  circle: 'ellipse',
  'isosceles-triangle': 'triangle',
  'right-triangle': 'rtTriangle',
  parallelogram: 'parallelogram',
  trapezoid: 'trapezoid',
  diamond: 'diamond',
  pentagon: 'pentagon',
  hexagon: 'hexagon',
  octagon: 'octagon',
  cross: 'plus',
  star4: 'star4',
  star5: 'star5',
  star6: 'star6',
  star8: 'star8',
  star12: 'star12',
  star24: 'star24',
  'right-arrow': 'rightArrow',
  'left-arrow': 'leftArrow',
  'up-arrow': 'upArrow',
  'down-arrow': 'downArrow',
  'left-right-arrow': 'leftRightArrow',
  'up-down-arrow': 'upDownArrow',
  'notched-right-arrow': 'notchedRightArrow',
  chevron: 'chevron',
  pentagon_right: 'homePlate',
  can: 'can',
  cube: 'cube',
  heart: 'heart',
  smiley: 'smileyFace',
  sun: 'sun',
  moon: 'moon',
  lightning: 'lightningBolt',
  cloud: 'cloud',
  frame: 'frame',
  ring: 'donut',
  'block-arc': 'blockArc',
  'flowchart-process': 'flowChartProcess',
  'flowchart-alternate-process': 'flowChartAlternateProcess',
  'flowchart-decision': 'flowChartDecision',
  'flowchart-data': 'flowChartInputOutput',
  'flowchart-predefined-process': 'flowChartPredefinedProcess',
  'flowchart-document': 'flowChartDocument',
  'flowchart-terminator': 'flowChartTerminator',
  'flowchart-preparation': 'flowChartPreparation',
  'flowchart-manual-input': 'flowChartManualInput',
  'flowchart-manual-operation': 'flowChartManualOperation',
  'flowchart-connector': 'flowChartConnector',
  'flowchart-off-page-connector': 'flowChartOffpageConnector',
  'flowchart-merge': 'flowChartMerge',
  'flowchart-extract': 'flowChartExtract',
  'flowchart-delay': 'flowChartDelay',
  'flowchart-sort': 'flowChartSort',
  'flowchart-or': 'flowChartOr',
  'flowchart-magnetic-disk': 'flowChartMagneticDisk',
  'rectangular-callout': 'wedgeRectCallout',
  'round-rectangular-callout': 'wedgeRoundRectCallout',
  'round-callout': 'wedgeEllipseCallout',
  'cloud-callout': 'cloudCallout',
  'left-bracket': 'leftBracket',
  'right-bracket': 'rightBracket',
  'left-brace': 'leftBrace',
  'right-brace': 'rightBrace',
  'bracket-pair': 'bracketPair',
  'brace-pair': 'bracePair',
  'horizontal-scroll': 'rect',
  'vertical-scroll': 'rect',
};

/* ================================================================= reader */

interface ShapeEnv {
  scope: 'content' | 'styles';
  /** Placeholders on a master page carry layout defaults rather than content. */
  master: boolean;
  /** xml:id / draw:id of shapes and paragraphs → element id and paragraph index (for animations). */
  ids: Map<string, { el: string; para?: number }>;
  /** A click action inherited from an enclosing draw:a. */
  link?: string;
  /** The page has a subtitle placeholder, so its title is a centred title. */
  titleSlide?: boolean;
  footerText?: string;
}

interface TextEnv {
  scope: 'content' | 'styles';
  /** Text defaults of the shape (presentation/graphic style and text style). */
  text: Props;
  para: Props;
  list?: Element;
  /** Outline placeholders: the master's style for deeper levels ("…-outline2" and on). */
  levelStyle?: (level: number) => Resolved | undefined;
  /** Paragraph ids → paragraph index. */
  paraIds?: Map<string, number>;
  footerText?: string;
}

type RawRun = Omit<Run, 'text'> & { text: string };

class OdpReader {
  files: Record<string, Uint8Array> = {};
  flat: Document | null = null;
  readonly styles = new Styles();
  theme: Theme;
  readonly warnings = new Set<string>();
  generator = '';
  pageIndex = new Map<string, number>();
  private pictures = new Map<string, { src: string; mime: string; bytes: Uint8Array } | null>();
  /** Placeholder and text defaults used to trim imported formatting. */
  pres!: Presentation;

  constructor() {
    this.theme = createPresentation('affice').theme;
  }

  parseXml(text: string): Document | null {
    try {
      const d = new DOMParser().parseFromString(text.replace(/^﻿/, ''), 'application/xml');
      if (!d.documentElement || d.getElementsByTagName('parsererror').length) return null;
      return d;
    } catch {
      return null;
    }
  }

  part(path: string): Document | null {
    const p = path.replace(/^\.\//, '');
    const bytes = this.files[p];
    return bytes ? this.parseXml(strFromU8(bytes)) : null;
  }

  /* ------------------------------------------------------------ colours */

  /** A model colour from an ODF colour and theme reference (which wins when it gives the same colour). */
  color(hex: string | undefined, complex: string | undefined, alpha = 1): string | undefined {
    const h = normHex(hex);
    if (complex) {
      const [type, ...mods] = complex.split(';');
      const name = THEME_IN[type];
      if (name) {
        let lumMod = 1;
        let lumOff = 0;
        let ok = true;
        for (const m of mods) {
          const [t, v] = m.split(':');
          const n = Number(v) / 10000;
          if (t === 'lummod') lumMod = n;
          else if (t === 'lumoff') lumOff = n;
          else ok = false;
        }
        const lum = lumOff > 0 ? Math.round(lumOff * 100) : lumMod < 1 ? -Math.round((1 - lumMod) * 100) : 0;
        if (ok) {
          const ref = colorRef(name, lum, alpha);
          if (!h || closeHex(resolveHex(colorRef(name, lum), this.theme).hex, h)) return ref;
        }
      }
    }
    if (!h) return undefined;
    return withAlpha(h, alpha);
  }

  /* -------------------------------------------------------------- media */

  /** A picture from the package (or inline base64) as a data URL. */
  picture(href: string | undefined, binary?: string): { src: string; mime: string; bytes: Uint8Array } | null {
    const key = href ?? `binary:${binary?.length}:${binary?.slice(0, 64)}`;
    if (this.pictures.has(key)) return this.pictures.get(key)!;
    let bytes: Uint8Array | undefined;
    if (binary) {
      try {
        const clean = binary.replace(/\s+/g, '');
        const bin = atob(clean);
        bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      } catch {
        bytes = undefined;
      }
    } else if (href && !/^[a-z]+:/i.test(href)) bytes = this.files[decodeURIComponent(href.replace(/^\.\//, ''))] ?? this.files[href.replace(/^\.\//, '')];
    let out: { src: string; mime: string; bytes: Uint8Array } | null = null;
    if (bytes) {
      const mime = sniffMime(bytes, href);
      if (mime) out = { src: `data:${mime};base64,${bytesToBase64(bytes)}`, mime, bytes };
      else this.warnings.add('Some pictures use formats that can’t be shown here (such as EMF, WMF or SVM) and were left out.');
    } else if (href && /^https?:/i.test(href)) out = { src: href, mime: 'image/png', bytes: new Uint8Array() };
    this.pictures.set(key, out);
    return out;
  }

  /* -------------------------------------------------------------- fills */

  fill(g: Props | undefined, shapeDefault: boolean, size?: { w: number; h: number }): Fill | undefined {
    if (!g) return shapeDefault ? { type: 'solid', color: '#729fcf' } : undefined;
    const mode = g.get('draw:fill') ?? (shapeDefault ? 'solid' : 'none');
    const alpha = pctOf(g.get('draw:opacity')) ?? 1;
    const opacityGrad = g.get('draw:opacity-name') ? this.styles.named.get(`opacity|${g.get('draw:opacity-name')}`) : undefined;
    switch (mode) {
      case 'none':
        return { type: 'none' };
      case 'solid': {
        if (opacityGrad) return this.gradient(undefined, opacityGrad, this.color(g.get('draw:fill-color') ?? '#729fcf', g.get('x:fill-complex')) ?? '#729fcf');
        return { type: 'solid', color: this.color(g.get('draw:fill-color') ?? '#729fcf', g.get('x:fill-complex'), alpha) ?? '#729fcf' };
      }
      case 'gradient': {
        const def = this.styles.named.get(`gradient|${g.get('draw:fill-gradient-name')}`);
        const f = this.gradient(def, opacityGrad, undefined, alpha);
        return f ?? { type: 'solid', color: normHex(g.get('draw:fill-color')) ?? '#729fcf' };
      }
      case 'bitmap': {
        const def = this.styles.named.get(`fill-image|${g.get('draw:fill-image-name')}`);
        const bin = def ? kid(def, NS.office, 'binary-data')?.textContent ?? undefined : undefined;
        const pic = def ? this.picture(at(def, NS.xlink, 'href'), bin) : null;
        if (!pic) return { type: 'none' };
        const repeat = g.get('style:repeat');
        return { type: 'image', src: pic.src, mode: repeat === 'repeat' ? 'tile' : 'stretch', ...(alpha < 0.999 ? { alpha } : {}) };
      }
      case 'hatch': {
        const def = this.styles.named.get(`hatch|${g.get('draw:fill-hatch-name')}`);
        const bg = g.get('draw:fill-hatch-solid') === 'true' ? (normHex(g.get('draw:fill-color')) ?? '#ffffff') : undefined;
        // hatching is drawn into a picture the size of the shape
        if (def && size && size.w > 0 && size.h > 0) return { type: 'image', src: this.hatchPicture(def, bg, size.w, size.h), mode: 'stretch', ...(alpha < 0.999 ? { alpha } : {}) };
        if (bg) return { type: 'solid', color: this.color(g.get('draw:fill-color'), g.get('x:fill-complex'), alpha) ?? bg };
        const c = normHex(at(def, NS.draw, 'color')) ?? '#000000';
        return { type: 'solid', color: withAlpha(c, 0.3 * alpha) };
      }
      default:
        return { type: 'none' };
    }
  }

  /** An SVG picture of a hatch (single, crossed or triple lines) over an optional background colour. */
  hatchPicture(def: Element, bg: string | undefined, w: number, h: number): string {
    const color = normHex(at(def, NS.draw, 'color')) ?? '#000000';
    const d = Math.max(2, lengthPx(at(def, NS.draw, 'distance')) ?? 4);
    const rot = this.angleDeg(at(def, NS.draw, 'rotation'));
    const style = at(def, NS.draw, 'style') ?? 'single';
    const angles = style === 'triple' ? [0, 90, 45] : style === 'double' ? [0, 90] : [0];
    const W = Math.round(w * 100) / 100;
    const H = Math.round(h * 100) / 100;
    let defs = '';
    let body = bg ? `<rect width="${W}" height="${H}" fill="${bg}"/>` : '';
    angles.forEach((a, i) => {
      const gap = a === 45 ? d * Math.SQRT2 : d;
      defs += `<pattern id="h${i}" width="${gap}" height="${gap}" patternUnits="userSpaceOnUse" patternTransform="rotate(${-(rot + a)})"><line x1="0" y1="0.5" x2="${gap}" y2="0.5" stroke="${color}" stroke-width="1"/></pattern>`;
      body += `<rect width="${W}" height="${H}" fill="url(#h${i})"/>`;
    });
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}"><defs>${defs}</defs>${body}</svg>`;
    return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
  }

  /** A gradient fill from draw:gradient (colours) and draw:opacity (transparency) definitions. */
  gradient(def: Element | undefined, opacity: Element | undefined, solid?: string, alpha = 1): Fill | undefined {
    const src = def ?? opacity;
    if (!src) return undefined;
    const style = at(src, NS.draw, 'style') ?? 'linear';
    const border = pctOf(at(src, NS.draw, 'border')) ?? 0;
    const angle = this.angleDeg(at(src, NS.draw, 'angle'));
    // colour stops (0 = start) with intensities applied
    let stops: { pos: number; hex: string }[] = [];
    const refs = new Map<number, string>();
    if (def) {
      const lo = kidsOf(def, NS.loext, 'gradient-stop');
      for (const s of lo) {
        const ref = at(s, NS.affice, 'color');
        if (ref && parseColorRef(ref)) refs.set(Number(at(s, NS.svg, 'offset') ?? 0), ref);
      }
      if (lo.length) stops = lo.map((s) => ({ pos: Number(at(s, NS.svg, 'offset') ?? 0), hex: normHex(at(s, NS.loext, 'color-value')) ?? '#000000' }));
      else {
        const dim = (hex: string, k: number) => (k >= 0.999 ? hex : mixHex('#000000', hex, k));
        stops = [
          { pos: 0, hex: dim(normHex(at(def, NS.draw, 'start-color')) ?? '#000000', pctOf(at(def, NS.draw, 'start-intensity')) ?? 1) },
          { pos: 1, hex: dim(normHex(at(def, NS.draw, 'end-color')) ?? '#ffffff', pctOf(at(def, NS.draw, 'end-intensity')) ?? 1) },
        ];
      }
    } else stops = [{ pos: 0, hex: normHex(solid) ?? '#000000' }, { pos: 1, hex: normHex(solid) ?? '#000000' }];
    // opacity stops (0 = start)
    let ops: { pos: number; a: number }[] = [];
    if (opacity) {
      const lo = kidsOf(opacity, NS.loext, 'opacity-stop');
      if (lo.length) ops = lo.map((s) => ({ pos: Number(at(s, NS.svg, 'offset') ?? 0), a: Number(at(s, NS.svg, 'stop-opacity') ?? 1) }));
      else ops = [{ pos: 0, a: pctOf(at(opacity, NS.draw, 'start')) ?? 1 }, { pos: 1, a: pctOf(at(opacity, NS.draw, 'end')) ?? 1 }];
    }
    const alphaAt = (pos: number) => {
      if (!ops.length) return alpha;
      if (pos <= ops[0].pos) return ops[0].a * alpha;
      for (let i = 1; i < ops.length; i++)
        if (pos <= ops[i].pos) {
          const t = (pos - ops[i - 1].pos) / (ops[i].pos - ops[i - 1].pos || 1);
          return (ops[i - 1].a + (ops[i].a - ops[i - 1].a) * t) * alpha;
        }
      return ops[ops.length - 1].a * alpha;
    };
    const positions = [...new Set([...stops.map((s) => s.pos), ...ops.map((o) => o.pos)])].sort((a, b) => a - b);
    const colorAt = (pos: number) => {
      if (pos <= stops[0].pos) return stops[0].hex;
      for (let i = 1; i < stops.length; i++)
        if (pos <= stops[i].pos) return mixHex(stops[i - 1].hex, stops[i].hex, (pos - stops[i - 1].pos) / (stops[i].pos - stops[i - 1].pos || 1));
      return stops[stops.length - 1].hex;
    };
    // a stop's colour: our theme reference when we wrote one (with the stop's transparency), else the plain colour
    const stopColor = (p: number) => {
      const ref = refs.get(p);
      const r = ref ? parseColorRef(ref) : null;
      if (r) return colorRef(r.name, r.lum, Math.round(alphaAt(p) * 100) / 100);
      return withAlpha(colorAt(p), alphaAt(p));
    };
    let out: GradientStop[];
    const radial = style === 'radial' || style === 'ellipsoid' || style === 'square' || style === 'rectangular';
    if (radial) {
      // ODF runs from the edge (0) to the centre (1) over a circle through the corners, less the border
      const k = (1 - border) * Math.SQRT2;
      const raw = positions.map((p) => ({ pos: (1 - p) * k, color: stopColor(p) })).sort((a, b) => a.pos - b.pos);
      out = clampStops(raw);
    } else if (style === 'axial') {
      const half = positions.map((p) => ({ pos: border / 2 + (p * (1 - border)) / 2, color: stopColor(p) }));
      out = [...half, ...half.map((s) => ({ pos: 1 - s.pos, color: s.color })).reverse()];
      if (border > 0) out = [{ pos: 0, color: half[0].color }, ...out, { pos: 1, color: half[0].color }];
    } else {
      out = positions.map((p) => ({ pos: border + p * (1 - border), color: stopColor(p) }));
      if (border > 0) out.unshift({ pos: 0, color: out[0].color });
    }
    out = out.map((s) => ({ pos: Math.round(s.pos * 10000) / 10000, color: s.color }));
    return radial ? { type: 'gradient', stops: out, angle: 0, radial: true } : { type: 'gradient', stops: out, angle: (((90 - angle) % 360) + 360) % 360 };
  }

  angleDeg(v: string | undefined): number {
    if (!v) return 0;
    const m = /^\s*([-+]?[\d.]+)\s*(deg|rad|grad)?\s*$/.exec(v);
    if (!m) return 0;
    const n = Number(m[1]);
    if (m[2] === 'rad') return (n * 180) / Math.PI;
    if (m[2] === 'grad') return n * 0.9;
    if (m[2] === 'deg') return n;
    // without a unit, OpenOffice and LibreOffice before 7.6 wrote tenths of a degree
    const lo = /(?:OpenOffice|LibreOffice)[^/]*\/(\d+)\.(\d+)/.exec(this.generator);
    const old = !this.generator || /OpenOffice|StarOffice/.test(this.generator) || (lo ? Number(lo[1]) * 100 + Number(lo[2]) < 706 : false);
    return old || n > 360 ? n / 10 : n;
  }

  line(g: Props | undefined, shapeDefault: boolean): Line | null {
    if (!g) return shapeDefault ? { color: '#3465a4', width: 1 } : null;
    const mode = g.get('draw:stroke') ?? (shapeDefault ? 'solid' : 'none');
    if (mode === 'none') return null;
    const width = lengthPx(g.get('svg:stroke-width')) ?? 0;
    const alpha = pctOf(g.get('svg:stroke-opacity')) ?? 1;
    const out: Line = { color: this.color(g.get('svg:stroke-color') ?? '#000000', g.get('x:stroke-complex'), alpha) ?? '#000000', width: width > 0 ? r2(width) : 1 };
    if (mode === 'dash') {
      const def = this.styles.named.get(`stroke-dash|${g.get('draw:stroke-dash')}`);
      out.dash = dashOf(def, width || 1);
    }
    const cap = g.get('svg:stroke-linecap');
    if (cap === 'round' || cap === 'square') out.cap = cap;
    const head = this.marker(g.get('draw:marker-start'));
    const tail = this.marker(g.get('draw:marker-end'));
    if (head) out.head = head;
    if (tail) out.tail = tail;
    return out;
  }

  marker(name: string | undefined): ArrowType | undefined {
    if (!name) return undefined;
    const def = this.styles.named.get(`marker|${name}`);
    const label = `${at(def, NS.draw, 'display-name') ?? ''} ${name}`.toLowerCase().replace(/_20_/g, ' ');
    const direct = (['triangle', 'arrow', 'stealth', 'oval', 'diamond'] as ArrowType[]).find((t) => (at(def, NS.draw, 'display-name') ?? '').toLowerCase() === t);
    if (direct) return direct;
    if (/open|line arrow|line short/.test(label)) return 'arrow';
    if (/stealth|concave/.test(label)) return 'stealth';
    if (/oval|circle/.test(label)) return 'oval';
    if (/diamond|square 45|square_45/.test(label)) return 'diamond';
    return 'triangle';
  }

  shadow(g: Props | undefined): Shadow | undefined {
    if (!g || g.get('draw:shadow') !== 'visible') return undefined;
    const dx = lengthPx(g.get('draw:shadow-offset-x')) ?? 0;
    const dy = lengthPx(g.get('draw:shadow-offset-y')) ?? 0;
    const alpha = pctOf(g.get('draw:shadow-opacity')) ?? 1;
    return {
      color: this.color(g.get('draw:shadow-color') ?? '#808080', undefined, alpha) ?? '#808080',
      blur: r2((lengthPx(g.get('loext:shadow-blur')) ?? 0) * 2),
      dist: r2(Math.hypot(dx, dy)),
      angle: Math.round(((Math.atan2(dy, dx) * 180) / Math.PI + 360) % 360),
    };
  }

  /** Text frame settings (anchor, insets, wrapping, fitting, direction, columns) from graphic properties. */
  frameProps(g: Props | undefined, p: Props | undefined, frame = false): Partial<TextBody> {
    const out: Partial<TextBody> = {};
    if (!g) return out;
    const all = lengthPx(g.get('fo:padding'));
    const inset: [number, number, number, number] = [
      r2(lengthPx(g.get('fo:padding-left')) ?? all ?? DEFAULT_INSET[0]),
      r2(lengthPx(g.get('fo:padding-top')) ?? all ?? DEFAULT_INSET[1]),
      r2(lengthPx(g.get('fo:padding-right')) ?? all ?? DEFAULT_INSET[2]),
      r2(lengthPx(g.get('fo:padding-bottom')) ?? all ?? DEFAULT_INSET[3]),
    ];
    if (inset.some((v, i) => Math.abs(v - DEFAULT_INSET[i]) > 0.05)) out.inset = inset;
    const va = g.get('draw:textarea-vertical-align');
    if (va === 'middle') out.anchor = 'm';
    else if (va === 'bottom') out.anchor = 'b';
    // text frames wrap unless they grow sideways; shapes follow their word-wrap setting
    if (frame ? g.get('draw:auto-grow-width') === 'true' : g.get('fo:wrap-option') === 'no-wrap') out.wrap = false;
    const fit = g.get('draw:fit-to-size');
    if (g.get('style:shrink-to-fit') === 'true' || fit === 'shrink-to-fit' || fit === 'true' || fit === 'all') out.autofit = 'shrink';
    else if (g.get('draw:auto-grow-height') === 'true') out.autofit = 'resize';
    const wm = p?.get('style:writing-mode') ?? g.get('style:writing-mode');
    if (wm === 'tb-rl' || wm === 'tb') out.vert = 'vert';
    else if (wm === 'bt-lr') out.vert = 'vert270';
    const cols = Number(g.get('x:columns') ?? 1);
    if (cols > 1) out.columns = cols;
    return out;
  }

  /* --------------------------------------------------------------- text */

  runOf(p: Props): Omit<Run, 'text'> {
    const r: Omit<Run, 'text'> = {};
    const font = unquoteFont(p.get('fo:font-family')) ?? (p.get('style:font-name') ? this.styles.fonts.get(p.get('style:font-name')!) ?? p.get('style:font-name') : undefined);
    if (font) r.font = font;
    const fs = p.get('fo:font-size');
    const size = fs?.trim().endsWith('%') ? r2(18 * (pctOf(fs) ?? 1)) : ptOf(fs);
    if (size) r.size = size;
    const alpha = pctOf(p.get('loext:opacity')) ?? 1;
    const color = p.get('style:use-window-font-color') === 'true' && !p.get('fo:color') ? '@tx1' : this.color(p.get('fo:color'), p.get('x:char-complex'), alpha);
    if (color) r.color = color;
    // ODF text is plain unless a style says otherwise (our runs would inherit a bold title default)
    const weight = p.get('fo:font-weight');
    r.b = !!weight && (weight === 'bold' || Number(weight) >= 600);
    const style = p.get('fo:font-style');
    r.i = style === 'italic' || style === 'oblique';
    const u = p.get('style:text-underline-style');
    r.u = !!u && u !== 'none';
    const s = p.get('style:text-line-through-style');
    r.s = !!s && s !== 'none';
    const pos = p.get('style:text-position');
    if (pos) {
      const first = pos.trim().split(/\s+/)[0];
      const v = first === 'super' ? 1 : first === 'sub' ? -1 : (pctOf(first) ?? 0);
      if (v > 0) r.sup = true;
      else if (v < 0) r.sub = true;
    }
    if (p.get('fo:text-transform') === 'uppercase' || p.get('fo:font-variant') === 'small-caps') r.caps = true;
    const ls = p.get('fo:letter-spacing');
    if (ls && ls !== 'normal') {
      const v = ptOf(ls);
      if (v) r.spacing = v;
    }
    const bg = p.get('fo:background-color');
    if (bg && bg !== 'transparent') {
      const hl = normHex(bg);
      if (hl) r.hl = hl;
    }
    return r;
  }

  /** Paragraphs of a text container (draw:text-box, a shape, a table cell, a notes frame). */
  paragraphs(container: Element, env: TextEnv): Para[] {
    const paras: Para[] = [];
    const listProps = (list: Element | undefined, level: number): { bullet?: Bullet; marL?: number; indent?: number; bulletColorFont?: { color?: string; font?: string } } => {
      if (!list) return {};
      const lv = kids(list).find((k) => Number(at(k, NS.text, 'level') ?? 1) === level + 1);
      if (!lv) return {};
      const props = kid(lv, NS.style, 'list-level-properties');
      const mode = at(props, NS.text, 'list-level-position-and-space-mode');
      let marL: number;
      let indent: number;
      if (mode === 'label-alignment') {
        const la = kid(props, NS.style, 'list-level-label-alignment');
        marL = lengthPx(at(la, NS.fo, 'margin-left')) ?? 0;
        indent = lengthPx(at(la, NS.fo, 'text-indent')) ?? 0;
      } else {
        const sb = lengthPx(at(props, NS.text, 'space-before')) ?? 0;
        const mlw = lengthPx(at(props, NS.text, 'min-label-width')) ?? 0;
        marL = sb + mlw;
        indent = -mlw;
      }
      const tp = kid(lv, NS.style, 'text-properties');
      const color = tp && at(tp, NS.fo, 'color') ? normHex(at(tp, NS.fo, 'color')) : undefined;
      const font = unquoteFont(at(tp, NS.fo, 'font-family') ?? (at(tp, NS.style, 'font-name') ? this.styles.fonts.get(at(tp, NS.style, 'font-name')!) : undefined));
      if (lv.localName === 'list-level-style-bullet') {
        const char = at(lv, NS.text, 'bullet-char') ?? '•';
        return { bullet: { type: 'char', char }, marL, indent, bulletColorFont: { color, font } };
      }
      if (lv.localName === 'list-level-style-number') {
        const fmt = at(lv, NS.style, 'num-format') ?? '';
        if (!fmt) return { bullet: { type: 'none' }, marL, indent };
        const suffix = at(lv, NS.style, 'num-suffix') ?? '';
        const key = `${fmt}${suffix.includes(')') ? ')' : '.'}`;
        const style = NUM_IN[key] ?? (fmt === 'i' ? 'romanLcPeriod' : fmt === 'I' ? 'romanUcPeriod' : fmt === 'a' ? 'alphaLcPeriod' : fmt === 'A' ? 'alphaUcPeriod' : 'arabicPeriod');
        const start = Number(at(lv, NS.text, 'start-value') ?? 1);
        return { bullet: { type: 'num', style, ...(start !== 1 ? { start } : {}) }, marL, indent, bulletColorFont: { color } };
      }
      // picture bullets show as dots
      return { bullet: { type: 'char', char: '•' }, marL, indent };
    };

    const onePara = (p: Element, list: { el?: Element; level: number; header: boolean; start?: number } | null) => {
      const ps = this.styles.resolve(at(p, NS.text, 'style-name'), 'paragraph', env.scope);
      const pp = new Map(env.para);
      const tp = new Map(env.text);
      const lvl = list && list.level > 0 ? env.levelStyle?.(list.level) : undefined;
      if (lvl) {
        mergeProps(pp, lvl.c.paragraph ?? new Map());
        mergeProps(tp, lvl.c.text ?? new Map());
      }
      mergeProps(pp, ps.c.paragraph ?? new Map());
      mergeProps(tp, ps.c.text ?? new Map());
      const para: Para = { runs: [] };
      const align = pp.get('fo:text-align');
      if (align === 'center') para.align = 'center';
      else if (align === 'end' || align === 'right') para.align = 'right';
      else if (align === 'justify') para.align = 'justify';
      else if (align === 'start' || align === 'left') para.align = 'left';
      const size = ptOf(tp.get('fo:font-size')) ?? 18;
      const lh = pp.get('fo:line-height');
      if (lh && lh !== 'normal') {
        const pct = pctOf(lh);
        if (pct !== undefined) para.lineSpacing = r2(pct);
        else {
          const v = ptOf(lh);
          if (v) para.lineSpacing = r2(v / (size * 1.2));
        }
      } else if (pp.get('style:line-spacing')) {
        const v = ptOf(pp.get('style:line-spacing'));
        if (v !== undefined) para.lineSpacing = r2(1 + v / (size * 1.2));
      } else if (pp.get('style:line-height-at-least')) {
        const v = ptOf(pp.get('style:line-height-at-least'));
        if (v) para.lineSpacing = r2(Math.max(1, v / (size * 1.2)));
      }
      const before = ptOf(pp.get('fo:margin-top'));
      const after = ptOf(pp.get('fo:margin-bottom'));
      if (before !== undefined) para.spaceBefore = before;
      if (after !== undefined) para.spaceAfter = after;
      const ml = lengthPx(pp.get('fo:margin-left')) ?? 0;
      const ti = lengthPx(pp.get('fo:text-indent')) ?? 0;
      if (list) {
        para.level = list.level;
        const lp = list.header ? { bullet: { type: 'none' } as Bullet } : listProps(list.el, list.level);
        if (lp.bullet) para.bullet = lp.bullet;
        if (lp.bullet?.type === 'num' && list.start !== undefined) para.bullet = { ...lp.bullet, start: list.start };
        para.marL = r2((lp.marL ?? 0) + ml);
        para.indent = r2((lp.indent ?? 0) + ti);
        const bc = lp.bulletColorFont;
        if (bc && para.bullet?.type === 'char') {
          if (bc.color) para.bullet = { ...para.bullet, color: bc.color };
          if (bc.font && bc.font !== 'Arial' && bc.font !== 'OpenSymbol' && bc.font !== 'StarSymbol') para.bullet = { ...para.bullet, font: bc.font };
        }
      } else {
        para.bullet = { type: 'none' };
        para.marL = r2(ml);
        para.indent = r2(ti);
      }
      // runs
      const runs: RawRun[] = [];
      let space = true;
      const push = (text: string, props: Props, extra: Partial<Run> = {}) => {
        if (!text && !extra.field) return;
        const r = { ...this.runOf(props), ...extra, text };
        const last = runs[runs.length - 1];
        if (last && !extra.field && !last.field && sameRun(last, r)) last.text += text;
        else runs.push(r);
      };
      const walk = (el: Element, props: Props, link?: string) => {
        for (let n = el.firstChild; n; n = n.nextSibling) {
          if (n.nodeType === 3 || n.nodeType === 4) {
            let t = '';
            for (const ch of n.nodeValue ?? '') {
              if (ch === ' ' || ch === '\n' || ch === '\t' || ch === '\r') {
                if (!space) t += ' ';
                space = true;
              } else {
                t += ch;
                space = false;
              }
            }
            push(t, props, link ? { link } : {});
            continue;
          }
          if (n.nodeType !== 1) continue;
          const k = n as Element;
          if (k.namespaceURI === NS.text) {
            switch (k.localName) {
              case 'span': {
                const sp = this.styles.resolve(at(k, NS.text, 'style-name'), 'text', env.scope);
                const merged = new Map(props);
                mergeProps(merged, sp.c.text ?? new Map());
                walk(k, merged, link);
                break;
              }
              case 'a': {
                const href = at(k, NS.xlink, 'href');
                walk(k, props, href ? this.linkIn(href) : link);
                break;
              }
              case 's':
                push(' '.repeat(Math.max(1, Number(at(k, NS.text, 'c') ?? 1))), props, link ? { link } : {});
                space = true;
                break;
              case 'tab':
                push('\t', props, link ? { link } : {});
                space = false;
                break;
              case 'line-break':
                push('\n', props, link ? { link } : {});
                space = true;
                break;
              case 'page-number':
                push(k.textContent || '1', props, { field: 'slidenum' });
                space = false;
                break;
              case 'date':
              case 'time':
                push(k.textContent || new Date().toLocaleDateString(), props, { field: 'date' });
                space = false;
                break;
              case 'note':
              case 'annotation':
              case 'bookmark':
              case 'bookmark-start':
              case 'bookmark-end':
              case 'soft-page-break':
              case 'change':
              case 'change-start':
              case 'change-end':
                break;
              default:
                walk(k, props, link);
            }
          } else if (k.namespaceURI === NS.presentation) {
            if (k.localName === 'date-time') push(new Date().toLocaleDateString(), props, { field: 'date' });
            else if (k.localName === 'footer' && env.footerText) push(env.footerText, props);
            space = false;
          } else if (k.namespaceURI === NS.office && k.localName === 'annotation') {
            // comments are not kept
          } else if (k.namespaceURI !== NS.draw) walk(k, props, link);
        }
      };
      walk(p, tp);
      para.runs = runs;
      const endProps = this.runOf(tp);
      if (!runs.length) para.endRun = endProps;
      const id = at(p, NS.xml, 'id') ?? at(p, NS.text, 'id');
      if (id && env.paraIds) env.paraIds.set(id, paras.length);
      paras.push(para);
    };

    const visitList = (listEl: Element, level: number, inherited: Element | undefined) => {
      const style = this.styles.list(at(listEl, NS.text, 'style-name'), env.scope) ?? inherited;
      for (const item of kids(listEl)) {
        if (item.namespaceURI !== NS.text || (item.localName !== 'list-item' && item.localName !== 'list-header')) continue;
        const header = item.localName === 'list-header';
        const start = at(item, NS.text, 'start-value');
        let first = true;
        for (const c of kids(item)) {
          if (c.namespaceURI !== NS.text) continue;
          if (c.localName === 'p' || c.localName === 'h') {
            onePara(c, { el: style, level, header: header || !first, start: first && start ? Number(start) : undefined });
            first = false;
          } else if (c.localName === 'list') visitList(c, level + 1, style);
        }
      }
    };

    for (const c of kids(container)) {
      if (c.namespaceURI !== NS.text) continue;
      if (c.localName === 'p' || c.localName === 'h') onePara(c, null);
      else if (c.localName === 'list') visitList(c, 0, env.list);
      else if (c.localName === 'section') paras.push(...this.paragraphs(c, env));
    }
    return paras;
  }

  linkIn(href: string): string {
    if (href.startsWith('#')) {
      const name = decodeURIComponent(href.slice(1));
      const idx = this.pageIndex.get(name);
      if (idx !== undefined) return `#slide:${idx + 1}`;
      const m = /^(?:page|slide)\s*(\d+)$/i.exec(name);
      if (m) return `#slide:${m[1]}`;
    }
    return href;
  }

  /* ---------------------------------------------------------- geometry */

  /** Frame of a shape: position, size and rotation (draw:transform rotates about the origin, then moves). */
  frame(el: Element): { x: number; y: number; w: number; h: number; rot: number } {
    const w = Math.max(0, lengthPx(at(el, NS.svg, 'width')) ?? 0);
    const h = Math.max(0, lengthPx(at(el, NS.svg, 'height')) ?? 0);
    const x0 = lengthPx(at(el, NS.svg, 'x')) ?? 0;
    const y0 = lengthPx(at(el, NS.svg, 'y')) ?? 0;
    const tr = at(el, NS.draw, 'transform');
    if (!tr) return { x: x0, y: y0, w, h, rot: 0 };
    const m = parseTransform(tr);
    const lx = x0 + w / 2;
    const ly = y0 + h / 2;
    const cx = m[0] * lx + m[2] * ly + m[4];
    const cy = m[1] * lx + m[3] * ly + m[5];
    let rot = (Math.atan2(m[1], m[0]) * 180) / Math.PI;
    rot = Math.round(rot * 100) / 100;
    if (Math.abs(rot) < 0.01) rot = 0;
    return { x: r2(cx - w / 2), y: r2(cy - h / 2), w: r2(w), h: r2(h), rot: rot < 0 ? rot + 360 : rot };
  }

  /* ------------------------------------------------------------- shapes */

  shapeStyle(el: Element, scope: 'content' | 'styles'): Resolved {
    const pres = at(el, NS.presentation, 'style-name');
    const draw = at(el, NS.draw, 'style-name');
    let r = pres ? this.styles.resolve(pres, 'presentation', scope) : this.styles.resolve(draw, 'graphic', scope);
    if (pres && !['fo:padding', 'fo:padding-left', 'fo:padding-top', 'fo:padding-right', 'fo:padding-bottom'].some((k) => r.c.graphic?.has(k))) {
      // presentation styles don't inherit the drawing defaults: text sits right at the edge
      r = cloneResolved(r);
      r.c.graphic = r.c.graphic ?? new Map();
      r.c.graphic.set('fo:padding', '0cm');
    }
    const ts = at(el, NS.draw, 'text-style-name');
    if (!ts) return r;
    // the shape's text style (paragraph defaults) comes after the graphic style
    const t = this.styles.resolve(ts, 'paragraph', scope);
    const out = cloneResolved(r);
    for (const kind of ['paragraph', 'text']) {
      if (!t.c[kind]) continue;
      out.c[kind] = out.c[kind] ?? new Map();
      mergeProps(out.c[kind], t.c[kind]);
    }
    return out;
  }

  textBody(container: Element, style: Resolved, env: ShapeEnv, paraIds: Map<string, number>, frame = false, owner?: Element): TextBody {
    const tenv: TextEnv = { scope: env.scope, text: style.c.text ?? new Map(), para: style.c.paragraph ?? new Map(), list: style.list, paraIds, footerText: env.footerText, levelStyle: owner ? this.outlineLevels(owner, env.scope) : undefined };
    const paras = this.paragraphs(container, tenv);
    // an empty body keeps the style's formatting for text typed into it
    const body: TextBody = { paras: paras.length ? paras : [{ runs: [], endRun: this.runOf(tenv.text) }], ...this.frameProps(style.c.graphic, style.c.paragraph, frame) };
    return body;
  }

  /** For an outline placeholder, the master's styles for its deeper levels (found through its style's parents). */
  outlineLevels(el: Element, scope: 'content' | 'styles'): ((level: number) => Resolved | undefined) | undefined {
    if (at(el, NS.presentation, 'class') !== 'outline') return undefined;
    let name = at(el, NS.presentation, 'style-name');
    for (let i = 0; name && i < 10; i++) {
      if (/-outline1$/.test(name)) {
        const prefix = name.slice(0, -'-outline1'.length);
        return (level) => (this.styles.element(`${prefix}-outline${level + 1}`, 'presentation', 'styles') ? this.styles.resolve(`${prefix}-outline${level + 1}`, 'presentation', 'styles') : undefined);
      }
      name = at(this.styles.element(name, 'presentation', scope), NS.style, 'parent-style-name');
    }
    return undefined;
  }

  common(el: Element, base: { x: number; y: number; w: number; h: number; rot: number }, env: ShapeEnv): Pick<El, 'id' | 'x' | 'y' | 'w' | 'h'> & Partial<El> {
    const out: Pick<El, 'id' | 'x' | 'y' | 'w' | 'h'> & Partial<El> = { id: newId(), x: base.x, y: base.y, w: base.w, h: base.h };
    if (base.rot) out.rot = base.rot;
    const name = at(el, NS.draw, 'name');
    if (name) out.name = name;
    const desc = kid(el, NS.svg, 'desc')?.textContent || kid(el, NS.svg, 'title')?.textContent;
    if (desc) out.alt = desc;
    if (at(el, NS.draw, 'display') === 'none') out.hidden = true;
    const link = this.clickAction(el) ?? env.link;
    if (link) out.link = link;
    return out;
  }

  clickAction(el: Element): string | undefined {
    const listeners = kid(el, NS.office, 'event-listeners');
    for (const l of kids(listeners)) {
      if (l.localName !== 'event-listener') continue;
      const action = at(l, NS.presentation, 'action');
      const href = at(l, NS.xlink, 'href');
      switch (action) {
        case 'next-page':
          return '#next';
        case 'previous-page':
          return '#prev';
        case 'first-page':
          return '#first';
        case 'last-page':
          return '#last';
        case 'show':
          return href ? this.linkIn(href) : undefined;
      }
      if (href) return this.linkIn(href);
    }
    return undefined;
  }

  register(el: Element, id: string, env: ShapeEnv) {
    const xid = at(el, NS.xml, 'id') ?? at(el, NS.draw, 'id');
    if (xid) env.ids.set(xid, { el: id });
  }

  placeholder(el: Element, env: ShapeEnv): { ph?: PlaceholderType; phIdx?: number; empty: boolean } {
    const own = at(el, NS.affice, 'placeholder');
    const idx = at(el, NS.affice, 'placeholder-index');
    const cls = at(el, NS.presentation, 'class');
    let ph = own && PH_TYPES.has(own) ? (own as PlaceholderType) : cls ? PH_IN[cls] : undefined;
    if (!own && ph === 'title' && env.titleSlide) ph = 'ctrTitle';
    return { ph, phIdx: idx ? Number(idx) : undefined, empty: at(el, NS.presentation, 'placeholder') === 'true' };
  }

  shape(el: Element, env: ShapeEnv): El | null {
    if (el.namespaceURI === NS.draw) {
      switch (el.localName) {
        case 'custom-shape':
          return this.customShape(el, env);
        case 'rect':
        case 'ellipse':
        case 'circle':
        case 'caption':
          return this.basicShape(el, env);
        case 'line':
          return this.lineShape(el, env);
        case 'connector':
          return this.connector(el, env);
        case 'polyline':
        case 'polygon':
        case 'path':
          return this.pathShape(el, env);
        case 'regular-polygon':
          return this.regularPolygon(el, env);
        case 'frame':
          return this.frameShape(el, env);
        case 'g':
          return this.group(el, env);
        case 'a': {
          const href = at(el, NS.xlink, 'href');
          const inner = kids(el).map((k) => this.shape(k, { ...env, link: href ? this.linkIn(href) : env.link })).filter((x): x is El => !!x);
          return inner.length === 1 ? inner[0] : inner.length ? this.groupOf(inner, el) : null;
        }
        case 'page-thumbnail':
        case 'control':
        case 'measure':
          return null;
      }
    }
    if (el.namespaceURI === 'urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0') this.warnings.add('3D objects were left out.');
    return null;
  }

  withText(shape: ShapeEl, el: Element, style: Resolved, env: ShapeEnv, container: Element = el) {
    const hasText = kids(container).some((k) => k.namespaceURI === NS.text);
    if (!hasText && !shape.ph) return;
    const paraIds = new Map<string, number>();
    shape.text = this.textBody(container, style, env, paraIds, container !== el, el);
    for (const [pid, i] of paraIds) env.ids.set(pid, { el: shape.id, para: i });
  }

  customShape(el: Element, env: ShapeEnv): El | null {
    const f = this.frame(el);
    const style = this.shapeStyle(el, env.scope);
    const g = style.c.graphic;
    const geo = kid(el, NS.draw, 'enhanced-geometry');
    const type = at(geo, NS.draw, 'type') ?? 'non-primitive';
    const mods = (at(geo, NS.draw, 'modifiers') ?? '').trim().split(/\s+/).filter(Boolean).map(Number);
    const path = at(geo, NS.drawooo, 'enhanced-path') ?? at(geo, NS.draw, 'enhanced-path');
    const base = this.common(el, f, env);
    const shape: ShapeEl = { ...base, type: 'shape', geom: 'rect', fill: this.fill(g, true, f), line: this.line(g, true) };
    const sh = this.shadow(g);
    if (sh) shape.shadow = sh;
    if (at(geo, NS.draw, 'mirror-horizontal') === 'true') shape.flipH = true;
    if (at(geo, NS.draw, 'mirror-vertical') === 'true') shape.flipV = true;
    let textInset: [number, number, number, number] | undefined;
    const ooxml = type.startsWith('ooxml-') ? type.slice(6) : undefined;
    if (ooxml && isKnownPreset(ooxml)) {
      shape.geom = ooxml;
      const adj = adjFrom(ooxml, mods);
      if (adj) shape.adj = adj;
    } else if (type === 'rectangle' || type === 'rect' || (ODF_PRESET[type] === 'rect' && !path)) shape.geom = 'rect';
    else if (type === 'ellipse' || (type === 'circle' && !path)) shape.geom = 'ellipse';
    else if (!path && ODF_PRESET[type]) {
      shape.geom = ODF_PRESET[type];
      if (type === 'round-rectangle' && mods.length) shape.adj = { adj: Math.round((mods[0] / 21600) * 100000) };
    } else if (path) {
      const vb = (at(geo, NS.svg, 'viewBox') ?? '').trim().split(/[\s,]+/).map(Number);
      const equations = new Map<string, string>();
      for (const e of kidsOf(geo, NS.draw, 'equation')) equations.set(at(e, NS.draw, 'name') ?? '', at(e, NS.draw, 'formula') ?? '0');
      const ev = evaluateEnhancedGeometry({ viewBox: vb.length === 4 && vb.every(Number.isFinite) ? vb : undefined, path, equations, modifiers: mods, textAreas: at(geo, NS.draw, 'text-areas'), logicalText: !!ooxml }, f.w, f.h);
      if (ev.paths.length) {
        shape.geom = 'custom';
        shape.custom = ev.paths;
        textInset = ev.text;
      }
    }
    if (at(el, NS.affice, 'picture') === 'true' && shape.fill?.type === 'image') {
      // a picture cropped to a shape
      const img: ImageEl = { ...base, type: 'image', src: shape.fill.src, geom: shape.geom === 'custom' ? 'rect' : shape.geom };
      if (shape.line) img.line = shape.line;
      if (sh) img.shadow = sh;
      this.register(el, img.id, env);
      return img;
    }
    const ph = this.placeholder(el, env);
    if (ph.ph) {
      shape.ph = ph.ph;
      if (ph.phIdx) shape.phIdx = ph.phIdx;
    }
    this.withText(shape, el, style, env);
    if (shape.text && textInset) {
      const inset = shape.text.inset ?? DEFAULT_INSET;
      shape.text.inset = inset.map((v, i) => r2(v + Math.max(0, textInset![i]))) as [number, number, number, number];
    }
    this.register(el, shape.id, env);
    return shape;
  }

  basicShape(el: Element, env: ShapeEnv): El | null {
    let f = this.frame(el);
    if (el.localName === 'circle' && at(el, NS.svg, 'r')) {
      const r = lengthPx(at(el, NS.svg, 'r')) ?? 0;
      const cx = lengthPx(at(el, NS.svg, 'cx')) ?? 0;
      const cy = lengthPx(at(el, NS.svg, 'cy')) ?? 0;
      f = { x: cx - r, y: cy - r, w: 2 * r, h: 2 * r, rot: 0 };
    }
    const style = this.shapeStyle(el, env.scope);
    const g = style.c.graphic;
    const shape: ShapeEl = { ...this.common(el, f, env), type: 'shape', geom: 'rect', fill: this.fill(g, true, f), line: this.line(g, true) };
    const sh = this.shadow(g);
    if (sh) shape.shadow = sh;
    if (el.localName === 'rect') {
      const r = lengthPx(at(el, NS.draw, 'corner-radius'));
      if (r && r > 0) {
        shape.geom = 'roundRect';
        shape.adj = { adj: Math.round(Math.min(50000, (r / Math.max(1, Math.min(f.w, f.h))) * 100000)) };
      }
    } else if (el.localName === 'ellipse' || el.localName === 'circle') {
      const kind = at(el, NS.draw, 'kind') ?? 'full';
      shape.geom = 'ellipse';
      if (kind !== 'full') {
        const st = Number(at(el, NS.draw, 'start-angle') ?? 0);
        const en = Number(at(el, NS.draw, 'end-angle') ?? 360);
        shape.geom = kind === 'section' ? 'pie' : kind === 'cut' ? 'chord' : 'arc';
        // ODF angles run counter-clockwise, OOXML's clockwise
        shape.adj = { adj1: Math.round(((360 - en + 360) % 360) * 60000), adj2: Math.round(((360 - st + 360) % 360) * 60000) };
      }
    }
    const ph = this.placeholder(el, env);
    if (ph.ph) shape.ph = ph.ph;
    this.withText(shape, el, style, env);
    this.register(el, shape.id, env);
    return shape;
  }

  lineFrom(el: Element, env: ShapeEnv, x1: number, y1: number, x2: number, y2: number, geom = 'line'): ShapeEl {
    const style = this.shapeStyle(el, env.scope);
    const g = style.c.graphic;
    const shape: ShapeEl = {
      ...this.common(el, { x: r2(Math.min(x1, x2)), y: r2(Math.min(y1, y2)), w: r2(Math.abs(x2 - x1)), h: r2(Math.abs(y2 - y1)), rot: 0 }, env),
      type: 'shape',
      geom,
      line: this.line(g, true) ?? { color: '#000000', width: 1 },
    };
    if (x2 < x1) shape.flipH = true;
    if (y2 < y1) shape.flipV = true;
    const sh = this.shadow(g);
    if (sh) shape.shadow = sh;
    this.register(el, shape.id, env);
    return shape;
  }

  lineShape(el: Element, env: ShapeEnv): El | null {
    let x1 = lengthPx(at(el, NS.svg, 'x1')) ?? 0;
    let y1 = lengthPx(at(el, NS.svg, 'y1')) ?? 0;
    let x2 = lengthPx(at(el, NS.svg, 'x2')) ?? 0;
    let y2 = lengthPx(at(el, NS.svg, 'y2')) ?? 0;
    const tr = at(el, NS.draw, 'transform');
    if (tr) {
      const m = parseTransform(tr);
      [x1, y1] = [m[0] * x1 + m[2] * y1 + m[4], m[1] * x1 + m[3] * y1 + m[5]];
      [x2, y2] = [m[0] * x2 + m[2] * y2 + m[4], m[1] * x2 + m[3] * y2 + m[5]];
    }
    const geom = at(el, NS.affice, 'geometry');
    return this.lineFrom(el, env, x1, y1, x2, y2, geom && isKnownPreset(geom) ? geom : 'line');
  }

  connector(el: Element, env: ShapeEnv): El | null {
    const d = at(el, NS.svg, 'd');
    const type = at(el, NS.draw, 'type') ?? 'standard';
    if (!d || type === 'line') {
      const x1 = lengthPx(at(el, NS.svg, 'x1')) ?? 0;
      const y1 = lengthPx(at(el, NS.svg, 'y1')) ?? 0;
      const x2 = lengthPx(at(el, NS.svg, 'x2')) ?? 0;
      const y2 = lengthPx(at(el, NS.svg, 'y2')) ?? 0;
      return this.lineFrom(el, env, x1, y1, x2, y2);
    }
    return this.pathShape(el, env);
  }

  pathShape(el: Element, env: ShapeEnv): El | null {
    const vb = (at(el, NS.svg, 'viewBox') ?? '').trim().split(/[\s,]+/).map(Number);
    let f = this.frame(el);
    const cmds: CustomPath['cmds'] = [];
    const closed = el.localName === 'polygon';
    if (el.localName === 'path' || el.localName === 'connector') {
      for (const s of svgPathSegments(at(el, NS.svg, 'd') ?? '')) cmds.push({ c: s.c, p: s.p });
    } else {
      const nums = (at(el, NS.draw, 'points') ?? '').trim().split(/[\s,]+/).map(Number);
      for (let i = 0; i + 1 < nums.length; i += 2) cmds.push({ c: i ? 'L' : 'M', p: [nums[i], nums[i + 1]] });
      if (closed && cmds.length) cmds.push({ c: 'Z', p: [] });
    }
    if (!cmds.length) return null;
    // path coordinates are in view-box units; a connector's are in 1/100 mm of the page
    let [vx, vy, vw, vh] = vb.length === 4 && vb.every(Number.isFinite) ? vb : [0, 0, 0, 0];
    if (el.localName === 'connector' && !at(el, NS.svg, 'width')) {
      const xs = cmds.flatMap((c) => c.p.filter((_, i) => i % 2 === 0));
      const ys = cmds.flatMap((c) => c.p.filter((_, i) => i % 2 === 1));
      vx = Math.min(...xs);
      vy = Math.min(...ys);
      vw = Math.max(1, Math.max(...xs) - vx);
      vh = Math.max(1, Math.max(...ys) - vy);
      f = { x: (vx * 96) / 2540, y: (vy * 96) / 2540, w: (vw * 96) / 2540, h: (vh * 96) / 2540, rot: 0 };
    }
    if (!vw) vw = f.w || 1;
    if (!vh) vh = f.h || 1;
    const shifted = cmds.map((c) => ({ c: c.c, p: c.p.map((v, i) => (i % 2 === 0 ? v - vx : v - vy)) }));
    const style = this.shapeStyle(el, env.scope);
    const g = style.c.graphic;
    const open = el.localName === 'polyline' || el.localName === 'connector' || !shifted.some((c) => c.c === 'Z');
    const custom: CustomPath = { w: vw, h: vh, cmds: shifted, ...(open ? { fill: false } : {}) };
    const shape: ShapeEl = { ...this.common(el, f, env), type: 'shape', geom: 'custom', custom: [custom], fill: open ? { type: 'none' } : this.fill(g, true, f), line: this.line(g, true) };
    const sh = this.shadow(g);
    if (sh) shape.shadow = sh;
    this.withText(shape, el, style, env);
    this.register(el, shape.id, env);
    return shape;
  }

  regularPolygon(el: Element, env: ShapeEnv): El | null {
    const f = this.frame(el);
    const n = Math.max(3, Number(at(el, NS.draw, 'corners') ?? 5));
    const concave = at(el, NS.draw, 'concave') === 'true';
    const sharp = pctOf(at(el, NS.draw, 'sharpness')) ?? 0.5;
    const cmds: CustomPath['cmds'] = [];
    const pts = concave ? n * 2 : n;
    for (let i = 0; i < pts; i++) {
      const a = -Math.PI / 2 + (i * 2 * Math.PI) / pts;
      const k = concave && i % 2 ? 1 - sharp : 1;
      cmds.push({ c: i ? 'L' : 'M', p: [f.w / 2 + (f.w / 2) * k * Math.cos(a), f.h / 2 + (f.h / 2) * k * Math.sin(a)] });
    }
    cmds.push({ c: 'Z', p: [] });
    const style = this.shapeStyle(el, env.scope);
    const g = style.c.graphic;
    const shape: ShapeEl = { ...this.common(el, f, env), type: 'shape', geom: 'custom', custom: [{ w: f.w || 1, h: f.h || 1, cmds }], fill: this.fill(g, true, f), line: this.line(g, true) };
    this.withText(shape, el, style, env);
    this.register(el, shape.id, env);
    return shape;
  }

  groupOf(children: El[], el: Element): GroupEl | null {
    const b = unionBounds(children.map((c) => rotatedBounds(c)));
    if (!b) return null;
    const g: GroupEl = { id: newId(), type: 'group', x: r2(b.x), y: r2(b.y), w: r2(b.w), h: r2(b.h), children };
    const name = at(el, NS.draw, 'name');
    if (name) g.name = name;
    return g;
  }

  group(el: Element, env: ShapeEnv): El | null {
    const children = kids(el)
      .map((k) => this.shape(k, env))
      .filter((c): c is El => !!c);
    if (!children.length) return null;
    const g = this.groupOf(children, el);
    if (!g) return null;
    const desc = kid(el, NS.svg, 'desc')?.textContent;
    if (desc) g.alt = desc;
    const link = this.clickAction(el);
    if (link) g.link = link;
    this.register(el, g.id, env);
    return g;
  }

  frameShape(el: Element, env: ShapeEnv): El | null {
    const f = this.frame(el);
    const style = this.shapeStyle(el, env.scope);
    const g = style.c.graphic;
    const ph = this.placeholder(el, env);
    const table = kid(el, NS.table, 'table');
    if (table) return this.table(el, table, f, env);
    const object = kid(el, NS.draw, 'object') ?? kid(el, NS.draw, 'object-ole');
    if (object) {
      const chart = this.chart(el, object, f, env);
      if (chart) return chart;
    }
    const images = kidsOf(el, NS.draw, 'image');
    if (images.length && !(ph.ph === 'pic' && ph.empty)) {
      const img = this.image(el, images, f, style, env);
      if (img) return img;
      if (object) this.warnings.add('Some embedded objects could not be shown.');
      return null;
    }
    if (object) {
      this.warnings.add('Some embedded objects could not be shown.');
      return null;
    }
    const box = kid(el, NS.draw, 'text-box');
    if (!box && !ph.ph) return null;
    // like LibreOffice, a frame whose styles don't say otherwise is filled and outlined
    const shape: ShapeEl = { ...this.common(el, f, env), type: 'shape', geom: 'rect', fill: this.fill(g, true, f) ?? { type: 'none' }, line: this.line(g, true) };
    if (!ph.ph) shape.textbox = true;
    else {
      shape.ph = ph.ph;
      if (ph.phIdx) shape.phIdx = ph.phIdx;
    }
    const sh = this.shadow(g);
    if (sh) shape.shadow = sh;
    if (!shape.line) delete (shape as Partial<ShapeEl>).line;
    if (shape.fill?.type === 'none' && (shape.textbox || shape.ph)) delete shape.fill;
    if (box) this.withText(shape, el, style, env, box);
    else shape.text = { paras: [{ runs: [], endRun: this.runOf(style.c.text ?? new Map()) }], ...this.frameProps(g, style.c.paragraph, true) };
    if (shape.textbox && shape.text?.autofit === 'resize') delete shape.text.autofit;
    if (g?.get('style:protect')?.includes('position')) shape.locked = true;
    this.register(el, shape.id, env);
    return shape;
  }

  image(el: Element, images: Element[], f: { x: number; y: number; w: number; h: number; rot: number }, style: Resolved, env: ShapeEnv): El | null {
    const pics = images.map((i) => ({ el: i, pic: this.picture(at(i, NS.xlink, 'href'), kid(i, NS.office, 'binary-data')?.textContent ?? undefined) })).filter((p) => p.pic);
    if (!pics.length) return null;
    const svg = pics.find((p) => p.pic!.mime === 'image/svg+xml');
    const raster = pics.find((p) => p.pic!.mime !== 'image/svg+xml');
    const main = raster ?? svg!;
    const g = style.c.graphic;
    const img: ImageEl = { ...this.common(el, f, env), type: 'image', src: main.pic!.src };
    if (svg) {
      img.svg = strFromU8(svg.pic!.bytes);
      if (!raster) img.src = svg.pic!.src;
    }
    const line = this.line(g, false);
    if (line) img.line = line;
    const sh = this.shadow(g);
    if (sh) img.shadow = sh;
    const mirror = g?.get('style:mirror') ?? '';
    if (/horizontal/.test(mirror)) img.flipH = true;
    if (/vertical/.test(mirror)) img.flipV = true;
    const op = pctOf(g?.get('draw:image-opacity'));
    if (op !== undefined && op < 0.999) img.opacity = op;
    const clip = g?.get('fo:clip');
    if (clip && clip !== 'auto') {
      const m = /rect\(\s*([^,\s]+)[\s,]+([^,\s]+)[\s,]+([^,\s]+)[\s,]+([^,\s)]+)\s*\)/.exec(clip);
      const size = imageSize(main.pic!.bytes);
      if (m && size) {
        const wIn = size.w / size.dpiX;
        const hIn = size.h / size.dpiY;
        const inch = (v: string) => (lengthPx(v) ?? 0) / 96;
        const crop: [number, number, number, number] = [inch(m[4]) / wIn, inch(m[1]) / hIn, inch(m[2]) / wIn, inch(m[3]) / hIn].map((v) => Math.round(v * 10000) / 10000) as [number, number, number, number];
        if (crop.some((v) => v > 0.0001)) img.crop = crop;
      }
    }
    const ph = this.placeholder(el, env);
    if (ph.ph) {
      img.ph = ph.ph;
      if (ph.phIdx) img.phIdx = ph.phIdx;
    }
    this.register(el, img.id, env);
    return img;
  }

  chart(el: Element, object: Element, f: { x: number; y: number; w: number; h: number; rot: number }, env: ShapeEnv): ChartEl | null {
    const href = at(object, NS.xlink, 'href');
    let root: Document | Element | null = null;
    if (href) root = this.part(`${href.replace(/^\.\//, '').replace(/\/$/, '')}/content.xml`);
    else root = kid(object, NS.office, 'document') ?? null;
    if (!root) return null;
    const chart = readSlideChart(root, this.theme);
    if (!chart) return null;
    const out: ChartEl = { ...this.common(el, f, env), type: 'chart', chart };
    const ph = at(el, NS.affice, 'placeholder');
    if (ph && PH_TYPES.has(ph)) out.ph = ph as PlaceholderType;
    const idx = at(el, NS.affice, 'placeholder-index');
    if (idx) out.phIdx = Number(idx);
    this.register(el, out.id, env);
    return out;
  }

  table(el: Element, t: Element, f: { x: number; y: number; w: number; h: number; rot: number }, env: ShapeEnv): TableEl | null {
    // columns (with their default cell styles)
    const cols: { w?: number; cell?: string }[] = [];
    const colEls: Element[] = [];
    const collectCols = (p: Element) => {
      for (const k of kids(p)) {
        if (k.namespaceURI !== NS.table) continue;
        if (k.localName === 'table-column') colEls.push(k);
        else if (k.localName === 'table-columns' || k.localName === 'table-header-columns' || k.localName === 'table-column-group') collectCols(k);
      }
    };
    collectCols(t);
    for (const c of colEls) {
      const rep = Math.min(200, Number(at(c, NS.table, 'number-columns-repeated') ?? 1) || 1);
      const st = this.styles.resolve(at(c, NS.table, 'style-name'), 'table-column', env.scope);
      const w = lengthPx(st.c['table-column']?.get('style:column-width'));
      for (let i = 0; i < rep; i++) cols.push({ w, cell: at(c, NS.table, 'default-cell-style-name') });
    }
    const rowEls: Element[] = [];
    const collectRows = (p: Element) => {
      for (const k of kids(p)) {
        if (k.namespaceURI !== NS.table) continue;
        if (k.localName === 'table-row') rowEls.push(k);
        else if (k.localName === 'table-rows' || k.localName === 'table-header-rows' || k.localName === 'table-row-group') collectRows(k);
      }
    };
    collectRows(t);
    if (!rowEls.length) return null;
    // the template (LibreOffice table design) and our own table style
    const templateName = at(t, NS.table, 'template-name');
    const flags: TableStyleOpts = {};
    if (at(t, NS.table, 'use-first-row-styles') === 'true') flags.firstRow = true;
    if (at(t, NS.table, 'use-last-row-styles') === 'true') flags.lastRow = true;
    if (at(t, NS.table, 'use-first-column-styles') === 'true') flags.firstCol = true;
    if (at(t, NS.table, 'use-last-column-styles') === 'true') flags.lastCol = true;
    if (at(t, NS.table, 'use-banding-rows-styles') === 'true') flags.bandRows = true;
    if (at(t, NS.table, 'use-banding-columns-styles') === 'true') flags.bandCols = true;
    const ours = templateName ? /^affice-(medium2|light1|light2|medium1|dark1|none)(?:-([a-z0-9]+))?$/.exec(templateName) : null;
    const template = templateName && !ours ? this.styles.templates.get(templateName) : undefined;
    const rows: TableEl['rows'] = [];
    const nCols = Math.max(cols.length, ...rowEls.map((r) => kids(r).filter((c) => c.namespaceURI === NS.table && (c.localName === 'table-cell' || c.localName === 'covered-table-cell')).reduce((n, c) => n + (Number(at(c, NS.table, 'number-columns-repeated') ?? 1) || 1), 0)));
    const expandedRows: Element[] = [];
    for (const r of rowEls) {
      const rep = Math.min(200, Number(at(r, NS.table, 'number-rows-repeated') ?? 1) || 1);
      for (let i = 0; i < rep; i++) expandedRows.push(r);
    }
    const nRows = expandedRows.length;
    expandedRows.forEach((r, ri) => {
      const rs = this.styles.resolve(at(r, NS.table, 'style-name'), 'table-row', env.scope);
      const h = lengthPx(rs.c['table-row']?.get('style:row-height') ?? rs.c['table-row']?.get('style:min-row-height'));
      const cells: TableCell[] = [];
      let ci = 0;
      for (const c of kids(r)) {
        if (c.namespaceURI !== NS.table || (c.localName !== 'table-cell' && c.localName !== 'covered-table-cell')) continue;
        const rep = Math.min(200, Number(at(c, NS.table, 'number-columns-repeated') ?? 1) || 1);
        for (let k = 0; k < rep; k++, ci++) {
          if (c.localName === 'covered-table-cell') {
            cells.push({ text: { paras: [{ runs: [] }] }, merged: true });
            continue;
          }
          const styleName = at(c, NS.table, 'style-name') ?? at(r, NS.table, 'default-cell-style-name') ?? cols[ci]?.cell ?? (template ? templateCell(template, ri, ci, nRows, nCols, flags) : undefined);
          const cs = this.styles.resolve(styleName, 'table-cell', env.scope);
          const g = new Map<string, string>([...(cs.c['table-cell'] ?? new Map()), ...(cs.c.graphic ?? new Map())]);
          const para = new Map<string, string>([...(cs.c['table-cell'] ?? new Map()), ...(cs.c.paragraph ?? new Map())]);
          const text = this.textBody(c, { c: { text: cs.c.text ?? new Map(), paragraph: cs.c.paragraph ?? new Map(), graphic: g } }, env, new Map());
          const va = g.get('style:vertical-align');
          if (!text.anchor && va === 'middle') text.anchor = 'm';
          else if (!text.anchor && va === 'bottom') text.anchor = 'b';
          const cell: TableCell = { text };
          let fill = g.get('draw:fill') ? this.fill(g, false) : undefined;
          const bg = g.get('fo:background-color');
          if (!fill && bg && bg !== 'transparent') fill = { type: 'solid', color: this.color(bg, undefined) ?? bg };
          if (fill) cell.fill = fill;
          const border = (side: string) => borderOf(para.get(`fo:border-${side}`) ?? para.get('fo:border'));
          const borders: TableCell['borders'] = [border('top'), border('right'), border('bottom'), border('left')];
          if (borders.some((b) => b)) cell.borders = borders;
          const cs2 = Number(at(c, NS.table, 'number-columns-spanned') ?? 1);
          const rs2 = Number(at(c, NS.table, 'number-rows-spanned') ?? 1);
          if (cs2 > 1) cell.colSpan = cs2;
          if (rs2 > 1) cell.rowSpan = rs2;
          cells.push(cell);
        }
      }
      while (cells.length < nCols) cells.push({ text: { paras: [{ runs: [] }] } });
      rows.push({ h: h ?? 0, cells: cells.slice(0, nCols) });
    });
    const known = cols.filter((c) => c.w).reduce((a, c) => a + (c.w ?? 0), 0);
    const unknown = Math.max(0, nCols - cols.filter((c) => c.w).length);
    const widths = Array.from({ length: nCols }, (_, i) => r2(cols[i]?.w ?? (unknown ? Math.max(20, (f.w - known) / unknown) : f.w / nCols)));
    const knownH = rows.filter((r) => r.h).reduce((a, r) => a + r.h, 0);
    const unknownH = rows.filter((r) => !r.h).length;
    for (const r of rows) if (!r.h) r.h = r2(Math.max(20, unknownH ? (f.h - knownH) / unknownH : f.h / rows.length));
    const table: TableEl = { ...this.common(el, f, env), type: 'table', cols: widths, rows, w: r2(widths.reduce((a, b) => a + b, 0)), h: r2(rows.reduce((a, r) => a + r.h, 0)) };
    if (ours) {
      table.style = { family: ours[1] as TableStyleOpts['family'], ...(ours[2] ? { accent: ours[2] as ThemeColorName } : {}), ...flags };
      const sid = at(t, NS.affice, 'table-style-id');
      if (sid) table.style.styleId = sid;
    } else table.style = { family: 'none', ...flags };
    const ph = at(el, NS.affice, 'placeholder');
    if (ph && PH_TYPES.has(ph)) table.ph = ph as PlaceholderType;
    this.register(el, table.id, env);
    return table;
  }

  /* ---------------------------------------------------------- trimming */

  /** Drops text formatting that equals what the element's role gives anyway. */
  tidyText(e: El) {
    if (e.type === 'group') {
      e.children.forEach((c) => this.tidyText(c));
      return;
    }
    if (e.type === 'table') {
      // cell formatting that the table style gives anyway is left to the style
      e.rows.forEach((row, r) =>
        row.cells.forEach((cell, c) => {
          if (cell.merged) return;
          const look: CellLook = e.style?.family && e.style.family !== 'none' ? cellLook(e, r, c) : { borders: [null, null, null, null] };
          if (e.style?.family && e.style.family !== 'none') {
            if (sameFill(cell.fill, look.fill, this.theme)) delete cell.fill;
            if (sameBorders(cell.borders, look.borders, this.theme)) delete cell.borders;
          }
          stripCellText(cell.text, look, this);
        }),
      );
      return;
    }
    if (e.type !== 'shape' || !e.text) return;
    const tctx: TextContext = { pres: this.pres, role: roleOf(e.ph), ph: e.ph };
    tidyBody(e.text, tctx);
  }
}

/* ======================================================= small helpers */

function sniffMime(b: Uint8Array, href?: string): string | undefined {
  if (b.length > 4 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return 'image/png';
  if (b.length > 3 && b[0] === 0xff && b[1] === 0xd8) return 'image/jpeg';
  if (b.length > 3 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return 'image/gif';
  if (b.length > 2 && b[0] === 0x42 && b[1] === 0x4d) return 'image/bmp';
  if (b.length > 12 && b[0] === 0x52 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x46 && b[8] === 0x57 && b[9] === 0x45 && b[10] === 0x42 && b[11] === 0x50) return 'image/webp';
  const head = strFromU8(b.subarray(0, Math.min(b.length, 512))).trimStart();
  if (/^(<\?xml[^>]*>\s*)?(<!--[\s\S]*?-->\s*)*(<!DOCTYPE svg[^>]*>\s*)?<svg[\s>]/i.test(head) || (href?.toLowerCase().endsWith('.svg') && head.includes('<svg'))) return 'image/svg+xml';
  return undefined;
}

/** A 2D transform [a, b, c, d, e, f] from draw:transform (applied left to right). */
function parseTransform(s: string): number[] {
  let m = [1, 0, 0, 1, 0, 0];
  const mul = (t: number[]) => {
    // apply t after m
    m = [t[0] * m[0] + t[2] * m[1], t[1] * m[0] + t[3] * m[1], t[0] * m[2] + t[2] * m[3], t[1] * m[2] + t[3] * m[3], t[0] * m[4] + t[2] * m[5] + t[4], t[1] * m[4] + t[3] * m[5] + t[5]];
  };
  const re = /(\w+)\s*\(([^)]*)\)/g;
  let x: RegExpExecArray | null;
  while ((x = re.exec(s))) {
    const args = x[2].trim().split(/[\s,]+/).filter(Boolean);
    const num = (i: number) => Number(args[i] ?? 0);
    const len = (i: number) => lengthPx(args[i]) ?? 0;
    switch (x[1]) {
      case 'rotate': {
        // positive angles turn counter-clockwise on the page
        const a = num(0);
        mul([Math.cos(a), -Math.sin(a), Math.sin(a), Math.cos(a), 0, 0]);
        break;
      }
      case 'translate':
        mul([1, 0, 0, 1, len(0), args.length > 1 ? len(1) : 0]);
        break;
      case 'scale':
        mul([num(0), 0, 0, args.length > 1 ? num(1) : num(0), 0, 0]);
        break;
      case 'skewX':
        mul([1, 0, Math.tan(num(0)), 1, 0, 0]);
        break;
      case 'skewY':
        mul([1, Math.tan(num(0)), 0, 1, 0, 0]);
        break;
      case 'matrix':
        mul([num(0), num(1), num(2), num(3), len(4), len(5)]);
        break;
    }
  }
  return m;
}

function adjFrom(geom: string, mods: number[]): Record<string, number> | undefined {
  if (!mods.length) return undefined;
  const defs = DEFAULT_ADJ[geom];
  const keys = defs ? Object.keys(defs) : mods.length === 1 ? ['adj'] : mods.map((_, i) => `adj${i + 1}`);
  const out: Record<string, number> = {};
  mods.forEach((v, i) => {
    const k = keys[i];
    if (!k || !Number.isFinite(v)) return;
    if (defs && defs[k] === Math.round(v)) return;
    out[k] = Math.round(v);
  });
  return Object.keys(out).length ? out : undefined;
}

function dashOf(def: Element | undefined, width: number): Dash {
  if (!def) return 'dash';
  const shown = (at(def, NS.draw, 'display-name') ?? '').trim();
  if (DASHES.has(shown)) return shown as Dash;
  const rel = (v: string | undefined) => {
    if (!v) return 1;
    const p = pctOf(v);
    if (p !== undefined) return p;
    return (lengthPx(v) ?? width) / Math.max(0.1, width);
  };
  const d1 = rel(at(def, NS.draw, 'dots1-length'));
  const d2 = Number(at(def, NS.draw, 'dots2') ?? 0);
  const gap = rel(at(def, NS.draw, 'distance'));
  if (d2 > 0) return 'dashDot';
  if (d1 <= 1.5) return gap <= 1.5 ? 'sysDot' : 'dot';
  if (d1 >= 6) return 'longDash';
  return gap <= 1.5 ? 'sysDash' : 'dash';
}

function borderOf(v: string | undefined): Line | null {
  if (!v || v === 'none' || /^\s*0(\.0+)?\s*\w*\s*$/.test(v)) return null;
  const parts = v.trim().split(/\s+/);
  let width = 1;
  let color = '#000000';
  let style = 'solid';
  for (const p of parts) {
    if (/^#/.test(p)) color = normHex(p) ?? color;
    else if (lengthPx(p) !== undefined) width = lengthPx(p)!;
    else if (/^[a-z-]+$/.test(p)) style = p;
    else if (p === 'thin') width = 1;
  }
  if (style === 'none' || style === 'hidden' || width <= 0) return null;
  const out: Line = { color, width: r2(width) };
  if (style === 'dashed') out.dash = 'dash';
  else if (style === 'dotted') out.dash = 'dot';
  return out;
}

function templateCell(tpl: Element, r: number, c: number, nR: number, nC: number, f: TableStyleOpts): string | undefined {
  const part = (name: string) => at(kid(tpl, NS.table, name), NS.table, 'style-name');
  if (f.firstRow && r === 0) return part('first-row') ?? part('body');
  if (f.lastRow && r === nR - 1) return part('last-row') ?? part('body');
  if (f.firstCol && c === 0) return part('first-column') ?? part('body');
  if (f.lastCol && c === nC - 1) return part('last-column') ?? part('body');
  if (f.bandRows) return ((r - (f.firstRow ? 1 : 0)) % 2 === 0 ? part('odd-rows') : part('even-rows')) ?? part('body');
  if (f.bandCols) return ((c - (f.firstCol ? 1 : 0)) % 2 === 0 ? part('odd-columns') : part('even-columns')) ?? part('body');
  return part('body');
}

function sameColor(a: string | undefined, b: string | undefined, theme: Theme): boolean {
  if (!a || !b) return !a && !b;
  const x = resolveHex(a, theme);
  const y = resolveHex(b, theme);
  return closeHex(x.hex, y.hex, 3) && Math.abs(x.alpha - y.alpha) < 0.02;
}

/** Run colours: theme references match only the same reference (they follow the theme), plain colours by value. */
function sameRunColor(a: string, b: string | undefined, theme: Theme): boolean {
  if (!b) return false;
  const ref = (c: string) => {
    const r = parseColorRef(c);
    if (!r) return undefined;
    const alias: Record<string, string> = { dk1: 'tx1', lt1: 'bg1', dk2: 'tx2', lt2: 'bg2', text: 'tx1', background: 'bg1' };
    return colorRef(alias[r.name] ?? r.name, r.lum, r.alpha);
  };
  const ra = ref(a);
  const rb = ref(b);
  if (ra || rb) return ra === rb;
  return sameColor(a, b, theme);
}

function sameFill(a: Fill | undefined, b: Fill | undefined, theme: Theme): boolean {
  const none = (f: Fill | undefined) => !f || f.type === 'none';
  if (none(a) || none(b)) return none(a) && none(b);
  if (a!.type === 'solid' && b!.type === 'solid') return sameColor(a!.color, b!.color, theme);
  return false;
}

function sameBorders(a: TableCell['borders'], b: TableCell['borders'] | undefined, theme: Theme): boolean {
  const sides = (x: TableCell['borders'] | undefined) => x ?? [null, null, null, null];
  return sides(a).every((l, i) => {
    const m = sides(b)[i];
    if (!l || l.width <= 0) return !m || m.width <= 0;
    if (!m) return false;
    return Math.abs(l.width - m.width) < 0.2 && sameColor(l.color, m.color, theme);
  });
}

function sameRun(a: Omit<Run, 'text'>, b: Omit<Run, 'text'>): boolean {
  const keys = new Set([...Object.keys(a), ...Object.keys(b)].filter((k) => k !== 'text'));
  for (const k of keys) if ((a as Record<string, unknown>)[k] !== (b as Record<string, unknown>)[k]) return false;
  return true;
}

/** Removes run and paragraph properties equal to the defaults of the text's role. */
function tidyBody(body: TextBody, tctx: TextContext, extra: Omit<Run, 'text'> = {}) {
  const theme = tctx.pres.theme;
  const same = (k: string, a: unknown, b: unknown) => {
    if (a === undefined) return true;
    if (k === 'color' || k === 'hl') return sameRunColor(a as string, b as string | undefined, theme);
    if (k === 'size' || k === 'spacing') return Math.abs(Number(a) - Number(b ?? 0)) < 0.05;
    // with one font for headings and text, either theme reference means the same font
    if (k === 'font' && theme.fonts.major === theme.fonts.minor && /^\+(major|minor)$/.test(String(a)) && /^\+(major|minor)$/.test(String(b))) return true;
    if (typeof a === 'boolean') return a === !!b;
    return a === b;
  };
  const tidyRun = (r: Omit<Run, 'text'>, p: Para) => {
    const base = { ...effectiveRun({}, p, { ...body, defaults: undefined }, tctx), ...extra } as Record<string, unknown>;
    for (const k of Object.keys(r)) {
      if (k === 'text' || k === 'link' || k === 'field') continue;
      if (same(k, (r as Record<string, unknown>)[k], base[k])) delete (r as Record<string, unknown>)[k];
    }
  };
  for (const p of body.paras) {
    const lvl = p.level ?? 0;
    const e = effectivePara({ runs: [], level: lvl }, tctx);
    const first = p.runs.find((r) => r.text) ?? p.endRun ?? {};
    const firstColor = effectiveRun(first, p, body, tctx).color;
    // bullets
    if (p.bullet?.type === 'char' && p.bullet.color && sameColor(p.bullet.color, first.color ?? firstColor, theme)) p.bullet = { ...p.bullet, color: undefined };
    if (p.bullet?.type === 'char') {
      const b = { type: 'char' as const, char: p.bullet.char, ...(p.bullet.color ? { color: p.bullet.color } : {}), ...(p.bullet.font ? { font: p.bullet.font } : {}) };
      p.bullet = b;
    }
    const bulletSame = JSON.stringify(p.bullet ?? null) === JSON.stringify(e.bullet ?? null) || (p.bullet?.type === 'none' && e.bullet?.type === 'none');
    if (bulletSame) delete p.bullet;
    // indents are compared with what the (possibly default) bullet gives
    const pe = effectivePara({ runs: [], level: lvl, bullet: p.bullet }, tctx);
    if (p.marL !== undefined && Math.abs(p.marL - pe.marL) < 0.6) delete p.marL;
    if (p.indent !== undefined && Math.abs(p.indent - pe.indent) < 0.6) delete p.indent;
    if (p.lineSpacing !== undefined && Math.abs(p.lineSpacing - pe.lineSpacing) < 0.005) delete p.lineSpacing;
    if (p.spaceBefore !== undefined && Math.abs(p.spaceBefore - pe.spaceBefore) < 0.05) delete p.spaceBefore;
    if (p.spaceAfter !== undefined && Math.abs(p.spaceAfter - pe.spaceAfter) < 0.05) delete p.spaceAfter;
    if (p.align === 'left') delete p.align;
    if (!lvl) delete p.level;
    for (const r of p.runs) tidyRun(r, p);
    if (p.endRun) {
      tidyRun(p.endRun, p);
      if (!Object.keys(p.endRun).length) delete p.endRun;
    }
  }
}

function stripCellText(body: TextBody, look: { color?: string; bold?: boolean }, r: OdpReader) {
  const tctx: TextContext = { pres: r.pres, role: 'other' };
  const extra: Omit<Run, 'text'> = {};
  if (look.color) extra.color = look.color;
  if (look.bold) extra.b = true;
  tidyBody(body, tctx, extra);
}

function clampStops(raw: { pos: number; color: string }[]): GradientStop[] {
  const out: GradientStop[] = [];
  for (let i = 0; i < raw.length; i++) {
    const s = raw[i];
    if (s.pos <= 1) out.push({ pos: Math.max(0, s.pos), color: s.color });
    else {
      const prev = raw[i - 1];
      if (prev && prev.pos < 1) {
        const t = (1 - prev.pos) / (s.pos - prev.pos || 1);
        out.push({ pos: 1, color: mixColor(prev.color, s.color, t) });
      } else if (!out.length) out.push({ pos: 1, color: s.color });
      break;
    }
  }
  if (out.length === 1) out.push({ pos: 1, color: out[0].color });
  return out;
}

function mixColor(a: string, b: string, t: number): string {
  if (a.startsWith('@') || b.startsWith('@')) return t < 0.5 ? a : b;
  const alpha = (c: string) => (c.length === 9 ? parseInt(c.slice(7), 16) / 255 : 1);
  const hex = mixHex(a.slice(0, 7), b.slice(0, 7), t);
  return withAlpha(hex, alpha(a) + (alpha(b) - alpha(a)) * t);
}

/* ============================================================ animations */

const TRANSITION_IN: Record<string, TransitionType> = {
  fade: 'fade',
  pushWipe: 'push',
  slideWipe: 'cover',
  barWipe: 'wipe',
  barnDoorWipe: 'split',
  irisWipe: 'zoom',
  ellipseWipe: 'circle',
  dissolve: 'dissolve',
  zoom: 'zoom',
};

function transitionFrom(type: string | undefined, subtype: string | undefined, reverse: boolean): Transition | undefined {
  if (!type) return undefined;
  const t = TRANSITION_IN[type] ?? 'fade';
  const out: Transition = { type: t };
  const edge: Record<string, Dir> = { fromRight: 'l', fromLeft: 'r', fromBottom: 'u', fromTop: 'd' };
  if (t === 'push' || t === 'cover') {
    if (edge[subtype ?? '']) out.dir = edge[subtype!];
    if (t === 'cover' && reverse) out.type = 'reveal';
  } else if (t === 'wipe') {
    out.dir = subtype === 'topToBottom' ? (reverse ? 'u' : 'd') : reverse ? 'l' : 'r';
  } else if (t === 'split') out.dir = subtype === 'horizontal' ? 'u' : undefined;
  if (!out.dir) delete out.dir;
  return out;
}

const PRESET_EFFECT: [RegExp, AnimEffect][] = [
  [/appear|disappear/, 'appear'],
  [/fly-in|fly-out|peek|crawl|glide|fly/, 'fly'],
  [/ascend|descend|float|rise-up|sink-down/, 'float'],
  [/zoom|magnify|grow-and-turn|expand|contract|compress/, 'zoom'],
  [/wipe/, 'wipe'],
  [/split|barn/, 'split'],
  [/wheel|pinwheel|clock/, 'wheel'],
  [/bounce/, 'bounce'],
];

function effectFrom(cls: string, id: string): AnimEffect {
  if (cls === 'emphasis') {
    if (/grow-and-shrink|grow/.test(id)) return 'grow';
    if (/spin/.test(id)) return 'spin';
    if (/teeter|wave/.test(id)) return 'teeter';
    if (/transparency/.test(id)) return 'transparency';
    return 'pulse';
  }
  if (id === 'ooo-exit-disappear') return 'disappear';
  for (const [re, e] of PRESET_EFFECT) if (re.test(id)) return e;
  return 'fade';
}

function dirFrom(effect: AnimEffect, id: string, sub: string | undefined): Dir | undefined {
  if (effect === 'float') return /descend|sink/.test(id) ? (id.includes('exit') ? undefined : 'd') : 'u';
  const s = (sub ?? '').replace(/^to-/, 'from-');
  if (effect === 'fly' || effect === 'wipe') {
    if (s === 'from-left') return 'l';
    if (s === 'from-right') return 'r';
    if (s === 'from-top') return 'u';
    if (s === 'from-bottom') return effect === 'wipe' ? 'd' : undefined;
  }
  return undefined;
}

/** Duration of an effect node: the latest end of its animations, ignoring the instant visibility switches. */
function effectDuration(par: Element): number {
  let end = 0;
  const walk = (el: Element, offset: number) => {
    for (const k of kids(el)) {
      if (k.namespaceURI !== NS.anim) continue;
      const begin = clockMs(at(k, NS.smil, 'begin')) ?? 0;
      const dur = clockMs(at(k, NS.smil, 'dur'));
      const rev = at(k, NS.smil, 'autoReverse') === 'true' ? 2 : 1;
      const rep = Number(at(k, NS.smil, 'repeatCount') ?? 1) || 1;
      if (k.localName === 'par' || k.localName === 'seq' || k.localName === 'iterate') walk(k, offset + begin);
      else if (dur !== undefined && dur > 5) end = Math.max(end, offset + begin + dur * rev * rep);
    }
  };
  walk(par, 0);
  return Math.round(end || 500);
}

function firstTarget(el: Element): string | undefined {
  const t = at(el, NS.smil, 'targetElement');
  if (t) return t;
  for (const k of kids(el)) {
    const x = firstTarget(k);
    if (x) return x;
  }
  return undefined;
}

/* ============================================================ defaults */

function firstRun(e: El | undefined): (Omit<Run, 'text'> & { text?: string }) | undefined {
  if (!e || e.type !== 'shape' || !e.text) return undefined;
  const p = e.text.paras[0];
  if (!p) return undefined;
  const r = p.runs.find((x) => x.text) ?? p.endRun;
  return r ? { ...r } : undefined;
}

/** The most common value among runs (for text defaults when a file's styles don't say). */
function mostCommon<T>(values: (T | undefined)[]): T | undefined {
  const counts = new Map<string, { v: T; n: number }>();
  for (const v of values) {
    if (v === undefined) continue;
    const k = JSON.stringify(v);
    const hit = counts.get(k);
    if (hit) hit.n++;
    else counts.set(k, { v, n: 1 });
  }
  return [...counts.values()].sort((a, b) => b.n - a.n)[0]?.v;
}

/** Theme fonts and the master's text defaults, from our own note, the layouts' placeholders or the slides. */
function settleDefaults(R: OdpReader, pres: Presentation, masters: Element[], noted: string[] | undefined) {
  const layouts = pres.layouts;
  // the layout for ordinary content slides has a plain title (a title slide's is larger)
  const uses = new Map<string, number>();
  for (const s of pres.slides) uses.set(s.layout, (uses.get(s.layout) ?? 0) + 1);
  const content =
    layouts.find((l) => l.type === 'obj' && masters.some((m) => at(m, NS.affice, 'layout-type') === 'obj')) ??
    layouts.find((l) => /content/i.test(l.name) && !/title slide/i.test(l.name)) ??
    [...layouts].filter((l) => l.type !== 'title' && l.placeholders.some((p) => p.ph === 'title')).sort((a, b) => (uses.get(b.id) ?? 0) - (uses.get(a.id) ?? 0))[0] ??
    layouts[0];
  const slideEls = pres.slides.flatMap((s) => s.elements);
  // the first run of each placeholder says most about its default formatting
  const phRuns = (kinds: PlaceholderType[]) =>
    slideEls
      .filter((e) => e.type === 'shape' && e.ph && kinds.includes(e.ph))
      .map((e) => ((e as ShapeEl).text?.paras ?? []).filter((p) => !p.level).flatMap((p) => p.runs.filter((r) => r.text))[0])
      .filter((r): r is Run => !!r);
  const titleRuns = phRuns(['title']);
  const bodyRuns = phRuns(['obj', 'body']);
  const fromLayout = (kinds: PlaceholderType[]) => firstRun(content?.placeholders.find((p) => p.ph && kinds.includes(p.ph)));
  const titleEl = content?.placeholders.find((p) => p.ph === 'title' || p.ph === 'ctrTitle');
  const bodyEl = content?.placeholders.find((p) => p.ph === 'obj' || p.ph === 'body');
  const hasText = (e: El | undefined) => !!e && e.type === 'shape' && !!e.text?.paras.some((p) => p.runs.some((r) => r.text));
  const t = hasText(titleEl) ? fromLayout(['title', 'ctrTitle']) : undefined;
  const b = hasText(bodyEl) ? fromLayout(['obj', 'body']) : undefined;
  const tStyle = fromLayout(['title', 'ctrTitle']);
  const bStyle = fromLayout(['obj', 'body']);
  const pickFrom = <K extends keyof Run>(k: K, own: Partial<Run> | undefined, runs: Run[], style: Partial<Run> | undefined) => (own?.[k] as Run[K] | undefined) ?? mostCommon(runs.map((r) => r[k])) ?? (style?.[k] as Run[K] | undefined);
  const titleFont = pickFrom('font', t, titleRuns, tStyle);
  const bodyFont = pickFrom('font', b, bodyRuns, bStyle);
  const other = R.runOf(R.styles.resolve(undefined, 'graphic', 'styles').c.text ?? new Map());
  pres.theme.fonts = { major: noted?.[0] || titleFont || pres.theme.fonts.major, minor: noted?.[1] || bodyFont || other.font || pres.theme.fonts.minor };
  const bodyPara = bodyEl?.type === 'shape' ? bodyEl.text?.paras[0] : undefined;
  pres.master.text = {
    title: { font: '+major', size: pickFrom('size', t, titleRuns, tStyle) ?? 44, color: pickFrom('color', t, titleRuns, tStyle) ?? '@tx1', ...(pickFrom('b', t, titleRuns, tStyle) ? { bold: true } : {}) },
    body: { font: '+minor', size: pickFrom('size', b, bodyRuns, bStyle) ?? 28, color: pickFrom('color', b, bodyRuns, bStyle) ?? '@tx1', lineSpacing: bodyPara?.lineSpacing ?? 0.9 },
    other: { font: '+minor', size: other.size ?? 18, color: '@tx1' },
  };
  if (pres.master.text.title.size > 54 && titleEl?.ph === 'ctrTitle') pres.master.text.title.size = Math.round((pres.master.text.title.size * 44) / 60);
}

/** Fonts that are the theme's become theme references, so a new theme restyles them. */
function themeFontRefs(e: El, theme: Theme) {
  const map = (r: { font?: string } | undefined) => {
    if (!r?.font) return;
    if (r.font === theme.fonts.major) r.font = '+major';
    else if (r.font === theme.fonts.minor) r.font = '+minor';
  };
  const body = (t: TextBody | undefined) => {
    if (!t) return;
    map(t.defaults);
    for (const p of t.paras) {
      p.runs.forEach(map);
      map(p.endRun);
    }
  };
  if (e.type === 'group') e.children.forEach((c) => themeFontRefs(c, theme));
  else if (e.type === 'shape') body(e.text);
  else if (e.type === 'table') e.rows.forEach((r) => r.cells.forEach((c) => body(c.text)));
}

/** A master page placeholder becomes a layout placeholder whose prompt formatting is its text defaults. */
function placeholderDefaults(R: OdpReader, el: El): El {
  if (el.type !== 'shape' || !el.text) return el;
  const p0 = el.text.paras[0];
  if (!p0) return el;
  R.tidyText(el);
  const defaults: Omit<Run, 'text'> & { text?: string } = { ...(p0.runs[0] ?? p0.endRun ?? {}) };
  delete defaults.text;
  delete defaults.link;
  delete defaults.field;
  const para: Para = { runs: [], ...(p0.align ? { align: p0.align } : {}), ...(p0.bullet ? { bullet: p0.bullet } : {}) };
  const text: TextBody = { ...el.text, paras: [para] };
  if (Object.keys(defaults).length) text.defaults = defaults;
  else delete text.defaults;
  return { ...el, text };
}

/* ================================================================ import */

export async function importOdp(bytes: Uint8Array): Promise<LoadedPresentation> {
  const R = new OdpReader();
  let content: Document | null;
  let stylesDoc: Document | null;
  let metaDoc: Document | null;
  if (bytes[0] === 0x50 && bytes[1] === 0x4b) {
    try {
      R.files = unzipSync(bytes);
    } catch {
      throw new Error('This file is not a valid OpenDocument presentation.');
    }
    content = R.part('content.xml');
    stylesDoc = R.part('styles.xml');
    metaDoc = R.part('meta.xml');
  } else {
    // flat OpenDocument (.fodp): one XML file holds everything
    const flat = R.parseXml(strFromU8(bytes));
    content = flat;
    stylesDoc = flat;
    metaDoc = flat;
  }
  if (!content) throw new Error('This file is not a valid OpenDocument presentation.');
  const body = kid(kid(content.documentElement, NS.office, 'body'), NS.office, 'presentation');
  if (!body) throw new Error(kid(content.documentElement, NS.office, 'body') ? 'This OpenDocument file is not a presentation.' : 'This file is not a valid OpenDocument presentation.');
  const meta = kid(metaDoc?.documentElement, NS.office, 'meta');
  R.generator = kid(meta, NS.meta, 'generator')?.textContent ?? '';
  if (stylesDoc && stylesDoc !== content) R.styles.add(stylesDoc.documentElement, 'styles');
  R.styles.add(content.documentElement, 'content');
  if (stylesDoc === content) {
    // a flat file: its automatic styles serve the master pages too
    R.styles.add(content.documentElement, 'styles');
  }

  const userMeta = new Map<string, string>();
  for (const u of kidsOf(meta, NS.meta, 'user-defined')) userMeta.set(at(u, NS.meta, 'name') ?? '', u.textContent ?? '');

  // theme: colours from LibreOffice's theme, fonts from the title and outline styles (or our note of them)
  const masters = R.styles.masters;
  const themeEl = masters.map((m) => kid(m, NS.loext, 'theme')).find(Boolean);
  const theme: Theme = { ...R.theme, id: 'imported', colors: { ...R.theme.colors } };
  if (themeEl) {
    theme.name = at(themeEl, NS.loext, 'name') ?? 'Imported';
    for (const c of kids(kid(themeEl, NS.loext, 'theme-colors'))) {
      const name = THEME_IN[at(c, NS.loext, 'name') ?? ''];
      const hex = normHex(at(c, NS.loext, 'color'));
      if (name && hex) theme.colors[name] = hex;
    }
  } else theme.name = 'Imported';
  R.theme = theme;

  // page size from the first master's page layout
  const pageLayout = masters[0] ? R.styles.pageLayouts.get(at(masters[0], NS.style, 'page-layout-name') ?? '') : undefined;
  const plp = kid(pageLayout, NS.style, 'page-layout-properties');
  const size = { w: Math.round(lengthPx(at(plp, NS.fo, 'page-width')) ?? 1280), h: Math.round(lengthPx(at(plp, NS.fo, 'page-height')) ?? 720) };

  // master text defaults (from a master whose title is a plain title)
  const pres: Presentation = {
    format: 'affice-slides',
    version: 1,
    size,
    theme,
    master: { background: { type: 'solid', color: '@bg1' }, decor: [], text: { title: { font: '+major', size: 44, color: '@tx1' }, body: { font: '+minor', size: 28, color: '@tx1', lineSpacing: 0.9 }, other: { font: '+minor', size: 18, color: '@tx1' } } },
    layouts: [],
    slides: [],
    props: {},
  };
  R.pres = pres;
  // page names (for links between slides)
  const pages = kidsOf(body, NS.draw, 'page');
  pages.forEach((p, i) => {
    const n = at(p, NS.draw, 'name');
    if (n) R.pageIndex.set(n, i);
  });

  // footer declarations
  const footerDecls = new Map<string, string>();
  for (const d of kidsOf(body, NS.presentation, 'footer-decl')) footerDecls.set(at(d, NS.presentation, 'name') ?? '', d.textContent ?? '');
  // only today's date can be shown (a fixed date text has no place in the model)
  const liveDates = new Set(kidsOf(body, NS.presentation, 'date-time-decl').filter((d) => at(d, NS.presentation, 'source') === 'current-date').map((d) => at(d, NS.presentation, 'name') ?? ''));

  // layouts from master pages
  const layoutByMaster = new Map<string, Layout>();
  const masterFooters = new Map<string, Set<string>>();
  const bgKey = (f: Fill | undefined) => JSON.stringify(f ?? null);
  const masterBgs: (Fill | undefined)[] = [];
  const usedIds = new Set<string>();
  for (const m of masters) {
    const name = at(m, NS.style, 'name') ?? 'Master';
    const display = at(m, NS.style, 'display-name') ?? name.replace(/_([0-9a-f]+)_/gi, (_, h: string) => String.fromCodePoint(parseInt(h, 16)));
    const dp = R.styles.resolve(at(m, NS.draw, 'style-name'), 'drawing-page', 'styles');
    let bg = dp.c['drawing-page']?.get('draw:fill') && dp.c['drawing-page']?.get('draw:fill') !== 'none' ? R.fill(dp.c['drawing-page'], false, size) : undefined;
    if (!bg) {
      const bgStyle = R.styles.resolve(`${name}-background`, 'presentation', 'styles');
      const f = bgStyle.c.graphic?.get('draw:fill');
      if (f && f !== 'none') bg = R.fill(bgStyle.c.graphic, false, size);
    }
    masterBgs.push(bg);
    const env: ShapeEnv = { scope: 'styles', master: true, ids: new Map(), titleSlide: kids(m).some((k) => at(k, NS.presentation, 'class') === 'subtitle') };
    masterFooters.set(name, new Set(kids(m).map((k) => at(k, NS.presentation, 'class') ?? '').filter((c) => c === 'footer' || c === 'date-time' || c === 'page-number')));
    const decor: El[] = [];
    const placeholders: El[] = [];
    for (const k of kids(m)) {
      if (k.namespaceURI === NS.presentation && k.localName === 'notes') continue;
      if (k.namespaceURI === NS.loext) continue;
      const cls = at(k, NS.presentation, 'class');
      if (cls && ['date-time', 'footer', 'page-number', 'header', 'notes', 'page', 'handout'].includes(cls)) continue;
      const el = R.shape(k, env);
      if (!el) continue;
      if (el.ph) placeholders.push(el);
      else decor.push(el);
    }
    const typeAttr = at(m, NS.affice, 'layout-type');
    const named = LAYOUT_NAMES.find(([re]) => re.test(display.trim()))?.[1];
    if (named === 'title' && !typeAttr) {
      // a title slide converted from PowerPoint keeps only title and outline frames on its master: they are its title and subtitle
      const t = placeholders.find((p) => p.ph === 'title');
      const sub = placeholders.find((p) => p.ph === 'obj' || p.ph === 'body');
      if (t) t.ph = 'ctrTitle';
      if (sub && !placeholders.some((p) => p.ph === 'subTitle')) sub.ph = 'subTitle';
    }
    const phKinds = placeholders.map((p) => p.ph);
    const titles = phKinds.filter((k) => k === 'title' || k === 'ctrTitle').length;
    const outlines = phKinds.filter((k) => k === 'obj' || k === 'body').length;
    const type: LayoutType =
      typeAttr && LAYOUT_TYPES.has(typeAttr)
        ? (typeAttr as LayoutType)
        : named
          ? named
          : !phKinds.length
          ? 'blank'
          : phKinds.includes('subTitle')
            ? 'title'
            : titles && !outlines
              ? 'titleOnly'
              : outlines === 2
                ? 'twoObj'
                : outlines === 4
                  ? 'twoTxTwoObj'
                  : 'obj';
    let id = at(m, NS.affice, 'layout-id') ?? `layout-${name}`;
    while (usedIds.has(id)) id = `${id}-2`;
    usedIds.add(id);
    const layout: Layout = { id, name: display, type, decor, placeholders };
    if (bg) layout.background = bg;
    layoutByMaster.set(name, layout);
    pres.layouts.push(layout);
  }
  // the most common background is the master's; layouts keep only their own
  const counts = new Map<string, number>();
  for (const b of masterBgs) counts.set(bgKey(b), (counts.get(bgKey(b)) ?? 0) + 1);
  const topKey = [...counts.entries()].sort((a, b) => b[1] - a[1])[0]?.[0];
  const topBg = masterBgs.find((b) => bgKey(b) === topKey);
  if (topBg) pres.master.background = topBg;
  for (const l of pres.layouts) if (bgKey(l.background) === topKey) delete l.background;
  if (!pres.layouts.length) pres.layouts.push({ id: 'layout-blank', name: 'Blank', type: 'blank', decor: [], placeholders: [] });

  // slides
  const footerUse = { text: 0, number: 0, date: 0, titleShown: 0, titleSlides: 0 };
  let footerText: string | undefined;
  for (let i = 0; i < pages.length; i++) {
    const page = pages[i];
    const masterName = at(page, NS.draw, 'master-page-name') ?? '';
    const layout = layoutByMaster.get(masterName) ?? pres.layouts[0];
    const dp = R.styles.resolve(at(page, NS.draw, 'style-name'), 'drawing-page', 'content');
    const dpp = dp.c['drawing-page'] ?? new Map<string, string>();
    const fname = at(page, NS.presentation, 'use-footer-name');
    const env: ShapeEnv = {
      scope: 'content',
      master: false,
      ids: new Map(),
      titleSlide: kids(page).some((k) => at(k, NS.presentation, 'class') === 'subtitle'),
      footerText: fname ? footerDecls.get(fname) : undefined,
    };
    const slide: Slide = { id: newId('s'), layout: layout.id, elements: [] };
    for (const k of kids(page)) {
      if ((k.namespaceURI === NS.anim && k.localName === 'par') || (k.namespaceURI === NS.presentation && k.localName === 'notes') || k.namespaceURI === NS.office) continue;
      const el = R.shape(k, env);
      if (el) slide.elements.push(el);
    }
    // background, visibility, decorations
    const fill = dpp.get('draw:fill');
    if (fill && fill !== 'none') {
      const bg = R.fill(dpp, false, size);
      if (bg && bg.type !== 'none') slide.background = bg;
    }
    if (dpp.get('presentation:visibility') === 'hidden') slide.hidden = true;
    if (dpp.get('presentation:background-objects-visible') === 'false') slide.hideDecor = true;
    const section = at(page, NS.affice, 'section');
    if (section) slide.section = section;
    // footers (a title slide is one whose layout has a subtitle)
    const pageLayout = R.styles.presLayouts.get(at(page, NS.presentation, 'presentation-page-layout-name') ?? '');
    const titleLayout = layout.type === 'title' || (!!pageLayout && kids(pageLayout).some((k) => at(k, NS.presentation, 'object') === 'subtitle'));
    const frames = masterFooters.get(masterName) ?? new Set<string>();
    const shows = {
      text: dpp.get('presentation:display-footer') === 'true' && !!env.footerText && frames.has('footer'),
      number: dpp.get('presentation:display-page-number') === 'true' && frames.has('page-number'),
      date: dpp.get('presentation:display-date-time') === 'true' && liveDates.has(at(page, NS.presentation, 'use-date-time-name') ?? '') && frames.has('date-time'),
    };
    if (shows.text) {
      footerUse.text++;
      footerText ??= env.footerText;
    }
    if (shows.number) footerUse.number++;
    if (shows.date) footerUse.date++;
    if (titleLayout) {
      footerUse.titleSlides++;
      if (shows.text || shows.number || shows.date) footerUse.titleShown++;
    }
    // transition: the timing root's filter, else the page style
    const root = kids(page).find((k) => k.namespaceURI === NS.anim && k.localName === 'par' && at(k, NS.presentation, 'node-type') === 'timing-root');
    let tr: Transition | undefined;
    const filterPar = root ? kids(root).find((k) => k.localName === 'par' && (at(k, NS.smil, 'begin') ?? '').endsWith('.begin') && kid(k, NS.anim, 'transitionFilter')) : undefined;
    const filter = filterPar ? kid(filterPar, NS.anim, 'transitionFilter') : undefined;
    if (filter) {
      tr = transitionFrom(at(filter, NS.smil, 'type'), at(filter, NS.smil, 'subtype'), at(filter, NS.smil, 'direction') === 'reverse');
      const d = clockMs(at(filter, NS.smil, 'dur'));
      if (tr && d !== undefined) tr.dur = Math.round(d);
    } else if (dpp.get('smil:type')) {
      tr = transitionFrom(dpp.get('smil:type'), dpp.get('smil:subtype'), dpp.get('smil:direction') === 'reverse');
      const speed = dpp.get('presentation:transition-speed');
      if (tr) tr.dur = speed === 'slow' ? 1000 : speed === 'fast' ? 500 : 700;
    }
    const own = at(page, NS.affice, 'transition');
    if (own === 'morph' || own === 'flip') {
      tr = { ...(tr ?? { type: own }), type: own };
      const d = at(page, NS.affice, 'transition-direction');
      if (d === 'l' || d === 'r' || d === 'u' || d === 'd') tr.dir = d;
      else delete tr.dir;
    }
    if (dpp.get('presentation:transition-type') === 'automatic') {
      const after = durationMs(dpp.get('presentation:duration'));
      tr = { ...(tr ?? { type: 'none' }), after: after ?? 0 };
    }
    if (tr && at(page, NS.affice, 'advance-on-click') === 'false') tr.onClick = false;
    if (tr) slide.transition = tr;
    // animations
    const seq = root ? kids(root).find((k) => k.localName === 'seq' && at(k, NS.presentation, 'node-type') === 'main-sequence') : undefined;
    if (seq) {
      const anims: (Anim & { para?: number })[] = [];
      for (const click of kids(seq)) {
        if (click.localName !== 'par') continue;
        for (const group of kids(click)) {
          if (group.localName !== 'par') continue;
          for (const eff of kids(group)) {
            if (eff.localName !== 'par' && eff.localName !== 'iterate') continue;
            const nodeType = at(eff, NS.presentation, 'node-type');
            const cls = at(eff, NS.presentation, 'preset-class') ?? 'entrance';
            if (cls === 'motion-path' || cls === 'media-call' || cls === 'ole-action') continue;
            const presetId = at(eff, NS.presentation, 'preset-id') ?? '';
            const target = firstTarget(eff);
            const t = target ? env.ids.get(target) : undefined;
            if (!t) continue;
            const effect = effectFrom(cls, presetId);
            const a: Anim & { para?: number } = {
              id: newId('a'),
              el: t.el,
              cls: cls === 'exit' ? 'exit' : cls === 'emphasis' ? 'emph' : 'entr',
              effect,
              start: nodeType === 'with-previous' ? 'with' : nodeType === 'after-previous' ? 'after' : 'click',
              dur: effectDuration(eff),
              delay: Math.round(clockMs(at(eff, NS.smil, 'begin')) ?? 0),
            };
            const dir = dirFrom(effect, presetId, at(eff, NS.presentation, 'preset-sub-type'));
            if (dir) a.dir = dir;
            const rep = Number(at(eff, NS.smil, 'repeatCount') ?? 0);
            if (rep > 1) a.repeat = rep;
            if (t.para !== undefined) a.para = t.para;
            anims.push(a);
          }
        }
      }
      // effects on consecutive paragraphs of one shape become one "by paragraph" animation
      const merged: Anim[] = [];
      for (const a of anims) {
        const prev = merged[merged.length - 1] as (Anim & { para?: number; lastPara?: number }) | undefined;
        if (a.para !== undefined && prev && prev.el === a.el && prev.byPara && prev.effect === a.effect && prev.cls === a.cls && prev.dir === a.dir && a.para > (prev.lastPara ?? -1) && (a.start === 'click' ? prev.start === 'click' : a.start === 'after')) {
          prev.lastPara = a.para;
          continue;
        }
        const { para, ...rest } = a;
        const m: Anim & { lastPara?: number } = { ...rest };
        if (para !== undefined) {
          m.byPara = true;
          m.lastPara = para;
        }
        merged.push(m);
      }
      slide.anims = merged.map((m) => {
        const { lastPara: _l, ...rest } = m as Anim & { lastPara?: number };
        return rest;
      });
      if (!slide.anims.length) delete slide.anims;
    }
    // notes
    const notes = kid(page, NS.presentation, 'notes');
    if (notes) {
      const frame = kids(notes).find((k) => at(k, NS.presentation, 'class') === 'notes');
      const box = frame ? kid(frame, NS.draw, 'text-box') : undefined;
      if (box) {
        const text = R.paragraphs(box, { scope: 'content', text: new Map(), para: new Map() })
          .map((p) => p.runs.map((r) => r.text).join(''))
          .join('\n')
          .replace(/\s+$/, '');
        if (text.trim()) slide.notes = text;
      }
    }
    pres.slides.push(slide);
  }
  if (!pres.slides.length) pres.slides.push({ id: newId('s'), layout: pres.layouts[0].id, elements: [] });

  // theme fonts and default text: our note of them, else the layout placeholders, else what the slides use
  settleDefaults(R, pres, masters, userMeta.get('AfficeThemeFonts')?.split('|'));
  const allEls: El[] = [...pres.layouts.flatMap((l) => [...l.decor, ...l.placeholders]), ...pres.slides.flatMap((sl) => sl.elements)];
  for (const e of allEls) themeFontRefs(e, theme);
  for (const l of pres.layouts) {
    l.decor.forEach((d) => R.tidyText(d));
    l.placeholders = l.placeholders.map((ph) => placeholderDefaults(R, ph));
  }
  for (const sl of pres.slides) sl.elements.forEach((e) => R.tidyText(e));
  if (footerUse.text || footerUse.number || footerUse.date) {
    pres.footer = {
      ...(footerUse.text && footerText ? { text: footerText } : {}),
      ...(footerUse.number ? { slideNumber: true } : {}),
      ...(footerUse.date ? { date: true } : {}),
      ...(footerUse.titleSlides && !footerUse.titleShown && pres.slides.length > footerUse.titleSlides ? { skipTitle: true } : {}),
    };
  }

  // metadata
  const text = (ns: string, name: string) => kid(meta, ns, name)?.textContent?.trim() || undefined;
  const title = text(NS.dc, 'title');
  const subject = text(NS.dc, 'subject');
  const author = text(NS.meta, 'initial-creator') ?? text(NS.dc, 'creator');
  const created = text(NS.meta, 'creation-date');
  if (title) pres.props.title = title;
  if (subject) pres.props.subject = subject;
  if (author) pres.props.author = author;
  if (created) pres.props.created = created;
  const design = userMeta.get('AfficeDesign');
  if (design) pres.design = design;
  return { pres, warning: R.warnings.size ? [...R.warnings].join(' ') : undefined };
}
