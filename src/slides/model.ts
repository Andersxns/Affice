/**
 * Presentation model. Everything is plain JSON so it can be cloned, diffed for undo and saved as .afslides.
 *
 * Units: positions and sizes are CSS pixels at 96 dpi in slide space (a 16:9 slide is 1280 × 720,
 * i.e. 13.333 × 7.5 in); font sizes are points. Colours are CSS hex strings or theme references
 * written as "@accent1", "@accent1+40" (40 % lighter), "@accent1-25" (25 % darker), optionally with
 * "/60" for 60 % opacity, so a presentation recolours itself when its theme changes.
 */

export const DEFAULT_W = 1280;
export const DEFAULT_H = 720;
export const PT = 96 / 72;

export type ThemeColorName = 'bg1' | 'tx1' | 'bg2' | 'tx2' | 'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6' | 'hlink' | 'folHlink';
export const THEME_COLOR_NAMES: ThemeColorName[] = ['bg1', 'tx1', 'bg2', 'tx2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];

export interface Theme {
  id: string;
  name: string;
  colors: Record<ThemeColorName, string>;
  fonts: { major: string; minor: string };
}

export interface GradientStop {
  pos: number; // 0..1
  color: string;
}

export type Fill =
  | { type: 'none' }
  | { type: 'solid'; color: string }
  | { type: 'gradient'; stops: GradientStop[]; angle: number; radial?: boolean }
  | { type: 'image'; src: string; mode?: 'stretch' | 'tile' | 'cover'; alpha?: number };

export type Dash = 'solid' | 'dash' | 'dot' | 'dashDot' | 'longDash' | 'sysDash' | 'sysDot';
export type ArrowType = 'none' | 'triangle' | 'arrow' | 'stealth' | 'oval' | 'diamond';

export interface Line {
  color: string;
  width: number; // px
  dash?: Dash;
  head?: ArrowType;
  tail?: ArrowType;
  cap?: 'flat' | 'round' | 'square';
}

export interface Shadow {
  color: string;
  blur: number; // px
  dist: number; // px
  angle: number; // degrees, 0 = to the right, 90 = down
}

/* ----------------------------------------------------------------- text */

export interface Run {
  text: string;
  b?: boolean;
  i?: boolean;
  u?: boolean;
  s?: boolean;
  font?: string;
  size?: number; // pt
  color?: string;
  hl?: string;
  link?: string;
  sup?: boolean;
  sub?: boolean;
  caps?: boolean;
  spacing?: number; // pt letter spacing
  /** Live field: current slide number or today's date. */
  field?: 'slidenum' | 'date';
}

export type NumStyle = 'arabicPeriod' | 'arabicParenR' | 'alphaLcPeriod' | 'alphaUcPeriod' | 'alphaLcParenR' | 'romanLcPeriod' | 'romanUcPeriod';

export type Bullet = { type: 'none' } | { type: 'char'; char: string; color?: string; font?: string } | { type: 'num'; style: NumStyle; start?: number };

export interface Para {
  runs: Run[];
  align?: 'left' | 'center' | 'right' | 'justify';
  level?: number;
  bullet?: Bullet;
  /** Line spacing as a multiple of single spacing (1 = single). */
  lineSpacing?: number;
  spaceBefore?: number; // pt
  spaceAfter?: number; // pt
  /** Left margin and first-line indent in px (PowerPoint marL / indent). */
  marL?: number;
  indent?: number;
  /** Default run properties for empty paragraphs and new text. */
  endRun?: Omit<Run, 'text'>;
}

export interface TextBody {
  paras: Para[];
  anchor?: 't' | 'm' | 'b';
  /** Insets l, t, r, b in px. */
  inset?: [number, number, number, number];
  autofit?: 'none' | 'shrink' | 'resize';
  wrap?: boolean;
  vert?: 'horz' | 'vert' | 'vert270';
  /** Shrink factor applied by "shrink text on overflow" (1 = none). */
  fontScale?: number;
  columns?: number;
  /** Default run properties for the whole body (placeholder styling, like PowerPoint's lstStyle). */
  defaults?: Omit<Run, 'text'>;
}

/* ------------------------------------------------------------- elements */

export type PlaceholderType = 'title' | 'ctrTitle' | 'subTitle' | 'body' | 'obj' | 'pic' | 'dt' | 'ftr' | 'sldNum';

export interface ElBase {
  id: string;
  name?: string;
  x: number;
  y: number;
  w: number;
  h: number;
  rot?: number;
  flipH?: boolean;
  flipV?: boolean;
  opacity?: number;
  ph?: PlaceholderType;
  phIdx?: number;
  locked?: boolean;
  hidden?: boolean;
  /** Click action: URL or "#slide:N" / "#next" / "#prev" / "#first" / "#last". */
  link?: string;
  alt?: string;
}

export interface PathCmd {
  c: 'M' | 'L' | 'C' | 'Q' | 'A' | 'Z';
  /** Points in path space (M/L: [x,y]; C: [x1,y1,x2,y2,x,y]; Q: [x1,y1,x,y]; A: [wR,hR,stAng,swAng] in degrees). */
  p: number[];
}

export interface CustomPath {
  w: number;
  h: number;
  cmds: PathCmd[];
  fill?: boolean;
  stroke?: boolean;
}

export interface ShapeEl extends ElBase {
  type: 'shape';
  /** Preset geometry name (PowerPoint prstGeom), or "custom". */
  geom: string;
  adj?: Record<string, number>;
  custom?: CustomPath[];
  fill?: Fill;
  line?: Line | null;
  shadow?: Shadow;
  text?: TextBody;
  /** Text box semantics: no fill/line by default, grows with its text. */
  textbox?: boolean;
}

export interface ImageEl extends ElBase {
  type: 'image';
  src: string;
  /** Crop as fractions of the source image (l, t, r, b). */
  crop?: [number, number, number, number];
  line?: Line | null;
  shadow?: Shadow;
  /** Mask shape (preset geometry) — e.g. ellipse for round photos. */
  geom?: string;
  /** Vector source (icons) — rendered instead of src when present; src holds a PNG fallback. */
  svg?: string;
}

export interface TableCell {
  text: TextBody;
  fill?: Fill;
  /** Borders: top, right, bottom, left (null = none). */
  borders?: [Line | null, Line | null, Line | null, Line | null];
  rowSpan?: number;
  colSpan?: number;
  /** Covered by a merged neighbour. */
  merged?: boolean;
}

export interface TableRowData {
  h: number;
  cells: TableCell[];
}

export interface TableStyleOpts {
  /** Accent used by the built-in table style (theme colour name). */
  accent?: ThemeColorName;
  firstRow?: boolean;
  bandRows?: boolean;
  firstCol?: boolean;
  lastRow?: boolean;
  lastCol?: boolean;
  bandCols?: boolean;
  /** Style family (PowerPoint table style id is stored separately for round-trips). */
  family?: 'medium2' | 'light1' | 'light2' | 'medium1' | 'dark1' | 'none';
  styleId?: string;
}

export interface TableEl extends ElBase {
  type: 'table';
  cols: number[];
  rows: TableRowData[];
  style?: TableStyleOpts;
}

export type SlideChartType = 'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter' | 'radar';

export interface SlideChart {
  type: SlideChartType;
  title?: string;
  categories: string[];
  series: { name: string; values: (number | null)[]; color?: string }[];
  stacked?: boolean | 'percent';
  legend?: 'bottom' | 'right' | 'top' | 'left' | 'none';
  dataLabels?: boolean;
  smooth?: boolean;
  gridlines?: boolean;
  /** Text size of axis labels and the legend in pt (default 14); the title is 1.3× unless titleSize is set. */
  fontSize?: number;
  titleSize?: number;
}

export interface ChartEl extends ElBase {
  type: 'chart';
  chart: SlideChart;
}

export interface GroupEl extends ElBase {
  type: 'group';
  /** Children use slide coordinates (in the group's unrotated frame). */
  children: El[];
}

export type El = ShapeEl | ImageEl | TableEl | ChartEl | GroupEl;

/* ----------------------------------------------------------- animation */

export type TransitionType = 'none' | 'fade' | 'push' | 'wipe' | 'split' | 'cover' | 'reveal' | 'zoom' | 'morph' | 'dissolve' | 'circle' | 'flip';
export type Dir = 'l' | 'r' | 'u' | 'd';

export interface Transition {
  type: TransitionType;
  dir?: Dir;
  dur?: number; // ms
  /** Advance automatically after this many ms. */
  after?: number;
  /** Advance on mouse click (default true). */
  onClick?: boolean;
}

export type AnimClass = 'entr' | 'emph' | 'exit';
export type AnimEffect =
  | 'appear'
  | 'fade'
  | 'fly'
  | 'float'
  | 'zoom'
  | 'wipe'
  | 'split'
  | 'wheel'
  | 'bounce'
  | 'grow'
  | 'spin'
  | 'pulse'
  | 'teeter'
  | 'transparency'
  | 'disappear';

export interface Anim {
  id: string;
  el: string;
  cls: AnimClass;
  effect: AnimEffect;
  dir?: Dir;
  start: 'click' | 'with' | 'after';
  dur: number; // ms
  delay: number; // ms
  /** Animate text paragraph by paragraph. */
  byPara?: boolean;
  repeat?: number;
}

/* --------------------------------------------------------------- slides */

export type LayoutType = 'title' | 'obj' | 'secHead' | 'twoObj' | 'twoTxTwoObj' | 'titleOnly' | 'blank' | 'objTx' | 'picTx' | 'cust';

export interface Layout {
  id: string;
  name: string;
  type: LayoutType;
  background?: Fill;
  /** Decorative shapes drawn behind slide content. */
  decor: El[];
  /** Placeholder frames and default text for new slides using this layout. */
  placeholders: El[];
}

export interface Slide {
  id: string;
  layout: string;
  elements: El[];
  background?: Fill;
  notes?: string;
  transition?: Transition;
  anims?: Anim[];
  hidden?: boolean;
  /** Hide the layout's decorative shapes (PowerPoint "Hide background graphics"). */
  hideDecor?: boolean;
  section?: string;
}

/** Default text styling per placeholder role, resolved from the theme. */
export interface TextDefaults {
  title: { font: string; size: number; color: string; bold?: boolean; align?: Para['align'] };
  body: { font: string; size: number; color: string; lineSpacing?: number };
  other: { font: string; size: number; color: string };
}

export interface Presentation {
  format: 'affice-slides';
  version: 1;
  size: { w: number; h: number };
  theme: Theme;
  /** Built-in design id (for Design tab highlighting); undefined for imported files. */
  design?: string;
  master: { background: Fill; decor: El[]; text: TextDefaults };
  layouts: Layout[];
  slides: Slide[];
  props: { title?: string; author?: string; subject?: string; created?: string; modified?: string };
  /** Footer / slide number / date shown on slides (PowerPoint header & footer). */
  footer?: { text?: string; slideNumber?: boolean; date?: boolean; skipTitle?: boolean };
}

/* ============================================================== helpers */

let seq = 0;
export function newId(prefix = 'e'): string {
  seq = (seq + 1) % 1e6;
  return `${prefix}${Date.now().toString(36)}${seq.toString(36)}${Math.floor(Math.random() * 1296).toString(36)}`;
}

export function clone<T>(v: T): T {
  return structuredClone(v);
}

/** Deep copy of elements with fresh ids (keeps names so Morph can match them). */
export function cloneEls(els: El[]): El[] {
  return els.map((e) => {
    const c = clone(e);
    const re = (x: El) => {
      x.id = newId();
      if (x.type === 'group') x.children.forEach(re);
    };
    re(c);
    return c;
  });
}

export function plainText(t: TextBody | undefined): string {
  if (!t) return '';
  return t.paras.map((p) => p.runs.map((r) => r.text).join('')).join('\n');
}

export function textFromPlain(s: string, run: Omit<Run, 'text'> = {}, para: Omit<Para, 'runs'> = {}): Para[] {
  return s.split('\n').map((line) => ({ ...para, runs: line ? [{ ...run, text: line }] : [], endRun: Object.keys(run).length ? { ...run } : undefined }));
}

export function isEmptyText(t: TextBody | undefined): boolean {
  return !t || t.paras.every((p) => p.runs.every((r) => !r.text));
}

export function walkEls(els: El[], fn: (e: El, parent: GroupEl | null) => void, parent: GroupEl | null = null): void {
  for (const e of els) {
    fn(e, parent);
    if (e.type === 'group') walkEls(e.children, fn, e);
  }
}

export function findEl(els: El[], id: string): El | undefined {
  for (const e of els) {
    if (e.id === id) return e;
    if (e.type === 'group') {
      const f = findEl(e.children, id);
      if (f) return f;
    }
  }
  return undefined;
}

/** Axis-aligned bounds of an element after rotation. */
export function rotatedBounds(e: { x: number; y: number; w: number; h: number; rot?: number }): { x: number; y: number; w: number; h: number } {
  const rot = ((e.rot ?? 0) * Math.PI) / 180;
  if (!rot) return { x: e.x, y: e.y, w: e.w, h: e.h };
  const cx = e.x + e.w / 2;
  const cy = e.y + e.h / 2;
  const c = Math.abs(Math.cos(rot));
  const s = Math.abs(Math.sin(rot));
  const w = e.w * c + e.h * s;
  const h = e.w * s + e.h * c;
  return { x: cx - w / 2, y: cy - h / 2, w, h };
}

export function unionBounds(list: { x: number; y: number; w: number; h: number }[]): { x: number; y: number; w: number; h: number } | null {
  if (!list.length) return null;
  let x1 = Infinity;
  let y1 = Infinity;
  let x2 = -Infinity;
  let y2 = -Infinity;
  for (const b of list) {
    x1 = Math.min(x1, b.x);
    y1 = Math.min(y1, b.y);
    x2 = Math.max(x2, b.x + b.w);
    y2 = Math.max(y2, b.y + b.h);
  }
  return { x: x1, y: y1, w: x2 - x1, h: y2 - y1 };
}

/* ============================================================== colours */

function hexToRgb(hex: string): [number, number, number, number] | null {
  let h = hex.replace('#', '');
  if (h.length === 3 || h.length === 4) h = [...h].map((c) => c + c).join('');
  if (!/^[0-9a-f]{6}([0-9a-f]{2})?$/i.test(h)) return null;
  const n = (i: number) => parseInt(h.slice(i, i + 2), 16);
  return [n(0), n(2), n(4), h.length === 8 ? n(6) / 255 : 1];
}

export function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
  r /= 255;
  g /= 255;
  b /= 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h / 6, s, l];
}

export function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  if (s === 0) return [l * 255, l * 255, l * 255];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const f = (t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  return [f(h + 1 / 3) * 255, f(h) * 255, f(h - 1 / 3) * 255];
}

const hex2 = (v: number) => Math.round(Math.max(0, Math.min(255, v))).toString(16).padStart(2, '0');

/** Office-style luminance modulation: lumMod/lumOff as fractions. */
export function lumModOff(hex: string, lumMod: number, lumOff: number): string {
  const rgb = hexToRgb(hex);
  if (!rgb) return hex;
  const [h, s, l] = rgbToHsl(rgb[0], rgb[1], rgb[2]);
  const [r, g, b] = hslToRgb(h, s, Math.max(0, Math.min(1, l * lumMod + lumOff)));
  return `#${hex2(r)}${hex2(g)}${hex2(b)}`;
}

export interface ColorRef {
  name: string;
  /** Positive: lighter by n %, negative: darker by n %. */
  lum: number;
  alpha: number; // 0..1
}

export function parseColorRef(c: string): ColorRef | null {
  const m = /^@([a-zA-Z0-9]+)([+-]\d+)?(?:\/(\d+))?$/.exec(c);
  if (!m) return null;
  return { name: m[1], lum: m[2] ? parseInt(m[2], 10) : 0, alpha: m[3] ? parseInt(m[3], 10) / 100 : 1 };
}

export function colorRef(name: string, lum = 0, alpha = 1): string {
  return `@${name}${lum ? (lum > 0 ? `+${lum}` : `${lum}`) : ''}${alpha < 1 ? `/${Math.round(alpha * 100)}` : ''}`;
}

const ALIASES: Record<string, ThemeColorName> = { dk1: 'tx1', lt1: 'bg1', dk2: 'tx2', lt2: 'bg2', text: 'tx1', background: 'bg1' };

/** Applies an Office-like lighter/darker percentage. */
export function shiftLum(hex: string, lum: number): string {
  if (!lum) return hex;
  return lum > 0 ? lumModOff(hex, 1 - lum / 100, lum / 100) : lumModOff(hex, 1 + lum / 100, 0);
}

function withAlpha(hex: string, a: number): string {
  if (a >= 1) return hex;
  const rgb = hexToRgb(hex);
  if (!rgb) return hex;
  return `rgba(${rgb[0]},${rgb[1]},${rgb[2]},${Math.round(a * rgb[3] * 1000) / 1000})`;
}

/** Resolves a colour value (hex or theme reference) to a CSS colour. */
export function resolveColor(c: string | undefined | null, theme: Theme, fallback = '#000000'): string {
  if (!c) return fallback;
  if (c[0] !== '@') return c;
  const ref = parseColorRef(c);
  if (!ref) return fallback;
  const key = (ALIASES[ref.name] ?? ref.name) as ThemeColorName;
  const base = theme.colors[key] ?? fallback;
  return withAlpha(shiftLum(base, ref.lum), ref.alpha);
}

/** Resolves to #rrggbb plus alpha (for export formats). */
export function resolveHex(c: string | undefined | null, theme: Theme, fallback = '#000000'): { hex: string; alpha: number } {
  const css = resolveColor(c, theme, fallback);
  const m = /^rgba\((\d+),(\d+),(\d+),([\d.]+)\)$/.exec(css);
  if (m) return { hex: `#${hex2(+m[1])}${hex2(+m[2])}${hex2(+m[3])}`, alpha: +m[4] };
  const rgb = hexToRgb(css);
  if (!rgb) return { hex: fallback, alpha: 1 };
  return { hex: `#${hex2(rgb[0])}${hex2(rgb[1])}${hex2(rgb[2])}`, alpha: rgb[3] };
}

export function isDarkColor(css: string): boolean {
  const rgb = hexToRgb(css.startsWith('#') ? css : '#000000');
  if (!rgb) return false;
  return 0.2126 * rgb[0] + 0.7152 * rgb[1] + 0.0722 * rgb[2] < 128;
}

/** CSS background for a fill. */
export function fillCss(f: Fill | undefined, theme: Theme): string | undefined {
  if (!f || f.type === 'none') return undefined;
  if (f.type === 'solid') return resolveColor(f.color, theme);
  if (f.type === 'gradient') {
    const stops = [...f.stops].sort((a, b) => a.pos - b.pos).map((s) => `${resolveColor(s.color, theme)} ${Math.round(s.pos * 1000) / 10}%`).join(', ');
    return f.radial ? `radial-gradient(circle at center, ${stops})` : `linear-gradient(${(f.angle + 90) % 360}deg, ${stops})`;
  }
  const mode = f.mode ?? 'stretch';
  return `url("${f.src}") center / ${mode === 'tile' ? 'auto' : mode === 'cover' ? 'cover' : '100% 100%'} ${mode === 'tile' ? 'repeat' : 'no-repeat'}`;
}

/** Representative solid colour of a fill (for contrast decisions). */
export function fillBaseColor(f: Fill | undefined, theme: Theme): string | undefined {
  if (!f || f.type === 'none') return undefined;
  if (f.type === 'solid') return resolveHex(f.color, theme).hex;
  if (f.type === 'gradient') return resolveHex(f.stops[0]?.color, theme).hex;
  return undefined;
}

/* ======================================================== slide helpers */

export function layoutOf(pres: Presentation, slide: Slide): Layout | undefined {
  return pres.layouts.find((l) => l.id === slide.layout) ?? pres.layouts[0];
}

export function slideBackground(pres: Presentation, slide: Slide): Fill {
  return slide.background ?? layoutOf(pres, slide)?.background ?? pres.master.background;
}

export function slideSizeIn(pres: Presentation): { w: number; h: number } {
  return { w: pres.size.w / 96, h: pres.size.h / 96 };
}

/** Numbered labels for a list of paragraphs ("1.", "a)", "iv."…), or undefined for others. */
export function paragraphNumbers(paras: Para[]): (string | undefined)[] {
  const counters: number[] = [];
  const out: (string | undefined)[] = [];
  let prevLevel = -1;
  for (const p of paras) {
    const lvl = p.level ?? 0;
    const b = p.bullet;
    if (b?.type === 'num') {
      if (lvl > prevLevel || counters[lvl] === undefined) counters[lvl] = (b.start ?? 1) - 1;
      counters.length = lvl + 1;
      counters[lvl]++;
      out.push(formatNumber(counters[lvl], b.style));
    } else {
      if (lvl <= prevLevel || prevLevel < 0) counters.length = lvl;
      out.push(undefined);
    }
    prevLevel = lvl;
  }
  return out;
}

function roman(n: number): string {
  const map: [number, string][] = [
    [1000, 'm'],
    [900, 'cm'],
    [500, 'd'],
    [400, 'cd'],
    [100, 'c'],
    [90, 'xc'],
    [50, 'l'],
    [40, 'xl'],
    [10, 'x'],
    [9, 'ix'],
    [5, 'v'],
    [4, 'iv'],
    [1, 'i'],
  ];
  let s = '';
  for (const [v, r] of map)
    while (n >= v) {
      s += r;
      n -= v;
    }
  return s;
}

function alpha(n: number): string {
  let s = '';
  while (n > 0) {
    n--;
    s = String.fromCharCode(97 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}

export function formatNumber(n: number, style: NumStyle): string {
  switch (style) {
    case 'arabicParenR':
      return `${n})`;
    case 'alphaLcPeriod':
      return `${alpha(n)}.`;
    case 'alphaUcPeriod':
      return `${alpha(n).toUpperCase()}.`;
    case 'alphaLcParenR':
      return `${alpha(n)})`;
    case 'romanLcPeriod':
      return `${roman(n)}.`;
    case 'romanUcPeriod':
      return `${roman(n).toUpperCase()}.`;
    default:
      return `${n}.`;
  }
}

export const BULLET_CHARS = ['•', '◦', '▪', '–', '➢', '✓', '★', '◆', '→'];
