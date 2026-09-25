/**
 * OpenDocument presentation (.odp) writer, readable by LibreOffice Impress and other ODF apps.
 *
 * Each slide layout becomes a master page (with the design's background, decorations and placeholder
 * frames) and every slide is written with fully resolved formatting, so it looks the same in apps that
 * know nothing about Affice. Theme colours are kept as LibreOffice theme references, charts are embedded
 * chart documents with their own data table, and transitions and animations use LibreOffice's presets.
 * A few Affice-only details (placeholder kinds, layout types, Morph) ride along as attributes in an
 * Affice namespace that other apps ignore.
 */
import { strToU8, zipSync, type Zippable } from 'fflate';
import { escapeXml as x } from '@/lib/utils';
import { customGeometry, DEFAULT_ADJ, presetGeometry } from '../geometry';
import {
  isEmptyText,
  parseColorRef,
  resolveHex,
  type Anim,
  type ChartEl,
  type Dir,
  type El,
  type Fill,
  type GroupEl,
  type ImageEl,
  type Layout,
  type Line,
  type Para,
  type PlaceholderType,
  type Presentation,
  type Shadow,
  type ShapeEl,
  type Slide,
  type TableEl,
  type TextBody,
  type Theme,
  type Transition,
} from '../model';
import { DEFAULT_INSET, effectivePara, effectiveRun, roleOf, type TextContext } from '../render/text';
import { buildOf, type BuildItem } from '../show/engine';
import { cellLook, styledCellText } from '../tables';
import { resolveFont } from '../themes';
import { slideChartObject } from './odp-charts';
import { enhancedPath } from './odp-geometry';
import { decodeImage, IMAGE_MIME, imageSize, rasterizeImage } from './media';

export const NS_AFFICE = 'urn:org:affice:xmlns:slides:1.0';
const MIME = 'application/vnd.oasis.opendocument.presentation';
const XML_DECL = '<?xml version="1.0" encoding="UTF-8"?>\n';

const XMLNS = [
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
  'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"',
  'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"',
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
  'xmlns:xlink="http://www.w3.org/1999/xlink"',
  'xmlns:dc="http://purl.org/dc/elements/1.1/"',
  'xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"',
  'xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"',
  'xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"',
  'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"',
  'xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"',
  'xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"',
  'xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0"',
  'xmlns:anim="urn:oasis:names:tc:opendocument:xmlns:animation:1.0"',
  'xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"',
  'xmlns:drawooo="http://openoffice.org/2010/draw"',
  'xmlns:officeooo="http://openoffice.org/2009/office"',
  `xmlns:affice="${NS_AFFICE}"`,
].join(' ');

/* ================================================================= units */

const r3 = (v: number) => Math.round(v * 1000) / 1000;
const cm = (px: number) => `${r3((px * 2.54) / 96)}cm`;
const pt = (v: number) => `${Math.round(v * 100) / 100}pt`;
const pct = (v: number) => `${Math.round(v * 1000) / 10}%`;
const secs = (ms: number) => `${Math.round(Math.max(0, ms)) / 1000}s`;

/** An ODF style name (an XML name) for a display name, with other characters written as _hex_. */
function styleNameOf(display: string): string {
  let out = '';
  for (const ch of display) out += /[A-Za-z0-9]/.test(ch) ? ch : `_${ch.codePointAt(0)!.toString(16)}_`;
  if (!out) out = 'Master';
  if (/^[0-9]/.test(out)) out = `_${out.codePointAt(0)!.toString(16)}_${out.slice(1)}`;
  return out;
}

const quoteFont = (f: string) => (/[\s,'"]/.test(f) ? `'${f.replace(/'/g, '')}'` : f);

/* ================================================================ colours */

const THEME_TYPE: Record<string, string> = {
  tx1: 'dark1',
  bg1: 'light1',
  tx2: 'dark2',
  bg2: 'light2',
  dk1: 'dark1',
  lt1: 'light1',
  dk2: 'dark2',
  lt2: 'light2',
  text: 'dark1',
  background: 'light1',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hlink: 'hyperlink',
  folHlink: 'followed-hyperlink',
};

interface Col {
  hex: string;
  alpha: number;
  /** LibreOffice theme colour reference (loext:*-complex-color). */
  complex?: string;
}

function colorOf(c: string | undefined, theme: Theme, fallback = '#000000'): Col {
  const { hex, alpha } = resolveHex(c ?? fallback, theme, fallback);
  const ref = c ? parseColorRef(c) : null;
  const type = ref ? THEME_TYPE[ref.name] : undefined;
  if (!ref || !type) return { hex, alpha };
  let mods = '';
  if (ref.lum > 0) mods = `<loext:transformation loext:type="lummod" loext:value="${(100 - ref.lum) * 100}"/><loext:transformation loext:type="lumoff" loext:value="${ref.lum * 100}"/>`;
  else if (ref.lum < 0) mods = `<loext:transformation loext:type="lummod" loext:value="${(100 + ref.lum) * 100}"/>`;
  return { hex, alpha, complex: `loext:theme-type="${type}" loext:color-type="theme">${mods}` };
}

const complex = (kind: 'fill' | 'stroke' | 'char', c: Col) => (c.complex ? `<loext:${kind}-complex-color ${c.complex}</loext:${kind}-complex-color>` : '');

/* ================================================================= styles */

/** Automatic styles of one part (content.xml or styles.xml), shared when identical. */
class StyleSet {
  private byKey = new Map<string, string>();
  private counts = new Map<string, number>();
  readonly xml: string[] = [];
  constructor(private prefix: string) {}

  add(family: string, short: string, body: string, attrs = ''): string {
    const key = `${family}\n${attrs}\n${body}`;
    const hit = this.byKey.get(key);
    if (hit) return hit;
    const n = (this.counts.get(short) ?? 0) + 1;
    this.counts.set(short, n);
    const name = `${this.prefix}${short}${n}`;
    this.byKey.set(key, name);
    this.xml.push(family === 'list' ? `<text:list-style style:name="${name}">${body}</text:list-style>` : `<style:style style:name="${name}" style:family="${family}"${attrs}>${body}</style:style>`);
    return name;
  }
}

interface GProps {
  attrs: string[];
  kids: string[];
}

const gprops = (): GProps => ({ attrs: [], kids: [] });
const gxml = (tag: string, g: GProps) => (g.kids.length ? `<${tag} ${g.attrs.join(' ')}>${g.kids.join('')}</${tag}>` : `<${tag} ${g.attrs.join(' ')}/>`);

/* ============================================================= transforms */

/** An affine map from an element's own frame to the page (groups rotate their children). */
interface Xf {
  a: number;
  b: number;
  c: number;
  d: number;
  e: number;
  f: number;
  /** Total rotation in degrees (clockwise). */
  rot: number;
}

const IDENTITY: Xf = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0, rot: 0 };

function rotateAbout(t: Xf, deg: number, cx: number, cy: number): Xf {
  if (!deg) return t;
  const r = (deg * Math.PI) / 180;
  const cos = Math.cos(r);
  const sin = Math.sin(r);
  // q = R(p - c) + c, then apply t
  const ra = cos;
  const rb = sin;
  const rc = -sin;
  const rd = cos;
  const re = cx - cos * cx + sin * cy;
  const rf = cy - sin * cx - cos * cy;
  return {
    a: t.a * ra + t.c * rb,
    b: t.b * ra + t.d * rb,
    c: t.a * rc + t.c * rd,
    d: t.b * rc + t.d * rd,
    e: t.a * re + t.c * rf + t.e,
    f: t.b * re + t.d * rf + t.f,
    rot: t.rot + deg,
  };
}

const apply = (t: Xf, px: number, py: number): [number, number] => [t.a * px + t.c * py + t.e, t.b * px + t.d * py + t.f];

interface Placed {
  x: number;
  y: number;
  w: number;
  h: number;
  rot: number;
}

function place(e: El, t: Xf): Placed {
  const [cx, cy] = apply(t, e.x + e.w / 2, e.y + e.h / 2);
  let rot = (((e.rot ?? 0) + t.rot) % 360 + 360) % 360;
  if (rot > 180) rot -= 360;
  if (Math.abs(rot) < 0.01) rot = 0;
  return { x: cx - e.w / 2, y: cy - e.h / 2, w: Math.max(0, e.w), h: Math.max(0, e.h), rot };
}

/** svg:x/y/width/height, or a draw:transform for rotated frames. */
function frameAttrs(p: Placed): string {
  const size = `svg:width="${cm(p.w)}" svg:height="${cm(p.h)}"`;
  if (!p.rot) return `${size} svg:x="${cm(p.x)}" svg:y="${cm(p.y)}"`;
  const r = (p.rot * Math.PI) / 180;
  const cx = p.x + p.w / 2;
  const cy = p.y + p.h / 2;
  // where the unrotated top-left corner ends up after rotating about the centre
  const tlx = cx - (p.w / 2) * Math.cos(r) + (p.h / 2) * Math.sin(r);
  const tly = cy - (p.w / 2) * Math.sin(r) - (p.h / 2) * Math.cos(r);
  return `${size} draw:transform="rotate (${Number((-r).toFixed(12))}) translate (${cm(tlx)} ${cm(tly)})"`;
}

/* ============================================================ transitions */

interface SmilTransition {
  type: string;
  subtype: string;
  reverse?: boolean;
}

const EDGE: Record<Dir, string> = { l: 'fromRight', r: 'fromLeft', u: 'fromBottom', d: 'fromTop' };

function smilTransition(t: Transition): SmilTransition | null {
  const dir = t.dir ?? 'l';
  switch (t.type) {
    case 'fade':
    case 'morph':
    case 'flip':
      return { type: 'fade', subtype: 'crossfade' };
    case 'push':
      return { type: 'pushWipe', subtype: EDGE[dir] };
    case 'cover':
      return { type: 'slideWipe', subtype: EDGE[dir] };
    case 'reveal':
      return { type: 'slideWipe', subtype: EDGE[dir], reverse: true };
    case 'wipe':
      return dir === 'r' ? { type: 'barWipe', subtype: 'leftToRight' } : dir === 'l' ? { type: 'barWipe', subtype: 'leftToRight', reverse: true } : dir === 'd' ? { type: 'barWipe', subtype: 'topToBottom' } : { type: 'barWipe', subtype: 'topToBottom', reverse: true };
    case 'split':
      return { type: 'barnDoorWipe', subtype: dir === 'u' || dir === 'd' ? 'horizontal' : 'vertical' };
    case 'zoom':
      return { type: 'irisWipe', subtype: 'rectangle' };
    case 'circle':
      return { type: 'ellipseWipe', subtype: 'circle' };
    case 'dissolve':
      return { type: 'dissolve', subtype: 'default' };
    default:
      return null;
  }
}

/* ============================================================= animations */

interface Preset {
  id: string;
  sub?: string;
  xml: string;
}

const FLY_SUB: Record<Dir, string> = { l: 'from-left', r: 'from-right', u: 'from-top', d: 'from-bottom' };

function animPreset(a: Anim, target: string): Preset {
  const tgt = `smil:targetElement="${target}"`;
  const dur = Math.max(1, a.dur);
  const d = secs(dur);
  const show = `<anim:set smil:begin="0s" smil:dur="0.001s" smil:fill="hold" ${tgt} smil:attributeName="visibility" smil:to="visible"/>`;
  const hide = `<anim:set smil:begin="${secs(dur - 1)}" smil:dur="0.001s" smil:fill="hold" ${tgt} smil:attributeName="visibility" smil:to="hidden"/>`;
  const filter = (type: string, subtype: string, out: boolean, reverse = false) => `<anim:transitionFilter smil:dur="${d}" ${tgt} smil:type="${type}" smil:subtype="${subtype}"${out ? ' smil:mode="out"' : ''}${reverse ? ' smil:direction="reverse"' : ''}/>`;
  const move = (attr: string, values: string) => `<anim:animate smil:dur="${d}" smil:fill="hold" ${tgt} smil:attributeName="${attr}" smil:values="${values}" smil:keyTimes="0;1"/>`;
  const entr = a.cls === 'entr';
  const exit = a.cls === 'exit';
  const dir = a.dir ?? 'd';
  if (a.cls === 'emph') {
    switch (a.effect) {
      case 'spin':
        return { id: 'ooo-emphasis-spin', xml: `<anim:animateTransform smil:dur="${d}" smil:fill="hold" ${tgt} smil:by="360" svg:type="rotate"/>` };
      case 'grow':
        return { id: 'ooo-emphasis-grow-and-shrink', xml: `<anim:animateTransform smil:dur="${secs(dur / 2)}" smil:fill="hold" smil:autoReverse="true" ${tgt} smil:by="1.35,1.35" svg:type="scale"/>` };
      case 'teeter':
        return { id: 'ooo-emphasis-teeter', xml: `<anim:animateTransform smil:dur="${secs(dur / 4)}" smil:fill="hold" smil:autoReverse="true" smil:repeatCount="2" ${tgt} smil:by="5" svg:type="rotate"/>` };
      case 'transparency':
        return { id: 'ooo-emphasis-transparency', xml: `<anim:set smil:dur="${d}" smil:fill="hold" ${tgt} smil:attributeName="opacity" smil:to="0.5"/>` };
      default:
        return { id: 'ooo-emphasis-flash-bulb', xml: `<anim:animateTransform smil:dur="${secs(dur / 2)}" smil:fill="hold" smil:autoReverse="true" ${tgt} smil:by="1.05,1.05" svg:type="scale"/>` };
    }
  }
  switch (a.effect) {
    case 'appear':
    case 'disappear':
      return exit || a.effect === 'disappear' ? { id: 'ooo-exit-disappear', xml: `<anim:set smil:begin="0s" smil:dur="0.001s" smil:fill="hold" ${tgt} smil:attributeName="visibility" smil:to="hidden"/>` } : { id: 'ooo-entrance-appear', xml: show };
    case 'fly': {
      const x = dir === 'l' ? '0-width/2' : dir === 'r' ? '1+width/2' : 'x';
      const y = dir === 'd' ? '1+height/2' : dir === 'u' ? '0-height/2' : 'y';
      return entr
        ? { id: 'ooo-entrance-fly-in', sub: FLY_SUB[dir], xml: show + move('x', `${x};x`) + move('y', `${y};y`) }
        : { id: 'ooo-exit-fly-out', sub: FLY_SUB[dir].replace('from-', 'to-'), xml: move('x', `x;${x}`) + move('y', `y;${y}`) + hide };
    }
    case 'float':
      return entr
        ? { id: dir === 'd' ? 'ooo-entrance-descend' : 'ooo-entrance-ascend', xml: show + filter('fade', 'crossfade', false) + move('x', 'x;x') + move('y', dir === 'd' ? 'y-.1;y' : 'y+.1;y') }
        : { id: dir === 'u' ? 'ooo-exit-ascend' : 'ooo-exit-descend', xml: filter('fade', 'crossfade', true) + move('x', 'x;x') + move('y', dir === 'u' ? 'y;y-.1' : 'y;y+.1') + hide };
    case 'zoom':
      return entr
        ? { id: 'ooo-entrance-fade-in-and-zoom', xml: show + move('width', '0;width') + move('height', '0;height') + filter('fade', 'crossfade', false) }
        : { id: 'ooo-exit-fade-out-and-zoom', xml: move('width', 'width;0') + move('height', 'height;0') + filter('fade', 'crossfade', true) + hide };
    case 'wipe': {
      const f = dir === 'r' ? ['leftToRight', true] : dir === 'l' ? ['leftToRight', false] : dir === 'u' ? ['topToBottom', false] : ['topToBottom', true];
      // our direction names where the wipe comes from: d = from the bottom, u = from the top
      const sub = { d: 'from-bottom', u: 'from-top', l: 'from-left', r: 'from-right' }[dir];
      return entr ? { id: 'ooo-entrance-wipe', sub, xml: show + filter('barWipe', f[0] as string, false, f[1] as boolean) } : { id: 'ooo-exit-wipe', sub, xml: filter('barWipe', f[0] as string, true, f[1] as boolean) + hide };
    }
    case 'split':
      return entr ? { id: 'ooo-entrance-split', sub: 'vertical-out', xml: show + filter('barnDoorWipe', 'vertical', false) } : { id: 'ooo-exit-split', sub: 'vertical-in', xml: filter('barnDoorWipe', 'vertical', true, true) + hide };
    case 'wheel':
      return entr ? { id: 'ooo-entrance-wheel', sub: '1', xml: show + filter('pinWheelWipe', 'oneBlade', false) } : { id: 'ooo-exit-wheel', sub: '1', xml: filter('pinWheelWipe', 'oneBlade', true) + hide };
    case 'bounce':
      return entr ? { id: 'ooo-entrance-bounce', xml: show + filter('fade', 'crossfade', false) + move('y', 'y-0.25;y') } : { id: 'ooo-exit-bounce', xml: filter('fade', 'crossfade', true) + move('y', 'y;y+0.25') + hide };
    default:
      return entr ? { id: 'ooo-entrance-fade-in', xml: show + filter('fade', 'crossfade', false) } : { id: 'ooo-exit-fade-out', xml: filter('fade', 'crossfade', true) + hide };
  }
}

/* ================================================================= writer */

interface ElCtx {
  styles: StyleSet;
  layer: 'layout' | 'backgroundobjects';
  master: boolean;
  /** Master page style name prefix for presentation styles. */
  masterName: string;
  slideNumber?: number;
  /** Element ids that animations target → xml:id. */
  ids: Map<string, string>;
  /** Paragraph targets: element id → paragraph index → xml:id. */
  paraIds: Map<string, Map<number, string>>;
}

const PH_CLASS: Partial<Record<PlaceholderType, string>> = { title: 'title', ctrTitle: 'title', subTitle: 'subtitle', body: 'outline', obj: 'outline', pic: 'graphic' };
const PH_PARENT: Partial<Record<PlaceholderType, string>> = { title: 'title', ctrTitle: 'title', subTitle: 'subtitle', body: 'outline1', obj: 'outline1', pic: 'backgroundobjects' };
const ANCHOR: Record<string, string> = { t: 'top', m: 'middle', b: 'bottom' };
const ALIGN: Record<string, string> = { left: 'start', center: 'center', right: 'end', justify: 'justify' };
const NUM_FORMAT: Record<string, [string, string]> = {
  arabicPeriod: ['1', '.'],
  arabicParenR: ['1', ')'],
  alphaLcPeriod: ['a', '.'],
  alphaUcPeriod: ['A', '.'],
  alphaLcParenR: ['a', ')'],
  romanLcPeriod: ['i', '.'],
  romanUcPeriod: ['I', '.'],
};
const DASH: Record<string, [string, string, string, string]> = {
  // dots1 length, dots2 count, dots2 length, distance (relative to the line width)
  dash: ['400%', '0', '0%', '300%'],
  dot: ['100%', '0', '0%', '200%'],
  sysDot: ['100%', '0', '0%', '100%'],
  sysDash: ['300%', '0', '0%', '100%'],
  dashDot: ['400%', '1', '100%', '300%'],
  longDash: ['800%', '0', '0%', '300%'],
};
const MARKER: Record<string, [string, string]> = {
  triangle: ['0 0 300 300', 'M150 0L300 300L0 300Z'],
  arrow: ['0 0 300 300', 'M150 0L300 270L255 300L150 105L45 300L0 270Z'],
  stealth: ['0 0 300 300', 'M150 0L300 300L150 225L0 300Z'],
  oval: ['0 0 300 300', 'M150 0C233 0 300 67 300 150C300 233 233 300 150 300C67 300 0 233 0 150C0 67 67 0 150 0Z'],
  diamond: ['0 0 300 300', 'M150 0L300 150L150 300L0 150Z'],
};

class OdpWriter {
  readonly theme: Theme;
  readonly content = new StyleSet('');
  readonly masterStyles = new StyleSet('M');
  /** Named drawing styles (gradients, bitmaps, markers, dashes) in office:styles. */
  private named = new Map<string, { name: string; xml: string }>();
  private namedCount = 0;
  readonly files: Record<string, Uint8Array> = {};
  readonly manifest: string[] = [];
  private pictures = new Map<string, Promise<{ href: string; bytes: Uint8Array } | null>>();
  private picCount = 0;
  private objCount = 0;
  private idCount = 0;
  readonly pageNames: string[];

  constructor(readonly pres: Presentation) {
    this.theme = pres.theme;
    this.pageNames = pres.slides.map((_, i) => `page${i + 1}`);
  }

  nextId(prefix = 'id'): string {
    return `${prefix}${++this.idCount}`;
  }

  namedDef(key: string, make: (name: string) => string, base: string): string {
    const hit = this.named.get(key);
    if (hit) return hit.name;
    const name = `${base}${++this.namedCount}`;
    this.named.set(key, { name, xml: make(name) });
    return name;
  }

  namedXml(): string {
    return [...this.named.values()].map((n) => n.xml).join('');
  }

  /** Stores a picture once and returns its path in the package. */
  picture(src: string): Promise<{ href: string; bytes: Uint8Array } | null> {
    let hit = this.pictures.get(src);
    if (!hit) {
      hit = decodeImage(src).then((img) => {
        if (!img) return null;
        const ext = img.ext === 'jpeg' ? 'jpg' : img.ext;
        const href = `Pictures/image${++this.picCount}.${ext}`;
        this.files[href] = img.data;
        this.manifest.push(`<manifest:file-entry manifest:full-path="${href}" manifest:media-type="${IMAGE_MIME[img.ext] ?? 'image/png'}"/>`);
        return { href, bytes: img.data };
      });
      this.pictures.set(src, hit);
    }
    return hit;
  }

  addFile(path: string, bytes: Uint8Array, mime: string): string {
    this.files[path] = bytes;
    this.manifest.push(`<manifest:file-entry manifest:full-path="${x(path)}" manifest:media-type="${mime}"/>`);
    return path;
  }

  /* ------------------------------------------------------------- fills */

  gradientName(f: Extract<Fill, { type: 'gradient' }>, mul: number): { fill: string; opacity?: string } {
    const stops = [...f.stops].sort((a, b) => a.pos - b.pos);
    if (!stops.length) stops.push({ pos: 0, color: '#ffffff' }, { pos: 1, color: '#ffffff' });
    if (stops.length === 1) stops.push({ pos: 1, color: stops[0].color });
    const cols = stops.map((s) => colorOf(s.color, this.theme));
    // radial gradients run from the edge (offset 0) to the centre in ODF; an "ellipsoid" with a 29 %
    // border ends exactly at the shape's inscribed ellipse, like the on-screen gradient
    const radial = !!f.radial;
    const odfStops = stops.map((s, i) => ({ pos: radial ? 1 - s.pos : s.pos, c: cols[i], ref: parseColorRef(s.color) ? s.color.replace(/\/\d+$/, '') : undefined }));
    if (radial) odfStops.reverse();
    const angle = radial ? 0 : Math.round((((90 - f.angle) % 360) + 360) % 360);
    const style = radial ? 'draw:style="ellipsoid" draw:cx="50%" draw:cy="50%"' : `draw:style="linear" draw:angle="${angle}deg"`;
    const border = radial ? '29%' : '0%';
    const first = odfStops[0].c.hex;
    const last = odfStops[odfStops.length - 1].c.hex;
    // theme colours ride along for Affice (LibreOffice keeps plain colours on gradient stops)
    const stopXml = odfStops.map((s) => `<loext:gradient-stop svg:offset="${r3(s.pos)}" loext:color-type="rgb" loext:color-value="${s.c.hex}"${s.ref ? ` affice:color="${x(s.ref)}"` : ''}/>`).join('');
    const key = `g|${style}|${stopXml}`;
    const fill = this.namedDef(key, (name) => `<draw:gradient draw:name="${name}" draw:display-name="${name}" ${style} draw:start-color="${first}" draw:end-color="${last}" draw:start-intensity="100%" draw:end-intensity="100%" draw:border="${border}">${stopXml}</draw:gradient>`, 'Gradient_');
    const alphas = odfStops.map((s) => s.c.alpha * mul);
    if (alphas.every((a) => a >= 0.999)) return { fill };
    const ostops = odfStops.map((s, i) => `<loext:opacity-stop svg:offset="${r3(s.pos)}" svg:stop-opacity="${r3(alphas[i])}"/>`).join('');
    const okey = `o|${style}|${ostops}`;
    const opacity = this.namedDef(okey, (name) => `<draw:opacity draw:name="${name}" draw:display-name="${name}" ${style} draw:start="${pct(alphas[0])}" draw:end="${pct(alphas[alphas.length - 1])}" draw:border="${border}">${ostops}</draw:opacity>`, 'Transparency_');
    return { fill, opacity };
  }

  async fillProps(f: Fill | undefined, g: GProps, mul = 1): Promise<void> {
    if (!f || f.type === 'none') {
      g.attrs.push('draw:fill="none"');
      return;
    }
    if (f.type === 'solid') {
      const c = colorOf(f.color, this.theme);
      g.attrs.push('draw:fill="solid"', `draw:fill-color="${c.hex}"`);
      const a = c.alpha * mul;
      if (a < 0.999) g.attrs.push(`draw:opacity="${pct(a)}"`);
      g.kids.push(complex('fill', c));
      return;
    }
    if (f.type === 'gradient') {
      const { fill, opacity } = this.gradientName(f, mul);
      g.attrs.push('draw:fill="gradient"', `draw:fill-gradient-name="${fill}"`, `draw:gradient-step-count="0"`);
      if (opacity) g.attrs.push(`draw:opacity-name="${opacity}"`);
      return;
    }
    const pic = await this.picture(f.src);
    if (!pic) {
      g.attrs.push('draw:fill="none"');
      return;
    }
    const name = this.namedDef(`b|${pic.href}`, (n) => `<draw:fill-image draw:name="${n}" draw:display-name="${n}" xlink:href="${pic.href}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>`, 'Bitmap_');
    g.attrs.push('draw:fill="bitmap"', `draw:fill-image-name="${name}"`, `style:repeat="${f.mode === 'tile' ? 'repeat' : 'stretch'}"`);
    const a = (f.alpha ?? 1) * mul;
    if (a < 0.999) g.attrs.push(`draw:opacity="${pct(a)}"`);
  }

  lineProps(l: Line | null | undefined, g: GProps, mul = 1): void {
    if (!l || l.width <= 0) {
      g.attrs.push('draw:stroke="none"');
      return;
    }
    const c = colorOf(l.color, this.theme);
    const dash = l.dash && l.dash !== 'solid' ? DASH[l.dash] : undefined;
    if (dash) {
      const [len1, n2, len2, dist] = dash;
      const name = this.namedDef(`d|${l.dash}`, (n) => `<draw:stroke-dash draw:name="${n}" draw:display-name="${l.dash}" draw:style="rect" draw:dots1="1" draw:dots1-length="${len1}"${n2 !== '0' ? ` draw:dots2="${n2}" draw:dots2-length="${len2}"` : ''} draw:distance="${dist}"/>`, 'Dash_');
      g.attrs.push('draw:stroke="dash"', `draw:stroke-dash="${name}"`);
    } else g.attrs.push('draw:stroke="solid"');
    g.attrs.push(`svg:stroke-width="${cm(l.width)}"`, `svg:stroke-color="${c.hex}"`, 'draw:stroke-linejoin="round"', `svg:stroke-linecap="${l.cap === 'round' ? 'round' : l.cap === 'square' ? 'square' : 'butt'}"`);
    const a = c.alpha * mul;
    if (a < 0.999) g.attrs.push(`svg:stroke-opacity="${pct(a)}"`);
    const marker = (end: 'start' | 'end', type: string | undefined) => {
      if (!type || type === 'none' || !MARKER[type]) return;
      const [vb, d] = MARKER[type];
      const name = this.namedDef(`m|${type}`, (n) => `<draw:marker draw:name="${n}" draw:display-name="${type}" svg:viewBox="${vb}" svg:d="${d}"/>`, 'Arrow_');
      g.attrs.push(`draw:marker-${end}="${name}"`, `draw:marker-${end}-width="${cm(Math.max(7, l.width * 3.2))}"`, `draw:marker-${end}-center="false"`);
    };
    marker('start', l.head);
    marker('end', l.tail);
    g.kids.push(complex('stroke', c));
  }

  shadowProps(s: Shadow | undefined, g: GProps): void {
    if (!s) return;
    const c = colorOf(s.color, this.theme);
    const r = (s.angle * Math.PI) / 180;
    g.attrs.push('draw:shadow="visible"', `draw:shadow-offset-x="${cm(Math.cos(r) * s.dist)}"`, `draw:shadow-offset-y="${cm(Math.sin(r) * s.dist)}"`, `draw:shadow-color="${c.hex}"`, `draw:shadow-opacity="${pct(c.alpha)}"`, `loext:shadow-blur="${cm(s.blur / 2)}"`);
  }

  /** Text frame properties: anchoring, insets, wrapping, fitting, direction and columns. */
  textFrameProps(body: TextBody | undefined, g: GProps, textbox: boolean, frame = textbox): string {
    const b = body ?? { paras: [] };
    const inset = b.inset ?? DEFAULT_INSET;
    g.attrs.push(`draw:textarea-vertical-align="${ANCHOR[b.anchor ?? 't']}"`, 'draw:textarea-horizontal-align="justify"');
    g.attrs.push(`fo:padding-left="${cm(inset[0])}"`, `fo:padding-top="${cm(inset[1])}"`, `fo:padding-right="${cm(inset[2])}"`, `fo:padding-bottom="${cm(inset[3])}"`);
    g.attrs.push(`fo:wrap-option="${b.wrap === false ? 'no-wrap' : 'wrap'}"`);
    const grow = b.autofit === 'resize' || (textbox && b.autofit === undefined);
    // a text frame that doesn't wrap grows sideways (that is how LibreOffice keeps frame text on one line)
    g.attrs.push(`draw:auto-grow-height="${grow}"`, `draw:auto-grow-width="${b.wrap === false && frame}"`, 'draw:fit-to-size="false"', `style:shrink-to-fit="${b.autofit === 'shrink'}"`);
    if (textbox) g.attrs.push('fo:min-height="0cm"', 'fo:min-width="0cm"');
    if (b.columns && b.columns > 1) g.kids.push(`<style:columns fo:column-count="${b.columns}" fo:column-gap="${cm(24)}"/>`);
    return b.vert === 'vert' || b.vert === 'vert270' ? '<style:paragraph-properties style:writing-mode="tb-rl"/>' : '<style:paragraph-properties style:writing-mode="lr-tb"/>';
  }

  /* -------------------------------------------------------------- text */

  runProps(r: ReturnType<typeof effectiveRun>, mul = 1): string {
    const at: string[] = [];
    const font = resolveFont(r.font, this.theme);
    if (font) at.push(`fo:font-family="${x(quoteFont(font))}"`, `style:font-family-asian="${x(quoteFont(font))}"`, `style:font-family-complex="${x(quoteFont(font))}"`);
    at.push(`fo:font-size="${pt(r.size)}"`, `style:font-size-asian="${pt(r.size)}"`, `style:font-size-complex="${pt(r.size)}"`);
    const c = colorOf(r.color, this.theme);
    at.push(`fo:color="${c.hex}"`);
    const a = c.alpha * mul;
    if (a < 0.999) at.push(`loext:opacity="${pct(a)}"`);
    const bold = r.b ? 'bold' : 'normal';
    const italic = r.i ? 'italic' : 'normal';
    at.push(`fo:font-weight="${bold}"`, `style:font-weight-asian="${bold}"`, `style:font-weight-complex="${bold}"`);
    at.push(`fo:font-style="${italic}"`, `style:font-style-asian="${italic}"`, `style:font-style-complex="${italic}"`);
    at.push(r.u ? 'style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"' : 'style:text-underline-style="none"');
    at.push(r.s ? 'style:text-line-through-style="solid" style:text-line-through-type="single"' : 'style:text-line-through-style="none"');
    if (r.sup) at.push('style:text-position="30% 66%"');
    else if (r.sub) at.push('style:text-position="-25% 66%"');
    if (r.caps) at.push('fo:text-transform="uppercase"');
    if (r.spacing) at.push(`fo:letter-spacing="${pt(r.spacing)}"`);
    if (r.hl) at.push(`fo:background-color="${colorOf(r.hl, this.theme).hex}"`);
    const kids = complex('char', c);
    return kids ? `<style:text-properties ${at.join(' ')}>${kids}</style:text-properties>` : `<style:text-properties ${at.join(' ')}/>`;
  }

  /** Paragraphs (and lists) of a text body. */
  textXml(body: TextBody, tctx: TextContext, ctx: ElCtx, opts: { paraIds?: Map<number, string>; mul?: number } = {}): string {
    const paras = body.paras.length ? body.paras : [{ runs: [] } as Para];
    const styles = ctx.styles;
    let out = '';
    let i = 0;
    while (i < paras.length) {
      const e = effectivePara(paras[i], tctx);
      if (!e.bullet || e.bullet.type === 'none') {
        out += this.paraXml(paras[i], i, body, tctx, ctx, false, opts);
        i++;
        continue;
      }
      // a run of bulleted paragraphs becomes one list (so numbering continues) with one level style per depth
      let j = i;
      const levels = new Map<number, string>();
      while (j < paras.length) {
        const p = paras[j];
        const pe = effectivePara(p, tctx);
        if (!pe.bullet || pe.bullet.type === 'none') break;
        const lvl = Math.min(9, p.level ?? 0);
        if (!levels.has(lvl)) levels.set(lvl, this.listLevelXml(lvl + 1, p, pe, body, tctx));
        j++;
      }
      let listBody = '';
      for (let l = 0; l < 10; l++) listBody += levels.get(l) ?? `<text:list-level-style-bullet text:level="${l + 1}" text:bullet-char="•"><style:list-level-properties text:space-before="${cm(l * 48)}" text:min-label-width="${cm(24)}"/></text:list-level-style-bullet>`;
      const listName = styles.add('list', 'L', listBody);
      out += `<text:list text:style-name="${listName}">`;
      for (let k = i; k < j; k++) {
        const lvl = Math.min(9, paras[k].level ?? 0);
        out += `<text:list-item>${'<text:list><text:list-item>'.repeat(lvl)}${this.paraXml(paras[k], k, body, tctx, ctx, true, opts)}${'</text:list-item></text:list>'.repeat(lvl)}</text:list-item>`;
      }
      out += '</text:list>';
      i = j;
    }
    return out;
  }

  listLevelXml(level: number, p: Para, e: ReturnType<typeof effectivePara>, body: TextBody, tctx: TextContext): string {
    const firstRun = effectiveRun(p.runs.find((r) => r.text) ?? p.endRun ?? {}, p, body, tctx);
    const indent = Math.min(0, e.indent);
    const before = Math.max(0, e.marL + indent);
    const props = `<style:list-level-properties${before > 0.01 ? ` text:space-before="${cm(before)}"` : ''} text:min-label-width="${cm(-indent)}"/>`;
    const b = e.bullet!;
    const color = colorOf(b.type === 'char' && b.color ? b.color : firstRun.color, this.theme).hex;
    if (b.type === 'num') {
      const [fmt, suffix] = NUM_FORMAT[b.style] ?? ['1', '.'];
      return `<text:list-level-style-number text:level="${level}" style:num-format="${fmt}" style:num-suffix="${x(suffix)}"${b.start && b.start !== 1 ? ` text:start-value="${b.start}"` : ''}>${props}<style:text-properties fo:color="${color}" fo:font-size="100%"/></text:list-level-style-number>`;
    }
    const char = b.type === 'char' ? b.char : '•';
    // like PowerPoint, bullets without a font of their own use Arial: LibreOffice misplaces bullets whose
    // font isn't installed, and every system has Arial or a metric twin
    const font = (b.type === 'char' && b.font) || 'Arial';
    return `<text:list-level-style-bullet text:level="${level}" text:bullet-char="${x(char)}">${props}<style:text-properties fo:font-family="${x(quoteFont(font))}" fo:color="${color}" fo:font-size="100%"/></text:list-level-style-bullet>`;
  }

  paraXml(p: Para, index: number, body: TextBody, tctx: TextContext, ctx: ElCtx, inList: boolean, opts: { paraIds?: Map<number, string>; mul?: number }): string {
    const e = effectivePara(p, tctx);
    const mul = opts.mul ?? 1;
    const at: string[] = [];
    at.push(`fo:text-align="${ALIGN[p.align ?? 'left']}"`);
    if (inList) at.push('fo:margin-left="0cm"', 'fo:text-indent="0cm"');
    else at.push(`fo:margin-left="${cm(e.marL)}"`, `fo:text-indent="${cm(e.indent)}"`);
    at.push(`fo:margin-top="${pt(e.spaceBefore)}"`, `fo:margin-bottom="${pt(e.spaceAfter)}"`, `fo:line-height="${pct(e.lineSpacing)}"`);
    const def = effectiveRun(p.runs.length ? (p.runs.find((r) => r.text) ?? {}) : (p.endRun ?? {}), p, body, tctx);
    const pstyle = ctx.styles.add('paragraph', 'P', `<style:paragraph-properties ${at.join(' ')}/>${this.runProps(def, mul)}`);
    const id = opts.paraIds?.get(index);
    let inner = '';
    let prevSpace = true; // a leading space would be dropped
    for (const r of p.runs) {
      if (!r.text && !r.field) continue;
      const er = effectiveRun(r, p, body, tctx);
      const tstyle = ctx.styles.add('text', 'T', this.runProps(er, mul));
      let content: string;
      if (r.field === 'slidenum') content = `<text:page-number text:select-page="current">${x(tctx.slideNumber ? String(tctx.slideNumber) : r.text || '#')}</text:page-number>`;
      else if (r.field === 'date') content = `<text:date>${x(r.text || new Date().toLocaleDateString())}</text:date>`;
      else {
        const enc = encodeText(r.text, prevSpace);
        content = enc.xml;
        prevSpace = enc.endsWithSpace;
      }
      if (r.field) prevSpace = false;
      let span = `<text:span text:style-name="${tstyle}">${content}</text:span>`;
      const href = r.link ? this.linkHref(r.link) : undefined;
      if (href) span = `<text:a xlink:type="simple" xlink:href="${x(href)}">${span}</text:a>`;
      inner += span;
    }
    return `<text:p${id ? ` xml:id="${id}" text:id="${id}"` : ''} text:style-name="${pstyle}">${inner}</text:p>`;
  }

  /** A hyperlink target: URLs as they are, slide jumps as #pageN. */
  linkHref(link: string): string | undefined {
    const m = /^#slide:(\d+)$/.exec(link);
    if (m) return this.pageNames[Number(m[1]) - 1] ? `#${this.pageNames[Number(m[1]) - 1]}` : undefined;
    if (link.startsWith('#')) return undefined;
    return link;
  }

  /** Click action of a shape (next slide, a slide, a web page). */
  eventXml(link: string | undefined): string {
    if (!link) return '';
    const actions: Record<string, string> = { '#next': 'next-page', '#prev': 'previous-page', '#first': 'first-page', '#last': 'last-page' };
    let listener: string;
    if (actions[link]) listener = `<presentation:event-listener script:event-name="dom:click" presentation:action="${actions[link]}"/>`;
    else {
      const href = this.linkHref(link);
      if (!href) return '';
      listener = `<presentation:event-listener script:event-name="dom:click" presentation:action="show" xlink:href="${x(href)}" xlink:type="simple" xlink:show="embed" xlink:actuate="onRequest"/>`;
    }
    return `<office:event-listeners>${listener}</office:event-listeners>`;
  }

  /* ---------------------------------------------------------- elements */

  commonAttrs(e: El, ctx: ElCtx): string {
    const at: string[] = [];
    if (e.name) at.push(`draw:name="${x(e.name)}"`);
    at.push(`draw:layer="${ctx.layer}"`);
    const id = ctx.ids.get(e.id);
    if (id) at.push(`xml:id="${id}" draw:id="${id}"`);
    if (e.hidden) at.push('draw:display="none"');
    return at.join(' ');
  }

  phAttrs(e: El, ctx: ElCtx, cls: string | undefined, empty: boolean): string {
    if (!e.ph) return '';
    const at: string[] = [];
    if (cls) {
      at.push(`presentation:class="${cls}"`);
      if (empty && !ctx.master) at.push('presentation:placeholder="true"');
      at.push('presentation:user-transformed="true"');
    }
    at.push(`affice:placeholder="${e.ph}"`);
    if (e.phIdx) at.push(`affice:placeholder-index="${e.phIdx}"`);
    return at.join(' ');
  }

  /** A graphic style (or a presentation style for placeholders). */
  styleAttr(e: El, ctx: ElCtx, cls: string | undefined, g: GProps, extra = ''): string {
    const body = `${gxml('style:graphic-properties', g)}${extra}`;
    if (cls && e.ph && PH_PARENT[e.ph]) {
      const name = ctx.styles.add('presentation', 'pr', body, ` style:parent-style-name="${ctx.masterName}-${PH_PARENT[e.ph]}"`);
      return `presentation:style-name="${name}"`;
    }
    return `draw:style-name="${ctx.styles.add('graphic', 'gr', body)}"`;
  }

  async elXml(e: El, ctx: ElCtx, t: Xf = IDENTITY): Promise<string> {
    switch (e.type) {
      case 'shape':
        return this.shapeXml(e, ctx, t);
      case 'image':
        return this.imageXml(e, ctx, t);
      case 'table':
        return this.tableXml(e, ctx, t);
      case 'chart':
        return this.chartXml(e, ctx, t);
      case 'group':
        return this.groupXml(e, ctx, t);
    }
  }

  async groupXml(e: GroupEl, ctx: ElCtx, t: Xf): Promise<string> {
    // children use the group's unrotated frame: rotating the group rotates them about its centre
    const inner = rotateAbout(t, e.rot ?? 0, e.x + e.w / 2, e.y + e.h / 2);
    const kids = (await Promise.all(e.children.map((c) => this.elXml(c, ctx, inner)))).join('');
    const at = [e.name ? `draw:name="${x(e.name)}"` : '', ctx.ids.get(e.id) ? `xml:id="${ctx.ids.get(e.id)}" draw:id="${ctx.ids.get(e.id)}"` : '', e.hidden ? 'draw:display="none"' : ''].filter(Boolean).join(' ');
    return `<draw:g${at ? ` ${at}` : ''}>${e.alt ? `<svg:desc>${x(e.alt)}</svg:desc>` : ''}${this.eventXml(e.link)}${kids}</draw:g>`;
  }

  async shapeXml(e: ShapeEl, ctx: ElCtx, t: Xf): Promise<string> {
    const p = place(e, t);
    const mul = e.opacity ?? 1;
    const tctx: TextContext = { pres: this.pres, role: roleOf(e.ph), ph: e.ph, slideNumber: ctx.slideNumber };
    const straight = (e.geom === 'line' || e.geom === 'straightConnector1') && !e.text;
    if (straight) return this.lineXml(e, ctx, p);
    const g = gprops();
    await this.fillProps(e.fill, g, mul);
    this.lineProps(e.line, g, mul);
    this.shadowProps(e.shadow, g);
    if (e.locked) g.attrs.push('style:protect="position size"');
    const textbox = !!e.textbox;
    const para = this.textFrameProps(e.text, g, textbox, textbox || !!e.ph);
    // slide-level footers are plain text frames: LibreOffice only has them on master pages
    const cls = e.ph ? PH_CLASS[e.ph] : undefined;
    const empty = isEmptyText(e.text);
    const style = this.styleAttr(e, ctx, cls, g, para);
    const ph = this.phAttrs(e, ctx, cls, empty);
    const paraIds = ctx.paraIds.get(e.id);
    const text = e.text && !(e.ph && empty && !ctx.master) ? this.textXml(e.text, tctx, ctx, { paraIds, mul }) : '';
    const desc = e.alt ? `<svg:desc>${x(e.alt)}</svg:desc>` : '';
    const common = this.commonAttrs(e, ctx);
    if (e.ph || textbox) {
      // placeholders and text boxes are text frames
      if (e.ph === 'pic' && empty) return `<draw:frame ${style} ${common} ${frameAttrs(p)} ${ph}><draw:image/>${desc}</draw:frame>`;
      return `<draw:frame ${style} ${common} ${frameAttrs(p)}${ph ? ` ${ph}` : ''}><draw:text-box>${text}</draw:text-box>${this.eventXml(e.link)}${desc}</draw:frame>`;
    }
    return `<draw:custom-shape ${style} ${common} ${frameAttrs(p)}>${desc}${this.eventXml(e.link)}${text}${this.geometryXml(e, p)}</draw:custom-shape>`;
  }

  geometryXml(e: ShapeEl | ImageEl, p: Placed): string {
    const geom = e.type === 'shape' ? e.geom : (e.geom ?? 'rect');
    const custom = e.type === 'shape' && e.geom === 'custom' && e.custom?.length ? e.custom : undefined;
    const name = geom === 'textBox' ? 'rect' : geom;
    const adj = e.type === 'shape' ? e.adj : undefined;
    const g = custom ? customGeometry(custom, Math.max(p.w, 0.01), Math.max(p.h, 0.01)) : presetGeometry(name, Math.max(p.w, 0.01), Math.max(p.h, 0.01), adj);
    const type = custom ? 'non-primitive' : name === 'rect' ? 'rectangle' : name === 'ellipse' ? 'ellipse' : `ooxml-${name}`;
    const ep = enhancedPath(g.paths, p.w, p.h, g.text, type.startsWith('ooxml-'));
    const at = [`svg:viewBox="${ep.viewBox}"`, `draw:type="${x(type)}"`];
    if (!custom && name !== 'rect' && name !== 'ellipse') {
      const defs = DEFAULT_ADJ[name] ?? {};
      const keys = [...Object.keys(defs), ...Object.keys(adj ?? {}).filter((k) => !(k in defs)).sort()];
      if (keys.length) at.push(`draw:modifiers="${keys.map((k) => Math.round(adj?.[k] ?? defs[k] ?? 0)).join(' ')}"`);
    }
    at.push(`draw:enhanced-path="${ep.path}"`);
    if (ep.extended) at.push(`drawooo:enhanced-path="${ep.extended}"`);
    if (ep.textAreas) at.push(`draw:text-areas="${ep.textAreas}"`);
    if (e.flipH) at.push('draw:mirror-horizontal="true"');
    if (e.flipV) at.push('draw:mirror-vertical="true"');
    return ep.equations ? `<draw:enhanced-geometry ${at.join(' ')}>${ep.equations}</draw:enhanced-geometry>` : `<draw:enhanced-geometry ${at.join(' ')}/>`;
  }

  lineXml(e: ShapeEl, ctx: ElCtx, p: Placed): string {
    // end points in the shape's frame, flipped, then rotated about its centre
    let x1 = 0;
    let y1 = 0;
    let x2 = p.w;
    let y2 = p.h;
    if (e.flipH) [x1, x2] = [x2, x1];
    if (e.flipV) [y1, y2] = [y2, y1];
    const r = (p.rot * Math.PI) / 180;
    const cx = p.w / 2;
    const cy = p.h / 2;
    const rot = (px: number, py: number): [number, number] => [p.x + cx + (px - cx) * Math.cos(r) - (py - cy) * Math.sin(r), p.y + cy + (px - cx) * Math.sin(r) + (py - cy) * Math.cos(r)];
    const [ax, ay] = rot(x1, y1);
    const [bx, by] = rot(x2, y2);
    const g = gprops();
    this.lineProps(e.line ?? { color: '@tx1', width: 1.33 }, g, e.opacity ?? 1);
    this.shadowProps(e.shadow, g);
    g.attrs.push('draw:fill="none"');
    const style = ctx.styles.add('graphic', 'gr', gxml('style:graphic-properties', g));
    const ph = e.ph ? ` affice:placeholder="${e.ph}"` : '';
    const geom = e.geom !== 'line' ? ` affice:geometry="${x(e.geom)}"` : '';
    return `<draw:line draw:style-name="${style}" ${this.commonAttrs(e, ctx)} svg:x1="${cm(ax)}" svg:y1="${cm(ay)}" svg:x2="${cm(bx)}" svg:y2="${cm(by)}"${ph}${geom}>${e.alt ? `<svg:desc>${x(e.alt)}</svg:desc>` : ''}${this.eventXml(e.link)}<text:p/></draw:line>`;
  }

  async imageXml(e: ImageEl, ctx: ElCtx, t: Xf): Promise<string> {
    const p = place(e, t);
    const mul = e.opacity ?? 1;
    const g = gprops();
    const desc = e.alt ? `<svg:desc>${x(e.alt)}</svg:desc>` : '';
    // non-rectangular pictures are shapes filled with the picture
    if (e.geom && e.geom !== 'rect' && !e.svg) {
      await this.fillProps({ type: 'image', src: e.src, mode: 'stretch' }, g, mul);
      this.lineProps(e.line, g, mul);
      this.shadowProps(e.shadow, g);
      const style = ctx.styles.add('graphic', 'gr', gxml('style:graphic-properties', g));
      return `<draw:custom-shape draw:style-name="${style}" ${this.commonAttrs(e, ctx)} ${frameAttrs(p)} affice:picture="true">${desc}${this.eventXml(e.link)}${this.geometryXml(e, p)}</draw:custom-shape>`;
    }
    let images = '';
    let size: ReturnType<typeof imageSize> = null;
    if (e.svg) {
      const svg = this.addFile(`Pictures/image${++this.picCount}.svg`, strToU8(e.svg), 'image/svg+xml');
      images += `<draw:image xlink:href="${svg}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad" draw:mime-type="image/svg+xml"><text:p/></draw:image>`;
      const png = await rasterizeImage(`data:image/svg+xml;charset=utf-8,${encodeURIComponent(e.svg)}`, Math.max(64, e.w * 2), Math.max(64, e.h * 2));
      if (png) {
        const href = this.addFile(`Pictures/image${++this.picCount}.png`, png, 'image/png');
        images += `<draw:image xlink:href="${href}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad" draw:mime-type="image/png"><text:p/></draw:image>`;
      }
    } else {
      const pic = await this.picture(e.src);
      if (!pic) return '';
      images += `<draw:image xlink:href="${pic.href}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"><text:p/></draw:image>`;
      size = imageSize(pic.bytes);
    }
    g.attrs.push('draw:fill="none"');
    this.lineProps(e.line, g, mul);
    this.shadowProps(e.shadow, g);
    if (mul < 0.999) g.attrs.push(`draw:image-opacity="${pct(mul)}"`);
    if (e.flipH || e.flipV) g.attrs.push(`style:mirror="${[e.flipV ? 'vertical' : '', e.flipH ? 'horizontal' : ''].filter(Boolean).join(' ')}"`);
    if (e.crop && size) {
      // fo:clip is measured on the picture at its own size
      const wIn = size.w / size.dpiX;
      const hIn = size.h / size.dpiY;
      const [l, tp, r, b] = e.crop;
      const inch = (v: number) => `${r3(v)}in`;
      g.attrs.push(`fo:clip="rect(${inch(tp * hIn)}, ${inch(r * wIn)}, ${inch(b * hIn)}, ${inch(l * wIn)})"`);
    }
    const cls = e.ph ? 'graphic' : undefined;
    const style = this.styleAttr(e, ctx, cls, g);
    const ph = this.phAttrs(e, ctx, cls, false);
    return `<draw:frame ${style} ${this.commonAttrs(e, ctx)} ${frameAttrs(p)}${ph ? ` ${ph}` : ''}>${images}${this.eventXml(e.link)}${desc}</draw:frame>`;
  }

  async tableXml(e: TableEl, ctx: ElCtx, t: Xf): Promise<string> {
    const p = place(e, t);
    const st = e.style ?? {};
    const tctx: TextContext = { pres: this.pres, role: 'other', slideNumber: ctx.slideNumber };
    const cols = e.cols.map((w) => `<table:table-column table:style-name="${ctx.styles.add('table-column', 'co', `<style:table-column-properties style:column-width="${cm(w)}" style:use-optimal-column-width="false"/>`)}"/>`).join('');
    let rows = '';
    for (let r = 0; r < e.rows.length; r++) {
      const row = e.rows[r];
      const rs = ctx.styles.add('table-row', 'ro', `<style:table-row-properties style:row-height="${cm(row.h)}" style:use-optimal-row-height="false"/>`);
      let cells = '';
      for (let c = 0; c < row.cells.length; c++) {
        const cell = row.cells[c];
        if (cell.merged) {
          cells += '<table:covered-table-cell/>';
          continue;
        }
        const look = cellLook(e, r, c);
        const fill = cell.fill ?? look.fill;
        const borders = cell.borders ?? look.borders;
        const body = styledCellText(e, r, c, cell.text);
        const g = gprops();
        await this.fillProps(fill ?? { type: 'none' }, g);
        const inset = body.inset ?? DEFAULT_INSET;
        g.attrs.push(`draw:textarea-vertical-align="${ANCHOR[body.anchor ?? 't']}"`, `fo:padding-left="${cm(inset[0])}"`, `fo:padding-top="${cm(inset[1])}"`, `fo:padding-right="${cm(inset[2])}"`, `fo:padding-bottom="${cm(inset[3])}"`);
        const border = (l: Line | null | undefined) => (l && l.width > 0 ? `${pt(l.width * 0.75)} solid ${colorOf(l.color, this.theme).hex}` : 'none');
        const bxml = `fo:border-top="${border(borders[0])}" fo:border-right="${border(borders[1])}" fo:border-bottom="${border(borders[2])}" fo:border-left="${border(borders[3])}"`;
        const bg = fill?.type === 'solid' ? colorOf(fill.color, this.theme).hex : 'transparent';
        const cs = ctx.styles.add('table-cell', 'ce', `${gxml('loext:graphic-properties', g)}<style:paragraph-properties ${bxml}/><style:table-cell-properties fo:background-color="${bg}" ${bxml} style:vertical-align="${ANCHOR[body.anchor ?? 't']}" fo:padding-left="${cm(inset[0])}" fo:padding-top="${cm(inset[1])}" fo:padding-right="${cm(inset[2])}" fo:padding-bottom="${cm(inset[3])}"/>`);
        const span = `${(cell.colSpan ?? 1) > 1 ? ` table:number-columns-spanned="${cell.colSpan}"` : ''}${(cell.rowSpan ?? 1) > 1 ? ` table:number-rows-spanned="${cell.rowSpan}"` : ''}`;
        cells += `<table:table-cell table:style-name="${cs}"${span}>${this.textXml(body, tctx, ctx)}</table:table-cell>`;
      }
      rows += `<table:table-row table:style-name="${rs}">${cells}</table:table-row>`;
    }
    const flags = [
      st.family && st.family !== 'none' ? `table:template-name="affice-${st.family}-${st.accent ?? 'accent1'}"` : st.family === 'none' ? 'table:template-name="affice-none"' : '',
      st.firstRow ? 'table:use-first-row-styles="true"' : '',
      st.lastRow ? 'table:use-last-row-styles="true"' : '',
      st.firstCol ? 'table:use-first-column-styles="true"' : '',
      st.lastCol ? 'table:use-last-column-styles="true"' : '',
      st.bandRows ? 'table:use-banding-rows-styles="true"' : '',
      st.bandCols ? 'table:use-banding-columns-styles="true"' : '',
      st.styleId ? `affice:table-style-id="${x(st.styleId)}"` : '',
    ].filter(Boolean);
    const g = gprops();
    g.attrs.push('draw:fill="none"', 'draw:stroke="none"');
    const style = ctx.styles.add('graphic', 'gr', gxml('style:graphic-properties', g));
    const ph = e.ph ? ` affice:placeholder="${e.ph}"${e.phIdx ? ` affice:placeholder-index="${e.phIdx}"` : ''}` : '';
    return `<draw:frame draw:style-name="${style}" ${this.commonAttrs(e, ctx)} ${frameAttrs(p)}${ph}><table:table${flags.length ? ` ${flags.join(' ')}` : ''}>${cols}${rows}</table:table>${e.alt ? `<svg:desc>${x(e.alt)}</svg:desc>` : ''}</draw:frame>`;
  }

  async chartXml(e: ChartEl, ctx: ElCtx, t: Xf): Promise<string> {
    const p = place(e, t);
    const name = `Object ${++this.objCount}`;
    const obj = slideChartObject(e.chart, this.theme, e.w, e.h);
    this.files[`${name}/content.xml`] = strToU8(obj.content);
    this.files[`${name}/styles.xml`] = strToU8(obj.styles);
    this.manifest.push(`<manifest:file-entry manifest:full-path="${name}/content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="${name}/styles.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="${name}/" manifest:version="1.3" manifest:media-type="application/vnd.oasis.opendocument.chart"/>`);
    const g = gprops();
    g.attrs.push('draw:fill="none"', 'draw:stroke="none"');
    const style = ctx.styles.add('graphic', 'gr', gxml('style:graphic-properties', g));
    const ph = e.ph ? ` affice:placeholder="${e.ph}"${e.phIdx ? ` affice:placeholder-index="${e.phIdx}"` : ''}` : '';
    return `<draw:frame draw:style-name="${style}" ${this.commonAttrs(e, ctx)} ${frameAttrs(p)}${ph}><draw:object draw:notify-on-update-of-ranges="" xlink:href="./${name}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"><loext:p/></draw:object>${e.alt ? `<svg:desc>${x(e.alt)}</svg:desc>` : ''}</draw:frame>`;
  }

  /* ------------------------------------------------------ backgrounds */


  /** A page background: attributes and theme colour children for drawing-page properties. */
  async pageFill(f: Fill | undefined): Promise<GProps | null> {
    if (!f) return null;
    const g = gprops();
    await this.fillProps(f, g);
    g.attrs.unshift('draw:background-size="border"');
    return g;
  }

  /* --------------------------------------------------------- masters */

  /** Presentation styles (text defaults per placeholder class) of one master page. */
  presentationStyles(master: string, display: string, l: Layout, background: GProps | null): string {
    const find = (kinds: PlaceholderType[]) => l.placeholders.find((p) => p.ph && kinds.includes(p.ph)) as ShapeEl | undefined;
    const title = find(['ctrTitle', 'title']);
    const sub = find(['subTitle']);
    const body = find(['obj', 'body']);
    const out: string[] = [];
    const style = (cls: string, content: string, parent?: string) =>
      out.push(`<style:style style:name="${master}-${cls}" style:display-name="${x(display)}-${cls}" style:family="presentation"${parent ? ` style:parent-style-name="${master}-${parent}"` : ''}>${content}</style:style>`);
    const frame = (el: ShapeEl | undefined, anchor: 't' | 'm' | 'b') => {
      const g = gprops();
      g.attrs.push('draw:stroke="none"', 'draw:fill="none"');
      const inner = this.textFrameProps(el?.text ?? { paras: [], anchor, autofit: 'shrink' }, g, false);
      return { g, inner };
    };
    const textOf = (el: ShapeEl | undefined, ph: PlaceholderType, level = 0) => {
      const tctx: TextContext = { pres: this.pres, role: roleOf(ph), ph };
      const bodyT = el?.text ?? { paras: [] };
      const p0: Para = { ...(bodyT.paras[0] ?? { runs: [] }), level, runs: [] };
      const e = effectivePara(p0, tctx);
      const r = effectiveRun({}, p0, bodyT, tctx);
      return { e, r, tctx, p0 };
    };
    // title
    {
      const t = textOf(title, title?.ph ?? 'title');
      const { g, inner } = frame(title, 'm');
      style('title', `${gxml('style:graphic-properties', g)}<style:paragraph-properties fo:text-align="${ALIGN[t.p0.align ?? 'left']}" fo:line-height="${pct(t.e.lineSpacing)}" fo:margin-top="0cm" fo:margin-bottom="0cm" style:writing-mode="lr-tb"/>${this.runProps(t.r)}${inner.replace('<style:paragraph-properties style:writing-mode="lr-tb"/>', '')}`);
    }
    // subtitle
    {
      const t = textOf(sub, 'subTitle');
      const { g, inner } = frame(sub, 't');
      style('subtitle', `${gxml('style:graphic-properties', g)}<style:paragraph-properties fo:text-align="${ALIGN[t.p0.align ?? 'center']}" fo:line-height="${pct(t.e.lineSpacing)}" fo:margin-top="0cm" fo:margin-bottom="0cm"/>${this.runProps(t.r)}${inner.replace('<style:paragraph-properties style:writing-mode="lr-tb"/>', '')}`);
    }
    // outline levels: bullets in a list style on the first level
    {
      const ph = body?.ph ?? 'obj';
      const { g } = frame(body, 't');
      let levels = '';
      for (let i = 0; i < 9; i++) {
        const t = textOf(body, ph, i);
        if (t.e.bullet && t.e.bullet.type !== 'none') levels += this.listLevelXml(i + 1, t.p0, t.e, body?.text ?? { paras: [] }, t.tctx);
        else levels += `<text:list-level-style-number text:level="${i + 1}" style:num-format=""><style:list-level-properties text:space-before="${cm(Math.max(0, t.e.marL))}"/></text:list-level-style-number>`;
      }
      g.kids.push(`<text:list-style style:name="${master}-outline1">${levels}</text:list-style>`);
      for (let i = 0; i < 9; i++) {
        const t = textOf(body, ph, i);
        const para = `<style:paragraph-properties fo:margin-top="${pt(t.e.spaceBefore)}" fo:margin-bottom="${pt(t.e.spaceAfter)}" fo:line-height="${pct(t.e.lineSpacing)}" fo:text-align="${ALIGN[t.p0.align ?? 'left']}"/>`;
        style(`outline${i + 1}`, `${i === 0 ? gxml('style:graphic-properties', g) : ''}${para}${this.runProps(t.r)}`, i === 0 ? undefined : `outline${i}`);
      }
    }
    const other = this.pres.master.text.other;
    const otherRun = { font: other.font, size: 12, color: other.color } as ReturnType<typeof effectiveRun>;
    style('notes', `<style:graphic-properties draw:stroke="none" draw:fill="none"/><style:paragraph-properties fo:margin-left="0cm" fo:text-indent="0cm"/>${this.runProps(otherRun)}`);
    // LibreOffice takes a master page's background from this style
    const bgProps = gprops();
    bgProps.attrs.push('draw:stroke="none"', ...(background ? background.attrs.filter((a) => !a.startsWith('draw:background-size')) : ['draw:fill="none"']));
    bgProps.kids.push(...(background?.kids ?? []));
    style('background', gxml('style:graphic-properties', bgProps));
    style('backgroundobjects', `<style:graphic-properties draw:shadow="hidden"/>${this.runProps({ ...otherRun, size: other.size })}`);
    return out.join('');
  }

  footerFramesXml(ctx: ElCtx): string {
    const sx = this.pres.size.w / 1280;
    const y = this.pres.size.h - 53;
    const frame = (cls: string, xx: number, w: number, align: string, content: string) => {
      const g = gprops();
      g.attrs.push('draw:stroke="none"', 'draw:fill="none"', 'draw:textarea-vertical-align="middle"', 'draw:auto-grow-height="false"', 'fo:padding-left="0.254cm"', 'fo:padding-right="0.254cm"', 'fo:padding-top="0.127cm"', 'fo:padding-bottom="0.127cm"');
      const style = ctx.styles.add('presentation', 'pr', gxml('style:graphic-properties', g), ` style:parent-style-name="${ctx.masterName}-backgroundobjects"`);
      const run = { font: '+minor', size: 12, color: '@tx1+45' } as ReturnType<typeof effectiveRun>;
      const ps = ctx.styles.add('paragraph', 'P', `<style:paragraph-properties fo:text-align="${align}"/>${this.runProps(run)}`);
      const ts = ctx.styles.add('text', 'T', this.runProps(run));
      return `<draw:frame presentation:style-name="${style}" draw:layer="backgroundobjects" svg:width="${cm(w * sx)}" svg:height="${cm(38)}" svg:x="${cm(xx * sx)}" svg:y="${cm(y)}" presentation:class="${cls}"><draw:text-box><text:p text:style-name="${ps}"><text:span text:style-name="${ts}">${content}</text:span></text:p></draw:text-box></draw:frame>`;
    };
    return frame('date-time', 88, 288, 'start', '<presentation:date-time/>') + frame('footer', 424, 432, 'center', '<presentation:footer/>') + frame('page-number', 904, 288, 'end', '<text:page-number>&lt;number&gt;</text:page-number>');
  }

  async masterPageXml(l: Layout, name: string, pageLayout: string): Promise<string> {
    const ctx: ElCtx = { styles: this.masterStyles, layer: 'backgroundobjects', master: true, masterName: name, ids: new Map(), paraIds: new Map() };
    const bg = (await this.pageFill(l.background ?? this.pres.master.background)) ?? { attrs: ['draw:fill="none"'], kids: [] };
    const dp = this.masterStyles.add('drawing-page', 'dp', gxml('style:drawing-page-properties', bg));
    const c = this.theme.colors;
    const theme = `<loext:theme loext:name="${x(this.theme.name)}"><loext:theme-colors loext:name="${x(this.theme.name)}">${(
      [
        ['dark1', c.tx1],
        ['light1', c.bg1],
        ['dark2', c.tx2],
        ['light2', c.bg2],
        ['accent1', c.accent1],
        ['accent2', c.accent2],
        ['accent3', c.accent3],
        ['accent4', c.accent4],
        ['accent5', c.accent5],
        ['accent6', c.accent6],
        ['hyperlink', c.hlink],
        ['followed-hyperlink', c.folHlink],
      ] as [string, string][]
    )
      .map(([n, v]) => `<loext:color loext:name="${n}" loext:color="${resolveHex(v, this.theme).hex}"/>`)
      .join('')}</loext:theme-colors></loext:theme>`;
    let shapes = '';
    for (const d of [...this.pres.master.decor, ...l.decor]) shapes += await this.elXml(d, ctx);
    for (const p of l.placeholders) shapes += await this.masterPlaceholderXml(p, ctx);
    shapes += this.footerFramesXml(ctx);
    const notes = `<presentation:notes style:page-layout-name="PM2"><draw:page-thumbnail presentation:style-name="${name}-background" draw:layer="backgroundobjects" svg:width="14cm" svg:height="${cm(((14 / 2.54) * 96 * this.pres.size.h) / this.pres.size.w)}" svg:x="3.5cm" svg:y="2.5cm" presentation:class="page"/><draw:frame presentation:style-name="${name}-notes" draw:layer="backgroundobjects" svg:width="16.8cm" svg:height="13.4cm" svg:x="2.1cm" svg:y="13.4cm" presentation:class="notes" presentation:placeholder="true"><draw:text-box/></draw:frame></presentation:notes>`;
    return `<style:master-page style:name="${name}" style:display-name="${x(l.name)}" style:page-layout-name="${pageLayout}" draw:style-name="${dp}" affice:layout-id="${x(l.id)}" affice:layout-type="${l.type}">${theme}${shapes}${notes}</style:master-page>`;
  }

  /** A layout placeholder on its master page, with a prompt that carries its text defaults. */
  async masterPlaceholderXml(p: El, ctx: ElCtx): Promise<string> {
    if (p.type !== 'shape') return this.elXml(p, ctx);
    const prompt: Record<string, string> = { title: 'Click to edit the title text format', ctrTitle: 'Click to edit the title text format', subTitle: 'Click to edit the subtitle text format', pic: 'Picture', dt: 'Date', ftr: 'Footer', sldNum: '#' };
    const text: TextBody = p.text ? { ...p.text, paras: [{ ...(p.text.paras[0] ?? { runs: [] }), runs: [{ ...(p.text.paras[0]?.endRun ?? {}), text: prompt[p.ph ?? ''] ?? 'Click to edit the outline text format' }] }] } : { paras: [{ runs: [{ text: prompt[p.ph ?? ''] ?? 'Text' }] }] };
    return this.shapeXml({ ...p, text }, ctx, IDENTITY);
  }

  /* ----------------------------------------------------------- slides */

  async slideXml(s: Slide, i: number, masterName: string, pageLayout: string | undefined): Promise<string> {
    const pres = this.pres;
    const layout = pres.layouts.find((l) => l.id === s.layout) ?? pres.layouts[0];
    const pageId = this.nextId('page');
    const ids = new Map<string, string>();
    const paraIds = new Map<string, Map<number, string>>();
    const build = s.anims?.length ? buildOf(s) : null;
    const items: BuildItem[] = build ? [...(build.auto ?? []), ...build.steps.flat()] : [];
    for (const it of items) {
      if (it.para !== undefined) {
        const m = paraIds.get(it.anim.el) ?? new Map<number, string>();
        if (!m.has(it.para)) m.set(it.para, this.nextId('p'));
        paraIds.set(it.anim.el, m);
      } else if (!ids.has(it.anim.el)) ids.set(it.anim.el, this.nextId('id'));
    }
    const ctx: ElCtx = { styles: this.content, layer: 'layout', master: false, masterName, slideNumber: i + 1, ids, paraIds };
    let shapes = '';
    for (const e of s.elements) shapes += await this.elXml(e, ctx);

    // page style: background, transition, visibility, footers
    const dpAt: string[] = [];
    const bg = s.background ? await this.pageFill(s.background) : null;
    if (bg) dpAt.push(...bg.attrs, 'presentation:background-visible="true"');
    dpAt.push(`presentation:background-objects-visible="${!s.hideDecor}"`);
    const f = pres.footer;
    const showFooters = !!f && !(f.skipTitle && layout?.type === 'title');
    dpAt.push(`presentation:display-footer="${showFooters && !!f?.text}"`, `presentation:display-page-number="${showFooters && !!f?.slideNumber}"`, `presentation:display-date-time="${showFooters && !!f?.date}"`, 'presentation:display-header="false"');
    if (s.hidden) dpAt.push('presentation:visibility="hidden"');
    const tr = s.transition;
    const smil = tr ? smilTransition(tr) : null;
    if (tr?.after !== undefined) dpAt.push('presentation:transition-type="automatic"', `presentation:duration="PT${Math.round(tr.after) / 1000}S"`);
    if (tr && smil) {
      const dur = tr.dur ?? 700;
      dpAt.push(`presentation:transition-speed="${dur < 500 ? 'fast' : dur < 1000 ? 'medium' : 'slow'}"`, `smil:type="${smil.type}"`, `smil:subtype="${smil.subtype}"`);
      if (smil.reverse) dpAt.push('smil:direction="reverse"');
    }
    const dp = this.content.add('drawing-page', 'dp', gxml('style:drawing-page-properties', { attrs: dpAt, kids: bg?.kids ?? [] }));

    // timing: the transition, then the main sequence of effects
    let timing = '';
    const filter = tr && smil ? `<anim:par smil:begin="${pageId}.begin"><anim:transitionFilter smil:dur="${secs(tr.dur ?? 700)}" smil:type="${smil.type}" smil:subtype="${smil.subtype}"${smil.reverse ? ' smil:direction="reverse"' : ''}/></anim:par>` : '';
    const seq = build ? this.mainSequence(build.auto, build.steps, ids, paraIds) : '';
    if (filter || seq) timing = `<anim:par presentation:node-type="timing-root">${filter}${seq}</anim:par>`;

    const notes = s.notes?.trim() ? this.notesXml(s.notes, i) : '';
    const footerRefs = showFooters ? `${f?.text ? ' presentation:use-footer-name="ftr1"' : ''}${f?.date ? ' presentation:use-date-time-name="dtd1"' : ''}` : '';
    const extra = `${s.section ? ` affice:section="${x(s.section)}"` : ''}${tr && (tr.type === 'morph' || tr.type === 'flip') ? ` affice:transition="${tr.type}"${tr.dir ? ` affice:transition-direction="${tr.dir}"` : ''}` : ''}${tr?.onClick === false ? ' affice:advance-on-click="false"' : ''}`;
    return `<draw:page draw:name="${this.pageNames[i]}" draw:style-name="${dp}" draw:master-page-name="${masterName}"${pageLayout ? ` presentation:presentation-page-layout-name="${pageLayout}"` : ''}${footerRefs} xml:id="${pageId}" draw:id="${pageId}"${extra}>${shapes}${timing}${notes}</draw:page>`;
  }

  mainSequence(auto: BuildItem[] | null, steps: BuildItem[][], ids: Map<string, string>, paraIds: Map<string, Map<number, string>>): string {
    const effect = (it: BuildItem, nodeType: string) => {
      const target = it.para !== undefined ? paraIds.get(it.anim.el)?.get(it.para) : ids.get(it.anim.el);
      if (!target) return '';
      const p = animPreset(it.anim, target);
      const cls = it.anim.cls === 'entr' ? 'entrance' : it.anim.cls === 'exit' ? 'exit' : 'emphasis';
      const repeat = it.anim.repeat && it.anim.repeat > 1 ? ` smil:repeatCount="${it.anim.repeat}"` : '';
      return `<anim:par smil:begin="${secs(it.anim.delay)}" smil:fill="hold"${repeat} presentation:node-type="${nodeType}" presentation:preset-class="${cls}" presentation:preset-id="${p.id}"${p.sub ? ` presentation:preset-sub-type="${p.sub}"` : ''}>${p.xml}</anim:par>`;
    };
    const step = (items: BuildItem[], click: boolean) => {
      // effects that start together share a group; "after previous" effects start new groups
      const groups: { at: number; items: BuildItem[] }[] = [];
      for (const it of items) {
        const start = it.offset - it.anim.delay;
        const g = groups.find((gg) => Math.abs(gg.at - start) < 1);
        if (g) g.items.push(it);
        else groups.push({ at: start, items: [it] });
      }
      const inner = groups
        .map((g, gi) => {
          const nodes = g.items.map((it, ii) => effect(it, gi === 0 && ii === 0 ? (click ? 'on-click' : it.anim.start === 'with' ? 'with-previous' : 'after-previous') : ii === 0 ? 'after-previous' : 'with-previous')).join('');
          return `<anim:par smil:begin="${secs(g.at)}" smil:fill="hold">${nodes}</anim:par>`;
        })
        .join('');
      return `<anim:par smil:begin="${click ? 'next' : '0s'}" smil:fill="hold">${inner}</anim:par>`;
    };
    let out = '';
    if (auto?.length) out += step(auto, false);
    for (const s of steps) out += step(s, true);
    return out ? `<anim:seq smil:dur="indefinite" presentation:node-type="main-sequence">${out}</anim:seq>` : '';
  }

  notesXml(notes: string, i: number): string {
    const ctx: ElCtx = { styles: this.content, layer: 'layout', master: false, masterName: '', ids: new Map(), paraIds: new Map() };
    const tctx: TextContext = { pres: this.pres, role: 'other' };
    const body: TextBody = { paras: notes.split('\n').map((line) => ({ runs: line ? [{ text: line, size: 12 }] : [], endRun: { size: 12 } })) };
    const pr = this.content.add('presentation', 'pr', '<style:graphic-properties draw:fill="none" draw:stroke="none" draw:auto-grow-height="false" fo:min-height="13.4cm"/>');
    const h = r3(((14 / 2.54) * 96 * this.pres.size.h) / this.pres.size.w);
    return `<presentation:notes draw:style-name="${this.content.add('drawing-page', 'dp', '<style:drawing-page-properties presentation:display-header="false" presentation:display-footer="false" presentation:display-page-number="false" presentation:display-date-time="false"/>')}"><draw:page-thumbnail draw:layer="layout" svg:width="14cm" svg:height="${cm(h)}" svg:x="3.5cm" svg:y="2.5cm" draw:page-number="${i + 1}" presentation:class="page"/><draw:frame presentation:style-name="${pr}" draw:layer="layout" svg:width="16.8cm" svg:height="13.4cm" svg:x="2.1cm" svg:y="13.4cm" presentation:class="notes" presentation:user-transformed="true"><draw:text-box>${this.textXml(body, tctx, ctx)}</draw:text-box></draw:frame></presentation:notes>`;
  }
}

/** Encodes run text for ODF: runs of spaces, tabs and line breaks become elements. */
function encodeText(text: string, prevSpace: boolean): { xml: string; endsWithSpace: boolean } {
  let out = '';
  let spaces = 0;
  let after = prevSpace;
  const flush = () => {
    if (!spaces) return;
    if (after) out += spaces === 1 ? '<text:s/>' : `<text:s text:c="${spaces}"/>`;
    else {
      out += ' ';
      if (spaces > 1) out += spaces === 2 ? '<text:s/>' : `<text:s text:c="${spaces - 1}"/>`;
    }
    after = true;
    spaces = 0;
  };
  for (const ch of text) {
    if (ch === ' ') {
      spaces++;
      continue;
    }
    flush();
    if (ch === '\t') {
      out += '<text:tab/>';
      after = false;
    } else if (ch === '\n' || ch === '\v' || ch === '\u2028') {
      out += '<text:line-break/>';
      after = true;
    } else {
      out += x(ch);
      after = false;
    }
  }
  const trailing = spaces > 0;
  flush();
  return { xml: out, endsWithSpace: trailing || (after && text.length > 0 && /[\n\v\u2028]$/.test(text)) };
}

/** Presentation page layout (the placeholder kinds LibreOffice's layout panel matches). */
function pageLayoutXml(name: string, l: Layout): string {
  const obj = (ph: PlaceholderType | undefined) => (ph === 'title' || ph === 'ctrTitle' ? 'title' : ph === 'subTitle' ? 'subtitle' : ph === 'pic' ? 'graphic' : ph === 'body' || ph === 'obj' ? 'outline' : undefined);
  const phs = l.placeholders
    .map((p) => ({ p, o: obj(p.ph) }))
    .filter((q) => q.o)
    .map(({ p, o }) => `<presentation:placeholder presentation:object="${o}" svg:x="${cm(p.x)}" svg:y="${cm(p.y)}" svg:width="${cm(p.w)}" svg:height="${cm(p.h)}"/>`)
    .join('');
  return `<style:presentation-page-layout style:name="${name}">${phs}</style:presentation-page-layout>`;
}

function autoLayoutNumber(l: Layout): number {
  const kinds = l.placeholders.map((p) => p.ph);
  const titles = kinds.filter((k) => k === 'title' || k === 'ctrTitle').length;
  const subs = kinds.filter((k) => k === 'subTitle').length;
  const outlines = kinds.filter((k) => k === 'obj' || k === 'body').length;
  if (!kinds.length) return 20;
  if (titles && subs) return 0;
  if (titles && !outlines) return 19;
  if (outlines === 2) return 3;
  if (outlines === 4) return 18;
  return 1;
}

/** Writes a presentation as an OpenDocument presentation (.odp). */
export async function exportOdp(pres: Presentation): Promise<Uint8Array> {
  const w = new OdpWriter(pres);
  const layouts = pres.layouts.length ? pres.layouts : [{ id: 'blank', name: 'Blank', type: 'blank' as const, decor: [], placeholders: [] }];

  // master pages: one per layout, with unique names
  const used = new Set<string>();
  const masterName = new Map<string, string>();
  for (const l of layouts) {
    let n = styleNameOf(l.name || l.type);
    let k = 2;
    while (used.has(n)) n = `${styleNameOf(l.name || l.type)}${k++}`;
    used.add(n);
    masterName.set(l.id, n);
  }
  const pageLayoutName = new Map<string, string>();
  const pageLayouts: string[] = [];
  layouts.forEach((l, i) => {
    const name = `AL${i + 1}T${autoLayoutNumber(l)}`;
    pageLayoutName.set(l.id, name);
    pageLayouts.push(pageLayoutXml(name, l));
  });

  const masters: string[] = [];
  const presStyles: string[] = [];
  for (const l of layouts) {
    const name = masterName.get(l.id)!;
    const bg = await w.pageFill(l.background ?? pres.master.background);
    presStyles.push(w.presentationStyles(name, l.name || l.type, l, bg));
    masters.push(await w.masterPageXml(l, name, 'PM1'));
  }

  const pages: string[] = [];
  for (let i = 0; i < pres.slides.length; i++) {
    const s = pres.slides[i];
    const l = layouts.find((ll) => ll.id === s.layout) ?? layouts[0];
    pages.push(await w.slideXml(s, i, masterName.get(l.id)!, pageLayoutName.get(l.id)));
  }

  const f = pres.footer;
  const decls = `${f?.text ? `<presentation:footer-decl presentation:name="ftr1">${x(f.text)}</presentation:footer-decl>` : ''}${f?.date ? '<presentation:date-time-decl presentation:name="dtd1" presentation:source="current-date" style:data-style-name="D1"/>' : ''}`;
  const dateStyle = '<number:date-style style:name="D1" number:automatic-order="true"><number:month/><number:text>/</number:text><number:day/><number:text>/</number:text><number:year number:style="long"/></number:date-style>';
  const content = `${XML_DECL}<office:document-content ${XMLNS} office:version="1.3"><office:automatic-styles>${dateStyle}${w.content.xml.join('')}</office:automatic-styles><office:body><office:presentation>${decls}${pages.join('')}</office:presentation></office:body></office:document-content>`;

  const other = pres.master.text.other;
  const minor = resolveFont(other.font, pres.theme) ?? pres.theme.fonts.minor;
  const defaults = `<style:default-style style:family="graphic"><style:graphic-properties svg:stroke-color="#3465a4" draw:fill-color="#729fcf" fo:wrap-option="no-wrap"/><style:paragraph-properties style:text-autospace="ideograph-alpha" style:punctuation-wrap="simple" style:line-break="strict" style:writing-mode="page" style:font-independent-line-spacing="false"><style:tab-stops/></style:paragraph-properties><style:text-properties style:use-window-font-color="true" fo:font-family="${x(quoteFont(minor))}" fo:font-size="${pt(other.size)}" style:font-size-asian="${pt(other.size)}" style:font-size-complex="${pt(other.size)}" style:letter-kerning="true"/></style:default-style>`;
  const size = pres.size;
  const pm = `<style:page-layout style:name="PM1"><style:page-layout-properties fo:margin-top="0cm" fo:margin-bottom="0cm" fo:margin-left="0cm" fo:margin-right="0cm" fo:page-width="${cm(size.w)}" fo:page-height="${cm(size.h)}" style:print-orientation="${size.w >= size.h ? 'landscape' : 'portrait'}"/></style:page-layout><style:page-layout style:name="PM2"><style:page-layout-properties fo:margin-top="0cm" fo:margin-bottom="0cm" fo:margin-left="0cm" fo:margin-right="0cm" fo:page-width="21cm" fo:page-height="29.7cm" style:print-orientation="portrait"/></style:page-layout>`;
  const layers = '<draw:layer-set><draw:layer draw:name="layout"/><draw:layer draw:name="background"/><draw:layer draw:name="backgroundobjects"/><draw:layer draw:name="controls"/><draw:layer draw:name="measurelines"/></draw:layer-set>';
  const styles = `${XML_DECL}<office:document-styles ${XMLNS} office:version="1.3"><office:styles>${w.namedXml()}${defaults}${pageLayouts.join('')}${presStyles.join('')}</office:styles><office:automatic-styles>${pm}${w.masterStyles.xml.join('')}</office:automatic-styles><office:master-styles>${layers}${masters.join('')}</office:master-styles></office:document-styles>`;

  const props = pres.props;
  const now = new Date().toISOString().slice(0, 19);
  const created = props.created ? new Date(props.created).toISOString().slice(0, 19) : now;
  const meta = `${XML_DECL}<office:document-meta ${XMLNS} office:version="1.3"><office:meta><meta:generator>Affice</meta:generator>${props.title ? `<dc:title>${x(props.title)}</dc:title>` : ''}${props.subject ? `<dc:subject>${x(props.subject)}</dc:subject>` : ''}${props.author ? `<meta:initial-creator>${x(props.author)}</meta:initial-creator><dc:creator>${x(props.author)}</dc:creator>` : ''}<meta:creation-date>${created}</meta:creation-date><dc:date>${now}</dc:date>${pres.design ? `<meta:user-defined meta:name="AfficeDesign">${x(pres.design)}</meta:user-defined>` : ''}<meta:user-defined meta:name="AfficeThemeFonts">${x(`${pres.theme.fonts.major}|${pres.theme.fonts.minor}`)}</meta:user-defined><meta:document-statistic meta:object-count="${pres.slides.reduce((n, s) => n + s.elements.length, 0)}"/></office:meta></office:document-meta>`;

  const files: Zippable = { mimetype: [strToU8(MIME), { level: 0 }] };
  files['content.xml'] = strToU8(content);
  files['styles.xml'] = strToU8(styles);
  files['meta.xml'] = strToU8(meta);
  for (const [k, v] of Object.entries(w.files)) files[k] = v;
  files['META-INF/manifest.xml'] = strToU8(
    `${XML_DECL}<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="${MIME}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>${w.manifest.join('')}</manifest:manifest>`,
  );
  // the mimetype entry must come first and be stored uncompressed
  return zipSync(files, { level: 6 });
}

